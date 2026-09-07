[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Executable,
    [string]$OutputDirectory,
    [string]$WixDirectory,
    [switch]$SkipIceValidation
)
$ErrorActionPreference = "Stop"
$project = $PSScriptRoot
$release = Get-Content -LiteralPath (Join-Path $project "installer\release.json") -Raw | ConvertFrom-Json
$executable = (Resolve-Path -LiteralPath $Executable).Path
function Assert-AnyCpu([string]$path) {
    $bytes = [IO.File]::ReadAllBytes($path)
    $pe = [BitConverter]::ToInt32($bytes, 60)
    if ([BitConverter]::ToUInt32($bytes, $pe) -ne 0x4550 -or
        [BitConverter]::ToUInt16($bytes, $pe + 4) -ne 0x14c -or
        [BitConverter]::ToUInt16($bytes, $pe + 24) -ne 0x10b) {
        throw "The MSI requires the existing PE32 AnyCPU CLR executable, not a native/x64/ARM64 binary."
    }
    $clrRva = [BitConverter]::ToUInt32($bytes, $pe + 24 + 208)
    $sections = [BitConverter]::ToUInt16($bytes, $pe + 6)
    $sectionTable = $pe + 24 + [BitConverter]::ToUInt16($bytes, $pe + 20)
    for ($i = 0; $i -lt $sections; $i++) {
        $section = $sectionTable + $i * 40
        $virtualAddress = [BitConverter]::ToUInt32($bytes, $section + 12)
        $rawSize = [BitConverter]::ToUInt32($bytes, $section + 16)
        if ($clrRva -ge $virtualAddress -and $clrRva -lt ([long]$virtualAddress + $rawSize)) {
            $offset = [BitConverter]::ToUInt32($bytes, $section + 20) + $clrRva - $virtualAddress
            $flags = [BitConverter]::ToUInt32($bytes, $offset + 16)
            if (($flags -band 1) -eq 0 -or ($flags -band 0x20002) -ne 0) {
                throw "The payload is not IL-only AnyCPU (32-bit-required/preferred binaries are not accepted)."
            }
            return
        }
    }
    throw "The payload has no readable CLR header."
}
Assert-AnyCpu $executable
$version = [Diagnostics.FileVersionInfo]::GetVersionInfo($executable).FileVersion
if ($version -ne "$($release.version).0") {
    throw "Expected release executable $($release.version).0; found $version. Build the matching candidate first."
}
if ([IO.Path]::GetFileName($executable) -ne "JarvisPowerPoint.exe") {
    throw "The payload must be named JarvisPowerPoint.exe."
}
$assemblyName = [Reflection.AssemblyName]::GetAssemblyName($executable)
if ($assemblyName.Version.ToString() -ne "$($release.version).0") {
    throw "The assembly version must also match the MSI release version."
}
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { $OutputDirectory = Join-Path $project "bin\msi" }
$output = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
if ([string]::IsNullOrWhiteSpace($WixDirectory)) {
    $candleCommand = Get-Command candle.exe -ErrorAction SilentlyContinue
    $WixDirectory = if ($candleCommand) { Split-Path -Parent $candleCommand.Source } else {
        Join-Path $project "installer\.tools\wix3"
    }
}
$candle = Join-Path $WixDirectory "candle.exe"
$light = Join-Path $WixDirectory "light.exe"
if (-not (Test-Path -LiteralPath $candle) -or -not (Test-Path -LiteralPath $light)) {
    throw "WiX 3.14.1 is missing. Run .\installer\Restore-Wix.ps1 for a project-local restore, then retry."
}
$candle = (Resolve-Path -LiteralPath $candle).Path
$light = (Resolve-Path -LiteralPath $light).Path
foreach ($tool in @($candle, $light)) {
    if ([Diagnostics.FileVersionInfo]::GetVersionInfo($tool).FileVersion -ne $release.wixVersion) {
        throw "Use the pinned WiX $($release.wixVersion) toolset (-WixDirectory)."
    }
}
New-Item -ItemType Directory -Path $output -Force | Out-Null
$work = Join-Path $project ("installer\.work\" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $work -Force | Out-Null
$oldTemp = $env:TEMP
$oldTmp = $env:TMP
try {
    $env:TEMP = $work
    $env:TMP = $work
    $object = Join-Path $work "Product.wixobj"
    & $candle -nologo -wx -arch x86 "-dProductVersion=$($release.version)" `
        "-dProductCode=$($release.productCode)" "-dUpgradeCode=$($release.upgradeCode)" "-dExecutablePath=$executable" `
        "-dSourceDirectory=$project" -out $object (Join-Path $project "installer\Product.wxs")
    if ($LASTEXITCODE -ne 0) { throw "WiX compilation failed ($LASTEXITCODE)." }
    $finalMsi = Join-Path $output "JarvisPowerPoint-$($release.version)-x86.msi"
    $msi = Join-Path $work ([IO.Path]::GetFileName($finalMsi))
    $linkArguments = @("-nologo", "-wx", "-out", $msi, "-pdbout", (Join-Path $work "Product.wixpdb"))
    if ($SkipIceValidation) {
        Write-Warning "ICE execution was explicitly skipped. Validate MSI tables/payload now and run ICE on an approved build machine before release."
        $linkArguments += "-sval"
    }
    & $light @linkArguments $object
    if ($LASTEXITCODE -ne 0) { throw "WiX linking/ICE validation failed ($LASTEXITCODE)." }
    $installer = $null
    $database = $null
    $summary = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $database = $installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $installer, @([string]$msi, 0))
        $summary = $database.GetType().InvokeMember("SummaryInformation", "GetProperty", $null, $database, @(0))
        $packageCode = $summary.GetType().InvokeMember("Property", "GetProperty", $null, $summary, @(9))
    } finally {
        if ($summary) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($summary) }
        if ($database) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($database) }
        if ($installer) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($installer) }
    }
    $metadata = [ordered]@{
        product = "Jarvis PowerPoint"; version = $release.version
        productCode = $release.productCode; upgradeCode = $release.upgradeCode
        packageCode = $packageCode
        packageArchitecture = "x86"; executableArchitecture = "AnyCPU (.NET Framework)"
        msi = [IO.Path]::GetFileName($msi)
        sha256 = (Get-FileHash -LiteralPath $msi -Algorithm SHA256).Hash
        executableSha256 = (Get-FileHash -LiteralPath $executable -Algorithm SHA256).Hash
        executableSignature = (Get-AuthenticodeSignature -LiteralPath $executable).Status.ToString()
        signature = (Get-AuthenticodeSignature -LiteralPath $msi).Status.ToString()
        iceValidation = if ($SkipIceValidation) { "Skipped explicitly; required on approved build machine" } else { "Passed" }
    }
    $metadata | ConvertTo-Json | Set-Content -LiteralPath "$msi.json" -Encoding UTF8
    Copy-Item -LiteralPath $msi -Destination $finalMsi -Force
    Copy-Item -LiteralPath "$msi.json" -Destination "$finalMsi.json" -Force
    Write-Host "MSI: $finalMsi"
    Write-Host "SHA-256: $($metadata.sha256)"
    Write-Host "ProductCode: $($release.productCode)"
} finally {
    $env:TEMP = $oldTemp
    $env:TMP = $oldTmp
    Remove-Item -LiteralPath $work -Recurse -Force
}
