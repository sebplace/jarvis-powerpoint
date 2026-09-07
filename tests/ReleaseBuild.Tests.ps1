[CmdletBinding()]
param()
$ErrorActionPreference = "Stop"
$project = Split-Path -Parent $PSScriptRoot
$root = Join-Path $PSScriptRoot (".release-fixtures-" + [Guid]::NewGuid().ToString("N"))
$compiler = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe"
$build = Join-Path $project "build-msi.ps1"
$sign = Join-Path $project "sign-release.ps1"
$assertions = 0
$oldTemp = $env:TEMP
$oldTmp = $env:TMP
function Assert-Rejected([scriptblock]$action, [string]$expected) {
    $failure = $null
    try { & $action | Out-Null } catch { $failure = $_.Exception.Message }
    if (-not $failure -or $failure -notmatch $expected) {
        throw "Expected rejection '$expected'; observed '$failure'."
    }
    $script:assertions++
}
try {
    New-Item -ItemType Directory -Path $root | Out-Null
    $env:TEMP = $root
    $env:TMP = $root
    foreach ($case in @(
        @{ Name = "x86"; Platform = "x86"; Assembly = "1.5.0.0"; File = "1.5.0.0"; Error = "not IL-only AnyCPU" },
        @{ Name = "preferred32"; Platform = "anycpu32bitpreferred"; Assembly = "1.5.0.0"; File = "1.5.0.0"; Error = "not IL-only AnyCPU" },
        @{ Name = "old"; Platform = "anycpu"; Assembly = "1.4.0.0"; File = "1.4.0.0"; Error = "Expected release executable" },
        @{ Name = "assembly"; Platform = "anycpu"; Assembly = "1.4.0.0"; File = "1.5.0.0"; Error = "assembly version must also match" }
    )) {
        $folder = Join-Path $root $case.Name
        New-Item -ItemType Directory -Path $folder | Out-Null
        $source = Join-Path $folder "Fixture.cs"
        $exe = Join-Path $folder "JarvisPowerPoint.exe"
        ('using System.Reflection; [assembly: AssemblyVersion("' + $case.Assembly +
            '")] [assembly: AssemblyFileVersion("' + $case.File +
            '")] class Fixture { static void Main() {} }') | Set-Content -LiteralPath $source -Encoding ASCII
        & $compiler /nologo /langversion:5 /target:exe "/platform:$($case.Platform)" "/out:$exe" $source
        if ($LASTEXITCODE -ne 0) { throw "Release guard fixture compilation failed." }
        Assert-Rejected { & $build -Executable $exe -OutputDirectory (Join-Path $folder "output") } $case.Error
        if (Test-Path -LiteralPath (Join-Path $folder "output")) { throw "Rejected payload created installer output." }
        $script:assertions++
    }
    foreach ($url in @("http://timestamp.example.invalid", "https://user@timestamp.example.invalid", "https://timestamp.example.invalid/#fragment")) {
        Assert-Rejected {
            & $sign -Files $exe -CertificateThumbprint ("0" * 40) -TimestampUrl $url
        } "HTTPS RFC3161"
    }
    Write-Host "Passed $assertions release architecture/version and signing-preflight assertions; no signing or network requests performed."
} finally {
    $env:TEMP = $oldTemp
    $env:TMP = $oldTmp
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}
