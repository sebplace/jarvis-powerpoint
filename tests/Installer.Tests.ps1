[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Msi,
    [Parameter(Mandatory = $true)][string]$Executable,
    [switch]$SkipAdminImage
)
$ErrorActionPreference = "Stop"
$msi = (Resolve-Path -LiteralPath $Msi).Path
$executable = (Resolve-Path -LiteralPath $Executable).Path
$project = Split-Path -Parent $PSScriptRoot
$release = Get-Content -LiteralPath (Join-Path $project "installer\release.json") -Raw | ConvertFrom-Json
$root = Join-Path $PSScriptRoot (".installer-fixtures-" + [Guid]::NewGuid().ToString("N"))
$assertions = 0
$installer = $null
$database = $null
$passed = $false
$oldTemp = $env:TEMP
$oldTmp = $env:TMP
function Assert-True([bool]$condition, [string]$message) {
    if (-not $condition) { throw $message }
    $script:assertions++
}
function Invoke-Com($object, [string]$name, [object[]]$arguments) {
    $object.GetType().InvokeMember($name, "InvokeMethod", $null, $object, $arguments)
}
function Read-Com($object, [string]$name, [object[]]$arguments) {
    $object.GetType().InvokeMember($name, "GetProperty", $null, $object, $arguments)
}
function Set-MsiProperty($session, [string]$name, [string]$value) {
    [void]$session.GetType().InvokeMember("Property", "SetProperty", $null, $session, @($name, $value))
}
function Read-Table([string]$table, [string[]]$columns) {
    $query = 'SELECT ' + (($columns | ForEach-Object { '`' + $_ + '`' }) -join ',') + ' FROM `' + $table + '`'
    $view = Invoke-Com $database "OpenView" @($query)
    try {
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        $env:TEMP = $root
        $env:TMP = $root
        [void](Invoke-Com $view "Execute" @())
        while ($record = Invoke-Com $view "Fetch" @()) {
            try {
                $row = [ordered]@{}
                for ($i = 0; $i -lt $columns.Count; $i++) {
                    $row[$columns[$i]] = Read-Com $record "StringData" @($i + 1)
                }
                [pscustomobject]$row
            } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($record) }
        }
    } finally {
        [void](Invoke-Com $view "Close" @())
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($view)
    }
}
try {
    $installer = New-Object -ComObject WindowsInstaller.Installer
    $database = Invoke-Com $installer "OpenDatabase" @($msi, 0)
    $properties = @{}
    Read-Table "Property" @("Property", "Value") | ForEach-Object { $properties[$_.Property] = $_.Value }
    Assert-True ($properties.ProductCode -eq $release.productCode) "Unexpected product code."
    Assert-True ($properties.UpgradeCode -eq $release.upgradeCode) "Upgrade family changed."
    Assert-True ($properties.ProductVersion -eq $release.version) "Version mismatch."
    Assert-True ($properties.ALLUSERS -eq "1" -and -not $properties.ContainsKey("MSIINSTALLPERUSER")) "Package is not strictly per-machine."
    Assert-True ($properties.MSIRESTARTMANAGERCONTROL -eq "Disable") "Installer could automatically shut down apps."
    Assert-True ($properties.MSIDISABLERMRESTART -eq "1") "Installer could restart Jarvis under the wrong user."
    Assert-True ($properties.REBOOT -eq "ReallySuppress") "Installer could reboot without explicit IT control."
    $summary = Read-Com $database "SummaryInformation" @(0)
    try {
        Assert-True ((Read-Com $summary "Property" @(7)) -eq "Intel;1033") "Expected x86 MSI, not native ARM64."
        Assert-True ((Read-Com $summary "Property" @(14)) -ge 500) "Windows Installer 5 requirement missing."
        $packageCode = Read-Com $summary "Property" @(9)
        Assert-True ($packageCode -match '^\{[0-9A-F-]{36}\}$') "Invalid package code."
        Assert-True ($packageCode -ne $release.productCode) "Package and product codes must differ."
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($summary) }
    if (Test-Path -LiteralPath "$msi.json") {
        $metadata = Get-Content -LiteralPath "$msi.json" -Raw | ConvertFrom-Json
        Assert-True ($metadata.packageCode -eq $packageCode) "Package code sidecar is stale."
        Assert-True ($metadata.sha256 -eq (Get-FileHash -LiteralPath $msi -Algorithm SHA256).Hash) "MSI sidecar hash is stale."
        Assert-True ($metadata.executableSha256 -eq (Get-FileHash -LiteralPath $executable -Algorithm SHA256).Hash) "Payload sidecar hash is stale."
    }
    $files = @(Read-Table "File" @("File", "Component_", "FileName", "FileSize", "Version"))
    Assert-True ($files.Count -eq 3) "Only executable, managed marker, and license belong in payload."
    $exeRow = $files | Where-Object { $_.File -eq "JarvisExe" }
    Assert-True ($exeRow.Version -eq "$($release.version).0") "Executable file version mismatch."
    Assert-True ([long]$exeRow.FileSize -eq (Get-Item -LiteralPath $executable).Length) "Executable size mismatch."
    $directories = @(Read-Table "Directory" @("Directory", "Directory_Parent", "DefaultDir"))
    $appDirectory = $directories | Where-Object { $_.Directory -eq "JarvisInstallFolder" }
    Assert-True ($appDirectory.Directory_Parent -eq "ProgramFilesFolder" -and
        $appDirectory.DefaultDir -match 'Jarvis PowerPoint$') "Application must use the protected Program Files directory."
    $components = @(Read-Table "Component" @("Component", "ComponentId", "Directory_", "Attributes", "KeyPath"))
    Assert-True ($components.Count -eq 4) "Unexpected installation components."
    Assert-True (@($components | Where-Object { ([int]$_.Attributes -band 16) -ne 0 }).Count -eq 0) "Permanent component prevents complete uninstall."
    Assert-True (@($components | Where-Object { ([int]$_.Attributes -band 256) -ne 0 }).Count -eq 0) "x86 package must not have 64-bit components."
    $shortcut = @(Read-Table "Shortcut" @("Shortcut", "Directory_", "Target", "Arguments", "WkDir"))
    Assert-True ($shortcut.Count -eq 1 -and $shortcut[0].Directory_ -eq "ProgramMenuFolder") "Missing all-users Start menu shortcut."
    Assert-True ($shortcut[0].Target -eq "Jarvis" -and $shortcut[0].WkDir -eq "JarvisInstallFolder" -and
        -not $shortcut[0].Arguments) "Advertised shortcut must target app feature without helper arguments."
    $registry = @(Read-Table "Registry" @("Registry", "Root", "Key", "Name", "Value"))
    Assert-True ($registry.Count -eq 3 -and @($registry | Where-Object { $_.Root -ne "2" }).Count -eq 0) "Installer must not write HKCU or user settings."
    Assert-True (@($registry | Where-Object { $_.Key -ne "Software\JarvisPowerPoint\Deployment" }).Count -eq 0) "Unrelated registry ownership."
    Assert-True (@($registry | Where-Object { $_.Name -eq "InstallPath" -and $_.Value -eq "[JarvisInstallFolder]" }).Count -eq 1) "Managed registration is not path-bound."
    $tables = @(Read-Table "_Tables" @("Name") | ForEach-Object { $_.Name })
    foreach ($forbidden in @("CustomAction", "ServiceInstall", "ServiceControl", "Environment", "RemoveRegistry", "MoveFile", "DuplicateFile")) {
        Assert-True ($tables -notcontains $forbidden) "Unexpected action table: $forbidden."
    }
    $removals = @(Read-Table "RemoveFile" @("FileKey", "FileName", "DirProperty", "InstallMode"))
    Assert-True ($removals.Count -eq 1 -and -not $removals[0].FileName -and
        $removals[0].DirProperty -eq "JarvisInstallFolder" -and $removals[0].InstallMode -eq "2") "Uninstall may remove only its empty installation folder, never user files."
    $upgrades = @(Read-Table "Upgrade" @("UpgradeCode", "VersionMin", "VersionMax", "Attributes", "ActionProperty"))
    $old = $upgrades | Where-Object { $_.ActionProperty -eq "WIX_UPGRADE_DETECTED" }
    $newer = $upgrades | Where-Object { $_.ActionProperty -eq "WIX_DOWNGRADE_DETECTED" }
    Assert-True ($upgrades.Count -eq 2 -and $old.VersionMax -eq $release.version -and
        ([int]$old.Attributes -band 2) -eq 0 -and ([int]$old.Attributes -band 512) -eq 0) "Upgrade must remove only strictly older versions."
    Assert-True ($newer.VersionMin -eq $release.version -and ([int]$newer.Attributes -band 2) -ne 0 -and
        ([int]$newer.Attributes -band 256) -eq 0) "Downgrade detection must detect strictly newer versions."
    $conditions = @(Read-Table "LaunchCondition" @("Condition", "Description"))
    Assert-True (@($conditions | Where-Object { $_.Condition -match 'NOT WIX_DOWNGRADE_DETECTED' }).Count -eq 1) "Newer installed version does not block downgrade."
    Assert-True (@($conditions | Where-Object { $_.Condition -match 'Installed OR.*NETFRAMEWORK48.*528040' }).Count -eq 1) "Framework requirement/maintenance exemption missing."
    Assert-True (@($conditions | Where-Object { $_.Condition -match 'ALLUSERS = 1' }).Count -eq 1) "Per-user override is not rejected."
    $search = @(Read-Table "RegLocator" @("Signature_", "Root", "Key", "Name", "Type"))
    Assert-True ($search.Count -eq 1 -and $search[0].Root -eq "2" -and $search[0].Name -eq "Release" -and
        ([int]$search[0].Type -band 16) -eq 0) "Framework discovery must use the .NET 32-bit registry view, without requiring PowerPoint."
    $sequence = @{}
    Read-Table "InstallExecuteSequence" @("Action", "Condition", "Sequence") |
        ForEach-Object { $sequence[$_.Action] = [int]$_.Sequence }
    Assert-True ($sequence.RemoveExistingProducts -gt $sequence.InstallInitialize -and
        $sequence.RemoveExistingProducts -lt $sequence.ProcessComponents) "Major upgrade removal must be inside rollback transaction."
    foreach ($action in @("InstallFiles", "RemoveFiles", "CreateShortcuts", "RemoveShortcuts", "WriteRegistryValues", "RemoveRegistryValues")) {
        Assert-True ($sequence.ContainsKey($action)) "Repair/uninstall standard action missing: $action."
    }
    Assert-True (-not $sequence.ContainsKey("ForceReboot") -and -not $sequence.ContainsKey("ScheduleReboot")) "Forced reboot action present."
    $media = @(Read-Table "Media" @("DiskId", "Cabinet"))
    Assert-True ($media.Count -eq 1 -and $media[0].Cabinet.StartsWith("#")) "MSI is not a self-contained embedded-cabinet package."

    # OpenPackage(ignoreMachineState) and condition evaluation run no installation actions.
    $session = Invoke-Com $installer "OpenPackage" @($msi, 1)
    try {
        $frameworkCondition = ($conditions | Where-Object { $_.Condition -match 'Installed OR.*NETFRAMEWORK48' }).Condition
        $downgradeCondition = ($conditions | Where-Object { $_.Condition -match 'NOT WIX_DOWNGRADE_DETECTED' }).Condition
        Set-MsiProperty $session "Installed" ""
        foreach ($case in @(
            @{ Release = ""; Expected = 0 },
            @{ Release = "#378389"; Expected = 0 },
            @{ Release = "#461808"; Expected = 0 },
            @{ Release = "#528039"; Expected = 0 },
            @{ Release = "#528040"; Expected = 1 },
            @{ Release = "#533320"; Expected = 1 }
        )) {
            Set-MsiProperty $session "NETFRAMEWORK48" $case.Release
            Assert-True ((Invoke-Com $session "EvaluateCondition" @($frameworkCondition)) -eq $case.Expected) "Framework launch condition failed for '$($case.Release)'."
        }
        Set-MsiProperty $session "Installed" "1"
        Set-MsiProperty $session "NETFRAMEWORK48" ""
        Assert-True ((Invoke-Com $session "EvaluateCondition" @($frameworkCondition)) -eq 1) "Repair/uninstall is blocked by a missing prerequisite."
        Set-MsiProperty $session "WIX_DOWNGRADE_DETECTED" ""
        Assert-True ((Invoke-Com $session "EvaluateCondition" @($downgradeCondition)) -eq 1) "Fresh installation is blocked as a downgrade."
        Set-MsiProperty $session "WIX_DOWNGRADE_DETECTED" "{11111111-1111-1111-1111-111111111111}"
        Assert-True ((Invoke-Com $session "EvaluateCondition" @($downgradeCondition)) -eq 0) "A newer product does not block downgrade."
        Set-MsiProperty $session "ALLUSERS" ""
        Assert-True ((Invoke-Com $session "EvaluateCondition" @("ALLUSERS = 1")) -eq 0) "Per-user override is allowed."
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($session) }

    if (-not $SkipAdminImage) {
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        $image = Join-Path $root "admin-image"
        $log = Join-Path $root "admin-image.log"
        $stateBefore = Read-Com $installer "ProductState" @($release.productCode)
        $hashBefore = (Get-FileHash -LiteralPath $executable -Algorithm SHA256).Hash
        $msiHashBefore = (Get-FileHash -LiteralPath $msi -Algorithm SHA256).Hash
        $process = New-Object Diagnostics.Process
        $process.StartInfo = New-Object Diagnostics.ProcessStartInfo
        $process.StartInfo.FileName = Join-Path $env:WINDIR "System32\msiexec.exe"
        $process.StartInfo.Arguments = "/a `"$msi`" /qn /norestart TARGETDIR=`"$image`" /l*v `"$log`""
        $process.StartInfo.UseShellExecute = $false
        $process.StartInfo.EnvironmentVariables["TEMP"] = $root
        $process.StartInfo.EnvironmentVariables["TMP"] = $root
        try {
            Assert-True ($process.Start()) "Administrative image process could not start."
            if (-not $process.WaitForExit(120000)) {
                $process.Kill()
                $process.WaitForExit()
                throw "Administrative image creation timed out; stopped only test process $($process.Id). Inspect $log."
            }
            Assert-True ($process.ExitCode -eq 0) "Administrative extraction failed: exit $($process.ExitCode); $log"
        } finally { $process.Dispose() }
        Assert-True ((Read-Com $installer "ProductState" @($release.productCode)) -eq $stateBefore) "Administrative extraction changed real installed product state."
        Assert-True ((Get-FileHash -LiteralPath $executable -Algorithm SHA256).Hash -eq $hashBefore) "Source application changed during extraction."
        Assert-True ((Get-FileHash -LiteralPath $msi -Algorithm SHA256).Hash -eq $msiHashBefore) "Source MSI changed during extraction."
        $extracted = @(Get-ChildItem -LiteralPath $image -Filter "JarvisPowerPoint.exe" -Recurse)
        Assert-True ($extracted.Count -eq 1) "Administrative image must contain exactly one executable."
        Assert-True ((Get-FileHash -LiteralPath $extracted[0].FullName -Algorithm SHA256).Hash -eq $hashBefore) "MSI payload differs from selected candidate."
        Assert-True (Test-Path -LiteralPath (Join-Path $extracted[0].DirectoryName "JarvisPowerPoint.managed")) "Managed marker missing from administrative image."
        Assert-True (Test-Path -LiteralPath (Join-Path $extracted[0].DirectoryName "LICENSE")) "License missing from administrative image."
        foreach ($payload in @("JarvisPowerPoint.managed", "LICENSE")) {
            $source = if ($payload -eq "LICENSE") { Join-Path $project $payload } else { Join-Path $project "installer\$payload" }
            Assert-True ((Get-FileHash -LiteralPath (Join-Path $extracted[0].DirectoryName $payload) -Algorithm SHA256).Hash -eq
                (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash) "Extracted $payload differs from its source."
        }
    }
    Write-Host "Passed $assertions MSI metadata, offline lifecycle-contract, and administrative-image assertions."
    Write-Host "No real installation, repair, upgrade, or uninstall was performed; validate these on an isolated IT test VM."
    $passed = $true
} finally {
    if ($database) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($database) }
    if ($installer) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($installer) }
    $env:TEMP = $oldTemp
    $env:TMP = $oldTmp
    if ($passed -and (Test-Path -LiteralPath $root)) { Remove-Item -LiteralPath $root -Recurse -Force }
    elseif (Test-Path -LiteralPath $root) { Write-Warning "Failed test fixtures/logs retained: $root" }
}
