[CmdletBinding()]
param([string]$Executable = (Join-Path (Split-Path -Parent $PSScriptRoot) "bin\JarvisPowerPoint.exe"))

$ErrorActionPreference = "Stop"
$projectDirectory = Split-Path -Parent $PSScriptRoot
$executable = [IO.Path]::GetFullPath($Executable)
$assembly = [Reflection.Assembly]::LoadFrom($executable)
$shortcutType = $assembly.GetType("JarvisPowerPoint.StartMenuShortcut", $true)
$id = [Guid]::NewGuid().ToString("N")
$tempDirectory = Join-Path $PSScriptRoot ".shortcut-fixtures-$id"
$registrySubKey = "Software\JarvisPowerPoint.Tests\$id"
$settingsKey = "HKEY_CURRENT_USER\$registrySubKey"
$script:promptCount = 0
$assertions = 0
$shell = $null

function Assert-True([bool]$condition, [string]$message) {
    if (-not $condition) { throw $message }
    $script:assertions++
}

function New-ShortcutManager([string]$folder, [string]$target, [string]$key) {
    [Activator]::CreateInstance($shortcutType, [object[]]@($folder, $target, $key))
}

function Assert-Shortcut([string]$path, [string]$target) {
    $link = $shell.CreateShortcut($path)
    try {
        Assert-True ($link.TargetPath -eq $target) "Wrong shortcut target."
        Assert-True ($link.WorkingDirectory -eq (Split-Path -Parent $target)) "Wrong working directory."
        Assert-True ([string]::IsNullOrEmpty($link.Arguments)) "Shortcut unexpectedly passes CLI arguments."
        Assert-True ($link.IconLocation -eq "$target,0") "Wrong shortcut icon."
    } finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($link)
    }
}

$accept = [Func[bool]]{
    $script:promptCount++
    return $true
}
$decline = [Func[bool]]{
    $script:promptCount++
    return $false
}
$unexpectedPrompt = [Func[bool]]{ throw "An unexpected prompt was displayed." }

try {
    # Isolate all filesystem and HKCU changes from the real user's shortcut/settings.
    $sourceFolder = Join-Path $tempDirectory ("Portable app " + [char]0x00E9)
    New-Item -ItemType Directory -Path $sourceFolder -Force | Out-Null
    $target = Join-Path $sourceFolder "JarvisPowerPoint.exe"
    Copy-Item -LiteralPath $executable -Destination $target
    $shell = New-Object -ComObject WScript.Shell
    $programs = Join-Path $tempDirectory "Start Menu\Programs"
    $manager = New-ShortcutManager $programs $target $settingsKey
    [Microsoft.Win32.Registry]::SetValue($settingsKey, "Language", "en-US")

    $result = $manager.Configure($decline, $false)
    Assert-True ($result.ToString() -eq "Declined") "First-launch refusal was not honored."
    Assert-True ($script:promptCount -eq 1) "First launch must prompt exactly once."
    Assert-True (-not (Test-Path -LiteralPath $manager.ShortcutPath)) "Refusing must not create a shortcut."
    Assert-True ([Microsoft.Win32.Registry]::GetValue($settingsKey, "StartMenuShortcutPrompted", 0) -eq 1) "Refusal was not persisted."
    $result = $manager.Configure($unexpectedPrompt, $false)
    Assert-True ($result.ToString() -eq "NotNeeded") "Refusal must suppress subsequent startup prompts."
    # A new object simulates the state after restarting the application.
    $manager = New-ShortcutManager $programs $target $settingsKey
    Assert-True ($manager.Configure($unexpectedPrompt, $false).ToString() -eq "NotNeeded") "Refusal was lost on restart."

    $result = $manager.Configure($accept, $true)
    Assert-True ($result.ToString() -eq "Created") "Manual creation after refusal failed."
    Assert-True (Test-Path -LiteralPath $manager.ShortcutPath) "Shortcut is missing."
    Assert-Shortcut $manager.ShortcutPath $target
    Assert-True ([Microsoft.Win32.Registry]::GetValue($settingsKey, "Language", "") -eq "en-US") "Existing language setting was modified."
    Assert-True ($manager.Configure($unexpectedPrompt, $false).ToString() -eq "NotNeeded") "Existing shortcut triggers a startup prompt."

    $newKey = "$settingsKey\Fresh"
    $fresh = New-ShortcutManager (Join-Path $tempDirectory "Fresh Programs") $target $newKey
    Assert-True ($fresh.Configure($accept, $false).ToString() -eq "Created") "Accepting on first launch failed."
    Assert-Shortcut $fresh.ShortcutPath $target
    Assert-True ([Microsoft.Win32.Registry]::GetValue($newKey, "StartMenuShortcutPrompted", 0) -eq 1) "Acceptance was not persisted."

    $existing = New-ShortcutManager $programs $target "$settingsKey\NoPreference"
    $before = (Get-FileHash -LiteralPath $manager.ShortcutPath).Hash
    Assert-True ($existing.Configure($unexpectedPrompt, $false).ToString() -eq "NotNeeded") "Pre-existing shortcut must not prompt."
    Assert-True ((Get-FileHash -LiteralPath $manager.ShortcutPath).Hash -eq $before) "Startup modified a pre-existing shortcut."

    $movedFolder = Join-Path $tempDirectory "Moved app"
    New-Item -ItemType Directory -Path $movedFolder | Out-Null
    $movedTarget = Join-Path $movedFolder "JarvisPowerPoint.exe"
    Move-Item -LiteralPath $target -Destination $movedTarget
    $moved = New-ShortcutManager $programs $movedTarget $settingsKey
    Assert-True ($moved.Configure($decline, $true).ToString() -eq "Declined") "Manual refusal was ignored."
    Assert-True ((Get-FileHash -LiteralPath $manager.ShortcutPath).Hash -eq $before) "Manual refusal changed the shortcut."
    Assert-True ($moved.Configure($accept, $true).ToString() -eq "Created") "Manual repair after moving failed."
    Assert-Shortcut $moved.ShortcutPath $movedTarget
    Assert-True ((Get-ChildItem -LiteralPath $programs -Filter "*.lnk").Count -eq 1) "Repair created a duplicate shortcut."

    $blockedFolder = Join-Path $tempDirectory "Not a directory"
    New-Item -ItemType File -Path $blockedFolder | Out-Null
    $blockedKey = "$settingsKey\Blocked"
    $blocked = New-ShortcutManager $blockedFolder $movedTarget $blockedKey
    $failed = $false
    try { [void]$blocked.Configure($accept, $false) }
    catch {
        if ($_.Exception.GetBaseException() -isnot [IO.IOException]) { throw }
        $failed = $true
    }
    Assert-True $failed "A filesystem failure must propagate."
    Assert-True ($null -eq [Microsoft.Win32.Registry]::GetValue($blockedKey, "StartMenuShortcutPrompted", $null)) "Failed creation incorrectly marked the prompt handled."
    Remove-Item -LiteralPath $blockedFolder
    Assert-True ($blocked.Configure($accept, $false).ToString() -eq "Created") "Retry after failure did not offer creation."

    $missingKey = "$settingsKey\Missing"
    $missing = New-ShortcutManager (Join-Path $tempDirectory "Missing Programs") $target $missingKey
    $failed = $false
    try { [void]$missing.Configure($accept, $false) }
    catch {
        if ($_.Exception.GetBaseException() -isnot [IO.FileNotFoundException]) { throw }
        $failed = $true
    }
    Assert-True $failed "Missing executable must fail, not create a broken shortcut."
    Assert-True (-not (Test-Path -LiteralPath $missing.ShortcutPath)) "Missing executable created a broken shortcut."
    Assert-True ($null -eq [Microsoft.Win32.Registry]::GetValue($missingKey, "StartMenuShortcutPrompted", $null)) "Missing executable failure incorrectly marked the prompt handled."

    $managedKey = "$settingsKey\Managed"
    $managed = New-ShortcutManager (Join-Path $tempDirectory "Managed Programs") $movedTarget $managedKey
    Set-Content -LiteralPath (Join-Path $movedFolder "JarvisPowerPoint.managed") -Value "managed fixture"
    Assert-True ($managed.Configure($unexpectedPrompt, $false).ToString() -eq "NotNeeded") "Managed startup must not prompt."
    Assert-True ($managed.Configure($unexpectedPrompt, $true).ToString() -eq "NotNeeded") "Managed menu must not mutate shortcuts."
    Assert-True (-not (Test-Path -LiteralPath $managed.ShortcutPath)) "Managed install created a per-user shortcut."
    Assert-True ($null -eq [Microsoft.Win32.Registry]::GetValue($managedKey, "StartMenuShortcutPrompted", $null)) "Managed guard mutated user preference."

    $alias = Join-Path $tempDirectory "Directory alias"
    New-Item -ItemType Junction -Path $alias -Target $movedFolder | Out-Null
    try {
        $managedType = $assembly.GetType("JarvisPowerPoint.ManagedDeployment", $true)
        $matches = $managedType.GetMethod("MatchesInstallPath", [Reflection.BindingFlags]"Static,NonPublic")
        Assert-True ($matches.Invoke($null, @([string]$alias, [string]$movedFolder))) "Directory alias evades path-bound registration."
        Assert-True ($managedType.GetMethod("IsManagedExecutable").Invoke($null, @([string](Join-Path $alias "JarvisPowerPoint.exe")))) "Directory alias evades managed marker."
        $aliasManager = New-ShortcutManager (Join-Path $tempDirectory "Alias Programs") (Join-Path $alias "JarvisPowerPoint.exe") "$managedKey\Alias"
        Assert-True ($aliasManager.Configure($unexpectedPrompt, $true).ToString() -eq "NotNeeded") "Managed directory alias mutated shortcuts."
    } finally { [IO.Directory]::Delete($alias) }

    Write-Host "Passed $assertions shortcut assertions."
} finally {
    if ($shell) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell) }
    [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($registrySubKey, $false)
    if (Test-Path -LiteralPath $tempDirectory) { Remove-Item -LiteralPath $tempDirectory -Recurse -Force }
}
