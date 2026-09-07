[CmdletBinding()]
param(
    [string]$Executable = (Join-Path (Split-Path -Parent $PSScriptRoot) "bin\JarvisPowerPoint.exe"),
    [switch]$CheckGithub
)

$ErrorActionPreference = "Stop"
if (!$CheckGithub) { throw "This opt-in check contacts the public GitHub API. Pass -CheckGithub explicitly." }
$assembly = [Reflection.Assembly]::LoadFrom([IO.Path]::GetFullPath($Executable))
$manager = $assembly.GetType("JarvisPowerPoint.UpdateManager", $true)
$method = $manager.GetMethod("Check", [type[]]@([Threading.CancellationToken]))
$cancellation = New-Object Threading.CancellationTokenSource
try {
    $cancellation.CancelAfter(45000)
    try { $release = $method.Invoke($null, [object[]]@($cancellation.Token)) }
    catch [Reflection.TargetInvocationException] { throw $_.Exception.InnerException }
    if ($null -eq $release) {
        Write-Host "Anonymous GitHub check succeeded: no newer stable release. No asset downloaded."
    } else {
        Write-Host "Anonymous GitHub check succeeded: newer stable release found. No asset downloaded."
    }
} finally { $cancellation.Dispose() }
