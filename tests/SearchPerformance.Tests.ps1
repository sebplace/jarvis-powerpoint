[CmdletBinding()]
param([string]$Executable)

$ErrorActionPreference = "Stop"
$arguments = @{ SearchBenchmark = $true }
if ([string]::IsNullOrWhiteSpace($Executable)) {
    $arguments.SourceOnly = $true
} else {
    $arguments.Executable = $Executable
}

# Reuses the real C#5 controller/session implementation with synthetic in-memory
# slides. It never launches PowerPoint, captures audio, or touches an open deck.
& (Join-Path $PSScriptRoot "PresentationSession.Tests.ps1") @arguments
