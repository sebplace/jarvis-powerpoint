[CmdletBinding()]
param(
    [string]$OutputDirectory
)

$ErrorActionPreference = "Stop"
$projectDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $projectDirectory "bin"
}
$outputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
$outputPath = Join-Path $outputDirectory "JarvisPowerPoint.exe"
$sourcePaths = Get-ChildItem -LiteralPath $projectDirectory -Filter "*.cs" |
    Select-Object -ExpandProperty FullName

$compilerCandidates = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)
$compiler = $compilerCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
$speechAssemblyCandidates = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\WPF\System.Speech.dll",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\WPF\System.Speech.dll"
)
$speechAssembly = $speechAssemblyCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
$powerPointInterop = Get-ChildItem `
    "$env:WINDIR\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint" `
    -Filter "Microsoft.Office.Interop.PowerPoint.dll" `
    -Recurse `
    -ErrorAction SilentlyContinue |
    Select-Object -First 1 -ExpandProperty FullName
$officeInterop = Get-ChildItem `
    "$env:WINDIR\assembly\GAC_MSIL\office" `
    -Filter "OFFICE.DLL" `
    -Recurse `
    -ErrorAction SilentlyContinue |
    Select-Object -First 1 -ExpandProperty FullName

if (-not $compiler) {
    throw "Le compilateur .NET Framework de Windows est introuvable."
}

if (-not $speechAssembly) {
    throw "La bibliothèque de reconnaissance vocale System.Speech est introuvable."
}

if (-not $powerPointInterop -or -not $officeInterop) {
    throw "Les bibliothèques d'intégration PowerPoint sont introuvables."
}

New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null

& $compiler `
    /nologo `
    /target:winexe `
    /platform:anycpu `
    /optimize+ `
    "/win32manifest:$projectDirectory\app.manifest" `
    "/out:$outputPath" `
    /reference:System.dll `
    /reference:System.Core.dll `
    /reference:System.Drawing.dll `
    /reference:System.Xml.dll `
    /reference:System.Xml.Linq.dll `
    /reference:System.Runtime.Serialization.dll `
    "/reference:$speechAssembly" `
    /reference:System.Windows.Forms.dll `
    "/reference:$officeInterop" `
    "/reference:$powerPointInterop" `
    $sourcePaths

if ($LASTEXITCODE -ne 0) {
    throw "La compilation a échoué avec le code $LASTEXITCODE."
}

Write-Host "Application créée : $outputPath"
