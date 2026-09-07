[CmdletBinding()]
param()
$ErrorActionPreference = "Stop"
$release = Get-Content -LiteralPath (Join-Path $PSScriptRoot "release.json") -Raw | ConvertFrom-Json
$tools = Join-Path $PSScriptRoot ".tools"
$destination = Join-Path $tools "wix3"
New-Item -ItemType Directory -Path $tools -Force | Out-Null
$archive = Join-Path $tools "wix314-binaries.zip"
if (-not (Test-Path -LiteralPath $archive)) {
    & curl.exe --fail --location --retry 2 --output $archive $release.wixArchiveUrl
    if ($LASTEXITCODE -ne 0) { throw "Could not download official WiX archive. No tools were installed." }
}
if ((Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash -ne $release.wixArchiveSha256) {
    throw "WiX archive SHA-256 does not match the pinned release. Refusing extraction."
}
Expand-Archive -LiteralPath $archive -DestinationPath $destination -Force
foreach ($tool in @("candle.exe", "light.exe")) {
    if ([Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $destination $tool)).FileVersion -ne $release.wixVersion) {
        throw "Restored WiX tool version differs from the pinned release: $tool."
    }
}
Write-Host "Restored WiX locally: $destination"
