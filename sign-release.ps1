[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string[]]$Files,
    [Parameter(Mandatory = $true)][ValidatePattern('^[A-Fa-f0-9]{40}$')][string]$CertificateThumbprint,
    [Parameter(Mandatory = $true)][uri]$TimestampUrl,
    [ValidateSet("CurrentUser", "LocalMachine")][string]$CertificateStore = "CurrentUser",
    [string]$SignToolPath
)
$ErrorActionPreference = "Stop"
if ($TimestampUrl.Scheme -ne "https" -or $TimestampUrl.UserInfo -or $TimestampUrl.Fragment) {
    throw "Use the certificate provider's HTTPS RFC3161 timestamp endpoint without credentials or fragments."
}
$certificate = Get-Item -LiteralPath "Cert:\$CertificateStore\My\$CertificateThumbprint" -ErrorAction Stop
if (-not $certificate.HasPrivateKey -or $certificate.NotBefore -gt (Get-Date) -or $certificate.NotAfter -le (Get-Date)) {
    throw "The existing certificate must be valid now and have an accessible private key."
}
$codeSigning = @($certificate.EnhancedKeyUsageList | Where-Object { $_.ObjectId -eq "1.3.6.1.5.5.7.3.3" })
if ($codeSigning.Count -eq 0) { throw "The selected certificate does not have the Code Signing EKU." }
if ([string]::IsNullOrWhiteSpace($SignToolPath)) {
    $command = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if ($command) { $SignToolPath = $command.Source }
}
if (-not $SignToolPath -or -not (Test-Path -LiteralPath $SignToolPath -PathType Leaf)) {
    throw "An existing Windows SDK signtool.exe is required. Pass -SignToolPath; no SDK or trust changes are performed."
}
$paths = @($Files | ForEach-Object {
    $path = (Resolve-Path -LiteralPath $_).Path
    if ([IO.Path]::GetExtension($path) -notin @(".exe", ".msi")) { throw "Only EXE/MSI release artifacts may be signed." }
    $path
})
if (@($paths | ForEach-Object { [IO.Path]::GetExtension($_).ToLowerInvariant() } | Select-Object -Unique).Count -gt 1) {
    throw "Sign the EXE first, build the MSI from that signed EXE, then sign only the MSI in a separate invocation."
}
foreach ($path in $paths) {
    $arguments = @("sign", "/sha1", $CertificateThumbprint, "/s", "My", "/fd", "SHA256",
        "/tr", $TimestampUrl.AbsoluteUri, "/td", "SHA256")
    if ($CertificateStore -eq "LocalMachine") { $arguments += "/sm" }
    & $SignToolPath @arguments $path
    if ($LASTEXITCODE -ne 0) { throw "Signing/timestamping failed ($LASTEXITCODE): $path. Do not distribute it." }
    & $SignToolPath verify /pa /all /tw /v $path
    if ($LASTEXITCODE -ne 0) { throw "Authenticode verification failed ($LASTEXITCODE): $path. Do not distribute it." }
    $signature = Get-AuthenticodeSignature -LiteralPath $path
    if ($signature.Status -ne "Valid" -or
        $signature.SignerCertificate.Thumbprint -ne $CertificateThumbprint -or
        $null -eq $signature.TimeStamperCertificate) {
        throw "Signature trust, signer identity, or timestamp validation failed: $path. Do not distribute it."
    }
    $metadataPath = "$path.json"
    if (Test-Path -LiteralPath $metadataPath) {
        $metadata = Get-Content -LiteralPath $metadataPath -Raw | ConvertFrom-Json
        $metadata.sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $metadata.signature = $signature.Status.ToString()
        $metadata | ConvertTo-Json | Set-Content -LiteralPath $metadataPath -Encoding UTF8
    }
    Write-Host "Verified timestamped signature: $path"
    Write-Host "SHA-256: $((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash)"
}
