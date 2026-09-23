#Requires -Version 7.0
[CmdletBinding()]
param([switch]$Build, [switch]$TrustCertificate)
$ErrorActionPreference = 'Stop'
$repoPath = Split-Path -Parent $PSScriptRoot
$secretDir = Join-Path $env:LOCALAPPDATA 'WordAiSecure\secrets'
if (-not (Test-Path -LiteralPath $secretDir)) {
    New-Item -ItemType Directory -Path $secretDir -Force | Out-Null
}
# Protect the directory BEFORE writing secrets. No chmod-on-Windows assumption.
$ownerSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
$directoryInfo = [IO.DirectoryInfo]::new($secretDir)
$acl = [IO.FileSystemAclExtensions]::GetAccessControl($directoryInfo, [Security.AccessControl.AccessControlSections]::Access)
$acl.SetAccessRuleProtection($true, $false)
foreach ($existingRule in $acl.GetAccessRules($true, $false, [Security.Principal.SecurityIdentifier])) {
    $acl.RemoveAccessRuleSpecific($existingRule)
}
foreach ($sid in @($ownerSid, [System.Security.Principal.SecurityIdentifier]::new('S-1-5-18'))) {
    $rule = [System.Security.AccessControl.FileSystemAccessRule]::new($sid, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
    $acl.AddAccessRule($rule)
}
[IO.FileSystemAclExtensions]::SetAccessControl($directoryInfo, $acl)
$tokenPath = Join-Path $secretDir 'pairing-token'
if (-not (Test-Path -LiteralPath $tokenPath)) {
    [IO.File]::WriteAllText($tokenPath, [Convert]::ToBase64String([Security.Cryptography.RandomNumberGenerator]::GetBytes(32)))
}
$certPath = Join-Path $secretDir 'localhost.crt'
$keyPath = Join-Path $secretDir 'localhost.key'
if (-not (Test-Path -LiteralPath $certPath) -and -not (Test-Path -LiteralPath $keyPath)) {
    $rsa = [Security.Cryptography.RSA]::Create(3072)
    $request = [Security.Cryptography.X509Certificates.CertificateRequest]::new('CN=Word AI Secure localhost', $rsa, 'SHA256', [Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $san = [Security.Cryptography.X509Certificates.SubjectAlternativeNameBuilder]::new()
    $san.AddDnsName('localhost')
    $san.AddIpAddress([Net.IPAddress]::Loopback)
    $request.CertificateExtensions.Add($san.Build())
    $request.CertificateExtensions.Add([Security.Cryptography.X509Certificates.X509BasicConstraintsExtension]::new($false, $false, 0, $true))
    $oids = [Security.Cryptography.OidCollection]::new()
    $null = $oids.Add([Security.Cryptography.Oid]::new('1.3.6.1.5.5.7.3.1'))
    $request.CertificateExtensions.Add([Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]::new($oids, $false))
    $cert = $request.CreateSelfSigned([DateTimeOffset]::Now.AddMinutes(-5), [DateTimeOffset]::Now.AddMonths(6))
    [IO.File]::WriteAllText($certPath, $cert.ExportCertificatePem())
    [IO.File]::WriteAllText($keyPath, $rsa.ExportPkcs8PrivateKeyPem())
    $rsa.Dispose()
}
if (-not (Test-Path -LiteralPath $certPath) -or -not (Test-Path -LiteralPath $keyPath)) {
    throw 'Incomplete certificate material. Inspect the secret directory; no files were overwritten.'
}
$publicCert = [Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPem([IO.File]::ReadAllText($certPath))
if ($publicCert.NotAfter -le (Get-Date)) { throw 'Local TLS certificate expired. Renew it before starting.' }
if ($TrustCertificate -and -not (Test-Path "Cert:\CurrentUser\Root\$($publicCert.Thumbprint)")) {
    # Trust only this non-CA localhost leaf, only for the current Windows user.
    $store = [Security.Cryptography.X509Certificates.X509Store]::new('Root', 'CurrentUser')
    $store.Open('ReadWrite')
    try { $store.Add($publicCert) } finally { $store.Close() }
}
# Rancher/WSL cannot reliably read an ACL-restricted Windows bind directory.
# Copy only these three files over STDIN into a private Linux volume. Values
# never appear in command arguments, environment, logs or a Docker image layer.
$copyScript = @'
import json, os, pathlib, sys
data = json.load(sys.stdin)
root = pathlib.Path('/secrets')
os.chown(root, 0, 0)
for name in ('pairing-token', 'localhost.crt', 'localhost.key'):
    path = root / name
    if path.exists():
        os.chown(path, 0, 0)
    path.write_text(data[name], encoding='utf-8')
    path.chmod(0o400)
    os.chown(path, 10001, 10001)
root.chmod(0o700)
os.chown(root, 10001, 10001)
'@
$secretPayload = @{
    'pairing-token' = [IO.File]::ReadAllText($tokenPath)
    'localhost.crt' = [IO.File]::ReadAllText($certPath)
    'localhost.key' = [IO.File]::ReadAllText($keyPath)
} | ConvertTo-Json -Compress
$secretPayload | & docker run --rm -i --network none --read-only --cap-drop ALL --cap-add CHOWN --cap-add DAC_OVERRIDE --mount type=volume,source=word-ai-secure-secrets,target=/secrets --entrypoint python python:3.12-slim@sha256:78387bc3881b8273120a12ebe6c1ab22b018ccc2c9adf565ae1ac9b536e184ea -c $copyScript
$secretPayload = $null
if ($LASTEXITCODE -ne 0) { throw 'Provisioning the protected Docker secrets volume failed.' }
$composeArgs = @('compose', '-f', (Join-Path $repoPath 'compose.secure.yaml'), 'up', '-d')
if ($Build) { $composeArgs += '--build' }
& docker @composeArgs
if ($LASTEXITCODE -ne 0) { throw 'Docker compose failed.' }
Write-Output 'Word AI secure-live started on https://localhost:3100.'
Write-Output "Pairing token is protected in $tokenPath (not printed)."
if (-not (Test-Path "Cert:\CurrentUser\Root\$($publicCert.Thumbprint)")) {
    Write-Output 'Word still needs certificate trust. Run this script with -TrustCertificate and accept the Windows localhost certificate prompt.'
}
