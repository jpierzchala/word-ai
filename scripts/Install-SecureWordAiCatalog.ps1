#Requires -Version 7.0
[CmdletBinding()]
param([switch]$PrepareOnly, [switch]$RequestElevation)
# RequestElevation is retained for compatibility with the first installer.
# No elevation, new SMB shares, firewall changes or machine-wide trust changes.
$ErrorActionPreference = 'Stop'
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$catalogId = '{d9a54e63-a71e-4cc5-87dd-dc6555d60a5b}'
$catalogPath = Join-Path $env:LOCALAPPDATA 'WordAiSecure\catalog'
$source = Join-Path (Split-Path -Parent $PSScriptRoot) 'secure-addin\manifest.xml'
$target = Join-Path $catalogPath 'manifest.xml'
if (-not (Test-Path -LiteralPath $catalogPath)) {
    New-Item -ItemType Directory -Path $catalogPath -Force | Out-Null
}
$directory = Get-Item -LiteralPath $catalogPath
if (($directory.Attributes -band [IO.FileAttributes]::ReparsePoint) -or
    ($directory.Parent.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
    throw 'Catalog and its parent must not be junctions or symbolic links.'
}
if ((Get-Acl -LiteralPath $catalogPath).GetOwner([Security.Principal.SecurityIdentifier]).Value -ne $identity.User.Value) {
    throw 'Catalog must belong to the current user.'
}
if (Test-Path -LiteralPath $target) {
    if ((Get-FileHash -LiteralPath $target).Hash -ne (Get-FileHash -LiteralPath $source).Hash) {
        throw 'Published manifest differs. Review the existing file before updating it.'
    }
} else { Copy-Item -LiteralPath $source -Destination $target }
$items = @(Get-ChildItem -LiteralPath $catalogPath -Force)
if ($items.Count -ne 1 -or $items[0].Name -ne 'manifest.xml' -or
    ($items[0].Attributes -band [IO.FileAttributes]::ReparsePoint)) {
    throw 'The catalog must contain only the regular manifest.xml file.'
}

# MSIX can redirect AppData writes into Packages/<family>/LocalCache/Local.
# Resolve the actual file handle, not the logical path seen by this process.
if (-not ('WordAiCatalogPaths' -as [type])) {
    Add-Type -TypeDefinition @'
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
public static class WordAiCatalogPaths {
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern uint GetFinalPathNameByHandle(SafeFileHandle handle, StringBuilder path, uint length, uint flags);
}
'@
}
$handle = [IO.File]::OpenHandle($target)
try {
    $buffer = [Text.StringBuilder]::new(4096)
    $length = [WordAiCatalogPaths]::GetFinalPathNameByHandle($handle, $buffer, 4096, 0)
    if ($length -eq 0 -or $length -ge 4096) { throw 'Could not resolve the physical manifest path.' }
    $physicalFile = $buffer.ToString()
} finally { $handle.Dispose() }
if (-not $physicalFile.StartsWith('\\?\')) { throw 'Unexpected final file path.' }
$physicalFile = $physicalFile.Substring(4)
if ($physicalFile -notmatch '^[A-Za-z]:\\' -or -not $physicalFile.EndsWith('\WordAiSecure\catalog\manifest.xml', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Only the current local WordAiSecure manifest is supported.'
}
$drive = $physicalFile.Substring(0, 1)
$shareName = $drive + '$'
$share = Get-SmbShare -Name $shareName -ErrorAction Stop
if ($share.Path.TrimEnd('\') -ne ($drive + ':')) { throw 'Existing drive share points to an unexpected location.' }
$uncFile = '\\{0}\{1}\{2}' -f $env:COMPUTERNAME, $shareName, $physicalFile.Substring(3)
if ((Get-FileHash -LiteralPath $uncFile -ErrorAction Stop).Hash -ne (Get-FileHash -LiteralPath $target).Hash) {
    throw 'Current user cannot read the exact manifest through the existing share.'
}
$url = Split-Path -Parent $uncFile

# MSIX can virtualize HKCU as well. StdRegProv reads/writes the real user's hive
# outside the package overlay, under the caller's existing permissions.
# Never target another SID or HKEY_LOCAL_MACHINE.
$registryPath = $identity.User.Value + '\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\' + $catalogId
$base = @{hDefKey=[uint32]2147483651; sSubKeyName=$registryPath} # HKEY_USERS/current SID
function Invoke-CatalogRegistry([string]$Method, [hashtable]$Values = @{}) {
    $arguments = @{} + $base + $Values
    Invoke-CimMethod -Namespace root/default -ClassName StdRegProv -MethodName $Method -Arguments $arguments
}
$existingUrl = Invoke-CatalogRegistry 'GetStringValue' @{sValueName='Url'}
if ($existingUrl.ReturnValue -notin @(0, 2)) { throw "Cannot inspect the real user catalog: $($existingUrl.ReturnValue)" }
if ($existingUrl.ReturnValue -eq 0 -and $existingUrl.sValue -ne $url) {
    throw 'Catalog ID already belongs to another location; it was not changed.'
}
if ($PrepareOnly) { Write-Output "Verified physical manifest, existing share and real user hive. Catalog URL: $url"; exit 0 }
if ((Invoke-CatalogRegistry 'CreateKey').ReturnValue -ne 0) { throw 'Cannot create the catalog in the real user hive.' }
foreach ($pair in @(@('Id', $catalogId), @('Url', $url))) {
    if ((Invoke-CatalogRegistry 'SetStringValue' @{sValueName=$pair[0];sValue=$pair[1]}).ReturnValue -ne 0) { throw "Cannot write catalog $($pair[0])." }
}
if ((Invoke-CatalogRegistry 'SetDWORDValue' @{sValueName='Flags';uValue=[uint32]1}).ReturnValue -ne 0) { throw 'Cannot enable Show in Menu.' }
$verified = Invoke-CatalogRegistry 'GetStringValue' @{sValueName='Url'}
$flags = Invoke-CatalogRegistry 'GetDWORDValue' @{sValueName='Flags'}
if ($verified.ReturnValue -ne 0 -or $verified.sValue -ne $url -or $flags.uValue -ne 1) { throw 'Catalog read-back verification failed.' }
Write-Output "Installed in the real profile of $($identity.Name). Catalog: $url"
Write-Output 'Restart Word, then Home > Add-ins > Advanced > Shared Folder > Word AI Secure Live > Add.'
