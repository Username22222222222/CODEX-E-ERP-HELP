param(
    [Parameter(Mandatory = $true)]
    [string]$PlainText
)

$ErrorActionPreference = 'Stop'
Add-Type -Path 'C:\X-ERP-Encryption\AesEncryption.cs'
$aes = [AesEncryption]::new()
$aes.EncryptString($PlainText)
