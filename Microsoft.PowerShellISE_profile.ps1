# Last Updated 18/04/2023
$MyPublicPowerShellFolder = 'C:\Users\Public\Desktop\PowerShell'
$ErrorActionPreference = 'SilentlyContinue'
# If the Public Desktop PowerShell folder exists, set that as the current path for the PowerShell session
if (Test-Path -LiteralPath $MyPublicPowerShellFolder)
{
    Set-Location -Path $MyPublicPowerShellFolder    
}
else
{
    Write-Warning "The Public Desktop PowerShell folder does not exist on this host machine."
}
# Setting TLS 1.2 explicitly to ensure not inadvertently using vulnerable TLS mechanisms
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12