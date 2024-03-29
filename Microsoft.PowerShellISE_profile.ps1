#=========================================================================================================================
# Script:               Microsoft.PowerShellISE_profile.ps1
# Purpose:              PowerShell profile script
#                       Loads whenever PowerShell ISE is loaded
#                       Opportunity to set variables or check for things or initialise things 
#                       that might be valuable for every session.
#
#                       This module file includes a simple function that allows operators to easily
#                       search Active Directory in a forgiving way.
#
# Prerequisites:        Requires ActiveDirectory module on local machine
#
# Author:               Daniel Verberne
# Updated:              14/06/2023
#=========================================================================================================================
# Variables
#-------------------------------------------------------------------------------------------------------------------------
$ErrorActionPreference = 'Stop'
$MyPublicPowerShellFolder = 'C:\Users\Public\Desktop\PowerShell'
$thisScript = $MyInvocation.MyCommand.Path
$File = Get-Item -Path $thisScript
#-------------------------------------------------------------------------------------------------------------------------
# Main script logic
#-------------------------------------------------------------------------------------------------------------------------
# If the Public Desktop PowerShell folder exists, set that as the current path for the PowerShell session
if (Test-Path -LiteralPath $MyPublicPowerShellFolder) {Set-Location -Path $MyPublicPowerShellFolder} `
    else {Write-Warning "Public `'PowerShell`' folder does not exist on this machine"}


Clear-Host
Write-Output "Loaded Profile Module File, Last Modified $(Get-Date ($File.LastWriteTime) -f g)"

# Setting TLS 1.2 explicitly to ensure not inadvertently using vulnerable TLS mechanisms
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Simple function to find an AD object from just a partial match and output results using the Out-GridView.
Write-Output "Creating function `'DanSearchAD' to facilitate easy Active Directory searching"
Write-Output "Usage:   `'DanSearchAD {keyword}`' [ENTER] | `'searchad {keyword]`' [ENTER]"
# Carriage return
Write-Output "`n"
function DanSearchAD 
{
    [alias("searchad","search")]param([string]$keyword)

    # 20230614 - add in the SamAccountName property, as it's useful to know the username of accounts
    Get-ADObject -LDAPFilter "(name=*$keyword*)" -Properties SamAccountName `
        | Select-Object Name, `
                        SamAccountName, `
                        DistinguishedName, `
                        ObjectClass, `
                        ObjectGUID `
                        | Sort-Object -Property Name `
                        | Out-GridView -PassThru -Title 'Which of these objects?'
}
#=========================================================================================================================
