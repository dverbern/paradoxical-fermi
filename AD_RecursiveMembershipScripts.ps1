#==================================================================================================================================================================================
# Purpose:          Intended as something of a one-stop-shop for getting group membership of users, of other groups, using nested (recursive) or direct means.
# Prerequisites:    ADEssentials PowerShell module {Install-Module -Name ADEssentials -Repository PSGallery -Force -Verbose}
# Author:           Daniel Verberne
# Updated:          14/06/2023
#==================================================================================================================================================================================
$ErrorActionPreference = 'Stop'
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Fetch recursive GROUP MEMBERSHIP for a given USER - display HOW each membership is inherited visually
# Utilises function 'Get-ADUserGroupMembershipNestedStructure' (Based on script 'Get-ADUserGroupMembershipNestedStructure.ps1')
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function Get-ADUserGroupMembershipNestedStructure
{
[CmdletBinding()]
    param(
        [Parameter(Mandatory = $TRUE, ValueFromPipeline = $TRUE)][String[]]$UserName,
        [Parameter(Mandatory = $TRUE, ValueFromPipeline = $TRUE)][String[]]$UserFullName
    )
    begin {
        # introduce two lookup hashtables. First will contain cached AD groups,
        # second will contain user groups. We will reuse it for each user.
        # format: Key = group distinguished name, Value = ADGroup object
        $ADGroupCache = @{}
        $UserGroups = @{}
        $OutObject = @()
        # define recursive function to recursively process groups.
        function __findPath ([string]$currentGroup, [string]$comment) 
        {
            Write-Verbose "Processing group: $currentGroup"
            # we must do processing only if the group is not already processed.
            # otherwise we will get an infinity loop
            if (!$UserGroups.ContainsKey($currentGroup)) 
            {
                # retrieve group object, either, from cache (if is already cached)
                # or from Active Directory
                $groupObject = if ($ADGroupCache.ContainsKey($currentGroup)) 
                {
                    Write-Verbose "Found group in cache: $currentGroup"
                    $ADGroupCache[$currentGroup].Psobject.Copy()
                } 
                else 
                {
                    Write-Verbose "Group: $currentGroup is not presented in cache. Retrieve and cache."
                    $g = Get-ADGroup -Identity $currentGroup -Property objectclass,sid,whenchanged,whencreated,samaccountname,displayname,enabled,distinguishedname,memberof,groupscope,groupcategory
                    # immediately add group to local cache:
                    $ADGroupCache.Add($g.DistinguishedName, $g)
                    $g
                }
                
                $c = $comment + "->" + $groupObject.SamAccountName
                
                $UserGroups.Add($c, $groupObject)
                                
                Write-Verbose "Membership Path:  $c"
                foreach ($p in $groupObject.MemberOf) 
                {
                       __findPath $p $c
                }
            } 
            else 
            { 
                Write-Verbose "Closed walk or duplicate on '$currentGroup'. Skipping."
            }
        }
    }
    process 
    {
        $enus = 'en-US' -as [Globalization.CultureInfo]
        foreach ($user in $UserName) {
            Write-Verbose "========== $user =========="
            # clear group membership prior to each user processing
            $UserObject = Get-ADUser -Identity $user -Property objectclass,sid,whenchanged,whencreated,samaccountname,displayname,enabled,distinguishedname,memberof
            $UserObject.MemberOf | ForEach-Object {__findPath $_ $UserObject.SamAccountName}
        }
            foreach($g in $UserGroups.GetEnumerator())
            {
                
                # 20230614 - Merge the GroupScope and GroupCategory to simplify the columns
                $GroupInfo = "$($g.value.ObjectClass)|$($g.value.GroupScope)|$($g.value.GroupCategory)"


                $OutObject += [pscustomobject]@{
                    
                    # 20230614 - don't need the username, it is present in 'InheritancePath'
                    #UserName = $UserObject.SamAccountName;
                    InheritancePath = $g.key;
                    MemberOf = $g.value.SamAccountName;
                    GroupInfo = $GroupInfo;
                    
                    # 20230614 - streamlining the properties returned
                    #ObjectClass = $g.value.ObjectClass;
                    #GroupScope = $g.value.GroupScope;
                    #GroupCategory = $g.value.GroupCategory;
                    #WhenCreated2 = $g.value.WhenCreated.ToString("MM/dd/yyyy hh:mm tt", $enus);
                    #WhenChanged = $g.value.WhenChanged.ToString("MM/dd/yyyy hh:mm tt", $enus);
                }
            }
            
            $returnObject = $OutObject | Sort-Object -Property UserName,InheritancePath,MemberOf
            Return $returnObject
            
            $UserGroups.Clear()
    }
} # End Function
Clear-Host
do
{
    $UserKeyword = Read-Host "Enter some part of user's first or last name for whom you'd like to retrieve recursive group membership"
    $UserObject = Get-ADUser -LDAPFilter "(name=*$UserKeyword*)" | Out-GridView -PassThru -Title 'Select the USER'
    
}
until ($UserObject)
Add-Type -AssemblyName System.DirectoryServices.AccountManagement            
$ct = [System.DirectoryServices.AccountManagement.ContextType]::Domain            
$user = [System.DirectoryServices.AccountManagement.Principal]::FindByIdentity($ct,$($UserObject.SamAccountName))  
$FullName = "$($user.GivenName) $($user.Surname)"
# Echo back what operator chose
Write-Output "Details as follows for user `'$FullName`' ($($UserObject.SamAccountName))"

# Call the function, return the data
$Results = Get-ADUserGroupMembershipNestedStructure -UserName $UserObject.SamAccountName -UserFullName $FullName

# Prompt operator how they want results output
$responseOutputCSV = [System.Windows.Forms.MessageBox]::Show("Would you like the resulting information output to CSV?","Output to CSV?", "YesNo", "Information", "Button1")
if ($responseOutputCSV -eq 'Yes')
{
    # Export results to CSV file
    $Results | Export-Csv -LiteralPath "Get-ADGroupMembershipNestedStructure.$($UserObject.SamAccountName).csv" -NoTypeInformation
}

# Irrespective of operator choice, always output results to screen using the GridView control
$Results | Out-Gridview -Title "Recursive Group Membership for $($user.GivenName) $($user.Surname)"
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Get recursive SECURITY GROUP members for a particular USER (Does not include Distribution Groups)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Clear-Host
do
{
    $UserKeyword = Read-Host "Enter the USER name or part of user's account for whom you'd like to retrieve RECURSIVE group membership"
    $UserObject = Get-ADUser -LDAPFilter "(name=*$UserKeyword*)" | Out-GridView -PassThru -Title 'Select the USER'
    
}
until ($UserObject)
# Echo back what operator chose
Add-Type -AssemblyName System.DirectoryServices.AccountManagement            
$ct = [System.DirectoryServices.AccountManagement.ContextType]::Domain            
$user = [System.DirectoryServices.AccountManagement.Principal]::FindByIdentity($ct,$($UserObject.SamAccountName))  
# Validate the user
Write-Host "Details as follows for user `'$($user.GivenName) $($user.Surname)`' ($($UserObject.SamAccountName))`n" -ForegroundColor Yellow                
$Groups = $User.GetAuthorizationGroups() | Select-Object SamAccountName , Description | Sort-Object -Property SamAccountName -Descending
# Output results both to screen and to operators' clipboard (AND to Out-GridView)
$Groups | Out-GridView
$Groups | Select-Object -ExpandProperty SamAccountName | Set-Clipboard
$Groups
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Get recursive 'MEMBER OF' for a given ROLE GROUP
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# This code requires installation of PowerShell Gallery module 'ADEssentials'
# i.e. 'Tell me what members of role X will have access to'
Clear-Host
$RoleGroup = Read-Host "Enter name of the role or security group whose RECURSIVE group membership you'd like to analyse:"
$Results = Get-WinADGroupMemberOf -Identity $RoleGroup
# From any duplicate groups from results
$Results = $Results | Sort-Object -Property Name -Unique
# Output RECURSIVE membership
Write-Host "`n`n$($RoleGroup) is itself a member (RECURSIVELY) of the following groups:`n" -ForegroundColor Yellow
$Results.Name | Sort-Object | Format-Table -AutoSize
#$Results | Select-Object -ExpandProperty Name | Out-GridHtml -HideFooter -HideButtons -PagingStyle full -Style display -Title $RoleGroup -Simplify
#$Results | Out-GridHtml -Title $RoleGroup -Simplify
$Results | Select-Object -property Name | Out-GridView
# Filter to DIRECT membership
Write-Host "`n`n$($RoleGroup) DIRECT group membership:`n" -ForegroundColor Yellow
$Results | Where-Object -FilterScript {$PSItem.Nesting -eq 0} | Select-Object -ExpandProperty Name | Sort-Object | Format-Table -AutoSize
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Compare RECURSIVE 'member of' (Any Group Type) of TWO ROLE GROUPS
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# This code requires installation of PowerShell Gallery module 'ADEssentials' and 'ImportExcel'
# i.e. 'Tell me what groups that of role X is a member of recursively, compared to Role Y'
$objDifferences = $NULL; $Differences = $NULL
Clear-Host
# Prompt user for a keyword belonging to role / groups to help find the correct one.
do
{
    $Group1Keyword = Read-Host "Enter a keyword to find ROLE GROUP #1"
    # Below:  We are filtering to ONLY retrieve role groups with the underscore '_' character at the start of their name, as per our naming standards.
    $RoleGroup1 = Get-ADGroup -LDAPFilter "(name=*$Group1Keyword*)" | Where-Object -FilterScript {$PSItem.Name -like "_*"} | Select-Object -ExpandProperty Name | Sort-Object | Out-GridView -PassThru -Title 'Choosing ROLE GROUP #1'
    
}
until ($RoleGroup1)

do
{
    $Group2Keyword = Read-Host "Enter a keyword to find ROLE GROUP #2"
    # Below:  We are filtering to ONLY retrieve role groups with the underscore '_' character at the start of their name, as per our naming standards.
    $RoleGroup2 = Get-ADGroup -LDAPFilter "(name=*$Group2Keyword*)" | Where-Object -FilterScript {$PSItem.Name -like "_*"} | Select-Object -ExpandProperty Name | Sort-Object |Out-GridView -PassThru -Title 'Choosing ROLE GROUP #2'
}
until ($RoleGroup2)
Write-Host ""
# Fetch results for each group and filter out duplicate results
Write-Host "$($RoleGroup1) - Fetching RECURSIVE group membership ..."
$RoleGroup1MemberOfRecursive = Get-WinADGroupMemberOf -Identity $RoleGroup1 | Sort-Object -Property Name -Unique -Descending
Write-Host "$($RoleGroup2) - Fetching RECURSIVE group membership ..."
$RoleGroup2MemberOfRecursive = Get-WinADGroupMemberOf -Identity $RoleGroup2 | Sort-Object -Property Name -Unique -Descending
# Compare two sets of objects, specifically by the 'Name' of the groups in each
Write-Host "Comparing set of group membership ..."
$Differences = Compare-Object -ReferenceObject $RoleGroup1MemberOfRecursive -DifferenceObject $RoleGroup2MemberOfRecursive -Property Name
$MemberOfUniqueToRoleGroup1 = $Differences | Where-Object -FilterScript {$PSItem.SideIndicator -eq '<='} | Select-Object -ExpandProperty Name
$MemberOfUniqueToRoleGroup2 = $Differences | Where-Object -FilterScript {$PSItem.SideIndicator -eq '=>'} | Select-Object -ExpandProperty Name
Write-Host "Writing comparison details to Excel spreadsheet ..."
$objDifferences = [pscustomobject] @{
    "Unique to $($RoleGroup1)" = $MemberOfUniqueToRoleGroup1 -join "`r`n" # To prevent multi-valued cells being stored as single line in square braces {}
    "Unique to $($RoleGroup2)" = $MemberOfUniqueToRoleGroup2 -join "`r`n" # To prevent multi-valued cells being stored as single line in square braces {}
}
$ExcelFileName = "CompareMultipleRolesMemberOf.$(Get-Date -f 'yyyyMMdd').xlsx"
# Remove existing file if already exists, we want to create afresh each run of script.
if (Test-Path -Path $ExcelFileName) {Remove-Item -Path $ExcelFileName -Force -ErrorAction SilentlyContinue}
# Write differences between two lists to its own worksheet in same spreadsheet
$objDifferences | Export-Excel -Path $ExcelFileName -WorksheetName 'Differences' -AutoFilter -FreezeTopRow -TableStyle Medium13 -AutoSize
$RoleGroup1MemberOfRecursive | Select-Object -Property Name | Export-Excel -Path $ExcelFileName -WorksheetName $RoleGroup1 -FreezeTopRow -AutoFilter -TableStyle Medium10 -AutoSize -Append
$RoleGroup2MemberOfRecursive | Select-Object -Property Name | Export-Excel -Path $ExcelFileName -WorksheetName $RoleGroup2 -FreezeTopRow -AutoFilter -TableStyle Medium8 -AutoSize -Append
Write-Host `n"Script finished - please see Excel spreadsheet at $($ExcelFileName)" -ForegroundColor Yellow
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Get recursive USER members of a particular role group or division
Clear-Host
$Group1Keyword = $NULL
$RoleGroup1 = $NULL
# Prompt user for a keyword belonging to role / groups to help find the correct one.
do
{
    $Group1Keyword = Read-Host "Enter a keyword to find a particular GROUP"
    $RoleGroup1 = Get-ADGroup -LDAPFilter "(name=*$Group1Keyword*)" | Select-Object -ExpandProperty Name | Sort-Object | Out-GridView -PassThru -Title 'Choosing GROUP'
    
}
until ($RoleGroup1)
Write-Host "You've selected role group - $($RoleGroup1)" -ForegroundColor Green

Get-ADGroupMember -Identity $RoleGroup1 -Recursive | Where-Object -FilterScript {$PSItem.ObjectClass -eq 'user'} | Select-Object -ExpandProperty name
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Get ROLE GROUP membership for given user (Direct Membership Only)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Clear-Host
$userName = Read-Host "Enter user name you wish to check DIRECT group membership for: "
# Validate the user
$user = Get-ADUser -Identity $userName
Write-Host "Details as follows for user `'$($user.GivenName) $($user.Surname)`'`n" -ForegroundColor Yellow
$Groups = Get-ADPrincipalGroupMembership -Identity $userName | Where-Object -FilterScript {$PSItem.Name -like '_*'} | Select-Object -ExpandProperty name | Sort-Object
# Output results both to screen and to operators' clipboard
$Groups | Set-Clipboard
$Groups
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Get DISTRIBUTION LIST membership for a particular USER (Direct Membership Only)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Clear-Host
$username = Read-Host "Enter user name you wish to check DIRECT group membership for: "
# Validate the user
$user = Get-ADUser -Identity $userName
Write-Host "Details as follows for user `'$($user.GivenName) $($user.Surname)`'`n" -ForegroundColor Yellow
$GroupMembershipForUser = Get-ADPrincipalGroupMembership -Identity $userName | Where-Object -FilterScript {$PSItem.GroupCategory -eq 'Distribution'} | Select-Object -ExpandProperty name
$GroupMembershipForUser
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Get DISTRIBUTION LIST membership inherited by members of a particular GROUP (Outputs each DL and level each is nested)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# This code requires installation of PowerShell Gallery module 'ADEssentials'
Clear-Host
do
{
    $GroupKeyword = Read-Host "Enter a keyword to find role / security group"
    $RoleGroup = Get-ADGroup -LDAPFilter "(name=*$GroupKeyword*)" | Select-Object -ExpandProperty Name | Out-GridView -PassThru -Title 'Choosing Security Group/Role'
    
}
until ($RoleGroup)

Write-Host "`nA user who is a member of group, `'$($RoleGroup)`' would inherit membership of these DISTRIBUTION LISTS:`n" -ForegroundColor Yellow
Get-WinADGroupMemberOf -Identity $RoleGroup `
    | Where-Object -FilterScript {($PSItem.Name -like '!*')} `
    | Sort-Object -Property Name -Unique -Descending `
    | Select-Object Name, Nesting
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Determine DIRECT or NESTED group membership for a SPECIFIC GROUP for a SPECIFIC USER
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Envisaged as being useful for apps like Brand Portal (Bynder) that require user to have the security group membership DIRECTLY, nested is insufficient.
# This code requires installation of PowerShell Gallery module 'ADEssentials'
Clear-Host
$Username = Read-Host "Enter the username (username only, no domain) that youy wish to check DIRECT or NESTED group membership for"
$GroupKeyWord = Read-Host "Enter a keyword of group or system name you'd like to filter group membership by"
Get-WinADGroupMemberOf -Identity $Username | Where-Object -FilterScript {$PSItem.Name -like "*$GroupKeyWord*"} | Select-Object Name, Nesting
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Get recursive user members for a particular GROUP (Any type of group)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Clear-Host
do
{
    $GroupKeyword = Read-Host "Enter a keyword to find a specific GROUP"
    $Group = Get-ADGroup -LDAPFilter "(name=*$GroupKeyword*)" | Out-GridView -PassThru -Title 'Choosing a GROUP'
    
}
until ($Group)
# Echo back what operator chose
Write-Host "`nSelected group:  $($Group.Name).  Fetching recursive user membership ..."
Start-Sleep -Seconds 1
# Below, have to pipe users through to Get-ADUser in order to extract properties not returned by Get-ADGroupMember.  Purpose here is to filter out disabled users.
$Users = Get-ADGroupMember -Recursive -Identity $Group.DistinguishedName `
        | ForEach-Object -Process {Get-ADUser -Identity $PSItem -Properties * `
        | Where-Object -FilterScript {$PSItem.Enabled -eq $TRUE} `
        | Select-Object -ExpandProperty Name}
$UsersCount = $Users | Measure-Object | Select-Object -ExpandProperty Count
Write-Host "`nHere are the $($UsersCount) RECURSIVE USER members of SECURITY GROUP, `'$($RoleGroup)`'`n" -ForegroundColor Yellow
# Output results both to screen and to operators' clipboard
$Users | Set-Clipboard
$Users
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Get recursive user members for a particular DISTRIBUTION GROUP
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Clear-Host
$groupname = Read-Host "Enter the display name of the distribution list whose recursive membership you wish to enumerate: "
$groupDN = Get-ADGroup $groupname | Select-Object -ExpandProperty DistinguishedName
$LDAPFilterString = "(memberOf:1.2.840.113556.1.4.1941:=" + $groupDN + ")"
$Users = Get-ADUser -LDAPFilter $LDAPFilterString | Select-Object UserPrincipalName
# Output results both to screen and to operators' clipboard
$Users | Set-Clipboard
$Users
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Get DIRECT (Non-Recursive) GROUP MEMBERSHIP for a given USER
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Clear-Host
do
{
    $UserKeyword = Read-Host "Enter the USER name or part of user's account to find them in Active Directory"
    $User = Get-ADUser -LDAPFilter "(name=*$UserKeyword*)" | Out-GridView -PassThru -Title 'Select the USER'
    
}
until ($User)
# Echo back what operator chose
Write-Host "`nSelected user:  $($User.Name).  Fetching this user's DIRECT GROUP MEMBERSHIP ..."
Start-Sleep -Seconds 1
# Fetch AD User again, this time fetching just the group membership (which is non-recursive using this method)
$DirectMemberOfGroups = Get-ADUser -Identity $User -Properties memberOf | Select-Object -ExpandProperty memberOf
$DirectMemberOfGroupsCount = $DirectMemberOfGroups | Measure-Object | Select-Object -ExpandProperty Count
Write-Host "`n$($User.Name) has DIRECT membership of these $($DirectMemberOfGroupsCount) GROUPS:`n" -ForegroundColor Yellow
# Output results both to screen and to operators' clipboard
$DirectMemberOfGroups | Set-Clipboard
$DirectMemberOfGroups
#==================================================================================================================================================================================