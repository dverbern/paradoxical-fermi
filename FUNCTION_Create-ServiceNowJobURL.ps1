#=========================================================================================================================
# Script:               FUNCTION_Create-ServiceNowJobURL.ps1
# Purpose:              Generate a hyperlink to navigate direct to a given ServiceNow job, such as an Incident (INCxxxx),
#                       Change task (CTASKxxxx) and regular task.
# Author:               Daniel Verberne
# Updated:              20/06/2023
#=========================================================================================================================
# Global variables
#-------------------------------------------------------------------------------------------------------------------------
$SNOWURL = "https://yarravalley.service-now.com/nav_to.do?uri=TICKETTYPE.do?sys_id=JOBNUMBER"
$RegexValidateSNOWJobNumber = '^(INC|RITM|S?CTASK)\d{7}$'

#-------------------------------------------------------------------------------------------------------------------------
# Create hyperlinks to ServiceNow tickets
function Create-ServiceNowJobURL
{
    Param(
        [Parameter(Mandatory=$TRUE,Position=0)][ValidatePattern("^(INC|RITM|S?CTASK)\d{7}$")][string]$job_number
    )
    
    # Constructing URL to this ticket, part 1
    switch -Wildcard ($job_number)
    {            
        'RITM*'  {$job_URL = $SNOWURL -replace 'TICKETTYPE','sc_req_item'}
        'INC*'   {$job_URL = $SNOWURL -replace 'TICKETTYPE','incident'}
        'SCTASK*'{$job_URL = $SNOWURL -replace 'TICKETTYPE','sc_task'}
        'CTASK*' {$job_URL = $SNOWURL -replace 'TICKETTYPE','change_task'}
    }

    # Constructing URL to this ticket, part 2
    $url += $job_URL -replace 'JOBNUMBER', $job_number

    Return $url

} # End function Create-ServiceNowJobURLS
#-------------------------------------------------------------------------------------------------------------------------
Clear-Host
$TicketNumber = (Get-Clipboard).Trim()

# Verify we've got something valid before proceeding
if ($TicketNumber -match $RegexValidateSNOWJobNumber)
{
    Write-Output "Ticket Number:  $($TicketNumber)"
    Write-Output "Generating URL to ticket in ServiceNow ..."
    $URL = Create-ServiceNowJobURL -job_number $TicketNumber

    # Sanity-check, make sure count of URLs same as count of ticket numbers...
    if ($TicketNumber.Count -ne $URL.Count)
    {

        Write-Warning "Mismatch of results..."
    }
    else
    {
        $URL | Set-Clipboard
        Write-Output "Ticket URL:  $($URL) (Sent to your clipboard)" 

    }

}
else
{
    Write-Output "No valid SNOW ticket number found in clipboard"
}
Write-Output "Script completed"
#=========================================================================================================================