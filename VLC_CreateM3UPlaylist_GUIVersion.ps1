#=========================================================================================================================

# Create an *.m3u playlist of a specific folder, assuming it contains *.flac or *.mp3 files.
# Playlist uses relative paths to the audio files so users can move the folders where they wish and 
# shouldn't break the playlist.
#
# Using tentative GUI skills I've been developing!
#               
# Author:       Daniel Verberne
# Last Updated: 15/06/2023
#=========================================================================================================================

Function Generate-Form
{
    # Init PowerShell Gui
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $formDefault                     = New-Object System.Windows.Forms.Form
    $formDefault.ClientSize          = '724,234'
    $formDefault.text                = "Dan`'s M3U Playlist Creator"
    $formDefault.BackColor           = "#4a4a4a"
    $formDefault.TopMost             = $false

    $buttonSpecificFolder            = New-Object system.Windows.Forms.Button
    $buttonSpecificFolder.text       = "Create Playlist!"
    $buttonSpecificFolder.width      = 161
    $buttonSpecificFolder.height     = 30
    $buttonSpecificFolder.location   = New-Object System.Drawing.Point(32,70)
    $buttonSpecificFolder.BackColor  = "#ffffff"
    $buttonSpecificFolder.Font       = 'Calibri,13'

    $txtStatus                       = New-Object System.Windows.Forms.TextBox
    $txtStatus.multiline             = $false
    $txtStatus.width                 = 377
    $txtStatus.height                = 20
    $txtStatus.visible               = $false
    $txtStatus.location              = New-Object System.Drawing.Point(19,180)
    $txtStatus.BackColor             = "#ffffff"
    $txtStatus.Font                  = 'Calibri,13'

    $formDefault.Controls.AddRange(@($buttonSpecificFolder,$txtStatus))

    # Here is where I find syntax and logic a little confusing, but basically below we are 
    # telling PowerShell that 'here is a new event' for these buttons, on user clicking.  
    # In the brackets and braces, we're giving the name of the Function that PowerShell is to run
    # upon a click event.
    $buttonSpecificFolder.Add_Click({SpecificFolderButtonClick})
    
    # We're about to now request that the form be shown, using 'ShowDialog'.
    # Instead of suppressing this command's output by declaring the line below [void] or 
    # outputting to Out-Null, we instead want to put the output of this command in a variable,
    # because it tells us whether the user hit Cancel or closed the program, etc, which
    # gives us more control over the rest of the program.

    $UserInteractionToForm = $formDefault.ShowDialog()
    if ($UserInteractionToForm –eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        Write-Output "User has cancelled out of the program"
    }

} # End Function Generate-Form

Function GetAudioFileLengthInSeconds
{
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
    [string]$Path)

    $Shell = New-Object -COMObject Shell.Application
    $Folder = Split-Path $Path
    $File = Split-Path $Path -Leaf
    $Shellfolder = $Shell.Namespace($Folder)
    $Shellfile = $Shellfolder.ParseName($File)
    ($Shellfolder.GetDetailsOf($Shellfile, 27)) -match $TrackLengthExtractionRegEx
    [int]$Hours = $Matches['Hours']
    [int]$Minutes = $Matches['Minutes']
    [int]$Seconds = $Matches['Seconds']
    [int]$TotalTrackTimeInSeconds = (($Hours*60)*60)+($Minutes*60)+$Seconds
        
    $Band = $Shellfolder.GetDetailsOf($Shellfile, 20)
    $TrackOnly = $Shellfolder.GetDetailsOf($Shellfile, 21)
    $TidyTrackName = "$Band - $TrackOnly"
    
    Return $TotalTrackTimeInSeconds, $TidyTrackName
} # End Function GetAudioFileLengthInSeconds


Function SpecificFolderButtonClick()
{
    $VLCExecutable = "C:\Program Files\VideoLAN\VLC\vlc.exe"
    
    # Change status label to visible and populate so we know program is doing something.
    $txtStatus.Visible = $TRUE
    $txtStatus.Text = "Starting procedure"
    $M3UHeaderString = '#EXTM3U'
    $M3UAdditionalHeaderString = "#Created via PowerShell script by Daniel Verberne"
    $TrackLengthExtractionRegEx = '^(?<Hours>\d+):(?<Minutes>\d+):(?<Seconds>\d+)$'
   
    $Foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $Foldername.Description = "Select a folder you'd like to create a playlist for"
    $Foldername.rootfolder = "MyComputer"
    $Foldername.SelectedPath = $InitialDirectory
    if($foldername.ShowDialog() -eq "OK")
    {
        $FolderToScan += $Foldername.SelectedPath
    }
    $txtStatus.Text = $FolderToScan
  
    # Script can and does bomb at line below if the process doesn't result
    # in a valid playlist being created
    If (!(Test-Path -Path $FolderToScan))
    {
        [System.Windows.Forms.MessageBox]::Show("No valid folder chosen","Abort", "OK" , "Information" , "Button1")
        Exit;    
    }
    else
    {
        
        $M3UFileName = "$(Split-Path -Path $FolderToScan -Leaf).m3u"   
        $M3UFileName | Out-Host

        $M3UFileNameFullPath = "$FolderToScan\$M3UFileName"

        # Note about Get-ChildItem and -Include:
        # When the Include parameter is used, the Path parameter needs a trailing asterisk (*) wildcard to specify the directory's contents. For example, -Path C:\Test\*.
        # Source: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/get-childitem?view=powershell-7

        # UPDATE 30/11/2020 - My script fails# to execute properly if the audio file names contain characters like square braces '[ ]'.  
        # Really need to remove those if they exist before proceeding further.
        
        Write-Host "Checking if files need renaming to remove any unacceptable characters..."
        $Files = Get-ChildItem -Path "$FolderToScan\*" -Filter *.* -Include *.flac, *.mp3 | Select-Object -Property *
        foreach ($File in $Files)
        {
            $CurrentFileName = $File.Name
            $CurrentFilePath = "$FolderToScan\$($CurrentFileName)"

            # Update 11/01/2021 - added an extra iteration of replace, to remove any hash '#' symbols.
            $NewFileName = (($CurrentFileName -replace '\[','') -replace '\]','') -replace '#',''
            $NewFilePath = "$FolderToScan\$($NewFileName)"

            Rename-Item -LiteralPath $CurrentFilePath -NewName $NewFileName
        }

        # Now below we're doing this 'Get-ChildItem' thing all over again, but it's pretty quick so shouldn't matter.
        $AudioFiles = Get-ChildItem -Path "$FolderToScan\*" -Filter *.* -Include *.flac, *.mp3 | Select-Object -Property *
        $AudioFiles | Out-Host

        $M3UHeaderString | Out-File -FilePath $M3UFileNameFullPath -Encoding utf8 -Append:$FALSE -Force:$TRUE
        $M3UAdditionalHeaderString | Out-File -FilePath $M3UFileNameFullPath -Encoding utf8 -Append:$TRUE

        foreach ($AudioFile in $AudioFiles)
        {
            Write-Host "Processing track `'$AudioFile`'"
            $ReturnTrackDetails = GetAudioFileLengthInSeconds($AudioFile.FullName)
            $TrackLength = $ReturnTrackDetails[1]
            $TrackName = $ReturnTrackDetails[2]
            "#EXTINF:$TrackLength,$TrackName" | Out-File -FilePath $M3UFileNameFullPath -Encoding utf8 -Append:$TRUE
            
            #
            # M3U format string replacement rules
            #
            #
            # 1. "%20" replaces any whitespace
            # 2. "%23" replaces any hash # symbol
            # 3. "%27" replaces any apostrophe '
            
            If ($AudioFile.Name -match '(#)')
            {
                $AudioFile.Name -replace '#','%23' | Out-File -FilePath $M3UFileNameFullPath -Encoding utf8 -Append:$TRUE
            }
            If ($AudioFile.Name -match '(\s)')
            {
                $AudioFile.Name -replace '\s','%20' | Out-File -FilePath $M3UFileNameFullPath -Encoding utf8 -Append:$TRUE
            }
            If ($AudioFile.Name -match "(')")
            {
                $AudioFile.Name -replace "'","%27" | Out-File -FilePath $M3UFileNameFullPath -Encoding utf8 -Append:$TRUE
            }

        }
        $txtStatus.Font = 'Arial,10,style=Bold'
        $txtStatus.Text = $M3UFileNameFullPath
		$PlaylistFullPath = $txtStatus.Text
	
        # Offer to load the newly-created playlist in VLC, if present on this machine.
        If (Test-Path -Path $VLCExecutable)
        {
            # Ask user if they want to launch the newy-created playlist in VLC
            $Response = [System.Windows.Forms.MessageBox]::Show("Would you like to launch this playlist with VLC?","Playlist Created", "YesNo" , "Information" , "Button1")
            If ($Response -eq 'Yes')
            {
                # call the VLC executable with the playlist as the sole parameter
                &$VLCExecutable $PlaylistFullPath
            }
        }
    }

} # End Function SpecificFolderButtonClick


#---------------------------------------------------------[Script]--------------------------------------------------------
# Script logic proper begins here
#-------------------------------------------------------------------------------------------------------------------------
# Guts of program start below by calling the first function.
Generate-Form
Set-Location -Path "C:\Users\Public\Desktop\PowerShell"
#=========================================================================================================================
