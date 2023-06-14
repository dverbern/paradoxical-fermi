#=========================================================================================================================
# Script:               YTDLP_AllInOneYouTubeVideoDownloader.ps1
# Purpose:              For a given YouTube Video URL:
#                       1)  Downloads the poster image used by YouTube itself for that video.
#                       2)  Downloads the English subtitle stream (where available) in *.vtt format.
#                       3)  Downloads the YouTube video itself in MP4 using YT-DLP utility.
#
# Prerequisites:        Requires presence of utility YT-DLP installed on system.
#
# Author:               Daniel Verberne
# Updated:              14/06/2023
#=========================================================================================================================
# Variables
#-------------------------------------------------------------------------------------------------------------------------
# Update the output path to suit your needs
$OutputPath = "$($env:USERPROFILE)\Downloads"
$YouTubeImageURLFormat = 'https://i.ytimg.com/vi/VIDEO_UNIQUE_IDENTIFIER/maxresdefault.jpg'
$RegExExtractVideoUniqueIdentifier = 'https:\/\/www.youtube.com\/watch\?v=(?<VideoID>[a-zA-Z-_\d]+)'
$RegExExtractSubtitleDefaultFileNameInformation = `
    "^(?<VideoTitle>.*)\s\[(?<VideoID>.*)\].(?<LanguageCode>\w+).(?<Extension>(vtt|srt))$"
#-------------------------------------------------------------------------------------------------------------------------
# Purpose:  Utilise .NET to determine the dimensions of an image file.
# Intended to be used in the context of downloading the YouTube 'poster image' for a given video file.
# Source: https://powershelladministrator.com/2021/09/20/move-images-based-on-dimension/
#-------------------------------------------------------------------------------------------------------------------------
function Get-ImageDimensions
{
    [CmdletBinding()]Param(
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=0)][string]$image_file_path)

    Add-Type -AssemblyName System.Drawing
    
    # Get the image information
    $image_bitmap_Object = New-Object System.Drawing.Bitmap $image_file_path
    # Get the image dimensions and piece together as a single string
    $image_dimensions = "$($image_bitmap_Object.Width) x $($image_bitmap_Object.Height)"

    # Close the image
    $image_bitmap_Object.Dispose()

    Return $image_dimensions
}
#-------------------------------------------------------------------------------------------------------------------------
# Purpose:  Given a YouTube Video URL, fetch it's webpage and process the content using 
# regular expressions to fetch video name and channel name.
# End result is providing a filename to use for the downloaded video file and it's subtitles and image.
#-------------------------------------------------------------------------------------------------------------------------
function Get-YouTubeVideoNameAndChannelName
{
    [CmdletBinding()]Param(
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=0)][string]$video_url)
    
    $get_video_title_regex = "\<title\>(?<Title>.*)\<\/title\>"
    $get_video_channel_name_regex = "`"author`":`"(?<ChannelName>.*?)`""
    $string_removals = '.\-.YouTube','\&quot\;','\&\#39\;','\?','\:','\|','\[','\]'    
    
    $web_request = Invoke-WebRequest -UseBasicParsing -Uri $video_url
    $web_request.Content -match $get_video_title_regex | Out-Null
    $video_title = $Matches['Title']
    # Carry out string replacement (removal) as we can only deal with characters allowed in file names.
    foreach ($string_removal in $string_removals)
    {
        $video_title = $video_title -replace $string_removal,''
    }

    $web_request.Content -match $get_video_channel_name_regex | Out-Null
    $channel_name = $Matches['ChannelName']
    Write-Output "Channelname:  $($channel_name)"
    Write-Output "Videotitle:  $($video_title)"
    
    $suggested_video_filename = $video_title
    Return $suggested_video_filename

}
#-------------------------------------------------------------------------------------------------------------------------
# Purpose:  Use the YT-DLP external utility to download ONLY the subtitle file (english is auto selected)
# for a given YouTube video, where one exists.
#-------------------------------------------------------------------------------------------------------------------------
function Download-YouTubeVideoSubtitles
{
    [CmdletBinding()]Param(
    [Parameter(Mandatory=$TRUE,Position=0)][string]$video_url,
    [Parameter(Mandatory=$TRUE,Position=1)][string]$output_subtitle_full_path)

    # GOTCHYA - In order to properly evaluated whether a video has subtitles or not,
    # we must NOT use --quiet mode, because that suppresses the messages we're looking for!
    $sub_check = yt-dlp.exe $video_url --skip-download --write-subs -o $output_subtitle_full_path
    Return $sub_check
}
#-------------------------------------------------------------------------------------------------------------------------
function Download-YouTubeVideo
{
    [CmdletBinding()]Param(
    [Parameter(Mandatory=$TRUE,Position=0)][string]$vid_url,
    [Parameter(Mandatory=$TRUE,Position=1)][string]$output_filename)
    
	# Note:  Syntax below instructs YT-DLP to fetch a stream that has MP4 extension and contains M4A format 
	# audio.  This represents what I consider a well-supported combination for most media playing purposes.
    yt-dlp.exe -f "bestvideo[ext=mp4]+bestaudio[ext=m4a]" $vid_url -o $output_filename --quiet --progress
}
#-------------------------------------------------------------------------------------------------------------------------
# Script starts
#-------------------------------------------------------------------------------------------------------------------------
Clear-Host
Write-Output "Starting Dan's YouTube video / poster / subtitle fetcher...`n"
do
{
    $VideoURL = Read-Host "Please enter YouTube video URL"
}
while ($VideoURL -notmatch $RegExExtractVideoUniqueIdentifier)

if ($VideoURL -match $RegExExtractVideoUniqueIdentifier)
{
    
#-------------------------------------------------------------------------------------------------------------------------
#   FETCH VIDEO ID AND DEFINE OUTPUT FILENAMES
#-------------------------------------------------------------------------------------------------------------------------

    # CALL FUNCTION - Get-YouTubeVideoNameAndChannelName
    $SuggestedFileNamePattern = (Get-YouTubeVideoNameAndChannelName -video_url $VideoURL)[2]

    Write-Host "Video:  $($SuggestedFileNamePattern)" -ForegroundColor Green

    # Use regular expressions to attempt to extract just the YouTube video identifier and therefore likely image
    # hyperlink.
    $VideoURL -match $RegExExtractVideoUniqueIdentifier | Out-Null
    $VideoUniqueIdentifier = $Matches['VideoID']
    $VideoImageURL = $YouTubeImageURLFormat -replace 'VIDEO_UNIQUE_IDENTIFIER', $VideoUniqueIdentifier
    
    # Define the output JPG file name and path
    $OutputImageFullPath = "$($OutputPath)\$($SuggestedFileNamePattern).jpg"
    
    # Define the output VTT file name and path (Note that I don't need to supply the file extension, applies automatically)
    $OutputSubtitleFullPath = "$($OutputPath)\$($SuggestedFileNamePattern)"

    # Define the output MP4 file name and path
    $OutputVideoFullPath = "$($OutputPath)\$($SuggestedFileNamePattern).mp4"
    
#-------------------------------------------------------------------------------------------------------------------------
#   DOWNLOAD VIDEO'S JPEG TITLE IMAGE AND CALCULATE IMAGE DIMENSIONS
#-------------------------------------------------------------------------------------------------------------------------

    try
    {
        
        Write-Output "Fetching POSTER IMAGE for this YouTube video..."
        # Download the YouTube poster JPG image by its direct URL
        Invoke-WebRequest -Uri $VideoImageURL -OutFile $OutputImageFullPath -ErrorAction SilentlyContinue
        if (Test-Path -Path $OutputImageFullPath)
        {
            
            $ImageFileSize = Get-Item -Path $OutputImageFullPath | Select-Object -ExpandProperty Length
            # CALL FUNCTION - Get-ImageDimensions
            $ImageFileDimensions = Get-ImageDimensions -image_file_path $OutputImageFullPath
            # $ImageDimensions typically has values like '1280 x 720'

            # NOTE:  Not currently utilising data returned from dimensions function, but 
            # could use it if needed.
        }
    }
    catch [System.UnauthorizedAccessException]
    {
        Write-Warning "Cannot save file to intended location. Make sure you have permission"
    }
}
else
{
    Write-Warning "Could not perform a successful regex match against this URL"
}
#-------------------------------------------------------------------------------------------------------------------------
#   DOWNLOAD VIDEO SUBTITLES USING YT-DLP AND RENAME OUTPUT FILE
#-------------------------------------------------------------------------------------------------------------------------

    Write-Output "Fetching SUBTITLES for this YouTube video..."
    # CALL FUNCTION - Download-YouTubeVideoSubtitles
    $SubtitleCheck = Download-YouTubeVideoSubtitles -video_url $VideoURL -output_subtitle_full_path $OutputSubtitleFullPath
    # Check the resulting string array from trying to download video subtitles
    # If the video doesn't actually HAVE any subtitles, we'll expect to see something to the effect
    # of 'There's no subtitle' in the last line of that returned string array.
    if ($SubtitleCheck[$($SubtitleCheck.Count-1)] -like "*There's no subtitles*")
    {
        Write-Host "Video DOES NOT have subtitles, skipping this part..." -ForegroundColor Yellow
    }
    else # i.e. There ARE subtitles
    {
        Write-Output "Confirmed that video DOES have subtitles..."
        $SubtitleFile = $NULL # Reset this to $NULL each run   
        $SubtitleFile = Get-ChildItem -Path $OutputPath -Filter "$($SuggestedFileNamePattern)*.vtt" -File

        if ($SubtitleFile -ne $NULL)
        {
            Write-Output "Located newly-created subtitle VTT file, attempting to rename it ..."
            # Define the new name we want for the subtitle file
            $DesiredSubtitleFileName = "$($SuggestedFileNamePattern)$($SubtitleFile.Extension)"
            # Rename the subtitle file
            Rename-Item $SubtitleFile.FullName -NewName $DesiredSubtitleFileName -Force
        }
        else
        {
            Write-Warning "Could not find the newly-downloaded subtitle VTT file in order to rename it!"
        }

    }

#-------------------------------------------------------------------------------------------------------------------------
#   DOWNLOAD YOUTUBE VIDEO USING YT-DLP
#-------------------------------------------------------------------------------------------------------------------------

    Write-Output "Downloading actual AUDIO/VIDEO file for this YouTube video..."
    # CALL FUNCTION - Download-YouTubeVideo
    Download-YouTubeVideo -vid_url $VideoURL -output_filename $OutputVideoFullPath

    # Verify if MP4 file exists
    Write-Output "Checking if video exists..."
    if (Test-Path -Path $OutputVideoFullPath -ErrorAction SilentlyContinue)
    {
        Write-Host "Finished Downloading Video:  $($SuggestedFileNamePattern)" -ForegroundColor Green
    }
    else
    {
        Write-Warning "Could not find the expected video file, `'$($OutputVideoFullPath)`'. Download may have failed."
    }


Write-Output "`nScript finished"
#=========================================================================================================================

