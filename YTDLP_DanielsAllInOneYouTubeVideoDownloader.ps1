#=========================================================================================================================
# Script:               YTDLP_DanielsAllInOneYouTubeVideoDownloader.ps1
# Purpose:              For a given YouTube Video URL:
#                       1)  Downloads the poster image used by YouTube itself for that video.
#                       2)  Downloads the English subtitle stream (where available) format.
#                       3)  Downloads the YouTube video.
#
# Prerequisites:        Requires utility YT-DLP (https://github.com/yt-dlp/yt-dlp)
#
# Author:               Daniel Verberne
# Updated:              19/06/2023
#=========================================================================================================================
# Variables
#-------------------------------------------------------------------------------------------------------------------------
# Change download folder to suit your preferences!
$OutputPath = "$($env:USERPROFILE)\Downloads"

$YouTubeImageURLFormat = 'https://i.ytimg.com/vi/VIDEO_UNIQUE_IDENTIFIER/maxresdefault.jpg'
$RegExExtractVideoUniqueIdentifier = 'https:\/\/www.youtube.com\/watch\?v=(?<VideoID>[a-zA-Z-_\d]+)'
$RegExExtractSubtitleDefaultFileNameInformation = `
    "^(?<VideoTitle>.*)\s\[(?<VideoID>.*)\].(?<LanguageCode>\w+).(?<Extension>(vtt|srt))$"
#-------------------------------------------------------------------------------------------------------------------------
# Purpose:  Given a YouTube Video URL, fetch it's webpage in order to retrieve video title, channel name.
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
# Purpose:  Download the poster image for a given YouTube video.
# This function has the ability to try more than one variant of a YouTube video image file,
# which are named in reference to their size/dimensions.
#-------------------------------------------------------------------------------------------------------------------------
function Download-YouTubeVideoImage
{

    [CmdletBinding()]Param(
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=0)][string]$video_poster_url,
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=1)][string]$output_poster_full_path)

    try
    { 
        Write-Output "Fetching poster image for this YouTube video ..."

        Add-Type -AssemblyName System.Net.Http
        $http_client = New-Object System.Net.Http.HttpClient
        # Check if image file exists at this URL
        $response = $http_client.GetAsync($video_poster_url)
        $response.Wait()

        if ($response.Result.StatusCode -eq 200)
        {
            Write-Output "Found video image file at $($video_poster_url)"
            Invoke-WebRequest -Uri $video_poster_url -OutFile $OutputImageFullPath -ErrorAction SilentlyContinue
        }
        else
        {
            Write-Output "Could not find video image at $($video_poster_url) ... trying different filename"
            $video_poster_url = $video_poster_url -replace 'maxresdefault','hqdefault'
            $response = $NULL
            $response = $http_client.GetAsync($video_poster_url)
            $response.Wait()

            if ($response.Result.StatusCode -eq 200)
            {
                Write-Output "Found video image file at $($video_poster_url)"
                Invoke-WebRequest -Uri $video_poster_url -OutFile $OutputImageFullPath -ErrorAction SilentlyContinue
            }
            else
            {
                Write-Output "Could not find video image at $($video_poster_url) ... trying different filename"
                # Try a different filename on same host
                $video_poster_url = $video_poster_url -replace 'hqdefault','mqdefault'
                # Reset the $response and try again
                $response = $NULL
                $response = $http_client.GetAsync($video_poster_url)
                $response.Wait()   
        
                # Final attempt!
                if ($response.Result.StatusCode -eq 200)
                {
                    Write-Output "Found video image file at $($video_poster_url)"
                    Invoke-WebRequest -Uri $video_poster_url -OutFile $OutputImageFullPath -ErrorAction SilentlyContinue
                }
                else
                {
                    Write-Warning "Despite multiple attempts, unable to find an image for this YouTube video!"
                }     
            }
        }
    } # End Try
    catch [System.UnauthorizedAccessException]
    {
        Write-Warning "Cannot save file to intended location. Make sure you have permission"
    }
} # end Function
#-------------------------------------------------------------------------------------------------------------------------
# Purpose:  Use the YT-DLP program to download ONLY the subtitle file (english is auto selected)
# for a given YouTube video, where one exists.
#-------------------------------------------------------------------------------------------------------------------------
function Download-YouTubeVideoSubtitles
{
    [CmdletBinding()]Param(
    [Parameter(Mandatory=$TRUE,Position=0)][string]$video_url,
    [Parameter(Mandatory=$TRUE,Position=1)][string]$output_subtitle_full_path)

    # Note:  Must NOT use the --quiet switch, because that suppresses the particular 
    # message I am looking for to determine the availability of a subtitle stream.
    $sub_check = yt-dlp.exe $video_url --skip-download --write-subs -o $output_subtitle_full_path
    Return $sub_check
}
#-------------------------------------------------------------------------------------------------------------------------
# Purpose:  Use the YT-DLP program to download the actual YouTube video file (assumed to have audio and video streams)
# YT-DLP has significant flexibility over which particular streams are downloaded for a given YouTube or other stream.
# This script uses a 'FORMAT SELECTION' syntax whereby we specify the best-quality streams available for a given
# extension, which in turn implies certain codecs used.
# By default, the format selection string will fetch the best quality video stream with an MP4 extension
# and the best quality audio stream with an M4A extension.
# Operator may send their own valid format selection string if they so wish.
#-------------------------------------------------------------------------------------------------------------------------
function Download-YouTubeVideo
{
    [CmdletBinding()]Param(
    [Parameter(Mandatory=$TRUE,Position=0)][string]$vid_url,
    [Parameter(Mandatory=$TRUE,Position=1)][string]$output_filename,
    [Parameter(Mandatory=$FALSE,Position=2)][string]$format_selection_string = "bestvideo[ext=mp4]+bestaudio[ext=m4a]")

    Write-Output "Format selection string in use:  $($format_selection_string)"
    yt-dlp.exe -f $format_selection_string $vid_url -o $output_filename --quiet --progress
}
#-------------------------------------------------------------------------------------------------------------------------
#   SCRIPT STARTS - PROMPT OPERATOR FOR YOUTUBE VIDEO URL - FETCH VIDEO ID FOR POSTER IMAGE URL
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
    # Use regex to extract the unique identifier for the YouTube video URL
    $VideoURL -match $RegExExtractVideoUniqueIdentifier | Out-Null
    $VideoUniqueIdentifier = $Matches['VideoID']
    
    # Use the YouTube video unique identifier to construct a likely URL to that video's 'poster image'
    $VideoImageURL = $YouTubeImageURLFormat -replace 'VIDEO_UNIQUE_IDENTIFIER', $VideoUniqueIdentifier

#-------------------------------------------------------------------------------------------------------------------------
#   FETCH VIDEO ID AND DEFINE OUTPUT FILENAMES
#-------------------------------------------------------------------------------------------------------------------------

    # CALL FUNCTION - Get-YouTubeVideoNameAndChannelName
    $SuggestedFileNamePattern = (Get-YouTubeVideoNameAndChannelName -video_url $VideoURL)[2]
    Write-Output "Video:  $($SuggestedFileNamePattern)"
   
    # Construct output file names for the YouTube video poster image, subtitle and the video itself.
    $OutputImageFullPath = "$($OutputPath)\$($SuggestedFileNamePattern).jpg"
    $OutputSubtitleFullPath = "$($OutputPath)\$($SuggestedFileNamePattern)"
    $OutputVideoFullPath = "$($OutputPath)\$($SuggestedFileNamePattern).mp4"
    
#-------------------------------------------------------------------------------------------------------------------------
#   DOWNLOAD YOUTUBE VIDEO POSTER IMAGE
#-------------------------------------------------------------------------------------------------------------------------

    # CALL FUNCTION TO DOWNLOAD YOUTUBE VIDEO IMAGE
    Write-Output "Fetching IMAGE for this YouTube video..."
    Download-YouTubeVideoImage -video_poster_url $VideoImageURL -output_poster_full_path $OutputImageFullPath
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
    if ($SubtitleCheck[$($SubtitleCheck.Count-1)] -like "*There's no subtitles*")
    {
        Write-Output "Video DOES NOT have subtitles, skipping this part..."
    }
    else # i.e. There ARE subtitles
    {
        Write-Output "Confirmed that video DOES have subtitles..."
        $SubtitleFile = $NULL # Reset this to $NULL each run   

        $SubtitleFile = Get-ChildItem -Path $OutputPath -Filter "$($SuggestedFileNamePattern)*.vtt" -File
        $SubtitleFileCount = $SubtitleFile | Measure-Object | Select-Object -ExpandProperty Count

        if ($SubtitleFile -ne $NULL)
        {
            
            if ($SubtitleFileCount -gt 1)
            {
                Write-Output "We may already have downloaded this subtitle and given it the correct name, skipping any renaming!"
            }
            else
            {
                Write-Output "Located newly-created subtitle VTT file, attempting to rename it ..."
                # Define the new name we want for the subtitle file
                $DesiredSubtitleFileName = "$($SuggestedFileNamePattern)$($SubtitleFile.Extension)"
                # Rename the subtitle file
                Rename-Item $SubtitleFile.FullName -NewName $DesiredSubtitleFileName -Force
            }               
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
        Write-Output "Finished Downloading Video:  $($SuggestedFileNamePattern)"
    }
    else
    {
        Write-Warning "Could not find the expected video file, `'$($OutputVideoFullPath)`'. Download may have failed."
    }


Write-Output "`nScript finished"
#=========================================================================================================================

