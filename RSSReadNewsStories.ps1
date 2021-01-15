#=====================================================================================================================================================#
# Purpose: Retrieve current top stories from New York Times via its RSS reader and read out loud using Windows speech synthesizer.
#
# Author:  Daniel Verberne
# Date:    15/01/2021
#=====================================================================================================================================================#
Clear-Host
$URL = 'https://rss.nytimes.com/services/xml/rss/nyt/HomePage.xml'
#-----------------------------------------------------------------------------------------------------------------------------------------------------#
# Main Script Logic Starts Here
#-----------------------------------------------------------------------------------------------------------------------------------------------------#
$Title = "Daniel's New York Times RSS Story reader"
$DodgyCharacters = "â€œ"

# I keep a record of words that the speech synthesiser clearly has issues pronouncing properly, together with the word I think will ensure
# correct pronunciation.
$HashSpeechFixes = @{'Biden' = 'Byden';}

# Load audio/speech synthesizer
Add-Type -AssemblyName System.speech
$Speak = New-Object System.Speech.Synthesis.SpeechSynthesizer

$WebClientObject = New-Object Net.WebClient 
$WebClientObject.UseDefaultCredentials = $True
$RSS = [xml]$WebClientObject.DownloadString($URL)

# Compare RSS last update to the same value this script has 'remembered'
$LastBuildDateFile = 'RSSReadNewsStories.txt'
if (!(Test-Path -Path $LastBuildDateFile))
{
    [datetime]$LastBuildDate = $RSS.rss.channel.lastBuildDate    
    $LastBuildDate | Out-File -FilePath $LastBuildDateFile -Append:$FALSE -Force -NoNewline
}
else
{
    [datetime]$LastBuildDate = Get-Content -Path $LastBuildDateFile
    # Compare to current same information in the RSS feed
    if ([datetime]$RSS.rss.channel.lastBuildDate -gt $LastBuildDate)
    {
        # RSS has been updated since script last checked
        # Fetch news stories that are published since last checked
        $NewStories = $RSS.rss.channel.item | Where-Object -FilterScript {$PSItem.pubDate -gt $LastBuildDate} | Select-Object -First 5 -Property pubDate, description, link

        # update the last update time in our file.
        $LastBuildDate = [datetime]$RSS.rss.channel.lastBuildDate
        $LastBuildDate | Out-File -FilePath $LastBuildDateFile -Append:$FALSE -Force -NoNewline
    }
}

# Below, the introductory line features 'welcome' twice, but that's merely to cater for 
# the fact that on the audio platform I use, wireless headphones, the adapter only turns on 
# the very moment it is asked to play a sound, so I often get the first second or so cutoff, so 
# this is my attempt to work around that.
$Speak.Speak("Welcome.............. Welcome... to $Title")
# Liberal uses of breaks or pauses.
$Speak.Speak("....")

foreach ($NewStory in $NewStories)
{
    # if the story being read was published today...
    if ((Get-Date $NewStory.pubDate -Format d) -eq (Get-Date -Format d))
    {
        $Speak.Speak("Article published today at $(Get-Date $NewStory.pubDate -Format t)..")
    }
    else
    {
        
        # if the story was published a previous day
        $Speak.Speak("From $([datetime]$NewStory.pubDate)..")
    }
    # Remove any dodgy characters from story description first...
    $Description = $NewStory.description -replace $DodgyCharacters,''
    # Replace any words I know that the speech program misproncounces, replacing with a word that I think forces correct pronunciation
    $HashSpeechFixes.Keys | ForEach-Object -Process {$Description = $Description -replace $PSItem,$HashSpeechFixes.Item($PSItem)}
    $Speak.Speak($Description)
    $Speak.Speak("....")
}

$Speak.Speak("Thank you for listening to $Title")
#=====================================================================================================================================================#