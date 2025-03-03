Import-Module MicrosoftTeams -ErrorAction Stop
Import-Module "$PSScriptRoot\Modules\TeamsManagement\TeamsManagement.psm1" -Force

Write-Host "|   IPL - Teams - Add Students to Channels |" -BackgroundColor DarkGreen -ForegroundColor White
Write-Host "Connecting to Teams - Sign In "

if (-not (Connect-TeamSession)) {
    exit
}

$TeamName = Read-Host "Type the Team Name"
$Team = Get-TeamByName -TeamName $TeamName
if (-not $Team) {
    exit
}

$GroupId = $Team.GroupId
$ChannelsPath = ".\inputs\channels\"

if (-Not (Test-Path -Path $ChannelsPath)) {
    Write-Error "The directory '$ChannelsPath' does not exist. Please ensure the path is correct."
    exit
}

$CSVFiles = Get-ChildItem -Path $ChannelsPath -Filter *.csv
if ($CSVFiles.Count -eq 0) {
    Write-Error "No CSV files found in '$ChannelsPath'. Please ensure the directory contains the necessary CSV files."
    exit
}

$ExistingChannels = Get-TeamChannels -GroupId $GroupId
if (-not $ExistingChannels) {
    exit
}
$ExistingChannelNames = $ExistingChannels.DisplayName

$TotalFiles = $CSVFiles.Count
$ProcessedFiles = 0
$SkippedFiles = 0
$AddedUsers = 0
$FailedAdds = 0

foreach ($File in $CSVFiles) {
    $FileName = $File.Name
    $ChannelName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    
    Write-Host "`nProcessing File: '$FileName' for Channel: '$ChannelName'" -ForegroundColor Cyan

    if ($ExistingChannelNames -notcontains $ChannelName) {
        Write-Warning "Channel '$ChannelName' does not exist in team '$TeamName'. Skipping file '$FileName'."
        $SkippedFiles++
        continue
    }

    $CSVFilePath = Join-Path -Path $ChannelsPath -ChildPath $FileName
    $Users = Import-ChannelUsers -FilePath $CSVFilePath
    
    if (-not $Users) {
        $SkippedFiles++
        continue
    }

    foreach ($User in $Users) {
        $Success = Add-UserToChannel -GroupId $GroupId -ChannelName $ChannelName -UserEmail $User.email
        if ($Success) {
            $AddedUsers++
        } else {
            $FailedAdds++
        }
    }

    $ProcessedFiles++
}

Write-OperationSummary -TotalFiles $TotalFiles -ProcessedFiles $ProcessedFiles -SkippedFiles $SkippedFiles -AddedUsers $AddedUsers -FailedAdds $FailedAdds

Write-Host "DONE" -ForegroundColor Cyan
