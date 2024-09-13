Import-Module MicrosoftTeams -ErrorAction Stop

Write-Host "|   IPL - Teams - Add Students to Channels |" -BackgroundColor DarkGreen -ForegroundColor White
Write-Host "Connecting to Teams - Sign In "

try {
    Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
    Write-Host "Successfully connected to Microsoft Teams." -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to Microsoft Teams. Please check your credentials and try again."
    exit
}

$TeamName = Read-Host "Type the Team Name"

try {
    $Team = Get-Team -DisplayName $TeamName -ErrorAction Stop
    $GroupId = $Team.GroupId
    Write-Host "Team '$TeamName' found with GroupId: $GroupId" -ForegroundColor Green
} catch {
    Write-Error "Team '$TeamName' not found. Please ensure the team exists and you have the necessary permissions."
    exit
}

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

try {
    Write-Host "`nRetrieving existing channels in the team..."
    $ExistingChannels = Get-TeamChannel -GroupId $GroupId -ErrorAction Stop
    $ExistingChannelNames = $ExistingChannels.DisplayName
    Write-Host "Found $($ExistingChannels.Count) channels in team '$TeamName'." -ForegroundColor Green
} catch {
    Write-Error "Failed to retrieve channels for team '$TeamName'. Error: $_"
    exit
}

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

    try {
        $Users = Import-Csv -Path $CSVFilePath -ErrorAction Stop
        if ($Users.Count -eq 0) {
            Write-Warning "CSV file '$FileName' is empty. Skipping."
            $SkippedFiles++
            continue
        }
    } catch {
        Write-Warning "Failed to import CSV file '$FileName'. Error: $_. Skipping."
        $SkippedFiles++
        continue
    }

    foreach ($User in $Users) {
        $UserEmail = $User.email

        if ([string]::IsNullOrWhiteSpace($UserEmail)) {
            Write-Host "`tSkipping empty email entry in file '$FileName'." -ForegroundColor Yellow
            continue
        }

        Write-Host "`tAdding user '$UserEmail' to channel '$ChannelName'..." -NoNewLine

        try {
            Add-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName -User $UserEmail -ErrorAction Stop
            Write-Host " Success" -ForegroundColor Green
            $AddedUsers++
        } catch {
            Write-Host " Failed" -ForegroundColor Red
            Write-Host "`t`tError adding '$UserEmail' to '$ChannelName': $_" -ForegroundColor Red
            $FailedAdds++
        }
    }

    $ProcessedFiles++
}

# Summary of operations
Write-Host "`n=== Operation Summary ===" -ForegroundColor Yellow
Write-Host "Total CSV Files Found: $TotalFiles"
Write-Host "Files Processed: $ProcessedFiles"
Write-Host "Files Skipped: $SkippedFiles"
Write-Host "Total Users Added Successfully: $AddedUsers" -ForegroundColor Green
Write-Host "Total Failed User Additions: $FailedAdds" -ForegroundColor Red
Write-Host "=========================`n" -ForegroundColor Yellow

Write-Host "DONE" -ForegroundColor Cyan
