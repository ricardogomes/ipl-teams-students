Import-Module MicrosoftTeams

Write-Host "|   IPL - Teams - Set Up Team  |" -BackgroundColor DarkGreen -ForegroundColor White
Write-Host "Connecting to Teams - Sign In "

try {
    Connect-MicrosoftTeams -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Teams." -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to Microsoft Teams. Please check your credentials and try again."
    exit
}

$coTeacherID = "<CHANGE TO ADD ANOTHER TEACH AS OWNER OF THE TEAM - NEEDS TO BE THE USER ID - Get-TeamUser -GroupId X -Role Owner>"

$TeamName = Read-Host "Type the Team Name"

do {
    $dInput = Read-Host "How many day practical shifts?"
    if ([int]::TryParse($dInput, [ref]$null)) {
        $d = [int]$dInput
        if ($d -lt 0) {
            Write-Host "Please enter a non-negative integer." -ForegroundColor Yellow
            $d = $null
        }
    } else {
        Write-Host "Invalid input. Please enter a valid integer." -ForegroundColor Yellow
    }
} until ($d -ne $null)

do {
    $nInput = Read-Host "How many night practical shifts?"
    if ([int]::TryParse($nInput, [ref]$null)) {
        $n = [int]$nInput
        if ($n -lt 0) {
            Write-Host "Please enter a non-negative integer." -ForegroundColor Yellow
            $n = $null
        }
    } else {
        Write-Host "Invalid input. Please enter a valid integer." -ForegroundColor Yellow
    }
} until ($n -ne $null)


try {
    $team = Get-Team -DisplayName $TeamName -ErrorAction Stop
    Write-Host "Team '$TeamName' found with GroupId: $($team.GroupId)" -ForegroundColor Green
} catch {
    Write-Error "Team '$TeamName' not found. Please ensure the team exists and you have the necessary permissions."
    exit
}

$groupId = $team.GroupId

# Function to create a private channel
function New-PrivateChannel {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ChannelName
    )

    Write-Host "Creating private channel: $ChannelName"

    try {
        New-TeamChannel `
            -GroupId $groupId `
            -DisplayName $ChannelName `
            -Description "Private channel for $ChannelName" `
            -MembershipType Private `
            -Owner $coTeacherID `
            -ErrorAction Stop

        Write-Host "Channel '$ChannelName' created successfully." -ForegroundColor Green
    } catch {
        Write-Error "Failed to create channel '$ChannelName'. Error: $_"
    }
}

# Create Theoretical Channels
$theoreticalDayChannel = "D - T"
$theoreticalNightChannel = "PL - T"

New-PrivateChannel -ChannelName $theoreticalDayChannel
New-PrivateChannel -ChannelName $theoreticalNightChannel

# Create Practical Day Shift Channels
for ($i = 1; $i -le $d; $i++) {
    $channelName = "D - PL$i"
    New-PrivateChannel -ChannelName $channelName
}

# Create Practical Night Shift Channels
for ($i = 1; $i -le $n; $i++) {
    $channelName = "PL - PL$i"
    New-PrivateChannel -ChannelName $channelName
}

Write-Host "All specified channels have been created successfully." -ForegroundColor Cyan