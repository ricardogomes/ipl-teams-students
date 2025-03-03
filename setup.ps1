Import-Module MicrosoftTeams -ErrorAction Stop
Import-Module "$PSScriptRoot\Modules\TeamsManagement\TeamsManagement.psm1" -Force

Write-Host "|   IPL - Teams - Set Up Team  |" -BackgroundColor DarkGreen -ForegroundColor White
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

# Get team owners
$TeamOwners = Get-TeamOwners -GroupId $GroupId
if (-not $TeamOwners) {
    Write-Warning "No team owners found or failed to retrieve owners. Channels will be created without owners."
}

# Get number of shifts
$DayShifts = Read-ValidInteger -Prompt "How many day practical shifts?"
$NightShifts = Read-ValidInteger -Prompt "How many night practical shifts?"

# Create Theoretical Channels
$TheoreticalChannels = @(
    @{ Name = "D - T"; Description = "Theoretical classes - Day shift" },
    @{ Name = "PL - T"; Description = "Theoretical classes - Night shift" }
)

$SuccessCount = 0
$FailedCount = 0
$CreatedChannels = @()

# Create theoretical channels
foreach ($Channel in $TheoreticalChannels) {
    if (New-PrivateChannel -GroupId $GroupId -ChannelName $Channel.Name -Description $Channel.Description) {
        $SuccessCount++
        $CreatedChannels += $Channel.Name
    } else {
        $FailedCount++
    }
}

# Create Practical Day Shift Channels
for ($i = 1; $i -le $DayShifts; $i++) {
    $ChannelName = "D - PL$i"
    $Description = "Practical classes - Day shift group $i"
    if (New-PrivateChannel -GroupId $GroupId -ChannelName $ChannelName -Description $Description) {
        $SuccessCount++
        $CreatedChannels += $ChannelName
    } else {
        $FailedCount++
    }
}

# Create Practical Night Shift Channels
for ($i = 1; $i -le $NightShifts; $i++) {
    $ChannelName = "PL - PL$i"
    $Description = "Practical classes - Night shift group $i"
    if (New-PrivateChannel -GroupId $GroupId -ChannelName $ChannelName -Description $Description) {
        $SuccessCount++
        $CreatedChannels += $ChannelName
    } else {
        $FailedCount++
    }
}

# Add owners to all created channels
if ($TeamOwners) {
    Write-Host "`nAdding team owners to channels..." -ForegroundColor Cyan
    $OwnerSuccessCount = 0
    $OwnerFailedCount = 0
    
    foreach ($Channel in $CreatedChannels) {
        foreach ($Owner in $TeamOwners) {
            if (Add-ChannelOwner -GroupId $GroupId -ChannelName $Channel -UserEmail $Owner.User) {
                $OwnerSuccessCount++
            } else {
                $OwnerFailedCount++
            }
        }
    }
}

Write-Host "`n=== Channel Creation Summary ===" -ForegroundColor Yellow
Write-Host "Successfully created channels: $SuccessCount" -ForegroundColor Green
Write-Host "Failed to create channels: $FailedCount" -ForegroundColor Red
if ($TeamOwners) {
    Write-Host "`nOwner Assignment Summary:" -ForegroundColor Yellow
    Write-Host "Successfully added owners: $OwnerSuccessCount" -ForegroundColor Green
    Write-Host "Failed owner assignments: $OwnerFailedCount" -ForegroundColor Red
}
Write-Host "=========================`n" -ForegroundColor Yellow

Write-Host "DONE" -ForegroundColor Cyan