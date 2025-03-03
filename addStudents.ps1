Import-Module MicrosoftTeams -ErrorAction Stop
Import-Module "$PSScriptRoot\Modules\TeamsManagement\TeamsManagement.psm1" -Force

Write-Host "|   IPL - Teams - Add Students  |" -BackgroundColor DarkGreen -ForegroundColor White
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
$StudentsFile = ".\inputs\students.csv"

if (-not (Test-Path -Path $StudentsFile)) {
    Write-Error "Students file not found at '$StudentsFile'. Please ensure the file exists."
    exit
}

Write-Host "`nAdding Students to $TeamName [$GroupId]"

$Students = Import-Csv -Path $StudentsFile
$TotalStudents = $Students.Count
$SuccessCount = 0
$FailedCount = 0

foreach ($Student in $Students) {
    $UserEmail = $Student.email
    
    if ([string]::IsNullOrWhiteSpace($UserEmail)) {
        Write-Warning "Skipping empty email entry"
        $FailedCount++
        continue
    }
    
    if (Add-TeamMember -GroupId $GroupId -UserEmail $UserEmail) {
        $SuccessCount++
    } else {
        $FailedCount++
    }
}

Write-Host "`n=== Student Addition Summary ===" -ForegroundColor Yellow
Write-Host "Total students processed: $TotalStudents"
Write-Host "Successfully added: $SuccessCount" -ForegroundColor Green
Write-Host "Failed to add: $FailedCount" -ForegroundColor Red
Write-Host "=========================`n" -ForegroundColor Yellow

Write-Host "DONE" -ForegroundColor Cyan
