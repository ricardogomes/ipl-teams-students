Import-Module MicrosoftTeams

Write-Host "|   IPL - Teams - Add Students  |" -BackgroundColor DarkGreen -ForegroundColor White
Write-Host "Connecting to Teams - Sign In "

Connect-MicrosoftTeams | Out-Null

Write-Host "Type the Team Name:"
$TeamName = Read-Host

$Team = Get-Team -DisplayName $TeamName
$GroupId = $Team.GroupId

Write-Host "`nAdding Students to $TeamName [$GroupId]"

Import-CSV ".\inputs\students.csv" | ForEach-Object {

    $User = $_.'email'

    Write-Host "`tProcessing $User"

    Add-TeamUser -GroupId $GroupId -User $User

}
