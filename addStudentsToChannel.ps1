Import-Module MicrosoftTeams

Write-Host "|   IPL - Teams - Add Students to Channel |" -BackgroundColor DarkGreen -ForegroundColor White
Write-Host "Connecting to Teams - Sign In "
Connect-MicrosoftTeams | Out-Null

Write-Host "Type the Team Name:"
$TeamName = Read-Host

$Team = Get-Team -DisplayName $TeamName
$GroupId = $Team.GroupId

Write-Host "`nGetting Channels"
$Channels = Get-TeamChannel -GroupId $GroupId
for($index=0; $index -lt $Channels.Length; $index++)
{
  $ChannelName = $Channels[$index].DisplayName
  Write-Host "[$index] $ChannelName"
}
Write-Host "`nChoose one Channel [#]: " -NoNewLine
$Option = Read-Host
$Channel = $Channels[$Option]
$ChannelName = $Channel.DisplayName
Write-Host "`nWorking on Channel - $ChannelName"

Write-Host "`nGetting Files"
$ChannelsPath = ".\inputs\channels\"
$Files = Get-ChildItem $ChannelsPath
for($index=0; $index -lt $Files.Length; $index++)
{
  $FileName = $Files[$index].Name
  Write-Host "[$index] $FileName"
}
Write-Host "`nChoose one CSV File [#]: " -NoNewLine
$Option = Read-Host
$File = $Files[$Option]
$FileName = $File.Name
Write-Host "`nWorking on File - $FileName"

Import-CSV "$ChannelsPath\$FileName" | ForEach-Object {

  $User = $_.'email'

  Write-Host "`tProcessing $User for $ChannelName"

  Add-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName -User $User

}

Write-Host "`nDONE"
