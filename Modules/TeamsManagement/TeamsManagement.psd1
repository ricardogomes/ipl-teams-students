@{
    RootModule = 'TeamsManagement.psm1'
    ModuleVersion = '1.0.0'
    GUID = 'f8b59e40-9c6d-4346-9c7a-6c53a8c3f2b2'
    Author = 'IPLeiria Teams Management'
    Description = 'Module for managing Microsoft Teams users and channels'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Connect-TeamSession',
        'Get-TeamByName',
        'Get-TeamChannels',
        'Import-ChannelUsers',
        'Add-UserToChannel',
        'Write-OperationSummary',
        'Read-ValidInteger',
        'New-PrivateChannel',
        'Add-TeamMember'
    )
    PrivateData = @{
        PSData = @{
            Tags = @('MicrosoftTeams', 'UserManagement')
        }
    }
}
