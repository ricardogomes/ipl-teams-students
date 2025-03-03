function Connect-TeamSession {
    try {
        Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
        Write-Host "Successfully connected to Microsoft Teams." -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Failed to connect to Microsoft Teams. Please check your credentials and try again."
        return $false
    }
}

function Get-TeamByName {
    param (
        [Parameter(Mandatory=$true)]
        [string]$TeamName
    )
    
    try {
        $Team = Get-Team -DisplayName $TeamName -ErrorAction Stop
        Write-Host "Team '$TeamName' found with GroupId: $($Team.GroupId)" -ForegroundColor Green
        return $Team
    } catch {
        Write-Error "Team '$TeamName' not found. Please ensure the team exists and you have the necessary permissions."
        return $null
    }
}

function Get-TeamChannels {
    param (
        [Parameter(Mandatory=$true)]
        [string]$GroupId
    )
    
    try {
        Write-Host "Retrieving existing channels in the team..."
        $ExistingChannels = Get-TeamChannel -GroupId $GroupId -ErrorAction Stop
        Write-Host "Found $($ExistingChannels.Count) channels in team." -ForegroundColor Green
        return $ExistingChannels
    } catch {
        Write-Error "Failed to retrieve channels. Error: $_"
        return $null
    }
}

function Import-ChannelUsers {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    try {
        $Users = Import-Csv -Path $FilePath -ErrorAction Stop
        if ($Users.Count -eq 0) {
            Write-Warning "CSV file '$FilePath' is empty."
            return $null
        }
        return $Users
    } catch {
        Write-Warning "Failed to import CSV file '$FilePath'. Error: $_"
        return $null
    }
}

function Add-UserToChannel {
    param (
        [Parameter(Mandatory=$true)]
        [string]$GroupId,
        
        [Parameter(Mandatory=$true)]
        [string]$ChannelName,
        
        [Parameter(Mandatory=$true)]
        [string]$UserEmail
    )
    
    if ([string]::IsNullOrWhiteSpace($UserEmail)) {
        Write-Host "`tSkipping empty email entry." -ForegroundColor Yellow
        return $false
    }

    Write-Host "`tAdding user '$UserEmail' to channel '$ChannelName'..." -NoNewLine

    try {
        Add-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName -User $UserEmail -ErrorAction Stop
        Write-Host " Success" -ForegroundColor Green
        return $true
    } catch {
        Write-Host " Failed" -ForegroundColor Red
        Write-Host "`t`tError adding '$UserEmail' to '$ChannelName': $_" -ForegroundColor Red
        return $false
    }
}

function Write-OperationSummary {
    param (
        [int]$TotalFiles,
        [int]$ProcessedFiles,
        [int]$SkippedFiles,
        [int]$AddedUsers,
        [int]$FailedAdds
    )
    
    Write-Host "`n=== Operation Summary ===" -ForegroundColor Yellow
    Write-Host "Total CSV Files Found: $TotalFiles"
    Write-Host "Files Processed: $ProcessedFiles"
    Write-Host "Files Skipped: $SkippedFiles"
    Write-Host "Total Users Added Successfully: $AddedUsers" -ForegroundColor Green
    Write-Host "Total Failed User Additions: $FailedAdds" -ForegroundColor Red
    Write-Host "=========================`n" -ForegroundColor Yellow
}

function Read-ValidInteger {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Prompt,
        [Parameter(Mandatory=$false)]
        [int]$MinValue = 0
    )
    
    do {
        $input = Read-Host $Prompt
        if ([int]::TryParse($input, [ref]$null)) {
            $value = [int]$input
            if ($value -ge $MinValue) {
                return $value
            }
            Write-Host "Please enter a number greater than or equal to $MinValue." -ForegroundColor Yellow
        } else {
            Write-Host "Invalid input. Please enter a valid integer." -ForegroundColor Yellow
        }
    } until ($false)
}

function New-PrivateChannel {
    param (
        [Parameter(Mandatory=$true)]
        [string]$GroupId,
        
        [Parameter(Mandatory=$true)]
        [string]$ChannelName,
        
        [Parameter(Mandatory=$false)]
        [string]$Description = "Private channel for $ChannelName"
    )
    
    Write-Host "Creating private channel: $ChannelName"

    try {
        New-TeamChannel `
            -GroupId $GroupId `
            -DisplayName $ChannelName `
            -Description $Description `
            -MembershipType Private

        Write-Host "Channel '$ChannelName' created successfully." -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Failed to create channel '$ChannelName'. Error: $_"
        return $false
    }
}

function Add-TeamMember {
    param (
        [Parameter(Mandatory=$true)]
        [string]$GroupId,
        
        [Parameter(Mandatory=$true)]
        [string]$UserEmail
    )
    
    try {
        Add-TeamUser -GroupId $GroupId -User $UserEmail -ErrorAction Stop
        Write-Host "`tSuccessfully added $UserEmail" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "`tFailed to add $UserEmail. Error: $_"
        return $false
    }
}

function Get-TeamOwners {
    param (
        [Parameter(Mandatory=$true)]
        [string]$GroupId
    )
    
    try {
        $TeamUsers = Get-TeamUser -GroupId $GroupId -Role Owner -ErrorAction Stop
        Write-Host "Found $($TeamUsers.Count) team owners." -ForegroundColor Green
        return $TeamUsers
    } catch {
        Write-Error "Failed to retrieve team owners. Error: $_"
        return $null
    }
}

function Add-ChannelOwner {
    param (
        [Parameter(Mandatory=$true)]
        [string]$GroupId,
        
        [Parameter(Mandatory=$true)]
        [string]$ChannelName,
        
        [Parameter(Mandatory=$true)]
        [string]$UserEmail
    )
    
    try {
        Add-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName -User $UserEmail -Role Owner -ErrorAction Stop
        Write-Host "`tAdded $UserEmail as owner to channel '$ChannelName'" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "`tFailed to add $UserEmail as owner to channel '$ChannelName'. Error: $_"
        return $false
    }
}
