# IPL Teams User Creator

These scripts allow for the insertion of students into an existing Team, and also Channels (for a subset of those students).

The inputs are csv files with emails.

## Preparation

These scripts are Powershell scripts, and so require it. It's possible to install Powershell in MacOS and Linux.

Having Powershell, the next step is to install the Teams Module. The recomendend way is to use the PowerShellGet tool.

The following commands need to be executed in a Adminstrator priviledged Powershell (right-click Powershell -> Run as Administrator).

```powershell
Install-Module -Name PowerShellGet -Force -AllowClobber
```

Next install the Teams Module

```powershell
Install-Module -Name MicrosoftTeams -Force -AllowClobber
```

Now, depending on the current Powershell config, there are some things to address, the most important is to set the Execution Policy to AllSigned.

```powershell
Set-ExecutionPolicy AllSigned
```

Following that the next is to import the Teams module.

```powershell
Import-Module MicrosoftTeams
```

## Data

These scripts require CSV files under the input directory (and channels subdirectory). For the initial student creation the script expects a file under inputs called **students.csv**. For the channels the user can choose which file to process for each Channel.

For both the initial students and the channels the files need to have a field called **email** (can be seen in the examples). The files can have other fields (they are not considered).

## Execution

Team creation carries the normal issues like having an unique name.

**For simplicity these scripts assume that both the Team and the Channels are already created**.

These scripts are not digitally signed so before their execution the Execution Policy must be set to **Unrestricted**.

```powershell
Set-ExecutionPolicy Unrestricted
```

The first script to run should be the `.\addStudents.ps1`. This script adds each student in the file `./inputs/students.csv` to the specified Team (The script will ask for the Team name).

After that one, the `.\addStudentsToChannel.ps1` can opcionally be run for each Channel. This script will allow for the Channel selection and list the files under `./inputs/channels/` also allowing for its selection.

The first step in the execution is to logon to the Teams instance. This happens in every execution, on multiple subsequent executions the `Connect-MicrosoftTeams | Out-Null` line can be commented (#) as the session will persist.

After their execution the Execution Policy should be reset to **AllSigned**.

```powershell
Set-ExecutionPolicy AllSigned
```

## Resources

- [Installing Powershell in Linux](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-linux?view=powershell-7.3).
- [Installing Powershell in MacOS](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-macos?view=powershell-7.3)
- [Installing Powershell Teams Module](https://learn.microsoft.com/en-us/microsoftteams/teams-powershell-install)
- [Teams Powershell Cmdlet Reference](https://learn.microsoft.com/en-us/powershell/teams/?view=teams-ps)
