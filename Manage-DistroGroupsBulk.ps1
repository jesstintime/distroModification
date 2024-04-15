<#
    .SYNOPSIS
    Easily allows bulk adding or removal of distribution group members. Includes error handling.

    .DESCRIPTION
    The script will provide a fully guided and prompted user experience to manage the distribution groups.
    Inputs can accept both TXT and CSV formats and the script will terminate and abort all processes if there
    are mistakes made during the initial input phase.

    .PARAMETER operationChoice
    Specifies if the script will add or remove members.

    .PARAMETER groupChoice
    Stores the name of the group to be modified.

    .PARAMETER userList
    Allows the upload of either CSV or TXT for member processing.

    .PARAMETER confirmCheck
    Final confirmation to ensure the correct choices were selected.

    .EXAMPLE
    PS> .\Manage-DistroGroupsBulk.ps1

    .NOTES
    The Exchange Online PS Module is required to run this script.
    ErrorActionPreference is modified during the script. Please take note to examine the preference after the script
    to ensure it is set as desired.
#>

# Setting up logging and error handling.
$ErrorActionPreference = 'Stop'
$logFolder = $env:TEMP + "\ManageDistroGroups_$(Get-Date -Format MMddyyyy_mmssss)"
if (!(Test-Path $logFolder)) {
    New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
}
$successLog = $logFolder + "\Success - $(Get-Date -Format MM-dd-yyyy).log"
$errorLog = $logFolder + "\Errors - $(Get-Date -Format MM-dd-yyyy).log"
$addedLog = $logFolder + "\Added Successfully - $(Get-Date -Format MM-dd-yyyy).txt"
$notaddedLog = $logFolder + "\Unable to Add - $(Get-Date -Format MM-dd-yyyy).txt"
$error.Clear()

# Guided prompts to begin and choice group and input file for modification.
Write-Host "Welcome to the bulk management Distro script. Please begin by authenticating to Exchange Online."
Write-Host ""
Start-Sleep -Seconds 3
Connect-ExchangeOnline
Write-Host ""
$operationChoice = Read-Host "Please enter the desired operation - [A]dd, [R]emove"
if (($operationChoice -ne 'A') -and ($operationChoice -ne 'R')) {
    Write-Error "Invalid entry. Please retry script."
}
$groupChoice = Read-Host "Please enter the EXACT name of the Group you wish to manage"
$group = Get-DistributionGroup -Identity $groupChoice -ErrorAction 'SilentlyContinue'
if ($group.Count -eq 0) {
    Write-Error "Invalid entry. Please retry script."
}
Write-Host ""
Write-Host "A selection screen will appear shortly. Please select your data file for the script."
Start-Sleep -Seconds 5
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Title = 'Choose data file for script'
}
$null = $FileBrowser.ShowDialog()
if ($FileBrowser.FileName -like '*.csv') {
    $userList = Import-Csv -Path $FileBrowser.FileName
    $inputType = "CSV"
}
elseif ($FileBrowser.FileName -like '*.txt') {
    $userList = Get-Content -Path $FileBrowser.FileName
    $inputType = "TXT"
}
else {
    Write-Error "Invalid file type chosen. Please retry script."
}
$firstUser = $userList[0] -replace "@{","" -replace "}",""
$secondUser = $userList[1] -replace "@{","" -replace "}",""
$thirdUser = $userList[2] -replace "@{","" -replace "}",""

Write-Host ""
Write-Host "Summary"
Write-Host "-------"
Write-Host ""
Write-Host "First 3 entries of User List:"
Write-Host "$($firstUser)"
Write-Host "$($secondUser)"
Write-Host "$($thirdUser)"
Write-Host ""
Write-Host "Total number of users in list: $($userList.Count)"
Write-Host ""
if ($operationChoice -eq 'A') {
    Write-Host "Being ADDED to the group $($group.DisplayName)/$($group.PrimarySmtpAddress)"
}
else {
    Write-Host "Being REMOVED from the group $($group.DisplayName)/$($group.PrimarySmtpAddress)"  
}
Write-Host ""
$confirmCheck = Read-Host "Does this information look correct? [Y]es, [N]o"
if ($confirmCheck -eq 'N') {
    Exit
}
elseif (($confirmCheck -ne 'N') -and ($confirmCheck -ne 'Y')) {
    Write-Error "Invalid entry. Please retry script."
}
# Proceed only if the user confirms the input information is as intended.
elseif ($confirmCheck -eq 'Y') {
    Write-Host ""
    $ErrorActionPreference = "SilentlyContinue"
    $count = 1
    $successReport = New-Object System.Collections.ArrayList @()
    $failReport = New-Object System.Collections.ArrayList @()
    if ($inputType -eq 'CSV') {
        if ($operationChoice -eq 'A') {
            ForEach ($user in $userList) {
                Write-Output "Processing $($user.Email) - $($count) of $($userList.Count)..."
                $removeSpace = $user.Email -replace " ",""
                Try {
                    Add-DistributionGroupMember -Identity $group -Member $removeSpace -ErrorAction Stop
                    Add-Content -Path $successLog -Value "$($removeSpace) has been added to $($group) successfully."
                    [void]$successReport.Add($removeSpace)
                }
                Catch {
                    Add-Content -Path $errorLog -Value "Failed to add $($removeSpace) to $($group) with the following error:"
                    Add-Content -Path $errorLog -Value $error[0].Exception
                    Add-Content -Path $errorLog -Value ""
                    [void]$failReport.Add($removeSpace)
                }
                $count++
            }
        }
        elseif ($operationChoice -eq 'R') {
            ForEach ($user in $userList) {
                Write-Output "Processing $($user.Email) - $($count) of $($userList.Count)..."
                $removeSpace = $user.Email -replace " ",""
                Try {
                    Remove-DistributionGroupMember -Identity $group -Member $removeSpace -ErrorAction Stop -Confirm:$false
                    Add-Content -Path $successLog -Value "$($removeSpace) has been removed from $($group) successfully."
                    [void]$successReport.Add($removeSpace)
                }
                Catch {
                    Add-Content -Path $errorLog -Value "Failed to remove $($removeSpace) from $($group) with the following error:"
                    Add-Content -Path $errorLog -Value $error[0].Exception
                    Add-Content -Path $errorLog -Value ""
                    [void]$failReport.Add($removeSpace)
                }
                $count++
            }
        }
        $successReport | Out-File -FilePath $addedLog
        $failReport | Out-File -FilePath $notaddedLog
    }
    elseif ($inputType -eq 'TXT') {
        if ($operationChoice -eq 'A') {
            ForEach ($user in $userList) {
                Write-Output "Processing $($user) - $($count) of $($userList.Count)..."
                $removeSpace = $user -replace " ",""
                Try {
                    Add-DistributionGroupMember -Identity $group -Member $removeSpace -ErrorAction Stop
                    Add-Content -Path $successLog -Value "$($removeSpace) has been added to $($group) successfully."
                    [void]$successReport.Add($removeSpace)
                }
                Catch {
                    Add-Content -Path $errorLog -Value "Failed to add $($removeSpace) to $($group) with the following error:"
                    Add-Content -Path $errorLog -Value $error[0].Exception
                    Add-Content -Path $errorLog -Value ""
                    [void]$failReport.Add($removeSpace)
                }
                $count++
            }
        }
        elseif ($operationChoice -eq 'R') {
            ForEach ($user in $userList) {
                Write-Output "Processing $($user) - $($count) of $($userList.Count)..."
                $removeSpace = $user -replace " ",""
                Try {
                    Remove-DistributionGroupMember -Identity $group -Member $removeSpace -ErrorAction Stop -Confirm:$false
                    Add-Content -Path $successLog -Value "$($removeSpace) has been removed from $($group) successfully."
                    [void]$successReport.Add($removeSpace)
                }
                Catch {
                    Add-Content -Path $errorLog -Value "Failed to remove $($removeSpace) from $($group) with the following error:"
                    Add-Content -Path $errorLog -Value $error[0].Exception
                    Add-Content -Path $errorLog -Value ""
                    [void]$failReport.Add($removeSpace)
                }
                $count++
            }
        }
        $successReport | Out-File -FilePath $addedLog
        $failReport | Out-File -FilePath $notaddedLog
    }
    Write-Host ""
    Write-Host ""
    # Report on results and display logging folder for review of results and errors, if applicable.
    Write-Host "Script has completed successfully. In total, $($error.Count) errors occurred during the script. In a moment, the error logging folder will open for review as needed."
    Start-Sleep -Seconds 10
    Invoke-Item $logFolder
}
Exit