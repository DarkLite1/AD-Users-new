#Requires -Version 5.1
#Requires -Modules ActiveDirectory, ImportExcel
#Requires -Modules Toolbox.ActiveDirectory, Toolbox.HTML, Toolbox.EventLog

<#
    .SYNOPSIS
        Create a list of all the user accounts that have been created in the
        last x days within the active directory.

    .DESCRIPTION
        The active directory is scanned for the last created users. An e-mail
        is sent with an Excel file in attachment containing the results.

    .PARAMETER ImportFile
        A .json file containing the script arguments.

    .PARAMETER LogFolder
        Location for the log files.
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\AD Reports\AD Users new\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import input file
        $File = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json

        if (-not ($MailTo = $File.MailTo)) {
            throw "Input file '$ImportFile': No 'MailTo' addresses found."
        }

        if (-not ($adOUs = $File.AD.OU)) {
            throw "Input file '$ImportFile': No 'AD.OU' found."
        }

        if (-not ($Days = $File.NewerThanDays)) {
            throw "Input file '$ImportFile': No 'NewerThanDays' found."
        }

        try {
            [int]$Days
        }
        catch {
            throw "Property 'NewerThanDays' needs to be a number, the value '$NewerThanDays' is not supported."
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        $NewUsers = Get-ADUserNewHC -OU $adOUs -Days $Days -EA Stop
        $adOUsHtml = $adOUs | ConvertTo-OuNameHC -OU | Sort-Object | ConvertTo-HtmlListHC -Header 'Organizational units:'

        Switch (($NewUsers | Measure-Object).Count) {
            '0' {
                $Intro = "<p><b>No new users</b> have been created in the last <b>$Days days</b>.</p>"
                $Subject = "No new users created in the last $Days days"
            }
            '1' {
                $Intro = "<p>Only <b>1 user</b> has been created in the last <b>$Days days</b>:</p>"
                $Subject = "1 new user created in the last $Days days"
            }
            Default {
                $Intro = "<p><b>$_ new users</b> have been created in the last <b>$Days days</b>:</p>"
                $Subject = "$_ new users created in the last $Days days"
            }
        }

        if ($NewUsers) {
            $ExcelParams = @{
                Path          = $LogFile + '.xlsx'
                AutoSize      = $true
                FreezeTopRow  = $true
                TableName     = "Users"
                WorkSheetName = "New users last $Days days"
            }
            $NewUsers | Export-Excel @ExcelParams -NoNumberConversion @(
                'Employee ID', 'OfficePhone', 'HomePhone', 'MobilePhone',
                'ipPhone', 'Fax', 'Pager'
            )

            $Table = $NewUsers | Group-Object Country |
            Select-Object @{
                Name       = "Country"
                Expression = { $_.Name }
            },
            @{
                Name       = "Total"
                Expression = { $_.Count }
            } |
            Sort-Object Count -Descending |
            ConvertTo-Html -As Table -Fragment

            $Message = "$Intro
                        $Table
                        <h3>Summary:</h3>
                        $($NewUsers | Sort-Object 'Country', 'Display name' | Select-Object 'Display name',
                            'Manager','Company','Type of account',
                            @{Name='Account expires';Expression={if($_.'Account expires' -eq 'Never'){'Never'}
                            else{$_.'Account expires'.ToString('dd/MM/yyyy')}}} |
                            ConvertTo-Html -Fragment -As Table)
                        <p><i>* Check the attachment(s) for details</i></p>"
        }
        else {
            $Message = $Intro
        }

        $EmailParams = @{
            To          = $MailTo
            Bcc         = $ScriptAdmin
            Subject     = $Subject
            Message     = $Message, $adOUsHtml
            Attachments = $ExcelParams.Path
            LogFolder   = $LogParams.LogFolder
            Header      = $ScriptName
            Save        = $LogFile + ' - Mail.html'
        }
        Remove-EmptyParamsHC $EmailParams
        Get-ScriptRuntimeHC -Stop
        Send-MailHC @EmailParams
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"; Exit 1
    }
    Finally {
        Get-Job | Remove-Job -Force
        Write-EventLog @EventEndParams
    }
}