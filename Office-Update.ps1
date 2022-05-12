<#PSScriptInfo

.VERSION 22.03.25

.GUID 72cb5483-744e-4a7d-bcad-e04462ea2c2e

.AUTHOR Mike Galvin Contact: mike@gal.vin / twitter.com/mikegalvin_ / discord.gg/5ZsnJ5k

.COMPANYNAME Mike Galvin

.COPYRIGHT (C) Mike Galvin. All rights reserved.

.TAGS Office 2022 2019 365 Click-to-run C2R updates

.LICENSEURI

.PROJECTURI https://gal.vin/utils/office-update-utility/

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

#>

<#
    .SYNOPSIS
    Office Update Utility - Microsoft Office Update Manager.

    .DESCRIPTION
    Checks for updates of Microsoft Office

    To send a log file via e-mail using ssl and an SMTP password you must generate an encrypted password file.
    The password file is unique to both the user and machine.

    To create the password file run this command as the user and on the machine that will use the file:

    $creds = Get-Credential
    $creds.Password | ConvertFrom-SecureString | Set-Content C:\scripts\ps-script-pwd.txt

    .PARAMETER Office
    The folder containing the Office Deployment Tool (ODT).

    .PARAMETER Config
    The name of the configuration xml file for the Office Deployment Tool.
    It must be located in the same folder as the Office deployment tool.

    .PARAMETER Days
    The number of days that you wish to keep old update files for.
    If you do not configure this option, no old update files will be removed.

    .PARAMETER NoBanner
    Use this option to hide the ASCII art title in the console.

    .PARAMETER L
    The path to output the log file to.
    The file name will be Office-Update_YYYY-MM-dd_HH-mm-ss.log
    Do not add a trailing \ backslash.

    .PARAMETER Subject
    The subject line for the e-mail log. Encapsulate with single or double quotes.
    If no subject is specified, the default of "Office Update Utility Log" will be used.

    .PARAMETER SendTo
    The e-mail address the log should be sent to.

    .PARAMETER From
    The e-mail address the log should be sent from.

    .PARAMETER Smtp
    The DNS name or IP address of the SMTP server.

    .PARAMETER Port
    The Port that should be used for the SMTP server.

    .PARAMETER User
    The user account to authenticate to the SMTP server.

    .PARAMETER Pwd
    The txt file containing the encrypted password for SMTP authentication.

    .PARAMETER UseSsl
    Configures the utility to connect to the SMTP server using SSL.

    .EXAMPLE
    Office-Update.ps1 -Office \\Apps01\Software\Office365 -Config config-365-x64.xml -Days 30 -L C:\scripts\logs -Subject 'Server: Office Update'
    -SendTo me@contoso.com -From Office-Update@contoso.com -Smtp smtp-mail.outlook.com -User me@contoso.com -Pwd P@ssw0rd -UseSsl

    The above command will download any Office updates for the version and channel configured in config-365-x64.xml to the Office files
    directory \\Apps01\Software\Office365. Any update files older than 30 days will be removed. If the download is successful the log
    file will be output to C:\scripts\logs and e-mailed with a custom subject line.
#>

## Set up command line switches.
[CmdletBinding()]
Param(
    [alias("Office")]
    [ValidateScript({Test-Path $_ -PathType 'Container'})]
    $OfficeSrc,
    [alias("Config")]
    $Cfg,
    [alias("Days")]
    $Time,
    [alias("L")]
    $LogPath,
    [alias("LogRotate")]
    $LogHistory,
    [alias("Subject")]
    $MailSubject,
    [alias("SendTo")]
    $MailTo,
    [alias("From")]
    $MailFrom,
    [alias("Smtp")]
    $SmtpServer,
    [alias("Port")]
    $SmtpPort,
    [alias("User")]
    $SmtpUser,
    [alias("Pwd")]
    [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
    $SmtpPwd,
    [switch]$UseSsl,
    [switch]$Help,
    [switch]$NoBanner)

If ($NoBanner -eq $False)
{
    Write-Host -Object ""
    Write-Host -ForegroundColor Yellow -BackgroundColor Black -Object "     ___  __  __ _                            _       _         "
    Write-Host -ForegroundColor Yellow -BackgroundColor Black -Object "    /___\/ _|/ _(_) ___ ___   /\ /\ _ __   __| | __ _| |_ ___   "
    Write-Host -ForegroundColor Yellow -BackgroundColor Black -Object "   //  // |_| |_| |/ __/ _ \ / / \ \ '_ \ / _  |/ _  | __/ _ \  "
    Write-Host -ForegroundColor Yellow -BackgroundColor Black -Object "  / \_//|  _|  _| | (_|  __/ \ \_/ / |_) | (_| | (_| | ||  __/  "
    Write-Host -ForegroundColor Yellow -BackgroundColor Black -Object "  \___/ |_| |_| |_|\___\___|  \___/| .__/ \__,_|\__,_|\__\___|  "
    Write-Host -ForegroundColor Yellow -BackgroundColor Black -Object "                                   |_|                          "
    Write-Host -ForegroundColor Yellow -BackgroundColor Black -Object "         _   _ _ _ _                                            "
    Write-Host -ForegroundColor Yellow -BackgroundColor Black -Object "   /\ /\| |_(_) (_) |_ _   _             Mike Galvin            "
    Write-Host -ForegroundColor Yellow -BackgroundColor Black -Object "  / / \ \ __| | | | __| | | |          https://gal.vin          "
    Write-Host -ForegroundColor Yellow -BackgroundColor Black -Object "  \ \_/ / |_| | | | |_| |_| |                                   "
    Write-Host -ForegroundColor Yellow -BackgroundColor Black -Object "   \___/ \__|_|_|_|\__|\__, |                                   "
    Write-Host -ForegroundColor Yellow -BackgroundColor Black -Object "                       |___/          Version 22.03.25          "
    Write-Host -ForegroundColor Yellow -BackgroundColor Black -Object "                                     See -help for usage        "
    Write-Host -Object ""
}

If ($PSBoundParameters.Values.Count -eq 0 -or $Help)
{
    Write-Host "Usage:"
    Write-Host "From an elevated terminal run: [path\]Office-Update.ps1 -Office [path\]Office365 -Config config-365-x64.xml -Days 30"
    Write-Host "This will update the office installation files in the specified directory."
    Write-Host ""
    Write-Host ""
    Write-Host "To output a log: -L [path]. To remove logs produced by the utility older than X days: -LogRotate [number]."
    Write-Host "Run with no ASCII banner: -NoBanner"
    Write-Host ""
}

else {
    ## If logging is configured, start logging.
    ## If the log file already exists, clear it.
    If ($LogPath)
    {
        ## Make sure the log directory exists.
        $LogPathFolderT = Test-Path $LogPath

        If ($LogPathFolderT -eq $False)
        {
            New-Item $LogPath -ItemType Directory -Force | Out-Null
        }

        $LogFile = ("Office-Update_{0:yyyy-MM-dd_HH-mm-ss}.log" -f (Get-Date))
        $Log = "$LogPath\$LogFile"

        $LogT = Test-Path -Path $Log

        If ($LogT)
        {
            Clear-Content -Path $Log
        }

        Add-Content -Path $Log -Encoding ASCII -Value "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") [INFO] Log started"
    }

    ## Function to get date in specific format.
    Function Get-DateFormat
    {
        Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }

    ## Function for logging.
    Function Write-Log($Type, $Evt)
    {
        If ($Type -eq "Info")
        {
            If ($Null -ne $LogPath)
            {
                Add-Content -Path $Log -Encoding ASCII -Value "$(Get-DateFormat) [INFO] $Evt"
            }
            
            Write-Host -Object "$(Get-DateFormat) [INFO] $Evt"
        }

        If ($Type -eq "Succ")
        {
            If ($Null -ne $LogPath)
            {
                Add-Content -Path $Log -Encoding ASCII -Value "$(Get-DateFormat) [SUCCESS] $Evt"
            }

            Write-Host -ForegroundColor Green -Object "$(Get-DateFormat) [SUCCESS] $Evt"
        }

        If ($Type -eq "Err")
        {
            If ($Null -ne $LogPath)
            {
                Add-Content -Path $Log -Encoding ASCII -Value "$(Get-DateFormat) [ERROR] $Evt"
            }

            Write-Host -ForegroundColor Red -BackgroundColor Black -Object "$(Get-DateFormat) [ERROR] $Evt"
        }

        If ($Type -eq "Conf")
        {
            If ($Null -ne $LogPath)
            {
                Add-Content -Path $Log -Encoding ASCII -Value "$Evt"
            }

            Write-Host -ForegroundColor Cyan -Object "$Evt"
        }
    }

    ## getting Windows Version info
    $OSVMaj = [environment]::OSVersion.Version | Select-Object -expand major
    $OSVMin = [environment]::OSVersion.Version | Select-Object -expand minor
    $OSVBui = [environment]::OSVersion.Version | Select-Object -expand build
    $OSV = "$OSVMaj" + "." + "$OSVMin" + "." + "$OSVBui"

    ##
    ## Display the current config and log if configured.
    ##

    Write-Log -Type Conf -Evt "************ Running with the following config *************."
    Write-Log -Type Conf -Evt "Utility Version:.......22.03.25"
    Write-Log -Type Conf -Evt "Hostname:..............$Env:ComputerName."
    Write-Log -Type Conf -Evt "Windows Version:.......$OSV."
    If ($Null -ne $OfficeSrc)
    {
        Write-Log -Type Conf -Evt "Office folder:.........$OfficeSrc."
    }

    If ($Null -ne $Cfg)
    {
        Write-Log -Type Conf -Evt "Config file:...........$Cfg."
    }

    If ($Null -ne $Time)
    {
        Write-Log -Type Conf -Evt "Days to keep updates:..$Time days."
    }

    If ($Null -ne $LogPath)
    {
        Write-Log -Type Conf -Evt "Logs directory:........$LogPath."
    }

    If ($Null -ne $LogHistory)
    {
        Write-Log -Type Conf -Evt "Logs to keep:..........$LogHistory days"
    }

    If ($MailTo)
    {
        Write-Log -Type Conf -Evt "E-mail log to:.........$MailTo."
    }

    If ($MailFrom)
    {
        Write-Log -Type Conf -Evt "E-mail log from:.......$MailFrom."
    }

    If ($MailSubject)
    {
        Write-Log -Type Conf -Evt "E-mail subject:........$MailSubject."
    }

    else {
        Write-Log -Type Conf -Evt "E-mail subject:........Default"
    }

    If ($SmtpServer)
    {
        Write-Log -Type Conf -Evt "SMTP server is:........$SmtpServer."
    }

    If ($SmtpPort)
    {
        Write-Log -Type Conf -Evt "SMTP Port:.............$SmtpPort."
    }

    else {
        Write-Log -Type Conf -Evt "SMTP Port:.............Default"
    }

    If ($SmtpUser)
    {
        Write-Log -Type Conf -Evt "SMTP user is:..........$SmtpUser."
    }

    If ($SmtpPwd)
    {
        Write-Log -Type Conf -Evt "SMTP pwd file:.........$SmtpPwd."
    }

    Write-Log -Type Conf -Evt "-UseSSL switch is:.....$UseSsl."
    Write-Log -Type Conf -Evt "************************************************************"
    Write-Log -Type Info -Evt "Process started"

    ##
    ## Display current config ends here.
    ##

    #Run update process.
    & $OfficeSrc\setup.exe /download $OfficeSrc\$Cfg

    ## Location of the office source files.
    $UpdateFolder = "$OfficeSrc\Office\Data"

    ## Check the last write time of the office source files folder if it is greater than the previous day.
    $Updated = (Get-ChildItem -Path $UpdateFolder | Where-Object CreationTime -gt (Get-Date).AddDays(-1)).Count

    ## If the Updated variable returns as not 0 then continue.
    If ($Updated -ne 0)
    {
        $VerName = Get-ChildItem -Path $UpdateFolder -Directory | Sort-Object LastWriteTime | Select-Object -last 1 | Select-Object -ExpandProperty Name
        Write-Log -Type Info -Evt "Office source files were updated."
        Write-Log -Type Info -Evt "Latest version is: $VerName"

        If ($Null -ne $Time)
        {
            $FilesToDel = Get-ChildItem -Path $UpdateFolder | Where-Object LastWriteTime –lt (Get-Date).AddDays(-$Time)

            If ($FilesToDel.count -ne 0)
            {
                Write-Log -Type Info -Evt "The following old Office files were removed:"
                Get-ChildItem -Path $UpdateFolder | Where-Object LastWriteTime –lt (Get-Date).AddDays(-$Time)
                Get-ChildItem -Path $UpdateFolder | Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-$Time)} | Select-Object -Property Name, LastWriteTime | Format-Table -HideTableHeaders | Out-File -Append $Log -Encoding ASCII

                ## If configured, remove the old files.
                Get-ChildItem $UpdateFolder | Where-Object {$_.LastWriteTime –lt (Get-Date).AddDays(-$Time)} | Remove-Item -Recurse
            }
        }

        Write-Log -Type Info -Evt "Process finished"

        ## If logging is configured then finish the log file.
        If ($LogPath)
        {
            Add-Content -Path $Log -Encoding ASCII -Value "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") [INFO] Log finished"

            ## This whole block is for e-mail, if it is configured.
            If ($SmtpServer)
            {
                ## Default e-mail subject if none is configured.
                If ($Null -eq $MailSubject)
                {
                    $MailSubject = "Office Update Utility Log"
                }

                ## Default Smtp Port if none is configured.
                If ($Null -eq $SmtpPort)
                {
                    $SmtpPort = "25"
                }

                ## Setting the contents of the log to be the e-mail body. 
                $MailBody = Get-Content -Path $Log | Out-String

                ## If an smtp password is configured, get the username and password together for authentication.
                ## If an smtp password is not provided then send the e-mail without authentication and obviously no SSL.
                If ($SmtpPwd)
                {
                    $SmtpPwdEncrypt = Get-Content $SmtpPwd | ConvertTo-SecureString
                    $SmtpCreds = New-Object System.Management.Automation.PSCredential -ArgumentList ($SmtpUser, $SmtpPwdEncrypt)

                    ## If -ssl switch is used, send the email with SSL.
                    ## If it isn't then don't use SSL, but still authenticate with the credentials.
                    If ($UseSsl)
                    {
                        Send-MailMessage -To $MailTo -From $MailFrom -Subject $MailSubject -Body $MailBody -SmtpServer $SmtpServer -Port $SmtpPort -UseSsl -Credential $SmtpCreds
                    }

                    else {
                        Send-MailMessage -To $MailTo -From $MailFrom -Subject $MailSubject -Body $MailBody -SmtpServer $SmtpServer -Port $SmtpPort -Credential $SmtpCreds
                    }
                }

                else {
                    Send-MailMessage -To $MailTo -From $MailFrom -Subject $MailSubject -Body $MailBody -SmtpServer $SmtpServer -Port $SmtpPort
                }
            }
        }

        If ($Null -ne $LogHistory)
        {
            ## Cleanup logs.
            Write-Log -Type Info -Evt "Deleting logs older than: $LogHistory days"
            Get-ChildItem -Path "$LogPath\Office-Update_*" -File | Where-Object CreationTime -lt (Get-Date).AddDays(-$LogHistory) | Remove-Item -Recurse
        }
    }

    Write-Log -Type Info -Evt "No updates."
    Write-Log -Type Info -Evt "Process finished"
}
## End