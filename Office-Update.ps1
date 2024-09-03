<#PSScriptInfo

.VERSION 23.04.28

.GUID 72cb5483-744e-4a7d-bcad-e04462ea2c2e

.AUTHOR Mike Galvin Contact: digressive@outlook.com

.COMPANYNAME Mike Galvin

.COPYRIGHT (C) Mike Galvin. All rights reserved.

.TAGS Office 2022 2019 365 Click-to-run C2R updates

.LICENSEURI https://github.com/Digressive/Office-Update?tab=MIT-1-ov-file

.PROJECTURI https://gal.vin/utils/office-update-utility/

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

#>

<#
    .SYNOPSIS
    Office Update Utility - Microsoft Office Update Manager

    .DESCRIPTION
    Checks for updates of Microsoft Office and removes old versions.
    Run with -help or no arguments for usage.
#>

## Set up command line switches.
[CmdletBinding()]
Param(
    [alias("Office")]
    $OfficeSrc,
    [alias("Config")]
    $Cfg,
    [alias("L")]
    $LogPathUsr,
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
    [Alias("Webhook")]
    [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
    [string]$Webh,
    [switch]$UseSsl,
    [switch]$Help,
    [switch]$NoBanner)

If ($NoBanner -eq $False)
{
    Write-Host -ForegroundColor Yellow -BackgroundColor Black -Object "
           ___  __  __ _                            _       _           
          /___\/ _|/ _(_) ___ ___   /\ /\ _ __   __| | __ _| |_ ___     
         //  // |_| |_| |/ __/ _ \ / / \ \ '_ \ / _  |/ _  | __/ _ \    
        / \_//|  _|  _| | (_|  __/ \ \_/ / |_) | (_| | (_| | ||  __/    
        \___/ |_| |_| |_|\___\___|  \___/| .__/ \__,_|\__,_|\__\___|    
                                         |_|                            
               _   _ _ _ _                                              
         /\ /\| |_(_) (_) |_ _   _             Mike Galvin              
        / / \ \ __| | | | __| | | |          https://gal.vin            
        \ \_/ / |_| | | | |_| |_| |                                     
         \___/ \__|_|_|_|\__|\__, |         Version 23.04.28            
                             |___/         See -help for usage          
                                                                        
                  Donate: https://www.paypal.me/digressive              
"
}

If ($PSBoundParameters.Values.Count -eq 0 -or $Help)
{
    Write-Host -Object "Usage:
    From a terminal run: [path\]Office-Update.ps1 -Office [path\] -Config [file name.xml]
    This will update the office installation files in the specified directory, and delete the old update files.

    To output a log: -L [path\].
    To remove logs produced by the utility older than X days: -LogRotate [number].
    Run with no ASCII banner: -NoBanner

    To use the 'email log' function:
    Specify the subject line with -Subject ""'[subject line]'"" If you leave this blank a default subject will be used
    Make sure to encapsulate it with double & single quotes as per the example for Powershell to read it correctly.

    Specify the 'to' address with -SendTo [example@contoso.com]
    For multiple address, separate with a comma.

    Specify the 'from' address with -From [example@contoso.com]
    Specify the SMTP server with -Smtp [smtp server name]

    Specify the port to use with the SMTP server with -Port [port number].
    If none is specified then the default of 25 will be used.

    Specify the user to access SMTP with -User [example@contoso.com]
    Specify the password file to use with -Pwd [path\]ps-script-pwd.txt.
    Use SSL for SMTP server connection with -UseSsl.

    To generate an encrypted password file run the following commands
    on the computer and the user that will run the script:
"
    Write-Host -Object '    $creds = Get-Credential
    $creds.Password | ConvertFrom-SecureString | Set-Content [path\]ps-script-pwd.txt'
}

else {
    ## If logging is configured, start logging.
    ## If the log file already exists, clear it.
    If ($LogPathUsr)
    {
        ## Clean User entered string
        $LogPath = $LogPathUsr.trimend('\')

        ## Make sure the log directory exists.
        If ((Test-Path -Path $LogPath) -eq $False)
        {
            New-Item $LogPath -ItemType Directory -Force | Out-Null
        }

        $LogFile = ("Office-Update_{0:yyyy-MM-dd_HH-mm-ss}.log" -f (Get-Date))
        $Log = "$LogPath\$LogFile"

        If (Test-Path -Path $Log)
        {
            Clear-Content -Path $Log
        }
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
            If ($LogPathUsr)
            {
                Add-Content -Path $Log -Encoding ASCII -Value "$(Get-DateFormat) [INFO] $Evt"
            }
            
            Write-Host -Object "$(Get-DateFormat) [INFO] $Evt"
        }

        If ($Type -eq "Succ")
        {
            If ($LogPathUsr)
            {
                Add-Content -Path $Log -Encoding ASCII -Value "$(Get-DateFormat) [SUCCESS] $Evt"
            }

            Write-Host -ForegroundColor Green -Object "$(Get-DateFormat) [SUCCESS] $Evt"
        }

        If ($Type -eq "Err")
        {
            If ($LogPathUsr)
            {
                Add-Content -Path $Log -Encoding ASCII -Value "$(Get-DateFormat) [ERROR] $Evt"
            }

            Write-Host -ForegroundColor Red -BackgroundColor Black -Object "$(Get-DateFormat) [ERROR] $Evt"
        }

        If ($Type -eq "Conf")
        {
            If ($LogPathUsr)
            {
                Add-Content -Path $Log -Encoding ASCII -Value "$Evt"
            }

            Write-Host -ForegroundColor Cyan -Object "$Evt"
        }
    }

    ## Function for Update Check
    Function UpdateCheck()
    {
        $ScriptVersion = "23.04.28"
        $RawSource = "https://raw.githubusercontent.com/Digressive/Office-Update/master/Office-Update.ps1"

        try {
            $SourceCheck = Invoke-RestMethod -uri "$RawSource"
            $VerCheck = $SourceCheck -split '\n' | Select-String -Pattern ".VERSION $ScriptVersion" -SimpleMatch -CaseSensitive -Quiet

            If ($VerCheck -ne $True)
            {
                Write-Log -Type Conf -Evt "*** There is an update available. ***"
            }
        }

        catch {
        }
    }

    If ($Null -eq $OfficeSrc)
    {
        Write-Log -Type Err -Evt "You must specify -Office [path\]."
        Exit
    }

    else {
        If ($Null -eq $LogPathUsr -And $SmtpServer)
        {
            Write-Log -Type Err -Evt "You must specify -L [path\] to use the email log function."
            Exit
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
    Write-Log -Type Conf -Evt "--- Running with the following config ---"
    Write-Log -Type Conf -Evt "Utility Version: 23.04.28"
    UpdateCheck ## Run Update checker function
    Write-Log -Type Conf -Evt "Hostname: $Env:ComputerName."
    Write-Log -Type Conf -Evt "Windows Version: $OSV."
    If ($OfficeSrc)
    {
        Write-Log -Type Conf -Evt "Office folder: $OfficeSrc."
    }

    If ($Cfg)
    {
        Write-Log -Type Conf -Evt "Config file: $Cfg."
    }

    If ($LogPathUsr)
    {
        Write-Log -Type Conf -Evt "Logs directory: $LogPath."
    }

    If ($Null -ne $LogHistory)
    {
        Write-Log -Type Conf -Evt "Logs to keep: $LogHistory days."
    }

    If ($Webh)
    {
        Write-Log -Type Conf -Evt "Webhook file: $Webh."
    }

    If ($MailTo)
    {
        Write-Log -Type Conf -Evt "E-mail log to: $MailTo."
    }

    If ($MailFrom)
    {
        Write-Log -Type Conf -Evt "E-mail log from: $MailFrom."
    }

    If ($MailSubject)
    {
        Write-Log -Type Conf -Evt "E-mail subject: $MailSubject."
    }

    If ($SmtpServer)
    {
        Write-Log -Type Conf -Evt "SMTP server is: $SmtpServer."
    }

    If ($SmtpPort)
    {
        Write-Log -Type Conf -Evt "SMTP Port: $SmtpPort."
    }

    If ($SmtpUser)
    {
        Write-Log -Type Conf -Evt "SMTP auth: Configured"
    }
    Write-Log -Type Conf -Evt "---"
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
        $VerName = Get-ChildItem -Path $UpdateFolder -Name -Directory | Measure-Object -Maximum
        Write-Log -Type Info -Evt "Office source files were updated."
        Write-Log -Type Info -Evt "Latest version is: $($VerName.Maximum)"

        ## Remove old update folders and then files
        $Content = Get-ChildItem -Path $UpdateFolder

        If ($Content.count -gt 3)
        {
            $OffDirs = Get-ChildItem -Path $UpdateFolder -Name -Directory
            $resultDirs = $OffDirs | Measure-Object -Maximum

            Get-ChildItem -Path $UpdateFolder -Directory -Exclude $resultDirs.Maximum | Remove-Item -Recurse

            $OffFiles = Get-ChildItem -Path $UpdateFolder -Name -File -Exclude v64.cab
            $resultFiles = $OffFiles | Measure-Object -Maximum

            Get-ChildItem -Path $UpdateFolder -File | Where-Object {$_.Name -NotMatch "v64.cab" -And $_.Name -NotMatch "v32.cab" -And $_.Name -NotMatch $resultFiles.Maximum} | Remove-Item
        }

        Write-Log -Type Info -Evt "Process finished"

        ## This whole block is for e-mail, if it is configured.
        If ($SmtpServer)
        {
            If (Test-Path -Path $Log)
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

                ForEach ($MailAddress in $MailTo)
                {
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
                            Send-MailMessage -To $MailAddress -From $MailFrom -Subject $MailSubject -Body $MailBody -SmtpServer $SmtpServer -Port $SmtpPort -UseSsl -Credential $SmtpCreds
                        }

                        else {
                            Send-MailMessage -To $MailAddress -From $MailFrom -Subject $MailSubject -Body $MailBody -SmtpServer $SmtpServer -Port $SmtpPort -Credential $SmtpCreds
                        }
                    }

                    else {
                        Send-MailMessage -To $MailAddress -From $MailFrom -Subject $MailSubject -Body $MailBody -SmtpServer $SmtpServer -Port $SmtpPort
                    }
                }
            }

            else {
                Write-Host -ForegroundColor Red -BackgroundColor Black -Object "There's no log file to email."
            }
        }
        ## End of Email block

        ## Webhook block
        If ($Webh)
        {
            $WebHookUri = Get-Content $Webh
            $WebHookArr = @()

            $title       = "Office Update Utility"
            $description = Get-Content -Path $Log | Out-String

            $WebHookObj = [PSCustomObject]@{
                title = $title
                description = $description
            }

            $WebHookArr += $WebHookObj
            $payload = [PSCustomObject]@{
                embeds = $WebHookArr
            }

            Invoke-RestMethod -Uri $WebHookUri -Body ($payload | ConvertTo-Json -Depth 2) -Method Post -ContentType 'application/json'
        }
    }

    else {
        Write-Log -Type Info -Evt "No updates."
        Write-Log -Type Info -Evt "Process finished"
    }

    If ($Null -ne $LogHistory)
    {
        ## Cleanup logs.
        Write-Log -Type Info -Evt "Deleting logs older than: $LogHistory days"
        Get-ChildItem -Path "$LogPath\Office-Update_*" -File | Where-Object CreationTime -lt (Get-Date).AddDays(-$LogHistory) | Remove-Item -Recurse
    }
}
## End