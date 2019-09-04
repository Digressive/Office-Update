<#PSScriptInfo

.VERSION 1.1

.GUID 72cb5483-744e-4a7d-bcad-e04462ea2c2e

.AUTHOR Mike Galvin Contact: mike@gal.vin twitter.com/mikegalvin_

.COMPANYNAME Mike Galvin

.COPYRIGHT (C) Mike Galvin. All rights reserved.

.TAGS Office 2019 365 Click-to-run C2R updates

.LICENSEURI

.PROJECTURI https://gal.vin/2019/06/16/automated-office-updates

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

#>

<#
    .SYNOPSIS
    Checks for updates to Office Click-to-run source files.
    
    .DESCRIPTION
    A simple script to check for updates to Office click-to-run source files. If updates are available they are downloaded.
    An email notification can be configured to be sent.

    Please note: to send a log file using ssl and an SMTP password you must generate an encrypted
    password file. The password file is unique to both the user and machine.
    
    The command is as follows:

    $creds = Get-Credential
    $creds.Password | ConvertFrom-SecureString | Set-Content c:\foo\ps-script-pwd.txt
   
    .PARAMETER Office
    The location of the Office Deployment Tool. Must be a folder.

    .PARAMETER Config
    The name of the configuration xml file for the Office Deployment Tool. It must be in the same folder of the Office deployment tool.
    
    .PARAMETER Days
    The number of days that you wish to keep old update files for. If you do not configure this option, no old update files will be removed.
    
    .PARAMETER L
    The path to output the log file to.
    The file name will be Office-Update.log.

    .PARAMETER Subject
    The email subject that the email should have. Encapulate with single or double quotes.

    .PARAMETER SendTo
    The e-mail address the log should be sent to.

    .PARAMETER From
    The from address the log should be sent from.

    .PARAMETER Smtp
    The DNS name or IP address of the SMTP server.

    .PARAMETER User
    The user account to connect to the SMTP server.

    .PARAMETER Pwd
    The password for the user account.

    .PARAMETER UseSsl
    Connect to the SMTP server using SSL.

    .EXAMPLE
    Office-Update.ps1 -Office C:\officesrc -Config config.xml -Days 60 -L C:\scripts\logs -Subject 'Server: Office Updates' -SendTo me@contoso.com -From Office-Update@contoso.com -Smtp exch01.contoso.com -User me@contoso.com -Pwd P@ssw0rd -UseSsl

    The above command will run the script, download the Office files to C:\officesrc\Office.
    It will use a configuration file called config.xml in the C:\officesrc folder.
    Any update files older than 60 days will be removed.
    If the download is successful a log file is generated and it can be e-mailed with a custom subject line as a notification that a download occurred.
#>

## Set up command line switches and what variables they map to.
[CmdletBinding()]
Param(
    [parameter(Mandatory=$True)]
    [alias("Office")]
    [ValidateScript({Test-Path $_ -PathType 'Container'})]
    $OfficeSrc,
    [parameter(Mandatory=$True)]
    [alias("Config")]
    $Cfg,
    [alias("Days")]
    $Time,
    [alias("L")]
    [ValidateScript({Test-Path $_ -PathType 'Container'})]
    $LogPath,
    [alias("Subject")]
    $MailSubject,
    [alias("SendTo")]
    $MailTo,
    [alias("From")]
    $MailFrom,
    [alias("Smtp")]
    $SmtpServer,
    [alias("User")]
    $SmtpUser,
    [alias("Pwd")]
    [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
    $SmtpPwd,
    [switch]$UseSsl)

#Run update process.
& $OfficeSrc\setup.exe /download $OfficeSrc\$Cfg

## Location of the office source files.
$UpdateFolder = "$OfficeSrc\Office\Data"

## Check the last write time of the office source files folder if it is greater than the previous day.
$Updated = (Get-ChildItem -Path $UpdateFolder | Where-Object CreationTime -gt (Get-Date).AddDays(-1)).Count

## If the Updated variable returns as not 0...
If ($Updated -ne 0)
{
    ## If logging is configured, start logging.
    If ($LogPath)
    {
        $LogFile = "Office-Update.log"
        $Log = "$LogPath\$LogFile"

        ##Test for the existence of the log file.
        $LogT = Test-Path -Path $Log

        ## If the log file already exists, clear it.
        If ($LogT)
        {
            Clear-Content -Path $Log
        }

        Add-Content -Path $Log -Value "****************************************"
        Add-Content -Path $Log -Value "$(Get-Date -Format G) Log started"
        Add-Content -Path $Log -Value " "
    }

    Write-Host
    Write-Host -Object "Office source files were updated."
    Write-Host -Object "New version is:"

    ## List the update folder contents and the last write time.
    Get-ChildItem -Path $UpdateFolder -Directory | Where-Object CreationTime –gt (Get-Date).AddDays(-1) | Select-Object -ExpandProperty Name

    ## If logging was configured, write to the log.
    If ($LogPath)
    {
        Add-Content -Path $Log -Value "Office source files were updated"
        Add-Content -Path $Log -Value "New version is:"
        Get-ChildItem -Path $UpdateFolder | Where-Object {$_.CreationTime -gt (Get-Date).AddDays(-1)} | Select-Object -Property Name, CreationTime | Out-File -Append $Log -Encoding ASCII
        Add-Content -Path $Log -Value " "
    }

    If ($Null -ne $Time)
    {
        ## If logging was configured, write to the log.
        If ($LogPath)
        {
            Add-Content -Path $Log -Value "Old Office source files were removed:"
            Get-ChildItem -Path $UpdateFolder | Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-$Time)} | Select-Object -Property Name, LastWriteTime | Out-File -Append $Log -Encoding ASCII
            Add-Content -Path $Log -Value " "
        }

        Write-Host
        Write-Host "Old Office source files were removed:"
        Get-ChildItem -Path $UpdateFolder | Where-Object LastWriteTime –lt (Get-Date).AddDays(-$Time)

        ## If configured, remove the old files.
        Get-ChildItem $UpdateFolder | Where-Object {$_.LastWriteTime –lt (Get-Date).AddDays(-$Time)} | Remove-Item -Recurse
    }

    ## If logging was configured stop the log.
    If ($LogPath)
    {
        Add-Content -Path $Log -Value " "
        Add-Content -Path $Log -Value "$(Get-Date -Format G) Log finished"
        Add-Content -Path $Log -Value "****************************************"

        ## If email was configured, set the variables for the email subject and body.
        If ($SmtpServer)
        {
            # If no subject is set, use the string below.
            If ($Null -eq $MailSubject)
            {
                $MailSubject = "Office Update"
            }

            $MailBody = Get-Content -Path $Log | Out-String

            ## If an email password was configured, create a variable with the username and password.
            If ($SmtpPwd)
            {
                $SmtpPwdEncrypt = Get-Content $SmtpPwd | ConvertTo-SecureString
                $SmtpCreds = New-Object System.Management.Automation.PSCredential -ArgumentList ($SmtpUser, $SmtpPwdEncrypt)

                ## If ssl was configured, send the email with ssl.
                If ($UseSsl)
                {
                    Send-MailMessage -To $MailTo -From $MailFrom -Subject $MailSubject -Body $MailBody -SmtpServer $SmtpServer -UseSsl -Credential $SmtpCreds
                }

                ## If ssl wasn't configured, send the email without ssl.
                Else
                {
                    Send-MailMessage -To $MailTo -From $MailFrom -Subject $MailSubject -Body $MailBody -SmtpServer $SmtpServer -Credential $SmtpCreds
                }
            }

            ## If an email username and password were not configured, send the email without authentication.
            Else
            {
                Send-MailMessage -To $MailTo -From $MailFrom -Subject $MailSubject -Body $MailBody -SmtpServer $SmtpServer
            }
        }
    }
}

## End