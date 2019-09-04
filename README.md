# Automated-Office-Updates

PowerShell script to automate the downloading and maintenance of click-to-run Office 365/2019 update files.

Automated Office Updates can also be downloaded from:

* [The Microsoft TechNet Gallery](https://gallery.technet.microsoft.com/Automated-Office-Updates-4fef21d3?redir=0)
* [The PowerShell Gallery](https://www.powershellgallery.com/packages/Office-Update/1.0)
* For full instructions and documentation, [visit my blog post](https://gal.vin/2019/06/16/automated-office-updates/)

-Mike

Tweet me if you have questions: [@mikegalvin_](https://twitter.com/mikegalvin_)

## Features and Requirements

* This script will check for and download update files for click-to-run Office installs such as Office 365 and Office 2019.
* The script can also remove old update files to prevent bloat.
* The script can optionally create a log file and e-mail the log file to an address of your choice.
* The script has been tested on Windows Server 2016 and Windows Server 2019, updating Office 2019 volume licensed installations on Windows 10 1809 and 1903.

### Generating A Password File

The password used for SMTP server authentication must be in an encrypted text file. To generate the password file, run the following command in PowerShell, on the computer that is going to run the script and logged in with the user that will be running the script. When you run the command you will be prompted for a username and password. Enter the username and password you want to use to authenticate to your SMTP server.

Please note: This is only required if you need to authenticate to the SMTP server when send the log via e-mail.

``` powershell
$creds = Get-Credential
$creds.Password | ConvertFrom-SecureString | Set-Content c:\scripts\ps-script-pwd.txt
```

After running the commands, you will have a text file containing the encrypted password. When configuring the -Pwd switch enter the path and file name of this file.

### Configuration

Here’s a list of all the command line switches and example configurations.

``` txt
-Office
```

The folder containing the Office Deployment Tool (ODT).

``` txt
-Config
```

The name of the configuration xml file for the Office ODT. It must be located in the same folder as the ODT.

``` txt
-Days
```

The number of days that you wish to keep old update files for. If you do not configure this option, no old files will be removed.

``` txt
-L
```

The path to output the log file to. The file name will be Office-Update.log.

``` txt
-Subject
```

The email subject that the email should have. Encapulate with single or double quotes.

``` txt
-SendTo
```

The e-mail address the log should be sent to.

``` txt
-From
```

The from address the log should be sent from.

``` txt
-Smtp
```

The DNS name or IP address of the SMTP server.

``` txt
-User
```

The user account to connect to the SMTP server.

``` txt
-Pwd
```

The password for the user account.

``` txt
-UseSsl
```

Connect to the SMTP server using SSL.

### Example

``` txt
Office-Update.ps1 -Office C:\officesrc -Config config.xml -Days 60 -L C:\scripts\logs -Subject 'Server: Office Update' -SendTo me@contoso.com -From Office-Update@contoso.com -Smtp exch01.contoso.com -User me@contoso.com -Pwd P@ssw0rd -UseSsl
```

The above command will run the script, download the Office files to C:\officesrc\Office. It will use a configuration file called config.xml in the C:\officesrc folder. Any update files older than 60 days will be removed. If the download is successful a log file is generated and it can be e-mailed with a custom subject line as a notification that a download occurred.
