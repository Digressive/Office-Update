# Office Update Utility

Office Update Manager

```txt
     ___  __  __ _                            _       _
    /___\/ _|/ _(_) ___ ___   /\ /\ _ __   __| | __ _| |_ ___
   //  // |_| |_| |/ __/ _ \ / / \ \ '_ \ / _  |/ _  | __/ _ \  
  / \_//|  _|  _| | (_|  __/ \ \_/ / |_) | (_| | (_| | ||  __/  
  \___/ |_| |_| |_|\___\___|  \___/| .__/ \__,_|\__,_|\__\___|  
                                   |_|
         _   _ _ _ _
   /\ /\| |_(_) (_) |_ _   _
  / / \ \ __| | | | __| | | |           3 6 5
  \ \_/ / |_| | | | |_| |_| |          2 0 1 9
   \___/ \__|_|_|_|\__|\__, |        Click-to-Run
                       |___/

    Mike Galvin    https://gal.vin    Version 21.12.04
```

For full instructions and documentation, [visit my site.](https://gal.vin/posts/automated-office-updates/)

Please consider supporting my work:

* Sign up using [**Patreon**](https://www.patreon.com/mikegalvin).
* Support with a one-time payment using [**PayPal**](https://www.paypal.me/digressive).

Office Update Utility can also be downloaded from:

* [The Microsoft PowerShell Gallery](https://www.powershellgallery.com/packages/Office-Update)

Join the [Discord](http://discord.gg/5ZsnJ5k) or Tweet me if you have questions: [@mikegalvin_](https://twitter.com/mikegalvin_)

-Mike

## Features and Requirements

* This utility will check for and download update files for Office 365 and Office 2019.
* It can be configured to remove old update files.
* It can be configured to create and e-mail a log file.
* The utility requires the Office Deployment Tool [a free download available here.](https://www.microsoft.com/en-us/download/details.aspx?id=49117)
* The utility requires at least PowerShell 5.0.
* This utility has been tested on Windows 11, Windows 10, Windows Server 2016 and Windows Server 2019 and updating Office 2019 and 365 installations.

### Folder Structure

This utility requires a specific folder structure in order to operate, it expects the Office Deployment Tool and the configuration xml file to be in the same folder. Additionally, the source path of the Office installation files in the configuration xml file should be set to the same location. For example:

* Office Deployment Tool location: ```\\server\share\Office-365-x64\setup.exe```
* Configuration xml file location: ```\\server\share\Office-365-x64\config-2019-x64.xml```
* Source path in the configuration xml file: ```\\server\share\Office-365-x64```

This configuration will result in the Office update files being downloaded and stored in: ```\\server\share\Office-2019-x64\Office\Data```

### Generating A Password File

The password used for SMTP server authentication must be in an encrypted text file. To generate the password file, run the following command in PowerShell on the computer and logged in with the user that will be running the utility. When you run the command, you will be prompted for a username and password. Enter the username and password you want to use to authenticate to your SMTP server.

Please note: This is only required if you need to authenticate to the SMTP server when send the log via e-mail.

``` powershell
$creds = Get-Credential
$creds.Password | ConvertFrom-SecureString | Set-Content c:\scripts\ps-script-pwd.txt
```

After running the commands, you will have a text file containing the encrypted password. When configuring the -Pwd switch enter the path and file name of this file.

### Configuration

Here’s a list of all the command line switches and example configurations.

| Command Line Switch | Description | Example |
| ------------------- | ----------- | ------- |
| -Office | The folder containing the Office Deployment Tool (ODT). | ```\\server\share\office-365-x64\setup.exe``` |
| -Config | The name of the configuration xml file for the Office ODT. It must be located in the same folder as the ODT. | config-365-x64.xml |
| -Days | The number of days that you wish to keep old update files for. If you do not configure this option, no old files will be removed. | 30 |
| -NoBanner | Use this option to hide the ASCII art title in the console. | N/A |
| -L | The path to output the log file to. The file name will be Office-Update_YYYY-MM-dd_HH-mm-ss.log. Do not add a trailing \ backslash. | ```C:\scripts\logs``` |
| -Subject | The subject line for the e-mail log. Encapsulate with single or double quotes. If no subject is specified, the default of "Office Update Utility Log" will be used. | 'Server: Notification' |
| -SendTo | The e-mail address the log should be sent to. | me@contoso.com |
| -Port | The Port that should be used for the SMTP server. If none is specified then the default of 25 will be used. | 587 |
| -From | The e-mail address the log should be sent from. | OffUpdate@contoso.com |
| -Smtp | The DNS name or IP address of the SMTP server. | smtp.live.com OR smtp.office365.com |
| -User | The user account to authenticate to the SMTP server. | example@contoso.com |
| -Pwd | The txt file containing the encrypted password for SMTP authentication. | ```C:\scripts\ps-script-pwd.txt``` |
| -UseSsl | Configures the utility to connect to the SMTP server using SSL. | N/A |

### Example

``` txt
Office-Update.ps1 -Office \\Apps01\Software\Office365 -Config config-365-x64.xml -Days 30 -L C:\scripts\logs -Subject 'Server: Office Update' -SendTo me@contoso.com -From OffUpdate@contoso.com -Smtp smtp.outlook.com -User me@contoso.com -Pwd P@ssw0rd -UseSsl
```

The above command will download any Office updates for the version and channel configured in config-365-x64.xml to the Office files directory ```\\Apps01\Software\Office365```. Any update files older than 30 days will be removed. If the download is successful the log file will be output to ```C:\scripts\logs``` and e-mailed with a custom subject line.
