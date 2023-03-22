# Office Update Utility

## Microsoft Office Update Manager

For full change log and more information, [visit my site.](https://gal.vin/utils/office-update-utility/)

Office Update Utility is available from:

* [GitHub](https://github.com/Digressive/Office-Update)
* [The Microsoft PowerShell Gallery](https://www.powershellgallery.com/packages/Office-Update)

Please consider supporting my work:

* Sign up using [Patreon](https://www.patreon.com/mikegalvin).
* Support with a one-time donation using [PayPal](https://www.paypal.me/digressive).

Please report issues on Github via the issues tab.

Thanks
-Mike

## Features and Requirements

* This utility will check for and download update files for Microsoft Office.
* It can also remove old update files.
* This utility requires the Office Deployment Tool [a free download available here.](https://www.microsoft.com/en-us/download/details.aspx?id=49117)
* This utility requires at least PowerShell 5.0.
* This utility has been tested on Windows 11, Windows 10, Windows Server 2022, Windows Server 2019 and Windows Server 2016.
* The update log can be sent via email and/or webhook.

## Folder Structure

This utility requires a specific folder structure in order to operate, it expects the Office Deployment Tool and the configuration xml file to be in the same folder. Additionally, the source path of the Office installation files in the configuration xml file should be set to the same location. For example:

* Office Deployment Tool location: ```\\server\share\Office-365-x64\setup.exe```
* Configuration xml file location: ```\\server\share\Office-365-x64\config-2019-x64.xml```
* Source path in the configuration xml file: ```\\server\share\Office-365-x64```

This configuration will result in the Office update files being downloaded and stored in: ```\\server\share\Office-2019-x64\Office\Data```

## Generating A Password File For SMTP Authentication

The password used for SMTP server authentication must be in an encrypted text file. To generate the password file, run the following command in PowerShell on the computer and logged in with the user that will be running the utility. When you run the command, you will be prompted for a username and password. Enter the username and password you want to use to authenticate to your SMTP server.

Please note: This is only required if you need to authenticate to the SMTP server when send the log via e-mail.

``` powershell
$creds = Get-Credential
$creds.Password | ConvertFrom-SecureString | Set-Content c:\scripts\ps-script-pwd.txt
```

After running the commands, you will have a text file containing the encrypted password. When configuring the -Pwd switch enter the path and file name of this file.

## Configuration

Here’s a list of all the command line switches and example configurations.

| Command Line Switch | Description | Example |
| ------------------- | ----------- | ------- |
| -Office | The folder containing the Office Deployment Tool (ODT). | [path\] |
| -Config | The name of the configuration xml file for the Office ODT. It must be located in the same folder as the ODT. | [file name.xml] |
| -Days | The number of days that you wish to keep old update files for. If you do not configure this option, no old files will be removed. | [number] |
| -L | The path to output the log file to. | [path\] |
| -LogRotate | Remove logs produced by the utility older than X days | [number] |
| -NoBanner | Use this option to hide the ASCII art title in the console. | N/A |
| -Help | Display usage information. No arguments also displays help. | N/A |
| -Webhook | The txt file containing the URI for a webhook to send the log file to. | [path\]webhook.txt |
| -Subject | Specify a subject line. If you leave this blank the default subject will be used | "'[Server: Notification]'" |
| -SendTo | The e-mail address the log should be sent to. For multiple address, separate with a comma. | [example@contoso.com] |
| -From | The e-mail address the log should be sent from. | [example@contoso.com] |
| -Smtp | The DNS name or IP address of the SMTP server. | [smtp server address] |
| -Port | The Port that should be used for the SMTP server. If none is specified then the default of 25 will be used. | [port number] |
| -User | The user account to authenticate to the SMTP server. | [example@contoso.com] |
| -Pwd | The txt file containing the encrypted password for SMTP authentication. | [path\]ps-script-pwd.txt |
| -UseSsl | Configures the utility to connect to the SMTP server using SSL. | N/A |

## Example

``` txt
[path\]Office-Update.ps1 -Office [path\] -Config [file name.xml] -Days [number]
```

This will update the office installation files in the specified directory, and delete update files older than X days

## Change Log

### 2023-02-07: Version 23.02.07

* Added script update checker - shows if an update is available in the log and console.
* Added webhook option to send log file to.
* Removed SMTP authentication details from the 'Config' report. Now it just shows as 'configured' if SMTP user is configured. To be clear: no passwords were ever shown or stored in plain text.

### 2022-07-31: Version 22.07.30

* Changed how the removal of old files works. Old versions will be removed regardless, -Days option has been removed.

### 2022-06-22: Version 22.06.22

* Fixed an issue where If -L [path\] not configured then a non fatal error would occur as no log path was specified for the log to be output to.

### 2022-06-14: Version 22.05.25

* Added new feature: log can now be emailed to multiple addresses.
* Added checks and balances to help with configuration as I'm very aware that the initial configuration can be troublesome. Running the utility manually is a lot more friendly and step-by-step now.
* Added -Help to give usage instructions in the terminal. Running the script with no options will also trigger the -help switch.
* Cleaned user entered paths so that trailing slashes no longer break things or have otherwise unintended results.
* Added -LogRotate [days] to removed old logs created by the utility.
* Streamlined config report so non configured options are not shown.
* Added donation link to the ASCII banner.
* Cleaned up code, removed unneeded log noise.

### 2022-05-25: Version 22.05.25

* Added -Help to give usage instructions in the terminal. Also running the script with no options will also trigger the -help switch.
* Streamlined config report so non configured options are not shown.
* Added a -LogRotate option to delete logs older than X number of days.

### 2021-12-06: Version 21.12.06

* Fixed problem with Hostname not displaying.

### 2021-12-04: Version 21.12.04

* Configured logs path now is created, if it does not exist.
* Added OS version info.
* Added an option to specify the Port for SMTP communication.

### 2020-03-03: Version 2020.03.01 'Crosshair'

New features:

* Refactored code.
* Fully backwards compatible.
* Added ASCII banner art when run in the console.
* Added option to disable the ASCII banner art.
* Config report matches design of Image Factory Utility.

### 2019-09-04 v1.1

* Added custom subject line for e-mail.

### 2019-06-16 v1.0

* Initial release.
