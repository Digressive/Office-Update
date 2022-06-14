# Office Update Utility

Microsoft Office Update Manager

For full change log and more information, [visit my site.](https://gal.vin/utils/office-update-utility/)

Office Update Utility is available from:

* [GitHub](https://github.com/Digressive/Office-Update)
* [The Microsoft PowerShell Gallery](https://www.powershellgallery.com/packages/Office-Update)

Please consider supporting my work:

* Sign up using [Patreon](https://www.patreon.com/mikegalvin).
* Support with a one-time donation using [PayPal](https://www.paypal.me/digressive).

If you’d like to contact me, please leave a comment, send me a [tweet or DM](https://twitter.com/mikegalvin_), or you can join my [Discord server](https://discord.gg/5ZsnJ5k).

-Mike

## Features and Requirements

* This utility will check for and download update files for Microsoft Office.
* It can also remove old update files.
* It can create and e-mail a log file when there are updates.
* This utility requires the Office Deployment Tool [a free download available here.](https://www.microsoft.com/en-us/download/details.aspx?id=49117)
* This utility requires at least PowerShell 5.0.
* This utility has been tested on Windows 11, Windows 10, Windows Server 2022, Windows Server 2019 and Windows Server 2016.

## Folder Structure

This utility requires a specific folder structure in order to operate, it expects the Office Deployment Tool and the configuration xml file to be in the same folder. Additionally, the source path of the Office installation files in the configuration xml file should be set to the same location. For example:

* Office Deployment Tool location: ```\\server\share\Office-365-x64\setup.exe```
* Configuration xml file location: ```\\server\share\Office-365-x64\config-2019-x64.xml```
* Source path in the configuration xml file: ```\\server\share\Office-365-x64```

This configuration will result in the Office update files being downloaded and stored in: ```\\server\share\Office-2019-x64\Office\Data```

## Generating A Password File

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
