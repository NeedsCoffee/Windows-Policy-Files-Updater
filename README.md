# Windows-Policy-Files-Updater
> [!NOTE]
> A tool to download Windows Administrative Templates and update the PolicyDefinitions folder with the ADMX/ADML files from the template package.

Microsoft provides no way to update ADMX files in the c:\Windows\PolicyDefinitions folder and in fact makes it extraordinarily difficult for anyone, even a skilled administrator, to replace these files. However if you can act as the TrustedInstaller service you can actually do this quite easily it turns out.

This PowerShell script solves this problem simply and quickly by downloading the msi, extracting it, and then copying the files using the necessary privileges required to complete the task.

The current latest templates are the "Administrative Templates (.admx) for Windows 11 2024 Update (24H2)" found here: [https://www.microsoft.com/en-us/download/details.aspx?id=105667](https://www.microsoft.com/en-gb/download/details.aspx?id=106254)

## How to run the script
> Download **Update-ADMX-Policy-Files.ps1** from the repo, then right-click the file and choose "Run with PowerShell"

- The script will launch and request admin privilages
- It will then create a temporary folder
- Temporarily installs the NtObjectManager module (used to access the TrustedInstaller service)
- Downloads the administrative templates msi it knows about
- Extracts the msi file using MSIEXEC /A
- Runs an XCOPY command as the TrustedInstaller service. This updates the local machine's C:\Windows\PolicyDefinitions folder with ADMX and ADML files from the extracted MSI
- Finally the script will cleanup after itself by deleting the temporary folder

## Using a different MSI
> [!NOTE]
> Run the script with **`-MSI_URL <url>`** to specify an alternative admin templates package

e.g. override the url with the package for Windows 10 22H2 by doing the following from a PowerShell prompt:

```
.\Update-ADMX-Policy-Files.ps1 -MSI_URL "https://download.microsoft.com/download/c/3/c/c3cd85c0-0785-4cf7-a48e-cdc9b8e20108/Administrative%20Templates%20(.admx)%20for%20Windows%2010%20October%202022%20Update.msi"
```
