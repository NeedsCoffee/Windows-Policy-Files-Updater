[CmdletBinding()]
param (
    [Parameter()][string]$MSI_URL = '',
    [Parameter(DontShow)][int]$AcquireModuleInWrapper = 0,
    [Parameter(DontShow)][switch]$Wrapped,
    [Parameter(DontShow)][string]$Working_Dir
)

$MSI_URL_default = 'https://download.microsoft.com/download/9/5/b/95be347e-c49e-4ede-a205-467c85eb1674/Administrative%20Templates%20(.admx)%20for%20Windows%2011%20Sep%202024%20Update.msi'
$startingPath = Get-Location

function TestForModule {
    if(-not (Get-Module -ListAvailable -Name NtObjectManager)){
        $script:AcquireModuleInWrapper = 1
    } else {
        Write-Output "NtObjectManager is available, importing it"
        Import-Module -Name NtObjectManager -Scope Local -Force
    }
}
function LaunchFromWrapper {
    if([string]::IsNullOrWhiteSpace($script:MSI_URL)){
        $script:MSI_URL = $script:MSI_URL_default
    }
    Write-Output "Launching wrapped admin-console"
    Start-Process PowerShell.exe -WorkingDirectory $script:working_dir -ArgumentList "-NoProfile -File `"$PSCommandPath`" -Wrapped -MSI_URL:`"$script:MSI_URL`" -AcquireModuleInWrapper:$script:AcquireModuleInWrapper -Working_Dir:`"$script:working_dir`"" -Verb RunAs -Wait
}
function PrepareEnvironment {
    Write-Output 'Preparing script environment'
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
    $script:working_dir = Set-Location (New-Item -ItemType Directory -Path "$($env:TEMP)\$(Get-Date -Format FileDateTime)") -PassThru
    Write-Output "Created temporary folder [$script:working_dir]"
}

function AcquireModule {
    Write-Output 'Acquring module NtObjectManager'
    Import-Module "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement" -Force
    Import-Module "$env:ProgramFiles\WindowsPowerShell\Modules\PowerShellGet" -Force
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ForceBootstrap -EA:Si | Out-Null
    Import-PackageProvider -Name NuGet -RequiredVersion 2.8.5.201 -Force -EA:Si
    Register-PSRepository -Default -InstallationPolicy Trusted -EA:Si
    Save-Module -Name NtObjectManager -Repository PSGallery -Path "$script:working_dir\module" -Force
    Import-Module "$script:working_dir\module\NtObjectManager" -Force
}

function PrepareTrustedInstaller {
    Write-Output 'Starting TrustedInstaller service'
    Start-Service -Name TrustedInstaller -Confirm:$false
    Start-Sleep -Seconds 3
    $script:TrustedInstaller = Get-NtProcess -ServiceName TrustedInstaller
}
function MsiActions {
    Write-Output 'Downloading msi file'
    Invoke-WebRequest -UseBasicParsing -OutFile "$script:working_dir\installer.msi" -Uri $script:MSI_URL
    Unblock-File -Path "$script:working_dir\installer.msi" -Confirm:$false
    Write-Output "Extracting msi file to $script:working_dir\msi_extract"
    Start-Process -FilePath "$env:SystemRoot\system32\msiexec.exe" -WorkingDirectory $script:working_dir -NoNewWindow -ArgumentList "/A `"$working_dir\installer.msi`" TARGETDIR=`"$script:working_dir\msi_extract`" /QB" -Wait
    $script:policy_source = Get-ChildItem -Path "$script:working_dir\msi_extract\*.admx" -Recurse -File `
        | Select-Object -First 1 `
        | Select-Object -ExpandProperty Directory `
        | Select-Object -ExpandProperty FullName
    $script:policy_store = "$env:SystemRoot\PolicyDefinitions\"
    Write-Output "Policy store: $script:policy_store"
}

function CleanupEnvironment {
    Write-Output 'Cleaning-up environment'
    $script:startingPath | Set-Location
    $script:working_dir | Remove-Item -Recurse -Force -Confirm:$False
    Write-Output 'Done.'
}

function CopyPolicyFiles {
    Write-Output "Will copy policy files to `"$script:policy_store`" from"
    Write-Output "`"$script:policy_source`""
    Write-Output 'Invoking xcopy as TrustedInstaller'
    $xcopyProcess = New-Win32Process -CurrentDirectory $script:policy_source -ApplicationName $env:ComSpec -CommandLine "/C TITLE Copying Policy Files... & XCOPY `"$script:policy_source\*`" `"$script:policy_store`" /S /V /F /G /H /R /Y & TIMEOUT /T 10" -CreationFlags NewConsole -ParentProcess $script:TrustedInstaller -Wait
}

if(-not $Wrapped){
    TestForModule
    PrepareEnvironment
    LaunchFromWrapper
}

if($Wrapped){
    if($AcquireModuleInWrapper){ AcquireModule }
    PrepareTrustedInstaller
    MsiActions
    CopyPolicyFiles
}

if(-not $Wrapped){ CleanupEnvironment }
