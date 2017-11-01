<# : ----------------- begin batch part -----------------
@echo off

:: get Documents folder location from registry
rem for /f "tokens=2*" %%A in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v "Personal" 2^>nul ^| find "REG_"') do call set "DocumentsFolder=%%B"
for /f "tokens=2*" %%A in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v "Personal" 2^>nul') do call set "DocumentsFolder=%%B"
:: add Modules location to PSModulePath variable
set "PSModulePath=%DocumentsFolder%\WindowsPowerShell\Modules\;%PSModulePath%"

:: PowerShell location
set POWERSHELL=%windir%\system32\WindowsPowerShell\v1.0\powershell.exe
:: use 32-bit version on 64-bit systems!
if exist "%windir%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" (set POWERSHELL=%windir%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe)

:: from http://www.dostips.com/forum/viewtopic.php?f=3&t=5526&start=15#p45502
:: invoke embedded PowerShell code + code specified on command line
setlocal enabledelayedexpansion
set _args=%*
rem this is to prevent PS errors if launched with empty command line
if not defined _args set "_args= "
set _args=!_args:'=''!
rem set _args=!_args:"=''!
set _args=!_args:"="""!
type "%~f0" | %POWERSHELL% -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "%POWERSHELL% -NoLogo -NoProfile -ExecutionPolicy Bypass -Command ([ScriptBlock]::Create([Console]::In.ReadToEnd()+';!_args!'))"
rem type "%~f0" | %POWERSHELL% -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "%POWERSHELL% -NoLogo -NoProfile -ExecutionPolicy Bypass -Command ('Set-Location -LiteralPath ''%CD%'';'+[ScriptBlock]::Create([Console]::In.ReadToEnd()+';!_args!'))"
endlocal

exit /b 0
----------------- end batch part ----------------- #>

[cultureinfo]::DefaultThreadCurrentCulture = 'en-US'

# import SDL Powershell Toolkit modules
$StudioVersion = "Studio5"
$Modules = @("TMHelper","ProjectHelper","GetGuids","PackageHelper")
ForEach ($Module in $Modules) {Import-Module -Name $Module -ArgumentList $StudioVersion}