SDL Trados Studio Automation Kit (STraSAK)
=============

## Introduction
In localization engineer's daily work, many projects use more-or-less fixed process and structure:
- master TMs are stored in a fixed location
- individual handoffs use fixed (sub)folders structure *(location of the prepared source files is known, location where the Trados Studio project should be created is known, etc...)*
- Trados Studio projects naming follows certain pattern, *e.g. {client_name}\_{current_date}_{handoff_name}*
- source language is (almost) always the same... target languages are also often the same...
- and so on...

So... it should not be necessary to enter **manually** all these parameters required for Trados Studio operations over and over again... some **script** should be able to use all this known data and do everything automatically.

And that's the purpose of this automation kit – to allow engineers to automate Trados-related operation using simple scriptable commands:
Command                 | Description
------------------------|------------
**New-Project**         | Create new project – optionally based on project template or another project – in specified location, using specified source and target languages and TMs from specified location. Get source files from specified location and automatically convert them to translatable format and copy them to target languages. Optionally also pre-translate and analyze the files, saving results to Trados 2007-formatted log.
**Export-Package**      | Create translation packages from specified project, using specified package options, and save  them to specified location
**Import-Package**      | Import return packages from specified location in a specified project
**ConvertTo-TradosLog** | Convert Studio XML-based report to Trados 2007-formatted log. _Note:_ _This command_ _**operates on individual files only**__, not on entire directories._
**Export-TargetFiles**  | Export target files from specified project to specified location
==Work in progress==    | *(Commands work, but some optimizations are still in progress)*
**New-FileBasedTM**     | Create new translation memory in specified location, using specified options
**Import-TMX**          | Import content from TMX file in a specified TM
**Update-MainTM**       | Update main translation memories of specified project


## Technical info
Automation is based on SDL PowerShell Toolkit (https://github.com/sdl/Sdl-studio-powershell-toolkit) with own extensive enhancements and customizations.
The kit consists of two parts:

 - **PowerShell modules** containing the actual automation
 - **Windows batch wrapper/launcher script** for easy invocation of the PowerShell functions from command line

## Installation and setup
0. **Pre-requisite**: Windows PowerShell 4.0 or newer (https://www.microsoft.com/en-us/download/details.aspx?id=40855)
_This may be required only for Windows 7 and 8._
_**Windows 8.1 has PowerShell 4.0 already built-in, Windows 10 has PowerShell 5.0 already built-in.**_

1.	**Create `WindowsPowerShell` subfolder** in your `Documents` folder  
(i.e. the result will be `C:\Users\<YourProfile>\Documents\WindowsPowerShell`)  
**NOTE**: If you moved your Documents folder to another location, create the subfolder in that location.
2.	**Copy the entire `Modules` folder** (including the folder structure) into the created `WindowsPowerShell` folder.
3.	**Put the `TS2015.cmd` wrapper script** to any preferred location and **add the location to your PATH environment variable**, so that you can run the script without specifying its full path.  
See https://www.java.com/en/download/help/path.xml for more information about PATH variable and how to edit its content in different operating systems.  
_(Optionally you can put the script to a location which is already listed in the PATH variable... but that may be uncomfortable, depeding on particular system setup, etc.)_

That's all... the kit is now ready for use!

## Usage
### In Windows batch script
Call the wrapper script with desired action command and its parameters as command line arguments:  
`call TS2015 New-Project -Name "My project" -Location "D:\My project" ...`

If the wrapper script's location is not listed in PATH environment variable, you need to use full path to script:  
`call "C:\My scripts\TS2015.cmd" New-Project -Name "My project" -Location "D:\My project" ...`

### In Powershell script
Call the desired PowerShell function with corresponding parameters directly from your PowerShell script:  
`New-Project -Name "My project" -Location "D:\Projects\My project" ...`

### In other scripting languages  
Use the language's appropriate method to call either the batch wrapper, or the PowerShell function.

## Usage Examples
Here is a few Windows batch scripts as examples of automation implementation.
Target languages list for "CreateProject", "ExportPackages" and "ExportFiles" scripts is passed as script parameter, e.g.:  
`03_CreateProject.cmd "de-DE fr-FR it-IT"`
*(language codes can be separated by space, comma or pipe)*

==03_CreateProject.cmd==:
```
@echo off
set TARGETLANGUAGES=%~1
for %%D in ("%CD%") do set "PROJECTNAME=%%~nxD"
call TS2015 New-Project ^
     -Name "%PROJECTNAME%" ^
     -SourceLocation "02_Prep" ^
     -ProjectLocation "03_Studio" ^
     -LogLocation "04_ForTrans" ^
     -TargetLanguages "%TARGETLANGUAGES%" ^
     -TMLocation "X:\Projects\My Project\_TMs" ^
     -ProjectTemplate "X:\Projects\My Project\_Template\MyProject.sdltpl" ^
     -Pretranslate -Analyze
```
==04_ExportPackages.cmd==:
```
@echo off
set TARGETLANGUAGES=%~1
call TS2015 Export-Package ^
     -ProjectLocation "03_Studio" ^
     -PackageLocation "04_ForTrans" ^
     -TargetLanguages "%TARGETLANGUAGES%" ^
     -IncludeMainTMs -IncludeTermbases
```
==05_ImportPackages.cmd==:
```
@echo off
call TS2015 Import-Package ^
     -ProjectLocation "03_Studio" ^
     -PackageLocation "05_FromTrans"
```
==06_ExportFiles.cmd==:
```
@echo off
set TARGETLANGUAGES=%~1
call TS2015 Export-TargetFiles ^
     -ProjectLocation "03_Studio" ^
     -ExportLocation "06_Post" ^
     -TargetLanguages "%TARGETLANGUAGES%"
```

## Known issues
### "log4net:ERROR: XmlConfigurator..." message displayed each time automation is started
This is Trados Studio API bug. It's just a cosmetic issue and does not influence automation functionality.

### No progress is displayed during Analysis task
This seems to be a Trados Studio API bug – activating analysis progress display causes analysis task to completely fail.
Currently the only known 'workaround' is to not display progress during analysis.

### Out Of Memory error during analysis or package creation  
This seems to be caused by some weird memory leak in Studio API if a huge TM is used for analysis or for creating Project TM (either during project creation, or when a "Create new TM" package creation option is used), which may cause Out Of Memory exception.
As a workaround you can either use smaller/less TMs, or create packages or analyze  
individual languages separately instead of all at once... or both.

