SDL Trados Studio Automation Kit (STraSAK)
==========================================

## Introduction
SDL Trados Studio Automation Kit allows to automate repetitive tasks like project creation, translation packages creation, return packages import, files export, etc.

In localization engineer's daily work, projects often use more-or-less fixed process and structure:
- master TMs are stored in a fixed location
- individual handoffs use fixed (sub)folders structure *(location of the prepared source files is known, location where the Trados Studio project should be created is known, etc...)*
- Trados projects are usually named using a fixed scheme *(e.g. after the handoff name)*
- source language is (almost) always the same... target languages are also often the same...
- and so on...

This means that all these _parameters_ required for Trados Studio operations should not need to be manually entered over and over again, requiring gazzillion of clicks in the Studio GUI.
A script should be able to use all this known data and do everything automatically...

And that's exactly what this automation kit is for – to allow engineers to automate Trados-related operation using simple scripts.

## Technical info
Automation is based on SDL PowerShell Toolkit (https://github.com/sdl/Sdl-studio-powershell-toolkit) with own customizations and enhancements. The kit consists of two parts:

 - **PowerShell modules** containing the actual automation
 - **Windows batch wrapper/launcher script** for easy invocation of the PowerShell functions from command line

## Installation and setup
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
