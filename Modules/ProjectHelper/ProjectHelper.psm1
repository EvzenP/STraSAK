﻿param([String]$StudioVersion = "Studio5")

if ("${Env:ProgramFiles(x86)}") {
	$ProgramFilesDir = "${Env:ProgramFiles(x86)}"
}
else {
	$ProgramFilesDir = "${Env:ProgramFiles}"
}

Add-Type -Path "$ProgramFilesDir\SDL\SDL Trados Studio\$StudioVersion\Sdl.ProjectAutomation.FileBased.dll"
Add-Type -Path "$ProgramFilesDir\SDL\SDL Trados Studio\$StudioVersion\Sdl.ProjectAutomation.Core.dll"
Add-Type -Path "$ProgramFilesDir\SDL\SDL Trados Studio\$StudioVersion\Sdl.ProjectApi.dll"

# Helper for handling PowerShell runspace issues with multithreaded event handlers
# https://stackoverflow.com/questions/53788232/system-threading-timer-kills-the-powershell-console/53789011
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Management.Automation.Runspaces;

public class RunspacedDelegateFactory
{
    public static Delegate NewRunspacedDelegate(Delegate _delegate, Runspace runspace)
    {
        Action setRunspace = () => Runspace.DefaultRunspace = runspace;

        return ConcatActionToDelegate(setRunspace, _delegate);
    }

    private static Expression ExpressionInvoke(Delegate _delegate, params Expression[] arguments)
    {
        var invokeMethod = _delegate.GetType().GetMethod("Invoke");

        return Expression.Call(Expression.Constant(_delegate), invokeMethod, arguments);
    }

    public static Delegate ConcatActionToDelegate(Action a, Delegate d)
    {
        var parameters =
            d.GetType().GetMethod("Invoke").GetParameters()
            .Select(p => Expression.Parameter(p.ParameterType, p.Name))
            .ToArray();

        Expression body = Expression.Block(ExpressionInvoke(a), ExpressionInvoke(d, parameters));

        var lambda = Expression.Lambda(d.GetType(), body, parameters);

        var compiled = lambda.Compile();

        return compiled;
    }
}
'@

$LanguagesSeparator = "\s+|;\s*|,\s*"

$StudioVersionsMap = @{
	Studio2  = "10.0.0.0"
	Studio3  = "11.0.0.0"
	Studio4  = "12.0.0.0"
	Studio5  = "14.0.0.0"
	Studio15 = "15.0.0.0"
	Studio16 = "16.0.0.0"
}

function Get-DefaultProjectTemplate {
	##########################################################################################################
	# Due to API bug basing new projects on "Default.sdltpl" template instead of actual default project template,
	# we need to find the real default template configured in Trados Studio by reading the configuration files
	$StudioVersionAppData = $StudioVersionsMap[$StudioVersion]
	# Get default project template GUID from the user settings file
	$UserSettingsFilePath = "${Env:AppData}\SDL\SDL Trados Studio\$StudioVersionAppData\UserSettings.xml"
	$DefaultProjectTemplateGuid = (Select-Xml -Path $UserSettingsFilePath -XPath "//Setting[@Id='DefaultProjectTemplateGuid']").Node.InnerText
	# Get user's project templates and get the default one using the GUID
	$ProjectTemplates = [Sdl.ProjectApi.ApplicationFactory]::CreateApplication().LocalProjectServers[0].ProjectTemplates
	$DefaultProjectTemplate = ($ProjectTemplates | Where-Object -Property Guid -eq $DefaultProjectTemplateGuid).FilePath
	
	return $DefaultProjectTemplate
}

function New-Project {
<#
.SYNOPSIS
Creates new Trados Studio file based project.
.DESCRIPTION
Creates new Trados Studio file based project in specified location, using specified source and target languages.
Project can be optionally based on specified project template or other reference project.
Project location is created automatically; existing location is emptied before creating the project.
Translation memories (*.sdltm) are searched in specified location according to source and target languages.
Source files are added from specified location recursively including folders.

Following tasks are run automatically after project creation:
- Scan
- Convert to translatable format
- Copy to target languages

Optionally also following tasks can be run:
- PerfectMatch
- Pretranslate
- Analyze
.EXAMPLE
New-Project -Name "Project" -ProjectLocation "D:\Project" -SourceLocation "D:\Sources"

Creates project named "Project" based on default Trados Studio project template in "D:\Project" folder;
source files are taken from "D:\Sources" folder;
source language, target languages and translation memories are taken from the default project template;
only default Scan, Convert and Copy to target languages tasks are run.
.EXAMPLE
New-Project -Name "Project" -ProjectLocation "D:\Project" -SourceLocation "D:\Sources" -SourceLanguage "en-US" -TargetLanguages "fi-FI" -TMLocation "D:\TMs" -Pretranslate -Analyze

Creates project named "Project" based on default Trados Studio project template in "D:\Project" folder, with American English as source language and Finnish as target language; source files are taken from "D:\Sources" folder and translation memories are taken from "D:\TMs" folder; and runs Pretranslate and Analyze as additional tasks after scanning, converting and copying to target languages.
.EXAMPLE
New-Project -Name "Sample Project" -ProjectLocation "D:\Projects\Trados Studio Automation\Sample" -SourceLocation "D:\Projects\Trados Studio Automation\Source files" -SourceLanguage "en-GB" -TargetLanguages "de-DE,ja-JP" -TMLocation "D:\Projects\TMs\Samples" -ProjectTemplate "D:\ProjectTemplates\SampleTemplate.sdltpl" -Analyze

Creates project named "Sample Project" in "D:\Projects\Trados Studio Automation\Sample" folder, based on "D:\ProjectTemplates\SampleTemplate.sdltpl" project template.
Source language is set to British English and target languages are German and Japanese.
Source files are taken from "D:\Projects\Trados Studio Automation\Source files" folder.
Translation memories are taken from "D:\Projects\TMs\Samples" folder.
Analyze task is run after scanning, converting and copying to target languages.
#>

	[CmdletBinding(DefaultParametersetName="ProjectTemplate")]

	param(
		# Project name. Must not contain invalid characters such as \ / : * ? " < > |
		[Parameter (Mandatory = $true)]
		[Alias("ProjectName","PrjName")]
		[String] $Name,

		# Path to directory where the project will be created.
		# Any existing content of the directory will be deleted before creating the project.
		# If the directory does not exist, it will be created.
		[Parameter (Mandatory = $true)]
		[Alias("Location","PrjLoc")]
		[String] $ProjectLocation,

		# Path to directory containing project source files.
		# Complete directory structure present in that directory will be added as project source.
		[Parameter (Mandatory = $true)]
		[Alias("SrcLoc")]
		[String] $SourceLocation,
		
		# Project source language locale code.
		# For locale codes, see https://msdn.microsoft.com/en-us/goglobal/bb896001.aspx
		[Alias("SrcLng")]
		[String] $SourceLanguage,
		
		# Space-, comma- or semicolon-separated list of locale codes of project target languages.
		# For locale codes, see https://msdn.microsoft.com/en-us/goglobal/bb896001.aspx
		[Alias("TrgLng")]
		[String] $TargetLanguages,

		# Path to directory containing Trados Studio translation memories for project language pairs.
		# Directory will be searched for all TMs with language pairs defined for the project and found TMs will be assigned to the project languages.
		# Additional TMs defined in project template or reference project will be retained (unless "OverrideTM" parameter is specified).
		# Note: directory is NOT searched recursively!
		[Alias("TMLoc")]
		[String] $TMLocation,

		# Path to directory where Trados 2007-formatted analysis logs will be created.
		# Any existing logs will be overwritten.
		[Alias("LogLoc")]
		[String] $LogLocation,

		# Path to project template (*.sdltpl) on which the created project will be based.
		# If this parameter is not specified, default project template set in Trados Studio will be used.
		[Parameter (ParameterSetName = "ProjectTemplate")]
		[Alias("PrjTpl")]
		[String] $ProjectTemplate = $(Get-DefaultProjectTemplate),

		# Path to project file (*.sdlproj) on which the created project will be based.
		[Parameter (ParameterSetName = "ProjectReference")]
		[Alias("PrjRef")]
		[String] $ProjectReference,

		# Ignore TMs defined in project template or reference project and use only TMs specified by "TMLocation" parameter.
		# If this parameter is not present, TMs specified by "TMLocation" parameter will be added to TMs defined in project template.
		[Alias("OvrTM")]
		[Switch] $OverrideTM,

		# Path to directory containing bilingual files to be used for applying PerfectMatch.
		# Directory must contain subdirectories for individual languages and bilingual files in these language subdirectories must be structured exactly the same way as files in project.
		[Alias("PM")]
		[String] $PerfectMatch,

		# Run pre-translation task for target languages during project creation.
		[Alias("PreTra","Translate")]
		[Switch] $Pretranslate,

		# Run analysis task for target languages during project creation.
		[Alias("Analyse")]
		[Switch] $Analyze,

		# Run all batch tasks for each target language separately, as opposed to default Studio behavior which runs batch tasks for all languages at once.
		[Alias("PerLng","PerLang")]
		[Switch] $PerLanguage,

		# Save analysis reports in Excel XLSX format.
		# Reports will be saved to the same location as the Trados 2007 logs.
		[Alias("Excel")]
		[Switch] $ExcelLog
	)

	# If project location does not exist, create it... if it does exist, empty it
	if (!(Test-Path -LiteralPath $ProjectLocation)) {
		New-Item $ProjectLocation -Force -ItemType Directory | Out-Null
	}
	else {
		Get-ChildItem -LiteralPath $ProjectLocation * | Remove-Item -Force -Recurse
	}

	$ProjectLocation = (Resolve-Path -LiteralPath $ProjectLocation).ProviderPath

	# Create file based project
	Write-Host "`nCreating new project..." -ForegroundColor White

	# Get project creation reference, depending on provided parameters
	switch ($PsCmdlet.ParameterSetName) {
		"ProjectTemplate" {
			$ProjectTemplate = (Resolve-Path -LiteralPath $ProjectTemplate).ProviderPath
			$ProjectCreationReference = New-Object Sdl.ProjectAutomation.Core.ProjectTemplateReference $ProjectTemplate
			break
		}
		"ProjectReference" {
			$ProjectReference = (Resolve-Path -LiteralPath $ProjectReference).ProviderPath
			$ProjectCreationReference =  New-Object Sdl.ProjectAutomation.Core.ProjectReference $ProjectReference
			break
		}
	}

	# Create project info required for project creation
	$ProjectInfo = New-Object Sdl.ProjectAutomation.Core.ProjectInfo
	$ProjectInfo.Name = $Name
	$ProjectInfo.LocalProjectFolder = $ProjectLocation

	# If source/target languages were specified, use it as project languages
	# otherwise use only those (eventually) specified in project template/reference
	if ($SourceLanguage -ne $null -and $SourceLanguage -ne "") {
		$ProjectInfo.SourceLanguage = Get-Language $SourceLanguage
	}
	if ($TargetLanguages -ne $null -and $TargetLanguages -ne "") {
		# Parse target languages into array
		$TargetLanguagesList = $TargetLanguages -Split $LanguagesSeparator
		$ProjectInfo.TargetLanguages = Get-Languages $TargetLanguagesList
	}

	# If PerfectMatch parameter was specified, verify the specified path
	if ($PerfectMatch -ne $null -and $PerfectMatch -ne "") {
		$BilingualsPath = (Resolve-Path -LiteralPath $PerfectMatch).ProviderPath
	}

	# Crete new project using creation reference and constructed project info
	$Project = New-Object Sdl.ProjectAutomation.FileBased.FileBasedProject ($ProjectInfo, $ProjectCreationReference)

	# Get project languages
	$TargetLanguagesList = @($Project.GetProjectInfo().TargetLanguages.IsoAbbreviation)
	$SourceLanguage = $Project.GetProjectInfo().SourceLanguage.IsoAbbreviation

	# If TMs location was specified, assign TMs to project languages
	# Otherwise use only TMs (eventually) specified in project template/reference
	if ($TMLocation -ne $null -and $TMLocation -ne "") {
		Write-Host "`nAssigning TMs to project..." -ForegroundColor White

		# Loop through all TMs present in TMs location
		$TMPaths = Get-ChildItem -LiteralPath $TMLocation *.sdltm | ForEach {$_.FullName}
		ForEach($TMPath in $TMPaths) {
			# Get TM language pair
			$TMSourceLanguage = Get-TMSourceLanguage -Path $TMPath | ForEach {Get-Language $_.Name}
			$TMTargetLanguage = Get-TMTargetLanguage -Path $TMPath | ForEach {Get-Language $_.Name}

			# If TM languages are not one of the project lang pairs, skip to next TM
			if ($tmSourceLanguage -ne $SourceLanguage -or $tmTargetLanguage -notin $TargetLanguagesList) {continue}

			# Create new TranslationProviderCascadeEntry entry object for currently processed TM
			$TMentry = New-Object Sdl.ProjectAutomation.Core.TranslationProviderCascadeEntry ($TMPath, $true, $true, $true)

			# Get existing translation provider configuration which can be defined in project template or reference project
			[Sdl.ProjectAutomation.Core.TranslationProviderConfiguration] $TMConfig = $Project.GetTranslationProviderConfiguration($TMTargetLanguage)

			# If OverrideTM parameter was specified, remove all existing TMs from translation provider configuration
			if ($OverrideTM) {$TMConfig.Entries.Clear()}

			# Get list of TM URIs from existing translation provider configuration
			$TMUris = $TMConfig.Entries | ForEach {$_.MainTranslationProvider.Uri}

			# If the TM is not in the existing TMs list, add it
			if ($TMentry.MainTranslationProvider.Uri -notin $TMUris) {
				$TMConfig.Entries.Add($TMentry)
				$TMConfig.OverrideParent = $true
				$Project.UpdateTranslationProviderConfiguration($TMTargetLanguage, $TMConfig)
				Write-Host "$TMPath added to project"
			}
			else {
				Write-Host "$TMPath already in project"
			}
		}
	}
	
	# task progress event handler
	$OnTaskProgress = ${function:Write-TaskProgress} -as [System.EventHandler[Sdl.ProjectAutomation.Core.TaskStatusEventArgs]]
	$OnTaskProgressDelegate = [RunspacedDelegateFactory]::NewRunspacedDelegate($OnTaskProgress, [Runspace]::DefaultRunspace)
	
	# task messages event handler
	$OnTaskMessage = ${function:Write-TaskMessage} -as [System.EventHandler[Sdl.ProjectAutomation.Core.TaskMessageEventArgs]]
	$OnTaskMessageDelegate = [RunspacedDelegateFactory]::NewRunspacedDelegate($OnTaskMessage, [Runspace]::DefaultRunspace)
	
	# scriptblock for running automatic tasks (see below)
	$ProcessAutomaticTasks = {
		
		param($LangsToProcess)
		
		# get IDs of target files to be processed
		[System.Guid[]] $TargetFilesGuids = @()
		ForEach ($TargetLanguage in $LangsToProcess) {
			$TargetFiles = $Project.GetTargetLanguageFiles($TargetLanguage)
			$TargetFilesGuids += Get-Guids $TargetFiles
		}
		
		# run (and then validate) the task sequence
		$Tasks = $Project.RunAutomaticTasks($TargetFilesGuids, $TaskSequence, $OnTaskProgressDelegate, $OnTaskMessageDelegate)
		Validate-TaskSequence $Tasks
		
		# save analysis logs
		if ($Analyze -and $LogLocation) {
			if (!($PerLanguage)) {Write-Host "`n" -NoNewLine}
			Write-Host "Saving logs..." -ForegroundColor White
			ForEach ($T in $Tasks.SubTasks) {
				ForEach ($R in $T.Reports) {
					if ($R.TaskTemplateId -eq [string] [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::AnalyzeFiles) {
						# Save report to temporary file
						$TempFile = [System.IO.Path]::GetTempFileName()
						$Project.SaveTaskReportAs($R.Id, $TempFile, [Sdl.ProjectAutomation.Core.ReportFormat]::Xml)
						
						# Read the LCID information from temporarily saved report and use it to create CultureInfo, which is then used to get the culture code
						[xml] $TempReport = Get-Content -LiteralPath $TempFile
						[int] $LCID = $TempReport.task.taskInfo.language.lcid
						# all custom cultures share the same LCID 4096
						# and since the only other available info in the report is the language name, we have to find the name in list of all custom cultures
						if ($LCID -eq 4096) {
							$CustomCulturesList = [System.Globalization.CultureInfo]::GetCultures([System.Globalization.CultureTypes]::UserCustomCulture)
							$ReportCulture = $CustomCulturesList | Where-Object -Property EnglishName -eq $TempReport.task.taskInfo.language.name
						}
						else {
							$ReportCulture = New-Object System.Globalization.CultureInfo($LCID)
						}
						
						# construct the log file name (without extension, which is dependent on log save format)
						$ReportTargetLanguage = $ReportCulture.Name
						$ReportName = ($R.Name -replace 'Report','')  # "Analyze Files Report" => "Analyze Files "
						$LogFileName = "$ReportName$($SourceLanguage)_$ReportTargetLanguage"
						
						# If log location does not exist, create it
						if (!(Test-Path -LiteralPath $LogLocation)) {
							New-Item $LogLocation -Force -ItemType Directory | Out-Null
						}
						$LogLocation = (Resolve-Path -LiteralPath $LogLocation).ProviderPath
						
						# Save log in Trados 2007 format
						Write-Host "  $LogFileName.log"
						ConvertTo-TradosLog $TempFile "$LogLocation\$LogFileName.log"
						
						# Save log in Excel format
						if ($ExcelLog) {
							Write-Host "  $LogFileName.xlsx"
							$Project.SaveTaskReportAs($R.Id, "$LogLocation\$LogFileName.xlsx", [Sdl.ProjectAutomation.Core.ReportFormat]::Excel)
						}
						
						# Delete temporarily saved report
						Remove-Item $TempFile -Force
					}  # if $R.TaskTemplateId -eq AnalyzeFiles
				}  # ForEach $R in $T.Reports
			}  # ForEach $T in $Tasks.SubTasks
		}  # if $Analyze and $LogLocation
	}

	# Add project source files
	Write-Host "`nAdding source files..." -ForegroundColor White
	$ProjectFiles = $Project.AddFolderWithFiles($SourceLocation, $true)
	Write-Host "Done"

	# Get source language project files IDs
	[Sdl.ProjectAutomation.Core.ProjectFile[]] $ProjectFiles = $Project.GetSourceLanguageFiles()
	[System.Guid[]] $SourceFilesGuids = Get-Guids $ProjectFiles

	# Run preparation tasks
	Write-Host "`nRunning preparation tasks..." -ForegroundColor White
	Write-Host "Task Scan"
	Validate-Task $Project.RunAutomaticTask($SourceFilesGuids, [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::Scan, $OnTaskProgressDelegate, $OnTaskMessageDelegate)
	Write-Host "Task Convert to Translatable Format"
	Validate-Task $Project.RunAutomaticTask($SourceFilesGuids, [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::ConvertToTranslatableFormat, $OnTaskProgressDelegate, $OnTaskMessageDelegate)
	Write-Host "Task Copy to Target Languages"
	Validate-Task $Project.RunAutomaticTask($SourceFilesGuids, [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::CopyToTargetLanguages, $OnTaskProgressDelegate, $OnTaskMessageDelegate)
	Write-Host "Done"

	if ($PerfectMatch) {
		Write-Host "`nAssigning bilingual files for PerfectMatch..." -ForegroundColor White
		$BilingualFileMappings = Get-BilingualFileMappings -LanguagesList $ProjectInfo.TargetLanguages -TranslatableFilesList $ProjectFiles -BilingualsPath $BilingualsPath
		$Project.AddBilingualReferenceFiles($BilingualFileMappings)
		Write-Host "Done"
	}


	# "safety save" before running automatic tasks, so that project is not completely lost in case of fatal crash
	$Project.Save()

	if ($Pretranslate -or $Analyze -or $PerfectMatch) {
		Write-Host "`nRunning automatic tasks..." -ForegroundColor White
		
		# construct task sequence
		[String[]] $TaskSequence = @()
		if ($PerfectMatch) {
			$TaskSequence += [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::PerfectMatch
		}
		if ($Pretranslate) {
			$TaskSequence += [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::PreTranslateFiles
		}
		if ($Analyze) {
			$TaskSequence += [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::AnalyzeFiles
		}

		# run the task sequence, either per language or all languages at once
		if ($PerLanguage) {
			Write-Host "Per-language mode selected"
			ForEach ($TargetLanguage in $TargetLanguagesList) {
				Write-Host "Processing $TargetLanguage" -ForegroundColor Yellow
				& $ProcessAutomaticTasks -LangsToProcess $TargetLanguage
			}
		}
		else {
			& $ProcessAutomaticTasks -LangsToProcess $TargetLanguagesList
		}
		Write-Host "Done"
	}

	# Save the project
	Write-Host "`nSaving project..." -ForegroundColor White
	$Project.Save()
	Write-Host "Done"
}

function Get-Project {
<#
.SYNOPSIS
Opens Trados Studio file based project.
.DESCRIPTION
Opens Trados Studio file based project in specified location.
#>
	param(
		# Path to directory where the project file is located.
		[Parameter (Mandatory = $true)]
		[Alias("Location","PrjLoc")]
		[String] $ProjectLocation
	)

	# get project file path
	$ProjectFilePath = (Get-ChildItem -LiteralPath $ProjectLocation *.sdlproj).FullName

	# get file based project
	$Project = New-Object Sdl.ProjectAutomation.FileBased.FileBasedProject $ProjectFilePath

	return $Project
}

function Remove-Project {
<#
.SYNOPSIS
Deletes Trados Studio file based project.
.DESCRIPTION
Deletes Trados Studio file based project in specified location.
Complete project is deleted, including the project location directory.
#>
	param (
		# Path to directory where the project file is located.
		[Parameter (Mandatory = $true)]
		[Alias("Location","PrjLoc")]
		[String] $ProjectLocation
	)

	$Project = Get-Project (Resolve-Path -LiteralPath $ProjectLocation).ProviderPath
	$Project.Delete()
}

function Get-TaskFileInfoFiles {
	param(
		[Sdl.ProjectAutomation.FileBased.FileBasedProject] $Project,
		[Sdl.Core.Globalization.Language] $Language
	)

	[Sdl.ProjectAutomation.Core.TaskFileInfo[]]$TaskFilesList = @()
	ForEach ($Taskfile in $Project.GetTargetLanguageFiles($Language)) {
		$FileInfo = New-Object Sdl.ProjectAutomation.Core.TaskFileInfo
		$FileInfo.ProjectFileId = $Taskfile.Id
		$FileInfo.ReadOnly = $false
		$TaskFilesList += $FileInfo
	}
	return $TaskFilesList
}

function Write-TaskProgress {
	param(
	$Caller,
	$ProgressEventArgs
	)

	$Percent = $ProgressEventArgs.PercentComplete
	if ($Percent -eq 100) {
		$Status = "Completed"
		$Status = $ProgressEventArgs.Status
	}
	else {
		$Status = $ProgressEventArgs.Status
	}

	$Cancel = $ProgressEventArgs.Cancel
	$Task = $ProgressEventArgs.TaskTemplateIds

	# write textual progress percentage in console
	if ($host.name -eq 'ConsoleHost') {
		Write-Host "$($Percent.ToString().PadLeft(5))%	$Status	$StatusMessage`r" -NoNewLine
		# when all is done, output nothing WITH NEW LINE
		if ($Percent -eq 100 -and $Status -eq "Completed") {
			Write-Host $null
		}
	}
	# use PowerShell progress bar in PowerShell environment since it does not support writing on the same line using `r
	else {
		Write-Progress -Activity "Processing task" -PercentComplete $Percent -Status $Status
		# when all is done, remove the progress bar
		if ($Percent -eq 100 -and $Status -eq "Completed") {
			Write-Progress -Activity "Processing task" -Completed
		}
	}
}

function Write-TaskMessage {
	param(
	$Caller,
	[Sdl.ProjectAutomation.Core.TaskMessageEventArgs] $MessageEventArgs
	)

	$Message = $MessageEventArgs.Message
	Write-Host "`n$($Message.Level): $($Message.Message)" -ForegroundColor DarkYellow

	if ($($Message.Exception) -ne $null) {
		Write-Host "$($Message.Exception)" -ForegroundColor Magenta
	}
}

function ConvertTo-TradosLog {
<#
.SYNOPSIS
Converts Studio XML-formatted task report to Trados 2007 plain text format
.DESCRIPTION
Converts Studio XML-formatted task report to Trados 2007 plain text format.
Output can optionally contain only total summary, omitting details for individual files.
.EXAMPLE
ConvertTo-TradosLog -Path "D:\Project\Reports\Analyze Files_en-US_fi-FI.xml" -Destination "D:\Finnish analysis.log"

Converts the "D:\Project\Reports\Analyze Files_en-US_fi-FI.xml" report to Trados 2007-like plain text log and saves the result to "D:\Finnish analysis.log" file.
.EXAMPLE
ConvertTo-TradosLog -Path "D:\Project\Reports" -TotalOnly -Recurse

Converts the all reports found in "D:\Project\Reports" directory and its subdirectories Trados 2007-like plain text logs, which will include only total summaries. Converted files will be saved in the directories along the source reports.
#>
	[CmdletBinding()]
	
	param(
		# Path to either a single XML-formatted Studio report, or a directory where multiple XML reports are located.
		[Parameter (Mandatory = $true)]
		[String] $Path,
		
		# Path to a single output text file. Can be used only if Path specifies a single file, not directory.
		# If this parameter is omitted, output file will be created in the same location as input file, using input file name and ".log" extension.
		[String] $Destination,
		
		# Include only total summary (do not include analyses for individual files)
		[Switch] $TotalOnly,
		
		# If Path specifies a directory, convert also all task reports found in subdirectories of the specified path.
		[Alias("r")]
		[Switch] $Recurse
	)
	#

	function Write-LogItem {
		param($Item)
		"{0,-12}{1,10}{2,13}{3,8}{4,11}" -f " Match Types",
							"Segments",
							"Words",
							"Percent",
							"Placeables"
		"{0,-12}{1,10}{2,13}{3,8}{4,11}" -f " Context TM",
							([int]$Item.analyse.perfect.segments + [int]$Item.analyse.inContextExact.segments).ToString("n0",$culture),
							([int]$Item.analyse.perfect.words + [int]$Item.analyse.inContextExact.words).ToString("n0",$culture),
							$(
								if ([int]$Item.analyse.total.words -gt 0) {
									([int]$Item.analyse.perfect.words + [int]$Item.analyse.inContextExact.words) / [int]$Item.analyse.total.words * 100
								}
								else {
									[int]$Item.analyse.total.words
								}
							).ToString("n0",$culture),
							([int]$Item.analyse.perfect.placeables + [int]$Item.analyse.inContextExact.placeables).ToString("n0",$culture)
		"{0,-12}{1,10}{2,13}{3,8}{4,11}" -f " Repetitions",
							([int]$Item.analyse.crossFileRepeated.segments + [int]$Item.analyse.repeated.segments).ToString("n0",$culture),
							([int]$Item.analyse.crossFileRepeated.words + [int]$Item.analyse.repeated.words).ToString("n0",$culture),
							$(
								if ([int]$Item.analyse.total.words -gt 0) {
									([int]$Item.analyse.crossFileRepeated.words + [int]$Item.analyse.repeated.words) / [int]$Item.analyse.total.words * 100
								}
								else {
									[int]$Item.analyse.total.words
								}
							).ToString("n0",$culture),
							([int]$Item.analyse.crossFileRepeated.placeables + [int]$Item.analyse.repeated.placeables).ToString("n0",$culture)
		"{0,-12}{1,10}{2,13}{3,8}{4,11}" -f " 100%",
							([int]$Item.analyse.exact.segments).ToString("n0",$culture),
							([int]$Item.analyse.exact.words).ToString("n0",$culture),
							$(
								if ([int]$Item.analyse.total.words -gt 0) {
									[int]$Item.analyse.exact.words / [int]$Item.analyse.total.words * 100
								}
								else {
									[int]$Item.analyse.total.words
								}
							).ToString("n0",$culture),
							([int]$Item.analyse.exact.placeables).ToString("n0",$culture)
		$Item.analyse.fuzzy | ForEach {
			$intfuzzy = $Item.analyse.internalFuzzy | where min -eq $_.min
			"{0,-12}{1,10}{2,13}{3,8}{4,11}" -f " $($_.min)% - $($_.max)%",
								([int]$_.segments + [int]$intfuzzy.segments).ToString("n0",$culture),
								([int]$_.words + [int]$intfuzzy.words).ToString("n0",$culture),
								$(
									if ([int]$Item.analyse.total.words -gt 0) {
										([int]$_.words + [int]$intfuzzy.words) / [int]$Item.analyse.total.words * 100
									}
									else {
										[int]$Item.analyse.total.words
									}
								).ToString("n0",$culture),
								([int]$_.placeables + [int]$intfuzzy.placeables).ToString("n0",$culture)
		} | Sort-Object -Descending
		"{0,-12}{1,10}{2,13}{3,8}{4,11}" -f " No Match",
							([int]$Item.analyse.new.segments).ToString("n0",$culture),
							([int]$Item.analyse.new.words).ToString("n0",$culture),
							$(
								if ([int]$Item.analyse.total.words -gt 0) {[int]$Item.analyse.new.words / [int]$Item.analyse.total.words * 100
								}
								else {
									[int]$Item.analyse.total.words
								}
							).ToString("n0",$culture),
							([int]$Item.analyse.new.placeables).ToString("n0",$culture)
		"{0,-12}{1,10}{2,13}{3,8}{4,11}" -f " Total",
							([int]$Item.analyse.total.segments).ToString("n0",$culture),
							([int]$Item.analyse.total.words).ToString("n0",$culture),
							$(
								if ([int]$Item.analyse.total.words -gt 0) {
									[int]$Item.analyse.total.words / [int]$Item.analyse.total.words * 100
								}
								else {
									[int]$Item.analyse.total.words
								}
							).ToString("n0",$culture),
							([int]$Item.analyse.total.placeables).ToString("n0",$culture)
	}

	function Write-LogFile {
		param($File)
		Write-Output "Start Analyse: $($report.task.taskInfo.runAt)"
		Write-Output ""

		# Temporarily redefine Output Field Separator to have possible multiple TM names nicely aligned
		# e.g.:
		# Translation Memory: foo.sdltm
		#                     bar.sdltm
		#                     foobar.sdltm
		$tempOFS = $OFS
		$OFS="`n".PadRight(21)

		Write-Output "Translation Memory: $($report.task.taskInfo.tm.name)"

		# Restore original Output Field Separator
		$OFS = $tempOFS

		Write-Output ""

		$filescount = 0
		$File.task.file | ForEach {
			$filescount++
			if (!($TotalOnly)) {
					Write-Output $_.name
					Write-Output ""
					Write-Output $(Write-LogItem $_)
					Write-Output ""
			}
		}

		Write-Output "Analyse Total ($($filescount) files):"
		Write-Output ""
		Write-Output $(Write-LogItem $report.task.batchTotal)
		Write-Output ""
		Write-Output "Analyse finished successfully without errors!"
		Write-Output ""
		Write-Output $(Get-Date).ToString("F",$culture)
		Write-Output "================================================================================"
	}

	# main ConvertTo-TradosLog function body
	$culture = New-Object System.Globalization.CultureInfo("en-US")

	Get-ChildItem $Path *.xml -File -Recurse:$Recurse | ForEach {
		[xml] $report = Get-Content ($_.FullName)
		# Only "analyse" report is supported
		if ($report.task.name -ne "analyse") {
			return
		} else {
			if ($Destination -ne $null -and $Destination -ne "") {
				$OutputFile = $Destination
			} else {
				$OutputFile = ($_.Fullname -replace 'xml$','log')
			}
			Write-LogFile $report | Out-File -LiteralPath $OutputFile -Encoding UTF8 -Force
		}
	}
}

function Export-TargetFiles {
<#
.SYNOPSIS
Exports target files from Trados Studio file based project.
.DESCRIPTION
Exports target files from Trados Studio file based project to specified location.
.EXAMPLE
Export-TargetFiles -ProjectLocation "D:\Project" -ExportLocation "D:\Export"

Exports target files for all target languages defined in project located in "D:\Project" folder;
files will be created in "D:\Export" folder.
.EXAMPLE
Export-TargetFiles -ProjectLocation "D:\Project" -ExportLocation "D:\Export" -TargetLanguages "fi-FI,sv-SE"

Exports target files for Finnish and Swedish languages from project located in "D:\Project" folder;
files will be created in "D:\Export" folder.
#>
	param (
		# Path to directory where the project file is located.
		[Parameter (Mandatory = $true)]
		[Alias("Location","PrjLoc")]
		[String] $ProjectLocation,

		# Path to directory where the exported files will be placed.
		# If the directory does not exist, it will be created.
		[Parameter (Mandatory = $true)]
		[Alias("ExpLoc","Export")]
		[String] $ExportLocation,
		
		# Space-, comma- or semicolon-separated list of locale codes of project target languages.
		# For locale codes, see https://msdn.microsoft.com/en-us/goglobal/bb896001.aspx
		[Alias("TrgLng")]
		[String] $TargetLanguages
	)

	# if export location does not exist, create it
	if (!(Test-Path -LiteralPath $ExportLocation)) {
		New-Item $ExportLocation -Force -ItemType Directory | Out-Null
	}
	$ExportLocation = (Resolve-Path -LiteralPath $ExportLocation).ProviderPath

	# get project and its settings
	$Project = Get-Project (Resolve-Path -LiteralPath $ProjectLocation).ProviderPath
	$ProjectSettings = $Project.GetSettings()

	# update the ExportLocation value in project settings
	# (need to use this harakiri due to missing direct support for generic methods invocation in PowerShell)
	$method = $ProjectSettings.GetType().GetMethod("GetSettingsGroup",[System.Type]::EmptyTypes)
	$closedMethod = $method.MakeGenericMethod([Sdl.ProjectAutomation.Settings.ExportFilesSettings])
	$closedMethod.Invoke($ProjectSettings,[System.Type]::EmptyTypes).ExportLocation.Value = $ExportLocation
	
	$Project.UpdateSettings($ProjectSettings)
	$Project.Save()
	
	if ($TargetLanguages -ne $null -and $TargetLanguages -ne "") {
		# Parse target languages from provided parameter
		$TargetLanguagesList = $TargetLanguages -Split $LanguagesSeparator
	}
	else {
		# Get project languages
		$TargetLanguagesList = @($Project.GetProjectInfo().TargetLanguages.IsoAbbreviation)
	}
	
	Write-Host "`nExporting files..." -ForegroundColor White
	
	# get IDs of all target project files
	[System.Guid[]] $TargetFilesGuids = @()
	ForEach ($TargetLanguage in $TargetLanguagesList) {
		$TargetFiles = $Project.GetTargetLanguageFiles($TargetLanguage)
		$TargetFilesGuids += Get-Guids $TargetFiles
	}
	
	# task progress event handler
	$OnTaskProgress = ${function:Write-TaskProgress} -as [System.EventHandler[Sdl.ProjectAutomation.Core.TaskStatusEventArgs]]
	$OnTaskProgressDelegate = [RunspacedDelegateFactory]::NewRunspacedDelegate($OnTaskProgress, [Runspace]::DefaultRunspace)
	
	# task messages event handler
	$OnTaskMessage = ${function:Write-TaskMessage} -as [System.EventHandler[Sdl.ProjectAutomation.Core.TaskMessageEventArgs]]
	$OnTaskMessageDelegate = [RunspacedDelegateFactory]::NewRunspacedDelegate($OnTaskMessage, [Runspace]::DefaultRunspace)
	
	# run (and then validate) the task sequence
	$Task = $Project.RunAutomaticTask($TargetFilesGuids, [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::ExportFiles, $OnTaskProgressDelegate, $OnTaskMessageDelegate)
	Validate-Task $Task
}

function Update-MainTMs {
<#
.SYNOPSIS
Updates main TMs of Trados Studio file based project.
.DESCRIPTION
Updates main TMs of Trados Studio file based project with content from the bilingual target language files.
.EXAMPLE
Update-MainTMs -ProjectLocation "D:\Project"

Updates main TMs for all target languages defined in project located in "D:\Project" folder.
.EXAMPLE
Update-MainTMs -ProjectLocation "D:\Project" -TargetLanguages "fi-FI,sv-SE"

Updates main TMs for Finnish and Swedish languages from project located in "D:\Project" folder.
#>
	param (
		# Path to directory where the project file is located.
		[Parameter (Mandatory = $true)]
		[Alias("Location","PrjLoc")]
		[String] $ProjectLocation,
		
		# Space-, comma- or semicolon-separated list of locale codes of project target languages.
		# For locale codes, see https://msdn.microsoft.com/en-us/goglobal/bb896001.aspx
		[Alias("TrgLng")]
		[String] $TargetLanguages
	)

	# get project and its settings
	$Project = Get-Project (Resolve-Path -LiteralPath $ProjectLocation).ProviderPath
	$ProjectSettings = $Project.GetSettings()

	if ($TargetLanguages -ne $null -and $TargetLanguages -ne "") {
		# Parse target languages from provided parameter
		$TargetLanguagesList = $TargetLanguages -Split $LanguagesSeparator
	}
	else {
		# Get project languages
		$TargetLanguagesList = @($Project.GetProjectInfo().TargetLanguages.IsoAbbreviation)
	}

	Write-Host "`nUpdating main translation memories..." -ForegroundColor White

	# get IDs of all target project files
	[System.Guid[]] $TargetFilesGuids = @()
	ForEach ($TargetLanguage in $TargetLanguagesList) {
		$TargetFiles = $Project.GetTargetLanguageFiles($TargetLanguage)
		$TargetFilesGuids += Get-Guids $TargetFiles
	}
	
	# task progress event handler
	$OnTaskProgress = ${function:Write-TaskProgress} -as [System.EventHandler[Sdl.ProjectAutomation.Core.TaskStatusEventArgs]]
	$OnTaskProgressDelegate = [RunspacedDelegateFactory]::NewRunspacedDelegate($OnTaskProgress, [Runspace]::DefaultRunspace)
	
	# task messages event handler
	$OnTaskMessage = ${function:Write-TaskMessage} -as [System.EventHandler[Sdl.ProjectAutomation.Core.TaskMessageEventArgs]]
	$OnTaskMessageDelegate = [RunspacedDelegateFactory]::NewRunspacedDelegate($OnTaskMessage, [Runspace]::DefaultRunspace)
	
	# run (and then validate) the task sequence
	$Task = $Project.RunAutomaticTask($TargetFilesGuids, [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::UpdateMainTranslationMemories, $OnTaskProgressDelegate, $OnTaskMessageDelegate)
	Validate-Task $Task

	Write-Host "Done"
}

function PseudoTranslate {
<#
.SYNOPSIS
Pseudo-translates Trados Studio project.
.DESCRIPTION
Pseudo-translates Trados Studio project, allowing to define specific pseudo-translation options.
Optionally also exports the pseudo-translated target files.
.EXAMPLE
PseudoTranslate -ProjectLocation "D:\Project"

Pseudo-translates all target languages defined in project located in "D:\Project" folder.
.EXAMPLE
PseudoTranslate -ProjectLocation "D:\Project" -TargetLanguages "fi-FI,sv-SE"

Pseudo-translates Finnish and Swedish languages from project located in "D:\Project" folder.
.EXAMPLE
Export-Package -PrjLoc "D:\Project" -TrgLng "fi-FI,sv-SE" -ExpLoc "D:\Export"

Pseudo-translates Finnish and Swedish languages from project located in "D:\Project" folder and exports the pseudo-translated files to "D:\Export" folder.
#>
	param (
		# Path to directory where the project file is located.
		[Parameter (Mandatory = $true)]
		[Alias("Location","PrjLoc")]
		[String] $ProjectLocation,
		
		# Space-, comma- or semicolon-separated list of locale codes of project target languages.
		# For locale codes, see https://msdn.microsoft.com/en-us/goglobal/bb896001.aspx
		[Alias("TrgLng")]
		[String] $TargetLanguages,

		# Generate the same translations every time pseudo-translate is run. Otherwise, random words will be used.
		[Alias("Determ","Det")]
		[Switch] $Deterministic,

		# String that will be appended to the start of each pseudo-translation.
		[Alias("Start")]
		[String] $AppendStart = "_",

		# String that will be appended to the end of each pseudo-translation.
		[Alias("End")]
		[String] $AppendEnd = "_",

		# AppendStart and AppendEnd strings are applied at paragraph level. Otherwise they are applied at segment level.
		[Alias("Paragraph","Par")]
		[Switch] $AppendToParagraph,

		# The factor by which the length of pseudo-translated content will differ to the source content.
		[ValidateRange(0,5.9)]
		[Alias("Expand","Exp")]
		[Double] $ExpandBy = 1.3,

		# Use dollar signs to represent characters in the pseudo-translations.
		[Switch] $Dollar,

		# Path to directory where to export pseudo-translated files.
		# If the directory does not exist, it will be created.
		[Alias("ExpLoc","Export")]
		[String] $ExportLocation

	)

	# get project and its settings
	$Project = Get-Project (Resolve-Path -LiteralPath $ProjectLocation).ProviderPath
	$ProjectSettings = $Project.GetSettings()

	$method = $ProjectSettings.GetType().GetMethod("GetSettingsGroup",[System.Type]::EmptyTypes)
	$closedMethod = $method.MakeGenericMethod([Sdl.ProjectAutomation.Settings.PseudoTranslateSettings])
	
	if ($Deterministic) {
		$closedMethod.Invoke($ProjectSettings,[System.Type]::EmptyTypes).DeterministicPseudoTranslation.Value = $true
	} else {
		$closedMethod.Invoke($ProjectSettings,[System.Type]::EmptyTypes).DeterministicPseudoTranslation.Value = $false
	}

	if ($AppendStart -eq $null -or $AppendStart -eq "") {
		$closedMethod.Invoke($ProjectSettings,[System.Type]::EmptyTypes).AppendStart.Value = $false
	} else {
		$closedMethod.Invoke($ProjectSettings,[System.Type]::EmptyTypes).AppendStart.Value = $true
		$closedMethod.Invoke($ProjectSettings,[System.Type]::EmptyTypes).AppendStartString.Value = $AppendStart
	}

	if ($AppendEnd -eq $null -or $AppendEnd -eq "") {
		$closedMethod.Invoke($ProjectSettings,[System.Type]::EmptyTypes).AppendEnd.Value = $false
	} else {
		$closedMethod.Invoke($ProjectSettings,[System.Type]::EmptyTypes).AppendEnd.Value = $true
		$closedMethod.Invoke($ProjectSettings,[System.Type]::EmptyTypes).AppendEndString.Value = $AppendEnd
	}

	if ($AppendToParagraph) {
		$closedMethod.Invoke($ProjectSettings,[System.Type]::EmptyTypes).AppendToUnit.Value = [Sdl.ProjectAutomation.Settings.AppendToUnitType]::Paragraph
	} else {
		$closedMethod.Invoke($ProjectSettings,[System.Type]::EmptyTypes).AppendToUnit.Value = [Sdl.ProjectAutomation.Settings.AppendToUnitType]::Segment
	}

	if ($ExpandBy -ne $null) {
		$closedMethod.Invoke($ProjectSettings,[System.Type]::EmptyTypes).ExpansionLimit.Value = $ExpandBy
	}

	if ($Dollar) {
		$closedMethod.Invoke($ProjectSettings,[System.Type]::EmptyTypes).UseDictionary.Value = $false
	} else {
		$closedMethod.Invoke($ProjectSettings,[System.Type]::EmptyTypes).UseDictionary.Value = $true
	}

	$Project.UpdateSettings($ProjectSettings)

	if ($TargetLanguages -ne $null -and $TargetLanguages -ne "") {
		# Parse target languages from provided parameter
		$TargetLanguagesList = $TargetLanguages -Split $LanguagesSeparator
	}
	else {
		# Get project languages
		$TargetLanguagesList = @($Project.GetProjectInfo().TargetLanguages.IsoAbbreviation)
	}

	Write-Host "`nPseudo-translating..." -ForegroundColor White

	# get IDs of all target project files
	[System.Guid[]] $TargetFilesGuids = @()
	ForEach ($TargetLanguage in $TargetLanguagesList) {
		$TargetFiles = $Project.GetTargetLanguageFiles($TargetLanguage)
		$TargetFilesGuids += Get-Guids $TargetFiles
	}
	
	# task progress event handler
	$OnTaskProgress = ${function:Write-TaskProgress} -as [System.EventHandler[Sdl.ProjectAutomation.Core.TaskStatusEventArgs]]
	$OnTaskProgressDelegate = [RunspacedDelegateFactory]::NewRunspacedDelegate($OnTaskProgress, [Runspace]::DefaultRunspace)
	
	# task messages event handler
	$OnTaskMessage = ${function:Write-TaskMessage} -as [System.EventHandler[Sdl.ProjectAutomation.Core.TaskMessageEventArgs]]
	$OnTaskMessageDelegate = [RunspacedDelegateFactory]::NewRunspacedDelegate($OnTaskMessage, [Runspace]::DefaultRunspace)
	
	# run (and then validate) the task sequence
	$Task = $Project.RunAutomaticTask($TargetFilesGuids, [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::PseudoTranslateFiles, $OnTaskProgressDelegate, $OnTaskMessageDelegate)
	Validate-Task $Task

	Write-Host "Done"

	if ($ExportLocation) {
		Export-TargetFiles -ProjectLocation $ProjectLocation -ExportLocation $ExportLocation -TargetLanguages $TargetLanguages
	}

	# Save the project
	Write-Host "`nSaving project..." -ForegroundColor White
	$Project.Save()
	Write-Host "Done"
}

function Get-BilingualFileMappings {
	param (
		[Sdl.Core.Globalization.Language[]] $LanguagesList,
		[Sdl.ProjectAutomation.Core.ProjectFile[]] $TranslatableFilesList,
		[String] $BilingualsPath
	)

	[Sdl.ProjectAutomation.Core.BilingualFileMapping[]] $mappings = @()
	ForEach ($Language in $LanguagesList) {
		Write-Host "Processing $Language" -ForegroundColor Yellow
		$BilingualsCount = 0
		$SearchPath = Join-Path -Path $BilingualsPath -ChildPath $Language.IsoAbbreviation
		foreach ($file in $TranslatableFilesList) {
			if ($file.Name.EndsWith(".sdlxliff")) {
				$suffix = ""
			}
			else {
				$suffix = ".sdlxliff"
			}
			$BilingualFile = $(Join-Path -Path $SearchPath -ChildPath $file.Folder | Join-Path -ChildPath $file.Name) + $suffix
			if (Test-Path $BilingualFile) {
				$mapping = New-Object Sdl.ProjectAutomation.Core.BilingualFileMapping
				$mapping.BilingualFilePath = $BilingualFile
				$mapping.Language = $Language
				$mapping.FileId = $file.Id
				$mappings += $mapping
				$BilingualsCount += 1
			}
		}
 		Write-Host "  Assigned $BilingualsCount of $($TranslatableFilesList.Count) files"
	}
	return $mappings
}

function Validate-Task {
	param (
		[Sdl.ProjectAutomation.Core.AutomaticTask] $taskToValidate
	)

	if ($taskToValidate.Status -eq [Sdl.ProjectAutomation.Core.TaskStatus]::Completed) {
		Write-Host "Task $($taskToValidate.Name) successfully completed." -ForegroundColor green
	}
	else {
		switch ($taskToValidate.Status) {
			([Sdl.ProjectAutomation.Core.TaskStatus]::Failed).ToString() {
				Write-Host "Task $($taskToValidate.Name) failed." -ForegroundColor red
			}
			([Sdl.ProjectAutomation.Core.TaskStatus]::Invalid).ToString() {
				Write-Host "Task $($taskToValidate.Name) not valid." -ForegroundColor red
			}
			([Sdl.ProjectAutomation.Core.TaskStatus]::Rejected).ToString() {
				Write-Host "Task $($taskToValidate.Name) rejected." -ForegroundColor red
			}
			([Sdl.ProjectAutomation.Core.TaskStatus]::Cancelled).ToString() {
				Write-Host "Task $($taskToValidate.Name) cancelled." -ForegroundColor red
			}
			Default {
				Write-Host "Task $($taskToValidate.Name) status:  $($taskToValidate.Status)" -ForegroundColor cyan
			}
		}
		ForEach ($message in $taskToValidate.Messages) {
			if ($message.ProjectFileId -ne $null) {
				$AffectedFile = @($Project.GetTargetLanguageFiles()).Where({$_.Id -eq $message.ProjectFileId})
				Write-Host "$($AffectedFile.Language)`t$($AffectedFile.Folder)$($AffectedFile.Name)"
			}
			Write-Host "$($message.Message -Replace '(`n|`r)+$','')" -ForegroundColor red
		}
	}
}

function Validate-TaskSequence {
	param (
		[Sdl.ProjectAutomation.FileBased.TaskSequence] $TaskSequenceToValidate
	)

	ForEach ($Task in $TaskSequenceToValidate.SubTasks) {
		Validate-Task $Task
	}
}

Set-Alias -Name Pseudo -Value PseudoTranslate
Export-ModuleMember New-Project
Export-ModuleMember Get-Project
Export-ModuleMember Remove-Project
Export-ModuleMember ConvertTo-TradosLog
Export-ModuleMember Export-TargetFiles
Export-ModuleMember Update-MainTMs
Export-ModuleMember PseudoTranslate -Alias Pseudo
Export-ModuleMember Get-BilingualFileMappings
Export-ModuleMember Get-TaskFileInfoFiles
Export-ModuleMember Get-DefaultProjectTemplate
