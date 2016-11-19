param([String]$StudioVersion = "Studio4")

if ("${Env:ProgramFiles(x86)}") {
	$ProgramFilesDir = "${Env:ProgramFiles(x86)}"
}
else {
	$ProgramFilesDir = "${Env:ProgramFiles}"
}

Add-Type -Path "$ProgramFilesDir\SDL\SDL Trados Studio\$StudioVersion\Sdl.ProjectAutomation.FileBased.dll"
Add-Type -Path "$ProgramFilesDir\SDL\SDL Trados Studio\$StudioVersion\Sdl.ProjectAutomation.Core.dll"

##########################################################################################################
# Due to API bug basing new projects on "Default.sdltpl" template instead of actual default project template,
# we need to find the real default template configured in Trados Studio by reading the configuration files
switch ($StudioVersion) {
	"Studio2" {$StudioVersionAppData = "10.0.0.0"}
	"Studio3" {$StudioVersionAppData = "11.0.0.0"}
	"Studio4" {$StudioVersionAppData = "12.0.0.0"}
	"Studio5" {$StudioVersionAppData = "14.0.0.0"}
}
# Get default project template GUID from the user settings file
$DefaultProjectTemplateGuid = Select-Xml -Path "${Env:AppData}\SDL\SDL Trados Studio\$StudioVersionAppData\UserSettings.xml" -XPath "//Setting[@Id='DefaultProjectTemplateGuid']" | ForEach {$_.node.InnerXml}
# Get the location of local projects storage from ProjectApi configuration file
$LocalDataFolder = Select-Xml -Path "${Env:AppData}\SDL\ProjectApi\$StudioVersionAppData\SDL.ProjectApi.xml" -XPath "//LocalProjectServerInfo/@LocalDataFolder" | ForEach {$_.node.Value}
# Finally, get the default project template path from local project storage file
$DefaultProjectTemplate = Select-Xml -Path "$LocalDataFolder\projects.xml" -XPath "//ProjectTemplateListItem[@Guid='$DefaultProjectTemplateGuid']/@ProjectTemplateFilePath" | ForEach {$_.node.Value}
##########################################################################################################

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

		# Locale code of project source language.
		# See (incomplete) list of codes at https://www.microsoft.com/resources/msdn/goglobal/
		# Hint: Code for Latin American Spanish is "es-419" ;-)
		[Alias("SrcLng")]
		[String] $SourceLanguage,

		# Space- or comma- or semicolon-separated list of locale codes of project target languages.
		# See (incomplete) list of codes at https://www.microsoft.com/resources/msdn/goglobal/
		# Hint: Code for Latin American Spanish is "es-419" ;-)
		[Alias("TrgLng")]
		[String] $TargetLanguages,

		# Path to directory containing Trados Studio translation memories for project language pairs.
		# Directory will be searched for all TMs with language pairs defined for the project and found TMs will be assigned to the project languages.
		# Additional TMs defined in project template or reference project will be retained (unless "OverrideTM" parameter is specified).
		# Note: directory is NOT searched recursively!
		[Alias("TMLoc")]
		[String] $TMLocation,

		# Path to directory where Trados 2007-formatted textual logs will be created.
		# Any existing logs will be overwritten.
		[Alias("LogLoc")]
		[String] $LogLocation,

		# Path to project template (*.sdltpl) on which the created project will be based.
		# If this parameter is not specified, default project template set in Trados Studio will be used.
		[Parameter (ParameterSetName = "ProjectTemplate")]
		[Alias("PrjTpl")]
		[String] $ProjectTemplate = $DefaultProjectTemplate,

		# Path to project file (*.sdlproj) on which the created project will be based.
		[Parameter (ParameterSetName = "ProjectReference")]
		[Alias("PrjRef")]
		[String] $ProjectReference,

		# Ignore TMs defined in project template or reference project and use only TMs specified by "TMLocation" parameter.
		[Alias("OvrTM")]
		[Switch] $OverrideTM,

		# Run pre-translation task for each target language during project creation.
		[Alias("PreTra","Translate")]
		[Switch] $Pretranslate,

		# Run analysis task for each target language during project creation.
		[Alias("Analyse")]
		[Switch] $Analyze
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
		$TargetLanguagesList = $TargetLanguages -Split " |;|,"
		$ProjectInfo.TargetLanguages = Get-Languages $TargetLanguagesList
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
			$TMSourceLanguage = Get-TMSourceLanguage $TMPath | ForEach {Get-Language $_.Name}
			$TMTargetLanguage = Get-TMTargetLanguage $TMPath | ForEach {Get-Language $_.Name}

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
			else {Write-Host "$TMPath already in project"}
		}
	}

	# Add project source files
	Write-Host "`nAdding source files..." -ForegroundColor White
	$ProjectFiles = $Project.AddFolderWithFiles($SourceLocation, $true)

	# Get source language project files IDs
	[Sdl.ProjectAutomation.Core.ProjectFile[]] $ProjectFiles = $Project.GetSourceLanguageFiles()
	[System.Guid[]] $SourceFilesGuids = Get-Guids $ProjectFiles

	# Run preparation tasks
	Write-Host "`nRunning preparation tasks..." -ForegroundColor White
	Write-Host "Task Scan"
	Validate-Task $Project.RunAutomaticTask($SourceFilesGuids, [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::Scan, ${function:Write-TaskProgress}, ${function:Write-TaskMessage})
	Write-Host "Task Convert to Translatable Format"
	Validate-Task $Project.RunAutomaticTask($SourceFilesGuids, [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::ConvertToTranslatableFormat, ${function:Write-TaskProgress}, ${function:Write-TaskMessage})
	Write-Host "Task Copy to Target Languages"
	Validate-Task $Project.RunAutomaticTask($SourceFilesGuids, [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::CopyToTargetLanguages, ${function:Write-TaskProgress}, ${function:Write-TaskMessage})

	if ($Pretranslate -or $Analyze) {
		Write-Host "`nRunning pre-translation / analysis tasks..." -ForegroundColor White
		
		# get IDs of all target project files
		[System.Guid[]] $TargetFilesGuids = @()
		ForEach ($TargetLanguage in $TargetLanguagesList) {
			$TargetFiles = $Project.GetTargetLanguageFiles($TargetLanguage)
			$TargetFilesGuids += Get-Guids $TargetFiles
		}
		
		# construct task sequence
		[String[]] $TasksSequence = @()
		if ($Pretranslate) {
			$TasksSequence += [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::PreTranslateFiles
		}
		if ($Analyze) {
			$TasksSequence += [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::AnalyzeFiles
		}
		
		# run (and then validate) the task sequence
		# $Tasks = $Project.RunAutomaticTasks($TargetFilesGuids, $TasksSequence, ${function:Write-TaskProgress}, ${function:Write-TaskMessage})
		$Tasks = $Project.RunAutomaticTasks($TargetFilesGuids, $TasksSequence, $null, ${function:Write-TaskMessage})
		Validate-TaskSequence $Tasks
		
		# save analysis logs
		if ($Analyze -and $LogLocation) {
			Write-Host "`nSaving logs..." -ForegroundColor White
			ForEach ($T in $Tasks.SubTasks) {
				ForEach ($R in $T.Reports) {
					if ($R.TaskTemplateId -eq [string] [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::AnalyzeFiles) {
						# Save report to temporary file
						$TempFile = [System.IO.Path]::GetTempFileName()
						$Project.SaveTaskReportAs($R.Id, $TempFile, [Sdl.ProjectAutomation.Core.ReportFormat]::Xml)
						
						# Read the target language information from temporarily saved report
						[xml] $TempReport = Get-Content -LiteralPath $TempFile
						$ReportCulture = New-Object System.Globalization.CultureInfo([int]$TempReport.task.taskInfo.language.lcid)
						
						# If log location does not exist, create it
						if (!(Test-Path -LiteralPath $LogLocation)) {
							New-Item $LogLocation -Force -ItemType Directory | Out-Null
						}
						$LogLocation = (Resolve-Path -LiteralPath $LogLocation).ProviderPath
						$LogName = ($R.Name -replace 'Report','') + $SourceLanguage + "_" + $ReportCulture.Name + ".log"
						Write-Host $LogName
						ConvertTo-TradosLog $TempFile | Out-File -LiteralPath ($LogLocation + "\" + $LogName) -Encoding UTF8 -Force
						
						# Delete temporarily saved report
						#Remove-Item $TempFile -Force
					}
				}
			}
		}
	}

	# Save the project
	Write-Host "`nSaving project..." -ForegroundColor White
	$Project.Save()
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
		$TaskFilesList = $TaskFilesList + $FileInfo
	}
	return $TaskFilesList
}

function Write-TaskProgress {
	param(
	[Sdl.ProjectAutomation.FileBased.FileBasedProject] $Project,
	[Sdl.ProjectAutomation.Core.TaskStatusEventArgs] $ProgressEventArgs
	)

	$Percent = $ProgressEventArgs.PercentComplete
	$Status = $ProgressEventArgs.Status

	Write-Host "  $Percent%	$Status`r" -NoNewLine

#	if ($Status -ne $null -and $Status -ne "") {
#		Write-Progress -Activity "Processing task" -PercentComplete $Percent -Status $Status
#	}
}

function Write-TaskMessage {
	param(
	[Sdl.ProjectAutomation.FileBased.FileBasedProject] $Project,
	[Sdl.ProjectAutomation.Core.TaskMessageEventArgs] $MessageEventArgs
	)

	$Message = $MessageEventArgs.Message

	Write-Host "`n$($Message.Level): $($Message.Message)" -ForegroundColor DarkYellow
	if ($($Message.Exception) -ne $null) {Write-Host "$($Message.Exception)" -ForegroundColor Magenta}
}

function ConvertTo-TradosLog {
<#
.SYNOPSIS
Converts Studio XML-formatted task report to Trados 2007 plain text format
.DESCRIPTION
Converts Studio XML-formatted task report to Trados 2007 plain text format.
Output can optionally contain only total summary, omitting details for individual files.
.EXAMPLE
ConvertTo-TradosLog -Path "D:\Project\Reports\Analyze Files_en-US_fi-FI.xml" -TotalOnly > "D:\Analyze Files_en-US_fi-FI.log"

Converts the "D:\Project\Reports\Analyze Files_en-US_fi-FI.xml" report to Trados 2007-like plain text log and saves the result to "D:\Analyze Files_en-US_fi-FI.log" file.
.EXAMPLE
ConvertTo-TradosLog -Path "D:\Project\Reports\Analyze Files_en-US_fi-FI.xml" -TotalOnly | Out-File -Path "D:\Analyze Files_en-US_fi-FI.log" -Encoding UTF-8

Same as previous example, but using more 'PowerShell-ish' way – converts the "D:\Project\Reports\Analyze Files_en-US_fi-FI.xml" report to Trados 2007-like plain text log and saves the result to UTF-8 encoded  "D:\Analyze Files_en-US_fi-FI.log" file.
#>
	[CmdletBinding()]
	
	param(
		# Path to input XML-formatted Studio report
		[Parameter (Mandatory = $true)]
		[String] $Path,
		
		# Include only total summary (do not include analyses for individual files)
		[Switch] $TotalOnly
	)
	# 

	function Write-LogItem {
		param($Item)
		"{0,-12}{1,11}{2,13}{3,8}{4,11}" -f " Match Types","Segments","Words","Percent","Placeables"
		"{0,-12}{1,11}{2,13}{3,8}{4,11}" -f " Context TM",([int]$Item.analyse.perfect.segments + [int]$Item.analyse.inContextExact.segments).ToString("n0",$culture),([int]$Item.analyse.perfect.words + [int]$Item.analyse.inContextExact.words).ToString("n0",$culture),$null,([int]$Item.analyse.perfect.placeables + [int]$Item.analyse.inContextExact.placeables).ToString("n0",$culture)
		"{0,-12}{1,11}{2,13}{3,8}{4,11}" -f " Repetitions",([int]$Item.analyse.crossFileRepeated.segments + [int]$Item.analyse.repeated.segments).ToString("n0",$culture),([int]$Item.analyse.crossFileRepeated.words + [int]$Item.analyse.repeated.words).ToString("n0",$culture),$null,([int]$Item.analyse.crossFileRepeated.placeables + [int]$Item.analyse.repeated.placeables).ToString("n0",$culture)
		"{0,-12}{1,11}{2,13}{3,8}{4,11}" -f " 100%",([int]$Item.analyse.exact.segments).ToString("n0",$culture),([int]$Item.analyse.exact.words).ToString("n0",$culture),$null,([int]$Item.analyse.exact.placeables).ToString("n0",$culture)
		$Item.analyse.fuzzy | ForEach {
			$intfuzzy = $Item.analyse.internalFuzzy | where {$_.min -eq $min}
			"{0,-12}{1,11}{2,13}{3,8}{4,11}" -f " $($_.min)% - $($_.max)%",([int]$_.segments + [int]$intfuzzy.segments).ToString("n0",$culture),([int]$_.words + [int]$intfuzzy.words).ToString("n0",$culture),$null,([int]$_.placeables + [int]$intfuzzy.placeables).ToString("n0",$culture)
		} | Sort-Object -Descending
		"{0,-12}{1,11}{2,13}{3,8}{4,11}" -f " No Match",([int]$Item.analyse.new.segments).ToString("n0",$culture),([int]$Item.analyse.new.words).ToString("n0",$culture),$null,([int]$Item.analyse.new.placeables).ToString("n0",$culture)
		"{0,-12}{1,11}{2,13}{3,8}{4,11}" -f " Total",([int]$Item.analyse.total.segments).ToString("n0",$culture),([int]$Item.analyse.total.words).ToString("n0",$culture),$null,([int]$Item.analyse.total.placeables).ToString("n0",$culture)
	}

	$culture = New-Object System.Globalization.CultureInfo("en-US")

	[xml] $report = Get-Content (Resolve-Path -LiteralPath $Path)

	# Only "analyse" report is supported
	if ($report.task.name -ne "analyse") {return}

	Write-Output "Start Analyse: $($report.task.taskInfo.runAt)"
	Write-Output ""

	# Temporarily redefine Output Field Separator to have possible multiple TM names nicely aligned
	# e.g.:
	# Translation Memory: foo.sdltm
	#                     bar.sdltm
	#                     foobar.sdltm
	$tempOFS = $OFS
	$OFS="`n                    "

	Write-Output "Translation Memory: $($report.task.taskInfo.tm.name)"

	# Restore original Output Field Separator
	$OFS = $tempOFS

	Write-Output ""

	$filescount = 0
	$report.task.file | ForEach {
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
Export-TargetFiles -PrjLoc "D:\Project" -ExpLoc "D:\Export" -TrgLng "fi-FI,sv-SE"

Exports target files for Finnish and Swedish languages from project located in "D:\Project" folder;
files will be created in "D:\Export" folder.
#>
	param (
		# Path to directory where the project file is located.
		[Parameter (Mandatory = $true)]
		[Alias("Location","PrjLoc")]
		[String] $ProjectLocation,

		# Path to directory where the exported files will be placed.
		# Complete directory structure present in that directory will be added as project source.
		[Parameter (Mandatory = $true)]
		[Alias("ExpLoc")]
		[String] $ExportLocation,

		# Space- or comma- or semicolon-separated list of locale codes of project target languages.
		# See (incomplete) list of codes at https://www.microsoft.com/resources/msdn/goglobal/
		# Hint: Code for Latin American Spanish is "es-419" ;-)
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
	$Settings = $Project.GetSettings()

	# update the ExportLocation value in project settings
	# (need to use this harakiri due to missing direct support for generic methods invocation in PowerShell)
	$method = $Settings.GetType().GetMethod("GetSettingsGroup",[System.Type]::EmptyTypes)
	$closedMethod = $method.MakeGenericMethod([Sdl.ProjectAutomation.Settings.ExportFilesSettings])
	$closedMethod.Invoke($Settings,[System.Type]::EmptyTypes).ExportLocation.Value = $ExportLocation
	
	$Project.UpdateSettings($Settings)
	$Project.Save()
	
	if ($TargetLanguages -ne $null -and $TargetLanguages -ne "") {
		# Parse target languages from provided parameter
		$TargetLanguagesList = $TargetLanguages -Split " |;|,"
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
	
	# run (and then validate) the task sequence
	$Task = $Project.RunAutomaticTask($TargetFilesGuids, [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::ExportFiles, ${function:Write-TaskProgress}, ${function:Write-TaskMessage})
	Validate-Task $Task
}

function Validate-Task {
	param ([Sdl.ProjectAutomation.Core.AutomaticTask] $taskToValidate)

	if ($taskToValidate.Status -eq [Sdl.ProjectAutomation.Core.TaskStatus]::Failed) {
		Write-Host "Task"$taskToValidate.Name"failed."
		ForEach($message in $taskToValidate.Messages) {
			Write-Host $message.Message -ForegroundColor red
		}
	}
	if ($taskToValidate.Status -eq [Sdl.ProjectAutomation.Core.TaskStatus]::Invalid) {
		Write-Host "Task"$taskToValidate.Name"not valid."
		ForEach($message in $taskToValidate.Messages) {
			Write-Host $message.Message -ForegroundColor red
		}
	}
	if ($taskToValidate.Status -eq [Sdl.ProjectAutomation.Core.TaskStatus]::Rejected) {
		Write-Host "Task"$taskToValidate.Name"rejected."
		ForEach($message in $taskToValidate.Messages) {
			Write-Host $message.Message -ForegroundColor red
		}
	}
	if ($taskToValidate.Status -eq [Sdl.ProjectAutomation.Core.TaskStatus]::Cancelled) {
		Write-Host "Task"$taskToValidate.Name"cancelled."
		ForEach($message in $taskToValidate.Messages) {
			Write-Host $message.Message -ForegroundColor red
		}
	}
	if ($taskToValidate.Status -eq [Sdl.ProjectAutomation.Core.TaskStatus]::Completed) {
		Write-Host "Task"$taskToValidate.Name"successfully completed." -ForegroundColor green
	}
}

function Validate-TaskSequence {
	param ([Sdl.ProjectAutomation.FileBased.TaskSequence] $TaskSequenceToValidate)

	ForEach ($Task in $TaskSequenceToValidate.SubTasks) {
		Validate-Task $Task
	}
}


Export-ModuleMember New-Project
Export-ModuleMember Get-Project
Export-ModuleMember Remove-Project
Export-ModuleMember ConvertTo-TradosLog
Export-ModuleMember Export-TargetFiles
Export-ModuleMember Get-TaskFileInfoFiles