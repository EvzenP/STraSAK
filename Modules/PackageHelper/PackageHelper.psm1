param([String]$StudioVersion = "Studio4")

if ("${Env:ProgramFiles(x86)}") {
    $ProgramFilesDir = "${Env:ProgramFiles(x86)}"
}
else {
    $ProgramFilesDir = "${Env:ProgramFiles}"
}

Add-Type -Path "$ProgramFilesDir\SDL\SDL Trados Studio\$StudioVersion\Sdl.ProjectAutomation.FileBased.dll"
Add-Type -Path "$ProgramFilesDir\SDL\SDL Trados Studio\$StudioVersion\Sdl.ProjectAutomation.Core.dll"

$ProjectPackageExtension = ".sdlppx"
$ReturnPackageExtension = ".sdlrpx"

function Export-Package {
<#
.SYNOPSIS
Exports Trados Studio project packages from project.
.DESCRIPTION
Exports Trados Studio project packages from project to a specified location, allowing to define specific export options.
Packages location is created automatically. Separate package is created for every target language.
.EXAMPLE
Export-Package -ProjectLocation "D:\Project" -PackageLocation "D:\Packages"

Creates translation packages for all target languages defined in project located in "D:\Project" folder;
packages will be created in "D:\Packages" folder;
no project TMs, main TMs, termbases, etc. will be included in the packages.
.EXAMPLE
Export-Package -ProjectLocation "D:\Project" -PackageLocation "D:\Packages" -TargetLanguages "fi-FI,sv-SE" -IncTM -IncTB

Creates translation package for Finnish and Swedish target languages from project located in "D:\Project" folder;
package will be created in "D:\Packages" folder;
project TM will not be included in package, main TM and termbase will be included in package.
#>
	[CmdletBinding()]

	param(
		# Path to directory where the project is located.
		[Parameter (Mandatory = $true)]
		[Alias("Location","PrjLoc")]
		[String] $ProjectLocation,

		# Path to directory where the package will be created.
		[Parameter (Mandatory = $true)]
		[Alias("PkgLoc")]
		[String] $PackageLocation,

		# Space- or comma- or semicolon-separated list of locale codes of project target languages for which the package will be created.
		# See (incomplete) list of codes at https://www.microsoft.com/resources/msdn/goglobal/
		# Hint: Code for Latin American Spanish is "es-419" ;-)
		# If this parameter is omitted, packages for all project target languages are created.
		[Alias("TrgLng")]
		[String] $TargetLanguages,

		# Name of manual task which will be associated with files in the package
		[ValidateSet("Translate","Review")]
		[String] $Task = "Translate",

		# Option for project translation memory to be included in the package
		# None - do not include any project TM
		# UseExisting - include existing project TM
		# CreateNew - create new project TM
		[ValidateSet("None","UseExisting","CreateNew")]
		[Alias("PrjTM","ProjectTMs")]
		[String] $ProjectTM = "None",

		# Optional short comment to be included in package
		[Alias("PkgCmt")]
		[String] $PackageComment = "",

		# Include AutoSuggest dictionaries in package
		[Alias("IncAS","IncludeAutoSuggest")]
		[Switch] $IncludeAutoSuggestDictionaries,

		# Include main translation memories in package
		[Alias("IncTM","IncludeMainTM")]
		[Switch] $IncludeMainTMs,

		# Include termbases in package
		[Alias("IncTB","IncludeTermbase")]
		[Switch] $IncludeTermbases,

		# Recompute wordcount and analysis to update cross-file repetition counts
		# and include the recomputed reports in package
		[Alias("RecAna","RecomputeAnalyse","RecomputeAnalyze")]
		[Switch] $RecomputeAnalysis,

		# Include existing wordcount reports in package
		[Alias("IncRep","IncludeExistingReport","IncludeReports","IncludeReport")]
		[Switch] $IncludeExistingReports,

		# Keep automated translation providers information in package
		[Alias("KeepAT","KeepATProviders","KeepATProvider")]
		[Switch] $KeepAutomatedTranslationProviders,

		# Remove links to server-based translation memories from package
		[Alias("RmvSrvTM","RemoveServerTMs","RemoveServerTM")]
		[Switch] $RemoveServerBasedTMs
	)

	# If package location does not exist, create it
	if (!(Test-Path $PackageLocation)) 	{
		New-Item -Path $PackageLocation -Force -ItemType Directory | Out-Null
	}

	# Initialize default package creation options
	$PackageOptions = [Sdl.ProjectAutomation.Core.ProjectPackageCreationOptions] (New-DefaultPackageOptions)

	# Cast ProjectTM string value to corresponding enumeration value
	$PackageOptions.ProjectTranslationMemoryOptions = [Sdl.ProjectAutomation.Core.ProjectTranslationMemoryPackageOptions] $ProjectTM

	# Workaround for "create new project TM" option not actually working unless "include reports" is also set
	if ($ProjectTM -eq "CreateNew") {
		# if the IncludeReports property exists, set it to true
		# (this property was introduced only in Studio 2015 SR2 CU7)
		if ($PackageOptions.IncludeReports) {
			$PackageOptions.IncludeReports = $true
		}
	}

	# Set options according to provided switches
	if ($IncludeAutoSuggestDictionaries) {
		$PackageOptions.IncludeAutoSuggestDictionaries = $true
	}
	if ($IncludeMainTMs) {
		$PackageOptions.IncludeMainTranslationMemories = $true
	}
	if ($IncludeTermbases) {
		$PackageOptions.IncludeTermbases = $true
	}
	if ($KeepAutomatedTranslationProviders) {
		$PackageOptions.RemoveAutomatedTranslationProviders = $false
	}
	if ($RemoveServerBasedTMs) {
		$PackageOptions.RemoveServerBasedTranslationMemories = $true
	}
	if ($IncludeExistingReports) {
		# if the IncludeExistingReports property exists, set it to true
		# (this property was introduced only in Studio 2015 SR2 CU7)
		if ($PackageOptions.IncludeExistingReports) {
			$PackageOptions.IncludeExistingReports = $true
		}
		# if the IncludeReports property exists, set it to true
		# (this property was introduced only in Studio 2015 SR2 CU7)
		if ($PackageOptions.IncludeReports) {
			$PackageOptions.IncludeReports = $true
		}
	}
	if ($RecomputeAnalysis) {
		$PackageOptions.RecomputeAnalysisStatistics = $true
		# if the IncludeReports property exists, set it to true
		# (this property was introduced only in Studio 2015 SR2 CU7)
		if ($PackageOptions.IncludeReports) {
			$PackageOptions.IncludeReports = $true
		}
	}

	# According to info from SDL developer forum, using [DateTime]::MaxValue sets "no package due date"
	$PackageDueDate = [DateTime]::MaxValue

	$Project = Get-Project (Resolve-Path -LiteralPath $ProjectLocation).ProviderPath

	if ($TargetLanguages -ne $null -and $TargetLanguages -ne "") {
		# Parse target languages from provided parameter
		$TargetLanguagesList = $TargetLanguages -Split " |;|,"
	}
	else {
		# Get project languages
		$TargetLanguagesList = @($Project.GetProjectInfo().TargetLanguages.IsoAbbreviation)
	}

	Write-Host "`nCreating packages..." -ForegroundColor White

	# Loop through target languages and create package for each one
	Get-Languages $TargetLanguagesList | ForEach {

		$Language = $_
		$User = "$($Language.IsoAbbreviation) translator"
		# Set package name to project name with target language ISO code suffix
		$PackageName = "$($Project.GetProjectInfo().Name)_$($Language.IsoAbbreviation)"

		Write-Host "$PackageName$ProjectPackageExtension"

		# Get TaskFileInfo (files list) data for the target language's project files
		[Sdl.ProjectAutomation.Core.TaskFileInfo[]] $TaskFiles = Get-TaskFileInfoFiles $Project $Language
		# Create the manual task which will be associated with files being included in the package
		[Sdl.ProjectAutomation.Core.ManualTask] $ManualTask = $Project.CreateManualTask($Task, $User, $PackageDueDate, $TaskFiles)

		# Create package containing the manual task
		[Sdl.ProjectAutomation.Core.ProjectPackageCreation] $Package = $Project.CreateProjectPackage($ManualTask.Id, $PackageName, $PackageComment, $PackageOptions, ${function:Write-PackageProgress}, ${function:Write-PackageMessage})

		# Save the package to file in specified location
		if ($Package.Status -eq [Sdl.ProjectAutomation.Core.PackageStatus]::Completed) {
			$Project.SavePackageAs($Package.PackageId, "$PackageLocation\$PackageName$ProjectPackageExtension")
		}
		else {
			Write-Host "Package creation failed, cannot save it!"
		}

		Remove-Variable TaskFiles, ManualTask, Package
	}
}

function Import-Package {
<#
.SYNOPSIS
Imports Trados Studio return packages in project.
.DESCRIPTION
Imports Trados Studio return packages from specified location into a project stored in specified location.
.EXAMPLE
Import-Package -ProjectLocation "D:\Project" -PackageLocation "D:\Packages"

Imports all translation packages found in "D:\Packages" directory (and all its eventual subdirectories) into a project located in "D:\Project" folder.
.EXAMPLE
Import-Package -ProjectLocation "D:\Project" -PackageLocation "D:\Packages\Handback_en-US_fi-FI.sdlrpx"

Imports single  translation package from "D:\Packages\Handback_en-US_fi-FI.sdlrpx" folder  into a project located in "D:\Project" folder.
#>
	[CmdletBinding()]

	param(
		# Path to directory where the project is located.
		[Parameter (Mandatory = $true)]
		[Alias("Location","PrjLoc")]
		[String] $ProjectLocation,

		# Path to either a single package, or directory where multiple packages are located.
		[Parameter (Mandatory = $true)]
		[Alias("PkgLoc")]
		[String] $PackageLocation,

		# Imports also all return packages found in subdirectories of the specified path.
		[Alias("r")]
		[switch] $Recurse
	)

	$Project = Get-Project (Resolve-Path -LiteralPath $ProjectLocation).ProviderPath

	Write-Host "`nImporting packages..." -ForegroundColor White

	Get-ChildItem $PackageLocation *.sdlrpx -File -Recurse:$Recurse | ForEach {
		Write-Host "$($_.Name)"
		$PackageImport = $Project.ImportReturnPackage($_.FullName, ${function:Write-PackageProgress}, ${function:Write-PackageMessage})
	}
}

function Write-PackageProgress {
	param(
	$Caller,
	$ProgressEventArgs
	)

	$Cancel = $ProgressEventArgs.Cancel
	$Message = $ProgressEventArgs.StatusMessage

	if ($Message -ne $null -and $Message -ne "") {
		$Percent = $ProgressEventArgs.PercentComplete
		if ($Percent -eq 100) {
			$Message = "Completed"
		}

		# write textual progress percentage in console
		if ($host.name -eq 'ConsoleHost') {
			Write-Host "$($Percent.ToString().PadLeft(5))%	$Message"
			Start-Sleep -Seconds 1
		}
		# use PowerShell progress bar in PowerShell environment since it does not support writing on the same line using `r
		else {
			Write-Progress -Activity "Processing task" -PercentComplete $Percent -Status $Message
			# when all is done, remove the progress bar
			if ($Percent -eq 100 -and $Message -eq "Completed") {
				Write-Progress -Activity "Processing task" -Completed
			}
		}
	}
}

function Write-PackageMessage {
	param(
	$Caller,
	$MessageEventArgs
	)

	$Message = $MessageEventArgs.Message

	# do not pollute output with potentially unnecessary lines
	if ($Message.Source -ne "Package import") {
		Write-Host "$($Message.Source)" -ForegroundColor DarkYellow
	}
	Write-Host "$($Message.Level): $($Message.Message)" -ForegroundColor Magenta
	if ($Message.Exception) {
		Write-Host "$($Message.Exception)" -ForegroundColor Magenta
	}
}

function New-PackageOptions {
	[Sdl.ProjectAutomation.Core.ProjectPackageCreationOptions] $PackageOptions = New-Object Sdl.ProjectAutomation.Core.ProjectPackageCreationOptions
	return $PackageOptions
}

function New-DefaultPackageOptions {
	[Sdl.ProjectAutomation.Core.ProjectPackageCreationOptions] $PackageOptions = New-Object Sdl.ProjectAutomation.Core.ProjectPackageCreationOptions
	$PackageOptions.IncludeAutoSuggestDictionaries = $false
	$PackageOptions.IncludeMainTranslationMemories = $false
	$PackageOptions.IncludeTermbases = $false
	$PackageOptions.ProjectTranslationMemoryOptions = [Sdl.ProjectAutomation.Core.ProjectTranslationMemoryPackageOptions]::None
	if ($PackageOptions.IncludeReports) {$PackageOptions.IncludeReports = $false}
	if ($PackageOptions.IncludeExistingReports) {$PackageOptions.IncludeExistingReports = $false}
	$PackageOptions.RecomputeAnalysisStatistics = $false
	$PackageOptions.RemoveAutomatedTranslationProviders = $true
	$PackageOptions.RemoveServerBasedTranslationMemories = $false
	return $PackageOptions
}

Export-ModuleMember Export-Package
Export-ModuleMember Import-Package
#Export-ModuleMember New-PackageOptions
#Export-ModuleMember New-DefaultPackageOptions
