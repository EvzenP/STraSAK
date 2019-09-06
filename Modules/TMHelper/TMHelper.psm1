param([String]$StudioVersion = "Studio5")

if ("${Env:ProgramFiles(x86)}") {
    $ProgramFilesDir = "${Env:ProgramFiles(x86)}"
}
else {
    $ProgramFilesDir = "${Env:ProgramFiles}"
}

Add-Type -Path "$ProgramFilesDir\SDL\SDL Trados Studio\$StudioVersion\Sdl.LanguagePlatform.TranslationMemoryApi.dll"
Add-Type -Path "$ProgramFilesDir\SDL\SDL Trados Studio\$StudioVersion\Sdl.LanguagePlatform.TranslationMemory.dll"

$LanguagesSeparator = "\s+|;\s*|,\s*"

$SDLTMFileExtension = ".sdltm"
$TMXFileExtension = ".tmx"

function New-FileBasedTM {
<#
.SYNOPSIS
Creates a new file based TM.
.DESCRIPTION
Creates new file based translation memory using specified source language and target language(s).
Optionally also a TM description and further TM behavior options can be specified.
TM can be created
– using specified name and location
 (where the name will be used for the internal TM 'display name' and for filename of the created *.sdltm file)
– using specified full file path
 (where the *.sdltm file name will be used for the internal TM 'display name')
If using only single target language, the specified TM name or .sdltm filename is used as-is.
If using multiple target languages, the source- and target language ISO abbreviations are added as suffix to the specified TM name or .sdltm filename.
.EXAMPLE
New-FileBasedTM -Name "Contoso English-German Main" -TMLocation "D:\Projects\TMs" -SourceLanguage "en-US" -TargetLanguages "de-DE"

Creates "D:\Projects\TMs\Contoso English-German Main.sdltm" translation memory with "English (United States)" source language and "German (Germany)" target language and "Contoso English-German Main" internal friendly name.
.EXAMPLE
New-FileBasedTM -Name "Contoso Main" -TMLocation "D:\Projects\TMs" -SourceLanguage "en-US" -TargetLanguages "de-DE fr-FR it-IT"

Creates "Contoso Main en-US_de-DE.sdltm", "Contoso Main en-US_fr-FR.sdltm" and "Contoso Main en-US_it-IT.sdltm" translation memories in "D:\Projects\TMs" location, with "English (United States)" source language and "German (Germany)", "French (France)" and "Italian (Italy)" respective target languages. Internal friendly TM names will be set to "Contoso Main en-US_de-DE", "Contoso Main en-US_fr-FR" and "Contoso Main en-US_it-IT" respectively.
.EXAMPLE
New-FileBasedTM -Path "D:\Projects\TMs\Contoso_English-German_Main.sdltm" -SrcLng "en-US" -TrgLng "de-DE"

Creates "D:\Projects\TMs\Contoso_English-German_Main.sdltm" translation memory with "English (United States)" source language and "German (Germany)" target language and "Contoso_English-German_Main" internal friendly name.
.EXAMPLE
New-FileBasedTM -Path "D:\Projects\TMs\Contoso Main.sdltm" -SrcLng "en-US" -TrgLng "de-DE fr-FR it-IT"

Creates "Contoso Main en-US_de-DE.sdltm", "Contoso Main en-US_fr-FR.sdltm" and "Contoso Main en-US_it-IT.sdltm" translation memories in "D:\Projects\TMs" location, with "English (United States)" source language and "German (Germany)", "French (France)" and "Italian (Italy)" respective target languages. Internal friendly TM names will be set to "Contoso Main en-US_de-DE", "Contoso Main en-US_fr-FR" and "Contoso Main en-US_it-IT" respectively.
#>
	[CmdletBinding(DefaultParametersetName="Location")]

	param(
		# Translation memory name. Must not contain invalid characters such as \ / : * ? " < > |
		[Parameter (ParametersetName="Location", Mandatory = $true)]
		[Alias("TMName")]
		[String] $Name,

		# Path to directory where the translation memory will be created.
		# If the directory does not exist, it will be created.
		[Parameter (ParametersetName="Location", Mandatory = $true)]
		[Alias("Location","TMLoc")]
		[String] $TMLocation,

		# Path of translation memory file (including the ".sdltm" extension!) to be created.
		[Parameter (ParametersetName="Path", Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path,

#		# Path of translation memory file (including the ".sdltm" extension!) to be used as 'template' on which the new TM will be based.
#		[Parameter (ParametersetName="Template")]
#		[Alias("Template","TMTemplate")]
#		[String] $From,

		# Locale code of translation memory source language.
		# For locale codes, see https://msdn.microsoft.com/en-us/goglobal/bb896001.aspx
		[Parameter (Mandatory = $true)]
		[Alias("SrcLng")]
		[String] $SourceLanguage,
		
		# Space-, comma- or semicolon-separated list of locale codes of translation memory target languages.
		# For locale codes, see https://msdn.microsoft.com/en-us/goglobal/bb896001.aspx
		[Parameter (Mandatory = $true)]
		[Alias("TrgLng","TargetLanguage")]
		[String] $TargetLanguages,

		# Optional translation memory description
		[Alias("TMDesc")]
		[String] $Description = "",

		# Set of fuzzy indexes that should be created in translation memory.
		[Sdl.LanguagePlatform.TranslationMemory.FuzzyIndexes] $FuzzyIndexes = $(Get-DefaultFuzzyIndexes),

		# Recognizer settings.
		[Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers] $Recognizers = $(Get-DefaultRecognizers),

		# Flags affecting tokenizer behavior for translation memory
		[Sdl.LanguagePlatform.Core.Tokenization.TokenizerFlags] $TokenizerFlags = $(Get-DefaultTokenizerFlags),

		# Flags affecting word count behavior for translation memory
		[Sdl.LanguagePlatform.TranslationMemory.WordCountFlags] $WordCountFlags = $(Get-DefaultWordCountFlags)
	)

	if ($TargetLanguages -ne $null -and $TargetLanguages -ne "") {
		# Parse target languages from provided parameter
		$TargetLanguagesList = $TargetLanguages -Split $LanguagesSeparator
	}

	# Loop through target languages and create package for each one
	$TargetLanguagesList | ForEach {
	
		$Language = $_
	
		$TMSourceLanguage = Get-CultureInfo $SourceLanguage
		$TMTargetLanguage = Get-CultureInfo $Language

		switch ($PsCmdlet.ParameterSetName) {
			"Path" {
				# Parse TM filename and path to separate variables
				$TMFileName = Split-Path -Path $Path -Leaf
				$TMLocation = Split-Path -Path $Path
				# If only filename was specified, i.e. path is empty, set path to current directory
				if ($TMLocation -eq "") {
					$TMLocation = "."
				}
			}
			"Location" {
				$TMFileName = $Name + $SDLTMFileExtension
			}
		}

		# If TM location does not exist, create it
		if (!(Test-Path -LiteralPath $TMLocation)) {
			New-Item $TMLocation -Force -ItemType Directory | Out-Null
		}

		# Construct full path to TM to be created
		if ($TargetLanguagesList.Length -eq 1) {
			# If only single target language was specified, use the TM filename as-is
			$TMPath = Join-Path $TMLocation $TMFileName
		}
		else {
			# If multiple target languages were specified, construct separate TM filename for each language
			$TMBaseFileName = [System.IO.Path]::GetFileNameWithoutExtension($TMFileName)
			$TMPath = Join-Path $TMLocation "$TMBaseFileName $($TMSourceLanguage.Name)_$($TMTargetLanguage.Name)$SDLTMFileExtension"
		}

		# Create TM
		if ($StudioVersion -le "Studio3") {
			$TM = New-Object Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory ($TMPath, $Description, $TMSourceLanguage, $TMTargetLanguage, $FuzzyIndexes, $Recognizers)
		}
		else {
			$TM = New-Object Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory ($TMPath, $Description, $TMSourceLanguage, $TMTargetLanguage, $FuzzyIndexes, $Recognizers, $TokenizerFlags, $WordCountFlags)
		}
	}
}

function Get-FileBasedTM {
	param(
		# Path of translation memory file (including the ".sdltm" extension!) to be opened.
		[Parameter (Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path
	)
	
	$TM = New-Object Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory ($Path)
	return $TM
}

function Get-ServerBasedTM {
	param(
		# Server URL
		[Parameter (Mandatory = $true)]
		[Alias("ServerUrl","Url","SrvUrl")]
		[String] $Server,
		
		# translation memory server path
		[Parameter (Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path,
		
		# server user login name
		[Parameter (Mandatory = $true)]
		[Alias("Login","User")]
		[String] $Username,
		
		# server user password
		[Parameter (Mandatory = $true)]
		[Alias("Pwd")]
		[String] $Password,
		
		[ValidateSet("None","LanguageDirections","Fields","LanguageResources","LanguageResourceData","Container","ScheduledOperations","All")]
		[Alias("TMProp")]
		[String] $TMProperties = "All"
	)
	
	# Cast TMProperties string value to corresponding enumeration value
	$AdditionalProperties = [Sdl.LanguagePlatform.TranslationMemoryApi.TranslationMemoryProperties] $TMProperties
	
	$TMServer = New-Object Sdl.LanguagePlatform.TranslationMemoryApi.TranslationProviderServer ($Server, $false, $Username, $Password)
	$TM = $TMServer.GetTranslationMemory($Path, $AdditionalProperties)
	return $TM
}

function Get-TMSourceLanguage {
	param(
		# Path of translation memory file (including the ".sdltm" extension!).
		[Parameter (ParameterSetName = "File", Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path,
		
		# Translation memory object.
		[Parameter (ParameterSetName = "Object", Mandatory = $true)]
		[Alias("TMObject")]
		[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory] $TM
	)
	
	if ($PsCmdlet.ParameterSetName -eq "File") {
		[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory] $TM = Get-FileBasedTM $Path
	}
	
	$Direction = $TM.LanguageDirection
	return $Direction.SourceLanguage
}

function Get-TMTargetLanguage {
	param(
		# Path of translation memory file (including the ".sdltm" extension!).
		[Parameter (ParameterSetName = "File", Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path,
		
		# Translation memory object.
		[Parameter (ParameterSetName = "Object", Mandatory = $true)]
		[Alias("TMObject")]
		[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory]$TM
	)
	
	if ($PsCmdlet.ParameterSetName -eq "File") {
		[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory]$TM = Get-FileBasedTM $Path
	}
	
	$Direction = $TM.LanguageDirection
	return $Direction.TargetLanguage
}

function Get-Language {
	param(
		[String] $Language
	)
	
	# temporary object used to get properly lower-/uppercased language code
	$tmp = New-Object Sdl.Core.Globalization.Language ($Language)

	# use the temporary object's language code to create actual language object
	# (e.g if lowercased "sr-latn-rs" parameter was passed to function, returned
	# language object is created using properly cased "sr-Latn-RS" parameter)
	$Lang = New-Object Sdl.Core.Globalization.Language ($tmp.CultureInfo.Name)
	return $Lang
}

function Get-Languages {
	param(
		[String[]] $Languages
	)

	[Sdl.Core.Globalization.Language[]]$Langs = @()
	foreach($Lang in $Languages)
	{
		$NewLang = Get-Language $Lang
		$Langs = $Langs + $newlang
	}

	return $Langs
}

function Get-CultureInfo {
	param(
		[String] $Language
	)

	$CultureInfo = Get-Language $Language
	return $CultureInfo.CultureInfo
}

function Get-DefaultFuzzyIndexes {
	 return [Sdl.LanguagePlatform.TranslationMemory.FuzzyIndexes]::SourceWordBased -bOr
			[Sdl.LanguagePlatform.TranslationMemory.FuzzyIndexes]::TargetWordBased
}

function Get-TMFuzzyIndexes {
	param(
		# Path of translation memory file (including the ".sdltm" extension!).
		[Parameter (ParameterSetName = "File", Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path,
		
		# Translation memory object.
		[Parameter (ParameterSetName = "Object", Mandatory = $true)]
		[Alias("TMObject")]
		[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory]$TM
	)
	
	if ($PsCmdlet.ParameterSetName -eq "File") {
		[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory]$TM = Get-FileBasedTM $Path
	}
	
	$FuzzyIndexes = $TM.FuzzyIndexes
	return $FuzzyIndexes
}

function Get-DefaultRecognizers {
<#{
	return [Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeDates -bAnd
	[Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeTimes -bAnd
	[Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeNumbers -bAnd
	[Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeAcronyms -bAnd
	[Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeVariables -bAnd
	[Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeMeasurements -bAnd
	[Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeAlphaNumeric
}#>
	return [Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeAll
}

function Get-TMRecognizers {
	param(
		# Path of translation memory file (including the ".sdltm" extension!).
		[Parameter (ParameterSetName = "File", Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path,
		
		# Translation memory object.
		[Parameter (ParameterSetName = "Object", Mandatory = $true)]
		[Alias("TMObject")]
		[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory]$TM
	)
	
	if ($PsCmdlet.ParameterSetName -eq "File") {
		[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory]$TM = Get-FileBasedTM $Path
	}
	
	$Recognizers = $TM.Recognizers
	return $Recognizers
}

function Get-DefaultTokenizerFlags {
	return [Sdl.LanguagePlatform.Core.Tokenization.TokenizerFlags]::DefaultFlags
}

function Get-TMTokenizerFlags {
	param(
		# Path of translation memory file (including the ".sdltm" extension!).
		[Parameter (ParameterSetName = "File", Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path,
		
		# Translation memory object.
		[Parameter (ParameterSetName = "Object", Mandatory = $true)]
		[Alias("TMObject")]
		[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory]$TM
	)
	
	if ($PsCmdlet.ParameterSetName -eq "File") {
		[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory]$TM = Get-FileBasedTM $Path
	}
	
	$TokenizerFlags = $TM.TokenizerFlags
	return $TokenizerFlags
}

function Get-DefaultWordCountFlags {
	return [Sdl.LanguagePlatform.TranslationMemory.WordCountFlags]::DefaultFlags
}

function Get-TMWordCountFlags {
	param(
		# Path of translation memory file (including the ".sdltm" extension!).
		[Parameter (ParameterSetName = "File", Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path,
		
		# Translation memory object.
		[Parameter (ParameterSetName = "Object", Mandatory = $true)]
		[Alias("TMObject")]
		[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory]$TM
	)
	
	if ($PsCmdlet.ParameterSetName -eq "File") {
		[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory]$TM = Get-FileBasedTM $Path
	}
	
	$WordcountFlags = $TM.WordCountFlags
	return $WordcountFlags
}

function Get-FilterString {
	param(
		# Path to *.sdltm.filters file containing filter definition(s) exported from Studio
		[parameter(Mandatory = $true)]
		[String]$FilterFilePath,
		
		# Name of the filter to be loaded from the provided filters file
		[parameter(Mandatory = $true)]
		[String]$FilterName
	)
	
	$FilterString = $null
	if ($FilterFilePath -ne $null -and $FilterFilePath -ne "") {
		$FilterFilePath = (Resolve-Path -LiteralPath $FilterFilePath).ProviderPath
		
		[xml]$Filters = Get-Content -Path $FilterFilePath
		$FilterString = ($Filters.TranslationMemoryFilters.Filter | Where-Object -Property Name -eq $FilterName).Expression
	}
	return $FilterString
}

function Import-TMX {
<#
.SYNOPSIS
Imports TMX file to Trados Studio TM, optionally applying a filter.
.DESCRIPTION
Imports content of TMX file to Trados Studio translation memory.
Optionally, a filter can be applied during import, allowing to import only translation units meeting certain criteria.
Filter is loaded from a filter file with ".sdltm.filters" extension, which can be obtained by exporting it from Trados Studio (from Translation Memories view using "Export filters to file" function).
.EXAMPLE
Import-TMX -Path "D:\Projects\TMs\EN-DE.sdltm" -TMXPath "D:\TMX\English-German.tmx"

Imports "D:\TMX\English-German.tmx" file to "D:\Projects\TMs\EN-DE.sdltm" translation memory.
.EXAMPLE
Import-TMX -Path "EN-DE.sdltm" -TMXPath "English-German.tmx"

Imports "English-German.tmx" file to "EN-DE.sdltm" translation memory. Both files are located in current folder.
.EXAMPLE
Import-TMX -Path "EN-DE.sdltm" -TMXPath "English-German.tmx" -FilterFile "D:\Projects\Filters\Microsoft.sdltm.filters" -FilterName "Word 2016"

Imports "English-German.tmx" file to "EN-DE.sdltm" translation memory. Both files are located in current folder.
A filter named "Word 2016" from a filter file "D:\Projects\Filters\Microsoft.sdltm.filters" will be applied during import.
Only translation units meeting criteria defined in the "Word 2016" filter will be actually imported into the translation memory.
#>
	param(
		# Path to translation memory file (including the ".sdltm" extension!).
		[Parameter (Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path,
		
		# Path of TMX file (including the ".tmx" extension!) to be imported.
		[Parameter (Mandatory = $true)]
		[String] $TMXPath,
		
		# Path to TM filter file, containing definition of filter to be used for import.
		# Only translation units matching the filter criteria will be imported.
		# (TM filter file can be created in Studio in Translation Memories view using "Export filters to file" function)
		[Alias("FltFile", "FltPath", "FilterPath")]
		[String] $FilterFile,
		
		# Name of the filter to be used for export.
		# Only translation units matching the filter criteria will be exported.
		# Filter must exist in the filter file specified using FilterFile parameter.
		[Alias("FltName")]
		[String] $FilterName
	)
	
	# Get filter string from the provided filter file
	$FilterString = Get-FilterString -FilterFilePath $FilterFile -FilterName $FilterName
	if ($FilterString -eq $null) {
		Write-Host "Filter name not found, importing complete TMX..." -ForegroundColor Yellow
	}
	
	# Create BatchImported event handler type
	$BatchImportedEventHandlerType = [System.Type] "System.EventHandler[Sdl.LanguagePlatform.TranslationMemoryApi.BatchImportedEventArgs]"
	
	# BatchImported event handler scriptblock
	$OnBatchImported = {
		param($sender, $e)
		$Stats = $e.Statistics
		$TotalRead = $Stats.TotalRead
		$TotalImported = $Stats.TotalImported
		$TotalAdded = $Stats.AddedTranslationUnits
		$TotalDiscarded = $Stats.DiscardedTranslationUnits
		$TotalMerged = $Stats.MergedTranslationUnits
		$TotalErrors = $Stats.Errors
		Write-Host "TUs processed: $TotalRead, imported: $TotalImported (added: $TotalAdded, merged: $TotalMerged), discarded: $TotalDiscarded, errors: $TotalErrors`r" -NoNewLine
	} -as $BatchImportedEventHandlerType
	
	# Get full TM path
	if ($Path -ne $null -and $Path -ne "") {
		$Path = (Resolve-Path -LiteralPath $Path).ProviderPath
	}
	
	# Get full TMX path
	if ($TMXPath -ne $null -and $TMXPath -ne "") {
		$TMXPath = (Resolve-Path -LiteralPath $TMXPath).ProviderPath
	}
	
	# Display info about import source and target
	Write-Host "$(Split-Path $TMXPath -Leaf) -> $(Split-Path $Path -Leaf)" -ForegroundColor White
	
	# Get the source TM object and create importer object
	$TM = Get-FileBasedTM $Path
	$Importer = New-Object Sdl.LanguagePlatform.TranslationMemoryApi.TranslationMemoryImporter ($TM.LanguageDirection)
	
	# Parse the filter string and set the importer filter expression
	# TODO: error handling (e.g. when filter uses fields which do not exist in the TM)
	if ($FilterString -ne $null) {
		$FilterExpr = [Sdl.LanguagePlatform.TranslationMemory.FilterExpressionParser]::Parse($FilterString, $TM.FieldDefinitions)
		$Importer.ImportSettings.Filter = $FilterExpr
	}
	
	# Register import event handler, do the import and unregister event handler afterwards
	$Importer.add_BatchImported($OnBatchImported)
	$Importer.Import($TMXPath)
	$Importer.remove_BatchImported($OnBatchImported)
	
	# when all is done, output nothing WITH NEW LINE
	# so that the last progress output from event handler is kept
	Write-Host $null
}

function Export-TMX {
<#
.SYNOPSIS
Exports Trados Studio translation memory to TMX file, optionally applying a filter.
.DESCRIPTION
Exports one or more Trados Studio translation memories to TMX file(s).
Optionally, a filter can be applied during export, allowing to export only translation units meeting certain criteria.
Filter is loaded from a filter file with ".sdltm.filters" extension, which can be obtained by exporting it from Trados Studio (from Translation Memories view using "Export filters to file" function).
.EXAMPLE
Export-TMX

Exports all Trados Studio translation memories found in current directory to TMX files. Existing TMX files will be preserved.
.EXAMPLE
Export-TMX -TMLocation "D:\Projects\TMs" -TMXLocation "D:\TMX_Exports" -Recurse -Force

Exports to TMX all Trados Studio TMs present in "D:\Projects\TMs" folder and its subfolders. Exported files will be stored to "D:\TMX_Exports" folder. Existing TMX files will be overwritten.
.EXAMPLE
Export-TMX -TMLocation "D:\Projects\TMs\EN-DE.sdltm" -FilterFile "D:\Projects\Filters\Microsoft.sdltm.filters" -FilterName "Word 2016"

Exports a single "D:\Projects\TMs\EN-DE.sdltm" translation memory to TMX. Exported file will be stored in the same location as source file (i.e. "D:\Projects\TMs").
A filter named "Word 2016" from a filter file "D:\Projects\Filters\Microsoft.sdltm.filters" will be applied during export.
Exported TMX will then contain only translation units meeting criteria defined in the "Word 2016" filter.
#>
	param(
		# Path to either single TM file, or to directory where one or more TMs are located.
		# If this parameter is omitted, current directory is used by default.
		[Alias("Location","TMLoc")]
		[String] $TMLocation = ".",
		
		# Path to directory where the exported TMX will be created.
		# If this parameter is omitted, exported TMX will be created in the same location as source TM.
		[Alias("TMXLoc")]
		[String] $TMXLocation,
		
		# Path to TM filter file, containing definition of filter to be used for export.
		# Only translation units matching the filter criteria will be exported.
		# (TM filter file can be created in Studio in Translation Memories view using "Export filters to file" function)
		[Alias("FltFile","FltPath","FilterPath")]
		[String] $FilterFile,
		
		# Name of the filter to be used for export.
		# Only translation units matching the filter criteria will be exported.
		# Filter must exist in the filter file specified using FilterFile parameter.
		[Alias("FltName")]
		[String] $FilterName,
		
		# Allows to overwrite any existing file
		[Alias("Overwrite")]
		[Switch] $Force,
		
		# Exports also all TMs found in subdirectories of the specified path.
		[Alias("r")]
		[Switch] $Recurse
	)
	
	# Get filter string from the provided filter file
	$FilterString = Get-FilterString -FilterFilePath $FilterFile -FilterName $FilterName
	if ($FilterString -eq $null) {
		Write-Host "Filter name not found, exporting complete TM..." -ForegroundColor Yellow
	}
	
	# Create BatchExported event handler type
	$BatchExportedEventHandlerType = [System.Type] "System.EventHandler[Sdl.LanguagePlatform.TranslationMemoryApi.BatchExportedEventArgs]"
	
	# BatchExported event handler scriptblock
	$OnBatchExported = {
		param([System.Object]$sender, [Sdl.LanguagePlatform.TranslationMemoryApi.BatchExportedEventArgs]$e)
		$TotalProcessed = $e.TotalProcessed
		$TotalExported = $e.TotalExported
		Write-Host "TUs processed: $TotalProcessed, exported: $TotalExported`r" -NoNewLine
	} -as $BatchExportedEventHandlerType
	
	# Get full TMX location path
	if ($TMXLocation -ne $null -and $TMXLocation -ne "") {
		$TMXLocation = (Resolve-Path -LiteralPath $TMXLocation).ProviderPath
	}
	
	# Get full TM location path and iterate over all *.sdltm files in the location
	# (if single file is specified, process only that file)
	$TMLocation = (Resolve-Path -LiteralPath $TMLocation).ProviderPath
	Get-ChildItem -Path $TMLocation -Filter "*$SDLTMFileExtension" -File -Recurse:$Recurse | ForEach-Object {
		$SDLTM = $_
		$TMXName = $SDLTM.Name.Replace($SDLTMFileExtension, $TMXFileExtension)
		
		# If TMX export location was not specified, use the location of source SDLTM file,
		# otherwise use the specified location
		if ($TMXLocation -eq "") {
			$TMXPath = $SDLTM.DirectoryName
		}
		else {
			$TMXPath = $TMXLocation
		}
		
		# Get the source TM object and create exporter object
		$TM = Get-FileBasedTM $SDLTM.FullName
		$Exporter = New-Object Sdl.LanguagePlatform.TranslationMemoryApi.TranslationMemoryExporter $TM.LanguageDirection
		
		# Parse the filter string and set the exporter filter expression
		# TODO: error handling (e.g. when filter uses fields which do not exist in the TM)
		if ($FilterString -ne $null) {
			$FilterExpr = [Sdl.LanguagePlatform.TranslationMemory.FilterExpressionParser]::Parse($FilterString, $TM.FieldDefinitions)
			$Exporter.FilterExpression = $FilterExpr
		}
		
		# Display info about export source and target
		Write-Host "$($SDLTM.Name) -> $TMXName" -ForegroundColor White
		
		# Register export event handler, do the export and unregister event handler afterwards
		$Exporter.add_BatchExported($OnBatchExported)
		$Exporter.Export("$TMXPath\$TMXName", ($Force.IsPresent))
		$Exporter.remove_BatchExported($OnBatchExported)
		
		# when all is done, output nothing WITH NEW LINE
		# so that the last progress output from event handler is kept
		Write-Host $null
	}
}

Export-ModuleMember New-FileBasedTM
Export-ModuleMember Get-FileBasedTM
Export-ModuleMember Get-ServerBasedTM
Export-ModuleMember Get-DefaultFuzzyIndexes
Export-ModuleMember Get-DefaultRecognizers
Export-ModuleMember Get-DefaultTokenizerFlags
Export-ModuleMember Get-DefaultWordCountFlags
Export-ModuleMember Get-Language
Export-ModuleMember Get-Languages
Export-ModuleMember Get-TMSourceLanguage
Export-ModuleMember Get-TMTargetLanguage
Export-ModuleMember Get-TMFuzzyIndexes
Export-ModuleMember Get-TMRecognizers
Export-ModuleMember Get-TMTokenizerFlags
Export-ModuleMember Get-TMWordCountFlags
Export-ModuleMember Import-TMX
Export-ModuleMember Export-TMX
