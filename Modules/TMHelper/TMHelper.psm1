param([String]$StudioVersion = "Studio4")

if ("${Env:ProgramFiles(x86)}") {
    $ProgramFilesDir = "${Env:ProgramFiles(x86)}"
}
else {
    $ProgramFilesDir = "${Env:ProgramFiles}"
}

Add-Type -Path "$ProgramFilesDir\SDL\SDL Trados Studio\$StudioVersion\Sdl.LanguagePlatform.TranslationMemoryApi.dll"
Add-Type -Path "$ProgramFilesDir\SDL\SDL Trados Studio\$StudioVersion\Sdl.LanguagePlatform.TranslationMemory.dll"

function New-FileBasedTM {
<#
.SYNOPSIS
Creates a new file based TM.
.DESCRIPTION
Creates new file based translation memory using specified source and target language.
Optionally also a TM description and further TM behavior options can be specified.
.EXAMPLE
New-FileBasedTM -Name "Contoso English-German Main" -TMLocation "D:\Projects\TMs" -SourceLanguage "en-US" -TargetLanguage "de-DE"

Creates "D:\Projects\TMs\Contoso English-German Main.sdltm" translation memory with "English (United States)" source language and "German (Germany)" target language and "Contoso English-German Main" internal friendly name.
.EXAMPLE
New-FileBasedTM -Path "D:\Projects\TMs\Contoso_English-German_Main.sdltm" -SrcLng "en-US" -TrgLng "de-DE"

Creates "D:\Projects\TMs\Contoso_English-German_Main.sdltm" translation memory with "English (United States)" source language and "German (Germany)" target language and "Contoso_English-German_Main" internal friendly name.
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
#		[Alias("Template")]
#		[String] $From,

		# Locale code of translation memory source language.
		# For locale codes, see https://msdn.microsoft.com/en-us/goglobal/bb896001.aspx
		[Parameter (Mandatory = $true)]
		[Alias("SrcLng")]
		[String] $SourceLanguage,
		
		# Locale code of translation memory target language.
		# For locale codes, see https://msdn.microsoft.com/en-us/goglobal/bb896001.aspx
		[Parameter (Mandatory = $true)]
		[Alias("TrgLng")]
		[String] $TargetLanguage,

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

	$TMSourceLanguage = Get-CultureInfo $SourceLanguage
	$TMTargetLanguage = Get-CultureInfo $TargetLanguage

	switch ($PsCmdlet.ParameterSetName) {
		"Path" {
			$TMLocation = Split-Path -Path $Path
			if ($TMLocation -eq "") {$TMLocation = "."}
			$TMFileName = Split-Path -Path $Path -Leaf
		}
		"Location" {
			$TMFileName = $Name + ".sdltm"
		}
	}

	# If TM location does not exist, create it
	if (!(Test-Path -LiteralPath $TMLocation)) {
		New-Item $TMLocation -Force -ItemType Directory | Out-Null
	}

	$TMPath = Join-Path $TMLocation $TMFileName

	if ($StudioVersion -le "Studio3") {
		$TM = New-Object Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory ($TMPath, $Description, $TMSourceLanguage, $TMTargetLanguage, $FuzzyIndexes, $Recognizers)
	}
	else {
		$TM = New-Object Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory ($TMPath, $Description, $TMSourceLanguage, $TMTargetLanguage, $FuzzyIndexes, $Recognizers, $TokenizerFlags, $WordCountFlags)
	}
}

function Get-FilebasedTM {
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
		[Parameter (Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path
	)
	
	$TM = Get-FilebasedTM $Path
	$Direction = $TM.LanguageDirection
	return $Direction.SourceLanguage
}

function Get-TMTargetLanguage {
	param(
		# Path of translation memory file (including the ".sdltm" extension!).
		[Parameter (Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path
	)

	$TM = Get-FilebasedTM $Path
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
		[Parameter (Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path
	)

	$TM = Get-FilebasedTM $Path
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
		[Parameter (Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path
	)

	$TM = Get-FilebasedTM $Path
	$Recognizers = $TM.Recognizers
	return $Recognizers
}

function Get-DefaultTokenizerFlags {
	return [Sdl.LanguagePlatform.Core.Tokenization.TokenizerFlags]::DefaultFlags
}

function Get-TMTokenizerFlags {
	param(
		# Path of translation memory file (including the ".sdltm" extension!).
		[Parameter (Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path
	)

	$TM = Get-FilebasedTM $Path
	$TokenizerFlags = $TM.TokenizerFlags
	return $TokenizerFlags
}

function Get-DefaultWordCountFlags {
	return [Sdl.LanguagePlatform.TranslationMemory.WordCountFlags]::DefaultFlags
}

function Get-TMWordCountFlags {
	param(
		# Path of translation memory file (including the ".sdltm" extension!).
		[Parameter (Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path
	)

	$TM = Get-FilebasedTM $Path
	$WordcountFlags = $TM.WordCountFlags
	return $WordcountFlags
}

function Import-TMX {
<#
.SYNOPSIS
Imports TMX file to Trados Studio TM.
.DESCRIPTION
Imports content of TMX file to Trados Studio translation memory.
.EXAMPLE
Import-TMX -Path "D:\Projects\TMs\EN-DE.sdltm" -TMXPath "D:\TMX\English-German.tmx"

Imports "D:\TMX\English-German.tmx" file to "D:\Projects\TMs\EN-DE.sdltm" translation memory.
.EXAMPLE
Import-TMX -Path "EN-DE.sdltm" -TMXPath "English-German.tmx"

Imports "English-German.tmx" file to "EN-DE.sdltm" translation memory. Both files are located in current folder.
#>
	param(
		# Path to translation memory file (including the ".sdltm" extension!).
		[Parameter (Mandatory = $true)]
		[Alias("TMPath")]
		[String] $Path,
		
		# Path of TMX file (including the ".tmx" extension!) to be imported.
		[Parameter (Mandatory = $true)]
		[String] $TMXPath
	)

	# Event handler scriptblock
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
	}
	
	$Path = (Resolve-Path -LiteralPath $Path).ProviderPath
	$TMXPath = (Resolve-Path -LiteralPath $TMXPath).ProviderPath
	
	Write-Host "$(Split-Path $TMXPath -Leaf) -> $(Split-Path $Path -Leaf)" -ForegroundColor White
	
	$TM = Get-FilebasedTM $Path
	$Importer = New-Object Sdl.LanguagePlatform.TranslationMemoryApi.TranslationMemoryImporter ($TM.LanguageDirection)
	$Importer.Add_BatchImported($OnBatchImported)
	$Importer.Import($TMXPath)
	$Importer.Remove_BatchImported($OnBatchImported)
	
	# when all is done, output nothing WITH NEW LINE
	Write-Host $null
}

function Export-TMX {
<#
.SYNOPSIS
Exports Trados Studio translation memory to TMX file.
.DESCRIPTION
Exports one or more Trados Studio translation memories to TMX file(s).
.EXAMPLE
Export-TMX

Exports all Trados Studio translation memories found in current directory to TMX files. Existing TMX files will be preserved.
.EXAMPLE
Export-TMX -TMLocation "D:\Projects\TMs" -TMXLocation "D:\TMX_Exports" -Recurse -Force

Exports to TMX all Trados Studio TMs present in "D:\Projects\TMs" folder and its subfolders. Exported files will be stored to "D:\TMX_Exports" folder. Existing TMX files will be overwritten.
.EXAMPLE
Export-TMX -TMLocation "D:\Projects\TMs\EN-DE.sdltm"

Exports a single "D:\Projects\TMs\EN-DE.sdltm" translation memory to TMX. Exported file will be stored in the same location as source file (i.e. "D:\Projects\TMs").
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
		
		# Allows to overwrite any existing file
		[Alias("Overwrite")]
		[Switch] $Force,
		
		# Exports also all TMs found in subdirectories of the specified path.
		[Alias("r")]
		[Switch] $Recurse
	)

	$TMLocation = (Resolve-Path -LiteralPath $TMLocation).ProviderPath
	
	if ($TMXLocation -ne "") {
		$TMXLocation = (Resolve-Path -LiteralPath $TMXLocation).ProviderPath
	}
	
	# Event handler scriptblock
	$OnBatchExported = {
		param($sender, $e)
		$TotalProcessed = $e.TotalProcessed
		$TotalExported = $e.TotalExported
		Write-Host "TUs processed: $TotalProcessed, exported: $TotalExported`r" -NoNewLine
	}
	
	Get-ChildItem $TMLocation *.sdltm -File -Recurse:$Recurse | ForEach-Object {
		$SDLTM = $_
		$TMXName = $SDLTM.Name.Replace(".sdltm", ".tmx")
		
		if ($TMXLocation -eq "") {
			$TMXPath = $SDLTM.DirectoryName
		}
		else {
			$TMXPath = $TMXLocation
		}
		
		Write-Host "$($SDLTM.Name) -> $TMXName" -ForegroundColor White
		
		$TM = Get-FilebasedTM $SDLTM.FullName
		$Exporter = New-Object Sdl.LanguagePlatform.TranslationMemoryApi.TranslationMemoryExporter ($TM.LanguageDirection)
		$Exporter.Add_BatchExported($OnBatchExported)
		$Exporter.Export("$TMXPath\$TMXName", ($Force.IsPresent))
		$Exporter.Remove_BatchExported($OnBatchExported)
		
		# when all is done, output nothing WITH NEW LINE
		Write-Host $null
	}
}

Export-ModuleMember New-FileBasedTM
Export-ModuleMember Get-FilebasedTM
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
