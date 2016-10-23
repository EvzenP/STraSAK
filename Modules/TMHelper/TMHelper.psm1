param([String]$StudioVersion = "Studio4")

if ("${Env:ProgramFiles(x86)}")
{
    $ProgramFilesDir = "${Env:ProgramFiles(x86)}"
}
else
{
    $ProgramFilesDir = "${Env:ProgramFiles}"
}

Add-Type -Path "$ProgramFilesDir\SDL\SDL Trados Studio\$StudioVersion\Sdl.LanguagePlatform.TranslationMemoryApi.dll"
Add-Type -Path "$ProgramFilesDir\SDL\SDL Trados Studio\$StudioVersion\Sdl.LanguagePlatform.TranslationMemory.dll"

function New-FileBasedTM
{
<#
.DESCRIPTION
Creates a new file based TM.
#>
	param([String] $filePath,[String] $description, [String] $sourceLanguageName, [String] $targetLanguageName,
		[Sdl.LanguagePlatform.TranslationMemory.FuzzyIndexes] $fuzzyIndexes,
		[Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers] $recognizers,
		[Sdl.LanguagePlatform.Core.Tokenization.TokenizerFlags] $tokenizerFlags,
		[Sdl.LanguagePlatform.TranslationMemory.WordCountFlags] $wordcountFlags)


	$sourceLanguage = Get-CultureInfo $sourceLanguageName
	$targetLanguage = Get-CultureInfo $targetLanguageName

	if ($StudioVersion -le "Studio3")
	{
		[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory] $tm =
		New-Object Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory ($filePath,
		$description, $sourceLanguage, $targetLanguage, $fuzzyIndexes, $recognizers)
	}
	else
	{
		[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory] $tm =
		New-Object Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory ($filePath,
		$description, $sourceLanguage, $targetLanguage, $fuzzyIndexes, $recognizers, $tokenizerFlags, $wordcountFlags)
	}
}

function Open-FileBasedTM
{
	param([String] $filePath)
	[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory] $tm =
	New-Object Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory ($filePath)

	return $tm
}

function Get-TMSourceLanguage
{
	param([String] $filePath)

	[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory] $tm = Open-FileBasedTM $filePath
	[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemoryLanguageDirection] $direction = $tm.LanguageDirection
	return $direction.SourceLanguage
}

function Get-TMTargetLanguage
{
	param([String] $filePath)

	[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory] $tm = Open-FileBasedTM $filePath
	[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemoryLanguageDirection] $direction = $tm.LanguageDirection
	return $direction.TargetLanguage
}

function Get-Language
{
	param([String] $languageName)

	[Sdl.Core.Globalization.Language] $language = New-Object Sdl.Core.Globalization.Language ($languageName)
	return $language
}

function Get-Languages
{
	param([String[]] $languageNames)
	[Sdl.Core.Globalization.Language[]]$languages = @()
	foreach($lang in $languageNames)
	{
		$newlang = Get-Language $lang

		$languages = $languages + $newlang
	}

	return $languages
}

function Get-CultureInfo
{
	param([String] $languageName)
	$cultureInfo = Get-Language $languageName
	return [System.Globalization.CultureInfo] $cultureInfo.CultureInfo
}

function Get-DefaultFuzzyIndexes
{
	 return [Sdl.LanguagePlatform.TranslationMemory.FuzzyIndexes]::SourceCharacterBased -bAnd
	 	[Sdl.LanguagePlatform.TranslationMemory.FuzzyIndexes]::SourceWordBased -bAnd
		[Sdl.LanguagePlatform.TranslationMemory.FuzzyIndexes]::TargetCharacterBased -bAnd
		[Sdl.LanguagePlatform.TranslationMemory.FuzzyIndexes]::TargetWordBased
}

function Get-TMFuzzyIndexes
{
	param([String] $filePath)

	[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory] $tm = Open-FileBasedTM $filePath
	[Sdl.LanguagePlatform.TranslationMemoryApi.FuzzyIndexes] $fuzzyIndexes = $tm.FuzzyIndexes
	return $fuzzyIndexes
}

#function Get-DefaultRecognizers
<#{
	return [Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeDates -bAnd
	[Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeTimes -bAnd
	[Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeNumbers -bAnd
	[Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeAcronyms -bAnd
	[Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeVariables -bAnd
	[Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeMeasurements -bAnd
	[Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeAlphaNumeric
}#>
{
	return [Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers]::RecognizeAll
}

function Get-TMRecognizers
{
	param([String] $filePath)

	[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory] $tm = Open-FileBasedTM $filePath
	[Sdl.LanguagePlatform.Core.Tokenization.BuiltinRecognizers] $recognizers = $tm.Recognizers
	return $recognizers
}

function Get-DefaultTokenizerFlags
{
	return [Sdl.LanguagePlatform.Core.Tokenization.TokenizerFlags]::DefaultFlags
}

function Get-TMTokenizerFlags
{
	param([String] $filePath)

	[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory] $tm = Open-FileBasedTM $filePath
	[Sdl.LanguagePlatform.Core.Tokenization.TokenizerFlags] $tokenizerFlags = $tm.TokenizerFlags
	return $tokenizerFlags
}

function Get-DefaultWordCountFlags
{
	return [Sdl.LanguagePlatform.TranslationMemory.WordCountFlags]::DefaultFlags
}

function Get-TMWordCountFlags
{
	param([String] $filePath)

	[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory] $tm = Open-FileBasedTM $filePath
	[Sdl.LanguagePlatform.TranslationMemory.WordCountFlags] $wordcountFlags = $tm.WordCountFlags
	return $wordcountFlags
}

function Import-TMX
{
	param([String] $tmFilePath, [String] $importFilePath)
	[Sdl.LanguagePlatform.TranslationMemoryApi.FileBasedTranslationMemory] $tm = Open-FileBasedTM $tmFilePath
	[Sdl.LanguagePlatform.TranslationMemoryApi.TranslationMemoryImporter] $importer =
	New-Object Sdl.LanguagePlatform.TranslationMemoryApi.TranslationMemoryImporter ($tm.LanguageDirection)
	$importer.Import($importFilePath)
}

Export-ModuleMember New-FileBasedTM
Export-ModuleMember Open-FileBasedTM
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
