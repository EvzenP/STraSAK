param ([String] $studioVersion = "Studio18")

# Helper module that loads all the others and resolve the versioning conflicts between powershell versions and Trados versions
$scriptPath = $MyInvocation.MyCommand.Path
$scriptParentDiv = Split-Path $scriptPath -Parent;
$moduleNames = @("GetGuids", "PackageHelper", "ProjectHelper", "TMHelper")

Add-Dependencies $studioVersion;

foreach ($moduleName in $moduleNames)
{
    Import-Module -Name $moduleName -Scope Global -ArgumentList $studioVersion
}

function Add-Dependencies {
    param([String] $StudioVersion)

    $assemblyResolverPath = $scriptParentDiv + "\DependencyResolver.dll"
    $versionNumber = [regex]::Match($StudioVersion, "\d+").Value;

    if ("${Env:ProgramFiles(x86)}") {
        $ProgramFilesDir = "${Env:ProgramFiles(x86)}"
    }
    else {
        $ProgramFilesDir = "${Env:ProgramFiles}"
    }

    if ($versionNumber -le 16)
    {
        $appPath = "$ProgramFilesDir\Sdl\Sdl Trados Studio\$StudioVersion\"
    }
    else {
        $appPath = "$ProgramFilesDir\Trados\Trados Studio\$StudioVersion\"
    }

    # Solve dependency conficts
    Add-Type -Path $assemblyResolverPath;
    $assemblyResolver = New-Object DependencyResolver.AssemblyResolver("$appPath\");
    $assemblyResolver.Resolve();

    Add-Type -Path "$appPath\Sdl.ProjectAutomation.Core.dll"
    Add-Type -Path "$appPath\Sdl.ProjectAutomation.FileBased.dll"
    Add-Type -Path "$appPath\Sdl.ProjectAutomation.Settings.dll"
    Add-Type -Path "$appPath\Sdl.LanguagePlatform.TranslationMemory.dll"
}