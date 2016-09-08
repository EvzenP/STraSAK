﻿function Get-Guids
{
	param([Sdl.ProjectAutomation.Core.ProjectFile[]] $files)
	[System.Guid[]] $guids = New-Object System.Guid[] ($files.Count);
	$i = 0;
	foreach($file in $files)
	{
		$guids.Set($i,$file.Id);
		$i++;
	}
	return $guids
}
 
Export-ModuleMember Get-Guids 

