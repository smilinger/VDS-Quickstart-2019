function InitializeFileNameValidation()
{
    if ($Prop["_EditMode"].Value -ne $true)
    {
	    if($dsWindow.Name -eq 'InventorWindow')
	    {
			#new QS 2019
			#generate file name preview for validation. Note-remove the extension during OnPostClose!
			if($Prop["_SaveCopyAsMode"].Value -eq $true)
			{
				$dsWindow.FindName("Format").add_SelectionChanged({
					mPreviewExportFileName
				})
				$dsWindow.FindName("NumSchms").add_SelectionChanged({
					mPreviewExportFileName
				})
				$dsWindow.FindName("FILENAME").add_TextChanged({
					mPreviewExportFileName
				})
			}


            $Prop["DocNumber"].CustomValidation = { FileNameCustomValidation } #overrides any validation if set; the type is Scriptblock {} containing functions or direct calls; results will drive the "ValidProperties/Invalid"
		}
        elseif($dsWindow.Name -eq 'AutoCADWindow')
		{
		    $Prop["GEN-TITLE-DWG"].CustomValidation = { FileNameCustomValidation }
		}                
    }
}

function FileNameCustomValidation
{
    $DSNumSchmsCtrl = $dsWindow.FindName("DSNumSchmsCtrl")
    if ($DSNumSchmsCtrl -and -not $DSNumSchmsCtrl.NumSchmFieldsEmpty)
    {
        return $true
    }
    if($dsWindow.Name -eq 'InventorWindow')
	{
        $propertyName = "DocNumber"
	}
    elseif($dsWindow.Name -eq 'AutoCADWindow')
	{
		$propertyName = "GEN-TITLE-DWG"
	}
	
	
	if($Prop["_SaveCopyAsMode"].Value -eq $true) #validate the preview file name 
	{
		#$_temp = $Prop["DocNumber"].Value
		#$dsDiag.Trace("PreviewExport File Validation $_temp")
		$rootFolder = GetVaultRootFolder
		$fullFileName = [System.IO.Path]::Combine($Prop["_FilePath"].Value, $Prop["DocNumber"].Value)
		if ([System.IO.File]::Exists($fullFileName))
		{
			$Prop["$($propertyName)"].CustomValidationErrorMessage = "$($UIString["MSG4"])" #tooltip
			return $false
		}
		$isinvault = FileExistsInVault($rootFolder.FullName + "/" + $Prop["Folder"].Value.Replace(".\", "") + $Prop["DocNumber"].Value)
		if ($isinvault)
		{
			$Prop["$($propertyName)"].CustomValidationErrorMessage = "$($UIString["VAL12"])"
			return $false
		}
		if ($vault.DocumentService.GetUniqueFileNameRequired())
		{    
			$result = FindFile -fileName $Prop["DocNumber"].Value
			if ($result)
			{
				$Prop["$($propertyName)"].CustomValidationErrorMessage = "$($UIString["VAL13"])"
				return $false
			}
		}
		return $true
	} #end SaveCopyAsMode
    
	#$dsDiag.Trace("general validation")
	$rootFolder = GetVaultRootFolder
    $fullFileName = [System.IO.Path]::Combine($Prop["_FilePath"].Value, $Prop["_FileName"].Value)
    if ([System.IO.File]::Exists($fullFileName))
    {
		$Prop["$($propertyName)"].CustomValidationErrorMessage = "$($UIString["MSG4"])" #tooltip
        return $false
    }
    $isinvault = FileExistsInVault($rootFolder.FullName + "/" + $Prop["Folder"].Value.Replace(".\", "") + $Prop["_FileName"].Value)
    if ($isinvault)
    {
		$Prop["$($propertyName)"].CustomValidationErrorMessage = "$($UIString["VAL12"])"
        return $false
    }
    if ($vault.DocumentService.GetUniqueFileNameRequired())
    {    
        $result = FindFile -fileName $Prop["_FileName"].Value
        if ($result)
        {
			$Prop["$($propertyName)"].CustomValidationErrorMessage = "$($UIString["VAL13"])"
            return $false
        }
    }
    return $true
}

function FindFile($fileName)
{
    $filePropDefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE")
    $fileNamePropDef = $filePropDefs | where {$_.SysName -eq "ClientFileName"}
    $srchCond = New-Object 'Autodesk.Connectivity.WebServices.SrchCond'
    $srchCond.PropDefId = $fileNamePropDef.Id
    $srchCond.PropTyp = "SingleProperty"
    $srchCond.SrchOper = 3 #is equal
    $srchCond.SrchRule = "Must"
    $srchCond.SrchTxt = $fileName

    $bookmark = ""
    $status = $null
    $totalResults = @()
    while ($status -eq $null -or $totalResults.Count -lt $status.TotalHits)
    {
        $results = $vault.DocumentService.FindFilesBySearchConditions(@($srchCond),$null, $null, $false, $true, [ref]$bookmark, [ref]$status)

        if ($results -ne $null)
        {
            $totalResults += $results
        }
        else {break}
    }
    return $totalResults;
}

function FileExistsInVault($vaultPath)
{
    $pathWithoutDot = $vaultPath.Replace("/.", "/")
    $pathInVaultFormat = $pathWithoutDot.Replace("\", "/")
    try
    {
        $files = $vault.DocumentService.FindLatestFilesByPaths(@($pathInVaultFormat))
        if ($files.Count -gt 0)
        {
            if ($files[0].Id -ne -1)
            { return $true }
        }
    }
    catch
    {
        #$dsDiag.Inspect()
    }    
    return $false
}

#new QS 2019
function mPreviewExportFileName()
{
	If ($Global:OnPostAction)
	{
		return #we must not add an extension automatically, as soon as the dialog shuts down; otherwise the final name gets 2 extensions
	}
	If($Prop["_Format"].Value -ne ".jt")
	{
		$newFileName = @()
		$newFileName += ($Prop["DocNumber"].Value.Split("."))
		If($newFileName.Count -gt 1) 
		{
			$newExt = $newFileName[$newFileName.Count-1]
			$Prop["DocNumber"].Value = $Prop["DocNumber"].Value.Replace("." + $newExt, ".stp")
		}
		Else 
		{ 
			$newExt =  ".stp" 
			$Prop["DocNumber"].Value = $Prop["DocNumber"].Value + $newExt
		}
	}
	If($Prop["_Format"].Value -eq ".jt")
	{
		$newFileName += ($Prop["DocNumber"].Value.Split("."))
		If($newFileName.Count -gt 1) 
		{
			$newExt = $newFileName[$newFileName.Count-1]
			$Prop["DocNumber"].Value = $Prop["DocNumber"].Value.Replace("." + $newExt, ".jt")
		}
		Else 
		{ 
			$newExt =  ".jt" 
			$Prop["DocNumber"].Value = $Prop["DocNumber"].Value + $newExt
		}
	}
}
