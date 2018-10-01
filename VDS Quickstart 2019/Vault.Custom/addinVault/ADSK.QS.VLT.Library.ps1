#region disclaimer
#=============================================================================#
# PowerShell script sample for Vault Data Standard Quickstart Configuration   #
#                                                                             #
# Copyright (c) Autodesk - All rights reserved.                               #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#
#endregion

#region - version history
# Version Info - VDS Quickstart Vault Library 2019.1.1
	# fixed failure in getting PropertyTranslations for default DSLanguages settings
	# added mGetProjectFolderPropToVaultFile

# Version Info - VDS Quickstart Vault Library 2019.1.0
	# new name, aligned to Quickstart naming convention and added CAD extension library

# Version Info - VDS Quickstart Extension 2019.1.0
	# new function mSearchCustentOfCat([String]$mCatDispName)
	# new function mSearchItemOfCat([String]$mCatDispName)

# Version Info - VDS Quickstart Extension 2019.0.0
	# new function ADSK.GroupMemberOf()

#endregion

#retrieve property value given by displayname from folder (ID)
function mGetFolderPropValue ([Int64] $mFldID, [STRING] $mDispName)
{
	$PropDefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("FLDR")
	$propDefIds = @()
	$PropDefs | ForEach-Object {
		$propDefIds += $_.Id
	} 
	$mPropDef = $propDefs | Where-Object { $_.DispName -eq $mDispName}
	$mEntIDs = @()
	$mEntIDs += $mFldID
	$mPropDefIDs = @()
	$mPropDefIDs += $mPropDef.Id
	$mProp = $vault.PropertyService.GetProperties("FLDR",$mEntIDs, $mPropDefIDs)
	$mProp | Where-Object { $mPropVal = $_.Val }
	Return $mPropVal
}

#retrieve the definition ID for given property by displayname
function mGetFolderPropertyDefId ([STRING] $mDispName) {
	$PropDefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("FLDR")
	$propDefIds = @()
	$PropDefs | ForEach-Object {
		$propDefIds += $_.Id
	} 
	$mPropDef = $propDefs | Where-Object { $_.DispName -eq $mDispName}
	Return $mPropDef.Id
}

#retrieve Category definition ID by display name
function mGetCategoryDef ([String] $mEntType, [String] $mDispName)
{
	$mEntityCategories = $vault.CategoryService.GetCategoriesByEntityClassId($mEntTyp, $true)
	$mEntCatId = ($mEntityCategories | Where-Object {$_.Name -eq $mDispName }).ID
	return $mEntCatId
}

#update single property. Parameters: Folder ID, UDP display name and UDP value
function mUpdateFldrProperties([Long] $FldId, [String] $mDispName, [Object] $mVal)
{
	$ent_idsArray = @()
	$ent_idsArray += $FldId
	$propInstParam = New-Object Autodesk.Connectivity.WebServices.PropInstParam
	$propInstParamArray = New-Object Autodesk.Connectivity.WebServices.PropInstParamArray
	$mPropDefId = mGetFolderPropertyDefId $mDispName
 	$propInstParam.PropDefId = $mPropDefId
	$propInstParam.Val = $mVal
	$propInstParamArray.Items += $propInstParam
	$propInstParamArrayArray += $propInstParamArray
	Try{
        $vault.DocumentServiceExtensions.UpdateFolderProperties($ent_idsArray, $propInstParamArrayArray)
	    return $true
    }
    catch { return $false}
}

#show current runspace ID as input parameter to be used in step by step debugging
function ShowRunspaceID
{
            $id = [runspace]::DefaultRunspace.Id
            $app = [System.Diagnostics.Process]::GetCurrentProcess()
            [System.Windows.Forms.MessageBox]::Show("application: $($app.name)"+[Environment]::NewLine+"runspace ID: $id")
}

#create folder structure based on seqential file numbering; 
# parameters: Filenumber (has to be number in string format) and number of files per folder as digit, e.g. 3 for max. 999 files.
function mGetFolderNumber($_FileNumber, $_nChar)
{
	#$_FileNumber = "1000000"
	$_l = $_FileNumber.Length
	#$_nChar = 3 # number of files per folder
	$_nO = [math]::Ceiling( $_FileNumber.Length/$_nChar)
	$_NumberArray = @()
	$_d = 0
	$_n = 0

	do{
	if ($_l-$_nChar -ge 0) { 
			$_NumberArray += $_FileNumber.Substring($_l-$_nChar,$_nChar)
		}
		else {
			if ($_d -gt 0) {
				$_NumberArray += $_FileNumber.Substring(0, $_d)
			}
		}
		$_l -= $_nChar
		$_d = $_FileNumber.Length-(($_n+1)*$_nChar)
		$_n +=1
	}
	while ($_n -le $_nO+1) 

	$_Folders = @()
	for ($_i = 1; $_i -lt $_NumberArray.Count; $_i++)
	{
		if ($_NumberArray[$_i] -eq "000") { $_Folder = 0 }
		else { [int16]$_Folder = $_NumberArray[$_i] }
		$_Folders += $_Folder
	}

	$_ItemFilePath = "$/xDMS/"
	for ($_i = 0; $_i -lt $_Folders.Count; $_i++) {
		$_ItemFilePath = $_ItemFilePath + $_Folders[$_i] + "/"
	}
	return $_ItemFilePath

} #end function mGetFolderNumber


# VDS Dialogs and Tabs share UIString according DSLanguage.xml override or default powerShell UI culture;
# VDS MenuCommand scripts don't read as a default; call this function in case $UIString[] key value pairs are needed
function mGetUIStrings
{
	# check language override settings of VDS
	[xml]$mDSLangFile = Get-Content "C:\ProgramData\Autodesk\Vault 2019\Extensions\DataStandard\Vault\DSLanguages.xml"
	$mUICodes = $mDSLangFile.SelectNodes("/DSLanguages/Language_Code")
	$mLCode = @{}
	Foreach ($xmlAttr in $mUICodes)
	{
		$mKey = $xmlAttr.ID
		$mValue = $xmlAttr.InnerXML
		$mLCode.Add($mKey, $mValue)
	}
	#If override exists, apply it, else continue with $PSUICulture
	If ($mLCode["UI"]){
		$mVdsUi = $mLCode["UI"]
	} 
	Else{$mVdsUi=$PSUICulture}
	[xml]$mUIStrFile = get-content ("C:\ProgramData\Autodesk\Vault 2019\Extensions\DataStandard\" + $mVdsUi + "\UIStrings.xml")
	$UIString = @{}
	$xmlUIStrs = $mUIStrFile.SelectNodes("/UIStrings/UIString")
	Foreach ($xmlAttr in $xmlUIStrs) {
		$mKey = $xmlAttr.ID
		$mValue = $xmlAttr.InnerXML
		$UIString.Add($mKey, $mValue)
		}
	return $UIString
}

# VDS Dialogs and Tabs share property name translations $Prop[_XLTN_*] according DSLanguage.xml override or default powerShell UI culture;
# VDS MenuCommand scripts don't read as a default; call this function in case $UIString[] key value pairs are needed
function mGetPropTranslations
{
	# check language override settings of VDS
	[xml]$mDSLangFile = Get-Content "C:\ProgramData\Autodesk\Vault 2019\Extensions\DataStandard\Vault\DSLanguages.xml"
	$mUICodes = $mDSLangFile.SelectNodes("/DSLanguages/Language_Code")
	$mLCode = @{}
	Foreach ($xmlAttr in $mUICodes)
	{
		$mKey = $xmlAttr.ID
		$mValue = $xmlAttr.InnerXML
		$mLCode.Add($mKey, $mValue)
	}
	#If override exists, apply it, else continue with $PSUICulture
	If ($mLCode["DB"]){
		$mVdsDb = $mLCode["DB"]
	} 
	Else{
		$mVdsDb=$PSUICulture
	}
	[xml]$mPrpTrnsltnFile = get-content ("C:\ProgramData\Autodesk\Vault 2019\Extensions\DataStandard\" + $mVdsDb + "\PropertyTranslations.xml")
	$mPrpTrnsltns = @{}
	$xmlPrpTrnsltns = $mPrpTrnsltnFile.SelectNodes("/PropertyTranslations/PropertyTranslation")
	Foreach ($xmlAttr in $xmlPrpTrnsltns) {
		$mKey = $xmlAttr.Name
		$mValue = $xmlAttr.InnerXML
		$mPrpTrnsltns.Add($mKey, $mValue)
		}
	return $mPrpTrnsltns
}

function Adsk.CreateTcFileLink([string]$FileFullVaultPath )
{
    $FullPaths = @($FileFullVaultPath)

    $Files = $vault.DocumentService.FindLatestFilesByPaths($FullPaths)

    $IDs = @($Files[0].Id)
     
    $PersIDs = $vault.KnowledgeVaultService.GetPersistentIds("FILE", $IDs, [Autodesk.Connectivity.WebServices.EntPersistOpt]::Latest)
    $PersID = $PersIDs[0].TrimEnd("=")

    $serverUri = [System.Uri]$Vault.InformationService.Url

    $vaultName = $VaultConnection.Vault
    $Server = $VaultConnection.Server

    $TCLink = $serverUri.Scheme + "://" + $Server + "/AutodeskTC/" + $Server + "/" + $vaultName + "/#/Entity/Details?id=m" + "$PersID" + "&itemtype=File"
    
    return $TCLink
}

function Adsk.CreateTcFolderLink([string]$FolderFullVaultPath)
{
    $Folder = $vault.DocumentService.GetFolderByPath($FolderFullVaultPath)

    $IDs = @($Folder.Id)
     
    $PersIDs = $vault.KnowledgeVaultService.GetPersistentIds("FLDR", $IDs, [Autodesk.Connectivity.WebServices.EntPersistOpt]::Latest)
    $PersID = $PersIDs[0].TrimEnd("=")

    $serverUri = [System.Uri]$Vault.InformationService.Url

    $vaultName = $VaultConnection.Vault
    $Server = $VaultConnection.Server

    $TCLink = $serverUri.Scheme + "://" + $Server + "/AutodeskTC/" + $Server + "/" + $vaultName + "/#/Entity/Entities?folder=m" + "$PersID" + "&start=0"
    return $TCLink
}

function Adsk.CreateTCItemLink ([Long]$ItemId)
{
	$IDs = @($ItemId)
    $PersIDs = $vault.KnowledgeVaultService.GetPersistentIds("ITEM", $IDs, [Autodesk.Connectivity.WebServices.EntPersistOpt]::Latest)
    $PersID = $PersIDs[0].TrimEnd("=")

    $serverUri = [System.Uri]$Vault.InformationService.Url

    $vaultName = $VaultConnection.Vault
    $Server = $VaultConnection.Server

    $TCLink = $serverUri.Scheme + "://" + $Server + "/AutodeskTC/" + $Server + "/" + $vaultName + "/#/Entity/Details?id=m" + "$PersID" + "&itemtype=Item"
    return $TCLink
}

#function to check that the current user is member of a named group; returns true or false
function Adsk.GroupMemberOf([STRING]$mGroupName)
{
	$mGroupInfo = New-Object Autodesk.Connectivity.WebServices.GroupInfo
	$mGroup = $vault.AdminService.GetGroupByName($mGroupName)
	$mGroupInfo = $vault.AdminService.GetGroupInfoByGroupId($mGroup.Id)
	foreach ($user in $mGroupInfo.Users)
	{
		if($vault.AdminService.SecurityHeader.UserId -eq $user.Id)
		{				
			return $true
		}
	}
	return $false
}

function mSearchCustentOfCat([String]$mCatDispName)
{
	$mSearchString = $mCatDispName
	$srchCond = New-Object autodesk.Connectivity.WebServices.SrchCond
	$propDefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("CUSTENT")
	$propDef = $propDefs | Where-Object { $_.SysName -eq "CategoryName" }
	$srchCond.PropDefId = $propDef.Id
	$srchCond.SrchOper = 3 
	$srchCond.SrchTxt = $mSearchString
	$srchCond.PropTyp = [Autodesk.Connectivity.WebServices.PropertySearchType]::SingleProperty
	$srchCond.SrchRule = [Autodesk.Connectivity.WebServices.SearchRuleType]::Must
	$srchSort = New-Object autodesk.Connectivity.WebServices.SrchSort
	$searchStatus = New-Object autodesk.Connectivity.WebServices.SrchStatus
	$bookmark = ""     
	$mResultAll = New-Object 'System.Collections.Generic.List[Autodesk.Connectivity.WebServices.CustEnt]'
	
	while(($searchStatus.TotalHits -eq 0) -or ($mResultAll.Count -lt $searchStatus.TotalHits))
	{
		$mResultPage = $vault.CustomEntityService.FindCustomEntitiesBySearchConditions(@($srchCond),@($srchSort),[ref]$bookmark,[ref]$searchStatus)			
		If ($searchStatus.IndxStatus -ne "IndexingComplete" -or $searchStatus -eq "IndexingContent")
		{
			#check the indexing status; you might return a warning that the result bases on an incomplete index, or even return with a stop/error message, that we need to have a complete index first
		}
		if($mResultPage.Count -ne 0)
		{
			$mResultAll.AddRange($mResultPage)
		}
		else { break;}

		return $mResultAll				
		break; #limit the search result to the first result page; page scrolling not implemented in this snippet release
	}
}

function mGetProjectFolderPropToVaultFile ([String] $mFolderSourcePropertyName, [String] $mFileTargetPropertyName)
{
	$mPath = $Prop["_FilePath"].Value
	$mFld = $vault.DocumentService.GetFolderByPath($mPath)

	IF ($mFld.Cat.CatName -eq $UIString["CAT6"]) { $Global:mProjectFound = $true}
	ElseIf ($mPath -ne "$"){
		Do {
			$mParID = $mFld.ParID
			$mFld = $vault.DocumentService.GetFolderByID($mParID)
			IF ($mFld.Cat.CatName -eq $UIString["CAT6"]) { $Global:mProjectFound = $true}
		} Until (($mFld.Cat.CatName -eq $UIString["CAT6"]) -or ($mFld.FullName -eq "$"))
	}

	If ($mProjectFound -eq $true) {
		#Project's property Value copied to file property
		$Prop[$mFileTargetPropertyName].Value = mGetFolderPropValue $mFld.Id $mFolderSourcePropertyName
	}
	Else{
		#empty field value if file will not link to a project
		$Prop[$mFileTargetPropertyName].Value = ""
	}
}