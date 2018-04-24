$fileId=$vaultContext.CurrentSelectionSet[0].Id
$folderId = $vaultContext.NavSelectionSet[0].Id
$folder = $vault.DocumentService.GetFolderById($folderId)

$mFile = $vault.DocumentService.GetLatestFileByMasterId($fileId)

$mFileNameSegm += ($mFile.Name.Split("."))
If($mFileNameSegm.Count -gt 1) 
{
	$mExt = $mFileNameSegm[$mFileNameSegm.Count-1]
}

#proceed only for ipt 
if($mExt -eq 'ipt')
{
	[System.Reflection.Assembly]::LoadFrom('C:\_VLTMKDE2019\Libraries\iLogic\QuickstartiLogicLibrary.dll')
	$iLogicVault = New-Object QuickstartiLogicLibrary.QuickstartiLogicLib

	$mVltFullFileName = $folder.FullName + "/" + $mFile.Name
	$mLocalFile = $iLogicVault.mGetFileByFullFileName($vaultConnection, $mVltFullFileName)

	[System.Reflection.Assembly]::LoadFrom($Env:ProgramData + "\Autodesk\Vault 2019\Extensions\DataStandard" + '\Vault.Custom\addinVault\QuickstartUtilityLibrary.dll')
	$_mInvHelpers = New-Object QuickstartUtilityLibrary.InvHelpers
	$_mInvHelpers.
	$_mVaultHelpers = New-Object QuickstartUtilityLibrary.VltHelpers
	$mInventorApplication = $_mInvHelpers.m_InventorApplication()
	$mInvActiveDocFullFileName = $_mInvHelpers.m_ActiveDocFullFileName($mInventorApplication)
	$mInvActiveDoc = Get-Item -Path $mInvActiveDocFullFileName
	If(!$mInvActiveDoc -and $mInvActiveDoc.Extension -ne ".iam" ) #proceed only for active doc = assembly
	{
		[System.Windows.MessageBox]::Show("This command expects Inventor having an assembly file active!
			Did you save the assembly?" , "Insert CAD: Component Copy")
		return
	}
	
	$mNumSchms = $vault.DocumentService.GetNumberingSchemesByType([Autodesk.Connectivity.WebServices.NumSchmType]::ApplicationDefault)
	$mNs = $mNumSchms[0]
	$NumGenArgs = @("") #add arguments in case the default is not just a sequence
	$mNewFileNumber = $vault.DocumentService.GenerateFileNumber($mNs.SchmID, $NumGenArgs)

	$path = $mInvActiveDoc.Directory.FullName

	Set-ItemProperty -Path $mLocalFile -Name IsReadOnly -Value $false
	$mCompCopy = Copy-Item ($mLocalFile) -Destination ($path + '\' + $mNewFileNumber + '.' + $mExt) -PassThru
	if($mCompCopy) 
	{
		$_mInvHelpers.m_PlaceComponent($mInventorApplication, $mCompCopy.FullName)
	}

} #end if IPT
Else
{
	[System.Windows.MessageBox]::Show("Command supports Inventor part files only!" , "Insert CAD: Component Copy")
}