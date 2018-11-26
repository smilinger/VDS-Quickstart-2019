﻿
#=============================================================================#
# PowerShell script sample for Vault Data Standard                            
#			 Autodesk Vault - Quickstart 2019  								  
# This sample is based on VDS 2018 RTM and adds functionality and rules		  
#                                                                             
# Copyright (c) Autodesk - All rights reserved.                               
#                                                                             
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES 
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  
#=============================================================================#

#this function will be called to check if the Ok button can be enabled
function ActivateOkButton
{
		return Validate;
}

# sample validation function
# finds all function definition with names beginning with
# ValidateFile, ValidateFolder and ValidateTask respectively
# these funcions should return a boolean value, $true if the Property is valid
# $false otherwise
# As soon as one property validation function returns $false the entire Validate function will return $false
function Validate
{
	$mWindowName = $dsWindow.Name
	switch($mWindowName)
	{
		"FileWindow"
		{
			foreach ($func in dir function:ValidateFile*) { if(!(&$func)) { return $false } }
			return $true
		}
		"FolderWindow"
		{
			foreach ($func in dir function:ValidateFolder*) { if(!(&$func)) { return $false } }
			return $true
		}
		"CustomObjectWindow"
		{
			foreach ($func in dir function:ValidateCustomObject*) { if(!(&$func)) { return $false } }
			return $true
		}
		default { return $true }
	}
    
}

# sample validation function for the FileName property
# if the File Name is empty the validation will fail
function ValidateFileName
{
	if($Prop["_FileName"].Value -or !$dsWindow.FindName("DSNumSchmsCtrl").NumSchmFieldsEmpty)
	{
		return $true;
	}
	return $false;
}

function ValidateFolderName
{
	if($Prop["_FolderName"].Value -or !$dsWindow.FindName("DSNumSchmsCtrl").NumSchmFieldsEmpty)
	{
		return $true;
	}
	return $false;
}

function ValidateCustomObjectName
{
	if($Prop["_CustomObjectName"].Value -or !$dsWindow.FindName("DSNumSchmsCtrl").NumSchmFieldsEmpty)
	{
		return $true;
	}
	return $false;
}

function InitializeTabWindow
{
	#$dsDiag.ShowLog()
	#$dsDiag.Inspect()
}

function InitializeWindow
{	      

        #$dsDiag.ShowLog()
        #$dsDiag.Clear()
      
	#begin rules applying commonly
	$Prop["_Category"].add_PropertyChanged({
        if ($_.PropertyName -eq "Value")
        {
			#region quickstart
				#$Prop["_NumSchm"].Value = $Prop["_Category"].Value
				m_CategoryChanged
			#endregion
        }		
    })
	#end rules applying commonly
	$mWindowName = $dsWindow.Name
	switch($mWindowName)
	{
		"FileWindow"
		{
			#rules applying for File
			$dsWindow.Title = SetWindowTitle $UIString["LBL24"] $UIString["LBL25"] $Prop["_FileName"].Value
			if ($Prop["_CreateMode"].Value)
			{
				if ($Prop["_IsOfficeClient"].Value)
				{
					$Prop["_Category"].Value = $UIString["CAT2"]
				}
				else
				{
					$Prop["_Category"].Value = $UIString["CAT1"]
				}
			}
			#region Quickstart to get category from selected template
			$dsWindow.FindName("DocTypeCombo").add_SelectionChanged({
					mResetTemplates
				})
			$dsWindow.FindName("TemplateCB").add_SelectionChanged({
				m_TemplateChanged})

			if($dsWindow.FindName("tabItemProperties")) { mInitializeTabItemProps}

			if($dsWindow.FindName("tabClassification")) 
			{ 
				$dsWindow.FindName("tabClassification").add_GotFocus({
					
						if(-not $Global:mClsTabInitialized)
						{
							mInitializeClassificationTab -ParentType "Dialog" -file $null
						}
					}) #add selection changed
			}
			#endregion Quickstart
			
		}
		"FolderWindow"
		{
			#rules applying for Folder
			$dsWindow.Title = SetWindowTitle $UIString["LBL29"] $UIString["LBL30"] $Prop["_FolderName"].Value
			if ($Prop["_CreateMode"].Value)
			{
				$Prop["_Category"].Value = $UIString["CAT5"]
			}
			#region Quickstart - for imported folders easily set title to folder name on edit
			If ($Prop["_EditMode"].Value) {
				If ($Prop["_XLTN_TITLE"]){
					IF ($Prop["_XLTN_TITLE"].Value -eq $null) {
						$Prop["_XLTN_TITLE"].Value = $Prop["Name"].Value
					}
				}
			}
			#endregion Quickstart
		}
		"CustomObjectWindow"
		{
			#rules applying for Custom Object
			$dsWindow.Title = SetWindowTitle $UIString["LBL61"] $UIString["LBL62"] $Prop["_CustomObjectName"].Value
			if ($Prop["_CreateMode"].Value)
			{
				$Prop["_Category"].Value = $Prop["_CustomObjectDefName"].Value

				#region Quickstart
					$dsWindow.FindName("Categories").IsEnabled = $false
					$dsWindow.FindName("NumSchms").Visibility = "Collapsed"
					$Prop["_NumSchm"].Value = $Prop["_Category"].Value
				#endregion
			}
		}
		#region GoToInvSibling
		"GoToInvSibling"
		{
			$_t = $Prop["Internal ID"].Value

			$mTargetFile = Get-Content $env:TEMP"\mStrTabClick.txt"

			#create search conditions for 
			$srchConds = New-Object autodesk.Connectivity.WebServices.SrchCond[] 1
				$srchCond = New-Object autodesk.Connectivity.WebServices.SrchCond
				$propDefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE")
				$propNames = @("Internal ID")
				$propDefIds = @{}
				foreach($name in $propNames) {
					$propDef = $propDefs | Where-Object { $_.DispName -eq $name }
					$propDefIds[$propDef.Id] = $propDef.DispName
				}
				$srchCond.PropDefId = $propDef.Id
				$srchCond.SrchOper = 3
				$srchCond.SrchTxt = $Prop["Internal ID"].Value

				$srchCond.PropTyp = [Autodesk.Connectivity.WebServices.PropertySearchType]::SingleProperty
				$srchCond.SrchRule = [Autodesk.Connectivity.WebServices.SearchRuleType]::Must
				$srchConds[0] = $srchCond
				$srchSort = New-Object autodesk.Connectivity.WebServices.SrchSort
				$mSearchStatus = New-Object autodesk.Connectivity.WebServices.SrchStatus
				$mBookmark = ""     
				$mResultAll = New-Object 'System.Collections.Generic.List[Autodesk.Connectivity.WebServices.File]'
				$mResultFiltered = New-Object 'System.Collections.Generic.List[Autodesk.Connectivity.WebServices.File]'
	
				while(($mSearchStatus.TotalHits -eq 0) -or ($mResultAll.Count -lt $mSearchStatus.TotalHits))
				{
					$mResultPage = $vault.DocumentService.FindFilesBySearchConditions($srchConds,@($srchSort),@(($vault.DocumentService.GetFolderRoot()).Id),$true,$false,[ref]$mBookmark,[ref]$mSearchStatus)
					if($mResultPage.Count -ne 0)
					{
						$mResultAll.AddRange($mResultPage)
					}
					else { break;}
		
					break; #limit the search result to the first result page; page scrolling not implemented in this snippet release
				}

				If($mResultAll.Count -lt $mSearchStatus.TotalHits)
				{
					$dsWindow.FindName("txtBlck_Notification1").Text = ([String]::Format($UIString["ADSK-GoToInvSibl_MSG00"], $mResultAll.Count, $mSearchStatus.TotalHits))
				}
				Else{
					$dsWindow.FindName("txtBlck_Notification1").Visibility = "Collapsed"
				}

				IF($mResultAll.Count -gt 0)
				{
					$mResultFiltered += $mResultAll | Where-Object {$_.Name -ne ($Prop["_FileName"].Value + $Prop["_FileExt"].Value)} # don't list the selected file's versions as parallel copies
					If ($mResultFiltered.Count -gt 0)
					{
						$dsWindow.FindName("mDerivatives1").ItemsSource = @(mGetResultMeta $mResultFiltered) # @() #the array of each file iterations meta data
					}
					$dsWindow.FindName("mDerivatives1").add_SelectionChanged({
						mDerivatives1Click
						If($dsWindow.FindName("mDerivatives1").SelectedItem)
							{
								$dsWindow.FindName("btnGoTo").IsEnabled = $true
							}
						})
				}
				IF($mResultFiltered.Count -lt 1)
				{
					$dsWindow.FindName("txtBlck_Notification1").Visibility = "Visible"
					$dsWindow.FindName("txtBlck_Notification1").Text = $UIString["ADSK-GoToInvSibl_MSG01"]
				}
		}
		#endregion GoToInvSibling

		#region CustomObjectTermWindow-CatalogTermsTranslations
		"CustomObjectTermWindow"
		{      
			IF ($Prop["_CreateMode"].Value -eq $true) 
			{
				$Prop["_Category"].Value = $Prop["_CustomObjectDefName"].Value

					$dsWindow.FindName("Categories").IsEnabled = $false
					$dsWindow.FindName("NumSchms").Visibility = "Collapsed"
					$Prop["_NumSchm"].Value = $Prop["_Category"].Value

				IF($Prop["_XLTN_IDENTNUMBER"]){ $Prop["_XLTN_IDENTNUMBER"].Value = $UIString["LBL27"]}
			}

			#region EditMode
			IF ($Prop["_EditMode"].Value -eq $true) 
			{
				#read existing classification elements
				$_classes = @()
				Try{ #likely not all properties are used...
					If ($Prop["_XLTN_SEGMENT"].Value.Length -gt 1){
						$_classes += $Prop["_XLTN_SEGMENT"].Value
						If ($Prop["_XLTN_MAINGROUP"].Value.Length -gt 1){
							$_classes += $Prop["_XLTN_MAINGROUP"].Value
							If ($Prop["_XLTN_GROUP"].Value.Length -gt 1){
								$_classes += $Prop["_XLTN_GROUP"].Value
								If ($Prop["_XLTN_SEGMENT"].Value.Length -gt 1){
									$_classes += $Prop["_XLTN_SUBGROUP"].Value
								}
							}
						}
					}
				}
				catch {}
			}
			#endregion EditMode
			mAddCoCombo -_CoName "Segment" -_classes $_classes #enables classification for catalog of terms
			# ToDo: createmode: activate last used classification
			
		} # objectterm Window
		#endregion CatalogTermsTranslations-CustomObjectTermsWindow

		"ActivateSchedTaskWindow"
		{
			ADSK.QS.ReadSchedTasks
			$dsWindow.FindName("dataGrdSchedTasks").add_SelectionChanged({ ADSK.QS.TaskSelectionChanged})
		}

		"FileImportWindow"
		{
			mInitializeFileImport
		}
	}
}

function SetWindowTitle($newFile, $editFile, $name)
{
	if ($Prop["_CreateMode"].Value)
    {
		$windowTitle = $newFile
	}
	elseif ($Prop["_EditMode"].Value)
	{
		$windowTitle = "$($editFile) - $($name)"
	}
	elseif ($Prop["_ReadOnly"].Value)
	{
		$windowTitle = "$($editFile) - $($name)$($UIString["LBL26"])"
	}
	return $windowTitle
}

function OnLogOn
{
	#Executed when User logs on Vault; $vaultUsername can be used to get the username, which is used in Vault on login
	
	#check Vault Client Version to match this configuration requirements; note - Office Client registers as WG or PRO 
	$mVaultVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Autodesk\Vault Workgroup\24.0\VWG-2440:407\").ProductVersion.Split(".")
	If(-not $mVaultVersion) { $mVaultVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Autodesk\Vault Professional\24.0\VPRO-2440:407\").ProductVersion.Split(".")}

	If($mVaultVersion[0] -ne "24" -or $mVaultVersion[1] -lt "1" )
	{
		[System.Windows.MessageBox]::Show("This machine's Vault Data Standard configuration requires Vault Client 2019 Update 1 or newer installed; contact your system administrator.", "Vault Quickstart Client Configuration")
	}
}
function OnLogOff
{
	#Executed when User logs off Vault
}

function GetTitleWindow
{
	$message = "Autodesk Data Standard - Create/Edit "+$Prop["_FileName"]
	return $message
}

#fired when the file selection changes
function OnTabContextChanged
{      
	$xamlFile = [System.IO.Path]::GetFileName($VaultContext.UserControl.XamlFile)

	if ($VaultContext.SelectedObject.TypeId.SelectionContext -eq "FileMaster" -and $xamlFile -eq "ADSK.QS.CAD BOM.xaml")
	{
		$fileMasterId = $vaultContext.SelectedObject.Id
		$file = $vault.DocumentService.GetLatestFileByMasterId($fileMasterId)
		$bom = @(GetFileBOM($file.id))
		$dsWindow.FindName("bomList").ItemsSource = $bom
	}
	if ($VaultContext.SelectedObject.TypeId.SelectionContext -eq "ItemMaster" -and $xamlFile -eq "Associated Files.xaml")
	{
		$items = $vault.ItemService.GetItemsByIds(@($vaultContext.SelectedObject.Id))
		$item = $items[0]
		$itemids = @($item.Id)
		$assocFiles = @(GetAssociatedFiles $itemids $([System.IO.Path]::GetDirectoryName($VaultContext.UserControl.XamlFile)))
		$dsWindow.FindName("AssoicatedFiles").ItemsSource = $assocFiles
	}
	if ($VaultContext.SelectedObject.TypeId.SelectionContext -eq "FileMaster" -and $xamlFile -eq "ADSK.QS.FileDataSheet.xaml")
	{
		$fileMasterId = $vaultContext.SelectedObject.Id
		$file = $vault.DocumentService.GetLatestFileByMasterId($fileMasterId)
		mInitializeClassificationTab -ParentType $null -file $file
	}
	
	if ($VaultContext.SelectedObject.TypeId.SelectionContext -eq "ItemMaster" -and $xamlFile -eq "ADSK.QS.ItemEdit.xaml")
	{
		$items = $vault.ItemService.GetItemsByIds(@($vaultContext.SelectedObject.Id))
		$item = $items[0]
		$itemids = @($item.Id)
		$mItemEditable = mItemEditable($itemids) #note - checks the current state to activate buttons, but this might change over time; therefore the state is local
	}
	
	if ($VaultContext.SelectedObject.TypeId.SelectionContext -eq "ChangeOrder" -and $xamlFile -eq "ADSK.QS.EcoEdit.xaml")
	{
		$number=$vaultContext.SelectedObject.Label
		$mECO = $vault.ChangeOrderService.GetChangeOrderByNumber($number)
		$mEcoEditable = mEcoEditable($mECO.Id) #note - checks the current state to activate buttons, but this might change over time; therefore the state is local
	}

	if ($VaultContext.SelectedObject.TypeId.SelectionContext -eq "FileMaster" -and $xamlFile -eq "ADSK.QS.FileItemDataSheet.xaml")
	{
		#clear any data from previous selections:
		mClearItemView
		$fileMasterId = $vaultContext.SelectedObject.Id
		$file = $vault.DocumentService.GetLatestFileByMasterId($fileMasterId)
		#fill item system and user properties
		mFillItemView -file $file
		#fill grid of associated files; primary only Yes/No
		$itemids = @($item.Id)
		$assocFiles = @(mFileItemTabGetAssocFiles $itemids $([System.IO.Path]::GetDirectoryName($VaultContext.UserControl.XamlFile)))
		$dsWindow.FindName("AssociatedFiles").ItemsSource = $assocFiles
	}

	#region Documentstructure Extension
		if ($VaultContext.SelectedObject.TypeId.SelectionContext -eq "FileMaster" -and $xamlFile -eq "ADSK.QS.FileDocStructure.xaml")
		{
			Add-Type -Path 'C:\ProgramData\Autodesk\Vault 2019\Extensions\DataStandard\Vault.Custom\addinVault\UsesWhereUsed.dll'
			$file = $vault.DocumentService.GetLatestFileByMasterId($vaultContext.SelectedObject.Id)
			$treeNode = New-Object UsesWhereUsed.TreeNode($file, $vaultConnection)
			$dsWindow.FindName("Uses").ItemsSource = @($treeNode)

			$dsWindow.FindName("Uses").add_SelectedItemChanged({
				mUwUsdChldrnClick
				#$dsDiag.Trace("Child selected")
			})
			$dsWindow.FindName("WhereUsed").add_SelectedItemChanged({
				mUwUsdPrntClick
				$dsDiag.Trace("Child selected")
			})

		}
	#endregion documentstructure

	#region derivation tree
	if ($VaultContext.SelectedObject.TypeId.SelectionContext -eq "FileMaster" -and $xamlFile -eq "ADSK.QS.DerivationTree.xaml")
	{
		if($dsWindow.FindName("chkAutoUpdate").IsChecked)
		{
			mDerivationTreeUpdateView #$vaultContext.SelectedObject.Id
			$dsWindow.FindName("btnUpdate").IsEnabled = $false
		}
		if($dsWindow.FindName("chkAutoUpdate").IsChecked -eq $false)
		{ 
			mDerivationTreeResetView
		}
	}
	#endregion derivation tree

	#region ItemTab-FileImport
if ($VaultContext.SelectedObject.TypeId.SelectionContext -eq "ItemMaster" -and $xamlFile -eq "ADSK.QS.AttachFileImport.xaml")
	{
		$items = $vault.ItemService.GetItemsByIds(@($vaultContext.SelectedObject.Id))
		$item = $items[0]
		
		If (!$item.Locked)
		{
			$dsWindow.FindName("mDragAreaEnabled").Source = "C:\ProgramData\Autodesk\Vault 2019\Extensions\DataStandard\Vault.Custom\DragFilesActive.png"
			$dsWindow.FindName("mDragAreaEnabled").Visibility = "Visible"
			$dsWindow.FindName("mDragAreaDisabled").Visibility = "Collapsed"
			$dsWindow.FindName("txtActionInfo").Visibility = "Visible"
		}
		Else
		{
			$dsWindow.FindName("mDragAreaDisabled").Source = "C:\ProgramData\Autodesk\Vault 2019\Extensions\DataStandard\Vault.Custom\DragFilesLocked.png"
			$dsWindow.FindName("mDragAreaDisabled").Visibility = "Visible"
			$dsWindow.FindName("mDragAreaEnabled").Visibility = "Collapsed"
			$dsWindow.FindName("txtActionInfo").Visibility = "Collapsed"
		}

		Try{
			Import-Module powerVault
		}
		catch{
		   [System.Windows.MessageBox]::Show("This feature requires powerVault installed; check for its availability", "Extension Title")
		   return
		}

		$dsWindow.FindName("mImportProgress").Value = 0
		$dsWindow.FindName("mDragAreaEnabled").add_Drop({			
			param( $sender, $e)
			$items = $vault.ItemService.GetItemsByIds(@($vaultContext.SelectedObject.Id))
			$item = $items[0]

			#check that the item is editable for the current user, if not, we shouldn't add the files, before we try to attach
			try{
				$vault.ItemService.EditItems(@($item.RevId))
				#[System.Windows.MessageBox]::Show("Item is accessible", "Item-File Attachment Import")
				$_ItemIsEditable = $true
			}
			catch {
				#[System.Windows.MessageBox]::Show("Item is NOT accessible", "Item-File Attachment Import")
				$_ItemIsEditable = $false
			}
			If($_ItemIsEditable)
			{
				$vault.ItemService.UndoEditItems(@($item.RevId))
				$vault.ItemService.DeleteUncommittedItems($true)
				#[System.Windows.MessageBox]::Show("Item Lock Removed to continue", "Item-File Attachment Import")
			}
			
			[System.Windows.DataObject]$mDragData = $e.Data
			$mFileList = $mDragData.GetFileDropList()
			#Filter folders, we attach files directly selected only
			$mFileList = $mFileList | Where { (get-item $_).PSIsContainer -eq $false }
			If ($mFileList -and $_ItemIsEditable)
			{
				$dsWindow.Cursor = "Wait"
				$_NumFiles = $mFileList.Count
				$_n = 0
				$dsWindow.FindName("mImportProgress").Value = 0
				$mExtExclude = @(".ipt", ".iam", ".ipn", ".dwg", ".idw", ".slddrw", ".sldprt", ".sldasm")
				$m_ImpFileList = @() #filepath array of imported files to be attached
				ForEach ($_file in $mFileList)
				{
					$m_FileName = [System.IO.Path]::GetFileNameWithoutExtension($_file)
					$m_Ext = [System.IO.Path]::GetExtension($_file)
					If ($mExtExclude -contains $m_Ext){
						$mCADWarning = $true
						break;
					}
					$m_Dir = [System.IO.Path]::GetDirectoryName($_file)
					
					#get new number and create new file name
					[System.Collections.ArrayList]$numSchems = @($vault.DocumentService.GetNumberingSchemesByType('Activated'))
					if ($numSchems.Count -gt 1)
					{							
						$_DfltNumSchm = $numSchems | Where { $_.Name -eq $UIString["ADSK-ItemFileImport_00"]}
						if($_DfltNumSchm)
						{
							$NumGenArgs = @("")
							$_newFile=$vault.DocumentService.GenerateFileNumber($_DfltNumSchm.SchmID, $NumGenArgs)
						}		
					}

					#add file
					If($_newFile)
					{
						#get appropriate folder number (limit 1k files per folder)
						Try{
							$mTargetPath = mGetFolderNumber $_newFile 3 #hand over the file number (name) and number of files / folder
						}
						catch { 
							[System.Windows.MessageBox]::Show($UIString["ADSK-ItemFileImport_01"], "Item-File Attachment Import")
						}
						#add extension to number
						$_newFile = $_newFile + $m_Ext
						$mFullTargetPath = $mTargetPath + $_newFile
						$m_ImportedFile = Add-VaultFile -From $_file -To $mFullTargetPath -Comment $UIString["ADSK-ItemFileImport_02"]
						$m_ImpFileList += $m_ImportedFile._FullPath
					}
					Else #continue with the given file name
					{
						$mTargetPath = "$/xDMS/"
						$mFullTargetPath = $mTargetPath + $m_FileName
						$m_ImportedFile = Add-VaultFile -From $_file -To $mFullTargetPath -Comment $UIString["ADSK-ItemFileImport_02"]
						$m_ImpFileList += $m_ImportedFile._FullPath
					}
					$_n += 1
					$dsWindow.FindName("mImportProgress").Value = (($_n/$_NumFiles)*100)-10

				} #for each file
				#attach file to current item
				$parent = Get-VaultItem -Number $item.ItemNum	
				$parentUpdated = Update-VaultItem -Number $parent._Number -AddAttachments $m_ImpFileList -Comment $UIString["ADSK-ItemFileImport_03"]
				$dsWindow.FindName("mImportProgress").Value = (($_n/$_NumFiles)*100)
				If ($mCADWarning)
				{
					[System.Windows.MessageBox]::Show($UIString["ADSK-ItemFileImport_04"], "Item-File Attachment Import")
				}
			}
			$mFileList = $null
			$dsWindow.Cursor = "Arrow"
			$dsWindow.FindName("mDragAreaEnabled").remove_Drop()
			}) #end drag & drop
	}
#endregion ItemTab-FileImport
}

function GetNewCustomObjectName
{
	#region Quickstart
		$m_Cat = $Prop["_Category"].Value
		switch ($m_Cat)
		{
			$UIString["MSDCE_CO02"] #Person
			{
				if($dsWindow.FindName("DSNumSchmsCtrl").NumSchmFieldsEmpty -eq $false)
				{
					$Prop["_XLTN_IDENTNUMBER"].Value = $Prop["_GeneratedNumber"].Value
				}
				$customObjectName = $Prop["_XLTN_FIRSTNAME"].Value + " " + $Prop["_XLTN_LASTNAME"].Value
				return $customObjectName
			}

			$UIString["ClassTerms_00"] #CatalogTermsTranslations
			{
				if($dsWindow.FindName("DSNumSchmsCtrl").NumSchmFieldsEmpty -eq $false)
				{
					Try {$Prop["_XLTN_IDENTNUMBER"].Value = $Prop["_GeneratedNumber"].Value}
					catch { $dsDiag.Trace("UDP to save ident number does not exist")}
				}
				$customObjectName = $Prop["_XLTN_TERM"].Value
				return $customObjectName
			}

			Default 
			{
				#region - default configuration's behavior
				if($dsWindow.FindName("DSNumSchmsCtrl").NumSchmFieldsEmpty)
				{	
					#$dsDiag.Trace("read text from TextBox CUSTOMOBJECTNAME")
					$customObjectName = $dsWindow.FindName("CUSTOMOBJECTNAME").Text
					#$dsDiag.Trace("customObjectName = $customObjectName")
				}
				else{
					#$dsDiag.Trace("-> GenerateNumber")
					$customObjectName = $Prop["_GeneratedNumber"].Value
					#$dsDiag.Trace("customObjectName = $customObjectName")
				}
				return $customObjectName
				#endregion - default configuration's behavior
			}
		}
	#endregion Quickstart
}

#Constructs the filename(numschems based or handtyped)and returns it.
function GetNewFileName
{
	$dsDiag.Trace(">> GetNewFileName")
	if($dsWindow.FindName("DSNumSchmsCtrl").NumSchmFieldsEmpty)
	{	
		$dsDiag.Trace("read text from TextBox FILENAME")
		$fileName = $dsWindow.FindName("FILENAME").Text
		$dsDiag.Trace("fileName = $fileName")
	}
	else{
		$dsDiag.Trace("-> GenerateNumber")
		$fileName = $Prop["_GeneratedNumber"].Value
		$dsDiag.Trace("fileName = $fileName")
		#Quickstart
			If($Prop["_XLTN_PARTNUMBER"]) { $Prop["_XLTN_PARTNUMBER"].Value = $Prop["_GeneratedNumber"].Value }
		#Quickstart
	}
	$newfileName = $fileName + $Prop["_FileExt"].Value
	$dsDiag.Trace("<< GetNewFileName $newfileName")
	return $newfileName
}

function GetNewFolderName
{
	$dsDiag.Trace(">> GetNewFolderName")
	if($dsWindow.FindName("DSNumSchmsCtrl").NumSchmFieldsEmpty)
	{	
		$dsDiag.Trace("read text from TextBox FOLDERNAME")
		$folderName = $dsWindow.FindName("FOLDERNAME").Text
		$dsDiag.Trace("folderName = $folderName")
	}
	else{
		$dsDiag.Trace("-> GenerateNumber")
		$folderName = $Prop["_GeneratedNumber"].Value
		$dsDiag.Trace("folderName = $folderName")
	}
	$dsDiag.Trace("<< GetNewFolderName $folderName")
	return $folderName
}

# This function can be used to force a specific folder when using "New Standard File" or "New Standard Folder" functions.
# If an empty string is returned the selected folder is used
# ! Do not remove the function
function GetParentFolderName
{
	$folderName = ""
	return $folderName
}

function GetCategories
{
	if ($dsWindow.Name -eq "FileWindow")
	{
		#return $vault.CategoryService.GetCategoriesByEntityClassId("FILE", $true)
		#region quickstart
			$global:mFileCategories = $vault.CategoryService.GetCategoriesByEntityClassId("FILE", $true)
			return $global:mFileCategories
		#endregion
	}
	elseif ($dsWindow.Name -eq "FolderWindow")
	{
		return $vault.CategoryService.GetCategoriesByEntityClassId("FLDR", $true)
	}
	elseif ($dsWindow.Name -eq "CustomObjectWindow")
	{
		return $vault.CategoryService.GetCategoriesByEntityClassId("CUSTENT", $true)
	}
	elseif ($dsWindow.Name -eq "CustomObjectTermWindow")
	{
		return $vault.CategoryService.GetCategoriesByEntityClassId("CUSTENT", $true)
	}
}

function GetNumSchms
{
	if ($Prop["_CreateMode"].Value)
	{
		try
		{
			[System.Collections.ArrayList]$numSchems = @($vault.DocumentService.GetNumberingSchemesByType('Activated'))
			if ($numSchems.Count -gt 1)
			{
				#$numSchems = $numSchems | Sort-Object -Property IsDflt -Descending
				#region Quickstart
					$mWindowName = $dsWindow.Name
					switch($mWindowName)
					{
						"FileWindow"
						{
							$_FilteredNumSchems = $numSchems | Where { $_.IsDflt -eq $true}
							$Prop["_NumSchm"].Value = $_FilteredNumSchems[0].Name
							$dsWindow.FindName("NumSchms").IsEnabled = $false
							return $_FilteredNumSchems
						}

						"FolderWindow" 
						{
							#numbering schemes are available for items and files specificly; 
							#for folders we use the file numbering schemes and filter to these, that have a corresponding name in folder categories
							$_FolderCats = $vault.CategoryService.GetCategoriesByEntityClassId("FLDR", $true)
							$_FilteredNumSchems = @()
							Foreach ($item in $_FolderCats) 
							{
								$_temp = $numSchems | Where { $_.Name -eq $item.Name}
								$_FilteredNumSchems += ($_temp)
							}
							#we need an option to unselect a previosly selected numbering; to achieve that we add a virtual one, named "None"
							$noneNumSchm = New-Object 'Autodesk.Connectivity.WebServices.NumSchm'
							$noneNumSchm.Name = "None"
							$_FilteredNumSchems += ($noneNumSchm)

							return $_FilteredNumSchems
						}

						"CustomObjectWindow"
						{
							$_FilteredNumSchems = $numSchems | Where { $_.Name -eq $Prop["_Category"].Value}
							return $_FilteredNumSchems
						}

						"CustomObjectTermWindow"
						{
							$_FilteredNumSchems = $numSchems | Where-Object { $_.Name -eq $Prop["_Category"].Value}
							return $_FilteredNumSchems
						}

						default
						{
							$numSchems = $numSchems | Sort-Object -Property IsDflt -Descending
							return $numSchems
						}
					}
				#region
			}
			Else {
				$dsWindow.FindName("NumSchms").IsEnabled = $false				
			}
			return $numSchems
		}
		catch [System.Exception]
		{		
			#[System.Windows.MessageBox]::Show($error)
		}
	}
}


# Decides if the NumSchmes field should be visible
function IsVisibleNumSchems
{
	$ret = "Collapsed"
	$numSchems = $vault.DocumentService.GetNumberingSchemesByType([Autodesk.Connectivity.WebServices.NumSchmType]::Activated)
	if($numSchems.Length -gt 0)
	{	$ret = "Visible" }
	return $ret
}

#Decides if the FileName should be enabled, it should only when the NumSchmField isnt
function ShouldEnableFileName
{
	$ret = "true"
	$numSchems = $vault.DocumentService.GetNumberingSchemesByType([Autodesk.Connectivity.WebServices.NumSchmType]::Activated)
	if($numSchems.Length -gt 0)
	{	$ret = "false" }
	return $ret
}

function ShouldEnableNumSchms
{
	$ret = "false"
	$numSchems = $vault.DocumentService.GetNumberingSchemesByType([Autodesk.Connectivity.WebServices.NumSchmType]::Activated)
	if($numSchems.Length -gt 0)
	{	$ret = "true" }
	return $ret
}

#define the parametrisation for the number generator here
function GenerateNumber
{
	$dsDiag.Trace(">> GenerateNumber")
	$selected = $dsWindow.FindName("NumSchms").Text
	if($selected -eq "") { return "na" }

	$ns = $global:numSchems | Where-Object { $_.Name.Equals($selected) }
	switch ($selected) {
		"Sequential" { $NumGenArgs = @(""); break; }
		default      { $NumGenArgs = @(""); break; }
	}
	$dsDiag.Trace("GenerateFileNumber($($ns.SchmID), $NumGenArgs)")
	$vault.DocumentService.GenerateFileNumber($ns.SchmID, $NumGenArgs)
	$dsDiag.Trace("<< GenerateNumber")
}

#define here how the numbering preview should look like
function GetNumberPreview
{
	$selected = $dsWindow.FindName("NumSchms").Text
	switch ($selected) {
		"Sequential" { $Prop["_FileName"].Value="???????"; break; }
		"Short" { $Prop["_FileName"].Value=$Prop["Project"].Value + "-?????"; break; }
		"Long" { $Prop["_FileName"].Value=$Prop["Project"].Value + "." + $Prop["Material"].Value + "-?????"; break; }
		default { $Prop["_FileName"].Value="NA" }
	}
}

#Workaround for Property names containing round brackets
#Xaml fails to parse
function ItemTitle
{
    if ($Prop)
	{
       $val = $Prop["_XLTN_TITLE_ITEM_CO"].Value
	   return $val
    }
}

#Workaround for Property names containing round brackets
#Xaml fails to parse
function ItemDescription
{
	if($Prop)#Tab gets loaded before the SelectionChanged event gets fired resulting with null Prop. Happens when the vault is started with Change Order as the last view.
    {
       $val = $Prop["_XLTN_DESCRIPTION_ITEM_CO"].Value
	   return $val
    }
}

 
function m_TemplateChanged {
	#$dsDiag.Trace(">> Template Changed ...")
	$mContext = $dsWindow.DataContext
	$mTemplatePath = $mContext.TemplatePath
	$mTemplateFile = $mContext.SelectedTemplate
	$mTemplate = $mTemplatePath + "/" + $mTemplateFile
	$mFolder = $vault.DocumentService.GetFolderByPath($mTemplatePath)
	$mFiles = $vault.DocumentService.GetLatestFilesByFolderId($mFolder.Id,$false)
	$mTemplateFile = $mFiles | Where-Object { $_.Name -eq $mTemplateFile }
	$Prop["_Category"].Value = $mTemplateFile.Cat.CatName
	$mCatName = $mTemplateFile.Cat.CatName
	$dsWindow.FindName("Categories").SelectedValue = $mCatName
	If ($mCatName) #if something went wrong the user should be able to select a category
	{
		$dsWindow.FindName("Categories").IsEnabled = $false #comment out this line if admins like to release the choice to the user
	}
	#$dsDiag.Trace(" ... TemplateChanged finished <<")
}

function m_CategoryChanged 
{
	$mWindowName = $dsWindow.Name
    switch($mWindowName)
	{
		"FileWindow"
		{
			#Quickstart uses the default numbering scheme for files; GoTo GetNumSchms function to disable this filter incase you'd like to apply numbering per category for files as well
			$Prop['_XLTN_AUTHOR'].Value = $VaultConnection.UserName

			#Copy Parent Project Number to file property "Project" if exists
			If($Prop["_XLTN_PROJECT"]){
				mGetProjectFolderPropToVaultFile -mFolderSourcePropertyName "Name" -mFileTargetPropertyName $Prop["_XLTN_PROJECT"].Name
			}
		}

		"FolderWindow" 
		{
			$dsWindow.FindName("NumSchms").SelectedItem = $null
			$dsWindow.FindName("NumSchms").Visibility = "Collapsed"
			$dsWindow.FindName("DSNumSchmsCtrl").Visibility = "Collapsed"
			$dsWindow.FindName("FOLDERNAME").Visibility = "Visible"
					
			$Prop["_NumSchm"].Value = $Prop["_Category"].Value
			IF ($dsWindow.FindName("DSNumSchmsCtrl").Scheme.Name -eq $Prop["_Category"].Value) 
			{
				$dsWindow.FindName("DSNumSchmsCtrl").Visibility = "Visible"
				$dsWindow.FindName("FOLDERNAME").Visibility = "Collapsed"
			}
			Else
			{
				$Prop["_NumSchm"].Value = "None" #we need to reset in case a user switches back from existing numbering scheme to manual input
			}
			
			#set the start date = today for project category
			If ($Prop["_Category"].Value -eq $UIString["CAT6"] -and $Prop["_XLTN_DATESTART"] )		
			{
				$Prop["_XLTN_DATESTART"].Value = Get-Date -displayhint date
			}
		}

		"CustomObjectWindow"
		{
			#categories are bound to CO type name
		}
		default
		{
			#nothing for 'unknown' new window type names
		}			
	} #end switch window
} #end function m_CategoryChanged

function mHelp ([Int] $mHContext) {
	Try
	{
		Switch ($mHContext){
			500 {
				$mHPage = "V.2File.html";
			}
			550 {
				$mHPage = "V.3OfficeFile.html";
			}
			600 {
				$mHPage = "V.1Folder.html";
			}
			700 {
				$mHPage = "V.6CustomObject.html";
			}
			Default {
				$mHPage = "Index.html";
			}
		}
		$mHelpTarget = "C:\ProgramData\Autodesk\Vault 2019\Extensions\DataStandard\HelpFiles\"+$mHPage
		$mhelpfile = Invoke-Item $mHelpTarget 
	}
	Catch
	{
		[System.Windows.MessageBox]::Show("Help Target not found", "Vault Quickstart Client")
	}
}
function mResetTemplates
{
	$dsWindow.FindName("TemplateCB").ItemsSource = $dsWindow.DataContext.Templates
}

function mFldrNameValidation
{
	$rootFolder = $vault.DocumentService.GetFolderByPath($Prop["_FolderPath"].Value)
	$mFldExist = mFindFolder $Prop["_FolderName"].Value $rootFolder

	If($mFldExist)
	{
		$dsWindow.FindName("FOLDERNAME").ToolTip = "Foldername exists, select a new unique one."
		$dsWindow.FindName("FOLDERNAME").BorderBrush = "Red"
		$dsWindow.FindName("FOLDERNAME").BorderThickness = "1,1,1,1"
		$dsWindow.FindName("FOLDERNAME").Background = "#FFFFDADA"
		return $false
	}
	Else
	{
		$dsWindow.FindName("FOLDERNAME").ToolTip = "Foldername is valid, press OK to create."
		$dsWindow.FindName("FOLDERNAME").BorderBrush = "#FFABADB3" #the default color
		$dsWindow.FindName("FOLDERNAME").BorderThickness = "0,1,1,0" #the default border top, right
		$dsWindow.FindName("FOLDERNAME").Background = "#FFFFFFFF"
		return $true
	}
}

function mFindFolder($FolderName, $rootFolder)
{
	$FolderPropDefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("FLDR")
    $FolderNamePropDef = $FolderPropDefs | where {$_.SysName -eq "Name"}
    $srchCond = New-Object 'Autodesk.Connectivity.WebServices.SrchCond'
    $srchCond.PropDefId = $FolderNamePropDef.Id
    $srchCond.PropTyp = "SingleProperty"
    $srchCond.SrchOper = 3 #is equal
    $srchCond.SrchRule = "Must"
    $srchCond.SrchTxt = $FolderName

    $bookmark = ""
    $status = $null
    $totalResults = @()
    while ($status -eq $null -or $totalResults.Count -lt $status.TotalHits)
    {
        $results = $vault.DocumentService.FindFoldersBySearchConditions(@($srchCond),$null, @($rootFolder.Id), $false, [ref]$bookmark, [ref]$status)

        if ($results -ne $null)
        {
            $totalResults += $results
        }
        else {break}
    }
    return $totalResults;
}