﻿#=============================================================================
# PowerShell script sample for Vault Data Standard                            
#			 Autodesk Vault - Quickstart 2019  								  
# This sample is based on VDS 2018/2019 RTM and adds functionality and rules    
#                                                                             
# Copyright (c) Autodesk - All rights reserved.                               
#                                                                             
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES 
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  
#=============================================================================

function InitializeWindow
{
	#region rules applying commonly
    $dsWindow.Title = SetWindowTitle
	InitializeFileNameValidation
	#InitializeCategory #Quickstart differentiates for Inventor and AutoCAD
	#InitializeNumSchm #Quickstart differentiates for Inventor and AutoCAD
	#InitializeBreadCrumb #Quickstart differentiates Inventor, Inventor C&H, T&P, FG, DA dialogs
	#endregion rules applying commonly

	$mWindowName = $dsWindow.Name
	switch($mWindowName)
	{
		"InventorWindow"
		{
			InitializeBreadCrumb
			#	there are some custom functions to enhance functionality:
			[System.Reflection.Assembly]::LoadFrom($Env:ProgramData + "\Autodesk\Vault 2019\Extensions\DataStandard" + '\Vault.Custom\addinVault\QuickstartUtilityLibrary.dll')

			#	initialize the context for Drawings or presentation files as these have Vault Option settings
			$global:mGFN4Special = $Prop["_GenerateFileNumber4SpecialFiles"].Value
					
			if ($global:mGFN4Special -eq $true)
			{
				$dsWindow.FindName("GFN4Special").IsChecked = $true # this checkbox is used by the XAML dialog styles, to enable / disable or show / hide controls
			}
			$mInvDocuFileTypes = (".IDW", ".DWG", ".IPN") #to compare that the current new file is one of the special files the option applies to
			if ($mInvDocuFileTypes -contains $Prop["_FileExt"].Value) {
				$global:mIsInvDocumentationFile = $true
				$dsWindow.FindName("chkBxIsInvDocuFileType").IsChecked = $true
				If ($global:mIsInvDocumentationFile-eq $true -and $global:mGFN4Special -eq $false) #IDW/DWG, IPN - Don't generate new document number
				{ 
					$dsWindow.FindName("BreadCrumb").IsEnabled = $false
					$dsWindow.FindName("GroupFolder").Visibility = "Collapsed"
					$dsWindow.FindName("expShortCutPane").Visibility = "Collapsed"
				}
				Else {$dsWindow.FindName("BreadCrumb").IsEnabled = $true} #IDW/DWG, IPN - Generate new document number
			}

			$global:_ModelPath = $null
			switch ($Prop["_CreateMode"].Value) 
			{
				$true 
				{
					$Prop["Part Number"].Value = "" #reset the part number for new files as Inventor writes the file name (no extension) as a default.
					#$dsDiag.Trace(">> CreateMode Section executes...")
					# set the category: VDS Quickstart 2019 supports extended category differentiation for 3D components
					InitializeInventorCategory
					InitializeInventorNumSchm

					#region FDU Support --------------------------------------------------------------------------
					
					# Read FDS related internal meta data; required to manage particular workflows
					$_mInvHelpers = New-Object QuickstartUtilityLibrary.InvHelpers
					If ($_mInvHelpers.m_FDUActive($Application))
					{
						#[System.Windows.MessageBox]::Show("Active FDU-AddIn detected","Vault MFG Quickstart")
						$_mFdsKeys = $_mInvHelpers.m_GetFdsKeys($Application, @{})

						# some FDS workflows require VDS cancellation; add the conditions to the event handler _Loaded below
						$dsWindow.add_Loaded({
							IF ($mSkipVDS -eq $true)
							{
								$dsWindow.CancelWindowCommand.Execute($this)
								#$dsDiag.Trace("FDU-VDS EventHandler: Skip Dialog executed")	
							}
						})

						# FDS workflows with individual settings					
						$dsWindow.FindName("Categories").add_SelectionChanged({
							If ($Prop["_Category"].Value -eq "Factory Asset" -and $Document.FileSaveCounter -eq 0) #don't localize name according FDU fixed naming
							{
								$paths = @("Factory Asset Library Source")
								mActivateBreadCrumbCmbs $paths
								$dsWindow.FindName("NumSchms").SelectedIndex = 1
							}
						})
				
						If($_mFdsKeys.ContainsKey("FdsType") -and $Document.FileSaveCounter -eq 0 )
						{
							#$dsDiag.Trace(" FDS File Type detected")
							# for new assets we suggest to use the source file folder name, nothing else
							If($_mFdsKeys.Get_Item("FdsType") -eq "FDS-Asset")
							{
								# only the MSDCE FDS configuration template provides a category for assets, check for this otherwise continue with the selection done before
								$mCatName = GetCategories | Where {$_.Name -eq "Factory Asset"}
								IF ($mCatName) { $Prop["_Category"].Value = "Factory Asset"}
							}
							# skip for publishing the 3D temporary file save event for VDS
							If($_mFdsKeys.Get_Item("FdsType") -eq "FDS-Asset" -and $Application.SilentOperation -eq $true)
							{ 
								#$dsDiag.Trace(" FDS publishing 3D - using temporary assembly silent mode: need to skip VDS!")
								$global:mSkipVDS = $true
							}
							If($_mFdsKeys.Get_Item("FdsType") -eq "FDS-Asset" -and $Document.InternalName -ne $Application.ActiveDocument.InternalName)
							{
								#$dsDiag.Trace(" FDS publishing 3D: ActiveDoc.InternalName different from VDSDoc.Internalname: Verbose VDS")
								$global:mSkipVDS = $true
							}

							# 
							If($_mFdsKeys.Get_Item("FdsType") -eq "FDS-Layout" -and $_mFdsKeys.Count -eq 1)
							{
								#$dsDiag.Trace("3DLayout, not synced")
								# only the MSDCE FDS configuration template provides a category for layouts, check for this otherwise continue with the selection done before
								$mCatName = GetCategories | Where {$_.Name -eq "Factory Layout"}
								IF ($mCatName) { $Prop["_Category"].Value = "Factory Layout"}
							}

							# FDU 2019.22.0.2 allows to skip dynamically, instead of skipping in general by the SkipVDSon1stSave.IAM template
							If($_mFdsKeys.Get_Item("FdsType") -eq "FDS-Layout" -and $_mFdsKeys.Count -gt 1 -and $Document.FileSaveCounter -eq 0)
							{
								#$dsDiag.Trace("3DLayout not saved yet, but already synced")
								$dsWindow.add_Loaded({
									$dsWindow.CancelWindowCommand.Execute($this)
									#$dsDiag.Trace("FDU-VDS EventHandler: Skip Dialog executed")	
								})
							}
						}
					}
					#endregion FDU Support --------------------------------------------------------------------------

					#retrieve 3D model properties (Inventor captures these also, but too late; we are currently before save event transfers model properties to drawing properties) 
					# but don't do this, if the copy mode is active
					if ($Prop["_CopyMode"].Value -eq $false) 
					{	
						if (($Prop["_FileExt"].Value -eq "idw") -or ($Prop["_FileExt"].Value -eq "dwg" )) 
						{
							$_mInvHelpers = New-Object QuickstartUtilityLibrary.InvHelpers #NEW 2019 hand over the parent inventor application, to ensure the correct instance
							$_ModelFullFileName = $_mInvHelpers.m_GetMainViewModelPath($Application)#NEW 2019 hand over the parent inventor application, to ensure the correct instance
							$Prop["Title"].Value = $_mInvHelpers.m_GetMainViewModelPropValue($Application, $_ModelFullFileName,"Title")
							$Prop["Description"].Value = $_mInvHelpers.m_GetMainViewModelPropValue($Application, $_ModelFullFileName,"Description")
							$Prop["Part Number"].Value = $_mInvHelpers.m_GetMainViewModelPropValue($Application, $_ModelFullFileName,"Part Number") 
						}

						if ($Prop["_FileExt"].Value -eq "ipn") 
						{
							$_mInvHelpers = New-Object QuickstartUtilityLibrary.InvHelpers #NEW 2019 hand over the parent inventor application, to ensure the correct instance
							$_ModelFullFileName = $_mInvHelpers.m_GetMainViewModelPath($Application)#NEW 2019 hand over the parent inventor application, to ensure the correct instance
							$Prop["Title"].Value = $_mInvHelpers.m_GetMainViewModelPropValue($Application, $_ModelFullFileName,"Title")
							$Prop["Description"].Value = $_mInvHelpers.m_GetMainViewModelPropValue($Application, $_ModelFullFileName,"Description")
							$Prop["Part Number"].Value = $_mInvHelpers.m_GetMainViewModelPropValue($Application, $_ModelFullFileName,"Part Number")
							$Prop["Stock Number"].Value = $_mInvHelpers.m_GetMainViewModelPropValue($Application, $_ModelFullFileName,"Stock Number")
							# for custom properties there is always a risk that any does not exist
							try {
								$Prop[$_iPropSemiFinished].Value = $_mInvHelpers.m_GetMainViewModelPropValue($Application, $_ModelFullFileName,$_iPropSemiFinished)
								$_t1 = $_mInvHelpers.m_GetMainViewModelPropValue($Application, $_ModelFullFileName, $_iPropSpearWearPart)
								if ($_t1 -ne "") {
									$Prop[$_iPropSpearWearPart].Value = $_t1
								}
							} 
							catch {
								$dsDiag.Trace("Set path, filename and properties for IPN: At least one custom property failed, most likely it did not exist and is not part of the cfg ")
							}
						}

						if (($_ModelFullFileName -eq "") -and ($global:mGFN4Special -eq $false)) 
						{ 
							[System.Windows.MessageBox]::Show($UIString["MSDCE_MSG00"],"Vault MFG Quickstart")
							$dsWindow.add_Loaded({
										# Will skip VDS Dialog for Drawings without model view; 
										$dsWindow.CancelWindowCommand.Execute($this)})
						}
					} # end of copy mode = false check

					if ($Prop["_CopyMode"].Value -and @("DWG","IDW","IPN") -contains $Prop["_FileExt"].Value)
					{
						$Prop["DocNumber"].Value = $Prop["DocNumber"].Value.TrimStart($UIString["CFG2"])
					}
					
					#} #end of copymode = true
				}
				$false # EditMode = True
				{
					#add specific action rules for edit mode here
				}
				default
				{

				}
			} #end switch Create / Edit Mode

		}
		"AutoCADWindow"
		{
			InitializeBreadCrumb
			switch ($Prop["_CreateMode"].Value) 
			{
				$true 
				{
					#$dsDiag.Trace(">> CreateMode Section executes...")
					# set the category: Quickstart = "AutoCAD Drawing"
					$mCatName = GetCategories | Where {$_.Name -eq $UIString["MSDCE_CAT01"]}
					IF ($mCatName) { $Prop["_Category"].Value = $UIString["MSDCE_CAT01"]}
						# in case the current vault is not quickstart, but a plain MFG default configuration
					Else {
						$mCatName = GetCategories | Where {$_.Name -eq $UIString["CAT1"]} #"Engineering"
						IF ($mCatName) { $Prop["_Category"].Value = $UIString["CAT1"]}
					}

					#region FDU Support ------------------
					$_FdsUsrData = $Document.UserData #Items FACT_* are added by FDU
					[System.Reflection.Assembly]::LoadFrom($Env:ProgramData + "\Autodesk\Vault 2019\Extensions\DataStandard" + '\Vault.Custom\addinVault\QuickstartUtilityLibrary.dll')
					$_mAcadHelpers = New-Object QuickstartUtilityLibrary.AcadHelpers
					$_FdsBlocksInDrawing = $_mAcadHelpers.mFdsDrawing($Application)
					If($_FdsUsrData.Get_Item("FACT_FactoryDocument") -and $_FdsBlocksInDrawing )
					{
						#try to activate category "Factory Layout"
						$Prop["_Category"].Value = "Factory Layout"
					}
					#endregion FDU Support ---------------
				}
			}

			#endregion quickstart
		}
		default
		{
			#rules applying for other windows, e.g. FG, DA, TP and CH functional dialogs; SaveCopyAs dialog
		}
	} #end switch windows
	$global:expandBreadCrumb = $true
	#$dsDiag.Trace("... Initialize window end <<")
}#end InitializeWindow

function AddinLoaded
{
	#Executed when DataStandard is loaded in Inventor/AutoCAD
		$m_File = $env:TEMP + "\Folder2019.xml"
		if (!(Test-Path $m_File)){
			$source = $Env:ProgramData + "\Autodesk\Vault 2019\Extensions\DataStandard\Vault.Custom\Folder2019.xml"
			Copy-Item $source $env:TEMP\Folder2019.xml
		}
}

function AddinUnloaded
{
	#Executed when DataStandard is unloaded in Inventor/AutoCAD
}

function SetWindowTitle
{
	$mWindowName = $dsWindow.Name
    switch($mWindowName)
 	{
  		"InventorFrameWindow"
  		{
   			$windowTitle = $UIString["LBL54"]
  		}
  		"InventorDesignAcceleratorWindow"
  		{
   			$windowTitle = $UIString["LBL50"]
  		}
  		"InventorPipingWindow"
  		{
   			$windowTitle = $UIString["LBL39"]
  		}
  		"InventorHarnessWindow"
  		{
   			$windowTitle = $UIString["LBL44"]
  		}
  		default #applies to InventorWindow and AutoCADWindow
  		{
   			if ($Prop["_CreateMode"].Value)
   			{
    			if ($Prop["_CopyMode"].Value)
    			{
     				$windowTitle = "$($UIString["LBL60"]) - $($Prop["_OriginalFileName"].Value)"
    			}
    			elseif ($Prop["_SaveCopyAsMode"].Value)
    			{
     				$windowTitle = "$($UIString["LBL72"]) - $($Prop["_OriginalFileName"].Value)"
    			}else
    			{
     				$windowTitle = "$($UIString["LBL24"]) - $($Prop["_OriginalFileName"].Value)"
    			}
   			}
   			else
   			{
    			$windowTitle = "$($UIString["LBL25"]) - $($Prop["_FileName"].Value)"
   			} 
  		}
 	}
  	return $windowTitle
}

function InitializeInventorNumSchm
{
	if ($Prop["_SaveCopyAsMode"].Value -eq $true)
    {
        $Prop["_NumSchm"].Value = $UIString["LBL77"]
    }
	if($Prop["_Category"].Value -eq $UIString["MSDCE_CAT12"]) #Substitutes, as reference parts should not retrieve individual new number
	{
		$Prop["_NumSchm"].Value = $UIString["LBL77"]
	}
	if($dsWindow.Name -eq "InventorFrameWindow")
	{
		$Prop["_NumSchm"].Value = $UIString["LBL77"]
	}
}

function InitializeInventorCategory
{
	$mDocType = $Document.DocumentType
	$mDocSubType = $Document.SubType #differentiate part/sheet metal part and assembly/weldment assembly
	switch ($mDocType)
	{
		'12291' #assembly
		{ 
			$mCatName = GetCategories | Where {$_.Name -eq $UIString["MSDCE_CAT10"]} #assembly, available in Quickstart Advanced, e.g. INV-Samples Vault
			IF ($mCatName) 
			{ 
				$Prop["_Category"].Value = $UIString["MSDCE_CAT10"]
			}
			$mCatName = GetCategories | Where {$_.Name -eq $UIString["MSDCE_CAT02"]}
			IF ($mCatName) 
			{ 
				$Prop["_Category"].Value = $UIString["MSDCE_CAT02"] #3D Component, Quickstart, e.g. MFG-2019-PRO-EN
			}
			Else 
			{
				$mCatName = GetCategories | Where {$_.Name -eq $UIString["CAT1"]} #"Engineering"
				IF ($mCatName) 
				{ 
					$Prop["_Category"].Value = $UIString["CAT1"]
				}
			}
			If($mDocSubType -eq "{28EC8354-9024-440F-A8A2-0E0E55D635B0}") #weldment assembly
			{
				$mCatName = GetCategories | Where {$_.Name -eq $UIString["MSDCE_CAT11"]} # weldment assembly
				IF ($mCatName) 
				{ 
					$Prop["_Category"].Value = $UIString["MSDCE_CAT10"]
				}
			} 
		}
		'12290' #part
		{
			$mCatName = GetCategories | Where {$_.Name -eq $UIString["MSDCE_CAT08"]} #Part, available in Quickstart Advanced, e.g. INV-Samples Vault
			IF ($mCatName) 
			{ 
				$Prop["_Category"].Value = $UIString["MSDCE_CAT08"]
			}
			$mCatName = GetCategories | Where {$_.Name -eq $UIString["MSDCE_CAT02"]}
			IF ($mCatName) 
			{ 
				$Prop["_Category"].Value = $UIString["MSDCE_CAT02"] #3D Component, Quickstart, e.g. MFG-2019-PRO-EN
			}
			Else 
			{
				$mCatName = GetCategories | Where {$_.Name -eq $UIString["CAT1"]} #"Engineering"
				IF ($mCatName) 
				{ 
					$Prop["_Category"].Value = $UIString["CAT1"]
				}
			}
			If($mDocSubType -eq "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}") 
			{
				$mCatName = GetCategories | Where {$_.Name -eq $UIString["MSDCE_CAT09"]} #sheet metal part, available in Quickstart Advanced, e.g. INV-Samples Vault
				IF ($mCatName) 
				{ 
					$Prop["_Category"].Value = $UIString["MSDCE_CAT09"]
				}
			}
			If($Document.IsSubstitutePart -eq $true) 
			{
				$mCatName = GetCategories | Where {$_.Name -eq $UIString["MSDCE_CAT12"]} #substitute, available in Quickstart Advanced, e.g. INV-Samples Vault
				IF ($mCatName) 
				{ 
					$Prop["_Category"].Value = $UIString["MSDCE_CAT12"]
				}
			}			
		}
		'12292' #drawing
		{
			$mCatName = GetCategories | Where {$_.Name -eq $UIString["MSDCE_CAT00"]}
			IF ($mCatName) { $Prop["_Category"].Value = $UIString["MSDCE_CAT00"]}
			Else # in case the current vault is not quickstart, but a plain MFG default configuration
			{
				$mCatName = GetCategories | Where {$_.Name -eq $UIString["CAT1"]} #"Engineering"
				IF ($mCatName) { $Prop["_Category"].Value = $UIString["CAT1"]}
			}
		}
		'12293' #presentation
		{
			$mCatName = GetCategories | Where {$_.Name -eq $UIString["MSDCE_CAT13"]} #presentation, available in Quickstart Advanced, e.g. INV-Samples Vault
			IF ($mCatName) 
			{ 
				$Prop["_Category"].Value = $UIString["MSDCE_CAT13"]
			}
			$mCatName = GetCategories | Where {$_.Name -eq $UIString["MSDCE_CAT02"]} #3D Component, Quickstart, e.g. MFG-2019-PRO-EN
			IF ($mCatName) 
			{ 
				$Prop["_Category"].Value = $UIString["MSDCE_CAT02"]
			}
			Else 
			{
				$mCatName = GetCategories | Where {$_.Name -eq $UIString["CAT1"]} #"Engineering"
				IF ($mCatName) 
				{ 
					$Prop["_Category"].Value = $UIString["CAT1"]
				}
			}
		}
	} #DocType Switch
}

function GetNumSchms
{
	try
	{
		if (-Not $Prop["_EditMode"].Value)
        {
            #quickstart - there is the use case that we don't need a number: IDW/DWG, IPN and Option Generate new file number = off
			If ($global:mIsInvDocumentationFile-eq $true -and $global:mGFN4Special -eq $false) 
			{ 
				return
			}
			[System.Collections.ArrayList]$numSchems = @($vault.DocumentService.GetNumberingSchemesByType('Activated'))
			$_FilteredNumSchems = @()
			$_temp = $numSchems | Where { $_.IsDflt -eq $true}
			$_FilteredNumSchems += ($_temp)
			if ($Prop["_NumSchm"].Value) { $Prop["_NumSchm"].Value = $_FilteredNumSchems[0].Name} #note - functional dialogs don't have the property _NumSchm, therefore we conditionally set the value
			$dsWindow.FindName("NumSchms").IsEnabled = $true
			$noneNumSchm = New-Object 'Autodesk.Connectivity.WebServices.NumSchm'
			$noneNumSchm.Name = $UIString["LBL77"] # None 
			$_FilteredNumSchems += $noneNumSchm

			#reverse order for these cases; none is added latest; reverse the list, if None is pre-set to index = 0

			If($dsWindow.Name-eq "InventorWindow" -and $Prop["DocNumber"].Value -notlike "Assembly*" -and $Prop["_FileExt"].Value -eq ".iam") #you might find better criteria based on then numbering scheme
			{
				$_FilteredNumSchems = $_FilteredNumSchems | Sort-Object -Descending
				return $_FilteredNumSchems
			}
			If($dsWindow.Name-eq "InventorWindow" -and $Prop["DocNumber"].Value -notlike "Part*" -and $Prop["_FileExt"].Value -eq ".ipt") #you might find better criteria based on then numbering scheme
			{
				$_FilteredNumSchems = $_FilteredNumSchems | Sort-Object -Descending
				return $_FilteredNumSchems
			}
			If($dsWindow.Name-eq "InventorFrameWindow")
			{ 
				#$_FilteredNumSchems = $_FilteredNumSchems | Sort-Object -Descending
				return $_FilteredNumSchems
			}
	
			return $_FilteredNumSchems
        }
	}
	catch [System.Exception]
	{		
		[System.Windows.MessageBox]::Show($error)
	}	
}

function GetCategories
{
	$mAllCats =  $vault.CategoryService.GetCategoriesByEntityClassId("FILE", $true)
	$mFDSFilteredCats = $mAllCats | Where { $_.Name -ne "Asset Library"}
	return $mFDSFilteredCats
}

function OnPostCloseDialog
{
	$mWindowName = $dsWindow.Name
	switch($mWindowName)
	{
		"InventorWindow"
		{
				if (!($Prop["_CopyMode"].Value -and !$Prop["_GenerateFileNumber4SpecialFiles"].Value -and @("DWG","IDW","IPN") -contains $Prop["_FileExt"].Value))
				{
					mWriteLastUsedFolder
				}
			#new 2019 QS, remove file extentions if used in validation for preview new file name
			if($Prop["_SaveCopyAsMode"].Value -eq $true)
			{
				$Global:OnPostAction = $true
				$newFileName = @()
				$newFileName += ($Prop["DocNumber"].Value.Split("."))
				If($newFileName.Count -gt 1) 
				{
					$newExt = $newFileName[$newFileName.Count-1]
					$Prop["DocNumber"].Value = $Prop["DocNumber"].Value.Replace("." + $newExt, "")
				}
				Else 
				{ 
					$Prop["DocNumber"].Value = $newFileName[0]
				}
			}#end _SaveCopyAsMode

				if ($Prop["_CreateMode"].Value -and !$Prop["Part Number"].Value) #we empty the part number on initialize: if there is no other function to provide part numbers we should apply the Inventor default
				{
					$Prop["Part Number"].Value = $Prop["DocNumber"].Value
				}
		}

		"AutoCADWindow"
		{
			mWriteLastUsedFolder
		}
		default
		{
			#rules applying for windows non specified
		}
	} #switch Window Name
	
}

function mHelp ([Int] $mHContext) {
	try
	{
		switch ($mHContext){
			100 {
				$mHPage = "C.2Inventor.html";
			}
			110 {
				$mHPage = "C.2.11FrameGenerator.html";
			}
			120 {
				$mHPage = "C.2.13DesignAccelerator.html";
			}
			130 {
				$mHPage = "C.2.12TubeandPipe.html";
			}
			140 {
				$mHPage = "C.2.14CableandHarness.html";
			}
			200 {
				$mHPage = "C.3AutoCADAutoCAD.html";
			}
			Default {
				$mHPage = "Index.html";
			}
		}
		$mHelpTarget = $Env:ProgramData + "\Autodesk\Vault 2019\Extensions\DataStandard\HelpFiles\"+$mHPage
		$mhelpfile = Invoke-Item $mHelpTarget 
	}
	catch
	{
		[System.Windows.MessageBox]::Show($UIString["MSDCE_MSG02"], "Vault Quickstart Client")
	}
}

function mReadShortCuts {
	if ($Prop["_CreateMode"].Value -eq $true) {
		#$dsDiag.Trace(">> Looking for Shortcuts...")
		$m_Server = $VaultConnection.Server
		$m_Vault = $VaultConnection.Vault
		$m_AllFiles = @()
		$m_FiltFiles = @()
		$m_Path = $env:APPDATA + '\Autodesk\VaultCommon\Servers\Services_Security_1_16_2018\'
		$m_AllFiles += Get-ChildItem -Path $m_Path -Filter 'Shortcuts.xml' -Recurse
		$m_AllFiles | ForEach-Object {
			if ($_.FullName -like "*"+$m_Server + "*" -and $_.FullName -like "*"+$m_Vault + "*") 
			{
				$m_FiltFiles += $_
			} 
		}
		$global:mScFile = $m_FiltFiles.SyncRoot[$m_FiltFiles.Count-1].FullName
		if (Test-Path $global:mScFile) {
			#$dsDiag.Trace(">> Start reading Shortcuts...")
			$global:m_ScXML = New-Object XML 
			$global:m_ScXML.Load($mScFile)
			$m_ScAll = $m_ScXML.Shortcuts.Shortcut
			#the shortcuts need to get filtered by type of document.folder and path information related to CAD workspace
			$global:m_ScCAD = @{}
			$mScNames = @()
			#$dsDiag.Trace("... Filtering Shortcuts...")
			$m_ScAll | ForEach-Object { 
				if (($_.NavigationContextType -eq "Connectivity.Explorer.Document.DocFolder") -and ($_.NavigationContext.URI -like "*"+$global:CAx_Root + "/*"))
				{
					try
					{
						$_t = $global:m_ScCAD.Add($_.Name, $_.NavigationContext.URI)
						$mScNames += $_.Name
					}
					catch {
						$dsDiag.Trace("... ERROR Filtering Shortcuts...")
					}
				}
			}
		}
		$dsDiag.Trace("... returning Shortcuts: $mScNames")
		return $mScNames
	}
}

function mScClick {
	try 
	{
		$_key = $dsWindow.FindName("lstBoxShortCuts").SelectedValue
		$_Val = $global:m_ScCAD.get_item($_key)
		$_SPath = @()
		$_SPath = $_Val.Split("/")

		$m_DesignPathNames = $null
		[System.Collections.ArrayList]$m_DesignPathNames = @()
		#differentiate AutoCAD and Inventor: AutoCAD is able to start in $, but Inventor starts in it's mandatory Workspace folder (IPJ)
		IF ($dsWindow.Name -eq "InventorWindow") {$indexStart = 2}
		If ($dsWindow.Name -eq "AutoCADWindow") {$indexStart = 1}
		for ($index = $indexStart; $index -lt $_SPath.Count; $index++) 
		{
			$m_DesignPathNames += $_SPath[$index]
		}
		if ($m_DesignPathNames.Count -eq 1) { $m_DesignPathNames += "."}
		mActivateBreadCrumbCmbs $m_DesignPathNames
		$global:expandBreadCrumb = $true
	}
	catch
	{
		#$dsDiag.Trace("mScClick function - error reading selected value")
	}
	
}

function mAddSc {
	try
	{
		$mNewScName = $dsWindow.FindName("txtNewShortCut").Text
		mAddShortCutByName ($mNewScName)
		$dsWindow.FindName("lstBoxShortCuts").ItemsSource = mReadShortCuts
	}
	catch {}
}

function mRemoveSc {
	try
	{
		$_key = $dsWindow.FindName("lstBoxShortCuts").SelectedValue
		mRemoveShortCutByName $_key
		$dsWindow.FindName("lstBoxShortCuts").ItemsSource = mReadShortCuts
	}
	catch { }
}

function mAddShortCutByName([STRING] $mScName)
{
	try #simply check that the name is unique
	{
		#$dsDiag.Trace(">> Start to add ShortCut, check for used name...")
		$global:m_ScCAD.Add($mScName,"Dummy")
		$global:m_ScCAD.Remove($mScName)
	}
	catch #no reason to continue in case of existing name
	{
		[System.Windows.MessageBox]::Show($UIString["MSDCE_MSG01"], "Vault Quickstart Client")
		end function
	}

	try 
	{
		#$dsDiag.Trace(">> Continue to add ShortCut, creating new from template...")	
		#read from template
		$m_File = $env:TEMP + "\Folder2019.xml"
		if (Test-Path $m_File)
		{
			#$dsDiag.Trace(">>-- Started to read Folder2019.xml...")
			$global:m_XML = New-Object XML
			$global:m_XML.Load($m_File)
		}
		$mShortCut = $global:m_XML.Folder.Shortcut | where { $_.Name -eq "Template"}
		#clone the template completely and update name attribute and navigationcontext element
		$mNewSc = $mShortCut.Clone() #.CloneNode($true)
		#rename "Template" to new name
		$mNewSc.Name = $mScName 

		#derive the path from current selection
		$breadCrumb = $dsWindow.FindName("BreadCrumb")
		$newURI = "vaultfolderpath:" + $global:CAx_Root
		foreach ($cmb in $breadCrumb.Children) 
		{
			$_N = $cmb.SelectedItem.Name
			#$dsDiag.Trace(" - selecteditem.Name of cmb: $_N ")
			if (($cmb.SelectedItem.Name.Length -gt 0) -and !($cmb.SelectedItem.Name -eq "."))
			{ 
				$newURI = $newURI + "/" + $cmb.SelectedItem.Name
				#$dsDiag.Trace(" - the updated URI  of the shortcut: $newURI")
			}
			else { break}
		}
		
		#hand over the path in shortcut navigation format
		$mNewSc.NavigationContext.URI = $newURI
		#append the new shortcut and save back to file
		$mImpNode = $global:m_ScXML.ImportNode($mNewSc,$true)
		$global:m_ScXML.Shortcuts.AppendChild($mImpNode)
		$global:m_ScXML.Save($mScFile)
		$dsWindow.FindName("txtNewShortCut").Text = ""
		#$dsDiag.Trace("..successfully added ShortCut <<")
		return $true
	}
	catch 
	{
		$dsDiag.Trace("..problem encountered adding ShortCut <<")
		return $false
	}
}

function mRemoveShortCutByName ([STRING] $mScName)
{
	try 
	{
		#$dsDiag.Trace(">> Start to remove ShortCut from list")
		$mShortCut = @() #Vault allows multiple shortcuts equally named
		$mShortCut = $global:m_ScXML.Shortcuts.Shortcut | where { $_.Name -eq $mScName}
		$mShortCut | ForEach-Object {
			$global:m_ScXML.Shortcuts.RemoveChild($_)
		}
		$global:m_ScXML.Save($global:mScFile)
		#$dsDiag.Trace("..successfully removed ShortCut <<")
		return $true
	}
	catch 
	{
		return $false
	}
}