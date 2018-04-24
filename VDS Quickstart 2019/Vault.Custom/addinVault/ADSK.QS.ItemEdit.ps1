#region disclaimer
#=============================================================================#
# PowerShell script sample for Vault Data Standard                            #
#                                                                             #
# Copyright (c) Autodesk - All rights reserved.                               #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#
#endregion

function mItemEditable($ids)
{
	Try{
		$vault.ItemService.EditItems(@($ids))
		#[System.Windows.MessageBox]::Show("Item is accessible", "VDS Item Edit")
		$dsWindow.FindName("btnCancel").IsEnabled = $false
		$dsWindow.FindName("btnSave").IsEnabled = $false
		$dsWindow.FindName("btnEdit").IsEnabled = $true
		$vault.ItemService.UndoEditItems(@($ids))
		return $true
	}
	catch{
		#[System.Windows.MessageBox]::Show("Item is NOT accessible", "VDS Item Edit")
		return $false
	}
}

function EditItemProps 
{
	#reserve the item while editing
	$number=$vaultContext.SelectedObject.Label
	$item = $vault.ItemService.GetLatestItemByItemNumber($number)
	$ItemIds = @($item.RevId)
	#again check the items accessibility as it might have changed since latest selection
	if(mItemEditable($ItemIds) -eq $true)
	{
		#store the current UDP values
		$Global:CustomValue = $Prop["CustomValue"].Value
		$Global:CustomCompany = $Prop["CustomCompany"].Value
		$Global:Comments = $Prop["Comments"].Value	
		
		#load the Source for combobox and other controls
		$dsWindow.FindName("cmbCustomCompany"). ItemsSource= GetCompanies
		
		#update button state
		$dsWindow.FindName("btnCancel").IsEnabled = $true
		$dsWindow.FindName("btnSave").IsEnabled = $true
		$dsWindow.FindName("btnEdit").IsEnabled = $false
		$Global:mEditItems = @()
		$Global:mEditItems += $vault.ItemService.EditItems($ItemIds)
		
		#store the current item system property values (editable without property update)
		$Global:ItemTitle = $Global:mEditItems[0].Title #direct access of Title (Item,CO)
		$Global:ItemDescr = $Global:mEditItems[0].Detail #direct access of Description (Item,CO)
	}
	Else{
		[System.Windows.MessageBox]::Show("This Item is currently NOT accessible", "VDS Item Edit")
	}
}

function CancelItemUpdate 
{
	$dsWindow.FindName("btnCancel").IsEnabled = $False
	$dsWindow.FindName("btnSave").IsEnabled = $False
	$dsWindow.FindName("btnEdit").IsEnabled = $True
	$vault.ItemService.UndoEditItems(@($Global:mEditItems[0].RevId))
	$vaultContext.Refresh() = $true
}

function ItemUpdate 
{
	$dsDiag.Trace(">> Item Property Edit starts...")
	try
	{
		$data = @{}
		$sysDataChanged = $False

		#only store the changed values
		if ($Global:CustomValue -ne $dsWindow.FindName("txtCustomValue").Text) 
		{
			$PropDefID = $Prop["CustomValue"].Id
			$data[$PropDefID] = $dsWindow.FindName("txtCustomValue").Text 
		}
		if ($Global:CustomCompany -ne $dsWindow.FindName("cmbCustomCompany").Text)
		{
			$PropDefID = $Prop["CustomCompany"].Id
			$data[$PropDefID] = $dsWindow.FindName("cmbCustomCompany").Text
		}
		if ($Global:Comments -ne $dsWindow.FindName("txtComments").Text)
		{
			$PropDefID = $Prop["Comments"].Id
			$data[$PropDefID] = $dsWindow.FindName("txtComments").Text
		}
		
		#only update UDP(s) if property values are changed
		if($data.Count -gt 0)
		{
			$propValues = New-Object Autodesk.Connectivity.WebServices.PropInstParamArray
			$propValues.Items = New-Object Autodesk.Connectivity.WebServices.PropInstParam[] $data.Count
		
			$i = 0
			foreach($d in $data.GetEnumerator()) 
			{
				$propValues.Items[$i] = New-Object Autodesk.Connectivity.WebServices.PropInstParam -Property @{PropDefId = $d.Key;Val = $d.Value}
				$i++
			}

			# update UDPs
			$UpdateItems = $vault.ItemService.UpdateItemProperties(@($Global:mEditItems[0].RevId), $propValues)	
			$Global:mEditItems = $UpdateItems
		}
		
		#only update System Prop(s) if property values are changed
		if($Global:ItemTitle -ne $dsWindow.FindName("txtItemTitle").Text)
		{
			$Global:mEditItems[0].Title = $dsWindow.FindName("txtItemTitle").Text
			$sysDataChanged = $True
			$dsDiag.Trace("title valued really changed")
		}
		if($Global:ItemDescr -ne $dsWindow.FindName("txtItemDescr").Text)
		{
			$Global:mEditItems[0].Details = $dsWindow.FindName("txtItemDescr").Text
			$sysDataChanged = $True
			$dsDiag.Trace("Detail valued really changed")
		}
		
		if($data.Count -gt 0 -or $sysDataChanged -eq $True )
		{
			$Global:mEditItems[0].Comm = "Property Edit via VDS Item Edit Tab"
			$vault.ItemService.UpdateAndCommitItems($Global:mEditItems)
			$dsWindow.FindName("btnCancel").IsEnabled = $False
			$dsWindow.FindName("btnSave").IsEnabled = $False
			$dsWindow.FindName("btnEdit").IsEnabled = $True
			$vaultContext.Refresh() = $true
		}
		else
		{
			#nothing changed really, therefore undo item edit and reset UI
			CancelItemUpdate
		}

	} #end try
	catch 
	{
	  	#$dsDiag.Trace("...Error during edit properties...")
		CancelItemUpdate
	}
	#$dsDiag.Trace(" ...Item Property Edit finished successfully<<")
}

function GetCompanies
{
	$companies = mSearchCustentOfCat "Organisation"
	$companyNames = @()
	$companies | ForEach-Object { $companyNames += $_.Name }
	return $companyNames
}

