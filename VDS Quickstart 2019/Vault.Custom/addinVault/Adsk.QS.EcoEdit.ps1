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

function mEcoEditable($mEcoId)
{
	Try
	{
		$vault.ChangeOrderService.EditChangeOrder($mEcoId)
		$dsWindow.FindName("btnCancel").IsEnabled = $false
		$dsWindow.FindName("btnSave").IsEnabled = $false
		$dsWindow.FindName("btnEdit").IsEnabled = $true
		$vault.ChangeOrderService.UndoEditChangeOrder($mEcoId)
		return $true
	}
	catch
	{
		#[System.Windows.MessageBox]::Show("Item is not accessible for Edit.", "VDS ECO Edit")
		return $false
	}
	return $false
}

function EditEcoProps
{
	$number=$vaultContext.SelectedObject.Label
	$mECO = $vault.ChangeOrderService.GetChangeOrderByNumber($number)
	#again check the ECO accessibility as it might have changed since latest selection
	$_temp = mEcoEditable($mECO.Id) #the function returns an array of 1 - 3 objects, the latest is the coded return $true/$false
	$_temp = $_temp[$_temp.count-1]
	if($_temp -eq $true)
	{		
		#store the current values
		$Global:ItemTitle = $Prop["ItemTitle"].Value
		$Global:ItemDescription = $Prop["ItemDescription"].Value
		$Global:DueDate =  Convert-StringToDateTime $Prop["_XLTN_DUE DATE"].Value
		$Global:CustomValue = $Prop["CustomValue"].Value
		$Global:CustomCompany = $Prop["CustomCompany"].Value
		$Global:ECO = $mECO
		$vault.ChangeOrderService.EditChangeOrder($Global:ECO.Id)
		$global:EcoEditState = $true
		$dsWindow.FindName("btnCancel").IsEnabled = $true
		$dsWindow.FindName("btnSave").IsEnabled = $true
		$dsWindow.FindName("btnEdit").IsEnabled = $false

		#load the Source for combobox and other controls
		$dsWindow.FindName("cmbCustomCompany"). ItemsSource= GetCompanies
		#$vaultContext.Refresh() = $true
	}
	else{
		[System.Windows.MessageBox]::Show("This ECO is accessible for Edit", "VDS ECO Edit")
	}	
}

function CancelEcoUpdate 
{
	Try{
		$global:EcoEditState = $false
		$vault.ChangeOrderService.UndoEditChangeOrder($Global:ECO.Id)
		$dsWindow.FindName("btnCancel").IsEnabled = $False
		$dsWindow.FindName("btnSave").IsEnabled = $False
		$dsWindow.FindName("btnEdit").IsEnabled = $True
	}
	catch{}
	$vaultContext.Refresh() = $true
}

#Update ECO Properties
function EcoUpdate
{
	$dsDiag.Trace(">> Start Save ECO Props ...")
	try
	{
		$data = @{}

		#only store the changed values
		$ValuesChanged = $False
		if($Global:ItemTitle -ne $dsWindow.FindName("txtItemTitle").Text)
		{
			$Global:ItemTitle = $dsWindow.FindName("txtItemTitle").Text
			$ValuesChanged = $true
		}

		if($Global:ItemDescription -ne $dsWindow.FindName("txtItemDescr").Text)
		{
			$Global:ItemDescription = $dsWindow.FindName("txtItemDescr").Text
			$ValuesChanged = $true
		}
		
		if ($Global:CustomValue -ne $dsWindow.FindName("txtCustomValue").Text) 
		{
			$PropDefID = $Prop["CustomValue"].Id
			$data[$PropDefID] = $dsWindow.FindName("txtCustomValue").Text 
			$ValuesChanged = $true
		}

		if ($Global:CustomCompany -ne $dsWindow.FindName("cmbCustomCompany").Text)
		{
			$PropDefID = $Prop["CustomCompany"].Id
			$data[$PropDefID] = $dsWindow.FindName("cmbCustomCompany").Text
			$ValuesChanged = $true
		}

		#only update ECO if property values are changed
		$propValues =@()
		if($data.Count -gt 0)
		{
			#$i = 0
			foreach($d in $data.GetEnumerator()) 
			{
				$PropInst = New-Object Autodesk.Connectivity.WebServices.PropInst		
				$PropInst.EntityId = $Global:ECO.Id
				$PropInst.PropDefId = $d.Key
				$PropInst.Val = $d.Value
				$propValues += $PropInst
			}
		}
		
		if($ValuesChanged){
			#Update ECO Properties
			$vault.ChangeOrderService.UpdateChangeOrder($Global:ECO.Id, #changeOrderId 
													$null, #changeOrderNumber 
													$Global:ItemTitle, #ECO title 
													$Global:ItemDescription, #ECO description 
													$Global:DueDate, #approveDeadline
													$null, #addItemMasterIds
													$null, #delItemMasterIds
													$null, #addAttmtMasterIds
													$null, #delAttmtMasterIds
													$null, #addFileMasterIds
													$null, #delFileMasterIds
													$propValues, #addProperties
													$null, #delPropDefIds
													$null, #addComments
													$null,  #AddEmailNotifications
													$null, #addAssocProperties
													$null,  #delAssocPropIds
													-1, #routingId 
													$null, #addMembers
													$null) #delMembers
			$dsWindow.FindName("btnCancel").IsEnabled = $False
			$dsWindow.FindName("btnSave").IsEnabled = $False
			$dsWindow.FindName("btnEdit").IsEnabled = $True
			$global:EcoEditState = $false
		}
		else
		{
			$dsDiag.Trace("EcoUpdate: nothing changed really, therefore undo item edit and reset UI")
			cancelEcoUpdate
		}
	} #end try
	catch 
	{
	  	$dsDiag.Trace("...Error during edit properties...")
		cancelEcoUpdate
	}
$dsDiag.Trace("...finished Eco Save Props<<")
}

function GetCompanies
{
	$companies = mSearchCustentOfCat "Organisation"
	$companyNames = @()
	$companies | ForEach-Object { $companyNames += $_.Name }
	return $companyNames
}

function Convert-StringToDateTime
{
param
(
[Parameter(Mandatory = $true)]
[String] $DateTimeStr
)

$DateFormatParts = (Get-Culture).DateTimeFormat.ShortDatePattern -split ‘/|-|\.’

$Month_Index = ($DateFormatParts | Select-String -Pattern ‘M’).LineNumber – 1
$Day_Index = ($DateFormatParts | Select-String -Pattern ‘d’).LineNumber – 1
$Year_Index = ($DateFormatParts | Select-String -Pattern ‘y’).LineNumber – 1

$DateTimeParts = $DateTimeStr -split ‘/|-|\.| ‘

$DateTimeParts_LastIndex = $DateTimeParts.Count – 1

$DateTime = [DateTime] $($DateTimeParts[$Month_Index] + ‘/’ + $DateTimeParts[$Day_Index] + ‘/’ + $DateTimeParts[$Year_Index] + ‘ ‘ + $DateTimeParts[3..$DateTimeParts_LastIndex] -join ‘ ‘)

return $DateTime
}
