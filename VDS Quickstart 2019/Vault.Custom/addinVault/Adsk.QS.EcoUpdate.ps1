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

function EditECOProps 
{
	#buttons visibility
	$dsWindow.FindName("Cancel").IsEnabled = $True
	$dsWindow.FindName("Save").IsEnabled = $True
	$dsWindow.FindName("EditECO").IsEnabled = $False

	#store the current values
	$Global:ItemTitle = $Prop["ItemTitle"].Value
	$Global:ItemDescription = $Prop["ItemDescription"].Value
	$Global:DueDate =  Convert-StringToDateTime $Prop["_XLTN_DUE DATE"].Value
	$Global:CustomValue = $Prop["CustomValue"].Value
	$Global:CustomCompany = $Prop["CustomCompany"].Value

	$number=$vaultContext.SelectedObject.Label
	$Global:ECO = $vault.ChangeOrderService.GetChangeOrderByNumber($number)
	$vault.ChangeOrderService.EditChangeOrder($Global:ECO.Id)

	#load the Source for combobox and other controls
	$dsWindow.FindName("CustomCompany"). ItemsSource= GetCustomObjects "Organisation"
		$vaultContext.Refresh() = $true
}

function CancelECOUpdate 
{
	#buttons visibility
	$dsWindow.FindName("Cancel").IsEnabled = $False
	$dsWindow.FindName("Save").IsEnabled = $False
	$dsWindow.FindName("EditECO").IsEnabled = $True

	$vault.ChangeOrderService.UndoEditChangeOrder($Global:ECO.Id)

	#restore values
	$vaultContext.Refresh() = $true
}

#Update ECO Properties
function ECOUpdate 
{

	$dsWindow.FindName("Cancel").IsEnabled = $False
	$dsWindow.FindName("Save").IsEnabled = $False
	
	$dsDiag.Trace(">> Start Save ECO Props ...")
	
	try
	{
		#Grab The ECO object and set it to edit
		#$number=$vaultContext.SelectedObject.Label
		#$Global:ECO = $vault.ChangeOrderService.GetChangeOrderByNumber($number)
		#$vault.ChangeOrderService.EditChangeOrder($Global:ECO.Id)

		$data = @{}

		#only store the changed values
		$ValuesChanged = $False
		if($Global:ItemTitle -ne $dsWindow.FindName("ItemTitle").Text)
		{
			$Global:ItemTitle = $dsWindow.FindName("ItemTitle").Text
			$ValuesChanged = $true
		}

		if($Global:ItemDescription -ne $dsWindow.FindName("Description").Text)
		{
			$Global:ItemDescription = $dsWindow.FindName("Description").Text
			$ValuesChanged = $true
		}
		
		if ($Global:CustomValue -ne $dsWindow.FindName("CustomValue").Text) 
		{
			$PropDefID = $Prop["CustomValue"].Id
			$data[$PropDefID] = $dsWindow.FindName("CustomValue").Text 
			$ValuesChanged = $true
		}

		if ($Global:CustomCompany -ne $dsWindow.FindName("CustomCompany").Text)
		{
			$PropDefID = $Prop["CustomCompany"].Id
			$data[$PropDefID] = $dsWindow.FindName("CustomCompany").Text
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
		}
		else
		{
			$vault.ChangeOrderService.UndoEditChangeOrder($Global:ECO.Id)
		}
	} #end try
	catch 
		{
	  		$dsDiag.Trace("...Error during edit properties...")
			$dsWindow.FindName("EditECO").IsEnabled = $True
			$vault.ChangeOrderService.UndoEditChangeOrder($Global:ECO.Id)
			$vaultContext.Refresh() = $true
	  	}

	$dsWindow.FindName("EditECO").IsEnabled = $True
	$vaultContext.Refresh() = $true
	$dsDiag.Trace("... ECO Property Edit finished successfully <<")
}

#Get a list of Custom Objects (e.g. Organisation)
function GetCustomObjects([string]$CustomObjectDisplayName)
{
	#$dsWindow.FindName("Cancel").IsEnabled = $False
	#$dsWindow.FindName("Save").IsEnabled = $False
	#$dsWindow.FindName("EditECO").IsEnabled = $True

	$dsDiag.Trace(">> GetCustomObjects")
	$propDefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("CUSTENT")
    $prop = $propDefs | Where-Object{$_.DispName -eq "Custom Object Name"}
	
	$srchConds = New-Object autodesk.Connectivity.WebServices.SrchCond[] 1
	$srchCond = New-Object autodesk.Connectivity.WebServices.SrchCond
	$srchCond.PropDefId = $prop.Id 
	$srchCond.SrchOper = 3
	$srchCond.SrchTxt = $CustomObjectDisplayName
	$srchCond.PropTyp = [Autodesk.Connectivity.WebServices.PropertySearchType]::SingleProperty
	$srchCond.SrchRule = [Autodesk.Connectivity.WebServices.SearchRuleType]::Must
	$srchConds[0] = $srchCond
	$dsDiag.Trace(" search conditions build")	
	$srchSort = New-Object autodesk.Connectivity.WebServices.SrchSort
	$searchStatus = New-Object autodesk.Connectivity.WebServices.SrchStatus
	$bookmark = ""

	$global:customObjects = $vault.CustomEntityService.FindCustomEntitiesBySearchConditions(@($srchCond),@($srchSort),[ref]$bookmark,[ref]$searchStatus)
	$dsDiag.Trace(" search performed. "+ $global:companies.Count+" elements found")	
	$CustEntNames = @()
	$global:customObjects | ForEach-Object { $CustEntNames += $_.Name }
	$dsDiag.Trace("<< GetCustomObjects")
	return $CustEntNames
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
