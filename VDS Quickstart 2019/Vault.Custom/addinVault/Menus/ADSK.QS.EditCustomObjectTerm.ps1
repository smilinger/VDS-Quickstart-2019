$vaultContext.ForceRefresh = $true
$id=$vaultContext.CurrentSelectionSet[0].Id
$dialog = $dsCommands.GetEditCustomObjectDialog($id)

$xamlFile = New-Object CreateObject.WPF.XamlFile "CustomEntityXaml", "%ProgramData%\Autodesk\Vault 2019\Extensions\DataStandard\Vault.Custom\Configuration\ADSK.QS.CustomObjectTerm.xaml"
$dialog.XamlFile = $xamlFile

$result = $dialog.Execute()
$dsDiag.Trace($result)