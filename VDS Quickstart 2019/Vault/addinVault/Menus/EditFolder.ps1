
$vaultContext.ForceRefresh = $true
$folderId=$vaultContext.CurrentSelectionSet[0].Id
$dialog = $dsCommands.GetEditFolderDialog($folderId)

$dialog.Execute()

