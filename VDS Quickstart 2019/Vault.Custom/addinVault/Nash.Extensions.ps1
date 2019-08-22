function CreateFolderRecursive($parentId, $node)
{
    $folder = $vault.DocumentServiceExtensions.AddFolderWithCategory($node.Name, $parentId, $false, 21)

    if ($node.Folder)
    {
        foreach($cldFolder in $node.Folder)
        {
            CreateFolderRecursive $folder.Id $cldFolder
        }
    }
}