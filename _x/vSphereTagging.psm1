
Function Update-Tags($tags) {
    $Log4n.info("Updating tags")

    foreach ($curTag in $tags) {
        New-Tag -Name $curTag["name"] -Description $curTag["description"] -Category $curTag["group"]
    }

    $Log4n.info("Updating tags completed")
}

Function Update-TagCategories($tagCategories) {
    $Log4n.info("Updating tag categories")

    Get-TagCategory | Remove-TagCategory

    foreach ($curTagCategory in $tagCategories) {
        New-TagCategory -Name $curTagCategory["name"] -Description $curTagCategory["description"] -Cardinality $curTagCategory["policy"] -EntityType $curTagCategory["validFor"]
    }
    
    $Log4n.info("Updating tag categories completed")
}


Function Update-Objects($objects, $typeName) {
    $Log4n.info("Updating tags of type $typeName")

    foreach ($curVmObject in $vmObjects) {
        foreach ($curTag in $curVmObject["tags"]) {
            New-TagAssignment -Tag $curTag -Entity $curVmObject["name"]
        }
    }
}

Function New-ObjectDeltaEntry ($action, $type, $value) {
    $objectProperties = @{
        "Action" = $action
        "Type"   = $type
        "Value"  = $value
    }

    return New-Object -TypeName psobject -Property $objectProperties
}

Function Find-TagCategoryDelta($tagCategories) {
    $tagCategoryActions = @()
    
    $tagCategoryNames = ($tagCategories).name
    $definedTagCategoryNames = (Get-TagCategory).Name

    foreach ($curTagCategoryName in $tagCategoryNames) {
        if ($definedTagCategoryNames.Contains($curTagCategoryName)) {
            $tagCategoryActions += New-ObjectDeltaEntry "O" "Tag Category" $curTagCategoryName
        }
        else {
            $tagCategoryActions += New-ObjectDeltaEntry "+" "Tag Category" $curTagCategoryName
        }
    }

    foreach ($curDefinedTagCategoryName in $definedTagCategoryNames) {
        if (-not ($tagCategoryNames.Contains($curDefinedTagCategoryName))) {
            $tagCategoryActions += New-ObjectDeltaEntry "-" "Tag Category" $curDefinedTagCategoryName
        }
    }

    return $tagCategoryActions
}

Function Find-ObjectDelta($objects) {
    $objectActions = @()
    $currentTagAssignmets = Get-TagAssignment
    $currentTagAssignmentsByObjectName = @{ }

    foreach ($currentTagAssignment in $currentTagAssignmets) {
        $currentTagName = $currentTagAssignment.Tag.Name
        $currentObjectName = $currentTagAssignment.Entity.Name

        if (-not($currentTagAssignmentsByObjectName.ContainsKey($currentObjectName))) {
            $currentTagAssignmentsByObjectName.Add($currentObjectName, @())
        }
        $currentTagAssignmentsByObjectName[$currentObjectName] += $currentTagName
    }
   
    foreach ($curObject in $objects) {
        $objectName = $curObject["name"]
        if ($currentTagAssignmentsByObjectName.ContainsKey($objectName)) {
            $currentTags = $currentTagAssignmentsByObjectName[$objectName]
            #tag assignments already existing. Search for changes
            foreach ($curTagAssignment in $currentTags) {
                if ($curObject["tags"].Contains($curTagAssignment)) {
                    $objectActions += New-ObjectDeltaEntry "O" "Tag Assignment" "$objectName -> $curTagAssignment"
                }
                else {
                    $objectActions += New-ObjectDeltaEntry "-" "Tag Assignment" "$objectName -> $curTagAssignment"
                }
            }

            foreach ($curTag in $curObject["tags"]) {
                if (-not($currentTags.Contains($curTag))) {
                    $objectActions += New-ObjectDeltaEntry "+" "Tag Assignment" "$objectName -> $curTag"
                }
            }
        }
        else {
            # no tag assignments currently existing 
            foreach ($curTag in $curObject["tags"]) {
                $objectActions += New-ObjectDeltaEntry "+" "Tag Assignment" "$objectName -> $curTag"
            }
        }
    }
    $objectNames = ($objects).name

    foreach ($currentObjectName in $currentTagAssignmentsByObjectName.Keys) {
        $currentTagName = $currentTagAssignmentsByObjectName[$currentObjectName]

        if (-not($objectNames.Contains($currentObjectName))) {
            $objectActions += New-ObjectDeltaEntry "-" "Tag Assignment" "$currentObjectName -> $currentTagName"
        }
    }
    return $objectActions
}

Function Find-TagDelta($tags) {
    $tagActions = @()
    $definedTags = Get-Tag

    foreach ($curTag in $tags) {
        $found = $false

        foreach ($curDefinedTag in $definedTags) {
            if (($curDefinedTag.Name -eq $curTag["name"]) -and ($curDefinedTag.Category.Name -eq $curTag["group"])) {
                $found = $true
                $tagGroup = $curTag["group"]
                $tagName = $curTag["name"]
                $tagActions += New-ObjectDeltaEntry "O" "Tag" "$tagGroup/$tagName"

                break
            }
        }
        if (-not $found) {
            $tagGroup = $curTag["group"]
            $tagName = $curTag["name"]
            $tagActions += New-ObjectDeltaEntry "+" "Tag" "$tagGroup/$tagName"
        }
    }

    foreach ($curDefinedTag in $definedTags) {
        $found = $false

        foreach ($curTag in $tags) {
            if (($curTag["name"] -eq $curDefinedTag.Name) -and ($curTag["group"] -eq $curDefinedTag.Category)) {
                $found = $true
                break
            }
        }
        
        if (-not $found) {
            $categoryName = $curDefinedTag.Category
            $tagName = $curDefinedTag.Name
            $tagActions += New-ObjectDeltaEntry "-" "Tag" "$categoryName/$tagName"
        }
    }
    
    return $tagActions
}



#TODO: Need to get implemented 
Function Show-DeltaInformation($tagCategories, $tags, $tagAssignments) {
    $tagCategoryDelta = Find-TagCategoryDelta $tagCategories
    $tagDelta = Find-TagDelta $tags
    $tagAssingmentDelta = Find-ObjectDelta $tagAssignments

    ($tagCategoryDelta + $tagDelta + $tagAssingmentDelta) | Format-Table -Property Action, Type, Value

}



Function Update-vSphereTags($excelFilePath, $force) {
    $tagCategories = Read-TagCategorySheet $excelFilePath
    $tags = Read-TagSheet $excelFilePath
    $vmTags = Read-VMConfigurationSheet $excelFilePath
    $dataStoreTags = Read-DataStoreConfigurationSheet $excelFilePath
    $clusterTags = Read-ClusterConfigurationSheet $excelFilePath
    $hostTags = Read-HostConfigurationSheet $excelFilePath
    $folderTags = Read-FolderConfigurationSheet $excelFilePath    
    $tagAssignments = ($vmTags + $dataStoreTags + $clusterTags + $folderTags) 
    Show-DeltaInformation $tagCategories $tags $tagAssignments
    if ($force -ne $true) {
        $reply = Read-Host -Prompt "Do you want to apply the changes?[y/n]"
        if ( -not($reply -match "[yY]") ) { 
            exit 
        }
    }
    Update-TagCategories $tagCategories
    Update-Tags $tags
    Update-Objects $vmTags "VM"
    Update-Objects $dataStoreTags "DataStore"
    Update-Objects $clusterTags "Cluster"
    Update-Objects $hostTags "Host"
    Update-Objects $folderTags "folder"
    
}

Function Update-HostTags($excelFilePath, $force) {
    $tagCategories = Read-TagCategorySheet $excelFilePath
    $tags = Read-TagSheet $excelFilePath
    $hostTags = Read-HostConfigurationSheet $excelFilePath
    Show-DeltaInformation $tagCategories $tags $tagAssignments
    if ($force -ne $true) {
        $reply = Read-Host -Prompt "Do you want to apply the changes?[y/n]"
        if ( -not($reply -match "[yY]") ) { 
            exit 
        }
    }
    Update-TagCategories $tagCategories
    Update-Tags $tags
    Update-Objects $hostTags "Host"
}

#### MAIN ####

if ($null -ne (get-module|where-object {$_.name -eq "support"})){
    remove-module support
}
Import-Module ./support.psm1

if ($null -ne (get-module|where-object {$_.name -eq "excelfilereader"})){
    remove-module excelfilereader
}
Import-Module ./excelfilereader.psm1


