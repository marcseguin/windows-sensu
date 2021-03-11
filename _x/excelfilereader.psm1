

#This script can be used to convert information saved inside the excel sheet 
#to script usable objects
#
#@author NKO

### main

    $lastPath = $null
    $configuration = $null

    $CONFIGURATION_SHEET_NAME = "Script-Configuration"
    $CONFIGURATION_OPTION_GROUP_COLUMN = 1
    $CONFIGURATION_OPTION_COLUMN = 2
    $CONFIGURATION_VALUE_COLUMN = 4
    $CONFIGURATION_START_ROW = 2

    if ($null -ne (get-module|where-object {$_.name -eq "psexcel"})){
        remove-module psexcel
    }
    Import-Module psexcel

Function Read-ConfigurationSheet($path){
    $log4n.Debug("Reading Configuration from file $path ...")
    $excelFile = New-Excel -Path $path
    $workBook = $excelFile | Get-Workbook
    $configurationWorksheet = $workBook | Get-Worksheet -Name $CONFIGURATION_SHEET_NAME

    $optionGroups = @{}
    for($i = $CONFIGURATION_START_ROW; $i -le $configurationWorksheet.Dimension.Rows; $i++ ){
        $currentOptionGroup = $configurationWorksheet.Cells.Item($i, $CONFIGURATION_OPTION_GROUP_COLUMN).text
        $currentOption = $configurationWorksheet.Cells.Item($i, $CONFIGURATION_OPTION_COLUMN).text
        $currentOptionValue = $configurationWorksheet.Cells.Item($i, $CONFIGURATION_VALUE_COLUMN).text

        if(-not ($optionGroups.ContainsKey($currentOptionGroup))){
            $optionGroups.Add($currentOptionGroup, @{})
        } 

        $optionGroups[$currentOptionGroup].Add($currentOption, $currentOptionValue)
    }
    return $optionGroups
}

Function Read-TagSheet($path){
    $log4n.Debug("Reading Tags Definitions from file $path ...")
    if(-not ($lastPath -eq $path)){
        $configuration = Read-ConfigurationSheet $path
    }

    $TAG_SHEET_NAME = $configuration["General"]["TAG_SHEET_NAME"]
    $TAG_NAME_COLUMN = $configuration["Tag"]["TAG_NAME_COLUMN"] -as [int]
    $TAG_DESCRIPTION_COLUMN = $configuration["Tag"]["TAG_DESCRIPTION_COLUMN"] -as [int]
    $TAG_CATEGORY_COLUMN = $configuration["Tag"]["TAG_CATEGORY_COLUMN"] -as [int]
    $TAG_DATA_START = $configuration["Tag"]["TAG_DATA_START"] -as [int]

    $excelFile = New-Excel -Path $path
    $workBook = $excelFile | Get-Workbook
    $tag_worksheet = $workBook | Get-Worksheet -Name $TAG_SHEET_NAME

    $tags = @()
    for($i = $TAG_DATA_START; $i -le $tag_worksheet.Dimension.Rows; $i++){
        $curTag = @{}
        
        $curTag.Add("name", $tag_worksheet.Cells.Item($i, $TAG_NAME_COLUMN).text)
        $curTag.Add("description", $tag_worksheet.Cells.Item($i, $TAG_DESCRIPTION_COLUMN).text)
        $curTag.Add("group", $tag_worksheet.Cells.Item($i, $TAG_CATEGORY_COLUMN).text)

        $tags += $curTag
    }

    $lastPath = $path
    return $tags
}

Function Read-TagCategorySheet($path){
    $log4n.Debug("Reading Tag Category Definitions from file $path ...")
    if(-not($lastPath -eq $path)){
        $configuration = Read-ConfigurationSheet $path
    }

    $TAG_CATEGORY_SHEET_NAME = $configuration["General"]["TAG_CATEGORY_SHEET_NAME"] 
    $TAG_CATEGORY_NAME_COLUMN = $configuration["Tag Category"]["TAG_CATEGORY_NAME_COLUMN"] -as [int]
    $TAG_CATEGORY_DESCRIPTION_COLUMN = $configuration["Tag Category"]["TAG_CATEGORY_DESCRIPTION_COLUMN"] -as [int]
    $TAG_CATEGORY_POLICY_COLUMN = $configuration["Tag Category"]["TAG_CATEGORY_POLICY_COLUMN"] -as [int]
    $TAG_CATEGORY_VALID_FOR_START_COLUMN = $configuration["Tag Category"]["TAG_CATEGORY_VALID_FOR_START_COLUMN"] -as [int]
    $TAG_CATEGORY_DATA_START = $configuration["Tag Category"]["TAG_CATEGORY_DATA_START"] -as [int]
    $TAG_CATEGORY_VALID_FOR_NAME_ROW = $configuration["Tag Category"]["TAG_CATEGORY_VALID_FOR_NAME_ROW"] -as [int]

    $excelFile = New-Excel -Path $path
    $workBook = $excelFile | Get-Workbook
    $tag_category_worksheet = $workBook | Get-Worksheet -Name $TAG_CATEGORY_SHEET_NAME

    $tag_categories = @()
    for($i = $TAG_CATEGORY_DATA_START; $i -le $tag_category_worksheet.Dimension.Rows; $i++ ){
        $curTagCategory = @{}

        $curTagCategory.Add("name", $tag_category_worksheet.Cells.Item($i, $TAG_CATEGORY_NAME_COLUMN).text)
        $curTagCategory.Add("description", $tag_category_worksheet.Cells.Item($i, $TAG_CATEGORY_DESCRIPTION_COLUMN).text)
        $curTagCategory.Add("policy", $tag_category_worksheet.Cells.Item($i, $TAG_CATEGORY_POLICY_COLUMN).text)

        $tagCategoryValidFor = @()
        for($k = $TAG_CATEGORY_VALID_FOR_START_COLUMN; $k -le $tag_category_worksheet.Dimension.Columns; $k++){
            $curValidForField = $tag_category_worksheet.Cells.Item($i, $k).text
            if (($curValidForField -eq "x") -or ($curValidForField -eq "X")){
                $tagCategoryValidFor += $tag_category_worksheet.Cells.Item($TAG_CATEGORY_VALID_FOR_NAME_ROW, $k).text
            }
        }
        $curTagCategory.Add("validFor", $tagCategoryValidFor)

        $tag_categories += $curTagCategory
    }

    $lastPath = $path
    return $tag_categories
}


Function Read-ObjectToTagConfigurationSheet($path, $sheetName, $nameColumn, $tagStartColumn, $dataStartRow, $tagNameRow){
    $excelFile = New-Excel -Path $path
    $workBook = $excelFile | Get-Workbook
    $objectWorksheet = $workBook | Get-Worksheet -Name $sheetName

    $vmObjects = @()
    for($i = $dataStartRow; $i -le $objectWorksheet.Dimension.Rows; $i++){
        $curVmObject = @{}

        $curVmObject.Add("name", $objectWorksheet.Cells.Item($i, $nameColumn).text)

        $curVmObjectTags = @()
        for($k = $tagStartColumn; $k -le $objectWorksheet.Dimension.Columns; $k++){
            $curVmObjectHasTagCell = $objectWorksheet.Cells.Item($i, $k).text
            if(($curVmObjectHasTagCell -eq "x") -or ($curVmObjectHasTagCell -eq "X")){
                $curVmObjectTags += $objectWorksheet.Cells.Item($tagNameRow, $k).text
            } 
        }
        $curVmObject.Add("tags", $curVmObjectTags)
        if($curVmObject -ne $null){
            $vmObjects += $curVmObject
        }
    }

    return $vmObjects
}

function Read-VMConfigurationSheet($path){
    $log4n.Debug("Reading VM Tags from file $path ...")
    if(-not($lastPath -eq $path)){
        $configuration = Read-ConfigurationSheet $path
    }

    $VM_SHEET_NAME = $configuration["General"]["VM_CONFIGURATION_SHEET_NAME"]
    $VM_NAME_COLUMN = $configuration["VM Configuration"]["VM_NAME_COLUMN"] -as [int]
    $TAG_START_COLUMN = $configuration["VM Configuration"]["TAG_START_COLUMN"] -as [int]
    $VM_DATA_START = $configuration["VM Configuration"]["VM_DATA_START"] -as [int]
    $TAG_NAME_ROW = $configuration["VM Configuration"]["TAG_NAME_ROW"] -as [int]

    $lastPath = $path
    return Read-ObjectToTagConfigurationSheet $path $VM_SHEET_NAME $VM_NAME_COLUMN $TAG_START_COLUMN $VM_DATA_START $TAG_NAME_ROW
}

function Read-DataStoreConfigurationSheet($path){
    $log4n.Debug("Reading Datastore Tags from file $path ...")
    if(-not ($lastPath -eq $path)){
        $configuration = Read-ConfigurationSheet $path
    }

    $DATA_STORE_SHEET_NAME = $configuration["General"]["DATA_STORE_CONFIGURATION_SHEET_NAME"]
    $DATA_STORE_NAME_COLUMN = $configuration["Data Store Configuration"]["DATA_STORE_NAME_COLUMN"] -as [int]
    $TAG_START_COLUMN = $configuration["Data Store Configuration"]["TAG_START_COLUMN"] -as [int]
    $DATA_STORE_DATA_START = $configuration["Data Store Configuration"]["DATA_STORE_DATA_START"] -as [int]
    $TAG_NAME_ROW = $configuration["Data Store Configuration"]["TAG_NAME_ROW"] -as [int]

    $lastPath = $path
    return Read-ObjectToTagConfigurationSheet $path $DATA_STORE_SHEET_NAME $DATA_STORE_NAME_COLUMN $TAG_START_COLUMN $DATA_STORE_DATA_START $TAG_NAME_ROW
}

function Read-ClusterConfigurationSheet($path){
    $log4n.Debug("Reading Cluster Tags from file $path ...")
    if(-not ($lastPath -eq $path)){
        $configuration = Read-ConfigurationSheet $path
    }

    $CLUSTER_SHEET_NAME = $configuration["General"]["CLUSTER_CONFIGURATION_SHEET_NAME"]
    $CLUSTER_NAME_COLUMN = $configuration["Cluster Configuration"]["CLUSTER_NAME_COLUMN"] -as [int]
    $TAG_START_COLUMN = $configuration["Cluster Configuration"]["TAG_START_COLUMN"] -as [int]
    $CLUSTER_DATA_START = $configuration["Cluster Configuration"]["CLUSTER_DATA_START"] -as [int]
    $TAG_NAME_ROW = $configuration["Cluster Configuration"]["TAG_NAME_ROW"] -as [int]

    $lastPath = $path
    return Read-ObjectToTagConfigurationSheet $path $CLUSTER_SHEET_NAME $CLUSTER_NAME_COLUMN $TAG_START_COLUMN $CLUSTER_DATA_START $TAG_NAME_ROW
}

function Read-HostConfigurationSheet($path){
    $log4n.Debug("Reading Host Tags from file $path ...")
    if(-not($lastPath -eq $path)){
        $configuration = Read-ConfigurationSheet $path
    }

    $HOST_SHEET_NAME = $configuration["General"]["HOST_CONFIGURATION_SHEET_NAME"]
    $HOST_NAME_COLUMN = $configuration["Host Configuration"]["HOST_NAME_COLUMN"] -as [int]
    $TAG_START_COLUMN = $configuration["Host Configuration"]["TAG_START_COLUMN"] -as [int]
    $HOST_DATA_START = $configuration["Host Configuration"]["HOST_DATA_START"] -as [int]
    $TAG_NAME_ROW = $configuration["Host Configuration"]["TAG_NAME_ROW"] -as [int]

    $lastPath = $path
    return Read-ObjectToTagConfigurationSheet $path $HOST_SHEET_NAME $HOST_NAME_COLUMN $TAG_START_COLUMN $HOST_DATA_START $TAG_NAME_ROW
}

function Read-FolderConfigurationSheet($path){
    $log4n.Debug("Reading Folder Tags from file $path ...")
    if(-not ($lastPath -eq $path)){
        $configuration = Read-ConfigurationSheet $path
    }

    $FOLDER_SHEET_NAME = $configuration["General"]["FOLDER_CONFIGURATION_SHEET_NAME"]
    $FOLDER_NAME_COLUMN = $configuration["Folder Configuration"]["FOLDER_NAME_COLUMN"] -as [int]
    $TAG_START_COLUMN = $configuration["Folder Configuration"]["TAG_START_COLUMN"] -as [int]
    $FOLDER_DATA_START = $configuration["Folder Configuration"]["FOLDER_DATA_START"] -as [int]
    $TAG_NAME_ROW = $configuration["Folder Configuration"]["TAG_NAME_ROW"] -as [int]

    $lastPath = $path
    return Read-ObjectToTagConfigurationSheet $path $FOLDER_SHEET_NAME $FOLDER_NAME_COLUMN $TAG_START_COLUMN $FOLDER_DATA_START $TAG_NAME_ROW
}