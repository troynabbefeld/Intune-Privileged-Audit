# trnabbefeld@hbs.net - 2024-09-30, 2024-09-30
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

Connect-MgGraph -Scope "DeviceManagementConfiguration.Read.All", "DeviceManagementApps.Read.All" -NoWelcome

# Define the company name and current date
$companyName = (Get-MgOrganization).DisplayName
$date = (Get-Date -Format "yyyy-MM-dd")

# Get the desktop path of the current user
$desktopPath = [Environment]::GetFolderPath("Desktop")

# Create the folder name
$folderName = "$companyName Intune Reports $date"

# Combine the desktop path and folder name
$parentFolder = Join-Path -Path $desktopPath -ChildPath $folderName

function New-Report {
    param (
        [string]$ReportName
    )
    
    return @{
        ReportName = $ReportName
        Data       = @()  # Initialize the Data property as an array
    }
}

# Function to add data to a report
function Add-DataToReport {
    param (
        [hashtable]$report,
        [PSCustomObject]$data
    )
    
    $report.Data += $data
}

function New-Folder {
    param(
        [string]$folderPath
    )
    if (-not (Test-Path -Path $folderPath)) {
        New-Item -Path $folderPath -ItemType Directory | Out-Null
    }
}

#Initialize Variable for all Intune Settings

$allSettings = Invoke-MgGraphRequest GET -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationSettings'


# Function to check for choiceSettingValue, simpleSettingCollectionValue, and simpleSettingValue, including children
function Get-SettingValue {
    param (
        [PSCustomObject]$settingInstance
    )
    
    # Initialize an array to hold the setting data
    $settingData = @()

    $specificSetting = $allSettings.value | Where-Object { $_.id -eq $settingInstance.settingDefinitionId }
    $settingDisplayName = $specificSetting.displayName

    try {
        # Handle deviceManagementConfigurationGroupSettingCollectionInstance
        if ($settingInstance.'@odata.type' -eq "#microsoft.graph.deviceManagementConfigurationGroupSettingCollectionInstance") {
            $groupSettings = $settingInstance.groupSettingCollectionValue
            foreach ($groupSetting in $groupSettings) {
                if ($groupSetting.children) {
                    # Process each child in the group settings
                    foreach ($childSetting in $groupSetting.children) {
                        $settingData += Get-SettingValue -settingInstance $childSetting
                    }
                }
            }
        }

        # Attempt to get choiceSettingValue
        $choiceValue = $settingInstance.choiceSettingValue.value
        if ($choiceValue) {
            $value = $specificSetting.options | Where-Object { $_.itemId -eq $choiceValue }
            $settingData += [PSCustomObject]@{
                Name        = $settingDisplayName
                Value       = $value.displayName
                Type        = "Choice"
                Description = $specificSetting.description
            }
        }
        $children = $settingInstance.choiceSettingValue.children
        if ($children) {
            $children | ForEach-Object {
                $settingData += Get-SettingValue -settingInstance $_
            }
        }
    }
    catch {}

    try {
        # Attempt to get simpleSettingCollectionValue
        $collectionValue = $settingInstance.simpleSettingCollectionValue.value
        if ($collectionValue) {
            $settingData += [PSCustomObject]@{
                Name        = $settingDisplayName
                Value       = $collectionValue
                Type        = "Collection"
                Description = $specificSetting.description
            }
        }
        $children = $settingInstance.simpleSettingCollectionValue.children
        if ($children) {
            $children | ForEach-Object {
                $settingData += Get-SettingValue -settingInstance $_
            }
        }
    }
    catch {}

    try {
        # Attempt to get simpleSettingValue
        $simpleValue = $settingInstance.simpleSettingValue.value
        if ($simpleValue) {
            $settingData += [PSCustomObject]@{
                Name        = $settingDisplayName
                Value       = $simpleValue
                Type        = "Simple"
                Description = $specificSetting.description
            }
        }
        $children = $settingInstance.simpleSettingValue.children
        if ($children) {
            $children | ForEach-Object {
                $settingData += Get-SettingValue -settingInstance $_
            }
        }
    }
    catch {}

    return $settingData
}
function Get-SettingCatalogs {
    # Log
    write-host "Fetching Setting Catalogs"

    # Create ConfigurationProfiles folder within the parent folder
    $SettingCatalogFolder = Join-Path -Path $parentFolder -ChildPath "SettingCatalogs"
    New-Folder -folderPath $SettingCatalogFolder

    # Create Security Baselines folder within the parent folder
    $SecurityBaselinesFolder = Join-Path -Path $ParentFolder -ChildPath "SecurityBaselines"
    New-Folder -folderPath $SecurityBaselinesFolder

    #Get Setting Catalogs (Limit 100)
    $settingCatalogs = Invoke-MgGraphRequest GET -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationpolicies?$top=100'

    # Initialize row counter
    $rowIndex = 1

    # Setting Catalogs
    $settingCatalogs.Value | ForEach-Object {
        # Log
        $Id = $_.Id
        $name = $_.name
        $_.templateReference.TemplateFamily
        write-host "Fetching Setting Catalog $name ($Id)"

        #Get Settings (Limit 1000)
        $settings = Invoke-MgGraphRequest GET -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$Id')/settings?`$top=1000"

        # Collect settings for the current policy with row numbers
        $data = @()
        $settings.value | ForEach-Object {
            $settingInstance = $_.settingInstance

            # Get the settings data and add to the policy's settings collection with a row number as the first column
            $settingValues = Get-SettingValue -settingInstance $settingInstance
            foreach ($value in $settingValues) {
                # Add Row property as the first column
                $indexedValue = [PSCustomObject]@{
                    Row         = $rowIndex
                    Name        = $value.Name
                    Value       = $value.Value
                    Description = $value.Description
                    Type        = $value.Type
                }
                $data += $indexedValue
                $rowIndex++
            }
        }

        # Display the policy settings in Out-GridView
        $data | Out-GridView -Title $_.Name

        # Save report to csv file
        if ($_.templateReference.TemplateFamily -ne "baseline") {
            $csvFilePath = Join-Path -Path $SettingCatalogFolder -ChildPath "$($name).csv"
        }
        else {
            $csvFilePath = Join-Path -Path $SecurityBaselinesFolder -ChildPath "$($name).csv"
        }
        $data | Export-Csv -Path $csvFilePath 
        
        # Reset row index for each new policy if desired
        $rowIndex = 1
    }
}

# Function to expand nested properties
function Expand-NestedProperties {
    param (
        [hashtable]$settings
    )

    $data = @()

    # Define properties to exclude
    $propertiesToExclude = @(
        'deviceManagementApplicabilityRuleOSEdition',
        'lastModifiedDateTime',
        '@odata.type', # Exclude @odata.type
        'createdDateTime',
        'supportsScopeTags',
        'deviceManagementApplicabilityRuleOSVersion',
        'id',
        'description',
        'roleScopeTagIds',
        'deviceManagementApplicabilityRuleDeviceMode',
        'version',
        'displayName'
    )

    foreach ($key in $settings.Keys) {
        if ($propertiesToExclude -notcontains $key) {
            $value = $settings[$key]
            if ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) {
                # If the value is a collection, expand it
                foreach ($item in $value) {
                    if ($item -is [hashtable]) {
                        # If the item is a hashtable, exclude any @odata.type properties
                        $nestedProperties = $item.Keys | Where-Object { $_ -ne '@odata.type' }
                        foreach ($nestedKey in $nestedProperties) {
                            $data += [PSCustomObject]@{
                                SettingName = "$key.$nestedKey"  # Use a dot notation to indicate the hierarchy
                                Value       = $item[$nestedKey]
                            }
                        }
                    }
                    else {
                        # For non-hashtable items, add a row directly
                        $data += [PSCustomObject]@{
                            SettingName = $key
                            Value       = $item
                        }
                    }
                }
            }
            else {
                # Regular property, add to expanded settings
                $data += [PSCustomObject]@{
                    SettingName = $key
                    Value       = $value
                }
            }
        }
    }
    return $data
}

function Get-ConfigurationPolicies {
    # Log
    write-host "Fetching Configuration Policies"

    # Create ConfigurationProfiles folder within the parent folder
    $ConfigurationPoliciesFolder = Join-Path -Path $parentFolder -ChildPath "ConfigurationPolicies"
    New-Folder -folderPath $ConfigurationPoliciesFolder

    # Get Configuration Policies (Limit 100)
    $configurationPolicies = Invoke-MgGraphRequest GET -Uri 'https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?$top=100'

    # Loop through each configuration policy
    $configurationPolicies.value | ForEach-Object {
        $Id = $_.Id
        $name = $_.displayName

        #Log
        write-host "Fetching Configuration Policy $name ($Id)"
        
        # Get the base object which contains the settings
        $settings = $_

        # Expand settings
        $data = Expand-NestedProperties -settings $settings

        # Display the current policy settings in Out-GridView
        $data | Out-GridView -Title $name

        # Save report to csv file
        $csvFilePath = Join-Path -Path $ConfigurationPoliciesFolder -ChildPath "$($name).csv"
        $data | Export-Csv -Path $csvFilePath 
    }
}

function Get-AdmxPolicies {
    # Log
    write-host "Fetching Administrative Templates"

    # Create AdmxTemplates folder within the parent folder
    $AdmxTemplatesFolder = Join-Path -Path $ParentFolder -ChildPath "AdmxTemplates"
    New-Folder -folderPath $AdmxTemplatesFolder

    # Get Administrative Templates (Limit 100)
    $admxPolicies = Invoke-MgGraphRequest GET -Uri 'https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations?$top=100'

    $admxPolicies.Value | ForEach-Object {
        $data = @()
        $Id = $_.Id
        $name = $_.displayName

        # Log
        write-host "Fetching Administrative Policy $name ($Id)"

        $settings = Invoke-MgGraphRequest GET -Uri "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations('$id')/definitionValues"
        $settings.value | ForEach-Object {
            $settingId = $_.ID
            $setting = Invoke-MgGraphRequest GET -Uri "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations('$id')/definitionValues('$settingId')"
            $settingInfo = Invoke-MgGraphRequest GET -Uri "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations('$id')/definitionValues('$settingId')/definition"
            $settingValue = Invoke-MgGraphRequest GET -Uri "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations('$id')/definitionValues('$settingId')/PresentationValues"
            $policyResult = [PSCustomObject]@{
                Name    = $settingInfo.displayName
                Enabled = $setting.enabled
                Label   = $null
                value   = $null
            }
            $data += $policyResult
            $settingValue.Value | ForEach-Object {
                try {
                    $newId = $_.id
                    $value = $_.value
                    $newSetting = Invoke-MgGraphRequest GET -Uri "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations('$id')/definitionValues('$settingId')/PresentationValues/$newid/presentation"
                    try {
                        $newValue = $newSetting.items | Where-Object { $_.value -eq $value }
                        $value = $newValue.displayName
                    }
                    catch {
                    }
                    $label = $newSetting.label
                }
                catch {
                    $value = $null
                
                }
                $policyResult = [PSCustomObject]@{
                    Name    = $settingInfo.displayName
                    Enabled = $null
                    Label   = $label
                    value   = $value
                }
                $data += $policyResult
            }
        }

        $data | Out-GridView -Title $name

        # Save report to csv file
        $csvFilePath = Join-Path -Path $AdmxTemplatesFolder -ChildPath "$($name).csv"
        $data | Export-Csv -Path $csvFilePath 
    }
}

function Get-CompliancePolicies {

    # Log
    write-host "Fetching Compliance Policies"

    # Create AdmxTemplates folder within the parent folder
    $CompliancePoliciesFolder = Join-Path -Path $ParentFolder -ChildPath "CompliancePolicies"
    New-Folder -folderPath $CompliancePoliciesFolder

    # Retrieve all device compliance policies
    $compliancePolicies = Get-MgDeviceManagementDeviceCompliancePolicy -All

    # Loop through each compliance policy to get its settings using ForEach-Object
    $compliancePolicies | ForEach-Object {
        $Id = $_.Id
        $name = $_.displayName

        # Log
        write-host "Fetching Complaince Policy $name ($Id)"

        # Initialize an array to hold settings for the current policy
        $data = @()

        # Extract additional properties
        $additionalProperties = $_.AdditionalProperties

        # Loop through each setting in AdditionalProperties (which is a Dictionary)
        $additionalProperties.GetEnumerator() | ForEach-Object {
            # Exclude the @odata.type property
            if ($_.Key -ne '@odata.type') {
                # Create a hashtable to hold the key-value pairs
                $setting = [PSCustomObject]@{
                    SettingName = $_.Key   # Key is the setting name
                    Value       = $_.Value # Value is the setting value
                }

                # Add the setting to the current policy settings array
                $data += $setting
            }
        }

        # Sort settings alphabetically by SettingName
        $data = $data | Sort-Object SettingName

        # Display the current policy settings in Out-GridView
        $data | Out-GridView -Title $name

        # Save report to csv file
        $csvFilePath = Join-Path -Path $CompliancePoliciesFolder -ChildPath "$($name).csv"
        $data | Export-Csv -Path $csvFilePath 
    }
}

function Get-Applications {
    write-host "Fetching Applications"

    # Create AdmxTemplates folder within the parent folder
    $ApplicationsFolder = Join-Path -Path $ParentFolder -ChildPath "Applications"
    New-Folder -folderPath $ApplicationsFolder

    # Get Applications (Limit 100)
    $applications = Invoke-MgGraphRequest GET -Uri 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$top=100&$filter=(microsoft.graph.managedApp/appAvailability%20eq%20null%20or%20microsoft.graph.managedApp/appAvailability%20eq%20%27lineOfBusiness%27%20or%20isAssigned%20eq%20true)&$orderby=displayName&'

    $applications.value | ForEach-Object {
        $Id = $_.Id
        $name = $_.displayName
        write-host "Fetching Application $name ($Id)"

        # Initialize an empty array to hold the application details in row format
        $data = @()
    
        # Check if '@odata.type' is '#microsoft.graph.win32LobApp'
        if ($_.'@odata.type' -eq "#microsoft.graph.win32LobApp") {
            # Populate the array with the specific application details for Win32 LOB Apps
            $data += [PSCustomObject]@{ Name = "App Name"; Value = $_.displayName }
            $data += [PSCustomObject]@{ Name = "File Name"; Value = $_.fileName }
            $data += [PSCustomObject]@{ Name = "Setup File Path"; Value = $_.setupFilePath }
            $data += [PSCustomObject]@{ Name = "Install Command Line"; Value = $_.installCommandLine }
            $data += [PSCustomObject]@{ Name = "Uninstall Command Line"; Value = $_.uninstallCommandLine }
            $data += [PSCustomObject]@{ Name = "Minimum Supported Windows Release"; Value = $_.minimumSupportedWindowsRelease }
            $data += [PSCustomObject]@{ Name = "Applicable Architectures"; Value = $_.applicableArchitectures }
            
            # Add detection rules as separate rows
            $_.detectionRules | ForEach-Object {
                if ($_.'@odata.type' -eq "#microsoft.graph.win32LobAppFileSystemDetection") {
                    $data += [PSCustomObject]@{ Name = "Detection Rule Type"; Value = "File" }
                    $data += [PSCustomObject]@{ Name = "Path"; Value = $_.path }
                    $data += [PSCustomObject]@{ Name = "File/Folder Name"; Value = $_.fileOrFolderName }
                    $data += [PSCustomObject]@{ Name = "Detection Method"; Value = $_.detectionType }
                }
                elseif ($_.'@odata.type' -eq "#microsoft.graph.win32LobAppRegistryDetection") {
                    $data += [PSCustomObject]@{ Name = "Detection Rule Type"; Value = "Registry" }
                    $data += [PSCustomObject]@{ Name = "Path"; Value = "$($_.keyPath)\$($_.valueName)" }
                    $data += [PSCustomObject]@{ Name = "Detection Method"; Value = $_.detectionType }
                    if ($_.detectionValue -ne "") {
                        $data += [PSCustomObject]@{ Name = "Operator "; Value = $_.operator }
                        $data += [PSCustomObject]@{ Name = "Value "; Value = $_.detectionValue }
                    }
                }
                elseif ($_.'@odata.type' -eq "#microsoft.graph.win32LobAppProductCodeDetection") {
                    $data += [PSCustomObject]@{ Name = "Detection Rule Type"; Value = "MSI" }
                    $data += [PSCustomObject]@{ Name = "Product Code"; Value = $_.productCode }
                }
            }
        }
        # Check if '@odata.type' is '#microsoft.graph.windowsMobileMSI'
        elseif ($_.'@odata.type' -eq "#microsoft.graph.windowsMobileMSI") {
            # Populate the array with the specific application details for Windows Mobile MSI Apps
            $data += [PSCustomObject]@{ Name = "App Name"; Value = $_.displayName }
            $data += [PSCustomObject]@{ Name = "File Name"; Value = $_.fileName }
            $data += [PSCustomObject]@{ Name = "Command Line"; Value = $_.commandLine }
        }
        # Check if '@odata.type' is '#microsoft.graph.androidManagedStoreApp'
        elseif ($_.'@odata.type' -eq "#microsoft.graph.androidManagedStoreApp") {
            # Populate the array with the specific application details for Android Managed Store Apps
            $data += [PSCustomObject]@{ Name = "App Name"; Value = $_.displayName }
            $data += [PSCustomObject]@{ Name = "App Store URL"; Value = $_.appStoreUrl }
        }
    
        # Display the results in Out-GridView, using the display name as the title
        $data | Out-GridView -Title $name

        # Save report to csv file
        $csvFilePath = Join-Path -Path $ApplicationsFolder -ChildPath "$($name).csv"
        $data | Export-Csv -Path $csvFilePath 
    }
}
    
function Get-Scripts {
    # Log
    write-host "Fetching Scripts"

    # Create Scripts folder within the parent folder
    $ScriptsFolder = Join-Path -Path $ParentFolder -ChildPath "Scripts"
    New-Folder $ScriptsFolder

    # Retrieve device management scripts (Limit 100)
    $scripts = Invoke-MgGraphRequest GET -Uri 'https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts?$top=100'
    
    # Loop through each script and create an individual report
    foreach ($script in $scripts.value) {
        $Id = $script.Id
        $name = $script.displayName

        #Log
        write-host "Fetching Script $name ($Id)"

        # Prepare data as rows (key-value pairs) for Out-GridView
        $data = @(
            [PSCustomObject]@{ Property = 'FileName'; Value = $script.fileName }
            [PSCustomObject]@{ Property = 'RunAsAccount'; Value = $script.runAsAccount }
            [PSCustomObject]@{ Property = 'EnforceSignatureCheck'; Value = $script.enforceSignatureCheck }
            [PSCustomObject]@{ Property = 'RunAs32Bit'; Value = $script.runAs32Bit }
        )

        # Display each script in a separate Out-GridView window, titled with the script's display name
        $data | Out-GridView -Title $name

        # Save report to csv file
        $csvFilePath = Join-Path -Path $ScriptsFolder -ChildPath "$($name).csv"
        $data | Export-Csv -Path $csvFilePath
    }
}

function Get-Intents {
    # Log
    write-host "Fetching Intents"

    # Create Intents folder within the parent folder
    $IntentsFolder = Join-Path -Path $ParentFolder -ChildPath "Intents"
    New-Folder $IntentsFolder
    
    $intents = Invoke-MgGraphRequest GET -Uri "https://graph.microsoft.com/beta/deviceManagement/intents"

    $intents.value | ForEach-Object {
        $Id = $_.Id
        $name = $_.displayName
        
        # Log
        write-host "Fetching Intent "$name" ($Id)"
        
        $data = @()
        $settings = Invoke-MgGraphRequest GET -Uri "https://graph.microsoft.com/beta/deviceManagement/intents('$Id')/settings"
        $settings.value | ForEach-Object {
            $definitionId = $_.definitionId
            $settingName = ((Invoke-MgGraphRequest GET -Uri "https://graph.microsoft.com/beta/deviceManagement/settingDefinitions").value | where-object {$_.id -eq $definitionId }).displayName
            try {
                $value = $_.value
            }
            catch {
            }
            try {
                $value = $_.valueJson
            }
            catch {
            }
            $setting = [PSCustomObject]@{
                name  = $settingName
                value = $value
            }
            $data += $setting
        }
        
        $data | Out-GridView -Title $name

        # Save report to csv file
        $csvFilePath = Join-Path -Path $IntentsFolder -ChildPath "$($name).csv"
        $data | Export-Csv -Path $csvFilePath
    }
}


# Create Parent Folder
New-Folder -folderPath $parentFolder

# Get Setting Catalogs
Get-SettingCatalogs

# Get Configuration Policies
Get-ConfigurationPolicies

# Get Administrative Templates
Get-AdmxPolicies

# Get Compliance Policies
Get-CompliancePolicies

# Get Applications
Get-Applications

# Get Scripts
Get-Scripts

# Get Security Baselines
Get-Intents
