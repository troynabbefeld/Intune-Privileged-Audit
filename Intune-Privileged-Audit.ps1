# trnabbefeld@hbs.net - 2024-09-30, 2024-09-30
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

Connect-MgGraph -Scope "DeviceManagementConfiguration.Read.All", "DeviceManagementApps.Read.All" -NoWelcome

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

$allSettings = Invoke-MgGraphRequest GET -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationSettings"


# Function to check for choiceSettingValue, simpleSettingCollectionValue, and simpleSettingValue, including children
function Get-SettingValue {
    param (
        [PSCustomObject]$settingInstance
    )
    
    # Initialize an array to hold the setting data
    $settingData = @()

    $specificSetting = $allSettings.value | Where-Object {$_.id -eq $settingInstance.settingDefinitionId}
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
            $value = $specificSetting.options | Where-Object {$_.itemId -eq $choiceValue}
            $settingData += [PSCustomObject]@{
                Name  = $settingDisplayName
                Value = $value.displayName
                Type = "Choice"
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
                Name  = $settingDisplayName
                Value = $collectionValue
                Type = "Collection"
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
                Name  = $settingDisplayName
                Value = $simpleValue
                Type = "Simple"
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
    $settingCatalogs = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationpolicies"

    # Initialize row counter
    $rowIndex = 1

    # Setting Catalogs
    $settingCatalogs.Value | ForEach-Object {
        $Id = $_.Id
        Write-Host "Policy Name:" $_.Name
        Write-Host "Created Date:" $_.CreatedDateTime
        Write-Host "Setting Count:" $_.settingCount
        Write-Host "Description:" $_.description
        Write-Host "Platforms:" $_.platforms
        Write-Host "Policy ID:" $_.Id

        $settings = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$Id')/settings"

        # Collect settings for the current policy with row numbers
        $policySettingsData = @()
        $settings.value | ForEach-Object {
            $settingInstance = $_.settingInstance
            Write-Host "Start Settings"
            Write-Host "Setting Definition ID:" $settingInstance.settingDefinitionId

            # Get the settings data and add to the policy's settings collection with a row number as the first column
            $settingValues = Get-SettingValue -settingInstance $settingInstance
            foreach ($value in $settingValues) {
                # Add Row property as the first column
                $indexedValue = [PSCustomObject]@{
                    Row   = $rowIndex
                    Name  = $value.Name
                    Value = $value.Value
                    Type = $value.Type
                }
                $policySettingsData += $indexedValue
                $rowIndex++
            }

            Write-Host "End Settings"
        }

        # Display the policy settings in Out-GridView
        if ($policySettingsData.Count -gt 0) {
            $policySettingsData | Out-GridView -Title $_.Name
        }
    
        Write-Host "---------------------------------------------------------------------------"
        # Reset row index for each new policy if desired
        $rowIndex = 1
    }
}

Get-SettingCatalogs

#Configuration Policies
$configurationPolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations" -Method Get

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

# Function to expand nested properties
function Expand-NestedProperties {
    param (
        [hashtable]$settings
    )

    $expandedSettings = @()

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
                            $expandedSettings += [PSCustomObject]@{
                                SettingName = "$key.$nestedKey"  # Use a dot notation to indicate the hierarchy
                                Value       = $item[$nestedKey]
                            }
                        }
                    }
                    else {
                        # For non-hashtable items, add a row directly
                        $expandedSettings += [PSCustomObject]@{
                            SettingName = $key
                            Value       = $item
                        }
                    }
                }
            }
            else {
                # Regular property, add to expanded settings
                $expandedSettings += [PSCustomObject]@{
                    SettingName = $key
                    Value       = $value
                }
            }
        }
    }
    return $expandedSettings
}

# Loop through each configuration policy
$configurationPolicies.value | ForEach-Object {
    # Get the base object which contains the settings
    $settings = $_

    # Expand settings
    $expandedSettings = Expand-NestedProperties -settings $settings

    # Sort settings alphabetically by SettingName
    $sortedSettings = $expandedSettings | Sort-Object SettingName

    # Display the current policy settings in Out-GridView
    $sortedSettings | Out-GridView -Title $settings.displayName
}



$admxPolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations"
$admxPolicies.Value | ForEach-Object {
    $policyResults = @()
    $Id = $_.Id
    $displayName = $_.displayName
    $settings = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations('$id')/definitionValues"
    $settings.value | ForEach-Object {
        $settingId = $_.ID
        $setting = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations('$id')/definitionValues('$settingId')"
        $settingInfo = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations('$id')/definitionValues('$settingId')/definition"
        $settingValue = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations('$id')/definitionValues('$settingId')/PresentationValues"
        if ($settingValue.value -ne $null) {
            $value = $settingValue.value.value
        }
        else {
            $value = $null
        }
        $policyResult = [PSCustomObject]@{
            Name    = $settingInfo.displayName
            Enabled = $setting.enabled
            value   = $value
        }
        $policyResults += $policyResult
    }
    $policyResults | Out-GridView -Title $displayName
}

# Retrieve all device compliance policies
$compliancePolicies = Get-MgDeviceManagementDeviceCompliancePolicy

# Loop through each compliance policy to get its settings using ForEach-Object
$compliancePolicies | ForEach-Object {
    # Initialize an array to hold settings for the current policy
    $currentPolicySettings = @()

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
            $currentPolicySettings += $setting
        }
    }

    # Sort settings alphabetically by SettingName
    $currentPolicySettings = $currentPolicySettings | Sort-Object SettingName

    # Display the current policy settings in Out-GridView
    $currentPolicySettings | Out-GridView -Title $_.DisplayName
}

$applications = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$filter=(microsoft.graph.managedApp/appAvailability%20eq%20null%20or%20microsoft.graph.managedApp/appAvailability%20eq%20%27lineOfBusiness%27%20or%20isAssigned%20eq%20true)&$orderby=displayName&'

$applications.value | ForEach-Object {
    # Initialize an empty array to hold the application details in row format
    $appDetails = @()

    # Check if '@odata.type' is '#microsoft.graph.win32LobApp'
    if ($_.'@odata.type' -eq "#microsoft.graph.win32LobApp") {
        # Populate the array with the specific application details for Win32 LOB Apps
        $appDetails += [PSCustomObject]@{ Name = "App Name"; Value = $_.displayName }
        $appDetails += [PSCustomObject]@{ Name = "File Name"; Value = $_.fileName }
        $appDetails += [PSCustomObject]@{ Name = "Setup File Path"; Value = $_.setupFilePath }
        $appDetails += [PSCustomObject]@{ Name = "Install Command Line"; Value = $_.installCommandLine }
        $appDetails += [PSCustomObject]@{ Name = "Uninstall Command Line"; Value = $_.uninstallCommandLine }
        $appDetails += [PSCustomObject]@{ Name = "Minimum Supported Windows Release"; Value = $_.minimumSupportedWindowsRelease }
        $appDetails += [PSCustomObject]@{ Name = "Applicable Architectures"; Value = $_.applicableArchitectures }
        
        # Add detection rules as separate rows
        $_.detectionRules | ForEach-Object {
            if ($_.'@odata.type' -eq "#microsoft.graph.win32LobAppFileSystemDetection") {
                $appDetails += [PSCustomObject]@{ Name = "Detection Rule Type"; Value = "File" }
                $appDetails += [PSCustomObject]@{ Name = "Path"; Value = $_.path }
                $appDetails += [PSCustomObject]@{ Name = "File/Folder Name"; Value = $_.fileOrFolderName }
                $appDetails += [PSCustomObject]@{ Name = "Detection Method"; Value = $_.detectionType }
            }
            elseif ($_.'@odata.type' -eq "#microsoft.graph.win32LobAppRegistryDetection") {
                $appDetails += [PSCustomObject]@{ Name = "Detection Rule Type"; Value = "Registry" }
                $appDetails += [PSCustomObject]@{ Name = "Path"; Value = "$($_.keyPath)\$($_.valueName)" }
                $appDetails += [PSCustomObject]@{ Name = "Detection Method"; Value = $_.detectionType }
                if ($_.detectionValue -ne "") {
                    $appDetails += [PSCustomObject]@{ Name = "Operator "; Value = $_.operator }
                    $appDetails += [PSCustomObject]@{ Name = "Value "; Value = $_.detectionValue }
                }
            }
            elseif ($_.'@odata.type' -eq "#microsoft.graph.win32LobAppProductCodeDetection") {
                $appDetails += [PSCustomObject]@{ Name = "Detection Rule Type"; Value = "MSI" }
                $appDetails += [PSCustomObject]@{ Name = "Product Code"; Value = $_.productCode }
            }
        }
    }
    # Check if '@odata.type' is '#microsoft.graph.windowsMobileMSI'
    elseif ($_.'@odata.type' -eq "#microsoft.graph.windowsMobileMSI") {
        # Populate the array with the specific application details for Windows Mobile MSI Apps
        $appDetails += [PSCustomObject]@{ Name = "App Name"; Value = $_.displayName }
        $appDetails += [PSCustomObject]@{ Name = "File Name"; Value = $_.fileName }
        $appDetails += [PSCustomObject]@{ Name = "Command Line"; Value = $_.commandLine }
    }
    # Check if '@odata.type' is '#microsoft.graph.androidManagedStoreApp'
    elseif ($_.'@odata.type' -eq "#microsoft.graph.androidManagedStoreApp") {
        # Populate the array with the specific application details for Android Managed Store Apps
        $appDetails += [PSCustomObject]@{ Name = "App Name"; Value = $_.displayName }
        $appDetails += [PSCustomObject]@{ Name = "App Store URL"; Value = $_.appStoreUrl }
    }

    # Display the results in Out-GridView, using the display name as the title
    if ($appDetails.Count -gt 0) {
        $appDetails | Out-GridView -Title $_.displayName
    }
}


# Retrieve device management scripts
$scripts = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts'

# Loop through each script and create an individual report
foreach ($script in $scripts.value) {
    # Prepare data as rows (key-value pairs) for Out-GridView
    $reportData = @(
        [PSCustomObject]@{ Property = 'FileName'; Value = $script.fileName }
        [PSCustomObject]@{ Property = 'RunAsAccount'; Value = $script.runAsAccount }
        [PSCustomObject]@{ Property = 'EnforceSignatureCheck'; Value = $script.enforceSignatureCheck }
        [PSCustomObject]@{ Property = 'RunAs32Bit'; Value = $script.runAs32Bit }
    )

    # Display each script in a separate Out-GridView window, titled with the script's display name
    $reportData | Out-GridView -Title $script.displayName
}

$securityBaselinePolicies = Invoke-MgGraphRequest GET -Uri "https://graph.microsoft.com/beta/deviceManagement/intents"
$securityBaselinePolicies.value | ForEach-Object {
    $report = @()
    $Id = $_.Id
    $settings = Invoke-MgGraphRequest GET -Uri "https://graph.microsoft.com/beta/deviceManagement/intents('$Id')/settings"
    $settings.value | ForEach-Object {
        $definitionId = $_.definitionId
        $name = $allSettings2.value | Where-Object {$_.id -eq $definitionId}
        $name
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
            name  = $name.displayName
            value = $value
        }
        $report += $setting
    }
    $report | Out-GridView -Title $_.displayName
}

$allSettings = Invoke-MgGraphRequest GET -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationSettings"
$specificSetting = $allSettings.value | Where-Object {$_.id -eq 'device_vendor_msft_policy_config_remoteassistance_unsolicitedremoteassistance_ra_unsolicit_dacl_edit'}
$specificSetting.options | Where-Object {$_.itemId -eq 'device_vendor_msft_policy_config_remoteassistance_unsolicitedremoteassistance_ra_unsolicit_control_list_1'} | Select-Object displayName
$specificSetting.displayName

$allSettings2 = Invoke-MgGraphRequest GET -Uri "https://graph.microsoft.com/beta/deviceManagement/settingDefinitions"
$allSettings2.value | Where-Object {$_.id -eq 'deviceConfiguration--windows10GeneralConfiguration_startMenuPinnedFolderSettings'}