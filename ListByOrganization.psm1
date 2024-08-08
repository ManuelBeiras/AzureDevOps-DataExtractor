Install-Module ImportExcel -Scope CurrentUser

function Invoke-FilterOrganizations {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )
    
    # Organizations
    $organizations = @("Organization1", "Organization2", "Organization3")  # Array with the names of the organizations you want to work with.

    if ($Filter -eq "All") {
        return $organizations
    }
    else {
        $filterArray = $Filter -split ', '
        $filteredOrganizations = @()

        foreach ($org in $filterArray) {
            if ($org -notin $organizations) {
                throw "The organization '$org' does not exist."
            }
            $filteredOrganizations += $org
        }

        return $filteredOrganizations
    }
}

function Invoke-ReplaceCharacters {
    param (
        [string]$jsonString
    )

    # Replace incorrect characters in the JSON string
    $jsonString = $jsonString -replace 'Ý', 'í'  # Replace 'Ý' with 'í'
    $jsonString = $jsonString -replace '¾', 'ó'  # Replace '¾' with 'ó'
    $jsonString = $jsonString -replace 'Ú', 'é'  # Replace 'Ú' with 'é'
    $jsonString = $jsonString -replace '±', 'ñ'  # Replace '±' with 'ñ'
    $jsonString = $jsonString -replace 'ß', 'á'  # Replace 'ß' with 'á'
    $jsonString = $jsonString -replace '┴', 'Á'  # Replace '┴' with 'Á'
    $jsonString = $jsonString -replace '·', 'ú'  # Replace '·' with 'ú'

    return $jsonString
}

function Get-Users {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )

    $filteredOrganizations = Invoke-FilterOrganizations -Filter $Filter

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    # Initialize an array to store all Users information
    $allUsers = @()

    # Process each filtered organization
    foreach ($org in $filteredOrganizations) {

        Export-Excel -Path "Users.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow

        # Initialize an array to store organization's items information
        $orgItems = @()

        # Execute the Azure CLI command to get the JSON response as a string
        $jsonString = az devops user list --organization "https://dev.azure.com/$org"

        # Convert the JSON string to objects with correct encoding
        $jsonResponse = $jsonString | ConvertFrom-Json -Depth 100

        # Convert the response objects back to JSON with corrected encoding
        $correctedJsonString = $jsonResponse | ConvertTo-Json -Depth 100

        # Replace incorrect characters in the JSON string
        $correctedJsonString = Invoke-ReplaceCharacters -jsonString $correctedJsonString
        
        # Convert the corrected JSON string back to objects
        $jsonResponse = $correctedJsonString | ConvertFrom-Json

        # Extract the 'items' array from the JSON response
        $items = $jsonResponse.items

        # Process each item in the 'items' array
        $output = foreach ($item in $items) {
            [PSCustomObject]@{
                "Name"               = $item.user.displayName
                "Email"              = $item.user.mailAddress
                "Organization"       = $org
                "AccountLicenseType" = $item.accessLevel.accountLicenseType
                "LicenseDisplayName" = $item.accessLevel.licenseDisplayName
                "LicensingSource"    = $item.accessLevel.licensingSource
                "MsdnLicenseType"    = $item.accessLevel.msdnLicenseType
                "DateCreated"        = $item.dateCreated
                "LastAccessedDate"   = $item.lastAccessedDate
                "ExecutionDate"      = $ExecutionDate
            }
        }

        # Export the processed data to a CSV file with multiple sheets
        $output | Export-Excel -Path "Users.xlsx" -WorksheetName $org -AutoSize -AutoFilter -FreezeTopRow

        # Append users to organization's items
        $orgItems += $output

        # Append organization's items to all items
        $allUsers += $orgItems
    }
    # Export all Users to "All Information" worksheet
    $allUsers | Export-Excel -Path "Users.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
}

function Get-Groups {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )
  
    $filteredOrganizations = Invoke-FilterOrganizations -Filter $Filter

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    # Initialize an array to store all Groups information
    $allGroups = @()
  
    # Process each filtered organization
    foreach ($org in $filteredOrganizations) {
        Export-Excel -Path "Groups.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
        
        # Initialize an array to store organization's items information
        $orgItems = @()

        # Execute the Azure CLI command to get the JSON response as a string
        $jsonString = az devops security group list --scope organization --organization "https://dev.azure.com/$org"

        # Convert the JSON string to objects with correct encoding
        $jsonResponse = $jsonString | ConvertFrom-Json -Depth 100

        # Convert the response objects back to JSON with corrected encoding
        $correctedJsonString = $jsonResponse | ConvertTo-Json -Depth 100

        # Replace incorrect characters in the JSON string
        $correctedJsonString = Invoke-ReplaceCharacters -jsonString $correctedJsonString
        
        # Convert the corrected JSON string back to objects
        $jsonResponse = $correctedJsonString | ConvertFrom-Json

        # Extract the 'graphGroups' array from the corrected JSON response
        $graphGroups = $jsonResponse.graphGroups

        # Process each graphGroup in the 'graphGroups' array
        $output = foreach ($graphGroup in $graphGroups) {
            [PSCustomObject]@{
                "Name"           = $graphGroup.displayName
                "Organization"   = $org
                "PrincipalName"  = $graphGroup.principalName
                "IsCrossProject" = $graphGroup.isCrossProject
                "Email"          = $graphGroup.mailAddress
                "Description"    = $graphGroup.description
                "ExecutionDate"  = $ExecutionDate
            }
        }

        # Export the processed data to a CSV file with multiple sheets
        $output | Export-Excel -Path "Groups.xlsx" -WorksheetName $org -AutoSize -AutoFilter -FreezeTopRow

        # Append Groups to organization's items
        $orgItems += $output

        # Append organization's items to all Groups
        $allGroups += $orgItems
    }
    # Export all Groups to "All Information" worksheet
    $allGroups  | Export-Excel -Path "Groups.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
}

function Get-Projects {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )
  
    $filteredOrganizations = Invoke-FilterOrganizations -Filter $Filter

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"
  
    # Initialize an array to store all Projects information
    $allProjects = @()

    # Process each filtered organization
    foreach ($org in $filteredOrganizations) {
        Export-Excel -Path "Projects.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow

        # Initialize an array to store organization's items information
        $orgItems = @()

        # Execute the Azure CLI command to get the JSON response as a string
        $jsonString = az devops project list --organization "https://dev.azure.com/$org"
    
        # Convert the JSON string to objects with correct encoding
        $jsonResponse = $jsonString | ConvertFrom-Json -Depth 100

        # Convert the response objects back to JSON with corrected encoding
        $correctedJsonString = $jsonResponse | ConvertTo-Json -Depth 100

        # Replace incorrect characters in the JSON string
        $correctedJsonString = Invoke-ReplaceCharacters -jsonString $correctedJsonString
        
        # Convert the corrected JSON string back to objects
        $jsonResponse = $correctedJsonString | ConvertFrom-Json
    
        # Extract the 'value' array from the corrected JSON response
        $values = $jsonResponse.value
    
        # Process each value in the 'values' array
        $output = foreach ($value in $values) {
            [PSCustomObject]@{
                "Name"           = $value.name
                "Organization"   = $org
                "LastUpdateTime" = $value.lastUpdateTime
                "Visibility"     = $value.visibility
                "Description"    = $value.description
                "ExecutionDate"  = $ExecutionDate
            }
        }
    
        # Export the processed data to a CSV file with multiple sheets
        $output | Export-Excel -Path "Projects.xlsx" -WorksheetName $org -AutoSize -AutoFilter -FreezeTopRow

        # Append Projects to organization's items
        $orgItems += $output

        # Append organization's items to all values
        $allProjects += $orgItems
    }
    # Export all Projects to "All Information" worksheet
    $allProjects | Export-Excel -Path "Projects.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
}

function Get-AgentPools {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )
  
    $filteredOrganizations = Invoke-FilterOrganizations -Filter $Filter

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"
  
    # Initialize an array to store all AgentPools information
    $allAgentPools = @()

    # Process each filtered organization
    foreach ($org in $filteredOrganizations) {
        Export-Excel -Path "AgentsPools.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow

        # Initialize an array to store organization's items information
        $orgItems = @()

        # Execute the Azure CLI command to get the JSON response as a string
        $jsonString = az pipelines pool list --organization "https://dev.azure.com/$org"
    
        # Convert the JSON string to objects with correct encoding
        $jsonResponse = $jsonString | ConvertFrom-Json -Depth 100

        # Convert the response objects back to JSON with corrected encoding
        $correctedJsonString = $jsonResponse | ConvertTo-Json -Depth 100

        # Replace incorrect characters in the JSON string
        $correctedJsonString = Invoke-ReplaceCharacters -jsonString $correctedJsonString
        
        # Convert the corrected JSON string back to objects
        $jsonResponse = $correctedJsonString | ConvertFrom-Json
    
        # Extract the 'value' array from the corrected JSON response
        $values = $jsonResponse
    
        # Process each value in the 'values' array
        $output = foreach ($value in $values) {
            [PSCustomObject]@{
                "AgentPoolName"  = $value.name
                "Organization"   = $org
                "CreatedOn"      = $value.createdOn
                "IsLegacy"       = $value.isLegacy
                "IsHosted"       = $value.isHosted
                "Options"        = $value.options
                "PoolType"       = $value.poolType
                "Size"           = $value.size
                "TargetSize"     = $value.targetSize
                "CreatorName"    = $value.createdBy.displayName
                "CreatorEmail"   = $value.createdBy.uniqueName
                "OwnerName"      = $value.owner.displayName
                "OwnerEmail"     = $value.owner.uniqueName
                "ExecutionDate"  = $ExecutionDate
            }
        }
    
        # Export the processed data to a CSV file with multiple sheets
        $output | Export-Excel -Path "AgentsPools.xlsx" -WorksheetName $org -AutoSize -AutoFilter -FreezeTopRow

        # Append AgentPools to organization's items
        $orgItems += $output

        # Append organization's items to all values
        $allAgentPools += $orgItems
    }
    # Export all AgentPools to "All Information" worksheet
    $allAgentPools | Export-Excel -Path "AgentsPools.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
}

function Get-EverythingByOrganization {
    Get-Users
    Get-Groups
    Get-Projects
    Get-AgentPools
}

Export-ModuleMember -Function Get-EverythingByOrganization, Get-Users, Get-Groups, Get-Projects, Get-AgentPools
