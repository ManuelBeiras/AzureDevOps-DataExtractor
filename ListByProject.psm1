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

function Invoke-RetrieveDevOpsProjects {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )
    
    $filteredOrganizations = Invoke-FilterOrganizations -Filter $Filter

    # Initialize hashtable to store projects linked to each organization
    $projectsHashtable = @{}

    # Process each filtered organization
    foreach ($org in $filteredOrganizations) {
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
    
        # Initialize arrays to store project names and IDs for the current organization
        $projects = @()
        $ids = @()
    
        # Process each value in the 'values' array
        foreach ($value in $values) {
            $projects += $value.name
            $ids += $value.id
        }

        # Add projects and ids to the hashtable
        $projectsHashtable[$org] = @{
            Projects = $projects
            Ids = $ids
        }
    }

    # Return the hashtable
    return $projectsHashtable
}

function Invoke-RetrieveDevOpsTeams {
    Param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        [Parameter(Mandatory = $true)]
        [string]$ProjectId
    )

    # Execute the Azure CLI command to get the JSON response as a string
    $jsonString = az devops team list --org "https://dev.azure.com/$Organization" --project "$ProjectId"
    
    # Convert the JSON string to objects with correct encoding
    $jsonResponse = $jsonString | ConvertFrom-Json -Depth 100

    # Convert the response objects back to JSON with corrected encoding
    $correctedJsonString = $jsonResponse | ConvertTo-Json -Depth 100

    # Replace incorrect characters in the JSON string
    $correctedJsonString = Invoke-ReplaceCharacters -jsonString $correctedJsonString
    
    # Convert the corrected JSON string back to objects
    $jsonResponse = $correctedJsonString | ConvertFrom-Json

    # Process each team
    $teamsArray = foreach ($team in $jsonResponse) {
        # Create a custom object for the team
        [PSCustomObject]@{
            "TeamName"    = $team.name
            "ProjectName" = $team.projectName
            "TeamID"      = $team.id
            "Description"   = $team.description
        }
    }

    # Return the array of team objects
    return $teamsArray
}

function Invoke-RetrieveDevOpsRepositories {
    Param (
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        [Parameter(Mandatory = $true)]
        [string]$ProjectId
    )

    # Execute the Azure CLI command to get the JSON response as a string
    $jsonString = az repos list --project "$ProjectId" --organization "https://dev.azure.com/$Organization"
    
    # Convert the JSON string to objects with correct encoding
    $jsonResponse = $jsonString | ConvertFrom-Json -Depth 100

    # Convert the response objects back to JSON with corrected encoding
    $correctedJsonString = $jsonResponse | ConvertTo-Json -Depth 100

    # Replace incorrect characters in the JSON string
    $correctedJsonString = Invoke-ReplaceCharacters -jsonString $correctedJsonString
    
    # Convert the corrected JSON string back to objects
    $jsonResponse = $correctedJsonString | ConvertFrom-Json

    # Process each Repository
    $repositoriesArray = foreach ($Repository in $jsonResponse) {
        # Create a custom object for the Repository
        [PSCustomObject]@{
            "ProjectName"    = $Repository.project.name
            "RepoName"       = $Repository.name
            "RepoID"         = $Repository.id
            "LastUpdateTime" = $Repository.project.lastUpdateTime
            "Visibility"     = $Repository.project.visibility
            "Size"           = $Repository.size
            "Description"    = $Repository.project.description
            "WebUrl"         = $Repository.webUrl
            "RemoteUrl"      = $Repository.remoteUrl
            "SSHUrl"         = $Repository.sshUrl
        }
    }

    # Return the array of Repository objects
    return $repositoriesArray
}

Function Invoke-AzureDevOpsAuth {
    $token = az account get-access-token --query accessToken -o tsv
    $AzureDevOpsAuthenicationHeader = @{Authorization=("Bearer {0}" -f $token)}
  
    $AzureDevOpsAuthenicationHeader
}

function Get-ServiceConnections {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )

    # Retrieve projects linked to each organization along with IDs
    $projectsHashtable = Invoke-RetrieveDevOpsProjects -Filter $Filter

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    # Process each organization and its associated projects
    foreach ($org in $projectsHashtable.Keys) {
        # Initialize an array to store all Service Connections information
        $allserviceConnection = @()

        Export-Excel -Path "ServiceConnections/ServiceConnections_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
        Write-Host "Organization: $org"
        Write-Host "Projects:"

        $projects = $projectsHashtable[$org].Projects
        $ids = $projectsHashtable[$org].Ids

        for ($i = 0; $i -lt $projects.Count; $i++) {
            $project = $projects[$i]
            $id = $ids[$i]
            Write-Host "  $project with id: $id"
            
            # Execute the Azure CLI command to get the JSON response as a string
            $jsonString = az devops service-endpoint list --project "$id" --organization "https://dev.azure.com/$org"

            # Convert the JSON string to objects with correct encoding
            $jsonResponse = $jsonString | ConvertFrom-Json -Depth 100

            # Convert the response objects back to JSON with corrected encoding
            $correctedJsonString = $jsonResponse | ConvertTo-Json -Depth 100

            # Replace incorrect characters in the JSON string
            $correctedJsonString = Invoke-ReplaceCharacters -jsonString $correctedJsonString
            
            # Convert the corrected JSON string back to objects
            $jsonResponse = $correctedJsonString | ConvertFrom-Json

            # Extract specific fields from the JSON response
            $output = foreach ($serviceConnection in $jsonResponse) {
                [PSCustomObject]@{
                    "ServiceConnectionName" = $serviceConnection.serviceEndpointProjectReferences.name
                    "ProjectName"           = $serviceConnection.serviceEndpointProjectReferences.projectReference.name
                    "ProjectId"             = $serviceConnection.serviceEndpointProjectReferences.projectReference.id
                    "ApplicationId"         = $serviceConnection.authorization.parameters.applicationId
                    "UserName"              = $serviceConnection.createdBy.displayName
                    "UserId"                = $serviceConnection.createdBy.id
                    "Email"                 = $serviceConnection.createdBy.uniqueName
                    "Type"                  = $serviceConnection.type
                    "IsOutdated"            = $serviceConnection.isOutdated
                    "IsReady"               = $serviceConnection.isReady
                    "IsShared"              = $serviceConnection.isShared
                    "ExecutionDate"         = $ExecutionDate
                }
            }
            
            # Export Service Connections to Excel
            $output | Export-Excel -Path "ServiceConnections/ServiceConnections_$org.xlsx" -WorksheetName "$project" -AutoSize -AutoFilter -FreezeTopRow

            # Append project Service Connections to all Service Connections
            $allserviceConnection += $output
        }
        # Export all Service Connections to "All Information" worksheet
        $allserviceConnection | Export-Excel -Path "ServiceConnections/ServiceConnections_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
    }
}

function Get-RepositoriesGIT {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )

    # Retrieve projects linked to each organization along with IDs
    $projectsHashtable = Invoke-RetrieveDevOpsProjects -Filter $Filter

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    # Process each organization and its associated projects
    foreach ($org in $projectsHashtable.Keys) {
        # Initialize an array to store all Repositories information
        $allRepository = @()

        Export-Excel -Path "RepositoriesGIT/RepositoriesGIT_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
        Write-Host "Organization: $org"
        Write-Host "Projects:"

        $projects = $projectsHashtable[$org].Projects
        $ids = $projectsHashtable[$org].Ids

        for ($i = 0; $i -lt $projects.Count; $i++) {
            $project = $projects[$i]
            $id = $ids[$i]
            Write-Host "  $project with id: $id"
            
            # Execute the Azure CLI command to get the JSON response as a string
            $jsonString = az repos list --project "$id" --organization "https://dev.azure.com/$org"

            # Convert the JSON string to objects with correct encoding
            $jsonResponse = $jsonString | ConvertFrom-Json -Depth 100

            # Convert the response objects back to JSON with corrected encoding
            $correctedJsonString = $jsonResponse | ConvertTo-Json -Depth 100

            # Replace incorrect characters in the JSON string
            $correctedJsonString = Invoke-ReplaceCharacters -jsonString $correctedJsonString
            
            # Convert the corrected JSON string back to objects
            $jsonResponse = $correctedJsonString | ConvertFrom-Json

            # Extract specific fields from the JSON response
            $output = foreach ($Repository in $jsonResponse) {
                [PSCustomObject]@{
                    "ProjectName"    = $Repository.project.name
                    "RepoName"       = $Repository.name
                    "LastUpdateTime" = $Repository.project.lastUpdateTime
                    "Visibility"     = $Repository.project.visibility
                    "Size"           = $Repository.size
                    "Description"    = $Repository.project.description
                    "WebUrl"         = $Repository.webUrl
                    "RemoteUrl"      = $Repository.remoteUrl
                    "SSHUrl"         = $Repository.sshUrl
                    "RepositoryType" = "GIT"
                    "ExecutionDate"  = $ExecutionDate
                }
            }
            
            # Export Repositories to Excel
            $output | Export-Excel -Path "RepositoriesGIT/RepositoriesGIT_$org.xlsx" -WorksheetName "$project" -AutoSize -AutoFilter -FreezeTopRow

            # Append project Repositories to all Repositories
            $allRepository += $output
        }
        # Export all Repositories to "All Information" worksheet
        $allRepository | Export-Excel -Path "RepositoriesGIT/RepositoriesGIT_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
    }
}

function Get-RepositoriesTFVC {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )
    # Retrieve projects linked to each organization along with IDs
    $projectsHashtable = Invoke-RetrieveDevOpsProjects -Filter $Filter

    Write-Host "Retrieved Organizations: $($projectsHashtable.Count)"

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    # Process each organization and its associated projects
    foreach ($org in $projectsHashtable.Keys) {
        # Initialize an array to store all Repositories Branches information
        $allrepositoriesTFVC = @()

        Export-Excel -Path "RepositoriesTFVC/RepositoriesTFVC_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow

        Write-Host "Processing Organization: $org"
        Write-Host "Projects:"

        $projects = $projectsHashtable[$org].Projects
        $ids = $projectsHashtable[$org].Ids

        # Process each project within the organization
        for ($i = 0; $i -lt $projects.Count; $i++) {
            $project = $projects[$i]
            $id = $ids[$i]
            Write-Host "  Project: $project with id: $id"

            Write-Host "Fetching Branches for Project: $project"

            # Initialize an array to store Repositories Branches information for this project
            $repositoriesTFVC = @() 
            
            if ($null -eq $AzureDevOpsAuthenicationHeader) {
                $AzureDevOpsAuthenicationHeader = Invoke-AzureDevOpsAuth
            }

            $bodyRequest = @{}

            $bodyRequest = $bodyRequest | ConvertTo-Json -Depth 100
            
            $uriGetListProjects = "https://dev.azure.com/$org/$id/_apis/tfvc/items?api-version=7.1-preview.1"
            # $uriGetListProjects = "https://almsearch.dev.azure.com/$org/$id/_apis/search/status/tfvc?api-version=7.1-preview.1"

            # Execute the Azure CLI command to get Repositories Branches
            $jsonString2 = Invoke-RestMethod -Uri $uriGetListProjects -Method Get -Headers $AzureDevOpsAuthenicationHeader -Body $bodyRequest -ContentType "application/json"

            if ($jsonString2.count -eq 0) {
                Write-Host "No Repositories TFVC found"
            }         
            else {
                # # Convert the corrected JSON string back to objects
                $repositoriesTFVCArray = $jsonString2.value

                # Extract specific fields and add to repoBranchesTFVCArray array
                foreach ($repoTFVC in $repositoriesTFVCArray) {
                    if ($repoTFVC.path -match '^\$\/[\w-]+$') {
                        $repositoriesTFVC += [PSCustomObject]@{
                            "ProjectName"     = $project
                            "RepoName"        = $repoTFVC.path
                            "RepositoryType"  = "TFVC"
                            "ExecutionDate"   = $ExecutionDate
                        }
                    }
                }
            }

            # Export Repositories Branches to Excel
            $repositoriesTFVC | Export-Excel -Path "RepositoriesTFVC/RepositoriesTFVC_$org.xlsx" -WorksheetName $project -AutoSize -AutoFilter -FreezeTopRow

            # Append project Repositories Branches to all Repositories Branches
            $allrepositoriesTFVC += $repositoriesTFVC
        }
        # Export all Repositories Branches to "All Information" worksheet
        $allrepositoriesTFVC | Export-Excel -Path "RepositoriesTFVC/RepositoriesTFVC_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
    }
    Write-Host "Exporting TFVC Branches to Excel completed."
}

function Get-AllRepositories {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )

    # Retrieve projects linked to each organization along with IDs
    $projectsHashtable = Invoke-RetrieveDevOpsProjects -Filter $Filter

    Write-Host "Retrieved Organizations: $($projectsHashtable.Count)"

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    # Process each organization and its associated projects
    foreach ($org in $projectsHashtable.Keys) {
        # Initialize arrays to store all repositories information
        $allRepositories = @()

        # Initialize Excel file for all repositories
        Export-Excel -Path "AllRepositories/AllRepositories_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow

        Write-Host "Processing Organization: $org"
        Write-Host "Projects:"

        $projects = $projectsHashtable[$org].Projects
        $ids = $projectsHashtable[$org].Ids

        # Process each project within the organization
        for ($i = 0; $i -lt $projects.Count; $i++) {
            $project = $projects[$i]
            $id = $ids[$i]
            Write-Host "  Project: $project with id: $id"

            # Initialize arrays to store repositories information for this project
            $repositories = @()

            # Fetch Git repositories
            Write-Host "Fetching Git repositories for Project: $project"

            # Execute the Azure CLI command to get the JSON response as a string
            $jsonString = az repos list --project "$id" --organization "https://dev.azure.com/$org"

            # Convert the JSON string to objects with correct encoding
            $jsonResponse = $jsonString | ConvertFrom-Json -Depth 100

            # Convert the response objects back to JSON with corrected encoding
            $correctedJsonString = $jsonResponse | ConvertTo-Json -Depth 100

            # Replace incorrect characters in the JSON string
            $correctedJsonString = Invoke-ReplaceCharacters -jsonString $correctedJsonString
            
            # Convert the corrected JSON string back to objects
            $jsonResponse = $correctedJsonString | ConvertFrom-Json

            # Extract specific fields from the JSON response and add to repositories array
            foreach ($repository in $jsonResponse) {
                $repositories += [PSCustomObject]@{
                    "ProjectName"    = $repository.project.name
                    "RepoName"       = $repository.name
                    "LastUpdateTime" = $repository.project.lastUpdateTime
                    "Visibility"     = $repository.project.visibility
                    "Size"           = $repository.size
                    "RepositoryType" = "GIT"
                    "ExecutionDate"  = $ExecutionDate
                    "Description"    = $repository.project.description
                    "WebUrl"         = $repository.webUrl
                    "RemoteUrl"      = $repository.remoteUrl
                    "SSHUrl"         = $repository.sshUrl
                }
            }

            if ($null -eq $AzureDevOpsAuthenicationHeader) {
                $AzureDevOpsAuthenicationHeader = Invoke-AzureDevOpsAuth
            }

            # Fetch TFVC repositories
            Write-Host "Fetching TFVC repositories for Project: $project"

            $bodyRequest = @{}

            $bodyRequest = $bodyRequest | ConvertTo-Json -Depth 100

            # Execute the Azure DevOps REST API to get repositories
            # https://almsearch.dev.azure.com/$org/$id/_apis/search/status/tfvc On some repositories that were tfvc got parsed like normal repositories. So this don't work.
            $uriGetListProjects = "https://dev.azure.com/$org/$id/_apis/tfvc/items?api-version=7.1-preview.1"
            $jsonString2 = Invoke-RestMethod -Uri $uriGetListProjects -Method Get -Headers $AzureDevOpsAuthenicationHeader -Body $bodyRequest -ContentType "application/json"

            # Extract specific fields and add to repositories array
            if ($jsonString2.count -ne 0) {
                foreach ($repoTFVC in $jsonString2.value) {
                    if ($repoTFVC.path -match '^\$\/[\w-]+$') {
                        $repositories += [PSCustomObject]@{
                            "ProjectName"     = $project
                            "RepoName"        = $repoTFVC.path
                            "RepositoryType"  = "TFVC"
                            "ExecutionDate"   = $ExecutionDate
                        }
                    }
                }
            }

            # Export repositories to Excel
            $repositories | Export-Excel -Path "AllRepositories/AllRepositories_$org.xlsx" -WorksheetName $project -AutoSize -AutoFilter -FreezeTopRow

            # Append project repositories to all repositories
            $allRepositories += $repositories
        }

        # Export all repositories to "All Information" worksheet
        $allRepositories | Export-Excel -Path "AllRepositories/AllRepositories_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
    }
    Write-Host "Exporting repositories to Excel completed."
}

function Get-RepositoriesPRs {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )

    # Retrieve projects linked to each organization along with IDs
    $projectsHashtable = Invoke-RetrieveDevOpsProjects -Filter $Filter

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    # Process each organization and its associated projects
    foreach ($org in $projectsHashtable.Keys) {
        # Initialize an array to store all Repositories Pull Requests information
        $allRepositoryPRs = @()

        Export-Excel -Path "RepositoriesPRs/RepositoriesPRs_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
        Write-Host "Organization: $org"
        Write-Host "Projects:"

        $projects = $projectsHashtable[$org].Projects
        $ids = $projectsHashtable[$org].Ids

        for ($i = 0; $i -lt $projects.Count; $i++) {
            $project = $projects[$i]
            $id = $ids[$i]
            Write-Host "  $project with id: $id"
            
            # Execute the Azure CLI command to get the JSON response as a string
            $jsonString = az repos pr list --status 'all' --project "$id" --organization "https://dev.azure.com/$org"

            # Convert the JSON string to objects with correct encoding
            $jsonResponse = $jsonString | ConvertFrom-Json -Depth 100

            # Convert the response objects back to JSON with corrected encoding
            $correctedJsonString = $jsonResponse | ConvertTo-Json -Depth 100

            # Replace incorrect characters in the JSON string
            $correctedJsonString = Invoke-ReplaceCharacters -jsonString $correctedJsonString
            
            # Convert the corrected JSON string back to objects
            $jsonResponse = $correctedJsonString | ConvertFrom-Json

            # Extract specific fields from the JSON response
            $output = foreach ($RepositoryPRs in $jsonResponse) {
                [PSCustomObject]@{
                    "Title"         = $RepositoryPRs.title
                    "ProjectName"   = $project
                    "Status"        = $RepositoryPRs.status
                    "Branch"        = $RepositoryPRs.sourceRefName
                    "Project"       = $RepositoryPRs.repository.project.name
                    "Repository"    = $RepositoryPRs.repository.name
                    "CreatorName"   = $RepositoryPRs.createdBy.displayName
                    "Email"         = $RepositoryPRs.createdBy.uniqueName
                    "CreationDate"  = $RepositoryPRs.creationDate
                    "Description"   = $RepositoryPRs.description
                    "ExecutionDate" = $ExecutionDate
                }
            }
            
            # Export Repositories Pull Requests to Excel
            $output | Export-Excel -Path "RepositoriesPRs/RepositoriesPRs_$org.xlsx" -WorksheetName "$project" -AutoSize -AutoFilter -FreezeTopRow

            # Append project Repositories Pull Requests to all Repositories Pull Requests
            $allRepositoryPRs += $output
        }
        # Export all Repositories Pull Requests to "All Information" worksheet
        $allRepositoryPRs | Export-Excel -Path "RepositoriesPRs/RepositoriesPRs_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
    }
}

function Get-Teams {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )

    # Retrieve projects linked to each organization along with IDs
    $projectsHashtable = Invoke-RetrieveDevOpsProjects -Filter $Filter

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    # Process each organization and its associated projects
    foreach ($org in $projectsHashtable.Keys) {
        # Initialize an array to store all Teams information
        $allTeams = @()

        Export-Excel -Path "Teams/Teams_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
        Write-Host "Organization: $org"
        Write-Host "Projects:"

        $projects = $projectsHashtable[$org].Projects
        $ids = $projectsHashtable[$org].Ids

        for ($i = 0; $i -lt $projects.Count; $i++) {
            $project = $projects[$i]
            $id = $ids[$i]
            Write-Host "  $project with id: $id"
            
            # Execute the Azure CLI command to get the JSON response as a string
            $jsonString = az devops team list --project "$id" --organization "https://dev.azure.com/$org"

            # Convert the JSON string to objects with correct encoding
            $jsonResponse = $jsonString | ConvertFrom-Json -Depth 100

            # Convert the response objects back to JSON with corrected encoding
            $correctedJsonString = $jsonResponse | ConvertTo-Json -Depth 100

            # Replace incorrect characters in the JSON string
            $correctedJsonString = Invoke-ReplaceCharacters -jsonString $correctedJsonString
            
            # Convert the corrected JSON string back to objects
            $jsonResponse = $correctedJsonString | ConvertFrom-Json

            # Extract specific fields from the JSON response
            $output = foreach ($Teams in $jsonResponse) {
                [PSCustomObject]@{
                    "Name"          = $Teams.name
                    "ProjectName"   = $Teams.projectName
                    "TeamID"        = $Teams.id
                    "Description"   = $Teams.description
                    "ExecutionDate" = $ExecutionDate
                }
            }
            
            # Export Teams to Excel
            $output | Export-Excel -Path "Teams/Teams_$org.xlsx" -WorksheetName "$project" -AutoSize -AutoFilter -FreezeTopRow

            # Append project Teams to all Teams
            $allTeams += $output
        }
        # Export all Teams to "All Information" worksheet
        $allTeams | Export-Excel -Path "Teams/Teams_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
    }
}

function Get-TeamsMembers {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )

    # Retrieve projects linked to each organization along with IDs
    $projectsHashtable = Invoke-RetrieveDevOpsProjects -Filter $Filter

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    Write-Host "Retrieved Organizations: $($projectsHashtable.Count)"

    # Process each organization and its associated projects
    foreach ($org in $projectsHashtable.Keys) {
        # Initialize an array to store all Teams Members information
        $allTeamMembers = @()

        Export-Excel -Path "TeamsMembers/TeamsMembers_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
        Write-Host "Processing Organization: $org"
        Write-Host "Projects:"

        $projects = $projectsHashtable[$org].Projects
        $ids = $projectsHashtable[$org].Ids

        # Process each project within the organization
        for ($i = 0; $i -lt $projects.Count; $i++) {
            $project = $projects[$i]
            $id = $ids[$i]
            Write-Host "  Project: $project with id: $id"

            Write-Host "Fetching teams for Project: $project"

            # Get teams for the current project
            $teams = Invoke-RetrieveDevOpsTeams -Organization $org -ProjectId $id

            # Initialize an array to store Teams Members information for this project
            $teamMembers = @()

            # Iterate over each team to fetch Teams Members
            foreach ($team in $teams) {
                $teamId = $team.TeamID

                # Execute the Azure CLI command to get team details
                $jsonString1 = az devops team show --team $teamId --project $id --organization "https://dev.azure.com/$org"
                
                # Convert the JSON string to objects with correct encoding
                $jsonResponse1 = $jsonString1 | ConvertFrom-Json -Depth 100

                # Convert the response objects back to JSON with corrected encoding
                $correctedJsonString1 = $jsonResponse1 | ConvertTo-Json -Depth 100

                # Replace incorrect characters in the JSON string
                $correctedJsonString1 = Invoke-ReplaceCharacters -jsonString $correctedJsonString1

                # Convert the corrected JSON string back to objects
                $jsonResponse1 = $correctedJsonString1 | ConvertFrom-Json

                # Extract team name from team details
                $teamName = $jsonResponse1.name

                # Execute the Azure CLI command to get team members
                $jsonString2 = az devops team list-member --team $teamId --project $id --organization "https://dev.azure.com/$org"

                # Convert the JSON string to objects with correct encoding
                $jsonResponse2 = $jsonString2 | ConvertFrom-Json -Depth 100

                # Convert the response objects back to JSON with corrected encoding
                $correctedJsonString2 = $jsonResponse2 | ConvertTo-Json -Depth 100

                # Replace incorrect characters in the JSON string
                $correctedJsonString2 = Invoke-ReplaceCharacters -jsonString $correctedJsonString2

                # Convert the corrected JSON string back to objects
                $teamMembersArray = $correctedJsonString2 | ConvertFrom-Json

                # Extract specific fields and add to teamMembers array
                foreach ($member in $teamMembersArray) {
                    $teamMembers += [PSCustomObject]@{
                        "UserName"        = $member.identity.displayName
                        "Email"           = $member.identity.uniqueName
                        "TeamName"        = $teamName
                        "ProjectName"     = $project
                        "IsAdministrator" = $member.isTeamAdmin
                        "ExecutionDate"   = $ExecutionDate
                    }
                }
            }

            # Export Teams Members to Excel
            $teamMembers | Export-Excel -Path "TeamsMembers/TeamsMembers_$org.xlsx" -WorksheetName $project -AutoSize -AutoFilter -FreezeTopRow

            # Append project Teams Members to all Teams Members
            $allTeamMembers += $teamMembers
        }
        # Export all Teams Members to "All Information" worksheet
        $allTeamMembers | Export-Excel -Path "TeamsMembers/TeamsMembers_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
    }

    Write-Host "Exporting team members to Excel completed."
}

function Get-VariablesGroups {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )

    # Retrieve projects linked to each organization along with IDs
    $projectsHashtable = Invoke-RetrieveDevOpsProjects -Filter $Filter

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    # Process each organization and its associated projects
    foreach ($org in $projectsHashtable.Keys) {
        # Initialize an array to store all Variable Groups information
        $allVariableGroups = @()

        Export-Excel -Path "VariableGroups/VariableGroups_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
        Write-Host "Organization: $org"
        Write-Host "Projects:"

        $projects = $projectsHashtable[$org].Projects
        $ids = $projectsHashtable[$org].Ids

        for ($i = 0; $i -lt $projects.Count; $i++) {
            $project = $projects[$i]
            $id = $ids[$i]
            Write-Host "  $project with id: $id"

            # Execute the Azure CLI command to get the JSON response as a string
            $jsonString = az pipelines variable-group list --project "$id" --organization "https://dev.azure.com/$org"

            # Convert the JSON string to objects with correct encoding
            $jsonResponse = $jsonString | ConvertFrom-Json -Depth 100

            # Convert the response objects back to JSON with corrected encoding
            $correctedJsonString = $jsonResponse | ConvertTo-Json -Depth 100

            # Replace incorrect characters in the JSON string
            $correctedJsonString = Invoke-ReplaceCharacters -jsonString $correctedJsonString
            
            # Convert the corrected JSON string back to objects
            $jsonResponse = $correctedJsonString | ConvertFrom-Json

            # Extract specific fields from the JSON response
            $output = foreach ($VariableGroups in $jsonResponse) {
                # Convert variables object to a single string
                $variablesString = ($VariableGroups.variables | ConvertTo-Json -Compress)           
                [PSCustomObject]@{
                    "VariableGroupName" = $VariableGroups.name
                    "ProjectName"       = $project
                    "IsShared"          = $VariableGroups.isShared
                    "CreatedBy"         = $VariableGroups.createdBy.displayName
                    "CreatedByEmail"    = $VariableGroups.createdBy.uniqueName
                    "ModifiedBy"        = $VariableGroups.modifiedBy.displayName
                    "ModifiedByEmail"   = $VariableGroups.modifiedBy.uniqueName
                    "CreatedOn"         = $VariableGroups.createdOn
                    "ModifiedOn"        = $VariableGroups.modifiedOn
                    "Description"       = $VariableGroups.description
                    "Variables"         = $variablesString  # Save variables as plain text
                    "ExecutionDate"     = $ExecutionDate
                }
            }
            
            # Export Variable Groups to Excel
            $output | Export-Excel -Path "VariableGroups/VariableGroups_$org.xlsx" -WorksheetName "$project" -AutoSize -AutoFilter -FreezeTopRow

            # Append project Variable Groups to all Variable Groups
            $allVariableGroups += $output
        }
        # Export all Variable Groups to "All Information" worksheet
        $allVariableGroups | Export-Excel -Path "VariableGroups/VariableGroups_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
    }
}

function Get-Pipelines {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )

    # Retrieve projects linked to each organization along with IDs
    $projectsHashtable = Invoke-RetrieveDevOpsProjects -Filter $Filter

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    # Process each organization and its associated projects
    foreach ($org in $projectsHashtable.Keys) {
        # Initialize an array to store Pipelines information
        $allPipelines = @()

        Export-Excel -Path "Pipelines/Pipelines_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
        Write-Host "Organization: $org"
        Write-Host "Projects:"

        $projects = $projectsHashtable[$org].Projects
        $ids = $projectsHashtable[$org].Ids

        for ($i = 0; $i -lt $projects.Count; $i++) {
            $project = $projects[$i]
            $id = $ids[$i]
            Write-Host "  $project with id: $id"
            
            # Execute the Azure CLI command to get the JSON response as a string
            $jsonString = az pipelines list --project "$id" --organization "https://dev.azure.com/$org"

            # Convert the JSON string to objects with correct encoding
            $jsonResponse = $jsonString | ConvertFrom-Json -Depth 100

            # Convert the response objects back to JSON with corrected encoding
            $correctedJsonString = $jsonResponse | ConvertTo-Json -Depth 100

            # Replace incorrect characters in the JSON string
            $correctedJsonString = Invoke-ReplaceCharacters -jsonString $correctedJsonString
            
            # Convert the corrected JSON string back to objects
            $jsonResponse = $correctedJsonString | ConvertFrom-Json

            # Extract specific fields from the JSON response
            $output = foreach ($Pipelines in $jsonResponse) {
                [PSCustomObject]@{
                    "PipelineName"    = $Pipelines.name
                    "AuthoredByName"  = $Pipelines.authoredBy.displayName
                    "AuthoredByEmail" = $Pipelines.authoredBy.uniqueName
                    "ProjectName"     = $Pipelines.project.name
                    "AgentPoolName"   = $Pipelines.queue.pool.name
                    "QueueStatus"     = $Pipelines.queueStatus
                    "CreatedDate"     = $Pipelines.createdDate
                    "ExecutionDate"   = $ExecutionDate
                }
            }
            
            # Export Pipelines to Excel
            $output | Export-Excel -Path "Pipelines/Pipelines_$org.xlsx" -WorksheetName "$project" -AutoSize -AutoFilter -FreezeTopRow

            # Append project Pipelines to all Pipelines
            $allPipelines += $output
        }
        # Export all Pipelines to "All Information" worksheet
        $allPipelines | Export-Excel -Path "Pipelines/Pipelines_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
    }
}

function Get-PipelinesReleases {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )

    # Retrieve projects linked to each organization along with IDs
    $projectsHashtable = Invoke-RetrieveDevOpsProjects -Filter $Filter

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    # Process each organization and its associated projects
    foreach ($org in $projectsHashtable.Keys) {
        # Initialize an array to store Pipelines Releases information
        $allPipelinesReleases = @()

        Export-Excel -Path "PipelinesReleases/PipelinesReleases_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
        Write-Host "Organization: $org"
        Write-Host "Projects:"

        $projects = $projectsHashtable[$org].Projects
        $ids = $projectsHashtable[$org].Ids

        for ($i = 0; $i -lt $projects.Count; $i++) {
            $project = $projects[$i]
            $id = $ids[$i]
            Write-Host "  $project with id: $id"
            
            # Execute the Azure CLI command to get the JSON response as a string
            $jsonString = az pipelines release list --project "$id" --organization "https://dev.azure.com/$org"

            # Convert the JSON string to objects with correct encoding
            $jsonResponse = $jsonString | ConvertFrom-Json -Depth 100

            # Convert the response objects back to JSON with corrected encoding
            $correctedJsonString = $jsonResponse | ConvertTo-Json -Depth 100

            # Replace incorrect characters in the JSON string
            $correctedJsonString = Invoke-ReplaceCharacters -jsonString $correctedJsonString
            
            # Convert the corrected JSON string back to objects
            $jsonResponse = $correctedJsonString | ConvertFrom-Json

            # Extract specific fields from the JSON response
            $output = foreach ($PipelinesReleases in $jsonResponse) {
                [PSCustomObject]@{
                    "ReleaseName"               = $PipelinesReleases.name
                    "AuthoredByName"            = $PipelinesReleases.createdBy.displayName
                    "AuthoredByEmail"           = $PipelinesReleases.createdBy.uniqueName
                    "CreatedOn"                 = $PipelinesReleases.createdOn
                    "ModifiedByName"            = $PipelinesReleases.modifiedBy.displayName
                    "ModifiedByEmail"           = $PipelinesReleases.modifiedBy.uniqueName
                    "ProjectName"               = $PipelinesReleases.projectReference.name
                    "releaseDefinitionName"     = $PipelinesReleases.releaseDefinition.name
                    "releaseDefinitionRevision" = $PipelinesReleases.releaseDefinitionRevision
                    "releaseNameFormat"         = $PipelinesReleases.releaseNameFormat
                    "Status"                    = $PipelinesReleases.status
                    "ExecutionDate"             = $ExecutionDate
                }
            }
            
            # Export Pipelines Releases to Excel
            $output | Export-Excel -Path "PipelinesReleases/PipelinesReleases_$org.xlsx" -WorksheetName "$project" -AutoSize -AutoFilter -FreezeTopRow

            # Append project Pipelines Releases to all Pipelines
            $allPipelinesReleases += $output
        }
        # Export all Pipelines Releases to "All Information" worksheet
        $allPipelinesReleases | Export-Excel -Path "PipelinesReleases/PipelinesReleases_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
    }
}

function Get-PipelinesBuilds {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )

    # Retrieve projects linked to each organization along with IDs
    $projectsHashtable = Invoke-RetrieveDevOpsProjects -Filter $Filter

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    # Process each organization and its associated projects
    foreach ($org in $projectsHashtable.Keys) {
        # Initialize an array to store Pipelines Builds information
        $allPipelinesBuilds = @()

        Export-Excel -Path "PipelinesBuilds/PipelinesBuilds_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
        Write-Host "Organization: $org"
        Write-Host "Projects:"

        $projects = $projectsHashtable[$org].Projects
        $ids = $projectsHashtable[$org].Ids

        for ($i = 0; $i -lt $projects.Count; $i++) {
            $project = $projects[$i]
            $id = $ids[$i]
            Write-Host "  $project with id: $id"
            
            # Execute the Azure CLI command to get the JSON response as a string
            $jsonString = az pipelines build list --project "$id" --organization "https://dev.azure.com/$org"

            # Convert the JSON string to objects with correct encoding
            $jsonResponse = $jsonString | ConvertFrom-Json -Depth 100

            # Convert the response objects back to JSON with corrected encoding
            $correctedJsonString = $jsonResponse | ConvertTo-Json -Depth 100

            # Replace incorrect characters in the JSON string
            $correctedJsonString = Invoke-ReplaceCharacters -jsonString $correctedJsonString
            
            # Convert the corrected JSON string back to objects
            $jsonResponse = $correctedJsonString | ConvertFrom-Json

            # Extract specific fields from the JSON response
            $output = foreach ($PipelinesBuilds in $jsonResponse) {
                [PSCustomObject]@{
                    "PipelineName"      = $PipelinesBuilds.definition.name
                    "BuildNumber"       = $PipelinesBuilds.buildNumber
                    "BuildProjectName"  = $PipelinesBuilds.definition.project.name
                    "BuildProjectID"    = $PipelinesBuilds.definition.project.id
                    "QueueStatus"       = $PipelinesBuilds.definition.queueStatus
                    "PipelineStatus"    = $PipelinesBuilds.status
                    "LastChangedDate"   = $PipelinesBuilds.lastChangedDate
                    "AgentPool"         = $PipelinesBuilds.queue.pool.name
                    "BuildReason"       = $PipelinesBuilds.reason
                    "SourceBranch"      = $PipelinesBuilds.sourceBranch
                    "SourceVersion"     = $PipelinesBuilds.sourceVersion
                    "RepositoryName"    = $PipelinesBuilds.repository.name
                    "RequestedByName"   = $PipelinesBuilds.requestedBy.displayName
                    "RequestedByEmail"  = $PipelinesBuilds.requestedBy.uniqueName
                    "RequestedForName"  = $PipelinesBuilds.requestedFor.displayName
                    "RequestedForEmail" = $PipelinesBuilds.requestedFor.uniqueName
                    "TriggerMessage"    = $PipelinesBuilds.triggerInfo.'ci.message'
                    "ExecutionDate"     = $ExecutionDate
                }
            }
            
            # Export Pipelines Builds to Excel
            $output | Export-Excel -Path "PipelinesBuilds/PipelinesBuilds_$org.xlsx" -WorksheetName "$project" -AutoSize -AutoFilter -FreezeTopRow

            # Append project Pipelines Builds to all Pipelines Builds
            $allPipelinesBuilds += $output
        }
        # Export all Pipelines Builds to "All Information" worksheet
        $allPipelinesBuilds | Export-Excel -Path "PipelinesBuilds/PipelinesBuilds_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
    }
}

function Get-BranchesGIT {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )

    # Retrieve projects linked to each organization along with IDs
    $projectsHashtable = Invoke-RetrieveDevOpsProjects -Filter $Filter

    Write-Host "Retrieved Organizations: $($projectsHashtable.Count)"

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    # Process each organization and its associated projects
    foreach ($org in $projectsHashtable.Keys) {
        # Initialize an array to store all Repositories Branches information
        $allRepoBranches = @()

        Export-Excel -Path "ReposBranchesGIT/ReposBranchesGIT_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
        Write-Host "Processing Organization: $org"
        Write-Host "Projects:"

        $projects = $projectsHashtable[$org].Projects
        $ids = $projectsHashtable[$org].Ids

        # Process each project within the organization
        for ($i = 0; $i -lt $projects.Count; $i++) {
            $project = $projects[$i]
            $id = $ids[$i]
            Write-Host "  Project: $project with id: $id"

            Write-Host "Fetching Branches for Project: $project"

            # Get Repositories for the current project
            $repositories = Invoke-RetrieveDevOpsRepositories -Organization $org -ProjectId $id

            # Initialize an array to store Repositories Branches information for this project
            $repoBranches = @() 

            # Iterate over each Repository to fetch Repositories Branches
            foreach ($repo in $repositories) {
                $repoId = $repo.RepoID

                # Execute the Azure CLI command to get Repository details
                $jsonString1 = az repos show --repository $repoId --project $id --organization "https://dev.azure.com/$org"
                
                # Convert the JSON string to objects with correct encoding
                $jsonResponse1 = $jsonString1 | ConvertFrom-Json -Depth 100

                # Convert the response objects back to JSON with corrected encoding
                $correctedJsonString1 = $jsonResponse1 | ConvertTo-Json -Depth 100

                # Replace incorrect characters in the JSON string
                $correctedJsonString1 = Invoke-ReplaceCharacters -jsonString $correctedJsonString1

                # Convert the corrected JSON string back to objects
                $jsonResponse1 = $correctedJsonString1 | ConvertFrom-Json

                # Extract Repository name from Repository details
                $repoName = $jsonResponse1.name

                # Execute the Azure CLI command to get Repositories Branches
                $jsonString2 = az repos ref list --repository $repoId --project $id --organization "https://dev.azure.com/$org"

                # Convert the JSON string to objects with correct encoding
                $jsonResponse2 = $jsonString2 | ConvertFrom-Json -Depth 100

                # Convert the response objects back to JSON with corrected encoding
                $correctedJsonString2 = $jsonResponse2 | ConvertTo-Json -Depth 100

                # Replace incorrect characters in the JSON string
                $correctedJsonString2 = Invoke-ReplaceCharacters -jsonString $correctedJsonString2

                # Convert the corrected JSON string back to objects
                $repoBranchesArray = $correctedJsonString2 | ConvertFrom-Json

                # Extract specific fields and add to repoBranchesArray array
                foreach ($repoBranch in $repoBranchesArray) {
                    $repoBranches += [PSCustomObject]@{
                        "BranchName"      = $repoBranch.name
                        "RepoName"        = $repoName
                        "ProjectName"     = $project
                        "CreatorName"     = $repoBranch.creator.displayName
                        "CreatorEmail"    = $repoBranch.creator.uniqueName
                        "ExecutionDate"   = $ExecutionDate
                    }
                }
            }

            # Export Repositories Branches to Excel
            $repoBranches | Export-Excel -Path "ReposBranchesGIT/ReposBranchesGIT_$org.xlsx" -WorksheetName $project -AutoSize -AutoFilter -FreezeTopRow

            # Append project Repositories Branches to all Repositories Branches
            $allRepoBranches += $repoBranches
        }
        # Export all Repositories Branches to "All Information" worksheet
        $allRepoBranches | Export-Excel -Path "ReposBranchesGIT/ReposBranchesGIT_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
    }
    Write-Host "Exporting Branches to Excel completed."
}

function Get-BranchesTFVC {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )
    # Retrieve projects linked to each organization along with IDs
    $projectsHashtable = Invoke-RetrieveDevOpsProjects -Filter $Filter

    Write-Host "Retrieved Organizations: $($projectsHashtable.Count)"

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    # Process each organization and its associated projects
    foreach ($org in $projectsHashtable.Keys) {
        # Initialize an array to store all Repositories Branches information
        $allrepoBranchesTFVC = @()

        Export-Excel -Path "ReposBranchesTFVC/ReposBranchesTFVC_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow

        Write-Host "Processing Organization: $org"
        Write-Host "Projects:"

        $projects = $projectsHashtable[$org].Projects
        $ids = $projectsHashtable[$org].Ids

        # Process each project within the organization
        for ($i = 0; $i -lt $projects.Count; $i++) {
            $project = $projects[$i]
            $id = $ids[$i]
            Write-Host "  Project: $project with id: $id"

            Write-Host "Fetching Branches for Project: $project"

            # Initialize an array to store Repositories Branches information for this project
            $repoBranchesTFVC = @() 
            
            if ($null -eq $AzureDevOpsAuthenicationHeader) {
                $AzureDevOpsAuthenicationHeader = Invoke-AzureDevOpsAuth
            }

            $bodyRequest = @{}

            $bodyRequest = $bodyRequest | ConvertTo-Json -Depth 100
         
            $uriGetListProjects = "https://dev.azure.com/$org/$id/_apis/tfvc/branches?api-version=7.1-preview.1"

            # Execute the Azure CLI command to get Repositories Branches
            $jsonString2 = Invoke-RestMethod -Uri $uriGetListProjects -Method Get -Headers $AzureDevOpsAuthenicationHeader -Body $bodyRequest -ContentType "application/json"
            if ($jsonString2.count -eq 0) {
                Write-Host "No Branches TFVC Found"
            }
            else {
                # # Convert the corrected JSON string back to objects
                $repoBranchesTFVCArray = $jsonString2.value

                # Extract specific fields and add to repoBranchesTFVCArray array
                foreach ($repoBranchTFVC in $repoBranchesTFVCArray) {
                    $repoBranchesTFVC += [PSCustomObject]@{
                        "ProjectName"     = $project
                        "RepoName"        = $repoBranchTFVC.path.Split("/")[1]
                        "BranchName"      = $repoBranchTFVC.path.Split("/")[-1]
                        "BranchPath"      = $repoBranchTFVC.path
                        "CreatorName"     = $repoBranchTFVC.owner.displayName
                        "CreatorEmail"    = $repoBranchTFVC.owner.uniqueName
                        "CreatedDate"     = $repoBranchTFVC.createdDate
                        "ExecutionDate"   = $ExecutionDate
                    }
                }
            }

            # Export Repositories Branches to Excel
            $repoBranchesTFVC | Export-Excel -Path "ReposBranchesTFVC/ReposBranchesTFVC_$org.xlsx" -WorksheetName $project -AutoSize -AutoFilter -FreezeTopRow

            # Append project Repositories Branches to all Repositories Branches
            $allrepoBranchesTFVC += $repoBranchesTFVC
        }
        # Export all Repositories Branches to "All Information" worksheet
        $allrepoBranchesTFVC | Export-Excel -Path "ReposBranchesTFVC/ReposBranchesTFVC_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
    }
    Write-Host "Exporting TFVC Branches to Excel completed."
}

function Get-AllBranches {
    Param (
        [Parameter(Mandatory = $false)]
        [string]$Filter = "All"
    )

    # Retrieve projects linked to each organization along with IDs
    $projectsHashtable = Invoke-RetrieveDevOpsProjects -Filter $Filter

    Write-Host "Retrieved Organizations: $($projectsHashtable.Count)"

    $ExecutionDate = Get-Date -Format "dd/MM/yyyy"

    # Process each organization and its associated projects
    foreach ($org in $projectsHashtable.Keys) {
        # Initialize arrays to store all branches information
        $allBranches = @()

        # Initialize Excel file for all branches
        Export-Excel -Path "AllBranches/AllBranches_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow

        Write-Host "Processing Organization: $org"
        Write-Host "Projects:"

        $projects = $projectsHashtable[$org].Projects
        $ids = $projectsHashtable[$org].Ids

        # Process each project within the organization
        for ($i = 0; $i -lt $projects.Count; $i++) {
            $project = $projects[$i]
            $id = $ids[$i]
            Write-Host "  Project: $project with id: $id"

            # Initialize arrays to store branches information for this project
            $branches = @()

            # Fetch Git branches
            Write-Host "Fetching Git branches for Project: $project"

            # Get Repositories for the current project
            $repositories = Invoke-RetrieveDevOpsRepositories -Organization $org -ProjectId $id

            # Iterate over each Repository to fetch branches
            foreach ($repo in $repositories) {
                $repoId = $repo.RepoID
                $repoName = $repo.RepoName

                # Execute the Azure CLI command to get branches
                $jsonString1 = az repos ref list --repository $repoId --project $id --organization "https://dev.azure.com/$org"

                # Convert the JSON string to objects with correct encoding
                $jsonResponse1 = $jsonString1 | ConvertFrom-Json -Depth 100

                # Convert the response objects back to JSON with corrected encoding
                $correctedJsonString1 = $jsonResponse1 | ConvertTo-Json -Depth 100

                # Replace incorrect characters in the JSON string
                $correctedJsonString1 = Invoke-ReplaceCharacters -jsonString $correctedJsonString1

                # Convert the corrected JSON string back to objects
                $repoBranchesArray = $correctedJsonString1 | ConvertFrom-Json

                # Extract specific fields and add to branches array
                foreach ($repoBranch in $repoBranchesArray) {
                    $branches += [PSCustomObject]@{
                        "BranchName"      = $repoBranch.name
                        "RepoName"        = $repoName
                        "ProjectName"     = $project
                        "CreatorName"     = $repoBranch.creator.displayName
                        "CreatorEmail"    = $repoBranch.creator.uniqueName
                        "ExecutionDate"   = $ExecutionDate
                        "Source"          = "Git"
                    }
                }
            }

            if ($null -eq $AzureDevOpsAuthenicationHeader) {
                $AzureDevOpsAuthenicationHeader = Invoke-AzureDevOpsAuth
            }

            # Fetch TFVC branches
            Write-Host "Fetching TFVC branches for Project: $project"

            $bodyRequest = @{}

            $bodyRequest = $bodyRequest | ConvertTo-Json -Depth 100

            # Execute the Azure DevOps REST API to get TFVC branches
            $uriGetListProjects = "https://dev.azure.com/$org/$id/_apis/tfvc/branches?api-version=7.1-preview.1"
            $jsonString2 = Invoke-RestMethod -Uri $uriGetListProjects -Method Get -Headers $AzureDevOpsAuthenicationHeader -Body $bodyRequest -ContentType "application/json"

            # Extract specific fields and add to branches array
            if ($jsonString2.count -ne 0) {
                foreach ($repoBranchTFVC in $jsonString2.value) {
                    $branches += [PSCustomObject]@{
                        "BranchName"      = $repoBranchTFVC.path.Split("/")[-1]
                        "RepoName"        = $repoBranchTFVC.path.Split("/")[1]
                        "ProjectName"     = $project
                        "CreatorName"     = $repoBranchTFVC.owner.displayName
                        "CreatorEmail"    = $repoBranchTFVC.owner.uniqueName
                        "ExecutionDate"   = $ExecutionDate
                        "Source"          = "TFVC"
                    }
                }
            }

            # Export branches to Excel
            $branches | Export-Excel -Path "AllBranches/AllBranches_$org.xlsx" -WorksheetName $project -AutoSize -AutoFilter -FreezeTopRow

            # Append project branches to all branches
            $allBranches += $branches
        }

        # Export all branches to "All Information" worksheet
        $allBranches | Export-Excel -Path "AllBranches/AllBranches_$org.xlsx" -WorksheetName "All Information" -AutoSize -AutoFilter -FreezeTopRow
    }
    Write-Host "Exporting branches to Excel completed."
}

function Get-EverythingByProject {
    Get-ServiceConnections
    Get-RepositoriesGIT
    Get-RepositoriesTFVC
    Get-AllRepositories
    Get-RepositoriesPRs
    Get-Teams
    Get-TeamsMembers
    Get-VariablesGroups
    Get-Pipelines
    Get-PipelinesReleases
    Get-PipelinesBuilds
    Get-BranchesGIT
    Get-BranchesTFVC
    Get-AllBranches
}

Export-ModuleMember -Function Get-EverythingByProject, Get-ServiceConnections, Get-RepositoriesGIT, Get-RepositoriesTFVC, Get-AllRepositories, Get-RepositoriesPRs, Get-Teams, Get-TeamsMembers, Get-VariablesGroups, Get-Pipelines, Get-PipelinesReleases, Get-PipelinesBuilds, Get-BranchesGIT, Get-BranchesTFVC, Get-AllBranches
