# Azure DevOps Information Extraction

This repository contains two PowerShell modules designed to help you extract and manage useful information from Azure DevOps. The scripts focus on listing and organizing data based on projects or organizations, making them useful for generating reports, analyzing data, or managing resources within Azure DevOps. Saving all of them in Excel in .xlsx format.

## How It Works

1. **ListByProject.psm1**  
    A PowerShell module that provides functions to list and manage items grouped by Azure DevOps projects. Core functionalities may include:

    - Extracting information such as work items, builds, or repositories associated with specific projects.
    - Filtering or sorting items based on project attributes (e.g., organization name, project name).
    - Aggregating data or generating summaries for each project.

    These are the functions inside the powershell module:
    - Get-EverythingByProject: Call all the functions together.
    - Get-ServiceConnections: List all the service connections.
    - Get-RepositoriesGIT: List all GIT repositories.
    - Get-RepositoriesTFVC: List all TFVC repositories.
    - Get-AllRepositories: List both TFVC and GIT repositories.
    - Get-RepositoriesPRs: List Repositorioes Pull Requests.
    - Get-Teams: List Azure DevOps Teams.
    - Get-TeamsMembers: List Members inside Azure DevOps Teams.
    - Get-VariablesGroups: List all variables groups and values inside. ***Can show sensitive data.***
    - Get-Pipelines: List all pipelines.
    - Get-PipelinesReleases: List all Pipelines Releases.
    - Get-PipelinesBuilds: List all Pipelines Builds.
    - Get-BranchesGIT: List all Branches in GIT format.
    - Get-BranchesTFVC: List all Branches in TFVC format.
    - Get-AllBranches: List all Branches in both TFVC and GIT format.

2. **ListByOrganization.psm1**  
    A PowerShell module that provides functions to list and manage items grouped by Azure DevOps organizations. The functionalities may include:

    - Extracting information such as users, teams, or agents associated with specific organizations.
    - Filtering or sorting items based on organizational attributes (e.g., organization name).
    - Aggregating data or generating summaries for each organization.

    These are the functions inside the powershell module:
    - Get-EverythingByOrganization: Call all the functions together.
    - Get-Users: List all Users.
    - Get-Groups: List all Groups.
    - Get-Projects: List all projects.
    - Get-AgentPools: List all AgentPools.

## How to Use

### Prerequisites

Ensure that you have the following:

- PowerShell installed on your machine (version 5.0 or higher recommended).
- Necessary permissions to execute scripts.
- User account with privileged permissions on the organizations to access the Azure DevOps data.

### 1. Update the Array of organizations

```powershell
# Array with the names of the organizations you want to work with.
$organizations = @("Organization1", "Organization2", "Organization3") 
```

### 2. Importing the Modules

Before using the functions provided in these modules, you'll need to import them into your PowerShell session.

```powershell
# Import ListByProject Module
Import-Module -Path "Path\To\ListByProject.psm1"

# Import ListByOrganization Module
Import-Module -Path "Path\To\ListByOrganization.psm1"
```
### 3. Function usage

All the functions are created to be used with or without the flag -Filter.

```powershell
# Just use all the organizations inside of the array
Get-EverythingByProject -Filter All

# This also means the same of using Get-EverythingByProject -Filter All
Get-EverythingByProject

# FIlter the organizations by using the flag -Filter 
Get-EverythingByProject - Filter 'Organization1, Organization2'
```

### Contributing

This is the first iteration of the PowerShell scripts designed to extract and manage data from Azure DevOps. I'm excited to share this project and hope it will be helpful in your work.

As this is an early version, i know there’s plenty of room for improvement. Your feedback, suggestions, and contributions would be incredibly valuable to help me refine and enhance these scripts.

Whether you spot a bug, have ideas for new features, or want to improve the existing code, I would love to hear from you. Every bit of help makes a difference!

Thank you for your support !
