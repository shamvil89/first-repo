# Pipeline Cloner PowerShell Script

# Parameters
param(
    [Parameter(Mandatory=$false)]
    [string]$OrganizationUrl,

    [Parameter(Mandatory=$false)]
    [string]$ProjectName,

    [Parameter(Mandatory=$false)]
    [int]$SourcePipelineId,

    [Parameter(Mandatory=$false)]
    [string]$NewPipelineName,

    [Parameter(Mandatory=$false)]
    [string]$Pat
)

# Function to get Azure DevOps PAT from environment or user input
function Get-AzureDevOpsPat {
    $pat = $Pat
    if (-not $pat) {
        $pat = $env:AZURE_DEVOPS_PAT
        if (-not $pat) {
            $pat = Read-Host "Enter your Azure DevOps Personal Access Token" -AsSecureString
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pat)
            $pat = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        }
    }
    return $pat
}

# Function to get organization URL
function Get-OrgUrl {
    $url = $OrganizationUrl
    if (-not $url) {
        $url = $env:AZURE_DEVOPS_ORG_URL
        if (-not $url) {
            $url = Read-Host "Enter your Azure DevOps Organization URL (e.g., https://dev.azure.com/your-org)"
        }
    }
    return $url.TrimEnd('/')
}

# Function to get project name
function Get-ProjectName {
    $project = $ProjectName
    if (-not $project) {
        $project = Read-Host "Enter your Azure DevOps project name"
    }
    return $project
}

# Function to get source pipeline ID
function Get-SourcePipelineId {
    $pipelineId = $SourcePipelineId
    if (-not $pipelineId) {
        $pipelineId = Read-Host "Enter the source pipeline ID"
    }
    return $pipelineId
}

# Function to get new pipeline name
function Get-NewPipelineName {
    $newName = $NewPipelineName
    if (-not $newName) {
        $newName = Read-Host "Enter the name for the new pipeline"
    }
    return $newName
}

try {
    # Get required information
    $pat = Get-AzureDevOpsPat
    $orgUrl = Get-OrgUrl
    $project = Get-ProjectName
    $sourcePipelineId = Get-SourcePipelineId
    $newPipelineName = Get-NewPipelineName

    # Create authorization header
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($pat)"))
    $headers = @{
        Authorization = "Basic $base64AuthInfo"
        'Content-Type' = 'application/json'
    }

    # Get the source pipeline definition
    $getUrl = "$orgUrl/$project/_apis/build/definitions/$($sourcePipelineId)?api-version=7.1"
    Write-Host "Getting source pipeline definition..."
    $sourcePipeline = Invoke-RestMethod -Uri $getUrl -Headers $headers -Method Get

    # Modify the pipeline definition for the new pipeline
    $sourcePipeline.PSObject.Properties.Remove('id')
    $sourcePipeline.PSObject.Properties.Remove('revision')
    $sourcePipeline.name = $newPipelineName
    
    # Convert to JSON
    $body = $sourcePipeline | ConvertTo-Json -Depth 100

    # Create the new pipeline
    $createUrl = "$orgUrl/$project/_apis/build/definitions?api-version=7.1"
    Write-Host "Creating new pipeline '$newPipelineName'..."
    $newPipeline = Invoke-RestMethod -Uri $createUrl -Headers $headers -Method Post -Body $body

    Write-Host "Successfully created new pipeline: $($newPipeline.name) (ID: $($newPipeline.id))" -ForegroundColor Green
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Response: $($_.ErrorDetails.Message)" -ForegroundColor Red
    exit 1
} 