# Pipeline Cloner PowerShell Script

# Parameters
param(
    [Parameter(Mandatory=$false)]
    [string]$OrganizationUrl,

    [Parameter(Mandatory=$false)]
    [string]$ProjectName,

    [Parameter(Mandatory=$false)]
    [string]$SourcePipelineName,

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

# Function to get source pipeline name
function Get-SourcePipelineName {
    $pipelineName = $SourcePipelineName
    if (-not $pipelineName) {
        $pipelineName = Read-Host "Enter the source pipeline name"
    }
    return $pipelineName
}

# Function to get new pipeline name
function Get-NewPipelineName {
    $newName = $NewPipelineName
    if (-not $newName) {
        $newName = Read-Host "Enter the name for the new pipeline"
    }
    return $newName
}

# Function to get pipeline details from name
function Get-PipelineIdFromName {
    param (
        [string]$OrgUrl,
        [string]$Project,
        [string]$PipelineName,
        [hashtable]$Headers
    )
    
    # Get all pipeline definitions
    $listUrl = "$OrgUrl/$Project/_apis/build/definitions?api-version=7.1&name=$PipelineName"
    Write-Host "Looking up pipeline '$PipelineName'..."
    $response = Invoke-RestMethod -Uri $listUrl -Headers $Headers -Method Get

    if ($response.count -eq 0) {
        throw "No pipeline found with name: $PipelineName"
    }
    elseif ($response.count -gt 1) {
        Write-Host "Multiple pipelines found with name '$PipelineName'. Listing all matches:" -ForegroundColor Yellow
        Write-Host "----------------------------------------"
        foreach ($def in $response.value) {
            $pipelineType = if ($def.process.type -eq 1) { "GUI-based (Classic)" } else { "YAML-based" }
            Write-Host "Name: $($def.name)"
            Write-Host "ID: $($def.id)"
            Write-Host "Type: $pipelineType"
            Write-Host "Path: $($def.path)"
            Write-Host "Created: $($def.createdDate)"
            Write-Host "----------------------------------------"
        }
        
        Write-Host "Using the first pipeline in the list." -ForegroundColor Yellow
    }

    $selectedPipeline = $response.value[0]
    $pipelineType = if ($selectedPipeline.process.type -eq 1) { "GUI-based (Classic)" } else { "YAML-based" }
    
    Write-Host "Selected pipeline details:" -ForegroundColor Cyan
    Write-Host "- Name: $($selectedPipeline.name)"
    Write-Host "- Type: $pipelineType"
    Write-Host "- ID: $($selectedPipeline.id)"
    Write-Host "- Path: $($selectedPipeline.path)"
    
    return $selectedPipeline.id
}

# Function to preserve GUI-specific properties
function Update-PipelineDefinition {
    param (
        [object]$Definition,
        [string]$NewName
    )
    
    # Remove properties that should not be copied
    $Definition.PSObject.Properties.Remove('id')
    $Definition.PSObject.Properties.Remove('revision')
    $Definition.name = $NewName
    
    # Handle GUI-based pipeline specific properties
    if ($Definition.process.type -eq 1) {
        Write-Host "Processing GUI-based pipeline properties..." -ForegroundColor Cyan
        
        # Preserve important GUI pipeline properties
        if ($Definition.PSObject.Properties['quality']) {
            Write-Host "- Preserving build quality settings"
        }
        if ($Definition.PSObject.Properties['jobAuthorizationScope']) {
            Write-Host "- Preserving job authorization scope"
        }
        if ($Definition.PSObject.Properties['jobTimeoutInMinutes']) {
            Write-Host "- Preserving job timeout settings"
        }
        
        # Preserve task groups if present
        if ($Definition.process.phases) {
            foreach ($phase in $Definition.process.phases) {
                if ($phase.steps) {
                    Write-Host "- Preserving task group configurations"
                }
            }
        }
    }
    
    return $Definition
}

try {
    # Get required information
    $pat = Get-AzureDevOpsPat
    $orgUrl = Get-OrgUrl
    $project = Get-ProjectName
    $sourcePipelineName = Get-SourcePipelineName
    $newPipelineName = Get-NewPipelineName

    # Create authorization header
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($pat)"))
    $headers = @{
        Authorization = "Basic $base64AuthInfo"
        'Content-Type' = 'application/json'
    }

    # Get the source pipeline ID from name
    $sourcePipelineId = Get-PipelineIdFromName -OrgUrl $orgUrl -Project $project -PipelineName $sourcePipelineName -Headers $headers

    # Get the source pipeline definition
    $getUrl = "$orgUrl/$project/_apis/build/definitions/$($sourcePipelineId)?api-version=7.1"
    Write-Host "Getting source pipeline definition..."
    $sourcePipeline = Invoke-RestMethod -Uri $getUrl -Headers $headers -Method Get

    # Update the pipeline definition
    $newDefinition = Update-PipelineDefinition -Definition $sourcePipeline -NewName $newPipelineName
    
    # Convert to JSON
    $body = $newDefinition | ConvertTo-Json -Depth 100

    # Create the new pipeline
    $createUrl = "$orgUrl/$project/_apis/build/definitions?api-version=7.1"
    Write-Host "Creating new pipeline '$newPipelineName'..."
    $newPipeline = Invoke-RestMethod -Uri $createUrl -Headers $headers -Method Post -Body $body

    Write-Host "`nPipeline cloned successfully!" -ForegroundColor Green
    Write-Host "- New Pipeline Name: $($newPipeline.name)"
    Write-Host "- New Pipeline ID: $($newPipeline.id)"
    Write-Host "- Type: $(if ($newPipeline.process.type -eq 1) { 'GUI-based (Classic)' } else { 'YAML-based' })"
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Response: $($_.ErrorDetails.Message)" -ForegroundColor Red
    exit 1
} 