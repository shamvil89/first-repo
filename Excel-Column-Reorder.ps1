# Check if ImportExcel module is installed, if not install it
if (!(Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..."
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Import the module
Import-Module ImportExcel

function Reorder-ExcelColumns {
    param (
        [Parameter(Mandatory = $true)]
        [string]$InputFile,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputFile,
        
        [Parameter(Mandatory = $false)]
        [string]$WorksheetName = "Sheet1",

        [Parameter(Mandatory = $false)]
        [string[]]$ColumnOrder
    )
    
    try {
        # Verify input file exists
        if (!(Test-Path -Path $InputFile)) {
            throw "Input file does not exist: $InputFile"
        }
        
        # Import Excel file
        Write-Host "Reading Excel file: $InputFile"
        $data = Import-Excel -Path $InputFile -WorksheetName $WorksheetName
        
        if ($data.Count -eq 0) {
            throw "No data found in the Excel file"
        }
        
        # Get current column headers
        $currentHeaders = $data[0].PSObject.Properties.Name
        Write-Host "`nCurrent column order:"
        $currentHeaders | ForEach-Object { Write-Host "- $_" }
        
        # If no column order specified, prompt user to arrange columns
        if (!$ColumnOrder) {
            Write-Host "`nEnter the numbers corresponding to the columns in the desired order."
            Write-Host "Example: 3,1,4,2 (comma-separated)"
            
            # Display numbered list of columns
            for ($i = 0; $i -lt $currentHeaders.Count; $i++) {
                Write-Host "$($i + 1): $($currentHeaders[$i])"
            }
            
            $orderInput = Read-Host "`nEnter column order"
            $ColumnOrder = $orderInput.Split(',') | ForEach-Object { 
                $index = [int]$_ - 1
                $currentHeaders[$index]
            }
        }
        
        # Verify all columns exist
        $invalidColumns = $ColumnOrder | Where-Object { $_ -notin $currentHeaders }
        if ($invalidColumns) {
            throw "Invalid column names: $($invalidColumns -join ', ')"
        }
        
        # Add any columns that weren't specified to the end
        $remainingColumns = $currentHeaders | Where-Object { $_ -notin $ColumnOrder }
        $finalColumnOrder = $ColumnOrder + $remainingColumns
        
        Write-Host "`nNew column order:"
        $finalColumnOrder | ForEach-Object { Write-Host "- $_" }
        
        # Create new ordered object
        $reorderedData = $data | Select-Object $finalColumnOrder
        
        # Create directory for output file if it doesn't exist
        $outputDirectory = Split-Path -Parent $OutputFile
        if (!(Test-Path -Path $outputDirectory)) {
            New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
        }
        
        # Export reordered data
        Write-Host "`nExporting reordered data to: $OutputFile"
        $reorderedData | Export-Excel -Path $OutputFile -WorksheetName $WorksheetName -AutoSize -AutoFilter
        
        Write-Host "Column reordering completed successfully!" -ForegroundColor Green
        Write-Host "Output file: $OutputFile"
    }
    catch {
        Write-Error "An error occurred: $_"
    }
}

# Interactive mode
$inputFile = Read-Host "Enter the input Excel file path"
$outputFile = Read-Host "Enter the output Excel file path"
$worksheetName = Read-Host "Enter worksheet name (press Enter for default 'Sheet1')"

if ([string]::IsNullOrWhiteSpace($worksheetName)) {
    $worksheetName = "Sheet1"
}

# Example of specifying column order directly:
# $columnOrder = @("LastName", "FirstName", "Email", "Phone")
# Reorder-ExcelColumns -InputFile $inputFile -OutputFile $outputFile -WorksheetName $worksheetName -ColumnOrder $columnOrder

# Interactive column ordering
Reorder-ExcelColumns -InputFile $inputFile -OutputFile $outputFile -WorksheetName $worksheetName 