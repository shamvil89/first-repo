# Check if ImportExcel module is installed, if not install it
if (!(Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..."
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Import the module
Import-Module ImportExcel

function Merge-CSVToExcel {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SourceFolder,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputFile,
        
        [Parameter(Mandatory = $false)]
        [string]$FilePattern = "*.csv",

        [Parameter(Mandatory = $false)]
        [string]$Delimiter = ","
    )
    
    try {
        # Verify source folder exists
        if (!(Test-Path -Path $SourceFolder)) {
            throw "Source folder does not exist: $SourceFolder"
        }
        
        # Get all CSV files in the folder
        $csvFiles = Get-ChildItem -Path $SourceFolder -Filter $FilePattern
        
        if ($csvFiles.Count -eq 0) {
            throw "No CSV files found in $SourceFolder matching pattern $FilePattern"
        }
        
        Write-Host "Found $($csvFiles.Count) CSV files to merge"
        
        # Initialize array to store all data
        $allData = @()
        $firstFile = $true
        $columnHeaders = @()
        
        foreach ($file in $csvFiles) {
            Write-Host "Processing: $($file.Name)"
            
            try {
                # Import CSV file
                $data = Import-Csv -Path $file.FullName -Delimiter $Delimiter
                
                # Check column headers on first file
                if ($firstFile) {
                    $columnHeaders = $data[0].PSObject.Properties.Name
                    $firstFile = $false
                    Write-Host "Column headers found: $($columnHeaders -join ', ')"
                }
                else {
                    # Verify column headers match
                    $currentHeaders = $data[0].PSObject.Properties.Name
                    $headerDiff = Compare-Object -ReferenceObject $columnHeaders -DifferenceObject $currentHeaders
                    
                    if ($headerDiff) {
                        Write-Warning "Skipping $($file.Name) - Column headers do not match"
                        Write-Warning "Expected: $($columnHeaders -join ', ')"
                        Write-Warning "Found: $($currentHeaders -join ', ')"
                        continue
                    }
                }
                
                # Add data to collection
                $allData += $data
                Write-Host "Added $($data.Count) rows from $($file.Name)"
            }
            catch {
                Write-Warning "Error processing $($file.Name): $_"
                continue
            }
        }
        
        if ($allData.Count -eq 0) {
            throw "No data was collected from the CSV files"
        }
        
        # Create directory for output file if it doesn't exist
        $outputDirectory = Split-Path -Parent $OutputFile
        if (!(Test-Path -Path $outputDirectory)) {
            New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
        }
        
        # Export combined data to new Excel file
        Write-Host "Exporting $($allData.Count) total rows to $OutputFile"
        $allData | Export-Excel -Path $OutputFile -AutoSize -AutoFilter -WorksheetName "Merged_Data"
        
        Write-Host "Merge completed successfully!" -ForegroundColor Green
        Write-Host "Output file: $OutputFile"
    }
    catch {
        Write-Error "An error occurred: $_"
    }
}

# Interactive mode
$sourceFolder = Read-Host "Enter the folder path containing CSV files"
$outputFile = Read-Host "Enter the output Excel file path (including filename.xlsx)"
$delimiter = Read-Host "Enter CSV delimiter (press Enter for default ',')"
$filePattern = Read-Host "Enter file pattern (press Enter for default '*.csv')"

if ([string]::IsNullOrWhiteSpace($delimiter)) {
    $delimiter = ","
}

if ([string]::IsNullOrWhiteSpace($filePattern)) {
    $filePattern = "*.csv"
}

# Create the output directory if it doesn't exist
$outputDirectory = Split-Path -Parent $outputFile
if (!(Test-Path -Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

Merge-CSVToExcel -SourceFolder $sourceFolder -OutputFile $outputFile -FilePattern $filePattern -Delimiter $delimiter 