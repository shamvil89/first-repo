# Check if ImportExcel module is installed, if not install it
if (!(Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..."
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Import the module
Import-Module ImportExcel

function Merge-ExcelFiles {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SourceFolder,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputFile,
        
        [Parameter(Mandatory = $false)]
        [string]$FilePattern = "*.xlsx"
    )
    
    try {
        # Verify source folder exists
        if (!(Test-Path -Path $SourceFolder)) {
            throw "Source folder does not exist: $SourceFolder"
        }
        
        # Get all Excel files in the folder
        $excelFiles = Get-ChildItem -Path $SourceFolder -Filter $FilePattern
        
        if ($excelFiles.Count -eq 0) {
            throw "No Excel files found in $SourceFolder matching pattern $FilePattern"
        }
        
        Write-Host "Found $($excelFiles.Count) Excel files to merge"
        
        # Initialize array to store all data
        $allData = @()
        $firstFile = $true
        $columnHeaders = @()
        
        foreach ($file in $excelFiles) {
            Write-Host "Processing: $($file.Name)"
            
            try {
                # Import Excel file
                $data = Import-Excel -Path $file.FullName -ErrorAction Stop
                
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
            throw "No data was collected from the Excel files"
        }
        
        # Export combined data to new Excel file
        Write-Host "Exporting $($allData.Count) total rows to $OutputFile"
        $allData | Export-Excel -Path $OutputFile -AutoSize -AutoFilter
        
        Write-Host "Merge completed successfully!" -ForegroundColor Green
        Write-Host "Output file: $OutputFile"
    }
    catch {
        Write-Error "An error occurred: $_"
    }
}

# Example usage:
# Merge-ExcelFiles -SourceFolder "C:\ExcelFiles" -OutputFile "C:\Merged\CombinedFile.xlsx"

# Interactive mode
$sourceFolder = Read-Host "Enter the folder path containing Excel files"
$outputFile = Read-Host "Enter the output file path (including filename.xlsx)"
$filePattern = Read-Host "Enter file pattern (default: *.xlsx)"

if ([string]::IsNullOrWhiteSpace($filePattern)) {
    $filePattern = "*.xlsx"
}

Merge-ExcelFiles -SourceFolder $sourceFolder -OutputFile $outputFile -FilePattern $filePattern 