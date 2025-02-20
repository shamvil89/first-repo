# Define variables
$SourceFile = "D:\Desktop\Book1.xlsx"  # Input Excel file
$SheetNames = @("Sheet1", "Sheet2")    # Sheets to extract
$NewFile = "D:\Desktop\Output.xlsx"    # Output Excel file
$MergedSheetName = "MergedSheet"       # Sheet for merged data
$PivotSheetName = "PivotTableSheet"    # Pivot Table sheet
$TableName = "MergedDataTable"         # Table name for Pivot Table

# Define filtering condition
$FilterColumn = "col8"  # Column to filter
$FilterValue = @("row8", "row6")  # Values to keep

# Ensure ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ERROR: The ImportExcel module is not installed. Please install it using:"
    Write-Host "`n Install-Module -Name ImportExcel -Scope CurrentUser -Force"
    exit
}

# Check if source file exists
if (-Not (Test-Path $SourceFile)) {
    Write-Host "ERROR: Source file not found: $SourceFile"
    exit
}

# Initialize an empty array to store filtered data
$AllFilteredData = @()

foreach ($SheetName in $SheetNames) {
    try {
        # Import the sheet
        $SheetData = Import-Excel -Path $SourceFile -WorksheetName $SheetName -ErrorAction Stop

        # Check if data exists
        if ($null -eq $SheetData -or $SheetData.Count -eq 0) {
            Write-Host "WARNING: Sheet '$SheetName' is empty."
            continue
        }

        # Extract column names
        $ActualColumns = $SheetData[0].PSObject.Properties.Name
        Write-Host "Available columns in '$SheetName': $($ActualColumns -join ', ')"

        # Ensure the filter column exists
        if ($ActualColumns -notcontains $FilterColumn) {
            Write-Host "ERROR: Column '$FilterColumn' not found in '$SheetName'. Available columns: $($ActualColumns -join ', ')"
            continue
        }

        # Apply filtering and Debug print
        $FilteredData = $SheetData | Where-Object { $_.$FilterColumn -in $FilterValue }

        # Debug: Check how many rows are filtered
        Write-Host "DEBUG: Filtered $($FilteredData.Count) rows from '$SheetName'"

        # Check if filtered data is empty
        if ($FilteredData.Count -eq 0) {
            Write-Host "WARNING: No matching records found in '$SheetName' for '$FilterColumn'."
            continue
        }

        # Convert filtered data to text format to prevent formatting issues
        $FilteredData = $FilteredData | ForEach-Object {
            $_.PSObject.Properties | ForEach-Object { $_.Value = $_.Value -as [string] }
            $_
        }

        # Add to combined dataset
        $AllFilteredData += $FilteredData

    } catch {
        Write-Host "ERROR: Unable to read '$SheetName'. Ensure the file exists and the sheet name is correct. Error: $_"
    }
}

# Debug: Check count of filtered data before exporting
Write-Host "DEBUG: Total filtered records to export: $($AllFilteredData.Count)"

# Stop if no data found
if ($AllFilteredData.Count -eq 0) {
    Write-Host "ERROR: No matching data found. Stopping execution."
    exit
}

# Step 1: Export filtered data to Excel as a named table
try {
    Write-Host "INFO: Writing filtered data to '$MergedSheetName'..."
    $AllFilteredData | Export-Excel -Path $NewFile -WorksheetName $MergedSheetName -AutoSize -TableName $TableName -ClearSheet -ErrorAction Stop
    Write-Host "SUCCESS: Filtered data saved in '$MergedSheetName'"
} catch {
    Write-Host "ERROR: Unable to write to '$NewFile'. Ensure it is not open. Error: $_"
    exit
}

# Step 2: Create Pivot Table Sheet
try {
    Export-Excel -Path $NewFile -WorksheetName $PivotSheetName -AutoSize -ClearSheet
    Write-Host "SUCCESS: Created Pivot Table Sheet '$PivotSheetName'"
} catch {
    Write-Host "ERROR: Unable to create pivot sheet. Error: $_"
    exit
}

# Step 3: Create Pivot Table
try {
    # Define Pivot Table parameters
    $ExcelPackage = Open-ExcelPackage -Path $NewFile
    $ws = $ExcelPackage.Workbook.Worksheets[$PivotSheetName]

    $PivotParams = @{
        PivotTableName = "PivotTable1"
        Address = $ws.Cells["A1"]
        SourceWorksheet = $ExcelPackage.Workbook.Worksheets[$MergedSheetName]
        PivotRows = "col8"
        PivotData = @{"col6" = "Count"}
    }

    # Generate Pivot Table
    Add-PivotTable @PivotParams
    Close-ExcelPackage $ExcelPackage -Show

    Write-Host "SUCCESS: Pivot Table created in '$PivotSheetName'"
} catch {
    Write-Host "ERROR: Unable to create Pivot Table. Error: $_"
}