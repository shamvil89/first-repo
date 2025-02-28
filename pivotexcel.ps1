# Define variables
$SourceFile = "D:\Desktop\Azure_Virtual_Machine_Inventory.xlsx"
$NewFile = "D:\Desktop\Output.xlsx"

# Define sheets to combine
$SheetsToInclude = @(
    "Azure_Virtual_Machine_Inventory",
    "Sheet1",   # Add your sheet names here
    "Sheet2"
)

# Clean up existing output file
if (Test-Path $NewFile) {
    Remove-Item -Path $NewFile -Force
}

# Combine specified sheets into one
try {
    $CombinedData = @()
    
    foreach ($SheetName in $SheetsToInclude) {
        Write-Host "Importing data from sheet: $SheetName"
        try {
            $Data = Import-Excel -Path $SourceFile -WorksheetName $SheetName
            $CombinedData += $Data
        } catch {
            Write-Host "WARNING: Could not import sheet '$SheetName'. Error: $_"
        }
    }
    
    # Export combined data to the new file
    if ($CombinedData.Count -gt 0) {
        $CombinedData | Export-Excel -Path $NewFile -WorksheetName "Combined_Data" -AutoSize -TableStyle Medium2
        Write-Host "Successfully created Combined_Data sheet in new file"
        $AllData = $CombinedData  # Use combined data for further processing
    } else {
        Write-Host "ERROR: No data found in specified sheets"
        exit
    }
} catch {
    Write-Host "ERROR: Failed to combine Excel sheets. Error: $_"
    exit
}

# Sheet names
$SummarySheet = "Summary"
$BillableSheet = "Billable VMs"
$InScopeSheet = "InScope VMs"
$AllVMsSheet = "All VMs"

# Filter data for each sheet
$BillableVMs = $AllData | Where-Object { $_.Billable -eq "Yes" }
$InScopeVMs = $AllData | Where-Object { $_.InScope -eq "Yes" }

try {
    # Create Summary sheet first (empty for now)
    @{} | Export-Excel -Path $NewFile -WorksheetName $SummarySheet -ClearSheet

    # Create Billable VMs sheet with count
    $BillableCount = $BillableVMs.Count
    Write-Host "Creating Billable VMs sheet with $BillableCount rows..."
    $BillableVMs | Export-Excel -Path $NewFile -WorksheetName "$BillableSheet - $BillableCount" `
        -AutoSize -TableName "BillableTable" -TableStyle Medium2 -Append

    # Create InScope VMs sheet with count
    $InScopeCount = $InScopeVMs.Count
    Write-Host "Creating InScope VMs sheet with $InScopeCount rows..."
    $InScopeVMs | Export-Excel -Path $NewFile -WorksheetName "$InScopeSheet - $InScopeCount" `
        -AutoSize -TableName "InScopeTable" -TableStyle Medium2 -Append

    # Create All VMs sheet with count
    $AllVMsCount = $AllData.Count
    Write-Host "Creating All VMs sheet with $AllVMsCount rows..."
    $AllData | Export-Excel -Path $NewFile -WorksheetName "$AllVMsSheet - $AllVMsCount" `
        -AutoSize -TableName "AllVMsTable" -TableStyle Medium2 -Append

    # Update sheet name variables to match new names with counts
    $BillableSheet = "$BillableSheet - $BillableCount"
    $InScopeSheet = "$InScopeSheet - $InScopeCount"
    $AllVMsSheet = "$AllVMsSheet - $AllVMsCount"

    # Now create pivot table
    $ExcelPackage = Open-ExcelPackage -Path $NewFile
    $ws = $ExcelPackage.Workbook.Worksheets[$SummarySheet]
    
    # Clear any existing content in the Summary sheet
    $ws.Cells.Clear()
    
    $PivotParams = @{
        PivotTableName = "PivotTable1"
        Address = $ws.Cells["A1"]
        SourceWorksheet = $ExcelPackage.Workbook.Worksheets[$InScopeSheet]
        PivotRows = "Status"
        PivotColumn = "OS Type"
        PivotData = @{"VM Name" = "Count"}
        PivotDataToColumn = $true
        NoTotalsInPivot = $false
        PivotTotals = "Both"
        PivotNumberFormat = "0"
        PivotTableStyle = "Light16"
    }

    # Generate Pivot Table
    Add-PivotTable @PivotParams -ExcelPackage $ExcelPackage
    
    # Additional pivot table formatting
    $pivotTable = $ws.PivotTables[0]
    
    # Format the pivot table
    $ws.View.ShowGridLines = $false
    $ws.Column(1).Width = 20
    
    # Additional formatting for better readability
    $ws.Cells["A1:Z1"].Style.Font.Bold = $true
    $ws.Cells.Style.HorizontalAlignment = 'Center'
    $ws.Column(1).Style.HorizontalAlignment = 'Left'

    # Rearrange sheets in desired order
    try {
        # Move Summary to position 1
        $ExcelPackage.Workbook.Worksheets.MoveToStart($SummarySheet)
        
        # Move other sheets in order
        $ExcelPackage.Workbook.Worksheets.MoveBefore($BillableSheet, $InScopeSheet)
        $ExcelPackage.Workbook.Worksheets.MoveBefore($InScopeSheet, $AllVMsSheet)

        Write-Host "Successfully reordered worksheets"
    } catch {
        Write-Host "WARNING: Unable to reorder worksheets. Error: $_"
        # Continue execution even if reordering fails
    }

    # Remove the Combined_Data sheet since it's no longer needed
    try {
        $CombinedSheet = $ExcelPackage.Workbook.Worksheets["Combined_Data"]
        if ($CombinedSheet) {
            $ExcelPackage.Workbook.Worksheets.Delete("Combined_Data")
            Write-Host "Removed temporary Combined_Data sheet"
        }
    } catch {
        Write-Host "WARNING: Unable to remove Combined_Data sheet. Error: $_"
    }

    # Save and close package
    Close-ExcelPackage $ExcelPackage -Show

    Write-Host "SUCCESS: Excel file created with all sheets"
} catch {
    Write-Host "ERROR: Unable to create workbook. Error: $_"
    if ($ExcelPackage) {
        Close-ExcelPackage $ExcelPackage -NoSave
    }
    exit
}