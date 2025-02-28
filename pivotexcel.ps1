# Define variables
$SourceFile = "D:\Desktop\Azure_Virtual_Machine_Inventory.xlsx"
$NewFile = "D:\Desktop\Output.xlsx"

# Clean up existing files
if (Test-Path $NewFile) {
    Remove-Item -Path $NewFile -Force
}

# Sheet names
$SummarySheet = "Summary"
$BillableSheet = "Billable VMs"
$InScopeSheet = "InScope VMs"
$AllVMsSheet = "All VMs"

# Import original data
try {
    $AllData = Import-Excel -Path $SourceFile
    Write-Host "Successfully imported data from source file"
} catch {
    Write-Host "ERROR: Unable to read source file. Error: $_"
    exit
}

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
        SourceWorksheet = $ExcelPackage.Workbook.Worksheets[$AllVMsSheet]
        PivotRows = @("InScope","Status")
        PivotColumn = "OS Type"
        PivotData = @{"VM Name" = "Count"}
        PivotDataToColumn = $true
        NoTotalsInPivot = $false
        PivotTotals = "Rows"
        PivotNumberFormat = "1"
        PivotTableStyle = "Light16"
    }

    # Generate Pivot Table
    Add-PivotTable @PivotParams -ExcelPackage $ExcelPackage
    
    # Format Summary sheet
    $ws.View.ShowGridLines = $false
    $ws.Column(1).Width = 20
    
    # Additional formatting for better readability
    $ws.Cells["A1:Z1"].Style.Font.Bold = $true
    $ws.Cells.Style.HorizontalAlignment = 'Center'
    $ws.Column(1).Style.HorizontalAlignment = 'Left'

    # Rearrange sheets in desired order
    $ws_Summary = $ExcelPackage.Workbook.Worksheets[$SummarySheet]
    $ws_Billable = $ExcelPackage.Workbook.Worksheets[$BillableSheet]
    $ws_InScope = $ExcelPackage.Workbook.Worksheets[$InScopeSheet]

    # Move Summary to position 1
    if ($ws_Summary.Index -ne 1) {
        $ws_Summary.MoveBefore(1)
    }

    # Move Billable VM to position 2
    if ($ws_Billable.Index -ne 2) {
        $ws_Billable.MoveBefore(2)
    }

    # Save and close package
    Close-ExcelPackage $ExcelPackage -Show

    Write-Host "SUCCESS: Excel file created with all sheets in correct order"
} catch {
    Write-Host "ERROR: Unable to create workbook. Error: $_"
    if ($ExcelPackage) {
        Close-ExcelPackage $ExcelPackage -NoSave
    }
    exit
}