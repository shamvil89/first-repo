# Script Name: new_script.ps1
# Description: Add VLOOKUP and calculate total disk sizes for matching VMs
# Date: [Current date]

function Write-Log {
    param([string]$Message)
    Write-Host "$(Get-Date -Format 'HH:mm:ss'): $Message"
}

try {
    Write-Log "Starting script..."
    
    # Create Excel COM object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.DisplayAlerts = $false
    
    # Open disk report
    $diskPath = "D:\Desktop\Azure_Disk_Report_Updated.xlsx"
    Write-Log "Opening disk report: $diskPath"
    $diskWorkbook = $excel.Workbooks.Open($diskPath)
    $diskSheet = $diskWorkbook.Worksheets.Item(1)
    
    # Create temporary sheet for inventory data
    Write-Log "Creating temporary inventory sheet..."
    $tempSheet = $diskWorkbook.Worksheets.Add()
    $tempSheet.Name = "TempInventory"
    
    # Open and copy inventory data
    Write-Log "Copying inventory data..."
    $inventoryPath = "D:\Desktop\Azure_Virtual_Machine_Inventory.xlsx"
    $inventoryWorkbook = $excel.Workbooks.Open($inventoryPath)
    $inventorySheet = $inventoryWorkbook.Worksheets.Item(1)
    $inventoryRange = $inventorySheet.UsedRange
    $inventoryRange.Copy($tempSheet.Range("A1"))
    $inventoryWorkbook.Close($false)
    
    # Create new worksheet for report and copy disk data
    $newSheet = $diskWorkbook.Worksheets.Add()
    $newSheet.Name = "Report_With_Lookup"
    $usedRange = $diskSheet.UsedRange
    $usedRange.Copy($newSheet.Range("A1"))
    
    # Insert new column C
    $newSheet.Columns("C:C").Insert()
    $newSheet.Cells(1, 3).Value2 = "Lookup_Result"
    
    # Get last row
    $lastRow = $newSheet.UsedRange.Rows.Count
    Write-Log "Total rows to process: $lastRow"
    
    # Add VLOOKUP formula to entire column at once
    $range = $newSheet.Range("C2:C$lastRow")
    $formula = "=VLOOKUP(B2,TempInventory!`$A:`$A,1,0)"
    $range.Formula = $formula
    Write-Log "Added VLOOKUP formula: $formula"
    
    # Remove rows with #N/A
    Write-Log "Removing rows with #N/A..."
    for ($row = $lastRow; $row -ge 2; $row--) {
        $lookupValue = $newSheet.Cells($row, 3).Text
        if ($lookupValue -eq "#N/A") {
            Write-Log "Removing row $row - No match found"
            $newSheet.Rows($row).Delete()
        }
    }
    
    # Get new last row after deletions
    $lastRow = $newSheet.UsedRange.Rows.Count
    Write-Log "Remaining rows after filtering: $($lastRow - 1)"
    
    # Delete the Lookup_Result column (column C)
    Write-Log "Removing Lookup_Result column..."
    $newSheet.Columns("C:C").Delete()
    
    # Format as table
    Write-Log "Formatting as table..."
    $dataRange = $newSheet.UsedRange
    $table = $newSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $dataRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
    $table.TableStyle = "TableStyleLight9"
    
    # Check if disk sizes are already calculated
    Write-Log "Checking if disk sizes are already calculated..."
    $lastRow = $newSheet.UsedRange.Rows.Count
    $lastCell = $newSheet.Cells($lastRow, 1).Text
    
    if ($lastCell -match "Total Disk Size") {
        Write-Log "Disk sizes are already calculated. Skipping calculation..."
    } else {
        # Calculate total disk size
        Write-Log "Calculating total disk size..."
        $totalSize = 0
        
        # Get all values at once for processing
        $diskValues = $newSheet.Range("D2:AC$lastRow").Value2
        
        # Process values (now all rows are valid)
        for ($row = 1; $row -le $diskValues.GetLength(0); $row++) {
            $vmName = $newSheet.Cells($row + 1, 2).Text
            Write-Log "Processing VM: $vmName"
            
            for ($col = 1; $col -le $diskValues.GetLength(1); $col++) {
                $value = $diskValues[$row, $col]
                if ($value -match '^\d+\.?\d*$') {
                    $totalSize += [double]$value
                    Write-Log "Added disk size: $value GB"
                }
            }
        }
        
        Write-Log "Total disk size: $totalSize GB"
        
        # Write results
        $resultRow = $lastRow + 3
        
        # Write GB value
        $newSheet.Cells($resultRow, 1).Value2 = "Total Disk Size (GB):"
        $newSheet.Cells($resultRow, 2).Value2 = $totalSize
        $newSheet.Cells($resultRow, 2).NumberFormat = "#,##0.00"
        
        # Write TB value
        $tbValue = [math]::Round($totalSize / 1024, 2)
        $newSheet.Cells($resultRow + 1, 1).Value2 = "Total Disk Size (TB):"
        $newSheet.Cells($resultRow + 1, 2).Value2 = $tbValue
        $newSheet.Cells($resultRow + 1, 2).NumberFormat = "#,##0.00"
        
        # Format TB cell
        $tbCell = $newSheet.Cells($resultRow + 1, 2)
        $tbCell.Interior.ColorIndex = 6  # Yellow
        $tbCell.Font.Bold = $true
    }
    
    # AutoFit columns
    $newSheet.UsedRange.Columns.AutoFit()
    
    # Clean up extra sheets
    Write-Log "Cleaning up extra sheets..."
    $diskWorkbook.Sheets.Item("TempInventory").Delete()
    $diskWorkbook.Sheets.Item("Sheet1").Delete()
    
    # Move the report sheet to first position
    $newSheet.Move($diskWorkbook.Sheets.Item(1))
    
    # Save as new file
    $newPath = "D:\Desktop\Azure_Disk_Report_Updated_with_Totals.xlsx"
    Write-Log "Saving as: $newPath"
    
    if (Test-Path $newPath) {
        Remove-Item $newPath -Force
    }
    
    $diskWorkbook.SaveAs($newPath)
    $diskWorkbook.Close($false)
    $excel.Quit()
    
    # Cleanup
    Write-Log "Cleaning up..."
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($newSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($diskSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($diskWorkbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Log "Script completed successfully!"

} catch {
    Write-Error "Error at $(Get-Date -Format 'HH:mm:ss'): $($_.Exception.Message)"
} finally {
    Get-Process excel -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq "" } | Stop-Process -Force
} 