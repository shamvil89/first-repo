# Create Word application instance
$word = New-Object -ComObject Word.Application
$word.Visible = $true

# Create a new document
$document = $word.Documents.Add()

# Add a title
$selection = $word.Selection
$selection.Font.Size = 24
$selection.Font.Bold = $true
$selection.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
$selection.TypeText("Excel Files")
$selection.TypeParagraph()
$selection.TypeParagraph()

# Define the Excel files and icon
$excelFiles = @(
    "D:\Desktop\Azure_Virtual_Machine_Inventory.xlsx",
    "D:\Desktop\Output.xlsx"
)
$iconFile = "D:\Desktop\excel.ico"

# Verify icon file exists
if (-not (Test-Path $iconFile)) {
    Write-Error "Icon file not found at: $iconFile"
    exit
}

# Add each Excel file
foreach ($file in $excelFiles) {
    if (Test-Path $file) {
        try {
            # Create a new paragraph for the icon
            $selection.ParagraphFormat.Alignment = 1  # Center
            
            # Get filename for label
            $fileName = [System.IO.Path]::GetFileName($file)
            
            # Insert Excel file as icon with custom icon and filename as label
            $shape = $selection.InlineShapes.AddOLEObject(
                "Excel.Sheet",  # ClassType
                $file,         # FileName
                $false,        # LinkToFile
                $true,        # DisplayAsIcon
                $iconFile,    # IconFileName
                0,           # IconIndex
                $fileName    # IconLabel
            )
            
            # Add extra space after icon
            $selection.TypeParagraph()
            $selection.TypeParagraph()
            
            Write-Host "Successfully added $file"
        }
        catch {
            Write-Warning "Error adding file $file with custom icon: $_"
            try {
                # Try with default icon but still use filename as label
                $shape = $selection.InlineShapes.AddOLEObject(
                    "Excel.Sheet",  # ClassType
                    $file,         # FileName
                    $false,        # LinkToFile
                    $true,        # DisplayAsIcon
                    [Type]::Missing, # Default icon
                    0,            # IconIndex
                    $fileName     # IconLabel
                )
                
                # Add extra space after icon
                $selection.TypeParagraph()
                $selection.TypeParagraph()
                
                Write-Host "Successfully added $file with default icon"
            }
            catch {
                Write-Warning "Failed to add file with default icon: $file"
                Write-Warning "Error details: $_"
            }
        }
    }
    else {
        Write-Warning "File not found: $file"
    }
}

# Save the document
$documentPath = "D:\Desktop\ExcelAttachments.docx"
$document.SaveAs([ref]$documentPath)
Write-Host "Document saved as: $documentPath"

# Clean up COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers() 