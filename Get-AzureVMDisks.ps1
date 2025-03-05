# Check if required modules are installed
$requiredModules = @('Az.Accounts', 'Az.Compute')
foreach ($module in $requiredModules) {
    if (!(Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing $module module..."
        Install-Module -Name $module -Force -Scope CurrentUser
    }
    Import-Module $module -Force
}

function Get-AzureVMDiskInfo {
    param (
        [Parameter(Mandatory = $false)]
        [string]$OutputFile = ".\AzureVMDisks.xlsx"
    )
    
    try {
        # Connect to Azure if not already connected
        $context = Get-AzContext
        if (!$context) {
            Connect-AzAccount
        }
        
        Write-Host "Fetching all subscriptions..."
        $subscriptions = Get-AzSubscription
        
        # Initialize results array
        $results = @()
        
        foreach ($sub in $subscriptions) {
            Write-Host "`nProcessing Subscription: $($sub.Name)"
            Select-AzSubscription -SubscriptionId $sub.Id | Out-Null
            
            # Get all VMs in the subscription
            $vms = Get-AzVM
            
            if ($vms.Count -eq 0) {
                Write-Host "No VMs found in subscription $($sub.Name)"
                continue
            }
            
            Write-Host "Found $($vms.Count) VMs in subscription $($sub.Name)"
            
            foreach ($vm in $vms) {
                Write-Host "Processing VM: $($vm.Name) in Resource Group: $($vm.ResourceGroupName)"
                
                try {
                    # Get OS Disk details
                    $osDisk = Get-AzDisk -ResourceGroupName $vm.ResourceGroupName -DiskName $vm.StorageProfile.OsDisk.Name
                    
                    # Get VM Status
                    $vmStatus = (Get-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name -Status).Statuses |
                        Where-Object Code -like "PowerState*" |
                        Select-Object -ExpandProperty DisplayStatus
                    
                    # Add OS Disk to results
                    $results += [PSCustomObject]@{
                        'Subscription'    = $sub.Name
                        'SubscriptionId'  = $sub.Id
                        'ResourceGroup'   = $vm.ResourceGroupName
                        'VM_Name'        = $vm.Name
                        'VM_Size'        = $vm.HardwareProfile.VmSize
                        'VM_Status'      = $vmStatus
                        'Disk_Name'      = $osDisk.Name
                        'Disk_Type'      = 'OS Disk'
                        'Disk_SKU'       = $osDisk.Sku.Name
                        'Disk_Size_GB'   = $osDisk.DiskSizeGB
                        'Disk_State'     = $osDisk.DiskState
                        'Disk_Location'  = $osDisk.Location
                        'Encryption'     = $osDisk.Encryption.Type
                        'Created_Date'   = $osDisk.TimeCreated.ToString('yyyy-MM-dd')
                    }
                    
                    # Get Data Disk details
                    foreach ($dataDisk in $vm.StorageProfile.DataDisks) {
                        $disk = Get-AzDisk -ResourceGroupName $vm.ResourceGroupName -DiskName $dataDisk.Name
                        
                        $results += [PSCustomObject]@{
                            'Subscription'    = $sub.Name
                            'SubscriptionId'  = $sub.Id
                            'ResourceGroup'   = $vm.ResourceGroupName
                            'VM_Name'        = $vm.Name
                            'VM_Size'        = $vm.HardwareProfile.VmSize
                            'VM_Status'      = $vmStatus
                            'Disk_Name'      = $disk.Name
                            'Disk_Type'      = 'Data Disk'
                            'Disk_SKU'       = $disk.Sku.Name
                            'Disk_Size_GB'   = $disk.DiskSizeGB
                            'Disk_State'     = $disk.DiskState
                            'Disk_Location'  = $disk.Location
                            'Encryption'     = $disk.Encryption.Type
                            'Created_Date'   = $disk.TimeCreated.ToString('yyyy-MM-dd')
                        }
                    }
                }
                catch {
                    Write-Warning "Error processing VM $($vm.Name): $_"
                    continue
                }
            }
        }
        
        if ($results.Count -eq 0) {
            Write-Warning "No VM or disk data found in any subscription"
            return
        }
        
        # Create output directory if it doesn't exist
        $outputDirectory = Split-Path -Parent $OutputFile
        if (![string]::IsNullOrWhiteSpace($outputDirectory) -and !(Test-Path -Path $outputDirectory)) {
            New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
        }
        
        # Export to Excel if ImportExcel module is available
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Write-Host "`nExporting $($results.Count) records to Excel: $OutputFile"
            $results | Export-Excel -Path $OutputFile -WorksheetName 'VM_Disks' -AutoSize -AutoFilter -FreezeTopRow
            Write-Host "Excel report generated: $OutputFile"
        }
        else {
            # Export to CSV if ImportExcel is not available
            $csvPath = [System.IO.Path]::ChangeExtension($OutputFile, "csv")
            Write-Host "`nExporting $($results.Count) records to CSV: $csvPath"
            $results | Export-Csv -Path $csvPath -NoTypeInformation
            Write-Host "CSV report generated: $csvPath"
        }
        
        return $results
    }
    catch {
        Write-Error "An error occurred: $_"
        throw $_
    }
}

# Interactive mode
$outputFile = Read-Host "Enter output file path (press Enter for default './AzureVMDisks.xlsx')"

if ([string]::IsNullOrWhiteSpace($outputFile)) {
    $outputFile = ".\AzureVMDisks.xlsx"
}

# Call the function with provided parameters
Get-AzureVMDiskInfo -OutputFile $outputFile 