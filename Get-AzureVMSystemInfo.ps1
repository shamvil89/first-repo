# Check if required modules are installed
$requiredModules = @('Az.Accounts', 'Az.Compute', 'ImportExcel')
foreach ($module in $requiredModules) {
    if (!(Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing $module module..."
        Install-Module -Name $module -Force -Scope CurrentUser
    }
    Import-Module $module -Force
}

function Get-VMSystemInfo {
    param (
        [Parameter(Mandatory = $false)]
        [string]$OutputFile = ".\AzureVMSystemInfo.xlsx"
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
                    # Get VM Status
                    $vmStatus = (Get-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name -Status).Statuses |
                        Where-Object Code -like "PowerState*" |
                        Select-Object -ExpandProperty DisplayStatus
                    
                    if ($vmStatus -ne "VM running") {
                        Write-Host "VM $($vm.Name) is not running. Current status: $vmStatus"
                        
                        # Add non-running VM to results
                        $results += [PSCustomObject]@{
                            'Subscription'    = $sub.Name
                            'SubscriptionId'  = $sub.Id
                            'ResourceGroup'   = $vm.ResourceGroupName
                            'VM_Name'        = $vm.Name
                            'VM_Size'        = $vm.HardwareProfile.VmSize
                            'VM_Status'      = $vmStatus
                            'OS_Name'        = "VM not running"
                            'OS_Version'     = "VM not running"
                            'OS_Build'       = "VM not running"
                            'System_Type'    = "VM not running"
                        }
                        continue
                    }
                    
                    # Command to get system info
                    $command = "systeminfo | findstr /B /C:'OS'"
                    
                    # Try Windows command first
                    try {
                        $result = Invoke-AzVMRunCommand -ResourceGroupName $vm.ResourceGroupName -VMName $vm.Name `
                            -CommandId 'RunPowerShellScript' -ScriptString $command
                        
                        $osInfo = $result.Value[0].Message -split "`n"
                        
                        # Parse Windows systeminfo output
                        $osName = ($osInfo | Where-Object { $_ -like "*OS Name:*" }) -replace "OS Name:\s+", ""
                        $osVersion = ($osInfo | Where-Object { $_ -like "*OS Version:*" }) -replace "OS Version:\s+", ""
                        $osBuild = ($osInfo | Where-Object { $_ -like "*OS Build Type:*" }) -replace "OS Build Type:\s+", ""
                        $systemType = ($osInfo | Where-Object { $_ -like "*System Type:*" }) -replace "System Type:\s+", ""
                    }
                    catch {
                        # If Windows command fails, try Linux command
                        $command = "cat /etc/os-release"
                        $result = Invoke-AzVMRunCommand -ResourceGroupName $vm.ResourceGroupName -VMName $vm.Name `
                            -CommandId 'RunShellScript' -ScriptString $command
                        
                        $osInfo = $result.Value[0].Message -split "`n"
                        
                        # Parse Linux os-release output
                        $osName = ($osInfo | Where-Object { $_ -like "NAME=*" }) -replace 'NAME="|"', ""
                        $osVersion = ($osInfo | Where-Object { $_ -like "VERSION=*" }) -replace 'VERSION="|"', ""
                        $osBuild = ($osInfo | Where-Object { $_ -like "VERSION_ID=*" }) -replace 'VERSION_ID="|"', ""
                        $systemType = "Linux"
                    }
                    
                    # Add results
                    $results += [PSCustomObject]@{
                        'Subscription'    = $sub.Name
                        'SubscriptionId'  = $sub.Id
                        'ResourceGroup'   = $vm.ResourceGroupName
                        'VM_Name'        = $vm.Name
                        'VM_Size'        = $vm.HardwareProfile.VmSize
                        'VM_Status'      = $vmStatus
                        'OS_Name'        = $osName
                        'OS_Version'     = $osVersion
                        'OS_Build'       = $osBuild
                        'System_Type'    = $systemType
                    }
                }
                catch {
                    Write-Warning "Error processing VM $($vm.Name): $_"
                    
                    # Add error VM to results
                    $results += [PSCustomObject]@{
                        'Subscription'    = $sub.Name
                        'SubscriptionId'  = $sub.Id
                        'ResourceGroup'   = $vm.ResourceGroupName
                        'VM_Name'        = $vm.Name
                        'VM_Size'        = $vm.HardwareProfile.VmSize
                        'VM_Status'      = $vmStatus
                        'OS_Name'        = "Error: $_"
                        'OS_Version'     = "Error"
                        'OS_Build'       = "Error"
                        'System_Type'    = "Error"
                    }
                    continue
                }
            }
        }
        
        if ($results.Count -eq 0) {
            Write-Warning "No VM data found in any subscription"
            return
        }
        
        # Create output directory if it doesn't exist
        $outputDirectory = Split-Path -Parent $OutputFile
        if (![string]::IsNullOrWhiteSpace($outputDirectory) -and !(Test-Path -Path $outputDirectory)) {
            New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
        }
        
        # Export to Excel
        Write-Host "`nExporting $($results.Count) records to Excel: $OutputFile"
        $results | Export-Excel -Path $OutputFile -WorksheetName 'VM_SystemInfo' -AutoSize -AutoFilter -FreezeTopRow
        Write-Host "Excel report generated: $OutputFile"
        
        return $results
    }
    catch {
        Write-Error "An error occurred: $_"
        throw $_
    }
}

# Interactive mode
$outputFile = Read-Host "Enter output file path (press Enter for default './AzureVMSystemInfo.xlsx')"

if ([string]::IsNullOrWhiteSpace($outputFile)) {
    $outputFile = ".\AzureVMSystemInfo.xlsx"
}

# Call the function
Get-VMSystemInfo -OutputFile $outputFile 