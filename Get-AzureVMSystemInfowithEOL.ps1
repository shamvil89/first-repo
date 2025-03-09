# Check if required modules are installed
$requiredModules = @('Az.Accounts', 'Az.Compute', 'ImportExcel')
foreach ($module in $requiredModules) {
    if (!(Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing $module module..."
        Install-Module -Name $module -Force -Scope CurrentUser
    }
    Import-Module $module -Force
}

function Get-OSLifecycleStatus {
    param (
        [string]$OSName,
        [string]$OSVersion,
        [string]$SystemType
    )
    
    # Define EOL/EOS dates for common operating systems
    $osLifecycleDates = @{
        # Windows Server
        "Windows Server 2008" = @{EOL = "2020-01-14"; EOS = "2020-01-14"}
        "Windows Server 2008 R2" = @{EOL = "2020-01-14"; EOS = "2020-01-14"}
        "Windows Server 2012" = @{EOL = "2023-10-10"; EOS = "2023-10-10"}
        "Windows Server 2012 R2" = @{EOL = "2023-10-10"; EOS = "2023-10-10"}
        "Windows Server 2016" = @{EOL = "2027-01-12"; EOS = "2022-01-11"}
        "Windows Server 2019" = @{EOL = "2029-01-09"; EOS = "2024-01-09"}
        "Windows Server 2022" = @{EOL = "2031-10-14"; EOS = "2026-10-13"}
        
        # Windows Client
        "Windows 7" = @{EOL = "2020-01-14"; EOS = "2020-01-14"}
        "Windows 8" = @{EOL = "2016-01-12"; EOS = "2016-01-12"}
        "Windows 8.1" = @{EOL = "2023-01-10"; EOS = "2023-01-10"}
        "Windows 10" = @{EOL = "2025-10-14"; EOS = "2025-10-14"}
        "Windows 11" = @{EOL = "2027-10-14"; EOS = "2027-10-14"}
        
        # Linux - Ubuntu LTS
        "Ubuntu 16.04" = @{EOL = "2021-04-30"; EOS = "2021-04-30"}
        "Ubuntu 18.04" = @{EOL = "2023-04-30"; EOS = "2023-04-30"}
        "Ubuntu 20.04" = @{EOL = "2025-04-30"; EOS = "2025-04-30"}
        "Ubuntu 22.04" = @{EOL = "2027-04-30"; EOS = "2027-04-30"}
        
        # RHEL
        "Red Hat Enterprise Linux 7" = @{EOL = "2024-06-30"; EOS = "2024-06-30"}
        "Red Hat Enterprise Linux 8" = @{EOL = "2029-05-31"; EOS = "2029-05-31"}
        "Red Hat Enterprise Linux 9" = @{EOL = "2032-05-31"; EOS = "2032-05-31"}
        
        # CentOS
        "CentOS Linux 7" = @{EOL = "2024-06-30"; EOS = "2024-06-30"}
        "CentOS Linux 8" = @{EOL = "2021-12-31"; EOS = "2021-12-31"}
        
        # SUSE
        "SUSE Linux Enterprise Server 12" = @{EOL = "2027-10-31"; EOS = "2027-10-31"}
        "SUSE Linux Enterprise Server 15" = @{EOL = "2031-07-31"; EOS = "2031-07-31"}
    }
    
    # Initialize result
    $result = @{
        EOLDate = "Unknown"
        EOSDate = "Unknown"
        IsEOL = $false
        IsEOS = $false
        Status = "Unknown"
    }
    
    # Find matching OS
    $matchedOS = $null
    
    foreach ($os in $osLifecycleDates.Keys) {
        if ($OSName -like "*$os*") {
            $matchedOS = $os
            break
        }
    }
    
    # Special handling for Windows 10/11 versions
    if ($OSName -like "*Windows 10*" -or $OSName -like "*Windows 11*") {
        # Extract build number from version
        if ($OSVersion -match "\d+\.\d+\.(\d+)\.") {
            $buildNumber = $matches[1]
            
            # Windows 10 builds and EOL dates
            $win10Builds = @{
                "10240" = @{EOL = "2017-05-09"; EOS = "2017-05-09"; Version = "1507"}  # Initial Release
                "10586" = @{EOL = "2017-10-10"; EOS = "2017-10-10"; Version = "1511"}  # November Update
                "14393" = @{EOL = "2018-04-10"; EOS = "2018-04-10"; Version = "1607"}  # Anniversary Update
                "15063" = @{EOL = "2018-10-09"; EOS = "2018-10-09"; Version = "1703"}  # Creators Update
                "16299" = @{EOL = "2019-04-09"; EOS = "2019-04-09"; Version = "1709"}  # Fall Creators Update
                "17134" = @{EOL = "2019-11-12"; EOS = "2019-11-12"; Version = "1803"}  # April 2018 Update
                "17763" = @{EOL = "2020-11-10"; EOS = "2020-11-10"; Version = "1809"}  # October 2018 Update
                "18362" = @{EOL = "2020-11-10"; EOS = "2020-11-10"; Version = "1903"}  # May 2019 Update
                "18363" = @{EOL = "2021-05-11"; EOS = "2021-05-11"; Version = "1909"}  # November 2019 Update
                "19041" = @{EOL = "2021-12-14"; EOS = "2021-12-14"; Version = "2004"}  # May 2020 Update
                "19042" = @{EOL = "2022-05-10"; EOS = "2022-05-10"; Version = "20H2"}  # October 2020 Update
                "19043" = @{EOL = "2022-12-13"; EOS = "2022-12-13"; Version = "21H1"}  # May 2021 Update
                "19044" = @{EOL = "2023-06-13"; EOS = "2023-06-13"; Version = "21H2"}  # November 2021 Update
                "19045" = @{EOL = "2025-10-14"; EOS = "2025-10-14"; Version = "22H2"}  # October 2022 Update
            }
            
            # Windows 11 builds and EOL dates
            $win11Builds = @{
                "22000" = @{EOL = "2023-10-10"; EOS = "2023-10-10"; Version = "21H2"}  # Initial Release
                "22621" = @{EOL = "2024-10-08"; EOS = "2024-10-08"; Version = "22H2"}  # 2022 Update
                "22631" = @{EOL = "2025-10-14"; EOS = "2025-10-14"; Version = "23H2"}  # 2023 Update
            }
            
            if ($OSName -like "*Windows 10*" -and $win10Builds.ContainsKey($buildNumber)) {
                $buildInfo = $win10Builds[$buildNumber]
                $result.EOLDate = $buildInfo.EOL
                $result.EOSDate = $buildInfo.EOS
                $matchedOS = "Windows 10 version $($buildInfo.Version) (build $buildNumber)"
            }
            elseif ($OSName -like "*Windows 11*" -and $win11Builds.ContainsKey($buildNumber)) {
                $buildInfo = $win11Builds[$buildNumber]
                $result.EOLDate = $buildInfo.EOL
                $result.EOSDate = $buildInfo.EOS
                $matchedOS = "Windows 11 version $($buildInfo.Version) (build $buildNumber)"
            }
        }
    }
    
    # If we found a match in the general list and haven't set dates from build numbers
    if ($matchedOS -and $result.EOLDate -eq "Unknown") {
        $result.EOLDate = $osLifecycleDates[$matchedOS].EOL
        $result.EOSDate = $osLifecycleDates[$matchedOS].EOS
    }
    
    # Check if OS is EOL or EOS
    $today = Get-Date
    
    if ($result.EOLDate -ne "Unknown") {
        $eolDate = [DateTime]::ParseExact($result.EOLDate, "yyyy-MM-dd", $null)
        $result.IsEOL = ($today -gt $eolDate)
    }
    
    if ($result.EOSDate -ne "Unknown") {
        $eosDate = [DateTime]::ParseExact($result.EOSDate, "yyyy-MM-dd", $null)
        $result.IsEOS = ($today -gt $eosDate)
    }
    
    # Set status
    if ($result.IsEOL) {
        $result.Status = "End of Life"
    }
    elseif ($result.IsEOS) {
        $result.Status = "End of Support"
    }
    else {
        $result.Status = "Supported"
    }
    
    return $result
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
                        
                        # Try to get OS info from VM properties
                        $osType = $vm.StorageProfile.OsDisk.OsType
                        $osName = "Unknown"
                        $osVersion = "Unknown"
                        $osBuild = "Unknown"
                        $systemType = $osType.ToString()
                        
                        # Try to get OS info from image reference
                        if ($vm.StorageProfile.ImageReference) {
                            $imageRef = $vm.StorageProfile.ImageReference
                            
                            if ($imageRef.Publisher -and $imageRef.Offer -and $imageRef.Sku) {
                                if ($osType -eq "Windows") {
                                    # Windows OS
                                    if ($imageRef.Offer -like "WindowsServer*") {
                                        $osName = "Windows Server"
                                        $osVersion = $imageRef.Sku
                                        if ($osVersion -match "(\d+)") {
                                            $osName = "Windows Server $($matches[1])"
                                        }
                                    }
                                    elseif ($imageRef.Offer -like "Windows-*") {
                                        $osName = "Windows"
                                        if ($imageRef.Sku -like "*-win10-*") {
                                            $osName = "Windows 10"
                                        }
                                        elseif ($imageRef.Sku -like "*-win11-*") {
                                            $osName = "Windows 11"
                                        }
                                        $osVersion = $imageRef.Sku
                                    }
                                }
                                else {
                                    # Linux OS
                                    $osName = $imageRef.Offer
                                    $osVersion = $imageRef.Sku
                                    
                                    # Clean up common Linux distro names
                                    if ($imageRef.Publisher -eq "Canonical") {
                                        $osName = "Ubuntu"
                                        if ($osVersion -match "(\d+)_(\d+)") {
                                            $osVersion = "$($matches[1]).$($matches[2])"
                                        }
                                    }
                                    elseif ($imageRef.Publisher -eq "RedHat") {
                                        $osName = "Red Hat Enterprise Linux"
                                        if ($osVersion -match "(\d+)") {
                                            $osVersion = $matches[1]
                                        }
                                    }
                                    elseif ($imageRef.Publisher -eq "OpenLogic") {
                                        $osName = "CentOS"
                                        if ($osVersion -match "(\d+)") {
                                            $osVersion = $matches[1]
                                        }
                                    }
                                    elseif ($imageRef.Publisher -eq "SUSE") {
                                        $osName = "SUSE Linux Enterprise Server"
                                        if ($osVersion -match "(\d+)") {
                                            $osVersion = $matches[1]
                                        }
                                    }
                                }
                            }
                        }
                        
                        # Get OS lifecycle status based on extracted info
                        $lifecycleStatus = Get-OSLifecycleStatus -OSName $osName -OSVersion $osVersion -SystemType $systemType
                        
                        # Add non-running VM to results with extracted OS info
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
                            'EOL_Date'       = $lifecycleStatus.EOLDate
                            'EOS_Date'       = $lifecycleStatus.EOSDate
                            'Support_Status' = $lifecycleStatus.Status
                            'OS_Source'      = "VM Image Reference (VM not running)"
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
                    
                    # Get OS lifecycle status
                    $lifecycleStatus = Get-OSLifecycleStatus -OSName $osName -OSVersion $osVersion -SystemType $systemType
                    
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
                        'EOL_Date'       = $lifecycleStatus.EOLDate
                        'EOS_Date'       = $lifecycleStatus.EOSDate
                        'Support_Status' = $lifecycleStatus.Status
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
                        'EOL_Date'       = "Unknown"
                        'EOS_Date'       = "Unknown"
                        'Support_Status' = "Unknown"
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
        
        # Export to Excel with conditional formatting
        Write-Host "`nExporting $($results.Count) records to Excel: $OutputFile"
        
        $excel = $results | Export-Excel -Path $OutputFile -WorksheetName 'VM_SystemInfo' -AutoSize -AutoFilter -FreezeTopRow -PassThru
        
        # Add conditional formatting for support status
        $worksheet = $excel.Workbook.Worksheets['VM_SystemInfo']
        $lastRow = $worksheet.Dimension.End.Row
        
        # Find the Support_Status column
        $supportStatusColumn = $null
        for ($i = 1; $i -le $worksheet.Dimension.End.Column; $i++) {
            if ($worksheet.Cells[1, $i].Value -eq "Support_Status") {
                $supportStatusColumn = $i
                break
            }
        }
        
        if ($supportStatusColumn) {
            $colLetter = [char](64 + $supportStatusColumn)
            
            # Format EOL cells
            Add-ConditionalFormatting -Worksheet $worksheet -Range "$colLetter`2:$colLetter`$lastRow" -RuleType Equal -ConditionValue "End of Life" -BackgroundColor Red -ForegroundColor White
            
            # Format EOS cells
            Add-ConditionalFormatting -Worksheet $worksheet -Range "$colLetter`2:$colLetter`$lastRow" -RuleType Equal -ConditionValue "End of Support" -BackgroundColor Orange
            
            # Format Supported cells
            Add-ConditionalFormatting -Worksheet $worksheet -Range "$colLetter`2:$colLetter`$lastRow" -RuleType Equal -ConditionValue "Supported" -BackgroundColor Green -ForegroundColor White
        }
        
        # Save and close the Excel package
        Close-ExcelPackage $excel
        
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