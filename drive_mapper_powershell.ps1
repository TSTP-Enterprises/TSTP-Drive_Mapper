# PowerShell Script to Manage Network Drive Mappings
# Enhanced Features:
# - Import and export drive mappings from/to config.json
# - Import drive mappings from an external JSON file
# - Map, unmap, and force reconnect drives using CMD 'net use' commands
# - Manage startup settings
# - Implement threading to prevent GUI freezing
# - Comprehensive error handling and logging
# Generated on: 2024-11-14

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global Variables
$Global:ScriptPath = $MyInvocation.MyCommand.Path

# Function to add or remove the script from Windows startup
function Set-Startup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [bool]$Enable
    )

    if (-not $Global:ScriptPath) {
        [System.Windows.Forms.MessageBox]::Show("Cannot determine the script path. Please run the script directly.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $shell = New-Object -ComObject WScript.Shell
    $shortcutPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\DriveMapper.lnk"

    if ($Enable) {
        if (-not (Test-Path $shortcutPath)) {
            # Create a shortcut in the Startup folder
            $shortcut = $shell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = "powershell.exe"
            $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$Global:ScriptPath`""
            $shortcut.WorkingDirectory = Split-Path $Global:ScriptPath
            $shortcut.IconLocation = "shell32.dll,3"  # Standard PowerShell icon
            $shortcut.Save()
        }
    } else {
        # Remove the shortcut if it exists
        if (Test-Path $shortcutPath) {
            Remove-Item $shortcutPath -Force
        }
    }
}

# Function to export configuration to a JSON file
function Export-Configuration {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath,

        [Parameter(Mandatory = $true)]
        [object]$ConfigData
    )
    $ConfigData | ConvertTo-Json -Depth 5 | Set-Content -Path $ConfigPath -Force
}

# Function to import configuration from a JSON file
function Import-Configuration {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )
    if (Test-Path $ConfigPath) {
        try {
            return Get-Content -Path $ConfigPath | ConvertFrom-Json
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to parse config.json. Ensure it's properly formatted.", "Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return $null
        }
    } else {
        return @{
            Drives = @()
            ReAddAtStart = $false
            StartOnWindowsStart = $false
        }
    }
}

# Function to import external JSON configuration
function Import-External-Configuration {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExternalPath
    )
    if (Test-Path $ExternalPath) {
        try {
            return Get-Content -Path $ExternalPath | ConvertFrom-Json
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to import JSON file. Ensure it's properly formatted.", "Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return $null
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Selected file does not exist.", "Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $null
    }
}

# Function to map a single drive using CMD 'net use' command
function New-DriveMapping {  # Renamed from 'Map-Drive' to 'New-DriveMapping'
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$DriveConfig
    )

    $driveLetter = $DriveConfig.Drive
    $uncPath = $DriveConfig.UNCPath
    $useCred = $DriveConfig.UseCredentials
    $username = $DriveConfig.Username
    $password = $DriveConfig.Password  # Renamed from 'pwd' to 'password'

    Write-Output "Processing drive ${driveLetter} -> $uncPath"

    # Prepare the 'net use' command
    if ($useCred -and $username -and $password) {
        $cmd = "net use ${driveLetter} `"$uncPath`" `"$password`" /user:`"$username`" /persistent:yes"
    } else {
        $cmd = "net use ${driveLetter} `"$uncPath`" /persistent:yes"
    }

    try {
        # Execute the 'net use' command
        $result = cmd.exe /c $cmd 2>&1
        if ($result -match "The command completed successfully") {
            Write-Output "Successfully mapped drive ${driveLetter} to $uncPath."
            return $true
        } else {
            Write-Error "Failed to map drive ${driveLetter} to $uncPath. Result: $result"
            return $false
        }
    } catch {
        Write-Error "Exception occurred while mapping drive ${driveLetter}: $_"
        return $false
    }
}

# Function to unmap a single drive using CMD 'net use' command
function Remove-DriveMapping {  # Renamed from 'Unmap-Drive' to 'Remove-DriveMapping'
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$DriveConfig
    )

    $driveLetter = $DriveConfig.Drive

    Write-Output "Unmapping drive ${driveLetter}"

    # Prepare the 'net use' command
    $cmd = "net use ${driveLetter} /delete /yes"

    try {
        # Execute the 'net use' command
        $result = cmd.exe /c $cmd 2>&1
        if ($result -match "The command completed successfully") {
            Write-Output "Successfully unmapped drive ${driveLetter}."
            return $true
        } else {
            Write-Error "Failed to unmap drive ${driveLetter}. Result: $result"
            return $false
        }
    } catch {
        Write-Error "Exception occurred while unmapping drive ${driveLetter}: $_"
        return $false
    }
}

# Function to perform drive mapping operations asynchronously using Runspaces
function Start-DriveOperations {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Operation  # "Map" or "Unmap"
    )

    # Create a RunspacePool
    $runspacePool = [RunspaceFactory]::CreateRunspacePool(1, [Environment]::ProcessorCount)
    $runspacePool.Open()

    $jobs = @()

    foreach ($drive in $drives) {
        # Skip drives that are not selected
        if (-not $drive.Selected) {
            continue
        }

        # Create a PowerShell instance
        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $runspacePool

        if ($Operation -eq "Map") {
            $ps.AddScript({
                param($d)
                New-DriveMapping -DriveConfig $d
            }).AddArgument($drive)
        }
        elseif ($Operation -eq "Unmap") {
            $ps.AddScript({
                param($d)
                Remove-DriveMapping -DriveConfig $d
            }).AddArgument($drive)
        }
        else {
            Write-Error "Invalid operation: $Operation"
            continue
        }

        # Invoke asynchronously
        $job = $ps.BeginInvoke()
        $jobs += @{ PowerShell = $ps; AsyncResult = $job }
    }

    # Collect results
    foreach ($job in $jobs) {
        $job.PowerShell.EndInvoke($job.AsyncResult)
        $job.PowerShell.Dispose()
    }

    # Close the RunspacePool
    $runspacePool.Close()
    $runspacePool.Dispose()
}

# Function to create and display the Drive Mapping form
function Show-DriveMappingForm {
    [CmdletBinding()]
    param ()

    # Define the path to config.json in the same directory as the script
    $scriptDirectory = (Get-Item -Path $PSCommandPath).DirectoryName
    if ($null -eq $scriptDirectory) {
        Write-Error "Failed to determine script directory."
        return
    }
    $configPath = Join-Path $scriptDirectory "config.json"
    if (-not (Test-Path $configPath)) {
        Write-Warning "Config file not found at $configPath. Creating a new one."
        $defaultConfig = @{
            Drives = @()
            ReAddAtStart = $false
            StartOnWindowsStart = $false
        }
        $defaultConfig | ConvertTo-Json | Set-Content -Path $configPath
    }
    $config = Import-Configuration -ConfigPath $configPath

    # Initialize Form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Drive Mapping Configuration"
    $form.Size = New-Object System.Drawing.Size(900, 800)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false

    # Instructions Label
    $labelInstructions = New-Object System.Windows.Forms.Label
    $labelInstructions.Text = "Configure your drive mappings below:"
    $labelInstructions.Location = New-Object System.Drawing.Point(10, 10)
    $labelInstructions.AutoSize = $true
    $form.Controls.Add($labelInstructions)

    # DataGridView for Drive Mappings
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Location = New-Object System.Drawing.Point(10, 40)
    $dataGridView.Size = New-Object System.Drawing.Size(860, 400)
    $dataGridView.AutoGenerateColumns = $false
    $dataGridView.AllowUserToAddRows = $true
    $dataGridView.AllowUserToDeleteRows = $true
    $dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $dataGridView.MultiSelect = $false
    $dataGridView.ReadOnly = $false
    $dataGridView.RowHeadersVisible = $false
    $form.Controls.Add($dataGridView)

    # Define Columns
    # Drive Column
    $colDrive = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colDrive.HeaderText = "Drive"
    $colDrive.Name = "Drive"
    $colDrive.Width = 60
    $dataGridView.Columns.Add($colDrive)

    # UNC Path Column
    $colUNCPath = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colUNCPath.HeaderText = "UNC Path"
    $colUNCPath.Name = "UNCPath"
    $colUNCPath.Width = 300
    $dataGridView.Columns.Add($colUNCPath)

    # Added Date Column
    $colAddedDate = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAddedDate.HeaderText = "Added Date"
    $colAddedDate.Name = "AddedDate"
    $colAddedDate.Width = 150
    $dataGridView.Columns.Add($colAddedDate)

    # Mapped Column
    $colMapped = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMapped.HeaderText = "Mapped"
    $colMapped.Name = "Mapped"
    $colMapped.Width = 60
    $colMapped.ReadOnly = $true
    $dataGridView.Columns.Add($colMapped)

    # Selected Column
    $colSelected = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colSelected.HeaderText = "Selected"
    $colSelected.Name = "Selected"
    $colSelected.Width = 60
    $dataGridView.Columns.Add($colSelected)

    # Force Connect Button Column
    $colForceConnect = New-Object System.Windows.Forms.DataGridViewButtonColumn
    $colForceConnect.HeaderText = "Force Connect"
    $colForceConnect.Name = "ForceConnect"
    $colForceConnect.Text = "Force Connect"
    $colForceConnect.UseColumnTextForButtonValue = $true
    $colForceConnect.Width = 120
    $dataGridView.Columns.Add($colForceConnect)

    # Populate DataGridView with existing mappings
    if ($config.Drives -and $config.Drives.Count -gt 0) {
        foreach ($mapping in $config.Drives) {
            $rowIndex = $dataGridView.Rows.Add()
            $dataGridView.Rows[$rowIndex].Cells["Drive"].Value = $mapping.Drive
            $dataGridView.Rows[$rowIndex].Cells["UNCPath"].Value = $mapping.UNCPath
            $dataGridView.Rows[$rowIndex].Cells["AddedDate"].Value = $mapping.AddedDate
            $dataGridView.Rows[$rowIndex].Cells["Mapped"].Value = $mapping.Mapped
            $dataGridView.Rows[$rowIndex].Cells["Selected"].Value = $mapping.Selected
        }
    }

    # Event Handler for Force Connect Button Click
    $dataGridView.add_CellContentClick({
        param($eventSender, $e)
        if ($e.ColumnIndex -eq ($dataGridView.Columns["ForceConnect"].Index) -and $e.RowIndex -ge 0) {
            $row = $dataGridView.Rows[$e.RowIndex]
            $driveLetter = $row.Cells["Drive"].Value
            $uncPath = $row.Cells["UNCPath"].Value
            $useCredentials = $config.Drives[$e.RowIndex].UseCredentials
            $username = $config.Drives[$e.RowIndex].Username
            $password = $config.Drives[$e.RowIndex].Password
    
            if ([string]::IsNullOrWhiteSpace($driveLetter) -or [string]::IsNullOrWhiteSpace($uncPath)) {
                [System.Windows.Forms.MessageBox]::Show("Drive Letter and UNC Path cannot be empty.", "Invalid Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
    
            # Prompt for credentials if required
            if ($useCredentials -and (-not $username -or -not $password)) {
                $credential = Get-Credential -Message "Enter credentials for mapping drive ${driveLetter}:"
                $username = $credential.UserName
                $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($credential.Password)
                )
            }
    
            # Start mapping/unmapping in a separate job
            $job = Start-Job -ScriptBlock {
                param($d, $u, $useCreds, $user, $pass)
                try {
                    # Check if the drive is already mapped
                    $checkResult = cmd.exe /c "net use ${d}" 2>&1
                    if ($checkResult -notmatch "The network connection could not be found") {
                        # Unmap the existing drive if it is already mapped
                        cmd.exe /c "net use ${d} /delete /yes" | Out-Null
                    }
    
                    # Prepare and execute the 'net use' command
                    if ($useCreds) {
                        $cmd = "net use ${d} `"$u`" `"$pass`" /user:`"$user`" /persistent:yes"
                    } else {
                        $cmd = "net use ${d} `"$u`" /persistent:yes"
                    }
    
                    # Execute the command
                    $result = cmd.exe /c $cmd 2>&1
                    if ($result -match "The command completed successfully") {
                        Write-Output "Success"
                    } else {
                        Write-Error "Failed: $result"
                    }
                } catch {
                    Write-Error "Error while processing drive ${d}: $($_.Exception.Message)"
                }
            } -ArgumentList ($driveLetter, $uncPath, $useCredentials, $username, $password)
    
            # Wait for job completion and update UI
            Wait-Job -Job $job
            $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
            Remove-Job -Job $job
    
            if ($result -eq "Success") {
                # Update UI for success
                $row.Cells["Mapped"].Value = "Yes"
                $row.Cells["AddedDate"].Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                [System.Windows.Forms.MessageBox]::Show("Drive ${driveLetter} has been successfully connected.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                # Update UI for failure
                $row.Cells["Mapped"].Value = "No"
                [System.Windows.Forms.MessageBox]::Show("Failed to connect drive ${driveLetter}: $result", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    

    # Panel to hold action buttons
    $panelActions = New-Object System.Windows.Forms.Panel
    $panelActions.Size = New-Object System.Drawing.Size(860, 60)
    $panelActions.Location = New-Object System.Drawing.Point(10, 460)
    $panelActions.Anchor = "Top, Left, Right"
    $form.Controls.Add($panelActions)

    # Create Buttons
    $buttonMap = New-Object System.Windows.Forms.Button
    $buttonMap.Text = "Map Drives"
    $buttonMap.Size = New-Object System.Drawing.Size(120, 40)

    $buttonUnmap = New-Object System.Windows.Forms.Button
    $buttonUnmap.Text = "Unmap Drives"
    $buttonUnmap.Size = New-Object System.Drawing.Size(120, 40)

    $buttonImport = New-Object System.Windows.Forms.Button
    $buttonImport.Text = "Import JSON"
    $buttonImport.Size = New-Object System.Drawing.Size(120, 40)

    $buttonExport = New-Object System.Windows.Forms.Button
    $buttonExport.Text = "Export JSON"
    $buttonExport.Size = New-Object System.Drawing.Size(120, 40)

    $buttonClose = New-Object System.Windows.Forms.Button
    $buttonClose.Text = "Close"
    $buttonClose.Size = New-Object System.Drawing.Size(120, 40)

    # Add Buttons to Panel
    $panelActions.Controls.Add($buttonMap)
    $panelActions.Controls.Add($buttonUnmap)
    $panelActions.Controls.Add($buttonImport)
    $panelActions.Controls.Add($buttonExport)
    $panelActions.Controls.Add($buttonClose)

    # Arrange buttons in a single row with equal spacing using FlowLayoutPanel
    $flowLayout = New-Object System.Windows.Forms.FlowLayoutPanel
    $flowLayout.Dock = "Fill"
    $flowLayout.FlowDirection = "LeftToRight"
    $flowLayout.WrapContents = $false
    $flowLayout.AutoSize = $false
    $flowLayout.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)
    $flowLayout.Controls.AddRange(@($buttonMap, $buttonUnmap, $buttonImport, $buttonExport, $buttonClose))
    $panelActions.Controls.Clear()
    $panelActions.Controls.Add($flowLayout)

    # Event Handlers for Map, Unmap, Import, Export Buttons
    $buttonMap.Add_Click({
        # Disable buttons to prevent multiple clicks
        $buttonMap.Enabled = $false
        $buttonUnmap.Enabled = $false
        $buttonImport.Enabled = $false
        $buttonExport.Enabled = $false

        # Gather selected drives
        $selectedDrives = @()
        foreach ($row in $dataGridView.Rows) {
            if (-not $row.IsNewRow -and $row.Cells["Selected"].Value -eq $true) {
                $index = $dataGridView.Rows.IndexOf($row)
                $selectedDrives += @{
                    Drive = $row.Cells["Drive"].Value
                    UNCPath = $row.Cells["UNCPath"].Value
                    AddedDate = $row.Cells["AddedDate"].Value
                    Mapped = $row.Cells["Mapped"].Value
                    Selected = $row.Cells["Selected"].Value
                    UseCredentials = $config.Drives[$index].UseCredentials
                    Username = $config.Drives[$index].Username
                    Password = $config.Drives[$index].Password
                }
            }
        }

        if ($selectedDrives.Count -eq 0) {
            $result = [System.Windows.Forms.MessageBox]::Show("No drives selected. Do you want to map all drives?", "No Selection", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                foreach ($row in $dataGridView.Rows) {
                    if (-not $row.IsNewRow) {
                        $index = $dataGridView.Rows.IndexOf($row)
                        $selectedDrives += @{
                            Drive = $row.Cells["Drive"].Value
                            UNCPath = $row.Cells["UNCPath"].Value
                            AddedDate = $row.Cells["AddedDate"].Value
                            Mapped = $row.Cells["Mapped"].Value
                            Selected = $row.Cells["Selected"].Value
                            UseCredentials = $config.Drives[$index].UseCredentials
                            Username = $config.Drives[$index].Username
                            Password = $config.Drives[$index].Password
                        }
                    }
                }
            } else {
                # Re-enable buttons
                $buttonMap.Enabled = $true
                $buttonUnmap.Enabled = $true
                $buttonImport.Enabled = $true
                $buttonExport.Enabled = $true
                return
            }
        }

        # Start mapping in separate jobs to prevent GUI freezing
        foreach ($drive in $selectedDrives) {
            Start-Job -ScriptBlock {
                param($d, $u, $useCredentials, $username, $password)
                if ($d -and $u) {
                    # Unmap existing drive if mapped
                    cmd.exe /c "net use $d /delete /yes" | Out-Null

                    # Prepare 'net use' command
                    if ($useCredentials) {
                        $cmd = "net use $d `"$u`" `"$password`" /user:`"$username`" /persistent:yes"
                    } else {
                        $cmd = "net use $d `"$u`" /persistent:yes"
                    }

                    # Execute the command
                    $result = cmd.exe /c $cmd 2>&1
                    if ($result -match "The command completed successfully") {
                        Write-Output "Success"
                    } else {
                        Write-Error "Failed: $result"
                    }
                } else {
                    Write-Error "Drive Letter or UNC Path is missing."
                }
            } -ArgumentList ($drive.Drive, $drive.UNCPath, $drive.UseCredentials, $drive.Username, $drive.Password) | Out-Null
        }

        # Inform user that mapping has been initiated
        [System.Windows.Forms.MessageBox]::Show("Drive mapping operations have been initiated. Please wait for completion.", "Mapping Initiated", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        # Monitor and handle job results
        $jobs = Get-Job | Where-Object { $_.Command -like "*net use $" }

        foreach ($job in $jobs) {
            Wait-Job -Job $job
            $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
            Remove-Job -Job $job

            if ($result -eq "Success") {
                # Update UI with success status
                foreach ($drive in $selectedDrives) {
                    for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                        $row = $dataGridView.Rows[$i]
                        if ($row.Cells["Drive"].Value -eq $drive.Drive -and $row.Cells["UNCPath"].Value -eq $drive.UNCPath) {
                            $row.Cells["Mapped"].Value = "Yes"
                            $row.Cells["AddedDate"].Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                        }
                    }
                }
                [System.Windows.Forms.MessageBox]::Show("Drive mappings have been successfully completed.", "Mapping Completed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                # Handle failures
                foreach ($drive in $selectedDrives) {
                    for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                        $row = $dataGridView.Rows[$i]
                        if ($row.Cells["Drive"].Value -eq $drive.Drive -and $row.Cells["UNCPath"].Value -eq $drive.UNCPath) {
                            $row.Cells["Mapped"].Value = "No"
                            [System.Windows.Forms.MessageBox]::Show("Failed to map drive ${drive.Drive}: $result", "Mapping Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        }
                    }
                }
            }
        }

        # Save current mappings to configuration
        $mappedDrives = @()
        foreach ($row in $dataGridView.Rows) {
            if (-not $row.IsNewRow) {
                $index = $dataGridView.Rows.IndexOf($row)
                $mappedDrives += @{
                    Drive = $row.Cells["Drive"].Value
                    UNCPath = $row.Cells["UNCPath"].Value
                    AddedDate = $row.Cells["AddedDate"].Value
                    Mapped = $row.Cells["Mapped"].Value
                    Selected = $row.Cells["Selected"].Value
                    UseCredentials = $config.Drives[$index].UseCredentials
                    Username = $config.Drives[$index].Username
                    Password = $config.Drives[$index].Password
                }
            }
        }
        $newConfig = @{
            Drives = $mappedDrives
            ReAddAtStart = $checkboxReAddAtStart.Checked
            StartOnWindowsStart = $checkboxStartOnWindowsStart.Checked
        }
        Export-Configuration -ConfigPath $configPath -ConfigData $newConfig

        # Set startup based on checkbox
        Set-Startup -Enable $checkboxStartOnWindowsStart.Checked

        # Re-enable buttons
        $buttonMap.Enabled = $true
        $buttonUnmap.Enabled = $true
        $buttonImport.Enabled = $true
        $buttonExport.Enabled = $true
    })

    $buttonUnmap.Add_Click({
        # Disable buttons to prevent multiple clicks
        $buttonMap.Enabled = $false
        $buttonUnmap.Enabled = $false
        $buttonImport.Enabled = $false
        $buttonExport.Enabled = $false

        # Gather selected drives
        $selectedDrives = @()
        foreach ($row in $dataGridView.Rows) {
            if (-not $row.IsNewRow -and $row.Cells["Selected"].Value -eq $true) {
                $index = $dataGridView.Rows.IndexOf($row)
                $selectedDrives += @{
                    Drive = $row.Cells["Drive"].Value
                    UNCPath = $row.Cells["UNCPath"].Value
                    AddedDate = $row.Cells["AddedDate"].Value
                    Mapped = $row.Cells["Mapped"].Value
                    Selected = $row.Cells["Selected"].Value
                    UseCredentials = $config.Drives[$index].UseCredentials
                    Username = $config.Drives[$index].Username
                    Password = $config.Drives[$index].Password
                }
            }
        }

        if ($selectedDrives.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one drive to unmap.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            # Re-enable buttons
            $buttonMap.Enabled = $true
            $buttonUnmap.Enabled = $true
            $buttonImport.Enabled = $true
            $buttonExport.Enabled = $true
            return
        }

        # Start unmapping in separate jobs to prevent GUI freezing
        foreach ($drive in $selectedDrives) {
            Start-Job -ScriptBlock {
                param($d)
                $cmd = "net use $d /delete /yes"
                $result = cmd.exe /c $cmd 2>&1
                if ($result -match "The command completed successfully") {
                    Write-Output "Success"
                } else {
                    Write-Error "Failed: $result"
                }
            } -ArgumentList ($drive.Drive) | Out-Null
        }

        # Inform user that unmapping has been initiated
        [System.Windows.Forms.MessageBox]::Show("Drive unmapping operations have been initiated. Please wait for completion.", "Unmapping Initiated", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        # Monitor and handle job results
        $jobs = Get-Job | Where-Object { $_.Command -like "*net use $" }

        foreach ($job in $jobs) {
            Wait-Job -Job $job
            $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
            Remove-Job -Job $job

            if ($result -eq "Success") {
                # Update UI with success status
                foreach ($drive in $selectedDrives) {
                    for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                        $row = $dataGridView.Rows[$i]
                        if ($row.Cells["Drive"].Value -eq $drive.Drive -and $row.Cells["UNCPath"].Value -eq $drive.UNCPath) {
                            $row.Cells["Mapped"].Value = "No"
                            $row.Cells["AddedDate"].Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                        }
                    }
                }
                [System.Windows.Forms.MessageBox]::Show("Drive unmapping operations have been successfully completed.", "Unmapping Completed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                # Handle failures
                foreach ($drive in $selectedDrives) {
                    for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                        $row = $dataGridView.Rows[$i]
                        if ($row.Cells["Drive"].Value -eq $drive.Drive -and $row.Cells["UNCPath"].Value -eq $drive.UNCPath) {
                            $row.Cells["Mapped"].Value = "No"
                            [System.Windows.Forms.MessageBox]::Show("Failed to unmap drive ${drive.Drive}: $result", "Unmapping Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        }
                    }
                }
            }
        }

        # Save current mappings to configuration
        $mappedDrives = @()
        foreach ($row in $dataGridView.Rows) {
            if (-not $row.IsNewRow) {
                $index = $dataGridView.Rows.IndexOf($row)
                $mappedDrives += @{
                    Drive = $row.Cells["Drive"].Value
                    UNCPath = $row.Cells["UNCPath"].Value
                    AddedDate = $row.Cells["AddedDate"].Value
                    Mapped = $row.Cells["Mapped"].Value
                    Selected = $row.Cells["Selected"].Value
                    UseCredentials = $config.Drives[$index].UseCredentials
                    Username = $config.Drives[$index].Username
                    Password = $config.Drives[$index].Password
                }
            }
        }
        $newConfig = @{
            Drives = $mappedDrives
            ReAddAtStart = $checkboxReAddAtStart.Checked
            StartOnWindowsStart = $checkboxStartOnWindowsStart.Checked
        }
        Export-Configuration -ConfigPath $configPath -ConfigData $newConfig

        # Set startup based on checkbox
        Set-Startup -Enable $checkboxStartOnWindowsStart.Checked

        # Re-enable buttons
        $buttonMap.Enabled = $true
        $buttonUnmap.Enabled = $true
        $buttonImport.Enabled = $true
        $buttonExport.Enabled = $true
    })

    $buttonImport.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.InitialDirectory = $scriptDirectory
        $openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        $openFileDialog.Title = "Import Drive Mappings JSON"

        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $externalPath = $openFileDialog.FileName
            $externalConfig = Import-External-Configuration -ExternalPath $externalPath
            if ($externalConfig -and $externalConfig.Drives) {
                # Clear existing rows
                $dataGridView.Rows.Clear()

                # Populate DataGridView with imported mappings
                foreach ($mapping in $externalConfig.Drives) {
                    $rowIndex = $dataGridView.Rows.Add()
                    $dataGridView.Rows[$rowIndex].Cells["Drive"].Value = $mapping.Drive
                    $dataGridView.Rows[$rowIndex].Cells["UNCPath"].Value = $mapping.UNCPath
                    $dataGridView.Rows[$rowIndex].Cells["AddedDate"].Value = $mapping.AddedDate
                    $dataGridView.Rows[$rowIndex].Cells["Mapped"].Value = $mapping.Mapped
                    $dataGridView.Rows[$rowIndex].Cells["Selected"].Value = $mapping.Selected
                }

                # Update other settings
                $checkboxReAddAtStart.Checked = $externalConfig.ReAddAtStart
                $checkboxStartOnWindowsStart.Checked = $externalConfig.StartOnWindowsStart

                # Save the imported configuration
                Export-Configuration -ConfigPath $configPath -ConfigData $externalConfig

                # Update the global config variable
                #$config = Import-Configuration -ConfigPath $configPath

                # Set startup based on checkbox
                Set-Startup -Enable $externalConfig.StartOnWindowsStart

                [System.Windows.Forms.MessageBox]::Show("JSON configuration imported successfully.", "Import Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                [System.Windows.Forms.MessageBox]::Show("Failed to import configuration. The file may be invalid or missing required data.", "Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

    $buttonExport.Add_Click({
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.InitialDirectory = $scriptDirectory
        $saveFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        $saveFileDialog.Title = "Export Drive Mappings to JSON"

        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $exportPath = $saveFileDialog.FileName

            # Gather all mappings
            $mappedDrives = @()
            for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                $currentRow = $dataGridView.Rows[$i]
                if (-not $currentRow.IsNewRow) {
                    $mappedDrives += @{
                        Drive = $currentRow.Cells["Drive"].Value
                        UNCPath = $currentRow.Cells["UNCPath"].Value
                        AddedDate = $currentRow.Cells["AddedDate"].Value
                        Mapped = $currentRow.Cells["Mapped"].Value
                        Selected = $currentRow.Cells["Selected"].Value
                        UseCredentials = $config.Drives[$i].UseCredentials
                        Username = $config.Drives[$i].Username
                        Password = $config.Drives[$i].Password
                    }
                }
            }

            $exportConfig = @{
                Drives = $mappedDrives
                ReAddAtStart = $checkboxReAddAtStart.Checked
                StartOnWindowsStart = $checkboxStartOnWindowsStart.Checked
            }

            try {
                Export-Configuration -ConfigPath $exportPath -ConfigData $exportConfig
                [System.Windows.Forms.MessageBox]::Show("Configuration exported successfully to $exportPath.", "Export Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to export configuration. Error: $($_.Exception.Message)", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

    $buttonClose.Add_Click({
        $form.Close()
    })

    # Panel for Checkboxes
    $panelCheckboxes = New-Object System.Windows.Forms.Panel
    $panelCheckboxes.Size = New-Object System.Drawing.Size(860, 50)
    $panelCheckboxes.Location = New-Object System.Drawing.Point(10, 530)
    $panelCheckboxes.Anchor = "Top, Left, Right"
    $form.Controls.Add($panelCheckboxes)

    # Re-add at Start Checkbox
    $checkboxReAddAtStart = New-Object System.Windows.Forms.CheckBox
    $checkboxReAddAtStart.Text = "Re-add Drives at Start"
    $checkboxReAddAtStart.AutoSize = $true
    $checkboxReAddAtStart.Checked = $config.ReAddAtStart
    $checkboxReAddAtStart.Location = New-Object System.Drawing.Point(0, 15)
    $checkboxReAddAtStart.Anchor = "Top"
    $checkboxReAddAtStart.Add_CheckedChanged({
        $config.ReAddAtStart = $checkboxReAddAtStart.Checked
        Export-Configuration -ConfigPath $configPath -ConfigData $config
    })
    $panelCheckboxes.Controls.Add($checkboxReAddAtStart)

    # Start on Windows Startup Checkbox
    $checkboxStartOnWindowsStart = New-Object System.Windows.Forms.CheckBox
    $checkboxStartOnWindowsStart.Text = "Start on Windows Startup"
    $checkboxStartOnWindowsStart.AutoSize = $true
    $checkboxStartOnWindowsStart.Checked = $config.StartOnWindowsStart
    $checkboxStartOnWindowsStart.Location = New-Object System.Drawing.Point(300, 15)  # Positioned to the right
    $checkboxStartOnWindowsStart.Anchor = "Top"
    $checkboxStartOnWindowsStart.Add_CheckedChanged({
        $config.StartOnWindowsStart = $checkboxStartOnWindowsStart.Checked
        Export-Configuration -ConfigPath $configPath -ConfigData $config
        Set-Startup -Enable $checkboxStartOnWindowsStart.Checked
    })
    $panelCheckboxes.Controls.Add($checkboxStartOnWindowsStart)

    # Arrange checkboxes in a single row with spacing using FlowLayoutPanel
    $flowCheckboxes = New-Object System.Windows.Forms.FlowLayoutPanel
    $flowCheckboxes.Dock = "Fill"
    $flowCheckboxes.FlowDirection = "LeftToRight"
    $flowCheckboxes.WrapContents = $false
    $flowCheckboxes.AutoSize = $false
    $flowCheckboxes.Padding = New-Object System.Windows.Forms.Padding(0)
    $flowCheckboxes.Controls.AddRange(@($checkboxReAddAtStart, $checkboxStartOnWindowsStart))
    $panelCheckboxes.Controls.Clear()
    $panelCheckboxes.Controls.Add($flowCheckboxes)

    # Function to re-add drives on script start if enabled
    function Invoke-ReAddDrivesAtStart {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$ConfigPath
        )
        $currentConfig = Import-Configuration -ConfigPath $ConfigPath
        if ($currentConfig.ReAddAtStart -eq $true -and $currentConfig.Drives.Count -gt 0) {
            foreach ($mapping in $currentConfig.Drives) {
                try {
                    $driveLetter = $mapping.Drive
                    $uncPath = $mapping.UNCPath
                    $useCred = $mapping.UseCredentials
                    $username = $mapping.Username
                    $password = $mapping.Password  # Renamed from 'pwd' to 'password'

                    # Prepare 'net use' command
                    if ($useCred -and $username -and $password) {
                        $cmd = "net use ${driveLetter} `"$uncPath`" `"$password`" /user:`"$username`" /persistent:yes"
                    } else {
                        $cmd = "net use ${driveLetter} `"$uncPath`" /persistent:yes"
                    }

                    # Execute the command
                    $result = cmd.exe /c $cmd 2>&1
                    if ($result -match "The command completed successfully") {
                        Write-Output "Successfully re-mapped drive ${driveLetter}: to $uncPath."
                        # Update the DataGridView
                        for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                            $row = $dataGridView.Rows[$i]
                            if ($row.Cells["Drive"].Value -eq $driveLetter -and $row.Cells["UNCPath"].Value -eq $uncPath) {
                                $row.Cells["Mapped"].Value = "Yes"
                                $row.Cells["AddedDate"].Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                            }
                        }
                    } else {
                        Write-Error "Failed to re-map drive ${driveLetter}: $result"
                    }
                } catch {
                    Write-Output "Error re-adding drive ${driveLetter}: $($_.Exception.Message)"
                }
            }

            # Save updated mappings to configuration
            $mappedDrives = @()
            for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
                $currentRow = $dataGridView.Rows[$i]
                if (-not $currentRow.IsNewRow) {
                    $mappedDrives += @{
                        Drive = $currentRow.Cells["Drive"].Value
                        UNCPath = $currentRow.Cells["UNCPath"].Value
                        AddedDate = $currentRow.Cells["AddedDate"].Value
                        Mapped = $currentRow.Cells["Mapped"].Value
                        Selected = $currentRow.Cells["Selected"].Value
                        UseCredentials = $config.Drives[$i].UseCredentials
                        Username = $config.Drives[$i].Username
                        Password = $config.Drives[$i].Password
                    }
                }
            }
            $newConfig = @{
                Drives = $mappedDrives
                ReAddAtStart = $currentConfig.ReAddAtStart
                StartOnWindowsStart = $currentConfig.StartOnWindowsStart
            }
            Export-Configuration -ConfigPath $ConfigPath -ConfigData $newConfig
        }
    }

    # Re-add drives at start if enabled
    Invoke-ReAddDrivesAtStart -ConfigPath $configPath

    # Show the form
    $form.ShowDialog()
}

# Main Execution
Show-DriveMappingForm
