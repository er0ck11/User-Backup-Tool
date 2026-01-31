
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Configuration & Path Setup ---
$ScriptPath = $PSScriptRoot
$RclonePath = Join-Path $ScriptPath "rclone.exe"
$JsonPath   = Join-Path $ScriptPath "smb_targets.json"

# Check for Rclone
if (-not (Test-Path $RclonePath)) {
    [System.Windows.Forms.MessageBox]::Show("rclone.exe not found in script directory.`nPlease download it and place it alongside this script.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# --- Global State for Cancellation ---
$Global:JobState = $null

# --- UI Creation ---
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "User Backup Tool"
$Form.Size = New-Object System.Drawing.Size(600, 750)
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = "FixedDialog"
$Form.MaximizeBox = $false

# 1. Source Selection (Users)
$GroupSource = New-Object System.Windows.Forms.GroupBox
$GroupSource.Location = New-Object System.Drawing.Point(20, 20)
$GroupSource.Size = New-Object System.Drawing.Size(540, 150)
$GroupSource.Text = "1. Select Source Users (C:\Users)"
$Form.Controls.Add($GroupSource)

$CheckListUsers = New-Object System.Windows.Forms.CheckedListBox
$CheckListUsers.Location = New-Object System.Drawing.Point(15, 25)
$CheckListUsers.Size = New-Object System.Drawing.Size(510, 110)
$CheckListUsers.CheckOnClick = $true
$GroupSource.Controls.Add($CheckListUsers)

# Populate Users (Exclude Public, Default, All Users)
Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notin @("Public", "Default", "All Users", "Default User") } | ForEach-Object {
    $CheckListUsers.Items.Add($_.Name)
}

# 2. Exclusions
$GroupExclude = New-Object System.Windows.Forms.GroupBox
$GroupExclude.Location = New-Object System.Drawing.Point(20, 180)
$GroupExclude.Size = New-Object System.Drawing.Size(540, 60)
$GroupExclude.Text = "2. AppData Exclusions"
$Form.Controls.Add($GroupExclude)

$ChkExLocal = New-Object System.Windows.Forms.CheckBox
$ChkExLocal.Text = "Exclude AppData\Local"
$ChkExLocal.Location = New-Object System.Drawing.Point(20, 25)
$ChkExLocal.AutoSize = $true
$GroupExclude.Controls.Add($ChkExLocal)

$ChkExLocalLow = New-Object System.Windows.Forms.CheckBox
$ChkExLocalLow.Text = "Exclude AppData\LocalLow"
$ChkExLocalLow.Location = New-Object System.Drawing.Point(180, 25)
$ChkExLocalLow.AutoSize = $true
$GroupExclude.Controls.Add($ChkExLocalLow)

$ChkExRoaming = New-Object System.Windows.Forms.CheckBox
$ChkExRoaming.Text = "Exclude AppData\Roaming"
$ChkExRoaming.Location = New-Object System.Drawing.Point(360, 25)
$ChkExRoaming.AutoSize = $true
$GroupExclude.Controls.Add($ChkExRoaming)

# 3. Destination Type & Config
$GroupDest = New-Object System.Windows.Forms.GroupBox
$GroupDest.Location = New-Object System.Drawing.Point(20, 250)
$GroupDest.Size = New-Object System.Drawing.Size(540, 220)
$GroupDest.Text = "3. Destination Configuration"
$Form.Controls.Add($GroupDest)

$RadioLocal = New-Object System.Windows.Forms.RadioButton
$RadioLocal.Text = "Local / Mapped Drive"
$RadioLocal.Location = New-Object System.Drawing.Point(20, 30)
$RadioLocal.AutoSize = $true
$RadioLocal.Checked = $true
$GroupDest.Controls.Add($RadioLocal)

$RadioSMB = New-Object System.Windows.Forms.RadioButton
$RadioSMB.Text = "SMB Share (Network)"
$RadioSMB.Location = New-Object System.Drawing.Point(180, 30)
$RadioSMB.AutoSize = $true
$GroupDest.Controls.Add($RadioSMB)

# SMB Controls Panel
$PanelSMB = New-Object System.Windows.Forms.Panel
$PanelSMB.Location = New-Object System.Drawing.Point(10, 60)
$PanelSMB.Size = New-Object System.Drawing.Size(520, 100)
$PanelSMB.Visible = $false # Hidden by default
$GroupDest.Controls.Add($PanelSMB)

$LblShare = New-Object System.Windows.Forms.Label
$LblShare.Text = "Select Share:"
$LblShare.Location = New-Object System.Drawing.Point(10, 10)
$LblShare.AutoSize = $true
$PanelSMB.Controls.Add($LblShare)

$ComboSMB = New-Object System.Windows.Forms.ComboBox
$ComboSMB.Location = New-Object System.Drawing.Point(100, 5)
$ComboSMB.Size = New-Object System.Drawing.Size(400, 25)
$ComboSMB.DropDownStyle = "DropDownList"
$PanelSMB.Controls.Add($ComboSMB)

# Load JSON
if (Test-Path $JsonPath) {
    try {
        $JsonData = Get-Content $JsonPath | ConvertFrom-Json
        foreach ($item in $JsonData) {
            $ComboSMB.Items.Add($item) # Store object, DisplayMember handles text
        }
        $ComboSMB.DisplayMember = "Name"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error reading smb_targets.json", "Warning")
    }
} else {
    $ComboSMB.Items.Add("No JSON found")
}

$LblUser = New-Object System.Windows.Forms.Label
$LblUser.Text = "Username:"
$LblUser.Location = New-Object System.Drawing.Point(10, 45)
$LblUser.AutoSize = $true
$PanelSMB.Controls.Add($LblUser)

$TxtUser = New-Object System.Windows.Forms.TextBox
$TxtUser.Location = New-Object System.Drawing.Point(100, 40)
$TxtUser.Size = New-Object System.Drawing.Size(150, 20)
$PanelSMB.Controls.Add($TxtUser)

$LblPass = New-Object System.Windows.Forms.Label
$LblPass.Text = "Password:"
$LblPass.Location = New-Object System.Drawing.Point(260, 45)
$LblPass.AutoSize = $true
$PanelSMB.Controls.Add($LblPass)

$TxtPass = New-Object System.Windows.Forms.TextBox
$TxtPass.Location = New-Object System.Drawing.Point(330, 40)
$TxtPass.Size = New-Object System.Drawing.Size(170, 20)
$TxtPass.UseSystemPasswordChar = $true
$PanelSMB.Controls.Add($TxtPass)

# Destination Path Input (Shared by both modes)
$LblDestPath = New-Object System.Windows.Forms.Label
$LblDestPath.Text = "Target Folder:"
$LblDestPath.Location = New-Object System.Drawing.Point(20, 180)
$LblDestPath.AutoSize = $true
$GroupDest.Controls.Add($LblDestPath)

$TxtDestPath = New-Object System.Windows.Forms.TextBox
$TxtDestPath.Location = New-Object System.Drawing.Point(110, 175)
$TxtDestPath.Size = New-Object System.Drawing.Size(350, 20)
# Prepopulate Date ONLY (Removed "_Backup")
$TxtDestPath.Text = (Get-Date).ToString("yyyy-MM-dd")
$GroupDest.Controls.Add($TxtDestPath)

$BtnBrowse = New-Object System.Windows.Forms.Button
$BtnBrowse.Text = "..."
$BtnBrowse.Location = New-Object System.Drawing.Point(470, 173)
$BtnBrowse.Size = New-Object System.Drawing.Size(50, 23)
$GroupDest.Controls.Add($BtnBrowse)

# 4. Progress & Actions
$LblStatus = New-Object System.Windows.Forms.Label
$LblStatus.Text = "Status: Ready"
$LblStatus.Location = New-Object System.Drawing.Point(20, 485)
$LblStatus.Size = New-Object System.Drawing.Size(540, 20)
$Form.Controls.Add($LblStatus)

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(20, 510)
$ProgressBar.Size = New-Object System.Drawing.Size(540, 30)
$Form.Controls.Add($ProgressBar)

$BtnStart = New-Object System.Windows.Forms.Button
$BtnStart.Text = "Start Backup"
$BtnStart.Location = New-Object System.Drawing.Point(20, 550)
$BtnStart.Size = New-Object System.Drawing.Size(120, 40)
$BtnStart.BackColor = "LightGreen"
$Form.Controls.Add($BtnStart)

$BtnCancel = New-Object System.Windows.Forms.Button
$BtnCancel.Text = "Cancel"
$BtnCancel.Location = New-Object System.Drawing.Point(440, 550)
$BtnCancel.Size = New-Object System.Drawing.Size(120, 40)
$BtnCancel.Enabled = $false
$Form.Controls.Add($BtnCancel)

$TxtLog = New-Object System.Windows.Forms.TextBox
$TxtLog.Multiline = $true
$TxtLog.ScrollBars = "Vertical"
$TxtLog.Location = New-Object System.Drawing.Point(20, 600)
$TxtLog.Size = New-Object System.Drawing.Size(540, 100)
$TxtLog.ReadOnly = $true
$Form.Controls.Add($TxtLog)

# --- Event Handlers ---

# Toggle SMB Panel
$RadioSMB.Add_Click({ $PanelSMB.Visible = $true; $BtnBrowse.Enabled = $false })
$RadioLocal.Add_Click({ $PanelSMB.Visible = $false; $BtnBrowse.Enabled = $true })

# Browse Button (Only for Local) - Updated to remove "_Backup"
$BtnBrowse.Add_Click({
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($FolderBrowser.ShowDialog() -eq "OK") {
        # Joins the selected path with ONLY the current date
        $TxtDestPath.Text = Join-Path $FolderBrowser.SelectedPath ((Get-Date).ToString("yyyy-MM-dd"))
    }
})

# Start Button Logic
$BtnStart.Add_Click({
    
    # 1. Validation
    if ($CheckListUsers.CheckedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one user.", "Error")
        return
    }

    $DestBase = $TxtDestPath.Text
    $IsSMB = $RadioSMB.Checked
    
    # Gather SMB Credentials if needed
    $SmbShare = $null
    $SmbUser = $TxtUser.Text
    $SmbPass = $TxtPass.Text
    
    if ($IsSMB) {
        if ($ComboSMB.SelectedItem -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("Please select an SMB Share.", "Error")
            return
        }
        $SmbShare = $ComboSMB.SelectedItem.Path
        # Construct full UNC path for rclone destination
        # Rclone treats UNC paths as local paths
        $FinalDestPath = Join-Path $SmbShare $DestBase
    } else {
        $FinalDestPath = $DestBase
    }

    # Gather Exclusions
    $Exclusions = @()
    if ($ChkExLocal.Checked) { $Exclusions += "--exclude", "/AppData/Local/**" }
    if ($ChkExLocalLow.Checked) { $Exclusions += "--exclude", "/AppData/LocalLow/**" }
    if ($ChkExRoaming.Checked) { $Exclusions += "--exclude", "/AppData/Roaming/**" }
    
    # Gather Users
    $SelectedUsers = @($CheckListUsers.CheckedItems)

    # 2. UI State Update
    $BtnStart.Enabled = $false
    $BtnCancel.Enabled = $true
    $GroupSource.Enabled = $false
    $GroupDest.Enabled = $false
    $ProgressBar.Style = "Marquee"
    $TxtLog.Text = "Initializing..."

    # 3. Create Background Job (to keep UI responsive)
    $ScriptBlock = {
        param($Users, $RcloneExe, $DestPath, $IsSMBMode, $SmbSharePath, $SmbU, $SmbP, $ExcludeList)

        $Log = @()
        
        # Helper to add log
        function Write-Log ($msg) { 
            Write-Output "LOG:$msg"
        }

        # SMB Connection Logic
        if ($IsSMBMode) {
            Write-Log "Attempting SMB connection to $SmbSharePath..."
            # Remove existing mapping if any
            net use $SmbSharePath /delete /y 2>&1 | Out-Null
            
            if ($SmbU -ne "") {
                $netCmd = "net use `"$SmbSharePath`" `"$SmbP`" /user:`"$SmbU`""
                cmd /c $netCmd 2>&1 | Out-String | ForEach-Object { Write-Log $_ }
            } else {
                cmd /c "net use `"$SmbSharePath`"" 2>&1 | Out-String | ForEach-Object { Write-Log $_ }
            }
            
            if (-not (Test-Path $SmbSharePath)) {
                Write-Log "ERROR: Could not connect to SMB share."
                return
            }
        }

        # Loop through users
        foreach ($User in $Users) {
            $Source = "C:\Users\$User"
            $UserDest = Join-Path $DestPath $User
            
            Write-Log "Backing up $User..."
            Write-Log "Source: $Source"
            Write-Log "Dest: $UserDest"

            # Build Rclone Args
            # Using --transfers 4 and -P (progress) is hard to parse in non-interactive, 
            # so we use -v for verbose logging which we can parse if needed, or just --stats
            
            $ArgsList = @("copy", "`"$Source`"", "`"$UserDest`"", "--stats", "2s", "--ignore-errors")
            $ArgsList += $ExcludeList

            # Run Rclone
            $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
            $ProcessInfo.FileName = $RcloneExe
            $ProcessInfo.Arguments = $ArgsList -join " "
            $ProcessInfo.RedirectStandardOutput = $true
            $ProcessInfo.RedirectStandardError = $true
            $ProcessInfo.UseShellExecute = $false
            $ProcessInfo.CreateNoWindow = $true

            $Process = New-Object System.Diagnostics.Process
            $Process.StartInfo = $ProcessInfo
            $Process.Start() | Out-Null

            # Monitor Process for Cancel or Output
            while (-not $Process.HasExited) {
                # Check for cancellation signal (File check logic is clumsy, 
                # we will rely on stopping the job from the parent)
                
                # Read stream (non-blocking is hard in pure PS job, so we just wait a bit)
                Start-Sleep -Milliseconds 500
            }
            $StdOut = $Process.StandardOutput.ReadToEnd()
            $StdErr = $Process.StandardError.ReadToEnd()
            
            Write-Log "Rclone Output: $StdErr" # Rclone stats usually go to stderr
        }

        # Cleanup SMB
        if ($IsSMBMode) {
            net use $SmbSharePath /delete /y 2>&1 | Out-Null
        }

        Write-Log "Backup Complete."
    }

    # Start the Job
    $Global:JobState = Start-Job -ScriptBlock $ScriptBlock -ArgumentList $SelectedUsers, $RclonePath, $FinalDestPath, $IsSMB, $SmbShare, $SmbUser, $SmbPass, $Exclusions
    
    # 4. Timer to monitor Job
    $Timer = New-Object System.Windows.Forms.Timer
    $Timer.Interval = 1000
    $Timer.Add_Tick({
        if ($Global:JobState.State -eq "Running") {
            # Fetch generic output
            $NewData = Receive-Job -Job $Global:JobState
            if ($NewData) {
                foreach ($line in $NewData) {
                    if ($line -like "LOG:*") {
                        $TxtLog.AppendText($line.Substring(4) + "`r`n")
                    }
                }
            }
        } else {
            # Job Finished
            $Timer.Stop()
            $BtnStart.Enabled = $true
            $BtnCancel.Enabled = $false
            $GroupSource.Enabled = $true
            $GroupDest.Enabled = $true
            $ProgressBar.Style = "Blocks"
            $ProgressBar.Value = 100
            $LblStatus.Text = "Status: " + $Global:JobState.State
            
            # Get final output
            $FinalData = Receive-Job -Job $Global:JobState
            if ($FinalData) { 
                foreach ($line in $FinalData) {
                    if ($line -like "LOG:*") { $TxtLog.AppendText($line.Substring(4) + "`r`n") }
                }
            }
            Remove-Job -Job $Global:JobState
        }
    })
    $Timer.Start()
})

# Cancel Button Logic
$BtnCancel.Add_Click({
    if ($Global:JobState -and $Global:JobState.State -eq "Running") {
        $LblStatus.Text = "Status: Cancelling..."
        Stop-Job -Job $Global:JobState
        $TxtLog.AppendText("User requested cancel.`r`n")
        
        # Kill stray rclone processes? 
        # CAUTION: This kills ALL rclone processes. 
        # In a production environment, you'd want to track the specific PID.
        Stop-Process -Name "rclone" -ErrorAction SilentlyContinue
    }
})

# --- Run ---
$Form.Add_Shown({$Form.Activate()})
[void]$Form.ShowDialog()