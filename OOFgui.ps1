Add-Type -AssemblyName System.Windows.Forms

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Out of Office Manager"
$form.Size = New-Object System.Drawing.Size(400,400)
$form.StartPosition = "CenterScreen"

# Label & Button for CSV Import
$csvLabel = New-Object System.Windows.Forms.Label
$csvLabel.Text = "CSV File:"
$csvLabel.Location = New-Object System.Drawing.Point(10,20)
$csvLabel.AutoSize = $true
$form.Controls.Add($csvLabel)

$csvTextBox = New-Object System.Windows.Forms.TextBox
$csvTextBox.Location = New-Object System.Drawing.Point(70, 18)
$csvTextBox.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($csvTextBox)

$csvButton = New-Object System.Windows.Forms.Button
$csvButton.Text = "Browse"
$csvButton.Location = New-Object System.Drawing.Point(280,15)
$csvButton.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = "CSV Files (*.csv)|*.csv"
    if ($fileDialog.ShowDialog() -eq "OK") {
        $csvTextBox.Text = $fileDialog.FileName
    }
})
$form.Controls.Add($csvButton)

# Button to Create Example CSV
$exampleCsvButton = New-Object System.Windows.Forms.Button
$exampleCsvButton.Text = "Create Example CSV"
$exampleCsvButton.Location = New-Object System.Drawing.Point(10, 50)
$exampleCsvButton.Size = New-Object System.Drawing.Size(150, 25)
$exampleCsvButton.Add_Click({
    $exampleCsvPath = "C:\OutOfOffice-main\importautoreply_example.csv"
    @"
User
example@domain.com
anotheruser@domain.com
"@ | Out-File -Encoding utf8 -FilePath $exampleCsvPath
    [System.Windows.Forms.MessageBox]::Show("Example CSV created at $exampleCsvPath", "Info", "OK", "Information")
})
$form.Controls.Add($exampleCsvButton)

# Enable Scheduling Checkbox
$scheduleLabel = New-Object System.Windows.Forms.Label
$scheduleLabel.Text = "Schedule Auto-Reply:"
$scheduleLabel.Location = New-Object System.Drawing.Point(170, 50)
$scheduleLabel.AutoSize = $true
$form.Controls.Add($scheduleLabel)

$scheduleCheckbox = New-Object System.Windows.Forms.CheckBox
$scheduleCheckbox.Location = New-Object System.Drawing.Point(300, 50)
$scheduleCheckbox.Checked = $true
$scheduleCheckbox.Add_CheckStateChanged({
    $startDatePicker.Enabled = $scheduleCheckbox.Checked
    $endDatePicker.Enabled = $scheduleCheckbox.Checked
})
$form.Controls.Add($scheduleCheckbox)

# Start Date & Time Picker
$startLabel = New-Object System.Windows.Forms.Label
$startLabel.Text = "Start Date & Time:"
$startLabel.Location = New-Object System.Drawing.Point(10, 90)
$startLabel.AutoSize = $true
$form.Controls.Add($startLabel)

$startDatePicker = New-Object System.Windows.Forms.DateTimePicker
$startDatePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$startDatePicker.CustomFormat = "MM/dd/yyyy HH:mm:ss"
$startDatePicker.Location = New-Object System.Drawing.Point(150, 88)
$form.Controls.Add($startDatePicker)

# End Date & Time Picker
$endLabel = New-Object System.Windows.Forms.Label
$endLabel.Text = "End Date & Time:"
$endLabel.Location = New-Object System.Drawing.Point(10, 120)
$endLabel.AutoSize = $true
$form.Controls.Add($endLabel)

$endDatePicker = New-Object System.Windows.Forms.DateTimePicker
$endDatePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$endDatePicker.CustomFormat = "MM/dd/yyyy HH:mm:ss"
$endDatePicker.Location = New-Object System.Drawing.Point(150, 118)
$form.Controls.Add($endDatePicker)

# Message Input
$msgLabel = New-Object System.Windows.Forms.Label
$msgLabel.Text = "Auto-Reply Message:"
$msgLabel.Location = New-Object System.Drawing.Point(10, 150)
$msgLabel.AutoSize = $true
$form.Controls.Add($msgLabel)

$msgTextBox = New-Object System.Windows.Forms.TextBox
$msgTextBox.Location = New-Object System.Drawing.Point(10, 170)
$msgTextBox.Size = New-Object System.Drawing.Size(360, 80)
$msgTextBox.Multiline = $true
$msgTextBox.ScrollBars = "Vertical"
$form.Controls.Add($msgTextBox)

# Enable/Disable Toggle
$toggleLabel = New-Object System.Windows.Forms.Label
$toggleLabel.Text = "Enable Auto-Reply:"
$toggleLabel.Location = New-Object System.Drawing.Point(10, 250)
$toggleLabel.AutoSize = $true
$form.Controls.Add($toggleLabel)

$toggleCheckbox = New-Object System.Windows.Forms.CheckBox
$toggleCheckbox.Location = New-Object System.Drawing.Point(150, 250)
$toggleCheckbox.Checked = $true
$form.Controls.Add($toggleCheckbox)

# Log File Path Input
$logPathLabel = New-Object System.Windows.Forms.Label
$logPathLabel.Text = "Log File Path:"
$logPathLabel.Location = New-Object System.Drawing.Point(10, 280)
$logPathLabel.AutoSize = $true
$form.Controls.Add($logPathLabel)

$logPathTextBox = New-Object System.Windows.Forms.TextBox
$logPathTextBox.Location = New-Object System.Drawing.Point(100, 280)
$logPathTextBox.Size = New-Object System.Drawing.Size(200, 20)
$logPathTextBox.Text = "C:\OutOfOffice-main\log\AutoReplyLog.csv"  # Default log path
$form.Controls.Add($logPathTextBox)

# Apply Button - Updated to use inputted log path
$applyButton = New-Object System.Windows.Forms.Button
$applyButton.Text = "Apply"
$applyButton.Location = New-Object System.Drawing.Point(150, 320)
$applyButton.Add_Click({
    try {
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -ShowBanner:$false
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to connect to Exchange Online.", "Error", "OK", "Error")
        return
    }

    $csvFile = $csvTextBox.Text
    $startDate = $startDatePicker.Value.ToString("MM/dd/yyyy HH:mm:ss")
    $endDate = $endDatePicker.Value.ToString("MM/dd/yyyy HH:mm:ss")
    $message = $msgTextBox.Text
    $enableAutoReply = $toggleCheckbox.Checked
    $useSchedule = $scheduleCheckbox.Checked
    $logPath = $logPathTextBox.Text  # Use the inputted log file path

    if (-not (Test-Path $csvFile)) {
        [System.Windows.Forms.MessageBox]::Show("CSV file not found!", "Error", "OK", "Error")
        return
    }

    $Users = Import-Csv $csvFile
    $logResults = @()
    $timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")

    foreach ($user in $Users) {
        try {
            if ($enableAutoReply) {
                if ($useSchedule) {
                    Set-MailboxAutoreplyConfiguration -Identity $user.user -AutoReplyState Scheduled -StartTime $startDate -EndTime $endDate -InternalMessage $message -ExternalMessage $message
                    $status = "Success"
                } else {
                    Set-MailboxAutoreplyConfiguration -Identity $user.user -AutoReplyState Enabled -InternalMessage $message -ExternalMessage $message
                    $status = "Success"
                }
            } else {
                Set-MailboxAutoreplyConfiguration -Identity $user.user -AutoReplyState Disabled
                $status = "Disabled"
            }
            $logResults += [PSCustomObject]@{ UserEmail = $user.user; Status = $status; Timestamp = (Get-Date).ToString("MM/dd/yyyy HH:mm:ss") }
        } catch {
            $logResults += [PSCustomObject]@{ UserEmail = $user.user; Status = "Failed"; Timestamp = (Get-Date).ToString("MM/dd/yyyy HH:mm:ss") }
        }
    }

    $logResults | Export-Csv -Path $logPath -NoTypeInformation
    Disconnect-ExchangeOnline -Confirm:$false
    [System.Windows.Forms.MessageBox]::Show("Auto-reply update complete. Log saved to $logPath.", "Success", "OK", "Information")
})
$form.Controls.Add($applyButton)

$form.ShowDialog()
