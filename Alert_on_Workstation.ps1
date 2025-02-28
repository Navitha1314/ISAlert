#####################################################################################################
### AUTHOR : NAVITHA GOVINDRAJ
### DATE   : 05 MARCH 2024
### REASON : SCRIPT TO DISPLAY THE IS ALERT POPUP ON THE CLIENT WORKSTATION
#####################################################################################################
Param(
       [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Param1
)
$alert  = $Param1.replace("GH","`n")
Start-Transcript "$env:Systemdrive\OHMSILogs\ISAlert_on_workstation.log" -Append -Force
Add-Type -AssemblyName system.windows.Forms
$Form = New-Object system.windows.forms.form
$Form.ControlBox = $False
[System.Windows.Forms.Application]::EnableVisualStyles() 
$monitor = [System.Windows.Forms.Screen]::PrimaryScreen
[void]::$monitor.WorkingArea.Width
#$Form.Text = "  OHIOHEALTH - IMPORTANT UPDATE" 
$Form.Height = '500'
$Form.Width = '600'
$Form.AutoSize = $False
$Form.AutoScale = $False
$Form.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#0071aa')
$Form.StartPosition = "Manual"
$Form.Top = $monitor.WorkingArea.Height - $form.Height
$Form.Left = $monitor.WorkingArea.Width - $Form.Width
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$Form.TopMost = $true
$Icon = [System.Drawing.SystemIcons]::Asterisk
$Form.Icon = $Icon
$l1Font = New-Object System.Drawing.Font("Calibri",15,[System.Drawing.FontStyle]::Bold)
$l2Font = New-Object System.Drawing.Font("Calibri",11,[System.Drawing.FontStyle]::Regular)
############### WHITE BACKGROUND ##################
$Label = New-Object System.Windows.Forms.Label
$Label.Text = ""
$Label.AutoSize = $false
$Label.Height = '300'
$Label.Width = '575'
$Label.BorderStyle = "Fixed3D"
$Label.BackColor = [System.Drawing.Color]::White
$Label.Location = New-Object System.Drawing.Size(10,140)
$Label.SendToBack()
############### LOGO #############################
$Pic1 = New-Object System.Windows.Forms.PictureBox
$Pic1.AutoSize = $false
$Pic1.Height = '100'
$Pic1.Width = '575'
$Pic1.Location = New-Object System.Drawing.Size(10,20)
$Pic1base64 = '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'
$Pic1Bytes = [Convert]::FromBase64String($Pic1base64)
$Pic1stream = [System.IO.MemoryStream]::new($Pic1Bytes, 0, $Pic1Bytes.Length)
$Pic1.Image = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($Pic1stream).GetHIcon()))
$Pic1.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Normal
$Pic1.BackColor = "White"
$Pic1.BorderStyle = "FixedSingle"
#################### NOT REQD ################
$Label1 = New-Object System.Windows.Forms.Label
$Label1.Text = "Important alert"
$Label1.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$Label1.AutoSize = $false
$Label1.Height = '89'
$Label1.Width = '250'
$Label1.Location = New-Object System.Drawing.Size(333,21)
$l1Font = New-Object System.Drawing.Font("Century",15,[System.Drawing.FontStyle]::Bold)
$l2Font = New-Object System.Drawing.Font("Calibri",11,[System.Drawing.FontStyle]::Regular)
$Label1.Font = $l1Font
$Label1.ForeColor = "Red"
$Label1.BackColor = "White"
#$Label1.BorderStyle = "FixedSingle"
################ MAIN TEXT BOX ###########################
$Label2 = New-Object System.Windows.Forms.Label
$Label2.Text = $alert
$Label2.BackColor = "White"
$Label2.AutoSize = $false
$Label2.Height = '299'
$Label2.Width = '504'
$Label2.Font = $l2Font
$label2.BringToFront()
#$Label2.MaximumSize = New-Object System.Drawing.Size(450,300)
$Label2.Location = New-Object System.Drawing.Size(80,141)
################# ALERT ICON ##############################
$alerticon = New-Object System.Windows.Forms.PictureBox
$alerticonbase64 = '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'
$alerticonBytes = [Convert]::FromBase64String($alerticonbase64)
$alertstream = [System.IO.MemoryStream]::new($alerticonBytes, 0, $alerticonBytes.Length)
$alerticon.Image = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($alertstream).GetHIcon()))
$alerticon.Location = New-Object System.Drawing.Size(20,150)
$alerticon.Height = 55
################# OK BUTTON ###########################
$Button1 = New-Object System.Windows.Forms.Button
$Button1.Text = "OK"
$Button1.Width = 70
$Button1.Height = 40
$button1.BackColor = "White"
$Button1.Location = New-Object System.Drawing.Point(272,450)
$Button1.Font = "Microsoft Sans Serif,10"
$Button1.Add_Click({$Form.Close()})

<#$label3 = New-Object Windows.Forms.Label
$label3.Text = "Time remaining: 10 seconds"
$label3.Location = New-Object Drawing.Point(100,300)

# Create a timer
$timer1 = New-Object Windows.Forms.Timer
$timer1.Interval = 1000  # 1 second
$timer1.Add_Tick({
    $timeLeft = [Math]::Max(0, $timeLeft - 1)
    $label3.Text = "Time remaining: $timeLeft seconds"
    if ($timeLeft -eq 0) {
        $timer1.Stop()
        $label3.Text = "Time's up!"
    }
})
$timeLeft = 10  # Set the initial countdown time
$timer1.Start()#>
Write-Output "Script is being executed"
$Form.Controls.Add($Label1)

$Form.Controls.Add($Pic1)
$Form.Controls.Add($Label2)
$Form.Controls.Add($Button1)
$Form.Controls.Add($alerticon)
$Form.Controls.Add($Label)
Stop-transcript
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 60000
$timer.Add_Tick({
$form.Close() | Out-Null
})
$timer.Start()
$form.ShowDialog() | Out-Null
$timer.Dispose()
$form.Dispose()
