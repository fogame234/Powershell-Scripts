
# Windows shell for window pop
$wshell = (New-Object -ComObject Wscript.Shell)

# Importing modules

import-module AzureAD
import-module MSOnline

$local = (Get-Date)

# ScriptBlock for background job
$scriptBlock = {
    Connect-MsolService 
    while($local -ine "05:00 PM")
    {
        $time = (Get-MsolCompanyInformation | Select-Object lastdirsynctime -ExpandProperty lastdirsynctime)

        $msttime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($time, 'Mountain Standard Time')
        $msttimeformat = ([DateTime]$msttime).ToString("hh:mm tt")

        $start = (get-date)
        $timediff = ([datetime]$start - [datetime]$msttimeformat)
        $Hrs = $timediff.Hours
        $Mins = $timediff.Minutes
        
        $mstmsg = "Last dirsync was run at $msttimeformat `r`n`r`n"
        Write-output $mstmsg
        
        $difference = "Last dirsync was {0:00} hours and {1:00} minutes ago `r`n`r`n" -f $Hrs,$Mins
        Write-Output $difference
        Start-Sleep -Seconds 5
    }
}

# Start background job
Start-Job -ScriptBlock $scriptBlock -Name 'AzureDirsync'

#Timer and tick add
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 15000
$timer.add_tick({AddToLog $clock})

# Referenced functions
function AddToLog($logtext)
{ 
    $script:clock = Receive-Job -Name AzureDirsync
    $txtlog.Text = " "
    $txtLog.Text = $txtLog.Text + $logtext
    $txtLog.ScrolltoCaret
}

function startTimer() { 

   $timer.start()
   $wshell.Popup('Start Job',5,"Start",4096)

}

function stopTimer() {

    $timer.Enabled = $false
    $wshell.Popup('Stop Job',5,"Stop",4096)

}


########################
# Setup User Interface
########################

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

# Form Object
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Azure AD Sync Timer"
$objForm.Size = New-Object System.Drawing.Size(370,210) 
$objForm.StartPosition = "CenterScreen"


# Start Button
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Location = New-Object System.Drawing.Size(80,130)
$btnStart.AutoSize = $true
$btnStart.Text = "Start"
$btnStart.Add_Click({StartTimer; })

#Stop Button
$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Location = New-Object System.Drawing.Size(200,130)
$btnStop.AutoSize = $true
$btnStop.Text = "Stop"
$btnStop.Add_Click({StopTimer; })
$btnStop.Enabled  = $true

# Heading
$heading = New-Object System.Windows.Forms.Label
$heading.Location = New-Object System.Drawing.Size(85,10) 
$heading.AutoSize = $true 
$heading.Text = "Azure AD Sync Timer:"
$heading.Font = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular)

# Text Box
$txtLog = New-Object System.Windows.Forms.Textbox
$txtLog.Location = New-Object System.Drawing.Size(10,40)
$txtLog.Size = New-Object System.Drawing.Size(330,80)
$txtLog.Multiline = $True
$txtlog.ReadOnly = $true
#$txtlog.Text = "test`r`n`r`ntest"
$txtlog.Font = New-Object System.Drawing.Font("Times New Roman",12,[System.Drawing.FontStyle]::Regular)
#$txtlog.BorderStyle = 0
$txtLog.Add_Click({$txtLog.SelectAll(); $txtLog.Copy()})

# Draw object forms
$objForm.Controls.Add($btnStart)
$objForm.Controls.Add($btnStop)
$objForm.Controls.Add($heading) 
$objForm.Controls.Add($txtLog)
$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()


<#
stop-Job -Name AzureDirsync
remove-job -Name AzureDirsync
get-job#>
