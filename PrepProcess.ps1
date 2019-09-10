Start-Sleep 1
[console]::Beep(250,300)
Write-Host "Success"

IF ($host.Name -eq 'Windows PowerShell ISE Host') {
        IF ($psISE.CurrentFile.FullPath -eq $Null) {$Global:CRSPath = Split-Path $PSCommandPath}
        IF ($psISe.CurrentFile.FullPath -ne $Null) {$Global:CRSPath = Split-Path $psISE.CurrentFile.FullPath;Write-Host "The CRSPAth is: $CRSPath"}
    }
    
IF ($host.name -eq 'ConsoleHost') {$Global:CRSPath = Split-Path $PSCommandPath}

###################
## Variable Pull ##
###################
#Set Variable for WindowTitle

$WindowTitle = (Get-Item -Path HKCU:\Environment).GetValue("PSWindowTitle")

#Set-Variable -Name ItemAutoEnteredCommand -Value (Get-Content -Path $Global:CRSPath\VarTransfer-ItemAutoEnteredCommand.txt) -Force -Scope Global
$ItemAutoEnteredCommand = (Get-Item -Path HKCU:\Environment).GetValue("PSItemAutoEnteredCommand")

Start-Sleep -Milliseconds 200
$wshellPrepProcess = New-Object -ComObject wscript.shell
$wshellPrepProcess.AppActivate($WindowTitle) | Out-Null
Start-Sleep -Milliseconds 150
    
$wshellPrepProcess.SendKeys('^i')
$wshellPrepProcess.SendKeys('^a')
$wshellPrepProcess.SendKeys('^a')
Start-Sleep -Milliseconds 100
    
$wshellPrepProcess.SendKeys('{F8}')
Start-Sleep -Milliseconds 100
$wshellPrepProcess.SendKeys('^{END}')
$wshellPrepProcess.SendKeys('^{END}')
$wshellPrepProcess.SendKeys('{HOME}')        
Start-Sleep -Milliseconds 100

$wshellPrepProcess.SendKeys('^d')
$wshellPrepProcess.SendKeys('{ESC}')
$wshellPrepProcess.SendKeys("CLS")
$wshellPrepProcess.SendKeys('{ENTER}')
Start-Sleep -Milliseconds 300

[console]::Beep(500,300)