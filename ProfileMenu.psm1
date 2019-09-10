
Function ProfileMenu {
        
    #Return if Left Shift is held down
    $KeyTest = [System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftShift)
    If ($KeyTest -eq "True") {
        $Global:WindowTitle = $Host.UI.RawUI.WindowTitle
        Start-Sleep -Milliseconds 200
        $wshell2 = New-Object -ComObject wscript.shell
        $wshell2.AppActivate($Global:WindowTitle)
        Start-Sleep -Milliseconds 150
        Write-Host "Running Standard "
        $wshell2.SendKeys('^n')
        return}
    
    ##############################
    ## Session Entry State Test ##
    ##############################
    #region EntryStateTest
    
        #Set menu to initially run
        $ProfileMenuRun = 1

        ## Figure out what is going on with the Session you are opening up to...
        #This is primarily for handling if you right-click on a file and open it, the menu will not run, because the file will pass a Test-Path, and will not be Utitled*

        #If the Tab Count is gt 0, Test-Path of current open tab(s) 
        If ($psISE.PowerShellTabs.files.FullPath.Count -gt 0) {
            
            $TestPath = Test-Path $psISE.CurrentPowerShellTab.Files.fullpath  

            ## Single File Open - GetType is Boolean
            IF ($TestPath.GetType().Name -eq "Boolean") {
                IF ($TestPath -eq "False") {$ProfileMenuRun = 1} #If there is a single file open, and it is not an actual file ($TestPath -eq "False") then run the Profile Menu becaues its an untitled, default opened file.
                
                IF ($psISE.CurrentPowerShellTab.Files.IsSaved -eq $False) {$ProfileMenuRun --} #Don't run if the untitled file has been modified and not saved
                IF (($psISE.CurrentPowerShellTab.Files.IsSaved -eq $True) -AND ($TestPath -eq $True)) {$ProfileMenuRun --} #Don't run if the untitled file is saved, named Untitled and is a real file
            }

            #Mulitiple files open - GetType is Object[]
            IF ($TestPath.GetType().Name -eq "Object[]") {
                IF ($TestPath.Contains("True")) {$ProfileMenuRun --} #If one of them are a real file, do not run
            }
        }

        #If there are NO Tabs open, this is a fresh session, so run the ProfileMenu
        IF ($psISE.PowerShellTabs.files.FullPath.Count -eq 0) {$ProfileMenuRun = 1} 
        
        If ($ProfileMenuRun -lt 1) {return}
    #endregion EntryStateTest
    
    ## Set up CRSPath variable to reference the current running directory
    IF ($host.Name -eq 'Windows PowerShell ISE Host') {
        If ($psISE.CurrentFile.FullPath -eq $Null) {$Global:CRSPath = Split-Path $PSCommandPath}
        If ($psISe.CurrentFile.FullPath -ne $Null) {$Global:CRSPath = Split-Path $psISE.CurrentFile.FullPath;Write-Host "The CRSPAth is: $CRSPath";pause}
    }
    
    If ($host.name -eq 'ConsoleHost') {$Global:CRSPath = Split-Path $PSCommandPath}
    
    #Load DATA / SETTINGS
    $Global:MenuItemsCSV = Get-Content $Global:CRSPath\Profile_MenuItems.csv | ConvertFrom-Csv
      
    #Cleanup
    $Global:CMDLineRun = ""
    $Global:CMDLine = ""

    # Formatting Variables
    $Divider = {Write-Host -no " == "}

    ##Close the Untitled File
    $wshell = New-Object -ComObject wscript.shell
    $wshell.SendKeys("^{F4}")

    ###############
    ## Main Menu ##
    ###############
    ##Main Menu Header
    Write-Host "##############################################"
    Write-Host -no "## ";Write-Host -no "Profile: RunCommon / Change Window Title" -f Red;Write-Host " ##"
    Write-Host "##############################################"
    Write-Host ""

    #Generate Main Menu
    foreach ($item in $MenuItemsCSV) {
    Write-Host -no $item.ItemNum -f Red;&$Divider;Write-Host $item.ItemName -f $item.ItemColor
    Write-Host ""
    }

    $Global:StartChoice = Read-Host "Type in your choice and hit ENTER:  > "

    #Off by 1 Preventer (Because of INDEX number)
    $Global:StartChoice = $Global:StartChoice -1

    ## Process Menu Item / Run Items ##
    $Global:ItemToOpen = $MenuItemsCSV[$StartChoice].ItemToOpen
    $Global:WindowTitle = $MenuItemsCSV[$StartChoice].ItemName;$WindowTitle
    $Global:ItemAutoEnteredCommand = $MenuItemsCSV[$StartChoice].ItemAutoEnteredCommand

    Clear-Host;Clear-Host

    #######################
    ## File Open Process ## 
    #######################
    #Open the File\Open prompt
    Invoke-Command -ScriptBlock {
        $wshell = New-Object -ComObject wscript.shell;
        $wshell.AppActivate($Global:WindowTitle)
        Sleep 1
        $wshell.SendKeys('^{F4}')
        $wshell.SendKeys('^{o}')
        
        #Go to the Default Directory if ItemToOpen IS Blank
        IF ($MenuItemsCSV[$StartChoice].ItemToOpen -le "") {
        $wshell.SendKeys("$DefaultOpenDir")
        $wshell.SendKeys('{ENTER}')
        }
        
        #Enter in the ItemToOpen, if it is NOT Blank
        If ($MenuItemsCSV[$StartChoice].ItemToOpen -gt "") {
        $wshell.SendKeys("$ItemToOpen")
        $wshell.SendKeys('{ENTER}')
        Start-Sleep 2 #Wait for file to open
        }    
        
        $host.ui.RawUI.WindowTitle = “$WindowTitle”
        
        }

    Function Test-RegistryEntryExistence {
        Param (
        [parameter(Mandatory)]$Path,
        [parameter(Mandatory)]$ItemToTest
        )

    $Result = (Get-Item -Path $Path).GetValueNames() -match $ItemToTest

    If ($Result -notcontains "$ItemToTest") {return $false}
    If ($Result -contains $ItemToTest) {return $true}
    }
   
    ##############################################
    ## Start the PrepProcess if PrepProcess = 1 ##
    ##############################################
    IF ($MenuItemsCSV[$StartChoice].PrepProcess -eq 1) {
        #Push out variables for the Jobs if the MenuItem involves PrepProcess

        #If Reg Entry DOES EXIST, update value
        IF ((Test-RegistryEntryExistence -Path HKCU:\Environment -ItemToTest PSItemAutoEnteredCommand) -eq $True) {
            #Update Value
            Set-ItemProperty -Path "HKCU:\Environment" -Name "PSItemAutoEnteredCommand" -Value "$ItemAutoEnteredCommand"
            }
        IF ((Test-RegistryEntryExistence -Path HKCU:\Environment -ItemToTest PSWindowTitle) -eq $True) {
            #Update Value
            Set-ItemProperty -Path "HKCU:\Environment" -Name "PSWindowTitle" -Value "$WindowTitle"
            }

        #If Reg Entry DOES NOT EXIST, create and update
        IF ((Test-RegistryEntryExistence -Path HKCU:\Environment -ItemToTest PSItemAutoEnteredCommand) -eq $False) {
            #Create Entry
            New-ItemProperty -Path "HKCU:\Environment" -Name "PSItemAutoEnteredCommand" -Value "$ItemAutoEnteredCommand"  -PropertyType "ExpandString"    
            #Update Value
            }
        IF ((Test-RegistryEntryExistence -Path HKCU:\Environment -ItemToTest PSWindowTitle) -eq $False) {
            #Create Entry
            New-ItemProperty -Path "HKCU:\Environment" -Name "PSWindowTitle" -Value "$WindowTitle"  -PropertyType "ExpandString"    
            #Update Value
            }

        #Run the Prep Process (as a JOB (background process)), if PrepProcess -eq 1
        $Global:BreakProfile = 1;Start-Job -Name PrepProcess -FilePath "$Global:CRSPath\PrepProcess.ps1" | Out-Null
        Wait-Job -Name PrepProcess | Out-Null
        
        #Run the ItemAEC Scriptblock
        Start-Job -Name ItemAEC -ScriptBlock {
            #Set up Variables
            $Global:WindowTitle = (Get-Item -Path HKCU:\Environment).GetValue("PSWindowTitle")
            $ItemAutoEnteredCommand = (Get-Item -Path HKCU:\Environment).GetValue("PSItemAutoEnteredCommand")
            Start-Sleep 3
            $wshellItemAEC = New-Object -ComObject wscript.shell;
            $wshellItemAEC.AppActivate($WindowTitle) | Out-Null
            $wshellItemAEC.SendKeys('^d')
            Start-Sleep -Milliseconds 300
            If ($ItemAutoEnteredCommand -eq 'F5') {$wshellItemAEC.SendKeys('{F5}')}
            If ($ItemAutoEnteredCommand -ne 'F5') {$wshellItemAEC.SendKeys($ItemAutoEnteredCommand)}
            }
        }
}