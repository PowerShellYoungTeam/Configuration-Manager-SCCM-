#Modules
import-module ActiveDirectory
#For Guis
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Start-transcript "C:\temp\fun\Posh_Outputs\CCMAppInstallerLog$(get-date -f yyyyMMdd-HHmm).txt"

$Domain = 'POSHYT'

#########################
#
# Functions
#
#########################
Function Start-SCCMAppInstall{

    Param
    ([String][Parameter(Mandatory=$True, Position=1)] $Computername,
    [String][Parameter(Mandatory=$True, Position=2)] $AppName,
    [ValidateSet("Install","Uninstall")]
    [String][Parameter(Mandatory=$True, Position=3)] $Method
    )
    
    Begin {
    $Application = (Get-CimInstance -ClassName CCM_Application -Namespace "root\ccm\clientSDK" -ComputerName (get-adcomputer $Computername -server $Domain).dnshostname | Where-Object {$_.Name -like $AppName})
    
    $Args = @{EnforcePreference = [UINT32] 0
        Id = "$($Application.id)"
        IsMachineTarget = $Application.IsMachineTarget
        IsRebootIfNeeded = $False
        Priority = 'High'
        Revision = "$($Application.Revision)" 
    }
    
    } #End of Begin Block
    
    Process{
    
    Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Application -ComputerName (get-adcomputer $Computername -server $Domain).dnshostname -MethodName $Method -Arguments $Args
    
    } #End of Process Block
    
    End {}
    
} #End of Start-SCCMAppInstall Function

Function Start-ConfigmanangerActions {
    param($Hostname) 
    
    $Computer =  (Get-ADComputer -Identity $Hostname -server $Domain)
    
    if($null -eq $Computer){

        write-host -ForegroundColor Red "Not in AD - $Computer"

    }
    
    if($null -ne $Computer){

        Write-Progress -Activity 'Running Config Manager' -CurrentOperation $Hostname 

        $PathTest = Test-Connection -Computername $Computer.DNSHostName -BufferSize 16 -Count 1 -Quiet

        if($PathTest -eq $False) {
    
            write-host -ForegroundColor Red "OFFLINE - $Computer"

        }Else{
        
            #out message to console $Computer
            write-host -ForegroundColor Green "Online - $($Computer.DNSHostName)" 
            Invoke-Command -ComputerName $Computer.DNSHostName -ScriptBlock {
            Invoke-WmiMethod -Class SMS_Client -Name TriggerSchedule '{00000000-0000-0000-0000-000000000121}'  -Namespace root\ccm
            Invoke-WmiMethod -Class SMS_Client -Name TriggerSchedule '{00000000-0000-0000-0000-000000000003}'  -Namespace root\ccm
            Invoke-WmiMethod -Class SMS_Client -Name TriggerSchedule '{00000000-0000-0000-0000-000000000001}'  -Namespace root\ccm
            Invoke-WmiMethod -Class SMS_Client -Name TriggerSchedule '{00000000-0000-0000-0000-000000000021}'  -Namespace root\ccm
            Invoke-WmiMethod -Class SMS_Client -Name TriggerSchedule '{00000000-0000-0000-0000-000000000022}'  -Namespace root\ccm
            Invoke-WmiMethod -Class SMS_Client -Name TriggerSchedule '{00000000-0000-0000-0000-000000000108}'  -Namespace root\ccm
            Invoke-WmiMethod -Class SMS_Client -Name TriggerSchedule '{00000000-0000-0000-0000-000000000113}'  -Namespace root\ccm
            Invoke-WmiMethod -Class SMS_Client -Name TriggerSchedule '{00000000-0000-0000-0000-000000000111}'  -Namespace root\ccm
            Invoke-WmiMethod -Class SMS_Client -Name TriggerSchedule '{00000000-0000-0000-0000-000000000032}'  -Namespace root\ccm}

        }#End of IF Else

    }#end of if

}#End of Start-ConfigmanangerActions Function

Function Select-CCMMachineGUI{

    ### Enter Hostname GUI ###
    $ChooseHostnameForm = New-Object System.Windows.Forms.Form
    $ChooseHostnameForm.Text = 'SCCM App Installer'
    $ChooseHostnameForm.Size = New-Object System.Drawing.Size(350,250)
    $ChooseHostnameForm.StartPosition = 'CenterScreen'

    #Enter Hostname Label and Textbox
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(50,20)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = 'Enter Hostname:'
    $ChooseHostnameForm.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(50,40)
    $textBox.Size = New-Object System.Drawing.Size(260,20)
    $textbox.Text = "Enter hostname here"
    $ChooseHostnameForm.Controls.Add($textBox)

    $ConfigMgrActionsCheckbox = New-Object System.Windows.Forms.Checkbox 
    $ConfigMgrActionsCheckbox.Location = New-Object System.Drawing.Size(50,80) 
    $ConfigMgrActionsCheckbox.Size = New-Object System.Drawing.Size(500,20)
    $ConfigMgrActionsCheckbox.Text = "Run Config Manager Actions"
    $ConfigMgrActionsCheckbox.TabIndex = 1
    $ConfigMgrActionsCheckbox.Checked = $true
    $ChooseHostnameForm.Controls.Add($ConfigMgrActionsCheckbox)

    $Methodlabel = New-Object System.Windows.Forms.Label
    $Methodlabel.Location = New-Object System.Drawing.Point(50,110)
    $Methodlabel.Size = New-Object System.Drawing.Size(280,20)
    $Methodlabel.Text = 'Choose Action'
    $ChooseHostnameForm.Controls.Add($Methodlabel)

    $MethodList = New-Object system.Windows.Forms.ComboBox
    $MethodList.width = 170
    $MethodList.autosize = $true
    # Add the items in the dropdown list
    @('Install','Uninstall') | ForEach-Object {[void] $MethodList.Items.Add($_)}
    # Select the default value
    $MethodList.SelectedIndex = 0
    $MethodList.location = New-Object System.Drawing.Point(50,130)
    $ChooseHostnameForm.Controls.Add($MethodList)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(75,180)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $ChooseHostnameForm.AcceptButton = $okButton
    $ChooseHostnameForm.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(200,180)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $ChooseHostnameForm.CancelButton = $cancelButton
    $ChooseHostnameForm.Controls.Add($cancelButton)

    $ChooseHostnameForm.Topmost = $true

    $ChooseHostnameForm.Add_Shown({$textBox.Select()})
    $result = $ChooseHostnameForm.ShowDialog()

    $result
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $Hostname = $textbox.Text
        $Script:Hostname = $Hostname
        $Script:Method = $MethodList.Text

        if($true -eq $ConfigMgrActionsCheckbox.Checked){

            Start-ConfigmanangerActions $Hostname
    
        }

        $Script:CCMApps = (Get-CimInstance -ClassName CCM_Application -Namespace "root\ccm\clientSDK" -ComputerName (get-adcomputer $Hostname -server $Domain).dnshostname  | Select-Object name, installstate, ErrorCode, AllowedActions, Publisher, SoftwareVersion, LastInstallTime, LastEvalTime |Sort-Object name)
        
    }else{

    Write-Warning "Exiting Script"
    break
    }

}# End of Select-CCMMachineGUI Function

    #########################
#
# End of Functions
#
#########################

$CCMApps = @()

$Hostname = $Null
$Method = $Null

Do{

    Select-CCMMachineGUI

}while ($Null -eq $CCMApps)

$ActionApps = ($CCMApps | Out-GridView -PassThru -Title "Select what apps you want to install" )

Foreach($ActionApp in $ActionApps){

    Start-SCCMAppInstall $Hostname $ActionApp.name Install

}

Stop-Transcript
