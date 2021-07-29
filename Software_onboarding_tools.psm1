Function Connect-SCCM{
    import-module 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1' -Global

    Set-Location "LPC:"
}

function Deploy-Collection2ManyMachines {
    <#
    .SYNOPSIS
    Function to Deploye machines listed in CSV to a SCCM Collection
    by Steven Wight
    .DESCRIPTION
    Deploy-Collection2ManyMachines -collectionname <collectionname>
    -inputfile <PathandFileName.csv> Default = "E:\PowerShell\Software_Onboarding_Tools\Input\Hostnames.csv"
    -Namespace <Namespace> Default = "root\SMS\site_LPC"
    -SiteServer <SiteServer> Default = "SCCM01.POSHYT.com"
    .EXAMPLE
    Deploy-Collection2ManyMachines SCCM_Collection "E:\PowerShell\Software_Onboarding_Tools\Input\AlphaTestMachines.csv"
    .NOTES
    Should never need to change Namespace or SiteServer once set, Inputfile should be headerless file with hostnames on each row
    #>
    [cmdletbinding()]
    param (
        [Parameter()] [String] [ValidateNotNullOrEmpty()] $collectionname, 
        [Parameter()] [String] [ValidateNotNullOrEmpty()] $inputfile = "E:\PowerShell\Software_Onboarding_Tools\Input\Hostnames.csv",
        [Parameter()] [String] [ValidateNotNullOrEmpty()] [string] $Namespace = "root\SMS\site_LPC",
        [Parameter()] [String] [ValidateNotNullOrEmpty()] $SiteServer = "SCCM01.POSHYT.com"
    )
    #Open Connection to SCCM
    Connect-SCCM

    #Loop through the hostnames in the $inputfile.csv
    Import-Csv $infile -header Computer | ForEach-Object {

        #store hostname in variable for ease of use
        $Computer = $_.Computer
        
        #Get computers SCCM Resource ID
        $ResIDQuery = Get-WmiObject -ComputerName $SiteServer -Namespace $Namespace -Class "SMS_R_SYSTEM" -Filter "Name='$Computer'"
        $ResouceID = $ResIDQuery.resourceid
        
        #Output action to be taken to console
        Write-Host -ForegroundColor Cyan "adding $($collectionname) to $($Computer)"

        
        try{# Try and Add the computer to the collection and output to console if successful
            Add-CMDeviceCollectionDirectMembershipRule -CollectionName $collectionname -ResourceID  $ResouceID -ErrorAction Stop
            Write-Host -ForegroundColor Green "added $($collectionname) to $($Computer)"
        }catch{#If any issues, output to the console 
            Write-Host -ForegroundColor RED "Can't add $($collectionname) to $($Computer)" 
        }#End of Try...Catch  
    }#End of Foreach-Object
}#End of Function

function Get-ApplicationDetailsAndCollections {
    <#
    .SYNOPSIS
    Function to grab details of application and what collections it is part of
    by Steven Wight
    .DESCRIPTION
    Get-ApplicationDetailsAndCollections -Application <Application>
    -outputfile <PathandFileName.csv> Default = "E:\PowerShell\Software_Onboarding_Tools\output\$($Application)_Collections.csv"
    -Namespace <Namespace> Default = "root\SMS\site_LPC"
    -SiteServer <SiteServer> Default = "SCCM01.POSHYT.com"
    .EXAMPLE
    Get-ApplicationDetailsAndCollections "Bloomberg Terminal 3064.1.80.1" 
    .NOTES
    Should never need to change Namespace or SiteServer once set
    #>
    [cmdletbinding()]
    param (
        [Parameter()] [String] [ValidateNotNullOrEmpty()] $Application, 
        [Parameter()] [String] [ValidateNotNullOrEmpty()] $outputfile = "E:\PowerShell\Software_Onboarding_Tools\output\$($Application)_Collections.csv",
        [Parameter()] [String] [ValidateNotNullOrEmpty()] [string] $Namespace = "root\SMS\site_LPC",
        [Parameter()] [String] [ValidateNotNullOrEmpty()] $SiteServer = "SCCM01.POSHYT.com"
    )
    #Open Connection to SCCM
    Connect-SCCM

    Get-CMApplication $Application | Select-object LocalizedDisplayName,LocalizedDescription,Manufacturer, SoftwareVersion,NumberOfDeployments,NumberOfDevicesWithApp,NumberOfDevicesWithFailure
    $Collections = (Get-CMDeployment -FeatureType "Application" | Where-Object -Property ApplicationName -EQ -Value $Application | Select-Object CollectionName)
    $Collections | Format-Table
    $Collections  | export-csv $outputfile -NoTypeInformation -Append

}#End of Function

function Mass-DeployAndRemoveApplications {
    <#
    .SYNOPSIS
    Function to Deploy and remove appliactions from Collections
    by Steven Wight
    .DESCRIPTION
    Mass-DeployAndRemoveApplications inputfile <PathandFileName.csv> Default = "E:\PowerShell\Software_Onboarding_Tools\Input\DeployApplicationList.csv"
    -Namespace <Namespace> Default = "root\SMS\site_LPC"
    -SiteServer <SiteServer> Default = "SCCM01.POSHYT.com"
    .EXAMPLE
    Mass-DeployAndRemoveApplications
    .NOTES
    Should never need to change Namespace or SiteServer, Input file should be as Collection Name, New Application Name, Old Application Name
    e.g. - BloombergProfessional,Bloomberg Terminal 3064.1.80.1,Bloomberg Terminal 3062.1.80.3
    #>
    [cmdletbinding()]
    param ( 
        [Parameter()] [String] [ValidateNotNullOrEmpty()] $inputfile = "E:\PowerShell\Software_Onboarding_Tools\Input\DeployApplicationList.csv",
        [Parameter()] [String] [ValidateNotNullOrEmpty()] [string] $Namespace = "root\SMS\site_LPC",
        [Parameter()] [String] [ValidateNotNullOrEmpty()] $SiteServer = "SCCM01.POSHYT.com"
    )
    #Open Connection to SCCM
    Connect-SCCM

    #Initialise total counter
    $AppList = Import-Csv -Path $inputfile
    $AppList = $AppList.count

    #Initialise Progress and Error counter
    $i = 0
    $ErrorCount =0 

    #Loop through CSV rows
    Import-csv $inputfile -header CollectionName, Name, oldname | foreach-object{

        #Clear and load variables 
        $CollectionName = $ApplicationName = $null
        $CollectionName = $_.CollectionName
        $ApplicationName = $_.Name
        $oldApplicationName = $_.oldname

        #Start Progress Bar
        Write-Progress -Activity "Deploying Application to collection no. $($i) of $($AppList)" -CurrentOperation $CollectionName
        
        #increment progress counter
        $i++
        
        #Try to add deployemnt to collection
        Try {
            Start-CMApplicationDeployment -CollectionName $CollectionName -Name $ApplicationName -DeployAction install -DeployPurpose Required -RebootOutsideServiceWindow $true -SendWakeUpPacket $True 
        }Catch{

            Write-host "-----------------"
            Write-host -ForegroundColor Red ("There was an error adding $($ApplicationName)  to  $($CollectionName)   collection.")
            Write-host "-----------------"
            $ErrorCount += 1
            #print error message
            $ErrorMessage = $_.Exception.Message
            $ErrorMessage
        }#End of Try .... Catch 
        
        #Try Removing old deployment
        Try {
            Remove-CMApplicationDeployment -Name $oldApplicationName -CollectionName $CollectionName -force 
        }Catch{

            Write-host "-----------------"
            Write-host -ForegroundColor Red ("There was an error removing $($oldApplicationName)  to  $($CollectionName)   collection.")
            Write-host "-----------------"
            $ErrorCount += 1
            #print error message
            $ErrorMessage = $_.Exception.Message
            $ErrorMessage
        }#End of Try .... Catch 
        
    }#End of ForEach loop

    If ($ErrorCount -ge1) { #Notify if there has been errors (will be apparent if there was tho)
        Write-host "-----------------"
        Write-Host -ForegroundColor Red "The script execution completed, but with errors."
        Write-host "-----------------"
        Pause
    }Else{
        Write-host "-----------------"
        Write-Host -ForegroundColor Green "Script execution completed without error. All Deployments created sucessfully."
        Write-host "-----------------"
        Pause
    }

}#End of Function
