import-module 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1'
cd PYT:

#set domain and SCCM site details
[string] $Namespace = "root\SMS\site_PYT"
$SiteServer = "POSHSITE01.POSHYT.corp"
$Domain = 'POSHYT'

# Input output files names and paths
$InputFile = "c:\temp\Posh_inputs\SCCM_Collection_Discovery_Hostnames.csv"
$OutputFile = "c:\temp\Posh_outputs\SCCM_Collection_Discovery_$(get-date -f yyyy-MM-dd-HH-mm).csv"
$ErrorLog = "c:\temp\posh_outputs\SCCM_Collection_Discovery_$(get-date -f yyyy-MM-dd-HH-mm)_Errorlog.log"

#initialise error counter 
$ErrorCount=0

#initialise Progress counter 
$i=0

#get number of machines to check
$ComputersCount = Import-Csv $InputFile -Header Hostname
$ComputersCount = $ComputersCount.count


Import-Csv $InputFile -Header Hostname | ForEach-Object {

    # copy input into variables
    $hostname = $_.Hostname

    #clear other variables
    $ResIDQuery = $Computer = $null

    $i++ # main loop progress counter
    $percentageComplete = (($i / $ComputersCount) * 100)
    $percentageComplete = [math]::Round($percentageComplete,1)

    Write-Host -ForegroundColor Cyan "Querying Computer - $($hostname): $i of $($computerscount) - Percent complete $($percentageComplete) %"
    Write-host -ForegroundColor yellow "Testing if $hostname is in AD and online"

    #Find machine in AD
    Try{
        
        $Computer = (Get-ADComputer $hostname -server $Domain -ErrorAction Stop).name
        Write-host -ForegroundColor Green "Found $($hostname) in AD" 

    }catch{

        #increase error counter if something not right
        $ErrorCount += 1
        #print error message
        $ErrorMessage = $_.Exception.Message
        "Can't find $($hostname) in AD because $($ErrorMessage)" | Out-File $ErrorLog -Force -Append
        Write-host -ForegroundColor RED "Can't find $($hostname) in AD because $($ErrorMessage)"

    }

    #check machine is in SCCM
    try{

        $ResIDQuery = Get-WmiObject -ComputerName $SiteServer -Namespace $Namespace -Class "SMS_R_SYSTEM" -Filter "Name='$Computer'" -ErrorAction STOP
        Write-host -ForegroundColor Green "Found $($hostname) in SCCM" 

    }catch{

        #increase error counter if something not right
        $ErrorCount += 1
        #print error message
        $ErrorMessage = $_.Exception.Message
        "Can't find $($hostname) in SCCM because $($ErrorMessage)" | Out-File $ErrorLog -Force -Append
        Write-host -ForegroundColor RED "Can't find $($hostname) in SCCM because $($ErrorMessage)"

    }

    #If Machine in SCCM, extract collections
    if($null -ne $ResIDQuery){
        
        try{

            $Collections = (Get-WmiObject -ComputerName $SiteServer -Class sms_fullcollectionmembership -Namespace $Namespace -Filter "ResourceID = '$($ResIDQuery.ResourceId)'" -ErrorAction Stop)
            Write-host -ForegroundColor Green "Extracted $($hostname)'s collections" 

            $devicecollections = @()
            ForEach ($Collection in $collections){
    
                $colID = $Collection.CollectionID

                $collectioninfo = Get-WmiObject -ComputerName $SiteServer -Namespace $Namespace -Class "SMS_Collection" -Filter "CollectionID='$colID'"
                $object = New-Object -TypeName PSObject
                $object | Add-Member -MemberType NoteProperty -Name "Computer" -Value $hostname
                $object | Add-Member -MemberType NoteProperty -Name "Name" -Value $collectioninfo.Name
                $object | Add-Member -MemberType NoteProperty -Name "Commnent" -Value $collectioninfo.Comment
                
                $devicecollections += $object
            }#end of foreach collection in collections

            $devicecollections | Export-Csv $OutputFile -Append -NoTypeInformation



        }catch{

            #increase error counter if something not right
            $ErrorCount += 1
            #print error message
            $ErrorMessage = $_.Exception.Message
            "Can't Extract $($hostname)'s collections because $($ErrorMessage)" | Out-File $ErrorLog -Force -Append
            Write-host -ForegroundColor RED "Can't Extract $($hostname)'s collections because $($ErrorMessage)"

        }
        
        }#end of if Machine is in SCCM 


}#End of ForEach-object loop 


If ($ErrorCount -ge1) {

    Write-host "-----------------"
    Write-Host -ForegroundColor Red "The script execution completed, but with errors. See $($ErrorLog)"
    Write-host "-----------------"
    Pause
}Else{
    Write-host "-----------------"
    Write-Host -ForegroundColor Green "Script execution completed without error."
    Write-host "-----------------"
    Pause
}
