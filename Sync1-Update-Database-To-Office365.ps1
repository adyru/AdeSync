
<#
This script is to update an SQL database with AD Objects
This data will then be used in a subsequent script to create contacts in another directory
#>
# Below are the functions that we will be using in this script
Function Write-Log {
    <#
        . Notes
        =======================================
        v1  Created on: 25/05/2021
            Created by AU             
        =======================================
        . Description
        Params are Level and Message - it will use these as the output along with the date 
        Depending on how $OutputScreenInfo and  $OutputScreenNotInfo are set it will also output to screen
        You need to define the output file  $logfile  in the main script block
                write-log -level info -message "Something"

        #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $False)]
        [ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]
        [String]
        $Level = "INFO",

        [Parameter(Mandatory = $True)]
        [string]
        $Message
    
    )
    # Check if we have a logfile defined and quit script if not with error to screen
    If (!($logfile)) {
        write-host -foregroundcolor red "Logfile not defind so exiting - create a variable called $logfile  and try again"
        exit

    }
    # Set these to $true or $false - switches on and off output to screen
    # One if for info eventss the other for anything but info
    $OutputScreenInfo = $false
    $OutputScreenNotInfo = $true
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    # Create the line from the timestamp, level and message
    $Line = "$Stamp $Level $Message"
    # Create seperate log file for errors
    $ErrorLog = $logfile -replace ".txt", "-errors.txt"
    #write-host "Output is $($outputscreen)"
    # Check if the level isnt info and then write to error log
    If ($level -ne "info") {
        Add-Content $ErrorLog  -Value $Line
        # Check if we want this outputted to screen, if so it is unexpected so error to screen
        If ($OutputScreenNotInfo -eq $true) {
            write-host -ForegroundColor Red $Message 
        }
    }
    # Write to log for info events

        Add-Content $logfile -Value $Line
        # And if required output these to screen
        If ($OutputScreenInfo -eq $true) {
            write-host  $Message 
        }
    

}

Function Export-File {
    <#
        . Notes
        =======================================
        v1  Created on: 13/05/2021
            Created by AU  
        V1.1 - 21/05/2021 Added in a check so that if files exists it exports with a random digit inserted           
        =======================================
        . Description
        Params are variable and filename 
        it then exports the variable to a csv if it isnt null
        If files exists adds in some random digits to makes sure there is an export
            export-file $var $varout
        #>
    [CmdletBinding()]
        
    Param(
        [Parameter(Mandatory = $false)]
        [Array]
        $OutVar,

        [Parameter(Mandatory = $True)]
        [string]
        $OutFile

    )
    Begin {
        $checkFile = test-path $outfile
    }
    Process {
        If ($checkFile -eq $true) {
            # if the files exists we will use  a random number into the filename 
            write-host -ForegroundColor Red "$($outfile) already exists"

            If (!($OutVar)) {
                write-host "Variable is null"
                return;
            }
            Else {
                $Random = Get-Random -Minimum 1 -Maximum 100
                $newOut = $outfile -replace ".csv", "$($Random).csv"
                $OutVar | Export-Csv -NoClobber -NoTypeInformation -path $newout -Encoding UTF8
                $OutcheckFile = test-path $outfile
                If ($OutcheckFile -eq $true)
                { write-host -ForegroundColor Green "Variable exported to $($newout)" }
                Else
                { write-host -ForegroundColor red "Variable not exported to $($newout)" }
            }
            return;
        }
        Else {
            If (!($OutVar)) {
                write-host "Variable is null"
                return;
            }
            Else {
                $OutVar | Export-Csv -NoClobber -NoTypeInformation -path $OutFile -Encoding UTF8
                $OutcheckFile = test-path $outfile
                If ($OutcheckFile -eq $true)
                { write-host -ForegroundColor Green "Variable exported to $($OutFile)" }
                Else
                { write-host -ForegroundColor red "Variable not exported to $($OutFile)" }
            }
        }
            
 

    }
    End {

    }
}


Function Nettestserver {
    <#
        .Description
        test-server server port.Params are server and port - accepts either common ports or numbers
        it then trys to resolve the DNS name and catches the error if cant 
        if it succeeds it goes on to test the connection to a port 
        returns values back to script block in $tcpclient -  Timeout for connection is $Timer in MS
        #>
    [CmdletBinding()]

    Param(
        [Parameter(Mandatory = $True)]
        [string]
        $servertest,

        [Parameter(Mandatory = $True)]
        [string]
        $Porttest
    )
    try {
        # Check DNS and just keep the first one
        $lookups = $null
        $Lookups = (Resolve-DnsName $servertest -DnsOnly -ErrorAction Stop).IP4Address
        $DNSCheck = $Lookups | Select-Object -First 1
    }
    #Catch the error in the DNS record doesnt exist
    catch [System.ComponentModel.Win32Exception] {
        Write-host -ForegroundColor red $servertest " was not found"
        exit
    }
    
    # Null out array
    If ([STRING]::IsNullOrWhitespace($DNSCheck)) {
    }
    Else {
        # If it is a numerical port do a check
        if ($porttest -match "^\d{1,3}") {
            try {
                $tcpclient = New-Object System.Net.Sockets.TCPClient
                $Timer = 1500
                $StartConnection = $tcpclient.BeginConnect($servertest, $PortTest, $null, $null)
                $wait = $StartConnection.AsyncWaitHandle.WaitOne($timer, $false)
                return  Write-Output -NoEnumerate  $tcpclient
            }

            catch {
                return  Write-Output -NoEnumerate  $tcpclient
            }
        }
                    
        Else {
            write-host -ForegroundColor Red "You have entered and incorrect port $($porttest) - it needs to either be a number"
        }
    }
}




######################
# Start of Variables #
######################

# Static DC to use - needs to be FQDN this will be checked later on
$DC = "RPBDOMR01.RPB01.babgroup.co.uk"

#Instance No
$Instance = "1"

#Date that we will look for new/update users from
$Currentdate = get-date -Format dd-MM-yyyy--hh-mm
$StartDate = (get-date).AddDays(-10)


#Create new folder for output files if it doesnt exist using todays date
$FolderDate = Get-Date -Format dd-MM-yyyy
$newFolder = "C:\Scripts\RPA\Output\$($FolderDate)\Update-Database-To-O365\Instance$($instance)"
$TestFolder = test-path $newFolder
If ($testfolder -eq $false) { New-Item -Path $NewFolder -ItemType directory }
$OutputFolder = $NewFolder


#Need to create a logfile for the write-log function as the script will error out without that
$logfile = "$OutputFolder\Instance$($instance)-Update-Database-To-O36-Log-$($Currentdate).txt"
start-transcript -path "$OutputFolder\Instance$($instance)-Update-Database-To-O365-transcript-$($Currentdate).txt"

#Keep track of users processed by this script
$UsersArray =  @()
$UsersArrayCounter = 0

# List Of OUs to process
$OUS = get-content "$($PSScriptRoot)\Allowed-OUs.txt"

#SQL Variables
$SQLServer = "RPBWEBR11.RPB01.babgroup.co.uk"
$SQLDatabase = "AzureGuestReports"
$SQLTable = "tmpRPA"
$SetSQLTable = "dbo.tmpRPA"
$DatabaseData = Read-SqlTableData -serverInstance $SQLServer -database $SQLDatabase   -TableName $SQLTable  -SchemaName dbo
$sqlUpdateDate = (get-date).ToString('yyyy-MM-dd')
$RemoveDate = ((get-date).adddays(-3)).ToString('yyyy-MM-dd')

# We are going to do a garbage collection every 2 mins so 
# need to kick of a timer
#Start Stopwatch
$sw = [diagnostics.stopwatch]::StartNew()

########################
#  End Of Variables    #
########################

###Revisit

#DC Connectivity Check - Exits on failure, user is given 10 secs to check DC
#write-host "Domain Controller to use is  $($DC)"
#Start-Sleep -Seconds 10

Write-Log -Message "Setting Domain Controller to  $($DC )" #write-host
Write-Log -Message "Checking connectivity to Domain Controller $($DC )" #write-host

# Check using the nettestserver function 
$CheckCOnnection = Nettestserver $DC 135

#$CheckCOnnection = test-netconnection $DC -port 135
# Check what the status of the connection is
If ($CheckCOnnection.Connected -eq $false) {
    # It failed - report and exit script
    write-host -ForegroundColor red "Connection is $($CheckCOnnection.Connected) so exiting script, amend DC variable and re run" 
    Write-Log -Message "Connectivity to Domain Controller $($DC ) is $($CheckCOnnection.Connected ) so exiting script " #write-host -Level "Fatal"
    exit
}
Else {
    #All good so carry on and report to screen
    write-host -ForegroundColor Green "Connectivity to Domain Controller $($DC ) Succdded - check connection is  $($CheckCOnnection.Connected ) "
    Write-Log -Message "Connectivity to Domain Controller $($DC ) Succdded - check connection is  $($CheckCOnnection.Connected ) " #write-host
}


#Check SQL Connection 
$conn = New-Object System.Data.SqlClient.SqlConnection                                      
$conn.ConnectionString = "Server=$SQLServer;Database=$Database;Integrated Security=True;"                                                                        
$conn.Open()
IF($conn.State -ne "Open")
    {
    Write-Log -Message "Failed to connect to database $($Database) on server $($SQLServer) so exiting" -Level "Error"
    exit
    }
Else
    {
    Write-Log -Message "Connected to database $($Database) on server $($SQLServer)"
    }




# We will loop though the OUs in the text to pick up users
write-log -Message "Picking up OUs from text file" 
ForEach($OU in $OUs)
    {
    Write-Log -Message "Starting in $($OU)" 
    # We need to change the target address used based on OU - for this we will use the OUs to change the target address needed
    if($OU -like "*OU=QCS Objects,DC=RPB01,DC=babgroup,DC=co,DC=uk")
        {
        # It is a white user
        #$TargetAddressSearch = "smtp:*@RPW01.babgroup.co.uk" 
        $TargetAddressSearch = "smtp:$($ADuser.Samaccountname)@RPW01.babgroup.co.uk" 
        Write-Log -Message "We are in $($OU) so setting the target address email search to $($TargetAddressSearch)" 
        }
    Else
        {
        # It is a blue user
        #$TargetAddressSearch = "smtp:*@RPB01.babgroup.co.uk"
        $TargetAddressSearch = "smtp:$($ADuser.Samaccountname)@RPB01.babgroup.co.uk"
        Write-Log -Message "We are in $($OU) so setting the target address email search to $($TargetAddressSearch)" 
        }
    # We will go and find users if they have an email address and are not hidden from the address list 
    Write-Log -Message "We are going to get the AD users from $($OU)"    
    $ADUsers = get-aduser -filter {emailaddress -like "*@*"} -SearchBase $OU -properties * | ? {$_.msExchHideFromAddressLists -ne "True"} #| select -first 2
    #$ADUsers = get-aduser -filter {emailaddress -like "*@*" -and msExchHideFromAddressLists -ne "True"} -SearchBase $OU -properties * | ? {$_.msExchHideFromAddressLists -ne "True"} #| select -first 2
    #$ADUsers =get-aduser -filter {emailaddress -like "*@*" -and msExchHideFromAddressLists -ne "True"} -SearchBase $OU -properties *
        If(!($ADUsers))
            {
            Write-Log -Message "No users found in AD" -Level Error
            }
        Else
            {
            Write-Log -Message "$(($ADUsers | measure).count) found in AD" 
            }  


    #First Of Update Instance for users who have moved OU
    $SQLQuery = "SELECT *  FROM $($SetSQLTable) WHERE Instance = 'Moved'"
    $movedContacts = Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLDatabase -Query $SQLQuery
    If(!($movedContacts))
        {
        Write-Log -Message "No users flagged in the database as moved" 
        }
    Else
        {
        Write-Log -Message "$(($movedContacts | measure).count) moved users found in SQL Database" 
        }  
    #$movedContacts 
    start-sleep -Seconds 2 
    write-log -message  "Starting to look at objects flagged as moved in the database"
    ForEach($movedContact in $movedContacts)
        {
        write-host "Moved user address $($movedContact.PrimarySMTPAddress)"
        $CheckUserInAD = get-aduser -filter "emailaddress -like '$($movedContact.PrimarySMTPAddress)'"
        if(!($CheckUserInAD))
            {
            write-host "we are going to delete"
            write-log -message  "$($movedContact.PrimarySMTPAddress)  doesnt exists in AD - so flag for deletion"
            #$SQLDelete = "DELETE FROM $SetSQLTable WHERE PrimarySMTPAddress='$($movedContact.PrimarySMTPAddress -replace "'","''")'"
            $SQLQuery = "UPDATE $SetSQLTable SET RecordDeleted='True' WHERE PrimarySMTPAddress = '$($movedContact.PrimarySMTPAddress -replace "'","''")'" 
            Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase  -Query $SQLQuery
            }
        Else
            {
            $CheckUserInOU  = $ADUsers | ? { $_.mail -eq $movedContact.PrimarySMTPAddress} 
            If(!($CheckUserInOU))
                {
                write-host "here - exists in AD but not in this OU"
                write-log -message  "$($movedContact.PrimarySMTPAddres) exists in AD but not in this OU"
                }
            Else
                {
                write-log -message  "$($movedContact.PrimarySMTPAddress) found in $($OU) so updating database to point this user at instance $($Instance)"
                $SQLQuery = "UPDATE $SetSQLTable SET Instance=$($Instance) WHERE PrimarySMTPAddress= '$($movedContact.PrimarySMTPAddress -replace "'","''")'"
                Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase  -Query $SQLQuery
                }

  
            }
 
        }


    # We will tehn loop through these users
    Write-Log -Message "Starting to look through the $(($ADUsers | measure).count) found in AD" 
    ForEach($ADUser in $ADUsers)
        {
        #Do garbage collection every couple of minutes to stop memory going off piste
        # so check if it is around 2 mins
            if ( $Sw.Elapsed.minutes -eq 2) {
                # it is over 2 mins so start garbage collection
                Write-Log -Message "Doing Garbage Collection after $($Sw.Elapsed.minutes ) minutes"  #write-host 
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers();
                #Reset timer by stopping and starting a new one
                $Sw.Stop()
                $sw = [diagnostics.stopwatch]::StartNew()

            }
        # First of check if they are a exchange recipient 
        Write-Log -Message "Checking if $($ADUser.emailaddress) is an exchange recipient" 
        $checkifMB = get-recipient $ADUser.emailaddress -erroraction silentlycontinue
        # if they arent a exchagne recipient
        If(!($checkifMB))
            {
            # we will log it and move on to the next user
            write-log -message  "$($ADUser.emailaddress) isnt an mail enabled object" -Level Error
            }
        # if they are an exchange user we will carry on
        Else
            {
            write-log -message  "$($ADUser.emailaddress) is an mail enabled object"
        # A counter i might use later on
        $UsersArrayCounter++
        # We will now look at all the users email addresses as we need to pick a target address
        write-log -message  "Checking the email addresses on $($ADUser.emailaddress)"
                write-log -message  "Looking for an address that matches smtp:$($ADuser.Samaccountname)@RPB01.babgroup.co.uk on user $($ADuser.Samaccountname)"
        ForEach($EmailAddress in $ADuser.proxyAddresses)
            {
            # Here we look to get an email address that matches the variable $TargetAddressSearch
            If($EmailAddress -like "$($TargetAddressSearch)") {$TargetAddress = $EmailAddress -replace "smtp:",""}
            }
            #if they dont have a match for this we will go through and just find a match for any email address in teh domain $TargetAddressSearch
            write-log -message  "There were no matches for smtp:$($ADuser.Samaccountname)@RPB01.babgroup.co.uk so we are now looking for smtp:*@RPB01.babgroup.co.uk"
            if(!($TargetAddress))
                {
                $TargetAddressSearch = "smtp:*@RPB01.babgroup.co.uk"
                ForEach($EmailAddress in $ADuser.proxyAddresses)
                    {
                    # Here we look to get an email address that matches the variable $TargetAddressSearch
                    If($EmailAddress -like "$($TargetAddressSearch)") {$TargetAddress = $EmailAddress -replace "smtp:",""}
                    }
                }
            write-log -message  "We have picked $($TargetAddress) as the target address"
            #write-host "Target address is $($TargetAddress)"
            #write the matching address to screen
            #Write-host "Target address in $($TargetAddress)"
            #write-host " I am here $($ADUser.Mail)"
            #Add users to the users array - we use that later on to check if the database  has users that arent in AD
            write-log -message  "Adding $($ADUser.Mail) to the users array object"
            $UsersArrayObj = New-Object System.Object
            $UsersArrayObj | Add-Member -type NoteProperty -name PrimarySMTPAddress -Value $ADUser.Mail
            $UsersArray += $UsersArrayObj
            #Check users in in Database currently by using their primary smtp address
            write-log -message  "Checking to see if $($ADUser.Mail) is in the database"
            $CheckPrimaryInDB  = $DatabaseData | ? { $_.PrimarySMTPAddress.Trim() -eq $ADUser.Mail}
            #There arnt so we need to add them 
            [int]$random = get-random -Maximum 999999
            # The user is not in the databse so we need to add them
                If(!($CheckPrimaryInDB))
                    {
                    write-log -message  "$($ADUser.Mail) isnt in the database so we will add them"
                    #Remove the null from the database later in if the user doesnt have certain attributes set as it gets on my nervves
                    IF(!($ADUser.GivenName)){$ADUser.GivenName = ""}
                    IF(!($ADUser.Surname)){$ADUser.Surname = ""}
                    IF(!($ADUser.Mobile)){$ADUser.Mobile = ""}
                    IF(!($ADUser.Office)){$ADUser.Office = ""}
                    IF(!($ADUser.telephoneNumber)){$ADUser.telephoneNumber = ""}
                    IF(!($ADUser.Company)){$ADUser.Company = ""}
                    IF(!($ADUser.Department)){$ADUser.Department = ""}
                    IF(!($ADUser.Country)){$ADUser.Country = $null}
                    IF(!($ADUser.Title)){$ADUser.Title = ""}
                    IF(!($ADUser.StreetAddress)){$ADUser.StreetAddress = ""}
                    IF(!($ADUser.postalcode)){$ADUser.postalcode = ""}
                    IF(!($ADUser.city)){$ADUser.city = ""}
                    IF(!($ADUser.State)){$ADUser.State = ""}
                    # The user has not been found so we will add them to the database
                     write-host -ForegroundColor Red "$($AdUser.Mail) is not in Database"
                    # You always forget this but this command actually punts the data into the database by order
                    # The actual fields arent used - si check the SQL rows if you get weird SQL errors
                    $Obj = [PSCustomObject] @{PrimarySMTPAddress=$ADUser.Mail;`
                                 FirstName=$ADUser.GivenName; `
                                 Surname=$ADUser.Surname; `
                                 TargetAddress=$TargetAddress; `
                                 telephoneNumber=$ADUser.telephoneNumber; `
                                 Mobile=$ADUser.Mobile; `
                                 Office=$ADUser.office; `
                                 Company=$ADUser.Company; `
                                 Department=$ADUser.Department; `
                                 RecordAdded=(get-date).ToString('yyyy-MM-dd');`
                                 RecordUpdated=(get-date).ToString('yyyy-MM-dd');`
                                 RecordDeleted="False"; `
                                 Title=$ADUser.Title; `
                                 Country=$ADUser.Country; `
                                 StreetAddress=$ADUser.StreetAddress; `
                                 postalcode=$ADUser.postalcode; `
                                 city=$ADUser.city; `
                                 State=$ADUser.State; `
                                 Instance=$Instance; `
                                 FirstRunComplete="False"
                                 MovedDate=$null
                                 Random=$Random
                                 }
         # Write this to the SQL database
         Write-SqlTableData -serverInstance $SQLServer -database $SQLDatabase  -TableName $SQLTable -SchemaName dbo -InputData $Obj
                    }
                #They are in the database so we will check that the other details are the same
                # Basically we loop through each attribute and check the database and AD Match
                Else
                    {
                    write-log -message  "$($ADUser.Mail) is in the database"
                    # First of lets check to see if the first name is the same as in AD - we assume the user HAS to have a first name!
                    write-log -message  "Checking if their first name in the database $($CheckPrimaryInDB.Firstname) matchs the AD value $($ADUser.GivenName)"
                    If($($CheckPrimaryInDB.Firstname) -ne $ADUser.GivenName)
                        {
                        # If isnt so we need to update the database
                        write-host "first name change"
                        write-log -message  "The first name values dont match so updating database with $($ADUser.GivenName)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Firstname='$($ADUser.GivenName)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }

                    # check to see if the Surname  is the same as in AD
                    # First of 2 ways of seeing if both are null
                    write-log -message  "Checking if their surname name in the database $($CheckPrimaryInDB.Surname) matchs the AD value $($ADUser.Surname)"
                    If(($CheckPrimaryInDB.Surname -eq $null) -and $(!($ADUser.Surname))){write-log -message  "Both surname values are null"}
                    If(!($CheckPrimaryInDB.Surname) -and $(!($ADUser.Surname))){write-log -message  "Both surname values are null"}
                    # It isnt null in both
                    ElseIf($CheckPrimaryInDB.Surname -eq $null)
                        {
                        write-log -message  "The surnames value dont match so updating database with $($ADUser.Surname)"
                        # It is null in the database so we will update with the AD value
                        write-host "Surname change $($ADUser.Surname)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Surname='$($ADUser.Surname)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    ElseIf($($CheckPrimaryInDB.Surname) -ne $ADUser.Surname)
                        {
                        write-log -message  "The surnames value dont match so updating database with $($ADUser.Surname)"
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "Surname change $($ADUser.Surname)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Surname='$($ADUser.Surname)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    
                    #We will assume the target address cannot be null and it needs to be se
                    write-log -message  "Checking if their target address in the database $($CheckPrimaryInDB.TargetAddress) matchs the AD value $($TargetAddress)"
                    If($($CheckPrimaryInDB.TargetAddress) -ne $TargetAddress)
                        {
                        # It is a different value so we will update with the AD value
                        write-host "TargetAddress change $($TargetAddress)"
                        write-log -message  "The TargetAddresses value dont match so updating database with $($TargetAddress)"
                        $SQLQuery = "UPDATE $SetSQLTable SET TargetAddress='$($TargetAddress -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }

                    # First of 2 ways of seeing if both Mobile fields are null
                    write-log -message  "Checking if their Mobile in the database $($CheckPrimaryInDB.Mobile) matchs the AD value $($ADUser.Mobile)"
                    If(($CheckPrimaryInDB.Mobile -eq $null) -and $(!($ADUser.Mobile))){write-log -message  "Both Mobile values are null"}
                    If(!($CheckPrimaryInDB.Mobile ) -and $(!($ADUser.Mobile))){write-log -message  "Both surname Mobile are null"}
                    ElseIf($CheckPrimaryInDB.Mobile -eq $null)
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "Mobile change $($ADUser.Mobile)"
                        write-log -message  "The Mobiles value dont match so updating database with $($ADUser.Mobile)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Mobile='$($ADUser.Mobile)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    ElseIf($($CheckPrimaryInDB.Mobile) -ne $ADUser.Mobile)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "Mobile change $($ADUser.Mobile)"
                        write-log -message  "The Mobiles value dont match so updating database with $($ADUser.Mobile)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Mobile='$($ADUser.Mobile)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    
                    # First of 2 ways of seeing if both Phone fields are null
                    write-log -message  "Checking if their telephoneNumber in the database $($CheckPrimaryInDB.telephoneNumber) matchs the AD value $($ADUser.telephoneNumber)"
                    If(!($CheckPrimaryInDB.telephoneNumber) -and $(!($ADUser.telephoneNumber))){write-log -message  "Both Telephone values are null"}
                    ElseIf(!($CheckPrimaryInDB.telephoneNumber))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "telephoneNumber change $($ADUser.telephoneNumber)"
                        write-log -message  "The TelephoneNumbers value dont match so updating database with $($ADUser.telephoneNumber)"
                        $SQLQuery = "UPDATE $SetSQLTable SET telephoneNumber='$($ADUser.telephoneNumber)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    ElseIf($($CheckPrimaryInDB.telephoneNumber) -ne $ADUser.telephoneNumber)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "telephoneNumber change $($ADUser.telephoneNumber)"
                        write-log -message  "The TelephoneNumbers value dont match so updating database with $($ADUser.telephoneNumber)"
                        $SQLQuery = "UPDATE $SetSQLTable SET telephoneNumber='$($ADUser.telephoneNumber)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    
                    # First of 2 ways of seeing if both Office fields are null
                    write-log -message  "Checking if their Office in the database $($CheckPrimaryInDB.Office) matchs the AD value $($ADUser.Office)"
                    If(!($CheckPrimaryInDB.Office) -and $(!($ADUser.Office))){write-log -message  "Both Office values are null"}
                    ElseIf(!($CheckPrimaryInDB.Office))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "Office change $($ADUser.Office)"
                        write-log -message  "The Office value dont match so updating database with $($ADUser.Office)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Office='$($ADUser.Office)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    ElseIf($($CheckPrimaryInDB.Office) -ne $ADUser.Office)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "Office change $($ADUser.Office)"
                        write-log -message  "The Office value dont match so updating database with $($ADUser.Office)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Office='$($ADUser.Office)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    
                     write-log -message  "Checking if their Department in the database $($CheckPrimaryInDB.Department) matchs the AD value $($ADUser.Department)"
                    # First of 2 ways of seeing if both Department fields are null
                    If(!($CheckPrimaryInDB.Department) -and $(!($ADUser.Department))){write-log -message  "Both Department values are null"}
                    ElseIf(!($CheckPrimaryInDB.Department))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "Department change $($ADUser.Department)"
                        write-log -message  "The Office value dont match so updating database with $($ADUser.Department)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Department='$($ADUser.Department)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    ElseIf($($CheckPrimaryInDB.Department) -ne $ADUser.Department)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "Department change $($ADUser.Department)"
                        write-log -message  "The Office value dont match so updating database with $($ADUser.Department)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Department='$($ADUser.Department)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    write-log -message  "Checking if their Company in the database $($CheckPrimaryInDB.Company) matchs the AD value $($ADUser.Company)"
                    # First of 2 ways of seeing if both Company fields are null
                    If(!($CheckPrimaryInDB.Company) -and $(!($ADUser.Company))){write-log -message  "Both Company values are null"}
                    ElseIf(!($CheckPrimaryInDB.Company))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "Company change $($ADUser.Company)"
                        write-log -message  "The Office value dont match so updating database with $($ADUser.Company)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Company='$($ADUser.Company)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    ElseIf($($CheckPrimaryInDB.Company) -ne $ADUser.Company)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "Company change $($ADUser.Company)"
                        write-log -message  "The Office value dont match so updating database with $($ADUser.Company)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Company='$($ADUser.Company)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''" )'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    
                    write-log -message  "Checking if their Country in the database $($CheckPrimaryInDB.Country) matchs the AD value $($ADUser.Country)"
                    # First of 2 ways of seeing if both Country fields are null
                    If(!($CheckPrimaryInDB.Country) -and $(!($ADUser.Country))){write-log -message  "Both Country values are null"}
                    ElseIf(!($CheckPrimaryInDB.Country))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "Country change $($ADUser.Country)"
                        write-log -message  "The Country value dont match so updating database with $($ADUser.Country)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Country='$($ADUser.Country)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    ElseIf($($CheckPrimaryInDB.Country) -ne $ADUser.Country)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "Country change $($ADUser.Country)"
                        write-log -message  "The Country value dont match so updating database with $($ADUser.Country)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Country='$($ADUser.Country)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''" )'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    
                    write-log -message  "Checking if their Title in the database $($CheckPrimaryInDB.Title) matchs the AD value $($ADUser.Title)"
                    # First of 2 ways of seeing if both Title fields are null
                    If(!($CheckPrimaryInDB.Title) -and $(!($ADUser.Title))){write-log -message  "Both Title values are null"}
                    ElseIf(!($CheckPrimaryInDB.Country))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "Title change $($ADUser.Title)"
                        write-log -message  "The Country value dont match so updating database with $($ADUser.Title)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Title='$($ADUser.Title)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    ElseIf($($CheckPrimaryInDB.Title) -ne $ADUser.Title)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "Title change $($ADUser.Title)"
                        write-log -message  "The Country value dont match so updating database with $($ADUser.Title)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Title='$($ADUser.Title)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''" )'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }

                    write-log -message  "Checking if their StreetAddress in the database $($CheckPrimaryInDB.StreetAddress) matchs the AD value $($ADUser.StreetAddress)"
                    # First of 2 ways of seeing if both StreetAddress fields are null
                    If(!($CheckPrimaryInDB.StreetAddress) -and $(!($ADUser.StreetAddress))){write-log -message  "Both StreetAddress values are null"}
                    ElseIf(!($CheckPrimaryInDB.Country))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "StreetAddress change $($ADUser.StreetAddress)"
                        write-log -message  "The Country value dont match so updating database with $($ADUser.StreetAddress)"
                        $SQLQuery = "UPDATE $SetSQLTable SET StreetAddress='$($ADUser.StreetAddress)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    ElseIf($($CheckPrimaryInDB.StreetAddress) -ne $ADUser.StreetAddress)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "StreetAddress change $($ADUser.StreetAddress)"
                        write-log -message  "The Country value dont match so updating database with $($ADUser.StreetAddress)"
                        $SQLQuery = "UPDATE $SetSQLTable SET StreetAddress='$($ADUser.StreetAddress)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''" )'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    

                    write-log -message  "Checking if their PostalCode in the database $($CheckPrimaryInDB.PostalCode) matchs the AD value $($ADUser.PostalCode)"
                    # First of 2 ways of seeing if both PostalCode fields are null
                    If(!($CheckPrimaryInDB.PostalCode) -and $(!($ADUser.PostalCode))){write-log -message  "Both PostalCode values are null"}
                    ElseIf(!($CheckPrimaryInDB.Country))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "PostalCode change $($ADUser.PostalCode)"
                        write-log -message  "The Country value dont match so updating database with $($ADUser.PostalCode)"
                        $SQLQuery = "UPDATE $SetSQLTable SET PostalCode='$($ADUser.PostalCode)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    ElseIf($($CheckPrimaryInDB.PostalCode) -ne $ADUser.PostalCode)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "PostalCode change $($ADUser.PostalCode)"
                        write-log -message  "The Country value dont match so updating database with $($ADUser.PostalCode)"
                        $SQLQuery = "UPDATE $SetSQLTable SET PostalCode='$($ADUser.PostalCode)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''" )'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    
                    write-log -message  "Checking if their City in the database $($CheckPrimaryInDB.City) matchs the AD value $($ADUser.City)"
                    # First of 2 ways of seeing if both City fields are null
                    If(!($CheckPrimaryInDB.City) -and $(!($ADUser.City))){write-log -message  "Both City values are null"}
                    ElseIf(!($CheckPrimaryInDB.Country))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "City change $($ADUser.City)"
                        write-log -message  "The Country value dont match so updating database with $($ADUser.City)"
                        $SQLQuery = "UPDATE $SetSQLTable SET City='$($ADUser.City)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    ElseIf($($CheckPrimaryInDB.City) -ne $ADUser.City)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "City change $($ADUser.City)"
                        write-log -message  "The Country value dont match so updating database with $($ADUser.City)"
                        $SQLQuery = "UPDATE $SetSQLTable SET City='$($ADUser.City)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''" )'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
   
                    write-log -message  "Checking if their State in the database $($CheckPrimaryInDB.State) matchs the AD value $($ADUser.State)"
                    # First of 2 ways of seeing if both State fields are null
                    If(!($CheckPrimaryInDB.State) -and $(!($ADUser.State))){write-log -message  "Both State values are null"}
                    ElseIf(!($CheckPrimaryInDB.Country))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "State change $($ADUser.State)"
                        write-log -message  "The Country value dont match so updating database with $($ADUser.State)"
                        $SQLQuery = "UPDATE $SetSQLTable SET State='$($ADUser.State)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
                    ElseIf($($CheckPrimaryInDB.State) -ne $ADUser.State)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "State change $($ADUser.State)"
                        write-log -message  "The Country value dont match so updating database with $($ADUser.State)"
                        $SQLQuery = "UPDATE $SetSQLTable SET State='$($ADUser.State)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''" )'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        }
   
                    
                    
                    write-log -message  "Finished checking $($ADUser.PrimarySMTPAddress) so will move on to next user"
                    }
                }
       }

    }

    

# We will now compate the users in the $UserArrays array to what is in the database and delete from the database is they are
write-host -ForegroundColor Green  "Comparing database data to what the users found in AD  we will delete the ones not found"
write-log -message  "Comparing database data to what the users found in AD  we will delete the ones not found"
# First of lets grab the data from SQL again
#$DatabaseData = Read-SqlTableData -serverInstance $SQLServer -database $SQLDatabase   -TableName $SQLTable  -SchemaName dbo
$SQLQuery = "SELECT * FROM $($SetSQLTable) WHERE Instance = '$($Instance)'"
$DatabaseData = Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLDatabase -Query $SQLQuery 

# Get all the users from SQL
$SQLQuery = "SELECT * FROM $($SetSQLTable)"
$AllDatabaseData = Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLDatabase -Query $SQLQuery 
$ALLADUsers = get-aduser -filter {emailaddress -like "*@*"}  -properties * | ? {$_.msExchHideFromAddressLists -ne "True"}
# Now we will loop throught each user in the table
ForEach($DBUser in $DatabaseData)
    {
    #Do garbage collection every couple of minutes to stop memory going off piste
        # so check if it is around 2 mins
        if ( $Sw.Elapsed.minutes -eq 2) {
            # it is over 2 mins so start garbage collection
            Write-Log -Message "Doing Garbage Collection after $($Sw.Elapsed.minutes ) minutes"  #write-host 
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers();
            #Reset timer by stopping and starting a new one
            $Sw.Stop()
            $sw = [diagnostics.stopwatch]::StartNew()

        }
    write-log -message  "Checking if $($DBUser.PrimarySMTPAddress) exists in the AD Array"
    # For each of these we will see if they are in the $usersArray 
    # here we look to see if the database users email address exists in the $usersArray
    $CheckUserInDB  = $UsersArray  | ? { $_.PrimarySMTPAddress -eq $DBUser.PrimarySMTPAddress} 
    # they arent so we will mark the row by setting the RecordDeleted to True
    # this will be used in the next script to purge it from teh database after the contact is deleted
    If(!($CheckUserInDB))
        {
        write-log -Message "$($DBUser.PrimarySMTPAddress) doesnt exists in the instance AD Array so we check if it exists elsehwere"
        #$DBUser.PrimarySMTPAddress
        $CheckUserInAD  = $ALLADUsers  | ? { $_.mail -eq $($DBUser.PrimarySMTPAddress.Trim())} 
        #$CheckUserInAD.Mail
        If(!($CheckUserInAD))
                {
                write-log -message  "$($DBUser.PrimarySMTPAddress) doesnt exists in the AD Array so we will mark if for deletion"
                #Here is the bit where we flag deleted users
                write-host -ForegroundColor yellow "here"
                write-host -ForegroundColor Red   "$($DBUser.PrimarySMTPAddress.Trim()) not found"
                #$SQLQuery = "UPDATE $SetSQLTable SET RecordDeleted='True' WHERE PrimarySMTPAddress= '$($DBUser.PrimarySMTPAddress -replace "'","''")'"
                #Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase  -Query $SQLQuery
                write-log -message  "$($DBUser.PrimarySMTPAddress) doesnt exists in the AD Array so we will delete it"
                write-host  "$($DBUser.PrimarySMTPAddress) doesnt exists in the AD Array so we will delete it"
                $SQLDelete = "DELETE FROM $SetSQLTable WHERE PrimarySMTPAddress='$($DBUser.PrimarySMTPAddress -replace "'","''")'"
                Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase  -Query $SQLDelete
                }
            Else
                {
                #write-host  "$($DBUser.PrimarySMTPAddress) exists in the AD Array so we will flag it as moved"
                write-host -ForegroundColor Green  "$($DBUser.PrimarySMTPAddress.Trim()) found"
                $SQLQuery = "UPDATE $SetSQLTable SET Instance='Moved' WHERE PrimarySMTPAddress= '$($DBUser.PrimarySMTPAddress -replace "'","''")'"
                Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase  -Query $SQLQuery
                }
        }
                
    Else
        {
        # Here we flag users as not deleted
        write-log -message  "$($DBUser.PrimarySMTPAddress)  exists in the AD Array so we will mark it so it isnt flagged for deletion" 
        $SQLQuery = "UPDATE $SetSQLTable SET RecordDeleted='False' WHERE PrimarySMTPAddress= '$($DBUser.PrimarySMTPAddress -replace "'","''")'"
        Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase  -Query $SQLQuery
        }
    }

write-log -message  "Script has completed" 
# Last thing - stop transcript
Stop-Transcript
