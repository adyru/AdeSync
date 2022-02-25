
<#
This script is to update an SQL database with AD Objects
This data will then be used in a subsequent script to create contacts in another directory

Things that need to be changed are ...

For Different domains
$DC
$InstanceTargetAddress
$SQLServer
$SQLDatabase
$SQLTable
$SetSQLTable

For a New Instance in the same domain
$Instance
it 'should' create the new directories for the logs

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
    $script:ErrorLog  = $logfile -replace ".txt", "-errors.txt"
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
                $newOut = $outfile -replace ".csv", "--Random-$($Random).csv"
                move-item $outfile $newOut
                $OutVar | Export-Csv -NoClobber -NoTypeInformation -path $outfile -Encoding UTF8
                $OutcheckFile = test-path $outfile
                If ($OutcheckFile -eq $true)
                { write-host -ForegroundColor Green "Variable exported to $($outfile)" }
                Else
                { write-host -ForegroundColor red "Variable not exported to $($outfile)" }
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



Function Add-Attachments {
     [CmdletBinding()]
        
        Param(
        [Parameter(Mandatory=$true)]
        [string]
        $File
        )

        $Check = test-path $file
        If($Check -eq $true)
            {
            $script:attachments += $file
            }
        Else
            {
            #write-host -ForegroundColor red "File $($attachments) doesnt exist"
            }

}
$attachments = @()

######################
# Start of Variables #
######################

##################################
#First of logging and transcript #
##################################

###################################
# Need to be amended per instance #
###################################

#Instance No
$Instance = "1"


#Date that we will look for new/update users from
$Currentdate = get-date -Format dd-MM-yyyy--hh-mm


#Create new folder for output files if it doesnt exist using todays date
#Path
$Path = "D:\LogFiles\Sync"
$FolderDate = Get-Date -Format dd-MM-yyyy
#$newFolder = "C:\Scripts\RPA\Output\$($FolderDate)\Update-Database-To-O365\Instance$($instance)"
$newFolder = "$Path\$($FolderDate)\Update-Database-To-O365\Instance$($instance)"
$TestFolder = test-path $newFolder
If ($testfolder -eq $false) { New-Item -Path $NewFolder -ItemType directory }
$OutputFolder = $NewFolder


# Create event log
New-EventLog -LogName application -Source "Sync - Database Update" -ErrorAction silentlycontinue -WarningAction SilentlyContinue
Write-EventLog -LogName "Application" -Source "Sync - Database Update"  -EventID 1 -EntryType Information   -Message "Instance $($Instance) started"


#Need to create a logfile for the write-log function as the script will error out without that
$logfile = "$OutputFolder\Sync$($instance)-Update-Database-To-O36-Log-$($Currentdate).txt"
$transcript = "$OutputFolder\Sync$($instance)-Update-Database-To-O365-transcript-$($Currentdate).txt"
start-transcript $transcript 



#################################
# Need to be amended per domain #
#################################

# Static DC to use - needs to be FQDN this will be checked later on
$DC = "DC.company.co.uk"

#Exchange Server
$Exchange = "exchange.company.co.uk"

#Instance TargetAddress
$InstanceTargetAddress = "@baddress.co.uk"

### Load Modules ###
Write-Log -Message "Importing the sqlserver PS module)" 
Import-Module "D:\PowershellModules\sqlserver\sqlserver"
Write-Log -Message "Importing the Exchange PS module)"  
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$($Exchange)/PowerShell/ -Authentication Kerberos
Import-PSSession $Session

#SQL Variables
$SQLServer = "SQL.company.co.uk"
$SQLDatabase = "Database"
$SQLTable = "LiveTable"
$SetSQLTable = "dbo.LiveTable"
$DatabaseData = Read-SqlTableData -serverInstance $SQLServer -database $SQLDatabase   -TableName $SQLTable  -SchemaName dbo
$sqlUpdateDate = (get-date).ToString('yyyy-MM-dd')


# Domain Variable
$NetBios = (Get-ADDomain).NetBIOSName

#Production Vars for mail send at the end
$ReportsDL = "someone@somewhere.com"
$ManagmentReportsDL = "somedl@somewhere.com"
$Admin = "someone@somewhere.com"
$Sender = "someone@somewhere.com"
$SmtpServer = "smtp.domain.co.uk" # + $ADDomain.DNSRoot
$subject = "o365 Contact Creation - $NetBios Instance$($Instance)"
$CC = "someoneelse@somewhere.com"

# AD Attibutes we are interested In
$attributes = "proxyAddresses","emailaddress","Mail","GivenName","Surname","telephoneNumber","Mobile","office","Company","Department","Title","Country","StreetAddress","postalcode","city","State","msExchHideFromAddressLists","Enabled"

# Hold the mailbox databases at the end
$ALLADUsers =  @()
$AllExchangeDatabases = (Get-MailboxDatabase | ?{$_.name -notlike "jdb*" -and $_.name -notlike "qcsdb*"}).name | sort name
$ALLADUsersCount = 0

######################
# Standard Variables #
######################

#Keep track of users processed by this script
$UsersArray =  @()
$UsersArrayCounter = 0


# List Of OUs to process
Write-Log -Message "Checking OU Input script is available"
$CheckFile = test-path "$($PSScriptRoot)\Allowed-OUs.txt"
If(!($CheckFile ))
    {
    Write-Log -Message "OU file not found - exiting" -Level ERROR
    Write-EventLog -LogName "Application" -Source "Company Sync - Database Update"  -EventID 104 -EntryType Error -Message "OU import file not found so exiting for instance $($Instance) "
    $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as OU Input not found"
    send-mailmessage -to $ManagmentReportsDL -from $Emailsender -bcc $CC -subject $subject  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    Stop-Transcript
    exit
    }
Else
    {
    Write-Log -Message "OU file found so importing"
    Write-EventLog -LogName "Application" -Source "Company Sync - Database Update"  -EventID 114 -EntryType Information  -Message "OU import file found for instance $($Instance)"
    $OUSImport = get-content "$($PSScriptRoot)\Allowed-OUs.txt"
    $OUs = $OUSImport |select -unique
    }


# Sync Exclusion Group
$ExcludeUsers = Get-ADGroupMember GG-CompanyAfrica-Sync-Exclude | select samaccountname

#Reporting variables
$DeletedUserReport = @()
$UpdatedUserReport = @()
$NewUserReport = @()
$MovedUserIntoOUReport = @()
$MovedUserOutOUReport = @()

#Outfiles for extracting the above arrays
$NewUserReportOut = "$($OutputFolder)\NewUser-Report-Instance$($Instance)-$($Currentdate).csv"
$UpdatedUserReportOut = "$($OutputFolder)\UpdatedUser-Report-Instance$($Instance)-$($Currentdate).csv"
$DeletedUserReportOut = "$($OutputFolder)\Deleted-Report-Instance$($Instance)-$($Currentdate).csv"
$MovedUserIntoOUReportOut = "$($OutputFolder)\Moved-Into-Instance-Report-Instance$($Instance)-$($Currentdate).csv"
$MovedUserOutOUReportOut = "$($OutputFolder)\Moved-Out-Of-Instance-Report-Instance$($Instance)-$($Currentdate).csv"

#Counters
$NewUserReportCount  = 0
$UpdatedUserReportCount = 0
$DeletedUserReportCount = 0
$MovedUserReportIntoOUCount = 0
$MovedUserReportOutOUCount = 0

# We are going to do a garbage collection every 2 mins so 
# need to kick of a timer
#Start Stopwatch
$sw = [diagnostics.stopwatch]::StartNew()

########################
#  End Of Variables    #
########################

###Revisit

########################
# Connectivity Checks  #
########################

Write-Log -Message "### Starting the connectivity tests ###"

Write-Log -Message "Setting Domain Controller to  $($DC )" #write-host
Write-Log -Message "Checking connectivity to Domain Controller $($DC)" #write-host
# Check using the nettestserver function 
$CheckCOnnection = Nettestserver $DC 135

# Check what the status of the connection is
If ($CheckCOnnection.Connected -eq $false) {
    # It failed - report and exit script
    write-host -ForegroundColor red "Connection is $($CheckCOnnection.Connected) so exiting script, amend DC variable and re run" 
    Write-Log -Message "Connectivity to Domain Controller $($DC ) is $($CheckCOnnection.Connected ) so exiting script " #write-host -Level "Fatal"
    Write-EventLog -LogName "Application" -Source "Company Sync - Database Update"  -EventID 106 -EntryType Error -Message "Failed to connect to Domain Controller $($DC) so exiting for instance $($Instance)"
    $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as the domain controller check failed"
    send-mailmessage -to $ManagmentReportsDL -from $Emailsender -bcc $CC -subject $subject  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    Stop-Transcript
    exit
}
Else {
    #All good so carry on and report to screen
    write-host -ForegroundColor Green "Connectivity to Domain Controller $($DC ) Succdded - check connection is  $($CheckCOnnection.Connected) "
    Write-Log -Message "Connectivity to Domain Controller $($DC ) Succdded - check connection is  $($CheckCOnnection.Connected ) " #write-host
    Write-EventLog -LogName "Application" -Source "Company Sync - Database Update"  -EventID 116 -EntryType Information -Message "Connected to Domain Controller $($DC) for instance $($Instance)"
}

# Check SQL Module
Write-Log -Message "Checking SQL PS Module is loaded"
$SQLCheck = get-module sqlserver
IF(!($SQLCheck ))
    {
    Write-Log -Message "SQL Powershell Module Check failed - so exiting" -Level "Error"
    Write-EventLog -LogName "Application" -Source "Company Sync - Database Update"  -EventID 101 -EntryType Error -Message "SQL Powershell Module Check failed for instance $($Instance) "
    $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as the SQL PS module check failed"
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject "Script Failed To Run"  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    Stop-Transcript
    exit
    }
Else
    {
    Write-Log -Message "SQL Powershell Module loaded"
    Write-EventLog -LogName "Application" -Source "Company Sync - Database Update"  -EventID 111 -EntryType Information -Message "SQL Powershell Module Check Succedded for instance $($Instance)"
    }

Write-Log -Message "Checking the connection to the SQL database"
#Check SQL Connection 
$conn = New-Object System.Data.SqlClient.SqlConnection                                      
$conn.ConnectionString = "Server=$SQLServer;Database=$Database;Integrated Security=True;"                                                                        
$conn.Open()
IF($conn.State -ne "Open")
    {
    Write-Log -Message "Failed to connect to database $($Database) on server $($SQLServer) so exiting" -Level "Error"
    Write-EventLog -LogName "Application" -Source "Company Sync - Database Update"  -EventID 102 -EntryType Error -Message "Failed to connect to database $($Database) on server $($SQLServer) so exiting for instance $($Instance)"
    $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as the SQL check failed"
    send-mailmessage -to $ManagmentReportsDL -from $Emailsender -bcc $CC -subject $subject  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    Stop-Transcript
    exit
    }
Else
    {
    Write-Log -Message "Connected to database $($Database) on server $($SQLServer)"
    write-EventLog -LogName "Application" -Source "Company Sync - Database Update"  -EventID 112 -EntryType Information -Message "Connected to database $($Database) on server $($SQLServer) for instance $($Instance)"
    }

Write-Log -Message "Checking Exchange Powershell is loaded"
$ExchangeCheck = get-command get-mailcontact
IF(!($ExchangeCheck))
    {
    Write-Log -Message "Failed the exchange powershell command test" -Level "Error"
    Write-EventLog -LogName "Application" -Source "Company Sync - Database Update"  -EventID 103 -EntryType Error -Message "Exchange Powershell module failed against server $($Exchange) so exiting for instance $($Instance)"
    $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as the Exchange check failed"
    send-mailmessage -to $ManagmentReportsDL -from $Emailsender -bcc $CC -subject $subject  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    Stop-Transcript
    exit
    }
Else
    {
    Write-Log -Message "Exchange powershell command tests passed"
    Write-EventLog -LogName "Application" -Source "Company Sync - Database Update"  -EventID 113 -EntryType Information -Message "Exchange Powershell module succedded against server $($Exchange) for instance $($Instance)"
    }

Write-Log -Message "### End of  the connectivity tests ###"
# We will loop though the OUs in the text to pick up users
write-log -Message "Picking up OUs from text file" 
ForEach($OU in $OUs)
    {
    $ADUserRunningCount = 0
    Write-host "Starting in $($OU)" 
    Write-Log -Message "Starting in $($OU)" 
    # We will go and find users if they have an email address and are not hidden from the address list 
        # Trying to use a variable to select these
    Write-Log -Message "We are going to get the AD users from $($OU)"    
    $ADUsers = get-aduser -filter {emailaddress -like "*@*" -and enabled -eq $true} -SearchBase $OU -ResultSetSize $Null -properties $($attributes) -SearchScope 1   | Where-Object {$_.msExchHideFromAddressLists -ne "True"} 
            # -properties Mail,GivenName,Surname,telephoneNumber,Mobile,office,Company,Department,Title,Country,StreetAddress,postalcode,city,State `
            #-properties $($attributes)   | Where-Object {$_.msExchHideFromAddressLists -ne "True"} #| select -first 2
    
    #$ADUsers = get-aduser -filter {emailaddress -like "*@*" -and msExchHideFromAddressLists -ne "True"} -SearchBase $OU -properties * | ? {$_.msExchHideFromAddressLists -ne "True"} #| select -first 2
    #$ADUsers =get-aduser -filter {emailaddress -like "*@*" -and msExchHideFromAddressLists -ne "True"} -SearchBase $OU -properties *
        If(!($ADUsers))
            {
            Write-Log -Message "No users found in $($OU)" -Level Error
            }
        Else
            {
            Write-Log -Message "$(($ADUsers | Measure-Object).count) found in AD" 
            }  


    #First Of Update Instance for users who have moved OU
    Write-Log -Message "Checking database for users flagged as moved" 
    $SQLQuery = "SELECT *  FROM $($SetSQLTable) WHERE Instance = 'Moved'"
    $movedContacts = Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLDatabase -Query $SQLQuery
    If(!($movedContacts))
        {
        Write-Log -Message "No users flagged in the database as moved" 
        }
    Else
        {
        Write-Log -Message "$(($movedContacts | Measure-Object).count) moved users found in SQL Database" 
        }  
    #$movedContacts 
    #start-sleep -Seconds 2 
    write-log -message  "Starting to process the moved users"
    ForEach($movedContact in $movedContacts)
        {
        write-host "Processing moved user address $($movedContact.PrimarySMTPAddress)"
        ### Sunday ###
        #$MoveCheckAddress = $MailCheck = $movedContact.PrimarySMTPAddress -replace "'","''"
        $MoveCheckAddress = $movedContact.PrimarySMTPAddress -replace "'","''"
        ### Changed this ##
        #$CheckUserInAD = get-aduser -filter "emailaddress -like '$($MoveCheckAddress)'"
        write-host "Checking if $($movedContact.PrimarySMTPAddress) exists and isnt disabled"
        $CheckUserInAD = get-aduser -filter "emailaddress -like '$($MoveCheckAddress)' -and enabled -eq '$true'"
        if(!($CheckUserInAD))
            {
            write-host "$($movedContact.PrimarySMTPAddress)  doesnt exists in AD - so flag for deletion"
            write-log -message  "$($movedContact.PrimarySMTPAddress)  doesnt exists in AD - so flag for deletion"
            #$SQLDelete = "DELETE FROM $SetSQLTable WHERE PrimarySMTPAddress='$($movedContact.PrimarySMTPAddress -replace "'","''")'"
            $CheckUserinDeletedArray = $DeletedUserReport | Where-Object { $_.user -eq $($movedContact.PrimarySMTPAddress)}
                if(!($CheckUserinDeletedArray))
                    {
                    write-log -message  "$($movedContact.PrimarySMTPAddress)  not in the deleted array so adding"
                    $SQLQuery = "UPDATE $SetSQLTable SET RecordDeleted='True' WHERE PrimarySMTPAddress = '$($movedContact.PrimarySMTPAddress -replace "'","''")'" 
                    Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase  -Query $SQLQuery
                    $DeletedUserReportObj  = New-Object System.Object
                    $DeletedUserReportObj | Add-Member -type NoteProperty -name User -Value $movedContact.PrimarySMTPAddress
                    $DeletedUserReport += $DeletedUserReportObj
                    $DeletedUserReportCount++
                    }
                Else
                    {
                    write-log -message  "$($movedContact.PrimarySMTPAddress)  is in the deleted array so ignoring"
                    }

            }
        Else
            {
            $CheckUserInOU  = $ADUsers | Where-Object { $_.mail -eq $($movedContact.PrimarySMTPAddress)} 
            If(!($CheckUserInOU))
                {
                write-host "here - exists in AD but not in this OU"
                write-log -message  "$($movedContact.PrimarySMTPAddress) exists in AD but not in this OU"
                }
            Else
                {
                write-log -message  "$($movedContact.PrimarySMTPAddress) found in $($OU) so updating database to point this user at instance $($Instance)"
                #$SQLQuery = "UPDATE $SetSQLTable SET Instance=$($Instance) WHERE PrimarySMTPAddress= '$($movedContact.PrimarySMTPAddress -replace "'","''")'"
                $SQLQuery = "UPDATE $SetSQLTable SET Instance=$($Instance),RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($movedContact.PrimarySMTPAddress -replace "'","''")'"
                Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase  -Query $SQLQuery
                $MovedUserIntoOUReportObj  = New-Object System.Object
                $MovedUserIntoOUReportObj | Add-Member -type NoteProperty -name User -Value $movedContact.PrimarySMTPAddress
                $MovedUserIntoOUReport += $MovedUserIntoOUReportObj
                $MovedUserIntoOUReportCount++
                }

  
            }
 
        }


    # We will tehn loop through these users
    Write-Log -Message "Starting to look through the $(($ADUsers | Measure-Object).count) found in AD" 
    ForEach($ADUser in $ADUsers)
        {
        $ADUserRunningCount++
        $TargetAddress = $null
        #Do garbage collection every couple of minutes to stop memory going off piste
        # so check if it is around 2 mins
            if ( $Sw.Elapsed.minutes -eq 2) {
                # it is over 2 mins so start garbage collection
                Write-Log -Message "### Doing Garbage Collection after $($Sw.Elapsed.minutes ) minutes ###"
                Write-Log -Message "### Processing $($ADUserRunningCount) of $(($ADUsers | Measure-Object).count) users ###"  -Level Error
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers();
                #Reset timer by stopping and starting a new one
                $Sw.Stop()
                $sw = [diagnostics.stopwatch]::StartNew()

            }

        ### Friday change ###
        # Check if they are in exclude group
        $CheckExcludeUser = $ExcludeUsers| Where-Object { $_.samaccountname -eq $ADUser.samaccountname}
        
        # First of check if they are a exchange recipient 
        Write-Log -Message "Checking if $($ADUser.emailaddress) is an exchange recipient" 
        #$checkifMB = get-recipient $ADUser.emailaddress -erroraction silentlycontinue
        $MailCheck = $ADUser.emailaddress -replace "'","''"
        #write-host "here $($MailCheck )"
        $checkifMB = get-mailbox -Filter "primarysmtpaddress -eq '$MailCheck'" -erroraction silentlycontinue
        #write-host "now here $($checkifMB).alias"
        # if they arent a exchagne recipient
        If(!($checkifMB))
            {
            # we will log it and move on to the next user
            write-log -message  "$($ADUser.emailaddress) isnt an mail enabled object" -Level Error
            }
        # if they are an exchange user we will carry on
        Else
            {
            #write-host "now here"
            write-log -message  "$($ADUser.emailaddress) is an mail enabled object"
        # A counter i might use later on
        $UsersArrayCounter++
        # We will now look at all the users email addresses as we need to pick a target address
        write-log -message  "Checking the email addresses on $($ADUser.emailaddress)"
        $TargetAddressSearch = "smtp:$($ADuser.Samaccountname)$($InstanceTargetAddress)"
        write-log -message  "Looking for an address that matches smtp:$($ADuser.Samaccountname)$($InstanceTargetAddress) on user $($ADuser.Samaccountname)"
        ForEach($EmailAddress in $ADuser.proxyAddresses)
            {
            # Here we look to get an email address that matches the variable $TargetAddressSearch
            If($EmailAddress -like $TargetAddressSearch) 
                {
                $TargetAddress = $EmailAddress -replace "smtp:",""
                write-log -message  "Match found for 4x4 email address smtp:$($ADuser.Samaccountname)$($InstanceTargetAddress)"
                #Write-host -ForegroundColor GREEN "Here and TargetAddress id $($TargetAddress)"
                }
            }
            #if they dont have a match for this we will go through and just find a match for any email address in teh domain $TargetAddressSearch
        if(!($TargetAddress))
                {
                write-log -message  "There were no matches for smtp:$($ADuser.Samaccountname)$($InstanceTargetAddress) so we are now looking for smtp:*$($InstanceTargetAddress)"
                $TargetAddressSearch = "smtp:*$($InstanceTargetAddress)"
                ForEach($EmailAddress in $ADuser.proxyAddresses)
                    {
                    # Here we look to get an email address that matches the variable $TargetAddressSearch
                    If($EmailAddress -like "$($TargetAddressSearch)") 
                        {
                        $TargetAddress = $EmailAddress -replace "smtp:",""
                        #Write-host -ForegroundColor GREEN "Here and TargetAddress id $($TargetAddress)"
                        }
                    }
                }
        if(!($TargetAddress))
                {
                write-log -message  "Cannot find an email address with domain name $($InstanceTargetAddress) on user $($ADUser.Mail)" -Level "Error"
                }
            Else
                {
                write-log -message  "We have picked $($TargetAddress) as the target address so all good"
                }
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
            $CheckPrimaryInDB  = $DatabaseData | Where-Object { $_.PrimarySMTPAddress.Trim() -eq $ADUser.Mail}
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
         if(!($TargetAddress))
            {
            write-log -message  "Cannot find an email address with domain name $($InstanceTargetAddress) on user $($ADUser.Mail) - so not adding to database)" -Level "Error"
            }
        ### Friday Change ###
        ElseIf($CheckExcludeUSer -ne $null)
            {
            Write-Log -Message "### $($ADUser.samaccountname) is in the exclude group so not adding to the database ###"  -Level Error
            }
        Else
            {
            write-log -message  "Adding $($ADUser.Mail) to database)" 
            Write-SqlTableData -serverInstance $SQLServer -database $SQLDatabase  -TableName $SQLTable -SchemaName dbo -InputData $Obj
            $NewUserReportObj  = New-Object System.Object
            $NewUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
            $NewUserReport += $NewUserReportObj
            $NewUserReportCount++
            }
         
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
                        write-host "first name change $($ADUser.GivenName)"
                        write-log -message  "The first name values dont match so updating database with $($ADUser.GivenName)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Firstname='$($ADUser.GivenName -replace "'","''")',NameUpdated='1',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value Firstname
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.GivenName
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }

                    # check to see if the Surname  is the same as in AD
                    # First of 2 ways of seeing if both are null
                    write-log -message  "Checking if their surname name in the database $($CheckPrimaryInDB.Surname) matchs the AD value $($ADUser.Surname)"
                    If(($null -eq $CheckPrimaryInDB.Surname) -and $(!($ADUser.Surname))){write-log -message  "Both surname values are null"}
                    If(!($CheckPrimaryInDB.Surname) -and $(!($ADUser.Surname))){write-log -message  "Both surname values are null"}
                    # It isnt null in both
                    ElseIf($null -eq $CheckPrimaryInDB.Surname)
                        {
                        write-log -message  "The surnames value dont match so updating database with $($ADUser.Surname)"
                        # It is null in the database so we will update with the AD value
                        write-host "Surname change $($ADUser.Surname)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Surname='$($ADUser.Surname -replace "'","''")',NameUpdated='1',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value Surname
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.Surname
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }
                    ElseIf($($CheckPrimaryInDB.Surname) -ne $ADUser.Surname)
                        {
                        write-log -message  "The surnames value dont match so updating database with $($ADUser.Surname)"
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "Surname change $($ADUser.Surname)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Surname='$($ADUser.Surname -replace "'","''" )',NameUpdated='1',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value Surname
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.Surname
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }
                    
                    #We will assume the target address cannot be null and it needs to be se
                    write-log -message  "Checking if their target address in the database $($CheckPrimaryInDB.TargetAddress) matchs the AD value $($TargetAddress)"
                    If($($CheckPrimaryInDB.TargetAddress) -ne $TargetAddress)
                        {
                        # It is a different value so we will update with the AD value
                        write-host "TargetAddress change $($TargetAddress)"
                        write-log -message  "The TargetAddresses value dont match so updating database with $($TargetAddress)"
                        $SQLQuery = "UPDATE $SetSQLTable SET TargetAddress='$($TargetAddress -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                                 if(!($TargetAddress))
                                        {
                                        write-log -message  "Cannot find an email address with domain name $($InstanceTargetAddress) on user $($ADUser.Mail) - so not updating to database" -Level "Error"
                                        }
                                    Else
                                        {
                                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                                        $UpdatedUserReportObj  = New-Object System.Object
                                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value TargetAddress
                                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $TargetAddress
                                        $UpdatedUserReport+= $UpdatedUserReportObj
                                        $UpdatedUserReportCount++
                                        }
                        }

                    # First of 2 ways of seeing if both Mobile fields are null
                    write-log -message  "Checking if their Mobile in the database $($CheckPrimaryInDB.Mobile) matchs the AD value $($ADUser.Mobile)"
                    If(($null -eq $CheckPrimaryInDB.Mobile) -and $(!($ADUser.Mobile))){write-log -message  "Both Mobile values are null"}
                    ElseIf(!($CheckPrimaryInDB.Mobile ) -and $(!($ADUser.Mobile))){write-log -message  "Both surname Mobile are null"}
                    ElseIf($null -eq $CheckPrimaryInDB.Mobile)
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "Mobile change $($ADUser.Mobile)"
                        write-log -message  "The Mobiles value dont match so updating database with $($ADUser.Mobile)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Mobile='$($ADUser.Mobile)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value Mobile
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.Mobile
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }
                    ElseIf($($CheckPrimaryInDB.Mobile) -ne $ADUser.Mobile)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "Mobile change $($ADUser.Mobile)"
                        write-log -message  "The Mobiles value dont match so updating database with $($ADUser.Mobile)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Mobile='$($ADUser.Mobile)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value Mobile
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.Mobile
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
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
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value telephoneNumber
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.telephoneNumber
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }
                    ElseIf($($CheckPrimaryInDB.telephoneNumber) -ne $ADUser.telephoneNumber)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "telephoneNumber change $($ADUser.telephoneNumber)"
                        write-log -message  "The TelephoneNumbers value dont match so updating database with $($ADUser.telephoneNumber)"
                        $SQLQuery = "UPDATE $SetSQLTable SET telephoneNumber='$($ADUser.telephoneNumber)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value telephoneNumber
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.telephoneNumber
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }
                    
                    # First of 2 ways of seeing if both Office fields are null
                    write-log -message  "Checking if their Office in the database $($CheckPrimaryInDB.Office) matchs the AD value $($ADUser.Office)"
                    If(!($CheckPrimaryInDB.Office) -and $(!($ADUser.Office))){write-log -message  "Both Office values are null"}
                    ElseIf(!($CheckPrimaryInDB.Office))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "Office change $($ADUser.Office)"
                        write-log -message  "The Office value dont match so updating database with $($ADUser.Office)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Office='$($ADUser.Office -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value Office
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.Office
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }
                    ElseIf($($CheckPrimaryInDB.Office) -ne $ADUser.Office)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "Office change $($ADUser.Office)"
                        write-log -message  "The Office value dont match so updating database with $($ADUser.Office)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Office='$($ADUser.Office -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value Office
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.Office
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }
                    
                     write-log -message  "Checking if their Department in the database $($CheckPrimaryInDB.Department) matchs the AD value $($ADUser.Department)"
                    # First of 2 ways of seeing if both Department fields are null
                    If(!($CheckPrimaryInDB.Department) -and $(!($ADUser.Department))){write-log -message  "Both Department values are null"}
                    ElseIf(!($CheckPrimaryInDB.Department))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "Department change $($ADUser.Department)"
                        write-log -message  "The Office value dont match so updating database with $($ADUser.Department)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Department='$($ADUser.Department -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value Department
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.Department
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }
                    ElseIf($($CheckPrimaryInDB.Department) -ne $ADUser.Department)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "Department change $($ADUser.Department)"
                        write-log -message  "The Office value dont match so updating database with $($ADUser.Department)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Department='$($ADUser.Department -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value Department
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.Department
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }

                    write-log -message  "Checking if their Company in the database $($CheckPrimaryInDB.Company) matchs the AD value $($ADUser.Company)"
                    # First of 2 ways of seeing if both Company fields are null
                    If(!($CheckPrimaryInDB.Company) -and $(!($ADUser.Company))){write-log -message  "Both Company values are null"}
                    ElseIf(!($CheckPrimaryInDB.Company))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "Company change $($ADUser.Company)"
                        write-log -message  "The Office value dont match so updating database with $($ADUser.Company)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Company='$($ADUser.Company -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value Company
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.Company
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++

                        }
                    ElseIf($($CheckPrimaryInDB.Company) -ne $ADUser.Company)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "Company change $($ADUser.Company)"
                        write-log -message  "The Office value dont match so updating database with $($ADUser.Company)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Company='$($ADUser.Company -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''" )'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value Company
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.Company
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }
                    
                    write-log -message  "Checking if their Country in the database $($CheckPrimaryInDB.Country) matchs the AD value $($ADUser.Country)"
                    # First of 2 ways of seeing if both Country fields are null
                    If(($null -eq $CheckPrimaryInDB.Country) -and $(!($ADUser.Country))){write-log -message  "Country value - Database = null Active Directory = Empty"}
                    ElseIf(($null -eq $CheckPrimaryInDB.Country) -and ($null -eq $ADUser.Country)){write-log -message  "Country value - Database = null Active Directory = Null"}
                    ElseIf((!($CheckPrimaryInDB.Country)) -and ($null -eq $ADUser.Country)){write-log -message  "Country value - Database = Empty Active Directory = Null"}
                    ElseIf((!($CheckPrimaryInDB.Country)) -and $(!($ADUser.Country))){write-log -message  "Country value - Database = Empty Active Directory = Emtpy"}
                    ElseIf(($null -ne $CheckPrimaryInDB.Country) -and ($null -eq $ADUser.Country)){write-log -message  "Country value - Database = Not Null  Active Directory = Null"}
                    ElseIf(!($CheckPrimaryInDB.Country))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "Country change $($ADUser.Country)"
                        write-log -message  "The Country value dont match so updating database with $($ADUser.Country)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Country='$($ADUser.Country)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value Country
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.Country
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }
                    ElseIf($($CheckPrimaryInDB.Country) -ne $ADUser.Country)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "Country change $($ADUser.Country)"
                        write-log -message  "The Country value dont match so updating database with $($ADUser.Country)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Country='$($ADUser.Country)',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''" )'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value Country
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.Country
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }
                    
                    write-log -message  "Checking if their Title in the database $($CheckPrimaryInDB.Title) matchs the AD value $($ADUser.Title)"
                    # First of 2 ways of seeing if both Title fields are null
                    If(!($CheckPrimaryInDB.Title) -and $(!($ADUser.Title))){write-log -message  "Both Title values are null"}
                    ElseIf(!($CheckPrimaryInDB.Title))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "Title change $($ADUser.Title)"
                        write-log -message  "The Title value dont match so updating database with $($ADUser.Title)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Title='$($ADUser.Title -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value Title
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.Title
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }
                    ElseIf($($CheckPrimaryInDB.Title) -ne $ADUser.Title)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "Title change $($ADUser.Title)"
                        write-log -message  "The Title value dont match so updating database with $($ADUser.Title)"
                        $SQLQuery = "UPDATE $SetSQLTable SET Title='$($ADUser.Title -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''" )'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value Title
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.Title
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }

                    write-log -message  "Checking if their StreetAddress in the database $($CheckPrimaryInDB.StreetAddress) matchs the AD value $($ADUser.StreetAddress)"
                    # First of 2 ways of seeing if both StreetAddress fields are null
                    If(!($CheckPrimaryInDB.StreetAddress) -and $(!($ADUser.StreetAddress))){write-log -message  "Both StreetAddress values are null"}
                    ElseIf(!($CheckPrimaryInDB.StreetAddress))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "StreetAddress change $($ADUser.StreetAddress)"
                        write-log -message  "The StreetAddress value dont match so updating database with $($ADUser.StreetAddress)"
                        $SQLQuery = "UPDATE $SetSQLTable SET StreetAddress='$($ADUser.StreetAddress -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value StreetAddress
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.StreetAddress
                        $UpdatedUserReport+= $UpdatedUserReportObj 
                        $UpdatedUserReportCount++
                        }
                    ElseIf($($CheckPrimaryInDB.StreetAddress) -ne $ADUser.StreetAddress)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "StreetAddress change $($ADUser.StreetAddress)"
                        write-log -message  "The StreetAddress value dont match so updating database with $($ADUser.StreetAddress)"
                        $SQLQuery = "UPDATE $SetSQLTable SET StreetAddress='$($ADUser.StreetAddress -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''" )'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value StreetAddress
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.StreetAddress
                        $UpdatedUserReport+= $UpdatedUserReportObj
                        $UpdatedUserReportCount++
                        }
                    

                    write-log -message  "Checking if their PostalCode in the database $($CheckPrimaryInDB.PostalCode) matchs the AD value $($ADUser.PostalCode)"
                    # First of 2 ways of seeing if both PostalCode fields are null
                    If(!($CheckPrimaryInDB.PostalCode) -and $(!($ADUser.PostalCode))){write-log -message  "Both PostalCode values are null"}
                    ElseIf(!($CheckPrimaryInDB.PostalCode))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "PostalCode change $($ADUser.PostalCode)"
                        write-log -message  "The PostalCode value dont match so updating database with $($ADUser.PostalCode)"
                        $SQLQuery = "UPDATE $SetSQLTable SET PostalCode='$($ADUser.PostalCode -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value PostalCode
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.PostalCode
                        $UpdatedUserReport+= $UpdatedUserReportObj 
                        $UpdatedUserReportCount++
                        }
                    ElseIf($($CheckPrimaryInDB.PostalCode) -ne $ADUser.PostalCode)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "PostalCode change $($ADUser.PostalCode)"
                        write-log -message  "The Postal value dont match so updating database with $($ADUser.PostalCode)"
                        $SQLQuery = "UPDATE $SetSQLTable SET PostalCode='$($ADUser.PostalCode -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''" )'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value PostalCode
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.PostalCode
                        $UpdatedUserReport+= $UpdatedUserReportObj 
                        $UpdatedUserReportCount++
                        }
                    
                    write-log -message  "Checking if their City in the database $($CheckPrimaryInDB.City) matchs the AD value $($ADUser.City)"
                    # First of 2 ways of seeing if both City fields are null
                    If(!($CheckPrimaryInDB.City) -and $(!($ADUser.City))){write-log -message  "Both City values are null"}
                    ElseIf(!($CheckPrimaryInDB.City))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "City change $($ADUser.City)"
                        write-log -message  "The City value dont match so updating database with $($ADUser.City)"
                        $SQLQuery = "UPDATE $SetSQLTable SET City='$($ADUser.City -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value City
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.City
                        $UpdatedUserReport+= $UpdatedUserReportObj  
                        $UpdatedUserReportCount++
                        }
                    ElseIf($($CheckPrimaryInDB.City) -ne $ADUser.City)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "City change $($ADUser.City)"
                        write-log -message  "The City value dont match so updating database with $($ADUser.City)"
                        $SQLQuery = "UPDATE $SetSQLTable SET City='$($ADUser.City -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''" )'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value City
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.City
                        $UpdatedUserReport+= $UpdatedUserReportObj 
                        $UpdatedUserReportCount++
                        }
   
                    write-log -message  "Checking if their State in the database $($CheckPrimaryInDB.State) matchs the AD value $($ADUser.State)"
                    # First of 2 ways of seeing if both State fields are null
                    If(!($CheckPrimaryInDB.State) -and $(!($ADUser.State))){write-log -message  "Both State values are null"}
                    ElseIf(!($CheckPrimaryInDB.State))
                        {
                        # It is null in the database so we will update with the AD value
                        write-host "State change $($ADUser.State)"
                        write-log -message  "The State value dont match so updating database with $($ADUser.State)"
                        $SQLQuery = "UPDATE $SetSQLTable SET State='$($ADUser.State -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value State
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.State
                        $UpdatedUserReport+= $UpdatedUserReportObj 
                        $UpdatedUserReportCount++
                        }
                    ElseIf($($CheckPrimaryInDB.State) -ne $ADUser.State)
                        {
                        # It isnt null in the database but is a different value so we will update with the AD value
                        write-host "State change $($ADUser.State)"
                        write-log -message  "The State value dont match so updating database with $($ADUser.State)"
                        $SQLQuery = "UPDATE $SetSQLTable SET State='$($ADUser.State -replace "'","''")',RecordUpdated='$($sqlUpdateDate)' WHERE PrimarySMTPAddress= '$($ADUser.Mail -replace "'","''" )'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLdatabase  -Query $SQLQuery 
                        $UpdatedUserReportObj  = New-Object System.Object
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $ADUser.Mail
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name Attribute -Value State
                        $UpdatedUserReportObj | Add-Member -type NoteProperty -name NewValue -Value $ADUser.State
                        $UpdatedUserReport+= $UpdatedUserReportObj 
                        $UpdatedUserReportCount++
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
#$SQLQuery = "SELECT * FROM $($SetSQLTable)"
#$AllDatabaseData = Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLDatabase -Query $SQLQuery 
$SQLQuery = "SELECT * FROM $($SetSQLTable) WHERE Instance = '$($Instance)'"
$AllDatabaseData = Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLDatabase -Query $SQLQuery 

#$ALLADUsers = get-aduser -filter {emailaddress -like "*@*"}  -properties * | Where-Object {$_.msExchHideFromAddressLists -ne "True"}

write-log -message  "### Getting all mailboxes from exchange ###"
#$ALLADUsers = get-mailbox -ResultSize unlimited -warningaction silentlycontinue -RecipientTypeDetails UserMailbox  | ? {$_.HiddenFromAddressListsEnabled -eq $false}

### Chnaging here ###
ForEach($AllExchangeDatabase in $AllExchangeDatabases)
    {
    write-host "Processing database $($AllExchangeDatabase)"
    write-log -message  "Processing database $($AllExchangeDatabase)"
    $ExchangeMailboxs = Get-Mailbox -Database $($AllExchangeDatabase) -RecipientTypeDetails UserMailbox  | ? {$_.HiddenFromAddressListsEnabled -eq $false} #| select -first 10
        ForEach($ExchangeMailbox in $ExchangeMailboxs)
            {
            #write-host "here"
            $ALLADUsersObj  = New-Object System.Object
            $ALLADUsersObj | Add-Member -type NoteProperty -name PrimarySMTPAddress -Value $ExchangeMailbox.PrimarySMTPAddress
            $ALLADUsers += $ALLADUsersObj
            $ALLADUsersCount++
            }
    Start-Sleep -Seconds 10  
    write-host "Mailbox count is now $($ALLADUsersCount)"
    write-log -message  "Mailbox count is now $($ALLADUsersCount)"
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers();
    }

write-host "###Found $($ALLADUsersCount) mailboxes in total ###"
write-log -message  "### We found $(($ALLADUsers| Measure-Object).count) mailboxes in exchange ### "

If(!($ALLADUsers))
    {
    # None were found
    Write-Log -Message "ALLADUsers was null" -Level Error
    Write-EventLog -LogName "Application" -Source "Company Sync - Database Update"  -EventID 105 -EntryType Error -Message "Exchange Mailboxes Found is low or null for instance $($Instance)"
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject "Script Terminated Unexpectedly"  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    stop-transcript
    Exit
    }
ElseIF($(($ALLADUsers | measure).count) -lt 10)
    {
    # Users found write to log
     Write-Log -Message "ALLADUsers was less than 10 - $(($ALLADUsers | measure).count)  " -Level Error
    Write-EventLog -LogName "Application" -Source "Company Sync - Database Update"  -EventID 105 -EntryType Error -Message "Exchange Mailboxes Found is low or null for instance $($Instance)"
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject "Script Terminated Unexpectedly"  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    stop-transcript
    Exit
    }
Else
    {
    # Users found write to log
    Write-Log -Message "Value for ALLADUsers is $(($ALLADUsers| measure).count)" 
    Write-EventLog -LogName "Application" -Source "Company Sync - Database Update"  -EventID 115 -EntryType Information  -Message "Value for Exchange Mailboxes ALLADUsers is $(($ALLADUsers | measure).count) for instance $($Instance)"
    }

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
    $CheckUserInDB  = $UsersArray  | Where-Object { $_.PrimarySMTPAddress -eq $DBUser.PrimarySMTPAddress} 
    # they arent so we will mark the row by setting the RecordDeleted to True
    # this will be used in the next script to purge it from teh database after the contact is deleted
    If(!($CheckUserInDB))
        {
        write-log -Message "$($DBUser.PrimarySMTPAddress) doesnt exists in the instance AD Array so we check if it exists elsehwere"
        #$DBUser.PrimarySMTPAddress
        #$CheckUserInAD  = $ALLADUsers  | Where-Object { $_.mail -eq $($DBUser.PrimarySMTPAddress.Trim())} 
        $CheckUserInAD  = $ALLADUsers  | Where-Object { $_.PrimarySmtpAddress -eq $($DBUser.PrimarySMTPAddress.Trim())}
        #$CheckUserInAD.Mail
        If(!($CheckUserInAD))
                {
                write-log -message  "$($DBUser.PrimarySMTPAddress) doesnt exists in the AD Array so we will mark if for deletion"
                #Here is the bit where we flag deleted users
                #write-host -ForegroundColor yellow "here"
                write-host -ForegroundColor Red   "$($DBUser.PrimarySMTPAddress.Trim()) not found"
                $SQLQuery = "UPDATE $SetSQLTable SET RecordDeleted='True' WHERE PrimarySMTPAddress= '$($DBUser.PrimarySMTPAddress -replace "'","''")'"
                Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase  -Query $SQLQuery
                write-log -message  "$($DBUser.PrimarySMTPAddress) doesnt exists in the AD Array so we will delete it"
                write-host  "$($DBUser.PrimarySMTPAddress) doesnt exists in the AD Array so we will delete it"
                $DeletedUserReportObj  = New-Object System.Object
                $DeletedUserReportObj | Add-Member -type NoteProperty -name User -Value $DBUser.PrimarySMTPAddress
                $DeletedUserReport += $DeletedUserReportObj
                $DeletedUserReportCount++
                #$SQLDelete = "DELETE FROM $SetSQLTable WHERE PrimarySMTPAddress='$($DBUser.PrimarySMTPAddress -replace "'","''")'"
                #Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase  -Query $SQLDelete
                }
            Else
                {
                #write-host  "$($DBUser.PrimarySMTPAddress) exists in the AD Array so we will flag it as moved"
                write-host -ForegroundColor Green  "$($DBUser.PrimarySMTPAddress.Trim()) found"
                $SQLQuery = "UPDATE $SetSQLTable SET Instance='Moved' WHERE PrimarySMTPAddress= '$($DBUser.PrimarySMTPAddress -replace "'","''")'"
                Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase  -Query $SQLQuery
                $MovedUserOutOUReportObj  = New-Object System.Object
                $MovedUserOutOUReportObj | Add-Member -type NoteProperty -name User -Value $DBUser.PrimarySMTPAddress
                $MovedUserOutOUReport += $MovedUserOutOUReportObj
                $MovedUserOutOUReportCount++
                }
        }
                
    Else
        {
        # Here we flag users as not deleted
        write-log -message  "$($DBUser.PrimarySMTPAddress)  exists in the AD Array so all is good - to be sage we will mark it so it isnt flagged for deletion" 
        $SQLQuery = "UPDATE $SetSQLTable SET RecordDeleted='False' WHERE PrimarySMTPAddress= '$($DBUser.PrimarySMTPAddress -replace "'","''")'"
        Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase  -Query $SQLQuery
        }
    }


### friday change ###
# Delete users in exclude group if they are in the database
write-log -message  "Checking if the users in GG-CompanyAfrica-Sync-Exclude are in the database and if they are flagging them for deletion" 
ForEach($ExcludeUser in $ExcludeUsers)
    {
    $ExcludedMailbox = get-mailbox -Filter "samaccountname -eq '$($ExcludeUser.samaccountname)'"
    write-log -message  "If the record for  $($ExcludedMailbox.PrimarySMTPAddress) is in the database we will flag it for deletion" 
    $SQLQuery = "UPDATE $SetSQLTable SET RecordDeleted='True' WHERE PrimarySMTPAddress= '$($ExcludedMailbox.PrimarySMTPAddress -replace "'","''")'"
    Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase  -Query $SQLQuery
    }

$ALLADUsers = $null
write-log -message  "Script has completed run - moivng onto report generation" 
# Last thing - stop transcript
Stop-Transcript

############################
# Start Generating Reports #
############################

# We will output all the arrays to a csv if via the export-file function
export-file $NewUserReport $NewUserReportOut
export-file $UpdatedUserReport $UpdatedUserReportOut
export-file $DeletedUserReport $DeletedUserReportOut
export-file $MovedUserIntoOUReport $MovedUserIntoOUReportOut
export-file $MovedUserOutOUReport $MovedUserOutOUReportOut

copy-item "$($PSScriptRoot)\$($MyInvocation.MyCommand)" "$($OutputFolder)\$($Currentdate)-Backup-$($MyInvocation.MyCommand)"

# We will then add any files to the attachments array
Add-Attachments $NewUserReportOut
Add-Attachments $UpdatedUserReportOut
Add-Attachments $DeletedUserReportOut
Add-Attachments $ErrorLog
Add-Attachments $logfile
Add-Attachments $transcript
Add-Attachments $MovedUserIntoOUReportOut
Add-Attachments $MovedUserOutOUReportOut

# Create the body of the email
$body = "Here is the results from the last run of the $($MyInvocation.MyCommand) script which ran at $Currentdate on $env:computername<br/><br/> `
Results were <br/> `
$UsersArrayCounter processed in this instance <br/> `
$($NewUserReportCount) were added to the database<br/> `
$($UpdatedUserReportCount) updated attributes <br/> `
$($DeletedUserReportCount) were set to be deleted <br/> <br/>`
The logs can be found at $($OutputFolder) `
" + $html 


# First of it there aren't attachments - send email
If(!($attachments))
    {  
    write-log -message  "There are not attachments to attach to email" 
    send-mailmessage -to $ManagmentReportsDL -from $Emailsender -bcc $CC -subject $subject -body $Body -SmtpServer $SmtpServer -BodyAsHtml #
    }
#There are attachments so sent that email
Else
    {
    write-log -message  "There are $(($attachments| Measure-Object).count) to attach to email" 
    send-mailmessage -to $ManagmentReportsDL -from $Emailsender -bcc $CC -subject $subject -Attachments $attachments -body $Body -SmtpServer $SmtpServer -BodyAsHtml #
    }

write-log -message  "Clear sessions" 
Get-PSSession | Remove-PSSession
Write-EventLog -LogName "Application" -Source "Company Sync - Database Update"  -EventID 2 -EntryType Information  -Message "Instance $($Instance) finished" 
write-log -message  "Script Finished" 