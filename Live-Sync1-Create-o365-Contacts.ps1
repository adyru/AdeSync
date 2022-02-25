<# 
  
    Script to update an office 365 GAL from a database
    THis is set to work multi domain and with mulitple copies running
    Variable sections are per instance, per domain and then standard ones that should not need to be amended
#>
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
        stop-transcript
        exit

    }
    # Set these to $true or $false - switches on and off output to screen
    # One if for info eventss the other for anything but info
    $OutputScreenInfo = $false
    #$OutputScreenInfo = $true
    $OutputScreenNotInfo = $true
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    # Create the line from the timestamp, level and message
    $Line = "$Stamp $Level $Message"
    # Create seperate log file for errors
    $script:ErrorLog = $logfile -replace ".txt", "-errors.txt"
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

###################################
# Need to be amended per instance #
###################################

#Instance No
$Instance = "1"

##################################
#First of logging and transcript #
##################################

#Date that we will look for new/update users from
$Currentdate = get-date -Format dd-MM-yyyy--hh-mm
$StartDate = (get-date).AddDays(-1)
write-host "State date is $($StartDate)"
$ReportDate = Get-Date -Format MM/dd/yyyy
$StartDate

#Create new folder for output files if it doesnt exist using todays date
$Path = "D:\LogFiles\Company-Picasso----Africa-Sync"
$FolderDate = Get-Date -Format dd-MM-yyyy
$newFolder = "$Path\$($FolderDate)\Create-o365-Contacts\Instance$($instance)"
$TestFolder = test-path $newFolder
If ($testfolder -eq $false) { New-Item -Path $NewFolder -ItemType directory }
$OutputFolder = $NewFolder


#Need to create a logfile for the write-log function as the script will error out without that
$logfile = "$OutputFolder\Sync$($instance)-Create-Office-365-Contacts-Log-$($Currentdate).txt"
$transcript = "$OutputFolder\Sync$($instance)-Create-Office-365-Contacts-Log--transcript-$($Currentdate).txt"
start-transcript $transcript 

# Create event log
New-EventLog -LogName application -Source "Company Sync - Contact Creation" -ErrorAction silentlycontinue -WarningAction SilentlyContinue
Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 1 -EntryType Information   -Message "Instance $($Instance) started"


#################################
# Need to be amended per domain #
#################################

### Load Modules #######
Write-Log -Message "Importing the sqlserver PS module)" 
Import-Module "D:\PowershellModules\sqlserver\sqlserver"
Write-Log -Message "Importing the exchangeonline PS module" 
Import-Module "D:\PowershellModules\exchangeonlinemanagement.2.0.5\ExchangeOnlineManagement"

# Connect via service account
#Write-Log -Message "Connecting to Exchangonline" 
#$ProxyOptions = New-PSSessionOption -ProxyAccessType IEConfig
#Connect-ExchangeOnline  -PSSessionOption $ProxyOptions -Credential $credential 

#Connect to o365 - via certificate
Write-Log -Message "Connecting to Exchangonline"
$ProxyOptions = New-PSSessionOption -ProxyAccessType IEConfig
Connect-ExchangeOnline -CertificateThumbPrint "*****" -AppID "*******" -Organization "*******" -ShowProgress $true -PSSessionOption $ProxyOptions

#Instance TargetAddress
$InstanceTargetAddress = "@address.co.uk"
write-host -ForegroundColor Green "Instance address is $($InstanceTargetAddress)"
write-host -ForegroundColor Green "Instance is $($Instance)"
start-sleep -Seconds 10

# SQL Variables
$SQLServer = "SQL.domain.co.uk"
$SQLDatabase = "Database"
$SQLTable = "LiveTable"
$SetSQLTable = "dbo.LiveTable"


#Production Vars for mail send at the end
$ReportsDL = "someone@somewhere.com"
$ManagmentReportsDL = "somedl@somewhere.com"
$Admin = "someone@somewhere.com"
$Sender = "someone@somewhere.com"
$SmtpServer = "smtp.domain.co.uk" # + $ADDomain.DNSRoot
$subject = "o365 Contact Creation - $NetBios Instance$($Instance)"
$CC = "someoneelse@somewhere.com"

######################
# Standard Variables #
######################

# Limit on Deleted Users
$DeletedLimit = "50"
$DeletedUSersCheckCount = 0



#Reporting variables
$DeletedUserReport = @()
$UpdatedUserReport = @()
$NewUserReport = @()
$RecreatedUserReport = @()

#Outfiles for extracting the above arrays
$NewUserReportOut = "$($OutputFolder)\NewUser-Report-$($Currentdate).csv"
$UpdatedUserReportOut = "$($OutputFolder)\UpdatedUser-Report$($Currentdate).csv"
$DeletedUserReportOut = "$($OutputFolder)\Deleted-Report$($Currentdate).csv"
$RecreatedUserReportOut = "$($OutputFolder)\Recreated-Report$($Currentdate).csv"
$RecreatedUserReportCount = 0

#Counters
$NewUserReportCount  = 0
$UpdatedUserReportCount = 0
$DeletedUserReportCount = 0

# Domain Variable
$NetBios = (Get-ADDomain).NetBIOSName




# We are going to do a garbage collection every 2 mins so 
# need to kick of a timer
#Start Stopwatch
$sw = [diagnostics.stopwatch]::StartNew()

#####################
# End Of Variables  #
#####################

write-host -ForegroundColor Green "Instance address is $($InstanceTargetAddress)"
write-host -ForegroundColor Green  "Instance is $($Instance)"
write-host
write-host -ForegroundColor Green  "SQl Server is $($SQLServer)"
write-host -ForegroundColor Green  "SQl Server Database is $($SQLDatabase)"
write-host -ForegroundColor Green  "SQl Server Table is $($SQLTable)"
write-host
write-host -ForegroundColor Green  "Delete Limit is $($DeletedLimit)"
write-host -ForegroundColor Green  "Variable check before deletion is $($CheckLimit)"


########################
# Connectivity Checks  #
########################
Write-Log -Message "### Starting the connectivity tests ###"
# Check SQL Module
$SQLCheck = get-module sqlserver
IF(!($SQLCheck ))
    {
    Write-Log -Message "SQL Powershell Module Check failed - so exiting" -Level "Error"
    Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 101 -EntryType Error -Message "SQL Powershell Module Check failed for instance $($Instance)"
    $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as the SQL PS module check failed"
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject "Script Failed To Run"  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    stop-transcript
    exit
    }
Else
    {
    Write-Log -Message "Connected to database $($Database) on server $($SQLServer)"
    Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 111 -EntryType Information -Message "SQL Powershell Module Check Succedded for instance $($Instance)"
    }

#Check SQL Connection 
Write-Log -Message "Checking the SQL connection to database $($Database) on server $($SQLServer) " 
$conn = New-Object System.Data.SqlClient.SqlConnection                                      
$conn.ConnectionString = "Server=$SQLServer;Database=$Database;Integrated Security=True;"                                                                        
$conn.Open()
IF($conn.State -ne "Open")
    {
    Write-Log -Message "Failed to connect to database $($Database) on server $($SQLServer) so exiting" -Level "Error"
    Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 102 -EntryType Error -Message "Failed to connect to database $($Database) on server $($SQLServer) so exiting for instance $($Instance)"
    $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as the SQL check failed"
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject "Script Failed To Run"  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    stop-transcript
    exit
    }
Else
    {
    Write-Log -Message "Connected to database $($Database) on server $($SQLServer)"
    Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 112 -EntryType Information -Message "Connected to database $($Database) on server $($SQLServer) for instance $($Instance)"
    }


Write-Log -Message "Check exchange powershell commands" 
$ExchangeCheck = get-command get-mailcontact -ErrorAction SilentlyContinue
IF(!($ExchangeCheck))
    {
    Write-Log -Message "Failed to load Powershell modules from o365 - so exiting" -Level "Error"
    Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 103 -EntryType Error -Message "Failed to load Powershell modules from o365 - so exiting for instance $($Instance)"
    $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as the office 365 PS module check failed"
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject "Script Failed To Run"  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    stop-transcript
    exit
    }
Else
    {
     Write-Log -Message "Exchange powershell module loaded from office 365"
     Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 113 -EntryType Information -Message "Exchange powershell module loaded from office 365 for instance $($Instance) for instance $($Instance)"
    }

Write-Log -Message "### Connectivity checks complete ###"
###############################
#  End Of Connectivity Checks #
###############################


##############
#SQL Queries #
##############
# We will query the database for new users
Write-Log -Message "Querying the database for new users"
$SQLQuery = "SELECT * FROM $($SetSQLTable) WHERE RecordAdded > '$StartDate' AND RecordAdded = RecordUpdated AND RecordDeleted = 'False' AND Instance = '$($Instance)' AND FirstRunComplete = 'False'"
$NewContacts = Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLDatabase -Query $SQLQuery 

If(!($NewContacts))
    {
    # None were found
    Write-Log -Message "No new users found in the database" 
    }
Else
    {
    # Users found write to log
    Write-Log -Message "$(($NewContacts | measure).count) new users found in the database" 
    }

# We will query the database for updated users
Write-Log -Message "Querying the database for updated users"
## Use for first run
$SQLQuery = "SELECT *  FROM $($SetSQLTable) WHERE RecordUpdated > '$StartDate'AND RecordUpdated >= RecordAdded AND RecordDeleted =  'False' AND Instance = '$($Instance)' AND FirstRunComplete = 'True'"
#User after first run
#$SQLQuery = "SELECT *  FROM $($SetSQLTable) WHERE RecordUpdated > '$StartDate'AND RecordUpdated >= RecordAdded AND RecordDeleted =  'False' AND Instance = '$($Instance)'"
$UpdatedContacts = Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLDatabase -Query $SQLQuery 
If(!($UpdatedContacts))
    {
    # None were found
    Write-Log -Message "No updated users found in the database" 
    }
Else
    {
    # Users found write to log
    Write-Log -Message "$(($UpdatedContacts | measure).count) updated users found in the database" 
    }

# We will query the database for deleted users
Write-Log -Message "Querying the database for deleted users"
#$SQLQuery = "SELECT * FROM $($SetSQLTable) WHERE RecordDeleted = 'True' AND Instance = '$($Instance)'"
$SQLQuery = "SELECT * FROM $($SetSQLTable) WHERE RecordDeleted = 'True'"
$DeletedContacts = Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLDatabase -Query $SQLQuery 
If(!($DeletedContacts ))
    {
    # None were found
    Write-Log -Message "No deleted users found in the database" 
    }
Else
    {
    # Users found write to log
    Write-Log -Message "$(($DeletedContacts | measure).count) deleted users found in the database" 
    }

# Check if there are too many deleted users to process and exit if there are
If((($DeletedContacts | measure).count) -ge $DeletedLimit)
    {
    Write-Log -Message "$(($DeletedContacts | measure).count) deleted users is greater than the set limit - $($DeletedLimit)"   -Level Error
    $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as $(($DeletedContacts | measure).count) deleted users is greater than the set limit - $($DeletedLimit)"
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject "Script Failed To Run"  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    stop-transcript
    exit
    }
Else
    {
    Write-Log -Message "$(($DeletedContacts | measure).count) deleted users is less than the set limit - $($DeletedLimit)" 
    }

# Get all non delete users
$SQLQuery = "SELECT * FROM $($SetSQLTable) WHERE RecordDeleted = 'False' AND Instance = '$($Instance)'"
$AllSQLContacts = Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLDatabase -Query $SQLQuery 

IF(!($AllSQLContacts))
    {
    Write-Log -Message "No  users found in database"
    }

Else
    {
    Write-Log -Message "$(($AllSQLContacts| measure).count) users found in database"
    }
#####################
#End of SQL Queries #
#####################



##################################
# Process Deleted Database Users #
##################################
ForEach($DeletedContact  in $DeletedContacts)
    {
    write-host "we will delete $($DeletedContact.PrimarySMTPAddress)"
    Write-Log -Message "will delete $($DeletedContact.PrimarySMTPAddress)"
    remove-mailcontact $DeletedContact.PrimarySMTPAddress -confirm:$false -erroraction silentlycontinue
    #remove-mailcontact $AllADContact.PrimarySMTPAddress -confirm:$true -erroraction silentlycontinue -whatif
    Write-Log -Message "we will check $($DeletedContact.PrimarySMTPAddress) has been deleted"
    $DEleteCheckContact =  $DeletedContact.PrimarySMTPAddress -replace "'","''"
    #$CheckContact = get-mailcontact -filter $DEleteCheckContact -erroraction silentlycontinue
    $CheckContact = Get-MailContact -Filter "PrimarySMTPAddress -like '*$($DEleteCheckContact)'" 
    If(!($CheckContact))
        {
        write-host -ForegroundColor yellow "$($DeletedContact.PrimarySMTPAddress) has been deleted"
        Write-Log -Message "$($DeletedContact.PrimarySMTPAddress) has been deleted so we will remove from the database table"
        $SQLDelete = "DELETE FROM $SetSQLTable WHERE PrimarySMTPAddress='$($DeletedContact.PrimarySMTPAddress -replace "'","''")'"
        Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLdatabase  -Query $SQLDelete
        $DeletedUserReportObj  = New-Object System.Object
        $DeletedUserReportObj | Add-Member -type NoteProperty -name User -Value $DeletedContact.PrimarySMTPAddress
        $DeletedUserReport+= $DeletedUserReportObj
        $DeletedUserReportCount++


        }
    Else
        {
        Write-Log -Message "$($CheckContact.PrimarySMTPAddress) still exists" -Level Error
        
        }
    }


write-host -ForegroundColor Green "sleeping for 30 secs......"
start-sleep -Seconds 10


#####################
# Process New Users #
#####################

Write-Log -Message "Getting a list of all contacts on office 365"
### Friday Change ###
#$CheckNewContacts = Get-MailContact -Filter "ExternalEmailAddress -like '*$($InstanceTargetAddress)'"  -ResultSize unlimited | select name,PrimarySmtpAddress
$CheckNewContacts = Get-recipient -Filter "ExternalEmailAddress -like '*$($InstanceTargetAddress)'"  -ResultSize unlimited | select name,PrimarySmtpAddress

If(!($CheckNewContacts))
    {
    Write-Log -Message "No users found in AD" -Level Error
    }
Else
    {
    Write-Log -Message "$(($CheckNewContacts | measure).count) contacts found in AD" 
    }  
write-Log -Message "Starting to process new users"  #write-host 
# Start to process each new user
ForEach($NewContact in $NewContacts)
    {
    $CheckContactExists  = ""
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
        write-log -Message "### User was $($NewUserReportCount) out of $(($NewContacts | measure).count) new users ###"

    }
    Write-Log -Message "Starting to process new user  $($NewContact.PrimarySMTPAddress)" 
    #Get around the issue with blank surname
    Write-Log -Message "Checking if $($NewContact.PrimarySMTPAddress) has a surname value" 
    if(!($NewContact.Surname))
        {
        # Set name and display name to just the firstname
        $name = "$($NewContact.FirstName.Trim())"
        $DisplayName ="$($NewContact.FirstName.Trim())"
        Write-Log -Message "$($NewContact.PrimarySMTPAddress) doesnt have a surname value so setting name to $($name) and display name to $($DisplayName)" 
        }
    Else
        {
        # Set name and display name to firstname and firstname
        $name = "$($NewContact.Surname.Trim()), $($NewContact.FirstName.Trim())"
        $DisplayName = "$($NewContact.Surname.Trim()), $($NewContact.FirstName.Trim())"
        Write-Log -Message "$($NewContact.PrimarySMTPAddress) does have a surname value so setting name to $($name) and display name to $($DisplayName)" 
        }
    # Get around the issue with a blank country code
    Write-Log -Message "Checking if $($NewContact.PrimarySMTPAddress) has a Country value" 
    If(!($NewContact.Country))
        {
        $CountryCode -eq $null
        $CountryCode = $null
        Write-Log -Message "$($NewContact.PrimarySMTPAddress) doesnt have a Country value" 
        }
    ElseIf($NewContact.Country -eq [System.DBNull]::Value)
        {
        $CountryCode -eq $null
        $CountryCode = $null
        Write-Log -Message "$($NewContact.PrimarySMTPAddress) has value null value"
        }
    Else
        {
        $CountryCode = $NewContact.Country
        Write-Log -Message "$($NewContact.PrimarySMTPAddress) does have a Country value - $($NewContact.Country)" 
        }
    write-host "will create $($NewContact.PrimarySMTPAddress)"
    Write-Log -Message "will create $($NewContact.PrimarySMTPAddress)" 
    # We will check to see if there is an existing user
    Write-Log -Message "Checking if $($NewContact.TargetAddress) already exists" 
    #$CheckNewContact = get-mailcontact $NewContact.TargetAddress -erroraction silentlycontinue
    $CheckContactExists  = $CheckNewContacts  | ? { $_.PrimarySMTPAddress -eq $NewContact.PrimarySMTPAddress}
    #write-host "CheckContactExists is $($CheckContactExists.PrimarySMTPAddress)"
    If(!($CheckContactExists))
        {
        Write-Log -Message "$($NewContact.PrimarySMTPAddress) doesnt exists so we will create it"
        # write host we use the target address to create the contact as you cannot set the primary and external to different addresses when creating 
        ### changing here #######
        #new-mailcontact -ExternalEmailAddress $($NewContact.TargetAddress) -FirstName $($NewContact.FirstName.Trim()) -LastName $($NewContact.Surname.Trim()) -name $($name) `
        #                     -displayname $($displayName)
        if(!($NewContact.Surname))
                {
                Write-Log -Message "$($NewContact.PrimarySMTPAddress) doesnt have a surname value"
                $Number = 0
                $COmplete = 0
                Do {
                Write-Log -Message "Checking if name $($name) is in use for $($NewContact.PrimarySMTPAddress)"
                $name = $name -replace "'","''"
                $CheckAD = Get-recipient -Filter "name -eq '$name'" 
                $name = $name -replace "''","'"
                if(!($CheckAD))
                    {
                    write-log -Message  "$($name) doesnt exist so creating $($NewContact.PrimarySMTPAddress) ......"
                    new-mailcontact -ExternalEmailAddress $($NewContact.TargetAddress) -FirstName $($NewContact.FirstName.Trim()) -name $($name) `
                                -displayname $($displayName)
                    $SQLUpdate = "Update $SetSQLTable SET Name= '$($name -replace "'","''")' WHERE PrimarySMTPAddress='$($NewContact.PrimarySMTPAddress -replace "'","''")'"
                    Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $SQLUpdate 
                    $Number = 0
                    $COmplete = 1
                    break
                    }
                else{
                    write-log -Message  "$($name) exists so moving on to next number for $($NewContact.PrimarySMTPAddress)"
                    $Number = $Number+1
                    $name = "$($NewContact.FirstName.Trim())$($Number)" 
                }
                } # End of 'Do'
                While ($Complete -lt 1)
                }
            Else
                {
                Write-Log -Message "$($NewContact.PrimarySMTPAddress) does have a surname value"
                $Number = 0
                $COmplete = 0
                Do {
                Write-Log -Message "Checking if name $($name) is in use for $($NewContact.PrimarySMTPAddress)"
                $name = $name -replace "'","''"
                $CheckAD = Get-recipient -Filter "name -eq '$name'" 
                $name = $name -replace "''","'"
                if(!($CheckAD))
                    {
                    write-log -Message "$($name) doesnt exist so creating $($NewContact.PrimarySMTPAddress) ......"
                    new-mailcontact -ExternalEmailAddress $($NewContact.TargetAddress) -FirstName $($NewContact.FirstName.Trim()) -LastName $($NewContact.Surname.Trim()) -name $($name) `
                                -displayname $($displayName)
                    $SQLUpdate = "Update $SetSQLTable SET Name= '$($name -replace "'","''")' WHERE PrimarySMTPAddress='$($NewContact.PrimarySMTPAddress -replace "'","''")'"
                    Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $SQLUpdate 
                    $Number = 0
                    $COmplete = 1
                    break
                    }
                else{
                    write-log -Message  "$($name) exists so moving on to next number for $($NewContact.PrimarySMTPAddress)"
                    $Number = $Number+1
                    #write-host "Number is $($number)"
                    $name = "$($NewContact.Surname.Trim())$($Number), $($NewContact.FirstName.Trim())" 
                    #write-host $name
                }
                } # End of 'Do'
                While ($Complete -lt 1)
                }


        ### End of changes ###

        # We will now change the primary SMTP address
        Write-Log -Message "We will now set the mail settings on mail contact $($NewContact.TargetAddress) - this includes changing the primary address" 
        set-MailContact $($NewContact.TargetAddress)  -EmailAddresses "SMTP:$($NewContact.PrimarySMTPAddress)","smtp:$($NewContact.TargetAddress)" `
                    -displayname $($displayName) -CustomAttribute1 $($Instance)
        Write-Log -Message "We will now set the AD on contact $($NewContact.PrimarySMTPAddress)" 
        # We will now change various user variables
        set-contact $($NewContact.PrimarySMTPAddress)   -phone $($NewContact.telephoneNumber) -Mobile $($NewContact.Mobile) -Office $($NewContact.Office) `
            -Department $($NewContact.Department) -Company $($NewContact.Company) -CountryOrRegion $($CountryCode) -title $($NewContact.title)  `
            -StreetAddress $($NewContact.StreetAddress) -City $($NewContact.City) -PostalCode $($NewContact.postalcode) -StateOrProvince $($NewContact.State) `
        # Update SQL to say we have processed this user
        Write-Log -Message "We will now update the SQL Database to show $($NewContact.PrimarySMTPAddress) has been processed" 
        $SQLUpdate = "Update $SetSQLTable SET FirstRunComplete = 'True' WHERE PrimarySMTPAddress='$($NewContact.PrimarySMTPAddress -replace "'","''")'"
        Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $SQLUpdate 
        $NewUserReportObj  = New-Object System.Object
        $NewUserReportObj | Add-Member -type NoteProperty -name User -Value $NewContact.PrimarySMTPAddress
        $NewUserReport+= $NewUserReportObj
        $NewUserReportCount++
        #write-log -Message "$NewUserReportCount"
        }
    Else
        {
        write-host -ForegroundColor Red "Contact $($NewContact.TargetAddress) exists"
        Write-Log -Message "Contact $($NewContact.TargetAddress) exists" -Level "Error"
        }
    } 


#########################
# Process Updated Users #
#########################
write-Log -Message "Starting to process updated users"  #write-host 
# Start to process eachupdate users
ForEach($UpdatedContact in $UpdatedContacts)
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
    Write-Log -Message "Starting to process updated user  $($UpdatedContact.PrimarySMTPAddress)" 
    #Get around the issue with blank surname
    Write-Log -Message "Checking if $($UpdatedContact.PrimarySMTPAddress) has a surname value" 
    if(!($UpdatedContact.Surname))
        {
         # Set name and display name to just the firstname
        $name = "$($UpdatedContact.FirstName.Trim())"
        $DisplayName = "$($UpdatedContact.FirstName.Trim())"
        Write-Log -Message "$($UpdatedContact.PrimarySMTPAddress) doesnt have a surname value so setting name to $($name) and display name to $($DisplayName)"  
        }
    Else
        {
        # Set name and display name to firstname and firstname
        $name = "$($UpdatedContact.Surname.Trim()), $($UpdatedContact.FirstName.Trim())"
        $DisplayName = "$($UpdatedContact.Surname.Trim()), $($UpdatedContact.FirstName.Trim())"
       Write-Log -Message "$($UpdatedContact.PrimarySMTPAddress) does have a surname value so setting name to $($name) and display name to $($DisplayName)" 
        }
    # Get around the issue with a blank country code
    Write-Log -Message "Checking if $($UpdatedContact.PrimarySMTPAddress) has a Country value" 
    If(!($UpdatedContact.Country))
        {
        $CountryCode -eq $null
        $CountryCode = $null
        Write-Log -Message "$($UpdatedContact.PrimarySMTPAddress) doesnt have a Country value" 
        }
    ElseIf($UpdatedContact.Country -eq [System.DBNull]::Value)
        {
        $CountryCode = $null
        Write-Log -Message "$($UpdatedContact.PrimarySMTPAddress) has value null value"
        }
    Else
        {
        $CountryCode = $UpdatedContact.Country
        Write-Log -Message "$($UpdatedContact.PrimarySMTPAddress) does have a Country value - $($UpdatedContact.Country)" 
        }
    Write-Log -Message "will update $($UpdatedContact.PrimarySMTPAddress)" 
    write-host "will update $($UpdatedContact.PrimarySMTPAddress)"
    Write-Log -Message "Setting Mail attributes on $($UpdatedContact.PrimarySMTPAddress)" 
    set-mailcontact $($UpdatedContact.TargetAddress) -WindowsEmailAddress $($UpdatedContact.PrimarySMTPAddress) -EmailAddresses "SMTP:$($UpdatedContact.PrimarySMTPAddress)","smtp:$($UpdatedContact.TargetAddress)" -CustomAttribute1 $($Instance)
    Write-Log -Message "Setting AD attributes on $($UpdatedContact.PrimarySMTPAddress)"
    ### Changed start here ###
    If($UpdatedContact.NameUpdated -eq "1")
        {
        write-host "Nameupdated value on $($UpdatedContact.PrimarySMTPAddress) is $($UpdatedContact.NameUpdated)" -ForegroundColor Green
        if(!($UpdatedContact.Surname))
                    {
                    Write-Log -Message "$($UpdatedContact.PrimarySMTPAddress) doesnt have a surname value"
                    $Number = 0
                    $COmplete = 0
                    Do {
                    Write-Log -Message "Checking if name $($name) is in use for $($UpdatedContact.PrimarySMTPAddress)"
                    $name = $name -replace "'","''"
                    $CheckAD = Get-recipient -Filter "name -eq '$name'" 
                    $name = $name -replace "''","'"
                    if(!($CheckAD))
                        {
                        write-log -Message  "$($name) doesnt exist so updating $($UpdatedContact.PrimarySMTPAddress) ......"
                        set-contact $($UpdatedContact.PrimarySMTPAddress)   -phone $($UpdatedContact.telephoneNumber) -Mobile $($UpdatedContact.Mobile) -Office $($UpdatedContact.Office) `
                                        -Department $($UpdatedContact.Department) -Company $($UpdatedContact.Company) -displayname $($displayName)  -title $($UpdatedContact.title)  `
                                        -FirstName $($UpdatedContact.FirstName.Trim())  -Name $name -CountryOrRegion $($CountryCode) `
                                        -StreetAddress $($UpdatedContact.StreetAddress) -City $($UpdatedContact.City) -PostalCode $($UpdatedContact.postalcode) -StateOrProvince $($UpdatedContact.State) 
                        $SQLUpdate = "Update $SetSQLTable SET Name= '$($name -replace "'","''")',NameUpdated='0' WHERE PrimarySMTPAddress='$($UpdatedContact.PrimarySMTPAddress -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $SQLUpdate 
                        $Number = 0
                        $COmplete = 1
                        break
                        }
                    else{
                        write-log -Message  "$($name) exists so moving on to next number for $($UpdatedContact.PrimarySMTPAddress)"
                        $Number = $Number+1
                        $name = "$($UpdatedContact.FirstName.Trim())$($Number)" 
                    }
                    } # End of 'Do'
                    While ($Complete -lt 1)
                    }
                Else
                    {
                    Write-Log -Message "$($UpdatedContact.PrimarySMTPAddress) does have a surname value"
                    $Number = 0
                    $COmplete = 0
                    Do {
                    Write-Log -Message "Checking if name $($name) is in use for $($UpdatedContact.PrimarySMTPAddress)"
                    $name = $name -replace "'","''"
                    $CheckAD = Get-recipient -Filter "name -eq '$name'" 
                    $name = $name -replace "''","'"
                    if(!($CheckAD))
                        {
                        write-log -Message "$($name) doesnt exist so creating $($UpdatedContact.PrimarySMTPAddress) ......"
                        set-contact $($UpdatedContact.PrimarySMTPAddress)   -phone $($UpdatedContact.telephoneNumber) -Mobile $($UpdatedContact.Mobile) -Office $($UpdatedContact.Office) `
                                        -Department $($UpdatedContact.Department) -Company $($UpdatedContact.Company) -displayname $($displayName)  -title $($UpdatedContact.title)  `
                                        -FirstName $($UpdatedContact.FirstName.Trim()) -LastName $($UpdatedContact.Surname.Trim()) -Name $name -CountryOrRegion $($CountryCode) `
                                        -StreetAddress $($UpdatedContact.StreetAddress) -City $($UpdatedContact.City) -PostalCode $($UpdatedContact.postalcode) -StateOrProvince $($UpdatedContact.State) 
                        $SQLUpdate = "Update $SetSQLTable SET Name= '$($name -replace "'","''")',NameUpdated='0' WHERE PrimarySMTPAddress='$($UpdatedContact.PrimarySMTPAddress -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $SQLUpdate 
                        $Number = 0
                        $COmplete = 1
                        break
                        }
                    else{
                        write-log -Message  "$($name) exists so moving on to next number for $($UpdatedContact.PrimarySMTPAddress)"
                        $Number = $Number+1
                        #write-host "Number is $($number)"
                        $name = "$($UpdatedContact.Surname.Trim())$($Number), $($UpdatedContact.FirstName.Trim()) " 
                        #write-host $name
                    }
                    } # End of 'Do'
                    While ($Complete -lt 1)
                }
        }
    Else
        {
        #write-host -ForegroundColor Yello "here"
        #write-host "Nameupdated value on $($UpdatedContact.PrimarySMTPAddress) is $($UpdatedContact.NameUpdated)"
        write-log -Message  "No firstname or surname change on $($UpdatedContact.PrimarySMTPAddress)"
        set-contact $($UpdatedContact.PrimarySMTPAddress)   -phone $($UpdatedContact.telephoneNumber) -Mobile $($UpdatedContact.Mobile) -Office $($UpdatedContact.Office) `
                 -Department $($UpdatedContact.Department) -Company $($UpdatedContact.Company) -displayname $($displayName)  -title $($UpdatedContact.title)  `
                    -FirstName $($UpdatedContact.FirstName.Trim()) -LastName $($UpdatedContact.Surname.Trim())  -CountryOrRegion $($CountryCode) `
                    -StreetAddress $($UpdatedContact.StreetAddress) -City $($UpdatedContact.City) -PostalCode $($UpdatedContact.postalcode) -StateOrProvince $($UpdatedContact.State)
    #### changes stop here ### 
        }
    #set-contact $($UpdatedContact.PrimarySMTPAddress)   -phone $($UpdatedContact.telephoneNumber) -Mobile $($UpdatedContact.Mobile) -Office $($UpdatedContact.Office) `
    #             -Department $($UpdatedContact.Department) -Company $($UpdatedContact.Company) -displayname $($displayName)  -title $($UpdatedContact.title)  `
    #                -FirstName $($UpdatedContact.FirstName.Trim()) -LastName $($UpdatedContact.Surname.Trim())  -CountryOrRegion $($CountryCode) `
    #                -StreetAddress $($UpdatedContact.StreetAddress) -City $($UpdatedContact.City) -PostalCode $($UpdatedContact.postalcode) -StateOrProvince $($UpdatedContact.State)
    #### changes stop here ### 
                    # removed -name $($name)
    $SQLUpdate = "Update $SetSQLTable SET FirstRunComplete = 'True' WHERE PrimarySMTPAddress='$($UpdatedContact.PrimarySMTPAddress -replace "'","''")'"
    Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $SQLUpdate 
    $UpdatedUserReportObj  = New-Object System.Object
    $UpdatedUserReportObj | Add-Member -type NoteProperty -name User -Value $UpdatedContact.PrimarySMTPAddress
    $UpdatedUserReport+= $UpdatedUserReportObj
    $UpdatedUserReportCount++
                    
    }



##################################
# Process Deleted AD Contacts    #
##################################

write-host -ForegroundColor Green "sleeping for 30 secs......"
start-sleep -Seconds 10
#Last of we will check AD for deleted Contacts and recreate if they dont exist
Write-Log -Message "We will check AD for deleted Contacts and recreate if they dont exist"

Write-Log -Message "Getting all relevent contacts from Office 365"
$AllADInstanceContacts = Get-MailContact -Filter "ExternalEmailAddress -like '*$($InstanceTargetAddress)'" -resultsize unlimited| ? {$_.CustomAttribute1 -eq $($Instance)}
Write-Log -Message "We found $(($AllADInstanceContacts | measure).count) contacts in Office 365 for this instance"
# Check if contact value is null or less than 10
If(!($AllADInstanceContacts))
    {
    # None were found
    Write-Log -Message "AllADInstanceContacts was null" -Level Error
    Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 107 -EntryType Error -Message "AllADInstanceContacts  count low or null so exiting for instance $($Instance)"
    $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as  AllADInstanceContacts is null"
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject "Script Terminated Unexpectedly"  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    stop-transcript
    Exit
    }
ElseIF($(($AllADInstanceContacts | measure).count) -lt 100)
    {
    # Users found write to log
     Write-Log -Message "AllADInstanceContacts was less than 10 - $(($AllADInstanceContacts| measure).count)  " -Level Error
     Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 107 -EntryType Error -Message "AllADInstanceContacts  count low or null so exiting for instance $($Instance)"
    $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as AllADInstanceContacts is less thant the limit - $(($AllADInstanceContacts | measure).count)"
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject "Script Terminated Unexpectedly"  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    stop-transcript
    Exit
    }
Else
    {
    # Users found write to log
    Write-Log -Message "Value for AllADInstanceContacts is $(($AllADInstanceContacts| measure).count)" 
    Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 117 -EntryType Information -Message "Value for AllADInstanceContacts  is $(($AllADInstanceContacts | measure).count) for instance $($Instance)"
    }



### Sunday Change ###
#$AllADContacts = Get-MailContact -Filter "ExternalEmailAddress -like '*$($InstanceTargetAddress)'"  -resultsize unlimited
$AllADContacts = Get-MailContact -Filter "ExternalEmailAddress -like '*$($InstanceTargetAddress)'"  -resultsize unlimited| ? {$_.CustomAttribute1 -eq $($Instance)}
Write-Log -Message "We found $(($AllADContacts  | measure).count) contacts in Office 365 for this instance"
If(!($AllADContacts))
    {
    # None were found
    Write-Log -Message "AllADContacts was null" -Level Error
    Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 106 -EntryType Error -Message "AllADContacts count low or null so exiting for instance $($Instance)"
    $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as $AllADContacts is null"
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject "Script Terminated Unexpectedly"  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    stop-transcript
    Exit
    }
ElseIF($(($AllADContacts | measure).count) -lt 100)
    {
    # Users found write to log
     Write-Log -Message "AllADContacts was less than 10 - $(($AllADContacts| measure).count)  " -Level Error
     Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 106 -EntryType Error -Message "AllADContacts count low or null so exiting for instance $($Instance)"
     $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as  AllADContacts  is less thant the limit - $(($AllADContacts  | measure).count)"
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject "Script Terminated Unexpectedly"  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    stop-transcript
    Exit
    }
Else
    {
    # Users found write to log
    Write-Log -Message "Value for AllADContacts is $(($AllADInstanceContacts| measure).count)" 
    Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 116 -EntryType Information -Message "Value for AllADContacts   is $(($AllADContacts  | measure).count) for instance $($Instance)"
    }


# Get all non delete users from SQL for this instance



# Get all non delete users from SQL for this instance
$SQLQuery = "SELECT * FROM $($SetSQLTable) WHERE RecordDeleted = 'False'"
$AllSQLContacts = Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLDatabase -Query $SQLQuery 
Write-Log -Message "We found $(($AllSQLContacts| measure).count) active users in the database"
# Check if we picked up any and error it it is less than 10 or null
If(!($AllSQLContacts))
    {
    # None were found
    Write-Log -Message "AllSQLContacts was null" -Level Error
    Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 104 -EntryType Error -Message "AllSQLContacts count low or null so exiting for instance $($Instance)"
    $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as AllSQLContacts is null"
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject "Script Terminated Unexpectedly"  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    stop-transcript
    Exit
    }
ElseIF($(($AllSQLContacts | measure).count) -lt 100)
    {
    # Users found write to log
     Write-Log -Message "AllSQLContacts was less than 10 - $(($AllSQLContacts | measure).count)  " -Level Error
     Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 104 -EntryType Error -Message "AllSQLContacts count low or null so exiting for instance $($Instance)"
    $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as AllSQLContacts is less thant the limit - $(($AllSQLContacts | measure).count)"
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject "Script Terminated Unexpectedly"  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    stop-transcript
    Exit
    }
Else
    {
    # Users found write to log
    Write-Log -Message "Value for AllSQLContacts is $(($AllSQLInstanceContacts| measure).count)" 
    Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 114 -EntryType Information -Message "Value for AllSQLContacts is $(($AllSQLContacts| measure).count) for instance $($Instance)"
    }

Write-Log -Message "Querying the database for all the SQL instance Users"
$SQLQuery = "SELECT * FROM $($SetSQLTable) WHERE RecordDeleted = 'False' AND Instance = '$($Instance)'"
$AllSQLInstanceContacts = Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLDatabase -Query $SQLQuery 
Write-Log -Message "We found $(($AllSQLInstanceContacts | measure).count) active users in the database for this instance"
# Check if we picked up any and error it it is less than 10 or null
If(!($AllSQLInstanceContacts))
    {
    # None were found
    Write-Log -Message "AllSQLInstanceContacts was null" -Level Error
    Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 105 -EntryType Error -Message "AllSQLInstanceContacts count low or null so exiting for instance $($Instance)"
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject "Script Terminated Unexpectedly"  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    stop-transcript
    Exit
    }
ElseIF($(($AllSQLInstanceContacts | measure).count) -lt 100)
    {
    # Users found write to log
     Write-Log -Message "AllSQLInstanceContacts was less than 10 - $(($AllSQLInstanceContacts | measure).count)  " -Level Error
     Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 105 -EntryType Error -Message "AllSQLInstanceContacts count low or null so exiting for instance $($Instance)"
     $body = "Script  $($MyInvocation.MyCommand) failed to run at $Currentdate on $env:computername as AllSQLContacts is less than the limit - $(($AllSQLInstanceContacts | measure).count)"
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject "Script Terminated Unexpectedly"  -body $Body -SmtpServer $SmtpServer -BodyAsHtml
    stop-transcript
    Exit
    }
Else
    {
    # Users found write to log
    Write-Log -Message "Value for AllSQLInstanceContacts is $(($AllSQLInstanceContacts| measure).count)" 
    Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 115 -EntryType Information -Message "Value for AllSQLInstanceContacts is $(($AllSQLInstanceContacts | measure).count) for instance $($Instance)"
    }




# We will now compare all the AD and SQL Instance users
Write-Log -Message "### Starting to compare the users in AD against the users in the database ###"
Write-Log -Message "### We are going to check that each user in active directory exists in the database ###" 
ForEach($AllADContact in $AllADContacts)
    {
    if ( $Sw.Elapsed.minutes -eq 2) {
        # it is over 2 mins so start garbage collection
        Write-Log -Message "Doing Garbage Collection after $($Sw.Elapsed.minutes ) minutes"  #write-host 
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers();
        #Reset timer by stopping and starting a new one
        $Sw.Stop()
        $sw = [diagnostics.stopwatch]::StartNew()
        write-log -Message "### User was $($NewUserReportCount) out of $(($NewContacts | measure).count) new users ###"

    }
    $CheckUserInDB  = $AllSQLContacts  | ? { $_.PrimarySMTPAddress -eq $AllADContact.PrimarySMTPAddress}
    if(!($CheckUserInDB))
         {
         if($DeletedUSersCheckCount -lt $DeletedLimit)
            {
            Write-Log -Message "$($AllADContact.PrimarySMTPAddress) does not exist in the database so deleting user" 
            remove-mailcontact $AllADContact.PrimarySMTPAddress -confirm:$false  -erroraction silentlycontinue
            #remove-mailcontact $AllADContact.PrimarySMTPAddress -confirm:$true -erroraction silentlycontinue -whatif
            $DeletedUSersCheckCount++
            $SQLDelete = "DELETE FROM $SetSQLTable WHERE PrimarySMTPAddress='$($AllADContact.PrimarySMTPAddress -replace "'","''")'"
            Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLdatabase  -Query $SQLDelete
            $DeletedUserReportObj  = New-Object System.Object
            $DeletedUserReportObj | Add-Member -type NoteProperty -name User -Value $AllADContact.PrimarySMTPAddress 
            $DeletedUserReport+= $DeletedUserReportObj
            $DeletedUserReportCount++
            }
        Else
            {
            Write-Log -Message "### $($AllADContact.PrimarySMTPAddress) does not exist in the database but it was not deleted as it was over the delete limit of $($DeletedLimit) ###" -Level Error
            }

         }
    Else
        {
        Write-Log -Message "$($AllADContact.PrimarySMTPAddress) does exist in the database - moving on to next user" 
        }
    }

### Sunday Changes ###
Connect-ExchangeOnline -CertificateThumbPrint "*********" -AppID "*******" -Organization "********" -ShowProgress $true -PSSessionOption $ProxyOptions
start-sleep -Seconds 30

$NOInstances = Get-MailContact -resultsize unlimited -Filter {customattribute1 -eq $null}  |? { $_.ExternalEmailAddress -like "*$($InstanceTargetAddress)"}
ForEach($NOInstance in $NOInstances)
    {
    Write-Log -Message "### $($NOInstance.PrimarySMTPAddress) does not have an instance set ###" -Level Error
    }

# We will now compare all the AD and SQL Instance users
Write-Log -Message "### Starting to compare the users flagged for this instance in o365 with the users flagged for this instance in SQL ###"
ForEach($AllSQLInstanceContact in $AllSQLInstanceContacts)
    {
    if ( $Sw.Elapsed.minutes -eq 2) {
        # it is over 2 mins so start garbage collection
        Write-Log -Message "Doing Garbage Collection after $($Sw.Elapsed.minutes ) minutes"  #write-host 
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers();
        #Reset timer by stopping and starting a new one
        $Sw.Stop()
        $sw = [diagnostics.stopwatch]::StartNew()
        

    }
    $CheckUserInDB  = $AllADInstanceContacts  | ? { $_.PrimarySMTPAddress -eq $AllSQLInstanceContact.PrimarySMTPAddress}
    If(!($CheckUserInDB))
    {
    Write-Log -Message "Starting to process deleted contact  $($AllSQLInstanceContact.PrimarySMTPAddress) as it exists in the database so needs to be recreated" 
    #Get around the issue with blank surname

        Write-Log -Message "Checking if $($AllSQLInstanceContact.PrimarySMTPAddress) has a surname value" 
        if(!($AllSQLInstanceContact.Surname))
            {
            # Set name and display name to just the firstname
            $name = "$($AllSQLInstanceContact.FirstName.Trim())"
            $DisplayName ="$($AllSQLInstanceContact.FirstName.Trim())"
            Write-Log -Message "$($AllSQLInstanceContact.PrimarySMTPAddress) doesnt have a surname value so setting name to $($name) and display name to $($DisplayName)" 
            }
        Else
            {
            # Set name and display name to firstname and firstname
            $name = "$($AllSQLInstanceContact.Surname.Trim()), $($AllSQLInstanceContact.FirstName.Trim())"
            $DisplayName = "$($AllSQLInstanceContact.Surname.Trim()), $($AllSQLInstanceContact.FirstName.Trim())"
            Write-Log -Message "$($AllSQLInstanceContact.PrimarySMTPAddress) does have a surname value so setting name to $($name) and display name to $($DisplayName)" 
            }
    
        # Get around the issue with a blank country code
        Write-Log -Message "Checking if $($AllSQLInstanceContact.PrimarySMTPAddress) has a Country value" 
        If(!($AllSQLInstanceContact.Country))
            {
            $CountryCode -eq $null
            $CountryCode = $null
            Write-Log -Message "$($AllSQLInstanceContact.PrimarySMTPAddress) doesnt have a Country value" 
            }
        ElseIf($AllSQLInstanceContact.Country -eq [System.DBNull]::Value)
            {
            $CountryCode -eq $null
            $CountryCode = $null
            Write-Log -Message "$($UpdatedContact.PrimarySMTPAddress) has value null value"
            }
        Else
            {
            $CountryCode = $AllSQLInstanceContact.Country
            Write-Log -Message "$($AllSQLInstanceContact.PrimarySMTPAddress) does have a Country value - $($AllSQLInstanceContact.Country)" 
            }

        write-host "will create $($AllSQLInstanceContact.PrimarySMTPAddress)"
        Write-Log -Message "will create $($AllSQLInstanceContact.PrimarySMTPAddress)" 
        # We will check to see if there is an existing user
        Write-Log -Message "Checking if $($AllSQLInstanceContact.PrimarySMTPAddress) already exists" 
        $CheckNewContact = get-mailcontact $AllSQLInstanceContact.PrimarySMTPAddress -erroraction silentlycontinue
        #There isnt an exisiting user so we will create it
          If(!($CheckNewCOntact))
            {
            #write-host -ForegroundColor yellow "Here"
            Write-Log -Message "$($AllSQLInstanceContact.PrimarySMTPAddress) doesnt exists so we will re-create it"
            # write host we use the target address to create the contact as you cannot set the primary and external to different addresses when creating 
            #new-mailcontact -ExternalEmailAddress $($AllSQLInstanceContact.TargetAddress) -FirstName $($AllSQLInstanceContact.FirstName.Trim()) -LastName $($AllSQLInstanceContact.Surname.Trim()) -name $($name) `
            #                    -displayname $($displayName)

            ##### Changes here ####
                if(!($AllSQLInstanceContact.Surname))
                    {
                    Write-Log -Message "$($AllSQLInstanceContact.PrimarySMTPAddress) doesnt have a surname value"
                    $Number = 0
                    $COmplete = 0
                    Do {
                    Write-Log -Message "Checking if name $($name) is in use for $($AllSQLInstanceContact.PrimarySMTPAddress)"
                    $name = $name -replace "'","''"
                    $CheckAD = Get-recipient -Filter "name -eq '$name'" 
                    $name = $name -replace "''","'"
                    if(!($CheckAD))
                        {
                        write-log -Message  "$($name) doesnt exist so creating $($AllSQLInstanceContact.PrimarySMTPAddress) ......"
                        new-mailcontact -ExternalEmailAddress $($AllSQLInstanceContact.TargetAddress) -FirstName $($AllSQLInstanceContact.FirstName.Trim()) -name $($name) `
                                     -displayname $($displayName)
                        $SQLUpdate = "Update $SetSQLTable SET Name= '$($name -replace "'","''")' WHERE PrimarySMTPAddress='$($AllSQLInstanceContact.PrimarySMTPAddress -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $SQLUpdate 
                        $Number = 0
                        $COmplete = 1
                        break
                        }
                    else{
                        write-log -Message  "$($name) exists so moving on to next number for $($AllSQLInstanceContact.PrimarySMTPAddress)"
                        $Number = $Number+1
                        $name = "$($AllSQLInstanceContact.FirstName.Trim())$($Number)" 
                    }
                    } # End of 'Do'
                    While ($Complete -lt 1)
                    }
                Else
                    {
                    Write-Log -Message "$($AllSQLInstanceContact.PrimarySMTPAddress) does have a surname value"
                    $Number = 0
                    $COmplete = 0
                    Do {
                    Write-Log -Message "Checking if name $($name) is in use for $($AllSQLInstanceContact.PrimarySMTPAddress)"
                    $name = $name -replace "'","''"
                    $CheckAD = Get-recipient -Filter "name -eq '$name'" 
                    $name = $name -replace "''","'"
                    if(!($CheckAD))
                        {
                        write-log -Message "$($name) doesnt exist so creating $($AllSQLInstanceContact.PrimarySMTPAddress) ......"
                        new-mailcontact -ExternalEmailAddress $($AllSQLInstanceContact.TargetAddress) -FirstName $($AllSQLInstanceContact.FirstName.Trim()) -LastName $($AllSQLInstanceContact.Surname.Trim()) -name $($name) `
                                     -displayname $($displayName)
                        $SQLUpdate = "Update $SetSQLTable SET Name= '$($name -replace "'","''")' WHERE PrimarySMTPAddress='$($AllSQLInstanceContact.PrimarySMTPAddress -replace "'","''")'"
                        Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDatabase -Query $SQLUpdate 
                        $Number = 0
                        $COmplete = 1
                        break
                        }
                    else{
                        write-log -Message  "$($name) exists so moving on to next number for $($AllSQLInstanceContact.PrimarySMTPAddress)"
                        $Number = $Number+1
                        #write-host "Number is $($number)"
                        $name = "$($AllSQLInstanceContact.Surname.Trim())$($Number), $($AllSQLInstanceContact.FirstName.Trim())" 
                        #write-host $name
                    }
                    } # End of 'Do'
                    While ($Complete -lt 1)
                }


            #### End of changes

            # We will now change the primary SMTP address
            Write-Log -Message "We will now set the mail settings on mail contact $($AllSQLInstanceContact.TargetAddress) - this includes changing the primary address" 
            set-MailContact $($AllSQLInstanceContact.TargetAddress)  -EmailAddresses "SMTP:$($AllSQLInstanceContact.PrimarySMTPAddress)","smtp:$($AllSQLInstanceContact.TargetAddress)" `
                        -displayname $($displayName) -CustomAttribute1 $($Instance)
            Write-Log -Message "We will now set the AD on contact $($AllSQLInstanceContact.PrimarySMTPAddress)" 
            # We will now change various user variables
            set-contact $($AllSQLInstanceContact.PrimarySMTPAddress)   -phone $($AllSQLInstanceContact.telephoneNumber) -Mobile $($AllSQLInstanceContact.Mobile) -Office $($AllSQLInstanceContact.Office) `
                -Department $($AllSQLInstanceContact.Department) -Company $($AllSQLInstanceContact.Company) -CountryOrRegion $($CountryCode) -title $($AllSQLInstanceContact.title) `
                -StreetAddress $($AllSQLInstanceContact.StreetAddress) -City $($AllSQLInstanceContact.City) -PostalCode $($AllSQLInstanceContact.postalcode) -StateOrProvince $($AllSQLInstanceContact.State) 
                $RecreatedUserReportObj  = New-Object System.Object
                $RecreatedUserReportObj | Add-Member -type NoteProperty -name User -Value $NewContact.PrimarySMTPAddress
                $RecreatedUserReport+= $RecreatedUserReportObj
                $RecreatedUserReportCount++
            
            }
         Else
            {
            #There is a user in the database that is supposed to be in this instance so we will amend customattribute1
            #write-host -ForegroundColor Red "Contact $($AllSQLInstanceContact.TargetAddress) exists"
            write-host -ForegroundColor yellow "We are here"
            Write-Log -Message "Contact $($AllSQLInstanceContact.PrimarySMTPAddress) exists" -Level "Error"
            write-log -message "Updating the CustomAttribute1 for $($AllSQLInstanceContact.PrimarySMTPAddress) to $($instance)"
            set-mailcontact $AllSQLInstanceContact.PrimarySMTPAddress -CustomAttribute1 $instance
            }
    } 
    
    Else
        {
        write-log -message "Check if the user has the wrong instance value"
        write-host "CheckUserInDB.CustomAttribute1 value is $($CheckUserInDB.CustomAttribute1)"
        #$CheckInstanceValue = (get-mailcontact $AllSQLInstanceContact.PrimarySMTPAddress).CustomAttribute1
        #If($CheckInstanceValue -ne $Instance)
        If($CheckUserInDB.CustomAttribute1 -ne $Instance)
            {
            write-log -message "Updating the CustomAttribute1 for $($AllSQLInstanceContact.PrimarySMTPAddress) to $($instance)"
            set-mailcontact $AllSQLInstanceContact.PrimarySMTPAddress -CustomAttribute1 $instance
            }
        Else
            {
            Write-Log -Message "$($AllSQLInstanceContact.PrimarySMTPAddress) still exists in Office 365 and has the correct instance - $($instance)"
            }
        }





    }



############################
# Start Generating Reports #
############################

# We will output all the arrays to a csv if via the export-file function
export-file $NewUserReport $NewUserReportOut
export-file $UpdatedUserReport $UpdatedUserReportOut
export-file $DeletedUserReport $DeletedUserReportOut
export-file $RecreatedUserReport $RecreatedUserReportOut

#Stop transcipt before sending
Stop-Transcript

# We will then add any files to the attachments array
Add-Attachments $NewUserReportOut
Add-Attachments $UpdatedUserReportOut
Add-Attachments $DeletedUserReportOut
Add-Attachments $ErrorLog
Add-Attachments $logfile
Add-Attachments $transcript
Add-Attachments $RecreatedUserReportOut


copy-item "$($PSScriptRoot)\$($MyInvocation.MyCommand)" "$($OutputFolder)\$($Currentdate)-Backup-$($MyInvocation.MyCommand)"

$NewContacts = Get-MailContact -filter "whenCreated -lt '$ReportDate' -and whenchanged -gt '$ReportDate' -and customattribute1 -eq '$instance'"  -ResultSize unlimited
$UpdatedContacts = Get-MailContact -filter "whenCreated -gt '$ReportDate' -and whenchanged -gt '$ReportDate' -and customattribute1 -eq '$instance'" -ResultSize unlimited

# Create the body of the email
"$(($UpdatedContacts| measure).count) new contacts found in office 365"
"$(($NewContacts| measure).count) updated contacts found in office 365"

$body = "Here is the results from the last run of the $($MyInvocation.MyCommand) script which ran at $Currentdate on $env:computername<br/><br/> `
In Office 365 <br/> `
$(($UpdatedContacts| measure).count) new contacts found in office 365 <br/> `
$(($NewContacts| measure).count) updated contacts found in office 365 <br/><br/> `
In the database <br/> `
$(($NewContacts | measure).count) new users found in the database <br/> `
$(($UpdatedContacts | measure).count) updated users found in the database <br/> <br/>`
The logs can be found at $($OutputFolder) `
" + $html 


# First of it there aren't attachments - send email
If(!($attachments))
    {  
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject $subject -body $Body -SmtpServer $SmtpServer -BodyAsHtml #
    }
#There are attachments so sent that email
Else
    {
    send-mailmessage -to $ManagmentReportsDL -from $sender -bcc $CC -subject $subject -Attachments $attachments -body $Body -SmtpServer $SmtpServer -BodyAsHtml #
    }
 
 Write-EventLog -LogName "Application" -Source "Company Sync - Contact Creation"  -EventID 2 -EntryType Information  -Message "Instance $($Instance) finished"

#$DeletedContacts 

#Clear all sessions
get-PSSession | Remove-PSSession
