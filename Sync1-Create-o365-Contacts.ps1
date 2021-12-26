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

Function Check-Contact
{
<#
This function compares AD and database values
To see if they are the same
Has to be in order of emailaddress,Database value
AD Value and then the attribute name
#>
    [CmdletBinding()]
        Param(
        [Parameter(Mandatory=$True)]
        [String]$EmailAddress,

        [Parameter(Mandatory=$True)]
        [string]$ADAttribute,

        [Parameter(Mandatory=$True)]
        [AllowEmptyString()]
        [String]$DatabaseAttributeValue,

        [Parameter(Mandatory=$False)]
        [AllowEmptyString()]
        [Array]$ADAttributeValue



        )

    BEGIN
    {
        #Write-Host "In Begin block"
        IF($DatabaseAttributeValue = "NullValue"){$DatabaseAttributeValue =$null}
        IF($ADAttributeValue = "NullValue"){$ADAttributeValue =$null}
            If(!($DatabaseAttributeValue))
            {
            Write-Log -Message "Setting $($ADAttributeValue) to null for user $($EmailAddress)"
            $DatabaseAttributeValue = $null
            }
    }
    PROCESS
    {
        Write-Log -Message "Checking if $($EmailAddress) has $($ADAttribute) set to  $($ADAttributeValue)"
        If($($DatabaseAttributeValue) -match $($ADAttributeValue))
            {
            Write-Log -Message "$($ADAttributeValue) matches $($DatabaseAttributeValue)"
            }
        Else
            {
            write-host "Here"
            Write-Log -Message "$($ADAttributeValue) does not match $($DatabaseAttributeValue)" -Level "Error"
            $SetAttribute = "-$($ADAttribute)"
            #$cmd = "Set-Contact $($A) adrian"
            Write-Log -Message "Amending $($ADAttributeValue) to $($DatabaseAttributeValue)" 
            #$cmd = "set-contact $EmailAddress $($SetAttribute) $DatabaseAttributeValue"
            #Invoke-Expression $cmd
            }

    }
    END


    {
        #Write-Host "In End block"
    }
} #END Function Test-ScriptBlock


#Connect to o365
$ProxyOptions = New-PSSessionOption -ProxyAccessType IEConfig
#Connect-ExchangeOnline -CertificateThumbPrint "31F0787C9492FC124DB8831E59395A95D61B60E2" -AppID "3c9c74e2-54a6-4916-9708-ef22edce7251" -Organization "babcockafricapre1.onmicrosoft.com" -ShowProgress $true -PSSessionOption $ProxyOptions


######################
# Start of Variables #
######################

# Static DC to use - needs to be FQDN this will be checked later on
$DC = "RPBDOMR01.RPB01.babgroup.co.uk"

#Date that we will look for new/update users from
$Currentdate = get-date -Format dd-MM-yyyy--hh-mm
$StartDate = (get-date).AddDays(-1)
$StartDate

# SQL Variables
$SQLServer = "RPBWEBR11.RPB01.babgroup.co.uk"
$SQLDatabase = "AzureGuestReports"
$SQLTable = "tmpRPA"
$SetSQLTable = "dbo.tmpRPA"

#Instance No
$Instance = "1"

#Create new folder for output files if it doesnt exist using todays date
$FolderDate = Get-Date -Format dd-MM-yyyy
$newFolder = "C:\Scripts\RPA\Output\$($FolderDate)\Create-o365-Contacts\Instance$($instance)"
$TestFolder = test-path $newFolder
If ($testfolder -eq $false) { New-Item -Path $NewFolder -ItemType directory }
$OutputFolder = $NewFolder



#Need to create a logfile for the write-log function as the script will error out without that
$logfile = "$OutputFolder\Instance$($instance)-Create-o365-Contacts-Log-$($Currentdate).txt"
start-transcript -path "$OutputFolder\Instance$($instance)-Create-o365-Contacts-transcript-$($Currentdate).txt"

# We are going to do a garbage collection every 2 mins so 
# need to kick of a timer
#Start Stopwatch
$sw = [diagnostics.stopwatch]::StartNew()

########################
#  End Of Variables    #
########################


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
$SQLQuery = "SELECT *  FROM $($SetSQLTable) WHERE RecordUpdated > '$StartDate'AND RecordUpdated >= RecordAdded AND RecordDeleted =  'False' AND Instance = '$($Instance)' AND FirstRunComplete = 'True'"
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
$SQLQuery = "SELECT * FROM $($SetSQLTable) WHERE RecordDeleted = 'True' AND Instance = '$($Instance)'"
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

#####################
# Process New Users #
#####################

Write-Log -Message "Getting a list of all contacts on office 365"
$CheckNewContacts = Get-MailContact -Filter "ExternalEmailAddress -like '*@RPW01.babgroup.co.uk' -or ExternalEmailAddress -like '*@RPB01.babgroup.co.uk'"  -ResultSize unlimited | select name,PrimarySmtpAddress
If(!($CheckNewContacts))
    {
    Write-Log -Message "No users found in AD" -Level Error
    }
Else
    {
    Write-Log -Message "$(($CheckNewContacts | measure).count) contacts found in AD" 
    }  

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

    }
    Write-Log -Message "Starting to process new user  $($NewContact.PrimarySMTPAddress)" 
    #Get around the issue with blank surname
    Write-Log -Message "Checking if $($NewContact.PrimarySMTPAddress) has a surname value" 
    if(!($NewContact.Surname))
        {
        # Set name and display name to just the firstname
        $name = "$($NewContact.FirstName)"
        $DisplayName ="$($NewContact.FirstName)"
        Write-Log -Message "$($NewContact.PrimarySMTPAddress) doesnt have a surname value so setting name to $($name) and display name to $($DisplayName)" 
        }
    Else
        {
        # Set name and display name to firstname and firstname
        $name = "$($NewContact.FirstName) $($NewContact.Surname)"
        $DisplayName = "$($NewContact.Surname), $($NewContact.FirstName)"
        Write-Log -Message "$($NewContact.PrimarySMTPAddress) does have a surname value so setting name to $($name) and display name to $($DisplayName)" 
        }
    # Get around the issue with a blank country code
    Write-Log -Message "Checking if $($NewContact.PrimarySMTPAddress) has a Country value" 
    If(!($NewContact.Country))
        {
        $CountryCode -eq $null
        Write-Log -Message "$($NewContact.PrimarySMTPAddress) doesnt have a Country value" 
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
        Write-Log -Message "$($NewContact.TargetAddress) doesnt exists so we will create it"
        # write host we use the target address to create the contact as you cannot set the primary and external to different addresses when creating 
        new-mailcontact -ExternalEmailAddress $($NewContact.TargetAddress) -FirstName $($NewContact.FirstName) -LastName $($NewContact.Surname) -name $($name) `
                             -displayname $($displayName)
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
    Write-Log -Message "Starting to process new user  $($NewContact.PrimarySMTPAddress)" 
    #Get around the issue with blank surname
    Write-Log -Message "Checking if $($UpdatedContact.PrimarySMTPAddress) has a surname value" 
    if(!($UpdatedContact.Surname))
        {
         # Set name and display name to just the firstname
        $name = "$($UpdatedContact.FirstName)"
        $DisplayName = $name = "$($UpdatedContact.FirstName)"
        Write-Log -Message "$($UpdatedContact.PrimarySMTPAddress) doesnt have a surname value so setting name to $($name) and display name to $($DisplayName)"  
        }
    Else
        {
        # Set name and display name to firstname and firstname
        $name = "$($UpdatedContact.FirstName) $($UpdatedContact.Surname)"
        $DisplayName = "$($UpdatedContact.Surname), $($UpdatedContact.FirstName)"
       Write-Log -Message "$($UpdatedContact.PrimarySMTPAddress) does have a surname value so setting name to $($name) and display name to $($DisplayName)" 
        }
    # Get around the issue with a blank country code
    Write-Log -Message "Checking if $($UpdatedContact.PrimarySMTPAddress) has a Country value" 
    If(!($UpdatedContact.Country))
        {
        $CountryCode -eq $null
        Write-Log -Message "$($UpdatedContact.PrimarySMTPAddress) doesnt have a Country value" 
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
    set-contact $($UpdatedContact.PrimarySMTPAddress)   -phone $($UpdatedContact.telephoneNumber) -Mobile $($UpdatedContact.Mobile) -Office $($UpdatedContact.Office) `
                 -Department $($UpdatedContact.Department) -Company $($UpdatedContact.Company) -displayname $($displayName)  -title $($UpdatedContact.title)  `
                    -FirstName $($UpdatedContact.FirstName) -LastName $($UpdatedContact.Surname) -name $($name) -CountryOrRegion $($CountryCode) `
                    -StreetAddress $($UpdatedContact.StreetAddress) -City $($UpdatedContact.City) -PostalCode $($UpdatedContact.postalcode) -StateOrProvince $($UpdatedContact.State) 
                    
    }


##################################
# Process Deleted Database Users #
##################################
ForEach($DeletedContact  in $DeletedContacts)
    {
    write-host "we will delete $($DeletedContact.PrimarySMTPAddress)"
    Write-Log -Message "will delete $($DeletedContact.PrimarySMTPAddress)"
    #remove-mailcontact $DeletedContact.PrimarySMTPAddress -confirm:$false
    Write-Log -Message "we will check $($DeletedContact.PrimarySMTPAddress) has been deleted"
    $CheckContact = get-mailcontact $DeletedContact.PrimarySMTPAddress -erroraction silentlycontinue
    If(!($CheckContact))
        {
        write-host -ForegroundColor yellow "Here"
        Write-Log -Message "$($DeletedContact.PrimarySMTPAddress) has been deleted so we will remove from the database table"
        $SQLDelete = "DELETE FROM $SetSQLTable WHERE PrimarySMTPAddress='$($DeletedContact.PrimarySMTPAddress -replace "'","''")'"
        Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLdatabase  -Query $SQLDelete
        }
    Else
        {
        Write-Log -Message "$($CheckContact.PrimarySMTPAddress) still exists" -Level Error
        
        }
    }



##################################
# Process Deleted AD Contacts    #
##################################

write-host -ForegroundColor Green "sleeping for 30 secs......"
start-sleep -Seconds 10
#Last of we will check AD for deleted Contacts and recreate if they dont exist
Write-Log -Message "We will check AD for deleted Contacts and recreate if they dont exist"
Write-Log -Message "Getting all relevent contacts from Office 365"
$AllADInstanceContacts = Get-MailContact -Filter "ExternalEmailAddress -like '*@RPW01.babgroup.co.uk' -or ExternalEmailAddress -like '*@RPB01.babgroup.co.uk'" -resultsize unlimited| ? {$_.CustomAttribute1 -eq $($Instance)}
$AllADContacts = Get-MailContact -Filter "ExternalEmailAddress -like '*@RPW01.babgroup.co.uk' -or ExternalEmailAddress -like '*@RPB01.babgroup.co.uk'"  -resultsize unlimited
Write-Log -Message "We found $(($AllADInstanceContacts | measure).count) contacts in Office 365 for this instance"
Write-Log -Message "We found $(($AllADContacts  | measure).count) contacts in Office 365 for this instance"
# Get all non delete users from SQL for this instance
$SQLQuery = "SELECT * FROM $($SetSQLTable) WHERE RecordDeleted = 'False' AND Instance = '$($Instance)'"
$AllSQLInstanceContacts = Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLDatabase -Query $SQLQuery 
Write-Log -Message "We found $(($AllSQLContacts| measure).count) active users in the database for this instance"

# Get all non delete users from SQL for this instance
$SQLQuery = "SELECT * FROM $($SetSQLTable) WHERE RecordDeleted = 'False'"
$AllSQLContacts = Invoke-Sqlcmd -ServerInstance $SQLServer  -Database $SQLDatabase -Query $SQLQuery 
Write-Log -Message "We found $(($AllSQLContacts| measure).count) active users in the database"


# We will now compare all the AD and SQL Instance users
Write-Log -Message "Starting to compare the users in AD against the users in the database"
Write-Log -Message "We are going to check that each user in active directory exists in the database" 
ForEach($AllADContact in $AllADContacts)
    {
    $CheckUserInDB  = $AllSQLContacts  | ? { $_.PrimarySMTPAddress -eq $AllADContact.PrimarySMTPAddress}
    if(!($CheckUserInDB))
         {
         Write-Log -Message "$($AllADContact.PrimarySMTPAddress) does not exist in the database so deleting user" 
         remove-mailcontact $AllADContact.PrimarySMTPAddress -confirm:$false
        $SQLDelete = "DELETE FROM $SetSQLTable WHERE PrimarySMTPAddress='$($AllADContact.PrimarySMTPAddress -replace "'","''")'"
        Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLdatabase  -Query $SQLDelete
         }
    Else
        {
        Write-Log -Message "$($AllADContact.PrimarySMTPAddress) does exist in the database - moving on to next user" 
        }
    }

# We will now compare all the AD and SQL Instance users
Write-Log -Message "Starting to compare the users flagged for this instance in o365 with the users flagged for this instance in SQL"
ForEach($AllSQLInstanceContact in $AllSQLInstanceContacts)
    {
    $CheckUserInDB  = $AllADInstanceContacts  | ? { $_.PrimarySMTPAddress -eq $AllSQLInstanceContact.PrimarySMTPAddress}
    If(!($CheckUserInDB))
    {
    Write-Log -Message "Starting to process deleted contact  $($AllSQLInstanceContact.PrimarySMTPAddress) as it exists in the database so needs to be recreated" 
    #Get around the issue with blank surname

    Write-Log -Message "Checking if $($AllSQLInstanceContact.PrimarySMTPAddress) has a surname value" 
    if(!($AllSQLInstanceContact.Surname))
        {
        # Set name and display name to just the firstname
        $name = "$($AllSQLInstanceContact.FirstName)"
        $DisplayName ="$($AllSQLInstanceContact.FirstName)"
        Write-Log -Message "$($AllSQLInstanceContact.PrimarySMTPAddress) doesnt have a surname value so setting name to $($name) and display name to $($DisplayName)" 
        }
    Else
        {
        # Set name and display name to firstname and firstname
        $name = "$($AllSQLInstanceContact.FirstName) $($AllSQLInstanceContact.Surname)"
        $DisplayName = "$($AllSQLInstanceContact.Surname), $($AllSQLInstanceContact.FirstName)"
        Write-Log -Message "$($AllSQLInstanceContact.PrimarySMTPAddress) does have a surname value so setting name to $($name) and display name to $($DisplayName)" 
        }
    # Get around the issue with a blank country code
    Write-Log -Message "Checking if $($AllSQLInstanceContact.PrimarySMTPAddress) has a Country value" 
    If(!($AllSQLInstanceContact.Country))
        {
        $CountryCode -eq $null
        Write-Log -Message "$($AllSQLInstanceContact.PrimarySMTPAddress) doesnt have a Country value" 
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
      If(!($CheckNewCOntact))
        {
        write-host -ForegroundColor yellow "Here"
        Write-Log -Message "$($AllSQLInstanceContact.PrimarySMTPAddres) doesnt exists so we will re-create it"
        # write host we use the target address to create the contact as you cannot set the primary and external to different addresses when creating 
        new-mailcontact -ExternalEmailAddress $($AllSQLInstanceContact.TargetAddress) -FirstName $($AllSQLInstanceContact.FirstName) -LastName $($AllSQLInstanceContact.Surname) -name $($name) `
                             -displayname $($displayName)
        # We will now change the primary SMTP address
        Write-Log -Message "We will now set the mail settings on mail contact $($AllSQLInstanceContact.TargetAddress) - this includes changing the primary address" 
        set-MailContact $($AllSQLInstanceContact.TargetAddress)  -EmailAddresses "SMTP:$($AllSQLInstanceContact.PrimarySMTPAddress)","smtp:$($AllSQLInstanceContact.TargetAddress)" `
                    -displayname $($displayName) -CustomAttribute1 $($Instance)
        Write-Log -Message "We will now set the AD on contact $($AllSQLInstanceContact.PrimarySMTPAddress)" 
        # We will now change various user variables
        set-contact $($AllSQLInstanceContact.PrimarySMTPAddress)   -phone $($AllSQLInstanceContact.telephoneNumber) -Mobile $($AllSQLInstanceContact.Mobile) -Office $($AllSQLInstanceContact.Office) `
            -Department $($AllSQLInstanceContact.Department) -Company $($AllSQLInstanceContact.Company) -CountryOrRegion $($CountryCode) -title $($AllSQLInstanceContact.title) `
            -StreetAddress $($AllSQLInstanceContact.StreetAddress) -City $($AllSQLInstanceContact.City) -PostalCode $($AllSQLInstanceContact.postalcode) -StateOrProvince $($AllSQLInstanceContact.State) 
            
        }
    Else
        {
        write-host -ForegroundColor Red "Contact $($AllSQLInstanceContact.TargetAddress) exists"
        Write-Log -Message "Contact $($AllSQLInstanceContact.TargetAddress) exists" -Level "Error"
        }
    } 
    
    Else
        {
        write-log -message "Check if the user has the wrong instance value"
        $CheckInstanceValue = (get-mailcontact $AllSQLInstanceContact.PrimarySMTPAddress).CustomAttribute1
        If($CheckInstanceValue -ne $Instance)
            {
            write-log -message "Updating the CustomAttribute1 for $($AllSQLInstanceContact.PrimarySMTPAddress) to $($instance)"
            set-mailcontact $AllSQLInstanceContact.PrimarySMTPAddress -CustomAttribute1 $instance
            }
        Else
            {
            Write-Log -Message "$($AllSQLInstanceContact.PrimarySMTPAddress) still exists in Office 365"
            }
        }





    }


    
#write-host "New Contacts are"
#$NewContacts 

#write-host "Updated Contacts are"
#$UpdatedContacts  


#write-host "Deleted Contacts are"
#$DeletedContacts 

#Clear all sessions
#get-PSSession | Remove-PSSession
Stop-Transcript