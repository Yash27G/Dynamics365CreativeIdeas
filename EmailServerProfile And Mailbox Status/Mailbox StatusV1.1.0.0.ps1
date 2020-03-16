﻿##_______________This Script check the status of mailbox in Dynamics 365 CRM instance________________##
##_____________________________________Author : YASH GUPTA (YG)______________________________________##
##________________________________Company Name : Sopra Steria India__________________________________##
##____________________________________Execution Time : 1 min 58sec____________________________________##

#Checking Required Module
$PackageArray = "CredentialManager","Microsoft.Xrm.Data.Powershell"
foreach($i in $PackageArray){
    if(Get-Module -ListAvailable -Name $i){
        }
    else{
        Install-Module -Name $i -Force
        }
    }

#Saving required file.
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$LogFilePath =  $scriptPath+"\MailboxStatus.log"
$DataFilepath = $scriptPath+"\Mailbox.csv"

#This function is used to send email at higher priority of failed mailbox details to the User.
function Send-EmailToNotifyError{
    param(
        [string] $loc1,
        [string] $EmailTo,
        [string] $cc,
        [string] $SMTPServer, #Specify the Server Name here
        [string] $SMTPPort =25
        )
    try{
        #Subject, Body and SMTP Details
        $Subject = (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Alert: Mailbox Status failure !!"

        $StoredSMTPCredential = Get-StoredCredential -Target SMTPLogin -AsCredentialObject
        If (!$StoredSMTPCredential)
        {
            $psCred = Get-Credential -Message "Enter UserName and Password"
            $cre = New-StoredCredential -Target SMTPLogin -Credentials $psCred -Persist LocalMachine
            $StoredSMTPCredential = Get-StoredCredential -Target SMTPLogin
            }
        $SMTPAuthUsername = $StoredSMTPCredential.UserName
        $SMTPAuthPassword = $StoredSMTPCredential.Password

        $EmailFrom = $SMTPAuthUsername
        $Body = "Hello User,
        Please find Attachment of "+(Get-Date).ToString('MM-dd-yyyy hh:mm:ss') +" Mailbox Status failure.

        NOTE: This is an autogenerated mail. Please do not reply.

        Thanks and Regards,
        Administrator"
        $mailmessage = New-Object system.net.mail.mailmessage
        $mailmessage.from = ($EmailFrom)
        $mailmessage.To.add($EmailTo)
        $mailmessage.Subject = $Subject
        $mailmessage.Body = $Body
        $mailmessage.Priority = [System.Net.Mail.MailPriority]::High
        $attachment = New-Object System.Net.Mail.Attachment($loc1)
        $mailmessage.Attachments.Add($attachment)
        if($cc){
            $mailmessage.CC.add($cc)
            }
        $SMTPClient = New-Object Net.Mail.SmtpClient("$SMTPServer", "$SMTPPort")
        $SMTPClient.Credentials = New-Object System.Net.NetworkCredential("$SMTPAuthUsername", "$SMTPAuthPassword")
        $SMTPClient.Send($mailmessage)
        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Email Send Successfully to " +$EmailTo | Add-Content $LogFilePath
        "      " | Add-Content $LogFilePath
        }
    catch{
        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " An Error Has Ocuured : "+$_.Exception.Message | Add-Content $LogFilePath
        }

<#
.SYNOPSIS
 
This function is used to send Email if anything goes wrong in mailbox.
 
.DESCRIPTION
 
Author : Yash Gupta(YG).
This function is automatically triggered by Get-MailboxDetail.
This function is a parameterised function.
You should pass arguments to the function using Named parametered.
The credentials will be saved in Window's Credential Manager.
 
.OUTPUTS
 
Send Email

#>
    }


#This function is used get last tested date of mailbox.
function Get-MailboxLastTestDate{
    PARAM(
        [string] $MailboxName
        )
    $result = Get-CrmRecords -EntityLogicalName mailbox -FilterAttribute name -FilterOperator eq $MailboxName -Fields testmailboxaccesscompletedon
    $value = foreach($j in $result.CrmRecords){$j.testmailboxaccesscompletedon}
    return $value

<#
.SYNOPSIS
 
This function is used to get last test date and time of mailbox.
 
.DESCRIPTION
 
Author : Yash Gupta(YG).
This function is tiggered by Get-MailboxDetail

#>
    }


#This function is used get the status of mailbox and save the details in csv file.
function Get-MailboxDetail{
    PARAM(
        [string]$URL,
        [string]$EmailTo,
        [string]$ServerProfileName,
        [string]$MailboxName
        )
    try{
        $StoredCredential = Get-StoredCredential -Target PowershellLogin
        If(!$StoredCredential) {
            $psCred = Get-Credential -Message "Enter your Credentials"
            $strcred = New-StoredCredential -Target PowershellLogin -Credentials $psCred -Persist LocalMachine
            $StoredCredential = Get-StoredCredential -Target PowershellLogin
            }
        $trgtUserName = $StoredCredential.UserName
        $trgtPass = $StoredCredential.Password
        $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $trgtUserName, $trgtPass

        #________________________Connecting CRM______________________
        $trgtCRMOrg = Connect-CrmOnline -Credential $cred -ServerUrl $URL
        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Connected to Dynamics 365 " +$URL | Add-Content $LogFilePath

        $result = Get-CrmRecords -EntityLogicalName mailbox -FilterAttribute name -FilterOperator eq $MailboxName -Fields name,mailboxid
        $value = foreach($j in $result.CrmRecords){$j.name;$j.mailboxid.Guid}

        $Oldvalue = Get-MailboxLastTestDate -MailboxName $j.name

        #Approve the Email
        Set-CrmRecord -conn $trgtCRMOrg -EntityLogicalName mailbox -Id $j.mailboxid.Guid -Fields @{isemailaddressapprovedbyo365admin=$true}
        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Email is approved" | Add-Content $LogFilePath

        #Test and enabling the mailbox
        Set-CrmRecord -conn $trgtCRMOrg -EntityLogicalName mailbox -Id $j.mailboxid.Guid -Fields @{testemailconfigurationscheduled=$true}
        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " The Mailbox has been Tested and Enabled. Please wait... " | Add-Content $LogFilePath
        Start-Sleep 1

        $Newvalue = Get-MailboxLastTestDate -MailboxName $j.name

        #Checking last tested date and time for mailbox
        while($Oldvalue.Equals($Newvalue)){
            Start-Sleep 4
            $Newvalue = Get-MailboxLastTestDate -MailboxName $j.name
            }

        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " The Mailbox is Tested and Enabled." | Add-Content $LogFilePath

 #Fetch XML on Entity Mailbox Details
$fetchXml = @"
<fetch>
  <entity name="mailbox" >
    <attribute name="incomingemailstatus" />
    <attribute name="outgoingemailstatus" />
    <attribute name="averagetotalduration" />
    <attribute name="mailboxstatus" />
    <attribute name="name" />
    <attribute name="emailserverprofile" />
    <attribute name="mailboxid" />
    <attribute name="statuscode" />
    <attribute name="emailrouteraccessapproval" />
    <attribute name="isemailaddressapprovedbyo365admin" />
    <attribute name="testemailconfigurationscheduled" />
    <attribute name="testmailboxaccesscompletedon" />
    <filter>
      <condition attribute="name" operator="eq" value="$MailboxName" />
    </filter>
  </entity>
</fetch>
"@

        $FetchResult = Get-CrmRecordsByFetch -conn $trgtCRMOrg -Fetch $fetchXml -AllRow
            $CheckServerProfile = foreach($i in $FetchResult.CrmRecords){
                if(!($i.emailserverprofile).Equals($ServerProfileName)){
                    [System.Windows.MessageBox]::Show("The mailbox " +$MailboxName+ " is not present in " +$ServerProfileName+"",'No Result Found','Ok','warning')
                    exit
                    }
                }
        
        

        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Mailbox Status has been retrieved from  " +$URL | Add-Content $LogFilePath

        #Creating Tablubar format to store in CSV File
        $tableEntity = New-Object system.Data.DataTable "MailboxStaus"
        $tblcol1 = New-Object system.Data.DataColumn Date, ([string])
        $tblcol2 = New-Object system.Data.DataColumn Name, ([string])
        $tblcol3 = New-Object system.Data.DataColumn Incoming_Email_Status, ([string])
        $tblcol4 = New-Object system.Data.DataColumn Outgoing_Email_Status, ([string])
        $tblcol5 = New-Object system.Data.DataColumn Status_Code, ([string])
        $tblcol6 = New-Object system.Data.DataColumn Mailbox_Status, ([string])
        $tblcol7 = New-Object system.Data.DataColumn Average_Total_Duration, ([string])
        $tblcol8 = New-Object system.Data.DataColumn Email_Router_Access_Approval, ([string])
        $tblcol9 = New-Object system.Data.DataColumn IsEmailAddressApprovedbyo365Admin, ([string])
        $tblcol10 = New-Object system.Data.DataColumn Test_Emailconfigurationscheduled, ([string])
        $tblcol11 = New-Object system.Data.DataColumn Test_Mailboxaccesscompletedon, ([string])
        $tableEntity.columns.add($tblcol1)
        $tableEntity.columns.add($tblcol2)
        $tableEntity.columns.add($tblcol3)
        $tableEntity.columns.add($tblcol4)
        $tableEntity.columns.add($tblcol5)
        $tableEntity.columns.add($tblcol6)
        $tableEntity.columns.add($tblcol7)
        $tableEntity.columns.add($tblcol8)
        $tableEntity.columns.add($tblcol9)
        $tableEntity.columns.add($tblcol10)
        $tableEntity.columns.add($tblcol11)
        $FetchResult = Get-CrmRecordsByFetch -conn $trgtCRMOrg -Fetch $fetchXml -AllRows
        #$solutionEntity = New-Object System.Collections.Generic.List[Guid]

        $FetchResult.CrmRecords | ForEach-Object{
            $tblrow = $tableEntity.NewRow()
            $tblrow.Date = (Get-Date).ToString('MM-dd-yyyy hh:mm:ss')
            $tblrow.Name = $_.name
            $tblrow.Incoming_Email_Status = $_.incomingemailstatus
            $tblrow.Outgoing_Email_Status = $_.outgoingemailstatus
            $tblrow.Status_Code = $_.statuscode
            $tblrow.Mailbox_Status = $_.mailboxstatus
            $tblrow.Average_Total_Duration = $_.averagetotalduration
            $tblrow.Email_Router_Access_Approval = $_.emailrouteraccessapproval
            $tblrow.IsEmailAddressApprovedbyo365Admin = $_.isemailaddressapprovedbyo365admin
            $tblrow.Test_Emailconfigurationscheduled = $_.testemailconfigurationscheduled
            $tblrow.Test_Mailboxaccesscompletedon = $_.testmailboxaccesscompletedon
            $tableEntity.Rows.Add($tblrow)

            $FetchResult = Get-CrmRecordsByFetch -conn $trgtCRMOrg -Fetch $fetchXml -AllRow
            $result = foreach($i in $FetchResult.CrmRecords){
                $i.name
                $i.mailboxid.Guid
                $i.testemailconfigurationscheduled
                $i.testmailboxaccesscompletedon
                }
            }

        #Exporting Data in CSV File
        $tableEntity | Export-Csv -NoTypeInformation -Path $DataFilepath -Append
        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " The Mailbox Status has been exported." | Add-Content $LogFilePath
        "           " | Add-Content $DataFilepath

        if($tblrow.Name -eq $MailboxName){
            if(($tblrow.Incoming_Email_Status -eq "Failure") -or ($tblrow.Outgoing_Email_Status -eq "Failure") -or ($tblrow.Status_Code -eq "Inactive") -or ($tblrow.Mailbox_Status -eq "Not Run") -or ($tblrow.Mailbox_Status -eq "Failure") -or ($tblrow.Email_Router_Access_Approval -eq "Empty") -or ($tblrow.IsEmailAddressApprovedbyo365Admin -eq "No") -or ($tblrow.Outgoing_Email_Status -eq "Not Run") -or ($tblrow.Incoming_Email_Status -eq "Not Run")){
                (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Something went wrong in mailbox. Please Check your mail for more detailed error of mailbox " | Add-Content $LogFilePath
                Send-EmailToNotifyError -loc1 $DataFilepath -EmailTo $EmailTo
                }
            else{
                #send and receive email to mailbox in order to test functionality
                }
            }
        }
    catch{
        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " An Error has occured" +$_.Exception.Message| Add-Content $LogFilePath
        "      " | Add-Content $LogFilePath
        }

<#
.SYNOPSIS
 
This function is used to get Status Mailbox.
 
.DESCRIPTION
 
Author : Yash Gupta(YG).
This function triggered by Get-EmailServerProfile
This function is a parameterised function.
You should pass arguments to the function using Named parametered.
The credentials will be saved in Window's Credential Manager.
 
.OUTPUTS
 
Create CSV file in same folder where script is saved. Send Email to given Email Id.
 
.EXAMPLE
 
Get-MailboxLastTestDate -URL <Environment URL> -EmailTo <Email Address> -ServerProfileName <Server Profile Name> -MailboxName <Mailbox Name>
 
.EXAMPLE
 
Get-SystemSettingFromSource -URL "https:\\example.crm8.dynamics.com" -EmailTo "Someone@domain.com" -ServerProfileName "Server Name" -MailboxName "Mailbox"
 
#>
 }


#This function is used get details of Email Server profile and notify if it is Inactivate.
function Get-EmailServerProfile{
    PARAM(
        [string]$ServerProfileName,
        [string]$URL,
        [string]$EmailTo
        )
    try{
    $StoredCredential = Get-StoredCredential -Target PowershellLogin
    If(!$StoredCredential) {
        $psCred = Get-Credential -Message "Enter your Credentials"
        $strcred = New-StoredCredential -Target PowershellLogin -Credentials $psCred -Persist LocalMachine
        $StoredCredential = Get-StoredCredential -Target PowershellLogin
        }
    $trgtUserName = $StoredCredential.UserName
    $trgtPass = $StoredCredential.Password
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $trgtUserName, $trgtPass

    #________________________Connecting CRM______________________
    $trgtCRMOrg = Connect-CrmOnline -Credential $cred -ServerUrl $URL
    (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Connected to Dynamics 365 " +$URL | Add-Content $LogFilePath

$fetchXml = @"
<fetch>
<entity name="emailserverprofile" >
<attribute name="name" />
<attribute name="statecode" />
<attribute name="statuscode" />
<filter>
<condition attribute="name" operator="eq" value="$ServerProfileName" />
</filter>
</entity>
</fetch>
"@
    $FetchResult = Get-CrmRecordsByFetch -conn $trgtCRMOrg -Fetch $fetchXml -AllRow
    $result = foreach($i in $FetchResult.CrmRecords){
        $i.name
        $i.statecode
        $i.statuscode
        }
    if($i.statecode.Equals("Inactive") -and $i.statuscode.Equals("Inactive")){
        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Email Server Profile is Inactivated" | Add-Content $LogFilePath
        [System.Windows.MessageBox]::Show('Email Server profile is not Active','Not Active','Ok','warning')
        break
        }
    else{
        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Email Server Profile is Activated" | Add-Content $LogFilePath
        $MialboxName = Read-Host "Which Mailbox's status are you looking for?"
        Get-MailboxDetail -MailboxName $MialboxName -URL $URL -EmailTo $EmailTo -ServerProfileName $ServerProfileName
        }

    (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Function Get-EmailServerProfile run successfully"| Add-Content $LogFilePath
    }
    catch{
        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " An Error has occured" +$_.Exception.Message| Add-Content $LogFilePath
        "      " | Add-Content $LogFilePath
    }
    
<#
.SYNOPSIS
 
This function is used to get Status Email Server Profile and linked Mailbox and will notify if anything goes wrong.
 
.DESCRIPTION
 
Author : Yash Gupta(YG).
This function is a parameterised function.
You should pass arguments to the function using Named parametered.
The credentials will be saved in Window's Credential Manager.
 
.OUTPUTS
 
Ask for mailbox name, Create CSV file in same folder where script is saved, Send Email to given Email Id.
 
.EXAMPLE
 
Get-EmailServerProfile -ServerProfileName <Server Profile Name> -URL <Environment URL> -EmailTo <Email Id>
 
.EXAMPLE
 
Get-EmailServerProfile -ServerProfileName "Server Name" -URL "https:\\example.crm8.dynamics.com" -EmailTo "Someone@domain.com"
 
#>
}


#This function is used to test incoming and outgoing mail to and from Mailbox ID respectively.
function Test-Mailbox{
    PARAM(
        [string]$MailboxId,
        [string]$TestEmailId,
        [int]$PortForTestEmailId,
        [string]$SMTPServerForTestEmailId,
        [int]$PortForMailboxId,
        [string]$SMTPServerForMailboxId
    )
    try{
        $ToMailbox = $host.ui.PromptForCredential("Need credentials", "Please enter your user name and password for TestEmail ID.", "", "NetBiosUserName")
        Send-MailMessage -From $TestEmailId -Subject "Test Subject to check incoming" -To $MailboxId -Body "Test Body From PowerShell" -Credential $ToMailbox -Port $PortForTestEmailId -SmtpServer $SMTPServerForTestEmailId
        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Email Send Successfully to " +$MailboxId+ " from " +$TestEmailId | Add-Content $LogFilePath

        $FromMailbox = $host.ui.PromptForCredential("Need credentials", "Please enter your user name and password for Mailbox ID.", "", "NetBiosUserName")
        Send-MailMessage -From $MailboxId -Subject "Test Subject to check outgoing" -To $TestEmailId -Body "Test Body From PowerShell" -Credential $FromMailbox -Port $PortForMailboxId -SmtpServer $SMTPServerForMailboxId
        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Email Send Successfully to " +$TestEmailId+ " from " +$MailboxId | Add-Content $LogFilePath
        "      " | Add-Content $LogFilePath
    }
    catch{
        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " An Error has occured" +$_.Exception.Message| Add-Content $LogFilePath
        "      " | Add-Content $LogFilePath
    }
}
