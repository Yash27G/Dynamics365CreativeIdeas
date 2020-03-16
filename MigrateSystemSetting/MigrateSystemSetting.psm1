##____This Script migrating the system setting from source and setting in target instance of Dynamics 365 CRM_____##
##___________________________________________Author : YASH GUPTA (YG)_____________________________________________##
##_______________________________________Company Name : Sopra Steria India________________________________________##
##___________________Total Execution Time : 1min 52sec(Get) + 2min 39sec(Set) = 4min 31sec________________________##


function Get-SystemSettingFromSource{
# .ExternalHelp MigrateSystemSetting.help.xml
    [CmdletBinding(SupportsShouldProcess)]
    PARAM(
     [parameter(Mandatory=$true)]
     [string]$SrcURL,
     [parameter(Mandatory=$true)]
     [string]$DataFilepath,
     [parameter(Mandatory=$true)]
     [string]$LogFilePath,
     [parameter(Mandatory=$true)]
     [string]$ErrorLogFilePath
    )
try{

    #Checking Required Module
    $PackageArray = "CredentialManager","Microsoft.Xrm.Data.Powershell"
    foreach($i in $PackageArray){
        if(Get-Module -ListAvailable -Name $i){
        }
        else {
            Install-Module -Name $i -Force
      }
    }

    #User Logging Credentials
    $StoredCredential = Get-StoredCredential -Target PowershellLogin
    If(!$StoredCredential)
        {
            $PsCred = Get-Credential -Message "Enter your Credentials"
            $storcred = New-StoredCredential -Target PowershellLogin -Credentials $PsCred -Persist LocalMachine
            $StoredCredential = Get-StoredCredential -Target PowershellLogin
        }
    $TrgtUserName = $StoredCredential.UserName
    $TrgtPass = $StoredCredential.Password
    $Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $TrgtUserName, $TrgtPass

    #__________Connecting to Source CRM____________
    $SrcCRMOrg = Connect-CrmOnline -Credential $Cred -ServerUrl $SrcURL
    (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Connected to Source environment " + $SrcURL | Add-Content $LogFilePath

    #Saving attributes of organization in csv file
    $OrgAttribute=Get-CrmEntityAttributes organization | Where-Object {($_.IsValidForUpdate -eq $true) -and ($_.AttributeType -ne "virtual") -and ($_.AttributeType -ne "Uniqueidentifier") -and ($_.AttributeType -ne "memo") -and ($_.LogicalName -ne "TokenKey") -and ($_.LogicalName -ne "referencesitemapxml")} | Select-Object logicalname,AttributeType
    $OrgAttribute | Export-Csv -NoTypeInformation -Path $DataFilepath

    $objExcel = New-Object -ComObject Excel.Application
    $WorkBook = $objExcel.Workbooks.Open($DataFilepath)
    $Sheet = $WorkBook.sheets.item(1)
    $row = 1
    $Sheet.Cells.Item($row,4)= 'Source name'
    $Sheet.Cells.Item($row+1,4)= $SrcCRMOrg.ConnectedOrgFriendlyName
    $Sheet.Cells.Item($row+1,4).Font.Size = 14
    $Sheet.Cells.Item($row+1,4).Font.Bold = $True
    $Sheet.Cells.Item($row,3)= 'AttributeValue'

    #saving attributes value in same csv file
    foreach($i in $OrgAttribute){
        $logicalname = $i.LogicalName
        $OrgAttributeValueResult = (Get-CrmRecords -conn $SrcCRMOrg -Entitylogicalname organization -Fields $logicalname,organizationid).CrmRecords
        foreach($j in $OrgAttributeValueResult){
            $row+=1
            $Sheet.Cells.Item($row,3)= $j.$logicalname
        }
}
(Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " The System Setting is saved in csv file of instance " + $SrcURL + " in " +$DataFilepath | Add-Content $LogFilePath

 #Exiting csv file
 $objExcel.DisplayAlerts = 'False'
 $objExcel.ActiveWorkbook.SaveAs($DataFilepath)
 $workbook.Close()
 Start-Sleep 2
 $objExcel.Quit()
 (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Get-SystemSettingFromSource is successfully completed " | Add-Content $LogFilePath
 }
catch{
    Start-Sleep 2
    $objExcel.Quit()
    (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " An Error Has Ocuured " | Add-Content $LogFilePath
    "   ########   " | Add-content $LogFilePath
    $_.Exception.Message | Out-File $ErrorLogFilePath
    "   ########   " | Add-content $ErrorLogFilePath
 }
}



function Set-SystemSettingInTarget{
# .ExternalHelp MigrateSystemSetting.help.xml
    [CmdletBinding(SupportsShouldProcess)]
    PARAM(
        [parameter(Mandatory=$true)]
        [string]$TrgtURL,
        [parameter(Mandatory=$true)]
        [string]$DataFilepath,
        [parameter(Mandatory=$true)]
        [string]$LogFilePath,
        [parameter(Mandatory=$true)]
        [string]$ErrorLogFilePath
    )
try{

    #User Logging Credentials
    $StoredCredential = Get-StoredCredential -Target PowershellLogin
    If(!$StoredCredential)
        {
            $PsCred = Get-Credential -Message "Enter your Credentials"
            New-StoredCredential -Target PowershellLogin -Credentials $PsCred -Persist LocalMachine
            $StoredCredential = Get-StoredCredential -Target PowershellLogin
        }
    $TrgtUserName = $StoredCredential.UserName
    $TrgtPass = $StoredCredential.Password
    $Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $TrgtUserName, $TrgtPass
    #Connecting to Target Environment
    $TrgtCRMOrg = Connect-CrmOnline -Credential $Cred -ServerUrl $TrgtURL
    (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Connected to target environment " + $TrgtURL | Add-Content $LogFilePath

    $OrgResult = (Get-CrmRecords -conn $TrgtCRMOrg -Entitylogicalname organization -Fields organizationid).CrmRecords
    $OrgGuid=foreach($i in $OrgResult){$i.organizationid.Guid}

    $ObjExcel = New-Object -ComObject Excel.Application
    $WorkBook = $ObjExcel.Workbooks.Open($DataFilepath)
    $Sheet = $WorkBook.sheets.item(1)
    $ObjRange2 = $Sheet.Range("B2")
    [void] $ObjRange2.Sort($ObjRange2)
    $WorkBook.Save()

    #searching Attribute value in csv file
    $SearchText = "AttributeValue"
    $Found = $Sheet.Cells.Find($SearchText)
    if($Found){
       $row = $Found.row + 1
       $column = $Found.column
    }

    $WorksheetRange = $Sheet.UsedRange
    $RowCount = $WorksheetRange.Rows.Count

    Start-Transaction

    for($i=0; $i -le ($RowCount+1); $i++) {
        $DataType = $Sheet.Cells.Item($row,$column-1).Text
        $SettingName = $Sheet.Cells.Item($row,$column-2).Text  
        if(($Sheet.Cells.Item($row,$column).Text).Equals('')){
            $SettingValue = $Sheet.Cells.Item($row,$column).Text
        }
        else{
            $SettingValue = $Sheet.Cells.Item($row,$column).value()
        }

        Switch($DataType){

            #Setting interger values
            'Integer' {
                $cv=[int]$SettingValue
                Set-CrmRecord -conn $TrgtCRMOrg -EntityLogicalName organization $orgguid @{$SettingName=$cv}
                (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " " + $SettingName +" is updated in target instance" | Add-Content $LogFilePath
                $row+=1
                continue
            }

            #Setting String values
            'String' {if($SettingName.Equals('contentsecuritypolicyconfiguration')){
                        $row+=1
                        continue
                        }
                $cv=[string]$SettingValue
                Set-CrmRecord -conn $TrgtCRMOrg -EntityLogicalName organization $orgguid @{$SettingName=$cv}
                (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " " + $SettingName +" is updated in target instance" | Add-Content $LogFilePath
                $row+=1
                continue
                }

            #Setting Boolean values
            'Boolean' {If($SettingValue.Equals('Yes')){
                        Set-CrmRecord -conn $TrgtCRMOrg -EntityLogicalName organization $orgguid @{$SettingName=$true}
                        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " " + $SettingName +" is updated in target instance" | Add-Content $LogFilePath
                        }
                If($SettingValue.Equals('No') -or $SettingValue.Equals('Disable')){
                    Set-CrmRecord -conn $TrgtCRMOrg -EntityLogicalName organization $orgguid @{$SettingName=$false}
                    (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " " + $SettingName +" is updated in target instance" | Add-Content $LogFilePath
                    if(($SettingName.Equals('syncoptinselection')) -and ($SettingValue.Equals('Disable'))){
                        $syncoptinselectionValue = $SettingValue
                        $syncoptinselectionName = $SettingName
                    }
                }
                If($SettingValue.Equals('')){
                    Set-CrmRecord -conn $TrgtCRMOrg -EntityLogicalName organization $orgguid @{$SettingName=$null}
                    (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " " + $SettingName +" has no Value" | Add-Content $LogFilePath
                }
                $row+=1
                continue
                }

            #Setting DateTime values
            "DateTime" {
                $cv=[datetime]$SettingValue
                Set-CrmRecord -conn $TrgtCRMOrg -EntityLogicalName organization $orgguid @{$SettingName=$cv}
                (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " " + $SettingName +" is updated in target instance" | Add-Content $LogFilePath
                $row+=1
                continue
                }

            #Setting Picklist values
            "Picklist"{
                if(($SettingName.Equals('syncoptinselectionstatus')) -and ($syncoptinselectionValue.Equals('Disable')) -and ($syncoptinselectionName.Equals('syncoptinselection'))){
                    (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " " + $SettingName +" is diabled in source instance" | Add-Content $LogFilePath
                    $row+=1
                    continue
                }
                else{
                    $fg=Get-CrmEntityOptionSet -EntityLogicalName organization -conn $TrgtCRMOrg -FieldLogicalName $SettingName | Select-Object -Property Items
                    $picklistvalue = $fg.Items | Where-Object {($_.DisplayLabel -eq $SettingValue)}
                    $optionSetValue = New-Object -TypeName  Microsoft.Xrm.Sdk.OptionSetValue -ArgumentList $picklistvalue.PickListItemId
                    Set-CrmRecord -conn $TrgtCRMOrg -EntityLogicalName organization $orgguid @{$SettingName=$optionSetValue}
                }
                (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " " + $SettingName +" is updated in target instance" | Add-Content $LogFilePath
                $row+=1
                continue
                }

            #Setting Lookup values
            "Lookup"{
                if($SettingName.Equals('defaultemailserverprofileid')){
                    if($SettingValue.Equals('')){
                        Set-CrmRecord  -EntityLogicalName organization -id $OrgGuid -Fields @{$SettingName=$null}
                        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " " + $SettingName +" has no Value" | Add-Content $LogFilePath
                    }
                else{
                    $emailserverprofileresult = (Get-CrmRecords -conn $TrgtCRMOrg -Entitylogicalname emailserverprofile -FilterAttribute name -FilterOperator eq -FilterValue $SettingValue -Fields emailserverprofileid).CrmRecords
                    $emailguid = foreach($j in $emailserverprofileresult){$j.emailserverprofileid.Guid}
                    Set-CrmSystemSettings -conn $TrgtCRMOrg -DefaultEmailServerProfileId $emailguid
                    (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " " + $SettingName +" is updated in target instance" | Add-Content $LogFilePath
                    }
                }
                if($SettingName.Equals('acknowledgementtemplateid')){
                    if($SettingValue.Equals('')){
                        Set-CrmRecord  -EntityLogicalName organization -id $OrgGuid -Fields @{$SettingName=$null}
                        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " " + $SettingName +" has no Value" | Add-Content $LogFilePath
                    }
                    else{
                        $templateresult = (Get-CrmRecords -conn $TrgtCRMOrg -Entitylogicalname template -FilterAttribute title -FilterOperator eq -FilterValue $SettingValue -Fields templateid).CrmRecords
                        $tempid = foreach($k in $templateresult){$k.templateid.Guid}
                        Set-CrmSystemSettings -conn $TrgtCRMOrg -AcknowledgementTemplateId $tempid
                        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " " + $SettingName +" is updated in target instance" | Add-Content $LogFilePath
                        }
                }
                if($SettingName.Equals('defaultmobileofflineprofileid')){
                    if($SettingValue.Equals('')){
                        Set-CrmRecord  -EntityLogicalName organization -id $OrgGuid -Fields @{$SettingName=$null}
                        (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " " + $SettingName +" has no Value" | Add-Content $LogFilePath
                        }
                }
                $row+=1
                continue
                }
        }
  }
    Complete-Transaction

    (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " Set-SystemSettingInTarget is successfully completed " | Add-Content $LogFilePath
    "   ########   " | Add-content $LogFilePath
    #Exiting csv file
    $ObjExcel.DisplayAlerts = 'False'
    $ObjExcel.ActiveWorkbook.SaveAs($DataFilepath)
    $workbook.Close()
    Start-Sleep -Seconds 2
    $ObjExcel.Quit()
}
 catch{
    (Get-Date).ToString('MM-dd-yyyy hh:mm:ss') + " An Error Has Ocuured " | Add-Content $LogFilePath
    Undo-Transaction
    Start-Sleep -Seconds 2
    $ObjExcel.Quit()
    "   ########   " | Add-content $LogFilePath
    $_.Exception.Message | Out-File $ErrorLogFilePath
    "   ########   " | Add-content $ErrorLogFilePath
    }

}

