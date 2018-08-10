<#

    AWS CloudWatch Metric Extractor , Requires Module --> https://github.com/dfinke/ImportExcel

    Collecting metric / state data from AWS CloudWatch api and output's to CSV, run the script without parameters for instructions.

    Script requires Powershell v3 and above.

    Run the command below to store AccessKey and SecretKey in secure credential XML for each environment

        $cred = Get-Credential
        $cred | Export-Clixml -Path "d:\AWS\config\us-east-1.xml"

    Run the command below to store user and password which will login to the internet proxy.

        $cred = Get-Credential
        $cred | Export-Clixml -Path "d:\AWS\config\proxy.xml"


#>

param
(
    [String]$CollectionType = 'Collection',
    [String]$AWSMetricNameSpace = 'AWS/EC2',
    [String]$AWSRegion = 'us-east-1',
    [Array]$InstanceById = @('i-BleBlahBlue','i-BleBlahRed'),
    [Array]$Metrics = @('CPUUtilization','DiskReadBytes','DiskReadOps','DiskWriteBytes','DiskWriteOps','NetworkIn','NetworkOut','NetworkPacketsIn','NetworkPacketsOut'),
    [String]$rollUpType = 'Average',
    [String]$intervalType = 3600,
    [DateTime]$StartDate = (Get-date).addDays(-2),
    [DateTime]$EndDate = (Get-date),
    [String]$Format = 'XLS',
    [String]$Email,
    [String]$FileName, # = 'MyBoxes',
    [String]$OutputLocation, # = 'd:\Watcher\WebSSO\us-east-1',
    [String]$InstanceByIdLookupList # = 'd:\Watcher\Servers.csv'
)

#Params
$ScriptPath = 'D:\Wacther'
$LogFileLoc = $ScriptPath + '\Log\Logfile.log'
$RunDateTime = (Get-date)
$RunDateTime = $RunDateTime.tostring("yyyyMMddHHmmss") 
$StartDateFile = $StartDate.tostring("yyyyMMdd-HHmmss")            
$EndDateFile = $EndDate.tostring("yyyyMMdd-HHmmss")
$Share = '\\localhost\completed\'
$mailserver = 'mailserverbox.vMan.ch'
$mailport = 25
$ProxyCreds = 'MeProxy'
$ProxyHost = 'proxy.vMan.ch'
$ProxyPort = 8080

#Functions

#Logging Function
Function Log([String]$message, [String]$LogType, [String]$LogFile){
    $date = Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
    $message = $date + "`t" + $LogType + "`t" + $message
    $message >> $LogFile
}

#Send Email Function
Function SS64Mail($SMTPServer, $SMTPPort, $SMTPuser, $SMTPPass, $strSubject, $strBody, $strSenderemail, $strRecipientemail, $AttachFile)
   {
   [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { return $true }
      $MailMessage = New-Object System.Net.Mail.MailMessage
      $SMTPClient = New-Object System.Net.Mail.smtpClient ($SMTPServer, $SMTPPort)
	  $SMTPClient.EnableSsl = $true
	  $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTPuser, $SMTPPass)
      $Recipient = New-Object System.Net.Mail.MailAddress($strRecipientemail, "Recipient")
      $Sender = New-Object System.Net.Mail.MailAddress($strSenderemail, "vMan AWS Metrics")
     
      $MailMessage.Sender = $Sender
      $MailMessage.From = $Sender
      $MailMessage.Subject = $strSubject
      $MailMessage.To.add($Recipient)
      $MailMessage.Body = $strBody
      if ($AttachFile -ne $null) {$MailMessage.attachments.add($AttachFile) }
      $SMTPClient.Send($MailMessage)
   }


switch($CollectionType)
    {

Collection {

If ($InstanceByIdLookupList -gt ''){


Log -Message "Collecting $intervalType between $StartDateFile and $EndDateFile" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
echo "Collecting $intervalType between $StartDateFile and $EndDateFile"
Log -Message "Collecting metrics: $Metrics" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
echo "Collecting metrics: $Metrics"
Log -Message "Data collection for objects listed in the file $InstanceByIdLookupList" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
echo "Data collection for objects listed in the file $InstanceByIdLookupList"

$InstanceById = ''

[Array]$InstanceById = import-csv $InstanceByIdLookupList | where ENV -eq $AWSRegion | Select Server

[Array]$InstanceById = $InstanceById.Server

Log -Message "Data collection for $InstanceById" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
echo "Data collection for $InstanceById"

}

else {

Log -Message "Collecting $intervalType between $StartDateFile and $EndDateFile" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
echo "Collecting $intervalType between $StartDateFile and $EndDateFile"
Log -Message "Collecting metrics: $Metrics" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
echo "Collecting metrics: $Metrics"
Log -Message "Data collection for object: $InstanceById" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
echo "Data collection for object: $InstanceById"

}

#Set's the AWS Region to query
if($AWSRegion-gt ""){

    Set-DefaultAWSRegion -Region $AWSRegion

    }

    else {
        echo "AWS region not specified, bye!!"
        Exit
    }

#Sets the credentials to access the data.

if($AWSRegion -gt ""){

    $AWSCred = Import-Clixml -Path "$ScriptPath\config\$AWSRegion.xml"

    $AWSUser = $AWSCred.GetNetworkCredential().Username
    $AWSPassword = $AWSCred.GetNetworkCredential().Password

    Set-AWSCredentials -AccessKey $AWSUser -SecretKey $AWSPassword
    }
    else
    {
    Log -Message "AWS credentials not specified, bye!!" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    echo "AWS credentials not specified, bye!!"
    Exit
    }

#Check if any metrics specified
if($Metrics -eq ""){

        echo "No metrics specified, bye!!"
        Exit

    }

#Set the Internet Proxy if you cant get out directly the to net.

if($ProxyCreds -gt ""){

    $ProxyCred = Import-Clixml -Path "$ScriptPath\config\$ProxyCreds.xml"

    $ProxyUser = $ProxyCred.GetNetworkCredential().Username
    $ProxyPassword = $ProxyCred.GetNetworkCredential().Password

    Set-AWSProxy -Hostname $ProxyHost -Port $ProxyPort -Password $ProxyPassword -Username $ProxyUser
    }
    else
    {
    Log -Message "Proxy credentials not specified so not configuring proxy" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    echo "Proxy credentials not specified so not configuring proxy"
    }

if($Email -imatch '^.*@vMan\.ch$'){

    Log -Message "$email matches the vMan.ch domain" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    Echo "$email matches the vMan.ch domain"

    $SMTPcred = Import-Clixml -Path "$ScriptPath\config\smtp.xml"

    $SMTPUser = $SMTPcred.GetNetworkCredential().Username
    $SMTPPassword = $SMTPcred.GetNetworkCredential().Password
    }
    else
    {
    Log -Message "$email is not in the vMan.ch domain, will not send mail but report generation will continue" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    Echo "$email is not in the vMan.ch domain, will not send mail but report generation will continue"
	$Email = ''
    }

$report = @()

ForEach ($InstanceID in $InstanceById){

    echo "Running collection for instance: $InstanceID"
    Log -Message "Running collection for instance: $InstanceID" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    Get-Date -UFormat '%m-%d-%Y %H:%M:%S'

    $dimension = New-Object Amazon.CloudWatch.Model.Dimension

    $dimension.set_Name('InstanceId')
    $dimension.set_Value($InstanceID)

    ForEach ($Metric in $Metrics){

    $Data = @()
    echo "Collecting Metric $Metric for $InstanceID"
    Log -Message "Collecting Metric $Metric for $InstanceID" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    Get-Date -UFormat '%m-%d-%Y %H:%M:%S'

    $Data = Get-CWMetricStatistics -Namespace $AWSMetricNameSpace -dimension $dimension -MetricName $Metric -StartTime $StartDate -EndTime $EndDate -Period $intervalType -Statistics $rollUpType

            echo "Transforming Metric $Metric for $InstanceID"
            Log -Message "Transforming Metric $Metric for $InstanceID" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
            Get-Date -UFormat '%m-%d-%Y %H:%M:%S'

            If ($rollUpType -eq 'Average') {$values = @($Data.DataPoints.Average)}
            If ($rollUpType -eq 'Maximum') {$values = @($Data.DataPoints.Maximum)}
            If ($rollUpType -eq 'Minumum') {$values = @($Data.DataPoints.Minumum)}

            $Timestamps = @($Data.DataPoints.timestamp)

                for ($i=0; $i -lt $Values.Count -and $i -lt $Timestamps.Count; $i++) {
                    $report += New-Object PSObject -Property @{
                    Metric     = $Metric
                    InstanceId = $InstanceID
                    Timestamp  = $Timestamps[$i]
                    value      = $values[$i]
                    }

            }
    }
}

#Output merge to Excel
  if ($Format -eq 'XLS'){
    Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
    echo 'Generate XLS'
    Log -Message "Generate XLS" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    
    If ($FileName -gt ''){

    $OutputExcelFile = $OutputLocation + $FileName + '.xlsx'
    $ShareExcelFile = $OutputLocation + $FileName + '.xlsx'
    remove-item $OutputExcelFile -Force -ErrorAction SilentlyContinue -Recurse

    }
   else
   {

    $OutputExcelFile = $ScriptPath + '\completed\Collected_Metrics_' + $intervalType + '_' + [String]$rollUpType + '_' + [String]$StartDateFile + '_' + [String]$EndDateFile + '_' + $RunDateTime + '.xlsx'
    $ShareExcelFile = $share + 'Collected_Metrics_' + $intervalType + '_' + [String]$rollUpType + '_' + [String]$StartDateFile + '_' + [String]$EndDateFile + '_' + $RunDateTime + '.xlsx'  
   }

    $report | select Timestamp,InstanceId,Metric,value | Sort-Object { $_.Timestamp -as [datetime] } | Export-Excel $OutputExcelFile -WorkSheetname Data -ChartType Line -IncludePivotChart -IncludePivotTable -PivotRows Timestamp -PivotData value -PivotColumns InstanceId,Metric
    
    Log -Message "Task complete, pickup your file from the share $ShareExcelFile within the next 24 hours as it will be automatically deleted" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    Echo "Task complete, pickup your file from the share $ShareExcelFile within the next 24 hours as it will be automatically deleted"
    if($Email -gt "")
    {
        Log -Message "Email found and is vMan.ch domain, sending report to $email" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
        Echo "Email found and is vMan.ch domain, sending email to $email"
        SS64Mail $mailserver $mailport $SMTPUser $SMTPPassword "ENV $creds Report started at $RunDateTime complete" "Your report should be attached, if it's missing it was probably too large to send via the mail server.... Pickup it up from here $ShareExcelFile within the next 24H" 'info@vman.ch' $email $OutputExcelFile
        Log -Message "Email sent to $email" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
        Echo "Email sent to $email"
    }
    Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
    echo 'All done, Terminating Script'
    Log -Message "All done, Terminating Script" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    
    Remove-Variable *  -Force -ErrorAction SilentlyContinue
    Exit
  }


#Output to CSV
  if ($Format -eq 'CSV'){
    Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
    echo 'Generate CSV'
    Log -Message "Generate CSV" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc

    If ($FileName -gt ''){

    $OutputCSVFile = $OutputLocation + $FileName + '.csv'
    $ShareCSVFile = $OutputLocation + $FileName + '.csv'

    }
   else
   {

    $OutputCSVFile = $ScriptPath + '\completed\Collected_Metrics_' + $intervalType + '_' + [String]$rollUpType + '_' + [String]$StartDateFile + '_' + [String]$EndDateFile + '_' + $RunDateTime + '.csv'
    $ShareCSVFile = $share + 'Collected_Metrics_' + $intervalType + '_' + [String]$rollUpType + '_' + [String]$StartDateFile + '_' + [String]$EndDateFile + '_' + $RunDateTime + '.csv'
   
   }

    $report | select Timestamp,InstanceId,Metric,value | Sort-Object { $_.Timestamp -as [datetime] } | Export-csv $OutputCSVFile -NoTypeInformation
    Log -Message "Task complete, pickup your file from the share $ShareCSVFile within the next 24 hours as it will be automatically deleted" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    Echo "Task complete, pickup your file from the share $ShareCSVFile within the next 24 hours as it will be automatically deleted"
    if($Email -gt "")
    {
        Log -Message "Email found and is vMan.ch domain, sending CSV link to $email" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
        Echo "Email found and is vMan.ch domain, sending CSV link to $email"
        SS64Mail $mailserver $mailport $SMTPUser $SMTPPassword "ENV $creds Report started at $RunDateTime complete" "Your report is ready,  Pickup it up from here $ShareCSVFile within the next 24H" 'info@vman.ch' $email
        Log -Message "Email sent to $email" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
        Echo "Email sent to $email"
    }
    Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
    echo 'All done, Terminating Script'
    Log -Message "All done, Terminating Script" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    
    Remove-Variable *  -Force -ErrorAction SilentlyContinue
    Exit
  }

}
<#

#Debug, get list of Metrics
#$AWSCW = Get-CWMetrics

#>

default{"Usage

The script can be run by specifying all parameters otherwise it will use some default values.

.\CloudWatcher.ps1 -CollectionType 'Collection' -AWSRegion 'us-east-1' -StartDate '2016/09/16 10:00' -EndDate '2016/09/16 11:00' -Metrics 'CPUUtilization','DiskReadBytes','DiskReadOps','DiskWriteBytes','DiskWriteOps','NetworkIn','NetworkOut','NetworkPacketsIn','NetworkPacketsOut' -rollUpType 'Average' -intervalType '3600' -Format 'XLS' -FileName 'MyBoxes' -OutputLocation 'D:\Watcher\' -InstanceByIdLookupList 'D:\Watcher\Servers.csv'
        "}
}