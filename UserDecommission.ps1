Clear-Variable -Name body

#ServiceNow REST API, reaches out to the Request table associated with the decommission ticket and pulls information about the user, including name, user account number etc.
$cred = Get-Credential
$SNbody = @{'sysparm_limit'=500;'sysparm_display_value'='true'}
$ritmNumber = Read-Host -Prompt "Enter request number for User Decommission"
$SNbody.sysparm_query = "number=$ritmNumber"
$ritm=Invoke-RestMethod -Method GET -Uri "SERVICENOW_REST_TABLE" -Credential $cred -Body $SNbody
$req = Invoke-RestMethod -Method GET -Uri $ritm[0].result.request.link -Credential $cred
$sn_user = Invoke-RestMethod -Method GET -Uri $req[0].result.requested_for.link -Credential $cred

#Creates a PowerShell object for each user based on data pulled from the ServiceNow REST request
$user = New-Object -TypeName PSObject
  $user | Add-Member -MemberType NoteProperty -Name "Name" -Value $sn_user.result.name
  $user | Add-Member -MemberType NoteProperty -Name "First Name" -Value $sn_user.result.first_name
  $user | Add-Member -MemberType NoteProperty -Name "Last Name" -Value $sn_user.result.last_name
  $user | Add-Member -MemberType NoteProperty -Name "UA" -Value $sn_user.result.u_user_account

#Excel 90 Day Counter, creates a COM Object to interact with an Excel spreadsheet used to track user decomissions
$90DayPath = 'SPREADSHEET_PATH'
$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open($90DayPath)
$Page = $Excel.Worksheets.item('Sheet1')
$Page.activate()

$FullRow = $Page.UsedRange.rows.count
$NewRow = $Page.Cells.Item($FullRow+2,1) = $user.Name
$NewRow = $Page.Cells.Item($FullRow+2,2) = Get-Date -Format 'M/dd/yyyy'
$NewMonth = $Page.Cells.Item($FullRow+2,3) = (Get-Date).AddMonths(3).ToString('M/dd/yyyy')
$NewRow = $Page.Cells.Item($FullRow+2,6) = $ritmNumber
$Workbook.Save()
$Workbook.Close()
Stop-Process -Name 'EXCEL' -Force

Write-Host ' '
Write-Host -ForegroundColor Blue -BackgroundColor White 'User added to 90 Day Delete spreadsheet'
Write-Host ' '

#PageGate Check, reaches out to paging VM and checks a text file of recipients that is checked daily 
$pgCheck=(Get-ChildItem -Path "NETWORK_PATH_ARCHIVE" -Filter "*recipients*.txt" | Sort-Object -Property LastWriteTime -Descending)[0]
$pgRecipients = Get-Content $pgCheck.FullName
$firstName = $user.'First Name'
$lastName = $user.'Last Name'
$pgUser = $pgRecipients | Where-Object { $_ -like ("*$lastName*_*$firstName*")}

if($pgUser){
  Write-host ' '
  Write-Host -ForegroundColor Red -BackgroundColor white $user.'First Name'  $user.'Last Name'  'Exists in PageGate'
  Write-host ' '
}else{
  Write-Host ' '
  Write-Host -ForegroundColor Blue -BackgroundColor white $user.'First Name'  $user.'Last Name'  'is not in PageGate'
  Write-Host ' '
}

#Automic User Check, creates COM Object to interact with Excel spreadsheet
$uc4Check ="\\cds_1\corpdata\IS_DBA\Shared\Oracle\Applications\Atomic\automic.xlsx"
$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open($uc4Check)
$10prod = $Excel.Worksheets.item('10prod')
$12prod = $Excel.Worksheets.item('12prod')
$10dev = $Excel.Worksheets.item('10dev')
$12dev = $Excel.Worksheets.item('12dev')
$10prod.activate()
$12prod.activate()
$10dev.activate()
$12dev.activate()
#add samAccountName to the Oracle DB Query to search on
$userFound10P = ($10prod.Cells.Find($lastName)-and $10prod.Cells.Find($firstName))
$userFound12P = ($12prod.Cells.Find($lastName)-and $12prod.Cells.Find($firstName))
$userFound10D = ($10dev.Cells.Find($lastName)-and $10dev.Cells.Find($firstName))
$userFound12D = ($12dev.Cells.Find($lastName)-and $12dev.Cells.Find($firstName))

#Email Body Array, writes to the host and adds the user to an email template if found in any version of automation orchestration tool.
$body = @()
$body += "$ritmNumber <br>"
$body += "$($lastName, $firstName) <br>"
if($userFound10P -ne $False){
  $body += 'User is in v10 Production <br>'
  Write-Host ' '
  Write-Host -ForegroundColor Red -BackgroundColor white $firstName $lastName 'Exists in Version 10 Production'
  Write-host ' '
 }else{
    Write-Host ' '
    Write-Host -ForegroundColor Blue -BackgroundColor white $firstName $lastName 'is not in Automic Version 10 Production'
    Write-Host ' '
}
  if($userFound12P -ne $False){
    $body += 'User is in v12 Production <br>'
    Write-Host ' '
    Write-Host -ForegroundColor Red -BackgroundColor white $firstName $lastName 'Exists in Version 12 Production'
    Write-host ' '
  }else{
    Write-Host ' '
    Write-Host -ForegroundColor Blue -BackgroundColor white $firstName $lastName 'is not in Automic Version 12 Production'
    Write-Host ' '
 }
  if($userFound10D -ne $False){
    $body += 'User is in v10 Development <br>'
    Write-Host ' '
    Write-Host -ForegroundColor Red -BackgroundColor white $firstName $lastName 'Exists in Version 10 Development'
    Write-host ' '
  }else{
    Write-Host ' '
    Write-Host -ForegroundColor Blue -BackgroundColor white $firstName $lastName 'is not in Automic Version 10 Development'
    Write-Host ' '
  }
  if($userFound12D -ne $False){
    $body += "User is in v12 Development <br>"
    Write-Host ' '
    Write-Host -ForegroundColor Red -BackgroundColor white $firstName $lastName 'Exists in Version 12 Development'
    Write-host ' '
  }else{
    Write-Host ' '
    Write-Host -ForegroundColor Blue -BackgroundColor white $firstName $lastName 'is not in Automic Version 12 Development'
    Write-Host ' '
  }

#Finalize and convert email to HTML
$body += 'This is an automated email'
$body | ConvertTo-HTML
$body = $body | Out-String
$Workbook.Close()
Stop-Process -Name 'EXCEL' -Force

#PowerShell email parameters
$mailParams=@{
To = "EMAIL"
From = "EMAIL"
CC = "EMAIL"
Subject = "Action Requested: $ritmNumber - Automic User Decommission"
SMTPServer = 'MAIL_SERVER'
Body = $body
BodyAsHTML = $True
Credential = $cred
}

#if user found in any instance, send email
if($userFound10P -or $userFound12P -or $userFound10D -or $userFound12D -ne $False){
  Send-MailMessage @mailParams
} 