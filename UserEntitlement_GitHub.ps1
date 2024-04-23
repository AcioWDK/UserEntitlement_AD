<#
Script by @Aciowdk

Get all Users in $searchBase with hire date <2 weeks (extensionAttribute14)
Creates or overwrites LandingZoneUsers.csv -> saves to desktop -> opens it
requires -module ActiveDirectory

To customize:

CTRL - F and search for:    ####################### CHANGE ####################### 



#>


#                                 Location of users - filter - generate pw

Add-Type -AssemblyName PresentationFramework

#Check if run as admin function
function Check-IsElevated{

  $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()

  $p = New-Object System.Security.Principal.WindowsPrincipal($id)

  if ($p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)){ 
    return $true; 
  }else{
    return $false;
  }  

}



# Install RSAT if not installed

if (!(Get-Module -ListAvailable -Name ActiveDirectory)){
  if (Check-IsElevated) {
    Start-Process powershell -ArgumentList "-noexit"," Write-Host 'Unlocking New Level: RSAT - Active Directory' -ForegroundColor Green; Get-WindowsCapability -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' -Online | Where-Object { '$_.State' -ne 'Installed' } | Add-WindowsCapability -Online; Get-Module -ListAvailable -Name ActiveDirectory; Write-Host 'Congratulation, New Level Unlocked!' -ForegroundColor Green; Write-Warning 'You can now run the application without elevated rights.'; pause; Write-Host 'Terminating process...' -ForegroundColor Red; Start-Sleep 5; exit;"
    exit
  }else{
    Start-Process powershell -ArgumentList ' -command Write-Host " Run as Administrator to install RSAT! " -foregroundcolor red; pause; '
      exit
  }
}



                                           ####################### CHANGE #######################

$server = "aciowdk.domain"
$searchBase = "OU=WorkdayLandingZone,OU=Users,OU=_GlobalObjects,OU=_Organisation,DC=aciowdk,DC=domain"

#For e-mail sending
$emailBodyTemplate =  @'

<html>
    <style>
        p,ul,li {
            margin: 0;
        } 
        </style>

    <body lang=EN-US link="#0563C1" vlink="#954F72">

    <br>
        
    <p>Dear #firstname# #lastname#,</p>
        
    <br>
        
    <p>The account for your new Employee has been successfully
    created and these are the following credentials that #employeeFirstName# has to login
    with: </p>
    
    <br>
    
    <p><i>Windows Account: <b>#employeeID#</b></i><span
    style='font-size:9.0pt;font-family:"Helvetica",sans-serif;color:black'><o:p></o:p></span></p>
    
    <p><i>Windows Account Password: <b>#employeePw#</b><o:p></o:p></i></p>
    
    <p><i><span lang=IT style='mso-ansi-language:IT'>E-Mail ID</span></i><span
    lang=IT style='mso-ansi-language:IT'>:<span class=MsoHyperlink><i><span
    style='color:#4472C4;text-decoration:none;text-underline:none'>&nbsp;</span></i></span><i><span
    style='color:#4472C4'> </span></i></span><i><a
    href="mailto:#employeeMail#"><span lang=DE style='mso-ansi-language:
    DE'>#employeeMail#</span></a></i><i><span lang=DE
    style='mso-ansi-language:DE'><o:p></o:p></span></i></p>
    
    <p><i><span lang=DE style='mso-ansi-language:DE'><o:p>&nbsp;</o:p></span></i></p>
    
    <p><i><span lang=DE style='mso-ansi-language:DE'><o:p>&nbsp;</o:p></span></i></p>
    
    <p><b><span style='font-size:14.0pt'>Additional information:<o:p></o:p></span></b></p>
    
    <p><b>Your IT ServiceDesk Team </b></p>
    
    <p><o:p>&nbsp;</o:p></p>
    


    </body>

</html>

'@
$emailSubject = "New IT account - Onboarding"
# E-mails are automatically sent to the User's manager that is set in AD
$emailCC = "ITaccountsManagement@aciowdkdoamin.com;"


                                           ######################################################


$DesktopPath = [Environment]::GetFolderPath("Desktop")

$filterBy = 'enabled -eq "true" -and mail -like "*" -and samaccountname -notlike "*2*"'
$twoWeeks = (Get-Date).AddDays(+14).ToString("yyyy-MM-dd")
$randomPassword = @{Name='Password'; Expression={(Get-Date -Format "ddMMMM")+(Get-Random -Minimum 0 -Maximum 9999).ToString('0000')+"!"}}

#                                            User Properties

$manager = @{Name='Manager';Expression={(Get-ADUser -server $server ($_.Manager)).SAMAccountname}}
$managerDN = @{Name='Manager Location';Expression={(Get-ADUser -server $server ($_.Manager)).distinguishedname.Split(',')[6,5,4] -replace 'OU=','' -join ' | '}}
$hireDate = @{Name='HireDate'; Expression={[datetime]::parseexact(-join ($_.extensionAttribute14)[0..7], 'yyyyMMdd', $null).ToString('dd/MM/yyyy')}}
$samAccount = @{Name='Windows ID'; Expression={$_.SamAccountName}}
$mail = @{Name='E-Mail ID'; Expression={$_.mail}}
$moveTo = @{Name='ouName'; Expression={(Get-ADUser -server $server ($_.Manager)).distinguishedname -replace '^.+?(?<!\\),',''}}



$properties = @(
    'Enabled',
    'EmployeeID',
    'SamAccountName',
    'mail',
    'Title',                       
    'manager',
    'Country',
    'City',
    'GivenName',
    'extensionAttribute14'
  )


$customProperties = @(
    'Enabled',
    'EmployeeID',
    $samAccount,
    $randomPassword,
    $mail,
    'Title',                        
    $manager,
    $managerDN,
    'Country',
    'City',
    $hireDate,
    'GivenName',
    $moveTo
  )



#                                           Where magic happens

function sendManagerMail{

#runs on execute
begin {
  Write-Host "`n================================`n" -ForegroundColor Cyan
}

#runs for each object passed trought pipeline
process{


# HTML Onboarding Mail Format                             
                                           
$body = $emailBodyTemplate


# Prep to replace in HTML body
  try{
    
    # Avoid Breaking on missing mananger
    try {
      #Get manager for later use
      $currentManager = Get-ADUser ($_.manager) -Properties mail, givenname, Surname,DistinguishedName -server $server
      $user = $_.Name
    }
    catch {
      Write-Host "Manager Not Found" -ForegroundColor Red
      [reflection.assembly]::loadwithpartialname('System.Windows.Forms')
      [reflection.assembly]::loadwithpartialname('System.Drawing')
      $notify = new-object system.windows.forms.notifyicon
      $notify.icon = [System.Drawing.SystemIcons]::Information
      $notify.visible = $true                             
                                           ####################### CHANGE #######################
      $notify.showballoontip(10,'Manager Not Found for','Contact Aciowdkdomain',[system.windows.forms.tooltipicon]::None)
      $currentManager = ''
    }

    # If manager missing stop
    if(-not ($currentManager)){
      Write-Host "Creating Draft Mail" -ForegroundColor Cyan
      $parameters = @{
        <#change#>firstname = "ManagerName" 
        <#change#>lastname = "Here"      
        employeeFirstName = $_.GivenName
        employeeID = $_.'Windows ID'   
        employeeMail = $_.'E-Mail ID'
        employeePw = $_.Password
      }
  
      # Replace in HTML body
      foreach ($pair in $parameters.GetEnumerator()) {
        $body = $body -replace "#$($pair.Key)#",$pair.Value
      }
  
      $OL = New-Object -ComObject outlook.application
      $mItem = $OL.CreateItem(0)
  
                                           

      # Mail details
      $mItem.Subject = $emailSubject
      $mItem.CC = $emailCC
      $mItem.Importance = 2
      $mItem.HTMLBody = $body
  

      $inspector = $mItem.GetInspector
      $inspector.Display()

      # <#change#> $mItem.Send() 
  
    }else{
   
    $parameters = @{
        firstname = $currentManager.GivenName
        lastname = $currentManager.Surname
        employeeFirstName = $_.GivenName
        employeeID = $_.'Windows ID'   
        employeeMail = $_.'E-Mail ID'
        employeePw = $_.Password
    }

    # Replace in HTML body
    foreach ($pair in $parameters.GetEnumerator()) {
      $body = $body -replace "#$($pair.Key)#",$pair.Value
    }

    # Create and send the mail to each manager
    Write-Host  "Sending email to:" $currentManager.Mail  -ForegroundColor Blue

    $OL = New-Object -ComObject outlook.application
    $mItem = $OL.CreateItem("olMailItem")




    # Mail details
    $mItem.Subject = $emailSubject
    $mItem.To = $currentManager.mail
    $mItem.CC = $emailCC
    $mItem.HTMLBody = $body
    $mItem.Importance = 2
  
    $mItem.Send()
    
    Write-Host "Successfully sent" -ForegroundColor Green

    [reflection.assembly]::loadwithpartialname('System.Windows.Forms')
    [reflection.assembly]::loadwithpartialname('System.Drawing')
    $notify = new-object system.windows.forms.notifyicon
    $notify.icon = [System.Drawing.SystemIcons]::Information
    $notify.visible = $true  
    $notify.showballoontip(10,'Success','Mail sent to ' + $currentManager.mail + ' Successfully sent',[system.windows.forms.tooltipicon]::None)

  }

    Start-Sleep 1
    
  }
  catch{

  Write-Host "FAIL" -ForegroundColor Red
  [reflection.assembly]::loadwithpartialname('System.Windows.Forms')
  [reflection.assembly]::loadwithpartialname('System.Drawing')
  $notify = new-object system.windows.forms.notifyicon
  $notify.icon = [System.Drawing.SystemIcons]::Information
  $notify.visible = $true  
  $notify.showballoontip(10,'FAILED','Mail sent to ' + $currentManager.mail + ' failed',[system.windows.forms.tooltipicon]::None)
  
  }

}

#runs when 'process' finished
end{
 Write-Host "`nDone" -ForegroundColor Green
 [reflection.assembly]::loadwithpartialname('System.Windows.Forms')
 [reflection.assembly]::loadwithpartialname('System.Drawing')
 $notify = new-object system.windows.forms.notifyicon
 $notify.icon = [System.Drawing.SystemIcons]::Information
 $notify.visible = $true  
 $notify.showballoontip(10,'DONE','Script has finished',[system.windows.forms.tooltipicon]::None)
}

}


try {

  #=============-Send mails from .txt-=============  
  # 
  # $PATH = .\users.txt - just and example
  # $Users1 = Get-Content $PATH
  # $Users2 = @()
  # foreach ($users in $users1) {
  #   $users2 += @(Get-ADUser $Users -Properties $properties | Select-Object $customProperties)
  # }
  # $Users = $Users2 | Sort-Object -Property hireDate
  # 
  #================================================

  #================================================  
  #        Send mails from Landinzone - default
  #================================================
  
  $Users = @()
  $Users = Get-ADUser -Filter $filterBy -SearchBase $searchBase  -server $server  -Properties $properties| where{([datetime]::parseexact(-join ($_.extensionAttribute14)[0..7], 'yyyyMMdd', $null).ToString('yyyy-MM-dd')) -lt $twoWeeks} | Select-Object $customProperties
  $Users = $Users | Sort-Object -Property hireDate
  $count = 0
  $count = ($users | Measure-Object).count


  if ($count -eq 0) {
    $caption = "Yeeey" 
    $message = "No newHires today"
    $continue = [System.Windows.MessageBox]::Show($message, $caption, 'Ok','Information', 'none')
    Write-Host "Terminating process..." -ForegroundColor Red
    exit
  }


  #====================== Create ADM import csv==========================  
  #          
  #======================================================================
  # 
  # #check for import csv and refresh it
  # if(Test-Path($DesktopPath+"\LandingZoneUsers_To_Import.csv")){
  #   Remove-Item $DesktopPath"\LandingZoneUsers_To_Import.csv"
  # }
  # 
  # # Test OU path to move to
  # foreach ($userToImport in $Users){
  #   try {
  #     [string] $ouPath = $userToImport.ouName
  #   if (-not ([adsi]::Exists("LDAP://$ouPath"))) {
  #     $userToImport.ouName = ''
  #   }
  # 
  #   $toImport = New-Object -Type PSObject -Property @{
  #     'sAMAccountName'            = $userToImport.'Windows ID'
  #     'password'                  = $userToImport.Password
  #     'ouName'                    = $userToImport.ouName
  #   }
  #   $toImport | Export-Csv -Path $DesktopPath"\LandingZoneUsers_To_Import.csv" -Append -NoTypeInformation
  #   $count++
  #   }
  #   catch {
  #     write-host "no manager i guess" -ForegroundColor red
  #   }
  # 
  # }
  #================================================

  #Create the csv with all the info - on desktop
  $Users | Export-Csv -Path $DesktopPath"\LandingZoneUsers.csv" -NoTypeInformation
      
  #Open spearsheet
  Invoke-Item $DesktopPath"\LandingZoneUsers.csv"
  
  
  #========================= Check missing manager ===========================
  $managerNotFound = $false
  $missingManagers = @()

  $users | %{ if(!$_.manager)
                  {$managerNotFound=$true
                   $missingManagers +="`n"+$_.'Windows ID'
                  }
            }
  
  if($managerNotFound)
  {
    Write-Warning "Missing Managers for: $missingManagers"
    $caption = "Managers not set" 
    $message = "Please set the manager for: `n" + $missingManagers
    $continue = [System.Windows.MessageBox]::Show($message, $caption, 'Ok','Information', 'none')
    Write-Host "Terminating process..." -ForegroundColor Red
    Start-Sleep 5
    exit
  }
  


  #================== SendMail Confirmation Pop-up ===============

  
  $caption = "SENDING ------ $count ------ E-MAILS"    
  $message = "Are you Sure You Want To Proceed:"
  $continue = [System.Windows.MessageBox]::Show($message, $caption, 'YesNo')
  
  if ($continue -eq "No") {

    Write-Host "Terminating process..." -ForegroundColor Red
    Start-Sleep 1

  }else {
    
    #Send the emails
    $Users | sendManagerMail
    
  }



}
catch {
  [reflection.assembly]::loadwithpartialname('System.Windows.Forms')
  [reflection.assembly]::loadwithpartialname('System.Drawing')
  $notify = new-object system.windows.forms.notifyicon
  $notify.icon = [System.Drawing.SystemIcons]::Information
  $notify.visible = $true
  $notify.showballoontip(10,'We Talked About This...','To write new data.... LandingZone.csv must be closed and run again',[system.windows.forms.tooltipicon]::None)
}




