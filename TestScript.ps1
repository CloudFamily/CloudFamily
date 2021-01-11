#############################################################
# RDS Managent tool for managing RDS environment            #
#                                                           #
# In some cases this tool requires Administrator privileges #
#                                                           #
#                                                           #
# Author: Karel de Reus  / Jan Willem Buiten / Wim Eling    #
#                                                           #
#                                                           #
#IMPORTANT always use nl Culture otherwise things go wrong!!!
[System.Threading.Thread]::CurrentThread.CurrentCulture = 'nl-NL'

$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

import-module ActiveDirectory

$DC = "SRVAD01"
$Profielpad = "\\domein.local\dfs$\tsprofiles\"
$ADusersource = "OU=1.Gebruikers ZGIJV,OU=Zorgcombinatie,DC=domein,DC=local" 
$ADUitdienstsource = "OU=Uit Dienst,OU=Zorgcombinatie,DC=domein,DC=local" 
$domain = "@zgijv.nl"
$naamvoorstelVariant = "Partner+Eigen"

#Office licentie keys
$Office365E1 = "ZGIJV:STANDARDPACK"
$Office365E3 = "ZGIJV:ENTERPRISEPACK"
$Office365F1 = "ZGIJV:DESKLESSPACK"
$OfficeVisio = "ZGIJV:VISIOCLIENT"

$OfficeUsername = "admin@zgijv.onmicrosoft.com"

#key created for adm_wim
#$OfficePassword = "01000000d08c9ddf0115d1118c7a00c04fc297eb010000004bfd2a7df7b83041bfe47c8593cfe7040000000002000000000003660000c00000001000000040b8a41f9a0f31b610abf47bbbe066920000000004800000a0000000100000009df7e8cae7fb0dee9514e3b1bbdfbd8e200000005b9701ae5b39e358bc70b16593bd368198cf5c38ddb23eff62b10e48e527c28e14000000eef9dca2ca6cbe3dc5b07e123e9028cee8629260" | ConvertTo-SecureString
#key created for SVC-Export-HR-AD
###$OfficePassword  = "01000000d08c9ddf0115d1118c7a00c04fc297eb0100000001f10bdc59798b41bb4d9cbdaf7b4ff90000000002000000000003660000c0000000100000001560f1624abda4c391ea43852df210320000000004800000a0000000100000006b78fd4848955e749e5a95135325f47f20000000cdc670fb7f8b7296159dc89fa8cf049ca76c4ac058026b349710898b767a982914000000dba376bb1e8289dfd79d0244ea7dea56d5fb0817" | ConvertTo-SecureString
#CSV 
$today = Get-Date;
$nrOfDaysInTheFuture = 5;
$daysInTheFuture = $today.AddDays($nrOfDaysInTheFuture).Date;

$sftpUsername = "ZGIJV_Intercept"
$sftpPassword = ConvertTo-SecureString "6rj784Vw" -AsPlainText -Force
$sftpRemotePath = "/"
$sftpServer = "sftp.mijnaag.nl"
$sftpPort = 22
$csvFilename = 'export_beaufort.csv';
$xmlFilename = 'export_beaufort.xml';
$sftpDownloadDirectory = "C:\scripts\Download ZGIJV"
$sftpDownloadDirectoryXML = "C:\scripts\Download ZGIJV\XML"

$global:accountsWithoutInformation = @();
$global:accountsToBeCreated = @();

#Office365
# install-module AzureAD
# install-module MSOnline
# import-module MSOnline
#import-Module -Name Posh-SSH
###import-Module "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\Posh-SSH-master\posh-ssh"
import-Module "C:\Program Files (x86)\WindowsPowerShell\Modules\Posh-SSH\2.1\posh-ssh"

#Email settings 
#In PWS you can find the login details of this Sendgrid account under the name SMTP Resource
###Install-Module -Name Posh-SSH -RequiredVersion 2.1$sendGridUsername = "azure_b3e7619415b0d075f3fce06bb6e11ca5@azure.com"
$sendGridPassword = ConvertTo-SecureString "DFGdf@iperjlkgfds4543_SFD" -AsPlainText -Force
$sendGridCredential = New-Object System.Management.Automation.PSCredential $sendGridUsername, $sendGridPassword
$sendGridSmtp = "smtp.sendgrid.net"

$CCEmailAddresses = @("QiC@zgijv.nl", "HR@zgijv.nl", "Applicatiebeheer-Aysist@zgijv.nl")
$AdminEmailAddress = "helpdesk@zgijv.nl"
$FromEmailAddress = "noreply@zgijv.nl"
$ManagerNrsWithPersonalEmail = @( "99199195" )


$emailTemplate = Join-Path $executingScriptDirectory "newUserEmailTemplate.txt"
$emailTemplateManagerNotFound = Join-Path $executingScriptDirectory "managerNotFoundEmailTemplate.txt" 

###########
# Functions
###########

function KRS_Loginoffice365 {
    param([Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credentials           )
    
    try {
        $office365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://"outlook.office365.com"/powershell-liveid/ -Credential $Credentials -Authentication Basic -AllowRedirection
        Connect-MsolService -Credential $Credentials
        Import-PSSession $office365Session
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        $emailBody = "Het inloggen op de Office 365 omgeving is mislukt via het In - Uitdienst script. Er zijn geen veranderingen aan office licenties doorgevoerd. Controleer de inlog informatie. Message: $ErrorMessage"
        SendEmail -Priority "High" -EmailTo $AdminEmailAddress -EmailFrom $FromEmailAddress -Subject "Inloggen Office 365 mislukt" -Body  $emailBody
    }
}

function Search-Users {
    param([Parameter(Mandatory = $false)][string]$search)
    
    if (!$search) {
        $office365users = Get-MsolUser -all | where {$_.isLicensed -eq $true} | Select BlockCredential, Department, DisplayName, Licenses, Title, UserPrincipalName , Enabled
    }       
    else {
        $office365users = Get-MsolUser -SearchString $search | where {$_.isLicensed -eq $true} 
    }
    
    $UsersearchO365 = @()
    foreach ($user in $office365users) {
        $userarray0365 = New-Object psobject
        $userarray0365 | Add-Member -type NoteProperty -Name "UserName" -Value $user.DisplayName
        $userarray0365 | Add-Member -type NoteProperty -Name "Functie" -Value $user.Title
        $userarray0365 | Add-Member -type NoteProperty -Name "Afdeling" -Value $user.Department
        $userarray0365 | Add-Member -type NoteProperty -Name "Uitgeschakeld" -Value $user.BlockCredential
        $userarray0365 | Add-Member -type NoteProperty -Name "Licentie" -Value $user.Licenses.AccountSkuId
        $userarray0365 | Add-Member -type NoteProperty -Name "UserPrincipalName" -Value $user.UserPrincipalName
        $UsersearchO365 += $userarray0365
    }
    return $UsersearchO365
    
    
}

function ConvertNullProperties {
    param([Parameter(Mandatory = $true)]$Elements)

    foreach ($Element in $Elements) {
        foreach ($object_property in $Element.PsObject.Properties) {
            if ($object_property.Value -eq "NULL" -or $object_property.Value -eq "null") {
                $object_property.Value = $null;
            }
        }
    }
}

function ConvertToDate {
    param([Parameter(Mandatory = $false)]$Date)

    if ($Date) {
        return Get-Date $Date;
    }
    return $Date;
}

#
# Send email via SendGrid
#
function SendEmail {
    param([Parameter(Mandatory = $true)][string]$EmailFrom,
        [Parameter(Mandatory = $true)][string]$EmailTo,
        [Parameter(Mandatory = $false)][string[]]$CC,
        [Parameter(Mandatory = $true)][string]$Subject,
        [Parameter(Mandatory = $true)][string]$Body,
        [Parameter(Mandatory = $false)][string]$Priority)

        if(!$Priority){
            $Priority = "Normal"
        }
        if ($CC) {
            Send-MailMessage -Priority $Priority -smtpServer $sendGridSmtp -Credential $sendGridCredential -Usessl -Port 587 -from $EmailFrom -to $EmailTo -Cc $CC -subject $Subject -Body $Body
        }
        else {
            Send-MailMessage -Priority $Priority -smtpServer $sendGridSmtp -Credential $sendGridCredential -Usessl -Port 587 -from $EmailFrom -to $EmailTo -subject $Subject -Body $Body
        }
}

function Read-AD-Gegevens {
    param(
        [Parameter(Mandatory = $true)][string]$ADSourceOU
    )

    $gebruikers = get-aduser -SearchBase "$ADSourceOU" -properties * -filter * | Select userPrincipalName, employeeID, accountExpires, proxyAddresses, cn , extensionAttribute1, extensionAttribute2 , sAMAccountName , Enabled 
    
    return $gebruikers
}

#
# Geef een naamvoorstel op basis van voornaam, achternaam en optioneel partner achternaam
#
function Get-Naam-voorstel {
    param(
        [Parameter(Mandatory = $true)][string]$Voornaam,
        [Parameter(Mandatory = $true)][string]$Eigenachternaam,
        [Parameter(Mandatory = $false)][string]$Partnerachternaam,
        [Parameter(Mandatory = $false)][string]$SamengesteldeNaam
    )

    $voorstel = @{}

    $Voornaam = Remove-Diacritics $Voornaam
    $Eigenachternaam = Remove-Diacritics $Eigenachternaam
    $Partnerachternaam = Remove-Diacritics $Partnerachternaam
    $SamengesteldeNaam = Remove-Diacritics $SamengesteldeNaam

    #$gebruikersnaam = ($Voornaam[0] + $Voornaam.Substring(1).ToLower() + $Eigenachternaam) 

    #gebruikersnaam mag nooit spaties bevatten
    $voorstel.Gebruikersnaam = $Voornaam[0] + ($Voornaam -replace " ", "").Substring(1).ToLower() + ($Eigenachternaam -replace '.+\s(.+)', '$1')[0]
    $voorstel.AlternativeGebruikersnaam = $Voornaam[0] + ($Voornaam -replace " ", "").Substring(1).ToLower() + ($Eigenachternaam -replace '.+\s(.+)', '$1').substring(0,2)
    $voorstel.AlternativeGebruikersnaam2 = $Voornaam[0] + ($Voornaam -replace " ", "").Substring(1).ToLower() + ($Eigenachternaam -replace '.+\s(.+)', '$1').substring(0,3)
    $samengesteld = ($SamengesteldeNaam -replace " ", "");

    $upnFirstElement = $Voornaam[0]  + "." + $samengesteld
    $voorstel.UserPrincipalName =  $upnFirstElement + "$domain" | ForEach-Object { $_.ToLower() }
    $voorstel.AlternativeUserPrincipalName = $Voornaam + "." + $samengesteld + "$domain" | ForEach-Object { $_.ToLower() }

    $voorstel.Mailadres = ""
    $voorstel.Weergavenaam = ""

    $voorstel.Mailadres = $Voornaam[0] + "." + ($SamengesteldeNaam -replace " ", "") + "$domain" | ForEach-Object { $_.ToLower() }

    $voorstel.Weergavenaam = $Voornaam + " " + $SamengesteldeNaam
    $voorstel.AlternativeWeergavenaam = $Voornaam + " " + $Partnerachternaam + " - " + $Eigenachternaam
    $ResultObj = (New-Object PSObject -Property $voorstel)
    return $ResultObj
}


function Remove-Diacritics 
{
  param ([String]$sToModify = [String]::Empty)

  foreach ($s in $sToModify) # Param may be a string or a list of strings
  {
    if ($sToModify -eq $null) {return [string]::Empty}

    $sNormalized = $sToModify.Normalize("FormD")

    foreach ($c in [Char[]]$sNormalized)
    {
      $uCategory = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c)
      if ($uCategory -ne "NonSpacingMark") {$res += $c}
    }

    return $res
  }
}

function Controleer-Gebruiker {
    param(
        [Parameter(Mandatory = $true)][string]$Personeelsnr,
        [Parameter(Mandatory = $true)][string]$Mailadres,
        [Parameter(Mandatory = $true)][string]$Weergavenaam
    )

    $Logging = $null
    $time = get-date -Format T 

    foreach ($checkuser in $global:gebruikers) {
        if ($checkuser.employeeID -match $Personeelsnr) { $Logging += "$time match persnr `n" } 
        if ($checkuser.proxyAddresses -match $Mailadres) { $Logging += "$time match mail `n" } 
        if ($checkuser.userPrincipalName -contains $Mailadres) { $Logging += "$time match UPN `n" }
        if ($checkuser.cn -contains $Weergavenaam) { $Logging += "$time match CN `n" }
        write-host $checkuser.proxyAddresses
    }
}


#
# Maak gebruiker aan in Active Directory en maak gebruiker lid van groep(en)
#
function Maak-gebruiker {
    param(
        [string]$UitDienstDatum,
        [Parameter(Mandatory = $true)][string]$Personeelsnr,
        [Parameter(Mandatory = $true)][string]$Gebruikersnaam,
        [Parameter(Mandatory = $true)][string]$Weergavenaam,
        [Parameter(Mandatory = $true)][string]$Voornaam,
        [Parameter(Mandatory = $true)][string]$Achternaam,
        [Parameter(Mandatory = $true)][string]$Mailadres,
        [Parameter(Mandatory = $true)][string]$AfdelingFunctieCode,
        [Parameter(Mandatory = $true)][string]$UserPrincipalName,
        [Parameter(Mandatory = $true)][string]$FunctieOmschrijving,
        [Parameter(Mandatory = $true)][string]$AfdelingCode
    )

    if (!$Personeelsnr) { $logging = "Vul Personeelsnummer in"}
    Else {

        $samaccname = $Gebruikersnaam

        $password = Maak-AD-gebruiker -DC $DC -Weergavenaam $Weergavenaam -Voornaam $Voornaam -Achternaam $Achternaam `
            -Samaccname $Gebruikersnaam -Mailadres $Mailadres -Personeelsid $Personeelsnr -Gebruikersou $ADusersource `
            -UitDienstDatum $UitDienstDatum -UserPrincipalName $UserPrincipalName
                
     
                # AfdelingFunctieCode bepaald welke rechten de user heeft op basis van functie en afdeling code
              $group =  Get-ADGroup -Identity  $AfdelingFunctieCode  
            

            if( $group)
            {
                    # AfdelingFunctieCode bepaald welke rechten de user heeft op basis van functie en afdeling code
                    Add-ADGroupMember -Identity  $AfdelingFunctieCode -Member $samaccname 
                
            }
            else {
                $emailBody = "De groep met afdelingscode $($AfdelingFunctieCode) kon niet worden gevonden in AD koppeling voor $($Gebruikersnaam) is mislukt functieomschrijving: $($FunctieOmschrijving)"
                SendEmail -Priority "High" -EmailTo $AdminEmailAddress -EmailFrom $FromEmailAddress -Subject "Groep is niet gevonden in AD" -Body  $emailBody
            }

            #stagaire leerling vakantiekracht
            if($AfdelingCode -ne "999" -and $AfdelingCode -ne "296" -and $AfdelingCode -ne "396")
            {
                Add-ADGroupMember -Identity  "AP-Nordined-Users" -Member $samaccname 
            }

            if($password.GetType() -eq [string]) {
                return $password
            }
            else {
               return $password[$password.Count-1]
            }

    }    
}

#
# Maak gebruiker aan in Active Directory
#
function Maak-AD-gebruiker {
    param(
        [Parameter(Mandatory = $true)][string]$DC,
        [Parameter(Mandatory = $true)][string]$Weergavenaam,
        [Parameter(Mandatory = $true)][string]$Voornaam,
        [Parameter(Mandatory = $true)][string]$Achternaam,
        [Parameter(Mandatory = $true)][string]$Samaccname,
        [Parameter(Mandatory = $true)][string]$Mailadres,
        [Parameter(Mandatory = $true)][string]$UserPrincipalName,
        [Parameter(Mandatory = $true)][string]$Personeelsid,
        [Parameter(Mandatory = $true)][string]$Gebruikersou,
        [Parameter(Mandatory = $false)][string]$UitDienstDatum
    )
 
    $Profilename = $Samaccname + ".domein" 
    $generatedPassword = GenerateCustomPassword
    $homeDirectory =  "\\domein.local\DFS$\Home\$Samaccname"

        $newAdUser = New-ADUser -Name $Weergavenaam  -GivenName $Voornaam -Surname $Achternaam `
            -SamAccountName $Samaccname -UserPrincipalName $UserPrincipalName -DisplayName $Weergavenaam `
            -accountPassword (ConvertTo-SecureString -AsPlainText $generatedPassword -Force) `
            -EmailAddress $UserPrincipalName -EmployeeID $Personeelsid `
            -Path $Gebruikersou -enable $True


            Set-ADUser -identity $Samaccname -add @{proxyAddresses = "SMTP:$UserPrincipalName"} -Replace @{HomeDrive = "H:"; HomeDirectory = $homeDirectory} -ChangePasswordAtLogon $True
            Set-ADUser -identity $Samaccname -Add @{extensionAttribute1=$UitDienstDatum}
            $user = [ADSI] "LDAP://CN=$Weergavenaam,$Gebruikersou"
            $user.psbase.Invokeset("terminalservicesprofilepath", "\\domein.local\dfs$\tsprofiles\$Profilename")
            $user.setinfo() 

            # check if folder is available if not then create
            If(!(test-path $homeDirectory))
            {
                New-Item -ItemType Directory -Force -Path $homeDirectory
            }
            
            $newAdUser = Get-ADUser -Filter {sAMAccountName -eq $Samaccname}
            
                # Always set FullControl access rights to a new user 
                $rule= new-object System.Security.AccessControl.FileSystemAccessRule ($Samaccname,"FullControl","ContainerInherit,ObjectInherit","None","Allow")
                $acl = Get-ACL -Path $homeDirectory
                $acl.SetAccessRule($rule)
                Set-ACL -Path $homeDirectory -AclObject $acl

    return $generatedPassword
}

$dagen = @("Zondag","Maandag","Dinsdag","Woensdag","Donderdag","Vrijdag","Zaterdag")

function GenerateCustomPassword {

    $randomDay = Get-Random -Minimum 0 -Maximum 6
    $randomNumber = Get-Random -Maximum 99
    $dayAsText = $dagen[$randomDay]
    return  "$($dayAsText)$($randomNumber)"
}

#
# Process .csv file
#
function ProcessCsvFile {
    param([Parameter(Mandatory = $true)][string]$FilePath)

    $CsvUsers = Import-Csv $FilePath"" -Delimiter ';' -Encoding UTF8

    ConvertNullProperties -Elements $CsvUsers;

    foreach ($CsvUser in $CsvUsers) {
        if ($CsvUser.indnst_dt) {
            $CsvUser.indnst_dt = ConvertToDate -Date $CsvUser.indnst_dt;
        }
        if ( $CsvUser.uitdnst_dt) {
            $CsvUser.uitdnst_dt = ConvertToDate -Date $CsvUser.uitdnst_dt;
        }
    }

    foreach ($CsvUser in $CsvUsers) {

        $oeKort =  $CsvUser."oe_kort,"
        $UserInfo = [PSCustomObject]@{ 
            Afdeling                = $CsvUser.func_oms
            AfdelingCode            = $CsvUser.func_kd
            AfdelingOmschrijving    = $CsvUser.oe_vol
            Personeelsnummer        = $CsvUser.pers_nr
            Voornaam                = Remove-Diacritics $CsvUser.roepnaam
            Achternaam              = Remove-Diacritics $CsvUser.geboorte_achternaam
            PriveEmailadres         = $CsvUser.prive_email
            PersoneelsnummerManager = $CsvUser.mngr_pers_nr
            AfdelingFunctieCode     = "$($oeKort)_$($CsvUser.func_kd)"
            SamenGesteldeNaam = Remove-Diacritics $CsvUser.naam_samengesteld
            GebruikAchternaam = $CsvUser."gebruik achternaam"
            UitDienstDatum = $null
            InDienstDatum = $CsvUser.indnst_dt.Date.ToString()
            NaamLeidinggevende = "$(Remove-Diacritics $CsvUser.mngr_e_roepnaam) $(Remove-Diacritics $CsvUser.mngr_naam_samen)"
        }

        # Stagaire met Welzijn in de afdelingomschrijving hoeven niet aangemaakt te worden.
        if($UserInfo.AfdelingCode -eq "999" -and $UserInfo.AfdelingOmschrijving.Contains("Welzijn"))
        {
            continue
        } 

        # Boventallige medewerkers niet meenemen
        if($CsvUser."oe_kort," -eq "11110")
        {
            continue
        }

        if($CsvUser.uitdnst_dt)
        {
            $UserInfo.UitDienstDatum  = $CsvUser.uitdnst_dt.Date.ToString()
        }

        # these properties are mandatory without this information no valid account can be created or deleted
        if (!$CsvUser.roepnaam -or !$CsvUser.geboorte_achternaam) {
            $global:accountsWithoutInformation += $UserInfo;
            continue
        }

        # Select all users that have a InDienst date and not a UitDienst date these are potential new users
        if (($CsvUser.indnst_dt) -and ($CsvUser.indnst_dt.Date -le $daysInTheFuture)) { #-and () ) {
            if(!($CsvUser.uitdnst_dt) -or ($CsvUser.uitdnst_dt -gt  $today.Date ))
            {
                $voorstel = Get-Naam-voorstel -Voornaam $CsvUser.roepnaam -Eigenachternaam $CsvUser.geboorte_achternaam -Partnerachternaam $CsvUser.partner_achternaam -SamengesteldeNaam $CsvUser.naam_samengesteld
                $UserInfo | Add-Member -Name 'Voorstel' -Type NoteProperty -Value $voorstel
                $global:accountsToBeCreated += $UserInfo;
            }
        }
    
        # Select all users with a UitDienst date then ad property must be updated
        if (($CsvUser.uitdnst_dt)) {
            $nrOfEmptyElements   = @($CsvUsers | Where-Object { $_.pers_nr -eq $CsvUser.pers_nr -and !$_.uitdnst_dt })
            if($nrOfEmptyElements.Count -eq  0)
            {
                $laatseUitdienstDatum =   @($CsvUsers | Where-Object { $_.pers_nr -eq $CsvUser.pers_nr } ) | Sort-Object uitdnst_dt -Descending | Select-Object -first 1
                
                $uitdienstDatum =  $laatseUitdienstDatum.uitdnst_dt.ToString()
           
                Get-ADUser  -Filter "EmployeeID -eq $($UserInfo.Personeelsnummer)" | Set-ADUser -Replace @{ extensionAttribute1=$uitdienstDatum}
                Write-Host "Uitdienst datum is aangepast voor employeeId:"  $UserInfo.Personeelsnummer
            }
            else {
                Write-Host "Uitdienst niet aangepast voor employeeId:"  $UserInfo.Personeelsnummer " Medewerker heeft andere actieve dienstverbanden"
            }
        }
        else {
            Get-ADUser  -Filter "EmployeeID -eq $($UserInfo.Personeelsnummer)" | Set-ADUser -Clear extensionAttribute1
            Write-Host "Uitdienst datum is geleegd voor employeeId:"  $UserInfo.Personeelsnummer
        }
    }
}

#
# SFTP functions
#
function CreateSession {
    param([Parameter(Mandatory = $true)][string]$ComputerName, [Parameter(Mandatory = $true)][string]$Username, [Parameter(Mandatory = $true)][SecureString]$Password,[string]$Port)
 
    $credentials = New-Object System.Management.Automation.PSCredential ($Username, $Password)
    $sftpSession = New-SFTPSession -ComputerName $ComputerName -Credential $credentials -Port $Port -AcceptKey -Verbose
    return $sftpSession
}

function DownloadFile {
    param([Parameter(Mandatory = $true)][int]$SessionId,
        [Parameter(Mandatory = $true)][string]$LocalPath,
        [Parameter(Mandatory = $true)][string]$RemotePath)
   
    Get-SFTPFile -SessionId $SessionId -LocalPath $LocalPath -RemoteFile $RemotePath -Overwrite 
}

function UploadFile {
    param([Parameter(Mandatory = $true)][int]$SessionId,
        [Parameter(Mandatory = $true)][string]$LocalPath,
        [Parameter(Mandatory = $true)][string]$RemotePath)
   
    Set-SFTPFile -SessionId $SessionId -LocalFile $LocalPath -RemotePath $RemotePath
}

function CloseSession {
    param([Parameter(Mandatory = $true)][int]$SessionId)

    $session = Get-SFTPSession -SessionId  $SessionId

    if ($session) {
        $session.Disconnect()
        $null = Remove-SftpSession -SftpSession $session
    }
}

function DownloadCsvFile {
    $strippedFileName = [System.IO.Path]::GetFileNameWithoutExtension($csvFilename);
    $extension = [System.IO.Path]::GetExtension($csvFilename);
    $csvNewFilename = $strippedFileName + "-" + [DateTime]::Now.ToString("yyyyMMdd-HHmmss") + $extension;

    $session = CreateSession -ComputerName $sftpServer -Username $sftpUsername -Password $sftpPassword -Port $sftpPort

    $remoteFilelocation = $sftpRemotePath + $csvFilename

    $localFileLocation = Join-Path $sftpDownloadDirectory $csvNewFilename

    $downloadedFile = Join-Path $sftpDownloadDirectory $csvFilename

    DownloadFile -SessionId $session.SessionId -LocalPath $sftpDownloadDirectory -RemotePath $remoteFilelocation 

    type $downloadedFile -Encoding:String | Out-File $localFileLocation -Encoding:UTF8

    Remove-Item -Path $downloadedFile

    CloseSession -SessionId $session.SessionId
    return $localFileLocation;
}


###Added by Gregor Suttie 21/12/2020

###START

###$xmlFileLocation = DownloadXmlFile

function DownloadXmlFile {
    $strippedFileName = [System.IO.Path]::GetFileNameWithoutExtension($xmlFilename);
    $extension = [System.IO.Path]::GetExtension($xmlFilename);
    $xmlNewFilename = $strippedFileName + "-" + [DateTime]::Now.ToString("yyyyMMdd-HHmmss") + $extension;

    $session = CreateSession -ComputerName $sftpServer -Username $sftpUsername -Password $sftpPassword -Port $sftpPort

    $remoteFilelocation = $sftpRemotePath + $xmlFilename

    $localFileLocation = Join-Path $sftpDownloadDirectory $xmlFilename

    $downloadedFile = Join-Path $sftpDownloadDirectory $xmlFilename

    DownloadFile -SessionId $session.SessionId -LocalPath $sftpDownloadDirectoryXML -RemotePath $remoteFilelocation 

    type $downloadedFile -Encoding:String | Out-File $localFileLocation -Encoding:UTF8

    Remove-Item -Path $downloadedFile

    CloseSession -SessionId $session.SessionId
    return $localFileLocation;
}

function ConvertXMLFileToCSV
{
    param([Parameter(Mandatory = $true)][string]$FilePath)

    # Here we need to grab the xml file contents and convert this to a csv file format
    [xml]$xmlContent = get-content $FilePath

    [System.Xml.XmlDocument] $xd = new-object System.Xml.XmlDocument
    $xd.load($FilePath)
    $employeelist = $xd.selectnodes("/data/company/employees/employee") # XPath is case sensitive
    $array = @()
    
    foreach ($employee in $employeelist) {

      $pers_nr = $employee.employeeid

      if($employee.salutationid -eq 'MW')
      {
        $dv_vlgnr = 1
      }
      elseif($employee.salutationid -eq 'DHR')
      {
        $dv_vlgnr = 2
      }
      else
      {
        $dv_vlgnr = 3
      }

      $roepnaam = $employee.nickname
      $naam_samengesteld = $employee.birthname
      $geboorte_voorvoegsels = $employee.prefixbirthname
      $geboorte_achternaam = $employee.birthname
      $partner_voorvoegsels = $employee.prefixpartnername
      $partner_achternaam = $employee.partnername
      $gebruik_achternaam =  $employee.nameusage
      $werk_telefoon = ''

      #Phone Numbers
      $mobiel_list = $employee.SelectNodes("phonenumbers/phonenumber") 

      if ($mobiel_list)
      {
        $mobiel_telefoon = $null
        #We have some phone numbers now try to get the phone number where phonetypeid = 3 i.e. mobile number
        foreach ($mobiel in $mobiel_list)
        {
            # Attempt to ge the private mobile number
            if ($mobiel.phonetypeid -eq 3)
            {
                $mobiel_telefoon = $mobiel.phoneno
            }
        }
      }
      else
      {
        $mobiel_telefoon = $null
      }

      #Emails
      $emails_list = $employee.SelectNodes("emails/email") 

      $werk_email = $null
      $prive_email = $null
      if ($emails_list)
      {
        #We have some emails now try to get the work email
        foreach ($email in $emails_list)
        {
            # Attempt to get email addresses - 1 = home, 2 = work
            if ($email.emailtypeid -eq 2)
            {
                $werk_email = $email.emailaddress
            }
            elseif ($email.emailtypeid -eq 1)
            {
                $prive_email = $email.emailaddress
            }
        }
      }
      else
      {
        $werk_email = $null
        $prive_email = $null
      }

      $Locatie = '' #TODO WE DOONT HAVE

      $oe_kort = $null
      $oe_kort = $employee.selectSingleNode("contracts/contract/subcontracts/subcontract/departments/department/orgunit").InnerText
      
      $oe_vol = $null
      $oe_vol = $employee.selectSingleNode("contracts/contract/subcontracts/subcontract/departments/department/orgunitname").InnerText

      $func_kd = $null
      $func_kd = $employee.selectSingleNode("contracts/contract/subcontracts/subcontract/functions/function/functionid").InnerText

      $func_oms = $null
      $func_oms = $employee.selectSingleNode("contracts/contract/subcontracts/subcontract/functions/function/functionname").InnerText

      #Format dates here into dd/mm/yyyy
      $indnst_dt = $null
      $indnst_date = $employee.selectSingleNode("contracts/contract/subcontracts/subcontract/functions/function/validfrom")
      if ($indnst_date)
      {
        $indnst_dt =  [System.DateTime]::ParseExact($indnst_date.'#text','yyyyddmm',$null).ToString('dd-MM-yyyy') 
      }
      else
      {
        $indnst_dt = ''
      }

      $uitdnst_date = $employee.selectSingleNode("contracts/contract/subcontracts/subcontract/functions/function/validuntil").InnerText
      if ($uitdnst_date.Value)
      {
        $uitdnst_dt = [System.DateTime]::ParseExact($uitdnst_date.'#text','yyyyddmm',$null).ToString('dd-MM-yyyy')
      }
      else
      {
        $uitdnst_dt = ''
      }

      $mngr_pers_nr = ''     #Not within the XML example
      $mngr_e_roepnaam = ''  #Not within the XML example
      $mngr_vrvg_samen = ''  #Not within the XML example
      $mngr_Ing_datum = ''   #Not within the XML example
      $mngr_Eind_datum = ''  #Not within the XML example

      $customobject = New-Object psobject
      
      $customobject | Add-Member -MemberType NoteProperty -Name pers_nr -Value  $pers_nr
      $customobject | Add-Member -MemberType NoteProperty -Name dv_vlgnr -Value $dv_vlgnr
      $customobject | Add-Member -MemberType NoteProperty -Name roepnaam -Value $roepnaam
      $customobject | Add-Member -MemberType NoteProperty -Name naam_samengesteld -Value  $naam_samengesteld
      $customobject | Add-Member -MemberType NoteProperty -Name geboorte_voorvoegsels -Value $geboorte_voorvoegsels
      $customobject | Add-Member -MemberType NoteProperty -Name geboorte_achternaam -Value $geboorte_achternaam
      $customobject | Add-Member -MemberType NoteProperty -Name partner_voorvoegsels -Value $partner_voorvoegsels
      $customobject | Add-Member -MemberType NoteProperty -Name partner_achternaam -Value $partner_achternaam
      $customobject | Add-Member -MemberType NoteProperty -Name gebruik_achternaam -Value $gebruik_achternaam
      $customobject | Add-Member -MemberType NoteProperty -Name werk_telefoon -Value $werk_telefoon
      $customobject | Add-Member -MemberType NoteProperty -Name mobiel_telefoon -Value $mobiel_telefoon
      $customobject | Add-Member -MemberType NoteProperty -Name werk_email -Value $werk_email
      $customobject | Add-Member -MemberType NoteProperty -Name prive_email -Value $prive_email
      $customobject | Add-Member -MemberType NoteProperty -Name Locatie -Value $Locatie
      $customobject | Add-Member -MemberType NoteProperty -Name oe_kort -Value $oe_kort
      $customobject | Add-Member -MemberType NoteProperty -Name oe_vol -Value $oe_vol
      $customobject | Add-Member -MemberType NoteProperty -Name func_kd -Value $func_kd
      $customobject | Add-Member -MemberType NoteProperty -Name func_oms -Value $func_oms
      $customobject | Add-Member -MemberType NoteProperty -Name indnst_dt -Value $indnst_dt
      $customobject | Add-Member -MemberType NoteProperty -Name uitdnst_dt -Value $uitdnst_dt
      $customobject | Add-Member -MemberType NoteProperty -Name mngr_pers_nr -Value $mngr_pers_nr
      $customobject | Add-Member -MemberType NoteProperty -Name mngr_e_roepnaam -Value $mngr_e_roepnaam
      $customobject | Add-Member -MemberType NoteProperty -Name mngr_vrvg_samen -Value $mngr_vrvg_samen
      $customobject | Add-Member -MemberType NoteProperty -Name mngr_Ing_datum -Value $mngr_Ing_datum
      $customobject | Add-Member -MemberType NoteProperty -Name mngr_Eind_datum -Value $mngr_Eind_datum

      # Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
      $array += $customobject
      $customobject = $null
    }

    #$sftpDownloadDirectoryXML
    $array | Export-Csv -Path  C:\export_beaufort.csv -NoTypeInformation
    set-content C:\export_beaufort.csv ((get-content C:\export_beaufort.csv) -replace '"')
    set-content C:\export_beaufort.csv ((get-content C:\export_beaufort.csv) -replace ",", ";")
    #Exit # for testing only
}


# 1.1 Connect SFTP
###$csvFileLocation = DownloadCsvFile
###$csvFileLocation = DownloadXMLFile

#$csvFileLocation = 'C:\work\projects\ZGIJV\ZGIJV files\ZGIJV files\prodfile.csv'

$xmlFileLocation = 'C:\Work\Projects\ZGIJV\ZGIJV files\ZGIJV files\testfile.xml'

ConvertXMLFileToCSV $xmlFileLocation

$csvFileLocation = 'C:\export_beaufort.csv'

Exit #test purposes only

###END

#2. Verwerk CSV
ProcessCsvFile -FilePath $csvFileLocation
# Moeten we nog een check doen op de bestaande lijst met gebruikers?

# # 3. Haal alle AD users op
$global:gebruikers =  Read-AD-Gegevens -ADSourceOU $ADusersource

$created = $global:accountsToBeCreated

foreach ($adGebruiker in $global:gebruikers) {
    #Filter out the accounts that are allready available in the AD
    $global:accountsToBeCreated = $global:accountsToBeCreated | Where-Object {$_.Personeelsnummer -ne $adGebruiker.employeeID }
}

$created = $global:accountsToBeCreated
$information = $global:accountsWithoutInformation;

#Create new AD Accounts
foreach ($newUser in  $global:accountsToBeCreated) {

    $voorstel = $newUser.Voorstel;
    $afdeling = $newUser.Afdeling;
    $personeelsnr = $newUser.Personeelsnummer;
    $voornaam = $newUser.Voornaam;
    $achternaam = $newUser.SamenGesteldeNaam;
    $uitdienstDatum = $newUser.UitDienstDatum
    $afdelingFunctieCode = $newUser.AfdelingFunctieCode

    # Check if user added to UitDienstOU if so move to InDienstOU and enable again 
    $adUserExistsInUitDienstOu =  Get-ADUser -SearchBase "$ADUitdienstsource" -Filter "EmployeeID -eq $($personeelsnr.Trim())" 
    if ($adUserExistsInUitDienstOu) {

        Enable-ADAccount -Identity $adUserExistsInUitDienstOu
        Move-ADObject -Identity $adUserExistsInUitDienstOu -TargetPath $ADusersource     
        # send email to support
        $emailBody = "Gebruiker met gebruikersnaam: $($voorstel.Gebruikersnaam) bestaat in de UitDienstOu deze wordt verplaatst naar InDienstOu. `r`n Personeelsnummer: $($personeelsnr)  `r`n  `r`n UserPrinicipalname: $($voorstel.UserPrincipalName)"
        SendEmail -EmailTo $AdminEmailAddress -EmailFrom $FromEmailAddress -Subject "Gebruiker bestaat al in UitDienstOu en wordt verplaatst" -Body  $emailBody
        continue
}
    # Check if user allready exist in AD (Same user starts on the same day at more then one new job)
    $adUserWithSameNumber =  Get-ADUser -SearchBase "$ADusersource" -Filter "EmployeeID -eq $($personeelsnr.Trim())" 
    $weergavenaam = $voorstel.Weergavenaam
    if ($adUserWithSameNumber) {

        # send email to support
        $emailBody = "Gebruiker met gebruikersnaam: $($voorstel.Gebruikersnaam) kan niet worden aangemaakt omdat een gebruiker met dit personleelsnummer al bestaat in de AD. `r`n Personeelsnummer: $($personeelsnr)  `r`n  `r`n UserPrinicipalname: $($voorstel.UserPrincipalName)"
        SendEmail -EmailTo $AdminEmailAddress -EmailFrom $FromEmailAddress -Subject "Gebruiker met dezelfde naam bestaat al in AD" -Body  $emailBody
        continue
    }
    else {
        $gebruikersnaam = $voorstel.Gebruikersnaam
        $adUserWithSameName = $global:gebruikers | Where-Object { $_.sAMAccountName -eq $voorstel.Gebruikersnaam  } | Select-Object -First 1
        if( $adUserWithSameName -ne $null)
        {
            $gebruikersnaam = $voorstel.AlternativeGebruikersnaam
            $adUserWithSameName = $global:gebruikers | Where-Object { $_.sAMAccountName -eq $voorstel.AlternativeGebruikersnaam  } | Select-Object -First 1
            if( $adUserWithSameName -ne $null)
            {
            $gebruikersnaam = $voorstel.AlternativeGebruikersnaam2
            $adUserWithSameName = $global:gebruikers | Where-Object {$_.sAMAccountName -eq $voorstel.AlternativeGebruikersnaam2  } | Select-Object -First 1
            }
            If($adUserWithSameName -ne $Null)
            {
                $emailBody = "Gebruiker met gebruikersnaam: $($voorstel.Gebruikersnaam) kan niet worden aangemaakt omdat een gebruiker met gebruikersnaam al bestaat `r`n Personeelsnummer: $($personeelsnr)  `r`n  `r`n Gebruikersnaam: $($voorstel.Gebruikersnaam) `r`n Alternatieuve Gebruikersnaam: $($voorstel.AlternativeGebruikersnaam)"
                SendEmail -EmailTo $AdminEmailAddress -EmailFrom $FromEmailAddress -Subject "Gebruiker met dezelfde gebruikersnaam bestaat al" -Body  $emailBody
                continue
            }
        }

        $userPrincipalName = $voorstel.UserPrincipalName
        $adUserWithSameUserPrincipalName = $global:gebruikers | Where-Object { $_.UserPrincipalName -eq $voorstel.UserPrincipalName  } | Select-Object -First 1
        if( $adUserWithSameUserPrincipalName -ne $null)
        {
            $userPrincipalName = $voorstel.AlternativeUserPrincipalName
            $weergavenaam = $voorstel.AlternativeWeergavenaam
            $adUserWithSameUserPrincipalName = $global:gebruikers | Where-Object { $_.UserPrincipalName -eq $voorstel.AlternativeUserPrincipalName  } | Select-Object -First 1
            if( $adUserWithSameUserPrincipalName -ne $null)
            {
                $emailBody = "Gebruiker met UserPrincipalName: $($voorstel.UserPrincipalName) kan niet worden aangemaakt omdat een gebruiker met UserPrincipalName al bestaat `r`n Personeelsnummer: $($personeelsnr)  `r`n  `r`n UserPrincipalName: $($voorstel.UserPrincipalName) `r`n Alternatieuve UserPrincipalName: $($voorstel.AlternativeUserPrincipalName)"
                SendEmail -EmailTo $AdminEmailAddress -EmailFrom $FromEmailAddress -Subject "Gebruiker met dezelfde UserPrincipalName bestaat al" -Body  $emailBody
                continue
            }
        }
      
        
        Write-Host $voorstel.Weergavenaam $personeelsnr "wordt aangemaakt in de AD"
        $password = Maak-gebruiker -Gebruikersnaam  $gebruikersnaam -Mailadres $userPrincipalName -Weergavenaam $weergavenaam -Voornaam $voornaam -Achternaam $achternaam  -UitDienstDatum $uitdienstDatum -Personeelsnr $personeelsnr -AfdelingFunctieCode $afdelingFunctieCode -UserPrincipalName $userPrincipalName -FunctieOmschrijving "$($newUser.AfdelingOmschrijving)_$($newUser.Afdeling)" -AfdelingCode $newUser.AfdelingCode

        $toEmailAddress = ""
        $toManagerEmailAddress = ""
        #Send email to manager or send the email to the private emailaddress of the employee
        if ($ManagerNrsWithPersonalEmail.Contains($newUser.PersoneelsnummerManager)) {

            $toEmailAddress = $newUser.PriveEmailadres
            $managerInformation = $global:gebruikers | Where-Object { $_.employeeID -eq $newUser.PersoneelsnummerManager  } | Select-Object -First 1
            if($managerInformation) {
                $toManagerEmailAddress = $managerInformation.userPrincipalName 
            }
        }
        else {
            $managerInformation = $global:gebruikers | Where-Object { $_.employeeID -eq $newUser.PersoneelsnummerManager  } | Select-Object -First 1
            if($managerInformation) {
                $toEmailAddress = $managerInformation.userPrincipalName 
            }
            else {
                $emailBody = [System.IO.File]::ReadAllText($emailTemplateManagerNotFound)

                $emailBody = $emailBody.Replace("[username]", $gebruikersnaam)
                $emailBody = $emailBody.Replace("[password]", $password)
                $emailBody = $emailBody.Replace("[emailaddress]", $userPrincipalName)
                $emailBody = $emailBody.Replace("[managerEmployeeId]", $newUser.PersoneelsnummerManager)
              
                SendEmail -Priority "High" -EmailTo $AdminEmailAddress -EmailFrom $FromEmailAddress -Subject "Mail indiensttreding is niet verstuurd" -Body  $emailBody
                continue
            }
        }
        
        $emailBody = [System.IO.File]::ReadAllText($emailTemplate)
        $emailBody = $emailBody.Replace("[username]", $gebruikersnaam)
        $emailBody = $emailBody.Replace("[password]", $password)
        $emailBody = $emailBody.Replace("[emailaddress]", $userPrincipalName)
        $emailBody = $emailBody.Replace("[weergavenaam]", $voorstel.Weergavenaam)
        $emailBody = $emailBody.Replace("[oe_vol_func_oms]", "$($newUser.AfdelingOmschrijving) $($newUser.Afdeling)")

        $emailBody = $emailBody.Replace("[datumindienst]", $newUser.InDienstDatum)
        $emailBody = $emailBody.Replace("[datumuitdienst]", $uitdienstDatum)
        $emailBody = $emailBody.Replace("[personeelsnummer]", $personeelsnr)
        $emailBody = $emailBody.Replace("[naamleidinggevende]", $newUser.NaamLeidinggevende)
        
        $CC = @("$($AdminEmailAddress)")

        if($CCEmailAddresses)
        {
            $CC += $CCEmailAddresses
        }

        if($toManagerEmailAddress)
        {
            $CC += @("$($toManagerEmailAddress)")
        }

        SendEmail -EmailTo $toEmailAddress -CC $CC -EmailFrom $FromEmailAddress -Subject "Nieuwe indiensttreding" -Body  $emailBody
    }
}

#Haal uit dienst lijst op uit Uitdient OU
$uitDienstUsers =  Read-AD-Gegevens -ADSourceOU $ADUitdienstsource

#AD accounts with the datum 
foreach($adUser in $global:gebruikers)
{
    $date = $adUser.extensionAttribute1
    if($date)
    {
        $date = ConvertToDate -Date $date;
        if (($date) -and ( $date.Date -lt $today.Date ))
        {
            $gebruikersnaam = $adUser.userPrincipalName
            $personeelsnr = $adUser.employeeID
            $mailAddress = $adUser.userPrincipalName


                $adDeleteUser =  Get-ADUser -SearchBase "$ADusersource" -Filter "EmployeeID -eq $($personeelsnr)" 
                $userDeactivated = $false

                # $uitDienstExtraTimePeriod = $adUser.extensionAttribute2
                # if($uitDienstExtraTimePeriod)
                # {
                #     $uitDienstExtraTimePeriod = ConvertToDate -Date $uitDienstExtraTimePeriod;
                # }

                if ($adDeleteUser -and $adDeleteUser.Enabled ){   #-and !($uitDienstExtraTimePeriod) -or ( $uitDienstExtraTimePeriod.Date -lt $today.Date )) {
                    Disable-ADAccount -Identity $adDeleteUser
                    Move-ADObject -Identity $adDeleteUser -TargetPath $ADUitdienstsource        
                    $userDeactivated = $true
                }
                # else {

                #     Write-Host $gebruikersnaam $personeelsnr "nog niet verplaatst naar uitdienst extra periode tot:"  $uitDienstExtraTimePeriod
                #     continue
                # }

                if($userDeactivated) {
                    Write-Host $gebruikersnaam $personeelsnr "verplaatst naar uitdienst"
                    #Email sturen naar beheerder dat gebruiker uit dienst is gehaald
                    $emailBody = "Gebruiker met gebruikersnaam: $($gebruikersnaam) en personeelsnummer: $($personeelsnr). `r`n Deze gebruiker is verplaatst naar de UitDienst AD en is op InActief gezet."
                    SendEmail -EmailTo $AdminEmailAddress -EmailFrom $FromEmailAddress -Subject "Gebruiker is verplaatst naar UitDienst" -Body  $emailBody
                }
                else {
                    $emailBody = "Het uitdiensttreed proces voor gebruiker: $($gebruikersnaam) met personeelsnummer $($personeelsnr) is mislukt gebruiker niet gevonden in AD"
                    SendEmail -Priority "High" -EmailTo $AdminEmailAddress -EmailFrom $FromEmailAddress -Subject "Uitdiensttreding gebruiker mislukt" -Body  $emailBody
                }
        }
    }
}

$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $OfficeUsername, $OfficePassword

KRS_Loginoffice365 -Credentials $Credentials

$Office365Users = Search-Users 

 $E3LicentieGroupMembers = Get-ADGroupMember -Identity "Ap-OfficelicentieE3" -Recursive | Select -ExpandProperty SamAccountName    

 $E1LicentieGroupMembers = Get-ADGroupMember -Identity "Ap-OfficelicentieE1" -Recursive | Select -ExpandProperty SamAccountName    

 $VisioLicentieGroupMembers = Get-ADGroupMember -Identity "Ap-Visio" -Recursive | Select -ExpandProperty SamAccountName   

#Create Office356 Licenses for all users that are in AD but do not have a Office license
foreach ($adGebruiker in $global:gebruikers) {   

    $officeUser = $Office365Users | Where-Object { $_.UserPrincipalName -eq $adGebruiker.userPrincipalName  } | Select-Object -First 1

        if (!$officeUser -or (!$officeUser.Licentie -and !$officeUser.Uitgeschakeld)) {

            if($E3LicentieGroupMembers -contains  $adGebruiker.sAMAccountName)
            {   
               try {
                    Write-Host $adGebruiker.sAMAccountName "add office E3"
                    Set-MsolUser -UserPrincipalName $adGebruiker.userPrincipalName -UsageLocation NL
                    Set-MsolUserLicense -UserPrincipalName $adGebruiker.userPrincipalName -AddLicenses $Office365E3 
                }
                catch {
                    $emailBody = "Het koppelen van Office 356 E3 licenties is mislukt voor gebruiker: $($adGebruiker.userPrincipalName)"
                    SendEmail -Priority "High" -EmailTo $AdminEmailAddress -EmailFrom $FromEmailAddress -Subject "Office E3 licentie niet gekoppeld" -Body  $emailBody
                }

            }
            if( $E1LicentieGroupMembers -contains  $adGebruiker.sAMAccountName)
            {
                try {
                    Write-Host $adGebruiker.sAMAccountName "add office E1"
                    Set-MsolUser -UserPrincipalName $adGebruiker.userPrincipalName -UsageLocation NL
                    Set-MsolUserLicense -UserPrincipalName $adGebruiker.userPrincipalName -AddLicenses $Office365E1  
                }
                catch {
                    $emailBody = "Het koppelen van Office 356 E1 licenties is mislukt voor gebruiker: $($adGebruiker.userPrincipalName)"
                    SendEmail -Priority "High" -EmailTo $AdminEmailAddress -EmailFrom $FromEmailAddress -Subject "Office E1 licentie niet gekoppeld" -Body  $emailBody
                }
            }
            if( $VisioLicentieGroupMembers -contains  $adGebruiker.sAMAccountName)
            {
                try {
                    Write-Host $adGebruiker.sAMAccountName "add office Visio"
                    Set-MsolUser -UserPrincipalName $adGebruiker.userPrincipalName -UsageLocation NL
                    Set-MsolUserLicense -UserPrincipalName $adGebruiker.userPrincipalName -AddLicenses $OfficeVisio  
                }
                catch {
                    $emailBody = "Het koppelen van Office Visio licenties is mislukt voor gebruiker: $($adGebruiker.userPrincipalName)"
                    SendEmail -Priority "High" -EmailTo $AdminEmailAddress -EmailFrom $FromEmailAddress -Subject "Office Visio licentie niet gekoppeld" -Body  $emailBody
                }
            }
        }
}



#Remove Office356 Licenses
foreach ($officeUser in $Office365Users) {
    $mailAddress = $officeUser.UserPrincipalName

    $adUser =  $uitDienstUsers | Where-Object { $_.userPrincipalName.ToLower() -eq   $mailAddress.ToLower()  } | Select-Object -First 1

    if ($adUser -and !$adUser.Enabled -and $officeUser.Licentie) {
        Set-MsolUserLicense -UserPrincipalName $officeUser.UserPrincipalName -RemoveLicenses $officeUser.Licentie
        Write-Host $officeUser.UserName "Remove office"
    }
}





  
