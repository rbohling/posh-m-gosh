#Import Active Directory Module
Import-Module ActiveDirectory
# You will need to have Powershell Credential Manager module installed. You can install this from the Powershell Gallery with the following command:
# Install-Module -Name CredentialManager
Import-Module CredentialManager
# Import MS Online (Office 365) module to create O365 mailbox
Import-Module MSOnline

#Load system.web assembly. This is used to generate secure random password.
[Reflection.Assembly]::LoadWithPartialName("System.Web")

# Function will create a HTML email using the details during the new AD account process to send to the employee and manager.
function Send-WelcomeEmail {
    param (
        [Parameter(Mandatory=$true)][string]$Recipient,
        [Parameter(Mandatory=$false)][string]$Sender = "itsupport@company.com",
        [Parameter(Mandatory=$true)][string]$UserEmailAddress,
        [Parameter(Mandatory=$true)][string]$UserDisplayName,
        [Parameter(Mandatory=$true)][string]$UsersAMAccountName,
        [Parameter(Mandatory=$true)][string]$UserPassword,
        [Parameter(Mandatory=$false)][string]$CC,
        [Parameter(Mandatory=$true)][string]$UserManager
        )

        $MessageBody = @"
        <p>Dear $UserManager,</p>
        <p>Please make sure that the following information is provided to your new hire. If you have any questions, please do not hesitate to contact IT.</p>
        <p>&nbsp;</p>
        <p>Dear $UserDisplayName,</p>
        <p>Welcome to CCI. Please find your IT credentials and information below.</p>
        <p><strong>Windows AD Username:</strong> $UsersaMAccountName<br /><strong>O365 Username and Email Address:</strong> $UserEmailAddress<br /><strong>Password:</strong> $UserPassword<br /><br /><strong>To change your password:</strong><br />1. Logon to your computer.<br />2. Press CTRL + ALT + DEL and click Change Password.</p>
        <p><strong>Information:</strong><br />1. Your windows account is synchronized to Office 365. If you change your windows password, it will automatically be replicated to Azure AD/Office 365.<br />2. To access the Office 365 portal, go to https://portal.office.com<br />3. Once you login to Office 365, you will be asked to add your mobile or alternative email address for Multi Factor Authentication. We recommend that you download the Microsoft Authenticator app (alternatively Google Authenticator will work) and set it up for push MFA requests. To do this, follow the instructions at the link below:<br /><br /><a href="https://docs.microsoft.com/en-us/azure/multi-factor-authentication/end-user/microsoft-authenticator-app-how-to" target="_blank" rel="noopener">https://docs.microsoft.com/en-us/azure/multi-factor-authentication/end-user/microsoft-authenticator-app-how-to</a></p>
        <p><strong>Support:</strong><br />If you have any questions or concerns, please lodge a support ticket at <a href="mailto:servicedesk@company.com">servicedesk@company.com</a>.</p>
        <p><strong>Important links:</strong></p>
        <ul>
          <li>SharePoint: <a href="https://company.sharepoint.com">company.sharepoint.com</a></li>
          <li>Jira/Confluence: <a href="https://company.atlassian.net">company.atlassian.net</a></li>
          <li>VPN: <a href="https://vpn.company.com:1234">vpn.company.com:1234</a></li>
          <li>Company Workflow: <a href="http://workflowserver.company.com/company">workflowserver.company.com/company</a></li>
          <li>Learning Management System: <a href="https://learning.company.com">learning.company.com</a></li>
          <li>Yammer: <a href="https://www.yammer.com/company.com">www.yammer.com/company.com</a></li>
        </ul>
        <p>&nbsp;</p>
        <p>Once again, welcome and please contact support if you require any further assistance.</p>
        <p>Thanks,</p>
        <p>company IT</p>
"@

        $MailMessage = @{
            From = $Sender
            To = $Recipient
            Body = $MessageBody
            SmtpServer = "smtp.company.com"
            Subject = "$UserDisplayName Login Credentials "
            CC = $CC
        }
        Send-MailMessage @MailMessage -BodyAsHtml
}

# Misc Variables
$Company = "Contract Callers Inc."
$AADServer = "auoolmgmt01.company.com"
$Domain = "cci.local"
$SPList = "https://company.sharepoint.com/corporate/hr"
$SPCredentials = "SharePoint_Admin"
$DomainCredential = "RCAdmin"
$UsageLocation = "AU"
$O365License = "company:STANDARDPACK"
$DYNLicense = "company:DYN365_ENTERPRISE_PLAN1"
$AzureADCredential = Get-StoredCredential -Target $SPCredentials
$ADCredential = Get-StoredCredential -Target $DomainCredential
$ITSupport = "itsupport@company.com"

# Office Location information in hash tables
$OOL_Office = @{
    Path = "OU=Users,OU=AU,OU=company,DC=company,DC=com"
    Street = "2/183 Some Parade"
    POBox = "PO Box 111, Some Town, QLD AU 4789"
    State = "QLD"
    PostalCode = "4227"
    Country = "AU"
}
$SEA_Office = @{
    Path = "OU=Users,OU=NA,OU=company,DC=company,DC=com"
    Street = "25 42nd Ave, Suite 5015"
    State = "MA"
    PostalCode = "38121"
    Country = "US"
}
$MNL_Office = @{
    Path = "OU=Users,OU=AS,OU=company,DC=company,DC=com"
    Street = "1 Miguel Avenue"
    PostalCode = "4800"
    Country = "PH"
}
$BRU_Office = @{
    Path = "OU=Users,OU=EU,OU=company,DC=company,DC=com"
    Street = "Co.Station, Oktrooiplein"
    PostalCode = "9000"
    Country = "BE"
}
$EWR_Office = @{
    Path = "OU=Users,OU=NA,OU=company,DC=company,DC=com"
    Street = "Suite 1, 122 Bossmore Drive"
    PostalCode = "86531"
    Country = "US"
}

# Standard AD SG's
$STD_SG = @("company Employees", "AADconnect_sync", "VPN Access", "PrjSrv_Team Members")

# Office Distribution Lists
$OOL_DL = "!Staff_AU"
$SEA_DL = "!Staff_US"
$MNL_DL = "!Staff_PH"
$BRU_DL = "!Staff_EU"
$SA_DL = "!Staff_SA"

# Organisational Distribution Lists
$PS_DL = "!PS"
$Sales_DL = @("!Sales","!Sales_Internal")
$RnD_DL = "!Dev"
$Support_DL = "!Support"
$Exec_DL = @("!Executive","Managers")


# Connect to SharePoint via PnP
Connect-PnPOnline -Url $SPList -Credentials $SPCredentials

# Gather all new employees from SharePoint List
$ListItems = (Get-PnPListItem -List 'New User Request').FieldValues

# Loop through SharePoint list, create new AD account for any employee that has been added and has a status of 'New'
foreach ($item in $ListItems) {
    if ($item.Status -eq 'New') {
        $Name = $($item.Preferred_x0020_First_x0020_Name + "." + $item.Last_x0020_Name).ToLower() -replace " ",""
        $DisplayName = $($item.Preferred_x0020_First_x0020_Name + " " + $item.Last_x0020_Name)
        $EmailAddress = $($Name + "@" + $Domain)
        $Manager = $item.Manager.LookupValue
        $ManagerDN = (Get-ADUser -Filter {DisplayName -eq $Manager} | Select-Object DistinguishedName -ExpandProperty DistinguishedName)
        $ManagerEmail = (Get-ADUser -Filter {DisplayName -eq $Manager} | Select-Object UserPrincipalName -ExpandProperty UserPrincipalName)
        $Password = [System.Web.Security.Membership]::GeneratePassword((Get-Random -Minimum 12 -Maximum 16), 2)
        $AccountPassword = ConvertTo-SecureString -String $password -AsPlainText -Force
        $ID = $item.ID
        $Office = $item.Office
        $Department = $item.Department

        $newuser = @{
            Name = $DisplayName
            UserPrincipalName = $EmailAddress
            sAMACcountName = $Name
            Description = $item.Position
            DisplayName = $DisplayName
            EmailAddress = $EmailAddress
            GivenName = $item.Preferred_x0020_First_x0020_Name
            Surname = $item.Last_x0020_Name
            Manager = $ManagerDN
            EmployeeID = $item.Title
            City = $Office
            Department = $Department
            Title = $item.Position
            AccountPassword = $AccountPassword
            Credential = $ADCredential
            Office = $Office
            Company = $Company
            Enabled = $true
        }

        Connect-MsolService -Credential $AzureADCredential
        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $AzureADCredential -Authentication "Basic" -AllowRedirection
        Import-PSSession $ExchangeSession -DisableNameChecking -AllowClobber

        # Create new AD user based on data pulled from SharePoint list, sync with Azure AD, assign O365 license, create mailbox and add to country DL.
        if ($Office -eq "Some Town" -Or $Office -eq "Remote - Australia") {
            $User = $newuser + $OOL_Office
            New-ADUser @User
            Set-ADUser -Identity $Name -Replace @{Co="AUS";countryCode="036";C="AU"}  -Credential $ADCredential
                foreach ($SG in $STD_SG) {
                    Add-ADGroupMember -Identity $SG -Members $Name -Credential $ADCredential
                }

            Start-ADSyncSyncCycle -PolicyType Delta
            Start-Sleep 120

            Set-MsolUser -UserPrincipalName $EmailAddress -UsageLocation $UsageLocation
            Start-Sleep 15

            Set-MsolUserLicense -UserPrincipalName $EmailAddress -AddLicenses $O365License
            Write-Host "Waiting for Exchange mailbox creation (2.5mins)"
            Start-Sleep 150

            Add-DistributionGroupMember -Identity $OOL_DL -Member $EmailAddress
        }

        if ($Office -eq "Seattle" -Or $Office -eq "Remote - North America") {
            $User = $newuser + $SEA_Office
            New-ADUser @User
            Set-ADUser -Identity $Name -Replace @{Co="USA";countryCode="840";C="US"}  -Credential $ADCredential
                foreach ($SG in $STD_SG) {
                    Add-ADGroupMember -Identity $SG -Members $Name -Credential $ADCredential
                }

            Start-ADSyncSyncCycle -PolicyType Delta
            Start-Sleep 120

            Set-MsolUser -UserPrincipalName $EmailAddress -UsageLocation $UsageLocation
            Start-Sleep 15

            Set-MsolUserLicense -UserPrincipalName $EmailAddress -AddLicenses $O365License
            Write-Host "Waiting for Exchange mailbox creation (2.5mins)"
            Start-Sleep 150

            Add-DistributionGroupMember -Identity $SEA_DL -Member $EmailAddress
        }

        if ($Office -eq "New York") {
            $User = $newuser + $EWR_Office
            New-ADUser @User
            Set-ADUser -Identity $Name -Replace @{Co="USA";countryCode="840";C="US"}  -Credential $ADCredential
                foreach ($SG in $STD_SG) {
                    Add-ADGroupMember -Identity $SG -Members $Name -Credential $ADCredential
                }

            Start-ADSyncSyncCycle -PolicyType Delta
            Start-Sleep 120

            Set-MsolUser -UserPrincipalName $EmailAddress -UsageLocation $UsageLocation
            Start-Sleep 15

            Set-MsolUserLicense -UserPrincipalName $EmailAddress -AddLicenses $O365License
            Write-Host "Waiting for Exchange mailbox creation (2.5mins)"
            Start-Sleep 150

            Add-DistributionGroupMember -Identity $SEA_DL -Member $EmailAddress
        }

        if ($Office -eq "Remote - South America") {
            $User = $newuser + $SEA_Office
            New-ADUser @User
            Set-ADUser -Identity $Name -Replace @{Co="USA";countryCode="840";C="US"}  -Credential $ADCredential
                foreach ($SG in $STD_SG) {
                    Add-ADGroupMember -Identity $SG -Members $Name -Credential $ADCredential
                }

            Start-ADSyncSyncCycle -PolicyType Delta
            Start-Sleep 120

            Set-MsolUser -UserPrincipalName $EmailAddress -UsageLocation $UsageLocation
            Start-Sleep 15

            Set-MsolUserLicense -UserPrincipalName $EmailAddress -AddLicenses $O365License
            Write-Host "Waiting for Exchange mailbox creation (2.5mins)"
            Start-Sleep 150

            Add-DistributionGroupMember -Identity $SA_DL -Member $EmailAddress
        }

        if ($Office -eq "Manila" -Or $Office -eq "Remote - Asia") {
            $User = $newuser + $MNL_Office
            New-ADUser @User
            Set-ADUser -Identity $Name -Replace @{Co="PHL";countryCode="608";C="PH"}  -Credential $ADCredential
                foreach ($SG in $STD_SG) {
                    Add-ADGroupMember -Identity $SG -Members $Name -Credential $ADCredential
                }

            Start-ADSyncSyncCycle -PolicyType Delta
            Start-Sleep 120

            Set-MsolUser -UserPrincipalName $EmailAddress -UsageLocation $UsageLocation
            Start-Sleep 15

            Set-MsolUserLicense -UserPrincipalName $EmailAddress -AddLicenses $O365License
            Write-Host "Waiting for Exchange mailbox creation (2.5mins)"
            Start-Sleep 150

            Add-DistributionGroupMember -Identity $MNL_DL -Member $EmailAddress
        }

        if ($Office -eq "Belgium" -Or $Office -eq "Remote - Europe") {
            $User = $newuser + $BRU_Office
            New-ADUser @User
            Set-ADUser -Identity $Name -Replace @{Co="BEL";countryCode="056";C="BE"}  -Credential $ADCredential
                foreach ($SG in $STD_SG) {
                    Add-ADGroupMember -Identity $SG -Members $Name -Credential $ADCredential
                }

            Start-ADSyncSyncCycle -PolicyType Delta
            Start-Sleep 120

            Set-MsolUser -UserPrincipalName $EmailAddress -UsageLocation $UsageLocation
            Start-Sleep 15

            Set-MsolUserLicense -UserPrincipalName $EmailAddress -AddLicenses $O365License
            Write-Host "Waiting for Exchange mailbox creation (2.5mins)"
            Start-Sleep 150

            Add-DistributionGroupMember -Identity $BRU_DL -Member $EmailAddress
        }

        # Add employees to department AD security groups
        if ($Department -eq "Research and Development" -Or $Department -eq "Analytics" -Or $Department -eq "Engineering" -Or $Department -eq "Custom Solutions") {
            Add-ADGroupMember -Identity "Research and Development" -Members $Name -Credential $ADCredential
            Add-DistributionGroupMember -Identity $RnD_DL -Member $EmailAddress

        }

        if ($Department -eq "PMO" -Or $Department -eq "Professional Services APAC/EMEA" -Or $Department -eq "Professional Services North America" -Or $Department -eq "Professional Services/PMO") {
            Add-ADGroupMember -Identity "Professional Services" -Members $Name -Credential $ADCredential
            Add-DistributionGroupMember -Identity $PS_DL -Member $EmailAddress
        }

        if ($Department -eq "Sales") {
            Add-ADGroupMember -Identity "Sales and Marketing" -Members $Name -Credential $ADCredential
            Add-DistributionGroupMember -Identity $Sales_DL -Member $EmailAddress
            Set-MsolUserLicense -UserPrincipalName $EmailAddress -AddLicenses $DYNLicense
        }

        if ($Department -eq "Support") {
            Add-ADGroupMember -Identity "Support Services" -Members $Name -Credential $ADCredential
            Add-DistributionGroupMember -Identity $Support_DL -Member $EmailAddress
        }

        if ($Department -eq "Executive") {
            Add-ADGroupMember -Identity "Company Administration" -Members $Name -Credential $ADCredential
            Add-DistributionGroupMember -Identity $Exec_DL -Member $EmailAddress
        }

        if ($Department -eq "Quality Assurance") {
            Add-ADGroupMember -Identity "Quality Assurance" -Members $Name -Credential $ADCredential
            Add-ADGroupMember -Identity "Research and Development" -Members $Name -Credential $ADCredential
            Add-DistributionGroupMember -Identity $RnD_DL -Member $EmailAddress
        }

        # Close remote PS session to Exchange Online.
        Remove-PSSession $exchangeSession

        # Update the employee record in SharePoint list to Completed so the script doesn't process the request again.
        $Update_List = Set-PnPListItem -List 'New User Request' -Identity $ID -Values @{"Status" = "Completed"}


        # Email to employee and manager with information and credentials
        Send-WelcomeEmail -Recipient $ManagerEmail -CC $ITSupport -UserManager $Manager -UserEmailAddress $EmailAddress -UserDisplayName $DisplayName -UsersAMAccountName $Name -UserPassword $Password

        # Small delay before next loop object. This allows the Azure AD Connect to finish up its sync process before restarting.
        Start-Sleep 60

    }
}
