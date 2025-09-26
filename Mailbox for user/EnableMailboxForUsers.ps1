#*********************************************************#
#* Enable Email for Existing AD Users                    *#
#* 17.08.2025                                            *#
#* Author: CYRIAC JOSE                                   *#
#*********************************************************#

# Import AD Module
Import-Module ActiveDirectory

$ExchangeServer = ""
$UserCredential = Get-Credential

# Start a remote PowerShell session to Exchange
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos -Credential $UserCredential

# Import the session into your current session
Import-PSSession $session -DisableNameChecking -AllowClobber

# Store the data from your CSV file in the $ADUsers variable
$ADUsers = Import-Csv "Same CSV as the one for account creation"

# Mailbox DB Mapping eveything we need to change it according to requirement
$MailboxDB_AtoM = ""
$MailboxDB_NtoZ = ""

$logResults = @()

foreach ($User in $ADUsers) {
    $Firstname = $User.firstname
    $Lastname = $User.lastname
    $Initial   = $User.initials
    $Username = $User.name
    $email = $User.email
    $streetaddress = $User.address
    $Description = $User.Description  
    $city = $User.city
    $zipcode = $User.zipcode
    $state = $User.state
    $country = $User.country
    $department = $User.department
    $Password = $User.password
    $telephone = $User.telephone
    $jobtitle = $User.title
    $company = $User.company
    $OU = $User.OU
    $UPN = $User.Name + "@" + $User.Maildomain

    # This will determine first letter of first name and move to DB
    $firstChar = $FirstName.Substring(0,1);
    if ($firstChar -ge 'A' -and $firstChar -le 'M') {
        $MailboxDatabase = $MailboxDB_AtoM
    } else {
        $MailboxDatabase = $MailboxDB_NtoZ
    }

    $result = New-Object PSObject -property @{
        Username      = $Username
        Firstname     = $Firstname
        Lastname      = $Lastname
        Status        = "Pending"
        ErrorMessage  = ""
    }

    # Check whether the user already exists in AD
    $existingUser = Get-ADUser -Filter {SamAccountName -eq $Username}

    if ($existingUser) {
        # If the user already exists, update status
        $result.Status = "User Exists"
        $result.ErrorMessage = "A user account with username $Username already exists in Active Directory."
        
        $existingMailbox = Get-Mailbox -Identity $email -ErrorAction SilentlyContinue
        if ($existingMailbox) {
            $result.Status = "Mailbox Exists"
            $result.ErrorMessage = "The mailbox for $email already exists in Exchange."
        } else {

            try {
                Enable-Mailbox -Identity $email -Database $MailboxDatabase -Alias $Username
                $result.Status = "Mailbox Enabled"
            } catch {
                $result.Status = "Error"
                $result.ErrorMessage = "Error enabling mailbox: $($_.Exception.Message)"
            }
        }
    }
    else {
        # If the user doesn't exist, log the status as "User Not Found"
        $result.Status = "User Not Found"
        $result.ErrorMessage = "No user found with username $Username in Active Directory."
    }

    # Add the result to the log array
    $logResults += $result
}

# Export the results to a CSV file
$logResults | Export-Csv "PATH TO SAVE THE FILES"

Write-Host "User creation log has been exported"
Remove-PSSession $session







