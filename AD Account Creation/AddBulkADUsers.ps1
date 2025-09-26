#*********************************************************#
#* Creating Bulk Users in AD                             *#
#* 12.08.2025                                            *#
#* Author: CYRIAC JOSE                                   *#
#*********************************************************#

# Importing AD modules.
Import-Module ActiveDirectory

# This Variable will call the users details in the CSV.
$ADUsers = Import-Csv "Path to your CSV having user data."

$logResults = @()

# For each user in the Varibale.
foreach ($User in $ADUsers) {
    $Firstname = $User.firstname
    $Lastname = $User.lastname
    $Initial   = $User.initials
    $Username = $User.name
    $streetaddress = $User.address
    $Description = $User.Description 
    $Office = $User.Office  
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

    # Define a result object to store the status for the user
    $result = New-Object PSObject -property @{
        Username      = $Username
        Firstname     = $Firstname
        Lastname      = $Lastname
        Status        = ""
        ErrorMessage  = ""
    }

    # Check whether the user already exists in AD
    $existingUser = Get-ADUser -Filter {SamAccountName -eq $Username}

    if ($existingUser) {
        # If the user already exists, this create log as username exist
        $result.Status = "User Exists"
        $result.ErrorMessage = "A user account with username $Username already exists in Active Directory."
        Write-Warning "A user account with username $Username already exists in Active Directory."
    }
    else {
        # Otherwise, create the new user account in the specified OU
        try {
            New-ADUser -SamAccountName $Username `
                -UserPrincipalName $UPN `
                -DisplayName "$Firstname $Lastname" `
                -Name "$Firstname $Lastname" `
                -GivenName $Firstname `
                -Surname $Lastname `
                -Enabled $True `
                -Description $Description `
                -Office $Office `
                -Path $OU `
                -City $city `
                -Company $company `
                -State $state `
                -StreetAddress $streetaddress `
                -OfficePhone $telephone `
                -EmailAddress $email `
                -Title $jobtitle `
                -Department $department `
                -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) `
                -ChangePasswordAtLogon $True

            # If the user creation is successful, log the status as "User Created"
            $result.Status = "User Created"
        }
        catch {
            # If an error occurs, log it in the result
            $result.Status = "Error"
            $result.ErrorMessage = "Error creating user: $($_.Exception.Message)"
        }
    }

    # Add the result to the log array
    $logResults += $result
}

# Export the results to a CSV file
$logResults | Export-Csv "Put the path where you want to store the log file" -NoTypeInformation

Write-Host "User creation log has been exported."
