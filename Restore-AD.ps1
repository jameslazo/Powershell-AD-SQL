# James Lazo 011529541

# Bypass run option prompts
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Try block
Try {
    # Set pwd as location
    Set-Location -Path $PSScriptRoot

    # Import AD
    Import-Module ActiveDirectory

    # Check for Finance OU
    if (Get-ADOrganizationalUnit -Filter { Name -eq "Finance" }) {
        # Create users object for deletion
        $users = Get-ADUser -Filter * -SearchBase "OU=Finance,DC=consultingfirm,DC=com" -Property Name

        # Delete users from OU
        foreach ($user in $users) {
            Remove-ADUser -Identity $users -Confirm:$false
        }

        # Output deleted users from Finance
        Write-Host "Deleted Finance users"

        # Delete if exists
        Remove-ADOrganizationalUnit -Identity "OU=Finance,DC=consultingfirm,DC=com" -Confirm:$false
        Write-Host "Deleted Finance"
    }

    # Output Finance OU does not exist
    else {
        Write-Host "Finance does not exist"
    }

    # Create Finance OU
    New-ADOrganizationalUnit -Name "Finance" -Path "DC=consultingfirm,DC=com" -ProtectedFromAccidentalDeletion $false
    Write-Host "Created Finance"

    # Import csv and iterate over objects
    Import-Csv .\financePersonnel.csv | ForEach-Object {
        # Create variables for user principal name and display name
        $userPrincipal = $_."samAccount" + "@consultingfirm.com"
        $firstLast = $_."First_Name" + " " + $_."Last_Name"

        # Remove users for debugging
        if (Get-ADUser -Identity "CN=$firstLast,OU=Finance,DC=consultingfirm,DC=com") {
            Remove-ADUser -Identity "CN=$firstLast,OU=Finance,DC=consultingfirm,DC=com" -Confirm:$false
        }

        # Create user per csv object
        New-ADUser -Name $firstLast -GivenName $_.First_Name -Surname $_.Last_Name -City $_.City -State $_.County -PostalCode $_.PostalCode -OfficePhone $_.OfficePhone -MobilePhone $_.MobilePhone -DisplayName $firstLast -UserPrincipalName $userPrincipal -Path "OU=Finance,DC=consultingfirm,DC=com"
    }

    # Output users to file
    Get-ADUser -Filter * -SearchBase "OU=Finance,DC=consultingfirm,DC=com" -Properties DisplayName,PostalCode,OfficePhone,MobilePhone > .\ADResults.txt
}

# Catch block
Catch {
    Write-Host "Exception"
}