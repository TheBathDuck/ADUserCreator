$csvFilePath = "C:\Users\Administrator\Desktop\klanten.csv" #Zet naar data
$csvData = Import-Csv -Path $csvFilePath

$DefaultPassword = (ConvertTo-SecureString "Welkom01" -AsPlainText -Force)
$ADLocation = "OU=TestUsers,DC=testbedrijf,DC=local" #Verander OU en DC naar je Active Domain.
$EMail = "@testbedrijf.local"

Write-Host "Starting.."
$createdUsersCount = 0

foreach($Row in $csvData) {
    $CompanyUserId = $Row.Personeelsnummer
    $FirstName = $Row.Roepnaam
    $Tussenvoegsel = $Row.Tussenvoegsel
    $LastName = $(if ($Tussenvoegsel -eq "") { "" } else { "${Tussenvoegsel} " }) + $Row.Achternaam
    $MailAddress = $CompanyUserId + $EMail
    $DisplayName = $FirstName + " " + $LastName

    Write-Host "Attempting to create user with name: ${FirstName} (${Username})"

    $newUser = @{
        Path = $ADLocation
        SamAccountName = $CompanyUserId
        Name = $CompanyUserId
        GivenName = $FirstName
        Surname = $LastName
        DisplayName = $DisplayName
        EmailAddress = $MailAddress
        UserPrincipalName = $MailAddress
        Enabled = $true
        ChangePasswordAtLogon = $true
        AccountPassword = $DefaultPassword
        City = $Row.Plaats
        StreetAddress = $Row.Adres
        MobilePhone = $Row.Tel
    }

    try {
        New-ADUser @newUser
        $createdUsersCount++
    } catch {
        Write-Host "Failed to create user: $_"
    }
}

Write-Host "Ended, created ${createdUsersCount} user(s)."
