$csvFilePath = "C:\Users\Administrator\Desktop\AD-Powershell\smallklanten.csv"
$csvData = Import-Csv -Path $csvFilePath

$DefaultPassword = (ConvertTo-SecureString "Welkom01" -AsPlainText -Force)
$ADLocation = "OU=TestUsers,DC=testbedrijf,DC=local" #Verander OU en DC naar je Active Domain.
$EMail = "@testbedrijf.local"

#Change windows accent color
$registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Accent"
$valueName = "AccentColorMenu"
$accentColor = "FFFF00"
Set-ItemProperty -Path $registryPath -Name $valueName -Value $accentColor
Write-Host "Edited registery color, only works after restart."

#Create share function
function CreateUserShare {
    Param (
        [string]$CompanyUserId
    )

    $FolderPath = 'C:\Shares\' + $CompanyUserId
    if (-not (Test-Path $FolderPath)) {
        New-Item -ItemType Directory -Path $FolderPath | Out-Null
    }

    # fix share permissions
    $ShareName = "Share-" + $CompanyUserId
    $AccountName = 'TESTBEDRIJF\' + $CompanyUserId

    $Parameters = @{
        Name = $ShareName
        Path = $FolderPath
        FullAccess = $AccountName
    }

    New-SmbShare @Parameters

    Get-SmbShareAccess -Name $ShareName | ForEach-Object {
        if ($_.AccountName -ne $AccountName) {
            Remove-SmbShareAccess -Name $ShareName -AccountName $_.AccountName -Force
        }
    }
    
    Grant-SmbShareAccess -Name $ShareName -AccountName $CompanyUserId -AccessRight Full -Force

    Write-Host "Share created for user: $CompanyUserId with path: $FolderPath"
}

# Download notepad ++
$notepadPPPath = "C:\Program Files (x86)\Notepad++\notepad++.exe"
if (Test-Path $notepadPPPath) {
    Write-Host "Notepad++ is already installed, skipping."
} else {
    $notepadPPInstallerUrl = "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.3.1/npp.8.3.1.Installer.exe"
    $notepadPPInstallerPath = "$env:TEMsP\npp.8.3.1.Installer.exe"
    Write-Host "Downloading Notepad++"
    Invoke-WebRequest -Uri $notepadPPInstallerUrl -OutFile $notepadPPInstallerPath -UseBasicParsing -Method Get
    Start-Process -FilePath $notepadPPInstallerPath -ArgumentList "/S" -Wait
    Write-Host "Notepad++ Installed."
}

Write-Host "Attempting to create users."
$createdUsersCount = 0

#Create users
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
        CreateUserShare -CompanyUserId $CompanyUserId
        $createdUsersCount++
    } catch {
        Write-Host "Failed to create user: $_"
    }
}

Write-Host "Ended, created ${createdUsersCount} user(s)."