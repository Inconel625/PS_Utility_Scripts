<#
=============================================================================================
Bulk creates users in AD from a CSV file, including their location, groups, password, etc.
Template for csv, headers must remain intact:
FirstName,LastName,DisplayName,SAM,Description,Email,Password,OU,Groups,ChangePWAtLogin,PasswordNeverExpires,MailDomain
Bob,Bobert,Bob Bobert,bob.bobert,TestUser,bbobert@test.local,SomePasswordThingIGuess$231,"OU=Test,OU=Users,DC=Branch,DC=Company,DC=com","TestGroup,BestGroup",FALSE,TRUE,branch.company.com
============================================================================================
#>
$CSVLocation = "C:\Usersadmin\Documents\NewUsers.csv"
$Users = Import-Csv -Path $CSVLocation
foreach ($User in $Users)
{
    $Displayname = $User.'Firstname' + " " + $User.'Lastname'
    $UserFirstname = $User.'Firstname'
    $UserLastname = $User.'Lastname'
    $OU = $User.'OU'
    $SAM = $User.'SAM'
    $UPN = $User.'Firstname' + "." + $User.'Lastname' + "@" + $User.'Maildomain'
    $Description = $User.'Description'
    $Password = $User.'Password'
    $Email = $User.'Email'
    $Groups = $User.'Groups'.split(",") #Splits out the group list in the CSV into a string array
    $PwChgLogin = [System.Convert]::ToBoolean($User.'ChangePWAtLogin') #Converts the value from the CSV into a boolean that PS will understand
    $PwExpire = [System.Convert]::ToBoolean($User.'PasswordNeverExpires')
    Write-Host "Creating User " $Displayname " with Username " $SAM -ForegroundColor Green
    New-ADUser -Name "$Displayname" -DisplayName "$Displayname" -SamAccountName $SAM -UserPrincipalName $UPN -GivenName "$UserFirstname" -Surname "$UserLastname" -Description "$Description" -EmailAddress $Email -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) -Enabled $True -Path "$OU" -ChangePasswordAtLogon $PwChgLogin -PasswordNeverExpires $PwExpire
    ForEach ($Group in $Groups)
    {
        Write-Host "Adding " $Displayname " to: " $Groups -ForegroundColor Blue
        Add-ADPrincipalGroupMembership -Identity $SAM -MemberOf $Group
    }
}