# Install required modules
	#Install-Module CredentialManager
	#Install-Module ActiveDirectory
	
# Get stored credentials from Credential Manager
$cred = Get-StoredCredential -Target '' #complete with name of stored credential in credential manager

# Set text encoding
$textEncoding = [System.Text.Encoding]::UTF8

#Get users from AD that are Enabled, have disabled option PasswordNeverExpires and Password is not Expired
$users = Get-ADUser -Filter * -Properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress |
    Where-Object {$_.Enabled -eq "True"} |
    Where-Object {$_.PasswordNeverExpires -eq $false} |
    Where-Object {$_.PasswordExpired -eq $false}

#Get the Domain Password Policy for users that we are not able to read password policy
$DefaultmaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge

foreach ($user in $users)
{
    $Name = $user.Name
    $emailaddress = $user.emailaddress
    $passwordSetDate = $user.PasswordLastSet
    $PasswordPol = (Get-AduserResultantPasswordPolicy $user)
    if (($PasswordPol) -ne $null) {$maxPasswordAge = ($PasswordPol).MaxPasswordAge}
	else {$maxPasswordAge = $DefaultmaxPasswordAge}

# Calculate days to password expiration
    $expireson = $passwordsetdate + $maxPasswordAge
    $today = (get-date)
    $daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days

# Construct message based on days to expiration
    if (($daystoexpire) -gt "1")
    {
        $messageDays1 = "wygaśnie za " + "<b>$daystoexpire</b>" + " dni."
		$messageDays2 = "wygaśnie za " + "$daystoexpire" + " dni."
    }
	else
    {
        $messageDays1 = "wygasa w dniu <b>DZISIEJSZYM</b>."
		$messageDays2 = "wygasa w dniu DZISIEJSZYM."
    }

    $subject="Twoje hasło $messageDays2"
    $body ="
    Cześć $name.
    <p> Wygląda na to że Twoje hasło domenowe $messageDays1<br><br>
    By zmienić Twoje hasło do Windowsa (komputera), najpierw połącz się z siecią firmową<br>
	&emsp;&emsp;Możesz to zrobić:<br>
	<i>&emsp;&emsp;&emsp;&emsp;a) w firmie po kablu sieciowym lub przez sieć WiFi Border</i><br>
	<i>&emsp;&emsp;&emsp;&emsp;b) poza firmą po zestawieniu VPNa</i><br>
	po połączeniu nacisnij <b>CTRL</b>+<b>ALT</b>+<b>Delete</b> i wybierz Zmień Hasło <br><br>
	<p>Przypominam, że połączenie VPN jak i dostęp do zasobów (np. INFOMEX) czy dostep do panelu Helpdesk jest skorelowane z Twoim kontem Windows. Oznacza to, że dane do autoryzacji w tych miejscach są pobierane z danych Twojego konta Windows.<br> 
	Dlatego tak ważne jest zmienienie hasła przed jego wygaśnięciem. Nie zmienienie hasła przed jego wygaśnięciem zablokuje dostęp do wyżej wymienionych usług. 
    <p>Pozdrawiam, Helpdesk IT<br>
    </P>"

    if (($daystoexpire -ge "0") -and (($daystoexpire -eq "1") -or ($daystoexpire -eq "3") -or ($daystoexpire -eq "7")))
    {
		$mailParams = @{
			SmtpServer		= 'smtp.office365.com' #complete with smtp server
			Port			= '587' # or '25' if not using TLS
			UseSSL			= $true # or not if using non-TLS
			Credential		= $cred
			From			= '' #complete with from address
			To			= $emailaddress
			Subject			= $subject
			Body			= $body
			Priority		= 'High'
			bodyasHTML		= $true
		}
		
		Send-Mailmessage @mailParams -Encoding $textEncoding
		
    }
}
