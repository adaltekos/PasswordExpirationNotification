# Password Expiration Notification Script

This PowerShell script retrieves a list of enabled Active Directory users whose password is not set to never expire and whose password has not yet expired. For each user, the script calculates the number of days until the password expires based on the maximum password age policy, and sends an email notification if the number of days matches a specified set of values (0, 1, 3, or 7).

The notification email reminds the user to change their password before it expires, provides instructions on how to change the password, and warns about the consequences of not changing the password before expiration.

## Prerequisites
- PowerShell 5.1 or later
- Installed CredentialManager module
- Installed ActiveDirectory module

## Configuration
- Set up stored credentials in Credential Manager with the name of the user account that has permission to access Active Directory.
- Update the following variables:
  - `$cred`: Set to the name of the stored credential in Credential Manager.
  - `$DefaultmaxPasswordAge`: Set to the default maximum password age for users that we are not able to read the password policy.
  - `$mailParams`: Set the SMTP server, the sender email address, and the email content according to your needs.
