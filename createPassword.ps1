Read-Host -Prompt "Please enter the password you want to store" -AsSecureString | ConvertFrom-SecureString | Out-File "ftppassword" -Force