$cred = get-credential
$UserName = $cred.UserName.Substring( 1, $cred.UserName.Length-1 )
$Password = $cred.Password | ConvertFrom-SecureString
"[Auth]" | Set-Content "Auth.ini"
"UserName=" + $UserName | Add-Content "Auth.ini"
"Password=" + $Password | Add-Content "Auth.ini"
