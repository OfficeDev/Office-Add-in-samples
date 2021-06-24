if ($args.Count -ne 3) {
    throw "Usage: addAppSecret.ps1 <SsoAppName> <User> <Secret>"
}

$ssoAppName = $args[0]
$user = $args[1]
$secret = $args[2]
[void][Windows.Security.Credentials.PasswordVault, Windows.Security.Credentials, ContentType = WindowsRuntime]
$creds = New-Object Windows.Security.Credentials.PasswordCredential
$creds.Resource = $ssoAppName
$creds.UserName = $user
$creds.Password = $secret
$vault = New-Object Windows.Security.Credentials.PasswordVault
$vault.Add($creds)
