if ($args.Count -ne 2) {
    throw "Usage: getAppSecret.ps1 <SsoAppName> <User>"
}

$ssoAppName = $args[0]
$user = $args[1]
[void][Windows.Security.Credentials.PasswordVault, Windows.Security.Credentials, ContentType = WindowsRuntime]
$vault = New-Object Windows.Security.Credentials.PasswordVault
$retrievedSecret = $vault.Retrieve($ssoAppName, $user)
return $retrievedSecret.Password