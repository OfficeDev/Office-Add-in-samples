
# Get the project item for the scripts folder
try {
    $scriptsFolderProjectItem = $project.ProjectItems.Item("Scripts")
    $projectScriptsFolderPath = $scriptsFolderProjectItem.FileNames(1)
}
catch {
    # No Scripts folder
    Write-Host "No scripts folder found"
}


$packageScriptsFolder = Join-Path $installPath Content\Scripts\Office
Write-Host "packageScriptsFolder" + $packageScriptsFolder

$officeVersionFolder = $packageScriptsFolder | Get-ChildItem | where {$_.PsIsContainer} | Split-Path -Leaf
Write-Host "officeVersionFolder: " + $officeVersionFolder

# $officeVersionFolder -match "(\d+)"
$officeVersion = $officeVersionFolder | Select-String -Pattern "(\d+)"

Write-Host "officeVersion" + $officeVersion

$officeBetaRegEx = "[Oo][Ff][Ff][Ii][Cc][Ee]/[Oo][Ff][Ff][Ii][Cc][Ee]\.[Jj][Ss]"
    
$officeintellisensefile = "_officeintellisense.js"
$officeLocalRegEx = "[Oo][Ff][Ff][Ii][Cc][Ee]/((\d+\.)(\d+))/[Oo][Ff][Ff][Ii][Cc][Ee]\.[Jj][Ss]"
$officeCDNRegEx = "[Hh][Tt][Tt][Pp][Ss]://[Aa][Pp][Pp][Ss][Ff][Oo][Rr][Oo][Ff][Ff][Ii][Cc][Ee]\.[Mm][Ii][Cc][Rr][Oo][Ss][Oo][Ff][Tt]\.[Cc][Oo][Mm]/[Ll][Ii][Bb]/(\d+\.)(\d+)/[Hh][Oo][Ss][Tt][Ee][Dd]/[Oo][Ff][Ff][Ii][Cc][Ee]\.[Jj][Ss]"
$outlook15LocalRegEx = "[Oo][Ff][Ff][Ii][Cc][Ee]/((\d+\.)(\d+))/[Oo][Uu][Tt][Ll][Oo][Oo][Kk]-15\.[Dd][Ee][Bb][Uu][Gg]\.[Jj][Ss]"
$outlook15CDNRegEx = "[Hh][Tt][Tt][Pp][Ss]://[Aa][Pp][Pp][Ss][Ff][Oo][Rr][Oo][Ff][Ff][Ii][Cc][Ee]\.[Mm][Ii][Cc][Rr][Oo][Ss][Oo][Ff][Tt]\.[Cc][Oo][Mm]/[Ll][Ii][Bb]/(\d+\.)(\d+)/[Hh][Oo][Ss][Tt][Ee][Dd]/[Oo][Uu][Tt][Ll][Oo][Oo][Kk]-15\.[Dd][Ee][Bb][Uu][Gg]\.[Jj][Ss]"
$outlookLocalRegEx = "[Oo][Ff][Ff][Ii][Cc][Ee]/((\d+\.)(\d+))/[Oo][Uu][Tt][Ll][Oo][Oo][Kk]-[Ww][Ii][Nn]32\.[Dd][Ee][Bb][Uu][Gg]\.[Jj][Ss]"
$outlookCDNRegEx = "[Hh][Tt][Tt][Pp][Ss]://[Aa][Pp][Pp][Ss][Ff][Oo][Rr][Oo][Ff][Ff][Ii][Cc][Ee]\.[Mm][Ii][Cc][Rr][Oo][Ss][Oo][Ff][Tt]\.[Cc][Oo][Mm]/[Ll][Ii][Bb]/(\d+\.)(\d+)/[Hh][Oo][Ss][Tt][Ee][Dd]/[Oo][Uu][Tt][Ll][Oo][Oo][Kk]-[Ww][Ii][Nn]32\.[Dd][Ee][Bb][Uu][Gg]\.[Jj][Ss]"

$newOfficeLocalPath = "office/$officeVersion/office.js"
$newOutlookLocalPath = "office/$officeVersion/outlook-win32.debug.js"
$newOutlookCDNPath = "https://appsforoffice.microsoft.com/lib/$officeVersion/hosted/outlook-win32.debug.js"
$newOfficeCDNPath = "https://appsforoffice.microsoft.com/lib/$officeVersion/hosted/office.js"

$officeCommentRegEx = "/\* Required to correctly initalize Office.js for intellisense \*/"
$onlineCommentRegEx = "/\* Use online copy of Office.js for intellisense \*/"
$offlineCommentRegEx = "/\* Use offline copy of Office.js for intellisense \*/"

$officeComment = "/* Required to correctly initalize Office.js for intellisense */"
$onlineComment = "/* Use online copy of Office.js for intellisense */"
$offlineComment = "/* Use offline copy of Office.js for intellisense */"
    

function AddOrUpdate-Reference($scriptsFolderProjectItem, $regExPattern, $newFullPath , $commentOut) {
    try {
        $referencesFileProjectItem = $scriptsFolderProjectItem.ProjectItems.Item("_references.js")
    }
    catch {
        # _references.js file not found
        return
    }

    if ($referencesFileProjectItem -eq $null) {
        # _references.js file not found
        return
    }

    $referencesFilePath = $referencesFileProjectItem.FileNames(1)
    $referencesTempFilePath = Join-Path $env:TEMP "_references.tmp.js"    

  
    $addCondition = Select-String $referencesFilePath -pattern $regExPattern -quiet
    #Write-Host "Add condition check $addCondition"
    if ($addCondition -eq $false) {
        #Write-Host "No existing reference found"
	# File has no existing matching reference line
        # Add the full reference line to the beginning of the file
        if ($regExPattern -eq $officeintellisensefile) {
        $officeComment | Add-Content $referencesTempFilePath -Encoding UTF8
	#Write-Host "Add Comment for intellisense"
        }
        elseif ($regExPattern -eq $outlookLocalRegEx) {
        $offlineComment | Add-Content $referencesTempFilePath -Encoding UTF8
	#Write-Host "Add comment for Local Office.js reference"
        }
        elseif ($regExPattern -eq $outlookCDNRegEx) {
        $onlineComment | Add-Content $referencesTempFilePath -Encoding UTF8
	#Write-Host "Add comment for CDN reference"
        }
        if ($commentOut -eq "True"){
        "// /// <reference path=""$newFullPath"" />" | Add-Content $referencesTempFilePath -Encoding UTF8
       	#Write-Host "Add Comment to $newFullPath"
	}
        else {
        "/// <reference path=""$newFullPath"" />" | Add-Content $referencesTempFilePath -Encoding UTF8
	#Write-Host "Add Reference to $newFullPath"
        }
         Get-Content $referencesFilePath | Add-Content $referencesTempFilePath
    }
    else {
        #Write-Host "Existing reference found"
	# Loop through file and replace old file name with new file name
        Get-Content $referencesFilePath | ForEach-Object { $_ -replace $regExPattern, $newFullPath } > $referencesTempFilePath
    }


    # Copy over the new _references.js file
    Copy-Item $referencesTempFilePath $referencesFilePath -Force
    Remove-Item $referencesTempFilePath -Force
}


function Remove-Reference($scriptsFolderProjectItem , $regExPattern) {
    try {
        $referencesFileProjectItem = $scriptsFolderProjectItem.ProjectItems.Item("_references.js")
    }
    catch {
        # _references.js file not found
        return
    }

    if ($referencesFileProjectItem -eq $null) {
        return
    }

    $referencesFilePath = $referencesFileProjectItem.FileNames(1)
    $referencesTempFilePath = Join-Path $env:TEMP "_references.tmp.js"
   
    $removeCondition = Select-String $referencesFilePath -pattern $regExPattern -quiet
    #Write-Host "Remove condition check $removeCondition"
    if ($removeCondition -eq $True) {
        #Write-Host "Removing Reference $regExPattern"
	# Delete the line referencing the file
        Get-Content $referencesFilePath | ForEach-Object { if (-not ($_ -match $regExPattern)) { $_ } } > $referencesTempFilePath

        # Copy over the new _references.js file
        Copy-Item $referencesTempFilePath $referencesFilePath -Force
        Remove-Item $referencesTempFilePath -Force
    }
}


# SIG # Begin signature block
# MIIkGwYJKoZIhvcNAQcCoIIkDDCCJAgCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCATO9USWLW8EGEl
# NLiZGw2NjSWBVkG0L+NDYD/tKzPmAaCCDZIwggYQMIID+KADAgECAhMzAAAAOI0j
# bRYnoybgAAAAAAA4MA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMTQxMDAxMTgxMTE2WhcNMTYwMTAxMTgxMTE2WjCBgzEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjENMAsGA1UECxMETU9Q
# UjEeMBwGA1UEAxMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMIIBIjANBgkqhkiG9w0B
# AQEFAAOCAQ8AMIIBCgKCAQEAwt7Wz+K3fxFl/7NjqfNyufEk61+kHLJEWetvnPtw
# 22VpmquQMV7/3itkEfXtbOkAIYLDkMyCGaPjmWNlir3T1fsgo+AZf7iNPGr+yBKN
# 5dM5701OPoaWTBGxEYSbJ5iIOy3UfRjzBeCtSwQ+Q3UZ5kbEjJ3bidgkh770Rye/
# bY3ceLnDZaFvN+q8caadrI6PjYiRfqg3JdmBJKmI9GNG6rsgyQEv2I4M2dnt4Db7
# ZGhN/EIvkSCpCJooSkeo8P7Zsnr92Og4AbyBRas66Boq3TmDPwfb2OGP/DksNp4B
# n+9od8h4bz74IP+WGhC+8arQYZ6omoS/Pq6vygpZ5Y2LBQIDAQABo4IBfzCCAXsw
# HwYDVR0lBBgwFgYIKwYBBQUHAwMGCisGAQQBgjdMCAEwHQYDVR0OBBYEFMbxyhgS
# CySlRfWC5HUl0C8w12JzMFEGA1UdEQRKMEikRjBEMQ0wCwYDVQQLEwRNT1BSMTMw
# MQYDVQQFEyozMTY0MitjMjJjOTkzNi1iM2M3LTQyNzEtYTRiZC1mZTAzZmE3MmMz
# ZjAwHwYDVR0jBBgwFoAUSG5k5VAF04KqFzc3IrVtqMp1ApUwVAYDVR0fBE0wSzBJ
# oEegRYZDaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljQ29k
# U2lnUENBMjAxMV8yMDExLTA3LTA4LmNybDBhBggrBgEFBQcBAQRVMFMwUQYIKwYB
# BQUHMAKGRWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWlj
# Q29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNydDAMBgNVHRMBAf8EAjAAMA0GCSqG
# SIb3DQEBCwUAA4ICAQCecm6ourY1Go2EsDqVN+I0zXvsz1Pk7qvGGDEWM3tPIv6T
# dVZHTXRrmYdcLnSIcKVGb7ScG5hZEk00vtDcdbNdDDPW2AX2NRt+iUjB5YmlLTo3
# J0ce7mjTaFpGoqyF+//Q6OjVYFXnRGtNz73epdy71XqL0+NIx0Z7dZhz+cPI7IgQ
# C/cqLRN4Eo/+a6iYXhxJzjqmNJZi2+7m4wzZG2PH+hhh7LkACKvkzHwSpbamvWVg
# Dh0zWTjfFuEyXH7QexIHgbR+uKld20T/ZkyeQCapTP5OiT+W0WzF2K7LJmbhv2Xj
# 97tj+qhtKSodJ8pOJ8q28Uzq5qdtCrCRLsOEfXKAsfg+DmDZzLsbgJBPixGIXncI
# u+OKq39vCT4rrGfBR+2yqF16PLAF9WCK1UbwVlzypyuwLhEWr+KR0t8orebVlT/4
# uPVr/wLnudvNvP2zQMBxrkadjG7k9gVd7O4AJ4PIRnvmwjrh7xy796E3RuWGq5eu
# dXp27p5LOwbKH6hcrI0VOSHmveHCd5mh9yTx2TgeTAv57v+RbbSKSheIKGPYUGNc
# 56r7VYvEQYM3A0ABcGOfuLD5aEdfonKLCVMOP7uNQqATOUvCQYMvMPhbJvgfuS1O
# eQy77Hpdnzdq2Uitdp0v6b5sNlga1ZL87N/zsV4yFKkTE/Upk/XJOBbXNedrODCC
# B3owggVioAMCAQICCmEOkNIAAAAAAAMwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNV
# BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
# HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29m
# dCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDExMB4XDTExMDcwODIwNTkw
# OVoXDTI2MDcwODIxMDkwOVowfjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEoMCYGA1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAx
# MTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAKvw+nIQHC6t2G6qghBN
# NLrytlghn0IbKmvpWlCquAY4GgRJun/DDB7dN2vGEtgL8DjCmQawyDnVARQxQtOJ
# DXlkh36UYCRsr55JnOloXtLfm1OyCizDr9mpK656Ca/XllnKYBoF6WZ26DJSJhIv
# 56sIUM+zRLdd2MQuA3WraPPLbfM6XKEW9Ea64DhkrG5kNXimoGMPLdNAk/jj3gcN
# 1Vx5pUkp5w2+oBN3vpQ97/vjK1oQH01WKKJ6cuASOrdJXtjt7UORg9l7snuGG9k+
# sYxd6IlPhBryoS9Z5JA7La4zWMW3Pv4y07MDPbGyr5I4ftKdgCz1TlaRITUlwzlu
# ZH9TupwPrRkjhMv0ugOGjfdf8NBSv4yUh7zAIXQlXxgotswnKDglmDlKNs98sZKu
# HCOnqWbsYR9q4ShJnV+I4iVd0yFLPlLEtVc/JAPw0XpbL9Uj43BdD1FGd7P4AOG8
# rAKCX9vAFbO9G9RVS+c5oQ/pI0m8GLhEfEXkwcNyeuBy5yTfv0aZxe/CHFfbg43s
# TUkwp6uO3+xbn6/83bBm4sGXgXvt1u1L50kppxMopqd9Z4DmimJ4X7IvhNdXnFy/
# dygo8e1twyiPLI9AN0/B4YVEicQJTMXUpUMvdJX3bvh4IFgsE11glZo+TzOE2rCI
# F96eTvSWsLxGoGyY0uDWiIwLAgMBAAGjggHtMIIB6TAQBgkrBgEEAYI3FQEEAwIB
# ADAdBgNVHQ4EFgQUSG5k5VAF04KqFzc3IrVtqMp1ApUwGQYJKwYBBAGCNxQCBAwe
# CgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0j
# BBgwFoAUci06AjGQQ7kUBU7h6qfHMdEjiTQwWgYDVR0fBFMwUTBPoE2gS4ZJaHR0
# cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2Vy
# QXV0MjAxMV8yMDExXzAzXzIyLmNybDBeBggrBgEFBQcBAQRSMFAwTgYIKwYBBQUH
# MAKGQmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljUm9vQ2Vy
# QXV0MjAxMV8yMDExXzAzXzIyLmNydDCBnwYDVR0gBIGXMIGUMIGRBgkrBgEEAYI3
# LgMwgYMwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lv
# cHMvZG9jcy9wcmltYXJ5Y3BzLmh0bTBABggrBgEFBQcCAjA0HjIgHQBMAGUAZwBh
# AGwAXwBwAG8AbABpAGMAeQBfAHMAdABhAHQAZQBtAGUAbgB0AC4gHTANBgkqhkiG
# 9w0BAQsFAAOCAgEAZ/KGpZjgVHkaLtPYdGcimwuWEeFjkplCln3SeQyQwWVfLiw+
# +MNy0W2D/r4/6ArKO79HqaPzadtjvyI1pZddZYSQfYtGUFXYDJJ80hpLHPM8QotS
# 0LD9a+M+By4pm+Y9G6XUtR13lDni6WTJRD14eiPzE32mkHSDjfTLJgJGKsKKELuk
# qQUMm+1o+mgulaAqPyprWEljHwlpblqYluSD9MCP80Yr3vw70L01724lruWvJ+3Q
# 3fMOr5kol5hNDj0L8giJ1h/DMhji8MUtzluetEk5CsYKwsatruWy2dsViFFFWDgy
# cScaf7H0J/jeLDogaZiyWYlobm+nt3TDQAUGpgEqKD6CPxNNZgvAs0314Y9/HG8V
# fUWnduVAKmWjw11SYobDHWM2l4bf2vP48hahmifhzaWX0O5dY0HjWwechz4GdwbR
# BrF1HxS+YWG18NzGGwS+30HHDiju3mUv7Jf2oVyW2ADWoUa9WfOXpQlLSBCZgB/Q
# ACnFsZulP0V3HjXG0qKin3p6IvpIlR+r+0cjgPWe+L9rt0uX4ut1eBrs6jeZeRhL
# /9azI2h15q/6/IvrC4DqaTuv/DDtBEyO3991bWORPdGdVk5Pv4BXIqF4ETIheu9B
# CrE/+6jMpF3BoYibV3FWTkhFwELJm3ZbCoBIa/15n8G9bW1qyVJzEw16UM0xghXf
# MIIV2wIBATCBlTB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQSAyMDExAhMzAAAA
# OI0jbRYnoybgAAAAAAA4MA0GCWCGSAFlAwQCAQUAoIHKMBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqG
# SIb3DQEJBDEiBCAAjoN50utyXHZkcE8hc7selB46Am7buIpZ8prY77Fu5TBeBgor
# BgEEAYI3AgEMMVAwTqAWgBQAQwBvAG0AbQBvAG4ALgBwAHMAMaE0gDJodHRwOi8v
# d3d3Lm51Z2V0Lm9yZy9wYWNrYWdlcy9NaWNyb3NvZnQuT2ZmaWNlLmpzLzANBgkq
# hkiG9w0BAQEFAASCAQCMm5V7sGqSZC9B6MoHpcu37lIc6LR/Hb842wcN/7zW6GCO
# Eku2RyQGJvnCKTJG4sLXzRT1IvGUmIhIgebhUwUpXGxgCNxW2UQXyV12CQUNoXMm
# kd9xmy18kPMTjvMwPgak2E5iFe2Zee5ogeLLdcsy8iEs/gF0WBrS15Hr3KQx7y92
# VNPUSBesbmwLz/jbiEI9L3hTA6rOZBT3qI/0dk+45xxPBipmpUz8lBfp2RaZjV1q
# +EsGeqTO/IpOQMrQ+QRQUfu39YmsK627A2N8cM1fKNX6duBmMZiB2y64PGqJRWD5
# r1IlFXvbN68sFBjaezGleJgUfp4aP8LqnPEx2VpAoYITTTCCE0kGCisGAQQBgjcD
# AwExghM5MIITNQYJKoZIhvcNAQcCoIITJjCCEyICAQMxDzANBglghkgBZQMEAgEF
# ADCCAT0GCyqGSIb3DQEJEAEEoIIBLASCASgwggEkAgEBBgorBgEEAYRZCgMBMDEw
# DQYJYIZIAWUDBAIBBQAEIFXs+vv5+Q173zGvoi0+n0mjbbIQeo+KKL//DXvPr0gT
# AgZVW05VMbkYEzIwMTUwNTI3MDE0OTUxLjA5N1owBwIBAYACAfSggbmkgbYwgbMx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xDTALBgNVBAsTBE1P
# UFIxJzAlBgNVBAsTHm5DaXBoZXIgRFNFIEVTTjpGNTI4LTM3NzctOEE3NjElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaCCDtAwggZxMIIEWaAD
# AgECAgphCYEqAAAAAAACMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBD
# ZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAxMDAeFw0xMDA3MDEyMTM2NTVaFw0yNTA3
# MDEyMTQ2NTVaMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMIIBIjANBgkq
# hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAqR0NvHcRijog7PwTl/X6f2mUa3RUENWl
# CgCChfvtfGhLLF/Fw+Vhwna3PmYrW/AVUycEMR9BGxqVHc4JE458YTBZsTBED/Fg
# iIRUQwzXTbg4CLNC3ZOs1nMwVyaCo0UN0Or1R4HNvyRgMlhgRvJYR4YyhB50YWeR
# X4FUsc+TTJLBxKZd0WETbijGGvmGgLvfYfxGwScdJGcSchohiq9LZIlQYrFd/Xcf
# PfBXday9ikJNQFHRD5wGPmd/9WbAA5ZEfu/QS/1u5ZrKsajyeioKMfDaTgaRtogI
# Neh4HLDpmc085y9Euqf03GS9pAHBIAmTeM38vMDJRF1eFpwBBU8iTQIDAQABo4IB
# 5jCCAeIwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFNVjOlyKMZDzQ3t8RhvF
# M2hahW1VMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIBhjAP
# BgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFNX2VsuP6KJcYmjRPZSQW9fOmhjE
# MFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kv
# Y3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNybDBaBggrBgEF
# BQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9w
# a2kvY2VydHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3J0MIGgBgNVHSABAf8E
# gZUwgZIwgY8GCSsGAQQBgjcuAzCBgTA9BggrBgEFBQcCARYxaHR0cDovL3d3dy5t
# aWNyb3NvZnQuY29tL1BLSS9kb2NzL0NQUy9kZWZhdWx0Lmh0bTBABggrBgEFBQcC
# AjA0HjIgHQBMAGUAZwBhAGwAXwBQAG8AbABpAGMAeQBfAFMAdABhAHQAZQBtAGUA
# bgB0AC4gHTANBgkqhkiG9w0BAQsFAAOCAgEAB+aIUQ3ixuCYP4FxAz2do6Ehb7Pr
# psz1Mb7PBeKp/vpXbRkws8LFZslq3/Xn8Hi9x6ieJeP5vO1rVFcIK1GCRBL7uVOM
# zPRgEop2zEBAQZvcXBf/XPleFzWYJFZLdO9CEMivv3/Gf/I3fVo/HPKZeUqRUgCv
# OA8X9S95gWXZqbVr5MfO9sp6AG9LMEQkIjzP7QOllo9ZKby2/QThcJ8ySif9Va8v
# /rbljjO7Yl+a21dA6fHOmWaQjP9qYn/dxUoLkSbiOewZSnFjnXshbcOco6I8+n99
# lmqQeKZt0uGc+R38ONiU9MalCpaGpL2eGq4EQoO4tYCbIjggtSXlZOz39L9+Y1kl
# D3ouOVd2onGqBooPiRa6YacRy5rYDkeagMXQzafQ732D8OE7cQnfXXSYIghh2rBQ
# Hm+98eEA3+cxB6STOvdlR3jo+KhIq/fecn5ha293qYHLpwmsObvsxsvYgrRyzR30
# uIUBHoD7G4kqVDmyW9rIDVWZeodzOwjmmC3qjeAzLhIp9cAvVCch98isTtoouLGp
# 25ayp0Kiyc8ZQU3ghvkqmqMRZjDTu3QyS99je/WZii8bxyGvWbWu3EQ8l1Bx16HS
# xVXjad5XwdHeMMD9zOZN+w2/XU/pnR4ZOC+8z1gFLu8NoFA12u8JJxzVs341Hgi6
# 2jbb01+P3nSISRIwggTaMIIDwqADAgECAhMzAAAAU8oCK/B0cFZsAAAAAABTMA0G
# CSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTE1
# MDMyMDE3MzIyNloXDTE2MDYyMDE3MzIyNlowgbMxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xDTALBgNVBAsTBE1PUFIxJzAlBgNVBAsTHm5DaXBo
# ZXIgRFNFIEVTTjpGNTI4LTM3NzctOEE3NjElMCMGA1UEAxMcTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgU2VydmljZTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEB
# AM54iPxu7jL0i6DtCPc+aXff2CrpQBDWGw2doNox0B4YXL7jx4/bMMNKBNrGJE4k
# tuU/FnOoC4rYgVWX4R9YEI1GkxHn32gOwHSsHQI6OViRLZI25y74/uh3CTpEmPRV
# +3kVDBGRuXhkwU4mGCaS+6Ph+FnvY5ax2NnjtqHIOxS7GEtjMvucBA9OjR6twB/l
# wc0s6lIK/qjEGLIo0JRPuAkE25oy55RZbEtcNz0p0+64izbFpe2QPkN4ltCtRzRF
# kN5oRopH3qmCb8n0P8DLxZdSA9NHzi4S6kq/xhwzCuV81N2ACXmpYpEiNFQUgwNt
# l+ej7NP8nkjwl+gIGjoYNT8CAwEAAaOCARswggEXMB0GA1UdDgQWBBQ/eyp4u9cj
# NRGndQ5ohPUqiAKp1TAfBgNVHSMEGDAWgBTVYzpcijGQ80N7fEYbxTNoWoVtVTBW
# BgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNUaW1TdGFQQ0FfMjAxMC0wNy0wMS5jcmwwWgYIKwYBBQUH
# AQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# L2NlcnRzL01pY1RpbVN0YVBDQV8yMDEwLTA3LTAxLmNydDAMBgNVHRMBAf8EAjAA
# MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEBCwUAA4IBAQCEAHONgXDn
# RSyPHAYbnO3615WFo65xQrSizlGqG2WVAwSrpEZvnoYMaXAqLfXVvG57F+Rp1d6g
# 90g1Qzsd4f06JNMKMhYJZOch39hAVeZm5x5s2yhrOhej5b1e1sE2D7seUA6CgcJl
# DJEbDePqXEbxiTKDgmTAyn4t9lxQr3LRgUFpZqO42BHDevk8SyZolzAF0k4rtA3h
# 5jDSlbnR85SivAF2Rf16O3vgKShEjDxYfl1/YhxA9UTQO4nJTTxxD1LRCyg02cdc
# g7aJpBTaeXyJipAiMev8rhaVdbgKpEpO0Ua3LCXwiyqYwdPHJiACJFBVr0RU4shh
# e4ejpmp69LE3oYIDeTCCAmECAQEwgeOhgbmkgbYwgbMxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xDTALBgNVBAsTBE1PUFIxJzAlBgNVBAsTHm5D
# aXBoZXIgRFNFIEVTTjpGNTI4LTM3NzctOEE3NjElMCMGA1UEAxMcTWljcm9zb2Z0
# IFRpbWUtU3RhbXAgU2VydmljZaIlCgEBMAkGBSsOAwIaBQADFQDVhi+Wt0SXrds5
# ZjSm4BKLexiVg6CBwjCBv6SBvDCBuTELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBOVFMg
# RVNOOjU3RjYtQzFFMC01NTRDMSswKQYDVQQDEyJNaWNyb3NvZnQgVGltZSBTb3Vy
# Y2UgTWFzdGVyIENsb2NrMA0GCSqGSIb3DQEBBQUAAgUA2Q+MiDAiGA8yMDE1MDUy
# NzAwMjEyOFoYDzIwMTUwNTI4MDAyMTI4WjB3MD0GCisGAQQBhFkKBAExLzAtMAoC
# BQDZD4yIAgEAMAoCAQACAhEqAgH/MAcCAQACAhlcMAoCBQDZEN4IAgEAMDYGCisG
# AQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwGgCjAIAgEAAgMW42ChCjAIAgEAAgMH
# oSAwDQYJKoZIhvcNAQEFBQADggEBAHd9LbIPaTjFe6lzYAkRUnYKeGLWYFq/uhu0
# w1u359Eqel/CxdESmBDtF9oI16W5kWleYlRp0mznb+buCTNM2QYRnM4ZY/X19TNT
# BjhamUhTpZjmYQ8NwmaP97u4NmQqQUPT5BMuV2cX064An9mZ/dYqw9cEhP80CTcn
# rcpc7yXCbRfa7sTjOyBQrIHeCXX+F4dnvuh0kKTHIB7cpoOIBDCwV8eEnBRew+3L
# jbUekCoc0Yhr02cpYRr6/SGgkFrWIemYuOkGhjqYVve1mykp0yu4Q9qaUbqnwBvY
# sFx1FKXNOwZZsdAvguHM0pm+OKGxkt5D0s3gmGfcBYCdq86/ccsxggL1MIIC8QIB
# ATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAFPKAivwdHBW
# bAAAAAAAUzANBglghkgBZQMEAgEFAKCCATIwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3
# DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCDmAVqG5JpShyf01FVcYchOi7wbjTkWvjGL
# 9A8CrAn1DDCB4gYLKoZIhvcNAQkQAgwxgdIwgc8wgcwwgbEEFNWGL5a3RJet2zlm
# NKbgEot7GJWDMIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAC
# EzMAAABTygIr8HRwVmwAAAAAAFMwFgQUtZ+dNr2pmS85u5MhY6q8T3bzK+cwDQYJ
# KoZIhvcNAQELBQAEggEAm560VAcJB99g+FqBgkUeRvL7U7/bF/usF+vGpQF2Ncws
# 9JIR/m2j0kfMZugBCUld5SwHFvGHS2qKHPkg1EGlXDeOc0EFmgjwjHMkkUiyNDE2
# bmjDj584EDkAwJqJFkgCjJzYx+ZojxoxbtRXIdSo2gDdMN3gj6cF/5mMvCCRp+8T
# xafjfk9l1nQin4zr4vcD8ckKd9oNwtavU6L4JXjan/LYp3RisrPzBOpXmt5CVZaa
# r+e7lhg94omjsR0qp2BwqY0YJUHRWwkyJ5iTdFnpg3Rf/btc0Pr20xmBeBUA7E+z
# cseHY9hQGzMaDi906E9db5ul/pVHp8/et/UijJItSQ==
# SIG # End signature block
