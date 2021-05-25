################################################################################
# Code Written by Hans Halbmayr
# Created On: 21.03.2021
# Last change on: 21.05.2021
#
# Module: OPX
#
# Version 0.90
#
# Purpos: Constans for module OPX
################################################################################

# !!! IMPORTANT!!!
# If variables are changed, the module must be removed and loaded again.
# Loading the moudle with the parameter Force may not load the new value(s)

# Target for logging. If a different value as File, the cmdlets will only send some output to screen
New-Variable -Name 'LoggingTarget' -Value 'File' -Option Constant -Scope Script; # supported values File and None
# The time format for logging
New-Variable -Name 'LogTimeFormat' -Value 'dd.MM.yyyy HH:mm:ss' -Option Constant -Scope Script;
# If set to TRUE UTC time will be logged. The time offset will be logged anyway
New-Variable -Name 'LogUTCTime' -Value $false -Option Constant -Scope Script;
# Delimiter for log files
New-Variable -Name 'LogFileCsvDelimiter' -Value ',' -Option Constant -Scope Script;
# If set to TRUE the user name will be written to the log file
New-Variable -Name 'LogUserName' -Value $true -Option Constant -Scope Script;
# For internal use, please dont' change it
New-Variable -Name 'NewGUIDForEverycmdletCall' -Value $false -Option Constant -Scope Script;
# The file name for the additional virtual directory configuration (*-OPXVirtualDirectoryConfigurationTemplate cmdlets and Set/Get-OPXVirtualDirectories). 
# The default vaule 'VdirCfg.xml' points to the file Vdircfg.xml in the direcotry <moduleRoot>\cmfg\Vdirs.
New-Variable -Name 'VirtualDirCfgFileName' -Value '\\HCCMSX10-01.corp.hcc.lab\PSModulesCfg\opx\VDirs\VdirCfg.xml' -Option Constant -Scope Script;
# If set to TRUE all Exchange cmdlets will be imported, if FALSE (default) only the cmdeltes listed in the file <moduleRoot>\cfg\CmdletsToImport.list, are imported (no need to change)
New-Variable -Name 'ImportAllExchangeCmdlets' -Value $false -Option Constant -Scope Script;
# The number of server components are expected to be inactive, if not in maintenance mode
New-Variable -Name 'ServerComponentNotExpectedActive' -Value 2 -Option Constant -Scope Script;
# For debugging purpos, defaut is FALSE
New-Variable -Name 'ExportFaultyServerData'-Value $false -Option Constant -Scope Script;
# For debugging purpos, defaut is FALSE
New-Variable -Name 'ShowDebuggingInfo' -Value $false -Option Constant -Scope Script;
# Time interval in seconds for chacking it the PowerShell job is still running
New-Variable -Name 'JobMonitoringIntervalInSeconds' -Value 10 -Option Constant -Scope Script;
# Force to run scripts in a PowerShell job
New-Variable -Name 'RunScriptsInJob' -Value $true -Option Constant -Scope Script;
# set custom log file directory
New-Variable -Name 'CustomLogFileDir' -Value '\\HCCMSX10-01\PSModuleLogs\OPX' -Option Constant -Scope Script;

# SIG # Begin signature block
# MIIImQYJKoZIhvcNAQcCoIIIijCCCIYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUkG7TtsjoqpmL1G8aqmuvC6wm
# RNSgggXxMIIF7TCCBNWgAwIBAgITGwAAAEVNRiQcpqxVgwAAAAAARTANBgkqhkiG
# 9w0BAQsFADBaMRMwEQYKCZImiZPyLGQBGRYDbGFiMRMwEQYKCZImiZPyLGQBGRYD
# aGNjMRQwEgYKCZImiZPyLGQBGRYEY29ycDEYMBYGA1UEAxMPSENDLUNvcnAtTEFC
# LUNBMB4XDTIxMDQyMDE1MjUzMloXDTIzMDQyMDE1MzUzMlowaDETMBEGCgmSJomT
# 8ixkARkWA2xhYjETMBEGCgmSJomT8ixkARkWA2hjYzEUMBIGCgmSJomT8ixkARkW
# BGNvcnAxDjAMBgNVBAMTBVVzZXJzMRYwFAYDVQQDEw1BZG1pbmlzdHJhdG9yMIIB
# IjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAv2Nhp92kDUopvTSIrL9yGU4z
# e5O1VfefmUqIUnA5+LoNvT2qq5d5jTKNVrXee1WArdXcL7P9n7T0sSuCB5b77/yE
# qEKKM0SmvME+0tImwMZs1XmfnQWWYFA7dss2tLFT0A0NqYJqGkpwLh/AwNAt4y5S
# ohrI4KIMDMzKEAPXfYKT8F+z4CBJCZlj9ZXKT2Z8XRFE1/zNrH16/jNX3wCJCOu9
# y9O0xGpXKK9nNE0rKXs9ebxgxtG0hfj+Y9O0wHTQig4iaCQFrUvHNtm3baV/oGGR
# zHI+7ZBdX+4+7cTCEa1J87qgtG6Pk8X8oWC4+D52kGscqh2rXmg/qG8+QnKxHwID
# AQABo4ICnDCCApgwPQYJKwYBBAGCNxUHBDAwLgYmKwYBBAGCNxUIgvLoV4TXwTOH
# 2YcUg4fvZ4HVqxJzhtmRO4KwqBwCAWUCAQAwEwYDVR0lBAwwCgYIKwYBBQUHAwMw
# DgYDVR0PAQH/BAQDAgeAMBsGCSsGAQQBgjcVCgQOMAwwCgYIKwYBBQUHAwMwHQYD
# VR0OBBYEFPe7LSN2so+S1CqRmPxzA7O8kiAjMB8GA1UdIwQYMBaAFFaiGKwi0xVF
# pzjh58WJVwkxRp+IMIHVBgNVHR8Egc0wgcowgceggcSggcGGgb5sZGFwOi8vL0NO
# PUhDQy1Db3JwLUxBQi1DQSxDTj1IQ0NNU1hQS0kwMSxDTj1DRFAsQ049UHVibGlj
# JTIwS2V5JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29uZmlndXJhdGlvbixE
# Qz1jb3JwLERDPWhjYyxEQz1sYWI/Y2VydGlmaWNhdGVSZXZvY2F0aW9uTGlzdD9i
# YXNlP29iamVjdENsYXNzPWNSTERpc3RyaWJ1dGlvblBvaW50MIHFBggrBgEFBQcB
# AQSBuDCBtTCBsgYIKwYBBQUHMAKGgaVsZGFwOi8vL0NOPUhDQy1Db3JwLUxBQi1D
# QSxDTj1BSUEsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049U2VydmljZXMs
# Q049Q29uZmlndXJhdGlvbixEQz1jb3JwLERDPWhjYyxEQz1sYWI/Y0FDZXJ0aWZp
# Y2F0ZT9iYXNlP29iamVjdENsYXNzPWNlcnRpZmljYXRpb25BdXRob3JpdHkwNQYD
# VR0RBC4wLKAqBgorBgEEAYI3FAIDoBwMGkFkbWluaXN0cmF0b3JAY29ycC5oY2Mu
# bGFiMA0GCSqGSIb3DQEBCwUAA4IBAQBzknd62mplja6+1yqQKxiU0Zvl2+4k71HI
# IMo/RjFC6E33JeyUgxpXdMOnnRfZ/RWmLmusarBfvsE/U7JWq5RPcL3UqRd187uw
# M+KVnWC6yitI2d3qSjYcNMuVlb+npHoMhja1dVnofNJvio5b++XJ8q80MjF5yYxu
# ls4JpPBX9G0FkhJYm4XRHd/bpLBcatT2v+hjgpgiJ9d+sPo36RhIRJsww1NYIHWk
# 5MhbPTXr1AebnMNVX14QfbgJBsZf0KX0atTqRuagG6MYeZZSMN5KAL8mV6GHNeWG
# /B8NGeM1QHASKipY5n/oq+eP0OYSY5eh2xMjYk6bMmFRTFGiMOIRMYICEjCCAg4C
# AQEwcTBaMRMwEQYKCZImiZPyLGQBGRYDbGFiMRMwEQYKCZImiZPyLGQBGRYDaGNj
# MRQwEgYKCZImiZPyLGQBGRYEY29ycDEYMBYGA1UEAxMPSENDLUNvcnAtTEFCLUNB
# AhMbAAAARU1GJBymrFWDAAAAAABFMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEM
# MQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQB
# gjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQdLKIt4N1IYhKn
# 3M5pT9AFt/wHvzANBgkqhkiG9w0BAQEFAASCAQAQs02Xiih9eESu7fjiWoazRxA9
# 5r0ObY8SBBidbLmqcDb4MJsPL4FnHXF6fQI0Z9Du4NKPC6jJJu+d2s3SUhGnnqUT
# JwEcT3594/CXx4oNdzoyZf4SCv0fBHSJpW1h4pttU4dkJ5rX5IF4k5CarQndn0Yo
# A6AOfUakbfWqUK25x/gSsYNzoEyeMEzGpev38/DMQy8zPdSh+YDfDu3mxoDe/1x6
# 9Xh8Gd/JnTuzubvRi74Y1viFRjDSjN7v0oWjGmDb9fFMQRF2NHhF6iueqkn4AWi2
# OpZYO/DIlTuC2pe9HB72T8UU8paLltUdpiwbJc1ERNXIMjL0gZT05W/pdjBl
# SIG # End signature block
