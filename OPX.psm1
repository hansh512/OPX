###############################################################################
# Code Written by Hans Halbmayr
# Created On: 14.03.2021
# Last change on: 21.05.2021
#
# Module: OPX 
#
# Version 0.90
#
# Purpose: Loads cmdlets for Exchange server preperation and administration
################################################################################
#requires -version 5

if ($PSVersionTable.PSEdition -eq 'Core')
{
    throw('Unsupported PowerShell edition.')
};
$localPath=Split-Path -Path ($MyInvocation.MyCommand.path ) -Parent;
$Script:cfgDir=[System.IO.Path]::Combine($localPath,'cfg');  # set var for config dir

# list of files to load
$filesToLoad=@(([System.IO.Path]::Combine($localPath,'Libs\cmdlets.ps1')),
               ([System.IO.Path]::Combine($localPath,'Libs\helperfunctions.ps1')),
               ([System.IO.Path]::Combine($localPath,'cfg\Constants.ps1'))
    );# end filesToLad


$fileCount=$filesToLoad.Count;
for ($i=0;$i -lt $fileCount;$i++)
{
    try {
        . $filesToLoad[$i] # load files
    } # end try
    catch {
        Write-Host ('Error: Failed to load code from file ' + $filesToLoad[$i]) -BackgroundColor Black -ForegroundColor Red;
        Write-Host ($_.Exception.Message) -BackgroundColor Black -ForegroundColor Red;
        return; # in case of an error stop
    }; # end catch   
} # end foreach

if (Test-Path -Path $script:VirtualDirCfgFileName)
{
    New-Variable -Name 'VirtualDirCfgFileNamePath' -Value $script:VirtualDirCfgFileName -Scope Script -Option Constant;
} # end if
else {
    # if file not exist, set default
    New-Variable -Name 'VirtualDirCfgFileNamePath' -Value ([system.io.path]::Combine($Script:cfgDir,'VDirs\VdirCfg.xml')) -Scope Script -Option Constant;   
}; # end else

if ([system.string]::IsNullOrEmpty($script:CustomLogFileDir))
{
    $Script:logDir=([System.IO.Path]::Combine($localPath,'logs')); # set var for log dir
} # end if
else {
    $Script:logDir=$script:CustomLogFileDir;
}; # end else

$timeLogType=@('Now','UTCNow'); # time can be logged in local or UTC time
$script:LogTimeUTC=$timeLogType[[int]$script:LogUTCTime]; # set local or UTC time
$script:CallingID=[guid]::NewGuid().guid; # generate GUID

$mName=([System.IO.Path]::GetFileNameWithoutExtension(($MyInvocation.MyCommand.path ))).ToUpper(); # extract the name of the module
$Script:CallingCmdlet=('Module_'+$mName); # name for the calling cmdlet (needed for logging)
$script:pCallingCmdlet=''; # prev cmdlet (is empty for module start)
if ($script:LogUserName)
{
    if (Test-Path -Path variable:PSSenderInfo.ConnectedUser)  # check if constrained endpoint
    {
        $script:LogedOnUser=($PSSenderInfo.ConnectedUser).Split('\')[1];
    } # end if
    else
    {
        $script:LogedOnUser=[System.Security.Principal.WindowsIdentity]::GetCurrent().Name;
    }; # end if
} # end if
else
{
    $script:LogedOnUser='N/A';
}; # end else
# calculate log file path
$logFileDir=[System.IO.Path]::Combine($script:LogDir,(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Replace('\','_')));
if (! (Test-Path -Path $logFileDir)) # verify if directory for log file exist
{
    New-Item -Path ($script:LogDir) -ItemType Directory -Name (([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Replace('\','_')) -Force;
}; # end if
#region init log file name
$LogDateTime=(([System.DateTimeOffset]::($script:timeLogType[[int]$script:LogUTCTime])).ToString())
$script:CurrentLogFileName=([System.IO.Path]::Combine($LogFileDir,('Log_')+($LogDateTime.Split(' ')[0].replace('/','-').replace('.','-'))+'.log'));;
$script:lastLogFileName=$script:CurrentLogFileName;
#endregion init log file name
writeToLog -LogString ('Loading module ' + $mName) -LogType Info;
if (Test-Path -Path variable:'__OPX_ModuleData') # clean up vars if exist
{
    Remove-Variable -Name __OPX_ModuleData;
    Remove-Variable -Name TmpConMsxSrv;
    
}; # end if
setMSXSearchBase; # init varibale for AD search root (for Exchange organization)
New-Variable -Name 'TmpConMsxSrv' -Scope Script;
$Script:cmdletList=[System.Collections.ArrayList]::new(); # init array for cmdlets to export
if ((loadMSXCmdlets -ConMsxSrv ([Ref]$Script:TmpConMsxSrv)) -eq $false) # if Exchange not found, export only a limited set of cmdlets
{
    writeToLog -LogString 'Loading only a limited set of commands.' -LogType Warning -CallingID ('Module_'+$mName);
    [void]$Script:cmdletList.AddRange(@(
            'Get-OPXExchangeSchemaVersion',
            'Save-OPXExchangeServiceStartupTypeToFile',
            'Restore-OPXExchangeServiceStartupType',
            'Get-OPXLastComputerBootTime',
            'Send-OPXTestMailMessages'
        ) # end list of cmdlets
    ); # end Script:cmdletList
} # end if
else
{
    [void]$Script:cmdletList.AddRange(@(
            'Add-OPXMailboxDatabaseCopies',
            'Copy-OPXExchangeCertificateToServers',
            'Get-OPXPreferredServerForMailboxDatabase',
            'Get-OPXExchangeSchemaVersion',
            'Test-OPXExchangeServerMaintenanceState',
            'Get-OPXVirtualDirectories',
            'New-OPXExchangeAuthCertificate',
            'New-OPXExchangeCertificateRequest',
            'Remove-OPXExchangeCertificate',
            'Remove-OPXFailedExchangeServerFromDAG',
            'Resolve-OPXVirtualDirectoriesURLs',
            'Get-OPXExchangeServer',
            'Restore-OPXExchangeServiceStartupType',
            'Save-OPXExchangeServiceStartupTypeToFile',
            'Send-OPXTestMailMessages',
            'Set-OPXVirtualDirectories',
            'Test-OPXExchangeAuthCertificateRollout',
            'Test-OPXMailboxDatabaseMountStatus',
            'Get-OPXExchangeCertificate',
            'Get-OPXVirtualDirectoryConfigurationTemplate',
            'Set-OPXVirtualDirectoryConfigurationTemplate',
            'Remove-OPXVirtualDirectoryConfigurationFromTemplate',
            'Clear-OPXVirtualDirectoryConfigurationTemplate',
            'Restart-OPXExchangeService',
            'Add-OPXKeyToConfigFile',
            'Get-OPXExchangeServerInMaintenance',
            'Get-OPXLastComputerBootTime',
            'Get-OPXCumulatedSizeOfMailboxesInGB'
        ) # end cmdletlist        
    ); # end Script:cmdletList
    if (Get-PSSnapin -Name 'Microsoft.Exchange.Management.PowerShell.E2010' -Registered -ErrorAction SilentlyContinue)
    {
        [void]$Script:cmdletList.AddRange(@(
            'Remove-OPXExchangeServerFromMaintenance',
            'Start-OPXExchangeServerMaintenance',
            'Start-OPXMailboxDatabaseRedistribution'
        )); # end cmdletlist
    }; # end if
}; # end else
if ($script:LoggingTarget -eq 'File')
{
    [void]$cmdletList.Add('Get-OPXLogfileEntries');
}; # end if
writeToLog -LogString ('Exporting ' + ($Script:cmdletList.Count.ToString()) + ' module members') -LogType Info -CallingID ('Module_'+$mName);
Export-ModuleMember -Function $Script:cmdletList;
New-Variable -Name __OPX_ModuleData -Value ([getModuleData]::new());  # init calass
$host.ui.RawUI.WindowTitle=('!!! PSModule ' + $mName + ' !!! Connected to Exchange server ' + $__OPX_ModuleData.ConnectedToMSXServer);
writeToLog -LogString ('Connected to Exchange server ' + $__OPX_ModuleData.ConnectedToMSXServer) -ShowInfo;
Export-ModuleMember -Variable __OPX_ModuleData;
initEnumerator 

# SIG # Begin signature block
# MIIImQYJKoZIhvcNAQcCoIIIijCCCIYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU4OeuW280HWZLbsa90D2DVILg
# l0CgggXxMIIF7TCCBNWgAwIBAgITGwAAAEVNRiQcpqxVgwAAAAAARTANBgkqhkiG
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
# gjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBRQnokRC2R//FSt
# iX3osXxHLiQzojANBgkqhkiG9w0BAQEFAASCAQAntd7YAHrE5P4E6MYVUmVFFjHi
# sEQXwo47/jweGynGaR1VpPWhcjH9cMSM2Lq7i8IGFk1tF/tmj37NeAl2hhYPwHsF
# yHFUYpAoBQPPWjqxZfmX6iHVPScWAQ8SMki7j7876LUe+YsZ8IejkT3uSYKUS4ef
# OEbcCAsIGgOOlxPOiFyStXjDW2Otn8JuXJQPF2o2fBCvfLzAvwxdlfvgxf6bj2/9
# +nSnesVmHJE1/Jepp52tGmUvXi7X32/2MVy2ab4jW15hkcBiVy6se6bzrmDDrCIm
# U/O61DT5UjdEE3mUlpcdp9n83QLrPEFJ2aNcn0N+p3QTQ6GyvKHbWrUWa0nJ
# SIG # End signature block
