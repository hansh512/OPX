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

