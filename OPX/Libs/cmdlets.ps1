################################################################################
# Code Written by Hans Halbmayr
# Created On: 14.03.2021
# Last change on: 07.05.2021
#
# Module: OPX
#
# Version 0.80
#
# Purpos: Module for Exchange server configuration and administration
################################################################################
#Region Set strict mode
Set-StrictMode -Version Latest
#EndRegion Set strict mode


# Tested on:
# Exchange 2016 CU 19
# Exchange 2019 CU6
# Windows Server 2016 and PowerShell 5.1
# Windows Server 2019 and PowerShell 5.1

#Requires -Version 5.1


function Remove-OPXFailedExchangeServerFromDAG
{
<#
.SYNOPSIS 
Remove the configuration of a failed Exchange server from a DAG
	
.DESCRIPTION
Remove-OPXFailedExchangeServerFromDAG removes the configuration for the database copies and the DAG membership of the server. In addition, a file with the database copy configuration will be created. After the server was installed with the switch RecoverServer, the server can be added to the DAG and the database copies can be configured with the command Add-OPXMailboxDatabaseCopies, which uses the file with the configuration of the database copies.
If the server is in a healthy state, the removal will fail.

.PARAMETER Server
The parameter is mandatory.
The name of the failed server. You can type the name of the server or step with TAB through the list of your Exchange mailbox servers.
	
.PARAMETER CreateDatabaseListFileOnly
The parameter is optional.
Use this parameter if you only want to get the database copy configuration. No configuration for database copy or DAG membership will be removed.
It is recommended to run the command with this parameter and copy the file to a save location before removing the configuration.

.PARAMETER DatabaseListFilePath
The parameter is mandatory.
The path to the file for the configuration of the database copy. It the file exist, the command will stop. 

.PARAMETER Force
The Parameter is optional.
If the file with the configuration of the database copies cannot be created, the removal of the configuration will fail, unless you use the parameter Force.

.PARAMETER CSVDelimiter
The parameter is optional.
With this parameter you can determine the delimiter for the CSV file. Default is the delimiter configured in the Constants.ps1, in the cfg directory of the module.

.EXAMPLE
Remove-OPXFailedExchangeServerFromDAG -Server <ServerName> -CreateDatabaseListFileOnly -DatabaseListFilePath <full file path for the csv file>
The example only creates configuration file for the database copies. No configuration will be removed.
	
.EXAMPLE
Remove-OPXFailedExchangeServerFromDAG -Server <ServerName> -DatabaseListFilePath <full file path for the csv file
The example removes the configuration of the database copies and the server will be removed from the DAG. If the server is online, the removal will fail.
#>
[cmdletbinding(SupportsShouldProcess=$true)]
param([Parameter(Mandatory = $true, Position = 0)][ArgumentCompleter( {
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($true,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*" });
      } )]
      [String]$Server,
      [Parameter(Mandatory = $false, Position = 1)][switch]$CreateDatabaseListFileOnly=$false,
      [Parameter(Mandatory = $true, Position = 2)][string]$DatabaseListFilePath,
      [Parameter(Mandatory = $false, Position = 3)][switch]$Force=$false,
      [Parameter(Mandatory = $false, Position = 4)][string]$CSVDelimiter=$Script:LogFileCsvDelimiter
     )
     
    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin 
    
    process { 
        if ($PSBoundParameters.ContainsKey('WhatIf'))
        {
            $msg='WhatIf is present, no configuration will be removed and no configuration will be written to file.';
            writeToLog -LogString $msg;
            Write-Verbose $msg;
        }; # end if
        $dirPath=[System.IO.Path]::GetDirectoryName($DatabaseListFilePath);
        $msg=('Verifying if directory ' + $dirPath + ' exist');
        Write-Verbose $msg;
        writeToLog -LogString $msg;
        if (!(Test-Path -Path $dirPath -PathType Container))
        {
            writeToLog -LogString ('The directory ' + $dirPath + ' does not exist, stopping.') -LogType Error;
            return;
        }; # end if
        if (Test-Path -Path $DatabaseListFilePath -PathType Leaf)
        {
            writeToLog -LogString ('The file ' + $DatabaseListFilePath + ' exist. Please copy the file to a save location and run the command again.') -LogType Warning;
            return;
        }; # end if
        Write-Verbose ('Verifying if server ' + $Server + ' is member of a DAG.');
        writeTolog -LogString ('Verifying if server ' + $Server + ' is member of a DAG.');
        try {
            $DAGName = (Get-MailboxServer -Identity $Server -ErrorAction Stop).DatabaseAvailabilityGroup;
            if ([System.String]::IsNullOrEmpty($DAGName))
            {   
                writeToLog -LogType Warning -LogString  ('The server ' + $server + ' is not a DAG member.');
                return;
            }; # end if
        } # end try
        catch {
            writeToLog -LogType Error -LogString  ('Server ' + $Server + ' not found or not member of a DAG.');
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
            return;
        }; # end catch
            
        if ($CreateDatabaseListFileOnly)
        {
            Write-Verbose 'Skipping removing database copies, creating file with list of database copies only.';
            writeTolog -LogString 'Skipping removing database copies, creating file with list of database copies only.';
        } # end if
        else
        {
            $msg=('Verifying if server ' + $Server + ' is online.');
            Write-Verbose $msg;
            writeTolog -LogString $msg;
            try {
                [void](Test-ServiceHealth -Server $Server -ErrorAction stop)
                writeToLog -LogType Warning -LogString  ('The server ' + $Server + ' is operational. Cannot remove a running server from DAG.');
                return;
            } # end try    
            catch {
            # do nothing, faild to connect to server (server down, is expected because a failed server should be removed)
            writeTolog -LogString ('Verifyed that server ' + $Server + ' is down.');
            }; # end catch
        }; # end else
        $fieldList=@(
            @('ServerName',[System.String]),
            @('DBName',[System.String]),
            @('ActivationPreference',[System.Int16]),
            @('ReplayLagTimes',[System.String]),
            @('TruncationLagTimes',[System.String])
        ); # end fieldList
        $DBTable=createNewTable -TableName 'Databases' -FieldList $fieldList;
        Write-Verbose ('Building list of databases on server ' + $server);
        writeTolog -LogString ('Building list of databases on server ' + $server);
        $DBList = @(Get-MailboxDatabase -Server $server);  # get list of DBs
        
        foreach ($DB in $DBList)                       # iterate through the list of DBs
        {
            $activationPreference=($DB.ActivationPreference | ForEach-Object {[string]$_}).TrimStart('[').TrimEnd(']').Replace(' ','');
            $ReplayLagTimes=($DB.ReplayLagTimes | ForEach-Object {[string]$_}).TrimStart('[').TrimEnd(']').Replace(' ','');
            $TruncationLagTimes=($DB.TruncationLagTimes | ForEach-Object {[string]$_}).TrimStart('[').TrimEnd(']').Replace(' ','');
            $srvCount=$activationPreference.count;
            for ($i=0;$i -lt $srvCount;$i++)  # for each server with a mail box database copy
            {
                $dataList=@(
                    $activationPreference[$i].split(',')[0],
                    [string]$DB.Identity,
                    $activationPreference[$i].split(',')[1],
                    $ReplayLagTimes[$i].Split(',')[1],
                    $TruncationLagTimes[$i].Split(',')[1]
                ); # end dataList            
                [void]$DBTable.rows.Add($dataList);
            }; # end foreach                   
        }; # end foreach

        Write-Verbose ('Writing database copy configuration to file ' + $DatabaseListFilePath);
        writeToLog -LogString ('Found ' + $dbList.count.ToString() + ' databases on server ' +  $Server);
        writeTolog -LogString ('Writing database copy configuration to file ' + $DatabaseListFilePath);
        try {
            $DBTable | Export-Csv -Path $DatabaseListFilePath -NoTypeInformation -Delimiter $CSVDelimiter;
        } # end try
        catch {
            writeToLog -LogType Error -LogString  'Failed to export the list of database copies.';
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
            $DBTable;
            if ($Force.IsPresent -eq $false)
            {                
                return; # if removal of configuration is not forced, stopp and return
            }; # end if       
        }; # end catch  
        if ($CreateDatabaseListFileOnly.IsPresent -eq $false)
        {
            try
            {
                Write-Verbose ('Removing configuration for malbox databases from server ' + $server);
                writeTolog -LogString ('Removing configuration for malbox databases from server ' + $server);
                foreach ($db in $dblist)
                {
                    try {
                        $dbIdentity=($db.name +'\'+($Server.split('.'))[0]);
                        Write-Verbose ('Removing database copy ' + $dbIdentity);
                        writeTolog -LogString ('Removing database copy ' + $dbIdentity);
                        Remove-MailboxDatabaseCopy -Identity $dbIdentity -ErrorAction Stop -WhatIf:($WhatIfPreference);   # remove database copy
                    } # end try
                    catch {
                        writeToLog -LogType Error -LogString  ('Failed to remove the database copy ' + $dbIdentity);
                        writeToLog -LogString ($_.Exception.Message) -LogType Error;
                    }; # end catch
                }; # end foreach 
                try {
                    $msg=('Removing configuration for mailbox server ' + $server + ' from DAG ' + $DAGName);
                    Write-Verbose $msg;
                    writeTolog -LogString $msg;
                    $pList=@{
                        Identity=$DAGName;
                        MailboxServer=$Server;
                        ConfigurationOnly=$true;
                        WhatIf=($WhatIfPreference);
                        ErrorAction='Stop';
                    }; # end pList
                    if ($WhatIfPreference)
                    {
                        $pList.ErrorAction='SilentlyContinue';
                    }; # end if
                    if ($PSBoundParameters.ContainsKey('WhatIf'))
                    {
                        writeToLog -LogString ('What if: Removing server from DAG on target ' + $DAGName) -ShowInfo;
                    } # end if
                    else
                    {
                        Remove-DatabaseAvailabilityGroupServer @pList;   # remove the failed server from the DAG
                    }; # end else
                    try
                    {
                        Write-Verbose ('Evicting cluster node ' + $Server);
                        writeTolog -LogString ('Evicting cluster node ' + $Server);
                        if ($PSBoundParameters.ContainsKey('WhatIf'))
                        {
                            writeToLog -LogString ('What if: Removing server from cluster on target ' + $DAGName) -ShowInfo;
                        } # end if
                        else
                        {
                            Invoke-Command -ComputerName ($__OPX_ModuleData.ConnectedToMSXServer) -ScriptBlock {Param($server,$wfp) Remove-ClusterNode -Name $Server -ErrorAction Stop  -WhatIf:($wfp)} -ArgumentList $server,$WhatIfPreference;
                        }; # end else
                    } # end try
                    catch
                    {
                        writeToLog -LogType Error -LogString  ('Failed to evict cluster node ' +$Server);
                        writeToLog -LogString ($_.Exception.Message) -LogType Error;
                    }; # end catch
                } # end try
                catch {
                    writeToLog -LogType Error -LogString  ('Failed to remove the mailbox server ' + $Server + ' from the DAG ' + $DAGName);
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;
                }; # end catch
            } # end try
            catch
            {
                writeToLog -LogType Error -LogString  ('Faeild to remove the database copy ' + $DB + '\' + $Server);
                writeToLog -LogString ($_.Exception.Message) -LogType Error;
                writeToLog -LogString 'The mailbox database copies were not removed successfully. Please remove the copies manually and if successfull please run the following command:' -LogType Warning;
                writeToLog -LogString ('Remove-DatabaseAvailabilityGroupServer -Identity ' + $DAGName + ' -MailboxServer ' + $Server + ' -ConfigurationOnly') -LogType Warning;
                writeToLog -LogString ('Evict the node from cluster with the command: Remove-ClusterNode -Name ' + $server) -LogType Warning;
            }; # end catch
        }; # end if
    }; # end process
    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
} # end function Remove-FailedExchangeServer

function Add-OPXMailboxDatabaseCopies
{
<#
.SYNOPSIS 
Add database copies to a server in a DAG.
	
.DESCRIPTION
The command adds database copies to a recovered DAG member server. To add the database copies, a configuration file is required. The command Remove-OPXFailedExchangeServerFromDAG creates the appropriate configuration file.

.PARAMETER InputFile
The parameter is mandatory.
The file path to the configuration file created with the command Remove-OPXFailedExchangeServerFromDAG.
	
.PARAMETER Server
The parameter is mandatory.
The name of the (recovered) server where the database copies should be added.

.PARAMETER DomainController
The parameter is optional.
The name of a domain controller which should be used for adding the database copies.

.PARAMETER ConfigurationOnly
The parameter is optional.
The parameter allows to add the database copies without invoking seeding.

.PARAMETER CSVDelimiter
The parameter is optional.
With this parameter you can determine the delimiter for the CSV file. Default is the delimiter configured in the Constants.ps1, in the cfg directory of the module.

.EXAMPLE
Add-OPXFailedMailboxDatabaseCopies -Server <ServerName> -InputFile <path to the input file>
The mailbox database copies will be added and seeding will be invoked.
.EXAMPLE
Add-OPXFailedMailboxDatabaseCopies -Server <ServerName> -InputFile <path to the input file> -ConfigurationOnly
The mailbox database copies will be added without invoking seeding.
#>
[CmdLetBinding()]
param([Parameter(Mandatory = $true, Position = 0)][String]$InputFile,
      [Parameter(Mandatory = $true, Position = 1)]
      [ArgumentCompleter( {   
        param ( $CommandName,
                $ParameterName,
                $WordToComplete,
                $CommandAst,
                $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($true,$false,$true));
        $srvList.Where({ $_ -like "$wordToComplete*"}); 
      } )]
      [String]$Server,
      [Parameter(Mandatory = $false, Position = 3)][Switch]$OnlyDisplayConfiguration = $false,
      [Parameter(Mandatory = $false, Position = 4)][Switch]$RecoverServer=$false,
      [Parameter(Mandatory = $false, Position = 5)][string]$DomainController,
      [Parameter(Mandatory = $false, Position = 5)][switch]$ConfigurationOnly=$false,
      [Parameter(Mandatory = $false, Position = 6)][string]$CSVDelimiter=$Script:LogFileCsvDelimiter
     )
    
    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        try
        {
            Write-Verbose "Importing data from file $InputFile.";
            writeTolog -LogString "Importing data from file $InputFile.";
            $dbCopyList=(Import-Csv -Path $InputFile -Delimiter ',').where({($_.ServerName -eq $Server) -and ([int]$_.ActivationPreference -ne (1 -band [int]($RecoverServer.IsPresent -eq $false)))});            
        } # end try
        catch
        {
            writeToLog -LogType Error -LogString  "Failed to read the input file $InputFile.";
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
            return;
        }; # end catch
        
        $aPList=@{};
        if($PSBoundParameters.ContainsKey('DomainController'))
        {        
            Write-Verbose ('Using domain controller ' + $DomainController);
            writeTolog -LogString ('Using domain controller ' + $DomainController);
            if (($DomainController=testIfFQDN -ComputerName $DomainController) -eq $false)
            {
                return;
            }; # end if
            if (! (istDCOnline -DomainController $DomainController))
            {
                writeToLog -LogType Error -LogString  ('Failed to connect to DC ' + $DomainController);
                return;
            }; #end if
            $aPList.Add('DomainController',$DomainController);
        }; # end if
        if($PSBoundParameters.ContainsKey('ConfigurationOnly'))
        {
            Write-Verbose ('Using the parameter ConfigurationOnly for adding the mailbox database copies (Add-MailboxDatabaseCopy)');
            writeTolog -LogString ('Using the parameter ConfigurationOnly for adding the mailbox database copies (Add-MailboxDatabaseCopy)');
            $aPList.Add('ConfigurationOnly',$true);
        }; # end if

        foreach ($dbCopy in $dbCopyList)
        {        
            try
            {
                if ($OnlyDisplayConfiguration)
                {
                    writeTolog -LogString ("Server: " + $dbCopy.ServerName + ", DB: " + $dbCopy.DBName + ", ReplayLagTime: " + $dbCopy.ReplayLagTimes + ", TruncationLagTime: " + $dbCopy.TruncationLagTimes + ", ActivationPreference: " + $dbCopy.ActivationPreference) -LogType Info -ShowInfo;
                } # end if
                else
                {
                    $dbCopyName=($dbCopy.dbName + '\' + $dbCopy.ServerName);
                    Write-Verbose ("Creating copy for database " + $dbCopy.DBName + " on server " + $dbCopy.ServerName + "...");
                    writeTolog -LogString ("Creating copy for database " + $dbCopy.DBName + " on server " + $dbCopy.ServerName);
                    if (Get-MailboxDatabaseCopyStatus -Identity $dbCopyName -ErrorAction silentlycontinue) # verify that copy does not exist
                    {
                        writeTolog -LogString ('A mailbox database copy with the identity ' + $dbCopyName + ' already exist.') -LogType Warning;
                    } # end if
                    else {
                        $addParams=@{
                            MailboxServer=$dbCopy.ServerName;
                            Identity=$dbCopy.DBName;
                            ActivationPreference=$dbCopy.ActivationPreference;
                            ErrorAction='Stop';
                        }; # end addParams

                        switch ($dbCopy) # verify if ReplayLagTime or TruncationLagTime is not 0
                        {
                            {[int](($_.ReplayLagTimes).Replace(':','').Replace('.','')) -gt 0} {
                                $addParams.Add('ReplayLagTime',$dbCopy.ReplayLagTimes)
                            };
                            {[int](($_.TruncationLagTimes).Replace('.','').Replace(':','')) -gt 0 } {
                                $addParams.Add('TruncationLagTime',$dbCopy.TruncationLagTimes)
                            };
                        }; # end switch
                        Add-MailboxDatabaseCopy @addParams @aPList;  
                    }; # end else                              
                }; # end else
            } # end try
            catch
            {
                writeToLog -LogType Error -LogString  ("Failed to add the database copy " + $dbCopy.DBName + " on server " + $dbCopy.ServerName);
                writeToLog -LogString ($_.Exception.Message) -LogType Error;
            }; # end catch
        } # end foreach
    }; # end process
    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
} # end function Add-OPXMailboxDatabaseCopies


function Save-OPXExchangeServiceStartupTypeToFile
{
<#
.SYNOPSIS 
Saves the startup type information for a set of services to a file.
	
.DESCRIPTION
The command saves the startup type of a set of services to a file. Per default the startup type of most of the Exchange relevant services are saved to a file. The startup type of the Exchange relevant services is changed when updates are implemented. When the update fails, it can occur, that the startup type of the services is not restored. With this command the appropriate configuration file can be created. The list of services, for which the startup type will be stored, can be found in the file Exchange.csv. The file can be found under the <module root directory>\cfg\Services. In this directory additional files with services can be created.

.PARAMETER ServiceFileDirectoryPath
The parameter is mandatory.
The directory path where the configuration is stored. 

.PARAMETER CompterName
The parameter is optional.
The parameter excepts input from the pipeline. 
The name of the computer from which the service startup configuration is saved.

.PARAMETER AllServices
The parameter is optional.
If this parameter is used, the startup type of all services is saved.

.PARAMETER ServiceConfig
The parameter is optional.
If you have created a CSV configuration file with a customized set of services, which is stored under <module root>\cfg\Services you can use this parameter to point to your customized configuration file.

.EXAMPLE
Save-OPXExchangeServiceStartupType -Computer <ComputerName> -ServiceFileDirectoryPath <path to the directory where the configuration is saved>
The Exchange relevant services on a given computer will be saved.
.EXAMPLE
Save-OPXExchangeServiceStartupType -AllServices -ServiceFileDirectoryPath <path to the directory where the configuration is saved>
All services on the current computer will be saved.
#>

[cmdletbinding(DefaultParametersetName='Default')]
param([Parameter(Mandatory = $true, Position = 0)][string]$ServiceFileDirectoryPath,
      [Parameter(Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position = 1)][string]$ComputerName=[System.Net.Dns]::GetHostByName('').HostName,
      [Parameter(ParametersetName='AllServices')]
      [Parameter(Mandatory = $false, Position = 2)][switch]$AllServices=$false,
      [Parameter(ParametersetName='SetOfServices')]
      [Parameter(Mandatory = $false, Position = 2)]      
      [ArgumentCompleter( {             
        $fileList=(Get-ChildItem -File -Path $([System.IO.Path]::Combine($__OPX_ModuleData.cfgDir,'Services')))       
        foreach ($item in $fileList) {        
            $path = $item.FullName
            if ($path -like '* *') 
            { 
                $path = "'$path'"
            }; # end if
            [Management.Automation.CompletionResult]::new([System.IO.Path]::GetFileNameWithoutExtension($path), [System.IO.Path]::GetFileNameWithoutExtension($path), 'ParameterValue', [System.IO.Path]::GetFileNameWithoutExtension($path));                 
        }; # end foreach       
      } )]    
      [string]$ServiceConfig
      )

    begin
    {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    Process 
    {
        try {
            if (!(Test-Path -Path $ServiceFileDirectoryPath -PathType Container))
            {
                writeToLog -LogString ($ServiceFileDirectoryPath + ' is not a valid and/or existing directory path.') -LogType Error;
                return;
            }; # end if
            if (($computerName=testIfFQDN -ComputerName $ComputerName) -eq $false)
            {
                return;
            };
            Write-Verbose ('Collecting services from computer ' + $ComputerName);
            writeTolog -LogString ('Collecting services from computer ' + $ComputerName);
            switch ($PSCmdlet.ParameterSetName)
            {
                {$_ -eq 'Default'}
                {
                    try {
                        saveServiceInfo -ServiceConfig 'Exchange' -ComputerName $ComputerName -ServiceFileDirectoryPath $ServiceFileDirectoryPath -ServiceFilePrefix 'Exchange';
                        break;
                    } # end try
                    catch {
                        writeToLog -LogType Warning -LogString  ('Config for service config Exchange not found. Using hardcoded default.');
                        writeTolog -LogString ($_.Exception.Message) -LogType Warning;
                        $msxServicesList=[System.Collections.ArrayList]::new();
                        [void]$msxServicesList.Add([PSCustomObject]@{'DisplayName'='Microsoft Exchange*'});
                        [void]$msxServicesList.Add([PSCustomObject]@{'DisplayName'='Microsoft Filtering Management Service*'});
                        [void]$msxServicesList.Add([PSCustomObject]@{'DisplayName'='IIS Admin Service'});
                        [void]$msxServicesList.Add([PSCustomObject]@{'DisplayName'='World Wide Web Publishing Service'});                                                            
                        saveServiceInfo -ServiceCfgObject $msxServicesList -ComputerName $ComputerName -ServiceFileDirectoryPath $ServiceFileDirectoryPath -ServiceFilePrefix 'Exchange';
                    }; # end catch                    
                    break;
                }; # end Default
                {$_ -eq 'AllServices'} {
                    $msxServicesList=[System.Collections.ArrayList]::new();
                    [void]$msxServicesList.Add([PSCustomObject]@{Name='*'})
                    saveServiceInfo -ServiceCfgObject $msxServicesList -ComputerName $ComputerName -ServiceFileDirectoryPath $ServiceFileDirectoryPath -ServiceFilePrefix 'AllServices';
                    break;
                }; # end AllServices
                {$_ -eq 'SetOfServices'} {
                    saveServiceInfo -ServiceConfig $ServiceConfig -ComputerName $ComputerName -ServiceFileDirectoryPath $ServiceFileDirectoryPath -ServiceFilePrefix $ServiceConfig;
                    }; # end SetOfServices
            }; # end switch
        } # end try
        catch {
            writeToLog -LogType Error -LogString  ('Faild to save the service startup for computer ' + $ComputerName);
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch        
    }; # end process

    end
    {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
   
}; # end function Save-OPXExchangeServiceStartupTypeToFile

function Restore-OPXExchangeServiceStartupType
{
<#
.SYNOPSIS 
Restores the startup type information for a set of services from a configuration file.
	
.DESCRIPTION
The command restores the startup type of a set of services from a configuration file created with the command Save-OPXEchangeServiceStartupTypeToFile. The startup type of a service will be restored, when the optional parameter RetoreStartupType is used.
Please note, that the services will NOT be restarted. It is up to you to restart the services.

.PARAMETER ServiceFilePath
The parameter is mandatory.
The file path to the file where the configuration is stored. 

.PARAMETER CompterName
The parameter is mandatory. 
The name of the computer to which the service startup configuration is restored.

.PARAMETER RestoreStartupType
The parameter is optional.
If this parameter is used, the startup type of the services will be restored. If you don’t use the parameter, the command will only display the startup type from the configuration file, if it differs from the current startup type.

.PARAMETER ServiceConfigFilePath
The parameter is optional.
If you have created a configuration file with a customized set of services, you can use this parameter to point to your customized configuration file.

.EXAMPLE
Restore-OPXExchangeServiceStartupType -Computer <ComputerName> -ServiceFilePath <path to the file where the configuration is saved>
The command will display which services, based on the content of the configuration file, have a different startup type.
.EXAMPLE
Restore-OPXExchangeServiceStartupType -Computer <ComputerName> -ServiceFilePath <path to the file where the configuration is saved> -RestoreStartupType
The command will restore the startup type for services, based on the content of the configuration file.
#>

[cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
param([Parameter(Mandatory = $true, Position = 0)][string]$ServiceFilePath,
      [Parameter(Mandatory = $true, Position = 1)][string]$ComputerName,
      [Parameter(Mandatory = $false, Position = 2)][switch]$RestoreStartupType=$false
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin
    
    process {
        try {
            $msxSrvList = Import-Clixml -Path $ServiceFilePath;
        } # end try
        catch {
            writeToLog -LogType Error -LogString  ('Failed to load the file ' + $ServiceFilePath);
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
            return;
        }; # end catch
        if (($computerName = testIfFQDN -ComputerName $computerName) -eq $false)
        {
            return;
        }; # end if
        $allSrvStartUpTypeOk=$true;
        foreach ($service in $msxSrvList)
        {
            writeToLog -LogString ('Verifying startup type for service ' + $service.name + ' on computer ' + $computerName);
            <#$svcPList=@{
                ComputerName=$computerName;
                ErrorAction='Stop';
                Query=('select * from win32_service WHERE Name='+"'$serviceName'")
            }; # end svcPlist
            #>
            try {
                #if (($srv=Get-CimInstance @svcPList).StartMode -ne $service.StartMode)            
                if (($srv=Get-Service -Name ($service.Name) -ComputerName $ComputerName).StartType -ne $service.StartType)
                {
                    if ($RestoreStartupType.IsPresent)
                    {
                        Write-Verbose ('Setting startType for service ' + ($service.Name) + ' to ' + $service.StartType.Value);
                        writeTolog -LogString ('Setting startType for service ' + ($service.Name) + ' to ' + $service.StartType.Value);
                        try
                        {
                            
                            if ($PSCmdlet.ShouldProcess('Do you want to set the startupType for service ' +($service.name)+ ' from ' + $srv.StartType + ' to ' + ($service.StartType.value) + '.'))
                            {
                                Set-Service -Name ($service.name) -StartupType ($service.StartType.Value) -ComputerName $ComputerName -ErrorAction Stop;
                            }; # end if
                        } # end try
                        catch
                        {
                            writeToLog -LogType Warning -LogString  ('Faild to set the startType for the service ' + ($service.Name) + ' on computer ' +$ComputerName + ' to ' + ($service.StartType.Value))
                        }; #end catch
                    } # end if
                    else {
                        writeTolog -LogString ('The servcie ' + $service.Name + ' on computer ' +$ComputerName + ' has a startup type of ' + $srv.StartType + '. The reference startup type is ' + $service.StartType.Value) -LogType Info -ShowInfo;
                        $allSrvStartUpTypeOk=$false;
                    }; # end else
                }; # end if    
            } # end try
            catch {
                $msg=('Failed to configure the start mode for service ' + $service.DisplayName + ' on computer ' + $ComputerName);
                writeToLog -LogString $msg -LogType Error;
                writeToLog -LogString ($_.Exception.Message) -LogType Error;
            }; # end catch            
        }; # end forach
        if ($allSrvStartUpTypeOk -eq $false)
        {
            writeToLog -LogType Warning -LogString  ('Not all services have the correct startup type configurd. To fix the issue, run the command with the parameter RestoreStartupType')
        }; # end if
    }; # end process
    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
}; # end function Restore-OPXExchangeServiceStartupType
function New-OPXExchangeCertificateRequest
{
<#
.SYNOPSIS 
Request, install, assign services and copy a certificate to Exchange servers with CAS.

	
.DESCRIPTION
The command allows you to request, install, assign services to the certificate and copy the certificate to other Exchange servers with CAS. Dependent what parameter you use, you can perform the tasks without additional input. You only have to confirm if the default SMTP certificate should be overwritten.

.PARAMETER RequestUNCFilePath
The parameter is mandatory.
The file path of the request file. The file will be created according your configuration input.

.PARAMETER C
The parameter is optional. 
Enter the the country code for the certificate request (subject in certificate).

.PARAMETER O
The parameter is optional.
Enter the name of the organization for the certificate request (subject in certificate).

.PARAMETER OU
The parameter is optional.
Enter the name of the organization unit for the certificate request (subject in certificate).

.PARAMETER L
The parameter is optional.
Enter the name of the location for the certificate request (subject in certificate).

.PARAMETER S
The parameter is optional.
Enter the name of the state for the certificate request (subject in certificate).

.PARAMETER CertificateFreindlyName
The parameter is optional.
The friendly name for the certificate.

.PARAMETER Server
The parameter is optional.
The name of the Exchange server where you want to run this command. FQDN and NetBiOS name are supported. If you don’t use the parameter, the name of the server, on which you run the cmdlet, is used.

.PARAMETER ServiceList
The parameter is optional.
The list of services the certificate will be enabled for. Per default the certificate will be enabled for the services IMAP, POP, IIS and SMTP.

.PARAMETER OnlyDisplayConfiguration
The parameter is optional.
If the parameter is used the request will not be created, only the values used for the request will be displayed.

.PARAMETER AdditionalNamespaces
The parameter is optional.
Pre default the following namespaces will be included in the certificate: all virtual directories, internal Autodiscover uri, the Autodiscover namespace of every accepted domain. If you need additional namespaces in the certificate, you can add a list of additional namespaces.

.PARAMETER GlobalAutodiscoverNamespace
The parameter is optional.
If you don’t want an Autodiscover namespace for every accepted domain, with this parameter you can add an appropriate Autodiscover namespace. No Autodiscover namespace for an accepted domain will be added.

.PARAMETER CopyCertificateToServers
The parameter is optional.
If you want to copy the new certificate to additional server, you can add the list of servers. With TAB you can step through the list of Exchange servers with CAS installed.
You can’t use the parameter with the parameters CopyCertificateToAllCAS and CopyCertificateToMembersOfDAG.

.PARAMETER CopyCertificateToAllCAS
The parameter is optional.
If you want to copy the new certificate to all Exchange servers with CAS installed, you can use this parameter.
You can’t use the parameter with the parameters CopyCertificateToServers and CopyCertificateToMembersOfDAG.

.PARAMETER CopyCertificateToMembersOfDAG
The parameter is optional.
If you want to copy the new certificate to all members of the DAG where this server is member, you can use this parameter.
You can’t use the parameter with the parameters CopyCertificateToServers and CopyCertificateToAllCAS.

.PARAMETER RequestType
The parameter is optional.
With this parameter you can select between the request types
	RequestAndInstall (default if the parameter is not used)
	RequestOnly
	InstallOnly
With the RequestType RequestAndInstall a request file will be created. After you have signed the request the certificate will be installed on the server where the request was created. The appropriate services will be assigned. If you have selected an option to copy the certificate to other servers, the certificate will be copied to these servers.
With the RequestType RequestOnly, only the request will be created. You can re-run the command with the RequestType of InstallOnly to install the certificate.
With the RequestRype InstallOnly the certificate will be installed on the server. The appropriate services will be assigned. If you have selected an option to copy the certificate to other servers, the certificate will be copied to these servers.

.PARAMETER CertificateTemplateName
The parameter is optional.
With this parameter you can select the name of the certificate for the request. The parameter is only used if the RequestType RequestAndInstall or InstallOnly is used. If you don’t specify a template the template with the name Webserver is used.

.PARAMETER CAServerFQDN
The parameter is optional.
The parameter expects a FQDN of a certificate authority server. If the parameter is used together with the parameter CertificateAuthorityName, the certificate will be requested without user intervention. It is required that the user who runs the command has the right to request the certificate.

.PARAMETER CertificateAuthorityName
The parameter is optional.
The parameter expects the name of a certificate authority. If the parameter is used together with the parameter CAServerFQDN, the certificate will be requested without user intervention. It is required that the user who runs the command has the right to request the certificate.

.PARAMETER CertificateUNCFilePath
The parameter is optional.
The parameter an UNC file path to the certificate. If the parameter is omitted, the certificate will be stored in the directory of the request file. The default name is certnew.cer.

.PARAMETER VerifyCertificatesAfterRollout
The parameter is optional.
If the parameter is used and the certificate will be copied to other servers, the command will try to list the certificate on these servers.

.PARAMETER CollectSubjectNameEntriesFrom
The parameter is optional.
If the parameter is omitted, the virtual directories on the server where the certificate was requested will be queried for SAN entries. The parameters accept the value
	ALLCAS (all servers with CAS installed will be queried)
	DAG (all servers of the DAG, where the requesting server is member, will be queried)

.PARAMETER BinaryEncoded
The parameter is optional.
If the parameter is used the request will be encoded with DER.

.PARAMETER KeySize
The parameter is optional.
The keysize for the certificate. Default is 2048.

.PARAMETER DomainController
The parameter is optional.
You can specify a domain controller which is used for the request, install and distribution of the certificate to other servers.

.EXAMPLE
$newCertRequest=@{
    RequestUNCFilePath=<path to request file>;
    C=<country code> 
    O=<organization name>;
    OU=<organization unit name>;
    CertificateFreindlyName='Exchange 2019 Certificate';
    RequestType='RequestAndInstall';
    CopyCertificateToAllCAS=$true;
    CAServerFQDN=fqdn.of.pki.server;
    CertificateAuthorityName=<certificate-authority-name>;
    OnlyDisplayConfiguration=$true;
    VerifyCertificatesAfterRollout=$true;
    CollectSubjectNameEntriesFrom='Server';
    CertificateUNCFilePath=<\\servername\share name\subfolder name\cerfile.name>;
};
New-OPXExchangeCertificateRequest @newCertRequest
Because OnlyDisplayConfiguration is set to TRUE, the configuration for the request will be displayed. To request and install the certificate change the value for OnlyDisplayConfiguration to FALSE
#>
[CmdLetBinding(DefaultParameterSetName='Default')]
param([Parameter(Mandatory = $false, Position = 0)][string]$RequestUNCFilePath,
      [Parameter(Mandatory = $false, Position = 1)][alias("Country")][string]$C,
      [Parameter(Mandatory = $false, Position = 2)][alias("Organization")][string]$O,
      [Parameter(Mandatory = $false, Position = 3)][alias("Department")][string]$OU,
      [Parameter(Mandatory = $false, Position = 4)][alias("Location")][string]$L,
      [Parameter(Mandatory = $false, Position = 5)][alias("State")][string]$S,
      [Parameter(Mandatory = $false, Position = 6)][string]$CertificateFreindlyName='Exchange Certificate',
      [Parameter(Mandatory = $false, Position = 7)]
      [string]$Server=$__OPX_ModuleData.ConnectedToMSXServer,
      [Parameter(Mandatory = $false, Position = 8)][array]$ServicesList=@('IMAP','POP','IIS','SMTP'),
      [Parameter(Mandatory = $false, Position = 9)][switch]$OnlyDisplayConfiguration = $false,
      [Parameter(Mandatory = $false, Position = 10)][array]$AdditionalNamespaces,
      [Parameter(Mandatory = $false, Position = 11)][string]$GlobalAutodiscoverNamespace,
      [Parameter(ParameterSetName='CopyToServers',Mandatory = $false, Position = 12)][ArgumentCompleter( { 
        param ( $CommandName,
                $ParameterName,
                $WordToComplete,
                $CommandAst,
                $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($false,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*"});
      } )]
      [array]$CopyCertificateToServers,
      [Parameter(ParameterSetName='CopyToAllCAS',Mandatory = $false, Position = 12)][switch]$CopyCertificateToAllCAS=$false,
      [Parameter(ParameterSetName='CopyToDAG',Mandatory = $false, Position = 12)]
      [switch]$CopyCertificateToMembersOfDAG=$false,      
      [Parameter(Mandatory = $false, Position = 13)][ValidateSet('RequestAndInstall','RequestOnly','InstallOnly')][string]$RequestType='RequestAndInstall',
      [Parameter(Mandatory = $false, Position = 14)][string]$CertificateTemplateName='WebServer',
      [Parameter(Mandatory = $false, Position = 14)][string]$CAServerFQDN,
      [Parameter(Mandatory = $false, Position = 15)][string]$CertificateAuthorityName,
      [Parameter(Mandatory = $false, Position = 16)][string]$CertificateUNCFilePath,
      [Parameter(Mandatory = $false, Position = 17)][switch]$VerifyCertificatesAfterRollout,
      [Parameter(Mandatory = $false, Position = 18)][ValidateSet('AllCAS','DAG','Server')][string]$CollectSubjectNameEntriesFrom='AllCAS',
      [Parameter(Mandatory = $false, Position = 19)][ValidateSet(1024,2048,4096)][int]$KeySize=2048,
      [Parameter(Mandatory = $false, Position = 20)][switch]$BinaryEncoded=$false,
      [Parameter(Mandatory = $false, Position = 21)][string]$DomainController
     )

    begin {
        if (!($PSBoundParameters.ContainsKey('RequestUNCFilePath')))
        {
            switch ($RequestType)
            {
                'InstallOnly' {$RequestUNCFilePath=[system.io.path]::Combine([System.IO.Path]::GetTempPath(),'certnew.cer')};
                default {
                    $RequestUNCFilePath=Read-Host -Prompt "Please supply a value for the  parameter`nRequestUNCFilePath"
                    if ([System.String]::IsNullOrEmpty($RequestUNCFilePath))
                    {
                        throw('The parameter RequestUNCFilePath requires a value.')
                    }; # end if
                }; # end default
            }; # end switch
        }; # end if
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        if (($server=(testIfFQDN -ComputerName $server)) -ne $false)
        {
            $server=$Server.ToLower();
        }; # end if
        
        if (!([System.Linq.Enumerable]::Contains([string[]]($__OPX_ModuleData.getExchangeServerList($false,$true,$true)).ToLower(),$server)))
        {
            writeToLog -LogString ('The server ' + $server + ' is not an Exchange server with CAS.') -LogType Warning;
            return;
        }; # end if
        if ((($PSBoundParameters.ContainsKey('CAServerFQDN') -or $PSBoundParameters.ContainsKey('CertificateAuthorityName')) -and ($RequestType -eq 'RequestAndInstall')) -and -not ($PSBoundParameters.ContainsKey('CAServerFQDN') -and $PSBoundParameters.ContainsKey('CertificateAuthorityName')))
        {
            writeToLog -LogType Warning -LogString  'For an fully automated certificate rollout the parameters CAServerFQDN AND CertificateAuthorityName are required.'
            return;
        }; # end if
        
        if (Test-Path -Path $RequestUNCFilePath -PathType Container)
        {
            writeTolog -LogString ('The path ' + $RequestUNCFilePath + ' is a directory path. A path to the request file is required.') -LogType Error;
            return;
        }; # end if
        if (($PSBoundParameters.ContainsKey('CertificateUNCFilePath') -and (Test-Path -Path $CertificateUNCFilePath -PathType Container)))
        {
            writeTolog -LogString ('The path ' + $CertificateUNCFilePath + ' is a directory path. A path to the certificate (cer) file is required.') -LogType Error;
            return;
        }; # end if
        
        if (($server=testIfFQDN -ComputerName $server) -eq $false)
        {
            return;
        };
        if (([System.Uri]$RequestUNCFilePath).IsUnc -eq $false) # check if path is an UNC path
        {
            $RequestUNCFilePath=[System.IO.Path]::Combine(('\\'+$Server),($RequestUNCFilePath).Replace(':','$'));        
        }; # end if
        if (([System.Uri]$CertificateUNCFilePath).IsUnc -eq $false) # check if path is an UNC path
        {
            $CertificateUNCFilePath=[System.IO.Path]::Combine(('\\'+$Server),($CertificateUNCFilePath).Replace(':','$'));        
        }; # end if

        $requestDir=[System.IO.Path]::GetDirectoryName($RequestUNCFilePath)
        if (($RequestType -in ('RequestAndInstall','RequestOnly')) -or ($OnlyDisplayConfiguration.IsPresent -eq $true)) #$PSBoundParameters.ContainsKey('OnlyDisplayConfiguration'))
        {
            if ($RequestType -ne 'InstallOnly')
            {
                if (! (Test-Path -Path ($requestDir) -PathType Container))
                {
                    writeToLog -LogType Warning -LogString  ('The directory ' + $requestDir + ' does not exist.')
                    return;
                }; # end if
                $subjectName = ''
                foreach ($element in  $PSBoundParameters.Keys)
                {
                    if ($element -in ("C","O","OU","L","S"))
                    {
                        $subjectName += (','+$element + "=" + $PSBoundParameters.$element)
                    }; # end if
                }; # end foreach
        
                Write-Verbose 'Building subject name';
                writeTolog -LogString 'Building subject name';
                $cn = @();      
                $cn += [string](Get-OutlookAnywhere -Server $Server).ExternalHostName;
                $domainEntryList=[System.Collections.ArrayList]::new();
            
                if ($PSBoundParameters.ContainsKey('AdditionalNamespaces'))
                {
                    Write-Verbose 'Adding additional namespaces';
                    writeTolog -LogString  'Adding additional namespaces';                
                    [void]$domainEntryList.AddRange(@([System.Linq.Enumerable]::Distinct([string[]]$AdditionalNamespaces)));
                }; #end if
            
                Write-Verbose 'Building list of SAN entries';
                writeTolog -LogString 'Building list of SAN entries';
                switch ($CollectSubjectNameEntriesFrom)
                {
                    'AllCAS'    {
                        try
                        {
                            $serverList=@($__OPX_ModuleData.getExchangeServerList($false,$true,$true));
                        } # end try
                        catch
                        {
                            writeToLog -LogType Error -LogString  ('Failed to connect to client access service on server ' + $Server);
                            writeToLog -LogString ($_.Exception.Message) -LogType Error;
                            return;
                        }; # end catch
                        break;
                    }; # end AllCAS
                    'DAG'    {
                        try
                        {
                            $memberOfDAG=[string](Get-MailboxServer -Identity $server).DatabaseAvailabilityGroup;
                            # extract servers with CAS
                            $serverList=@([System.Linq.Enumerable]::Intersect([string[]]$__OPX_ModuleData.getDagList($true,$memberOfDAG),[string[]]$__OPX_ModuleData.getExchangeServerList($false,$true,$true)));
                        } # end try
                        catch
                        {
                            writeToLog -LogType Error -LogString  ('Failed to connect to client access service on server ' + $Server);
                            writeToLog -LogString ($_.Exception.Message) -LogType Error;
                            return;
                        }; # end catch
                        break;
                    }; # end DAG
                    'Server'    {
                        try
                        {
                            # extract severs with CAS
                            $serverList=@([System.Linq.Enumerable]::Intersect([string[]]$Server.ToLower(),[string[]]$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower()))
                        } # end try
                        catch
                        {
                            writeToLog -LogType Error -LogString  ('Failed to connect to client access service on server ' + $Server);
                            writeToLog -LogString ($_.Exception.Message) -LogType Error;
                            return;
                        }; # end catch 
                        break;               
                    }; # end server
                }; # end switch
                $urlParamList=@{
                    SrvList=$serverList;
                    ADPropertiesOnly=$true
                }; # end urlParamList
                if ($PSBoundParameters.ContainsKey('GlobalAutodiscoverNamespace'))
                {
                    [void]$domainEntryList.Add($GlobalAutodiscoverNamespace);
                    $urlParamList.Add('SkipAcceptedDomains',$true);
                }; # end if
                if ($domainEntryList.Count -gt 0)
                {
                    $urlParamList.Add('DomainEntryList',$domainEntryList);   
                }; # end if
                $domainEntryList = getURLList @urlParamList;

                If ($domainEntryList -eq $false)
                {
                    writeToLog -LogType Error -LogString  "Failed to collect the data for the certificate request.";
                    return;
                }; # end if
            }; # end if InstallOnly
            if ($OnlyDisplayConfiguration)
            {
                Write-Output 'Displaying certificate configuration:';
                if ($RequestType -ne 'InstallOnly')
                {
                    Write-Output ('Request created on server: ' + $server);
                    Write-Output ('Template name: ' + $CertificateTemplateName);
                    Write-Output ('Key size: ' + $KeySize.ToString());
                    Write-Output ('Subject name: CN=' + $cn[0] + $subjectName);
                    Write-Output 'List of subject alternative name entries:';
                    @([System.Linq.Enumerable]::Distinct([string[]]$domainEntryList));
                    if ($PSBoundParameters.ContainsKey('CertificateAuthorityName'))
                    {
                        Write-Output ('Send request to CA: ' + $CertificateAuthorityName)
                    }; # end if
                }; # end if
                switch ($PsCmdlet.ParameterSetName)
                {
                    'CopyToAllCAS'  {
                        Write-Output 'Certificate will be copied to all CAS';
                        break;
                    }; # end copyToAllCAS
                    'CopyToDAG'     {
                        Write-Output 'Certificate will be copied to all DAG mebers';
                        break;
                    }; # end copyToAllCAS
                    'CopyToServers' {
                        Write-Output 'Certificate will be copied to the following servers:';
                        $CopyCertificateToServers;
                    }; # end copyToAllCAS                    
                }; # end switch
                return;
            }; # end if
            
            Write-Verbose 'Generate certificate request'; 
            writeTolog -LogString  'Generate certificate request';                                
        }; # end if requestAndInstall or requestOnly

        $parmList=@{
            ServerName=$Server;
            ServicesList=$ServicesList;
            RequestUNCFilePath=$RequestUNCFilePath;
            FriendlyName=$CertificateFreindlyName;        
            RequestType=$RequestType;
            KeySize=$KeySize;
        }; # end paramList
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $parmList.Add('DomainController',$DomainController);
        }; # end if
        if (($RequestType -in ('RequestAndInstall','RequestOnly'))) # add values for cert install and rollout
        {       
            $parmList.Add('SubjectName',('CN=' + $cn[0] + $subjectName));
            $parmList.Add('DomainList',$domainEntryList);            
            if (($RequestType -eq 'RequestAndInstall') -and ($PSBoundParameters.ContainsKey('CAServerFQDN')-and $PSBoundParameters.ContainsKey('CertificateAuthorityName')))
            {
                if (($CAServerFQDN=testIfFQDN -ComputerName $CAServerFQDN) -ne $false) # is the PKI server resoveable in DNS
                {
                    $parmList.Add('CAServerFQDN',$CAServerFQDN);
                    $parmList.Add('CertificateAuthorityName',$CertificateAuthorityName);                
                    $parmList.Add('CertificateTemplateName',$CertificateTemplateName);                               
                }; # end if
            }; # end if
        }; # end if

        $copyCertToOtherServers=(($RequestType -in ('RequestAndInstall','InstallOnly')) -and (($PsCmdlet.ParameterSetName -eq 'CopyToServers' -or $PsCmdlet.ParameterSetName -eq 'CopyToDAG' -or $PsCmdlet.ParameterSetName -eq 'CopyToAllCAS')))
        if ($copyCertToOtherServers -eq $false)
        {
            $parmList.Add('EnableCertificate',$true); 
        };
        if ($PSBoundParameters.ContainsKey('CertificateUNCFilePath'))
        {
            $parmList.Add('CertificateFilePath',$CertificateUNCFilePath);
        }; # end if 
        $certTP=generateCertRequest @parmList;

        if ($copyCertToOtherServers -and ($certTP -ne $false))     
        {        
            switch ($PsCmdlet.ParameterSetName)
            {
                {$_ -eq 'CopyToDAG'}      {
                    try {
                        $dagName=[string](Get-MailboxServer -Identity $server).DatabaseAvailabilityGroup;
                        $CopyCertificateToServers=@([System.Linq.Enumerable]::Intersect([string[]]$__OPX_ModuleData.getDagList($true,$DAGName).ToLower(),[string[]]$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower()));                    
                    } # end try
                    catch {
                        writeToLog -LogType Error -LogString  ('Faild to enumerate the DAG member servers.');
                        writeToLog -LogString ($_.Exception.Message) -LogType Error;
                        return;
                    }; # end catch
                    break;
                }; # end DAG
                {$_ -eq 'CopyToAllCAS'}   {
                    try
                    {
                        Write-Verbose 'Building list of Exchange servers with Client Access service';
                        writeTolog -LogString 'Building list of Exchange servers with Client Access service';
                        $CopyCertificateToServers=$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower();
                    } # end try
                    catch
                    {
                        writeToLog -LogType Error -LogString  ('Failed to list Exchange servers with Client Access service.');
                        writeToLog -LogString ($_.Exception.Message) -LogType Error;    
                    return;
                    }; # end catch
                    break;  
                }; # end AllCAS
            }; # end switch

            $targetSrvList=[System.Collections.ArrayList]::new();
            foreach($srv in $CopyCertificateToServers)
            {                
                if (($srvFQDN=(testIfFQDN -ComputerName $srv)) -notin ($false,$server))               
                {
                    [void]$targetSrvList.Add($srvFQDN);
                }; # end if 
            }; # end foreach            
            $parmList=@{
                TargetServerList=$targetSrvList;
                SourceServer=$Server;
                ServicesList=$ServicesList;
                CertificateTmpFilePath=(Join-Path -Path $requestDir -ChildPath (([guid]::NewGuid().guid)+'.pfx'));
                Thumbprint=$certTP;
                DeploymentType='CopyOnly';
                ReturnStatus=$true;
            }; # end ParamList
            if ($PSBoundParameters.ContainsKey('DomainController'))
            {
                $parmList.Add('DomainController',$DomainController);
            }; # end if
            writeToLog -LogString 'Starting copying certificate.'
            if (Copy-OPXExchangeCertificateToServers @parmList) # test if cert was copied to all servers
            {
                [void]$targetSrvList.Add($server);                
                $parmList.TargetServerList=$targetSrvList;
                $parmList.Remove('CertificateTmpFilePath');
                $parmList.DeploymentType='EnableOnly';
                $parmList.ReturnStatus=$false;
                $parmList.Add('EnableOnSourceServer',$true);
                writeTolog -LogString 'Starting enabeling the certificates';
                Copy-OPXExchangeCertificateToServers @parmList;
            }; # end
            if ($VerifyCertificatesAfterRollout.IsPresent)
            {
                $msg=('Verifying certificate with thumbprint ' + ($parmList.Thumbprint) + ' on server ' + ($parmList.TargetServerList -join ' '));
                Write-Verbose $msg;
                writeTolog -LogString $msg;
                Get-OPXExchangeCertificate -ServerList ($parmList.TargetServerList) -Thumbprint ($parmList.Thumbprint) -Format List;
            }; # end if
        }; # end if
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
}; # end function New-OPXExchangeCertificateRequest

function New-OPXExchangeAuthCertificate
{
<#
.SYNOPSIS 
Create a new Exchange Auth Certificate. 
	
.DESCRIPTION
With the command a new Exchange Auth certificate can be created. The thumbprint of will be written to the Exchange auth configuration. The command can initiate the rollout of the certificate to all Exchange servers with CAS installed.

.PARAMETER Server
The parameter is optional.
The name of the where the certificate should be created or the IIS should be reset.
	
.PARAMETER DomainName
The parameter is optional create certificate).
Use this parameter if you only want to get the database copy configuration. No configuration for database copy or DAG membership will be removed.

.PARAMETER EnableCertificateInHours
The parameter is optional.
The date when the new certificate should be use (n hours from now). Default is 48 hours.

.PARAMETER PublishingAndClearCert
The Parameter is optional.
The parameter offers the options
	Publish, the certificate will be created and published (Set-AuthConfig)
	PublishAndClearPrevious, the certificate will be created, published and the previous certificate will be cleared(Set-AuthConfig)
	None, only create the certificate

.PARAMETER ResetIISOnAllCAS
The parameter is optional.
The IIS on all Exchange servers with CAS will be reset. This parameter cannot be used with any other parameter. When this parameter is used, the new auth certificate should be created and published.

.PARAMETER OnlyResetIISOCAS
The parameter is optional.
The IIS on the Exchange server, where the certificate was created, will be reset.

.PARAMETER DomainController
The parameter is optional.
The name of a domain controller which is used for the Exchange PowerShell cmdlets.

.PARAMETER KeySize
The parameter is optional.
The key size for the certificate. Default is 2048

.EXAMPLE
New-OPXExchangeAuthCertificate -DomainName <domain.name> -PublishingAndClearCert Publish
#>
[cmdletbinding(DefaultParametersetName='RolloutCert')]
param([Parameter(ParametersetName='RolloutCert')][Parameter(ParametersetName='ResetOnly')]
        [Parameter(Mandatory = $false, Position = 0)]
        [ArgumentCompleter( {             
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($false,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*"});
        } )]
        [string]$Server=$__OPX_ModuleData.ConnectedToMSXServer,     
        [Parameter(ParametersetName='RolloutCert',Mandatory = $true, Position = 1)][string]$DomainName,
        [Parameter(ParametersetName='RolloutCert',Mandatory = $false, Position = 2)][int]$EnableCertificateInHours=48,
        [Parameter(ParametersetName='RolloutCert',Mandatory = $false, Position = 3)][ValidateSet('None','Publish','PublishAndClearPrevious')][string]$PublishingAndClearCert='None',
        [Parameter(ParametersetName='ResetOnlyOnAllCAS',Mandatory = $false, Position = 4)][switch]$ResetIISOnAllCAS=$false,
        [Parameter(ParametersetName='ResetOnly',Mandatory = $false, Position = 4)][switch]$OnlyResetIISOnCAS,
        [Parameter(ParametersetName='RolloutCert',Mandatory = $false, Position = 5)][string]$DomainController,
        [Parameter(ParametersetName='RolloutCert',Mandatory = $false, Position = 6)][ValidateSet(1024,2048,4096)][int]$KeySize=2048
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        try {
            $server=testIfFQDN -ComputerName $server;
            if (!([System.Linq.Enumerable]::Contains([string[]]($__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower()),$server.toLower())))
            {
                writeToLog -LogString ('The server ' + $server + ' is not an Exchange server with CAS.') -LogType Warning;
                return;
            }; # end if
            if ($PSCmdlet.ParameterSetName -eq 'RolloutCert')
            {
                $paramList=@{
                    Server=$Server;
                    KeySize=$KeySize;
                    PrivateKeyExportable=$true;
                    SubjectName='cn= Microsoft Exchange Server Auth Certificate';
                    DomainName=$DomainName;
                    FriendlyName='Microsoft Exchange Server Auth Certificate';
                    Services='SMTP';
                    ErrorAction='Stop';
                    BinaryEncoded=($BinaryEncoded.IsPresent);
                }; # end paramList
                if ($PSBoundParameters.ContainsKey('DomainController'))
                {
                    $parmList.Add('DomainController',$DomainController);
                }; # end if
                Write-Verbose 'Creating certificate';
                writeTolog -LogString 'Creating certificate';
                $cert=New-ExchangeCertificate @paramList;
                Write-Verbose ('Updateing AuthConfig with the new certificate thumbprint ' + $cert.Thumbprint); 
                writeTolog -LogString ('Updateing AuthConfig with the new certificate thumbprint ' + $cert.Thumbprint); 
                Set-AuthConfig -NewCertificateThumbprint $cert.Thumbprint -NewCertificateEffectiveDate ((Get-Date).AddHours($EnableCertificateInHours));                           
                if ($PublishingAndClearCert -in ('Publish','PublishAndClearPrevious'))
                {
                    Write-Verbose 'Publishing certificate';     
                    writeTolog -LogString 'Publishing certificate';
                    Set-AuthConfig -PublishCertificate;
                }; # end if
                if ($PublishingAndClearCert -eq 'PublishAndClearPrevious')
                {
                    Write-Verbose 'Clearing certificate';
                    writeTolog -LogString 'Clearing certificate';
                    Set-AuthConfig -ClearPreviousCertificate;
                }; # end if
            } ; # end if
            if (($ResetIISOnAllCAS.IsPresent) -or ($PsCmdlet.ParameterSetName -eq 'ResetOnly'))
            {
                
                if (($PsCmdlet.ParameterSetName -eq 'ResetOnly') -and ($PSBoundParameters.ContainsKey('Server')))
                {
                    $casList=@(@{FQDN=$server});
                } # end if
                else {
                    Write-Verbose 'Collecting list of client access services';
                    writeTolog -LogString 'Collecting list of client access services';
                    $casList=$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower();
                }; # end else
                
                foreach ($cas in $casList)
                {
                    $msg=('Restarting OWA and ECP app pool on server ' + ($cas));
                    Write-Verbose $msg;
                    writeTolog -LogString $msg; 
                    try {
                        Invoke-Command -ComputerName ($cas) -ScriptBlock {Restart-WebAppPool MSExchangeOWAAppPool;Restart-WebAppPool MSExchangeECPAppPool};   
                    } # end try
                    catch {
                        writeToLog -LogType Error -LogString  ('Failed restarting OWA and ECP app pool on server ' + ($cas));
                        writeToLog -LogString ($_.Exception.Message) -LogType Error;    
                    }; # end catch                
                } # end foreach
            } # end if
            else {
                Write-Verbose ('Resetting OWA and ECP app pool on server ' + $Server);
                writeTolog -LogString ('Resetting OWA and ECP app pool on server ' + $Server);
                Restart-WebAppPool MSExchangeOWAAppPool;
                Restart-WebAppPool MSExchangeECPAppPool;
                writeToLog -LogType Warning -LogString  'Reset the web app pools MSExchangeOWAAppPool and MSExchangeECPAppPool on the client access servers';
            }; # end else
            
        } # end try
        catch {
            writeToLog -LogType Error -LogString  ('Failed to configure the authentication certificate');
            writeToLog -LogString ($_.Exception.Message) -LogType Error;    
        }; #end catch
    }; # process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
}; #end function New-OPXExchangeAuthCertificate

function Test-OPXExchangeAuthCertificateRollout
{
<#
.SYNOPSIS 
Checks if the certificate configured in the auth config is present on all Exchange servers with CAS installed. 
	
.DESCRIPTION
Checks if the certificate configured in the auth config is present on all Exchange servers with CAS installed.

.PARAMETER DomainController
The parameter is optional.
The name of a domain controller which is used for the Exchange PowerShell cmdlets.
	

.EXAMPLE
Test-OPXExchangeAuthCertificateRollout
#>

[CmdLetBinding()]
param([Parameter(Mandatory = $false, Position = 0)][string]$DomainController
     )
    
    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
    try {        
        $DCParam=@{
            ErrorAction='Stop';
        }; # end DCParam
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $DCParam.Add('DomainController',$DomainController);
        }; # end if
        
        $serverCertList=[system.Collections.arraylist]::New(); # @();
        $authCfgNotValid=$false;
        Write-Verbose 'Building list of Exchange servers with Client Access service';
        writeTolog -LogString 'Building list of Exchange servers with Client Access service';
        #$casServiceList = ((Get-ClientAccessService @DCParam).name);
        $casServiceList=$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower();
        Write-Verbose 'Reading Auth certificate thumbprint from configuration';
        writeTolog -LogString 'Reading Auth certificate thumbprint from configuration';
        $authCertThumbprint=(Get-AuthConfig @dcparam).CurrentCertificateThumbprint;
        Write-Verbose ('Auth certificate thubprint ' + $authCertThumbprint + ' found.');
        writeTolog -LogString ('Auth certificate thubprint ' + $authCertThumbprint + ' found.');
        
        foreach ($casServer in $casServiceList)
        {
            try {
                Write-Verbose ('Searching for certificate with thumbprint ' + $authCertThumbprint + ' on server ' + $casServer);
                writeTolog -LogString ('Searching for certificate with thumbprint ' + $authCertThumbprint + ' on server ' + $casServer);
                [void](Get-ExchangeCertificate -Thumbprint $authCertThumbprint -Server $casServer @DCParam);
                writeTolog -LogString ('Certificate with thumbprint ' + $authCertThumbprint + ' found.');
            } # end try
            catch {
                writeToLog -LogType Warning -LogString  ('Auth certificate with thumbprint ' + $authCertThumbprint + ' not found on server ' + $casServer);
                $authCfgNotValid=$true;
                try {
                    $authCert=(@(Get-ExchangeCertificate -Server $casServer @DCParam).Where({$_.Subject -eq 'CN=Microsoft Exchange Server Auth Certificate'}));
                    if (! ($null -eq $authCert))
                    {
                        $serverCertList.Add($authCert.Thumbprint);
                    }; #end if
                } # end try
                catch {
                    writeToLog -LogType Error -LogString  ('Failed to get certificate from server ' + $casServer);
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;
                }; # end catch                
            }; # end catch
        }; # end if
    } # end try
    catch {
        writeToLog -LogType Error -LogString  ('Failed to prepare authentication certificate rollout verification.');
        writeToLog -LogString ($_.Exception.Message) -LogType Error;
    }; # end catch
    
    if (($serverCertList.count -gt 0) -and ($authCfgNotValid -eq $true))
    {        
        [array]$tbList=@([System.Linq.Enumerable]::Distinct([string[]]$serverCertList));
        if ($tbList.Count -eq 1)
        {
            writeToLog -LogType Warning -LogString  ('The Exchange servers in your organization are configurte with the auth certificate: ' +$tbList[0]);
            writeToLog -LogType Warning -LogString  ('In the auth configuration a different thumbprint is used. ')                 
        }; # end if
    }; # end if
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
}; # end function Test-OPXExchangeAuthCertificateRollout

function Copy-OPXExchangeCertificateToServers
{
<#
.SYNOPSIS 
Copies an Exchange certificate to a server and enables services.
	
.DESCRIPTION
Copies an Exchange certificate to one or more servers and enables the certificate for a set of services.

.PARAMETER TargetServerList
The parameter is mandatory if the certificate should be copied to a list of servers.
The parameter cannot be used with the parameters DAGName and AllCAS.
List of target servers for the certificate.
	
.PARAMETER DAGName
The parameter is mandatory if the certificate should be copied to DAG members.
The parameter cannot be used with the parameters TargetServerList and AllCAS.
The certificate will be copied to all members of the DAG (all DAG members should have CAS installed).

.PARAMETER AllCAS
The parameter is mandatory if the certificate should be copied to all CAS.
The parameter cannot be used with the parameters DAGName and TargetServerList.
The certificate will be copied to all Exchange servers with CAS installed.

.PARAMETER SourcServer
The Parameter is optional.
The name of the Exchange server with the certificate. If the parameter is omitted the name of the connected exchange server will be used.

.PARAMETER ServicesList
The Parameter is optional.
The list of services for which the certificate should be enabled. Default services are IMAP, POP, IIS and SMTP.

.PARAMETER CertificateTmpFilePath
The Parameter is optional.
If the parameter DeploymentType is not EnableOnly, the parameter is mandatory. The path to the file the certificate is stored (temporary) for deployment to servers.

.PARAMETER DeploymentType
The parameter is optional.
The type of the deployment. The certificate can be copied, enabled or both.

.PARAMETER Thumbprint
The parameter is mandatory.
The thumbprint of the certificate.

.PARAMETER EnableOnSourceServer
The parameter is optional.
If the switch is used, the certificate will bei enabled for the services in the parameter ServicesList.

.PARAMETER ReturnStatus
The parameter is optional and for module internal use.
Returns if the certificate was successfully depolyed.

.PARAMETER DomainController
The parameter is optional.
The name of a domain controller which is used for the Exchange PowerShell cmdlets.

.EXAMPLE
Copy-OPXExchangeCertificate -Thumbprint <certificate thumbprint> -AllCAS
#>
[cmdletbinding(DefaultParametersetName='Default')]
param([Parameter(ParametersetName='ServerList',Mandatory = $true, Position = 0)]
      [ArgumentCompleter( {
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($false,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*"});
      } )]
      [array]$TargetServerList,
      [Parameter(ParametersetName='DAG',Mandatory = $true, Position = 0)]
      [ArgumentCompleter( { 
            param ( $CommandName,
                $ParameterName,
                $WordToComplete,
                $CommandAst,
                $FakeBoundParameters )  
            $dagList=($__OPX_ModuleData.getDagList($false,$null));
            $dagList.Where({ $_ -like "$wordToComplete*"}) ; 
      } )]
      [string]$DAGName,
      [Parameter(ParametersetName='AllCAS',Mandatory = $true, Position = 0)][switch]$AllCAS,
      [Parameter(Mandatory = $false, Position = 1)][ArgumentCompleter( { 
        param ( $CommandName,
                $ParameterName,
                $WordToComplete,
                $CommandAst,
                $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($false,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*"});
      } )]
      [string]$SourceServer=$__OPX_ModuleData.ConnectedToMSXServer,
      [Parameter(Mandatory = $false, Position = 2)][array]$ServicesList=@('IMAP','POP','IIS','SMTP'),
      [Parameter(Mandatory = $false, Position = 3)][string]$CertificateTmpFilePath,
      [Parameter(Mandatory = $false, Position = 4)][string]$DomainController,
      [Parameter(Mandatory = $true, Position = 5)][string]$Thumbprint,
      [Parameter(Mandatory = $false, Position = 6)][CertDeploymentType]$DeploymentType='CopyOnly',
      [Parameter(Mandatory = $false, Position = 7)][switch]$ReturnStatus,
      [Parameter(Mandatory = $false, Position = 8)][switch]$EnableOnSourceServer
     )

    begin {
        if (($DeploymentType.value__ -lt 2) -and (!($PSBoundParameters.ContainsKey('CertificateTmpFilePath'))))
        {
            throw [System.ArgumentException]'The parameter CertificateTmpFilePath is requiered.' 
        }; #end if
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());               
    }; # end begin
    
    process {        
        if ($PSBoundParameters.ContainsKey('CertificateTmpFilePath') -and (Test-Path -Path $CertificateTmpFilePath -PathType Container))
        {
            writeToLog -LogString ('The parameter CertificateTmpFilePath must not be a directory path.') -LogType Error;
            return;
        }; # end if
        $srvTmp=(testIfFQDN -ComputerName $SourceServer);
        if ($srvTmp -eq $false)
        {
            writeToLog -LogString ('Failed to resolve the server ' + $SourceServer);
            return;
        } # end if
        else {
            $SourceServer=$srvTmp.ToLower();
        }; # end else
        if (!([System.Linq.Enumerable]::Contains([string[]]($__OPX_ModuleData.getExchangeServerList($false,$true,$true)).ToLower(),$SourceServer)))
        {
            writeToLog -LogString ('The server ' + $server + ' is not an Exchange server with CAS.') -LogType Warning;
            return;
        }; # end if
        switch ($PSCmdlet.ParameterSetName)
        {
            'ServerList'    {               
                break;
            }; # end ServerList
            'DAG'           {
                try {
                    $dagMembers=$__OPX_ModuleData.getDagList($true,$DAGName);
                    if ($dagMembers.count -eq 0)
                    {
                        writeToLog -LogString ('The DAG ' + $dagName + ' has no member server.') -LogType Warning;
                        return;
                    };
                    $TargetServerList=@([System.Linq.Enumerable]::Intersect([string[]]$dagMembers.ToLower(),[string[]]$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower()));
                    break;
                } # end try
                catch {
                    writeToLog -LogType Error -LogString  ('Faild to enumerate the member servers of the DAG ' + $DAGName);
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;
                    return;
                }; # end catch                
            }; # end DAG
            'AllCAS'        {
                $TargetServerList=$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower(); 
            }; # end AllCAS
        }; # end switch
        
        $status=$true; # set default for status
        ### copy
        if ($DeploymentType.value__ -lt 2) 
        {
            try   # copy certificate from source server
            {
                $TargetServerList=($TargetServerList -ne $SourceServer);
                # create random password and store it in a secure string
                $secPwd = ConvertTo-SecureString  (([char[]]([char]33..[char]95) + ([char[]]([char]97..[char]126)) + 0..9 | Sort-Object {Get-Random})[0..28] -join '') -AsPlainText -Force
                Write-Verbose ('Copy certificate with thumbprint ' + $Thumbprint + ' from server ' + $SourceServer);
                writeTolog -LogString ('Copy certificate with thumbprint ' + $Thumbprint + ' from server ' + $SourceServer);
                Write-Verbose ('Temporaly save certificate to file ' + $CertificateTmpFilePath);
                writeTolog -LogString ('Temporaly save certificate to file ' + $CertificateTmpFilePath);
                Set-Content -Path $CertificateTmpFilePath -Encoding Byte -Value (Export-ExchangeCertificate -Thumbprint $Thumbprint -Password $secPwd -Server $SourceServer -BinaryEncoded:$True -ErrorAction Stop).FileData -ErrorAction Stop;
            } # end try
            catch
            {
                writeTolog -LogString ('Failed to export the Exchange certificate to ' + $CertificateTmpFilePath);
                writeToLog -LogString ($_.Exception.Message) -LogType Error;
                return;
            } # end catch
        }; # end if

        $dcPList=@{};
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $dcPList.Add('DomainController',$DomainController);
        }; # end if
        if ($EnableOnSourceServer)
        {
            $TargetServerList+=$SourceServer;
        }; # end if
        $TargetServerList=[System.Linq.Enumerable]::Distinct([string[]]$TargetServerList);
        # filter Exchange servers with CAS
        $TargetServerList=@([System.Linq.Enumerable]::Intersect([string[]]$TargetServerList.ToLower(),[string[]]$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower()))
        if ($TargetServerList.count -gt 0)
        {
            foreach ($server in $TargetServerList) #  import the cert (all CAS)
            {
                Write-Verbose ('Processing server ' + $server);
                writeTolog -LogString  ('Processing server ' + $server);
                try
                {
                    if ($DeploymentType.value__ -lt 2)
                    {
                    Write-Verbose ('Importing certificate with thumbprint ' + $Thumbprint + '.');
                    writeTolog -LogString ('Importing certificate with thumbprint ' + $Thumbprint + '.');
                    $paramList=@{
                        FileData=([Byte[]](Get-Content -Path $CertificateTmpFilePath -Encoding byte -ReadCount 0));
                        Server=$server;
                        Password=$secPwd;
                        ErrorAction='Stop';
                    }; # end ParamList
                    Import-ExchangeCertificate @paramList @dcPList;
                    }; # end if

                    if ($DeploymentType.value__ -gt 0)
                    {
                        Write-Verbose ('Enabeling certificate for services: ' + $($ServicesList -join ','));
                        writeTolog -LogString ('Enabeling certificate for services: ' + $($ServicesList -join ','));
                        $paramList=@{
                            Thumbprint=$Thumbprint;
                            Server=$server;
                            Services=$ServicesList;
                            ErrorAction='Stop';
                        }; # end paramList                    
                        Enable-ExchangeCertificate @paramList @dcPList;
                    }; # end if
                } # end try
                catch
                {
                    writeToLog  -LogString ('Failed to enable the certificate on server ' + $server);
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;  
                    $status=$false;          
                } # end catch
            } # end foreach
        } # end if
        else {
          writeTolog -LogString 'No Exchange server with CAS found.' -logType Warning;  
        };
    }; # end process

    end {
        if ($DeploymentType.value__ -lt 2)
        {
            if (Test-Path $CertificateTmpFilePath -PathType Leaf)
            {
                Remove-Item -Path $CertificateTmpFilePath -Force; # delete temp certificate
            }; # end if
        }; # end if
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
        if ($ReturnStatus)
        {
            return $status;
        }; # end if        
    }; #end END
} # end function Copy-OPXExchangeCertificateToServers

function Remove-OPXExchangeCertificate
{
<#
.SYNOPSIS 
Removes an Exchange certificate from a server.
	
.DESCRIPTION
Removes an Exchange certificate from a single server or a list of servers. The command can remove an outdated Exchange certificate form a number of Exchange servers with CAS installed.

.PARAMETER ServerList
The parameter is optional.
The parameter cannot be used with the parameters RemoveFromAllCAS and DAGName.
Default for the parameter is the connected exchange server. This is true if neither the parameter RemoveFromAllCAS nor the parameter DAGName is used.
	
.PARAMETER RemoveFromAllCAS
The parameter is optional.
The parameter cannot be used with the parameters ServerList and DAGName.
The certificate will be removed from all Exchange servers with CAS installed.

.PARAMETER DAGName
The parameter is optional.
The parameter cannot be used with the parameters ServerList and RemoveFromAllCAS.
The certificate will be removed from all DAG member servers with CAS installed.

.PARAMETER Thumbprint
The parameter is mandatory.
The thumbprint of the certificate.

.PARAMETER DomainController
The parameter is optional.
The name of a domain controller which is used for the Exchange PowerShell cmdlets.

.EXAMPLE
Remove-OPXExchangeCertificate -Thumbprint <certificate thumbprint> -AllCAS
#>

[cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High',DefaultParametersetName = 'ServerList')]
param([Parameter(ParametersetName='ServerList',Mandatory = $false, Position = 0)]
      [ArgumentCompleter( { 
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($false,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*"});
      } )]
      [array]$ServerList=$__OPX_ModuleData.ConnectedToMSXServer,
      [Parameter(ParametersetName='AllCAS',Mandatory = $false, Position = 0)][switch]$RemoveFromAllCAS=$false,
      [Parameter(ParametersetName='DAG',Mandatory = $true, Position = 0)]
      [ArgumentCompleter( {      
            param ( $CommandName,
                $ParameterName,
                $WordToComplete,
                $CommandAst,
                $FakeBoundParameters )  
            $dagList=($__OPX_ModuleData.getDagList($false,$null));
            $dagList.Where({ $_ -like "$wordToComplete*"}) ; 
      } )]
      [string]$DAGName,
      [Parameter(Mandatory = $true, ValueFromPipeline=$True, Position = 1)][string]$Thumbprint,
      [Parameter(Mandatory = $false, Position = 2)][string]$DomainController
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
        $DCParam=@{
            ErrorAction='Stop';
        }; # end DCParam
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $DCParam.Add('DomainController',$DomainController);
        }; # end if
        switch ($PsCmdlet.ParameterSetName)
        {
            {$_ -eq 'DAG'}      {
                try {
                    Write-Verbose ('Building list of Exchange servers in DAG ' + $DAGName) ; 
                    writeTolog -LogString ('Building list of Exchange servers in DAG ' + $DAGName) ; 
                    $dagMembers=$__OPX_ModuleData.getDagList($true,$DAGName);
                    if ($dagMembers.count -eq 0)
                    {
                        writeToLog -LogString ('The DAG ' + $dagName + ' has no member server.') -LogType Warning;
                        return;
                    };
                    $ServerList=@([System.Linq.Enumerable]::Intersect([string[]]$dagMembers.ToLower(),[string[]]$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower()));
                    break;
                } # end try
                catch {
                    writeToLog -LogType Error -LogString  ('Faild to enumerate the member servers of the DAG ' + $DAGName);
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;
                    return;
                }; # end catch
            }; # end DAG
            {$_ -eq 'AllCAS'}   {                
                try
                {
                    Write-Verbose 'Building list of Exchange servers with Client Access service';
                    writeTolog -LogString 'Building list of Exchange servers with Client Access service';
                    $serverList=$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower();
                } # end try
                catch
                {
                    writeToLog -LogType Error -LogString  ('Failed to list Exchange servers witch Client Access service.');
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;    
                    return;
                }; # end catch
            }; # end AllCAS
        }; # end switch   

        $serverList=@([System.Linq.Enumerable]::Distinct([string[]]$serverList));
    }; # end begin

    process {                                
        foreach ($server in $ServerList)
        {
            try
            {
                Write-Verbose ('Removeing certificate with thumbprint ' + $Thumbprint + ' from server ' + $Server + '.');
                writeTolog -LogString ('Removeing certificate with thumbprint ' + $Thumbprint + ' from server ' + $Server + '.');
                if ($PSCmdlet.ShouldProcess('Do you want to remove the certificate with the thumbprint ' +$Thumbprint+ ' from server ' + $Server + '.'))
                {
                    Remove-ExchangeCertificate -Server $server -Thumbprint $Thumbprint @DCParam -Confirm:$false;
                }; # end if
            } # end try
            catch
            {
                writeToLog -LogType Error -LogString  ('Failed to remove the certificate with the thumbprint ' + $Thumbprint + ' from server ' + $server + '.');
                writeToLog -LogString ($_.Exception.Message) -LogType Error;
            }; # end catch
        }; # end foreach
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
}; # end function Remove-OPXExchangeCertificate

function Get-OPXVirtualDirectories
{
<#
.SYNOPSIS 
Lists the virtual directories.
	
.DESCRIPTION
Lists the virtual directories for
	a given server
	a list of servers
	for a given service
	for all services
The command lists the attributes which are configured with the command Set-OPXVirtualDirectories. With the use of a configuration file additional attributes can be listed.
If you only want verify that the configuration of the virtual directories is correct, use the switch OnlyVerifyURLs (OAB url and Hafnium). If there is something wrong warnings will be displayed (not configuration data).

.PARAMETER ServiceList
The parameter is optional.
You can specify a list of services. The parameter is omitted all services are queried.

.PARAMETER ServerList
The parameter is optional.
The parameter cannot be used with the parameters AllCAS and DAGName.
Default for the parameter is the connected exchange server. This is true if neither the parameter AllCAS nor the parameter DAGName is used.
	
.PARAMETER AllCAS
The parameter is optional.
The parameter cannot be used with the parameters ServerList and DAGName.
The virtual directories from all Exchange servers with CAS installed will be displayed.

.PARAMETER DAGName
The parameter is optional.
The parameter cannot be used with the parameters ServerList and RemoveFromAllCAS.
The virtual directories from all DAG member servers with CAS installed will be displayed.

.PARAMETER ADPropertiesOnly
The parameter is optional.
Only AD properties will be displayed.

.PARAMETER OnlyVerifyUrls
The parameter is optional.
The configuration will be verified, no data will be displayed (except warnings if there are any problems).

.PARAMETER IncludeAttributesFromConfigFile
The parameter is optional.
If you provide an appropriate configuration file (xml), additional data will be displayed. More can be found in the help for the *-OPXExchangeVirtualDirectoryConfigurationTemplate cmdlets.

.PARAMETER DomainController
The parameter is optional.
The name of a domain controller which is used for the Exchange PowerShell cmdlets.

.EXAMPLE
Get-OPXExchangeVirtualDirectories -AllCAS -ADPropertiesOnly
Lists the configuration of all virtual directories on all Exchange servers with CAS installed.

.EXAMPLE
Get-OPXExchangeVirtualDirectories -AllCAS -ADPropertiesOnly -ServiceList OWA,ActiveSync
Lists the configuration for the OWA and ACtivSync virtual directories on all Exchange servers with CAS installed.

.EXAMPLE
Get-OPXExchangeVirtualDirectories -AllCAS -ADPropertiesOnly -OnlyVerivyUrls
Verifys the configuration for all virtual directories on all Exchange servers with CAS installed.

#>
[cmdletbinding(DefaultParametersetName = 'ServerList')]
param([Parameter(Mandatory = $false, Position = 0)][ValidateSet('OWA','ECP','OAB','WebServices','ActiveSync','Mapi','OutlookAnywhere','Autodiscover')][array]$ServiceList,
      [Parameter(ParametersetName='ServerList',Mandatory = $false, Position = 1)]
      [ArgumentCompleter( { 
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($false,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*"});
      } )]
      [array]$ServerList=$__OPX_ModuleData.ConnectedToMSXServer,
      [Parameter(ParametersetName='AllCAS',Mandatory = $false, Position = 1)][switch]$AllCAS,      
      [Parameter(ParametersetName='DAG',Mandatory = $false, Position = 1)]
      [ArgumentCompleter( {     
            param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $dagList=($__OPX_ModuleData.getDagList($false,$null));
        $dagList.Where({ $_ -like "$wordToComplete*"}); 
      } )]
      [string]$DAGName,      
      [Parameter(Mandatory = $false, Position = 2)][switch]$ADPropertiesOnly=$false,
      [Parameter(Mandatory = $false, Position = 3)][switch]$OnlyVerifyUrls=$false,
      [Parameter(Mandatory = $false, Position = 4)][switch]$IncludeAttributesFromConfigFile=$false,
      [Parameter(Mandatory = $false, Position = 5)][string]$DomainController
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {                
        switch ($PSCmdlet.ParameterSetName)
        {
            'AllCAS'    {
                if (($ServerList=getAllCAS) -eq $false)
                {
                    return;
                }; # end if
                break;
            }; # end AllCAS
            'DAG'       {                
                if (($ServerList =getDAGMemberServer -DAGName $DAGName)[0] -eq $false)
                {
                    return;
                }; # end if
                break;
            }; # end DAG
        }; # end switch
        
        $serverList=@([System.Linq.Enumerable]::Distinct([string[]]$serverList));
        if (!($PSBoundParameters.ContainsKey('ServiceList')))
        {
            $ServiceList=@('OWA','ECP','OAB','WebServices','ActiveSync','Mapi','OutlookAnywhere','Autodiscover');
        }; # end if
        $paramlist=@{
            ServerList=$serverList;
            ServiceList=$ServiceList;
            OnlyVerifyUrls=$OnlyVerifyUrls;
        }; # end paramList
        if ($PSBoundParameters.ContainsKey('IncludeAttributesFromConfigFile'))
        {
            try {
                $vDirCfgFile=([System.IO.Path]::Combine($__OPX_ModuleData.cfgDir,'VDirs\' + $__OPX_ModuleData.VirtualDirCfgFileName));
                $addCfg=Import-Clixml -Path $vDirCfgFile;
                $paramlist.Add('AdditionalCfg',$addCfg);
            } # end try
            catch {
                writeToLog -LogString ('Failed to load the virtual directory config file (' + $vDirCfgFile + ')') -LogType Error;
                writeToLog -LogString ($_.Exception.Message) -LogType Error;            
            }; # end catch
        }; # end if
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $paramList.Add('DomainController',$DomainController);
        }; # end if
        listVirtualDirectories @paramList;       
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
} # end function Get-OPXVirtualDirectories

function Test-OPXMailboxDatabaseMountStatus
{
<#
.SYNOPSIS 
Tests if all databases are mounted on their preferred server.
	
.DESCRIPTION
The command verifies if a mailbox database is mounted on the preferred server. If not, a message will be displayed on screen.

.PARAMETER DAGName
The parameter is optional.
If a DAG name is provided only this DAG will be verified. If the parameter is omitted, all DAGs are verified.

.PARAMETER PassValue
The parameter is optonal.
The parameter is for internal use.

.PARAMETER DomainController
The parameter is optional.
The name of a domain controller which is used for the Exchange PowerShell cmdlets.

.EXAMPLE
Test-OPXMailboxDatabaseMountStatus -DAGName <dag name>
Test a DAG with the name <dag name>

.EXAMPLE
Test-OPXMailboxDatabaseMountStatus
Test all DAGs
#>
[cmdletbinding()]
param([Parameter(Mandatory = $False, Position = 0)]
      [ArgumentCompleter( {             
            param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $dagList=($__OPX_ModuleData.getDagList($false,$null));
        $dagList.Where({ $_ -like "$wordToComplete*"});
      } )]
      [string]$DAGName,
      [Parameter(Mandatory = $False, Position = 1)][switch]$PassValue = $false,
      [Parameter(Mandatory = $False, Position = 2)][string]$DomainController
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        $DCParam=@{
            ErrorAction='Stop';
        }; # end DCParam
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $DCParam.Add('DomainController',$DomainController);
        }; # end if
        if (!($PSBoundParameters.ContainsKey('DAGName')))
        {
            Write-Verbose 'Query AD for list of DAGs';
            writeTolog -LogString 'Query AD for list of DAGs';
            try
            {
                $dagList=$__OPX_ModuleData.getDagList($false,$null); # get list of DAGs
                if ($DAGList.Count -eq 0)
                {
                    writeToLog -LogType Warning -LogString  'No DAG fond';
                    return;
                }; # end if            
            } # end try
            catch
            {
                writeToLog -LogType Warning -LogString  'Failed to locate a DAG.';
                if ($PassValue)
                {
                    return $False;
                }; # end if
                return;
            }; # end catch
        } # end if psBoundParameters
        else
        {        
            #  verify that DAG exist
            if ([System.Linq.Enumerable]::Contains([string[]]$__OPX_ModuleData.getDagList($false,$null).ToLower(),$DagName.ToLower()) -eq $false)
            {
                writeToLog -LogType Warning -LogString  ('No DAG with name ' + $DAGName + ' found');
                if ($PassValue)
                {
                    return $False;
                } # end if
                else
                {
                    return;
                }; # end else
            } # end if
            else {
                $dagList=@($DAGName);
            }; # end else 
        }; # end else psBoundParameter DAGName

        foreach ($DAGName in $dagList)
        {
            $msg=('Query DAG ' + $DAGName + ' for member servers');
            Write-Verbose $msg;
            writeTolog -LogString $msg;
            try
            {               
                $dagMembers=$__OPX_ModuleData.getDagList($true,$DAGName);
                if (($dagMembers.count -eq 0) -or ([system.string]::IsNullOrEmpty($dagMembers[0])))
                {
                    writeToLog ('DAG ' + $DAGName + ' has no member server.') -LogType Warning;
                    if ($PassValue)
                    {
                        return $NULL;
                    } # end if
                    else {
                        continue;
                    }; # end else               
                } # end if
                else {
                    # get DAG members with CAS
                    $dagMembers=$dagMembers.ToLower();
                    $srvList=@([System.Linq.Enumerable]::Intersect([string[]]$dagMembers,[string[]]$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower()));
                }; # end if                
            } # end try
            catch
            {
                writeToLog -LogType Error -LogString  "Failed to enumurate servers for DAG $DAGName.";
                if ($PassValue)
                {
                    return $true;
                } # end if
                else
                {
                    return;
                }; # end else
            }; # end catch            
            Write-Verbose "Searching for databases which are not mounted on their preferred servers.";
            writeTolog -LogString "Searching for databases which are not mounted on their preferred servers.";
            $db2DistributeList=[System.Collections.ArrayList]::New();
            foreach ($srv in $srvList)
            {
                try {
                    $db2DistributeList.AddRange(@(Get-MailboxDatabaseCopyStatus -Server $srv).where({$_.status -eq 'Mounted' -and ($_.ActivationPreference -ne 1)}));
                } # end try
                catch {
                    $msg=('Failed to verify databases with pref. status > 1 on server ' + $srv);
                    writeToLog -LogString $msg -logType Warning;
                }; # end catch
            }; # end foreach
            if ($db2DistributeList)
            {
                if ($PassValue)
                {
                    return $False;
                } # end if
                else
                {
                    writeToLog -LogType Warning -LogString  'Not all mailbox databases are mounted on their preferred servers.';
                    $tblFields=@(
                        @('Database Name',[system.string]),
                        @('Mounted on Mailbox Server',[system.string]),
                        @('Activation Pref.',[system.int32]),
                        @('Originating Mailbox Server',[system.string]),
                        @('DAG Membership',[system.string])
                    ); # tableFields
                    $dbTable=createNewTable -TableName 'Databases' -FieldList $tblFields;
                    $dbCount=$db2DistributeList.count;
                    for ($i=0;$i -lt $dbCount; $i++)
                    {                        
                        if ($dbCount -gt 5)
                        {
                            Write-Progress -Activity 'Analysing database' -Status ('processing DB ' + ($i+1) + ' of ' + $dbCount) -PercentComplete (100-(100/($i+1)));
                        }; # end if
                        try {
                            $oriDB=(Get-MailboxDatabase -Identity $db2DistributeList[$i].DatabaseName);
                            $oriSrv=(([string]$oriDB.ActivationPreference).Replace(', ',',').split(' ').Replace('[','').Replace(']','')[0].split(',')[0]);
                            $currentSrv=(testIfFQDN -ComputerName $db2DistributeList[$i].MailboxServer);
                            $tmp=(testIfFQDN -ComputerName $oriSrv);
                            if ($tmp -ne $false)
                            {
                                $oriSrv=$tmp;
                            };
                            writeTolog -LogString ('The database ' + ([string]$db2DistributeList[$i].DatabaseName) + ', member in DAG ' + ($oriDB.MasterServerOrAvailabilityGroup) + ', is mounted on server ' + ([string]$currentSrv) + ' (activation preference ' + [string]$db2DistributeList[$i].ActivationPreference + '). The originating server is ' + $oriSrv) -LogType Warning -SupressScreenOutput;
                            [void]($dbTable.rows.Add([string]$db2DistributeList[$i].DatabaseName,[string]$currentSrv,[int]$db2DistributeList[$i].ActivationPreference,$oriSrv,$oriDB.MasterServerOrAvailabilityGroup));
                        } # end try
                        catch {
                            $msg=('Failed to get the data for the database ' +($db2DistributeList[$i].DatabaseName));
                            writeToLog -LogString $msg -LogType Warning;
                            writeToLog -LogString ($_.Exception.Message) -logType Error;
                            continue;
                        }; # end catch                        
                    }; # end for                    
                    Write-Output ''; # create spacer line
                    $dbTable | Format-Table;
                    writeToLog -LogString  ('Please verify if all mailbox servers in the DAG ' + $DAGName + ' are up and running.' + "`n" + 'If databases are not automatically redistributed, run the command Start-OPXMailboxDatabaseRedistribution to redistibute the databases.') -ShowInfo;                   
                }; # end else
            } # end if
            else
            {
                if ($PassValue)
                {
                    return $true;
                } # end if
                else
                {
                    writeToLog -LogType Info -LogString  ('All databases on DAG ' + $DAGName + ' are mounted on their preferred mailbox servers.') -ShowInfo;
                }; # end else
            }; # end else
        }; # end foreach
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
    
} # end function Test-OPXMailboxDatabaseMountStatus

function Get-OPXPreferredServerForMailboxDatabase
{
<#
.SYNOPSIS 
Returns the preferred server for a mailbox database.
	
.DESCRIPTION
Returns the preferred mailbox server for a given mailbox database.

.PARAMETER Identity
The parameter is mandatory.
Name of the mailbox database.

.PARAMETER DomainController
The parameter is optional.
The name of a domain controller which is used for the Exchange PowerShell cmdlets.

.EXAMPLE
Get-OPXPreferredServerForMailboxDatabase  -Identity <database name>
#>

[cmdletbinding()]
param([Parameter(Mandatory = $true, ValueFromPipeline=$True, Position = 0)][string]$Identity,
      [Parameter(Mandatory = $false, Position = 5)][string]$DomainController
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin
    
    Process
    {
        $DCParam=@{
            ErrorAction='Stop';
        }; # end DCParam
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $DCParam.Add('DomainController',$DomainController);
        }; # end if
        Write-Verbose "Retrieving information for database $Identity.";
        writeTolog -LogString "Retrieving information for database $Identity.";
        try
        { 
        $prefServer=(Get-MailboxDatabase -Identity $Identity @DCParam).Server;
        writeTolog -LogString ("The preferred mailbox server for Database $Identity is:`t" + $prefServer) -ShowInfo;
        } # end try
        catch
        {
            writeToLog -LogType Error -LogString  "Mailbox database $Identity was not found.";
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
} # end function Get-OPXPreferredServerForMailboxDatabase


function Resolve-OPXVirtualDirectoriesURLs
{
<#
.SYNOPSIS 
Resolves the FQDNs in the urls of the virtual directories.
	
.DESCRIPTION
Resolves the FQDNs in the urls of the virtual directories.

.PARAMETER ServerList
The parameter is optional.
List of servers where the urls of the virtual directories should be resolved. If the parameter is omitted, the local host will be used.

.PARAMETER DomainController
The parameter is optional.
The name of a domain controller which is used for the Exchange PowerShell cmdlets.

.EXAMPLE
Resolve-OPXVirtualDirectories -ServerList <list of servers>
#>
[CmdLetBinding()]
param(
      [Parameter(Mandatory = $false, Position = 0)]
      [ArgumentCompleter( {             
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($false,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*"});                    
      } )]
      [array]$ServerList,
      [Parameter(Mandatory = $false, Position = 1)][string]$DomainController

     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {       
        $DCParam=@{
            ErrorAction='Stop';
        }; # end DCParam
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $DCParam.Add('DomainController',$DomainController);
        }; # end if
        try
        {        
            Write-Verbose "Building server list";
            writeTolog -LogString "Building server list";
            if (!($PSBoundParameters.ContainsKey("ServerList")))
            {            
                $srvList=($__OPX_ModuleData.getExchangeServerList($false,$true,$true));
                if (! $srvList)
                {
                    writeToLog -LogType Error -LogString  ('No Exchange Client Access server found.');
                    return;
                } # end if     
            }  # end if     
            else
            {           
                $serverList=@([System.Linq.Enumerable]::Distinct([string[]]$serverList)); # get uniqu server entries
                # verify if all servers in list are servers (FQDN) with CAS insalled
                $slc=$serverList.Count;
                for ($i=0;$i -lt $slc;$i++)
                {
                    if (($tmp=testIfFQDN -ComputerName $serverList[$i]) -ne $false)
                    {
                        $serverList[$i]=($tmp.ToLower());
                    }; # end if
                }; # end foreach
                $srvList=@([System.Linq.Enumerable]::Intersect([string[]]$serverList,[string[]]$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower()));                
            }; # end else
        } # end try
        catch
        {
            writeToLog -LogType Error -LogString  'Faild to build a list of Exchange CAS servers.';
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
            return;
        }; # end catch
        
        if (! $srvList)
        {
            writeToLog -LogType Error -LogString  "Failed to collect the list of CAS." 
        } # end if
        else
        {
            $URLList = getURLList -SrvList $SrvList -ADPropertiesOnly;
            Write-Verbose "Resolving URLs";
            writeTolog -LogString "Resolving URLs";
            foreach ($URL in $URLList)
            {
                try
                {
                    Write-Verbose ('Resolving url ' + $url);
                    writeTolog -LogString ('Resolving url ' + $url);
                    [void]([System.Net.DNS]::GetHostAddresses($url));
                } # end try
                catch
                {
                    writeToLog -LogType Warning -LogString  ('Warning: The URL ' + $URL + ' couldnot be resolved.');
                    writeTolog -LogString ($_.Exception.Message) -LogType Warning;
                }; # end catch
            }; # end foreach
        }; # end else
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
} # end function Resolve-OPXVirtualDirectoriesURLs


function Start-OPXMailboxDatabaseRedistribution
{
<#
.SYNOPSIS 
Starts mailbox database redistribution.
	
.DESCRIPTION
The command calls the Exchange script RedistributeActiveDatabases.ps1 with the parameters
	DagName
	BalanceDbsByActivationPreference
	ShowFinalDatabaseDistribution
	Confirm:$false
The command is for older Exchange server versions. Exchange 2016 and newer can redistribute the databases automatically.
.PARAMETER DAGName
The parameter is optional.
Name of the DAG where the databases should be redistributed. If the parameter is omitted, all DAGs will be checked.

.PARAMETER Force
The parameter is optional.
Per default the command verifies if there any mailbox databases mounted on not preferred servers. If this is true the Microsoft script will not be called. With the parameter Force you can force the command to run RedistributeActiveDatabases.ps1.

.EXAMPLE
Start-OPXMailboxDatabaseRedistribution -DAGName <DAG name>
#>
[cmdletbinding()]
param([Parameter(Mandatory = $False, Position = 0)]
      [ArgumentCompleter( {             
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $dagList=($__OPX_ModuleData.getDagList($false,$null));
        $dagList.Where({ $_ -like "$wordToComplete*"});  
      } )]
      [string]$DAGName,
      [Parameter(Mandatory = $False, Position = 1)][switch]$Force = $false
     )
    
    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        if ($null -eq $Script:MSXScriptDir)
        {
            writeToLog -LogType Warning -LogString  'Please run Start-OPXMailboxDatabaseRedistribution on a computer with the Exchange managemt tools installed.';
            return;
        }; # end if
        if (!($PSBoundParameters.ContainsKey('DAGName')))
        {
            Write-Verbose 'Query AD for list of DAGs';  
            writeTolog -LogString 'Query AD for list of DAGs';  
            try
            {
                $DAGList=$__OPX_ModuleData.getDagList($false,$null); # get list of DAGs in Exchange org
                switch ($DAGList.count)
                {
                    0       { 
                        writeToLog -LogType Warning -LogString  'No DAG found';
                        return;
                    }; # end count 0
                    1       {
                        $DAGName = $DAGList[0]; # 1 DAG found, assign name to var
                    }; # end count 1
                    default { # more than 1 DAG found stopp
                        writeToLog -LogType Warning -LogString  'More than one DAG found. Please specify the DAG with the parameter DAGName.';
                        return;
                    }; # end more then 1
                }; # end switch
            } # end try
            catch
            {
                writeToLog -LogType Error -LogString  'Failed to locate a DAG.';
                writeToLog -LogString ($_.Exception.Message) -LogType Error;
                return;
            }; # end catch
        }; # end if
        
        Write-Verbose 'Verifying if databases are mounted on the preferred server.';
        writeTolog -LogString 'Verifiying if databases are mounted on the preferred server.';

        if ((($rv=Test-OPXMailboxDatabaseMountStatus -DAGName $DAGName -PassValue) -eq $false) -or $Force.IsPresent)
        {
            Write-Verbose "Query DAG $DAGName for member servers";
            writeTolog -LogString "Query DAG $DAGName for member servers";
            try
            {
                Write-Verbose "Checking for DAG members.";
                writeTolog -LogString "Checking for DAG members.";
                if (($memberCount=($__OPX_ModuleData.getDagList($true,$DagName)).count) -lt 1)
                {
                    writeToLog -LogType Error -LogString  ('Faild to enumerate the member servers for the DAG ' + $DAGName);
                    return;
                }; # end if
                if ($memberCount -eq 1) # verify member count
                {
                    writeToLog -LogString ('The DAG ' + $dagName + ' has only one member server. No redistribution needed.');
                    return;
                };
            } # end try
            catch
            {
                writeToLog -LogType Error -LogString  ('Faild to enumerate the member servers for the DAG ' + $DAGName + ', or the DAG has no member.');
                writeToLog -LogString ($_.Exception.Message) -LogType Error;
                return;            
            }; # end catch

            $pList=@{
                DAGName=$DAGName;
            }; # end pList
            $msg='Running script RedistributeActiveDatabases.ps1 from Exchange scripts directory'
            Write-Verbose $msg;
            writeToLog -LogString $msg;
            runScriptFromExScriptsDir -ScriptName 'RedistributeActiveDatabases' -ScriptParams $pList;
            $msg='Finished running script RedistributeActiveDatabases.ps1 from Exchange scripts directory'
            Write-Verbose $msg;
            writeToLog -LogString $msg;
        } # end if
        else
        {
            if ($rv)
            {
                writeTolog -LogString  'All mailbox databases are mounted on the preferred server. No redistribution needed.' -ShowInfo;
            }; # end if            
        }; # end else
    }; # end process
    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END    
} # end function Start-OPXMailboxDatabaseRedistribution


function Start-OPXExchangeServerMaintenance
{
<#
.SYNOPSIS 
Starts maintenance of an Exchange server.
	
.DESCRIPTION
Start the maintenance of an Exchange server.

.PARAMETER ServerFQDN
The parameter is optional.
The FQDN of the server which should be configured for maintenance. If you omit the parameter, the name of the locale computer is used.

.PARAMETER TargetTransportServerFQDN
The parameter is mandatory.
The name of a transport server for redirecting the messages. If the switch SingleServerEnvironment is use, the parameter cannot be used.

.PARAMETER SingleServerEnvironment
If only one Exchange server is available, this parameter should be used.

.PARAMETER DomainController
The parameter is optional.
The name of a domain controller which is used for the Exchange PowerShell cmdlets.

.EXAMPLE
Start-OPXExchangeServerMaintenance -ServerFQDN <server FQDN> -TargetTransportServerFQDN <target transport server FQDN>
#>
[cmdletbinding(DefaultParameterSetName='RedirToTransport')]
param([Parameter(Mandatory = $false, Position = 0)]
      [ArgumentCompleter( {
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($true,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*"});
      } )]
      [string]$ServerFQDN=$__OPX_ModuleData.ConnectedToMSXServer,
      [Parameter(ParameterSetName='RedirToTransport',Mandatory = $true, Position = 1)]
      [ArgumentCompleter( {   
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($true,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*"}); 
      } )]
      [string]$TargetTransportServerFQDN,
      [Parameter(ParameterSetName='SingelServerEnv',Mandatory = $true, Position = 1)][switch]$SingleServerEnvironment,
      [Parameter(Mandatory = $false, Position = 2)][string]$DomainController
     )
    
    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        if ($null -eq $Script:MSXScriptDir)
        {
            writeToLog -LogType Warning -LogString  'Please run Start-OPXExchangeServerMaintenance on a computer with the Exchange managemt tools installed.';
            return;
        }; # end if
        $DCParam=@{
            ErrorAction='Stop';
        }; # end DCParam
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $DCParam.Add('DomainController',$DomainController);
        }; # end if
        try
        {
            if ($PsCmdlet.ParameterSetName -eq 'RedirToTransport')
            {
                try {
                    # verify if the target transport servr is a transport server and the component HubTransport is active
                    if (($TargetTransportServerFQDN=testIfFQDN -ComputerName $TargetTransportServerFQDN ) -eq $false)
                    {
                        writeTolog -LogString 'Target transport server not found.' -LogType Error;
                        return;
                    }; # end if
                    # verify component HubTransport
                    if (! (($srvObj=Get-ServerComponentState -Component HubTransport -Identity $TargetTransportServerFQDN @DCParam).State -eq 'Active'))
                    {
                        $msg=('The component HubTransprot on server ' + $TargetTransportServerFQDN + ' is not active. Current state is ' + $srvObj.State + '.');
                        writeToLog -LogString $msg -LogType Error;
                        writeTolog -LogString 'Please select a different Target-Transport-Server!' -LogType Warning;
                        return;
                    }; # end if
                } # end try
                catch {
                    writeToLog -LogType Warning -LogString  ('Failed to verivy if the state of the component HubTransport is active on server' + $TargetTransportServerFQDN);
                    writeTolog -LogString ($_.Exception.Message) -LogType Warning;
                    return;
                }; # end catch
            }; # end if
            if (($serverFQDN=testIfFQDN -ComputerName $serverFQDN) -eq $false)
            {
                writeTolog -LogString 'Server not found.' -LogType Error;
                return;
            }; # end if            
            if ($serverFQDN -eq $TargetTransportServerFQDN)
            {
                writeTolog -LogString ('The values for ServerFQDN and TargetTransportServer must not be the same.') -LogType Error;
                return;
            }; # end if
        } # end try
        catch
        {
            writeToLog -LogType Warning -LogString  ('Failed to build FQDN for ' + $ServerFQDN);
            writeTolog -LogString ($_.Exception.Message) -LogType Warning;
            return;
        }; # end catch
        
        $isMemberOfDAG = isDAGMember -ServerName $ServerFQDN;
        $sCParams=@{
            Server=$ServerFQDN;
            Component='HubTransport';
            State='Draining';
            Requester='Maintenance';
        }; # sCParams
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $sCParams.Add('DomainController',$DomainController);
        }; # end if
        setComponentState @sCParams;
        try {
            $msg=('Restarting MSExchangeTransport on server ' + $ServerFQDN);
            Write-Verbose $msg;
            writeTolog -LogString $msg;
            Get-Service -ComputerName $ServerFQDN -Name MSExchangeTransport | Restart-Service -Confirm:$False -ErrorAction Stop;
                      
        } # end try
        catch {
            $msg=('Failed to configure transport for maintenaance on server ' + $ServerFQDN);
            writeTolog -LogString $msg -logType Error;
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch
        
        $DCParam.ErrorAction='SilentlyContinue';
        if (Get-ServerComponentState -Identity $ServerFQDN -Component UMCallRouter @DCParam)
        {
            $sCParams.Component='UMCallRouter';
            setComponentState @sCParams;            
        } # end if
        else {
            $DCParam.ErrorAction='SilentlyContinue';
        }; # end else
               
        try {
            if ($isMemberOfDAG)
            {                          
                $pList=@{
                    ServerName=($serverFQDN.split('.'))[0];
                    MoveComment='Maintenance';
                    PauseClusterNode=$true;
                }; # end pList
                $msg='Running script StartDagServerMaintenance.ps1 from Exchange scripts directory'
                Write-Verbose $msg;
                writeToLog -LogString $msg;
                runScriptFromExScriptsDir -ScriptName 'StartDagServerMaintenance' -ScriptParams $pList;
                $msg='Finished running script StartDagServerMaintenance.ps1 from Exchange scripts directory'
                Write-Verbose $msg;
                writeToLog -LogString $msg;
                try {
                    writeTolog -LogString "Disable database copy move to server $ServerFQDN";
                    Set-MailboxServer $ServerFQDN -DatabaseCopyActivationDisabledAndMoveNow $True @DCParam;
                } # end try
                catch {
                    $msg=('Failed to configure DatabaseCopyActivationDisabledAndMoveNow on mailbox server '  + $ServerFQDN );
                writeTolog -LogString $msg -logType Error;
                writeToLog -LogString ($_.Exception.Message) -LogType Error;
                }; # end catch
            }; # end if            
        } # end try
        catch {
            $msg=('Failed to configure on server '  + $ServerFQDN + ' database copy management for maintenance ');
            writeTolog -LogString $msg -logType Error;
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch
       
        if ($PSCmdlet.ParameterSetName -eq 'RedirToTransport')
        {
            Write-Verbose "Redirect messages";
            writeTolog -LogString "Redirect messages";
            Redirect-Message -Server $ServerFQDN -Target $TargetTransportServerFQDN -Confirm:$False @DCParam;
        } # end if
        else {
            $msg='Option SingeleServerEnvironment selected. Do not redirect messages.';
            Write-Verbose $msg;
            writeToLog -LogString $msg;
        }; # end else

        $sCParams.Component='ServerWideOffline';
        $sCParams.State='Inactive';
        setComponentState @sCParams;
       
        try {    
            Write-Verbose "Checking configuration";
            writeTolog -LogString "Checking configuration";
            Get-ServerComponentState $ServerFQDN @DCParam| Format-Table Component,State -Autosize;
            
        } # end try
        catch {
            $msg=('Failed to verify the components on '  + $ServerFQDN);
            writeTolog -LogString $msg -logType Error;
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch
        try {
            Get-MailboxServer $ServerFQDN @DCParam| Format-Table DatabaseCopy* -Autosize;
        }
        catch {
            $msg=('Failed to verify the database copies on mailbox server '  + $ServerFQDN);
            writeTolog -LogString $msg -logType Error;
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch
        try {
            if ($isMemberOfDAG)
            {           
                Invoke-Command -ComputerName ($__OPX_ModuleData.ConnectedToMSXServer) -ScriptBlock {param($p1) Get-ClusterNode $p1} -ArgumentList $ServerFQDN;              
            }; #end if

            Get-Queue -Server $ServerFQDN @DCParam;
        } # end try
        catch {
            $msg=('Failed to verify some components for maintenance on '  + $ServerFQDN );
            writeTolog -LogString $msg -logType Error;
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
        }; # catch
            
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END  
} # end function Start-OPXExchangeServerMaintenance

function Remove-OPXExchangeServerFromMaintenance
{
<#
.SYNOPSIS 
Remove an Exchange server from maintenance.
	
.DESCRIPTION
Remove an Exchange server from maintenance.

.PARAMETER ServerFQDN
The parameter is optional.
The FQDN of the server which should be removed from maintenance. If you omit the parameter, the name of the locale computer is used.

.PARAMETER Force
The command tries to check if the server is in maintenance. If it thinks the server is not in maintenance it will return. With the Force switch, the command will not rely on the check.

.PARAMETER DomainController
The parameter is optional.
The name of a domain controller which is used for the Exchange PowerShell cmdlets.

.EXAMPLE
Remove-OPXExchangeServerFromMaintenance -ServerFQDN <server FQDN>
#>

[cmdletbinding()]
param([Parameter(Mandatory = $false, Position = 0)]
      [ArgumentCompleter( { 
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($true,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*"});
      } )]
      [string]$ServerFQDN=$__OPX_ModuleData.ConnectedToMSXServer,# = [System.Net.DNS]::GetHostByName('').HostName,
      [Parameter(Mandatory = $false, Position = 1)][string]$DomainController,
      [Parameter(Mandatory = $false, Position = 2)][switch]$Force
     )
   
    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        $DCParam=@{
            ErrorAction='Stop';
        }; # end DCParam
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $DCParam.Add('DomainController',$DomainController);
        }; # end if
        try
        {
            $ServerFQDN = [System.Net.DNS]::GetHostByName($ServerFQDN).HostName;    
        } # end try
        catch
        {
            writeToLog -LogType Warning -LogString  ('Failed to build FQDN for ' + $ServerFQDN);
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch
        
        try {
            $msg=('Verifyng maintenance state of server ' + $ServerFQDN);
            writeToLog -LogString $msg;
            Write-Verbose $msg;
            $activeComponents=(((Get-ServerComponentState -Identity $serverFQDN  @DCParam).state).Where({$_ -eq 'active'})).count
            if (($Force.IsPresent -eq $false) -and ($activeComponents -gt 2))
            {
                $msg=('The server ' + $serverFQDN + ' is not in an expected maintenance state. Currently are ' + $activeComponents + ' components activ.');
                writeToLog -LogString $msg -LogType Warning;
                $msg=('Use the Force switch to forcibly move the server ' + $serverFQDN + ' out of maintenance.')
                writeToLog -LogString $msg -LogType Warning;
                return;
            };
        } # end try
        catch {
            writeToLog -LogType Error -LogString  ('Failed to get the maintenance state of server ' + $serverFQDN);
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
            return;
        }; # end catch

        $isMemberOfDAG = isDAGMember -ServerName $ServerFQDN;
        $sCParams=@{
            Server=$ServerFQDN;
            Component='ServerWideOffline';
            State='Active';
            Requester='Maintenance';
        }; # sCParams
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $sCParams.Add('DomainController',$DomainController);
        }; # end if
        try {
            setComponentState @sCParams;
            if (Get-ServerComponentState -Identity $ServerFQDN -Component UMCallRouter @DCParam)
            {
                $sCParams.Component='UMCallRouter';
                setComponentState @sCParams;                    
            }; # end if            
        } # end try
        catch {
            $msg=('Failed to configure some component states on server ' + $ServerFQDN);
            writeTolog -LogString $msg -logType Error;
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
            $DCParam.ErrorAction='Stop';  
        }; # end catch        

        if ($isMemberOfDAG)
        {
            $pList=@{
                ServerName=($serverFQDN.split('.'))[0];                
            }; # end pList
            $msg='Running script StopDagServerMaintenance.ps1 from Exchange scripts directory'
            Write-Verbose $msg;
            writeToLog -LogString $msg;
            runScriptFromExScriptsDir -ScriptName 'StopDagServerMaintenance' -ScriptParams $pList;
            $msg='Finished running script StopDagServerMaintenance.ps1 from Exchange scripts directory'
            Write-Verbose $msg;
            writeToLog -LogString $msg;                
        }; # end if
        $sCParams.Component='HubTransport';
        setComponentState @sCParams;
        try {
            Write-Verbose "Checking configuration";
            writeTolog -LogString "Checking configuration";
            Get-Service -ComputerName $ServerFQDN -Name MSExchangeTransport | Restart-Service -Confirm:$False;
        } # end try
        catch {
            $msg=('Failed to start the transport on ' + $ServerFQDN);
            writeTolog -LogString $msg -logType Error;
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch
        
            
        Get-ServerComponentState $ServerFQDN @DCParam| Format-Table Component,State -Autosize;        
        if ($isMemberOfDAG)
        {
            writeToLog -LogType Warning -LogString  'If the databases are not automatically redistributed, please run the command Start-OPXMailboxDatabaseRedistribution to redistribute the databases.';
        }; # end if
    }; # end process
        
    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
} # end Remove-OPXExchangeServerFromMaintenance

function Test-OPXExchangeServerMaintenanceState
{
<#
.SYNOPSIS 
Test the maintenance state of an Exchange server.
	
.DESCRIPTION
Test the maintenance state of an Exchange server.

.PARAMETER ServerFQDN
The parameter is optional.
The FQDN of the server which should be tested if in maintenance. If you omit the parameter, the name of the locale computer is used.

.PARAMETER DomainController
The parameter is optional.
The name of a domain controller which is used for the Exchange PowerShell cmdlets.

.EXAMPLE
Test-OPXExchangeServerMaintenanceState -ServerFQDN <server FQDN>
#>
[cmdletbinding()]
param([Parameter(Mandatory = $false, Position = 0)]
      [ArgumentCompleter( {             
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($true,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*"});
      } )]
      [string]$ServerFQDN=$__OPX_ModuleData.ConnectedToMSXServer,
      [Parameter(Mandatory = $false, Position = 1)][string]$DomainController
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        $DCParam=@{
            ErrorAction='Stop';
        }; # end DCParam
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $DCParam.Add('DomainController',$DomainController);
        }; # end if
        try {
            $ServerFQDN=[System.Net.DNS]::GetHostByName($ServerFQDN).HostName;
            $isMemberOfDAG = isDAGMember -ServerName $ServerFQDN;
            Write-Verbose "Checking component state on server $ServerFQDN";
            writeTolog -LogString "Checking component state on server $ServerFQDN";
            $cS=Get-ServerComponentState $ServerFQDN @DCParam | Format-Table Component,State -Autosize;
            $cs
            Write-Verbose "Checking database activation on server $ServerFQDN";
            writeTolog -LogString "Checking database activation on server $ServerFQDN";
            Get-MailboxServer -Identity $ServerFQDN @DCParam | Format-List DatabaseCopyAutoActivationPolicy,DatabaseCopyActivationDisabledAndMoveNow;
        } # end try
        catch {
            $msg=('Failed to test the maintenace state for server ' + $ServerFQDN);
            writeTolog -LogString $msg -logType Error;
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch
        
        try {
            if ($isMemberOfDAG)
            {
                Write-Verbose "Checking cluster node $ServerFQDN";
                writeTolog -LogString "Checking cluster node $ServerFQDN";
                Invoke-Command -ComputerName ($__OPX_ModuleData.ConnectedToMSXServer) -ScriptBlock {param($p1) Get-ClusterNode $p1} -ArgumentList $ServerFQDN;
                
            }; # end if
        } # end try
        catch {
            $msg=('Failed to checking the cluster on server ' + $ServerFQDN);
            writeTolog -LogString $msg -logType Error;
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch
        

        Write-Verbose "Checking queue on server $ServerFQDN";
        writeTolog -LogString "Checking queue on server $ServerFQDN";
        try
        {
            Get-Queue -Server $ServerFQDN -ErrorAction 'Stop' | Format-Table;
        } # end try
        catch
        {
            writeToLog -LogType Error -LogString  'Transport not running';
        }; #end catch
    }; # end process
    
    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
} # end function Test-OPXExchaneServerMaintenanceState


function Get-OPXExchangeSchemaVersion
{
<#
.SYNOPSIS 
Display the Exchange schema and object versoion.
	
.DESCRIPTION
Display the exchange schema version and the Exchange object version for forest and domain. The command can query all GCs for the versions. If an upgrade of Exchange will be implemented, with the switch WaitForSchemaReplication, the replication can be monitored. The command will stop when all DCs have the same version. To stop the command prior reaching that point, press Escape.

.PARAMETER ServerList
The parameter is optional.
A list of servers which should be queried for the versions. If the parameter is omitted, the command quires for the versions without specifying a DC.
The parameter cannot be used with the parameters QueryAllDomainControllers or WaitForSchemaReplication.

.PARAMETER QueryAllDomainControllers
The parameter is optional.
The command queries all GCs for the versions.
The parameter cannot be used with the parameters ServerList or WaitForSchemaReplication.

.PARAMETER WaitForSchemaReplication
The parameter is optional.
The command queries, in intervals, all GCs for the versions. The script continues until all GCs have the same version numbers. If you want to stop the script earlier, press Escape.
The parameter cannot be used with the parameters ServerList or QueryAllDomainControllers.

.PARAMETER WaitForSchemaReplicationIntervalInSeconds
The parameter is optional.
The parameter specifies the interval for the parameter WaitForSchemaReplication in seconds. If the parameter is omitted, the command will pause 30 seconds between queries.

.EXAMPLE
Get-OPXExchangeSchemaVersion
#>
[cmdletbinding(DefaultParameterSetName='Default')]
param([Parameter(ParametersetName='ServerList',Mandatory = $False, Position = 0)][array]$ServerFQDNList,
      [Parameter(ParametersetName='QueryAllDCs',Mandatory = $false, Position = 1)][switch]$QueryAllDomainControllers,
      [Parameter(ParametersetName='WaitForRepl',Mandatory = $true, Position = 2)][switch]$WaitForSchemaReplication,      
      [Parameter(ParametersetName='WaitForRepl',Mandatory = $false, Position = 3)][int]$WaitForSchemaReplicationIntervalInSeconds=30     
     )
     
    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        try
        {            
            $SchemaRoot=([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Schema.name);
            $configRoot=$schemaRoot.Remove(0,10);
            $msxCfgRoot='CN=Microsoft Exchange,CN=Services,'+$configRoot;
            $msxSchemaRoot='CN=ms-Exch-Schema-Version-Pt,'+$SchemaRoot;
            $columnList=@(  # list of table columns
                @('Domain controller name',[system.string]),
                @('RangeUpper',[System.int32]),
                @('Forest objVersion',[System.int32]),
                @('Domain objVersion',[System.int32])
            ); # end filedList
            $tblPL=@{
                    TableName='Results';
                    FieldList=$columnList;
                }; # end tblPl
                $outTbl=createNewTable @tblPL;
            $fParamList=@{
                LDAPFilter='(objectClass=msExchOrganizationContainer)';
                SearchBase='';#$msxSchemaRoot;
                FieldList=@('objectversion');
            }; # end paramList       
            $dParamList=@{
                LDAPFilter='(objectClass=msExchSystemObjectsContainer)';
                SearchBase='CN=Microsoft Exchange System Objects,'+([ADSI]'LDAP://RootDSE').defaultNamingContext;
                FieldList=@('objectversion');
            }; # end dParamList
                
            if ($PsCmdlet.ParameterSetName -eq 'ServerList') # convert server names to FQDN, if not already
            {
                #$tmpList=@();
                $ServerFQDNList=@([System.Linq.Enumerable]::Distinct([string[]]$ServerFQDNList));
                $srvC=$ServerFQDNList.count;
                for ($i=0;$i -lt $srvC; $i++)
                {
                    if (($tmp=testIfFQDN -ComputerName $ServerFQDNList[$i]) -ne $false)
                    {
                        $ServerFQDNList[$i]=$tmp;       
                    };
                }; # end if
                $fParamList.Add('Server','');
                $dParamList.Add('Server','');
            }; # end if
            if (($PSCmdlet.ParameterSetName -eq 'WaitForRepl')  -or ($PSCmdlet.ParameterSetName -eq 'QueryAllDCs'))# build list of DC's in forest and 
            {
                $roleOwner=@([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().SchemaRoleOwner.name);
                $ServerFQDNList=$roleOwner;
                $ServerFQDNList+=(@([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().GlobalCatalogs.name) -ne $ServerFQDNList);
                $fParamList.Add('Server',''); # params for forest
                $dParamList.Add('Server',''); # params for domain       
            }; # end if
            
            if (($PSBoundParameters.ContainsKey('ServerFQDNList')) -or ($PSCmdlet.ParameterSetName -eq 'WaitForRepl') -or ($PSCmdlet.ParameterSetName -eq 'QueryAllDCs'))
            {               
                # init ref value 
                $refVal=@($null,$null,$null,$null);                     
                $srvLc=$ServerFQDNList.Count;
                $replNOK=(($WaitForSchemaReplication.IsPresent) -and ($srvLc -gt 1));  
                do 
                {                
                    for ($i=0; $i -lt $srvLc; $i++)
                    {
                        writeTolog -LogString ('Quering DC ' + $ServerFQDNList[$i] + ' for Exchange schema and object version.')
                        $replOk=$true;
                        $serverRoot=([ADSI]('LDAP://'+$ServerFQDNList[$i]+'/RootDSE')).defaultNamingContext;
                        try
                        {
                            [void][System.Net.DNS]::GetHostByName($ServerFQDNList[$i]).HostName;
                        } # end try
                        catch
                        {
                            writeToLog -LogType Warning -LogString  ('Failed to resolve the server ' + $ServerFQDNList[$i]);
                            continue;
                        }; # end catch
                        $outTblFieldsVal=getMSXVersions -Server $ServerFQDNList[$i] -fParamList $fParamList -dParamList $dParamList -serverRoot $serverRoot                    
                        
                        [void]($outTblFields=[System.Collections.ArrayList]::new());
                        [void]$outTblFields.Add((($ServerFQDNList[$i].ToLower())));
                        if ($null -ne $outTblFieldsVal.msxSchemaInfo)
                        {
                            [void]$outTblFields.Add(([string]$outTblFieldsVal.msxSchemaInfo.Properties.rangeupper));
                        } # end if
                        else
                        {
                            [void]$outTblFields.Add(($null));
                        }; # end else
                        if ($null -ne $outTblFieldsVal.msxConfigInfo)
                        {
                            [void]$outTblFields.Add((([string]$outTblFieldsVal.msxConfigInfo.Properties.objectversion)));
                        } # end if
                        else
                        {
                            [void]$outTblFields.Add(($null));
                        }; # end else
                        if ($null -ne $outTblFieldsVal.domInfo)
                        {
                            [void]$outTblFields.Add((([string]$outTblFieldsVal.domInfo.Properties.objectversion)));
                        } # end if
                        else
                        {
                            [void]$outTblFields.Add(($null));
                        }; # end else

                        [void]($outTbl.rows.Add($outTblFields.ToArray())); # add row to table 
                    
                        if($PsCmdlet.ParameterSetName -eq 'WaitForRepl')
                        {                                                                                         
                            if (($refVal[1] -ne $outTblFields[1]) -or ($refVal[2] -ne $outTblFields[2]) -or ($refVal[3] -ne $outTblFields[3]))
                            {
                                switch ($true)
                                {
                                    {$refVal[1] -ne $outTblFields[1]} {
                                        $refval[1]=[math]::Max($refval[1],$outTblFields[1]);
                                    };
                                    {$refVal[2] -ne $outTblFields[2]} {
                                        $refval[2]=[math]::Max($refval[2],$outTblFields[2]);
                                    };
                                    {$refVal[3] -ne $outTblFields[3]} {
                                        $refval[3]=[math]::Max($refval[3],$outTblFields[3]);
                                    };
                                }; #end switch
                                $replOk=$false;
                            }; # end if                        
                           $replNOK=(! ($replOk -or ($srvLc -eq 1))); # verify if data was replicated and more than one GC exists
                        }; # end if
                    }; # end for(each)

                    $outTbl | Format-Table;
                    [void]($outTbl.rows.clear());   # remove entries from table         
                    Write-Output '';          
                    if (($replNOk)) # replication is not ok
                    {
                        if ($host.Name -in ('ConsoleHost','Visual Studio Code Host')) # verify host and wait until WaitForSchemaReplicationIntervalInSeconds has passed
                        {
                            for ($i=0;$i -le ($WaitForSchemaReplicationIntervalInSeconds * 10);$i++)
                            {
                                if ([Console]::KeyAvailable) 
                                {
                                    if (($host.UI.RawUI.ReadKey('NoEcho, IncludeKeyDown')).VirtualKeyCode -eq 27)  
                                    {
                                        $replNOK=$false;
                                        break;
                                    } # end if                                
                                }; # end if 
                                Start-Sleep -Milliseconds 100;                           
                            }; # end for
                        } #end if
                        else
                        {
                            for ($i=0;$i -le ($WaitForSchemaReplicationIntervalInSeconds * 10);$i++)
                            {
                                if ([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::Escape))
                                {
                                    $replNOK=$false;
                                    break;                                
                                }; #end if
                                Start-Sleep -Milliseconds 100;
                            }; # end for                       
                        }; # end else                    
                    }; # end if (replOk)
                    
                } while ($replNOK); # do while loop
            } # end if
            else
            {
                $fParamList.LDAPFilter='(objectClass=attributeSchema)';
                $fParamList.SearchBase=$msxSchemaRoot; 
                $fParamList.FieldList=@('rangeUpper');
                $msxSchemaInfo=queryADForObjects @fParamList;
                $fParamList.SearchBase=$msxCfgRoot; 
                $fParamList.FieldList=@('objectversion');
                $fParamList.Add('SearchScope','OneLevel');
                $fParamList.LDAPFilter='(objectClass=msExchOrganizationContainer)';
                $msxConfigInfo=queryADForObjects @fParamList;
                $dParamList.FieldList=@('objectversion');
                $domInfo=queryADForObjects @dParamList;
                $outTblFields=@('N/A');
                
                if ($null -ne $msxSchemaInfo)
                {
                    $outTblFields+=(([string]$msxSchemaInfo.Properties.rangeupper));
                } # end if
                else
                {
                        $outTblFields+=($null);
                }; # end else
                if ($null -ne $msxConfigInfo)
                {
                    $outTblFields+=(([string]$msxConfigInfo.Properties.objectversion));
                } # end if
                else
                {
                        $outTblFields+=($null);
                }; # end else
                if ($null -ne $domInfo)
                {
                    $outTblFields+=(([string]$domInfo.Properties.objectversion));
                } # end if
                else
                {
                        $outTblFields+=($null);
                }; # end else                
                    [void]($outTbl.rows.Add($outTblFields)); # add row to table                                
                $outTbl | Format-Table RangeUpper,'*objVersion';            
            }; # end else       
        } # end try
        catch
        {
            writeToLog -LogType Error -LogString  'Failed to read the Exchange schema version.';
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch  
    }; # end process
    
    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END 
} # end function Get-OPXExchangeSchemaVersion


function Send-OPXTestMailMessages
{
<#
.SYNOPSIS 
Send e-mails for testing purpose.
	
.DESCRIPTION
To send test mail messages in lab environment. The command is experimental.

.PARAMETER SMTPServer
The parameter is optional.
The name of the SMTP server. If the parameter is omitted, the command randomly sends the message to an Exchange mailbox server.

.PARAMETER MailFrom
The parameter is mandatory.
The parameter expects one of the following:
    e-mail address of the sender
    path to a file with e-mail addresses from sender
    AD path to OU with mailbox users, mail users or mail contacts
The command will try to find out what type of input (e-mail address, file path or AD path)  is provided

.PARAMETER SendMailToDirectoryPath
The parameter is mandatory for the parameter set 'SendToOU'.
An OU path is expected. Sends mail messages to mailbox user in an organizational unit.
The parameter cannot be used with the parameter SendMailToInputFilePath

.PARAMETER SendMailToInputFilePath
The parameter is mandatory for the parameter set 'SendToInputFile'.
A file path is expected. Sends mail messages to mail addresses form a file.
The parameter cannot be used with the parameter SendMailToDirectoryPath

.PARAMETER MailBodyDirectoryPath
The parameter is mandatory.
A directory path is expected. In the directory path are text files with content for mail bodies expected.

.PARAMETER MailAttachmetDirectoryPath
The parameter is mandatory.
A directory path is expected. In the directory path are files, which can be used as attachment, expected.

.PARAMETER MailBatchCount
The parameter is optional.
The number of times mails will be sent. The default for the parameter is 1.

.PARAMETER MailSendIntervallInMilliseconds
The parameter is optional.
The interval between mails in milliseconds. The default for the parameter is 100.
#>
[cmdletbinding(DefaultParameterSetName='SendToSingle')]
param([Parameter(Mandatory = $false, Position = 0)][string]$SMTPServer,
        [Parameter(Mandatory = $true, Position = 1)][string]$MailFrom,
        [Parameter(ParametersetName='SendToOU',Mandatory = $true, Position = 2)][string]$SendMailToDirectoryPath,
        [Parameter(ParametersetName='SendToInputFile',Mandatory = $true, Position = 2)][string]$SendMailToInputFilePath,      
        [Parameter(Mandatory = $true, Position = 3)][string]$MailBodyDirectoryPath,
        [Parameter(Mandatory = $true, Position = 4)][string]$MailAttachmentDirectoryPath,
        [Parameter(Mandatory = $false, Position = 5)][int]$MailCountToSend=1,
        [Parameter(Mandatory = $false, Position = 7)][int]$MailSendIntervallInMilliseconds=100
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        try {
            if ($PSCmdlet.ParameterSetName -eq 'SendToInputFile')
            {
                $msg=('Reading content of file ' + $SendMailToInputFilePath);
                Write-Verbose $msg;
                writeToLog -LogString $msg;
                $sendMailsToList=@(Get-Content -Path $SendMailToInputFilePath);
            } # end if
            else {
                $msg=('Reading e-mail addresses form OU ' + $SendMailToDirectoryPath);
                Write-Verbose $msg;
                writeToLog -LogString $msg;
                $pL=@{
                    LDAPFilter='(&(objectClass=user)(mail=*)(!(objectClass=computer)))';
                    SearchBase=$SendMailToDirectoryPath;
                    FieldList=@('cn','mail');            
                    #SearchScope='One';
                }; # end Pl
                if (($tmp=(queryADForObjects @pl)).count -eq 0)
                {
                    $msg='The list of recipients is empty. Cannot send test mails.'
                    writeToLog -LogString $msg -LogType Warning;
                    return;
                }; # end if                
                $sendMailsToList=@($tmp.properties.mail)
            }; # end else
            
            $mailBodyFileList=@((Get-ChildItem -Path $MailBodyDirectoryPath -File -Filter '*.txt').FullName);
            $mailAttachmentFileList=@((Get-ChildItem -Path $mailAttachmentDirectoryPath).FullName);
        } # end try
        catch {
            writeToLog -LogType Error -LogString  'Failed to init script.';
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
            return;
        }; # end catch
        switch ($MailFrom)
        {
            {$_ -as [System.Net.Mail.MailAddress]}  {
                $senderList = @($MailFrom;)
            }; # end e-mail address
            {Test-Path -PathType Leaf -Path $_}     {
                try {
                    $senderList=Get-Content -Path $MailFrom -ErrorAction Stop;
                } # end try
                catch {
                    $msg=('Failed to read the content of the file ' + $MailFrom)
                }; # end catch
            }; # end path
            Default                                 {
                try {
                    $pL=@{
                        LDAPFilter='(|(&(objectClass=user)(mail=*)(!(objectClass=computer)))(&((objectClass=contact)(TargetAddress=*))))';
                        SearchBase=$MailFrom;
                        FieldList=@('cn','mail');
                    }; # end Pl
                    Remove-Variable Tmp;
                    if (($tmp=(queryADForObjects @pl)).count -eq 0)
                    {
                        $msg='The list of senders is empty. Cannot send test mails.';
                        writeToLog -LogString $msg -LogType Warning;
                        return;
                    }; # end if                   
                    $senderList=@($tmp.properties.mail);
                } # end try
                catch {
                    $msg='Failed to query AD for list of senders.'
                    writeTolog -LogString $msg -LogType Error;
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;
                    return;
                }; # end catch
                #>
            }; # end query AD
        }; # end swtitch
        
        $wpParams=@{
            Activity=('Sending test mails');
            Status=('');
            PercentComplete=0;
        }; # end wpParams
        $sendToRandomSMTP=(! $PSBoundParameters.ContainsKey('SMTPServer'));
        writeToLog ('Sending to random SMTP server is ' + $sendToRandomSMTP);
        for ($j=0; $j -lt $MailCountToSend; $j++)
        {
            $cMc=$j+1;
            if ($sendToRandomSMTP)
            {
                $SMTPServer=getSMTPServer;
            }; # end if
            if ($SMTPServer -ne $false)
            {
                $mailSendParams=@{
                    From=($senderList | Get-Random);
                    SMTPServer=$SMTPServer;
                    Subject=('Mail Test ' + (Get-Date).ToString());
                    Body=[string](Get-Content -Path ($mailBodyFileList | Get-Random));
                    ErrorAction='Stop';
                    To=($sendMailsToList -ne $mailFrom | Get-Random);
                }; # end mailSendParams
                $wpParams.Status=('Sending test mail ' + $cMc + '/' + $MailCountToSend + ' Sender: ' + $mailSendParams.From + ' Recipient: ' + $mailSendParams.To);
                $wpParams.PercentComplete=(($cMc/$MailCountToSend)*100);
                Write-Progress @wpParams;
                if (@($true,$false) | Get-Random)
                {
                    $mailSendParams.Add('Attachment',($mailAttachmentFileList | Get-Random));
                }; #end if
                try {
                    $msg=($mailSendParams.From + ' is sending mail to ' + $mailSendParams.To);
                    writeToLog -LogString $msg;
                    Send-MailMessage @mailSendParams;
                    Start-Sleep -Milliseconds $MailSendIntervallInMilliseconds;
                } # end try
                catch {
                    writeToLog -LogType Error -LogString  ('Failed to send mail to '+($mailSendParams.To));
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;
                }; # end catch 
            } # end if  
            else {
                writeToLog -LogString ('No SMTP server found.') -LogType Warning;
            }; # end else      
        }; # end for                
    }; # end process
    
    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
}; # end function Send-OPXTestMailMessages


function Set-OPXVirtualDirectories
{
<#
.SYNOPSIS 
Configures the virtual directores.
	
.DESCRIPTION
Configures the internal and external urls/hostnames and the internal Autodiscover uri. With customized configuration files, different options can be configured. For the OWA virtual directory the LogonFormat and the DefaultDomain could be configured. To create the custom configuration files the cmdlets *-OPXVirtualDirectoryTemplates can be used.
Not all parameter can be used together.
For the urls the most specific url will be used. If a DefaultGlobalNamespace and an MapiInternalUrl is configured, the value in parameter MapiInternalUrl will be used.
If none of the Default*Namespace parameter is used and only the parameter OABExternalUrl is used, only the OABExternalUrl will be configured.
For the urls is the FQDN for the namespace expected. The command creates the various urls.

.PARAMETER DefaultGlobalNamespace
The parameter is optional.
The FQDN for the namespace used for the most internal and external urls/hostnames.

.PARAMETER DefaultExternalNamespace
The parameter is optional.
The FQDN for the external namespaces. This parameter is only useful when the parameter DefaultGlobalNamespace is not used.

.PARAMETER DefaultInternalNamespace
The parameter is optional.
The FQDN for the internal namespaces. This parameter is only useful when the parameter DefaultGlobalNamespace is not used.

.PARAMETER OWAExternalUrl
The parameter is optional
The FQDN for the OWAExternalUrl

.PARAMETER OWAInternalUrl
The parameter is optional
The FQDN for the OWAInternalUrl

.PARAMETER OABExternalUrl
The parameter is optional
The FQDN for the OABExternalUrl

.PARAMETER OABInternalUrl
The parameter is optional
The FQDN for the OABInternalUrl

.PARAMETER WebservicesExternalUrl
The parameter is optional
The FQDN for the WebservicesExternalUrl

.PARAMETER WebservicesInternalUrl
The parameter is optional
The FQDN for the WebservicesInternalUrl

.PARAMETER ActiveSyncExternalUrl
The parameter is optional
The FQDN for the ActiveSyncExternalUrl

.PARAMETER ActiveSyncInternalUrl
The parameter is optional
The FQDN for the ActiveSyncInternalUrl

.PARAMETER MapiInternalUrl 
The parameter is optional
The FQDN for the MapiInternalUrl 

.PARAMETER MapiExternalUrl
The parameter is optional
The FQDN for the MapiExternalUrl

.PARAMETER MapiInternalUrl
The parameter is optional
The FQDN for the MapiInternalUrl

.PARAMETER OutlookAnywhereInternalHostname
The parameter is optional.
The FQDN for the OutlookAnywhereInternalHostname

.PARAMETER OutlookAnywhereExternalHostname
The parameter is optional.
The FQDN for the OutlookAnywhereExternalHostname

.PARAMETER OutlookAnywhereExternalClientAuthenticationMethode
The parameter is optional.
The authentication method for the external outlook client. Valid methods are Basic, Ntlm or Negotiate,
The parameter cannot be used with the parameter OutlookAnywhereDefaultAuthenticationMethode.

.PARAMETER OutlookAnywhereDefaultAuthenticationMethode
The parameter is optional.
The default authentication method for outlook client. Valid methods are Basic, Ntlm or Negotiate,
The parameter cannot be used with the parameter OutlookAnywhereExternalClientAuthenticationMethode.

.PARAMETER InternalAutodiscoverURI
The parameter is optional.
FQDN for the internal Autodiscover uri.

.PARAMETER ServerList
The parameter is optional.
The list of servers where the configuration should be applied. If this parameter and the parameter DAGName or AllCAS are not used, the name of the connected exchange server is used.
The parameter cannot be used with the parameters DAGName or AllCAS.

.PARAMETER DAGName
The parameter is optional.
Name of the DAG for which member the urls should be configured.
The parameter cannot be used with the parameters ServerList or AllCAS.
.PARAMETER AllCAS
The parameter is optional.
If the parameter is used, the virtual directory configuration is applied to all Exchange server with CAS installed.
The parameter cannot be used with the parameters ServerList or DAGName.

.PARAMETER LeaveExternalECPUrlBlank
The parameter is optional.
If the parameter is used, the external ECP url will be left blank.

.PARAMETER IncludeConfigurationFromConfigFile
The parameter is optional.
If a configuration file is used, the parameter expects the path to the configuration file.

.PARAMETER OnlyDisplayConfiguration
The parameter is optional.
If the parameter is used, no configuration will be applied. The configuration will be displayed. 

.PARAMETER ResolveVDirURLs 
The parameter is optional.
If the parameter is applied, the command tries to resolve the FQDNs of the urls.

.PARAMETER DomainController 
The parameter is optional.
The name of a domain controller which is used for the Exchange PowerShell cmdlets.

.EXAMPLE
Set-OPXVirtualDirectories -DefaultGlobalNamespace <namespace FQDN> -MapiInternalUrl <namespace FQDN> -OutlookAnywhereInternalHostname <Namespace FQND> -OnlyDisplayConfiguration -DomainController <domain.controller.fqdn>
#>
[cmdletbinding(DefaultParametersetName = 'Default')]
param([Parameter(Mandatory = $false, Position = 0)][string]$DefaultGlobalNamespace,    
    [Parameter(Mandatory = $false, Position = 1)][string]$DefaultExternalNamespace,
    [Parameter(Mandatory = $false, Position = 2)][string]$DefaultInternalNamespace,
    [Parameter(Mandatory = $false, Position = 3)][string]$OWAExternalUrl,
    [Parameter(Mandatory = $false, Position = 4)][string]$OWAInternalUrl,
    [Parameter(Mandatory = $false, Position = 5)][string]$OABExternalUrl,
    [Parameter(Mandatory = $false, Position = 6)][string]$OABInternalUrl,
    [Parameter(Mandatory = $false, Position = 7)][string]$WebservicesExternalUrl,
    [Parameter(Mandatory = $false, Position = 8)][string]$WebservicesInternalUrl,
    [Parameter(Mandatory = $false, Position = 9)][string]$ActiveSyncExternalUrl,
    [Parameter(Mandatory = $false, Position = 10)][string]$ActiveSyncInternalUrl,
    [Parameter(Mandatory = $false, Position = 12)][string]$MapiInternalUrl, 
    [Parameter(Mandatory = $false, Position = 11)][string]$MapiExternalUrl,
    [Parameter(Mandatory = $false, Position = 14)][string]$OutlookAnywhereInternalHostname,          
    [Parameter(Mandatory = $false, Position = 13)][string]$OutlookAnywhereExternalHostname,     
    [Parameter(Mandatory = $false, Position = 15)][ValidateSet('Basic','Ntlm','Negotiate')][string]$OutlookAnywhereExternalClientAuthenticationMethode,      
    [Parameter(Mandatory = $false, Position = 16)][ValidateSet('Basic','Ntlm','Negotiate')][string]$OutlookAnywhereDefaultAuthenticationMethode,      
    [Parameter(Mandatory = $false, Position = 17)][string]$InternalAutodiscoverURI,
    [Parameter(ParametersetName = "ServerList",Mandatory = $false, Position = 18)]
    [ArgumentCompleter( {  
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($false,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*"}); 
    } )]
    [array]$ServerList=$__OPX_ModuleData.ConnectedToMSXServer,
    [Parameter(ParametersetName = "DAG",Mandatory = $false, Position = 18)]
    [ArgumentCompleter( { 
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $dagList=($__OPX_ModuleData.getDagList($false,$null));
        $dagList.Where({ $_ -like "$wordToComplete*"}); 
    } )]
    [string]$DAGName,
    [Parameter(ParametersetName = "AllCas",Mandatory = $false, Position = 18)][switch]$AllCAS,
    [Parameter(Mandatory = $false, Position = 19)][switch]$LeaveExternalECPUrlBlank=$false,
    [Parameter(Mandatory = $false, Position = 20)][switch]$IncludeConfigurationFromConfigFile=$false,
    [Parameter(Mandatory = $false, Position = 21)][switch]$OnlyDisplayConfiguration = $false,        
    [Parameter(Mandatory = $false, Position = 22)][switch]$ResolveVDirURLs = $false,     
    [Parameter(Mandatory = $false, Position = 23)][string]$DomainController         
    )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
#region check parameters        
        $DCParam=@{
            ErrorAction='Stop';
        }; # end DCParam
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $DCParam.Add('DomainController',$DomainController);
        }; # end if
        if ($PSBoundParameters.ContainsKey('OutlookAnywhereDefaultAuthenticationMethode') -and $PSBoundParameters.ContainsKey('OutlookAnywhereExternalClientAuthenticationMethode'))
        {
            writeToLog -LogType Warning -LogString  'The parameters OutlookAnywhereDefaultAuthenticationMethode and OutlookAnywhereExternalClientAuthenticationMethode are mutual exclusive.';
            return;
        }; # end if
        $pList=@('DefaultExternalNamespace','DefaultGlobalNamespace','OutlookAnywhereExternalHostname');
        foreach ($entry in $pList)
        {
            if (($PSBoundParameters.ContainsKey($entry)) -and (! ($PSBoundParameters.ContainsKey('OutlookAnywhereDefaultAuthenticationMethode') -or $PSBoundParameters.ContainsKey('OutlookAnywhereExternalClientAuthenticationMethode'))))
            {
                writeToLog -LogType Warning -LogString  "To configure the external hostname for Outlook Anywehre either the parameter`n`tOutlookAnywhereExternalClientAuthenticationMethode or`n`tOutlookAnywhereDefaultAuthenticationMethode`nis required.";
                return;
            }; # end if
        }; # end foreach
#endregion check parameters
#region calulate namespces    
        Write-Verbose 'Calulating values for configuration fo virtual directories';
        writeTolog -LogString 'Calulating values for configuration fo virtual directories';
        $paramList=@((Get-Command ($MyInvocation.MyCommand.Name)).Parameters.Keys);
        $paramListVdirExt=($paramList.Where({$_ -Like '*ExternalUrl'})); 
        $paramListVdirInt=($paramList.Where({$_ -Like '*InternalUrl'}));
        $plvCount=$paramListVdirExt.count;
        $paramListHostName=($paramList.Where({$_ -Like '*ternalHostName'}));
        $paramListHostName+=($paramList.Where({$_ -Like '*AutodiscoverUri'})); 

        if ($PSBoundParameters.ContainsKey('IncludeConfigurationFromConfigFile'))
        {
            try {
                $ConfigurationFilePath=([System.IO.Path]::Combine($__OPX_ModuleData.cfgDir,'VDirs\' + $__OPX_ModuleData.VirtualDirCfgFileName ));
                $additionalConfig=Import-Clixml $ConfigurationFilePath;
                $allowedKeys=@('OWA','OAB','Webservices','ActiveSync','Mapi','ECP','OutlookAnywhere','Autodiscover');
                $keysToremove=@();
                foreach ($key in $additionalConfig.Keys)
                {
                    if ($key -notin $allowedKeys)
                    {
                        writeToLog -LogType Warning -LogString  ('Key ' + $key + ' not supported.');
                        $keysToremove+=$key;
                    }; # end if
                }; # end foreach check if keys are supported
                foreach ($item in $keysToremove)
                {
                    $additionalConfig.Remove($item);
                }; # end foreach remove unsupported keys
                
            } # end try
            catch {
                writeToLog -LogType Warning -LogString  ('Faild to import the configuration from file ' + $ConfigurationFilePath);
                writeTolog -LogString ($_.Exception.Message) -LogType Warning;
                $additionalConfig=@{};
            }; # end catch            
        } # end if
        else {
            $additionalConfig=@{};
        }; # end else

        if ($PSBoundParameters.ContainsKey('OutlookAnywhereExternalClientAuthenticationMethode'))
        {
            if (! ($additionalConfig.ContainsKey('OutlookAnywhere')))
            {
                $additionalConfig.Add('OutlookAnywhere',@{});
            }; # end if
            $additionalConfig.OutlookAnywhere.Add('ExternalClientAuthenticationMethod',$OutlookAnywhereExternalClientAuthenticationMethode);
            if (!($additionalConfig.OutlookAnywhere.ContainsKey('ExternalClientsRequireSsl')))
            {
                $additionalConfig.OutlookAnywhere.Add('ExternalClientsRequireSsl',$true);
            }; # end if
        }; # end if
        if ($PSBoundParameters.ContainsKey('OutlookAnywhereDefaultAuthenticationMethode'))
        {
            if (! ($additionalConfig.ContainsKey('OutlookAnywhere')))
            {
                $additionalConfig.Add('OutlookAnywhere',@{});
            }; # end if
            $additionalConfig.OutlookAnywhere.Add('ClientAuthenticationMethod',$OutlookAnywhereDefaultAuthenticationMethode);
            if (!($additionalConfig.OutlookAnywhere.ContainsKey('ExternalClientsRequireSsl')))
            {
                $additionalConfig.OutlookAnywhere.Add('ExternalClientsRequireSsl',$true);
            }; # end if
        }; # end if

        $vDirParams=[ordered]@{
            OWA=@{};
            ECP=@{};
            OAB=@{};
            Webservices=@{};
            ActiveSync=@{};
            Mapi=@{};
            OutlookAnywhere=@{};
            Autodiscover=@{};
        }; # end vDirParams
        
        for ($i=0; $i -lt $plvCount; $i++)
        {
            $keyVal=$paramListVdirExt[$i].Substring(0,($paramListVdirExt[$i].IndexOf('External')));
            if ($PSBoundParameters.ContainsKey($paramListVdirExt[$i]))
            {
                [void]$vDirParams.$keyVal.Add('ExternalUrl',(Get-Variable -Name ($paramListVdirExt[$i]) -ValueOnly));
            } # end if
            else
            {
                if ($PSBoundParameters.ContainsKey('DefaultExternalNamespace'))
                {
                    [void]$vDirParams.$keyVal.Add('ExternalUrl',$DefaultExternalNamespace);
                } # end if
                else
                {
                    if ($PSBoundParameters.ContainsKey('DefaultGlobalNamespace'))
                    {
                        [void]$vDirParams.$keyVal.Add('ExternalUrl',$DefaultGlobalNamespace);
                    };
                }; # end else
            }; # end else

            if ($PSBoundParameters.ContainsKey($paramListVdirInt[$i]))
            {
                [void]$vDirParams.$keyVal.Add('InternalUrl',(Get-Variable -Name ($paramListVdirInt[$i]) -ValueOnly));
            } # end if
            else
            {
                if ($PSBoundParameters.ContainsKey('DefaultInternalNamespace'))
                {
                    [void]$vDirParams.$keyVal.Add('InternalUrl',$DefaultInternalNamespace);
                } # end if
                else
                {
                    if ($PSBoundParameters.ContainsKey('DefaultGlobalNamespace'))
                    {
                        [void]$vDirParams.$keyVal.Add('InternalUrl',$DefaultGlobalNamespace);
                    }; # end if
                }; # end else
            }; # end else
        }; # end foreach

        if ($vDirParams.owa.count -gt 0)
        {
            foreach ($entry in $vDirParams.owa.Keys)
            {
                [void]$vDirParams.ECP.Add($entry,$vDirParams.owa.$entry);
            }; # end foreach
        }; # end if

        $keyVal='OutlookAnywhere'; 
        $indexPos=($paramListHostName.IndexOf('OutlookAnywhereInternalHostname'));
        if ($PSBoundParameters.ContainsKey($paramListHostName[$indexPos]))  
        {
            [void]$vDirParams.$keyVal.Add('InternalHostname',(Get-Variable -Name ($paramListHostName[$indexPos]) -ValueOnly));
        } # end if
        else
        {
            if ($PSBoundParameters.ContainsKey('DefaultExternalNamespace'))
            {
                [void]$vDirParams.$keyVal.Add('InternalHostname',$DefaultExternalNamespace);
            } # end if
            else
            {
                if ($PSBoundParameters.ContainsKey('DefaultGlobalNamespace'))
                {
                    [void]$vDirParams.$keyVal.Add('InternalHostname',$DefaultGlobalNamespace);
                }; # end if
            }; # end else
        }; # end else  
        
        $indexPos=($paramListHostName.IndexOf('OutlookAnywhereExternalHostname'));
        if ($PSBoundParameters.ContainsKey($paramListHostName[$indexPos]))  
        {
            [void]$vDirParams.$keyVal.Add('ExternalHostname',(Get-Variable -Name ($paramListHostName[$indexPos]) -ValueOnly));
        } # end if
        else
        {
            if ($PSBoundParameters.ContainsKey('DefaultExternalNamespace'))
            {
                [void]$vDirParams.$keyVal.Add('ExternalHostname',$DefaultExternalNamespace);
            } # end if
            else
            {
                if ($PSBoundParameters.ContainsKey('DefaultGlobalNamespace'))
                {
                    [void]$vDirParams.$keyVal.Add('ExternalHostname',$DefaultGlobalNamespace);
                };
            }; # end else
        }; # end else 
        
        $keyVal='Autodiscover';
        $indexPos=($paramListHostName.IndexOf('InternalAutodiscoverURI'));
        if ($PSBoundParameters.ContainsKey($paramListHostName[$indexPos]))  
        {
            [void]$vDirParams.$keyVal.Add('AutoDiscoverServiceInternalUri',(Get-Variable -Name ($paramListHostName[$indexPos]) -ValueOnly));
        } # end if
        else
        {
            if ($PSBoundParameters.ContainsKey('DefaultExternalNamespace'))
            {
                [void]$vDirParams.$keyVal.Add('AutoDiscoverServiceInternalUri',$DefaultExternalNamespace);
            } # end if
            else
            {
                if ($PSBoundParameters.ContainsKey('DefaultGlobalNamespace'))
                {
                    [void]$vDirParams.$keyVal.Add('AutoDiscoverServiceInternalUri',$DefaultGlobalNamespace);
                };
            }; # end else
        }; # end else
#endregion calculate namespaces
        foreach ($keyVal in $additionalConfig.Keys)
        {
            if ($vDirParams.Contains($keyVal))
            {
                [array]$itemList=$additionalConfig.$keyVal.keys;
                foreach ($item in $itemList)
                {
                    $vDirParams.$keyVal.Add($item,$additionalConfig.$keyVal.$item);
                }; # end foreach
            }; # end if
        }; # end foreach
        if ($LeaveExternalECPUrlBlank)
        {
            if ($vDirParams.ECP.Contains('ExternalURL'))
            {
                $vDirParams.ECP.ExternalURL=$NULL;
            } # end if
            else
            {
                [void]$vDirParams.ECP.Add('ExternalURL',$NULL);
            }; # end else
        } # end if
        if ($vDirParams.OutlookAnywhere.ContainsKey('InternalHostname'))
        {
            [void]$vDirParams.OutlookAnywhere.Add('InternalClientsRequireSsl',$true); # if interal hostname will be configured add InternalClientsRequireSsl
        }; # end if 

        switch ($PsCmdlet.ParameterSetName)
        {
            {$_ -eq 'DAG'}      {
                try {
                    $dagMembers=$__OPX_ModuleData.getDagList($true,$DAGName);
                    if ($dagMembers.count -eq 0)
                    {
                        writeToLog -LogString ('The DAG ' + $dagName + ' has no member server.') -LogType Warning;
                        return;
                    };
                    $serverList=@([System.Linq.Enumerable]::Intersect([string[]]$dagMembers.ToLower(),[string[]]$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower()));
                    break;
                } # end try
                catch {
                    writeToLog -LogType Error -LogString  ('Faild to enumerate the member servers of the DAG ' + $DAGName);
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;
                    return;
                }; # end catch
            }; # end DAG
            {$_ -eq 'AllCAS'}   {
                $serverList=$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower();
            }; # end AllCAS
        }; # end switch
        $serverList=@([System.Linq.Enumerable]::Distinct([string[]]$serverList));
        $slc=$serverList.count;
        for ($i=0;$i -lt $slc; $i++)  # test if Exchange servers are resoveable in DNS
        {
            if (!($serverList[$i] -eq (testIfFQDN -ComputerName ($serverList[$i]))))
            {
                return;
            }; # end if
        }; # end for
        if ($ResolveVDirURLs)
        {
            Write-Verbose ('Verify if URLs and hostnames are resolveable');
            writeTolog -LogString ('Verify if URLs and hostnames are resolveable');
            resolveURLs -ParameterHash $vDirParams;
        }; # end if
#region format URLs
    # configure URLs
    foreach($service in $vDirParams.keys)
    {
        if ($vDirParams.$service.count -gt 0)
        {
            $pl=$vDirParams.$service;
            switch ($service)
            {
                {$_ -in ('OWA','ECP','OAB','Mapi')} {
                    $ProtocolString=$_;                    
                    formatURL -URLHash $pl -ProtocolString $ProtocolString;
                    break;
                };
                'WebServices'       {
                    $ProtocolString='EWS/Exchange.asmx';
                    formatURL -URLHash $pl -ProtocolString $ProtocolString; 
                    break;
                }; # end WebServices
                'ActiveSync'        {
                    $ProtocolString='Microsoft-Server-ActiveSync';
                    formatURL -URLHash $pl -ProtocolString $ProtocolString;
                    break;
                }; # end WebServices
                'OutlookAnywhere'   {
                    [void]$pl.Add('ErrorAction','Stop'); 
                }; # end Autodiscover
                'Autodiscover'      {
                    $pl.AutoDiscoverServiceInternalUri=('https://' + $vDirParams.Autodiscover.AutoDiscoverServiceInternalUri + '/autodiscover/Autodiscover.xml');
                    [void]$pl.Add('ErrorAction','Stop'); 
                }; #end if
            }; # end switch
        }; # end count of params -gt 0
    }; # end foreach configure URLs
#endregion format URLs
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            foreach($item in $vDirParams.Keys)
            {
                if ($vDirParams.$item.count -gt 0)
                {
                    $vDirParams.$item.Add('DomainController',$DomainController);
                }; # endif
            }; # end foreach
        }; # end if
        if ($OnlyDisplayConfiguration.IsPresent)
        {
            $outTbl=createNewTable -TableName 'Output' -FieldList @(@('Service',[system.string]),@('Parameter',[system.string]),@('Parameter Value',[system.string])); #,@('Additional configuration',[system.string]));        
            foreach ($vDir in $vDirParams.Keys)
            {
                if ($vDirParams.$vDir.count -gt 0)
                {
                    $dirName=$vdir;
                    foreach ($param in $vDirParams.$vDir.keys)
                    {
                        if ($param -ne 'ErrorAction')
                        {
                            [void]$outTbl.Rows.Add($dirName,$param,$vDirParams.$vDir.$param);
                            $dirName=$null;
                        }; # end if
                    }; # end foreach                                
                }; # end if count of vDir entries gt 0            
            }; # end foreach        
            $outTbl | Format-Table -AutoSize;
            Write-Output 'The configuration will be implemented on the following server(s):';
            $serverList;
        } # end if
        else # configure list of servers
        {                       
            foreach ($server in $ServerList) # configure vDirs on Exchange servers
            {            
                writeToLog -LogString ('Configuring virtual directories/hostnames on server ' + $server) -logType Info -ShowInfo;
                if (($server=testIfFQDN -ComputerName $server) -eq $false) # is the server up and running
                {
                    continue;
                };
                if ($null -eq ($Host.UI.RawUI.WindowSize))
                {
                    $windowWidth=120;
                } # end if
                else
                {
                    $windowWidth=$Host.UI.RawUI.WindowSize.Width;
                }; # end if
                $strLength=(67 + ($server.length));            
                $xTimes=([math]::Floor(($windowWidth-$strLength)/2));
                Write-Verbose ('#'*$xTimes +' Starting configuration of virtual directories on server ' + $server + ' ' + ('#'*$xTimes));
                writeTolog -LogString ('Starting configuration of virtual directories on server ' + $server);
                
                switch ($true)
                {
                    {$vDirParams.owa.count -gt 0} {
                        Write-Verbose ('Configuring OWA virtual directory urls on server ' + $server);
                        writeTolog -LogString ('Configuring OWA virtual directory urls on server ' + $server);
                        $pl=$vDirParams.OWA;                
                        try {
                            Set-OwaVirtualDirectory -Identity ($server.Split('.')[0] + '\owa (Default Web Site)') @pl;
                        } # end try
                        catch {
                            writeToLog -LogType Error -LogString  ('Failed to configure the OWA virtual directorys on server ' + $server);
                            writeToLog -LogString ($_.Exception.Message) -LogType Error;
                        }; # end catch                    
                    }; # end OWA

                    {$vDirParams.ECP.count -gt 0} {
                        $pl=$vDirParams.ECP;
                        Write-Verbose ('Configuring ECP virtual directory urls on server ' + $server);                   
                        writeTolog -LogString ('Configuring ECP virtual directory urls on server ' + $server);                   
                        try {
                        Set-EcpVirtualDirectory -Identity ($server.Split('.')[0] + '\ecp (Default Web Site)') @pl;
                        } # end try
                        catch {
                            writeToLog -LogType Error -LogString  ('Failed to configure the ECP virtual directorys on server ' + $server);
                            writeToLog -LogString ($_.Exception.Message) -LogType Error;
                        }; # end catch   
                    }; # ECP
                    {$vDirParams.OAB.count -gt 0} {
                        $pl=$vDirParams.OAB;
                        Write-Verbose ('Configuring OAB virtual directory urls on server ' + $server);
                        writeTolog -LogString ('Configuring OAB virtual directory urls on server ' + $server);
                        try {
                            Set-OabVirtualDirectory -Identity ($server.Split('.')[0] + '\OAB (Default Web Site)') @pl;
                        } # end try
                        catch {
                            writeToLog -LogType Error -LogString  ('Failed to configure the OAB virtual directorys on server ' + $server);
                            writeToLog -LogString ($_.Exception.Message) -LogType Error;
                        }; # end catch                    
                    }; # end OAB

                    {$vDirParams.Webservices.count -gt 0} {
                        $pl=$vDirParams.Webservices;
                        Write-Verbose ('Configuring Webservices virtual directory urls on server ' + $server);                    
                        writeTolog -LogString ('Configuring Webservices virtual directory urls on server ' + $server);                    
                        try {
                            Set-WebServicesVirtualDirectory -Identity ($server.Split('.')[0] + '\EWS (Default Web Site)') @pl;
                        } # end try
                        catch {
                            writeToLog -LogType Error -LogString  ('Failed to configure the EWS virtual directorys on server ' + $server);
                            writeToLog -LogString ($_.Exception.Message) -LogType Error;
                        }; # end catch                    
                    }; # end Webservices

                    {$vDirParams.ActiveSync.count -gt 0} {
                        $pl=$vDirParams.ActiveSync;
                        Write-Verbose ('Configuring ActiveSync virtual directory urls on server ' + $server);                    
                        writeTolog -LogString ('Configuring ActiveSync virtual directory urls on server ' + $server);                    
                        try {
                            Set-ActiveSyncVirtualDirectory -Identity ($server.Split('.')[0] + '\Microsoft-Server-ActiveSync (Default Web Site)') @pl;
                        } # end try                   
                        catch {
                            writeToLog -LogType Error -LogString  ('Failed to configure the ActiveSync virtual directorys on server ' + $server);
                            writeToLog -LogString ($_.Exception.Message) -LogType Error;
                        }; # end catch                    
                    }; # end ActiveSync

                    {$vDirParams.Mapi.count -gt 0} {
                        $pl=$vDirParams.Mapi;
                        Write-Verbose ('Configuring Mapi virtual directory urls on server ' + $server);                   
                        writeTolog -LogString ('Configuring Mapi virtual directory urls on server ' + $server);                   
                        try {
                        Set-MapiVirtualDirectory -Identity ($server.Split('.')[0] + '\mapi (Default Web Site)') @pl;
                        } # end try
                        catch {
                            writeToLog -LogType Error -LogString  ('Failed to configure the Mapi virtual directorys on server ' + $server);
                            writeToLog -LogString ($_.Exception.Message) -LogType Error;
                        }; # end catch                    
                    }; # end Mapi

                    {$vDirParams.OutlookAnywhere.count -gt 0} {
                        $pl=$vDirParams.OutlookAnywhere;
                        Write-Verbose ('Configuring Outlook Anywhere hostnames on server ' + $server);
                        writeTolog -LogString ('Configuring Outlook Anywhere hostnames on server ' + $server);
                        try {
                            Set-OutlookAnywhere -Identity ($server.Split('.')[0] + '\Rpc (Default Web Site)') @pl;
                        } # end try
                        catch {
                            writeToLog -LogType Error -LogString  ('Failed to configure the Outlook Anywehre hostnames on server ' + $server);
                            writeToLog -LogString ($_.Exception.Message) -LogType Error;
                        }; # end catch                    
                    }; # end OutlookAnywhere

                    {$vDirParams.Autodiscover.count -gt 0} {
                        Write-Verbose ('Configuring the internal Autodiscover URI  on server ' + $server);
                        writeTolog -LogString ('Configuring the internal Autodiscover URI  on server ' + $server);                        
                        $pl=$vDirParams.Autodiscover;
                        try {
                            Set-ClientAccessService -Identity $server @pl;
                        } # end try
                        catch {
                            writeToLog -LogType Error -LogString  ('Failed to configure the internal Autodiscover URI on server ' + $server);
                            writeToLog -LogString ($_.Exception.Message) -LogType Error;
                        }; # end catch
                    }; # end Autodiscover
                }; # end switch            
            }; # end foreach
        }; # end else
    }; # end process
    
    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
}; # end function Set-OPXVirtualDirectories

function Get-OPXLogFileEntries
{
<#
.SYNOPSIS 
Display content of log files.

.DESCRIPTION
The module MSX be configured to log info, warnings and errors to a log file. Every day a new log file will be created. The log files are stored in the directory <path of module MSX>\Logs. 

.PARAMETER LogfilePath
The parameter is optional
The path to the log file which should be displayed. If the parameter is omitted, the current log file will be displayed.

.PARAMETER IncludeInfo
The parameter is optional
If you call the cmdlet, in most cases, only warning and errors are displayed. To include info entries, use the parameter IncludeInfo.

.PARAMETER IncludeUserName
The parameter is optional
If the parameter is used, the user name is included.

.PARAMETER Format
The parameter is optional
If the parameter is omitted, the log is displayed as table. Additional options are:
	Lise
	PassValue (unformatted output from a CSV import)

.PARAMETER GetLoggingFromLastCmdletCall
The parameter is optional
Displays the log entries from the last command call. Info entries are included. In the header of the output the command name and the CallingID are displayed. With the CallingID the log entries of a command call can be accessed and displayed again.

.PARAMETER CallingID
The parameter is optional
For every command call a CallingID is generated. With this parameter you can query the log file for entries with a given CallingID. Info entries are included.

.PARAMETER CmdletName
The parameter is optional
To get log entries from an MSX command call, this parameter can be used. If an MSX command calls another MSX cmdlet, these entries are not included. You can step through the list of cmdlets with TAB or get a list of cmdlets with Ctr-Space.

.PARAMETER ListOnlyCmdlets
The parameter is optional
Only lists the names of the MSX cmdlets found in the log.

.EXAMPLE
Get-OPXLogFileEntries -GetLoggingFromLastCmdletCall
#>
[cmdletbinding(DefaultParametersetName='Default')]
param([Parameter(Mandatory = $false, Position = 0)]      
      [ArgumentCompleter( {             
        $fileList=(Get-ChildItem -Path ([System.IO.Path]::Combine($__OPX_ModuleData.logDir,(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Replace('\','_')))))       
        $fileList = ($fileList.Where({$_.name -like 'Log_*.log'}));
        foreach ($item in $fileList) {        
            $path = $item.FullName
            if ($path -like '* *') 
            { 
                $path = "'$path'"
            }; # end if
            [Management.Automation.CompletionResult]::new($path, [System.IO.Path]::GetFileNameWithoutExtension($path), 'ParameterValue', [System.IO.Path]::GetFileNameWithoutExtension($path));                 
        }; # end foreach       
      } )]      
      [string]$LogfilePath,
      [Parameter(Mandatory = $false, Position = 1)][switch]$IncludeInfo=$false,
      [Parameter(Mandatory = $false, Position = 2)][switch]$IncludeUserName=$false,
      [Parameter(Mandatory = $false, Position = 3)][OutFormat]$Format='Table',
      [Parameter(ParameterSetName='LastCall')]
      [Parameter(Mandatory = $false, Position = 5)][switch]$GetLoggingFromLastCmdletCall=$false,
      [Parameter(ParameterSetName='IdAndName')]
      [Parameter(Mandatory = $false, Position = 6)][string]$CallingID,
      [Parameter(ParameterSetName='IdAndName')]
      [Parameter(Mandatory = $false, Position = 7)]
          [ArgumentCompleter( { 
            param ( $CommandName,
                $ParameterName,
                $WordToComplete,
                $CommandAst,
                $FakeBoundParameters )  
            $cmdList=@($__OPX_ModuleData.cmdletList);
            $cmdList.Where({ $_ -like "$wordToComplete*"}); 
          } )]
          [string]$CmdletName,
      [Parameter(ParameterSetName='ListCmdlet')]
      [Parameter(Mandatory = $false, Position = 8)][switch]$ListOnlyCmdlets=$false
     )

    begin {
        $prevCallingID=$script:CallingID; # get CallingID  from last command call
        $prevLogFilePath=$script:lastLogFileName; # get the name of the file where the last log string was written to
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());        
    }; # end begin
    process {
        if ((!($PSBoundParameters.ContainsKey('LogfilePath'))) -or ($PSBoundParameters.ContainsKey('GetLoggingFromLastCmdletCall')))
        {            
            $LogfilePath=$prevLogFilePath;
        }; # end if       
        try {
            $logs=Import-Csv -Path $LogfilePath -Delimiter ($Script:LogFileCsvDelimiter);
            if ($PSBoundParameters.ContainsKey('ListOnlyCmdlets'))
            {
                $cmdletList=[System.Collections.ArrayList]::new();
                $cmdletList.AddRange($logs.CallingCmdlet);
                $cmdletList=@([System.Linq.Enumerable]::Distinct([string[]]$cmdletList));
                $cmdletList;
                return;
            }; # end if
            $fieldList=[System.Collections.ArrayList]::new();
            [void]('DateTime','UtcOff' |  ForEach-Object {$fieldList.Add($_)});            
            if ($PSBoundParameters.ContainsKey('IncludeUserName'))
            {
                $fieldList+='UserName';
                $PSBoundParameters.ContainsKey('GetLoggingFromLastCmdletCall') 
                $PSBoundParameters.ContainsKey('CallingID')
            }; # end if
                 
            $QueryFilter=@();
            [void]('LogType','CallingCmdlet','LogMessage' | ForEach-Object {$fieldList.Add($_)});  # create list field list 
            
            if ((!($PSBoundParameters.ContainsKey('GetLoggingFromLastCmdletCall'))) -and  (!($PSBoundParameters.ContainsKey('CallingID'))) -and (!($PSBoundParameters.ContainsKey('IncludeInfo'))))
            {
                $QueryFilter+='($_.logtype -in @("Warning","Error"))';
            }; # end if
            
            switch ($PSBoundParameters)
            {                                
                {$_.ContainsKey('GetLoggingFromLastCmdletCall')} {                        
                    try {
                        $fieldList.Remove('CallingCmdlet');              
                        $QueryFilter+='($_.CallingID -eq $prevCallingID)';
                        Write-Output ('Calling ID:  ' + $prevCallingID);
                        Write-Output ('Cmdlet name: ' + $logs.where({$_.CallingID -eq $prevCallingID})[0].CallingCmdlet);
                    } # end try
                    catch {
                        $qStr=(getQueryString -QueryFilter $QueryFilter);
                        writeToLog -LogString ('Failed to find entries in log which match the filter "' + $qStr + '" in log file ' + $LogfilePath) -LogType Warning;
                        if ($ShowDebuggingInfo)
                        {
                            writeToLog -LogString ($_.Exception.Message) -LogType Error;
                        } # end if
                        else {
                            writeToLog -LogString ($_.Exception.Message);
                        }; # end else                        
                        return;
                    }; # end catch                     
                    break;           
                }; # end 'GetLoggingFromLastCmdletCall'
                {$_.ContainsKey('CallingID')}       {
                    $QueryFilter+='($_.CallingID -eq $CallingID)';
                }; # end CallingID
                {$_.ContainsKey('CmdletName')}       {
                    $QueryFilter+='($_.CallingCmdlet -like $cmdletName)';
                }; # end CallingID
            }; # end swith
            
            if ($QueryFilter.count -gt 0)
            {
                try {
                    $logs=$logs.Where((([System.Management.Automation.ScriptBlock]::Create($QueryFilter -join ' -and '))));
                } # end try
                catch {
                    $qStr=(getQueryString -QueryFilter $QueryFilter);
                    writeToLog -LogString ('Failed to find entries in log which match the filter "' + $qStr + '" in log file ' + $LogfilePath) -LogType Warning;
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;
                    return;
                }; # end switch                
            }; # end if
            
            switch ($Format)
            {
                'Table' {
                    $logs | Format-Table $fieldList;
                    break;
                }; # format tabel
                'List'  {
                    $logs | Format-List $fieldList;
                    break;
                }; # format list
                'PassValue'  {
                    $logs;
                    break;
                }; # format list                                        
            }; # end switch                     
        } # end try
        catch {
            writeTolog -LogString ('Failed to read the log file ' + $LogfilePath) -LogType Error;
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch        
    }; # end process
    
    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
        $script:lastLogFileName= $script:CurrentLogFileName;
    }; # end END
}; # end function Get-OPXLogFileEntries      

function Restart-OPXExchangeService
{
<#
.SYNOPSIS 
Restarts a service.

.DESCRIPTION
With the command an Exchange service can be restarted. With TAB, the service can be selected. The list of services is configurable. The names of the services can be configured in the file <module root directory>\cfg\Services\cfgFile\ServiceRestart.csv>. Per default the Exchange Information Store and the Mailbox Replication Service are included.

.PARAMETER Service, 
The parameter is mandatory.
The name of the service to restart. The service can be selected with TAB.

.PARAMETER ServerList
The parameter is optional
The name(s) of computer(s). The names of the computers can be selected with TAB.

.PARAMETER AllCAS,
The parameter is optional
The service will be restarted on all Exchange servers with CAS.

.PARAMETER DAGName
The parameter is optional
The service will be restarted on all DAG member server.

.EXAMPLE
Restart-OPXExchangeService -Service MSExchangeMailboxReplication -DAGName <name of DAG>
#>
[cmdletbinding(DefaultParametersetName = 'ServerList')]
param([Parameter(Mandatory = $true, Position = 0)]      
      [ArgumentCompleter( {                        
          param ( $CommandName,
          $ParameterName,
          $WordToComplete,
          $CommandAst,
          $FakeBoundParameters )  
          $svcList=(Get-Content -Path $([System.IO.Path]::Combine($__OPX_ModuleData.cfgDir,'Services\cfgFiles\ServiceRestart.csv')));
          $svcList.Where({ $_ -like "$wordToComplete*"});
      } )]      
      [string]$Service, 
      [Parameter(ParametersetName='ServerList',Mandatory = $false, Position = 1)]
      [ArgumentCompleter( {  
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($true,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*"}); 
      } )]
      [array]$ServerList=$__OPX_ModuleData.ConnectedToMSXServer,
      [Parameter(ParametersetName='AllCAS',Mandatory = $false, Position = 1)][switch]$AllCAS,      
      [Parameter(ParametersetName='DAG',Mandatory = $false, Position = 1)]
      [ArgumentCompleter( {  
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $dagList=($__OPX_ModuleData.getDagList($false,$null));
        $dagList.Where({ $_ -like "$wordToComplete*"}); 
      } )]
      [string]$DAGName
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {        
        switch ($PSCmdlet.ParameterSetName)
        {
            'AllCAS'    {
                if (($ServerList=getAllCAS) -eq $false)
                {
                    return;
                }; # end if
            }; # end AllCAS
            'DAG'       {                
                if (($ServerList =getDAGMemberServer -DAGName $DAGName)[0] -eq $false)
                {
                    return;
                }; # end if               
            }; # end DAG
        }; # end switch    
        
        $serverList=@([System.Linq.Enumerable]::Distinct([string[]]$serverList));
        $srvCount=$serverList.Count;
        for ($i=0;$i -lt $srvCount;$i++)
        {
            if ($srvCount -gt 1)
            {
                Write-Progress ('Restarting servcie ' + $service + ' on server ' + $serverList[$i]) -Status ([Math]::Round(($i/$srvCount)*100,2).ToString() + '% Complete:') -PercentComplete (($i/$srvCount)*100);
            }; # end if
            if (($serverList[$i]=testIfFQDN -ComputerName $ServerList[$i]) -eq $false)
            {
                continue;
            };
            $msg=('Restarting service ' + $Service + ' on Server ' + $serverList[$i]);
            Write-Verbose $msg;
            writeTolog -LogString $msg;
            try
            {
                Get-Service -Name $Service -ComputerName $serverList[$i] | Restart-Service
            } # end try
            catch
            {
                $msg=('Failed to restart the service ' + $service + ' on computer ' + $serverList[$i]);
                writeToLog -LogType Error -LogString  $msg;
                writeToLog -LogString ($_.Exception.Message) -LogType Error;
            }; # end catch
        } # end foreach
        if ($Service-eq 'MSExchangeIS')
        {
            Write-Output '';
            writeToLog -LogType Warning -LogString  'If the server(s) is/are a DAG member(s) and the databases are not automatically redistributed, please run the command Start-OPXMailboxDatabaseRedistribution to redistribute the databases.';
        }; # end if        
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
}; # end function Restart-OPXExchangeService

function Get-OPXExchangeCertificate
{
<#
.SYNOPSIS 
Displays Exchange certificates.

.DESCRIPTION
Displays all Exchange certificate, or a Exchange certificate with a given thumbprint, to verify the thumbprints, start date, end date and subject name.

.PARAMETER ServerList
The parameter is optional
List of Exchange servers. With TAB the servers can be selected.
If the parameter is omitted, the name of the connected exchange server is assumed.

.PARAMETER AllCAS
The parameter is optional
Displays (a) certificate(s) from all Exchange server with CAS.

.PARAMETER DAGName
The parameter is optional
Displays (a) certificate(s) from all Exchange which are member of the DAG.

.PARAMETER Thumbprint
The parameter is optional
The thumbprint of the certificate to display.

.PARAMETER DomainController
The parameter is optional
The name of a domain controller which is used for the Exchange PowerShell cmdlets.

.PARAMETER Format
The parameter is optional, default is Table
Valid options are Table, List and PassValue. PassValue returns the unformatted raw data.

.EXAMPLE
Get-OPXExchangeCertificate -Thumbprint <certificate thumbprint> -AllCAS
#>
[cmdletbinding(DefaultParametersetName = 'ServerList')]
param([Parameter(Mandatory = $false, Position = 0)][string]$Thumbprint,
      [Parameter(ParametersetName='ServerList',Mandatory = $false, Position = 1)]
      [ArgumentCompleter( {  
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($false,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*"});                          
      } )]
      [array]$ServerList=$__OPX_ModuleData.ConnectedToMSXServer,
      [Parameter(ParametersetName='AllCAS',Mandatory = $false, Position = 1)][switch]$AllCAS=$false,
      [Parameter(ParametersetName='DAG',Mandatory = $true, Position = 1)]
      [ArgumentCompleter( { 
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $dagList=($__OPX_ModuleData.getDagList($false,$null));
        $dagList.Where({ $_ -like "$wordToComplete*"});                                
      } )]
      [string]$DAGName,      
      [Parameter(Mandatory = $false, Position = 3)][string]$DomainController,
      [Parameter(Mandatory = $false, Position = 4)][OutFormat]$Format='Table'
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {        
        switch ($PsCmdlet.ParameterSetName)
        {
            {$_ -eq 'DAG'}      {
                try {
                    Write-Verbose ('Building list of Exchange servers in DAG ' + $DAGName) ; 
                    writeTolog -LogString ('Building list of Exchange servers in DAG ' + $dagName) ; 
                    $dagMembers=$__OPX_ModuleData.getDagList($true,$DAGName);
                    if ($dagMembers.count -eq 0)
                    {
                        writeToLog -LogString ('The DAG ' + $dagName + ' has no member server.') -LogType Warning;
                        return;
                    };
                    $serverList=@([System.Linq.Enumerable]::Intersect([string[]]$dagMembers.ToLower(),[string[]]$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower()));
                    break;
                } # end try
                catch {
                    writeToLog -LogType Error -LogString  ('Faild to enumerate the member servers of the DAG ' + $DAGName);
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;
                    return;
                }; # end catch
            }; # end DAG
            {$_ -eq 'AllCAS'}   {
                try
                {
                    Write-Verbose 'Building list of Exchange servers with Client Access service';
                    writeTolog -LogString 'Building list of Exchange servers with Client Access service';
                    $serverList=@($__OPX_ModuleData.getExchangeServerList($false,$true,$true).toLower());
                } # end try
                catch
                {
                    writeToLog -LogType Error -LogString  ('Failed to list Exchange servers witch Client Access service.');
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;    
                    return;
                }; # end catch
            }; # end AllCAS
        }; # end switch        
        $serverList=@([System.Linq.Enumerable]::Distinct([string[]]$serverList));
        foreach ($server in $ServerList)
        {
            try
            {
                $paramList=@{
                    Server=$server;
                    ErrorAction='Stop';
                }; # end paramlist
                if ($PSBoundParameters.ContainsKey('Thumbprint'))
                {
                    $msg=('Searching for certificate with thumbprint ' + $Thumbprint + ' on server ' + $Server + '.');
                    $paramList.Add('Thumbprint',$Thumbprint)
                } # end if
                else {
                    $msg=('Searching for certificates on server ' + $Server + '.');
                }; # end else
                Write-Verbose $msg;
                writeTolog -LogString $msg;
                If ($PSBoundParameters.ContainsKey('DomainController'))
                {
                    $paramList.Add('DomainController',$DomainController);
                }; # end if
                $msxCert=Get-ExchangeCertificate @paramList;
                $fieldList=[System.Collections.ArrayList]::new();
                [void]$fieldList.AddRange(@('Thumbprint','NotBefore','NotAfter'));
                if (Get-Member -InputObject $msxCert[0] -Name 'Services' -ErrorAction SilentlyContinue)
                {
                    [void]$fieldList.Add('Services');
                }; # end if
                [void]$fieldList.AddRange(@('Subject','DnsNameList'));
                $padCount=14;
                switch ($Format)
                {
                    'Table' {
                        $msxCert | Format-Table $fieldList;
                        break;
                    }; # format tabel
                    'List'  {                       
                        'Server Name'.PadRight($padCount,' ') + ': '+ $server;
                        foreach($cert in $msxCert)
                        {
                            for ($i=0;$i -lt ($FieldList.Count -1);$i++)
                            {
                                $fieldList[$i].PadRight($padCount,' ') + ': ' + $cert.($FieldList[$i]);
                            }; # end for
                            
                            $domList=$cert.($FieldList[$i])
                            $fieldList[$i].PadRight($padCount,' ') + ': '  + $domList[0];
                            for ($i=1;$i -lt $domList.count;$i++)
                            {
                                (''.PadRight(($padCount+2),' ') +  $domList[$i]);
                            }; # end for
                            Write-Output '';
                        }; # end foreach
                        break;
                    }; # format list
                    'PassValue'  {
                        $msxCert;
                        break;
                    }; # format list                                        
                }; # end switch         
                
            } # end try
            catch
            {
                writeToLog -LogType Error -LogString  ('Failed to find the certificate with the thumbprint ' + $Thumbprint + ' on server ' + $server + '.');
                writeToLog -LogString ($_.Exception.Message) -LogType Error;
            }; # end catch
        }; # end foreach
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
}; # end function Get-OPXExchangeCertificate


#region experimental cmdlets
function Get-OPXExchangeServerInMaintenance
{
<#
.SYNOPSIS 
Lists all Exchange server which are in maintenance.

.DESCRIPTION
Displays all or an Exchange certificate with a given thumbprint.
The command returns a list of Exchange server which seems to be in maintenance mode. Exchange server, where some components are not active are listed too. The default value for inactive components is 2. The value is configurable (<module root directory>\cfg\Constans.ps1, ServerComponentNotExpectedActive).

.PARAMETER NumberOfInactiveComponents
The number of components expected to be inactive.

.EXAMPLE
Get-OPXExchangeServerInMaintenance
#>
[cmdletbinding()]
param([Parameter(Mandatory = $false, Position = 0)][int]$NumberOfInactiveComponents=$script:ServerComponentNotExpectedActive
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        try {
            $msg='Searching for Exchange servers'
            writeTolog -LogString $msg;
            Write-Verbose $msg;
            ###$serverList=Get-ExchangeServer -ErrorAction Stop;
            $serverList=getAllExchangeServer;
        } # end try
        catch {
            $msg='Failed to search for Exchange server.'
            writeToLog -LogString $msg -logType Error;
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
            return;
        }; # end catch
        
        $serverCount=$serverList.count;
        $tblFields=@(
            @('Server in Maintenance',[system.string]),
            @('Full Maint.',[System.String]),
            @('Act.Comp.',[system.string]),
            @('DAG Member',[System.String]),
            @('DBCopyAutoActivPolicy',[System.String]),
            @('DBCopyActivDisabledAndMoveNow',[System.String]),
            @('Cluster',[System.String])
        ); # end tblFields
        
        $serverTable=createNewTable -TableName 'Server' -FieldList $tblFields;
        for ($i=0; $i -lt $serverCount;$i++)
        {
            ###$srvName=(testIfFQDN -ComputerName ([string]$serverList[$i].identity));
            $srvName=(testIfFQDN -ComputerName ([string]$serverList[$i]));
            $pc=([math]::Round((($i/$serverCount)*100),2));            
            $wpParams=@{
                Activity=('Verifying if server ' + $srvName + ' is in maintenance');
                Status=('Server ' + ($i+1) + ' of ' + $serverCount);
                PercentComplete=$pc;
            }; # end if
            Write-Progress @wpParams;        
            try {
                $msg=('Verifiying if server ' + $srvName + ' is in maintenance.')
                writeTolog -LogString $msg;
                Write-Verbose $msg;
                $componentsOnServer=(Get-ServerComponentState -Identity $srvName -ErrorAction Stop);
                $activeComponents=(($componentsOnServer).state).Where({$_ -eq 'active'}).count;
                
                $dagInfo=@{
                    Name='N/A';
                    DatabaseCopyAutoActivationPolicy='N/A';
                    DatabaseCopyActivationDisabledAndMoveNow='N/A';
                    ClusterNode='N/A';
                }; # end dagInfo
                
                $componetesOk=7;
                If($srvObj=Get-MailboxServer -Identity $srvName -ErrorAction SilentlyContinue)
                {
                    if (!([system.String]::IsNullOrEmpty(([string]$srvObj.DatabaseAvailabilityGroup)))) # check if server is DAG member
                    {
                        $dagInfo.Name=[string]$srvObj.DatabaseAvailabilityGroup; # get name of DAG
                        # collect current confiuration settings
                        $dagInfo.DatabaseCopyAutoActivationPolicy=[string]$srvObj.DatabaseCopyAutoActivationPolicy; 
                        $dagInfo.DatabaseCopyActivationDisabledAndMoveNow=[string]$srvObj.DatabaseCopyActivationDisabledAndMoveNow;
                        try {
                            # get cluster status
                            $dagInfo.ClusterNode=(Invoke-Command -ComputerName ($__OPX_ModuleData.ConnectedToMSXServer) -ScriptBlock {param($p1) (get-ClusterNode $p1).state} -ArgumentList $srvName).value;                    
                        } # end try
                        catch {
                            $dagInfo.ClusterNode='Failed to query';
                            $msg=('Failed to query the cluster status on server ' + $srvName);
                            writeToLog -LogString $msg -logType Error;
                            writeToLog -LogString ($_.Exception.Message) -LogType Error;
                        }; # end catch
                        # check components
                        $componetesOk=($componetesOk -bxor (1 * [int]($dagInfo.DatabaseCopyAutoActivationPolicy -ne 'Unrestricted')));
                        $componetesOk=($componetesOk -bxor (2 * [int]($dagInfo.DatabaseCopyActivationDisabledAndMoveNow -eq $true)));
                        $componetesOk=($componetesOk -bxor (4 * [int]($dagInfo.ClusterNode -ne 'UP')));                        
                    }; # end if
                }; # end if
                
                if (($activeComponents -lt (($componentsOnServer.count) -$NumberOfInactiveComponents)) -or ($componetesOk -lt 7))
                {
                    $padVal=(($serverTable.Columns[2].Caption.Length)-((([string]$componentsOnServer.count).Length) + 1));
                    $fullMaintenance=(($activeComponents -le $NumberOfInactiveComponents) -and ($componetesOk -eq 0));
                    $tblFieldsToadd=@(
                        $srvName,
                        $fullMaintenance,
                        (([string]$activeComponents).PadLeft($padVal,' ')+'/'+[string](($componentsOnServer.count) -$NumberOfInactiveComponents)),
                        $dagInfo.Name,
                        $dagInfo.DatabaseCopyAutoActivationPolicy,
                        $dagInfo.DatabaseCopyActivationDisabledAndMoveNow,
                        $dagInfo.ClusterNode
                    ); # end tblFieldsToAdd
                    
                    [void]($serverTable.rows.Add($tblFieldsToadd));
                    if ($fullMaintenance)
                    {
                        $msg=('!!! ' + $srvName + ' is in maintenance !!!');
                        writeToLog -LogString $msg;
                    } # end if
                    else {
                        $msg=($srvName + ' seams to be partialy in maintenance. Some components are not in the expected state.');
                        writeTolog -LogString $msg -logType Warning;
                    }; # end else                    
                }; # end if
            } # end try
            catch {                
                $msg=('Failed to determine if server ' + $srvName + ' is in maintenance.');
                writeToLog -LogString $msg -logType Error;
                writeToLog -LogString ($_.Exception.Message) -LogType Error;
            }; # end catch        
        }; # end for

        if ($serverTable.rows.count -gt 0)
        {
            $serverTable.Select('', '[DAG Member] ASC') | Format-Table;                
        } # end if
        else {
            $msg='No server is in maintenance.';
            writeToLog -LogString $msg;
            Write-Verbose $msg;
        }; # end if
    }; # end process
    
    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
}; # end function Get-OPXExchangeServerInMaintenance

function Add-OPXKeyToConfigFile
{
<#
.SYNOPSIS 
Add entries to configuration file.

.DESCRIPTION
Add configuration entries to a configuration file like the webconfig for an Exchange service. 

.PARAMETER ConfigFileUNCPath
The parameter is mandatory
The UNC file path to the config file where the entry should be added. If the entry should be added on multiple servers, the local path must be the same on all servers.

.PARAMETER KeyToAdd
The parameter is mandatory
An array of entries that should be added. In a webconfig, the first entry could be the description and the second entry the key. The command checks if the last entry in the array is already present in the config file. If the entry exists, adding the entry will be skipped. It is recommended to provide with the first entry in parameter the description and in the second entry the value. After the last entry, an additional line will be added.

.PARAMETER LineNumberToStartInsert
The parameter is mandatory for the parameterset LineNumber
The number of the line in the config file, where the data should be added.
The parameter cannot be used with the parameters AddBeforePatter or AddAfterPattern.

.PARAMETER AddBeforePattern
The parameter is mandatory for the parameterset AddBefore
A search string, which represents the content of the line. Before this line the data will be added.
The parameter cannot be used with the parameters LineNumberToStart or AddAfterPattern.

.PARAMETER AddAfterPattern
The parameter is mandatory for the parameterset AddAfter
A search string, which represents the content of the line. After this line the data will be added.
The parameter cannot be used with the parameters LineNumberToStart or AddBeforePattern.

.PARAMETER OverwriteOlderBackupFile
The parameter is optional
Before the entries will be added, a backup of the config file will be created. The backup file will be named <original file name.extension>.msx.backup. If the parameter is omitted and the backup file exist, the command will not overwrite it and will not update the configuration file.

.PARAMETER Server
The parameter is optional
The name of the server where the configuration should be updated. If the parameter is omitted, the name of the connected exchange server will be used.

.PARAMETER AllCAS
The parameter is optional
In addition to the server defined in the parameter Server, on all Exchange servers with CAS, the configuration will be updated.
The parameter cannot be used with the parameters DAGName or ServerList.
.PARAMETER DAGName
The parameter is optional
In addition to the server defined in the parameter Server, on all Exchange servers, which are member of a given DAG, the configuration will be updated too.
The parameter cannot be used with the parameters AllCAS or ServerList.

.PARAMETER ServerList
The parameter is optional
In addition to the server defined in the parameter Server, on all Exchange servers, listed in the parameter ServerList, the configuration will be updated too.
The parameter cannot be used with the parameters AllCAS or DAGName.

.PARAMETER OnlyVerifyIfKeyExist
The parameter is optional
Verifies if any of the entries in the parameter KeyToAdd exists in the configuration file. 
No configuration will be updted.

.EXAMPLE
$paramList=@{
    ConfigFileUNCPath='C:\Program Files\Microsoft\Exchange Server\V15\ClientAccess\ecp\Web.config';
    KeyToAdd=@(
    ("`t"+'<!-- allows the OU picker when placing a new mailbox in its designated organizational unit to retrieve all OUs - default value is 500 NOEL val = 1000 -->'),
    ("`t"+'<add key="GetListDefaultResultSize" value="1000" />')     
    ); # end KeyToAdd
     LineNumberToStartInsert=124;
}; # end paramList

Add-OPXKeyToConfigFile @ParamList -AllCAS -OnlyVerifyIfKeyExist
The command will verify if both entries exist on the Exchange servers with CAS. No configuration change will be made.

.EXAMPLE
$paramList=@{
    ConfigFileUNCPath='C:\Program Files\Microsoft\Exchange Server\V15\ClientAccess\ecp\Web.config';
    KeyToAdd=@(
    ("`t"+"`t"+'<!-- allows the OU picker when placing a new mailbox in its designated organizational unit to retrieve all OUs - default value is 500 new val = 1000 -->'),
    ("`t"+"`t"+'<add key="GetListDefaultResultSize" value="1000" />')   
    ); # end KeyToAdd
    AddBeforePattern='  </appSettings>';
}; # end paramList

Add-OPXKeyToConfigFile @ParamList -AllCAS 
The command will update the web config for ECP on the Exchange servers with CAS.
#>
[cmdletbinding(DefaultParametersetName='AddBefore')]
param([Parameter(Mandatory = $true, Position = 0)][string]$ConfigFileUNCPath,
      [Parameter(Mandatory = $true, Position = 1)][array]$KeyToAdd,
      [Parameter(ParameterSetName='LineNumber',Mandatory = $true, Position = 2)][int]$LineNumberToStartInsert,
      [Parameter(ParameterSetName='AddBefore',Mandatory = $true, Position = 2)][string]$AddBeforePattern,
      [Parameter(ParameterSetName='AddAfter',Mandatory = $true, Position = 2)][string]$AddAfterPattern,
      [Parameter(Mandatory = $false, Position = 3)][switch]$OverwriteOlderBackupFile=$false,      
      [Parameter(Mandatory = $false, Position = 4)]
      [ArgumentCompleter( {             
        param ( $CommandName,
                $ParameterName,
                $WordToComplete,
                $CommandAst,
                $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($false,$true,$true))
        $srvList.Where({ $_ -like "$wordToComplete*"})        
      } )]
      [string]$Server=$__OPX_ModuleData.ConnectedToMSXServer,
      [Parameter(Mandatory = $false, Position = 5)][switch]$AllCAS,       
      [Parameter(Mandatory = $false, Position = 5)]
      [ArgumentCompleter( {  
        param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )  
        $dagList=($__OPX_ModuleData.getDagList($false,$null));
        $dagList.Where({ $_ -like "$wordToComplete*"});                               
      } )]
      [string]$DAGName,
      [Parameter(Mandatory = $false, Position = 5)]
      [ArgumentCompleter( {             
        param ( $CommandName,
                $ParameterName,
                $WordToComplete,
                $CommandAst,
                $FakeBoundParameters )  
        $srvList=($__OPX_ModuleData.getExchangeServerList($false,$true,$true));
        $srvList.Where({ $_ -like "$wordToComplete*" });      
      } )]
      [Array]$ServerList,
      [Parameter(Mandatory = $false, Position = 6)][switch]$OnlyVerifyIfKeyExist
     )

    begin {
        if (([int]($PSBoundParameters.ContainsKey('ServerList')) + [int]($PSBoundParameters.ContainsKey('DAGName')) + [int]($PSBoundParameters.ContainsKey('AllCAS'))) -gt 1)
        {
            Throw ('Parameter set cannot be resolved using the specified named parameters.')            
        }; # end if        
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        if ($PSBoundParameters.ContainsKey('ServerList'))
        {
            $serverList=@([System.Linq.Enumerable]::Distinct([string[]]$serverList));
        }; # end if
        $webCfgContent=[System.Collections.ArrayList]::new();
        $serverCollection=[System.Collections.ArrayList]::New();
        try
        {
            if (([System.Uri]$ConfigFileUNCPath).IsUnc -eq $false) # check if path is an UNC path
            {
                if (($tmp=testIfFQDN -ComputerName $Server) -ne $false)
                {
                    [void]$serverCollection.Add($tmp.toLower());
                } # end if
                else
                {
                    writeToLog -LogString ('Faied to resolve the FQDN for computer ' + $Server) -LogType Error;
                    return;
                }; # end if
                $ConfigFileUNCPath=[System.IO.Path]::Combine(('\\'+$serverCollection[0]),($ConfigFileUNCPath).Replace(':','$'));        
            } # end if
            else { # is UNC
                $hostInUNC=([System.Uri]$ConfigFileUNCPath).Host;
                if (($tmp=testIfFQDN -ComputerName $Server) -ne $false)
                {
                    $ConfigFileUNCPath=($ConfigFileUNCPath -replace ('\\' + $hostInUNC + '\'), ('\\' + $tmp + '\'));
                    writeToLog -LogString ('Faied to resolve the FQDN for computer in the UNC path of the config file.' + $Server) -LogType Error;
                    return;
                } # end if
            }; # end else
            #>           
        } # end try
        catch
        {
            $msg=('Failed to create the UNC configuration file path for ' + $ConfigFileUNCPath);
            writeToLog -LogString $msg -LogType Error;
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
            return;
        }; # end catch 
        
        $tmpList=[System.Collections.ArrayList]::new();
        switch ($PSBoundParameters)
        {
            {$_.ContainsKey('AllCAS')}        {
                [void]$tmpList.AddRange((getAllCAS))      
                break;
            }; # end AllCAS
            {$_.ContainsKey('DAGName')}          {
                $dagSrvList=(getDAGMemberServer -DAGName $DAGName);
                if ($dagSrvList[0] -ne $false)
                {
                    [void]$tmpList.AddRange($dagSrvList);
                }; # end if                
                break;
            }; # end DAG
            {$_.ContainsKey('ServerList')}    {
                [void]$tmpList.AddRange($ServerList);
            }; # end ServerList                
        }; # end switch
        
        foreach ($entry in $tmpList)
        {
            if ((!([system.string]::IsNullOrEmpty($entry))) -and ($tmp=testIfFQDN -ComputerName $entry)) # make sure FQDN can be resolved
            {
                [void]$serverCollection.Add($tmp.tolower());
            } # end if
            else
            {
                writeToLog -LogString ('Faied to resolve the FQDN for computer ' + $entry) -LogType Error;
            }; # end if
        }; # end foreach
        $serverCollection=@([System.Linq.Enumerable]::Distinct([string[]]$serverCollection)); # get list of unique entries
        $srvCount=$serverCollection.count;
        for ($i=0;$i -lt $srvCount;$i++)
        {
            $msg=('Preparing update for config file on server ' + $serverCollection[$i])
            writeToLog -LogString $msg;
            Write-Verbose $msg;
            $ConfigFileUNCPath=$ConfigFileUNCPath.Replace(('\\' + $serverCollection[([math]::max(0,$i-1))] + '\'),('\\' + $serverCollection[$i] + '\'));
            try {
                $msg=('Reading config file ' + $ConfigFileUNCPath);
                Write-Verbose $msg;
                writeToLog -LogString $msg;
                [void]$webCfgContent.Clear();
                [void]$webCfgContent.AddRange((Get-Content -Path $ConfigFileUNCPath)); # read the current config file
            } # end try
            catch {
                $msg=('Faild to read the configuraion file ' + $ConfigFileUNCPath);
                writeToLog -LogString $msg -LogType Error;
                continue;
            }; # end catch
            $numOfEntries=$KeyToAdd.Count;
            if ($OnlyVerifyIfKeyExist.IsPresent)
            {
                $msg=('Verifiying if entries on server ' + $serverCollection[$i]);
                Write-Verbose $msg;
                writeToLog -LogString $msg;                
                for ($j=0;$j -lt $numOfEntries;$j++)
                {
                    if ($webCfgContent -like ('*'+$KeyToAdd[$j].Trim()+'*') )
                    {
                        $msg=('The entry ' + $KeyToAdd[$j] + ' already exists in the config file ' + $ConfigFileUNCPath);
                        writeTolog -LogString $msg -LogType Warning;
                    }; # end if
                }; # end for
            } # end if
            else {
                try {                    
                    $msg=('Verifying if key exist (last entry in array of parameter KeyToAdd)');
                    writeToLog -LogString $msg;
                    Write-Verbose $msg;
                    $entryExist=$false;
                    for ($j=0;$j -lt $numofEntries;$j++)
                    {
                        if ($webCfgContent -like ('*'+$KeyToAdd[$j].Trim()+'*') )
                        {
                            $msg=('The entry ' + $KeyToAdd[$j] + ' already exists in the config file ' + $ConfigFileUNCPath);
                            writeTolog -LogString $msg -LogType Warning;
                            $entryExist=$true;
                        }; # end if
                    }; # end if
                    if ($entryExist -eq $false) # verify that the value or one of the values is not already present
                    {
                        writeToLog -LogString 'Looking for Line to start with insert'
                        switch ($PsCmdlet.ParameterSetName)
                        {
                            'AddBefore'     {
                                $startAdd=($webCfgContent.IndexOf($AddBeforePattern));
                            }; # end AddBefore
                            'AddAfter'     {
                                $startAdd=($webCfgContent.IndexOf($AddAfterPattern)+1);
                            }; # end AddAfter
                            'LineNumber'    {
                                $startAdd=($LineNumberToStartInsert-1);
                            }; # end LineNumber
                        }; # end switch         
                        
                        $msg=('Adding data starting with line ' + ($startAdd+1).ToString());
                        Write-Verbose $msg;
                        writeToLog -LogString $msg;
                        if ($startAdd -lt 0)
                        {
                            writeToLog -LogString ($serverCollection[$i] +  ': No valid startposition for adding key found.') -LogType Warning;
                            continue;
                        }; # end if
                        $webCfgContent.InsertRange(($startAdd),$KeyToAdd);
                        try
                        {
                            $backupFileName=([System.IO.Path]::ChangeExtension($ConfigFileUNCPath,([System.IO.Path]::GetExtension($ConfigFileUNCPath)+'.msx.backup')));
                            $msg=('Creating backup file ' + $backupFileName);
                            Write-Verbose $msg;
                            writeToLog -LogString $msg;
                            [System.IO.File]::Copy($ConfigFileUNCPath,$backupFileName,($OverwriteOlderBackupFile.isPresent));
                            $msg=('Saving data to file ' + $ConfigFileUNCPath);
                            Write-Verbose $msg;
                            writeToLog -LogString $msg;
                            $webCfgContent | Out-File -FilePath $ConfigFileUNCPath;                
                        } # end try
                        catch
                        {
                            $msg=('Failed to update the config file ' + $ConfigFileUNCPath);
                            writeToLog -LogString  $msg -LogType Error;
                            writeToLog -LogString ($_.Exception.Message) -LogType Error;
                        }; # end catch
                    } # end if
                    else
                    {
                        $msg=('Server ' + $serverCollection[$i] + ': The key already exist. File not updated.');
                        WriteToLog -LogString $msg -LogType Warning;
                    }; # end else
                } # end try
                catch {
                    $msg=('Unhandled expeption for configuraion file ' + $ConfigFileUNCPath);
                    writeToLog -LogString $msg -LogType Error;
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;
                    continue;
                }; # end catch
                $msg=('Finished with config file update on server ' + $serverCollection[$i]);
                writeTolog -LogString $msg;
                Write-Verbose $msg;
            }; # end else
            
        }; # end foreach        
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END      
}; # end function Add-OPXKeyToConfigFile
function Clear-OPXVirtualDirectoryConfigurationTemplate
{
<#
.SYNOPSIS 
Clears the entry for a service in the VirtualDirectoryConfigurationTemplate.

.DESCRIPTION
To configure or to list the various virtual directories, an optional configuration file can be used. The configuration file contains the following service nodes
	OWA
	ECP
	OAB
	Webservices
	ActiveSync
	OutlookAnywhere
	Mapi
	Autodiscover
The command can clear the configuration of any of these nodes.

.PARAMETER Service
The parameter is mandatory for clearing a service template
Allows to select the service, for which the configuration should be cleared.

.PARAMETER Service
The parameter is mandatory for resetting the template
Reset the template to default value. The cmdelet will not check if the template exists.

.EXAMPLE
Clear-OPXVirtualDirectoryConfigurationTemplate -Service OWA

.EXAMPLE
Clear-OPXVirtualDirectoryConfigurationTemplate -ResetTemplate
#>
[cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High',DefaultParametersetName='ClearService')]
param([Parameter(ParameterSetName='ClearService',Mandatory = $true, Position = 0)][ValidateSet('OWA','ECP','OAB','Webservices','ActiveSync','OutlookAnywhere','Mapi','Autodiscover')][string]$Servcice,
      [Parameter(ParameterSetName='ResetTemplate',Mandatory = $true, Position = 0)][switch]$ResetTemplate
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        $cfgFilePath=[system.io.path]::Combine( $__OPX_ModuleData.cfgDir,'VDirs\' + $__OPX_ModuleData.VirtualDirCfgFileName);
        if ($PsCmdlet.ParameterSetName -eq 'ClearService')
        {
        if ($vDirConfig=getVDirConfig -ConfigFilePath $cfgFilePath)  # reading configuration
        {
            try {
                writeTolog -LogString ('Clearing configuration for service ' + $Servcice);
                $vDirConfig.$Servcice=@{};
                writeTolog -LogString ('Saving configuration to ' + $cfgFilePath);
                $vDirConfig | Export-Clixml -Path $cfgFilePath -ErrorAction stop -Force;
            } # end try
            catch {
                writeTolog -LogString ('Faild to save the configuration file ' + $vDirConfig) -LogType Error;
                writeTolog -LogString ($_.Exception.Message) -LogType Error;
            }; # end catch
        }; #end if
        } # end if
        else {
            if ($PSCmdlet.ShouldProcess('Do you want to reset the VDir configuration template?'))
            {
                resetVdirConfig -ConfigFilePath $ConfigFilePath;
            } # end if
            else {
                writeToLog -LogString 'Unser canceled clearing VDir template.'
            }; # end else
        }; # end else
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END     

}; # end function Clear-OPXVirtualDirectoryConfigurationTemplate

function Get-OPXVirtualDirectoryConfigurationTemplate
{
<#
.SYNOPSIS 
List the entry for the configuration template.

.DESCRIPTION
Lists the configuration template for all services nodes, or in detail, for a particular service.
The configuration file contains the following service nodes
	OWA
	ECP
	OAB
	Webservices
	ActiveSync
	OutlookAnywhere
	Mapi
	Autodiscover

.PARAMETER Service
The parameter is mandatory
Allows to select the service, for which the configuration should be listed.

.EXAMPLE
Get-OPXVirtualDirectoryConfigurationTemplate -Service OWA
Get detailed info for the additional OWA configuration.

.EXAMPLE
Get-OPXVirtualDirectoryConfigurationTemplate
Get an overview of the configuration for all service nodes.
#>
[cmdletbinding(DefaultParametersetName='AllServices')]
param([Parameter(ParametersetName='Service',Mandatory = $true, Position = 0)][ValidateSet('OWA','ECP','OAB','Webservices','ActiveSync','OutlookAnywhere','Mapi','Autodiscover')][string]$Servcice,
      [Parameter(ParametersetName='AllServices',Mandatory = $false, Position = 0)][switch]$AllServices=$false
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        $cfgFilePath=[system.io.path]::Combine( $__OPX_ModuleData.cfgDir,'VDirs\' + $__OPX_ModuleData.VirtualDirCfgFileName);
        if ($vDirConfig=getVDirConfig -ConfigFilePath $cfgFilePath)  # reading configuration
        {
            if ($PSCmdlet.ParameterSetName -eq 'Service') 
            {
                writeTolog -LogString ('Reading configuration for service ' + $Servcice);
                $vDirConfig.$Servcice;    
            } # end if
            else
            {
                writeTolog -LogString ('Reading configuration');
                $vDirConfig;
            }; # end if
        }; #end if
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END     

}; # end function Clear-OPXVirtualDirectoryConfigurationTemplate

function Set-OPXVirtualDirectoryConfigurationTemplate
{
<#
.SYNOPSIS 
Configures the entries for the configuration template.

.DESCRIPTION
Configures the configuration template for service nodes.
The configuration file contains the following service nodes
	OWA
	ECP
	OAB
	Webservices
	ActiveSync
	OutlookAnywhere
	Mapi
	Autodiscover
To configure a service node the following items are needed:
	name of parameter (of the corresponding Exchange PowerShell cmdlet)
	value (a valid value for the parameter)
Possible parameter names are preconfigured in the <serviceName>.params files, which are stored under <module root directory>\cfg\VDirs>. For a single service node, multiple entries can be configured.

.PARAMETER Service
The parameter is mandatory
Allows to select the service to configure.

.PARAMETER Parameter
The parameter is mandatory
Select the parameter for the service to configure.

.PARAMETER Value
The parameter is mandatory
A valid value for the parameter. 

.EXAMPLE
Set-OPXVirtualDirectoryConfigurationTemplate -Service OWA -Parameter DefaultDomain -Value <domain.name>
Configures a default domain name for OWA.

.EXAMPLE
Set-OPXVirtualDirectoryConfigurationTemplate -Service OWA -Parameter LogonFormat -Value UserName
Configures the logon format for OWA.
#>
[cmdletbinding()]
param([parameter(Mandatory = $true)]
      [ValidateSet('OWA','ECP','OAB','Webservices','ActiveSync','OutlookAnywhere','Mapi','Autodiscover')]
      [string] $Service,
      [parameter(Mandatory = $true)]
      [ArgumentCompleter( {
            param ( $CommandName,
                $ParameterName,
                $WordToComplete,
                $CommandAst,
                $FakeBoundParameters )
            $ParameterPerService = @{
                OWA = @(get-content -Path ([System.IO.Path]::Combine($__OPX_ModuleData.cfgDir,'VDirs\OWA.params')));
                ECP = @(get-content -Path ([System.IO.Path]::Combine($__OPX_ModuleData.cfgDir,'VDirs\ECP.params')));
                OAB = @(get-content -Path ([System.IO.Path]::Combine($__OPX_ModuleData.cfgDir,'VDirs\OAB.params')));
                ActiveSync = @(get-content -Path ([System.IO.Path]::Combine($__OPX_ModuleData.cfgDir,'VDirs\ActiveSync.params')));
                Webservices = @(get-content -Path ([System.IO.Path]::Combine($__OPX_ModuleData.cfgDir,'VDirs\Webservices.params')));
                OutlookAnywhere=@(get-content -Path ([System.IO.Path]::Combine($__OPX_ModuleData.cfgDir,'VDirs\OutlookAnywhere.params')));
                Mapi = @(get-content -Path ([System.IO.Path]::Combine($__OPX_ModuleData.cfgDir,'VDirs\Mapi.params')));
                Autodiscover = @(get-content -Path ([System.IO.Path]::Combine($__OPX_ModuleData.cfgDir,'VDirs\Autodiscover.params')));
            }
            if ($fakeBoundParameters.ContainsKey('Service'))
            {
                @($ParameterPerService[$fakeBoundParameters.Service]).Where({ $_ -like "$wordToComplete*" });
            } # end if
            else
            {
                'Error: Missing service'
            }; # end else
        })]
      [string] $Parameter,
      [Parameter(Mandatory = $true, Position = 2)]$Value
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        $cfgFilePath=[system.io.path]::Combine( $__OPX_ModuleData.cfgDir,'VDirs\' + $__OPX_ModuleData.VirtualDirCfgFileName);
        if ($vDirConfig=getVDirConfig -ConfigFilePath $cfgFilePath)  # reading configuration
        {
            try {
                writeTolog -LogString ('Setting parameter ' + $Parameter + ' for service ' + $Service + ' to ' + $Value.toString());
                $vDirConfig.$service.$Parameter=$Value; # setting configuration
                try {
                    saveVDirConfig -ConfigFilePath $cfgFilePath -CfgHash $vDirConfig; # saving configuration
                } # end try
                catch {
                    writeTolog -LogString ('Failed to save the configuration to file ' + $cfgFilePath) -LogType Error;
                    writeTolog -LogString ($_.Exception.Message) -LogType Error;
                }; # end catch
            } # end try
            catch {
                writeTolog -LogString ('Failed to set the parameter ' + $Parameter + ' for the service ' + $service) -LogType Error;
                writeTolog -LogString ($_.Exception.Message) -LogType Error;
            }; # end catch
        }; # end if
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
}; # end function Set-OPXVirtualDirectoryConfigurationTemplate

function Remove-OPXVirtualDirectoryConfigurationFromTemplate
{
<#
.SYNOPSIS 
Removes an entry from the configuration template.

.DESCRIPTION
Removes an entry for a service in the configuration.
The configuration file contains the following service nodes
	OWA
	ECP
	OAB
	Webservices
	ActiveSync
	OutlookAnywhere
	Mapi
	Autodiscover

.PARAMETER Service
The parameter is mandatory
Allows to select the service to configure.

.PARAMETER Parameter
The parameter is mandatory
Select the parameter to remove.

.EXAMPLE
Remove-OPXVirtualDirectoryConfigurationFromTemplate -Service OWA -Parameter DefaultDomain 
Removes the parameter DefaultDomain for OWA configuration.
#>
[cmdletbinding()]
param([parameter(Mandatory = $true)]
      [ValidateSet('OWA','ECP','OAB','Webservices','ActiveSync','OutlookAnywhere','Mapi','Autodiscover')]
      [string] $Service,
      [parameter(Mandatory = $true)]
      [ArgumentCompleter( {
            param ( $CommandName,
                $ParameterName,
                $WordToComplete,
                $CommandAst,
                $FakeBoundParameters )
            $cfg=Import-Clixml -Path ([System.IO.Path]::Combine($__OPX_ModuleData.cfgDir,'VDirs\' + $__OPX_ModuleData.VirtualDirCfgFileName));
            $ParameterPerService = @{
                OWA = @([array]($cfg.OWA.keys));
                ECP = @($cfg.ECP.keys);
                OAB = @($cfg.OAB.keys);
                ActiveSync = @($cfg.ActiveSync.keys);
                Webservices = @($cfg.Webservices.keys);
                OutlookAnywhere=@($cfg.OutlookAnywhere.keys);
                Mapi = @($cfg.Mapi.keys);
                Autodiscover = @($cfg.Autodiscover.keys);
            }
            if ($fakeBoundParameters.ContainsKey('Service'))
            {
                ($ParameterPerService[$fakeBoundParameters.Service]).Where({$_ -like "$wordToComplete*"});
            } # end if
            else
            {
                'Error: Missing service'
            }; # end else
        })]
      [string] $Parameter,
      [Parameter(Mandatory = $false, Position = 2)]$Value
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        $cfgFilePath=[system.io.path]::Combine( $__OPX_ModuleData.cfgDir,'VDirs\' + $__OPX_ModuleData.VirtualDirCfgFileName)
        if ($vDirConfig=getVDirConfig -ConfigFilePath $cfgFilePath)  # reading configuration
        {
            try {
                writeTolog -LogString ('Removing parameter ' + $Parameter + ' from configuration for service ' + $Service);
                $vDirConfig.$service.Remove($Parameter); # removing configuration
                try {
                    saveVDirConfig -ConfigFilePath $cfgFilePath -CfgHash $vDirConfig; # saving configuration
                } # end try
                catch {
                    writeTolog -LogString ('Failed to save the configuration to file ' + $cfgFilePath) -LogType Error;
                    writeTolog -LogString ($_.Exception.Message) -LogType Error;
                }; # end catch
            } # end try
            catch {
                writeTolog -LogString ('Failed to remove the parameter ' + $Parameter + ' for the service ' + $service) -LogType Error;
                writeTolog -LogString ($_.Exception.Message) -LogType Error;
            }; # end catch
        }; # end if
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
}; # end function Remove-OPXVirtualDirectoryConfigurationFromTemplate

#endretion experimental cmdlets

#region in-module use
function Get-OPXExchangeServer
{
<#
.SYNOPSIS 
Lists Exchange server.

.DESCRIPTION
Lists Exchange server with the role CAS or mailbox. The command returns the following values:
	name of the server
	Exchange version installed
	DAG membership
	AD site
	component and maintenance state overview

.PARAMETER ReturnMailboxServer
The parameter is optional
If the parameter is use, Exchange servers with the role mailbox are returned. If the parameter is omitted, Exchange servers with the role CAS are returned. The parameter is only required for older Exchange servers (Exchange 2013).

.PARAMETER ReturnFQDN
The parameter is optional
If the parameter is used the FQDN of server is returned. If the parameter is omitted, the NetBIOS name of the server is returned.

.PARAMETER ReturnFQDN
The parameter is optional
With the parameter the format of the output returned can be selected. Valid options are
	Table (default)
	List

.PARAMETER ListComputerNameOnly
The parameter is optional
Only the names of the Exchange servers are returned. With this switch the command can be used to pipe the names of the Exchange servers to an other command, like Get-OPXLastComputerBootTime.

.EXAMPLE
Get-OPXExchangeServer -ReturnFQND
The FQDN of the Exchange servers is returned.
#>
[cmdletbinding()]
param([Parameter(Mandatory = $false, Position = 0)][switch]$ReturnMailboxServer=$false,
      [Parameter(Mandatory = $false, Position = 1)][switch]$ReturnFQDN=$false,
      [Parameter(Mandatory = $false, Position = 2)][OutFormat]$Format='Table',
      [Parameter(Mandatory = $false, Position = 2)][switch]$ListComputerNameOnly=$false
     )
    
    
    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin
    process {
        
        if ($ListComputerNameOnly.IsPresent)
        {
            $__OPX_ModuleData.getExchangeServerList($ReturnMailboxServer,$ReturnFQDN,$true); 
        } # end if
        else {
           switch ($Format)
            {
                'Table'     {
                    $__OPX_ModuleData.getExchangeServerList($ReturnMailboxServer,$ReturnFQDN,$false) | Format-Table;
                }; # end table
                'List'     {
                    $__OPX_ModuleData.getExchangeServerList($ReturnMailboxServer,$ReturnFQDN,$false) | Format-List;
                }; # end table
                'PassValue'     {
                    $__OPX_ModuleData.getExchangeServerList($ReturnMailboxServer,$ReturnFQDN,$false);
                }; # end table
            }; # end switch 
        }; # end else
        
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
}; # end function Get-OPXExchangeServer

function Get-OPXLastComputerBootTime
{
<#
.SYNOPSIS 
Get the last boot time of a computer.

.DESCRIPTION
With the command, the last boot time of a computer can be returned. For the parameter ComputerName the pipeline is supported.

.PARAMETER ComputerName
The parameter is optional
Name for the computer which last boot time should be displayed. If the parameter is omitted the local computer is used.
For the parameter the pipeline can be used (postion and name).

.PARAMETER ReturnFQDN
The parameter is optional
With the parameter the format of the output returned can be selected. Valid options are
	Table (default)
	List

.EXAMPLE
Get-OPXLastComputerBootTime -ComputerName <computer name>
The last boot time for a computer is returned.

.EXAMPLE
Get-OPXExchangeServer -ReturnFQDN -ListComputerNameOnly | Get-OPXLastComputerBootTime
A list of Exchange server with the last boot time will be displayed.
#>
[cmdletbinding()]
param([Parameter(Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position = 0)][string]$ComputerName=([System.Net.Dns]::GetHostByName('').hostname),
      [Parameter(Mandatory = $false, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, Position = 1)][ValidateSet('Table','List')][string]$Format='Table'
     )

    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
        $fieldList=@(
            @('ComputerName',[System.String]),
            @('LastBootTime',[System.DateTime]),
            @('OS Version',[System.String])
        ); # end fieldList
        [void]($tblOutput=createNewTable -TableName 'OutputTable' -FieldList $fieldList);   
    }; # end begin

    process {       
        Write-Verbose ('Processing computer ' + $ComputerName);
        try
        {
            writeTolog -LogString ('Querying computer ' + $computerName + ' for last boot time.')
            $rv=Get-CimInstance -ComputerName $ComputerName -ClassName win32_operatingsystem -Verbose:$false -ErrorAction stop;
            [void]$tblOutput.rows.add($ComputerName,$rv.lastbootuptime,$rv.version);            
        } # end try
        catch
        {
            Write-Warning ('Failed to get the boot information for computer ' + $ComputerName);
        }; # end catch
        
    }; # end process

    end {
        if ($Format -eq 'List')
        {
            $tblOutput | Format-List;
        } # end if
        else
        {
            $tblOutput;
        }; # end else
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
}; # end function

function Get-OPXCumulatedSizeOfMailboxesInGB
{
<#
.SYNOPSIS 
Calculate the size of mailboxes.

.DESCRIPTION
The command is intended for mailbox migrations/moves when circular logging is disabled.

.PARAMETER OrganizationalUnit
The parameter is mandatory
Name of an organizational unit where the mailbox sizes should be calculated. The parameter cannot be used with the parameter CSVInputFilePath.
If size for all mailboxes in the current domian should be calculated, provide $NULL as name for the organizational unit.

.PARAMETER CSVInputFilePath
The parameter is mandatory
Name of the CSV file with the identiy of the mailboxes where the size should be calculated. A CSV file with the singele column, named Identity, is expected.
The  parameter cannot be used with the parameter OrganizationalUnit.

.EXAMPLE
Get-OPXCumulatedSizeOfMailboxesInGB -OrganizationalUnit <OU name>

.EXAMPLE
Get-OPXCumulatedSizeOfMailboxesInGB -CSVFileNamePath <full file name>
#>
[cmdletbinding(DefaultParametersetName='OU')]
param([Parameter(ParametersetName='OU',Mandatory = $true, Position = 0)][AllowEmptyString()][string]$OrganizationalUnit,
      [Parameter(ParametersetName='File',Mandatory = $false, Position = 2)][string]$CSVInputFilePath
     )
    begin {
        writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ParameterList ($MyInvocation.MyCommand.Parameters.GetEnumerator());
    }; # end begin

    process {
        if ($PSCmdlet.ParameterSetName -eq 'OU')
        {
            try
            {
                if ([System.String]::IsNullOrEmpty($OrganizationalUnit))
                {
                    $OU=$null;
                } # end if
                else
                {
                    $OU=$OrganizationalUnit;
                    $msg=('Calculating size of mailboxes from OU ' + $OrganizationalUnit);
                    Write-Verbose $msg;
                    writeTolog -LogString $msg;
                }; # end if
                $val=[math]::round([System.Linq.Enumerable]::Sum([int[]]((((Get-Mailbox -OrganizationalUnit $OU -ResultSize Unlimited -ErrorAction Stop | Get-MailboxStatistics -ErrorAction Stop).TotalItemSize.Value) | ForEach-Object {$_.ToString().split('(')[1].split(' ')[0].replace(',','')})))/1GB,2);
                Write-Host ($val.ToString() + ' GB');
            } # end try
            catch {
                $msg=('Failed to calculate the mailbox size.')
                writeToLog -LogString $msg -LogType Error;
                writeTolog -LogString ($_.Exception.Message) -LogType Error;
            }; # end catch
        } # end if
        else {
            $msg=('Caluclating size of mailboxes from file ' + $CSVInputFilePath);
            Write-Verbose $msg;
            writeTolog -LogString $msg;
            try {
                $val=[math]::round([System.Linq.Enumerable]::Sum([int[]]((((Import-csv -Path $CSVInputFilePath -ErrorAction Stop | Get-MailboxStatistics -ErrorAction Stop).TotalItemSize.Value) | ForEach-Object {$_.ToString().split('(')[1].split(' ')[0].replace(',','')})))/1GB,2);
                Write-Host ($val.ToString() + ' GB');
            } # end try
            catch
            {
                $msg=('Failed to calculate the mailbox size.')
                writeToLog -LogString $msg -LogType Error;
                writeTolog -LogString ($_.Exception.Message) -LogType Error;
            }; # end catch
        }; # end else
    }; # end process

    end {
        [void](writeCmdletInitToDataLog -CallStack (Get-PSCallStack) -ExitLogEntry);
    }; # end END
} # end function Get-OPXCumulatedSizeOfMailboxesInGB 
#endregion in-moudule use
