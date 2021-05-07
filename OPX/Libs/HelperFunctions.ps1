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

#region helper

function getMSXVersions
{
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)][string]$Server,
      [Parameter(Mandatory = $true, Position = 1)][hashtable]$fParamList,
      [Parameter(Mandatory = $true, Position = 2)][hashtable]$dParamList,
      [Parameter(Mandatory = $true, Position = 3)][string]$serverRoot
     )
    $serverRoot=([ADSI]('LDAP://'+$server+'/RootDSE')).defaultNamingContext;
    try
    {
        [void][System.Net.DNS]::GetHostByName($server).HostName;
    } # end try
    catch
    {
        writeTolog -LogString ('Failed to resolve the server ' + $server) -LogType Warning -CallingID $Script:CallingID;
        #Write-Host ($_.Exception.Message);
        continue;
    }; # end catch
    $rvHash=@{};
    $fParamList.server=$server;
    $fParamList.LDAPFilter='(objectClass=attributeSchema)';
    $fParamList.SearchBase=$msxSchemaRoot;
    $dParamList.server=$server;
    $fParamList.FieldList=@('rangeUpper');
    $rvHash.Add('msxSchemaInfo',(queryADForObjects @fParamList));
                        
    $fParamList.SearchBase=$msxCfgRoot; #$msxCfg+$serverRoot;
    $fParamList.LDAPFilter='(objectClass=msExchOrganizationContainer)';
    $fParamList.FieldList=@('objectversion');
    $fParamList.Add('SearchScope','OneLevel');
    $rvHash.Add('msxConfigInfo',(queryADForObjects @fParamList));

    $fParamList.Remove('SearchScope'); 
    $dParamList.searchBase=('CN=Microsoft Exchange System Objects,'+$serverRoot); 
    $rvHash.Add('domInfo',(queryADForObjects @dParamList));
    
    return ,$rvHash;
}; # end function getMSXVersions

function isDAGMember 
{
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)][string]$ServerName
     )
    
    try
    {
        if ((Get-MailboxServer -Identity $ServerName).DatabaseAvailabilityGroup)
        {
            return $true;
        } # end if
        else
        {
            return $false;
        }; # end else
    } # end try
    catch
    {
        return $false;
    }; # end catch
} # end function isDAGMember

function queryADForObjects
{
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)][string]$LDAPFilter,
      [Parameter(Mandatory = $false, Position = 1)][string]$SearchBase=(([ADSI]"LDAP://RootDSE").defaultNamingContext),
      [Parameter(Mandatory = $false, Position = 2)][array]$FieldList,
      [Parameter(Mandatory = $false, Position = 3)][string]$Server,
      [Parameter(Mandatory = $false, Position = 4)][int]$PageSize=1000,
      [Parameter(Mandatory = $false, Position = 3)][string]$SearchScope='Subtree'
     )
        
    if ($PSBoundParameters.ContainsKey('server'))
    {
        try
        {
            $serverFQDN=[System.Net.DNS]::GetHostByName($server).hostname;
            $searcherStr='LDAP://'+$serverFQDN+'/'+$SearchBase;
        } # end try
        catch
        {
            writeTolog -LogString ('Failed to resolve server ' + $Server) -LogType Error;
            writeTolog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch
    }# end if
    else
    {
        $searcherStr=('LDAP://'+$SearchBase);
    }; #end else
    
    $searchRoot = New-Object System.DirectoryServices.DirectoryEntry($searcherStr);
    $ADSearcher = New-Object System.DirectoryServices.DirectorySearcher;
    
    $ADSearcher.SearchRoot = $searchRoot;
    $ADSearcher.Filter = $LDAPFilter;
    $ADSearcher.SearchScope = $SearchScope;
    $ADSearcher.PageSize = $PageSize;
    
    if ($PSBoundParameters.ContainsKey("FieldList"))
    {
        [void]$ADSearcher.PropertiesToLoad.AddRange($FieldList);        
    }; # end if
    try
    {
        return ,$ADSearcher.FindAll();
    } # end try
    catch
    {
        return ,$null;
    }; # end catch
    

} # end function queryADForObjects

function getURLList
{
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)][array]$SrvList,
      [Parameter(Mandatory = $false, Position = 1)][switch]$ADPropertiesOnly=$false,
      [Parameter(Mandatory = $false, Position = 2)][switch]$SkipAcceptedDomains=$false,
      [Parameter(Mandatory = $false, Position = 3)][array]$DomainEntryList
      )

    try
    {
        $msg=('Collecting inforamtion from server(s) ' + ($srvlist -join ', '));
        Write-Verbose $msg;
        writeTolog -LogString $msg;
        Write-Verbose 'Collecting namespace information for Autodiscover...';
        writeTolog -LogString 'Collecting namespace information for Autodiscover' -LogType Info; 
        $domEntryList=[System.Collections.ArrayList]::new();
        $progressBarMaxIterations=($srvlist.count*8+1);
        if ($PSBoundParameters.ContainsKey('SkipAcceptedDomains'))
        {            
            $progressBarMaxIterations--;
        } # end if
        if ($PSBoundParameters.ContainsKey('DomainEntryList'))
        {
            [void]$domEntryList.AddRange($DomainEntryList);
        }; # end if
        $server='';
        $service='Autodiscover';
        $progressBarIteration=1;
        $progressParams=@{
            Activity=('Collecting virtual directory information for service ');
            Status=($service +' on server ' + $server);
            PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
        }; # end progressParam
        
        foreach ($server in $srvList) 
        { 
            $progressBarIteration++;
            $progressParams.Status=('Autodiscover on server ' + $server);
            $progressParams.PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
            Write-Progress @progressParams;
            $tmp = '';
            $tmp= (Get-ClientAccessService -Identity $server).AutoDiscoverServiceInternalUri;
            if ($null -ne $tmp )
            {
                [void]$domEntryList.Add($tmp.host.toLower());
            } # end if
            else
            {
                writeTolog -LogString ("The AutoDiscoverServiceInternalUri on server " + $server + " is empty.") -LogType Warning;
            }; # end else
        }; # end foreach
    
        # get owa url 
        Write-Verbose "Collecting namespace information for OWA..."
        writeTolog -LogString 'Collecting namespace information for OWA' -LogType Info;
        
        foreach ($server in $srvList)
        {
            $progressBarIteration++;
            $progressParams.Status=('OWA on server ' + $server);
            $progressParams.PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
            Write-Progress @progressParams;
            $OWA = Get-OwaVirtualDirectory -Server $server -ADPropertiesOnly:$ADPropertiesOnly;
            $tmpInt = $OWA.InternalUrl
            $tmpEX = $OWA.ExternalUrl
            If ($null -ne $tmpInt)
            {
                [void]$domEntryList.Add($tmpInt.host.toLower());
            } # end if
            else
            {
                writeTolog -LogString ('The inernal URL for the OWA virtual directory on Server '+ $server + ' is empty.') -LogType Warning;
            }; # end else

            if ($null -ne $tmpEX)
            {
                [void]$domEntryList.Add($tmpEx.host.toLower());
            } # end if
            else
            {
                writeTolog -LogString ('The external URL for the OWA virtual directory on Server '+ $server + ' is empty.') -LogType Warning;
            }; # end else
            $tmpInt = ""
            $tmpEX = ""    
        } # foreach CAS in list
    
        $tmpInt = ""
        $tmpEX = ""

        # get ecp url
        Write-Verbose "Collecting namespace information for ECP..."
        writeTolog -LogString 'Collecting namespace information for ECP' -LogType Info;
        foreach ($server in $srvList)
        {               
            $progressBarIteration++;
            $progressParams.Status=('ECP on server ' + $server);
            $progressParams.PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
            Write-Progress @progressParams;
            $Ecp = Get-EcpVirtualDirectory -Server $Server -ADPropertiesOnly:$ADPropertiesOnly;
            $tmpInt = $Ecp.InternalUrl
            $tmpEX = $Ecp.ExternalUrl
            If ($null -ne $tmpInt)
            {
                [void]$domEntryList.Add($tmpInt.host.toLower());
            } # end if
            else
            {
                writeTolog -LogString ('The inernal URL for the ECP virtual directory on Server '+ $server + 'is empty.') -LogType Warning;
            }; # end else

            if ($null -ne $tmpEX)
            {
                [void]$domEntryList.Add($tmpEx.host.toLower());
            } # end if
            else
            {
                writeTolog -LogString ('The external URL for the ECP virtual directory on Server '+ $server + ' is empty.') -LogType Warning;
            }; # end else
            $tmpInt = ""
            $tmpEX = ""
        }; #end foreach

        #get oab url
        Write-Verbose "Collecting namespace information for OAB..."
        writeTolog -LogString 'Collecting namespace information for OAB' -LogType Info;
        foreach ($server in $srvList)
        {               
            $progressBarIteration++;
            $progressParams.Status=('OAB on server ' + $server);
            $progressParams.PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
            Write-Progress @progressParams;
            $Oab = Get-OabVirtualDirectory -Server $server -ADPropertiesOnly:$ADPropertiesOnly;
            $tmpInt = $Oab.InternalUrl
            $tmpEX = $Oab.ExternalUrl
            If ($null -ne $tmpInt)
            {
                [void]$domEntryList.Add($tmpInt.host.toLower());
            } # end if
            else
            {
                writeTolog -LogString ('The inernal URL for the OAB virtual directory on Server '+ $server + ' is empty.') -LogType Warning;
            }; # end else

            if ($null -ne $tmpEX)
            {
                [void]$domEntryList.Add($tmpEx.host.toLower());
            } # end if
            else
            {
                writeTolog -LogString ('The external URL for the OAB virtual directory on Server '+ $server + ' is empty.') -LogType Warning;
            }; # end else

            $tmpInt = ""
            $tmpEX = ""
        }; # end foreach
    
        # get ActiveSync url
        Write-Verbose "Collecting namespace information for ActiveSync..."
        writeTolog -LogString 'Collecting namespace information for ActiveSync' -LogType Info;
        foreach ($server in $srvList)
        {               
            $progressBarIteration++;
            $progressParams.Status=('ActiveSync on server ' + $server);
            $progressParams.PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
            Write-Progress @progressParams;
            $AS = Get-ActiveSyncVirtualDirectory -Server $server -ADPropertiesOnly:$ADPropertiesOnly;
            $tmpInt = $AS.InternalUrl
            $tmpEX = $AS.ExternalUrl
            If ($null -ne $tmpInt)
            {
                [void]$domEntryList.Add($tmpint.host.toLower());
            } # end if
            else
            {
                writeTolog -LogString ('The inernal URL for the ActiveSync virtual directory on Server '+ $server + ' is empty.') -LogType Warning;
            }; # end else

            if ($null -ne $tmpEX)
            {
                [void]$domEntryList.Add($tmpEX.host.toLower());
            } # end if
            else
            {
                writeTolog -LogString ('The external URL for the ActiveSync virtual directory on Server '+ $server + ' is empty.') -LogType Warning;
            }; # end else
            $tmpInt = ""
            $tmpEX = ""
        }; # end foreach

        # get ews url
        Write-Verbose "Collecting namespace information for EWS..."
        writeTolog -LogString 'Collecting namespace information for EWS' -LogType Info;
        foreach ($server in $srvList)
        {     
            $progressBarIteration++;
            $progressParams.Status=('WebServices on server ' + $server);
            $progressParams.PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
            Write-Progress @progressParams;          
            $Ews = Get-WebServicesVirtualDirectory -Server $server -ADPropertiesOnly:$ADPropertiesOnly;
            $tmpInt = $Ews.InternalUrl
            $tmpEX = $Ews.ExternalUrl
            If ($null -ne $tmpInt)
            {
                [void]$domEntryList.Add($tmpInt.host.toLower());
            } # end if
            else
            {
                writeTolog -LogString ('The inernal URL for the EWS virtual directory on Server '+ $server + ' is empty.') -LogType Warning;
            }; # end else
            if ($null -ne $tmpEX)
            {
                [void]$domEntryList.Add($tmpEX.host.toLower());
            } # end if
            else
            {
                writeTolog -LogString ('The external URL for the EWS virtual directory on Server '+ $server + ' is empty.') -LogType Warning;
            }; # end else

            $tmpInt = ""
            $tmpEX = ""
        }; # end foreach

        Write-Verbose "Collecting namespace information for Mapi..."
        writeTolog -LogString 'Collecting namespace information for Mapi' -LogType Info;
        foreach ($server in $srvList)
        {
            $progressBarIteration++;
            $progressParams.Status=('Mapi on server ' + $server);
            $progressParams.PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
            Write-Progress @progressParams;               
            $mapi = Get-MapiVirtualDirectory -Server $server -ADPropertiesOnly:$ADPropertiesOnly;
            $tmpInt = $mapi.InternalUrl
            $tmpEX = $mapi.ExternalUrl
            If ($null -ne $tmpInt)
            {
                [void]$domEntryList.Add($tmpInt.host.toLower());
            } # end if
            else
            {
                writeTolog -LogString ('The inernal URL for the Mapi virtual directory on Server '+ $server + ' is empty.') -LogType Warning;
            }; # end else
            if ($null -ne $tmpEX)
            {
                [void]$domEntryList.Add($tmpEX.host.toLower());
            } # end if
            else
            {
                writeTolog -LogString ('The external URL for the Mapi virtual directory on Server '+ $server + ' is empty.') -LogType Warning;
            }; # end else

            $tmpInt = ""
            $tmpEX = ""
        }; # end foreach

        # get outlookanywhere url
        Write-Verbose "Collecting namespace information for OutlookAnywhere..."
        writeTolog -LogString 'Collecting namespace information for OutlookAnywhere' -LogType Info;
        foreach ($server in $srvList)
        {
            $progressBarIteration++;
            $progressParams.Status=('OutlookAnywhere on server ' + $server);
            $progressParams.PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
            Write-Progress @progressParams;                           
            $OA = Get-OutlookAnywhere -Server $server -ADPropertiesOnly:$ADPropertiesOnly;          
            $tmpInt = $OA.InternalHostName
            $tmpEX = $OA.ExternalHostName

            If ($null -ne $tmpInt )
            {
                [void]$domEntryList.Add([string]$tmpInt);
            } # end if
            else
            {
                writeTolog -LogString ('The inernal URL for the OutlookAnywhere virtual directory on Server '+ $server + ' is empty.') -LogType Warning;
            }; # end else
       
            if ($null -ne $tmpEX)
            {            
                [void]$domEntryList.Add([string]$tmpEx);               
            } # end if
            else
            {
                writeTolog -LogString ('The external URL for the OutlookAnywhere virtual directory on Server '+ $server + ' is empty.') -LogType Warning;
            }; # end else            
        }; #end foreach
    
        Write-Verbose "Checking accepted domains..."
        writeTolog -LogString 'Checking accepted domains' -LogType Info;        
        if (!($PSBoundParameters.ContainsKey('SkipAcceptedDomains'))) 
        {            
            $progressBarIteration++;
            Write-Progress @progressParams;
            try {
                $acceptedDomainList=([string](Get-AcceptedDomain).DomainName).split(' ');                
                $progressParams.PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
                foreach ($domain in $acceptedDomainList)
                {
                    $progressParams.status=('Accepted domain: ' + $domain.ToLower())
                    [void]$domEntryList.Add('autodiscover.'+$domain.ToLower());
                }; # end foreach
            } # end try
            catch {
                writeToLog -LogString 'Failed to add accepted domains.' -LogType Error;
                writeTolog -LogString ($_.Exception.Message) -LogType Error;
            }; # end catch  
            Write-Progress @progressParams;    
            Write-Progress $progressParams.Activity -Completed;  
        }; # end if
        return [System.Linq.Enumerable]::Distinct([string[]]$domEntryList);
    } # end try
    catch
    {
        return $false
    } # end catch

} # end function getURLList

function createNewTable
{
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)][string]$TableName,
      [Parameter(Mandatory = $true, Position = 1)][array]$FieldList
     )
    
    $tmpTable = New-Object System.Data.DataTable $TableName; # init table   
    $fc=$FieldList.Count;
        
    for ($i=0;$i -lt $fc;$i++)
    {
        if ((!($null -eq $FieldList[$i][1])) -and ($FieldList[$i][1].GetType().name -eq 'runtimetype'))
        {
            [void]($tmpTable.Columns.Add(( New-Object System.Data.DataColumn($FieldList[$i][0],$FieldList[$i][1])))); # add columns to table
        } # end if
        else
        {
            [void]($tmpTable.Columns.Add(( New-Object System.Data.DataColumn($FieldList[$i][0],[System.String])))); # add columns to table
        }; #end else
    }; #end for
    
    return ,$tmpTable;
}; # end createNewTable

function setMSXSearchBase
{
    try {
        $pL=@{
            LDAPFilter='(objectClass=msExchOrganizationContainer)';
            SearchBase=(([ADSI]('LDAP://'+('CN=Microsoft Exchange,CN=Services,'+([ADSI]"LDAP://RootDSE").configurationNamingContext.Value))).distinguishedName)[0];
            FieldList=@('cn');            
            SearchScope='One';
        }; # end Pl
        $Script:msxOrgRootStr = ((queryADForObjects @pL).properties.adspath).TrimStart('LDAP://');
        $Script:msxAdmGroupSearchRootStr=[system.string]::Concat(('CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,',$msxOrgRootStr));
        $script:msxServerSearchRootStr=[system.string]::Concat('CN=Servers,',$Script:msxAdmGroupSearchRootStr);
    } # end try
    catch {
        writeToLog -LogString 'Failed to query the path to the Exchange organization.' -LogType Error;
        writeTolog -LogString ($_.Exception.Message) -LogType Error;
    }; # end catch 
}; # end funciton setMSXSearchBase
function getNewestExchangeServer
{
    try
    { 
        $pL=@{
            LDAPFilter='(objectClass=msExchExchangeServer)';  
            SearchBase=$Script:msxServerSearchRootStr;
            FieldList=@('cn','serialnumber','msexchserversite','msexchinstallpath','networkAddress');
            SearchScope='One';
        }; # end Pl
        writeTolog -LogString 'Querying AD for Exchange servers' -LogType Info;
        $serverList = queryADForObjects @pL; # get list of Exchange servers
        if ($null -ne $serverList)
        {
            writeTolog -LogString (($serverList.count).ToString()+' Exchange servers found.') -LogType Info;  
            writeTolog -LogString ('Exchange server list: ' + $serverList.Properties.cn);   
            $Script:ExchangeServerList = [System.Collections.ArrayList]::new();        
            writeTolog -LogString 'Retrieving FQDN for Exchange servers';
        } # end if
        else {
            #writeTolog -LogString ('No Exchange server found') -LogType Warning;
        }; # end else
        $msxTable = createNewTable -TableName 'MSXServers' -FieldList @(@('ServerName',[System.String]),@('Version',[System.String]),@('ADSite',[System.String]),@('InstallPath',[System.String]));
        
        foreach ($entry in $serverList)
        {
            try {
                $serverFQDN=@(($entry.properties.networkaddress).Where({$_.StartsWith('ncacn_ip_tcp:')}).split(':'))[1].ToLower(); # get FQDN from Exchange server
                [void]($msxTable.Rows.Add([string]$serverFQDN,[string]$entry.Properties.serialnumber.trim('Version ').replace('(','').trim(')').replace(' Build ','.'),[string]$entry[0].Properties.msexchserversite.split(',')[0].replace('CN=',''),[string]$entry[0].Properties.msexchinstallpath));
                [void]$Script:ExchangeServerList.Add($serverFQDN);
            } # end try
            catch {
                writeTolog -LogString ('Failed to add server ' + $entry.Properties.adspath + ' to table.') -LogType Error;
                writeTolog -LogString ($_.Exception.Message) -LogType Error;
                if ($script:ExportFaultyServerData)
                {
                    $fileName=[System.IO.Path]::Combine($Script:logDir,(([System.DateTime]::now.Ticks.ToString())+'.xml'));
                    $entry | Export-Clixml -Path ($fileName);
                    writeTolog ('Exported data of Exchange server to ' +  $fileName) -ShowInfo;
                }; # end if
            }; # end catch            
        }; # end foreach
        # get an Exchange servers with the highest version
        if ($msxTable.rows.count -eq 0) 
        {
            return; # return if no Exchange server found
        }; # end if

        try {
            writeTolog -LogString 'Searching AD site info for local computer.';
            $hostName=[System.Net.Dns]::GetHostByName('').HostName;
            $serverNotValidated=$true;
            $computerSite=([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()).name; # get AD site of computer
            writeTolog -LogString ('Local computer (' + $hostName + ') is member of AD site ' + $computerSite);
        } # end try
        catch {
            writeTolog -LogString ('Failed to find AD site for computer.') -LogType Error;
            writeTolog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch
        
        writeTolog -LogString 'Validating servers';
        while ($serverNotValidated)
        {
            $rv=$msxTable.Select(("[Version] = MAX([Version]) AND [ServerName] = '$hostName'")); # verify if the script is running on an Exchange server (highest version)
            if ($rv.count -eq 0) # if script does not run on an Exchange server
            {
                try {
                    $Rv=(($msxTable.Select(("[Version] = MAX([Version]) AND [ADSite]='$computerSite'"))) | Get-Random); # get a random Exchange server from server list with highest version number and in AD site of computer
                } # end try
                catch {
                    $rv=(($msxTable.Select('[Version] = MAX([Version])')) | Get-Random); # pick up an any Exchange server with the highest version number
                }; # end catch
            }; # end if
            if (testPort -ComputerName ($rv.ServerName) -Port 444)
            {
                $serverNotValidated=$false;
            } # end if
            else
            {
                writeTolog -LogString ('Server ' + ($rv.ServerName + ' not operational')) -LogType Warning;
                $msxTable.rows.Remove($rv); # remove server if not healthy
            }; # end else        
        }; # end while
        writeTolog -LogString 'Finished validating servers';    
        return ,$rv.ServerName;
    } # end try
    catch
    {
        writeTolog -LogString 'Failed to locate Exchange in current forest.' -LogType Warning;
        if ($script:ExportFaultyServerData)        
        {
            writeTolog -LogString ('Debug info: ' + ($_.Exception.Message)) -LogType Info -ShowInfo;
        }; # end if
        return $false;
    }; # end catch
}; # end function getNewestExchangeServer

function loadMSXCmdlets
{
[cmdletbinding()]
param([ref]$ConMsxSrv)
        
    if (! (Get-Command -Name 'Get-Mailbox' -ErrorAction SilentlyContinue))
    {
        try
        {                                               
            if ($ExchangePSTarget = getNewestExchangeServer)
            {
                writeTolog -LogString 'Cleaning up existing sessions to Exchange servers.' -LogType Info;             
                foreach ($psSession in ((Get-PSSession).Where({$_.ConfigurationName -eq 'Microsoft.Exchange'})))
                {
                    Remove-PSSession -Session $psSession;
                }; # end foreach                
                
                try {
                    if ($script:ImportAllExchangeCmdlets -eq $true)
                    {
                        $Script:CommandsToImport='*';
                    } # end if
                    else {
                        try {
                            $Script:CommandsToImport=Get-Content -Path ([System.IO.Path]::Combine($Script:cfgDir,'CmdletsToImport.list'));
                        } # end try
                        catch {
                            $Script:CommandsToImport='*';
                            writeTolog -LogString ('Failed to import the list of required Exchange cmdlets, loading all cmdlets.') -LogType Warning;
                            writeTolog -LogString ($_.Exception.Message) -LogType Warning;
                        }; # end catch                        
                    }; # end if                    
                    $sessionParams=@{
                        ConfigurationName='Microsoft.Exchange';
                        ConnectionUri=('http://' + $ExchangePSTarget + '/PowerShell/');
                        Authentication='Kerberos';
                        ErrorAction='Stop';                        
                    }; # end sessionParams
                    $Script:ExSession= New-PSSession @sessionParams;
                    $ConMsxSrv.Value=$Script:ExSession.ComputerName; 
                    $sessionParams=@{
                        CommandName=$Script:CommandsToImport;
                        DisableNameChecking=$true;
                        Session=$Script:ExSession;
                    }; # end sessionParams
                    writeTolog -LogString ('Loading Exchange cmdlets from server ' + $ExchangePSTarget) -LogType Info;
                    $Script:SessionModule=Import-PSSession @sessionParams;
                    $numOfMSXCmdlets=(Get-module -name ($Script:SessionModule.name)).ExportedCommands.count;
                    writeTolog ('Imported ' + $numOfMSXCmdlets.ToString() + ' Exchange cmdlets from server ' +$Script:ExSession.ComputerName);                    
                } # end try
                catch {
                    writeTolog -LogString ('Failed to load the Exchange cmdlets.') -LogType Error;
                    writeTolog -LogString ($_.Exception.Message) -LogType Error;
                    return $false;
                }; # end catch                             
            } # end if
            else {
                writeTolog -LogString 'No Exchange server found.' -LogType Warning;
                $ConMsxSrv.Value='--- no Exchange server found ---'
                return $false;
            }; # end else
        } # end try
        catch
        {
            writeTolog -LogString ('Failed to load the Exchange PowerShell cmdlets.') -LogType Error;
            writeTolog -LogString ($_.Exception.Message) -LogType Error;
            $ConMsxSrv.Value='--- no Exchange server found ---'
            return $false;
        }; # end catch      
    } # end if
    else
    {
        $ConMsxSrv.Value=(([System.Net.Dns]::GetHostByName('').HostName).ToLower());
    }; # end if
    if (($PSVersionTable.PSEdition -ne 'core') -and ($EMS=Get-PSSnapin -Registered -Name 'Microsoft.Exchange.Management.PowerShell.E2010' -ErrorAction SilentlyContinue))
    {
        $script:MSXScriptDir=(Join-Path -Path ([system.io.path]::GetDirectoryName($ems.ApplicationBase)) -ChildPath 'scripts\');
    } # end if
    else {
        $script:MSXScriptDir=$null;
    }; # end else
    return $true;
}; # end function loadMSXCmdlets

function generateCertRequest
{
[cmdletbinding()]
param([Parameter(Mandatory = $false, Position = 0)][string]$ServerName=$__OPX_ModuleData.ConnectedToMSXServer,
      [Parameter(Mandatory = $false, Position = 1)][array]$ServicesList=@('IMAP','POP','IIS','SMTP'),
      [Parameter(Mandatory = $true, Position = 2)][string]$RequestUNCFilePath,
      [Parameter(Mandatory = $false, Position = 3)][string]$FriendlyName='Exchange Server Certificate',
      [Parameter(Mandatory = $false, Position = 4)][string]$SubjectName,
      [Parameter(Mandatory = $false, Position = 5)][array]$DomainList,
      [Parameter(Mandatory = $false, Position = 6)][ValidateSet('RequestAndInstall','RequestOnly','InstallOnly')][string]$RequestType='RequestAndInstall',
      [Parameter(Mandatory = $false, Position = 7)][string]$CAServerFQDN,
      [Parameter(Mandatory = $false, Position = 8)][string]$CertificateAuthorityName,
      [Parameter(Mandatory = $false, Position = 9)][string]$CertificateTemplateName,
      [Parameter(Mandatory = $false, Position = 10)][string]$CertificateFilePath,
      [Parameter(Mandatory = $false, Position = 11)][switch]$EnableCertificate=$false,
      [Parameter(Mandatory = $false, Position = 12)][int]$KeySize=2048,
      [Parameter(Mandatory = $false, Position = 13)][string]$DomainController
     )
         
    if ($RequestType -in ('RequestAndInstall','RequestOnly'))
    {
        if (!($PSBoundParameters.ContainsKey('SubjectName') -and $PSBoundParameters.ContainsKey('DomainList')))
        {
            throw [System.ArgumentException]'The parameters SubjectName and DomainList are requiered'
        }; #end if
        $msxCertReq=@{
            Server=$ServerName;
            GenerateRequest=$True;
            FriendlyName=$FriendlyName;
            PrivateKeyExportable=$True;
            SubjectName=$SubjectName;
            DomainName = $DomainList;
            ErrorAction='Stop'; 
            KeySize=$KeySize;  
        } # end msxCertRequ
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $msxCertReq.Add('DomainController',$DomainController);
        }; # end if

    #  create cert request and store it to a file
        try
        {
            $msg=('Generating certificate request on server ' + $ServerName)
            Write-Verbose $msg;
            writeTolog -LogString $msg -LogType Info;
            $CertReq = New-ExchangeCertificate  @msxCertReq;
        } # end try
        catch
        {
            $msg=('Certificate request on server ' + $ServerName + ' failed.')
            writeTolog -LogString $msx -LogType Error;
            writeTolog -LogString ($_.Exception.Message) -LogType Error;
            return;
        } # end catch
        try
        {
            Write-Verbose 'Writing request to disk';
            writeTolog -LogString 'Writing request to disk' -LogType Info;
            Set-Content -Path $RequestUNCFilePath -Value $CertReq -ErrorAction Stop;
            Write-Verbose ('Request file written to ' + $RequestUNCFilePath);
            writeTolog -LogString ('Request file written to ' + $RequestUNCFilePath) -LogType Info;
        } # end try
        catch
        {
            writeTolog -LogString  ('Failed to write the request to ' + $RequestUNCFilePath) -LogType Error;
            writeTolog -LogString ($_.Exception.Message) -LogType Error;
            return;
        } # end catch
    }; # end if 'RequestAndInstall','RequestOnly'

    if (!($PSBoundParameters.ContainsKey('CertificateFilePath')))
    {
        $CertificateFilePath=([system.io.path]::Combine(([System.IO.Path]::GetDirectoryName($RequestUNCFilePath)),'Certnew.cer'));
    }; # end if    
    if ($RequestType -in 'RequestAndInstall','InstallOnly')
    {        
        $cRequestDone=$false
        if (($PSBoundParameters.ContainsKey('CAServerFQDN')) -and ($PSBoundParameters.ContainsKey('CertificateAuthorityName')))
        {
            writeTolog -LogString ('Running certreq -submit -config ' + ($CAServerFQDN+'\'+$CertificateAuthorityName) + ' -attrib ' + ('CertificateTemplate:'+$CertificateTemplateName) + ' ' + $RequestUNCFilePath + ' ' + $CertificateFilePath)
            $LASTEXITCODE=0; # reset LASTEXITCODE
            $requestInfo=certreq -submit -config ($CAServerFQDN+'\'+$CertificateAuthorityName) -attrib ('CertificateTemplate:'+$CertificateTemplateName) $RequestUNCFilePath $CertificateFilePath;
            writeTolog -LogString ('Certreq response: ' + ($requestInfo -join ' ')) -ShowInfo;
            if ($LASTEXITCODE -ne 0)
            {
                writeTolog -LogString 'The automatic certificate request failed.' -LogType Warning;                
            } # end if
            else {
                $cRequestDone=$true;
            }; # end else
        }; # end if
        
        if ($cRequestDone -eq $false)
        {
            writeTolog -LogString 'Waiting for the request to complete' -LogType Info;
            if ($RequestType -eq 'InstallOnly' -and ($cRequestDone -eq $false))
            {
                Write-Verbose 'Waiting for the request to complete';
                $tmp=Read-Host -Prompt ('Please enter the path to the certificate file or leave it blank for ' + $CertificateFilePath);         
            } # end if
            else {
                Write-Verbose 'Waiting for the request to complete';
                $tmp=Read-Host -Prompt ('Request file prepared. Please enter the path to the certificate file or leave it blank for ' + $CertificateFilePath);          
            }; # end else
            if (!([System.String]::IsNullOrEmpty($tmp)))
            {
                $CertificateFilePath=$tmp;
            }; # end if
        }; # end if
        
        if (! (Test-Path -Path $CertificateFilePath -PathType Leaf))
        {
            writeTolog -LogString ('Certificate ' + $CertificateFilePath + ' not found.') -LogType Warning;
            return $false;  # if cert file not found, return FALSE
        }; # end if
    # import the cert to the requesting server
        try
        {
            Write-Verbose ('Importing certificate');
            writeTolog -LogString 'Importing certificate' -LogType Info;
            $msxCertImport=(Import-ExchangeCertificate -Server $ServerName -FileName $CertificateFilePath -PrivateKeyExportable:$true -ErrorAction 'Stop');
            $Thumbprint = $msxCertImport.Thumbprint
            if($EnableCertificate.IsPresent)
            {
                Write-Verbose ('Enabeling certificate for services: ' +$($ServicesList -join ','));
                writeTolog -LogString ('Enabeling certificate for services: ' +$($ServicesList -join ','));
                $paramList=@{
                    Thumbprint=$Thumbprint;
                    Server=$ServerName;
                    Services=$servicesList;
                    ErrorAction='Stop';
                }; # end parmaList
                if ($PSBoundParameters.ContainsKey('DomainController'))
                {
                    $paramList.Add('DomainController',$DomainController);
                }; # end if
                Enable-ExchangeCertificate @paramList;
                #>            
            }; # end if
        } #end try
        catch
        {
            writeTolog -LogString 'Failed to install the certificate' -LogType Info;
            writeTolog -LogString ($_.Exception.Message) -LogType Error;
            return $false;
        }; # end catch    
        return $Thumbprint;
    } # end if 'RequestAndInstall','InstallOnly'
    else {
        writeTolog -LogString ('Request file ' + $RequestUNCFilePath + ' prepared.') -LogType Info -ShowInfo;
    }; # end else
}; # end function generateCertRequest

function testPort
{
[CmdLetBinding()]            
param([Parameter(Mandatory = $true, Position = 0)][string]$ComputerName,
      [Parameter(Mandatory = $true, Position = 1)][int]$Port, 
      [Parameter(Mandatory = $false, Position = 2)][int]$TcpTimeout=100    
     )

    begin {        
    }; #end begin
    
    process {
        writeTolog -LogString ('Testing port ' + $Port.ToString() + ' on computer ' + $computerName);
        $TcpClient = New-Object System.Net.Sockets.TcpClient
        $Connect = $TcpClient.BeginConnect($ComputerName, $Port, $null, $null)
        $Wait = $Connect.AsyncWaitHandle.WaitOne($TcpTimeout, $false)
        if (!$Wait) 
        {
	        writeTolog -LogString ('Server ' + $computerName + ' failed to answer on port ' + $Port.ToString()) -LogType Warning;
            return $false;
        } # end if
        else 
        {	        
	        return $true;
        }; # end else        
    } # end process

    end {        
        $TcpClient.Close();
        $TcpClient.Dispose();
    } # end END

}; # end function testPort

function formatURL
{
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)][hashtable]$URLHash,
      [Parameter(Mandatory = $true, Position = 1)][string]$ProtocolString
     )

    [array]$entryList=@(($URLHash.Keys).Where({$_ -in ('InternalUrl','ExternalUrl')})); # filter  int/ext URLs
    $entryListCount=$entryList.count;

    for ($i=0;$i -lt $entryListCount;$i++)
    {
        if (!([system.string]::IsNullOrEmpty($URLHash.($entryList[$i])))) # check if the value is NULL, if NULL don't modify
        {
            $URLHash.($entryList[$i]) = [string]([system.uri]::New(([uri]('https://' + $URLHash.($entryList[$i]))) ,$ProtocolString));  # build uri     
        } # end if
        else
        {
            $URLHash.($entryList[$i])=$null;
        }; # end else
    }; # end for
    [void]$URLHash.Add('ErrorAction','Stop');
}; # end function format URL


function istDCOnline
{  
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)][string]$DomainController
     )      
    
    if (testPort -ComputerName $DomainController -Port 636 ) # verify if DC is available
    {
        return $true;
    } # end if
    else {
        writeTolog -LogString ('DC ' + $DomainController + ' failed to answer on port 636') -LogType Warning;
        return $false;
    }; # end else
}; # end funciton getDCList

function testIfFQDN
{
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)][string]$ComputerName
     )
    
    try {
        $tmpName=[System.Net.Dns]::GetHostByName($ComputerName);
        if (!(($tmpName.hostname).contains('.')))
        {
            try {
                $tmpName=[System.Net.Dns]::GetHostByAddress($tmpName.AddressList[0].IPAddressToString);
            } # end try
            catch {
                writeTolog -LogString ('Failed to convert the computer name ' + $ComputerName + ' to a fully qualified domain name (FQDN).') -LogType Warning;
                writeToLog -LogString ($_.Exception.Message) -LogType warning;
            }; # catch            
        }; # end if
        return $tmpName.HostName;
    } # end try
    catch {
        writeTolog -LogString ('Failed to resolve the computer name ' + $ComputerName) -LogType Warning;
        return $false;
    }; # end catch
    
}; #end function 
    
function resolveURLs 
{
[cmdletbinding()]
param([Parameter(Mandatory = $true, Position = 0)][hashtable]$ParameterHash
     )
    [void]($urlList=[System.Collections.ArrayList]::new());
    Clear-DnsClientCache;
    foreach($item in $ParameterHash.Keys)
    {
        if (($ParameterHash.$item).count -gt 0)
        {
            foreach ($subItem in $ParameterHash.$item.Keys)
            {
                if ($subitem -in ('InternalUrl','ExternalUrl','InternalHostname','ExternalHostname','AutoDiscoverServiceInternalUri'))
                {
                    $entry=$ParameterHash.$item.$subItem
                    if (!([system.string]::IsNullOrEmpty($entry)))
                    {
                        [void]($urlList.Add($entry)); # add to list
                    }; # is URL/hostname not empty string
                }; # end if is URL or hostname
            }; # end foreach subitem
        }; # end if parameter count
    }; # end foreach entry in hash    
    $urlList=@([System.Linq.Enumerable]::Distinct([string[]]$urlList));
    foreach ($url in $urlList)
    {
        try {
            [System.Net.Dns]::Resolve($url);
        } # end try
        catch {
            writeTolog -LogString ('Failed to resolve the FQDN ' + $url) -LogType Warning;
            writeToLog -LogString ($_.Exception.Message) -LogType Warning;
        }; # end catch
    }; # end foreach
}; # end function resolveURLs


function writeTolog
{
    [CmdLetBinding()]
    param([Parameter(Mandatory = $true, Position = 0)]
          [string]$LogString,
          [Parameter(Mandatory = $false, Position = 2)]
          [ValidateSet('Info','Warning','Error')] 
          [string]$LogType="Info",
          [Parameter(Mandatory = $false, Position = 3)]
          [string]$ComputerName=([System.Net.Dns]::GetHostEntry('').hostname),                                   
          [Parameter(Mandatory = $false, Position = 8)]
          [string]$CallingCmdlet=$Script:CallingCmdlet,
          [Parameter(Mandatory = $false, Position = 9)]
          [string]$CallingID=$Script:CallingID,
          [Parameter(Mandatory = $false, Position = 10)]
          [switch]$ShowInfo=$false,
          [Parameter(Mandatory = $false, Position = 12)]
          [switch]$SupressScreenOutput=$false
         )
    
    begin {        
        $LogDateTime=([System.DateTimeOffset]::($script:LogTimeUTC)).ToString();        
        $logFileName=$script:CurrentLogFileName;
    }; # end begin

    process {
        switch ($script:LoggingTarget)
        {
            'File'          {
                $lt=$LogDateTime.split(' ')[0,1] -join ' ';
                $utc=$LogDateTime.split(' ')[2]
                $csvExport=New-Object -TypeName PSObject -Property ([ordered]@{
                    DateTime=$lt;
                    UtcOff=$utc;
                    Computer=$ComputerName;
                    UserName= $script:LogedOnUser;
                    LogType=$LogType;
                    CallingCmdlet=$CallingCmdlet;
                    LogMessage=$LogString;
                    CallingID=$Script:CallingID;
                }); # end csvExport
                $csvParams=@{
                    Path = $logFileName;
                    Append=$true;
                    Delimiter=$script:LogFileCsvDelimiter;
                    NoTypeInformation=$true;
                    WhatIF=$false;
                }; # end csvParams
                try {
                    $csvExport | Export-Csv @csvParams;
                } # end try
                catch {
                    Write-Host ('ERROR: Failed to write to log file ' + $logFileName) -BackgroundColor Black -ForegroundColor Red;
                    Write-Host ($_.Exception.Message) -BackgroundColor Black -ForegroundColor Red;
                }; # end catch
                break;
            }; # end file
            'None'       {
                # log nothing
            }; # None
        }; # end switch LoggingTarget
         
        if ((($LogType -in ('Warning','Error')) -or $ShowInfo.IsPresent) -and (! $SupressScreenOutput.isPresent))
        {
            switch ($LogType)
            {
                'Warning'   {
                    Write-Warning ('!!! ' + $LogString);
                }; # end warning
                'Error'     {
                    Write-Host ('ERROR: ' +$LogString) -ForegroundColor Red -BackgroundColor Black;
                }; # end error
                'Info'      {
                    $logParams=@{
                        Object=$LogString;
                    };                
                    Write-Host @logParams;                                
                }; # end info
            }; # end switch
        }; # end if
    }; # end process

    end {

    }; # end END
}; # end function wirteToLog

function writeCmdletInitToDataLog
{
[CmdLetBinding()]
Param([Parameter(Mandatory = $true, Position = 0)][array]$CallStack,
      [Parameter(Mandatory = $false, Position = 1)][switch]$ExitLogEntry=$false,
      [Parameter(Mandatory = $false, Position = 4)][System.Collections.Generic.Dictionary`2+Enumerator[System.String,System.Management.Automation.ParameterMetadata]]$ParameterList
     )
    
    begin 
    {                        
        if ($ExitLogEntry.IsPresent -eq $false)
        {
            $LogDateTime=([System.DateTimeOffset]::($script:LogTimeUTC)).ToString();        
            $script:CurrentLogFileName=([System.IO.Path]::Combine($script:LogFileDir,('Log_')+($LogDateTime.Split(' ')[0].replace('/','-').replace('.','-'))+'.log'));
        }; # end if
        $cmdCount=2+[int]('<ScriptBlock>' -in (Get-PSCallStack).Command);
        if ((! $ExitLogEntry.IsPresent) -and (((Get-PSCallStack).Command.count -eq $cmdCount) -or $NewGUIDForEverycmdletCall))
        {
            $Script:CallingID=[guid]::NewGuid().Guid;
            $Script:CallingCmdlet=$CallStack[0].Command;
        } # end if
        $csArgumentString=$CallStack[0].Arguments.TrimStart('{').TrimEnd('}')
        $tmp=$csArgumentString.Split(',');
        for ($i=0;$i -lt $tmp.Count;$i++)
        {
            if ($tmp[$i].EndsWith('='))
            {
                $Tmp[$i]+='False';
            }; # end if
        }; # end for
        $csArgumentString=$tmp -join ',';
        if ($PSBoundParameters.ContainsKey('ParameterList')) # check for parameters with default value
        {                   
            $excludeList=[System.Collections.ArrayList]::new();
            [void]$excludeList.AddRange(@([System.Management.Automation.PSCmdlet]::OptionalCommonParameters));
            [void]$excludeList.AddRange(@([System.Management.Automation.PSCmdlet]::CommonParameters));            
            foreach ($line in ($csArgumentString.split(',').Trim()))
            {
                [void]$excludeList.Add($line.split('=')[0]);
            }; # end foreach
            $pList=@([System.Linq.Enumerable]::Except([string[]]@($ParameterList).key, [string[]]$excludeList)); # get diff between parameterList and excludeList (items in parameterList and not in excludeList)
            foreach ($p in $pList)
            {
                $csArgumentString += ', ' + $p + '=' + (Get-Variable -Name ($p) -Scope 1 -ValueOnly); # add parameter and value
            }; # end foreach           
        } # end if
        
        $ParamList = @{
            LogString = ('Running ' + $CallStack[0].Command + ' ' + $csArgumentString.TrimStart(', '))
            LogType = 'Info'           
        } # end paramList  
        #[void](writeToLog @ParamList)

        if ($ExitLogEntry)
        {
            $ParamList.logString = ('Finished running command ' + $CallStack[0].Command + ' ' + $csArgumentString.TrimStart(', '))
            [void](writeToLog @ParamList)
            $Script:CallingCmdlet=$CallStack[1].Command;            
        }   # end if 
        else {
            $Script:CallingCmdlet=$CallStack[0].Command;          
            [void](writeToLog @ParamList)
        } # end if
        #>
     } # end begin
     process{};
     end {       
     } # end END
} # end function writeCmdletInitToDataLog


function listVirtualDirectories
{
[CmdLetBinding()]
Param([Parameter(Mandatory = $true, Position = 0)][array]$ServiceList,
      [Parameter(Mandatory = $true, Position = 1)][array]$ServerList,
      [Parameter(Mandatory = $false, Position = 2)][switch]$OnlyVerifyUrls=$false,
      [Parameter(Mandatory = $false, Position = 3)][hashtable]$AdditionalCfg=@{},
      [Parameter(Mandatory = $false, Position = 4)][string]$DomainController
     )
    
    $srvCount=$ServerList.count;
    $progressBarMaxIterations=($ServiceList.Count * $srvCount);
    for ($i=0;$i -lt $srvCount; $i++)
    {
        if ($ServerList[$i].Contains('.'))
        {
            $ServerList[$i]=($ServerList[$i].split('.'))[0]; # remove domain
        }; # end if
    }; # end for
    # init arrays and var
    $formatList=[System.Collections.ArrayList]::new();
    $tmpVdir=[System.Collections.ArrayList]::new();
    $progressBarIteration=0;
    $svcCount=$serviceList.Count;
    $msg='Collecting data from selected server(s), please wait'
    Write-Verbose ($msg + '...');
    writeTolog -LogString $msg;
    $DCParam=@{
        ErrorAction='Stop';
    }; # end DCParam
    if ($PSBoundParameters.ContainsKey('DomainController'))
    {
        $DCParam.Add('DomainController',$DomainController)
    }; # end if

    try {
        for ($i=0;$i -lt $svcCount; $i++)
        {                
            switch ($serviceList[$i])
            {
                'OWA' {
                    for ($k=0;$k -lt $srvCount; $k++)
                    {
                        $progressBarIteration++;
                        $progressParams=@{
                            Activity=('Collecting virtual directory information for service ' + $ServiceList[$i] +' on server ' + $serverList[$k]);
                            Status=('Server ' + ($k+1).ToString() + '/' + $srvCount.ToString());
                            PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
                        }; # end progressParams
                        $msg=('Collecting OWA virtual directory data on server ' + $serverList[$k])
                        Write-Verbose ($msg + '...');
                        writeTolog -LogString $msg;
                        Write-Progress @progressParams
                        [void]$tmpVdir.Add((Get-OwaVirtualDirectory -Server $serverList[$k] -ADPropertiesOnly:$ADPropertiesOnly.IsPresent @DCParam));
                    }; # end for server
                    $tmp=@('Identity','InternalURL','ExternalURL');
                    if (($AdditionalCfg.ContainsKey(($ServiceList[$i]))) -and ($AdditionalCfg.($ServiceList[$i]).Count -gt 0))
                    {
                        $tmp+=@($AdditionalCfg.($ServiceList[$i]).keys);
                    }; # end if
                    [void]$formatList.Add(@($tmp));
                    if (! ($OnlyVerifyUrls.IsPresent))
                    {
                        Write-Output ('Displaying data for service ' + ($ServiceList[$i]));
                        $tmpVdir | Format-List $tmp;
                    }; # end if display output
                    break;
                }; # end owa
                'ECP' {
                    [void]$tmpVdir.Clear();
                    for ($k=0;$k -lt $srvCount; $k++)
                    {
                        $progressBarIteration++;
                        $progressParams=@{
                            Activity=('Collecting virtual directory information for service ' + $ServiceList[$i] +' on server ' + $serverList[$k]);
                            Status=('Server ' + ($k+1).ToString() + '/' + $srvCount.ToString());
                            PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
                        }; # end progressParams
                        $msg=('Collecting ECP virtual directory data on server ' + $serverList[$k])
                        Write-Verbose ($msg + '...');
                        writeTolog -LogString $msg;
                        Write-Progress @progressParams
                        [void]$tmpVdir.Add((Get-EcpVirtualDirectory -Server $serverList[$k] -ADPropertiesOnly:$ADPropertiesOnly.IsPresent @DCParam));
                    }; # end for server
                    $tmp=@('Identity','InternalURL','ExternalURL');
                    if (($AdditionalCfg.ContainsKey(($ServiceList[$i]))) -and ($AdditionalCfg.($ServiceList[$i]).Count -gt 0))
                    {
                        $tmp+=@($AdditionalCfg.($ServiceList[$i]).keys);
                    }; # end if
                    [void]$formatList.Add(@($tmp));
                    if (! ($OnlyVerifyUrls.IsPresent))
                    {
                        Write-Output ('Displaying data for service ' + ($ServiceList[$i]));
                        $tmpVdir | Format-List $tmp;
                    }; # end if display output   
                    break;                 
                }; # end ECP
                'OAB' {
                    [void]$tmpVdir.Clear();
                    for ($k=0;$k -lt $srvCount; $k++)
                    {
                        $progressBarIteration++;
                        $progressParams=@{
                            Activity=('Collecting virtual directory information for service ' + $ServiceList[$i] +' on server ' + $serverList[$k]);
                            Status=('Server ' + ($k+1).ToString() + '/' + $srvCount.ToString());
                            PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
                        }; # end progressParams
                        $msg=('Collecting OAB virtual directory data on server ' + $serverList[$k])
                        Write-Verbose ($msg + '...');
                        writeTolog -LogString $msg;
                        Write-Progress @progressParams
                        [void]$tmpVdir.Add((Get-OabVirtualDirectory -Server $serverList[$k] -ADPropertiesOnly:$ADPropertiesOnly.IsPresent @DCParam));
                    }; # end for server
                    $tmp=@('Identity','InternalURL','ExternalURL');
                    if (($AdditionalCfg.ContainsKey(($ServiceList[$i]))) -and ($AdditionalCfg.($ServiceList[$i]).Count -gt 0))
                    {
                        $tmp+=@($AdditionalCfg.($ServiceList[$i]).keys);
                    }; # end if
                    [void]$formatList.Add(@($tmp));
                    if (! ($OnlyVerifyUrls.IsPresent))
                    {
                        Write-Output ('Displaying data for service ' + ($ServiceList[$i]));
                        $tmpVdir | Format-List $tmp;
                    }; # end if display output  
                    break;                  
                }; # end OAB
                'Webservices' {
                    [void]$tmpVdir.Clear();
                    for ($k=0;$k -lt $srvCount; $k++)
                    {
                        $progressBarIteration++;
                        $progressParams=@{
                            Activity=('Collecting virtual directory information for service ' + $ServiceList[$i] +' on server ' + $serverList[$k]);
                            Status=('Server ' + ($k+1).ToString() + '/' + $srvCount.ToString());
                            PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
                        }; # end progressParams
                        $msg=('Collecting Webservices virtual directory data on server ' + $serverList[$k])
                        Write-Verbose ($msg + '...');
                        writeTolog -LogString $msg;
                        Write-Progress @progressParams
                        [void]$tmpVdir.Add((Get-WebservicesVirtualDirectory -Server $serverList[$k] -ADPropertiesOnly:$ADPropertiesOnly.IsPresent @DCParam));
                    }; # end for server
                    $tmp=@('Identity','InternalURL','ExternalURL');
                    if (($AdditionalCfg.ContainsKey(($ServiceList[$i]))) -and ($AdditionalCfg.($ServiceList[$i]).Count -gt 0))
                    {
                        $tmp+=@($AdditionalCfg.($ServiceList[$i]).keys);
                    }; # end if
                    [void]$formatList.Add(@($tmp));
                    if (! ($OnlyVerifyUrls.IsPresent))
                    {
                        Write-Output ('Displaying data for service ' + ($ServiceList[$i]));
                        $tmpVdir | Format-List $tmp;
                    }; # end if display output
                    break;
                }; # end Webservices
                'ActiveSync' {
                    [void]$tmpVdir.Clear();
                    for ($k=0;$k -lt $srvCount; $k++)
                    {
                        $progressBarIteration++;
                        $progressParams=@{
                            Activity=('Collecting virtual directory information for service ' + $ServiceList[$i] +' on server ' + $serverList[$k]);
                            Status=('Server ' + ($k+1).ToString() + '/' + $srvCount.ToString());
                            PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
                        }; # end progressParams
                        $msg=('Collecting ActiveSync virtual directory data on server ' + $serverList[$k])
                        Write-Verbose ($msg + '...');
                        writeTolog -LogString $msg;
                        Write-Progress @progressParams
                        [void]$tmpVdir.Add((Get-ActiveSyncVirtualDirectory -Server $serverList[$k] -ADPropertiesOnly:$ADPropertiesOnly.IsPresent @DCParam));
                    }; # end for server
                    $tmp=@('Identity','InternalURL','ExternalURL');
                    if (($AdditionalCfg.ContainsKey(($ServiceList[$i]))) -and ($AdditionalCfg.($ServiceList[$i]).Count -gt 0))
                    {
                        $tmp+=@($AdditionalCfg.($ServiceList[$i]).keys);
                    }; # end if
                    [void]$formatList.Add(@($tmp));
                    if (! ($OnlyVerifyUrls.IsPresent))
                    {
                        Write-Output ('Displaying data for service ' + ($ServiceList[$i]));
                        $tmpVdir | Format-List $tmp;
                    }; # end if display output
                    break;
                }; # end Activesync
                'Mapi' {
                    [void]$tmpVdir.Clear();
                    for ($k=0;$k -lt $srvCount; $k++)
                    {
                        $progressBarIteration++;
                        $progressParams=@{
                            Activity=('Collecting virtual directory information for service ' + $ServiceList[$i] +' on server ' + $serverList[$k]);
                            Status=('Server ' + ($k+1).ToString() + '/' + $srvCount.ToString());
                            PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
                        }; # end progressParams
                        $msg=('Collecting Mapi virtual directory data on server ' + $serverList[$k])
                        Write-Verbose ($msg + '...');
                        writeTolog -LogString $msg;
                        Write-Progress @progressParams
                        [void]$tmpVdir.Add((Get-MapiVirtualDirectory -Server $serverList[$k] -ADPropertiesOnly:$ADPropertiesOnly.IsPresent @DCParam));
                    }; # end for server
                    $tmp=@('Identity','InternalURL','ExternalURL');
                    if (($AdditionalCfg.ContainsKey(($ServiceList[$i]))) -and ($AdditionalCfg.($ServiceList[$i]).Count -gt 0))
                    {
                        $tmp+=@($AdditionalCfg.($ServiceList[$i]).keys);
                    }; # end if
                    [void]$formatList.Add(@($tmp));
                    if (! ($OnlyVerifyUrls.IsPresent))
                    {
                        Write-Output ('Displaying data for service ' + ($ServiceList[$i]));
                        $tmpVdir | Format-List $tmp;
                    }; # end if display output
                    break;
                }; # end Mapi
                'OutlookAnywhere' {
                    [void]$tmpVdir.Clear();
                    for ($k=0;$k -lt $srvCount; $k++)
                    {
                        $progressBarIteration++;
                        $progressParams=@{
                            Activity=('Collecting virtual directory information for service ' + $ServiceList[$i] +' on server ' + $serverList[$k]);
                            Status=('Server ' + ($k+1).ToString() + '/' + $srvCount.ToString());
                            PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
                        }; # end progressParams
                        $msg=('Collecting OutlookAnywhere data on server ' + $serverList[$k])
                        Write-Verbose ($msg + '...');
                        writeTolog -LogString $msg;
                        Write-Progress @progressParams
                        [void]$tmpVdir.Add((Get-OutlookAnywhere -Server $serverList[$k] -ADPropertiesOnly:$ADPropertiesOnly.IsPresent @DCParam));
                    }; # end for server
                    $tmp=@('Identity',
                        'InternalHostname',
                        'ExternalHostname',
                        'ExternalClientsRequireSsl',
                        'InternalClientsRequireSsl',
                        'ClientAuthenticationMethod',
                        'ExternalClientAuthenticationMethod'
                    );
                    if (($AdditionalCfg.ContainsKey(($ServiceList[$i]))) -and ($AdditionalCfg.($ServiceList[$i]).Count -gt 0))
                    {
                        $tmp+=@($AdditionalCfg.($ServiceList[$i]).keys);
                    }; # end if
                    [void]$formatList.Add(@($tmp));
                    if (! ($OnlyVerifyUrls.IsPresent))
                    {
                        Write-Output ('Displaying data for service ' + ($ServiceList[$i]));
                        $tmpVdir | Format-List $tmp;
                    }; # end if display output
                    break;
                }; # end OutlookAnywhere
                'Autodiscover' {
                    [void]$tmpVdir.Clear();
                    for ($k=0;$k -lt $srvCount; $k++)
                    {
                        $progressBarIteration++;
                        $progressParams=@{
                            Activity=('Collecting virtual directory information for service ' + $ServiceList[$i] +' on server ' + $serverList[$k]);
                            Status=('Server ' + ($k+1).ToString() + '/' + $srvCount.ToString());
                            PercentComplete=([math]::min(100,(($progressBarIteration /$progressBarMaxIterations)*100)));
                        }; # end progressParams
                        $msg=('Collecting Autodiscover data on server ' + $serverList[$k])
                        Write-Verbose ($msg + '...');
                        writeTolog -LogString $msg;
                        Write-Progress @progressParams
                        [void]$tmpVdir.Add((Get-ClientAccessService -Identity $serverList[$k]  @DCParam));
                    }; # end for server                    
                    $tmp=@('Identity','AutoDiscoverServiceInternalUri');
                    if (($AdditionalCfg.ContainsKey(($ServiceList[$i]))) -and ($AdditionalCfg.($ServiceList[$i]).Count -gt 0))
                    {
                        $tmp+=@($AdditionalCfg.($ServiceList[$i]).keys);
                    }; # end if
                    [void]$formatList.Add(@($tmp));
                    if (! ($OnlyVerifyUrls.IsPresent))
                    {
                        Write-Output ('Displaying data for service ' + ($ServiceList[$i]));
                        $tmpVdir | Format-List $tmp;
                    }; # end if display output
                    break
                }; # end Mapi
            }; # end switch            
        }; # end for services list
    } # end try
    catch {
        writeTolog -LogString ('Failed to list the services') -LogType Error;
        writeTolog -LogString ($_.Exception.Message) -LogType Error;
    }; # end catch
    
}; # end function listVirtualDirectories

function setComponentState
{
[CmdLetBinding()]
Param([Parameter(Mandatory = $true, Position = 0)][string]$Server,
      [Parameter(Mandatory = $true, Position = 1)][string]$Component,
      [Parameter(Mandatory = $true, Position = 2)][string]$State,
      [Parameter(Mandatory = $true, Position = 3)][string]$Requester,
      [Parameter(Mandatory = $false, Position = 4)][string]$DomainController
     )
    
    try {
        $msg=('Setting component ' + $Component + ' on server ' + $server + ' to ' + $state);
        Write-Verbose $msg;
        writeTolog -LogString $msg;
        $DCParam=@{ErrorAction='Stop'};
        if ($PSBoundParameters.ContainsKey('DomainController'))
        {
            $DCParam.Add('DomainController',$DomainController);
        }; # end if
        Set-ServerComponentState $ServerFQDN -Component $Component -State $State -Requester $Requester @DCParam;
    } # end try
    catch {
        writeToLog -LogType Error -LogString  ('Failed to set the component ' + $Component + ' on server ' + $Server + ' to ' + $State);
        writeTolog -LogString ($_.Exception.Message) -LogType Error;
    }; # end catch    
}; # end function setComponentState

function runScriptFromExScriptsDir
{
[CmdLetBinding()]
Param([Parameter(Mandatory = $true, Position = 0)][string]$ScriptName,
      [Parameter(Mandatory = $true, Position = 1)][hashtable]$ScriptParams
     )

    if (($Script:RunScriptsInJob -eq $false) -and (Get-PSSnapin -Name 'Microsoft.Exchange.Management.PowerShell.E2010' -ErrorAction SilentlyContinue))
    {
        $msg='Saving current locaton.'
        Write-Verbose $msg;
        writeTolog -LogString $msg;
        $tmpLoc = Get-Location
        $msg=('Changing to directory ' + $Script:MSXScriptDir)
        Write-Verbose $msg;
        writeTolog -LogString $msg;
        Set-Location $Script:MSXScriptDir;
        $runJob=$false;
    } # end if
    else {
        $runJob=$true
    }; # end else
    switch ($ScriptName)
    {
        'StartDagServerMaintenance'     {
                if ($runJob)
                {
                    $SB = {
                        Param (
                        [string]$Dir,
                        [string]$serverName,
                        [string]$MoveComment
                        )
                        Set-Location $dir;
                        $DagScriptTesting=$false;                        
                        .\StartDagServerMaintenance.ps1 -serverName $serverName -MoveComment $MoveComment -pauseClusterNode;                        
                    }; # end SB
                    writeToLog -LogString ('Starting job StartDagServerMaintenance');
                    try {
                        $msxJob= Start-Job -ScriptBlock $Sb -ArgumentList $script:MSXScriptDir,$ScriptParams.ServerName,$ScriptParams.MoveComment -ErrorAction Stop;   
                    } # end try
                    catch {
                        writeToLog -LogString ('Failed to run job StartDagServerMaintenance') -LogType Error;
                        writeToLog -LogString ($_.Exception.Message) -LogType Error;
                    }; # end catch                    
                } # end if
                else {
                    try {
                        $DagScriptTesting=$false;
                        writeTolog -LogString 'Starting script StartDagServerMaintenance.ps1';
                        .\StartDagServerMaintenance.ps1 -serverName $ScriptParams.ServerName -MoveComment $ScriptParams.MoveComment -pauseClusterNode;
                    } # end try
                    catch {
                        writeToLog -LogString ('Failed to run script StartDagServerMaintenance.ps1') -LogType Error;
                        writeToLog -LogString ($_.Exception.Message) -LogType Error;
                    }; # end catch
                }; # end else
                break;                
        }; # end start maintenance
        'StopDagServerMaintenance'     {
            if ($runJob)
            {
                $SB = {
                    Param (
                    [string]$Dir,
                    [string]$serverName                   
                    )
                    Set-Location $dir;
                    $DagScriptTesting=$false;                        
                    .\StopDagServerMaintenance.ps1 -serverName $serverName ;                        
                }; # end SB
                writeToLog ('Starting job StopDagServerMaintenance');
                try {
                    $msxJob= Start-Job -ScriptBlock $Sb -ArgumentList $script:MSXScriptDir,$ScriptParams.ServerName -ErrorAction Stop;   
                } # end try
                catch {
                    writeToLog -LogString ('Failed to run job StopDagServerMaintenance') -LogType Error;
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;
                }; # end catch                    
            } # end if
            else {
                try {
                    $DagScriptTesting=$false;
                    writeTolog -LogString 'Starting script StopDagServerMaintenance.ps1';
                    .\StopDagServerMaintenance.ps1 -serverName $ScriptParams.ServerName -MoveComment $ScriptParams.MoveComment -pauseClusterNode;
                } # end try
                catch {
                    writeToLog -LogString ('Failed to run script StopDagServerMaintenance.ps1') -LogType Error;
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;
                }; # end catch
            }; # end else
            break;                
    }; # end start maintenance  
        'RedistributeActiveDatabases'   {
            writeTolog -LogString 'Running script RedistributeActiveDatabases.ps1';
            if ($runJob)
            {
                $SB = {
                    Param (
                    [string]$Dir,
                    [string]$DAGName
                    )
                    Set-Location $dir                       
                    .\RedistributeActiveDatabases.ps1 -DagName $DAGName -BalanceDbsByActivationPreference -ShowFinalDatabaseDistribution -Confirm:$false;                        
                }; # end SB
                writeToLog ('Starting job RedistributeActiveDatabases');
                try {
                    $msxJob= Start-Job -ScriptBlock $Sb -ArgumentList $script:MSXScriptDir,$ScriptParams.DAGName;
                } # end try
                catch {
                    writeToLog -LogString ('Failed to run script RedistributeActiveDatabases') -LogType Error;
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;
                }; # end catch                
            } # end if
            else {
                try {
                    writeTolog -LogString 'Starting script RedistributeActiveDatabases.ps1';
                    .\RedistributeActiveDatabases.ps1 -DagName $ScriptParams.DAGName -BalanceDbsByActivationPreference -ShowFinalDatabaseDistribution -Confirm:$false;
                } # end try
                catch {
                    writeToLog -LogString ('Failed to run script RedistributeActiveDatabases.ps1') -LogType Error;
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;
                }; # end catch
            }; # end else
            break;                
        }; # end RedistributeActiveDatabases
    }; # end switch
    if ($runJob)  # monitor job
    {        
        $monitorJob=$true;
        while ($monitorJob)
        {
            Start-Sleep -Seconds $Script:JobMonitoringIntervalInSeconds;
            if ($msxJob.State -ne 'Running')
            {
                $monitorJob=$false;
            }; # end if
        }; # end while
        Start-Sleep -Seconds 3; 
        If ($msxJob.State -ne 'Completed')
        {
            writeToLog -LogString ('The job for ' +$ScriptName + ' finished with a state of ' + $msxJob.State) -LogType Warning;
        } # end if
        else {
            writeToLog -LogString ('The job for ' +$ScriptName + ' finished with a state of ' + $msxJob.State) -LogType Info;
        }; # end else
        try {
            Receive-Job  -Id $msxJob.Id -ErrorAction Stop;
            Remove-Job -Id $msxJob.Id -ErrorAction Stop;   
        } # end try
        catch {
            writeToLog -LogString ('Failed to cleanup job') -LogType Error;
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch
        
    } # end if
    else {
        Write-Verbose ('Changing to directory ' + $tmpLoc);
        writeTolog -LogString ('Changing to directory ' + $tmpLoc);
        Set-Location $tmpLoc;   
    }; # end else
     
}; # end function

function getAllExchangeServer
{
[CmdLetBinding()]
Param([Parameter(Mandatory = $false, Position = 0)][switch]$ReturnFQDN=$false
     ) 
    $pL=@{
        LDAPFilter='(objectClass=msExchExchangeServer)'; 
        SearchBase=$Script:msxServerSearchRootStr;
        FieldList=@('cn','networkAddress');
        SearchScope='One';
    }; # end Pl
    
    $serverList=[system.Collections.ArrayList]::new();    
    try {
        $resultList=queryADForObjects @pL;
        if ($ReturnFQDN.isPresent)
        {
            foreach ($entry in $resultList)
            {
                try {            
                    [void]$serverList.Add((($entry.properties.networkaddress).Where({$_.StartsWith('ncacn_ip_tcp:')}).split(':'))[1].ToLower());
                }
                catch {
                    writeToLog -LogString ('Failed get FQDN for Exchange server ' + $entry.properties.cn) -LogType Error;
                    writeToLog -LogString ($_.Exception.Message) -LogType Error;
                }; # end catch
            }; # end foreach   
        } # end if
        else {
            [void]$serverList.AddRange($resultList.properties.cn);
        }; # end else        
    } # end try
    catch {
        writeToLog -LogString ('Failed to query AD for Exchange server.') -LogType Error;
        writeToLog -LogString ($_.Exception.Message) -LogType Error;
        return $false;
    }; # end catch 
    return $serverList;   
}; # end function getAllExchangeServer
function getAllCAS
{
[CmdLetBinding()]
Param([Parameter(Mandatory = $false, Position = 0)][string]$DomainController
     )

    $DCParam=@{
        ErrorAction='Stop';
    }; # end DCParam
    if ($PSBoundParameters.ContainsKey('DomainController'))
    {
        $DCParam.Add('DomainController',$DomainController);
    }; # end if
    try
    {
        return ,$__OPX_ModuleData.getExchangeServerList($false,$true,$true).ToLower();;
    } # end try
    catch
    {
        writeToLog -LogType Error -LogString  ('Failed to connect to client access service on server ' + $Server);
        writeToLog -LogString ($_.Exception.Message) -LogType Error;
        return $false;
    }; # end catch
}; # end if getAllCAS

function getDAGMemberServer
{
[CmdLetBinding()]
Param([Parameter(Mandatory = $true, Position = 0)][string]$DAGName,
      [Parameter(Mandatory = $false, Position = 1)][string]$DomainController
     )
    
    $DCParam=@{
        ErrorAction='Stop';
    }; # end DCParam
    if ($PSBoundParameters.ContainsKey('DomainController'))
    {
        $DCParam.Add('DomainController',$DomainController);
    }; # end if
    try
    {
        $msg=('Quereing DAG ' + $DAGName + ' for member server');
        writeTolog -LogString $msg;
        Write-Verbose $msg;
        $dag=$__OPX_ModuleData.getDagList($true,$DAGName);
        if (($dag.Count -eq 0) -or ([System.String]::IsNullOrEmpty($dag[0])))
        {
            writeTolog -LogString ('The DAG ' + $dagName + ' has no meber server.') -LogType Warning;
            return $false;
        }; # end if
        return ,$dag;
    } # end try
    catch
    {
        writeToLog -LogType Error -LogString  ('Failed to locate the DAG ' + $DAGName);
        writeToLog -LogString ($_.Exception.Message) -LogType Error;
        return $false;
    }; # end catch
}; # end if getDAGMemberServer

function getUinqueFromArrray
{
[CmdLetBinding()]
Param([Parameter(Mandatory = $true, Position = 0)][array]$List
     )

    return @([System.Linq.Enumerable]::Distinct([string[]]$List));
    
}; # end function getUinqueFromArrray

function getVDirConfig
{
[CmdLetBinding()]
Param([Parameter(Mandatory = $true, Position = 0)][string]$ConfigFilePath
     )  
    try {
        writeTolog -LogString ('Reading configuration file ' + $ConfigFilePath);
        return ,(Import-Clixml -Path $ConfigFilePath -ErrorAction Stop);        
    } # end try
    catch {
        writeTolog -LogString ('Faild to read the configuration file ' + $ConfigFilePath) -LogType Error;
        writeTolog -LogString ($_.Exception.Message) -LogType Error;
        return $false;
    }; # end catch
}; # end function getVDirConfig

function resetVDirConfg
{
[CmdLetBinding()]
Param([Parameter(Mandatory = $true, Position = 0)][string]$ConfigFilePath
     )

    $servicesList=@('OWA','ECP','OAB','Webservices','ActiveSync','OutlookAnywhere','Mapi','Autodiscover');
    $cfgTemplate=@{};
    $srvCount=$servicesList.count;
    for ($i=0;$i -lt $srvCount;$i++)
    {
        $cfgTemplate.Add($servicesList[$i],@{});
    }; # end for
    try {
        writeToLog -LogString ('Resetting VDir template.')
        $cfgTemplate | Export-Clixml -Path $ConfigFilePath;
    } # end try
    catch {
        writeTolog -LogString ('Faild to save the resetted configuration file to ' + $ConfigFilePath) -LogType Error;
        writeTolog -LogString ($_.Exception.Message) -LogType Error;
    }; # end catch
}; # end function resetVdirConfig

function saveVDirConfig
{
[CmdLetBinding()]
Param([Parameter(Mandatory = $true, Position = 0)][string]$ConfigFilePath,
      [Parameter(Mandatory = $true, Position = 1)][hashtable]$CfgHash
     )  

    try {
        writeTolog -LogString ('Saving  configuration to file ' + $ConfigFilePath);
        $CfgHash | Export-Clixml -Path $ConfigFilePath -ErrorAction Stop;        
    } # end try
    catch {
        writeTolog -LogString ('Faild to save the configuration file to ' + $ConfigFilePath) -LogType Error;
        writeTolog -LogString ($_.Exception.Message) -LogType Error;
    }; # end catch
}; # end function saveVDirConfig

function getQueryString
{
[CmdLetBinding()]
Param([Parameter(Mandatory = $true, Position = 0)][array]$queryFilter
     )  
    
    $qFEntries=$queryFilter.TrimStart('(').TrimEnd(')').Split(' ');
    $qFC=$qFEntries.count;
    $qStr='';
    for ($i=0;$i -lt $qFC;$i++)
    {
        switch ($qFEntries[$i])
        {
            {$_.StartsWith('$_.')} {$qStr+=(($_.Remove(0,3)) + ' '); continue};
            {$_.StartsWith('$')} {$qStr+=(Get-Variable -Name ($_.Remove(0,1)) -ValueOnly) + ' '; continue}
            Default {$qStr+=($_) + ' '};
        }; # end switch
    }; # end for
    return $qStr.TrimEnd();
}; # end function getQueryString

function saveServiceInfo
{
[CmdLetBinding()]
Param([Parameter(ParametersetName='File',Mandatory = $true, Position = 0)][string]$ServiceConfig,
      [Parameter(ParametersetName='cfgObject', Mandatory = $true, Position = 0)][array]$ServiceCfgObject,
      [Parameter(Mandatory = $true, Position = 1)][string]$ComputerName,
      [Parameter(Mandatory = $true, Position = 2)][string]$ServiceFileDirectoryPath,
      [Parameter(Mandatory = $true, Position = 3)][string]$ServiceFilePrefix
     ) 

    $srvList=[System.Collections.ArrayList]::new();
    if ($PsCmdlet.ParametersetName -eq 'File')
    {                                                            
        $srvFilter=(Import-Csv -Path ([System.IO.Path]::Combine($__OPX_ModuleData.cfgDir,'Services\'+$ServiceConfig+'.csv')) -ErrorAction Stop);
    } # end if
    else {
        $srvFilter=$ServiceCfgObject;
    }; # end else
    $paramName=($srvFilter | Get-Member -MemberType NoteProperty)[0].Name;
    $pList=@{
        $paramName='';
        #Query='';
        ComputerName=$ComputerName;
        ErrorAction='Stop';
    }; # end pList              
    foreach ($srv in $srvFilter)
    {
        $pList.$paramName=$srv.$paramName
        writeTolog -LogString ('Searching service(s) with ' + $paramName + ' ' + $srv.$paramName);        
        [void]$srvList.AddRange(@(Get-Service @pList)); 
    }; # end forach
    $outFile=([System.IO.Path]::Combine($ServiceFileDirectoryPath,($ServiceFilePrefix + '_' + $ComputerName))+ '.xml');
    if ($srvList.count -gt 0)
    {
        $msg=('Exporting service start-up configuration to file ' + $outFile)
        Write-Verbose $msg;
        writeTolog -LogString $msg;
        try {
            $srvList | Export-Clixml -Path $outFile  -ErrorAction Stop;
        } # end try
        catch {
            writeTolog -LogString ('Failed to export the service start-up configuration to file ' + $outFile) -LogType Error;
            writeToLog -LogString ($_.Exception.Message) -LogType Error;
        }; # end catch
        
    } # end if
    else {
        writeToLog -LogType Warning -LogString  ('On computer ' + $ComputerName + ' no services, whilch match your selection, found.');
    }; # end else
}; #end function saveServiceInfo

function getSMTPServer
{
    $serverList=[System.Collections.ArrayList]::new();
    [void]$serverList.AddRange($__OPX_ModuleData.getExchangeServerList($true,$true,$true));
    $mailServerNotFound=$true;
    $mailServer=$false;
    while ($mailServerNotFound)
    {
        $smtpServer=($serverList | Get-Random);
        $mailServerNotFound=(! ((testPort -ComputerName $smtpServer -Port 25 -TcpTimeout 50) -and ((Get-ServerComponentState -Component hubtransport -Identity $smtpServer).state -eq 'Active')));
        if (! ($mailServerNotFound))
        {
            $mailServer=$smtpServer;
        } # end if
        else
        {
            [void]$serverList.Remove($smtpServer);
        }; # end else
    }; # end while
    return $mailServer
}; # end funciton getSMTPServer

class getModuleData
{        
    getModuleData () {
        $this.PSObject.Properties.Add(
            (New-Object PSScriptProperty 'cmdletList', {$this._cmdletList})
        ) # end cmdletList
        $this.PSObject.Properties.Add(
            (New-Object PSScriptProperty 'cfgDir', {$this._cfgDir})
        ) # end cfgDir
        $this.PSObject.Properties.Add(
            (New-Object PSScriptProperty 'logDir', {$this._logDir})
        ) # end logDir
        
        $this.PSObject.Properties.Add(
            (New-Object PSScriptProperty 'VirtualDirCfgFileName', {$this._VirtualDirCfgFileName})
        ) # end logDir
        $this.PSObject.Properties.Add(
            (New-Object PSScriptProperty 'ConnectedToMSXServer', {$this._ConnectedToMSXServer})
        ) # end connectedToMSXServer
        #>
    } # end getModuleData

    [array]getDagList ([switch]$returnMember,
                        [string]$DAGName)
    {
        $pL=@{
            LDAPFilter='(objectClass=msExchMDBAvailabilityGroup)' ;
            SearchBase=$Script:msxAdmGroupSearchRootStr;
            FieldList=@('cn','name','msexchstartedmailboxservers');
        }; # end Pl

        if (!([System.String]::IsNullOrEmpty($DAGName)))
        {
            $pl.LDAPFilter='(&(objectClass=msExchMDBAvailabilityGroup)(cn='+$dagName+'))';
        }; # end if
        $rv = (queryADForObjects @pL);
        $dagList=@()
        foreach($dag in $rv)
        {
            try {
                if ($returnMember.IsPresent)
                {
                    if ([System.Linq.Enumerable]::Contains([string[]]$dag.Properties.PropertyNames,'msexchstartedmailboxservers'))
                    {
                        $dagList+=$dag.Properties.msexchstartedmailboxservers;
                    };
                } # end if
                else {
                    $dagList+=$dag.properties.name;
                }; # end else                
            } # end try
            catch {
                writeTolog -LogString ('Faild to add data for DAG ' + $DAGName) -LogType Error;
                writeTolog -LogString ($_.Exception.Message) -LogType Error;
            }; # end catch            
        }; # end foreach 
        return [array]$dagList       
    } # end method getDagList

    [array]getExchangeServerList ([switch]$IsMailboxServer,
                                  [switch]$ReturnFQDN=$false,
                                  [switch]$SupressFullInfo)  
    {
        $pL=@{        
            LDAPFilter='(&((objectClass=msExchExchangeTransportServer)(objectClass=msExchExchangeServer)(name=Frontend)))';
            SearchBase=$Script:msxServerSearchRootStr;
            FieldList=@('cn','networkaddress');
        }; # end Pl
        if ($IsMailboxServer.ispresent)
        {
            $pl.LDAPFilter='(&((objectClass=msExchExchangeTransportServer)(objectClass=msExchExchangeServer)(name=Mailbox)))' 
        }; # end if

        if ($ReturnFQDN.IsPresent)
        {
            $searchStr='ncacn_ip_tcp:';
        }
        else
        {
            $searchStr='netbios:';
        }; # end else
        $serverTbl='';
        if ($SupressFullInfo.IsPresent -eq $false)
        {
            $pl.FieldList=@('cn','networkaddress','msExchMDBAvailabilityGroupLink','msExchServerSite','serialNumber','msExchComponentStates','description');
            $pl.LDAPFilter='(&(objectClass=msExchExchangeServer)(!(objectClass=msExchExchangeTransportServer)))';
            $FieldList = @(
            @('Server Name',[System.String]),
            @('Server Version',[System.String]),
            @('Member in DAG',[System.String]),
            @('AD Site',[System.String]),
            @('Component',[System.String]),
            @('Maint. State',[System.Boolean])
            ); # end FieldList
            $serverTbl=createNewTable -TableName 'MSXServers' -FieldList $FieldList;
        };
        $rv = (queryADForObjects @pL);
        $rvList=@();
        $srvCount=$rv.Count;
            
        for ($i=0;$i -lt $srvCount;$i++)
        {                
            if ($SupressFullInfo) # if true only return the names of server
            {
                try {
                    foreach ($prop in $rv[$i].Properties.networkaddress)
                    {                
                        if ($prop.startsWith($searchStr))
                        {
                            $rvlist+=$prop.split(':')[1];
                            break;
                        }; # end if
                    }; # end foreach
                } # end try
                catch {
                    writeTolog -LogString ('Entry for Exchange server ' + $rv[$i].Properties.cn + ' is not valid.') -LogType Error;
                }; # end catch                
            } # end if
            else { # return more info
                $maintInfo=@()
                try {
                    foreach ($prop in $rv[$i].Properties.msexchcomponentstates)
                    {
                        if ($prop -like '*Maintenance*')
                        {
                        $maintInfo+=$prop
                        }; # end if
                    }; # end foreach
                    $mIc=$maintInfo.count;
                    if ($rv[$i].Properties.Contains('msExchMDBAvailabilityGroupLink'))
                    {
                        $dagName=(($rv[$i].Properties.msexchmdbavailabilitygrouplink).split(',')[0]).Remove(0,3);
                    }
                    else
                    {
                        $dagName='N/A';
                    }; # end else
                    $name=''
                    
                    foreach ($prop in ($rv[$i].Properties.networkaddress))
                    {
                        if ($prop.startsWith($searchStr))
                        {
                            $name = $prop.split(':')[1];
                            break;
                        }; #end if
                    }; # end foreach
                    $fieldsToAdd=@(
                        $name, 
                        [string]$rv[$i].Properties.serialnumber,
                        $dagName,
                        (($rv[$i].Properties.msexchserversite).split(','))[0].Remove(0,3)
                        ($maintInfo[0].Split(':')[1] + ' '),
                        ($maintInfo[0].Split(':')[3] -ne 1)
                    ); # end fieldsToAdd
                    [void]$serverTbl.Rows.Add($fieldsToAdd);
                    for ($j=1;$j -lt $mIc;$j++)
                    {
                        [void]$serverTbl.Rows.Add('','',$dagName,'',($maintInfo[$j].Split(':')[1]), ($maintInfo[$j].Split(':')[3] -ne 1))
                    }; # end for
                } # end try
                catch {
                    writeTolog -LogString ('Entry for Exchange server ' + $rv[$i].Properties.cn + ' is not valid.') -LogType Error;
                }; # end catch                                
            }; # end else        
        }; # end for 
        if ($SupressFullInfo.IsPresent -eq $false)
        {
            ($serverTbl.Select('', '[Member in DAG] ASC'));
            $rc=$serverTbl.rows.count;
            for ($i=0;$i -lt $rc;$i++)
            {
                if ([system.string]::IsNullOrEmpty(($serverTbl.rows[$i].'Server Version')))
                {
                    $serverTbl.rows[$i].'Member in DAG'='';
                }; # end if
            }; # end for
            $rvList+=$serverTbl;
        }; # end if
        return $rvList
    } # end methode getExchangeServerList
    # list of constants needed in Argument Completer
    hidden [array]$_cmdletList=$Script:cmdletList;
    hidden [string]$_cfgDir=$Script:cfgDir;
    hidden [string]$_logDir=$Script:logDir;
    hidden [string]$_VirtualDirCfgFileName=$Script:VirtualDirCfgFileName;
    hidden [string]$_ConnectedToMSXServer=$Script:TmpConMsxSrv;
    
}; # end class getModuleData

function initEnumerator
{

try {
Add-Type -TypeDefinition @'
public enum CertDeploymentType {
    CopyOnly,
    CopyAndEnable,
    EnableOnly
}
'@    

Add-Type -TypeDefinition @'
public enum OutFormat {
    Table,
    List,
    PassValue
}
'@

}
catch {
    # type exists
}; # end catch

}; # end
#endregion helper function
