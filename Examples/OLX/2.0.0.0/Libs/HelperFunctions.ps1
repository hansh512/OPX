

function initCredentialData
{

    if ($tmp=$__CMM_ModuleData.GetDefaultConfig($Script:moduleName,$Script:ModuleVersion)) 
    {
        $script:DefCfg=$tmp.split('_')[4]         
        setActiveConfig -ConfigName $script:DefCfg;
        Write-Host ('Configuration ' + ($script:DefCfg) + ' as default config loaded'); 
    
    } # end if
    else
    {
        Write-Warning 'No default configuration found.'
        $script:ActiveConfig=@{};
    }; # end else

}; # end function initCredentialData


function setActiveConfig
{
[CmdLetBinding()]            
param([Parameter(Mandatory = $true, Position = 0)][string]$ConfigName,
      [Parameter(Mandatory = $false, Position = 1)][Version]$Version=$Script:ModuleVersion,
      [Parameter(Mandatory = $false, Position = 2)][Switch]$SkipTemplateFiltering=$Global:SkipTemplateVersionFiltering
     )  

    $tmpCfg=$__CMM_ModuleData.GetConfig($Script:ModuleName,$Version,$ConfigName,$SkipTemplateFiltering);
    $script:ActiveConfig=$tmpCfg.Data;
    $script:ActiveCfgNuV=@{
        Name=$tmpCfg.ConfigName;
        Version=$tmpCfg.ConfigVersion;
    }; # end ActiveCfgNuV
}; # end function setActiveConfig


class getModuleData
{        
    getModuleData () {
        $this.PSObject.Properties.Add(
            (New-Object PSScriptProperty 'ModuleName', {$this._ModuleName})
        ) # end vaultName
        
        $this.PSObject.Properties.Add(
            (New-Object PSScriptProperty 'ModuleVersion', {$this._ModuleVersion})
        ) # end cfgNamePrefix
                
    } # end getModuleData

    UpdateConfig ([string]$ConfigName,
                      [version]$Version
                     )
    {                
        if (($script:ActiveCfgNuV.Version -eq $Version) -and ($script:ActiveCfgNuV.Name -eq $ConfigName))
        {
            Write-Host 'Updating...' -ForegroundColor Green;
            setActiveConfig -ConfigName $ConfigName -Version $Version
        } # end if
        else {
            Write-Host 'No update needed' -ForegroundColor DarkMagenta;
        }; # end else
           #return $true
    } # end method getConfig

    
    # list of constants needed in Argument Completer and default values
    hidden [string]$_ModuleName=$Script:ModuleName;
    hidden [version]$_ModuleVersion=$Script:ModuleVersion
    
}; # end class getModuleData