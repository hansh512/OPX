


function Get-ActiveConfig
{
    $script:ActiveCfgNuV;
    foreach ($entry in $script:ActiveConfig.keys)
    {
        if (($script:ActiveConfig.$entry).GetType().name -eq 'PSCredential')
        {
            Write-Host ($entry + ':  ' + ($script:ActiveConfig.$entry).UserName) -ForegroundColor Green;
        } # end if
        else
        {
            Write-Host ($entry + ':  ' + ($script:ActiveConfig.$entry).ToString());
        }; # end else
    }; # end foreach
  

}; # end function Get-ActiveConfig

function Get-Config
{
Param([Parameter(Mandatory = $true, Position = 0)]
        [ArgumentCompleter( {  
            param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )           
            $cfgList=($__CMM_ModuleData.GetConfigerationList($__OLX_2_0_0_0.ModuleName,$__OLX_2_0_0_0.ModuleVersion,$Global:SkipTemplateVersionFiltering));
            $cfgList.Where({ $_ -like "$wordToComplete*" });              
        } )]    
      [string]$Configuration,
      [Parameter(Mandatory = $false, Position = 1)][switch]$SkipTemplateFiltering
    )
        
    $data=$__CMM_ModuleData.GetConfig($Script:ModuleName,$Script:ModuleVersion,$Configuration,$SkipTemplateFiltering)
    if ($null -ne $data)
    {
        $data;
        Write-Host 'Variables:'
        foreach ($entry in $data.Data.keys)
        {
            if ($data.Data.$entry.getType().name -eq 'PSCredential')
            {
                Write-Host ($entry + ':  ' + $data.data.$entry.UserName);
            } # end if
            else {
                Write-Host ($entry + ':  ' + $data.data.$entry.toString());
            }; # end if
        }; # end if
        if ($data.InconsistentAttributes.count -gt 0)
        {
            Write-Warning 'InconsistentAttributes:'
            $data.InconsistentAttributes
        }

    } # end if
    else
    {
        Write-Warning 'No configuration found';
    }; # end else
}; # end function Get-Config


function Set-ConfigAsActive
{
Param([Parameter(Mandatory = $true, Position = 0)]
        [ArgumentCompleter( {  
            param ( $CommandName,
            $ParameterName,
            $WordToComplete,
            $CommandAst,
            $FakeBoundParameters )           
            $cfgList=($__CMM_ModuleData.GetConfigerationList($__OLX_2_0_0_0.ModuleName,$__OLX_2_0_0_0.ModuleVersion,$Global:SkipTemplateVersionFiltering));
            $cfgList.Where({ $_ -like "$wordToComplete*" });              
        } )]    
      [string]$Configuration,
      [Parameter(Mandatory = $false, Position = 1)][switch]$StrictTemplateFiltering=$false
    )

    setActiveConfig -ConfigName $Configuration -Version ($__OLX_2_0_0_0.ModuleVersion);
}; # end function Set-ConfigAsActive

function Test-RegXString
{
[cmdletbinding()]
param([Parameter(Mandatory = $false, Position = 0)][ValidateScript({$_ -match $RexHelpMsg})][string]$TestVal,
      [Parameter(Mandatory = $false, Position = 0)][ValidateScript({$_ -match $RexCfgName})][string]$CfgName
     )


    Write-Host 'OK'
    $cfgName
}
