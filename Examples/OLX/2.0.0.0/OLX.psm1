
$filesToLoad=@(    
    ($PSScriptroot + '\\Libs\HelperFunctions.ps1'),
    ($PSScriptroot + '\\Libs\cmdlets.ps1')
); # end filesToLoad

foreach ($file in $filesToLoad)
{
    try
    {
    . $file;
    } # end try
    catch
    {
        Write-Warning ('Failed to load the file ' + $file);
        Write-Host ($_.Exception.Message);
        #return;
    }; # end catch
}; # end foreach

$mfp=($MyInvocation.MyCommand.path)
#New-Variable -Name 'moduleName' -Value (([System.IO.Path]::GetFileNameWithoutExtension(($mfp))).ToUpper()) -Scope Global; #-Option Constant;
$Script:ModuleName=(([System.IO.Path]::GetFileNameWithoutExtension(($mfp))).ToUpper())
$mList=(Get-Module -Name $Script:ModuleName -ListAvailable);
$refPath=[System.IO.Path]::ChangeExtension($mfp,'psd1');
foreach ($entry in $mList)
{
    if ($entry.path -eq $refPath)
    {
        $Script:ModuleVersion=($entry.version)
        break;
    }; # end if
}; # end foreach

$Global:SkipTemplateVersionFiltering=$false

Export-ModuleMember @(
                        'Get-Config';
                        'Get-ActiveConfig',
                        'Set-ConfigAsActive',
                        'Test-RegXString'
                     )
# load default configuration
initCredentialData

$Script:HelpMsgRS='^[a-zA-Z0-9 ,.=\d-<>();@]+$'
$script:CfgName='^[a-zA-Z0-9'']+$'
$RexHelpMsg=[regex]::new($Script:HelpMsgRS)
$RexCfgName=[regex]::new($Script:CfgName)

$vStr=(($Script:ModuleVersion).ToString()).Replace('.','_');
$vName='__' + $Script:ModuleName + '_' + $vStr; 
New-Variable -Name $vName -Value ([GetModuleData]::new());
Export-ModuleMember -Variable $vName;



