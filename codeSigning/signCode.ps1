###############################################################################
# Code Written by Hans Halbmayr
# Created On: 20.04.2021
# Last change on: 24.04.2021
#
# Module: N/A 
#
# Version 1.00
#
# Purpose: Sign PowerShell module files
################################################################################
#requires -version 5

# Tested with PowerShell 5.1 on Windows Server 2016

function Set-AuthenticodeSignatureForPSModuleFiles
{
<#
.SYNOPSIS 
Signs PowerShell script files.
    
.DESCRIPTION
The command signs script files for a PowerShell module of a single file. If files for a module are singned, the command lists all *.ps* files under the root directory (and all subdirectories) of a given PowerShell module and signes the *.ps* files.

.PARAMETER ModuleName
The parameter is mandatory.
The name of the PowerShell module. If the module cannot be loaded automatically, enter the path to the module.
The parameter ModuleName and the parameter FilePath are mutual exclusive.

.PARAMETER FilePath
The parameter is mandatory.
The path to the PowerShell file to sign.
The parameter FilePath and the parameter ModuleName are mutual exclusive.

.PARAMETER SigningCertificate
The parameter is optional.
The parameter expects the code signing certificate. If the parameter is omitted, the command looks up in the certificate store of the user for a code signing certificate. The command will pick up the first code signing certificate.

.PARAMETER KeepLastFileWriteTime
The parameter is optional.
If the parameter is used, the code signing will not change the file property LastWriteTime. 

.EXAMPLE
Set-AuthenticodeSignatureForPSModuleFiles -ModuleName <name of module>
The files, found under the PS module root directoy, will be signed with the code signing certificate found in the certificate store of the user.
.EXAMPLE
Set-AuthenticodeSignatureForPSModuleFiles -ModuleName <name of module> -SigningCertificate $CodeSigningCert
The files, found under the PS module root directoy, will be signed with the code signing certificate provided by the parameter SigningCertificate.
#>
[cmdletbinding(DefaultParameterSetName='Module')]
param([Parameter(ParametersetName='Module',Mandatory = $true, Position = 0)][string]$ModuleName,
      [Parameter(ParametersetName='File',Mandatory = $false, Position = 0)][string]$FilePath,
      [Parameter(Mandatory = $false, Position = 1)][System.Security.Cryptography.X509Certificates.X509Certificate]$SigningCertificate,
      [Parameter(Mandatory = $false, Position = 2)][switch]$KeepLastFileWriteTime=$false
     )

    if (! ($PSBoundParameters.ContainsKey('SigningCertificate')))
    {
        try {
            $SigningCertificate=(Get-ChildItem cert:\CurrentUser\my -CodeSigningCert -ErrorAction Stop)[0]; # load code signing certificate from user certificate store
        } # end try
        catch {
            Write-Warning 'Faild to load the code signing certificate from user certificate store.';
            Write-Host ($_.Exception.Message);
            return;
        }; # end catch        
    }; # end if

    if ($PsCmdlet.ParameterSetName -eq 'Module')
    {
        try {
            $moduleRoot=[System.IO.Path]::GetDirectoryName((Get-Module -Name ($ModuleName.TrimEnd('\')) -ListAvailable).path); # get the path to the module files
        } # end try
        catch {
            Write-Warning ('Module ' + $ModuleName + ' not found.');
            Write-Host ($_.Exception.Message);
            return;
        }; # end catch
        try {
            Write-Verbose ('Building file list ('+$moduleRoot+')');
            $fileList=@(Get-ChildItem -Path $moduleRoot -Recurse -Filter '*.ps*' -ErrorAction Stop); # get list of *.ps* files
        } # end try
        catch {
            Write-Warning ('Failed read the files for module ' + $ModuleName);
            Write-Host ($_.Exception.Message);
        }; # end catch        
    } # end if
    else {
        if (Test-Path -Path $FilePath)
        {
            $fileList=@{
                FullName=$FilePath;
            }; # end fileList
        } # end if
        else {
            Write-Warning ('File ' + $FilePath + ' not found');
            return;
        }; # end else
    }; # end else

    foreach ($fileName in $fileList.FullName)
    {
        try {
            Write-Verbose ('Verifying authenticodeSignature for file ' + $fileName);
            if ((Get-AuthenticodeSignature -FilePath $fileName).status -ne 'Valid')
            {
                Write-Verbose ('Signing file ' + $fileName);
                $fileInfo=Get-Item -Path $fileName -ErrorAction Stop ;
                $lwt=$fileInfo.LastWriteTime; # save timestamp for LastWriteTime
                [void](Set-AuthenticodeSignature -Certificate $SigningCertificate -FilePath $fileName -ErrorAction Stop); # sign file
                if ($KeepLastFileWriteTime.IsPresent)
                {
                    Write-Verbose ('Setting LastWriteTime to ' + $lwt.toString())
                    $fileInfo.LastWriteTime=$lwt; # reset LastWriteTime
                }; # end if
            } # end if
            else {
                Write-Host ('Skipping file ' + $fileName + ', authenticodeSignature valid.');
            }; # end if
        } # end try
        catch {
            Write-Warning ('Failed to sign file ' + $fileName);
            Write-Host ($_.Exception.Message);
        }; # end catch    
    }; # end foreach

}; # end function Set-AuthenticodeSignatureForPSModuleFiles
