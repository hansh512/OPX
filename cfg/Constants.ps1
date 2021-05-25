################################################################################
# Code Written by Hans Halbmayr
# Created On: 21.03.2021
# Last change on: 25.05.2021
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
New-Variable -Name 'VirtualDirCfgFileName' -Value 'VdirCfg.xml' -Option Constant -Scope Script;
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
New-Variable -Name 'CustomLogFileDir' -Value '' -Option Constant -Scope Script;

