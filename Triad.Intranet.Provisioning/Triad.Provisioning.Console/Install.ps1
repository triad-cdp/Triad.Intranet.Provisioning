#
# Script1.ps1
#

Write-Host "Started installation"

Dir

$path = Split-Path -parent $MyInvocation.MyCommand.Definition

if ($env:PSModulePath -notlike "*$path\Modules\*")
{
	"Adding ;$path\Modules to PSModulePath" | Write-Debug 
	$env:PSModulePath += ";$path\Modules\"
}

Write-Host $env:PSModulePath

Set-PnPTraceLog -On -Level Debug

Write-Host "Completed installation"
