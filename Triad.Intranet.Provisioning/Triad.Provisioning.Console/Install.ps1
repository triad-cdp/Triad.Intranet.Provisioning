#
# Install.ps1
#

Param(
  [string]$Tenant,
  [string]$Site
)



Write-Host "Started installation"

Write-Host "Tenant: $Tenant"

Write-Host "Site: $site"

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
