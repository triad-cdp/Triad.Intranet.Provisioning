#
# Install.ps1
#

Param(
  [string]$Tenant,
  [string]$Site
)

if ($Tenant -eq $null -or $Tenant -eq "")
{
	$Tenant = "https://zephyrgroup.sharepoint.com"
	$Site =  "/sites/pieter"
}

Write-Host "Started installation"

Write-Host "Tenant: $Tenant"

Write-Host "Site: $Site"

Dir

$path = Split-Path -parent $MyInvocation.MyCommand.Definition

if ($env:PSModulePath -notlike "*$path\Modules\*")
{
	"Adding ;$path\Modules to PSModulePath" | Write-Debug 
	$env:PSModulePath += ";$path\Modules\"
}

Write-Host $env:PSModulePath

Set-PnPTraceLog -On -Level Debug

$url = $Tenant + $Site

$password = "W3ybr00k"
$username = "matt@zephyrgroup.onmicrosoft.com"

$password = convertto-securestring -String $password -AsPlainText -Force

$cred = new-object -typename System.Management.Automation.PSCredential `
         -argumentlist $username, $password


Connect-PnPOnline -Url $url -Credentials $cred 

Write-Host "Connected to PnP Online"

Write-Host "Applying template to $url"

Apply-PnPProvisioningTemplate -Path "$path\Templates\Home.pnp" -ExcludeHandlers Files

Write-Host "Applied"

Write-Host "Completed installation"
