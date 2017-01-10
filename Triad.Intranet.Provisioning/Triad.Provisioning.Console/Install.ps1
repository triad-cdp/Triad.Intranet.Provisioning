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

Write-Host "Tenant: $Tenant"
Write-Host "Site: $Site"

Write-Host "Started installation"

$path = Split-Path -parent $MyInvocation.MyCommand.Definition


[xml]$config = Get-Content -Path $path config.xml



if ($env:PSModulePath -notlike "*$path\Modules\*")
{
	"Adding ;$path\Modules to PSModulePath" | Write-Debug 
	$env:PSModulePath += ";$path\Modules\"
}

Write-Host $env:PSModulePath

$url = $Tenant + $Site

whoami


$password = "W3ybr00k"
$username = "matt@zephyrgroup.onmicrosoft.com"

$encpassword = convertto-securestring -String $password -AsPlainText -Force

$cred = new-object -typename System.Management.Automation.PSCredential `
         -argumentlist $username, $encpassword


Connect-PnPOnline -Url $url -Credentials $cred 

Write-Host "Connected to PnP Online"

Write-Host "Applying template to $url"

Set-PnPTraceLog -On -Level Debug

$web = Get-PnPWeb

Apply-PnPProvisioningTemplate -Web $web -Path "$path\Templates\Home\Home.xml" -ResourceFolder "$path\Templates\Home"

Write-Host "Applied template"

Write-Host "Completed installation"


