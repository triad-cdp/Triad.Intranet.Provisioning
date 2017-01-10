#
# Install.ps1
#

Param(
  [string]$Tenant,
  [string]$Site,
  [string]$Username,
  [string]$Password
)

Write-Host "Tenant: $Tenant"
Write-Host "Site: $Site"
Write-Host "User: $Username"
Write-Host "Password: $Password"

Write-Host "Started installation"

$path = Split-Path -parent $MyInvocation.MyCommand.Definition


[xml]$config = Get-Content -Path $path/config.xml



if ($env:PSModulePath -notlike "*$path\Modules\*")
{
	"Adding ;$path\Modules to PSModulePath" | Write-Debug 
	$env:PSModulePath += ";$path\Modules\"
}

Write-Host $env:PSModulePath

$url = $Tenant + $Site

whoami


$encpassword = convertto-securestring -String $Password -AsPlainText -Force

$cred = new-object -typename System.Management.Automation.PSCredential `
         -argumentlist $Username, $encpassword


Connect-PnPOnline -Url $url -Credentials $cred 

Write-Host "Connected to PnP Online"

Write-Host "Applying template to $url"

Set-PnPTraceLog -On -Level Debug

$web = Get-PnPWeb

Apply-PnPProvisioningTemplate -Web $web -Path "$path\Templates\Home\Home.xml" -ResourceFolder "$path\Templates\Home"

Write-Host "Applied template"

Write-Host "Completed installation"


