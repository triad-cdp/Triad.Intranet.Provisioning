#
# Install.ps1
#

Param(
  [string]$Tenant,
  [string]$Site,
  [string]$Username,
  [string]$Password
)

begin
{
	Clear-Host
	Write-Host "Tenant: $Tenant"
	Write-Host "Site: $Site"
	Write-Host "User: $Username"
	Write-Host "Password: $Password"

	Write-Host "Started installation"

}

process
{

	$path = Split-Path -parent $MyInvocation.MyCommand.Definition


	$config = [xml](Get-Content $path/config.xml -ErrorAction Stop)



	if ($env:PSModulePath -notlike "*$path\Modules\*")
	{
		"Adding ;$path\Modules to PSModulePath" | Write-Debug 
		$env:PSModulePath += ";$path\Modules\"
	}

	Write-Host $env:PSModulePath

	#set the url of the site collection
	$url = $Tenant + $Site

	$encpassword = convertto-securestring -String $Password -AsPlainText -Force

	$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $Username, $encpassword


	Connect-PnPOnline -Url $url -Credentials $cred 
	Write-Host "Connected to PnP Online"


	$sitesConfig = $config.Configurations.Configuration.Sites.Site

	foreach ($siteConfig in $sitesConfig)
	{
		$siteUrl = $null


		$siteTitle = $siteConfig.Name
		$siteBaseTemplate = $siteConfig.BaseTemplate
		$siteUrl = $siteConfig.Url

		if ($siteUrl -eq "")
		{ 
			$web = Get-PnPWeb 
		}
		else
		{
			$web = Get-PnPWeb $siteUrl -ErrorAction SilentlyContinue
	    }

		if ($web  -eq $null)
		{
			$newWebUrl = $siteUrl.Split("/")[$siteUrl.Split("/").Count-1]

			$parentWebUrl = $siteUrl.Replace("/"+$newWebUrl,"")
			
			New-PnPWeb -Web $parentWebUrl -Url "$newWebUrl" -Title "$siteTitle" -Template "$siteBaseTemplate"
			$web = Get-PnPWeb $siteUrl

		}


		if ($web  -eq $null)
		{
			Write-Error "Failed to find site at $siteUrl"
			exit
		}
		
		$template = $siteConfig.Template		

		Write-Host "Applying template $template to $siteUrl"
		
		Set-PnPTraceLog -On -Level Debug
				

		Apply-PnPProvisioningTemplate -Web $web -Path "$path\Templates\$template\Home.xml" -ResourceFolder "$path\Templates\$template"

	}

	
	

	
	Write-Host "Applied template"

}

end
{
    Write-Host "Completed installation!"
}


