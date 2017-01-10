#
# TriadOffice365.psm1
#
# Functions for handling Office 365 operations
#
##############################################
$global:credentials


function Connect-TriadOffice365
{
[CmdletBinding()]
	param(

		[Parameter(Mandatory=$True,ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True,HelpMessage=’The site to apply the template to’)]		
		[ValidateLength(3,254)]
		[string]$url
	)

    try
    {
        Write-Host $connecterror
    }
    catch [exception]
    {    

	 If ($Error.Count -gt 0)
        {
            switch ($Error[0].Exception)
            {
               
                "Connect-SPOnline : Identity Client Runtime Library (IDCRL) could not look up the realm information for a federated sign-in."
                {
                    $global:credentials = Get-Credential
                    Connect-PnPOnline -Url $url -Credentials($global:credentials)
                }

                "The remote server returned an error: (403) Forbidden."        
                {
                    Write-Host $url " is a subweb"
                    $siteUrl = $url.Substring($url.Length - $url.IndexOf("sharepoint.com/")+15)

                    exit
                    New-SPOWeb -Url $url -Title "Title" -Description "Description" -InheritNavigation $true -Locale 1033 -Template "STS#0"
                    Connect-SPOnline -Url $url -Credentials($global:credentials)        
                }
        
                default
                {
                    Write-Error "Unhandled exception"
                    exit
                }
            }
        }
     }
}

function Add-Site {

	<# 
	 .SYNOPSIS
	 Add a site
	 .DESCRIPTION
	 Add a site
	 .EXAMPLE
	 Add-Site "https://mytenant.sharepoint.com" 
	 .PARAMETER param1
	 My parameter
	#>

	[CmdletBinding()]
	param(
        
		[Parameter(Mandatory=$True,ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True,HelpMessage='The site to apply the template to')]		
		[ValidateLength(0,254)]
		[string]$parentweburl,

		[Parameter(Mandatory=$True,ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True,HelpMessage='The site to apply the template to')]		
		[ValidateLength(0,254)]
		[string]$weburl,

		[Parameter(Mandatory=$True,ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True,HelpMessage=’The template to apply')]		
		[ValidateLength(0,254)]
		[string]$sitetitle,

		[Parameter(Mandatory=$True,ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True,HelpMessage=’The template to apply')]		
		[ValidateLength(0,254)]
		[string]$siteassets,

		[Parameter(Mandatory=$True,ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True,HelpMessage=’The template to apply')]		
		[ValidateLength(0,254)]
		[string]$applytemplate,

		[Parameter(Mandatory=$True,ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True,HelpMessage='The xml template to apply')]		
		[ValidateLength(0,254)]
		[string]$sitetemplate

	)

    begin
    {
  
        write-debug "Begin Add-Site"
        write-Host "Begin Add-Site"

		if ($global:credentials -eq $null)
		{
			$global:credentials = Get-Credential
			if ($global:credentials -eq $null)
			{
				Write-Warning "Cannot create site without credentials"
				exit
			}
		}

        $XMLFilePath = "$path\configuration\config.xml"
        $siteurl = Get-Setting -settingName "siteurl" -XMLFilePath $XMLFilePath
    }

    process
    {
        write-debug "Process Add-Site"

        Connect-SPOnline -Url $siteurl -Credentials($global:credentials) -ErrorVariable $connecterror
        if ($? -eq $false)
        {
            Write-Error "Failed to connect to $siteurl, resetting credentials"
            $global:credentials = $null
            exit
        }
        
        $fullParentUrl = $siteurl + $parentweburl

        if ($parentweburl -eq "/")
        {
           $parentweb = Get-SPOWeb
        }
        else
        {
            $parentweb = Get-SPOSubWebs -Recurse | where {$_.Url -eq $fullParentUrl} 
        }

        if ($parentweb  -eq  $null)
        {
			Write-Error "Create Parent web $fullparenturl first"

            $wshell = New-Object -ComObject Wscript.Shell
            $intBtn = $wshell.Popup("Create parent site first: $fullparenturl",0,"Parent Site does no Exist",0x0 + 0x40 + 0x1000)

			exit
		}
		else
        {
            $web = $null
         
            if ($weburl -eq "/")
            {
               $fullWebUrl = $siteurl + $parentweburl

               $web = Get-SPOWeb
            }
            else
            {
                if ($parentweburl -eq "/")
                {
                    $fullWebUrl = $siteurl + "/" + $weburl 
                }
                else
                {
                    $fullWebUrl = $siteurl + $parentweburl + "/" + $weburl 
                }


                $web = Get-SPOSubWebs -Recurse | where {$_.Url -eq $fullWebUrl}

                if ($web -eq $null)
                {
                    Write-Host "Creating site: $siteTitle at url: $weburl using template: $sitetemplate"

                    $web = New-SPOWeb -Url $weburl -Title $sitetitle -Template $sitetemplate -Web $parentweb

                    $web = Get-SPOSubWebs -Recurse | where {$_.Url -eq $fullWebUrl}
                }
            }
 
            $web.SiteLogoUrl = "/Style%20Library/bluesource/logo_block.png"
            $web.AlternateCssUrl = "/Style%20Library/bluesource/bluesource.css"
            #$web.AllProperties["__InheritsAlternateCssUrl"] = $True
            $web.Update()
        } 

        #Set the Publishing Web NOT NEEDED AS PAGES NOT TICKED ANYWAY
        #    write-host "fullWebUrl = $fullWebUrl" 
        #    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($fullWebUrl)
        #    $SPWeb = $ctx.Web
        #    $SPPubWeb = [Microsoft.SharePoint.Client.Publishing.PublishingWeb]::GetPublishingWeb($ctx, $SPWeb)
        #    $ctx.Load($SPPubWeb)
        #    $SPPubWeb.AllowUnsafeUpdates = $true
        #    $SPPubWeb.Navigation.InheritCurrent = $false
        #    $SPPubWeb.Navigation.ShowSiblings = $false
        #    $SPPubWeb.Navigation.CurrentIncludeSubSites = $false
        #    $SPPubWeb.Navigation.CurrentIncludePages = $false
        #    $SPPubWeb.AllowUnsafeUpdates = $false
        #    $SPPubWeb.Update()
        
        if($applytemplate -eq "true")
        {
            $siteAssetTemplateXML = "$path\Data\WebTemplates\$siteassets\template.xml"
            
            
            # Where did the template get generated from 
            $sourceSiteUrl = Get-Setting -settingName "source" -XMLFilePath $XMLFilePath  
 

            $sspidSource = Get-EnvironmentSetting -environment $sourceSiteUrl -settingName "SSPID"  -XMLFilePath $XMLFilePath  
            $sspidNew = Get-EnvironmentSetting -environment $siteurl -settingName "SSPID" -XMLFilePath $XMLFilePath  

            $updatedFile = "$siteAssetTemplateXML.updated"
            # deleting previous version in case the creation of the fiel doesn't work.
            
            Remove-Item $updatedFile |Out-Null

            set-spotracelog -on -level debug             
 
            (Get-Content $siteAssetTemplateXML).replace($sspidSource,$sspidNew ).Replace("’","&apos;") | Set-Content -Path $updatedFile -Encoding UTF8

            

            Apply-SPOProvisioningTemplate -Path $updatedFile  -Web $web -Verbose -ExcludeHandlers Files -schema

            
            if ($? -eq $false)
            {
                "Apply-SPOProvisioningTemplate -Path $updatedFile -Web $web failed" | Write-Error 
                exit
            }
        }

        Add-Pages -siteassets $siteassets -Web $web

        $wshell = New-Object -ComObject Wscript.Shell
        $intBtn = $wshell.Popup("Your site has been created: $fullWebUrl",0,"Site Created",0x0 + 0x40 + 0x1000)

    }

    end
    {

        write-debug "End Add-Site"
        write-Host "End Add-Site"

    }


}