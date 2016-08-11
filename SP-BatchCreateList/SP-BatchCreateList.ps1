<###################################################################################
# Creates SharePoint Document Library with unput file (SP-BatchListCreate.ps1)
#
# DESCRIPTION
#
# Created By: Matthias Rieder, 2016-08-02
#
# Variables:  $WebUrl - Url to SharePoint Web
#             $SourceFile - Path to the input file (csv format)              
#             $Template - Template for library creation (optional)
#
# Usage: ./SP-BatchCreateList.ps1 -web <String> -sourcefile <String>
###################################################################################>

#Parameter Definition
param
(
  #[PARAMETER(Mandatory=$false,HelpMessage="Keyword for Running Mode")][ValidateSet("WhatIf", "Execute")][String]$RunningMode="WhatIf",
  [PARAMETER(Mandatory=$true,HelpMessage="Url to SharePoint Web")][String]$WebUrl,
  [PARAMETER(Mandatory=$true,HelpMessage="Path to input file in csv format")][String]$SourceFile,
  [PARAMETER(Mandatory=$false,HelpMessage="Template for library creation")][String]$TemplateName
)

#Check if PS snap-in is loaded an load if not
$spsnapin = Get-PSSnapin -Name Microsoft.Sharepoint.Powershell –erroraction SilentlyContinue 
if (!$spsnapin) { "Loading VMware PowerShell snap-in..."; Add-PSSnapin Microsoft.Sharepoint.Powershell }

#Open Web and Site Collection
try 
{
  $Web = Get-SPWeb -Identity $WebUrl -ErrorAction Stop
  $Site = Get-SPSite -Identity $web.Site.RootWeb.Url.ToString()
}
catch
{
  Write-Host "Please insert a valid URL for the SharePoint web" -ForegroundColor Red
  exit
}
#import Input-File
try 
{
  $inputlist = Import-Csv $SourceFile -ErrorAction Stop
}
catch
{
  Write-Host "Please insert a valid Path to csv File" -ForegroundColor Red
  exit
}

#Checking Template
if($TemplateName) 
{
    Write-Host "Specified template" $TemplateName -ForegroundColor Yellow
    $CustomTemplate = $site.GetCustomListTemplates($Web) | ? {$_.Name -eq $TemplateName.ToString()}
    if ($CustomTemplate)
    {
        Write-Host "Template found." -ForegroundColor Green
        Write-Host "Using template" $CustomTemplate.Name

        #Creating the Document Library with custom template     
        foreach ($item in $inputlist)
        {
            Write-Host "Creating Library " $item.Title -ForegroundColor Yellow
            $Web.Lists.Add($item.Number,$item.Description,$CustomTemplate) | Out-Null
            #change List Name
            $Library = $Web.Lists[$item.Number.ToString()]
            $Library.Title = $item.Title
            $Library.Update()
        }
    }
    else
    {
        Write-Host "No custom template with name"$TemplateName "found" -ForegroundColor Red
        Write-Host "Usable custom templates for this site collection are:"
        Write-Host $site.GetCustomListTemplates($Web).Name -ForegroundColor Cyan
        exit
    }
}
else
{
    write-host "No template specified" -ForegroundColor Yellow
    write-host "...using default Document Library (to implement)"
    $template = $web.ListTemplates | ? {$_.InternalName -eq "doclib"}
    #CREATE LISTS with default templates
}
