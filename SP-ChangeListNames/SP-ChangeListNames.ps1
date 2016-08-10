<###################################################################################
# Creates SharePoint Document Library with unput file (SP-BatchListCreate.ps1)
#
# DESCRIPTION
#
# Created By: Matthias Rieder, 2016-08-02
#
#             $WebUrl - Url to SharePoint Web
#             $SourceFile - Path to the input file (csv format)              
#
# Usage: ./SP-BatchListCreate.ps1 -web <String> -sourcefile <String>
###################################################################################>

#Parameter Definition
param
(
  #[PARAMETER(Mandatory=$false,HelpMessage="Keyword for Running Mode")][ValidateSet("WhatIf", "Execute")][String]$RunningMode="WhatIf",
  [PARAMETER(Mandatory=$true,HelpMessage="Url to SharePoint Web")][String]$WebUrl,
  [PARAMETER(Mandatory=$true,HelpMessage="Path to input file in csv format")][String]$SourceFile
)

#Check if PS snap-in is loaded an load if not
$spsnapin = Get-PSSnapin -Name Microsoft.Sharepoint.Powershell –erroraction SilentlyContinue 
if (!$spsnapin) { "Loading VMware PowerShell snap-in..."; Add-PSSnapin Microsoft.Sharepoint.Powershell }

#Open Web and Site Collextion
try 
{
  $Web = Get-SPWeb -Identity $WebUrl -ErrorAction Stop
  $Site = Get-SPSite -Identity $Web.Site.RootWeb.Url.ToString()
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


#Changing the Document Library Title
Write-Host "Changing Liste Title and Description"
foreach ($item in $inputlist)
{
    #change List Name
    $Library = $Web.Lists | ? {$_.EntityTypeName -eq $item.Number.ToString()}
    Write-Host "Set List Title " -NoNewline;
    Write-Host $Library.Title -ForegroundColor Yellow -NoNewline
    Write-Host " to " -NoNewline
    Write-Host $item.Title -ForegroundColor Green

    $Library.Title = $item.Title
    $Library.Description = $item.Description
    $Library.Update()
}
