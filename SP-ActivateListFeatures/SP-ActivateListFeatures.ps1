<###################################################################################
# Creates SharePoint Document Library with unput file (SP-BatchListCreate.ps1)
#
# DESCRIPTION
#
# Created By: Matthias Rieder, 2016-08-02
#
#             $WebUrl - Url to SharePoint Web          
#
# Usage: ./SP-ActivateCustomerListFeatures.ps1 -web <String> -sourcefile <String>
###################################################################################>

#Parameter Definition
param
(
  #[PARAMETER(Mandatory=$false,HelpMessage="Keyword for Running Mode")][ValidateSet("WhatIf", "Execute")][String]$RunningMode="WhatIf",
  [PARAMETER(Mandatory=$true,HelpMessage="Url to SharePoint Web")][String]$WebUrl
  )

#Check if PS snap-in is loaded an load if not
$spsnapin = Get-PSSnapin -Name Microsoft.Sharepoint.Powershell –erroraction SilentlyContinue 
if (!$spsnapin) { "Loading VMware PowerShell snap-in..."; Add-PSSnapin Microsoft.Sharepoint.Powershell }

#define Folders to create in docLibrary
$FoldersToCreate = @("Archive","Incident Reports","Licenses","Mail Communication")

#Open Web and Site Collection
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

#Changing the Document Library Title
Write-Host "Changing List Settings to defined policy for Customer Library"
foreach ($docLibrary in ($Web.Lists | ? {$_.Title -match '^[0-9]{1,6}.*$' -and $_.Title -notmatch '^000000.*$'}))
{
    Write-Host "Updating... " -NoNewline; Write-Host $Library -ForegroundColor Yellow
  
    try 
    {
        #Enabling "Allow management of content types"
        $docLibrary.ContentTypesEnabled = $true
         #Add content type to list
        foreach ($ctToAdd in $Web.Site.RootWeb.ContentTypes |? {$_.Group -eq "ACS Custom Content Types"})
        {
            $docLibrary.ContentTypes.Add($ctToAdd)
            write-host "Content type" $ctToAdd.Name "added to list" $docLibrary.Title
        }
        
        #create defined Folders
        foreach ($name in $FoldersToCreate)
        {
            $folder = $docLibrary.AddItem("",[Microsoft.SharePoint.SPFileSystemObjectType]::Folder,$name.ToString())
            $folder.Update()
        }

        #creating OneNote notebook

        #check for exisiting NoteBook
        $NoteBookPresent = $false
          
        if($docLibrary.Items | ? { $_.Name -match '^\D+\.onetoc2$'})
        {
            $NoteBookPresent = $true
        }


        If(!$NoteBookPresent)
        {
            $folder = $docLibrary.AddItem("",[Microsoft.SharePoint.SPFileSystemObjectType]::Folder,$item.ToString())
            $folder.Update()
        }


        $docLibrary.Items


        ###

        $docLibrary.Update()
    }
    catch 
    {
        write-host “Caught an exception:” -ForegroundColor Red
        write-host “Exception Type: $($_.Exception.GetType().FullName)” -ForegroundColor Red
        write-host “Exception Message: $($_.Exception.Message)” -ForegroundColor Red
        exit
    }   
}
