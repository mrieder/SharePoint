<###################################################################################
# Adds defined Group of Content Types, creates defined Folders and creates a OneNote
# Notebook in a list of Document Libraries
#
# DESCRIPTION
#
# Created By: Matthias Rieder, 2016-08-02
#
#             $WebUrl - Url to SharePoint Web
#             $SourceFile - Path to inputfile          
#
# Usage: ./SP-ActivateListFeatures.ps1 -web <String> -sourcefile <String>
###################################################################################>

#Parameter Definition
param
(
  #[PARAMETER(Mandatory=$false,HelpMessage="Keyword for Running Mode")][ValidateSet("WhatIf", "Execute")][String]$RunningMode="WhatIf",
  [PARAMETER(Mandatory=$true,HelpMessage="Url to SharePoint Web")][String]$WebUrl,
  [PARAMETER(Mandatory=$true,HelpMessage="Path to input file in txt format")][String]$SourceFile
  )

#Check if PS snap-in is loaded an load if not
$spsnapin = Get-PSSnapin -Name Microsoft.Sharepoint.Powershell –erroraction SilentlyContinue 
if (!$spsnapin) { "Loading VMware PowerShell snap-in..."; Add-PSSnapin Microsoft.Sharepoint.Powershell }

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

#import List for Folder creation
try 
{
  $inputlist = Get-Content $SourceFile -ErrorAction Stop
}
catch
{
  Write-Host "Please insert a valid Path to txt File" -ForegroundColor Red
  exit
}

#Changing the Document Library Title
Write-Host "Changing List Settings to defined policy for Customer Library" -ForegroundColor Green
foreach ($docLibrary in ($Web.Lists | ? {$_.Title -match '^[0-9]{1,6}.*$' -and $_.Title -notmatch '^000000.*$'}))
{
    try 
    {
        #Enabling "Allow management of content types"
        $docLibrary.ContentTypesEnabled = $true
         #Add content type to list
        foreach ($ctToAdd in $Web.Site.RootWeb.ContentTypes |? {$_.Group -eq "ACS Custom Content Types"})
        {
            #check if Content Type is already in list
            if (!$docLibrary.ContentTypes[$ctToAdd.Name]) {
                $docLibrary.ContentTypes.Add($ctToAdd) | Out-Null
                Write-Host "List "-NoNewline; Write-Host $docLibrary.Title -ForegroundColor Yellow -NoNewline
                Write-Host " - added content type: " -NoNewline; Write-Host $ctToAdd.Name -ForegroundColor Yellow 
            }
        }
        
        $docLibrary.Update()

        #create defined Folders
        foreach ($name in $inputlist)
        {
            #check if Folder already exits
            if (!($docLibrary.Folders | ? {$_.Name -eq $name}))
            {
                $folder = $docLibrary.AddItem("",[Microsoft.SharePoint.SPFileSystemObjectType]::Folder,$name.ToString())
                $folder.Update()
                Write-Host "List "-NoNewline; Write-Host $docLibrary.Title -ForegroundColor Yellow -NoNewline
                Write-Host " - added folder: " -NoNewline; Write-Host $name -ForegroundColor Yellow 
            }
        }
                
        #check for exisiting NoteBook
        $NoteBookPresent = $false
          
        if($docLibrary.Folders | ? {$_.ProgId -eq "OneNote.Notebook"})
        {
            $NoteBookPresent = $true
        }
      
        #creating OneNote notebook
        if(!$NoteBookPresent)
        {
            $folder = $docLibrary.AddItem("",[Microsoft.SharePoint.SPFileSystemObjectType]::Folder,"Notebook")
            $folder.ProgId ="OneNote.Notebook"
            $folder.Update()
            Write-Host "List "-NoNewline; Write-Host $docLibrary.Title -ForegroundColor Yellow -NoNewline
            Write-Host " - Notebook created"
        }

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

