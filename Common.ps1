﻿# JLL Claims Auth
# $ctx = New-Object MSDN.Samples.ClaimsAuth.ClaimClientContext.GetAuthenticatedContext($siteUrl)

##
#
# Allow Powershell to use CSOM
# http://soerennielsen.wordpress.com/2013/08/25/use-csom-from-powershell/
#
##
$myScriptPath = (Split-Path -Parent $MyInvocation.MyCommand.Path)
Write-Host "Script Path $myScriptPath"
function Add-CSOM {
    $CSOMdlls = Get-Item "$myScriptPath\dlls\*.dll"
    
    ForEach ($dll in $CSOMdlls) {
        [System.Reflection.Assembly]::LoadFrom($dll.FullName)
    }
    $assemblies = $CSOMdlls | Select -ExpandProperty FullName
    Add-Type -ReferencedAssemblies $assemblies -TypeDefinition @"
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
namespace SharePointClient
{
    public class PSClientContext: ClientContext
    {
        public PSClientContext(string siteUrl)
            : base(siteUrl)
        {
        }
        // need a plain Load method here, the base method is a generic method
        // which isn't supported in PowerShell.
        public void Load(ClientObject objectToLoad)
        {
            base.Load(objectToLoad);
        }
        public static TaxonomyField CastToTaxonomyField (ClientContext ctx, Field field)
        {
            return ctx.CastTo<TaxonomyField>(field);
        }
        public static void Load (ClientContext ctx, ClientObject objectToLoad)
        {
            ctx.Load(objectToLoad);
        }
        public TaxonomyField CastToTaxonomyField (Field field)
        {
            return base.CastTo<TaxonomyField>(field);
        }
    }
}
"@
}

function Get-ContentType {
 
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$ContentTypeName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web] $web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $contentTypes = $web.ContentTypes

        $context.Load($contentTypes)
        $context.ExecuteQuery()

        $contentType = $contentTypes | Where {$_.Name -eq $ContentTypeName}
        $contentType
    }
    end {}
}
function Delete-ContentType {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$ContentTypeName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web] $web, 
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        
        $contentType = Get-ContentType -ContentTypeName $ContentTypeName -Web $web -Context $context
        if($contentType -ne $null) {
            $contentType.DeleteObject()
            $context.ExecuteQuery()
        }
    }
    end {}
}
function Add-ContentType {
 
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$Name,
        [parameter(ValueFromPipeline=$true)][string]$Description = "Create a new $Name",
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$Group,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$ParentContentType,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web] $web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $ctx = $context
        $contentTypes = $web.ContentTypes

        $ctx.Load($contentTypes)
        $ctx.ExecuteQuery()

        $parentCT = $contentTypes | Where {$_.Name -eq $ParentContentType}
        $ctx.Load($parentCT)
        $ctx.ExecuteQuery()

        $contentType = $null
        if($parentCT -eq $null) {
            Write-Host "Error loading parent content type $ParentContentType"
        } else {
            $ctCreation = New-Object Microsoft.SharePoint.Client.ContentTypeCreationInformation
            $ctCreation.Name = $Name
            $ctCreation.Description = $Description
            $ctCreation.Group = $Group
            $ctCreation.ParentContentType = $parentCT
            $contentType = $contentTypes.Add($ctCreation)
            $ctx.ExecuteQuery()
        }
        $contentType
    }
    end {}
}

function Add-ContentTypeRGB {
 
    param (
        [parameter(ValueFromPipeline=$true)][string]$Name = "RGB Document",
        [parameter(ValueFromPipeline=$true)][string]$Description = "Create a new $Name",
        [parameter(ValueFromPipeline=$true)][string]$Group = "RGB Content Types",
        [parameter(ValueFromPipeline=$true)][string]$ParentContentType = "Document",
        [parameter(ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $siteUrl = "http://jll.spdev13.local"
        $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
        #$ctx = $context
        $web = $ctx.Web
        $contentTypes = $web.ContentTypes

        $ctx.Load($contentTypes)
        $ctx.ExecuteQuery()

        $parentCT = $contentTypes | Where {$_.Name -eq $ParentContentType}
        $ctx.Load($parentCT)
        $ctx.ExecuteQuery()

        $contentType = $null
        if($parentCT -eq $null) {
            Write-Host "Error loading parent content type $ParentContentType"
        } else {
            $ctCreation = New-Object Microsoft.SharePoint.Client.ContentTypeCreationInformation
            $ctCreation.Name = $Name
            $ctCreation.Description = $Description
            $ctCreation.Group = $Group
            $ctCreation.ParentContentType = $parentCT
            $contentType = $contentTypes.Add($ctCreation)
            #$contentType
            #$ctx.ExecuteQuery()
        }
        $contentType
    }
    end {}
}

function Add-FieldToContentType {
 
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$FieldId,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ContentType]$ContentType,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web] $web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        
        $field = Get-SiteColumn -fieldId $FieldId -Web $web -context $context
        $fieldlink = $null
        if($field -eq $null) {
            Write-Host "Error getting field $FieldId"
        } else {
            $context.Load($ContentType.FieldLinks)
            $context.ExecuteQuery()
            $fieldlinkCreation = New-Object Microsoft.SharePoint.Client.FieldLinkCreationInformation
            $fieldlinkCreation.Field = $field
            $fieldlink = $ContentType.FieldLinks.Add($fieldlinkCreation)
            $ContentType.Update($true)
            $context.ExecuteQuery()
        }
        $fieldlink
    }
    end {}
}
function Get-FieldForContentType {
 
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$FieldId,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ContentType]$ContentType,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $fields = $ContentType.Fields
        $context.Load($fields)
        $context.ExecuteQuery()

        $field = $null
        $field = $fields | Where {$_.Id -eq $FieldId}
        $field
    }
    end {}
}
function Remove-FieldFromContentType {
 
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$FieldId,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ContentType]$ContentType,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $fieldLinks = $ContentType.FieldLinks
        $context.Load($fieldLinks)
        $context.ExecuteQuery()

        $fieldLink = $fieldLinks | Where {$_.Id -eq $FieldId}
        if($fieldLink -ne $null) {
            $fieldLink.DeleteObject()
            $ContentType.Update($true)
            $context.ExecuteQuery()
            Write-Output "Deleted field $fieldId from content type $($ContentType.Name)"
        } else {
            Write-Output "Field $fieldId already deleted from content type $($ContentType.Name)"
        }
    }
    end {}
}
function Update-ContentTypeFieldLink {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$FieldId,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][System.Nullable[bool]]$Required,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][System.Nullable[bool]]$Hidden,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ContentType]$ContentType,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $fieldLinks = $ContentType.FieldLinks
        $context.Load($fieldLinks)
        $context.ExecuteQuery()
        
        $fieldLink = $fieldLinks | Where {$_.Id -eq $FieldId}
        if($fieldLink -ne $null) {
            
            $needsUpdating = $false
            if($Required -ne $null -and $fieldLink.Required -ne $Required) {
                $fieldLink.Required = $Required
                $needsUpdating = $true
            }
            if($Hidden -ne $null -and $fieldLink.Hidden -ne $Hidden) {
                $fieldLink.Hidden = $Hidden
                $needsUpdating = $true
            }
            if($needsUpdating) {
                $ContentType.Update($true)
                $context.ExecuteQuery()
                Write-Output "Updated field link $fieldId for content type $($ContentType.Name)"
            } else {
                Write-Output "Did not update field link $fieldId for content type $($ContentType.Name)"
            }
        } else {
            Write-Error "Could not find field link $fieldId for content type $($ContentType.Name)"
        }
    }
    end {}
}

function Add-SiteColumn {
 
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$fieldXml,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web] $web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
   )
    process {
        $field = $web.Fields.AddFieldAsXml($fieldXml, $false, ([Microsoft.SharePoint.Client.AddFieldOptions]::AddToNoContentType))
        $context.ExecuteQuery()
        $field
    }
    end {} 
}
function Get-SiteColumn {
 
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$fieldId,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web] $web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
   )
    process {
        $fields = $web.Fields
        $context.Load($fields)
        $context.ExecuteQuery()

        $field = $null
        $field = $fields | Where {$_.Id -eq $fieldId}
        $field
    }
    end {} 
}
function Connect-ManagedMetadataColumn {
 
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$fieldId,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$termStore,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$termGroup,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$termSet,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web] $web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
   )
    process {
        $field = $web.Fields.GetById($fieldId)

        $session =  [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($context)
        $store = $session.TermStores.GetByName($termStore)
        $group = $store.Groups.GetByName($termGroup)
        $set = $group.TermSets.GetByName($termSet)
        
        $context.Load($field)
        $context.Load($store)
        $context.Load($set)        
        $context.ExecuteQuery()

        $taxField = [SharePointClient.PSClientContext]::CastToTaxonomyField($context, $field)
        $taxField.SspId = $store.Id
        $taxField.TermSetId = $set.Id
        $taxField.UpdateAndPushChanges($true)
        $context.ExecuteQuery()
    }
    end {}
}

function New-List {
 
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$ListName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$Type,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$Url,        
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web] $web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context

   )
    process {
        
        $listCreationInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
        $listCreationInfo.Title = $ListName
        $listCreationInfo.TemplateType = $Type
        $listCreationInfo.Url = $Url
        
        $list = $web.Lists.Add($listCreationInfo)
        $context.ExecuteQuery()
        $list
    }
    end {}
}
function Get-List {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$ListName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web] $web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $lists = $web.Lists
        $context.Load($lists)
        $context.ExecuteQuery()
        
        $list = $null
        $list = $lists | Where {$_.Title -eq $ListName}
        if($list -ne $null) {
            $context.Load($list)
            $context.ExecuteQuery()
        }
        $list
    }
}

function Get-ListView {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.List]$List,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$ViewName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $views = $list.Views
        $context.load($views)
        $context.ExecuteQuery()
        
        $view = $null
        $view = $views | Where {$_.Title -eq $ViewName}
        if($view -ne $null) {
            $context.load($view)
            $context.ExecuteQuery()
        }
        $view
    }
}
function New-ListView {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.List]$List,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$ViewName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][bool]$DefaultView,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][bool]$Paged,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][bool]$PersonalView,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$Query,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][int]$RowLimit,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string[]]$ViewFields,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$ViewType,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        
        $ViewTypeKind
        switch($ViewType) {
            "none"{$ViewTypeKind = [Microsoft.SharePoint.Client.ViewType]::None}
            "html"{$ViewTypeKind = [Microsoft.SharePoint.Client.ViewType]::Html}
            "grid"{$ViewTypeKind = [Microsoft.SharePoint.Client.ViewType]::Grid}
            "calendar"{$ViewTypeKind = [Microsoft.SharePoint.Client.ViewType]::Calendar}
            "recurrence"{$ViewTypeKind = [Microsoft.SharePoint.Client.ViewType]::Recurrence}
            "chart"{$ViewTypeKind = [Microsoft.SharePoint.Client.ViewType]::Chart}
            "gantt"{$ViewTypeKind = [Microsoft.SharePoint.Client.ViewType]::Gantt}
        }
        $vCreation = New-Object Microsoft.SharePoint.Client.ViewCreationInformation
        $vCreation.Paged = $Paged
        $vCreation.PersonalView = $PersonalView
        $vCreation.Query = $Query
        $vCreation.RowLimit = $RowLimit
        $vCreation.SetAsDefaultView = $DefaultView
        $vCreation.Title = $ViewName
        $vCreation.ViewFields = $ViewFields
        $vCreation.ViewTypeKind = $ViewTypeKind

        $view = $list.Views.Add($vCreation)
        $list.Update()
        $context.ExecuteQuery()
        $view
    }
}
function Update-ListView {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.List]$List,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$ViewName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][bool]$DefaultView,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][bool]$Paged,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$Query,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][int]$RowLimit,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string[]]$ViewFields,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        
        $view = Get-ListView -List $List -ViewName $ViewName -context $context
        
        if($view -ne $null) {
            $view.Paged = $Paged
            $view.ViewQuery = $Query
            $view.RowLimit = $RowLimit
            $view.DefaultView = $DefaultView
            #Write-Host $ViewFields
            $view.ViewFields.RemoveAll()
            ForEach ($vf in $ViewFields) {
                $view.ViewFields.Add($vf)
                #$ctx.Load($view.ViewFields)
                #$view.Update()
                #$List.Update()
                #$context.ExecuteQuery()
                #Write-Host "Add column $vf to view"
                #Write-Host $view.ViewFields
            }

            $view.Update()
            $List.Update()
            $context.ExecuteQuery()
        }
        $view
    }
}

function Get-ListContentType {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.List]$List,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$ContentTypeName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $contentTypes = $List.ContentTypes
        $context.load($contentTypes)
        $context.ExecuteQuery()
        
        $contentType = $null
        $contentType = $contentTypes | Where {$_.Name -eq $ContentTypeName}
        if($contentType -ne $null) {
            $context.load($contentType)
            $context.ExecuteQuery()
        }
        $contentType
    }
}
function Add-ListContentType {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.List]$List,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$ContentTypeName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web] $web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context

   )
    process {
        $contentTypes = $web.AvailableContentTypes
        $context.Load($contentTypes)
        $context.ExecuteQuery()

        $contentType = $contentTypes | Where {$_.Name -eq $ContentTypeName}
        if($contentType -ne $null) {
            if(!$List.ContentTypesEnabled) {
                $List.ContentTypesEnabled = $true
            }
            $ct = $List.ContentTypes.AddExistingContentType($contentType);
            $List.Update()
            $context.ExecuteQuery()
        } else {
            $ct = $null
        }
        $ct
    }
    end {}
}
function Delete-ListContentType {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.List]$List,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$ContentTypeName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context

   )
    process {
        $contentTypeToDelete = Get-ListContentType $List $context -ContentTypeName $ContentTypeName
        
        if($contentTypeToDelete -ne $null) {
            if($contentTypeToDelete.Sealed) {
                $contentTypeToDelete.Sealed = $false
            }
            $contentTypeToDelete.DeleteObject()
            $List.Update()
            $context.ExecuteQuery()
        }
    }
}

function New-ListField {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$FieldXml,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.List]$List,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
   )
    process {
        $field = $list.Fields.AddFieldAsXml($FieldXml, $true, ([Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue))
        $context.Load($field)
        $context.ExecuteQuery()
        $field
    }
    end {}
}
function Get-ListField {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$FieldName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.List]$List,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $Fields = $List.Fields
        $context.Load($Fields)
        $context.ExecuteQuery()
        
        $Field = $null
        $Field = $Fields | Where {$_.InternalName -eq $FieldName}
        $Field
    }
}
function Delete-ListField{
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$FieldName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.List]$List,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $Fields = $List.Fields
        $context.Load($Fields)
        $context.ExecuteQuery()
        
        $Field = $null
        $Field = $Fields | Where {$_.InternalName -eq $FieldName}
        if($Field -ne $null) {
            $Field.DeleteObject()
            $List.Update()
            $context.ExecuteQuery()
            Write-Output "`t`tDeleted List Field: $FieldName"
        } else {
            Write-Output "`t`tField not found in list: $FieldName"
        }
    }
}

function Add-Webs {

 
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web]$web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][System.Xml]$xml,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
   )
    process {

        foreach ($webInfo in $xml.Elements.Web) {
             $webCreationInfo = New-Object Microsoft.SharePoint.Client.WebCreationInformation

             $webCreationInfo.Url = $webInfo.URL
             $webCreationInfo.Title = $webInfo.Title
             $webCreationInfo.Description = $webInfo.Description
             $webCreationInfo.WebTemplate = $webInfo.WebTemplate

             $newWeb = $web.Webs.Add($webCreationInfo); 
             $context.Load($newWeb);
             $context.ExecuteQuery()
        }
      
    }
    end {} 
}



function Setup-Web {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web]$web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Xml]$xml,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        foreach ($list in $xml.elements.List) {
            $splist = Get-List -ListName $list.Title -Web $web -context $ctx
            if($splist -eq $null) {
                $splist = New-List -ListName $list.Title -Type $list.Type -Url $list.Url -Web $web -context $ctx
                Write-Output "List created: $($list.Title)"
            } else {
                Write-Output "List already created: $($list.Title)"
            }

            Write-Output "`tContent Types"
	        foreach ($ct in $list.ContentType) {
                $spContentType = Get-ListContentType -List $splist -ContentTypeName $ct.Name -context $ctx
		        if($spContentType -eq $null) {
                    $spContentType = Add-ListContentType -List $splist -ContentTypeName $ct.Name -Web $web -context $ctx
                    if($spContentType -eq $null) {
                        Write-Error "`t`tContent Type could not be added: $($ct.Name)"
                    } else {
                        Write-Output "`t`tContent Type added: $($ct.Name)"
                    }
                } else {
                    Write-Output "`t`tContent Type already added: $($ct.Name)"
                }
	        }
            foreach ($ct in $list.RemoveContentType) {
                $spContentType = Get-ListContentType -List $splist -ContentTypeName $ct.Name -context $ctx
		        if($spContentType -ne $null) {
                    Delete-ListContentType -List $splist -ContentTypeName $ct.Name -context $ctx
                    Write-Output "`t`tContent Type deleted: $($ct.Name)"
                } else {
                    Write-Output "`t`tContent Type already deleted: $($ct.Name)"
                }
            }

            Write-Output "`tViews"
            foreach ($view in $list.Views.View) {
                $spView = Get-ListView -List $splist -ViewName $view.DisplayName -context $ctx
                if($spView -ne $null) {
            
                    $Paged = [bool]::Parse($view.RowLimit.Paged)
                    $DefaultView = [bool]::Parse($view.DefaultView)
                    $RowLimit = $view.RowLimit.InnerText
                    $Query = $view.Query.InnerXml.Replace(" xmlns=`"http://schemas.microsoft.com/sharepoint/`"", "")
                    $ViewFields = $view.ViewFields.FieldRef | Select -ExpandProperty Name

                    $spView = Update-ListView -List $splist -ViewName $view.DisplayName -Paged $Paged -Query $Query -RowLimit $RowLimit -DefaultView $DefaultView -ViewFields $ViewFields -context $ctx
                    Write-Output "`t`tUpdated List View: $($view.DisplayName)"
                } else {
            
                    $Paged = [bool]::Parse($view.RowLimit.Paged)
                    $PersonalView = [bool]::Parse($view.PersonalView)
                    $DefaultView = [bool]::Parse($view.DefaultView)
                    $RowLimit = $view.RowLimit.InnerText
                    $Query = $view.Query.InnerXml.Replace(" xmlns=`"http://schemas.microsoft.com/sharepoint/`"", "")
                    $ViewFields = $view.ViewFields.FieldRef | Select -ExpandProperty Name
                    $ViewType = $view.Type
                    $spView = New-ListView -List $splist -ViewName $view.DisplayName -Paged $Paged -PersonalView $PersonalView -Query $Query -RowLimit $RowLimit -DefaultView $DefaultView -ViewFields $ViewFields -ViewType $ViewType -context $ctx
                    Write-Output "`t`tCreated List View: $($view.DisplayName)"
                }
            }

            Write-Output "`tFields"
            foreach($field in $list.Fields.Field){
                $spField = Get-ListField -List $splist -FieldName $Field.Name -Context $ctx
                if($spField -eq $null) {
                    $fieldStr = $field.OuterXml.Replace(" xmlns=`"http://schemas.microsoft.com/sharepoint/`"", "")
                    $spField = New-ListField -FieldXml $fieldStr -List $splist -context $ctx
                    Write-Output "`t`tCreated Field: $($Field.Name)"
                } else {
                    Write-Output "`t`tField already added: $($Field.Name)"
                }
            }
            foreach($Field in $list.Fields.UpdateField) {
                $spField = Get-ListField -List $splist -FieldName $Field.Name -Context $ctx
                $needsUpdate = $false
                if($Field.ValidationFormula) {
                    $ValidationFormula = $Field.ValidationFormula
                    $ValidationFormula = $ValidationFormula -replace "&lt;","<"
                    $ValidationFormula = $ValidationFormula -replace "&gt;",">"
                    $ValidationFormula = $ValidationFormula -replace "&amp;","&"
                    if($spField.ValidationFormula -ne $ValidationFormula) {
                        $spField.ValidationFormula = $ValidationFormula
                        $needsUpdate = $true
                    }
                }

                if($Field.ValidationMessage) {
                    if($spField.ValidationMessage -ne $Field.ValidationMessage) {
                        $spField.ValidationMessage = $Field.ValidationMessage
                        $needsUpdate = $true
                    }
                }

                if($needsUpdate -eq $true) {
                    $spField.Update()
                    $ctx.ExecuteQuery()
                    Write-Output "`t`tUpdated Field: $($Field.Name)"
                } else {
                    Write-Output "`t`tDid not need to update Field: $($Field.Name)"
                }
            }
            foreach($Field in $list.Fields.RemoveField) {
                Delete-ListField -List $splist -FieldName $Field.Name -Context $ctx
            }
        }
    }
}