##
#
# Allow Powershell to use CSOM
# http://soerennielsen.wordpress.com/2013/08/25/use-csom-from-powershell/
# SharePoint 2013 - http://www.microsoft.com/en-us/download/details.aspx?id=35585
# SharePoint Online - http://www.microsoft.com/en-us/download/details.aspx?id=42038
##

function Add-CSOM {
    $CSOMdir = "${env:CommonProgramFiles}\microsoft shared\Web Server Extensions\16\ISAPI"
    $excludeDlls = "*.Portable.dll"
    
    if ((Test-Path $CSOMdir -pathType container) -ne $true)
    {
        $CSOMdir = "${env:CommonProgramFiles}\microsoft shared\Web Server Extensions\15\ISAPI"
        if ((Test-Path $CSOMdir -pathType container) -ne $true)
        {
            Throw "Please install the SharePoint 2013[1] or SharePoint Online[2] Client Components`n `n[1] http://www.microsoft.com/en-us/download/details.aspx?id=35585`n[2] http://www.microsoft.com/en-us/download/details.aspx?id=42038`n `n "
        }
    }
    
    
    $CSOMdlls = Get-Item "$CSOMdir\*.dll" -exclude $excludeDlls
    
    ForEach ($dll in $CSOMdlls) {
        [System.Reflection.Assembly]::LoadFrom($dll.FullName) | Out-Null
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
        $contentTypes = $web.AvailableContentTypes
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
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$ParentContentTypeName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web] $web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        
        $parentContentType = Get-ContentType -ContentTypeName $ParentContentTypeName -Web $web -Context $context
        $contentType = $null
        if($parentContentType -eq $null) {
            Write-Host "Error loading parent content type $ParentContentTypeName"
        } else {

            $contentTypeCreationInformation = New-Object Microsoft.SharePoint.Client.ContentTypeCreationInformation
            $contentTypeCreationInformation.Name = $Name
            $contentTypeCreationInformation.Description = "Create a new $Name"
            $contentTypeCreationInformation.Group = $Group
            $contentTypeCreationInformation.ParentContentType = $parentContentType
            
            $contentType = $web.ContentTypes.Add($contentTypeCreationInformation)
            $context.load($contentType)
            $context.ExecuteQuery()
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
                Write-Output "`tUpdated field link $fieldId for content type $($ContentType.Name)"
            } else {
                Write-Verbose "`tDid not update field link $fieldId for content type $($ContentType.Name)"
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
        $context.load($field)
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
function Delete-SiteColumn {
    param(
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$fieldId,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web] $web, 
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $field = Get-SiteColumn -FieldId $fieldId -Web $web -Context $context
        if($field -ne $null) {
            $field.DeleteObject()
            $context.ExecuteQuery()
        }
    }
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
function Delete-List {
    param(
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$ListName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web] $web, 
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $list = Get-List -ListName $ListName -Web $web -Context $context
        if($list -ne $null) {
            $list.DeleteObject()
            $context.ExecuteQuery()
        }
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


function Add-Web {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web]$web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][System.Xml.XmlElement]$xml,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {

        $webCreationInfo = New-Object Microsoft.SharePoint.Client.WebCreationInformation

        $webCreationInfo.Url = $xml.URL
        $webCreationInfo.Title = $xml.Title
        $webCreationInfo.Description = $xml.Description
        $webCreationInfo.WebTemplate = $xml.WebTemplate

        $newWeb = $web.Webs.Add($webCreationInfo); 
        $context.Load($newWeb);
        $context.ExecuteQuery()

        Setup-Web -web $newweb -xml $webInfo -context $context
        $newWeb
    }
    end {} 
}
function Add-Webs {

 
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web]$web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][System.Xml.XmlElement]$xml,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
   )
    process {

        foreach ($webInfo in $xml.Web) {
            $newweb = Add-Web -web $web -xml $webInfo -context $context 
        }
      
    }
    end {} 
}



function Setup-Web {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web]$web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][System.Xml.XmlElement]$xml,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        foreach ($List in $xml.Lists.RemoveList) {
            Delete-List -ListName $List.Title -Web $web -context $context
        }
        foreach ($ContentType in $xml.ContentTypes.RemoveContentType) {
            Delete-ContentType -ContentTypeName $ContentType.Name -Web $web -context $context
        }
        foreach ($Field in $xml.Fields.RemoveField) {
            Delete-SiteColumn -FieldId $Field.ID -Web $web -context $context
        }

        foreach ($Field in $xml.Fields.Field) {
            $fieldStr = $Field.OuterXml.Replace(" xmlns=`"http://schemas.microsoft.com/sharepoint/`"", "")
	        $SPfield = Get-SiteColumn -FieldId $Field.ID -Web $web -context $context
	        if($SPfield -eq $null) {
		        $SPfield = Add-SiteColumn -FieldXml $fieldStr -Web $web -context $context
                if(($Field.Type -eq "TaxonomyFieldType") -or ($Field.Type -eq "TaxonomyFieldTypeMulti")) {
                    $termSetId = $null
                    foreach($property in $Field.Customization.ArrayOfProperty.Property) {
                        if($property.Name -eq "TermSetId") {
                            $termSetId = $property.Value.InnerText
                        }
                    }

                    $taxonomySession = Get-TaxonomySession -context $context
                    $defaultSiteCollectionTermStore = Get-DefaultSiteCollectionTermStore -TaxonomySession $taxonomySession -context $context
                    $SPfield = [SharePointClient.PSClientContext]::CastToTaxonomyField($context, $SPfield)
                    $SPfield.SspId = $defaultSiteCollectionTermStore.Id
                    $SPfield.TermSetId = $termSetId
                    $SPfield.UpdateAndPushChanges($true)
                    $context.load($SPfield)
                    $context.ExecuteQuery()
                } 
		        Write-Output "Created Site Column $($Field.Name)"
	        } else {
		        Write-Verbose "Site Column $($Field.Name) already exists"
	        }
        }

        foreach ($ContentType in $xml.ContentTypes.ContentType) {
            $SPContentType = Get-ContentType -ContentTypeName $ContentType.Name -Web $web -context $context
            if($SPContentType -eq $null) {
                $SPContentType = Add-ContentType -Name $ContentType.Name -Description $ContentType.Description -Group $ContentType.Group -ParentContentTypeName $ContentType.ParentContentType -Web $web -context $context
                if($SPContentType -eq $null) {
                    Write-Error "Could Not Create Content Type $($ContentType.Name)"
                    break;
                } else {
                    Write-Output "Created Content Type $($ContentType.Name)"
                }
            } else  {
                Write-Verbose "Content Type $($ContentType.Name)  already created."
            }

            foreach ($FieldRef in $ContentType.FieldRefs.FieldRef) {
                $SPField = Get-FieldForContentType -FieldId $FieldRef.ID -ContentType $SPContentType -context $context
                if($SPField -eq $null) {
                    $SPFieldLink = Add-FieldToContentType -FieldId $FieldRef.ID -ContentType $SPContentType -Web $web -context $context
                    Write-Output "`tAdded field $($FieldRef.ID) to Content Type $($ContentType.Name)"
                } else {
                    Write-Verbose "`tField $($FieldRef.ID) already added to Content Type $($ContentType.Name)"
                }

                $Required = $null
                if($FieldRef.Required) {
                    $Required = [bool]::Parse($FieldRef.Required)
                }
                $Hidden = $null
                if($FieldRef.Hidden) {
                    $Hidden = [bool]::Parse($FieldRef.Hidden)
                }
                Update-ContentTypeFieldLink -FieldId $FieldRef.ID -Required $Required -Hidden $Hidden -ContentType $SPContentType -context $context
            }

            foreach ($RemoveFieldRef in $ContentType.FieldRefs.RemoveFieldRef) {
                Remove-FieldFromContentType -FieldId $RemoveFieldRef.ID -ContentType $SPContentType -context $context
            }
        }

        foreach ($List in $xml.Lists.List) {
            $SPList = Get-List -ListName $List.Title -Web $web -context $context
            if($SPList -eq $null) {
                $SPList = New-List -ListName $List.Title -Type $List.Type -Url $List.Url -Web $web -context $context
                Write-Output "List created: $($List.Title)"
            } else {
                Write-Verbose "`List already created: $($List.Title)"
            }

            Write-Output "`tContent Types"
	        foreach ($ct in $List.ContentType) {
                $spContentType = Get-ListContentType -List $SPList -ContentTypeName $ct.Name -context $context
		        if($spContentType -eq $null) {
                    $spContentType = Add-ListContentType -List $SPList -ContentTypeName $ct.Name -Web $web -context $context
                    if($spContentType -eq $null) {
                        Write-Error "`t`tContent Type could not be added: $($ct.Name)"
                    } else {
                        Write-Output "`t`tContent Type added: $($ct.Name)"
                    }
                } else {
                    Write-Verbose "`t`tContent Type already added: $($ct.Name)"
                }
	        }
            foreach ($ct in $List.RemoveContentType) {
                $spContentType = Get-ListContentType -List $SPList -ContentTypeName $ct.Name -context $context
		        if($spContentType -ne $null) {
                    Delete-ListContentType -List $SPList -ContentTypeName $ct.Name -context $context
                    Write-Output "`t`tContent Type deleted: $($ct.Name)"
                } else {
                    Write-Verbose "`t`tContent Type already deleted: $($ct.Name)"
                }
            }

            

            Write-Output "`tFields"
            foreach($field in $List.Fields.Field){
                $spField = Get-ListField -List $SPList -FieldName $Field.Name -Context $context
                if($spField -eq $null) {
                    $fieldStr = $field.OuterXml.Replace(" xmlns=`"http://schemas.microsoft.com/sharepoint/`"", "")
                    $spField = New-ListField -FieldXml $fieldStr -List $splist -context $context
                    Write-Output "`t`tCreated Field: $($Field.DisplayName)"
                } else {
                    Write-Verbose "`t`tField already added: $($Field.DisplayName)"
                }
            }
            foreach($Field in $List.Fields.UpdateField) {
                $spField = Get-ListField -List $SPList -FieldName $Field.Name -Context $context
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
                    $context.ExecuteQuery()
                    Write-Output "`t`tUpdated Field: $($Field.DisplayName)"
                } else {
                    Write-Verbose "`t`tDid not need to update Field: $($Field.DisplayName)"
                }
            }
            foreach($Field in $List.Fields.RemoveField) {
                Delete-ListField -List $SPList -FieldName $Field.Name -Context $context
            }

            Write-Output "`tViews"
            foreach ($view in $List.Views.View) {
                $spView = Get-ListView -List $SPList -ViewName $view.DisplayName -context $context
                if($spView -ne $null) {
            
                    $Paged = [bool]::Parse($view.RowLimit.Paged)
                    $DefaultView = [bool]::Parse($view.DefaultView)
                    $RowLimit = $view.RowLimit.InnerText
                    $Query = $view.Query.InnerXml.Replace(" xmlns=`"http://schemas.microsoft.com/sharepoint/`"", "")
                    $ViewFields = $view.ViewFields.FieldRef | Select -ExpandProperty Name

                    $spView = Update-ListView -List $splist -ViewName $view.DisplayName -Paged $Paged -Query $Query -RowLimit $RowLimit -DefaultView $DefaultView -ViewFields $ViewFields -context $context
                    Write-Output "`t`tUpdated List View: $($view.DisplayName)"
                } else {
            
                    $Paged = [bool]::Parse($view.RowLimit.Paged)
                    $PersonalView = [bool]::Parse($view.PersonalView)
                    $DefaultView = [bool]::Parse($view.DefaultView)
                    $RowLimit = $view.RowLimit.InnerText
                    $Query = $view.Query.InnerXml.Replace(" xmlns=`"http://schemas.microsoft.com/sharepoint/`"", "")
                    $ViewFields = $view.ViewFields.FieldRef | Select -ExpandProperty Name
                    $ViewType = $view.Type
                    $spView = New-ListView -List $splist -ViewName $view.DisplayName -Paged $Paged -PersonalView $PersonalView -Query $Query -RowLimit $RowLimit -DefaultView $DefaultView -ViewFields $ViewFields -ViewType $ViewType -context $context
                    Write-Output "`t`tCreated List View: $($view.DisplayName)"
                }
            }
        }

        if($xml.Webs) {
            Add-Webs -Web $web -Xml $xml.Webs -Context $ctx
        }
    }
}

function UnSetup-Web {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Web]$web,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][System.Xml.XmlElement]$xml,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        foreach ($List in $xml.Lists.List) {
            Delete-List -ListName $ContentType.Title -Web $web -context $context
        }
        foreach ($ContentType in $xml.ContentTypes.ContentType) {
            Delete-ContentType -ContentTypeName $ContentType.Name -Web $web -context $context
        }
        foreach ($Field in $xml.Fields.Field) {
            Delete-SiteColumn -FieldId $Field.ID -Web $web -context $context
        }
    }
}

# The taxonomy code is untested

function Get-TaxonomySession {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($context)
        $session.UpdateCache()
        $session
    }
}
function Get-DefaultSiteCollectionTermStore {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]$TaxonomySession,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $store = $TaxonomySession.GetDefaultSiteCollectionTermStore()
        $context.Load($store)
        $context.ExecuteQuery()
        $store
    }
}

function Get-TermGroup {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$GroupName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Taxonomy.TermStore]$TermStore,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $group = $TermStore.Groups.GetByName($GroupName)
        $context.Load($group)
        $context.ExecuteQuery()
        $group
    }
}
function Add-TermGroup {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$Name,
        [parameter(ValueFromPipelineByPropertyName = $true)][guid]$Id = [guid]::NewGuid(),
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Taxonomy.TermStore]$TermStore,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $group = $TermStore.CreateGroup($Name,$Id)
        $TermStore.CommitAll()
        $context.load($group)
        $context.ExecuteQuery()
        $group
    }
}

function Get-TermSet {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$SetName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Taxonomy.TermGroup]$TermGroup,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $termSet = $TermGroup.TermSets.GetByName($SetName)
        $context.Load($termSet)
        $context.ExecuteQuery()
        $termSet
    }
}
function Add-TermSet {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$Name,
        [parameter(ValueFromPipelineByPropertyName = $true)][int]$Language = 1033,
        [parameter(ValueFromPipelineByPropertyName = $true)][guid]$Id = [guid]::NewGuid(),
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Taxonomy.TermGroup]$TermGroup,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $termSet = $TermGroup.CreateTermSet($Name, $Id, $Language)
        $TermGroup.TermStore.CommitAll()
        $context.load($termSet)
        $context.ExecuteQuery()
        $termSet
    }
}
function Add-Term {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName = "Name")][string]$Name,
        [parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = "Language")][int]$Language = 1033,
        [parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = "Id")][guid]$Id = [guid]::NewGuid(),
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Taxonomy.TermSet]$TermSet,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $term = $TermSet.CreateTerm($Name, $Language, $Id)

        $TermSet.TermStore.CommitAll()
        $context.load($term)
        $context.ExecuteQuery()
        $term
    }
}
function Get-Term {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][guid]$Id,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Taxonomy.TermSet]$TermSet,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $term = $TermSet.GetTerm($Id)
        $context.Load($term)
        $context.ExecuteQuery()
        $term
    }
}
function Get-Terms {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Taxonomy.TermSet]$TermSet,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $terms = $TermSet.Terms
        $context.Load($terms)
        $context.ExecuteQuery()
        $terms
    }
}
function Get-ChildTerms {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Taxonomy.Term]$Term,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $terms = $Term.Terms
        $context.Load($terms)
        $context.ExecuteQuery()
        $terms
    }
}

function Get-TermsByName {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$Name,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Taxonomy.TermSet]$TermSet,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $LabelMatchInformation = New-Object Microsoft.SharePoint.Client.Taxonomy.LabelMatchInformation($context);
        $LabelMatchInformation.Lcid = 1033
        $LabelMatchInformation.TrimUnavailable = $false         
        $LabelMatchInformation.TermLabel = $Name

        $terms = $TermSet.GetTerms($LabelMatchInformation)
        $context.Load($terms)
        $context.ExecuteQuery()
        $terms
    }
}

function Add-ChildTerm {
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$Name,
        [parameter(ValueFromPipelineByPropertyName = $true)][int]$Language = 1033,
        [parameter(ValueFromPipelineByPropertyName = $true)][guid]$Id = [guid]::NewGuid(),
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Taxonomy.Term]$parentTerm,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$context
    )
    process {
        $term = $parentTerm.CreateTerm($Name, $Language, $Id)

        $parentTerm.TermStore.CommitAll()
        $context.load($term)
        $context.ExecuteQuery()
        $term
    }
}