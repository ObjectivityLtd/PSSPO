<#
The MIT License (MIT)

Copyright (c) 2017 Objectivity Bespoke Software Specialists

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>

function Ensure-SPOWeb {
    <#

    .SYNOPSIS

    Ensures that SharePoint site of given URL address exists in parent's site subsites collection.



    .DESCRIPTION

    The creation is performed only if the site does not exist.

    Parameters necessary to create site are expected as pipeline properties of an object, or may be passed to the -Title, -Description, -Url and -Template parameters.



    .PARAMETER Web 

    The parent site.



    .PARAMETER Title

    Title of the child site.



    .PARAMETER Description

    Description of the child site.



    .PARAMETER Template

    Template that will be applied to the child site (like 'STS#0').



    .PARAMETER Url

    Url of the child site (relative to parent's site url).



    .PARAMETER UniquePermissions

    Switch parameter that indicates if new site should be crerated with unique permissions. If used, the following default groups are provided:

    - site owners (with 'full control' permissions to the site),

    - site members (with 'edit' permissions to the site),

    - site visitors (with 'read' permissions to the site).



    .EXAMPLE

    Ensure that 'Sample' site exists as a root site's subsite. If not, the site is created with unique permissions.


    Ensure-SPOWeb -Web $context.Web -Title "Sample" -Description "Sample site" -Template "STS#0" -Url "sample" -UniquePermissions



    .EXAMPLE 

    Read site's children from xml file and create missing ones (if there are any).


    (Get-Content -Path "c:\portal.xml).site.web.webs | Ensure-SPOWeb -Web $context.Site.RootWeb



    .Example

    Create (if not exists) 'Sample' child site for the root site by specifying parameters and passing parent web to the pipeline. The permissions will be inherited.


    $ctx.Web | Ensure-SPOWeb -Title "Sample" -Description "Sample site" -Template "STS#0" -Url "sample"



    .NOTES

    You need to pass 'Web' argument that is loaded in the context of a user who has privileges to create child sites on this web.

    Notice that function call does not change state of existing site (if found by url). This applies both to parameters (title, description and template) and site permissions.

    #>
    [CmdletBinding()]
    [OutputType([Microsoft.SharePoint.Client.Web])]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [Microsoft.SharePoint.Client.Web]$Web,

        [Parameter(Mandatory=$true, Position=2, ValueFromPipelineByPropertyName=$true)]
        [string]$Url,

        [Parameter(Mandatory=$true, Position=3, ValueFromPipelineByPropertyName=$true)]
        [string]$Template,

        [Parameter(Mandatory=$true, Position=4, ValueFromPipelineByPropertyName=$true)]
        [string]$Title,

        [Parameter(Mandatory=$false, Position=5, ValueFromPipelineByPropertyName=$true)]
        [string]$Description,

        [Parameter(Mandatory=$false, Position=6, ValueFromPipelineByPropertyName=$true)]
        [switch]$UniquePermissions
    )

    begin {
        Write-Debug -Message "Ensure-SPOWeb begin"
        $ctx = $Web.Context
        $ctx.Load($Web)
        $ctx.Load($Web.Webs)
        $ctx.ExecuteQuery()
    }

    process {
        Write-Debug -Message "Ensure-SPOWeb process: $Url"
                
        $childWebUrl = "$($Web.Url)/$Url"
        $childWeb = ($Web.Webs | Where-Object { $_.Url.ToLower() -eq $childWebUrl.ToLower() } | Select-Object -first 1)

        if ($childWeb) {
            Write-Verbose -Message "Site '$($childWeb.Url)' already exists - creation skipped."
        }
        else {
            try {
                Write-Verbose -Message "Creating child site '$Url' for '$($Web.Url)'"
                
                $webInfo = New-Object Microsoft.SharePoint.Client.WebCreationInformation
                $webInfo.WebTemplate = $Template
                $webInfo.Description = $Description
                $webInfo.Title = $Title
                $webInfo.Url = $Url
                $webInfo.Language = $Web.Language
                $webInfo.UseSamePermissionsAsParentSite = (-not $UniquePermissions)
                $childWeb = $Web.Webs.Add($webInfo)
                $ctx.Load($childWeb)
                $ctx.ExecuteQuery()
                
                Write-Host -Object "Site '$childWebUrl' was successfully created."

            } catch {
                Write-Error -Message "Site '$childWebUrl' creation failed."
                Write-Error $_
            }


            if ($UniquePermissions) {
                try {
                    Write-Verbose -Message "Breaking permission inheritance on site '$($childWeb.Url)'."
                    $childWeb.BreakRoleInheritance($false, $false)
                    $childWeb.Update()
                    $ctx.Load($childWeb)
                    $ctx.ExecuteQuery()
                    Write-Verbose -Message "Permission inheritance successfully broken."

                    $childWeb.AssociatedOwnerGroup = Ensure-SPODefaultGroup -Web $childWeb -GroupType "Owners" | Add-SPOGroupPermissions -TargetWeb $childWeb -PermissionLevel "Design"
                    $childWeb.Update()
                    $childWeb.AssociatedMemberGroup = Ensure-SPODefaultGroup -Web $childWeb -GroupType "Members" | Add-SPOGroupPermissions -TargetWeb $childWeb -PermissionLevel "Contribute"
                    $childWeb.Update()
                    $childWeb.AssociatedVisitorGroup = Ensure-SPODefaultGroup -Web $childWeb -GroupType "Visitors" | Add-SPOGroupPermissions -TargetWeb $childWeb -PermissionLevel "Read"
                    $childWeb.Update()
                    
                    # Site property 'vti_createdassociategroups' are comma-separated identifiers of associated groups. When site with unique permissions is created using gui, this property is set.
                    # Setting associated groups does not set the property. The code below is a workaround.
                    $ids = @($childWeb.AssociatedOwnerGroup.Id, $childWeb.AssociatedMemberGroup.Id, $childWeb.AssociatedVisitorGroup.Id)
                    $associatedGroups = $ids -join ';'
                    $childWeb.AllProperties["vti_createdassociategroups"] = $associatedGroups
                    $childWeb.Update()

                    $ctx.ExecuteQuery()

                    Write-Host -Object "Unique permissions successfully set for site '$($childWeb.Url)'."
                } catch {
                    Write-Error -Message "Setting unique permissions for site '$($childWeb.Url)' failed."
                    Write-Error $_
                }
            }
        }
    
        return $childWeb
    }

    end {
        Write-Debug -Message "Ensure-SPOWeb end"
    }
}