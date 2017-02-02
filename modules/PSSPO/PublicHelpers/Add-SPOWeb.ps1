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

function Add-SPOWeb {
    <#

    .SYNOPSIS

    Adds child site.



    .DESCRIPTION

    Site creation is performed only if the site does not exist in parent's site child sites collection.

    Parameters necessary to create site are expected as pipeline properties of an object, or may be passed to the -Title, -Description, -Url and -Template parameters.



    .PARAMETER ParentWeb 

    Parent site.



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

    - site members (with 'contribute' permissions to the site),

    - site visitors (with 'read' permissions to the site).



    .EXAMPLE

    Add-SPOWeb -ParentWeb $web -Title "Sample" -Description "Sample site" -Template "STS#0" -Url "sample" -UniquePermissions



    .EXAMPLE 

    $siteInfo | Add-SPOWeb -ParentWeb $web



    .Example

    $web | Add-SPOWeb -Title "Sample" -Description "Sample site" -Template "STS#0" -Url "sample"



    .NOTES

    You need to pass 'ParentWeb' argument that is loaded in the context of a user who has privileges to create child sites for this site.

    Notice that function call does not change state of existing site (if found by url). This applies both to parameters (title, description and template) and site permissions.

    #>
    [CmdletBinding()]
    [OutputType([Microsoft.SharePoint.Client.Web])]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [Microsoft.SharePoint.Client.Web]$ParentWeb,

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
        Write-Debug -Message "### Add-SPOWeb begin ###"
        Write-Debug -Message "Loading client objects."
        $ctx = $ParentWeb.Context
        $ctx.Load($ParentWeb)
        $ctx.Load($ParentWeb.Webs)
        $ctx.ExecuteQuery()
        Write-Debug -Message "Query execution finished."
    }

    process {
        Write-Debug -Message "### Add-SPOWeb process: $Url ###"
                
        $childWebUrl = "$($ParentWeb.Url)/$Url"
        $childWeb = ($ParentWeb.Webs | Where-Object { $_.Url.ToLower() -eq $childWebUrl.ToLower() } | Select-Object -first 1)

        if ($childWeb) {
            Write-Warning -Message "Site '$($childWeb.Url)' already exists - creation skipped."
            Write-Debug -Message "Loading child web."
            $ctx.Load($childWeb)
            $ctx.ExecuteQuery()
            Write-Debug -Message "Query execution finished."
        }
        else {
            Write-Verbose -Message "Creating child site '$Url' for '$($ParentWeb.Url)'"
            try {
                Write-Debug -Message "Adding child site."
                $webInfo = New-Object Microsoft.SharePoint.Client.WebCreationInformation
                $webInfo.WebTemplate = $Template
                $webInfo.Description = $Description
                $webInfo.Title = $Title
                $webInfo.Url = $Url
                $webInfo.Language = $ParentWeb.Language
                $webInfo.UseSamePermissionsAsParentSite = (-not $UniquePermissions)
                $childWeb = $ParentWeb.Webs.Add($webInfo)
                $ctx.Load($childWeb)
                $ctx.ExecuteQuery()
                Write-Debug -Message "Query execution finished."
                Write-Verbose -Message "Site '$childWebUrl' was successfully created."

            } catch {
                Write-Error -Message "Site '$childWebUrl' creation failed."
            }

            if ($UniquePermissions) {
                Write-Verbose -Message "Setting up unique permissions on site '$($childWeb.Url)'."
                try {
                    Write-Debug -Message "Associating security groups."
                    $owners = Add-SPOWebAssociatedGroup -Web $childWeb -GroupType "Owners"
                    Set-SPOPermissions -Target $childWeb -Principal $owners -PermissionLevel "Full Control"
                    $childWeb.AssociatedOwnerGroup =  $owners
                    $childWeb.Update()
                    $members = Add-SPOWebAssociatedGroup -Web $childWeb -GroupType "Members"
                    Set-SPOPermissions -Target $childWeb -Principal $members -PermissionLevel "Contribute"
                    $childWeb.AssociatedMemberGroup =  $members
                    $childWeb.Update()
                    $visitors = Add-SPOWebAssociatedGroup -Web $childWeb -GroupType "Visitors"
                    Set-SPOPermissions -Target $childWeb -Principal $visitors -PermissionLevel "Read"
                    $childWeb.AssociatedVisitorGroup = $visitors
                    $childWeb.Update()
                    # Site property 'vti_createdassociategroups' are comma-separated identifiers of associated groups. When site with unique permissions is created using gui, this property is set.
                    # Setting associated groups does not set the property. The code below is a workaround.
                    $ids = @($childWeb.AssociatedOwnerGroup.Id, $childWeb.AssociatedMemberGroup.Id, $childWeb.AssociatedVisitorGroup.Id)
                    $associatedGroups = $ids -join ';'
                    Write-Debug -Message "Associated groups: $associatedGroups."
                    $childWeb.AllProperties["vti_createdassociategroups"] = $associatedGroups
                    $childWeb.Update()
                    $ctx.Load($childWeb)
                    $ctx.ExecuteQuery()
                    Write-Debug -Message "Query execution finished."
                    Write-Verbose -Message "Unique permissions successfully set up for site '$($childWeb.Url)'."
                } catch {
                    Write-Error -Message "Setting up unique permissions for site '$($childWeb.Url)' failed."
                }
            }
        }
    
        return $childWeb
    }

    end {
        Write-Debug -Message "### Add-SPOWeb end ###"
    }
}