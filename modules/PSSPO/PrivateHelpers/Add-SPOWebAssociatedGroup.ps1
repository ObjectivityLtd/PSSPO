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

function Add-SPOWebAssociatedGroup {
    <#

    .SYNOPSIS

    Add SharePoint default security group of given type.



    .DESCRIPTION

    The Add-SPOWebAssociatedGroup function checks if SharePoint security group (members, owners or visitors) exists for given site. If not, the group with default name and description is created.

    Parameters necessary to create group are expected as pipeline properties of an object, or may be passed to the -GroupType and -PermissionName parameters.



    .PARAMETER Web 

    The site default group is created for.



    .PARAMETER GroupType

    Associated group type. The parameter takes one of the following values: Members, Owners, Visitors.



    .EXAMPLE

    Add-SPOWebAssociatedGroup -Web $web -GroupType "Owners" -PermissionName "design"



    .EXAMPLE

    Add-SPOWebAssociatedGroup -Web $web -GroupType "Members"


    
    .EXAMPLE

    $groupInfo | Add-SPOWebAssociatedGroup -Web $web


    
    .NOTES

    You need to pass 'Web' argument that is loaded in the context of a user who has privileges to create SharePoint security groups.

    #>
	[CmdletBinding()]
    [OutputType([Microsoft.SharePoint.Client.Group])]
	param(
        [Parameter(Mandatory=$true, Position=1)]
        [Microsoft.SharePoint.Client.Web]$Web,

        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=2)]
        [ValidateSet("Members", "Owners", "Visitors")]
        [string]$GroupType,

        [Parameter(Mandatory=$false, ValueFromPipeline=$true, Position=3)]
        [string]$PermissionName
    )

    begin {
        Write-Debug -Message "### Add-SPOWebAssociatedGroup begin ###"

        $permissions = @{
            "Members" = "contribute"
            "Owners" = "full control"
            "Visitors" = "read"
        }

        Write-Debug -Message "Loading client objects."
        $ctx = $Web.Context
        $site = $ctx.Site
        $ctx.Load($site)
        $ctx.Load($Web)
        $ctx.ExecuteQuery()
        Write-Debug -Message "Query execution finished."
	}
	
	process {
        
        Write-Debug -Message "### Add-SPOWebAssociatedGroup process ###"

        if (-not $PermissionName) {
            Write-Verbose -Message "Permission name not set - default permission name will be used for $GroupType group description."
            $PermissionName = $permissions[$GroupType]
        }
        $name = "$($Web.Title) $GroupType"
        $description = "Use this group to grant people $PermissionName permissions to the SharePoint site&#58; <a href=""$($Web.ServerRelativeUrl)"">$($Web.Title)</a>"

        $group = Add-SPOGroup -Site $site -Owner "$($Web.Title) Owners" -Name $name -Description $description

        Write-Verbose -Message "Associated $GroupType group for site $($Web.Url): $name"

        return $group
	}
	
	end {
        Write-Debug -Message "### Add-SPOWebAssociatedGroup end ###"
	}
}