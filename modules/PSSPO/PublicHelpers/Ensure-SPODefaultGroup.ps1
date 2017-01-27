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

function Ensure-SPODefaultGroup {
    <#

    .SYNOPSIS

    Ensures SharePoint default security group of given type exists for site.



    .DESCRIPTION

    The Ensure-SPODefaultGroup function checks if SharePoint security group of given type exists for given site. If not, the group with default name and description is created.

    Parameters necessary to create site are expected as pipeline properties of an object, or may be passed to the -Name, -Description and -OwnerGroupName parameters.



    .PARAMETER web 

    The parent web.



    .PARAMETER Name

    Name of the security group.



    .PARAMETER Description

    Description of the security group.



    .PARAMETER OwnerGroupName

    Name of the group security group that will be used as an owner of the ensured group (only if the group is created).



    .EXAMPLE

    Ensure 'SAMPLE Owners' security group exists. If the group does not exist, create it and make it owner of itself.


    Ensure-SPOGroup -Web $context.Web -Name "SAMPLE Owners" -Description "Use this group to grant people full control permissions to the SharePoint site" -OwnerGroupName "SAMPLE Owners"



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
        [string]$GroupType
    )

    begin {
        Write-Debug -Message "Ensure-SPODefaultGroup begin"

        $permissions = @{
            "Members" = "contribute"
            "Owners" = "full control"
            "Visitors" = "read"
        }
	}
	
	process {
        
        Write-Debug -Message "Ensure-SPODefaultGroup process"

        $ctx = $Web.Context

        $name = "$($Web.Title) $GroupType"
        $description = "Use this group to grant people $($permissions[$GroupType]) permissions to the SharePoint site&#58; <a href=""$($Web.ServerRelativeUrl)"">$($Web.Title)</a>"

        $group = Ensure-SPOGroup -Web $Web -Owner "$($Web.Title) Owners" -Name $name -Description $description

        return $group
	}
	
	end {
        Write-Debug -Message "Ensure-SPODefaultGroup end"
	}
}