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

function Ensure-SPOGroup {
    <#

    .SYNOPSIS

    Ensures that SharePoint security group of given name exists.



    .DESCRIPTION

    The Ensure-SPOGroup function checks if SharePoint security group of given name exists in site collection groups. If not, the group is created.

    Parameters necessary to create site are expected as pipeline properties of an object, or may be passed to the -Name, -Description and -Owner parameters.



    .PARAMETER Web 

    Parent site of the group.



    .PARAMETER Name

    Name of the SharePoint security group.



    .PARAMETER Description

    Description of the SharePoint security group.



    .PARAMETER Owner

    Name of the SharePoint security group that will be used as an owner of the created group.



    .EXAMPLE

    Ensure 'Sample Owners' SharePoint security group exists. If the group does not exist, create it and make it owner of itself.


    Ensure-SPOGroup -Web $Web -Name "Sample Owners" -Description "Use this group to grant people full control permissions to the SharePoint site" -Owner "Sample Owners"



    .NOTES

    You need to pass 'Web' argument that is loaded in the context of a user who has privileges to create SharePoint security groups.

    #>
	[CmdletBinding()]
    [OutputType([Microsoft.SharePoint.Client.Group])]
	param(
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline=$false)]
        [Microsoft.SharePoint.Client.Web]$Web,

        [Parameter(Mandatory=$true, Position=2, ValueFromPipelineByPropertyName=$true)]
        [string]$Name,

        [Parameter(Mandatory=$true, Position=3, ValueFromPipelineByPropertyName=$true)]
        [string]$Description,

        [Parameter(Mandatory=$false, Position=4, ValueFromPipelineByPropertyName=$true)]
        [string]$Owner
    )

	begin {
        Write-Debug -Message "Ensure-SPOGroup begin"

        $ctx = $Web.Context
        $ctx.Load($Web)
        $ctx.Load($Web.SiteGroups)

        $site = $ctx.Site

        $ctx.Load($site)
        $ctx.Load($site.RootWeb.SiteGroups)
        $ctx.ExecuteQuery()
	}
	
	process {
        
        Write-Debug -Message "Ensure-SPOGroup process: $Name"

        $group = $Web.SiteGroups | Where-Object { $_.Title -eq $Name } | Select-Object -First 1

        if ($group) {
            Write-Verbose -Message "SharePoint security group '$Name' already exists - creation skipped."
        } else {
            try {

                Write-Verbose -Message "Creating SharePoint security group '$Name'."
                $newGroup = New-Object Microsoft.SharePoint.Client.GroupCreationInformation
                $newGroup.Title = $Name
                $newGroup.Description = $Description
            
                $group = $Web.SiteGroups.Add($newGroup)
                
                $ctx.Load($group)
                $ctx.Load($Web.SiteGroups)
                $ctx.ExecuteQuery()

                Write-Verbose -Message "'$($group.Title)' group was created."
                
                if ($Owner) {
                    Write-Verbose -Message "Looking for owner group '$Owner' for '$Name'."
                    $ownerGroup = $site.RootWeb.SiteGroups | Where-Object { $_.Title -eq $Owner } | Select-Object -First 1
                    if ($ownerGroup) {
                        $group.Owner = $ownerGroup
                        $group.Update()
                        Write-Verbose -Message "'$Owner' group set as the owner of '$Name' group."
                    } else {
                        Write-Verbose -Message "Owner group '$Owner' not found - skipping."
                    }
                             
                } else {
                    Write-Verbose -Message "Owner group name not set - skipping."
                }

                # If group description contains HTML formatting, it must be set the following way. Otherwise HTML tags are automatically escaped.
                # See http://sharepoint.stackexchange.com/questions/26228/html-in-spgroup-description
                $groupInfo = $Web.SiteUserInfoList.GetItemById($group.Id)
                $groupInfo["Notes"] = $Description
                $groupInfo.Update()

                $ctx.Load($group)
                $ctx.ExecuteQuery()

                Write-Host -Object "SharePoint security group '$Name' successfully created."
            } catch {
                Write-Error -Message "Creating SharePoint security group '$Name' failed."
                Write-Error $_
            }
        }

        return $group
	}
	
	end {
        Write-Debug -Message "Ensure-SPOGroup end"
	}
}