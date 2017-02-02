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

function Add-SPOGroup {
    <#

    .SYNOPSIS

    Adds SharePoint security group.



    .DESCRIPTION

    The Add-SPOGroup function checks if SharePoint security group of given name exists in site collection groups. If not, the group is created.

    Parameters necessary to create site are expected as pipeline properties of an object, or may be passed to the -Name, -Description and -Owner parameters.



    .PARAMETER Site 

    Site collection to add group to.



    .PARAMETER Name

    Name of the SharePoint security group.



    .PARAMETER Description

    Description of the SharePoint security group.



    .PARAMETER Owner

    Name of the SharePoint security group that will be used as an owner of the created group.



    .EXAMPLE

    Add-SPOGroup -Web $Web -Name "Sample Owners" -Description "Use this group to grant people full control permissions to the SharePoint site" -Owner "Sample Owners"



    .NOTES

    You need to pass 'Web' argument that is loaded in the context of a user who has privileges to create SharePoint security groups.

    Regular users od domain groups cannot be set as group owners yet.

    #>
	[CmdletBinding()]
    [OutputType([Microsoft.SharePoint.Client.Group])]
	param(
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline=$false)]
        [Microsoft.SharePoint.Client.Site]$Site,

        [Parameter(Mandatory=$true, Position=2, ValueFromPipelineByPropertyName=$true)]
        [string]$Name,

        [Parameter(Mandatory=$true, Position=3, ValueFromPipelineByPropertyName=$true)]
        [string]$Description,

        [Parameter(Mandatory=$false, Position=4, ValueFromPipelineByPropertyName=$true)]
        [string]$Owner
    )

	begin {
        Write-Debug -Message "### Add-SPOGroup begin ###"
        Write-Debug -Message "Loading client objects."
        $ctx = $Site.Context
        $ctx.Load($site)
        $web = $Site.RootWeb
        $ctx.Load($web)
        $groups = $web.SiteGroups
        $ctx.Load($groups)
        $ctx.ExecuteQuery()
        Write-Debug -Message "Query execution finished."
	}
	
	process {
        Write-Debug -Message "### Add-SPOGroup process: $Name ###"
        
        $group = $groups | Where-Object { $_.Title -eq $Name } | Select-Object -First 1

        if ($group) {
            Write-Warning -Message "SharePoint security group '$Name' already exists - creation skipped."
            Write-Debug -Message "Retrieving group."
            $ctx.Load($group)
            $ctx.ExecuteQuery()
            Write-Debug -Message "Query execution finished."
        } else {
            try {
                Write-Verbose -Message "Creating SharePoint security group '$Name'."
                Write-Debug -Message "Adding SharePoint group."
                $newGroup = New-Object Microsoft.SharePoint.Client.GroupCreationInformation
                $newGroup.Title = $Name
                $newGroup.Description = $Description
                $group = $groups.Add($newGroup)
                $ctx.Load($group)
                $ctx.Load($groups)
                $ctx.ExecuteQuery()
                Write-Debug -Message "Query execution finished."
                Write-Verbose -Message "'$($group.Title)' group was created."
            } catch {
                Write-Error -Message "Creating SharePoint security group '$Name' failed."
            }

            if ($Owner) {
                Write-Verbose -Message "Updating group owner of '$Name'."
                try {
                    $ownerGroup = $groups | Where-Object { $_.Title -eq $Owner } | Select-Object -First 1
                    if ($ownerGroup) {
                        Write-Debug -Message "Setting group owner."
                        $group.Owner = $ownerGroup
                        $group.Update()
                        $ctx.ExecuteQuery()
                        Write-Debug -Message "Query execution finished."
                        Write-Verbose -Message "'$Owner' group set as the owner of '$Name' group."
                    } else {
                        Write-Warning -Message "Owner group '$Owner' for '$Name' group not found - skipping."
                    }
                } catch {
                    Write-Error -Message "Setting owner for group '$Name' failed"
                }

            } else {
                Write-Verbose -Message "Owner group name not set - skipping."
            }

            try {
                # If group description contains HTML formatting, it must be set the following way. Otherwise HTML tags are automatically escaped.
                # See http://sharepoint.stackexchange.com/questions/26228/html-in-spgroup-description
                Write-Verbose -Message "Updating group description."
                Write-Debug -Message "Setting group description."
                $groupInfo = $web.SiteUserInfoList.GetItemById($group.Id)
                $groupInfo["Notes"] = $Description
                $groupInfo.Update()
                $ctx.Load($group)
                $ctx.ExecuteQuery()
                Write-Debug -Message "Query execution finished."
                Write-Verbose -Message "Description of SharePoint security group '$Name' successfully updated."
            } catch {
                Write-Error -Message "Updating SharePoint security group description failed."
            }
        }

        return $group
	}
	
	end {
        Write-Debug -Message "### Add-SPOGroup end ###"
	}
}