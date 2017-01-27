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

function Add-SPOGroupPermissions {
    <#

    .SYNOPSIS

    Grant SharePoint security group permissions to site.



    .DESCRIPTION

    The Add-SPOGroupPermissions cmdled grants SharePoint security group site permissions. The permissions are described by permission level.

    Parameters necessary to grant permissions are passed to the -Group,  -TargetWebUrl and -PermissionLevel parameters. Group may be passed

    as a pipeline parameter.



    .PARAMETER Group 

    SharePoint security group permissions will be granted to.



    .PARAMETER TargetWeb

    Target site the group will be granted permissions to.



    .PARAMETER PermissionLevel

    Permission level name.



    .EXAMPLE

    Grant site owners 'Full Control' permissions to the site.


    Add-SPOGroupPermissions -Group $siteOwners -TargetWeb $web -PermissionLevel "Full Control" -Force



    .EXAMPLE

    Grant site members 'Edit' permissions to the site using pipeline to pass 'Group' parameter to the cmdlet.


    $siteMembers | Add-SPOGroupPermissions -TargetWeb $web -PermissionLevel "Edit" -Force



    .NOTES

    You need to pass 'Group' argument that is loaded in the context of a user who has privileges to grant SharePoint security groups permissions.

    #>	
	[CmdletBinding()]
    [OutputType([Microsoft.SharePoint.Client.Group])]
	param(
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [Microsoft.SharePoint.Client.Group]$Group,

        [Parameter(Mandatory=$true, Position=2)]
        [Microsoft.SharePoint.Client.Web]$TargetWeb,

        [Parameter(Mandatory=$true, Position=3)]
        [string]$PermissionLevel
    )

	begin {
        Write-Debug -Message "Ensure-SPOGroupPermissions begin"
        
        $originalCtx = $TargetWeb.context
        $originalCtx.Load($TargetWeb)
        $originalCtx.ExecuteQuery()
	}
	
	process {
        try {
            Write-Debug -Message "Ensure-SPOGroupPermissions process: $($Group.Title)"
        
            Write-Debug -Message "Creating new client context."
            $ctx = New-Object Microsoft.SharePoint.Client.ClientContext -ArgumentList @($TargetWeb.Url)
            $ctx.Credentials = $Group.Context.Credentials
            Write-Debug "Client contxt created for URL address '$($TargetWeb.Url)'."

            $web = $ctx.Web
        
            $ctx.Load($web)
            $ctx.Load($web.RoleAssignments)
            $ctx.Load($web.RoleDefinitions)
            $ctx.Load($web.SiteGroups)
            $ctx.ExecuteQuery()

            $gr = $web.SiteGroups.GetByName($Group.Title)
            $ctx.Load($gr)
   
            $permissions = $web.RoleDefinitions.GetByName($PermissionLevel)

            $roleDefinitionBinding = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
            $roleDefinitionBinding.Add($permissions)

            $assignments = $web.RoleAssignments
            $assignment = $assignments.Add($gr, $roleDefinitionBinding)

            $ctx.ExecuteQuery()

            Write-Host -Object "'$PermissionLevel' granted to site $($web.Url) for '$($gr.Title)'"
        } catch {
            Write-Error -Message "Granting '$($gr.Title)' '$PermissionLevel' permissions failed."
            Write-Error $_
        } finally {
            if ($ctx) {
                $ctx.Dispose()
            }
        }
        return $Group
	}
	
	end {
        Write-Debug -Message "Ensure-SPOGroupPermissions end"
	}
}