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

function Set-SPOPermissions {
    <#

    .SYNOPSIS

    Set SharePoint permissions on SharePoint object.



    .DESCRIPTION

    The Add-SPOPermissions cmdled grants SharePoint principal (seccurity group or user) permissions to SharePoint securable object.  The permissions are described by permission level.
    
    Parameters necessary to grant permissions are passed to the -Principal, -Target and -PermissionLevel parameters. Principal and Target may be passed as a pipeline parameter.

    If target object inherits the permissions from its parent, the inheritance is broken. Parent permissions can be copied or cleared, depending on user's choice.
    


    .PARAMETER Target

    Target object the principal will be granted permissions to.



    .PARAMETER Principal 

    SharePoint principal (security group or user) the permissions will be granted to.



    .PARAMETER PermissionLevel

    Permission level name.



    .PARAMETER CopyPermissions

    If role inheritance needs to be broken, this parameter determines if parent's permissions should be copied or cleared.


    
    .EXAMPLE

    Add-SPOPermissions -Principal $siteOwners -Target $web -PermissionLevel "Full Control"



    .EXAMPLE

    $siteMembers | Add-SPOPermissions -Target $web -PermissionLevel "Contribute"



    .EXAMPLE

    $list | Add-SPOPermissions -Principal 'admin@contoso.com' -PermissionLevel 'Read'



    .NOTES

    You need to pass 'Target' argument that is loaded in the context of a user who has privileges to grant SharePoint permissions.
    
    Permissions of all child objects that do not inherit permissions from target object will remain unchanged.

    #>	
	[CmdletBinding()]
    [OutputType([Microsoft.SharePoint.Client.Principal])]
	param(
        [Parameter(Mandatory=$true, Position=1)]
        [Microsoft.SharePoint.Client.SecurableObject]$Target,

        [Parameter(Mandatory=$true, Position=2, ValueFromPipeline=$true)]
        [Microsoft.SharePoint.Client.Principal]$Principal,

        [Parameter(Mandatory=$true, Position=3)]
        [string]$PermissionLevel,

        [Parameter(Mandatory=$false, Position=4)]
        [switch]$CopyPermissions
    )

	begin {
        Write-Debug -Message "### Set-SPOPermissions begin ###"
        Write-Debug -Message "Loading client objects."
        $ctx = $Target.Context
        $ctx.Load($Principal)
        $ctx.Load($Target)
        Invoke-LoadMethod -Object $Target -PropertyName "HasUniqueRoleAssignments"
        $ctx.ExecuteQuery()
        Write-Debug -Message "Query execution finished."
	}
	
	process {
        try {
            Write-Debug -Message "### Set-SPOPermissions process ###"

            if ($Target.HasUniqueRoleAssignments -eq $false) {
                Write-Warning -Message "Target object inherits permissions from its parent. The inheritance will be broken."
                Write-Debug "Breaking role inheritance."
                # Breaking role inheritance will not affect child objects that do not interit permissions from target object (clearSubscopes = $false)
                # See more: https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.securableobject.breakroleinheritance.aspx
                $Target.BreakRoleInheritance($CopyPermissions, $false)
                $ctx.Load($Target)
                $ctx.ExecuteQuery()
                Write-Debug -Message "Query execution finished."
            }
   
            $permissions = $ctx.Web.RoleDefinitions.GetByName($PermissionLevel)
            $roleDefinitionBinding = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
            $roleDefinitionBinding.Add($permissions)
            $assignments = $Target.RoleAssignments
            $assignment = $assignments.Add($Principal, $roleDefinitionBinding)
            $ctx.ExecuteQuery()
            Write-Debug -Message "Query execution finished."
            Write-Verbose -Message "'$PermissionLevel' permissions set up for $($Principal.LoginName)."
        } catch {
            Write-Error -Message "Setting permissions on target object failed."
        }

        return $Principal
	}
	
	end {
        Write-Debug -Message "### Set-SPOPermissions end ###"
	}
}