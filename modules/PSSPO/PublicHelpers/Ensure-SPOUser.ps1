﻿<#
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

function Ensure-SPOUser {
    <#

    .SYNOPSIS

    Ensures that SharePoint Online user exists on target site collection.



    .DESCRIPTION

    Adds SharePoint Online user to site users list if the user entry not exists yet.

    User logins are expected as pipeline, or may be passed to the -LoginName parameter.


    .PARAMETER Site

    The site collection in context of which the user is ensured.



    .PARAMETER LoginName

    Login name of the user to ensure.



    .Example

    Ensure-SPOUser -Site $ctx.Site -LoginName "jdoe@objectivity.co.uk"



    .NOTES

    You need to pass 'Site' argument that is loaded in the context of a user who has privileges to browse user info.

    #>
    [CmdletBinding()]
    [OutputType([Microsoft.SharePoint.Client.User])]

    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$false, Position=1)]
        [Microsoft.SharePoint.Client.Site]$Site,

        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=2)]
        [Alias("User", "UserName", "Name")]
        [string]$LoginName
    )

    begin {
        Write-Debug -Message "Ensure-SPOUser begin"
        $ctx = $Site.Context
        $web = $ctx.Web
        $ctx.Load($web)
        $ctx.ExecuteQuery()
    }

    process {
        
        try {
            Write-Debug -Message "Ensure-SPOUser process"
        
            $user = $web.EnsureUser($LoginName)
            $ctx.Load($user)
            $ctx.ExecuteQuery()
            Write-Host -Object "User '$LoginName' successfully ensured on $($web.Url)"

            return $user

        } catch {
            Write-Error -Message "Ensuring user '$LoginName' failed."
            Write-Error $_
            return $null
        }
    }

    end {
        Write-Debug -Message "Ensure-SPOUser end"
    }
}