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

function Get-SPOContext {
    <#

    .SYNOPSIS

    Gets SharePoint Online client context.



    .DESCRIPTION

    Returns client context for given user and url address.



    .PARAMETER Url

    The URL address to create context for.




    .PARAMETER User

    Name of the user to create context for.



    .PARAMETER Password

    Password for the user specified by 'User' parameter. If not set, a prompt will be displayed.



    .PARAMETER SecurePassword

    Secure string password. This parameter can be set instead of plain text password.



    .EXAMPLE

    $ctx = Get-SPOContext -Url "https://contoso.sharepoin.com" -User "admin@contoso.com" -Password "P@ssw0rd"



    .EXAMPLE

    $ctx = Get-SPOContext -Url "https://contoso.sharepoin.com" -User "admin@contoso.com" -SecurePassword $securePassword
    


    .EXAMPLE

    $ctx = Get-SPOContext -Url "https://contoso.sharepoin.com" -User "admin@contoso.com"



    .NOTES

    The password is passed to the cmdlet in plain text.

    #>
    [CmdletBinding(DefaultParameterSetName="Secure")]
    [OutputType([Microsoft.SharePoint.Client.ClientContext])]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]$Url,

        [Parameter(Mandatory=$true, Position=2)]
        [string]$User,

        [Parameter(Mandatory = $true, ParameterSetName = "Secure", Position=3)]
        [Security.SecureString]$SecurePassword,

        [Parameter(Mandatory = $true, ParameterSetName = "Plain", Position=3)]
        [string]$Password
    )

    if (-not $SecurePassword) {
        $SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
    }

    try {
        Write-Verbose -Message "Creating client context for user '$User' and URL '$Url'."
        $ctx = New-Object Microsoft.SharePoint.Client.ClientContext -ArgumentList @($Url)
        $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials -ArgumentList @($User, $SecurePassword)
        Write-Verbose -Message "Client context successfully created."

    } catch {
        Write-Error -Message "Context creation failed"
        Write-Error $_.Exception.Message
    }
    
    return $ctx
}