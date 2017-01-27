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

function Get-AssemblyPath {
    <#

    .SYNOPSIS

    Get assembly path.



    .DESCRIPTION

    Gets assembly path.

    Assembly name is passed to the -AssemblyName parameter. Default path that is set to SharePoint installation catalog in Program Files can be overriden

    by passing -DefaultPath paremeter. If not found in default location, assembly is looked for in current script catalog.



    .PARAMETER AssemblyName

    Name of the assembly.



    .PARAMETER DefaultPath

    Default path of the assembly where it is expected to be installed.



    .Example

    Get-AssemblyPath -AssemblyName "Microsoft.SharePoint.Client"



    .NOTES

    If assembly is not found, null value is returned.

    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]$AssemblyName,

        [Parameter(Mandatory=$false, Position=2)]
        [string]$DefaultPath="$env:15\ISAPI"
    )

    $fullPath = Join-Path -Path $DefaultPath -ChildPath $AssemblyName

    if (-not (Test-Path -Path $fullPath)) {
        Write-Debug -Message "$fullPath not found"
        Write-Debug -Message "Looking for '$AssemblyName' in $PSScriptRoot\Binaries"
        $fullPath = Join-Path -Path "$PSScriptRoot\Binaries" -ChildPath $AssemblyName
        if (-not (Test-Path -Path  $fullPath)) {
            Write-Verbose -Message "$fullPath not found"
            $fullPath = $null
        } else {
            Write-Verbose -Message "$AssemblyName found in '$PSScriptRoot\Binaries'"
        }
    } else {
        Write-Verbose -Message "$AssemblyName found in $DefaultPath"
    }
    return $fullPath
}