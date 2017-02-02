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

function Enable-SPOWebFeature {
    <#

    .SYNOPSIS

    Activates site feature.



    .DESCRIPTION

    The activation is performed only if the feature is not activated on the site.



    .PARAMETER Web 

    The site to activate feature on.



    .PARAMETER Id

    Feature definition identifier.



    .PARAMETER Custom

    Determines if feature identifier represents a custom or SharePoint built-in feature.



    .EXAMPLE

    Enable-SPOWebFeature -Web $web -Id $featureId



    .NOTES

    You need to pass 'Web' argument that is loaded in the context of a user who has privileges to activate features.

    #>
    [CmdletBinding()]
    [OutputType([Microsoft.SharePoint.Client.Feature])]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [Microsoft.SharePoint.Client.Web]$Web,

        [Parameter(Mandatory=$true, Position=2)]
        [GUID]$Id,

        [Parameter(Mandatory=$false, Position=3)]
        [switch]$Custom
    )

    $ctx = $Web.Context
    $ctx.Load($Web)
    Write-Debug -Message "Retrieving site features."
    $ctx.Load($Web.Features)
    $ctx.ExecuteQuery()
    Write-Debug -Message "Query execution finished."

    $feature = $Web.Features | Where-Object { $_.DefinitionId.Guid -eq $Id.Guid } | Select-Object -first 1

    if ($feature) {
        Write-Warning -Message "Feature '$($feature.DefinitionId)' is already active on site $($Web.Url) - activation skipped."
        $ctx.Load($feature)
        $ctx.ExecuteQuery()
        Write-Debug -Message "Query execution finished."
    } else {
        Write-Verbose -Message "Feature '$FeatureId' is not active on site $($Web.Url) - activating."
        try {
            $scope = Get-FeatureScope -Custom:$Custom
            Write-Debug -Message "Adding feature to features list."
            $feature = $Web.Features.Add($id, $false, $scope)
            $ctx.Load($feature)
            $ctx.ExecuteQuery()
            Write-Debug -Message "Query execution finished."
            Write-Verbose -Message "Site feature '$FeatureId' was activated on $($Web.Url)."
        } catch {
            Write-Error -Message "Feature '$FeatureId' activation failed on $($Web.Url)."
            Write-Error $_.Exception.Message
        }
    }

    return $feature
}