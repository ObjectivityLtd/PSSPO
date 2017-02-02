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

function Disable-SPOSiteFeature {
    <#

    .SYNOPSIS

    Deactivates SharePoint site collection feature.



    .DESCRIPTION

    The deactivation is performed only if the feature is active on site collection.

    Feature identifiers are expected as pipeline, or may be passed to the -FeatureId parameter.



    .PARAMETER Site 

    The site collection to deactivate feature on.



    .PARAMETER FeatureId

    Feature definition identifier.



    .EXAMPLE

    Disable-SPOSiteFeature -Site $context.Site -FeatureId "f6924d36-2fa8-4f0b-b16d-06b7250180fa"



    .EXAMPLE 

    $identifiers | Disable-SPOSiteFeature -Site $site



    .NOTES

    You need to pass 'Site' argument that is loaded in the context of a user who has privileges to deactivate features.

    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [Microsoft.SharePoint.Client.Site]$Site,

        [Parameter(Mandatory=$true, Position=2)]
        [GUID]$Id
    )

    $ctx = $Site.Context
    $ctx.Load($Site)
    Write-Debug -Message "Retrieving site collection features."
    $ctx.Load($Site.Features)
    $ctx.ExecuteQuery()
    Write-Debug -Message "Query execution finished."

    $feature = $Site.Features | Where-Object { $_.DefinitionId.Guid -eq $Id.Guid } | Select-Object -first 1

    if ($feature) {
        Write-Verbose -Message "Feature '$($feature.DefinitionId)' is active on site collection $($Site.Url) - deactivating."
        try {
            Write-Debug -Message "Removing feature from site collection features list."
            $Site.Features.Remove($id, $true)
            $ctx.ExecuteQuery()
            Write-Debug -Message "Query execution finished."
            Write-Verbose -Message "Site collection feature '$FeatureId' was deactivated on $($Site.Url)."
        } catch {
            Write-Error -Message "Feature '$FeatureId' deactivation failed on $($Site.Url)."
            Write-Error $_.Exception.Message
        }
    } else {
        Write-Warning -Message "Feature '$FeatureId' is not active on site collection $($Site.Url) - deactivation skipped."    
    } 
}