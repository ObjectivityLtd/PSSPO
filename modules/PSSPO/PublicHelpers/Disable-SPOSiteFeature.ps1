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

    The deactivation is performed only if the feature is activated on site collection.

    Feature identifiers are expected as pipeline, or may be passed to the -FeatureId parameter.



    .PARAMETER Site 

    The site collection to deactivate feature on.



    .PARAMETER FeatureId

    Feature definition identifier (string that represents a correct GUID in xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx format).



    .EXAMPLE

    Deactivate 'SharePoint Server Publishing Infrastructure' feature on a site collection.


    Disable-SPOSiteFeature -Site $context.Site -FeatureId "f6924d36-2fa8-4f0b-b16d-06b7250180fa"



    .EXAMPLE 

    Read site features to activate from xml file and activate them on a site


    (Get-Content -Path "c:\portalMap.xml).site.features | Disable-SPOSiteFeature -Site $context.Site



    .NOTES

    You need to pass 'Site' argument that is loaded in the context of a user who has privileges to activate features.

    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [Microsoft.SharePoint.Client.Site]$Site,

        [Parameter(Mandatory=$true, Position=2, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [Alias("ID")]
        [string]$FeatureId
    )

    begin {
        Write-Debug -Message "Disable-SPOSiteFeature begin"
        $ctx = $Site.Context
        $ctx.Load($Site)
        $ctx.Load($Site.Features)
        $ctx.ExecuteQuery()
    }

    process {
        Write-Debug -Message "Ensure-SPOSiteFeature process: $FeatureId"

        $id = [GUID]$FeatureId
        $feature = $Site.Features | Where-Object { $_.DefinitionId -eq $id } | Select-Object -first 1

        if ($feature) {
            try {
                Write-Verbose -Message "Feature '$($feature.DefinitionId)' is already activated on site collection $($Site.Url) - deactivating."
                $Site.Features.Remove($id, $true) | Out-Null
                $Site.RootWeb.Update()
                $ctx.ExecuteQuery()
                Write-Host -Object "Site collection feature '$FeatureId' was deactivated on $($Site.Url)."
            } catch {
                Write-Error -Message "Feature '$FeatureId' deactivation failed on $($Site.Url)."
                Write-Error $_
            }
        } else {
            Write-Verbose -Message "Feature '$FeatureId' is not activated on site collection $($Site.Url) - deactivation skipped."    
        } 
    }

    end {
        Write-Debug -Message "Disable-SPOSiteFeature end"
    }
}