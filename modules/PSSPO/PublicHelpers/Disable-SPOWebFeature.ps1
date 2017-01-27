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

function Disable-SPOWebFeature {
    <#

    .SYNOPSIS

    Deactivates SharePoint site feature.



    .DESCRIPTION

    The deactivation is performed only if the feature is activated on site.

    Feature identifiers are expected as pipeline, or may be passed to the -FeatureId parameter.



    .PARAMETER Web 

    The site to deactivate feature on.



    .PARAMETER FeatureId

    Feature definition identifier (string that represents a correct GUID in xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx format).



    .EXAMPLE

    Dectivate 'SharePoint Server Publishing' feature on a site.


    Disable-SPOWebFeature -Web $context.Web -featureId "94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb"



    .EXAMPLE 

    Read site features to activate from xml file and deactivate them on a site


    (Get-Content -Path "c:\portalMap.xml).site.web.features | Select-Object -Property ID, @{ Name="Custom"; Expression = { [System.Convert]::ToBoolean($_.Custom) }} | Disable-SPOFeature -Web $context.Web



    .NOTES

    You need to pass 'Web' argument that is loaded in the context of a user who has privileges to deactivate features.

    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [Microsoft.SharePoint.Client.Web]$Web,

        [Parameter(Mandatory=$true, Position=2, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [Alias("ID")]
        [string]$FeatureId
    )

    begin {
        Write-Debug -Message "Disable-SPOWebFeature begin"
        $ctx = $Web.Context
        $ctx.Load($Web)
        $ctx.Load($Web.Features)
        $ctx.ExecuteQuery()
    }

    process {
        Write-Debug -Message "Disable-SPOWebFeature process: $FeatureId"

        $id = [GUID]$FeatureId
        $feature = ($Web.Features | Where-Object { $_.DefinitionId -eq $id } | Select-Object -first 1)

        if ($feature) {
            try {
                
                Write-Verbose -Message "Feature '$($feature.DefinitionId)' is activated on site $($Web.Url) - deactivating."
				
                $Web.Features.Remove($id, $true) | Out-Null
                $Web.Update()
                $ctx.ExecuteQuery()
                Write-Host -Object "Site feature '$FeatureId' was deactivated on $($Web.Url)."
            } catch {
                Write-Error -Message "Feature '$FeatureId' deactivation failed on $($Web.Url)."
                Write-Error $_
            } 
        } else {
            Write-Verbose -Message "Feature '$FeatureId' is not activated on site $($Web.Url) - deactivation skipped."
        } 
    }

    end {
        Write-Debug -Message "Disable-SPOWebFeature end"
    }
}