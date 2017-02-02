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

function Enable-SPOFeature {
    <#

    .SYNOPSIS

    Enables SharePoint feature.



    .DESCRIPTION

    The activation is performed only if the feature is not active yet.

    Feature identifiers are expected as pipeline, or may be passed to the -FeatureId parameter.



    .PARAMETER Site 

    The site collection to activate feature on.



    .PARAMETER Web

    The site to activate feature on.



    .PARAMETER FeatureId

    Feature definition identifier (string that represents a correct GUID in xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx format).



    .EXAMPLE

    Ensure-SPOFeature -Site $site -FeatureId "f6924d36-2fa8-4f0b-b16d-06b7250180fa"



    .EXAMPLE 

    $identifiers | Enable-SPOFeature -Site $site



    .EXAMPLE

    Ensure-SPOFeature -Web $web -FeatureId "94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb"



    .EXAMPLE 

    $identifiers | Enable-SPOFeature -Web $web



    .NOTES

    You need to pass parent element ('Site' or 'Web') argument that is loaded in the context of a user who has privileges to activate features.

    #>
    [CmdletBinding(DefaultParameterSetName="Site")]
    [OutputType([Microsoft.SharePoint.Client.Feature])]
    param(
        [Parameter(Mandatory=$true, ParameterSetName = "SiteCollection", Position=1)]
        [Microsoft.SharePoint.Client.Site]$Site,

        [Parameter(Mandatory=$true, ParameterSetName = "Site", Position=1)]
        [Microsoft.SharePoint.Client.Web]$Web,

        [Parameter(Mandatory=$true, Position=2, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [Alias("ID")]
        [string]$FeatureId,

        [Parameter(Mandatory=$false, Position=3, ValueFromPipelineByPropertyName=$true)]
        [switch]$Custom
    )

    begin {
        Write-Debug -Message "### Enable-SPOFeature begin ###"
    }

    process {
        Write-Debug -Message "### Enable-SPOFeature process: $FeatureId ###"

        $id = [GUID]$FeatureId
        
        if ($PsCmdlet.ParameterSetName -eq "SiteCollection") {
            $feature = Enable-SPOSiteFeature -Site $Site -Id $id -Custom:$Custom
        } elseif ($PsCmdlet.ParameterSetName -eq "Site") {
            $feature = Enable-SPOWebFeature -Web $Web -Id $id -Custom:$Custom
        }

        return $feature
    }

    end {
        Write-Debug -Message "### Enable-SPOFeature end ###"
    }
}