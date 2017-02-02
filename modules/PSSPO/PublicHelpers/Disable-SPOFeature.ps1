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

function Disable-SPOFeature {
    <#

    .SYNOPSIS

    Deactivates SharePoint feature.



    .DESCRIPTION

    The deactivation is performed only if the feature is active.

    Feature identifiers are expected as pipeline, or may be passed to the -FeatureId parameter.



    .PARAMETER Site 

    The site collection to deactivate feature on.



    .PARAMETER Web 

    The site to deactivate feature on.



    .PARAMETER FeatureId

    Feature definition identifier (string that represents a correct GUID in xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx format).



    .EXAMPLE

    Disable-SPOFeature -Site $site -featureId "f6924d36-2fa8-4f0b-b16d-06b7250180fa"
    


    .EXAMPLE

    Disable-SPOFeature -Web $web -featureId "94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb"


    .EXAMPLE 

    $identifiers | Disable-SPOFeature -Web $web



    .NOTES

    You need to pass parent element ('Site' or 'Web') argument that is loaded in the context of a user who has privileges to deactivate features.

    #>
    [CmdletBinding(DefaultParameterSetName="Site")]
    [OutputType([void])]
    param(
        [Parameter(Mandatory=$true, ParameterSetName = "SiteCollection", Position=1)]
        [Microsoft.SharePoint.Client.Site]$Site,

        [Parameter(Mandatory=$true, ParameterSetName = "Site", Position=1)]
        [Microsoft.SharePoint.Client.Web]$Web,

        [Parameter(Mandatory=$true, Position=2, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [Alias("ID")]
        [string]$FeatureId
    )

    begin {
        Write-Debug -Message "### Disable-SPOFeature begin ###"
    }

    process {
        Write-Debug -Message "### Disable-SPOFeature process: $FeatureId ###"

        $id = [GUID]$FeatureId
        
        if ($PsCmdlet.ParameterSetName -eq "SiteCollection") {
            Disable-SPOSiteFeature -Site $Site -Id $id
        } elseif ($PsCmdlet.ParameterSetName -eq "Site") {
            Disable-SPOWebFeature -Web $Web -Id $id
        }

        return $feature
    }

    end {
        Write-Debug -Message "### Disable-SPOFeature end ###"
    }
}