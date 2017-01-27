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

    Ensures that SharePoint site feature is activated.



    .DESCRIPTION

    The activation is performed only if the feature is not activated on site.

    Feature identifiers are expected as pipeline, or may be passed to the -FeatureId parameter.



    .PARAMETER Web 

    The site to activate feature on.



    .PARAMETER FeatureId

    Feature definition identifier (string that represents a correct GUID in xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx format).



    .PARAMETER Custom

    Switch parameter determines if activated feature is out-of-the-box SharePoint feature (false) or custom one (true)



    .EXAMPLE

    Activate 'SharePoint Server Publishing' feature on a site.


    Ensure-SPOWebFeature -Web $context.Web -featureId "94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb"



    .EXAMPLE 

    Read site features to activate from xml file and activate them on a site


    (Get-Content -Path "c:\portalMap.xml).site.web.features | Select-Object -Property ID, @{ Name="Custom"; Expression = { [System.Convert]::ToBoolean($_.Custom) }} | Ensure-SPOFeature -Web $context.Web



    .NOTES

    You need to pass 'Web' argument that is loaded in the context of a user who has privileges to activate features.

    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [Microsoft.SharePoint.Client.Web]$Web,

        [Parameter(Mandatory=$true, Position=2, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [Alias("ID")]
        [string]$FeatureId,
		
		[Parameter(Mandatory=$false, Position=3, ValueFromPipelineByPropertyName=$true)]
        [switch]$Custom
    )

    begin {
        Write-Debug -Message "Ensure-SPOWebFeature begin"
        $ctx = $Web.Context
        $ctx.Load($Web)
        $ctx.Load($Web.Features)
        $ctx.ExecuteQuery()
    }

    process {
        Write-Debug -Message "Ensure-SPOWebFeature process: $FeatureId"

        $id = [GUID]$FeatureId
        $feature = ($Web.Features | Where-Object { $_.DefinitionId -eq $id } | Select-Object -first 1)

        if ($feature) {
            Write-Verbose -Message "Feature '$($feature.DefinitionId)' is already activated on site $($Web.Url) - activation skipped."
        } else {
            try {
                Write-Verbose -Message "Feature '$FeatureId' is not activated on site $($Web.Url) - activating."
                
                $scope = Get-FeatureScope -Custom:$Custom
				
                $Web.Features.Add($id, $false, $scope) | Out-Null
                $Web.Update()
                $ctx.ExecuteQuery()
                Write-Host -Object "Site feature '$FeatureId' was activated on $($Web.Url)."
            } catch {
                Write-Error -Message "Feature '$FeatureId' activation failed on $($Web.Url)."
                Write-Error $_
            }
        } 
    }

    end {
        Write-Debug -Message "Ensure-SPOWebFeature end"
    }
}