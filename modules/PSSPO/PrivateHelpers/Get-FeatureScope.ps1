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

function Get-FeatureScope {
    <#

    .SYNOPSIS

    Gets scope of the web feature depending on feature origin (SharePoint OOTB feature or a custom one).



    .NOTES

    This function is intended for internal use only.

    #>
    [CmdletBinding()]
    [OutputType([Microsoft.SharePoint.Client.FeatureDefinitionScope])]
    param(
		[Parameter(Mandatory=$false, Position=1)]
        [switch]$Custom
    )
    
    # FeatureDefinitionScope is misleading in case of web-scoped feature.
    # The documentation says: "It must have the value of FeatureDefinitionScope.Site or FeatureDefinitionScope.Farm" (although FeatureDefinitionScope.Web is available)
    # Using 'Site' value causes an error, but for 'None' it works fine (for SharePoint OOTB features).
    # See: https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.featurecollection.add.aspx
    # See also: http://sadomovalex.blogspot.com/2014/08/reactivate-web-scoped-features-from.html
    # "None" works fine for OOTB features, for custom ones use "Site" instead
	$scope = [Microsoft.SharePoint.Client.FeatureDefinitionScope]::None
				
	if ($Custom) {
		$scope = [Microsoft.SharePoint.Client.FeatureDefinitionScope]::Site
	}

    return $scope
}