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

function Invoke-LoadMethod() {
    <#

    .SYNOPSIS

    Load property of a SharePoint object.



    .DESCRIPTION

    Gets url of the container site for sites of given type (guilds or tribes).



    .PARAMETER PropertyName

    Path to the configuration xml file the setting is stored in.



    .Example

    Load property "RequestAccessEmail" for a web

    Invoke-LoadMethod -Object $Web -PropertyName "RequestAccessEmail"



    .NOTES

    Original function was posted on http://sharepoint.stackexchange.com/questions/126221/spo-retrieve-hasuniqueroleassignements-property-using-powershell.

    It was copied and modified a bit.

    #>
    [CmdletBinding()]
    param(
       [Parameter(Mandatory=$true, Position = 1)]
       [Microsoft.SharePoint.Client.ClientObject]$Object,
       [Parameter(Mandatory=$true, Position = 2)]
       [string]$PropertyName
    ) 

    $ctx = $Object.Context
    $load = [Microsoft.SharePoint.Client.ClientContext].GetMethod("Load") 
    $type = $Object.GetType()
    $clientLoad = $load.MakeGenericMethod($type) 


    $Parameter = [System.Linq.Expressions.Expression]::Parameter(($type), $type.Name)
    $Expression = [System.Linq.Expressions.Expression]::Lambda(
                    [System.Linq.Expressions.Expression]::Convert(
                        [System.Linq.Expressions.Expression]::PropertyOrField($Parameter,$PropertyName),
                        [System.Object]
                    ),
                    $($Parameter)
                  )

    $ExpressionArray = [System.Array]::CreateInstance($Expression.GetType(), 1)
    $ExpressionArray.SetValue($Expression, 0)
    $clientLoad.Invoke($ctx,@($Object,$ExpressionArray))
}