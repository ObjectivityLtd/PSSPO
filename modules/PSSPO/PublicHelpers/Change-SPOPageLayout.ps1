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

function Change-SPOPageLayout {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [Microsoft.SharePoint.Client.Web]$Web,

        [Parameter(Mandatory=$true, Position=2, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [Alias("Title")]
        [string]$PageName,

        [Parameter(Mandatory=$true, Position=3, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [Alias("Page")]
        [string]$Layout
    )

    begin {
        Write-Debug -Message "Change-SPOPageLayout begin"
        $ctx = $Web.Context
        $ctx.Load($Web)
        $ctx.Load($Web.Lists)
        $ctx.ExecuteQuery()

        $list = $Web.Lists.GetByTitle("Pages")
        $ctx.Load($list)
        $ctx.ExecuteQuery()

        $camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
        $camlQuery.ViewXml = '<View><Query><Where><Eq><FieldRef Name="FileLeafRef" /><Value Type="Text">'+ $PageName +'</Value></Eq></Where></Query></View>'

        $pageItem = $list.GetItems($camlQuery)
        $ctx.Load($pageItem)
        $ctx.ExecuteQuery()

        

        if ($pageItem -ne $null) {
            
            $page = $pageItem.File

            $ctx.Load($page)
            $ctx.ExecuteQuery()

            if ($page.CheckOutType -eq [Microsoft.SharePoint.Client.CheckOutType]::None) {
                try {
                    Write-Host -Object "Checking out page $($Web.Url)/Pages/$PageName... " -NoNewLine
                    $page.CheckOut()
                    $ctx.ExecuteQuery()
                    Write-Host -Object "DONE"
                } catch {
                    Write-Error -Message "FAILED"
                }
            } else {
                $ctx.Load($page.CheckedOutByUser)
                $ctx.ExecuteQuery()

                $checkoutUserLogin = $page.CheckedOutByUser.Email
                $curentUser = $ctx.Credentials.UserName
                
                if ($checkoutUserLogin -ne $curentUser) {
                    try {
                        Write-Warning -Message "Page $($Web.Url)/Pages/$PageName is checked out for $checkoutUserLogin."
                        Write-Host -Object "Undoing checkout... " -NoNewline
                        $page.UndoCheckOut()
                        $ctx.ExecuteQuery()
                        Write-Host -Object "DONE"
                    
                        Write-Host "Checking out page $($Web.Url)/Pages/$PageName... " -NoNewline
                        $page.CheckOut()
                        $ctx.ExecuteQuery()
                        Write-Host "DONE"
                    } catch {
                        Write-Host "FAILED"
                        Write-Error -Message $_
                    }
                } else {
                    Write-Warning -Message "Page $($Web.Url)/Pages/$PageName is already checked out for current user."
                }
            }

        } else {
            Write-Warning -Message "Page $PageName not found at site $($Web.Url)"
        }
    }

    process {
        Write-Debug -Message "Delete-SPOWebPart process: $WebPartTitle"

        if ($page) {
            try {
                
                Write-Host -Object "Page '$($page.Name)' was found on site $($Web.Url) - changing layout... " -NoNewline
				
                $pageItem.Set_Item("PublishingPageLayout", $Layout);
                $pageItem.Update()
                $ctx.ExecuteQuery()
                Write-HOST "DONE"

                
            } catch {
                Write-Host -Object "FAILED"
                Write-Error $_
            } 
        } else {
            Write-Verbose -Message "Page '$($page.Name)' was not found on site $($Web.Url) - changing layout skipped."
        } 
    }

    end {
        if ($page -ne $null) {
            try {
                Write-Host -Object "Checking in... " -NoNewline
                $page.CheckIn("Checked in by PS script", [Microsoft.SharePoint.Client.CheckinType]::MajorCheckIn)
                $ctx.ExecuteQuery()
                Write-Host -Object "DONE"
            } catch {
                Write-Host -Object "FAILED"
            }
        }

        Write-Debug -Message "Change-SPOPageLayout end"
    }
}