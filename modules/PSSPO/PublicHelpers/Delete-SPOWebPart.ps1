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

function Delete-SPOWebPart {
    <#

    .SYNOPSIS

    Deletes SharePoint WebPart on a publishing page.



    .DESCRIPTION

    The deletion is performed only if the WebPart is found on the page. The page is checked out and checked in after the WebPArt is deleted.

    If the page is already checked out, the existing checkout is overriden.

    WebPart names are expected as pipeline, or may be passed to the -WebPartTitle parameter.



    .PARAMETER Web 

    The site that contains publishing page to delete WebPart on.



    .PARAMETER WebPartTitle

    Name (title) of the WebPart to delete.



    .PARAMETER PageName

    Name (library relative url) of the page to delete WebPart on.



    .EXAMPLE

    Delete 'Yammer' WebPart on home page.


    Delete-SPOWebPart -Web $context.Web -WebPartTitle "Yammer" -PageName "default.aspx"



    .NOTES

    You need to pass 'Web' argument that is loaded in the context of a user who has privileges to delete WebParts.

    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [Microsoft.SharePoint.Client.Web]$Web,

        [Parameter(Mandatory=$true, Position=2, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [Alias("Title")]
        [string]$WebPartTitle,

        [Parameter(Mandatory=$true, Position=3, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [Alias("Page")]
        [string]$PageName
    )

    begin {
        Write-Debug -Message "Delete-SPOWebPart begin"
        $ctx = $Web.Context
        $ctx.Load($Web)
        $ctx.Load($Web.Lists)
        $ctx.ExecuteQuery()

        $list = $Web.Lists | Where-Object { $_.Title -eq "Pages" } | Select-Object -First 1

        if ($list -ne $null) {
            $ctx.Load($list)
            $ctx.ExecuteQuery()

            $pages = $list.RootFolder.Files
            $ctx.Load($pages)
            $ctx.ExecuteQuery()

            $page = $pages | Where-Object { $_.Name -eq $PageName } | Select-Object -first 1

            if ($page -ne $null) {
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
        } else {
            Write-Warning -Message "'Pages' library not found at site $($Web.Url)"
            $page = $null
        }
    }

    process {
        Write-Debug -Message "Delete-SPOWebPart process: $WebPartTitle"

        if ($page) {
            try {
                
                Write-Verbose -Message "Page '$($page.Name)' was found on site $($Web.Url) - deleting WebPart."
				
                $wpm = $page.GetLimitedWebPartManager("Shared")
                $ctx.Load($wpm)
                $ctx.ExecuteQuery()

                $webParts = $wpm.WebParts
                $ctx.Load($webParts)
                $ctx.ExecuteQuery()

                $webParts | ForEach-Object {
                    $ctx.Load($_.WebPart)
                    $ctx.ExecuteQuery()
                }

                $webPart = $webParts | Where-Object { $_.WebPart.Title -eq $WebPartTitle } | Select-Object -first 1
                if ($webPart -ne $null) {
                    $webPart.DeleteWebPart()
                    $ctx.ExecuteQuery()
                    Write-Host -Object "WebPart '$WebPartTitle' was deleted on $($Web.Url)/Pages/$PageName."
                }
                else {
                    Write-Host -Object "WebPart '$WebPartTitle' was not found on $($Web.Url)/Pages/$PageName."
                }
                
            } catch {
                Write-Error -Message "WebPart '$WebPartTitle' deleting failed on $($Web.Url)."
                Write-Error $_
            } 
        } elseif ($list -ne $null) {
            Write-Verbose -Message "Page '$PageName' was not found on site $($Web.Url) - deleting WebPart skipped."
        } 
    }

    end {
        if ($page -ne $null) {
            try {
                Write-Host -Object "Checking in... " -NoNewline
                $page.CheckIn("Checked in by PS script (Delete WebPart)", [Microsoft.SharePoint.Client.CheckinType]::MajorCheckIn)
                $ctx.ExecuteQuery()
                Write-Host -Object "DONE"
            } catch {
                Write-Host -Object "FAILED"
            }
        }

        Write-Debug -Message "Delete-SPOWebPart end"
    }
}