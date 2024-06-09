<#
.SYNOPSIS
Converts Microsoft Edge bookmarks from JSON to HTML format for import into other browsers.

.DESCRIPTION
This PowerShell script reads the bookmarks from a Microsoft Edge JSON file and converts them into an HTML format compatible with most web browsers' bookmark import functionality.

.PARAMETER bookmarkPath
The path to the Microsoft Edge bookmarks JSON file. If not provided and promptBookmarkPath is false, the script will use the default path for the current user's Edge bookmarks.

.PARAMETER promptBookmarkPath
If set to true, prompt for a bookmarks JSON file via a system window dialog.

.PARAMETER outputPath
Location where the favorites will be exported to.

.PARAMETER promptOutputPath
If set to true, prompt for save location via a system window dialog.

.EXAMPLE
Convert-EdgeBookmarksToHtml

Converts the bookmarks from the default Edge bookmarks file for the current user.

.EXAMPLE
Convert-EdgeBookmarksToHtml -bookmarkPath "C:\path\to\custom\bookmarks.json"

Converts the bookmarks from a custom JSON file specified by the user.

.EXAMPLE
Convert-EdgeBookmarksToHtml -promptBookmarkPath $true

Prompts the user to select the bookmarks JSON file via a system window dialog.

.NOTES
Author:  Erik Plachta
Created: 20240609
Version: 0.0.1

#>
function Convert-EdgeBookmarksToHtml {
    param (
        [string]$bookmarkPath = "$HOME\AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks",
        [bool]$promptBookmarkPath = $false,
        [string]$outputPath = "$HOME\Desktop\bookmarks_$(((get-date).ToUniversalTime()).ToString("yyyyMMddTHHmmssZ")).html",
        [bool]$promptOutputPath = $false
    )

    # If requested, prompt for the path, opens UI allowing the user to manually select a file.
    if ($promptBookmarkPath -or $promptOutputPath) {
        # Make sure forms are imported for saving and selecting file UI.
        Add-Type -AssemblyName System.Windows.Forms

        # Prompt for bookmark file
        if ($promptBookmarkPath) {
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
            $openFileDialog.ShowDialog() | Out-Null

            # If selected use it, otherwise default.
            if ($openFileDialog.FileName) {
                $bookmarkPath = $openFileDialog.FileName
            }
        }

        # Prompt for output path
        if ($promptOutputPath) {
            $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveFileDialog.Filter = "HTML File (*.html)|*.html"
            $saveFileDialog.ShowDialog() | Out-Null

            # If selected use it, otherwise default.
            if ($saveFileDialog.FileName) {
                $outputPath = $saveFileDialog.FileName
            }
        }
    }

    # Helper function to handle folders within bookmarks.
    function Process-Bookmark-Folders($folder) {
        $html = "<DT><H3>$($folder.name)</H3>`n<DL><p>`n"
        foreach ($child in $folder.children) {
            if ($child.type -eq 'url') {
                $html += "<DT><A HREF=`"$($child.url)`" ADD_DATE=`"$($child.date_added)`">$($child.name)</A>`n"
            } elseif ($child.type -eq 'folder') {
                $html += Process-Bookmark-Folders $child
            }
        }
        $html += "</DL><p>`n"
        return $html
    }

    # Get Edge custom JSON Bookmarks file.
    try {
        $jsonContent = Get-Content -Path $bookmarkPath -Raw | ConvertFrom-Json
    } catch {
        Write-Host "Failed to read or parse the bookmarks JSON file. Please check the path and try again."
        return
    }

    # Starting HTML file content
    $htmlOutput = @"
<!DOCTYPE NETSCAPE-Bookmark-file-1>
<!-- This is an automatically generated file. It will be read and overwritten. DO NOT EDIT! -->
<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=UTF-8'>
<TITLE>Bookmarks</TITLE>
<H1>Bookmarks</H1>
<DL><p>
"@

    foreach ($child in $jsonContent.roots.bookmark_bar.children) {
        if ($child.type -eq 'url') {
            $htmlOutput += "<DT><A HREF=`"$($child.url)`" ADD_DATE=`"$($child.date_added)`">$($child.name)</A>`n"
        } elseif ($child.type -eq 'folder') {
            $htmlOutput += Process-Bookmark-Folders $child
        }
    }

    $htmlOutput += "</DL><p>"

    # Writing to HTML file
    try {
        $htmlOutput | Out-File -FilePath $outputPath -Encoding UTF8
        Write-Host "Bookmarks saved to $outputPath"
    } catch {
        Write-Host "Failed to save the HTML file. Please check the output path and try again."
    }
}

Convert-EdgeBookmarksToHtml
