<#
.SYNOPSIS
Converts Microsoft Edge bookmarks from JSON to HTML format for import into other browsers.

.DESCRIPTION
This PowerShell script reads the bookmarks from a Microsoft Edge JSON file and converts them into an HTML format compatible with most web browsers' bookmark import functionality.

.PARAMETER JsonFilePath
The path to the Microsoft Edge bookmarks JSON file. If not provided and PromptForPath is false, the script will use the default path for the current user's Edge bookmarks.

.PARAMETER PromptForPath
If set to true, the script will prompt the user with a file selection dialog to choose the bookmarks JSON file.

.EXAMPLE
Convert-EdgeBookmarksToHtml

Converts the bookmarks from the default Edge bookmarks file for the current user.

.EXAMPLE
Convert-EdgeBookmarksToHtml -JsonFilePath "C:\path\to\custom\bookmarks.json"

Converts the bookmarks from a custom JSON file specified by the user.

.EXAMPLE
Convert-EdgeBookmarksToHtml -PromptForPath $true

.NOTES
Author:  Erik Plachta
Created: 20240609
Version: 0.0.1

#>
function Convert-EdgeBookmarksToHtml() {
    param (
        [string]$JsonFilePath = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Bookmarks",
        [bool]$PromptForPath = $false
    )

    # Make sure forms imported for saving and selecting file UI.
    Add-Type -AssemblyName System.Windows.Forms

    # If requested prompt for path, opens UI allowing user to manually select file.
    if ($PromptForPath)
    {
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
        $openFileDialog.ShowDialog() | Out-Null
        $JsonFilePath = $openFileDialog.FileName
    }

    function Process-Folder($folder) {
        $html = "<DT><H3>$($folder.name)</H3>`n<DL><p>`n"
        foreach ($child in $folder.children) {
            if ($child.type -eq 'url') {
                $html += "<DT><A HREF=`"$($child.url)`" ADD_DATE=`"$($child.date_added)`">$($child.name)</A>`n"
            } elseif ($child.type -eq 'folder') {
                $html += Process-Folder $child
            }
        }
        $html += "</DL><p>`n"
        return $html
    }

    # Reading JSON file
    $jsonContent = Get-Content -Path $JsonFilePath -Raw | ConvertFrom-Json

    # Starting HTML file content
    $htmlOutput = "<!DOCTYPE NETSCAPE-Bookmark-file-1>
        <!-- This is an automatically generated file. It will be read and overwritten. DO NOT EDIT! -->
        <META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=UTF-8'>
        <TITLE>Bookmarks</TITLE>
        <H1>Bookmarks</H1>
        <DL>
        <p>"
    foreach ($child in $jsonContent.roots.bookmark_bar.children) {
        if ($child.type -eq 'url') {
            $htmlOutput += "<DT><A HREF=`"$($child.url)`" ADD_DATE=`"$($child.date_added)`">$($child.name)</A>`n"
        } elseif ($child.type -eq 'folder') {
            $htmlOutput += Process-Folder $child
        }
    }

    $htmlOutput += "</DL><p>"

    # Prompt user for output file path
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.filter = "HTML File (*.html)|*.html"
    $saveFileDialog.ShowDialog() | Out-Null
    $HtmlOutputPath = $saveFileDialog.FileName

    if ($HtmlOutputPath -ne "") {
        # Writing to HTML file
        $htmlOutput | Out-File -FilePath $HtmlOutputPath -Encoding UTF8
        Write-Host "Bookmarks saved to $HtmlOutputPath"
    } else {
        Write-Host "No file selected. Operation cancelled."
    }
}

Convert-EdgeBookmarksToHtml
