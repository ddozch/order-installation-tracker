<#
.SYNOPSIS
    Removes the folder path that Excel saves inside a workbook, then checks that no local
    path or computer name is left in the file.

.DESCRIPTION
    On every save Excel writes the folder of the workbook into xl/workbook.xml
    (<x15ac:absPath url="C:\...\"/> inside an mc:AlternateContent block). Excel does not need
    it to open the file, so the block is removed. Run this after Excel has closed the file;
    tools\build_workbook.ps1 runs it at the end of the build.

    The check reads every part of the package, vbaProject.bin included, as UTF-8, Latin-1 and
    UTF-16 text. It stops with an error if a part contains the user profile folder, the folder
    of the workbook, the computer name or the Windows user name as a folder of a path (\name\),
    or if the author fields hold the Windows user name.
    The VBA source inside vbaProject.bin is compressed and is not covered: it is the code in src\.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\scrub_workbook.ps1 InstallationTracker.xlsm
#>
param(
    [string]$Path = (Join-Path (Split-Path $PSScriptRoot -Parent) 'InstallationTracker.xlsm')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$Path = [IO.Path]::GetFullPath($Path)
$utf8 = New-Object Text.UTF8Encoding $false
$latin1 = [Text.Encoding]::GetEncoding(28591)

function ReadBytes($stream) {
    $m = New-Object IO.MemoryStream
    $stream.CopyTo($m)
    return , $m.ToArray()
}

# 1. Remove the absPath block from xl/workbook.xml. The entry keeps its place and time stamp.
$zip = [IO.Compression.ZipFile]::Open($Path, 'Update')
try {
    $entry = $zip.GetEntry('xl/workbook.xml')
    $stamp = $entry.LastWriteTime
    $s = $entry.Open()
    try {
        $xml = $utf8.GetString((ReadBytes $s))
        $clean = [regex]::Replace($xml, '<mc:AlternateContent\b[^>]*>\s*<mc:Choice\b[^>]*>\s*<x15ac:absPath\b[^>]*/>\s*</mc:Choice>\s*</mc:AlternateContent>', '')
        if ($clean -ne $xml) {
            $bytes = $utf8.GetBytes($clean)
            $s.SetLength(0)
            $s.Write($bytes, 0, $bytes.Length)
            Write-Host 'Removed the saved folder path from xl/workbook.xml'
        }
    }
    finally { $s.Dispose() }
    $entry.LastWriteTime = $stamp
}
finally { $zip.Dispose() }

# 2. Check every part
$needles = @(@($env:USERPROFILE, (Split-Path $Path -Parent), $env:COMPUTERNAME) | Where-Object { $_ } | Select-Object -Unique)
# The user name as a folder also catches a path that is cut off before C:\Users
if ($env:USERNAME) { $needles += @("\$env:USERNAME\", "/$env:USERNAME/") }
$author = '<(dc:creator|cp:lastModifiedBy)>' + [regex]::Escape($env:USERNAME) + '</'
$found = @()
$zip = [IO.Compression.ZipFile]::OpenRead($Path)
try {
    foreach ($e in $zip.Entries) {
        $s = $e.Open()
        try { $b = ReadBytes $s } finally { $s.Dispose() }
        $texts = @($utf8.GetString($b), $latin1.GetString($b), [Text.Encoding]::Unicode.GetString($b))
        if ($b.Length -gt 1) { $texts += [Text.Encoding]::Unicode.GetString($b, 1, $b.Length - 1) }
        foreach ($n in $needles) {
            foreach ($t in $texts) {
                if ($t.IndexOf($n, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $found += "$($e.FullName): $n"; break }
            }
        }
        if ($e.FullName -eq 'docProps/core.xml' -and $env:USERNAME -and $texts[0] -cmatch $author) {
            $found += "$($e.FullName): Windows user name as author"
        }
    }
}
finally { $zip.Dispose() }

if ($found.Count -gt 0) { throw ("Local data left in $Path`n  " + ($found -join "`n  ")) }
Write-Host 'No local path, computer name or user name found'
