<#
.SYNOPSIS
    Lists files of a given kind (video by default) under a folder, using the Windows Search index.

.DESCRIPTION
    Queries the Windows Search index instead of walking the disk, so results come back fast.
    Only folders that Windows Search indexes are covered.

.PARAMETER Path
    Folder to search, including its subfolders. Defaults to the current folder.

.PARAMETER Kind
    Windows Search "kind" to match, for example video, music, picture or document. Defaults to video.

.EXAMPLE
    .\Find-VideoFiles.ps1 -Path 'D:\Videos'

.EXAMPLE
    .\Find-VideoFiles.ps1 -Path 'D:\Music' -Kind music
#>
[CmdletBinding()]
param(
    [string]$Path = (Get-Location).Path,
    [string]$Kind = 'video'
)

$folder = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath.TrimEnd('\')

# Escape single quotes so folder names like "Bob's Videos" don't break the query.
$folderSql = $folder.Replace("'", "''")
$kindSql = $Kind.Replace("'", "''")

$query = "SELECT System.ItemPathDisplay FROM SYSTEMINDEX " +
         "WHERE System.Kind = '$kindSql' AND SCOPE = 'file:$folderSql'"

$connection = New-Object -ComObject ADODB.Connection
$recordSet = New-Object -ComObject ADODB.Recordset
try {
    $connection.Open("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';")
    $recordSet.Open($query, $connection)

    while (-not $recordSet.EOF) {
        $recordSet.Fields.Item('System.ItemPathDisplay').Value
        $recordSet.MoveNext()
    }
}
finally {
    if ($recordSet.State -ne 0) { $recordSet.Close() }
    if ($connection.State -ne 0) { $connection.Close() }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($recordSet)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($connection)
}
