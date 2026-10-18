#Requires -Version 5.1
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$TemplatePath = Join-Path $Root "templates\User_list_EXXXENTENCION.xlsx"
$Prefix = "http://127.0.0.1:8765/"

function Get-JsonBody {
    param($Request)
    $reader = New-Object System.IO.StreamReader($Request.InputStream, [System.Text.Encoding]::UTF8)
    try {
        $text = $reader.ReadToEnd()
    } finally {
        $reader.Close()
    }
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    return $text | ConvertFrom-Json
}

function Export-UserList {
    param($Payload)

    if (-not (Test-Path $TemplatePath)) {
        throw "Template not found"
    }
    if (-not $Payload -or -not $Payload.rows) {
        throw "No rows"
    }

    $outPath = Join-Path $env:TEMP ("exfix_export_" + [guid]::NewGuid().ToString() + ".xlsx")
    Copy-Item -Path $TemplatePath -Destination $outPath -Force

    $excel = $null
    $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        $wb = $excel.Workbooks.Open((Resolve-Path $outPath).Path)
        $ws = $wb.Worksheets.Item(1)

        $i = 0
        foreach ($row in @($Payload.rows)) {
            $excelRow = $i + 2
            $phone = [string]$row.phone
            if ([string]::IsNullOrWhiteSpace($phone)) { continue }

            $cell = $ws.Cells.Item($excelRow, 1)
            $cell.NumberFormat = "@"
            $cell.Value2 = $phone

            if ($row.firstName) {
                $ws.Cells.Item($excelRow, 2).NumberFormat = "@"
                $ws.Cells.Item($excelRow, 2).Value2 = [string]$row.firstName
            }
            if ($row.lastName) {
                $ws.Cells.Item($excelRow, 3).NumberFormat = "@"
                $ws.Cells.Item($excelRow, 3).Value2 = [string]$row.lastName
            }
            $i++
        }

        $wb.Save()
        $wb.Close($false)
        $wb = $null
        return $outPath
    }
    finally {
        if ($wb) { try { $wb.Close($false) } catch {} }
        if ($excel) {
            try { $excel.Quit() } catch {}
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Send-Bytes {
    param($Response, [int]$Status, [string]$ContentType, [byte[]]$Bytes)
    $Response.StatusCode = $Status
    $Response.ContentType = $ContentType
    $Response.ContentLength64 = $Bytes.Length
    $Response.OutputStream.Write($Bytes, 0, $Bytes.Length)
    $Response.Close()
}

$mime = @{
    ".html" = "text/html; charset=utf-8"
    ".js"   = "application/javascript; charset=utf-8"
    ".css"  = "text/css; charset=utf-8"
    ".xlsx" = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ".png"  = "image/png"
    ".ico"  = "image/x-icon"
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($Prefix)
try {
    $listener.Start()
} catch {
    Write-Host "Port 8765 is busy. Close the old window and run Start.bat again."
    throw
}

Write-Host "EXfix ready at $Prefix"

while ($listener.IsListening) {
    $ctx = $listener.GetContext()
    $req = $ctx.Request
    $res = $ctx.Response
    try {
        if ($req.HttpMethod -eq "POST" -and $req.Url.AbsolutePath -eq "/export") {
            $payload = Get-JsonBody $req
            $file = Export-UserList $payload
            $bytes = [IO.File]::ReadAllBytes($file)
            Remove-Item $file -Force -ErrorAction SilentlyContinue
            $name = if ($payload.filename) { [string]$payload.filename } else { "User_list.xlsx" }
            $res.AddHeader("Content-Disposition", "attachment; filename=""$name""")
            Send-Bytes $res 200 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" $bytes
            continue
        }

        $path = [Uri]::UnescapeDataString($req.Url.AbsolutePath.TrimStart("/"))
        if ([string]::IsNullOrWhiteSpace($path)) { $path = "index.html" }
        $full = Join-Path $Root $path
        if (-not (Test-Path $full -PathType Leaf)) {
            Send-Bytes $res 404 "text/plain; charset=utf-8" ([Text.Encoding]::UTF8.GetBytes("not found"))
            continue
        }
        $ext = [IO.Path]::GetExtension($full).ToLowerInvariant()
        $type = if ($mime.ContainsKey($ext)) { $mime[$ext] } else { "application/octet-stream" }
        Send-Bytes $res 200 $type ([IO.File]::ReadAllBytes($full))
    } catch {
        $msg = [Text.Encoding]::UTF8.GetBytes([string]$_.Exception.Message)
        try { Send-Bytes $res 500 "text/plain; charset=utf-8" $msg } catch {}
    }
}
