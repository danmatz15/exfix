#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$InputPath,
    [string]$OutputPath,
    [string]$TemplatePath,
    [switch]$Gui
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $TemplatePath) {
    $TemplatePath = Join-Path $ScriptRoot "templates\User_list_EXXXENTENCION.xlsx"
}

function Normalize-Header {
    param([string]$Value)
    if ($null -eq $Value) { return "" }
    return (($Value -replace "[\u200f\u200e\ufeff]", "").Trim().ToLowerInvariant())
}

function ConvertTo-InternationalPhone {
    param($Raw)

    if ($null -eq $Raw -or [string]::IsNullOrWhiteSpace([string]$Raw)) {
        return [pscustomobject]@{ Ok = $false; Value = ""; Reason = "ריק" }
    }

    $text = ([string]$Raw).Trim()
    if ($text -match "^[+-]?\d+(\.\d+)?[eE][+-]?\d+$") {
        $text = [decimal]::Parse($text, [System.Globalization.NumberStyles]::Float).ToString("0")
    }

    $digits = [regex]::Replace($text, "[^\d]", "")
    if (-not $digits) {
        return [pscustomobject]@{ Ok = $false; Value = ""; Reason = "אין ספרות" }
    }

    if ($digits.StartsWith("00")) { $digits = $digits.Substring(2) }
    if ($digits.StartsWith("9720")) { $digits = "972" + $digits.Substring(4) }

    if ($digits.StartsWith("972")) {
        if ($digits.Length -eq 12) {
            return [pscustomobject]@{ Ok = $true; Value = $digits; Reason = "" }
        }
        if ($digits.Length -ge 11 -and $digits.Length -le 15) {
            return [pscustomobject]@{ Ok = $true; Value = $digits; Reason = "אורך לא סטנדרטי" }
        }
        return [pscustomobject]@{ Ok = $false; Value = $digits; Reason = "אורך שגוי אחרי קידומת 972" }
    }

    if ($digits -match "^0?5\d{8}$") {
        $national = if ($digits.StartsWith("0")) { $digits.Substring(1) } else { $digits }
        return [pscustomobject]@{ Ok = $true; Value = "972$national"; Reason = "" }
    }

    if ($digits -match "^0[2-489]\d{7,8}$") {
        return [pscustomobject]@{ Ok = $true; Value = "972" + $digits.Substring(1); Reason = "" }
    }

    if (-not $digits.StartsWith("0") -and $digits.Length -ge 10 -and $digits.Length -le 15) {
        return [pscustomobject]@{ Ok = $true; Value = $digits; Reason = "" }
    }

    return [pscustomobject]@{ Ok = $false; Value = $digits; Reason = "לא זוהה פורמט" }
}

function Get-PhoneColumnIndex {
    param($Headers)

    $bestIndex = 0
    $bestScore = -1
    $bestName = ""
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $header = [string]$Headers[$i]
        $norm = Normalize-Header $header
        $score = 0
        if ($norm -eq "phone numbers" -or $norm -eq "phone number") { $score = 100 }
        elseif ($norm -match "phone|טלפון|נייד|פלאפון|mobile|cell") { $score = 80 }
        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestIndex = $i
            $bestName = $header
        }
    }
    return [pscustomobject]@{ Index = $bestIndex; Header = $bestName; Score = $bestScore }
}

function Release-Com {
    param($Object)
    if ($null -ne $Object) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}

function Convert-ExcelPhoneFile {
    param(
        [Parameter(Mandatory = $true)][string]$SourceFile,
        [Parameter(Mandatory = $true)][string]$DestinationFile,
        [Parameter(Mandatory = $true)][string]$TemplateFile
    )

    if (-not (Test-Path $SourceFile)) { throw "קובץ המקור לא נמצא: $SourceFile" }
    if (-not (Test-Path $TemplateFile)) { throw "קובץ התבנית לא נמצא: $TemplateFile" }

    $excel = $null
    $sourceWb = $null
    $templateWb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        $sourceWb = $excel.Workbooks.Open((Resolve-Path $SourceFile).Path)
        $sourceWs = $sourceWb.Worksheets.Item(1)
        $used = $sourceWs.UsedRange
        $rowCount = [int]$used.Rows.Count
        $colCount = [int]$used.Columns.Count

        $headers = @()
        for ($c = 1; $c -le $colCount; $c++) {
            $headers += [string]$sourceWs.Cells.Item(1, $c).Text
        }
        $phoneCol = Get-PhoneColumnIndex -Headers $headers
        $phoneIndex = $phoneCol.Index + 1

        $firstIndex = 0
        $lastIndex = 0
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $norm = Normalize-Header $headers[$c]
            if ($firstIndex -eq 0 -and $norm -match "first name|firstname|שם פרטי|^שם$|^name$") { $firstIndex = $c + 1 }
            if ($lastIndex -eq 0 -and $norm -match "last name|lastname|שם משפחה|surname") { $lastIndex = $c + 1 }
        }

        $seen = @{}
        $valid = New-Object System.Collections.Generic.List[object]
        $preview = New-Object System.Collections.Generic.List[object]

        for ($r = 2; $r -le $rowCount; $r++) {
            $raw = [string]$sourceWs.Cells.Item($r, $phoneIndex).Text
            if ([string]::IsNullOrWhiteSpace($raw)) { continue }
            foreach ($part in ($raw -split "[\r\n,;/|]+")) {
                $piece = $part.Trim()
                if (-not $piece) { continue }
                $formatted = ConvertTo-InternationalPhone $piece
                $item = [pscustomobject]@{
                    SourceRow = $r
                    Raw       = $piece
                    Formatted = $formatted.Value
                    Ok        = [bool]$formatted.Ok
                    Reason    = $formatted.Reason
                    FirstName = if ($firstIndex) { [string]$sourceWs.Cells.Item($r, $firstIndex).Text } else { "" }
                    LastName  = if ($lastIndex) { [string]$sourceWs.Cells.Item($r, $lastIndex).Text } else { "" }
                }
                if ($item.Ok) {
                    if ($seen.ContainsKey($item.Formatted)) {
                        $item.Ok = $false
                        $item.Reason = "כפילות"
                    } else {
                        $seen[$item.Formatted] = $true
                        $valid.Add($item)
                    }
                }
                $preview.Add($item)
            }
        }

        Copy-Item -Path $TemplateFile -Destination $DestinationFile -Force
        $templateWb = $excel.Workbooks.Open((Resolve-Path $DestinationFile).Path)
        $outWs = $templateWb.Worksheets.Item(1)

        # Write only values, as text, into the Phone number column.
        for ($i = 0; $i -lt $valid.Count; $i++) {
            $row = $i + 2
            $outWs.Cells.Item($row, 1).NumberFormat = "@"
            $outWs.Cells.Item($row, 1).Value2 = [string]$valid[$i].Formatted
            if ($valid[$i].FirstName) {
                $outWs.Cells.Item($row, 2).NumberFormat = "@"
                $outWs.Cells.Item($row, 2).Value2 = [string]$valid[$i].FirstName
            }
            if ($valid[$i].LastName) {
                $outWs.Cells.Item($row, 3).NumberFormat = "@"
                $outWs.Cells.Item($row, 3).Value2 = [string]$valid[$i].LastName
            }
        }

        $templateWb.Save()
        $templateWb.Close($false)
        $sourceWb.Close($false)

        return [pscustomobject]@{
            PhoneColumn = $phoneCol.Header
            Preview     = $preview
            ValidCount  = $valid.Count
            Output      = $DestinationFile
        }
    }
    finally {
        if ($templateWb) { try { $templateWb.Close($false) } catch {} }
        if ($sourceWb) { try { $sourceWb.Close($false) } catch {} }
        if ($excel) {
            $excel.Quit()
            Release-Com $excel
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Show-Gui {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "תיקון מספרי טלפון"
    $form.RightToLeft = "Yes"
    $form.RightToLeftLayout = $true
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(980, 680)
    $form.BackColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
    $form.ForeColor = [System.Drawing.Color]::WhiteSmoke
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = "העלאת אקסל עם מספרי פלאפון והמרה לפורמט User list"
    $title.AutoSize = $true
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $title.Location = New-Object System.Drawing.Point(20, 18)
    $form.Controls.Add($title)

    $info = New-Object System.Windows.Forms.Label
    $info.Text = "המספרים ייכתבו כערכים בלבד בעמודת Phone number"
    $info.AutoSize = $true
    $info.Location = New-Object System.Drawing.Point(20, 52)
    $form.Controls.Add($info)

    $pickBtn = New-Object System.Windows.Forms.Button
    $pickBtn.Text = "בחירת קובץ מקור"
    $pickBtn.Size = New-Object System.Drawing.Size(180, 36)
    $pickBtn.Location = New-Object System.Drawing.Point(20, 86)
    $form.Controls.Add($pickBtn)

    $saveBtn = New-Object System.Windows.Forms.Button
    $saveBtn.Text = "שמירה בפורמט הנכון"
    $saveBtn.Size = New-Object System.Drawing.Size(180, 36)
    $saveBtn.Location = New-Object System.Drawing.Point(210, 86)
    $saveBtn.Enabled = $false
    $form.Controls.Add($saveBtn)

    $status = New-Object System.Windows.Forms.Label
    $status.Text = "לא נבחר קובץ"
    $status.AutoSize = $true
    $status.Location = New-Object System.Drawing.Point(20, 132)
    $form.Controls.Add($status)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(20, 168)
    $grid.Size = New-Object System.Drawing.Size(920, 450)
    $grid.ReadOnly = $true
    $grid.AutoSizeColumnsMode = "Fill"
    $grid.BackgroundColor = [System.Drawing.Color]::FromArgb(17, 24, 39)
    $form.Controls.Add($grid)

    $script:LastResult = $null
    $script:LastSource = $null

    $pickBtn.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "Excel|*.xlsx;*.xls;*.csv"
        if ($dialog.ShowDialog() -eq "OK") {
            $script:LastSource = $dialog.FileName
            $tempOut = Join-Path $env:TEMP ("exfix_preview_" + [guid]::NewGuid().ToString() + ".xlsx")
            try {
                $script:LastResult = Convert-ExcelPhoneFile -SourceFile $dialog.FileName -DestinationFile $tempOut -TemplateFile $TemplatePath
                $grid.DataSource = [System.Collections.ArrayList]@($script:LastResult.Preview)
                $status.Text = "עמודה שזוהתה: $($script:LastResult.PhoneColumn) | תקינים: $($script:LastResult.ValidCount)"
                $saveBtn.Enabled = $script:LastResult.ValidCount -gt 0
                Remove-Item $tempOut -Force -ErrorAction SilentlyContinue
            } catch {
                [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "שגיאה")
            }
        }
    })

    $saveBtn.Add_Click({
        if (-not $script:LastSource) { return }
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = "Excel|*.xlsx"
        $dialog.FileName = ([IO.Path]::GetFileNameWithoutExtension($script:LastSource) + "_User_list.xlsx")
        if ($dialog.ShowDialog() -eq "OK") {
            try {
                $saved = Convert-ExcelPhoneFile -SourceFile $script:LastSource -DestinationFile $dialog.FileName -TemplateFile $TemplatePath
                [System.Windows.Forms.MessageBox]::Show("נשמר: $($saved.Output)`nמספרים תקינים: $($saved.ValidCount)", "הסתיים")
            } catch {
                [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "שגיאה")
            }
        }
    })

    [void]$form.ShowDialog()
}

if ($Gui -or -not $InputPath) {
    Show-Gui
    return
}

if (-not $OutputPath) {
    $OutputPath = Join-Path (Split-Path $InputPath) (([IO.Path]::GetFileNameWithoutExtension($InputPath)) + "_User_list.xlsx")
}

$result = Convert-ExcelPhoneFile -SourceFile $InputPath -DestinationFile $OutputPath -TemplateFile $TemplatePath
Write-Output $result
