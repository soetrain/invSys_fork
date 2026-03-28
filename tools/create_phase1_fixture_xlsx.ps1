Param(
    [string]$OutputDir = "tests/fixtures"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Utf8File {
    Param(
        [string]$Path,
        [string]$Content
    )

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function Convert-ToColumnName {
    Param([int]$Index)
    $name = ""
    $i = $Index
    while ($i -gt 0) {
        $rem = ($i - 1) % 26
        $name = [char](65 + $rem) + $name
        $i = [int](($i - 1) / 26)
    }
    return $name
}

function Escape-Xml {
    Param([string]$Value)
    if ($null -eq $Value) { return "" }
    return [System.Security.SecurityElement]::Escape($Value)
}

function New-SheetXml {
    Param(
        [object[][]]$Rows
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sb.AppendLine('  <sheetData>')

    for ($r = 0; $r -lt $Rows.Count; $r++) {
        $rowIdx = $r + 1
        [void]$sb.AppendLine("    <row r=""$rowIdx"">")
        $rowVals = $Rows[$r]
        for ($c = 0; $c -lt $rowVals.Count; $c++) {
            $col = Convert-ToColumnName ($c + 1)
            $cellRef = "$col$rowIdx"
            $val = $rowVals[$c]

            if ($null -eq $val -or $val -eq "") {
                [void]$sb.AppendLine("      <c r=""$cellRef"" t=""inlineStr""><is><t></t></is></c>")
            } elseif ($val -is [int] -or $val -is [long] -or $val -is [double] -or $val -is [decimal]) {
                [void]$sb.AppendLine("      <c r=""$cellRef""><v>$val</v></c>")
            } elseif ($val -is [bool]) {
                $boolInt = if ($val) { 1 } else { 0 }
                [void]$sb.AppendLine("      <c r=""$cellRef"" t=""b""><v>$boolInt</v></c>")
            } else {
                $text = Escape-Xml ([string]$val)
                [void]$sb.AppendLine("      <c r=""$cellRef"" t=""inlineStr""><is><t>$text</t></is></c>")
            }
        }
        [void]$sb.AppendLine('    </row>')
    }

    [void]$sb.AppendLine('  </sheetData>')
    [void]$sb.AppendLine('</worksheet>')
    return $sb.ToString()
}

function New-MinimalXlsx {
    Param(
        [string]$Path,
        [hashtable[]]$Sheets
    )

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("xlsx_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tempRoot | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot "_rels") | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot "docProps") | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl") | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl/_rels") | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot "xl/worksheets") | Out-Null

    $contentTypes = @(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '  <Default Extension="xml" ContentType="application/xml"/>',
        '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    )
    for ($i = 0; $i -lt $Sheets.Count; $i++) {
        $sheetNum = $i + 1
        $contentTypes += "  <Override PartName=""/xl/worksheets/sheet$sheetNum.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>"
    }
    $contentTypes += '</Types>'
    Write-Utf8File -Path (Join-Path $tempRoot "[Content_Types].xml") -Content ($contentTypes -join "`r`n")

    $rels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'@
    Write-Utf8File -Path (Join-Path $tempRoot "_rels/.rels") -Content $rels

    $app = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
</Properties>
'@
    Write-Utf8File -Path (Join-Path $tempRoot "docProps/app.xml") -Content $app

    $created = (Get-Date).ToUniversalTime().ToString("s") + "Z"
    $core = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>invSys fixture</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">$created</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$created</dcterms:modified>
</cp:coreProperties>
"@
    Write-Utf8File -Path (Join-Path $tempRoot "docProps/core.xml") -Content $core

    $wb = New-Object System.Text.StringBuilder
    [void]$wb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$wb.AppendLine('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    [void]$wb.AppendLine('  <sheets>')
    for ($i = 0; $i -lt $Sheets.Count; $i++) {
        $sheetNum = $i + 1
        $name = Escape-Xml $Sheets[$i].Name
        [void]$wb.AppendLine("    <sheet name=""$name"" sheetId=""$sheetNum"" r:id=""rId$sheetNum""/>")
    }
    [void]$wb.AppendLine('  </sheets>')
    [void]$wb.AppendLine('</workbook>')
    Write-Utf8File -Path (Join-Path $tempRoot "xl/workbook.xml") -Content $wb.ToString()

    $wbRels = New-Object System.Text.StringBuilder
    [void]$wbRels.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$wbRels.AppendLine('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
    for ($i = 0; $i -lt $Sheets.Count; $i++) {
        $sheetNum = $i + 1
        [void]$wbRels.AppendLine("  <Relationship Id=""rId$sheetNum"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet$sheetNum.xml""/>")
    }
    $styleRid = $Sheets.Count + 1
    [void]$wbRels.AppendLine("  <Relationship Id=""rId$styleRid"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>")
    [void]$wbRels.AppendLine('</Relationships>')
    Write-Utf8File -Path (Join-Path $tempRoot "xl/_rels/workbook.xml.rels") -Content $wbRels.ToString()

    $styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="1"><xf xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
'@
    Write-Utf8File -Path (Join-Path $tempRoot "xl/styles.xml") -Content $styles

    for ($i = 0; $i -lt $Sheets.Count; $i++) {
        $sheetNum = $i + 1
        $xml = New-SheetXml -Rows $Sheets[$i].Rows
        Write-Utf8File -Path (Join-Path $tempRoot "xl/worksheets/sheet$sheetNum.xml") -Content $xml
    }

    $zipPath = [System.IO.Path]::ChangeExtension($Path, ".zip")
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    if (Test-Path $Path) { Remove-Item $Path -Force }

    Compress-Archive -Path (Join-Path $tempRoot "*") -DestinationPath $zipPath -Force
    Move-Item -Path $zipPath -Destination $Path -Force

    Remove-Item -Path $tempRoot -Recurse -Force
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$configSheets = @(
    @{
        Name = "WarehouseConfig"
        Rows = @(
            @("WarehouseId", "WarehouseName", "Timezone", "DefaultLocation", "BatchSize", "LockTimeoutMinutes", "HeartbeatIntervalSeconds", "MaxLockHoldMinutes", "SnapshotCadence", "BackupCadence", "PathDataRoot", "PathBackupRoot", "PathSharePointRoot", "DesignsEnabled", "PoisonRetryMax", "AuthCacheTTLSeconds", "ProcessorServiceUserId", "FF_DesignsEnabled", "FF_OutlookAlerts", "FF_AutoSnapshot", "AutoRefreshIntervalSeconds"),
            @("WH1", "Main Warehouse", "UTC", "A1", 500, 3, 30, 2, "PER_BATCH", "DAILY", "C:\invSys\WH1\", "C:\invSys\Backups\WH1\", "", $false, 3, 300, "svc_processor", $false, $false, $true, 0)
        )
    },
    @{
        Name = "StationConfig"
        Rows = @(
            @("StationId", "WarehouseId", "StationName", "RoleDefault"),
            @("S1", "WH1", "STATION-01", "RECEIVE")
        )
    }
)

$authSheets = @(
    @{
        Name = "Users"
        Rows = @(
            @("UserId", "DisplayName", "PinHash", "Status", "ValidFrom", "ValidTo"),
            @("user1", "User One", "", "Active", "", ""),
            @("user2", "User Two", "", "Active", "", "")
        )
    },
    @{
        Name = "Capabilities"
        Rows = @(
            @("UserId", "Capability", "WarehouseId", "StationId", "Status", "ValidFrom", "ValidTo"),
            @("user1", "RECEIVE_POST", "WH1", "S1", "Active", "", ""),
            @("user1", "INBOX_PROCESS", "WH1", "*", "Active", "", "")
        )
    }
)

New-MinimalXlsx -Path (Join-Path $OutputDir "WH1.invSys.Config.sample.xlsx") -Sheets $configSheets
New-MinimalXlsx -Path (Join-Path $OutputDir "WH1.invSys.Auth.sample.xlsx") -Sheets $authSheets

Write-Host "Created fixture workbooks in $OutputDir"
