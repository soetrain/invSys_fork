param(
  [string]$Root = (Resolve-Path ".").Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

function New-Icon {
  param(
    [Parameter(Mandatory)] [string]$Path,
    [Parameter(Mandatory)] [int]$Size,
    [Parameter(Mandatory)] [System.Drawing.Color]$Bg,
    [Parameter(Mandatory)] [scriptblock]$Draw
  )

  $bmp = New-Object System.Drawing.Bitmap $Size, $Size
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
  $g.Clear([System.Drawing.Color]::Transparent)

  $bgBrush = New-Object System.Drawing.SolidBrush $Bg
  $g.FillEllipse($bgBrush, 0, 0, $Size-1, $Size-1)

  $stroke = [Math]::Max(2, [int]([Math]::Round($Size / 16.0)))
  $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::White), $stroke
  $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
  $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
  $pen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round

  & $Draw $g $pen $Size

  $dir = Split-Path -Parent $Path
  if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
  $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)

  $pen.Dispose()
  $bgBrush.Dispose()
  $g.Dispose()
  $bmp.Dispose()
}

function Pt([int]$Size, [double]$x, [double]$y) {
  return [System.Drawing.PointF]::new([single]($x*$Size), [single]($y*$Size))
}

function DrawDoc($g, $pen, $Size, [double]$x, [double]$y, [double]$w, [double]$h) {
  $rx = [int]([Math]::Round($x*$Size))
  $ry = [int]([Math]::Round($y*$Size))
  $rw = [int]([Math]::Round($w*$Size))
  $rh = [int]([Math]::Round($h*$Size))
  $g.DrawRectangle($pen, $rx, $ry, $rw, $rh)
  $g.DrawLine($pen, $rx + $rw*0.65, $ry, $rx + $rw, $ry + $rh*0.35)
}

function DrawCheck($g, $pen, $Size, [double]$cx, [double]$cy) {
  $p1 = Pt $Size ($cx-0.18) ($cy+0.02)
  $p2 = Pt $Size ($cx-0.04) ($cy+0.16)
  $p3 = Pt $Size ($cx+0.20) ($cy-0.10)
  $g.DrawLine($pen, $p1, $p2)
  $g.DrawLine($pen, $p2, $p3)
}

function DrawPlus($g, $pen, $Size, [double]$cx, [double]$cy) {
  $p1 = Pt $Size ($cx-0.16) $cy
  $p2 = Pt $Size ($cx+0.16) $cy
  $p3 = Pt $Size $cx ($cy-0.16)
  $p4 = Pt $Size $cx ($cy+0.16)
  $g.DrawLine($pen, $p1, $p2)
  $g.DrawLine($pen, $p3, $p4)
}

function DrawX($g, $pen, $Size, [double]$cx, [double]$cy) {
  $a = 0.15
  $g.DrawLine($pen, (Pt $Size ($cx-$a) ($cy-$a)), (Pt $Size ($cx+$a) ($cy+$a)))
  $g.DrawLine($pen, (Pt $Size ($cx-$a) ($cy+$a)), (Pt $Size ($cx+$a) ($cy-$a)))
}

function DrawArrow($g, $pen, $Size, [string]$dir) {
  if ($dir -eq 'down') {
    $g.DrawLine($pen, (Pt $Size 0.50 0.22), (Pt $Size 0.50 0.62))
    $g.DrawLine($pen, (Pt $Size 0.38 0.50), (Pt $Size 0.50 0.62))
    $g.DrawLine($pen, (Pt $Size 0.62 0.50), (Pt $Size 0.50 0.62))
  } else {
    $g.DrawLine($pen, (Pt $Size 0.50 0.70), (Pt $Size 0.50 0.30))
    $g.DrawLine($pen, (Pt $Size 0.38 0.42), (Pt $Size 0.50 0.30))
    $g.DrawLine($pen, (Pt $Size 0.62 0.42), (Pt $Size 0.50 0.30))
  }
}

function DrawBox($g, $pen, $Size) {
  $g.DrawRectangle($pen, [int]($Size*0.23), [int]($Size*0.55), [int]($Size*0.54), [int]($Size*0.22))
  $g.DrawLine($pen, (Pt $Size 0.23 0.63), (Pt $Size 0.77 0.63))
}

function DrawList($g, $pen, $Size) {
  for ($i=0; $i -lt 3; $i++) {
    $y = 0.34 + $i*0.14
    $g.DrawLine($pen, (Pt $Size 0.30 $y), (Pt $Size 0.74 $y))
    $g.DrawLine($pen, (Pt $Size 0.22 $y), (Pt $Size 0.24 $y))
  }
}

function DrawPin($g, $pen, $Size) {
  $g.DrawEllipse($pen, [int]($Size*0.36), [int]($Size*0.22), [int]($Size*0.28), [int]($Size*0.28))
  $g.DrawLine($pen, (Pt $Size 0.50 0.50), (Pt $Size 0.50 0.74))
  $g.DrawLine($pen, (Pt $Size 0.44 0.62), (Pt $Size 0.50 0.74))
  $g.DrawLine($pen, (Pt $Size 0.56 0.62), (Pt $Size 0.50 0.74))
}

function DrawSliders($g, $pen, $Size) {
  $xs = @(0.34, 0.50, 0.66)
  $ys = @(0.40, 0.56, 0.34)
  for ($i=0; $i -lt 3; $i++) {
    $x = $xs[$i]
    $g.DrawLine($pen, (Pt $Size $x 0.28), (Pt $Size $x 0.72))
    $g.DrawEllipse($pen, [int](($x-0.06)*$Size), [int](($ys[$i]-0.06)*$Size), [int]($Size*0.12), [int]($Size*0.12))
  }
}

function DrawTruck($g, $pen, $Size) {
  $g.DrawRectangle($pen, [int]($Size*0.22), [int]($Size*0.44), [int]($Size*0.34), [int]($Size*0.20))
  $g.DrawRectangle($pen, [int]($Size*0.56), [int]($Size*0.50), [int]($Size*0.20), [int]($Size*0.14))
  $g.DrawEllipse($pen, [int]($Size*0.28), [int]($Size*0.64), [int]($Size*0.12), [int]($Size*0.12))
  $g.DrawEllipse($pen, [int]($Size*0.62), [int]($Size*0.64), [int]($Size*0.12), [int]($Size*0.12))
}

function DrawInfo($g, $pen, $Size) {
  $g.DrawEllipse($pen, [int]($Size*0.28), [int]($Size*0.22), [int]($Size*0.44), [int]($Size*0.44))
  $g.DrawLine($pen, (Pt $Size 0.50 0.40), (Pt $Size 0.50 0.58))
  $g.DrawLine($pen, (Pt $Size 0.50 0.34), (Pt $Size 0.50 0.34))
}

function DrawFactory($g, $pen, $Size) {
  $g.DrawRectangle($pen, [int]($Size*0.24), [int]($Size*0.50), [int]($Size*0.52), [int]($Size*0.26))
  $g.DrawLine($pen, (Pt $Size 0.24 0.50), (Pt $Size 0.36 0.40))
  $g.DrawLine($pen, (Pt $Size 0.36 0.40), (Pt $Size 0.48 0.50))
  $g.DrawLine($pen, (Pt $Size 0.48 0.50), (Pt $Size 0.60 0.40))
  $g.DrawLine($pen, (Pt $Size 0.60 0.40), (Pt $Size 0.76 0.50))
  $g.DrawLine($pen, (Pt $Size 0.30 0.50), (Pt $Size 0.30 0.34))
}

function DrawNodes($g, $pen, $Size) {
  $pts = @(
    (Pt $Size 0.32 0.36),
    (Pt $Size 0.68 0.36),
    (Pt $Size 0.50 0.64)
  )
  $g.DrawLine($pen, $pts[0], $pts[2])
  $g.DrawLine($pen, $pts[1], $pts[2])
  foreach ($p in $pts) {
    $g.DrawEllipse($pen, [int]($p.X-$Size*0.06), [int]($p.Y-$Size*0.06), [int]($Size*0.12), [int]($Size*0.12))
  }
}

function DrawUsers($g, $pen, $Size) {
  $g.DrawEllipse($pen, [int]($Size*0.28), [int]($Size*0.26), [int]($Size*0.18), [int]($Size*0.18))
  $g.DrawEllipse($pen, [int]($Size*0.50), [int]($Size*0.30), [int]($Size*0.18), [int]($Size*0.18))
  $g.DrawLine($pen, (Pt $Size 0.24 0.70), (Pt $Size 0.48 0.70))
  $g.DrawLine($pen, (Pt $Size 0.46 0.74), (Pt $Size 0.76 0.74))
}

function DrawWrench($g, $pen, $Size) {
  $g.DrawLine($pen, (Pt $Size 0.34 0.66), (Pt $Size 0.66 0.34))
  $g.DrawEllipse($pen, [int]($Size*0.26), [int]($Size*0.58), [int]($Size*0.16), [int]($Size*0.16))
  $g.DrawLine($pen, (Pt $Size 0.62 0.30), (Pt $Size 0.74 0.22))
}

function DrawLock($g, $pen, $Size) {
  $g.DrawRectangle($pen, [int]($Size*0.34), [int]($Size*0.48), [int]($Size*0.32), [int]($Size*0.26))
  $g.DrawArc($pen, [int]($Size*0.34), [int]($Size*0.32), [int]($Size*0.32), [int]($Size*0.32), 200, 140)
}

function DrawStar($g, $pen, $Size) {
  $g.DrawLine($pen, (Pt $Size 0.50 0.30), (Pt $Size 0.50 0.68))
  $g.DrawLine($pen, (Pt $Size 0.34 0.42), (Pt $Size 0.66 0.56))
  $g.DrawLine($pen, (Pt $Size 0.66 0.42), (Pt $Size 0.34 0.56))
}

$pal = @{
  Inventory = [System.Drawing.Color]::FromArgb(46,125,50)
  Receiving = [System.Drawing.Color]::FromArgb(21,101,192)
  Shipping = [System.Drawing.Color]::FromArgb(239,108,0)
  Production = [System.Drawing.Color]::FromArgb(106,27,154)
  Designs = [System.Drawing.Color]::FromArgb(0,105,92)
  Admin = [System.Drawing.Color]::FromArgb(66,66,66)
}

$targets = @(
  @{Dir="src/InventoryDomain/Ribbon/images"; Bg=$pal.Inventory; Icons=@(
    @{Name="stock_on_hand"; Draw={ param($g,$pen,$s) DrawList $g $pen $s }},
    @{Name="inventory_logs"; Draw={ param($g,$pen,$s) DrawDoc $g $pen $s 0.28 0.22 0.44 0.56; DrawList $g $pen $s }},
    @{Name="locations"; Draw={ param($g,$pen,$s) DrawPin $g $pen $s }},
    @{Name="adjustments"; Draw={ param($g,$pen,$s) DrawSliders $g $pen $s }}
  )},
  @{Dir="src/Receiving/Ribbon/images"; Bg=$pal.Receiving; Icons=@(
    @{Name="receive_goods"; Draw={ param($g,$pen,$s) DrawBox $g $pen $s; DrawArrow $g $pen $s 'down' }},
    @{Name="post_receipt"; Draw={ param($g,$pen,$s) DrawDoc $g $pen $s 0.28 0.22 0.44 0.56; DrawCheck $g $pen $s 0.54 0.58 }},
    @{Name="verify_delivery"; Draw={ param($g,$pen,$s) $g.DrawEllipse($pen,[int]($s*0.28),[int]($s*0.30),[int]($s*0.28),[int]($s*0.28)); $g.DrawLine($pen,(Pt $s 0.52 0.54),(Pt $s 0.68 0.70)); DrawCheck $g $pen $s 0.60 0.38 }},
    @{Name="purchasing_optional"; Draw={ param($g,$pen,$s) $g.DrawLine($pen,(Pt $s 0.30 0.36),(Pt $s 0.40 0.36)); $g.DrawLine($pen,(Pt $s 0.40 0.36),(Pt $s 0.46 0.62)); $g.DrawRectangle($pen,[int]($s*0.42),[int]($s*0.42),[int]($s*0.30),[int]($s*0.18)); $g.DrawEllipse($pen,[int]($s*0.46),[int]($s*0.62),[int]($s*0.10),[int]($s*0.10)); $g.DrawEllipse($pen,[int]($s*0.62),[int]($s*0.62),[int]($s*0.10),[int]($s*0.10)) }}
  )},
  @{Dir="src/Shipping/Ribbon/images"; Bg=$pal.Shipping; Icons=@(
    @{Name="ship_order"; Draw={ param($g,$pen,$s) DrawBox $g $pen $s; DrawArrow $g $pen $s 'up' }},
    @{Name="post_shipment"; Draw={ param($g,$pen,$s) DrawTruck $g $pen $s; DrawCheck $g $pen $s 0.62 0.40 }},
    @{Name="packing_list"; Draw={ param($g,$pen,$s) DrawDoc $g $pen $s 0.28 0.22 0.44 0.56; DrawList $g $pen $s }},
    @{Name="carrier_info"; Draw={ param($g,$pen,$s) DrawInfo $g $pen $s }}
  )},
  @{Dir="src/Production/Ribbon/images"; Bg=$pal.Production; Icons=@(
    @{Name="start_run"; Draw={ param($g,$pen,$s) $g.DrawEllipse($pen,[int]($s*0.28),[int]($s*0.24),[int]($s*0.44),[int]($s*0.44)); $g.DrawLine($pen,(Pt $s 0.46 0.38),(Pt $s 0.62 0.46)); $g.DrawLine($pen,(Pt $s 0.46 0.54),(Pt $s 0.62 0.46)); $g.DrawLine($pen,(Pt $s 0.46 0.38),(Pt $s 0.46 0.54)) }},
    @{Name="post_production"; Draw={ param($g,$pen,$s) DrawFactory $g $pen $s; DrawCheck $g $pen $s 0.62 0.62 }},
    @{Name="consume_parts"; Draw={ param($g,$pen,$s) $g.DrawEllipse($pen,[int]($s*0.30),[int]($s*0.28),[int]($s*0.40),[int]($s*0.40)); $g.DrawLine($pen,(Pt $s 0.36 0.48),(Pt $s 0.64 0.48)) }},
    @{Name="finished_goods"; Draw={ param($g,$pen,$s) DrawBox $g $pen $s; DrawStar $g $pen $s }}
  )},
  @{Dir="src/DesignsDomain/Ribbon/images"; Bg=$pal.Designs; Icons=@(
    @{Name="view_bom"; Draw={ param($g,$pen,$s) DrawNodes $g $pen $s }},
    @{Name="new_version"; Draw={ param($g,$pen,$s) DrawDoc $g $pen $s 0.28 0.22 0.44 0.56; DrawPlus $g $pen $s 0.62 0.62 }},
    @{Name="release_design"; Draw={ param($g,$pen,$s) $g.DrawLine($pen,(Pt $s 0.34 0.62),(Pt $s 0.70 0.26)); $g.DrawLine($pen,(Pt $s 0.58 0.26),(Pt $s 0.70 0.26)); $g.DrawLine($pen,(Pt $s 0.70 0.26),(Pt $s 0.70 0.38)) }},
    @{Name="obsolete_design"; Draw={ param($g,$pen,$s) DrawDoc $g $pen $s 0.28 0.22 0.44 0.56; DrawX $g $pen $s 0.60 0.58 }}
  )},
  @{Dir="src/Admin/Ribbon/images"; Bg=$pal.Admin; Icons=@(
    @{Name="users_roles"; Draw={ param($g,$pen,$s) DrawUsers $g $pen $s }},
    @{Name="maintenance"; Draw={ param($g,$pen,$s) DrawWrench $g $pen $s }},
    @{Name="repair_locks"; Draw={ param($g,$pen,$s) DrawLock $g $pen $s; DrawWrench $g $pen $s }},
    @{Name="audit_logs"; Draw={ param($g,$pen,$s) $g.DrawEllipse($pen,[int]($s*0.28),[int]($s*0.22),[int]($s*0.44),[int]($s*0.44)); DrawList $g $pen $s }}
  )}
)

$sizes = @(16, 32)
foreach ($t in $targets) {
  $dir = Join-Path $Root $t.Dir
  foreach ($icon in $t.Icons) {
    foreach ($sz in $sizes) {
      $path = Join-Path $dir ("{0}_{1}.png" -f $icon.Name, $sz)
      New-Icon -Path $path -Size $sz -Bg $t.Bg -Draw $icon.Draw
    }
  }
}

Write-Output "Generated ribbon icons under src/*/Ribbon/images"
