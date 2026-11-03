param(
    [string]$OutDir = "C:\Users\tooth\AppData\Local\Temp\opencode\modelling\DCF-Model"
)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ============================================================
# PROJECT MERIDIAN - Discounted Cash Flow (uFCF) model builder
# Excel COM: live formulas, then Excel calculates and we verify
# every key output against this independent mirror.
# All figures in USD millions unless per share. Fictitious data.
# ============================================================

function rgb([int]$r, [int]$g, [int]$b) { return [int]($r + 256 * $g + 65536 * $b) }
$NAVY   = rgb 31 78 121
$LIGHT  = rgb 234 242 251
$BLUEF  = rgb 0 0 255      # input font
$GREENF = rgb 0 128 0      # cross-sheet link font
$GRAYT  = rgb 84 101 121
$YELLOW = rgb 255 249 231  # input fill

# ---------- mirror math (ground truth) ----------
$rf = 0.040; $erp = 0.050; $beta = 1.0
$wacc = $rf + $erp * $beta                       # 9.00 percent
$midYear = $true
$gRev = @(0.08, 0.07, 0.06, 0.05, 0.04)
$ebitdaMargin = 0.30; $daPct = 0.07; $capexPct = 0.08; $nwcPct = 0.12; $tax = 0.25
$rev0 = 1000.0
$rev = @(); $prev = $rev0
foreach ($gr in $gRev) { $prev = $prev * (1.0 + $gr); $rev += $prev }
$ebitda = @(0..4 | ForEach-Object { $rev[$_] * $ebitdaMargin })
$da     = @(0..4 | ForEach-Object { $rev[$_] * $daPct })
$ebit   = @(0..4 | ForEach-Object { $ebitda[$_] - $da[$_] })
$nopat  = @(0..4 | ForEach-Object { $ebit[$_] * (1.0 - $tax) })
$capex  = @(0..4 | ForEach-Object { $rev[$_] * $capexPct })
$dNwc   = @(); $pv = $rev0
for ($i = 0; $i -lt 5; $i++) { $dNwc += (($rev[$i] - $pv) * $nwcPct); $pv = $rev[$i] }
$fcf = @(0..4 | ForEach-Object { $nopat[$_] + $da[$_] - $capex[$_] - $dNwc[$_] })
$exps = @(0.5, 1.5, 2.5, 3.5, 4.5)
$dfs = @(0..4 | ForEach-Object { 1.0 / [Math]::Pow(1.0 + $wacc, $exps[$_]) })
$pvFcf = @(0..4 | ForEach-Object { $fcf[$_] * $dfs[$_] })
$sumPV = ($pvFcf | Measure-Object -Sum).Sum
$tg = 0.025; $exitMult = 9.0
$tvG = $fcf[4] * (1.0 + $tg) / ($wacc - $tg); $pvTvG = $tvG * $dfs[4]
$tvX = $exitMult * $ebitda[4];                 $pvTvX = $tvX * $dfs[4]
$debt = 800.0; $cash = 100.0; $netDebt = $debt - $cash
$shares = 50.0; $curPrice = 45.0
$evX = $sumPV + $pvTvX; $eqX = $evX - $netDebt; $pxX = $eqX / $shares; $upX = $pxX / $curPrice - 1.0
$evG = $sumPV + $pvTvG; $eqG = $evG - $netDebt; $pxG = $eqG / $shares; $upG = $pxG / $curPrice - 1.0
$impliedXofG = $tvG / $ebitda[4]
function PriceGordon([double]$w, [double]$g) {
    $s = 0.0
    for ($i = 0; $i -lt 5; $i++) { $s += $fcf[$i] / [Math]::Pow(1.0 + $w, $exps[$i]) }
    $tv = $fcf[4] * (1.0 + $g) / ($w - $g)
    return ($s + $tv * [Math]::Pow(1.0 + $w, -$exps[4]) - $netDebt) / $shares
}
function PriceExit([double]$w, [double]$m) {
    $s = 0.0
    for ($i = 0; $i -lt 5; $i++) { $s += $fcf[$i] / [Math]::Pow(1.0 + $w, $exps[$i]) }
    return ($s + $m * $ebitda[4] * [Math]::Pow(1.0 + $w, -$exps[4]) - $netDebt) / $shares
}
Write-Output ("MIRROR  uFCF Y1..Y5: " + (($fcf | ForEach-Object { [Math]::Round($_, 3) }) -join ", "))
Write-Output ("MIRROR  sumPV={0:N3}  TVg={1:N3} PVg={2:N3}  TVx={3:N3} PVx={4:N3}" -f $sumPV, $tvG, $pvTvG, $tvX, $pvTvX)
Write-Output ("MIRROR  EV exit={0:N3}  EV gordon={1:N3}  price exit={2:N2} (+{3:P1})  price gordon={4:N2} (+{5:P1})  impliedTerminalMultiple={6:N2}x" -f $evX, $evG, $pxX, $upX, $pxG, $upG, $impliedXofG)

# ---------- helpers ----------
function SetText($ws, [string]$addr, $v) {
    # PS 5.1's COM binding is inconsistent: strings/dates via InvokeMember,
    # numbers via the plain setter, with the other path as fallback.
    $r = $ws.Range($addr)
    if ($v -is [datetime]) { $v = $v.ToOADate() }
    if ($v -is [bool]) { $v = [double][int]$v }
    try {
        if ($v -is [string]) {
            [void]$r.GetType().InvokeMember([string]"Value2", [System.Reflection.BindingFlags]::SetProperty, $null, $r, [object[]]@([System.Object]$v))
        } else {
            $r.Value2 = $v
        }
    } catch {
        try {
            $r = $ws.Range($addr)
            if ($v -is [string]) {
                $r.Value2 = $v
            } else {
                [void]$r.GetType().InvokeMember([string]"Value2", [System.Reflection.BindingFlags]::SetProperty, $null, $r, [object[]]@([System.Object]$v))
            }
        } catch {
            throw ("SetText failed addr='" + $addr + "' v=" + $v + " :: " + $_.Exception.Message)
        }
    }
}
function SetF($ws, [string]$addr, [string]$f) { $r = $ws.Range($addr); $r.Formula = $f }
function SetFmt($ws, [string]$addr, [string]$fmt) { $r = $ws.Range($addr); $r.NumberFormat = $fmt }
function SetFont($ws, [string]$addr, [int]$color, [bool]$bold, [double]$size, [string]$name = "Calibri") {
    $r = $ws.Range($addr); $r.Font.Color = $color; $r.Font.Bold = $bold; $r.Font.Size = $size; $r.Font.Name = $name
}
function SectionHeader($ws, [string]$addr) {
    $r = $ws.Range($addr); $r.Merge()
    $r.Interior.Color = $NAVY; $r.Font.Color = 16777215; $r.Font.Bold = $true; $r.Font.Size = 11
    $r.HorizontalAlignment = -4131; $r.VerticalAlignment = -4108
}
function SubHeader($ws, [string]$addr) {
    $r = $ws.Range($addr); $r.Merge()
    $r.Interior.Color = $LIGHT; $r.Font.Color = $NAVY; $r.Font.Bold = $true; $r.Font.Size = 10
}
function Freeze($xl, $ws, [string]$cell) { $ws.Activate(); $ws.Range($cell).Select() | Out-Null; $xl.ActiveWindow.FreezePanes = $true }
function ColW($ws, [string]$col, [double]$w) { $ws.Columns.Item($col).ColumnWidth = $w }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$xl.AskToUpdateLinks = $false
try {
    $wb = $xl.Workbooks.Add()
    $wb.Worksheets.Item(1).Name = "Cover"
    foreach ($n in @("Assumptions", "OperatingModel", "DCFValuation", "Sensitivity")) {
        $s = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
        $s.Name = $n
    }

    # ============ COVER ============
    $c = $wb.Worksheets.Item("Cover")
    $c.Tab.Color = $NAVY
    ColW $c "A" 3; ColW $c "B" 62; ColW $c "C" 20
    SetText $c "B2" "DCF VALUATION MODEL - PROJECT MERIDIAN"
    SetFont $c "B2" $NAVY $true 20 "Calibri"
    SetText $c "B3" "Target: Meridian Industrials, Inc. | Unlevered free cash flow, mid-year convention"
    SetFont $c "B3" $GRAYT $false 11
    SetText $c "B5" "Contents"
    SetFont $c "B5" $NAVY $true 12
    SetText $c "B6" "Assumptions    Market data, CAPM WACC, operating drivers, terminal value and capital structure"
    SetText $c "B7" "OperatingModel  Revenue to unlevered free cash flow, 5-year forecast (2027E-2031E)"
    SetText $c "B8" "DCFValuation    Discounting, Gordon-growth and exit-multiple terminal value, enterprise to equity"
    SetText $c "B9" "Sensitivity     Implied share price by WACC x terminal growth and WACC x exit multiple"
    SetFont $c "B6:B9" (rgb 40 40 40) $false 10
    SetText $c "B11" "Conventions: blue font = input (change these); black = formula; green = link from another tab. All formulas are live and auditable."
    SetFont $c "B11" $NAVY $false 10
    SetText $c "B12" "Units: USD millions unless per-share. Illustrative training model with fictitious data - not investment advice."
    SetFont $c "B12" $GRAYT $false 10
    SetText $c "B14" ("Prepared {0} by aminsaifanoa-web" -f (Get-Date -Format "d MMM yyyy"))
    SetFont $c "B14" $GRAYT $false 10

    # ============ ASSUMPTIONS ============
    $a = $wb.Worksheets.Item("Assumptions")
    $a.Tab.Color = rgb 191 143 34
    ColW $a "A" 2; ColW $a "B" 38
    foreach ($l in @("C", "D", "E", "F", "G", "H")) { ColW $a $l 12.5 }
    SetText $a "B2" "ASSUMPTIONS & DRIVERS"; SetFont $a "B2" $NAVY $true 14
    SetText $a "B3" "Blue = input"; SetFont $a "B3" $BLUEF $false 9

    SectionHeader $a "B4:G4"; SetText $a "B4" "  MARKET / CAPITAL"
    SetText $a "B5" "Valuation date"; SetText $a "C5" ([datetime]"2026-12-31"); SetFont $a "C5" $BLUEF $true 10; SetFmt $a "C5" "dd mmm yyyy"
    $r = $a.Range("B5"); $r.Font.Bold = $true

    SectionHeader $a "B7:G7"; SetText $a "B7" "  WACC (CAPM)"
    SetText $a "B8" "Risk-free rate";              SetText $a "C8" 0.040
    SetText $a "B9" "Equity risk premium";         SetText $a "C9" 0.050
    SetText $a "B10" "Levered beta";               SetText $a "C10" 1.0
    SetText $a "B11" "WACC (CAPM)";                SetF   $a "C11" "=C8+C9*C10"
    SetText $a "B12" "Mid-year convention (1 = yes, 0 = no)"; SetText $a "C12" 1.0
    SetFmt $a "C12" "0"
    SetFont $a "C8:C10" $BLUEF $true 10; SetFmt $a "C8:C9" "0.00%"; SetFmt $a "C10" "0.00"
    SetFont $a "C11" 0 $true 11; SetFmt $a "C11" "0.00%"; SetFont $a "B11" $NAVY $true 10
    SetFont $a "C12" $BLUEF $true 10; SetFont $a "B8:B10,B12" 40 40 40 $false 10

    SectionHeader $a "B14:G14"; SetText $a "B14" "  OPERATING DRIVERS (forecast years)"
    SetText $a "B15" "Year"
    foreach ($i in 0..4) { $l = [char](67 + $i); SetText $a ([string]$l + "15") ("202" + (7 + $i) + "E"); SetFont $a ([string]$l + "15") $NAVY $true 10 }
    $rows = @(
        @{ n = 16; label = "Revenue growth (%)"        ; vals = @(0.08, 0.07, 0.06, 0.05, 0.04) }
        @{ n = 17; label = "EBITDA margin (%)"         ; vals = @(0.30, 0.30, 0.30, 0.30, 0.30) }
        @{ n = 18; label = "D&A (% of revenue)"        ; vals = @(0.07, 0.07, 0.07, 0.07, 0.07) }
        @{ n = 19; label = "CapEx (% of revenue)"      ; vals = @(0.08, 0.08, 0.08, 0.08, 0.08) }
        @{ n = 20; label = "Net working capital (% rev)"; vals = @(0.12, 0.12, 0.12, 0.12, 0.12) }
    )
    foreach ($row in $rows) {
        SetText $a ("B" + $row.n) $row.label; SetFont $a ("B" + $row.n) 40 40 40 $false 10
        for ($i = 0; $i -lt 5; $i++) {
            $l = [char](67 + $i)
            SetText $a ([string]$l + $row.n) $row.vals[$i]
            SetFmt $a ([string]$l + $row.n) "0.0%"; SetFont $a ([string]$l + $row.n) $BLUEF $true 10
        }
    }
    SetText $a "B21" "Tax rate"; SetText $a "C21" 0.25
    SetFont $a "C21" $BLUEF $true 10; SetFmt $a "C21" "0.0%"; SetFont $a "B21" 40 40 40 $false 10
    SetText $a "B22" "LTM revenue (2026A, base)"; SetText $a "C22" 1000.0
    SetFont $a "C22" $BLUEF $true 10; SetFmt $a "C22" "#,##0.0"; SetFont $a "B22" 40 40 40 $false 10

    SectionHeader $a "B24:G24"; SetText $a "B24" "  TERMINAL VALUE"
    SetText $a "B25" "Gordon growth rate (g)";         SetText $a "C25" 0.025
    SetText $a "B26" "Exit EV / EBITDA (terminal)";    SetText $a "C26" 9.0
    SetFont $a "C25" $BLUEF $true 10; SetFmt $a "C25" "0.0%"
    SetFont $a "C26" $BLUEF $true 10; SetFmt $a "C26" '0.0"x"'
    SetFont $a "B25:B26" 40 40 40 $false 10

    SectionHeader $a "B28:G28"; SetText $a "B28" "  CAPITAL STRUCTURE & SHARES"
    SetText $a "B29" "Total debt";                 SetText $a "C29" 800.0
    SetText $a "B30" "Cash and equivalents";       SetText $a "C30" 100.0
    SetText $a "B31" "Net debt";                   SetF   $a "C31" "=C29-C30"
    SetText $a "B32" "Diluted shares (mm)";        SetText $a "C32" 50.0
    SetText $a "B33" "Current share price";        SetText $a "C33" 45.0
    SetFont $a "C29:C30" $BLUEF $true 10; SetFmt $a "C29:C30" "#,##0.0"; SetFont $a "C32" $BLUEF $true 10; SetFmt $a "C32" "#,##0.0"
    SetFont $a "C33" $BLUEF $true 10; SetFmt $a "C33" '$#,##0.00'
    SetFont $a "C31" 0 $true 10; SetFmt $a "C31" "#,##0.0"
    SetFont $a "B29:B33" 40 40 40 $false 10

    # ============ OPERATING MODEL ============
    $o = $wb.Worksheets.Item("OperatingModel")
    $o.Tab.Color = rgb 84 101 121
    ColW $o "A" 2; ColW $o "B" 36
    foreach ($l in @("C", "D", "E", "F", "G", "H")) { ColW $o $l 12.5 }
    SetText $o "B2" "OPERATING MODEL (USD mm)"; SetFont $o "B2" $NAVY $true 14
    SetText $o "B4" "USD mm"; SetFont $o "B4" $NAVY $true 10
    foreach ($i in 0..4) { $l = [char](67 + $i); SetText $o ([string]$l + "4") ("202" + (7 + $i) + "E"); SetFont $o ([string]$l + "4") $NAVY $true 10; $o.Range([string]$l + "4").Borders.Item(9).LineStyle = 1 }
    SetText $o "H4" "Total"; SetFont $o "H4" $NAVY $true 10

    # Revenue row 5 (green cross-sheet for base ref)
    SetF $o "C5" '=Assumptions!$C$22*(1+Assumptions!C$16)'
    for ($i = 1; $i -le 4; $i++) {
        $prev = [char](66 + $i); $cur = [char](67 + $i)
        SetF $o ([string]$cur + "5") ('=' + $prev + '5*(1+Assumptions!' + $cur + '$16)')
    }
    SetText $o "B5" "Revenue"; SetFmt $o "C5:H5" "#,##0.0"; SetFont $o "C5:G5" $GREENF $false 10
    # EBITDA row 6
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $o ([string]$l + "6") ('=' + $l + '5*Assumptions!' + $l + '$17') }
    SetText $o "B6" "EBITDA"; SetFmt $o "C6:H6" "#,##0.0"; SetFont $o "B6" $NAVY $true 10
    # margin row 7
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $o ([string]$l + "7") ("=" + $l + "6/" + $l + "5") }
    SetText $o "B7" "EBITDA margin"; SetFmt $o "C7:G7" "0.0%"
    # D&A row 8
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $o ([string]$l + "8") ('=' + $l + '5*Assumptions!' + $l + '$18') }
    SetText $o "B8" "Depreciation & amortisation"; SetFmt $o "C8:H8" "#,##0.0"
    # EBIT row 9
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $o ([string]$l + "9") ("=" + $l + "6-" + $l + "8") }
    SetText $o "B9" "EBIT"; SetFmt $o "C9:H9" "#,##0.0"; SetFont $o "B9" $NAVY $true 10
    # EBIT margin row 10
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $o ([string]$l + "10") ("=" + $l + "9/" + $l + "5") }
    SetText $o "B10" "EBIT margin"; SetFmt $o "C10:G10" "0.0%"
    # NOPAT row 11
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $o ([string]$l + "11") ('=' + $l + '9*(1-Assumptions!$C$21)') }
    SetText $o "B11" "NOPAT (EBIT x (1 - tax))"; SetFmt $o "C11:H11" "#,##0.0"
    # add D&A row 12
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $o ([string]$l + "12") ("=" + $l + "8") }
    SetText $o "B12" "Add: D&A"; SetFmt $o "C12:H12" "#,##0.0"
    # less CapEx row 13
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $o ([string]$l + "13") ('=-' + $l + '5*Assumptions!' + $l + '$19') }
    SetText $o "B13" "Less: capital expenditure"; SetFmt $o "C13:H13" "#,##0.0"
    # less change in NWC row 14
    SetF $o "C14" '=-(C5-Assumptions!$C$22)*Assumptions!C$20'
    for ($i = 1; $i -le 4; $i++) {
        $prev = [char](66 + $i); $cur = [char](67 + $i)
        SetF $o ([string]$cur + "14") ('=-(' + $cur + '5-' + $prev + '5)*Assumptions!' + $cur + '$20')
    }
    SetText $o "B14" "Less: change in net working capital"; SetFmt $o "C14:H14" "#,##0.0"
    # uFCF row 15
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $o ([string]$l + "15") ("=" + $l + "11+" + $l + "12+" + $l + "13+" + $l + "14") }
    SetText $o "B15" "Unlevered free cash flow (uFCF)"; SetFmt $o "C15:H15" "#,##0.0"; SetFont $o "B15" $NAVY $true 10
    $o.Range("C15:H15").Borders.Item(9).LineStyle = 1
    # margin row 16
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $o ([string]$l + "16") ("=" + $l + "15/" + $l + "5") }
    SetText $o "B16" "uFCF margin"; SetFmt $o "C16:G16" "0.0%"
    # totals
    SetF $o "H5" "=SUM(C5:G5)"; SetF $o "H6" "=SUM(C6:G6)"; SetF $o "H15" "=SUM(C15:G15)"
    SetFmt $o "H5:H6" "#,##0.0"; SetFmt $o "H15" "#,##0.0"; SetFont $o "H5:H6,H15" 0 $true 10
    Freeze $xl $o "C5"

    # ============ DCF VALUATION ============
    $v = $wb.Worksheets.Item("DCFValuation")
    $v.Tab.Color = rgb 0 128 96
    ColW $v "A" 2; ColW $v "B" 44
    foreach ($l in @("C", "D", "E", "F", "G", "H")) { ColW $v $l 12.5 }
    SetText $v "B2" "DCF VALUATION"; SetFont $v "B2" $NAVY $true 14
    foreach ($i in 0..4) { $l = [char](67 + $i); SetText $v ([string]$l + "4") ("202" + (7 + $i) + "E"); SetFont $v ([string]$l + "4") $NAVY $true 10 }
    SetText $v "H4" "Total"; SetFont $v "H4" $NAVY $true 10

    # FCF link row + exponents + df + PV
    for ($i = 0; $i -le 4; $i++) {
        $l = [char](67 + $i)
        SetF $v ([string]$l + "5") ("=OperatingModel!" + $l + "15")
        SetF $v ([string]$l + "6") ('=' + (1 + $i) + '-0.5*Assumptions!$C$12')
        SetF $v ([string]$l + "7") ('=1/(1+Assumptions!$C$11)^' + $l + '6')
        SetF $v ([string]$l + "8") ("=" + $l + "5*" + $l + "7")
    }
    SetText $v "B5" "Unlevered free cash flow"; SetFmt $v "C5:H5" "#,##0.0"; SetFont $v "C5:G5" $GREENF $false 10
    SetText $v "B6" "Discount period (years, mid-year adjusted)"; SetFmt $v "C6:G6" "0.0"
    SetText $v "B7" "Discount factor"; SetFmt $v "C7:G7" "0.0000"
    SetText $v "B8" "PV of uFCF"; SetFmt $v "C8:H8" "#,##0.0"; SetFont $v "B8" $NAVY $true 10
    SetF $v "H8" "=SUM(C8:G8)"
    $v.Range("C8:H8").Borders.Item(9).LineStyle = 1
    SetFont $v "B5:B8" 40 40 40 $false 10; SetFont $v "B8" $NAVY $true 10

    # terminal value block
    SetF $v "H10" '=G5*(1+Assumptions!$C$25)/(Assumptions!$C$11-Assumptions!$C$25)'
    SetF $v "H11" '=Assumptions!$C$26*OperatingModel!$G$6'
    SetF $v "H12" "=G7"
    SetF $v "H13" "=H10*H12"
    SetF $v "H14" "=H11*H12"
    SetText $v "B10" "Terminal value - Gordon growth (g)"; SetText $v "B11" "Terminal value - exit EV/EBITDA"
    SetText $v "B12" "Discount factor at exit (year 5)";  SetText $v "B13" "PV of TV - Gordon"
    SetText $v "B14" "PV of TV - exit multiple"
    SetFmt $v "H10:H14" "#,##0.0"; SetFont $v "B10:B14" 40 40 40 $false 10

    # enterprise -> equity
    SetF $v "H16" "=H8+H14"; SetText $v "B16" "Enterprise value - exit multiple (primary)"; SetFont $v "B16" $NAVY $true 10
    SetF $v "H17" "=H8+H13"; SetText $v "B17" "Enterprise value - Gordon growth";           SetFont $v "B17" $NAVY $true 10
    SetF $v "H18" '=-Assumptions!$C$31'; SetText $v "B18" "Less: net debt at valuation"
    SetF $v "H19" "=H16+H18"; SetText $v "B19" "Equity value - exit multiple"; SetFont $v "B19" $NAVY $true 10
    SetF $v "H20" "=H17+H18"; SetText $v "B20" "Equity value - Gordon growth"; SetFont $v "B20" $NAVY $true 10
    SetF $v "H21" '=H19/Assumptions!$C$32'; SetText $v "B21" "Implied share price - exit multiple"; SetFont $v "B21" $NAVY $true 10
    SetF $v "H22" '=H20/Assumptions!$C$32'; SetText $v "B22" "Implied share price - Gordon growth"
    SetF $v "H23" '=H21/Assumptions!$C$33-1'; SetText $v "B23" "Upside vs current price - exit multiple"
    SetF $v "H24" '=H22/Assumptions!$C$33-1'; SetText $v "B24" "Upside vs current price - Gordon growth"
    SetF $v "H25" '=H10/OperatingModel!$G$6'; SetText $v "B25" "Implied terminal EV/EBITDA from Gordon TV"
    SetFmt $v "H16:H20" "#,##0.0"; SetFmt $v "H21:H22" '$#,##0.00'; SetFmt $v "H23:H24" "+0.0%;-0.0%"; SetFmt $v "H25" '0.0"x"'
    $v.Range("H16:H20").Borders.Item(9).LineStyle = 1
    SetFont $v "B18,B21:B25" 40 40 40 $false 10; SetFont $v "H22:H25" 0 $false 10
    SubHeader $v "B27:G27"; SetText $v "B27" "  Check: sum of PV + PV(TV) equals enterprise value"
    SetF $v "H27" '=IF(ABS(H8+H14-H16)<0.001,"OK","ERROR")'; SetFont $v "H27" (rgb 0 128 96) $true 10
    Freeze $xl $v "C5"

    # ============ SENSITIVITY ============
    $s = $wb.Worksheets.Item("Sensitivity")
    $s.Tab.Color = rgb 197 90 17
    ColW $s "A" 2; ColW $s "B" 12
    foreach ($l in @("C", "D", "E", "F", "G", "H", "I")) { ColW $s $l 12.5 }
    SetText $s "B2" "SENSITIVITY - IMPLIED SHARE PRICE (USD)"; SetFont $s "B2" $NAVY $true 14
    $waccs = @(0.075, 0.08, 0.085, 0.09, 0.095, 0.10)
    $gs    = @(0.015, 0.02, 0.025, 0.03, 0.035)
    $mults = @(7.0, 7.5, 8.0, 8.5, 9.0, 9.5, 10.0)

    SubHeader $s "B4:I4"; SetText $s "B4" "  Grid A - Gordon-growth terminal value (WACC rows x g columns)"
    SetText $s "B5" "WACC \ g"; SetFont $s "B5" $NAVY $true 10
    for ($j = 0; $j -lt $gs.Count; $j++) {
        $l = [char](67 + $j); SetText $s ([string]$l + "5") $gs[$j]; SetFmt $s ([string]$l + "5") "0.0%"; SetFont $s ([string]$l + "5") $NAVY $true 10
    }
    for ($r = 0; $r -lt $waccs.Count; $r++) {
        $row = 6 + $r
        SetText $s ("B" + $row) $waccs[$r]; SetFmt $s ("B" + $row) "0.0%"; SetFont $s ("B" + $row) $NAVY $true 10
        for ($j = 0; $j -lt $gs.Count; $j++) {
            $l = [char](67 + $j)
            $f = '=(SUMPRODUCT(DCFValuation!$C$5:$G$5,1/(1+$B' + $row + ')^DCFValuation!$C$6:$G$6)+DCFValuation!$G$5*(1+' + $l + '$5)/(($B' + $row + '-' + $l + '$5)*(1+$B' + $row + ')^DCFValuation!$G$6)-Assumptions!$C$31)/Assumptions!$C$32'
            SetF $s ([string]$l + $row) $f
        }
    }
    SetFmt $s "C6:G11" '$#,##0.00'
    try {
        $cs = $s.Range("C6:G11").FormatConditions.AddColorScale(5)
        $cs.ColorScaleCriteria.Item(1).FormatColor.Color = rgb 210 228 245
        $cs.ColorScaleCriteria.Item(3).FormatColor.Color = rgb 255 255 255
        $cs.ColorScaleCriteria.Item(5).FormatColor.Color = rgb 26 179 148
    } catch { }

    SubHeader $s "B13:I13"; SetText $s "B13" "  Grid B - Exit-multiple terminal value (WACC rows x exit EV/EBITDA columns)"
    SetText $s "B14" "WACC \ exit"; SetFont $s "B14" $NAVY $true 10
    for ($j = 0; $j -lt $mults.Count; $j++) {
        $l = [char](67 + $j); SetText $s ([string]$l + "14") $mults[$j]; SetFmt $s ([string]$l + "14") '0.0"x"'; SetFont $s ([string]$l + "14") $NAVY $true 10
    }
    for ($r = 0; $r -lt $waccs.Count; $r++) {
        $row = 15 + $r
        SetText $s ("B" + $row) $waccs[$r]; SetFmt $s ("B" + $row) "0.0%"; SetFont $s ("B" + $row) $NAVY $true 10
        for ($j = 0; $j -lt $mults.Count; $j++) {
            $l = [char](67 + $j)
            $f = '=(SUMPRODUCT(DCFValuation!$C$5:$G$5,1/(1+$B' + $row + ')^DCFValuation!$C$6:$G$6)+' + $l + '$14*OperatingModel!$G$6/(1+$B' + $row + ')^DCFValuation!$G$6-Assumptions!$C$31)/Assumptions!$C$32'
            SetF $s ([string]$l + $row) $f
        }
    }
    SetFmt $s "C15:I20" '$#,##0.00'
    try {
        $cs2 = $s.Range("C15:I20").FormatConditions.AddColorScale(5)
        $cs2.ColorScaleCriteria.Item(1).FormatColor.Color = rgb 210 228 245
        $cs2.ColorScaleCriteria.Item(3).FormatColor.Color = rgb 255 255 255
        $cs2.ColorScaleCriteria.Item(5).FormatColor.Color = rgb 26 179 148
    } catch { }
    SetText $s "B22" "All cells are live formulas recomputing the full DCF at each WACC / terminal assumption."
    SetFont $s "B22" $GRAYT $false 10

    # ============ CALCULATE + VERIFY ============
    $xl.CalculateFullRebuild()
    function Assert([string]$name, [double]$exp, [double]$act, [double]$tol = 0.0005) {
        if ([Math]::Abs($exp - $act) -gt $tol) { throw ("VERIFY FAIL {0}: expected {1}, Excel says {2}" -f $name, $exp, $act) }
        Write-Output ("ok   {0,-46} {1,14:N6}" -f $name, $act)
    }
    $v = $wb.Worksheets.Item("DCFValuation")
    Assert "sum PV of FCF (H8)"          $sumPV  $v.Range("H8").Value2
    Assert "TV Gordon (H10)"             $tvG    $v.Range("H10").Value2
    Assert "TV exit multiple (H11)"      $tvX    $v.Range("H11").Value2
    Assert "PV of TV Gordon (H13)"       $pvTvG  $v.Range("H13").Value2
    Assert "PV of TV exit (H14)"         $pvTvX  $v.Range("H14").Value2
    Assert "EV exit (H16)"               $evX    $v.Range("H16").Value2
    Assert "EV Gordon (H17)"             $evG    $v.Range("H17").Value2
    Assert "equity exit (H19)"           $eqX    $v.Range("H19").Value2
    Assert "price exit (H21)"            $pxX    $v.Range("H21").Value2  0.000005
    Assert "price Gordon (H22)"          $pxG    $v.Range("H22").Value2  0.000005
    Assert "upside exit (H23)"           $upX    $v.Range("H23").Value2  0.00001
    Assert "implied mult from Gordon TV (H25)" $impliedXofG $v.Range("H25").Value2 0.0005
    $s = $wb.Worksheets.Item("Sensitivity")
    Assert "grid A center (E9) = Gordon price"   $pxG                     $s.Range("E9").Value2  0.000005
    Assert "grid A corner (C6)"                  (PriceGordon 0.075 0.015) $s.Range("C6").Value2 0.000005
    Assert "grid B center (G18) = exit price"    $pxX                     $s.Range("G18").Value2 0.000005
    Assert "grid B corner (I20)"                 (PriceExit 0.10 10.0)    $s.Range("I20").Value2 0.000005
    $chk = $v.Range("H27").Text; if ($chk -ne "OK") { throw ("check cell says: " + $chk) }
    Write-Output "EXCEL CHECK CELL: $chk"
    Write-Output ("EXCEL uFCF row: " + ((0..4 | ForEach-Object { [Math]::Round($wb.Worksheets.Item("OperatingModel").Range([string][char](67 + $_) + "15").Value2, 3) }) -join ", "))
    Write-Output ("EXCEL prices: exit={0:N2} gordon={1:N2}" -f $v.Range("H21").Value2, $v.Range("H22").Value2)

    # ============ SAVE ============
    if (Test-Path (Join-Path $OutDir "DCF_Model.xlsx")) { Remove-Item (Join-Path $OutDir "DCF_Model.xlsx") -Force }
    $wb.SaveAs((Join-Path $OutDir "DCF_Model.xlsx"), 51)
    $wb.ExportAsFixedFormat(0, (Join-Path $OutDir "DCF_Model.pdf"))
    $wb.Close($false)
    Write-Output "SAVED DCF_Model.xlsx + DCF_Model.pdf"
}
finally {
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
