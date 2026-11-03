param(
    [string]$OutDir = "C:\Users\tooth\AppData\Local\Temp\opencode\modelling\LBO-Model"
)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ============================================================
# PROJECT LIGHTYEAR - Leveraged Buyout model builder
# Excel COM: live formulas, then Excel calculates and we verify
# every key output against this independent mirror.
# Interest on beginning-of-period balances (no circularity).
# All figures USD millions unless per share. Fictitious data.
# ============================================================

function rgb([int]$r, [int]$g, [int]$b) { return [int]($r + 256 * $g + 65536 * $b) }
$NAVY   = rgb 31 78 121
$LIGHT  = rgb 234 242 251
$BLUEF  = rgb 0 0 255      # input font
$GREENF = rgb 0 128 0      # cross-sheet link font
$GRAYT  = rgb 84 101 121

# ---------- mirror math (ground truth) ----------
$ebitda0 = 500.0
$entryMult = 8.0; $feePct = 0.03; $holdYears = 5
$tlbLev = 4.4; $tlbAmt = $tlbLev * $ebitda0; $tlbRate = 0.09; $tlbAmortPct = 0.05
$ntsLev = 1.2; $ntsAmt = $ntsLev * $ebitda0; $ntsRate = 0.11
$sweepPct = 1.00
$rev0 = 1250.0; $revG = 0.05; $ebM = 0.40; $daP = 0.04; $capP = 0.04; $nwcP = 0.10; $tax = 0.25
$exitMult = 8.0

$ev = $entryMult * $ebitda0; $fees = $ev * $feePct; $uses = $ev + $fees
$eq0 = $uses - $tlbAmt - $ntsAmt
$rev = @(); $prev = $rev0
for ($t = 1; $t -le $holdYears; $t++) { $prev = $prev * (1.0 + $revG); $rev += $prev }
$ebitda = @(0..($holdYears - 1) | ForEach-Object { $rev[$_] * $ebM })
$da = @(0..($holdYears - 1) | ForEach-Object { $rev[$_] * $daP })
$ebit = @(0..($holdYears - 1) | ForEach-Object { $ebitda[$_] - $da[$_] })
$tlbBeg = @(); $tlbEnd = @(); $mand = @(); $sweep = @(); $int = @(); $tTax = @(); $fcf = @(); $cashEnd = @(); $netDebt = @()
$tb = $tlbAmt; $cb = 0.0
$interestNts = $ntsAmt * $ntsRate
for ($t = 0; $t -lt $holdYears; $t++) {
    $tb0 = $tb; $cb0 = $cb
    $interest = $tb0 * $tlbRate + $interestNts
    $tt = [Math]::Max(0.0, $ebit[$t] - $interest) * $tax
    $capex = $rev[$t] * $capP
    $dnwc = if ($t -eq 0) { ($rev[0] - $rev0) * $nwcP } else { ($rev[$t] - $rev[$t - 1]) * $nwcP }
    $f = $ebitda[$t] - $interest - $tt - $capex - $dnwc
    $m = [Math]::Min($tlbAmt * $tlbAmortPct, $tb0)
    $sw = [Math]::Max(0.0, [Math]::Min($sweepPct * ($f - $m), $tb0 - $m))
    $te = $tb0 - $m - $sw
    $ce = $cb0 + $f - $m - $sw
    $tlbBeg += $tb0; $mand += $m; $sweep += $sw; $int += $interest; $tTax += $tt; $fcf += $f
    $tlbEnd += $te; $cashEnd += $ce
    $netDebt += ($te + $ntsAmt - $ce)
    $tb = $te; $cb = $ce
}
$exitEv = $exitMult * $ebitda[$holdYears - 1]
$exitEq = $exitEv - $netDebt[$holdYears - 1]
$moic = $exitEq / $eq0
$irr = [Math]::Pow($moic, 1.0 / $holdYears) - 1.0
$paydown = $tlbAmt + $ntsAmt - $netDebt[$holdYears - 1]
Write-Output ("MIRROR  equity0={0:N3} (uses={1:N3})  EBITDA path: {2}" -f $eq0, $uses, (($ebitda | ForEach-Object { [Math]::Round($_, 2) }) -join ", "))
Write-Output ("MIRROR  FCF: {0}" -f (($fcf | ForEach-Object { [Math]::Round($_, 3) }) -join ", "))
Write-Output ("MIRROR  TLB end: {0}" -f (($tlbEnd | ForEach-Object { [Math]::Round($_, 3) }) -join ", "))
Write-Output ("MIRROR  exitEV={0:N3} exitNetDebt={1:N3} exitEq={2:N3}  MOIC={3:N4}x  IRR={4:P4}  paydown={5:N3}" -f $exitEv, $netDebt[$holdYears - 1], $exitEq, $moic, $irr, $paydown)
Write-Output ("MIRROR  net leverage: {0}" -f ((0..($holdYears - 1) | ForEach-Object { [Math]::Round($netDebt[$_] / $ebitda[$_], 2) }) -join ", "))

function LboIrr([double]$eMult, [double]$xMult) {
    $e0 = $eMult * $ebitda0 * (1.0 + $feePct) - $tlbAmt - $ntsAmt
    $xq = $xMult * $ebitda[$holdYears - 1] - $netDebt[$holdYears - 1]
    return ([Math]::Pow($xq / $e0, 1.0 / $holdYears) - 1.0)
}
function LboMoic([double]$eMult, [double]$xMult) {
    $e0 = $eMult * $ebitda0 * (1.0 + $feePct) - $tlbAmt - $ntsAmt
    $xq = $xMult * $ebitda[$holdYears - 1] - $netDebt[$holdYears - 1]
    return ($xq / $e0)
}

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

# Row map for DebtSchedule:
# 18 sub TLB | 19 beg | 20 mandatory | 21 sweep | 22 end
# 23 sub notes | 24 beg | 25 end
# 26 sub cash | 27 beg | 28 build | 29 end
# 30 sub stats | 31 total debt | 32 net debt | 33 net leverage | 34 coverage

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$xl.AskToUpdateLinks = $false
try {
    $wb = $xl.Workbooks.Add()
    $wb.Worksheets.Item(1).Name = "Cover"
    foreach ($n in @("Assumptions", "SourcesUses", "DebtSchedule", "Returns", "Sensitivity")) {
        $s = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
        $s.Name = $n
    }

    # ============ COVER ============
    $c = $wb.Worksheets.Item("Cover")
    $c.Tab.Color = $NAVY
    ColW $c "A" 3; ColW $c "B" 62; ColW $c "C" 20
    SetText $c "B2" "LEVERAGED BUYOUT MODEL - PROJECT LIGHTYEAR"
    SetFont $c "B2" $NAVY $true 20 "Calibri"
    SetText $c "B3" "Target: WidgetCo Industrials, Inc. | Sources & uses, debt schedule with cash sweep, sponsor returns"
    SetFont $c "B3" $GRAYT $false 11
    SetText $c "B5" "Contents"
    SetFont $c "B5" $NAVY $true 12
    SetText $c "B6" "Assumptions     Transaction terms, financing structure, operating drivers, exit"
    SetText $c "B7" "SourcesUses     Uses and sources at close with the sponsor equity plug and a balance check"
    SetText $c "B8" "DebtSchedule    Operating FCF, mandatory amortisation + 100 percent cash sweep, credit statistics"
    SetText $c "B9" "Returns         Exit equity, MOIC and IRR over the 5-year hold, with an IRR cross-check"
    SetText $c "B10" "Sensitivity     IRR and MOIC grids across entry and exit EV/EBITDA multiples (exact recomputation)"
    SetFont $c "B6:B10" (rgb 40 40 40) $false 10
    SetText $c "B12" "Conventions: blue font = input (change these); black = formula; green = link from another tab. All formulas are live and auditable."
    SetFont $c "B12" $NAVY $false 10
    SetText $c "B13" "Interest is charged on beginning-of-period balances (no circular references, slightly conservative)."
    SetFont $c "B13" $NAVY $false 10
    SetText $c "B14" "Units: USD millions unless per share. Illustrative training model with fictitious data - not investment advice."
    SetFont $c "B14" $GRAYT $false 10
    SetText $c "B16" ("Prepared {0} by aminsaifanoa-web" -f (Get-Date -Format "d MMM yyyy"))
    SetFont $c "B16" $GRAYT $false 10

    # ============ ASSUMPTIONS ============
    $a = $wb.Worksheets.Item("Assumptions")
    $a.Tab.Color = rgb 191 143 34
    ColW $a "A" 2; ColW $a "B" 40
    foreach ($l in @("C", "D", "E", "F", "G")) { ColW $a $l 12.5 }
    SetText $a "B2" "TRANSACTION ASSUMPTIONS"; SetFont $a "B2" $NAVY $true 14
    SetText $a "B3" "Blue = input"; SetFont $a "B3" $BLUEF $false 9

    SectionHeader $a "B5:G5"; SetText $a "B5" "  TRANSACTION"
    SetText $a "B6" "LTM EBITDA (2026A)";           SetText $a "C6"  $ebitda0
    SetText $a "B7" "Entry EV / EBITDA";           SetText $a "C7"  $entryMult
    SetText $a "B8" "Transaction fees (% of EV)";  SetText $a "C8"  $feePct
    SetText $a "B9" "Holding period (years)";      SetText $a "C9"  $holdYears
    SetFont $a "C6" $BLUEF $true 10; SetFmt $a "C6" "#,##0.0"
    SetFont $a "C7" $BLUEF $true 10; SetFmt $a "C7" "0.0x"
    SetFont $a "C8" $BLUEF $true 10; SetFmt $a "C8" "0.0%"
    SetFont $a "C9" $BLUEF $true 10; SetFmt $a "C9" "0"
    SetFont $a "B6:B9" 40 40 40 $false 10

    SectionHeader $a "B11:G11"; SetText $a "B11" "  FINANCING"
    SetText $a "B12" "Term Loan B - leverage (x EBITDA)"; SetText $a "C12" $tlbLev
    SetText $a "B13" "Term Loan B - amount";              SetF   $a "C13" "=C12*C6"
    SetText $a "B14" "Term Loan B - interest rate";       SetText $a "C14" $tlbRate
    SetText $a "B15" "Term Loan B - mandatory amortisation (% original, annual)"; SetText $a "C15" $tlbAmortPct
    SetText $a "B16" "Senior notes - leverage (x EBITDA)"; SetText $a "C16" $ntsLev
    SetText $a "B17" "Senior notes - amount";              SetF   $a "C17" "=C16*C6"
    SetText $a "B18" "Senior notes - interest rate";       SetText $a "C18" $ntsRate
    SetText $a "B19" "Cash sweep (% of FCF after amortisation)"; SetText $a "C19" $sweepPct
    SetFont $a "C12" $BLUEF $true 10; SetFmt $a "C12" "0.0x"
    SetFmt $a "C13" "#,##0.0"; SetFont $a "C13" 0 $true 10
    SetFont $a "C14" $BLUEF $true 10; SetFmt $a "C14" "0.0%"
    SetFont $a "C15" $BLUEF $true 10; SetFmt $a "C15" "0.0%"
    SetFont $a "C16" $BLUEF $true 10; SetFmt $a "C16" "0.0x"
    SetFmt $a "C17" "#,##0.0"; SetFont $a "C17" 0 $true 10
    SetFont $a "C18" $BLUEF $true 10; SetFmt $a "C18" "0.0%"
    SetFont $a "C19" $BLUEF $true 10; SetFmt $a "C19" "0%"
    SetFont $a "B12:B19" 40 40 40 $false 10

    SectionHeader $a "B21:G21"; SetText $a "B21" "  OPERATING DRIVERS"
    SetText $a "B22" "LTM revenue (2026A)";      SetText $a "C22" $rev0
    SetText $a "B23" "Revenue growth (annual)";  SetText $a "C23" $revG
    SetText $a "B24" "EBITDA margin";            SetText $a "C24" $ebM
    SetText $a "B25" "D&A (% of revenue)";       SetText $a "C25" $daP
    SetText $a "B26" "CapEx (% of revenue)";     SetText $a "C26" $capP
    SetText $a "B27" "Change in NWC (% of revenue change)"; SetText $a "C27" $nwcP
    SetText $a "B28" "Cash tax rate";            SetText $a "C28" $tax
    SetFont $a "C22" $BLUEF $true 10; SetFmt $a "C22" "#,##0.0"
    foreach ($rr in 23..27) { SetFont $a ("C" + $rr) $BLUEF $true 10; SetFmt $a ("C" + $rr) "0.0%" }
    SetFont $a "C28" $BLUEF $true 10; SetFmt $a "C28" "0.0%"
    SetFont $a "B22:B28" 40 40 40 $false 10

    SectionHeader $a "B30:G30"; SetText $a "B30" "  EXIT"
    SetText $a "B31" "Exit EV / EBITDA";         SetText $a "C31" $exitMult
    SetFont $a "C31" $BLUEF $true 10; SetFmt $a "C31" "0.0x"
    SetFont $a "B31" 40 40 40 $false 10

    # ============ SOURCES & USES ============
    $u = $wb.Worksheets.Item("SourcesUses")
    $u.Tab.Color = rgb 84 101 121
    ColW $u "A" 2; ColW $u "B" 36; ColW $u "C" 13; ColW $u "D" 13
    SetText $u "B2" "SOURCES & USES AT CLOSE (USD mm)"; SetFont $u "B2" $NAVY $true 14

    SectionHeader $u "B4:D4"; SetText $u "B4" "  USES"
    SetText $u "B5" "Purchase of enterprise (EV)"; SetF $u "C5" "=Assumptions!C6*Assumptions!C7"
    SetText $u "B6" "Transaction fees";            SetF $u "C6" "=C5*Assumptions!C8"
    SetText $u "B7" "Total uses";                  SetF $u "C7" "=C5+C6"
    SetFmt $u "C5:C7" "#,##0.0"; SetFont $u "C5:C6" $GREENF $false 10
    SetFont $u "B5:B6" 40 40 40 $false 10; SetFont $u "B7" $NAVY $true 10
    $u.Range("C7").Borders.Item(9).LineStyle = 1

    SectionHeader $u "B9:D9"; SetText $u "B9" "  SOURCES"
    SetText $u "B10" "Term Loan B";           SetF $u "C10" "=Assumptions!C13"
    SetText $u "B11" "Senior notes";          SetF $u "C11" "=Assumptions!C17"
    SetText $u "B12" "Sponsor equity (plug)"; SetF $u "C12" "=C7-C10-C11"
    SetText $u "B13" "Total sources";         SetF $u "C13" "=C10+C11+C12"
    SetFmt $u "C10:C13" "#,##0.0"; SetFont $u "C10:C11" $GREENF $false 10
    SetFont $u "B10:B11" 40 40 40 $false 10; SetFont $u "B12" $NAVY $true 10; SetFont $u "B13" $NAVY $true 10
    $u.Range("C13").Borders.Item(9).LineStyle = 1

    SectionHeader $u "B15:D15"; SetText $u "B15" "  % OF TOTAL FUNDING"
    SetText $u "B16" "Term Loan B";    SetF $u "D16" "=C10/C13"; SetFmt $u "D16" "0.0%"
    SetText $u "B17" "Senior notes";   SetF $u "D17" "=C11/C13"; SetFmt $u "D17" "0.0%"
    SetText $u "B18" "Sponsor equity"; SetF $u "D18" "=C12/C13"; SetFmt $u "D18" "0.0%"
    SetFont $u "B16:B18" 40 40 40 $false 10
    SubHeader $u "B20:D20"; SetText $u "B20" "  Entry statistics"
    SetText $u "B21" "Gross leverage (total debt / EBITDA)"; SetF $u "C21" "=(C10+C11)/Assumptions!C6"; SetFmt $u "C21" "0.0x"
    SetText $u "B22" "Sponsor equity % of capitalisation";   SetF $u "C22" "=C12/C13"; SetFmt $u "C22" "0.0%"
    SetFont $u "B21:B22" 40 40 40 $false 10

    # ============ DEBT SCHEDULE ============
    $d = $wb.Worksheets.Item("DebtSchedule")
    $d.Tab.Color = rgb 0 128 96
    ColW $d "A" 2; ColW $d "B" 42
    foreach ($l in @("C", "D", "E", "F", "G")) { ColW $d $l 12.5 }
    SetText $d "B2" "DEBT SCHEDULE & FREE CASH FLOW (USD mm)"; SetFont $d "B2" $NAVY $true 14
    SetText $d "B4" "Forecast year"; SetFont $d "B4" $NAVY $true 10
    for ($i = 1; $i -le 5; $i++) {
        $l = [char](66 + $i)
        SetText $d ([string]$l + "4") ("Year " + $i)
        SetFont $d ([string]$l + "4") $NAVY $true 10
        $d.Range([string]$l + "4").Borders.Item(9).LineStyle = 1
    }

    SubHeader $d "B5:G5"; SetText $d "B5" "  Operating results"
    SetText $d "B6" "Revenue"
    SetF $d "C6" "=Assumptions!C22*(1+Assumptions!C23)"
    for ($i = 1; $i -le 4; $i++) {
        $prev = [char](66 + $i); $cur = [char](67 + $i)
        SetF $d ([string]$cur + "6") ("=" + $prev + "6*(1+Assumptions!C23)")
    }
    SetFont $d "C6:G6" $GREENF $false 10
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "7") ("=" + $l + "6*Assumptions!C24") }
    SetText $d "B7" "EBITDA"; SetFont $d "B7" $NAVY $true 10
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "8") ("=" + $l + "6*Assumptions!C25") }
    SetText $d "B8" "Depreciation & amortisation"
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "9") ("=" + $l + "7-" + $l + "8") }
    SetText $d "B9" "EBIT"; SetFont $d "B9" $NAVY $true 10
    SetFmt $d "C6:G9" "#,##0.0"; SetFont $d "B6:B9" 40 40 40 $false 10

    SubHeader $d "B11:G11"; SetText $d "B11" "  Levered free cash flow"
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "12") ("=-" + $l + "19*Assumptions!C14") }
    SetText $d "B12" "Less: cash interest - Term Loan B"
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "13") ("=-" + $l + "24*Assumptions!C18") }
    SetText $d "B13" "Less: cash interest - senior notes"
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "14") ("=-Assumptions!C28*MAX(0," + $l + "9+" + $l + "12+" + $l + "13)") }
    SetText $d "B14" "Less: cash taxes"
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "15") ("=-" + $l + "6*Assumptions!C26") }
    SetText $d "B15" "Less: capital expenditure"
    SetF $d "C16" "=-(C6-Assumptions!C22)*Assumptions!C27"
    for ($i = 1; $i -le 4; $i++) {
        $prev = [char](66 + $i); $cur = [char](67 + $i)
        SetF $d ([string]$cur + "16") ("=-(" + $cur + "6-" + $prev + "6)*Assumptions!C27")
    }
    SetText $d "B16" "Less: increase in net working capital"
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "17") ("=" + $l + "7+" + $l + "12+" + $l + "13+" + $l + "14+" + $l + "15+" + $l + "16") }
    SetText $d "B17" "Free cash flow available for debt paydown"
    SetFmt $d "C12:G17" "#,##0.0"
    $d.Range("C17:G17").Borders.Item(9).LineStyle = 1
    SetFont $d "B12:B17" 40 40 40 $false 10; SetFont $d "B17" $NAVY $true 10

    SubHeader $d "B18:G18"; SetText $d "B18" "  Term Loan B"
    SetF $d "C19" "=SourcesUses!C10"
    for ($i = 1; $i -le 4; $i++) {
        $prev = [char](66 + $i); $cur = [char](67 + $i)
        SetF $d ([string]$cur + "19") ("=" + $prev + "22")
    }
    SetText $d "B19" "Beginning balance"; SetFont $d "C19" $GREENF $false 10
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "20") ("=-MIN(Assumptions!C13*Assumptions!C15," + $l + "19)") }
    SetText $d "B20" "Mandatory amortisation"
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "21") ("=-MAX(0,MIN(Assumptions!C19*(" + $l + "17-ABS(" + $l + "20))," + $l + "19-ABS(" + $l + "20)))") }
    SetText $d "B21" "Cash sweep"
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "22") ("=" + $l + "19+" + $l + "20+" + $l + "21") }
    SetText $d "B22" "Ending balance"; SetFont $d "B22" $NAVY $true 10
    SetFmt $d "C19:G22" "#,##0.0"

    SubHeader $d "B23:G23"; SetText $d "B23" "  Senior notes (bullet, non-amortising)"
    SetF $d "C24" "=SourcesUses!C11"
    for ($i = 1; $i -le 4; $i++) {
        $prev = [char](66 + $i); $cur = [char](67 + $i)
        SetF $d ([string]$cur + "24") ("=" + $prev + "25")
    }
    SetText $d "B24" "Beginning balance"; SetFont $d "C24" $GREENF $false 10
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "25") ("=" + $l + "24") }
    SetText $d "B25" "Ending balance"; SetFmt $d "C24:G25" "#,##0.0"; SetFont $d "B24:B25" 40 40 40 $false 10

    SubHeader $d "B26:G26"; SetText $d "B26" "  Cash on balance sheet"
    SetText $d "B27" "Beginning cash"; SetText $d "C27" 0.0; SetFmt $d "C27" "#,##0.0"
    for ($i = 1; $i -le 4; $i++) {
        $prev = [char](66 + $i); $cur = [char](67 + $i)
        SetF $d ([string]$cur + "27") ("=" + $prev + "29")
    }
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "28") ("=" + $l + "17+" + $l + "20+" + $l + "21") }
    SetText $d "B28" "Build-up (FCF less total repayment)"
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "29") ("=" + $l + "27+" + $l + "28") }
    SetText $d "B29" "Ending cash"
    SetFmt $d "C27:G29" "#,##0.0"; SetFont $d "B27:B29" 40 40 40 $false 10

    SubHeader $d "B30:G30"; SetText $d "B30" "  Credit statistics (year-end)"
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "31") ("=" + $l + "22+" + $l + "25") }
    SetText $d "B31" "Total debt"
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "32") ("=" + $l + "31-" + $l + "29") }
    SetText $d "B32" "Net debt"; SetFont $d "B32" $NAVY $true 10
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "33") ("=" + $l + "32/" + $l + "7") }
    SetText $d "B33" "Net leverage (x)"; SetFont $d "B33" $NAVY $true 10; SetFmt $d "C33:G33" "0.0x"
    for ($i = 0; $i -le 4; $i++) { $l = [char](67 + $i); SetF $d ([string]$l + "34") ("=" + $l + "7/ABS(" + $l + "12+" + $l + "13)") }
    SetText $d "B34" "Interest coverage (EBITDA / cash interest)"; SetFmt $d "C34:G34" "0.0x"
    SetFmt $d "C31:G32" "#,##0.0"
    SetFont $d "B31:B34" 40 40 40 $false 10
    Freeze $xl $d "C6"

    # ============ RETURNS ============
    $r2 = $wb.Worksheets.Item("Returns")
    $r2.Tab.Color = rgb 197 90 17
    ColW $r2 "A" 2; ColW $r2 "B" 46; ColW $r2 "C" 13
    SetText $r2 "B2" "SPONSOR RETURNS (USD mm)"; SetFont $r2 "B2" $NAVY $true 14
    SubHeader $r2 "B4:C4"; SetText $r2 "B4" "  Entry"
    SetText $r2 "B5" "Sponsor equity invested"; SetF $r2 "C5" "=SourcesUses!C12"
    SetFmt $r2 "C5" "#,##0.0"; SetFont $r2 "C5" $GREENF $false 10; SetFont $r2 "B5" 40 40 40 $false 10
    SubHeader $r2 "B7:C7"; SetText $r2 "B7" "  Exit (year 5)"
    SetText $r2 "B8" "Exit enterprise value";   SetF $r2 "C8" "=Assumptions!C31*DebtSchedule!G7"
    SetText $r2 "B9" "Less: net debt at exit";  SetF $r2 "C9" "=-DebtSchedule!G32"
    SetText $r2 "B10" "Exit equity value";      SetF $r2 "C10" "=C8+C9"
    SetFmt $r2 "C8:C10" "#,##0.0"; SetFont $r2 "B10" $NAVY $true 10
    $r2.Range("C10").Borders.Item(9).LineStyle = 1
    SetFont $r2 "B5,B8:B9" 40 40 40 $false 10
    SubHeader $r2 "B12:C12"; SetText $r2 "B12" "  Returns over 5-year hold"
    SetText $r2 "B13" "Profit";                 SetF $r2 "C13" "=C10-C5"
    SetText $r2 "B14" "MOIC (multiple of invested capital)"; SetF $r2 "C14" "=C10/C5"
    SetText $r2 "B15" "IRR";                    SetF $r2 "C15" "=(C10/C5)^(1/Assumptions!C9)-1"
    SetFmt $r2 "C13" "#,##0.0"; SetFmt $r2 "C14" "0.00x"; SetFmt $r2 "C15" "0.0%"
    SetFont $r2 "B14,B15" $NAVY $true 10; SetFont $r2 "C14,C15" $NAVY $true 11
    SubHeader $r2 "B17:C17"; SetText $r2 "B17" "  Cross-checks and statistics"
    SetText $r2 "B18" "Check: (1+IRR)^hold equals MOIC"
    SetF $r2 "C18" '=IF(ABS((1+C15)^Assumptions!C9-C14)<0.0001,"OK","ERROR")'
    SetText $r2 "B19" "Average net leverage during hold (x)"
    SetF $r2 "C19" "=AVERAGE(DebtSchedule!C33:G33)"; SetFmt $r2 "C19" "0.0x"
    SetText $r2 "B20" "Total net debt paydown over hold"
    SetF $r2 "C20" "=DebtSchedule!C19+DebtSchedule!C24-DebtSchedule!G32"
    SetFmt $r2 "C20" "#,##0.0"
    SetText $r2 "B21" "EBITDA growth over hold"
    SetF $r2 "C21" "=DebtSchedule!G7/Assumptions!C6-1"; SetFmt $r2 "C21" "0.0%"
    SetFont $r2 "B18:B21" 40 40 40 $false 10

    # ============ SENSITIVITY ============
    $s = $wb.Worksheets.Item("Sensitivity")
    $s.Tab.Color = rgb 197 90 17
    ColW $s "A" 2; ColW $s "B" 13
    foreach ($l in @("C", "D", "E", "F", "G", "H", "I", "J", "K")) { ColW $s $l 11 }
    SetText $s "B2" "SENSITIVITY - RETURN GRIDS"; SetFont $s "B2" $NAVY $true 14
    $axes = @(6.0, 6.5, 7.0, 7.5, 8.0, 8.5, 9.0, 9.5, 10.0)

    SubHeader $s "B4:K4"; SetText $s "B4" "  Grid A - IRR by entry multiple (rows) x exit multiple (columns)"
    SetText $s "B5" "entry \ exit"; SetFont $s "B5" $NAVY $true 10
    for ($j = 0; $j -lt $axes.Count; $j++) {
        $l = [char](67 + $j)
        SetText $s ([string]$l + "5") $axes[$j]
        SetFmt $s ([string]$l + "5") "0.0x"
        SetFont $s ([string]$l + "5") $NAVY $true 10
    }
    for ($i = 0; $i -lt $axes.Count; $i++) {
        $row = 6 + $i
        SetText $s ("B" + $row) $axes[$i]
        SetFmt $s ("B" + $row) "0.0x"
        SetFont $s ("B" + $row) $NAVY $true 10
        for ($j = 0; $j -lt $axes.Count; $j++) {
            $l = [char](67 + $j)
            $f = '=((' + $l + '$5*DebtSchedule!$G$7-DebtSchedule!$G$32)/(($B' + $row + '*Assumptions!$C$6*(1+Assumptions!$C$8))-(Assumptions!$C$13+Assumptions!$C$17)))^(1/Assumptions!$C$9)-1'
            SetF $s ([string]$l + $row) $f
        }
    }
    SetFmt $s "C6:K14" "0.0%"
    try {
        $cs = $s.Range("C6:K14").FormatConditions.AddColorScale(5)
        $cs.ColorScaleCriteria.Item(1).FormatColor.Color = rgb 244 204 204
        $cs.ColorScaleCriteria.Item(3).FormatColor.Color = rgb 255 249 231
        $cs.ColorScaleCriteria.Item(5).FormatColor.Color = rgb 26 179 148
    } catch { }

    SubHeader $s "B16:K16"; SetText $s "B16" "  Grid B - MOIC by entry multiple (rows) x exit multiple (columns)"
    SetText $s "B17" "entry \ exit"; SetFont $s "B17" $NAVY $true 10
    for ($j = 0; $j -lt $axes.Count; $j++) {
        $l = [char](67 + $j)
        SetText $s ([string]$l + "17") $axes[$j]
        SetFmt $s ([string]$l + "17") "0.0x"
        SetFont $s ([string]$l + "17") $NAVY $true 10
    }
    for ($i = 0; $i -lt $axes.Count; $i++) {
        $row = 18 + $i
        SetText $s ("B" + $row) $axes[$i]
        SetFmt $s ("B" + $row) "0.0x"
        SetFont $s ("B" + $row) $NAVY $true 10
        for ($j = 0; $j -lt $axes.Count; $j++) {
            $l = [char](67 + $j)
            $f = '=((' + $l + '$17*DebtSchedule!$G$7-DebtSchedule!$G$32)/(($B' + $row + '*Assumptions!$C$6*(1+Assumptions!$C$8))-(Assumptions!$C$13+Assumptions!$C$17)))'
            SetF $s ([string]$l + $row) $f
        }
    }
    SetFmt $s "C18:K26" "0.00x"
    try {
        $cs2 = $s.Range("C18:K26").FormatConditions.AddColorScale(5)
        $cs2.ColorScaleCriteria.Item(1).FormatColor.Color = rgb 244 204 204
        $cs2.ColorScaleCriteria.Item(3).FormatColor.Color = rgb 255 249 231
        $cs2.ColorScaleCriteria.Item(5).FormatColor.Color = rgb 26 179 148
    } catch { }
    SetText $s "B28" "Grids recompute the full buyout at each combination of multiples. Because the debt paydown path does not depend on the entry or exit multiples, these grids are exact, not approximations."
    SetFont $s "B28" $GRAYT $false 10

    # ============ CALCULATE + VERIFY ============
    $xl.CalculateFullRebuild()
    function Assert([string]$name, [double]$exp, [double]$act, [double]$tol = 0.0005) {
        if ([Math]::Abs($exp - $act) -gt $tol) { throw ("VERIFY FAIL {0}: expected {1}, Excel says {2}" -f $name, $exp, $act) }
        Write-Output ("ok   {0,-48} {1,14:N6}" -f $name, $act)
    }
    $u2 = $wb.Worksheets.Item("SourcesUses")
    Assert "purchase EV (uses)"          $ev    $u2.Range("C5").Value2
    Assert "total uses"                  $uses  $u2.Range("C7").Value2
    Assert "sponsor equity plug"         $eq0   $u2.Range("C12").Value2
    Assert "total sources = total uses"  $uses  $u2.Range("C13").Value2
    $d2 = $wb.Worksheets.Item("DebtSchedule")
    for ($t = 0; $t -lt 5; $t++) {
        $l = [char](67 + $t)
        Assert ("FCF year " + ($t + 1))            $fcf[$t]    $d2.Range([string]$l + "17").Value2
        Assert ("TLB end year " + ($t + 1))        $tlbEnd[$t] $d2.Range([string]$l + "22").Value2
        Assert ("net leverage year " + ($t + 1))   ([double]$netDebt[$t] / [double]$ebitda[$t]) $d2.Range([string]$l + "33").Value2 0.001
    }
    $r3 = $wb.Worksheets.Item("Returns")
    Assert "exit EV"              $exitEv  $r3.Range("C8").Value2
    Assert "exit equity"          $exitEq  $r3.Range("C10").Value2
    Assert "MOIC"                 $moic    $r3.Range("C14").Value2 0.0002
    Assert "IRR"                  $irr     $r3.Range("C15").Value2 0.00002
    Assert "net debt paydown"     $paydown $r3.Range("C20").Value2
    $chk = $r3.Range("C18").Text; if ($chk -ne "OK") { throw ("MOIC/IRR check cell says: " + $chk) }
    Write-Output "EXCEL IRR/MOIC CROSS-CHECK: $chk"
    $s2 = $wb.Worksheets.Item("Sensitivity")
    Assert "IRR grid center (G10, both 8.0x)"     $irr              $s2.Range("G10").Value2 0.00002
    Assert "MOIC grid center (G22, both 8.0x)"    $moic             $s2.Range("G22").Value2 0.0002
    Assert "IRR grid corner C6 (entry 6.0, exit 6.0)"   (LboIrr 6.0 6.0)   $s2.Range("C6").Value2  0.0002
    Assert "IRR grid corner K14 (entry 10.0, exit 10.0)" (LboIrr 10.0 10.0) $s2.Range("K14").Value2 0.0002

    # ============ SAVE ============
    if (Test-Path (Join-Path $OutDir "LBO_Model.xlsx")) { Remove-Item (Join-Path $OutDir "LBO_Model.xlsx") -Force }
    $wb.SaveAs((Join-Path $OutDir "LBO_Model.xlsx"), 51)
    $wb.ExportAsFixedFormat(0, (Join-Path $OutDir "LBO_Model.pdf"))
    $wb.Close($false)
    Write-Output "SAVED LBO_Model.xlsx + LBO_Model.pdf"
}
finally {
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
