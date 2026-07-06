Attribute VB_Name = "DiscountEngine"
Option Explicit

'╔══════════════════════════════════════════════════════════════════════════════╗
'║        DISCOUNT CALCULATION ENGINE  v3.1  —  UNIVERSAL EDITION             ║
'╠══════════════════════════════════════════════════════════════════════════════╣
'║  KEY UPGRADE FROM v3.0:                                                     ║
'║  v3.0 had hardcoded BLR column numbers → broke on CCU sheet                ║
'║  v3.1 reads header row of ACTIVE sheet → auto-detects every column         ║
'║  Works on: BLR-2025-26, CCU_2025-26, any future sheet                      ║
'╠══════════════════════════════════════════════════════════════════════════════╣
'║  WHY CCU BROKE:                                                             ║
'║  CCU has 2 extra cols at position 7,8 (AIRLINE_MATCH, FLIGHT_CLEAN)        ║
'║  This shifts AIRLINES: BLR=col7, CCU=col9                                  ║
'║       AWB OWNED BY: BLR=col10, CCU=col12  (all key cols shift by +2)       ║
'║  Engine now finds "AIRLINES" header wherever it lives                       ║
'╚══════════════════════════════════════════════════════════════════════════════╝



'════════════════════════════════════════════════════════════════════════════════
'  §1  CONFIGURATION
'  Fallback column numbers used ONLY if header auto-detection fails
'  In normal use, engine reads header row and ignores these fallbacks
'════════════════════════════════════════════════════════════════════════════════

'── Discount sheet names ─────────────────────────────────────────────────────
Private Const CFG_DISC_SHEET        As String = "DISCOUNT"
Private Const CFG_DISC_SHEET_ALT    As String = "DISCOUNT (2)"

'── Fallback column numbers (BLR-2025-26 layout) ─────────────────────────────
'   Used only if header auto-detection fails for that field
Private Const FB_DEST               As Long = 5
Private Const FB_AIRLINE            As Long = 7
Private Const FB_CONSIGNOR          As Long = 9
Private Const FB_AWB_OWNER          As Long = 10
Private Const FB_MATERIAL           As Long = 11
Private Const FB_WEIGHT             As Long = 13
Private Const FB_RATE               As Long = 14
Private Const FB_BASIC              As Long = 15
Private Const FB_DUE_AGENT          As Long = 16
Private Const FB_GROSS              As Long = 20
Private Const FB_DISC_OUT           As Long = 21
Private Const FB_TOTAL_GST          As Long = 26
Private Const FB_VALIDATE_UPTO      As Long = 19

'── Data sheet header keywords (partial match, case-insensitive) ──────────────
'   Engine searches for these in row 1 of the active sheet
Private Const KW_DEST               As String = "DEST"
Private Const KW_AIRLINES           As String = "AIRLINES"    ' NOT "AIRLINE_MATCH"
Private Const KW_CONSIGNOR          As String = "CONSIGNOR"
Private Const KW_AWB_OWNER          As String = "AWB OWNED"
Private Const KW_MATERIAL           As String = "MATERIAL"
Private Const KW_WEIGHT             As String = "CH. WT"
Private Const KW_WEIGHT2            As String = "CH.WT"        ' alternate spelling
Private Const KW_RATE               As String = "RATE"
Private Const KW_BASIC              As String = "BASIC"
Private Const KW_DUE_AGENT          As String = "DUE AGENT"
Private Const KW_GROSS              As String = "GROSS"
Private Const KW_DISC_OUT           As String = "DISCOUNT"
Private Const KW_TOTAL_GST          As String = "AMOUNT AFTER GST"
Private Const KW_VALIDATE_END       As String = "OCDC"         ' Last col to validate

'── DISCOUNT sheet header keywords ───────────────────────────────────────────
Private Const DK_OWNER              As String = "AWB OWNED"
Private Const DK_DEST               As String = "DEST"
Private Const DK_CONS               As String = "CONSIGNOR"
Private Const DK_AIR                As String = "AIRLINE"
Private Const DK_RATE               As String = "RATE"
Private Const DK_MAT                As String = "MATERIAL"

'── Matching behaviour ───────────────────────────────────────────────────────
Private Const CFG_AWB_PARTIAL       As Boolean = True
Private Const CFG_GST_RATE          As Double = 0.18

'── Scoring constants ────────────────────────────────────────────────────────
Private Const SC_EXACT              As Long = 5
Private Const SC_WILDCARD           As Long = 1
Private Const SC_CONSIGNOR          As Long = 3
Private Const SC_RATE_RANGE         As Long = 3
Private Const SC_MATERIAL           As Long = 10
Private Const SC_AWB_PARTIAL        As Long = 3
Private Const SC_KILL               As Long = -1



'════════════════════════════════════════════════════════════════════════════════
'  §2  TYPE DEFINITIONS
'════════════════════════════════════════════════════════════════════════════════

'── Dynamically detected column map for any data sheet ───────────────────────
Private Type DataSheetMap
    ColDest         As Long
    ColAirline      As Long
    ColConsignor    As Long
    ColAwbOwner     As Long
    ColMaterial     As Long
    ColWeight       As Long
    ColRate         As Long
    ColBasic        As Long
    ColDueAgent     As Long
    ColGross        As Long
    ColDiscOut      As Long
    ColTotalGST     As Long
    ColValidateUpto As Long    ' Validate cols 1 to this column
    SheetName       As String  ' Which sheet was detected
    DetectedFrom    As String  ' "HEADERS" or "FALLBACK"
End Type

Private Type WeightSlab
    Header          As String
    UpperBound      As Double
    ColIdx          As Long
End Type

Private Type DiscountRule
    AwbOwner        As String
    Dest            As String
    Consignor       As String
    Airline         As String
    RateFilter      As String
    Material        As String
    MinVal          As Double
    SlabVals()      As String
    SlabCount       As Long
    SourceRow       As Long
End Type

Private Type DiscSheetMap
    ColOwner        As Long
    ColDest         As Long
    ColCons         As Long
    ColAir          As Long
    ColRate         As Long
    ColMaterial     As Long
    Slabs()         As WeightSlab
    SlabCount       As Long
    FirstSlabCol    As Long
    FormatStyle     As String
End Type

Private Type RowValues
    AwbOwner        As String
    Dest            As String
    Consignor       As String
    Airline         As String
    Material        As String
    Weight          As Double
    Rate            As Double
    BasicVal        As Double
    DueAgent        As Double
    GrossAmt        As Double
    TotalGST        As Double
End Type

Private Type ProcessResult
    DiscountApplied As Double
    FormulaUsed     As String
    RuleOwner       As String
    MatchScore      As Long
    MinValUsed      As Double
    SkipReason      As String
    WasSkipped      As Boolean
End Type



'════════════════════════════════════════════════════════════════════════════════
'  §3  PUBLIC ENTRY POINTS
'════════════════════════════════════════════════════════════════════════════════

Public Sub CALCULATEDISCOUNT()
    Call RunEngine(False, False)
End Sub

Public Sub CalculateDiscountSelected()
    Call RunEngine(True, False)
End Sub

Public Sub DiscountPreview()
    Call RunEngine(False, True)
End Sub

Public Sub ShowDiagnostics()
    Dim wsData    As Worksheet
    Dim wsDis     As Worksheet
    Dim dm        As DataSheetMap
    Dim dsMap     As DiscSheetMap
    Dim rules()   As DiscountRule
    Dim ruleCount As Long

    Set wsData = ActiveSheet
    If Not GetSheet(wsDis, CFG_DISC_SHEET, CFG_DISC_SHEET_ALT) Then Exit Sub
    If Not MapDataSheetHeaders(wsData, dm) Then Exit Sub
    If Not LoadDiscountRules(wsDis, dsMap, rules, ruleCount) Then Exit Sub

    Dim msg As String
    msg = "DISCOUNT ENGINE v3.1 — Diagnostics" & vbNewLine
    msg = msg & String(45, "-") & vbNewLine
    msg = msg & "Active Sheet      : " & dm.SheetName & vbNewLine
    msg = msg & "Column Detection  : " & dm.DetectedFrom & vbNewLine
    msg = msg & "Discount Sheet    : " & wsDis.Name & vbNewLine
    msg = msg & "Formula Format    : " & dsMap.FormatStyle & vbNewLine
    msg = msg & "Rules Loaded      : " & ruleCount & vbNewLine
    msg = msg & "Weight Slabs      : " & dsMap.SlabCount & vbNewLine
    msg = msg & "AWB Partial Match : " & IIf(CFG_AWB_PARTIAL, "ON", "OFF") & vbNewLine
    msg = msg & vbNewLine

    msg = msg & "── Detected Column Map ──" & vbNewLine
    msg = msg & PadR("DEST", 16) & "= Col " & dm.ColDest & vbNewLine
    msg = msg & PadR("AIRLINES", 16) & "= Col " & dm.ColAirline & vbNewLine
    msg = msg & PadR("CONSIGNOR", 16) & "= Col " & dm.ColConsignor & vbNewLine
    msg = msg & PadR("AWB OWNED BY", 16) & "= Col " & dm.ColAwbOwner & vbNewLine
    msg = msg & PadR("MATERIAL", 16) & "= Col " & dm.ColMaterial & vbNewLine
    msg = msg & PadR("CH.WT", 16) & "= Col " & dm.ColWeight & vbNewLine
    msg = msg & PadR("RATE", 16) & "= Col " & dm.ColRate & vbNewLine
    msg = msg & PadR("BASIC", 16) & "= Col " & dm.ColBasic & vbNewLine
    msg = msg & PadR("DUE AGENT", 16) & "= Col " & dm.ColDueAgent & vbNewLine
    msg = msg & PadR("GROSS", 16) & "= Col " & dm.ColGross & vbNewLine
    msg = msg & PadR("DISCOUNT OUT", 16) & "= Col " & dm.ColDiscOut & vbNewLine
    msg = msg & PadR("AMOUNT AFT GST", 16) & "= Col " & dm.ColTotalGST & vbNewLine
    msg = msg & PadR("VALIDATE UPTO", 16) & "= Col " & dm.ColValidateUpto & vbNewLine
    msg = msg & vbNewLine

    msg = msg & "── Weight Slabs ──" & vbNewLine
    Dim s As Long
    For s = 0 To dsMap.SlabCount - 1
        msg = msg & "  [" & (s + 1) & "] " & dsMap.Slabs(s).Header & _
              " → col " & dsMap.Slabs(s).ColIdx & _
              " (upto " & dsMap.Slabs(s).UpperBound & " kg)" & vbNewLine
    Next s

    MsgBox msg, vbInformation, "Discount Engine v3.1"
End Sub

Public Sub ShowUnmatchedRows()
    Call RunEngine(False, True, True)
End Sub



'════════════════════════════════════════════════════════════════════════════════
'  §4  CORE ENGINE
'════════════════════════════════════════════════════════════════════════════════

Private Sub RunEngine(selectedOnly As Boolean, dryRun As Boolean, _
                      Optional unmatchedOnly As Boolean = False)

    Dim wsData      As Worksheet
    Dim wsDis       As Worksheet
    Dim dm          As DataSheetMap   ' <── Active sheet column map (auto-detected)
    Dim dsMap       As DiscSheetMap
    Dim rules()     As DiscountRule
    Dim ruleCount   As Long
    Dim startRow    As Long
    Dim endRow      As Long
    Dim lastRow     As Long
    Dim dataArr     As Variant
    Dim i           As Long

    Dim cProcessed  As Long
    Dim cSkipped    As Long
    Dim cNoRule     As Long
    Dim cMinApplied As Long
    Dim cFormula    As Long

    Dim unmatchedLog As String
    unmatchedLog = "Row|AWB Owner|Dest|Material|Airline|Reason" & vbNewLine
    unmatchedLog = unmatchedLog & String(90, "-") & vbNewLine

    Set wsData = ActiveSheet

    '── Step 1: Auto-detect columns from active sheet header row ─────────────
    If Not MapDataSheetHeaders(wsData, dm) Then
        MsgBox "Sheet '" & wsData.Name & "' ki header row se columns detect" & _
               " nahi hue. Row 1 mein headers hain?", vbCritical
        Exit Sub
    End If

    '── Step 2: Load DISCOUNT rules ──────────────────────────────────────────
    If Not GetSheet(wsDis, CFG_DISC_SHEET, CFG_DISC_SHEET_ALT) Then Exit Sub
    If Not LoadDiscountRules(wsDis, dsMap, rules, ruleCount) Then Exit Sub
    If ruleCount = 0 Then
        MsgBox "Koi valid discount rule nahi mila!", vbExclamation: Exit Sub
    End If

    '── Step 3: Row range ────────────────────────────────────────────────────
    lastRow = GetLastUsedRow(wsData)
    If selectedOnly Then
        startRow = Selection.Row
        endRow = startRow + Selection.Rows.Count - 1
        If startRow < 2 Then startRow = 2
        If endRow > lastRow Then endRow = lastRow
    Else
        startRow = 2: endRow = lastRow
    End If

    If startRow > endRow Or endRow < 2 Then
        MsgBox "Valid rows nahi mili!": Exit Sub
    End If

    '── Step 4: Performance mode ON ─────────────────────────────────────────
    If Not dryRun Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    End If

    '── Step 5: Array read ───────────────────────────────────────────────────
    Dim maxCol As Long
    maxCol = MaxColNeeded(dm)
    dataArr = wsData.Range( _
                  wsData.Cells(startRow, 1), _
                  wsData.Cells(endRow, maxCol) _
              ).Value

    '── Step 6: Process rows ─────────────────────────────────────────────────
    Dim outArr() As Variant
    ReDim outArr(1 To UBound(dataArr, 1), 1 To 1)

    For i = 1 To UBound(dataArr, 1)
        Dim actualRow As Long
        actualRow = startRow + i - 1

        Dim res As ProcessResult
        res = ProcessSingleRow(dataArr, i, rules, ruleCount, dsMap, dm)

        If res.WasSkipped Then
            cSkipped = cSkipped + 1
            outArr(i, 1) = SafeGet(dataArr, i, dm.ColDiscOut)
        ElseIf res.RuleOwner = "" Then
            cNoRule = cNoRule + 1
            outArr(i, 1) = SafeGet(dataArr, i, dm.ColDiscOut)
            unmatchedLog = unmatchedLog & _
                actualRow & "|" & _
                SafeStr(dataArr, i, dm.ColAwbOwner) & "|" & _
                SafeStr(dataArr, i, dm.ColDest) & "|" & _
                SafeStr(dataArr, i, dm.ColMaterial) & "|" & _
                SafeStr(dataArr, i, dm.ColAirline) & "|" & _
                res.SkipReason & vbNewLine
        Else
            cProcessed = cProcessed + 1
            outArr(i, 1) = res.DiscountApplied
            If res.MinValUsed > 0 And res.DiscountApplied <= res.MinValUsed Then
                cMinApplied = cMinApplied + 1
            Else
                cFormula = cFormula + 1
            End If
        End If
    Next i

    '── Step 7: Batch write ──────────────────────────────────────────────────
    If Not dryRun And Not unmatchedOnly Then
        wsData.Range( _
            wsData.Cells(startRow, dm.ColDiscOut), _
            wsData.Cells(endRow, dm.ColDiscOut) _
        ).Value = outArr
    End If

    '── Step 8: Restore ──────────────────────────────────────────────────────
    If Not dryRun Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    End If

    '── Step 9: Report ───────────────────────────────────────────────────────
    If unmatchedOnly Then
        If cNoRule = 0 Then
            MsgBox "Sab rows mein rule mila! Koi unmatched nahi.", vbInformation
        Else
            WriteUnmatchedReport unmatchedLog, cNoRule, wsData.Name
        End If
    Else
        Dim tag As String
        tag = IIf(dryRun, " [DRY RUN]", "")
        Dim sm As String
        sm = "DISCOUNT ENGINE v3.1" & tag & vbNewLine
        sm = sm & "Sheet: " & dm.SheetName & vbNewLine
        sm = sm & String(38, "-") & vbNewLine
        sm = sm & "Applied   : " & cProcessed & " rows" & vbNewLine
        sm = sm & " Formula  : " & cFormula & " rows" & vbNewLine
        sm = sm & " Minimum  : " & cMinApplied & " rows" & vbNewLine
        sm = sm & "No Rule   : " & cNoRule & " rows" & vbNewLine
        sm = sm & "Skipped   : " & cSkipped & " rows" & vbNewLine
        If cNoRule > 0 Then
            sm = sm & vbNewLine & cNoRule & " rows unmatched." & vbNewLine
            sm = sm & "Run ShowUnmatchedRows() for detail."
        End If
        MsgBox sm, vbInformation, "Done"
    End If
End Sub

Private Function ProcessSingleRow(dataArr As Variant, rowIdx As Long, _
                                   rules() As DiscountRule, ruleCount As Long, _
                                   dsMap As DiscSheetMap, _
                                   dm As DataSheetMap) As ProcessResult
    Dim res As ProcessResult

    If Not IsRowComplete(dataArr, rowIdx, dm) Then
        res.WasSkipped = True
        res.SkipReason = "Blank in cols 1-" & dm.ColValidateUpto
        ProcessSingleRow = res
        Exit Function
    End If

    Dim rv As RowValues
    Call ExtractRowValues(dataArr, rowIdx, rv, dm)

    Dim bestIdx As Long, bestScore As Long
    bestIdx = FindBestRule(rules, ruleCount, rv, bestScore)

    If bestIdx < 1 Then
        res.RuleOwner = ""
        res.SkipReason = "No rule for '" & rv.AwbOwner & "'"
        ProcessSingleRow = res
        Exit Function
    End If

    Dim formula As String
    formula = GetSlabFormula(rules(bestIdx), rv.Weight, dsMap)

    Dim calcAmt As Double, minAmt As Double
    calcAmt = CalcDiscount(formula, rv)
    minAmt = rules(bestIdx).MinVal

    res.DiscountApplied = IIf(calcAmt > minAmt, calcAmt, minAmt)
    res.FormulaUsed = formula
    res.RuleOwner = rules(bestIdx).AwbOwner
    res.MatchScore = bestScore
    res.MinValUsed = minAmt
    ProcessSingleRow = res
End Function



'════════════════════════════════════════════════════════════════════════════════
'  §5  DATA SHEET COLUMN AUTO-DETECTOR  ← v3.1 core upgrade
'
'  Reads Row 1 of any sheet and maps fields to column numbers.
'  BLR:  AIRLINES at col 7  → dm.ColAirline = 7
'  CCU:  AIRLINES at col 9  → dm.ColAirline = 9  (auto!)
'  Future sheet: wherever AIRLINES header is → auto-detected
'════════════════════════════════════════════════════════════════════════════════

Private Function MapDataSheetHeaders(ws As Worksheet, ByRef dm As DataSheetMap) As Boolean
    MapDataSheetHeaders = False
    dm.SheetName = ws.Name

    Dim lastCol As Long
    lastCol = GetLastUsedCol(ws)
    If lastCol < 5 Then Exit Function

    ' Read entire header row into array (fast)
    Dim hArr As Variant
    hArr = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value

    ' Reset all to 0
    dm.ColDest = 0: dm.ColAirline = 0: dm.ColConsignor = 0
    dm.ColAwbOwner = 0: dm.ColMaterial = 0: dm.ColWeight = 0
    dm.ColRate = 0: dm.ColBasic = 0: dm.ColDueAgent = 0
    dm.ColGross = 0: dm.ColDiscOut = 0: dm.ColTotalGST = 0
    dm.ColValidateUpto = 0

    Dim c      As Long
    Dim hdr    As String
    Dim hdrRaw As String

    For c = 1 To lastCol
        If Not IsEmpty(hArr(1, c)) Then
            hdrRaw = Trim(CStr(hArr(1, c)))
            hdr = UCase(hdrRaw)

            ' ── DEST ───────────────────────────────────────────────────────
            If hdr = UCase(KW_DEST) And dm.ColDest = 0 Then
                dm.ColDest = c

            ' ── AIRLINES (exact keyword "AIRLINES" avoids "AIRLINE_MATCH") ─
            ElseIf hdr = UCase(KW_AIRLINES) And dm.ColAirline = 0 Then
                dm.ColAirline = c

            ' ── CONSIGNOR ──────────────────────────────────────────────────
            ElseIf InStr(1, hdr, UCase(KW_CONSIGNOR)) > 0 And dm.ColConsignor = 0 Then
                dm.ColConsignor = c

            ' ── AWB OWNED BY ───────────────────────────────────────────────
            ElseIf InStr(1, hdr, UCase(KW_AWB_OWNER)) > 0 And dm.ColAwbOwner = 0 Then
                dm.ColAwbOwner = c

            ' ── MATERIALS DESCRIPTION ──────────────────────────────────────
            ElseIf InStr(1, hdr, UCase(KW_MATERIAL)) > 0 And dm.ColMaterial = 0 Then
                dm.ColMaterial = c

            ' ── CH. WT (weight) ────────────────────────────────────────────
            ElseIf (InStr(1, hdr, UCase(KW_WEIGHT)) > 0 Or _
                    InStr(1, hdr, UCase(KW_WEIGHT2)) > 0) And dm.ColWeight = 0 Then
                dm.ColWeight = c

            ' ── RATE (standalone — not inside longer phrase) ───────────────
            ElseIf hdr = UCase(KW_RATE) And dm.ColRate = 0 Then
                dm.ColRate = c

            ' ── BASIC ──────────────────────────────────────────────────────
            ElseIf hdr = UCase(KW_BASIC) And dm.ColBasic = 0 Then
                dm.ColBasic = c

            ' ── DUE AGENT ──────────────────────────────────────────────────
            ElseIf InStr(1, hdr, UCase(KW_DUE_AGENT)) > 0 And dm.ColDueAgent = 0 Then
                dm.ColDueAgent = c

            ' ── GROSS AMMOUNT / GROSS AMOUNT ───────────────────────────────
            ElseIf InStr(1, hdr, UCase(KW_GROSS)) > 0 And dm.ColGross = 0 Then
                dm.ColGross = c

            ' ── DISCOUNT output column (first occurrence = the output col) ─
            ElseIf hdr = UCase(KW_DISC_OUT) And dm.ColDiscOut = 0 Then
                dm.ColDiscOut = c

            ' ── AMOUNT AFTER GST ───────────────────────────────────────────
            ElseIf InStr(1, hdr, UCase(KW_TOTAL_GST)) > 0 And dm.ColTotalGST = 0 Then
                dm.ColTotalGST = c

            ' ── OCDC (last column to validate completeness) ─────────────────
            ElseIf hdr = UCase(KW_VALIDATE_END) And dm.ColValidateUpto = 0 Then
                dm.ColValidateUpto = c
            End If
        End If
    Next c

    ' ── Apply fallbacks for any undetected columns ────────────────────────
    Dim usedFallback As Boolean
    usedFallback = False

    If dm.ColDest = 0 Then dm.ColDest = FB_DEST: usedFallback = True
    If dm.ColAirline = 0 Then dm.ColAirline = FB_AIRLINE: usedFallback = True
    If dm.ColConsignor = 0 Then dm.ColConsignor = FB_CONSIGNOR: usedFallback = True
    If dm.ColAwbOwner = 0 Then dm.ColAwbOwner = FB_AWB_OWNER: usedFallback = True
    If dm.ColMaterial = 0 Then dm.ColMaterial = FB_MATERIAL: usedFallback = True
    If dm.ColWeight = 0 Then dm.ColWeight = FB_WEIGHT: usedFallback = True
    If dm.ColRate = 0 Then dm.ColRate = FB_RATE: usedFallback = True
    If dm.ColBasic = 0 Then dm.ColBasic = FB_BASIC: usedFallback = True
    If dm.ColDueAgent = 0 Then dm.ColDueAgent = FB_DUE_AGENT: usedFallback = True
    If dm.ColGross = 0 Then dm.ColGross = FB_GROSS: usedFallback = True
    If dm.ColDiscOut = 0 Then dm.ColDiscOut = FB_DISC_OUT: usedFallback = True
    If dm.ColTotalGST = 0 Then dm.ColTotalGST = FB_TOTAL_GST: usedFallback = True
    If dm.ColValidateUpto = 0 Then dm.ColValidateUpto = FB_VALIDATE_UPTO: usedFallback = True

    dm.DetectedFrom = IIf(usedFallback, "HEADERS + FALLBACK", "HEADERS")

    ' Minimum viable check: AWB Owner and Discount output must be found
    If dm.ColAwbOwner = 0 Or dm.ColDiscOut = 0 Then
        Exit Function
    End If

    MapDataSheetHeaders = True
End Function



'════════════════════════════════════════════════════════════════════════════════
'  §6  RULE FINDER — Scoring Pipeline
'════════════════════════════════════════════════════════════════════════════════

Private Function FindBestRule(rules() As DiscountRule, ruleCount As Long, _
                               rv As RowValues, ByRef bestScore As Long) As Long
    FindBestRule = -1: bestScore = -1
    Dim j As Long

    For j = 1 To ruleCount
        Dim r As DiscountRule: r = rules(j)

        Dim awbSc As Long: awbSc = ScoreAwb(rv.AwbOwner, r.AwbOwner)
        If awbSc = SC_KILL Then GoTo NR

        Dim dSc As Long: dSc = ScoreExact(rv.Dest, r.Dest)
        If dSc = SC_KILL Then GoTo NR

        Dim cSc As Long: cSc = ScoreCons(rv.Consignor, r.Consignor)
        If cSc = SC_KILL Then GoTo NR

        Dim aSc As Long: aSc = ScoreExact(rv.Airline, r.Airline)
        If aSc = SC_KILL Then GoTo NR

        Dim rtSc As Long: rtSc = ScoreRate(rv.Rate, r.RateFilter)
        If rtSc = SC_KILL Then GoTo NR

        Dim mSc As Long: mSc = ScoreMat(rv.Material, r.Material)
        If mSc = SC_KILL Then GoTo NR

        Dim tot As Long: tot = awbSc + dSc + cSc + aSc + rtSc + mSc
        If tot > bestScore Then bestScore = tot: FindBestRule = j

NR: Next j
End Function

Private Function ScoreAwb(actual As String, pattern As String) As Long
    If pattern = actual Then ScoreAwb = SC_EXACT: Exit Function
    If CFG_AWB_PARTIAL Then
        If InStr(1, actual, pattern, vbTextCompare) > 0 Or _
           InStr(1, pattern, actual, vbTextCompare) > 0 Then
            ScoreAwb = SC_AWB_PARTIAL: Exit Function
        End If
    End If
    ScoreAwb = SC_KILL
End Function

Private Function ScoreExact(actual As String, pattern As String) As Long
    If pattern = actual Then ScoreExact = SC_EXACT: Exit Function
    If pattern = "ALL" Or pattern = "" Then ScoreExact = SC_WILDCARD: Exit Function
    ScoreExact = SC_KILL
End Function

Private Function ScoreCons(actual As String, pattern As String) As Long
    If pattern = "" Or pattern = "ALL" Then ScoreCons = 0: Exit Function
    If InStr(1, actual, pattern, vbTextCompare) > 0 Or _
       InStr(1, pattern, actual, vbTextCompare) > 0 Then
        ScoreCons = SC_CONSIGNOR: Exit Function
    End If
    ScoreCons = SC_KILL
End Function

Private Function ScoreRate(actualRate As Double, rf As String) As Long
    ScoreRate = SC_KILL
    If rf = "" Or rf = "ALL" Then ScoreRate = SC_WILDCARD: Exit Function
    If IsNumeric(rf) Then
        If CDbl(rf) = actualRate Then ScoreRate = SC_EXACT
        Exit Function
    End If
    If Len(rf) >= 2 Then
        Dim op As String: op = Left(rf, 1)
        Dim vs As String: vs = Trim(Mid(rf, 2))
        If IsNumeric(vs) Then
            Dim t As Double: t = CDbl(vs)
            If op = "<" And actualRate < t Then ScoreRate = SC_RATE_RANGE
            If op = ">" And actualRate > t Then ScoreRate = SC_RATE_RANGE
        End If
    End If
End Function

Private Function ScoreMat(shipMat As String, ruleMat As String) As Long
    If ruleMat = "" Or ruleMat = "ALL" Then ScoreMat = 0: Exit Function
    If shipMat = "" Then ScoreMat = SC_KILL: Exit Function
    If UCase(shipMat) = UCase(ruleMat) Then ScoreMat = SC_MATERIAL: Exit Function
    ScoreMat = SC_KILL
End Function



'════════════════════════════════════════════════════════════════════════════════
'  §7  DISCOUNT CALCULATOR
'  Handles both formula syntaxes found in this workbook:
'  OLD: Q*5%, O*3, R, Z-(O*34.36)
'  NEW: BASIC*5%, WIEGHT*3/=, DUE AGENT, TOTAL GST-W*34.36
'════════════════════════════════════════════════════════════════════════════════

Private Function CalcDiscount(formulaStr As String, rv As RowValues) As Double
    CalcDiscount = 0
    Dim f As String: f = UCase(Trim(formulaStr))
    If Len(f) = 0 Then Exit Function

    ' Direct number
    If IsNumeric(f) Then CalcDiscount = CDbl(f): Exit Function

    ' R / DUE AGENT
    If f = "R" Or InStr(1, f, "DUE") > 0 Then
        CalcDiscount = rv.DueAgent: Exit Function
    End If

    ' Z / TOTAL GST / GST type → totalGST - weight * X
    If Left(f, 1) = "Z" Or InStr(1, f, "TOTAL") > 0 Or InStr(1, f, "GST") > 0 Then
        Dim ls As Long: ls = InStrRev(f, "*")
        If ls > 0 Then
            Dim zn As String: zn = CN(Mid(f, ls + 1))
            If IsNumeric(zn) Then
                CalcDiscount = Round(rv.TotalGST - (rv.Weight * CDbl(zn)), 2)
            End If
        End If
        Exit Function
    End If

    Dim sp As Long: sp = InStr(1, f, "*")
    If sp = 0 Then Exit Function

    Dim pre As String: pre = Left(f, sp - 1)
    Dim suf As String: suf = CN(Mid(f, sp + 1))
    If Not IsNumeric(suf) Then Exit Function
    Dim nv As Double: nv = CDbl(suf)

    ' Q*X%  or  BASIC*X  → basicVal * X / 100
    If Left(pre, 1) = "Q" Or InStr(1, pre, "BASIC", vbTextCompare) > 0 Then
        CalcDiscount = Round(rv.BasicVal * nv / 100, 2): Exit Function
    End If

    ' O*X  or  WIEGHT*X  or  WEIGHT*X  → weight * X
    ' Note: "WIEGHT" is a real typo in DISCOUNT(2) sheet — handled intentionally
    If Left(pre, 1) = "O" Or _
       InStr(1, pre, "WEIGHT", vbTextCompare) > 0 Or _
       InStr(1, pre, "WIEGHT", vbTextCompare) > 0 Then
        CalcDiscount = Round(rv.Weight * nv, 2): Exit Function
    End If
End Function

' Strip parens, %, /=, spaces from number strings in formulas
Private Function CN(s As String) As String
    s = Replace(s, ")", ""): s = Replace(s, "(", "")
    s = Replace(s, "%", ""): s = Replace(s, "/=", "")
    CN = Trim(s)
End Function



'════════════════════════════════════════════════════════════════════════════════
'  §8  DISCOUNT SHEET LOADER
'════════════════════════════════════════════════════════════════════════════════

Private Function LoadDiscountRules(wsDis As Worksheet, ByRef dsMap As DiscSheetMap, _
                                   ByRef rules() As DiscountRule, _
                                   ByRef ruleCount As Long) As Boolean
    LoadDiscountRules = False: ruleCount = 0

    Dim lastRow As Long: lastRow = GetLastUsedRow(wsDis)
    Dim lastCol As Long: lastCol = GetLastUsedCol(wsDis)
    If lastRow < 2 Or lastCol < 3 Then
        MsgBox "DISCOUNT sheet mein data nahi hai!", vbCritical: Exit Function
    End If

    Dim dd As Variant
    dd = wsDis.Range(wsDis.Cells(1, 1), wsDis.Cells(lastRow, lastCol)).Value

    If Not MapDiscHeaders(dd, dsMap) Then
        MsgBox "DISCOUNT sheet headers nahi mile!", vbCritical: Exit Function
    End If
    If Not DetectSlabs(dd, dsMap) Then
        MsgBox "Weight slab columns nahi mile!", vbCritical: Exit Function
    End If

    dsMap.FormatStyle = DetectFormatStyle(dd, dsMap, lastRow)

    ReDim rules(1 To lastRow)
    Dim j As Long

    For j = 2 To lastRow
        Dim ov As Variant: ov = dd(j, dsMap.ColOwner)
        If IsEmpty(ov) Or Len(Trim(CStr(ov))) = 0 Then GoTo SR

        Dim rule As DiscountRule
        rule.SourceRow = j
        rule.AwbOwner = UCase(Trim(CStr(ov)))
        rule.Dest = DC(dd, j, dsMap.ColDest, "ALL")
        rule.Consignor = DC(dd, j, dsMap.ColCons, "ALL")
        rule.Airline = DC(dd, j, dsMap.ColAir, "ALL")
        rule.RateFilter = DC(dd, j, dsMap.ColRate, "ALL")
        rule.Material = DC(dd, j, dsMap.ColMaterial, "")
        If Len(rule.AwbOwner) = 0 Then GoTo SR

        rule.SlabCount = dsMap.SlabCount
        ReDim rule.SlabVals(0 To dsMap.SlabCount - 1)
        Dim s As Long
        For s = 0 To dsMap.SlabCount - 1
            Dim ci As Long: ci = dsMap.Slabs(s).ColIdx
            rule.SlabVals(s) = IIf(ci <= UBound(dd, 2), Trim(CStr(dd(j, ci))), "")
        Next s

        Dim fv As Variant: fv = dd(j, dsMap.FirstSlabCol)
        If IsNumeric(fv) Then
    rule.MinVal = CDbl(fv)
Else
    rule.MinVal = 0
End If

        ruleCount = ruleCount + 1
        rules(ruleCount) = rule
SR: Next j

    If ruleCount > 0 Then ReDim Preserve rules(1 To ruleCount)
    LoadDiscountRules = True
End Function

Private Function DetectFormatStyle(dd As Variant, dsMap As DiscSheetMap, _
                                   lastRow As Long) As String
    Dim j As Long, s As Long
    For j = 2 To Application.Min(lastRow, 10)
        For s = 0 To dsMap.SlabCount - 1
            Dim ci As Long: ci = dsMap.Slabs(s).ColIdx
            If ci <= UBound(dd, 2) Then
                Dim v As String: v = UCase(Trim(CStr(dd(j, ci))))
                If InStr(1, v, "BASIC") > 0 Or InStr(1, v, "WEIGHT") > 0 Or _
                   InStr(1, v, "WIEGHT") > 0 Or InStr(1, v, "DUE AGENT") > 0 Then
                    DetectFormatStyle = "NEW": Exit Function
                End If
                If Left(v, 1) = "Q" Or Left(v, 1) = "O" Or _
                   Left(v, 1) = "Z" Or v = "R" Then
                    DetectFormatStyle = "OLD": Exit Function
                End If
            End If
        Next s
    Next j
    DetectFormatStyle = "UNKNOWN"
End Function

Private Function MapDiscHeaders(dd As Variant, ByRef dsMap As DiscSheetMap) As Boolean
    MapDiscHeaders = False
    dsMap.ColOwner = 0: dsMap.ColDest = 0: dsMap.ColCons = 0
    dsMap.ColAir = 0: dsMap.ColRate = 0: dsMap.ColMaterial = 0
    Dim c As Long
    For c = 1 To UBound(dd, 2)
        If Not IsEmpty(dd(1, c)) Then
            Dim h As String: h = UCase(Trim(CStr(dd(1, c))))
            If InStr(1, h, DK_OWNER) > 0 And dsMap.ColOwner = 0 Then dsMap.ColOwner = c
            If InStr(1, h, DK_DEST) > 0 And dsMap.ColDest = 0 Then dsMap.ColDest = c
            If InStr(1, h, DK_CONS) > 0 And dsMap.ColCons = 0 Then dsMap.ColCons = c
            If InStr(1, h, DK_AIR) > 0 And dsMap.ColAir = 0 Then dsMap.ColAir = c
            If InStr(1, h, DK_RATE) > 0 And dsMap.ColRate = 0 Then dsMap.ColRate = c
            If InStr(1, h, DK_MAT) > 0 And dsMap.ColMaterial = 0 Then dsMap.ColMaterial = c
        End If
    Next c
    If dsMap.ColOwner = 0 Or dsMap.ColDest = 0 Or dsMap.ColCons = 0 Then Exit Function
    MapDiscHeaders = True
End Function

Private Function DetectSlabs(dd As Variant, ByRef dsMap As DiscSheetMap) As Boolean
    DetectSlabs = False
    Dim c As Long, n As Long
    Dim tC() As Long: ReDim tC(1 To UBound(dd, 2))
    Dim tB() As Double: ReDim tB(1 To UBound(dd, 2))
    Dim tH() As String: ReDim tH(1 To UBound(dd, 2))
    n = 0
    For c = 1 To UBound(dd, 2)
        If Not IsEmpty(dd(1, c)) Then
            Dim h As String: h = UCase(Trim(CStr(dd(1, c))))
            Dim ub As Double: ub = ParseSH(h)
            If ub > 0 Then n = n + 1: tC(n) = c: tB(n) = ub: tH(n) = CStr(dd(1, c))
        End If
    Next c
    If n = 0 Then
        n = 4
        tC(1) = 4: tB(1) = 100: tH(1) = "0-100(auto)"
        tC(2) = 5: tB(2) = 150: tH(2) = "101-150(auto)"
        tC(3) = 6: tB(3) = 300: tH(3) = "151-300(auto)"
        tC(4) = 7: tB(4) = 999999: tH(4) = "301+(auto)"
    End If
    dsMap.SlabCount = n
    dsMap.FirstSlabCol = tC(1)
    ReDim dsMap.Slabs(0 To n - 1)
    Dim s As Long
    For s = 1 To n
        dsMap.Slabs(s - 1).ColIdx = tC(s)
        dsMap.Slabs(s - 1).UpperBound = tB(s)
        dsMap.Slabs(s - 1).Header = tH(s)
    Next s
    DetectSlabs = True
End Function

Private Function ParseSH(h As String) As Double
    ParseSH = 0: h = UCase(Trim(h))
    If InStr(1, h, "ABOVE") > 0 Or InStr(1, h, "+") > 0 Then ParseSH = 999999: Exit Function
    If Left(h, 1) = ">" And IsNumeric(Mid(h, 2)) Then ParseSH = 999999: Exit Function
    Dim d As Long: d = InStr(1, h, "-")
    If d > 1 Then
        Dim rp As String: rp = Trim(Mid(h, d + 1))
        If IsNumeric(rp) Then ParseSH = CDbl(rp)
        If InStr(1, rp, "ABOVE") > 0 Then ParseSH = 999999
    End If
End Function

Private Function GetSlabFormula(rule As DiscountRule, wt As Double, _
                                 dsMap As DiscSheetMap) As String
    GetSlabFormula = ""
    Dim s As Long
    For s = 0 To dsMap.SlabCount - 1
        If wt <= dsMap.Slabs(s).UpperBound Then
            If s < rule.SlabCount Then GetSlabFormula = rule.SlabVals(s)
            Exit Function
        End If
    Next s
    If rule.SlabCount > 0 Then GetSlabFormula = rule.SlabVals(rule.SlabCount - 1)
End Function

' Safe cell read from discount data array
Private Function DC(dd As Variant, r As Long, c As Long, def As String) As String
    If c <= 0 Or c > UBound(dd, 2) Then DC = def: Exit Function
    Dim v As String: v = UCase(Trim(CStr(dd(r, c))))
    DC = IIf(Len(v) > 0, v, def)
End Function



'════════════════════════════════════════════════════════════════════════════════
'  §9  ROW EXTRACTION & VALIDATION
'════════════════════════════════════════════════════════════════════════════════

Private Sub ExtractRowValues(dataArr As Variant, ri As Long, _
                              ByRef rv As RowValues, dm As DataSheetMap)
    rv.AwbOwner = UCase(Trim(SafeStr(dataArr, ri, dm.ColAwbOwner)))
    rv.Dest = UCase(Trim(SafeStr(dataArr, ri, dm.ColDest)))
    rv.Consignor = UCase(Trim(SafeStr(dataArr, ri, dm.ColConsignor)))
    rv.Airline = UCase(Trim(SafeStr(dataArr, ri, dm.ColAirline)))
    rv.Material = UCase(Trim(SafeStr(dataArr, ri, dm.ColMaterial)))
    rv.Weight = SafeNum(dataArr, ri, dm.ColWeight)
    rv.Rate = SafeNum(dataArr, ri, dm.ColRate)
    rv.BasicVal = SafeNum(dataArr, ri, dm.ColBasic)
    rv.DueAgent = SafeNum(dataArr, ri, dm.ColDueAgent)
    rv.GrossAmt = SafeNum(dataArr, ri, dm.ColGross)
    rv.TotalGST = SafeNum(dataArr, ri, dm.ColTotalGST)
    If rv.TotalGST = 0 And rv.GrossAmt > 0 Then
        rv.TotalGST = rv.GrossAmt * (1 + CFG_GST_RATE)
    End If
End Sub

Private Function IsRowComplete(dataArr As Variant, ri As Long, _
                                dm As DataSheetMap) As Boolean
    Dim c As Long
    Dim maxC As Long: maxC = Application.Min(dm.ColValidateUpto, UBound(dataArr, 2))
    For c = 1 To maxC
        If Len(Trim(CStr(dataArr(ri, c)))) = 0 Then
            IsRowComplete = False: Exit Function
        End If
    Next c
    IsRowComplete = True
End Function



'════════════════════════════════════════════════════════════════════════════════
'  §10  REPORTING
'════════════════════════════════════════════════════════════════════════════════

Private Sub WriteUnmatchedReport(logText As String, rowCount As Long, srcSheet As String)
    ' Remove old unmatched sheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 9) = "Unmatched" Then
            Application.DisplayAlerts = False: ws.Delete: Application.DisplayAlerts = True
        End If
    Next ws

    Dim shName As String: shName = "Unmatched_" & Format(Now, "DDMMM_HHMM")
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsLog.Name = shName

    wsLog.Range("A1").Value = "Discount Engine v3.1 — Unmatched Rows from: " & srcSheet
    wsLog.Range("A2").Value = "Generated: " & Now & "  |  Total: " & rowCount

    Dim lines() As String: lines = Split(logText, vbNewLine)
    Dim i As Long
    For i = 0 To UBound(lines)
        If Len(Trim(lines(i))) > 0 Then
            Dim parts() As String: parts = Split(lines(i), "|")
            Dim p As Long
            For p = 0 To UBound(parts)
                wsLog.Cells(4 + i, p + 1).Value = Trim(parts(p))
            Next p
        End If
    Next i

    wsLog.Rows(4).Font.Bold = True
    wsLog.Columns("A:G").AutoFit
    wsLog.Activate

    MsgBox rowCount & " unmatched rows → sheet '" & shName & "'" & vbNewLine & vbNewLine & _
           "Sabse common reason: DISCOUNT sheet ka AWB Owner naam" & vbNewLine & _
           "data sheet ke naam se alag hai (e.g. MANISH vs MANISH SINGH)." & vbNewLine & _
           "CFG_AWB_PARTIAL = True already set hai — naam ki spelling check karo.", _
           vbInformation, "Unmatched Report"
End Sub



'════════════════════════════════════════════════════════════════════════════════
'  §11  UTILITIES
'════════════════════════════════════════════════════════════════════════════════

Private Function GetSheet(ByRef ws As Worksheet, pName As String, _
                           Optional aName As String = "") As Boolean
    GetSheet = False
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(pName): On Error GoTo 0
    If ws Is Nothing And Len(aName) > 0 Then
        On Error Resume Next: Set ws = ThisWorkbook.Sheets(aName): On Error GoTo 0
    End If
    If ws Is Nothing Then
        MsgBox "Sheet '" & pName & "' nahi mili!", vbCritical
    Else
        GetSheet = True
    End If
End Function

Private Function GetLastUsedRow(ws As Worksheet) As Long
    GetLastUsedRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
End Function

Private Function GetLastUsedCol(ws As Worksheet) As Long
    GetLastUsedCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
End Function

Private Function MaxColNeeded(dm As DataSheetMap) As Long
    MaxColNeeded = Application.Max( _
        dm.ColDest, dm.ColAirline, dm.ColConsignor, dm.ColAwbOwner, _
        dm.ColMaterial, dm.ColWeight, dm.ColRate, dm.ColBasic, _
        dm.ColDueAgent, dm.ColGross, dm.ColDiscOut, dm.ColTotalGST, _
        dm.ColValidateUpto)
End Function

Private Function SafeStr(arr As Variant, r As Long, c As Long) As String
    If c >= 1 And c <= UBound(arr, 2) Then SafeStr = CStr(arr(r, c)) Else SafeStr = ""
End Function

Private Function SafeNum(arr As Variant, r As Long, c As Long) As Double
    SafeNum = 0
    If c >= 1 And c <= UBound(arr, 2) Then
        If IsNumeric(arr(r, c)) Then SafeNum = CDbl(arr(r, c))
    End If
End Function

Private Function SafeGet(arr As Variant, r As Long, c As Long) As Variant
    If c >= 1 And c <= UBound(arr, 2) Then SafeGet = arr(r, c) Else SafeGet = 0
End Function

Private Function PadR(s As String, n As Long) As String
    PadR = s & Space(Application.Max(0, n - Len(s)))
End Function
