Option Explicit

' ============================================================================
'  SAP2000 - Box Culvert Model Build  (writes a .$2k text model)
'
'  Writes <workbook name>.$2k next to this workbook - e.g. 3.00m.xlsm gives
'  3.00m.$2k in the same folder - overwriting any file of that name. Open it
'  in SAP2000 with File > Import > SAP2000 .s2k/.$2k Text File. Nothing is
'  sent anywhere and no worksheet is changed; SAP_INPUT and SAP_RESULTS are
'  not read or written (SAP_INPUT is an older formula generator, left as-is).
'
'  SAME MODEL AS THE MIDAS BUILDER. Inputs, geometry, loads and combinations
'  come from procedures copied BYTE FOR BYTE out of
'  midas-culvert-model-build.bas (ReadInputs, GenerateGeometry,
'  GenerateBeamLoads, GenerateLoadCombinations and their helpers, plus the
'  seismic gate). tests/verify_sap2000_build.py fails if a copy drifts, so a
'  change to the MIDAS model has to be carried over here too. Only the
'  SAP-specific part is this module's own: dividing and renumbering, the
'  load translation and the table writer.
'
'  What the MIDAS builder does through the API happens here in the text:
'    ope/DIVIDEELEM   elements 18, 10, 11, 4 split into INPUT!K15 equal
'                     frames; each piece carries its share of a trapezoid
'    db/NSPR          joint springs on every z = 0 joint, same tributary rule
'    db/BODF          DL = SelfWtMult 1; ATA = FRAME LOADS - GRAVITY along X
'  Joints and frames are numbered 1..n in the MIDAS element order
'  (foundation, walls, slab), division joints after the original ones.
'
'  Differences from the hand-built reference model (SAP2000/*.$2k in the
'  repo), all by the owner's decision on 2026-09-24:
'    - full MIDAS load set (20 cases less the seismic pair when the gate is
'      off) and all 33/35 combinations, not the reference's subset
'    - rigid-zone and haunch members get their own sections (MIDAS_INPUT
'      rows 10-16) instead of the plain slab/wall/foundation one
'    - XZ plane frame (UX UZ RY active), so springs at z = 0 alone make it
'      stable; the reference had all six DOF
'    - load-case names lose their "_" (EHS2_L -> EHS2L, as the reference
'      names them); combination names are kept exactly as in MIDAS
'    - no GUIDs: SAP2000 assigns them on import
'  Ec is MIDAS_INPUT!B43, falling back to 33 GPa (EN C30/37, the reference's
'  value) - NOT to the MIDAS builder's 26.291 GPa fallback.
'
'  No restraints are written: the foundation springs are the supports, as
'  in the reference.
' ============================================================================


' ---------------------------------------------------------------------------
'  CONFIG
' ---------------------------------------------------------------------------

' Stamped into the report title. Bump with every change to this file.
Private Const SCRIPT_VERSION As String = "2026-09-24a"

' One line, no "_" continuation, no "|" - read by the updater's manifest.
Private Const SCRIPT_CHANGELOG As String = "First version: writes <workbook>.$2k beside the workbook from MIDAS_INPUT, same geometry, loads and combinations as the MIDAS culvert builder."

' Identifies this module to the updater whatever it was named in Excel.
Private Const SCRIPT_ID As String = "sap2000-culvert-model-build"

' Written into PROGRAM CONTROL. The reference file came from 17.3.0; later
' versions import older text files.
Private Const SAP_PROGRAM_VERSION As String = "17.3.0"

Private Const SAP_FILE_EXTENSION As String = ".$2k"

' Concrete, as the reference file defines it (EN 1992-1-1 C30/37).
Private Const MATERIAL_NAME As String = "C30/37"
Private Const MATERIAL_FC As Double = 30000#          ' kN/m2
Private Const MATERIAL_POISN As Double = 0.2
Private Const MATERIAL_THERMAL As Double = 0.00001
Private Const MATERIAL_ELAST_DEFAULT As Double = 33000000#  ' fallback if B43 is blank/invalid
Private Const MATERIAL_DEN_DEFAULT As Double = 25           ' fallback if INPUT!B5 is blank/invalid
Private Const GRAVITY_ACCEL As Double = 9.80665

' Every section is a solid rectangle <depth> x 1.0 m, the per-metre strip.
Private Const SECTION_WIDTH As Double = 1#

' Beam-load values are rounded to 3 dp, like the MIDAS builder's AddLoad.
Private Const LOAD_ROUND_DP As Long = 3

' Coordinates and section properties are written with at most this many
' decimals, so 0.1 + 0.2 comes out as 0.3, not 0.30000000000000004.
Private Const COORD_ROUND_DP As Long = 6

' Joints within this distance of z = 0 are foundation joints (springs).
Private Const FOUNDATION_Z_TOL As Double = 0.001

' Section colours, the MIDAS builder's palette (RGB per section number), so
' the two models read the same on screen.
Private Const SECTION_COLOR_LIST As String = _
    "1|45,110,200;2|60,160,75;3|150,150,150;4|110,90,70;" & _
    "5|70,150,205;6|95,185,110;7|205,50,45;8|225,130,35;9|185,70,45"

' MIDAS load-case type -> SAP2000 DesignType. LLacc is typed "E" in the
' MIDAS list but is an accidental live load, not an earthquake - it is
' mapped to LIVE by name in SapDesignType.
Private Const SAP_DESIGN_TYPE_LIST As String = _
    "D|DEAD;EV|DEAD;EH|DEAD;L|LIVE;LS|LIVE;E|QUAKE;FP|OTHER"


' ---------------------------------------------------------------------------
'  INPUT SHEET - the MIDAS builder's cells, read by the copied ReadInputs.
'  See midas-culvert-model-build.bas for the full cell map; in short:
'    B4 slab, B5 wall, B7 foundation thickness, B18 span, B19 wall height
'    (all required > 0); B8/B9 haunch, B10-B13/B16 haunch and rigid-zone
'    sections, B15 kh, B17 foundation extension; loads B21-B40; B42 subgrade
'    modulus ks; B43 Ec. INPUT!B5 unit weight, INPUT!K15 division count,
'    INPUT!G8/G9/G11/N15/N16 project information.
' ---------------------------------------------------------------------------
Private Const INPUT_SHEET_NAME As String = "MIDAS_INPUT"
Private Const PROJECT_SHEET_NAME As String = "INPUT"
Private Const CELL_PROJECT_TYPE As String = "G9"
Private Const CELL_PROJECT_NO As String = "G8"
Private Const CELL_PROJECT_PART As String = "G11"
Private Const CELL_PROJECT_ENGINEER As String = "N15"
Private Const CELL_PROJECT_REVISION As String = "N16"
Private Const CELL_DIVIDE_AMOUNT As String = "K15"

Private Const PROJINFO_COMPANY As String = "DEHA"

' ---------------------------------------------------------------------------
'  SEISMIC GATE - identical to the MIDAS builder (see its header):
'    -1 auto (1_GIRIS!P81, then O81 < Q81, then the INPUT Z/H ratio)
'     0 force OFF, 1 force ON.
'  Off drops the EQ and ATA cases, the EQ wall loads, the ATA inertia, EQ-1
'  and ENV_EQ. Auto-detect that finds nothing usable stops the export.
' ---------------------------------------------------------------------------
Private Const SEISMIC_GATE_OVERRIDE As Long = -1
Private SEISMIC_ACTIVE As Boolean
Private SEISMIC_GATE_SOURCE As String
Private SEISMIC_GATE_KNOWN As Boolean

' Geometry + section inputs (names shared with the copied procedures)
Private DIM_SLAB_T As Double         ' B4
Private DIM_WALL_T As Double         ' B5
Private DIM_SECT3 As Double          ' B6
Private DIM_FOUND_T As Double        ' B7
Private DIM_SLAB_HAUNCH As Double    ' B8
Private DIM_WALL_HAUNCH As Double    ' B9
Private DIM_SECT5 As Double          ' B10
Private DIM_SECT6 As Double          ' B11
Private DIM_SECT7 As Double          ' B12
Private DIM_SECT8 As Double          ' B13
Private DIM_SECT9 As Double          ' B16
Private DIM_ATA_FACTOR As Double     ' B15
Private DIM_EXT As Double            ' B17
Private DIM_SPAN As Double           ' B18
Private DIM_WALL_H As Double         ' B19
Private DIVIDE_AMOUNT As Long        ' INPUT!K15
Private SPRING_KZ_MODULUS As Double  ' MIDAS_INPUT!B42
Private MATERIAL_ELAST As Double     ' MIDAS_INPUT!B43
Private MATERIAL_DEN As Double       ' INPUT!B5
Private INPUT_FALLBACKS As String

' Load magnitude inputs
Private LD_EV1 As Double, LD_EV2 As Double
Private LD_EV1_EXT As Double, LD_EV2_EXT As Double
Private LD_EHS1_TOP As Double, LD_EHS1_BOT As Double
Private LD_EHA1_TOP As Double, LD_EHA1_BOT As Double
Private LD_EHS2_TOP As Double, LD_EHS2_BOT As Double
Private LD_EHA2_TOP As Double, LD_EHA2_BOT As Double
Private LD_LL1 As Double, LD_LL As Double, LD_LLACC As Double
Private LD_LSS1_L As Double, LD_LSS2_L As Double, LD_LSA2_L As Double
Private LD_EQ_TOP As Double, LD_EQ_BOT As Double

Private PROJECT_NAME As String
Private PROJECT_ENGINEER As String
Private PROJECT_REVISION As String

Private ORD_EHS1(1 To 8) As Double
Private ORD_EHA1(1 To 8) As Double
Private ORD_EHS2(1 To 8) As Double
Private ORD_EHA2(1 To 8) As Double
Private ORD_EQ(1 To 8) As Double

' Generated MIDAS-numbered model (same formats as the MIDAS builder)
Private NODE_LIST As String       ' "no|x|y|z;..."
Private ELEM_LIST As String       ' "no|sect|n1|n2|angle;..."
Private SECT_LIST As String       ' "no|name|depth;..."
Private BEAMLOAD_LIST As String   ' "elem|lcname|cmd|dir|p1|p2;..."
Private LOADCOMB_LIST As String   ' "name|ACTIVE|iTYPE|ANAL:LC:F,...|tier;..."

Private HAS_EXT As Boolean
Private SLAB_E As Long
Private WALL_E As Long

' The MIDAS static load cases, in the MIDAS builder's order.
Private Const STLDCASE_LIST As String = _
    "DL|D;EV1|EV;EV2|EV;EVin|EV;EHS1|EH;EHA1|EH;EHS2_L|EH;EHA2_L|EH;" & _
    "EHS2_R|EH;EHA2_R|EH;LL1|L;LL|L;LLin|L;LLacc|E;LSS1_L|LS;LSS2_L|LS;" & _
    "LSA2_L|LS;WA|FP;EQ|E;ATA|E"

' ---------------------------------------------------------------------------
'  SAP MODEL - the divided, renumbered model ExpandModel builds.
'  Joints 1..JT_COUNT, frames 1..FR_COUNT. ELEM_FIRST_FRAME/ELEM_PIECES
'  map a MIDAS element number to its frames (indexed by element number).
' ---------------------------------------------------------------------------
Private JT_COUNT As Long
Private JT_X() As Double
Private JT_Z() As Double
Private JT_SPECIAL() As Boolean
Private FR_COUNT As Long
Private FR_I() As Long
Private FR_J() As Long
Private FR_SECT() As Long
Private FR_ANGLE() As Double
Private ELEM_FIRST_FRAME() As Long
Private ELEM_PIECES() As Long
Private DIVIDED_ELEMS As String   ' for the report, e.g. "18, 10, 11, 4"
Private SPRING_COUNT As Long
Private LOAD_COUNT As Long
' Warnings collected while the text is assembled (e.g. springs skipped).
Private TEXT_WARNINGS As String

' Elements the MIDAS builder divides (ope/DIVIDEELEM), in its order.
Private Const DIVIDE_TARGETS As String = "18,10,11,4"

' Output text, one entry per line; OUT_COUNT lines used.
Private OUT_LINES() As String
Private OUT_COUNT As Long


' ===========================================================================
'  ENTRY POINT
' ===========================================================================

Public Sub BuildSap2000Model()

    Dim report As String
    Dim ok As Boolean
    Dim path As String
    Dim res As String
    Dim stage As String

    report = "SAP2000 culvert model  [" & SCRIPT_VERSION & "]" & vbCrLf & _
             String(40, "-") & vbCrLf

    On Error GoTo Crashed

    stage = "seismic gate"
    Call InitSeismicGate
    If Not SEISMIC_GATE_KNOWN Then
        MsgBox report & "STOPPED - no file was written." & vbCrLf & vbCrLf & _
               "The seismic gate " & SEISMIC_GATE_SOURCE & "." & vbCrLf & vbCrLf & _
               "Fix those cells, or set SEISMIC_GATE_OVERRIDE to 0 (off) or 1 (on) " & _
               "at the top of this module.", vbExclamation
        Exit Sub
    End If
    report = report & "NOTE - seismic gate is " & IIf(SEISMIC_ACTIVE, "ON", "OFF") & _
             " (" & SEISMIC_GATE_SOURCE & ")." & vbCrLf & String(40, "-") & vbCrLf

    stage = "output path"
    path = SapOutputPath()
    ok = StepResult(report, "Output file", IIf(Len(path) = 0, _
             "the workbook has never been saved, so there is no folder to write " & _
             "into. Save it first.", ""))

    stage = "read inputs"
    If ok Then
        res = ReadInputs()
        If Len(res) = 0 And Len(INPUT_FALLBACKS) > 0 Then
            res = "WARN: used built-in defaults -" & INPUT_FALLBACKS
        End If
        ok = StepResult(report, "Read inputs", res)
    End If

    stage = "generate model"
    If ok Then
        Call ReadProjectName
        Call GenerateGeometry
        Call ComputeOrdinates
        Call GenerateBeamLoads
        Call GenerateLoadCombinations
        ok = StepResult(report, "Divide + number", ExpandModel())
    End If

    stage = "assemble tables"
    If ok Then ok = StepResult(report, "Model text", WriteModelText())

    stage = "write file"
    If ok Then ok = StepResult(report, "Write " & path, WriteOutputFile(path))

    If ok Then
        report = report & String(40, "-") & vbCrLf & _
                 JT_COUNT & " joints, " & FR_COUNT & " frames (divided: " & DIVIDED_ELEMS & _
                 "), " & SPRING_COUNT & " springs, " & LOAD_COUNT & " frame loads, " & _
                 (UBound(Split(LOADCOMB_LIST, ";")) + 1) & " combinations." & vbCrLf & _
                 "Import in SAP2000: File > Import > SAP2000 .s2k/.$2k Text File." & vbCrLf
    End If

    report = VerdictFirst(report, ok)
    MsgBox FitReport(report, SCRIPT_ID), IIf(ok, vbInformation, vbExclamation)
    Exit Sub

Crashed:
    report = report & "FAIL - " & stage & ": VBA error " & Err.Number & " - " & _
             Err.Description & vbCrLf
    report = VerdictFirst(report, False)
    MsgBox FitReport(report, SCRIPT_ID), vbExclamation

End Sub

' <folder of this workbook>\<workbook name without extension>.$2k, or ""
' when the workbook has never been saved.
Private Function SapOutputPath() As String

    Dim nm As String
    Dim p As Long

    If Len(ThisWorkbook.Path) = 0 Then Exit Function

    nm = ThisWorkbook.Name
    p = InStrRev(nm, ".")
    If p > 1 Then nm = Left$(nm, p - 1)

    SapOutputPath = ThisWorkbook.Path & "\" & nm & SAP_FILE_EXTENSION

End Function


' ===========================================================================
'  COPIED FROM midas-culvert-model-build.bas - DO NOT EDIT HERE.
'  Byte-identical copies; tests/verify_sap2000_build.py compares them. Make
'  the change in the MIDAS builder and copy the procedure across.
' ===========================================================================

Private Sub InitSeismicGate()

    SEISMIC_GATE_KNOWN = True

    If SEISMIC_GATE_OVERRIDE = 0 Then
        SEISMIC_ACTIVE = False
        SEISMIC_GATE_SOURCE = "forced OFF by SEISMIC_GATE_OVERRIDE = 0"
    ElseIf SEISMIC_GATE_OVERRIDE = 1 Then
        SEISMIC_ACTIVE = True
        SEISMIC_GATE_SOURCE = "forced ON by SEISMIC_GATE_OVERRIDE = 1"
    Else
        SEISMIC_ACTIVE = DetectSeismicGate()
    End If

End Sub

Private Function DetectSeismicGate() As Boolean

    Dim wsGiris As Worksheet
    Dim wsInput As Worksheet
    Dim pVal As String
    Dim oVal As Double, qVal As Double
    Dim hCover As Double, hWall As Double, tTop As Double, tBot As Double
    Dim ratio As Double

    On Error Resume Next
    Set wsGiris = ThisWorkbook.Worksheets("1_GIRIS")
    If Not wsGiris Is Nothing Then
        pVal = Trim$(CStr(wsGiris.Range("P81").Value))
        If pVal = "<" Then
            DetectSeismicGate = True
            SEISMIC_GATE_SOURCE = "1_GIRIS!P81 = ""<"""
            Exit Function
        ElseIf pVal = ">" Then
            DetectSeismicGate = False
            SEISMIC_GATE_SOURCE = "1_GIRIS!P81 = "">"""
            Exit Function
        End If
        If IsNumeric(wsGiris.Range("O81").Value) And IsNumeric(wsGiris.Range("Q81").Value) Then
            oVal = CDbl(wsGiris.Range("O81").Value)
            qVal = CDbl(wsGiris.Range("Q81").Value)
            DetectSeismicGate = (oVal < qVal)
            SEISMIC_GATE_SOURCE = "1_GIRIS!O81 = " & Format$(oVal, "0.000") & _
                                  IIf(oVal < qVal, " < ", " >= ") & Format$(qVal, "0.000")
            Exit Function
        End If
    End If

    Set wsInput = ThisWorkbook.Worksheets(PROJECT_SHEET_NAME)
    If wsInput Is Nothing Then
        Set wsInput = ThisWorkbook.Worksheets("INPUT")
    End If
    If Not wsInput Is Nothing Then
        If IsNumeric(wsInput.Range("B21").Value) And IsNumeric(wsInput.Range("B23").Value) And _
           IsNumeric(wsInput.Range("B12").Value) And IsNumeric(wsInput.Range("B14").Value) Then
            hCover = CDbl(wsInput.Range("B21").Value)
            hWall = CDbl(wsInput.Range("B23").Value)
            tTop = CDbl(wsInput.Range("B12").Value)
            tBot = CDbl(wsInput.Range("B14").Value)
            If (hWall + tTop + tBot) > 0 Then
                ratio = hCover / (hWall + tTop + tBot)
                DetectSeismicGate = (ratio < 0.5)
                SEISMIC_GATE_SOURCE = "Z/H = INPUT!B21/(B23+B12+B14) = " & Format$(ratio, "0.000") & _
                                      IIf(ratio < 0.5, " < ", " >= ") & "0.500"
                Exit Function
            End If
        End If
    End If
    On Error GoTo 0

    DetectSeismicGate = False
    SEISMIC_GATE_KNOWN = False
    SEISMIC_GATE_SOURCE = "could not be determined - neither 1_GIRIS!P81/O81/Q81 nor " & _
                          "INPUT!B21/B23/B12/B14 hold usable values"

End Function

Private Function ReadInputs() As String

    INPUT_FALLBACKS = ""

    Dim ws As Worksheet
    Dim wsInput As Worksheet
    Dim nameSect9 As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(INPUT_SHEET_NAME)
    On Error GoTo 0

    Call InitSeismicGate

    If ws Is Nothing Then
        ReadInputs = "Sheet """ & INPUT_SHEET_NAME & """ not found."
        Exit Function
    End If

    ReadInputs = RequirePositive(ws, "B4", "slab thickness", DIM_SLAB_T)
    If Len(ReadInputs) > 0 Then Exit Function
    ReadInputs = RequirePositive(ws, "B5", "wall thickness", DIM_WALL_T)
    If Len(ReadInputs) > 0 Then Exit Function
    ReadInputs = RequirePositive(ws, "B7", "foundation thickness", DIM_FOUND_T)
    If Len(ReadInputs) > 0 Then Exit Function
    ReadInputs = RequirePositive(ws, "B18", "clear span", DIM_SPAN)
    If Len(ReadInputs) > 0 Then Exit Function
    ReadInputs = RequirePositive(ws, "B19", "wall height", DIM_WALL_H)
    If Len(ReadInputs) > 0 Then Exit Function

    ' Optional / zero-allowed inputs - a blank cell reads as 0, which is the
    ' MCT module's own "feature off" value for B6/B8/B9/B17.
    If Not TryReadCell(ws, "B6", DIM_SECT3) Then ReadInputs = BadOptionalCell("B6"): Exit Function
    If Not TryReadCell(ws, "B8", DIM_SLAB_HAUNCH) Then ReadInputs = BadOptionalCell("B8"): Exit Function
    If Not TryReadCell(ws, "B9", DIM_WALL_HAUNCH) Then ReadInputs = BadOptionalCell("B9"): Exit Function
    If Not TryReadCell(ws, "B10", DIM_SECT5) Then ReadInputs = BadOptionalCell("B10"): Exit Function
    If Not TryReadCell(ws, "B11", DIM_SECT6) Then ReadInputs = BadOptionalCell("B11"): Exit Function
    If Not TryReadCell(ws, "B12", DIM_SECT7) Then ReadInputs = BadOptionalCell("B12"): Exit Function
    If Not TryReadCell(ws, "B13", DIM_SECT8) Then ReadInputs = BadOptionalCell("B13"): Exit Function
    If Not TryReadCell(ws, "B16", DIM_SECT9) Then ReadInputs = BadOptionalCell("B16"): Exit Function
    If Not TryReadCell(ws, "B15", DIM_ATA_FACTOR) Then ReadInputs = BadOptionalCell("B15"): Exit Function
    If Not TryReadCell(ws, "B17", DIM_EXT) Then ReadInputs = BadOptionalCell("B17"): Exit Function

    ' Divide amount from INPUT!K15 (subdivision count for elements 18, 10, 11, 4)
    DIVIDE_AMOUNT = 0
    On Error Resume Next
    Set wsInput = ThisWorkbook.Worksheets(PROJECT_SHEET_NAME)
    If Not wsInput Is Nothing Then
        If IsNumeric(wsInput.Range(CELL_DIVIDE_AMOUNT).Value) Then
            DIVIDE_AMOUNT = CLng(wsInput.Range(CELL_DIVIDE_AMOUNT).Value)
        End If
    End If
    If DIVIDE_AMOUNT <= 0 Then
        If IsNumeric(ws.Range(CELL_DIVIDE_AMOUNT).Value) Then
            DIVIDE_AMOUNT = CLng(ws.Range(CELL_DIVIDE_AMOUNT).Value)
        End If
    End If
    On Error GoTo 0

    If Not TryReadCell(ws, "B21", LD_EV1) Then ReadInputs = BadOptionalCell("B21"): Exit Function
    If Not TryReadCell(ws, "B22", LD_EV2) Then ReadInputs = BadOptionalCell("B22"): Exit Function
    If Not TryReadCell(ws, "B37", LD_EV1_EXT) Then ReadInputs = BadOptionalCell("B37"): Exit Function
    If Not TryReadCell(ws, "B36", LD_EV2_EXT) Then ReadInputs = BadOptionalCell("B36"): Exit Function
    If Not TryReadCell(ws, "B23", LD_EHS1_TOP) Then ReadInputs = BadOptionalCell("B23"): Exit Function
    If Not TryReadCell(ws, "B24", LD_EHS1_BOT) Then ReadInputs = BadOptionalCell("B24"): Exit Function
    If Not TryReadCell(ws, "B25", LD_EHA1_TOP) Then ReadInputs = BadOptionalCell("B25"): Exit Function
    If Not TryReadCell(ws, "B26", LD_EHA1_BOT) Then ReadInputs = BadOptionalCell("B26"): Exit Function
    If Not TryReadCell(ws, "B27", LD_EHS2_TOP) Then ReadInputs = BadOptionalCell("B27"): Exit Function
    If Not TryReadCell(ws, "B28", LD_EHS2_BOT) Then ReadInputs = BadOptionalCell("B28"): Exit Function
    If Not TryReadCell(ws, "B29", LD_EHA2_TOP) Then ReadInputs = BadOptionalCell("B29"): Exit Function
    If Not TryReadCell(ws, "B30", LD_EHA2_BOT) Then ReadInputs = BadOptionalCell("B30"): Exit Function
    If Not TryReadCell(ws, "B31", LD_LL1) Then ReadInputs = BadOptionalCell("B31"): Exit Function
    If Not TryReadCell(ws, "B32", LD_LL) Then ReadInputs = BadOptionalCell("B32"): Exit Function
    If Not TryReadCell(ws, "B39", LD_LLACC) Then ReadInputs = BadOptionalCell("B39"): Exit Function
    If Not TryReadCell(ws, "B38", LD_LSS1_L) Then ReadInputs = BadOptionalCell("B38"): Exit Function
    If Not TryReadCell(ws, "B33", LD_LSS2_L) Then ReadInputs = BadOptionalCell("B33"): Exit Function
    If Not TryReadCell(ws, "B40", LD_LSA2_L) Then ReadInputs = BadOptionalCell("B40"): Exit Function
    If Not TryReadCell(ws, "B34", LD_EQ_TOP) Then ReadInputs = BadOptionalCell("B34"): Exit Function
    If Not TryReadCell(ws, "B35", LD_EQ_BOT) Then ReadInputs = BadOptionalCell("B35"): Exit Function
    If Not TryReadCell(ws, "B42", SPRING_KZ_MODULUS) Then ReadInputs = BadOptionalCell("B42"): Exit Function

    ' Ec from MIDAS_INPUT!B43; unit weight from the "INPUT" sheet's B5.
    ' Neither blocks the build if missing/invalid - falls back to the
    ' MCT row's own defaults instead (see MATERIAL_ELAST_DEFAULT/
    ' MATERIAL_DEN_DEFAULT).
    If Not TryReadCell(ws, "B43", MATERIAL_ELAST) Or MATERIAL_ELAST <= 0 Then
        MATERIAL_ELAST = MATERIAL_ELAST_DEFAULT
        INPUT_FALLBACKS = INPUT_FALLBACKS & " Ec (" & INPUT_SHEET_NAME & _
                          "!B43) missing or <= 0, used " & JsonNum(MATERIAL_ELAST_DEFAULT) & "."
    End If
    MATERIAL_DEN = 0
    On Error Resume Next
    If Not wsInput Is Nothing Then
        If IsNumeric(wsInput.Range("B5").Value) Then
            MATERIAL_DEN = CDbl(wsInput.Range("B5").Value)
        End If
    End If
    On Error GoTo 0
    If MATERIAL_DEN <= 0 Then
        MATERIAL_DEN = MATERIAL_DEN_DEFAULT
        INPUT_FALLBACKS = INPUT_FALLBACKS & " Unit weight (" & PROJECT_SHEET_NAME & _
                          "!B5) missing or <= 0, used " & JsonNum(MATERIAL_DEN_DEFAULT) & "."
    End If

    ' Section names come from column A, alongside each dimension.
    SECT_LIST = SectRow(1, ws.Range("A4").Value, DIM_SLAB_T)
    SECT_LIST = SECT_LIST & ";" & SectRow(2, ws.Range("A5").Value, DIM_WALL_T)
    If DIM_SECT3 > 0 Then SECT_LIST = SECT_LIST & ";" & SectRow(3, ws.Range("A6").Value, DIM_SECT3)
    SECT_LIST = SECT_LIST & ";" & SectRow(4, ws.Range("A7").Value, DIM_FOUND_T)
    ' Sections 5 and 6 are gated on B8/B9 but take their name/size from
    ' rows 10/11 - that cross-reference is in the MCT module, kept as-is.
    If DIM_SLAB_HAUNCH > 0 Then SECT_LIST = SECT_LIST & ";" & SectRow(5, ws.Range("A10").Value, DIM_SECT5)
    If DIM_WALL_HAUNCH > 0 Then SECT_LIST = SECT_LIST & ";" & SectRow(6, ws.Range("A11").Value, DIM_SECT6)
    SECT_LIST = SECT_LIST & ";" & SectRow(7, ws.Range("A12").Value, DIM_SECT7)
    SECT_LIST = SECT_LIST & ";" & SectRow(8, ws.Range("A13").Value, DIM_SECT8)
    If DIM_SECT9 > 0 Then
        nameSect9 = Trim$(CStr(ws.Range("A16").Value))
        If Len(nameSect9) = 0 Or StrComp(nameSect9, "TEMEL RIJIT", vbTextCompare) = 0 Then
            nameSect9 = "TEM RIJIT"
        End If
        SECT_LIST = SECT_LIST & ";" & SectRow(9, nameSect9, DIM_SECT9)
    End If

    ReadInputs = ""

End Function

Private Function SectRow(ByVal no As Long, ByVal nm As Variant, ByVal depth As Double) As String
    SectRow = no & "|" & Trim$(CStr(nm)) & "|" & JsonNum(depth)
End Function

Private Function TryReadCell(ByVal ws As Worksheet, ByVal addr As String, _
                             ByRef result As Double) As Boolean
    Dim v As Variant
    v = ws.Range(addr).Value
    If IsEmpty(v) Then
        result = 0
        TryReadCell = True
    ElseIf IsNumeric(v) Then
        result = CDbl(v)
        TryReadCell = True
    Else
        ' Never leave the caller holding the previous run's value - these
        ' targets are module-level and survive between runs.
        result = 0
        TryReadCell = False
    End If
End Function

Private Function RequirePositive(ByVal ws As Worksheet, ByVal addr As String, _
                                 ByVal label As String, ByRef result As Double) As String

    Dim v As Variant

    v = ws.Range(addr).Value

    If IsEmpty(v) Or IsError(v) Then
        RequirePositive = ws.Name & "!" & addr & " (" & label & ") is blank or an error value."
    ElseIf Not IsNumeric(v) Then
        RequirePositive = ws.Name & "!" & addr & " (" & label & ") is not a number."
    ElseIf CDbl(v) <= 0 Then
        RequirePositive = ws.Name & "!" & addr & " (" & label & ") must be greater than 0."
    Else
        result = CDbl(v)
    End If

End Function

Private Function BadOptionalCell(ByVal addr As String) As String
    BadOptionalCell = INPUT_SHEET_NAME & "!" & addr & " holds text or an error value. " & _
                      "Leave it blank for 0, or fix the formula."
End Function

Private Sub ReadProjectName()

    Dim ws As Worksheet
    Dim pType As String, pNo As String, pPart As String
    Dim s As String

    PROJECT_NAME = ""
    PROJECT_ENGINEER = ""
    PROJECT_REVISION = ""

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PROJECT_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    pType = Trim$(CStr(ws.Range(CELL_PROJECT_TYPE).Value))
    pNo = Trim$(CStr(ws.Range(CELL_PROJECT_NO).Value))
    pPart = Trim$(CStr(ws.Range(CELL_PROJECT_PART).Value))
    PROJECT_ENGINEER = Trim$(CStr(ws.Range(CELL_PROJECT_ENGINEER).Value))
    PROJECT_REVISION = Trim$(CStr(ws.Range(CELL_PROJECT_REVISION).Value))
    On Error GoTo 0

    s = pType
    If Len(pNo) > 0 Then s = s & IIf(Len(s) > 0, " ", "") & pNo
    If Len(pPart) > 0 Then s = s & IIf(Len(s) > 0, " - ", "") & pPart

    PROJECT_NAME = s

End Sub

Private Sub GenerateGeometry()

    Dim dx1 As Double, dx2 As Double, dz2 As Double
    Dim sHaunchSlab As Long, sHaunchWall As Long, sRigidFound As Long
    Dim n As String, e As String

    ' Haunch members fall back to the plain slab/wall section when their
    ' gate dimension is zero. NOTE the cross-gating, straight from the MCT:
    ' the slab-haunch members switch on B8 and the wall-haunch ones on B9.
    sHaunchSlab = IIf(DIM_SLAB_HAUNCH > 0, 5, 1)
    sHaunchWall = IIf(DIM_WALL_HAUNCH > 0, 6, 2)
    sRigidFound = IIf(DIM_SECT9 > 0, 9, 4)

    HAS_EXT = (DIM_EXT > 0)

    If HAS_EXT Then

        SLAB_E = 16
        WALL_E = 8

        dx1 = DIM_EXT + DIM_WALL_T / 2
        dx2 = DIM_EXT + DIM_WALL_T + DIM_SPAN + DIM_WALL_T / 2
        dz2 = DIM_FOUND_T / 2 + DIM_WALL_H + DIM_SLAB_T / 2

        n = NodeRow(1, 0, 0)
        n = n & ";" & NodeRow(2, DIM_EXT, 0)
        n = n & ";" & NodeRow(3, dx1, 0)
        n = n & ";" & NodeRow(4, DIM_EXT + DIM_WALL_T, 0)
        n = n & ";" & NodeRow(5, DIM_EXT + DIM_WALL_T + DIM_SPAN, 0)
        n = n & ";" & NodeRow(6, dx2, 0)
        n = n & ";" & NodeRow(7, DIM_EXT + DIM_WALL_T + DIM_SPAN + DIM_WALL_T, 0)
        n = n & ";" & NodeRow(8, DIM_EXT + DIM_WALL_T + DIM_SPAN + DIM_WALL_T + DIM_EXT, 0)
        n = n & ";" & NodeRow(9, dx1, DIM_FOUND_T / 2)
        n = n & ";" & NodeRow(10, dx2, DIM_FOUND_T / 2)
        n = n & ";" & NodeRow(11, dx1, DIM_FOUND_T / 2 + DIM_WALL_H - DIM_SLAB_HAUNCH)
        n = n & ";" & NodeRow(12, dx2, DIM_FOUND_T / 2 + DIM_WALL_H - DIM_SLAB_HAUNCH)
        n = n & ";" & NodeRow(13, dx1, DIM_FOUND_T / 2 + DIM_WALL_H)
        n = n & ";" & NodeRow(14, dx2, DIM_FOUND_T / 2 + DIM_WALL_H)
        n = n & ";" & NodeRow(15, dx1, dz2)
        n = n & ";" & NodeRow(16, dx1 + DIM_WALL_T / 2, dz2)
        n = n & ";" & NodeRow(17, dx1 + DIM_WALL_T / 2 + DIM_WALL_HAUNCH, dz2)
        n = n & ";" & NodeRow(18, dx1 + DIM_WALL_T / 2 + DIM_SPAN - DIM_WALL_HAUNCH, dz2)
        n = n & ";" & NodeRow(19, dx1 + DIM_WALL_T / 2 + DIM_SPAN, dz2)
        n = n & ";" & NodeRow(20, dx1 + DIM_WALL_T + DIM_SPAN, dz2)

        e = ElemRow(1, 4, 1, 2, 0)
        e = e & ";" & ElemRow(2, sRigidFound, 2, 3, 0)
        e = e & ";" & ElemRow(3, sRigidFound, 3, 4, 0)
        e = e & ";" & ElemRow(4, 4, 4, 5, 0)
        e = e & ";" & ElemRow(5, sRigidFound, 5, 6, 0)
        e = e & ";" & ElemRow(6, sRigidFound, 6, 7, 0)
        e = e & ";" & ElemRow(7, 4, 7, 8, 0)
        e = e & ";" & ElemRow(8, 8, 9, 3, -180)
        e = e & ";" & ElemRow(9, 8, 10, 6, 0)
        e = e & ";" & ElemRow(10, 2, 11, 9, -180)
        e = e & ";" & ElemRow(11, 2, 12, 10, 0)
        e = e & ";" & ElemRow(12, sHaunchWall, 13, 11, -180)
        e = e & ";" & ElemRow(13, sHaunchWall, 14, 12, 0)
        e = e & ";" & ElemRow(14, 8, 15, 13, -180)
        e = e & ";" & ElemRow(15, 8, 20, 14, 0)
        e = e & ";" & ElemRow(16, 7, 15, 16, 0)
        e = e & ";" & ElemRow(17, sHaunchSlab, 16, 17, 0)
        e = e & ";" & ElemRow(18, 1, 17, 18, 0)
        e = e & ";" & ElemRow(19, sHaunchSlab, 18, 19, 0)
        e = e & ";" & ElemRow(20, 7, 19, 20, 0)

    Else

        SLAB_E = 16
        WALL_E = 8

        dx1 = DIM_WALL_T / 2
        dx2 = DIM_WALL_T + DIM_SPAN + DIM_WALL_T / 2
        dz2 = DIM_FOUND_T / 2 + DIM_WALL_H + DIM_SLAB_T / 2

        n = NodeRow(3, dx1, 0)
        n = n & ";" & NodeRow(4, DIM_WALL_T, 0)
        n = n & ";" & NodeRow(5, DIM_WALL_T + DIM_SPAN, 0)
        n = n & ";" & NodeRow(6, dx2, 0)
        n = n & ";" & NodeRow(9, dx1, DIM_FOUND_T / 2)
        n = n & ";" & NodeRow(10, dx2, DIM_FOUND_T / 2)
        n = n & ";" & NodeRow(11, dx1, DIM_FOUND_T / 2 + DIM_WALL_H - DIM_SLAB_HAUNCH)
        n = n & ";" & NodeRow(12, dx2, DIM_FOUND_T / 2 + DIM_WALL_H - DIM_SLAB_HAUNCH)
        n = n & ";" & NodeRow(13, dx1, DIM_FOUND_T / 2 + DIM_WALL_H)
        n = n & ";" & NodeRow(14, dx2, DIM_FOUND_T / 2 + DIM_WALL_H)
        n = n & ";" & NodeRow(15, dx1, dz2)
        n = n & ";" & NodeRow(16, dx1 + DIM_WALL_T / 2, dz2)
        n = n & ";" & NodeRow(17, dx1 + DIM_WALL_T / 2 + DIM_WALL_HAUNCH, dz2)
        n = n & ";" & NodeRow(18, dx1 + DIM_WALL_T / 2 + DIM_SPAN - DIM_WALL_HAUNCH, dz2)
        n = n & ";" & NodeRow(19, dx1 + DIM_WALL_T / 2 + DIM_SPAN, dz2)
        n = n & ";" & NodeRow(20, dx1 + DIM_WALL_T + DIM_SPAN, dz2)

        e = ElemRow(3, sRigidFound, 3, 4, 0)
        e = e & ";" & ElemRow(4, 4, 4, 5, 0)
        e = e & ";" & ElemRow(5, sRigidFound, 5, 6, 0)
        e = e & ";" & ElemRow(8, 8, 9, 3, -180)
        e = e & ";" & ElemRow(9, 8, 10, 6, 0)
        e = e & ";" & ElemRow(10, 2, 11, 9, -180)
        e = e & ";" & ElemRow(11, 2, 12, 10, 0)
        e = e & ";" & ElemRow(12, sHaunchWall, 13, 11, -180)
        e = e & ";" & ElemRow(13, sHaunchWall, 14, 12, 0)
        e = e & ";" & ElemRow(14, 8, 15, 13, -180)
        e = e & ";" & ElemRow(15, 8, 20, 14, 0)
        e = e & ";" & ElemRow(16, 7, 15, 16, 0)
        e = e & ";" & ElemRow(17, sHaunchSlab, 16, 17, 0)
        e = e & ";" & ElemRow(18, 1, 17, 18, 0)
        e = e & ";" & ElemRow(19, sHaunchSlab, 18, 19, 0)
        e = e & ";" & ElemRow(20, 7, 19, 20, 0)

    End If

    NODE_LIST = n
    ELEM_LIST = e

End Sub

Private Function NodeRow(ByVal no As Long, ByVal x As Double, ByVal z As Double) As String
    NodeRow = no & "|" & JsonNum(x) & "|0|" & JsonNum(z)
End Function

Private Function ElemRow(ByVal no As Long, ByVal sect As Long, ByVal n1 As Long, _
                         ByVal n2 As Long, ByVal angle As Double) As String
    ElemRow = no & "|" & sect & "|" & n1 & "|" & n2 & "|" & JsonNum(angle)
End Function

Private Sub ComputeOrdinates()

    Call EhBlock(LD_EHS1_TOP, LD_EHS1_BOT, ORD_EHS1)
    Call EhBlock(LD_EHA1_TOP, LD_EHA1_BOT, ORD_EHA1)
    Call EhBlock(LD_EHS2_TOP, LD_EHS2_BOT, ORD_EHS2)
    Call EhBlock(LD_EHA2_TOP, LD_EHA2_BOT, ORD_EHA2)
    Call EqBlock(ORD_EQ)

End Sub

Private Sub EhBlock(ByVal pTop As Double, ByVal pBot As Double, ByRef o() As Double)

    Dim l As Double

    l = ((-DIM_SLAB_T / 2 - DIM_FOUND_T / 2 - DIM_WALL_H) * pBot) / (pTop - pBot)

    o(1) = pBot
    o(2) = pBot * (l - DIM_FOUND_T / 2) / l
    o(3) = o(2)
    o(4) = pBot * (l - DIM_FOUND_T / 2 - DIM_WALL_H + DIM_SLAB_HAUNCH) / l
    o(5) = o(4)
    o(6) = pBot * (l - DIM_FOUND_T / 2 - DIM_WALL_H) / l
    o(7) = o(6)
    o(8) = pBot * (l - DIM_FOUND_T / 2 - DIM_WALL_H - DIM_SLAB_T / 2) / l

End Sub

Private Sub EqBlock(ByRef o() As Double)

    Dim zTop As Double

    zTop = DIM_FOUND_T / 2 + DIM_WALL_H + DIM_SLAB_T / 2

    o(1) = EqAt(zTop, zTop)
    o(2) = EqAt(DIM_FOUND_T / 2 + DIM_WALL_H, zTop)
    o(3) = o(2)
    o(4) = EqAt(DIM_FOUND_T / 2 + DIM_WALL_H - DIM_SLAB_HAUNCH, zTop)
    o(5) = o(4)
    o(6) = EqAt(DIM_FOUND_T / 2, zTop)
    o(7) = o(6)
    o(8) = EqAt(0, zTop)

End Sub

Private Function EqAt(ByVal z As Double, ByVal zTop As Double) As Double

    If zTop = 0 Then
        EqAt = LD_EQ_BOT
    Else
        EqAt = LD_EQ_BOT + (LD_EQ_TOP - LD_EQ_BOT) * z / zTop
    End If

End Function

Private Sub GenerateBeamLoads()

    BEAMLOAD_LIST = ""

    ' --- Vertical loads on the 5 top-slab elements (global Z) ---
    Call AddSlabLoads("EV1", LD_EV1)
    Call AddSlabLoads("EV2", LD_EV2)
    Call AddSlabLoads("LL1", LD_LL1)
    Call AddSlabLoads("LL", LD_LL)
    Call AddSlabLoads("LLacc", LD_LLACC)

    ' EV1/EV2 also load the two outer foundation members when the
    ' foundation extends past the walls (elements 1 and 7 only).
    If HAS_EXT Then
        Call AddLoad(1, "EV1", "BEAM", "GZ", -LD_EV1_EXT, -LD_EV1_EXT)
        Call AddLoad(7, "EV1", "BEAM", "GZ", -LD_EV1_EXT, -LD_EV1_EXT)
        Call AddLoad(1, "EV2", "BEAM", "GZ", -LD_EV2_EXT, -LD_EV2_EXT)
        Call AddLoad(7, "EV2", "BEAM", "GZ", -LD_EV2_EXT, -LD_EV2_EXT)
    End If

    ' --- Uniform surcharge on the LEFT wall only (local Z) ---
    Call AddLeftWallLoads("LSS1_L", LD_LSS1_L)
    Call AddLeftWallLoads("LSS2_L", LD_LSS2_L)
    Call AddLeftWallLoads("LSA2_L", LD_LSA2_L)

    ' --- Trapezoidal earth pressure, both walls (local Z) ---
    ' EHS1 and EHA1 put BOTH sides into one load case; the "2" pair splits
    ' into _L and _R cases. Straight from the MCT module.
    Call AddWallProfile("EHS1", "EHS1", ORD_EHS1)
    Call AddWallProfile("EHA1", "EHA1", ORD_EHA1)
    Call AddWallProfile("EHS2_L", "EHS2_R", ORD_EHS2)
    Call AddWallProfile("EHA2_L", "EHA2_R", ORD_EHA2)

    ' Seismic ground pressure - skipped with the rest of the seismic
    ' chain when this culvert is buried too deep to need it.
    If SEISMIC_ACTIVE Then Call AddEqLoads

End Sub

Private Sub AddSlabLoads(ByVal lcname As String, ByVal v As Double)
    Dim j As Long
    For j = 0 To 4
        Call AddLoad(SLAB_E + j, lcname, "BEAM", "GZ", -v, -v)
    Next j
End Sub

Private Sub AddLeftWallLoads(ByVal lcname As String, ByVal v As Double)
    Dim j As Long
    For j = 0 To 3
        Call AddLoad(WALL_E + j * 2, lcname, "BEAM", "LZ", -v, -v)
    Next j
End Sub

Private Sub AddWallProfile(ByVal lcLeft As String, ByVal lcRight As String, ByRef o() As Double)

    Dim p As Long

    For p = 1 To 4
        Call AddLoad(WALL_E + (p - 1) * 2, lcLeft, "LINE", "LZ", -o(p * 2), -o(p * 2 - 1))
    Next p

    For p = 1 To 4
        Call AddLoad(WALL_E + (p - 1) * 2 + 1, lcRight, "LINE", "LZ", -o(p * 2), -o(p * 2 - 1))
    Next p

End Sub

Private Sub AddEqLoads()

    Dim p As Long

    For p = 1 To 4
        Call AddLoad(WALL_E + 12 - (4 + p * 2), "EQ", "LINE", "LZ", _
                     -ORD_EQ(p * 2 - 1), -ORD_EQ(p * 2))
    Next p

End Sub

Private Sub AddLoad(ByVal elemNo As Long, ByVal lcname As String, ByVal cmd As String, _
                    ByVal loadDir As String, ByVal p1 As Double, ByVal p2 As Double)

    If Len(BEAMLOAD_LIST) > 0 Then BEAMLOAD_LIST = BEAMLOAD_LIST & ";"

    BEAMLOAD_LIST = BEAMLOAD_LIST & elemNo & "|" & lcname & "|" & cmd & "|" & loadDir & _
                    "|" & JsonNum(Round(p1, LOAD_ROUND_DP)) & _
                    "|" & JsonNum(Round(p2, LOAD_ROUND_DP))

End Sub

Private Sub GenerateLoadCombinations()

    Dim i As Long
    Dim serSpec As String, strSpec As String

    LOADCOMB_LIST = ""

    ' --- Serviceability ---
    Call AddCombo("SLS-1", 0, "ST", "DL:1,EV1:1,EHS1:1,LL1:1,LSS1_L:1", 1)
    Call AddCombo("SLS-2", 0, "ST", "DL:1,EV1:1,EHS1:1", 1)
    Call AddCombo("SLS-3", 0, "ST", "DL:1,EV1:1,EHS1:0.5,LL1:1,LSS1_L:1", 1)
    Call AddCombo("SLS-4", 0, "ST", "DL:1,EV1:1,EHS1:0.5", 1)
    Call AddCombo("SLS-5", 0, "ST", "DL:1,EV2:1,EHS2_L:1,EHS2_R:1", 1)
    Call AddCombo("SLS-6", 0, "ST", "DL:1,EV2:1,EHA2_L:1,EHA2_R:1", 1)
    Call AddCombo("SLS-7", 0, "ST", "DL:1,EV2:1,EHS2_L:1,EHS2_R:1,LL:1", 1)
    Call AddCombo("SLS-8", 0, "ST", "DL:1,EV2:1,EHA2_L:1,EHA2_R:1,LL:1", 1)
    Call AddCombo("SLS-9", 0, "ST", "DL:1,EV2:1,EHS2_L:1,EHS2_R:1,LSS2_L:1", 1)
    Call AddCombo("SLS-10", 0, "ST", "DL:1,EV2:1,EHA2_L:1,EHA2_R:1,LSA2_L:1", 1)
    Call AddCombo("SLS-11", 0, "ST", "DL:1,EV2:1,EHS2_L:1,EHS2_R:1,LL:1,LSS2_L:1", 1)
    Call AddCombo("SLS-12", 0, "ST", "DL:1,EV2:1,EHA2_L:1,EHA2_R:1,LL:1,LSA2_L:1", 1)
    Call AddCombo("SLS-13", 0, "ST", "DL:1,EV2:1,EHS2_L:0.5,EHS2_R:0.5,LL:1", 1)
    Call AddCombo("SLS-14", 0, "ST", "DL:1,EV2:1,EHA2_L:0.5,EHA2_R:0.5,LL:1", 1)

    ' --- Ultimate ---
    Call AddCombo("ULS-1", 0, "ST", "DL:1.35,EV1:1.35,EHS1:1.5,LL1:1.45,LSS1_L:1.45", 1)
    Call AddCombo("ULS-2", 0, "ST", "DL:1.35,EV1:1.35,EHS1:1.5", 1)
    Call AddCombo("ULS-3", 0, "ST", "DL:1.35,EV1:1.35,EHS1:0.75,LL1:1.45,LSS1_L:1.45", 1)
    Call AddCombo("ULS-4", 0, "ST", "DL:1.35,EV1:1.35,EHS1:0.75", 1)
    Call AddCombo("ULS-5", 0, "ST", "DL:1.35,EV2:1.35,EHS2_L:1.35,EHS2_R:1.35", 1)
    Call AddCombo("ULS-6", 0, "ST", "DL:1.35,EV2:1.35,EHA2_L:1.35,EHA2_R:1.35", 1)
    Call AddCombo("ULS-7", 0, "ST", "DL:1.35,EV2:1.35,EHS2_L:1.35,EHS2_R:1.35,LL:1.45", 1)
    Call AddCombo("ULS-8", 0, "ST", "DL:1.35,EV2:1.35,EHA2_L:1.35,EHA2_R:1.35,LL:1.45", 1)
    Call AddCombo("ULS-9", 0, "ST", "DL:1.35,EV2:1.35,EHS2_L:1.35,EHS2_R:1.35,LSS2_L:1.45", 1)
    Call AddCombo("ULS-10", 0, "ST", "DL:1.35,EV2:1.35,EHA2_L:1.35,EHA2_R:1.35,LSA2_L:1.45", 1)
    Call AddCombo("ULS-11", 0, "ST", _
        "DL:1.35,EV2:1.35,EHS2_L:1.35,EHS2_R:1.35,LL:1.45,LSS2_L:1.45", 1)
    Call AddCombo("ULS-12", 0, "ST", _
        "DL:1.35,EV2:1.35,EHA2_L:1.35,EHA2_R:1.35,LL:1.45,LSA2_L:1.45", 1)
    Call AddCombo("ULS-13", 0, "ST", "DL:1.35,EV2:1.35,EHS2_L:0.75,EHS2_R:0.75,LL:1.45", 1)
    Call AddCombo("ULS-14", 0, "ST", "DL:1.35,EV2:1.35,EHA2_L:0.75,EHA2_R:0.75,LL:1.45", 1)

    ' --- Accidental and seismic ---
    ' ACC-1 is an accidental (impact) case, not a seismic one, so it is
    ' built regardless. EQ-1 goes only when the seismic gate is on.
    '
    ' EQ-1 deliberately does NOT reference ATA (by request, 2026-09-23) -
    ' only the EQ lateral earth pressure case. ATA is still built as a
    ' static load case (db/STLD) and still carries its self-weight
    ' inertia record (db/BODF) whenever SEISMIC_ACTIVE, per the
    ' SEISMIC GATE block above - it is simply not pulled into any
    ' combination any more, EQ-1 or otherwise.
    Call AddCombo("ACC-1", 0, "ST", "DL:1,EV2:1,EHA2_L:1,EHA2_R:1,LLacc:1", 1)
    If SEISMIC_ACTIVE Then
        Call AddCombo("EQ-1", 0, "ST", "DL:1,EV2:1,EHA2_L:1,EHA2_R:1,EQ:1", 1)
    End If

    ' --- Envelopes ---
    For i = 1 To 14
        serSpec = serSpec & IIf(i > 1, ",", "") & "SLS-" & i & ":1"
        strSpec = strSpec & IIf(i > 1, ",", "") & "ULS-" & i & ":1"
    Next i

    Call AddCombo("ENV_SER", 1, "CB", serSpec, 2)
    Call AddCombo("ENV_STR", 1, "CB", strSpec, 2)

    If SEISMIC_ACTIVE Then
        Call AddCombo("ENV_ALL", 1, "CB", "EQ-1:1,ENV_SER:1,ENV_STR:1", 3)
    Else
        Call AddCombo("ENV_ALL", 1, "CB", "ENV_SER:1,ENV_STR:1", 3)
    End If

    Call AddCombo("ENV_DEAD", 1, "CB", "SLS-2:1,SLS-4:1,SLS-5:1,SLS-6:1", 2)

    ' ENV_EQ, the seismic envelope the displacement capture reads. Last, so
    ' every other combination keeps the key it always had; seismic only,
    ' like EQ-1 itself (the capture reports it SKIPPED otherwise).
    If SEISMIC_ACTIVE Then
        Call AddCombo("ENV_EQ", 1, "CB", "EQ-1:1", 2)
    End If

End Sub

Private Sub AddCombo(ByVal nm As String, ByVal iType As Long, _
                     ByVal anal As String, ByVal spec As String, _
                     ByVal tier As Long)

    Dim pairs() As String
    Dim kv() As String
    Dim i As Long
    Dim factors As String

    pairs = Split(spec, ",")

    For i = LBound(pairs) To UBound(pairs)
        kv = Split(pairs(i), ":")
        factors = factors & IIf(i > LBound(pairs), ",", "") & _
                  anal & ":" & kv(0) & ":" & kv(1)
    Next i

    If Len(LOADCOMB_LIST) > 0 Then LOADCOMB_LIST = LOADCOMB_LIST & ";"
    LOADCOMB_LIST = LOADCOMB_LIST & nm & "|ACTIVE|" & iType & "|" & factors & _
                    "|" & tier

End Sub

Private Function IsSeismicCase(ByVal nm As String) As Boolean
    IsSeismicCase = (nm = "EQ" Or nm = "ATA")
End Function

Private Function StepResult(ByRef report As String, ByVal stepName As String, _
                            ByVal result As String) As Boolean
    If Len(result) = 0 Then
        report = report & "OK   - " & stepName & vbCrLf
        StepResult = True
    ElseIf Left$(result, 5) = "WARN:" Then
        report = report & "WARN - " & stepName & ":" & Mid$(result, 6) & vbCrLf
        StepResult = True
    Else
        report = report & "FAIL - " & stepName & ": " & result & vbCrLf
        StepResult = False
    End If
End Function

Private Function JsonNum(ByVal v As Double) As String

    Dim s As String
    s = Trim$(Str$(v))

    If Left$(s, 1) = "." Then
        s = "0" & s
    ElseIf Left$(s, 2) = "-." Then
        s = "-0" & Mid$(s, 2)
    End If

    JsonNum = s

End Function

Private Function VerdictFirst(ByVal report As String, ByVal ok As Boolean) As String

    Dim verdict As String
    Dim p As Long

    If ok Then
        verdict = "All steps completed."
    Else
        p = InStrRev(report, "FAIL - ")
        If p > 0 Then
            verdict = "STOPPED - " & Mid$(report, p) & "Fix it and re-run."
        Else
            verdict = "STOPPED after a failed step - fix it and re-run."
        End If
    End If

    VerdictFirst = Replace(report, vbCrLf, vbCrLf & verdict & vbCrLf & vbCrLf, 1, 1)

End Function

Private Function FitReport(ByVal report As String, ByVal logName As String) As String

    Const MAX_LEN As Long = 900
    Dim path As String
    Dim fileNo As Integer

    If Len(report) <= MAX_LEN Then
        FitReport = report
        Exit Function
    End If

    If Len(ThisWorkbook.Path) > 0 Then
        path = ThisWorkbook.Path & "\" & logName & "_log.txt"
    Else
        path = Environ$("TEMP") & "\" & logName & "_log.txt"
    End If

    On Error Resume Next
    fileNo = FreeFile
    Open path For Output As #fileNo
    Print #fileNo, report
    Close #fileNo
    If Err.Number <> 0 Then path = "(could not be written: " & Err.Description & ")"
    Err.Clear
    On Error GoTo 0

    FitReport = Left$(report, MAX_LEN) & vbCrLf & "..." & vbCrLf & "Full report: " & path

End Function


' ===========================================================================
'  DIVIDE + RENUMBER
'  What ope/DIVIDEELEM does in the MIDAS builder, done on the lists:
'  elements 18, 10, 11 and 4 are split into DIVIDE_AMOUNT equal frames.
'  Original nodes become joints 1..m in NODE_LIST order; division joints
'  follow. Frames are numbered 1..n in ELEM_LIST order, a divided element's
'  pieces consecutively from its I end. Each piece keeps the element's
'  section and angle.
'
'  Numbers come back out of NODE_LIST with Val, never CDbl: the lists hold
'  period decimals (JsonNum), and CDbl reads "0.225" as 225 on a Turkish
'  Windows.
' ===========================================================================

Private Function ExpandModel() As String

    Dim nodes() As String
    Dim elems() As String
    Dim f() As String
    Dim i As Long, k As Long
    Dim maxNode As Long, maxElem As Long
    Dim nodeMap() As Long
    Dim cap As Long
    Dim elemNo As Long, pieces As Long
    Dim ni As Long, nj As Long
    Dim jPrev As Long, jNext As Long
    Dim t As Double

    nodes = Split(NODE_LIST, ";")
    elems = Split(ELEM_LIST, ";")

    For i = 0 To UBound(nodes)
        f = Split(nodes(i), "|")
        If CLng(f(0)) > maxNode Then maxNode = CLng(f(0))
    Next i
    For i = 0 To UBound(elems)
        f = Split(elems(i), "|")
        If CLng(f(0)) > maxElem Then maxElem = CLng(f(0))
    Next i

    ReDim nodeMap(0 To maxNode)
    ReDim ELEM_FIRST_FRAME(0 To maxElem)
    ReDim ELEM_PIECES(0 To maxElem)

    cap = (UBound(nodes) + 1) + (UBound(elems) + 1) * IIf(DIVIDE_AMOUNT > 1, DIVIDE_AMOUNT, 1)
    ReDim JT_X(1 To cap)
    ReDim JT_Z(1 To cap)
    ReDim JT_SPECIAL(1 To cap)
    ReDim FR_I(1 To cap)
    ReDim FR_J(1 To cap)
    ReDim FR_SECT(1 To cap)
    ReDim FR_ANGLE(1 To cap)

    JT_COUNT = 0
    For i = 0 To UBound(nodes)
        f = Split(nodes(i), "|")
        JT_COUNT = JT_COUNT + 1
        JT_X(JT_COUNT) = Val(f(1))
        JT_Z(JT_COUNT) = Val(f(3))
        JT_SPECIAL(JT_COUNT) = True
        nodeMap(CLng(f(0))) = JT_COUNT
    Next i

    DIVIDED_ELEMS = ""
    FR_COUNT = 0
    For i = 0 To UBound(elems)

        f = Split(elems(i), "|")
        elemNo = CLng(f(0))
        ni = nodeMap(CLng(f(2)))
        nj = nodeMap(CLng(f(3)))
        If ni = 0 Or nj = 0 Then
            ExpandModel = "element " & elemNo & " references a node that was not generated."
            Exit Function
        End If

        pieces = 1
        If DIVIDE_AMOUNT > 1 And IsDivideTarget(elemNo) Then
            pieces = DIVIDE_AMOUNT
            DIVIDED_ELEMS = DIVIDED_ELEMS & IIf(Len(DIVIDED_ELEMS) > 0, ", ", "") & elemNo
        End If

        ELEM_FIRST_FRAME(elemNo) = FR_COUNT + 1
        ELEM_PIECES(elemNo) = pieces

        jPrev = ni
        For k = 1 To pieces
            If k = pieces Then
                jNext = nj
            Else
                t = k / pieces
                JT_COUNT = JT_COUNT + 1
                JT_X(JT_COUNT) = JT_X(ni) + (JT_X(nj) - JT_X(ni)) * t
                JT_Z(JT_COUNT) = JT_Z(ni) + (JT_Z(nj) - JT_Z(ni)) * t
                JT_SPECIAL(JT_COUNT) = False
                jNext = JT_COUNT
            End If
            FR_COUNT = FR_COUNT + 1
            FR_I(FR_COUNT) = jPrev
            FR_J(FR_COUNT) = jNext
            FR_SECT(FR_COUNT) = CLng(f(1))
            FR_ANGLE(FR_COUNT) = Val(f(4))
            jPrev = jNext
        Next k

    Next i

    If Len(DIVIDED_ELEMS) = 0 Then
        DIVIDED_ELEMS = "none"
        ExpandModel = "WARN: nothing divided - " & PROJECT_SHEET_NAME & "!" & _
                      CELL_DIVIDE_AMOUNT & " is " & DIVIDE_AMOUNT & " (must be >= 2)."
    End If

End Function

Private Function IsDivideTarget(ByVal elemNo As Long) As Boolean
    IsDivideTarget = (InStr(1, "," & DIVIDE_TARGETS & ",", "," & elemNo & ",", vbBinaryCompare) > 0)
End Function

Private Function FrameLength(ByVal fr As Long) As Double
    FrameLength = Sqr((JT_X(FR_J(fr)) - JT_X(FR_I(fr))) ^ 2 + _
                      (JT_Z(FR_J(fr)) - JT_Z(FR_I(fr))) ^ 2)
End Function


' ===========================================================================
'  MODEL TEXT
'  Tables in the order SAP2000's own export uses (see the reference file).
'  Each table is its header, its rows, then a line holding one space.
' ===========================================================================

Private Function WriteModelText() As String

    Dim res As String

    OUT_COUNT = 0
    ReDim OUT_LINES(1 To 512)
    TEXT_WARNINGS = ""
    LOAD_COUNT = 0
    SPRING_COUNT = 0

    Call Emit("File " & SapOutputPath() & " was saved on " & Month(Now) & "." & Day(Now) & _
              "." & Format$(Year(Now) Mod 100, "00") & " at " & Format$(Now, "hh:mm:ss"))
    Call Emit(" ")

    ' XZ plane frame, as MIDAS's STYP X-Z: springs at z = 0 then leave no
    ' mechanism (with all six DOF the frame could spin about global X).
    Call TableStart("ACTIVE DEGREES OF FREEDOM")
    Call Emit("   UX=Yes   UY=No   UZ=Yes   RX=No   RY=Yes   RZ=No")
    Call TableEnd

    Call TableStart("ANALYSIS OPTIONS")
    Call Emit("   Solver=Advanced   SolverProc=Auto   Force32Bit=No   StiffCase=None   GeomMod=None")
    Call TableEnd

    Call TableStart("CASE - MODAL 1 - GENERAL")
    Call Emit("   Case=MODAL   ModeType=Eigen   MaxNumModes=12   MinNumModes=1   EigenShift=0" & _
              "   EigenCutoff=0   EigenTol=1E-09   AutoShift=Yes")
    Call TableEnd

    Call WriteStaticCaseTable

    res = WriteCombinationTable()
    If Len(res) > 0 Then WriteModelText = res: Exit Function

    Call WriteConnectivityTable

    Call TableStart("COORDINATE SYSTEMS")
    Call Emit("   Name=GLOBAL   Type=Cartesian   X=0   Y=0   Z=0   AboutZ=0   AboutY=0   AboutX=0")
    Call TableEnd

    Call TableStart("DATABASE FORMAT TYPES")
    Call Emit("   UnitsCurr=Yes   OverrideE=No")
    Call TableEnd

    Call WritePerFrameTable("FRAME AUTO MESH ASSIGNMENTS", _
        "   AutoMesh=Yes   AtJoints=Yes   AtFrames=No   NumSegments=0   MaxLength=0   MaxDegrees=0")
    Call WritePerFrameTable("FRAME DESIGN PROCEDURES", "   DesignProc=""From Material""")

    res = WriteDistributedLoadTable()
    If Len(res) > 0 Then WriteModelText = res: Exit Function
    Call WriteGravityLoadTable

    Call WritePerFrameTable("FRAME LOAD TRANSFER OPTIONS", "   Transfer=Yes")
    Call WriteLocalAxesTable
    Call WritePerFrameTable("FRAME OUTPUT STATION ASSIGNMENTS", _
        "   StationType=MaxStaSpcg   MaxStaSpcg=0.5   AddAtElmInt=Yes   AddAtPtLoad=Yes")

    res = WriteSectionTables()
    If Len(res) > 0 Then WriteModelText = res: Exit Function

    Call TableStart("GROUPS 1 - DEFINITIONS")
    Call Emit("   GroupName=ALL   Selection=Yes   SectionCut=Yes   Steel=Yes   Concrete=Yes" & _
              "   Aluminum=Yes   ColdFormed=Yes   Stage=Yes   Bridge=Yes   AutoSeismic=No" & _
              "   AutoWind=No   SelDesSteel=No   SelDesAlum=No   SelDesCold=No   MassWeight=Yes   Color=Red")
    Call TableEnd

    Call WriteJointTable

    res = WriteSpringTable()
    If Len(res) > 0 Then WriteModelText = res: Exit Function

    Call WriteLoadCaseTables
    Call WriteMaterialTables

    Call TableStart("PREFERENCES - DIMENSIONAL")
    Call Emit("   MergeTol=0.001   FineGrid=0.25   Nudge=0.25   SelectTol=3   SnapTol=12" & _
              "   SLineThick=2   PLineThick=4   MaxFont=8   MinFont=3   AutoZoom=10" & _
              "   ShrinkFact=70   TextFileLen=240")
    Call TableEnd

    Call TableStart("PROGRAM CONTROL")
    Call Emit("   ProgramName=SAP2000   Version=" & SAP_PROGRAM_VERSION & _
              "   CurrUnits=""KN, m, C""   SteelCode=""AISC 360-10""   ConcCode=""ACI 318-14""" & _
              "   AlumCode=""AA-ASD 2000""   ColdCode=AISI-ASD96   RegenHinge=Yes")
    Call TableEnd

    Call WriteProjectInfoTable
    Call WriteRebarSizeTable

    Call Emit("END TABLE DATA")

    If Len(TEXT_WARNINGS) > 0 Then WriteModelText = "WARN:" & TEXT_WARNINGS

End Function

' One row per load pattern (the MIDAS static load cases, less EQ/ATA when
' the seismic gate is off).
Private Sub WriteStaticCaseTable()

    Dim rows() As String
    Dim f() As String
    Dim i As Long

    Call TableStart("CASE - STATIC 1 - LOAD ASSIGNMENTS")
    rows = Split(STLDCASE_LIST, ";")
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
        If SEISMIC_ACTIVE Or Not IsSeismicCase(f(0)) Then
            Call Emit("   Case=" & SapName(f(0)) & "   LoadType=""Load pattern""   LoadName=" & _
                      SapName(f(0)) & "   LoadSF=1")
        End If
    Next i
    Call TableEnd

End Sub

' "ST" entries name a load pattern (SapName drops the "_"); "CB" entries name
' another combination, kept as MIDAS spells it. LOADCOMB_LIST is already in
' dependency order (ENV_* after the combinations they envelope).
Private Function WriteCombinationTable() As String

    Dim rows() As String
    Dim f() As String
    Dim items() As String
    Dim it() As String
    Dim i As Long, j As Long
    Dim caseName As String
    Dim s As String

    If Len(LOADCOMB_LIST) = 0 Then
        WriteCombinationTable = "no load combinations were generated."
        Exit Function
    End If

    Call TableStart("COMBINATION DEFINITIONS")
    rows = Split(LOADCOMB_LIST, ";")
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
        items = Split(f(3), ",")
        For j = 0 To UBound(items)
            it = Split(items(j), ":")
            If it(0) = "ST" Then
                caseName = SapName(it(1))
            ElseIf it(0) = "CB" Then
                caseName = it(1)
            Else
                WriteCombinationTable = f(0) & ": unknown reference type """ & it(0) & """."
                Exit Function
            End If
            s = "   ComboName=" & SapQuote(f(0))
            If j = 0 Then
                s = s & "   ComboType=" & IIf(f(2) = "1", "Envelope", """Linear Add""") & _
                    "   AutoDesign=No"
            End If
            s = s & "   CaseName=" & SapQuote(caseName) & "   ScaleFactor=" & it(2)
            If j = 0 Then
                s = s & "   SteelDesign=None   ConcDesign=None   AlumDesign=None   ColdDesign=None"
            End If
            Call Emit(s)
        Next j
    Next i
    Call TableEnd

End Function

Private Sub WriteConnectivityTable()

    Dim fr As Long

    Call TableStart("CONNECTIVITY - FRAME")
    For fr = 1 To FR_COUNT
        Call Emit("   Frame=" & fr & "   JointI=" & FR_I(fr) & "   JointJ=" & FR_J(fr) & "   IsCurved=No")
    Next fr
    Call TableEnd

End Sub

' The same fields on every frame.
Private Sub WritePerFrameTable(ByVal tableName As String, ByVal fields As String)

    Dim fr As Long

    Call TableStart(tableName)
    For fr = 1 To FR_COUNT
        Call Emit("   Frame=" & fr & fields)
    Next fr
    Call TableEnd

End Sub

' BEAMLOAD_LIST rows are "elem|case|cmd|dir|p1|p2", p1 at the element's I
' end and p2 at its J end over the full length (MIDAS D = [0,1]). A divided
' element's pieces each get their slice of that trapezoid.
'   GZ (global Z, MIDAS writes -v)  ->  CoordSys=GLOBAL  Dir=Gravity  +v
'   LZ (member local z)             ->  CoordSys=Local   Dir=2        same sign
' The LZ mapping is the reference file's: its walls run top-down like the
' MIDAS ones, the left wall with Angle=180, and carry the MIDAS values
' unchanged in Dir=2.
Private Function WriteDistributedLoadTable() As String

    Dim rows() As String
    Dim f() As String
    Dim i As Long, k As Long
    Dim elemNo As Long, pieces As Long, fr As Long
    Dim p1 As Double, p2 As Double, va As Double, vb As Double
    Dim loadSign As Double
    Dim coordSys As String, loadDir As String

    If Len(BEAMLOAD_LIST) = 0 Then
        WriteDistributedLoadTable = "no beam loads were generated."
        Exit Function
    End If

    Call TableStart("FRAME LOADS - DISTRIBUTED")
    rows = Split(BEAMLOAD_LIST, ";")
    For i = 0 To UBound(rows)

        f = Split(rows(i), "|")
        elemNo = CLng(f(0))

        Select Case f(3)
            Case "GZ": coordSys = "GLOBAL": loadDir = "Gravity": loadSign = -1
            Case "LZ": coordSys = "Local": loadDir = "2": loadSign = 1
            Case Else
                WriteDistributedLoadTable = "element " & elemNo & ", case " & f(1) & _
                    ": load direction """ & f(3) & """ has no SAP2000 mapping."
                Exit Function
        End Select

        If elemNo > UBound(ELEM_PIECES) Then
            WriteDistributedLoadTable = "a load names element " & elemNo & ", which does not exist."
            Exit Function
        End If
        pieces = ELEM_PIECES(elemNo)
        If pieces = 0 Then
            WriteDistributedLoadTable = "a load names element " & elemNo & ", which does not exist."
            Exit Function
        End If

        p1 = Val(f(4))
        p2 = Val(f(5))
        For k = 0 To pieces - 1
            fr = ELEM_FIRST_FRAME(elemNo) + k
            va = loadSign * (p1 + (p2 - p1) * k / pieces)
            vb = loadSign * (p1 + (p2 - p1) * (k + 1) / pieces)
            Call Emit("   Frame=" & fr & "   LoadPat=" & SapName(f(1)) & "   CoordSys=" & coordSys & _
                      "   Type=Force   Dir=" & loadDir & "   DistType=RelDist   RelDistA=0   RelDistB=1" & _
                      "   AbsDistA=0   AbsDistB=" & SapNum(FrameLength(fr)) & _
                      "   FOverLA=" & SapNum(Round(va, LOAD_ROUND_DP)) & _
                      "   FOverLB=" & SapNum(Round(vb, LOAD_ROUND_DP)))
            LOAD_COUNT = LOAD_COUNT + 1
        Next k

    Next i
    Call TableEnd

End Function

' ATA: the MIDAS db/BODF record FV = [kh, 0, 0] - the self-weight applied
' along global X. Seismic only, like the ATA pattern itself.
Private Sub WriteGravityLoadTable()

    Dim fr As Long

    If Not SEISMIC_ACTIVE Then Exit Sub

    Call TableStart("FRAME LOADS - GRAVITY")
    For fr = 1 To FR_COUNT
        Call Emit("   Frame=" & fr & "   LoadPat=ATA   CoordSys=GLOBAL   MultiplierX=" & _
                  SapNum(DIM_ATA_FACTOR) & "   MultiplierY=0   MultiplierZ=0")
    Next fr
    Call TableEnd

End Sub

' MIDAS beta -180 (left wall) is the same rotation as SAP's 180; angles are
' normalised into (-180, 180] and 0 is left out, as SAP's export does.
Private Sub WriteLocalAxesTable()

    Dim fr As Long
    Dim a As Double
    Dim wrote As Boolean

    For fr = 1 To FR_COUNT
        a = FR_ANGLE(fr)
        Do While a > 180: a = a - 360: Loop
        Do While a <= -180: a = a + 360: Loop
        If a <> 0 Then
            If Not wrote Then Call TableStart("FRAME LOCAL AXES ASSIGNMENTS 1 - TYPICAL")
            wrote = True
            Call Emit("   Frame=" & fr & "   Angle=" & SapNum(a))
        End If
    Next fr
    If wrote Then Call TableEnd

End Sub

' FRAME SECTION ASSIGNMENTS, 01 - GENERAL and 02 - CONCRETE COLUMN. Only
' the sections in SECT_LIST are defined - the same set the MIDAS builder
' writes, section 3 included when B6 > 0 although nothing uses it.
Private Function WriteSectionTables() As String

    Dim rows() As String
    Dim f() As String
    Dim i As Long, j As Long, fr As Long
    Dim nm As String, nm2 As String
    Dim d As Double, b As Double
    Dim a As Double, i33 As Double, i22 As Double, tors As Double
    Dim longSide As Double, shortSide As Double

    rows = Split(SECT_LIST, ";")

    ' Names must be unique - two sections sharing one would silently merge.
    For i = 0 To UBound(rows)
        nm = SectionNameOf(CLng(Split(rows(i), "|")(0)))
        For j = i + 1 To UBound(rows)
            nm2 = SectionNameOf(CLng(Split(rows(j), "|")(0)))
            If StrComp(nm, nm2, vbTextCompare) = 0 Then
                WriteSectionTables = "two sections are both named """ & nm & """ - rename one in " & _
                                     INPUT_SHEET_NAME & " column A."
                Exit Function
            End If
        Next j
    Next i

    For fr = 1 To FR_COUNT
        If Len(SectionNameOf(FR_SECT(fr))) = 0 Then
            WriteSectionTables = "frame " & fr & " uses section " & FR_SECT(fr) & _
                                 ", which was not defined."
            Exit Function
        End If
    Next fr

    Call TableStart("FRAME SECTION ASSIGNMENTS")
    For fr = 1 To FR_COUNT
        Call Emit("   Frame=" & fr & "   AutoSelect=N.A.   AnalSect=" & _
                  SapQuote(SectionNameOf(FR_SECT(fr))) & "   MatProp=Default")
    Next fr
    Call TableEnd

    ' Solid rectangle, depth d (t3) x width b (t2). The torsion constant is
    ' the usual series approximation, which reproduces the reference file's
    ' 0.0217931 for 0.45 x 1.0.
    Call TableStart("FRAME SECTION PROPERTIES 01 - GENERAL")
    b = SECTION_WIDTH
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
        d = Val(f(2))
        a = b * d
        i33 = b * d ^ 3 / 12
        i22 = d * b ^ 3 / 12
        longSide = IIf(b > d, b, d)
        shortSide = IIf(b > d, d, b)
        tors = longSide * shortSide ^ 3 * _
               (1 / 3 - 0.21 * (shortSide / longSide) * (1 - shortSide ^ 4 / (12 * longSide ^ 4)))
        Call Emit("   SectionName=" & SapQuote(SectionNameOf(CLng(f(0)))) & "   Material=" & MATERIAL_NAME & _
                  "   Shape=Rectangular   t3=" & SapNum(d) & "   t2=" & SapNum(b) & _
                  "   Area=" & SapNum(a) & "   TorsConst=" & SapNum(tors) & _
                  "   I33=" & SapNum(i33) & "   I22=" & SapNum(i22) & "   I23=0" & _
                  "   AS2=" & SapNum(a * 5 / 6) & "   AS3=" & SapNum(a * 5 / 6) & _
                  "   S33=" & SapNum(b * d ^ 2 / 6) & "   S22=" & SapNum(d * b ^ 2 / 6) & " _")
        Call Emit("        Z33=" & SapNum(b * d ^ 2 / 4) & "   Z22=" & SapNum(d * b ^ 2 / 4) & _
                  "   R33=" & SapNum(Sqr(i33 / a)) & "   R22=" & SapNum(Sqr(i22 / a)) & _
                  "   Color=" & SectionColorOf(CLng(f(0))) & "   FromFile=No   AMod=1   A2Mod=1" & _
                  "   A3Mod=1   JMod=1   I2Mod=1   I3Mod=1   MMod=1   WMod=1")
    Next i
    Call TableEnd

    ' Reinforcement layout as in the reference file (design, not check).
    Call TableStart("FRAME SECTION PROPERTIES 02 - CONCRETE COLUMN")
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
        Call Emit("   SectionName=" & SapQuote(SectionNameOf(CLng(f(0)))) & _
                  "   RebarMatL=A615Gr60   RebarMatC=A615Gr60   ReinfConfig=Rectangular" & _
                  "   LatReinf=Ties   Cover=0.04   NumBars3Dir=3   NumBars2Dir=3   BarSizeL=#9" & _
                  "   BarSizeC=#4   SpacingC=0.15   NumCBars2=3   NumCBars3=3   ReinfType=Design")
    Next i
    Call TableEnd

End Function

' The sheet's column-A name for a section number, "" when it is not defined.
' A blank name becomes SECT<n>.
Private Function SectionNameOf(ByVal sectNo As Long) As String

    Dim rows() As String
    Dim f() As String
    Dim i As Long

    rows = Split(SECT_LIST, ";")
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
        If CLng(f(0)) = sectNo Then
            SectionNameOf = Trim$(f(1))
            If Len(SectionNameOf) = 0 Then SectionNameOf = "SECT" & sectNo
            Exit Function
        End If
    Next i

End Function

' SECTION_COLOR_LIST's RGB as the integer SAP2000 stores (R + 256 G + 65536 B).
Private Function SectionColorOf(ByVal sectNo As Long) As String

    Dim entries() As String
    Dim f() As String
    Dim c() As String
    Dim i As Long

    SectionColorOf = "Gray8Dark"
    entries = Split(SECTION_COLOR_LIST, ";")
    For i = 0 To UBound(entries)
        f = Split(entries(i), "|")
        If CLng(f(0)) = sectNo Then
            c = Split(f(1), ",")
            SectionColorOf = CStr(CLng(c(0)) + 256& * CLng(c(1)) + 65536 * CLng(c(2)))
            Exit Function
        End If
    Next i

End Function

Private Sub WriteJointTable()

    Dim jt As Long

    Call TableStart("JOINT COORDINATES")
    For jt = 1 To JT_COUNT
        Call Emit("   Joint=" & jt & "   CoordSys=GLOBAL   CoordType=Cartesian   XorR=" & _
                  SapNum(JT_X(jt)) & "   Y=0   Z=" & SapNum(JT_Z(jt)) & _
                  "   SpecialJt=" & IIf(JT_SPECIAL(jt), "Yes", "No"))
    Next jt
    Call TableEnd

End Sub

' The MIDAS builder's db/NSPR rule on every joint at z = 0, sorted by X:
'   interior   L = (x[i+1] - x[i-1]) / 2
'   end        L = (x[2] - x[1]) / 2, plus half a wall thickness when the
'              foundation stops at the wall centreline (no extension)
'   Kz = ks * L,  Kx = Ky = Kz / 2
' The reference file's springs follow exactly this rule.
Private Function WriteSpringTable() As String

    Dim ids() As Long
    Dim xs() As Double
    Dim n As Long, jt As Long
    Dim i As Long, j As Long
    Dim tmpId As Long, tmpX As Double
    Dim ltrib As Double, kz As Double

    If SPRING_KZ_MODULUS <= 0 Then
        TEXT_WARNINGS = TEXT_WARNINGS & " no springs - " & INPUT_SHEET_NAME & "!B42 subgrade " & _
                        "modulus is " & SapNum(SPRING_KZ_MODULUS) & ", so the model has no supports."
        Exit Function
    End If

    ReDim ids(1 To JT_COUNT)
    ReDim xs(1 To JT_COUNT)
    For jt = 1 To JT_COUNT
        If Abs(JT_Z(jt)) < FOUNDATION_Z_TOL Then
            n = n + 1
            ids(n) = jt
            xs(n) = JT_X(jt)
        End If
    Next jt

    If n < 2 Then
        WriteSpringTable = "found " & n & " foundation joints at z = 0 (expected at least 2)."
        Exit Function
    End If

    For i = 1 To n - 1
        For j = i + 1 To n
            If xs(j) < xs(i) Then
                tmpX = xs(i): xs(i) = xs(j): xs(j) = tmpX
                tmpId = ids(i): ids(i) = ids(j): ids(j) = tmpId
            End If
        Next j
    Next i

    Call TableStart("JOINT SPRING ASSIGNMENTS 1 - UNCOUPLED")
    For i = 1 To n
        If i = 1 Then
            ltrib = (xs(2) - xs(1)) / 2
            If Not HAS_EXT Then ltrib = ltrib + DIM_WALL_T / 2
        ElseIf i = n Then
            ltrib = (xs(n) - xs(n - 1)) / 2
            If Not HAS_EXT Then ltrib = ltrib + DIM_WALL_T / 2
        Else
            ltrib = (xs(i + 1) - xs(i - 1)) / 2
        End If
        kz = SPRING_KZ_MODULUS * ltrib
        Call Emit("   Joint=" & ids(i) & "   CoordSys=Local   U1=" & SapNum(kz / 2) & _
                  "   U2=" & SapNum(kz / 2) & "   U3=" & SapNum(kz) & "   R1=0   R2=0   R3=0")
    Next i
    Call TableEnd

    SPRING_COUNT = n

End Function

' LOAD CASE DEFINITIONS (one linear static case per pattern, plus MODAL)
' and LOAD PATTERN DEFINITIONS. DL alone carries self-weight.
Private Sub WriteLoadCaseTables()

    Dim rows() As String
    Dim f() As String
    Dim i As Long
    Dim dt As String

    rows = Split(STLDCASE_LIST, ";")

    Call TableStart("LOAD CASE DEFINITIONS")
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
        If SEISMIC_ACTIVE Or Not IsSeismicCase(f(0)) Then
            dt = SapDesignType(f(0), f(1))
            Call Emit("   Case=" & SapName(f(0)) & "   Type=LinStatic   InitialCond=Zero" & _
                      "   DesTypeOpt=""Prog Det""   DesignType=" & dt & _
                      "   DesActOpt=""Prog Det""   DesignAct=" & SapDesignAct(dt) & _
                      "   AutoType=None   RunCase=Yes")
            If i = 0 Then
                Call Emit("   Case=MODAL   Type=LinModal   InitialCond=Zero   DesTypeOpt=""Prog Det""" & _
                          "   DesignType=OTHER   DesActOpt=""Prog Det""   DesignAct=Other" & _
                          "   AutoType=None   RunCase=Yes")
            End If
        End If
    Next i
    Call TableEnd

    Call TableStart("LOAD PATTERN DEFINITIONS")
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
        If SEISMIC_ACTIVE Or Not IsSeismicCase(f(0)) Then
            Call Emit("   LoadPat=" & SapName(f(0)) & "   DesignType=" & SapDesignType(f(0), f(1)) & _
                      "   SelfWtMult=" & IIf(f(0) = "DL", "1", "0"))
        End If
    Next i
    Call TableEnd

    Call TableStart("MASS SOURCE")
    Call Emit("   MassSource=MSSSRC1   Elements=Yes   Masses=Yes   Loads=No   IsDefault=Yes")
    Call TableEnd

End Sub

Private Function SapDesignType(ByVal caseName As String, ByVal midasType As String) As String

    Dim entries() As String
    Dim f() As String
    Dim i As Long

    If caseName = "LLacc" Then
        SapDesignType = "LIVE"
        Exit Function
    End If

    SapDesignType = "OTHER"
    entries = Split(SAP_DESIGN_TYPE_LIST, ";")
    For i = 0 To UBound(entries)
        f = Split(entries(i), "|")
        If f(0) = midasType Then
            SapDesignType = f(1)
            Exit Function
        End If
    Next i

End Function

Private Function SapDesignAct(ByVal designType As String) As String
    Select Case designType
        Case "DEAD": SapDesignAct = "Non-Composite"
        Case "OTHER": SapDesignAct = "Other"
        Case Else: SapDesignAct = """Short-Term Composite"""
    End Select
End Function

' C30/37 plus the A615Gr60 rebar the concrete-column sections reference.
' Unit weight is INPUT!B5, Ec MIDAS_INPUT!B43 (see ReadInputs); G = E/2(1+v).
Private Sub WriteMaterialTables()

    Call TableStart("MATERIAL PROPERTIES 01 - GENERAL")
    Call Emit("   Material=A615Gr60   Type=Rebar   SymType=Uniaxial   TempDepend=No   Color=Cyan")
    Call Emit("   Material=" & MATERIAL_NAME & "   Type=Concrete   SymType=Isotropic   TempDepend=No   Color=Green")
    Call TableEnd

    Call TableStart("MATERIAL PROPERTIES 02 - BASIC MECHANICAL PROPERTIES")
    Call Emit("   Material=A615Gr60   UnitWeight=76.9728639422648   UnitMass=7.84904737995992" & _
              "   E1=199947978.795958   A1=1.16999994421006E-05")
    Call Emit("   Material=" & MATERIAL_NAME & "   UnitWeight=" & SapNum(MATERIAL_DEN) & _
              "   UnitMass=" & SapNum(MATERIAL_DEN / GRAVITY_ACCEL) & _
              "   E1=" & SapNum(MATERIAL_ELAST) & _
              "   G12=" & SapNum(MATERIAL_ELAST / (2 * (1 + MATERIAL_POISN))) & _
              "   U12=" & SapNum(MATERIAL_POISN) & "   A1=" & SapNum(MATERIAL_THERMAL))
    Call TableEnd

    Call TableStart("MATERIAL PROPERTIES 03B - CONCRETE DATA")
    Call Emit("   Material=" & MATERIAL_NAME & "   Fc=" & SapNum(MATERIAL_FC) & _
              "   LtWtConc=No   SSCurveOpt=Mander   SSHysType=Takeda   SFc=0.00181818" & _
              "   SCap=0.005   FinalSlope=-0.1   FAngle=0   DAngle=0")
    Call TableEnd

    Call TableStart("MATERIAL PROPERTIES 03E - REBAR DATA")
    Call Emit("   Material=A615Gr60   Fy=413685.473370947   Fu=620528.21005642" & _
              "   EffFy=455054.020708041   EffFu=682581.031062062   SSCurveOpt=Simple" & _
              "   SSHysType=Kinematic   SHard=0.01   SCap=0.09   FinalSlope=-0.1   UseCTDef=No")
    Call TableEnd

End Sub

' Company fixed; project name, engineer and revision from the INPUT sheet
' (ReadProjectName), model name = the workbook's name, model description =
' this module's SCRIPT_ID and SCRIPT_VERSION.
Private Sub WriteProjectInfoTable()

    Call TableStart("PROJECT INFORMATION")
    Call Emit("   Item=""Company Name""   Data=" & SapQuote(PROJINFO_COMPANY))
    Call Emit("   Item=""Project Name""" & DataField(PROJECT_NAME))
    Call Emit("   Item=""Model Name""" & DataField(ThisWorkbook.Name))
    ' Which macro build wrote the file - tests/verify_sap2000_build.py reads it.
    Call Emit("   Item=""Model Description""" & DataField(SCRIPT_ID & " " & SCRIPT_VERSION))
    Call Emit("   Item=""Revision Number""" & DataField(PROJECT_REVISION))
    Call Emit("   Item=Engineer" & DataField(PROJECT_ENGINEER))
    Call TableEnd

End Sub

Private Function DataField(ByVal s As String) As String
    If Len(Trim$(s)) > 0 Then DataField = "   Data=" & SapQuote(Trim$(s))
End Function

' Only the two sizes the concrete-column sections name; SAP2000 adds its
' own standard list on import.
Private Sub WriteRebarSizeTable()

    Call TableStart("REBAR SIZES")
    Call Emit("   RebarID=#4   Area=0.000129032001922727   Diameter=0.0127")
    Call Emit("   RebarID=#9   Area=0.00064516   Diameter=0.0286512005329132")
    Call TableEnd

End Sub


' ===========================================================================
'  TEXT HELPERS
' ===========================================================================

Private Sub TableStart(ByVal tableName As String)
    Call Emit("TABLE:  """ & tableName & """")
End Sub

Private Sub TableEnd()
    Call Emit(" ")
End Sub

Private Sub Emit(ByVal s As String)
    OUT_COUNT = OUT_COUNT + 1
    If OUT_COUNT > UBound(OUT_LINES) Then ReDim Preserve OUT_LINES(1 To UBound(OUT_LINES) * 2)
    OUT_LINES(OUT_COUNT) = s
End Sub

' MIDAS case names -> SAP pattern names: EHS2_L -> EHS2L, as the reference
' file spells them. Combination names do not go through this.
Private Function SapName(ByVal midasName As String) As String
    SapName = Replace(midasName, "_", "")
End Function

' A value holding a space or "=" is quoted; a double quote inside it becomes
' a single one, since the text format has no escape for it.
Private Function SapQuote(ByVal s As String) As String
    s = Replace(s, """", "'")
    If Len(s) = 0 Or InStr(s, " ") > 0 Or InStr(s, "=") > 0 Then
        SapQuote = """" & s & """"
    Else
        SapQuote = s
    End If
End Function

' Locale-invariant number, rounded so float noise (0.30000000000000004)
' never reaches the file.
Private Function SapNum(ByVal v As Double) As String
    SapNum = JsonNum(Round(v, 9))
End Function

' Written in one go at the end, so a failure while assembling never leaves
' a half-written model behind. Print # uses the system code page, which is
' what SAP2000 reads.
Private Function WriteOutputFile(ByVal path As String) As String

    Dim fileNo As Integer
    Dim i As Long

    On Error GoTo Failed

    fileNo = FreeFile
    Open path For Output As #fileNo
    For i = 1 To OUT_COUNT
        Print #fileNo, OUT_LINES(i)
    Next i
    Close #fileNo
    Exit Function

Failed:
    WriteOutputFile = "could not write the file (" & Err.Description & ") - is it open " & _
                      "in another program, or is the folder read-only?"
    On Error Resume Next
    Close #fileNo

End Function
