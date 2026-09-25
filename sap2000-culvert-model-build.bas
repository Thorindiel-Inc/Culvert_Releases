Option Explicit

' ============================================================================
'  SAP2000 - Box Culvert Model Build  (through the SAP2000 API)
'
'  Builds the culvert in SAP2000 object by object, saves it as
'  <workbook folder>\SAP2000\<workbook name>.sdb, runs the analysis and
'  leaves SAP2000 open with the solved model. It starts its OWN SAP2000 -
'  one already open is never touched. The frame forces come back into
'  MIDAS_RESULTS's two tables, replacing the MIDAS results there (see
'  RESULTS below); no other worksheet is changed.
'
'  HOW IT REACHES SAP2000: Excel is 64-bit and the SAP2000 17 API is 32-bit
'  only, so the macro writes a PowerShell script (SAP2000\<name>_build.ps1,
'  kept for inspection) and runs it with Windows' 32-bit PowerShell; see
'  the SAP2000 API BRIDGE block. The script checks every API call and every
'  name SAP2000 hands back, and logs to SAP2000\<name>_build.log.
'
'  SAME MODEL AS THE MIDAS BUILDER. Inputs, geometry, loads and combinations
'  come from procedures copied BYTE FOR BYTE out of
'  midas-culvert-model-build.bas (ReadInputs, GenerateGeometry,
'  GenerateBeamLoads, GenerateLoadCombinations and their helpers, plus the
'  seismic gate). tests/verify_sap2000_build.py fails if a copy drifts, so a
'  change to the MIDAS model has to be carried over here too. Only the
'  SAP-specific part is this module's own: dividing and renumbering, the
'  load translation and the build script.
'
'  What the MIDAS builder does through its API happens here as well:
'    ope/DIVIDEELEM   elements 18, 10, 11, 4 split into INPUT!K15 equal
'                     frames; each piece carries its share of a trapezoid
'    db/NSPR          joint springs on every z = 0 joint, same tributary rule
'    db/BODF          DL = self-weight multiplier 1; ATA = gravity along X
'  Joints and frames carry the numbers MIDAS gives them, division pieces
'  included (see ExpandModel), so a frame in SAP2000 and in MIDAS_RESULTS
'  is the same member under the same number.
'
'  Owner's decisions (2026-09-24), against the hand-built reference model
'  old/SAP2000/3.00x3.00_Hd_3.0m.$2k:
'    - full MIDAS load set (20 cases less the seismic pair when the gate is
'      off) and all 33/35 combinations, not the reference's subset
'    - rigid-zone and haunch members get their own sections (MIDAS_INPUT
'      rows 10-16) instead of the plain slab/wall/foundation one
'    - XZ plane frame (UX UZ RY active), so springs at z = 0 alone make it
'      stable; the reference had all six DOF
'    - load-case names lose their "_" (EHS2_L -> EHS2L, as the reference
'      names them); combination names are kept exactly as in MIDAS
'    - built through the API (2026-09-24), replacing the .$2k text file
'  Ec is MIDAS_INPUT!B43, falling back to 33 GPa (EN C30/37, the reference's
'  value) - NOT to the MIDAS builder's 26.291 GPa fallback.
'
'  No restraints are written: the foundation springs are the supports, as
'  in the reference. Design-only data of the old text file (column rebar,
'  A615Gr60) is not built - the model is for analysis.
' ============================================================================


' ---------------------------------------------------------------------------
'  CONFIG
' ---------------------------------------------------------------------------

' Stamped into the report title. Bump with every change to this file.
Private Const SCRIPT_VERSION As String = "2026-09-25a"

' One line, no "_" continuation, no "|" - read by the updater's manifest.
Private Const SCRIPT_CHANGELOG As String = "Calculation always goes back to Automatic at the end (it used to restore the starting mode, so one interrupted run left Excel on Manual for good)"

' Identifies this module to the updater whatever it was named in Excel.
Private Const SCRIPT_ID As String = "sap2000-culvert-model-build"

' SAP2000 runs hidden while the model is built and solved, then its window
' is shown and left open. True keeps it visible throughout (slower, and a
' click in SAP2000 mid-build can interfere).
Private Const SAP_VISIBLE As Boolean = False

' Longest the macro waits for SAP2000, in seconds, before it gives up and
' closes the SAP2000 it started (a hidden SAP2000 waiting on a dialog would
' otherwise block Excel for good). One culvert takes a few seconds.
Private Const SAP_TIMEOUT_SEC As Long = 600

' Folder next to the workbook that receives the .sdb, SAP2000's analysis
' files, the build script and its log.
Private Const SAP_FOLDER_NAME As String = "SAP2000"

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

' RESULTS - SAP2000's forces go into MIDAS_RESULTS's two data tables,
' replacing what is there, row for row as the MIDAS builder writes them
' (headers, summary block and helper formulas stay as they are):
'   table 1  B:J from row 3    every SLS-*, ULS-* and EQ-1 combination
'   table 2  S:AA from row 36  ENV_SER/ENV_ALL max, then ENV_SER/ENV_ALL min
' Combination outer, element ascending, then I[node], 2/4, J[node]; values
' to 2 dp (MIDAS STYLES PLACE 2). MIDAS columns from SAP (measured against a
' MIDAS build, same signs): Axial = P, Shear-y = V3, Shear-z = V2,
' Torsion = T, Moment-y = M3, Moment-z = M2.
Private Const RESULT_SHEET_NAME As String = "MIDAS_RESULTS"
Private Const FORCE_ROUND_DP As Long = 2

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

' ---------------------------------------------------------------------------
'  LIVE-LOAD GATE - identical to the MIDAS builder (see its header):
'  INPUT!B26 ("HAREKETLI YUK"), trimmed case-insensitive: "YOK" -> OFF,
'  blank or anything else -> ON, an error value stops the export. Off
'  drops LL/LLin/LLacc/LSS2_L/LSA2_L cases and loads, their terms out of
'  every combination, and ACC-1 entirely. LL1/LSS1_L are unaffected.
' ---------------------------------------------------------------------------
Private LIVE_LOAD_ACTIVE As Boolean
Private LIVE_LOAD_GATE_SOURCE As String

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

' The MIDAS static load cases, in the MIDAS builder's order. EQ/ATA are
' the seismic pair (IsSeismicCase); LL/LLin/LLacc/LSS2_L/LSA2_L are the
' live-load ones (IsLiveLoadCase) - both dropped by EmitLoadPatterns when
' their gate is off.
Private Const STLDCASE_LIST As String = _
    "DL|D;EV1|EV;EV2|EV;EVin|EV;EHS1|EH;EHA1|EH;EHS2_L|EH;EHA2_L|EH;" & _
    "EHS2_R|EH;EHA2_R|EH;LL1|L;LL|L;LLin|L;LLacc|E;LSS1_L|LS;LSS2_L|LS;" & _
    "LSA2_L|LS;WA|FP;EQ|E;ATA|E"

' ---------------------------------------------------------------------------
'  SAP MODEL - the divided model ExpandModel builds. Joints are indexed
'  1..JT_COUNT and frames 1..FR_COUNT; JT_NAME/FR_NAME hold the MIDAS number
'  each is created under in SAP2000. ELEM_FIRST_FRAME/ELEM_PIECES map a MIDAS
'  element number to its frame indices (indexed by element number).
' ---------------------------------------------------------------------------
Private JT_COUNT As Long
Private JT_NAME() As Long
Private FR_NAME() As Long
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
    Dim ok As Boolean, built As Boolean
    Dim folder As String, stem As String
    Dim res As String
    Dim stage As String
    Dim liveLoadErr As String

    report = "SAP2000 culvert model  [" & SCRIPT_VERSION & "]" & vbCrLf & _
             String(40, "-") & vbCrLf

    On Error GoTo Crashed

    stage = "seismic gate"
    Call InitSeismicGate
    If Not SEISMIC_GATE_KNOWN Then
        MsgBox report & "STOPPED - nothing was built." & vbCrLf & vbCrLf & _
               "The seismic gate " & SEISMIC_GATE_SOURCE & "." & vbCrLf & vbCrLf & _
               "Fix those cells, or set SEISMIC_GATE_OVERRIDE to 0 (off) or 1 (on) " & _
               "at the top of this module.", vbExclamation
        Exit Sub
    End If
    report = report & "NOTE - seismic gate is " & IIf(SEISMIC_ACTIVE, "ON", "OFF") & _
             " (" & SEISMIC_GATE_SOURCE & ")." & vbCrLf & String(40, "-") & vbCrLf

    stage = "live-load gate"
    liveLoadErr = InitLiveLoadGate()
    If Len(liveLoadErr) > 0 Then
        MsgBox report & "STOPPED - nothing was built." & vbCrLf & vbCrLf & _
               liveLoadErr, vbExclamation
        Exit Sub
    End If
    report = report & "NOTE - live load is " & IIf(LIVE_LOAD_ACTIVE, "ON", "OFF") & _
             " (" & LIVE_LOAD_GATE_SOURCE & ")." & IIf(LIVE_LOAD_ACTIVE, "", _
             " No LL/LLin/LLacc/LSS2_L/LSA2_L cases or loads, no ACC-1.") & _
             vbCrLf & String(40, "-") & vbCrLf

    stage = "SAP2000 folder"
    folder = SapFolder()
    ok = StepResult(report, "SAP2000 folder", IIf(Len(folder) = 0, _
             "the workbook has never been saved, so there is no folder to put the model " & _
             "in. Save it first.", ""))
    If ok Then stem = SapBasePath(folder)

    stage = "read inputs"
    If ok Then
        res = ReadInputs()
        If Len(res) = 0 And Len(INPUT_FALLBACKS) > 0 Then
            res = "WARN: used built-in defaults -" & INPUT_FALLBACKS
        End If
        ok = StepResult(report, "Read inputs", res)
    End If

    ' Without springs the model has no supports and SAP2000 stops on an
    ' instability dialog - refuse before anything is started.
    If ok And SPRING_KZ_MODULUS <= 0 Then
        ok = StepResult(report, "Springs", INPUT_SHEET_NAME & "!B42 (subgrade modulus) is " & _
                        SapNum(SPRING_KZ_MODULUS) & " - without springs the model has no supports " & _
                        "and SAP2000 cannot solve it.")
    End If

    stage = "generate model"
    If ok Then
        Call ReadProjectName
        Call GenerateGeometry
        Call ComputeOrdinates
        Call GenerateBeamLoads
        Call GenerateLoadCombinations
        ok = StepResult(report, "Divide + number", ExpandModel())
        If SEISMIC_ACTIVE And DIM_ATA_FACTOR = 0 Then
            report = report & "NOTE - kh (" & INPUT_SHEET_NAME & "!B15) is 0: ATA load pattern kept, " & _
                     "no ATA self-weight." & vbCrLf
        End If
    End If

    stage = "build script"
    If ok Then
        res = WriteBuildScript(stem & ".sdb", stem & "_build.log", stem & "_forces.txt")
        If Len(res) = 0 Then res = WriteScriptFile(stem & "_build.ps1")
        ok = StepResult(report, "Build script", res)
    End If

    ' Last run's forces must never be read as this run's.
    stage = "SAP2000"
    If ok Then
        With CreateObject("Scripting.FileSystemObject")
            If .FileExists(stem & "_forces.txt") Then .DeleteFile stem & "_forces.txt", True
        End With
        ok = StepResult(report, "SAP2000: build, save, analyse, frame forces", _
                        RunSapScript(stem & "_build.ps1", stem & "_build.log"))
    End If
    built = ok

    stage = "results"
    If ok Then ok = StepResult(report, "Frame forces to " & RESULT_SHEET_NAME, _
                               WriteSapResults(stem & "_forces.txt"))

    If built Then
        report = report & String(40, "-") & vbCrLf & _
                 JT_COUNT & " joints, " & FR_COUNT & " frames (divided: " & DIVIDED_ELEMS & _
                 "), " & SPRING_COUNT & " springs, " & LOAD_COUNT & " frame loads, " & _
                 (UBound(Split(LOADCOMB_LIST, ";")) + 1) & " combinations." & vbCrLf & _
                 "SAP2000 is open with the analysed model:" & vbCrLf & stem & ".sdb" & vbCrLf
    End If

    report = VerdictFirst(report, ok)
    MsgBox FitReport(report, SCRIPT_ID), IIf(ok, vbInformation, vbExclamation)
    Exit Sub

Crashed:
    Application.StatusBar = False
    report = report & "FAIL - " & stage & ": VBA error " & Err.Number & " - " & _
             Err.Description & vbCrLf
    report = VerdictFirst(report, False)
    MsgBox FitReport(report, SCRIPT_ID), vbExclamation

End Sub


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

' Reads INPUT!B26 and sets LIVE_LOAD_ACTIVE/LIVE_LOAD_GATE_SOURCE. Returns
' "" when the cell could be evaluated, otherwise a message naming it - an
' error value there stops the build, same pattern as RequirePositive.
Private Function InitLiveLoadGate() As String

    Dim ws As Worksheet
    Dim v As Variant
    Dim s As String

    InitLiveLoadGate = ""
    ' Reset every run - module-level values survive between runs.
    LIVE_LOAD_ACTIVE = True
    LIVE_LOAD_GATE_SOURCE = ""

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PROJECT_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        LIVE_LOAD_ACTIVE = True
        LIVE_LOAD_GATE_SOURCE = "sheet """ & PROJECT_SHEET_NAME & """ not found, defaulted to ON"
        Exit Function
    End If

    v = ws.Range("B26").Value

    If IsError(v) Then
        InitLiveLoadGate = PROJECT_SHEET_NAME & "!B26 (HAREKETLI YUK) is an error value. " & _
                           "Fix the formula, or set it to a valid option or blank."
        Exit Function
    End If

    s = Trim$(CStr(v))

    If StrComp(s, "YOK", vbTextCompare) = 0 Then
        LIVE_LOAD_ACTIVE = False
        LIVE_LOAD_GATE_SOURCE = PROJECT_SHEET_NAME & "!B26 = ""YOK"""
    Else
        LIVE_LOAD_ACTIVE = True
        LIVE_LOAD_GATE_SOURCE = PROJECT_SHEET_NAME & "!B26 = " & _
                                IIf(Len(s) = 0, "(blank)", """" & s & """")
    End If

End Function

' The live-load cases dropped entirely when LIVE_LOAD_ACTIVE is False. LL1
' and LSS1_L are deliberately NOT here - the owner keeps them on regardless
' (2026-09-24).
Private Function IsLiveLoadCase(ByVal nm As String) As Boolean
    IsLiveLoadCase = (nm = "LL" Or nm = "LLin" Or nm = "LLacc" Or _
                      nm = "LSS2_L" Or nm = "LSA2_L")
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
    ReadInputs = InitLiveLoadGate()
    If Len(ReadInputs) > 0 Then Exit Function

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
    Dim nWallTopL As Long, nWallTopR As Long, nSlabL As Long, nSlabR As Long

    ' Where the plain wall and slab members end when a haunch is absent.
    nWallTopL = IIf(DIM_SLAB_HAUNCH > 0, 11, 13)
    nWallTopR = IIf(DIM_SLAB_HAUNCH > 0, 12, 14)
    nSlabL = IIf(DIM_WALL_HAUNCH > 0, 17, 16)
    nSlabR = IIf(DIM_WALL_HAUNCH > 0, 18, 19)

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
        If DIM_SLAB_HAUNCH > 0 Then
            n = n & ";" & NodeRow(11, dx1, DIM_FOUND_T / 2 + DIM_WALL_H - DIM_SLAB_HAUNCH)
            n = n & ";" & NodeRow(12, dx2, DIM_FOUND_T / 2 + DIM_WALL_H - DIM_SLAB_HAUNCH)
        End If
        n = n & ";" & NodeRow(13, dx1, DIM_FOUND_T / 2 + DIM_WALL_H)
        n = n & ";" & NodeRow(14, dx2, DIM_FOUND_T / 2 + DIM_WALL_H)
        n = n & ";" & NodeRow(15, dx1, dz2)
        n = n & ";" & NodeRow(16, dx1 + DIM_WALL_T / 2, dz2)
        If DIM_WALL_HAUNCH > 0 Then
            n = n & ";" & NodeRow(17, dx1 + DIM_WALL_T / 2 + DIM_WALL_HAUNCH, dz2)
            n = n & ";" & NodeRow(18, dx1 + DIM_WALL_T / 2 + DIM_SPAN - DIM_WALL_HAUNCH, dz2)
        End If
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
        e = e & ";" & ElemRow(10, 2, nWallTopL, 9, -180)
        e = e & ";" & ElemRow(11, 2, nWallTopR, 10, 0)
        If DIM_SLAB_HAUNCH > 0 Then
            e = e & ";" & ElemRow(12, sHaunchWall, 13, 11, -180)
            e = e & ";" & ElemRow(13, sHaunchWall, 14, 12, 0)
        End If
        e = e & ";" & ElemRow(14, 8, 15, 13, -180)
        e = e & ";" & ElemRow(15, 8, 20, 14, 0)
        e = e & ";" & ElemRow(16, 7, 15, 16, 0)
        If DIM_WALL_HAUNCH > 0 Then e = e & ";" & ElemRow(17, sHaunchSlab, 16, 17, 0)
        e = e & ";" & ElemRow(18, 1, nSlabL, nSlabR, 0)
        If DIM_WALL_HAUNCH > 0 Then e = e & ";" & ElemRow(19, sHaunchSlab, 18, 19, 0)
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
        If DIM_SLAB_HAUNCH > 0 Then
            n = n & ";" & NodeRow(11, dx1, DIM_FOUND_T / 2 + DIM_WALL_H - DIM_SLAB_HAUNCH)
            n = n & ";" & NodeRow(12, dx2, DIM_FOUND_T / 2 + DIM_WALL_H - DIM_SLAB_HAUNCH)
        End If
        n = n & ";" & NodeRow(13, dx1, DIM_FOUND_T / 2 + DIM_WALL_H)
        n = n & ";" & NodeRow(14, dx2, DIM_FOUND_T / 2 + DIM_WALL_H)
        n = n & ";" & NodeRow(15, dx1, dz2)
        n = n & ";" & NodeRow(16, dx1 + DIM_WALL_T / 2, dz2)
        If DIM_WALL_HAUNCH > 0 Then
            n = n & ";" & NodeRow(17, dx1 + DIM_WALL_T / 2 + DIM_WALL_HAUNCH, dz2)
            n = n & ";" & NodeRow(18, dx1 + DIM_WALL_T / 2 + DIM_SPAN - DIM_WALL_HAUNCH, dz2)
        End If
        n = n & ";" & NodeRow(19, dx1 + DIM_WALL_T / 2 + DIM_SPAN, dz2)
        n = n & ";" & NodeRow(20, dx1 + DIM_WALL_T + DIM_SPAN, dz2)

        e = ElemRow(3, sRigidFound, 3, 4, 0)
        e = e & ";" & ElemRow(4, 4, 4, 5, 0)
        e = e & ";" & ElemRow(5, sRigidFound, 5, 6, 0)
        e = e & ";" & ElemRow(8, 8, 9, 3, -180)
        e = e & ";" & ElemRow(9, 8, 10, 6, 0)
        e = e & ";" & ElemRow(10, 2, nWallTopL, 9, -180)
        e = e & ";" & ElemRow(11, 2, nWallTopR, 10, 0)
        If DIM_SLAB_HAUNCH > 0 Then
            e = e & ";" & ElemRow(12, sHaunchWall, 13, 11, -180)
            e = e & ";" & ElemRow(13, sHaunchWall, 14, 12, 0)
        End If
        e = e & ";" & ElemRow(14, 8, 15, 13, -180)
        e = e & ";" & ElemRow(15, 8, 20, 14, 0)
        e = e & ";" & ElemRow(16, 7, 15, 16, 0)
        If DIM_WALL_HAUNCH > 0 Then e = e & ";" & ElemRow(17, sHaunchSlab, 16, 17, 0)
        e = e & ";" & ElemRow(18, 1, nSlabL, nSlabR, 0)
        If DIM_WALL_HAUNCH > 0 Then e = e & ";" & ElemRow(19, sHaunchSlab, 18, 19, 0)
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
    Dim i As Long

    If pTop = pBot Then
        For i = 1 To 8
            o(i) = pBot
        Next i
        Exit Sub
    End If

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

    ' A haunch member left out by GenerateGeometry (no haunch) carries
    ' nothing; its neighbours already span the whole height/length.
    If Not ElementExists(elemNo) Then Exit Sub

    ' Live load off: LL/LLin/LLacc/LSS2_L/LSA2_L carry nothing. LL1 and
    ' LSS1_L are not live-load cases here and are unaffected.
    If Not LIVE_LOAD_ACTIVE And IsLiveLoadCase(lcname) Then Exit Sub

    If Len(BEAMLOAD_LIST) > 0 Then BEAMLOAD_LIST = BEAMLOAD_LIST & ";"

    BEAMLOAD_LIST = BEAMLOAD_LIST & elemNo & "|" & lcname & "|" & cmd & "|" & loadDir & _
                    "|" & JsonNum(Round(p1, LOAD_ROUND_DP)) & _
                    "|" & JsonNum(Round(p2, LOAD_ROUND_DP))

End Sub

Private Function ElementExists(ByVal elemNo As Long) As Boolean
    ElementExists = (InStr(1, ";" & ELEM_LIST, ";" & elemNo & "|", vbBinaryCompare) > 0)
End Function

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
    ' gated on LIVE_LOAD_ACTIVE, not SEISMIC_ACTIVE - LLacc is its only
    ' load, so with live load off it would be DL,EV2,EHA2_L,EHA2_R alone,
    ' the same as SLS-6, and is dropped entirely instead (owner's
    ' decision, 2026-09-24). EQ-1 goes only when the seismic gate is on.
    '
    ' EQ-1 deliberately does NOT reference ATA (by request, 2026-09-23) -
    ' only the EQ lateral earth pressure case. ATA is still built as a
    ' static load case (db/STLD) and still carries its self-weight
    ' inertia record (db/BODF) whenever SEISMIC_ACTIVE, per the
    ' SEISMIC GATE block above - it is simply not pulled into any
    ' combination any more, EQ-1 or otherwise.
    If LIVE_LOAD_ACTIVE Then
        Call AddCombo("ACC-1", 0, "ST", "DL:1,EV2:1,EHA2_L:1,EHA2_R:1,LLacc:1", 1)
    End If
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
        ' Live load off: drop this pair from every "ST" combination that
        ' references a live-load case, so SLS-*/ULS-*/ACC-1 keep their
        ' names and every other term, just without that one - centralised
        ' here so no individual combination needs its own gate.
        If Not (anal = "ST" And Not LIVE_LOAD_ACTIVE And IsLiveLoadCase(kv(0))) Then
            factors = factors & IIf(Len(factors) > 0, ",", "") & _
                      anal & ":" & kv(0) & ":" & kv(1)
        End If
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
    Dim tg() As String
    Dim nextElem As Long, nextNode As Long
    Dim pieceElem() As Long, pieceNode() As Long

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

    ' MIDAS numbering of the division (read off a real MIDAS build, K15 = 6:
    ' 18 -> 18, 21-25; 10 -> 10, 26-30; 11 -> 11, 31-35; 4 -> 4, 36-40).
    ' ope/DIVIDEELEM runs on the targets in DIVIDE_TARGETS order; each keeps
    ' its own number for the piece at its I end, and its other pieces and new
    ' nodes take the next free numbers, running from the I end to the J end.
    ' pieceElem/pieceNode hold each target's first new number.
    ReDim pieceElem(0 To maxElem)
    ReDim pieceNode(0 To maxElem)
    nextElem = maxElem + 1
    nextNode = maxNode + 1
    If DIVIDE_AMOUNT > 1 Then
        tg = Split(DIVIDE_TARGETS, ",")
        For i = 0 To UBound(tg)
            elemNo = CLng(tg(i))
            If ElementExists(elemNo) Then
                pieceElem(elemNo) = nextElem
                pieceNode(elemNo) = nextNode
                nextElem = nextElem + DIVIDE_AMOUNT - 1
                nextNode = nextNode + DIVIDE_AMOUNT - 1
            End If
        Next i
    End If

    cap = (UBound(nodes) + 1) + (UBound(elems) + 1) * IIf(DIVIDE_AMOUNT > 1, DIVIDE_AMOUNT, 1)
    ReDim JT_NAME(1 To cap)
    ReDim JT_X(1 To cap)
    ReDim JT_Z(1 To cap)
    ReDim JT_SPECIAL(1 To cap)
    ReDim FR_NAME(1 To cap)
    ReDim FR_I(1 To cap)
    ReDim FR_J(1 To cap)
    ReDim FR_SECT(1 To cap)
    ReDim FR_ANGLE(1 To cap)

    JT_COUNT = 0
    For i = 0 To UBound(nodes)
        f = Split(nodes(i), "|")
        JT_COUNT = JT_COUNT + 1
        JT_NAME(JT_COUNT) = CLng(f(0))
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
                JT_NAME(JT_COUNT) = pieceNode(elemNo) + k - 1
                JT_X(JT_COUNT) = JT_X(ni) + (JT_X(nj) - JT_X(ni)) * t
                JT_Z(JT_COUNT) = JT_Z(ni) + (JT_Z(nj) - JT_Z(ni)) * t
                JT_SPECIAL(JT_COUNT) = False
                jNext = JT_COUNT
            End If
            FR_COUNT = FR_COUNT + 1
            FR_NAME(FR_COUNT) = IIf(k = 1, elemNo, pieceElem(elemNo) + k - 2)
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



' ===========================================================================
'  BUILD SCRIPT
'  One checked SAP2000 API call per object, in dependency order: material,
'  sections, joints, frames, springs, load patterns, loads, combinations,
'  DOF, project information - then save, analyse, show. The values are the
'  ones the .$2k writer used to write (and that the owner's reference model
'  matched), so the model is the same; only the way it reaches SAP changed.
' ===========================================================================

Private Function WriteBuildScript(ByVal sdbPath As String, ByVal resultPath As String, _
                                  ByVal forcesPath As String) As String

    Dim res As String

    OUT_COUNT = 0
    ReDim OUT_LINES(1 To 1024)
    LOAD_COUNT = 0
    SPRING_COUNT = 0

    res = ValidateSections()
    If Len(res) > 0 Then WriteBuildScript = res: Exit Function

    Call EmitSapPreamble(resultPath)

    ' Helpers for this model. Every add compares the name SAP2000 hands back
    ' with the one asked for - a clash is renamed silently otherwise.
    Call Emit("  function SapPt($nm, $x, $z) { $n = ''; SapChk ""Joint $nm"" ($m.PointObj.AddCartesian([double]$x, 0.0, [double]$z, [ref]$n, $nm, 'Global', $true, 0)); SapNamed ""Joint $nm"" $nm $n }")
    Call Emit("  function SapFr($nm, $i, $j, $sec, $ang) { $n = ''; SapChk ""Frame $nm"" ($m.FrameObj.AddByPoint($i, $j, [ref]$n, $sec, $nm)); SapNamed ""Frame $nm"" $nm $n; if ($ang -ne 0) { SapChk ""Frame $nm local axes"" ($m.FrameObj.SetLocalAxes($nm, [double]$ang, 0)) }; SapChk ""Frame $nm output stations"" ($m.FrameObj.SetOutputStations($nm, 2, 0.0, 3, $false, $false, 0)) }")
    Call Emit("  function SapSpr($nm, $kh, $kv) { $k = [double[]]($kh, $kh, $kv, 0, 0, 0); SapChk ""Spring on joint $nm"" ($m.PointObj.SetSpring($nm, [ref]$k, 0, $true, $true)) }")
    Call Emit("  function SapLd($fr, $pat, $dir, $a, $b) { $cs = 'Local'; if ($dir -eq 10) { $cs = 'Global' }; SapChk ""Load $pat on frame $fr"" ($m.FrameObj.SetLoadDistributed($fr, $pat, 1, $dir, 0.0, 1.0, [double]$a, [double]$b, $cs, $true, $false, 0)) }")
    Call Emit("  function SapGr($fr, $pat, $gx) { SapChk ""Self-weight $pat on frame $fr"" ($m.FrameObj.SetLoadGravity($fr, $pat, [double]$gx, 0.0, 0.0, $false, 'Global', 0)) }")
    Call Emit("  function SapPat($nm, $type, $sw) { SapChk ""Load pattern $nm"" ($m.LoadPatterns.Add($nm, $type, [double]$sw, $true)) }")
    Call Emit("  function SapCmb($nm, $type) { SapChk ""Combination $nm"" ($m.RespCombo.Add($nm, $type)) }")
    Call Emit("  function SapCmbAdd($nm, $ct, $case, $sf) { $t = $ct; SapChk ""Combination $nm + $case"" ($m.RespCombo.SetCaseList($nm, [ref]$t, $case, [double]$sf)) }")

    Call EmitMaterial
    Call EmitSections
    Call EmitJointsAndFrames

    res = EmitSprings()
    If Len(res) > 0 Then WriteBuildScript = res: Exit Function

    Call EmitLoadPatterns

    res = EmitDistributedLoads()
    If Len(res) > 0 Then WriteBuildScript = res: Exit Function
    Call EmitSeismicSelfWeight

    res = EmitCombinations()
    If Len(res) > 0 Then WriteBuildScript = res: Exit Function

    ' XZ plane frame (MIDAS STYP X-Z): springs at z = 0 then leave no
    ' mechanism - with all six DOF the frame could spin about global X.
    Call Emit("  $dof = [bool[]]($true, $false, $true, $false, $true, $false)")
    Call Emit("  SapChk 'Active DOF' ($m.Analyze.SetActiveDOF([ref]$dof))")

    Call EmitProjectInfo
    Call Emit("  SapCount 'joints' $m.PointObj " & JT_COUNT)
    Call Emit("  SapCount 'frames' $m.FrameObj " & FR_COUNT)
    Call Emit("  SapLog 'OK|Model built'")
    Call EmitSapFinish(sdbPath, ForcePullScript(forcesPath))

End Function

' Script lines that read the solved model's frame forces - the table 1 and
' table 2 combinations only, envelopes as Max/Min rows - for every frame
' (group ALL) and write them to forcesPath, one result per line:
'   frame <TAB> combination <TAB> step (""/Max/Min) <TAB> station <TAB>
'   P <TAB> V2 <TAB> V3 <TAB> T <TAB> M2 <TAB> M3
' in full precision with invariant number formatting.
Private Function ForcePullScript(ByVal forcesPath As String) As String

    Dim s As String
    Dim c As Variant
    Dim names As String

    For Each c In ResultCombos(1)
        names = names & IIf(Len(names) > 0, ", ", "") & PsQ(CStr(c))
    Next c
    For Each c In ResultCombos(2)
        names = names & ", " & PsQ(CStr(c))
    Next c

    s = "  $rs = $m.Results.Setup" & vbCrLf
    s = s & "  SapChk 'Results: deselect all' ($rs.DeselectAllCasesAndCombosForOutput())" & vbCrLf
    s = s & "  SapChk 'Results: envelopes as Max/Min' ($rs.SetOptionMultiValuedCombo(1))" & vbCrLf
    s = s & "  foreach ($cb in @(" & names & ")) { SapChk ""Results: select $cb"" ($rs.SetComboSelectedForOutput($cb, $true)) }" & vbCrLf
    s = s & "  $nr = 0; $ob = [string[]]@(); $os = [double[]]@(); $el = [string[]]@(); $es = [double[]]@()" & vbCrLf
    s = s & "  $lc = [string[]]@(); $sty = [string[]]@(); $snum = [double[]]@()" & vbCrLf
    s = s & "  $fP = [double[]]@(); $fV2 = [double[]]@(); $fV3 = [double[]]@(); $fT = [double[]]@(); $fM2 = [double[]]@(); $fM3 = [double[]]@()" & vbCrLf
    s = s & "  SapChk 'Frame forces' ($m.Results.FrameForce('ALL', 2, [ref]$nr, [ref]$ob, [ref]$os, [ref]$el, [ref]$es, [ref]$lc, [ref]$sty, [ref]$snum, [ref]$fP, [ref]$fV2, [ref]$fV3, [ref]$fT, [ref]$fM2, [ref]$fM3))" & vbCrLf
    s = s & "  if ($nr -lt 1) { throw 'SAP2000 returned no frame forces' }" & vbCrLf
    s = s & "  $inv = [Globalization.CultureInfo]::InvariantCulture" & vbCrLf
    s = s & "  $rows = New-Object 'System.Collections.Generic.List[string]'" & vbCrLf
    s = s & "  for ($k = 0; $k -lt $nr; $k++) { $rows.Add(($ob[$k], $lc[$k], $sty[$k], $os[$k].ToString('R', $inv), $fP[$k].ToString('R', $inv), $fV2[$k].ToString('R', $inv), $fV3[$k].ToString('R', $inv), $fT[$k].ToString('R', $inv), $fM2[$k].ToString('R', $inv), $fM3[$k].ToString('R', $inv)) -join ""`t"") }" & vbCrLf
    s = s & "  [IO.File]::WriteAllLines(" & PsQ(forcesPath) & ", $rows, (New-Object Text.UTF8Encoding($true)))" & vbCrLf
    s = s & "  SapLog ""OK|Frame forces: $nr results"""

    ForcePullScript = s

End Function

' The combinations of results table 1 (every SLS-*, ULS-* and EQ-1 the
' model has, in LOADCOMB_LIST order - the MIDAS builder's rule, ACC-1 left
' out) or table 2 (the two envelopes).
Private Function ResultCombos(ByVal tableNo As Long) As Collection

    Dim rows() As String
    Dim nm As String
    Dim i As Long

    Set ResultCombos = New Collection
    If tableNo = 2 Then
        ResultCombos.Add "ENV_SER"
        ResultCombos.Add "ENV_ALL"
        Exit Function
    End If

    rows = Split(LOADCOMB_LIST, ";")
    For i = 0 To UBound(rows)
        nm = Split(rows(i), "|")(0)
        If Left$(nm, 4) = "SLS-" Or Left$(nm, 4) = "ULS-" Or nm = "EQ-1" Then ResultCombos.Add nm
    Next i

End Function


' ===========================================================================
'  RESULTS - into MIDAS_RESULTS (see CONFIG)
' ===========================================================================

' Reads the forces file, checks every frame has exactly its three stations
' for every combination, and replaces MIDAS_RESULTS's two data tables with
' SAP2000's values.
Private Function WriteSapResults(ByVal forcesPath As String) As String

    Dim raw As String
    Dim lines() As String
    Dim f() As String
    Dim idx As Object
    Dim sta() As Double, vals() As Double, cnt() As Long
    Dim n As Long, i As Long, k As Long, key As String
    Dim order() As Long
    Dim t1 As Collection, t2 As Collection
    Dim data1() As Variant, data2() As Variant
    Dim c As Variant, stepName As Variant
    Dim r As Long, res As String

    raw = ReadUtf8File(forcesPath)
    If Len(raw) = 0 Then
        WriteSapResults = "SAP2000 wrote no frame forces (" & forcesPath & ")."
        Exit Function
    End If

    ' One entry per frame|combination|step, its three stations as they come.
    lines = Split(Replace(raw, vbCr, ""), vbLf)
    Set idx = CreateObject("Scripting.Dictionary")
    ReDim sta(1 To UBound(lines) + 1, 1 To 3)
    ReDim vals(1 To UBound(lines) + 1, 1 To 3, 1 To 6)
    ReDim cnt(1 To UBound(lines) + 1)
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            f = Split(lines(i), vbTab)
            If UBound(f) <> 9 Then
                WriteSapResults = "line " & (i + 1) & " of " & forcesPath & " has " & (UBound(f) + 1) & _
                                  " fields, expected 10."
                Exit Function
            End If
            key = f(0) & "|" & f(1) & "|" & f(2)
            If Not idx.Exists(key) Then
                n = n + 1
                idx.Add key, n
            End If
            r = idx(key)
            cnt(r) = cnt(r) + 1
            If cnt(r) > 3 Then
                WriteSapResults = "frame " & f(0) & ", " & f(1) & " " & f(2) & ": more than 3 output stations."
                Exit Function
            End If
            sta(r, cnt(r)) = Val(f(3))
            For k = 1 To 6
                vals(r, cnt(r), k) = Val(f(3 + k))
            Next k
        End If
    Next i

    ' Frame indices by MIDAS number, ascending - the MIDAS table's order.
    ReDim order(1 To FR_COUNT)
    For i = 1 To FR_COUNT
        order(i) = i
    Next i
    For i = 1 To FR_COUNT - 1
        For k = i + 1 To FR_COUNT
            If FR_NAME(order(k)) < FR_NAME(order(i)) Then
                r = order(i): order(i) = order(k): order(k) = r
            End If
        Next k
    Next i

    Set t1 = ResultCombos(1)
    Set t2 = ResultCombos(2)
    ReDim data1(1 To t1.Count * FR_COUNT * 3, 1 To 9)
    ReDim data2(1 To t2.Count * 2 * FR_COUNT * 3, 1 To 9)

    r = 0
    For Each c In t1
        res = FillTableRows(data1, r, order, CStr(c), "", CStr(c), idx, sta, vals, cnt)
        If Len(res) > 0 Then WriteSapResults = res: Exit Function
    Next c
    r = 0
    For Each stepName In Array("Max", "Min")
        For Each c In t2
            res = FillTableRows(data2, r, order, CStr(c), CStr(stepName), _
                                c & "(" & LCase$(stepName) & ")", idx, sta, vals, cnt)
            If Len(res) > 0 Then WriteSapResults = res: Exit Function
        Next c
    Next stepName

    WriteSapResults = PutResultsSheet(data1, data2)

End Function

' Appends one combination's rows - every frame in order, I / 2/4 / J - to
' data at row r. The three stations are sorted along the frame.
Private Function FillTableRows(ByRef data() As Variant, ByRef r As Long, ByRef order() As Long, _
                               ByVal comboName As String, ByVal stepName As String, _
                               ByVal loadLabel As String, ByVal idx As Object, _
                               ByRef sta() As Double, ByRef vals() As Double, _
                               ByRef cnt() As Long) As String

    Dim i As Long, p As Long, q As Long, e As Long
    Dim key As String
    Dim pos(1 To 3) As Long
    Dim tmp As Long

    For i = 1 To FR_COUNT
        key = FR_NAME(order(i)) & "|" & comboName & "|" & stepName
        If Not idx.Exists(key) Then
            FillTableRows = "SAP2000 returned no " & comboName & IIf(Len(stepName) > 0, " " & stepName, "") & _
                            " forces for frame " & FR_NAME(order(i)) & "."
            Exit Function
        End If
        e = idx(key)
        If cnt(e) <> 3 Then
            FillTableRows = "frame " & FR_NAME(order(i)) & ", " & comboName & ": " & cnt(e) & _
                            " output stations, expected 3 (I, middle, J)."
            Exit Function
        End If

        pos(1) = 1: pos(2) = 2: pos(3) = 3
        For p = 1 To 2
            For q = p + 1 To 3
                If sta(e, pos(q)) < sta(e, pos(p)) Then tmp = pos(p): pos(p) = pos(q): pos(q) = tmp
            Next q
        Next p

        For p = 1 To 3
            r = r + 1
            data(r, 1) = FR_NAME(order(i))
            data(r, 2) = loadLabel
            Select Case p
                Case 1: data(r, 3) = "I[" & JT_NAME(FR_I(order(i))) & "]"
                Case 2: data(r, 3) = "2/4"
                Case 3: data(r, 3) = "J[" & JT_NAME(FR_J(order(i))) & "]"
            End Select
            ' SAP P V2 V3 T M2 M3 -> MIDAS Axial Shear-y Shear-z Torsion Moment-y Moment-z
            data(r, 4) = Round(vals(e, pos(p), 1), FORCE_ROUND_DP)
            data(r, 5) = Round(vals(e, pos(p), 3), FORCE_ROUND_DP)
            data(r, 6) = Round(vals(e, pos(p), 2), FORCE_ROUND_DP)
            data(r, 7) = Round(vals(e, pos(p), 4), FORCE_ROUND_DP)
            data(r, 8) = Round(vals(e, pos(p), 6), FORCE_ROUND_DP)
            data(r, 9) = Round(vals(e, pos(p), 5), FORCE_ROUND_DP)
        Next p
    Next i

End Function

' MIDAS_RESULTS's two data tables emptied and filled with data1 (B:J from
' row 3) and data2 (S:AA from row 36) - exactly where the MIDAS builder puts
' its results, so the headers, the summary block (read by 6_DONATI) and the
' helper formulas carry on unchanged. The helper formulas (L:O, AC:AF) are
' extended when a table outgrows them, as the MIDAS builder does.
Private Function PutResultsSheet(ByRef data1() As Variant, ByRef data2() As Variant) As String

    Dim ws As Worksheet
    Dim last As Long, n1 As Long, n2 As Long
    Dim stage As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RESULT_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        PutResultsSheet = "sheet " & RESULT_SHEET_NAME & " not found in this workbook."
        Exit Function
    End If

    n1 = UBound(data1, 1)
    n2 = UBound(data2, 1)

    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    stage = "empty the tables"
    last = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If last >= 3 Then ws.Range(ws.Cells(3, 2), ws.Cells(last, 10)).ClearContents
    last = ws.Cells(ws.Rows.Count, 19).End(xlUp).Row
    If last >= 36 Then ws.Range(ws.Cells(36, 19), ws.Cells(last, 27)).ClearContents

    ' Part stays text ("2/4" would otherwise turn into a date).
    stage = "write the tables"
    ws.Range(ws.Cells(3, 4), ws.Cells(2 + n1, 4)).NumberFormat = "@"
    ws.Range(ws.Cells(36, 21), ws.Cells(35 + n2, 21)).NumberFormat = "@"
    ws.Range(ws.Cells(3, 2), ws.Cells(2 + n1, 10)).Value = data1
    ws.Range(ws.Cells(36, 19), ws.Cells(35 + n2, 27)).Value = data2

    stage = "extend the helper formulas"
    last = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row
    If last < 2 + n1 And last >= 3 Then
        ws.Range(ws.Cells(3, 12), ws.Cells(3, 15)).AutoFill _
            Destination:=ws.Range(ws.Cells(3, 12), ws.Cells(2 + n1, 15))
    End If
    last = ws.Cells(ws.Rows.Count, 29).End(xlUp).Row
    If last < 35 + n2 And last >= 36 Then
        ws.Range(ws.Cells(36, 29), ws.Cells(36, 32)).AutoFill _
            Destination:=ws.Range(ws.Cells(36, 29), ws.Cells(35 + n2, 32))
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Function

    ' A half-written table would mix the previous results with SAP's - leave
    ' both empty instead.
Failed:
    PutResultsSheet = stage & ": VBA error " & Err.Number & " - " & Err.Description & _
                      " - " & RESULT_SHEET_NAME & "'s tables were left empty."
    On Error Resume Next
    ws.Range(ws.Cells(3, 2), ws.Cells(ws.Rows.Count, 10)).ClearContents
    ws.Range(ws.Cells(36, 19), ws.Cells(ws.Rows.Count, 27)).ClearContents
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Function

' C30/37: Ec (MIDAS_INPUT!B43, else 33 GPa), unit weight (INPUT!B5), fc -
' the values the .$2k's material tables carried. SAP derives G and the mass.
Private Sub EmitMaterial()

    Dim nm As String
    nm = PsQ(MATERIAL_NAME)

    Call Emit("  SapChk 'Material' ($m.PropMaterial.SetMaterial(" & nm & ", 2, -1, '', ''))")
    Call Emit("  SapChk 'Material E, nu, alpha' ($m.PropMaterial.SetMPIsotropic(" & nm & ", " & _
              SapNum(MATERIAL_ELAST) & ", " & SapNum(MATERIAL_POISN) & ", " & SapNum(MATERIAL_THERMAL) & ", 0))")
    Call Emit("  SapChk 'Material unit weight' ($m.PropMaterial.SetWeightAndMass(" & nm & ", 1, " & _
              SapNum(MATERIAL_DEN) & ", 0))")
    Call Emit("  SapChk 'Material fc' ($m.PropMaterial.SetOConcrete_1(" & nm & ", " & SapNum(MATERIAL_FC) & _
              ", $false, 1, 2, 2, 0.00181818, 0.005, -0.1, 0, 0, 0))")

End Sub

' Every section is a solid rectangle depth x 1.0 m (the per-metre strip).
Private Sub EmitSections()

    Dim rows() As String
    Dim f() As String
    Dim i As Long

    rows = Split(SECT_LIST, ";")
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
        Call Emit("  SapChk " & PsQ("Section " & SectionNameOf(CLng(f(0)))) & " ($m.PropFrame.SetRectangle(" & _
                  PsQ(SectionNameOf(CLng(f(0)))) & ", " & PsQ(MATERIAL_NAME) & ", " & SapNum(Val(f(2))) & ", " & _
                  SapNum(SECTION_WIDTH) & ", " & SectionColorOf(CLng(f(0))) & ", '', ''))")
    Next i

End Sub

' Joints and frames under the MIDAS numbers ExpandModel gave them (every
' frame gets 3 output stations: I end, middle, J end - the MIDAS table's
' I, 2/4, J). MIDAS beta -180 (left wall) is SAP's 180; 0 is left alone.
Private Sub EmitJointsAndFrames()

    Dim jt As Long, fr As Long
    Dim a As Double

    For jt = 1 To JT_COUNT
        Call Emit("  SapPt '" & JT_NAME(jt) & "' " & PsNum(JT_X(jt)) & " " & PsNum(JT_Z(jt)))
    Next jt

    For fr = 1 To FR_COUNT
        a = FR_ANGLE(fr)
        Do While a > 180: a = a - 360: Loop
        Do While a <= -180: a = a + 360: Loop
        Call Emit("  SapFr '" & FR_NAME(fr) & "' '" & JT_NAME(FR_I(fr)) & "' '" & JT_NAME(FR_J(fr)) & "' " & _
                  PsQ(SectionNameOf(FR_SECT(fr))) & " " & PsNum(a))
    Next fr

End Sub

' The MIDAS builder's db/NSPR rule on every joint at z = 0, sorted by X:
'   interior   L = (x[i+1] - x[i-1]) / 2
'   end        L = (x[2] - x[1]) / 2, plus half a wall thickness when the
'              foundation stops at the wall centreline (no extension)
'   Kz = ks * L,  Kx = Ky = Kz / 2
' The reference model's springs follow exactly this rule.
Private Function EmitSprings() As String

    Dim ids() As Long
    Dim xs() As Double
    Dim n As Long, jt As Long
    Dim i As Long, j As Long
    Dim tmpId As Long, tmpX As Double
    Dim ltrib As Double, kz As Double

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
        EmitSprings = "found " & n & " foundation joints at z = 0 (expected at least 2)."
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
        Call Emit("  SapSpr '" & JT_NAME(ids(i)) & "' " & PsNum(kz / 2) & " " & PsNum(kz))
    Next i

    SPRING_COUNT = n

End Function

' One pattern (and its linear static case) per MIDAS static case, less
' EQ/ATA when the seismic gate is off and less LL/LLin/LLacc/LSS2_L/
' LSA2_L when the live-load gate is off; DL alone carries self-weight.
' The blank model's own DEAD pattern and case go once DL exists.
Private Sub EmitLoadPatterns()

    Dim rows() As String
    Dim f() As String
    Dim i As Long

    rows = Split(STLDCASE_LIST, ";")
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
        If (SEISMIC_ACTIVE Or Not IsSeismicCase(f(0))) And _
           (LIVE_LOAD_ACTIVE Or Not IsLiveLoadCase(f(0))) Then
            Call Emit("  SapPat " & PsQ(SapName(f(0))) & " " & SapPatternTypeCode(SapDesignType(f(0), f(1))) & _
                      " " & IIf(f(0) = "DL", "1", "0"))
        End If
    Next i
    Call Emit("  SapChk 'Remove the default DEAD case' ($m.LoadCases.Delete('DEAD'))")
    Call Emit("  SapChk 'Remove the default DEAD pattern' ($m.LoadPatterns.Delete('DEAD'))")

End Sub

' BEAMLOAD_LIST rows are "elem|case|cmd|dir|p1|p2", p1 at the element's I
' end and p2 at its J end over the full length (MIDAS D = [0,1]). A divided
' element's pieces each get their slice of that trapezoid.
'   GZ (global Z, MIDAS writes -v)  ->  gravity (dir 10), +v
'   LZ (member local z)             ->  local 2 (dir 2), same sign
' The LZ mapping is the reference model's: its walls run top-down like the
' MIDAS ones, the left wall at 180 degrees, carrying the MIDAS values.
Private Function EmitDistributedLoads() As String

    Dim rows() As String
    Dim f() As String
    Dim i As Long, k As Long
    Dim elemNo As Long, pieces As Long
    Dim p1 As Double, p2 As Double
    Dim loadSign As Double, loadDir As Long

    If Len(BEAMLOAD_LIST) = 0 Then
        EmitDistributedLoads = "no beam loads were generated."
        Exit Function
    End If

    rows = Split(BEAMLOAD_LIST, ";")
    For i = 0 To UBound(rows)

        f = Split(rows(i), "|")
        elemNo = CLng(f(0))

        Select Case f(3)
            Case "GZ": loadDir = 10: loadSign = -1
            Case "LZ": loadDir = 2: loadSign = 1
            Case Else
                EmitDistributedLoads = "element " & elemNo & ", case " & f(1) & _
                    ": load direction """ & f(3) & """ has no SAP2000 mapping."
                Exit Function
        End Select

        If elemNo > UBound(ELEM_PIECES) Then
            EmitDistributedLoads = "a load names element " & elemNo & ", which does not exist."
            Exit Function
        End If
        pieces = ELEM_PIECES(elemNo)
        If pieces = 0 Then
            EmitDistributedLoads = "a load names element " & elemNo & ", which does not exist."
            Exit Function
        End If

        p1 = Val(f(4))
        p2 = Val(f(5))
        For k = 0 To pieces - 1
            Call Emit("  SapLd '" & FR_NAME(ELEM_FIRST_FRAME(elemNo) + k) & "' " & PsQ(SapName(f(1))) & " " & loadDir & " " & _
                      PsNum(Round(loadSign * (p1 + (p2 - p1) * k / pieces), LOAD_ROUND_DP)) & " " & _
                      PsNum(Round(loadSign * (p1 + (p2 - p1) * (k + 1) / pieces), LOAD_ROUND_DP)))
            LOAD_COUNT = LOAD_COUNT + 1
        Next k

    Next i

End Function

' ATA: the MIDAS db/BODF record FV = [kh, 0, 0] - self-weight along global
' X. Seismic only, like the ATA pattern itself; with kh (B15) = 0 the ATA
' pattern stays but carries nothing, as in the MIDAS builder.
Private Sub EmitSeismicSelfWeight()

    Dim fr As Long

    If Not SEISMIC_ACTIVE Or DIM_ATA_FACTOR = 0 Then Exit Sub
    For fr = 1 To FR_COUNT
        Call Emit("  SapGr '" & FR_NAME(fr) & "' 'ATA' " & PsNum(DIM_ATA_FACTOR))
    Next fr

End Sub

' "ST" entries name a load pattern's case (SapName drops the "_"), "CB"
' another combination, spelled as in MIDAS. LOADCOMB_LIST is in dependency
' order (ENV_* after the combinations they envelope).
Private Function EmitCombinations() As String

    Dim rows() As String
    Dim f() As String
    Dim items() As String
    Dim it() As String
    Dim i As Long, j As Long

    If Len(LOADCOMB_LIST) = 0 Then
        EmitCombinations = "no load combinations were generated."
        Exit Function
    End If

    rows = Split(LOADCOMB_LIST, ";")
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
        Call Emit("  SapCmb " & PsQ(f(0)) & " " & IIf(f(2) = "1", "1", "0"))
        items = Split(f(3), ",")
        For j = 0 To UBound(items)
            it = Split(items(j), ":")
            If it(0) = "ST" Then
                Call Emit("  SapCmbAdd " & PsQ(f(0)) & " 0 " & PsQ(SapName(it(1))) & " " & PsNum(Val(it(2))))
            ElseIf it(0) = "CB" Then
                Call Emit("  SapCmbAdd " & PsQ(f(0)) & " 1 " & PsQ(it(1)) & " " & PsNum(Val(it(2))))
            Else
                EmitCombinations = f(0) & ": unknown reference type """ & it(0) & """."
                Exit Function
            End If
        Next j
    Next i

End Function

' Company fixed; project name, engineer and revision from the INPUT sheet;
' model name = the workbook; model description = this module and version
' (the tests read it back from SAP2000's own .$2k export).
Private Sub EmitProjectInfo()

    Call EmitInfo("Company Name", PROJINFO_COMPANY)
    Call EmitInfo("Project Name", PROJECT_NAME)
    Call EmitInfo("Model Name", ThisWorkbook.Name)
    Call EmitInfo("Model Description", SCRIPT_ID & " " & SCRIPT_VERSION)
    Call EmitInfo("Revision Number", PROJECT_REVISION)
    Call EmitInfo("Engineer", PROJECT_ENGINEER)

End Sub

Private Sub EmitInfo(ByVal item As String, ByVal value As String)
    If Len(Trim$(value)) = 0 Then Exit Sub
    Call Emit("  SapChk " & PsQ("Project information " & item) & " ($m.SetProjectInfo(" & PsQ(item) & ", " & _
              PsQ(Trim$(value)) & "))")
End Sub

' Section names must be unique (two sharing one would silently merge) and
' every frame's section must be defined.
Private Function ValidateSections() As String

    Dim rows() As String
    Dim i As Long, j As Long, fr As Long
    Dim nm As String, nm2 As String

    rows = Split(SECT_LIST, ";")
    For i = 0 To UBound(rows)
        nm = SectionNameOf(CLng(Split(rows(i), "|")(0)))
        For j = i + 1 To UBound(rows)
            nm2 = SectionNameOf(CLng(Split(rows(j), "|")(0)))
            If StrComp(nm, nm2, vbTextCompare) = 0 Then
                ValidateSections = "two sections are both named """ & nm & """ - rename one in " & _
                                   INPUT_SHEET_NAME & " column A."
                Exit Function
            End If
        Next j
    Next i

    For fr = 1 To FR_COUNT
        If Len(SectionNameOf(FR_SECT(fr))) = 0 Then
            ValidateSections = "frame " & FR_NAME(fr) & " uses section " & FR_SECT(fr) & ", which was not defined."
            Exit Function
        End If
    Next fr

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

' SECTION_COLOR_LIST's RGB as the integer SAP2000 stores (R + 256 G + 65536
' B); -1 (SAP's own choice) for a section without one.
Private Function SectionColorOf(ByVal sectNo As Long) As String

    Dim entries() As String
    Dim f() As String
    Dim c() As String
    Dim i As Long

    SectionColorOf = "-1"
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

' SAP2000's eLoadPatternType for a design type name.
Private Function SapPatternTypeCode(ByVal designType As String) As String
    Select Case designType
        Case "DEAD": SapPatternTypeCode = "1"
        Case "LIVE": SapPatternTypeCode = "3"
        Case "QUAKE": SapPatternTypeCode = "5"
        Case Else: SapPatternTypeCode = "8"
    End Select
End Function


' ===========================================================================
'  SAP2000 API BRIDGE - identical in both SAP2000 builders (drift-locked).
'
'  Excel is 64-bit; the SAP2000 17 API is a 32-bit .NET class that VBA
'  cannot call. So the macro writes a PowerShell script, runs it hidden with
'  Windows' own 32-bit PowerShell, and waits. The script starts its OWN
'  SAP2000 (never one the user has open), builds or opens the model, saves
'  the .sdb, runs the analysis, shows the window and exits - SAP2000 stays
'  open. It reports one line per outcome in a log the macro reads back:
'      OK|<step>      FAIL|<message>      DONE
'  Every fact behind this (the 32-bit class, silent renaming, the UTF-8 BOM,
'  the modal-dialog hang) is in CLAUDE.md, "SAP2000 audit ... and the API".
' ===========================================================================

' <workbook folder>\SAP2000, created when missing; "" if the workbook has
' never been saved. FileSystemObject, not Dir/MkDir, so a folder name with
' characters outside the code page still works.
Private Function SapFolder() As String

    Dim fso As Object
    Dim p As String

    If Len(ThisWorkbook.Path) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    p = ThisWorkbook.Path & "\" & SAP_FOLDER_NAME
    If Not fso.FolderExists(p) Then fso.CreateFolder p
    SapFolder = p

End Function

' <SAP2000 folder>\<workbook name without extension> - the stem of the
' .sdb, the build script and its log.
Private Function SapBasePath(ByVal folder As String) As String

    Dim nm As String
    Dim p As Long

    nm = ThisWorkbook.Name
    p = InStrRev(nm, ".")
    If p > 1 Then nm = Left$(nm, p - 1)
    SapBasePath = folder & "\" & nm

End Function

' A PowerShell single-quoted literal: nothing inside is interpreted, and a
' single quote is written twice.
Private Function PsQ(ByVal s As String) As String
    PsQ = "'" & Replace(s, "'", "''") & "'"
End Function

' A number as a PowerShell argument: locale-invariant, in parentheses so a
' leading minus can never read as a parameter name.
Private Function PsNum(ByVal v As Double) As String
    PsNum = "(" & SapNum(v) & ")"
End Function

Private Function PsBool(ByVal b As Boolean) As String
    PsBool = IIf(b, "$true", "$false")
End Function

' The start of every build script: helpers, and a fresh hidden (or visible,
' SAP_VISIBLE) SAP2000 with a blank kN-m model. Everything after it runs
' inside the try block that EmitSapFinish closes.
Private Sub EmitSapPreamble(ByVal resultPath As String)

    Call Emit("# Written by " & SCRIPT_ID & " " & SCRIPT_VERSION & " - run by the macro with 32-bit PowerShell.")
    Call Emit("# It starts its own SAP2000, builds the model, saves, analyses and leaves SAP2000 open.")
    Call Emit("$ErrorActionPreference = 'Stop'")
    Call Emit("$sapLog = " & PsQ(resultPath))
    Call Emit("function SapLog($s) { Add-Content -LiteralPath $sapLog -Value $s -Encoding UTF8 }")
    Call Emit("function SapChk($what, $ret) { if ($ret -ne 0) { throw ""$what failed (SAP2000 returned $ret)"" } }")
    Call Emit("function SapNamed($what, $want, $got) { if ($got -ne $want) { throw ""$what came back named '$got', not '$want'"" } }")
    Call Emit("function SapCount($what, $obj, $want) { $n = 0; $a = [string[]]@(); SapChk ""Count $what"" ($obj.GetNameList([ref]$n, [ref]$a)); if ($n -ne $want) { throw ""SAP2000 holds $n $what, the macro made $want"" } }")
    Call Emit("$sap = $null")
    Call Emit("try {")
    Call Emit("  $sap = New-Object -ComObject CSI.SAP2000.API.SapObject")
    Call Emit("  SapChk 'Start SAP2000' ($sap.ApplicationStart(6, " & PsBool(SAP_VISIBLE) & ", ''))")
    Call Emit("  $m = $sap.SapModel")
    Call Emit("  SapChk 'New model' ($m.InitializeNewModel(6))")
    Call Emit("  SapChk 'Blank model' ($m.File.NewBlank())")
    Call Emit("  SapChk 'Units kN, m, C' ($m.SetPresentUnits(6))")
    Call Emit("  SapLog 'OK|SAP2000 started'")

End Sub

' Save, analyse, show, and close the try block. On any failure the script
' closes its SAP2000 again, so nothing half-built is left behind.
' afterAnalysis: script lines run on the solved model before it is shown
' (the culvert pulls its frame forces there); "" for none.
Private Sub EmitSapFinish(ByVal sdbPath As String, Optional ByVal afterAnalysis As String = "")

    Call Emit("  SapChk 'Save' ($m.File.Save(" & PsQ(sdbPath) & "))")
    Call Emit("  SapLog " & PsQ("OK|Saved " & sdbPath))
    Call Emit("  SapChk 'Analysis' ($m.Analyze.RunAnalysis())")
    Call Emit("  SapLog 'OK|Analysis run'")
    If Len(afterAnalysis) > 0 Then Call Emit(afterAnalysis)
    Call Emit("  if (-not $sap.Visible()) { SapChk 'Show SAP2000' ($sap.Unhide()) }")
    Call Emit("  SapLog 'DONE'")
    Call Emit("} catch {")
    Call Emit("  SapLog ('FAIL|' + $_.Exception.Message)")
    Call Emit("  if ($sap) { try { [void]$sap.ApplicationExit($false) } catch { } }")
    Call Emit("}")

End Sub

' OUT_LINES as a UTF-8 file WITH a byte-order mark - Windows PowerShell 5.1
' reads a file without one as ANSI and garbles every non-ASCII character.
Private Function WriteScriptFile(ByVal path As String) As String

    Dim st As Object
    Dim lines() As String
    Dim i As Long

    On Error GoTo Failed

    ReDim lines(1 To OUT_COUNT)
    For i = 1 To OUT_COUNT
        lines(i) = OUT_LINES(i)
    Next i

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText Join(lines, vbCrLf) & vbCrLf
    st.SaveToFile path, 2
    st.Close
    Exit Function

Failed:
    WriteScriptFile = "could not write " & path & " (" & Err.Description & ")."

End Function

Private Function ReadUtf8File(ByVal path As String) As String

    Dim st As Object

    On Error GoTo Failed
    If Not CreateObject("Scripting.FileSystemObject").FileExists(path) Then Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8File = st.ReadText
    st.Close
    Exit Function

Failed:
    ReadUtf8File = ""

End Function

' Runs the script in 32-bit PowerShell and waits for it (SAP_TIMEOUT_SEC at
' most), keeping Excel responsive. Returns "" when the log ends in DONE,
' otherwise the reason. On a timeout it ends the PowerShell process and the
' SAP2000 it started - a hidden SAP2000 stuck on a dialog would otherwise
' block forever - and nothing else.
Private Function RunSapScript(ByVal scriptPath As String, ByVal resultPath As String) As String

    Dim fso As Object, wmi As Object, startup As Object
    Dim psExe As String, cmd As String
    Dim pid As Variant
    Dim rc As Long
    Dim t0 As Single, waited As Single
    Dim outcome As String
    Dim p As Long, q As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(resultPath) Then fso.DeleteFile resultPath, True

    psExe = Environ$("WINDIR") & "\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    If Not fso.FileExists(psExe) Then psExe = Environ$("WINDIR") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
    cmd = """" & psExe & """ -NoProfile -NonInteractive -ExecutionPolicy Bypass -File """ & scriptPath & """"

    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set startup = wmi.Get("Win32_ProcessStartup").SpawnInstance_
    startup.ShowWindow = 0
    rc = wmi.Get("Win32_Process").Create(cmd, Null, startup, pid)
    If rc <> 0 Then
        RunSapScript = "could not start 32-bit PowerShell (" & psExe & ", WMI code " & rc & ")."
        Exit Function
    End If

    t0 = Timer
    Do
        waited = SecondsSince(t0)
        If wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE ProcessId = " & pid).Count = 0 Then Exit Do
        If waited > SAP_TIMEOUT_SEC Then
            Call EndSapScript(wmi, pid)
            Application.StatusBar = False
            RunSapScript = "SAP2000 did not finish within " & SAP_TIMEOUT_SEC & " s - most likely a " & _
                           "SAP2000 dialog was waiting for an answer (an unstable or degenerate model does " & _
                           "that). The script and its SAP2000 were closed. Last step reached: " & _
                           LastOkStep(ReadUtf8File(resultPath))
            Exit Function
        End If
        Application.StatusBar = "SAP2000: building and analysing the model - " & Int(waited) & " s"
        Call PauseSeconds(0.5)
    Loop
    Application.StatusBar = False

    outcome = ReadUtf8File(resultPath)
    If InStr(1, outcome, "DONE", vbBinaryCompare) > 0 Then Exit Function

    p = InStr(1, outcome, "FAIL|", vbBinaryCompare)
    If p > 0 Then
        q = InStr(p, outcome, vbCr)
        If q = 0 Then q = Len(outcome) + 1
        RunSapScript = Mid$(outcome, p + 5, q - p - 5)
    Else
        RunSapScript = "the SAP2000 script ended without reporting success. Last step reached: " & _
                       LastOkStep(outcome) & ". Run " & scriptPath & " by hand in 32-bit PowerShell to see why."
    End If

End Function

' Ends one PowerShell process and every SAP2000 it started.
Private Sub EndSapScript(ByVal wmi As Object, ByVal pid As Variant)

    Dim proc As Object

    On Error Resume Next
    For Each proc In wmi.ExecQuery("SELECT * FROM Win32_Process WHERE ParentProcessId = " & pid)
        If LCase$(proc.Name) = "sap2000.exe" Then proc.Terminate
    Next proc
    For Each proc In wmi.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = " & pid)
        proc.Terminate
    Next proc
    On Error GoTo 0

End Sub

Private Function LastOkStep(ByVal outcome As String) As String

    Dim p As Long, q As Long

    p = InStrRev(outcome, "OK|")
    If p = 0 Then
        LastOkStep = "none (SAP2000 may not have started)"
    Else
        q = InStr(p, outcome, vbCr)
        If q = 0 Then q = Len(outcome) + 1
        LastOkStep = Mid$(outcome, p + 3, q - p - 3)
    End If

End Function

' Seconds since t0, across midnight (Timer restarts at 0).
Private Function SecondsSince(ByVal t0 As Single) As Single
    SecondsSince = Timer - t0
    If SecondsSince < 0 Then SecondsSince = SecondsSince + 86400
End Function

' DoEvents + Timer, not Application.Wait (whole seconds only, and it freezes
' Excel).
Private Sub PauseSeconds(ByVal secs As Single)
    Dim t0 As Single
    t0 = Timer
    Do While SecondsSince(t0) < secs
        DoEvents
    Loop
End Sub

' MIDAS case names -> SAP pattern names: EHS2_L -> EHS2L.
Private Function SapName(ByVal midasName As String) As String
    SapName = Replace(midasName, "_", "")
End Function

' Locale-invariant number, rounded so float noise never reaches SAP2000.
Private Function SapNum(ByVal v As Double) As String
    SapNum = JsonNum(Round(v, 9))
End Function

Private Sub Emit(ByVal s As String)
    OUT_COUNT = OUT_COUNT + 1
    If OUT_COUNT > UBound(OUT_LINES) Then ReDim Preserve OUT_LINES(1 To UBound(OUT_LINES) * 2)
    OUT_LINES(OUT_COUNT) = s
End Sub
