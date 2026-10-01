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
'  value) - the same fallback as the MIDAS culvert builder (since 2026-09-28).
'
'  No restraints are written: the foundation springs are the supports, as
'  in the reference. Design-only data of the old text file (column rebar,
'  A615Gr60) is not built - the model is for analysis.
' ============================================================================


' ---------------------------------------------------------------------------
'  CONFIG
' ---------------------------------------------------------------------------

' Stamped into the report title. Bump with every change to this file.
Private Const SCRIPT_VERSION As String = "2026-10-01b"

' One line, no "_" continuation, no "|" - read by the updater's manifest.
Private Const SCRIPT_CHANGELOG As String = "Run record switched on: each build sends one record (inputs, loads, gates, governing section forces; project fields encrypted) to the culvert run database when INPUT!J21/J22 are filled"

' Identifies this module to the updater whatever it was named in Excel.
Private Const SCRIPT_ID As String = "sap2000-culvert-model-build"

' Report window font (ShowReport): 9 pt Segoe UI is the MsgBox font; the
' window's A- / A+ buttons change it on the spot.
Private Const REPORT_FONT_NAME As String = "Segoe UI"
Private Const REPORT_FONT_PT As Long = 9

' ---------------------------------------------------------------------------
'  RUN RECORD (run database, docs/run-database/) - see RecordRun at the end
'  of this module. INPUT!J21 = Firebase project ID, J22 = web API key; both
'  blank = recording off. A failed upload only WARNs.
'  REC_ENABLED False = nothing is collected or sent (shipped off until the
'  owner's Firebase project exists). REC_DUMP_BODY True also writes the
'  body to %TEMP%\<SCRIPT_ID>_record_body.json.
' ---------------------------------------------------------------------------
Private Const REC_ENABLED As Boolean = True
Private Const REC_DUMP_BODY As Boolean = False
Private Const REC_BUILDER As String = "sap2000"
Private Const REC_PROJECT_ID_CELL As String = "J21"
Private Const REC_API_KEY_CELL As String = "J22"
' The owner's RSA-3072 PUBLIC key (.NET XML, ONE line) and its key id from
' scripts/run-database/make_rsa_keys.py. It can only lock, never unlock.
' Blank = the project fields are left out (key_id "none"), never sent plain.
Private Const REC_PUBLIC_KEY_XML As String = "<RSAKeyValue><Modulus>8VDzVu3qkUxA2yJqeExds39SzuTXKvKX18ubYRiq71J8nAlPAtgdrUE/Ko8+OryUhNnDTPh14wCS6EJ1D2opFRPlat4WKrMzVI8JK721Wzf9GVmLHUq52/kXK6MgyLPc6bgbkVEU0JI3TdPvjPn9yU1sd31HyJ5Ff0BoKA05O9sht1pxkwZ3YSi6q0U47w+209KACW9cg1w+qGGIocjcZDpAKhs4oEsucVZyMufNOLJfn8BdzsWU/JsN+g5plGs+BLz8eTxExVJ6++12iC8kTJJ7YVNbMfho2oE7b5cwjHSiO5vvQvXBqRVu/MB7OdiwYQBoRaPrPX9F+nI2yIC5G4ylXzvtF56/m+BiD8V8WZoR6YtHPTW8t5y25IlChZIqykYUN5MRnRcRlcwab/2Bj0hbDC0u2iK7cOVaqOT5vpNVJDPM46B3tzow3OaNuZXZ5lR6o2fsC53+hBgRTA2EnagsDDfSPsPahuyimWRhsoFt8gj9c6CfnfS8rdo3NjVB</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>"
Private Const REC_KEY_ID As String = "04bce0f3fa8473f7"
Private Const REC_UID_PLACEHOLDER As String = "__REC_UID__"
Private Const REC_ENCRYPT_TIMEOUT_SEC As Long = 15
' Refresh, sign-up after a dead refresh token, create, one retry, and one
' re-sign-in after a 401.
Private Const REC_MAX_REQUESTS As Long = 5
' Sign-in cached per Excel session, keyed by project ID + API key.
Private REC_TOKEN_CACHE As String
Private REC_TOKEN_UID As String
Private REC_TOKEN_FOR As String
Private REC_TOKEN_EXPIRES As Date
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef pguid As Byte) As Long
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As Integer)

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

' ---------------------------------------------------------------------------
'  MIN-VERTICAL GATE - identical to the MIDAS builder (see its header):
'  INPUT!K17 = 1 adds ULS-15..20 (DL/EV x0.90, no vertical live load) to
'  the combinations and ENV_STR; 0 or blank -> as before; anything else
'  stops the build.
' ---------------------------------------------------------------------------
Private Const MIN_VERTICAL_GATE_CELL As String = "K17"
Private MIN_VERTICAL_ACTIVE As Boolean
Private MIN_VERTICAL_GATE_SOURCE As String

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
    Dim minVertErr As String
    Dim resultsWritten As Boolean
    Dim recRunId As String, recStartedAt As String, recRes As String
    Dim recT0 As Single

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

    stage = "min-vertical gate"
    minVertErr = InitMinVerticalGate()
    If Len(minVertErr) > 0 Then
        MsgBox report & "STOPPED - nothing was built." & vbCrLf & vbCrLf & _
               minVertErr, vbExclamation
        Exit Sub
    End If
    If MIN_VERTICAL_ACTIVE Then
        report = report & "NOTE - min-vertical combinations ON (" & MIN_VERTICAL_GATE_SOURCE & _
                 "): ULS-15..20 (DL/EV x0.90) added to ENV_STR." & vbCrLf & String(40, "-") & vbCrLf
    End If

    ' The run record's id and start (after the gate stops: a stopped run
    ' sends nothing).
    recRunId = RecNewRunId()
    recStartedAt = RecUtcNow()
    recT0 = Timer
    INPUT_FALLBACKS = ""

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
    If ok Then
        res = WriteSapResults(stem & "_forces.txt")
        resultsWritten = (Len(res) = 0)
        ok = StepResult(report, "Frame forces to " & RESULT_SHEET_NAME, res)
    End If

    If built Then
        report = report & String(40, "-") & vbCrLf & _
                 JT_COUNT & " joints, " & FR_COUNT & " frames (divided: " & DIVIDED_ELEMS & _
                 "), " & SPRING_COUNT & " springs, " & LOAD_COUNT & " frame loads, " & _
                 (UBound(Split(LOADCOMB_LIST, ";")) + 1) & " combinations." & vbCrLf & _
                 "SAP2000 is open with the analysed model:" & vbCrLf & stem & ".sdb" & vbCrLf
    End If

    ' Reads the report before the verdict; its own line goes after it.
    stage = "run record"
    recRes = RecordRun(ok, resultsWritten, report, "", recRunId, recStartedAt, recT0)

    report = VerdictFirst(report, ok)
    If Len(recRes) > 0 Then Call StepResult(report, "Run record", IIf(recRes = "SENT", "", recRes))
    Call ShowReport(report, ok)
    Exit Sub

Crashed:
    Application.StatusBar = False
    report = report & "FAIL - " & stage & ": VBA error " & Err.Number & " - " & _
             Err.Description & vbCrLf
    ' RecordRun has its own error handler (an error in this handler could
    ' not be trapped here).
    If Len(recRunId) > 0 And stage <> "run record" Then
        recRes = RecordRun(False, False, report, stage, recRunId, recStartedAt, recT0)
    End If
    report = VerdictFirst(report, False)
    If Len(recRes) > 0 Then Call StepResult(report, "Run record", IIf(recRes = "SENT", "", recRes))
    Call ShowReport(report, False)

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

' Reads INPUT!K17 and sets MIN_VERTICAL_ACTIVE/MIN_VERTICAL_GATE_SOURCE.
' Returns "" when the cell holds 1, 0 or nothing, otherwise a message
' naming it - the build stops on it, same pattern as InitLiveLoadGate.
Private Function InitMinVerticalGate() As String

    Dim ws As Worksheet
    Dim v As Variant
    Dim s As String

    InitMinVerticalGate = ""
    ' Reset every run - module-level values survive between runs.
    MIN_VERTICAL_ACTIVE = False
    MIN_VERTICAL_GATE_SOURCE = ""

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PROJECT_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        MIN_VERTICAL_GATE_SOURCE = "sheet """ & PROJECT_SHEET_NAME & """ not found, defaulted to OFF"
        Exit Function
    End If

    v = ws.Range(MIN_VERTICAL_GATE_CELL).Value

    If IsError(v) Then
        InitMinVerticalGate = PROJECT_SHEET_NAME & "!" & MIN_VERTICAL_GATE_CELL & _
                              " (min-vertical combinations) is an error value. Set it to 1, 0 or blank."
        Exit Function
    End If

    ' CStr of a whole number has no decimal separator, so "1"/"0" are
    ' locale-safe; a typed text "1" counts the same as the number.
    s = Trim$(CStr(v))

    If s = "1" Then
        MIN_VERTICAL_ACTIVE = True
        MIN_VERTICAL_GATE_SOURCE = PROJECT_SHEET_NAME & "!" & MIN_VERTICAL_GATE_CELL & " = 1"
    ElseIf s = "0" Or Len(s) = 0 Then
        MIN_VERTICAL_GATE_SOURCE = PROJECT_SHEET_NAME & "!" & MIN_VERTICAL_GATE_CELL & " = " & _
                                   IIf(Len(s) = 0, "(blank)", "0")
    Else
        InitMinVerticalGate = PROJECT_SHEET_NAME & "!" & MIN_VERTICAL_GATE_CELL & _
                              " (min-vertical combinations) is """ & s & """. Set it to 1, 0 or blank."
    End If

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
    ReadInputs = InitMinVerticalGate()
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

    ' --- Minimum vertical / maximum horizontal (INPUT!K17 = 1 only) ---
    ' ULS-1/2/5/6/9/10 with DL and EV at the AASHTO minimum 0.90 and no
    ' vertical live load (LL1/LL); the horizontal terms keep their factors.
    ' Named ULS-* so ENV_STR and MIDAS_RESULTS table 1 take them in. With
    ' live load off, ULS-19/20 lose LSS2_L/LSA2_L and repeat ULS-17/18, as
    ' SLS-7 repeats SLS-5 (accepted). Gate off: nothing here, every later
    ' key unchanged.
    If MIN_VERTICAL_ACTIVE Then
        Call AddCombo("ULS-15", 0, "ST", "DL:0.9,EV1:0.9,EHS1:1.5,LSS1_L:1.45", 1)
        Call AddCombo("ULS-16", 0, "ST", "DL:0.9,EV1:0.9,EHS1:1.5", 1)
        Call AddCombo("ULS-17", 0, "ST", "DL:0.9,EV2:0.9,EHS2_L:1.35,EHS2_R:1.35", 1)
        Call AddCombo("ULS-18", 0, "ST", "DL:0.9,EV2:0.9,EHA2_L:1.35,EHA2_R:1.35", 1)
        Call AddCombo("ULS-19", 0, "ST", "DL:0.9,EV2:0.9,EHS2_L:1.35,EHS2_R:1.35,LSS2_L:1.45", 1)
        Call AddCombo("ULS-20", 0, "ST", "DL:0.9,EV2:0.9,EHA2_L:1.35,EHA2_R:1.35,LSA2_L:1.45", 1)
    End If

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
    Next i
    For i = 1 To IIf(MIN_VERTICAL_ACTIVE, 20, 14)
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

' ===========================================================================
'  REPORT WINDOW (shared, byte-identical in every module - see CLAUDE.md)
' ===========================================================================

' Shows a final report whole, every line, in a small window instead of
' MsgBox, which cuts text past ~1024 characters and whose font cannot be
' set (owner's request, 2026-09-29). The text goes to a UTF-8 file in
' %TEMP% and a hidden PowerShell draws the window; the script deletes both
' files. The macro does not wait for it. If anything fails on the way, the
' old MsgBox + FitReport log file is shown instead.
Private Sub ShowReport(ByVal report As String, ByVal ok As Boolean)

    Dim basePath As String
    Dim txtPath As String
    Dim ps1Path As String
    Dim fileNo As Integer
    Dim stm As Object

    On Error GoTo UseMsgBox

    basePath = Environ$("TEMP") & "\" & SCRIPT_ID & "_report_" & Format$(Now, "yyyymmdd_hhnnss")
    txtPath = basePath & ".txt"
    ps1Path = basePath & ".ps1"

    ' UTF-8: a report can carry non-ASCII names and paths.
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText report
    stm.SaveToFile txtPath, 2
    stm.Close

    ' The script itself is plain ASCII.
    fileNo = FreeFile
    Open ps1Path For Output As #fileNo
    Print #fileNo, ReportWindowScript(ok)
    Close #fileNo

    Shell "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & _
          ps1Path & """ """ & txtPath & """", vbHide
    Exit Sub

UseMsgBox:
    MsgBox FitReport(report, SCRIPT_ID), IIf(ok, vbInformation, vbExclamation)

End Sub

' The PowerShell that draws the report window (generated from a tested
' script - keep it plain ASCII; symbols are [char] codes). DPI-aware, so
' 9 pt is the size MsgBox shows. The window title is the report's title
' line minus its [version]. Line styling by prefix: OK / WARN / FAIL / NOTE
' step lines get a green tick / orange warning / red cross / blue info; a
' FAILED: / WARNINGS: / SKIPPED / OUT OF DATE / OK: / UP TO DATE / UPDATED
' heading colours itself and gives its '  - ' items its symbol; gate
' ON/OFF lines a filled/empty circle; title and verdict bold, dashed rules
' grey, indented continuation lines grey. A- / A+ zoom; the window fits
' the text up to 80 % of the screen, scrollbar beyond.
Private Function ReportWindowScript(ByVal ok As Boolean) As String

    Dim s As String

    s = "param([string]$TextPath)" & vbCrLf
    s = s & "$ErrorActionPreference = 'Stop'" & vbCrLf
    s = s & "$text = [IO.File]::ReadAllText($TextPath, [Text.Encoding]::UTF8)" & vbCrLf
    ' [IO.File]::Delete, not Remove-Item: PowerShell 5.1's Remove-Item rejects the
    ' 8.3 %TEMP% path (GKAY~1) with a terminating error, which killed the
    ' script before the window opened (culvert builder 29c, 2026-09-29).
    s = s & "try { [IO.File]::Delete($TextPath); [IO.File]::Delete($PSCommandPath) } catch { }" & vbCrLf
    s = s & "Add-Type -AssemblyName System.Windows.Forms, System.Drawing" & vbCrLf
    s = s & "Add-Type -Namespace CulvertReport -Name Dpi -MemberDefinition '[DllImport(""user32.dll"")] public static extern bool SetProcessDPIAware();'" & vbCrLf
    s = s & "[void][CulvertReport.Dpi]::SetProcessDPIAware()" & vbCrLf
    s = s & "[Windows.Forms.Application]::EnableVisualStyles()" & vbCrLf
    s = s & "$fontName = '" & REPORT_FONT_NAME & "'" & vbCrLf
    s = s & "$pt = " & REPORT_FONT_PT & vbCrLf
    s = s & "$plain = New-Object Drawing.Font($fontName, $pt)" & vbCrLf
    s = s & "$bold = New-Object Drawing.Font($fontName, $pt, [Drawing.FontStyle]::Bold)" & vbCrLf
    s = s & "$title = New-Object Drawing.Font($fontName, ($pt + 2), [Drawing.FontStyle]::Bold)" & vbCrLf
    s = s & "$symFont = New-Object Drawing.Font('Segoe UI Symbol', $pt)" & vbCrLf
    s = s & "$green = [Drawing.Color]::FromArgb(16, 124, 16)" & vbCrLf
    s = s & "$orange = [Drawing.Color]::FromArgb(200, 122, 0)" & vbCrLf
    s = s & "$red = [Drawing.Color]::FromArgb(196, 43, 28)" & vbCrLf
    s = s & "$blue = [Drawing.Color]::FromArgb(0, 90, 158)" & vbCrLf
    s = s & "$grey = [Drawing.Color]::FromArgb(150, 150, 150)" & vbCrLf
    s = s & "$ink = [Drawing.SystemColors]::WindowText" & vbCrLf
    s = s & "$lines = $text -split '\r?\n'" & vbCrLf
    s = s & "$caption = 'MIDAS macros'" & vbCrLf
    s = s & "foreach ($l in $lines) { if ($l -match '^(.+?)\s+\[\d{4}-\d{2}-\d{2}\w*\]') { $caption = $matches[1]; break } }" & vbCrLf
    s = s & "$form = New-Object Windows.Forms.Form" & vbCrLf
    s = s & "$form.Text = $caption" & vbCrLf
    s = s & "$form.Icon = [Drawing.SystemIcons]::" & IIf(ok, "Information", "Warning") & vbCrLf
    s = s & "$form.StartPosition = 'CenterScreen'" & vbCrLf
    s = s & "$form.TopMost = $true" & vbCrLf
    s = s & "$form.MinimizeBox = $false" & vbCrLf
    s = s & "$form.BackColor = [Drawing.SystemColors]::Window" & vbCrLf
    s = s & "$form.Font = New-Object Drawing.Font($fontName, 9)" & vbCrLf
    s = s & "$box = New-Object Windows.Forms.RichTextBox" & vbCrLf
    s = s & "$box.ReadOnly = $true" & vbCrLf
    s = s & "$box.DetectUrls = $false" & vbCrLf
    s = s & "$box.WordWrap = $true" & vbCrLf
    s = s & "$box.ScrollBars = 'Vertical'" & vbCrLf
    s = s & "$box.BorderStyle = 'None'" & vbCrLf
    s = s & "$box.BackColor = [Drawing.SystemColors]::Window" & vbCrLf
    s = s & "$box.Dock = 'Fill'" & vbCrLf
    s = s & "$box.Font = $plain" & vbCrLf
    s = s & "$box.TabStop = $false" & vbCrLf
    s = s & "$nl = [string][char]10" & vbCrLf
    s = s & "$sb = New-Object Text.StringBuilder" & vbCrLf
    s = s & "$runs = New-Object Collections.ArrayList" & vbCrLf
    s = s & "function Add-Run([string]$s, $font, $color) {" & vbCrLf
    s = s & "  [void]$runs.Add(@($sb.Length, $s.Length, $font, $color))" & vbCrLf
    s = s & "  [void]$sb.Append($s)" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function Add-Line($sym, $symColor, [string]$rest, $font, $color) {" & vbCrLf
    s = s & "  if ($null -eq $color) { $color = $ink }" & vbCrLf
    s = s & "  if ($sym) { Add-Run ([string]$sym + '  ') $symFont $symColor }" & vbCrLf
    s = s & "  Add-Run ($rest + $nl) $font $color" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "$tick = [char]0x2714; $cross = [char]0x2716; $warn = [char]0x26A0; $info = [char]0x2139" & vbCrLf
    s = s & "$ring = [char]0x25CB; $dot = [char]0x25C9; $cycle = [char]0x21BB" & vbCrLf
    s = s & "$secSym = $null; $secColor = $ink" & vbCrLf
    s = s & "foreach ($l in $lines) {" & vbCrLf
    s = s & "  if ($l -match '^.+\s+\[\d{4}-\d{2}-\d{2}\w*\]') { Add-Line $null $ink $l $title }" & vbCrLf
    s = s & "  elseif ($l -match '^OK   - (.*)$') { Add-Line $tick $green $matches[1] $plain }" & vbCrLf
    s = s & "  elseif ($l -match '^WARN - (.*)$') { Add-Line $warn $orange $matches[1] $plain }" & vbCrLf
    s = s & "  elseif ($l -match '^FAIL - (.*)$') { Add-Line $cross $red $matches[1] $plain }" & vbCrLf
    s = s & "  elseif ($l -match '^NOTE - (.*)$') { Add-Line $info $blue $matches[1] $plain }" & vbCrLf
    s = s & "  elseif ($l -match '^[-=]{10,}$') { Add-Run (([string][char]0x2500 * 40) + $nl) $plain $grey }" & vbCrLf
    s = s & "  elseif ($l -match '^(All steps completed|OK - )') { Add-Line $tick $green $l $bold }" & vbCrLf
    s = s & "  elseif ($l -match '^(STOPPED|CANNOT CHECK)') { Add-Line $cross $red $l $bold }" & vbCrLf
    s = s & "  elseif ($l -match '^(FAILED|FAIL)\b.*:\s*$') { $secSym = $cross; $secColor = $red; Add-Line $null $ink $l $bold $red }" & vbCrLf
    s = s & "  elseif ($l -match '^WARNINGS?\b.*:\s*$') { $secSym = $warn; $secColor = $orange; Add-Line $null $ink $l $bold $orange }" & vbCrLf
    s = s & "  elseif ($l -match '^(SKIPPED|NOT MANAGED|AHEAD OF SERVER)\b.*:\s*$') { $secSym = $ring; $secColor = $grey; Add-Line $null $ink $l $bold $grey }" & vbCrLf
    s = s & "  elseif ($l -match '^(OUT OF DATE|UPDATER ITSELF)\b.*:\s*$') { $secSym = $cycle; $secColor = $blue; Add-Line $null $ink $l $bold $blue }" & vbCrLf
    s = s & "  elseif ($l -match '^(OK|UP TO DATE|UPDATED)\b.*:\s*$') { $secSym = $tick; $secColor = $green; Add-Line $null $ink $l $bold $green }" & vbCrLf
    s = s & "  elseif ($l -match '^\s+- (.*)$' -and $secSym) { Add-Run '    ' $plain $ink; Add-Line $secSym $secColor $matches[1] $plain }" & vbCrLf
    s = s & "  elseif ($l -match '^\s{4,}\S') { Add-Line $null $ink $l $plain $grey }" & vbCrLf
    s = s & "  elseif ($l -match '^WARN(ING)?\b') { Add-Line $warn $orange $l $plain }" & vbCrLf
    s = s & "  elseif ($l -match '^(Seismic|Live load|Min-vertical) ON\b') { Add-Line $dot $blue $l $plain }" & vbCrLf
    s = s & "  elseif ($l -match '^(Seismic|Live load|Min-vertical) OFF\b') { Add-Line $ring $grey $l $plain }" & vbCrLf
    s = s & "  elseif ($l -match '^Clean pass.*: OK$') { Add-Line $tick $green $l $plain }" & vbCrLf
    s = s & "  elseif ($l -match '^\d+ captures?: ') { Add-Line $null $ink $l $bold }" & vbCrLf
    s = s & "  else { Add-Line $null $ink $l $plain }" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "$box.Text = $sb.ToString()" & vbCrLf
    s = s & "foreach ($r in $runs) { $box.Select($r[0], $r[1]); $box.SelectionFont = $r[2]; $box.SelectionColor = $r[3] }" & vbCrLf
    s = s & "$box.Select(0, 0)" & vbCrLf
    s = s & "$pad = New-Object Windows.Forms.Panel" & vbCrLf
    s = s & "$pad.Dock = 'Fill'" & vbCrLf
    s = s & "$pad.Padding = New-Object Windows.Forms.Padding(16, 14, 6, 6)" & vbCrLf
    s = s & "$pad.Controls.Add($box)" & vbCrLf
    s = s & "$bar = New-Object Windows.Forms.FlowLayoutPanel" & vbCrLf
    s = s & "$bar.Dock = 'Bottom'" & vbCrLf
    s = s & "$bar.FlowDirection = 'RightToLeft'" & vbCrLf
    s = s & "$bar.AutoSize = $true" & vbCrLf
    s = s & "$bar.Padding = New-Object Windows.Forms.Padding(8)" & vbCrLf
    s = s & "$bar.BackColor = [Drawing.SystemColors]::Control" & vbCrLf
    s = s & "function New-Btn([string]$label, [int]$width) {" & vbCrLf
    s = s & "  $b = New-Object Windows.Forms.Button; $b.Text = $label; $b.Width = $width; $b.Height = 28; $b" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "$okBtn = New-Btn 'OK' 88" & vbCrLf
    s = s & "$okBtn.DialogResult = 'OK'" & vbCrLf
    s = s & "$bigger = New-Btn 'A+' 40" & vbCrLf
    s = s & "$smaller = New-Btn 'A-' 40" & vbCrLf
    s = s & "$bigger.Add_Click({ $box.ZoomFactor = [Math]::Min(3.0, $box.ZoomFactor + 0.1) })" & vbCrLf
    s = s & "$smaller.Add_Click({ $box.ZoomFactor = [Math]::Max(0.6, $box.ZoomFactor - 0.1) })" & vbCrLf
    s = s & "$bar.Controls.AddRange(@($okBtn, $bigger, $smaller))" & vbCrLf
    s = s & "$form.Controls.Add($pad)" & vbCrLf
    s = s & "$form.Controls.Add($bar)" & vbCrLf
    s = s & "$form.AcceptButton = $okBtn" & vbCrLf
    s = s & "$form.CancelButton = $okBtn" & vbCrLf
    s = s & "$area = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea" & vbCrLf
    s = s & "$need = [Windows.Forms.TextRenderer]::MeasureText($text, $bold)" & vbCrLf
    s = s & "$w = [Math]::Max(420, [Math]::Min($need.Width + 110, [int]($area.Width * 0.8)))" & vbCrLf
    s = s & "$h = [Math]::Min($need.Height + $bar.PreferredSize.Height + 70, [int]($area.Height * 0.8))" & vbCrLf
    s = s & "$form.ClientSize = New-Object Drawing.Size($w, $h)" & vbCrLf
    s = s & "$form.Add_Shown({ $form.Activate(); $form.ActiveControl = $okBtn; $box.Refresh() })" & vbCrLf
    s = s & "[void]$form.ShowDialog()" & vbCrLf

    ReportWindowScript = s

End Function


' ===========================================================================
'  RUN RECORD - the run database (docs/run-database/)
'
'  After each build, one record of the run - inputs, loads, gates and the
'  governing section forces, the project fields encrypted - is created in
'  the owner's Firestore project. scripts/run-database/record_body.py is
'  the REFERENCE for the body; docs/run-database/recordrun-design-review.md
'  the list this code follows.
'
'  Byte-identical in both culvert builders (the drift test locks every
'  Rec* procedure); the module-specific value is the Const REC_BUILDER.
'  Recording never fails a build: RecordRun returns "" (off - no report
'  line), "SENT" or "WARN: ...". It never uses the MIDAS request helpers,
'  so the MAPI key cannot reach Google.
' ===========================================================================

' ok             - the build's verdict
' resultsWritten - the results step returned exactly "" (else MIDAS_RESULTS
'                  may still hold an earlier run: no section forces sent)
' report         - the build report so far (its last FAIL line names the
'                  failed step), BEFORE VerdictFirst
' failedStep     - overrides that step name (the SAP builder's crash stage)
Private Function RecordRun(ByVal ok As Boolean, ByVal resultsWritten As Boolean, _
                           ByVal report As String, ByVal failedStep As String, _
                           ByVal runId As String, ByVal startedAt As String, _
                           ByVal t0 As Single) As String

    Dim projectId As String, apiKey As String
    Dim cfg As String, body As String, notes As String, res As String
    Dim finishedAt As String
    Dim duration As Double

    On Error GoTo Crashed

    If Not REC_ENABLED And Not REC_DUMP_BODY Then Exit Function

    finishedAt = RecUtcNow()
    duration = RecSecondsSince(t0)

    ' Both cells blank = recording off, no report line.
    cfg = RecConfig(projectId, apiKey)
    If Not REC_DUMP_BODY Then
        If cfg = "OFF" Then Exit Function
        If Len(cfg) > 0 Then
            RecordRun = cfg
            Exit Function
        End If
    End If

    Application.StatusBar = "Run record: collecting ..."
    body = RecBuildBody(ok, resultsWritten, report, failedStep, runId, startedAt, _
                        finishedAt, duration, notes)

    If REC_DUMP_BODY Then
        Call RecWriteText(Environ$("TEMP") & "\" & SCRIPT_ID & "_record_body.json", body)
    End If

    If Not REC_ENABLED Or cfg = "OFF" Then
        Application.StatusBar = False
        Exit Function
    End If
    If Len(cfg) > 0 Then
        Application.StatusBar = False
        RecordRun = cfg
        Exit Function
    End If

    Application.StatusBar = "Run record: sending ..."
    res = RecSend(projectId, apiKey, runId, body)
    Application.StatusBar = False

    If Len(res) > 0 Then
        RecordRun = "WARN: not sent - " & res
    ElseIf Len(notes) > 0 Then
        RecordRun = "WARN: sent, but" & notes
    Else
        RecordRun = "SENT"
    End If
    Exit Function

Crashed:
    RecordRun = "WARN: not sent - VBA error " & Err.Number & " (" & Err.Description & ")."
    On Error Resume Next
    Application.StatusBar = False

End Function

' Public test hook: writes this workbook's record body to
' %TEMP%\<SCRIPT_ID>_record_body.json with a fixed run id and timestamps,
' as if the build had just succeeded - nothing is sent and no model is
' touched. tests/ compare it with scripts/run-database/record_body.py.
Public Sub RecDumpBodyForTest()

    Dim path As String, body As String, notes As String

    Call InitSeismicGate
    Call InitLiveLoadGate
    Call InitMinVerticalGate

    body = RecBuildBody(True, True, "", "", "00000000-0000-4000-8000-000000000000", _
                        "2026-10-01T00:00:00Z", "2026-10-01T00:00:00Z", 0, notes)
    body = Replace(body, REC_UID_PLACEHOLDER, "test-uid")
    path = Environ$("TEMP") & "\" & SCRIPT_ID & "_record_body.json"
    Call RecWriteText(path, body)

    MsgBox "Run record body written to " & path & IIf(Len(notes) > 0, vbCrLf & "Notes:" & notes, ""), _
           vbInformation, SCRIPT_ID

End Sub

' INPUT!J21 (project ID) and J22 (web API key). Returns "" when both are
' usable, "OFF" when both are blank, otherwise a WARN.
Private Function RecConfig(ByRef projectId As String, ByRef apiKey As String) As String

    Dim ws As Worksheet
    Dim blank As Boolean
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("INPUT")
    On Error GoTo 0
    If ws Is Nothing Then
        RecConfig = "OFF"
        Exit Function
    End If

    projectId = Trim$(RecCellText(ws.Range(REC_PROJECT_ID_CELL).Value2, blank))
    apiKey = Trim$(RecCellText(ws.Range(REC_API_KEY_CELL).Value2, blank))

    If Len(projectId) = 0 And Len(apiKey) = 0 Then
        RecConfig = "OFF"
    ElseIf Len(projectId) = 0 Or Len(apiKey) = 0 Then
        RecConfig = "WARN: not sent - only one of INPUT!" & REC_PROJECT_ID_CELL & " (Firebase project ID) " & _
                    "and INPUT!" & REC_API_KEY_CELL & " (web API key) is filled in. Fill in both, " & _
                    "or clear both to switch recording off."
    Else
        If Len(projectId) < 6 Or Len(projectId) > 30 Then RecConfig = "bad"
        For i = 1 To Len(projectId)
            If Not (Mid$(projectId, i, 1) Like "[a-z0-9-]") Then RecConfig = "bad"
        Next i
        If Len(RecConfig) > 0 Then
            RecConfig = "WARN: not sent - INPUT!" & REC_PROJECT_ID_CELL & " is not a Firebase project ID " & _
                        "(6-30 characters: lower-case letters, digits, hyphens)."
        End If
    End If

End Function

' ---------------------------------------------------------------------------
'  The body: Firestore REST typed JSON {"fields": {...}}, ASCII only. The
'  uid is left as REC_UID_PLACEHOLDER (it is known only after sign-in).
'  Every cell is read here, BEFORE the encryption and the requests, which
'  both wait with DoEvents.
' ---------------------------------------------------------------------------
Private Function RecBuildBody(ByVal ok As Boolean, ByVal resultsWritten As Boolean, _
                              ByVal report As String, ByVal failedStep As String, _
                              ByVal runId As String, ByVal startedAt As String, _
                              ByVal finishedAt As String, ByVal duration As Double, _
                              ByRef notes As String) As String

    Dim f As String, failed As String, sections As String, plain As String
    Dim token As String, keyId As String, encErr As String
    Dim fb As Variant, arr As String

    notes = ""

    ' failed_step: the step name only (the line can carry MIDAS's raw body).
    If Not ok Then
        failed = failedStep
        If Len(failed) = 0 Then failed = RecLastFailStep(report)
        If Len(failed) = 0 Then failed = "unknown"
        failed = Left$(failed, 200)
    End If

    sections = ""
    If ok And resultsWritten Then
        sections = RecSections(notes)
    End If

    plain = RecProjectJson()

    ' Everything is read - only now the encryption (it waits with DoEvents).
    encErr = RecEncrypt(plain, token, keyId)
    plain = ""
    If Len(encErr) > 0 Then
        token = ""
        keyId = "none"
        notes = notes & " the project fields were not encrypted (" & encErr & ") and were left out."
    End If

    Call RecAdd(f, "schema_version", RecTInt(1))
    Call RecAdd(f, "run_id", RecTStr(runId))
    Call RecAdd(f, "uid", RecTStr(REC_UID_PLACEHOLDER))
    Call RecAdd(f, "source", RecTStr("manual"))
    Call RecAdd(f, "builder", RecTStr(REC_BUILDER))
    Call RecAdd(f, "script_id", RecTStr(SCRIPT_ID))
    Call RecAdd(f, "script_version", RecTStr(SCRIPT_VERSION))
    Call RecAdd(f, "status", RecTStr(IIf(ok, "ok", "fail")))
    Call RecAdd(f, "failed_step", IIf(ok, RecTNull(), RecTStr(failed)))
    Call RecAdd(f, "started_at", "{""timestampValue"":""" & startedAt & """}")
    Call RecAdd(f, "finished_at", "{""timestampValue"":""" & finishedAt & """}")
    Call RecAdd(f, "duration_s", RecTDouble(Round(duration, 2)))
    Call RecAdd(f, "step_durations_s", RecTMap(""))

    ' The builder's own default-value notes (" Ec (...) ..., used 33000000.").
    arr = ""
    For Each fb In Split(Trim$(INPUT_FALLBACKS), ". ")
        fb = Trim$(fb)
        If Right$(fb, 1) = "." Then fb = Left$(fb, Len(fb) - 1)
        If Len(fb) > 0 Then
            If Len(arr) > 0 Then arr = arr & ","
            arr = arr & RecTStr(CStr(fb))
        End If
    Next fb
    Call RecAdd(f, "fallbacks", "{""arrayValue"":{""values"":[" & arr & "]}}")

    Call RecAdd(f, "sweep_id", RecTNull())
    Call RecAdd(f, "sweep_point", RecTNull())
    Call RecAdd(f, "program_version", RecTNull())
    Call RecAdd(f, "key_id", RecTStr(keyId))
    Call RecAdd(f, "project_enc", RecTStr(token))
    Call RecAdd(f, "inputs", RecTMap(RecInputs()))
    Call RecAdd(f, "loads", RecTMap(RecLoads()))
    Call RecAdd(f, "gates", RecTMap(RecGates()))
    Call RecAdd(f, "sections", RecTMap(sections))

    RecBuildBody = "{""fields"":{" & f & "}}"

End Function

' "FAIL - Load cases: [Error] ..." (the LAST such line) -> "Load cases".
Private Function RecLastFailStep(ByVal report As String) As String

    Dim p As Long, q As Long
    Dim s As String

    p = InStrRev(report, "FAIL - ")
    If p = 0 Then Exit Function
    s = Mid$(report, p + 7)
    q = InStr(s, vbCr)
    If q = 0 Then q = InStr(s, vbLf)
    If q > 0 Then s = Left$(s, q - 1)
    q = InStr(s, ": ")
    If q > 0 Then s = Left$(s, q - 1)
    RecLastFailStep = Trim$(s)

End Function

' ---- plaintext collectors: explicit cell lists, never a project cell ----

' schema_v1.INPUT_CELLS, in order: name|sheet|cell|kind (d = double,
' s = text, m = the DD1-DD3 map of RecSeismicMap).
Private Function RecInputCellList() As String
    Dim s As String
    s = "phi_deg|INPUT|B3|d;gamma_fill_kN_m3|INPUT|B4|d;gamma_concrete_kN_m3|INPUT|B5|d;"
    s = s & "gamma_ballast_kN_m3|INPUT|B6|d;k0|INPUT|B7|d;ka|INPUT|B8|d;kv_kN_m3|INPUT|B9|d;"
    s = s & "t_slab_m|INPUT|B12|d;t_wall_m|INPUT|B13|d;t_found_m|INPUT|B14|d;"
    s = s & "haunch_v_m|INPUT|B15|d;haunch_h_m|INPUT|B16|d;fill_m|INPUT|B19|d;"
    s = s & "ballast_m|INPUT|B20|d;cover_m|INPUT|B21|d;span_m|INPUT|B22|d;height_m|INPUT|B23|d;"
    s = s & "toe_m|INPUT|B24|d;live_load|INPUT|B26|s;fc_MPa|INPUT|B27|d;fy_MPa|INPUT|B28|d;"
    s = s & "ss|INPUT|G|m;s1|INPUT|H|m;soil_class|INPUT|J4|s;seismic_level|1_GIRIS|AE60|s;"
    s = s & "sds_design|1_GIRIS|AE61|d;mesh_divisions|INPUT|K15|d;min_vertical|INPUT|K17|d"
    RecInputCellList = s
End Function

' schema_v1.LOAD_CELLS, in order: name|MIDAS_INPUT cell, all double.
Private Function RecLoadCellList() As String
    Dim s As String
    s = "ev1|B21;ev2|B22;ehs1_t|B23;ehs1_b|B24;eha1_t|B25;eha1_b|B26;ehs2_t|B27;ehs2_b|B28;"
    s = s & "eha2_t|B29;eha2_b|B30;ll1|B31;ll|B32;lss2|B33;eq_t|B34;eq_b|B35;ev2_side|B36;"
    s = s & "ev1_side|B37;lss1|B38;llacc|B39;lsa2|B40;kh|B15;spring_kN_m3|B42;ec_kN_m2|B43;"
    s = s & "haunch_v_used_m|B8;haunch_h_used_m|B9"
    RecLoadCellList = s
End Function

Private Function RecInputs() As String

    Dim item As Variant, p() As String
    Dim f As String, blank As Boolean, txt As String

    For Each item In Split(RecInputCellList(), ";")
        p = Split(item, "|")
        Select Case p(3)
            Case "d"
                Call RecAdd(f, p(0), RecTDouble(RecCell(p(1), p(2))))
            Case "s"
                txt = RecCellText(RecCell(p(1), p(2)), blank)
                Call RecAdd(f, p(0), IIf(blank, RecTNull(), RecTStr(txt)))
            Case "m"
                Call RecAdd(f, p(0), RecTMap(RecSeismicMap(p(1), p(2))))
        End Select
    Next item
    RecInputs = f

End Function

' DD1..DD3 = rows 4..6 of one column (Ss: G, S1: H).
Private Function RecSeismicMap(ByVal sheetName As String, ByVal col As String) As String
    Dim f As String, k As Long
    For k = 1 To 3
        Call RecAdd(f, "DD" & k, RecTDouble(RecCell(sheetName, col & (3 + k))))
    Next k
    RecSeismicMap = f
End Function

Private Function RecLoads() As String
    Dim item As Variant, p() As String, f As String
    For Each item In Split(RecLoadCellList(), ";")
        p = Split(item, "|")
        Call RecAdd(f, p(0), RecTDouble(RecCell("MIDAS_INPUT", p(1))))
    Next item
    RecLoads = f
End Function

' The gates as this run applied them. seismic_source is rebuilt from the
' cells with a locale-free number format (the report's text uses Format$,
' which writes 0,769 on a Turkish Windows).
Private Function RecGates() As String

    Dim f As String, src As String
    Dim pv As Variant, ov As Variant, qv As Variant
    Dim cv As Variant, hv As Variant, tv As Variant, bv As Variant
    Dim ratio As Variant

    ov = RecCell("1_GIRIS", "O81")
    qv = RecCell("1_GIRIS", "Q81")
    cv = RecCell("INPUT", "B21")
    hv = RecCell("INPUT", "B23")
    tv = RecCell("INPUT", "B12")
    bv = RecCell("INPUT", "B14")
    ratio = Empty
    If RecIsNum(cv) And RecIsNum(hv) And RecIsNum(tv) And RecIsNum(bv) Then
        If CDbl(hv) + CDbl(tv) + CDbl(bv) > 0 Then ratio = CDbl(cv) / (CDbl(hv) + CDbl(tv) + CDbl(bv))
    End If

    src = SEISMIC_GATE_SOURCE
    If SEISMIC_GATE_OVERRIDE = -1 Then
        pv = RecCell("1_GIRIS", "P81")
        If VarType(pv) = vbString Then pv = Trim$(pv) Else pv = ""
        If pv = "<" Or pv = ">" Then
            src = "1_GIRIS!P81 = """ & pv & """"
        ElseIf RecIsNum(ov) And RecIsNum(qv) Then
            src = "1_GIRIS!O81 = " & RecFixed3(CDbl(ov)) & IIf(CDbl(ov) < CDbl(qv), " < ", " >= ") & _
                  RecFixed3(CDbl(qv))
        ElseIf Not IsEmpty(ratio) Then
            src = "Z/H = INPUT!B21/(B23+B12+B14) = " & RecFixed3(CDbl(ratio)) & _
                  IIf(ratio < 0.5, " < ", " >= ") & "0.500"
        End If
    End If

    Call RecAdd(f, "seismic_active", RecTBool(SEISMIC_ACTIVE))
    Call RecAdd(f, "seismic_source", RecTStr(src))
    If RecIsNum(ov) Then
        Call RecAdd(f, "zh_ratio", RecTDouble(ov))
    Else
        Call RecAdd(f, "zh_ratio", RecTDouble(ratio))
    End If
    Call RecAdd(f, "live_load_active", RecTBool(LIVE_LOAD_ACTIVE))
    Call RecAdd(f, "min_vertical_active", RecTBool(MIN_VERTICAL_ACTIVE))
    RecGates = f

End Function

' ---- the section forces ----

' KESIT|elem|part|face|mode|V elem|V part|As,req cell|phi Vc cell -
' analysis/governing.py's SECTION_POSITIONS / SHEAR_POSITIONS and
' schema_v1's 6_DONATI cells. elem/part are the fallback when the summary
' block's typed position is blank or unreadable; no V position = no shear.
Private Function RecSectionList() As String
    Dim s As String
    s = "1-1|19|J|UST|min|19|M|M9|J64;2-2|19|I|UST|min|19|I|M10|J65;3-3|23|I|ALT|max|||M11|;"
    s = s & "4-4|4|I|ALT|max|4|M|M12|J66;5-5|39|I|UST|min|||M13|;6-6|30|J|DIS|min|30|M|M14|J67;"
    s = s & "7-7|27|I|DIS|min|||M15|;8-8|26|I|IC|max|||M16|"
    RecSectionList = s
End Function

' Per face, over MIDAS_RESULTS table 1 (B:J from row 3): the largest
' moment of the face's tension sign over ULS-* + EQ-1 (m_str) / SLS-*
' (m_ser) at the summary block's position, Nu from the same row; V = the
' largest |Shear-z| over ULS-* + EQ-1. Strict < / > - the first row wins a
' tie. "" (an empty map) when table 1 does not hold exactly this run's
' combinations. Returns the map's fields.
Private Function RecSections(ByRef notes As String) As String

    Dim ws As Worksheet, wsD As Worksheet
    Dim data As Variant, sm As Variant
    Dim lastRow As Long, i As Long, n As Long, skipped As Long
    Dim rElem() As Long, rComb() As String, rPos() As String
    Dim rAx() As Double, rVz() As Double, rMy() As Double
    Dim have As Object, want As Object, k As Variant
    Dim pos As String, comb As String
    Dim item As Variant, p() As String, r As Long
    Dim f As String, sec As String, sh As String, calcDone As Boolean
    Dim mStr As String, mSer As String, vv As String
    Dim eM As Long, pM As String, eS As Long, pS As String, eV As Long, pV As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("MIDAS_RESULTS")
    Set wsD = ThisWorkbook.Worksheets("6_DONATI")
    On Error GoTo 0
    If ws Is Nothing Then
        notes = notes & " no section forces (no MIDAS_RESULTS sheet)."
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 3 Then
        notes = notes & " no section forces (MIDAS_RESULTS table 1 is empty)."
        Exit Function
    End If
    data = ws.Range("B3:J" & lastRow).Value2
    sm = ws.Range("Q3:AI10").Value2

    n = 0
    ReDim rElem(1 To UBound(data, 1)), rComb(1 To UBound(data, 1)), rPos(1 To UBound(data, 1))
    ReDim rAx(1 To UBound(data, 1)), rVz(1 To UBound(data, 1)), rMy(1 To UBound(data, 1))
    For i = 1 To UBound(data, 1)
        If IsEmpty(data(i, 1)) Then Exit For
        pos = ""
        If VarType(data(i, 3)) = vbString Then
            If Left$(Trim$(data(i, 3)), 1) = "I" Then pos = "I"
            If Left$(Trim$(data(i, 3)), 1) = "J" Then pos = "J"
            If Trim$(data(i, 3)) = "2/4" Then pos = "M"
        End If
        If Len(pos) > 0 And RecIsWhole(data(i, 1)) And VarType(data(i, 2)) = vbString And _
           RecIsNum(data(i, 4)) And RecIsNum(data(i, 5)) And RecIsNum(data(i, 6)) And _
           RecIsNum(data(i, 7)) And RecIsNum(data(i, 8)) And RecIsNum(data(i, 9)) Then
            n = n + 1
            rElem(n) = CLng(data(i, 1))
            rComb(n) = Trim$(data(i, 2))
            rPos(n) = pos
            rAx(n) = CDbl(data(i, 4))
            rVz(n) = CDbl(data(i, 6))
            rMy(n) = CDbl(data(i, 8))
        Else
            skipped = skipped + 1
        End If
    Next i

    ' Table 1 must hold exactly this run's SLS-/ULS-/EQ-1 combinations.
    Set have = CreateObject("Scripting.Dictionary")
    Set want = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        have(rComb(i)) = True
    Next i
    For i = 1 To 14
        want("SLS-" & i) = True
    Next i
    For i = 1 To IIf(MIN_VERTICAL_ACTIVE, 20, 14)
        want("ULS-" & i) = True
    Next i
    If SEISMIC_ACTIVE Then want("EQ-1") = True
    comb = ""
    If have.Count <> want.Count Then comb = "x"
    For Each k In want.Keys
        If Not have.Exists(k) Then comb = "x"
    Next k
    If Len(comb) > 0 Then
        notes = notes & " no section forces (MIDAS_RESULTS table 1 holds other combinations than this run's)."
        Exit Function
    End If

    ' The sheet's own values only from a finished recalculation.
    On Error Resume Next
    Application.Calculate
    calcDone = (Application.CalculationState = xlDone)
    On Error GoTo 0

    r = 0
    For Each item In Split(RecSectionList(), ";")
        p = Split(item, "|")
        r = r + 1
        sec = p(0)

        ' Positions typed in the summary block (row r): R/T, X/Z, AD/AF.
        eM = RecSummaryElem(sm(r, 2), CLng(p(1))): pM = RecSummaryPart(sm(r, 4), p(2))
        eS = RecSummaryElem(sm(r, 8), CLng(p(1))): pS = RecSummaryPart(sm(r, 10), p(2))
        If Len(p(5)) > 0 Then
            eV = RecSummaryElem(sm(r, 14), CLng(p(5))): pV = RecSummaryPart(sm(r, 16), p(6))
        End If

        mStr = RecGoverning(rElem, rComb, rPos, rMy, rAx, n, eM, pM, "STR", p(4), "value_kNm")
        mSer = RecGoverning(rElem, rComb, rPos, rMy, rAx, n, eS, pS, "SER", p(4), "value_kNm")

        sh = ""
        If calcDone Then
            Call RecAdd(sh, "m_str_value", RecTDouble(sm(r, 6)))
            Call RecAdd(sh, "m_str_comb", RecTTrimText(sm(r, 3)))
            Call RecAdd(sh, "m_ser_value", RecTDouble(sm(r, 12)))
            Call RecAdd(sh, "m_ser_comb", RecTTrimText(sm(r, 9)))
            If Len(p(5)) > 0 Then
                Call RecAdd(sh, "v_value", RecTDouble(sm(r, 18)))
                Call RecAdd(sh, "v_comb", RecTTrimText(sm(r, 15)))
            Else
                Call RecAdd(sh, "v_value", RecTNull())
                Call RecAdd(sh, "v_comb", RecTNull())
            End If
            If wsD Is Nothing Then
                Call RecAdd(sh, "as_req_mm2", RecTNull())
                Call RecAdd(sh, "phi_vc_kN", RecTNull())
            Else
                Call RecAdd(sh, "as_req_mm2", RecTDouble(wsD.Range(p(7)).Value2))
                If Len(p(8)) > 0 Then
                    Call RecAdd(sh, "phi_vc_kN", RecTDouble(wsD.Range(p(8)).Value2))
                Else
                    Call RecAdd(sh, "phi_vc_kN", RecTNull())
                End If
            End If
        Else
            Call RecAdd(sh, "m_str_value", RecTNull())
            Call RecAdd(sh, "m_str_comb", RecTNull())
            Call RecAdd(sh, "m_ser_value", RecTNull())
            Call RecAdd(sh, "m_ser_comb", RecTNull())
            Call RecAdd(sh, "v_value", RecTNull())
            Call RecAdd(sh, "v_comb", RecTNull())
            Call RecAdd(sh, "as_req_mm2", RecTNull())
            Call RecAdd(sh, "phi_vc_kN", RecTNull())
        End If

        vv = ""
        Call RecAdd(vv, "face", RecTStr(p(3)))
        Call RecAdd(vv, "element", RecTInt(eM))
        Call RecAdd(vv, "part", RecTStr(pM))
        Call RecAdd(vv, "m_str", mStr)
        Call RecAdd(vv, "m_ser", mSer)
        If Len(p(5)) > 0 Then
            Call RecAdd(vv, "v", RecGoverning(rElem, rComb, rPos, rVz, rAx, n, eV, pV, "STR", "abs", "value_kN"))
        End If
        Call RecAdd(vv, "sheet", RecTMap(sh))
        Call RecAdd(f, sec, RecTMap(vv))
    Next item

    If Not calcDone Then notes = notes & " the sheet's own section values were left out (recalculation not finished)."
    RecSections = f

End Function

' The governing row of one section position: limitState "STR" (ULS-* and
' EQ-1) or "SER" (SLS-*); mode "min" / "max" (signed) or "abs". A typed
' map, or null when no row matches.
Private Function RecGoverning(ByRef rElem() As Long, ByRef rComb() As String, ByRef rPos() As String, _
                              ByRef rVal() As Double, ByRef rAx() As Double, ByVal n As Long, _
                              ByVal elemNo As Long, ByVal pos As String, ByVal limitState As String, _
                              ByVal mode As String, ByVal valueKey As String) As String

    Dim i As Long, best As Long
    Dim keep As Boolean, better As Boolean
    Dim f As String

    best = 0
    For i = 1 To n
        If rElem(i) = elemNo And rPos(i) = pos Then
            If limitState = "STR" Then
                keep = (Left$(rComb(i), 4) = "ULS-" Or rComb(i) = "EQ-1")
            Else
                keep = (Left$(rComb(i), 4) = "SLS-")
            End If
            If keep Then
                If best = 0 Then
                    better = True
                ElseIf mode = "min" Then
                    better = (rVal(i) < rVal(best))
                ElseIf mode = "max" Then
                    better = (rVal(i) > rVal(best))
                Else
                    better = (Abs(rVal(i)) > Abs(rVal(best)))
                End If
                If better Then best = i
            End If
        End If
    Next i

    If best = 0 Then
        RecGoverning = RecTNull()
        Exit Function
    End If
    Call RecAdd(f, valueKey, RecTDouble(rVal(best)))
    Call RecAdd(f, "comb", RecTStr(rComb(best)))
    Call RecAdd(f, "element", RecTInt(rElem(best)))
    Call RecAdd(f, "part", RecTStr(rPos(best)))
    Call RecAdd(f, "nu_kN", RecTDouble(rAx(best)))
    RecGoverning = RecTMap(f)

End Function

' A summary-block element cell: a whole number, else the fallback.
Private Function RecSummaryElem(ByVal v As Variant, ByVal fallback As Long) As Long
    If RecIsWhole(v) Then RecSummaryElem = CLng(v) Else RecSummaryElem = fallback
End Function

' A summary-block part cell: I... / J... / 2/4 / M, or the number 0.5 (the
' sheet types 0.5 for the middle) - VarType, never CStr (0,5 in Turkish).
Private Function RecSummaryPart(ByVal v As Variant, ByVal fallback As String) As String
    RecSummaryPart = fallback
    If RecIsNum(v) Then
        If Abs(CDbl(v) - 0.5) < 0.000000000001 Then RecSummaryPart = "M"
    ElseIf VarType(v) = vbString Then
        If Left$(Trim$(v), 1) = "I" Then
            RecSummaryPart = "I"
        ElseIf Left$(Trim$(v), 1) = "J" Then
            RecSummaryPart = "J"
        ElseIf Trim$(v) = "2/4" Or Trim$(v) = "M" Then
            RecSummaryPart = "M"
        End If
    End If
End Function

' ---- the project fields (the ONLY collector of project cells) ----

' One JSON object (raw UTF-8 text, it never leaves the computer
' unencrypted) - schema_v1.PROJECT_FIELDS, in order.
Private Function RecProjectJson() As String

    Dim item As Variant, p() As String
    Dim s As String, txt As String, blank As Boolean
    Dim sh As Object

    For Each item In Split("project_type|G9;project_number|G8;project_part|G11;km|G12;" & _
                           "title|G13;engineer|N15;revision|N16", ";")
        p = Split(item, "|")
        txt = RecCellText(RecCell("INPUT", p(1)), blank)
        If Len(s) > 0 Then s = s & ","
        s = s & """" & p(0) & """:" & IIf(blank, "null", """" & RecJsonText(txt, False) & """")
    Next item

    ' WScript.Shell reads the environment as Unicode (Environ$ is ANSI).
    Set sh = CreateObject("WScript.Shell")
    s = s & ",""workbook"":""" & RecJsonText(ThisWorkbook.Name, False) & """"
    s = s & ",""windows_user"":""" & RecJsonText(sh.ExpandEnvironmentStrings("%USERNAME%"), False) & """"
    s = s & ",""computer"":""" & RecJsonText(sh.ExpandEnvironmentStrings("%COMPUTERNAME%"), False) & """"
    RecProjectJson = "{" & s & "}"

End Function

' Locks the project fields as a "CF1" token (scripts/run-database/
' encrypt_fields.ps1, run by PowerShell/.NET). The plaintext goes through a
' UTF-8 temp file that the script deletes (and this code again), never the
' command line. Returns "" or the reason; on any failure nothing is sent in
' its place - never the plaintext.
Private Function RecEncrypt(ByVal plain As String, ByRef token As String, ByRef keyId As String) As String

    Dim fso As Object, wmi As Object, startup As Object
    Dim base As String, inPath As String, outPath As String, ps1Path As String
    Dim psExe As String, cmd As String, outText As String
    Dim pid As Variant, rc As Long
    Dim t0 As Single, tw As Single

    token = ""
    keyId = "none"
    If Len(REC_PUBLIC_KEY_XML) = 0 Or Len(REC_KEY_ID) <> 16 Then
        RecEncrypt = "no public key in REC_PUBLIC_KEY_XML"
        Exit Function
    End If

    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    base = Environ$("TEMP") & "\" & SCRIPT_ID & "_rec_" & Replace(RecNewRunId(), "-", "")
    inPath = base & "_in.txt"
    outPath = base & "_out.txt"
    ps1Path = base & ".ps1"

    Call RecWriteText(inPath, plain)
    Call RecWriteText(ps1Path, Replace(Replace(RecEncryptScript(), "__PUBLIC_KEY_XML__", _
                      REC_PUBLIC_KEY_XML), "__KEY_ID__", REC_KEY_ID))

    psExe = Environ$("WINDIR") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
    cmd = """" & psExe & """ -NoProfile -NonInteractive -ExecutionPolicy Bypass -File """ & _
          ps1Path & """ """ & inPath & """ """ & outPath & """"
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set startup = wmi.Get("Win32_ProcessStartup").SpawnInstance_
    startup.ShowWindow = 0
    rc = wmi.Get("Win32_Process").Create(cmd, Null, startup, pid)
    If rc <> 0 Then
        RecEncrypt = "PowerShell did not start, WMI code " & rc
        GoTo Cleanup
    End If

    ' Wait for OUR process only; on a timeout end it and nothing else.
    t0 = Timer
    Do
        If wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE ProcessId = " & pid).Count = 0 Then Exit Do
        If RecSecondsSince(t0) > REC_ENCRYPT_TIMEOUT_SEC Then
            Call RecEndProcess(wmi, pid)
            RecEncrypt = "PowerShell did not finish within " & REC_ENCRYPT_TIMEOUT_SEC & " s"
            GoTo Cleanup
        End If
        tw = Timer
        Do While RecSecondsSince(tw) < 0.2
            DoEvents
        Loop
    Loop

    outText = ""
    If fso.FileExists(outPath) Then outText = Trim$(RecReadText(outPath))
    outText = Replace(Replace(outText, vbCr, ""), vbLf, "")
    If Left$(outText, 7) = "OK|CF1:" And Len(outText) - 3 <= 4000 Then
        token = Mid$(outText, 4)
        keyId = REC_KEY_ID
    ElseIf Left$(outText, 5) = "FAIL|" Then
        RecEncrypt = "PowerShell: " & Left$(Mid$(outText, 6), 200)
    Else
        RecEncrypt = "PowerShell gave no usable answer"
    End If

Cleanup:
    On Error Resume Next
    If fso.FileExists(inPath) Then fso.DeleteFile inPath, True
    If fso.FileExists(outPath) Then fso.DeleteFile outPath, True
    If fso.FileExists(ps1Path) Then fso.DeleteFile ps1Path, True
    Exit Function

Failed:
    RecEncrypt = "VBA error " & Err.Number & " (" & Err.Description & ")"
    token = ""
    keyId = "none"
    Resume Cleanup

End Function

Private Sub RecEndProcess(ByVal wmi As Object, ByVal pid As Variant)
    Dim proc As Object
    On Error Resume Next
    For Each proc In wmi.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = " & pid)
        proc.Terminate
    Next proc
End Sub

' The PowerShell that locks the project fields: scripts/run-database/
' encrypt_fields.ps1 from its param( line on, generated from that file
' (tests/verify_record_run.py compares them). The two placeholders are
' replaced by RecEncrypt.
Private Function RecEncryptScript() As String

    Dim s As String

    s = s & "param(" & vbCrLf
    s = s & "    [Parameter(Mandatory = $true, Position = 0)][string]$InFile," & vbCrLf
    s = s & "    [Parameter(Mandatory = $true, Position = 1)][string]$OutFile" & vbCrLf
    s = s & ")" & vbCrLf
    s = s & "$ErrorActionPreference = 'Stop'" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "$PublicKeyXml = '__PUBLIC_KEY_XML__'" & vbCrLf
    s = s & "$KeyId = '__KEY_ID__'" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "function Write-Result([string]$line) {" & vbCrLf
    s = s & "    [IO.File]::WriteAllText($OutFile, $line, [Text.Encoding]::ASCII)" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "$rsa = $null; $aes = $null; $enc = $null; $hmac = $null; $rng = $null" & vbCrLf
    s = s & "$k = $null" & vbCrLf
    s = s & "try {" & vbCrLf
    s = s & "    if ($PublicKeyXml.StartsWith('__') -or $KeyId.StartsWith('__')) {" & vbCrLf
    s = s & "        throw 'public key placeholders were not replaced'" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "    if ($KeyId -notmatch '^[0-9a-f]{16}$') { throw ""bad key id '$KeyId'"" }" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    # ReadAllText with UTF8 drops a BOM; re-encode WITHOUT one." & vbCrLf
    s = s & "    $text = [IO.File]::ReadAllText($InFile, [Text.Encoding]::UTF8)" & vbCrLf
    s = s & "    $plain = (New-Object Text.UTF8Encoding($false)).GetBytes($text)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    $rng = New-Object Security.Cryptography.RNGCryptoServiceProvider" & vbCrLf
    s = s & "    $k = New-Object byte[] 64" & vbCrLf
    s = s & "    $rng.GetBytes($k)" & vbCrLf
    s = s & "    $iv = New-Object byte[] 16" & vbCrLf
    s = s & "    $rng.GetBytes($iv)" & vbCrLf
    s = s & "    $kEnc = New-Object byte[] 32" & vbCrLf
    s = s & "    $kMac = New-Object byte[] 32" & vbCrLf
    s = s & "    [Array]::Copy($k, 0, $kEnc, 0, 32)" & vbCrLf
    s = s & "    [Array]::Copy($k, 32, $kMac, 0, 32)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    $rsa = New-Object Security.Cryptography.RSACng" & vbCrLf
    s = s & "    $rsa.FromXmlString($PublicKeyXml)" & vbCrLf
    s = s & "    if ($rsa.KeySize -lt 3072) { throw ""public key is only $($rsa.KeySize) bits"" }" & vbCrLf
    s = s & "    $w = $rsa.Encrypt($k, [Security.Cryptography.RSAEncryptionPadding]::OaepSHA256)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    $aes = [Security.Cryptography.Aes]::Create()" & vbCrLf
    s = s & "    $aes.Mode = [Security.Cryptography.CipherMode]::CBC" & vbCrLf
    s = s & "    $aes.Padding = [Security.Cryptography.PaddingMode]::PKCS7" & vbCrLf
    s = s & "    $aes.KeySize = 256" & vbCrLf
    s = s & "    $aes.Key = $kEnc" & vbCrLf
    s = s & "    $aes.IV = $iv" & vbCrLf
    s = s & "    $enc = $aes.CreateEncryptor()" & vbCrLf
    s = s & "    $c = $enc.TransformFinalBlock($plain, 0, $plain.Length)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    $body = 'CF1:' + $KeyId + ':' + [Convert]::ToBase64String($w) + ':' +" & vbCrLf
    s = s & "        [Convert]::ToBase64String($iv) + ':' + [Convert]::ToBase64String($c)" & vbCrLf
    s = s & "    $hmac = New-Object Security.Cryptography.HMACSHA256 (, $kMac)" & vbCrLf
    s = s & "    $t = $hmac.ComputeHash([Text.Encoding]::ASCII.GetBytes($body))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Write-Result ('OK|' + $body + ':' + [Convert]::ToBase64String($t))" & vbCrLf
    s = s & "    $code = 0" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "catch {" & vbCrLf
    s = s & "    $msg = ($_.Exception.Message -replace '[\r\n|]+', ' ').Trim()" & vbCrLf
    s = s & "    try { Write-Result ('FAIL|' + $msg) } catch { }" & vbCrLf
    s = s & "    $code = 1" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "finally {" & vbCrLf
    s = s & "    if ($k) { [Array]::Clear($k, 0, $k.Length) }" & vbCrLf
    s = s & "    if ($kEnc) { [Array]::Clear($kEnc, 0, $kEnc.Length) }" & vbCrLf
    s = s & "    if ($kMac) { [Array]::Clear($kMac, 0, $kMac.Length) }" & vbCrLf
    s = s & "    foreach ($o in @($enc, $aes, $hmac, $rsa, $rng)) { if ($o) { $o.Dispose() } }" & vbCrLf
    s = s & "    # Never Remove-Item: it fails on the 8.3 %TEMP% path (CLAUDE.md)." & vbCrLf
    s = s & "    try { if ([IO.File]::Exists($InFile)) { [IO.File]::Delete($InFile) } } catch { }" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "exit $code" & vbCrLf
    RecEncryptScript = s

End Function

' ---- sending ----

' Signs in (anonymous; cached token, else the refresh token kept in
' %APPDATA%\CulvertDB, else a new sign-up) and creates runs/<runId>.
' Returns "" when Firestore confirmed the document, otherwise the reason.
' At most REC_MAX_REQUESTS requests; tokens, the API key and the URLs never
' go into the reason.
Private Function RecSend(ByVal projectId As String, ByVal apiKey As String, _
                         ByVal runId As String, ByVal body As String) As String

    Dim idToken As String, uid As String, res As String
    Dim url As String, docPath As String, resp As String
    Dim status As Long, nReq As Long

    res = RecSignIn(projectId, apiKey, False, idToken, uid, nReq)
    If Len(res) > 0 Then
        RecSend = res
        Exit Function
    End If

    url = "https://firestore.googleapis.com/v1/projects/" & projectId & _
          "/databases/(default)/documents/runs?documentId=" & runId
    docPath = """projects/" & projectId & "/databases/(default)/documents/runs/" & runId & """"

    Call RecPost(url, "application/json", Replace(body, REC_UID_PLACEHOLDER, uid), idToken, status, resp, nReq)
    If status = 200 And InStr(resp, docPath) > 0 Then Exit Function

    If status = 0 Then
        ' The first attempt may have landed before the connection broke:
        ' only on this retry does 409 (already exists) count as sent.
        Call RecPost(url, "application/json", Replace(body, REC_UID_PLACEHOLDER, uid), idToken, status, resp, nReq)
        If status = 200 And InStr(resp, docPath) > 0 Then Exit Function
        If status = 409 Then Exit Function
    ElseIf status = 401 Then
        res = RecSignIn(projectId, apiKey, True, idToken, uid, nReq)
        If Len(res) > 0 Then
            RecSend = res
            Exit Function
        End If
        Call RecPost(url, "application/json", Replace(body, REC_UID_PLACEHOLDER, uid), idToken, status, resp, nReq)
        If status = 200 And InStr(resp, docPath) > 0 Then Exit Function
    End If

    RecSend = RecHttpProblem("the database", status, resp)

End Function

' idToken + uid for this project, from the cache, the refresh token or a
' new anonymous sign-up. forceRefresh skips the cache (after a 401).
Private Function RecSignIn(ByVal projectId As String, ByVal apiKey As String, _
                           ByVal forceRefresh As Boolean, ByRef idToken As String, _
                           ByRef uid As String, ByRef nReq As Long) As String

    Dim cacheFor As String, refreshTok As String, newRefresh As String
    Dim resp As String, msg As String, expires As String
    Dim status As Long

    cacheFor = projectId & "|" & apiKey
    If Not forceRefresh And REC_TOKEN_FOR = cacheFor And Len(REC_TOKEN_CACHE) > 0 And _
       Now < REC_TOKEN_EXPIRES Then
        idToken = REC_TOKEN_CACHE
        uid = REC_TOKEN_UID
        Exit Function
    End If
    REC_TOKEN_CACHE = ""
    REC_TOKEN_UID = ""
    REC_TOKEN_FOR = ""

    refreshTok = RecReadToken(projectId)
    If Len(refreshTok) > 0 Then
        Call RecPost("https://securetoken.googleapis.com/v1/token?key=" & apiKey, _
                     "application/x-www-form-urlencoded", _
                     "grant_type=refresh_token&refresh_token=" & RecUrlEncode(refreshTok), "", _
                     status, resp, nReq)
        If status = 200 Then
            idToken = RecJsonValue(resp, "id_token")
            newRefresh = RecJsonValue(resp, "refresh_token")
            expires = RecJsonValue(resp, "expires_in")
            uid = RecJsonValue(resp, "user_id")
        Else
            msg = RecJsonValue(resp, "message")
            If status <> 400 Or Not (Left$(msg, 21) = "INVALID_REFRESH_TOKEN" Or _
               Left$(msg, 14) = "USER_NOT_FOUND" Or Left$(msg, 13) = "TOKEN_EXPIRED") Then
                RecSignIn = RecHttpProblem("sign-in", status, resp)
                Exit Function
            End If
        End If
    End If

    If Len(idToken) = 0 Then
        Call RecPost("https://identitytoolkit.googleapis.com/v1/accounts:signUp?key=" & apiKey, _
                     "application/json", "{""returnSecureToken"":true}", "", status, resp, nReq)
        If status <> 200 Then
            RecSignIn = RecHttpProblem("anonymous sign-in", status, resp)
            Exit Function
        End If
        idToken = RecJsonValue(resp, "idToken")
        newRefresh = RecJsonValue(resp, "refreshToken")
        expires = RecJsonValue(resp, "expiresIn")
        uid = RecJsonValue(resp, "localId")
    End If

    If Len(idToken) = 0 Or Len(uid) = 0 Then
        RecSignIn = "sign-in answered without a token."
        Exit Function
    End If
    If Len(newRefresh) > 0 And newRefresh <> refreshTok Then Call RecSaveToken(projectId, newRefresh)

    REC_TOKEN_CACHE = idToken
    REC_TOKEN_UID = uid
    REC_TOKEN_FOR = cacheFor
    REC_TOKEN_EXPIRES = Now + (IIf(Val(expires) > 600, Val(expires), 600) - 300) / 86400#

End Function

' One POST with its own WinHTTP client (never the MIDAS one, which carries
' the MAPI key). status 0 = no answer (resp = the reason), -1 = the request
' cap was reached.
Private Sub RecPost(ByVal url As String, ByVal contentType As String, ByVal body As String, _
                    ByVal bearer As String, ByRef status As Long, ByRef resp As String, _
                    ByRef nReq As Long)

    Dim http As Object

    status = 0
    resp = ""
    If nReq >= REC_MAX_REQUESTS Then
        status = -1
        resp = "request limit reached"
        Exit Sub
    End If
    nReq = nReq + 1

    On Error GoTo NoAnswer
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 5000, 5000, 10000, 10000
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", contentType
    If Len(bearer) > 0 Then http.SetRequestHeader "Authorization", "Bearer " & bearer
    http.Send body
    status = http.Status
    resp = http.ResponseText
    Exit Sub

NoAnswer:
    status = 0
    resp = "no answer: " & Err.Description

End Sub

' A short reason for the report - Google's error message, never the URL.
Private Function RecHttpProblem(ByVal what As String, ByVal status As Long, ByVal resp As String) As String

    Dim msg As String

    If status <= 0 Then
        RecHttpProblem = what & ": " & Left$(Trim$(Replace(Replace(resp, vbCr, " "), vbLf, " ")), 150)
        If status = 0 Then RecHttpProblem = RecHttpProblem & " (offline?)"
        Exit Function
    End If
    msg = Left$(RecJsonValue(resp, "message"), 150)
    RecHttpProblem = what & " answered HTTP " & status & IIf(Len(msg) > 0, " - " & msg, "")
    Select Case status
        Case 400
            If InStr(msg, "API key") > 0 Then RecHttpProblem = RecHttpProblem & " (check INPUT!" & REC_API_KEY_CELL & ")"
            If InStr(msg, "OPERATION_NOT_ALLOWED") > 0 Then RecHttpProblem = RecHttpProblem & _
                " (turn on anonymous sign-in in the Firebase console)"
        Case 403
            RecHttpProblem = RecHttpProblem & " (refused - the security rules, or a restricted API key)"
        Case 404
            RecHttpProblem = RecHttpProblem & " (check INPUT!" & REC_PROJECT_ID_CELL & " and that the database exists)"
        Case 409
            RecHttpProblem = RecHttpProblem & " (run id already exists)"
    End Select

End Function

' The refresh token of this computer, %APPDATA%\CulvertDB\<project>.token
' (FSO / ADODB - VBA's Open cannot take a Turkish user name).
Private Function RecTokenPath(ByVal projectId As String) As String
    RecTokenPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%APPDATA%") & _
                   "\CulvertDB\" & projectId & ".token"
End Function

Private Function RecReadToken(ByVal projectId As String) As String
    Dim path As String
    On Error GoTo NoToken
    path = RecTokenPath(projectId)
    If CreateObject("Scripting.FileSystemObject").FileExists(path) Then
        RecReadToken = Trim$(Replace(Replace(Replace(RecReadText(path), ChrW(65279), ""), vbCr, ""), vbLf, ""))
    End If
    Exit Function
NoToken:
    RecReadToken = ""
End Function

Private Sub RecSaveToken(ByVal projectId As String, ByVal tok As String)
    Dim fso As Object, folder As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.GetParentFolderName(RecTokenPath(projectId))
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    Call RecWriteText(RecTokenPath(projectId), tok)
End Sub

' ---- small helpers ----

' A cell's .Value2, or Empty when the sheet or cell cannot be read.
Private Function RecCell(ByVal sheetName As String, ByVal addr As String) As Variant
    On Error GoTo Unreadable
    RecCell = ThisWorkbook.Worksheets(sheetName).Range(addr).Value2
    Exit Function
Unreadable:
    RecCell = Empty
End Function

' Text of a cell value as record_body.cell_text: 140 -> "140", 0.5 ->
' "0.5" (locale-free), blank / error -> blank.
Private Function RecCellText(ByVal v As Variant, ByRef blank As Boolean) As String
    blank = False
    If IsEmpty(v) Or IsError(v) Or IsNull(v) Then
        blank = True
    ElseIf VarType(v) = vbBoolean Then
        RecCellText = IIf(v, "True", "False")
    ElseIf RecIsWhole(v) Then
        RecCellText = Format$(v, "0")
    ElseIf RecIsNum(v) Then
        RecCellText = JsonNum(CDbl(v))
    Else
        RecCellText = CStr(v)
    End If
End Function

Private Function RecIsNum(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal
            RecIsNum = True
    End Select
End Function

Private Function RecIsWhole(ByVal v As Variant) As Boolean
    If RecIsNum(v) Then
        If Abs(CDbl(v)) < 1000000000# Then RecIsWhole = (CDbl(v) = Fix(CDbl(v)))
    End If
End Function

' 0.7692 -> "0.769" whatever the Windows number format.
Private Function RecFixed3(ByVal x As Double) As String
    Dim n As Double, s As String
    n = Round(Abs(x) * 1000)
    s = Format$(n, "0")
    Do While Len(s) < 4
        s = "0" & s
    Loop
    RecFixed3 = IIf(x < 0 And n > 0, "-", "") & Left$(s, Len(s) - 3) & "." & Right$(s, 3)
End Function

Private Sub RecAdd(ByRef acc As String, ByVal fieldName As String, ByVal typed As String)
    If Len(acc) > 0 Then acc = acc & ","
    acc = acc & """" & fieldName & """:" & typed
End Sub

Private Function RecTNull() As String
    RecTNull = "{""nullValue"":null}"
End Function

Private Function RecTInt(ByVal n As Long) As String
    RecTInt = "{""integerValue"":""" & CStr(n) & """}"
End Function

' doubleValue for a number (whole numbers too), null for anything else.
Private Function RecTDouble(ByVal v As Variant) As String
    If RecIsNum(v) Then
        RecTDouble = "{""doubleValue"":" & JsonNum(CDbl(v)) & "}"
    Else
        RecTDouble = RecTNull()
    End If
End Function

Private Function RecTStr(ByVal s As String) As String
    RecTStr = "{""stringValue"":""" & RecJsonText(s, True) & """}"
End Function

Private Function RecTTrimText(ByVal v As Variant) As String
    If VarType(v) = vbString Then RecTTrimText = RecTStr(Trim$(v)) Else RecTTrimText = RecTNull()
End Function

Private Function RecTBool(ByVal b As Boolean) As String
    RecTBool = "{""booleanValue"":" & IIf(b, "true", "false") & "}"
End Function

Private Function RecTMap(ByVal fields As String) As String
    If Len(fields) = 0 Then
        RecTMap = "{""mapValue"":{}}"
    Else
        RecTMap = "{""mapValue"":{""fields"":{" & fields & "}}}"
    End If
End Function

' JSON string contents. asciiOnly: every character above 126 as \uXXXX
' (the request body); otherwise non-ASCII stays as is (the UTF-8 file
' that is encrypted).
Private Function RecJsonText(ByVal s As String, ByVal asciiOnly As Boolean) As String
    Dim i As Long, c As Long, out As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        If c = 34 Then
            out = out & "\"""
        ElseIf c = 92 Then
            out = out & "\\"
        ElseIf c < 32 Or (asciiOnly And c > 126) Then
            out = out & "\u" & Right$("000" & LCase$(Hex$(c)), 4)
        Else
            out = out & Mid$(s, i, 1)
        End If
    Next i
    RecJsonText = out
End Function

' The value of "key" in a Google JSON answer (string or bare number); "" if
' absent. Enough for the flat sign-in answers and error messages.
Private Function RecJsonValue(ByVal json As String, ByVal key As String) As String
    Dim p As Long, c As String, out As String
    p = InStr(json, """" & key & """")
    If p = 0 Then Exit Function
    p = InStr(p + Len(key) + 2, json, ":")
    If p = 0 Then Exit Function
    p = p + 1
    Do While p <= Len(json) And InStr(" " & vbTab & vbCr & vbLf, Mid$(json, p, 1)) > 0
        p = p + 1
    Loop
    If Mid$(json, p, 1) = """" Then
        p = p + 1
        Do While p <= Len(json)
            c = Mid$(json, p, 1)
            If c = "\" Then
                p = p + 1
                c = Mid$(json, p, 1)
                If c = "n" Then c = " "
            ElseIf c = """" Then
                Exit Do
            End If
            out = out & c
            p = p + 1
        Loop
    Else
        Do While p <= Len(json) And InStr(",}] " & vbCr & vbLf, Mid$(json, p, 1)) = 0
            out = out & Mid$(json, p, 1)
            p = p + 1
        Loop
    End If
    RecJsonValue = out
End Function

Private Function RecUrlEncode(ByVal s As String) As String
    Dim i As Long, c As String, out As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c Like "[A-Za-z0-9._~-]" Then
            out = out & c
        Else
            out = out & "%" & Right$("0" & Hex$(AscW(c) And 255), 2)
        End If
    Next i
    RecUrlEncode = out
End Function

' A random UUID (CoCreateGuid; VBA's Rnd repeats per session, and a repeated
' id would be refused by the database as "already exists").
Private Function RecNewRunId() As String
    Dim g(0 To 15) As Byte, i As Long, s As String
    If CoCreateGuid(g(0)) <> 0 Then
        Randomize
        For i = 0 To 15
            g(i) = Int(Rnd * 256)
        Next i
    End If
    g(6) = (g(6) And 15) Or 64
    g(8) = (g(8) And 63) Or 128
    For i = 0 To 15
        If i = 4 Or i = 6 Or i = 8 Or i = 10 Then s = s & "-"
        s = s & Right$("0" & LCase$(Hex$(g(i))), 2)
    Next i
    RecNewRunId = s
End Function

' Now in UTC, RFC 3339 "YYYY-MM-DDTHH:MM:SSZ" (Now is local time).
Private Function RecUtcNow() As String
    Dim st(0 To 7) As Integer
    Call GetSystemTime(st(0))
    RecUtcNow = Format$(st(0), "0000") & "-" & Format$(st(1), "00") & "-" & Format$(st(3), "00") & _
                "T" & Format$(st(4), "00") & ":" & Format$(st(5), "00") & ":" & Format$(st(6), "00") & "Z"
End Function

Private Function RecSecondsSince(ByVal t0 As Single) As Double
    RecSecondsSince = Timer - t0
    If RecSecondsSince < 0 Then RecSecondsSince = RecSecondsSince + 86400#
End Function

' UTF-8 (ADODB writes a BOM, which every reader here drops).
Private Sub RecWriteText(ByVal path As String, ByVal txt As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText txt
    st.SaveToFile path, 2
    st.Close
End Sub

Private Function RecReadText(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    RecReadText = st.ReadText
    st.Close
End Function
