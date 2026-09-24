Option Explicit

' ============================================================================
'  SAP2000 - Wingwall Model Build  (.$2k model, opened through the API)
'
'  Writes the shell model as <workbook folder>\SAP2000\<workbook name>.$2k,
'  then starts its OWN SAP2000 through the API, opens that file, saves
'  <workbook name>.sdb beside it, runs the analysis and leaves SAP2000 open
'  with the solved model. An already-open SAP2000 is never touched, and no
'  worksheet is changed. Results are not pulled back yet.
'
'  WHY A FILE AND NOT OBJECT BY OBJECT like the culvert: the wall earth
'  pressures are joint patterns, and SAP2000 17's API cannot set joint
'  pattern values (SetPatternByXYZ/SetPatternByPressure store 0 - CSI bug
'  84182, fixed in v18.0.0). The .$2k carries them exactly; the owner chose
'  this hybrid on 2026-09-24. The bridge that runs SAP2000 (32-bit
'  PowerShell, because Excel is 64-bit) is the culvert builder's, byte for
'  byte - see the SAP2000 API BRIDGE block.
'
'  SAME INPUTS AND GEOMETRY AS THE MIDAS WINGWALL BUILDER. The cell readers
'  and GenerateGeometry are copied BYTE FOR BYTE from
'  midas-wingwall-model-build.bas, and the .$2k helpers from
'  sap2000-culvert-model-build.bas; tests/verify_sap2000_wingwall_build.py
'  and verify_no_helper_drift.py fail if a copy drifts. Change the original
'  and copy it across.
'
'  Owner's decisions (2026-09-24), against the hand-built example
'  SAP2000/3.00x3.00_H_3.75m.$2k:
'    - geometry from the workbook (the example's wall lengths differ)
'    - WALL MESH AS THE EXAMPLE: from z = 0 a band up to the rigid zone,
'      bands up to the lower wall end, then rows of height
'      (Hnear - Hfar) / N meeting the sloped top on the N length
'      divisions, with a triangle where the top crosses each column
'    - foundation = the MIDAS plates 1-4, 9, divided as DIVIDE_TABLE
'    - earth pressure = the sheet's four corner values per wall (rows
'      27-32), interpolated bilinearly over the wall face, as the MIDAS plane
'      loads do - not the example's depth-only profile
'    - springs: kv as area springs normal to the foundation (as the
'      example), 0.5 kv horizontally as joint springs of 0.5 kv x tributary
'      area - SAP2000's simple area spring acts normal to the face only
'    - MIDAS load cases with "_" dropped (EHS2_L -> EHS2L) and all eight
'      MIDAS combinations
'  Ec is INPUT!B19, falling back to 33 GPa as in the culvert SAP builder.
' ============================================================================


' ---------------------------------------------------------------------------
'  CONFIG
' ---------------------------------------------------------------------------

' Stamped into the report title and the file. Bump with every change.
Private Const SCRIPT_VERSION As String = "2026-09-24d"

' One line, no "_" continuation, no "|" - read by the updater's manifest.
Private Const SCRIPT_CHANGELOG As String = "No change to the model: the SAP2000 API bridge it shares with the culvert builder can now run steps after the analysis (used by the culvert results pull)."

' Identifies this module to the updater whatever it was named in Excel.
Private Const SCRIPT_ID As String = "sap2000-wingwall-model-build"

Private Const SAP_PROGRAM_VERSION As String = "17.3.0"
' SAP2000 runs hidden while the model is opened and solved, then its window
' is shown and left open. True keeps it visible throughout.
Private Const SAP_VISIBLE As Boolean = False

' Longest the macro waits for SAP2000, in seconds, before it gives up and
' closes the SAP2000 it started.
Private Const SAP_TIMEOUT_SEC As Long = 600

' Folder next to the workbook that receives the .$2k, the .sdb, SAP2000's
' analysis files, the script and its log.
Private Const SAP_FOLDER_NAME As String = "SAP2000"

' Where WriteModelText's .$2k goes (its first line names the file).
Private TWOK_PATH As String

Private Const PI_CONST As Double = 3.14159265358979

' Concrete - C30/37 as in the example. Unit weight is the MIDAS wingwall
' builder's fixed 25.
Private Const MATERIAL_NAME As String = "C30/37"
Private Const MATERIAL_FC As Double = 30000#
Private Const MATERIAL_POISN As Double = 0.2
Private Const MATERIAL_THERMAL As Double = 0.00001
Private Const MATERIAL_DEN As Double = 25
Private Const MATERIAL_ELAST_DEFAULT As Double = 33000000#  ' fallback if B19 is blank/invalid
Private Const CELL_MATERIAL_ELAST As String = "B19"
Private Const GRAVITY_ACCEL As Double = 9.80665
Private MATERIAL_ELAST As Double
Private INPUT_FALLBACKS As String

' Area sections, named as in the example.
Private Const SECTION_NAME_FOUNDATION As String = "TEMEL"
Private Const SECTION_NAME_WALL As String = "DUVAR"
Private Const SECTION_COLOR_FOUNDATION As String = "4610670"   ' 110,90,70 as R + 256 G + 65536 B
Private Const SECTION_COLOR_WALL As String = "4956220"         ' 60,160,75

' MIDAS load-case type -> SAP2000 DesignType (shared with the culvert).
Private Const SAP_DESIGN_TYPE_LIST As String = _
    "D|DEAD;EV|DEAD;EH|DEAD;L|LIVE;LS|LIVE;E|QUAKE;FP|OTHER"

' Joints closer than this are one joint - how walls and foundation come
' to share their base line.
Private Const MERGE_TOL As Double = 0.000001

Private Const PROJINFO_COMPANY As String = "DEHA"

' ---------------------------------------------------------------------------
'  INPUT SHEET - the MIDAS wingwall builder's cells (see CLAUDE.md
'  "Wingwall builder"): B20 clear opening, B21 stem, B24 foundation,
'  rows 22/23 LEFT/RIGHT length, angle, Hnear, Hfar; rows 27-32 pressures;
'  B34/C34 surcharge, B35 seismic coefficient; B8 kv; F18/F19 wall and
'  G18/G19 foundation mesh counts; G8/G9/K17/K18 project information.
' ---------------------------------------------------------------------------
Private Const INPUT_SHEET_NAME As String = "INPUT"
Private Const CELL_OPENING_WIDTH As String = "B20"
Private Const CELL_STEM_THICKNESS As String = "B21"
Private Const CELL_FOUNDATION_THICKNESS As String = "B24"
Private Const CELL_LEFT_LENGTH As String = "B22"
Private Const CELL_LEFT_ANGLE As String = "C22"
Private Const CELL_LEFT_HNEAR As String = "D22"
Private Const CELL_LEFT_HFAR As String = "E22"
Private Const CELL_RIGHT_LENGTH As String = "B23"
Private Const CELL_RIGHT_ANGLE As String = "C23"
Private Const CELL_RIGHT_HNEAR As String = "D23"
Private Const CELL_RIGHT_HFAR As String = "E23"
Private Const PROJECT_SHEET_NAME As String = "INPUT"
Private Const CELL_PROJECT_NO As String = "G8"
Private Const CELL_PROJECT_ANO As String = "G9"
Private Const CELL_PROJECT_ENGINEER As String = "K17"
Private Const CELL_PROJECT_REVISION As String = "K18"
Private Const CELL_MESH_X_WALL As String = "F18"
Private Const CELL_MESH_Y_WALL As String = "F19"
Private Const CELL_MESH_X_FOUND As String = "G18"
Private Const CELL_MESH_Y_FOUND As String = "G19"
Private Const CELL_SUBGRADE_MODULUS As String = "B8"
Private Const CELL_EP_LEFT_ATREST_ROW As Long = 27
Private Const CELL_EP_LEFT_ACTIVE_ROW As Long = 28
Private Const CELL_EP_LEFT_SEISMIC_ROW As Long = 29
Private Const CELL_EP_RIGHT_ATREST_ROW As Long = 30
Private Const CELL_EP_RIGHT_ACTIVE_ROW As Long = 31
Private Const CELL_EP_RIGHT_SEISMIC_ROW As Long = 32
Private Const CELL_SURCHARGE_LEFT As String = "B34"
Private Const CELL_SURCHARGE_RIGHT As String = "C34"
Private Const CELL_SEISMIC_COEF As String = "B35"

' Same values as the MIDAS builder (see its comments).
Private Const RIGID_ZONE_RATIO As Double = 0.5
Private Const GEN_STYPE As Long = 3
Private Const SECT_FOUNDATION As Long = 1
Private Const SECT_STEM As Long = 2
Private Const SPRING_H_RATIO As Double = 0.5

Private Type SideParams
    L As Double
    AngleDeg As Double
    Hnear As Double
    Hfar As Double
End Type

' Filled by the copied readers and GenerateGeometry.
Private PROJECT_NAME As String
Private PROJECT_ENGINEER As String
Private PROJECT_REVISION As String
Private SPRING_KV As Double
Private MESH_X_WALL As Long, MESH_Y_WALL As Long
Private MESH_X_FOUND As Long, MESH_Y_FOUND As Long
Private NODE_LIST As String
Private ELEMENT_LIST As String
Private GEO_LEFT_NEAR_X As Double, GEO_LEFT_NEAR_Y As Double
Private GEO_LEFT_FAR_X As Double, GEO_LEFT_FAR_Y As Double
Private GEO_LEFT_L As Double, GEO_LEFT_HNEAR As Double, GEO_LEFT_HFAR As Double
Private GEO_RIGHT_NEAR_X As Double, GEO_RIGHT_NEAR_Y As Double
Private GEO_RIGHT_FAR_X As Double, GEO_RIGHT_FAR_Y As Double
Private GEO_RIGHT_L As Double, GEO_RIGHT_HNEAR As Double, GEO_RIGHT_HFAR As Double
Private GEO_LEFT_ANGLE_DEG As Double
Private LOAD_L_ATREST(1 To 4) As Double
Private LOAD_L_ACTIVE(1 To 4) As Double
Private LOAD_L_SEISMIC(1 To 4) As Double
Private LOAD_R_ATREST(1 To 4) As Double
Private LOAD_R_ACTIVE(1 To 4) As Double
Private LOAD_R_SEISMIC(1 To 4) As Double
Private LOAD_SURCHARGE_LEFT As Double
Private LOAD_SURCHARGE_RIGHT As Double
Private LOAD_SEISMIC_COEF As Double

' The MIDAS wingwall builder's cases and combinations, verbatim (the test
' compares them). Case names lose their "_" on the way out (SapName);
' combination names are kept.
Private Const STLDCASE_LIST As String = _
    "DL|D;" & _
    "EHS2_L|EH;" & _
    "EHA2_L|EH;" & _
    "EHS2_R|EH;" & _
    "EHA2_R|EH;" & _
    "LSS1_L|LS;" & _
    "LSS1_R|LS;" & _
    "EQ_L|E;" & _
    "EQ_R|E;" & _
    "ATA_L|E;" & _
    "ATA_R|E"
Private Const LOADCOMB_LIST As String = _
    "SLS|ACTIVE|0|ST:DL:1,ST:EHS2_L:1,ST:EHS2_R:1,ST:LSS1_L:1,ST:LSS1_R:1|1;" & _
    "ULS|ACTIVE|0|ST:DL:1.35,ST:EHS2_L:1.35,ST:EHS2_R:1.35,ST:LSS1_L:1.45,ST:LSS1_R:1.45|1;" & _
    "EQ - 1|ACTIVE|0|ST:DL:1,ST:EHA2_L:1,ST:EQ_L:1,ST:ATA_L:1|1;" & _
    "EQ - 2|ACTIVE|0|ST:DL:1,ST:EHA2_R:1,ST:EQ_R:1,ST:ATA_R:1|1;" & _
    "ENV_SER|ACTIVE|1|CB:SLS:1|2;" & _
    "ENV_STR|ACTIVE|1|CB:ULS:1|2;" & _
    "ENV_EQ|ACTIVE|1|CB:EQ - 1:1,CB:EQ - 2:1|2;" & _
    "ENV_ALL|ACTIVE|1|CB:ENV_SER:1,CB:ENV_STR:1,CB:ENV_EQ:1|3"

' The MIDAS builder's plate division table, verbatim; only the foundation
' rows (1-4, 9) are used - the walls are meshed as the owner's example.
Private Const DIVIDE_TABLE As String = _
    "1|MESHX|1;" & _
    "2|MESHX|MESHY;" & _
    "3|MESHX|1;" & _
    "4|MESHX|1;" & _
    "5|MESHY|MESHX;" & _
    "6|1|MESHX;" & _
    "7|MESHY|MESHX;" & _
    "8|1|MESHX;" & _
    "9|MESHX|1"

' Wall pressure cases per side: case | load row (1 at-rest, 2 active,
' 3 seismic). The joint pattern of each case is named after it.
Private Const WALL_PRESSURE_CASES As String = "EHS2_#|1;EHA2_#|2;EQ_#|3"

' ---------------------------------------------------------------------------
'  SAP MODEL, built by BuildMesh. Joints 1..JT_COUNT; areas 1..AR_COUNT
'  (3 or 4 joints; AR_J(a, 4) = 0 for a triangle). AR_KIND: 0 foundation,
'  1 LEFT wall, 2 RIGHT wall. Each joint's wall-face position (JT_S along the
'  wall from the near end, JT_Z height) is kept for the pressure patterns.
' ---------------------------------------------------------------------------
Private JT_COUNT As Long
Private JT_X() As Double
Private JT_Y() As Double
Private JT_Z() As Double
Private JT_WALL() As Long         ' 0 foundation only, 1 LEFT wall, 2 RIGHT wall
Private JT_S() As Double          ' distance along its wall from the near end
Private AR_COUNT As Long
Private AR_J() As Long
Private AR_KIND() As Long
Private MESH_N As Long           ' length divisions (walls and foundation)
Private BAND_INFO As String      ' bands between rigid zone and lower wall end, "LEFT/RIGHT"
Private SPRING_COUNT As Long
Private TEXT_WARNINGS As String

' Output text, one entry per line; OUT_COUNT lines used.
Private OUT_LINES() As String
Private OUT_COUNT As Long


' ===========================================================================
'  ENTRY POINT
' ===========================================================================

Public Sub BuildSap2000WingwallModel()

    Dim report As String
    Dim ok As Boolean
    Dim folder As String, stem As String
    Dim res As String
    Dim stage As String
    Dim openingWidth As Double, foundT As Double, stemT As Double
    Dim leftSide As SideParams, rightSide As SideParams

    report = "SAP2000 wingwall model  [" & SCRIPT_VERSION & "]" & vbCrLf & _
             String(40, "-") & vbCrLf

    On Error GoTo Crashed

    stage = "SAP2000 folder"
    folder = SapFolder()
    ok = StepResult(report, "SAP2000 folder", IIf(Len(folder) = 0, _
             "the workbook has never been saved, so there is no folder to put the model " & _
             "in. Save it first.", ""))
    If ok Then stem = SapBasePath(folder)

    stage = "read inputs"
    If ok Then
        res = ReadWingwallInputs(openingWidth, foundT, stemT, leftSide, rightSide)
        If Len(res) = 0 Then res = ReadLoadInputs()
        If Len(res) = 0 And Len(INPUT_FALLBACKS) > 0 Then
            res = "WARN: used built-in defaults -" & INPUT_FALLBACKS
        End If
        ok = StepResult(report, "Read inputs", res)
    End If

    stage = "generate model"
    If ok Then
        Call ReadProjectName
        Call GenerateGeometry(openingWidth, stemT, leftSide, rightSide)
        ok = StepResult(report, "Mesh", BuildMesh(stemT))
    End If

    stage = "model file"
    If ok Then
        TWOK_PATH = stem & ".$2k"
        ok = StepResult(report, "Model text", WriteModelText(foundT, stemT))
    End If
    If ok Then ok = StepResult(report, "Write " & TWOK_PATH, WriteOutputFile(TWOK_PATH))

    stage = "SAP2000 script"
    If ok Then
        res = WriteOpenScript(stem)
        If Len(res) = 0 Then res = WriteScriptFile(stem & "_build.ps1")
        ok = StepResult(report, "SAP2000 script", res)
    End If

    stage = "SAP2000"
    If ok Then ok = StepResult(report, "SAP2000: open, save, analyse", _
                               RunSapScript(stem & "_build.ps1", stem & "_build.log"))

    If ok Then
        report = report & String(40, "-") & vbCrLf & _
                 JT_COUNT & " joints, " & AR_COUNT & " areas (" & MESH_N & " along each wall, " & _
                 BAND_INFO & " band(s) below the lower wall end, LEFT/RIGHT), " & SPRING_COUNT & _
                 " spring joints, " & (UBound(Split(LOADCOMB_LIST, ";")) + 1) & " combinations." & _
                 vbCrLf & "SAP2000 is open with the analysed model:" & vbCrLf & stem & ".sdb" & vbCrLf
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
'  COPIED - DO NOT EDIT HERE. Byte-identical copies from
'  midas-wingwall-model-build.bas (readers, geometry) and
'  sap2000-culvert-model-build.bas (.$2k helpers). Change the original and
'  copy the procedure across; the tests compare them.
' ===========================================================================

Private Function ReadWingwallInputs(ByRef openingWidth As Double, ByRef foundationThickness As Double, _
                                    ByRef stemThickness As Double, _
                                    ByRef leftSide As SideParams, ByRef rightSide As SideParams) As String

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(INPUT_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        ReadWingwallInputs = "Sheet not found: " & INPUT_SHEET_NAME
        Exit Function
    End If

    ReadWingwallInputs = RequirePositive(ws, CELL_OPENING_WIDTH, "opening width", openingWidth)
    If Len(ReadWingwallInputs) > 0 Then Exit Function
    ReadWingwallInputs = RequirePositive(ws, CELL_STEM_THICKNESS, "stem thickness", stemThickness)
    If Len(ReadWingwallInputs) > 0 Then Exit Function
    ReadWingwallInputs = RequirePositive(ws, CELL_FOUNDATION_THICKNESS, "foundation thickness", foundationThickness)
    If Len(ReadWingwallInputs) > 0 Then Exit Function

    ReadWingwallInputs = RequirePositive(ws, CELL_LEFT_LENGTH, "LEFT length", leftSide.L)
    If Len(ReadWingwallInputs) > 0 Then Exit Function
    ReadWingwallInputs = RequireNumber(ws, CELL_LEFT_ANGLE, "LEFT angle", leftSide.AngleDeg)
    If Len(ReadWingwallInputs) > 0 Then Exit Function
    ReadWingwallInputs = RequirePositive(ws, CELL_LEFT_HNEAR, "LEFT height @ face", leftSide.Hnear)
    If Len(ReadWingwallInputs) > 0 Then Exit Function
    ReadWingwallInputs = RequirePositive(ws, CELL_LEFT_HFAR, "LEFT height @ tip", leftSide.Hfar)
    If Len(ReadWingwallInputs) > 0 Then Exit Function

    ReadWingwallInputs = RequirePositive(ws, CELL_RIGHT_LENGTH, "RIGHT length", rightSide.L)
    If Len(ReadWingwallInputs) > 0 Then Exit Function
    ReadWingwallInputs = RequireNumber(ws, CELL_RIGHT_ANGLE, "RIGHT angle", rightSide.AngleDeg)
    If Len(ReadWingwallInputs) > 0 Then Exit Function
    ReadWingwallInputs = RequirePositive(ws, CELL_RIGHT_HNEAR, "RIGHT height @ face", rightSide.Hnear)
    If Len(ReadWingwallInputs) > 0 Then Exit Function
    ReadWingwallInputs = RequirePositive(ws, CELL_RIGHT_HFAR, "RIGHT height @ tip", rightSide.Hfar)
    If Len(ReadWingwallInputs) > 0 Then Exit Function

    ' Ec is not required to build - falls back to MATERIAL_ELAST_DEFAULT
    ' if B19 is blank/invalid, same non-blocking pattern as
    ' midas-culvert-model-build.bas's B43 read.
    ' The stem's 2-row split (GenerateGeometry) degenerates when a wall is
    ' not taller than its rigid zone, so refuse that before anything posts.
    ReadWingwallInputs = RigidZoneProblem("LEFT", leftSide, stemThickness)
    If Len(ReadWingwallInputs) > 0 Then Exit Function
    ReadWingwallInputs = RigidZoneProblem("RIGHT", rightSide, stemThickness)
    If Len(ReadWingwallInputs) > 0 Then Exit Function

    INPUT_FALLBACKS = ""
    If Not TryReadCell(ws, CELL_MATERIAL_ELAST, MATERIAL_ELAST) Or MATERIAL_ELAST <= 0 Then
        MATERIAL_ELAST = MATERIAL_ELAST_DEFAULT
        INPUT_FALLBACKS = INPUT_FALLBACKS & " Ec (" & INPUT_SHEET_NAME & "!" & CELL_MATERIAL_ELAST & _
                          ") missing or <= 0, used " & JsonNum(MATERIAL_ELAST_DEFAULT) & "."
    End If

    ' Subgrade modulus, likewise non-blocking: a model without it simply
    ' gets no foundation springs, and PostFoundationSprings warns.
    If Not TryReadCell(ws, CELL_SUBGRADE_MODULUS, SPRING_KV) Then
        SPRING_KV = 0
    End If

    ' Mesh counts, also non-blocking - a blank or < 1 cell just leaves that
    ' direction undivided, and PostDivideElements warns if nothing at all
    ' would be divided.
    MESH_X_WALL = ReadCount(ws, CELL_MESH_X_WALL)
    MESH_Y_WALL = ReadCount(ws, CELL_MESH_Y_WALL)
    MESH_X_FOUND = ReadCount(ws, CELL_MESH_X_FOUND)
    MESH_Y_FOUND = ReadCount(ws, CELL_MESH_Y_FOUND)

    ReadWingwallInputs = ""

End Function

Private Function ReadLoadInputs() As String

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(INPUT_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        ReadLoadInputs = "Sheet not found: " & INPUT_SHEET_NAME
        Exit Function
    End If

    ReadLoadInputs = ReadLoadRow(ws, CELL_EP_LEFT_ATREST_ROW, LOAD_L_ATREST(1), LOAD_L_ATREST(2), _
                                 LOAD_L_ATREST(3), LOAD_L_ATREST(4), "LEFT at-rest earth pressure")
    If Len(ReadLoadInputs) > 0 Then Exit Function

    ReadLoadInputs = ReadLoadRow(ws, CELL_EP_LEFT_ACTIVE_ROW, LOAD_L_ACTIVE(1), LOAD_L_ACTIVE(2), _
                                 LOAD_L_ACTIVE(3), LOAD_L_ACTIVE(4), "LEFT active earth pressure")
    If Len(ReadLoadInputs) > 0 Then Exit Function

    ReadLoadInputs = ReadLoadRow(ws, CELL_EP_LEFT_SEISMIC_ROW, LOAD_L_SEISMIC(1), LOAD_L_SEISMIC(2), _
                                 LOAD_L_SEISMIC(3), LOAD_L_SEISMIC(4), "LEFT seismic earth pressure")
    If Len(ReadLoadInputs) > 0 Then Exit Function

    ReadLoadInputs = ReadLoadRow(ws, CELL_EP_RIGHT_ATREST_ROW, LOAD_R_ATREST(1), LOAD_R_ATREST(2), _
                                 LOAD_R_ATREST(3), LOAD_R_ATREST(4), "RIGHT at-rest earth pressure")
    If Len(ReadLoadInputs) > 0 Then Exit Function

    ReadLoadInputs = ReadLoadRow(ws, CELL_EP_RIGHT_ACTIVE_ROW, LOAD_R_ACTIVE(1), LOAD_R_ACTIVE(2), _
                                 LOAD_R_ACTIVE(3), LOAD_R_ACTIVE(4), "RIGHT active earth pressure")
    If Len(ReadLoadInputs) > 0 Then Exit Function

    ReadLoadInputs = ReadLoadRow(ws, CELL_EP_RIGHT_SEISMIC_ROW, LOAD_R_SEISMIC(1), LOAD_R_SEISMIC(2), _
                                 LOAD_R_SEISMIC(3), LOAD_R_SEISMIC(4), "RIGHT seismic earth pressure")
    If Len(ReadLoadInputs) > 0 Then Exit Function

    If Not TryReadCell(ws, CELL_SURCHARGE_LEFT, LOAD_SURCHARGE_LEFT) Then
        ReadLoadInputs = "Non-numeric cell: " & CELL_SURCHARGE_LEFT & " (LEFT LS surcharge pressure)"
        Exit Function
    End If
    If Not TryReadCell(ws, CELL_SURCHARGE_RIGHT, LOAD_SURCHARGE_RIGHT) Then
        ReadLoadInputs = "Non-numeric cell: " & CELL_SURCHARGE_RIGHT & " (RIGHT LS surcharge pressure)"
        Exit Function
    End If
    If Not TryReadCell(ws, CELL_SEISMIC_COEF, LOAD_SEISMIC_COEF) Then
        ReadLoadInputs = "Non-numeric cell: " & CELL_SEISMIC_COEF & " (seismic inertia coefficient)"
        Exit Function
    End If
    ' Round the inertia coefficient to 3 decimals. The sheet usually
    ' computes it rather than having it typed in, so it arrives with a long
    ' float tail that just makes the ATA_L/ATA_R vectors unreadable in
    ' Civil NX. Rounded once here, at the point of reading, so both
    ' directions get the identical magnitude.
    LOAD_SEISMIC_COEF = Round(LOAD_SEISMIC_COEF, 3)

    ReadLoadInputs = ""

End Function

Private Function ReadLoadRow(ByVal ws As Worksheet, ByVal rowNo As Long, _
                             ByRef v1 As Double, ByRef v2 As Double, _
                             ByRef v3 As Double, ByRef v4 As Double, _
                             ByVal label As String) As String

    If Not TryReadCell(ws, "B" & rowNo, v1) Then
        ReadLoadRow = "Non-numeric cell: B" & rowNo & " (" & label & ", bottom near)"
        Exit Function
    End If
    If Not TryReadCell(ws, "C" & rowNo, v2) Then
        ReadLoadRow = "Non-numeric cell: C" & rowNo & " (" & label & ", bottom far)"
        Exit Function
    End If
    If Not TryReadCell(ws, "D" & rowNo, v3) Then
        ReadLoadRow = "Non-numeric cell: D" & rowNo & " (" & label & ", top far)"
        Exit Function
    End If
    If Not TryReadCell(ws, "E" & rowNo, v4) Then
        ReadLoadRow = "Non-numeric cell: E" & rowNo & " (" & label & ", top near)"
        Exit Function
    End If

    ReadLoadRow = ""

End Function

Private Function ReadCount(ByVal ws As Worksheet, ByVal addr As String) As Long

    Dim v As Double

    If Not TryReadCell(ws, addr, v) Then
        ReadCount = 1
    ElseIf v < 1 Then
        ReadCount = 1
    Else
        ReadCount = CLng(Int(v))
    End If

End Function

Private Function TryReadCell(ByVal ws As Worksheet, ByVal addr As String, ByRef result As Double) As Boolean
    If Not IsNumeric(ws.Range(addr).Value) Then
        TryReadCell = False
        Exit Function
    End If
    result = CDbl(ws.Range(addr).Value)
    TryReadCell = True
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

Private Function RequireNumber(ByVal ws As Worksheet, ByVal addr As String, _
                               ByVal label As String, ByRef result As Double) As String

    Dim v As Variant

    v = ws.Range(addr).Value

    If IsEmpty(v) Or IsError(v) Then
        RequireNumber = ws.Name & "!" & addr & " (" & label & ") is blank or an error value."
    ElseIf Not IsNumeric(v) Then
        RequireNumber = ws.Name & "!" & addr & " (" & label & ") is not a number."
    Else
        result = CDbl(v)
    End If

End Function

Private Function RigidZoneProblem(ByVal side As String, ByRef p As SideParams, _
                                  ByVal stemThickness As Double) As String

    Dim rz As Double

    rz = RIGID_ZONE_RATIO * stemThickness

    If p.Hnear <= rz Or p.Hfar <= rz Then
        RigidZoneProblem = side & " wall: heights " & JsonNum(p.Hnear) & " (face) / " & _
            JsonNum(p.Hfar) & " (tip) must both exceed the rigid-zone height " & _
            JsonNum(rz) & " (RIGID_ZONE_RATIO x stem thickness)."
    End If

End Function

Private Sub ReadProjectName()

    Dim ws As Worksheet
    Dim pNo As String, pAno As String
    Dim s As String

    PROJECT_NAME = ""
    PROJECT_ENGINEER = ""
    PROJECT_REVISION = ""

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PROJECT_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    pNo = Trim$(CStr(ws.Range(CELL_PROJECT_NO).Value))
    pAno = Trim$(CStr(ws.Range(CELL_PROJECT_ANO).Value))
    PROJECT_ENGINEER = Trim$(CStr(ws.Range(CELL_PROJECT_ENGINEER).Value))
    PROJECT_REVISION = Trim$(CStr(ws.Range(CELL_PROJECT_REVISION).Value))
    On Error GoTo 0

    s = pNo
    If Len(pAno) > 0 Then s = s & IIf(Len(s) > 0, " - ", "") & pAno

    PROJECT_NAME = s

End Sub

Private Sub GenerateGeometry(ByVal openingWidth As Double, ByVal stemThickness As Double, _
                             ByRef leftSide As SideParams, ByRef rightSide As SideParams)

    ' Shared source for both the stem's Z rigid-zone height and the
    ' foundation inner/outer offset magnitude (a true perpendicular
    ' distance now - see the offset vectors below).
    Dim rzF As Double
    rzF = RIGID_ZONE_RATIO * stemThickness

    Dim angleL As Double, angleR As Double
    angleL = DegToRad(leftSide.AngleDeg)
    angleR = DegToRad(rightSide.AngleDeg)

    ' Wall-line (centerline) points, placed directly in FINAL orientation -
    ' no separate rotation step. A is always the origin; Tl always lands on
    ' +X (LEFT wall's own edge = the global X axis).
    Dim aX As Double, aY As Double, bX As Double, bY As Double
    Dim tlX As Double, tlY As Double, trX As Double, trY As Double

    aX = 0: aY = 0
    tlX = leftSide.L: tlY = 0

    Dim edgeAngle As Double
    edgeAngle = PI_CONST / 2 + angleL

    ' Clear opening -> centreline spacing. NOT just "add one thickness":
    ' each wall crosses the face at its splay angle, so its half thickness
    ' covers halfT/Cos(angle) ALONG the face. Skipping this silently
    ' narrows the opening (2.000 clear comes out 1.938). halfT is the
    ' WALL's half thickness - equals rzF only because the ratio is 0.5.
    Dim halfT As Double, centreSpacing As Double
    halfT = stemThickness / 2
    centreSpacing = openingWidth + halfT / Cos(angleL) + halfT / Cos(angleR)

    bX = aX + centreSpacing * Cos(edgeAngle)
    bY = aY + centreSpacing * Sin(edgeAngle)

    Dim rightAngle As Double
    rightAngle = angleL + angleR
    trX = bX + rightSide.L * Cos(rightAngle)
    trY = bY + rightSide.L * Sin(rightAngle)

    ' All 8 offset corners sit rzF PERPENDICULAR off their own wall. What
    ' varies is where each is CUT, and both cuts are structure faces:
    ' NEAR on the culvert face (line A-B), FAR on the tip-to-tip edge
    ' (line Tl-Tr). That puts all four near corners on one straight face
    ' and all four far corners on one straight far edge - a true trapezoid
    ' for a symmetric pair, a clean quadrilateral otherwise. Two other cuts
    ' were tried live and rejected; see CLAUDE.md before changing this.
    '
    ' "in" = toward the opening. LEFT's direction is (1,0) so its inner
    ' perpendicular is +Y; RIGHT's rotated -90 deg gives (Sin, -Cos).
    Dim dLx As Double, dLy As Double, nLx As Double, nLy As Double
    dLx = 1: dLy = 0
    nLx = 0: nLy = 1

    Dim dRx As Double, dRy As Double, nRx As Double, nRy As Double
    dRx = Cos(rightAngle): dRy = Sin(rightAngle)
    nRx = Sin(rightAngle): nRy = -Cos(rightAngle)

    ' The two cut lines. Direction vectors only - OffsetCorner's math is
    ' scale-invariant, so neither needs normalising.
    Dim faceDx As Double, faceDy As Double
    faceDx = Cos(edgeAngle): faceDy = Sin(edgeAngle)

    Dim farDx As Double, farDy As Double
    farDx = trX - tlX: farDy = trY - tlY

    ' Inner/outer offset points.
    Dim aInX As Double, aInY As Double, tlInX As Double, tlInY As Double
    Dim aOutX As Double, aOutY As Double, tlOutX As Double, tlOutY As Double
    Dim bInX As Double, bInY As Double, trInX As Double, trInY As Double
    Dim bOutX As Double, bOutY As Double, trOutX As Double, trOutY As Double

    ' NEAR corners - cut on the culvert face (A and B both lie on it, so
    ' either one serves as the line's reference point).
    OffsetCorner aX, aY, rzF * nLx, rzF * nLy, dLx, dLy, _
                 aX, aY, faceDx, faceDy, aInX, aInY
    OffsetCorner aX, aY, -rzF * nLx, -rzF * nLy, dLx, dLy, _
                 aX, aY, faceDx, faceDy, aOutX, aOutY
    OffsetCorner bX, bY, rzF * nRx, rzF * nRy, dRx, dRy, _
                 aX, aY, faceDx, faceDy, bInX, bInY
    OffsetCorner bX, bY, -rzF * nRx, -rzF * nRy, dRx, dRy, _
                 aX, aY, faceDx, faceDy, bOutX, bOutY

    ' FAR corners - cut on the far edge (Tl and Tr both lie on it).
    OffsetCorner tlX, tlY, rzF * nLx, rzF * nLy, dLx, dLy, _
                 tlX, tlY, farDx, farDy, tlInX, tlInY
    OffsetCorner tlX, tlY, -rzF * nLx, -rzF * nLy, dLx, dLy, _
                 tlX, tlY, farDx, farDy, tlOutX, tlOutY
    OffsetCorner trX, trY, rzF * nRx, rzF * nRy, dRx, dRy, _
                 tlX, tlY, farDx, farDy, trInX, trInY
    OffsetCorner trX, trY, -rzF * nRx, -rzF * nRy, dRx, dRy, _
                 tlX, tlY, farDx, farDy, trOutX, trOutY

    NODE_LIST = _
        "7|" & JsonNum(aX) & "|" & JsonNum(aY) & "|0;" & _
        "5|" & JsonNum(bX) & "|" & JsonNum(bY) & "|0;" & _
        "8|" & JsonNum(tlX) & "|" & JsonNum(tlY) & "|0;" & _
        "6|" & JsonNum(trX) & "|" & JsonNum(trY) & "|0;" & _
        "9|" & JsonNum(aInX) & "|" & JsonNum(aInY) & "|0;" & _
        "10|" & JsonNum(tlInX) & "|" & JsonNum(tlInY) & "|0;" & _
        "2|" & JsonNum(aOutX) & "|" & JsonNum(aOutY) & "|0;" & _
        "3|" & JsonNum(tlOutX) & "|" & JsonNum(tlOutY) & "|0;" & _
        "11|" & JsonNum(bInX) & "|" & JsonNum(bInY) & "|0;" & _
        "12|" & JsonNum(trInX) & "|" & JsonNum(trInY) & "|0;" & _
        "1|" & JsonNum(bOutX) & "|" & JsonNum(bOutY) & "|0;" & _
        "4|" & JsonNum(trOutX) & "|" & JsonNum(trOutY) & "|0;" & _
        "13|" & JsonNum(aX) & "|" & JsonNum(aY) & "|" & JsonNum(leftSide.Hnear) & ";" & _
        "14|" & JsonNum(bX) & "|" & JsonNum(bY) & "|" & JsonNum(rightSide.Hnear) & ";" & _
        "15|" & JsonNum(tlX) & "|" & JsonNum(tlY) & "|" & JsonNum(leftSide.Hfar) & ";" & _
        "16|" & JsonNum(trX) & "|" & JsonNum(trY) & "|" & JsonNum(rightSide.Hfar) & ";" & _
        "17|" & JsonNum(aX) & "|" & JsonNum(aY) & "|" & JsonNum(rzF) & ";" & _
        "18|" & JsonNum(tlX) & "|" & JsonNum(tlY) & "|" & JsonNum(rzF) & ";" & _
        "19|" & JsonNum(bX) & "|" & JsonNum(bY) & "|" & JsonNum(rzF) & ";" & _
        "20|" & JsonNum(trX) & "|" & JsonNum(trY) & "|" & JsonNum(rzF)

    ELEMENT_LIST = _
        "1|6|5|11|12|" & GEN_STYPE & "|" & SECT_FOUNDATION & ";" & _
        "2|12|11|9|10|" & GEN_STYPE & "|" & SECT_FOUNDATION & ";" & _
        "3|10|9|7|8|" & GEN_STYPE & "|" & SECT_FOUNDATION & ";" & _
        "4|8|7|2|3|" & GEN_STYPE & "|" & SECT_FOUNDATION & ";" & _
        "5|19|14|16|20|" & GEN_STYPE & "|" & SECT_STEM & ";" & _
        "6|5|19|20|6|" & GEN_STYPE & "|" & SECT_STEM & ";" & _
        "7|13|17|18|15|" & GEN_STYPE & "|" & SECT_STEM & ";" & _
        "8|17|7|8|18|" & GEN_STYPE & "|" & SECT_STEM & ";" & _
        "9|4|1|5|6|" & GEN_STYPE & "|" & SECT_FOUNDATION

    ' Hand the wall geometry to the load builders (see the GEO_* block).
    GEO_LEFT_NEAR_X = aX: GEO_LEFT_NEAR_Y = aY
    GEO_LEFT_FAR_X = tlX: GEO_LEFT_FAR_Y = tlY
    GEO_LEFT_L = leftSide.L
    GEO_LEFT_HNEAR = leftSide.Hnear: GEO_LEFT_HFAR = leftSide.Hfar

    GEO_RIGHT_NEAR_X = bX: GEO_RIGHT_NEAR_Y = bY
    GEO_RIGHT_FAR_X = trX: GEO_RIGHT_FAR_Y = trY
    GEO_RIGHT_L = rightSide.L
    GEO_RIGHT_HNEAR = rightSide.Hnear: GEO_RIGHT_HFAR = rightSide.Hfar

    ' The LEFT wall's splay angle, kept for PostNamedUcs - it is the
    ' rotation of the foundation's UCS. See the db/NUCS block.
    GEO_LEFT_ANGLE_DEG = leftSide.AngleDeg

End Sub

Private Function DegToRad(ByVal deg As Double) As Double
    DegToRad = deg * PI_CONST / 180
End Function

Private Sub OffsetCorner(ByVal px As Double, ByVal py As Double, _
                         ByVal offX As Double, ByVal offY As Double, _
                         ByVal dirX As Double, ByVal dirY As Double, _
                         ByVal cutPx As Double, ByVal cutPy As Double, _
                         ByVal cutDx As Double, ByVal cutDy As Double, _
                         ByRef outX As Double, ByRef outY As Double)

    Dim sx As Double, sy As Double
    Dim ex As Double, ey As Double
    Dim t As Double

    sx = px + offX: sy = py + offY
    ex = sx - cutPx: ey = sy - cutPy

    ' 2-D cross products. The denominator only vanishes if a wall runs
    ' parallel to its own cut line - a 90 deg splay, not a real wingwall.
    t = -(ex * cutDy - ey * cutDx) / (dirX * cutDy - dirY * cutDx)

    outX = sx + t * dirX
    outY = sy + t * dirY

End Sub

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

Private Sub WriteRebarSizeTable()

    Call TableStart("REBAR SIZES")
    Call Emit("   RebarID=#4   Area=0.000129032001922727   Diameter=0.0127")
    Call Emit("   RebarID=#9   Area=0.00064516   Diameter=0.0286512005329132")
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

Private Sub TableStart(ByVal tableName As String)
    Call Emit("TABLE:  """ & tableName & """")
End Sub

Private Sub TableEnd()
    Call Emit(" ")
End Sub



Private Function SapQuote(ByVal s As String) As String
    s = Replace(s, """", "'")
    If Len(s) = 0 Or InStr(s, " ") > 0 Or InStr(s, "=") > 0 Then
        SapQuote = """" & s & """"
    Else
        SapQuote = s
    End If
End Function


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


' ===========================================================================
'  MESH
'  Foundation: the MIDAS plates 1-4 and 9 (ELEMENT_LIST, section 1), each
'  divided bilinearly as DIVIDE_TABLE says (MESHX -> MESH_N along the walls,
'  MESHY -> G19 across). Walls: see BuildWall. Joints are merged by
'  coordinate, which is how the walls and the foundation come to share the
'  wall base line - the length division must therefore be the same on both
'  (F18 = G18).
'
'  Coordinates come out of NODE_LIST with Val, never CDbl: the list holds
'  period decimals, and CDbl reads "0.2" as 2 on a Turkish Windows.
' ===========================================================================

Private Function BuildMesh(ByVal stemT As Double) As String

    Dim rows() As String
    Dim f() As String
    Dim i As Long, no As Long
    Dim nx(1 To 20) As Double, ny(1 To 20) As Double
    Dim divX As Long, divY As Long
    Dim msg As String

    JT_COUNT = 0
    AR_COUNT = 0
    ReDim JT_X(1 To 256): ReDim JT_Y(1 To 256): ReDim JT_Z(1 To 256)
    ReDim JT_WALL(1 To 256): ReDim JT_S(1 To 256)
    ReDim AR_J(1 To 256, 1 To 4): ReDim AR_KIND(1 To 256)

    If SPRING_KV <= 0 Then
        BuildMesh = INPUT_SHEET_NAME & "!" & CELL_SUBGRADE_MODULUS & " (kv) is not above 0 - without " & _
                    "springs the model has no supports and SAP2000 cannot solve it."
        Exit Function
    End If

    If MESH_X_FOUND <> MESH_X_WALL Then
        BuildMesh = INPUT_SHEET_NAME & "!" & CELL_MESH_X_WALL & " (" & MESH_X_WALL & ") and " & _
                    CELL_MESH_X_FOUND & " (" & MESH_X_FOUND & ") must be equal - walls and " & _
                    "foundation share their joints along the wall base."
        Exit Function
    End If
    MESH_N = MESH_X_WALL

    rows = Split(NODE_LIST, ";")
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
        no = CLng(f(0))
        nx(no) = Val(f(1))
        ny(no) = Val(f(2))
    Next i

    ' ELEMENT_LIST rows: "no|n1|n2|n3|n4|stype|sect"
    rows = Split(ELEMENT_LIST, ";")
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
        If CLng(f(6)) = SECT_FOUNDATION Then
            divX = DivideCount(CLng(f(0)), 1)
            divY = DivideCount(CLng(f(0)), 2)
            Call MeshFoundationPlate(nx(CLng(f(1))), ny(CLng(f(1))), nx(CLng(f(2))), ny(CLng(f(2))), _
                                     nx(CLng(f(3))), ny(CLng(f(3))), nx(CLng(f(4))), ny(CLng(f(4))), _
                                     divX, divY)
        End If
    Next i

    BAND_INFO = ""
    msg = BuildWall(1, GEO_LEFT_NEAR_X, GEO_LEFT_NEAR_Y, GEO_LEFT_FAR_X, GEO_LEFT_FAR_Y, _
                    GEO_LEFT_HNEAR, GEO_LEFT_HFAR, GEO_RIGHT_NEAR_X, GEO_RIGHT_NEAR_Y, stemT)
    If Len(msg) = 0 Then
        msg = BuildWall(2, GEO_RIGHT_NEAR_X, GEO_RIGHT_NEAR_Y, GEO_RIGHT_FAR_X, GEO_RIGHT_FAR_Y, _
                        GEO_RIGHT_HNEAR, GEO_RIGHT_HFAR, GEO_LEFT_NEAR_X, GEO_LEFT_NEAR_Y, stemT)
    End If
    BuildMesh = msg

End Function

' DIVIDE_TABLE's count for one plate: axis 1 = local x, 2 = local y.
Private Function DivideCount(ByVal plateNo As Long, ByVal axis As Long) As Long

    Dim rows() As String
    Dim f() As String
    Dim i As Long

    DivideCount = 1
    rows = Split(DIVIDE_TABLE, ";")
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
        If CLng(f(0)) = plateNo Then
            Select Case f(axis)
                Case "MESHX": DivideCount = MESH_N
                Case "MESHY": DivideCount = MESH_Y_FOUND
                Case Else: DivideCount = CLng(f(axis))
            End Select
            Exit Function
        End If
    Next i

End Function

' One foundation plate, corners in MIDAS node order (local x = 1->2, local
' y = 1->4), divided bilinearly; every area faces up (+Z).
Private Sub MeshFoundationPlate(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                                ByVal x3 As Double, ByVal y3 As Double, ByVal x4 As Double, ByVal y4 As Double, _
                                ByVal divX As Long, ByVal divY As Long)

    Dim g() As Long
    Dim i As Long, j As Long
    Dim u As Double, v As Double

    ReDim g(0 To divX, 0 To divY)
    For i = 0 To divX
        For j = 0 To divY
            u = i / divX
            v = j / divY
            g(i, j) = AddJoint((1 - u) * (1 - v) * x1 + u * (1 - v) * x2 + u * v * x3 + (1 - u) * v * x4, _
                               (1 - u) * (1 - v) * y1 + u * (1 - v) * y2 + u * v * y3 + (1 - u) * v * y4, _
                               0, 0, 0)
        Next j
    Next i

    For i = 0 To divX - 1
        For j = 0 To divY - 1
            Call AddArea(g(i, j), g(i + 1, j), g(i + 1, j + 1), g(i, j + 1), 0, 0, 0, 1)
        Next j
    Next i

End Sub

' One wingwall, meshed as the owner's example: columns at MESH_N equal
' divisions of the wall line; up each column the levels
'   0, rigid-zone top, BAND equal bands up to the lower wall end (zLow),
'   then rows of height h = |Hnear - Hfar| / MESH_N up to that column's top.
' Column tops step by exactly one row per column, so the sloped top passes
' through grid points and each column strip ends in one triangle. BAND =
' round((zLow - rigid zone) / h), at least 1 (the owner's rule: bands about
' as tall as the rows above); a wall with Hnear = Hfar uses the column width.
' Every area faces the soil (away from the other wall), as in the example,
' so positive pressure on the Top face pushes toward the opening.
Private Function BuildWall(ByVal side As Long, ByVal nearX As Double, ByVal nearY As Double, _
                           ByVal farX As Double, ByVal farY As Double, ByVal hNear As Double, _
                           ByVal hFar As Double, ByVal otherX As Double, ByVal otherY As Double, _
                           ByVal stemT As Double) As String

    Dim wallLen As Double, dx As Double, dy As Double
    Dim outX As Double, outY As Double
    Dim rz As Double, zLow As Double, h As Double
    Dim nb As Long, nLow As Long, maxUp As Long
    Dim k As Long, lv As Long, m As Long
    Dim s As Double, z As Double, zTop As Double
    Dim g() As Long
    Dim up() As Long
    Dim lowZ() As Double
    Dim t As Long, u As Long, mMin As Long

    wallLen = Sqr((farX - nearX) ^ 2 + (farY - nearY) ^ 2)
    dx = (farX - nearX) / wallLen
    dy = (farY - nearY) / wallLen

    ' Outward = the perpendicular pointing away from the other wall.
    outX = -dy: outY = dx
    If outX * (otherX - nearX) + outY * (otherY - nearY) > 0 Then outX = -outX: outY = -outY

    rz = RIGID_ZONE_RATIO * stemT
    zLow = IIf(hNear < hFar, hNear, hFar)
    h = Abs(hNear - hFar) / MESH_N
    If h > MERGE_TOL Then
        nb = Int((zLow - rz) / h + 0.5)
    Else
        nb = Int((zLow - rz) / (wallLen / MESH_N) + 0.5)
    End If
    If nb < 1 Then nb = 1

    ' Levels shared by every column: 0, rz, rz + j (zLow - rz) / nb.
    nLow = nb + 1
    ReDim lowZ(0 To nLow)
    lowZ(0) = 0
    For lv = 1 To nLow
        lowZ(lv) = rz + (lv - 1) * (zLow - rz) / nb
    Next lv
    lowZ(nLow) = zLow

    maxUp = IIf(h > MERGE_TOL, MESH_N, 0)
    ReDim g(0 To MESH_N, 0 To nLow + maxUp)
    ReDim up(0 To MESH_N)

    For k = 0 To MESH_N
        s = wallLen * k / MESH_N
        zTop = hNear + (hFar - hNear) * k / MESH_N
        If h > MERGE_TOL Then up(k) = Int((zTop - zLow) / h + 0.5) Else up(k) = 0
        For lv = 0 To nLow + up(k)
            If lv <= nLow Then
                z = lowZ(lv)
            ElseIf lv = nLow + up(k) Then
                z = zTop
            Else
                z = zLow + (lv - nLow) * h
            End If
            g(k, lv) = AddJoint(nearX + s * dx, nearY + s * dy, z, side, s)
        Next lv
    Next k

    For k = 0 To MESH_N - 1
        For lv = 0 To nLow - 1
            Call AddArea(g(k, lv), g(k + 1, lv), g(k + 1, lv + 1), g(k, lv + 1), side, outX, outY, 0)
        Next lv
        mMin = IIf(up(k) < up(k + 1), up(k), up(k + 1))
        For m = 0 To mMin - 1
            Call AddArea(g(k, nLow + m), g(k + 1, nLow + m), g(k + 1, nLow + m + 1), g(k, nLow + m + 1), _
                         side, outX, outY, 0)
        Next m
        If up(k) <> up(k + 1) Then
            If Abs(up(k) - up(k + 1)) <> 1 Then
                BuildWall = "wall " & side & ": column tops " & up(k) & "/" & up(k + 1) & _
                            " rows apart - the sloped top does not step one row per column."
                Exit Function
            End If
            If up(k) > up(k + 1) Then t = k: u = k + 1 Else t = k + 1: u = k
            Call AddArea(g(t, nLow + mMin + 1), g(t, nLow + mMin), g(u, nLow + mMin), 0, _
                         side, outX, outY, 0)
        End If
    Next k

    BAND_INFO = BAND_INFO & IIf(Len(BAND_INFO) > 0, "/", "") & nb

End Function

' Returns the joint at (x, y, z), adding it if new. side/s record which
' wall face the joint is on (0 = foundation only) and how far along it.
Private Function AddJoint(ByVal x As Double, ByVal y As Double, ByVal z As Double, _
                          ByVal side As Long, ByVal s As Double) As Long

    Dim i As Long

    For i = 1 To JT_COUNT
        If Abs(JT_X(i) - x) < MERGE_TOL And Abs(JT_Y(i) - y) < MERGE_TOL And Abs(JT_Z(i) - z) < MERGE_TOL Then
            If side > 0 Then JT_WALL(i) = side: JT_S(i) = s
            AddJoint = i
            Exit Function
        End If
    Next i

    JT_COUNT = JT_COUNT + 1
    If JT_COUNT > UBound(JT_X) Then
        ReDim Preserve JT_X(1 To JT_COUNT * 2): ReDim Preserve JT_Y(1 To JT_COUNT * 2)
        ReDim Preserve JT_Z(1 To JT_COUNT * 2): ReDim Preserve JT_WALL(1 To JT_COUNT * 2)
        ReDim Preserve JT_S(1 To JT_COUNT * 2)
    End If
    JT_X(JT_COUNT) = x: JT_Y(JT_COUNT) = y: JT_Z(JT_COUNT) = z
    JT_WALL(JT_COUNT) = side: JT_S(JT_COUNT) = s
    AddJoint = JT_COUNT

End Function

' Adds one area (j4 = 0 for a triangle), wound so its local 3 axis (right-
' hand rule over the joint order, as SAP2000 defines it) points along
' (wantX, wantY, wantZ).
Private Sub AddArea(ByVal j1 As Long, ByVal j2 As Long, ByVal j3 As Long, ByVal j4 As Long, _
                    ByVal kind As Long, ByVal wantX As Double, ByVal wantY As Double, ByVal wantZ As Double)

    Dim jl As Long
    Dim ax As Double, ay As Double, az As Double
    Dim bx As Double, by As Double, bz As Double
    Dim tmp As Long

    jl = IIf(j4 > 0, j4, j3)
    ax = JT_X(j2) - JT_X(j1): ay = JT_Y(j2) - JT_Y(j1): az = JT_Z(j2) - JT_Z(j1)
    bx = JT_X(jl) - JT_X(j1): by = JT_Y(jl) - JT_Y(j1): bz = JT_Z(jl) - JT_Z(j1)
    If (ay * bz - az * by) * wantX + (az * bx - ax * bz) * wantY + (ax * by - ay * bx) * wantZ < 0 Then
        If j4 > 0 Then
            tmp = j2: j2 = j4: j4 = tmp
        Else
            tmp = j2: j2 = j3: j3 = tmp
        End If
    End If

    AR_COUNT = AR_COUNT + 1
    If AR_COUNT > UBound(AR_KIND) Then
        AR_J = GrowAreaJoints(AR_J, AR_COUNT * 2)
        ReDim Preserve AR_KIND(1 To AR_COUNT * 2)
    End If
    AR_J(AR_COUNT, 1) = j1: AR_J(AR_COUNT, 2) = j2
    AR_J(AR_COUNT, 3) = j3: AR_J(AR_COUNT, 4) = j4
    AR_KIND(AR_COUNT) = kind

End Sub

' ReDim Preserve can only grow the LAST dimension, so the area table is
' copied into a longer one instead.
Private Function GrowAreaJoints(ByRef src() As Long, ByVal newCount As Long) As Long()

    Dim out() As Long
    Dim i As Long, c As Long

    ReDim out(1 To newCount, 1 To 4)
    For i = 1 To UBound(src, 1)
        For c = 1 To 4
            out(i, c) = src(i, c)
        Next c
    Next i
    GrowAreaJoints = out

End Function

Private Function AreaSize(ByVal a As Long) As Double

    Dim p As Long, q As Long, r As Long
    Dim cx As Double, cy As Double, cz As Double

    ' Half the cross product of the diagonals (a quad) or of two edges.
    If AR_J(a, 4) > 0 Then
        p = AR_J(a, 1): q = AR_J(a, 3): r = AR_J(a, 2)
        cx = (JT_Y(q) - JT_Y(p)) * (JT_Z(AR_J(a, 4)) - JT_Z(r)) - (JT_Z(q) - JT_Z(p)) * (JT_Y(AR_J(a, 4)) - JT_Y(r))
        cy = (JT_Z(q) - JT_Z(p)) * (JT_X(AR_J(a, 4)) - JT_X(r)) - (JT_X(q) - JT_X(p)) * (JT_Z(AR_J(a, 4)) - JT_Z(r))
        cz = (JT_X(q) - JT_X(p)) * (JT_Y(AR_J(a, 4)) - JT_Y(r)) - (JT_Y(q) - JT_Y(p)) * (JT_X(AR_J(a, 4)) - JT_X(r))
    Else
        p = AR_J(a, 1): q = AR_J(a, 2): r = AR_J(a, 3)
        cx = (JT_Y(q) - JT_Y(p)) * (JT_Z(r) - JT_Z(p)) - (JT_Z(q) - JT_Z(p)) * (JT_Y(r) - JT_Y(p))
        cy = (JT_Z(q) - JT_Z(p)) * (JT_X(r) - JT_X(p)) - (JT_X(q) - JT_X(p)) * (JT_Z(r) - JT_Z(p))
        cz = (JT_X(q) - JT_X(p)) * (JT_Y(r) - JT_Y(p)) - (JT_Y(q) - JT_Y(p)) * (JT_X(r) - JT_X(p))
    End If
    AreaSize = Sqr(cx * cx + cy * cy + cz * cz) / 2

End Function


' ===========================================================================
'  MODEL TEXT - tables in the order SAP2000's own export uses (see the
'  example file).
' ===========================================================================

Private Function WriteModelText(ByVal foundT As Double, ByVal stemT As Double) As String

    Dim res As String

    OUT_COUNT = 0
    ReDim OUT_LINES(1 To 1024)
    TEXT_WARNINGS = ""
    SPRING_COUNT = 0

    Call Emit("File " & TWOK_PATH & " was saved on " & Month(Now) & "." & Day(Now) & _
              "." & Format$(Year(Now) Mod 100, "00") & " at " & Format$(Now, "hh:mm:ss"))
    Call Emit(" ")

    Call TableStart("ACTIVE DEGREES OF FREEDOM")
    Call Emit("   UX=Yes   UY=Yes   UZ=Yes   RX=Yes   RY=Yes   RZ=Yes")
    Call TableEnd

    Call TableStart("ANALYSIS OPTIONS")
    Call Emit("   Solver=Advanced   SolverProc=Auto   Force32Bit=No   StiffCase=None   GeomMod=None")
    Call TableEnd

    Call WriteAreaLoadTables
    Call WriteAreaTables(foundT, stemT)

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

    Call TableStart("GROUPS 1 - DEFINITIONS")
    Call Emit("   GroupName=ALL   Selection=Yes   SectionCut=Yes   Steel=Yes   Concrete=Yes" & _
              "   Aluminum=Yes   ColdFormed=Yes   Stage=Yes   Bridge=Yes   AutoSeismic=No" & _
              "   AutoWind=No   SelDesSteel=No   SelDesAlum=No   SelDesCold=No   MassWeight=Yes   Color=Red")
    Call TableEnd

    Call WriteJointTable
    Call WriteJointPatternTables
    Call WriteJointSpringTable

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

' GRAVITY: ATA_L / ATA_R, the self-weight along +Y / -Y (the MIDAS db/BODF
' records FV = [0, +coef, 0] and [0, -coef, 0]) on every area.
' SURFACE PRESSURE: each wall's three earth-pressure cases, Pressure 1 times
' that case's joint pattern (WallPressure) on the Top face - the face
' toward the soil, so it pushes toward the opening.
' UNIFORM: the surcharge, local 3 (toward the soil), value as on the sheet
' (-10 pushes toward the opening), on every area of its wall.
Private Sub WriteAreaLoadTables()

    Dim a As Long, c As Long
    Dim cases() As String
    Dim f() As String
    Dim sfx As String

    Call TableStart("AREA LOADS - GRAVITY")
    For a = 1 To AR_COUNT
        Call Emit("   Area=" & a & "   LoadPat=" & SapName("ATA_L") & "   CoordSys=GLOBAL   MultiplierX=0" & _
                  "   MultiplierY=" & SapNum(LOAD_SEISMIC_COEF) & "   MultiplierZ=0")
        Call Emit("   Area=" & a & "   LoadPat=" & SapName("ATA_R") & "   CoordSys=GLOBAL   MultiplierX=0" & _
                  "   MultiplierY=" & SapNum(-LOAD_SEISMIC_COEF) & "   MultiplierZ=0")
    Next a
    Call TableEnd

    cases = Split(WALL_PRESSURE_CASES, ";")
    Call TableStart("AREA LOADS - SURFACE PRESSURE")
    For a = 1 To AR_COUNT
        If AR_KIND(a) > 0 Then
            sfx = IIf(AR_KIND(a) = 1, "L", "R")
            For c = 0 To UBound(cases)
                f = Split(cases(c), "|")
                Call Emit("   Area=" & a & "   LoadPat=" & SapName(Replace(f(0), "#", sfx)) & _
                          "   Face=Top   Pressure=1   JtPattern=" & SapName(Replace(f(0), "#", sfx)))
            Next c
        End If
    Next a
    Call TableEnd

    Call TableStart("AREA LOADS - UNIFORM")
    For a = 1 To AR_COUNT
        If AR_KIND(a) = 1 Then
            Call Emit("   Area=" & a & "   LoadPat=" & SapName("LSS1_L") & "   CoordSys=Local   Dir=3" & _
                      "   UnifLoad=" & SapNum(LOAD_SURCHARGE_LEFT))
        ElseIf AR_KIND(a) = 2 Then
            Call Emit("   Area=" & a & "   LoadPat=" & SapName("LSS1_R") & "   CoordSys=Local   Dir=3" & _
                      "   UnifLoad=" & SapNum(LOAD_SURCHARGE_RIGHT))
        End If
    Next a
    Call TableEnd

End Sub

' Local axes (foundation turned by the LEFT wall's angle, like the MIDAS
' "FOUND" UCS and the example), sections, and the vertical springs: kv
' normal to every foundation area, as the example writes them.
Private Sub WriteAreaTables(ByVal foundT As Double, ByVal stemT As Double)

    Dim a As Long

    If GEO_LEFT_ANGLE_DEG <> 0 Then
        Call TableStart("AREA LOCAL AXES ASSIGNMENTS 1 - TYPICAL")
        For a = 1 To AR_COUNT
            If AR_KIND(a) = 0 Then Call Emit("   Area=" & a & "   Angle=" & SapNum(GEO_LEFT_ANGLE_DEG))
        Next a
        Call TableEnd
    End If

    Call TableStart("AREA SECTION ASSIGNMENTS")
    For a = 1 To AR_COUNT
        Call Emit("   Area=" & a & "   Section=" & IIf(AR_KIND(a) = 0, SECTION_NAME_FOUNDATION, _
                  SECTION_NAME_WALL) & "   MatProp=Default")
    Next a
    Call TableEnd

    Call TableStart("AREA SECTION PROPERTIES")
    Call Emit(AreaSectionRow(SECTION_NAME_WALL, stemT, SECTION_COLOR_WALL))
    Call Emit("        MMod=1   WMod=1")
    Call Emit(AreaSectionRow(SECTION_NAME_FOUNDATION, foundT, SECTION_COLOR_FOUNDATION))
    Call Emit("        MMod=1   WMod=1")
    Call TableEnd

    Call TableStart("AREA SECTION PROPERTY DESIGN PARAMETERS")
    Call Emit("   Section=" & SECTION_NAME_WALL & "   RebarMat=None   RebarOpt=Default")
    Call Emit("   Section=" & SECTION_NAME_FOUNDATION & "   RebarMat=None   RebarOpt=Default")
    Call TableEnd

    If SPRING_KV > 0 Then
        Call TableStart("AREA SPRING ASSIGNMENTS")
        For a = 1 To AR_COUNT
            If AR_KIND(a) = 0 Then
                Call Emit("   Area=" & a & "   Type=Simple   Stiffness=" & SapNum(SPRING_KV) & _
                          "   SimpleType=""Tension and Compression""   Face=Top" & _
                          "   Dir1Type=""Normal To Face""   NormalDir=Inward")
            End If
        Next a
        Call TableEnd
    End If

End Sub

' First line of an area section row; the caller adds the continuation.
Private Function AreaSectionRow(ByVal nm As String, ByVal t As Double, ByVal colour As String) As String
    AreaSectionRow = "   Section=" & nm & "   Material=" & MATERIAL_NAME & "   MatAngle=0" & _
                     "   AreaType=Shell   Type=Shell-Thick   DrillDOF=Yes   Thickness=" & SapNum(t) & _
                     "   BendThick=" & SapNum(t) & "   Color=" & colour & "   F11Mod=1   F22Mod=1" & _
                     "   F12Mod=1   M11Mod=1   M22Mod=1   M12Mod=1   V13Mod=1   V23Mod=1 _"
End Function

Private Sub WriteStaticCaseTable()

    Dim rows() As String
    Dim f() As String
    Dim i As Long

    Call TableStart("CASE - STATIC 1 - LOAD ASSIGNMENTS")
    rows = Split(STLDCASE_LIST, ";")
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
        Call Emit("   Case=" & SapName(f(0)) & "   LoadType=""Load pattern""   LoadName=" & _
                  SapName(f(0)) & "   LoadSF=1")
    Next i
    Call TableEnd

End Sub


Private Sub WriteConnectivityTable()

    Dim a As Long
    Dim s As String

    Call TableStart("CONNECTIVITY - AREA")
    For a = 1 To AR_COUNT
        s = "   Area=" & a & "   Joint1=" & AR_J(a, 1) & "   Joint2=" & AR_J(a, 2) & "   Joint3=" & AR_J(a, 3)
        If AR_J(a, 4) > 0 Then s = s & "   Joint4=" & AR_J(a, 4)
        Call Emit(s)
    Next a
    Call TableEnd

End Sub

' Wall-face corners are SpecialJt=Yes; the rest are mesh joints.
Private Sub WriteJointTable()

    Dim jt As Long

    Call TableStart("JOINT COORDINATES")
    For jt = 1 To JT_COUNT
        Call Emit("   Joint=" & jt & "   CoordSys=GLOBAL   CoordType=Cartesian   XorR=" & _
                  SapNum(JT_X(jt)) & "   Y=" & SapNum(JT_Y(jt)) & "   Z=" & SapNum(JT_Z(jt)) & _
                  "   SpecialJt=No")
    Next jt
    Call TableEnd

End Sub

' One pattern per wall pressure case, named after the case, holding the
' pressure at each joint of that wall (WallPressure); other joints stay 0.
Private Sub WriteJointPatternTables()

    Dim cases() As String
    Dim f() As String
    Dim c As Long, side As Long, jt As Long
    Dim sfx As String

    cases = Split(WALL_PRESSURE_CASES, ";")

    Call TableStart("JOINT PATTERN ASSIGNMENTS")
    For side = 1 To 2
        sfx = IIf(side = 1, "L", "R")
        For c = 0 To UBound(cases)
            f = Split(cases(c), "|")
            For jt = 1 To JT_COUNT
                If JT_WALL(jt) = side Then
                    Call Emit("   Joint=" & jt & "   Pattern=" & SapName(Replace(f(0), "#", sfx)) & _
                              "   Value=" & SapNum(Round(WallPressure(side, CLng(f(1)), JT_S(jt), JT_Z(jt)), 4)))
                End If
            Next jt
        Next c
    Next side
    Call TableEnd

    Call TableStart("JOINT PATTERN DEFINITIONS")
    Call Emit("   Pattern=Default")
    For side = 1 To 2
        sfx = IIf(side = 1, "L", "R")
        For c = 0 To UBound(cases)
            f = Split(cases(c), "|")
            Call Emit("   Pattern=" & SapName(Replace(f(0), "#", sfx)))
        Next c
    Next side
    Call TableEnd

End Sub

' Earth pressure at a point of a wall face from the sheet's four corner
' values - 1 bottom near, 2 bottom far, 3 top far, 4 top near - bilinear in
' u = s / L along the wall and v = z / (that column's top): the same
' 4-point area load the MIDAS builder writes to db/PNLD over the face
' (0,0) (L,0) (L,Hfar) (0,Hnear).
Private Function WallPressure(ByVal side As Long, ByVal loadRow As Long, _
                              ByVal s As Double, ByVal z As Double) As Double

    Dim q(1 To 4) As Double
    Dim wallLen As Double, hN As Double, hF As Double
    Dim u As Double, v As Double, i As Long

    For i = 1 To 4
        If side = 1 Then
            Select Case loadRow
                Case 1: q(i) = LOAD_L_ATREST(i)
                Case 2: q(i) = LOAD_L_ACTIVE(i)
                Case 3: q(i) = LOAD_L_SEISMIC(i)
            End Select
        Else
            Select Case loadRow
                Case 1: q(i) = LOAD_R_ATREST(i)
                Case 2: q(i) = LOAD_R_ACTIVE(i)
                Case 3: q(i) = LOAD_R_SEISMIC(i)
            End Select
        End If
    Next i

    If side = 1 Then
        wallLen = GEO_LEFT_L: hN = GEO_LEFT_HNEAR: hF = GEO_LEFT_HFAR
    Else
        wallLen = GEO_RIGHT_L: hN = GEO_RIGHT_HNEAR: hF = GEO_RIGHT_HFAR
    End If

    u = s / wallLen
    v = z / (hN + (hF - hN) * u)
    WallPressure = (1 - u) * (1 - v) * q(1) + u * (1 - v) * q(2) + u * v * q(3) + (1 - u) * v * q(4)

End Function

' The horizontal part of the MIDAS springs (0.5 kv x area): a joint spring
' of SPRING_H_RATIO x kv x its tributary area (a quarter of each foundation
' area it corners) in U1 and U2. The vertical kv is the area spring.
Private Sub WriteJointSpringTable()

    Dim trib() As Double
    Dim a As Long, c As Long, jt As Long
    Dim k As Double

    If SPRING_KV <= 0 Then
        TEXT_WARNINGS = TEXT_WARNINGS & " no springs"
        Exit Sub
    End If

    ReDim trib(1 To JT_COUNT)
    For a = 1 To AR_COUNT
        If AR_KIND(a) = 0 Then
            For c = 1 To 4
                trib(AR_J(a, c)) = trib(AR_J(a, c)) + AreaSize(a) / 4
            Next c
        End If
    Next a

    Call TableStart("JOINT SPRING ASSIGNMENTS 1 - UNCOUPLED")
    For jt = 1 To JT_COUNT
        If trib(jt) > 0 Then
            k = SPRING_H_RATIO * SPRING_KV * trib(jt)
            Call Emit("   Joint=" & jt & "   CoordSys=Local   U1=" & SapNum(Round(k, 3)) & "   U2=" & _
                      SapNum(Round(k, 3)) & "   U3=0   R1=0   R2=0   R3=0")
            SPRING_COUNT = SPRING_COUNT + 1
        End If
    Next jt
    Call TableEnd

End Sub

' LOAD CASE DEFINITIONS (one linear static case per pattern, plus MODAL),
' LOAD PATTERN DEFINITIONS (DL alone carries self-weight) and MASS SOURCE.
Private Sub WriteLoadCaseTables()

    Dim rows() As String
    Dim f() As String
    Dim i As Long
    Dim dt As String

    rows = Split(STLDCASE_LIST, ";")

    Call TableStart("LOAD CASE DEFINITIONS")
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
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
    Next i
    Call TableEnd

    Call TableStart("LOAD PATTERN DEFINITIONS")
    For i = 0 To UBound(rows)
        f = Split(rows(i), "|")
        Call Emit("   LoadPat=" & SapName(f(0)) & "   DesignType=" & SapDesignType(f(0), f(1)) & _
                  "   SelfWtMult=" & IIf(f(0) = "DL", "1", "0"))
    Next i
    Call TableEnd

    Call TableStart("MASS SOURCE")
    Call Emit("   MassSource=MSSSRC1   Elements=Yes   Masses=Yes   Loads=No   IsDefault=Yes")
    Call TableEnd

End Sub


' ===========================================================================
'  OPEN SCRIPT - the .$2k just written, opened in a fresh SAP2000, checked
'  (SAP2000 must hold exactly the joints and areas the mesh made - a
'  silently dropped row would show here), saved, analysed and shown.
' ===========================================================================

Private Function WriteOpenScript(ByVal stem As String) As String

    OUT_COUNT = 0
    ReDim OUT_LINES(1 To 64)

    Call EmitSapPreamble(stem & "_build.log")
    Call Emit("  SapChk 'Open the model file' ($m.File.OpenFile(" & PsQ(stem & ".$2k") & "))")
    Call Emit("  SapChk 'Units kN, m, C' ($m.SetPresentUnits(6))")
    Call Emit("  SapCount 'joints' $m.PointObj " & JT_COUNT)
    Call Emit("  SapCount 'areas' $m.AreaObj " & AR_COUNT)
    Call Emit("  SapLog 'OK|Model opened'")
    Call EmitSapFinish(stem & ".sdb")

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
