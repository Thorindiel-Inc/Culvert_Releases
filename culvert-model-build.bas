Option Explicit

' ============================================================================
'  MIDAS Civil NX API - Box Culvert Model Build
'  Replaces the three MCT-text generators in ex/ (MCT_MIDAS_INTRO,
'  MCT_MIDAS_GEO, MCT_MIDAS_LOAD) with direct db/* API writes. Same inputs,
'  same geometry, same loads as the MCT modules with ONE deliberate
'  exception: the EQ profile is now a proper top-to-bottom interpolation
'  (see EqBlock).
'  (A "64KB module cap" is folklore - CLAUDE.md records this project
'  compiling a 75KB module fine. Per-PROCEDURE size is the real limit.)
' ----------------------------------------------------------------------------
'  db/UNIT STYP MATL SECT NODE ELEM STLD BODF BMLD LCOM-GEN - all written
'  with PUT, never POST (POST dies on "Key Already Exist"; PUT is
'  create-or-update). Success = 2xx AND no "error" key in the body: this API
'  returns HTTP 200 with an error body on a bad request. See CLAUDE.md.
'
'  No boundary conditions are written - the MCT output had none either, and
'  supports are applied by hand in Civil NX after the build (confirmed).
' ============================================================================


' ---------------------------------------------------------------------------
'  CONFIG
' ---------------------------------------------------------------------------

' Stamped into the report title. The module is pasted into Excel by hand,
' so a screenshot of a run does not otherwise say which build produced it -
' bump this whenever the file changes and check it matches before
' diagnosing anything from a report.
Private Const SCRIPT_VERSION As String = "2026-09-23f"

' One-line summary of what changed in THIS version, shown by the updater
' next to this module when it's stale. Update alongside SCRIPT_VERSION -
' must stay on ONE physical line (no "_" continuation - the parser that
' reads this out does not resolve continuations) and must not contain "|"
' (breaks manifest.txt's pipe-delimited format).
Private Const SCRIPT_CHANGELOG As String = "Audit rev 2: adds ENV_EQ (seismic only), required dimensions must be present and above 0, an error value in an optional cell stops the build, the seismic gate reads this workbook and shows its source (stops if undetermined), long timeout for doc/ANAL and post/TABLE, result-table combinations derived from the generated list, verdict first in the report."

' Identifies this module to the updater regardless of what it was
' named when pasted into Excel - these files carry no VB_Name, so the
' module name in the VBA project is whatever the user typed.
Private Const SCRIPT_ID As String = "culvert-model-build"

' DESTRUCTIVE, and on by default. BuildCulvertModel deletes everything it
' is about to write before writing it, so a run always starts from a known
' empty state instead of merging into whatever the model already held.
'
' Why it is the default: PUT replaces a record BY KEY, but names are
' unique MODEL-WIDE. Writing into a model that already has these records
' therefore fails as soon as the two numberings disagree - e.g. the old
' ENV_SER sitting at key 31 while this run writes ENV_SER to key 30. That
' happens whenever the record COUNT changes, which SEISMIC_ACTIVE alone is
' enough to do (34 combinations vs 33, 20 load cases vs 18).
'
' Set False only to write into a model you want left otherwise intact -
' and expect name collisions if it already holds any of these records.
Private Const CLEAN_BEFORE_BUILD As Boolean = True

Private Const API_BASE_URL As String = "https://moa-engineers.midasit.com:443/civil"
' The MAPI key lives on the worksheet, NOT in this file. Civil NX issues a
' new key whenever you press Refresh in Tools > API > API Setting, and
' these modules are pasted into workbooks by hand - keeping the key in a
' cell means rotating it is a one-cell edit instead of re-pasting every
' module, and it keeps the secret out of the published code entirely.
Private Const MAPI_KEY_SHEET As String = "INPUT"
Private Const MAPI_KEY_CELL As String = "J20"

' Read once per run and remembered, so a long job does not re-read the
' cell on every request.
Private MAPI_KEY_CACHE As String

' One WinHTTP client reused for every request in a run. Declared here
' because ALL module-level declarations must precede the first procedure
' in a VBA module. Created on first use, released by ClearProgress.
Private HTTP_CLIENT As Object

Private Const PROJECT_FORCE_UNIT As String = "KN"
Private Const PROJECT_DIST_UNIT As String = "M"
Private Const PROJECT_HEAT_UNIT As String = "KCAL"
Private Const PROJECT_TEMPER_UNIT As String = "C"

' *STRUCTYPE "1, 1, 1, NO, YES, 9.806, 0, NO, NO, NO" - note iSTYP=1
' (X-Z plane), NOT the wingwall builder's 0 (3-D): this culvert is a plane
' frame in X-Z, which is why every node has Y=0.
Private Const STRUCTYPE_TYPE As Long = 1
Private Const STRUCTYPE_MASS As Long = 1
Private Const STRUCTYPE_SMASS As Long = 1
Private Const STRUCTYPE_GRAV As Double = 9.806
Private Const STRUCTYPE_TEMP As Double = 0

' *MATERIAL "1, CONC, C30/37, 0, 0, , C, NO, 0.05, 2, 2.6291e+07, 0.2, 1.0000e-05, 25, 0"
' DATA1 form 2 (Isotropic user-defined): ELAST, POISN, THERMAL, DEN, MASS.
' ELAST comes from MIDAS_INPUT!B43 (see MATERIAL_ELAST below), DEN from the
' "INPUT" sheet's B5 (see MATERIAL_DEN below); POISN/THERMAL/MASS are fixed,
' matching this MCT row.
Private Const MATERIAL_NO As Long = 1
Private Const MATERIAL_NAME As String = "C30/37"
Private Const MATERIAL_DAMP_RATIO As Double = 0.05
Private Const MATERIAL_POISN As Double = 0.2
Private Const MATERIAL_THERMAL As Double = 0.00001
Private Const MATERIAL_MASS As Double = 0
Private Const MATERIAL_ELAST_DEFAULT As Double = 26291000#  ' fallback if B43 is blank/invalid
Private Const MATERIAL_DEN_DEFAULT As Double = 25           ' fallback if INPUT!B5 is blank/invalid

' Every *SECTION row is "DBUSER ... SB, 2, <depth>, 1, 0 x8" - a solid
' rectangle (SHAPE "SB", DATATYPE 2 = user value) of <depth> x 1.0m, i.e.
' the usual per-metre-run strip. vSIZE is [depth, width, 0 x8].
Private Const SECTION_WIDTH As Double = 1#

' Beam-load values are rounded to 3 dp before being sent, matching the
' Round(x, 3) the MCT module applied when writing its text.
Private Const LOAD_ROUND_DP As Long = 3


' ---------------------------------------------------------------------------
'  INPUT SHEET - same cell addresses the MCT modules used, now bound to a
'  named sheet instead of whatever happened to be active.
'
'  Column A holds the label, column B the value. Sheet labels in ().
'    A4/B4   sect 1  slab thickness        (DOSEME)
'    A5/B5   sect 2  wall thickness        (DUVAR)
'    A6/B6   sect 3  - only if B6>0; defined but no element uses it
'    A7/B7   sect 4  foundation thickness  (TEMEL)
'    A10/B10 sect 5  slab haunch (DOS GUSE) - GATED ON B8, not B10
'    A11/B11 sect 6  wall haunch (DUV GUSE) - GATED ON B9, not B11
'    A12/B12 sect 7  slab rigid zone       (DOS RIJIT)
'    A13/B13 sect 8  wall rigid zone       (DUV RIJIT)
'    A16/B16 sect 9  foundation rigid zone (TEM RIJIT)
'    B8  haunch HEIGHT (GUSE,YUK)      B9  haunch WIDTH (GUSE,EN)
'        Both dimensions of the same haunch - B8 sets the length of the
'        vertical haunch members and B9 the horizontal ones. The section
'        gating crosses over (see above); harmless in practice because a
'        real haunch has both non-zero together.
'    B14 fill/cover depth (DOLGU) - read by the workbook, NOT by this
'        script or by the MCT modules it replaces
'    B15 ATA self-weight X factor (kh)
'    B17 foundation outer extension (MENFEZ DIS GENISLIGI); >0 selects
'        the 20-node branch
'    B18 clear span (MENFEZ, L)        B19 wall clear height (MENFEZ, H)
'    B21 EV1 slab   B22 EV2 slab   B37 EV1 on ext. elems   B36 EV2 on ext.
'    B23/B24 EHS1 top/bottom   B25/B26 EHA1   B27/B28 EHS2   B29/B30 EHA2
'    B31 LL1   B32 LL   B39 LLacc
'    B38 LSS1_L   B33 LSS2_L   B40 LSA2_L
'    B34 EQ intensity at the TOP (EQ_T)   B35 at the BOTTOM (EQ_B)
'  Load combinations are NOT read from the sheet - they are fixed in
'  GenerateLoadCombinations below.
'  Columns E-H are OUTPUT: the earth-pressure/seismic ordinate table, kept
'  for inspection exactly as the MCT module wrote it.
' ---------------------------------------------------------------------------
Private Const INPUT_SHEET_NAME As String = "MIDAS_INPUT"

' Project Information (db/PJCF). The project name is NOT on MIDAS_INPUT -
' it comes off the workbook's own "INPUT" sheet, assembled as
'   G9 & " " & G8 & " - " & G11
' e.g. "MNF" + "151A" + "A1-A2-A3" -> "MNF 151A - A1-A2-A3".
' A missing sheet or blank cells are not fatal: whatever parts are present
' are joined, and an entirely empty result makes the step WARN rather than
' block the build.
Private Const PROJECT_SHEET_NAME As String = "INPUT"
Private Const CELL_PROJECT_TYPE As String = "G9"
Private Const CELL_PROJECT_NO As String = "G8"
Private Const CELL_PROJECT_PART As String = "G11"
Private Const CELL_PROJECT_ENGINEER As String = "N15"
Private Const CELL_PROJECT_REVISION As String = "N16"

' Element division (ope/DIVIDEELEM)
Private Const CELL_DIVIDE_AMOUNT As String = "K15"

' Non-rigid element list (written to INPUT sheet prior to analysis)
Private Const CELL_ELEM_LIST As String = "G15"

' Results sheet name (written after analysis)
Private Const RESULT_SHEET_NAME As String = "MIDAS_RESULTS"

' Fixed, same values the wingwall builder uses.
Private Const PROJINFO_USER As String = "Deha"
Private Const PROJINFO_ADDRESS As String = "DEHA"

' Main Control Data (db/ACTL), carried over from the wingwall builder.
' These are the JSON Manual's own example defaults, NOT derived from any
' MCT block in ex/ - the MCT modules never wrote this record.
Private Const MAIN_CTRL_ARDC As Boolean = True
Private Const MAIN_CTRL_ANRC As Boolean = True
Private Const MAIN_CTRL_ITER As Long = 20
Private Const MAIN_CTRL_TOL As Double = 0.001
Private Const MAIN_CTRL_CSECF As Boolean = False
Private Const MAIN_CTRL_TRS As Boolean = True
Private Const MAIN_CTRL_CRBAR As Boolean = False
Private Const MAIN_CTRL_BMSTRESS As Boolean = False
Private Const MAIN_CTRL_CLATS As Boolean = False

' ---------------------------------------------------------------------------
'  SECTION COLOURS (db/CO_S)
'
'  Colour is NOT an attribute of db/SECT - there is no colour field on that
'  endpoint at all. The API manual lists a separate "View" group with
'  db/CO_M (material), db/CO_S (section), db/CO_T (thickness) and
'  db/CO_F (floor load). db/CO_S is the one that colours sections.
'
'  db/CO_S is GET/PUT only - no POST and, importantly, NO DELETE, which is
'  why RunCleanSteps does not try to clear it. Its Assign key is the
'  SECTION number.
'
'  THREE colour triplets per section, which is how 2D and 3D are separated:
'    W_*   WireFrame   - the 2D line view
'    HF_*  HiddenFill  - the 3D rendered surface
'    HE_*  HiddenEdge  - the 3D rendered edges
'  Plus bBLEMD (opacity on/off) and FACT (opacity value, used when on).
'
'  Only the BASE colour is listed below. It goes to the wireframe and the
'  hidden edge; the hidden fill is that colour lightened by
'  SECTION_FILL_LIGHTEN, so a 3D view reads as a lighter wash inside
'  darker edges.
'
'  That follows the SHAPE of the JSON Manual's example - it pairs W/HE
'  111,142,91 with a lighter HF 159,205,131 - but not its exact arithmetic:
'  back-solving that pair per channel gives 1.4324 / 1.4437 / 1.4396, so it
'  is NOT a uniform multiply, just MIDAS's default pair for that section.
'  1.44 is our own choice, near the middle of that range.
'
'  Haunches are brighter variants of the member they belong to and the
'  rigid zones are hot colours, so the structure reads at a glance:
'    1 DOSEME (slab)          blue        5 DOS GUSE (slab haunch) mid blue
'    2 DUVAR (wall)           green       6 DUV GUSE (wall haunch) mid green
'    3 unused                 grey        7 DOS RIJIT (slab rigid) red
'    4 TEMEL (foundation)     brown       8 DUV RIJIT (wall rigid) orange
'                                         9 TEM RIJIT (found rigid) red-brown
'  Section 3 gets an entry for completeness but no element references it,
'  and PostSectionColors walks SECT_LIST rather than this list, so a
'  section that was never written is never coloured either.
' ---------------------------------------------------------------------------
Private Const SECTION_COLOR_LIST As String = _
    "1|45,110,200;2|60,160,75;3|150,150,150;4|110,90,70;" & _
    "5|70,150,205;6|95,185,110;7|205,50,45;8|225,130,35;9|185,70,45"

' Hidden-fill tint factor. Channels round, then clamp at 255 - a bright
' base therefore shifts hue slightly as one channel saturates, which is
' fine for a fill sitting inside its own darker edge colour.
Private Const SECTION_FILL_LIGHTEN As Double = 1.44

' Opacity. False = solid, which is what a culvert wants; FACT only applies
' when bBLEMD is True.
Private Const SECTION_COLOR_TRANSLUCENT As Boolean = False
Private Const SECTION_COLOR_OPACITY As Double = 0.5

' ---------------------------------------------------------------------------
'  SEISMIC GATE - whether this culvert is designed for earthquake at all.
'
'  On the sheet this is evaluated as:
'    IF('1_GIRIS'!O81 < '1_GIRIS'!Q81, <with EQ-1>, <without>)
'  i.e. seismic counts only when the burial ratio Z/H is BELOW 0.5:
'    O81 = Z/H = INPUT!B21 / (INPUT!B23 + INPUT!B12 + INPUT!B14)
'    Q81 = 0.5 (fixed threshold)
'  Above 0.5 the sheet's own note (1_GIRIS!M83) says dynamic effects are
'  disregarded for buried box culverts, because the static case governs
'  more unfavourably than the seismic one.
'
'  SEISMIC_GATE_OVERRIDE controls evaluation:
'    -1 = Auto-detect from workbook ('1_GIRIS'!P81 / '1_GIRIS'!O81 < Q81 or INPUT ratio < 0.5)
'     0 = Force OFF (False)
'     1 = Force ON (True)
'
'  When FALSE, the WHOLE seismic chain is skipped - nothing seismic is
'  written to the model at all:
'    db/STLD    no EQ, no ATA load case
'    db/BMLD    no EQ beam loads          (AddEqLoads)
'    db/BODF    no ATA self-weight        (PostSelfWeight)
'    db/LCOM    no EQ-1, and ENV_ALL envelopes only ENV_SER + ENV_STR
'  When TRUE, the complete seismic chain is active:
'    db/STLD    EQ and ATA load cases created
'    db/BMLD    EQ lateral earth pressure beam loads applied
'    db/BODF    ATA self-weight inertia acceleration applied
'    db/LCOM    EQ-1 combo created and added to ENV_ALL envelope
'    post/TABLE PostBeamForceResults retrieves EQ-1(CB) into Table 1
'               and reflects it in Table 2 envelopes
' ---------------------------------------------------------------------------
Private Const SEISMIC_GATE_OVERRIDE As Long = -1
Private SEISMIC_ACTIVE As Boolean
' Where the gate's value came from, for the build report (e.g.
' "1_GIRIS!O81 = 0.769 >= 0.5"), and whether it could be determined at
' all. Auto-detect that finds neither source stops the build instead of
' silently dropping the whole seismic chain.
Private SEISMIC_GATE_SOURCE As String
Private SEISMIC_GATE_KNOWN As Boolean

' ---------------------------------------------------------------------------
'  PROGRESS BAR
'
'  Drawn in Excel's status bar rather than a UserForm, so this module stays
'  a single file that can be pasted into a workbook - a form would need a
'  separate .frm plus its binary .frx.
'
'  The step counts below are advisory: Progress clamps at 100%, so adding a
'  step without bumping them makes the bar finish slightly early rather
'  than misbehave. Bump them anyway when you add or remove a step.
' ---------------------------------------------------------------------------
Private Const PROGRESS_STEPS_BUILD As Long = 19
Private Const PROGRESS_STEPS_CLEAN As Long = 6
Private Const PROGRESS_STEPS_VERIFY As Long = 2
Private Const PROGRESS_BAR_WIDTH As Long = 24

Private PROGRESS_STEP As Long
Private PROGRESS_TOTAL As Long

' Geometry + section inputs
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
Private DIVIDE_AMOUNT As Long        ' INPUT!K15 (subdivision count for elements 18, 10, 11, 4)
Private SPRING_KZ_MODULUS As Double  ' MIDAS_INPUT!B42 (subgrade modulus, ks)
Private MATERIAL_ELAST As Double     ' MIDAS_INPUT!B43 (Ec), falls back to MATERIAL_ELAST_DEFAULT
Private MATERIAL_DEN As Double       ' "INPUT"!B5 (unit weight), falls back to MATERIAL_DEN_DEFAULT
' Collects "fell back to a default" notes from ReadInputs so they surface
' as a WARN on the build report instead of passing silently. A wrong unit
' weight or Ec changes every self-weight load without any other sign.
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
Private LD_EQ_TOP As Double, LD_EQ_BOT As Double   ' B34 (EQ_T), B35 (EQ_B)

' Project name, read from the INPUT sheet by ReadProjectName
Private PROJECT_NAME As String
' Engineer name, read from INPUT!N15 by ReadProjectName
Private PROJECT_ENGINEER As String
' Revision info, read from INPUT!N16 by ReadProjectName
Private PROJECT_REVISION As String

' Computed ordinate tables (the 8 values each block writes to cols E-H)
Private ORD_EHS1(1 To 8) As Double
Private ORD_EHA1(1 To 8) As Double
Private ORD_EHS2(1 To 8) As Double
Private ORD_EHA2(1 To 8) As Double
Private ORD_EQ(1 To 8) As Double

' Generated model data
Private NODE_LIST As String       ' "no|x|y|z;..."
Private ELEM_LIST As String       ' "no|sect|n1|n2|angle;..."
Private SECT_LIST As String       ' "no|name|depth;..."
Private BEAMLOAD_LIST As String   ' "elem|lcname|cmd|dir|p1|p2;..."
Private LOADCOMB_LIST As String   ' "name|ACTIVE|iTYPE|ANAL:LC:F,...|tier;..."

' Element numbering bases, set by GenerateGeometry:
Private HAS_EXT As Boolean
Private SLAB_E As Long            ' first of the 5 top-slab elements (always 16)
Private WALL_E As Long            ' first of the 8 wall elements (always 8)

' Wall elements always start at WALL_E = 8 and top-slab elements at SLAB_E = 16
' in both extension and no-extension models.

' The 20 static load cases, in the MCT module's own order. The last two
' (EQ, ATA) are the seismic pair and are dropped when SEISMIC_ACTIVE is
' False - see IsSeismicCase.
Private Const STLDCASE_LIST As String = _
    "DL|D;EV1|EV;EV2|EV;EVin|EV;EHS1|EH;EHA1|EH;EHS2_L|EH;EHA2_L|EH;" & _
    "EHS2_R|EH;EHA2_R|EH;LL1|L;LL|L;LLin|L;LLacc|E;LSS1_L|LS;LSS2_L|LS;" & _
    "LSA2_L|LS;WA|FP;EQ|E;ATA|E"


' ===========================================================================
'  SEISMIC GATE EVALUATION
'  Moved below the declarations block 2026-09-21 - VBA requires every
'  module-level Const/Dim/Private declaration to sit before the module's
'  FIRST procedure. These two used to sit between DetectSeismicGate's own
'  comment block and PROGRESS BAR's declarations, which put declarations
'  AFTER a procedure and broke compilation ("Only comments may appear
'  after End Sub, End Function, or End Property") - confirmed live. See
'  CLAUDE.md's VBA gotchas section.
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

' Sets SEISMIC_GATE_SOURCE, and SEISMIC_GATE_KNOWN = False when neither
' 1_GIRIS nor the INPUT cells could be evaluated. Reads THIS workbook, not
' the active one.
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


' ===========================================================================
'  ENTRY POINTS
' ===========================================================================

Sub BuildCulvertModel()

    Dim report As String
    Dim cleanReport As String
    Dim ok As Boolean
    Dim vLcom As String, vStld As String

    Call InitSeismicGate

    report = "Box culvert model build  [" & SCRIPT_VERSION & "]" & vbCrLf & _
             String(40, "-") & vbCrLf

    ' Stop before anything is deleted or written: silently building without
    ' the seismic chain is the unsafe default.
    If Not SEISMIC_GATE_KNOWN Then
        MsgBox report & "STOPPED - nothing was changed in the model." & vbCrLf & vbCrLf & _
               "The seismic gate " & SEISMIC_GATE_SOURCE & "." & vbCrLf & vbCrLf & _
               "Fix those cells, or set SEISMIC_GATE_OVERRIDE to 0 (off) or 1 (on) " & _
               "at the top of this module.", vbExclamation
        Exit Sub
    End If

    If SEISMIC_ACTIVE Then
        report = report & "NOTE - seismic gate is ON (" & SEISMIC_GATE_SOURCE & "): EQ/ATA " & _
                 "cases and loads, EQ-1 and ENV_EQ active." & vbCrLf & String(40, "-") & vbCrLf
    Else
        report = report & "NOTE - seismic gate is OFF (" & SEISMIC_GATE_SOURCE & "): no EQ/ATA " & _
                 "cases or loads." & vbCrLf & String(40, "-") & vbCrLf
    End If

    ' Wipe first, so every PUT below lands in an empty endpoint and the
    ' key/name numbering cannot disagree with whatever was there before.
    ' Deliberately NOT gated on success: a FAIL here usually just means
    ' the endpoint was already empty, which is the state we wanted.
    Call ResetProgress(PROGRESS_STEPS_BUILD + _
         IIf(CLEAN_BEFORE_BUILD, PROGRESS_STEPS_CLEAN + PROGRESS_STEPS_VERIFY, 0))

    If CLEAN_BEFORE_BUILD Then
        cleanReport = ""
        Call RunCleanSteps(cleanReport)
        vLcom = VerifyEmpty("db/LCOM-GEN", "Load combinations")
        vStld = VerifyEmpty("db/STLD", "Static load cases")
        Call StepResult(cleanReport, Progress("verify Load Combinations cleared"), vLcom)
        Call StepResult(cleanReport, Progress("verify Static Load Cases cleared"), vStld)

        ' Keep the build report concise so it does not exceed VBA's ~1024-char MsgBox limit.
        ' If clean steps were all OK, summarise in one line; if any warned/failed, show in full.
        If InStr(cleanReport, "FAIL") > 0 Or InStr(cleanReport, "WARN") > 0 Then
            report = report & "Clean pass (CLEAN_BEFORE_BUILD):" & vbCrLf & cleanReport & String(40, "-") & vbCrLf
        Else
            report = report & "Clean pass (CLEAN_BEFORE_BUILD): OK" & vbCrLf & String(40, "-") & vbCrLf
        End If
    End If


    ok = StepResult(report, Progress("Unit System"), PostUnitSystem())
    If ok Then ok = StepResult(report, Progress("Structure Type"), PostStructureType())
    If ok Then ok = StepResult(report, Progress("Read inputs + generate geometry"), GeometrySetup())
    ' After GeometrySetup - that is where the project name is read.
    If ok Then ok = StepResult(report, Progress("Project Information"), PostProjectInfo())
    If ok Then ok = StepResult(report, Progress("Main Control Data"), PostMainControlData())
    If ok Then ok = StepResult(report, Progress("Material"), PostMaterial())
    If ok Then ok = StepResult(report, Progress("Sections"), PostSections())
    If ok Then ok = StepResult(report, Progress("Section Colours"), PostSectionColors())
    If ok Then ok = StepResult(report, Progress("Nodes"), PostNodes())
    If ok Then ok = StepResult(report, Progress("Elements"), PostElements())
    If ok Then ok = StepResult(report, Progress("Static Load Cases"), PostStaticLoadCases())
    If ok Then ok = StepResult(report, Progress("Self-Weight"), PostSelfWeight())
    If ok Then ok = StepResult(report, Progress("Beam Loads"), PostBeamLoads())
    If ok Then ok = StepResult(report, Progress("Load Combinations"), PostLoadCombinations())
    If ok Then ok = StepResult(report, Progress("Divide Elements"), PostDivideElements())
    If ok Then ok = StepResult(report, Progress("Foundation Springs"), PostFoundationSprings())
    If ok Then ok = StepResult(report, Progress("Element List"), PostNonRigidElementList())
    If ok Then ok = StepResult(report, Progress("Perform Analysis"), PostPerformAnalysis())
    If ok Then ok = StepResult(report, Progress("Beam Force Results"), PostBeamForceResults())

    report = VerdictFirst(report, ok)

    ' Before the MsgBox, and on the failure path too - see ClearProgress.
    Call ClearProgress

    MsgBox FitReport(report, SCRIPT_ID), IIf(ok, vbInformation, vbExclamation)

End Sub

' Reverse of the build order, so a re-run starts from a clean model instead
' of colliding on "Key Already Exist"/"Duplicate Name". A FAIL on an
' already-empty endpoint is expected and harmless.
Sub CleanCulvertModel()

    Dim report As String
    report = "Box culvert model clean  [" & SCRIPT_VERSION & "]" & vbCrLf & _
             String(40, "-") & vbCrLf

    Call ResetProgress(PROGRESS_STEPS_CLEAN)
    Call RunCleanSteps(report)
    Call ClearProgress

    report = report & vbCrLf & "Clean pass done - review any unexpected FAILs " & _
             "before running BuildCulvertModel()."

    MsgBox report, vbInformation

End Sub

' The delete sequence, shared by CleanCulvertModel and by the
' CLEAN_BEFORE_BUILD pass at the top of BuildCulvertModel.
'
' Order is the REVERSE of the build order - loads before the load cases
' and elements they reference, elements before nodes, SECT/MATL last since
' ELEM assigns them by number.
'
' No step gates on another: a FAIL on an already-empty endpoint is
' expected and harmless, so every delete is attempted regardless and the
' report is there to review.
'
' db/CO_S is deliberately absent: it lists GET and PUT as its only active
' methods, so there is nothing to delete. Section colours are simply
' overwritten by the next PostSectionColors.
Private Sub RunCleanSteps(ByRef report As String)

    Call StepResult(report, Progress("delete Load Combinations"), DeleteAndCheck("db/LCOM-GEN"))
    Call StepResult(report, Progress("delete Static Load Cases"), DeleteAndCheck("db/STLD"))
    Call StepResult(report, Progress("delete Elements"), DeleteAndCheck("db/ELEM"))
    Call StepResult(report, Progress("delete Nodes"), DeleteAndCheck("db/NODE"))
    Call StepResult(report, Progress("delete Sections"), DeleteAndCheck("db/SECT"))
    Call StepResult(report, Progress("delete Material"), DeleteAndCheck("db/MATL"))

End Sub


' ===========================================================================
'  STEP PLUMBING (shared with the other build script - see CLAUDE.md)
' ===========================================================================

' ---------------------------------------------------------------------------
'  PROGRESS
' ---------------------------------------------------------------------------

Private Sub ResetProgress(ByVal total As Long)
    PROGRESS_STEP = 0
    ' Re-read INPUT!J20 on every run, so a key rotated since the last run
    ' is actually used.
    MAPI_KEY_CACHE = ""
    PROGRESS_TOTAL = IIf(total < 1, 1, total)
    ' Paired with ClearProgress, which already runs on every exit path.
    ' ScreenUpdating only: EnableEvents and Calculation are deliberately NOT
    ' touched here. This Sub has no error handler, and VBA restores
    ' ScreenUpdating by itself when execution halts, whereas EnableEvents
    ' stays off and would leave the workbook silently ignoring events. The
    ' builders spend their time in MIDAS and HTTP, not repainting, so the
    ' extra risk buys almost nothing.
    Application.ScreenUpdating = False
End Sub

' Hands the status bar back to Excel. Must run on EVERY exit path - a
' status bar left set stays stuck there for the rest of the session.
Private Sub ClearProgress()
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = True
    ' Release the shared WinHTTP client at the end of a run (this Sub runs
    ' on every exit path); the next run builds a fresh one.
    Set HTTP_CLIENT = Nothing
    On Error GoTo 0
End Sub

' Advances the bar and RETURNS ITS OWN ARGUMENT, so it drops straight into
' the existing call without an extra line per step:
'
'     StepResult(report, Progress("Sections"), PostSections())
'
' VBA evaluates arguments left to right, so the bar names the step that is
' about to run rather than the one that just finished - which matters
' because the HTTP call is the slow part. If that order ever changed the
' label would simply lag by one step; it cannot affect what gets written.
Private Function Progress(ByVal stepName As String) As String

    Dim pct As Long
    Dim filled As Long

    Progress = stepName

    PROGRESS_STEP = PROGRESS_STEP + 1
    If PROGRESS_TOTAL < 1 Then PROGRESS_TOTAL = 1

    pct = CLng((PROGRESS_STEP - 1) * 100 / PROGRESS_TOTAL)
    If pct < 0 Then pct = 0
    If pct > 100 Then pct = 100

    filled = CLng(pct * PROGRESS_BAR_WIDTH / 100)
    If filled > PROGRESS_BAR_WIDTH Then filled = PROGRESS_BAR_WIDTH

    ' ChrW keeps this file ASCII: 9608 is a full block, 9617 a light shade.
    On Error Resume Next
    Application.StatusBar = "MIDAS culvert  [" & _
        String(filled, ChrW(9608)) & String(PROGRESS_BAR_WIDTH - filled, ChrW(9617)) & _
        "] " & pct & "%   (" & PROGRESS_STEP & "/" & PROGRESS_TOTAL & ")  " & stepName
    On Error GoTo 0

    ' Lets Excel actually repaint the bar between synchronous HTTP calls.
    DoEvents

End Function


Private Function DeleteAndCheck(ByVal path As String) As String
    DeleteAndCheck = PostAndCheck(path, "{}", "DELETE")
End Function

' GETs an endpoint after a delete and reports whether anything survived.
' Crude on purpose: a record of any of these types carries a "NAME", so
' its presence in the response means the endpoint is not empty. Returns
' "" (treated as OK) when it looks clear.
Private Function VerifyEmpty(ByVal path As String, ByVal what As String) As String

    Dim resp As String
    Dim statusCode As Long

    Call SendApiRequest("GET", path, "", resp, statusCode)

    If InStr(1, resp, """NAME""", vbTextCompare) > 0 Then
        VerifyEmpty = "WARN: " & what & " STILL HOLD RECORDS after the delete - the " & _
            "keyless DELETE did not clear this endpoint. Clear it by hand in Civil NX " & _
            "(or delete the records there) and re-run; otherwise the writes below will " & _
            "collide on duplicate names."
    End If

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

' db/STLD and db/LCOM-GEN carry model-wide-unique NAMEs, so a re-run hits a
' name collision (HTTP 200 with an error body) that PUT does not solve.
' Downgraded to a warning: loads reference their case BY NAME, so existing
' names still leave the build valid. Does NOT verify their contents match.
Private Function TolerateDuplicateName(ByVal result As String, ByVal what As String) As String

    If InStr(1, result, "Duplicate Name", vbTextCompare) > 0 _
        Or InStr(1, result, "already exists", vbTextCompare) > 0 Then
        TolerateDuplicateName = "WARN: " & what & " already exist in this model - kept " & _
            "the existing ones. Verify their types/contents if this model wasn't empty."
    Else
        TolerateDuplicateName = result
    End If

End Function


' ===========================================================================
'  PROJECT SETUP
' ===========================================================================

Private Function PostUnitSystem() As String

    Dim b As String
    b = "{""Assign"": {""1"": {"
    b = b & """FORCE"": """ & PROJECT_FORCE_UNIT & ""","
    b = b & """DIST"": """ & PROJECT_DIST_UNIT & ""","
    b = b & """HEAT"": """ & PROJECT_HEAT_UNIT & ""","
    b = b & """TEMPER"": """ & PROJECT_TEMPER_UNIT & """"
    b = b & "}}}"

    PostUnitSystem = PostAndCheck("db/UNIT", b, "PUT")

End Function


Private Function PostStructureType() As String

    Dim b As String
    b = "{""Assign"": {""1"": {"
    b = b & """STYP"": " & STRUCTYPE_TYPE & ","
    b = b & """MASS"": " & STRUCTYPE_MASS & ","
    b = b & """bMASSOFFSET"": false,"
    b = b & """bSELFWEIGHT"": true,"
    b = b & """SMASS"": " & STRUCTYPE_SMASS & ","
    b = b & """GRAV"": " & JsonNum(STRUCTYPE_GRAV) & ","
    b = b & """TEMP"": " & JsonNum(STRUCTYPE_TEMP) & ","
    b = b & """bALIGNBEAM"": false,"
    b = b & """bALIGNSLAB"": false,"
    b = b & """bROTRIGID"": false"
    b = b & "}}}"

    PostStructureType = PostAndCheck("db/STYP", b, "PUT")

End Function


' db/PJCF. PUT, not POST: Civil NX pre-populates a Project Information
' record (key "1") in every model, so POST-to-create collides with it
' ("Key Already Exist", confirmed live on the wingwall builder).
Private Function PostProjectInfo() As String

    Dim b As String

    If Len(PROJECT_NAME) = 0 Then
        PostProjectInfo = "WARN: no project name found at " & PROJECT_SHEET_NAME & "!" & _
            CELL_PROJECT_TYPE & "/" & CELL_PROJECT_NO & "/" & CELL_PROJECT_PART & _
            " - Project Information not written."
        Exit Function
    End If

    b = "{""Assign"": {""1"": {"
    b = b & """PROJECT"": """ & JsonEsc(PROJECT_NAME) & ""","
    b = b & """USER"": """ & JsonEsc(PROJINFO_USER) & ""","
    b = b & """ADDRESS"": """ & JsonEsc(PROJINFO_ADDRESS) & ""","
    b = b & """TITLE"": """ & JsonEsc(PROJECT_NAME) & """"
    If Len(PROJECT_ENGINEER) > 0 Then
        b = b & ","
        b = b & """ENGINEER"": """ & JsonEsc(PROJECT_ENGINEER) & ""","
        b = b & """EDATE"": """ & Format$(Date, "yyyy.mm.dd") & """"
    End If
    If Len(PROJECT_REVISION) > 0 Then
        b = b & ","
        b = b & """REVISION"": """ & JsonEsc(PROJECT_REVISION) & """"
    End If
    b = b & "}}}"

    PostProjectInfo = PostAndCheck("db/PJCF", b, "PUT")

End Function


' db/ACTL. PUT, not POST - same "pre-populated singleton record" reasoning
' as db/PJCF above.
Private Function PostMainControlData() As String

    Dim b As String
    b = "{""Assign"": {""1"": {"
    b = b & """ARDC"": " & LCase(MAIN_CTRL_ARDC) & ","
    b = b & """ANRC"": " & LCase(MAIN_CTRL_ANRC) & ","
    b = b & """ITER"": " & MAIN_CTRL_ITER & ","
    b = b & """TOL"": " & JsonNum(MAIN_CTRL_TOL) & ","
    b = b & """CSECF"": " & LCase(MAIN_CTRL_CSECF) & ","
    b = b & """TRS"": " & LCase(MAIN_CTRL_TRS) & ","
    b = b & """CRBAR"": " & LCase(MAIN_CTRL_CRBAR) & ","
    b = b & """BMSTRESS"": " & LCase(MAIN_CTRL_BMSTRESS) & ","
    b = b & """CLATS"": " & LCase(MAIN_CTRL_CLATS)
    b = b & "}}}"

    PostMainControlData = PostAndCheck("db/ACTL", b, "PUT")

End Function


Private Function PostMaterial() As String

    Dim b As String
    b = "{""Assign"": {""" & MATERIAL_NO & """: {"
    b = b & """TYPE"": ""CONC"","
    b = b & """NAME"": """ & MATERIAL_NAME & ""","
    b = b & """HE_SPEC"": 0,"
    b = b & """HE_COND"": 0,"
    b = b & """PLMT"": 0,"
    b = b & """P_NAME"": """","
    b = b & """bMASS_DENS"": false,"
    b = b & """DAMP_RAT"": " & JsonNum(MATERIAL_DAMP_RATIO) & ","
    b = b & """PARAM"": [{"
    b = b & """P_TYPE"": 2,"
    b = b & """ELAST"": " & JsonNum(MATERIAL_ELAST) & ","
    b = b & """POISN"": " & JsonNum(MATERIAL_POISN) & ","
    b = b & """THERMAL"": " & JsonNum(MATERIAL_THERMAL) & ","
    b = b & """DEN"": " & JsonNum(MATERIAL_DEN) & ","
    b = b & """MASS"": " & JsonNum(MATERIAL_MASS)
    b = b & "}]"
    b = b & "}}}"

    PostMaterial = PostAndCheck("db/MATL", b, "PUT")

End Function


' ===========================================================================
'  INPUT + GEOMETRY
' ===========================================================================

Private Function GeometrySetup() As String

    Dim msg As String

    msg = ReadInputs()
    If Len(msg) > 0 Then
        GeometrySetup = msg
        Exit Function
    End If

    Call ReadProjectName
    Call GenerateGeometry
    Call ComputeOrdinates
    Call WriteOrdinateTable
    Call GenerateBeamLoads
    Call GenerateLoadCombinations

    If Len(INPUT_FALLBACKS) > 0 Then
        GeometrySetup = "WARN: used built-in defaults -" & INPUT_FALLBACKS
    Else
        GeometrySetup = ""
    End If

End Function


' Project name = INPUT!G9 & " " & INPUT!G8 & " - " & INPUT!G11,
' e.g. "MNF 151A - A1-A2-A3". Left blank rather than raised as an error if
' the sheet or the cells are missing - PostProjectInfo turns that into a
' WARN so a naming problem never blocks the model build. Each separator is
' only inserted when there is something on both sides of it, so a blank
' cell does not leave a dangling " - " or a leading space.
' Also reads the engineer's name off INPUT!N15 and revision info off
' INPUT!N16 - unlike the project name, blank cells there are not warned
' about, since ENGINEER/REVISION are both Optional per the db/PJCF doc
' and not every build has a reviewer or revision tag yet.
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

' "" when addr holds a number > 0, otherwise a message naming the cell.
' TryReadCell accepts a blank cell as 0 - right for an optional feature
' switch, wrong for a thickness or a span, which would build a degenerate
' model with no error from MIDAS.
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

' The message for an optional cell holding text or an error value. A blank
' optional cell is fine (it means 0 / feature off); #VALUE!, #REF! or text
' almost always means a broken formula, so the build stops and names it.
Private Function BadOptionalCell(ByVal addr As String) As String
    BadOptionalCell = INPUT_SHEET_NAME & "!" & addr & " holds text or an error value. " & _
                      "Leave it blank for 0, or fix the formula."
End Function


' Builds NODE_LIST and ELEM_LIST. Two branches:
'
'  DIM_EXT > 0 ("dis uzantili", foundation extends past the walls):
'    20 nodes, 20 elements. Foundation 1-7, walls 8-15, top slab 16-20.
'  DIM_EXT = 0 ("dis uzantisiz", pure rectangle):
'    16 nodes, 16 elements. Foundation 3-5, walls 8-15, top slab 16-20.
'    Foundation is flush at wall centrelines dx1 and dx2 (elements 1, 2, 6, 7 omitted).
'    Element numbering strictly follows the extension=true case.
'
'  All nodes have Y=0 (plane frame in X-Z). Foundation centreline is Z=0,
'  so its top face is at Z = foundThickness/2. Element ANGLE -180 on the
'  left-hand members reproduces the MCT's beta angles.
'
'  TWO DELIBERATE DEVIATIONS from the original MCT_MIDAS_GEO.bas (old/),
'  both confirmed 2026-09-21 - "no physics change" no longer holds for
'  this Sub, on purpose. Neither is caught by comparing against old/
'  directly any more; see tests/verify_geometry_loads.py's own header.
'   1. DIM_EXT=0 foundation is shorter by one wall thickness. The
'      original ran the foundation from outer wall face to outer wall
'      face (2 extra stub elements past each wall centreline); this
'      version stops at the wall centrelines (dx1..dx2), matching the
'      "Foundation is flush at wall centrelines" comment above.
'   2. Elements 8 and 9 (both branches) - the two wall members nearest
'      the foundation - now use SECT 8 (DUV RIJIT / rigid) instead of
'      SECT 2 (DUVAR / thin wall). This adds a rigid zone at the base of
'      each wall, mirroring the rigid zone that already existed at the
'      top (elements 14/15, also SECT 8) where the wall meets the slab.
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


' ===========================================================================
'  EARTH-PRESSURE / SEISMIC ORDINATES
'  Each block yields 8 values, one pair per wall member. Odd entries
'  repeat the previous even entry - that is how one member's end pressure
'  is chained into the next member's start, leaving no step in the profile.
'  The four EH blocks run BOTTOM-UP; the EQ block runs TOP-DOWN, because
'  AddWallProfile and AddEqLoads walk the wall in opposite directions.
' ===========================================================================

Private Sub ComputeOrdinates()

    Call EhBlock(LD_EHS1_TOP, LD_EHS1_BOT, ORD_EHS1)
    Call EhBlock(LD_EHA1_TOP, LD_EHA1_BOT, ORD_EHA1)
    Call EhBlock(LD_EHS2_TOP, LD_EHS2_BOT, ORD_EHS2)
    Call EhBlock(LD_EHA2_TOP, LD_EHA2_BOT, ORD_EHA2)
    Call EqBlock(ORD_EQ)

End Sub

' Linear earth-pressure profile. "l" is the distance from the wall top up
' to where the linear profile would reach zero, back-figured from the top
' and bottom intensities, then each member's end pressure is read off it.
' Division by zero if pTop = pBot (a uniform profile) - same as the MCT.
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

' Seismic (EQ) profile: a straight linear interpolation up the wall
' between the two sheet values - EQ_B (B35) at the bottom and EQ_T (B34)
' at the top. Equal values give a uniform block, unequal ones a trapezoid.
'
' Unlike the EH blocks this array runs TOP-DOWN, because AddEqLoads walks
' the left wall from element 14 down to element 8. The five sampling
' heights are the wall members' shared node levels:
'   ft/2 + H + ts/2   slab centreline        o(1)  elem 14 i-end
'   ft/2 + H          wall top               o(2)  14 j-end / 12 i-end
'   ft/2 + H - hg     haunch node            o(4)  12 j-end / 10 i-end
'   ft/2              foundation top face    o(6)  10 j-end / 8  i-end
'   0                 foundation centreline  o(8)  8  j-end
'
' DELIBERATE PHYSICS CHANGE (2026-09-21, by request) - the one place this
' script departs from the MCT module. The old form spliced B34 in at the
' very top and scaled everything below it off B35/(ft/2), a ramp that
' rose with height but never actually reached B34, so B34 and the rest of
' the profile were on unrelated scales.
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

' EQ intensity at height z: EQ_BOT at z = 0, EQ_TOP at z = zTop. Falls
' back to EQ_BOT for a zero-height wall rather than dividing by zero.
Private Function EqAt(ByVal z As Double, ByVal zTop As Double) As Double

    If zTop = 0 Then
        EqAt = LD_EQ_BOT
    Else
        EqAt = LD_EQ_BOT + (LD_EQ_TOP - LD_EQ_BOT) * z / zTop
    End If

End Function


' Reproduces the MCT module's scratch table in columns E-H so the computed
' pressures stay inspectable on the sheet. Purely informational - the API
' build reads the ORD_* arrays, not these cells.
'   E/F rows  3-10 EHS1, 12-19 EHA1, 21-28 EHS2, 30-37 EHA2 (sep. rows
'   11/20/29/38);  G/H rows 3-10 EQ (sep. row 11).
' The labels' element numbers are the MCT's own sequential ones and do NOT
' all match the elements the loads actually land on - cosmetic, kept as-is.
Private Sub WriteOrdinateTable()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(INPUT_SHEET_NAME)

    ws.Range("E3:H500").ClearContents

    Call WriteOrdBlock(ws, 3, 5, "EHS1", WALL_E, ORD_EHS1)
    Call WriteOrdBlock(ws, 12, 5, "EHA1", WALL_E, ORD_EHA1)
    ' EHS2's labels carry on from EHA1's counter rather than resetting -
    ' hence the +6 offset. Cosmetic, matches the MCT.
    Call WriteOrdBlock(ws, 21, 5, "EHS2", WALL_E + 6, ORD_EHS2)
    Call WriteOrdBlock(ws, 30, 5, "EHA2", WALL_E, ORD_EHA2)

    ' The EQ ordinates are only written when they are actually applied -
    ' a filled-in table under an inactive seismic gate would read as if
    ' the loads had gone into the model.
    If SEISMIC_ACTIVE Then
        Call WriteOrdBlock(ws, 3, 7, "EQ", WALL_E, ORD_EQ)
    Else
        ws.Cells(3, 7).Value = "EQ not applied"
        ws.Cells(4, 7).Value = "SEISMIC_ACTIVE = False"
    End If

End Sub

Private Sub WriteOrdBlock(ByVal ws As Worksheet, ByVal firstRow As Long, _
                          ByVal labelCol As Long, ByVal tag As String, _
                          ByVal baseElem As Long, ByRef o() As Double)

    Dim p As Long, r As Long

    For p = 1 To 4
        r = firstRow + (p - 1) * 2
        ws.Cells(r, labelCol).Value = tag & "-" & (baseElem + p - 1) & "-0"
        ws.Cells(r, labelCol + 1).Value = o(p * 2 - 1)
        ws.Cells(r + 1, labelCol).Value = tag & "-" & (baseElem + p - 1) & "-1"
        ws.Cells(r + 1, labelCol + 1).Value = o(p * 2)
    Next p

    ws.Cells(firstRow + 8, labelCol).Value = "-------------"
    ws.Cells(firstRow + 8, labelCol + 1).Value = "-------------"

End Sub


' ===========================================================================
'  BEAM LOADS
'  Every load the MCT emitted is a UNILOAD with D = [0, 1, 0, 0], so only
'  the two end intensities vary. Collected into BEAMLOAD_LIST and grouped
'  by element in PostBeamLoads - db/BMLD keys on the ELEMENT number and its
'  ITEMS array is that element's whole load list, so all of an element's
'  load cases must go in one entry or the last one would wipe the rest.
' ===========================================================================

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

' The 8 wall members: WALL_E+0,2,4,6 are the left wall bottom-to-top and
' WALL_E+1,3,5,7 the right. Both sides read the SAME ordinate pair, and
' the pair is applied J-end first (P1 = o(2p), P2 = o(2p-1)) - that is the
' MCT module's own ordering, opposite to the EQ block below.
Private Sub AddWallProfile(ByVal lcLeft As String, ByVal lcRight As String, ByRef o() As Double)

    Dim p As Long

    For p = 1 To 4
        Call AddLoad(WALL_E + (p - 1) * 2, lcLeft, "LINE", "LZ", -o(p * 2), -o(p * 2 - 1))
    Next p

    For p = 1 To 4
        Call AddLoad(WALL_E + (p - 1) * 2 + 1, lcRight, "LINE", "LZ", -o(p * 2), -o(p * 2 - 1))
    Next p

End Sub

' EQ walks the 4 left-wall elements from top to bottom:
'   WALL_E + 6 (top), WALL_E + 4 (haunch), WALL_E + 2 (lower), WALL_E + 0 (bottom).
' WALL_E is 8 in both branches, matching all other wall loads.
Private Sub AddEqLoads()

    Dim p As Long

    For p = 1 To 4
        Call AddLoad(WALL_E + 12 - (4 + p * 2), "EQ", "LINE", "LZ", _
                     -ORD_EQ(p * 2 - 1), -ORD_EQ(p * 2))
    Next p

End Sub

' loadDir, not "dir": Dir is a VBA intrinsic, and naming a parameter after
' one makes the whole procedure fail to compile while the error points at
' the CALL SITE ("Sub or Function not defined"). See CLAUDE.md.
Private Sub AddLoad(ByVal elemNo As Long, ByVal lcname As String, ByVal cmd As String, _
                    ByVal loadDir As String, ByVal p1 As Double, ByVal p2 As Double)

    If Len(BEAMLOAD_LIST) > 0 Then BEAMLOAD_LIST = BEAMLOAD_LIST & ";"

    BEAMLOAD_LIST = BEAMLOAD_LIST & elemNo & "|" & lcname & "|" & cmd & "|" & loadDir & _
                    "|" & JsonNum(Round(p1, LOAD_ROUND_DP)) & _
                    "|" & JsonNum(Round(p2, LOAD_ROUND_DP))

End Sub


' ===========================================================================
'  LOAD COMBINATIONS - fixed set, transcribed from the sheet's old
'  *LOADCOMB block. Nothing is read from the worksheet any more.
'
'  35 combinations (33 when SEISMIC_ACTIVE is False - no EQ-1, no ENV_EQ):
'    SLS-1..14   serviceability, factors 1.0 / 0.5
'    ULS-1..14   ultimate, 1.35 permanent / 1.5 / 1.45 / 0.75
'    ACC-1       accidental (LLacc) - NOT seismic, always built
'    EQ-1        seismic (EQ lateral pressure only, no ATA inertia) -
'                only when SEISMIC_ACTIVE
'    ENV_SER     envelope of all SLS
'    ENV_STR     envelope of all ULS
'    ENV_ALL     envelope of ENV_SER + ENV_STR, plus EQ-1 when
'                SEISMIC_ACTIVE (see the gate at the top of the module)
'    ENV_DEAD    envelope of the four no-live-load SLS cases
'    ENV_EQ      envelope of EQ-1 - only when SEISMIC_ACTIVE
'  The SLS/ULS/ACC/EQ combinations reference static load cases ("ST");
'  every ENV_* one references other combinations ("CB"), which is why
'  each combination's factor list is homogeneous and AddCombo can take a
'  single analysis type for the whole list.
'
'  ENV_* carry iTYPE 1 (Envelope); everything else iTYPE 0 (Add).
'  Order matters: ENV_ALL references ENV_SER/ENV_STR, so they are defined
'  first and the Assign keys come out in that order.
' ===========================================================================

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

' Appends one combination to LOADCOMB_LIST. "spec" is a comma-separated
' list of "<load case>:<factor>" pairs; anal ("ST" or "CB") applies to
' every pair in it, since no combination here mixes the two.
'
' "tier" is the dependency level, used by PostLoadCombinations to decide
' what has to be written before what:
'   1  references only static load cases           (SLS/ULS/ACC/EQ-1)
'   2  references tier-1 combinations              (ENV_SER/STR/DEAD)
'   3  references tier-2 combinations              (ENV_ALL)
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



' ===========================================================================
'  API WRITES
' ===========================================================================

' db/SECT, one record per *SECTION row. SHAPE "SB" + DATATYPE 2 is a
' user-defined solid rectangle; vSIZE is [depth, width, 0 x8].
Private Function PostSections() As String

    Dim rows() As String
    Dim f() As String
    Dim i As Long
    Dim b As String

    rows = Split(SECT_LIST, ";")
    b = "{""Assign"": {"

    For i = LBound(rows) To UBound(rows)
        f = Split(rows(i), "|")
        If i > LBound(rows) Then b = b & ","
        b = b & """" & f(0) & """: {"
        b = b & """SECTTYPE"": ""DBUSER"","
        b = b & """SECT_NAME"": """ & f(1) & ""","
        b = b & """SECT_BEFORE"": {"
        b = b & """OFFSET_PT"": ""CC"","
        b = b & """OFFSET_CENTER"": 0,"
        b = b & """USER_OFFSET_REF"": 0,"
        b = b & """HORZ_OFFSET_OPT"": 0,"
        b = b & """USERDEF_OFFSET_YI"": 0,"
        b = b & """VERT_OFFSET_OPT"": 0,"
        b = b & """USERDEF_OFFSET_ZI"": 0,"
        b = b & """USE_SHEAR_DEFORM"": true,"
        b = b & """USE_WARPING_EFFECT"": false,"
        b = b & """SHAPE"": ""SB"","
        b = b & """DATATYPE"": 2,"
        b = b & """SECT_I"": {""vSIZE"": [" & f(2) & "," & JsonNum(SECTION_WIDTH) & _
                ",0,0,0,0,0,0,0,0]}"
        b = b & "}}"
    Next i

    b = b & "}}"

    PostSections = PostAndCheck("db/SECT", b, "PUT")

End Function


' db/CO_S, keyed by section number. Walks SECT_LIST rather than
' SECTION_COLOR_LIST so only sections that were actually written get
' coloured - a section skipped by its gate is skipped here too.
'
' Wireframe (2D) and hidden edge (3D outline) take the base colour; hidden
' fill (3D surface) takes a lightened version of it. A section with no
' entry in SECTION_COLOR_LIST is left at whatever colour Civil NX chose,
' which is harmless.
Private Function PostSectionColors() As String

    Dim rows() As String
    Dim f() As String
    Dim i As Long
    Dim b As String
    Dim baseRgb As String
    Dim c() As String
    Dim wrote As Boolean

    rows = Split(SECT_LIST, ";")
    b = "{""Assign"": {"
    wrote = False

    For i = LBound(rows) To UBound(rows)

        f = Split(rows(i), "|")
        baseRgb = SectionColorFor(f(0))

        If Len(baseRgb) > 0 Then

            c = Split(baseRgb, ",")
            If wrote Then b = b & ","
            wrote = True

            b = b & """" & f(0) & """: {"
            b = b & """W_R"": " & c(0) & ",""W_G"": " & c(1) & ",""W_B"": " & c(2) & ","
            b = b & """HF_R"": " & Lighten(c(0)) & ",""HF_G"": " & Lighten(c(1)) & _
                    ",""HF_B"": " & Lighten(c(2)) & ","
            b = b & """HE_R"": " & c(0) & ",""HE_G"": " & c(1) & ",""HE_B"": " & c(2) & ","
            b = b & """bBLEMD"": " & LCase(SECTION_COLOR_TRANSLUCENT) & ","
            b = b & """FACT"": " & JsonNum(SECTION_COLOR_OPACITY)
            b = b & "}"

        End If

    Next i

    b = b & "}}"

    If Not wrote Then
        PostSectionColors = "WARN: no section colours matched the sections written."
        Exit Function
    End If

    ' PUT: db/CO_S lists GET and PUT as its only active methods.
    PostSectionColors = PostAndCheck("db/CO_S", b, "PUT")

End Function

' Base R,G,B for a section number, or "" when that section has no entry.
Private Function SectionColorFor(ByVal sectNo As String) As String

    Dim rows() As String
    Dim f() As String
    Dim i As Long

    rows = Split(SECTION_COLOR_LIST, ";")

    For i = LBound(rows) To UBound(rows)
        f = Split(rows(i), "|")
        If f(0) = sectNo Then
            SectionColorFor = f(1)
            Exit Function
        End If
    Next i

End Function

' One channel lightened for the hidden fill, clamped to 255.
Private Function Lighten(ByVal channel As String) As Long

    Dim v As Long
    v = CLng(Round(CDbl(channel) * SECTION_FILL_LIGHTEN, 0))
    If v > 255 Then v = 255
    If v < 0 Then v = 0
    Lighten = v

End Function


Private Function PostNodes() As String

    Dim rows() As String
    Dim f() As String
    Dim i As Long
    Dim b As String

    rows = Split(NODE_LIST, ";")
    b = "{""Assign"": {"

    For i = LBound(rows) To UBound(rows)
        f = Split(rows(i), "|")
        If i > LBound(rows) Then b = b & ","
        b = b & """" & f(0) & """: {"
        b = b & """X"": " & f(1) & ","
        b = b & """Y"": " & f(2) & ","
        b = b & """Z"": " & f(3)
        b = b & "}"
    Next i

    b = b & "}}"

    PostNodes = PostAndCheck("db/NODE", b, "PUT")

End Function


Private Function PostElements() As String

    Dim rows() As String
    Dim f() As String
    Dim i As Long
    Dim b As String

    rows = Split(ELEM_LIST, ";")
    b = "{""Assign"": {"

    For i = LBound(rows) To UBound(rows)
        f = Split(rows(i), "|")
        If i > LBound(rows) Then b = b & ","
        b = b & """" & f(0) & """: {"
        b = b & """TYPE"": ""BEAM"","
        b = b & """MATL"": " & MATERIAL_NO & ","
        b = b & """SECT"": " & f(1) & ","
        b = b & """NODE"": [" & f(2) & "," & f(3) & "],"
        b = b & """ANGLE"": " & f(4)
        b = b & "}"
    Next i

    b = b & "}}"

    PostElements = PostAndCheck("db/ELEM", b, "PUT")

End Function


Private Function PostStaticLoadCases() As String

    Dim rows() As String
    Dim f() As String
    Dim i As Long
    Dim key As Long
    Dim b As String

    rows = Split(STLDCASE_LIST, ";")
    b = "{""Assign"": {"
    key = 0

    For i = LBound(rows) To UBound(rows)
        f = Split(rows(i), "|")
        ' Skipping EQ/ATA uses a running key rather than i+1, so the
        ' remaining cases still come out numbered 1..n with no holes.
        If SEISMIC_ACTIVE Or Not IsSeismicCase(f(0)) Then
            key = key + 1
            If key > 1 Then b = b & ","
            b = b & """" & key & """: {"
            b = b & """NAME"": """ & f(0) & ""","
            b = b & """TYPE"": """ & f(1) & ""","
            b = b & """DESC"": """""
            b = b & "}"
        End If
    Next i

    b = b & "}}"

    PostStaticLoadCases = TolerateDuplicateName(PostAndCheck("db/STLD", b, "PUT"), _
                                                "Static load cases")

End Function

' The two seismic static load cases. EQ carries the horizontal ground
' pressure (AddEqLoads), ATA the inertia self-weight (PostSelfWeight).
Private Function IsSeismicCase(ByVal nm As String) As Boolean
    IsSeismicCase = (nm = "EQ" Or nm = "ATA")
End Function


' *SELFWEIGHT under DL is (0, 0, -1); under ATA it is (B15, 0, 0) - the
' seismic inertia factor applied along global X. The ATA record is
' dropped entirely when SEISMIC_ACTIVE is False, along with its load case.
Private Function PostSelfWeight() As String

    Dim b As String

    b = "{""Assign"": {"
    b = b & SelfWeightJson(1, "DL", 0, 0, -1)
    If SEISMIC_ACTIVE Then
        b = b & "," & SelfWeightJson(2, "ATA", DIM_ATA_FACTOR, 0, 0)
    End If
    b = b & "}}"

    PostSelfWeight = PostAndCheck("db/BODF", b, "PUT")

End Function

Private Function SelfWeightJson(ByVal key As Long, ByVal lcname As String, _
                                ByVal fx As Double, ByVal fy As Double, _
                                ByVal fz As Double) As String

    Dim b As String
    b = """" & key & """: {"
    b = b & """LCNAME"": """ & lcname & ""","
    b = b & """GROUP_NAME"": """","
    b = b & """FV"": [" & JsonNum(fx) & "," & JsonNum(fy) & "," & JsonNum(fz) & "]"
    b = b & "}"

    SelfWeightJson = b

End Function


' db/BMLD keys on the ELEMENT number (same convention as db/PRES), and its
' ITEMS array is that element's complete load list - so every load case
' acting on one element has to be written in a single entry. Grouping is
' O(n^2) over a couple of dozen elements, which is free at this size.
Private Function PostBeamLoads() As String

    Dim rows() As String
    Dim f() As String
    Dim g() As String
    Dim i As Long, j As Long, itemNo As Long
    Dim elemNo As String
    Dim seen As String
    Dim b As String
    Dim wroteElem As Boolean

    If Len(BEAMLOAD_LIST) = 0 Then
        PostBeamLoads = "No beam loads were generated."
        Exit Function
    End If

    rows = Split(BEAMLOAD_LIST, ";")
    b = "{""Assign"": {"
    seen = ""
    wroteElem = False

    For i = LBound(rows) To UBound(rows)

        f = Split(rows(i), "|")
        elemNo = f(0)

        ' Each element is emitted once, on its first appearance.
        If InStr(1, seen, "[" & elemNo & "]", vbBinaryCompare) = 0 Then

            seen = seen & "[" & elemNo & "]"
            If wroteElem Then b = b & ","
            wroteElem = True

            b = b & """" & elemNo & """: {""ITEMS"": ["
            itemNo = 0

            For j = LBound(rows) To UBound(rows)
                g = Split(rows(j), "|")
                If g(0) = elemNo Then
                    itemNo = itemNo + 1
                    If itemNo > 1 Then b = b & ","
                    b = b & "{"
                    b = b & """ID"": " & itemNo & ","
                    b = b & """LCNAME"": """ & g(1) & ""","
                    b = b & """GROUP_NAME"": """","
                    b = b & """CMD"": """ & g(2) & ""","
                    b = b & """TYPE"": ""UNILOAD"","
                    b = b & """DIRECTION"": """ & g(3) & ""","
                    b = b & """USE_PROJECTION"": false,"
                    b = b & """USE_ECCEN"": false,"
                    b = b & """D"": [0,1,0,0],"
                    b = b & """P"": [" & g(4) & "," & g(5) & ",0,0]"
                    b = b & "}"
                End If
            Next j

            b = b & "]}"

        End If

    Next i

    b = b & "}}"

    PostBeamLoads = PostAndCheck("db/BMLD", b, "PUT")

End Function


' Writes the combinations in THREE separate PUTs, one per dependency tier,
' instead of all 33 in a single Assign object.
'
' WHY: a combination that references another one ("ANAL":"CB") needs its
' target to already exist. Sending everything in one Assign object makes
' that depend on the order the server walks the object's KEYS - and JSON
' object keys carry no ordering guarantee. Walked as strings they come out
' "1","10","11".."19","2","20".."29","3","30",... so ENV_SER at key 30 is
' reached while SLS-4..SLS-9 (keys 4-9) do not exist yet, and the CB
' lookup fails. Confirmed live 2026-09-21 as
'   {"error":{"message":"[Error] The load combination name="}}
' - MIDAS could not even name the offender, which is what an unresolved
' reference looks like rather than a duplicate name.
'
' Keys stay globally unique and in definition order (i + 1) across tiers,
' so the numbering is the same as a single-request write would have given.
' The tier number is reported on failure, which narrows any future error
' to a third of the set straight away.
Private Function PostLoadCombinations() As String

    Dim tier As Long
    Dim res As String

    If Len(LOADCOMB_LIST) = 0 Then
        PostLoadCombinations = "No load combinations were generated."
        Exit Function
    End If

    For tier = 1 To 3
        res = PostCombTier(tier)
        If Len(res) > 0 Then
            If Left$(res, 5) = "WARN:" Then
                PostLoadCombinations = res
            Else
                PostLoadCombinations = "tier " & tier & " - " & res
            End If
            If Left$(res, 5) <> "WARN:" Then Exit Function
        End If
    Next tier

    If Left$(PostLoadCombinations, 5) <> "WARN:" Then PostLoadCombinations = ""

End Function

' One PUT for the combinations at the given dependency tier. Returns ""
' when the tier is empty or wrote cleanly.
Private Function PostCombTier(ByVal tier As Long) As String

    Dim rows() As String
    Dim f() As String
    Dim items() As String
    Dim it() As String
    Dim i As Long, j As Long
    Dim b As String
    Dim wrote As Boolean

    rows = Split(LOADCOMB_LIST, ";")
    b = "{""Assign"": {"
    wrote = False

    For i = LBound(rows) To UBound(rows)
        f = Split(rows(i), "|")
        If CLng(f(4)) = tier Then

            If wrote Then b = b & ","
            wrote = True

            b = b & """" & (i + 1) & """: {"
            b = b & """NAME"": """ & f(0) & ""","
            b = b & """ACTIVE"": """ & f(1) & ""","
            b = b & """iTYPE"": " & f(2) & ","
            b = b & """DESC"": """","
            b = b & """vCOMB"": ["

            items = Split(f(3), ",")
            For j = LBound(items) To UBound(items)
                it = Split(items(j), ":")
                If j > LBound(items) Then b = b & ","
                b = b & "{""ANAL"": """ & it(0) & """, ""LCNAME"": """ & it(1) & _
                         """, ""FACTOR"": " & it(2) & "}"
            Next j

            b = b & "]"
            b = b & "}"

        End If
    Next i

    b = b & "}}"

    If Not wrote Then Exit Function

    PostCombTier = TolerateDuplicateName(PostAndCheck("db/LCOM-GEN", b, "PUT"), _
                                         "Load combinations")

End Function


' ===========================================================================
'  DIVIDE ELEMENTS (ope/DIVIDEELEM)
' ===========================================================================

' Divides the main span members in order: element 18 (midspan slab), 10 (left
' wall), 11 (right wall), and 4 (foundation span) into DIVIDE_AMOUNT equal segments.
' Executed at the very end of BuildCulvertModel so that all standard element
' IDs, loads, and load combinations are already in place.
Private Function PostDivideElements() As String

    Dim targets(1 To 4) As Long
    Dim i As Long
    Dim b As String
    Dim res As String

    If DIVIDE_AMOUNT <= 1 Then
        PostDivideElements = "WARN: skipped - " & PROJECT_SHEET_NAME & "!" & _
                             CELL_DIVIDE_AMOUNT & " is " & DIVIDE_AMOUNT & _
                             " (must be >= 2 to divide)."
        Exit Function
    End If

    targets(1) = 18
    targets(2) = 10
    targets(3) = 11
    targets(4) = 4

    For i = 1 To 4
        b = "{""Argument"": {" & _
            """TARGETS"": [" & targets(i) & "]," & _
            """DIVIDE"": {" & _
                """ELEM_TYPE"": ""Frame""," & _
                """DIV_METHOD"": ""Equal""," & _
                """OPTION"": {" & _
                    """EQUAL_OPTION"": {" & _
                        """NUM_X"": " & DIVIDE_AMOUNT & _
                    "}}}" & _
            "}}"

        res = PostAndCheck("ope/DIVIDEELEM", b, "POST")
        If Len(res) > 0 Then
            PostDivideElements = "element " & targets(i) & " - " & res
            Exit Function
        End If
    Next i

    PostDivideElements = ""

End Function


' ===========================================================================
'  FOUNDATION SPRINGS (db/NSPR)
' ===========================================================================

' Assigns linear point springs (db/NSPR) to all foundation nodes at Z = 0.
' Subgrade modulus (Kz per unit area) is read from MIDAS_INPUT!B42 (Yay).
'
' Runs AFTER PostDivideElements() so that all newly divided foundation
' nodes are present in the model and included in the spring assignment.
'
' Tributary length formulation for sorted foundation nodes (x_1 <= x_2 <= ... <= x_m):
'   Intermediate nodes (2 <= i <= m-1):
'     L_trib = (x_{i+1} - x_{i-1}) / 2
'   Corner joints/nodes (Node 3 and Node 6 in no-extension case, index 1 and m):
'     Node 1: L_trib = (x_2 - x_1) / 2 + wall_thickness / 2
'     Node m: L_trib = (x_m - x_{m-1}) / 2 + wall_thickness / 2
'   Nodal spring stiffness:
'     Kz = ks * L_trib
'     Kx = Kz / 2
'     Ky = Kz / 2
Private Function PostFoundationSprings() As String

    Dim resp As String
    Dim statusCode As Long
    Dim nodeIds() As Long
    Dim nodeXs() As Double
    Dim nodeCount As Long
    Dim i As Long, j As Long
    Dim tmpId As Long, tmpX As Double
    Dim ltrib As Double
    Dim kz As Double, kx As Double, ky As Double
    Dim b As String
    Dim wrote As Boolean

    If SPRING_KZ_MODULUS <= 0 Then
        PostFoundationSprings = "WARN: skipped - " & INPUT_SHEET_NAME & "!B42 spring modulus is " & _
                                JsonNum(SPRING_KZ_MODULUS) & " (must be > 0 to assign springs)."
        Exit Function
    End If

    Call SendApiRequest("GET", "db/NODE", "", resp, statusCode)
    If statusCode < 200 Or statusCode >= 300 Then
        PostFoundationSprings = "HTTP " & statusCode & " - failed to query db/NODE for foundation springs."
        Exit Function
    End If

    Call ParseFoundationNodes(resp, nodeIds, nodeXs, nodeCount)

    If nodeCount < 2 Then
        PostFoundationSprings = "FAIL: found " & nodeCount & " foundation nodes at Z=0 (expected at least 2)."
        Exit Function
    End If

    ' Sort foundation nodes in ascending order of X coordinate
    For i = 1 To nodeCount - 1
        For j = i + 1 To nodeCount
            If nodeXs(j) < nodeXs(i) Then
                tmpX = nodeXs(i): nodeXs(i) = nodeXs(j): nodeXs(j) = tmpX
                tmpId = nodeIds(i): nodeIds(i) = nodeIds(j): nodeIds(j) = tmpId
            End If
        Next j
    Next i

    b = "{""Assign"": {"
    wrote = False

    For i = 1 To nodeCount
        If i = 1 Then
            ' Left corner node
            ltrib = (nodeXs(2) - nodeXs(1)) / 2
            If Not HAS_EXT Then
                ltrib = ltrib + DIM_WALL_T / 2
            End If
        ElseIf i = nodeCount Then
            ' Right corner node
            ltrib = (nodeXs(nodeCount) - nodeXs(nodeCount - 1)) / 2
            If Not HAS_EXT Then
                ltrib = ltrib + DIM_WALL_T / 2
            End If
        Else
            ' Intermediate foundation node
            ltrib = (nodeXs(i + 1) - nodeXs(i - 1)) / 2
        End If

        kz = SPRING_KZ_MODULUS * ltrib
        kx = kz / 2
        ky = kz / 2

        If wrote Then b = b & ","
        wrote = True

        ' TYPE "LINEAR" + SDR is the "by Point Spring Function" mode
        ' (FormType 0, the default - confirmed against the "Point Spring"
        ' doc's own Linear Type example, which omits FormType entirely).
        ' FormType 1 ("by Surface Spring Function") is a DIFFERENT mode
        ' that computes stiffness from EFFAREA/DK instead of SDR - do not
        ' set FormType: 1 here, since SDR already carries our own
        ' tributary-length-computed values and EFFAREA/DK are never sent.
        b = b & """" & nodeIds(i) & """: {"
        b = b & """ITEMS"": [{"
        b = b & """ID"": 1,"
        b = b & """TYPE"": ""LINEAR"","
        b = b & """SDR"": [" & JsonNum(kx) & "," & JsonNum(ky) & "," & JsonNum(kz) & ",0,0,0],"
        b = b & """F_S"": [false,false,false,false,false,false],"
        b = b & """DAMPING"": false,"
        b = b & """Cr"": [0,0,0,0,0,0]"
        b = b & "}]}"
    Next i

    b = b & "}}"

    PostFoundationSprings = PostAndCheck("db/NSPR", b, "PUT")

End Function


' Parses all nodes from db/NODE JSON where |Z| < 0.001.
Private Sub ParseFoundationNodes(ByVal json As String, ByRef nodeIds() As Long, _
                                ByRef nodeXs() As Double, ByRef count As Long)

    Dim pNode As Long, pos As Long, posColon As Long
    Dim posQ1 As Long, posQ2 As Long
    Dim posObjStart As Long, posObjEnd As Long
    Dim posNextQuote As Long
    Dim keyStr As String
    Dim nid As Long
    Dim xVal As Double, zVal As Double
    Dim maxCap As Long

    count = 0
    maxCap = 64
    ReDim nodeIds(1 To maxCap)
    ReDim nodeXs(1 To maxCap)

    pNode = InStr(1, json, """NODE""", vbTextCompare)
    If pNode = 0 Then Exit Sub

    pos = InStr(pNode, json, "{")
    If pos = 0 Then Exit Sub
    pos = pos + 1

    Do While pos < Len(json)
        posQ1 = InStr(pos, json, """")
        If posQ1 = 0 Then Exit Do
        posQ2 = InStr(posQ1 + 1, json, """")
        If posQ2 = 0 Then Exit Do

        keyStr = Mid$(json, posQ1 + 1, posQ2 - posQ1 - 1)

        posColon = InStr(posQ2 + 1, json, ":")
        If posColon = 0 Then Exit Do

        posObjStart = InStr(posColon + 1, json, "{")
        If posObjStart = 0 Then Exit Do

        posNextQuote = InStr(posColon + 1, json, """")
        If posNextQuote > 0 And posNextQuote < posObjStart Then
            pos = posQ2 + 1
        Else
            posObjEnd = InStr(posObjStart + 1, json, "}")
            If posObjEnd = 0 Then Exit Do

            If IsNumeric(keyStr) Then
                nid = CLng(keyStr)
                xVal = ExtractJsonNumFromBlock(json, posObjStart, posObjEnd, "X")
                zVal = ExtractJsonNumFromBlock(json, posObjStart, posObjEnd, "Z")

                If Abs(zVal) < 0.001 Then
                    count = count + 1
                    If count > maxCap Then
                        maxCap = maxCap * 2
                        ReDim Preserve nodeIds(1 To maxCap)
                        ReDim Preserve nodeXs(1 To maxCap)
                    End If
                    nodeIds(count) = nid
                    nodeXs(count) = xVal
                End If
            End If

            pos = posObjEnd + 1
        End If
    Loop

End Sub


' Extracts a numeric field value (e.g. "X": 0.225) from within a JSON object block.
' Uses Val() so that the extraction is locale-invariant and stops at delimiters.
Private Function ExtractJsonNumFromBlock(ByVal json As String, ByVal pStart As Long, _
                                        ByVal pEnd As Long, ByVal key As String) As Double

    Dim blk As String
    Dim pk As Long, pc As Long

    blk = Mid$(json, pStart, pEnd - pStart + 1)
    pk = InStr(1, blk, """" & key & """", vbTextCompare)
    If pk = 0 Then
        ExtractJsonNumFromBlock = 0
        Exit Function
    End If

    pc = InStr(pk + Len(key) + 2, blk, ":")
    If pc = 0 Then
        ExtractJsonNumFromBlock = 0
        Exit Function
    End If

    ExtractJsonNumFromBlock = Val(Mid$(blk, pc + 1))

End Function


' ===========================================================================
'  NON-RIGID ELEMENT LIST (INPUT!G15)
' ===========================================================================

' Queries all elements from db/ELEM, filters out rigid section elements
' (Section 7: DOS RIJIT, Section 8: DUV RIJIT, Section 9: TEM RIJIT),
' formats the non-rigid elements into range string syntax (e.g. "4 10to13 17to19 21to40"),
' and pastes the string into INPUT!G15 prior to running analysis.
Private Function PostNonRigidElementList() As String

    Dim resp As String
    Dim statusCode As Long
    Dim elemIds() As Long
    Dim elemCount As Long
    Dim i As Long, j As Long, tmpId As Long
    Dim elemListStr As String
    Dim wsInput As Worksheet

    Call SendApiRequest("GET", "db/ELEM", "", resp, statusCode)
    If statusCode < 200 Or statusCode >= 300 Then
        PostNonRigidElementList = "HTTP " & statusCode & " - failed to query db/ELEM for element list."
        Exit Function
    End If

    Call ParseNonRigidElements(resp, elemIds, elemCount)

    If elemCount = 0 Then
        PostNonRigidElementList = "WARN: no non-rigid elements found in model."
        Exit Function
    End If

    ' Sort element IDs ascending
    For i = 1 To elemCount - 1
        For j = i + 1 To elemCount
            If elemIds(j) < elemIds(i) Then
                tmpId = elemIds(i): elemIds(i) = elemIds(j): elemIds(j) = tmpId
            End If
        Next j
    Next i

    elemListStr = FormatElementRangeList(elemIds, elemCount)

    On Error Resume Next
    Set wsInput = ThisWorkbook.Worksheets(PROJECT_SHEET_NAME)
    On Error GoTo 0

    If wsInput Is Nothing Then
        PostNonRigidElementList = "WARN: sheet """ & PROJECT_SHEET_NAME & """ not found; " & _
                                  "could not write element list: " & elemListStr
        Exit Function
    End If

    On Error Resume Next
    wsInput.Range(CELL_ELEM_LIST).Value = elemListStr
    If Err.Number <> 0 Then
        PostNonRigidElementList = "WARN: failed to write to " & PROJECT_SHEET_NAME & "!" & _
                                  CELL_ELEM_LIST & ": " & Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    PostNonRigidElementList = ""

End Function


' Parses all elements from db/ELEM JSON excluding rigid sections 7, 8, 9.
Private Sub ParseNonRigidElements(ByVal json As String, ByRef elemIds() As Long, ByRef count As Long)

    Dim pElem As Long, pos As Long, posColon As Long
    Dim posQ1 As Long, posQ2 As Long
    Dim posObjStart As Long, posObjEnd As Long
    Dim posNextQuote As Long
    Dim keyStr As String
    Dim eid As Long
    Dim sectNo As Long
    Dim maxCap As Long

    count = 0
    maxCap = 64
    ReDim elemIds(1 To maxCap)

    pElem = InStr(1, json, """ELEM""", vbTextCompare)
    If pElem = 0 Then Exit Sub

    pos = InStr(pElem, json, "{")
    If pos = 0 Then Exit Sub
    pos = pos + 1

    Do While pos < Len(json)
        posQ1 = InStr(pos, json, """")
        If posQ1 = 0 Then Exit Do
        posQ2 = InStr(posQ1 + 1, json, """")
        If posQ2 = 0 Then Exit Do

        keyStr = Mid$(json, posQ1 + 1, posQ2 - posQ1 - 1)

        posColon = InStr(posQ2 + 1, json, ":")
        If posColon = 0 Then Exit Do

        posObjStart = InStr(posColon + 1, json, "{")
        If posObjStart = 0 Then Exit Do

        posNextQuote = InStr(posColon + 1, json, """")
        If posNextQuote > 0 And posNextQuote < posObjStart Then
            pos = posQ2 + 1
        Else
            posObjEnd = InStr(posObjStart + 1, json, "}")
            If posObjEnd = 0 Then Exit Do

            If IsNumeric(keyStr) Then
                eid = CLng(keyStr)
                sectNo = CLng(ExtractJsonNumFromBlock(json, posObjStart, posObjEnd, "SECT"))

                ' Rigid sections are 7 (DOS RIJIT), 8 (DUV RIJIT), 9 (TEM RIJIT)
                If sectNo <> 7 And sectNo <> 8 And sectNo <> 9 Then
                    count = count + 1
                    If count > maxCap Then
                        maxCap = maxCap * 2
                        ReDim Preserve elemIds(1 To maxCap)
                    End If
                    elemIds(count) = eid
                End If
            End If

            pos = posObjEnd + 1
        End If
    Loop

End Sub


' Compresses a sorted array of element IDs into MIDAS range string format (e.g. "4 10to13 17to19 21to40").
Private Function FormatElementRangeList(ByRef ids() As Long, ByVal count As Long) As String

    Dim out As String
    Dim startId As Long, prevId As Long
    Dim i As Long

    If count = 0 Then Exit Function

    startId = ids(1)
    prevId = ids(1)

    For i = 2 To count
        If ids(i) = prevId + 1 Then
            prevId = ids(i)
        Else
            If Len(out) > 0 Then out = out & " "
            If startId = prevId Then
                out = out & CStr(startId)
            Else
                out = out & CStr(startId) & "to" & CStr(prevId)
            End If
            startId = ids(i)
            prevId = ids(i)
        End If
    Next i

    If Len(out) > 0 Then out = out & " "
    If startId = prevId Then
        out = out & CStr(startId)
    Else
        out = out & CStr(startId) & "to" & CStr(prevId)
    End If

    FormatElementRangeList = out

End Function


' ===========================================================================
'  PERFORM ANALYSIS (doc/ANAL)
' ===========================================================================


' Runs structural analysis in MIDAS Civil NX after the entire model, loads,
' load combinations, element divisions, and foundation springs are in place.
Private Function PostPerformAnalysis() As String
    PostPerformAnalysis = PostAndCheck("doc/ANAL", "{}", "POST")
End Function


' ===========================================================================
'  BEAM FORCE RESULTS (post/TABLE)
'
'  Retrieves beam force result tables from MIDAS Civil NX after analysis
'  and populates both tables on the "MIDAS_RESULTS" sheet:
'    Table 1 (Cols B:J, starting row 3):
'      All 28 SLS & ULS combinations (plus EQ-1 if active; ACC-1 is
'      excluded from this pull by request even though it is still built
'      as a combination - see GenerateLoadCombinations)
'    Table 2 (Cols S:AA, starting row 36):
'      The 4 governing envelopes: ENV_SER(max), ENV_ALL(max),
'                                 ENV_SER(min), ENV_ALL(min)
'  Downstream sheet "6_DONATI" reads summary values calculated from these tables.
' ===========================================================================
Private Function PostBeamForceResults() As String

    Dim ws As Worksheet
    Dim t1Combos As String
    Dim t2Combos As String
    Dim i As Long
    Dim body As String
    Dim resp As String
    Dim statusCode As Long
    Dim dataArr() As Variant
    Dim rowCount1 As Long, rowCount2 As Long
    Dim lastUsedRow As Long
    Dim fRow As Long
    Dim combRows() As String
    Dim combName As String

    ' ThisWorkbook, not ActiveWorkbook: Progress() calls DoEvents, so the
    ' user can switch workbooks during a long build.
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RESULT_SHEET_NAME)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets("MIDAS_RESULT")
    End If
    On Error GoTo 0

    If ws Is Nothing Then
        PostBeamForceResults = "WARN: Worksheet '" & RESULT_SHEET_NAME & "' not found in this workbook."
        Exit Function
    End If

    ' --- Clean existing data in both tables before writing fresh results ---
    ' Table 1: Columns B:J from row 3 downwards
    lastUsedRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastUsedRow >= 3 Then
        ws.Range(ws.Cells(3, 2), ws.Cells(lastUsedRow, 10)).ClearContents
    End If

    ' Table 2: Columns S:AA from row 36 downwards
    lastUsedRow = ws.Cells(ws.Rows.Count, 19).End(xlUp).Row
    If lastUsedRow >= 36 Then
        ws.Range(ws.Cells(36, 19), ws.Cells(lastUsedRow, 27)).ClearContents
    End If

    ' --- Table 1: Primary load combinations ---
    ' Every SLS-*, ULS-* and EQ-1 combination GenerateLoadCombinations
    ' actually produced, in its order - taken from LOADCOMB_LIST rather than
    ' a hardcoded "1 To 14", so a combination added there cannot silently
    ' miss this table. ACC-1 stays out by request; EQ-1 is only in the list
    ' when the seismic gate is on.
    t1Combos = ""
    combRows = Split(LOADCOMB_LIST, ";")
    For i = LBound(combRows) To UBound(combRows)
        combName = Split(combRows(i), "|")(0)
        If Left$(combName, 4) = "SLS-" Or Left$(combName, 4) = "ULS-" Or combName = "EQ-1" Then
            t1Combos = t1Combos & IIf(Len(t1Combos) > 0, ",", "") & """" & combName & "(CB)"""
        End If
    Next i

    body = BuildTableReqJson("BeamForce_T1", t1Combos)
    Call SendApiRequest("POST", "post/TABLE", body, resp, statusCode)
    If statusCode < 200 Or statusCode >= 300 Or InStr(1, resp, """error""", vbTextCompare) > 0 Then
        PostBeamForceResults = "Table 1 request failed: HTTP " & statusCode & _
                               IIf(Len(ExtractJsonStringValue(resp, "message")) > 0, " - " & ExtractJsonStringValue(resp, "message"), "")
        Exit Function
    End If

    If Not ParseBeamForceTableData(resp, dataArr, rowCount1) Or rowCount1 = 0 Then
        PostBeamForceResults = "Failed to parse Table 1 response data."
        Exit Function
    End If

    ' Write Table 1 to Columns B:J starting at row 3
    ws.Range(ws.Cells(3, 2), ws.Cells(3 + rowCount1 - 1, 10)).Value = dataArr

    ' Clear any leftover rows below Table 1
    lastUsedRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastUsedRow > 3 + rowCount1 - 1 Then
        ws.Range(ws.Cells(3 + rowCount1, 2), ws.Cells(lastUsedRow, 10)).ClearContents
    End If

    ' Ensure helper formulas in Cols L:O extend down to row 3 + rowCount1 - 1 if needed
    fRow = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row
    If fRow < 3 + rowCount1 - 1 And fRow >= 3 Then
        ws.Range(ws.Cells(3, 12), ws.Cells(3, 15)).AutoFill _
            Destination:=ws.Range(ws.Cells(3, 12), ws.Cells(3 + rowCount1 - 1, 15))
    End If

    ' --- Table 2: Envelope load combinations ---
    t2Combos = """ENV_SER(CB:max)"",""ENV_ALL(CB:max)"",""ENV_SER(CB:min)"",""ENV_ALL(CB:min)"""

    body = BuildTableReqJson("BeamForce_T2", t2Combos)
    Call SendApiRequest("POST", "post/TABLE", body, resp, statusCode)
    If statusCode < 200 Or statusCode >= 300 Or InStr(1, resp, """error""", vbTextCompare) > 0 Then
        PostBeamForceResults = "Table 2 request failed: HTTP " & statusCode & _
                               IIf(Len(ExtractJsonStringValue(resp, "message")) > 0, " - " & ExtractJsonStringValue(resp, "message"), "")
        Exit Function
    End If

    If Not ParseBeamForceTableData(resp, dataArr, rowCount2) Or rowCount2 = 0 Then
        PostBeamForceResults = "Failed to parse Table 2 response data."
        Exit Function
    End If

    ' Write Table 2 to Columns S:AA starting at row 36
    ws.Range(ws.Cells(36, 19), ws.Cells(36 + rowCount2 - 1, 27)).Value = dataArr

    ' Clear any leftover rows below Table 2
    lastUsedRow = ws.Cells(ws.Rows.Count, 19).End(xlUp).Row
    If lastUsedRow > 36 + rowCount2 - 1 Then
        ws.Range(ws.Cells(36 + rowCount2, 19), ws.Cells(lastUsedRow, 27)).ClearContents
    End If

    ' Ensure helper formulas in Cols AC:AF extend down to row 36 + rowCount2 - 1 if needed
    fRow = ws.Cells(ws.Rows.Count, 29).End(xlUp).Row
    If fRow < 36 + rowCount2 - 1 And fRow >= 36 Then
        ws.Range(ws.Cells(36, 29), ws.Cells(36, 32)).AutoFill _
            Destination:=ws.Range(ws.Cells(36, 29), ws.Cells(36 + rowCount2 - 1, 32))
    End If

    PostBeamForceResults = ""

End Function

Private Function BuildTableReqJson(ByVal tableName As String, ByVal loadCaseListJson As String) As String

    Dim b As String

    b = "{" & _
        """Argument"": {" & _
            """TABLE_NAME"": """ & tableName & """," & _
            """TABLE_TYPE"": ""BEAMFORCE""," & _
            """UNIT"": {""FORCE"": ""kN"", ""DIST"": ""m""}," & _
            """STYLES"": {""FORMAT"": ""Fixed"", ""PLACE"": 2}," & _
            """COMPONENTS"": [""Elem"",""Load"",""Part"",""Axial"",""Shear-y"",""Shear-z"",""Torsion"",""Moment-y"",""Moment-z""]," & _
            """LOAD_CASE_NAMES"": [" & loadCaseListJson & "]," & _
            """PARTS"": [""PartI"",""Part2/4"",""PartJ""]" & _
        "}" & _
    "}"

    BuildTableReqJson = b

End Function

' Parses the DATA array from post/TABLE JSON into a 2D Variant array:
'   outArr(1 To rowCount, 1 To 9)
' Cols 1 to 9 correspond to:
'   1: Elem (Long)
'   2: Load (String)
'   3: Part (String)
'   4: Axial (Double)
'   5: Shear-y (Double)
'   6: Shear-z (Double)
'   7: Torsion (Double)
'   8: Moment-y (Double)
'   9: Moment-z (Double)
Private Function ParseBeamForceTableData(ByVal json As String, _
                                         ByRef outArr() As Variant, _
                                         ByRef outCount As Long) As Boolean

    Dim posData As Long, posOpen As Long, pos As Long
    Dim qStart As Long, qEnd As Long
    Dim pClose As Long
    Dim tok As String
    Dim colIdx As Long
    Dim cap As Long
    Dim rowIdx As Long
    Dim ch As String

    outCount = 0
    ParseBeamForceTableData = False

    posData = InStr(1, json, """DATA""", vbBinaryCompare)
    If posData = 0 Then Exit Function

    posOpen = InStr(posData, json, "[")
    If posOpen = 0 Then Exit Function

    cap = 500
    ReDim outArr(1 To 9, 1 To cap)

    pos = posOpen + 1
    colIdx = 0
    rowIdx = 0

    Do
        qStart = InStr(pos, json, """")
        If qStart = 0 Then Exit Do

        qEnd = InStr(qStart + 1, json, """")
        If qEnd = 0 Then Exit Do

        tok = Mid$(json, qStart + 1, qEnd - qStart - 1)
        pos = qEnd + 1

        Select Case colIdx
            Case 0
                ' Index token (e.g. "1") - skip, not written to sheet
                rowIdx = rowIdx + 1
                If rowIdx > cap Then
                    cap = cap * 2
                    ReDim Preserve outArr(1 To 9, 1 To cap)
                End If
            Case 1
                outArr(1, rowIdx) = CLng(Val(tok))
            Case 2
                outArr(2, rowIdx) = tok
            Case 3
                outArr(3, rowIdx) = tok
            Case 4
                outArr(4, rowIdx) = Val(tok)
            Case 5
                outArr(5, rowIdx) = Val(tok)
            Case 6
                outArr(6, rowIdx) = Val(tok)
            Case 7
                outArr(7, rowIdx) = Val(tok)
            Case 8
                outArr(8, rowIdx) = Val(tok)
            Case 9
                outArr(9, rowIdx) = Val(tok)
        End Select

        colIdx = colIdx + 1
        If colIdx = 10 Then
            colIdx = 0
            pClose = InStr(pos, json, "]")
            If pClose > 0 Then
                pos = pClose + 1
                Do While pos <= Len(json)
                    ch = Mid$(json, pos, 1)
                    If ch <> " " And ch <> vbCr And ch <> vbLf And ch <> vbTab Then
                        If ch = "]" Then
                            Exit Do
                        ElseIf ch = "," Then
                            Exit Do
                        End If
                    End If
                    pos = pos + 1
                Loop
                If ch = "]" Then Exit Do
            End If
        End If
    Loop

    If rowIdx = 0 Then Exit Function

    Dim finalArr() As Variant
    Dim r As Long, c As Long

    ReDim finalArr(1 To rowIdx, 1 To 9)
    For r = 1 To rowIdx
        For c = 1 To 9
            finalArr(r, c) = outArr(c, r)
        Next c
    Next r

    outArr = finalArr
    outCount = rowIdx
    ParseBeamForceTableData = True

End Function


' ===========================================================================
'  HTTP
' ===========================================================================

' Returns "" on success, else a short failure message. Success = 2xx AND no
' "error" key: PUT returns 200 and POST 201 (so "=200" would reject POSTs),
' and a bad body comes back as HTTP 200 with an error body and nothing
' written - status alone is not sufficient. See CLAUDE.md.
Private Function PostAndCheck(ByVal path As String, ByVal body As String, _
                              Optional ByVal httpMethod As String = "POST") As String

    Dim resp As String
    Dim statusCode As Long
    Dim msg As String

    Call SendApiRequest(httpMethod, path, body, resp, statusCode)

    If statusCode >= 200 And statusCode < 300 And InStr(1, resp, """error""", vbTextCompare) = 0 Then
        PostAndCheck = ""
        Exit Function
    End If

    ' The extracted message plus the RAW body. The raw copy is not
    ' redundant: it is the only thing that survives a bad extraction, and
    ' a truncated message once sent a whole debugging session the wrong
    ' way (see ExtractJsonStringValue). Keep both.
    msg = ExtractJsonStringValue(resp, "message")
    PostAndCheck = "HTTP " & statusCode & HttpStatusHint(statusCode) & _
                   IIf(Len(msg) > 0, " - " & msg, "") & _
                   vbCrLf & "       raw: " & Left$(resp, 400)

End Function

' Escapes a string for use as a JSON value. Everything else this module
' emits is a name it generated itself, but the project name is free text
' typed into a worksheet, so a stray quote or backslash there would
' otherwise produce a malformed body.
Private Function JsonEsc(ByVal s As String) As String

    Dim i As Long
    Dim c As String
    Dim out As String

    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        Select Case c
            Case """": out = out & "\"""
            Case "\": out = out & "\\"
            Case vbCr: out = out & "\r"
            Case vbLf: out = out & "\n"
            Case vbTab: out = out & "\t"
            Case Else
                If AscW(c) < 32 Then
                    out = out & " "
                Else
                    out = out & c
                End If
        End Select
    Next i

    JsonEsc = out

End Function

' Locale-invariant number -> JSON literal. Plain "&" is locale-aware (a
' Turkish Windows renders 0.4 as "0,4"), which corrupts JSON silently and
' can shift array element counts. Str$ is always period-decimal but drops
' the leading zero for |v|<1 (".001"), also invalid JSON - re-added here.
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

Private Sub SendApiRequest(ByVal httpMethod As String, ByVal path As String, ByVal body As String, _
                           ByRef responseText As String, ByRef statusCode As Long)

    Dim url As String

    url = API_BASE_URL & "/" & path

    If Len(MapiKey()) = 0 Then
        responseText = MapiKeyProblem()
        statusCode = 0
        Exit Sub
    End If

    ' Reuse one client across the run rather than paying COM instantiation
    ' and a fresh connection per request - a build makes hundreds.
    If HTTP_CLIENT Is Nothing Then
        Set HTTP_CLIENT = CreateObject("WinHttp.WinHttpRequest.5.1")
    End If

    On Error Resume Next

    ' resolve, connect, send, receive (ms). A long receive timeout ONLY for
    ' the solve and the request right after it (post/TABLE reads the fresh
    ' results): doc/ANAL is synchronous and a real model can take longer
    ' than WinHTTP's 30 s default. Everything else keeps the defaults, set
    ' explicitly because the one client is reused across calls.
    If path = "doc/ANAL" Or path = "post/TABLE" Then
        HTTP_CLIENT.SetTimeouts 0, 60000, 30000, 600000
    Else
        HTTP_CLIENT.SetTimeouts 0, 60000, 30000, 30000
    End If

    HTTP_CLIENT.Open httpMethod, url, False
    HTTP_CLIENT.SetRequestHeader "MAPI-Key", MapiKey()
    HTTP_CLIENT.SetRequestHeader "Content-Type", "application/json"
    If Err.Number <> 0 Then
        responseText = "WinHTTP error: " & Err.Description
        statusCode = 0
        Err.Clear
        Set HTTP_CLIENT = Nothing
        On Error GoTo 0
        Exit Sub
    End If

    If Len(body) > 0 Then
        HTTP_CLIENT.Send body
    Else
        HTTP_CLIENT.Send
    End If
    If Err.Number <> 0 Then
        responseText = "WinHTTP error: " & Err.Description
        statusCode = 0
        Err.Clear
        ' Never reuse a client that just failed - the next call builds one.
        Set HTTP_CLIENT = Nothing
        On Error GoTo 0
        Exit Sub
    End If

    On Error GoTo 0

    statusCode = HTTP_CLIENT.Status
    responseText = HTTP_CLIENT.ResponseText

End Sub

' Pulls the string value of "key": "value" out of a JSON response,
' honouring backslash escapes so an embedded \" does not cut the value
' short. Still not a general JSON parser - only for a top-level string
' field such as an error response's "message".
'
' WHY THE ESCAPE HANDLING MATTERS: MIDAS quotes the offending name inside
' its error text, e.g.
'   {"error":{"message":"[Error] The load combination name=\"ENV_SER\" ..."}}
' The old version took everything between the first two quote characters
' after the colon, so it stopped dead at that \" and returned
'   "[Error] The load combination name="
' That is not just a cosmetic truncation - it cut off the part of the
' message TolerateDuplicateName matches on, so a re-run collision that
' should have been a WARN surfaced as a hard FAIL with a message that
' trailed off mid-sentence. Confirmed live 2026-09-21.
Private Function ExtractJsonStringValue(ByVal json As String, ByVal key As String) As String

    Dim posKey As Long, posColon As Long, i As Long
    Dim c As String, nxt As String
    Dim out As String

    posKey = InStr(1, json, """" & key & """", vbTextCompare)
    If posKey = 0 Then Exit Function

    posColon = InStr(posKey + Len(key) + 2, json, ":")
    If posColon = 0 Then Exit Function

    ' Skip whitespace to the opening quote of the value. Anything else
    ' means this key does not hold a string, so there is nothing to take.
    i = posColon + 1
    Do While i <= Len(json)
        c = Mid$(json, i, 1)
        If c = """" Then Exit Do
        If c <> " " And c <> vbTab Then Exit Function
        i = i + 1
    Loop
    If i > Len(json) Then Exit Function
    i = i + 1

    Do While i <= Len(json)
        c = Mid$(json, i, 1)
        If c = "\" Then
            nxt = Mid$(json, i + 1, 1)
            Select Case nxt
                Case """": out = out & """"
                Case "\": out = out & "\"
                Case "/": out = out & "/"
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case "t": out = out & vbTab
                Case Else: out = out & nxt
            End Select
            i = i + 2
        ElseIf c = """" Then
            Exit Do
        Else
            out = out & c
            i = i + 1
        End If
    Loop

    ExtractJsonStringValue = out

End Function


' The MAPI key from MAPI_KEY_SHEET!MAPI_KEY_CELL, or "" if it cannot be
' read. Cached after the first successful read.
Private Function MapiKey() As String

    Dim ws As Worksheet

    If Len(MAPI_KEY_CACHE) > 0 Then
        MapiKey = MAPI_KEY_CACHE
        Exit Function
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(MAPI_KEY_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    MAPI_KEY_CACHE = Trim$(CStr(ws.Range(MAPI_KEY_CELL).Value))
    MapiKey = MAPI_KEY_CACHE

End Function


' "" when a key is available, otherwise the message to show the user.
' Every request goes through this first: without a key the API answers
' 404 "client does not exist" to everything, which sends people hunting
' through their JSON for a fault that is not there.
Private Function MapiKeyProblem() As String

    If Len(MapiKey()) = 0 Then
        MapiKeyProblem = "No MAPI key in " & MAPI_KEY_SHEET & "!" & MAPI_KEY_CELL & _
            ". Copy it from Civil NX (Tools > API > API Setting) into that cell " & _
            "and run again."
    End If

End Function

' Turns a MIDAS HTTP status into something a user can act on. The codes
' are from the API manual's own table:
'   200 Success   - request reached the model
'   201 Created   - POST succeeded (PUT answers 200), which is why every
'                   status test here accepts the whole 2xx range
'   400 Bad Request - wrong command or body
'   403 Forbidden   - the API is not switched on for this user
'   404 Not found   - the client never reached the API server at all
' 404 in particular is almost never a bug in the request: it means Civil
' NX is closed, the model is not open, or the MAPI key belongs to a
' session that has gone away. Printing the bare number sends people
' looking in the wrong place.
Private Function HttpStatusHint(ByVal statusCode As Long) As String

    Select Case statusCode
        Case 400
            HttpStatusHint = " (Bad Request - the command or body is wrong.)"
        Case 403
            HttpStatusHint = " (Forbidden - the API is not enabled. In Civil NX: " & _
                             "Tools > API > API Setting.)"
        Case 404
            HttpStatusHint = " (Not Found - not connected to the API server. Check " & _
                             "Civil NX is running with the model open, and that " & _
                             "MAPI_KEY matches the key in Tools > API > API Setting.)"
        Case 0
            HttpStatusHint = " (No response - the request never completed.)"
        Case Else
            HttpStatusHint = ""
    End Select

End Function


' Moves the verdict to the top of the report, right under the title line.
' MsgBox shows only about 1024 characters, and the failing step used to be
' the LAST line - the first thing to be cut off. The failure is the last
' "FAIL - " entry: the clean pass runs first, so any FAIL lines it logged
' come earlier.
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


' MsgBox shows only about 1024 characters and silently drops the rest.
' When the report is longer, the whole text goes to <logName>_log.txt next
' to the workbook (or in %TEMP% if it has never been saved) and the MsgBox
' shows the start of it plus where the rest is. Callers put the verdict
' first, so what gets cut is the least important part.
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
