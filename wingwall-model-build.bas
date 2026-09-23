Option Explicit

' ============================================================================
'  MIDAS Civil NX API - Wingwall Model Build
'  Full reference/derivations: CLAUDE.md. Keep comments here SHORT - VBA
'  caps a module at 64KB and this file has already hit it once, which shows
'  up as "Sub or Function not defined" on a function that is plainly there.
' ----------------------------------------------------------------------------
'  db/UNIT PJCF STYP ACTL MATL THIK CO_T NUCS NODE ELEM STLD BODF PRES PNLD
'  PNLA LCOM-GEN - all written with PUT, never POST (POST dies on "Key
'  Already Exist"; PUT is create-or-update). ope/DIVIDEELEM (mesh),
'  ope/SSPS (surface springs) and doc/ANAL (run the solver) are the
'  exceptions: POST is their only active method. Success = 2xx AND no
'  "error" key in the body:
'  this API returns HTTP 200 with an error body on a bad request.
'
'  Geometry and loads are both parametric, computed from an Excel INPUT
'  sheet (see the two INPUT blocks below). Run BuildWingwallModel() against
'  an EMPTY model - PUT silently overwrites nodes 1-20, elements 1-9,
'  MATL/THIK 1-2 if anything else is using those numbers.
' ============================================================================


' ---------------------------------------------------------------------------
'  CONFIG
' ---------------------------------------------------------------------------

' Bumped on EVERY edit to this file. Printed in the report title,
' because these modules are pasted into Excel by hand: the file in git
' and the code actually running can silently diverge, and a fix that
' looks ineffective is very often just not re-imported yet. Check this
' matches before diagnosing anything from a report screenshot.
Private Const SCRIPT_VERSION As String = "2026-09-23e"

' One-line summary of what changed in THIS version, shown by the updater
' next to this module when it's stale. Update alongside SCRIPT_VERSION -
' must stay on ONE physical line (no "_" continuation - the parser that
' reads this out does not resolve continuations) and must not contain "|"
' (breaks manifest.txt's pipe-delimited format).
Private Const SCRIPT_CHANGELOG As String = "PostPlaneLoadTypes now actually PUTs to db/PNLD (was building the body and never sending it); VerifyEmpty checks all 12 cleaned endpoints via the confirmed-live empty signal."

' Identifies this module to the updater regardless of what it was
' named when pasted into Excel - these files carry no VB_Name, so the
' module name in the VBA project is whatever the user typed.
Private Const SCRIPT_ID As String = "wingwall-model-build"

' DESTRUCTIVE, and on by default. BuildWingwallModel deletes everything
' it is about to write before writing it, so every PUT lands in an empty
' endpoint and key/name numbering cannot disagree with whatever was
' there before. Everything deleted is rewritten immediately afterwards.
Private Const CLEAN_BEFORE_BUILD As Boolean = True

' Base URL from Civil NX: Tools > API > API Setting
Private Const API_BASE_URL As String = "https://moa-engineers.midasit.com:443/civil"

' Module-level declarations (Const/Dim/Private) must all sit here, before
' any Sub/Function - VBA rejects one appearing between two procedures
' ("Only comments may appear after End Sub...", confirmed live).
Private Const PI_CONST As Double = 3.14159265358979

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

' Project unit system (db/UNIT), set once before everything else. Matches
' KD.txt's *UNIT line (KN, M, KCAL, C).
Private Const PROJECT_FORCE_UNIT As String = "KN"
Private Const PROJECT_DIST_UNIT As String = "M"
Private Const PROJECT_HEAT_UNIT As String = "KCAL"
Private Const PROJECT_TEMPER_UNIT As String = "C"

' Project Information (db/PJCF). USER/ADDRESS stay fixed; the project
' name is assembled off the INPUT sheet as  G8 & " - " & G9  (the sheet
' labels them KM and ANO in E8/E9), e.g. "151A - P1", and written to
' both PROJECT and TITLE - Civil NX's Project Information dialog shows
' Title separately and leaves it blank otherwise.
' ENGINEER (K17) and REVISION (K18) are written only when non-blank:
' both are Optional per the db/PJCF doc and not every build has a
' reviewer or a revision tag yet. A missing project NAME warns rather
' than blocking the build.
Private Const PROJINFO_USER As String = "Deha"
Private Const PROJINFO_ADDRESS As String = "DEHA"
Private Const PROJECT_SHEET_NAME As String = "INPUT"
Private Const CELL_PROJECT_NO As String = "G8"
Private Const CELL_PROJECT_ANO As String = "G9"
Private Const CELL_PROJECT_ENGINEER As String = "K17"
Private Const CELL_PROJECT_REVISION As String = "K18"
Private PROJECT_NAME As String
Private PROJECT_ENGINEER As String
Private PROJECT_REVISION As String

' Structure Type (db/STYP), per KD.txt's *STRUCTYPE line:
'   "0, 1, 1, NO, YES, 9.806, 0, NO, NO, NO, 1"
'   iSTYP=0(3D), iMASS=1(Lumped), iSMAS=1, bMASSOFFSET=NO, bSELFWEIGHT=YES,
'   GRAV=9.806, TEMPER=0, bALIGNBEAM=NO, bALIGNSLAB=NO, bROTRIGID=NO
Private Const STRUCTYPE_TYPE As Long = 0          ' 3-D
Private Const STRUCTYPE_MASS As Long = 1          ' Lumped Mass
Private Const STRUCTYPE_SMASS As Long = 1         ' Convert self-weight to X,Y,Z
Private Const STRUCTYPE_GRAV As Double = 9.806
Private Const STRUCTYPE_TEMP As Double = 0
Private Const STRUCTYPE_MASSOFFSET As Boolean = False
Private Const STRUCTYPE_SELFWEIGHT As Boolean = True
Private Const STRUCTYPE_ALIGNBEAM As Boolean = False
Private Const STRUCTYPE_ALIGNSLAB As Boolean = False
Private Const STRUCTYPE_ROTRIGID As Boolean = False

' Main Control Data (db/ACTL) - see caveat in the header comment above:
' these are the JSON Manual's own example defaults, NOT derived from
' KD.txt's *ANAL-CTRL block (a different, unrelated feature).
Private Const MAIN_CTRL_ARDC As Boolean = True
Private Const MAIN_CTRL_ANRC As Boolean = True
Private Const MAIN_CTRL_ITER As Long = 20
Private Const MAIN_CTRL_TOL As Double = 0.001
Private Const MAIN_CTRL_CSECF As Boolean = False
Private Const MAIN_CTRL_TRS As Boolean = True
Private Const MAIN_CTRL_CRBAR As Boolean = False
Private Const MAIN_CTRL_BMSTRESS As Boolean = False
Private Const MAIN_CTRL_CLATS As Boolean = False

' Single material, shared by foundation and stem (only thickness differs
' between them - see SECT_FOUNDATION/SECT_STEM below). User-defined
' isotropic concrete (P_TYPE 2), matching midas-culvert-model-build.bas's
' *MATERIAL row: "1, CONC, C30/37, 0, 0, , C, NO, 0.05, 2, <Ec>, 0.2,
' 1.0000e-05, 25, 0". Ec comes from "INPUT"!B19 (CELL_MATERIAL_ELAST,
' user-specified 2026-09-22); POISN/THERMAL/DEN/MASS are fixed, matching
' that row - DEN isn't read from a cell here (only Ec was asked for).
Private Const MATERIAL_NO As Long = 1
Private Const MATERIAL_NAME As String = "C30/37"
Private Const MATERIAL_DAMP_RATIO As Double = 0.05
Private Const MATERIAL_POISN As Double = 0.2
Private Const MATERIAL_THERMAL As Double = 0.00001
Private Const MATERIAL_DEN As Double = 25
Private Const MATERIAL_MASS As Double = 0
Private Const MATERIAL_ELAST_DEFAULT As Double = 26291000#  ' fallback if B19 is blank/invalid
Private Const CELL_MATERIAL_ELAST As String = "B19"
Private MATERIAL_ELAST As Double     ' "INPUT"!B19 (Ec), falls back to MATERIAL_ELAST_DEFAULT

' Two THIK records - foundation and stem have independent thicknesses.
Private Const SECT_FOUNDATION As Long = 1
Private Const SECT_STEM As Long = 2

' Passed to ParseElementsBySect to mean "every plate, whatever its
' thickness" - used for the before/after snapshots around a divide, which
' have to see walls and foundation alike.
Private Const SECT_ANY As Long = -1

' Thickness colours (db/CO_T) - the plate analogue of the culvert
' builder's db/CO_S, same field layout, keyed by THICKNESS number:
'   W_*  = WireFrame  (the 2D line view)
'   HF_* = HiddenFill (the 3D rendered surface)
'   HE_* = HiddenEdge (the 3D rendered edges)
' Wireframe and hidden edge take the base colour, hidden fill a
' THICK_FILL_LIGHTEN tint of it. Brown foundation / green stem, matching
' the culvert builder's own foundation and wall colours.
' GET and PUT are its only active methods - no DELETE, which is why
' RunCleanSteps leaves it alone; colours are simply overwritten.
Private Const THICK_COLOR_LIST As String = _
    "1|110,90,70;" & _
    "2|60,160,75"

' The 1.44 is OUR choice, not the doc's arithmetic - the manual's own
' example pair back-solves to 1.4324/1.4437/1.4396 per channel, i.e. a
' MIDAS default rather than a rule. Channels round, then clamp at 255.
Private Const THICK_FILL_LIGHTEN As Double = 1.44

' Opacity. False = solid; FACT only applies when bBLEMD is True.
Private Const THICK_COLOR_TRANSLUCENT As Boolean = False
Private Const THICK_COLOR_OPACITY As Double = 0.5


' ---------------------------------------------------------------------------
'  GEOMETRY INPUT (INPUT sheet)
'    B20 = INSIDE (clear) opening width, wall face to wall face along the
'          culvert face. NOT centreline spacing - GenerateGeometry converts.
'    B21 = stem thickness      B24 = foundation thickness
'    Row 22 = LEFT wingwall, row 23 = RIGHT: B=length, C=splay angle (deg),
'             D=height at culvert face (Hnear), E=height at tip (Hfar).
'  LEFT's own edge is placed on the global +X axis, so which row is "LEFT"
'  also sets the model's orientation. Flip the two CELL_ blocks below if the
'  wingwalls come out exchanged.
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

' Stem rigid zone height, and foundation rigid-strip width (both sides of
' each wall's reference line), both = this ratio x STEM thickness
' (confirmed by user: 0.5, matching KD.txt's 0.4m thickness -> 0.2m zone).
Private Const RIGID_ZONE_RATIO As Double = 0.5

' Element subtype applied to every generated plate. Thick+Drilling=3, per
' the Element doc's STYPE table - matches most of KD.txt's elements, but
' is a default here, not re-derived from the new geometry/inputs.
Private Const GEN_STYPE As Long = 3

Private Type SideParams
    L As Double
    AngleDeg As Double
    Hnear As Double
    Hfar As Double
End Type

' Filled in at runtime by GenerateGeometry (called from BuildWingwallModel
' before PostThickness/PostNodes/PostElements). No longer Const, unlike
' every other *_LIST in this file.
Private THICKNESS_FOUNDATION_VALUE As Double
Private THICKNESS_STEM_VALUE As Double
Private NODE_LIST As String
Private ELEMENT_LIST As String

' Also filled in by GenerateGeometry: the wall geometry the LOAD builders
' need. The plane-load steps have to place an area load on each wingwall
' face and project it onto that wall's own plane, which needs real
' coordinates and heights - NODE_LIST alone would mean re-parsing strings
' back into numbers. "Near"/"far" are the culvert-face and tip ends.
Private GEO_LEFT_NEAR_X As Double, GEO_LEFT_NEAR_Y As Double
Private GEO_LEFT_FAR_X As Double, GEO_LEFT_FAR_Y As Double
Private GEO_LEFT_L As Double, GEO_LEFT_HNEAR As Double, GEO_LEFT_HFAR As Double
Private GEO_RIGHT_NEAR_X As Double, GEO_RIGHT_NEAR_Y As Double
Private GEO_RIGHT_FAR_X As Double, GEO_RIGHT_FAR_Y As Double
Private GEO_RIGHT_L As Double, GEO_RIGHT_HNEAR As Double, GEO_RIGHT_HFAR As Double


' ---------------------------------------------------------------------------
'  PROGRESS BAR - Excel's own status bar, so this module stays a single
'  pasteable .bas (a UserForm would need .frm/.frx alongside it). The
'  step counts below are asserted against the actual call sites by
'  tests/verify_progress_steps.py - update them together.
' ---------------------------------------------------------------------------
Private Const PROGRESS_BAR_WIDTH As Long = 20
Private Const PROGRESS_STEPS_BUILD As Long = 21
Private Const PROGRESS_STEPS_CLEAN As Long = 12
Private Const PROGRESS_STEPS_VERIFY As Long = 12
Private PROGRESS_STEP As Long
Private PROGRESS_TOTAL As Long

' ---------------------------------------------------------------------------
'  ELEMENT DIVISION (ope/DIVIDEELEM)
'
'  Two mesh counts per structure part, off the INPUT sheet - row 17 heads
'  the columns (F = DUVAR/wall, G = TEMEL/foundation), row 18 is MESHX and
'  row 19 MESHY. Walls therefore mesh independently of the foundation.
'
'  DIVIDE_TABLE says, per element, which of those counts goes on the
'  element's LOCAL x and y - mirroring the table the owner laid out at
'  INPUT!T5:V14. A literal 1 means that direction is not divided. Note the
'  walls (5-8) deliberately SWAP the two: their local x runs up the height.
'
'      elem | x     | y          elem | x     | y
'        1  | MESHX | 1            6  | 1     | MESHX
'        2  | MESHX | MESHY        7  | MESHY | MESHX
'        3  | MESHX | 1            8  | 1     | MESHX
'        4  | MESHX | 1            9  | MESHX | 1
'        5  | MESHY | MESHX
'
'  The table is held HERE rather than read from those cells because the
'  element numbering it keys on is itself hardcoded, in GenerateGeometry -
'  a sheet table that disagreed with the built topology would divide the
'  wrong plates silently. Editing INPUT!T5:V14 alone changes nothing; edit
'  both, together.
'
'  Division runs AFTER the loads and combinations, matching the culvert
'  builder. CONFIRMED LIVE 2026-09-22: Civil NX carries the pressure loads
'  on plates 5-8 onto their sub-elements, so nothing has to be reassigned
'  afterwards and this step does not need to move ahead of
'  PostPressureLoads.
'
'  START_NUMBER is deliberately omitted, as the doc's own examples omit it.
'  Since elements start numbered 1..9 with no gaps, and a divide reuses the
'  original number for one child, there is never a gap for a new element to
'  fill - so elements not yet divided keep their numbers either way, which
'  is what lets this walk the table by element number.
' ---------------------------------------------------------------------------
Private Const CELL_MESH_X_WALL As String = "F18"
Private Const CELL_MESH_Y_WALL As String = "F19"
Private Const CELL_MESH_X_FOUND As String = "G18"
Private Const CELL_MESH_Y_FOUND As String = "G19"

' "elemNo|xSpec|ySpec", spec = MESHX, MESHY or a literal count.
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

Private MESH_X_WALL As Long, MESH_Y_WALL As Long
Private MESH_X_FOUND As Long, MESH_Y_FOUND As Long

' Every element as it stood just before a tracked plate was divided, and
' the non-rigid families accumulated by diffing against that snapshot.
' See PostDivideElements.
Private SNAP_IDS() As Long, SNAP_COUNT As Long
Private NR_FOUND_IDS() As Long, NR_FOUND_COUNT As Long
Private NR_WALL_IDS() As Long, NR_WALL_COUNT As Long

' Foundation and wall plates read back after the divide, shared by the
' element-list write-back and the surface springs so db/ELEM is fetched
' once rather than twice.
Private FOUND_IDS() As Long, FOUND_COUNT As Long
Private WALL_IDS() As Long, WALL_COUNT As Long
Private ELEM_SETS_READY As Boolean

' ---------------------------------------------------------------------------
'  FOUNDATION SURFACE SPRINGS (ope/SSPS)
'  B8 on the INPUT sheet is the VERTICAL subgrade modulus kv (kN/m3,
'  labelled ZEMIN DUSEY YAY KATSAYISI). The horizontal directions get
'  SPRING_H_RATIO x kv, mirroring what the culvert builder does with its
'  db/NSPR point springs.
'
'  Unlike the culvert - which discovers foundation nodes itself and
'  computes a tributary length per node - ope/SSPS takes the PLATES and
'  does the area conversion inside MIDAS, so no node discovery is needed.
'
'  The target is EVERY foundation plate AFTER the divide, opening strip
'  included (all of them bear on soil at Z=0) - read back from db/ELEM
'  into FOUND_IDS, not hardcoded, because the divide renumbers nothing but
'  adds a sub-element per cell.
'
'  UNVERIFIED LIVE: with BOUNDARY.TYPE "LINEAR" the doc calls STIFF
'  simply "Stiffness [Kx, Ky, Kz]" while the COMP/TENS forms name their
'  single value "Modulus of Subgrade Reaction". Civil NX's Surface Spring
'  dialog labels the linear boxes as moduli, which is what kv is - but
'  check the resulting db/NSPR values against kv x tributary area on the
'  first live run before trusting the magnitudes.
' ---------------------------------------------------------------------------
Private Const CELL_SUBGRADE_MODULUS As String = "B8"
Private Const SPRING_H_RATIO As Double = 0.5
Private Const SPRING_GROUP_NAME As String = "TEMEL_YAY"
Private SPRING_KV As Double

' ---------------------------------------------------------------------------
'  NAMED UCS (db/NUCS) - the coordinate system the FOUNDATION plate-force
'  pictures are reported in by midas-wingwall-plate-forces-capture.bas.
'
'  Why: local plate axes are not comparable across the foundation. Local x
'  runs node1 -> node2, and the five strips are wound per side, so "Mxx" on
'  a LEFT strip and "Mxx" on a RIGHT strip point about (angleL + angleR)
'  apart. A single UCS gives one grid for the whole slab.
'
'  Orientation, per the owner: origin at the global origin (node 7), X
'  rotated by the LEFT wall's own splay angle - the LEFT wall lies on global
'  +X, so rotating by C22 puts UCS x at that angle and UCS y at 90 + C22,
'  which is exactly along the culvert face A->B. So the foundation reads
'  "along the face" and "across the face".
'
'      VX = ( cos a,  sin a, 0)      VY = (-sin a,  cos a, 0)
'
'  CONFIRMED LIVE 2026-09-22 against the model, not guessed:
'    GET  info/db/NUCS -> NAME (string), ORG_ITEM/VX_ITEM/VY_ITEM (3 reals)
'    PUT  db/NUCS with the usual {"Assign": {"1": {...}}} -> HTTP 200, and a
'         GET reads the record straight back.
'  Note info/db/* wraps every record shape in "Argument" - db/NODE does the
'  same and definitely takes "Assign", so that is NOT a different body form.
'
'  This endpoint is absent from the cached API Online Manual's catalogue,
'  which predates it; its own page is "UCS"/[NUCS] Named UCS on the support
'  site. Looking for "UCS" in the endpoint list finds nothing - it is NUCS.
'
'  STILL OPEN: whether the capture's UCS_NAME accepts this name. The Plate
'  Forces doc's enum lists only the literal "CurrentUCS", and nothing
'  readable says which UCS is current - there is no info/view/* endpoint
'  (404) and GET view/RESULTGRAPHIC returns "error status" without results.
'  The capture script sends the name first; if the foundation pictures come
'  back unrotated, select this UCS by hand in Civil NX and switch that
'  script's UCS_NAME_FOR_RESULTS back to "CurrentUCS".
' ---------------------------------------------------------------------------
Private Const UCS_NAME As String = "FOUND"
Private GEO_LEFT_ANGLE_DEG As Double

' ---------------------------------------------------------------------------
'  ELEMENT LISTS written BACK to the INPUT sheet after the build. The
'  wingwall CAPTURE scripts already READ G10 (foundation) and G11 (wall)
'  to drive their ACTIVE element isolation, so filling these in here is
'  what stops them needing a hand edit after every rebuild.
'
'    G10  FOUND      - the NON-RIGID foundation plates
'    G11  WALL       - the NON-RIGID stem plates
'    G12  FOUND PRR  - every foundation plate, rigid ones included
'
'  G10 and G11 leave the rigid zones out, the same way the culvert
'  builder's own element list does, because a rigid member's forces are
'  not worth displaying. Of the nine as-built plates:
'
'    foundation  1, 9  half-bands under the RIGHT wall   RIGID
'                3, 4  half-bands under the LEFT wall    RIGID
'                2     the clear span between them       -> G10
'    stem        6, 8  rigid bands at the wall bases     RIGID
'                5, 7  the sloped wall bands             -> G11
'
'  Confirmed against the values this workbook carried from an earlier,
'  hand-built model: G10 held 64 elements (one 8x8 plate's worth, i.e.
'  element 2 alone), G11 held 128 (elements 5 and 7), and G12 held all 96.
'
'  Those old values were also NOT contiguous ranges, which is the whole
'  reason PostDivideElements diffs db/ELEM around each tracked divide
'  rather than assuming MIDAS hands out sub-element numbers in one block.
' ---------------------------------------------------------------------------
Private Const CELL_ELEM_LIST_FOUND As String = "G10"
Private Const CELL_ELEM_LIST_WALL As String = "G11"
Private Const CELL_ELEM_LIST_FOUND_PRR As String = "G12"

' As-built plate numbers whose families make up G10 and G11.
Private Const NONRIGID_FOUND_ELEMS As String = "2"
Private Const NONRIGID_WALL_ELEMS As String = "5,7"


' Static load cases: "NAME|TYPE;..." - TYPE per Static Load Cases doc's
' available-type table (D, EH, LS, E, ...).
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


' ---------------------------------------------------------------------------
'  LOAD VALUES (INPUT sheet) - only the MAGNITUDES are entered. Everything
'  geometric about a load is derived from the generated model.
'
'  Columns are the four corners of the trapezoidal wall face, where NEAR =
'  culvert-face end (height Hnear, the TALL end) and FAR = wall tip (Hfar):
'
'        E (top near)  o---___
'                      |       ---___ D (top far)
'        B (bot near)  o-------------o  C (bot far)
'                   culvert face    wall tip
'
'         |    B     |    C    |    D    |    E
'     27  | bot near | bot far | top far | top near   LEFT  at-rest  EHS2_L
'     28  |                                           LEFT  active   EHA2_L
'     29  |                                           LEFT  seismic  EQ_L
'     30  |                                           RIGHT at-rest  EHS2_R
'     31  |                                           RIGHT active   EHA2_R
'     32  |                                           RIGHT seismic  EQ_R
'     34  | LEFT LS  | RIGHT LS|                      surcharge on the stem
'     35  | seismic inertia coefficient (shared: ATA_L/ATA_R are the two
'          directions of ONE ground motion, not two walls' loads)
'
'  Triangular earth pressure puts the values in B and C with D/E zero, and
'  B > C since the near end is taller. KD.txt's reference numbers:
'    at-rest 24.73/8.95/0/0   active 14.98/5.42/0/0
'    seismic 3.13/3.13/13.14/36.29 (heaviest at the TOP)
'    surcharge -8.53   seismic coef 0.458
'
'  *** Absolute intensities (kN/m2), NOT per-metre-of-height coefficients -
'  they do NOT rescale if you change Hnear/Hfar. *** See CLAUDE.md.
' ---------------------------------------------------------------------------
Private Const CELL_EP_LEFT_ATREST_ROW As Long = 27
Private Const CELL_EP_LEFT_ACTIVE_ROW As Long = 28
Private Const CELL_EP_LEFT_SEISMIC_ROW As Long = 29
Private Const CELL_EP_RIGHT_ATREST_ROW As Long = 30
Private Const CELL_EP_RIGHT_ACTIVE_ROW As Long = 31
Private Const CELL_EP_RIGHT_SEISMIC_ROW As Long = 32
Private Const CELL_SURCHARGE_LEFT As String = "B34"
Private Const CELL_SURCHARGE_RIGHT As String = "C34"
Private Const CELL_SEISMIC_COEF As String = "B35"

' Per-side sign. An area load acts NORMAL to its own plane and the two
' walls' planes face opposite ways, so one side must be negated for both to
' push inward. KD.txt negates its LEFT ("uzun") side; matched here. Stays a
' constant rather than a cell - it's a modelling convention, not a load.
' VERIFY LIVE - the normal follows from the plane's origin/axis points, so
' this is exactly the kind of thing that silently flips when geometry moves.
Private Const EP_SIGN_LEFT As Double = -1
Private Const EP_SIGN_RIGHT As Double = 1

' Stem element numbers per side - the elements the surcharge pushes on.
' These follow THIS script's own topology (see GenerateGeometry's element
' numbering); KD.txt's equivalents were 36/37 (LEFT) and 34/35 (RIGHT).
Private Const ELEM_STEM_LEFT_SLOPED As Long = 7
Private Const ELEM_STEM_LEFT_RIGID As Long = 8
Private Const ELEM_STEM_RIGHT_SLOPED As Long = 5
Private Const ELEM_STEM_RIGHT_RIGID As Long = 6

' Filled in at runtime by ReadLoadInputs. Index order matches the sheet
' columns: 1 = bottom near, 2 = bottom far, 3 = top far, 4 = top near.
Private LOAD_L_ATREST(1 To 4) As Double
Private LOAD_L_ACTIVE(1 To 4) As Double
Private LOAD_L_SEISMIC(1 To 4) As Double
Private LOAD_R_ATREST(1 To 4) As Double
Private LOAD_R_ACTIVE(1 To 4) As Double
Private LOAD_R_SEISMIC(1 To 4) As Double
Private LOAD_SURCHARGE_LEFT As Double
Private LOAD_SURCHARGE_RIGHT As Double
Private LOAD_SEISMIC_COEF As Double

' Load combinations (db/LCOM-GEN):
'   "NAME|ACTIVE|ITYPE|ANAL:LCNAME:FACTOR,...|TIER;..."
' ITYPE per doc: Add=0, Envelope=1. ANAL: ST=static load case, CB=general
' combination (nested envelope-of-envelopes, as in KD.txt's ENV_ALL).
'
' TIER is the dependency level, and it is NOT cosmetic. A "CB" entry
' references another COMBINATION, which must already exist when the
' referring one is written - and sending them all in one Assign object
' makes that depend on the order the server walks the object's KEYS,
' which JSON does not guarantee (walked as strings they come out
' "1","10","11",...,"2",...). The culvert builder failed exactly this way,
' with an error MIDAS could not even name the offending combination in.
' PostLoadCombinations therefore sends one PUT per tier:
'   1 - references static load cases only
'   2 - references tier-1 combinations
'   3 - references tier-2 combinations
' Keys stay globally unique and in definition order across the tiers, so
' the numbering is exactly what a single request would have produced.
Private Const LOADCOMB_LIST As String = _
    "SLS|ACTIVE|0|ST:DL:1,ST:EHS2_L:1,ST:EHS2_R:1,ST:LSS1_L:1,ST:LSS1_R:1|1;" & _
    "ULS|ACTIVE|0|ST:DL:1.35,ST:EHS2_L:1.35,ST:EHS2_R:1.35,ST:LSS1_L:1.45,ST:LSS1_R:1.45|1;" & _
    "EQ - 1|ACTIVE|0|ST:DL:1,ST:EHA2_L:1,ST:EQ_L:1,ST:ATA_L:1|1;" & _
    "EQ - 2|ACTIVE|0|ST:DL:1,ST:EHA2_R:1,ST:EQ_R:1,ST:ATA_R:1|1;" & _
    "ENV_SER|ACTIVE|1|CB:SLS:1|2;" & _
    "ENV_STR|ACTIVE|1|CB:ULS:1|2;" & _
    "ENV_EQ|ACTIVE|1|CB:EQ - 1:1,CB:EQ - 2:1|2;" & _
    "ENV_ALL|ACTIVE|1|CB:ENV_SER:1,CB:ENV_STR:1,CB:ENV_EQ:1|3"


' ===========================================================================
'  MAIN - runs every step in order, aborting the remaining steps (but still
'  reporting what succeeded) on the first hard failure.
' ===========================================================================
Sub BuildWingwallModel()

    Dim report As String
    Dim cleanReport As String
    Dim ok As Boolean

    report = "Wingwall model build  [" & SCRIPT_VERSION & "]" & vbCrLf & _
             String(40, "-") & vbCrLf

    Call ResetProgress(PROGRESS_STEPS_BUILD + _
         IIf(CLEAN_BEFORE_BUILD, PROGRESS_STEPS_CLEAN + PROGRESS_STEPS_VERIFY, 0))

    ' Wipe first, so every PUT below lands in an empty endpoint and the
    ' key/name numbering cannot disagree with whatever was there before.
    ' Deliberately NOT gated on success: a FAIL here usually just means the
    ' endpoint was already empty, which is the state we wanted.
    '
    ' The READ-BACK is the part that matters. A keyless DELETE clearing a
    ' whole endpoint has been an untested assumption in this file since it
    ' was written; VerifyEmpty settles it per run instead of leaving the
    ' build to fail four steps later with a vague duplicate-name message.
    If CLEAN_BEFORE_BUILD Then
        cleanReport = ""
        Call RunCleanSteps(cleanReport)
        ' Every endpoint RunCleanSteps deletes gets read back, not just the
        ' two with model-wide-unique names - see VerifyEmpty's own comment
        ' for why a NAME-only check missed half of these. db/PNLD and
        ' db/PNLA get it for the same reason culvert's LCOM-GEN/STLD do:
        ' PNLD's keys can renumber when the endpoint isn't actually empty,
        ' and PNLA binds every area load to its type by bare integer
        ' PNLD_KEY with nothing else to catch a silent mismatch.
        Call StepResult(cleanReport, Progress("verify Load Combinations cleared"), _
                        VerifyEmpty("db/LCOM-GEN", "Load combinations"))
        Call StepResult(cleanReport, Progress("verify Plane Loads cleared"), _
                        VerifyEmpty("db/PNLA", "Plane loads"))
        Call StepResult(cleanReport, Progress("verify Plane Load Types cleared"), _
                        VerifyEmpty("db/PNLD", "Plane load types"))
        Call StepResult(cleanReport, Progress("verify Pressure Loads cleared"), _
                        VerifyEmpty("db/PRES", "Pressure loads"))
        Call StepResult(cleanReport, Progress("verify Self-Weight cleared"), _
                        VerifyEmpty("db/BODF", "Self-weight"))
        Call StepResult(cleanReport, Progress("verify Static Load Cases cleared"), _
                        VerifyEmpty("db/STLD", "Static load cases"))
        Call StepResult(cleanReport, Progress("verify Point Springs cleared"), _
                        VerifyEmpty("db/NSPR", "Point springs"))
        Call StepResult(cleanReport, Progress("verify Elements cleared"), _
                        VerifyEmpty("db/ELEM", "Elements"))
        Call StepResult(cleanReport, Progress("verify Nodes cleared"), _
                        VerifyEmpty("db/NODE", "Nodes"))
        Call StepResult(cleanReport, Progress("verify Thickness cleared"), _
                        VerifyEmpty("db/THIK", "Thickness"))
        Call StepResult(cleanReport, Progress("verify Material cleared"), _
                        VerifyEmpty("db/MATL", "Material"))
        Call StepResult(cleanReport, Progress("verify Named UCS cleared"), _
                        VerifyEmpty("db/NUCS", "Named UCS"))

        ' Keep the report inside VBA's ~1024-char MsgBox limit: one line when
        ' the clean pass was clean, the full text only when it was not.
        If InStr(cleanReport, "FAIL") > 0 Or InStr(cleanReport, "WARN") > 0 Then
            report = report & "Clean pass (CLEAN_BEFORE_BUILD):" & vbCrLf & _
                     cleanReport & String(40, "-") & vbCrLf
        Else
            report = report & "Clean pass (CLEAN_BEFORE_BUILD): OK" & vbCrLf & _
                     String(40, "-") & vbCrLf
        End If
    End If

    ok = StepResult(report, Progress("Unit System"), PostUnitSystem())
    If ok Then ok = StepResult(report, Progress("Project Information"), PostProjectInfo())
    If ok Then ok = StepResult(report, Progress("Structure Type"), PostStructureType())
    If ok Then ok = StepResult(report, Progress("Main Control Data"), PostMainControlData())
    If ok Then ok = StepResult(report, Progress("Geometry Setup (read inputs + compute)"), PostGeometrySetup())
    If ok Then ok = StepResult(report, Progress("Material"), PostMaterial())
    If ok Then ok = StepResult(report, Progress("Thickness"), PostThickness())
    If ok Then ok = StepResult(report, Progress("Thickness Colours"), PostThicknessColors())
    If ok Then ok = StepResult(report, Progress("Named UCS"), PostNamedUcs())
    If ok Then ok = StepResult(report, Progress("Nodes"), PostNodes())
    If ok Then ok = StepResult(report, Progress("Elements"), PostElements())

    ' Load steps. These run AFTER the geometry steps above, and must: the
    ' plane-load builders read the GEO_* variables that GenerateGeometry
    ' fills in, and every load targets element/node numbers that only exist
    ' once Nodes/Elements have posted.
    If ok Then ok = StepResult(report, Progress("Static Load Cases"), PostStaticLoadCases())
    If ok Then ok = StepResult(report, Progress("Self-Weight"), PostSelfWeight())
    If ok Then ok = StepResult(report, Progress("Pressure Loads"), PostPressureLoads())
    If ok Then ok = StepResult(report, Progress("Plane Load Types"), PostPlaneLoadTypes())
    If ok Then ok = StepResult(report, Progress("Plane Loads"), PostPlaneLoadAssignments())
    If ok Then ok = StepResult(report, Progress("Load Combinations"), PostLoadCombinations())

    ' Mesh, then everything that has to see the meshed model: the element
    ' lists are read back out of it, and the springs land on the foundation
    ' SUB-elements (per the owner: springs after dividing). Analysis last.
    If ok Then ok = StepResult(report, Progress("Divide Elements"), PostDivideElements())
    If ok Then ok = StepResult(report, Progress("Element Lists"), PostElementLists())
    If ok Then ok = StepResult(report, Progress("Foundation Springs"), PostFoundationSprings())
    If ok Then ok = StepResult(report, Progress("Perform Analysis"), PostPerformAnalysis())

    If ok Then
        report = report & vbCrLf & "All steps completed."
    Else
        report = report & vbCrLf & "Stopped after the first failed step - fix it and re-run."
    End If

    ' Before the MsgBox, and on the failure path too - see ClearProgress.
    Call ClearProgress

    MsgBox report, IIf(ok, vbInformation, vbExclamation)

End Sub

' Deletes everything the build script writes, so BuildWingwallModel can be
' re-run against a model that already holds a prior wingwall build instead
' of hitting "Key Already Exist"/"Duplicate Name" collisions. Run this by
' hand when CLEAN_BEFORE_BUILD is off, or to empty a model without
' rebuilding it.
Sub CleanWingwallModel()

    Dim report As String
    report = "Wingwall model clean  [" & SCRIPT_VERSION & "]" & vbCrLf & _
             String(40, "-") & vbCrLf

    Call ResetProgress(PROGRESS_STEPS_CLEAN)
    Call RunCleanSteps(report)
    Call ClearProgress

    report = report & vbCrLf & "Clean pass done - a FAIL on an endpoint that was already " & _
             "empty is expected and harmless. Review any other FAILs before running " & _
             "BuildWingwallModel()."

    MsgBox report, vbInformation

End Sub

' The delete sequence, shared by CleanWingwallModel and by the
' CLEAN_BEFORE_BUILD pass at the top of BuildWingwallModel.
'
' Order is the REVERSE of the build order - loads before the load cases and
' elements they reference, elements before nodes, THIK/MATL last since ELEM
' assigns them by number. No step gates on another: a FAIL on an
' already-empty endpoint is expected and harmless, so every delete is
' attempted regardless and the report is there to review.
'
' db/NSPR is listed even though deleting db/NODE clears the point springs
' with it - the springs are converted from the plates by ope/SSPS, and an
' explicit delete costs one request and removes the ordering assumption.
'
' db/CO_T is deliberately absent: GET and PUT are its only active methods,
' so there is nothing to delete. Colours are simply overwritten.
'
' UNVERIFIED LIVE: every endpoint below is documented "Active Methods:
' POST, GET, PUT, DELETE", so DELETE itself is real - but no doc page shows
' a worked DELETE JSON example the way they do for POST/PUT. Whether a
' keyless DELETE (this script's assumption, extending the documented GET
' convention: no key = whole endpoint) wipes every record, or needs each key
' named, is what the VerifyEmpty read-back in BuildWingwallModel checks.
Private Sub RunCleanSteps(ByRef report As String)

    Call StepResult(report, Progress("delete Load Combinations"), DeleteAndCheck("db/LCOM-GEN"))
    Call StepResult(report, Progress("delete Plane Loads"), DeleteAndCheck("db/PNLA"))
    Call StepResult(report, Progress("delete Plane Load Types"), DeleteAndCheck("db/PNLD"))
    Call StepResult(report, Progress("delete Pressure Loads"), DeleteAndCheck("db/PRES"))
    Call StepResult(report, Progress("delete Self-Weight"), DeleteAndCheck("db/BODF"))
    Call StepResult(report, Progress("delete Static Load Cases"), DeleteAndCheck("db/STLD"))
    Call StepResult(report, Progress("delete Point Springs"), DeleteAndCheck("db/NSPR"))
    Call StepResult(report, Progress("delete Elements"), DeleteAndCheck("db/ELEM"))
    Call StepResult(report, Progress("delete Nodes"), DeleteAndCheck("db/NODE"))
    Call StepResult(report, Progress("delete Thickness"), DeleteAndCheck("db/THIK"))
    Call StepResult(report, Progress("delete Material"), DeleteAndCheck("db/MATL"))
    Call StepResult(report, Progress("delete Named UCS"), DeleteAndCheck("db/NUCS"))

End Sub


' ---------------------------------------------------------------------------
'  PROGRESS
' ---------------------------------------------------------------------------

Private Sub ResetProgress(ByVal total As Long)
    PROGRESS_STEP = 0
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

' Hands the status bar back to Excel. Must run on EVERY exit path - a status
' bar left set stays stuck there for the rest of the session.
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
'     StepResult(report, Progress("Nodes"), PostNodes())
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
    Application.StatusBar = "MIDAS wingwall  [" & _
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
'
' Checks against the CONFIRMED-LIVE empty signal (an empty endpoint
' answers a bare {"message": ""} to every db/* read - see CLAUDE.md,
' 2026-09-22) rather than searching for a "NAME" field. That NAME-search
' version only worked for endpoints whose records happen to carry a NAME
' - db/LCOM-GEN, db/STLD, db/PNLD, db/MATL, db/THIK, db/NUCS - and stayed
' silently "clean" on every endpoint without one (db/PNLA, db/PRES,
' db/BODF, db/NSPR, db/ELEM, db/NODE: half of what RunCleanSteps
' deletes), which is exactly the gap that let a stale db/PNLD collision
' through undetected before this check covered every cleaned endpoint.
'
' Returns "" (treated as OK) when the endpoint reads back empty.
Private Function VerifyEmpty(ByVal path As String, ByVal what As String) As String

    Dim resp As String
    Dim statusCode As Long
    Dim t As String

    Call SendApiRequest("GET", path, "", resp, statusCode)

    t = resp
    t = Replace(t, " ", "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, vbTab, "")

    If InStr(1, t, """message"":""""", vbTextCompare) = 0 Then
        VerifyEmpty = "WARN: " & what & " STILL HOLD RECORDS after the delete - the " & _
            "keyless DELETE did not clear this endpoint. Clear it by hand in Civil NX " & _
            "and re-run; otherwise the writes below will collide on duplicate names."
    End If

End Function

' Appends one line to report and returns success, so the caller can chain
' "If ok Then ok = StepResult(...)" without repeating the MsgBox logic.
'
' A step can also return a "WARN: ..." string: logged, but NOT treated as a
' failure, so the remaining steps still run. Used for the re-run collisions
' that aren't actually harmful - see TolerateDuplicateName.
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

' db/STLD, db/PNLD and db/LCOM-GEN carry model-wide-unique NAMEs, so a
' re-run against an already-built model returns a name-collision error (as
' HTTP 200 with an error body) - worded "Duplicate Name" on STLD/PNLD, but
' "<thing> already exists" on LCOM-GEN. Downgraded to a warning: the loads
' reference their CASE by name (LCNAME - confirmed in the live schemas for
' db/BODF, db/PRES and db/PNLA), so pre-existing case names still leave the
' build valid. Does NOT check that the existing records' contents match.
'
' ONLY use this where the reference really is by name. db/PNLD is
' deliberately NOT wrapped in it: db/PNLA binds a type by integer PNLD_KEY,
' so tolerating a collision there silently mis-attaches every area load.
' See PostPlaneLoadTypes.
Private Function TolerateDuplicateName(ByVal result As String, ByVal what As String) As String

    If InStr(1, result, "Duplicate Name", vbTextCompare) > 0 _
        Or InStr(1, result, "already exists", vbTextCompare) > 0 Then
        TolerateDuplicateName = "WARN: " & what & " already exist in this model - kept " & _
            "the existing ones. Loads reference them by name, so the rest of the build " & _
            "is still valid, but verify their types/contents if this model wasn't empty."
    Else
        TolerateDuplicateName = result
    End If

End Function


' ===========================================================================
'  STEP FUNCTIONS - each returns "" on success, or an error message on
'  failure (so callers can both log and gate on it via StepResult).
' ===========================================================================

Private Function PostUnitSystem() As String

    Dim b As String
    b = "{""Assign"": {""1"": {"
    b = b & """FORCE"": """ & PROJECT_FORCE_UNIT & ""","
    b = b & """DIST"": """ & PROJECT_DIST_UNIT & ""","
    b = b & """HEAT"": """ & PROJECT_HEAT_UNIT & ""","
    b = b & """TEMPER"": """ & PROJECT_TEMPER_UNIT & """"
    b = b & "}}}"

    ' db/UNIT is GET/PUT only (per "Unit System" doc and CLAUDE.md) - not POST.
    PostUnitSystem = PostAndCheck("db/UNIT", b, "PUT")

End Function


Private Function PostProjectInfo() As String

    Dim b As String

    Call ReadProjectName

    If Len(PROJECT_NAME) = 0 Then
        PostProjectInfo = "WARN: no project name found at " & PROJECT_SHEET_NAME & "!" & _
            CELL_PROJECT_NO & "/" & CELL_PROJECT_ANO & _
            " - Project Information not written."
        Exit Function
    End If

    ' JsonEsc, because unlike every other string this module emits, these
    ' are free text off a worksheet.
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

    ' PUT, not POST: Civil NX pre-populates a Project Information record
    ' (key "1") by default in every model, so POST-to-create collides with
    ' it ("Key Already Exist", confirmed live). Same reasoning as db/UNIT
    ' and db/STYP above, even though the doc lists POST as an active method.
    PostProjectInfo = PostAndCheck("db/PJCF", b, "PUT")

End Function

' Project name = INPUT!G8 & " - " & INPUT!G9 (KM and ANO), e.g.
' "151A - P1", plus the engineer and revision cells. Nothing here is
' fatal: a missing sheet or blank cells just leave the parts that are
' present, and PostProjectInfo warns if the name comes out empty.
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


Private Function PostStructureType() As String

    Dim b As String
    b = "{""Assign"": {""1"": {"
    b = b & """STYP"": " & STRUCTYPE_TYPE & ","
    b = b & """MASS"": " & STRUCTYPE_MASS & ","
    b = b & """bMASSOFFSET"": " & LCase(STRUCTYPE_MASSOFFSET) & ","
    b = b & """bSELFWEIGHT"": " & LCase(STRUCTYPE_SELFWEIGHT) & ","
    b = b & """SMASS"": " & STRUCTYPE_SMASS & ","
    b = b & """GRAV"": " & JsonNum(STRUCTYPE_GRAV) & ","
    b = b & """TEMP"": " & JsonNum(STRUCTYPE_TEMP) & ","
    b = b & """bALIGNBEAM"": " & LCase(STRUCTYPE_ALIGNBEAM) & ","
    b = b & """bALIGNSLAB"": " & LCase(STRUCTYPE_ALIGNSLAB) & ","
    b = b & """bROTRIGID"": " & LCase(STRUCTYPE_ROTRIGID)
    b = b & "}}}"

    ' db/STYP is GET/PUT only (per "Structure Type" doc) - not POST.
    PostStructureType = PostAndCheck("db/STYP", b, "PUT")

End Function


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

    ' PUT, not POST - same "pre-populated singleton record" reasoning as
    ' db/PJCF above (untested for ACTL specifically, but same shape of risk).
    PostMainControlData = PostAndCheck("db/ACTL", b, "PUT")

End Function


' ---------------------------------------------------------------------------
'  PARAMETRIC GEOMETRY
' ---------------------------------------------------------------------------

' Reads inputs from INPUT_SHEET_NAME, then computes THICKNESS_FOUNDATION_
' VALUE/THICKNESS_STEM_VALUE/NODE_LIST/ELEMENT_LIST. Returns "" on success,
' or an error message (missing sheet, missing/non-numeric cell) on failure
' - it does NOT touch the API, so it can't itself return an HTTP-style
' failure.
Private Function PostGeometrySetup() As String

    Dim openingWidth As Double
    Dim leftSide As SideParams, rightSide As SideParams
    Dim errMsg As String

    errMsg = ReadWingwallInputs(openingWidth, THICKNESS_FOUNDATION_VALUE, THICKNESS_STEM_VALUE, _
                                 leftSide, rightSide)
    If Len(errMsg) > 0 Then
        PostGeometrySetup = errMsg
        Exit Function
    End If

    ' Read the load cells here too, not later next to the load steps, so a
    ' bad/empty load cell fails before ANY of it has posted - by the time
    ' the load steps run, nodes and elements are already in the model.
    errMsg = ReadLoadInputs()
    If Len(errMsg) > 0 Then
        PostGeometrySetup = errMsg
        Exit Function
    End If

    GenerateGeometry openingWidth, THICKNESS_STEM_VALUE, leftSide, rightSide

    PostGeometrySetup = ""

End Function

' Reads the LOAD VALUES cells (see that block for the sheet layout) into
' LOAD_*. Returns "" on success or an error message naming the bad cell.
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
        ReadLoadInputs = "Non-numeric or empty cell: " & CELL_SURCHARGE_LEFT & " (LEFT LS surcharge pressure)"
        Exit Function
    End If
    If Not TryReadCell(ws, CELL_SURCHARGE_RIGHT, LOAD_SURCHARGE_RIGHT) Then
        ReadLoadInputs = "Non-numeric or empty cell: " & CELL_SURCHARGE_RIGHT & " (RIGHT LS surcharge pressure)"
        Exit Function
    End If
    If Not TryReadCell(ws, CELL_SEISMIC_COEF, LOAD_SEISMIC_COEF) Then
        ReadLoadInputs = "Non-numeric or empty cell: " & CELL_SEISMIC_COEF & " (seismic inertia coefficient)"
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

' Reads one four-column earth-pressure row (B..E of rowNo) into v1..v4.
' Takes four ByRef scalars rather than the array they live in - the caller
' indexes the array at the call site. (That was originally done to chase a
' compile error whose real cause turned out to be elsewhere; it works, so
' it stayed.)
Private Function ReadLoadRow(ByVal ws As Worksheet, ByVal rowNo As Long, _
                             ByRef v1 As Double, ByRef v2 As Double, _
                             ByRef v3 As Double, ByRef v4 As Double, _
                             ByVal label As String) As String

    If Not TryReadCell(ws, "B" & rowNo, v1) Then
        ReadLoadRow = "Non-numeric or empty cell: B" & rowNo & " (" & label & ", bottom near)"
        Exit Function
    End If
    If Not TryReadCell(ws, "C" & rowNo, v2) Then
        ReadLoadRow = "Non-numeric or empty cell: C" & rowNo & " (" & label & ", bottom far)"
        Exit Function
    End If
    If Not TryReadCell(ws, "D" & rowNo, v3) Then
        ReadLoadRow = "Non-numeric or empty cell: D" & rowNo & " (" & label & ", top far)"
        Exit Function
    End If
    If Not TryReadCell(ws, "E" & rowNo, v4) Then
        ReadLoadRow = "Non-numeric or empty cell: E" & rowNo & " (" & label & ", top near)"
        Exit Function
    End If

    ReadLoadRow = ""

End Function

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

    If Not TryReadCell(ws, CELL_OPENING_WIDTH, openingWidth) Then
        ReadWingwallInputs = "Non-numeric or empty cell: " & CELL_OPENING_WIDTH & " (opening width)"
        Exit Function
    End If
    If Not TryReadCell(ws, CELL_STEM_THICKNESS, stemThickness) Then
        ReadWingwallInputs = "Non-numeric or empty cell: " & CELL_STEM_THICKNESS & " (stem thickness)"
        Exit Function
    End If
    If Not TryReadCell(ws, CELL_FOUNDATION_THICKNESS, foundationThickness) Then
        ReadWingwallInputs = "Non-numeric or empty cell: " & CELL_FOUNDATION_THICKNESS & " (foundation thickness)"
        Exit Function
    End If

    If Not TryReadCell(ws, CELL_LEFT_LENGTH, leftSide.L) Then
        ReadWingwallInputs = "Non-numeric or empty cell: " & CELL_LEFT_LENGTH & " (LEFT length)"
        Exit Function
    End If
    If Not TryReadCell(ws, CELL_LEFT_ANGLE, leftSide.AngleDeg) Then
        ReadWingwallInputs = "Non-numeric or empty cell: " & CELL_LEFT_ANGLE & " (LEFT angle)"
        Exit Function
    End If
    If Not TryReadCell(ws, CELL_LEFT_HNEAR, leftSide.Hnear) Then
        ReadWingwallInputs = "Non-numeric or empty cell: " & CELL_LEFT_HNEAR & " (LEFT height @ face)"
        Exit Function
    End If
    If Not TryReadCell(ws, CELL_LEFT_HFAR, leftSide.Hfar) Then
        ReadWingwallInputs = "Non-numeric or empty cell: " & CELL_LEFT_HFAR & " (LEFT height @ tip)"
        Exit Function
    End If

    If Not TryReadCell(ws, CELL_RIGHT_LENGTH, rightSide.L) Then
        ReadWingwallInputs = "Non-numeric or empty cell: " & CELL_RIGHT_LENGTH & " (RIGHT length)"
        Exit Function
    End If
    If Not TryReadCell(ws, CELL_RIGHT_ANGLE, rightSide.AngleDeg) Then
        ReadWingwallInputs = "Non-numeric or empty cell: " & CELL_RIGHT_ANGLE & " (RIGHT angle)"
        Exit Function
    End If
    If Not TryReadCell(ws, CELL_RIGHT_HNEAR, rightSide.Hnear) Then
        ReadWingwallInputs = "Non-numeric or empty cell: " & CELL_RIGHT_HNEAR & " (RIGHT height @ face)"
        Exit Function
    End If
    If Not TryReadCell(ws, CELL_RIGHT_HFAR, rightSide.Hfar) Then
        ReadWingwallInputs = "Non-numeric or empty cell: " & CELL_RIGHT_HFAR & " (RIGHT height @ tip)"
        Exit Function
    End If

    ' Ec is not required to build - falls back to MATERIAL_ELAST_DEFAULT
    ' if B19 is blank/invalid, same non-blocking pattern as
    ' midas-culvert-model-build.bas's B43 read.
    If Not TryReadCell(ws, CELL_MATERIAL_ELAST, MATERIAL_ELAST) Or MATERIAL_ELAST <= 0 Then
        MATERIAL_ELAST = MATERIAL_ELAST_DEFAULT
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

' A whole-number cell, floored at 1 - blank, non-numeric and < 1 all mean
' "do not divide in this direction".
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

' Computes the 20-node / 9-element geometry into NODE_LIST/ELEMENT_LIST
' (ELEMENT_LIST rows carry a 7th field, SECT). Construction reverse-
' engineered from KD.txt - full derivation and the rejected alternatives
' are in CLAUDE.md. Placement is FINAL, no build-then-rotate step:
'   A  (LEFT near)  = (0,0)                      Tl (LEFT far) = A + L along +X
'   B  (RIGHT near) = A + spacing at (90 + angleL)
'   Tr (RIGHT far)  = B + L at (angleL + angleR)   [SUM, not difference]
' A-to-B is CENTRELINE spacing, not the sheet's clear opening width.
'
' Node numbers match KD.txt's (Z=0 unless noted):
'   7/5   = LEFT/RIGHT near, wall line     8/6   = LEFT/RIGHT far, wall line
'   9/11  = LEFT/RIGHT near, inner offset  10/12 = LEFT/RIGHT far, inner
'   2/1   = LEFT/RIGHT near, outer offset  3/4   = LEFT/RIGHT far, outer
'   13/14 = near tops (Z=Hnear)            15/16 = far tops (Z=Hfar)
'   17/18 = LEFT rigid-zone tops           19/20 = RIGHT rigid-zone tops
' "Inner" = toward the opening. Near four land on the culvert face, far
' four on the tip-to-tip far edge.
'
' Elements - foundation: 1=RIGHT inner strip, 2=middle strip, 3=LEFT inner,
' 4=LEFT outer, 9=RIGHT outer.  Stem: 5/6=RIGHT sloped/rigid band,
' 7/8=LEFT sloped/rigid band.
'
' WARNING: the stem's 2-row split degenerates if a wall's Hnear or Hfar is
' SMALLER than the rigid-zone height (RIGID_ZONE_RATIO * stemThickness).
'
' leftSide/rightSide are ByRef - VBA can't pass a Type ByVal.
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

' Places one foundation offset corner. Steps rzF perpendicular off the wall
' (the offX/offY vector), then slides ALONG the wall until it meets the cut
' line - so the corner keeps its exact perpendicular distance from its own
' wall while still landing on the shared face/far-edge line. Results come
' back through outX/outY; VBA can't return a pair any other way without a
' Type, and a Type can't be passed ByVal (see the header note).
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

    ' PUT, not POST: confirmed live that POST fails ("Key Already Exist")
    ' once key "1" already exists from a prior run - PUT is idempotent
    ' (create-or-update) here the same way it is for db/UNIT/PJCF/STYP/ACTL.
    PostMaterial = PostAndCheck("db/MATL", b, "PUT")

End Function


' Posts both thickness records in one call: SECT_FOUNDATION (floor) and
' SECT_STEM (both wingwalls) - independent values, per user request.
Private Function PostThickness() As String

    Dim b As String
    b = "{""Assign"": {"
    b = b & ThicknessJson(SECT_FOUNDATION, THICKNESS_FOUNDATION_VALUE) & ","
    b = b & ThicknessJson(SECT_STEM, THICKNESS_STEM_VALUE)
    b = b & "}}"

    PostThickness = PostAndCheck("db/THIK", b, "PUT")

End Function

Private Function ThicknessJson(ByVal sectNo As Long, ByVal value As Double) As String

    Dim b As String
    b = """" & sectNo & """: {"
    b = b & """NAME"": """ & sectNo & ""","
    b = b & """TYPE"": ""VALUE"","
    b = b & """bINOUT"": false,"
    b = b & """T_IN"": " & JsonNum(value) & ","
    b = b & """T_OUT"": 0,"
    b = b & """OFFSET"": 0,"
    b = b & """O_VALUE"": 0"
    b = b & "}"

    ThicknessJson = b

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

    rows = Split(ELEMENT_LIST, ";")
    b = "{""Assign"": {"

    For i = LBound(rows) To UBound(rows)
        f = Split(rows(i), "|")
        If i > LBound(rows) Then b = b & ","
        b = b & """" & f(0) & """: {"
        b = b & """TYPE"": ""PLATE"","
        b = b & """MATL"": " & MATERIAL_NO & ","
        b = b & """SECT"": " & f(6) & ","
        b = b & """NODE"": [" & f(1) & "," & f(2) & "," & f(3) & "," & f(4) & "],"
        b = b & """ANGLE"": 0,"
        b = b & """STYPE"": " & f(5)
        b = b & "}"
    Next i

    b = b & "}}"

    PostElements = PostAndCheck("db/ELEM", b, "PUT")

End Function


Private Function PostStaticLoadCases() As String

    Dim rows() As String
    Dim f() As String
    Dim i As Long
    Dim b As String

    rows = Split(STLDCASE_LIST, ";")
    b = "{""Assign"": {"

    For i = LBound(rows) To UBound(rows)
        f = Split(rows(i), "|")
        If i > LBound(rows) Then b = b & ","
        b = b & """" & (i + 1) & """: {"
        b = b & """NAME"": """ & f(0) & ""","
        b = b & """TYPE"": """ & f(1) & ""","
        b = b & """DESC"": """""
        b = b & "}"
    Next i

    b = b & "}}"

    PostStaticLoadCases = TolerateDuplicateName(PostAndCheck("db/STLD", b, "PUT"), _
                                                "Static load cases")

End Function


' db/BODF: Assign key treated as a sequential index, not the load case
' number - unverified, but LCNAME in the body is explicit either way.
' Check Load > Self Weight shows DL/ATA_L/ATA_R and nothing else.
' DL is gravity; the ATA pair is seismic inertia in its two directions.
Private Function PostSelfWeight() As String

    Dim b As String

    b = "{""Assign"": {"
    b = b & SelfWeightJson(1, "DL", 0, 0, -1) & ","
    b = b & SelfWeightJson(2, "ATA_L", 0, LOAD_SEISMIC_COEF, 0) & ","
    b = b & SelfWeightJson(3, "ATA_R", 0, -LOAD_SEISMIC_COEF, 0)
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


' db/PRES: the Assign key is treated as the element number - inferred from
' the doc's examples (keys like "116"/"117"/"316" look like element/section
' IDs, not a plain sequence), but this is NOT stated in the doc's prose.
' Verify in Civil NX (Load > Pressure Loads) that the four loads below land
' on the stem elements 5/6/7/8, not on some other element or nothing at all.
Private Function PostPressureLoads() As String

    Dim b As String

    b = "{""Assign"": {"
    b = b & PressureJson(ELEM_STEM_LEFT_SLOPED, "LSS1_L", LOAD_SURCHARGE_LEFT) & ","
    b = b & PressureJson(ELEM_STEM_LEFT_RIGID, "LSS1_L", LOAD_SURCHARGE_LEFT) & ","
    b = b & PressureJson(ELEM_STEM_RIGHT_SLOPED, "LSS1_R", LOAD_SURCHARGE_RIGHT) & ","
    b = b & PressureJson(ELEM_STEM_RIGHT_RIGID, "LSS1_R", LOAD_SURCHARGE_RIGHT)
    b = b & "}}"

    PostPressureLoads = PostAndCheck("db/PRES", b, "PUT")

End Function

' One uniform FACE pressure on one plate, normal to it (LZ).
Private Function PressureJson(ByVal elemNo As Long, ByVal lcname As String, _
                              ByVal pressure As Double) As String

    Dim b As String
    b = """" & elemNo & """: {""ITEMS"": [{"
    b = b & """ID"": 1,"
    b = b & """LCNAME"": """ & lcname & ""","
    b = b & """GROUP_NAME"": """","
    b = b & """CMD"": ""PRES"","
    b = b & """ELEM_TYPE"": ""PLATE"","
    b = b & """FACE_EDGE_TYPE"": ""FACE"","
    b = b & """DIRECTION"": ""LZ"","
    b = b & """FORCES"": [" & JsonNum(pressure) & ",0,0,0,0]"
    b = b & "}]}"

    PressureJson = b

End Function


' Plane loads (db/PNLD + db/PNLA): six area-load types (at-rest / active /
' seismic, one pair per wingwall), each projected onto its own wall's
' plane. Points and planes are derived from the model; only the
' intensities are inputs. Named -L/-R rather than KD.txt's -kisa/-uzun,
' since with parametric lengths "short" is no longer fixed.
' PNLD_KEY in PostPlaneLoadAssignments must match the keys used here.
' Built INLINE, not via a helper. A PlaneLoadTypeJson() helper kept giving
' "Sub or Function not defined" at the call site with the function plainly
' in the module; four different theories were tried and none fixed it. No
' call means nothing to fail to bind.
'
' Each entry is a 4-point AREA load covering one wingwall face. The points
' ARE the face corners in the wall's own plane (local X along the wall from
' the near corner, local Y = height): (0,0) (L,0) (L,Hfar) (0,Hnear). The
' side sign flips the set for the wall whose plane normal faces the other
' way. Keys 1-6 must match the PNLD_KEY values in PostPlaneLoadAssignments.
Private Function PostPlaneLoadTypes() As String

    Dim b As String
    Dim i As Long
    Dim nm As String
    Dim wl As Double, hn As Double, hf As Double, ss As Double
    Dim q1 As Double, q2 As Double, q3 As Double, q4 As Double

    b = "{""Assign"": {"

    For i = 1 To 6

        If i <= 3 Then
            wl = GEO_RIGHT_L: hn = GEO_RIGHT_HNEAR: hf = GEO_RIGHT_HFAR: ss = EP_SIGN_RIGHT
        Else
            wl = GEO_LEFT_L: hn = GEO_LEFT_HNEAR: hf = GEO_LEFT_HFAR: ss = EP_SIGN_LEFT
        End If

        nm = PlaneLoadTypeName(i)

        Select Case i
            Case 1
                q1 = LOAD_R_ATREST(1): q2 = LOAD_R_ATREST(2): q3 = LOAD_R_ATREST(3): q4 = LOAD_R_ATREST(4)
            Case 2
                q1 = LOAD_R_ACTIVE(1): q2 = LOAD_R_ACTIVE(2): q3 = LOAD_R_ACTIVE(3): q4 = LOAD_R_ACTIVE(4)
            Case 3
                q1 = LOAD_R_SEISMIC(1): q2 = LOAD_R_SEISMIC(2): q3 = LOAD_R_SEISMIC(3): q4 = LOAD_R_SEISMIC(4)
            Case 4
                q1 = LOAD_L_ATREST(1): q2 = LOAD_L_ATREST(2): q3 = LOAD_L_ATREST(3): q4 = LOAD_L_ATREST(4)
            Case 5
                q1 = LOAD_L_ACTIVE(1): q2 = LOAD_L_ACTIVE(2): q3 = LOAD_L_ACTIVE(3): q4 = LOAD_L_ACTIVE(4)
            Case 6
                q1 = LOAD_L_SEISMIC(1): q2 = LOAD_L_SEISMIC(2): q3 = LOAD_L_SEISMIC(3): q4 = LOAD_L_SEISMIC(4)
        End Select

        If i > 1 Then b = b & ","
        b = b & """" & i & """: {"
        b = b & """NAME"": """ & nm & ""","
        b = b & """DESC"": ""Generated from INPUT sheet"","
        b = b & """LTYPE"": ""AREA"","
        b = b & """AREALOAD"": {"
        b = b & """bUNIFORM"": false,"
        b = b & """b3PNT"": false,"
        b = b & """X"": [0," & JsonNum(wl) & "," & JsonNum(wl) & ",0],"
        b = b & """Y"": [0,0," & JsonNum(hf) & "," & JsonNum(hn) & "],"
        b = b & """LOAD"": [" & JsonNum(ss * q1) & "," & JsonNum(ss * q2) & "," & JsonNum(ss * q3) & "," & JsonNum(ss * q4) & "]"
        b = b & "},"
        b = b & """COPY_X"": [],"
        b = b & """COPY_Y"": []"
        b = b & "}"

    Next i

    b = b & "}}"

    ' db/PNLD is NOT wrapped in TolerateDuplicateName. That wrapper matched
    ' on "Duplicate Name", and this endpoint never returns one: two records
    ' with the same NAME were both accepted live on 2026-09-22. So the
    ' wrapper was dead code here, not a live bug.
    '
    ' What IS true about db/PNLD, all confirmed live the same day:
    '   - it COMPACTS non-contiguous Assign keys. Writing keys 3 and 5 into
    '     an empty endpoint produced keys 1 and 2.
    '   - contiguous 1..6 is preserved, whether the endpoint was empty or
    '     already held six records (a PUT to an existing key replaces it).
    '   - duplicate NAMEs are accepted silently.
    ' This function only ever writes 1..6, so the normal path is safe.
    '
    ' THE ACTUAL PUT. This call went missing at some point - the function
    ' built the body and jumped straight to the read-back below without
    ' ever sending it, so VerifyPlaneLoadTypeKeys was reading an endpoint
    ' that RunCleanSteps had correctly emptied and nothing had written to
    ' since. That is exactly what "key 1 holds "" but should hold
    ' durgun-R" looks like when every key is missing, not just shifted -
    ' confirmed live 2026-09-23 after two runs reproduced the identical
    ' error even with every cleaned endpoint verified empty beforehand.
    Dim writeResult As String
    writeResult = PostAndCheck("db/PNLD", b, "PUT")
    If Len(writeResult) > 0 Then
        PostPlaneLoadTypes = writeResult
        Exit Function
    End If

    ' The read-back below is belt-and-braces, not a bug fix. db/PNLA binds
    ' each load to its type by integer PNLD_KEY - its 16 fields carry no
    ' name-based reference - so if the keys ever did shift, every area load
    ' would attach to the wrong pressure with nothing reported and a model
    ' that still solves. One GET is cheap next to that failure mode, and it
    ' proves the invariant instead of assuming it.
    PostPlaneLoadTypes = VerifyPlaneLoadTypeKeys()

End Function


' The six plane load types, in the key order PostPlaneLoadAssignments
' references them by. One source of truth: PostPlaneLoadTypes writes from
' it and VerifyPlaneLoadTypeKeys checks against it.
Private Function PlaneLoadTypeName(ByVal idx As Long) As String

    Select Case idx
        Case 1: PlaneLoadTypeName = "durgun-R"
        Case 2: PlaneLoadTypeName = "aktif-R"
        Case 3: PlaneLoadTypeName = "deprem-R"
        Case 4: PlaneLoadTypeName = "durgun-L"
        Case 5: PlaneLoadTypeName = "aktif-L"
        Case 6: PlaneLoadTypeName = "deprem-L"
    End Select

End Function


' The NAME held at one db/PNLD key. Finds the key marker, then the first
' "NAME" after it - each record is a flat object with exactly one, emitted
' in key order, so the first match belongs to that record. A wrong match
' fails the comparison, which is the safe direction.
Private Function PnldNameAtKey(ByVal json As String, ByVal keyNo As Long) As String

    Dim posKey As Long, posName As Long

    posKey = InStr(1, json, """" & keyNo & """:")
    If posKey = 0 Then Exit Function

    posName = InStr(posKey, json, """NAME""")
    If posName = 0 Then Exit Function

    PnldNameAtKey = ExtractJsonStringValue(Mid$(json, posName), "NAME")

End Function


' Reads db/PNLD back and confirms our six types actually occupy keys 1..6.
' See the comment in PostPlaneLoadTypes for why a 2xx is not enough.
Private Function VerifyPlaneLoadTypeKeys() As String

    Dim resp As String
    Dim statusCode As Long
    Dim i As Long
    Dim want As String, got As String

    Call SendApiRequest("GET", "db/PNLD", "", resp, statusCode)

    If statusCode < 200 Or statusCode >= 300 Then
        VerifyPlaneLoadTypeKeys = "wrote the plane load types but could not read " & _
            "db/PNLD back to check their keys (HTTP " & statusCode & ")."
        Exit Function
    End If

    For i = 1 To 6
        want = PlaneLoadTypeName(i)
        got = PnldNameAtKey(resp, i)
        If StrComp(got, want, vbTextCompare) <> 0 Then
            VerifyPlaneLoadTypeKeys = "plane load type key " & i & " holds """ & got & _
                """ but should hold """ & want & """. db/PNLD renumbers the keys it is " & _
                "given when the endpoint is not already empty, and db/PNLA binds each " & _
                "load to its type by integer PNLD_KEY - so the area loads would attach " & _
                "to the wrong pressures with nothing reported. Clear the plane load " & _
                "types in Civil NX and re-run."
            Exit Function
        End If
    Next i

End Function


' Each wall's load plane = three points from the generated geometry: near
' corner (origin), far corner (sets local X along the wall, matching the
' area load's X), near corner + Hnear (sets local Y, straight up).
Private Function PostPlaneLoadAssignments() As String

    Dim b As String

    ' Search tolerance for "which elements lie on this plane". KD.txt's own
    ' value (0.0009144 m = 0.003 ft, a MIDAS default in imperial guise).
    Const TOL As Double = 0.0009144

    b = "{""Assign"": {"
    ' One line each - see the note in PostPlaneLoadTypes.
    b = b & PlaneLoadAssignJson(1, "EHS2_L", 4, GEO_LEFT_NEAR_X, GEO_LEFT_NEAR_Y, GEO_LEFT_FAR_X, GEO_LEFT_FAR_Y, GEO_LEFT_HNEAR, TOL) & ","
    b = b & PlaneLoadAssignJson(2, "EHA2_L", 5, GEO_LEFT_NEAR_X, GEO_LEFT_NEAR_Y, GEO_LEFT_FAR_X, GEO_LEFT_FAR_Y, GEO_LEFT_HNEAR, TOL) & ","
    b = b & PlaneLoadAssignJson(3, "EQ_L", 6, GEO_LEFT_NEAR_X, GEO_LEFT_NEAR_Y, GEO_LEFT_FAR_X, GEO_LEFT_FAR_Y, GEO_LEFT_HNEAR, TOL) & ","
    b = b & PlaneLoadAssignJson(4, "EHS2_R", 1, GEO_RIGHT_NEAR_X, GEO_RIGHT_NEAR_Y, GEO_RIGHT_FAR_X, GEO_RIGHT_FAR_Y, GEO_RIGHT_HNEAR, TOL) & ","
    b = b & PlaneLoadAssignJson(5, "EHA2_R", 2, GEO_RIGHT_NEAR_X, GEO_RIGHT_NEAR_Y, GEO_RIGHT_FAR_X, GEO_RIGHT_FAR_Y, GEO_RIGHT_HNEAR, TOL) & ","
    b = b & PlaneLoadAssignJson(6, "EQ_R", 3, GEO_RIGHT_NEAR_X, GEO_RIGHT_NEAR_Y, GEO_RIGHT_FAR_X, GEO_RIGHT_FAR_Y, GEO_RIGHT_HNEAR, TOL)
    b = b & "}}"

    PostPlaneLoadAssignments = PostAndCheck("db/PNLA", b, "PUT")

End Function

' Builds one "key": {...} entry for db/PNLA, from one wall's near corner,
' far corner and near-end height. pnldKey must match the key used for the
' matching row in PostPlaneLoadTypes (1=durgun-R, 2=aktif-R, 3=deprem-R,
' 4=durgun-L, 5=aktif-L, 6=deprem-L).
Private Function PlaneLoadAssignJson(ByVal key As Long, ByVal lcname As String, ByVal pnldKey As Long, ByVal nearX As Double, ByVal nearY As Double, ByVal farX As Double, ByVal farY As Double, ByVal hNear As Double, ByVal tol As Double) As String

    Dim b As String
    b = """" & key & """: {"
    b = b & """LCNAME"": """ & lcname & ""","
    b = b & """LOAD_GROUP"": """","
    b = b & """PNLD_KEY"": " & pnldKey & ","
    b = b & """ELEM_TYPE"": ""PLATE"","
    ' Origin = the wall's near corner, at foundation level.
    b = b & """POINT_ORIGIN"": [" & JsonNum(nearX) & "," & JsonNum(nearY) & ",0],"
    ' Local X runs along the wall, toward the far corner.
    b = b & """AXIS_X"": [" & JsonNum(farX) & "," & JsonNum(farY) & ",0],"
    ' Local Y runs straight up, so the area load's Y really is height.
    b = b & """AXIS_Y"": [" & JsonNum(nearX) & "," & JsonNum(nearY) & "," & JsonNum(hNear) & "],"
    b = b & """TOL"": " & JsonNum(tol) & ","
    b = b & """SELECT_TYPE"": ""ON_PLANE"","
    b = b & """LOAD_DIR"": ""NORMAL_PLANE"","
    b = b & """PROJECT_TYPE"": ""NO"","
    b = b & """DESC"": ""Generated from INPUT sheet"""
    b = b & "}"

    PlaneLoadAssignJson = b

End Function


' One PUT per dependency tier - see LOADCOMB_LIST's own comment for why a
' single request is not safe once a combination references another one.
Private Function PostLoadCombinations() As String

    Dim tier As Long
    Dim res As String

    For tier = 1 To 3
        res = PostCombTier(tier)
        If Len(res) > 0 Then
            If Left$(res, 5) = "WARN:" Then
                PostLoadCombinations = res
            Else
                ' Naming the tier narrows any future failure to a third of
                ' the set immediately.
                PostLoadCombinations = "tier " & tier & " - " & res
                Exit Function
            End If
        End If
    Next tier

End Function

' The combinations at one dependency tier. Keys are the row's position in
' LOADCOMB_LIST (i + 1), so they stay globally unique and in definition
' order across all three tiers. Returns "" when the tier is empty or
' wrote cleanly.
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
'  THICKNESS COLOURS (db/CO_T)
' ===========================================================================

' Colour is not an attribute of db/THIK - that page has no colour field at
' all. It lives in the API manual's separate View group, keyed by thickness
' number, alongside db/CO_M (material), db/CO_S (section) and db/CO_F.
Private Function PostThicknessColors() As String

    Dim rows() As String
    Dim f() As String
    Dim c() As String
    Dim i As Long
    Dim b As String

    rows = Split(THICK_COLOR_LIST, ";")
    b = "{""Assign"": {"

    For i = LBound(rows) To UBound(rows)

        f = Split(rows(i), "|")
        c = Split(f(1), ",")

        If i > LBound(rows) Then b = b & ","
        b = b & """" & f(0) & """: {"
        b = b & """W_R"": " & c(0) & ",""W_G"": " & c(1) & ",""W_B"": " & c(2) & ","
        b = b & """HF_R"": " & Lighten(c(0)) & ",""HF_G"": " & Lighten(c(1)) & _
                ",""HF_B"": " & Lighten(c(2)) & ","
        b = b & """HE_R"": " & c(0) & ",""HE_G"": " & c(1) & ",""HE_B"": " & c(2) & ","
        b = b & """bBLEMD"": " & LCase(THICK_COLOR_TRANSLUCENT) & ","
        b = b & """FACT"": " & JsonNum(THICK_COLOR_OPACITY)
        b = b & "}"

    Next i

    b = b & "}}"

    ' PUT: db/CO_T lists GET and PUT as its only active methods.
    PostThicknessColors = PostAndCheck("db/CO_T", b, "PUT")

End Function

' ===========================================================================
'  NAMED UCS (db/NUCS)
' ===========================================================================

' Writes the foundation's coordinate system - see the db/NUCS block for the
' orientation and for what is confirmed vs still open. Origin at the global
' origin, X rotated by the LEFT wall's splay angle, Y perpendicular to it in
' plan, so Y lands along the culvert face.
Private Function PostNamedUcs() As String

    Dim b As String
    Dim a As Double
    Dim cs As Double, sn As Double

    a = DegToRad(GEO_LEFT_ANGLE_DEG)
    cs = Cos(a)
    sn = Sin(a)

    b = "{""Assign"": {""1"": {"
    b = b & """NAME"": """ & UCS_NAME & ""","
    b = b & """ORG_ITEM"": [0,0,0],"
    b = b & """VX_ITEM"": [" & JsonNum(cs) & "," & JsonNum(sn) & ",0],"
    b = b & """VY_ITEM"": [" & JsonNum(-sn) & "," & JsonNum(cs) & ",0]"
    b = b & "}}}"

    PostNamedUcs = PostAndCheck("db/NUCS", b, "PUT")

End Function


' One channel lightened for the hidden fill, clamped to 255.
Private Function Lighten(ByVal channel As String) As Long

    Dim v As Long
    v = CLng(Round(CDbl(channel) * THICK_FILL_LIGHTEN, 0))
    If v > 255 Then v = 255
    If v < 0 Then v = 0
    Lighten = v

End Function


' ===========================================================================
'  FOUNDATION SURFACE SPRINGS (ope/SSPS)
' ===========================================================================

' Converts the foundation plates' surface into point springs at their own
' nodes - MIDAS does the tributary-area arithmetic, so unlike the culvert's
' db/NSPR step nothing has to be discovered or distributed here.
'
' POST is the endpoint's ONLY active method: the one non-PUT write in this
' module. The result lands in db/NSPR, which is why RunCleanSteps deletes
' that endpoint.
Private Function PostFoundationSprings() As String

    Dim b As String
    Dim kh As Double
    Dim res As String
    Dim i As Long

    If SPRING_KV <= 0 Then
        PostFoundationSprings = "WARN: no subgrade modulus at " & INPUT_SHEET_NAME & _
            "!" & CELL_SUBGRADE_MODULUS & " - foundation springs not written."
        Exit Function
    End If

    ' Normally already loaded by PostElementLists, which runs first.
    res = RefreshElementSets()
    If Len(res) > 0 Then
        PostFoundationSprings = res
        Exit Function
    End If

    If FOUND_COUNT = 0 Then
        PostFoundationSprings = "WARN: no foundation plates found - springs not written."
        Exit Function
    End If

    kh = SPRING_KV * SPRING_H_RATIO

    b = "{""Argument"": {"
    b = b & """CONVERT_TO"": ""POINT_SPRING"","
    b = b & """GROUP_NAME"": """ & SPRING_GROUP_NAME & ""","
    b = b & """NODE_ELEMS"": {""KEYS"": ["
    For i = 1 To FOUND_COUNT
        If i > 1 Then b = b & ","
        b = b & FOUND_IDS(i)
    Next i
    b = b & "]},"
    b = b & """ELEMENT"": {""TYPE"": ""PLANAR""},"
    b = b & """BOUNDARY"": {"
    b = b & """TYPE"": ""LINEAR"","
    b = b & """STIFF"": [" & JsonNum(kh) & "," & JsonNum(kh) & "," & _
            JsonNum(SPRING_KV) & "],"
    b = b & """bDAMP"": false,"
    b = b & """DAMP"": [0,0,0]"
    b = b & "}}}"

    PostFoundationSprings = PostAndCheck("ope/SSPS", b, "POST")

End Function


' ===========================================================================
'  ELEMENT DIVISION (ope/DIVIDEELEM)
' ===========================================================================

' Walks DIVIDE_TABLE in element order, one POST per plate, because every
' plate has its own NUM_X/NUM_Y pair - there is no single request that
' covers them all.
'
' Around each plate whose family feeds G10/G11, the whole model is
' snapshotted before the divide and again after, so that plate's children
' are exactly the difference. Nothing in the model can tell them apart
' afterwards - a rigid sub-element carries the same thickness as its
' non-rigid neighbours.
Private Function PostDivideElements() As String

    Dim rows() As String
    Dim f() As String
    Dim i As Long
    Dim elemNo As Long
    Dim nx As Long, ny As Long
    Dim res As String
    Dim b As String
    Dim divided As Long
    Dim tracked As Long

    NR_FOUND_COUNT = 0
    NR_WALL_COUNT = 0
    ELEM_SETS_READY = False

    If MESH_X_WALL <= 1 And MESH_Y_WALL <= 1 And _
       MESH_X_FOUND <= 1 And MESH_Y_FOUND <= 1 Then
        PostDivideElements = "WARN: every mesh count at " & INPUT_SHEET_NAME & "!" & _
            CELL_MESH_X_WALL & "/" & CELL_MESH_Y_WALL & "/" & CELL_MESH_X_FOUND & "/" & _
            CELL_MESH_Y_FOUND & " is blank or < 2 - elements left undivided."
        Call SeedUndividedFamilies
        Exit Function
    End If

    rows = Split(DIVIDE_TABLE, ";")

    For i = LBound(rows) To UBound(rows)

        f = Split(rows(i), "|")
        elemNo = CLng(f(0))
        nx = MeshCount(f(1), elemNo)
        ny = MeshCount(f(2), elemNo)
        tracked = TrackedList(elemNo)

        ' Snapshot the whole model BEFORE a tracked plate is split, so the
        ' elements that appear are exactly its children.
        If tracked <> 0 Then Call SnapshotElements

        If nx > 1 Or ny > 1 Then

            b = "{""Argument"": {"
            b = b & """TARGETS"": [" & elemNo & "],"
            b = b & """DIVIDE"": {"
            b = b & """ELEM_TYPE"": ""Planar"","
            b = b & """DIV_METHOD"": ""Equal"","
            b = b & """OPTION"": {""EQUAL_OPTION"": {"
            b = b & """NUM_X"": " & nx & ","
            b = b & """NUM_Y"": " & ny
            b = b & "}}}}}"

            res = PostAndCheck("ope/DIVIDEELEM", b, "POST")
            If Len(res) > 0 Then
                PostDivideElements = "element " & elemNo & " (" & nx & "x" & ny & ") - " & res
                Exit Function
            End If

            divided = divided + 1

        End If

        If tracked <> 0 Then Call CaptureFamily(elemNo, tracked)

    Next i

    If divided = 0 Then
        PostDivideElements = "WARN: the mesh counts left every plate at 1x1 - " & _
                             "nothing was divided."
    End If

End Function

' 1 if this as-built plate's family belongs in G10, 2 if in G11, else 0.
Private Function TrackedList(ByVal elemNo As Long) As Long

    If InCsv(NONRIGID_FOUND_ELEMS, elemNo) Then
        TrackedList = 1
    ElseIf InCsv(NONRIGID_WALL_ELEMS, elemNo) Then
        TrackedList = 2
    End If

End Function

Private Function InCsv(ByVal csv As String, ByVal needle As Long) As Boolean

    Dim parts() As String
    Dim i As Long

    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        If Trim$(parts(i)) <> vbNullString Then
            If CLng(Trim$(parts(i))) = needle Then
                InCsv = True
                Exit Function
            End If
        End If
    Next i

End Function

' Resolves one DIVIDE_TABLE cell - "MESHX"/"MESHY" against the element's own
' part of the structure, anything else as a literal count.
Private Function MeshCount(ByVal spec As String, ByVal elemNo As Long) As Long

    Dim isWall As Boolean
    isWall = (SectOfBuiltElement(elemNo) = SECT_STEM)

    Select Case UCase$(spec)
        Case "MESHX"
            MeshCount = IIf(isWall, MESH_X_WALL, MESH_X_FOUND)
        Case "MESHY"
            MeshCount = IIf(isWall, MESH_Y_WALL, MESH_Y_FOUND)
        Case Else
            If IsNumeric(spec) Then
                MeshCount = CLng(spec)
            Else
                MeshCount = 1
            End If
    End Select

    If MeshCount < 1 Then MeshCount = 1

End Function

' The thickness an AS-BUILT element number carries, straight out of
' ELEMENT_LIST. Only valid for the original 9 - which is all DIVIDE_TABLE
' ever asks about, since it runs before anything has been subdivided.
Private Function SectOfBuiltElement(ByVal elemNo As Long) As Long

    Dim rows() As String
    Dim f() As String
    Dim i As Long

    rows = Split(ELEMENT_LIST, ";")
    For i = LBound(rows) To UBound(rows)
        f = Split(rows(i), "|")
        If CLng(f(0)) = elemNo Then
            SectOfBuiltElement = CLng(f(6))
            Exit Function
        End If
    Next i

End Function

' Nothing was divided, so each tracked plate is its own whole family.
Private Sub SeedUndividedFamilies()

    Dim rows() As String
    Dim i As Long
    Dim elemNo As Long

    rows = Split(DIVIDE_TABLE, ";")
    For i = LBound(rows) To UBound(rows)
        elemNo = CLng(Split(rows(i), "|")(0))
        Select Case TrackedList(elemNo)
            Case 1: Call AppendLong(NR_FOUND_IDS, NR_FOUND_COUNT, elemNo)
            Case 2: Call AppendLong(NR_WALL_IDS, NR_WALL_COUNT, elemNo)
        End Select
    Next i

End Sub

' Every element in the model right now. A failed GET leaves the snapshot
' empty, which CaptureFamily treats as "cannot tell" and falls back to the
' plate's own number.
Private Sub SnapshotElements()

    Dim resp As String
    Dim statusCode As Long

    SNAP_COUNT = 0
    Call SendApiRequest("GET", "db/ELEM", "", resp, statusCode)
    If statusCode < 200 Or statusCode >= 300 Then Exit Sub

    Call ParseElementsBySect(resp, SECT_ANY, SNAP_IDS, SNAP_COUNT)

End Sub

' Everything that was NOT in the snapshot, plus the plate's own number (a
' divide reuses it for one of the children, so it never appears as new).
' whichList picks the destination: 1 = G10's set, 2 = G11's.
Private Sub CaptureFamily(ByVal elemNo As Long, ByVal whichList As Long)

    Dim resp As String
    Dim statusCode As Long
    Dim nowIds() As Long
    Dim nowCount As Long
    Dim i As Long

    If whichList = 1 Then
        Call AppendLong(NR_FOUND_IDS, NR_FOUND_COUNT, elemNo)
    Else
        Call AppendLong(NR_WALL_IDS, NR_WALL_COUNT, elemNo)
    End If

    If SNAP_COUNT = 0 Then Exit Sub

    Call SendApiRequest("GET", "db/ELEM", "", resp, statusCode)
    If statusCode < 200 Or statusCode >= 300 Then Exit Sub

    Call ParseElementsBySect(resp, SECT_ANY, nowIds, nowCount)

    For i = 1 To nowCount
        If Not InLongArray(SNAP_IDS, SNAP_COUNT, nowIds(i)) Then
            If whichList = 1 Then
                Call AppendLong(NR_FOUND_IDS, NR_FOUND_COUNT, nowIds(i))
            Else
                Call AppendLong(NR_WALL_IDS, NR_WALL_COUNT, nowIds(i))
            End If
        End If
    Next i

End Sub

Private Sub AppendLong(ByRef ids() As Long, ByRef count As Long, ByVal v As Long)

    If count = 0 Then ReDim ids(1 To 64)
    count = count + 1
    If count > UBound(ids) Then ReDim Preserve ids(1 To UBound(ids) * 2)
    ids(count) = v

End Sub

Private Function InLongArray(ByRef ids() As Long, ByVal count As Long, _
                             ByVal needle As Long) As Boolean

    Dim i As Long

    For i = 1 To count
        If ids(i) = needle Then
            InLongArray = True
            Exit Function
        End If
    Next i

End Function

' ===========================================================================
'  ELEMENT LISTS -> INPUT!G10 / G11 / G12
' ===========================================================================

' Reads the elements back out of the model and writes three range strings
' (e.g. "1to4 9") to the INPUT sheet, so the capture scripts that read G10
' and G11 stay in step with the model without a hand edit. Foundation and
' stem are told apart by their thickness number, which is what db/ELEM's
' "SECT" holds for a plate.
Private Function PostElementLists() As String

    Dim ws As Worksheet
    Dim res As String
    Dim sFound As String, sWall As String, sFoundPrr As String

    res = RefreshElementSets()
    If Len(res) > 0 Then
        PostElementLists = res
        Exit Function
    End If

    If FOUND_COUNT = 0 And WALL_COUNT = 0 Then
        PostElementLists = "WARN: db/ELEM returned no plates - element lists not written."
        Exit Function
    End If

    ' G12 is every foundation plate, straight off db/ELEM. G10 and G11 are
    ' the NON-RIGID families that PostDivideElements tracked - they cannot
    ' be recovered from db/ELEM alone, since a rigid sub-element carries the
    ' same thickness as its non-rigid neighbours.
    Call SortLongs(NR_FOUND_IDS, NR_FOUND_COUNT)
    Call SortLongs(NR_WALL_IDS, NR_WALL_COUNT)

    sFoundPrr = FormatElementRangeList(FOUND_IDS, FOUND_COUNT)
    sFound = FormatElementRangeList(NR_FOUND_IDS, NR_FOUND_COUNT)
    sWall = FormatElementRangeList(NR_WALL_IDS, NR_WALL_COUNT)

    If NR_FOUND_COUNT = 0 Or NR_WALL_COUNT = 0 Then
        PostElementLists = "WARN: no non-rigid families were tracked - run " & _
            "Divide Elements first, or G10/G11 will be written empty."
        Exit Function
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PROJECT_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        PostElementLists = "WARN: sheet """ & PROJECT_SHEET_NAME & """ not found; " & _
            "element lists not written. FOUND: " & sFound & "  WALL: " & sWall
        Exit Function
    End If

    On Error Resume Next
    ws.Range(CELL_ELEM_LIST_FOUND).Value = sFound
    ws.Range(CELL_ELEM_LIST_WALL).Value = sWall
    ws.Range(CELL_ELEM_LIST_FOUND_PRR).Value = sFoundPrr
    If Err.Number <> 0 Then
        PostElementLists = "WARN: failed to write the element lists to " & _
            PROJECT_SHEET_NAME & ": " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

End Function

' GETs db/ELEM once and splits the plates into FOUND_IDS/WALL_IDS by
' thickness, sorted. Shared by the element-list write-back and the surface
' springs, both of which run after the divide and want the same data.
' Returns "" on success or an HTTP message.
Private Function RefreshElementSets() As String

    Dim resp As String
    Dim statusCode As Long

    If ELEM_SETS_READY Then Exit Function

    Call SendApiRequest("GET", "db/ELEM", "", resp, statusCode)
    If statusCode < 200 Or statusCode >= 300 Then
        RefreshElementSets = "HTTP " & statusCode & " - failed to query db/ELEM."
        Exit Function
    End If

    Call ParseElementsBySect(resp, SECT_FOUNDATION, FOUND_IDS, FOUND_COUNT)
    Call ParseElementsBySect(resp, SECT_STEM, WALL_IDS, WALL_COUNT)
    Call SortLongs(FOUND_IDS, FOUND_COUNT)
    Call SortLongs(WALL_IDS, WALL_COUNT)

    ELEM_SETS_READY = True

End Function

' Collects the element numbers whose "SECT" equals sectWanted out of a
' db/ELEM GET response, or every element when passed SECT_ANY. Hand-rolled
' because VBA has no JSON parser: walks
' the ELEM object's keys, skipping any value that is not itself an object.
Private Sub ParseElementsBySect(ByVal json As String, ByVal sectWanted As Long, _
                                ByRef elemIds() As Long, ByRef count As Long)

    Dim pElem As Long, pos As Long, posColon As Long
    Dim posQ1 As Long, posQ2 As Long
    Dim posObjStart As Long, posObjEnd As Long
    Dim posNextQuote As Long
    Dim keyStr As String
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
                If sectWanted = SECT_ANY Or _
                   CLng(ExtractJsonNumFromBlock(json, posObjStart, posObjEnd, "SECT")) _
                   = sectWanted Then
                    count = count + 1
                    If count > maxCap Then
                        maxCap = maxCap * 2
                        ReDim Preserve elemIds(1 To maxCap)
                    End If
                    elemIds(count) = CLng(keyStr)
                End If
            End If

            pos = posObjEnd + 1
        End If
    Loop

End Sub

Private Function ExtractJsonNumFromBlock(ByVal json As String, ByVal pStart As Long, _
                                         ByVal pEnd As Long, ByVal key As String) As Double

    Dim blk As String
    Dim pk As Long, pc As Long

    blk = Mid$(json, pStart, pEnd - pStart + 1)
    pk = InStr(1, blk, """" & key & """", vbTextCompare)
    If pk = 0 Then Exit Function

    pc = InStr(pk + Len(key) + 2, blk, ":")
    If pc = 0 Then Exit Function

    ExtractJsonNumFromBlock = Val(Mid$(blk, pc + 1))

End Function

Private Sub SortLongs(ByRef ids() As Long, ByVal count As Long)

    Dim i As Long, j As Long, tmp As Long

    For i = 1 To count - 1
        For j = i + 1 To count
            If ids(j) < ids(i) Then
                tmp = ids(i): ids(i) = ids(j): ids(j) = tmp
            End If
        Next j
    Next i

End Sub

' Compresses a SORTED array of element IDs into MIDAS range syntax, e.g.
' "4 10to13 17to19".
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

' Runs the solver once the model, loads, combinations and springs are all in
' place. POST with an empty body, exactly as the doc's own non-Pushover
' example does.
Private Function PostPerformAnalysis() As String
    PostPerformAnalysis = PostAndCheck("doc/ANAL", "{}", "POST")
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
    ' redundant: it is the only thing that survives a bad extraction, and a
    ' truncated message once sent a whole debugging session the wrong way
    ' on the culvert builder. Keep both.
    msg = ExtractJsonStringValue(resp, "message")
    PostAndCheck = "HTTP " & statusCode & HttpStatusHint(statusCode) & _
                   IIf(Len(msg) > 0, " - " & msg, "") & _
                   vbCrLf & "       raw: " & Left$(resp, 400)

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

' Escapes a string for use as a JSON value. Everything else this module
' emits is a name it generated itself, but the project name, engineer and
' revision are free text typed into a worksheet, so a stray quote or
' backslash there would otherwise produce a malformed body.
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

' Pulls the string value of "key": "value" out of a flat JSON response.
' Not a general JSON parser - only safe for a flat/top-level string field
' such as an error response's "message". Honours \" escapes: taking
' everything between the first two quotes after the colon truncates any
' message that quotes a name, which is exactly how a real MIDAS error once
' got misdiagnosed.
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
