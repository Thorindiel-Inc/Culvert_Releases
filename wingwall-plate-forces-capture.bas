Option Explicit

' ============================================================================
'  MIDAS Civil NX API - Wingwall Plate Forces/Moments Screenshot Capture
' ----------------------------------------------------------------------------
'  Endpoint : {base url} + view/CAPTURE
'  Docs     : "Capture"                            (JSON Manual, ed. 2025.04.24)
'             "Plate Forces/Moments - Result Display" (JSON Manual, ed. 2025.08.19)
'             - Argument.RESULT_GRAPHIC follows that same view/RESULTGRAPHIC
'             schema, with CURRENT_MODE = "PlateForces/Moments".
'             "Type of Display"                  (JSON Manual, ed. 2026.06.19)
'
'  The wingwall is a shell/plate model, not a beam/frame model - its results
'  are plate forces and moments (Fxx/Fyy/Fxy, Mxx/Myy/Mxy, Vxx/Vyy), not the
'  beam-element diagrams midas-beam-diagram-capture.bas captures. This is
'  the plate counterpart: same overall mechanism, own CONFIG (own output
'  folder, own JOB_LIST). Captures one plate-force image per job in
'  JOB_LIST (load case + component + target shape), replacing the matching
'  Picture/Shape on TARGET_SHEET_NAME with the freshly captured image.
' ============================================================================


' ---------------------------------------------------------------------------
'  CONFIG
' ---------------------------------------------------------------------------

' Printed in the summary MsgBox. These modules are pasted into Excel by
' hand, so the file in the repo and the code actually running can silently
' diverge - check this stamp matches the constant here before concluding
' anything from a run. Bump it whenever this file changes.
Private Const SCRIPT_VERSION As String = "2026-09-23g"

' Identifies this module to the updater regardless of what it was
' named when pasted into Excel - these files carry no VB_Name, so the
' module name in the VBA project is whatever the user typed.
Private Const SCRIPT_ID As String = "wingwall-plate-forces-capture"

' Base URL from Civil NX: Tools > API > API Setting
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

' One WinHTTP client reused for every request in a run. Must live here:
' ALL module-level declarations precede the first procedure in VBA.
' Created on first use, released by EndFastMode.
Private HTTP_CLIENT As Object

' Images are saved in a subfolder next to this workbook, nested under a
' folder named after the workbook itself, i.e. for "mnf151.xlsm":
' <workbook folder>\mnf151\WINGWALL_RESULTS\. Created automatically if
' missing. Kept distinct from the main-wall "RESULTS" folder deliberately -
' if this runs against the same workbook as midas-beam-diagram-capture.bas,
' the two sets of captures must not land in the same folder.
Private Const CAPTURE_SUBFOLDER As String = "WINGWALL_RESULTS"

' Sheet that holds the figures to replace. TODO: set this once you know
' which sheet holds the wingwall figures - left empty for now, which uses
' whichever sheet is active when the macro runs.
Private Const TARGET_SHEET_NAME As String = "4_SONUC"

' One capture per row: CombName|CombType|MinMax|Component|ShapeName
'   CombName   : the actual MIDAS load-combination name ("ULS", "SLS",
'                "ENV_EQ") - sent as LOAD_CASE_COMB.NAME, not a per-row label.
'   CombType   : "CB" (Load Combination) for every row here - ULS/SLS/ENV_EQ
'                are all named combinations, not raw load cases.
'   MinMax     : "Min" "Max" "All" - only ENV_EQ is a true envelope, so
'                MinMax is a required-but-mostly-unused placeholder ("Max")
'                on the ULS/SLS rows; the two "My Min"/"My Mak" rows are the
'                only ones where it's meaningful.
'   Component  : per the Plate Forces/Moments doc's COMP enum -
'                "Fxx" "Fyy" "Fxy" "Fmax" "Fmin" "FMax"
'                "Mxx" "Myy" "Mxy" "Mmax" "Mmin" "MMax"
'                "Vxx" "Vyy" "VMax" "Wood Armer Moment" "Fvector" "Mvector"
'                (reported in the coordinate system CoordSystemForShape
'                picks for the row - "dvr" rows Local element axes, "tml"
'                rows the current UCS. See the LOCAL_UCS_TYPE_* block.)
'   ShapeName  : encodes both the Picture/Shape name AND which element list
'                to activate - "dvr" (Duvar/wall) rows use ELEMENT_LIST_WALL,
'                "tml" (Temel/foundation) rows use ELEMENT_LIST_FOUND, per
'                ElementListForShape(). TODO: these are placeholder guesses
'                ('str_dvr_fx, ...) - rename to match the actual Picture/
'                Shape names once they exist on TARGET_SHEET_NAME.
' Rows separated by ";". Combinations below per the owner-supplied table
' (STR/EQ/SER checks at DVR/TML locations, ULS/ENV_EQ/SLS combinations).
Private Const JOB_LIST As String = _
    "ULS|CB|All|Fxx|str_dvr_fx;" & _
    "ULS|CB|All|Mxx|str_dvr_mx;" & _
    "ULS|CB|All|Vxx|str_dvr_vx;" & _
    "ULS|CB|All|Fyy|str_tml_fy;" & _
    "ULS|CB|All|Myy|str_tml_my;" & _
    "ULS|CB|All|Vyy|str_tml_vy;" & _
    "ENV_EQ|CB|All|Fxx|eq_dvr_fx;" & _
    "ENV_EQ|CB|Min|Mxx|eq_dvr_mx;" & _
    "ENV_EQ|CB|All|Vxx|eq_dvr_vx;" & _
    "ENV_EQ|CB|All|Fyy|eq_tml_fy;" & _
    "ENV_EQ|CB|Min|Myy|eq_tml_my_min;" & _
    "ENV_EQ|CB|Max|Myy|eq_tml_my_mak;" & _
    "ENV_EQ|CB|All|Vyy|eq_tml_vy;" & _
    "SLS|CB|All|Fxx|ser_dvr_fx;" & _
    "SLS|CB|All|Mxx|ser_dvr_mx;" & _
    "SLS|CB|All|Fyy|ser_tml_fy;" & _
    "SLS|CB|All|Myy|ser_tml_my"

' STEP_INDEX is Required per the doc's spec table, but is only meaningful
' for Construction Stage ("CS") load cases - 1 is the safe default for the
' plain static ("ST") cases in JOB_LIST above (matches the doc's own
' worked ST examples, which all use "STEP_INDEX": 1).
Private Const STEP_INDEX_DEFAULT As Long = 1

' Coordinate system the Component values are reported in (OPTIONS.LOCAL_UCS),
' chosen PER SHAPE on the same "dvr"/"tml" split as the element lists and
' camera angles - see CoordSystemForShape.
'   "Local" : element local axes - nothing needs defining first.
'   "UCS"   : whatever UCS is current in Civil NX (see the warning below).
'
' WALL rows stay Local: a stem plate's own axes follow the wall face, which
' is the natural way to read its forces.
'
' FOUNDATION rows use the UCS, because local axes are NOT comparable across
' the foundation. Local x runs node1 -> node2, and the five strips are wound
' per side: elements 3 and 4 (LEFT) point one way while 1, 2 and 9 (RIGHT)
' are rotated by the two splay angles - about 45 degrees apart on the
' reference model. "Mxx" on a left strip and "Mxx" on a right strip would be
' measuring different directions.
'
' The UCS itself is created by midas-wingwall-model-build.bas, which PUTs it
' to db/NUCS under the name below - origin at the global origin, X rotated by
' the LEFT wall's splay angle, so Y runs along the culvert face. See that
' script's db/NUCS block for the derivation. The endpoint is NUCS, not UCS,
' which is why searching the endpoint list for "UCS" finds nothing; it is
' also missing from the cached API manual, whose export predates it.
'
' *** UNRESOLVED: whether UCS_NAME accepts this name. *** The Plate Forces
' doc's enum lists only the literal "CurrentUCS", and nothing readable
' reports which UCS is current - there is no info/view/* endpoint (404) and
' GET view/RESULTGRAPHIC returns "error status" without results in the
' model. So this sends the real name and finds out from the pictures:
' IF THE FOUNDATION IMAGES COME BACK UNROTATED, select the UCS by hand in
' Civil NX and set UCS_NAME_FOR_RESULTS back to "CurrentUCS".
'
' Intended orientation (owner, 2026-09-22): the wall that lies on the global
' X axis - the LEFT wingwall - rotated by that wall's own splay angle
' (INPUT!C22). Nothing here computes or checks that; it is what should be
' set up in Civil NX, recorded so the pictures can be read later.
Private Const LOCAL_UCS_TYPE_DVR As String = "Local"
Private Const LOCAL_UCS_TYPE_TML As String = "UCS"
Private Const LOCAL_UCS_TYPE_OTHER As String = "Local"
' The name midas-wingwall-model-build.bas gives the UCS at db/NUCS. Fall
' back to "CurrentUCS" if the API turns out to ignore a real name.
Private Const UCS_NAME_FOR_RESULTS As String = "FOUND"

' Argument.DISPLAY.VIEW.UCS_AXIS (per the "Display" JSON Manual, ed.
' 2024.10.25 - the same schema Argument.DISPLAY follows inside
' view/CAPTURE), intended to draw the UCS axis triad on the capture
' itself as a visual sanity check.
'
' CONFIRMED LIVE 2026-09-23: this field is a NO-OP, and so is the
' equivalent toggle in Civil NX's own GUI - the UCS itself is genuinely
' applied (component values differ under "FOUND" vs "CurrentUCS"/"Local",
' see CLAUDE.md's "view/CAPTURE DOES honour a named UCS"), but nothing
' visibly marks it on screen or in a capture, in the API or by hand.
' MIDAS's own limitation, not a request-shape problem - left wired in
' (harmless) rather than ripped out, in case a future Civil NX version
' fixes it. "tml" rows only: "dvr" rows use Local.
Private Const SHOW_UCS_AXIS As Boolean = True

' How adjacent-element values are combined (OPTIONS.AVERAGE_NODAL.TYPE).
'   "Element"  : raw per-element values, no smoothing (default here).
'   "Avg.Nodal": nodal-averaged across adjacent elements.
Private Const AVERAGE_NODAL_TYPE As String = "Avg.Nodal"

' Image size in pixels, matched to the Excel figure boxes' physical size
' at 120 DPI, so the captured image isn't stretched/squished when
' ReplacePictureByName fits it into that same box.
Private Const IMG_WIDTH As Long = 1692
Private Const IMG_HEIGHT As Long = 1000

' Camera zoom on the whole model view.
'   25 <= value < 100 : zoom out
'   100                : zoom to fit (default)
'   100 < value <= 200 : zoom in
Private Const ZOOM_LEVEL As Long = 100

' Elements to activate before capture (only these are shown/plotted).
' Space- or comma-separated; supports MIDAS "AtoB" range shorthand.
' Read at runtime from the "INPUT" sheet instead of hardcoded here, since
' the wingwall's element numbers differ per workbook - G10 = foundation
' ("tml"), G11 = wall ("dvr"). Populated by ReadElementListsFromInput()
' at the start of CaptureWingwallPlateForces(); empty ("") if the cell
' is blank, which skips sending ACTIVE at all (shows/plots everything).
Private ELEMENT_LIST_WALL As String
Private ELEMENT_LIST_FOUND As String

' Camera view angle (Argument.ANGLE.HORIZONTAL/VERTICAL, degrees) - "dvr"
' (Duvar/wall) rows and "tml" (Temel/foundation) rows each need their own
' rotation, same ShapeName split as ELEMENT_LIST_WALL/FOUND above. Any
' ShapeName matching neither sends no ANGLE block, keeping whatever view
' is currently active in Civil NX.
Private Const VIEW_ANGLE_HORIZONTAL_FOR_DVR As Long = 255
Private Const VIEW_ANGLE_VERTICAL_FOR_DVR As Long = 7
Private Const VIEW_ANGLE_HORIZONTAL_FOR_TML As Long = 0
Private Const VIEW_ANGLE_VERTICAL_FOR_TML As Long = 90

' Perspective projection (Argument.PERSPECTIVE, boolean, default false per
' the Capture doc's spec table) - on for "dvr" (wall) rows only, off (not
' sent) for everything else.
Private Const PERSPECTIVE_FOR_DVR As Boolean = True

' Force/length units to switch to for these captures (db/UNIT, "Unit System"
' JSON Manual, ed. 2024.08.01: FORCE "N" "KN" "KGF" "TONF" "LBF" "KIPS",
' DIST "M" "CM" "MM" "FT" "IN"). The macro reads the model's current Unit
' System first, switches to these values, runs all captures, then always
' switches back to whatever it originally was - see CaptureWingwallPlateForces.
Private Const CAPTURE_FORCE_UNIT As String = "KN"
Private Const CAPTURE_DIST_UNIT As String = "M"

' Rotation of the printed value-label text itself (TYPE_OF_DISPLAY.VALUES.
' SET_ORIENT), in degrees, 0 to 180 in increments of 15 per the doc. This
' is text orientation only - it does not move the camera/viewpoint.
Private Const VALUE_TEXT_ORIENT As Long = 15


' ---------------------------------------------------------------------------
'  One row of JOB_LIST, parsed
' ---------------------------------------------------------------------------
Private Type CaptureJob
    CombName As String
    CombType As String
    MinMax As String
    CompName As String
    shapeName As String
    ElemList As String
End Type


' ===========================================================================
'  MAIN - one image per job in JOB_LIST
' ===========================================================================
Sub CaptureWingwallPlateForces()

    Dim jobs() As CaptureJob
    Dim i As Integer
    Dim body As String
    Dim resp As String
    Dim statusCode As Long
    Dim outFolder As String
    Dim exportPath As String
    Dim ws As Worksheet
    Dim origForce As String, origDist As String, origHeat As String, origTemper As String
    Dim unitsChanged As Boolean
    Dim prevCalc As XlCalculation
    Dim fastMode As Boolean

    Dim label As String
    Dim okLog As String, warnLog As String, failLog As String
    Dim okCount As Long, warnCount As Long, failCount As Long

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save this workbook first - it has no folder to save captures into yet.", vbCritical
        Exit Sub
    End If

    outFolder = ThisWorkbook.Path & "\" & WorkbookBaseName() & "\" & CAPTURE_SUBFOLDER & "\"

    If Not EnsureFolderExists(outFolder) Then
        MsgBox "Could not create folder: " & outFolder, vbCritical
        Exit Sub
    End If

    If Len(TARGET_SHEET_NAME) > 0 Then
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(TARGET_SHEET_NAME)
        On Error GoTo 0
        If ws Is Nothing Then
            MsgBox "Sheet not found: " & TARGET_SHEET_NAME, vbCritical
            Exit Sub
        End If
    Else
        Set ws = ActiveSheet
    End If

    ' Read the model's current Unit System so it can be restored afterwards,
    ' then switch to CAPTURE_FORCE_UNIT/CAPTURE_DIST_UNIT for these captures.
    If Not GetUnitSystem(origForce, origDist, origHeat, origTemper) Then
        MsgBox "Could not read the current Unit System (db/UNIT) - aborting before changing anything.", vbCritical
        Exit Sub
    End If

    If Not SetUnitSystem(CAPTURE_FORCE_UNIT, CAPTURE_DIST_UNIT, origHeat, origTemper) Then
        MsgBox "Could not switch the Unit System to " & CAPTURE_FORCE_UNIT & "/" & CAPTURE_DIST_UNIT & " - aborting.", vbCritical
        Exit Sub
    End If
    unitsChanged = True

    If Not ReadElementListsFromInput(ELEMENT_LIST_FOUND, ELEMENT_LIST_WALL) Then
        Call RestoreUnitsAndReport(origForce, origDist, origHeat, origTemper)
        MsgBox "Could not read element lists from the ""INPUT"" sheet (G10 = foundation, G11 = wall).", vbCritical
        Exit Sub
    End If

    On Error GoTo RestoreUnitsAndFail

    ' Suppress repaints, recalculation and events for the capture loop.
    ' Every ReplacePictureByName deletes and re-adds a Shape, which forces
    ' two repaints and walks the workbook's formula graph each time. Armed
    ' only AFTER the error handler above, so any failure still restores.
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    fastMode = True

    jobs = ParseJobList(JOB_LIST)

    For i = LBound(jobs) To UBound(jobs)

        exportPath = outFolder & jobs(i).shapeName & ".jpg"
        label = jobs(i).CombName & " (" & jobs(i).CompName & ")"

        body = BuildCaptureBody(exportPath, jobs(i))

        Call SendCaptureRequest(body, resp, statusCode)

        If IsApiSuccess(statusCode, resp) Then

            If WaitForFile(exportPath, 8) Then

                If ReplacePictureByName(ws, jobs(i).shapeName, exportPath) Then
                    okCount = okCount + 1
                    okLog = okLog & "  - " & label & vbCrLf
                Else
                    warnCount = warnCount + 1
                    warnLog = warnLog & "  - " & label & ": captured, but no shape named """ & _
                               jobs(i).shapeName & """ found on """ & ws.Name & """" & vbCrLf
                End If

            Else
                failCount = failCount + 1
                failLog = failLog & "  - " & label & ": API accepted the request, but no file appeared" & _
                          ShortApiError(resp) & vbCrLf
            End If

        Else
            failCount = failCount + 1
            failLog = failLog & "  - " & label & ": request failed (HTTP " & statusCode & ")" & HttpStatusHint(statusCode) & _
                      ShortApiError(resp) & vbCrLf
        End If

    Next i

    On Error GoTo 0

    If fastMode Then Call EndFastMode(prevCalc)

    MsgBox "Wingwall plate forces capture  [" & SCRIPT_VERSION & "]" & vbCrLf & vbCrLf & _
           BuildSummaryReport(okCount + warnCount + failCount, okCount, warnCount, failCount, _
                               okLog, warnLog, failLog) & _
           vbCrLf & RestoreUnitsAndReport(origForce, origDist, origHeat, origTemper), _
           IIf(failCount > 0, vbExclamation, vbInformation)

    Exit Sub

RestoreUnitsAndFail:
    Dim failMsg As String
    Dim restoreMsg As String
    failMsg = Err.Description
    If fastMode Then Call EndFastMode(prevCalc)
    If unitsChanged Then
        restoreMsg = RestoreUnitsAndReport(origForce, origDist, origHeat, origTemper)
    Else
        restoreMsg = "(Unit System was never changed.)"
    End If
    MsgBox "Unexpected error: " & failMsg & vbCrLf & restoreMsg, vbCritical

End Sub


' Builds the MsgBox report text: a one-line summary, then results grouped
' by outcome (OK / WARNINGS / FAILED) instead of interleaved in job order.
Private Function BuildSummaryReport(ByVal totalCount As Long, ByVal okCount As Long, _
                                    ByVal warnCount As Long, ByVal failCount As Long, _
                                    ByVal okLog As String, ByVal warnLog As String, _
                                    ByVal failLog As String) As String

    Dim r As String

    r = totalCount & " capture" & IIf(totalCount = 1, "", "s") & ": " & okCount & " OK"
    If warnCount > 0 Then r = r & ", " & warnCount & " warning" & IIf(warnCount = 1, "", "s")
    If failCount > 0 Then r = r & ", " & failCount & " failed"
    r = r & vbCrLf & String(40, "-")

    If okCount > 0 Then r = r & vbCrLf & vbCrLf & "OK:" & vbCrLf & okLog
    If warnCount > 0 Then r = r & vbCrLf & "WARNINGS:" & vbCrLf & warnLog
    If failCount > 0 Then r = r & vbCrLf & "FAILED:" & vbCrLf & failLog

    BuildSummaryReport = r

End Function

' Pulls a short human-readable reason out of a MIDAS API JSON response
' instead of dumping the raw JSON body into the report. "" if nothing usable.
Private Function ShortApiError(ByVal resp As String) As String

    Dim msg As String
    msg = ExtractJsonStringValue(resp, "message")

    If Len(msg) > 0 Then
        ShortApiError = " - " & msg
    ElseIf Len(resp) > 0 Then
        ShortApiError = " - " & Left(resp, 120)
    Else
        ShortApiError = ""
    End If

End Function


' PUTs the Unit System back to (force, dist, heat, temper), then GETs it
' again to confirm the change actually took (a failed PUT still returns a
' status - this catches it instead of assuming success). Returns a message
' describing what happened, for the caller to show the user.
Private Function RestoreUnitsAndReport(ByVal force As String, ByVal dist As String, _
                                       ByVal heat As String, ByVal temper As String) As String

    Dim putOk As Boolean
    Dim checkForce As String, checkDist As String, checkHeat As String, checkTemper As String
    Dim readOk As Boolean

    putOk = SetUnitSystem(force, dist, heat, temper)
    readOk = GetUnitSystem(checkForce, checkDist, checkHeat, checkTemper)

    If putOk And readOk And StrComp(checkDist, dist, vbTextCompare) = 0 _
       And StrComp(checkForce, force, vbTextCompare) = 0 Then
        RestoreUnitsAndReport = "Unit System restored to " & force & "/" & dist & "."
    Else
        RestoreUnitsAndReport = "WARNING: Unit System restore may have failed - tried to set FORCE=""" & _
            force & """/DIST=""" & dist & """, but a follow-up check now reads FORCE=""" & checkForce & _
            """/DIST=""" & checkDist & """. Set it back manually in Civil NX (Tools > Unit System) if needed."
    End If

End Function


' ===========================================================================
'  ELEMENT LISTS (from the "INPUT" sheet)
' ===========================================================================

' Reads the foundation (G10) and wall (G11) element lists off the "INPUT"
' sheet - kept out of source so the same script works unmodified across
' workbooks whose wingwall element numbers differ. Returns False (and
' leaves both ByRef args untouched) if the "INPUT" sheet doesn't exist.
Private Function ReadElementListsFromInput(ByRef foundList As String, _
                                           ByRef wallList As String) As Boolean

    Dim wsInput As Worksheet

    On Error Resume Next
    Set wsInput = ThisWorkbook.Sheets("INPUT")
    On Error GoTo 0

    If wsInput Is Nothing Then
        ReadElementListsFromInput = False
        Exit Function
    End If

    foundList = Trim(CStr(wsInput.Range("G10").Value))
    wallList = Trim(CStr(wsInput.Range("G11").Value))

    ReadElementListsFromInput = True

End Function


' ===========================================================================
'  UNIT SYSTEM (db/UNIT)
'  "Unit System" JSON Manual, ed. 2024.08.01. GET returns {"UNIT": {FORCE,
'  DIST, HEAT, TEMPER}} flat; PUT expects {"Assign": {"1": {...}}}.
' ===========================================================================

' Reads the model's current FORCE/DIST/HEAT/TEMPER units. Returns False if
' the request failed or DIST couldn't be found in the response.
Private Function GetUnitSystem(ByRef force As String, ByRef dist As String, _
                               ByRef heat As String, ByRef temper As String) As Boolean

    Dim resp As String
    Dim statusCode As Long

    Call SendUnitRequest("GET", "", resp, statusCode)

    If statusCode <> 200 Then
        GetUnitSystem = False
        Exit Function
    End If

    force = ExtractJsonStringValue(resp, "FORCE")
    dist = ExtractJsonStringValue(resp, "DIST")
    heat = ExtractJsonStringValue(resp, "HEAT")
    temper = ExtractJsonStringValue(resp, "TEMPER")

    GetUnitSystem = (Len(dist) > 0)

End Function


' Sets FORCE/DIST/HEAT/TEMPER via PUT db/UNIT. Returns True on HTTP 200.
Private Function SetUnitSystem(ByVal force As String, ByVal dist As String, _
                               ByVal heat As String, ByVal temper As String) As Boolean

    Dim body As String
    Dim resp As String
    Dim statusCode As Long

    body = "{""Assign"": {""1"": {"
    body = body & """FORCE"": """ & force & ""","
    body = body & """DIST"": """ & dist & ""","
    body = body & """HEAT"": """ & heat & ""","
    body = body & """TEMPER"": """ & temper & """"
    body = body & "}}}"

    Call SendUnitRequest("PUT", body, resp, statusCode)

    SetUnitSystem = (statusCode = 200)

End Function


' Pulls the string value of "key": "value" out of a flat JSON response.
' Not a general JSON parser - only safe for db/UNIT's flat string fields
' and for pulling a top-level/nested "message" out of an error response.
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


' GETs or PUTs db/UNIT. For GET, pass body = "".
Private Sub SendUnitRequest(ByVal httpMethod As String, ByVal body As String, _
                            ByRef responseText As String, ByRef statusCode As Long)

    Dim url As String

    url = API_BASE_URL & "/db/UNIT"

    If Len(MapiKey()) = 0 Then
        responseText = MapiKeyProblem()
        statusCode = 0
        Exit Sub
    End If

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
        ' Never reuse a client that just failed.
        Set HTTP_CLIENT = Nothing
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    statusCode = HTTP_CLIENT.Status
    responseText = HTTP_CLIENT.ResponseText

End Sub


' Parses JOB_LIST ("Comb|Type|MinMax|Comp|Shape;...") into an array.
Private Function ParseJobList(ByVal listStr As String) As CaptureJob()

    Dim rows() As String
    Dim fields() As String
    Dim result() As CaptureJob
    Dim i As Integer

    rows = Split(listStr, ";")
    ReDim result(LBound(rows) To UBound(rows))

    For i = LBound(rows) To UBound(rows)
        fields = Split(rows(i), "|")

        ' A missing "|" used to surface as "Subscript out of range"
        ' on the field read below, naming neither the row nor the cause.
        If UBound(fields) < 4 Then
            Err.Raise vbObjectError + 513, "ParseJobList", _
                "JOB_LIST row " & (i + 1) & " has " & (UBound(fields) + 1) & _
                " field(s), expected 5. Check the ""|"" separators in:  " & rows(i)
        End If
        result(i).CombName = Trim(fields(0))
        result(i).CombType = Trim(fields(1))
        result(i).MinMax = Trim(fields(2))
        result(i).CompName = Trim(fields(3))
        result(i).shapeName = Trim(fields(4))
        result(i).ElemList = ElementListForShape(result(i).shapeName)
    Next i

    ParseJobList = result

End Function

' Picks ELEMENT_LIST_WALL or ELEMENT_LIST_FOUND based on whether shapeName
' contains "dvr" (Duvar/wall) or "tml" (Temel/foundation); "" (no ACTIVE
' filter, whole model shown) if neither substring is present.
Private Function ElementListForShape(ByVal shapeName As String) As String

    If InStr(1, shapeName, "dvr", vbTextCompare) > 0 Then
        ElementListForShape = ELEMENT_LIST_WALL
    ElseIf InStr(1, shapeName, "tml", vbTextCompare) > 0 Then
        ElementListForShape = ELEMENT_LIST_FOUND
    Else
        ElementListForShape = ""
    End If

End Function

' "Local" or "UCS" for this job - same "dvr"/"tml" split as the element list
' above. See the LOCAL_UCS_TYPE_* block for why the foundation differs.
Private Function CoordSystemForShape(ByVal shapeName As String) As String

    If InStr(1, shapeName, "dvr", vbTextCompare) > 0 Then
        CoordSystemForShape = LOCAL_UCS_TYPE_DVR
    ElseIf InStr(1, shapeName, "tml", vbTextCompare) > 0 Then
        CoordSystemForShape = LOCAL_UCS_TYPE_TML
    Else
        CoordSystemForShape = LOCAL_UCS_TYPE_OTHER
    End If

End Function


' ===========================================================================
'  JSON BODY
'  Note: VBA caps a statement at 24 line continuations, which is why the
'  JSON is appended chunk by chunk instead of one long "& _" chain.
' ===========================================================================
Private Function BuildCaptureBody(ByVal exportPath As String, _
                                  ByRef job As CaptureJob) As String

    Dim b As String

    b = "{""Argument"": {"

    ' ---- capture settings ----
    b = b & """SET_MODE"": ""post"","
    b = b & """SET_HIDDEN"": false,"
    b = b & """EXPORT_PATH"": """ & Replace(exportPath, "\", "\\") & ""","
    b = b & """WIDTH"": " & IMG_WIDTH & ","
    b = b & """HEIGHT"": " & IMG_HEIGHT & ","
    b = b & """ZOOM_LEVEL"": " & ZOOM_LEVEL & ","

    ' Plain white background (both gradient stops), per Capture doc specs
    ' #12/#13 (BGCOLOR_TOP / BGCOLOR_BOTTOM).
    b = b & """BGCOLOR_TOP"": {""R"": 255, ""G"": 255, ""B"": 255},"
    b = b & """BGCOLOR_BOTTOM"": {""R"": 255, ""G"": 255, ""B"": 255},"

    ' Activate only the wall or foundation elements, per job.ElemList
    ' (resolved from "dvr"/"tml" in ShapeName - see ElementListForShape).
    If Len(job.ElemList) > 0 Then
        b = b & """ACTIVE"": {"
        b = b & """ACTIVE_MODE"": ""Active"","
        b = b & """E_LIST"": " & ElementListToJsonArray(job.ElemList)
        b = b & "},"
    End If

    ' Rotate the camera view (per Capture doc's own Argument.ANGLE.
    ' HORIZONTAL/VERTICAL example) - same "dvr"/"tml" split as the element
    ' list above.
    If InStr(1, job.shapeName, "dvr", vbTextCompare) > 0 Then
        b = b & """ANGLE"": {""HORIZONTAL"": " & VIEW_ANGLE_HORIZONTAL_FOR_DVR & _
                 ", ""VERTICAL"": " & VIEW_ANGLE_VERTICAL_FOR_DVR & "},"
    ElseIf InStr(1, job.shapeName, "tml", vbTextCompare) > 0 Then
        b = b & """ANGLE"": {""HORIZONTAL"": " & VIEW_ANGLE_HORIZONTAL_FOR_TML & _
                 ", ""VERTICAL"": " & VIEW_ANGLE_VERTICAL_FOR_TML & "},"
    End If

    ' Perspective projection - "dvr" (wall) rows only.
    If PERSPECTIVE_FOR_DVR And InStr(1, job.shapeName, "dvr", vbTextCompare) > 0 Then
        b = b & """PERSPECTIVE"": true,"
    End If

    ' Draws the UCS axis triad on the picture itself, "tml" rows only - see
    ' SHOW_UCS_AXIS's own comment.
    If SHOW_UCS_AXIS And InStr(1, job.shapeName, "tml", vbTextCompare) > 0 Then
        b = b & """DISPLAY"": {""VIEW"": {""UCS_AXIS"": true}},"
    End If

    ' ---- result display ----
    b = b & """RESULT_GRAPHIC"": {"
    b = b & """CURRENT_MODE"": ""PlateForces/Moments"","

    ' load case / combination - STEP_INDEX is Required by the doc's spec
    ' table even for plain static ("ST") cases; STEP_INDEX_DEFAULT (1) is
    ' the value the doc's own worked ST examples use.
    b = b & """LOAD_CASE_COMB"": {"
    b = b & """TYPE"": """ & job.CombType & ""","
    ' MINMAX only sent for true envelope combinations ("ENV_..." names) -
    ' ULS/SLS are plain combinations, not envelopes, so sending "Max" on
    ' them would imply an envelope semantic that doesn't apply.
    If Left(job.CombName, 4) = "ENV_" Then
        b = b & """MINMAX"": """ & job.MinMax & ""","
    End If
    b = b & """NAME"": """ & job.CombName & ""","
    b = b & """STEP_INDEX"": " & STEP_INDEX_DEFAULT
    b = b & "},"

    ' plate force options - coordinate system + averaging method. The
    ' coordinate system is per job: wall rows Local, foundation rows UCS
    ' (see CoordSystemForShape and the LOCAL_UCS_TYPE_* block).
    b = b & """OPTIONS"": {"
    b = b & """LOCAL_UCS"": {"
    b = b & """TYPE"": """ & CoordSystemForShape(job.shapeName) & """"
    If StrComp(CoordSystemForShape(job.shapeName), "UCS", vbTextCompare) = 0 Then
        b = b & ","
        b = b & """UCS_NAME"": """ & UCS_NAME_FOR_RESULTS & """"
    End If
    b = b & "},"
    b = b & """AVERAGE_NODAL"": {"
    b = b & """TYPE"": """ & AVERAGE_NODAL_TYPE & """"
    b = b & "}"
    b = b & "},"

    ' component
    b = b & """COMPONENTS"": {"
    b = b & """COMP"": """ & job.CompName & """"
    b = b & "},"

    ' ---- contour / values / legend ----
    ' NUM_OF_COLOR must be 6 / 12 / 18 / 24.
    ' VALUE_EXP defaults to true, which is why untouched legends print
    ' in exponential notation. UNDEFORMED is a plain boolean here (not the
    ' {"OPT_CHECK": ...} object form DisplacementContour uses) - confirmed
    ' against the cached "Plate Forces/Moments - Result Display" JSON Manual
    ' (ed. 2025.08.19): both its JSON Schema block and its Specifications
    ' table (item 5.(5)) agree UNDEFORMED is Boolean, default false, for
    ' this result mode specifically. The object form caused every capture
    ' to fail with "MIDAS CIVIL NX second query is wrong" (HTTP 200, no
    ' file ever written) - RESULT_GRAPHIC schema validation rejecting the
    ' malformed field. Don't copy DisplacementContour's object form back in
    ' without re-checking this doc; the two result modes disagree here.
    b = b & """TYPE_OF_DISPLAY"": {"
    b = b & """CONTOUR"": {"
    b = b & """OPT_CHECK"": true,"
    b = b & """NUM_OF_COLOR"": 12,"
    b = b & """COLOR_TYPE"": ""rgb"","
    b = b & """OPTIONS"": {"
    b = b & """CONTOUR_FILL"": true,"
    b = b & """GRADIENT_FILL"": false"
    b = b & "}"
    b = b & "},"
    b = b & """VALUES"": {"
    b = b & """OPT_CHECK"": true,"
    b = b & """VALUE_EXP"": false,"
    b = b & """DECIMAL_PT"": 0,"
    b = b & """SET_ORIENT"": " & VALUE_TEXT_ORIENT
    b = b & "},"
    b = b & """LEGEND"": {"
    b = b & """OPT_CHECK"": true,"
    b = b & """POSITION"": ""right"","
    b = b & """VALUE_EXP"": false,"
    b = b & """DECIMAL_PT"": 1"
    b = b & "},"
    b = b & """UNDEFORMED"": true"
    b = b & "}"

    b = b & "}"
    b = b & "}}"

    BuildCaptureBody = b

End Function


' Converts a MIDAS-style element list ("4 10to13 17to19 21to44", commas
' also accepted as separators) into a JSON number array "[4,10,11,...]".
Private Function ElementListToJsonArray(ByVal listStr As String) As String

    Dim tokens() As String
    Dim i As Integer
    Dim tok As String
    Dim toPos As Integer, byPos As Integer
    Dim startNum As Long, endNum As Long, stepNum As Long, n As Long
    Dim endPart As String
    Dim result As String

    listStr = Replace(listStr, ",", " ")
    tokens = Split(Trim(listStr), " ")

    For i = LBound(tokens) To UBound(tokens)
        tok = Trim(tokens(i))
        If Len(tok) > 0 Then
            toPos = InStr(1, tok, "to", vbTextCompare)
            If toPos > 0 Then
                startNum = CLng(Left(tok, toPos - 1))
                endPart = Mid(tok, toPos + 2)

                ' "AtoBbyN" step shorthand (e.g. "1850to2285by15") - MIDAS
                ' range-with-increment syntax, distinct from the plain
                ' "AtoB" (every integer) form.
                byPos = InStr(1, endPart, "by", vbTextCompare)
                If byPos > 0 Then
                    endNum = CLng(Left(endPart, byPos - 1))
                    stepNum = CLng(Mid(endPart, byPos + 2))
                Else
                    endNum = CLng(endPart)
                    stepNum = 1
                End If

                ' A zero or negative step would loop forever and hang
                ' Excel outright - "1to10by0" is a typo, not a request.
                If stepNum <= 0 Then stepNum = 1

                For n = startNum To endNum Step stepNum
                    If Len(result) > 0 Then result = result & ","
                    result = result & n
                Next n
            Else
                If Len(result) > 0 Then result = result & ","
                result = result & tok
            End If
        End If
    Next i

    ElementListToJsonArray = "[" & result & "]"

End Function


' ===========================================================================
'  EXCEL PICTURE REPLACEMENT
' ===========================================================================

' Deletes the existing shape called shapeName on ws (if any) and inserts
' imagePath in its place, keeping the same position/size and re-applying
' the same name so later runs keep finding it. Returns True if a shape
' was found and replaced; False if no such shape existed (nothing done
' except reporting - it will not guess where to put a brand new picture).
Private Function ReplacePictureByName(ByVal ws As Worksheet, _
                                      ByVal shapeName As String, _
                                      ByVal imagePath As String) As Boolean

    Dim shp As Shape
    Dim l As Single, t As Single, w As Single, h As Single
    Dim found As Boolean
    Dim newShp As Shape

    found = False

    For Each shp In ws.Shapes
        If StrComp(shp.Name, shapeName, vbTextCompare) = 0 Then
            l = shp.Left
            t = shp.Top
            w = shp.Width
            h = shp.Height
            shp.Delete
            found = True
            Exit For
        End If
    Next shp

    If Not found Then
        ReplacePictureByName = False
        Exit Function
    End If

    Set newShp = ws.Shapes.AddPicture(imagePath, msoFalse, msoCTrue, l, t, w, h)
    newShp.Name = shapeName

    ReplacePictureByName = True

End Function


' True only when the API both returned a 2xx AND did not put an "error"
' object in the body. MIDAS answers a bad request with HTTP 200 and
' {"error":{"message":"..."}}, writing no file - checking the status alone
' sends every rejected capture into the WaitForFile timeout below.
Private Function IsApiSuccess(ByVal statusCode As Long, ByVal responseText As String) As Boolean

    If statusCode < 200 Or statusCode > 299 Then Exit Function
    If InStr(1, responseText, """error""", vbTextCompare) > 0 Then Exit Function

    IsApiSuccess = True

End Function


' Puts Excel's screen updating, events and calculation back the way they
' were. Called on BOTH the normal exit and the error handler - if this is
' ever missed, the workbook is left with calculation switched off, which
' looks exactly like corrupted results. On Error Resume Next so a failure
' restoring one setting still lets the other two through.
Private Sub EndFastMode(ByVal prevCalc As XlCalculation)

    On Error Resume Next
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ' Release the shared client; the next run builds a fresh one.
    Set HTTP_CLIENT = Nothing
    On Error GoTo 0

End Sub


' Polls for a file to appear, up to timeoutSeconds.
'
' DoEvents rather than Application.Wait: Application.Wait cannot resolve
' below one second, so a render finishing in 200ms was still billed a full
' second, and it blocks Excel's thread outright ("Not Responding"). This
' version picks the file up within milliseconds of it appearing and keeps
' the UI alive while it waits.
Private Function WaitForFile(ByVal filePath As String, ByVal timeoutSeconds As Long) As Boolean

    Dim startTime As Single

    startTime = Timer

    Do While Len(Dir(filePath)) = 0
        ' Timer resets at midnight - without this the wait never times out.
        If Timer < startTime Then startTime = startTime - 86400!
        If (Timer - startTime) >= timeoutSeconds Then
            WaitForFile = (Len(Dir(filePath)) > 0)
            Exit Function
        End If
        DoEvents
    Loop

    WaitForFile = True

End Function


' POSTs the JSON body and returns the response text plus HTTP status.
' statusCode = 0 means the request never completed (network / TLS error).
Private Sub SendCaptureRequest(ByVal body As String, _
                               ByRef responseText As String, _
                               ByRef statusCode As Long)

    Dim url As String

    url = API_BASE_URL & "/view/CAPTURE"

    If Len(MapiKey()) = 0 Then
        responseText = MapiKeyProblem()
        statusCode = 0
        Exit Sub
    End If

    ' Reuse one client for the whole run rather than paying COM
    ' instantiation and a fresh connection on every capture.
    If HTTP_CLIENT Is Nothing Then
        Set HTTP_CLIENT = CreateObject("WinHttp.WinHttpRequest.5.1")
    End If

    On Error Resume Next

    HTTP_CLIENT.Open "POST", url, False
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

    HTTP_CLIENT.Send body
    If Err.Number <> 0 Then
        responseText = "WinHTTP error: " & Err.Description
        statusCode = 0
        Err.Clear
        ' Never reuse a client that just failed.
        Set HTTP_CLIENT = Nothing
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    statusCode = HTTP_CLIENT.Status
    responseText = HTTP_CLIENT.ResponseText

End Sub


' ===========================================================================
'  HELPERS
' ===========================================================================

' Workbook file name without its extension, e.g. "mnf151.xlsm" -> "mnf151".
Private Function WorkbookBaseName() As String

    Dim wbName As String
    Dim dotPos As Long

    wbName = ThisWorkbook.Name
    dotPos = InStrRev(wbName, ".")

    If dotPos > 0 Then
        WorkbookBaseName = Left(wbName, dotPos - 1)
    Else
        WorkbookBaseName = wbName
    End If

End Function

' Creates a folder plus any missing parent folders.
' Handles drive-letter paths (C:\a\b\c). UNC paths are not supported.
Private Function EnsureFolderExists(ByVal folderPath As String) As Boolean

    Dim parts() As String
    Dim i As Integer
    Dim built As String

    EnsureFolderExists = False

    folderPath = Replace(folderPath, "/", "\")

    Do While Right(folderPath, 1) = "\"
        folderPath = Left(folderPath, Len(folderPath) - 1)
    Loop

    If Len(folderPath) = 0 Then Exit Function

    parts = Split(folderPath, "\")
    built = parts(LBound(parts))          ' drive letter, e.g. "C:"

    On Error GoTo Failed

    For i = LBound(parts) + 1 To UBound(parts)
        built = built & "\" & parts(i)
        If Dir(built, vbDirectory) = "" Then
            MkDir built
        End If
    Next i

    EnsureFolderExists = (Dir(folderPath, vbDirectory) <> "")
    Exit Function

Failed:
    EnsureFolderExists = False

End Function


