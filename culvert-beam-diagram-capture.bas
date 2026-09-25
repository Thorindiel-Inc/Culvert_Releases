Option Explicit

' ============================================================================
'  MIDAS Civil NX API - Beam Diagram Screenshot Capture
' ----------------------------------------------------------------------------
'  Endpoint : {base url} + view/CAPTURE
'  Docs     : "Capture"                        (JSON Manual, ed. 2025.04.24)
'             "Beam Diagrams - Result Display" (JSON Manual, ed. 2026.01.06)
'             "Type of Display"                (JSON Manual, ed. 2026.06.19)
'
'  Captures one beam diagram image per job in JOB_LIST (combination +
'  component + target shape + zoom), replacing the matching Picture/Shape
'  on TARGET_SHEET_NAME with the freshly captured image.
' ============================================================================


' ---------------------------------------------------------------------------
'  CONFIG
' ---------------------------------------------------------------------------

' Printed in the summary MsgBox. These modules are pasted into Excel by
' hand, so the file in the repo and the code actually running can silently
' diverge - check this stamp matches the constant here before concluding
' anything from a run. Bump it whenever this file changes.
Private Const SCRIPT_VERSION As String = "2026-09-25a"

' One-line summary of what changed in THIS version, shown by the updater
' next to this module when it's stale. Update alongside SCRIPT_VERSION -
' must stay on ONE physical line (no "_" continuation - the parser that
' reads this out does not resolve continuations) and must not contain "|"
' (breaks manifest.txt's pipe-delimited format).
Private Const SCRIPT_CHANGELOG As String = "Calculation always goes back to Automatic at the end (it used to restore the starting mode, so one interrupted run left Excel on Manual for good)"

' Identifies this module to the updater regardless of what it was
' named when pasted into Excel - these files carry no VB_Name, so the
' module name in the VBA project is whatever the user typed.
Private Const SCRIPT_ID As String = "culvert-beam-diagram-capture"

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
' <workbook folder>\mnf151\RESULTS\. Created automatically if missing.
Private Const CAPTURE_SUBFOLDER As String = "RESULTS"

' Sheet that holds the figures to replace. Leave empty to use whichever
' sheet is active when the macro runs.
Private Const TARGET_SHEET_NAME As String = "5_SONUC"

' One capture per row: CombName|CombType|MinMax|Component|ShapeName|ZoomLevel|ScaleFactor
'   CombType   : "ST" "CS" "RS" "TH" "MV" "SM" "CB"
'   MinMax     : "Min" "Max" "All"
'   ZoomLevel  : 25<=v<100 zoom out, 100 fit, 100<v<=200 zoom in
'   ScaleFactor: diagram height multiplier (DISPLAY_OPTIONS.SCALE). Raise to
'                2.0 / 3.0 to exaggerate the plot; per-row since force and
'                moment diagrams often need different exaggeration.
' Rows separated by ";". EDIT to add/remove captures. Shape names must
' match existing Pictures on TARGET_SHEET_NAME (Name Box, top-left, when
' the picture is selected - or Home > Find & Select > Selection Pane).
Private Const JOB_LIST As String = _
    "ENV_ALL|CB|All|Fx|STR Fx|95|3;" & _
    "ENV_ALL|CB|All|Fz|STR Fz|90|3;" & _
    "ENV_ALL|CB|All|My|STR My|95|3;" & _
    "ENV_SER|CB|All|Fx|SER Fx|95|3;" & _
    "ENV_SER|CB|All|My|SER My|95|3"

' Image size in pixels, matched to the Excel figure boxes' physical size
' (33.16 x 19.61 cm) at 120 DPI, so the captured image isn't stretched/
' squished when ReplacePictureByName fits it into that same box.
Private Const IMG_WIDTH As Long = 1269
Private Const IMG_HEIGHT As Long = 750

' Elements to activate before capture (only these are shown/plotted).
' Space- or comma-separated; supports MIDAS "AtoB" range shorthand,
' e.g. "4 10to13 17to19 21to44" expands to 4,10,11,12,13,17,18,19,21..44.
' Empty shows/plots everything (skips sending ACTIVE at all). Read at
' runtime from ELEMENT_LIST_SHEET_NAME!ELEMENT_LIST_CELL (not a hardcoded
' Const, since VBA constants can't hold a formula/cell-derived value) -
' see ReadElementListFromSheet, called once per run from
' CaptureBeamDiagrams before the capture loop.
Private Const ELEMENT_LIST_SHEET_NAME As String = "INPUT"
Private Const ELEMENT_LIST_CELL As String = "G15"
Private ELEMENT_LIST As String

' Force/length units to switch to for these captures (db/UNIT, "Unit System"
' JSON Manual, ed. 2024.08.01: FORCE "N" "KN" "KGF" "TONF" "LBF" "KIPS",
' DIST "M" "CM" "MM" "FT" "IN"). The macro reads the model's current Unit
' System first, switches to these values, runs all captures, then always
' switches back to whatever it originally was - see CaptureBeamDiagrams.
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
    ShapeName As String
    ZoomLevel As Long
    ScaleFactor As String
End Type


' ===========================================================================
'  MAIN - one image per job in JOB_LIST
' ===========================================================================
Sub CaptureBeamDiagrams()

    Dim jobs() As CaptureJob
    Dim i As Integer
    Dim body As String
    Dim outFolder As String
    Dim exportPath As String
    Dim ws As Worksheet
    Dim origForce As String, origDist As String, origHeat As String, origTemper As String
    Dim unitsChanged As Boolean
    Dim unitWarn As String
    Dim fastMode As Boolean
    Dim report As String

    Dim label As String
    Dim okLog As String, warnLog As String, skipLog As String, failLog As String
    Dim okCount As Long, warnCount As Long, skipCount As Long, failCount As Long

    ' Re-read INPUT!J20 on every run, so a key rotated since the last run
    ' is actually used.
    MAPI_KEY_CACHE = ""

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

    If Not ReadElementListFromSheet(ELEMENT_LIST) Then
        MsgBox "Could not read the element list from """ & ELEMENT_LIST_SHEET_NAME & "!" & _
               ELEMENT_LIST_CELL & """ - aborting.", vbCritical
        Exit Sub
    End If

    ' Read the model's current Unit System so it can be restored afterwards,
    ' then switch to CAPTURE_FORCE_UNIT/CAPTURE_DIST_UNIT for these captures.
    If Not GetUnitSystem(origForce, origDist, origHeat, origTemper) Then
        MsgBox "Could not read the current Unit System (db/UNIT) - aborting before changing anything.", vbCritical
        Exit Sub
    End If

    ' A failed switch does not stop the run - the report says which units
    ' the pictures are really in.
    If SetUnitSystem(CAPTURE_FORCE_UNIT, CAPTURE_DIST_UNIT, origHeat, origTemper) Then
        unitWarn = UnitSwitchCheck(CAPTURE_FORCE_UNIT, CAPTURE_DIST_UNIT)
    Else
        unitWarn = "WARNING: the switch to " & CAPTURE_FORCE_UNIT & "/" & CAPTURE_DIST_UNIT & _
                   " was rejected - the pictures show values in " & origForce & "/" & origDist & "."
    End If
    unitsChanged = True

    On Error GoTo RestoreUnitsAndFail

    ' Suppress repaints, recalculation and events for the capture loop.
    ' Every ReplacePictureByName deletes and re-adds a Shape, which forces
    ' two repaints and walks the workbook's formula graph each time. Armed
    ' only AFTER the error handler above, so any failure still restores.
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    fastMode = True

    jobs = ParseJobList(JOB_LIST)

    For i = LBound(jobs) To UBound(jobs)

        exportPath = outFolder & jobs(i).ShapeName & ".jpg"
        label = jobs(i).CombName & " (" & jobs(i).CompName & ")"

        If Not CaseInModel(jobs(i).CombType, jobs(i).CombName) Then
            skipCount = skipCount + 1
            skipLog = skipLog & "  - " & label & vbCrLf
        Else
            body = BuildCaptureBody(exportPath, jobs(i))
            Call LogOutcome(CaptureToShape(ws, exportPath, body, jobs(i).ShapeName), label, _
                            okCount, okLog, warnCount, warnLog, failCount, failLog)
        End If

    Next i

    On Error GoTo 0

    If fastMode Then Call EndFastMode

    report = "Beam diagram capture  [" & SCRIPT_VERSION & "]" & vbCrLf & _
             RestoreUnitsAndReport(origForce, origDist, origHeat, origTemper) & vbCrLf
    If Len(unitWarn) > 0 Then report = report & unitWarn & vbCrLf
    report = report & vbCrLf & BuildSummaryReport(okCount, warnCount, skipCount, failCount, _
                                                  okLog, warnLog, skipLog, failLog)

    MsgBox FitReport(report, SCRIPT_ID), IIf(failCount > 0, vbExclamation, vbInformation)

    Exit Sub

RestoreUnitsAndFail:
    Dim failMsg As String
    Dim restoreMsg As String
    failMsg = Err.Description
    If fastMode Then Call EndFastMode
    If unitsChanged Then
        restoreMsg = RestoreUnitsAndReport(origForce, origDist, origHeat, origTemper)
    Else
        restoreMsg = "(Unit System was never changed.)"
    End If
    MsgBox "Unexpected error: " & failMsg & vbCrLf & restoreMsg, vbCritical

End Sub


' Turns Excel's screen updating and events back on and sets calculation to
' AUTOMATIC - always, not "whatever it was at the start": a run cut short
' (Esc, VBA Reset, a Civil NX dialog) never reaches this, and restoring the
' starting mode then kept every later run on Manual too (owner, 2026-09-25).
' Called on BOTH the normal exit and the error handler - if this is ever
' missed, the workbook is left with calculation switched off, which looks
' exactly like corrupted results. On Error Resume Next so a failure
' restoring one setting still lets the other two through.
Private Sub EndFastMode()

    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ' Release the shared client; the next run builds a fresh one.
    Set HTTP_CLIENT = Nothing
    On Error GoTo 0

End Sub


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
'  UNIT SYSTEM (db/UNIT)
'  "Unit System" JSON Manual, ed. 2024.08.01. GET and PUT use the same
'  nesting: {"UNIT": {"1": {FORCE, DIST, HEAT, TEMPER}}} back, and
'  {"Assign": {"1": {...}}} to write (confirmed live 2026-09-22).
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

    ' The body too, not just the status: MIDAS rejects with HTTP 200 plus
    ' an {"error":...} body.
    SetUnitSystem = IsApiSuccess(statusCode, resp)

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

' Pulls the string value of "key": "value" out of a flat JSON response.
' Not a general JSON parser - only safe for db/UNIT's flat string fields.
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

    Call SendDbRequest(httpMethod, "db/UNIT", body, responseText, statusCode)

End Sub


' Parses JOB_LIST ("Comb|Type|MinMax|Comp|Shape|Zoom;...") into an array.
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
        If UBound(fields) < 6 Then
            Err.Raise vbObjectError + 513, "ParseJobList", _
                "JOB_LIST row " & (i + 1) & " has " & (UBound(fields) + 1) & _
                " field(s), expected 7. Check the ""|"" separators in:  " & rows(i)
        End If
        result(i).CombName = Trim(fields(0))
        result(i).CombType = Trim(fields(1))
        result(i).MinMax = Trim(fields(2))
        result(i).CompName = Trim(fields(3))
        result(i).ShapeName = Trim(fields(4))
        result(i).ZoomLevel = CLng(Trim(fields(5)))
        result(i).ScaleFactor = Trim(fields(6))
    Next i

    ParseJobList = result

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
    b = b & """ZOOM_LEVEL"": " & job.ZoomLevel & ","

    ' Plain white background (both gradient stops), per Capture doc specs
    ' #12/#13 (BGCOLOR_TOP / BGCOLOR_BOTTOM).
    b = b & """BGCOLOR_TOP"": {""R"": 255, ""G"": 255, ""B"": 255},"
    b = b & """BGCOLOR_BOTTOM"": {""R"": 255, ""G"": 255, ""B"": 255},"

    ' Activate only specific elements, if ELEMENT_LIST is set.
    ' ACTIVE_MODE / N_LIST / E_LIST confirmed from one doc example only -
    ' no dedicated "view/ACTIVE" manual page seen yet.
    If Len(ELEMENT_LIST) > 0 Then
        b = b & """ACTIVE"": {"
        b = b & """ACTIVE_MODE"": ""Active"","
        b = b & """E_LIST"": " & ElementListToJsonArray(ELEMENT_LIST)
        b = b & "},"
    End If

    ' ---- result display ----
    b = b & """RESULT_GRAPHIC"": {"
    b = b & """CURRENT_MODE"": ""BeamDiagrams"","

    ' load case / combination
    b = b & """LOAD_CASE_COMB"": {"
    b = b & """TYPE"": """ & job.CombType & ""","
    b = b & """MINMAX"": """ & job.MinMax & ""","
    b = b & """NAME"": """ & job.CombName & """"
    b = b & "},"

    ' component
    b = b & """COMPONENTS"": {"
    b = b & """PART"": ""Total"","
    b = b & """COMP"": """ & job.CompName & """"
    b = b & "},"

    ' diagram appearance - REQUIRED for BeamDiagrams mode
    '   FIDELITY : "Exact" | "5Points"
    '   FILL     : "No" | "Line" | "Solid"
    b = b & """DISPLAY_OPTIONS"": {"
    b = b & """FIDELITY"": ""5Points"","
    b = b & """FILL"": ""Solid"","
    b = b & """SCALE"": " & job.ScaleFactor
    b = b & "},"

    ' ---- contour / values / legend ----
    ' NUM_OF_COLOR must be 6 / 12 / 18 / 24.
    ' VALUE_EXP defaults to true, which is why untouched legends print
    ' in exponential notation.
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
    b = b & """DECIMAL_PT"": 1,"
    b = b & """SET_ORIENT"": " & VALUE_TEXT_ORIENT
    b = b & "},"
    b = b & """LEGEND"": {"
    b = b & """OPT_CHECK"": true,"
    b = b & """POSITION"": ""right"","
    b = b & """VALUE_EXP"": false,"
    b = b & """DECIMAL_PT"": 1"
    b = b & "}"
    b = b & "},"

    ' ---- where values print along each member ----
    ' The GUI treats "By Element Locations" (I/Center/J) and
    ' "By Element Values" (Abs Max/Min-Max/All) as mutually exclusive -
    ' selecting one greys out the other. Always using the Abs Max group.
    ' OPT_BY_MEMBER is left false: with it on, members render detached
    ' from each other at the joints.
    b = b & """OUTPUT_SECT_LOCATION"": {"
    b = b & """OPT_BY_MEMBER"": false,"
    b = b & """OPT_MAX_MINMAX_ALL"": ""AbsMax"""
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
            ' A typo would otherwise go into the JSON as-is and come back
            ' as MIDAS's generic "second query is wrong".
            If Not IsElementToken(tok) Then
                Err.Raise vbObjectError + 514, "ElementListToJsonArray", _
                    "Element list entry """ & tok & """ is not a number or an AtoB / " & _
                    "AtoBbyN range. Check the element list cell. Full list: " & listStr
            End If
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

' Replaces the shape called shapeName on ws with imagePath, keeping its
' position, size and name. Returns 0 when replaced; 1 when no such shape
' exists (nothing is done - it will not guess where a new picture goes);
' 2 when inserting the new picture failed, with the reason in errText.
'
' The new picture is inserted BEFORE the old one is deleted, so a failed
' insert (a half-written or locked JPEG) leaves the old figure in place.
Private Function ReplacePictureByName(ByVal ws As Worksheet, _
                                      ByVal shapeName As String, _
                                      ByVal imagePath As String, _
                                      ByRef errText As String) As Long

    Dim shp As Shape
    Dim oldShp As Shape
    Dim newShp As Shape

    For Each shp In ws.Shapes
        If StrComp(shp.Name, shapeName, vbTextCompare) = 0 Then
            Set oldShp = shp
            Exit For
        End If
    Next shp

    If oldShp Is Nothing Then
        ReplacePictureByName = 1
        Exit Function
    End If

    On Error Resume Next
    Set newShp = ws.Shapes.AddPicture(imagePath, msoFalse, msoCTrue, _
                                      oldShp.Left, oldShp.Top, oldShp.Width, oldShp.Height)
    If Err.Number <> 0 Or newShp Is Nothing Then
        errText = Err.Description
        Err.Clear
        On Error GoTo 0
        ReplacePictureByName = 2
        Exit Function
    End If
    On Error GoTo 0

    ' Only now that the replacement exists.
    oldShp.Delete
    newShp.Name = shapeName

    ReplacePictureByName = 0

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


' Polls for a file to appear, up to timeoutSeconds. The API call returning
' HTTP 200 doesn't guarantee the file is flushed to disk yet.
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


' ===========================================================================
'  HELPERS
' ===========================================================================

' Reads ELEMENT_LIST_SHEET_NAME!ELEMENT_LIST_CELL into outValue (trimmed;
' empty is valid - it means "show/plot everything", same as before this
' was cell-driven). Returns False (and leaves outValue unchanged) only if
' the sheet itself doesn't exist.
Private Function ReadElementListFromSheet(ByRef outValue As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(ELEMENT_LIST_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        ReadElementListFromSheet = False
        Exit Function
    End If

    outValue = Trim(CStr(ws.Range(ELEMENT_LIST_CELL).Value))
    ReadElementListFromSheet = True

End Function

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


' Builds the report body: a one-line summary, then results grouped by
' outcome - FAILED first, then WARNINGS, SKIPPED and OK - instead of
' interleaved in job order.
Private Function BuildSummaryReport(ByVal okCount As Long, ByVal warnCount As Long, _
                                    ByVal skipCount As Long, ByVal failCount As Long, _
                                    ByVal okLog As String, ByVal warnLog As String, _
                                    ByVal skipLog As String, ByVal failLog As String) As String

    Dim r As String
    Dim totalCount As Long

    totalCount = okCount + warnCount + skipCount + failCount

    r = totalCount & " capture" & IIf(totalCount = 1, "", "s") & ": " & okCount & " OK"
    If failCount > 0 Then r = r & ", " & failCount & " failed"
    If warnCount > 0 Then r = r & ", " & warnCount & " warning" & IIf(warnCount = 1, "", "s")
    If skipCount > 0 Then r = r & ", " & skipCount & " skipped"
    r = r & vbCrLf & String(40, "-")

    ' Most important first: MsgBox cuts a long report off at the bottom.
    If failCount > 0 Then r = r & vbCrLf & vbCrLf & "FAILED:" & vbCrLf & failLog
    If warnCount > 0 Then r = r & vbCrLf & "WARNINGS:" & vbCrLf & warnLog
    If skipCount > 0 Then r = r & vbCrLf & "SKIPPED (not in this model):" & vbCrLf & skipLog
    If okCount > 0 Then r = r & vbCrLf & "OK:" & vbCrLf & okLog

    BuildSummaryReport = r

End Function


' ===========================================================================
'  CAPTURE FLOW - shared by all six capture scripts and kept byte-identical
'  by tests/verify_no_helper_drift.py. Edit every copy together.
' ===========================================================================

' One capture, end to end. Returns "" when the picture was replaced,
' "WARN:<why>" when the image was captured but could not be placed, and
' "FAIL:<why>" otherwise.
'
' The previous run's file is deleted FIRST. The export path is the same on
' every run, so a leftover file would satisfy WaitForFile at once and be
' inserted as if it were new.
Private Function CaptureToShape(ByVal ws As Worksheet, ByVal exportPath As String, _
                                ByVal body As String, ByVal shapeName As String) As String

    Dim resp As String
    Dim statusCode As Long
    Dim placeErr As String

    If Len(Dir(exportPath)) > 0 Then
        On Error Resume Next
        Kill exportPath
        Err.Clear
        On Error GoTo 0
        If Len(Dir(exportPath)) > 0 Then
            CaptureToShape = "FAIL:could not delete the previous capture " & exportPath & _
                             " - is it open in another program?"
            Exit Function
        End If
    End If

    Call SendCaptureRequest(body, resp, statusCode)

    If Not IsApiSuccess(statusCode, resp) Then
        CaptureToShape = "FAIL:request failed (HTTP " & statusCode & ")" & _
                         HttpStatusHint(statusCode) & ShortApiError(resp)
        Exit Function
    End If

    If Not WaitForFile(exportPath, 8) Then
        CaptureToShape = "FAIL:API accepted the request, but no file appeared" & ShortApiError(resp)
        Exit Function
    End If

    Select Case ReplacePictureByName(ws, shapeName, exportPath, placeErr)
        Case 0
            CaptureToShape = ""
        Case 1
            CaptureToShape = "WARN:captured, but no shape named """ & shapeName & _
                             """ found on """ & ws.Name & """"
        Case Else
            CaptureToShape = "FAIL:captured, but inserting the picture failed, so the old " & _
                             "figure was kept (" & placeErr & ")"
    End Select

End Function


' Files one CaptureToShape outcome under OK / WARNINGS / FAILED.
Private Sub LogOutcome(ByVal outcome As String, ByVal label As String, _
                       ByRef okCount As Long, ByRef okLog As String, _
                       ByRef warnCount As Long, ByRef warnLog As String, _
                       ByRef failCount As Long, ByRef failLog As String)

    If Len(outcome) = 0 Then
        okCount = okCount + 1
        okLog = okLog & "  - " & label & vbCrLf
    ElseIf Left$(outcome, 5) = "WARN:" Then
        warnCount = warnCount + 1
        warnLog = warnLog & "  - " & label & ": " & Mid$(outcome, 6) & vbCrLf
    Else
        failCount = failCount + 1
        failLog = failLog & "  - " & label & ": " & Mid$(outcome, 6) & vbCrLf
    End If

End Sub


' False only when the model definitely lacks this load case / combination,
' so the job can be reported as SKIPPED instead of failing with MIDAS's
' generic "second query is wrong". Anything this cannot check (another case
' type, a bare-model row with no name, a failed read) returns True and the
' capture is simply attempted.
'
' Why: the culvert builder writes EQ, ATA, EQ-1 and ENV_EQ only when its
' seismic gate is on, so on a non-seismic culvert those jobs have nothing
' to capture.
Private Function CaseInModel(ByVal caseType As String, ByVal caseName As String) As Boolean

    Dim path As String
    Dim resp As String
    Dim statusCode As Long

    CaseInModel = True
    If Len(caseName) = 0 Then Exit Function

    Select Case UCase$(caseType)
        Case "ST": path = "db/STLD"
        Case "CB": path = "db/LCOM-GEN"
        Case Else: Exit Function
    End Select

    Call SendDbRequest("GET", path, "", resp, statusCode)
    If Not IsApiSuccess(statusCode, resp) Then Exit Function

    If InStr(1, resp, """NAME"":""" & caseName & """", vbTextCompare) = 0 And _
       InStr(1, resp, """NAME"": """ & caseName & """", vbTextCompare) = 0 Then
        CaseInModel = False
    End If

End Function


' "" when the model now reports FORCE/DIST as asked, otherwise a warning
' line for the report. The switch is read back rather than trusted - the
' same check RestoreUnitsAndReport makes on the way out. The captures still
' run either way; the warning says which units the pictures are really in.
Private Function UnitSwitchCheck(ByVal wantForce As String, ByVal wantDist As String) As String

    Dim f As String, d As String, h As String, t As String

    If Not GetUnitSystem(f, d, h, t) Then
        UnitSwitchCheck = "WARNING: could not read the Unit System back after switching - " & _
                          "the values may not be in " & wantForce & "/" & wantDist & "."
    ElseIf StrComp(f, wantForce, vbTextCompare) <> 0 Or StrComp(d, wantDist, vbTextCompare) <> 0 Then
        UnitSwitchCheck = "WARNING: the model is in " & f & "/" & d & ", not " & wantForce & _
                          "/" & wantDist & " - the pictures show values in " & f & "/" & d & "."
    End If

End Function


' GET/PUT/POST/DELETE against {base url}/<path>. statusCode = 0 means the
' request never completed. Shares the run's one WinHTTP client.
Private Sub SendDbRequest(ByVal httpMethod As String, ByVal path As String, ByVal body As String, _
                          ByRef responseText As String, ByRef statusCode As Long)

    Dim url As String

    url = API_BASE_URL & "/" & path

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


' True for "12", "10to13" or "1850to2285by15" (case-insensitive).
Private Function IsElementToken(ByVal tok As String) As Boolean

    Dim p() As String
    Dim q() As String

    p = Split(LCase$(tok), "to")
    If UBound(p) = 0 Then
        IsElementToken = IsDigits(p(0))
        Exit Function
    End If
    If UBound(p) <> 1 Then Exit Function

    q = Split(p(1), "by")
    If UBound(q) > 1 Then Exit Function

    IsElementToken = IsDigits(p(0)) And IsDigits(q(0))
    If UBound(q) = 1 Then IsElementToken = IsElementToken And IsDigits(q(1))

End Function


Private Function IsDigits(ByVal s As String) As Boolean
    IsDigits = (Len(s) > 0 And Not s Like "*[!0-9]*")
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
