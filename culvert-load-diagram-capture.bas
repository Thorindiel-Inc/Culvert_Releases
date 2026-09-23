Option Explicit

' ============================================================================
'  MIDAS Civil NX API - Applied Load Diagram Screenshot Capture
' ----------------------------------------------------------------------------
'  Endpoint : {base url} + view/CAPTURE
'  Docs     : "Capture"  (JSON Manual, ed. 2025.04.24) - Argument.DISPLAY is
'             documented there as "refer to view/DISPLAY manual"
'             "Display" (JSON Manual, ed. 2024.10.25)  - that manual, confirms
'             Argument.DISPLAY.LOAD.{CASE_SELECTION, LOAD_VALUE, BEAM_LOAD,
'             NODAL_LOAD, ...}
'
'  Captures one image per load case in LOAD_JOB_LIST, showing the applied
'  loads on the (unanalyzed) model - no RESULT_GRAPHIC involved, since
'  loads don't need analysis results. Also replaces the matching
'  Picture/Shape on TARGET_SHEET_NAME with the freshly captured image,
'  same pattern as midas-beam-diagram-capture.bas.
' ============================================================================


' ---------------------------------------------------------------------------
'  CONFIG
' ---------------------------------------------------------------------------

' Printed in the summary MsgBox. These modules are pasted into Excel by
' hand, so the file in the repo and the code actually running can silently
' diverge - check this stamp matches the constant here before concluding
' anything from a run. Bump it whenever this file changes.
Private Const SCRIPT_VERSION As String = "2026-09-23g"

' One-line summary of what changed in THIS version, shown by the updater
' next to this module when it's stale. Update alongside SCRIPT_VERSION -
' must stay on ONE physical line (no "_" continuation - the parser that
' reads this out does not resolve continuations) and must not contain "|"
' (breaks manifest.txt's pipe-delimited format).
Private Const SCRIPT_CHANGELOG As String = "Audit rev 2: skips jobs whose load case or combination is not in the model, deletes the old JPEG before capturing, inserts the new picture before removing the old one, warns instead of aborting on a failed unit switch, re-reads the MAPI key every run, validates element lists, and puts failures first with a log file when the report is long. LSS1_L/LSS2_L/LSA2_L now save to their own files (LSS2_L and LSA2_L used to share LS2_L.jpg)."

' Identifies this module to the updater regardless of what it was
' named when pasted into Excel - these files carry no VB_Name, so the
' module name in the VBA project is whatever the user typed.
Private Const SCRIPT_ID As String = "culvert-load-diagram-capture"

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
' <workbook folder>\mnf151\LOADS\. Created automatically if missing.
Private Const CAPTURE_SUBFOLDER As String = "LOADS"

' Sheet that holds the figures to replace. Leave empty to use whichever
' sheet is active when the macro runs.
Private Const TARGET_SHEET_NAME As String = "4_MODEL"

' One capture per row: CaseType|CaseName|OutputFileName|ShapeName
'   CaseType : "ST" (Static Load) is the only value the Display doc confirms.
'              Other values (CS/RS/TH/MV/SM/CB) are guesses from the Capture
'              endpoint's own LOAD_CASE_COMB enum - untested here. Leave both
'              CaseType and CaseName empty for a bare model/mesh capture with
'              no load overlay at all (nodes/elements only).
'   ShapeName: name of the existing Picture/Shape on TARGET_SHEET_NAME to
'              replace (Name Box, top-left, when the picture is selected -
'              or Home > Find & Select > Selection Pane).
' Rows separated by ";". EDIT with your real load case names.
Private Const LOAD_JOB_LIST As String = _
    "||MODEL|'model;" & _
    "ST|EV1|EV1|'ev1;" & _
    "ST|EV2|EV2|'ev2;" & _
    "ST|EHS1|EHS1|'ehs1;" & _
    "ST|EHA1|EHA1|'eha1;" & _
    "ST|EHS2_L|EHS2_L|'ehs2_l;" & _
    "ST|EHA2_L|EHA2_L|'eha2_l;" & _
    "ST|EHS2_R|EHS2_R|'ehs2_r;" & _
    "ST|EHA2_R|EHA2_R|'eha2_r;" & _
    "ST|LL1|LL1|'ll1;" & _
    "ST|LL|LL|'ll;" & _
    "ST|LLacc|LLacc|'llacc;" & _
    "ST|LSS1_L|LSS1_L|'lss1_l;" & _
    "ST|LSS2_L|LSS2_L|'lss2_l;" & _
    "ST|LSA2_L|LSA2_L|'lsa2_l;" & _
    "ST|EQ|EQ|'eq"
    
'    "ST|ATA|ATA|'ATA"
'    "ST|DL|DL|LOAD DL;" & _
'    "ST|EVin|EVin|LOAD EVin;" & _
'    "ST|LLin|LLin|LOAD LLin;" & _
'    "ST|WA|WA|LOAD WA;" & _

' Image size in pixels.
Private Const IMG_WIDTH As Long = 1269
Private Const IMG_HEIGHT As Long = 750

' Camera zoom on the whole model view.
'   25 <= value < 100 : zoom out
'   100                : zoom to fit (default)
'   100 < value <= 200 : zoom in
Private Const ZOOM_LEVEL As Long = 90

' Elements to activate before capture (only these are shown/plotted).
' Space- or comma-separated; supports MIDAS "AtoB" range shorthand.
' Leave empty to show the whole model (skips sending ACTIVE at all).
Private Const ELEMENT_LIST As String = ""

' Which load types to draw. Booleans per the Display doc's LOAD schema.
' Matched to the two example screenshots: a distributed load along a beam
' (BEAM_LOAD) and point loads at nodes along a column (NODAL_LOAD).
' Add more (e.g. PRESSURE_LOAD, FLOOR_LOAD) if your model uses them -
' full list is in the Display doc.
Private Const SHOW_NODAL_LOAD As Boolean = True
Private Const SHOW_BEAM_LOAD As Boolean = True
Private Const SHOW_NODAL_BODY_FORCE As Boolean = True   ' self-weight etc.
Private Const SHOW_PRESSURE_LOAD As Boolean = False
Private Const SHOW_SPECIFIED_DISPLACEMENT As Boolean = False

' Boundary display options (per view/DISPLAY manual: BOUNDARY object).
' Shows point spring supports (e.g. foundation soil springs) and general supports.
Private Const SHOW_POINT_SPRING_SUPPORT As Boolean = True
Private Const SHOW_SUPPORT As Boolean = True

' Load value label format/decimals, per the Display doc's LOAD_VALUE object.
'   FORMAT: "Default" | "Fixed" | "Scientific"
Private Const LOAD_VALUE_FORMAT As String = "Fixed"
Private Const LOAD_VALUE_PLACE As Long = 1

' Force/length units to switch to for these captures (db/UNIT, "Unit System"
' JSON Manual, ed. 2024.08.01: FORCE "N" "KN" "KGF" "TONF" "LBF" "KIPS",
' DIST "M" "CM" "MM" "FT" "IN"). The macro reads the model's current Unit
' System first, switches to these values, runs all captures, then always
' switches back to whatever it originally was - see CaptureLoadDiagrams.
Private Const CAPTURE_FORCE_UNIT As String = "KN"
Private Const CAPTURE_DIST_UNIT As String = "M"


' ---------------------------------------------------------------------------
'  One row of LOAD_JOB_LIST, parsed
' ---------------------------------------------------------------------------
Private Type LoadJob
    CaseType As String
    CaseName As String
    FileName As String
    ShapeName As String
End Type


' ===========================================================================
'  MAIN - one image per job in LOAD_JOB_LIST
' ===========================================================================
Sub CaptureLoadDiagrams()

    Dim jobs() As LoadJob
    Dim i As Integer
    Dim body As String
    Dim outFolder As String
    Dim exportPath As String
    Dim ws As Worksheet
    Dim origForce As String, origDist As String, origHeat As String, origTemper As String
    Dim unitsChanged As Boolean
    Dim unitWarn As String
    Dim prevCalc As XlCalculation
    Dim fastMode As Boolean
    Dim report As String
    Dim dispResult As String

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
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    fastMode = True

    jobs = ParseLoadJobList(LOAD_JOB_LIST)

    For i = LBound(jobs) To UBound(jobs)

        If Len(jobs(i).CaseName) > 0 Then
            label = jobs(i).CaseName
        Else
            label = jobs(i).FileName
        End If

        exportPath = outFolder & "'" & jobs(i).FileName & ".jpg"

        ' EQ exists only when the culvert builder's seismic gate is on.
        If Not CaseInModel(jobs(i).CaseType, jobs(i).CaseName) Then
            skipCount = skipCount + 1
            skipLog = skipLog & "  - " & label & vbCrLf
        Else
            ' Explicit view/DISPLAY call, in addition to the inline
            ' DISPLAY.BOUNDARY field in the capture body itself - the inline
            ' field alone was not enough live: boundary/spring symbols set by
            ' an earlier capture kept showing on later ones even when this
            ' job's own body asked for them off. Same belt-and-braces pattern
            ' as SendActiveAllRequest in the displacement-contour script (see
            ' CLAUDE.md). Not fatal if it fails - the inline field is still
            ' sent as a fallback, so a WARN is enough.
            dispResult = SendDisplayBoundaryRequest(Len(jobs(i).CaseType) = 0)
            If Len(dispResult) > 0 Then
                warnCount = warnCount + 1
                warnLog = warnLog & "  - " & label & ": view/DISPLAY boundary toggle failed: " & _
                          dispResult & vbCrLf
            End If

            body = BuildLoadCaptureBody(exportPath, jobs(i))
            Call LogOutcome(CaptureToShape(ws, exportPath, body, jobs(i).ShapeName), label, _
                            okCount, okLog, warnCount, warnLog, failCount, failLog)
        End If

    Next i

    On Error GoTo 0

    If fastMode Then Call EndFastMode(prevCalc)

    report = "Culvert load diagram capture  [" & SCRIPT_VERSION & "]" & vbCrLf & _
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
    If fastMode Then Call EndFastMode(prevCalc)
    If unitsChanged Then
        restoreMsg = RestoreUnitsAndReport(origForce, origDist, origHeat, origTemper)
    Else
        restoreMsg = "(Unit System was never changed.)"
    End If
    MsgBox "Unexpected error: " & failMsg & vbCrLf & restoreMsg, vbCritical

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


' Parses LOAD_JOB_LIST ("Type|Name|FileName|ShapeName;...") into an array.
Private Function ParseLoadJobList(ByVal listStr As String) As LoadJob()

    Dim rows() As String
    Dim fields() As String
    Dim result() As LoadJob
    Dim i As Integer

    rows = Split(listStr, ";")
    ReDim result(LBound(rows) To UBound(rows))

    For i = LBound(rows) To UBound(rows)
        fields = Split(rows(i), "|")

        ' A missing "|" used to surface as "Subscript out of range"
        ' on the field read below, naming neither the row nor the cause.
        If UBound(fields) < 3 Then
            Err.Raise vbObjectError + 513, "ParseLoadJobList", _
                "LOAD_JOB_LIST row " & (i + 1) & " has " & (UBound(fields) + 1) & _
                " field(s), expected 4. Check the ""|"" separators in:  " & rows(i)
        End If
        result(i).CaseType = Trim(fields(0))
        result(i).CaseName = Trim(fields(1))
        result(i).FileName = Trim(fields(2))
        result(i).ShapeName = Trim(fields(3))
    Next i

    ParseLoadJobList = result

End Function


' ===========================================================================
'  JSON BODY
'  Note: VBA caps a statement at 24 line continuations, which is why the
'  JSON is appended chunk by chunk instead of one long "& _" chain.
' ===========================================================================
Private Function BuildLoadCaptureBody(ByVal exportPath As String, _
                                      ByRef job As LoadJob) As String

    Dim b As String

    b = "{""Argument"": {"

    ' ---- capture settings ----
    ' SET_MODE "pre": loads exist on the model before/without analysis.
    b = b & """SET_MODE"": ""pre"","
    b = b & """SET_HIDDEN"": false,"
    b = b & """EXPORT_PATH"": """ & Replace(exportPath, "\", "\\") & ""","
    b = b & """WIDTH"": " & IMG_WIDTH & ","
    b = b & """HEIGHT"": " & IMG_HEIGHT & ","
    b = b & """ZOOM_LEVEL"": " & ZOOM_LEVEL & ","

    ' Plain white background (both gradient stops), per Capture doc specs
    ' #12/#13 (BGCOLOR_TOP / BGCOLOR_BOTTOM).
    b = b & """BGCOLOR_TOP"": {""R"": 255, ""G"": 255, ""B"": 255},"
    b = b & """BGCOLOR_BOTTOM"": {""R"": 255, ""G"": 255, ""B"": 255},"

    ' Activate only specific elements, if ELEMENT_LIST is set.
    If Len(ELEMENT_LIST) > 0 Then
        b = b & """ACTIVE"": {"
        b = b & """ACTIVE_MODE"": ""Active"","
        b = b & """E_LIST"": " & ElementListToJsonArray(ELEMENT_LIST)
        b = b & "},"
    End If

    ' ---- DISPLAY.NODE / DISPLAY.BOUNDARY / DISPLAY.LOAD, per the "Display" JSON Manual ----
    ' Node markers are always shown - not config-gated like the load types
    ' below, since there's no scenario here where hiding nodes helps.
    ' An empty CaseType (job row with no load case, e.g. "||MODEL|'model")
    ' skips the whole LOAD block, capturing the bare 2D model/mesh with
    ' point spring supports shown along the foundation.
    b = b & """DISPLAY"": {"
    b = b & """NODE"": {"
    b = b & """NODE"": true"
    b = b & "}"

    ' BOUNDARY is sent EXPLICITLY on every job, not just gated on the
    ' SHOW_* constants - on (per those constants) for the bare-model row,
    ' explicitly false for every load-case row. Confirmed live 2026-09-23:
    ' Civil NX CAN render boundary symbols and a load overlay together, so
    ' omitting the field on load-case rows is not enough to guarantee it
    ' stays off (whatever the last capture left active could otherwise
    ' carry over) - only the bare-model row is meant to show springs, per
    ' request.
    If Len(job.CaseType) = 0 Then
        If SHOW_POINT_SPRING_SUPPORT Or SHOW_SUPPORT Then
            b = b & ","
            b = b & """BOUNDARY"": {"
            b = b & """POINT_SPRING_SUPPORT"": " & LCase(SHOW_POINT_SPRING_SUPPORT) & ","
            b = b & """SUPPORT"": " & LCase(SHOW_SUPPORT)
            b = b & "}"
        End If
    Else
        b = b & ","
        b = b & """BOUNDARY"": {"
        b = b & """POINT_SPRING_SUPPORT"": false,"
        b = b & """SUPPORT"": false"
        b = b & "}"
    End If

    If Len(job.CaseType) > 0 Then
        b = b & ","
        b = b & """LOAD"": {"
        b = b & """CASE_SELECTION"": {"
        b = b & """TYPE"": """ & job.CaseType & ""","
        b = b & """NAME"": """ & job.CaseName & """"
        b = b & "},"
        b = b & """LOAD_VALUE"": {"
        b = b & """FORMAT"": """ & LOAD_VALUE_FORMAT & ""","
        b = b & """PLACE"": " & LOAD_VALUE_PLACE
        b = b & "},"
        b = b & """NODAL_LOAD"": " & LCase(SHOW_NODAL_LOAD) & ","
        b = b & """BEAM_LOAD"": " & LCase(SHOW_BEAM_LOAD) & ","
        b = b & """NODAL_BODY_FORCE"": " & LCase(SHOW_NODAL_BODY_FORCE) & ","
        b = b & """PRESSURE_LOAD"": " & LCase(SHOW_PRESSURE_LOAD) & ","
        b = b & """SPECIFIED_DISPLACEMENT"": " & LCase(SHOW_SPECIFIED_DISPLACEMENT)
        b = b & "}"
    End If

    b = b & "}"

    b = b & "}}"

    BuildLoadCaptureBody = b

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


' Explicitly sets the boundary/spring display toggle in Civil NX via
' view/DISPLAY (a separate POST endpoint from view/CAPTURE - see the
' "Display" JSON Manual, ed. 2024.10.25). Confirmed live 2026-09-23:
' the DISPLAY.BOUNDARY field embedded in the capture body alone did not
' reliably turn springs off on later captures once an earlier capture had
' turned them on, so this makes it an explicit standalone toggle instead
' of trusting the per-capture field to fully reset the session's display
' state. Called once per job, right before the capture request, so the
' bare-model row shows springs and every load-case row explicitly hides
' them - same belt-and-braces pattern as SendActiveAllRequest in
' midas-culvert-displacement-contour-capture.bas.
'
' Returns "" on success (2xx, no "error" key), else a short failure
' message - same convention as SendCaptureRequest/ShortApiError. Not
' fatal if it fails: the caller logs a WARN and the capture loop
' continues, since the inline DISPLAY.BOUNDARY field is still sent as a
' fallback.
Private Function SendDisplayBoundaryRequest(ByVal showBoundary As Boolean) As String

    Dim url As String
    Dim body As String
    Dim resp As String
    Dim statusCode As Long

    url = API_BASE_URL & "/view/DISPLAY"
    body = "{""Argument"": {""BOUNDARY"": {"
    body = body & """SUPPORT"": " & LCase(showBoundary) & ","
    body = body & """POINT_SPRING_SUPPORT"": " & LCase(showBoundary)
    body = body & "}}}"

    If HTTP_CLIENT Is Nothing Then
        Set HTTP_CLIENT = CreateObject("WinHttp.WinHttpRequest.5.1")
    End If

    On Error Resume Next

    HTTP_CLIENT.Open "POST", url, False
    HTTP_CLIENT.SetRequestHeader "MAPI-Key", MapiKey()
    HTTP_CLIENT.SetRequestHeader "Content-Type", "application/json"
    If Err.Number <> 0 Then
        SendDisplayBoundaryRequest = "WinHTTP error: " & Err.Description
        Err.Clear
        Set HTTP_CLIENT = Nothing
        On Error GoTo 0
        Exit Function
    End If

    HTTP_CLIENT.Send body
    If Err.Number <> 0 Then
        SendDisplayBoundaryRequest = "WinHTTP error: " & Err.Description
        Err.Clear
        ' Never reuse a client that just failed.
        Set HTTP_CLIENT = Nothing
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    statusCode = HTTP_CLIENT.Status
    resp = HTTP_CLIENT.ResponseText

    If statusCode >= 200 And statusCode < 300 And InStr(1, resp, """error""", vbTextCompare) = 0 Then
        SendDisplayBoundaryRequest = ""
    Else
        SendDisplayBoundaryRequest = "HTTP " & statusCode & HttpStatusHint(statusCode) & ShortApiError(resp)
    End If

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
