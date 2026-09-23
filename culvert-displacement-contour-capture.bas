Option Explicit

' ============================================================================
'  MIDAS Civil NX API - Displacement Contour Screenshot Capture
' ----------------------------------------------------------------------------
'  Endpoint : {base url} + view/CAPTURE
'  Docs     : "Capture"                          (JSON Manual, ed. 2025.04.24)
'             "Displacement Contour - Result Display" (JSON Manual, ed. 2025.01.03)
'             - Argument.RESULT_GRAPHIC follows that same view/RESULTGRAPHIC
'             schema, with CURRENT_MODE = "DisplacementContour".
'             "Type of Display"                (JSON Manual, ed. 2026.06.19)
'
'  Captures one displacement-contour image per job in JOB_LIST (envelope
'  combination + component + target shape), replacing the matching
'  Picture/Shape on TARGET_SHEET_NAME with the freshly captured image.
'  Same pattern as midas-beam-diagram-capture.bas.
' ============================================================================


' ---------------------------------------------------------------------------
'  CONFIG
' ---------------------------------------------------------------------------

' Printed in the summary MsgBox. These modules are pasted into Excel by
' hand, so the file in the repo and the code actually running can silently
' diverge - check this stamp matches the constant here before concluding
' anything from a run. Bump it whenever this file changes.
Private Const SCRIPT_VERSION As String = "2026-09-23f"

' Identifies this module to the updater regardless of what it was
' named when pasted into Excel - these files carry no VB_Name, so the
' module name in the VBA project is whatever the user typed.
Private Const SCRIPT_ID As String = "culvert-displacement-contour-capture"

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
' <workbook folder>\mnf151\DISPLACEMENTS\. Created automatically if missing.
Private Const CAPTURE_SUBFOLDER As String = "DISPLACEMENTS"

' Sheet that holds the figures to replace. Leave empty to use whichever
' sheet is active when the macro runs.
Private Const TARGET_SHEET_NAME As String = "7_ZG"

' One capture per row: CombName|CombType|MinMax|Component|ShapeName
'   CombType  : "ST" "CS" "RS" "TH" "MV" "SM" "CB" - envelope combinations
'               (ENV_...) are "CB".
'   MinMax    : "Max" "Min" "All" (All = envelope abs max, per the doc).
'   Component : "DX" "DY" "DZ" "RX" "RY" "RZ" "RW" "DXY" "DYZ" "DXZ" "DXYZ"
'   ShapeName : name of the existing Picture/Shape on TARGET_SHEET_NAME to
'               replace (Name Box, top-left, when the picture is selected -
'               or Home > Find & Select > Selection Pane).
' Rows separated by ";". EDIT to add/remove captures.
Private Const JOB_LIST As String = _
    "ENV_SER|CB|Min|DZ|env_ser;" & _
    "ENV_STR|CB|Min|DZ|env_str;" & _
    "ENV_EQ|CB|Min|DZ|env_eq;" & _
    "ENV_DEAD|CB|Min|DZ|env_dead;" & _
    "LL|ST|Min|DZ|ll"

' Image size in pixels, matched to the Excel figure boxes' physical size
' (33.16 x 19.61 cm) at 120 DPI, so the captured image isn't stretched/
' squished when ReplacePictureByName fits it into that same box.
Private Const IMG_WIDTH As Long = 1269
Private Const IMG_HEIGHT As Long = 750

' Camera zoom on the whole model view.
'   25 <= value < 100 : zoom out
'   100                : zoom to fit (default)
'   100 < value <= 200 : zoom in
Private Const ZOOM_LEVEL As Long = 100

' Elements to activate before capture.
' Leave empty to show all elements (sends ACTIVE_MODE "All" so the full model is contoured).
' Or set to a space- or comma-separated list / range (e.g. "4 10to13 17to19 21to44").
Private Const ELEMENT_LIST As String = ""

' Number of decimal places for the printed displacement values/legend.
Private Const VALUE_DECIMAL_PT As Long = 2

' Rotation of the printed value-label text itself (TYPE_OF_DISPLAY.VALUES.
' SET_ORIENT), in degrees, 0 to 180 in increments of 15 per the doc. This
' is text orientation only - it does not move the camera/viewpoint.
Private Const VALUE_TEXT_ORIENT As Long = 15

' Deformed-shape scale factor (TYPE_OF_DISPLAY.DEFORM.SCALE_FACTOR).
Private Const DEFORM_SCALE_FACTOR As Double = 1#

' Length unit to switch to for these captures (db/UNIT, "Unit System" JSON
' Manual, ed. 2024.08.01: "M" "CM" "MM" "FT" "IN"). The macro reads the
' model's current Unit System first, switches DIST to this value, runs all
' captures, then always switches DIST back to whatever it originally was -
' see CaptureDisplacementContours.
Private Const CAPTURE_DIST_UNIT As String = "MM"


' ---------------------------------------------------------------------------
'  One row of JOB_LIST, parsed
' ---------------------------------------------------------------------------
Private Type CaptureJob
    CombName As String
    CombType As String
    MinMax As String
    CompName As String
    ShapeName As String
End Type


' ===========================================================================
'  MAIN - one image per job in JOB_LIST
' ===========================================================================
Sub CaptureDisplacementContours()

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
    ' then switch length to CAPTURE_DIST_UNIT for these captures.
    If Not GetUnitSystem(origForce, origDist, origHeat, origTemper) Then
        MsgBox "Could not read the current Unit System (db/UNIT) - aborting before changing anything.", vbCritical
        Exit Sub
    End If

    If Not SetUnitSystem(origForce, CAPTURE_DIST_UNIT, origHeat, origTemper) Then
        MsgBox "Could not switch the Unit System to " & CAPTURE_DIST_UNIT & " - aborting.", vbCritical
        Exit Sub
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

    jobs = ParseJobList(JOB_LIST)

    ' Ensure all elements are active (Active All) in Civil NX before
    ' capturing. Not fatal if it fails - the ACTIVE_MODE "All" sent with
    ' each individual capture below is the same (unconfirmed) value, so a
    ' failure here just means we proceed without the extra safety net.
    Dim activeAllResult As String
    activeAllResult = SendActiveAllRequest()
    If Len(activeAllResult) > 0 Then
        warnCount = warnCount + 1
        warnLog = warnLog & "  - Active All (view/ACTIVE): " & activeAllResult & vbCrLf
    End If

    For i = LBound(jobs) To UBound(jobs)

        exportPath = outFolder & jobs(i).CombName & ".jpg"
        label = jobs(i).CombName & " (" & jobs(i).CompName & ")"

        body = BuildCaptureBody(exportPath, jobs(i))

        Call SendCaptureRequest(body, resp, statusCode)

        If IsApiSuccess(statusCode, resp) Then

            If WaitForFile(exportPath, 8) Then

                If ReplacePictureByName(ws, jobs(i).ShapeName, exportPath) Then
                    okCount = okCount + 1
                    okLog = okLog & "  - " & label & vbCrLf
                Else
                    warnCount = warnCount + 1
                    warnLog = warnLog & "  - " & label & ": captured, but no shape named """ & _
                               jobs(i).ShapeName & """ found on """ & ws.Name & """" & vbCrLf
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

    MsgBox "Culvert displacement contour capture  [" & SCRIPT_VERSION & "]" & vbCrLf & vbCrLf & _
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
' by outcome (OK / WARNINGS / FAILED) instead of interleaved in job order -
' scans far faster than a flat chronological log, especially once there
' are failures mixed in with successes.
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

' Pulls a short human-readable reason out of a MIDAS API JSON response,
' e.g. " - MIDAS CIVIL NX second query is wrong", instead of dumping the
' raw JSON body into the report. Returns "" if there's nothing usable.
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

    If putOk And readOk And StrComp(checkDist, dist, vbTextCompare) = 0 Then
        RestoreUnitsAndReport = "Unit System restored to " & dist & "."
    Else
        RestoreUnitsAndReport = "WARNING: Unit System restore may have failed - tried to set DIST=""" & _
            dist & """, but a follow-up check now reads DIST=""" & checkDist & """. " & _
            "Set it back manually in Civil NX (Tools > Unit System) if needed."
    End If

End Function


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
        result(i).ShapeName = Trim(fields(4))
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
    ' SET_MODE "post": displacement contours are analysis results.
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

    ' Always ensure all elements are active (Active All) so contours cover the full model.
    b = b & """ACTIVE"": {"
    If Len(ELEMENT_LIST) > 0 Then
        b = b & """ACTIVE_MODE"": ""Active"","
        b = b & """E_LIST"": " & ElementListToJsonArray(ELEMENT_LIST)
    Else
        b = b & """ACTIVE_MODE"": ""All"""
    End If
    b = b & "},"

    ' ---- result display ----
    b = b & """RESULT_GRAPHIC"": {"
    b = b & """CURRENT_MODE"": ""DisplacementContour"","

    ' load case / combination
    b = b & """LOAD_CASE_COMB"": {"
    b = b & """TYPE"": """ & job.CombType & ""","
    b = b & """MINMAX"": """ & job.MinMax & ""","
    b = b & """NAME"": """ & job.CombName & """"
    b = b & "},"

    ' component
    b = b & """COMPONENTS"": {"
    b = b & """COMP"": """ & job.CompName & ""","
    b = b & """OPT_LOCAL_CHECK"": false"
    b = b & "},"

    ' ---- contour / deform / values / legend ----
    ' NUM_OF_COLOR must be 6 / 12 / 18 / 24.
    ' VALUE_EXP defaults to true, which is why untouched legends print
    ' in exponential notation. DEFORM shows the deformed shape (real,
    ' unscaled per REAL_DISP); UNDEFORMED additionally overlays the
    ' original outline for reference.
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
    b = b & """DEFORM"": {"
    b = b & """OPT_CHECK"": true,"
    b = b & """SCALE_FACTOR"": " & JsonNum(DEFORM_SCALE_FACTOR) & ","
    b = b & """REAL_DEFORM"": false,"
    b = b & """REL_DISP"": false,"
    b = b & """REAL_DISP"": false"
    b = b & "},"
    b = b & """VALUES"": {"
    b = b & """OPT_CHECK"": true,"
    b = b & """VALUE_EXP"": false,"
    b = b & """DECIMAL_PT"": " & VALUE_DECIMAL_PT & ","
    b = b & """SET_ORIENT"": " & VALUE_TEXT_ORIENT
    b = b & "},"
    b = b & """LEGEND"": {"
    b = b & """OPT_CHECK"": false,"
    b = b & """POSITION"": ""right"","
    b = b & """VALUE_EXP"": false,"
    b = b & """DECIMAL_PT"": " & VALUE_DECIMAL_PT
    b = b & "},"
    ' UNDEFORMED: bare boolean form - confirmed live 2026-09-21. The
    ' object form {"OPT_CHECK": true} was tried first and failed here
    ' ("MIDAS CIVIL NX second query is wrong", 200 OK, no file written);
    ' CLAUDE.md previously claimed the object form worked on this file,
    ' which was wrong and has been corrected. Don't reintroduce the
    ' object form without re-testing live.
    b = b & """UNDEFORMED"": true"

    b = b & "}"    ' closes TYPE_OF_DISPLAY
    b = b & "}"    ' closes RESULT_GRAPHIC
    b = b & "}}"   ' closes Argument, then the outer object

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


' Locale-invariant number -> JSON literal. Plain "&" is locale-aware (a
' Turkish Windows renders 0.4 as "0,4"), which corrupts JSON silently. Str$
' is always period-decimal but drops the leading zero for |v|<1 (".001"),
' also invalid JSON - re-added here.
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


' Explicitly sets Civil NX model view to Active All via view/ACTIVE.
'
' UNCONFIRMED FIELD: "ACTIVE_MODE": "All" is a guess by analogy, not a
' documented value - the Capture doc's own ACTIVE field just says "refer
' to the view/ACTIVE manual" for its schema, and that manual page was
' never fetched (same gap noted in CLAUDE.md). "Active" + N_LIST/E_LIST
' is the only value confirmed from a worked example. If this silently
' does nothing, captures still proceed using whatever was already active
' in Civil NX - which is exactly the state this call exists to avoid, so
' the caller surfaces a WARN rather than treating a bad response here as
' fatal to the whole run.
'
' Returns "" on success (2xx, no "error" key), else a short failure
' message - same convention as SendCaptureRequest/ShortApiError.
Private Function SendActiveAllRequest() As String

    Dim url As String
    Dim body As String
    Dim resp As String
    Dim statusCode As Long

    url = API_BASE_URL & "/view/ACTIVE"
    body = "{""Argument"": {""ACTIVE_MODE"": ""All""}}"

    If HTTP_CLIENT Is Nothing Then
        Set HTTP_CLIENT = CreateObject("WinHttp.WinHttpRequest.5.1")
    End If

    On Error Resume Next

    HTTP_CLIENT.Open "POST", url, False
    HTTP_CLIENT.SetRequestHeader "MAPI-Key", MapiKey()
    HTTP_CLIENT.SetRequestHeader "Content-Type", "application/json"
    If Err.Number <> 0 Then
        SendActiveAllRequest = "WinHTTP error: " & Err.Description
        Err.Clear
        Set HTTP_CLIENT = Nothing
        On Error GoTo 0
        Exit Function
    End If

    HTTP_CLIENT.Send body
    If Err.Number <> 0 Then
        SendActiveAllRequest = "WinHTTP error: " & Err.Description
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
        SendActiveAllRequest = ""
    Else
        SendActiveAllRequest = "HTTP " & statusCode & HttpStatusHint(statusCode) & ShortApiError(resp)
    End If

End Function
