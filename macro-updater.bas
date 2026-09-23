Option Explicit

' ============================================================================
'  MIDAS macro updater - check this workbook's modules against a published
'  manifest, and replace the stale ones on confirmation.
' ----------------------------------------------------------------------------
'  Why this exists: these modules are pasted into Excel by hand, so the code
'  in a workbook and the code in the repo drift apart silently. That has
'  already cost a debugging round trip - a fix was reported as ineffective
'  twice because the module had never been re-imported. SCRIPT_VERSION was
'  added so a run at least SAYS which version produced it; this closes the
'  loop by fetching the current one.
'
'  HOW MODULES ARE IDENTIFIED
'  The .bas files carry no "Attribute VB_Name", so the module name inside a
'  workbook is whatever the person pasting it typed. Matching on that would
'  be unreliable, so every managed module declares its own SCRIPT_ID and
'  SCRIPT_VERSION, and this updater reads those out of the code itself.
'
'  WHAT THE SERVER MUST PUBLISH
'    <UPDATE_BASE_URL>/manifest.txt      one "id|version" per line, # = comment
'    <UPDATE_BASE_URL>/<id>.bas          the full module text for that id
'  e.g.
'    culvert-beam-diagram-capture|2026-09-22d
'    wingwall-model-build|2026-09-22-live
'
'  REQUIREMENT - THIS ONE CATCHES PEOPLE OUT
'  Replacing code needs "Trust access to the VBA project object model":
'    File > Options > Trust Center > Trust Center Settings >
'    Macro Settings > tick "Trust access to the VBA project object model"
'  It is OFF by default and is per-machine, not per-workbook. Without it,
'  CheckMidasMacroUpdates still reports what is stale - it just cannot
'  install anything.
'
'  SAFETY
'   * nothing is replaced without an explicit Yes
'   * each module is exported to a backup folder beside the workbook first
'   * this module never replaces itself (a module cannot rewrite its own
'     code while that code is running)
'   * a download with no SCRIPT_ID is refused, so a 404 body or a proxy
'     login page can never be written into the VBA project
'
'  THE MAPI KEY IS NOT IN THESE MODULES AT ALL
'  Since 2026-09-23 every module reads it from INPUT!J20 at runtime, so an
'  update cannot cost the user their key and the published copies carry no
'  secret whatsoever. Rotating a key is a one-cell edit, not a re-paste of
'  nine modules.
'
'  WHAT AN UPDATE DOES NOT PRESERVE
'  Any local edit to a managed module - a tweaked JOB_LIST, a changed
'  sheet name, an adjusted zoom - is replaced by the published version.
'  The backup in BACKUP_SUBFOLDER is what you diff against to get it back.
'  Keep per-workbook customisation on the worksheet, not in the module.
' ============================================================================


' ---------------------------------------------------------------------------
'  CONFIG
' ---------------------------------------------------------------------------

Private Const SCRIPT_VERSION As String = "2026-09-23b"

' Identifies this module to the updater regardless of what it was
' named when pasted into Excel - these files carry no VB_Name, so the
' module name in the VBA project is whatever the user typed.
Private Const SCRIPT_ID As String = "macro-updater"

' Base URL serving manifest.txt and <id>.bas. No trailing slash.
' Points at a SEPARATE PUBLIC releases repo, mirroring how
' RC-Retaining_Wall publishes its desktop builds: the source repo stays
' private while the feed needs no token. Publish into it with
'   python scripts/make_manifest.py --publish <clone of that repo>
' Any plain HTTP(S) host would work just as well.
Private Const UPDATE_BASE_URL As String = _
    "https://raw.githubusercontent.com/Thorindiel-Inc/Culvert_Releases/main"

' Subfolder beside the workbook where the pre-update copies go.
Private Const BACKUP_SUBFOLDER As String = "macro-backups"

' Seconds to wait on each HTTP request.
Private Const REQUEST_TIMEOUT_SECONDS As Long = 20



' ===========================================================================
'  MAIN - the only entry point you need
' ===========================================================================
Public Sub CheckMidasMacroUpdates()

    Dim manifest As String
    Dim comp As Object
    Dim code As String
    Dim thisId As String, thisVer As String
    Dim latest As String
    Dim okLog As String, staleLog As String, skipLog As String
    Dim okCount As Long, staleCount As Long
    Dim staleNames As String
    Dim trustMsg As String
    Dim answer As VbMsgBoxResult
    Dim report As String

    If Not FetchText(UPDATE_BASE_URL & "/manifest.txt", manifest) Then
        MsgBox "Could not reach the update server." & vbCrLf & vbCrLf & _
               manifest & vbCrLf & vbCrLf & _
               "URL: " & UPDATE_BASE_URL & "/manifest.txt", vbCritical
        Exit Sub
    End If

    trustMsg = VbaTrustProblem()

    ' Enumerating VBComponents is what actually trips the Trust Center
    ' setting - VbaTrustProblem() already proved that above via its own
    ' On Error Resume Next. Doing it again here unguarded would crash with
    ' the same 1004 before the trust message ever gets shown, so skip the
    ' loop entirely when the project is not reachable.
    If Len(trustMsg) = 0 Then
        For Each comp In ThisWorkbook.VBProject.VBComponents
            code = ComponentCode(comp)
            thisId = ConstValue(code, "SCRIPT_ID")
            If Len(thisId) > 0 And StrComp(thisId, SCRIPT_ID, vbTextCompare) <> 0 Then
                thisVer = ConstValue(code, "SCRIPT_VERSION")
                latest = ManifestVersion(manifest, thisId)
                If Len(latest) = 0 Then
                    skipLog = skipLog & "  - " & comp.Name & " (" & thisId & _
                              "): not on the server" & vbCrLf
                ElseIf StrComp(thisVer, latest, vbTextCompare) = 0 Then
                    okCount = okCount + 1
                    okLog = okLog & "  - " & comp.Name & "  " & thisVer & vbCrLf
                Else
                    staleCount = staleCount + 1
                    staleLog = staleLog & "  - " & comp.Name & "  " & thisVer & _
                               "  ->  " & latest & vbCrLf
                    staleNames = staleNames & comp.Name & "|" & thisId & ";"
                End If
            End If
        Next comp
    End If

    report = "MIDAS macro update check  [" & SCRIPT_VERSION & "]" & vbCrLf & _
             String(46, "-") & vbCrLf

    If okCount > 0 Then report = report & "UP TO DATE (" & okCount & "):" & vbCrLf & okLog
    If staleCount > 0 Then report = report & vbCrLf & "OUT OF DATE (" & staleCount & "):" & vbCrLf & staleLog
    If Len(skipLog) > 0 Then report = report & vbCrLf & "NOT MANAGED:" & vbCrLf & skipLog

    If Len(trustMsg) > 0 Then
        MsgBox report & vbCrLf & "CANNOT CHECK:" & vbCrLf & trustMsg, vbExclamation
        Exit Sub
    End If

    If staleCount = 0 Then
        If okCount = 0 Then
            MsgBox report & vbCrLf & "No MIDAS modules found in this workbook.", vbInformation
        Else
            MsgBox report & vbCrLf & "OK - everything is current.", vbInformation
        End If
        Exit Sub
    End If

    answer = MsgBox(report & vbCrLf & _
             "Replace the " & staleCount & " out-of-date module(s)?" & vbCrLf & vbCrLf & _
             "A copy of each is exported to '" & BACKUP_SUBFOLDER & _
             "' beside this workbook first." & vbCrLf & vbCrLf & _
             "Save your work before continuing.", _
             vbYesNo + vbQuestion, "MIDAS macro updater")

    If answer <> vbYes Then Exit Sub

    Call InstallUpdates(staleNames)

End Sub


' ===========================================================================
'  INSTALL
' ===========================================================================

' staleList is "moduleName|scriptId;" repeated. Each module is exported as a
' backup, then its code is replaced IN PLACE - the component is never removed
' and re-imported, so the module keeps its name and any references to it.
Private Sub InstallUpdates(ByVal staleList As String)

    Dim rows() As String
    Dim f() As String
    Dim i As Long
    Dim comp As Object
    Dim newCode As String
    Dim backupDir As String
    Dim doneLog As String, failLog As String
    Dim doneCount As Long, failCount As Long

    backupDir = ThisWorkbook.Path & Application.PathSeparator & BACKUP_SUBFOLDER
    If Not EnsureFolderExists(backupDir) Then
        MsgBox "Could not create the backup folder:" & vbCrLf & backupDir & vbCrLf & _
               vbCrLf & "Nothing was changed.", vbCritical
        Exit Sub
    End If

    rows = Split(staleList, ";")

    For i = LBound(rows) To UBound(rows)
        If Len(Trim$(rows(i))) > 0 Then
            f = Split(rows(i), "|")

            If Not FetchText(UPDATE_BASE_URL & "/" & f(1) & ".bas", newCode) Then
                failCount = failCount + 1
                failLog = failLog & "  - " & f(0) & ": download failed - " & newCode & vbCrLf
            ElseIf Len(ConstValue(newCode, "SCRIPT_ID")) = 0 Then
                ' Refuse to install something that is not one of these modules -
                ' a proxy login page or a 404 body would otherwise be written
                ' straight into the VBA project.
                failCount = failCount + 1
                failLog = failLog & "  - " & f(0) & ": downloaded text has no " & _
                          "SCRIPT_ID, refusing to install it" & vbCrLf
            Else
                Set comp = Nothing
                On Error Resume Next
                Set comp = ThisWorkbook.VBProject.VBComponents(f(0))
                On Error GoTo 0

                If comp Is Nothing Then
                    failCount = failCount + 1
                    failLog = failLog & "  - " & f(0) & ": module vanished" & vbCrLf
                ElseIf Not BackupComponent(comp, backupDir) Then
                    failCount = failCount + 1
                    failLog = failLog & "  - " & f(0) & ": backup failed, left alone" & vbCrLf
                ElseIf ReplaceComponentCode(comp, newCode) Then
                    doneCount = doneCount + 1
                    doneLog = doneLog & "  - " & f(0) & "  ->  " & _
                              ConstValue(newCode, "SCRIPT_VERSION") & vbCrLf
                Else
                    failCount = failCount + 1
                    failLog = failLog & "  - " & f(0) & ": replace failed - restore " & _
                              "from " & BACKUP_SUBFOLDER & vbCrLf
                End If
            End If
        End If
    Next i

    MsgBox "MIDAS macro update" & vbCrLf & String(46, "-") & vbCrLf & _
           IIf(doneCount > 0, "UPDATED (" & doneCount & "):" & vbCrLf & doneLog, "") & _
           IIf(failCount > 0, vbCrLf & "FAILED (" & failCount & "):" & vbCrLf & failLog, "") & _
           vbCrLf & "Backups: " & backupDir & vbCrLf & vbCrLf & _
           "Save the workbook to keep the updated code.", _
           IIf(failCount > 0, vbExclamation, vbInformation)

End Sub


Private Function BackupComponent(ByVal comp As Object, ByVal folder As String) As Boolean

    Dim path As String

    path = folder & Application.PathSeparator & comp.Name & "_" & _
           Format$(Now, "yyyymmdd_hhnnss") & ".bas"

    On Error Resume Next
    comp.Export path
    BackupComponent = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

End Function


Private Function ReplaceComponentCode(ByVal comp As Object, ByVal newCode As String) As Boolean

    On Error Resume Next
    With comp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString newCode
    End With
    ReplaceComponentCode = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

End Function


' ===========================================================================
'  HELPERS
' ===========================================================================

' "" when the VBA project object model is reachable, otherwise the message
' to show the user. Reading .VBComponents is what actually trips the Trust
' Center setting, so that is what gets tested.
Private Function VbaTrustProblem() As String

    Dim n As Long

    On Error Resume Next
    n = ThisWorkbook.VBProject.VBComponents.count
    If Err.Number <> 0 Then
        Err.Clear
        VbaTrustProblem = "Excel is blocking access to the VBA project." & vbCrLf & _
            "Tick: File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
            "      Macro Settings > Trust access to the VBA project object model" & vbCrLf & _
            "then reopen the workbook. (This is a per-machine Excel setting.)"
    End If
    On Error GoTo 0

End Function


Private Function ComponentCode(ByVal comp As Object) As String

    On Error Resume Next
    With comp.CodeModule
        If .CountOfLines > 0 Then ComponentCode = .Lines(1, .CountOfLines)
    End With
    Err.Clear
    On Error GoTo 0

End Function


' Pulls the literal out of:  Private Const <name> As String = "value"
Private Function ConstValue(ByVal code As String, ByVal constName As String) As String

    Dim posConst As Long, q1 As Long, q2 As Long

    posConst = InStr(1, code, "Const " & constName & " As String", vbTextCompare)
    If posConst = 0 Then Exit Function

    q1 = InStr(posConst, code, """")
    If q1 = 0 Then Exit Function

    q2 = InStr(q1 + 1, code, """")
    If q2 = 0 Then Exit Function

    ConstValue = Mid$(code, q1 + 1, q2 - q1 - 1)

End Function


' Looks up one id in the manifest. Lines are "id|version"; blank lines and
' lines starting with "#" are ignored.
Private Function ManifestVersion(ByVal manifest As String, ByVal wantId As String) As String

    Dim rows() As String
    Dim f() As String
    Dim i As Long
    Dim ln As String

    manifest = Replace(manifest, vbCrLf, vbLf)
    manifest = Replace(manifest, vbCr, vbLf)
    rows = Split(manifest, vbLf)

    For i = LBound(rows) To UBound(rows)
        ln = Trim$(rows(i))
        If Len(ln) > 0 And Left$(ln, 1) <> "#" Then
            f = Split(ln, "|")
            If UBound(f) >= 1 Then
                If StrComp(Trim$(f(0)), wantId, vbTextCompare) = 0 Then
                    ManifestVersion = Trim$(f(1))
                    Exit Function
                End If
            End If
        End If
    Next i

End Function


' True on success with the body in outText; False with the reason in outText.
Private Function FetchText(ByVal url As String, ByRef outText As String) As Boolean

    Dim http As Object
    Dim statusCode As Long

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error Resume Next

    http.SetTimeouts REQUEST_TIMEOUT_SECONDS * 1000, REQUEST_TIMEOUT_SECONDS * 1000, _
                     REQUEST_TIMEOUT_SECONDS * 1000, REQUEST_TIMEOUT_SECONDS * 1000
    http.Open "GET", url, False
    http.Send
    If Err.Number <> 0 Then
        outText = "WinHTTP error: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Set http = Nothing
        Exit Function
    End If

    On Error GoTo 0

    statusCode = http.Status
    outText = http.ResponseText
    Set http = Nothing

    If statusCode < 200 Or statusCode > 299 Then
        outText = "HTTP " & statusCode & HttpStatusHint(statusCode)
        Exit Function
    End If

    FetchText = True

End Function


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
