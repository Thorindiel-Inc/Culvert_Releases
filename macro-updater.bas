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
'  be unreliable, so every managed module declares its own SCRIPT_ID,
'  SCRIPT_VERSION and SCRIPT_CHANGELOG, and this updater reads those out of
'  the code itself.
'
'  WHAT THE SERVER MUST PUBLISH
'    <UPDATE_BASE_URL>/manifest.txt      one "id|version|changelog" per
'                                         line, # = comment. changelog is
'                                         optional - a 2-field "id|version"
'                                         line still works, just with
'                                         nothing to show.
'    <UPDATE_BASE_URL>/<id>.bas          the full module text for that id
'  e.g.
'    culvert-beam-diagram-capture|2026-09-22d|Shares the unified JSON helpers.
'    wingwall-model-build|2026-09-22-live|PostPlaneLoadTypes now actually sends its PUT.
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
'   * a download with no SCRIPT_ID is refused, so a 404 body or a proxy
'     login page can never be written into the VBA project
'
'  HOW THIS MODULE UPDATES ITSELF
'  Rewriting a module while one of its own procedures is on the call stack
'  can crash Excel or corrupt the VBA project, so this module is never
'  touched by InstallUpdates. Instead, when it is itself stale:
'    1. every other stale module is replaced first, as normal
'    2. its new code is downloaded to a temp file and the workbook is SAVED
'       - so if anything below goes wrong, nothing else is lost
'    3. a throwaway module (BOOTSTRAP_MODULE) is added and scheduled with
'       Application.OnTime, and this macro returns
'    4. the bootstrap runs from Excel's idle loop, with no code from this
'       module running, and rewrites this module from the temp file
'    5. it then schedules FinishMidasSelfUpdate - which is the NEW code -
'       to delete the bootstrap and report
'  A bootstrap left behind by a crash is removed on the next check.
'
'  NO LOCAL BACKUP IS TAKEN BEFORE REPLACING. The published .bas files are
'  themselves the record of what a module looked like at any point - this
'  repo's git history is the backup, not a per-workbook copy. If a
'  workbook-specific edit needs recovering, it has to come from wherever
'  that edit was made, not from this updater.
'
'  THE MAPI KEY IS NOT IN THESE MODULES AT ALL
'  Since 2026-09-23 every module reads it from INPUT!J20 at runtime, so an
'  update cannot cost the user their key and the published copies carry no
'  secret whatsoever. Rotating a key is a one-cell edit, not a re-paste of
'  nine modules.
'
'  WHAT AN UPDATE DOES NOT PRESERVE
'  Any local edit to a managed module - a tweaked JOB_LIST, a changed
'  sheet name, an adjusted zoom - is replaced by the published version,
'  with NO local backup taken first (see above). Keep per-workbook
'  customisation on the worksheet, not in the module.
' ============================================================================


' ---------------------------------------------------------------------------
'  CONFIG
' ---------------------------------------------------------------------------

Private Const SCRIPT_VERSION As String = "2026-09-23e"

' One-line summary of what changed in THIS version, shown when the updater
' finds itself stale. One physical line, no "|".
Private Const SCRIPT_CHANGELOG As String = "The updater can now update itself - replaced last, after the workbook is saved, via a temporary bootstrap module run from Application.OnTime."

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

' Seconds to wait on each HTTP request.
Private Const REQUEST_TIMEOUT_SECONDS As Long = 20

' The throwaway module that performs a self-update - see HOW THIS MODULE
' UPDATES ITSELF above. Removed again by FinishMidasSelfUpdate, or by the
' next check if a crash left it behind.
Private Const BOOTSTRAP_MODULE As String = "MidasUpdaterBootstrap"

' vbext_ComponentType.vbext_ct_StdModule, spelled out so no reference to
' the VBIDE library is needed.
Private Const STD_MODULE As Long = 1



' ===========================================================================
'  MAIN - the only entry point you need
' ===========================================================================
Public Sub CheckMidasMacroUpdates()

    Dim manifest As String
    Dim comp As Object
    Dim code As String
    Dim thisId As String, thisVer As String
    Dim latest As String, latestChangelog As String
    Dim okLog As String, staleLog As String, skipLog As String
    Dim okCount As Long, staleCount As Long
    Dim staleNames As String
    Dim selfStale As Boolean, selfCompName As String, selfLog As String
    Dim trustMsg As String
    Dim answer As VbMsgBoxResult
    Dim report As String
    Dim prompt As String

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
        ' A bootstrap left behind by an interrupted self-update.
        Call RemoveBootstrapModule

        For Each comp In ThisWorkbook.VBProject.VBComponents
            code = ComponentCode(comp)
            thisId = ConstValue(code, "SCRIPT_ID")
            If StrComp(thisId, SCRIPT_ID, vbTextCompare) = 0 Then
                ' This module. Never handed to InstallUpdates - it is
                ' replaced last, by BeginSelfUpdate. Uses the constant, not
                ' the parsed code, since this is the code actually running.
                latest = ManifestVersion(manifest, thisId)
                If Len(latest) > 0 And StrComp(SCRIPT_VERSION, latest, vbTextCompare) <> 0 Then
                    selfStale = True
                    selfCompName = comp.Name
                    selfLog = "  - " & comp.Name & "  " & SCRIPT_VERSION & _
                              "  ->  " & latest & vbCrLf
                    latestChangelog = ManifestChangelog(manifest, thisId)
                    If Len(latestChangelog) > 0 Then
                        selfLog = selfLog & "      " & latestChangelog & vbCrLf
                    End If
                End If
            ElseIf Len(thisId) > 0 Then
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
                    latestChangelog = ManifestChangelog(manifest, thisId)
                    staleLog = staleLog & "  - " & comp.Name & "  " & thisVer & _
                               "  ->  " & latest & vbCrLf
                    If Len(latestChangelog) > 0 Then
                        staleLog = staleLog & "      " & latestChangelog & vbCrLf
                    End If
                    staleNames = staleNames & comp.Name & "|" & thisId & ";"
                End If
            End If
        Next comp
    End If

    report = "MIDAS macro update check  [" & SCRIPT_VERSION & "]" & vbCrLf & _
             String(46, "-") & vbCrLf

    If okCount > 0 Then report = report & "UP TO DATE (" & okCount & "):" & vbCrLf & okLog
    If staleCount > 0 Then report = report & vbCrLf & "OUT OF DATE (" & staleCount & "):" & vbCrLf & staleLog
    If selfStale Then report = report & vbCrLf & "UPDATER ITSELF (replaced last):" & vbCrLf & selfLog
    If Len(skipLog) > 0 Then report = report & vbCrLf & "NOT MANAGED:" & vbCrLf & skipLog

    If Len(trustMsg) > 0 Then
        MsgBox report & vbCrLf & "CANNOT CHECK:" & vbCrLf & trustMsg, vbExclamation
        Exit Sub
    End If

    If staleCount = 0 And Not selfStale Then
        If okCount = 0 Then
            MsgBox report & vbCrLf & "No MIDAS modules found in this workbook.", vbInformation
        Else
            MsgBox report & vbCrLf & "OK - everything is current.", vbInformation
        End If
        Exit Sub
    End If

    prompt = report & vbCrLf & "Replace the " & _
             (staleCount + IIf(selfStale, 1, 0)) & " out-of-date module(s)?" & vbCrLf & vbCrLf & _
             "No local backup is taken - any per-workbook edit to a managed " & _
             "module is lost." & vbCrLf & vbCrLf
    If selfStale Then
        prompt = prompt & "The updater itself goes LAST: the workbook is SAVED " & _
                 "automatically first, then it rewrites itself once this " & _
                 "macro has finished." & vbCrLf & vbCrLf
    End If
    prompt = prompt & "Save your work before continuing."

    answer = MsgBox(prompt, vbYesNo + vbQuestion, "MIDAS macro updater")

    If answer <> vbYes Then Exit Sub

    If staleCount > 0 Then Call InstallUpdates(staleNames)
    If selfStale Then Call BeginSelfUpdate(selfCompName)

End Sub


' ===========================================================================
'  INSTALL
' ===========================================================================

' staleList is "moduleName|scriptId;" repeated. Each module's code is
' replaced IN PLACE, with no backup taken first (see the header comment's
' NO LOCAL BACKUP note) - the component is never removed and re-imported,
' so the module keeps its name and any references to it.
Private Sub InstallUpdates(ByVal staleList As String)

    Dim rows() As String
    Dim f() As String
    Dim i As Long
    Dim comp As Object
    Dim newCode As String
    Dim installedChangelog As String
    Dim doneLog As String, failLog As String
    Dim doneCount As Long, failCount As Long

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
                ElseIf ReplaceComponentCode(comp, newCode) Then
                    doneCount = doneCount + 1
                    doneLog = doneLog & "  - " & f(0) & "  ->  " & _
                              ConstValue(newCode, "SCRIPT_VERSION") & vbCrLf
                    ' Read from the DOWNLOADED code, not the manifest - this
                    ' is what actually got installed, guaranteed consistent
                    ' with it even if the manifest and the module disagree.
                    installedChangelog = ConstValue(newCode, "SCRIPT_CHANGELOG")
                    If Len(installedChangelog) > 0 Then
                        doneLog = doneLog & "      " & installedChangelog & vbCrLf
                    End If
                Else
                    failCount = failCount + 1
                    failLog = failLog & "  - " & f(0) & ": replace failed" & vbCrLf
                End If
            End If
        End If
    Next i

    MsgBox "MIDAS macro update" & vbCrLf & String(46, "-") & vbCrLf & _
           IIf(doneCount > 0, "UPDATED (" & doneCount & "):" & vbCrLf & doneLog, "") & _
           IIf(failCount > 0, vbCrLf & "FAILED (" & failCount & "):" & vbCrLf & failLog, "") & _
           vbCrLf & "Save the workbook to keep the updated code.", _
           IIf(failCount > 0, vbExclamation, vbInformation)

End Sub


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
'  SELF-UPDATE - see HOW THIS MODULE UPDATES ITSELF in the header
' ===========================================================================

' Steps 2-3: download, save, plant the bootstrap, schedule it, return.
' Nothing in this module is rewritten here - that happens only once this
' procedure and CheckMidasMacroUpdates have both finished.
Private Sub BeginSelfUpdate(ByVal compName As String)

    Dim newCode As String
    Dim tempPath As String
    Dim finishCall As String
    Dim fileNo As Integer
    Dim boot As Object
    Dim stage As String

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "The updater did not update itself: this workbook has never been " & _
               "saved, and it must be saved before the updater rewrites itself." & _
               vbCrLf & "Save it and run the check again.", vbExclamation
        Exit Sub
    End If

    If Not FetchText(UPDATE_BASE_URL & "/" & SCRIPT_ID & ".bas", newCode) Then
        MsgBox "The updater could not download its own new version:" & vbCrLf & _
               newCode, vbExclamation
        Exit Sub
    End If

    ' Stricter than InstallUpdates: it must be THIS module, not merely some
    ' managed one, or another module's code would overwrite the updater.
    If StrComp(ConstValue(newCode, "SCRIPT_ID"), SCRIPT_ID, vbTextCompare) <> 0 Then
        MsgBox "The downloaded updater is not the updater (SCRIPT_ID is """ & _
               ConstValue(newCode, "SCRIPT_ID") & """) - refusing to install it.", _
               vbExclamation
        Exit Sub
    End If

    ' The feed serves LF line endings; AddFromFile wants CRLF lines.
    newCode = Replace(newCode, vbCrLf, vbLf)
    newCode = Replace(newCode, vbCr, vbLf)
    newCode = Replace(newCode, vbLf, vbCrLf)

    ' Chain to FinishMidasSelfUpdate only if the new code still has it -
    ' Application.OnTime on a missing macro is a runtime error.
    If InStr(1, newCode, "Sub FinishMidasSelfUpdate", vbTextCompare) > 0 Then
        finishCall = QualifiedMacro("FinishMidasSelfUpdate")
    End If

    tempPath = Environ$("TEMP") & Application.PathSeparator & _
               "midas-macro-updater-" & Format$(Now, "yyyymmdd_hhnnss") & ".bas"

    On Error GoTo Failed

    stage = "write the new code to " & tempPath
    fileNo = FreeFile
    Open tempPath For Output As #fileNo
    Print #fileNo, newCode;
    Close #fileNo

    ' BEFORE anything risky. Every module InstallUpdates just replaced is on
    ' disk after this, so a crash further down costs only this step.
    stage = "save the workbook"
    ThisWorkbook.Save

    stage = "add the " & BOOTSTRAP_MODULE & " module"
    Call RemoveBootstrapModule
    Set boot = ThisWorkbook.VBProject.VBComponents.Add(STD_MODULE)
    boot.Name = BOOTSTRAP_MODULE
    With boot.CodeModule
        ' The VBE may already have inserted "Option Explicit".
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString BootstrapCode(compName, tempPath, finishCall)
    End With

    stage = "schedule the bootstrap"
    Application.OnTime Now, QualifiedMacro("MidasUpdaterBootstrapRun")
    Exit Sub

Failed:
    MsgBox "The updater did not update itself - it could not " & stage & ":" & _
           vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "The other updates are unaffected. Nothing in the updater was changed.", _
           vbExclamation
    On Error Resume Next
    Close #fileNo
    Call RemoveBootstrapModule

End Sub

' The bootstrap module's source. It is the ONLY code running when the
' updater is rewritten: it replaces the updater from tempPath, deletes the
' file, then hands over to finishCall (the NEW updater's
' FinishMidasSelfUpdate) - or, if the new code has none, just says so.
Private Function BootstrapCode(ByVal compName As String, ByVal tempPath As String, _
                               ByVal finishCall As String) As String

    Dim b As String

    b = "Option Explicit" & vbCrLf & vbCrLf
    b = b & "' Temporary - written by the MIDAS macro updater to replace itself." & vbCrLf
    b = b & "' Safe to delete if it is ever left behind." & vbCrLf & vbCrLf
    b = b & "Public Sub MidasUpdaterBootstrapRun()" & vbCrLf
    b = b & "    Dim comp As Object" & vbCrLf
    b = b & "    On Error GoTo Failed" & vbCrLf
    b = b & "    Set comp = ThisWorkbook.VBProject.VBComponents(" & VbaLiteral(compName) & ")" & vbCrLf
    b = b & "    With comp.CodeModule" & vbCrLf
    b = b & "        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines" & vbCrLf
    b = b & "        .AddFromFile " & VbaLiteral(tempPath) & vbCrLf
    b = b & "    End With" & vbCrLf
    b = b & "    On Error Resume Next" & vbCrLf
    b = b & "    Kill " & VbaLiteral(tempPath) & vbCrLf
    If Len(finishCall) > 0 Then
        b = b & "    Application.OnTime Now, " & VbaLiteral(finishCall) & vbCrLf
    Else
        b = b & "    MsgBox " & VbaLiteral("The MIDAS macro updater replaced itself. Delete " & _
                "the " & BOOTSTRAP_MODULE & " module and save the workbook.") & _
                ", vbInformation" & vbCrLf
    End If
    b = b & "    Exit Sub" & vbCrLf
    b = b & "Failed:" & vbCrLf
    b = b & "    MsgBox " & VbaLiteral("The MIDAS macro updater could not replace itself: ") & _
            " & Err.Description & vbCrLf & vbCrLf & " & _
            VbaLiteral("The workbook was saved just before this step. Close it " & _
                       "WITHOUT saving and reopen it to get the previous updater " & _
                       "back, then paste midas-macro-updater.bas in by hand.") & _
            ", vbExclamation" & vbCrLf
    b = b & "End Sub" & vbCrLf

    BootstrapCode = b

End Function

' Step 5. Runs as the NEW code, scheduled by the bootstrap, so the
' bootstrap is no longer running and can be removed. Public because
' Application.OnTime can only reach public procedures.
Public Sub FinishMidasSelfUpdate()

    Call RemoveBootstrapModule

    MsgBox "MIDAS macro updater updated itself to " & SCRIPT_VERSION & "." & _
           vbCrLf & vbCrLf & SCRIPT_CHANGELOG & vbCrLf & vbCrLf & _
           "Save the workbook to keep it.", vbInformation, "MIDAS macro updater"

End Sub

' Deletes the bootstrap module if one exists. Safe whenever the bootstrap
' itself is not the code running.
Private Sub RemoveBootstrapModule()

    Dim comp As Object

    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(BOOTSTRAP_MODULE)
    If Not comp Is Nothing Then ThisWorkbook.VBProject.VBComponents.Remove comp
    Err.Clear
    On Error GoTo 0

End Sub

' A macro name Application.OnTime resolves in THIS workbook even when
' another workbook is active.
Private Function QualifiedMacro(ByVal procName As String) As String
    QualifiedMacro = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & procName
End Function

' s as a VBA string literal, quotes doubled.
Private Function VbaLiteral(ByVal s As String) As String
    VbaLiteral = """" & Replace(s, """", """""") & """"
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


' Same lookup as ManifestVersion, but the 3rd pipe field (the one-line
' changelog) instead of the 2nd. "" when the line has no 3rd field - an
' older 2-field "id|version" manifest line still works, it just has
' nothing to show here.
Private Function ManifestChangelog(ByVal manifest As String, ByVal wantId As String) As String

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
                    If UBound(f) >= 2 Then ManifestChangelog = Trim$(f(2))
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
