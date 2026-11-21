Attribute VB_Name = "IEControlScan"
'===============================================================================
' IEControlScan  v0.1.0
' Finds every legacy (Internet Explorer based) Web Browser Control in an Access
' database, the subforms and navigation targets that host them, and the VBA lines
' that use the legacy MSHTML object model.
'
' Why: Microsoft is retiring the legacy Web Browser Control in Access
' (Current Channel version 2705, about 2027-05-18; Access LTSC 2024 on 2027-05-11).
' After that the legacy control stops loading web content. Microsoft's own
' guidance says finding the controls is a manual, form-by-form job.
'
' Read-only: forms and reports are opened hidden in Design View and closed
' with acSaveNo. Nothing in your database is changed. Make a backup anyway.
'
' Usage: import this module (VBA editor > File > Import File), then in the
' Immediate window run:   ScanIEControls
' A CSV report is written next to the database and a summary is printed.
'
' MIT License. https://github.com/fengyiqicoder/access-ie-control-scan
'===============================================================================
Option Compare Database
Option Explicit

Private Const AC_WEB_BROWSER As Long = 128   ' acWebBrowser (legacy, IE based)
Private Const AC_EDGE_BROWSER As Long = 134  ' acEdgeBrowser (new, WebView2)
Private Const AC_CUSTOM_CONTROL As Long = 119 ' acCustomControl (ActiveX)
Private Const AC_SUBFORM As Long = 112       ' acSubform
Private Const AC_NAV_BUTTON As Long = 130    ' acNavigationButton

Private mRows As Collection
Private mLegacyNames As Collection   ' "FormName|ControlName"
Private mLegacy As Long, mEdge As Long, mActiveX As Long, mCodeMust As Long, mCodeCheck As Long

Public Sub ScanIEControls()
    Dim i As Long, outPath As String
    Set mRows = New Collection
    Set mLegacyNames = New Collection
    mLegacy = 0: mEdge = 0: mActiveX = 0: mCodeMust = 0: mCodeCheck = 0

    For i = 0 To CurrentProject.AllForms.Count - 1
        ScanObject "Form", CurrentProject.AllForms(i).Name
    Next i
    For i = 0 To CurrentProject.AllReports.Count - 1
        ScanObject "Report", CurrentProject.AllReports(i).Name
    Next i

    ScanCode

    outPath = WriteCsv()

    Debug.Print String(60, "-")
    Debug.Print "IEControlScan finished " & Format(Now, "yyyy-mm-dd hh:nn")
    Debug.Print "Legacy Web Browser Controls (must convert): " & mLegacy
    Debug.Print "IE ActiveX Shell.Explorer controls (check):  " & mActiveX
    Debug.Print "Edge Browser Controls (already new):        " & mEdge
    Debug.Print "VBA lines that must be rewritten:           " & mCodeMust
    Debug.Print "VBA lines to check:                         " & mCodeCheck
    Debug.Print "Report: " & outPath
    Debug.Print String(60, "-")
    MsgBox "Legacy Web Browser Controls: " & mLegacy & vbCrLf & _
           "VBA lines to rewrite: " & mCodeMust & "  (to check: " & mCodeCheck & ")" & vbCrLf & vbCrLf & _
           "Report saved to:" & vbCrLf & outPath, vbInformation, "IEControlScan"
End Sub

Private Sub ScanObject(ByVal kind As String, ByVal objName As String)
    Dim wasLoaded As Boolean, obj As Object, ctl As Control
    On Error GoTo Failed

    If kind = "Form" Then
        wasLoaded = CurrentProject.AllForms(objName).IsLoaded
        If Not wasLoaded Then DoCmd.OpenForm objName, acDesign, , , , acHidden
        Set obj = Forms(objName)
    Else
        wasLoaded = CurrentProject.AllReports(objName).IsLoaded
        If Not wasLoaded Then DoCmd.OpenReport objName, acViewDesign, , , acHidden
        Set obj = Reports(objName)
    End If

    For Each ctl In obj.Controls
        Select Case ctl.ControlType
            Case AC_WEB_BROWSER
                mLegacy = mLegacy + 1
                mLegacyNames.Add objName & "|" & ctl.Name
                AddRow kind, objName, ctl.Name, "LEGACY Web Browser Control", "MUST CONVERT", _
                       "ControlSource: " & SafeProp(ctl, "ControlSource")
            Case AC_EDGE_BROWSER
                mEdge = mEdge + 1
                AddRow kind, objName, ctl.Name, "Edge Browser Control", "OK", _
                       "ControlSource: " & SafeProp(ctl, "ControlSource")
            Case AC_CUSTOM_CONTROL
                If InStr(1, SafeProp(ctl, "Class"), "Shell.Explorer", vbTextCompare) > 0 Then
                    mActiveX = mActiveX + 1
                    mLegacyNames.Add objName & "|" & ctl.Name
                    AddRow kind, objName, ctl.Name, "ActiveX " & SafeProp(ctl, "Class"), "CHECK", _
                           "IE based ActiveX browser; not the control named in the notice, but it uses the same engine"
                End If
            Case AC_SUBFORM
                AddRow kind, objName, ctl.Name, "Subform control", "INFO", _
                       "Hosts: " & SafeProp(ctl, "SourceObject")
            Case AC_NAV_BUTTON
                If Len(SafeProp(ctl, "NavigationTargetName")) > 0 Then
                    AddRow kind, objName, ctl.Name, "Navigation button", "INFO", _
                           "Opens: " & SafeProp(ctl, "NavigationTargetName")
                End If
        End Select
    Next ctl

CloseIt:
    On Error Resume Next
    If Not wasLoaded Then
        If kind = "Form" Then DoCmd.Close acForm, objName, acSaveNo Else DoCmd.Close acReport, objName, acSaveNo
    End If
    Exit Sub
Failed:
    AddRow kind, objName, "", "Could not scan this object", "CHECK BY HAND", Err.Description
    Resume CloseIt
End Sub

Private Sub ScanCode()
    Dim proj As Object, comp As Object, cm As Object
    Dim n As Long, ln As Long, txt As String, hit As String, level As String
    On Error GoTo NoVBE

    Set proj = CurrentVBProject()
    If proj Is Nothing Then GoTo NoVBE

    For Each comp In proj.VBComponents
        If comp.Name <> "IEControlScan" Then
            Set cm = comp.CodeModule
            n = cm.CountOfLines
            For ln = 1 To n
                txt = cm.Lines(ln, 1)
                If Not IsComment(txt) Then
                    hit = CodeHit(txt, level)
                    If Len(hit) > 0 Then
                        If level = "MUST REWRITE" Then mCodeMust = mCodeMust + 1 Else mCodeCheck = mCodeCheck + 1
                        AddRow "Code", comp.Name, "line " & ln, hit, level, Trim$(txt)
                    End If
                End If
            Next ln
        End If
    Next comp
    Exit Sub
NoVBE:
    AddRow "Code", "", "", "VBA code not scanned", "CHECK BY HAND", _
           "Could not read the VBA project (" & Err.Description & "). Search your code for .Object, .Document and MSHTML."
End Sub

Private Function CurrentVBProject() As Object
    Dim p As Object
    For Each p In Application.VBE.VBProjects
        If StrComp(p.FileName, CurrentProject.FullName, vbTextCompare) = 0 Then
            Set CurrentVBProject = p
            Exit Function
        End If
    Next p
    Set CurrentVBProject = Application.VBE.ActiveVBProject
End Function

' Returns a short reason when the line touches the legacy object model.
Private Function CodeHit(ByVal txt As String, ByRef level As String) As String
    Dim lc As String, v As Variant, parts() As String
    lc = LCase$(txt)
    level = "MUST REWRITE"
    If InStr(lc, "mshtml") > 0 Then CodeHit = "MSHTML reference": Exit Function
    If InStr(lc, "htmldocument") > 0 Or InStr(lc, "ihtml") > 0 Then CodeHit = "MSHTML type": Exit Function
    If InStr(lc, "shell.explorer") > 0 Or InStr(lc, "internetexplorer.application") > 0 Then CodeHit = "IE automation": Exit Function
    If InStr(lc, "execwb") > 0 Then CodeHit = "ExecWB (legacy only)": Exit Function

    ' Lines that name a legacy control found above
    For Each v In mLegacyNames
        parts = Split(v, "|")
        If InStr(lc, LCase$(parts(1))) > 0 Then
            If InStr(lc, ".object") > 0 Then CodeHit = "Uses .Object of legacy control " & parts(1): Exit Function
            If InStr(lc, ".document") > 0 Then CodeHit = "Uses .Document of legacy control " & parts(1): Exit Function
            If InStr(lc, "navigate2") > 0 Then CodeHit = "Navigate2 on legacy control " & parts(1): Exit Function
            level = "CHECK"
            CodeHit = "Mentions legacy control " & parts(1)
            Exit Function
        End If
    Next v

    level = "CHECK"
    If InStr(lc, ".document.") > 0 Then CodeHit = ".Document access (check which control)": Exit Function
    If InStr(lc, "getelementbyid") > 0 Or InStr(lc, "getelementsbytagname") > 0 Then CodeHit = "DOM call (check which control)": Exit Function
    CodeHit = ""
End Function

Private Function IsComment(ByVal txt As String) As Boolean
    Dim t As String
    t = LTrim$(txt)
    IsComment = (Left$(t, 1) = "'") Or (LCase$(Left$(t, 4)) = "rem ")
End Function

Private Function SafeProp(ByVal ctl As Object, ByVal propName As String) As String
    On Error Resume Next
    SafeProp = CStr(ctl.Properties(propName).Value)
End Function

Private Sub AddRow(ParamArray f() As Variant)
    Dim s As String, i As Long
    For i = LBound(f) To UBound(f)
        If i > LBound(f) Then s = s & ","
        s = s & """" & Replace(CStr(f(i)), """", """""") & """"
    Next i
    mRows.Add s
End Sub

Private Function WriteCsv() As String
    Dim path As String, fn As Integer, r As Variant
    path = CurrentProject.Path & "\ie-control-scan-" & Format(Now, "yyyymmdd-hhnnss") & ".csv"
    fn = FreeFile
    Open path For Output As #fn
    Print #fn, """Type"",""Object"",""Control or line"",""Finding"",""Action"",""Detail"""
    For Each r In mRows
        Print #fn, r
    Next r
    Close #fn
    WriteCsv = path
End Function
