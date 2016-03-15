VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtReport 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   1080
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileExitSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TOKEN_FORM = "Form="
Private Const TOKEN_MODULE = "Module="
Private Const TOKEN_CLASS = "Class="
Private Const TOKEN_END = "End"
Private Const TOKEN_BEGIN = "Begin"
Private Const TOKEN_PRIVATE = "Private "
Private Const TOKEN_PUBLIC = "Public "
Private Const TOKEN_SUB = "Sub "
Private Const TOKEN_FUNCTION = "Function "
Private Const TOKEN_PROPERTY = "Property "
' Return true if the line contains a comment.
Private Function ContainsComment(ByVal new_line As String) As Boolean
Dim i As Integer
Dim quotes_open As Boolean

    ' Assume we will find a comment.
    ContainsComment = True

    quotes_open = False
    For i = 1 To Len(new_line)
        Select Case Mid$(new_line, i, 1)
            Case """"
                quotes_open = Not quotes_open
            Case "'"
                If Not quotes_open Then Exit Function
        End Select
    Next i

    ' No comment.
    ContainsComment = False
End Function

' Count the rest of the lines.
Private Sub CountTheRest(ByVal fnum As Integer, ByRef num_lines As Long, ByRef num_subs As Long, ByRef num_functions As Long, ByRef num_properties As Long, ByRef num_Comments As Long)
Dim new_line As String

    ' Examine the lines.
    Do While Not EOF(fnum)
        Line Input #fnum, new_line
        num_lines = num_lines + 1

        If ContainsComment(new_line) Then num_Comments = num_Comments + 1

        ' Remove leading Private or Public.
        If PrefixMatches(new_line, TOKEN_PRIVATE) Then
            new_line = Mid$(new_line, Len(TOKEN_PRIVATE) + 1)
        ElseIf PrefixMatches(new_line, TOKEN_PUBLIC) Then
            new_line = Mid$(new_line, Len(TOKEN_PUBLIC) + 1)
        End If

        ' See if we got anything interesting.
        If PrefixMatches(new_line, TOKEN_SUB) Then
            num_subs = num_subs + 1
        ElseIf PrefixMatches(new_line, TOKEN_FUNCTION) Then
            num_functions = num_functions + 1
        ElseIf PrefixMatches(new_line, TOKEN_PROPERTY) Then
            num_properties = num_properties + 1
        End If
    Loop
End Sub
' Return a report about a form.
Private Function ReportOnForm(ByVal project_dir As String, ByVal form_file As String, ByRef total_lines As Long, ByRef total_controls As Long, ByRef total_subs As Long, ByRef total_functions As Long, ByRef total_properties As Long, ByRef total_comments As Long) As String
Dim txt As String
Dim fnum As Integer
Dim new_line As String
Dim begins_open As Long
Dim num_controls As Long
Dim num_lines As Long
Dim num_functions As Long
Dim num_subs As Long
Dim num_properties As Long
Dim num_Comments As Long

    txt = "Form:       " & form_file & vbCrLf

    fnum = FreeFile
    On Error Resume Next
    Open project_dir & form_file For Input As fnum
    If Err.Number <> 0 Then
        txt = txt & "*** Error " & _
            Format$(Err.Number) & _
            " opening form file " & _
            project_dir & form_file & _
            vbCrLf & Err.Description
        ReportOnForm = txt
        Exit Function
    End If
    On Error GoTo 0

    ' Find the "Begin VB.Form" line.
    Do While Not EOF(fnum)
        Line Input #fnum, new_line
        new_line = Trim$(new_line)
        If PrefixMatches(new_line, TOKEN_BEGIN) _
            Then Exit Do
    Loop

    ' Search for the corresponding End.
    begins_open = 1
    Do While Not EOF(fnum)
        Line Input #fnum, new_line
        new_line = Trim$(new_line)
        If PrefixMatches(new_line, TOKEN_BEGIN) Then
            num_controls = num_controls + 1
            begins_open = begins_open + 1
        ElseIf PrefixMatches(new_line, TOKEN_END) Then
            begins_open = begins_open - 1
            If begins_open < 1 Then Exit Do
        End If
    Loop

    ' Count the remaining lines.
    CountTheRest fnum, num_lines, num_subs, num_functions, num_properties, num_Comments

    ' Close the form file.
    Close fnum

    total_lines = total_lines + num_lines
    total_controls = total_controls + num_controls
    total_subs = total_subs + num_subs
    total_functions = total_functions + num_functions
    total_properties = total_properties + num_properties
    total_comments = total_comments + num_Comments
    ReportOnForm = txt & _
        "Lines:      " & Format$(num_lines, "@@@@@@@") & vbCrLf & _
        "Controls:   " & Format$(num_controls, "@@@@@@@") & vbCrLf & _
        "Subs:       " & Format$(num_subs, "@@@@@@@") & vbCrLf & _
        "Functions:  " & Format$(num_functions, "@@@@@@@") & vbCrLf & _
        "Properties: " & Format$(num_properties, "@@@@@@@") & vbCrLf & _
        "Comments:   " & Format$(num_Comments, "@@@@@@@") & vbCrLf
End Function
' Return a report about a BAS or CLS module.
Private Function ReportOnModule(ByVal project_dir As String, ByVal module_file As String, ByVal module_type As String, ByRef total_lines As Long, ByRef total_controls As Long, ByRef total_subs As Long, ByRef total_functions As Long, ByRef total_properties As Long, ByRef total_comments As Long) As String
Dim txt As String
Dim fnum As Integer
Dim new_line As String
Dim num_lines As Long
Dim num_functions As Long
Dim num_subs As Long
Dim num_properties As Long
Dim num_Comments As Long

    module_type = module_type & ":"
    txt = Format$(module_type, "!@@@@@@@@@@@@") & module_file & vbCrLf

    fnum = FreeFile
    On Error Resume Next
    Open project_dir & module_file For Input As fnum
    If Err.Number <> 0 Then
        txt = txt & "*** Error " & _
            Format$(Err.Number) & _
            " opening module file " & _
            project_dir & module_file & _
            vbCrLf & Err.Description
        ReportOnModule = txt
        Exit Function
    End If
    On Error GoTo 0

    ' Count all the lines.
    CountTheRest fnum, num_lines, num_subs, num_functions, num_properties, num_Comments

    ' Close the module file.
    Close fnum

    total_lines = total_lines + num_lines
    total_subs = total_subs + num_subs
    total_functions = total_functions + num_functions
    total_properties = total_properties + num_properties
    total_comments = total_comments + num_Comments
    ReportOnModule = txt & _
        "Lines:      " & Format$(num_lines, "@@@@@@@") & vbCrLf & _
        "Subs:       " & Format$(num_subs, "@@@@@@@") & vbCrLf & _
        "Functions:  " & Format$(num_functions, "@@@@@@@") & vbCrLf & _
        "Properties: " & Format$(num_properties, "@@@@@@@") & vbCrLf & _
        "Comments:   " & Format$(num_Comments, "@@@@@@@") & vbCrLf
End Function
' Return True if txt begins with prefix.
Private Function PrefixMatches(ByVal txt As String, ByVal prefix As String) As Boolean
    PrefixMatches = (Left$(txt, Len(prefix)) = prefix)
End Function

' Return a report on a project.
Private Function ReportOnProject(ByVal project_file As String) As String
Dim project_dir As String
Dim fnum As Integer
Dim new_line As String
Dim project_forms As Collection
Dim project_modules As Collection
Dim project_classes As Collection
Dim i As Integer
Dim txt As String
Dim total_lines As Long
Dim total_controls As Long
Dim total_subs As Long
Dim total_functions As Long
Dim total_properties As Long
Dim total_comments As Long

    ' Get the project directory.
    project_dir = Left$(project_file, InStrRev(project_file, "\"))

    ' Process the project file.
    fnum = FreeFile
    Open project_file For Input As fnum

    Set project_forms = New Collection
    Set project_modules = New Collection
    Set project_classes = New Collection

    ' Read the lines in the project file.
    Do While Not EOF(fnum)
        ' Read a line.
        Line Input #fnum, new_line

        ' See what this line is.
        If PrefixMatches(new_line, TOKEN_FORM) Then
            ' It's a form.
            project_forms.Add Mid$(new_line, Len(TOKEN_FORM) + 1)
        ElseIf PrefixMatches(new_line, TOKEN_MODULE) Then
            ' It's a module.
            project_modules.Add Trim$(Mid$(new_line, InStr(new_line, ";") + 1))
        ElseIf PrefixMatches(new_line, TOKEN_CLASS) Then
            ' It's a class.
            project_classes.Add Trim$(Mid$(new_line, InStr(new_line, ";") + 1))
        End If
    Loop

    ' Close the project file.
    Close fnum

    ' Sort the collections.
    SortCollection project_forms
    SortCollection project_modules

    ' Start the report.
    txt = "Project:    " & project_file & vbCrLf

    ' Process the forms.
    If project_forms.Count > 0 Then
        txt = txt & vbCrLf
        txt = txt & "*************" & vbCrLf
        txt = txt & "*** FORMS ***" & vbCrLf
        txt = txt & "*************" & vbCrLf
        For i = 1 To project_forms.Count
            txt = txt & vbCrLf & ReportOnForm(project_dir, project_forms(i), total_lines, total_controls, total_subs, total_functions, total_properties, total_comments)
        Next i
    End If

    ' Process the modules.
    If project_modules.Count > 0 Then
        txt = txt & vbCrLf
        txt = txt & "***************" & vbCrLf
        txt = txt & "*** MODULES ***" & vbCrLf
        txt = txt & "***************" & vbCrLf
        For i = 1 To project_modules.Count
            txt = txt & vbCrLf & ReportOnModule(project_dir, project_modules(i), "Module", total_lines, total_controls, total_subs, total_functions, total_properties, total_comments)
        Next i
    End If

    ' Process the classes.
    If project_classes.Count > 0 Then
        txt = txt & vbCrLf
        txt = txt & "***************" & vbCrLf
        txt = txt & "*** CLASSES ***" & vbCrLf
        txt = txt & "***************" & vbCrLf
        For i = 1 To project_classes.Count
            txt = txt & vbCrLf & ReportOnModule(project_dir, project_classes(i), "Class", total_lines, total_controls, total_subs, total_functions, total_properties, total_comments)
        Next i
    End If

    ' Return the result.
    txt = txt & vbCrLf
    txt = txt & "***************" & vbCrLf
    txt = txt & "*** SUMMARY ***" & vbCrLf
    txt = txt & "***************" & vbCrLf
    txt = txt & vbCrLf
    txt = txt & "Forms:            " & Format$(project_forms.Count, "@@@@@@@") & vbCrLf
    txt = txt & "Modules:          " & Format$(project_modules.Count, "@@@@@@@") & vbCrLf
    txt = txt & "Classes:          " & Format$(project_classes.Count, "@@@@@@@") & vbCrLf
    txt = txt & "Total Controls:   " & Format$(total_controls, "@@@@@@@") & vbCrLf
    txt = txt & "Total Lines:      " & Format$(total_lines, "@@@@@@@") & vbCrLf
    txt = txt & "Total Subs:       " & Format$(total_subs, "@@@@@@@") & vbCrLf
    txt = txt & "Total Functions:  " & Format$(total_functions, "@@@@@@@") & vbCrLf
    txt = txt & "Total Properties: " & Format$(total_properties, "@@@@@@@") & vbCrLf
    txt = txt & "Total Comments:   " & Format$(total_comments, "@@@@@@@") & vbCrLf

    ReportOnProject = txt
End Function
' Sort the collection.
Private Sub SortCollection(ByRef col As Collection)
Dim new_col As Collection
Dim i As Integer
Dim j As Integer
Dim txt As String

    Set new_col = New Collection

    For i = 1 To col.Count
        txt = col(i)
        For j = 1 To new_col.Count
            If new_col(j) >= txt Then Exit For
        Next j

        If j > new_col.Count Then
            new_col.Add txt
        Else
            new_col.Add txt, , j
        End If
    Next i

    Set col = new_col
End Sub
Private Sub Form_Load()
    dlgFile.InitDir = App.Path
End Sub

Private Sub Form_Resize()
    txtReport.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub


' Select a project file and process it.
Private Sub mnuFileOpen_Click()
Dim pos As Integer
Dim txt As String

    dlgFile.Flags = _
        cdlOFNExplorer Or _
        cdlOFNFileMustExist Or _
        cdlOFNHideReadOnly Or _
        cdlOFNLongNames
    dlgFile.Filter = "Project Files (*.vbp)|*.vbp"
    dlgFile.CancelError = True

    On Error Resume Next
    dlgFile.ShowOpen
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & _
            " selecting project file" & _
            vbCrLf & Err.Description
        Exit Sub
    End If
    On Error GoTo 0

    ' Save the directory.
    pos = InStrRev(dlgFile.FileName, "\")
    If pos > 0 Then
        dlgFile.InitDir = Left$(dlgFile.FileName, pos)
        dlgFile.FileName = dlgFile.FileTitle
    End If

    txtReport.Text = ""
    Screen.MousePointer = vbHourglass
    DoEvents

    ' Process the file.
    txtReport.Text = ReportOnProject(dlgFile.FileName)

    Screen.MousePointer = vbDefault
End Sub
