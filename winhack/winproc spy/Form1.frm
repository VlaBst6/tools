VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Window HaViC"
   ClientHeight    =   4920
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   328
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   452
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4620
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Log to file in app.path"
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Window >"
      Height          =   300
      Left            =   60
      TabIndex        =   6
      ToolTipText     =   "use List Childs/all windows to find a new Handle then type in text box to right."
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox tbNewHwnd 
      Height          =   300
      Left            =   1380
      TabIndex        =   5
      Text            =   "0"
      Top             =   360
      Width           =   1125
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   6735
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   300
      Left            =   5760
      TabIndex        =   2
      Top             =   1080
      Width           =   810
   End
   Begin VB.CommandButton btnCapMsgs 
      Caption         =   "Cap Msgs"
      Height          =   300
      Left            =   60
      TabIndex        =   1
      Top             =   1080
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cap msg only works for window in current thread . so you can only cap this windows message"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   6600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current Window:"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   4380
   End
   Begin VB.Menu mnutest 
      Caption         =   "test"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlWindowhWnd As Long
Private msWindowText As String

Private Sub btnCapMsgs_Click()
    If mlWindowhWnd <> 0 Then
        If btnCapMsgs.Caption = "Cap Msgs (Stop)" Then
            Unhook
            btnCapMsgs.Caption = "Cap Msgs"
        Else
            Hook (mlWindowhWnd)
            btnCapMsgs.Caption = "Cap Msgs (Stop)"
        End If
    Else
        MsgBox "First, find a window"
    End If
End Sub

Private Sub btnClear_Click()
    List1.Clear
End Sub



Private Sub Command1_Click()
    Dim t&, s$
    Dim pp As POINTAPI
    Dim Nhwnd&
    
    Nhwnd = Val(tbNewHwnd.Text)
    If Nhwnd <> 0 Then
        s = Space$(512)
        GetWindowText Nhwnd, s, 512
        s = Trim$(s)
        gChileEnumContinue = False
        Unhook
        mlWindowhWnd = Nhwnd
        msWindowText = s
        Label1.Caption = "Current Window: (" & mlWindowhWnd & ")-" & msWindowText
    Else
        MsgBox "invalide handle, cannot = 0"
    End If
        
    
End Sub





Private Sub Form_Load()
    fillTable
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If gisHooked Then Unhook
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = ScaleWidth - List1.Left
    List1.Height = ScaleHeight - List1.Top
End Sub

Private Sub mnutest_Click()
Me.Caption = Me.Caption & "A"
End Sub


