VERSION 5.00
Begin VB.Form frmConv 
   Caption         =   "Character Conversion"
   ClientHeight    =   1260
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   7275
   Icon            =   "frmConversion2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1260
   ScaleWidth      =   7275
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBin 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   540
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   480
      Width           =   2670
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   195
      Left            =   1845
      TabIndex        =   10
      Top             =   1000
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   270
      Left            =   4980
      TabIndex        =   9
      Top             =   900
      Width           =   1005
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   165
      TabIndex        =   7
      Top             =   945
      Width           =   255
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   255
      Left            =   6180
      TabIndex        =   5
      Top             =   900
      Width           =   975
   End
   Begin VB.TextBox txtHex 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   4080
      TabIndex        =   4
      Top             =   480
      Width           =   3030
   End
   Begin VB.TextBox txtAsc 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      Top             =   60
      Width           =   3030
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   540
      TabIndex        =   0
      Top             =   105
      Width           =   2670
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Bin :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   540
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Space Hex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   2145
      TabIndex        =   11
      Top             =   945
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CGI Encode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Top             =   945
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Str"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   165
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hex :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3390
      TabIndex        =   2
      Top             =   525
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ascii :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3390
      TabIndex        =   1
      Top             =   120
      Width           =   525
   End
End
Attribute VB_Name = "frmConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal _
    hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx _
    As Long, ByVal cy As Long, ByVal wFlags As Long)
    
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
    
Dim inString As Boolean
Dim inHex As Boolean


Private Sub cmdClear_Click()
 Text1 = "": txtAsc = "": txtHex = ""
End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub


Private Sub Form_Load()
inString = True
inHex = False
SetWindowPos Me.hWnd, HWND_TOPMOST, Me.Left / 15, _
        Me.Top / 15, Me.Width / 15, _
        Me.Height / 15, SWP_SHOWWINDOW
End Sub

Private Sub Text1_GotFocus()
inString = True: inHex = False
End Sub


Private Sub txtasc_GotFocus()
inString = False: inHex = False
End Sub

Private Sub txthex_GotFocus()
inString = False: inHex = True
End Sub

Private Sub Text1_Change()
If Not inString Then Exit Sub
If Text1 = Empty Then txtHex = Empty: txtAsc = Empty
On Error Resume Next

Dim letter As String
p = " "

If Check1.value = 1 Then p = "%"
If Check2.value = 0 Then p = "": Check1.value = 0

If Len(Text1) > 1 Then letter = Mid(Text1, Len(Text1), 1) _
Else: letter = Text1

txtAsc = txtAsc & " " & Asc(letter)
txtHex = txtHex & p & Hex(Asc(letter))
txtBin = Empty
 
End Sub

Private Sub txthex_change()
If Not inHex Then Exit Sub
If txtHex = Empty Then Text1 = Empty: txtAsc = Empty

On Error Resume Next
If InStr(txtHex, " ") > 0 Then
    ary = Split(txtHex, " ")
    For i = 0 To UBound(ary)
        t1 = t1 & Chr(Int("&h" & ary(i))) & " "
        ta = ta & Int("&h" & ary(i)) & " "
    Next
    Text1 = t1: txtAsc = ta
Else
    Text1 = Chr(Int("&h" & txtHex))
    txtAsc = Int("&h" & txtHex)
    txtBin = dec2bin(Int("&h" & txtHex))
End If
    
    
End Sub

Private Sub txtasc_change()
If inString Or inHex Then Exit Sub
If txtAsc = Empty Then txtHex = Empty: Text1 = Empty

On Error Resume Next
Text1 = Chr(txtAsc)
txtHex = Hex(txtAsc)
txtBin = dec2bin(txtAsc)

End Sub


Public Function dec2bin(mynum As Variant) As String
    Dim loopcounter As Integer
    If mynum >= 2 ^ 31 Then Exit Function
    
    Do
        If (mynum And 2 ^ loopcounter) = 2 ^ loopcounter Then
            dec2bin = "1" & dec2bin
        Else
            dec2bin = "0" & dec2bin
        End If
        loopcounter = loopcounter + 1
    Loop Until 2 ^ loopcounter > mynum
    
    While Len(dec2bin) < 8
        dec2bin = "0" & dec2bin
    Wend
    
End Function

