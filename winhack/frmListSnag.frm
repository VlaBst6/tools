VERSION 5.00
Begin VB.Form frmListSnag 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUCase 
      Caption         =   "UCase"
      Height          =   375
      Left            =   4620
      TabIndex        =   12
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdLcase 
      Caption         =   "LCase"
      Height          =   375
      Left            =   3300
      TabIndex        =   11
      Top             =   960
      Width           =   1155
   End
   Begin VB.TextBox txtExp 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Text            =   "*DEL*"
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kill Lines Like"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   960
      Width           =   1515
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   540
      TabIndex        =   8
      Top             =   540
      Width           =   1455
   End
   Begin VB.TextBox txtReplace 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2820
      TabIndex        =   6
      Top             =   540
      Width           =   1635
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   375
      Left            =   4620
      TabIndex        =   5
      Top             =   540
      Width           =   1575
   End
   Begin VB.CommandButton cmdListItemData 
      Caption         =   "Snag All ItemData"
      Height          =   435
      Left            =   4620
      TabIndex        =   4
      Top             =   60
      Width           =   1575
   End
   Begin VB.CommandButton cmdLstText 
      Caption         =   "Snag All ListText"
      Height          =   435
      Left            =   3000
      TabIndex        =   3
      Top             =   60
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetCount 
      Caption         =   "List Item Count"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1500
      Width           =   6075
   End
   Begin VB.Label Label1 
      Caption         =   "Find :                                     Replace :"
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   540
      Width           =   3675
   End
End
Attribute VB_Name = "frmListSnag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LBGetLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LBGetString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Const LB_GETCOUNT = &H18B
Private Const LB_GETCURSEL = &H188
Private Const LB_GETITEMDATA = &H199
Private Const LB_GETTEXT = &H189


Private Function AllListData(arytoFill() As String, Optional blnItemData As Boolean = False)
    Dim kount   As Integer
    Dim iNdx    As Integer
    Dim lhWnd   As Long
    Dim strVal  As String
    Dim itmData As Long
    Dim iPos    As Integer
    Dim ret As String
    
    lhWnd = CLng(frmWinHack.txthWnd)
    
    kount = LBGetLong(lhWnd, LB_GETCOUNT, 0, 0)  'Find total items
    For iNdx = 0 To kount - 1
        'Allocate space for return string
        strVal = String(256, " ")
        
        If Not blnItemData Then
            Call LBGetString(lhWnd, LB_GETTEXT, iNdx, strVal)
            
            'Convert C String to VB String
            iPos = InStr(strVal, Chr(0))
            If (iPos > 0) Then strVal = Left$(strVal, iPos - 1)
                
            push arytoFill(), strVal
        Else
        
            itmData = LBGetString(lhWnd, LB_GETITEMDATA, iNdx, 0)
            push arytoFill(), itmData
            
        End If
        
       
    Next
    
     
    
End Function

Private Sub cmdGetCount_Click()
    Text2 = LBGetLong(CLng(frmWinHack.txthWnd), LB_GETCOUNT, 0, 0)
End Sub

Private Sub cmdLcase_Click()
    Text1 = LCase(Text1)
End Sub

Private Sub cmdListItemData_Click()
    Dim ret() As String
    AllListData ret(), True
    Text1 = Join(ret, vbCrLf)
End Sub

Private Sub cmdLstText_Click()
    Dim ret() As String
    AllListData ret()
    Text1 = Replace(Join(ret, vbCrLf), Chr(0), "\x00")
End Sub

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Sub cmdReplace_Click()
    Dim find As String
    Dim rep As String
    
    find = ExpandConstants(txtFind)
    rep = ExpandConstants(txtReplace)
    
    Text1 = Replace(Text1, find, rep, , , vbTextCompare)
    
End Sub

Private Sub cmdUCase_Click()
    Text1 = UCase(Text1)
End Sub

Private Sub Command3_Click()
    
    If Len(txtExp) = 0 Then
        MsgBox "Enter expression to match, uses VB LIKE keyword", vbInformation
        Exit Sub
    End If
    
    Dim tmp, i
    
    tmp = Split(Text1, vbCrLf)
    For i = 0 To UBound(tmp)
        If tmp(i) Like txtExp Then tmp(i) = ""
    Next
    
    tmp = Join(tmp, vbCrLf)
    tmp = Replace(tmp, vbCrLf & vbCrLf, vbCrLf)
    Text1 = tmp
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Text1.Width = Me.Width - Text1.Left - 200
    Text1.Height = Me.Height - Text1.Top - 500
End Sub

Function ExpandConstants(ByVal it)
    it = Replace(it, "<CRLF>", vbCrLf)
    it = Replace(it, "<CR>", vbCr)
    it = Replace(it, "<LF>", vbLf)
    it = Replace(it, "<TAB>", vbTab)
    ExpandConstants = it
End Function
