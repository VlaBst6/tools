VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   2700
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Menu mnuTopMenu 
      Caption         =   "Topmenu"
      Begin VB.Menu mnuSubmenu1 
         Caption         =   "Submenu1"
      End
      Begin VB.Menu mnuSubmenu2 
         Caption         =   "Submenu2"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSubmenu3 
         Caption         =   "Submenu3"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSubmenu4 
         Caption         =   "SubMenu4"
      End
      Begin VB.Menu mnuSubmenu5 
         Caption         =   "Submenu5"
         Begin VB.Menu mnuSubSubmenu1 
            Caption         =   "SubSubmenu1"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnuSubmenu2_Click()
    MsgBox "mnuSubmenu2_Click:  naughty boy!"
End Sub

Private Sub mnuSubmenu4_Click()
    MsgBox "submenu4 clicked"
End Sub

Private Sub Text1_Change()
    MsgBox "Text1 Changed"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
MsgBox "Text1_KeyDown Keycode=" & KeyCode
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Text1_KeyUp Keycode=" & KeyCode
End Sub
