VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AstroGrep Properties"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOther 
      Height          =   285
      Left            =   2925
      TabIndex        =   10
      Top             =   675
      Width           =   870
   End
   Begin VB.CheckBox chkOther 
      Appearance      =   0  'Flat
      Caption         =   "Other"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2070
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox lblPathMRUCount 
      Height          =   285
      Left            =   3015
      TabIndex        =   8
      Text            =   "10"
      Top             =   90
      Width           =   735
   End
   Begin VB.CheckBox chkLF 
      Appearance      =   0  'Flat
      Caption         =   "Line Feed"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox chkCR 
      Appearance      =   0  'Flat
      Caption         =   "Carriage Return"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtEditor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4815
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "(Supports \r\n\t)"
      Height          =   240
      Left            =   4005
      TabIndex        =   11
      Top             =   720
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "End of line is denoted by:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1785
   End
   Begin VB.Label Label2 
      Caption         =   "Program to open files with on double click."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number of recently used paths to store:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' AstroGrep File Searching Utility. Written by Theodore L. Ward
' Copyright (C) 2002 AstroComma Incorporated.
'
' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
' The author may be contacted at:
' TheodoreWard@Hotmail.com or TheodoreWard@Yahoo.com

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()

    NUM_STORED_PATHS = val(Me.lblPathMRUCount.Text)
    DEFAULT_EDITOR = Me.txtEditor

    ENDOFLINEMARKER = ""
    
    'dzzie: allow for custom end of line markers..
    If chkOther.Value = 1 Then
        If Len(Trim(txtOther)) = 0 Then
            MsgBox "Delimiter can not be empty or space", vbInformation
            Exit Sub
        End If
        ENDOFLINEMARKER = txtOther
        ENDOFLINEMARKER = Replace(ENDOFLINEMARKER, "\r", vbCr)
        ENDOFLINEMARKER = Replace(ENDOFLINEMARKER, "\n", vbLf)
        ENDOFLINEMARKER = Replace(ENDOFLINEMARKER, "\t", vbTab)
    Else
        If Me.chkCR.Value = vbChecked Then ENDOFLINEMARKER = ENDOFLINEMARKER + vbCr
        If Me.chkLF.Value = vbChecked Then ENDOFLINEMARKER = ENDOFLINEMARKER + vbLf
    End If

    ' Store settings in registry.
    UpdateRegistrySettings
    
    ' Only store as many paths as has been set in options.
    With frmMain.cboFilePath

        Do While .ListCount > NUM_STORED_PATHS
            ' Remove the last item in the list.
            .RemoveItem .ListCount - 1
        Loop
    
    End With
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    ' Center the form in the main form.
    Me.Left = frmMain.Left + (frmMain.Width - Me.Width) / 2
    Me.Top = frmMain.Top + (frmMain.Height - Me.Height) / 2

    ' Initialize the form.
    Me.lblPathMRUCount.Text = NUM_STORED_PATHS
    'Me.UpDown1.Max = MAX_STORED_PATHS
    Me.txtEditor.Text = DEFAULT_EDITOR
    
    Dim otherEOL As Long
    Dim tmp As String
    
    'dzzie: txtother support
    otherEOL = GetSetting(appname:="AstroGrep", section:="Startup", Key:="otherEOL", Default:=0)
    If otherEOL = 1 Then
        chkOther.Value = 1
        tmp = Replace(ENDOFLINEMARKER, vbCr, "\r")
        tmp = Replace(tmp, vbLf, "\n")
        tmp = Replace(tmp, vbTab, "\t")
        txtOther = tmp
    Else
    
        If InStr(ENDOFLINEMARKER, vbCr) Then
            Me.chkCR.Value = vbChecked
        Else
            Me.chkCR.Value = vbUnchecked
        End If
        
        If InStr(ENDOFLINEMARKER, vbLf) Then
            Me.chkLF.Value = vbChecked
        Else
            Me.chkLF.Value = vbUnchecked
        End If
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "AstroGrep", "Startup", "otherEOL", chkOther.Value
End Sub
