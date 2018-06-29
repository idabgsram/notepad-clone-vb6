VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5430
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chCase 
      Caption         =   "Match Case"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox cboFind 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Find What :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'      NotePad Clone Form : Find
'----------------------------------------------
'Created     : 10/22/2012
'Re-Modified : 09/06/2014
'**********************************************

Option Explicit
Dim st As Long
Dim mmatchCase As Integer
Dim mWholeword As Integer
Dim found As Long
Dim vStrPos As Long
Dim StopNow As Boolean
Dim Searching As Boolean
Private Sub cboFind_Change()
    If Len(Trim(cboFind.Text)) = 0 Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
    st = 0
End Sub
Private Sub Command1_Click()
    If chCase.Value = 1 Then
        mmatchCase = 4
    Else
        mmatchCase = 0
    End If
    vStrPos = SendMessageByString&(cboFind.hwnd, CB_FINDSTRINGEXACT, 0, cboFind.Text)
    If vStrPos - 1 Then
    End If
    With frmMain.rtf
        found = .Find(cboFind.Text, st, , mWholeword Or mmatchCase)
        If found <> -1 Then
            st = found + Len(cboFind.Text)
            frmMain.XX.Text = cboFind.Text
            .SetFocus
        Else
            st = 0
            MsgBox "Cannot find " & cboFind.Text
        End If
    End With
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

