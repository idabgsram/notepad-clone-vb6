VERSION 5.00
Begin VB.Form frmReplace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Replace"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chWord 
      Caption         =   "Whole Word"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox chCase 
      Caption         =   "Match Case"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace All"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton CmdReplace 
      Caption         =   "Replace"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox CboReplace 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox cboFind 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Replace With:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Find What :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'      NotePad Clone Form : Replace
'----------------------------------------------
'Created     : 10/22/2012
'Re-Modified : 09/06/2014
'**********************************************

Dim st As Long
Dim mmatchCase As Integer
Dim mWholeword As Integer
Dim found As Long
Dim vStrPos As Long
Dim StopNow As Boolean
Dim Searching As Boolean
Private Sub cboFind_Change()
    If Len(Trim(cboFind.Text)) = 0 Then
        cmdFindNext.Enabled = False
        CmdReplace.Enabled = False
        cmdReplaceAll.Enabled = False
    Else
        cmdFindNext.Enabled = True
        CmdReplace.Enabled = True
        cmdReplaceAll.Enabled = True
    End If
    st = 0
End Sub
Private Sub cmdCancel_Click()
    If Searching Then
        StopNow = True
        Searching = False
    Else
        Unload Me
    End If
End Sub
Private Sub cmdFindNext_Click()
    If chWord.Value = 1 Then
        mWholeword = 2
    Else
        mWholeword = 0
    End If
    If chCase.Value = 1 Then
        mmatchCase = 4
    Else
        mmatchCase = 0
    End If
    'No dupes thanks
    vStrPos = SendMessageByString&(cboFind.hwnd, CB_FINDSTRINGEXACT, 0, cboFind.Text)
    If vStrPos - 1 Then
     '   cboFind.AddItem cboFind.Text
    End If
    With frmMain.rtf
        LockWindowUpdate frmMain.hwnd
        found = .Find(cboFind.Text, st, , mWholeword Or mmatchCase)
        If found <> -1 Then
            st = found + Len(cboFind.Text)
            CmdReplace.Enabled = True
            .SetFocus
        Else
            st = 0
            CmdReplace.Enabled = False
            MsgBox "Text not found."
        End If
        LockWindowUpdate 0
    End With
End Sub
Private Sub cmdReplace_Click()
'    If frmMain.rtf.SelText = cboFind.Text Then
'        frmMain.Undo.InsertText CboReplace.Text
'    Else
'        cmdFindNext_Click
'        If frmMain.rtf.SelText = cboFind.Text Then frmMain.Undo.InsertText CboReplace.Text
'    End If
    'No dupes thanks
'    vStrPos = SendMessageByString&(cboFind.hwnd, CB_FINDSTRINGEXACT, 0, cboFind.Text)
'    If vStrPos = -1 Then
      '  cboFind.AddItem cboFind.Text
'    End If
'   vStrPos = SendMessageByString&(CboReplace.hwnd, CB_FINDSTRINGEXACT, 0, CboReplace.Text)
'    If vStrPos = -1 Then
      '  CboReplace.AddItem CboReplace.Text
'    End If
'    FileChanged = True
End Sub
Private Sub cmdReplaceAll_Click()
    Dim count As Long, beginST As Long
    Searching = True
    Screen.MousePointer = 11
    If chWord.Value = 1 Then
        mWholeword = 2
    Else
        mWholeword = 0
    End If
    If chCase.Value = 1 Then
        mmatchCase = 4
    Else
        mmatchCase = 0
    End If
    vStrPos = SendMessageByString&(cboFind.hwnd, CB_FINDSTRINGEXACT, 0, cboFind.Text)
    If vStrPos = -1 Then
       ' cboFind.AddItem cboFind.Text
    End If
    vStrPos = SendMessageByString&(CboReplace.hwnd, CB_FINDSTRINGEXACT, 0, CboReplace.Text)
    If vStrPos = -1 Then
       ' CboReplace.AddItem CboReplace.Text
    End If
    NoStatusUpdate = True
    With frmMain.rtf
        .SelStart = 0
        beginST = .SelStart
        LockWindowUpdate frmMain.hwnd
        Do
            DoEvents
            If StopNow Then Exit Do
            found = .Find(cboFind.Text, st, , mWholeword Or mmatchCase)
            If found <> -1 Then
                st = found + Len(cboFind.Text)
                count = count + 1
 '               frmMain.Undo.InsertText CboReplace.Text, False
            Else
                st = 0
                CmdReplace.Enabled = False
                If count = 0 Then
                    MsgBox "Cannot find " & cboFind.Text
                    Exit Do
                Else
                    FileChanged = True
                    Exit Do
                End If
            End If
            If StopNow Then Exit Do
        Loop
        StopNow = False
        .SelStart = beginST
        .SelLength = 0
        Searching = False
        NoStatusUpdate = False
'        If count > 0 Then frmMain.Undo.UpdateStateChange
        Screen.MousePointer = 0
      '  MsgBox count & IIf(count = 1, " item replaced", " items replaced")
      '  LockWindowUpdate 0
    End With
End Sub
Private Sub Command1_Click()
    cboFind.Text = frmMain.rtf.SelText
End Sub
Private Sub Command2_Click()
    CboReplace.Text = frmMain.rtf.SelText
End Sub
Private Sub cmdSelection_Click(Index As Integer)
    If Index = 0 Then
        cboFind.Text = frmMain.rtf.SelText
    Else
        CboReplace.Text = frmMain.rtf.SelText
    End If
End Sub
Private Sub Form_Load()
cboFind.BackColor = frmMain.rtf.BackColor
CboReplace.BackColor = frmMain.rtf.BackColor
End Sub


