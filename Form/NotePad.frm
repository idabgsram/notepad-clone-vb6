VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Untitled - NotePad Clone"
   ClientHeight    =   6405
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   8565
   Icon            =   "NotePad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox XX 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6150
      Visible         =   0   'False
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   14467
            MinWidth        =   14467
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6588
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"NotePad.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   480
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Statussave 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"NotePad.frx":094D
   End
   Begin VB.Menu mnfile 
      Caption         =   "File"
      Begin VB.Menu mnnewdoc 
         Caption         =   "New"
      End
      Begin VB.Menu mnopendoc 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnsavedoc 
         Caption         =   "Save"
      End
      Begin VB.Menu mnsaveasdoc 
         Caption         =   "Save As"
      End
      Begin VB.Menu strip1 
         Caption         =   "-"
      End
      Begin VB.Menu mnpagesetup 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu mnprintdoc 
         Caption         =   "Print..."
      End
      Begin VB.Menu strip2 
         Caption         =   "-"
      End
      Begin VB.Menu mnExitNow 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnedit 
      Caption         =   "Edit"
      Begin VB.Menu mnundoText 
         Caption         =   "Undo"
      End
      Begin VB.Menu strip3 
         Caption         =   "-"
      End
      Begin VB.Menu mncutText 
         Caption         =   "Cut"
      End
      Begin VB.Menu mncopyText 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnPasteText 
         Caption         =   "Paste"
      End
      Begin VB.Menu mndeleteText 
         Caption         =   "Delete"
      End
      Begin VB.Menu strip4 
         Caption         =   "-"
      End
      Begin VB.Menu mnfindNow 
         Caption         =   "Find..."
      End
      Begin VB.Menu mnfindNextNow 
         Caption         =   "Find Next"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnreplace 
         Caption         =   "Replace..."
      End
      Begin VB.Menu mngotoLine 
         Caption         =   "Go To..."
      End
      Begin VB.Menu Strip5 
         Caption         =   "-"
      End
      Begin VB.Menu selectallNow 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnadddate 
         Caption         =   "Date/Time"
      End
   End
   Begin VB.Menu mnformat 
      Caption         =   "Format"
      Begin VB.Menu mnwordWrapNow 
         Caption         =   "Word Wrap"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnfontNow 
         Caption         =   "Font..."
      End
   End
   Begin VB.Menu mnview 
      Caption         =   "View"
      Begin VB.Menu mnstatusbar 
         Caption         =   "Statusbar"
      End
   End
   Begin VB.Menu mnhelp 
      Caption         =   "Help"
      Begin VB.Menu mnaboutwin 
         Caption         =   "About"
      End
      Begin VB.Menu mnaboutclone 
         Caption         =   "About NotePad Clone"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'      NotePad Clone Form : Main
'----------------------------------------------
'Created     : 10/22/2012
'Re-Modified : 09/06/2014
'**********************************************

Option Explicit
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Dim st As Long
Dim mmatchCase As Integer
Dim mWholeword As Integer
Dim found As Long
Dim vStrPos As Long
Dim StopNow As Boolean
Dim Searching As Boolean
Dim sv As Integer
Dim LineA As String
Dim LineB As String
Dim LineC As String
Dim LineD As String
Dim LineE As String
Dim LineF As String
Dim LineG As String

Private Sub Form_Load()
InitConfig
sv = 0
    Dim mycommand As String
        mycommand = Command
        If Dir(mycommand) = "" Then
            If XMsgBox("Cannot find the " & FileOnly(mycommand) & "." & vbCrLf & "Do you want to create the file?", vbExclamation + vbDefaultButton1 + vbYesNoCancel, "File not found") = vbNo Or vbCancel Then
                Unload Me
            ElseIf vbYes Then
                Open App.Path & "\" & mycommand & ".txt" For Output As #1
                Print #1, ""
                Close #1
                Exit Sub
            End If
        End If
        If mycommand <> "" Then
            DoEvents
            mycommand = strUnQuoteString(mycommand)
            mycommand = GetLongFilename(mycommand)
            Select Case LCase(ExtOnly(mycommand))
                Case "txt"
                    rtf.SelText = OneGulp(mycommand)
                Case Else
                    rtf.SelText = OneGulp(mycommand)
            End Select
            Me.Caption = FileOnly(mycommand) & " - NotePad Clone"
            rtf.Tag = mycommand
            If FileLen(mycommand) > 100000 Then
                rtf.SelStart = 0
                rtf.SelLength = Len(rtf.Text)
                rtf.SelLength = 0
            End If
            rtf.SelStart = 0
        End If
sb1.Panels(1).Text = "| Ln " & Format(GetLineNum(frmMain.rtf) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMain.rtf) + 1)
End Sub

Private Sub InitConfig()
On Error Resume Next
    If Dir(App.Path & "\NotePad.cfg") = "" Then
        Open App.Path & "\NotePad.cfg" For Output As #1
        Print #1, "[notepad-config]"
        Print #1, "STATUSBAR=0"
        Print #1, "FONT="
        Print #1, "Lucida Console"
        Print #1, "FONTSIZE="
        Print #1, "10"
        Print #1, "[end-config]"
        Close #1
    Else
        Open App.Path & "\NotePad.cfg" For Input As #1
        Line Input #1, LineA
            If LineA <> "[notepad-config]" Then
                Close #1
                Open App.Path & "\NotePad.cfg" For Output As #1
                Print #1, "[notepad-config]"
                Print #1, "STATUSBAR=0"
                Print #1, "FONT="
                Print #1, "Lucida Console"
                Print #1, "FONTSIZE="
                Print #1, "10"
                Print #1, "[end-config]"
                Close #1
            End If
        Close #1
    End If
Open App.Path & "\NotePad.cfg" For Input As #1
Line Input #1, LineA
Line Input #1, LineB
    If Right(LineB, 1) = "0" Then
        sb1.Visible = False
        mnstatusbar.Checked = False
    End If
    If Right(LineB, 1) = "1" Then
        sb1.Visible = True
        mnstatusbar.Checked = True
    End If
    If Right(LineB, 1) <> "1" Then
        sb1.Visible = False
        mnstatusbar.Checked = False
    End If
Line Input #1, LineC
Line Input #1, LineD
    rtf.Font = LineD
Line Input #1, LineE
Line Input #1, LineF
    rtf.SelFontSize = LineF
Line Input #1, LineG
Close #1
End Sub

Private Sub SaveConfig()
If sb1.Visible = True Then
LineB = 1
Else
LineB = 0
End If
LineD = rtf.Font
LineF = rtf.SelFontSize
Open App.Path & "\NotePad.cfg" For Output As #1
Print #1, "[notepad-config]"
Print #1, "STATUSBAR=" & LineB
Print #1, "FONT="
Print #1, LineD
Print #1, "FONTSIZE="
Print #1, LineF
Print #1, "[end-config]"
Close #1
End Sub

Private Sub Form_Resize()
On Error GoTo jhx_cli
        rtf.Width = ScaleWidth
       If mnstatusbar.Checked = True Then
               rtf.Height = ScaleHeight - sb1.Height
               sb1.Panels(1).MinWidth = Me.Width - 500
               Else
            rtf.Height = ScaleHeight
            End If
        If mnwordWrapNow.Checked = True Then
            rtf.RightMargin = rtf.Width - 350
        Else
            rtf.RightMargin = rtf.Width + 10025
        End If
jhx_cli: Exit Sub
End Sub

Private Sub Form_Terminate()
SaveConfig
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveConfig
End
End Sub

Private Sub mnaboutclone_Click()
frmAbout.Show
End Sub

Private Sub mnaboutwin_Click()
ShellAbout Me.hwnd, "Software", "Modified by Idabgsram", Me.Icon
End Sub

Private Sub mnadddate_Click()
rtf.SelText = rtf.SelText & " " & time & " " & Date
End Sub

Private Sub mncopyText_Click()
Clipboard.Clear
Clipboard.SetText rtf.SelText
End Sub

Private Sub mncutText_Click()
Clipboard.Clear
Clipboard.SetText rtf.SelText
rtf.SelText = ""
End Sub

Private Sub mndeleteText_Click()
SendMessage rtf.hwnd, WM_CLEAR, 0&, 0&
End Sub

Private Sub mnExitNow_Click()
If sv = 1 Then
If Me.Caption = "Untitled - NotePad Clone" Then
Select Case XMsgBox("Do you want to save changes to Untitled?", vbQuestion + vbDefaultButton1 + vbYesNoCancel, "NotePad Clone")
    Case vbYes
    SaveDoc
    rtf.Text = ""
    sv = 0
    Unload Me
    Case vbNo
    Unload Me
    Case vbCancel
    Exit Sub
End Select
Else
Select Case XMsgBox("Do you want to save changes to " & cd1.filetitle & "?", vbQuestion + vbDefaultButton1 + vbYesNoCancel, "NotePad Clone")
    Case vbYes
    SaveDoc
    rtf.Text = ""
    Unload Me
    sv = 0
    Case vbNo
    rtf.Text = ""
    Unload Me
    sv = 0
    Case vbCancel
    Exit Sub
End Select
End If
Else
Unload Me
End If
End Sub

Private Sub mnfindNextNow_Click()
    If frmFind.chCase.Value = 1 Then
        mmatchCase = 4
    Else
        mmatchCase = 0
    End If
    vStrPos = SendMessageByString&(frmFind.cboFind.hwnd, CB_FINDSTRINGEXACT, 0, frmFind.cboFind.Text)
    If vStrPos - 1 Then
    End If
    With frmMain.rtf
        found = .Find(frmFind.cboFind.Text, st, , mWholeword Or mmatchCase)
        If found <> -1 Then
            st = found + Len(frmFind.cboFind.Text)
            frmMain.XX.Text = frmFind.cboFind.Text
            .SetFocus
        Else
            st = 0
            MsgBox "Cannot find " & frmFind.cboFind.Text
        End If
    End With
End Sub

Private Sub mnfindNow_Click()
On Error Resume Next
'frmfind.Show
  Dim s As String
  If rtf.SelLength > 0 Then s = rtf.SelText Else s = ""
  ShowFind Me, rtf, FR_DOWN, s
End Sub

Private Sub mnfontNow_Click()
cd1.FontName = rtf.SelFontName
cd1.FontSize = rtf.SelFontSize
cd1.FontBold = rtf.SelBold
cd1.FontItalic = rtf.SelItalic
cd1.FontUnderline = rtf.SelUnderline
cd1.FontStrikethru = rtf.SelStrikeThru
cd1.ShowFont
rtf.SelBold = cd1.FontBold
rtf.SelItalic = cd1.FontItalic
rtf.SelUnderline = cd1.FontUnderline
rtf.SelFontName = cd1.FontName
rtf.SelFontSize = cd1.FontSize
rtf.SelStrikeThru = cd1.FontStrikethru
End Sub

Private Sub mngotoLine_Click()
Dim linenumber
    linenumber = InputBox("Line Number : ", "Go To Line")
    If IsNumeric(linenumber) Then
        If LineCount < CLng(linenumber) Then
           MsgBox "The line number is beyond the total number of lines", , "NotePad Clone - Go To Line"
        Else
            rtf.SelStart = GetCharFromLine(CLng(linenumber) - 1)
        End If
    End If
End Sub

Private Sub mnnewdoc_Click()
If sv = 1 Then
If Me.Caption = "Untitled - NotePad Clone" Then
Select Case XMsgBox("Do you want to save changes to Untitled?", vbDefaultButton1 + vbYesNoCancel, "NotePad Clone")
    Case vbYes
    SaveDoc
    rtf.Text = ""
    sv = 0
    Me.Caption = "Untitled - NotePad Clone"
    Case vbNo
    rtf.Text = ""
    sv = 0
    Me.Caption = "Untitled - NotePad Clone"
    Case vbCancel
    Exit Sub
End Select
Else
Select Case XMsgBox("Do you want to save changes to " & cd1.filetitle & "?", vbDefaultButton1 + vbYesNoCancel, "NotePad Clone")
    Case vbYes
    SaveDoc
    rtf.Text = ""
    sv = 0
    Me.Caption = "Untitled - NotePad Clone"
    Case vbNo
    rtf.Text = ""
    sv = 0
    Me.Caption = "Untitled - NotePad Clone"
    Case vbCancel
    Exit Sub
End Select
End If
Else
rtf.Text = ""
sv = 0
Me.Caption = "Untitled - NotePad Clone"
End If
End Sub
Sub SaveAs()
  Dim sFile As String
        If Me Is Nothing Then Exit Sub
                rtf.Tag = sFile
                cd1.dialogtitle = "Save As"
                cd1.Filter = "Text Files (*.txt)|*.txt|All files (*.*)|*.*"
                cd1.flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
    cd1.CancelError = False
    cd1.ShowSave
   
                If Len(cd1.FileName) = 0 Then
                        Exit Sub
                End If
                 sFile = cd1.FileName
        rtf.SaveFile sFile
        Me.Caption = cd1.filetitle & " - NotePad Clone" '"RoteNote - " & FileOnly(sFile)
        Statussave.Text = "False"
     If Err Then Exit Sub
End Sub
Private Sub mnopendoc_Click()
If sv = 1 Then
If Me.Caption = "Untitled - NotePad Clone" Then
Select Case XMsgBox("Do you want to save changes to Untitled?", vbDefaultButton1 + vbYesNoCancel, "NotePad Clone")
    Case vbYes
    SaveDoc
    OpenDoc
    sv = 0
    Case vbNo
    OpenDoc
    sv = 0
    Case vbCancel
    Exit Sub
End Select
Else
Select Case XMsgBox("Do you want to save changes to " & cd1.filetitle & "?", vbDefaultButton1 + vbYesNoCancel, "NotePad Clone")
    Case vbYes
    SaveDoc
    OpenDoc
    sv = 0
    Case vbNo
    OpenDoc
    sv = 0
    Case vbCancel
    Exit Sub
End Select
End If
Else
OpenDoc
sv = 0
End If
End Sub
Sub OpenDoc()
   cd1.Filter = "Text Documents (*.txt)|*.txt|All Files (*.*)|*.*"
   cd1.FilterIndex = 2
   cd1.ShowOpen
  Dim LoadFileToTB As Boolean
  Dim TxtBox As Object
  Dim filepath As String
  Dim Append As Boolean
  Dim iFile As Integer
  Dim s As String
    If Dir(filepath) = "" Then Exit Sub
        On Error GoTo ErrorHandler:
        s = rtf.Text
        iFile = FreeFile
            Open cd1.FileName For Input As #iFile
            s = Input(LOF(iFile), #iFile)
                If Append Then
                    rtf.Text = rtf.Text & s
                    Else
                    rtf.Text = s
                End If
    sv = 0
    Me.Caption = cd1.filetitle & " - NotePad Clone"
ErrorHandler:
If iFile > 0 Then Close #iFile
End Sub
Sub SaveDoc()
    Dim Response As VbMsgBoxResult, sFile As String
    If Not FileExists(rtf.Tag) Then
         SaveAs
    Else
        Select Case LCase(ExtOnly(rtf.Tag))
            Case "rtf"
                Response = MsgBox("Any rich text formatting in this file will be lost." + vbCrLf + "Do you wish to save this file using a different name ?", vbYesNoCancel)
                Select Case Response
                    Case vbCancel
                        Exit Sub
                    Case vbYes
                       SaveAs ' GoTo DoSaveAs
                End Select
                rtf.SaveFile rtf.Tag
            Case "doc"
                'same as .rtf
                Response = MsgBox("Any document formatting in this file will be lost." + vbCrLf + "Do you wish to save this file using a different name ?", vbYesNoCancel)
                Select Case Response
                    Case vbCancel
                        Exit Sub
                    Case vbYes
                      SaveAs
                        'GoTo DoSaveAs
                End Select
            Case Else 'just plain text
                Kill rtf.Tag
                FileSave rtf.Text, rtf.Tag
                Me.Caption = cd1.filetitle & " - NotePad Clone"
                Statussave.Text = False
        End Select
    End If
DoSaveAs:
    With cd1
        .Filter = "Text Documents (*.txt)|*.txt|All files (*.*)|*.*"
        .flags = 5 Or 2
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        Select Case .FilterIndex
            Case 1
                If InStr(1, sFile, ".") = 0 Then
                    sFile = sFile + ".txt"
                Else
                    sFile = ChangeExt(sFile, "txt")
                End If
                FileSave rtf.Text, sFile 'Text Document
            Case 2
                If InStr(1, sFile, ".") = 0 Then sFile = sFile + ".txt"
                FileSave rtf.Text, .FileName 'plain text
        End Select
         If .filetitle = "" Then
        Select Case MsgBox("The file you save is not gived a name, do you sure want to continue?", vbQuestion + vbYesNoCancel, "No name")
        Case vbYes
        Me.Caption = "No Name - NotePad Clone"
        Case vbNo
        SaveAs
        Case vbCancel
        Exit Sub
        End Select
        Else
        Me.Caption = .filetitle & " - NotePad Clone"
        sv = 0
        End If
    End With
End Sub

Private Sub mnpagesetup_Click()
    ShowPageSetupDlg
End Sub

Private Sub mnPasteText_Click()
rtf.SelText = Clipboard.GetText()
End Sub

Private Sub mnprintdoc_Click()
 On Error GoTo Errhandler
  Dim BeginPage, EndPage, NumCopies, i
   cd1.CancelError = True
cd1.ShowPrinter
  BeginPage = cd1.FromPage
  EndPage = cd1.ToPage
  NumCopies = cd1.Copies
  For i = 1 To NumCopies
 Printer.Print rtf.Text
  Next i
  Exit Sub
Errhandler:
   Exit Sub
End Sub

Private Sub mnreplace_Click()
On Error Resume Next
'frmReplace.Show
   Dim s As String
  If rtf.SelLength > 0 Then s = rtf.SelText Else s = ""
  ShowFind Me, rtf, 0, s, True, ""
End Sub

Private Sub mnsaveasdoc_Click()
SaveAs
End Sub

Private Sub mnstatusbar_Click()
If mnstatusbar.Checked = True Then
sb1.Visible = False
'RTF Resizer
rtf.Height = Me.Height - 875
mnstatusbar.Checked = False
Else
sb1.Visible = True
mnstatusbar.Checked = True
'RTF Resizer
rtf.Height = rtf.Height - 255
End If
End Sub

Private Sub mnundoText_Click()
 SendMessage rtf.hwnd, EM_UNDO, 0&, 0&
End Sub

Private Sub mnwordWrapNow_Click()
    mnwordWrapNow.Checked = Not mnwordWrapNow.Checked
    rtf.RightMargin = IIf(mnwordWrapNow.Checked, 0, 200000)
End Sub

Private Sub rtf_Change()
sv = 1
End Sub
Private Sub Rtf_KeyUp(KeyCode As Integer, Shift As Integer)
sb1.Panels(1).Text = "| Ln " & Format(GetLineNum(frmMain.rtf) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMain.rtf) + 1)
End Sub


Private Sub Rtf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
sb1.Panels(1).Text = "| Ln " & Format(GetLineNum(frmMain.rtf) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMain.rtf) + 1)
End Sub
Private Sub selectallNow_Click()
rtf.SelStart = 0
rtf.SelLength = Len(rtf.Text)
rtf.SetFocus
End Sub

Private Sub Timer1_Timer()
If XX.Text = "" Then
mnfindNextNow.Enabled = False
Else
mnfindNextNow.Enabled = True
End If
If rtf.Text = "" Then
mnfindNow.Enabled = False
mnreplace.Enabled = False
selectallNow.Enabled = False
mndeleteText.Enabled = False
mnundoText.Enabled = False
mncopyText.Enabled = False
mncutText.Enabled = False
mngotoLine.Enabled = False
Else
mnfindNow.Enabled = True
mnreplace.Enabled = True
selectallNow.Enabled = True
mndeleteText.Enabled = True
mnundoText.Enabled = True
mncutText.Enabled = True
mncopyText.Enabled = True
mngotoLine.Enabled = True
End If
End Sub
