VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUndo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private WithEvents MyRichTextBox As RichTextBox
Attribute MyRichTextBox.VB_VarHelpID = -1
Private TrackUndo As Boolean
Private UndoStack As Collection
Private RedoStack As Collection
Private CatchOnKeyUp As Boolean
Private RichEdit As Boolean
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Public Event StateChanged()
Public Property Get canUndo() As Boolean
    canUndo = UndoStack.count > 0 And TrackUndo
End Property
Public Property Get canRedo() As Boolean
    canRedo = RedoStack.count > 0 And TrackUndo
End Property
Public Property Get UndoCount() As Long
    UndoCount = UndoStack.count
End Property
Public Property Get RedoCount() As Long
    RedoCount = RedoStack.count
End Property
Public Property Get RichBox() As RichTextBox
    Set RichBox = MyRichTextBox
End Property
Public Property Let RichBox(ByVal vNewValue As RichTextBox)
    Set MyRichTextBox = vNewValue
End Property
Public Property Get RichEditing() As Boolean
    RichEditing = RichEdit
End Property
Public Property Let RichEditing(ByVal vNewValue As Boolean)
    RichEdit = vNewValue
End Property
Public Sub Reset()
    Set UndoStack = New Collection
    Set RedoStack = New Collection
    TrackUndo = True
    RaiseEvent StateChanged
End Sub
Public Sub Redo()
    Dim UndoItem As New clsUndoElement
    If RedoStack.count < 1 Then Exit Sub
    LockWindowUpdate frmMain.hwnd
    TrackUndo = False 'dont do anything to the undo system
    'get the data from the collection
    Set UndoItem = RedoStack.Item(RedoStack.count)
    'make use of it
    MyRichTextBox.SelStart = UndoItem.SelStart
    MyRichTextBox.SelLength = UndoItem.SelRedoLength
    If RichEdit Then
        MyRichTextBox.SelRTF = UndoItem.SelRedoText
    Else
        MyRichTextBox.SelText = UndoItem.SelRedoText
    End If
    'if the undo was a tab, then set the start position to the beginning
    'of the text that follows the tab
    If UndoItem.RedoTabPosition <> -1 Then MyRichTextBox.SelStart = UndoItem.RedoTabPosition
    UndoStack.Add UndoItem 'add to the undo
    RedoStack.Remove RedoStack.count 'remove from the redo
    TrackUndo = True 'Ok monitor undo's again
    RaiseEvent StateChanged 'tell the form we're done, so update edit items
    LockWindowUpdate 0
End Sub
Public Sub Undo()
    Dim UndoItem As New clsUndoElement
    If UndoStack.count < 1 Then Exit Sub
    LockWindowUpdate frmMain.hwnd
    TrackUndo = False 'dont do anything to the undo system
    'get the data from the collection
    Set UndoItem = UndoStack.Item(UndoStack.count)
    'make use of it
    MyRichTextBox.SelStart = UndoItem.SelStart
    MyRichTextBox.SelLength = IIf(UndoItem.Action = True, 0, UndoItem.SelUndoLength)
    MyRichTextBox.SelRTF = UndoItem.SelUndoText
    'if the undo was a tab, then set the start position to the beginning
    'of the text that follows the tab
    If UndoItem.UndoTabPosition <> -1 Then MyRichTextBox.SelStart = UndoItem.UndoTabPosition
    RedoStack.Add UndoItem 'add to the redo
    UndoStack.Remove UndoStack.count 'remove from the undo
    TrackUndo = True 'Ok monitor undo's again
    RaiseEvent StateChanged 'tell the form we're done, so update edit items
    LockWindowUpdate 0
End Sub
Public Sub IgnoreChange(Ignore As Boolean)
   TrackUndo = Not Ignore
End Sub

Private Sub AddToStack(mItem As clsUndoElement)
    Dim z As Long, tmpItem As New clsUndoElement
    If Not TrackUndo Then Exit Sub
    If RedoStack.count > 0 Then
        Set RedoStack = New Collection
        UndoStack.Add mItem
    Else
        UndoStack.Add mItem
    End If
End Sub
'These subs make it easier to call and gives the class
'greater control over input

Public Sub Cut()
    MyRichTextBox_KeyDown 88, 2
End Sub

Public Sub Copy()
    SendMessage MyRichTextBox.hwnd, WM_COPY, 0, 0
End Sub

Public Sub Paste()
    MyRichTextBox_KeyDown 86, 2
End Sub
Public Sub InsertText(mText As String, Optional Update As Boolean = True)
    Dim UndoItem As New clsUndoElement 'not used here
    UndoItem.UndoTabPosition = -1
    UndoItem.SelStart = MyRichTextBox.SelStart
    UndoItem.SelUndoLength = Len(mText)
    UndoItem.SelRedoLength = MyRichTextBox.SelLength
    UndoItem.SelUndoText = MyRichTextBox.SelText
    UndoItem.SelRedoText = mText
    MyRichTextBox.SelText = mText
    UndoItem.Action = False
    AddToStack UndoItem
    If Update Then RaiseEvent StateChanged
End Sub

Public Sub Delete()
    MyRichTextBox_KeyDown vbKeyDelete, 0
End Sub
Public Sub UpdateStateChange()
    RaiseEvent StateChanged
End Sub

Private Sub MyRichTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim UndoItem As New clsUndoElement
    Dim curvl As Long, st As Long
    UndoItem.UndoTabPosition = -1
    UndoItem.RedoTabPosition = -1
    LockWindowUpdate frmMain.hwnd
    If Shift = 2 And KeyCode = vbKeyZ Then 'Undo
        Undo
        KeyCode = 0
    ElseIf Shift = 2 And KeyCode = vbKeyR Then 'Redo
        Redo
        KeyCode = 0
    ElseIf KeyCode = vbKeyDelete Then  'delete
        UndoItem.SelStart = MyRichTextBox.SelStart
        UndoItem.SelUndoLength = IIf(MyRichTextBox.SelLength = 0, 1, MyRichTextBox.SelLength)
        UndoItem.SelRedoLength = UndoItem.SelUndoLength
        If MyRichTextBox.SelLength = 0 Then
            MyRichTextBox.SelLength = 2
            If MyRichTextBox.SelText <> vbCrLf Then MyRichTextBox.SelLength = 1
        End If
        UndoItem.SelUndoText = MyRichTextBox.SelRTF
        UndoItem.SelRedoText = ""
        MyRichTextBox.SelRTF = ""
        UndoItem.Action = True
        AddToStack UndoItem
        KeyCode = 0
        RaiseEvent StateChanged
    ElseIf Shift = 2 And KeyCode = vbKeyX Then 'cut
        Clipboard.Clear
        Clipboard.SetText MyRichTextBox.SelText, IIf(RichEdit = True, vbCFRTF, vbCFText)
        UndoItem.SelStart = MyRichTextBox.SelStart
        UndoItem.SelUndoLength = IIf(MyRichTextBox.SelLength = 0, 1, MyRichTextBox.SelLength)
        UndoItem.SelRedoLength = UndoItem.SelUndoLength
        UndoItem.SelUndoText = MyRichTextBox.SelRTF
        UndoItem.SelRedoText = ""
        MyRichTextBox.SelRTF = ""
        UndoItem.Action = True
        AddToStack UndoItem
        KeyCode = 0
        RaiseEvent StateChanged
    ElseIf (Shift = 2 And KeyCode = vbKeyV) Or (Shift = 3 And KeyCode = 45) Then 'paste, Ctrl+Shift+Insert
        UndoItem.SelStart = MyRichTextBox.SelStart
        UndoItem.SelUndoLength = Len(Clipboard.GetText)
        UndoItem.SelRedoLength = MyRichTextBox.SelLength
        UndoItem.SelUndoText = MyRichTextBox.SelText
        UndoItem.SelRedoText = Clipboard.GetText(IIf(RichEdit = True, vbCFRTF, vbCFText))
        If RichEdit Then
            SendMessage MyRichTextBox.hwnd, WM_PASTE, 0, 0
        Else
            MyRichTextBox.SelText = Clipboard.GetText(vbCFText)
        End If
        UndoItem.Action = False
        AddToStack UndoItem
        KeyCode = 0
        RaiseEvent StateChanged
    ElseIf Shift = 0 And KeyCode = 9 Then 'tab
        LockWindowUpdate frmMain.hwnd
        UndoItem.UndoTabPosition = MyRichTextBox.SelStart
        MyRichTextBox.Span vbCrLf, True, True
        UndoItem.SelStart = MyRichTextBox.SelStart
        UndoItem.SelUndoLength = MyRichTextBox.SelLength
        UndoItem.SelUndoText = MyRichTextBox.SelText
        MyRichTextBox.SelStart = UndoItem.UndoTabPosition
        MyRichTextBox.SelLength = 0
        UndoItem.Action = False
        AddToStack UndoItem
        CatchOnKeyUp = True
        Exit Sub
    ElseIf Shift = 2 And KeyCode = vbKeyC Then 'copy
        Copy
        KeyCode = 0
    ElseIf Shift = 2 And KeyCode = vbKeyA Then 'select all
        MyRichTextBox.SelStart = 0
        MyRichTextBox.SelLength = Len(MyRichTextBox.Text)
        MyRichTextBox.SetFocus
        KeyCode = 0
    ElseIf Shift = 2 And KeyCode = vbKeyT Then 'select above
        Screen.MousePointer = 11
        LockWindowUpdate frmMain.hwnd 'dont paint till we're finished
        'this is where the top line and therefore scroll position is
        curvl = SendMessage(MyRichTextBox.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
        'do the selection
        st = MyRichTextBox.SelStart
        MyRichTextBox.SelStart = 0
        MyRichTextBox.SelLength = st
        'return the scroll postion back to what it was - see SetScrollPos sub
        SetScrollPos curvl, MyRichTextBox
        MyRichTextBox.SetFocus
        Screen.MousePointer = 0
        LockWindowUpdate 0
        KeyCode = 0
    ElseIf Shift = 2 And KeyCode = vbKeyB Then 'select below
        Screen.MousePointer = 11
        LockWindowUpdate frmMain.hwnd
        curvl = SendMessage(MyRichTextBox.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
        st = MyRichTextBox.SelStart + MyRichTextBox.SelLength
        MyRichTextBox.SelStart = st
        MyRichTextBox.SelLength = Len(MyRichTextBox.Text) - st
        SetScrollPos curvl, MyRichTextBox
        MyRichTextBox.SetFocus
        Screen.MousePointer = 0
        LockWindowUpdate 0
        KeyCode = 0
    ElseIf Shift < 2 Then 'any key
        UndoItem.SelStart = MyRichTextBox.SelStart
        UndoItem.SelUndoLength = MyRichTextBox.SelLength
        UndoItem.SelRedoLength = MyRichTextBox.SelLength
        UndoItem.SelUndoText = IIf(UndoItem.SelUndoLength = 0, "", MyRichTextBox.SelRTF)
        UndoItem.SelUndoLength = IIf(KeyCode = vbKeyReturn, 2, 1)
        Select Case KeyCode
            Case vbKeyReturn
                UndoItem.SelRedoText = vbCrLf
            Case vbKeyNumpad0
                UndoItem.SelRedoText = "0"
            Case vbKeyNumpad1
                UndoItem.SelRedoText = "1"
            Case vbKeyNumpad2
                UndoItem.SelRedoText = "2"
            Case vbKeyNumpad3
                UndoItem.SelRedoText = "3"
            Case vbKeyNumpad4
                UndoItem.SelRedoText = "4"
            Case vbKeyNumpad5
                UndoItem.SelRedoText = "5"
            Case vbKeyNumpad6
                UndoItem.SelRedoText = "6"
            Case vbKeyNumpad7
                UndoItem.SelRedoText = "7"
            Case vbKeyNumpad8
                UndoItem.SelRedoText = "8"
            Case vbKeyNumpad9
                UndoItem.SelRedoText = "9"
            Case Else
                UndoItem.SelRedoText = IIf(Shift = 0, LCase(Chr(KeyCode)), UCase(Chr(KeyCode)))
        End Select
        UndoItem.Action = False
        AddToStack UndoItem
        RaiseEvent StateChanged
    ElseIf Shift <> 4 Then 'ignore other combinations of keys
        KeyCode = 0
    End If
    LockWindowUpdate 0
End Sub
Private Sub MyRichTextBox_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim UndoItem As clsUndoElement, st As Long, st1 As Long, sL As Long
    'this is to catch the redo data for a tab
    'can be tricky to achieve without mentioning .text or .textRTF
    If CatchOnKeyUp Then
        CatchOnKeyUp = False
        Set UndoItem = UndoStack.Item(UndoStack.count)
        st = MyRichTextBox.SelStart
        MyRichTextBox.SelStart = st - 2
        MyRichTextBox.Span vbCrLf, True, True
        If MyRichTextBox.SelText = "" Then 'Doh! tab occurred just after a vbcrlf
            st1 = MyRichTextBox.SelStart + 2
            MyRichTextBox.SelStart = st
            MyRichTextBox.Span vbCrLf, True, True
            sL = MyRichTextBox.SelLength
            MyRichTextBox.SelStart = st1
            MyRichTextBox.SelLength = st - st1 + sL - 2
            UndoItem.SelRedoText = MyRichTextBox.SelText 'add data to class
            UndoItem.RedoTabPosition = st 'add data to class
        Else ' tab occurred mid-line
            UndoItem.SelRedoText = MyRichTextBox.SelText 'add data to class
            UndoItem.RedoTabPosition = st + 1 'add data to class
        End If
        UndoItem.SelRedoLength = MyRichTextBox.SelLength 'add data to class
        'set the start position to the beginning of the text that follows the tab
        MyRichTextBox.SelStart = st
        MyRichTextBox.SelLength = 0
        LockWindowUpdate 0
        RaiseEvent StateChanged
    End If
End Sub




