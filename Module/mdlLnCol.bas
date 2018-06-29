Attribute VB_Name = "mdlLnCol"
'**********************************************
'       ROTEJx MODULE TEXT COLLISION
'----------------------------------------------
'Modified by : Teztaz Enterprises
'Created     : 10/22/2012
'**********************************************

Option Explicit
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" _
       (ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long
Private Const EM_LINEFROMCHAR As Long = &HC9
Private Const EM_LINEINDEX As Long = &HBB
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_LINESCROLL As Long = &HB6
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)
Public There As Boolean
Public ret As Long
Public TXFileName As String
Public PasteKey As Boolean
Public MyMsg As String
Public CurColumn As Long
Public ToolbarOn As Boolean
Public FormatbarOn As Boolean
Public StatusbarOn As Boolean
Public RulerOn As Boolean
Public WrapOn As Integer
Public i As Integer
Public fCancel As Boolean
Public TXMeasurement As Integer
Public LineSpacing As Long
Public ColorHighLite As Long
Public HighColor As Boolean
Public tRow As Integer
Public tCol As Integer
Public tWidth As Long
Public tColWidth(1 To 20)
Public tCenter As Boolean
Public Function GetColPos(tBox As Object) As Long
  GetColPos = tBox.SelStart - SendMessageByNum(tBox.hwnd, EM_LINEINDEX, -1&, 0&)
End Function
Public Function GetLineNum(tBox As Object) As Long
  GetLineNum = SendMessageByNum(tBox.hwnd, EM_LINEFROMCHAR, tBox.SelStart, 0&)
End Function
Public Function LastPart(Text As String) As String
 Dim Temp As String
 Dim i As Integer
 Temp = Trim$(Text)
 For i = Len(Temp) To 1 Step -1
  If Mid$(Temp, i, 1) = "\" Then Exit For
 Next i
 If i = 0 Then
  LastPart = Temp
 Else
  LastPart = Mid$(Temp, i + 1)
 End If
End Function
Public Function FirstPart(Text As String) As String
 Dim Temp As String
 Dim i As Integer
 Temp = Trim$(Text)
 For i = Len(Temp) To 1 Step -1
  If Mid$(Temp, i, 1) = "\" Then Exit For
 Next i
 If i = 0 Then
  FirstPart = Temp
 Else
  FirstPart = Left$(Temp, i - 1)
 End If
End Function

Public Function FileExistsX(filename As String) As Boolean
On Error GoTo handle
    If FileLen(filename) >= 0 Then: FileExistsX = True: Exit Function
handle:
    FileExistsX = False
End Function
Public Function WordCount(Text As String) As Long
    Dim dest() As Byte
    Dim i As Long
    If LenB(Text) Then
        ReDim dest(LenB(Text))
        CopyMemory dest(0), ByVal StrPtr(Text), LenB(Text) - 1
        For i = 0 To UBound(dest) Step 2
            If dest(i) > 32 Then
                Do Until dest(i) < 33
                    i = i + 2
                Loop
                WordCount = WordCount + 1
            End If
        Next i
        Erase dest
    Else
        WordCount = 0
    End If
End Function


