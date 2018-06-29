VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUndoElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************
'ROTEform1 CLASS MODULE TEXT MANIPULATION CLIENT
'----------------------------------------------
'Modified by : Teztaz Enterprises
'Created     : 10/22/2012
'**********************************************

Public SelStart As Long
Public SelUndoLength As Long
Public SelRedoLength As Long
Public SelUndoText As String
Public SelRedoText As String
Public UndoTabPosition As Long
Public RedoTabPosition As Long
Public Action As Boolean



