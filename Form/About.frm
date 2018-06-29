VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "OK"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6480
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "http:\\thedarkenteztaz.blogspot.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Visit Me : "
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version x"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NotePad Clone"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   5400
      Picture         =   "About.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":08CA
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   6255
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'      NotePad Clone Form : About
'----------------------------------------------
'Created     : 10/22/2012
'Re-Modified : 09/06/2014
'**********************************************

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label2.Caption = "Version : " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
