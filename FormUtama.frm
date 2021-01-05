VERSION 5.00
Begin VB.Form FormUtama 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3d Label dan TextBox"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FormUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    PaintForm3D Me
    PaintControl3D Me, Label1  'Label1 is name of label
    PaintControl3D Me, Text1   'Text1 is name of textbox
End Sub
