VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "A Label"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Call AdjustForm(Me)
End Sub

Private Sub Form_Resize()
Call FormResize(Me)
End Sub
