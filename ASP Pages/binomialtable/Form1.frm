VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim prob As Double
  Dim lngX As Long
  Dim lngN As Long
  Dim dblP As Double
  
  lngX = CLng(Text1)
  lngN = CLng(Text2)
  dblP = CDbl(Text3)
  
  prob = CumBinProb(lngN, lngX, dblP)
  Text4 = Round(prob, 3)
  
End Sub
