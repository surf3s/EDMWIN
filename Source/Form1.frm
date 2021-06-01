VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2184
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6768
   LinkTopic       =   "Form1"
   ScaleHeight     =   2184
   ScaleWidth      =   6768
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Show Final Points"
      Height          =   540
      Left            =   2928
      TabIndex        =   7
      Top             =   288
      Width           =   1884
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Combo"
      Height          =   252
      Index           =   3
      Left            =   192
      TabIndex        =   6
      Top             =   1392
      Width           =   1068
   End
   Begin VB.OptionButton Option1 
      Caption         =   "YZ"
      Height          =   252
      Index           =   1
      Left            =   192
      TabIndex        =   5
      Top             =   1056
      Width           =   1068
   End
   Begin VB.OptionButton Option1 
      Caption         =   "XZ"
      Height          =   252
      Index           =   2
      Left            =   192
      TabIndex        =   4
      Top             =   720
      Width           =   1068
   End
   Begin VB.OptionButton Option1 
      Caption         =   "XY"
      Height          =   252
      Index           =   0
      Left            =   192
      TabIndex        =   3
      Top             =   384
      Value           =   -1  'True
      Width           =   1068
   End
   Begin VB.CommandButton Command3 
      Caption         =   "XZ Rotate"
      Height          =   252
      Left            =   3312
      TabIndex        =   2
      Top             =   1056
      Width           =   1404
   End
   Begin VB.CommandButton Command2 
      Caption         =   "YZ Rotate"
      Height          =   252
      Left            =   4992
      TabIndex        =   1
      Top             =   1056
      Width           =   1404
   End
   Begin VB.CommandButton Command1 
      Caption         =   "XY Rotate"
      Height          =   252
      Left            =   1632
      TabIndex        =   0
      Top             =   1056
      Width           =   1404
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmGetSlopeXY.Show

End Sub


Private Sub Command2_Click()
frmGetSlopeYZ.Show
End Sub


Private Sub Command3_Click()
frmGetSlopeXZ.Show

End Sub


Private Sub Command4_Click()
frmFinal.Show
End Sub

Private Sub Form_Load()
Option1_Click 0
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
    Case 0
        DatumFile = "testdatums-xy.txt"
        PointFile = "testpoints-xy.txt"
        TitleString = "Using XY Rotation"
    Case 1
        DatumFile = "testdatums-yz.txt"
        PointFile = "testpoints-yz.txt"
        TitleString = "Using YZ Rotation"
    Case 2
        DatumFile = "testdatums-xz.txt"
        PointFile = "testpoints-xz.txt"
        TitleString = "Using XZ Rotation"
    Case 3
        DatumFile = "testdatums-combo.txt"
        PointFile = "testpoints-combo.txt"
        TitleString = "Using Combo"
End Select

End Sub


