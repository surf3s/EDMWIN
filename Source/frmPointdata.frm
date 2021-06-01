VERSION 5.00
Begin VB.Form frmPointdata 
   Caption         =   "Point Data"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox poleh 
      Alignment       =   1  'Right Justify
      Height          =   372
      Left            =   1560
      TabIndex        =   6
      Top             =   3000
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   372
      Index           =   1
      Left            =   1920
      TabIndex        =   8
      Top             =   3480
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Accept"
      Default         =   -1  'True
      Height          =   372
      Index           =   0
      Left            =   840
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox x 
      Alignment       =   1  'Right Justify
      Height          =   372
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   972
   End
   Begin VB.TextBox y 
      Alignment       =   1  'Right Justify
      Height          =   372
      Left            =   1560
      TabIndex        =   4
      Top             =   2040
      Width           =   972
   End
   Begin VB.TextBox z 
      Alignment       =   1  'Right Justify
      Height          =   372
      Left            =   1560
      TabIndex        =   5
      Top             =   2520
      Width           =   972
   End
   Begin VB.TextBox hangle 
      Alignment       =   1  'Right Justify
      Height          =   372
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   972
   End
   Begin VB.TextBox vangle 
      Alignment       =   1  'Right Justify
      Height          =   372
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   972
   End
   Begin VB.TextBox sloped 
      Alignment       =   1  'Right Justify
      Height          =   372
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pole Height :"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label instruct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "meters"
      Height          =   195
      Index           =   6
      Left            =   2640
      TabIndex        =   21
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X :"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Y :"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Z :"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label instruct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "meters"
      Height          =   195
      Index           =   5
      Left            =   2640
      TabIndex        =   17
      Top             =   1680
      Width           =   465
   End
   Begin VB.Label instruct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "meters"
      Height          =   195
      Index           =   4
      Left            =   2640
      TabIndex        =   16
      Top             =   2160
      Width           =   465
   End
   Begin VB.Label instruct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "meters"
      Height          =   195
      Index           =   3
      Left            =   2640
      TabIndex        =   15
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Horizontal angle :"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vertical angle :"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Slope distance :"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label instruct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ddd.mmss"
      Height          =   195
      Index           =   0
      Left            =   2640
      TabIndex        =   11
      Top             =   240
      Width           =   750
   End
   Begin VB.Label instruct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ddd.mmss"
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   10
      Top             =   720
      Width           =   750
   End
   Begin VB.Label instruct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "meters"
      Height          =   195
      Index           =   2
      Left            =   2640
      TabIndex        =   9
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frmPointdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)

Unload Me

End Sub

Private Sub Form_Load()

Me.Width = 3765
Me.Height = 4350

End Sub
