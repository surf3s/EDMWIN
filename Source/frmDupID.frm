VERSION 5.00
Begin VB.Form frmDupID 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Duplicate ID options"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   348
      Left            =   4548
      TabIndex        =   4
      Top             =   684
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   348
      Left            =   4524
      TabIndex        =   3
      Top             =   180
      Width           =   780
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Convert to alpha ID"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   1572
      Width           =   3615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "New ID of "
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   1206
      Width           =   3615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Contuation (+ Shot) of "
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   840
      Value           =   -1  'True
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Duplicate ID.  Choose which option to follow:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   120
      TabIndex        =   5
      Top             =   108
      Width           =   4236
   End
End
Attribute VB_Name = "frmDupID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

For I = 0 To 2
    If Option1(I) Then
        frmMain.DupOption = I
        Exit For
    End If
Next I
Unload Me

End Sub

Private Sub Command2_Click()

Cancelling = True
Unload Me

End Sub

Private Sub Form_Load()

Cancelling = False
CenterForm Me

End Sub


