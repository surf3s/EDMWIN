VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Menu"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox menulist 
      Height          =   1425
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox menuitem 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label menutitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a points file:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

Select Case Index
Case 0
    MenuSelection$ = menuitem.Text
Case 1
    MenuSelection$ = ""
Case Else
End Select

Unload Me

End Sub

Private Sub menulist_Click()

menuitem.Text = menulist

End Sub
