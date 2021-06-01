VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form frmFields 
   Caption         =   "Fields"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin MSGrid.Grid dbfields 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   2355
      _StockProps     =   77
      BackColor       =   16777215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
End
Attribute VB_Name = "frmFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Form_Load()

dbfields.Cols = 3
dbfields.Rows = 20
dbfields.Row = 0
dbfields.Col = 0
dbfields.Text = "Field"
dbfields.Col = 1
dbfields.Text = "Type"
dbfields.Col = 2
dbfields.Text = "Options"
dbfields.FixedCols = 0
'dbfields.

End Sub
