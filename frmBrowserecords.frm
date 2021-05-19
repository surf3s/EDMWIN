VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmBrowserecords 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit/Create record"
   ClientHeight    =   5730
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   372
      Index           =   1
      Left            =   2280
      TabIndex        =   2
      Top             =   5160
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   372
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   5040
      Width           =   1092
   End
   Begin MSDBGrid.DBGrid recorddata 
      Height          =   3852
      Left            =   120
      OleObjectBlob   =   "frmBrowserecords.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   5052
   End
End
Attribute VB_Name = "frmBrowserecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

Unload Me

End Sub

