VERSION 5.00
Begin VB.Form AddField 
   Caption         =   "Add Field"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   Icon            =   "AddField.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Optional Fields"
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   7335
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3360
         TabIndex        =   15
         Top             =   1680
         Width           =   975
      End
      Begin VB.ListBox List2 
         Height          =   1815
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3360
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Carry Values to New Shots"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   1320
         Width           =   3135
      End
      Begin VB.OptionButton optType 
         Caption         =   "Numeric"
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optType 
         Caption         =   "Text"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   9
         Top             =   960
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optType 
         Caption         =   "Menu"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   5640
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Menu Values "
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Type:"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Maximum Length:"
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Field Name:"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Default Fields"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "AddField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()

End Sub

Private Sub List1_Click()

End Sub


