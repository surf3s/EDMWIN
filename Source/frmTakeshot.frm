VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTakeshot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recording point..."
   ClientHeight    =   1080
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar progress 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   972
   End
End
Attribute VB_Name = "frmTakeshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Form_Load()

progress.Left = (Me.Width) / 2 - progress.Width / 2
Command1.Left = (Me.Width) / 2 - Command1.Width / 2

Screen.MousePointer = 1

End Sub

