VERSION 5.00
Begin VB.Form frmDatummap 
   Caption         =   "Datum Map"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox datummap 
      Height          =   2535
      Left            =   720
      ScaleHeight     =   2475
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "frmDatummap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Call sizecontrols

End Sub
Private Sub sizecontrols()

datummap.width = Me.width - bannerwidth
datummap.Height = Me.Height - bannerheight
datummap.Top = 0
datummap.Left = 0

End Sub

Private Sub Form_Resize()

Call sizecontrols

End Sub

