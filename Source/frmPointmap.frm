VERSION 5.00
Begin VB.Form frmPointmap 
   Caption         =   "Point Map"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   5460
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pointmap 
      BackColor       =   &H00000000&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "frmPointmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Call sizecontrols

End Sub

Private Sub sizecontrols()

pointmap.width = Me.width - bannerwidth
pointmap.Height = Me.Height - bannerheight
pointmap.Top = 0
pointmap.Left = 0

End Sub

Private Sub Form_Resize()

Call sizecontrols

End Sub
