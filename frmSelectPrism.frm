VERSION 5.00
Begin VB.Form frmSelectPrism 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Prism"
   ClientHeight    =   1410
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   3495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   324
      Left            =   2568
      TabIndex        =   3
      Top             =   720
      Width           =   780
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   324
      Left            =   2568
      TabIndex        =   2
      Top             =   288
      Width           =   780
   End
   Begin VB.ComboBox txtprism 
      Height          =   288
      Left            =   192
      TabIndex        =   0
      Text            =   "Select Prism"
      Top             =   420
      Width           =   1485
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Prism"
      Height          =   192
      Left            =   708
      TabIndex        =   1
      Top             =   204
      Width           =   408
   End
End
Attribute VB_Name = "frmSelectPrism"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

If txtprism.ListIndex < 0 Then
    MsgBox ("Select a prism from the menu")
    Exit Sub
End If
edmshot.poleh = PoleHeight(txtprism.ItemData(txtprism.ListIndex))
edmshot.poleo = PoleOffset(txtprism.ItemData(txtprism.ListIndex))
Unload Me

End Sub

Private Sub Command1_Click()

Cancelling = True
Me.Hide

End Sub

Private Sub Form_Load()

txtprism.Clear
For I = 0 To frmMain.txtprism.ListCount - 1
        txtprism.AddItem frmMain.txtprism.List(I)
        txtprism.ItemData(txtprism.NewIndex) = frmMain.txtprism.ItemData(I)
Next I
Loading = True
If txtprism.ListCount > 0 Then
    txtprism.ListIndex = frmMain.txtprism.ListIndex
End If
Loading = False
CenterForm Me
Cancelling = False

End Sub


