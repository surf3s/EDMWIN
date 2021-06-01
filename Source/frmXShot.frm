VERSION 5.00
Begin VB.Form frmXShot 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "X-Shot"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRecord 
      Caption         =   "X-Shot"
      Default         =   -1  'True
      Height          =   405
      Left            =   5100
      TabIndex        =   13
      Top             =   105
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   405
      Left            =   5100
      TabIndex        =   12
      Top             =   615
      Width           =   1095
   End
   Begin VB.ComboBox txtprism 
      Height          =   315
      Left            =   3000
      TabIndex        =   0
      Text            =   "Select Prism"
      Top             =   690
      Width           =   1485
   End
   Begin VB.Label lblvalue 
      Caption         =   "value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   1740
      TabIndex        =   23
      Top             =   90
      Width           =   2745
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Unit:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   480
      TabIndex        =   22
      Top             =   90
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Y Distance from Station:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   2715
      TabIndex        =   21
      Top             =   1500
      Width           =   2340
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "X Distance from Station:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   2715
      TabIndex        =   20
      Top             =   1215
      Width           =   2340
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Horizontal Angle:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   105
      TabIndex        =   19
      Top             =   1245
      Width           =   1590
   End
   Begin VB.Label lblvalue 
      Alignment       =   1  'Right Justify
      Caption         =   "value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   5100
      TabIndex        =   18
      Top             =   1785
      Width           =   1000
   End
   Begin VB.Label lblvalue 
      Alignment       =   1  'Right Justify
      Caption         =   "value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   5100
      TabIndex        =   17
      Top             =   1500
      Width           =   1000
   End
   Begin VB.Label lblvalue 
      Alignment       =   1  'Right Justify
      Caption         =   "value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   5100
      TabIndex        =   16
      Top             =   1215
      Width           =   1000
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Z Distance from Station:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   2715
      TabIndex        =   15
      Top             =   1785
      Width           =   2340
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Prism"
      Height          =   195
      Left            =   3510
      TabIndex        =   14
      Top             =   480
      Width           =   405
   End
   Begin VB.Label lblvalue 
      Alignment       =   1  'Right Justify
      Caption         =   "value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   1710
      TabIndex        =   11
      Top             =   1815
      Width           =   1000
   End
   Begin VB.Label lblvalue 
      Alignment       =   1  'Right Justify
      Caption         =   "value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   1710
      TabIndex        =   10
      Top             =   1530
      Width           =   1000
   End
   Begin VB.Label lblvalue 
      Alignment       =   1  'Right Justify
      Caption         =   "value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1710
      TabIndex        =   9
      Top             =   1245
      Width           =   1000
   End
   Begin VB.Label lblvalue 
      Alignment       =   1  'Right Justify
      Caption         =   "value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1710
      TabIndex        =   8
      Top             =   960
      Width           =   1000
   End
   Begin VB.Label lblvalue 
      Alignment       =   1  'Right Justify
      Caption         =   "value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1710
      TabIndex        =   7
      Top             =   675
      Width           =   1000
   End
   Begin VB.Label lblvalue 
      Alignment       =   1  'Right Justify
      Caption         =   "value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1710
      TabIndex        =   6
      Top             =   375
      Width           =   1000
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Slope Distance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   105
      TabIndex        =   5
      Top             =   1815
      Width           =   1590
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Vertical Angle:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   105
      TabIndex        =   4
      Top             =   1530
      Width           =   1590
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Z:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   495
      TabIndex        =   3
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   495
      TabIndex        =   2
      Top             =   675
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   495
      TabIndex        =   1
      Top             =   390
      Width           =   1200
   End
End
Attribute VB_Name = "frmXShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()

Unload Me

End Sub

Private Sub cmdRecord_Click()

cmdRecord.Enabled = False
cmdClose.Enabled = False
frmMain.Command1_Click
If txtprism.ListIndex >= 0 Then
    lblvalue(2) = Format(lblvalue(2) + edmshot.poleh - PoleHeight(txtprism.ItemData(txtprism.ListIndex)), "####0.000")
    lblvalue(8) = Format(lblvalue(8) + edmshot.poleh - PoleHeight(txtprism.ItemData(txtprism.ListIndex)), "####0.000")
    edmshot.poleh = PoleHeight(txtprism.ItemData(txtprism.ListIndex))
End If
cmdRecord.Enabled = True
cmdClose.Enabled = True
mdiMain.StatusBar.Panels(6).Visible = False

End Sub

Private Sub Form_Load()

CenterForm Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

XShotShowing = False
Me.Hide
frmMain.Picture1.SetFocus

End Sub

Private Sub txtprism_Click()

If txtprism.ListIndex = -1 Then Exit Sub

If Not Loading Then
    
    lblvalue(2) = Format(lblvalue(2) + edmshot.poleh - PoleHeight(txtprism.ItemData(txtprism.ListIndex)), "####0.000")
    lblvalue(8) = Format(lblvalue(8) + edmshot.poleh - PoleHeight(txtprism.ItemData(txtprism.ListIndex)), "####0.000")
    edmshot.poleh = PoleHeight(txtprism.ItemData(txtprism.ListIndex))
End If

End Sub

Public Sub FindUnit(X As Single, Y As Single)

SqlString = "select * from [EDM_units] where minx< " & X & " and maxx>" & X & " and miny<" & Y & " and maxy>" & Y
Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
If Not rsTemp.EOF Then
    If Not IsNull(rsTemp("unit")) Then
        lblvalue(9) = rsTemp("unit")
        Exit Sub
    Else
        lblvalue(9) = "Not in defined Unit."
    End If
Else
    SqlString = "select * from [EDM_units] where abs(centerx-" & X & ")<=radius and abs(centery-" & Y & ")<=radius"
    Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("unit")) Then
            lblvalue(9) = rsTemp("unit")
            Exit Sub
        Else
            lblvalue(9) = "Not in defined Unit."
        End If
    Else
        lblvalue(9) = "Not in defined Unit."
    End If
End If
Set rsTemp = Nothing

End Sub

Private Sub txtprism_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub


