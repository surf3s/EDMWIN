VERSION 5.00
Begin VB.Form frmRecordDatum 
   Caption         =   "Record New Datum"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPoleHT 
      Enabled         =   0   'False
      Height          =   285
      Left            =   780
      TabIndex        =   11
      Top             =   2370
      Width           =   885
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   465
      Left            =   4620
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtZ 
      Height          =   285
      Left            =   600
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1620
      Width           =   1185
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   600
      TabIndex        =   8
      Top             =   1260
      Width           =   1185
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Top             =   900
      Width           =   1185
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1140
      TabIndex        =   6
      Top             =   90
      Width           =   1875
   End
   Begin VB.ComboBox txtPrism 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   2010
      Width           =   1785
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Record"
      Height          =   465
      Left            =   4620
      TabIndex        =   0
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label lblPrism 
      Caption         =   "Prism: "
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   525
   End
   Begin VB.Label lblPoleHT 
      Caption         =   "Height"
      Height          =   255
      Left            =   210
      TabIndex        =   12
      Top             =   2370
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Datum Name:"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Z"
      Height          =   285
      Index           =   2
      Left            =   270
      TabIndex        =   4
      Top             =   1650
      Width           =   285
   End
   Begin VB.Label Label1 
      Caption         =   "Y"
      Height          =   285
      Index           =   1
      Left            =   270
      TabIndex        =   3
      Top             =   1290
      Width           =   285
   End
   Begin VB.Label Label1 
      Caption         =   "X"
      Height          =   285
      Index           =   0
      Left            =   270
      TabIndex        =   2
      Top             =   960
      Width           =   285
   End
End
Attribute VB_Name = "frmRecordDatum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
X = Int((3) * Rnd) + 998 + Rnd
y = Int((1) * Rnd) + 1012 + Rnd
z = Rnd
txtX = Format(X, "######0.000")
txtY = Format(y, "######0.000")
txtZ = Format(z, "######0.000")

End Sub

Private Sub Form_Load()
PoleTB.MoveFirst
i = 0
txtPrism.Clear
While Not PoleTB.EOF
    i = i + 1
    txtPrism.AddItem PoleTB("pole")
    txtPrism.ItemData(txtPrism.NewIndex) = i
    PoleTB.MoveNext
Wend
Loading = True
txtPrism.ListIndex = 0
Loading = False
End Sub

Private Sub Text1_Change()

End Sub



Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DatumTB.Index = "datumname"
    DatumTB.Seek "=", txtName
    If Not DatumTB.NoMatch Then
        response = MsgBox(txtName + " already exists.  Replace?", vbYesNo)
        If response = vbYes Then
            DatumTB.Delete
        Else
            txtName.SetFocus
            txtName.SelStart = 0
            txtName.SelLength = Len(txtName)
        End If
    End If
End If
End Sub


Private Sub txtName_LostFocus()
txtName_KeyPress 13

End Sub


Private Sub txtPrism_Click()
If Not Loading Then
    If PoleHeight(txtPrism.ItemData(txtPrism.ListIndex)) <> OriginalPoleHT Then
        response = MsgBox("Update Z value from " & txtZ & " to " & Format(Val(txtZ) + OriginalPoleHT - PoleHeight(txtPrism.ItemData(txtPrism.ListIndex)), "#####0.000") & "?", vbYesNo)
        If response = vbYes Then
            txtPoleHT = PoleHeight(txtPrism.ItemData(txtPrism.ListIndex))
            txtZ = Val(txtZ) + OriginalPoleHT - PoleHeight(txtPrism.ItemData(txtPrism.ListIndex))
        Else
            txtPrism.ListIndex = OriginalPrismIndex
        End If
    End If
    txtPoleHT = PoleHeight(txtPrism.ItemData(txtPrism.ListIndex))
    OriginalPoleHT = txtPoleHT
    OriginalPrismIndex = txtPrism.ListIndex
Else
    txtPoleHT = PoleHeight(1)
End If

End Sub

Private Sub txtPrism_GotFocus()
OriginalPoleHT = txtPoleHT
OriginalPrismIndex = txtPrism.ListIndex
txtPrism.SelStart = 0
txtPrism.SelLength = Len(txtPrism)

End Sub

Private Sub txtPrism_LostFocus()
txtPrism.SelLength = 0
txtPoleHT = OriginalPoleHT
txtPrism.ListIndex = OriginalPrismIndex

End Sub


