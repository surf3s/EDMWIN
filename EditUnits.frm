VERSION 5.00
Begin VB.Form AddUnits 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Units"
   ClientHeight    =   4890
   ClientLeft      =   4125
   ClientTop       =   3030
   ClientWidth     =   4440
   Icon            =   "EditUnits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "OK"
      Height          =   495
      Left            =   3210
      TabIndex        =   23
      Top             =   3810
      Width           =   675
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   21
      Top             =   2550
      Width           =   2445
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Lookup"
      Height          =   375
      Left            =   1650
      TabIndex        =   20
      Top             =   1530
      Width           =   855
   End
   Begin VB.TextBox txtLastID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   810
      TabIndex        =   19
      Text            =   "0"
      Top             =   1560
      Width           =   705
   End
   Begin VB.ComboBox txtUnit 
      Height          =   315
      Left            =   1050
      Sorted          =   -1  'True
      TabIndex        =   17
      Text            =   "txtUnit"
      Top             =   150
      Width           =   1695
   End
   Begin VB.TextBox XY 
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   630
      Width           =   555
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2244
      TabIndex        =   4
      Text            =   "1"
      Top             =   630
      Width           =   555
   End
   Begin VB.TextBox XY 
      Height          =   285
      Index           =   1
      Left            =   1260
      TabIndex        =   1
      Top             =   630
      Width           =   555
   End
   Begin VB.TextBox XY 
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox XY 
      Height          =   285
      Index           =   3
      Left            =   1260
      TabIndex        =   3
      Top             =   990
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   330
      Left            =   3150
      TabIndex        =   6
      Top             =   48
      Width           =   825
   End
   Begin VB.TextBox txtZ 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2244
      TabIndex        =   5
      Text            =   "0"
      Top             =   990
      Width           =   555
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   330
      Left            =   3150
      TabIndex        =   7
      Top             =   468
      Width           =   825
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Select tables that might contain the last ID, then click OK."
      Height          =   885
      Left            =   2790
      TabIndex        =   22
      Top             =   2580
      Width           =   1485
   End
   Begin VB.Label Label8 
      Caption         =   "Last ID:"
      Height          =   255
      Left            =   60
      TabIndex        =   18
      Top             =   1590
      Width           =   615
   End
   Begin VB.Label X 
      Caption         =   "X2, Y2"
      Height          =   240
      Index           =   1
      Left            =   3750
      TabIndex        =   16
      Top             =   1110
      Width           =   630
   End
   Begin VB.Label X 
      Caption         =   "X1, Y1"
      Height          =   210
      Index           =   0
      Left            =   3030
      TabIndex        =   15
      Top             =   2070
      Width           =   600
   End
   Begin VB.Shape Shape1 
      Height          =   636
      Left            =   3312
      Top             =   1392
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Unit Name:"
      Height          =   195
      Left            =   150
      TabIndex        =   14
      Top             =   180
      Width           =   825
   End
   Begin VB.Label Label2 
      Caption         =   "X1"
      Height          =   240
      Left            =   90
      TabIndex        =   13
      Top             =   660
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Size:"
      Height          =   195
      Left            =   1890
      TabIndex        =   12
      Top             =   660
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Y1"
      Height          =   240
      Left            =   990
      TabIndex        =   11
      Top             =   660
      Width           =   240
   End
   Begin VB.Label Label5 
      Caption         =   "X2"
      Height          =   240
      Left            =   90
      TabIndex        =   10
      Top             =   1020
      Width           =   240
   End
   Begin VB.Label Label6 
      Caption         =   "Y2"
      Height          =   240
      Left            =   990
      TabIndex        =   9
      Top             =   1020
      Width           =   240
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Z"
      Height          =   240
      Left            =   1935
      TabIndex        =   8
      Top             =   1020
      Width           =   240
   End
End
Attribute VB_Name = "AddUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Combo1_Change()
End Sub


Private Sub Command1_Click()

If txtUnit = "" Then
    MsgBox ("Input Unit Name.")
    txtUnit.SetFocus
    Exit Sub
End If
txtUnit = Trim(txtUnit)
If Len(txtUnit) > UnitLength Then
    MsgBox ("Unit Name too long")
    txtUnit.SetFocus
    Exit Sub
End If
If XY(1) = "" Or XY(2) = "" Or XY(3) = "" Or XY(0) = "" Then
    MsgBox ("All coordinates must be entered")
    Exit Sub
End If
If Val(XY(2)) <= Val(XY(0)) Then
    MsgBox ("Invalid range:  X1 corner is lower left boundary in plan view.")
    Exit Sub
End If
If Val(XY(3)) <= Val(XY(1)) Then
    MsgBox ("Invalid range:  Y1 corner is lower left boundary in plan view.")
    Exit Sub
End If
txtUnit = UCase(txtUnit)
UnitTB.Index = "unitname"
UnitTB.Seek "=", txtUnit
If Not UnitTB.NoMatch Then
    response = MsgBox("Unit already defined -- overwrite?", vbYesNo)
    If response = vbNo Then
        txtUnit = ""
        Exit Sub
    End If
    Do While Not UnitTB.NoMatch
        UnitTB.Delete
        UnitTB.Seek "=", txtUnit
    Loop
End If


UnitTB.AddNew
UnitTB("UNIT") = txtUnit
UnitTB("ID") = PadID(txtLastID)
UnitTB("suffix") = 0
UnitTB("minX") = XY(0)
UnitTB("miny") = XY(1)
UnitTB("maxX") = XY(2)
UnitTB("maxy") = XY(1)
UnitTB.Update

For i = 0 To 3
    XY(i) = ""
Next i
txtUnit.AddItem txtUnit
txtUnit = ""
txtLastID = 0

txtUnit.SetFocus

End Sub

Public Sub Command2_Click()
UnitTB.Index = "unitname"
UnitTB.Seek "=", txtUnit
If Not UnitTB.NoMatch Then
    response = MsgBox("Permanently remove unit " + txtUnit + "?", vbYesNo)
    If response = vbNo Then
        Exit Sub
    End If
    UnitTB.Delete
    txtUnit.RemoveItem (txtUnit.ListIndex)
    
    txtUnit = ""
    txtLastID = 0
    txtUnit.SetFocus
Else
    MsgBox ("Unit not found.")
End If
End Sub

Private Sub Command3_Click()

Me.Height = 5295
GetMainTables
List1.Clear
For i = 1 To nMainTables
    List1.AddItem MainTable(i)
Next i

End Sub

Private Sub Command4_Click()
Dim RsTemp As Recordset
Dim MaxID As Long

If txtUnit = "" Then
    MsgBox ("Select or enter a unit name first.")
    Exit Sub
End If
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) = True Then
        SqlString = "select max(id) from " + List1.List(i) + " where id<'A' and unit='" + txtUnit + "'"
        Set RsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
        If Not RsTemp.EOF Then
            If Not IsNull(RsTemp(0)) Then
                If Val(Trim(RsTemp(0))) > MaxID Then
                    MaxID = Val(Trim(RsTemp(0)))
                End If
            End If
        End If
    End If
Next i
txtLastID = MaxID
End Sub


Private Sub Form_Load()
Me.Height = 2775

txtUnit.Clear
UnitTB.MoveFirst
While Not UnitTB.EOF
    txtUnit.AddItem UnitTB("Unit")
    UnitTB.MoveNext
Wend



End Sub


Private Sub Text1_Change()

End Sub


Private Sub Option1_Click(Index As Integer)
If Index = 0 Then
    Command2.Enabled = False
    Label1.Visible = False
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    For i = 0 To 3
        XY(i).Visible = False
    Next i
    txtUnit.Visible = False
    txtSize.Visible = False
Else
    Command2.Enabled = True
    Label1.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    For i = 0 To 3
        XY(i).Visible = True
    Next i
    txtUnit.Visible = True
    txtSize.Visible = True
End If
    
End Sub

Private Sub X1_Change()
x2 = Val(x1) + Val(txtSize)

End Sub


Private Sub X1_GotFocus()
x1.SelStart = 0
x1.SelLength = Len(x1)
End Sub


Private Sub x2_GotFocus()
x2.SelStart = 0
x2.SelLength = Len(x2)

End Sub


Private Sub y1_Change()
y2 = Val(y1) + Val(txtSize)

End Sub


Private Sub y1_GotFocus()
y1.SelStart = 0
y1.SelLength = Len(y1)

End Sub


Private Sub Y2_GotFocus()
y2.SelStart = 0
y2.SelLength = Len(y2)

End Sub


Private Sub XY_Change(Index As Integer)
If Index < 2 Then
    XY(Index + 2) = Val(XY(Index)) + Val(txtSize)
End If
End Sub


