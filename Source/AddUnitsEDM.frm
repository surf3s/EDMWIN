VERSION 5.00
Begin VB.Form AddUnits 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Units"
   ClientHeight    =   6135
   ClientLeft      =   4125
   ClientTop       =   3045
   ClientWidth     =   9885
   Icon            =   "AddUnitsEDM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9885
   Begin VB.Frame Frame1 
      Caption         =   "Unit Fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   60
      TabIndex        =   42
      Top             =   3900
      Width           =   9765
      Begin VB.ListBox lstUnitFields 
         Height          =   1410
         Left            =   150
         Style           =   1  'Checkbox
         TabIndex        =   43
         Top             =   360
         Width           =   3435
      End
      Begin VB.Label Label12 
         Caption         =   "To select a Unit Field, check the box next to the field name.  Uncheck the box to remove the field from the list of Unit Fields."
         Height          =   405
         Left            =   3720
         TabIndex        =   45
         Top             =   1320
         Width           =   5955
      End
      Begin VB.Label Label10 
         Caption         =   $"AddUnitsEDM.frx":08CA
         Height          =   825
         Left            =   3720
         TabIndex        =   44
         Top             =   360
         Width           =   5955
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Unit Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   96
      TabIndex        =   39
      Top             =   96
      Width           =   2772
      Begin VB.CommandButton Command1 
         Caption         =   "Clr"
         Height          =   240
         Left            =   2130
         TabIndex        =   41
         Top             =   345
         Width           =   525
      End
      Begin VB.ComboBox txtUnit 
         Height          =   315
         Left            =   195
         Sorted          =   -1  'True
         TabIndex        =   0
         Text            =   "txtUnit"
         Top             =   312
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Shape"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   2976
      TabIndex        =   38
      Top             =   96
      Width           =   2340
      Begin VB.OptionButton optType 
         Caption         =   "No Limits"
         Height          =   195
         Index           =   2
         Left            =   648
         TabIndex        =   3
         Top             =   528
         Width           =   1020
      End
      Begin VB.OptionButton optType 
         Caption         =   "Rectangle"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   264
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optType 
         Caption         =   "Circle"
         Height          =   195
         Index           =   1
         Left            =   1248
         TabIndex        =   2
         Top             =   264
         Width           =   855
      End
   End
   Begin VB.Frame TypeFrame 
      Caption         =   "Coordinates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1344
      Index           =   0
      Left            =   105
      TabIndex        =   18
      Top             =   915
      Width           =   6315
      Begin VB.TextBox XY 
         Height          =   285
         Index           =   3
         Left            =   1485
         TabIndex        =   20
         Top             =   615
         Width           =   825
      End
      Begin VB.TextBox XY 
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   19
         Top             =   615
         Width           =   825
      End
      Begin VB.TextBox XY 
         Height          =   228
         Index           =   1
         Left            =   1470
         TabIndex        =   5
         Top             =   255
         Width           =   825
      End
      Begin VB.TextBox txtSize 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   936
         TabIndex        =   6
         Text            =   "1"
         Top             =   924
         Width           =   375
      End
      Begin VB.TextBox XY 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   255
         Width           =   825
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Grid North"
         Height          =   192
         Left            =   5400
         TabIndex        =   29
         Top             =   1092
         Width           =   720
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   5916
         X2              =   5796
         Y1              =   576
         Y2              =   396
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   5676
         X2              =   5796
         Y1              =   606
         Y2              =   396
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   5796
         X2              =   5796
         Y1              =   996
         Y2              =   456
      End
      Begin VB.Label Label6 
         Caption         =   "Y2"
         Height          =   240
         Left            =   1215
         TabIndex        =   28
         Top             =   645
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "X2"
         Height          =   240
         Left            =   90
         TabIndex        =   27
         Top             =   645
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Y1"
         Height          =   240
         Left            =   1215
         TabIndex        =   26
         Top             =   285
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Size:"
         Height          =   252
         Left            =   588
         TabIndex        =   25
         Top             =   960
         Width           =   348
      End
      Begin VB.Label Label2 
         Caption         =   "X1"
         Height          =   240
         Left            =   90
         TabIndex        =   24
         Top             =   285
         Width           =   240
      End
      Begin VB.Shape Shape1 
         Height          =   636
         Left            =   4500
         Top             =   444
         Width           =   732
      End
      Begin VB.Label X 
         Caption         =   "X1, Y1"
         Height          =   216
         Index           =   0
         Left            =   3840
         TabIndex        =   23
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label X 
         Caption         =   "X2, Y2"
         Height          =   240
         Index           =   1
         Left            =   4752
         TabIndex        =   22
         Top             =   120
         Width           =   636
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   168
         Left            =   4428
         Shape           =   3  'Circle
         Top             =   972
         Width           =   168
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   168
         Left            =   5148
         Shape           =   3  'Circle
         Top             =   396
         Width           =   168
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "m"
         Height          =   192
         Left            =   1380
         TabIndex        =   21
         Top             =   948
         Width           =   132
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Update Last ID"
      Height          =   375
      Left            =   4935
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2460
      Width           =   1404
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Lookup Last ID"
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2460
      Width           =   1404
   End
   Begin VB.TextBox txtLastID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   996
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2484
      Width           =   1065
   End
   Begin VB.CommandButton cmdAddUpdate 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   330
      Left            =   5535
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   96
      Width           =   825
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   330
      Left            =   5535
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   530
      Width           =   825
   End
   Begin VB.Frame TypeFrame 
      Caption         =   "Circle Center and Radius"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   1
      Left            =   84
      TabIndex        =   30
      Top             =   936
      Visible         =   0   'False
      Width           =   6315
      Begin VB.TextBox XY 
         Height          =   240
         Index           =   7
         Left            =   432
         TabIndex        =   7
         Top             =   360
         Width           =   780
      End
      Begin VB.TextBox txtRadius 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1092
         TabIndex        =   8
         Text            =   "1"
         Top             =   684
         Width           =   408
      End
      Begin VB.TextBox XY 
         Height          =   285
         Index           =   6
         Left            =   1545
         TabIndex        =   9
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Radius"
         Height          =   192
         Left            =   4284
         TabIndex        =   40
         Top             =   360
         Width           =   516
      End
      Begin VB.Line Line8 
         X1              =   4884
         X2              =   4884
         Y1              =   264
         Y2              =   636
      End
      Begin VB.Line Line7 
         Index           =   1
         X1              =   4896
         X2              =   4968
         Y1              =   252
         Y2              =   396
      End
      Begin VB.Line Line7 
         Index           =   0
         X1              =   4872
         X2              =   4824
         Y1              =   252
         Y2              =   396
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "m"
         Height          =   192
         Left            =   1548
         TabIndex        =   36
         Top             =   696
         Width           =   120
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   168
         Left            =   4812
         Shape           =   3  'Circle
         Top             =   552
         Width           =   168
      End
      Begin VB.Label X 
         AutoSize        =   -1  'True
         Caption         =   "X, Y"
         Height          =   192
         Index           =   2
         Left            =   4776
         TabIndex        =   35
         Top             =   780
         Width           =   300
      End
      Begin VB.Shape Shape6 
         Height          =   780
         Left            =   4416
         Shape           =   3  'Circle
         Top             =   240
         Width           =   948
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   192
         Left            =   168
         TabIndex        =   34
         Top             =   372
         Width           =   108
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Radius:"
         Height          =   192
         Left            =   468
         TabIndex        =   33
         Top             =   696
         Width           =   540
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         Height          =   195
         Left            =   1335
         TabIndex        =   32
         Top             =   375
         Width           =   105
      End
      Begin VB.Line Line6 
         BorderWidth     =   3
         X1              =   5796
         X2              =   5796
         Y1              =   804
         Y2              =   264
      End
      Begin VB.Line Line5 
         BorderWidth     =   3
         X1              =   5676
         X2              =   5796
         Y1              =   414
         Y2              =   204
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   5916
         X2              =   5796
         Y1              =   384
         Y2              =   204
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Grid North"
         Height          =   192
         Left            =   5400
         TabIndex        =   31
         Top             =   900
         Width           =   720
      End
   End
   Begin VB.Label Label15 
      Caption         =   $"AddUnitsEDM.frx":09DF
      Height          =   1455
      Left            =   6540
      TabIndex        =   37
      Top             =   1590
      Width           =   3135
   End
   Begin VB.Shape Shape5 
      Height          =   660
      Left            =   75
      Top             =   2370
      Width           =   6375
   End
   Begin VB.Shape Shape4 
      Height          =   2304
      Left            =   72
      Top             =   24
      Width           =   6372
   End
   Begin VB.Label Label9 
      Caption         =   $"AddUnitsEDM.frx":0AEB
      Height          =   615
      Left            =   30
      TabIndex        =   16
      Top             =   3090
      Width           =   6435
   End
   Begin VB.Label Label7 
      Caption         =   $"AddUnitsEDM.frx":0BB6
      Height          =   1485
      Left            =   6540
      TabIndex        =   15
      Top             =   60
      Width           =   3135
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Last ID"
      Height          =   252
      Left            =   300
      TabIndex        =   12
      Top             =   2520
      Width           =   612
   End
End
Attribute VB_Name = "AddUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Loading As Boolean
Public CloseOnAdd As Boolean
Public Editing As Boolean

Private Sub cmdAddUpdate_Click()

Cancelling = False
If txtUnit = "" Then
    MsgBox ("Input Unit Name.")
    txtUnit.SetFocus
    Exit Sub
End If

txtUnit = Trim(txtUnit)
Gotit = False
For I = 0 To 2
    If optType(I) Then
        Gotit = True
    End If
Next I

If Not Gotit Then
    MsgBox ("Select type of unit")
    optType(0).SetFocus
    Cancelling = True
    Exit Sub
End If

If Len(txtUnit) > UnitLength Then
    MsgBox ("Unit Name too long. Maximum of " & UnitLength & " characters allowed")
    Cancelling = True
    txtUnit.SetFocus
    Exit Sub
End If

If optType(0) And (XY(1) = "" Or XY(2) = "" Or XY(3) = "" Or XY(0) = "") Then
    Cancelling = True
    MsgBox ("All coordinates must be entered")
    Exit Sub
ElseIf optType(1) And (XY(6) = "" Or XY(7) = "") Then
    Cancelling = True
    MsgBox ("Center point of circle unit must be entered")
    Exit Sub
End If

If optType(0) And Val(XY(2)) <= Val(XY(0)) Then
    Cancelling = True
    MsgBox ("Invalid range:  X1 corner is lower left boundary in plan view.")
    Exit Sub
End If

If optType(0) And Val(XY(3)) <= Val(XY(1)) Then
    Cancelling = True
    MsgBox ("Invalid range:  Y1 corner is lower left boundary in plan view.")
    Exit Sub
End If

If cmdAddUpdate.Caption = "Add" Then
    txtUnit = UCase(txtUnit)
    If UnitTB.RecordCount > 0 Then
    
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
        
        UnitTB.MoveFirst
        If optType(0) Then
            While Not UnitTB.EOF
                If Not IsNull(UnitTB("MinX")) And Not IsNull(UnitTB("MaxX")) And Not IsNull(UnitTB("MinY")) And Not IsNull(UnitTB("MaxY")) Then
                    If ((XY(0) > Val(UnitTB("minx")) And XY(0) < Val(UnitTB("maxx"))) = True Or (XY(2) > Val(UnitTB("minx")) And XY(2) < Val(UnitTB("maxx")))) And ((XY(1) > Val(UnitTB("miny")) And XY(1) < Val(UnitTB("maxy"))) Or (XY(3) > Val(UnitTB("miny")) And XY(3) < Val(UnitTB("maxy")))) Then
                        response = MsgBox("Unit boundaries overlap with another unit.  Re-enter coordinates.")
                        Exit Sub
                    End If
                End If
                UnitTB.MoveNext
            Wend
        End If
    End If
    UnitTB.AddNew
    UnitTB("UNIT") = txtUnit
    UnitTB("ID") = PadID(txtLastID)
    UnitTB("suffix") = 0
    If optType(0) Then
        UnitTB("minX") = XY(0)
        UnitTB("miny") = XY(1)
        UnitTB("maxX") = XY(2)
        UnitTB("maxy") = XY(3)
    ElseIf optType(1) Then
        UnitTB("centerx") = XY(7)
        UnitTB("centery") = XY(6)
        UnitTB("radius") = txtRadius
    ElseIf optType(2) Then
        UnitTB("minX") = -99999
        UnitTB("miny") = -99999
        UnitTB("maxX") = -99999
        UnitTB("maxy") = -99999
    End If
    UnitTB.Update

    For I = 0 To 3
        XY(I) = ""
    Next I
    txtUnit.AddItem txtUnit
Else
    txtUnit = UCase(txtUnit)
    If UnitTB.RecordCount > 0 Then
        UnitTB.Index = "unitname"
        UnitTB.Seek "=", txtUnit
        If UnitTB.NoMatch Then
            MsgBox ("Unit not found")
            cmdAddUpdate.Caption = "Add"
            Exit Sub
        End If
                
        UnitTB.MoveFirst
        If optType(0) Then
            While Not UnitTB.EOF
                If (XY(0) > UnitTB("minx") And XY(0) < UnitTB("maxx")) Or (XY(2) > UnitTB("minx") And XY(2) < UnitTB("maxx")) Or (XY(1) > UnitTB("miny") And XY(1) < UnitTB("maxy")) Or (XY(3) > UnitTB("miny") And XY(3) < UnitTB("maxy")) Then
                    response = ("Unit boundaries overlap with another unit.  Re-enter coordinates.")
                    Exit Sub
                End If
                UnitTB.MoveNext
            Wend
        End If
    End If
    UnitTB.Index = "unitname"
    UnitTB.Seek "=", txtUnit
    If Not UnitTB.NoMatch Then
        UnitTB.Edit
        UnitTB("ID") = PadID(txtLastID)
        UnitTB("suffix") = 0
        If optType(0) Then
            UnitTB("minX") = XY(0)
            UnitTB("miny") = XY(1)
            UnitTB("maxX") = XY(2)
            UnitTB("maxy") = XY(3)
        ElseIf optType(1) Then
            UnitTB("centerx") = XY(7)
            UnitTB("centery") = XY(6)
            UnitTB("radius") = txtRadius
        ElseIf optType(2) Then
            UnitTB("minX") = -99999
            UnitTB("miny") = -99999
            UnitTB("maxX") = -99999
            UnitTB("maxy") = -99999
        End If
        For I = 0 To 3
            XY(I) = ""
        Next I

        UnitTB.Update
    End If
End If

txtUnit = ""
txtLastID = 0
If CloseOnAdd Then
    CloseOnAdd = False
    Unload Me
End If
txtUnit.SetFocus
cmdAddUpdate.Enabled = False
Editing = False

End Sub

Private Sub Command1_Click()

For I = 0 To 2
    optType(I) = False
Next I
TypeFrame(1).Visible = False
TypeFrame(0).Visible = False
txtUnit = ""
txtLastID = 0
txtUnit.SetFocus
cmdAddUpdate.Enabled = False
Editing = False

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
    cmdAddUpdate.Caption = "Add"
    cmdAddUpdate.Enabled = True
    For I = 0 To 2
        optType(I) = False
    Next I
    TypeFrame(1).Visible = False
    TypeFrame(0).Visible = False
    Editing = False
    txtUnit.SetFocus
Else
    MsgBox ("Unit not found.")
End If
Editing = False

End Sub

Private Sub Command3_Click()

If txtUnit = "" Then
    MsgBox ("Select or enter a unit name first.")
    Exit Sub
End If

Dim rsTemp As Recordset
Dim MaxID As Long

If tablematch("context") Then
    SqlString = "select max(id) from context where id<'A' and unit='" + txtUnit + "'"
    Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp(0)) Then
            If Val(Trim(rsTemp(0))) > MaxID Then
                MaxID = Val(Trim(rsTemp(0)))
            End If
        End If
    End If
End If

If tablematch("xyz") Then
    SqlString = "select max(id) from xyz where id<'A' and unit='" + txtUnit + "'"
    Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp(0)) Then
            If Val(Trim(rsTemp(0))) > MaxID Then
                MaxID = Val(Trim(rsTemp(0)))
            End If
        End If
    End If
End If

txtLastID = MaxID
Command4_Click
Editing = True
Set rsTemp = Nothing

End Sub

Private Sub Command4_Click()

If txtUnit = "" Then
    MsgBox ("Select or Add Unit first")
    Exit Sub
End If

UnitTB.Index = "unitname"
UnitTB.Seek "=", txtUnit
If Not UnitTB.NoMatch Then
    UnitTB.Edit
    UnitTB("ID") = txtLastID
    UnitTB.Update
End If

End Sub

Private Sub Form_Load()

Cancelling = True
txtUnit.Clear

If UnitTB.BOF <> UnitTB.EOF Then
    UnitTB.MoveFirst
    While Not UnitTB.EOF
        txtUnit.AddItem UnitTB("Unit")
        UnitTB.MoveNext
    Wend
End If
cmdAddUpdate.Enabled = False
For I = 0 To 2
    optType(I) = False
Next I
TypeFrame(1).Visible = False
TypeFrame(0).Visible = False

For I = 1 To Vars
    Select Case LCase(VarList(I))
        Case "day", "time", "unit", "id", "suffix", "x", "y", "z", "vangle", "hangle"
        Case Else
            lstUnitFields.AddItem VarList(I)
            lstUnitFields.ItemData(lstUnitFields.NewIndex) = I
            For J = 1 To nUnitFields
                If LCase(Unitfield(J)) = LCase(VarList(I)) Then
                    lstUnitFields.Selected(lstUnitFields.NewIndex) = True
                    Exit For
                End If
            Next J
    End Select
Next I

If Not tablematch("context") And Not tablematch("xyz") Then Command3.Enabled = False
CenterForm Me

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim VarNumber(500) As Integer

If Editing Then
    response = MsgBox("Save Changes?", vbYesNo)
    If response = vbYes Then
        cmdAddUpdate_Click
        If Cancelling Then Exit Sub
    End If
End If

nUnitFields = 0
For I = 0 To lstUnitFields.ListCount - 1
    If lstUnitFields.Selected(I) Then
        nUnitFields = nUnitFields + 1
        Unitfield(nUnitFields) = lstUnitFields.List(I)
        VarNumber(nUnitFields) = lstUnitFields.ItemData(I)
    End If
Next I

Dim Inidata(1, 2) As String
Dim IniClass As String
Dim Status As Byte

IniClass = "[EDM]"
Inidata(1, 1) = "UnitFields"
Inidata(1, 2) = Unitfield(1)
For I = 2 To nUnitFields
    Inidata(1, 2) = Inidata(1, 2) + "," + Unitfield(I)
Next I
Call WriteIni(CFGName, IniClass, Inidata(), Status)

Set UnitTB = Nothing
For I = 1 To nUnitFields
    On Error Resume Next
    Gotit = False
    Gotit = LCase(SiteDB.TableDefs("edm_units").Fields(Unitfield(I)).Name) = LCase(Unitfield(I))
    On Error GoTo 0
    If Not Gotit Then
        With SiteDB.TableDefs("edm_units")
            Select Case LCase(VType(VarNumber(I)))
                Case "menu", "text"
                    .Fields.Append .CreateField(Unitfield(I), dbText, VLen(VarNumber(I)))
                Case "numeric", "instrument"
                    .Fields.Append .CreateField(Unitfield(I), dbSingle)
            End Select
        End With
    End If
Next I
Set UnitTB = SiteDB.OpenRecordset("EDM_Units")
Unload Me

End Sub

Private Sub optType_Click(Index As Integer)

Select Case Index
    Case 0
        TypeFrame(0).Visible = True
        TypeFrame(1).Visible = False

    Case 1
        TypeFrame(1).Visible = True
        TypeFrame(0).Visible = False
    
    Case 2
        TypeFrame(1).Visible = False
        TypeFrame(0).Visible = False

End Select

End Sub

Private Sub txtLastID_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 8, 48 To 57
    Case Else
        KeyAscii = 0
End Select
Editing = True

End Sub

Private Sub txtLastID_LostFocus()

Command4_Click

End Sub

Private Sub txtSize_Change()

If IsNumeric(txtSize) Then
    If IsNumeric(XY(0)) Then
        XY(2) = Val(XY(0)) + Val(txtSize)
    End If
    If IsNumeric(XY(1)) Then
        XY(3) = Val(XY(1)) + Val(txtSize)
    End If
    
End If

End Sub

Private Sub txtUnit_Click()

If txtUnit = "" Then
    Exit Sub
End If
UnitTB.Index = "unitname"
UnitTB.Seek "=", txtUnit

If Not UnitTB.NoMatch Then
    If Not IsNull(UnitTB("radius")) Then
        optType(1) = True
        If Not IsNull(UnitTB("centerx")) Then
            XY(7) = UnitTB("centerx")
        Else
            XY(7) = 0
        End If
        If Not IsNull(UnitTB("centery")) Then
            XY(6) = UnitTB("centery")
        Else
            XY(6) = 0
        End If
        txtRadius = UnitTB("radius")
    ElseIf UnitTB("minx") = -99999 Then
        optType(2) = True
    Else
        optType(0) = True
        If Not IsNull(UnitTB("minx")) Then
            XY(0) = UnitTB("minx")
        Else
            XY(0) = 0
        End If
        If Not IsNull(UnitTB("miny")) Then
            XY(1) = UnitTB("miny")
        Else
            XY(1) = 0
        End If
        If Not IsNull(UnitTB("maxx")) Then
            XY(2) = UnitTB("maxx")
        Else
            XY(2) = 0
        End If
        If Not IsNull(UnitTB("maxy")) Then
            XY(3) = UnitTB("maxy")
        Else
            XY(3) = 0
        End If
    End If
    If IsNull(UnitTB("ID")) Then
        txtLastID = 0
    Else
        txtLastID = UnitTB("ID")
    End If
    cmdAddUpdate.Caption = "Update"
    cmdAddUpdate.Enabled = True
End If

End Sub

Private Sub txtUnit_DropDown()

SqlString = "Select Unit from [EDM_Units]"
Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
txtUnit.Clear
While Not rsTemp.EOF
    txtUnit.AddItem rsTemp(0)
    rsTemp.MoveNext
Wend
Set rsTemp = Nothing

End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)
    
cmdAddUpdate.Caption = "Add"
cmdAddUpdate.Enabled = True
If UpperCase Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
Editing = True

End Sub

Private Sub XY_Change(Index As Integer)

If Loading Then
    If Index < 2 Then
        XY(Index + 2) = Val(XY(Index)) + Val(txtSize)
    End If
    Loading = False
End If

End Sub

Private Sub XY_GotFocus(Index As Integer)

SelStart = 0
SelLength = Len(XY(Index))

End Sub

Private Sub XY_KeyPress(Index As Integer, KeyAscii As Integer)

Loading = True
Select Case KeyAscii
    Case 8, 46, 48 To 57, Asc("-"), Asc(".")
    Case Else
        KeyAscii = 0
End Select
Editing = True

End Sub





