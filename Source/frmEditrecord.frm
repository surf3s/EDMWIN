VERSION 5.00
Begin VB.Form frmEditrecord 
   Caption         =   "Edit record"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   4200
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data pointdata 
      Caption         =   "Point Data"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   372
      Index           =   0
      Left            =   1080
      TabIndex        =   26
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Delete"
      Height          =   372
      Index           =   1
      Left            =   2160
      TabIndex        =   25
      Top             =   4440
      Width           =   1092
   End
   Begin VB.CommandButton varmenu 
      Caption         =   "M"
      Height          =   375
      Index           =   7
      Left            =   3360
      TabIndex        =   24
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox varvalue 
      Alignment       =   1  'Right Justify
      DataSource      =   "pointdata"
      Height          =   372
      Index           =   7
      Left            =   1560
      TabIndex        =   22
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton varmenu 
      Caption         =   "M"
      Height          =   375
      Index           =   6
      Left            =   3360
      TabIndex        =   21
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox varvalue 
      Alignment       =   1  'Right Justify
      DataSource      =   "pointdata"
      Height          =   372
      Index           =   6
      Left            =   1560
      TabIndex        =   19
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton varmenu 
      Caption         =   "M"
      Height          =   375
      Index           =   5
      Left            =   3360
      TabIndex        =   18
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox varvalue 
      Alignment       =   1  'Right Justify
      DataSource      =   "pointdata"
      Height          =   372
      Index           =   5
      Left            =   1560
      TabIndex        =   16
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton varmenu 
      Caption         =   "M"
      Height          =   375
      Index           =   4
      Left            =   3360
      TabIndex        =   15
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox varvalue 
      Alignment       =   1  'Right Justify
      DataSource      =   "pointdata"
      Height          =   372
      Index           =   4
      Left            =   1560
      TabIndex        =   13
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton varmenu 
      Caption         =   "M"
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   12
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox varvalue 
      Alignment       =   1  'Right Justify
      DataSource      =   "pointdata"
      Height          =   372
      Index           =   3
      Left            =   1560
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton varmenu 
      Caption         =   "M"
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   9
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox varvalue 
      Alignment       =   1  'Right Justify
      DataSource      =   "pointdata"
      Height          =   372
      Index           =   2
      Left            =   1560
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton varmenu 
      Caption         =   "M"
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   6
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox varvalue 
      Alignment       =   1  'Right Justify
      DataSource      =   "pointdata"
      Height          =   372
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.VScrollBar fieldscroll 
      Height          =   3735
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton varmenu 
      Caption         =   "M"
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox varvalue 
      Alignment       =   1  'Right Justify
      DataSource      =   "pointdata"
      Height          =   372
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label varcaption 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Variable Caption :"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   23
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label varcaption 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Variable Caption :"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label varcaption 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Variable Caption :"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label varcaption 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Variable Caption :"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label varcaption 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Variable Caption :"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label varcaption 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Variable Caption :"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label varcaption 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Variable Caption :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label varcaption 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Variable Caption :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmEditrecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

Select Case Index
Case 0
    Unload Me
Case 1
    Unload Me
Case Else
End Select

End Sub

Private Sub fieldscroll_Change()

fieldno = fieldscroll.value
For a = 0 To 7
    If fieldno >= pointdata.Recordset.Fields.Count Then
        varvalue(a).Text = ""
        varmenu(a).Visible = False
        varcaption(a).Caption = ""
        varvalue(a).Visible = False
    Else
        varvalue(a).Visible = True
        varcaption(a).Caption = pointdata.Recordset.Fields(fieldno).name
        varvalue(a).DataField = pointdata.Recordset.Fields(fieldno).name
        'If Not IsNull(pointdata.Recordset.Fields(fieldno)) Then
        '    varvalue(a).Text = pointdata.Recordset.Fields(fieldno)
        'Else
        '    varvalue(a).Text = ""
        'End If
        'If pointdata.Recordset.Fields(fieldno).Type = dbText Then
        '    varmenu(a).Visible = True
        'Else
        '    varmenu(a).Visible = False
        'End If
        varmenu(a).Tag = pointdata.Recordset.Fields(fieldno).name
        varvalue(a).Tag = pointdata.Recordset.Fields(fieldno).name
    End If
    fieldno = fieldno + 1
Next a

End Sub

Private Sub Form_Load()

Me.Width = 4320
Me.Height = 5310

Set pointdata.Recordset = PointsTB

Call fillform

End Sub

Sub fillform()

pointdata.Recordset.MoveLast

fieldscroll.min = 0
fieldscroll.Max = pointdata.Recordset.Fields.Count - 1
fieldscroll.SmallChange = 1
fieldscroll.LargeChange = 8
fieldscroll.value = 0
Call fieldscroll_Change

End Sub

Private Sub pointdata_Reposition()

Call fieldscroll_Change

End Sub

Private Sub varmenu_Click(Index As Integer)

fieldname$ = varmenu(Index).Tag
    
If fieldname$ = "Datum" Then
    If Not DatumTB.EOF Or Not DatumTB.BOF Then
        DatumTB.MoveFirst
        Do Until DatumTB.EOF
            frmMenu.menulist.AddItem DatumTB("Datum")
            DatumTB.MoveNext
        Loop
    End If
ElseIf fieldname$ = "Pole" Then
    If Not PoleTB.EOF Or Not PoleTB.BOF Then
        PoleTB.MoveFirst
        Do Until PoleTB.EOF
            frmMenu.menulist.AddItem PoleTB("Pole") + " " + Str$(PoleTB("height"))
            PoleTB.MoveNext
        Loop
    End If
Else
    menuitem$ = pointdata.Recordset.Fields(fieldname$).Properties("rowsource")
    Do Until menuitem$ = ""
        a = InStr(menuitem$, ";")
        If a = 0 And Len(menuitem$) <> 0 Then
            frmMenu.menulist.AddItem menuitem$
            menuitem$ = ""
        Else
            frmMenu.menulist.AddItem Left$(menuitem$, a - 1)
            menuitem$ = Mid$(menuitem$, a + 1)
        End If
    Loop
End If
frmMenu.Caption = fieldname$ + " menu"
frmMenu.menutitle = "Select from the following :"
frmMenu.Show 1
If MenuSelection$ <> "" Then
    Select Case fieldname$
    Case "Datum"
        varvalue(a).Text = MenuSelection$
    Case "Pole"
        b = InStr(MenuSelection$, " ")
        If b <> 0 Then
            varvalue(a).Text = Mid$(MenuSelection$, b + 1)
        End If
    Case Else
        varvalue(a).Text = MenuSelection$
    End Select
End If

End Sub


