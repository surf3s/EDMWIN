VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPolesheet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Poles"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9600
   ControlBox      =   0   'False
   Icon            =   "frmPolesheet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Add New Prism"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   5160
      TabIndex        =   12
      Top             =   180
      Width           =   3075
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1950
         Width           =   1035
      End
      Begin VB.TextBox txtPHeight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1860
         TabIndex        =   1
         Top             =   960
         Width           =   945
      End
      Begin VB.TextBox txtPoffset 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1860
         TabIndex        =   2
         Top             =   1410
         Width           =   945
      End
      Begin VB.TextBox txtPname 
         Height          =   285
         Left            =   1860
         TabIndex        =   0
         Top             =   540
         Width           =   945
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pole Height (m):"
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
         Left            =   315
         TabIndex        =   15
         Top             =   1020
         Width           =   1380
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Prism Offset (mm):"
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
         Left            =   150
         TabIndex        =   14
         Top             =   1440
         Width           =   1545
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Prism Name:"
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
         Left            =   630
         TabIndex        =   13
         Top             =   600
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Prisms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   90
      TabIndex        =   11
      Top             =   180
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "Delete Prism"
         Height          =   405
         Left            =   3390
         TabIndex        =   5
         Top             =   270
         Width           =   1425
      End
      Begin VB.Data poledata 
         Caption         =   "Poles"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   1140
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   0  'Table
         RecordSource    =   ""
         Top             =   2070
         Visible         =   0   'False
         Width           =   2535
      End
      Begin MSDBGrid.DBGrid polesheet 
         Bindings        =   "frmPolesheet.frx":000C
         Height          =   1635
         Left            =   150
         OleObjectBlob   =   "frmPolesheet.frx":0023
         TabIndex        =   4
         Top             =   810
         Width           =   4665
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   405
      Left            =   8400
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   270
      Width           =   1065
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPolesheet.frx":09F8
      Height          =   1455
      Left            =   90
      TabIndex        =   10
      Top             =   3600
      Width           =   9435
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To delete a prism, select the prism name in the grid, then click on the Delete Prism button."
      Height          =   555
      Left            =   90
      TabIndex        =   9
      Top             =   3330
      Width           =   9855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To add a new prism, fill in the values for Name, Height, and Offset, then click on the Add button."
      Height          =   255
      Left            =   90
      TabIndex        =   8
      Top             =   3060
      Width           =   9855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit existing prisms directly on the grid.  All changes will take effect immediately."
      Height          =   435
      Left            =   90
      TabIndex        =   7
      Top             =   2790
      Width           =   9855
   End
End
Attribute VB_Name = "frmPolesheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GotName, GotHeight, GotOffset As Boolean

Public Sub cmdAdd_Click()

If txtPname = "" Then
    MsgBox ("Enter Name for prism (15 chars or less)")
    Exit Sub
End If

If txtPHeight = "" Then
    MsgBox ("Enter Pole Height")
    Exit Sub
End If

If txtPoffset = "" Then
    MsgBox ("Enter offset for prism (in millimeters)")
    Exit Sub
End If

If poledata.Recordset.RecordCount > 0 Then
    poledata.Recordset.MoveFirst
    Do While Not poledata.Recordset.EOF
        If LCase(poledata.Recordset("name")) = LCase(txtPname) Then
            MsgBox ("Duplicate prism name.")
            txtPname.SetFocus
            Exit Sub
        End If
        poledata.Recordset.MoveNext
    Loop
End If

poledata.Recordset.AddNew
poledata.Recordset("Name") = txtPname
poledata.Recordset("height") = txtPHeight
poledata.Recordset("offset") = txtPoffset
poledata.Recordset.Update
poledata.Recordset.MoveLast
txtPname = ""
txtPHeight = ""
txtPoffset = ""
GotHeight = False
GotName = False
GotOffset = False
Command2.Default = True
txtPname.SetFocus

End Sub

Private Sub Command1_Click()

If poledata.Recordset.BOF Or poledata.Recordset.EOF Then
    Exit Sub
End If
poledata.Recordset.Delete
poledata.Recordset.MoveFirst

End Sub

Private Sub Command2_Click()

If txtPname <> "" Or txtPHeight <> "" Or txtPoffset <> "" Then
    response = MsgBox("Add new prism?", vbYesNo)
    If response = vbYes Then
        cmdAdd_Click
    End If
End If
Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

Me.Hide
Set poledata.Recordset = Nothing
frmMain.Picture1.SetFocus

End Sub

Private Sub polesheet_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case 13
    If polesheet.Col + 1 = poledata.Recordset.Fields.Count Then
        polesheet.Col = 0
    Else
        polesheet.Col = polesheet.Col + 1
    End If
Case Else
    If polesheet.Col > 0 Then
        Select Case KeyAscii
            Case 8, 46, 48 To 57, Asc("-"), Asc(".")
            Case Else
                KeyAscii = 0
        End Select
     Else
         If UpperCase Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
          End If
    End If
End Select

End Sub

Private Sub Form_Load()

Set poledata.Recordset = PoleTB
poledata.Refresh
polesheet.AllowRowSizing = False
CenterForm Me

End Sub

Private Sub txtPHeight_GotFocus()

txtPHeight.SelStart = 0
txtPHeight.SelLength = Len(txtPHeight)

End Sub

Private Sub txtPHeight_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 8, 48 To 57, Asc("-"), Asc(".")
    Case Else
        KeyAscii = 0
End Select
GotHeight = True
If GotName And GotHeight And GotOffset Then
    cmdAdd.Default = True
End If

End Sub

Private Sub txtPHeight_LostFocus()

txtPHeight = Format(txtPHeight, "###0.000")
If Val(txtPHeight) < -10 Or Val(txtPHeight) > 10 Then
    MsgBox ("Prism heights must be between -10 and 10 meters (tip: use negative numbers for measuring ceilings")
    txtPHeight.SetFocus
End If

End Sub

Private Sub txtPname_GotFocus()

txtPname.SelStart = 0
txtPname.SelLength = Len(txtPname)

End Sub

Private Sub txtPname_KeyPress(KeyAscii As Integer)

If UpperCase Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
Select Case KeyAscii
    Case 8
    Case Else
        If Len(txtPname) > 15 Then
            KeyAscii = 0
        End If
End Select
GotName = True
If GotName And GotHeight And GotOffset Then
    cmdAdd.Default = True
End If

End Sub

Private Sub txtPoffset_GotFocus()

txtPoffset.SelStart = 0
txtPoffset.SelLength = Len(txtPoffset)

End Sub

Private Sub txtPoffset_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 8, 48 To 57, Asc("-"), Asc(".")
    Case Else
        KeyAscii = 0
End Select
GotOffset = True
If GotName And GotHeight And GotOffset Then
    cmdAdd.Default = True
End If

End Sub

Private Sub txtPoffset_LostFocus()

If InStr(txtPoffset, ".") = 0 Then
    txtPoffset = Val(txtPoffset) / 1000
End If
txtPoffset = Format(txtPoffset, "###0.000")
If Val(txtPoffset) < -0.05 Or Val(txtPoffset) > 0.05 Then
    MsgBox ("Prism offsets must be between -0.05 and 0.05 meters")
    txtPoffset.SetFocus
End If

End Sub


