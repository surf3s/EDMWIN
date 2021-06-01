VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDatumSheet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit or Add Datums"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11220
   ControlBox      =   0   'False
   Icon            =   "frmDatumSheet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Add New Datum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2070
      Left            =   7680
      TabIndex        =   9
      Top             =   60
      Width           =   3465
      Begin VB.CommandButton Command4 
         Caption         =   "Save"
         Height          =   405
         Left            =   2430
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   900
         Width           =   915
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Record"
         Default         =   -1  'True
         Height          =   405
         Left            =   2430
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   765
         TabIndex        =   0
         Top             =   450
         Width           =   1575
      End
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   765
         TabIndex        =   1
         Top             =   825
         Width           =   915
      End
      Begin VB.TextBox txtY 
         Height          =   285
         Left            =   765
         TabIndex        =   2
         Top             =   1185
         Width           =   915
      End
      Begin VB.TextBox txtZ 
         Height          =   285
         Left            =   765
         TabIndex        =   3
         Top             =   1530
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   405
         Left            =   2430
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1395
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X"
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
         Left            =   525
         TabIndex        =   13
         Top             =   885
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Y"
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
         Left            =   525
         TabIndex        =   12
         Top             =   1245
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Z"
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
         Left            =   525
         TabIndex        =   11
         Top             =   1590
         Width           =   120
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1224
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Datums"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4068
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   7605
      Begin MSDBGrid.DBGrid DatumSheet 
         Bindings        =   "frmDatumSheet.frx":000C
         Height          =   3165
         Left            =   90
         OleObjectBlob   =   "frmDatumSheet.frx":0024
         TabIndex        =   16
         Top             =   810
         Width           =   7455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Delete Datum"
         Height          =   405
         Left            =   3300
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   270
         Width           =   1815
      End
      Begin VB.Data datumdata 
         Caption         =   "Datums"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2610
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1590
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin VB.Label Label4 
      Caption         =   $"frmDatumSheet.frx":09FA
      Height          =   1245
      Left            =   7680
      TabIndex        =   15
      Top             =   2175
      Width           =   3465
   End
   Begin VB.Label Label3 
      Caption         =   $"frmDatumSheet.frx":0B06
      Height          =   630
      Left            =   7680
      TabIndex        =   14
      Top             =   3480
      Width           =   3465
   End
End
Attribute VB_Name = "frmDatumSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Saving As Boolean
Dim Editing As Boolean

Private Sub sizecontrols()

Dim RowNum As Integer
'DatumSheet.Cols = 1
'DatumSheet.Rows = 1

datumdata.Refresh
DatumSheet.Refresh
For I = 0 To datumdata.Recordset.Fields.Count - 1
    Select Case UCase(datumdata.Recordset.Fields(I).Name)
        Case "NAME"
            DatumSheet.Columns(I).Visible = True
            DatumSheet.Columns(I).Width = 1700
        Case "X", "Y", "Z"
            DatumSheet.Columns(I).Visible = True
            DatumSheet.Columns(I).Width = Me.TextWidth("0000000.000")
            DatumSheet.Columns(I).NumberFormat = "#########0.000"
        Case Else
            DatumSheet.Columns(I).Visible = True
            DatumSheet.Columns(I).Width = Me.TextWidth("0000000000000")
    End Select
Next I

'DatumSheet.Columns(1).Width = Me.TextWidth("00/00/0000")
'DatumSheet.Columns(2).Width = Me.TextWidth("")
'DatumSheet.Columns(1).Visible = False
'DatumSheet.Columns(2).Visible = False
'
'DatumSheet.Columns(3).Width = Me.TextWidth("0000000.000")
'DatumSheet.Columns(4).Width = Me.TextWidth("0000000.000")
'DatumSheet.Columns(5).Width = Me.TextWidth("0000000.000")
'DatumSheet.Columns(3).NumberFormat = "#########0.000"
'DatumSheet.Columns(4).NumberFormat = "#########0.000"
'DatumSheet.Columns(5).NumberFormat = "#########0.000"


'DatumSheet.Cols = 6
'
'DatumSheet.ColWidth(0) = 300
'DatumSheet.ColWidth(1) = Me.TextWidth("00/00/0000")
'DatumSheet.ColWidth(2) = Me.TextWidth("00000000.000")
'DatumSheet.ColWidth(3) = Me.TextWidth("00000000.000")
'DatumSheet.ColWidth(4) = Me.TextWidth("00000000.000")
'DatumSheet.TextMatrix(0, 1) = "Name"
'DatumSheet.TextMatrix(0, 2) = "Creation"
'DatumSheet.TextMatrix(0, 3) = "X"
'DatumSheet.TextMatrix(0, 4) = "Y"
'DatumSheet.TextMatrix(0, 5) = "Z"

If Not DatumTB.EOF Or Not DatumTB.BOF Then DatumTB.MoveLast
If DatumTB.BOF Then Exit Sub

'DatumSheet.Rows = DatumTB.RecordCount + 1
'
'DatumTB.MoveFirst
'While Not DatumTB.EOF
'    RowNum = RowNum + 1
'    If Not IsNull(DatumTB("name")) Then
'        DatumSheet.TextMatrix(RowNum, 1) = DatumTB("Name")
'    Else
'        DatumSheet.TextMatrix(RowNum, 1) = ""
'    End If
'    If Not IsNull(DatumTB("day")) Then
'        DatumSheet.TextMatrix(RowNum, 2) = DatumTB("day")
'    Else
'        DatumSheet.TextMatrix(RowNum, 2) = ""
'    End If
'    DatumSheet.TextMatrix(RowNum, 3) = Format(DatumTB("x"), "#####0.000")
'    DatumSheet.TextMatrix(RowNum, 4) = Format(DatumTB("y"), "#####0.000")
'    DatumSheet.TextMatrix(RowNum, 5) = Format(DatumTB("z"), "#####0.000")
'    DatumTB.MoveNext
'Wend

End Sub

Private Sub Command2_Click()

If mdiMain.StatusBar.Panels(7).Visible Then
    Cancelling = True
    Exit Sub
ElseIf Shooting Then
    Exit Sub
Else
    If txtName <> "" Or txtX <> "" Or txtY <> "" Or txtZ <> "" Then
        response = MsgBox("Datum not saved.  Save before closing?", vbYesNo)
        If response = vbYes Then
            Command4_Click
            Exit Sub
        End If
    End If
    
    Unload Me
End If

End Sub

Private Sub Command3_Click()

If frmMain.lblPoleWarning.Visible Then
    MsgBox ("You must define prisms before recording a new datum with the EDM")
    Exit Sub
End If
Command3.Enabled = False
Call takeshot_core(AskForPrism)
mdiMain.StatusBar.Panels(6).Visible = False

If Cancelling Then
    Command3.Enabled = True
    Command3.Default = True
    Exit Sub
End If

Command3.Enabled = True
If errorcode = 0 And Not Cancelling Then
    txtX = Format(edmshot.X, "######0.000")
    txtY = Format(edmshot.y, "######0.000")
    txtZ = Format(edmshot.z, "######0.000")
End If
Cancelling = False
Command4.Default = True

End Sub

Private Sub Command4_Click()

If txtName = "" Then
    MsgBox ("Enter name for the datum")
    Exit Sub
End If
If txtX = "" Or txtY = "" Or txtZ = "" Then
    MsgBox ("Record new datum or manually enter coordinates before saving")
    Exit Sub
End If
If Not IsNumeric(txtX) Or Not IsNumeric(txtY) Or Not IsNumeric(txtZ) Then
    MsgBox ("Enter coordinates as numeric values.")
    Exit Sub
End If
    
DatumTB.Index = "datumname"
DatumTB.Seek "=", txtName
If Not DatumTB.NoMatch Then
    response = MsgBox(txtName + " already exists.  Replace?", vbYesNo)
    If response = vbYes Then
        DatumTB.Delete
    Else
        Exit Sub
    End If
End If

DatumTB.AddNew
DatumTB("Name") = txtName
DatumTB("X") = txtX
DatumTB("y") = txtY
DatumTB("z") = txtZ
DatumTB("day") = Date
DatumTB("time") = Time
DatumTB.Update
DatumTB.MoveLast
sizecontrols

'DoEvents
txtName = ""
txtX = ""
txtY = ""
txtZ = ""
txtName.SetFocus

End Sub

Private Sub Command5_Click()

If DatumSheet.Row > -1 And DatumTB.RecordCount > 0 Then
    response = MsgBox("Permanently delete datum " + DatumTB("name") + "?", vbYesNo)
    If response = vbYes Then
        datumdata.Recordset.Delete
        If DatumTB.RecordCount > 0 Then
            DatumTB.MoveFirst
        End If
        sizecontrols
        'DoEvents
    End If
End If

End Sub

Private Sub DatumSheet_AfterColEdit(ByVal ColIndex As Integer)

Command2.Cancel = True

End Sub

Private Sub DatumSheet_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)

Command2.Cancel = False
Command3.Default = False
Command4.Default = False
Cancelling = False

End Sub

Private Sub DatumSheet_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)

A = 1
If Cancelling Then
    Cancel = False
    DatumSheet.Columns(ColIndex) = OldValue
End If

End Sub

Private Sub DatumSheet_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyEscape Then
    Cancelling = True
End If

End Sub

Private Sub datumsheet_KeyPress(KeyAscii As Integer)

If DatumSheet.Col > 2 Then
    Select Case KeyAscii
        Case 8, 48 To 57, Asc("-"), Asc(".")
        Case Asc(",")
            KeyAscii = Asc(".")
        Case 27
            Cancelling = True
        Case Else
            KeyAscii = 0
    End Select
End If

End Sub

Private Sub DatumSheet_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

If Not DatumTB.BOF And Not DatumTB.EOF Then
    DatumTB.Bookmark = DatumSheet.Bookmark
End If

End Sub

'Private Sub datumsheet_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'If DatumTB.RecordCount = 0 Then
'    Exit Sub
'End If
'If Not Loading Then
'    datumdata.Recordset.Bookmark = datumsheet.Bookmark
'    DatumTB.Bookmark = datumsheet.Bookmark
'End If
'End Sub

Private Sub Form_Load()

DatumTB.Index = "datumname"
Set datumdata.Recordset = DatumTB
datumdata.Refresh

'DatumTB.MoveLast
'datumsheet.ReBind
'datumsheet.Refresh

sizecontrols
CenterForm Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set datumdata.Recordset = Nothing
Me.Hide
frmMain.Picture1.SetFocus

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

If UpperCase Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
Select Case KeyAscii
    Case 8
    Case Else
        If Len(txtName) >= 20 Then
            KeyAscii = 0
        End If
End Select

End Sub

Private Sub txtX_KeyPress(KeyAscii As Integer)

Command4.Default = True
Select Case KeyAscii
    Case 8, 48 To 57, Asc("-"), Asc(".")
    Case Asc(",")
        KeyAscii = Asc(".")
    Case Else
        KeyAscii = 0
End Select

End Sub

Private Sub txtY_KeyPress(KeyAscii As Integer)
    
Select Case KeyAscii
    Case 8, 48 To 57, Asc("-"), Asc(".")
    Case Asc(",")
        KeyAscii = Asc(".")
    Case Else
        KeyAscii = 0
End Select

End Sub

Private Sub txtZ_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 8, 48 To 57, Asc("-"), Asc(".")
    Case Asc(",")
        KeyAscii = Asc(".")
    Case Else
        KeyAscii = 0
End Select

End Sub

