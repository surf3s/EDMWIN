VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDataGrid 
   Caption         =   "Data"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9405
   Icon            =   "frmDataGrid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2895
   ScaleWidth      =   9405
   Begin MSDataGridLib.DataGrid DataGrid 
      Height          =   1965
      Left            =   2160
      TabIndex        =   4
      Top             =   420
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   3466
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   345
      Left            =   2130
      Top             =   2490
      Visible         =   0   'False
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   -1
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Filter by ..."
      Height          =   345
      Left            =   60
      TabIndex        =   3
      Top             =   2460
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.ListBox lstFields 
      Height          =   1860
      Left            =   60
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   1905
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Hit ESCAPE to return to main form for editing"
      Height          =   270
      Left            =   2160
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   6930
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Display Fields"
      Height          =   195
      Left            =   465
      TabIndex        =   1
      Top             =   90
      Width           =   1035
   End
End
Attribute VB_Name = "frmDataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Inidata(100, 2) As String
Dim IniClass As String
Dim Status As Byte
Dim GridFieldString As String
Dim nGridFields As Integer
Dim GridField(100) As String

Public Sub OpenGrid()

DataGrid.Visible = False
Data1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + SiteDBname + ";Persist Security Info=False"
Set Data1.Recordset = frmMain.PointsADO.Recordset.Clone
Set DataGrid.datasource = frmMain.PointsADO
DataGrid.Refresh
FormatGrid
If CountRecords = 0 Then Data1.Caption = "0\0"
DataGrid.Visible = True

End Sub

Private Sub Command1_Click()

frmFilter.Show 1

End Sub

Private Sub DataGrid_KeyPress(KeyAscii As Integer)

Dim Obj As Object

If GridLoading Then Exit Sub

If KeyAscii = 27 Then
    Set Obj = frmMain.Picture1
    Select Case UCase(Data1.Recordset.Fields(DataGrid.Col).Name)
        Case "UNIT"
            Set object = frmMain.txtUnit
        Case "ID"
            Set object = frmMain.txtID
        Case "SUFFIX"
            Exit Sub
        Case "PRISM"
            Set object = frmMain.txtprism
        Case "X"
            Set object = frmMain.txtXYZ(0)
        Case "Y"
            Set object = frmMain.txtXYZ(1)
        Case "Z"
            Set object = frmMain.txtXYZ(2)
        Case "HANGLE"
            Set object = frmMain.txtHangle
        Case "VANGLE"
            Set object = frmMain.txtVangle
        Case "SLOPED"
            Set object = frmMain.txtSloped
        Case Else
            Gotit = False
            For I = 1 To Vars
                If UCase(VarList(I)) = UCase(Data1.Recordset.Fields(DataGrid.Col).Name) Then
                    Select Case UCase(VType(I))
                        Case "MENU"
                            Set object = frmMain.MenuBox(I)
                        Case "NUMERIC", "INSTRUMENT"
                            Set object = frmMain.NumberBox(I)
                        Case "TEXT"
                            Set object = frmMain.TextBox(I)
                    End Select
                    Exit For
                End If
            Next I
    End Select
    Obj.SetFocus
End If

End Sub

Private Sub DataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

If GridLoading Then Exit Sub

'If Not Gotit Then
    GridLoading = True
    

    If frmMain.PointsADO.Recordset.EOF And frmMain.PointsADO.BOFAction Then
        GridLoading = False
        Exit Sub
    End If
    Gotit = True
    On Error GoTo errorhandler
    CurrentBookMark = DataGrid.Bookmark
    frmMain.PointsADO.Recordset.Bookmark = CurrentBookMark
    frmMain.ShowValues
    Gotit = False
    'Data1.Caption = PointNo(Data1.Recordset("recno")) & " \ " & Data1.Recordset.RecordCount
    GridLoading = False
'End If
'On Error GoTo 0

Exit Sub

errorhandler:
On Error GoTo 0
GridLoading = False

End Sub

Private Sub Form_Activate()

Label2.Visible = True

End Sub

Private Sub Form_Deactivate()

Label2.Visible = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyEscape
        frmMain.Picture1.SetFocus
    Case vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd
        frmMain.Form_KeyDown KeyCode, 0
        frmMain.Picture1.SetFocus
        Exit Sub
End Select

If Shift = 2 Or Shift = 4 Then
    frmMain.SetFocus
    frmMain.Form_KeyDown KeyCode, Shift
End If
   
End Sub

Public Sub Form_Load()

IniClass = "[EDM]"
Inidata(1, 1) = "GridFields"
Call ReadIni(CFGName, IniClass, Inidata(), Status)

If GridLeft = 0 Then GridLeft = Me.Left
If GridTop = 0 Then GridTop = Me.Top
GridLoading = True

GridFieldString = Inidata(1, 2)
ParseGridFields

GridShowing = True
'Data1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + SiteDBname + ";Persist Security Info=False"
'Set Data1.Recordset = frmMain.PointsADO.Recordset
'DataGrid.Refresh
'If CountRecords > 0 Then
'    Data1.Recordset.MoveLast
'End If
For Each cfield In frmMain.PointsADO.Recordset.Fields
    If LCase(cfield.Name) <> "recno" Then
        lstFields.AddItem LCase(cfield.Name)
        If nGridFields = 0 Then
            Select Case LCase(cfield.Name)
                Case "unit", "id", "suffix", "x", "y", "z"
                    lstFields.Selected(lstFields.NewIndex) = True
            End Select
        Else
            For I = 1 To nGridFields
                If LCase(cfield.Name) = LCase(GridField(I)) Then
                    lstFields.Selected(lstFields.NewIndex) = True
                End If
            Next I
        End If
    End If
Next
lstFields.ListIndex = -1
OpenGrid
DataGrid.AllowRowSizing = False
If GridWidth < 9525 Then
    Me.Width = 9525
Else
    Me.Width = GridWidth
End If
If GridHeight < 3165 Then
    Me.Height = 3165
Else
    Me.Height = GridHeight
End If
Me.Left = GridLeft
Me.Top = GridTop
Me.Show
DataGrid.SetFocus
Me.Refresh
GridLoading = False

End Sub

Private Sub Form_Resize()

If Me.WindowState <> 1 Then

    If Me.Width < 4000 Then Me.Width = 4000
    If Me.Height < 3300 Then Me.Height = 3300
    
    DataGrid.Width = Me.Width - DataGrid.Left - 200
    If Me.Height < 3000 Then Me.Height = 3000
    DataGrid.Height = Me.Height - DataGrid.Top - 900
    Data1.Width = DataGrid.Width
    Data1.Top = DataGrid.Top + DataGrid.Height + 50
    GridFormWidth = Me.Width
    GridFormHeight = Me.Height
End If

If Me.Width > DataGrid.Width + DataGrid.Left + 300 Then
    Me.Width = DataGrid.Width + DataGrid.Left + 300
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

mdiMain.mnuDataGrid.Caption = "Show Data Grid"
GridShowing = False
TempString = GridField(1)
For I = 2 To nGridFields
    TempString = TempString + "," + GridField(I)
Next I
Inidata(1, 2) = TempString

Call WriteIni(CFGName, IniClass, Inidata(), Status)
GridWidth = Me.Width
GridHeight = Me.Height
GridLeft = Me.Left
GridTop = Me.Top

inifile$ = fixpath(App.Path) + "edm.ini"
WriteEDMIni inifile$

'Data1.Recordset.Close
'Data1.RecordSource = ""
Unload Me
'frmMain.Picture1.SetFocus

End Sub

Private Sub lstFields_Click()

If Not GridLoading Then
    OpenGrid
End If

End Sub

Private Sub lstFields_GotFocus()

Label2.Visible = False

End Sub

Public Sub ParseGridFields()

Dim X As Integer

nGridFields = 0

TempString = Trim(GridFieldString)
X = InStr(TempString, ",")
Do While X > 0
    nGridFields = nGridFields + 1
    GridField(nGridFields) = Left(TempString, X - 1)
    TempString = Trim(Mid(TempString, X + 1))
    X = InStr(TempString, ",")
Loop

If Trim(TempString) <> "" Then
    nGridFields = nGridFields + 1
    GridField(nGridFields) = Trim(TempString)
End If

End Sub

Public Sub MoveGrid()

A = 1
'GridLoading = True
'DataGrid.Visible = False
''Data1.Recordset.MoveLast
'On Error Resume Next
'Do
'    Data1.Recordset.Requery
'    Data1.Recordset.Bookmark = CurrentBookMark
'Loop Until Data1.Recordset.Bookmark = CurrentBookMark
'
'On Error GoTo 0
'DataGrid.Visible = True
'GridLoading = False

End Sub

Public Sub FormatGrid()

NFIELDS = -1
nGridFields = 0
DataGrid.Columns(0).Visible = True
For Each cfield In Data1.Recordset.Fields
    NFIELDS = NFIELDS + 1
    Gotit = False
    For I = 0 To lstFields.ListCount - 1
        If UCase(cfield.Name) = UCase(lstFields.List(I)) And lstFields.Selected(I) = True Then
            nGridFields = nGridFields + 1
            GridField(nGridFields) = lstFields.List(I)
            Gotit = True
            Exit For
        End If
    Next I
    If Not Gotit Then
        DataGrid.Columns(NFIELDS).Visible = False
    Else
        DataGrid.Columns(NFIELDS).Visible = True
        If cfield.Type = dbText Then
            DataGrid.Columns(NFIELDS).Width = Me.TextWidth(String(cfield.Size - 1, Asc("W")))
        Else
            DataGrid.Columns(NFIELDS).Width = Me.TextWidth("000000000")
        End If
    End If
Next

End Sub
