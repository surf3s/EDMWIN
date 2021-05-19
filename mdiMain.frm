VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "EDM Windows"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   13230
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   4800
      Width           =   13230
      _ExtentX        =   23336
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   353
            MinWidth        =   353
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   13170
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   13230
      Begin VB.PictureBox UP 
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   3180
         Picture         =   "mdiMain.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   270
         TabIndex        =   0
         Top             =   120
         Width           =   330
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   330
         Index           =   3
         Left            =   90
         Picture         =   "mdiMain.frx":03C2
         ScaleHeight     =   270
         ScaleWidth      =   2865
         TabIndex        =   5
         Top             =   390
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   315
         Index           =   2
         Left            =   3570
         Picture         =   "mdiMain.frx":2C84
         ScaleHeight     =   255
         ScaleWidth      =   2865
         TabIndex        =   4
         Top             =   630
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   285
         Index           =   1
         Left            =   0
         Picture         =   "mdiMain.frx":5306
         ScaleHeight     =   225
         ScaleWidth      =   2880
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   330
         Index           =   0
         Left            =   4110
         Picture         =   "mdiMain.frx":7508
         ScaleHeight     =   270
         ScaleWidth      =   2880
         TabIndex        =   2
         Top             =   180
         Visible         =   0   'False
         Width           =   2940
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   360
      Top             =   552
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewCFG 
         Caption         =   "&New CFG"
      End
      Begin VB.Menu mnuOpenCFG 
         Caption         =   "&Open CFG"
      End
      Begin VB.Menu mnuSaveCFGas 
         Caption         =   "&Save CFG as..."
      End
      Begin VB.Menu filespace1 
         Caption         =   "-"
      End
      Begin VB.Menu createdefault 
         Caption         =   "Create Default CFG and MDB"
      End
      Begin VB.Menu filespace5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup Files"
      End
      Begin VB.Menu FileTransfer 
         Caption         =   "Transfer files to/from Pocket PC"
         Visible         =   0   'False
      End
      Begin VB.Menu Filespace4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrinter 
         Caption         =   "Setup &Printer"
         Visible         =   0   'False
      End
      Begin VB.Menu filespace2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Filelist 
         Caption         =   "filelist"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu Filelist 
         Caption         =   "filelist"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu Filelist 
         Caption         =   "Filelist"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu Filelist 
         Caption         =   "Filelist"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu Filelist 
         Caption         =   "Filelist"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu Filespace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditFields 
         Caption         =   "&Fields"
      End
      Begin VB.Menu mnuEditPrisms 
         Caption         =   "&Prisms"
      End
      Begin VB.Menu mnuEditUnits 
         Caption         =   "&Units"
      End
      Begin VB.Menu mnuCreateDatum 
         Caption         =   "&Datums"
      End
      Begin VB.Menu mnuButtons 
         Caption         =   "Shot &Buttons"
      End
      Begin VB.Menu mnuContextDependent 
         Caption         =   "&Context Dependent Defaults"
      End
      Begin VB.Menu editspace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteAll 
         Caption         =   "Delete &All points"
      End
   End
   Begin VB.Menu mnuStation 
      Caption         =   "&Station"
      Begin VB.Menu mnuTheodolite 
         Caption         =   "&Select Total Station Type and COM port settings"
      End
      Begin VB.Menu donothing2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "&Debug"
      End
      Begin VB.Menu donothing4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInitialize 
         Caption         =   "&Initialize Current Location"
      End
      Begin VB.Menu SetHangle 
         Caption         =   "Set Horizontal Angle"
      End
      Begin VB.Menu mnuStationStatus 
         Caption         =   "View Current &Coordinates"
      End
      Begin VB.Menu mnuStationVerify 
         Caption         =   "&Verify Location"
      End
      Begin VB.Menu donothing 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResetconnection 
         Caption         =   "Reset Connection"
      End
      Begin VB.Menu donothing3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetupLogFiles 
         Caption         =   "Setup log files"
      End
   End
   Begin VB.Menu mnuPlot 
      Caption         =   "&Plot"
      Begin VB.Menu mnuViewPoints 
         Caption         =   "&Points"
      End
      Begin VB.Menu mnuViewUnits 
         Caption         =   "&Units"
      End
      Begin VB.Menu mnuViewDatums 
         Caption         =   "&Datums"
      End
      Begin VB.Menu mnuViewAll 
         Caption         =   "&All"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuNewDB 
         Caption         =   "&New Site Database"
      End
      Begin VB.Menu mnuOpenDB 
         Caption         =   "&Open Site Database"
      End
      Begin VB.Menu dbspace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewPointsTB 
         Caption         =   "Select/Create &Points Table"
      End
      Begin VB.Menu dbspace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImportTables 
         Caption         =   "Import &Tables from External Database"
      End
      Begin VB.Menu mnuImportCFGField 
         Caption         =   "Import &Fields from CFG file"
      End
      Begin VB.Menu dbspace3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConvert2Newplot 
         Caption         =   "&Convert Database to Newplot Format"
         Visible         =   0   'False
      End
      Begin VB.Menu dbspace4 
         Caption         =   "-"
      End
      Begin VB.Menu DBRefresh 
         Caption         =   "&Refresh Database"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuspeech 
         Caption         =   "Speak Unit-ID on new shots"
      End
      Begin VB.Menu mnuFindUnit 
         Caption         =   "&Auto-Find Unit"
      End
      Begin VB.Menu mnuPrismPrompt 
         Caption         =   "&Prompt for Prism"
      End
      Begin VB.Menu mnuPrintShots 
         Caption         =   "P&rint Shots"
      End
      Begin VB.Menu mnuNoAlert 
         Caption         =   "No Update Alert"
      End
      Begin VB.Menu mnuSelectedUnit 
         Caption         =   "mnuSelectedUnit"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUpperCase 
         Caption         =   "UpperCase all Entries"
      End
      Begin VB.Menu mnuDataGrid 
         Caption         =   "Show Data Grid"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Show only Points from Current Station"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu HelpStatus 
         Caption         =   "&Status"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub createdefault_Click()

frmDefault.Show 1

End Sub

Private Sub DBRefresh_Click()

Screen.MousePointer = 11
CurrentPosition = frmMain.PointsADO.Recordset.Bookmark
frmMain.PointsADO.Refresh
' frmMain.PointsADO.Recordset.MoveLast
On Error GoTo BadBookMark
frmMain.PointsADO.Recordset.Bookmark = CurrentPosition

Start:
On Error GoTo 0
frmMain.txtTotalRecords = frmMain.PointsADO.Recordset.RecordCount

If GridShowing Then
    frmDataGrid.FormatGrid
End If
If PlotShowing Then
    frmPlot.PlotPoints
End If
Screen.MousePointer = 1

Exit Sub

BadBookMark:
frmMain.PointsADO.Recordset.MoveLast
GoTo Start

End Sub

Private Sub FileList_Click(Index As Integer)

CFGName = Filelist(Index).Caption
Cancelling = False
parsecfg A


'put code in here to rewrite cfg if necessary
If Cancelling Then
    Cancelling = False
    frmMain.FormatVarList
    Exit Sub
End If
LastPath = GetPath(CFGName)

End Sub

Private Sub FileTransfer_Click()

frmCEupdownload.Show 1

End Sub

Private Sub HelpStatus_Click()

frmStatus.Show 1

End Sub

Private Sub MDIForm_Load()

Me.Show
frmMain.Show

End Sub

Private Sub MDIForm_Resize()

For A = 1 To 7
    mdiMain.StatusBar.Panels(A).Width = mdiMain.Width / 7
Next A

End Sub

Private Sub mnuAbout_Click()

frmAbout.Show 1

End Sub

Private Sub mnuBackup_Click()

Dim TempGrid As Boolean
Dim TempDB As Database
Dim Tbl As TableDef

Start:
cd.CancelError = True
On Error GoTo cdcancel

cd.Filter = "*.mdb"
cd.filename = Mid(SiteDBname, Len(GetPath(SiteDBname)) + 2)
cd.DefaultExt = "mdb"
cd.InitDir = BackupFolder
cd.DialogTitle = "Select Backup Folder"
cd.ShowSave
On Error GoTo 0
If cd.filename = SiteDBname Then
    MsgBox ("You cannot copy a data file to itself.  Select new name or new destination folder")
    Exit Sub
End If

A = Dir(cd.filename)
If A <> "" Then
    response = MsgBox("Replace files?", vbYesNoCancel)
    If response = vbYes Then
        On Error GoTo BadFile
        Set TempDB = Workspaces(0).OpenDatabase(cd.filename)
        On Error GoTo 0
    Else
        Exit Sub
    End If
Else
    Set TempDB = Workspaces(0).CreateDatabase(cd.filename, dbLangGeneral)
End If
Screen.MousePointer = 11
X = InStr(cd.filename, cd.FileTitle)
BackupFolder = Left(cd.filename, X - 1)

For Each Tbl In SiteDB.TableDefs
    If Left(LCase(Tbl.Name), 4) <> "msys" Then
        Gotit = False
        For I = 0 To TempDB.TableDefs.Count - 1
            If Trim(UCase$(TempDB.TableDefs(I).Name)) = Trim(UCase$(Tbl.Name)) Then
                Gotit = True
                Exit For
            End If
        Next I
        If Gotit Then
            TempDB.TableDefs.Delete Tbl.Name
        End If
        SqlString = "SELECT [" + Tbl.Name + "].*  INTO [" + Tbl.Name + "] IN '" + cd.filename + "' from [" + Tbl.Name + "]"
        SiteDB.Execute SqlString
    End If
Next
MsgBox ("If you have backed up to a memory key, you must exit the program before ejecting")

Set TempDB = Nothing
Set Tbl = Nothing
Screen.MousePointer = 1
Exit Sub

BadFile:
MsgBox ("Error: " + Err.Description + ".  File must be valid Access database.")
Resume Start

cdcancel:

End Sub

Private Sub mnuButtons_Click()

If CFGName = "" Then
    MsgBox ("Open or Create CFG before performing this operation")
    Exit Sub
End If
frmButtons.Show 1

End Sub

Private Sub mnuContextDependent_Click()

If SiteDBOpen = False Then
    MsgBox ("Open Site Database first")
    Exit Sub
End If
frmSubUnits.Show

End Sub

Private Sub mnuConvert2Newplot_Click()

If CFGName = "" Then
    MsgBox ("Must open valid CFG file before performing this operation")
    Exit Sub
End If
If SiteDBname = "" Then
    MsgBox ("Must open database before performing this operation")
    Exit Sub
End If
Screen.MousePointer = 11
If Not tablematch("context") Then
    CreateContext
End If
If Not tablematch("xyz") Then
    CreateXYZ
End If
MsgBox ("Done")
Screen.MousePointer = 1
mnuConvert2Newplot.Enabled = False

End Sub

Private Sub mnuCreateDatum_Click()

If SiteDBname = "" Then
    MsgBox ("You must open site database before defining datums")
    Exit Sub
End If
frmDatumSheet.Show 1

End Sub

Public Sub mnuDataGrid_Click()

Dim Inidata(1, 2) As String
Dim IniClass As String
Dim Status As Byte

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then
    MsgBox ("Open points file before showing grid.")
    Exit Sub
End If
If mnuDataGrid.Caption = "Show Data Grid" Then
    mnuDataGrid.Caption = "Hide Data Grid"
    frmDataGrid.Show
    frmDataGrid.SetFocus
    Inidata(1, 2) = "YES"

Else
    mnuDataGrid.Caption = "Show Data Grid"
    Unload frmDataGrid
    Inidata(1, 2) = "NO"

End If

IniClass = "[EDM]"
Inidata(1, 1) = "ShowGrid"
Call WriteIni(CFGName, IniClass, Inidata(), Status)

End Sub

Private Sub mnuDebug_Click()

frmDebug.Show

End Sub

Private Sub mnuDeleteAll_Click()

If SiteDBname = "" Then
    MsgBox "Open database and points table before performing this operation.", vbInformation
    Exit Sub
End If
    
If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then
    MsgBox "Open points table before performing this operation.", vbInformation
    Exit Sub
End If

If frmMain.PointsADO.Recordset.EOF And frmMain.PointsADO.Recordset.BOF Then
    Exit Sub
End If

response = MsgBox("Warning: This action will permanently remove all records from the current points table.  Continue anyway?", vbYesNo)
If response = vbNo Then Exit Sub

With frmMain.PointsADO.Recordset
    .MoveFirst
    Do
        .Delete
        If Not .EOF Then .MoveNext
    Loop Until .EOF
End With

ClearFields
frmMain.PointsADO.Recordset.Requery

'OpenPointsTable
'frmMain.ShowValues

End Sub

Private Sub mnuEditFields_Click()

If CFGName = "" Then
    MsgBox ("Open or Create CFG before performing this operation")
    Exit Sub
End If
frmEditField.Show

End Sub

Public Sub mnuEditPrisms_Click()

Dim TempPrism As String

If CFGName = "" Then
    MsgBox ("Open a valid CFG file and database first.")
    Exit Sub
End If
If SiteDBname = "" Then
    MsgBox ("Open Site database first")
    Exit Sub
End If

TempPrism = frmMain.txtPrism

frmPolesheet.Show 1
frmMain.txtPrism.Clear
If PoleTB.RecordCount = 0 Then
    frmMain.lblPoleWarning.Visible = True
Else
    frmMain.lblPoleWarning.Visible = False
    Gotit = False
    PoleTB.MoveFirst
    If Not PoleTB.EOF Then
        PoleTB.MoveFirst
        While Not PoleTB.EOF
            nPoleHeights = nPoleHeights + 1
            frmMain.txtPrism.AddItem PoleTB("Name")
            frmMain.txtPrism.ItemData(frmMain.txtPrism.NewIndex) = nPoleHeights
            If IsNull(PoleTB("height")) Then
                PoleTB.Edit
                PoleTB("Height") = 0
                PoleTB.Update
                Gotit = True
            End If
            
            PoleHeight(nPoleHeights) = PoleTB("height")
            If IsNull(PoleTB("offset")) Then
                PoleTB.Edit
                PoleTB("offset") = 0
                PoleTB.Update
                Gotit = True
            End If
            PoleOffset(nPoleHeights) = PoleTB("offset")
            PoleTB.MoveNext
        Wend
        If Gotit Then
            MsgBox ("Invalid values found in Prisms table.  Verify prism heights and offsets")
        End If
        Gotit = False
        Loading = True
        For I = 0 To frmMain.txtPrism.ListCount - 1
            If UCase(TempPrism) = UCase(frmMain.txtPrism.List(I)) Then
                frmMain.txtPrism.ListIndex = I
                Gotit = True
                Exit For
            End If
        Next I
        If Not Gotit Then
            If frmMain.txtPrism.ListCount > 0 Then
                frmMain.txtPrism.ListIndex = 0
            End If
        End If
        Loading = False
        
    End If
End If
    
End Sub

Private Sub mnuEditUnits_Click()

If CFGName = "" Then
    MsgBox ("Open CFG and Database before performing this operation")
    Exit Sub
End If
If SiteDBname$ <> "" Then
    AddUnits.Show 1
Else
    MsgBox "Open or create a site first.", vbInformation
End If

End Sub

Private Sub mnuExit_Click()

If GridShowing Then
    Unload frmDataGrid
End If
If PlotShowing Then
    Unload frmPlot
End If

Unload frmMain
Unload Me
End

End Sub

Private Sub mnuFilter_Click()

'If mnuFilter.Checked Then
'    mnuFilter.Checked = False
'    PointsTb.Filter = ""
'
'Else
'    mnuFilter.Checked = True
'    PointsTb.Filter = "datumname='" + CurrentStation.Name + "'"
'End If
'
'Dim Inidata(1, 2) As String
'Dim IniClass As String
'Dim Status As Byte
'IniClass = "[EDM]"
'Inidata(1, 1) = "FilterPoints"
'If mnuFilter.Checked = True Then
'    Inidata(1, 2) = "YES"
'Else
'    Inidata(1, 2) = "NO"
'End If
'
'Call WriteIni(CFGName, IniClass, Inidata(), Status)

End Sub

Public Sub mnuFindUnit_Click()

If CFGName = "" Or SiteDBname = "" Then
    MsgBox ("Open database before using this option.")
    Exit Sub
End If

If mnuFindUnit.Checked Then
    mnuFindUnit.Checked = False
    LimitChecking = False
    frmMain.lblAutoFind.Visible = False
Else
    If UnitTB.RecordCount < 1 Then
        MsgBox ("No units defined -- Auto-find Unit option is not available")
        Exit Sub
    End If
    
    mnuFindUnit.Checked = True
    LimitChecking = True
    frmMain.lblAutoFind.Visible = True
End If

Dim Inidata(1, 2) As String
Dim IniClass As String
Dim Status As Byte
IniClass = "[EDM]"
Inidata(1, 1) = "Limitchecking"
If LimitChecking Then
    Inidata(1, 2) = "YES"
Else
    Inidata(1, 2) = "NO"
End If

Call WriteIni(CFGName, IniClass, Inidata(), Status)

End Sub

Private Sub mnuImportCFGField_Click()

If CFGName = "" Then
    MsgBox ("You must open or create a CFG file before importing fields.")
    Exit Sub
End If

cd.Filter = "CFG Files (*.cfg)|*.cfg"
cd.CancelError = True
On Error GoTo cdcancel
cd.DialogTitle = "Select Source CFG file"
cd.ShowOpen
On Error GoTo 0
If cd.filename <> "" Then
    frmImportFields.DonorCFG = cd.filename
    frmImportFields.Show
End If

cdcancel:
End Sub

Private Sub mnuImportTables_Click()

If SiteDBname = "" Then
    MsgBox ("Open database before performing this operation")
    Exit Sub
End If

On Error Resume Next
frmInputTables.Show
On Error GoTo 0

End Sub

Public Sub mnuInitialize_Click()

If SiteDBname = "" Then
    MsgBox ("Before you can initialize the station,  you need to open or create a site database.  Use Database Open or New first.")
    Exit Sub
End If

If frmMain.lblEDMWarning.Visible = True Then
    MsgBox ("The type of total station has not yet been set.  Use Station Select Total Station first.")
    Exit Sub
End If

If DatumTB.BOF And DatumTB.EOF Then
    MsgBox ("There are no datums defined.  Use Edit Datums first to add new datums and then initialize the station.")
    Exit Sub
End If
If UCase(EDMName) = "MICROSCRIBE" Then
    frmMSSetup.Show
Else
    frmStationSetup.Show 1
End If

End Sub

Private Sub mnuNewCFG_Click()

If CFGName <> "" Then
    response = MsgBox("Retain current settings and fields?", vbYesNo)
    If response = vbNo Then
        For I = 1 To Vars
            VarList(I) = ""
            VCarry(I) = False
            VMenu(I) = ""
            VLen(I) = 0
            VType(I) = ""
        Next I
        Vars = 7
        VarList(1) = "UNIT"
        VPrompt(1) = "Unit"
        VLen(1) = 6
        VType(1) = "TEXT"
        VarList(2) = "ID"
        VPrompt(2) = "ID"
        VLen(2) = 5
        VType(2) = "TEXT"
        VarList(3) = "SUFFIX"
        VPrompt(3) = "Suffix"
        VLen(3) = 10
        VType(3) = "NUMERIC"
        VarList(4) = "PRISM"
        VPrompt(4) = "Prism"
        VLen(4) = 10
        VType(4) = "NUMERIC"
        VarList(5) = "X"
        VPrompt(5) = "X"
        VLen(5) = 10
        VType(5) = "NUMERIC"
        VarList(6) = "Y"
        VPrompt(6) = "Y"
        VLen(6) = 10
        VType(6) = "NUMERIC"
        VarList(7) = "Z"
        VPrompt(7) = "Z"
        VLen(7) = 10
        VType(7) = "NUMERIC"
        mnuFindUnit.Checked = False
        frmMain.lblAutoFind.Visible = False
        LimitChecking = False
        frmMain.ClearDBfields
        For I = 1 To 6
            frmMain.Button(I).Visible = False
            nButtonVars(I) = 0
        Next I

        PointTableName = ""
        SiteDBname = ""
        DBName = ""
        DBPath = ""
        SqidCheck = False
        UnitFieldString = ""
    End If
End If
Set SiteDB = Nothing
frmMain.PointsADO.RecordSource = ""

Start:
cd.CancelError = True
On Error GoTo cdcancel
cd.Filter = "CFG Files (*.cfg)|*.cfg"
cd.DialogTitle = "Create New CFG file"
cd.filename = CFGName
cd.ShowSave
If cd.filename = "" Then Exit Sub
A = Dir(cd.filename)
If A <> "" Then
    response = MsgBox("Overwrite existing file?", vbYesNo)
    If response = vbNo Then
        GoTo Start
    End If
End If

CFGName = cd.filename

frmEditField.Show 1

cdcancel:

End Sub

Private Sub mnuNewDB_Click()

If CFGName = "" Then
    MsgBox ("You must open or create a CFG file before creating a database.")
    Exit Sub
End If
response = MsgBox("Create new database based on " + CFGName + "?", vbYesNo)
If response = vbNo Then Exit Sub
Loading = True
cd.filename = Left(CFGName, Len(CFGName) - 4)

cd.CancelError = True
On Error GoTo cdcancel

cd.Filter = "Site Files (*.mdb)|*.mdb"
cd.DialogTitle = "Create New Database"
Start:
cd.ShowSave

If cd.filename <> "" Then
    If Len(cd.FileTitle) > 12 Then
        response = MsgBox("If you plan to use this database with EDM CE (Pocket PC), database names cannot be longer than 8 characters.  Continue?", vbYesNo)
        If response = vbNo Then
            GoTo Start
        End If
    End If
    
    If Dir$(cd.filename) <> "" Then
        answer = MsgBox(cd.filename + " already exists.  Overwrite?", vbQuestion + vbYesNo)
        If answer = 7 Then Exit Sub
        frmMain.ClearDBfields
        On Error Resume Next
        
        Kill cd.filename
        If Err <> 0 Then
            MsgBox cd.filename + " could not be erased." + Chr$(13) + "Ensure that it is not already open in another application.", vbInformation + vbOKOnly
            Exit Sub
        End If
    End If
    
    frmMain.ClearDBfields
    Call createsitedb(cd.filename)
    SiteDBname = cd.filename
    
    OpenSite SiteDBname
    
    mdiMain.StatusBar.Panels(4) = "DB: " + LCase(DBName) + "   "
    response = MsgBox("Would you like to import tables (prisms, units, datums, etc) from an existing database?", vbYesNo)
    If response = vbYes Then
        mnuImportTables_Click
    End If
    
End If
cdcancel:

End Sub

Public Sub mnuNewPointsTB_Click()

If SiteDBname$ = "" Then
    MsgBox "Open or create a site database before creating a points file.", vbInformation
    Exit Sub
End If
If CFGName = "" Or Vars = 0 Then
    MsgBox ("Open or create a CFG file with fields before creating a points file")
    Exit Sub
End If
frmPointfiles.Show 1

If Cancelling Then Exit Sub
    
txtPT = PointTableName
Dim Inidata(1, 2) As String
Dim IniClass As String
Dim Status As Byte
IniClass = "[EDM]"
Inidata(1, 1) = "PointTable"
Inidata(1, 2) = PointTableName
Call WriteIni(CFGName, IniClass, Inidata(), Status)

End Sub

Public Sub mnuNoAlert_Click()

If mnuNoAlert.Checked Then
    mnuNoAlert.Checked = False
    NoAlert = False
Else
    mnuNoAlert.Checked = True
    NoAlert = True
End If
Dim Inidata(1, 2) As String
Dim IniClass As String
Dim Status As Byte
IniClass = "[EDM]"
Inidata(1, 1) = "UpdateAlerts"
If Not NoAlert Then
    Inidata(1, 2) = "YES"
Else
    Inidata(1, 2) = "NO"
End If

Call WriteIni(CFGName, IniClass, Inidata(), Status)

End Sub

Public Sub mnuOpenCFG_Click()

Start:
cd.CancelError = True
On Error GoTo cdcancel

cd.Filter = "CFG Files (*.cfg)|*.cfg"
cd.DialogTitle = "Open existing CFG file"
cd.ShowOpen
If cd.filename <> "" Then
    CFGName = cd.filename
    parsecfg A
    LastPath = GetPath(cd.filename)
End If
inifile$ = fixpath(App.Path) + "edm.ini"
Call WriteEDMIni(inifile$)
cdcancel:

End Sub

Public Sub mnuOpenDB_Click()

If CFGName = "" Then
    MsgBox ("You must open or create a CFG file before opening a database.")
    Exit Sub
End If

Dim A As Integer

cd.CancelError = True

If cd.filename <> "" Then
    If Left$(cd.filename, 3) <> "mdb" Then
        A = InStr(cd.filename, ".")
        If A <> 0 Then
            cd.filename = Left(cd.filename, A) + "mdb"
        End If
    End If
End If
Start:
On Error GoTo cdcancel

cd.Filter = "Site Files (*.mdb)|*.mdb"
cd.DialogTitle = "Open Existing Database"
cd.ShowOpen
On Error GoTo 0
If cd.filename <> "" Then
    If Len(cd.FileTitle) > 12 Then
        response = MsgBox("If you plan to use this database with EDM CE (Pocket PC), database names cannot be longer than 8 characters.  Continue?", vbYesNo)
        If response = vbNo Then
            GoTo Start
        End If
    End If
    If Dir(cd.filename) = "" Then
        response = MsgBox("Create new database based on " + CFGName + "?", vbYesNo)
        If response = vbNo Then Exit Sub
        PointTableName = ""
        'txtCurrentRecord = 0
        frmMain.txtTotalRecords = 0
        Loading = True
        frmMain.ClearDBfields
        SiteDBname = cd.filename
        Call createsitedb(SiteDBname)
        OpenSite SiteDBname
        mdiMain.StatusBar.Panels(4) = "DB: " + LCase(cd.filename) + "   "
    Else
        frmMain.ClearDBfields
        PointTableName = ""
        'txtCurrentRecord = 0
        frmMain.txtTotalRecords = 0
        SiteDBname$ = cd.filename
        Call OpenSite(SiteDBname$)
    End If
End If
mdiMain.StatusBar.Panels(4) = "DB: " + LCase(DBName) + "   "
cdcancel:

End Sub

Private Sub mnuPrinter_Click()

Screen.MousePointer = 11
frmPrinter.Show 1

End Sub

Private Sub mnuPrintShots_Click()

frmPrinter.Show 1

End Sub

Private Sub mnuPrismPrompt_Click()

If mnuPrismPrompt.Checked Then
    mnuPrismPrompt.Checked = False
Else
    mnuPrismPrompt.Checked = True
End If

Dim Inidata(1, 2) As String
Dim IniClass As String
Dim Status As Byte
IniClass = "[EDM]"
Inidata(1, 1) = "PrismPrompt"
If mnuPrismPrompt.Checked Then
    Inidata(1, 2) = "YES"
Else
    Inidata(1, 2) = "NO"
End If

Call WriteIni(CFGName, IniClass, Inidata(), Status)

End Sub

Private Sub mnuResetconnection_Click()

Select Case UCase(EDMName$)
Case "TOPCON", "WILD", "SOKKIA", "WILD2"
    If comport$ <> "" And comsettings <> "" Then
        Call initcomport(comport$, errorcode)
        MsgBox "The connection is reset.  Use this option when the computer powers off while the program is running.", vbOKOnly
    Else
        MsgBox "The COM port and its settings are not set.  Do Station Select Station Type to set these.", vbInformation
    End If
Case "SIMULATE"
    MsgBox ("A simulated station cannot be reset")
    
Case Else
    MsgBox "A valid total station type has not yet been selected so the connection can not be reset.", vbInformation
End Select

End Sub

Private Sub mnuSaveCFGas_Click()

If CFGName = "" Then
    MsgBox ("You must open or create a CFG file before creating a database.")
    Exit Sub
End If

Start:
cd.CancelError = True
On Error GoTo cdcancel
cd.Filter = "CFG Files (*.cfg)|*.cfg"
cd.DialogTitle = "Save CFG file as ...."
cd.filename = CFGName
cd.ShowSave
If cd.filename = "" Then Exit Sub
A = Dir(cd.filename)
If A <> "" Then
    response = MsgBox("Overwrite existing file?", vbYesNo)
    If response = vbNo Then
        GoTo Start
    End If
End If

CFGName = cd.filename

cdcancel:

End Sub

Private Sub mnuSelectedUnit_Click()

If mnuSelectedUnit.Checked = True Then
    frmMain.SelectedUnit = ""
    mnuSelectedUnit.Visible = False
End If

End Sub

Private Sub mnuSetupLogFiles_Click()

frmSetupLogFiles.Show 1

End Sub

Private Sub mnuspeech_Click()

If mnuspeech.Checked = True Then
    mnuspeech.Checked = False
    Speaking = False
Else
    Speaking = True
    On Error GoTo SAPINotFound
    Set Voice = New SpVoice
    Voice.Speak ("Speaking")
    On Error GoTo 0
    mnuspeech.Checked = True
    
    Exit Sub
    
SAPINotFound:
    If Err.Number = 459 Or Err.Number = 429 Then
        MsgBox "SAPI.dll (for speaking option) not found."
    Else
        MsgBox "Error encountered : " & Err.Number
    End If
    Speaking = False
End If

End Sub

Private Sub mnuStationStatus_Click()
    
If Not StationInitialized Then
    MsgBox ("Station Not Initialized")
Else
    TempString = "Current station: " + Trim(CurrentStation.Name) + Chr(13)
    TempString = TempString + "   X: " + Format(CurrentStation.X, "#####0.000") + Chr(13)
    TempString = TempString + "   Y: " + Format(CurrentStation.y, "#####0.000") + Chr(13)
    TempString = TempString + "   Z: " + Format(CurrentStation.z, "#####0.000") + Chr(13)
    MsgBox (TempString)
End If

End Sub

Private Sub mnuStationVerify_Click()

If Not StationInitialized Then
    MsgBox ("Station Not Initialized")
Else
    frmStationVerify.Show 1
End If

End Sub

Public Sub mnuTheodolite_Click()

Screen.MousePointer = 1
frmTheodolite.Show

End Sub

Private Sub mnuUpperCase_Click()

Dim TempString As String

If mnuUpperCase.Checked Then
    mnuUpperCase.Checked = False
    UpperCase = False
Else
    mnuUpperCase.Checked = True
    UpperCase = True
End If
Dim Inidata(1, 2) As String
Dim IniClass As String
Dim Status As Byte
IniClass = "[EDM]"
Inidata(1, 1) = "UpperCase"
If Not UpperCase Then
    Inidata(1, 2) = "no"
Else
    Inidata(1, 2) = "yes"
End If

Call WriteIni(CFGName, IniClass, Inidata(), Status)

If UpperCase Then
    For I = 1 To Vars
        If UCase(VType(I)) = "MENU" Then
            VMenu(I) = UCase(VMenu(I))
            MenuString = VMenu(I)
            TempString = frmMain.MenuBox(I)
            frmMain.MenuBox(I).Clear
            Gotit = False
            Do Until Gotit
                X = InStr(MenuString, ",")
                If X > 0 Then
                    frmMain.MenuBox(I).AddItem Left(MenuString, X - 1)
                    MenuString = Mid(MenuString, X + 1)
                Else
                    frmMain.MenuBox(I).AddItem MenuString
                    Gotit = True
                End If
            Loop
            frmMain.MenuBox(I) = TempString
        End If
    Next I
End If

End Sub

Private Sub mnuViewAll_Click()

If SiteDBname = "" Then
    MsgBox ("Open database before performing this operation")
    Exit Sub
End If
If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then
    MsgBox ("Open points table before performing this operation")
    Exit Sub
End If

If mnuViewAll.Caption = "&All" Then
    mnuViewDatums.Checked = True
    mnuViewUnits.Checked = True
    mnuViewPoints.Checked = True
    mnuViewAll.Caption = "&None"
    frmPlot.SetScale
    frmPlot.PlotPoints
    frmPlot.Show

Else
    mnuViewDatums.Checked = False
    mnuViewUnits.Checked = False
    mnuViewPoints.Checked = False
    mnuViewAll.Caption = "&All"
    Unload frmPlot
End If

End Sub

Public Sub mnuViewDatums_Click()

If SiteDBname = "" Then
    MsgBox ("Open database before performing this operation")
    Exit Sub
End If

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then
    MsgBox ("Open points table before performing this operation")
    Exit Sub
End If

If mnuViewDatums.Checked Then
    mnuViewDatums.Checked = False
    If mnuViewPoints.Checked = False And mnuViewUnits.Checked = False Then
        mnuViewAll.Caption = "All"
    Else
        mnuViewAll.Caption = "None"
    End If
    
    If mnuViewPoints.Checked Or mnuViewUnits.Checked Then
        frmPlot.SetScale
        frmPlot.PlotPoints
    Else
        Unload frmPlot
    End If
Else
    mnuViewDatums.Checked = True
    mnuViewAll.Caption = "None"
    frmPlot.Show
    frmPlot.SetScale
    frmPlot.PlotPoints
End If

End Sub

Private Sub mnuViewPoints_Click()

If SiteDBname = "" Then
    MsgBox ("Open database before performing this operation")
    Exit Sub
End If

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then
    MsgBox ("Open points table before performing this operation")
    Exit Sub
End If

If mnuViewPoints.Checked Then
    mnuViewPoints.Checked = False
    If mnuViewDatums.Checked = False And mnuViewUnits.Checked = False Then
        mnuViewAll.Caption = "All"
    Else
        mnuViewAll.Caption = "None"
    End If
    If mnuViewDatums.Checked Or mnuViewUnits.Checked Then
        frmPlot.SetScale
        frmPlot.PlotPoints
    Else
        Unload frmPlot
    End If
Else
    mnuViewPoints.Checked = True
    mnuViewAll.Caption = "None"
    frmPlot.Show
    frmPlot.SetScale
    frmPlot.PlotPoints
End If

End Sub

Private Sub mnuViewUnits_Click()

If SiteDBname = "" Then
    MsgBox ("Open database before performing this operation")
    Exit Sub
End If

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then
    MsgBox ("Open points table before performing this operation")
    Exit Sub
End If

If mnuViewUnits.Checked Then
    mnuViewUnits.Checked = False
    If mnuViewPoints.Checked = False And mnuViewDatums.Checked = False Then
        mnuViewAll.Caption = "All"
    Else
        mnuViewAll.Caption = "None"
    End If
    If mnuViewDatums.Checked Or mnuViewPoints.Checked Then
        frmPlot.SetScale
        frmPlot.PlotPoints
    Else
        Unload frmPlot
    End If
Else
    mnuViewUnits.Checked = True
    mnuViewAll.Caption = "None"
    frmPlot.Show
    frmPlot.SetScale
    frmPlot.PlotPoints
End If

End Sub

Private Sub SetHangle_Click()

frmHorizAngle.Show 1

End Sub

Public Sub save_cfgfile()

mdiMain.StatusBar.Panels(3) = "CFG: " + CFGName + "  "
Open CFGName For Output As 1
Print #1, "[EDM]"
Print #1, "Database="; DBName
Print #1, "DBPath="; DBPath
Print #1, "PointTable="; PointTableName
If SqidCheck = True Then
    Print #1, "SQID=YES"
Else
    Print #1, "SQID=NO"
End If
Print #1, "Unitfields="; UnitFieldString
If LimitChecking Then
    Print #1, "Limitchecking=Yes"
Else
    Print #1, "Limitchecking=No"
End If
If NoAlert Then
    Print #1, "UpdateAlerts=No"
Else
    Print #1, "UpdateAlerts=YES"
End If
If mdiMain.mnuPrismPrompt.Checked = True Then
    Print #1, "PrismPrompt=YES"
Else
    Print #1, "PrismPrompt=NO"
End If

Print #1, "Instrument="; EDMName
Print #1, "COMport="; comport
Print #1, "EdmDelayTime="; EDMDelayTime
Print #1, "StationName="; CurrentStation.Name
Print #1, "StationX="; CurrentStation.X
Print #1, "stationY="; CurrentStation.y
Print #1, "stationZ="; CurrentStation.z
Print #1, ""
For I = 1 To 6
    If nButtonVars(I) > 0 Then
        Print #1, "[BUTTON" + Trim(Str(I)) + "]"
        Print #1, "TITLE="; ButtonCaption(I)
        For J = 1 To nButtonVars(I)
            Print #1, VarList(ButtonVars(I, J, 1)) + "=" + ButtonVars(I, J, 2)
        Next J
        Print #1, ""
    End If
Next I

For I = 1 To Vars
    Print #1, "[" + VarList(I) + "]"
    Print #1, "Prompt="; VPrompt(I)
    Print #1, "Length="; VLen(I)
    Print #1, "Type="; VType(I)
    If VType(I) = "MENU" Then
        Print #1, "Menu=" + VMenu(I)
    End If
    If VCarry(I) Then
        Print #1, "Carry=True"
    End If
    Print #1, ""
Next I
Close 1

End Sub

