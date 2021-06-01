VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmNewMain 
   Caption         =   "EDM Windows"
   ClientHeight    =   7065
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10515
   Icon            =   "frmNewMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   8760
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8850
      TabIndex        =   15
      Top             =   300
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6840
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1575
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   6495
      Begin VB.ComboBox PrismMenu 
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   570
         Width           =   2115
      End
      Begin VB.Label Label6 
         Caption         =   "Z:"
         Height          =   255
         Left            =   3810
         TabIndex        =   9
         Top             =   1140
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Y:"
         Height          =   255
         Left            =   3810
         TabIndex        =   8
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "X:"
         Height          =   255
         Left            =   3810
         TabIndex        =   7
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Suffix:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Unit:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data"
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   5280
      Width           =   8175
      Begin MSDBGrid.DBGrid DBGrid1 
         Height          =   735
         Left            =   120
         OleObjectBlob   =   "frmNewMain.frx":000C
         TabIndex        =   13
         Top             =   360
         Width           =   7815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Optional Fields"
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   8175
      Begin VB.TextBox NumberBox 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox textBox 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox MenuBox 
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Shoot"
      Default         =   -1  'True
      Height          =   615
      Left            =   6840
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewDB 
         Caption         =   "New Site Database"
      End
      Begin VB.Menu mnuNewPointsTB 
         Caption         =   "New Points Table"
      End
      Begin VB.Menu space1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuOpenDB 
         Caption         =   "Open Site Database"
      End
      Begin VB.Menu mnuOpenPointsTB 
         Caption         =   "Open Points Table"
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSavePointsTBas 
         Caption         =   "Save Points Table As ...."
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save Site Database As ...."
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTheodolite 
         Caption         =   "Theodolite"
      End
      Begin VB.Menu mnuPrinter 
         Caption         =   "Printer"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditPoints 
         Caption         =   "Points"
      End
      Begin VB.Menu mnuEditPrisms 
         Caption         =   "Prisms"
      End
      Begin VB.Menu mnuEditUnits 
         Caption         =   "Units"
      End
      Begin VB.Menu mnuEditFields 
         Caption         =   "Fields"
      End
   End
   Begin VB.Menu mnuRecord 
      Caption         =   "Record"
      Begin VB.Menu mnuShootArtifact 
         Caption         =   "Artifact"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuShootTopo 
         Caption         =   "Topo"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuShootBucket 
         Caption         =   "Bucket"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuShootSample 
         Caption         =   "Sample"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuShootX 
         Caption         =   "X-shot"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuShootPlus 
         Caption         =   "+ Shot"
      End
   End
   Begin VB.Menu mnuStation 
      Caption         =   "Station"
      Begin VB.Menu mnuInitialize 
         Caption         =   "Initialize"
         Begin VB.Menu mnuDiect 
            Caption         =   "Direct"
         End
         Begin VB.Menu mnu1Reference 
            Caption         =   "1 Reference Point"
         End
         Begin VB.Menu mnu2References 
            Caption         =   "2 Reference Points"
         End
      End
      Begin VB.Menu mnuVerify 
         Caption         =   "Verify"
      End
      Begin VB.Menu mnuCreateStation 
         Caption         =   "Create"
      End
   End
   Begin VB.Menu mnuDatums 
      Caption         =   "Datums"
      Begin VB.Menu mnuCreateDatum 
         Caption         =   "Create"
      End
      Begin VB.Menu mnuEditDatums 
         Caption         =   "Edit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmNewMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuNew_Click()

End Sub


Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuDiect_Click()
frmInitialize.Show 1
End Sub

Private Sub mnuEditDatums_Click()
If DatumTableName$ <> "" Then
    frmDatumSheet.Show
Else
    MsgBox "Open or create a site first.", vbInformation
End If
End Sub

Private Sub mnuEditFields_Click()
frmFields.Show 1
End Sub

Private Sub mnuEditPoints_Click()
If PointTableName$ <> "" Then
    frmPointSheet.Show
Else
    MsgBox "Open or create a points file first.", vbInformation
End If
End Sub

Private Sub mnuEditPrisms_Click()
If PoleTableName$ <> "" Then
    frmPolesheet.Show
Else
    MsgBox "Open or create a site first.", vbInformation
End If
End Sub

Private Sub mnuEditUnits_Click()
If SiteDBname$ <> "" Then
    frmUnits.Show
Else
    MsgBox "Open or create a site first.", vbInformation
End If
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuNewDB_Click()
cd.Filter = "Site Files (*.mdb)|*.mdb"
cd.ShowSave

If cd.filename <> "" Then
    
    If Dir$(cd.filename) <> "" Then
        answer = MsgBox(cd.filename + " already exists.  Overwrite?", vbQuestion + vbYesNo)
        If answer = 7 Then Exit Sub
        On Error Resume Next
        Kill cd.filename
        If Err <> 0 Then
            MsgBox cd.filename + " could not be erased." + Chr$(13) + "Ensure that it is not already open in another application.", vbInformation + vbOKOnly
            Exit Sub
        End If
    End If
    
    Call createsitedb(cd.filename)
    
    SiteDBname$ = cd.filename
    frmMain.Caption = "EDM" + " - " + parsefilename$(cd.filename)
    
    If tablematch("Units") Then
        Set unitstb = SiteDB.OpenRecordset("Units")
    End If

End If

End Sub


Private Sub mnuNewPointsTB_Click()
If SiteDBname$ = "" Then
    MsgBox "Open or create a site file before creating a points file.", vbInformation
    Exit Sub
End If

frmPointfiles.Show 1
End Sub


Private Sub mnuOpenDB_Click()
cd.Filter = "Site Files (*.mdb)|*.mdb"
cd.ShowOpen
If cd.filename <> "" Then
    Call opensite(cd.filename)
    Call addtofilelist(dbname$)
End If
End Sub


Private Sub mnuOpenPointsTB_Click()

If SiteDBname$ = "" Then
    MsgBox "Open or create a site file before opening a points file.", vbInformation
    Exit Sub
End If

frmPointfiles.Show 1
End Sub


Private Sub mnuPrinter_Click()

Screen.MousePointer = 11
frmPrinter.Show 1
End Sub

Private Sub mnuShootArtifact_Click()
Select Case EDMName$
Case "None"
    frmManualshot.Show
Case Else
    frmTakeshot.Show
End Select
End Sub


Private Sub mnuTheodolite_Click()
Screen.MousePointer = 11
frmTheodolite.Show

End Sub


