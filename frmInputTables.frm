VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmInputTables 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Import Tables"
   ClientHeight    =   3060
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   372
      Left            =   2490
      TabIndex        =   2
      Top             =   900
      Width           =   1548
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Import Tables"
      Height          =   372
      Left            =   2490
      TabIndex        =   1
      Top             =   450
      Width           =   1548
   End
   Begin VB.ListBox lstTables 
      Height          =   2310
      Left            =   168
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   2028
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   3150
      Top             =   -30
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Select tables to import by checking them, then click on the Input Tables button.  "
      Height          =   816
      Left            =   2376
      TabIndex        =   5
      Top             =   1740
      Width           =   1692
   End
   Begin VB.Label Label2 
      Height          =   345
      Left            =   180
      TabIndex        =   4
      Top             =   2940
      Width           =   3330
   End
   Begin VB.Label Label1 
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   96
      Width           =   4980
   End
End
Attribute VB_Name = "frmInputTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempDB As Database

Private Sub Command1_Click()

Unload Me
OpenSite SiteDBname

End Sub

Private Sub Command2_Click()

Dim tablename As String
Dim ReOpenDB As Boolean
Dim nTables As Integer

Set UnitTB = Nothing
frmMain.PointsADO.RecordSource = ""
Set PoleTB = Nothing
Set DatumTB = Nothing
Screen.MousePointer = 11
For I = 0 To lstTables.ListCount - 1
    If lstTables.Selected(I) = True Then
        nTables = nTables + 1
        tablename = lstTables.List(I)
        GoSub Import
    End If
Next I
Label2 = ""
OpenSite SiteDBname
Loading = True
For I = 0 To lstTables.ListCount - 1
    lstTables.Selected(I) = False
Next I
Loading = False
Screen.MousePointer = 1
MsgBox (nTables & " tables imported")

Exit Sub

Import:

If tablematch(tablename) Then
        SiteDB.TableDefs.Delete tablename
End If

Label2 = "Importing " + tablename
SqlString = "SELECT [" + tablename + "].*  INTO [" + tablename + "] IN '" + SiteDBname + "' from [" + tablename + "]"
TempDB.Execute SqlString
SiteDB.TableDefs.Refresh

On Error Resume Next
Gotit = False
Gotit = "recno" = LCase(SiteDB.TableDefs(tablename).Fields("RecNo").Name)
On Error GoTo 0
If Gotit Then
    Set MainIndex = SiteDB.TableDefs(tablename$).CreateIndex("RecordCounter")
    With MainIndex
        .Fields = "RecNo"
        .Primary = True
        .Required = True
        .Unique = True
    End With
    SiteDB.TableDefs(tablename$).Indexes.Append MainIndex
    
    SiteDB.TableDefs.Refresh
    Set MainIndex = Nothing
End If
Return

End Sub

Private Sub Form_Load()

CenterForm Me
Cd1.Filter = "Site Files (*.mdb)|*.mdb"
Cd1.CancelError = True
On Error GoTo cd1cancel
Cd1.DialogTitle = "Select Source Database"
Cd1.ShowOpen

If Cd1.filename = SiteDBname Then
    MsgBox ("You cannot import to/from the same database")
    Unload Me
    Exit Sub
End If
If Cd1.filename <> "" Then
    Set TempDB = Workspaces(0).OpenDatabase(Cd1.filename)
    Label1 = Cd1.filename
    For Each A In TempDB.TableDefs
        If Left(LCase(A.Name), 4) <> "msys" And LCase(A.Name) <> "edm_cfg" Then
            lstTables.AddItem LCase(A.Name)
        End If
    Next
End If
'Set TempDB = Nothing
If lstTables.ListCount = 0 Then
    MsgBox ("No tables available for inport")
End If
Exit Sub

cd1cancel:
Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set TempDB = Nothing
Me.Hide
frmMain.Picture1.SetFocus

End Sub

Private Sub lstTables_Click()

If Loading Then Exit Sub

Dim tablename As String
If lstTables.Selected(lstTables.ListIndex) = True Then
    tablename = lstTables.List(lstTables.ListIndex)
    If tablematch(tablename) Then
        response = MsgBox("Overwrite existing table " + tablename + "?", vbYesNo)
        If response = vbNo Then
            lstTables.Selected(lstTables.ListIndex) = False
        End If
    End If
End If

End Sub


