VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConvertCFG 
   Caption         =   "Convert CFG to Windows Format"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   7230
      TabIndex        =   7
      Top             =   930
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create New CFG"
      Height          =   525
      Left            =   7230
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1320
      Top             =   1890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Dir..."
      Height          =   285
      Left            =   6450
      TabIndex        =   5
      Top             =   1080
      Width           =   525
   End
   Begin VB.TextBox txtCFG 
      Height          =   285
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1050
      Width           =   5385
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dir..."
      Height          =   285
      Left            =   6450
      TabIndex        =   2
      Top             =   690
      Width           =   525
   End
   Begin VB.TextBox txtDB 
      Height          =   285
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   690
      Width           =   5385
   End
   Begin VB.Label Label3 
      Caption         =   $"FrmStatus.frx":0000
      Height          =   465
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CFG Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1110
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Database:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   780
      Width           =   735
   End
End
Attribute VB_Name = "frmConvertCFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Gotit As Boolean
Private Sub Command1_Click()

cd.Filter = "Site Files (*.mdb)|*.mdb"
cd.filename = ""
cd.ShowSave

If cd.filename <> "" Then
    SiteDBname = cd.filename
    If Dir$(cd.filename) = "" Then
        Call createsitedb(cd.filename)
    End If
    OpenSite SiteDBname
    txtDB = LCase(SiteDBname)
End If

End Sub


Private Sub Command2_Click()
cd.Filter = "CFG Files (*.cfg)|*.cfg"
cd.filename = ""
cd.ShowSave
If cd.filename <> "" Then
    CFGName = cd.filename
    txtCFG = LCase(CFGName)
End If



End Sub


Private Sub Command3_Click()
Dim Inidata(100, 2) As String
Dim IniClass As String
Dim Status As Byte
If txtDB = "" Then
    MsgBox ("Select database before converting CFG file")
    Exit Sub
End If
If txtCFG = "" Then
    MsgBox ("Select CFG file to save converted format to")
    Exit Sub
End If

Set SiteDB = Workspaces(0).OpenDatabase(SiteDBname)
Set DatumTB = SiteDB.OpenRecordset("EDM_datums")
    On Error Resume Next
    For I = 1 To NTempDatums
        DatumTB.AddNew
        DatumTB("Name") = TempDatumName(I)
        DatumTB("x") = TempDatumX(I)
        DatumTB("y") = TempDatumY(I)
        DatumTB("z") = TempDatumZ(I)
        DatumTB.Update
    Next I
    On Error GoTo 0
Set DatumTB = Nothing
Set UnitTB = SiteDB.OpenRecordset("EDM_Units")
    On Error Resume Next
    For I = 1 To NTempUnits
        UnitTB.AddNew
        UnitTB("unit") = TempUnitName(I)
        UnitTB("minx") = TempUnitMinX(I)
        UnitTB("miny") = TempUnitMinY(I)
        UnitTB("maxx") = TempUnitMaxX(I)
        UnitTB("maxy") = TempUnitMaxY(I)
        UnitTB.Update
    Next I
    On Error GoTo 0
Set UnitTB = Nothing


Set PoleTB = SiteDB.OpenRecordset("EDM_Poles")
    On Error Resume Next
    For I = 1 To NTempPrisms
        PoleTB.AddNew
        PoleTB("Name") = TempPrismName(I)
        PoleTB("Height") = TempPrismHeight(I)
        PoleTB("offset") = TempPrismOffset(I)
        PoleTB.Update
    Next I
    On Error GoTo 0
Set PoleTB = Nothing


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
Print #1, "COMport"; comport
Print #1, "StationName="; CurrentStation.Name
Print #1, "StationX="; CurrentStation.X
Print #1, "stationY="; CurrentStation.y
Print #1, "stationZ="; CurrentStation.z
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
    Print #1, ""
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
Next I
Close 1


parsecfg A
Gotit = True
Unload Me

End Sub


Private Sub Command4_Click()
Gotit = True
Unload Me
End Sub

Private Sub Form_Load()
frmMain.ClearDBfields
Gotit = False
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not Gotit Then
    MsgBox ("The selected CFG is not in proper format.  Click on the Create New CFG button before closing, or hit the Cancel button.  If you cancel, select a new CFG file.")
    Cancel = 1
End If

End Sub

