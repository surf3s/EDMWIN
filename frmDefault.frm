VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDefault 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create a Default CFG"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox comportmenu 
      Height          =   315
      Left            =   2760
      TabIndex        =   27
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4560
      TabIndex        =   26
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write"
      Height          =   495
      Left            =   2520
      TabIndex        =   25
      Top             =   5280
      Width           =   1215
   End
   Begin VB.ComboBox stopbitmenu 
      Height          =   315
      Left            =   7200
      TabIndex        =   22
      Top             =   3840
      Width           =   735
   End
   Begin VB.ComboBox databitmenu 
      Height          =   315
      Left            =   6360
      TabIndex        =   20
      Top             =   3840
      Width           =   735
   End
   Begin VB.ComboBox paritymenu 
      Height          =   315
      Left            =   5160
      TabIndex        =   18
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ComboBox baudratemenu 
      Height          =   315
      Left            =   3960
      TabIndex        =   16
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ComboBox totalstationmenu 
      Height          =   315
      ItemData        =   "frmDefault.frx":0000
      Left            =   600
      List            =   "frmDefault.frx":0002
      TabIndex        =   14
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox pointsfiletext 
      Height          =   285
      Left            =   600
      TabIndex        =   8
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox databasetext 
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   7335
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   7440
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox cfgfiletext 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   7335
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COM Port :"
      Height          =   195
      Left            =   2760
      TabIndex        =   28
      Top             =   3600
      Width           =   780
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDefault.frx":0004
      Height          =   615
      Left            =   600
      TabIndex        =   24
      Top             =   4440
      Width           =   7455
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5."
      Height          =   195
      Left            =   360
      TabIndex        =   23
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stop Bits :"
      Height          =   195
      Left            =   7200
      TabIndex        =   21
      Top             =   3600
      Width           =   720
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Bits :"
      Height          =   195
      Left            =   6360
      TabIndex        =   19
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parity :"
      Height          =   195
      Left            =   5160
      TabIndex        =   17
      Top             =   3600
      Width           =   480
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Baud Rate :"
      Height          =   195
      Left            =   3960
      TabIndex        =   15
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Station Type :"
      Height          =   195
      Left            =   600
      TabIndex        =   13
      Top             =   3600
      Width           =   1395
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4."
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   3840
      Width           =   135
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3."
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2."
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Points Table :"
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Double-click inside box to browse)"
      Height          =   195
      Left            =   2640
      TabIndex        =   6
      Top             =   2400
      Width           =   2460
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database Name :"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Double-click inside box to browse)"
      Height          =   195
      Left            =   2640
      TabIndex        =   3
      Top             =   1440
      Width           =   2460
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CFG Filename :"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDefault.frx":0118
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "frmDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cfgfiletext_DblClick()

cd.ShowSave
If cd.filename <> "" Then
    cfgfiletext.Text = cd.filename
End If

End Sub

Private Sub Command1_Click()

If cfgfiletext.Text = "" Then
    MsgBox "Please supply a CFG filename.", vbInformation
    Exit Sub
End If

If databasetext.Text = "" Then
    MsgBox "Please supply a database filename.", vbInformation
    Exit Sub
End If

If pointsfiletext.Text = "" Then
    MsgBox "Please supply a points table name.", vbInformation
    Exit Sub
End If

If totalstationmenu.Text <> "Simulate" And (comportmenu.Text = "" Or baudratemenu.Text = "" Or paritymenu.Text = "" Or databitmenu.Text = "" Or stopbitmenu.Text = "") Then
    MsgBox "If the total station type is not set to Simulation, the baud rate, parity, data bits and stop bits must be set.", vbInformation
    Exit Sub
End If

On Error Resume Next
If Dir$(cfgfiletext.Text) <> "" Then
    If MsgBox("Overwrite existing file " + cfgfiletext.Text + "?", vbYesNo) = vbNo Then Exit Sub
End If

If Dir$(databasetext.Text) <> "" Then
    If MsgBox("Overwrite existing database " + databasetext.Text + "?", vbYesNo) = vbNo Then Exit Sub
End If

If LCase(Left(cfgfiletext.Text, Len("\my documents\edmwin\"))) = "\my documents\edmwin\" Then
    MkDir "\My Documents"
    MkDir "\My Documents\EDMWIN"
End If
Err.Clear

fno = FreeFile
Open cfgfiletext.Text For Output As #fno
Close fno
If Err <> 0 Then
    Call MsgBox("Could not create " + cfgfiletext.Text + ". Make sure the path exists and that it is not write-protected.", vbCritical)
    Exit Sub
End If
On Error GoTo 0

Screen.MousePointer = 11

frmMain.ClearDBfields
 
Vars = 20

For A = 1 To Vars
    VarList(A) = ""
    VType(A) = ""
    VLen(A) = 0
    VCarry(A) = False
    VIncr(A) = False
    VMenu(A) = ""
    VPrompt(A) = ""
Next A

VarList(1) = "Sitename"
VType(1) = "TEXT"
VLen(1) = 10
VCarry(1) = True

VarList(2) = "Unit"
VType(2) = "MENU"
VMenu(2) = "UNIT1,UNIT2"
VLen(2) = 5
VCarry(2) = True

VarList(3) = "ID"
VType(3) = "TEXT"
VIncr(3) = True
VCarry(3) = True
VLen(3) = 5

VarList(4) = "Suffix"
VType(4) = "NUMERIC"
VCarry(4) = True

VarList(5) = "Code"
VType(5) = "MENU"
VMenu(5) = "ARTIFACT,BONE,STONE,TOPO"
VLen(5) = 10
VCarry(5) = True

VarList(6) = "Level"
VLen(6) = 10
VMenu(6) = "Level1,Level2"
VType(6) = "MENU"
VCarry(6) = True

VarList(7) = "Excavator"
VLen(7) = 10
VMenu(7) = "Utsav,Steve"
VType(7) = "MENU"
VCarry(7) = True

VarList(8) = "X"
VType(8) = "NUMERIC"
VarList(9) = "Y"
VType(9) = "NUMERIC"
VarList(10) = "Z"
VType(10) = "NUMERIC"
VarList(11) = "PRISM"
VType(11) = "NUMERIC"
VarList(12) = "HANGLE"
VType(12) = "NUMERIC"
VarList(13) = "VANGLE"
VType(13) = "NUMERIC"
VarList(14) = "SLOPED"
VType(14) = "NUMERIC"
VarList(15) = "DAY"
VType(15) = "TEXT"
VarList(16) = "TIME"
VType(16) = "TEXT"
VarList(17) = "DATUMNAME"
VType(17) = "TEXT"
VarList(18) = "DATUMX"
VType(18) = "NUMERIC"
VarList(19) = "DATUMY"
VType(19) = "NUMERIC"
VarList(20) = "DATUMZ"
VType(20) = "NUMERIC"

CFGName = cfgfiletext.Text
cfgfile = cfgfiletext.Text
Call mdiMain.save_cfgfile

If totalstationmenu.Text <> "Simulate" Then
    comport = comportmenu.Text
    Select Case UCase$(paritymenu.Text)
    Case "EVEN"
        comsettings = baudratemenu.Text + ",E," + databitsmenu.Text + "," + stopbitsmenu.Text
    Case "ODD"
        comsettings = baudratemenu.Text + ",O," + databitsmenu.Text + "," + stopbitsmenu.Text
    Case "NONE"
        comsettings = baudratemenu.Text + ",N," + databitsmenu.Text + "," + stopbitsmenu.Text
    Case Else
    End Select
End If

If Dir$(databasetext.Text) <> "" Then
    Kill databasetext.Text
    If Err <> 0 Then
        MsgBox databasetext.Text + " could not be erased." + Chr$(13) + "Ensure that it is not already open in another application.", vbInformation + vbOKOnly
        Exit Sub
    End If
End If

frmMain.ClearDBfields
Call createsitedb(databasetext.Text)
SiteDBname = databasetext.Text

Dim db As Database
Dim poledata As Recordset
Set db = OpenDatabase(SiteDBname)
Set poledata = db.OpenRecordset("edm_poles")
poledata.AddNew
poledata("Name") = "Zero"
poledata("height") = 0
poledata("offset") = 0
poledata.Update
poledata.AddNew
poledata("Name") = "10cm"
poledata("height") = 0.1
poledata("offset") = 0
poledata.Update
poledata.Close
Set DatumTB = db.OpenRecordset("edm_datums")
DatumTB.AddNew
DatumTB("Name") = "Main"
DatumTB("X") = 1000
DatumTB("y") = 1000
DatumTB("z") = 0
DatumTB("day") = Date
DatumTB("time") = Time
DatumTB.Update
DatumTB.Close
Set UnitTB = db.OpenRecordset("edm_units")
UnitTB.AddNew
UnitTB("Unit") = "Unit1"
UnitTB("ID") = 0
UnitTB("minx") = 1000
UnitTB("miny") = 1000
UnitTB("maxx") = 1010
UnitTB("maxy") = 1010
UnitTB.Update
UnitTB.AddNew
UnitTB("Unit") = "Unit2"
UnitTB("ID") = 0
UnitTB("minx") = 1010
UnitTB("miny") = 1010
UnitTB("maxx") = 1020
UnitTB("maxy") = 1020
UnitTB.Update
UnitTB.Close
db.Close
Set db = Nothing

Dim Inidata(9, 2) As String
Dim IniClass As String
Dim Status As Byte
IniClass = "[EDM]"
Inidata(1, 1) = "PointTable"
Inidata(1, 2) = pointsfiletext.Text
Call parse_filename(SiteDBname$, fpath$, fname$, fext$)
Inidata(2, 1) = "Database"
Inidata(2, 2) = fname$ + "." + fext$
Inidata(3, 1) = "DBPath"
Inidata(3, 2) = fpath$
Inidata(4, 1) = "Instrument"
Inidata(4, 2) = totalstationmenu.Text
Inidata(5, 1) = "COMPort"
Inidata(5, 2) = comsettings
Inidata(6, 1) = "StationName"
Inidata(6, 2) = "Main"
Inidata(7, 1) = "StationX"
Inidata(7, 2) = 1000
Inidata(8, 1) = "StationY"
Inidata(8, 2) = 1000
Inidata(9, 1) = "StationZ"
Inidata(9, 2) = 0

Call WriteIni(CFGName, IniClass, Inidata(), Status)

'parsecfg (A)

OpenSite SiteDBname

mdiMain.StatusBar.Panels(4) = "DB: " + LCase(DBName) + "   "

flag = False
For A = 0 To SiteDB.TableDefs.Count - 1
    If LCase(SiteDB.TableDefs(A).Name) = LCase(pointsfiletext.Text) Then
        flag = True
        Exit For
    End If
Next A
PointTableName = pointsfiletext.Text

If flag = False Then
    Call CreatePointTB(pointsfiletext.Text)
    flag = True
End If
If Cancelling Then
    Exit Sub
End If
frmMain.lblPointsWarning.Visible = False
frmMain.txtXYZ(0).Enabled = True
frmMain.txtXYZ(1).Enabled = True
frmMain.txtXYZ(2).Enabled = True
frmMain.txtUnit.Enabled = True
frmMain.txtID.Enabled = True
frmMain.txtPrism.Enabled = True
OpenPointsTable

txtPT = pointsfiletext.Text
StationInitialized = True
frmMain.lblStationWarning.Visible = False
'mdiMain.StatusBar.Panels(5) = "Current Station: " + Station(0) + "  "
CurrentStation.Name = "Main"
CurrentStation.X = 1000
CurrentStation.y = 1000
CurrentStation.z = 0

parsecfg (A)

Screen.MousePointer = 1

Unload Me

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub databasetext_DblClick()

cd.ShowSave
If cd.filename <> "" Then
    databasetext.Text = cd.filename
End If

End Sub

Private Sub Form_Load()

baudratemenu.AddItem "300"
baudratemenu.AddItem "1200"
baudratemenu.AddItem "2400"
baudratemenu.AddItem "4800"
baudratemenu.AddItem "9600"

paritymenu.AddItem "Even"
paritymenu.AddItem "Odd"
paritymenu.AddItem "None"

databitmenu.AddItem "7"
databitmenu.AddItem "8"

stopbitmenu.AddItem "0"
stopbitmenu.AddItem "1"
stopbitmenu.AddItem "2"

totalstationmenu.AddItem "Topcon"
totalstationmenu.AddItem "Leica"
totalstationmenu.AddItem "Wild"
totalstationmenu.AddItem "Nikon"
totalstationmenu.AddItem "Sokkia"
totalstationmenu.AddItem "Simulate"
totalstationmenu.AddItem "Builder"
cfgfiletext.Text = "\My Documents\EDMWIN\Survey.CFG"
databasetext.Text = "\My Documents\EDMWIN\Survey.MDB"
pointsfiletext.Text = "Points"

totalstationmenu.Text = "Simulate"

End Sub

Private Sub totalstationmenu_Click()

Select Case totalstationmenu.Text
Case "Topcon"
    baudratemenu.Text = "1200"
    paritymenu.Text = "Even"
    stopbitmenu.Text = "1"
    databitmenu.Text = "7"
Case Else
End Select

End Sub
