Attribute VB_Name = "EDM"
'Declare Function DEVICETODESKTOP Lib "c:\program files\microsoft activesync\adofiltr.dll" (ByVal desktoplocn As String, ByVal tablelist As String, ByVal sync As Boolean, ByVal overwrite As Integer, ByVal devicelocn As String) As Long
'Declare Function DESKTOPTODEVICE Lib "c:\program files\microsoft activesync\adofiltr.dll" (ByVal desktoplocn As String, ByVal tablelist As String, ByVal sync As Boolean, ByVal overwrite As Integer, ByVal devicelocn As String) As Long
'Declare Function DEVICETODESKTOP Lib "adofiltr.dll" (ByVal desktoplocn As String, ByVal tablelist As String, ByVal sync As Boolean, ByVal overwrite As Integer, ByVal devicelocn As String) As Long
'Declare Function DESKTOPTODEVICE Lib "adofiltr.dll" (ByVal desktoplocn As String, ByVal tablelist As String, ByVal sync As Boolean, ByVal overwrite As Integer, ByVal devicelocn As String) As Long
Global Voice As SpVoice
Global Speaking As Boolean
Global DatumInfo As Boolean
Global FilterString As String
Public Const NewShot = True
Public Const XShot = False
Public Const INVALID_HANDLE = -1
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000 '(1073741824)
Public Const FILE_SHARE_READ = 1
Public Const FILE_SHARE_WRITE = 2
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3
Public Const INIT_SUCCESS = 0
Public Const WRITE_ERROR = 0
Public Const READ_ERROR = 0
Global EDMDelayTime As Single
Global SiteDBOpen As Boolean
Global PlotGridSize As Single
Global Shooting As Boolean
Global GridWidth As Integer
Global GridHeight As Integer
Global GridTop As Integer
Global GridLeft As Integer
'Global PointNo(100000) As Integer
Global GridFormWidth As Integer
Global GridFormHeight As Integer
Global RefDatum1 As String
Global RefDatum2 As String
Global SetupType As Integer
Global UnitFieldString As String
Global PPPath As String
Global CFGTitle As String
Global CFGpath As String
Global DBPath As String
Global DBName As String
Global GridShowing As Boolean
Global PCcfgPath As String
Global PCdbPath As String
Global PPcfgPath As String
Global PPdbPath As String
Global PlotShowing As Boolean
Global Const NoPrism = 0
Global Const AskForPrism = 1
Global XShotShowing As Boolean
Global UpperCase As Boolean
Global NoAlert As Boolean
Global StationInitialized As Boolean
Global StationName As String
Global stationx As Double
Global stationy As Double
Global StationZ As Double
Global stationheight As Double
Global NTempDatums As Integer
Global TempDatumName(100) As String
Global TempDatumX(100) As Double
Global TempDatumY(100) As Double
Global TempDatumZ(100) As Double
Global NTempPrisms As Integer
Global TempPrismName(100) As String
Global TempPrismHeight(100) As Single
Global TempPrismOffset(100) As Single
Global NTempUnits As Integer
Global TempUnitName(500) As String
Global TempUnitMinX(500) As Double
Global TempUnitMinY(500) As Double
Global TempUnitMaxX(500) As Double
Global TempUnitMaxY(500) As Double
Global comport As String
Global comsettings As String
Global ButtonShortCut(6) As String
Global ButtonCaption(6) As String
Global nButtonVars(6) As Integer
Global ButtonVars(6, 100, 2) As Variant
Global PlotX, PlotY As String
Global LastPath As String
Global nPointTables As Integer
Global PointTable(500) As String
Global IDLength As Integer
Global UnitLength As Integer
Global nPoleHeights As Integer
Global PoleHeight(100) As Single
Global PoleOffset(100) As Single
Global OriginalID As String
Global OriginalUnit As String
Global OriginalSuffix As Integer
Global OriginalPoleHT As Single
Global OriginalPrismIndex As Integer
Global rsTemp As recordset
Global edmshot As shotdata
Global GridTopRow As Long
Global GridCurrentRow As Long
Global SqlString As String
Global Loading As Boolean
Global Cancelling As Boolean
Global BannerHeight As Integer
Global BannerWidth As Integer
Global SiteDB As Database
Global DatumTB As recordset
Global UnitTB As recordset
Global PoleTB As recordset
Global cfgTB As recordset
Global limits(100, 4) As Double
Global limitname(100) As String
Global unitlimits As Integer
Global UnitName As String
Global nUnitFields As Integer
Global Unitfield(100) As String
Global datum(100, 4) As Double
Global DatumName(100) As String
Global ndatums As Integer
Global npoles As Integer
Global poledata(100, 2) As Single
Global polename(100) As String
Global Vars As Integer
Global VarList(100) As String
Global VType(100) As String
Global VDefault(100) As String
Global VPrompt(100) As String
Global VLen(100) As Integer
Global VMenu(100) As String
Global VPrint(100) As String
Global VCarry(100) As Boolean
Global VIncr(100) As Boolean
Global VUnique(100) As Boolean
Global SiteDBname As String 'name of database
Global PointTableName As String 'name of table for recorded points
Global DatumTableName As String 'name of table in database for datums
Global UnitTableName As String 'name of table in database for units
Global PoleTableName As String 'name of table in database for pole heights
Global cfgTableName As String 'name of table in database with config info
Global EDMName As String 'stores model name of edm or None
Global CFGName As String
Global vhd As Boolean 'determines whether manual entry is vhd or xyz
Global MenuSelection$ 'used to pass values back from menus
Global CurrentStation As xyz
Global SqidCheck As Boolean
Global LimitChecking As Boolean
Global BackupFolder As String
Global options(100) As String
Global CurrentBookMark As Variant
Global Filtering As Boolean
Global Gotit As Boolean
Global GridLoading As Boolean
Global MasterVar As String
Global MasterVal As Variant
Global nDependentVars As Integer
Global DependentVar(10) As String
Global DefaultsTB As recordset
Global PrinterOn As Boolean
Global UsingMicroscribe As Boolean
Global GeneralLog As Boolean
Global GeneralLogFile As String
Global TSLog As Boolean
Global TSLogFile As String

'option 1 - output filename (mdb)
'option 2 - com parameters
'option 3 - which com
'option 4 - edm type
'option 5
'option 6
'option 7
'option 8
'option 9
'option 10
'option 11 - computer type
'option 12 - sqid system yes or no
'option 13 - pole names
'option 14 - pole heights
'option 15 - pole offset
'option 16
'option 17 - unit fields
'option 18 - site name
'option 19 - printer yes or no
'option 20
'option 21 - use vhd or xyz for manual input
'option 22 - output table name
Type xyz
    Name As String
    X As Double
    y As Double
    z As Double
End Type

Sub CenterForm(formname As Form)

formname.Left = mdiMain.Width / 2 - formname.Width / 2
formname.Top = mdiMain.Height / 2 - formname.Height / 2

End Sub

Sub createsitedb(filename$)

Set SiteDB = Workspaces(0).CreateDatabase(filename$, dbLangGeneral)

Call createdatumtb
Call createpoletb
Call createunitstb
' Call createcfgtb

Set SiteDB = Nothing

End Sub

Sub createcfgtb()

Dim tdf As TableDef

Screen.MousePointer = 11

' Create a new TableDef object.
Set tdf = SiteDB.CreateTableDef("EDM_CFG")

With tdf

    .Fields.Append .CreateField("Field", dbText, 40)
    .Fields.Append .CreateField("Value", dbText, 40)
    
    SiteDB.TableDefs.Append tdf

End With

Screen.MousePointer = 1
Set tdf = Nothing

End Sub

Sub createpoletb()

Dim tdf As TableDef

Screen.MousePointer = 11

' Create a new TableDef object.
Set tdf = SiteDB.CreateTableDef("EDM_Poles")

With tdf

    .Fields.Append .CreateField("Name", dbText, 20)
    .Fields.Append .CreateField("Height", dbDouble)
    .Fields.Append .CreateField("Offset", dbDouble)
    
    SiteDB.TableDefs.Append tdf

End With
Set MainIndex = SiteDB.TableDefs("EDM_Poles").CreateIndex("PoleName")
With MainIndex
    .Fields = "Name"
    .Primary = False
    .Required = False
    .Unique = True
End With
SiteDB.TableDefs("EDM_Poles").Indexes.Append MainIndex
Screen.MousePointer = 1
Set tdf = Nothing
Set MainIndex = Nothing

End Sub

Sub createdatumtb()

Dim tdf As TableDef

Screen.MousePointer = 11

' Create a new TableDef object.
Set tdf = SiteDB.CreateTableDef("EDM_Datums")

With tdf

    .Fields.Append .CreateField("Name", dbText, 20)
    .Fields.Append .CreateField("day", dbDate)
    .Fields.Append .CreateField("time", dbDate)
    .Fields.Append .CreateField("X", dbDouble)
    .Fields.Append .CreateField("Y", dbDouble)
    .Fields.Append .CreateField("Z", dbDouble)
    
    SiteDB.TableDefs.Append tdf

End With
Set MainIndex = SiteDB.TableDefs("EDM_Datums").CreateIndex("DatumName")
With MainIndex
    .Fields = "Name"
    .Primary = False
    .Required = False
    .Unique = True
End With
SiteDB.TableDefs("EDM_Datums").Indexes.Append MainIndex
Screen.MousePointer = 1
Set tdf = Nothing
Set MainIndex = Nothing
End Sub

Function parsefilename$(filename$)
    
Dim A As Integer
Dim temp$

temp$ = filename$
For A = Len(temp$) To 1 Step -1
    If Mid$(temp$, A, 1) = "\" Then
        temp$ = Mid$(temp$, A + 1)
        A = InStr(temp$, ".")
        If A <> 0 Then
            temp$ = Left$(temp$, A - 1)
        End If
        parsefilename$ = temp$
        Exit Function
    End If
Next A

End Function

Function tablematch(tablename$) As Boolean

Dim A As Integer

For A = 0 To SiteDB.TableDefs.Count - 1
    If Trim(UCase$(SiteDB.TableDefs(A).Name)) = Trim(UCase$(tablename$)) Then
        tablematch = True
        Exit Function
    End If
Next A

tablematch = False

End Function

Sub CreatePointTB(tablename$)

tablename = Trim(tablename)

If tablematch(tablename) Then
    response = MsgBox("Overwrite existing table?", vbYesNo)
    If response = vbNo Then
        tablename = ""
        Exit Sub
    Else
        frmMain.PointsADO.RecordSource = ""
        SiteDB.TableDefs.Delete tablename
    End If
End If

Dim tdf As TableDef
Dim F As field
Screen.MousePointer = 11

' Create a new TableDef object.
Set tdf = SiteDB.CreateTableDef(tablename$)

'Now add fields for a points table
With tdf
    Set F = tdf.CreateField("RecNo", dbLong)
    F.Attributes = dbAutoIncrField
    tdf.Fields.Append F
    
    For I = 1 To Vars
        If VLen(I) = 0 Then VLen(I) = 20
        Select Case VType(I)
            Case "TEXT", "MENU", "UNIT", SQID
                Set F = tdf.CreateField(VarList(I), dbText, VLen(I))
                F.AllowZeroLength = True
                F.Required = False
                tdf.Fields.Append F
            Case "NUMERIC", "INSTRUMENT"
                Set F = tdf.CreateField(VarList(I), dbDouble)
                F.Required = False
                tdf.Fields.Append F
        End Select
    Next I
    
End With
SiteDB.TableDefs.Append tdf

Set MainIndex = SiteDB.TableDefs(tablename$).CreateIndex("RecordCounter")
With MainIndex
    .Fields = "RecNo"
    .Primary = True
    .Required = True
    .Unique = True
End With
SiteDB.TableDefs(tablename$).Indexes.Append MainIndex
SiteDB.TableDefs.Refresh
Screen.MousePointer = 1
Set tdf = Nothing
Set F = Nothing
Set MainIndex = Nothing

End Sub

Sub createunitstb()

Dim tdf As TableDef

Screen.MousePointer = 11

' Create a new TableDef object.
Set tdf = SiteDB.CreateTableDef("EDM_Units")
Set tdffield = tdf.CreateField("Unit", dbText, UnitLength)
tdffield.AllowZeroLength = True
tdf.Fields.Append tdffield

With tdf

'    .Fields.Append .CreateField("Unit", dbText, 40)
'    .Fields("unit").AllowZeroLength = True
    .Fields.Append .CreateField("ID", dbText, IDLength)
    .Fields.Append .CreateField("SUFFIX", dbInteger)
    .Fields.Append .CreateField("MinX", dbDouble)
    .Fields.Append .CreateField("MaxX", dbDouble)
    .Fields.Append .CreateField("MinY", dbDouble)
    .Fields.Append .CreateField("MaxY", dbDouble)
    .Fields.Append .CreateField("CenterX", dbDouble)
    .Fields.Append .CreateField("CenterY", dbDouble)
    .Fields.Append .CreateField("Radius", dbDouble)
    
    SiteDB.TableDefs.Append tdf

End With
Set MainIndex = SiteDB.TableDefs("EDM_Units").CreateIndex("UnitName")
With MainIndex
    .Fields = "Unit"
    .Primary = False
    .Required = False
    .Unique = False
End With
SiteDB.TableDefs("EDM_Units").Indexes.Append MainIndex
Screen.MousePointer = 1
Set tdf = Nothing
Set tdffield = Nothing
Set MainIndex = Nothing

End Sub

Sub computeangle(x1 As Double, y1 As Double, x2 As Double, y2 As Double, angle, minutes, seconds)

Dim rrun As Double, rise As Double
Dim slope As Double, pi As Double
Dim tangle As Double

pi = 3.14159265359

rrun = CDbl(x1) - CDbl(x2)
rise = CDbl(y2) - CDbl(y1)

If rrun = 0 Then
        tangle = 90
Else
        slope = rise / rrun
        tangle = CSng(Atn(slope) * 180# / pi)
End If

If (rrun >= 0) And (rise >= 0) Then
        tangle = 0 + tangle
ElseIf (rrun > 0) And (rise < 0) Then
        tangle = 360 + tangle
ElseIf (rrun <= 0) And (rise >= 0) Then
        tangle = 180 + tangle
ElseIf (rrun <= 0) And (rise < 0) Then
        tangle = 180 + tangle
End If

tangle = tangle + 90
If tangle >= 360 Then tangle = tangle - 360

angle = Int(tangle)
seconds = Int((tangle - CDbl(angle)) * 3600#)
minutes = Int(seconds / 60)
seconds = seconds Mod 60

End Sub

Function degtorad(degrees As Single) As Single

degtorad = CSng(degrees * 1.74532925199433E-02)

End Function

Sub insertpointintotb(edmshot As shotdata)

frmMain.PointsADO.recordset.AddNew
On Error Resume Next
frmMain.PointsADO.recordset("x") = Format(edmshot.X, "#########0.000")
frmMain.PointsADO.recordset("y") = Format(edmshot.y, "#########0.000")
frmMain.PointsADO.recordset("z") = Format(edmshot.z, "#########0.000")
frmMain.PointsADO.recordset("vangle") = Format(edmshot.vangle, "##0.0000")
frmMain.PointsADO.recordset("hangle") = Format(edmshot.hangle, "##0.0000")
frmMain.PointsADO.recordset("sloped") = Format(edmshot.sloped, "#########0.000")
frmMain.PointsADO.recordset("prism") = Format(edmshot.poleh, "#########0.000")
frmMain.PointsADO.recordset("datumx") = Format(CurrentStation.X, "#########0.000")
frmMain.PointsADO.recordset("datumy") = Format(CurrentStation.y, "#########0.000")
frmMain.PointsADO.recordset("datumz") = Format(CurrentStation.z, "#########0.000")
On Error GoTo 0
frmMain.PointsADO.recordset.Update

End Sub

Sub writecfg(fieldname$, Value$)

If Not cfgTB.EOF Or Not cfgTB.BOF Then
    cfgTB.MoveFirst
    Do Until cfgTB.EOF
        If LCase$(cfgTB("field")) = LCase$(fieldname$) Then
            cfgTB.Edit
            cfgTB("value") = Value$
            cfgTB.Update
            Exit Sub
        End If
        cfgTB.MoveNext
    Loop
End If

cfgTB.AddNew
cfgTB("field") = fieldname$
cfgTB("value") = Value$
cfgTB.Update

End Sub

Function readcfg(fieldname$) As String

readcfg = ""
If Not cfgTB.EOF Or Not cfgTB.BOF Then
    cfgTB.MoveFirst
    Do Until cfgTB.EOF
        If LCase$(cfgTB("field")) = LCase$(fieldname$) Then
            readcfg = cfgTB("value")
            Exit Function
        End If
        cfgTB.MoveNext
    Loop
End If

End Function

Public Sub ReadIni(inifile, IniClass$, Inidata$(), Status As Byte)

fileno = FreeFile
A$ = ""
On Error Resume Next
A$ = Dir(inifile)
On Error GoTo 0
If A$ = "" Then
    Open inifile For Output As #fileno
    Print #fileno, IniClass$
    Close #fileno
    Exit Sub
End If

Open inifile For Input As #fileno
Do Until EOF(fileno)
    Line Input #fileno, a1$
    a1$ = Trim(a1$)
    a2$ = UCase(a1$)
    If Left$(a2$, 1) = "[" Then
        If a2$ = UCase(IniClass$) Then
            Class = 1
        Else
            Class = 0
        End If
    ElseIf Class = 1 Then
        B = InStr(a2$, "=")
        If B <> 0 Then
            Var$ = Left$(a2$, B - 1)
            vardata$ = Mid$(a1$, B + 1)
            For C = 1 To UBound(Inidata$, 1)
                If UCase$(Inidata$(C, 1)) = UCase$(Var$) Then
                    If Inidata$(C, 2) = "" Then
                        Inidata$(C, 2) = vardata$
                    Else
                        Inidata$(C, 2) = Inidata$(C, 2) + Chr$(1) + vardata$
                    End If
                    Exit For
                End If
            Next C
        End If
    End If
Loop
Close fileno

End Sub

Sub ReadEDMini(inifile$)

If Dir$(inifile$) <> "" Then
    With mdiMain
        Open inifile$ For Input As #10
        Do Until EOF(10)
            Line Input #10, A$
            B = InStr(A$, "=")
            If B <> 0 Then
                Varname$ = UCase$(Left$(A$, B - 1))
                vardata = UCase$(Mid$(A$, B + 1))
                Select Case Varname$
                Case "FILENAME1"
                    If vardata <> "" Then
                        .Filelist(1).Caption = vardata
                        .Filelist(1).Visible = True
                    End If

                Case "FILENAME2"
                    If vardata <> "" Then
                        .Filelist(2).Caption = vardata
                        .Filelist(2).Visible = True
                    End If
                Case "FILENAME3"
                    If vardata <> "" Then
                        .Filelist(3).Caption = vardata
                        .Filelist(3).Visible = True
                    End If
                Case "FILENAME4"
                    If vardata <> "" Then
                        .Filelist(4).Caption = vardata
                        .Filelist(4).Visible = True
                    End If
                Case "FILENAME5"
                    If vardata <> "" Then
                        .Filelist(5).Caption = vardata
                        .Filelist(5).Visible = True
                    End If
                Case "CFGFILE"
                    If Trim(vardata) <> "" Then
                        CFGName = Trim(vardata)
                        parsecfg A
                    End If
                    
                Case "GRIDWIDTH"
                    If Val(vardata) = 0 Then
                        GridWidth = 9525
                    Else
                        GridWidth = Val(vardata)
                    End If
                Case "GRIDHEIGHT"
                    If Val(vardata) = 0 Then
                        GridHeight = 3165
                    Else
                        GridHeight = Val(vardata)
                    End If
                Case "GRIDTOP"
                    If Val(vardata) = 0 Then
                        GridTop = 0
                    Else
                        GridTop = Val(vardata)
                    End If
                Case "GRIDLEFT"
                    If Val(vardata) = 0 Then
                        GridLeft = 0
                    Else
                        GridLeft = Val(vardata)
                    End If
                Case "BACKUPFOLDER"
                    BackupFolder = vardata
                Case "USETSLOG"
                    If vardata = "TRUE" Then TSLog = True Else TSLog = False
                Case "TSLOGFILE"
                    TSLogFile = vardata
                Case "USEGENERALLOG"
                    If vardata = "TRUE" Then GeneralLog = True Else GeneralLog = False
                Case "GENERALLOGFILE"
                    GeneralLogFile = vardata
                Case Else
                End Select
            End If
        Loop
        Close #10
    End With
End If


End Sub

Public Sub WriteIni(inifile, IniClass$, Inidata$(), Status As Byte)

Dim gotit1 As Boolean
If inifile = "" Then Exit Sub
If Left(IniClass, 1) <> "[" Then
    IniClass = "[" + IniClass + "]"
End If
Start:
A = Dir(inifile)
If A = "" Then
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
    Print #1, "EDMDelayTime="; EDMDelayTime
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
Else
    fileno = FreeFile
    Open inifile For Input As #fileno
    filetwo = FreeFile
    Open "$temp$.ini" For Output As #filetwo
    Do Until EOF(fileno)
        Line Input #fileno, a1$
        a2$ = UCase$(Trim$(a1$))
        If Left$(a2$, 1) = "[" Then
            Print #filetwo, a1$
            If a2$ = UCase(IniClass$) Then
                Class = 1
                gotit1 = True
                For C = 1 To UBound(Inidata$, 1)
                    If Trim(Inidata$(C, 1)) <> "" And Trim(Inidata$(C, 2)) <> "" Then
                        Print #filetwo, Inidata$(C, 1) + "=" + Inidata$(C, 2)
                    End If
                Next C
            Else
                Class = 0
            End If
            
        ElseIf Class = 1 Then
            B = InStr(a2$, "=")
            If B <> 0 Then
                Var$ = UCase(Left$(a2$, B - 1))
                vardata$ = Mid$(a2$, B + 1)
                flag = 0
                For C = 1 To UBound(Inidata$, 1)
                    If UCase$(Inidata$(C, 1)) = Var$ Then
                        flag = 1
                        Exit For
                    End If
                Next C
                If flag = 0 Then
                    Print #filetwo, a1$
                End If
            Else
                Print #filetwo, a1$
            
            End If
        Else
            Print #filetwo, a1$
        End If
    Loop
    Close fileno, filetwo
    If Not gotit1 Then
        Open inifile For Append As #fileno
        Print #fileno, " "
        Print #fileno, IniClass$
        Close #fileno
        GoTo Start
    End If
    
    ' This looks silly, but when using DropBox there seem to be issues
    ' with deleting files.  And this seemed to fix it.
    For A = 1 To 20
        Kill inifile
        delay 0.1
        If Dir(inifile) = "" Then Exit For
    Next A
    
    Name "$temp$.ini" As inifile

End If

End Sub

Sub WriteEDMIni(inifile$)

Close 1
With mdiMain

    Open inifile$ For Output As #1
    If .Filelist(1).Visible Then Print #1, "Filename1=" + .Filelist(1).Caption
    If .Filelist(2).Visible Then Print #1, "Filename2=" + .Filelist(2).Caption
    If .Filelist(3).Visible Then Print #1, "Filename3=" + .Filelist(3).Caption
    If .Filelist(4).Visible Then Print #1, "Filename4=" + .Filelist(4).Caption
    If .Filelist(5).Visible Then Print #1, "Filename5=" + .Filelist(5).Caption
    Print #1, "Width=" + Str$(.Width)
    Print #1, "Height=" + Str$(.Height)
    If DBName$ <> "" And InStr(SiteDBname$, "edm.mdb") = 0 Then
        Print #1, "LASTOPEN=" + SiteDBname$
    End If
        
    Print #1, "CFGFILE=" + CFGName
    Print #1, "PCDBPATH=" + PCdbPath
    Print #1, "PCCFGPATH=" + PCcfgPath
    Print #1, "PPDBPATH=" + PPdbPath
    Print #1, "PPCFGPATH=" + PPcfgPath
    Print #1, "GridWidth=" & GridWidth
    Print #1, "Gridheight=" & GridHeight
    Print #1, "Gridleft=" & GridLeft
    Print #1, "Gridtop=" & GridTop
    Print #1, "BackupFolder=" & BackupFolder
    Print #1, "UseGeneralLog=" & GeneralLog
    Print #1, "GeneralLogFile=" & GeneralLogFile
    Print #1, "UseTSLog=" & TSLog
    Print #1, "TSLogFile=" & TSLogFile
    Close #1

End With

End Sub

Function fixpath(fpath$) As String

If Right$(fpath$, 1) <> "\" Then fpath$ = fpath$ + "\"
fixpath = fpath$

End Function

Sub OpenSite(filename$)

Dim errormessage As String
Dim TempName As String

TempName = filename
If Dir(filename) = "" Then
    MsgBox ("Database (" + filename + ") not found.")
    Exit Sub
End If

'***frmMain.PointsADO.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data source" + TempName + ";Persist Security Info=False"
frmMain.UnitsADO.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + TempName + " ;Persist Security Info=False"
frmMain.UnitsADO.RecordSource = "EDM_UNITS"
On Error Resume Next
Set SiteDB = Workspaces(0).OpenDatabase(TempName)
If Err Then
    MsgBox ("Could not open " & TempName & ". " & Err.Description)
    filename$ = ""
    Exit Sub
End If

TempString = TempName
For I = Len(TempString) To 1 Step -1
    If Mid(TempString, I, 1) = "\" Then
        TempString = Mid(TempString, I + 1)
        Exit For
    End If
Next I
mdiMain.StatusBar.Panels(4) = "DB: " + LCase(TempString) + "  "
SiteDBname = TempName
DBName = TempString
DBPath = Left(SiteDBname, Len(SiteDBname) - Len(DBName))
Dim Inidata(2, 2) As String
Dim IniClass As String
Dim Status As Byte
IniClass = "[EDM]"
Inidata(1, 1) = "Database"
Inidata(1, 2) = DBName
Inidata(2, 1) = "Dbpath"
Inidata(2, 2) = DBPath
Call WriteIni(CFGName, IniClass, Inidata(), Status)

If Not tablematch("EDM_datums") Then
        errormessage = errormessage + "Datums "
        Call createdatumtb
End If

If Not tablematch("EDM_Units") Then
        errormessage = errormessage + "Units "
        Call createunitstb
End If

If Not tablematch("EDM_Poles") Then
        errormessage = errormessage + "Pole "
        Call createpoletb
End If

'If Not tablematch("EDM_cfg") Then
'        errormessage = errormessage + "CFG "
'        Call createcfgtb
'End If

If errormessage <> "" Then
    MsgBox ("The following tables were created: " + errormessage)
End If

If Not tablematch("context") Or Not tablematch("xyz") Then
    mdiMain.mnuConvert2Newplot.Enabled = True
Else
    mdiMain.mnuConvert2Newplot.Enabled = False
End If

On Error GoTo 0
DatumTableName$ = "EDM_datums"
Set DatumTB = SiteDB.OpenRecordset("EDM_datums")
On Error GoTo CreateDatumIndex
DatumTB.Index = "DatumName"

UnitTableName$ = "EDM_Units"
Set UnitTB = SiteDB.OpenRecordset("EDM_Units")
'frmMain.UnitsADO.RecordSource = PointTableName
On Error GoTo CreateUnitIndex
UnitTB.Index = "UnitName"
On Error GoTo 0

Loading = True
frmMain.txtUnit.Clear
If Not UnitTB.EOF Then
    UnitTB.MoveFirst
    While Not UnitTB.EOF
        If Not IsNull(UnitTB("unit")) Then
            frmMain.txtUnit.AddItem UnitTB("unit")
        End If
        UnitTB.MoveNext
    Wend
    frmMain.txtUnit.ListIndex = 0
End If
Loading = False

nPoleHeights = 0
Set PoleTB = SiteDB.OpenRecordset("EDM_Poles", dbOpenTable)
On Error GoTo CreatePoleIndex
PoleTB.Index = "PoleName"
On Error GoTo 0

frmMain.txtprism.Clear
If Not PoleTB.EOF Then
    PoleTB.MoveFirst
    While Not PoleTB.EOF
        If Not IsNull(PoleTB("height")) And Not IsNull(PoleTB("offset")) And Not IsNull(PoleTB("Name")) Then
            nPoleHeights = nPoleHeights + 1
            frmMain.txtprism.AddItem PoleTB("Name")
            frmMain.txtprism.ItemData(frmMain.txtprism.NewIndex) = nPoleHeights
            PoleHeight(nPoleHeights) = PoleTB("height")
            PoleOffset(nPoleHeights) = PoleTB("offset")
        End If
        PoleTB.MoveNext
    Wend
    If nPoleHeights > 0 Then
        frmMain.txtprism.ListIndex = 0
        frmMain.txtPoleHT = PoleHeight(frmMain.txtprism.ItemData(frmMain.txtprism.ListIndex))
    End If
    frmMain.lblPoleWarning.Visible = False
Else
    frmMain.lblPoleWarning.Visible = True
End If

frmMain.txtID = ""
frmMain.txtSuffix = ""
frmMain.txtXYZ(0) = ""
frmMain.txtXYZ(1) = ""
frmMain.txtXYZ(2) = ""
frmMain.txtHangle = ""
frmMain.txtVangle = ""
frmMain.txtSloped = ""

frmMain.lblDBWarning.Visible = False

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then
    frmMain.lblPointsWarning.Visible = True
    frmMain.txtXYZ(0).Enabled = False
    frmMain.txtXYZ(1).Enabled = False
    frmMain.txtXYZ(2).Enabled = False
    frmMain.txtUnit.Enabled = False
    frmMain.txtID.Enabled = False
    frmMain.txtprism.Enabled = False
    
Else
    frmMain.lblPointsWarning.Visible = False
    frmMain.txtXYZ(0).Enabled = True
    frmMain.txtXYZ(1).Enabled = True
    frmMain.txtXYZ(2).Enabled = True
    frmMain.txtUnit.Enabled = True
    frmMain.txtID.Enabled = True
    frmMain.txtprism.Enabled = True

End If

PCdbPath = SiteDBname
SiteDBOpen = True

If tablematch("EDM_Defaults") Then
    For I = 0 To SiteDB.TableDefs("EDM_Defaults").Fields.Count - 1
        If I = 0 Then
            MasterVar = SiteDB.TableDefs("EDM_Defaults").Fields(I).Name
        Else
            nDependentVars = nDependentVars + 1
            DependentVar(nDependentVars) = SiteDB.TableDefs("EDM_Defaults").Fields(I).Name
        End If
    Next I
    frmMain.lblDefaults.Visible = True
    Set DefaultsTB = SiteDB.OpenRecordset("EDM_Defaults", dbOpenTable)
    On Error GoTo CreateMasterVarIndex
    DefaultsTB.Index = "MasterVar"
    On Error GoTo 0
Else
    nDependentVars = 0
    MasterVar = ""
End If

Exit Sub

CreateDatumIndex:

Set DatumTB = Nothing
Set MainIndex = SiteDB.TableDefs("EDM_Datums").CreateIndex("DatumName")
With MainIndex
    .Fields = "Name"
    .Primary = False
    .Required = False
    .Unique = True
End With
SiteDB.TableDefs("EDM_Datums").Indexes.Append MainIndex
Set MainIndex = Nothing
Set DatumTB = SiteDB.OpenRecordset("EDM_datums")
Resume
CreateUnitIndex:
Set UnitTB = Nothing
Set MainIndex = SiteDB.TableDefs("EDM_Units").CreateIndex("UnitName")
With MainIndex
    .Fields = "Unit"
    .Primary = False
    .Required = False
    .Unique = False
End With
SiteDB.TableDefs("EDM_Units").Indexes.Append MainIndex
Set UnitTB = SiteDB.OpenRecordset("EDM_Units")
Set MainIndex = Nothing
Resume

CreatePoleIndex:

Set PoleTB = Nothing
Set MainIndex = SiteDB.TableDefs("EDM_Poles").CreateIndex("PoleName")
With MainIndex
    .Fields = "Name"
    .Primary = False
    .Required = False
    .Unique = True
End With
SiteDB.TableDefs("EDM_Poles").Indexes.Append MainIndex
Set PoleTB = SiteDB.OpenRecordset("EDM_Poles", dbOpenTable)
Set MainIndex = Nothing
Resume

CreateMasterVarIndex:

Set DefaultsTB = Nothing
Set MainIndex = SiteDB.TableDefs("EDM_defaults").CreateIndex("MasterVar")
With MainIndex
    .Fields = MasterVar
    .Primary = False
    .Required = False
    .Unique = False
End With
SiteDB.TableDefs("EDM_defaults").Indexes.Append MainIndex
Set MainIndex = Nothing
Set DefaultsTB = SiteDB.OpenRecordset("EDM_Defaults", dbOpenTable)
Resume

End Sub

Sub addtofilelist(DBName$)

With mdiMain
    For I = 1 To 5
        If UCase(.Filelist(I).Caption) = UCase(CFGName) Then
            Exit Sub
        End If
    Next I
            
    For A = 5 To 2 Step -1
        If .Filelist(A - 1).Visible Then
            .Filelist(A).Caption = .Filelist(A - 1).Caption
            .Filelist(A).Visible = True
        End If
    Next A
    .Filelist(1).Caption = CFGName
    .Filelist(1).Visible = True
End With

End Sub

Public Sub ClearButtonIni()

Dim Gotit As Boolean
If CFGName = "" Then Exit Sub

fileno = FreeFile
Open CFGName For Input As #fileno
filetwo = FreeFile
Open "$temp$.ini" For Output As #filetwo
On Error GoTo readerror
Do Until EOF(fileno)
    Line Input #fileno, a1$
Start:
    a2$ = UCase$(Trim$(a1$))
    If Left(LCase(a2$), 7) = "[button" Then
        Line Input #fileno, a1$
        While Not EOF(fileno)
            If Left(a1$, 1) = "[" Then GoTo Start
            Line Input #fileno, a1$
        Wend
    Else
        Print #filetwo, a1$
    End If

Loop

Finish:
Close fileno, filetwo
Kill CFGName
Name "$temp$.ini" As CFGName
Exit Sub

readerror:
Resume Finish

End Sub

Sub parse_filename(filename$, fpath$, fname$, fext$)

A$ = filename$
B = InStr(A$, ".")
If B <> 0 Then
    fext$ = Mid$(A$, B + 1)
    A$ = Left$(A$, B - 1)
Else
    fext$ = ""
End If

fpath$ = ""
fname$ = A$
For B = Len(A$) To 1 Step -1
    If Mid$(A$, B, 1) = "\" Then
        fname$ = Mid$(A$, B + 1)
        fpath$ = Left$(A$, B)
        Exit Sub
    End If
Next B

End Sub

Public Sub CreateDefaultstb()

Dim tdf As TableDef
Dim fdf As field
Screen.MousePointer = 11

' Create a new TableDef object.
Set tdf = SiteDB.CreateTableDef("EDM_Defaults")


Set fdf = tdf.CreateField(MasterVar, dbText, 50)
tdf.Fields.Append fdf

For I = 1 To nDependentVars
    Set fdf = tdf.CreateField(DependentVar(I), dbText, 50)
    tdf.Fields.Append fdf
Next I

SiteDB.TableDefs.Append tdf


Set MainIndex = SiteDB.TableDefs("EDM_Defaults").CreateIndex("MasterVar")
With MainIndex
    .Fields = MasterVar
    .Primary = False
    .Required = False
    .Unique = True
End With
SiteDB.TableDefs("EDM_Defaults").Indexes.Append MainIndex
Screen.MousePointer = 1
Set tdf = Nothing
Set MainIndex = Nothing

frmMain.lblDefaults.Visible = True
Set DefaultsTB = SiteDB.OpenRecordset("EDM_Defaults", dbOpenTable)
DefaultsTB.Index = "MasterVar"

End Sub

Function field_in_recordset(recordset, fieldname) As Boolean

field_in_recordset = False
For Each field In recordset.Fields
    If LCase(field.Name) = LCase(fieldname) Then
        field_in_recordset = True
        Exit For
    End If
Next

End Function
