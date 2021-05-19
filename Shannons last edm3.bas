Attribute VB_Name = "ShanonslastEDMWindows"

Type shotdata
    X As Double
    y As Double
    z As Double
    sloped As Double
    hangle As Single
    vangle As Single
    poleh As Single
    poleo As Single
    edmpoffset As Single
End Type

Type displayinfo
    foreground As Integer           'foreground color
    background As Integer           'background color
    blink As Integer                'blink color on HP-95
    comptype As Integer             'type of computer PC or HP
    Width As Integer                'width in characters of screen
    datalines As Integer            'lines available for data display
    printer As Integer              'whether there is a printer attached
    sqidsys As Integer              'sqid system or no
    hpmodelno As Integer            'indicates 95, 100, or 200
    printerfile As Integer          'print file for 200
    datasource As Integer           'where to get edm data from
    vhdmode As Integer              'hand input xyz or vhd
    edmtype As Integer              'type of theodolite in use
    prismprompt As Integer          'whether we are prompting for prisms
End Type

Public Sub CreateContext()

ContextTableName = "Context"
Set ContextTableDef = SiteDB.CreateTableDef(ContextTableName)
Set ContextField = ContextTableDef.CreateField("Unit", dbText, UnitLength)
ContextField.AllowZeroLength = True
ContextTableDef.Fields.Append ContextField
Set ContextField = ContextTableDef.CreateField("ID", dbText, IDLength)
ContextField.AllowZeroLength = True
ContextTableDef.Fields.Append ContextField
For I = 1 To Vars
    Select Case UCase(VarList(I))
        Case "UNIT", "ID", "SUFFIX", "PRISM", "X", "Y", "Z", "VANGLE", "HANGLE", "SLOPED"
        Case Else
            Select Case UCase(VType(I))
                Case "TEXT", "MENU"
                    Set ContextField = ContextTableDef.CreateField(VarList(I), dbText, VLen(I))
                    ContextField.AllowZeroLength = True
                    ContextTableDef.Fields.Append ContextField
                Case "NUMERIC"
                    Set ContextField = ContextTableDef.CreateField(VarList(I), dbSingle)
                    ContextField.AllowZeroLength = True
                    ContextTableDef.Fields.Append ContextField
            End Select
    End Select
Next I

SiteDB.TableDefs.Append ContextTableDef
Set ContextIndex = SiteDB.TableDefs(ContextTableName).CreateIndex("SqidIndex")
With ContextIndex
    .Fields = "Unit;ID"
    .Primary = False
    .Required = False
    .Unique = False
End With
SiteDB.TableDefs(ContextTableName).Indexes.Append ContextIndex

End Sub


Public Sub CreateXYZ()
     
        
XYZTableName = "XYZ"
Set XYZTableDef = SiteDB.CreateTableDef(XYZTableName)
Set XYZField = XYZTableDef.CreateField("Unit", dbText, UnitLength)
XYZField.AllowZeroLength = True
XYZTableDef.Fields.Append XYZField
Set XYZField = XYZTableDef.CreateField("ID", dbText, IDLength)
XYZField.AllowZeroLength = True
XYZTableDef.Fields.Append XYZField
Set XYZField = XYZTableDef.CreateField("Suffix", dbInteger)
XYZTableDef.Fields.Append XYZField
Set XYZField = XYZTableDef.CreateField("X", dbSingle)
XYZTableDef.Fields.Append XYZField
Set XYZField = XYZTableDef.CreateField("Y", dbSingle)
XYZTableDef.Fields.Append XYZField
Set XYZField = XYZTableDef.CreateField("Z", dbSingle)
XYZTableDef.Fields.Append XYZField
Set XYZField = XYZTableDef.CreateField("RecordCounter", dbLong)
XYZField.Attributes = dbAutoIncrField
XYZTableDef.Fields.Append XYZField
For I = 1 To Vars
    Select Case UCase(VarList(I))
        Case "PRISM", "VANGLE", "HANGLE", "SLOPED"
            Set XYZField = XYZTableDef.CreateField(VarList(I), dbSingle)
            XYZTableDef.Fields.Append XYZField
    End Select
Next I
    

SiteDB.TableDefs.Append XYZTableDef
Set XYZIndex = SiteDB.TableDefs(XYZTableName).CreateIndex("SqidIndex")
With XYZIndex
    .Fields = "unit;ID"
    .Primary = False
    .Required = False
    .Unique = False
End With
SiteDB.TableDefs(XYZTableName).Indexes.Append XYZIndex

Set XYZIndex = XYZTableDef.CreateIndex("SqidSuffixIndex")
With XYZIndex
    .Fields = "unit;ID;Suffix"
    .Primary = False
    .Required = False
    .Unique = False
End With
SiteDB.TableDefs(XYZTableName).Indexes.Append XYZIndex

Set XYZIndex = XYZTableDef.CreateIndex("RecordCounter")
With XYZIndex
    .Fields = "RecordCounter"
    .Primary = False
    .Required = False
    .Unique = False
End With
SiteDB.TableDefs(XYZTableName).Indexes.Append XYZIndex
End Sub

Sub clearcom()

a$ = frmMain.theoport.Input

End Sub

Sub delay(delaytime As Single)

Dim t1 As Double

t1 = Timer

Do Until (Timer - t1) > delaytime Or inkey$ <> ""
Loop

End Sub

Sub directoutput(a$)

frmMain.theoport.Output = a$

End Sub

Sub displayvertangle()

Select Case EDMName$
Case "Topcon"
    d$ = "Z20"
    Call edmoutput(d$, errorcode)
Case 2
Case Else
End Select

End Sub

Sub edmack()

If EDMName$ = "Topcon" Then
    ack$ = Chr$(6) + "006" + Chr$(3) + Chr$(13) + Chr$(10)
    frmMain.theoport.Output = ack$
End If

End Sub

Sub edminput(a As String)
                                                         
timeone = Timer

a$ = ""
t$ = ""
Do
    t$ = frmMain.theoport.Input
    If t$ <> "" Then a$ = a$ + t$
Loop Until Right$(a$, 2) = Chr$(13) + Chr$(10) Or Timer - timeone > 15

End Sub

Sub edmoutput(d$, errorcode)

term$ = Chr$(13) + Chr$(10)
errorcode = 0
Select Case EDMName$
Case "Topcon"
    Call makebcc(d$, bcc$)
    a$ = d$ + bcc$ + Chr$(3) + term$
Case "Wild"
    a$ = d$ + term$
Case "Sokkia"
    a$ = d$
Case Else
    Exit Sub
End Select

Select Case EDMName$
Case "Topcon"
    frmMain.theoport.Output = a$
Case "Wild"
    frmMain.theoport.Output = a$
Case Else
End Select

End Sub

Sub getpresetdata(hangle As Single, X As Single, y As Single, z As Single, angleunit$, mesunits$, errormessage$)

errormessage$ = ""
If display.edmtype <> 1 Then Exit Sub

d$ = "L"
Call edmoutput(d$, errorcode)
Call edminput(B$)

Do
    Call edminput(a$)
Loop Until InStr(a$, "L") <> 0 Or a$ = "CANCEL"

If a$ <> "CANCEL" Then
    ack$ = Chr$(6) + "006" + Chr$(3) + Chr$(13) + Chr$(10)
    Call delay(0.05)
    Call directoutput(ack$)
    returndata$ = a$
    Call parsepreset(returndata$, hangle, X, y, z, angleunit$, mesunits$, errormessage$)
Else
    Call horizontal(errorcode)
End If

End Sub

Sub horizontal(errorcode)

errorcode = 0
Select Case EDMName$
Case "Topcon"
    d$ = "Z10"
    Call edmoutput(d$, errorcode)
    Call edminput(a$)
    If a$ = "CANCEL" Then errorcode = 27

Case Else
End Select

End Sub

Sub horizontalright()

Select Case EDMName$
Case "Topcon"
    d$ = "Z12"
    Call edmoutput(d$, errorcode)
Case 2
Case Else
End Select

End Sub

Sub hortizontalleft()

Select Case EDMName$
Case "Topcon"
    d$ = "Z13"
    Call edmoutput(d$, errorcode)
Case 2
Case Else
End Select

End Sub

Sub initcomport(comport$, errorcode)

'--------------------------------------------------------------------------
'The switch to horizontal isn't really necessary.
'I do it just to make the EDM beep in acknowledgement of the connection.
'If no EDM is present, however, the program will hange until ESCape.
'--------------------------------------------------------------------------

If frmMain.theoport.PortOpen Then frmMain.theoport.PortOpen = False
If comport$ = "" Or comsettings$ = "" Then
    MsgBox "The COMport settings (baud rate, parity, stop bits, data bits) have not yet been set. Go to select total station to do this.", vbInformation
    Exit Sub
End If

frmMain.theoport.Settings = comsettings
Select Case UCase$(comport)
Case "COM1"
    frmMain.theoport.CommPort = 1
Case "COM2"
    frmMain.theoport.CommPort = 2
Case "COM3"
    frmMain.theoport.CommPort = 3
Case "COM4"
    frmMain.theoport.CommPort = 4
Case Else
End Select

frmMain.theoport.PortOpen = True

Call initedm
Call horizontal(errorcode)

End Sub

Sub initedm()

Dim d$

Select Case EDMName$
Case "Topcon"
    d$ = "ST0"
    Call edmoutput(d$, errorcode)

Case "Wild"
    d$ = "SET/41/0"
    Call edmoutput(d$, errorcode)
    If errorcode = 0 Then
                    Call edminput(a$)
    End If
    d$ = "SET/149/2"
    Call edmoutput(d$, errorcode)
    If errorcode = 0 Then
                    Call edminput(a$)
    End If
    Call delay(0.5)
    Call clearcom

Case Else
End Select

End Sub

Sub makebcc(I$, O$)

Dim l As Integer

B = 0
For l = 1 To Len(I$)
                q = Asc(Mid$(I$, l, 1))
                b1 = q And (Not B)
                b2 = B And (Not q)
                B = b1 Or b2
Next l

O$ = LTrim$(Str$(B))
O$ = Right$("000" + O$, 3)

End Sub

Sub parsenez(nezdata$, edmshot As shotdata, edmpoffset As Single, mesunits$, angleunit$, errorcode)

Dim angle As Integer, minutes As Integer, seconds As Integer
Dim tangle As Single, dangle As Single, dist As Single

errorcode = 0

If nezdata$ = "" Then
    errorcode = -99
    Exit Sub
End If

Select Case EDMName$
Case "Topcon"
    Do Until Asc(Left$(nezdata$, 1)) > 32 Or Len(nezdata$) = 1
        nezdata$ = Mid$(nezdata$, 2)
    Loop
    a$ = Left$(nezdata$, 1)
    If a$ <> "?" And a$ <> "R" Then
        If a$ = "U" Then
            errorcode = 5
        Else
            errorcode = 1
        End If
        Exit Sub
    End If

    a = InStr(nezdata$, Chr$(3))
    If a <> 0 Then
        bcc1$ = Mid$(nezdata$, a - 3, 3)
        d$ = Left$(nezdata$, a - 4)
        Call makebcc(d$, bcc2$)
    End If

    If bcc1$ <> bcc2$ Then
        errorcode = 6
    Else
        nezdata$ = LTrim$(nezdata$)
        edmshot.sloped = Val(Mid$(nezdata$, 2, 9)) / 1000
        edmshot.hangle = Val(Mid$(nezdata$, 19, 4) + "." + Mid$(nezdata$, 23, 4))
        edmshot.vangle = Val(Mid$(nezdata$, 12, 3) + "." + Mid$(nezdata$, 15, 4))
        edmshot.edmpoffset = Val(Mid$(nezdata$, 43, 3)) / 1000
        mesunits$ = Mid$(nezdata$, 11, 1)
        angleunit$ = Mid$(nezdata$, 27, 1)
        If angleunit$ <> "d" Then
            errorcode = 2
        End If
    End If

Case 2
    a = InStr(nezdata$, "21.")
    If a = 0 Then
        errorcode = 1
        Exit Sub
    End If
    nezdata$ = Mid$(nezdata$, a)

    'IF LEN(nezdata$) <> 64 OR a = 0 THEN
    '        errorcode = 1
    '        EXIT SUB
    'END IF

    If Mid$(nezdata$, 6, 1) <> "4" Then
        errorcode = 2
        Exit Sub
    End If
    edmshot.hangle = Val(Mid$(nezdata$, 7, 4) + "." + Mid$(nezdata$, 11, 4))

    If Mid$(nezdata$, 22, 1) <> "4" Then
        errorcode = 2
        Exit Sub
    End If
    edmshot.vangle = Val(Mid$(nezdata$, 23, 4) + "." + Mid$(nezdata$, 27, 4))

    a = InStr(nezdata$, "31..0")
    edmshot.sloped = Val(Mid$(nezdata$, a + 6, 9)) / 1000
    If edmshot.sloped = 0 Then
        errorcode = 4
        Exit Sub
    End If
    mesunits$ = Mid$(nezdata$, a + 5)
    edmshot.edmpoffset = Val(Right$(nezdata$, 4)) / 1000

Case 3
    edmshot.edmpoffset = 0
    nezdata$ = RTrim$(LTrim$(nezdata$))
    a = InStr(nezdata$, " ")
    If a > 1 Then
        edmshot.sloped = Val(Left$(nezdata$, a - 1)) / 1000
        If edmshot.sloped = 0 Then
            errorcode = 10
        Else
            a$ = Mid$(nezdata$, a + 1)
            a = InStr(a$, " ")
            If a > 1 And Left$(a$, 1) <> "E" Then
                edmshot.vangle = Val(Left$(a$, 3) + "." + Mid$(a$, 4, 4))
                a$ = Mid$(a$, a + 1)
                a = InStr(a$, " ")
                If a > 1 And Left$(a$, 1) <> "E" Then
                    edmshot.hangle = Val(Left$(a$, 3) + "." + Mid$(a$, 4, 4))
                Else
                    errorcode = 1
                End If
            Else
                errorcode = 1
            End If
        End If
    Else
        errorcode = 1
    End If

Case Else
End Select

End Sub

Sub parsepreset(preset$, hangle As Single, X As Single, y As Single, z As Single, angleunit$, mesunits$, errormessage$)

a = InStr(preset$, "+")
B = InStr(preset$, "-")
If B < a And B <> 0 Then a = B
preset$ = " L" + Mid$(preset$, a)

bcc1$ = Mid$(preset$, 52, 3)
d$ = Mid$(preset$, 2, 50)
Call makebcc(d$, bcc2$)

If bcc1$ <> bcc2$ Then
                errormessage$ = "ERROR: BCC mismatch."
Else
                errorcode = 0
                hangle = Val(Mid$(preset$, 3, 8)) / 10000
                angleunit$ = Mid$(preset$, 11, 1)
                y = Val(Mid$(preset$, 12, 9)) / 1000
                X = Val(Mid$(preset$, 21, 9)) / 1000
                mesunits$ = Mid$(preset$, 30, 1)
                z = Val(Mid$(preset$, 31, 9)) / 1000

End If


End Sub

Sub parsevh(vhdata$, vangle As Single, hangle As Single, tilt, errorcode)

a = InStr(vhdata$, "<")
vhdata$ = Mid$(vhdata$, a)

bcc1$ = Mid$(vhdata$, 23, 3)
d$ = Mid$(vhdata$, 1, 22)
Call makebcc(d$, bcc2$)

If bcc1$ <> bcc2$ Then
        errorcode = 1
Else
        errorcode = 0
        vangle = Val(Mid$(vhdata$, 2, 7)) / 10000
        hangle = Val(Mid$(vhdata$, 9, 8)) / 10000
        tilt = Val(Mid$(vhdata$, 17, 5))
End If

End Sub

Sub recordpoint(returndata$)

Select Case UCase(EDMName$)
Case "TOPCON"
    ack$ = Chr$(6) + "006" + Chr$(3) + Chr$(13) + Chr$(10)
    notack$ = Chr$(21) + "021" + Chr$(3) + Chr$(13) + Chr$(10)
    nezback$ = ""
    returndata$ = ""
    bcc1$ = ""
    bcc2$ = ""

    '------------------------------------------------
    ' Take measurement in sloped and angle mode
    '------------------------------------------------
    d$ = "Z34"
    Call edmoutput(d$, errorcode)
    Call edminput(a$)
    frmTakeshot.progress.Value = 25
    'frmTakeshot.Show
    'frmTakeshot.Refresh
    
    Call delay(0.5)

    '------------------------------------------------
    ' Take the actual measurement
    '------------------------------------------------
    d$ = "C"
    Call edmoutput(d$, errorcode)
    Call edminput(a$)

    Call delay(0.5)
    
    'frmTakeshot.progress.Value = 50
    'frmTakeshot.progress.Refresh
    
    Call edminput(a$)

    If a$ <> "CANCEL" And a$ <> "" Then
        'CALL delay(.1)
        Call directoutput(ack$)
        Do Until Asc(Left$(a$, 1)) > 32 Or Len(a$) = 1
            a$ = Mid$(a$, 2)
        Loop
    Else
        'CALL horizontal( errorcode)
    End If

    returndata$ = a$
    Call delay(0.1)
    'frmTakeshot.progress.Value = 75
    'frmTakeshot.progress.Refresh
    Call horizontal(errorcode)
    Call clearcom

Case "WILD"
    d$ = "GET/M/WI11/WI21/WI22/WI31/WI51"
    Call edmoutput(d$, errorcode)
    Call edminput(a$)
    returndata$ = a$
    Call clearcom

Case "SOKKIA"
    d$ = Chr$(17)
    Call edmoutput(d$, errorcode)
    Call edminput(a$)
    returndata$ = a$
    If a$ = "CANCEL" Then
        d$ = Chr$(18)
        Call edmoutput(d$, errorcode)
    End If
    Call clearcom

Case Else
End Select

'frmTakeshot.Hide

End Sub

Sub sethortangle(angle$, deg, min, sec)

Dim a As String

If angle$ = "" Then
    angle$ = LTrim$(Str$(deg)) + LTrim$(Str$(min)) + LTrim$(Str$(sec))

ElseIf InStr(angle$, ".") <> 0 Then
    a = InStr(angle$, ".")
    angle$ = Left$(angle$ + "0000", a + 4)
    angle$ = Left$(angle$, a - 1) + Mid$(angle$, a + 1)

End If

Select Case UCase(EDMName)
Case "TOPCON"
    d$ = "J+" + LTrim$(angle$) + "d"
    Call edmoutput(d$, errorcode)
    Call edminput(a$)
    Call delay(1)
    Call clearcom

Case "WILD"
    d$ = "PUT/21...4+" + Right$("000" + LTrim$(angle$) + "0 ", 9)
    Call edmoutput(d$, errorcode)
    Call edminput(a$)

Case Else
End Select

End Sub

Sub setpresetdata(X As Single, y As Single, z As Single, errormessage$)

If displayinfo.edmtype <> 1 Then Exit Sub

px$ = LTrim$(Str$(X))
If X >= 0 Then px$ = "+" + px$
a = InStr(px$, ".")
If a <> 0 Then
                px$ = Left$(px$, a - 1) + Left$(Mid$(px$, a + 1), 3)
Else
                px$ = px$ + "000"
End If

py$ = LTrim$(Str$(y))
If X >= 0 Then py$ = "+" + py$
a = InStr(py$, ".")
If a <> 0 Then
                py$ = Left$(py$, a - 1) + Left$(Mid$(py$, a + 1), 3)
Else
                py$ = py$ + "000"
End If

d$ = "I" + px$ + py$ + "m"

Call edmoutput(d$, errorcode)
Call delay(0.5)
Call clearcom

If errorcode = 0 Then
                pz$ = LTrim$(Str$(z * -1))
                If (z * -1) >= 0 Then pz$ = "+" + pz$
                a = InStr(pz$, ".")
                If a <> 0 Then
                                pz$ = Left$(pz$, a - 1) + Left$(Mid$(pz$, a + 1), 3)
                Else
                                pz$ = pz$ + "000"
                End If
                d$ = "K" + pz$ + "mz"
                Call edmoutput(d$, errorcode)
                Call delay(0.5)
                Call clearcom
End If

End Sub


Sub vhdtonez(edmshot As shotdata)

Dim angle As Integer, minutes As Integer, seconds As Integer
Dim tangle As Single
Dim actuald As Double

Call parseangle(edmshot.vangle, angle, minutes, seconds)
tangle = angle + ((minutes * 60 + seconds) / 3600)

'Adjust the slope distance for the prism offset
If edmshot.poleo <> 0 And (edmshot.poleo <> edmshot.edmpoffset) Then
    If edmshot.edmpoffset <> 0 Then
        edmshot.sloped = edmshot.sloped - edmshot.edmpoffset
    End If
    If edmshot.poleo <> 0 Then
        edmshot.sloped = edmshot.sloped + edmshot.poleo
    End If
End If

edmshot.z = edmshot.sloped * Cos(degtorad(tangle))

actuald = Sqr(edmshot.sloped ^ 2 - edmshot.z ^ 2)

Call parseangle(edmshot.hangle, angle, minutes, seconds)
tangle = angle + ((minutes * 60 + seconds) / 3600)
tangle = 450 - tangle

edmshot.X = Cos(degtorad(tangle)) * actuald
edmshot.y = Sin(degtorad(tangle)) * actuald

End Sub

Function hash(hashlen)

Dim a As Integer
Randomize
hash = ""
For a = 1 To hashlen
    hash = hash + Chr(Rnd * 25 + Asc("A"))
Next a

End Function

Sub parsecfg(needtoconvert)

'--------------------------------------------------
' Open the config file passed from main module
' Note that we already know it does in fact exist.
'--------------------------------------------------

Dim lineno As Integer
Dim fl As String
Dim ts As String
Dim a As Integer
Dim B As Integer
Dim C As Integer
Dim d As Integer
Dim key As String
Dim dt As String

Dim temp(100) As String
Dim edmce As Boolean
Dim answer As Integer

Dim UnitFieldString As String
Dim errorcode As Integer
Dim errormessage As String

Dim Inidata(100, 2) As String
Dim IniClass As String
Dim Status As Byte

frmMain.ClearDBfields

Cancelling = False
TempString = ""
On Error Resume Next
TempString = Dir(CFGName)
On Error GoTo 0

If TempString = "" Then
    MsgBox ("CFG file not found.")
    frmMain.lblDBWarning.Visible = True
    frmMain.lblPointsWarning.Visible = True
    frmMain.lblCFGWarning.Visible = True
    Set PointsTB = Nothing
    Set UnitTB = Nothing
    Set PoleTB = Nothing
    Set DatumTB = Nothing
    Set SiteDB = Nothing
    PointTableName = ""
    SiteDBname = ""
    CFGName = ""
    mdiMain.StatusBar.panels(3) = ""
    mdiMain.StatusBar.panels(4) = ""
    frmMain.txtPT = ""
    frmMain.txtPrism.Clear
    frmMain.txtPoleHT = ""
    Cancelling = True
    Vars = 0
    Exit Sub
End If

TempString = CFGName
For I = Len(TempString) To 1 Step -1
    If Mid(TempString, I, 1) = "\" Then
        TempString = Mid(TempString, I + 1)
        Exit For
    End If
Next I
mdiMain.StatusBar.panels(3) = "CFG: " + LCase(TempString) + "  "
npoles = 0
Vars = 0
ndatums = 0
needtoconvert = 0

'first, check to see if the file is in old EDM format
'or new EDMCE/EDM format by looking for that section in the file

edmce = False
Open CFGName For Input As 1
Do While Not EOF(1)
    Line Input #1, ts
    If InStr(UCase(ts), "[EDM]") <> 0 Then
        edmce = True
        Exit Do
    End If
Loop
Close 1

If Not edmce Then
    lineno = 0
    Open CFGName For Input As 1
    Do While Not EOF(1)
            '--------------------------------------------------
            ' Look for a continuation character at the end of
            ' the line.  If it exists add the next line to a.
            ' In the end, a contains a full line from the CFG.
            '--------------------------------------------------
            fl = ""
            Do
                    lineno = lineno + 1
                    Line Input #1, ts
                    ts = Trim(ts)
                    fl = fl + ts
                    If Right(ts, 1) = "\" Then Mid(fl, Len(fl), 1) = " "
            Loop While Right(ts, 1) = "\" And Not EOF(1)
    
            '-------------------------------------------
            ' Check for an empty line or for a ' comment
            '-------------------------------------------
            If (Len(fl) <> 0 Or fl <> Space(Len(fl))) And Left(fl, 1) <> "'" Then
                    
                    'MsgBox fl
                    
                    '----------------------------------------------
                    ' Look for : as a marker for keywords
                    '----------------------------------------------
                    a = InStr(fl, ":")
                    If a <> 0 Then
    
                            '--------------------------------------------
                            ' Set key to the keyword then select case on
                            ' the different types of keyword.
                            '--------------------------------------------
                            key = UCase(Trim(Left(fl, a - 1)))
                            dt = Trim(Mid(fl, a + 1))
                            Select Case key
                            Case "FIELD"
    
                                    If dt <> "" Then
                                        Gotit = False
                                        For B = 1 To Vars
                                            If VarList(B) = dt Then
                                                Gotit = True
                                            End If
                                        Next B
        
                                        If Not Gotit Then
                                            Vars = Vars + 1
                                            VarList(Vars) = UCase(dt)
                                        End If
                                    End If
                            Case "POINTSFILE"
                                    PointTableName = UCase(dt)
    
                            Case "COM1"
                                    Options(2) = "COM1"
                                    Options(3) = UCase(dt)
    
                            Case "COM2"
                                    Options(2) = "COM2"
                                    Options(3) = UCase(dt)
    
                            Case "EDM"
                                    ts = UCase(dt)
                                    d = InStr(ts, " ")
                                    If ts <> "NONE" Then
                                            comport = Trim(Mid(ts, d))
                                            Options(4) = UCase(LTrim(Left(ts, d - 1)))
                                    Else
                                            Options(4) = "NONE"
                                    End If
                                    Select Case Options(4)
                                    Case "GTS-3B", "GTS-301", "GTS-302", "GTS-303", "GTS-3X", "TOPCON"
                                            Options(4) = "TOPCON" + " " + comport
                                    Case "NONE"
                                    Case "MANUAL"
                                    Case "TC-500", "WILD"
                                            Options(4) = "WILD" + " " + comport
                                    Case "SOKKIA"
                                            Options(4) = "SOKKIA" + " " + comport
                                    Case Else
                                    End Select
    
                            Case "PRINTFIELDS"
    
                            Case "TEXTFOR"
                            
                            Case "TEXTBACK"
    
                            Case "SQIDFILE"
    
                            '--------------------------------------------------
                            'Option to know computer type (ie. HP vs PC)
                            '--------------------------------------------------
                            Case "COMPUTER"
                                    Options(11) = UCase(dt)
                                    Select Case Options(11)
                                    Case "HP"
                                        Options(11) = "HP"
                                    Case "HP-95", "HP-95LX", "HP95"
                                        Options(11) = "HP95"
                                    Case "HP-100", "HP100"
                                        Options(11) = "HP100"
                                    Case "HP-200", "HP-200LX", "HP200LX", "HP200"
                                        Options(11) = "HP200"
                                    Case "HUSKY"
                                        Options(11) = "HUSKY"
                                    Case "PC", "PC80", "2000"
                                    Case Else
                                    End Select
    
                            Case "SQID"
                                    '-------------------------------------
                                    'Is this a Laq/CC/Cagny type setup
                                    '-------------------------------------
                                    Options(12) = Trim(dt)
    
                            Case "PRISMDEF"
                                    ts = dt
                                    d = InStr(ts, " ")
                                    If d > 0 Then
                                        NTempPrisms = NTempPrisms + 1
                                        TempPrismName(NTempPrisms) = UCase(Trim(Left(ts, d - 1)))
                                        ts = LTrim(Mid(ts, d))
                                        d = InStr(ts, " ")
                                        If d > 0 Then
                                            TempPrismHeight(NTempPrisms) = UCase(Trim(Left(ts, d - 1)))
                                            TempPrismOffset(NTempPrisms) = UCase(Trim(Mid(ts, d + 1)))
                                        End If
                                    End If
    
                            Case "DATUMDEF"
                                    ts = dt
                                    F = 0
                                    Do Until Len(ts) = 0
                                            E = InStr(ts, " ")
                                            If E = 0 Then E = Len(ts) + 1
                                            F = F + 1
                                            temp(F) = Left(ts, E - 1)
                                            ts = LTrim(Mid(ts, E))
                                    Loop
                                    If F = 5 Then
                                            NTempDatums = NTempDatums + 1
                                            TempDatumName(NTempDatums) = temp(1)
                                            TempDatumX(NTempDatums) = temp(2)
                                            TempDatumY(NTempDatums) = temp(3)
                                            TempDatumZ(NTempDatums) = temp(4)
                                    End If
    
                            '------------------------------------------
                            'The unit defaults file contains the default
                            'information for each unit.
                            '------------------------------------------
                            Case "UNITFILE"
    
                            '------------------------------------------
                            'These are the fields stored in the unit
                            'defaults file.  Make a packed string of them
                            '------------------------------------------
                            Case "UNITFIELDS"
                                UnitFieldString = dt
'                                options(17) = ""
'                                dt = UCase(dt)
'                                Do Until Len(dt) = 0
'                                    E = InStr(dt, " ")
'                                    If E = 0 Then E = Len(dt) + 1
'                                    If options(17) = "" Then
'                                        options(17) = Left(dt, E - 1)
'                                    Else
'                                        options(17) = options(17) + "," + Left(dt, E - 1)
'                                    End If
'                                    dt = LTrim(Mid(dt, E))
'                                Loop
'
'                                If options(17) <> "" Then
'                                    b = InStr(options(17), ",")
'                                    If b = 0 Then
'                                        UnitName = Trim(options(17))
'                                    Else
'                                        UnitName = Trim(Left(options(17), b - 1))
'                                    End If
'                                Else
'                                    UnitName = "unit"
'                                End If

                            '------------------------------------------
                            'Set limits for units
                            '------------------------------------------
                            Case "LIMIT"
                                    NTempUnits = NTempUnits + 1
                                    dt = UCase(dt)
                                    F = 0
                                    Do Until Len(dt) = 0
                                            E = InStr(dt, " ")
                                            If E = 0 Then E = Len(dt) + 1
                                            F = F + 1
                                            temp(F) = Left(dt, E - 1)
                                            dt = LTrim(Mid(dt, E))
                                    Loop
                                    TempUnitName(NTempUnits) = temp(1)
                                    
                                    Select Case temp(2)
                                    Case "RECT"
                                            If F = 6 Then
                                                TempUnitMinX(NTempUnits) = CDbl(temp(3))
                                                TempUnitMinY(NTempUnits) = CDbl(temp(4))
                                                TempUnitMaxX(NTempUnits) = CDbl(temp(5))
                                                TempUnitMaxY(NTempUnits) = CDbl(temp(6))
                                            End If
    
                                    End Select
    
                            '------------------------------------------
                            ' Optionally put the sitename here.
                            '------------------------------------------
                            Case "SITE"
                                    SiteName = dt
    
                            '------------------------------------------
                            ' For now this variable just indicates
                            ' whether there is a printer.  In the
                            ' future, it may say what kind of printer
                            ' (ie. ir printer or normal printer.
                            '------------------------------------------
                            Case "PRINTER"
                                    Options(19) = UCase(dt)
    
                            Case "TAMISRATE"
    
                            Case "VHDORXYZ"
                                    Options(21) = UCase(dt)
    
                            '--------------------------------------------------------
                            ' If we reach this point, the line must begin with either
                            ' a variable or unit name.  Otherwise it is an error.
                            '--------------------------------------------------------
                            Case Else
    
                                    ERRORFLAG = 1
                                    '-------------------------------------
                                    ' First check to see if its an already
                                    ' defined variable.
                                    '-------------------------------------
                                    For C = 1 To Vars
    
                                            'check if defined variable
    
                                            If VarList(C) = key Then
                                                    B = InStr(dt, " ") ' look for space
                                                    If B = 0 Then B = Len(dt) + 1 'no value list
    
                                                    '-------------------------------------
                                                    'variable name and data
                                                    '-------------------------------------
                                                    Var = UCase(LTrim(RTrim(Left(dt, B - 1))))
                                                    dat = LTrim(RTrim(Mid(dt, B + 1)))
    
                                                    '-------------------------------------
                                                     ' check for possible commands
                                                    '-------------------------------------
                                                    Select Case Var
                                                    Case "INPUT"
                                                            VType(C) = UCase(dat)
                                                            Select Case VType(C)
                                                                Case "TEXT", "NUMERIC", "INSTRUMENT", "MENU", "UNIT"
                                                                '---------------------
                                                                'For CC system only.
                                                                '---------------------
                                                                Case "SQID"
                                                                        VType(C) = 6
                                                                        VLen(C) = 11
                                                            End Select
    
                                                    Case "DEFAULT"
                                                            VDefault(C) = dat
    
                                                    Case "PRINT"
                                                            dat = UCase(Left(dat, 1))
                                                            If dat = "Y" Then VPrint(C) = dat
    
                                                    Case "PROMPT"
                                                            VPrompt(C) = dat
                                                            If Len(VPrompt(C)) > 60 Then VPrompt(C) = Left(VPrompt, 60)
    
                                                    Case "MENULIST"
                                                            VMenu(C) = ""
                                                            menuitems = 0
                                                            Do Until Len(dat) = 0
                                                                    E = InStr(dat, " ")
                                                                    If E = 0 Then E = Len(dat) + 1
                                                                    If VMenu(C) = "" Then
                                                                        VMenu(C) = UCase(Left(dat, E - 1))
                                                                    Else
                                                                        VMenu(C) = VMenu(C) + "," + UCase(Left(dat, E - 1))
                                                                    End If
                                                                    dat = LTrim(Mid(dat, E))
                                                                    menuitems = menuitems + 1
                                                            Loop
    
                                                    Case "VARLEN"
                                                            VLen(C) = dat
    
                                                    Case "CARRY"
                                                            VCarry(C) = 1
    
                                                    Case "INCREMENT"
                                                            VIncr(C) = 1
    
                                                    Case "VARLOC"
                                                    
                                                    End Select
    
                                                    'If VPrompt(c) = "" Then
                                                    '        If vlen(c) + Len(varlist(c)) + 2 + varloc(c, 2) > display.Width Then
                                                    '                errorcode = 90
                                                    '                errormessage = "Field name + length is too long."
                                                    '                Exit Do
                                                    '        End If
    
                                                    'ElseIf Len(VPrompt(c)) + vlen(c) + varloc(c, 2) > display.Width Then
                                                     '       errorcode = 90
                                                    '        errormessage = "Field name + prompt is too long."
                                                    '        Exit Do
    
                                                    'End If
                                                    ERRORFLAG = 0
                                                    Exit For
                                            End If
                                    Next C
                            End Select
                    End If
            End If
    Loop
    Close 1
    
    'For a = 1 To vars
    '    If varlist(a) = "DATE" Then
    '        varlist(a) = "DAY"
    '        Exit For
    '    End If
    'Next a
    
'    If errorcode = 0 Then
'        Screen.MousePointer = 1
'        answer = MsgBox("This file will be converted to EDMCE format.", vbOKCancel, "EDMCE")
'        If answer = 1 Then
'            needtoconvert = 1
'            Screen.MousePointer = 11
'        Else
'            errorcode = 100
'            errormessage = "The configuration file must be converted to EDMCE format for this program to work."
'        End If
'    End If
     frmConvertCFG.Show 1

Else
    
    IniClass = "[EDM]"
    Inidata(1, 1) = "Sitename"
    Inidata(2, 1) = "Database"
    Inidata(3, 1) = "PointTable"
    Inidata(4, 1) = "Instrument"
    Inidata(6, 1) = "COMport"
    Inidata(7, 1) = "SQID"
    Inidata(8, 1) = "Unitfields"
    Inidata(9, 1) = "Limitchecking"
    Inidata(10, 1) = "PrismPrompt"
    Inidata(11, 1) = "StationName"
    Inidata(12, 1) = "StationX"
    Inidata(13, 1) = "stationY"
    Inidata(14, 1) = "stationZ"
    Inidata(15, 1) = "Settings"
    Inidata(16, 1) = "UpdateAlerts"
    Call ReadIni(CFGName, IniClass, Inidata(), Status)
    
    SiteName = Trim(Inidata(1, 2))
    SiteDBname = Trim(Inidata(2, 2))
    PointTableName = Trim(Inidata(3, 2))
    If LCase(Inidata(7, 2)) = "y" Then SqidCheck = True Else SqidCheck = False
    UnitFieldString = Inidata(8, 2)
    If LCase(Inidata(9, 2)) = "yes" Then
        mdiMain.mnuFindUnit.Checked = True
        frmMain.lblAutoFind.Visible = True
        LimitChecking = True
    Else
        mdiMain.mnuFindUnit.Checked = False
        LimitChecking = False
        frmMain.lblAutoFind.Visible = False
    End If
    If LCase(Inidata(10, 2)) = "yes" Then
        mdiMain.mnuPrismPrompt.Checked = True
    Else
        mdiMain.mnuPrismPrompt.Checked = False
    End If
    EDMName = ""
    If Inidata(4, 2) <> "" Then
        Select Case UCase(Inidata(4, 2))
        Case "TOPCON", "SOKKIA", "WILD", "NONE", "SIMULATE"
            Options(4) = Inidata(4, 2)
            EDMName = Inidata(4, 2)
        Case Else
            errorcode = 21
            errormessage = "Unrecognized EDM type.  Recognized types are Topcon, Wild, Sokkia and None."
        End Select
    Else
        Options(4) = "NONE"
    End If
    If LCase(Inidata(16, 2)) = "yes" Then
        NoAlert = False
        mdiMain.mnuNoAlert.Checked = False
    Else
        NoAlert = True
        mdiMain.mnuNoAlert.Checked = True
    End If
    
    Options(3) = Inidata(5, 2)
    Options(2) = Inidata(6, 2)
    comport = Inidata(6, 2)
    comsettings = Inidata(15, 2)
    If (EDMName = "" And LCase(EDMName) <> "simulate") And (comport = "" Or comsettings = "") Then
        frmMain.lblEDMWarning.Visible = True
    Else
        frmMain.lblEDMWarning.Visible = False
    End If
    Options(17) = Inidata(8, 2)
        
    If Options(17) <> "" Then
        B = InStr(Options(17), ",")
        If B = 0 Then
            UnitName = Trim(Options(17))
    Else
            UnitName = Trim(Left(Options(17), B - 1))
        End If
    Else
        UnitName = "unit"
    End If

If Not StationInitialized Then
    If Inidata(12, 2) <> "" And Inidata(13, 2) <> "" And Inidata(14, 2) <> "" Then
        TempString = "Continue with station " + Inidata(11, 2) + Chr(13)
        TempString = TempString + "   X: " + Format(Val(Inidata(12, 2)), "#####0.000") + Chr(13)
        TempString = TempString + "   Y: " + Format(Val(Inidata(13, 2)), "#####0.000") + Chr(13)
        TempString = TempString + "   Z: " + Format(Val(Inidata(14, 2)), "#####0.000") + Chr(13)
        
        response = MsgBox(TempString, vbYesNo)
        If response = vbYes Then
            CurrentStation.Name = Inidata(11, 2)
            CurrentStation.X = Val(Inidata(12, 2))
            CurrentStation.y = Val(Inidata(13, 2))
            CurrentStation.z = Val(Inidata(14, 2))
            frmMain.lblStationWarning.Visible = False
            StationInitialized = True
        Else
            For I = 1 To 100
                Inidata(I, 1) = ""
                Inidata(I, 2) = ""
            Next I
            Inidata(1, 1) = "StationName"
            Inidata(2, 1) = "StationX"
            Inidata(3, 1) = "stationY"
            Inidata(4, 1) = "stationZ"
            Call WriteIni(CFGName, IniClass, Inidata(), Status)
            
            StationInitialized = False
            frmMain.lblStationWarning.Visible = True
        End If
    Else
        StationInitialized = False
        frmMain.lblStationWarning.Visible = True
    End If
Else
    frmMain.lblStationWarning.Visible = False
    
End If

    'now need to load variables
    'First, need to read through CFG looking for the field names
    
    Open CFGName For Input As 1
    Vars = 0
    Do While Not EOF(1)
        Do
            Line Input #1, ts
            ts = Trim(ts)
            If Left(ts, 1) = "[" Then
                ts = UCase(ts)
                Select Case ts
                Case "[EDM]", "[BUTTON1]", "[BUTTON2]", "[BUTTON3]", "[BUTTON4]", "[BUTTON5]", "[BUTTON6]"
                Case Else
                    Vars = Vars + 1
                    VarList(Vars) = ts
                End Select
            End If
        Loop Until EOF(1)
    Loop
    Close 1
    
    For a = 1 To 50
        Inidata(a, 1) = ""
        Inidata(a, 2) = ""
    Next a
    Inidata(1, 1) = "Type"
    Inidata(2, 1) = "Prompt"
    Inidata(3, 1) = "Menu"
    Inidata(4, 1) = "Length"
    Inidata(5, 1) = "Increment"
    Inidata(6, 1) = "Carry"
    Inidata(7, 1) = "Unique"
    
    For C = 1 To Vars
        VCarry(C) = False
        VIncr(C) = False
        VUnique(C) = False
        For a = 1 To 7
            Inidata(a, 2) = ""
        Next a
        IniClass = VarList(C)
        Call ReadIni(CFGName, IniClass, Inidata, Status)
        
        'strip the [ ]
        VarList(C) = Mid(VarList(C), 2)
        VarList(C) = Left(VarList(C), Len(VarList(C)) - 1)
        
        VType(C) = UCase(Inidata(1, 2))
        VPrompt(C) = Inidata(2, 2)
        VMenu(C) = Inidata(3, 2)
    
        If Inidata(4, 2) <> "" Then VLen(C) = CInt(Inidata(4, 2))
        If UCase(VarList(C)) = "ID" Then IDLength = VLen(C)
        If UCase(VarList(C)) = "UNIT" Then UnitLength = VLen(C)
    
        If LCase(Inidata(5, 2)) = "true" Or LCase(Inidata(5, 2)) = "yes" Then VIncr(C) = True
    
        If LCase(Inidata(6, 2)) = "true" Or LCase(Inidata(6, 2)) = "yes" Then VCarry(C) = True
    
        If LCase(Inidata(7, 2)) = "true" Or LCase(Inidata(7, 2)) = "yes" Then VUnique(C) = True
    
    Next C
    For I = 1 To 6
        For J = 1 To Vars + 1
            Inidata(J, 1) = ""
            Inidata(J, 2) = ""
        Next J
        IniClass = "[BUTTON" + Trim(Str(I)) + "]"
        For J = 1 To Vars
            Inidata(J, 1) = VarList(J)
        Next J
        Inidata(Vars + 1, 1) = "Title"
        Call ReadIni(CFGName, IniClass, Inidata, Status)
        nButtonVars(I) = 0
        frmMain.Button(I).Visible = False
        Gotit = False
        For J = 1 To Vars
            If Inidata(J, 2) <> "" Then
                Gotit = True
                nButtonVars(I) = nButtonVars(I) + 1
                ButtonVars(I, nButtonVars(I), 1) = J
                ButtonVars(I, nButtonVars(I), 2) = UCase(Inidata(J, 2))
            End If
        Next J
        If Gotit Then
            ButtonCaption(I) = Inidata(Vars + 1, 2)
            frmMain.Button(I).Visible = True
            frmMain.Button(I).Caption = "&" + Trim(Str(I)) + " " + ButtonCaption(I)
            ButtonCaption(I) = Inidata(Vars + 1, 2)
        End If
        
        
    Next I
    
End If

Close 1
frmMain.FormatVarList
If SiteDBname <> "" Then
    If InStr(SiteDBname, ".") > 0 Then
        SiteDBname = Left(SiteDBname, Len(SiteDBname) - 4)
    End If
    SiteDBname = SiteDBname + ".mdb"
    TempString = ""
    On Error Resume Next
    TempString = Dir(SiteDBname)
    On Error GoTo 0
    If TempString <> "" Then
        Call OpenSite(SiteDBname$)
        If PointTableName <> "" Then OpenPointsTable
        ParseUnitFields UnitFieldString
        frmMain.lblDBWarning.Visible = False
    Else
        MsgBox ("Database given in " + CFGName + " not found.")
        SiteDBname = ""
        frmMain.lblDBWarning.Visible = True
    
    End If
End If
frmMain.lblCFGWarning.Visible = False
addtofilelist CFGName

End Sub
Sub parseangle(hangle As Single, angle As Integer, minutes As Integer, seconds As Integer)

Dim B As Integer

a$ = Str$(hangle)
B = InStr(a$, ".")

If B = 0 Then
        angle = Val(a$)
        minutes = 0
        seconds = 0
Else
        a$ = a$ + "0000"
        If Val(Left$(a$, B - 1)) < 32000 Then angle = Val(Left$(a$, B - 1)) Else errorcode = -1
        If Val(Mid$(a$, B + 1, 2)) < 32000 Then minutes = Val(Mid$(a$, B + 1, 2)) Else errorcode = -1
        If Val(Mid$(a$, B + 3, 2)) < 32000 Then seconds = Val(Mid$(a$, B + 3, 2)) Else errorcode = -1
End If

End Sub


Public Sub OpenPointsTable()
Dim Xlength As Integer
Dim Gotit As Boolean

frmMain.lblBlankFields.Visible = False
frmMain.txtCurrentRecord = 0
frmMain.txtTotalRecords = 0
Set PointsTB = Nothing
If Not tablematch(PointTableName) Then
    PointTableName = ""
    response = MsgBox("Points table not present.  Create one now?", vbYesNo)
    If response = vbYes Then
        mdiMain.mnuNewPointsTB_Click
        frmMain.lblPointsWarning.Visible = False
    Else
        frmMain.lblPointsWarning.Visible = True
    End If
    Exit Sub
End If

For Each a In SiteDB.TableDefs(PointTableName).Fields
    Gotit = False
    If LCase(a.Name) <> "edm_reccounter" Then
        For I = 1 To Vars
            If UCase(a.Name) = UCase(VarList(I)) Then
                Gotit = True
                Exit For
            End If
        Next I
        If Gotit = False Then
            GoTo BadTable
        End If
    End If
Next
For I = 1 To Vars
    On Error Resume Next
    Gotit = False
    Gotit = UCase(VarList(I)) = UCase(SiteDB.TableDefs(PointTableName).Fields(VarList(I)).Name)
    On Error GoTo 0
    If Not Gotit Then
        GoTo BadTable
    End If
    If UCase(VType(I)) = "TEXT" Or UCase(VType(I)) = "MENU" Then
        If VLen(I) <> SiteDB.TableDefs(PointTableName).Fields(VarList(I)).Size Then
            FixFieldSize I
        End If
    End If
Next I
    

frmMain.txtPT = PointTableName
For I = 1 To Vars
    Select Case UCase(VType(I))
        Case "MENU", "TEXT"
            If SiteDB.TableDefs(PointTableName).Fields(VarList(I)).Size <> VLen(I) Then
                Dim tdf As TableDef
                Dim F As Field
                Set tdf = SiteDB.TableDefs(PointTableName)
                Set F = tdf.CreateField("$$Temp$$", dbText, VLen(I))
                F.AllowZeroLength = True
                tdf.Fields.Append F
                SqlString = "update " + PointTableName + " set [$$temp$$]=" + VarList(I)
                SiteDB.Execute SqlString
                SiteDB.TableDefs(PointTableName).Fields.Delete VarList(I)
                SiteDB.TableDefs(PointTableName).Fields("$$temp$$").Name = VarList(I)
                Set tdf = Nothing
                Set F = Nothing
            End If
        Case "NUMERIC", "INSTRUMENT"
    End Select
Next I
Set PointsTB = SiteDB.OpenRecordset(PointTableName, dbOpenDynaset)
If Not PointsTB.EOF Then
    PointsTB.MoveLast
End If
Loading = True
frmMain.ShowValues
Loading = False
frmMain.lblPointsWarning.Visible = False
Dim IniClass As String
Dim Inidata(1, 2) As String
Dim Status As Byte

IniClass = "[EDM]"
Inidata(1, 1) = "pointtable"
Inidata(1, 2) = PointTableName
Call WriteIni(CFGName, IniClass, Inidata(), Status)

Exit Sub

BadTable:
MsgBox ("Points table and CFG file do not match.  Select different Points table or create one.")
PointTableName = ""
frmMain.lblPointsWarning.Visible = True
frmMain.txtPT = ""

    
End Sub

Public Sub ParseUnitFields(UnitFieldString As String)
Dim X As Integer

nUnitFields = 0

UnitFieldString = Trim(UnitFieldString)
X = InStr(UnitFieldString, ",")
Do While X > 0
    nUnitFields = nUnitFields + 1
    Unitfield(nUnitFields) = Left(UnitFieldString, X - 1)
    UnitFieldString = Trim(Mid(UnitFieldString, X + 1))
    X = InStr(UnitFieldString, ",")
Loop
If Trim(UnitFieldString) <> "" Then
    nUnitFields = nUnitFields + 1
    Unitfield(nUnitFields) = Trim(UnitFieldString)
End If


On Error Resume Next
For I = 1 To nUnitFields
    Gotit = False
    Gotit = UCase(Unitfield(I)) = UCase(SiteDB.TableDefs("EDM_units").Fields(Unitfield(I)).Name)
    If Not Gotit Then GoTo BadTable:
Next I
For Each a In SiteDB.TableDefs("EDM_units").Fields
    Select Case UCase(a.Name)
        Case "MINX", "MINY", "MAXX", "MAXY", "CENTERX", "CENTERY", "RADIUS"
        Case Else
            Gotit = False
            For I = 1 To Vars
                If UCase(a.Name) = UCase(VarList(I)) Then
                    Gotit = True
                    Exit For
                End If
            Next I
            If Not Gotit Then
                GoTo BadTable
            End If
    End Select
Next

On Error GoTo 0
Exit Sub

BadTable:
On Error GoTo 0
If Loading Then
    response = MsgBox("Units table does not match list of required fields in CFG file.  Update now?", vbYesNo)
    If response = vbYes Then
        UpdateUnitsTable
    End If
Else
    UpdateUnitsTable
End If

Exit Sub




End Sub

Public Function FindColumn(Varname As String)
FindColumn = -1
For I = 0 To PointsTB.Fields.Count - 1
    If LCase(PointsTB.Fields(I).Name) = LCase(Varname) Then
        FindColumn = I
        Exit For
    End If
Next I


End Function

Public Function PadID(idstring As String)
PadID = Right(Space(IDLength) + Trim(idstring), IDLength)
End Function

Public Sub GetPointTables()

Dim Parity As Byte
Dim a As Boolean, B As Boolean, C As Boolean, d As Boolean, E As Boolean, F As Boolean, G As Boolean, H As Boolean, M As Boolean, N As Boolean, O As Boolean, P As Boolean
Dim NPtable As Boolean
nPlotFiles = 0
nPointTables = 0
nOverlayTables = 0
nWorkBooks = 0
nTempWorkBooks = 0
nLookupTables = 0

SiteDB.TableDefs.Refresh
For Each tdtemp In SiteDB.TableDefs
    a = False
    If Left(tdtemp.Name, 7) <> "WinPlot" And Left(tdtemp.Name, 4) <> "MSys" And Left(tdtemp.Name, 2) <> "$$" And Left(tdtemp.Name, 4) <> "EDM_" Then
        On Error Resume Next
        a = SiteDB.TableDefs(tdtemp.Name).Fields("EDM_reccounter").Size > 0
        If a Then
            nPointTables = nPointTables + 1
            PointTable(nPointTables) = LCase(tdtemp.Name)
        End If
    End If
Next
On Error GoTo 0

End Sub

Public Function GetPath(PathString As String)

GetPath = ""
If PathString = "" Then
    Exit Function
End If
For I = Len(PathString) To 1 Step -1
    If Mid(PathString, I, 1) = "\" Then
        GetPath = Left(PathString, I - 1)
        Exit Function
    End If
Next I


End Function

Public Sub UpdateUnitsTable()


Dim tdf As TableDef
Dim F As Field

Set UnitTB = Nothing
Set tdf = SiteDB.TableDefs("EDM_Units")
For I = 1 To nUnitFields
    Gotit = False
    On Error Resume Next
    Gotit = UCase(Unitfield(I)) = UCase(tdf.Fields(Unitfield(I)).Name)
    On Error GoTo 0
    If Not Gotit Then
        For J = 1 To Vars
            If UCase(Unitfield(I)) = UCase(VarList(J)) Then
                If VLen(J) = 0 Then VLen(J) = 20
                Select Case VType(J)
                    Case "TEXT", "MENU", "UNIT", SQID
                        Set F = tdf.CreateField(VarList(J), dbText, VLen(J))
                        tdf.Fields.Append F
                    Case "NUMERIC", "INSTRUMENT"
                        Set F = tdf.CreateField(VarList(J), dbDouble)
                        tdf.Fields.Append F
                End Select
                Exit For
            End If
        Next J
    End If
Next I
On Error GoTo 0


KeepLooking:
SiteDB.TableDefs.Refresh
For Each a In SiteDB.TableDefs("EDM_units").Fields
    Select Case UCase(a.Name)
        Case "UNIT", "ID", "MINX", "MINY", "MAXX", "MAXY", "CENTERX", "CENTERY", "RADIUS"
            Gotit = True
        Case Else
            Gotit = False
            For I = 1 To Vars
                If UCase(a.Name) = UCase(VarList(I)) Then
                    Gotit = True
                    Exit For
                End If
            Next I
            If Not Gotit Then
                SiteDB.TableDefs("EDM_units").Fields.Delete a.Name
                GoTo KeepLooking
            End If
    End Select
Next

Set tdf = Nothing
SiteDB.TableDefs.Refresh
nUnitFields = 0
For Each a In SiteDB.TableDefs("EDM_Units").Fields
    Select Case UCase(a.Name)
        Case "MINX", "MINY", "MAXX", "MAXY", "CENTERX", "CENTERY", "RADIUS"
        Case Else
            nUnitFields = nUnitFields + 1
            Unitfield(nUnitFields) = a.Name
    End Select
Next
Dim IniClass As String
Dim Inidata(1, 2) As String
Dim Status As Byte

IniClass = "[EDM]"
Inidata(1, 1) = "Unitfields"
Inidata(1, 2) = Unitfield(1)
For I = 2 To nUnitFields
    Inidata(1, 2) = Inidata(1, 2) + "," + Unitfield(I)
Next I
Call WriteIni(CFGName, IniClass, Inidata(), Status)

Set UnitTB = SiteDB.OpenRecordset("EDM_Units")

End Sub

Public Sub FixFieldSize(VarNum)
Set PointsTB = Nothing
Set tdf = SiteDB.TableDefs(PointTableName)

Set F = tdf.CreateField("*temp", dbText, VLen(VarNum))
tdf.Fields.Append F
SqlString = "update " + PointTableName + " set [*temp]=" + VarList(VarNum)
SiteDB.Execute SqlString
SiteDB.TableDefs(PointTableName).Fields.Delete VarList(VarNum)
SiteDB.TableDefs(PointTableName).Fields("*temp").Name = VarList(VarNum)
Set tdf = Nothing
Set F = Nothing


End Sub

Public Sub takeshot_core(edmshot As shotdata, errorcode)

Dim edmpoffset As Single

'Note - to use this routine - edmshot.poleh and edmshot.poleo have
'to be set prior to coming here.

Select Case UCase(EDMName)
Case "SIMULATE"
    Randomize
    X = Int((3) * Rnd) + 999 + Rnd
    y = Int((1) * Rnd) + 1013 + Rnd
    z = Rnd
    edmshot.X = X
    edmshot.y = y
    edmshot.z = z
    edmshot.hangle = 111.505
    edmshot.vangle = 98.2525
    edmshot.sloped = Sqr(edmshot.X ^ 2 + edmshot.y ^ 2 + edmshot.z ^ 2)

Case Else
    Call recordpoint(returndata$)
    Call parsenez(returndata$, edmshot, edmpoffset, mesunits$, angleunit$, errorcode)
    If errorcode = 0 Then
        Call vhdtonez(edmshot)
    Else
        Select Case errorcode
        Case 1, 5
            MsgBox "ERROR: Instrument was not in Vertical distance mode.  Reset instrument.", vbCritical
        Case 6
            MsgBox "ERROR: Parity error.  Shot must be retaken.", vbCritical
        Case 2
            MsgBox "ERORR: Angles not in degrees.  Reset instrument.", vbCritical
        Case Else
            MsgBox "ERROR: Unknown error.  Shot must be retaken.", vbCritical
        End Select
        Exit Sub
    End If
End Select

edmshot.X = CurrentStation.X + edmshot.X
edmshot.y = CurrentStation.y + edmshot.y
edmshot.z = CurrentStation.z + edmshot.z - edmshot.poleh

End Sub
