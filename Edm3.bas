Attribute VB_Name = "EDMWindows"
Type shotdata
    X As Single
    y As Single
    z As Single
    sloped As Single
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

Public Sub SpeakID(CurrentSquare, CurrentID)

Dim TempID

TempID = CurrentSquare + " dash "
CurrentID = Trim(CurrentID)

If Val(CurrentID) = 0 Then
        TempID = TempID + " " + Left(CurrentID, 1) + " " + Mid(CurrentID, 2, 1) + " " + Mid(CurrentID, 3, 1) + " " + Mid(CurrentID, 4, 1) + " " + Right(CurrentID, 1)
Else
    Select Case Len(CurrentID)
        Case 1, 2
            TempID = TempID + " " + CurrentID
        Case 3
            TempID = TempID + " " + Left(CurrentID, 1) + " " + Right(CurrentID, 2)
        Case 4
            TempID = TempID + " " + Left(CurrentID, 2) + " " + Right(CurrentID, 2)
        Case 5
            TempID = TempID + " " + Left(CurrentID, 2) + " " + Mid(CurrentID, 3, 1) + " " + Right(CurrentID, 2)
    End Select
End If
Voice.Speak TempID, SVSFlagsAsync

End Sub

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
Set ContextTableDef = Nothing
Set ContextIndex = Nothing
Set ContextField = Nothing

End Sub

Public Sub CreateXYZ()
     
XYZTableName = "XYZ"
Set xyztabledef = SiteDB.CreateTableDef(XYZTableName)
Set xyzfield = xyztabledef.CreateField("Unit", dbText, UnitLength)
xyzfield.AllowZeroLength = True
xyztabledef.Fields.Append xyzfield
Set xyzfield = xyztabledef.CreateField("ID", dbText, IDLength)
xyzfield.AllowZeroLength = True
xyztabledef.Fields.Append xyzfield
Set xyzfield = xyztabledef.CreateField("Suffix", dbInteger)
xyztabledef.Fields.Append xyzfield
Set xyzfield = xyztabledef.CreateField("X", dbSingle)
xyztabledef.Fields.Append xyzfield
Set xyzfield = xyztabledef.CreateField("Y", dbSingle)
xyztabledef.Fields.Append xyzfield
Set xyzfield = xyztabledef.CreateField("Z", dbSingle)
xyztabledef.Fields.Append xyzfield
Set xyzfield = xyztabledef.CreateField("RecordCounter", dbLong)
xyzfield.Attributes = dbAutoIncrField
xyztabledef.Fields.Append xyzfield
For I = 1 To Vars
    Select Case UCase(VarList(I))
        Case "PRISM", "VANGLE", "HANGLE", "SLOPED"
            Set xyzfield = xyztabledef.CreateField(VarList(I), dbSingle)
            xyztabledef.Fields.Append xyzfield
    End Select
Next I
    

SiteDB.TableDefs.Append xyztabledef
Set xyzindex = SiteDB.TableDefs(XYZTableName).CreateIndex("SqidIndex")
With xyzindex
    .Fields = "unit;ID"
    .Primary = False
    .Required = False
    .Unique = False
End With
SiteDB.TableDefs(XYZTableName).Indexes.Append xyzindex

Set xyzindex = xyztabledef.CreateIndex("SqidSuffixIndex")
With xyzindex
    .Fields = "unit;ID;Suffix"
    .Primary = False
    .Required = False
    .Unique = False
End With
SiteDB.TableDefs(XYZTableName).Indexes.Append xyzindex

Set xyzindex = xyztabledef.CreateIndex("RecordCounter")
With xyzindex
    .Fields = "RecordCounter"
    .Primary = False
    .Required = False
    .Unique = False
End With
SiteDB.TableDefs(XYZTableName).Indexes.Append xyzindex
Set xyztabledef = Nothing
Set xyzfield = Nothing
Set xyzindex = Nothing

End Sub

Sub clearcom()

On Error Resume Next
Do
    A$ = frmMain.theoport.Input
Loop Until A$ = ""

End Sub

Sub delay(DelayTime As Single)

Dim t1 As Double

t1 = Timer
Do Until (Timer - t1) > DelayTime Or Cancelling = True
    DoEvents
Loop

End Sub

Sub directoutput(A$)

frmMain.theoport.Output = A$

If frmDebug.Visible Then
    frmDebug.txtOutput = frmDebug.txtOutput + "->" + ack$
End If

End Sub

Sub displayvertangle()

Select Case LCase(EDMName$)
Case "topcon"
    d$ = "Z20"
    Call edmoutput(d$, errorcode)
Case Else
End Select

End Sub

Sub edmack()

If EDMName$ = "Topcon" Then
    ack$ = Chr$(6) + "006" + Chr$(3) + Chr$(13) + Chr$(10)
    frmMain.theoport.Output = ack$
End If

If frmDebug.Visible Then
    frmDebug.txtOutput = frmDebug.txtOutput + "->" + ack$
End If

End Sub

Sub edminput(response As String)
                                                         
Dim n As Integer
Dim fileno As Integer

On Error Resume Next

timeone = Timer

response = ""
t$ = ""
Do
    t$ = frmMain.theoport.Input
    If t$ <> "" Then response = response + t$
    If InStr(response, Chr$(13) + Chr$(10)) <> 0 Then
        n = InStr(response, Chr$(13) + Chr$(10))
        response = Left$(response, n + 1)
        Exit Do
    End If
    DoEvents
Loop Until Right$(response, 2) = Chr$(13) + Chr$(10) Or Timer - timeone > 30 Or Cancelling

If Not Cancelling And Not Right$(response, 2) = Chr$(13) + Chr$(10) Then
    Cancelling = True
    response = "CANCEL"
End If

' Code added by SPM in 2/2018
' Could affect timing of communications and fail.
' One option would be to open log file and leave it open.
' But I prefer to open and close when possible.
If TSLog Then
    If TSLogFile <> "" Then
        fileno = FreeFile
        Open TSLogFile For Append As #fileno
            Print #fileno, "Received->" & response & "<-"
        Close fileno
    End If
End If

If frmDebug.Visible Then
    frmDebug.txtOutput = frmDebug.txtOutput + "<-" + response
End If

Exit Sub

debug_it:
    MsgBox ("response '" + response + "' and t '" + t$ + "'" + " error is " + Err.Description + "COM Port is " + Str(frmMain.theoport.CommPort) + " Settings are " + frmMain.theoport.Settings + " and comport is open " + Str(frmMain.theoport.PortOpen = True))
    Exit Sub


End Sub

Sub edmoutput(d$, errorcode)

Dim transmit As String

term$ = Chr$(13) + Chr$(10)
errorcode = 0
Select Case LCase(EDMName$)
Case "nikon"
    Call makebccnikon("CT" + Chr$(2) + d$ + Chr$(3), bcc$)
    transmit = Chr$(1) + "CT" + Chr$(2) + d$ + Chr$(3) + bcc$ + Chr$(4) + term$
Case "topcon"
    Call makebcc(d$, bcc$)
    transmit = d$ + bcc$ + Chr$(3) + term$
Case "wild", "leica", "builder"
    transmit = d$ + term$
Case "wild2"
    transmit = Chr$(10) + d$ + term$
Case "sokkia"
    transmit = d$
Case Else
    Exit Sub
End Select

Select Case LCase(EDMName$)
Case "topcon", "wild", "wild2", "sokkia", "leica", "nikon", "builder"
    If frmMain.theoport.PortOpen Then
        frmMain.theoport.Output = transmit
    End If
Case Else
End Select

' Code added by SPM in 2/2018
' Could affect timing of communications and fail.
' One option would be to open log file and leave it open.
' But I prefer to open and close when possible.
If TSLog Then
    If TSLogFile <> "" Then
        fileno = FreeFile
        Open TSLogFile For Append As #fileno
            Print #fileno, "Sent->" & transmit & "<-"
        Close fileno
    End If
End If

If frmDebug.Visible Then
    frmDebug.txtOutput = frmDebug.txtOutput + "->" + transmit
End If
    
End Sub

Sub getpresetdata(hangle As Single, X As Single, y As Single, z As Single, angleunit$, mesunits$, errormessage$)

errormessage$ = ""
If display.edmtype <> 1 Then Exit Sub

d$ = "L"
Call edmoutput(d$, errorcode)
Call edminput(B$)

Do
    Call edminput(A$)
Loop Until InStr(A$, "L") <> 0 Or A$ = "CANCEL"

If A$ <> "CANCEL" Then
    ack$ = Chr$(6) + "006" + Chr$(3) + Chr$(13) + Chr$(10)
    Call delay(0.05)
    Call directoutput(ack$)
    returndata$ = A$
    Call parsepreset(returndata$, hangle, X, y, z, angleunit$, mesunits$, errormessage$)
Else
    Call horizontal(errorcode)
End If

End Sub

Sub horizontal(errorcode)

Dim response As String

errorcode = 0
Select Case LCase(EDMName$)
Case "topcon"
    d$ = "Z10"
    Call edmoutput(d$, errorcode)
    Call edminput(response)
    If response = "CANCEL" Then errorcode = 27
Case Else
End Select

End Sub

Sub horizontalright()

Select Case LCase(EDMName$)
Case "topcon"
    d$ = "Z12"
    Call edmoutput(d$, errorcode)
Case Else
End Select

End Sub

Sub hortizontalleft()

Select Case LCase(EDMName$)
Case "topcon"
    d$ = "Z13"
    Call edmoutput(d$, errorcode)
Case Else
End Select

End Sub

Sub initcomport(comport, errorcode)

'--------------------------------------------------------------------------
'The switch to horizontal isn't really necessary.
'I do it just to make the EDM beep in acknowledgement of the connection.
'If no EDM is present, however, the program will hange until ESCape.
'--------------------------------------------------------------------------

If frmMain.theoport.PortOpen Then frmMain.theoport.PortOpen = False

If comport = "" Or comsettings = "" Then
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
Case "COM5"
    frmMain.theoport.CommPort = 5
Case "COM6"
    frmMain.theoport.CommPort = 6
Case "COM7"
    frmMain.theoport.CommPort = 7
Case "COM8"
    frmMain.theoport.CommPort = 8
Case "COM9"
    frmMain.theoport.CommPort = 9
Case "COM10"
    frmMain.theoport.CommPort = 10
Case "COM11"
    frmMain.theoport.CommPort = 11
Case "COM12"
    frmMain.theoport.CommPort = 12
Case Else
End Select

On Error GoTo porterror
frmMain.theoport.PortOpen = True
On Error GoTo 0

Call initedm
Call horizontal(errorcode)
Exit Sub

porterror:
MsgBox ("Error: Tried to open " + UCase(comport) + " with '" + comsettings + "' and received error of " + Err.Description)
Exit Sub

End Sub

Sub initedm()

Dim d$

Select Case LCase(EDMName$)
Case "nikon"
    d$ = "$DNX"     'Set instrument to output after each measure
                    'format should be slope distance, horizontal angle, vertical angle
    Call edmoutput(d$, errorcode)
    Call edminput(A$)   'not sure whether an ACK will be sent back - remove this line if not

Case "topcon"
    d$ = "ST0"
    Call edmoutput(d$, errorcode)

Case "wild", "leica", "builder"
    d$ = "SET/41/0"
    Call edmoutput(d$, errorcode)
    If errorcode = 0 Then
        Call edminput(A$)
    End If
    d$ = "SET/149/2"
    Call edmoutput(d$, errorcode)
    If errorcode = 0 Then
        Call edminput(A$)
    End If
    Call delay(0.5)
    Call clearcom

Case "wild2"
    Call edmoutput("%R1Q,0:", errorcode)    'Checks that communications is working
    If errorcode = 0 Then
        Call edminput(A$)
        If A$ <> "%R1P,0,0:RC" Then         ' This is what should be returned
            errorcode = 1                   ' This code currently ignores errorcodes (not good)
        End If                              ' but I am hesitant to change just now.
    End If

Case Else
End Select

End Sub

Sub makebccnikon(I$, o$)

Dim l As Integer
Dim t As Long

'From the Nikon manual,
'first sum the asc values of the string to be sent
t = 0
For l = 1 To Len(I$)
    t = t + Asc(Mid$(I$, l, 1))
Next l

'Then take the lower byte of sum and transform as follows
t = t And &HFF
t = (t Mod &H40) + &H20

o$ = Chr$(t)

End Sub

Sub makebcc(I$, o$)

Dim l As Integer

B = 0
For l = 1 To Len(I$)
    q = Asc(Mid$(I$, l, 1))
    b1 = q And (Not B)
    b2 = B And (Not q)
    B = b1 Or b2
Next l

o$ = LTrim$(Str$(B))
o$ = Right$("000" + o$, 3)

End Sub

Function radtodeg(radians_angle) As Double

Dim pi As Double

pi = 3.14159265359

radtodeg = radians_angle / (2 * pi) * 360

End Function

Sub convert_edmshot_to_dms()

' This is for GeoCOM stations.  These return the V and H angles in decimal radians
' To keep compatibility with the rest of the program, these are then converted to
' DMS. The way this is done it not the best, but it works.

edmshot.vangle = radtodeg(edmshot.vangle)
edmshot.vangle = decimaldegrees_to_dms(edmshot.vangle)

edmshot.hangle = radtodeg(edmshot.hangle)
edmshot.hangle = decimaldegrees_to_dms(edmshot.hangle)

End Sub


Function decimaldegrees_to_dms(angle) As Double

degrees = Int(angle)
seconds = (angle - degrees) * 3600
minutes = Int(seconds / 60)
seconds = Int(seconds - (60 * minutes))
decimaldegrees_to_dms = Val(Str(degrees) + "." + Format(minutes, "00") + Format(seconds, "00"))

End Function


Sub parsenez(nezdata$, edmshot As shotdata, edmpoffset As Single, mesunits$, angleunit$, errorcode)

Dim angle As Integer, minutes As Integer, seconds As Integer
Dim tangle As Single, dangle As Single, dist As Single

errorcode = 0

If nezdata$ = "" Then
    errorcode = -99
    Exit Sub
End If

Select Case LCase(EDMName$)
Case "nikon"
    If InStr(nezdata$, "TC") = 0 Then
        errorcode = 100
    Else
        edmshot.edmpoffset = 0
        A = InStr(nezdata$, "SD:")
        If A <> 0 Then
            edmshot.sloped = Val(Mid$(nezdata$, A + 3, 9)) / 10000
        Else
            errorcode = 101
        End If
        A = InStr(nezdata$, "HA#")
        If A = 0 Then A = InStr(nezdata$, "HA:")
        If A <> 0 Then
            edmshot.hangle = Val(Mid$(nezdata$, A + 3, 4) + "." + Mid$(nezdata$, A + 7, 5))
        Else
            errorcode = 102
        End If
        A = InStr(nezdata$, "VA#")
        If A = 0 Then A = InStr(nezdata$, "VA:")
        If A <> 0 Then
            edmshot.vangle = Val(Mid$(nezdata$, A + 3, 4) + "." + Mid$(nezdata$, A + 7, 5))
        Else
            errorcode = 103
        End If
    End If

Case "topcon", "simulate"
    Do Until Asc(Left$(nezdata$, 1)) > 32 Or Len(nezdata$) = 1
        nezdata$ = Mid$(nezdata$, 2)
    Loop
    A = Left$(nezdata$, 1)
    If A <> "?" And A <> "R" Then
        If A = "U" Then
            errorcode = 5
        Else
            errorcode = 1
        End If
        Exit Sub
    End If

    A = InStr(nezdata$, Chr$(3))
    If A <> 0 Then
        bcc1$ = Mid$(nezdata$, A - 3, 3)
        d$ = Left$(nezdata$, A - 4)
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

Case "wild2"
    '%R1P,0,0:RC,Hz[double],V[double],SlopeDistance[double]
    If Left$(nezdata$, 9) <> "%R1P,0,0:" Then
        errorcode = 1001
        Exit Sub
    End If
    nezdata$ = Mid$(nezdata$, 10)
    A = InStr(nezdata$, ",")
    If A <> 0 Then
        nezdata$ = Mid$(nezdata$, A + 1)        'The first parameter can be ignored
        A = InStr(nezdata$, ",")
        If A <> 0 Then
            edmshot.hangle = Val(Left(nezdata$, A - 1))
            nezdata$ = Mid$(nezdata$, A + 1)
            A = InStr(nezdata$, ",")
            If A <> 0 Then
                edmshot.vangle = Val(Left(nezdata$, A - 1))
                nezdata$ = Mid$(nezdata$, A + 1)
                edmshot.sloped = Val(nezdata$)
            Else
                errorcode = 1004
                Exit Sub
            End If
        Else
            errorcode = 1003
            Exit Sub
        End If
        Call convert_edmshot_to_dms
    Else
        errorcode = 1002
        Exit Sub
    End If
        
Case "wild"
    
    ' This code is for Lieca as well.  Modified May, 2021, to accept either GSI 8 or 16
    
    ' Look for hangle and remove everything before it
    A = InStr(nezdata$, "21.")
    If A = 0 Then
        errorcode = 1
        Exit Sub
    End If
    nezdata$ = Mid$(nezdata$, A)

    ' Check the hangle is in degrees minutes seconds
    If Mid$(nezdata$, 6, 1) <> "4" Then
        errorcode = 2
        Exit Sub
    End If
    nezdata$ = Mid$(nezdata$, 7)
    B = InStr(nezdata$, " ")
    hangle$ = Left(nezdata$, B - 1)
    edmshot.hangle = Val(Left$(hangle$, Len(hangle$) - 5) + "." + Right$(hangle$, 5))

    nezdata$ = Mid$(nezdata$, B + 1)
    ' Check the hangle is in degrees minutes seconds
    If Mid$(nezdata$, 6, 1) <> "4" Then
        errorcode = 2
        Exit Sub
    End If
    nezdata$ = Mid$(nezdata$, 7)
    B = InStr(nezdata$, " ")
    vangle$ = Left(nezdata$, B - 1)
    edmshot.vangle = Val(Left$(vangle$, Len(vangle$) - 5) + "." + Right$(vangle$, 5))

    mesunits$ = Mid$(nezdata$, B + 6)

    nezdata$ = Mid$(nezdata$, B + 7)
    B = InStr(nezdata$, " ")
    sloped$ = Left(nezdata$, B - 1)
    edmshot.sloped = Val(sloped$) / 1000
    If edmshot.sloped = 0 Then
        errorcode = 4
        Exit Sub
    End If
    edmshot.edmpoffset = Val(Right$(nezdata$, 4)) / 1000

Case "builder"
    A = InStr(nezdata$, "21.")
    If A = 0 Then
        errorcode = 1
        Exit Sub
    End If
    nezdata$ = Mid$(nezdata$, A)

    'IF LEN(nezdata$) <> 64 OR a = 0 THEN
    '        errorcode = 1
    '        EXIT SUB
    'END IF

    If Mid$(nezdata$, 6, 1) <> "4" Then
        errorcode = 2
        Exit Sub
    End If
    edmshot.hangle = Val(Mid$(nezdata$, 15, 4) + "." + Mid$(nezdata$, 19, 4))
    A = InStr(nezdata$, "22.")
    If Mid$(nezdata$, A + 5, 1) <> "4" Then
        errorcode = 2
        Exit Sub
    End If
    edmshot.vangle = Val(Mid$(nezdata$, A + 14, 4) + "." + Mid$(nezdata$, A + 18, 4))

    A = InStr(nezdata$, "31..0")
    edmshot.sloped = Val(Mid$(nezdata$, A + 14, 9)) / 1000
    If edmshot.sloped = 0 Then
        errorcode = 4
        Exit Sub
    End If
    mesunits$ = Mid$(nezdata$, A + 5)
    edmshot.edmpoffset = Val(Right$(nezdata$, 4)) / 1000

Case "sokkia"
    edmshot.edmpoffset = 0
    nezdata$ = RTrim$(LTrim$(nezdata$))
    A = InStr(nezdata$, " ")
    If A > 1 Then
        If A > 8 Then
            edmshot.sloped = Val(Mid(Left$(nezdata$, A - 1), 2)) / 1000
        Else
            edmshot.sloped = Val(Left$(nezdata$, A - 1)) / 1000
        End If
        If edmshot.sloped = 0 Then
            errorcode = 10
        Else
            v$ = Mid$(nezdata$, A + 1)
            A = InStr(v$, " ")
            If A > 1 And Left$(v$, 1) <> "E" Then
                edmshot.vangle = Val(Left$(v$, 3) + "." + Mid$(v$, 4, 4))
                v$ = Mid$(v$, A + 1)
                A = InStr(v$, " ")
                If A > 1 And Left$(v$, 1) <> "E" Then
                    edmshot.hangle = Val(Left$(v$, 3) + "." + Mid$(v$, 4, 4))
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

A = InStr(preset$, "+")
B = InStr(preset$, "-")
If B < A And B <> 0 Then A = B
preset$ = " L" + Mid$(preset$, A)

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

A = InStr(vhdata$, "<")
vhdata$ = Mid$(vhdata$, A)

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

Cancelling = False
Screen.MousePointer = 11
mdiMain.StatusBar.Panels(6).Visible = True
mdiMain.StatusBar.Panels(7).Visible = True
Shooting = True

Select Case UCase(EDMName$)
Case "NIKON"
    d$ = "$MSR" 'take a measurement in precise mode and output data
    Call edmoutput(d$, errorcode)
    Call edminput(A$)   'I think this waits for ACK
    Call edminput(A$)   'and now wait for the data
    returndata$ = A$
    Call clearcom
    Shooting = False

Case "TOPCON"
    DoEvents
    If Cancelling Then GoTo ExitSub

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
    DoEvents
    If Cancelling Then GoTo ExitSub

    Call edminput(A$)

'    frmTakeshot.progress.Value = 25
'    frmTakeshot.Refresh
    If Cancelling Then
        GoTo ExitSub
    End If
    
    Call delay(0.5)

    '------------------------------------------------
    ' Take the actual measurement
    '------------------------------------------------
    d$ = "C"
    Call edmoutput(d$, errorcode)
    Call edminput(A$)
    
    Call delay(0.5)
    

'    frmTakeshot.progress.Value = 50
'    frmTakeshot.Refresh
    
    Call edminput(A$)
    mdiMain.StatusBar.Panels(6).Picture = mdiMain.Picture2(1).Picture
    mdiMain.StatusBar.Panels(7).Visible = False
    
    If A$ <> "CANCEL" And A$ <> "" Then
        'CALL delay(.1)
        Cancelling = False
        Call directoutput(ack$)
        Do Until Asc(Left$(A$, 1)) > 32 Or Len(A$) = 1
            A$ = Mid$(A$, 2)
        Loop
        returndata$ = A$
    Else
        Cancelling = True
        'GoTo ExitSub
    End If

    Call delay(0.1)
    mdiMain.StatusBar.Panels(6).Picture = mdiMain.Picture2(2).Picture
'    frmTakeshot.progress.Value = 75
'    frmTakeshot.Refresh
    Call clearcom
    Call horizontal(errorcode)
    Call clearcom
'    Unload frmTakeshot
    Shooting = False

Case "WILD", "LEICA", "BUILDER"
    Call clearcom
    returndata$ = ""
    d$ = "GET/M/WI11/WI21/WI22/WI31/WI51"
    Call edmoutput(d$, errorcode)
    Call edminput(A$)
    If A$ <> "CANCEL" Then returndata$ = A$
    Call clearcom
    Shooting = False
    mdiMain.StatusBar.Panels(6).Picture = mdiMain.Picture2(2).Picture

Case "WILD2"

    '%R1Q,2008:Command[long],Mode[long]     Request a measurement
    Call clearcom
    d$ = "%R1Q,2008:1,1"                    ' Use the defaults with 1 and 1
    Call edmoutput(d$, errorcode)
    Call edminput(A$)
    If A$ <> "CANCEL" Then
        '%R1Q,2108:WaitTime[long],Mode[long]
        'Waittime is in ms
        'Mode 1 = automatic
        Call clearcom
        d$ = "%R1Q,2108:10000,1"            ' Wait for the measurements and return the angles + sloped
        Call edmoutput(d$, errorcode)
        Call edminput(A$)
    End If
    returndata$ = A$
    
    Call clearcom
    d$ = "%R1Q,2008:3,1"                    ' Empty the measurement buffer as a precaution
    Call edmoutput(d$, errorcode)
    Call edminput(A$)
    
    Call clearcom
    Shooting = False
    mdiMain.StatusBar.Panels(6).Picture = mdiMain.Picture2(2).Picture
    
Case "SOKKIA"
    returndata$ = ""
    d$ = Chr$(17)
    Call edmoutput(d$, errorcode)
    Call edminput(A$)
    returndata$ = A$
    If Cancelling Or A$ = "CANCEL" Then
        d$ = Chr$(18)
        Call edmoutput(d$, errorcode)
        Cancelling = True
        GoTo ExitSub
    End If
    Call clearcom
    Shooting = False

Case "PENTOGRAPH"
    d$ = "GETPOS" 'take a measurement in precise mode and output data
    Call edmoutput(d$, errorcode)
    Call clearcom
    Shooting = False

Case Else
End Select

'mdiMain.StatusBar.Panels(6).Visible = False
mdiMain.StatusBar.Panels(7).Visible = False
Screen.MousePointer = 1
Shooting = False
Exit Sub
'frmMain.PBar.Visible = False

ExitSub:
    mdiMain.StatusBar.Panels(6).Visible = False
    mdiMain.StatusBar.Panels(7).Visible = False
    Call horizontal(errorcode)
    Call delay(1)
    Call clearcom
    Cancelling = True
    Screen.MousePointer = 1
    Shooting = False
    Exit Sub

End Sub

Sub sethortangle(angle$, deg, min, sec)

Dim A As String

If angle$ = "" Then
    da$ = Right("000" + LTrim$(Str$(deg)), 3)
    ma$ = Right("00" + LTrim$(Str$(min)), 2)
    sa$ = Right("00" + LTrim$(Str$(sec)), 2)
    If UCase(EDMName) = "SOKKIA" Then
        angle$ = da$ + "." + ma$ + sa$
    Else
        angle$ = da$ + ma$ + sa$
    End If
    'angle$ = LTrim$(Str$(deg)) + LTrim$(Str$(min)) + LTrim$(Str$(sec))

Else
    If InStr(angle$, ".") = 0 Then angle$ = angle$ + ".0000"
    A = InStr(angle$, ".")
    angle$ = Left$(angle$ + "0000", A + 4)
    If UCase(EDMName) = "SOKKIA" Then
        angle$ = Right("000" + angle$, 8)
        If Val(Left(angle$, 3)) >= 360 Then
            angle$ = Right("000" + Trim$(Str$(Val(Left(angle$, 3)) Mod 360)), 3) + Mid$(angle$, 4)
        End If
    ElseIf UCase(EDMName) = "NIKON" Then
        angle$ = Right("000" + angle$, 8)
        A = InStr(angle$, ".")
        angle$ = Left$(angle$, A - 1) + Mid$(angle$, A + 1)
    ElseIf UCase(EDMName) = "WILD2" Then
        angle$ = Right("000" + angle$, 8)
        A = InStr(angle$, ".")
        angle$ = Left$(angle$, A - 1) + Mid$(angle$, A + 1)
    Else
        angle$ = Left$(angle$, A - 1) + Mid$(angle$, A + 1)
    End If
End If

Select Case UCase(EDMName)
Case "NIKON"
    d$ = "!HAN" + Trim(angle$)
    MsgBox "About to output '" + d$ + "' - note this information for shannon.", vbOKOnly
    Call edmoutput(d$, errorcode)
    Call delay(1)
    Call clearcom
Case "TOPCON"
    d$ = "J+" + LTrim$(angle$) + "d"
    Call edmoutput(d$, errorcode)
    Call edminput(A$)
    Call delay(1)
    Call clearcom
Case "WILD", "LEICA", "BUILDER"
    d$ = "PUT/21...4+" + Right$("000" + LTrim$(angle$) + "0 ", 9)
    Call edmoutput(d$, errorcode)
    Call edminput(A$)
Case "WILD2"
    angledecdeg = Val(Left(angle$, 3)) + Val(Mid$(angle$, 4, 2)) / 60 + Val(Mid$(angle$, 6, 2)) / 3600
    anglerad = angledecdeg / 360 * (2 * 3.14159265359)
    ' %R1Q,2113:HzOrientation[double]
    d$ = "%R1Q,2113:" + Format(anglerad, "#0.0000000")
    Call edmoutput(d$, errorcode)
    Call edminput(A$)
Case "SOKKIA"
    d$ = "/Dc " + angle$ + Chr(13) + Chr(10)
    'd$ = "Gd" + Chr(13) + Chr(10) gets angle
    Call edmoutput(d$, errorcode)
Case "PENTOGRAPH"
    d$ = "GETPOS" + Chr(13) + Chr(10)
    Call edmoutput(d$, errorcode)
Case Else
End Select

End Sub

Public Sub takeshot_core(prismstatus)

Cancelling = False
Call takeshot_nostation(prismstatus)
If Cancelling Then Exit Sub
edmshot.X = CurrentStation.X + edmshot.X
edmshot.y = CurrentStation.y + edmshot.y
edmshot.z = CurrentStation.z + edmshot.z - edmshot.poleh

End Sub

Sub setpresetdata(X As Single, y As Single, z As Single, errormessage$)

If displayinfo.edmtype <> 1 Then Exit Sub

px$ = LTrim$(Str$(X))
If X >= 0 Then px$ = "+" + px$
A = InStr(px$, ".")
If A <> 0 Then
    px$ = Left$(px$, A - 1) + Left$(Mid$(px$, A + 1), 3)
Else
    px$ = px$ + "000"
End If

py$ = LTrim$(Str$(y))
If X >= 0 Then py$ = "+" + py$
A = InStr(py$, ".")
If A <> 0 Then
    py$ = Left$(py$, A - 1) + Left$(Mid$(py$, A + 1), 3)
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
    A = InStr(pz$, ".")
    If A <> 0 Then
        pz$ = Left$(pz$, A - 1) + Left$(Mid$(pz$, A + 1), 3)
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

Dim A As Integer
Randomize
hash = ""
For A = 1 To hashlen
    hash = hash + Chr(Rnd * 25 + Asc("A"))
Next A

End Function

Sub parsecfg(needtoconvert)

'--------------------------------------------------
' Open the config file passed from main module
' Note that we already know it does in fact exist.
'--------------------------------------------------

Dim lineno As Integer
Dim fl As String
Dim ts As String
Dim A As Integer
Dim B As Integer
Dim C As Integer
Dim d As Integer
Dim key As String
Dim dt As String
Dim TempString1 As String
Dim temp(100) As String
Dim edmce As Boolean
Dim answer As Integer

Dim errorcode As Integer
Dim errormessage As String

Dim Inidata(100, 2) As String
Dim IniClass As String
Dim Status As Byte
Dim TempName As String
Dim TempX As String
Dim TempY As String
Dim TempZ As String

frmMain.ClearDBfields
On Error Resume Next
For I = 1 To 6
    frmMain.Button(I).Visible = False
    nButtonVars(I) = 0
Next I
On Error GoTo 0
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
    frmMain.lblPoleWarning.Visible = True
    Set PoleTB = Nothing
    Set DatumTB = Nothing
    Set SiteDB = Nothing
    PointTableName = ""
    SiteDBname = ""
    DBName = ""
    DBPath = ""
    CFGName = ""
    mdiMain.StatusBar.Panels(3) = ""
    frmMain.txtXYZ(0).Enabled = False
    frmMain.txtXYZ(1).Enabled = False
    frmMain.txtXYZ(2).Enabled = False
    frmMain.txtUnit.Enabled = False
    frmMain.txtID.Enabled = False
    frmMain.txtprism.Enabled = False
    
    mdiMain.StatusBar.Panels(4) = ""
    frmMain.txtPT = ""
    frmMain.txtprism.Clear
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
mdiMain.StatusBar.Panels(3) = "CFG: " + LCase(TempString) + "  "
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
                    A = InStr(fl, ":")
                    If A <> 0 Then
    
                            '--------------------------------------------
                            ' Set key to the keyword then select case on
                            ' the different types of keyword.
                            '--------------------------------------------
                            key = UCase(Trim(Left(fl, A - 1)))
                            dt = Trim(Mid(fl, A + 1))
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
                                    options(2) = "COM1"
                                    options(3) = UCase(dt)
    
                            Case "COM2"
                                    options(2) = "COM2"
                                    options(3) = UCase(dt)
    
                            Case "EDM"
                                    ts = UCase(dt)
                                    d = InStr(ts, " ")
                                    If ts <> "NONE" Then
                                            comport = Trim(Mid(ts, d))
                                            options(4) = UCase(LTrim(Left(ts, d - 1)))
                                    Else
                                            options(4) = "NONE"
                                    End If
                                    Select Case options(4)
                                    Case "GTS-3B", "GTS-301", "GTS-302", "GTS-303", "GTS-3X", "TOPCON"
                                            options(4) = "TOPCON" + " " + comport
                                    Case "NONE"
                                    Case "MANUAL"
                                    Case "TC-500", "WILD"
                                            options(4) = "WILD" + " " + comport
                                    Case "WILD2"
                                            options(4) = "WILD2" + " " + comport
                                    Case "SOKKIA"
                                            options(4) = "SOKKIA" + " " + comport
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
                                    options(11) = UCase(dt)
                                    Select Case options(11)
                                    Case "HP"
                                        options(11) = "HP"
                                    Case "HP-95", "HP-95LX", "HP95"
                                        options(11) = "HP95"
                                    Case "HP-100", "HP100"
                                        options(11) = "HP100"
                                    Case "HP-200", "HP-200LX", "HP200LX", "HP200"
                                        options(11) = "HP200"
                                    Case "HUSKY"
                                        options(11) = "HUSKY"
                                    Case "PC", "PC80", "2000"
                                    Case Else
                                    End Select
    
                            Case "SQID"
                                    '-------------------------------------
                                    'Is this a Laq/CC/Cagny type setup
                                    '-------------------------------------
                                    options(12) = Trim(dt)
    
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
                                    options(19) = UCase(dt)
    
                            Case "TAMISRATE"
    
                            Case "VHDORXYZ"
                                    options(21) = UCase(dt)
    
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
     'frmConvertCFG.Show 1
    If errorcode = 0 Then
        MsgBox ("The configuration file should be converted to Windows EDM format.  Use the File|Save CFG as ... menu option.")
    End If
Else
    For I = Len(CFGName) To 1 Step -1
        If Mid(CFGName, I, 1) = "\" Then
            CFGTitle = Mid(CFGName, I + 1)
            CFGpath = Left(CFGName, I)
            Exit For
        End If
    Next I
    IniClass = "[EDM]"
    Inidata(1, 1) = ""
    Inidata(2, 1) = "Database"
    Inidata(3, 1) = "PointTable"
    Inidata(4, 1) = "Instrument"
    Inidata(6, 1) = "COMport"
    Inidata(7, 1) = "SQID"
    Inidata(8, 1) = "Unitfields"
    Inidata(9, 1) = "Limitchecking"
    Inidata(10, 1) = "PrismPrompt"
    Inidata(11, 1) = "CurrentStation"
    Inidata(12, 1) = "StationX"
    Inidata(13, 1) = "stationY"
    Inidata(14, 1) = "stationZ"
    Inidata(15, 1) = "Settings"
    Inidata(16, 1) = "UpdateAlerts"
    Inidata(17, 1) = "UpperCase"
    Inidata(18, 1) = "DBPath"
    Inidata(19, 1) = "PPPath"
    Inidata(20, 1) = "ReferenceDatum"
    Inidata(21, 1) = "ReferenceDatum2"
    Inidata(22, 1) = "SetupType"
    
    Call ReadIni(CFGName, IniClass, Inidata(), Status)
    
    SiteDBname = Trim(Inidata(18, 2)) + Trim(Inidata(2, 2))
    DBPath = Trim(Inidata(18, 2))
    PPPath = Trim(Inidata(19, 2))
    If PPPath = "" Then
        PPPath = "My Documents\"
    End If
    DBName = Trim(Inidata(2, 2))
    PointTableName = Trim(Inidata(3, 2))
    If LCase(Inidata(7, 2)) = "y" Then SqidCheck = True Else SqidCheck = False
    UnitFieldString = Inidata(8, 2)
    ParseUnitFields
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
    If LCase(Inidata(17, 2)) = "yes" Then
        mdiMain.mnuUpperCase.Checked = True
        UpperCase = True
    Else
        mdiMain.mnuUpperCase.Checked = False
        UpperCase = False
    End If
    
    EDMName = ""
    If Inidata(4, 2) <> "" Then
        Select Case UCase(Inidata(4, 2))
        Case "TOPCON", "SOKKIA", "WILD", "WILD2", "NONE", "SIMULATE", "NIKON", "MICROSCRIBE"
            options(4) = Inidata(4, 2)
            EDMName = Inidata(4, 2)
        Case Else
            errorcode = 21
            errormessage = "Unrecognized EDM type.  Recognized types are Topcon, Wild, Wild2, Sokkia and None."
        End Select
    Else
        options(4) = "NONE"
    End If
    If LCase(Inidata(16, 2)) = "yes" Then
        NoAlert = False
        mdiMain.mnuNoAlert.Checked = False
    Else
        NoAlert = True
        mdiMain.mnuNoAlert.Checked = True
    End If
    
    options(3) = Inidata(5, 2)
    options(2) = Inidata(6, 2)
    comport = Inidata(6, 2)
    comsettings = Inidata(15, 2)
    If (EDMName = "" And LCase(EDMName) <> "simulate") And (comport = "" Or comsettings = "") Then
        frmMain.lblEDMWarning.Visible = True
    Else
        frmMain.lblEDMWarning.Visible = False
    End If
    options(17) = Inidata(8, 2)
        
    If options(17) <> "" Then
        B = InStr(options(17), ",")
        If B = 0 Then
            UnitName = Trim(options(17))
    Else
            UnitName = Trim(Left(options(17), B - 1))
        End If
    Else
        UnitName = "unit"
    End If
    TempName = Inidata(11, 2)
    TempX = Inidata(12, 2)
    TempY = Inidata(13, 2)
    TempZ = Inidata(14, 2)
    RefDatum1 = Inidata(20, 2)
    RefDatum2 = Inidata(21, 2)
    StationName = TempName
    SetupType = Val(Inidata(22, 2))
    
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
                    Gotit = False
                    For I = 1 To Vars
                        If LCase(VarList(I)) = LCase(ts) Then Gotit = True
                    Next I
                    If Gotit Then
                        MsgBox ("Duplicate field names in " + CFGName + ".")
                    Else
                        Vars = Vars + 1
                        VarList(Vars) = ts
                    End If
                End Select
            End If
        Loop Until EOF(1)
    Loop
    Close 1
    
    For A = 1 To 50
        Inidata(A, 1) = ""
        Inidata(A, 2) = ""
    Next A
    Inidata(1, 1) = "Type"
    Inidata(2, 1) = "Prompt"
    Inidata(3, 1) = "Menu"
    Inidata(4, 1) = "Length"
    Inidata(5, 1) = "Increment"
    Inidata(6, 1) = "Carry"
    Inidata(7, 1) = "Unique"
    DatumInfo = False
    For C = 1 To Vars
        VCarry(C) = False
        VIncr(C) = False
        VUnique(C) = False
        For A = 1 To 7
            Inidata(A, 2) = ""
        Next A
        IniClass = VarList(C)
        Call ReadIni(CFGName, IniClass, Inidata, Status)
        
        'strip the [ ]
        VarList(C) = Mid(VarList(C), 2)
        VarList(C) = Left(VarList(C), Len(VarList(C)) - 1)
        If LCase(VarList(C)) = "datumname" Then
            DatumInfo = True
            VType(C) = "TEXT"
        Else
            VType(C) = UCase(Inidata(1, 2))
        End If
        VPrompt(C) = Inidata(2, 2)
        VMenu(C) = Inidata(3, 2)
        VLen(C) = Val(Inidata(4, 2))
        If LCase(VType(C)) = "menu" Then
            TempString1 = ""
            MenuString = VMenu(C)
            Gotit = False
            Do Until Gotit
                X = InStr(MenuString, ",")
                If X > 0 Then
                    TempString = Left(Trim(Left(MenuString, X - 1)), VLen(C))
                    TempString1 = TempString1 + TempString + ","
                    MenuString = Trim(Mid(MenuString, X + 1))
                Else
                    TempString = Left(Trim(MenuString), VLen(C))
                    TempString1 = TempString1 + TempString
                    VMenu(C) = TempString1
                    Gotit = True
                End If
            Loop

        End If
        If UCase(VarList(C)) = "ID" Then IDLength = VLen(C)
        If UCase(VarList(C)) = "UNIT" Then UnitLength = VLen(C)
    
        If LCase(Inidata(5, 2)) = "true" Or LCase(Inidata(5, 2)) = "yes" Then VIncr(C) = True
    
        If LCase(Inidata(6, 2)) = "true" Or LCase(Inidata(6, 2)) = "yes" Then VCarry(C) = True
    
        If LCase(Inidata(7, 2)) = "true" Or LCase(Inidata(7, 2)) = "yes" Then VUnique(C) = True
    
    Next C
    For I = 1 To 6
        For J = 1 To Vars + 2
            Inidata(J, 1) = ""
            Inidata(J, 2) = ""
        Next J
        IniClass = "[BUTTON" + Trim(Str(I)) + "]"
        For J = 1 To Vars
            Inidata(J, 1) = VarList(J)
        Next J
        Inidata(Vars + 1, 1) = "Title"
        Inidata(Vars + 2, 1) = "Shortcut"
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
            ButtonCaption(I) = Inidata(Vars + 1, 2)
            ButtonShortCut(I) = Inidata(Vars + 2, 2)
            frmMain.Button(I).Visible = True
            If Trim(ButtonShortCut(I)) = "" Then
                frmMain.Button(I).Caption = "&" + Trim(Str(I)) + " " + ButtonCaption(I)
            Else
                frmMain.Button(I).Caption = ButtonCaption(I) + " (" + ButtonShortCut(I) + ")"
            End If
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
    If TempString <> "" Then Call OpenSite(SiteDBname$)
    If TempString <> "" And SiteDBname$ <> "" Then
        If PointTableName <> "" Then OpenPointsTable
        ParseUnitFields
        frmMain.lblDBWarning.Visible = False
    Else
        MsgBox ("Database given in " + CFGName + " not found or not in the correct format.")
        SiteDBname = ""
        DBPath = ""
        DBName = ""
        frmMain.lblDBWarning.Visible = True
    End If
End If
addtofilelist CFGName

If Not StationInitialized Then
    If TempName <> "" And TempX <> "" And TempY <> "" And TempZ <> "" Then
        TempString = "Continue with station " + TempName + "?" + Chr(13) + Chr(13)
        TempString = TempString + "           X: " + Format(Val(TempX), "#####0.000") + Chr(13)
        TempString = TempString + "           Y: " + Format(Val(TempY), "#####0.000") + Chr(13)
        TempString = TempString + "           Z: " + Format(Val(TempZ), "#####0.000") + Chr(13)
        
        response = MsgBox(TempString, vbYesNo)
        If response = vbYes Then
            CurrentStation.Name = TempName
            CurrentStation.X = Val(TempX)
            CurrentStation.y = Val(TempY)
            CurrentStation.z = Val(TempZ)
            frmMain.lblStationWarning.Visible = False
            StationInitialized = True
            mdiMain.StatusBar.Panels(5) = "Current Station: " + CurrentStation.Name + "  "
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
PCcfgPath = CFGName
frmMain.lblCFGWarning.Visible = False

'inifile$ = fixpath(App.Path) + "edm.ini"
'Call WriteEDMIni(inifile$)
End Sub

Sub parseangle(hangle As Single, angle As Integer, minutes As Integer, seconds As Integer)

Dim B As Integer

A$ = Str$(hangle)
B = InStr(A$, ".")

If B = 0 Then
        angle = Val(A$)
        minutes = 0
        seconds = 0
Else
        A$ = A$ + "0000"
        If Val(Left$(A$, B - 1)) < 32000 Then angle = Val(Left$(A$, B - 1)) Else errorcode = -1
        If Val(Mid$(A$, B + 1, 2)) < 32000 Then minutes = Val(Mid$(A$, B + 1, 2)) Else errorcode = -1
        If Val(Mid$(A$, B + 3, 2)) < 32000 Then seconds = Val(Mid$(A$, B + 3, 2)) Else errorcode = -1
End If

End Sub

Public Sub OpenPointsTable()

Dim Xlength As Integer
Dim Gotit As Boolean

frmMain.lblBlankFields.Visible = False
frmMain.txtCurrentRecord = 0
frmMain.txtTotalRecords = 0
frmMain.PointsADO.RecordSource = ""
If Not tablematch(PointTableName) Then
    PointTableName = ""
    frmMain.lblPointsWarning.Visible = True
    frmMain.txtXYZ(0).Enabled = False
    frmMain.txtXYZ(1).Enabled = False
    frmMain.txtXYZ(2).Enabled = False
    frmMain.txtUnit.Enabled = False
    frmMain.txtID.Enabled = False
    frmMain.txtprism.Enabled = False
    Exit Sub
End If

For Each A In SiteDB.TableDefs(PointTableName).Fields
    Gotit = False
    If LCase(A.Name) <> "recno" Then
        For I = 1 To Vars
            If UCase(A.Name) = UCase(VarList(I)) Then
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
            'FixFieldSize I
        End If
    End If
Next I

frmMain.txtPT = PointTableName
For I = 1 To Vars
    Select Case UCase(VType(I))
        Case "MENU", "TEXT"
            If SiteDB.TableDefs(PointTableName).Fields(VarList(I)).Size <> VLen(I) Then
                'FixFieldSize I
            End If
        Case "NUMERIC", "INSTRUMENT"
    End Select
Next I

frmMain.PointsADO.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + SiteDBname + ";Persist Security Info=False"
frmMain.PointsADO.RecordSource = PointTableName
frmMain.PointsADO.Refresh
frmMain.PointsADO.Recordset.Requery
frmMain.txtXYZ(0).Enabled = True
frmMain.txtXYZ(1).Enabled = True
frmMain.txtXYZ(2).Enabled = True
frmMain.txtUnit.Enabled = True
frmMain.txtID.Enabled = True
frmMain.txtprism.Enabled = True

Dim Inidata(1, 2) As String
Dim IniClass As String
Dim Status As Byte
'IniClass = "[EDM]"
'Inidata(1, 1) = "ShowGrid"
'Call ReadIni(CFGName, IniClass, Inidata(), Status)
'If UCase(Inidata(1, 2)) = "YES" Then
'    mnuDataGrid.Caption = "Show Data Grid"
'    mdiMain.mnuDataGrid_Click
'End If

If GridShowing Then
    frmDataGrid.Form_Load
End If
frmMain.lblPointsWarning.Visible = False
Loading = True

If Not frmMain.PointsADO.Recordset.BOF And Not frmMain.PointsADO.Recordset.EOF Then
    frmMain.PointsADO.Recordset.MoveFirst
    frmMain.PointsADO.Recordset.MoveLast
    CurrentBookMark = frmMain.PointsADO.Recordset.Bookmark
    frmMain.ShowValues
Else
    ClearFields
    CurrentBookMark = Empty
End If
Loading = False
frmMain.lblPointsWarning.Visible = False

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
frmMain.txtXYZ(0).Enabled = False
frmMain.txtXYZ(1).Enabled = False
frmMain.txtXYZ(2).Enabled = False
frmMain.txtUnit.Enabled = False
frmMain.txtID.Enabled = False
frmMain.txtprism.Enabled = False

Cancelling = True

End Sub

Public Sub ParseUnitFields()

Dim X As Integer

nUnitFields = 0

TempString = Trim(UnitFieldString)
X = InStr(TempString, ",")
Do While X > 0
    nUnitFields = nUnitFields + 1
    Unitfield(nUnitFields) = Left(TempString, X - 1)
    TempString = Trim(Mid(TempString, X + 1))
    X = InStr(TempString, ",")
Loop
If Trim(TempString) <> "" Then
    nUnitFields = nUnitFields + 1
    Unitfield(nUnitFields) = Trim(TempString)
End If
If SiteDBOpen = False Then Exit Sub



On Error Resume Next
For I = 1 To nUnitFields
    Gotit = False
    Gotit = UCase(Unitfield(I)) = UCase(SiteDB.TableDefs("EDM_units").Fields(Unitfield(I)).Name)
    If Not Gotit Then GoTo BadTable:
Next I
For Each A In SiteDB.TableDefs("EDM_units").Fields
    Select Case UCase(A.Name)
        Case "MINX", "MINY", "MAXX", "MAXY", "CENTERX", "CENTERY", "RADIUS"
        Case Else
            Gotit = False
            For I = 1 To Vars
                If UCase(A.Name) = UCase(VarList(I)) Then
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
For I = 0 To frmMain.PointsADO.Recordset.Fields.Count - 1
    If LCase(frmMain.PointsADO.Recordset.Fields(I).Name) = LCase(Varname) Then
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
Dim A As Boolean, B As Boolean, C As Boolean, d As Boolean, E As Boolean, F As Boolean, G As Boolean, H As Boolean, M As Boolean, n As Boolean, o As Boolean, P As Boolean
Dim NPtable As Boolean
nPlotFiles = 0
nPointTables = 0
nOverlayTables = 0
nWorkBooks = 0
nTempWorkBooks = 0
nLookupTables = 0

SiteDB.TableDefs.Refresh
For Each tdtemp In SiteDB.TableDefs
    A = False
    If Left(tdtemp.Name, 7) <> "WinPlot" And Left(tdtemp.Name, 4) <> "MSys" And Left(tdtemp.Name, 2) <> "$$" And Left(tdtemp.Name, 4) <> "EDM_" Then
        On Error Resume Next
        A = SiteDB.TableDefs(tdtemp.Name).Fields("RecNo").Size > 0
        If A Then
            nPointTables = nPointTables + 1
            PointTable(nPointTables) = LCase(tdtemp.Name)
        End If
    End If
Next
On Error GoTo 0
Exit Sub

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
For Each A In SiteDB.TableDefs("EDM_units").Fields
    Select Case UCase(A.Name)
        Case "UNIT", "ID", "MINX", "MINY", "MAXX", "MAXY", "CENTERX", "CENTERY", "RADIUS"
            Gotit = True
        Case Else
            Gotit = False
            For I = 1 To Vars
                If UCase(A.Name) = UCase(VarList(I)) Then
                    Gotit = True
                    Exit For
                End If
            Next I
            If Not Gotit Then
                SiteDB.TableDefs("EDM_units").Fields.Delete A.Name
                GoTo KeepLooking
            End If
    End Select
Next

Set tdf = Nothing
SiteDB.TableDefs.Refresh
nUnitFields = 0
For Each A In SiteDB.TableDefs("EDM_Units").Fields
    Select Case UCase(A.Name)
        Case "MINX", "MINY", "MAXX", "MAXY", "CENTERX", "CENTERY", "RADIUS"
        Case Else
            nUnitFields = nUnitFields + 1
            Unitfield(nUnitFields) = A.Name
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
Set F = Nothing

End Sub

Public Sub FixFieldSize(VarNum)

On Error GoTo errorhandler

frmMain.PointsADO.RecordSource = ""
Set tdf = SiteDB.TableDefs(PointTableName)

Set F = tdf.CreateField("*temp", dbText, VLen(VarNum))
tdf.Fields.Append F
SqlString = "update [" + PointTableName + "] set [*temp]=" + VarList(VarNum)
SiteDB.Execute SqlString
SiteDB.TableDefs(PointTableName).Fields.Delete VarList(VarNum)
SiteDB.TableDefs(PointTableName).Fields("*temp").Name = VarList(VarNum)
Set tdf = Nothing
Set F = Nothing
frmMain.PointsADO.RecordSource = PointTableName
frmMain.PointsADO.Refresh
frmMain.PointsADO.Recordset.Requery
On Error GoTo 0
Set tdf = Nothing
Set F = Nothing
Exit Sub

errorhandler:
    MsgBox (Err.Description)
    frmMain.PointsADO.RecordSource = PointTableName
    frmMain.PointsADO.Refresh
    frmMain.PointsADO.Recordset.Requery
    

End Sub

Sub conv_angle_to_degminsec(angle As Double, degrees As Integer, minutes As Integer, seconds As Integer)

degrees = Int(angle)
seconds = Int((angle - CDbl(degrees)) * 3600#)
minutes = Int(seconds / 60)
seconds = seconds Mod 60

End Sub

Sub takeshot_nostation(prismstatus)

Dim edmpoffset As Single

If Speaking Then Voice.Speak "shoe" + "shoe" + "shoe" + "  " + "Shooting", SVSFlagsAsync
mdiMain.StatusBar.Panels(6).Visible = True
mdiMain.StatusBar.Panels(6).Picture = mdiMain.Picture2(0).Picture

'Do
'    DoEvents
'Loop Until Cancelling = True

mdiMain.StatusBar.Panels(6).Visible = True

Select Case UCase(EDMName)
Case "MICROSCRIBE"
    frmGetMSData.Show 1

Case "SIMULATE"
    w = 0
    Select Case w
    
    Case 0
        Randomize
        X = 2 * Rnd + 1018
        y = 1 * Rnd + 1006
        ''X = 1017 + Rnd
        'y = 1017 + Rnd
        z = Rnd
        edmshot.X = X
        edmshot.y = y
        edmshot.z = z
        edmshot.hangle = 111.505
        edmshot.vangle = 98.2525
        edmshot.sloped = Sqr(edmshot.X ^ 2 + edmshot.y ^ 2 + edmshot.z ^ 2)
        If prismstatus = AskForPrism Then
            frmSelectPrism.Show 1
        End If
    
    Case 1
        returndata$ = "?+00002413m0893100+0000000d+00002413t60+00+25107"
        If prismstatus = AskForPrism Then
            frmSelectPrism.Show 1
        End If
        errorcode = 0
        Call parsenez(returndata$, edmshot, edmpoffset, mesunits$, angleunit$, errorcode)
        If errorcode = 0 Then
            Call vhdtonez(edmshot)
        Else
            Select Case errorcode
            Case 1, 5
                MsgBox "ERROR: Shot could not be recorded.  Verify that prism is visible.", vbCritical
            Case 6
                MsgBox "ERROR: Parity error.  Shot must be retaken.", vbCritical
            Case 2
                MsgBox "ERORR: Angles not in degrees.  Reset instrument.", vbCritical
            Case Else
                MsgBox "ERROR: Unknown error.  Shot must be retaken.  Returned data was '" + returndata$ + "'", vbCritical
            End Select
            Exit Sub
        End If
    Case 2
        returndata$ = "?+00002260m0900735+0212020d+00002260t60+00+25099"
        If prismstatus = AskForPrism Then
            frmSelectPrism.Show 1
        End If
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
    Case Else
    End Select
    
Case Else
    Call recordpoint(returndata$)
    If Cancelling Then
        If Speaking Then Voice.Speak "Shot cancelled", SVSFlagsAsync
        Exit Sub
    End If
    
    If prismstatus = AskForPrism Then
        frmSelectPrism.Show 1
    End If
    
    If EDMName$ <> "PENTOGRAPH" Then
        Call parsenez(returndata$, edmshot, edmpoffset, mesunits$, angleunit$, errorcode)
    Else
        edmshot.vangle = 0
        edmshot.hangle = 0
        edmshot.sloped = 0
        edmshot.edmpoffset = 0
        edmshot.poleh = 0
    End If
    
    If edmshot.vangle > 180 Then
        MsgBox ("ERROR: The horizontal plane (of the vertical angle) is not set to 90, or the barrel of the theodolite has been reversed")
        Cancelling = True
        Exit Sub
    End If
    
    If UCase(EDMName$) = "LEICA" Or UCase(EDMName$) = "WILD" Or UCase(EDMName$) = "WILD2" Then
        If edmshot.edmpoffset = 0.004 Then
            frmMain.LblReflectorless.Visible = True
        Else
            frmMain.LblReflectorless.Visible = False
        End If
    ElseIf UCase(EDMName$) = "BUILDER" Then
        frmMain.LblReflectorless.Visible = True
    End If
    
    If errorcode = 0 Then
        If EDMName$ <> "PENTOGRAPH" Then
            Call vhdtonez(edmshot)
        Else
            errorcode = 200
            If Len(returndata$) > 14 Then
                If Left(returndata$, 4) = "POSX" Then
                    returndata$ = Mid$(returndata$, 5)
                    t = InStr(returndata$, ";")
                    If t <> 0 Then
                        edmshot.X = Val(Left(returndata$, t - 1))
                        returndata$ = Mid$(returndata$, t + 1)
                        If Left(returndata$, 4) = "POSY" Then
                            returndata$ = Mid$(returndata$, 5)
                            t = InStr(returndata$, ";")
                            If t <> 0 Then
                                edmshot.y = Val(Left(returndata$, t - 1))
                                returndata$ = Mid$(returndata$, t + 1)
                                If Left(returndata$, 4) = "POSZ" Then
                                    returndata$ = Mid$(returndata$, 5)
                                    edmshot.y = Val(returndata$)
                                    errorcode = 0
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If errorcode <> 0 Then
        If Speaking Then Voice.Speak "Shot error", SVSFlagsAsync
        Cancelling = True
        Select Case errorcode
        Case 1, 5
            MsgBox "ERROR: Shot could not be recorded.  Verify that prism is visible or if in reflectorless mode that the surface can reflect.", vbCritical
        Case 6
            MsgBox "ERROR: Parity error.  Shot must be retaken.", vbCritical
        Case 2
            MsgBox "ERORR: Angles not in degrees.  Reset instrument.", vbCritical
        Case 100, 101, 102, 103
            MsgBox "Tell Shannon the error number was " + Str(errorcode)
        Case Else
            MsgBox "ERROR: Unknown error.  Shot must be retaken.  Returned data was '" + returndata$ + "'", vbCritical
        End Select
        Cancelling = True
        Exit Sub
    End If
End Select

' mdiMain.StatusBar.Panels(6).Visible = False
mdiMain.StatusBar.Panels(7).Visible = False
If Speaking Then Voice.Speak "Got it", SVSFlagsAsync

End Sub

Public Function CountRecords()

CountRecords = frmMain.PointsADO.Recordset.RecordCount

End Function

Public Sub ClearFields()

frmMain.txtXYZ(0) = ""
frmMain.txtXYZ(1) = ""
frmMain.txtXYZ(2) = ""
frmMain.txtUnit = ""
frmMain.txtID = ""
frmMain.txtSuffix = ""
On Error Resume Next
For I = 1 To 50
    frmMain.MenuBox(I) = ""
    frmMain.TextBox(I) = ""
    frmMain.NumberBox(I) = ""
Next I
OriginalUnit = ""
OriginalID = ""
frmMain.txtCurrentRecord = 0
frmMain.txtTotalRecords = 0

On Error GoTo 0

End Sub
