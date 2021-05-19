Attribute VB_Name = "Module1"
'   f) need a browse option in file open to change dirs

Declare Function WriteFileL Lib "Coredll" Alias "WriteFile" (ByVal hFile As Long, lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Declare Function WriteFile Lib "Coredll" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

Public rs As Recordset      'This is a recordset object shared by all
                            'Once created, it must stay open otherwise memory leaks
Public rs2 As Recordset     'For those times when a second recordset is needed

Public rsprisms As Recordset
Public rsdatums As Recordset
Public rsunits As Recordset
Public rspoints As Recordset
Public tablesopen As Boolean

Public fpath As String
Public cfgfile As String
Public edmini As String

Public button_values(6, 100) As String

Public limits(100, 4) As Double
Public limitname(100) As String
Public unitlimits As Integer
Public use_limitchecking As Boolean
Public unitname As String

Public datum(100, 4) As Double
Public datumname(100) As String
Public ndatums As Integer

Public npoles As Integer
Public poledata(100, 2) As Single
Public polename(100)

'These variables handle the results of the last shot
Public currentstationx As Double
Public currentstationy As Double
Public currentstationz As Double
Public currentstationname As String
Public referencedatum As String

Public firstxp As Double
Public firstyp As Double

Public currentprismheight As Single
Public currentprismoffset As Single
Public currentprismname As String
Public currentxp As Double
Public currentyp As Double
Public currentzp As Double
Public currentsloped As Double
Public currentvangle As String
Public currenthangle As String
Public shottype As Integer      '0=no shot, >0 record no. of shot being edited, -1=new shot, -2=x-shot
Public button_no As Integer     'for user defined buttons
Public in_emulation As Boolean  '
Public back_to_edit As Boolean
Public back_to_main As Boolean
Public shot_canceled As Boolean
Public currentunit As String
Public refresh_point_grid As Boolean
Public last_id As String
Public last_unit As String

Public editvar As Integer       'which variable is being edited at any one moment
Public actualvar As Integer     'which variable in the list vs. editvar which is the control
Public responses(100) As String 'used to store the values for a point while they are edited

Public vars As Integer

Public vtype(100) As Integer
Public vardefault(100) As String
Public varprompt(100) As String
Public varlist(100) As String
Public vlen(100) As Integer
Public vmenu(100) As String

Public vunique(100) As Boolean
Public vcarry(100) As Boolean
Public vincr(100) As Boolean
Public varprint(100) As Boolean

Public options(100) As String
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

Sub parsecfg(needtoconvert)

'--------------------------------------------------
' Open the config file passed from main module
' Note that we already know it does in fact exist.
'--------------------------------------------------

Dim lineno As Integer
Dim fl As String
Dim ts As String
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim key As String
Dim dt As String
Dim comport As String
Dim temp(100) As String
Dim edmce As Boolean
Dim answer As Integer
Dim tempunitname As String

Dim errorcode As Integer
Dim errormessage As String

Dim inidata(100, 2)
Dim iniclass As String
Dim status As Integer

npoles = 0
vars = 0
ndatums = 0
needtoconvert = 0

'first, check to see if the file is in old EDM format
'or new EDMCE/EDM format by looking for that section in the file
edmce = False
frmFileOpen.File1.Open cfgfile, fsModeInput
Do While Not frmFileOpen.File1.EOF
    ts = frmFileOpen.File1.LineInputString
    If InStr(UCase(ts), "[EDM]") <> 0 Then
        edmce = True
        Exit Do
    End If
Loop
frmFileOpen.File1.Close

If Not edmce Then
    lineno = 0
    frmFileOpen.File1.Open cfgfile, fsModeInput
    Do While Not frmFileOpen.File1.EOF
            '--------------------------------------------------
            ' Look for a continuation character at the end of
            ' the line.  If it exists add the next line to a.
            ' In the end, a contains a full line from the CFG.
            '--------------------------------------------------
            fl = ""
            Do
                    lineno = lineno + 1
                    ts = frmFileOpen.File1.LineInputString
                    ts = Trim(ts)
                    fl = fl + ts
                    If Right(ts, 1) = "\" Then Mid(fl, Len(fl), 1) = " "
            Loop While Right(ts, 1) = "\" And Not frmFileOpen.File1.EOF
    
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
    
                                    If dt = "" Then
                                            errorcode = 10
                                            errormessage = "No field name given. "
                                            Exit Do
                                    End If
    
                                    '-------------------------------------
                                    ' Check against the array of all field
                                    ' names for duplicates.
                                    '-------------------------------------
                                    For b = 1 To vars
                                            If varlist(b) = dt Then
                                                    errorcode = 1
                                                    errormessage = "Duplicate FIELD definition for " + RTrim(Mid(a, a + 1))
                                                    Exit Do
                                            End If
                                    Next b
    
                                    '-------------------------------------
                                    ' Check the upper limit for number of
                                    ' possible variables in the program.
                                    '-------------------------------------
                                    If vars = UBound(varlist, 1) Then
                                            errorcode = 12
                                            errormessage = "Too many FIELD definitions.  Limit is" + Str(UBound(varlist, 1))
                                            Exit Do
                                    End If
    
                                    '-------------------------------------
                                    ' If a new field name add this to the
                                    ' field name list.
                                    '-------------------------------------
                                    vars = vars + 1
                                    varlist(vars) = UCase(dt)
    
                            Case "POINTSFILE"
                                    '-------------------------------------
                                    'Set option(1) to the output filename
                                    '-------------------------------------
                                    options(1) = UCase(dt)
    
                            Case "COM1"
                                    '-------------------------------------
                                    'Set option(2) to the com parameters
                                    '-------------------------------------
                                    options(2) = "COM1"
                                    options(3) = UCase(dt)
    
                            Case "COM2"
                                    '-------------------------------------
                                    'Set option(3) to the com parameters
                                    '-------------------------------------
                                    options(2) = "COM2"
                                    options(3) = UCase(dt)
    
                            Case "EDM"
                                    '-------------------------------------
                                    'Set option(4) to the type of instrument
                                    'and check that it is a valid type.
                                    '-------------------------------------
                                    ts = UCase(dt)
                                    d = InStr(ts, " ")
                                    If d = 0 And ts <> "NONE" Then
                                            errorcode = 11
                                            errormessage = "Missing COM port assignment."
                                            Exit Do
                                    ElseIf ts <> "NONE" Then
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
                                    Case "SOKKIA"
                                            options(4) = "SOKKIA" + " " + comport
                                    Case Else
                                            errorcode = 11
                                            errormessage = "Expect TOPCON, WILD or SOKKIA instrument."
                                            Exit Do
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
                                            errorcode = 11
                                            errormessage = "Expect HP, HP100, HP200, PC or PC80 computer type."
                                            Exit Do
                                    End Select
    
                            Case "SQID"
                                    '-------------------------------------
                                    'Is this a Laq/CC/Cagny type setup
                                    '-------------------------------------
                                    options(12) = Trim(dt)
    
                            Case "PRISMDEF"
                                    If npoles = UBound(poledata, 1) Then
                                            errorcode = 93
                                            errormessage = "The number of poles is limited to" + Str(UBound(poledata, 2))
                                            Exit Do
                                    End If
                                    ts = dt
                                    d = InStr(ts, " ")
                                    If d = 0 Then
                                            errorcode = 90
                                            errormessage = "Missing prism name or height."
                                            Exit Do
                                    End If
                                    npoles = npoles + 1
                                    polename(npoles) = UCase(Trim(Left(ts, d - 1)))
                                    ts = LTrim(Mid(ts, d))
                                    d = InStr(ts, " ")
                                    If d = 0 Then
                                            errorcode = 91
                                            errormessage = "Missing prism offset."
                                            npoles = npoles - 1
                                            Exit Do
                                    End If
                                    poledata(npoles, 1) = UCase(Trim(Left(ts, d - 1)))
                                    poledata(npoles, 2) = UCase(Trim(Mid(ts, d + 1)))
    
                            Case "DATUMDEF"
                                    If ndatums = UBound(datum, 1) Then
                                            errorcode = 93
                                            errormessage = "The number of datums is limited to" + Str(UBound(datum, 1))
                                            Exit Do
                                    End If
                                    ts = dt
                                    f = 0
                                    Do Until Len(ts) = 0
                                            e = InStr(ts, " ")
                                            If e = 0 Then e = Len(ts) + 1
                                            f = f + 1
                                            temp(f) = Left(ts, e - 1)
                                            ts = LTrim(Mid(ts, e))
                                    Loop
                                    If f <> 5 Then
                                            errorcode = 94
                                            errormessage = "Format DATUMDEF as Name X Y Z Height"
                                            Exit Do
                                    Else
                                            ndatums = ndatums + 1
                                            datum(ndatums, 1) = temp(2)
                                            datum(ndatums, 2) = temp(3)
                                            datum(ndatums, 3) = temp(4)
                                            datum(ndatums, 4) = temp(5)
                                            datumname(ndatums) = temp(1)
                                                                                    
                                            'NOT SURE WHAT THIS CODE DOES
                                            'For e = 2 To f - 4
                                            '        datumname(ndatums) = Trim(datumname(ndatums)) + " " + temp(e)
                                            'Next e
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
                                options(17) = ""
                                dt = UCase(dt)
                                Do Until Len(dt) = 0
                                    e = InStr(dt, " ")
                                    If e = 0 Then e = Len(dt) + 1
                                    If options(17) = "" Then
                                        options(17) = Left(dt, e - 1)
                                    Else
                                        options(17) = options(17) + "," + Left(dt, e - 1)
                                    End If
                                    dt = LTrim(Mid(dt, e))
                                Loop
    
                                If options(17) <> "" Then
                                    b = InStr(options(17), ",")
                                    If b = 0 Then
                                        unitname = Trim(options(17))
                                    Else
                                        unitname = Trim(leftstr(options(17), b - 1))
                                    End If
                                Else
                                    unitname = "unit"
                                End If

                            '------------------------------------------
                            'Set limits for units
                            '------------------------------------------
                            Case "LIMIT"
                                    dt = UCase(dt)
                                    f = 0
                                    Do Until Len(dt) = 0
                                            e = InStr(dt, " ")
                                            If e = 0 Then e = Len(dt) + 1
                                            f = f + 1
                                            temp(f) = Left(dt, e - 1)
                                            dt = LTrim(Mid(dt, e))
                                    Loop
                                    tempunitname = temp(1)
                                    For a = 1 To unitlimits
                                            If Trim(limitname(a)) = tempunitname Then
                                                    errorcode = 95
                                                    errormessage = tempunitname + " defined twice."
                                                    Exit Do
                                            End If
                                    Next a
                                    Select Case temp(2)
                                    Case "RECT"
                                            If f <> 6 Then
                                                    errorcode = 94
                                                    errormessage = "Format RECT as Name RECT X1 Y1 X2 Y2"
                                                    Exit Do
                                            End If
                                            unitlimits = unitlimits + 1
                                            limitname(unitlimits) = tempunitname
                                            limits(unitlimits, 1) = CDbl(temp(3))
                                            limits(unitlimits, 2) = CDbl(temp(4))
                                            limits(unitlimits, 3) = CDbl(temp(5))
                                            limits(unitlimits, 4) = CDbl(temp(6))
    
                                    Case Else
                                            errormessage = "Missing unit type."
                                            errorcode = 81
                                            Exit Do
    
                                    End Select
    
                            '------------------------------------------
                            ' Optionally put the sitename here.
                            '------------------------------------------
                            Case "SITE"
                                    options(18) = dt
    
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
                                    For c = 1 To vars
    
                                            'check if defined variable
    
                                            If varlist(c) = key Then
                                                    b = InStr(dt, " ") ' look for space
                                                    If b = 0 Then b = Len(dt) + 1 'no value list
    
                                                    '-------------------------------------
                                                    'variable name and data
                                                    '-------------------------------------
                                                    Var = UCase(LTrim(RTrim(Left(dt, b - 1))))
                                                    dat = LTrim(RTrim(Mid(dt, b + 1)))
    
                                                    '-------------------------------------
                                                     ' check for possible commands
                                                    '-------------------------------------
                                                    Select Case Var
                                                    Case "INPUT"
                                                            Select Case UCase(dat)
                                                            Case "TEXT"
                                                                    vtype(c) = 1
                                                            Case "NUMERIC"
                                                                    vtype(c) = 2
                                                            Case "INSTRUMENT"
                                                                    vtype(c) = 3
                                                            Case "MENU"
                                                                    vtype(c) = 4
                                                            Case "UNIT"
                                                                    vtype(c) = 5
    
                                                            '---------------------
                                                            'For CC system only.
                                                            '---------------------
                                                            Case "SQID"
                                                                    vtype(c) = 6
                                                                    vlen(c) = 11
                                                                    If options(10) = "" Then
                                                                            errorcode = 7
                                                                            errormessage = "SQIDFILE option must be specified first."
                                                                            Exit Do
                                                                    End If
    
                                                            Case Else
                                                                    errorcode = 7
                                                                    errormessage = Left(dat, 20) + " is an unrecognized INPUT type. "
                                                                    Exit Do
                                                            End Select
    
                                                    Case "DEFAULT"
                                                            vardefault(c) = dat
    
                                                    Case "PRINT"
                                                            dat = UCase(Left(dat, 1))
                                                            If dat = "Y" Then varprint(c) = dat
    
                                                    Case "PROMPT"
                                                            varprompt(c) = dat
                                                            If Len(varprompt(c)) > 60 Then varprompt(c) = Left(varprompt, 60)
    
                                                    Case "MENULIST"
                                                            vmenu(c) = ""
                                                            menuitems = 0
                                                            Do Until Len(dat) = 0
                                                                    e = InStr(dat, " ")
                                                                    If e = 0 Then e = Len(dat) + 1
                                                                    If vmenu(c) = "" Then
                                                                        vmenu(c) = UCase(Left(dat, e - 1))
                                                                    Else
                                                                        vmenu(c) = vmenu(c) + "," + UCase(Left(dat, e - 1))
                                                                    End If
                                                                    dat = LTrim(Mid(dat, e))
                                                                    menuitems = menuitems + 1
                                                            Loop
    
                                                    Case "VARLEN"
                                                            vlen(c) = dat
    
                                                    Case "CARRY"
                                                            vcarry(c) = 1
    
                                                    Case "INCREMENT"
                                                            vincr(c) = 1
    
                                                    Case "VARLOC"
                                                    
                                                    Case Else
                                                            errorcode = 3
                                                            errormessage = "Unrecognized command " + Var + " "
                                                            Exit Do
                                                    End Select
    
                                                    'If varprompt(c) = "" Then
                                                    '        If vlen(c) + Len(varlist(c)) + 2 + varloc(c, 2) > display.Width Then
                                                    '                errorcode = 90
                                                    '                errormessage = "Field name + length is too long."
                                                    '                Exit Do
                                                    '        End If
    
                                                    'ElseIf Len(varprompt(c)) + vlen(c) + varloc(c, 2) > display.Width Then
                                                     '       errorcode = 90
                                                    '        errormessage = "Field name + prompt is too long."
                                                    '        Exit Do
    
                                                    'End If
                                                    ERRORFLAG = 0
                                                    Exit For
                                            End If
                                    Next c
                                    If ERRORFLAG = 1 Then
                                            errorcode = 2
                                            errormessage = "No FIELD statement or unrecognized command " + key + ". "
                                            Exit Do
                                    End If
                            End Select
                    Else
                            errorcode = 5
                            errormessage = "Unrecognized statement or field."
                            Exit Do
                    End If
            End If
    Loop
    frmFileOpen.File1.Close
    
    'For a = 1 To vars
    '    If varlist(a) = "DATE" Then
    '        varlist(a) = "DAY"
    '        Exit For
    '    End If
    'Next a
    
    If errorcode = 0 Then
        Screen.MousePointer = 1
        answer = MsgBox("This file will be converted to EDMCE format.", vbOKCancel, "EDMCE")
        If answer = 1 Then
            needtoconvert = 1
            Screen.MousePointer = 11
        Else
            errorcode = 100
            errormessage = "The configuration file must be converted to EDMCE format for this program to work."
        End If
    End If

Else
    
    iniclass = "[EDM]"
    inidata(1, 1) = "Sitename"
    inidata(2, 1) = "Database"
    inidata(3, 1) = "PointTable"
    'inidata(4, 1) = "Instrument"
    'inidata(5, 1) = "COMport"
    'inidata(6, 1) = "COMparameters"
    inidata(7, 1) = "SQID"
    inidata(8, 1) = "Unitfields"
    inidata(9, 1) = "Limitchecking"
    
    Call ReadIni(cfgfile, iniclass, inidata, status)
    
    If inidata(1, 2) <> "" Then
        options(18) = inidata(1, 2)
    Else
        options(18) = "EDM"
    End If
    
    If inidata(2, 2) <> "" Then
        options(1) = inidata(2, 2)
    Else
        options(1) = "EDM"
    End If
    
    If inidata(3, 2) <> "" Then
        options(22) = inidata(3, 2)
    Else
        options(22) = "EDM"
    End If
    
    'If inidata(4, 2) <> "" Then
    '    Select Case UCase(inidata(4, 2))
    '    Case "TOPCON", "SOKKIA", "WILD", "NONE"
    '        options(4) = inidata(4, 2)
    '    Case Else
    '        errorcode = 21
    '        errormessage = "Unrecognized EDM type.  Recognized types are Topcon, Wild, Sokkia and None."
    '    End Select
    'Else
    '    options(4) = "NONE"
    'End If
    
    'inidata(5, 1) = "COMport"
    'options(3) = inidata(5, 2)
    'inidata(6, 1) = "COMparameters"
    'options(2) = inidata(6, 2)

    If LCase(inidata(7, 2)) = "yes" Then options(12) = "Yes"
    
    If LCase(inidata(9, 2)) = "yes" Then use_limitchecking = True Else use_limitchecking = False
    
    options(17) = inidata(8, 2)
        
    If options(17) <> "" Then
        b = InStr(options(17), ",")
        If b = 0 Then
            unitname = Trim(options(17))
        Else
            unitname = Trim(leftstr(options(17), b - 1))
        End If
    Else
        unitname = "unit"
    End If
    
    'now need to load variables
    'First, need to read through CFG looking for the field names
    
    frmFileOpen.File1.Open cfgfile, fsModeInput
    vars = 0
    Do While Not frmFileOpen.File1.EOF
        Do
            ts = frmFileOpen.File1.LineInputString
            ts = Trim(ts)
            If leftstr(ts, 1) = "[" Then
                ts = UCase(ts)
                Select Case ts
                Case "[EDM]", "[BUTTON1]", "[BUTTON2]", "[BUTTON3]", "[BUTTON4]", "[BUTTON5]", "[BUTTON6]"
                Case Else
                    vars = vars + 1
                    varlist(vars) = ts
                End Select
            End If
        Loop Until frmFileOpen.File1.EOF
    Loop
    frmFileOpen.File1.Close
    
    For a = 1 To 100
        inidata(a, 1) = ""
        inidata(a, 2) = ""
    Next a
    inidata(1, 1) = "Type"
    inidata(2, 1) = "Prompt"
    inidata(3, 1) = "Menu"
    inidata(4, 1) = "Length"
    inidata(5, 1) = "Increment"
    inidata(6, 1) = "Carry"
    inidata(7, 1) = "Unique"
    
    For c = 1 To vars
        
        For a = 1 To 7
            inidata(a, 2) = ""
        Next a
        iniclass = varlist(c)
        Call ReadIni(cfgfile, iniclass, inidata, status)
        
        'strip the [ ]
        varlist(c) = Mid(varlist(c), 2)
        varlist(c) = leftstr(varlist(c), Len(varlist(c)) - 1)
        
        Select Case LCase(inidata(1, 2))
        Case "text"
            vtype(c) = 1
        Case "numeric"
            vtype(c) = 2
        Case "instrument"
            vtype(c) = 3
        Case "menu"
            vtype(c) = 4
        Case "unit"
            vtype(c) = 5
        Case Else
        End Select
    
        varprompt(c) = inidata(2, 2)
        vmenu(c) = inidata(3, 2)
    
        If inidata(4, 2) <> "" Then vlen(c) = CInt(inidata(4, 2))
    
        If LCase(inidata(5, 2)) = "true" Or LCase(inidata(5, 2)) = "yes" Then vincr(c) = True
    
        If LCase(inidata(6, 2)) = "true" Or LCase(inidata(6, 2)) = "yes" Then vcarry(c) = True
    
        If LCase(inidata(7, 2)) = "true" Or LCase(inidata(7, 2)) = "yes" Then vunique(c) = True
    
    Next c
    
End If

End Sub

Sub write_datafile_ini(filename, status As Integer)

'This routine writes an ini for the data table within a site

Dim inidata(100, 2)
Dim iniclass
Dim a As Integer
Dim b As Integer
Dim wstatus

If LCase(Right(filename, 4)) <> ".cfg" Then
    filename = filename + ".cfg"
End If

'If it exists, kill it and start over
f = frmMain.FileSystem1.Dir(filename)
If f <> "" Then
    frmMain.FileSystem1.Kill filename
End If

'Write general ini variables
iniclass = "[EDM]"
inidata(1, 1) = "Sitename"
inidata(1, 2) = options(18)
inidata(2, 1) = "Database"
inidata(2, 2) = options(1)
inidata(3, 1) = "PointTable"
inidata(3, 2) = options(22)
inidata(4, 1) = "SQID"
inidata(4, 2) = options(12)
inidata(5, 1) = "Unitfields"
inidata(5, 2) = options(17)
inidata(6, 1) = "Limitchecking"
If use_limitchecking = True Then inidata(6, 2) = "Yes" Else inidata(6, 2) = "No"

'Write general info for this config
Call WriteIni(filename, iniclass, inidata, wstatus)

'Clear the arrays
For b = 1 To UBound(inidata, 1)
    inidata(b, 1) = ""
    inidata(b, 2) = ""
Next b

'Write out the buttons
If frmMain.button1.Visible = True Then Call save_button(1, frmMain.button1.Caption, filename)
If frmMain.button2.Visible = True Then Call save_button(2, frmMain.button2.Caption, filename)
If frmMain.button3.Visible = True Then Call save_button(3, frmMain.button3.Caption, filename)
If frmMain.button4.Visible = True Then Call save_button(4, frmMain.button4.Caption, filename)
If frmMain.button5.Visible = True Then Call save_button(5, frmMain.button5.Caption, filename)
If frmMain.button6.Visible = True Then Call save_button(6, frmMain.button6.Caption, filename)

'Now write section for each field
For a = 1 To vars
    Call savevar(a, filename)
Next a

End Sub
Sub WriteIni(inifile, iniclass, inidata, status)

Dim Gotit As Boolean
Dim f As String
Dim a1 As String
Dim a2 As String
Dim iclass As Integer
Dim c As Integer
Dim b As Integer
Dim varname As String
Dim vardata As String
Dim flag As Integer

f = frmMain.FileSystem1.Dir(inifile)
If f = "" Then
    frmMain.File1.Open inifile, fsModeOutput
    frmMain.File1.LinePrint iniclass
    frmMain.File1.Close
End If

Do
    frmMain.File1.Open inifile, fsModeInput
    frmMain.File2.Open "$temp$.ini", fsModeOutput
    
    Do Until frmMain.File1.EOF
        
        a1 = frmMain.File1.LineInputString
        a2 = UCase(Trim(a1))
        
        If Left(a2, 1) = "[" Then
            frmMain.File2.LinePrint a1
            If a2 = UCase(iniclass) Then
                iclass = 1
                Gotit = True
                For c = 1 To UBound(inidata, 1)
                    If Trim(inidata(c, 1)) <> "" And Trim(inidata(c, 2)) <> "" Then
                        frmMain.File2.LinePrint inidata(c, 1) + "=" + CStr(inidata(c, 2))
                    End If
                Next c
            Else
                iclass = 0
            End If
            
        ElseIf iclass = 1 Then
            b = InStr(a2, "=")
            If b <> 0 Then
                varname = UCase(Left(a2, b - 1))
                vardata = Mid(a2, b + 1)
                flag = 0
                For c = 1 To UBound(inidata, 1)
                    If UCase(inidata(c, 1)) = varname Then
                        flag = 1
                        Exit For
                    End If
                Next c
                If flag = 0 Then
                    frmMain.File2.LinePrint a1
                End If
            Else
                frmMain.File2.LinePrint a1
            End If
        Else
            frmMain.File2.LinePrint a1
        End If
    Loop
    
    frmMain.File1.Close
    frmMain.File2.Close
    
    'If the section was never found, add it and start over
    If Not Gotit Then
        frmMain.File1.Open inifile, fsModeAppend
        frmMain.File1.LinePrint " "
        frmMain.File1.LinePrint iniclass
        frmMain.File1.Close
    End If

Loop Until Gotit = True

'Delete the existing one, rename the new one as the existing one
frmMain.FileSystem1.Kill inifile
frmMain.FileSystem1.MoveFile "$temp$.ini", inifile

End Sub

Sub ReadIni(inifile, iniclass, inidata, status)

Dim f As String
Dim iclass As Integer
Dim b As Integer
Dim c As Integer
Dim varname As String
Dim vardata As String
Dim a1 As String
Dim a2 As String

f = frmMain.FileSystem1.Dir(inifile)
If f = "" Then
    frmMain.File1.Open inifile, fsModeOutput
    frmMain.File1.LinePrint iniclass
    frmMain.File1.Close
End If

frmMain.File1.Open inifile, fsModeInput

Do Until frmMain.File1.EOF
    a1 = frmMain.File1.LineInputString
    a1 = Trim(a1)
    a2 = UCase(a1)
    If Left(a2, 1) = "[" Then
        If a2 = UCase(iniclass) Then
            iclass = 1
        Else
            iclass = 0
        End If
    ElseIf iclass = 1 Then
        b = InStr(a2, "=")
        If b <> 0 Then
            varname = Left(a2, b - 1)
            vardata = Mid(a1, b + 1)
            For c = 1 To UBound(inidata, 1)
                If UCase(inidata(c, 1)) = UCase(varname) Then
                    If inidata(c, 2) = "" Then
                        inidata(c, 2) = vardata
                    Else
                        inidata(c, 2) = inidata(c, 2) + Chr(1) + vardata
                    End If
                    Exit For
                End If
            Next c
        End If
    End If
Loop

frmMain.File1.Close

End Sub

Sub create_db(dbname)

rs.Open "CREATE DATABASE '" + dbname + "'"

End Sub

Sub add_field_tb()

Dim a As Integer
Dim sql(100)
Dim sqlcmd

sql(0) = "CREATE TABLE " + options(22) + " (RECNO int, "
For a = 1 To vars
    If UCase(varlist(a)) = "DATE" Or UCase(varlist(a)) = "DAY" Then
        sql(a) = varlist(a) + " datetime "       'adDouble
    ElseIf UCase(varlist(a)) = "TIME" Then
        sql(a) = varlist(a) + " datetime "       'adDouble
    Else
        Select Case vtype(a)
        Case 1
            sql(a) = varlist(a) + " varchar(" + CStr(vlen(a)) + ") " 'adVarWChar
        Case 2
            sql(a) = varlist(a) + " float "       'adDouble
        Case 3
            sql(a) = varlist(a) + " float "       'adDouble
        Case 4
            sql(a) = varlist(a) + " varchar(" + CStr(vlen(a)) + ") "  'adVarWChar
        Case 5
            sql(a) = varlist(a) + " varchar(" + CStr(vlen(a)) + ") "  'adVarWChar
        Case Else
        End Select
    End If
Next a

sqlcmd = sql(0)
For a = 1 To vars - 1
  sqlcmd = sqlcmd & sql(a) & ","
Next
sqlcmd = sqlcmd + sql(vars) + ")"

dbname = options(1)
rs.Open sqlcmd, dbname

sqlcmd = "CREATE INDEX Recno ON " & options(22) & " (Recno)"
rs.Open sqlcmd, options(1)

End Sub

Sub set_comm_menus()
    
Dim t1 As String
Dim l(10) As String
Dim n As Integer

Select Case UCase(options(2))
Case "COM1"
    frmCommunications.commport.Text = "COM1"
Case "COM2"
    frmCommunications.commport.Text = "COM2"
Case "COM3"
    frmCommunications.commport.Text = "COM3"
Case "COM4"
    frmCommunications.commport.Text = "COM4"
Case Else
End Select

t1 = options(3)
If t1 <> "" Then
    Call parse_list(t1, ",", n, l)
    If l(1) <> "" Then frmCommunications.baudrate.Text = l(1)
    If l(2) <> "" Then
        Select Case LCase(l(2))
        Case "n"
            frmCommunications.parity.Text = "None"
        Case "e"
            frmCommunications.parity.Text = "Even"
        Case "o"
            frmCommunications.parity.Text = "Odd"
        Case Else
        End Select
    End If
    If l(3) <> "" Then frmCommunications.databits.Text = l(3)
    If l(4) <> "" Then frmCommunications.stopbits.Text = l(4)
End If

End Sub

Sub parse_list(ltext As String, d As String, n As Integer, l() As String)

Dim t1 As String
Dim a As Integer

t1 = Trim(ltext)

For a = 1 To UBound(l, 1)
    l(a) = ""
Next a

n = 0
Do Until t1 = "" Or n = UBound(l, 1)
    a = InStr(t1, d)
    If a <> 0 Then
        If Left(t1, a - 1) <> "" Then
            n = n + 1
            l(n) = Left(t1, a - 1)
        End If
        t1 = Mid(t1, a + 1)
    ElseIf t1 <> "" Then
        n = n + 1
        l(n) = t1
        Exit Do
    End If
Loop

End Sub

Sub fix_cfg_info()

Dim a As Integer

If options(1) = "" Then
    options(1) = fpath + "edm.cdb"
End If

If options(22) = "" Then
    options(22) = "edm"
End If

If InStr(options(1), ".") <> 0 Then
    a = InStr(options(1), ".")
    If LCase(Mid(options(1), a + 1)) <> "cdb" Then
        options(1) = Left(options(1), a) + "cdb"
    End If
End If

If InStr(options(1), "\") = 0 Then
    options(1) = fpath + options(1)
End If

For a = 1 To vars
    Select Case varlist(a)
    Case "X", "Y", "Z", "VANGLE", "HANGLE", "SLOPED", "DATUMX", "DATUMY", "DATUMZ"
        vtype(a) = 2
        vlen(a) = 13
    Case "SUFFIX"
        vtype(a) = 2
        vlen(a) = 4
    Case "PRISM"
        vtype(a) = 2
        vlen(a) = 8
    Case "DATUMNAME"
        vtype(a) = 1
        vlen(a) = 20
    Case "DAY", "TIME", "DATE"
        vtype(a) = 1
        vlen(a) = 10
    Case Else
        If vtype(a) = 0 Then vtype(a) = 1
    End Select
Next a

For a = 1 To vars
    Select Case varlist(a)
    Case "VANGLE"
        If varprompt(a) = "" Then varprompt(a) = "Vertical angle"
    Case "HANGLE"
        If varprompt(a) = "" Then varprompt(a) = "Horizontal angle"
    Case "SLOPED"
        If varprompt(a) = "" Then varprompt(a) = "Slope distance"
    Case Else
        varprompt(a) = UCase(Left(varlist(a), 1)) + LCase(Mid(varlist(a), 2))
    End Select
Next a

End Sub

Sub table_status(tbname, status)

Dim rc, strList, r

status = -1

'First see if the table exists
rs.Open "MSysTables", options(1), adOpenKeyset, adLockOptimistic
rc = rs.RecordCount
For r = 0 To rc - 1
    If LCase(rs.Fields("TableName").Value) = LCase(tbname) Then
        status = 0
        Exit For
    End If
    rs.MoveNext
Next
rs.Close

'If it does, see how many records it has
If status <> -1 Then
    rs.Open tbname, options(1), adOpenKeyset, adLockOptimistic
    status = rs.RecordCount
    rs.Close
End If

End Sub

Sub add_datum_table()

Dim sqlcmd

sqlcmd = "CREATE TABLE datums (name varchar(20), x float, y float, z float, day varchar(20), time varchar(20) )"
rs.Open sqlcmd, options(1)

sqlcmd = "CREATE INDEX name ON datums (name)"
rs.Open sqlcmd, options(1)

End Sub

Sub add_prism_table()

Dim sqlcmd

sqlcmd = "CREATE TABLE prisms (name varchar(20), height float, offset float)"
rs.Open sqlcmd, options(1)

sqlcmd = "CREATE INDEX name ON prisms (name)"
rs.Open sqlcmd, options(1)

End Sub

Sub add_unit_table()

Dim sqlcmd
Dim l As String
Dim a As Integer
Dim b As Integer
Dim n As Integer
Dim v As String
Dim sql1 As String

If unitname <> "" Then
    sqlcmd = "CREATE TABLE units (" + unitname + " varchar(20), x_sw float, y_sw float,x_ne float, y_ne float"
Else
    sqlcmd = "CREATE TABLE units (unit varchar(20), x_sw float, y_sw float,x_ne float, y_ne float"
End If

sql1 = ""

'if there are unit fields then add them to the end of these default fields
If options(17) <> "" Then
    l = LCase(options(17))
    sql1 = ""
    Do Until l = ""
        b = InStr(l, ",")
        If b = 0 Then
            v = Trim(l)
            l = ""
        Else
            v = Trim(leftstr(l, b - 1))
            l = Mid(l, b + 1)
        End If
        If v <> "" Then
            n = n + 1
            If n <> 1 Then
                For a = 1 To vars
                    If LCase(varlist(a)) = v Then
                        Select Case vtype(a)
                        Case 1
                            If sql1 = "" Then
                                sql1 = varlist(a) + " varchar(" + CStr(vlen(a)) + ") " 'adVarWChar
                            Else
                                sql1 = sql1 + ", " + varlist(a) + " varchar(" + CStr(vlen(a)) + ") " 'adVarWChar
                            End If
                        Case 2
                            If sql1 = "" Then
                                sql1 = varlist(a) + " float "    'adDouble
                            Else
                                sql1 = sql1 + ", " + varlist(a) + " float "     'adDouble
                            End If
                        Case 3
                            If sql1 = "" Then
                                sql1 = varlist(a) + " float "       'adDouble
                            Else
                                sql1 = sql1 + ", " + varlist(a) + " float "   'adDouble
                            End If
                        Case 4
                            If sql1 = "" Then
                                sql1 = varlist(a) + " varchar(" + CStr(vlen(a)) + ") "  'adVarWChar
                            Else
                                sql1 = sql1 + ", " + varlist(a) + " varchar(" + CStr(vlen(a)) + ") " 'adVarWChar
                            End If
                        Case 5
                            If sql1 = "" Then
                                sql1 = varlist(a) + " varchar(" + CStr(vlen(a)) + ") "  'adVarWChar
                            Else
                                sql1 = sql1 + ", " + varlist(a) + " varchar(" + CStr(vlen(a)) + ") " 'adVarWChar
                            End If
                        Case Else
                        End Select
                        Exit For
                    End If
                Next a
            End If
        End If
    Loop
End If

If sql1 = "" Then
    sqlcmd = sqlcmd + ")"
Else
    sqlcmd = sqlcmd + "," + sql1 + ")"
End If

rs.Open sqlcmd, options(1)

sqlcmd = "CREATE INDEX " & unitname & " ON units (" & unitname & ")"
rs.Open sqlcmd, options(1)

End Sub

Function leftstr(t, l)

'This kludge is to work around the fact that CE doesn't accept Left on forms

leftstr = Left(t, l)

End Function
Sub use_defaults()

Dim crlf
Dim status As Integer

vars = 17
varlist(1) = "Sitename"
vtype(1) = 1
vlen(1) = 20
vcarry(1) = True
varlist(2) = "ID"
vtype(2) = 2
vincr(2) = True
vcarry(2) = True
varlist(3) = "Suffix"
vtype(3) = 2
vcarry(3) = True
varlist(4) = "Code"
vtype(4) = 4
vmenu(4) = "TOPO,ARTIFACT,DATUM"
vlen(4) = 10
vcarry(4) = True
varlist(5) = "X"
varlist(6) = "Y"
varlist(7) = "Z"
varlist(8) = "PRISM"
varlist(9) = "HANGLE"
varlist(10) = "VANGLE"
varlist(11) = "SLOPED"
varlist(12) = "DAY"
varlist(13) = "TIME"
varlist(14) = "DATUMNAME"
varlist(15) = "DATUMX"
varlist(16) = "DATUMY"
varlist(17) = "DATUMZ"

options(2) = "COM1"
options(3) = "1200,E,7,1"
options(4) = "SIMULATION"
options(11) = "CE"
options(12) = "no"
options(13) = "Stadia"
options(14) = "2"
options(15) = "0"
options(17) = "Sitename"
options(19) = "no"

Call fix_cfg_info

use_limitchecking = False

dbname = options(1)
unitname = "Sitename"

If InStr(dbname, "\") = 0 Then
    dbname = fpath + dbname
End If

If InStr(dbname, ".") = 0 Then
    dbname = dbname + ".cdb"
End If

options(1) = dbname

Call setup_dbs

cfgfile = fpath + options(22) + ".cfg"
Call write_datafile_ini(cfgfile, status)

End Sub

Sub calculate_angle(x1 As Double, y1 As Double, x2 As Double, y2 As Double, angle As Double)

Dim rrun As Double, rise As Double
Dim slope As Double, pi As Double
Dim tangle As Double

pi = 3.14159265359

rrun = CDbl(x2) - CDbl(x1)
rise = CDbl(y2) - CDbl(y1)

If rrun = 0 Then
    tangle = 90
Else
    slope = rise / rrun
    tangle = CSng(Atn(slope) * 180# / pi)
End If

tangle = 90 - tangle

If (rrun >= 0) And (rise >= 0) Then
        tangle = 0 + tangle
ElseIf (rrun > 0) And (rise < 0) Then
        tangle = 360 + tangle
ElseIf (rrun <= 0) And (rise >= 0) Then
        tangle = 180 + tangle
ElseIf (rrun <= 0) And (rise < 0) Then
        tangle = 180 + tangle
End If

'tangle = tangle + 90
If tangle >= 360 Then tangle = tangle - 360

angle = tangle

End Sub

Sub conv_angle_to_degminsec(angle As Double, degrees As Integer, minutes As Integer, seconds As Integer)

degrees = Int(angle)
seconds = Int((angle - CDbl(degrees)) * 3600#)
minutes = Int(seconds / 60)
seconds = seconds Mod 60

End Sub

Sub record_point(status As Integer)


Select Case options(4)
Case "TOPCON", "WILD", "SOKKIA"
    
Case "NONE"
    'need hand entry section
    
Case "SIMULATION"

Case Else
End Select

End Sub

Function degtorad(degrees As Single) As Single

degtorad = CSng(degrees * 1.74532925199433E-02)

End Function

Sub open_com()

If Not in_emulation Then
    If frmMain.Comm1.PortOpen Then frmMain.Comm1.PortOpen = False
End If

Select Case options(2)
Case "COM1"
    portno = 1
Case "COM2"
    portno = 2
Case "COM3"
    portno = 3
Case "COM4"
    portno = 4
Case "COM5"
    portno = 5
Case "COM6"
    portno = 6
Case "COM7"
    portno = 7
Case "COM8"
    portno = 8
Case Else
    portno = 1
End Select

If Not in_emulation Then
    frmMain.Comm1.commport = portno
    frmMain.Comm1.Settings = options(3)
    On Error Resume Next
    If Not frmMain.Comm1.PortOpen Then frmMain.Comm1.PortOpen = True
    If Err <> 0 Then
        MsgBox "Could not open " & options(2) & ":" & options(3) & " Make sure you are cabled and use the Station Comport Settings to change the settings.", vbInformation, "EDM CE"
        Exit Sub
    End If
End If

Call initedm
Call horizontal(errorcode)

End Sub

Sub close_com()

If Not in_emulation Then
    If frmMain.Comm1.PortOpen Then frmMain.Comm1.PortOpen = False
End If

End Sub

Sub edminput(a As String)
                             
Dim t As String
Dim timeone As Long

timeone = Timer

a = ""
t = ""
If Not in_emulation Then
    Do
        t = frmMain.Comm1.Input
        If t <> "" Then a = a + t
    Loop Until Right(a, 2) = Chr(13) + Chr(10) Or Timer - timeone > 15
End If

End Sub

Sub edmoutput(d As String, errorcode As Integer)

Dim term As String
Dim bcc As String
Dim a As String
Dim b As Integer
Dim bTemp(200) As Byte
Dim y As Integer

term = Chr(13) + Chr(10)
errorcode = 0

Select Case options(4)
Case "TOPCON"
        Call makebcc(d, bcc)
        a = d + bcc + Chr(3) + term
        
Case "WILD"
        a = d + term
        
Case "SOKKIA"
        a = d

Case Else
        Exit Sub

End Select

Select Case options(4)
Case "TOPCON"
    frmMain.Comm1.InBufferCount = 0
    frmMain.Comm1.OutBufferCount = 0

    'Move the string into a byte array
    For y = 1 To Len(a)
        bTemp(y) = Asc(Mid(a, y, 1))
    Next y

    'Call microsoft routine for comm output
    'SendArrayData frmMain.Comm1.CommID, bTemp, Len(a)

    SendData frmMain.Comm1.CommID, a

Case "WILD"
    frmMain.Comm1.Output = a
    ''Print #display.DataSource, a;

Case Else
End Select


End Sub

Sub horizontal(errorcode As Integer)

Dim d As String
Dim a As String

errorcode = 0
Select Case options(4)
Case "TOPCON"
    d = "Z10"
    Call edmoutput(d, errorcode)
    Call edminput(a)
    If a = "CANCEL" Then errorcode = 27
Case Else
End Select

End Sub

Sub initedm()

Dim errorcode
Dim d As String
Dim a As String

Select Case options(4)
Case "TOPCON"
        d = "ST0"
        Call edmoutput(d, errorcode)

Case "WILD"
        d = "SET/41/0"
        Call edmoutput(d, errorcode)
        If errorcode = 0 Then
                Call edminput(a)
        End If
        d = "SET/149/2"
        Call edmoutput(d, errorcode)
        If errorcode = 0 Then
                Call edminput(a)
        End If
        ''Call delay(0.5)
        ''Call clearcom

Case Else
End Select

End Sub

Sub makebcc(itext As String, otext As String)

Dim b As Integer
Dim i As Integer
Dim q As Integer
Dim b1 As Integer
Dim b2 As Integer

b = 0
For i = 1 To Len(itext)
        q = Asc(Mid(itext, i, 1))
        b1 = q And (Not b)
        b2 = b And (Not q)
        b = b1 Or b2
Next i

otext = LTrim(CStr(b))
otext = Right("000" + otext, 3)

End Sub

Sub parsenez(nezdata As String, edmpoffset As Single, mesunits As String, angleunit As String, errorcode As Integer)

Dim angle As Integer, minutes As Integer, seconds As Integer
Dim tangle As Single, dangle As Single, dist As Single
Dim a As String
Dim bcc1 As String
Dim bcc2 As String
Dim d As String

errorcode = 0
Select Case options(4)
Case "TOPCON"
    
    Do Until Asc(Left(nezdata, 1)) > 32 Or Len(nezdata) = 1
        nezdata = Mid(nezdata, 2)
    Loop
    a = Left(nezdata, 1)
    If a <> "?" And a <> "R" Then
        If a = "U" Then
            errorcode = 5
        Else
            errorcode = 1
        End If
        Exit Sub
    End If

    a = InStr(nezdata, Chr(3))
    If a <> 0 Then
        bcc1 = Mid(nezdata, a - 3, 3)
        d = Left(nezdata, a - 4)
        Call makebcc(d, bcc2)
    End If

    If bcc1 <> bcc2 Then
        errorcode = 6
    Else
        nezdata = LTrim(nezdata)
        currentsloped = FormatNumber(Mid(nezdata, 2, 9)) / 1000
        currenthangle = FormatNumber(Mid(nezdata, 19, 4) + "." + Mid(nezdata, 23, 4), 4)
        currentvangle = FormatNumber(Mid(nezdata, 12, 3) + "." + Mid(nezdata, 15, 4), 4)
        edmpoffset = FormatNumber(Mid(nezdata, 43, 3)) / 1000
        mesunits = Mid(nezdata, 11, 1)
        angleunit = Mid(nezdata, 27, 1)
        If angleunit <> "d" Then
            errorcode = 2
        End If
    End If

Case "WILD"
    a = InStr(nezdata, "21.")
    If a = 0 Then
        errorcode = 1
        Exit Sub
    End If
    nezdata = Mid(nezdata, a)

    'IF LEN(nezdata) <> 64 OR a = 0 THEN
    '        errorcode = 1
    '        EXIT SUB
    'END IF

    If Mid(nezdata, 6, 1) <> "4" Then
        errorcode = 2
        Exit Sub
    End If
    currenthangle = Val(Mid(nezdata, 7, 4) + "." + Mid(nezdata, 11, 4))

    If Mid(nezdata, 22, 1) <> "4" Then
        errorcode = 2
        Exit Sub
    End If
    currentvangle = Val(Mid(nezdata, 23, 4) + "." + Mid(nezdata, 27, 4))

    a = InStr(nezdata, "31..0")
    currentsloped = Val(Mid(nezdata, a + 6, 9)) / 1000
    If currentsloped = 0 Then
        errorcode = 4
        Exit Sub
    End If
    mesunits = Mid(nezdata, a + 5)
    edmpoffset = Val(Right(nezdata, 4)) / 1000

Case "SOKKIA"
    edmpoffset = 0
    nezdata = RTrim(LTrim(nezdata))
    a = InStr(nezdata, " ")
    If a > 1 Then
        currentsloped = Val(Left(nezdata, a - 1)) / 1000
        If currentsloped = 0 Then
            errorcode = 10
        Else
            a = Mid(nezdata, a + 1)
            a = InStr(a, " ")
            If a > 1 And Left(a, 1) <> "E" Then
                currentvangle = Val(Left(a, 3) + "." + Mid(a, 4, 4))
                a = Mid(a, a + 1)
                a = InStr(a, " ")
                If a > 1 And Left(a, 1) <> "E" Then
                    currenthangle = Val(Left(a, 3) + "." + Mid(a, 4, 4))
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

Sub take_shot()

Dim d As String
Dim errorcode As Integer
Dim a As String

Select Case options(4)
Case "TOPCON"
    
    bcc1 = ""
    bcc2 = ""

    Call clearcom
    
    '------------------------------------------------
    ' Take measurement in sloped and angle mode
    '------------------------------------------------
    d = "Z34"
    Call edmoutput(d, errorcode)
    Call edminput(a)

    'Call delay(0.5)

    '------------------------------------------------
    ' Take the actual measurement
    '------------------------------------------------
    d = "C"
    Call edmoutput(d, errorcode)
    Call edminput(a)

    Call delay(0.5)
    
Case "WILD"
    d = "GET/M/WI11/WI21/WI22/WI31/WI51"
    Call edmoutput(d, errorcode)

Case "SOKKIA"
    d = Chr(17)
    Call edmoutput(d, errorcode)

Case Else
End Select

End Sub

Sub sethortangle(angle As String, deg As Integer, min As Integer, sec As Single)

Dim a As Integer
Dim d As String
Dim errorcode As Integer

If angle = "" Then
        angle = LTrim(CStr(deg)) + LTrim(CStr(min)) + LTrim(CStr(sec))

ElseIf InStr(angle, ".") <> 0 Then
        a = InStr(angle, ".")
        angle = Left(angle + "0000", a + 4)
        angle = Left(angle, a - 1) + Mid(angle, a + 1)
End If

Select Case options(4)
Case "TOPCON"
        d = "J+" + LTrim(angle) + "d"
        Call edmoutput(d, errorcode)
        Call edminput(a)
        Call delay(1)
        Call clearcom

Case "WILD"
        d = "PUT/21...4+" + Right("000" + LTrim(angle) + "0 ", 9)
        Call edmoutput(d, errorcode)
        Call edminput(a)

Case Else
End Select

End Sub

Sub vhdtonez()

Dim angle As Integer
Dim minutes As Integer
Dim seconds As Integer
Dim tangle As Double
Dim actuald As Double

Call parseangle(currentvangle, angle, minutes, seconds)
tangle = angle + ((minutes * 60 + seconds) / 3600)

'Calculate Z relative to total station
currentzp = currentsloped * Cos(degtorad(tangle))

'Calculate X and Y relative to total station

'***** is this an error ?  should this be currentzp?
actuald = Sqr(currentsloped ^ 2 - currentzp ^ 2)
Call parseangle(currenthangle, angle, minutes, seconds)
tangle = angle + ((minutes * 60 + seconds) / 3600)
tangle = 450 - tangle

currentxp = Cos(degtorad(tangle)) * actuald
currentyp = Sin(degtorad(tangle)) * actuald

End Sub

Sub directoutput(a)

'frmMain.Comm1.Output = a
SendData frmMain.Comm1.CommID, a

End Sub

Sub delay(a As Single)

Dim timeone As Double

timeone = Timer

Do
Loop Until Timer - timeone > a

End Sub

Sub clearcom()

Dim a As String

a = frmMain.Comm1.Input

End Sub

Public Sub SendArrayData(ByVal hCommID As Long, baData, nobytes)

'Microsoft web site code to fix issues with normal comm port output

Dim i, lRet, iWrite

For i = 1 To nobytes
    lRet = WriteFileL(hCommID, baData(i), 1, iWrite, 0)
Next

End Sub
    
Sub center(formname As Form)

formname.Left = Screen.Width / 2 - formname.Width / 2
formname.Top = Screen.Height / 2 - formname.Height / 2

End Sub

Sub write_edmce_ini(filename, status As Integer)

'This routine write an ini for the program itself

Dim inidata(100, 2)
Dim iniclass
Dim a As Integer
Dim b As Integer
Dim wstatus

'Write general ini variables
iniclass = "[EDM_CE]"
inidata(1, 1) = "CFG_File"
inidata(1, 2) = cfgfile
inidata(2, 1) = "Total Station"
inidata(2, 2) = options(4)
inidata(3, 1) = "COM"
inidata(3, 2) = options(2)
inidata(4, 1) = "Settings"
inidata(4, 2) = options(3)
inidata(5, 1) = "CurrentStation"
inidata(5, 2) = currentstationname
inidata(6, 1) = "CurrentStationX"
inidata(6, 2) = currentstationx
inidata(7, 1) = "CurrentStationY"
inidata(7, 2) = currentstationy
inidata(8, 1) = "CurrentStationZ"
inidata(8, 2) = currentstationz
inidata(9, 1) = "Referencedatum"
inidata(9, 2) = referencedatum

Call WriteIni(filename, iniclass, inidata, status)
For b = 1 To UBound(inidata, 1)
    inidata(b, 1) = ""
    inidata(b, 2) = ""
Next b

End Sub

Sub setup_dbs()

Dim f As String
Dim tbname As String
Dim dbname As String
Dim crlf As String
Dim status As Integer

dbname = options(1)
f = frmMain.FileSystem1.Dir(dbname)
If f = "" Then
    Call create_db(dbname)
End If

tbname = options(22)
Call table_status(tbname, status)
If status = -1 Then Call add_field_tb

crlf = Chr(13)

Call table_status("datums", status)
If status = -1 Then Call add_datum_table

Call table_status("prisms", status)
If status = -1 Then Call add_prism_table

Call table_status("units", status)
If status = -1 Then Call add_unit_table

End Sub

Function hash(hashlen)

Dim a As Integer

hash = ""
For a = 1 To hashlen
    hash = hash + Chr(Rnd * 25 + Asc("A"))
Next a

End Function

Public Sub SendData(ByVal hCommID As Long, sData As String)
    
Dim lRet, i, iWrite
For i = 1 To Len(sData)
    If Not in_emulation Then lRet = WriteFile(hCommID, ChrB(Asc(Mid(sData, i, 1))), 1, iWrite, 0)
Next
    
End Sub

Sub parseangle(hangle As Single, angle As Integer, minutes As Integer, seconds As Integer)

Dim a As String
Dim b As String
Dim c As Integer

a = CStr(hangle)
c = InStr(a, ".")
If c = 0 Then
        angle = FormatNumber(a, 0)
        minutes = 0
        seconds = 0
Else
        a = a + "0000"
        b = leftstr(a, c - 1)
        angle = FormatNumber(b, 0)
        minutes = FormatNumber(Mid(a, c + 1, 2), 0)
        seconds = FormatNumber(Mid(a, c + 3, 2), 0)
End If

End Sub

Sub setup_tbs()

Dim a As Integer

If Not tablesopen Then
    Call open_tbs
End If

If unitlimits > 0 Then
    
    For a = 1 To unitlimits
        rsunits.addnew
        rsunits.Fields(0) = limitname(a)
        rsunits.Fields("x_sw") = limits(a, 1)
        rsunits.Fields("y_sw") = limits(a, 2)
        rsunits.Fields("x_ne") = limits(a, 3)
        rsunits.Fields("y_ne") = limits(a, 4)
        rsunits.Update
    Next a
    
End If

If npoles > 0 Then
    
    For a = 1 To npoles
        rsprisms.addnew
        rsprisms.Fields("name") = polename(a)
        rsprisms.Fields("height") = poledata(a, 1)
        rsprisms.Fields("offset") = poledata(a, 2)
        rsprisms.Update
    Next a
    
End If

If ndatums > 0 Then
    
    For a = 1 To ndatums
        rsdatums.addnew
        rsdatums.Fields("name") = datumname(a)
        rsdatums.Fields("x") = datum(a, 1)
        rsdatums.Fields("y") = datum(a, 2)
        rsdatums.Fields("z") = datum(a, 2)
        rsdatums.Update
    Next a
    
End If

End Sub

Sub setup_buttons()

Dim inidata(100, 2)
Dim iniclass
Dim a As Integer
Dim b As Integer
Dim status As Integer
Dim flag As Boolean

For a = 1 To 6
    For b = 1 To 100
        button_values(a, b) = ""
    Next b
Next a

inidata(1, 1) = "TITLE"
For a = 1 To vars
    inidata(a + 1, 1) = varlist(a)
Next a

flag = False
For a = 1 To 6
    iniclass = "[BUTTON" & Trim(CStr(a)) & "]"
    For b = 1 To vars + 1
        inidata(b, 2) = ""
    Next b
    Call ReadIni(cfgfile, iniclass, inidata, status)
    If inidata(1, 2) <> "" Then
        Select Case a
        Case 1
            frmMain.button1.Caption = inidata(1, 2)
            frmMain.button1.Visible = True
            flag = True
        Case 2
            frmMain.button2.Caption = inidata(1, 2)
            frmMain.button2.Visible = True
            flag = True
        Case 3
            frmMain.button3.Caption = inidata(1, 2)
            frmMain.button3.Visible = True
            flag = True
        Case 4
            frmMain.button4.Caption = inidata(1, 2)
            frmMain.button4.Visible = True
            flag = True
        Case 5
            frmMain.button5.Caption = inidata(1, 2)
            frmMain.button5.Visible = True
            flag = True
        Case 6
            frmMain.button6.Caption = inidata(1, 2)
            frmMain.button6.Visible = True
            flag = True
        Case Else
        End Select
        For b = 2 To vars + 1
            button_values(a, b - 1) = inidata(b, 2)
        Next b
    End If
Next a

If flag = False Then
    frmMain.Frame2.Visible = False
Else
    frmMain.Frame2.Visible = True
End If

End Sub

Sub savevar(a As Integer, filename)

Dim b As Integer
Dim inidata(100, 2)
Dim iniclass
Dim wstatus As Integer

iniclass = "[" + varlist(a) + "]"

inidata(1, 1) = "Type"
Select Case vtype(a)
Case 1
    inidata(1, 2) = "Text"
Case 2
    inidata(1, 2) = "Numeric"
Case 3
    inidata(1, 2) = "Instrument"
Case 4
    inidata(1, 2) = "Menu"
Case 5
    inidata(1, 2) = "Unit"
End Select
    
inidata(2, 1) = "Prompt"
inidata(2, 2) = varprompt(a)
inidata(3, 1) = "Menu"
inidata(3, 2) = vmenu(a)
    
If vlen(a) <> 0 Then
    inidata(4, 1) = "Length"
    inidata(4, 2) = FormatNumber(vlen(a), 0)
End If

If vincr(a) = True Then
    inidata(5, 1) = "Increment"
    inidata(5, 2) = "True"
End If

If vcarry(a) = True Then
    inidata(6, 1) = "Carry"
    inidata(6, 2) = "True"
End If

If vunique(a) = True Then
    inidata(7, 1) = "Unique"
    inidata(7, 2) = "True"
End If
       
Call WriteIni(filename, iniclass, inidata, wstatus)

End Sub

Sub save_button(buttno As Integer, buttcaption, filename)

Dim iniclass
Dim inidata(20, 2)
Dim a As Integer
Dim b As Integer
Dim wstatus

iniclass = "[Button" & LTrim(CStr(buttno)) & "]"
inidata(1, 1) = "Title"
inidata(1, 2) = buttcaption

b = 1
For a = 1 To vars
    If button_values(buttno, a) <> "" Then
        b = b + 1
        inidata(b, 1) = varlist(a)
        inidata(b, 2) = button_values(buttno, a)
    End If
Next a

Call WriteIni(filename, iniclass, inidata, wstatus)

End Sub

Function dms_to_dd(degrees As Integer, minutes As Integer, seconds As Integer) As Double

dms_to_dd = CDbl(degrees) + CDbl(minutes / 60) + CDbl(seconds / 3600)

End Function
Sub open_tbs()

Dim sql As String

rsprisms.Open "prisms", options(1), adOpenKeyset, adLockOptimistic, adCmdTableDirect
rsdatums.Open "datums", options(1), adOpenKeyset, adLockOptimistic, adCmdTableDirect
sql = "SELECT * FROM " & options(22) & " ORDER BY Recno"
rspoints.Open sql, options(1), adOpenKeyset, adLockOptimistic, adCmdText
rsunits.Open "units", options(1), adOpenKeyset, adLockOptimistic, adCmdTableDirect
tablesopen = True

End Sub
