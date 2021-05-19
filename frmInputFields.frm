VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmImportFields 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Import Fields"
   ClientHeight    =   2625
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   372
      Left            =   2448
      TabIndex        =   3
      Top             =   900
      Width           =   1548
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Inport Fields"
      Height          =   372
      Left            =   2448
      TabIndex        =   2
      Top             =   390
      Width           =   1548
   End
   Begin VB.ListBox lstFields 
      Height          =   2085
      Left            =   168
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   2028
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   3810
      Top             =   30
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Select Fields to import by checking them, then click on the Input Fields button.  "
      Height          =   864
      Left            =   2376
      TabIndex        =   4
      Top             =   1392
      Width           =   1704
   End
   Begin VB.Label Label1 
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   96
      Width           =   4980
   End
End
Attribute VB_Name = "frmImportFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempVars As Integer
Dim TempVarlist(50) As String
Dim TempVType(50) As String
Dim TempVPrompt(50) As String
Dim TempVMenu(50) As String
Dim TempVCarry(50) As Boolean
Dim TempVLen(50) As Integer
Dim Inidata(8, 2) As String
Dim Status As Byte
Dim IniClass As String
Public DonorCFG As String

Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command2_Click()

For I = 0 To lstFields.ListCount - 1
    If lstFields.Selected(I) Then
        Gotit = False
        For J = 1 To Vars
            If LCase(lstFields.List(I)) = LCase(VarList(J)) Then
                response = MsgBox(lstFields.List(I) + " already exists in current CFG file." + Chr(13) + "Click Yes to overwrite, No to add anyway, and Cancel to skip", vbYesNoCancel)
                If response = vbCancel Then
                    GoTo Continue
                ElseIf response = vbYes Then
                    VarList(J) = TempVarlist(I)
                    VType(J) = TempVType(I)
                    VPrompt(J) = TempVPrompt(I)
                    VMenu(J) = TempVMenu(I)
                    VLen(J) = TempVLen(I)
                    VCarry(J) = TempVCarry(I)
                    GoTo Continue
                End If
                Exit For
            End If
        Next J
        Vars = Vars + 1
        VarList(Vars) = TempVarlist(I)
        VType(Vars) = TempVType(I)
        VPrompt(Vars) = TempVPrompt(I)
        VMenu(Vars) = TempVMenu(I)
        VLen(Vars) = TempVLen(I)
        VCarry(Vars) = TempVCarry(I)
    End If
Continue:
Next I
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
Unload Me
parsecfg A
frmMain.FormatVarList

End Sub

Private Sub Form_Load()

Me.Height = 3540
Me.Width = 5235
Label1 = DonorCFG
Open DonorCFG For Input As 1
TempVars = -1
Do While Not EOF(1)
    Do
        Line Input #1, ts
        ts = UCase(Trim(ts))
        If Left(ts, 1) = "[" Then
            Select Case ts
            Case "[EDM]", "[BUTTON1]", "[BUTTON2]", "[BUTTON3]", "[BUTTON4]", "[BUTTON5]", "[BUTTON6]", "[UNIT]", "[ID]", "[SUFFIX]", "[PRISM]", "[X]", "[Y]", "[Z]", "[HANGLE]", "[VANGLE]", "[SLOPED]"
            Case Else
                TempVars = TempVars + 1
                TempVarlist(TempVars) = Mid(ts, 2, Len(ts) - 2)
            End Select
        End If
    Loop Until EOF(1)
Loop
Close 1

lstFields.Clear
For I = 0 To TempVars
    lstFields.AddItem TempVarlist(I)
    IniClass = "[" + TempVarlist(I) + "]"
    Inidata(1, 1) = "Type"
    Inidata(2, 1) = "Prompt"
    Inidata(3, 1) = "Menu"
    Inidata(4, 1) = "Length"
    Inidata(5, 1) = "Carry"
    Inidata(1, 2) = ""
    Inidata(2, 2) = ""
    Inidata(3, 2) = ""
    Inidata(4, 2) = ""
    Inidata(5, 2) = ""
    Call ReadIni(DonorCFG, IniClass, Inidata, Status)
    TempVType(I) = UCase(Inidata(1, 2))
    TempVPrompt(I) = Inidata(2, 2)
    TempVMenu(I) = Inidata(3, 2)
    If Inidata(4, 2) <> "" Then TempVLen(I) = CInt(Inidata(4, 2))
    If LCase(Inidata(5, 2)) = "true" Or LCase(Inidata(6, 2)) = "yes" Then TempVCarry(I) = True
Next I
CenterForm Me
Exit Sub

cd1cancel:
Unload Me

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

Private Sub Form_Unload(Cancel As Integer)

Me.Hide
frmMain.Picture1.SetFocus

End Sub


