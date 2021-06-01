VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEditField 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Fields"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   ControlBox      =   0   'False
   Icon            =   "EditField.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "Add Datum Info"
      Height          =   315
      Left            =   5640
      TabIndex        =   34
      Top             =   1830
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton cmdAddDateTime 
      Caption         =   "Add Time"
      Height          =   315
      Index           =   1
      Left            =   7260
      TabIndex        =   33
      Top             =   1830
      Width           =   1065
   End
   Begin VB.CommandButton cmdAddDateTime 
      Caption         =   "Add Date"
      Height          =   315
      Index           =   0
      Left            =   4470
      TabIndex        =   32
      Top             =   1830
      Width           =   1065
   End
   Begin VB.Frame MenuFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3228
      Left            =   75
      TabIndex        =   29
      Top             =   2610
      Width           =   4356
      Begin VB.ListBox lstMenuItems 
         Height          =   2400
         Left            =   90
         TabIndex        =   11
         Top             =   555
         Width           =   2835
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   345
         Index           =   0
         Left            =   3060
         TabIndex        =   14
         Top             =   1980
         Width           =   1245
      End
      Begin VB.Frame Frame4 
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   3060
         TabIndex        =   30
         Top             =   444
         Width           =   1245
         Begin VB.CommandButton cmdUpDown 
            Caption         =   "Up"
            Height          =   375
            Index           =   0
            Left            =   270
            TabIndex        =   12
            Top             =   300
            Width           =   735
         End
         Begin VB.CommandButton cmdUpDown 
            Caption         =   "Down"
            Height          =   375
            Index           =   1
            Left            =   270
            TabIndex        =   13
            Top             =   870
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   345
         Left            =   3060
         TabIndex        =   15
         Top             =   2340
         Width           =   1245
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   345
         Index           =   0
         Left            =   3060
         TabIndex        =   16
         Top             =   2700
         Width           =   1245
      End
      Begin VB.Label Label6 
         Caption         =   "Menu Options:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   31
         Top             =   300
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   345
      Index           =   1
      Left            =   3120
      TabIndex        =   10
      Top             =   2112
      Width           =   1245
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   345
      Index           =   1
      Left            =   3120
      TabIndex        =   9
      Top             =   1752
      Width           =   1245
   End
   Begin VB.Frame Frame3 
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3120
      TabIndex        =   28
      Top             =   240
      Width           =   1245
      Begin VB.CommandButton cmdUpDown 
         Caption         =   "Down"
         Height          =   375
         Index           =   3
         Left            =   255
         TabIndex        =   8
         Top             =   870
         Width           =   735
      End
      Begin VB.CommandButton cmdUpDown 
         Caption         =   "Up"
         Height          =   375
         Index           =   2
         Left            =   255
         TabIndex        =   7
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add V-angle, H-angle and Slope-D"
      Height          =   345
      Left            =   4488
      TabIndex        =   17
      Top             =   2250
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.OptionButton optType 
      Caption         =   "Menu"
      Height          =   255
      Index           =   0
      Left            =   5208
      TabIndex        =   3
      Top             =   1128
      Width           =   735
   End
   Begin VB.OptionButton optType 
      Caption         =   "Text"
      Height          =   255
      Index           =   1
      Left            =   6315
      TabIndex        =   4
      Top             =   1110
      Width           =   615
   End
   Begin VB.OptionButton optType 
      Caption         =   "Numeric"
      Height          =   255
      Index           =   2
      Left            =   7335
      TabIndex        =   5
      Top             =   1110
      Width           =   975
   End
   Begin VB.CheckBox chkCarry 
      Alignment       =   1  'Right Justify
      Caption         =   "Carry Values to New Shots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4488
      TabIndex        =   6
      Top             =   1470
      Width           =   3825
   End
   Begin VB.TextBox txtLength 
      Height          =   285
      Left            =   6672
      TabIndex        =   2
      Top             =   750
      Width           =   1635
   End
   Begin VB.ListBox lstFields 
      Height          =   2010
      Left            =   180
      TabIndex        =   0
      Top             =   345
      Width           =   2835
   End
   Begin VB.TextBox txtPrompt 
      Height          =   285
      Left            =   5112
      TabIndex        =   1
      Top             =   360
      Width           =   3195
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8580
      Top             =   1770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   405
      Left            =   8550
      TabIndex        =   19
      Top             =   960
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   465
      Left            =   8550
      TabIndex        =   18
      Top             =   330
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "NOTE:  Adding or removing fields will require you to create a new point table."
      Height          =   390
      Index           =   3
      Left            =   4485
      TabIndex        =   27
      Top             =   5355
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   $"EditField.frx":000C
      Height          =   690
      Index           =   2
      Left            =   4485
      TabIndex        =   26
      Top             =   4665
      Width           =   5340
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   $"EditField.frx":00D7
      Height          =   1230
      Index           =   1
      Left            =   4485
      TabIndex        =   25
      Top             =   3630
      Width           =   5205
   End
   Begin VB.Label Label5 
      Caption         =   "Fields:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   195
      TabIndex        =   24
      Top             =   75
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "This form allows you to define, edit, or remove the data fields associated with the current CFG file."
      Height          =   450
      Index           =   0
      Left            =   4485
      TabIndex        =   23
      Top             =   3120
      Width           =   4395
   End
   Begin VB.Label Label2 
      Caption         =   "Maximum Length:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4488
      TabIndex        =   22
      Top             =   780
      Width           =   2088
   End
   Begin VB.Label Label3 
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4488
      TabIndex        =   21
      Top             =   1116
      Width           =   492
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Prompt:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   4488
      TabIndex        =   20
      Top             =   420
      Width           =   648
   End
End
Attribute VB_Name = "frmEditField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VarLen(100) As Integer
Dim VarType(100) As String
Dim VarCarry(100) As Boolean
Dim VarMenuList(100) As String
Dim VarPrompt(100) As String
Dim Loading As Boolean
Dim CurrentField As Integer
Dim TotalVars As Integer
Dim OldField(100) As Boolean

Private Sub chkCarry_Click()

If chkCarry = 0 Then
    VarCarry(lstFields.ItemData(lstFields.ListIndex)) = False
Else
    VarCarry(lstFields.ItemData(lstFields.ListIndex)) = True
End If

End Sub

Private Sub cmdAddDateTime_Click(Index As Integer)

Dim ListBox As Object
Set ListBox = lstFields
If Index = 0 Then
    TempString = "Date"
Else
    TempString = "Time"
End If
For I = 0 To ListBox.ListCount - 1
    If LCase(ListBox.List(I)) = LCase(TempString) Then
        cmdAddDateTime(Index).Enabled = False
        Exit Sub
    End If
Next I
    
ListBox.AddItem UCase(TempString), ListBox.ListCount
Loading = True
TotalVars = TotalVars + 1
ListBox.ItemData(ListBox.ListCount - 1) = TotalVars
ListBox.ListIndex = ListBox.ListCount - 1
CurrentField = TotalVars

txtPrompt = UCase(Left(TempString, 1)) + LCase(Mid(TempString, 2))
VarType(CurrentField) = "TEXT"
VarLen(CurrentField) = 12
optType(1) = True
MenuFrame.Visible = False
lstMenuItems.Clear
txtLength = 12
chkCarry = 0
Loading = False
cmdDelete(1).Enabled = True
optType(0).Enabled = False
optType(1).Enabled = False
optType(2).Enabled = False
chkCarry.Enabled = False
txtLength.Enabled = False
cmdAddDateTime(Index).Enabled = False

End Sub

Private Sub cmdDelete_Click(Index As Integer)

Dim ListBox As Control

If Index = 0 Then
    Set ListBox = lstMenuItems
Else
    Set ListBox = lstFields
End If

If ListBox.ListIndex >= 0 Then
    If ListBox.ListIndex < ListBox.ListCount - 1 Then
        TempIndex = ListBox.ListIndex
    Else
        TempIndex = ListBox.ListIndex - 1
    End If
    ListBox.RemoveItem ListBox.ListIndex
    If ListBox.ListCount > 0 Then
        ListBox.Selected(TempIndex) = True
    End If
    If Index = 0 Then
        VarMenuList(lstFields.ItemData(lstFields.ListIndex)) = lstMenuItems.List(0)
        For I = 1 To lstMenuItems.ListCount - 1
            VarMenuList(lstFields.ItemData(lstFields.ListIndex)) = VarMenuList(lstFields.ItemData(lstFields.ListIndex)) + "," + ListBox.List(I)
        Next I
    End If
End If
cmdAddDateTime(0).Enabled = True
cmdAddDateTime(1).Enabled = True

For I = 0 To ListBox.ListCount - 1
    If LCase(ListBox.List(I)) = "date" Then cmdAddDateTime(0).Enabled = False
    If LCase(ListBox.List(I)) = "time" Then cmdAddDateTime(1).Enabled = False
Next I

End Sub

Private Sub cmdEdit_Click()

If lstMenuItems.ListIndex >= 0 Then
    TempString = InputBox("Enter New Value", "Edit Menu Value", lstMenuItems.List(lstMenuItems.ListIndex))
    If TempString <> "" Then
        TempIndex = lstMenuItems.ListIndex
        Value = UCase(TempString)
        lstMenuItems.RemoveItem (TempIndex)
        lstMenuItems.AddItem Value, TempIndex
        lstMenuItems.Selected(TempIndex) = True
        VarMenuList(lstFields.ItemData(lstFields.ListIndex)) = lstMenuItems.List(0)
        For I = 1 To lstMenuItems.ListCount - 1
            VarMenuList(lstFields.ItemData(lstFields.ListIndex)) = VarMenuList(lstFields.ItemData(lstFields.ListIndex)) + "," + lstMenuItems.List(I)
        Next I
    
    End If
End If

End Sub

Private Sub cmdNew_Click(Index As Integer)

Dim ListBox As Control

If Index = 0 Then
    Set ListBox = lstMenuItems
    TempString = InputBox("Enter New Menu Item for " + lstFields.List(lstFields.ListIndex), "Add Value")
    If UpperCase Then TempString = UCase(TempString)
    If Len(TempString) > VarLen(CurrentField) Then
        MsgBox ("Maximum length for this field is set to " & VarLen(CurrentField) & " characters.  Enter a new value or change field size.")
        Exit Sub
    End If
    If TempString <> "" Then
        ListBox.AddItem UCase(TempString), ListBox.ListCount
    End If
    ListBox.ListIndex = ListBox.ListCount - 1
    VarMenuList(lstFields.ItemData(lstFields.ListIndex)) = lstMenuItems.List(0)
    For I = 1 To lstMenuItems.ListCount - 1
        VarMenuList(lstFields.ItemData(lstFields.ListIndex)) = VarMenuList(lstFields.ItemData(lstFields.ListIndex)) + "," + lstMenuItems.List(I)
    Next I
  
Else
    Set ListBox = lstFields
    If ListBox.ListCount = 30 Then
        MsgBox ("A maximum of 30 optional fields can be defined.")
        Exit Sub
    End If
    SaveData
    TempString = InputBox("Enter New Field name", "Add field")
    If UpperCase Then TempString = UCase(TempString)
    If TempString <> "" Then
        Gotit = False
        For I = 0 To ListBox.ListCount - 1
            If LCase(ListBox.List(I)) = LCase(TempString) Then
                MsgBox ("Duplicate field name.")
                Exit Sub
            End If
        Next I
        ListBox.AddItem UCase(TempString), ListBox.ListCount
            
        Loading = True
        TotalVars = TotalVars + 1
        ListBox.ItemData(ListBox.ListCount - 1) = TotalVars
        ListBox.ListIndex = ListBox.ListCount - 1
        CurrentField = TotalVars
        txtPrompt = UCase(Left(TempString, 1)) + LCase(Mid(TempString, 2))
        For I = 0 To 2
            optType(I) = False
        Next I
        MenuFrame.Visible = False
        lstMenuItems.Clear
        txtLength = 15
        VarLen(CurrentField) = 15
        chkCarry = 0
        Loading = False
        cmdDelete(1).Enabled = True
        optType(0).Enabled = True
        optType(1).Enabled = True
        optType(2).Enabled = True
        chkCarry.Enabled = True
        txtLength.Enabled = True

    End If
    
End If

Set ListBox = Nothing

End Sub

Private Sub cmdUpDown_Click(Index As Integer)

Dim TempIndex As Integer
Dim Value As String
Dim TempItemData As Integer
Dim ListBox As Control

If Index < 2 Then
    Set ListBox = lstMenuItems
Else
    Set ListBox = lstFields
End If

TempIndex = ListBox.ListIndex
Value = ListBox.List(ListBox.ListIndex)
TempItemData = ListBox.ItemData(ListBox.ListIndex)
ListBox.RemoveItem (ListBox.ListIndex)

Select Case Index
    Case 0, 2
            If TempIndex > 0 Then TempIndex = TempIndex - 1
    Case 1, 3
            If TempIndex <= ListBox.ListCount - 1 Then TempIndex = TempIndex + 1
End Select

ListBox.AddItem Value, TempIndex
ListBox.ItemData(ListBox.NewIndex) = TempItemData
CurrentField = TempItemData
ListBox.Selected(TempIndex) = True

If Index < 2 Then
    VarMenuList(lstFields.ItemData(lstFields.ListIndex)) = lstMenuItems.List(0)
    For I = 1 To lstMenuItems.ListCount - 1
        VarMenuList(lstFields.ItemData(lstFields.ListIndex)) = VarMenuList(lstFields.ItemData(lstFields.ListIndex)) + "," + lstMenuItems.List(I)
    Next I
End If
Set ListBox = Nothing

End Sub

Private Sub Command1_Click()

Dim GotHangle As Boolean
Dim GotVangle As Boolean
Dim GotSloped As Boolean

For I = 0 To lstFields.ListCount - 1
    Select Case UCase(lstFields.List(I))
        Case "HANGLE"
            GotHangle = True
        Case "VANGLE"
            GotVangle = True
        Case "SLOPED"
            GotSloped = True
    End Select
Next I
If Not GotHangle Then lstFields.AddItem "HANGLE"
If Not GotVangle Then lstFields.AddItem "VANGLE"
If Not GotSloped Then lstFields.AddItem "SLOPED"

End Sub

Private Sub Command2_Click()

For I = 0 To lstFields.ListCount - 1
    TempIndex = lstFields.ItemData(I)
    If VarPrompt(TempIndex) = "" Then
        MsgBox ("Enter prompt for " + lstFields.List(I))
        Exit Sub
    End If
    If VarLen(TempIndex) = 0 Then
        MsgBox ("Enter length for " + lstFields.List(I))
        Exit Sub
    End If
    If VarType(TempIndex) = "" Then
        MsgBox ("Select appropriate Type for " + lstFields.List(I))
        Exit Sub
    End If
        If VarType(TempIndex) = "MENU" And VarMenuList(TempIndex) = "" Then
        MsgBox ("Enter menu options for " + lstFields.List(I))
        Exit Sub
    End If
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
If Trim(CurrentStation.Name) <> "" Then
    Print #1, "StationName="; CurrentStation.Name
    Print #1, "StationX="; CurrentStation.X
    Print #1, "stationY="; CurrentStation.y
    Print #1, "stationZ="; CurrentStation.z
End If
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

For I = 0 To lstFields.ListCount - 1
    TempIndex = lstFields.ItemData(I)
    Print #1, "[" + lstFields.List(I) + "]"
    Print #1, "Prompt="; VarPrompt(TempIndex)
    Print #1, "Length="; VarLen(TempIndex)
    Print #1, "Type="; VarType(TempIndex)
    If VarType(TempIndex) = "MENU" Then
        If UpperCase Then
            Print #1, "Menu=" + UCase(VarMenuList(TempIndex))
        Else
            Print #1, "Menu=" + VarMenuList(TempIndex)
        End If
    End If
    If VarCarry(TempIndex) Then
        Print #1, "Carry=True"
    End If
    Print #1, ""
Next I

Close 1
parsecfg A
frmMain.FormatVarList
frmMain.ShowValues
Unload Me

End Sub

Private Sub Command3_Click()

Cancelling = True
Unload Me

End Sub

Private Sub Command4_Click()

Dim ListBox As Object
'Set ListBox = lstFields
'If Index = 0 Then
'    TempString = "Date"
'Else
'    TempString = "Time"
'End If
'For I = 0 To ListBox.ListCount - 1
'    If LCase(ListBox.List(I)) = LCase(TempString) Then
'        cmdAddDateTime(Index).Enabled = False
'        Exit Sub
'    End If
'Next I
'
'
'ListBox.AddItem UCase(TempString), ListBox.ListCount
'Loading = True
'TotalVars = TotalVars + 1
'ListBox.ItemData(ListBox.ListCount - 1) = TotalVars
'ListBox.ListIndex = ListBox.ListCount - 1
'CurrentField = TotalVars
'
'txtPrompt = UCase(Left(TempString, 1)) + LCase(Mid(TempString, 2))
'VarType(CurrentField) = "TEXT"
'VarLen(CurrentField) = 12
'optType(1) = True
'MenuFrame.Visible = False
'lstMenuItems.Clear
'txtLength = 12
'chkCarry = 0
'Loading = False
'cmdDelete(1).Enabled = True
'optType(0).Enabled = False
'optType(1).Enabled = False
'optType(2).Enabled = False
'chkCarry.Enabled = False
'txtLength.Enabled = False
'cmdAddDateTime(Index).Enabled = False

End Sub

Private Sub Form_Load()

Dim GotUnit As Boolean
Dim GotID As Boolean
Dim GotSuffix As Boolean
Dim GotPrism As Boolean
Dim GotX As Boolean
Dim GotY As Boolean
Dim GotZ As Boolean
Dim I As Integer

Loading = True
For I = 1 To Vars
    OldField(I) = True
    Select Case UCase(VarList(I))
        Case "UNIT"
            GotUnit = True
            VType(I) = "UNIT"
            VCarry(I) = False
        Case "ID"
            GotID = True
            VType(I) = "TEXT"
            VCarry(I) = False
        Case "SUFFIX"
            GotSuffix = True
            VLen(I) = 10
            VType(I) = "NUMERIC"
            VCarry(I) = False
        Case "PRISM"
            GotPrism = True
            VType(I) = "NUMERIC"
            VLen(I) = 10
            VCarry(I) = False
        Case "X"
            GotX = True
            VType(I) = "NUMERIC"
            VLen(I) = 10
            VCarry(I) = False
        Case "Y"
            GotY = True
            VType(I) = "NUMERIC"
            VLen(I) = 10
            VCarry(I) = False
        Case "Z"
            GotZ = True
            VType(I) = "NUMERIC"
            VLen(I) = 10
            VCarry(I) = False
        Case "DATE"
            cmdAddDateTime(0).Enabled = False
            VType(I) = "TEXT"
            VLen(I) = 12
            
        Case "TIME"
            VType(I) = "TEXT"
            VLen(I) = 12
            cmdAddDateTime(1).Enabled = False
            
    End Select
    lstFields.AddItem VarList(I)
    lstFields.ItemData(lstFields.NewIndex) = I
    lstFields.Selected(lstFields.NewIndex) = True
    VarPrompt(I) = VPrompt(I)
    VarLen(I) = VLen(I)
    VarType(I) = VType(I)
    VarCarry(I) = VCarry(I)
    VarMenuList(I) = VMenu(I)
    
Next I
If Not GotUnit Then
    lstFields.AddItem "UNIT"
    I = lstFields.NewIndex
    lstFields.Selected(I) = True
    lstFields.ItemData(I) = I
    VarPrompt(I) = "Unit"
    VarLen(I) = 6
    VarType(I) = "UNIT"
    VarCarry(I) = False
End If
If Not GotID Then
    lstFields.AddItem "ID"
    I = lstFields.NewIndex
    lstFields.Selected(I) = True
    lstFields.ItemData(I) = I
    VarPrompt(I) = "ID"
    VarLen(I) = 5
    VarType(I) = "TEXT"
    VarCarry(I) = False
End If
If Not GotSuffix Then
    lstFields.AddItem "SUFFIX"
    I = lstFields.NewIndex
    lstFields.Selected(I) = True
    lstFields.ItemData(I) = I
    VarPrompt(I) = "Suffix"
    VarLen(I) = 10
    VarType(I) = "NUMERIC"
    VarCarry(I) = False
End If
If Not GotPrism Then
    lstFields.AddItem "PRISM"
    I = lstFields.NewIndex
    lstFields.Selected(I) = True
    lstFields.ItemData(I) = I
    VarPrompt(I) = "Prism"
    VarLen(I) = 10
    VarType(I) = "NUMERIC"
    VarCarry(I) = False
End If
If Not GotX Then
    lstFields.AddItem "X"
    I = lstFields.NewIndex
    lstFields.Selected(I) = True
    lstFields.ItemData(I) = I
    VarPrompt(I) = "X"
    VarLen(I) = 10
    VarType(I) = "NUMERIC"
    VarCarry(I) = False
End If
If Not GotY Then
    lstFields.AddItem "Y"
    I = lstFields.NewIndex
    lstFields.Selected(I) = True
    lstFields.ItemData(I) = I
    VarPrompt(I) = "Y"
    VarLen(I) = 10
    VarType(I) = "NUMERIC"
    VarCarry(I) = False
End If
If Not GotZ Then
    lstFields.AddItem "Z"
    I = lstFields.NewIndex
    lstFields.Selected(I) = True
    lstFields.ItemData(I) = I
    VarPrompt(I) = "Z"
    VarLen(I) = 10
    VarType(I) = "NUMERIC"
    VarCarry(I) = False
End If

Loading = True
If lstFields.ListCount > 0 Then
    TotalVars = lstFields.ListCount
End If
Loading = False
CurrentField = 0
lstFields.ListIndex = 0
CenterForm Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

Me.Hide
frmMain.Picture1.SetFocus

End Sub

Private Sub lstFields_Click()

If Not Loading Then
    SaveData
Else
    Exit Sub
End If

lstMenuItems.Clear
Index = lstFields.ItemData(lstFields.ListIndex)
CurrentField = Index
Select Case UCase(lstFields.List(lstFields.ListIndex))
    Case "UNIT", "ID"
        cmdDelete(1).Enabled = False
        optType(0).Enabled = False
        optType(1).Enabled = False
        optType(2).Enabled = False
        chkCarry.Enabled = False
        txtLength.Enabled = True
    Case "SUFFIX", "PRISM", "X", "Y", "Z"
        cmdDelete(1).Enabled = False
        optType(0).Enabled = False
        optType(1).Enabled = False
        optType(2).Enabled = False
        chkCarry.Enabled = False
        txtLength.Enabled = False
    Case "DATE", "TIME"
        cmdDelete(1).Enabled = True
        optType(0).Enabled = False
        optType(1).Enabled = False
        optType(2).Enabled = False
        chkCarry.Enabled = False
        txtLength.Enabled = False
    
    Case Else
        cmdDelete(1).Enabled = True
        optType(0).Enabled = True
        optType(1).Enabled = True
        optType(2).Enabled = True
        chkCarry.Enabled = True
        txtLength.Enabled = True
End Select

txtLength = VarLen(Index)
txtPrompt = VarPrompt(Index)
For I = 0 To 2
    optType(I) = False
Next I
Select Case UCase(VarType(Index))
    Case "MENU"
        optType(0) = True
            MenuString = VarMenuList(Index)
            If UpperCase Then
                MenuString = UCase(MenuString)
            End If
            Gotit = False
            Do Until Gotit
                X = InStr(MenuString, ",")
                If X > 0 Then
                    lstMenuItems.AddItem Left(MenuString, X - 1)
                    MenuString = Trim(Mid(MenuString, X + 1))
                Else
                    lstMenuItems.AddItem Trim(MenuString)
                    Gotit = True
                End If
            Loop
            If lstMenuItems.ListCount > 0 Then
                cmdUpDown(0).Enabled = True
                cmdUpDown(0).Enabled = True
                lstMenuItems.Selected(0) = True
            Else
                cmdUpDown(0).Enabled = True
                cmdUpDown(0).Enabled = True
            End If
    Case "TEXT", "UNIT", "DATE", "TIME"
        optType(1) = True
        
    Case "NUMERIC", "INSTRUMENT", "X", "Y", "Z"
        optType(2) = True
End Select
If VarCarry(Index) Then
    chkCarry = 1
Else
    chkCarry = False
End If

End Sub

Public Sub SaveData()

If CurrentField = 0 Then Exit Sub

VarLen(CurrentField) = Val(txtLength)
If optType(0) Then
    VarType(CurrentField) = "MENU"
ElseIf optType(1) Then
    VarType(CurrentField) = "TEXT"
ElseIf optType(2) Then
    VarType(CurrentField) = "NUMERIC"
End If
If chkCarry = 1 Then
    VarCarry(CurrentField) = True
End If
If lstMenuItems.ListCount > 0 Then
    TempString = lstMenuItems.List(0)
    For I = 1 To lstMenuItems.ListCount - 1
        TempString = TempString + "," + lstMenuItems.List(I)
    Next I
    VarMenuList(CurrentField) = TempString
End If

End Sub

Private Sub optType_Click(Index As Integer)

If VarType(CurrentField) = "UNIT" Then
    optType(1) = True
End If
Select Case Index
    Case 0
        VarType(CurrentField) = "MENU"
        lstMenuItems.Clear
        MenuFrame.Visible = True
    Case 1
        
        VarType(CurrentField) = "TEXT"
        VarMenuList(CurrentField) = ""
        MenuFrame.Visible = False

    Case 2
        VarType(CurrentField) = "NUMERIC"
        txtLength = 10
        VarMenuList(CurrentField) = ""
        MenuFrame.Visible = False
End Select

End Sub

Private Sub txtLength_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 8, 48 To 57
    Case Else
        KeyAscii = 0
End Select

End Sub

Private Sub txtLength_LostFocus()

If Val(txtLength) > 255 Then
    MsgBox ("Enter length less than 256")
    txtLength.SetFocus
    txtLength.SelStart = 0
    txtLength.SelLength = Len(txtLength)
    Exit Sub
End If
If OldField(CurrentField) Then
    If VarLen(CurrentField) <> Val(txtLength) Then
        response = MsgBox("Changing the length of a field can have serious consequences, " + Chr(10) + Chr(13) + "including the need to manually alter field size in the receiving database." + Chr(10) + Chr(13) + "Continue?", vbOKCancel)
        If response = vbCancel Then
            txtLength.SetFocus
            txtLength.SelStart = 0
            txtLength.SelLength = Len(txtLength)
        Else
            VarLen(CurrentField) = Val(txtLength)
        End If
    End If
Else
    VarLen(CurrentField) = Val(txtLength)
End If

End Sub

Private Sub txtPrompt_Change()

VarPrompt(CurrentField) = txtPrompt

End Sub

Private Sub txtPrompt_KeyPress(KeyAscii As Integer)

If Len(txtPrompt) > 15 Then
    Select Case KeyAscii
        Case 8
        Case Else
            KeyAscii = 0
    End Select
End If

End Sub


