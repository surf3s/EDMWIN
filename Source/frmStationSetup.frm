VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStationSetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Station Setup"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   9075
   ControlBox      =   0   'False
   Icon            =   "frmStationSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Current Station"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   30
      Top             =   3720
      Width           =   2595
      Begin VB.ComboBox Station 
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   38
         Top             =   270
         Width           =   2205
      End
      Begin VB.TextBox txtX 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   288
         Index           =   0
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   720
         Width           =   1050
      End
      Begin VB.TextBox txtY 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   288
         Index           =   0
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1050
      End
      Begin VB.TextBox txtZ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   288
         Index           =   0
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1050
      End
      Begin VB.TextBox txtStationHeight 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   1485
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1830
         Width           =   930
      End
      Begin VB.TextBox txtReferenceAngle 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   1485
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2160
         Width           =   930
      End
      Begin VB.CommandButton cmdHangle 
         Caption         =   "Set Horizontal Angle"
         Height          =   330
         Left            =   120
         TabIndex        =   32
         Top             =   2640
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.CommandButton calccomp 
         Caption         =   "Flip 180 Degrees"
         Height          =   330
         Left            =   120
         TabIndex        =   31
         Top             =   3000
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   43
         Top             =   720
         Width           =   195
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y :"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   42
         Top             =   1080
         Width           =   195
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Z :"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   41
         Top             =   1440
         Width           =   195
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument Height :"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   1830
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horizontal Angle :"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   2160
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdAcceptStation 
      Caption         =   "Accept"
      Enabled         =   0   'False
      Height          =   500
      Left            =   2760
      TabIndex        =   28
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   500
      Left            =   4680
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1695
   End
   Begin VB.ComboBox txtPrism 
      Height          =   315
      Left            =   6510
      Sorted          =   -1  'True
      TabIndex        =   26
      Text            =   "Select Prism"
      Top             =   3345
      Width           =   2385
   End
   Begin VB.Frame frmCoordinates 
      Caption         =   "Station Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2880
      TabIndex        =   24
      Top             =   5640
      Visible         =   0   'False
      Width           =   5655
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   975
         Left            =   150
         TabIndex        =   25
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1720
         _Version        =   393216
         Rows            =   3
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   0
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin VB.Label lblST 
         Height          =   195
         Left            =   1410
         TabIndex        =   46
         Top             =   1260
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Station Height:"
         Height          =   195
         Left            =   150
         TabIndex        =   45
         Top             =   1260
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Click Accept button to continue with this station"
         Height          =   705
         Left            =   3360
         TabIndex        =   29
         Top             =   360
         Width           =   2010
      End
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "Record Datum"
      Height          =   855
      Index           =   0
      Left            =   4680
      TabIndex        =   22
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Setup Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   5780
      Begin VB.OptionButton SetUpTypes 
         Caption         =   $"frmStationSetup.frx":000C
         Height          =   675
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   5355
      End
      Begin VB.OptionButton SetUpTypes 
         Caption         =   $"frmStationSetup.frx":00C4
         Height          =   555
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   5295
      End
      Begin VB.OptionButton SetUpTypes 
         Caption         =   $"frmStationSetup.frx":017C
         Height          =   585
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   5340
      End
      Begin VB.OptionButton SetUpTypes 
         Caption         =   $"frmStationSetup.frx":0234
         Height          =   420
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   270
         Width           =   5355
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Secondary Reference Datum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Index           =   1
      Left            =   6000
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   2985
      Begin VB.CommandButton cmdRecord 
         Caption         =   "Record Datum 2"
         Default         =   -1  'True
         Height          =   855
         Index           =   1
         Left            =   1800
         TabIndex        =   44
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtZ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   288
         Index           =   2
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1050
      End
      Begin VB.TextBox txtY 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   288
         Index           =   2
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1050
      End
      Begin VB.TextBox txtX 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   288
         Index           =   2
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   720
         Width           =   1050
      End
      Begin VB.ComboBox Station 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   336
         Width           =   2205
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Z :"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   1440
         Width           =   225
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Y :"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   14
         Top             =   1080
         Width           =   225
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         Caption         =   "X :"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   13
         Top             =   750
         Width           =   225
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Primary Reference Datum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Index           =   0
      Left            =   2880
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   3015
      Begin VB.ComboBox Station 
         Height          =   315
         Index           =   1
         Left            =   168
         TabIndex        =   7
         Top             =   336
         Width           =   2205
      End
      Begin VB.TextBox txtX 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   288
         Index           =   1
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   720
         Width           =   1050
      End
      Begin VB.TextBox txtY 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   288
         Index           =   1
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1050
      End
      Begin VB.TextBox txtZ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   288
         Index           =   1
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "X :"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   720
         Width           =   195
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Y :"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   1080
         Width           =   195
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Z :"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   1440
         Width           =   195
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   3000
      TabIndex        =   47
      Top             =   5760
      Visible         =   0   'False
      Width           =   5505
   End
   Begin VB.Label lblPrism 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prism : "
      Height          =   195
      Left            =   5985
      TabIndex        =   23
      Top             =   3435
      Width           =   510
   End
   Begin VB.Label setuphelp 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   6000
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   2880
   End
End
Attribute VB_Name = "frmStationSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HangleSet As Boolean
Dim Datum1X As Double
Dim Datum1Y As Double
Dim Datum1Z As Double
Dim Datum2X As Double
Dim Datum2Y As Double
Dim Datum2Z As Double
Dim FirstXp As Double
Dim FirstYp As Double
Dim FirstZp As Double
Dim CurrentXP As Double
Dim CurrentYP As Double
Dim CurrentZP As Double

Private Sub calccomp_Click()

Dim hangle As Single
Dim angle As Integer
Dim minutes As Integer
Dim seconds As Integer

If txtReferenceAngle.Text <> "" Then
    hangle = CSng(txtReferenceAngle.Text)
    Call parseangle(hangle, angle, minutes, seconds)
    angle = angle + 180
    If angle >= 360 Then
        angle = angle - 360
    End If
    ms$ = Right$("00" + Trim$(Str$(minutes)), 2)
    ss$ = Right$("00" + Trim$(Str$(seconds)), 2)
    txtReferenceAngle.Text = Trim$(Str$(angle)) + "." + ms$ + ss$
End If

End Sub

Private Sub cmdAcceptStation_Click()

If txtstationheight = "" And SetUpTypes(0) Then
    MsgBox ("Warning: You have not included a station height.")
'     Exit Sub
End If

If SetUpTypes(0) And Not HangleSet Then
    MsgBox ("Set Horizontal Angle first")
    Exit Sub
End If

'If Not SetUpTypes(0) And txtPrism.ListIndex = -1 Then
'    MsgBox ("Select Prism first")
'    Exit Sub
'End If
    
'all points will now be offset to this currentstation

If SetUpTypes(2) Then
    ''txtX(0) = Format(CDbl(txtX(1)) - edmshot.X, "######0.000")
    ''txtY(0) = Format(CDbl(txtY(1)) - edmshot.y, "######0.000")
    txtZ(0) = Format(CDbl(txtZ(1)) - (edmshot.z - edmshot.poleh), "######0.000")
    stationheight = CSng(txtstationheight.Text)
Else
    stationheight = 0
End If
CurrentStation.X = txtX(0)
CurrentStation.y = txtY(0)

If SetUpTypes(0) Then
    If txtstationheight.Text <> "" Then
        CurrentStation.z = CSng(txtZ(0)) + CSng(txtstationheight.Text)
        stationheight = CSng(txtstationheight.Text)
    Else
        CurrentStation.z = CSng(txtZ(0))
        stationheight = 0
    End If
Else
    CurrentStation.z = CSng(txtZ(0))
End If

CurrentStation.Name = Station(0)
StationName = CurrentStation.Name
StationInitialized = True
frmMain.lblStationWarning.Visible = False
mdiMain.StatusBar.Panels(5) = "Current Station: " + Station(0) + "  "
'Update the cfg file so autoresume will work
Dim Inidata(7, 2) As String
Dim IniClass As String
Dim Status As Byte
IniClass = "[EDM]"
Inidata(1, 1) = "CurrentStation"
Inidata(1, 2) = CurrentStation.Name
Inidata(2, 1) = "StationX"
Inidata(2, 2) = CurrentStation.X
Inidata(3, 1) = "stationY"
Inidata(3, 2) = CurrentStation.y
Inidata(4, 1) = "stationZ"
Inidata(4, 2) = CurrentStation.z
Inidata(5, 1) = "ReferenceDatum"
Inidata(5, 2) = RefDatum1
Inidata(6, 1) = "ReferenceDatum2"
Inidata(6, 2) = RefDatum2
Inidata(7, 1) = "SetupType"

If SetUpTypes(0) Then
    Inidata(7, 2) = 0
ElseIf SetUpTypes(1) Then
    Inidata(7, 2) = 1
ElseIf SetUpTypes(2) Then
    Inidata(7, 2) = 2
ElseIf SetUpTypes(3) Then
    Inidata(7, 2) = 3
End If
Call WriteIni(CFGName, IniClass, Inidata(), Status)
MsgBox ("If possible, you should verify location using an independant datum.")

Unload Me

End Sub


Private Sub cmdHangle_Click()

answer = MsgBox("Aim the total station at datum and press OK to set the horizontal angle to " + txtReferenceAngle, vbOKCancel)
If answer = 2 Then
    Exit Sub
End If

Screen.MousePointer = 11

If frmCoordinates.Visible = True Then
    cmdAcceptStation.Default = True
Else
    cmdRecord(0).Default = True
End If

angle$ = txtReferenceAngle

If angle$ <> "" And LCase(EDMName) <> "simulate" Then
    Call sethortangle(angle$, deg, min, sec)
End If
HangleSet = True

Screen.MousePointer = 1

End Sub

Private Sub cmdRecord_Click(Index As Integer)

Dim x1 As Double
Dim x2 As Double
Dim y1 As Double
Dim y2 As Double
Dim edmpoffset As Single
Dim errorcode
Dim temp_stationx As Double
Dim temp_stationy As Double
Dim temp_stationz As Double
Dim degrees As Integer
Dim minutes As Integer
Dim seconds As Integer
Dim stationx As Double
Dim stationy As Double
Dim StationZ As Double

Dim CurrentstationX As Double
Dim CurrentstationY As Double
Dim CurrentStationZ As Double
Dim DistanceAB As Double
Dim DistanceAS As Double
Dim DistanceBS As Double
Dim SideOpposite, SideAdjacent As Double
Dim SinAngle, CosAngle, ForesightAB As Double
Dim ForesightBS As Double
Dim ForesightSB As Double
Dim AngleDifference As Double
Dim horizontalchange As Double
Dim DefinedDistance As Double
Dim MeasuredDistance As Double
Dim MeasuredAngle As Double
Dim DefinedAngle As Double

cmdRecord(Index).Enabled = False

If Not SetUpTypes(0) And frmMain.txtprism.ListCount = 0 Then
    MsgBox ("No prisms defined - cannot initialize station")
    cmdRecord(Index).Enabled = True
    Exit Sub
End If
Label6.Visible = False
lblST.Visible = False
If SetUpTypes(1) Then 'over unknown - shooting to known
    
    If Not HangleSet And txtReferenceAngle <> "" Then
        answer = MsgBox("Aim the total station at " + Station(1).Text + " and press OK to set the horizontal angle to " + txtReferenceAngle, vbOKCancel)
        If answer = 2 Then
            cmdRecord(Index).Enabled = True
            Exit Sub
        End If
        angle$ = txtReferenceAngle
        Call sethortangle(angle$, deg, min, sec)
        HangleSet = True
    End If
    
    answer = MsgBox("Press OK to record " + Station(1).Text, vbOKCancel)
    If answer = 2 Then
        cmdRecord(Index).Enabled = True
        Exit Sub
    End If
    
    Call takeshot_nostation(AskForPrism)
    If UCase(EDMName$) = "LEICA" Or UCase(EDMName$) = "WILD" Or UCase(EDMName$) = "WILD2" Or UCase(EDMName$) = "LEICA_AUTOTILT" Then
        If edmshot.edmpoffset = 0.004 Then
            response = MsgBox("Warning: Instrument in Reflectorless mode.", vbOKCancel)
            If response = vbCancel Then Cancelling = True
        End If
    ElseIf UCase(EDMName$) = "BUILDER" Then
        response = MsgBox("Warning: Instrument in Reflectorless mode.", vbOKCancel)
        If response = vbCancel Then Cancelling = True
    End If
    If Cancelling Or errorcode <> 0 Then
        cmdRecord(Index).Enabled = True
        Cancelling = False
        Exit Sub
    End If
    
    txtX(0) = Format(CDbl(txtX(1)) - edmshot.X, "######0.000")
    txtY(0) = Format(CDbl(txtY(1)) - edmshot.y, "######0.000")
    txtZ(0) = Format(CDbl(txtZ(1)) - edmshot.z - edmshot.poleh, "######0.000")
    
    cmdAcceptStation.Enabled = True
    
ElseIf SetUpTypes(2) Then  'over known and shooting to known
    
    If Station(1).Text = Station(0).Text Then
        MsgBox "You have to select two different datums: the one you are currently setup over and a reference datum.", vbInformation
        cmdRecord(Index).Enabled = True
        mdiMain.StatusBar.Panels(6).Visible = False
    Exit Sub
    End If
    
    x1 = txtX(0)
    x2 = txtX(1)
    y1 = txtY(0)
    y2 = txtY(1)
    
    If Not HangleSet Then
        Call computeangle(x2, y2, x1, y1, deg, min, sec)
        
        ms$ = Right$("00" + Trim$(Str$(min)), 2)
        ss$ = Right$("00" + Trim$(Str$(sec)), 2)
        temp$ = Trim$(Str$(deg)) + "." + ms$ + ss$
            
        answer = MsgBox("Aim the total station at " + Station(1).Text + " and press OK to set the horizontal angle to " + temp$, vbOKCancel)
        If answer = 2 Then
            cmdRecord(Index).Enabled = True
            Exit Sub
        End If
        
        Screen.MousePointer = 11
        angle$ = ""
        Call sethortangle(angle$, deg, min, sec)
        HangleSet = True
        delay 1
        Screen.MousePointer = 1
    End If
    
    answer = MsgBox("Press OK to record " + Station(1).Text, vbOKCancel)
    If answer = 2 Then
        cmdRecord(Index).Enabled = True
        Exit Sub
    End If
    Cancelling = False
    Call takeshot_nostation(AskForPrism)
    If UCase(EDMName$) = "LEICA" Or UCase(EDMName$) = "WILD" Or UCase(EDMName$) = "WILD2" Or UCase(EDMName$) = "LEICA_AUTOTILT" Then
        If edmshot.edmpoffset = 0.004 Then
            response = MsgBox("Warning: Instrument in Reflectorless mode.", vbOKCancel)
            If response = vbCancel Then Cancelling = True
        End If
    ElseIf UCase(EDMName$) = "BUILDER" Then
        response = MsgBox("Warning: Instrument in Reflectorless mode.", vbOKCancel)
        If response = vbCancel Then Cancelling = True
    End If
    If errorcode <> 0 Or Cancelling Then
        cmdRecord(Index).Enabled = True
        mdiMain.StatusBar.Panels(6).Visible = False
        Cancelling = False
    Exit Sub
    End If
    
'    temp_stationx = txtX(0)
'    temp_stationy = txtY(0)
'    temp_stationz = txtZ(0)
    temp_stationx = Format(CDbl(txtX(1)) - edmshot.X, "######0.000")
    temp_stationy = Format(CDbl(txtY(1)) - edmshot.y, "######0.000")
    temp_stationz = Format(CDbl(txtZ(1)) - (edmshot.z - edmshot.poleh), "######0.000")
'    txtX(0) = Format(CDbl(txtX(1)) - edmshot.x, "######0.000")
'    txtY(0) = Format(CDbl(txtY(1)) - edmshot.y, "######0.000")
'    txtZ(0) = Format(CDbl(txtZ(1)) - edmshot.z, "######0.000")
    
    cmdAcceptStation.Enabled = True
    GoTo ShowCoordinates
    
ElseIf SetUpTypes(3) Then
    
    Select Case Index
    Case 0
        If MsgBox("Aim the total station at " + Station(1).Text + " and press Ok.", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
    
        Call sethortangle("", 0, 0, 0)
        Call takeshot_nostation(AskForPrism)
        If UCase(EDMName$) = "LEICA" Or UCase(EDMName$) = "WILD" Or UCase(EDMName$) = "WILD2" Or UCase(EDMName$) = "LEICA_AUTOTILT" Then
            If edmshot.edmpoffset = 0.004 Then
                response = MsgBox("Warning: Instrument in Reflectorless mode.", vbOKCancel)
                If response = vbCancel Then Cancelling = True
            End If
        End If
        If Cancelling Or errorcode <> 0 Then
            cmdRecord(Index).Enabled = True
            mdiMain.StatusBar.Panels(6).Visible = False
            Cancelling = False
            Exit Sub
        End If
    
        Datum1X = txtX(1).Text
        Datum1Y = txtY(1).Text
        Datum1Z = txtZ(1).Text
        FirstXp = edmshot.X
        FirstYp = edmshot.y
'        FirstZp = edmshot.z
        FirstZp = edmshot.z - edmshot.poleh
        cmdRecord(1).Enabled = True
        cmdRecord(1).Visible = True
        cmdRecord(0).Enabled = True
        
    Case 1
        If txtX(2).Text = "" Or txtY(2).Text = "" Or txtZ(2).Text = "" Then
            MsgBox "Select a second datum point."
            Exit Sub
        End If
        
        If MsgBox("Aim the total station at " + Station(2).Text + " and press Ok.", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
           
        Call takeshot_nostation(AskForPrism)
        If UCase(EDMName$) = "LEICA" Or UCase(EDMName$) = "WILD" Or UCase(EDMName$) = "WILD2" Or UCase(EDMName$) = "LEICA_AUTOTILT" Then
            If edmshot.edmpoffset = 0.004 Then
                response = MsgBox("Warning: Instrument in Reflectorless mode.", vbOKCancel)
                If response = vbCancel Then Cancelling = True
            End If
        End If
        If Cancelling Or errorcode <> 0 Then
            cmdRecord(Index).Enabled = True
            mdiMain.StatusBar.Panels(6).Visible = False
            Cancelling = False
            Exit Sub
        End If
           
        Datum2X = txtX(2).Text
        Datum2Y = txtY(2).Text
        Datum2Z = txtZ(2).Text

        CurrentXP = edmshot.X
        CurrentYP = edmshot.y
        CurrentZP = edmshot.z - edmshot.poleh
'        CurrentZP = edmshot.z
        DefinedDistance = Sqr((Datum1X - Datum2X) ^ 2 + (Datum1Y - Datum2Y) ^ 2 + (Datum1Z - Datum2Z) ^ 2)
        MeasuredDistance = Sqr((FirstXp - CurrentXP) ^ 2 + (FirstYp - CurrentYP) ^ 2 + (FirstZp - CurrentZP) ^ 2)
        If Abs(DefinedDistance - MeasuredDistance) > 0.01 Then
            Label7.ForeColor = &HFF&
        Else
            Label7.ForeColor = 0
        End If
        Label7 = "Error in measured distance between first and second datums is " & Format(Abs(DefinedDistance - MeasuredDistance), "####0.000")
        Label7.Visible = True
        
        FirstYp = -FirstYp
        CalcAngle CurrentXP, CurrentYP + FirstYp, 0, 0, MeasuredAngle
        CalcAngle Datum2X, Datum2Y, Datum1X, Datum1Y, DefinedAngle
        AngleDifference = DefinedAngle - MeasuredAngle
        
        CosAngle = Cos(AngleDifference * 1.74532925199433E-02)
        SinAngle = Sin(AngleDifference * 1.74532925199433E-02)
        Distance = Abs(FirstYp)
        
        CurrentstationX = FirstYp * SinAngle + Datum1X
        CurrentstationY = FirstYp * CosAngle + Datum1Y
        CurrentStationZ = ((Datum1Z - FirstZp) + (Datum2Z - CurrentZP)) / 2
        
        txtX(0).Text = Format(CurrentstationX, "######0.000")
        txtY(0).Text = Format(CurrentstationY, "######0.000")
        txtZ(0).Text = Format(CurrentStationZ, "######0.000")
        
        CalcAngle Datum2X, Datum2Y, CurrentstationX, CurrentstationY, ForesightSB
        Call conv_angle_to_degminsec(ForesightSB, degrees, minutes, seconds)
        txtReferenceAngle = Trim(Str(degrees)) + "." + Right("00" + Trim(Str(minutes)), 2) + Right("00" + Trim(Str(seconds)), 2)
        Call sethortangle("", degrees, minutes, seconds)
        
    
    Case Else
    End Select

End If

cmdAcceptStation.Visible = True
cmdAcceptStation.Default = True
cmdAcceptStation.Enabled = True
cmdRecord(Index).Enabled = True

mdiMain.StatusBar.Panels(6).Visible = False

Exit Sub

ShowCoordinates:
Grid.Cols = 4
Grid.Rows = 3
Grid.ColWidth(0) = 300
Grid.TextMatrix(0, 1) = "Expected"
Grid.TextMatrix(0, 2) = "Recorded"
Grid.TextMatrix(0, 3) = "Difference"
Grid.TextMatrix(1, 0) = "X"
Grid.TextMatrix(2, 0) = "Y"
Grid.TextMatrix(1, 1) = Format(temp_stationx, "#######0.000")
Grid.TextMatrix(1, 2) = txtX(0)
Grid.TextMatrix(2, 1) = Format(temp_stationy, "########0.000")
Grid.TextMatrix(2, 2) = txtY(0)
Grid.TextMatrix(1, 3) = Format(txtX(0) - temp_stationx, "#######0.000")
Grid.TextMatrix(2, 3) = Format(txtY(0) - temp_stationy, "#######0.000")
txtstationheight = Format(txtZ(0) - temp_stationz, "#####0.000")
frmCoordinates.Visible = True
cmdAcceptStation.Visible = True
cmdAcceptStation.Default = True
cmdAcceptStation.Enabled = True
cmdRecord(Index).Enabled = True
Label6.Visible = True
lblST.Visible = True
lblST = Format(temp_stationz - txtZ(0), "#####0.000")
mdiMain.StatusBar.Panels(6).Visible = False

End Sub

Private Sub Command1_Click()

If mdiMain.StatusBar.Panels(7).Visible Then
    Cancelling = True
    Exit Sub
End If

If Shooting Then
    Exit Sub
End If

Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If mdiMain.StatusBar.Panels(7).Visible And KeyAscii = 27 Then
    Cancelling = True
End If

'If KeyAscii = 27 Then
'    Cancelling = True
'End If

End Sub

Private Sub Form_Load()

If DatumTB.RecordCount = 0 Then
    MsgBox ("You must define datums before initializing your station.  Use the Edit | Datums menu.")
    Exit Sub
End If

SetUpTypes(0).Value = True
setuphelp.Caption = "Select datum for Current Station. Enter Horizontal Angle and click on Set H-Angle button. Enter Station Height."
CenterForm Me
HangleSet = False

If DatumTB.BOF And DatumTB.EOF Then
    MsgBox ("No datums yet defined")
    Exit Sub
End If

For I = 0 To 2
    Station(I).Clear
    DatumTB.MoveFirst
    While Not DatumTB.EOF
        Station(I).AddItem DatumTB("Name")
        DatumTB.MoveNext
    Wend
    Gotit = False
    For J = 1 To Station(I).ListCount - 1
        Select Case I
        Case 0
            If LCase(Station(I).List(J)) = LCase(StationName) Then
                Station(I).ListIndex = J
                Gotit = True
                Exit For
            End If
            If Not Gotit Then
                Station(I) = "Select Datum"
            End If
        Case 1
            If LCase(Station(I).List(J)) = LCase(RefDatum1) Then
                Station(I).ListIndex = J
                Gotit = True
                Exit For
            End If
            If Not Gotit Then
                Station(I) = "Select Datum"
            End If
        Case 2
            If LCase(Station(I).List(J)) = LCase(RefDatum2) Then
                Station(I).ListIndex = J
                Gotit = True
                Exit For
            End If
            If Not Gotit Then
                Station(I) = "Select Datum"
            End If
        End Select
    Next J
Next I
SetUpTypes(SetupType) = True

End Sub

Private Sub setuptypes_Click(Index As Integer)

Select Case Index
Case 0
    setuphelp.Caption = "Select a datum for the Current Station from the pull-down list. Optionally enter a horizontal angle and click on Set Horizontal Angle button.  And optionally enter instrument height in meters."
Case 1
    setuphelp.Caption = "Select current Station datum and Primary Reference datum, set H-angle, then click on Record Datum button."
Case 2
    setuphelp.Caption = "Select current Station datum and Primary Reference datum and click on Record Datum Button."
Case 3
    setuphelp.Caption = "First select Primary and Secondary Reference datums.  Then click on Record Datum 1 button, then on Record Datum 2 button."
Case Else
End Select
setuphelp.Visible = True

cmdRecord(0).Visible = False

For I = 0 To 2
    If Station(I) = "" Then
        txtX(I) = ""
        txtY(I) = ""
        txtZ(I) = ""
    End If
Next I

'txtPrism.Visible = False
'lblPrism.Visible = False
frmCoordinates.Visible = False
cmdAcceptStation.Enabled = False
Select Case Index
    Case 0
        Frame1(0).Visible = False
        Frame1(1).Visible = False
        If Station(0) = "Unknown" Then Station(0) = "Select Datum"
        Station(0).Locked = False
        Station(0).BackColor = &H80000005
        If Station(0) <> "Select Datum" Then
            cmdAcceptStation.Enabled = True
        End If
        txtstationheight.Locked = False
        txtReferenceAngle.Enabled = True
        txtstationheight.BackColor = &H80000005
        txtReferenceAngle.BackColor = &H80000005
        lblPrism.Visible = False
        txtprism.Visible = False
    Case 1
        Station(0) = "Unknown"
        Station(0).Locked = True
        Station(0).BackColor = &HE0E0E0
        Frame1(0).Visible = True
        Frame1(1).Visible = False
        txtX(0) = ""
        txtY(0) = ""
        txtZ(0) = ""
        txtstationheight = "0"
        txtReferenceAngle = ""
        txtstationheight.Locked = True
        txtReferenceAngle.Enabled = True
        txtstationheight.BackColor = &HE0E0E0
        txtReferenceAngle.BackColor = &H80000005
    Case 2
        Frame1(0).Visible = True
        cmdHangle.Visible = False
        Frame1(1).Visible = False
        If Station(0) = "Unknown" Then Station(0) = "Select Datum"
        txtstationheight.Locked = True
        txtReferenceAngle.Enabled = False
        txtstationheight.BackColor = &HE0E0E0
        txtReferenceAngle.BackColor = &HE0E0E0
        txtstationheight = "0"
        txtReferenceAngle = ""
        Station(0).Locked = False
        Station(0).BackColor = &H80000005
        If Station(0) = "Select Datum" Then
            txtX(0) = ""
            txtY(0) = ""
            txtZ(0) = ""
        End If
        If Station(1) <> "" And Station(1) <> "Select Datum" Then
            Station_Click 1
        End If
        If Station(0) <> "Select Datum" Then
            cmdAcceptStation.Visible = True
        End If
            
    Case 3
        Frame1(0).Visible = True
        Frame1(1).Visible = True
        Station(0) = "Unknown"
        Station(0).Locked = True
        Station(0).BackColor = &HE0E0E0
        txtX(0) = ""
        txtY(0) = ""
        txtZ(0) = ""
        txtstationheight.Locked = True
        txtReferenceAngle.Enabled = False
        txtstationheight.BackColor = &HE0E0E0
        txtReferenceAngle.BackColor = &HE0E0E0
        txtstationheight = "0"
        txtReferenceAngle = ""
        If Station(1) <> "" And Station(1) <> "Select Datum" Then
            Station_Click 1
        End If
        If Station(2) <> "" And Station(2) <> "Select Datum" Then
            Station_Click 2
        End If
        
End Select

End Sub

Private Sub Station_Change(Index As Integer)

If Station(Index) = "" Then
    txtX(Index) = ""
    txtY(Index) = ""
    txtZ(Index) = ""
    If Index = 0 Then
        txtstationheight = "0"
        txtReferenceAngle = ""
    End If
End If

End Sub

Private Sub Station_Click(Index As Integer)

DatumTB.Index = "datumname"
DatumTB.Seek "=", Station(Index)
If Not DatumTB.NoMatch Then
    txtX(Index) = Format(DatumTB("x"), "#####0.000")
    txtY(Index) = Format(DatumTB("y"), "#####0.000")
    txtZ(Index) = Format(DatumTB("z"), "#####0.000")
End If

Select Case Index
    Case 0
        If SetUpTypes(0) Then
            cmdAcceptStation.Enabled = True
        ElseIf SetUpTypes(1) Then
            MsgBox ("For this type of setup, enter only the reference angle to the known Primary Site Datum")
            Station(0) = ""
            txtX(0) = ""
            txtY(0) = ""
            txtZ(0) = ""
            txtstationheight = "0"
        ElseIf SetUpTypes(2) Then
            If Station(1) <> "Select Datum" And Station(1) <> "" Then
                computeangle txtX(1), txtY(1), txtX(0), txtY(0), angle, minutes, seconds
                ma$ = Right$("00" + LTrim$(Str$(minutes)), 2)
                sa$ = Right$("00" + LTrim$(Str$(seconds)), 2)
                txtReferenceAngle = angle & "." & ma$ & sa$
                cmdRecord(0).Caption = "Record Datum"
                cmdRecord(0).Visible = True
                HangleSet = False
            End If
        ElseIf SetUpTypes(3) Then
            MsgBox ("For this type of setup do not choose a station name.")
            Station(0) = ""
            txtX(0) = ""
            txtY(0) = ""
            txtZ(0) = ""
            txtstationheight = "0"
            txtReferenceAngle = ""
        End If
        
    Case 1
        If SetUpTypes(2) Then
            If Station(0) <> "Select Datum" Then
                computeangle txtX(1), txtY(1), txtX(0), txtY(0), angle, minutes, seconds
                ma$ = Right$("00" + LTrim$(Str$(minutes)), 2)
                sa$ = Right$("00" + LTrim$(Str$(seconds)), 2)
                txtReferenceAngle = angle & "." & ma$ & sa$
                cmdRecord(0).Caption = "Record Datum"
                cmdRecord(0).Visible = True
                        
            End If
        ElseIf SetUpTypes(1) Then
            cmdRecord(0).Caption = "Record Datum"
            cmdRecord(0).Visible = True
        ElseIf SetUpTypes(3) Then
            If Station(2) <> "" Then
                cmdRecord(0).Caption = "Record Datum 1"
                cmdRecord(0).Visible = True
            End If
        End If
        RefDatum1 = Station(Index)
    Case 2
        If Station(1) <> "" Then
            cmdRecord(0).Caption = "Record Datum 1"
            cmdRecord(0).Visible = True
        End If
        RefDatum2 = Station(Index)
End Select
                
End Sub

Private Sub txtprism_Click()

If SetUpTypes(1) Then
    txtstationheight = edmshot.z - PoleHeight(txtprism.ItemData(txtprism.ListIndex))
Else
    txtZ(0) = edmshot.z - PoleHeight(txtprism.ItemData(txtprism.ListIndex))
    txtstationheight = 0
End If

End Sub

Private Sub txtPrism_DropDown()

txtprism.Clear
For I = 0 To frmMain.txtprism.ListCount - 1
        txtprism.AddItem frmMain.txtprism.List(I)
        txtprism.ItemData(txtprism.NewIndex) = frmMain.txtprism.ItemData(I)
Next I

End Sub

Private Sub txtReferenceAngle_Change()

If txtReferenceAngle <> "" Then
    cmdHangle.Visible = True
    cmdHangle.Default = True
    If SetUpTypes(0) Or SetUpTypes(1) Then
        calccomp.Visible = True
    End If
Else
    cmdHangle.Visible = False
    calccomp.Visible = False
End If

End Sub

Private Sub txtReferenceAngle_GotFocus()

txtReferenceAngle.SelStart = 0
txtReferenceAngle.SelLength = Len(txtReferenceAngle)

End Sub

Private Sub txtReferenceAngle_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 8, 46, 48 To 57
    Case Else
        KeyAscii = 0
End Select

End Sub

Private Sub txtReferenceAngle_LostFocus()

If Not IsNumeric(txtReferenceAngle) Or Val(txtReferenceAngle) < 0 Or Val(txtReferenceAngle) > 360 Then
    MsgBox ("Enter value between 0 and 360 degrees (ddd.mmss) for Angle to Reference Datum")
    Exit Sub
End If
txtReferenceAngle = Format(Val(txtReferenceAngle), "####0.0000")

End Sub

Private Sub txtStationHeight_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 8, 46, 48 To 57, Asc(".")
    Case Else
        KeyAscii = 0
End Select

End Sub

Private Sub txtStationHeight_LostFocus()

If Not IsNumeric(txtstationheight) Then
    MsgBox ("Enter Station Height as numeric value")
End If

End Sub

Public Sub CalcAngle(x1 As Double, y1 As Double, x2 As Double, y2 As Double, tangle As Double)

Dim rrun As Double, rise As Double
Dim slope As Double, pi As Double

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

