VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTheodolite 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Total Station"
   ClientHeight    =   5115
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   11280
   ControlBox      =   0   'False
   Icon            =   "frmTheodolite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnFindPorts 
      Caption         =   "Look for Ports"
      Height          =   615
      Left            =   360
      TabIndex        =   40
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   2400
      TabIndex        =   27
      Top             =   3840
      Width           =   8700
      Begin VB.TextBox Text1 
         Height          =   288
         Index           =   2
         Left            =   5430
         TabIndex        =   31
         Top             =   672
         Width           =   972
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Index           =   1
         Left            =   3870
         TabIndex        =   30
         Top             =   672
         Width           =   972
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Index           =   0
         Left            =   2310
         TabIndex        =   29
         Top             =   672
         Width           =   972
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Z:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   5190
         TabIndex        =   34
         Top             =   690
         Width           =   165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3630
         TabIndex        =   33
         Top             =   690
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2070
         TabIndex        =   32
         Top             =   690
         Width           =   165
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   $"frmTheodolite.frx":000C
         Height          =   495
         Left            =   1275
         TabIndex        =   28
         Top             =   150
         Width           =   6345
      End
   End
   Begin VB.Frame Frame1 
      Height          =   972
      Left            =   2400
      TabIndex        =   14
      Top             =   2760
      Width           =   8700
      Begin VB.CheckBox chkGeocom 
         Caption         =   "GeoCOM"
         Height          =   255
         Left            =   6480
         TabIndex        =   39
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Try"
         Default         =   -1  'True
         Height          =   495
         Index           =   2
         Left            =   7680
         TabIndex        =   36
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox baudrate 
         Height          =   315
         ItemData        =   "frmTheodolite.frx":00B7
         Left            =   1320
         List            =   "frmTheodolite.frx":00CA
         TabIndex        =   20
         Text            =   "1200"
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox databits 
         Height          =   315
         ItemData        =   "frmTheodolite.frx":00EB
         Left            =   2400
         List            =   "frmTheodolite.frx":00F5
         TabIndex        =   19
         Text            =   "7"
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox stopbits 
         Height          =   315
         ItemData        =   "frmTheodolite.frx":00FF
         Left            =   3360
         List            =   "frmTheodolite.frx":0109
         TabIndex        =   18
         Text            =   "1"
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox Parity 
         Height          =   315
         ItemData        =   "frmTheodolite.frx":0113
         Left            =   4320
         List            =   "frmTheodolite.frx":0120
         TabIndex        =   17
         Text            =   "Even"
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox txtcomport 
         Height          =   315
         ItemData        =   "frmTheodolite.frx":0135
         Left            =   240
         List            =   "frmTheodolite.frx":015D
         TabIndex        =   16
         Text            =   "COM1"
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox DelayTime 
         Height          =   315
         ItemData        =   "frmTheodolite.frx":01AC
         Left            =   5355
         List            =   "frmTheodolite.frx":01C5
         TabIndex        =   15
         Text            =   "3"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Baud rate :"
         Height          =   195
         Left            =   1320
         TabIndex        =   26
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Bits :"
         Height          =   195
         Left            =   2400
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stop Bits :"
         Height          =   195
         Left            =   3360
         TabIndex        =   24
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parity :"
         Height          =   195
         Left            =   4320
         TabIndex        =   23
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COM Port :"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delay Time:"
         Height          =   195
         Left            =   5355
         TabIndex        =   21
         Top             =   240
         Width           =   840
      End
   End
   Begin TabDlg.SSTab whichedm 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   9
      Tab             =   1
      TabsPerRow      =   9
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Topcon"
      TabPicture(0)   =   "frmTheodolite.frx":01E4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2(1)"
      Tab(0).Control(1)=   "Label1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Leica/Wild"
      TabPicture(1)   =   "frmTheodolite.frx":0200
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label18"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label19"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Sokkia"
      TabPicture(2)   =   "frmTheodolite.frx":021C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "None"
      TabPicture(3)   =   "frmTheodolite.frx":0238
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Option1(1)"
      Tab(3).Control(1)=   "Option1(0)"
      Tab(3).Control(2)=   "Label5"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Simulate"
      TabPicture(4)   =   "frmTheodolite.frx":0254
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label11"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Nikon"
      TabPicture(5)   =   "frmTheodolite.frx":0270
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label13"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Leica-Builder"
      TabPicture(6)   =   "frmTheodolite.frx":028C
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label14"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Microscribe"
      TabPicture(7)   =   "frmTheodolite.frx":02A8
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label15"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Pentograph"
      TabPicture(8)   =   "frmTheodolite.frx":02C4
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label17"
      Tab(8).ControlCount=   1
      Begin VB.OptionButton Option1 
         Caption         =   "X, Y and Z coordinates"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   8
         Top             =   1440
         Width           =   3495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Vertical angle, Horizontal Angle and Distance"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   7
         Top             =   1200
         Value           =   -1  'True
         Width           =   3615
      End
      Begin VB.Label Label19 
         Caption         =   "Newer instruments (like FlexLines) use a system called GEOCOM.  Check the GEOCOM box below to use this protocol."
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   1800
         Width           =   10545
      End
      Begin VB.Label Label18 
         Caption         =   $"frmTheodolite.frx":02E0
         Height          =   495
         Left            =   240
         TabIndex        =   37
         Top             =   1200
         Width           =   10545
      End
      Begin VB.Label Label17 
         Caption         =   $"frmTheodolite.frx":03DA
         Height          =   615
         Left            =   -74760
         TabIndex        =   35
         Top             =   480
         Width           =   10425
      End
      Begin VB.Label Label15 
         Caption         =   $"frmTheodolite.frx":0497
         Height          =   495
         Left            =   -74760
         TabIndex        =   13
         Top             =   480
         Width           =   10425
      End
      Begin VB.Label Label14 
         Caption         =   $"frmTheodolite.frx":0588
         Height          =   615
         Left            =   -74640
         TabIndex        =   12
         Top             =   600
         Width           =   10305
      End
      Begin VB.Label Label13 
         Caption         =   $"frmTheodolite.frx":0614
         Height          =   615
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   5385
      End
      Begin VB.Label Label11 
         Caption         =   $"frmTheodolite.frx":06DE
         Height          =   1215
         Left            =   -74760
         TabIndex        =   10
         Top             =   480
         Width           =   5505
      End
      Begin VB.Label Label4 
         Caption         =   $"frmTheodolite.frx":0796
         Height          =   1215
         Left            =   -74760
         TabIndex        =   9
         Top             =   480
         Width           =   5475
      End
      Begin VB.Label Label5 
         Caption         =   $"frmTheodolite.frx":0899
         Height          =   855
         Left            =   -74760
         TabIndex        =   6
         Top             =   480
         Width           =   5505
      End
      Begin VB.Label Label3 
         Caption         =   $"frmTheodolite.frx":095F
         Height          =   735
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   10545
      End
      Begin VB.Label Label2 
         Caption         =   $"frmTheodolite.frx":0A9D
         Height          =   615
         Index           =   1
         Left            =   -74760
         TabIndex        =   4
         Top             =   960
         Width           =   5475
      End
      Begin VB.Label Label1 
         Caption         =   "EDM has been tested with Topcon GTS-3, GTS 210, GTS-220, and GTS-320 series total stations."
         Height          =   495
         Left            =   -74760
         TabIndex        =   3
         Top             =   480
         Width           =   5505
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Accept"
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
   End
End
Attribute VB_Name = "frmTheodolite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function get_settings()

Select Case UCase$(Parity.Text)
Case "EVEN"
    get_settings = baudrate.Text + ",E," + databits.Text + "," + stopbits.Text
Case "ODD"
    get_settings = baudrate.Text + ",O," + databits.Text + "," + stopbits.Text
Case "NONE"
    get_settings = baudrate.Text + ",N," + databits.Text + "," + stopbits.Text
Case Else
End Select
    
End Function

Private Sub btnFindPorts_Click()

answer = MsgBox("EDMWIN will look for COM ports that appear to be available.  It will check ports 1-12 and report back here.  This can take a minute.  It will also change your current settings.  Are you sure you want to check for ports?", vbOKCancel)
If answer = 1 Then
    Screen.MousePointer = 11
    ports$ = ""
    If frmMain.theoport.PortOpen Then frmMain.theoport.PortOpen = False
    For A = 1 To 12
        frmMain.theoport.Settings = "1200,E,7,1"
        frmMain.theoport.CommPort = A
        On Error Resume Next
        frmMain.theoport.PortOpen = True
        If frmMain.theoport.PortOpen Then
            If ports$ = "" Then
                ports$ = "COM" + Trim(Str(A))
            Else
                ports$ = ports$ + ", COM" + Trim(Str(A))
            End If
        End If
        frmMain.theoport.PortOpen = False
        On Error GoTo 0
    Next A
    Screen.MousePointer = 1
    answer = MsgBox("The following ports are available for EDMWIN: " + ports$, vbInformation + vbOKOnly)
End If

End Sub

Private Sub Command1_Click(Index As Integer)

If Index = 1 Then
    Unload Me
End If

If whichedm.Tab = 7 Then
    UsingMicroscribe = True

ElseIf whichedm.Tab <> 3 Then
    UsingMicroscribe = False
    If txtcomport.Text = "" Or baudrate.Text = "" Or databits.Text = "" Or stopbits.Text = "" Or Parity.Text = "" Then
        MsgBox "Select the communications parameters: COMport, baudrate, databits, stopbits, and parity. You will need to match these to the settings in the instrument.  Consult the instrument's manual if you do not know them. If you don't have an total station connected, then select 'None'", vbInformation
        Exit Sub
    End If
    comport = txtcomport.Text
    comsettings = get_settings()
End If

Screen.MousePointer = 11

frmMain.lblEDMWarning.Visible = False

Do
    Cancelling = False
    Select Case Index
    Case 0, 2
        Select Case whichedm.Tab
        Case 0
            EDMName$ = "Topcon"
            answer = MsgBox("Cable the total station to the computer and communications will be initialized.", vbOKCancel)
            If answer = 1 Then
                Screen.MousePointer = 11
                Call initcomport(comport, errorcode)
                If Cancelling Then
                    MsgBox ("Communications error with total station.  Verify that it is turned on.")
                End If
            End If
        
        Case 1
            EDMName$ = "Wild"
            If chkGeocom.Value Then
                EDMName$ = "Wild2"
            End If
            answer = MsgBox("Cable the total station to the computer and communications will be initialized.", vbOKCancel)
            If answer = 1 Then
                Screen.MousePointer = 11
                Call initcomport(comport, errorcode)
                If Cancelling Then
                    MsgBox ("Communications error with total station.  Verify that it is turned on.")
                End If
            End If
        
        Case 2
            EDMName$ = "Sokkia"
            answer = MsgBox("Cable the total station to the computer and communications will be initialized.", vbOKCancel)
            If answer = 1 Then
                Screen.MousePointer = 11
                Call initcomport(comport, errorcode)
                If Cancelling Then
                    MsgBox ("Communications error with total station.  Verify that it is turned on.")
                End If
            End If
        
        Case 5
            EDMName$ = "Nikon"
            answer = MsgBox("Cable the total station to the computer and communications will be initialized.", vbOKCancel)
            If answer = 1 Then
                Screen.MousePointer = 11
                Call initcomport(comport, errorcode)
                If Cancelling Then
                    MsgBox ("Communications error with total station.  Verify that it is turned on.")
                End If
            End If
        
        Case 3
            EDMName$ = "None"
            If Option1(0).Value Then
                vhd = True
            Else
                vhd = False
            End If
        
        Case 4
            EDMName$ = "Simulate"
        
        Case 6
            EDMName$ = "builder"
            answer = MsgBox("Cable the total station to the computer and communications will be initialized.", vbOKCancel)
            If answer = 1 Then
                Screen.MousePointer = 11
                Call initcomport(comport, errorcode)
                If Cancelling Then
                    MsgBox ("Communications error with total station.  Verify that it is turned on.")
                End If
            End If
        
        Case 7
            EDMName$ = "Microscribe"
            
        Case 8
            EDMName$ = "Pentograph"
            answer = MsgBox("Make sure the pentograph software is running and communications will be initialized.", vbOKCancel)
            If answer = 1 Then
                Screen.MousePointer = 11
                Call initcomport(comport, errorcode)
                If Cancelling Then
                    MsgBox ("Communications error with the pentograph.  Verify that the interface software is active and configured correctly.")
                End If
            End If
           
        End Select
    End Select

Loop Until Not Cancelling

Screen.MousePointer = 1

If Index = 0 Then
    Dim IniClass As String
    Dim Inidata(5, 2) As String
    Dim Status As Byte
    
    IniClass = "[EDM]"
    Inidata(1, 1) = "Instrument"
    Inidata(1, 2) = EDMName
    Inidata(2, 1) = "COMport"
    Inidata(2, 2) = comport
    Inidata(3, 1) = "Settings"
    Inidata(3, 2) = comsettings
    Inidata(4, 1) = "VHD"
    Inidata(5, 1) = "EdmDelayTime"
    Inidata(5, 2) = EDMDelayTime
    
    If vhd Then Inidata(4, 2) = "True" Else Inidata(4, 2) = "False"
    
    Call WriteIni(CFGName, IniClass, Inidata(), Status)
    
    Unload Me
End If

End Sub

Private Sub DelayTime_Click()

EDMDelayTime = DelayTime.Text

End Sub

Private Sub Form_Load()

Screen.MousePointer = 1
CenterForm Me
If comport <> "" Then
    For I = 0 To txtcomport.ListCount - 1
        If LCase(txtcomport.List(I)) = LCase(comport) Then
            txtcomport.ListIndex = I
            Exit For
        End If
    Next I
End If

If EDMDelayTime <> 0 Then
    For I = 0 To DelayTime.ListCount - 1
        If LCase(DelayTime.List(I)) = LCase(EDMDelayTime) Then
            DelayTime.ListIndex = I
            Exit For
        End If
    Next I
End If

TempString = comsettings
X = InStr(TempString, ",")
If X > 0 Then
    baudrate.Text = Left(TempString, X - 1)
    TempString = Mid(TempString, X + 1)
    X = InStr(TempString, ",")
    If X > 0 Then
        Select Case UCase(Left(TempString, X - 1))
            Case "E"
                Parity = "Even"
            Case "N"
                Parity = "None"
            Case "0"
                Parity = "Odd"
        End Select
        TempString = Mid(TempString, X + 1)
        X = InStr(TempString, ",")
        If X > 0 Then
            databits.Text = Left(TempString, X - 1)
            TempString = Mid(TempString, X + 1)
            X = InStr(TempString, ",")
            If X > 0 Then
                stopbits.Text = Left(TempString, X - 1)
                TempString = Mid(TempString, X + 1)
            Else
                stopbits.Text = TempString
            End If
        End If
    End If
End If
        
comsettings = baudrate.Text + "," + Left(Parity, 1) + "," + databits.Text + "," + stopbits.Text

Frame1.Visible = True
Frame2.Visible = False
Select Case UCase(EDMName)
    Case "TOPCON"
        whichedm.Tab = 0
    Case "WILD", "LEICA"
        whichedm.Tab = 1
    Case "WILD2"
        whichedm.Tab = 1
        chkGeocom.Value = 1
    Case "SOKKIA"
        whichedm.Tab = 2
    Case "NONE"
        whichedm.Tab = 3
    Case "SIMULATE"
        whichedm.Tab = 4
    Case "NIKON"
        whichedm.Tab = 5
    Case "BUILDER"
        whichedm.Tab = 6
    Case "MICROSCRIBE"
        whichedm.Tab = 7
        Frame1.Visible = False
        Frame2.Visible = True
    Case "PENTOGRAPH"
        whichedm.Tab = 8
End Select

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

Select Case KeyAscii
    Case 13
        KeyAscii = 0
        Exit Sub
    Case 44, 46, 48 To 57, Asc("-"), Asc(".")
    Case Else
        KeyAscii = 0
        MsgBox ("Invalid data received from Microscribe/Pentograph")
End Select

If KeyAscii = 44 Then
    KeyAscii = 0
    If Index < 2 Then
        Text1(Index + 1).SetFocus
    End If
End If

End Sub

Private Sub txtcomport_Change()

comport = txtcomport

End Sub

Private Sub whichedm_Click(PreviousTab As Integer)

Select Case whichedm.Tab
Case 0, 1, 2, 5, 6
    Frame2.Visible = False
    Frame1.Visible = True
    baudrate.Enabled = True
    Parity.Enabled = True
    databits.Enabled = True
    stopbits.Enabled = True

Case 3, 4
    Frame2.Visible = False
    Frame1.Visible = False
    baudrate.Enabled = False
    Parity.Enabled = False
    databits.Enabled = False
    stopbits.Enabled = False

Case 7
    Frame2.Visible = True
    Frame1.Visible = False
    Me.Show
    Text1(0).SetFocus

Case 8
    Frame2.Visible = True
    Frame1.Visible = True
    Me.Show
    Text1(0).SetFocus

Case Else
    
End Select

End Sub

