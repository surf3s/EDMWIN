VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInitialize 
   Caption         =   "Station Initialize"
   ClientHeight    =   8832
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   12192
   LinkTopic       =   "Form1"
   ScaleHeight     =   8832
   ScaleWidth      =   12192
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   4305
      Left            =   4980
      TabIndex        =   32
      Top             =   6210
      Width           =   5595
      _ExtentX        =   9864
      _ExtentY        =   7599
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmInitialize.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "panels(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmInitialize.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmInitialize.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.PictureBox panels 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3372
         Index           =   0
         Left            =   180
         ScaleHeight     =   3348
         ScaleWidth      =   5148
         TabIndex        =   33
         Top             =   720
         Width           =   5172
         Begin VB.Frame Frame1 
            Caption         =   "Setup Type"
            Height          =   1815
            Left            =   120
            TabIndex        =   34
            Top             =   120
            Width           =   3135
            Begin VB.OptionButton setuptypes 
               Caption         =   "On datum (h-angle already set)."
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   38
               Top             =   270
               Value           =   -1  'True
               Width           =   2775
            End
            Begin VB.OptionButton setuptypes 
               Caption         =   "Sight to datum (h-angle already set)."
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   37
               Top             =   600
               Width           =   2895
            End
            Begin VB.OptionButton setuptypes 
               Caption         =   "Sight to datum and on datum."
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   36
               Top             =   960
               Width           =   2655
            End
            Begin VB.OptionButton setuptypes 
               Caption         =   "Sight to two datums."
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   35
               Top             =   1320
               Width           =   2175
            End
         End
         Begin VB.Label setupdesc 
            BackStyle       =   0  'Transparent
            Height          =   1215
            Left            =   150
            TabIndex        =   39
            Top             =   1890
            Width           =   3135
         End
      End
   End
   Begin VB.PictureBox panels 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3852
      Index           =   5
      Left            =   -1020
      ScaleHeight     =   3828
      ScaleWidth      =   5148
      TabIndex        =   20
      Top             =   4830
      Width           =   5172
      Begin VB.Label instructions5 
         BackStyle       =   0  'Transparent
         Caption         =   "Instructions"
         Height          =   2052
         Left            =   480
         TabIndex        =   21
         Top             =   240
         Width           =   2172
      End
   End
   Begin VB.PictureBox panels 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3852
      Index           =   1
      Left            =   1890
      ScaleHeight     =   3828
      ScaleWidth      =   5148
      TabIndex        =   4
      Top             =   3330
      Width           =   5172
      Begin VB.TextBox stationheight 
         Height          =   285
         Left            =   3120
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSDBGrid.DBGrid datumlist1 
         Bindings        =   "frmInitialize.frx":0054
         Height          =   2292
         Left            =   120
         OleObjectBlob   =   "frmInitialize.frx":006D
         TabIndex        =   5
         Top             =   1080
         Width           =   5052
      End
      Begin VB.Label stationheightlbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the station height over the datum : "
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label datum1desc 
         BackStyle       =   0  'Transparent
         Caption         =   "Select "
         Height          =   255
         Left            =   60
         TabIndex        =   6
         Top             =   180
         Width           =   4935
      End
   End
   Begin VB.PictureBox panels 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   6
      Left            =   3840
      ScaleHeight     =   4188
      ScaleWidth      =   6768
      TabIndex        =   22
      Top             =   2190
      Width           =   6795
      Begin VB.CommandButton Command2 
         Caption         =   "Accept current station"
         Height          =   372
         Left            =   960
         TabIndex        =   29
         Top             =   2880
         Width           =   2172
      End
      Begin VB.Label title2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Station Initialization Parameters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   3252
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "angle to reference point"
         Height          =   195
         Index           =   9
         Left            =   480
         TabIndex        =   27
         Top             =   2400
         Width           =   1680
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "reference points coordinates if known"
         Height          =   195
         Index           =   8
         Left            =   480
         TabIndex        =   26
         Top             =   2040
         Width           =   2640
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference Point :"
         Height          =   195
         Index           =   7
         Left            =   480
         TabIndex        =   25
         Top             =   1680
         Width           =   1230
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "current station coordinates if known"
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   24
         Top             =   1200
         Width           =   2475
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Station :"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   23
         Top             =   840
         Width           =   1110
      End
   End
   Begin VB.PictureBox panels 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3852
      Index           =   3
      Left            =   3270
      ScaleHeight     =   3828
      ScaleWidth      =   5148
      TabIndex        =   10
      Top             =   3540
      Width           =   5172
      Begin VB.Label title 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Station Initialization Parameters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   3252
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "angle to reference point"
         Height          =   192
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   2880
         Width           =   1680
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "reference points coordinates if known"
         Height          =   192
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Width           =   2640
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference Point :"
         Height          =   192
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   1224
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "current station coordinates if known"
         Height          =   192
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   2472
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Station :"
         Height          =   192
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   1104
      End
   End
   Begin VB.PictureBox panels 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3852
      Index           =   4
      Left            =   5610
      ScaleHeight     =   3828
      ScaleWidth      =   5148
      TabIndex        =   17
      Top             =   3990
      Width           =   5172
      Begin VB.CommandButton setangle 
         Caption         =   "Set Horizontal Angle"
         Height          =   372
         Left            =   1080
         TabIndex        =   19
         Top             =   2880
         Width           =   2172
      End
      Begin VB.Label instructions1 
         BackStyle       =   0  'Transparent
         Caption         =   "Instructions"
         Height          =   2055
         Left            =   2160
         TabIndex        =   18
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox panels 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3852
      Index           =   2
      Left            =   5760
      ScaleHeight     =   3828
      ScaleWidth      =   5148
      TabIndex        =   7
      Top             =   1050
      Width           =   5172
      Begin MSDBGrid.DBGrid datumlist2 
         Bindings        =   "frmInitialize.frx":0A43
         Height          =   2412
         Left            =   120
         OleObjectBlob   =   "frmInitialize.frx":0A5C
         TabIndex        =   8
         Top             =   1200
         Width           =   5052
      End
      Begin VB.Label datum2desc 
         Caption         =   "Select "
         Height          =   492
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   5052
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Finish"
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   3
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Next >>"
      Default         =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   2
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<< &Back"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   3600
      Width           =   855
   End
   Begin VB.Data datum2data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   570
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4590
      Visible         =   0   'False
      Width           =   3132
   End
   Begin VB.Data datum1data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   510
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5070
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   5760
      Y1              =   3480
      Y2              =   3480
   End
End
Attribute VB_Name = "frmInitialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim panelno As Integer
Dim datum1 As xyz
Dim datum2 As xyz
Dim tempstation As xyz

Private Sub Command1_Click(Index As Integer)

Select Case Index
Case 0
    Unload Me

Case 1
    If panelno = 3 Then
        If SetUpTypes(0) Or SetUpTypes(1) Then
            panelno = 1
        Else
            panelno = 2
        End If
    ElseIf panelno = 5 Then
        If SetUpTypes(1) Then
            panelno = 1
        Else
            If panelno <> 0 Then panelno = panelno - 1
        End If
    ElseIf panelno = 6 Then
        If SetUpTypes(0) Or SetUpTypes(1) Then
            panelno = 1
        Else
            panelno = 5
        End If
    Else
        If panelno <> 0 Then panelno = panelno - 1
    End If

Case 2
    If panelno = 1 Then
        If SetUpTypes(0) Then
            panelno = 6
        ElseIf SetUpTypes(1) Then
            panelno = 5
        Else
            panelno = 2
        End If
    ElseIf panelno = 5 Then
        Select Case EDMName$
        Case "None"
            frmManualshot.Show 1
        Case Else
            frmTakeshot.Show 1
        End Select
    Else
        If panelno <> 6 Then panelno = panelno + 1
    End If

Case 3
    panelno = 6

Case Else
End Select

If panelno = 3 Then
    If datum1.Name <> "" And datum2.Name <> "" Then
        Call computeangle(datum1.X, datum1.y, datum2.X, datum2.y, angle, minutes, seconds)
        Label1(4).Caption = "Angle between points is " + Str$(angle) + "." + Str$(minutes) + "." + Str$(seconds)
    End If
End If

For A = 0 To 6
    If A <> panelno Then
        panels(A).Visible = False
    Else
        panels(A).Visible = True
    End If
Next A

Select Case panelno
Case 5
    instructions5.Left = Me.Width / 2 - instructions5.Width / 2
    Command1(2).Enabled = True
    Command1(3).Enabled = True
    Command1(2).Default = True
    
Case 6
    If SetUpTypes(0) Then
        Label1(7).Visible = False
        Label1(8).Visible = False
        Label1(9).Visible = False
    End If
    Command2.Left = Me.Width / 2 - Command2.Width / 2
    Command2.Default = True
    Command1(2).Enabled = False
    Command1(3).Enabled = False
    Label1(6).Caption = Trim(tempstation.Name) + " -> X = " + Trim(Str$(tempstation.X)) + ", Y = " + Trim(Str$(tempstation.y)) + ", Z = " + Trim(Str$(tempstation.z))
    
Case Else
    Command1(2).Enabled = True
    Command1(3).Enabled = True
    Command1(2).Default = True
End Select

End Sub

Private Sub Command2_Click()

'all points will now be offset to this currentstation
CurrentStation.X = tempstation.X
CurrentStation.y = tempstation.y
CurrentStation.z = tempstation.z
StationName = tempstation.Name
StationInitialized = True
'Update the cfg file so autoresume will work



Unload Me

End Sub

Private Sub datum1data_Reposition()

If Not datum1data.Recordset.EOF And Not datum1data.Recordset.BOF Then
    
    If Not IsNull(datum1data.Recordset("Datum")) And Not IsNull(datum1data.Recordset("x")) And datum1data.Recordset("y") And datum1data.Recordset("z") Then
        Label1(1).Caption = datum1data.Recordset("Datum") + " - X = " + Str$(datum1data.Recordset("x")) + ", Y = " + Str$(datum1data.Recordset("y")) + ", Z = " + Str$(datum1data.Recordset("z"))
        datum1.Name = datum1data.Recordset("Datum")
        datum1.X = datum1data.Recordset("x")
        datum1.y = datum1data.Recordset("y")
        datum1.z = datum1data.Recordset("z")
        
        If SetUpTypes(0) Then
            tempstation.Name = datum1.Name
            tempstation.X = datum1.X
            tempstation.y = datum1.y
            tempstation.z = datum1.z + Val(stationheight.Text)
        ElseIf SetUpTypes(1) Then
            instructions5.Caption = "Aim the theodolite at the prism position on " + Trim$(datum1.Name) + " datum and then click on the 'Record Reference Point' button below to take the shot."
        End If
    
    Else
        Label1(1).Caption = "No valid datum selected."
        datum1.Name = ""
    
    End If

End If

End Sub

Private Sub datumlist2_Click()

If Not datum2data.Recordset.EOF And Not datum2data.Recordset.BOF Then
    If Not IsNull(datum2data.Recordset("Datum")) And Not IsNull(datum2data.Recordset("x")) And datum2data.Recordset("y") And datum2data.Recordset("z") Then
        Label1(8).Caption = datum2data.Recordset("Datum") + " - X = " + Str$(datum2data.Recordset("x")) + ", Y = " + Str$(datum2data.Recordset("y")) + ", Z = " + Str$(datum2data.Recordset("z"))
        Label1(3).Caption = datum2data.Recordset("Datum") + " - X = " + Str$(datum2data.Recordset("x")) + ", Y = " + Str$(datum2data.Recordset("y")) + ", Z = " + Str$(datum2data.Recordset("z"))
        datum2.Name = datum2data.Recordset("Datum")
        datum2.X = datum2data.Recordset("x")
        datum2.y = datum2data.Recordset("y")
        datum2.z = datum2data.Recordset("z")
    Else
        Label1(3).Caption = "No valid datum selected."
        datum2.Name = ""
    End If
End If

End Sub

Private Sub Form_Load()

Call centerform(Me)

panelno = 0

For A = 0 To 6
    panels(A).Visible = False
    panels(A).Width = Me.Width - 200
    panels(A).Top = 50
    panels(A).Height = Line1.y1 - panels(A).Top - 50
    panels(A).Left = 0
    panels(A).BorderStyle = 0
Next A

panels(0).Visible = True

Frame1.Top = 60
setupdesc.Top = Frame1.Top + Frame1.Height + 60
Frame1.Left = 100
setupdesc.Left = Frame1.Left


datum1desc.Top = 130
stationheightlbl.Top = datum1desc.Top + datum1desc.Height + 60
stationheight.Top = stationheightlbl.Top
datumlist1.Top = stationheightlbl.Top + stationheightlbl.Height + 100
datum2desc.Top = 130
datumlist2.Top = datum2desc.Top + datum2desc.Height + 60

datum1desc.Left = Frame1.Left
datumlist1.Left = Frame1.Left
datum2desc.Left = Frame1.Left
datumlist2.Left = Frame1.Left

title.Left = Me.Width / 2 - title.Width / 2
title2.Left = Me.Width / 2 - title2.Width / 2
title.Top = 60
title2.Top = 60

Label1(0).Left = Frame1.Left
Label1(1).Left = Frame1.Left
Label1(2).Left = Frame1.Left
Label1(3).Left = Frame1.Left
Label1(4).Left = Frame1.Left
Label1(5).Left = Frame1.Left
Label1(6).Left = Frame1.Left
Label1(7).Left = Frame1.Left
Label1(8).Left = Frame1.Left
Label1(9).Left = Frame1.Left

stationheightlbl.Left = Frame1.Left

datumlist1.Height = Line1.y1 - 350 - datumlist1.Top
datumlist2.Height = Line1.y1 - 350 - datumlist1.Top
datumlist1.Width = Line1.x2 - datumlist1.Left
datumlist2.Width = Line1.x2 - datumlist2.Left

Set datum1data.Recordset = DatumTB
Set datum2data.Recordset = DatumTB

Call setuptypes_Click(0)

Command1(2).Default = True

End Sub

Private Sub setuptypes_Click(Index As Integer)

Select Case Index
Case 0
    setupdesc.Caption = "This type of setup requires only that you are over a point listed in the datum file, that you can measure the height of the instrument over this point, and that the horizontal angle has already been set."
    datum1desc.Caption = "Select or enter the datum, and enter the station height?"
    datum2desc.Caption = ""
    stationheightlbl.Visible = True
    stationheight.Visible = True
    
Case 1
    setupdesc.Caption = "With this type of setup you are not over a known point, but the horizontal angle is correct and you can sight to a known point."
    datum1desc.Caption = "Sight to which datum point?"
    datum2desc.Caption = ""
    stationheightlbl.Visible = False
    stationheight.Visible = False

Case 2
    setupdesc.Caption = "This type of setup requires that you are over a known point and you can shoot another known point.  The horizontal angle and height of the station will be calculated automatically."
    datum1desc.Caption = "The theodolite is setup over which datum?"
    datum2desc.Caption = "Sight to which datum point?"
    stationheightlbl.Visible = False
    stationheight.Visible = False
    
Case 3
    setupdesc.Caption = "This type of setup requires that you can shoot to two known points.  The horizontal angle and the station height will be calculated automatically."
    datum1desc.Caption = "Sight to which datum point?"
    datum2desc.Caption = "Sight to which other datum point?"
    stationheightlbl.Visible = False
    stationheight.Visible = False
    
Case Else
End Select

End Sub

Private Sub sizecontols()

DBGrid1.Width = SSTab1.Width * 0.9
DBGrid1.Left = SSTab1.Width / 2 - DBGrid1.Width / 2

End Sub

Private Sub stationheight_Change()

tempstation.z = datum1.z + Val(stationheight.Text)

End Sub

