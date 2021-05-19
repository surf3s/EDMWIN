VERSION 5.00
Begin VB.Form frmMSSetup 
   Caption         =   "Microscribe Station Initialization"
   ClientHeight    =   2895
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6150
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Left            =   4704
      TabIndex        =   11
      Top             =   624
      Width           =   972
   End
   Begin VB.TextBox txtName 
      Height          =   288
      Left            =   1824
      TabIndex        =   7
      Top             =   96
      Width           =   2460
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Record Origin Point"
      Height          =   348
      Left            =   2400
      TabIndex        =   6
      Top             =   1152
      Width           =   1596
   End
   Begin VB.TextBox txtZ 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   3552
      TabIndex        =   5
      Text            =   "0"
      Top             =   720
      Width           =   732
   End
   Begin VB.TextBox txtY 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   2712
      TabIndex        =   4
      Text            =   "0"
      Top             =   720
      Width           =   732
   End
   Begin VB.TextBox txtX 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   1824
      TabIndex        =   3
      Text            =   "0"
      Top             =   720
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   300
      Left            =   4704
      TabIndex        =   0
      Top             =   144
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   $"frmMSSetup.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   924
      Left            =   336
      TabIndex        =   12
      Top             =   1776
      Width           =   5388
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   3840
      TabIndex        =   10
      Top             =   432
      Width           =   144
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2928
      TabIndex        =   9
      Top             =   432
      Width           =   156
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   2112
      TabIndex        =   8
      Top             =   432
      Width           =   156
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Station Coordinates:"
      Height          =   192
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Block Name/Number:"
      Height          =   192
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1548
   End
End
Attribute VB_Name = "frmMSSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempCurrentStationX As Single
Dim TempCurrentStationy As Single
Dim TempCurrentStationZ As Single

Private Sub Command1_Click()

CurrentStation.Name = txtName
StationName = CurrentStation.Name
StationInitialized = True
CurrentStation.X = txtX
CurrentStation.Y = txtY
CurrentStation.z = txtZ
frmMain.lblStationWarning.Visible = False
mdiMain.StatusBar.Panels(5) = "Current Station: " + StationName + "  "

Dim Inidata(7, 2) As String
Dim IniClass As String
Dim Status As Byte
IniClass = "[EDM]"
Inidata(1, 1) = "CurrentStation"
Inidata(1, 2) = CurrentStation.Name
Inidata(2, 1) = "StationX"
Inidata(2, 2) = CurrentStation.X
Inidata(3, 1) = "stationY"
Inidata(3, 2) = CurrentStation.Y
Inidata(4, 1) = "stationZ"
Inidata(4, 2) = CurrentStation.z
Inidata(5, 1) = "ReferenceDatum"
Inidata(5, 2) = ""
Inidata(6, 1) = "ReferenceDatum2"
Inidata(6, 2) = ""
Inidata(7, 1) = "SetupType"
Inidata(7, 2) = 0
Call WriteIni(CFGName, IniClass, Inidata(), Status)

Unload Me

End Sub

Private Sub Command2_Click()

TempCurrentStationX = CurrentStation.X
TempCurrentStationy = CurrentStation.Y
TempCurrentStationZ = CurrentStation.z
CurrentStation.X = 0
CurrentStation.Y = 0
CurrentStation.z = 0

Call takeshot_core(NoPrism)
txtX = Format(-edmshot.X, "####0.000")
txtY = Format(-edmshot.Y, "####0.000")
txtZ = Format(-edmshot.z, "####0.000")
mdiMain.StatusBar.Panels(6).Visible = False

End Sub

Private Sub Command3_Click()

CurrentStation.X = TempCurrentStationX
CurrentStation.Y = TempCurrentStationy
CurrentStation.z = TempCurrentStationZ

Unload Me

End Sub

Private Sub Form_Load()

Me.Height = 3444
Me.Width = 6360
txtName = CurrentStation.Name

End Sub


