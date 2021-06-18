VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStationVerify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Station Verify"
   ClientHeight    =   5730
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   7305
   ControlBox      =   0   'False
   Icon            =   "frmStationVerify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel Shot"
      Height          =   444
      Left            =   5040
      TabIndex        =   23
      Top             =   2070
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   444
      Left            =   3120
      TabIndex        =   22
      Top             =   5040
      Width           =   1065
   End
   Begin VB.TextBox txtStationName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   600
      Width           =   1665
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "Record Datum"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      TabIndex        =   19
      Top             =   2040
      Width           =   1575
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
      Height          =   1815
      Left            =   2640
      TabIndex        =   17
      Top             =   3000
      Visible         =   0   'False
      Width           =   4470
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   1350
         Left            =   150
         TabIndex        =   18
         Top             =   330
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2381
         _Version        =   393216
         Rows            =   4
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
   End
   Begin VB.TextBox txtReferenceAngle 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1440
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1410
   End
   Begin VB.TextBox txtX 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1410
   End
   Begin VB.TextBox txtY 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1410
   End
   Begin VB.TextBox txtZ 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reference Datum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   2880
      Width           =   2295
      Begin VB.TextBox txtZ 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   696
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1410
      End
      Begin VB.TextBox txtY 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   696
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1410
      End
      Begin VB.TextBox txtX 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   696
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   1410
      End
      Begin VB.ComboBox Station 
         Height          =   315
         Index           =   1
         Left            =   168
         TabIndex        =   1
         Top             =   336
         Width           =   1950
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Z"
         Height          =   195
         Index           =   1
         Left            =   525
         TabIndex        =   7
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Y"
         Height          =   195
         Index           =   1
         Left            =   510
         TabIndex        =   6
         Top             =   1080
         Width           =   105
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   195
         Index           =   1
         Left            =   525
         TabIndex        =   5
         Top             =   720
         Width           =   90
      End
      Begin VB.Image Image2 
         Height          =   276
         Index           =   0
         Left            =   120
         Picture         =   "frmStationVerify.frx":000C
         Stretch         =   -1  'True
         Top             =   936
         Width           =   288
      End
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      Caption         =   "Select  Reference datum, aim, and then click on Record Datum button. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   3
      Left            =   3240
      TabIndex        =   20
      Top             =   1560
      Width           =   3240
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Angle to Reference Datum:"
      Height          =   855
      Left            =   360
      TabIndex        =   16
      Top             =   2040
      Width           =   960
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   1275
      Left            =   120
      Picture         =   "frmStationVerify.frx":3C1E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   915
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "X"
      Height          =   195
      Index           =   0
      Left            =   1200
      TabIndex        =   14
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Y"
      Height          =   195
      Index           =   0
      Left            =   1200
      TabIndex        =   13
      Top             =   1440
      Width           =   105
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Z"
      Height          =   165
      Index           =   0
      Left            =   1200
      TabIndex        =   12
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
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
      Height          =   285
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   2010
   End
End
Attribute VB_Name = "frmStationVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRecord_Click()

cmdRecord.Enabled = False
takeshot_core AskForPrism
mdiMain.StatusBar.Panels(6).Visible = False

If Cancelling Then
    cmdRecord.Enabled = True
    Exit Sub
End If
cmdRecord.Enabled = True
If Cancelling Then Exit Sub

'need take shot code here
Grid.Cols = 4
Grid.Rows = 4
Grid.Row = 0
Grid.Col = 1
Grid.ColWidth(0) = 300
Grid = "Expected"
Grid.Col = 2
Grid = "Recorded"
Grid.Col = 3
Grid = "Difference"
Grid.Row = 1
Grid.Col = 0
Grid = "X"
Grid.Row = 2
Grid = "Y"
Grid.Row = 3
Grid = "Z"
Grid.Row = 1
Grid.Col = 1
Grid = txtX(1)
Grid.Row = 1
Grid.Col = 2
Grid = edmshot.X
Grid.Row = 1
Grid.Col = 3
Grid = Format(Abs(txtX(1) - edmshot.X), "#####0.000")
Grid.Row = 2
Grid.Col = 1
Grid = txtY(1)
Grid.Row = 2
Grid.Col = 2
Grid = edmshot.y
Grid.Row = 2
Grid.Col = 3
Grid = Format(Abs(txtY(1) - edmshot.y), "####0.000")
Grid.Row = 3
Grid.Col = 1
Grid = txtZ(1)
Grid.Col = 2
Grid = edmshot.z
Grid.Col = 3
Grid = Format(Abs(txtZ(1) - edmshot.z), "####0.000")
frmCoordinates.Visible = True

End Sub

Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command2_Click()

If mdiMain.StatusBar.Panels(7).Visible Then
    Cancelling = True
    Exit Sub
ElseIf Shooting Then
    Exit Sub
Else
    Unload Me
End If

End Sub

Private Sub Form_Load()

If DatumTB.RecordCount < 2 Then
    MsgBox ("Verification requires that at least two datums be defined.")
    Exit Sub
End If
txtstationname = CurrentStation.Name
txtX(0) = Format(CurrentStation.X, "#####0.000")
txtY(0) = Format(CurrentStation.y, "#####0.000")
txtZ(0) = Format(CurrentStation.z, "#####0.000")
CenterForm Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

Me.Hide
frmMain.Picture1.SetFocus

End Sub

Private Sub Station_Click(Index As Integer)

DatumTB.Index = "datumname"
DatumTB.Seek "=", Station(Index)
If Not DatumTB.NoMatch Then
    txtX(Index) = Format(DatumTB("x"), "#####0.000")
    txtY(Index) = Format(DatumTB("y"), "#####0.000")
    txtZ(Index) = Format(DatumTB("z"), "#####0.000")
End If

computeangle txtX(1), txtY(1), txtX(0), txtY(0), angle, minutes, seconds
txtReferenceAngle = Trim(Str(angle)) + "." + Right("00" + Trim(Str(minutes)), 2) + Right("00" + Trim(Str(seconds)), 2)
cmdRecord.Caption = "Record Datum"
cmdRecord.Enabled = True

End Sub

Private Sub Station_DropDown(Index As Integer)

If DatumTB.RecordCount = 0 Then
    MsgBox ("No datums yet defined for this site")
    Exit Sub
End If
Station(Index).Clear
DatumTB.MoveFirst
While Not DatumTB.EOF
    Station(Index).AddItem DatumTB("Name")
    DatumTB.MoveNext
Wend

End Sub

