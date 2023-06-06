VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Debug Total Station"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox ts_commands 
      Height          =   315
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy to Clipboard"
      Height          =   495
      Left            =   6120
      TabIndex        =   19
      Top             =   7560
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   3600
      TabIndex        =   18
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Frame Frame 
      Caption         =   "Results"
      Height          =   3855
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   3255
      Begin VB.TextBox txtPoleHeight 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   23
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox txtZ 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   21
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtY 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   16
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtX 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtSloped 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   12
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtVangle 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtHangle 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Pole H."
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Z"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Y"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "X"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Slope-D"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "V-Angle"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "H-Angle"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox txtComSettings 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton execute_command 
      Caption         =   "Send"
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton btnXShot 
      Caption         =   "X-Shot"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   3255
   End
   Begin VB.CommandButton btnInitialize 
      Caption         =   "Initialize"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox txtOutput 
      Height          =   7335
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label Label2 
      Caption         =   "COM Settings"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   $"frmDebug.frx":0000
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnInitialize_Click()

Call initcomport(comport, errorcode)

End Sub

Private Sub execute_command_Click()

cmd$ = ts_commands.Text
If cmd$ = "" Then
    response = MsgBox("Select a command from the dropdown next to this button.", vbInformation + vbOKOnly)
    Exit Sub
End If

If Not frmMain.theoport.PortOpen Then
    MsgBox ("The COM port is not open.  Try initialize first.")
    Exit Sub
End If

Select Case cmd$
Case "0 Horizontal Angle"
    angle$ = ""
    Call sethortangle(angle$, 0, 0, 0)
Case "Get date/time"
    Call get_date_time
Case "Get Instr. Name"
    Call get_station_name
Case "Launch measurement"
    Call launch_measurement
Case "Get meas. with tilt"
    measurement$ = get_measure_with_tilt()
    If measurement$ <> "" Then
        Call parse_tilt_measure(measurement$, edmshot, errorcode)
        If errorcode = 0 Then
            Call convert_edmshot_to_dms(edmshot)
            Call vhdtonez(edmshot)
            txtHangle = Format(edmshot.hangle, "####0.0000")
            txtVangle = Format(edmshot.vangle, "####0.0000")
            txtSloped = Format(edmshot.sloped, "####0.000")
            txtX = Format(edmshot.X, "####0.000")
            txtY = Format(edmshot.y, "####0.000")
            txtZ = Format(edmshot.z, "####0.000")
            txtPoleHeight = Format(edmshot.poleh, "####0.000")
        End If
    End If
Case Else
    response = MsgBox("Error: Unrecognize command selected.", vbCritical + vbOKOnly)
End Select

End Sub

Private Sub btnXShot_Click()

If Not frmMain.theoport.PortOpen Then
    MsgBox ("The COM port is not open.  Try initialize first.")
Else
    Call takeshot_core(NoPrism)
    txtHangle = Format(edmshot.hangle, "####0.0000")
    txtVangle = Format(edmshot.vangle, "####0.0000")
    txtSloped = Format(edmshot.sloped, "####0.000")
    txtX = Format(edmshot.X, "####0.000")
    txtY = Format(edmshot.y, "####0.000")
    txtZ = Format(edmshot.z, "####0.000")
    Screen.MousePointer = 1
End If

End Sub

Private Sub cmdClear_Click()

txtOutput.Text = ""

End Sub

Private Sub Command1_Click()

Clipboard.SetText (txtOutput.Text)

End Sub

Private Sub Form_Load()

txtComSettings = comport + ":" + comsettings

ts_commands.AddItem ("0 Horizontal Angle")
ts_commands.AddItem ("Get date/time")
ts_commands.AddItem ("Get Instr. Name")
ts_commands.AddItem ("Launch measurement")
ts_commands.AddItem ("Get meas. with tilt")

End Sub

