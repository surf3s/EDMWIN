VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Debug Total Station"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Copy to Clipboard"
      Height          =   495
      Left            =   4920
      TabIndex        =   21
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   3600
      TabIndex        =   20
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Frame Frame 
      Caption         =   "Results"
      Height          =   3495
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   3255
      Begin VB.TextBox txtZ 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   18
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
      Begin VB.Label Label8 
         Caption         =   "Z"
         Height          =   255
         Left            =   240
         TabIndex        =   19
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
      Width           =   1575
   End
   Begin VB.CommandButton btnSetAngle 
      Caption         =   "Set Angle to 0"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton btnXShot 
      Caption         =   "X-Shot"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton btnInitialize 
      Caption         =   "Initialize"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox txtOutput 
      Height          =   7095
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

Private Sub btnSetAngle_Click()

Call sethortangle(angle$, 0, 0, 0)

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

End Sub
