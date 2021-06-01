VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EDM Status"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtstationname 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2760
      Width           =   1875
   End
   Begin VB.TextBox txtstationheight 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   4200
      Width           =   1875
   End
   Begin VB.TextBox txtstationx 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3120
      Width           =   1875
   End
   Begin VB.TextBox txtstationy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3480
      Width           =   1875
   End
   Begin VB.TextBox txtstationz 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3840
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   372
      Index           =   0
      Left            =   8970
      TabIndex        =   14
      Top             =   150
      Width           =   1092
   End
   Begin VB.TextBox txtPointsTable 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1230
      Width           =   7035
   End
   Begin VB.TextBox txtsettings 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2340
      Width           =   7035
   End
   Begin VB.TextBox txtcomport 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1980
      Width           =   7035
   End
   Begin VB.TextBox txttotalstation 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1620
      Width           =   7035
   End
   Begin VB.TextBox txtdbname 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   7035
   End
   Begin VB.TextBox txtinifile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   7035
   End
   Begin VB.TextBox txtcfgfile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   7035
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "(Not necessarily accurate depending on setup type)"
      Height          =   255
      Left            =   3720
      TabIndex        =   26
      Top             =   4245
      Width           =   3855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "All points are relative to this location which is a point at the center of the total station and not the point on the ground."
      Height          =   495
      Left            =   4080
      TabIndex        =   25
      Top             =   3480
      Width           =   4575
   End
   Begin VB.Line Line2 
      X1              =   3720
      X2              =   3960
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      X1              =   3720
      X2              =   3720
      Y1              =   3120
      Y2              =   4080
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Station Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Station Height :"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Station X :"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Station Y :"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Station Z :"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Points Table: "
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1230
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Settings :"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2340
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "COMPort :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1980
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Station :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1620
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Database :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "EDM INI Filename :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CFG Filename :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

Unload Me

End Sub

Private Sub Form_Load()

txtinifile = fixpath(App.Path) + "edm.ini"
txtsettings.Text = comsettings
txtcomport.Text = comport
txttotalstation.Text = EDMName
txtcfgfile = CFGName
txtdbname = SiteDBname
txtPointsTable = PointTableName
txtstationx.Text = Format(CurrentStation.X, "########0.000")
txtstationy.Text = Format(CurrentStation.y, "########0.000")
txtstationz.Text = Format(CurrentStation.z, "########0.000")
txtstationname.Text = CurrentStation.Name
txtstationheight.Text = Format(stationheight, "########0.000")
CenterForm Me

End Sub

