VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8445
   ClientLeft      =   135
   ClientTop       =   -165
   ClientWidth     =   11490
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btn_fontsize 
      Caption         =   "+"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   67
      Top             =   8040
      Width           =   735
   End
   Begin VB.CommandButton btn_fontsize 
      Caption         =   "---"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   66
      Top             =   8040
      Width           =   735
   End
   Begin VB.CommandButton cmdScroll 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   8040
      Picture         =   "frmMain.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   528
      Width           =   735
   End
   Begin VB.CommandButton cmdScroll 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   8040
      Picture         =   "frmMain.frx":04AE
      Style           =   1  'Graphical
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdScroll 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   2
      Left            =   8040
      Picture         =   "frmMain.frx":0870
      Style           =   1  'Graphical
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdScroll 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   3
      Left            =   8040
      Picture         =   "frmMain.frx":0C32
      Style           =   1  'Graphical
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&X-Shot"
      Height          =   735
      Left            =   9105
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2145
   End
   Begin VB.CommandButton cmdPlusShot 
      Caption         =   "&Continue"
      Height          =   735
      Left            =   9105
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   1000
      Width           =   2145
   End
   Begin VB.CommandButton cmdShoot 
      Caption         =   "&New Object"
      Height          =   735
      Left            =   9105
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   200
      Width           =   2145
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel Shot"
      Height          =   1575
      Left            =   9105
      TabIndex        =   55
      Top             =   1000
      Visible         =   0   'False
      Width           =   2145
   End
   Begin MSAdodcLib.Adodc UnitsADO 
      Height          =   345
      Left            =   1080
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      ConnectMode     =   16
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   1
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc PointsADO 
      Height          =   420
      Left            =   120
      Top             =   4440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   741
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   1
      MaxRecords      =   0
      BOFAction       =   1
      EOFAction       =   1
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Obj"
      Height          =   735
      Left            =   9120
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   7400
      Width           =   2145
   End
   Begin VB.TextBox txtPoleHT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   792
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2250
      Width           =   1500
   End
   Begin VB.ComboBox txtXYZ 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   2850
      TabIndex        =   41
      Text            =   "txtXYZ"
      Top             =   2250
      Width           =   1635
   End
   Begin VB.CommandButton Button 
      Caption         =   "Button6"
      Height          =   735
      Index           =   6
      Left            =   9105
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.CommandButton Button 
      Caption         =   "Button5"
      Height          =   735
      Index           =   5
      Left            =   9105
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5800
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.CommandButton Button 
      Caption         =   "Button4"
      Height          =   735
      Index           =   4
      Left            =   9105
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5000
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.CommandButton Button 
      Caption         =   "Button3"
      Height          =   735
      Index           =   3
      Left            =   9105
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.CommandButton Button 
      Caption         =   "Button2"
      Height          =   735
      Index           =   2
      Left            =   9105
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3400
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.CommandButton Button 
      Caption         =   "Button1"
      Height          =   735
      Index           =   1
      Left            =   9105
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2600
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.ComboBox txtXYZ 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2850
      TabIndex        =   6
      Text            =   "txtXYZ"
      Top             =   1890
      Width           =   1635
   End
   Begin VB.ComboBox txtXYZ 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2850
      TabIndex        =   5
      Text            =   "txtXYZ"
      Top             =   1530
      Width           =   1635
   End
   Begin VB.ComboBox txtPrism 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmMain.frx":10D4
      Left            =   792
      List            =   "frmMain.frx":10D6
      TabIndex        =   4
      Text            =   "txtPrism"
      Top             =   1890
      Width           =   1500
   End
   Begin VB.TextBox txtSlopeD 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6165
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "txtSlopeD"
      Top             =   2220
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.TextBox txtVangle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6165
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "txtVangle"
      Top             =   1875
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.TextBox txtHangle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6165
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "txtHangle"
      Top             =   1530
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.ComboBox txtID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Text            =   "txtID"
      Top             =   970
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ComboBox txtUnit 
      DataSource      =   "PointsADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmMain.frx":10D8
      Left            =   1872
      List            =   "frmMain.frx":10DA
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "txtUnit"
      Top             =   970
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtSuffix 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   3
      Text            =   "txtSuffix"
      Top             =   970
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox MenuBox 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      ItemData        =   "frmMain.frx":10DC
      Left            =   4380
      List            =   "frmMain.frx":10DE
      TabIndex        =   9
      Text            =   "MenuBox"
      Top             =   2955
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox TextBox 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2985
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox NumberBox 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2832
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   2985
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4560
      Top             =   3390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm theoport 
      Left            =   3936
      Top             =   3456
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   90
      ScaleHeight     =   495
      ScaleWidth      =   7725
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   7755
      Begin VB.TextBox txtCurrentRecord 
         Alignment       =   1  'Right Justify
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
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtTotalRecords 
         Alignment       =   1  'Right Justify
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
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   120
         Width           =   855
      End
      Begin VB.ComboBox txtPT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   100
         Width           =   2580
      End
      Begin VB.Label lblCurrentRecord 
         AutoSize        =   -1  'True
         Caption         =   "Viewing record"
         Height          =   195
         Left            =   3960
         TabIndex        =   50
         Top             =   165
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "of"
         Height          =   195
         Left            =   6360
         TabIndex        =   49
         Top             =   165
         Width           =   135
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Points Table:"
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   165
         Width           =   930
      End
   End
   Begin VB.Label lbl_fontsize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fontsize"
      Height          =   195
      Left            =   120
      TabIndex        =   65
      Top             =   8160
      Width           =   585
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Last"
      Height          =   195
      Left            =   8040
      TabIndex        =   64
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "First"
      Height          =   195
      Left            =   8040
      TabIndex        =   63
      Top             =   285
      Width           =   735
   End
   Begin VB.Label lblDefaults 
      AutoSize        =   -1  'True
      Caption         =   "Context Defaults are ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Left            =   4680
      TabIndex        =   54
      Top             =   93
      Visible         =   0   'False
      Width           =   1992
   End
   Begin VB.Label LblReflectorless 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EDM Reflectorless"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6750
      TabIndex        =   53
      Top             =   90
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   90
      Top             =   1410
      Width           =   7755
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "With keyboard, to move from one record to another, use the Page-Up Page-Down keys.  To move between fields, use the TAB key."
      Height          =   390
      Left            =   120
      TabIndex        =   44
      Top             =   4020
      Width           =   6645
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAutoFind 
      AutoSize        =   -1  'True
      Caption         =   "Auto-Find Units set to ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Left            =   96
      TabIndex        =   40
      Top             =   96
      Visible         =   0   'False
      Width           =   2196
   End
   Begin VB.Label lblBlankFields 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Record contains blank fields "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2760
      TabIndex        =   39
      Top             =   2715
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Label lblPoleHT 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Height"
      Height          =   195
      Left            =   255
      TabIndex        =   30
      Top             =   2310
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "EDM Fields"
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
      Left            =   160
      TabIndex        =   26
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "X:"
      Height          =   195
      Left            =   2610
      TabIndex        =   25
      Top             =   1590
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Y:"
      Height          =   195
      Left            =   2610
      TabIndex        =   24
      Top             =   1950
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblZ 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Z:"
      Height          =   195
      Left            =   2610
      TabIndex        =   23
      Top             =   2310
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblHangle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Horizontal Angle:"
      Height          =   195
      Left            =   4920
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblVangle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Vertical Angle:"
      Height          =   195
      Left            =   5040
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label lblSlopeD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Slope Distance:"
      Height          =   195
      Left            =   4920
      TabIndex        =   20
      Top             =   2280
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblPrism 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Prism: "
      Height          =   195
      Left            =   270
      TabIndex        =   19
      Top             =   1920
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Object ID"
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
      Left            =   150
      TabIndex        =   15
      Top             =   1000
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Optional Fields"
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
      Left            =   135
      TabIndex        =   14
      Top             =   2685
      Width           =   1590
   End
   Begin VB.Label lblSuffix 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Suffix:"
      Height          =   195
      Left            =   6240
      TabIndex        =   13
      Top             =   1020
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblID 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   195
      Left            =   3720
      TabIndex        =   12
      Top             =   1020
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lblUnit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Unit:"
      Height          =   195
      Left            =   1455
      TabIndex        =   11
      Top             =   1020
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label VarLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "llllllllllllllllllllllllllllllllll"
      Height          =   315
      Index           =   0
      Left            =   -45
      TabIndex        =   10
      Top             =   3030
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblCFGWarning 
      Alignment       =   2  'Center
      Caption         =   "No CFG File Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   2340
      MouseIcon       =   "frmMain.frx":10E0
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   60
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Label lblDBWarning 
      Alignment       =   2  'Center
      Caption         =   "No Site Database Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   2340
      MouseIcon       =   "frmMain.frx":13EA
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   60
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Label lblPointsWarning 
      Alignment       =   2  'Center
      Caption         =   "No Points Table Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   2340
      MouseIcon       =   "frmMain.frx":16F4
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   60
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Label lblPoleWarning 
      Alignment       =   2  'Center
      Caption         =   "No Prisms Defined"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   2340
      MouseIcon       =   "frmMain.frx":19FE
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   60
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Label lblEDMWarning 
      Alignment       =   2  'Center
      Caption         =   "No EDM defined"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   2340
      MouseIcon       =   "frmMain.frx":1D08
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   60
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Label lblStationWarning 
      Alignment       =   2  'Center
      Caption         =   "Station not Initialized"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   2340
      MouseIcon       =   "frmMain.frx":2012
      MousePointer    =   99  'Custom
      TabIndex        =   43
      Top             =   60
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Label lblUnitsWarning 
      Alignment       =   2  'Center
      Caption         =   "No Units Defined"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   2340
      MouseIcon       =   "frmMain.frx":231C
      MousePointer    =   99  'Custom
      TabIndex        =   51
      Top             =   60
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Shape Shape3 
      Height          =   465
      Left            =   90
      Top             =   910
      Width           =   7755
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChangingXYZ As Boolean
Dim OriginalSuffix As Integer
Dim OrigIndex As Integer
Dim OrigValue As String
Dim Dropping As Boolean
Dim OriginalXYZ(0 To 2) As Single
Const IDonly = 1
Const Everything = 2
Const UnitFieldOnly = 3
Const GetNextID = 1
Const DecID = 2
Const SetID = 3
Const SetField = 4
Const DelRec = 5
Public SelectedUnit As String
Dim XYZChanged(3) As Boolean
Dim VarChanged(50) As Boolean
Dim IDChanged As Boolean
Dim SuffixChanged As Boolean
Dim DupID As Boolean
Public DupOption As Integer
Public OffsetValue As Single
'Dim Conn As New ADODB.Connection
'Dim JRO As New JRO.JetEngine
Dim PrevBookMark As Variant
Dim Status As Byte
Dim Need2Decrement As Boolean

Public Sub FormatVarList()

Dim LabelLeft As Integer
Dim LabelTop As Integer
Dim BoxLeft As Integer
Dim BoxTop As Integer
Dim Noptionals As Integer
Dim MenuString As String
Dim OriginalLeft As Integer
Dim LastOptional As Integer
Loading = True

On Error Resume Next
For I = 1 To 50
    Unload VarLabel(I)
    Unload TextBox(I)
    Unload NumberBox(I)
    Unload MenuBox(I)
Next I
On Error GoTo 0
'lblUnit.Visible = False
'txtUnit.Visible = False
'lblID.Visible = False
'txtID.Visible = False
'lblSuffix.Visible = False
'txtSuffix.Visible = False
'lblPrism.Visible = False
'txtPrism.Visible = False
'lblPoleHT.Visible = False
'txtPoleHT.Visible = False
'lblX.Visible = False
'txtXYZ(0).Visible = False
'lblY.Visible = False
'txtXYZ(1).Visible = False
'lblZ.Visible = False
'txtXYZ(2).Visible = False
lblVangle.Visible = False
txtVangle.Visible = False
lblHangle.Visible = False
txtHangle.Visible = False
lblSlopeD.Visible = False
txtSloped.Visible = False

LabelLeft = VarLabel(0).Left
OriginalLeft = LabelLeft
LabelTop = VarLabel(0).Top
BoxLeft = TextBox(0).Left
BoxTop = TextBox(0).Top
For I = 1 To Vars
    VPrompt(I) = Trim(LCase(VPrompt(I)))
    If VPrompt(I) = "" Then VPrompt(I) = Trim(LCase(VarList(I)))
    Select Case UCase(VarList(I))
        Case "UNIT"
            lblUnit.Visible = True
            lblUnit = VPrompt(I)
            txtUnit.Visible = True
            txtUnit = ""
            'Set txtUnit.datasource = PointsADO.Recordset
            'txtUnit.DataField = "unit"
        Case "ID"
            lblID.Visible = True
            lblID = VPrompt(I)
            txtID.Visible = True
            txtID = ""
            'Set txtID.datasource = PointsADO.Recordset
            'txtID.DataField = "id"
            
        Case "SUFFIX"
            lblSuffix.Visible = True
            lblSuffix = VPrompt(I)
            txtSuffix.Visible = True
            txtSuffix = ""
            'txtSuffix.DataField = "suffix"
        Case "PRISM"
            lblPrism.Visible = True
            lblPrism = VPrompt(I)
            txtprism.Visible = True
            lblPoleHT.Visible = True
            txtPoleHT.Visible = True
            txtprism = ""
        Case "X"
            lblX.Visible = True
            lblX = VPrompt(I)
            txtXYZ(0).Visible = True
        Case "Y"
            lblY.Visible = True
            lblY = VPrompt(I)
            txtXYZ(1).Visible = True
        Case "Z"
            lblZ.Visible = True
            lblZ = VPrompt(I)
            txtXYZ(2).Visible = True
        Case "VANGLE"
            lblVangle.Visible = True
            lblVangle = VPrompt(I)
            txtVangle.Visible = True
        Case "HANGLE"
            lblHangle.Visible = True
            lblHangle = VPrompt(I)
            txtHangle.Visible = True
        Case "SLOPED"
            lblSlopeD.Visible = True
            lblSlopeD = VPrompt(I)
            txtSloped.Visible = True
        Case Else
            Load VarLabel(I)
            VarLabel(I).Top = LabelTop
            VarLabel(I).Left = LabelLeft
            'VarLabel(I).Caption = UCase(Left(VarList(I), 1)) + LCase(Mid(VarList(I), 2)) + ": "
            VarLabel(I).Visible = True
            VarLabel(I) = VPrompt(I)
            BoxTop = LabelTop
            Select Case VType(I)
                Case "TEXT"
                    Load TextBox(I)
                    TextBox(I).Left = BoxLeft
                    TextBox(I).Top = BoxTop
                    TextBox(I) = ""
                    TextBox(I).Visible = True
                    If VarList(I) = "DATE" Then TextBox(I) = Date
                    If VarList(I) = "TIME" Then TextBox(I) = Time
                Case "NUMERIC"
                    Load NumberBox(I)
                    NumberBox(I).Left = BoxLeft
                    NumberBox(I).Top = BoxTop
                    NumberBox(I) = ""
                    NumberBox(I).Visible = True
                Case "MENU"
                    Load MenuBox(I)
                    MenuBox(I).Left = BoxLeft
                    MenuBox(I).Top = BoxTop
                    MenuBox(I).Visible = True
                    MenuString = VMenu(I)
                    If UpperCase Then
                        MenuString = UCase(MenuString)
                    End If
                    Gotit = False
                    Do Until Gotit
                        X = InStr(MenuString, ",")
                        If X > 0 Then
                            MenuBox(I).AddItem Left(MenuString, X - 1)
                            MenuString = Mid(MenuString, X + 1)
                        Else
                            MenuBox(I).AddItem MenuString
                            Gotit = True
                        End If
                    Loop
                    MenuBox(I) = ""
                    'Set MenuBox(I).datasource = PointsADO.Recordset
                    'MenuBox(I).DataField = VarList(I)
            End Select
            Noptionals = Noptionals + 1
            LastOptional = I
            LabelLeft = LabelLeft + VarLabel(I).Width + MenuBox(0).Width + 100
            BoxLeft = LabelLeft + VarLabel(I).Width + 50
            If Noptionals Mod 2 = 0 Then
                LabelTop = LabelTop + VarLabel(I).Height
                'LabelTop = VarLabel(0).Top
                LabelLeft = OriginalLeft
                'BoxTop = TextBox(0).Top
                BoxTop = LabelTop
                BoxLeft = LabelLeft + VarLabel(I).Width + 50
            End If
    End Select
Next I

bottom_speed_button = Command1.Top + Command1.Height + 100
For I = 1 To 6
    If Button(I).Visible Then
        If Button(I).Top + Button(I).Height + 100 > bottom_speed_button Then
            bottom_speed_button = Button(I).Top + Button(I).Height + 100
        End If
    End If
Next I

LabelTop = LabelTop + VarLabel(LastOptional).Height + 200
If bottom_speed_button > LabelTop Then
    LabelTop = bottom_speed_button
End If

If Label8.Top < LabelTop Then
    Label8.Top = LabelTop
    Command2.Top = LabelTop
End If
Command2.Left = Command1.Left

Me.Height = Command2.Top + Command2.Height + 200

Loading = False

Screen.MousePointer = 1

End Sub

Private Sub reformat_screen()

VarLabel(0).Left = 90
LabelLeft = VarLabel(0).Left
OriginalLeft = LabelLeft
LabelTop = VarLabel(0).Top
'BoxLeft = TextBox(0).Left
BoxLeft = VarLabel(0).Left + VarLabel(0).Width + 50
BoxTop = TextBox(0).Top

For I = 1 To Vars
    Select Case UCase(VarList(I))
        Case "UNIT"
        Case "ID"
        Case "SUFFIX"
        Case "PRISM"
        Case "X"
        Case "Y"
        Case "Z"
        Case "VANGLE"
        Case "HANGLE"
        Case "SLOPED"
        'Case "DATUMX"
        'Case "DATUMY"
        'Case "DATUMZ"
        Case Else
            VarLabel(I).Top = LabelTop
            VarLabel(I).Left = LabelLeft
            VarLabel(I).Width = VarLabel(0).Width
            VarLabel(I).Height = VarLabel(0).Height
            BoxTop = LabelTop
            Select Case VType(I)
                Case "TEXT"
                    TextBox(I).Left = BoxLeft
                    TextBox(I).Top = BoxTop
                    TextBox(I).Height = TextBox(0).Height
                    TextBox(I).Width = TextBox(0).Width
                Case "NUMERIC"
                    NumberBox(I).Left = BoxLeft
                    NumberBox(I).Top = BoxTop
                    NumberBox(I).Height = TextBox(0).Height
                    NumberBox(I).Width = TextBox(0).Width
                Case "MENU"
                    MenuBox(I).Left = BoxLeft
                    MenuBox(I).Top = BoxTop
                    'MenuBox(I).Height = TextBox(0).Height
                    MenuBox(I).Width = TextBox(0).Width
            End Select
            
            Noptionals = Noptionals + 1
            LastOptional = I
            LabelLeft = LabelLeft + VarLabel(0).Width + TextBox(0).Width + 100
            BoxLeft = LabelLeft + VarLabel(0).Width + 50
            If Noptionals Mod 2 = 0 Then
                LabelTop = LabelTop + VarLabel(0).Height + 50
                'LabelTop = VarLabel(0).Top
                LabelLeft = OriginalLeft
                'BoxTop = TextBox(0).Top
                BoxTop = LabelTop
                BoxLeft = LabelLeft + VarLabel(0).Width + 50
            End If
    End Select
Next I

bottom_speed_button = Command1.Top + Command1.Height + 100
For I = 1 To 6
    If Button(I).Visible Then
        If Button(I).Top + Button(I).Height + 100 > bottom_speed_button Then
            bottom_speed_button = Button(I).Top + Button(I).Height + 100
        End If
    End If
Next I

LabelTop = LabelTop + VarLabel(LastOptional).Height + 200
If bottom_speed_button > LabelTop Then
    LabelTop = bottom_speed_button
End If

If Label8.Top < LabelTop Then
    Label8.Top = LabelTop
    Command2.Top = LabelTop
End If
Command2.Left = Command1.Left

Me.Height = Command2.Top + Command2.Height + 200

titlebar_height = (Me.Height - Me.ScaleHeight) - (Me.Width - Me.ScaleWidth) / 2
btn_fontsize(0).Top = Me.Height - titlebar_height - btn_fontsize(0).Height - 50
btn_fontsize(1).Top = Me.Height - titlebar_height - btn_fontsize(1).Height - 50
lbl_fontsize.Top = btn_fontsize(0).Top + btn_fontsize(0).Height / 2 - lbl_fontsize.Height / 2

End Sub

Private Function min(A, B) As Integer
    If A < B Then
        min = A
    Else
        min = B
    End If
    
End Function

Private Function max(A, B) As Integer
    If A > B Then
        max = A
    Else
        max = B
    End If
    
End Function

Private Sub btn_fontsize_Click(Index As Integer)

Select Case Index
Case 0
    For I = 0 To Vars
        On Error Resume Next
        TextBox(I).FontSize = max(TextBox(I).FontSize - 2, 8)
        MenuBox(I).FontSize = max(MenuBox(I).FontSize - 2, 8)
        NumberBox(I).FontSize = max(NumberBox(I).FontSize - 2, 8)
        VarLabel(I).FontSize = max(VarLabel(I).FontSize - 2, 8)
    Next I
    TextBox(0).Height = max(TextBox(0).Height - 40, 315)
    TextBox(0).Width = max(TextBox(0).Width - 200, 2055)
    VarLabel(0).Height = max(VarLabel(0).Height - 40, 315)
    VarLabel(0).Width = max(VarLabel(0).Width - 100, 1200)
    reformat_screen
Case 1
    For I = 0 To Vars
        On Error Resume Next
        TextBox(I).FontSize = min(TextBox(I).FontSize + 2, 20)
        MenuBox(I).FontSize = min(MenuBox(I).FontSize + 2, 20)
        NumberBox(I).FontSize = min(NumberBox(I).FontSize + 2, 20)
        VarLabel(I).FontSize = min(VarLabel(I).FontSize + 2, 20)
    Next I
    TextBox(0).Height = min(TextBox(0).Height + 40, 515)
    TextBox(0).Width = min(TextBox(0).Width + 200, 3055)
    VarLabel(0).Height = min(VarLabel(0).Height + 40, 515)
    VarLabel(0).Width = min(VarLabel(0).Width + 100, 2200)
    reformat_screen
Case Else
End Select

On Error GoTo 0

End Sub

Private Sub Button_Click(Index As Integer)
    
If CheckStatus() = True Then Exit Sub

Screen.MousePointer = 11
Cancelling = False

cmdShoot_Click
mdiMain.StatusBar.Panels(6).Picture = mdiMain.Picture2(3).Picture
If Cancelling Then
    cmdShoot.Enabled = True
    cmdPlusShot.Enabled = True
    Command1.Enabled = True
    For I = 1 To 6
        Button(I).Enabled = True
    Next I
    
    Picture1.SetFocus
    Exit Sub
End If

For I = 1 To nButtonVars(Index)
    If LCase(VarList(ButtonVars(Index, I, 1))) = "unit" Then
        If txtUnit.Text <> ButtonVars(Index, I, 2) Then
            txtUnit = ButtonVars(Index, I, 2)
            'txtUnit_KeyPress 13
            txtUnit_Click
        End If
        
    ElseIf LCase(VarList(ButtonVars(Index, I, 1))) = "id" Then
        If LCase(ButtonVars(Index, I, 2)) = "alpha" Then
            Dim NewID As String
            OriginalUnit = txtUnit
            OriginalID = txtID
            OriginalSuffix = Val(txtSuffix)
            If Val(Trim(txtID)) > 0 Then
                DecrementID OriginalUnit, OriginalID, OriginalSuffix
                txtID = hash(5)
            End If
            'PointsADO.Refresh
            'PointsADO.Recordset.Bookmark = CurrentBookMark
            PointsADO.recordset.Update "id", txtID
            PointsADO.recordset.Update "suffix", 0
            OriginalID = txtID
        End If
    
    ElseIf LCase(VarList(ButtonVars(Index, I, 1))) = "prism" Then
        Gotit = False
        For J = 0 To txtprism.ListCount - 1
            If LCase(txtprism.List(J)) = LCase(ButtonVars(Index, I, 2)) Then
                Loading = True
                txtprism.ListIndex = J
                Loading = False
                Gotit = True
                Exit For
            End If
        Next J
        If Not Gotit Then
            MsgBox ("Prism name not found in current prism list")
        Else
            txtPoleHT = PoleHeight(txtprism.ItemData(txtprism.ListIndex))
            txtXYZ(2) = Format(Val(txtXYZ(2)) + OriginalPoleHT - PoleHeight(txtprism.ItemData(txtprism.ListIndex)), "#######0.000")
            'PointsADO.Refresh
            'PointsADO.Recordset.Bookmark = CurrentBookMark
            PointsADO.recordset.Update "prism", txtPoleHT
            PointsADO.recordset.Update "z", txtXYZ(2)
        End If
    Else
        Select Case VType(ButtonVars(Index, I, 1))
            Case "TEXT"
                TextBox(ButtonVars(Index, I, 1)).Text = ButtonVars(Index, I, 2)
                TextBox(ButtonVars(Index, I, 1)).SelLength = 0
                TextBox(ButtonVars(Index, I, 1)).Refresh
            Case "MENU"
                MenuBox(ButtonVars(Index, I, 1)) = ButtonVars(Index, I, 2)
                MenuBox(ButtonVars(Index, I, 1)).SelLength = 0
                MenuBox(ButtonVars(Index, I, 1)).Refresh
            Case "NUMERIC", "INSTRUMENT"
                NumberBox(ButtonVars(Index, I, 1)) = ButtonVars(Index, I, 2)
                NumberBox(ButtonVars(Index, I, 1)).SelLength = 0
                NumberBox(ButtonVars(Index, I, 1)).Refresh
        End Select
        'PointsADO.Recordset.Requery
        'PointsADO.Recordset.Update
        'PointsADO.Refresh
        'PointsADO.Recordset.Bookmark = CurrentBookMark
        PointsADO.recordset.Update VarList(ButtonVars(Index, I, 1)), ButtonVars(Index, I, 2)
        PointsADO.recordset.Update
    End If
Next I
mdiMain.StatusBar.Panels(6).Visible = False

ShowValues
FindBlankField
If Speaking Then
    SpeakID txtUnit, txtID
End If

cmdShoot.Enabled = True
cmdPlusShot.Enabled = True
Command1.Enabled = True
For I = 1 To 6
    Button(I).Enabled = True
Next I

Picture1.SetFocus
FindBlankField

Screen.MousePointer = 1

End Sub

Public Sub cmdCancel_Click()

If mdiMain.StatusBar.Panels(7).Visible Then
    Cancelling = True
    mdiMain.StatusBar.Panels(7).Visible = False

    Exit Sub
Else
    Picture1.SetFocus
End If

End Sub

Public Sub cmdPlusShot_Click()

Dim MaxSuffix As Integer

If CheckStatus() = True Then Exit Sub

If PointsADO.recordset.EOF And PointsADO.recordset.BOF Then
    MsgBox ("No initial record in this series has been recorded.  Shoot as new object")
    Exit Sub
End If

If txtUnit = "" Or txtID = "" Then
    MsgBox ("You cannot continue with an object unless it has a valid Unit and ID.")
    Exit Sub
End If

Picture1.SetFocus

currentrecord = PointsADO.recordset.Bookmark

GridLoading = True
PointsADO.recordset.MoveLast
If Not PointsADO.recordset.EOF Then
    If PointsADO.recordset("unit") <> txtUnit Or PointsADO.recordset("id") <> txtID Then
    
        response = MsgBox("Continue with object " + txtUnit + "-" + Trim(txtID) + "?" + Chr(13) + "(Press No to continue with " + PointsADO.recordset("unit") + "-" + PointsADO.recordset("id") + ")", vbYesNoCancel)
    End If
    If response = vbCancel Then
        PointsADO.recordset.Bookmark = currentrecord
        GridLoading = False
        Exit Sub
    ElseIf response = vbNo Then
        ShowValues
    End If
End If

MaxSuffix = -1
PointsADO.recordset.MoveFirst
Do
    If PointsADO.recordset("unit") = txtUnit And PointsADO.recordset("id") = txtID Then
        If PointsADO.recordset("suffix") > MaxSuffix Then
            MaxSuffix = PointsADO.recordset("suffix")
        End If
    End If
    PointsADO.recordset.MoveNext
Loop Until PointsADO.recordset.EOF

'PointsADO.Recordset.Filter = "unit='" + txtUnit + "' and id='" + txtID + "'"
'PointsADO.Recordset.MoveFirst
'MaxSuffix = PointsADO.Recordset("suffix")

'While Not PointsADO.Recordset.EOF
'    If PointsADO.Recordset("suffix") > MaxSuffix Then MaxSuffix = PointsADO.Recordset("suffix")
'    PointsADO.Recordset.MoveNext
'Wend

'PointsADO.Recordset.Filter = adFilterNone
Take_Shot XShot
mdiMain.StatusBar.Panels(6).Picture = mdiMain.Picture2(3).Picture

If Cancelling Then
    PointsADO.recordset.Bookmark = currentrecord
    cmdShoot.Enabled = True
    cmdPlusShot.Enabled = True
    Command1.Enabled = True
    For I = 1 To 6
        Button(I).Enabled = True
    Next I
    GridLoading = False
    Exit Sub
End If
txtSuffix = MaxSuffix + 1
'PointsADO.Recordset.Bookmark = CurrentBookMark
PointsADO.recordset.Update "unit", txtUnit
PointsADO.recordset.Update "id", txtID
PointsADO.recordset.Update "suffix", txtSuffix
mdiMain.StatusBar.Panels(6).Visible = False

ShowValues
If PlotShowing Then
    frmPlot.shpX.Visible = False

    If mdiMain.mnuViewPoints.Checked Then
        frmPlot.SetScale
        frmPlot.PlotPoints
    End If
End If
'Set rsTemp = Nothing
Picture1.SetFocus
FindBlankField

cmdShoot.Enabled = True
cmdPlusShot.Enabled = True
Command1.Enabled = True
For I = 1 To 6
    Button(I).Enabled = True
Next I
GridLoading = False

End Sub

Public Sub cmdScroll_Click(Index As Integer)

Select Case Index
    Case 0
        Form_KeyDown vbKeyHome, 0
    Case 1
        Form_KeyDown vbKeyPageUp, 0
    Case 2
        Form_KeyDown vbKeyPageDown, 0
    Case 3
        Form_KeyDown vbKeyEnd, 0
End Select

End Sub

Private Sub cmdShoot_Click()

If CheckStatus() = True Then Exit Sub

Picture1.SetFocus
Screen.MousePointer = 11

txtSuffix = 0
If Not (PointsADO.recordset.BOF Or PointsADO.recordset.EOF) Then
    PrevBookMark = PointsADO.recordset.Bookmark
End If

Take_Shot NewShot

If Cancelling Then
    If Not (PointsADO.recordset.BOF And PointsADO.recordset.EOF) Then
        PointsADO.recordset.Bookmark = PrevBookMark
        ShowValues
    End If
    Exit Sub
End If

mdiMain.StatusBar.Panels(6).Picture = mdiMain.Picture2(3).Picture
If LimitChecking And txtSuffix = 0 Then
    FindUnit
ElseIf Not LimitChecking And txtSuffix = 0 Then
    txtUnit_KeyPress 13
End If

If PlotShowing Then
    frmPlot.shpX.Visible = False
    If mdiMain.mnuViewPoints.Checked Then
        frmPlot.SetScale
        frmPlot.PlotPoints
    End If
End If

If Speaking Then
    SpeakID txtUnit, txtID
End If

cmdShoot.Enabled = True
cmdPlusShot.Enabled = True
Command1.Enabled = True
For I = 1 To 6
    Button(I).Enabled = True
Next I

Screen.MousePointer = 1
Picture1.SetFocus
FindBlankField
mdiMain.StatusBar.Panels(6).Visible = False

End Sub

Private Sub cmdShoot_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

A = 1

End Sub

Public Sub Command1_Click()

Dim edmpoffset As Single

If SiteDBname = "" Then
    MsgBox ("Open Site Database and define prisms before recording points")
    Exit Sub
End If
If Not StationInitialized Then
    MsgBox ("Total Station not initialized.  Initialize before recording points")
    Exit Sub
ElseIf PoleTB.BOF And PoleTB.EOF Then
    MsgBox ("No prisms defined.  Define before taking a shot")
    Exit Sub
End If
If Not frmMain.theoport.PortOpen And EDMName <> "Simulate" And EDMName <> "Microscribe" And EDMName <> "None" Then
    MsgBox ("Total Station not cabled")
    Exit Sub
End If
On Error Resume Next
Picture1.SetFocus

On Error GoTo 0
If XShotShowing Then
    If frmXShot.txtprism.ListIndex <> -1 Then
        edmshot.poleh = PoleHeight(frmXShot.txtprism.ItemData(frmXShot.txtprism.ListIndex))
        edmshot.poleo = PoleOffset(frmXShot.txtprism.ItemData(frmXShot.txtprism.ListIndex))
    Else
        edmshot.poleh = 0
        edmshot.poleo = 0
    End If
Else
    cmdCancel.Visible = True
    If txtprism.ListIndex <> -1 Then
        edmshot.poleh = PoleHeight(txtprism.ItemData(txtprism.ListIndex))
        edmshot.poleo = PoleOffset(txtprism.ItemData(txtprism.ListIndex))
    Else
        edmshot.poleh = 0
        edmshot.poleo = 0
    End If
End If
cmdShoot.Enabled = False
cmdPlusShot.Enabled = False
Command1.Enabled = False
For I = 1 To 6
    Button(I).Enabled = False
Next I
mdiMain.StatusBar.Panels(6).Picture = mdiMain.Picture2(3).Picture

Call takeshot_core(NoPrism)
cmdCancel.Visible = False
If Cancelling Then GoTo ExitSub

actuald = Sqr((CurrentStation.X - edmshot.X) ^ 2 + (CurrentStation.y - edmshot.y) ^ 2)

frmXShot.lblvalue(0) = Format(edmshot.X, "####0.000")
frmXShot.lblvalue(1) = Format(edmshot.y, "####0.000")
frmXShot.lblvalue(2) = Format(edmshot.z, "####0.000")
frmXShot.lblvalue(3) = Format(edmshot.hangle, "####0.0000")
frmXShot.lblvalue(4) = Format(edmshot.vangle, "####0.0000")
frmXShot.lblvalue(5) = Format(edmshot.sloped, "####0.000")
frmXShot.lblvalue(6) = Format(edmshot.X - CurrentStation.X, "####0.000")
frmXShot.lblvalue(7) = Format(edmshot.y - CurrentStation.y, "####0.000")
frmXShot.lblvalue(8) = Format(edmshot.z - CurrentStation.z, "####0.000")
frmXShot.FindUnit edmshot.X, edmshot.y
If Not XShotShowing Then
    frmXShot.txtprism.Clear
    For I = 0 To frmMain.txtprism.ListCount - 1
            frmXShot.txtprism.AddItem frmMain.txtprism.List(I)
            frmXShot.txtprism.ItemData(frmXShot.txtprism.NewIndex) = frmMain.txtprism.ItemData(I)
    Next I
    Loading = True
    If frmXShot.txtprism.ListCount > 0 Then
        frmXShot.txtprism.ListIndex = frmMain.txtprism.ListIndex
    End If
    Loading = False
    XShotShowing = True
    If PlotShowing Then
        frmPlot.shpX.Visible = True
        frmPlot.SetScale
        frmPlot.PlotPoints
    End If
    mdiMain.StatusBar.Panels(6).Visible = False
    frmXShot.Show 1
End If
If Cancelling Then GoTo ExitSub
If PlotShowing Then
    frmPlot.shpX.Visible = True
    frmPlot.SetScale
    frmPlot.PlotPoints
'    frmPlot.Show
End If

ExitSub:
cmdShoot.Enabled = True
cmdPlusShot.Enabled = True
Command1.Enabled = True
For I = 1 To 6
    Button(I).Enabled = True
Next I
mdiMain.StatusBar.Panels(6).Visible = False

End Sub

Private Sub Command2_Click()

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then
    MsgBox "Open points table before performing this operation", vbInformation
    Exit Sub
End If

If PointsADO.recordset.EOF And PointsADO.recordset.BOF Then Exit Sub

frmMain.DeleteRecord
ShowValues

If mdiMain.mnuViewPoints.Checked Then
    frmPlot.SetScale
    frmPlot.PlotPoints
End If

End Sub

Private Sub Command3_Click()

frmSubUnits.Show

End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then Exit Sub
If PointsADO.recordset.EOF And PointsADO.recordset.BOF Then Exit Sub
'Loading = True
Gotit = False

If KeyCode = vbKeyHome Then
    PointsADO.recordset.MoveFirst
    CurrentBookMark = PointsADO.recordset.Bookmark
    Gotit = True
ElseIf KeyCode = vbKeyEscape Then
    If mdiMain.StatusBar.Panels(7).Visible Then
        Cancelling = True
        mdiMain.StatusBar.Panels(7).Visible = False
        Exit Sub
    Else
        Picture1.SetFocus
    End If
ElseIf KeyCode = vbKeyEnd Then
    PointsADO.recordset.MoveLast
    CurrentBookMark = PointsADO.recordset.Bookmark
    Gotit = True
ElseIf KeyCode = vbKeyPageUp Then
    GridLoading = True
    PointsADO.recordset.MovePrevious
    If PointsADO.recordset.BOF Then
        PointsADO.recordset.MoveFirst
    End If
    GridLoading = False
    CurrentBookMark = PointsADO.recordset.Bookmark
    Gotit = True
ElseIf KeyCode = vbKeyPageDown Then
    GridLoading = True
    PointsADO.recordset.MoveNext
    If PointsADO.recordset.EOF Then
        PointsADO.recordset.MoveLast
    End If
    GridLoading = False
    CurrentBookMark = PointsADO.recordset.Bookmark
    Gotit = True
'ElseIf KeyCode = vbKeyDelete Then
'    DeleteRecord
ElseIf KeyCode = vbKeyC And Shift = 2 Then
    cmdPlusShot_Click
    Gotit = True
ElseIf KeyCode = vbKeyN And Shift = 2 Then
    cmdShoot_Click
    Gotit = True
ElseIf KeyCode = vbKeyX And Shift = 2 Then
    Command1_Click
    Gotit = True
End If
If Gotit Then
    txtUnit.ForeColor = 0
    txtUnit.FontBold = True
    ShowValues
    Loading = False
    If GridShowing Then
        frmDataGrid.DataGrid.Refresh
    End If
    'Loading = True
    Picture1.SetFocus
    Loading = False
    KeyCode = 0
    Exit Sub
End If

Gotit = False
If Shift = 2 And KeyCode <> 17 Then
    For I = 1 To 6
        If Button(I).Visible And ButtonShortCut(I) = Chr(KeyCode) Then
            Gotit = True
            Exit For
        End If
    Next I
    If Gotit Then Button_Click Int(I)
End If

End Sub

Private Sub Form_Load()

'Public variables initialized

For A = 1 To 7
    mdiMain.StatusBar.Panels(A).Width = mdiMain.Width / 7
Next A

BannerHeight = 400
BannerWidth = 150

Me.Left = 0
Me.Top = 0
Me.Height = 9000
Me.Width = cmdShoot.Left + cmdShoot.Width + 400

titlebar_height = (Me.Height - Me.ScaleHeight) - (Me.Width - Me.ScaleWidth) / 2
'titlebar_height = (Me.Height - Me.ScaleHeight)

Me.Show
Loading = True
inifile$ = fixpath(App.Path) + "edm.ini"
txtXYZ(0).Clear
txtXYZ(1).Clear
txtXYZ(2).Clear
txtprism.Clear
txtXYZ(0).AddItem "Offset East"
txtXYZ(0).AddItem "Offset West"
txtXYZ(1).AddItem "Offset North"
txtXYZ(1).AddItem "Offset South"
txtXYZ(2).AddItem "Offset Up"
txtXYZ(2).AddItem "Offset Down"

TSLog = False
GeneralLog = False

Call ReadEDMini(inifile$)
'desk.Height = Me.Height

Gotit = True
If SiteDBname = "" Then
    lblDBWarning.Visible = True
    Gotit = False
End If
If CFGName = "" Then
    lblCFGWarning.Visible = True
    Gotit = False
End If
If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then
    lblPointsWarning.Visible = True
    Gotit = False
End If

reformat_screen

'btn_fontsize(0).Top = Me.Height - titlebar_height - btn_fontsize(0).Height - 50
'btn_fontsize(1).Top = Me.Height - titlebar_height - btn_fontsize(1).Height - 50
'lbl_fontsize.Top = btn_fontsize(0).Top + btn_fontsize(0).Height / 2 - lbl_fontsize.Height / 2

Me.Show

Picture1.SetFocus
If Not Gotit Then
    If Dir(inifile) = "" Then
        txtXYZ(0).Enabled = False
        txtXYZ(1).Enabled = False
        txtXYZ(2).Enabled = False
    End If
    Exit Sub
End If

Start:
Screen.MousePointer = 1
If EDMName$ <> "" And comport <> "" And comsettings <> "" Then
    Select Case UCase(EDMName$)
    Case "TOPCON"
        answer = MsgBox("Cable the total station to the computer and communications will be initialized using station type " + EDMName$ + " on comport " + comport + ":" + comsettings + ".", vbOKCancel)
        If answer = 1 Then
            Screen.MousePointer = 11
            Call initcomport(comport, errorcode)
            If Cancelling Then
                MsgBox ("Communications error with total station.  Verify that it is turned on")
                GoTo Start
            End If
            Screen.MousePointer = 1
        End If
    Case "WILD", "LEICA", "BUILDER", "WILD2", "LEICA_AUTOTILT"
        If UCase(EDMName$) = "WILD2" Or UCase(EDMName$) = "LEICA_AUTOTILT" Then
            answer = MsgBox("Cable the total station to the computer and communications will be initialized using station type " + EDMName$ + " (Leica type station with GeoCOM format) on comport " + comport + ":" + comsettings + ".", vbOKCancel)
        Else
            answer = MsgBox("Cable the total station to the computer and communications will be initialized using station type " + EDMName$ + " on comport " + comport + ":" + comsettings + ".", vbOKCancel)
        End If
        If answer = 1 Then
            Screen.MousePointer = 11
            Call initcomport(comport, errorcode)
            If Cancelling Then
                MsgBox ("Communications error with total station.  Verify that it is turned on")
                GoTo Start
            End If
            Screen.MousePointer = 1
        End If
    Case "SOKKIA"
        answer = MsgBox("Cable the total station to the computer and communications will be initialized using station type " + EDMName$ + " on comport " + comport + ":" + comsettings + ".", vbOKCancel)
        If answer = 1 Then
            Screen.MousePointer = 11
            Call initcomport(comport, errorcode)
            If Cancelling Then
                MsgBox ("Communications error with total station.  Verify that it is turned on")
                GoTo Start
            End If
            Screen.MousePointer = 1
        End If
    Case "NONE"
    Case "SIMULATE"
    End Select
End If

'EDMName$ = readcfg("theodolite")
'If EDMName$ = "" Then
'    EDMName$ = "None"
'    Call writecfg("theodolite", "None")
'End If
'comport$ = readcfg("comport")

'temp$ = readcfg("vhd")
'If temp$ = "0" Or temp$ = "" Then
'    vhd = False
'Else
'    vhd = True
'End If

'need code here to deal with initializing the edm
'depending on which one they selected

'Call printstatus

Loading = False

ShowValues
Me.Show
Picture1.SetFocus
'On Error GoTo SAPINotFound
'Set Voice = New SpVoice
Speaking = False

Exit Sub
    
'SAPINotFound:
'    If Err.Number = 459 Or Err.Number = 429 Then
'        MsgBox "SAPI.dll (for speaking option) not found."
'    Else
'        MsgBox "Error encountered : " & Err.Number
'    End If
'    Speaking = False

End Sub

Private Sub Form_Unload(cancel As Integer)

inifile$ = fixpath(App.Path) + "edm.ini"
Call WriteEDMIni(inifile$)
Dim Inidata(100, 2) As String
Dim IniClass As String

IniClass = "[EDM]"
Inidata(2, 1) = "Database"
Inidata(2, 2) = DBName
Inidata(3, 1) = "PointTable"
Inidata(3, 2) = PointTableName
Inidata(4, 1) = "Instrument"
Inidata(4, 2) = EDMName
Inidata(5, 1) = "COMport"
Inidata(5, 2) = comport
Inidata(6, 1) = "Settings"
Inidata(6, 2) = comsettings
Inidata(7, 1) = "SQID"
Inidata(7, 2) = squidcheck
Inidata(8, 1) = "Unitfields"
Inidata(8, 2) = Unitfield(1)
Inidata(9, 1) = "DBPath"
Inidata(9, 2) = DBPath
Inidata(10, 1) = "EDMDelaytime"
Inidata(10, 2) = EDMDelayTime

For I = 2 To nUnitFields
    Inidata(8, 2) = Inidata(8, 2) + "," + Unitfield(I)
Next I
Inidata(9, 1) = "Limitchecking"
If LimitChecking Then
    Inidata(9, 2) = "Yes"
Else
    Inidata(9, 2) = "No"
End If

Dim Status As Byte
Call WriteIni(CFGName, IniClass, Inidata(), Status)

End Sub

Private Sub lblCFGWarning_Click()

mdiMain.mnuOpenCFG_Click

End Sub

Private Sub lblDBWarning_Click()

mdiMain.mnuOpenDB_Click

End Sub

Private Sub lblEDMWarning_Click()

mdiMain.mnuTheodolite_Click

End Sub

Private Sub lblPointsWarning_Click()

mdiMain.mnuNewPointsTB_Click

End Sub

Private Sub lblPoleWarning_Click()

mdiMain.mnuEditPrisms_Click

End Sub

Private Sub lblStationWarning_Click()

mdiMain.mnuInitialize_Click

End Sub

Private Sub MenuBox_Click(Index As Integer)

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then Exit Sub
If PointsADO.recordset.EOF And PointsADO.recordset.BOF Then Exit Sub
If Trim(MenuBox(Index)) = "" Then MenuBox(Index) = ""
If UpperCase Then MenuBox(Index) = UCase(MenuBox(Index))
UpdatePointsTable VarList(Index), MenuBox(Index), 1, 1

For I = 1 To nUnitFields
    If UCase(VarList(Index)) = UCase(Unitfield(I)) Then
        UpdateUnitTable txtUnit, txtID, UnitFieldOnly
        Exit For
    End If
Next I

If UCase(VarList(Index)) = UCase(MasterVar) Then
    MasterVal = MenuBox(Index).Text
    FillDependentFields
Else
    For I = 1 To nDependentVars
        If UCase(DependentVar(I)) = UCase(VarList(Index)) Then
            UpdateDependentVar DependentVar(I), MenuBox(Index).Text
            If Cancelling Then
                MenuBox(Index) = OrigValue
                Cancelling = False
                Exit Sub
            End If
            Exit For
        End If
    Next I
End If

OrigValue = MenuBox(Index)
If GridShowing Then frmDataGrid.DataGrid.Refresh
CheckFields

End Sub

Private Sub MenuBox_DropDown(Index As Integer)

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then
    MsgBox "Open point table first.", vbInformation
    MenuBox(Index) = OriginalValue
    Exit Sub
End If

'MenuString = VMenu(Index)
'If UpperCase Then
'    MenuString = UCase(MenuString)
'End If
'Gotit = False
'Do Until Gotit
'    X = InStr(MenuString, ",")
'    If X > 0 Then
'        MenuBox(I).AddItem Left(MenuString, X - 1)
'        MenuString = Mid(MenuString, X + 1)
'    Else
'        MenuBox(I).AddItem MenuString
'        Gotit = True
'    End If
'Loop
'A = 1

End Sub

Private Sub MenuBox_GotFocus(Index As Integer)

OrigValue = MenuBox(Index)
If Trim(MenuBox(Index)) = "" Then
    MenuBox(Index) = Space(30)
End If
MenuBox(Index).SelStart = 0
MenuBox(Index).SelLength = Len(MenuBox(Index))

End Sub

Private Sub MenuBox_KeyPress(Index As Integer, KeyAscii As Integer)

If Not VarChanged(Index) Then
    OrigValue = MenuBox(Index)
    VarChanged(Index) = True
End If

If KeyAscii = 27 Then
    MenuBox(Index) = OrigValue
    Picture1.SetFocus
    
ElseIf KeyAscii = 13 And Trim(MenuBox(Index)) <> "" Then
        If MenuBox(Index) = OrigValue Then
            Picture1.SetFocus
        Else
            If UpperCase Then
                MenuBox(Index) = UCase(MenuBox(Index))
            End If
            Gotit = False
            For I = 0 To MenuBox(Index).ListCount - 1
                If UCase(MenuBox(Index)) = UCase(MenuBox(Index).List(I)) Then
                    Gotit = True
                    Exit For
                End If
            Next I
            If Not Gotit Then
                If NoAlert Then
                    response = vbYes
                Else
                    response = MsgBox("Add " + MenuBox(Index) + " to list of terms for " + VarList(Index) + "?", vbYesNo)
                End If
                If response = vbNo Then Exit Sub
                MenuBox(Index).AddItem MenuBox(Index)
                If Len(VMenu(Index)) > 0 Then
                    VMenu(Index) = VMenu(Index) + "," + MenuBox(Index)
                Else
                    VMenu(Index) = MenuBox(Index)
                End If
                Dim Inidata(1, 2) As String
                Dim IniClass As String
                IniClass = VarList(Index)
                Inidata(1, 1) = "Menu"
                Inidata(1, 2) = VMenu(Index)
                Dim Status As Byte
                Call WriteIni(CFGName, IniClass, Inidata(), Status)
            End If
            UpdatePointsTable VarList(Index), MenuBox(Index), 1, 1
            For I = 1 To nUnitFields
                If UCase(VarList(Index)) = UCase(Unitfield(I)) Then
                    UpdateUnitTable txtUnit, txtID, UnitFieldOnly
                    Exit For
                End If
            Next I
            OrigValue = MenuBox(Index)
            MenuBox_Click Index
            CheckFields
        End If
ElseIf KeyAscii = 8 Then
Else
    If Len(Trim(MenuBox(Index))) = VLen(Index) And MenuBox(Index).SelLength = 0 Then
        KeyAscii = 0
    ElseIf UpperCase Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End If

End Sub

Private Sub MenuBox_LostFocus(Index As Integer)

If Not VarChanged(Index) Then OrigValue = MenuBox(Index)

If MenuBox(Index) <> OrigValue Then
    If NoAlert Then
        response = vbYes
    Else
        response = MsgBox("Update value of " & VarLabel(Index) & " to '" & MenuBox(Index) & "'", vbYesNo)
    End If
    If response = vbYes Then
        MenuBox_KeyPress Index, 13
    Else
        MenuBox(Index) = OrigValue
    End If
End If

'MenuBox(Index) = OrigValue
VarChanged(Index) = False
' If Trim(MenuBox(Index)) = "" Then
    MenuBox(Index).SelLength = 0
' End If
End Sub

Private Sub MenuBox_Scroll(Index As Integer)

A = 1

End Sub

Private Sub NumberBox_Change(Index As Integer)

If Trim(NumberBox(Index)) <> "" And Not IsNumeric(NumberBox(Index)) Then
    MsgBox ("This field requires numeric input only")
    NumberBox(Index).SelStart = 0
    NumberBox(Index).SelLength = Len(NumberBox(Index))
    Exit Sub
End If

End Sub

Private Sub NumberBox_Click(Index As Integer)

If Trim(NumberBox(Index)) = "" Then
    NumberBox(Index) = ""
End If

End Sub

Private Sub NumberBox_GotFocus(Index As Integer)

OrigValue = NumberBox(Index)
If NumberBox(Index) = "" Then
    NumberBox(Index) = Space(30)
End If
NumberBox(Index).SelStart = 0
NumberBox(Index).SelLength = Len(NumberBox(Index))

End Sub

Private Sub NumberBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

'Select Case KeyCode
'    Case vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd
'        Picture1.SetFocus
'        Picture1_KeyDown KeyCode, 0
'End Select

End Sub

Private Sub NumberBox_KeyPress(Index As Integer, KeyAscii As Integer)

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then Exit Sub
If CountRecords = 0 Then Exit Sub
If Not VarChanged(Index) Then
    OrigValue = NumberBox(Index)
    VarChanged(Index) = True
End If

If KeyAscii = 13 Then
    If OrigValue = NumberBox(Index) Then
        Picture1.SetFocus
    Else
        UpdatePointsTable VarList(Index), NumberBox(Index), 1, 1

        If UCase(VarList(Index)) = UCase(MasterVar) Then
            MasterVal = NumberBox(Index).Text
            FillDependentFields
        Else
            For I = 1 To nDependentVars
                If UCase(DependentVar(I)) = UCase(VarList(Index)) Then
                    UpdateDependentVar DependentVar(I), NumberBox(Index).Text
                    If Cancelling Then
                        MenuBox(Index) = OrigValue
                        Exit Sub
                    End If
                    Exit For
                End If
            Next I
        End If

        CheckFields
    End If
ElseIf KeyAscii = 27 Then
    NumberBox(Index) = OrigValue
    Picture1.SetFocus
Else
    Select Case KeyAscii
        Case 8, 46, 48 To 57, Asc("-"), Asc(".")
        Case Else
            KeyAscii = 0
    End Select
End If

End Sub

Private Sub NumberBox_LostFocus(Index As Integer)

If Not VarChanged(Index) Then OrigValue = NumberBox(Index)
If NumberBox(Index) <> OrigValue Then
    If NoAlert Then
        response = vbYes
    Else
        response = MsgBox("Update value of " + VarLabel(Index) + " to '" + NumberBox(Index) + "'", vbYesNo)
    End If
    If response = vbYes Then
        NumberBox_KeyPress Index, 13
    Else
        NumberBox(Index) = OrigValue
    End If
End If
VarChanged(Index) = False

End Sub

Private Sub TextBox_Click(Index As Integer)

If Trim(TextBox(Index)) = "" Then
    TextBox(Index) = ""
End If

End Sub

Private Sub TextBox_GotFocus(Index As Integer)

OrigValue = TextBox(Index)
If TextBox(Index) = "" Then
    TextBox(Index) = Space(30)
End If
TextBox(Index).SelStart = 0
TextBox(Index).SelLength = Len(TextBox(Index))

End Sub

Private Sub TextBox_KeyPress(Index As Integer, KeyAscii As Integer)

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then Exit Sub
If CountRecords = 0 Then Exit Sub

If Not VarChanged(Index) Then
    OrigValue = TextBox(Index)
    VarChanged(Index) = True
End If

If KeyAscii = 13 And Trim(TextBox(Index)) <> "" Then
    If OrigValue = TextBox(Index) Then
        Picture1.SetFocus
    Else
        UpdatePointsTable VarList(Index), TextBox(Index), 1, 1
        
        If UCase(VarList(Index)) = UCase(MasterVar) Then
            MasterVal = TextBox(Index).Text
            FillDependentFields
        Else
            For I = 1 To nDependentVars
                If UCase(DependentVar(I)) = UCase(VarList(Index)) Then
                    UpdateDependentVar DependentVar(I), TextBox(Index).Text
                    If Cancelling Then
                        TextBox(Index) = OrigValue
                        Cancelling = False
                        Exit Sub
                    End If
                    Exit For
                End If
            Next I
        End If
        CheckFields
        OrigValue = TextBox(Index)
        If GridShowing Then frmDataGrid.DataGrid.Refresh
    End If
ElseIf KeyAscii = 27 Then
    TextBox(Index) = OrigValue
    Picture1.SetFocus
ElseIf KeyAscii = 8 Then

Else
    If Len(Trim(TextBox(Index))) >= VLen(Index) And TextBox(Index).SelLength = 0 Then
        KeyAscii = 0
    ElseIf UpperCase Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End If

End Sub

Private Sub TextBox_LostFocus(Index As Integer)

If Trim(TextBox(Index)) = "" Then TextBox(Index) = ""
If Not VarChanged(Index) Then OrigValue = TextBox(Index)

If Len(TextBox(Index)) > VLen(Index) Then
    MsgBox ("Length of " & VarList(Index) & " set to maximum of " & VLen(Index))
    TextBox(Index).SelStart = 0
    TextBox(Index).SelLength = Len(TextBox(Index))
    Exit Sub
End If
If UCase(VarList(Index)) = "DATE" Or UCase(VarList(Index)) = "TIME" Then
    If Not IsDate(TextBox(Index)) And Trim(TextBox(Index)) <> "" Then
        MsgBox ("Invalid date.  Re-enter")
        TextBox(Index).SetFocus
    End If
End If
If TextBox(Index) <> OrigValue Then
    If NoAlert Then
        response = vbYes
    Else
        response = MsgBox("Update value of " & VarLabel(Index) & " to '" & TextBox(Index) & "'", vbYesNo)
    End If
    If response = vbYes Then
        TextBox_KeyPress Index, 13
    Else
        TextBox(Index) = OrigValue
    End If
End If
VarChanged(Index) = False

End Sub

Private Sub txtHangle_KeyDown(KeyCode As Integer, Shift As Integer)

'Select Case KeyCode
'    Case vbKeyUp, vbKeyPageUp, vbKeyPageDown, vbKeyDown, vbKeyHome, vbKeyEnd, vbKeyDelete, vbKeyRight, vbKeyLeft
'        Picture1.SetFocus
'End Select

End Sub

Private Sub txtID_Click()

Dim NewID As String
Dim TempSuffix As Integer
Dim OldID As String

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then
    MsgBox "Open point table first.", vbInformation
    txtID = OriginalID
    Exit Sub
End If

If Not Loading Then
    If txtID = OriginalID Then Exit Sub
    
    NewID = txtID
    If Val(txtSuffix) > 0 Then
        response = MsgBox("Change ID for all points in this series?", vbYesNoCancel)
        If response = vbYes Then
            UpdatePointsTable "id", NewID, 1, 1
            If Val(Trim(txtID)) = 0 Then
                OldID = OriginalID
                DecrementID OriginalUnit, OriginalID, OriginalSuffix
                OriginalID = OldID
            End If
        ElseIf response = vbCancel Then
            txtID = OriginalID
            Exit Sub
        Else
            UpdatePointsTable "id", NewID, 1, 1
            ADOAccess SetID, txtUnit, NewID, "", ""
            If Not PointsADO.recordset.EOF Then
                With PointsADO.recordset
                    currentrecord = .Bookmark
                    .MoveFirst
                    TempSuffix = 0
                    Do
                        If ("unit") = txtUnit And ("id") = OriginalID Then
                            .Update "suffix", TempSuffix
                        End If
                        .MoveNext
                        TempSuffix = TempSuffix + 1
                    Loop Until .EOF
                    If Not NoAlert Then
                        MsgBox TempSuffix + 1 & " records changed.", vbInformation
                    End If
                    .Bookmark = currentrecord
                End With
                UpdatePointsTable "id", NewID, 1, 0
                UpdatePointsTable "suffix", txtSuffix, 0, 0
                txtSuffix = 0
                OriginalID = Val(NewID)
            End If
        End If
    Else
        If Val(Trim(OriginalID)) > 0 Then
            OldID = OriginalID
            DecrementID OriginalUnit, OriginalID, OriginalSuffix
            OriginalID = OldID
'           ADOAccess SetID, OriginalUnit, NewID, "", ""
'            If Val(Trim(OriginalID)) < Val(Trim(NewID)) Then
'                response = MsgBox("Reset last ID? to " & NewID & "?", vbYesNo)
'                If response = vbYes Then
'                    OldID = OriginalID
'                    DecrementID OriginalUnit, OriginalID, OriginalSuffix
'                    OriginalID = OldID
'                End If
'            End If
        End If
        PointsADO.recordset.Update "ID", PadID(NewID)
        UpdatePointsTable "id", NewID, 1, 0
        'ADOAccess SetID, txtUnit, NewID, "", ""
        'PointsADO.Recordset.Filter = "unit='" + txtUnit + "' and id='" + OriginalID + "' and suffix>0"
        If Not PointsADO.recordset.EOF Then
            CurrentPosition = PointsADO.recordset.Bookmark
            GridLoading = True
            PointsADO.recordset.MoveFirst
            TempSuffix = 0
            While Not PointsADO.recordset.EOF
                If PointsADO.recordset("unit") = txtUnit And PointsADO.recordset("id") = OriginalID Then
                    TempSuffix = TempSuffix + 1
                    If TempSuffix = 1 Then
                        MsgBox ("Resequencing subsequent shot(s) from " + txtUnit + "-" + Trim(OriginalID) + " as continuation shots for " + txtUnit + "-" + Trim(txtID))
                    End If
                    PointsADO.recordset.Update "id", txtID
                    PointsADO.recordset.Update "suffix", TempSuffix
                End If
                PointsADO.recordset.MoveNext
            Wend
            MsgBox TempSuffix + 1 & " records changed", vbInformation
            PointsADO.recordset.Bookmark = CurrentPosition
        End If
        GridLoading = False
        'PointsADO.Recordset.Filter = adFilterNone

    End If
    'DoEvents
    txtID = NewID
    OriginalID = Val(txtID)
    OriginalSuffix = Val(txtSuffix)
    txtID.Refresh
    ShowValues
    Picture1.SetFocus
End If

End Sub

Private Sub txtID_DropDown()

Dim TempString As String
TempString = txtID
txtID.Clear
txtID = TempString
If Trim(txtID = "") Or Val(txtSuffix) > 0 Then
    SqlString = "select max(id) from [EDM_units] where unit='" + txtUnit + "' and id<'A'"
    Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
    If IsNull(rsTemp(0)) Then
        txtID.AddItem PadID("1")
    Else
        txtID.AddItem PadID(Str(rsTemp(0) + 1))
    End If
    txtID.AddItem hash(5)
ElseIf Val(Trim(txtID)) > 0 Then
    txtID.AddItem hash(5)
Else
    SqlString = "select max(id) from [EDM_units] where unit='" + txtUnit + "' and id<'A'"
    Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
    If IsNull(rsTemp(0)) Then
        txtID.AddItem PadID("1")
    Else
        txtID.AddItem PadID(Str(rsTemp(0) + 1))
    End If
End If
Set rsTemp = Nothing

End Sub

Private Sub txtID_GotFocus()
    OriginalUnit = txtUnit
    OriginalID = txtID
    OriginalSuffix = Val(txtSuffix)
    IDChanged = False
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
    Picture1.SetFocus
    txtID = OriginalID
ElseIf KeyAscii = 13 Then
    If txtID <> OriginalID Then
        response = MsgBox("Update ID to '" & txtID & "'?", vbYesNo)
        If response = vbNo Then
            txtID = OriginalID
            Exit Sub
        End If
    End If
    txtID = PadID(txtID)
    Loading = False
    txtID_Click
    
'    If DupID Then
'        UpdatePointsTable "id", txtID, 0, 0
'        UpdatePointsTable "suffix", txtSuffix, 0, 0
'        If Val(Trim(txtID)) = 0 Then
'            DecrementID OriginalUnit, OriginalID, OriginalSuffix
'        End If
'    Else
'    End If
Else
    Loading = False
    If Not IDChanged Then
        OriginalID = txtID
        IDChanged = True
    End If
    If UpperCase Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
'    KeyAscii = 0
End If

End Sub

Private Sub txtID_LostFocus()

If Not IDChanged Then OriginalID = txtID

If txtID <> OriginalID Then
        txtID_KeyPress 13
End If
IDChanged = False

End Sub

Private Sub txtPoleHT_KeyPress(KeyAscii As Integer)

'If KeyAscii = 27 Then
'    Picture1.SetFocus
'End If

End Sub

Private Sub txtprism_Click()

If txtprism.ListIndex > -1 Then
    txtPoleHT = Format(PoleHeight(txtprism.ItemData(txtprism.ListIndex)), "####0.000")
Else
    Exit Sub
End If
If Not Loading Then
    txtprism_KeyPress 13
End If

End Sub

Private Sub txtPrism_GotFocus()

OriginalPoleHT = Val(txtPoleHT)
OriginalPrismIndex = txtprism.ListIndex
txtprism.SelStart = 0
txtprism.SelLength = Len(txtprism)
Loading = False

End Sub

Private Sub txtprism_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
    Picture1.SetFocus
ElseIf KeyAscii = 13 Then
    If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then Exit Sub
    On Error GoTo ExitSub
    If Not Loading And CountRecords > 0 Then
        If PoleHeight(txtprism.ItemData(txtprism.ListIndex)) <> OriginalPoleHT Then
            txtPoleHT = Format(PoleHeight(txtprism.ItemData(txtprism.ListIndex)), "#####0.000")
            txtXYZ(2) = Format(Val(txtXYZ(2)) + OriginalPoleHT - PoleHeight(txtprism.ItemData(txtprism.ListIndex)), "#####0.000")
            UpdatePointsTable "prism", txtPoleHT, 0, 0
            UpdatePointsTable "z", txtXYZ(2), 1, 0
        End If
    End If
    txtPoleHT = PoleHeight(txtprism.ItemData(txtprism.ListIndex))
    OriginalPoleHT = txtPoleHT
    OriginalPrismIndex = txtprism.ListIndex
    For I = 1 To nUnitFields
        If UCase(Unitfield(I)) = "PRISM" Then
            UpdateUnitTable txtUnit, txtID, UnitFieldOnly
            Exit For
        End If
    Next I
Else
     KeyAscii = 0
End If

ExitSub:

End Sub

Private Sub txtPrism_LostFocus()

If Loading Then Exit Sub
txtprism.SelLength = 0
txtPoleHT = Format(OriginalPoleHT, "####0.000")
'If txtprism.ListCount > 0 Then
'    txtprism.ListIndex = OriginalPrismIndex
'End If

End Sub

Private Sub txtPT_Click()

If txtPT <> "" Then
    PointTableName = txtPT
    OpenPointsTable
End If
Picture1.SetFocus

End Sub

Private Sub txtPT_DropDown()

If SiteDBname <> "" Then
    txtPT.Clear
    GetPointTables
    For I = 1 To nPointTables
        txtPT.AddItem PointTable(I)
    Next I
End If

End Sub

Private Sub txtPT_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
    Picture1.SetFocus
Else
    KeyAscii = 0
End If

End Sub

Private Sub txtPT_LostFocus()

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then
    lblPointsWarning.Visible = True
Else
    txtPT = PointTableName
    txtXYZ(0).Enabled = True
    txtXYZ(1).Enabled = True
    txtXYZ(2).Enabled = True
    frmMain.txtUnit.Enabled = True
    frmMain.txtID.Enabled = True
    frmMain.txtprism.Enabled = True
End If

End Sub

Private Sub txtSlopeD_KeyDown(KeyCode As Integer, Shift As Integer)

'Select Case KeyCode
'    Case vbKeyUp, vbKeyPageUp, vbKeyPageDown, vbKeyDown, vbKeyHome, vbKeyEnd, vbKeyDelete, vbKeyRight, vbKeyLeft
'        Picture1.SetFocus
'End Select

End Sub

Private Sub txtSuffix_GotFocus()

txtSuffix.SelStart = 0
txtSuffix.SelLength = Len(txtSuffix)
OriginalSuffix = Val(txtSuffix)

End Sub

Public Sub ShowValues()

'Exit Sub

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then Exit Sub

If PointsADO.recordset.BOF And PointsADO.recordset.EOF Then
    txtCurrentRecord = 0
    txtTotalRecords = 0
    ClearFields
    
    Exit Sub
End If

Loading = True
txtCurrentRecord = ""

'If IsEmpty(CurrentBookMark) Then
'    PointsADO.Recordset.MoveLast
'    CurrentBookMark = PointsADO.Recordset.Bookmark
'Else
'    PointsADO.Recordset.Bookmark = CurrentBookMark
'End If

txtCurrentRecord = PointsADO.recordset.AbsolutePosition
txtTotalRecords = CountRecords
Set rsTemp = Nothing
'txtTotalRecords = countrecords
If CountRecords > 0 Then
    txtID.Enabled = True
    txtSuffix.Enabled = True
    txtXYZ(0).Enabled = True
    txtXYZ(1).Enabled = True
    txtXYZ(2).Enabled = True
    txtVangle.Enabled = True
    txtHangle.Enabled = True
    txtSloped.Enabled = True
    txtprism.Enabled = True
    On Error Resume Next
    For I = 1 To Vars
        TextBox(I).Enabled = True
        MenuBox(I).Enabled = True
        NumberBox(I).Enabled = True
    Next I
    On Error GoTo 0
    
    lblBlankFields.Visible = False
    If PointsADO.recordset.BOF Then
        PointsADO.recordset.MoveFirst
    ElseIf PointsADO.recordset.EOF Then
        PointsADO.recordset.MoveLast
    End If
    If IsNull(PointsADO.recordset("unit")) Then
        txtUnit = ""
    Else
        txtUnit = PointsADO.recordset("Unit")
    End If
    If IsNull(PointsADO.recordset("id")) Then
        txtID = ""
    Else
        txtID = PointsADO.recordset("id")
    End If
    If IsNull(PointsADO.recordset("suffix")) Then
        txtSuffix = 0
    Else
        txtSuffix = PointsADO.recordset("suffix")
    End If

    OriginalID = txtID
    OriginalUnit = txtUnit
    OriginalSuffix = Val(txtSuffix)
    'doevents
    If txtXYZ(0).Visible Then txtXYZ(0) = Format(PointsADO.recordset("x"), "#########0.000")
    If txtXYZ(1).Visible Then txtXYZ(1) = Format(PointsADO.recordset("y"), "#########0.000")
    If txtXYZ(2).Visible Then txtXYZ(2) = Format(PointsADO.recordset("z"), "#########0.000")
    If txtVangle.Visible Then txtVangle = Format(PointsADO.recordset("vangle"), "#########0.0000")
    If txtHangle.Visible Then txtHangle = Format(PointsADO.recordset("hangle"), "#########0.0000")
    If txtSloped.Visible Then txtSloped = Format(PointsADO.recordset("sloped"), "#########0.0000")
    txtXYZ(0).Refresh
    
    If txtPoleHT.Visible Then
        If field_in_recordset(PointsADO.recordset, "prism") Then
            OriginalPoleHT = PointsADO.recordset("prism")
            txtPoleHT = Format(PointsADO.recordset("prism"), "#####0.000")
            For I = 0 To txtprism.ListCount - 1
                If PoleHeight(txtprism.ItemData(I)) = PointsADO.recordset("prism") Then
                    txtprism.ListIndex = I
                    Exit For
                End If
            Next I
        Else
            OriginalPoleHT = 0
            txtPoleHT = 0
            txtprism = ""
            
        End If
    End If
    On Error Resume Next
    For I = 1 To Vars
        
        Select Case VType(I)
            Case "TEXT"
                If IsNull(PointsADO.recordset(VarList(I))) Then
                    TextBox(I) = ""
                Else
                    TextBox(I).Text = PointsADO.recordset(VarList(I))
                End If
            Case "MENU"
                If IsNull(PointsADO.recordset(VarList(I))) Then
                    MenuBox(I) = ""
                Else
                    MenuBox(I) = PointsADO.recordset(VarList(I))
                End If
            Case "NUMERIC", "INSTRUMENT"
                If IsNull(PointsADO.recordset(VarList(I))) Then
                    NumberBox(I) = ""
                Else
                    NumberBox(I) = PointsADO.recordset(VarList(I))
                End If
        End Select
    Next I
    On Error GoTo 0
    If mdiMain.mnuViewPoints.Checked Then
        frmPlot.shpPoint.Left = PointsADO.recordset(PlotX) - frmPlot.shpPoint.Width / 2
        frmPlot.shpPoint.Top = PointsADO.recordset(PlotY) + frmPlot.shpPoint.Height / 2
        frmPlot.Caption = txtUnit + "-" + Trim(txtID)
    End If
    Loading = False
    CheckFields
Else
    txtID.Enabled = False
    txtSuffix.Enabled = False
    txtXYZ(0).Enabled = False
    txtXYZ(1).Enabled = False
    txtXYZ(2).Enabled = False
    txtVangle.Enabled = False
    txtHangle.Enabled = False
    txtSloped.Enabled = False
    txtprism.Enabled = False
    On Error Resume Next
    For I = 1 To Vars
        TextBox(I).Enabled = False
        MenuBox(I).Enabled = False
        NumberBox(I).Enabled = False
    Next I
    On Error GoTo 0
    txtUnit = ""
    txtID = ""
    txtSuffix = ""
    OriginalID = ""
    OriginalUnit = ""
    OriginalSuffix = 0
    If txtXYZ(0).Visible Then txtXYZ(0) = ""
    If txtXYZ(1).Visible Then txtXYZ(1) = ""
    If txtXYZ(2).Visible Then txtXYZ(2) = ""
    If txtVangle.Visible Then txtVangle = ""
    If txtHangle.Visible Then txtHangle = ""
    If txtSloped.Visible Then txtSloped = ""
    On Error Resume Next
    For I = 1 To Vars
        Select Case VType(I)
            Case "TEXT"
                If VarList(I) = "DATE" Then
                    TextBox(I) = Date
                ElseIf VarList(I) = "TIME" Then
                    TextBox(I) = Time
                Else
                    TextBox(I) = ""
                End If
            Case "MENU"
                MenuBox(I) = ""
            Case "NUMERIC", "INSTRUMENT"
                NumberBox(I) = ""
        End Select
    Next I
    On Error GoTo 0
    Loading = False

End If
'If GridShowing Then
'    frmDataGrid.MoveGrid
'End If
'If Not GridShowing Then
'    Picture1.SetFocus
'End If

End Sub

Private Sub txtSuffix_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
    Picture1.SetFocus
Else
    Select Case KeyAscii
        Case 8, 48 To 57
        Case 13
            If Val(txtSuffix) <> OriginalSuffix Then
                
                
                GridLoading = True
                UpdatePointsTable "suffix", Val(txtSuffix), 1, 0
'                If Val(txtSuffix) <> 0 Then
'                    PointsADO.Recordset.MovePrevious
'                    If Not PointsADO.Recordset.BOF Then
'                        If Val(txtSuffix) < PointsADO.Recordset("suffix") Then
'                            MsgBox ("Cannot reverse Suffix sequence -- reenter")
'                            txtSuffix = OriginalSuffix
'                            Exit Sub
'                        End If
'                        PointsADO.Recordset.MoveNext
'                        UpdatePointsTable "suffix", Val(txtSuffix), 1, 0
'                    Else
'                        MsgBox ("First record on a point must have suffix=0")
'                        txtSuffix = 0
'                        Exit Sub
'                    End If
'                Else
'                    MsgBox ("First record on a point must have suffix=0")
'                    txtSuffix = 0
'                    UpdatePointsTable "suffix", Val(txtSuffix), 1, 0
'                    Exit Sub
'
'                End If
                GridLoading = False
            End If
        Case Else
            KeyAscii = 0
    End Select
    ' txtSuffix_GotFocus
End If

End Sub

Private Sub txtSuffix_LostFocus()

If Val(txtSuffix) <> OriginalSuffix Then
    txtSuffix_KeyPress 13
End If

End Sub

Private Sub txtUnit_Click()

Loading = False
If OriginalUnit = txtUnit Then Exit Sub

If Val(txtID) > 0 Then
    DecrementID OriginalUnit, OriginalID, OriginalSuffix
End If

txtUnit_KeyPress 13

End Sub

Private Sub txtUnit_DropDown()

txtUnit = OriginalUnit
If UnitTB.RecordCount > 0 Then
    txtUnit.Clear
    UnitTB.MoveFirst
    While Not UnitTB.EOF
        txtUnit.AddItem UnitTB("Unit")
        UnitTB.MoveNext
    Wend
End If
        
End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)

Dim NextID As String
Dim rsLock As recordset

If Loading Then Exit Sub

If KeyAscii = 27 Then
    Picture1.SetFocus

ElseIf KeyAscii = 13 Or KeyAscii = 9 Then
    If txtUnit = "" Then Exit Sub
    If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then Exit Sub
    
    If CountRecords = 0 Then
        OriginalUnit = txtUnit
        Exit Sub
    End If
    
    If Not Loading And CountRecords > 0 And PointTableName <> "" Then
            If txtSuffix = 0 And Not PointsADO.recordset.EOF And Not PointsADO.recordset.BOF Then
                If InStr(txtUnit, "-") > 0 Then
                    txtID = PadID(Mid(txtUnit, InStr(txtUnit, "-") + 1))
                    txtUnit = Left(txtUnit, InStr(txtUnit, "-") - 1)
                    PointsADO.recordset.Update "id", txtID
                    
                Else
                    ADOAccess GetNextID, txtUnit, NextID, "", ""
                    txtID = NextID
                    PointsADO.recordset.Update "id", PadID(NextID)
                    
                End If
                PointsADO.recordset.Update "unit", txtUnit
                OriginalUnit = txtUnit
                OriginalID = txtID
                OriginalSuffix = Val(txtSuffix)
                Screen.MousePointer = 1
                FillUnitFields
                ShowValues
                CheckFields
            ElseIf txtSuffix > 0 Then
                MsgBox ("You cannot change unit within a series")
                txtUnit = OriginalUnit
                Exit Sub
            End If
    ElseIf CountRecords = 0 Then
        OriginalUnit = txtUnit
    End If
    Picture1.SetFocus
Else
    KeyAscii = 0
End If

End Sub

Private Sub txtUnit_LostFocus()

txtUnit = OriginalUnit
txtID = OriginalID
txtSuffix = OriginalSuffix

End Sub

Private Sub txtVangle_KeyDown(KeyCode As Integer, Shift As Integer)

'Select Case KeyCode
'    Case vbKeyUp, vbKeyPageUp, vbKeyPageDown, vbKeyDown, vbKeyHome, vbKeyEnd, vbKeyDelete, vbKeyRight, vbKeyLeft
'        Picture1.SetFocus
'End Select

End Sub

Public Sub UpdatePointsTable(Variable As String, Value As String, ShowMessage As Byte, Scope As Byte)

Dim CurrentPosition
Dim NRecords As Integer

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then Exit Sub
If PointsADO.recordset.EOF And PointsADO.recordset.BOF Then Exit Sub

'GridLoading = True
'PointsADO.Recordset.Requery
'PointsADO.Refresh

'***PointsADO.Recordset.Bookmark = CurrentBookMark

'***CurrentPosition = CurrentBookMark

'If CountRecords = 0 Then Exit Sub

If Scope = 0 Or OriginalID = "" Then
    '***PointsADO.Recordset.Bookmark = CurrentBookMark
    If UCase(Variable) = "ID" Then
        PointsADO.recordset.Update Variable, PadID(Value)
    Else
        PointsADO.recordset.Update Variable, Value
    End If
    If UCase(Variable) = "ID" Then
        UpdateUnitTable txtUnit, PadID(Value), IDonly
    End If
    NRecords = 1
Else
    GridLoading = True
    CurrentPosition = PointsADO.recordset.Bookmark
    PointsADO.recordset.MoveFirst
    'PointsADO.Recordset.Filter = "unit='" + txtUnit + "' and id='" + OriginalID + "'"
    While Not PointsADO.recordset.EOF
        If PointsADO.recordset("unit") = txtUnit And PointsADO.recordset("id") = OriginalID Then
            NRecords = NRecords + 1
            If UCase(Variable) = "ID" Then
                PointsADO.recordset.Update Variable, PadID(Value)
                UpdateUnitTable txtUnit, PadID(Value), IDonly
            Else
                On Error Resume Next
                PointsADO.recordset.Fields(Variable) = Value
                PointsADO.recordset.Update
                If Err.Number <> 0 Then
                    PointsADO.recordset.CancelUpdate
                    MsgBox "could not update"
                End If
                'PointsADO.Recordset.Update Variable, Value
            End If
        End If
        PointsADO.recordset.MoveNext
    Wend
    '***PointsADO.Recordset.Filter = adFilterNone
    PointsADO.recordset.Bookmark = CurrentPosition
    GridLoading = False
    
    '***PointsADO.Recordset.Bookmark = CurrentBookMark
    For I = 1 To nUnitFields
        If UCase(Variable) = UCase(Unitfield(I)) Then
            UpdateUnitTable txtUnit, txtID, UnitFieldOnly
            Exit For
        End If
    Next I
End If

'***PointsADO.Recordset.Bookmark = CurrentBookMark
If ShowMessage = 1 And Not NoAlert And Not ChangingXYZ Then
    MsgBox (NRecords & " shot(s) on " + txtUnit + "-" + Trim(txtID) + " changed")
End If
ChangingXYZ = False
'Set rsTemp = Nothing

End Sub

Public Sub DecrementID(Unit As String, ID As String, Suffix As Integer)

Dim MaxID As Long

If Val(ID) = 0 Then Exit Sub
ADOAccess DecID, Unit, ID, "", ""

End Sub

Public Sub Take_Shot(NewObj As Boolean)

Dim edmpoffset As Single

'This has to happen first so that VHDTONEZ can have these values
If txtprism.ListIndex <> -1 Then
    edmshot.poleh = PoleHeight(txtprism.ItemData(txtprism.ListIndex))
    edmshot.poleo = PoleOffset(txtprism.ItemData(txtprism.ListIndex))
Else
    edmshot.poleh = 0
    edmshot.poleo = 0
End If

cmdShoot.Enabled = False
cmdPlusShot.Enabled = False
Command1.Enabled = False
For I = 1 To 6
    Button(I).Enabled = False
Next I
cmdCancel.Visible = True
Picture1.SetFocus
takeshot_core NoPrism
cmdCancel.Visible = False

If Cancelling Then
    mdiMain.StatusBar.Panels(6).Visible = False
    cmdShoot.Enabled = True
    cmdPlusShot.Enabled = True
    Command1.Enabled = True
    For I = 1 To 6
        Button(I).Enabled = True
    Next I
    Picture1.SetFocus
    GoTo ReEnable
End If
txtXYZ(0) = Format(edmshot.X, "######0.000")
txtXYZ(1) = Format(edmshot.y, "######0.000")
txtXYZ(2) = Format(edmshot.z, "######0.000")

txtVangle = edmshot.vangle
txtHangle = edmshot.hangle
txtSloped = edmshot.sloped

If mdiMain.mnuPrismPrompt.Checked Then
    MsgBox ("Verify Prism.  Currently set to " + txtprism.List(txtprism.ListIndex))
End If

GridLoading = True
'PointsADO.Recordset.Requery
If PointsADO.recordset.BOF And PointsADO.recordset.EOF Then
Else
    PointsADO.recordset.MoveLast
End If
'On Error Resume Next
On Error GoTo 0
Do
    PointsADO.recordset.AddNew
    PointsADO.recordset("unit") = " "
Loop Until Err.Number = 0

'On Error GoTo 0
PointsADO.recordset.Update
'CurrentBookMark = PointsADO.Recordset.Bookmark
PointsADO.recordset.MoveLast
PointsADO.recordset.Update "prism", edmshot.poleh
PointsADO.recordset.Update "suffix", txtSuffix
If txtXYZ(0).Visible Then PointsADO.recordset.Update "x", txtXYZ(0)
If txtXYZ(1).Visible Then PointsADO.recordset.Update "y", txtXYZ(1)
If txtXYZ(2).Visible Then PointsADO.recordset.Update "z", txtXYZ(2)
If txtVangle.Visible Then PointsADO.recordset.Update "vangle", txtVangle
If txtHangle.Visible Then PointsADO.recordset.Update "hangle", txtHangle
If txtSloped.Visible Then PointsADO.recordset.Update "sloped", txtSloped
If DatumInfo Then
    PointsADO.recordset.Update "DatumName", CurrentStation.Name
End If

On Error GoTo 0
'On Error GoTo Boxerror
For I = 1 To Vars
    Select Case UCase(VarList(I))
        Case "UNIT", "ID", "SUFFIX", "PRISM", "X", "Y", "Z", "VANGLE", "HANGLE", "SLOPED"
        Case "DATUMX"
            PointsADO.recordset.Update VarList(I), Format(CurrentStation.X, "######0.000")
        Case "DATUMY"
            PointsADO.recordset.Update VarList(I), Format(CurrentStation.y, "######0.000")
        Case "DATUMZ"
            PointsADO.recordset.Update VarList(I), Format(CurrentStation.z, "######0.000")
        Case "DATE"
            PointsADO.recordset.Update VarList(I), Date
        Case "TIME"
            PointsADO.recordset.Update VarList(I), Time
        Case Else
        Select Case VType(I)
            Case "TEXT"
                If VCarry(I) Or Not NewObj Then
                    If IsNull(TextBox(I)) Or Len(TextBox(I)) = 0 Then
                        PointsADO.recordset.Update VarList(I), " "
                    Else
                        PointsADO.recordset.Update VarList(I), TextBox(I).Text
                    End If
                End If
            Case "MENU"
                If VCarry(I) Or Not NewObj Then
                    If Trim(MenuBox(I)) <> "" Then PointsADO.recordset.Update VarList(I), MenuBox(I)
                End If
            Case "NUMERIC", "INSTRUMENT"
                If (VCarry(I) Or Not NewObj) And IsNumeric(NumberBox(I)) Then
                    PointsADO.recordset.Update VarList(I), NumberBox(I)
                End If
        End Select
        PointsADO.recordset.Update
    End Select
Continue:
Next I
On Error GoTo 0
PointsADO.recordset.Update
ReEnable:
GridLoading = False
Exit Sub

errorhandler:
    MsgBox ("Error when writing " + VarList(I) + ": " + Err.Description)
    Resume Next
    
Boxerror:
Resume Continue

End Sub

Public Sub DeleteRecord()

Dim NDeletions As Integer
Dim PreviousRecord As Integer
Dim BkMk
Dim NextRecno As Integer
Dim Data1BookMark As Variant
Dim NextBkMk As Variant
Dim OldUnit As String
Dim OldID As String
response = MsgBox("Warning:  Deleting will permanently remove records.  Continue anyway?", vbYesNo)

If response = vbNo Then Exit Sub
If GridShowing Then GridLoading = True

'PointsADO.Recordset.Bookmark = CurrentBookMark
'Do
'    PointsADO.Recordset.MovePrevious
'    If PointsADO.Recordset.BOF Then Exit Do
'Loop Until (PointsADO.Recordset("unit") <> txtUnit Or PointsADO.Recordset("id") <> txtID)

'If PointsADO.Recordset.BOF Then
'    PointsADO.Recordset.Bookmark = CurrentBookMark
'    Do
'        PointsADO.Recordset.MoveNext
'        If PointsADO.Recordset.EOF Then Exit Do
'    Loop Until PointsADO.Recordset("unit") <> txtUnit Or PointsADO.Recordset("id") <> txtID
'    If PointsADO.Recordset.EOF Then
'            NextBkMk = Empty
'            'ClearFields
'            'CurrentBookMark = Empty
'            'Exit Sub
'    Else
'        NextBkMk = PointsADO.Recordset.Bookmark
'    End If
'Else
'    NextBkMk = PointsADO.Recordset.Bookmark
'End If

Need2Decrement = False
OldUnit = txtUnit
OldID = txtID
ADOAccess DelRec, txtUnit, txtID, "", ""


If Not PointsADO.recordset.BOF Or Not PointsADO.recordset.EOF Then
    
    Do
        On Error Resume Next
        t$ = PointsADO.recordset("X")
        If Err.Number = 0 Then Exit Do
        PointsADO.recordset.MoveNext
    Loop Until PointsADO.recordset.EOF
    On Error GoTo 0
    
    If PointsADO.recordset.EOF Then
            On Error Resume Next
            PointsADO.recordset.MoveLast
            Do
                On Error Resume Next
                t$ = PointsADO.recordset("X")
                If Err.Number = 0 Then Exit Do
                PointsADO.recordset.MovePrevious
            Loop Until PointsADO.recordset.BOF
            On Error GoTo 0
    End If
            
    ShowValues

End If

If Need2Decrement Then
    DecrementID OldUnit, OldID, 0
End If


'PointsADO.Recordset.Requery
'PointsADO.Refresh
'If NextBkMk = Empty Then
'    ClearFields
'Else
'    CurrentBookMark = NextBkMk
'    PointsADO.Recordset.Bookmark = CurrentBookMark
'    ShowValues
'End If

GridLoading = False

Screen.MousePointer = 1
Picture1.SetFocus

End Sub

Private Sub txtXYZ_Click(Index As Integer)

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then
    MsgBox "Open point table first.", vbInformation
    Select Case Index
    Case 0
        txtXYZ(Index) = Format(PointsADO.recordset("X"), "#####0.000")
    Case 1
        txtXYZ(Index) = Format(PointsADO.recordset("Y"), "#####0.000")
    Case 2
        txtXYZ(Index) = Format(PointsADO.recordset("Z"), "#####0.000")
    End Select
    Exit Sub
End If

If LCase(Left(txtXYZ(Index), 6)) = "offset" Then
    frmOffSet.Hide
    frmOffSet.Text1 = ""
    frmOffSet.Caption = txtXYZ(Index)
    Select Case Index
    Case 0
        frmOffSet.OriginalXYZ = Format(PointsADO.recordset("x"), "####0.000")
    Case 1
        frmOffSet.OriginalXYZ = Format(PointsADO.recordset("y"), "####0.000")
    Case 2
        frmOffSet.OriginalXYZ = Format(PointsADO.recordset("z"), "####0.000")
    End Select
    'txtXYZ(Index).Clear
    'txtXYZ(Index) = Format(OriginalXYZ(Index), "####0.000")
    Set frmOffSet.CallingBox = txtXYZ(Index)
    Select Case Index
    Case 0
        frmOffSet.Varname = "X"
    Case 1
        frmOffSet.Varname = "Y"
    Case 2
        frmOffSet.Varname = "Z"
    End Select
    Loading = True
    Cancelling = False
    frmOffSet.Show 1
    Loading = False
    txtXYZ(Index) = Format(OffsetValue, "####0.000")
    Select Case Index
        Case 0
            UpdatePointsTable "x", txtXYZ(Index), 1, 0
        Case 1
            UpdatePointsTable "y", txtXYZ(Index), 1, 0
        Case 2
            UpdatePointsTable "z", txtXYZ(Index), 1, 0
    End Select
    'DoEvents
    
    Loading = False
    Picture1.SetFocus
    'ShowValues
End If

If Val(txtSuffix) = 0 Then
    Dim rsTemp As recordset
    OriginalUnit = txtUnit
    SqlString = "select * from [EDM_units] where minx<= " & PointsADO.recordset("x") & " and maxx>" & PointsADO.recordset("x") & " and miny<=" & PointsADO.recordset("y") & " and maxy>" & PointsADO.recordset("y")
    Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("unit")) Then
            If OriginalUnit <> rsTemp("unit") Then
                GoSub ChangeUnit
            End If
        End If
    Else
        SqlString = "select * from [EDM_units] where abs(centerx-" & PointsADO.recordset("x") & ")<=radius and abs(centery-" & PointsADO.recordset("y") & ")<=radius"
        Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
        If Not rsTemp.EOF Then
            If Sqr((rsTemp("centerx") - PointsADO.recordset("x")) ^ 2 + (rsTemp("centery") - PointsADO.recordset("y")) ^ 2) <= rsTemp("radius") Then
                If Not IsNull(rsTemp("unit")) Then
                    If OriginalUnit <> rsTemp("unit") Then
                        GoSub ChangeUnit
                    End If
                End If
            End If
        End If
    End If
    Set rsTemp = Nothing
End If
Exit Sub

ChangeUnit:
    MsgBox ("Because of the offset, the location of this point is now in unit " + rsTemp("Unit") + ".  Change Unit field from pull-down menu")

'txtXYZ(Index).Refresh

End Sub

Private Sub txtXYZ_DropDown(Index As Integer)

If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then Exit Sub
If PointsADO.recordset.EOF And PointsADO.recordset.BOF Then Exit Sub

txtXYZ(Index).Clear
Select Case Index
    Case 0
        txtXYZ(0).AddItem "Offset East"
        txtXYZ(0).AddItem "Offset West"
        txtXYZ(0).AddItem Format(PointsADO.recordset("x"), "####0.000")
        txtXYZ(0) = Format(PointsADO.recordset("x"), "####0.000")
    Case 1
        txtXYZ(1).AddItem "Offset North"
        txtXYZ(1).AddItem "Offset South"
        txtXYZ(1).AddItem Format(PointsADO.recordset("y"), "####0.000")
        txtXYZ(1) = Format(PointsADO.recordset("y"), "####0.000")
    Case 2
        txtXYZ(2).AddItem "Offset Up"
        txtXYZ(2).AddItem "Offset Down"
        txtXYZ(2).AddItem Format(PointsADO.recordset("z"), "####0.000")
        txtXYZ(2) = Format(PointsADO.recordset("z"), "####0.000")
End Select

End Sub

Private Sub txtXYZ_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 27 Then
    Select Case Index
    Case 0
        txtXYZ(Index) = Format(PointsADO.recordset("X"), "#####0.000")
    Case 1
        txtXYZ(Index) = Format(PointsADO.recordset("Y"), "#####0.000")
    Case 2
        txtXYZ(Index) = Format(PointsADO.recordset("Z"), "#####0.000")
    End Select
    Picture1.SetFocus

ElseIf KeyAscii = 13 Then
    If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then
        MsgBox "Open point table first.", vbInformation
        Select Case Index
        Case 0
            txtXYZ(Index) = Format(PointsADO.recordset("X"), "#####0.000")
        Case 1
            txtXYZ(Index) = Format(PointsADO.recordset("Y"), "#####0.000")
        Case 2
            txtXYZ(Index) = Format(PointsADO.recordset("Z"), "#####0.000")
        End Select
    End If
    txtXYZ(Index) = Format(txtXYZ(Index), "####0.000")
    Select Case Index
    Case 0
        UpdatePointsTable "x", txtXYZ(Index), 1, 0
    Case 1
        UpdatePointsTable "y", txtXYZ(Index), 1, 0
    Case 2
        UpdatePointsTable "z", txtXYZ(Index), 1, 0
    End Select

Else
    Select Case KeyAscii
    Case 8, 48 To 57, Asc("-"), Asc(".")
    Case Else
        KeyAscii = 0
    End Select

End If

End Sub

Public Sub UpdateUnitTable(UnitName As String, ID As String, Scope)

Dim rsTemp As ADODB.recordset

Screen.MousePointer = 11
If Scope <> UnitFieldOnly Then
    If Val(ID) > 0 Then
        ADOAccess SetID, UnitName, ID, "", ""
    End If
End If

NextCheck:
If Scope <> IDonly Then
    For I = 1 To nUnitFields
        Select Case UCase(Unitfield(I))
            Case "ID", "UNIT", "SUFFIX"
            Case "PRISM"
                'UnitTB("PRISM") = txtPoleHT
                SqlString = "Update [edm_units] set [edm_units].prism=" + txtPoleHT + " where [edm_units].unit='" + txtUnit + " '"
                SiteDB.Execute SqlString
            Case Else
                If PointsADO.recordset(Unitfield(I)) = "" Or IsNull(PointsADO.recordset(Unitfield(I))) Then
                    TempString = " "
                Else
                    TempString = PointsADO.recordset(Unitfield(I))
                End If
                SqlString = "Update [edm_units] set edm_units." + Unitfield(I) + "= '" + TempString + "' where edm_units.unit='" + txtUnit + "'"
                SiteDB.Execute SqlString
        End Select
    Next I
End If
        
Screen.MousePointer = 1

End Sub

Public Function FindUnit()

Dim X As Single
Dim y As Single
Dim rsTemp As recordset

If Cancelling Then Exit Function

Start:
Cancelling = False
Loading = False
'PointsADO.Recordset.Bookmark = CurrentBookMark
SqlString = "select * from [EDM_units] where minx<= " & PointsADO.recordset("x") & " and maxx>" & PointsADO.recordset("x") & " and miny<=" & PointsADO.recordset("y") & " and maxy>" & PointsADO.recordset("y")
Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
If Not rsTemp.EOF Then
    If Not IsNull(rsTemp("unit")) Then
        If OriginalUnit <> rsTemp("unit") Then
            txtUnit.ForeColor = QBColor(12)
        Else
            txtUnit.ForeColor = 0
        End If
        OriginalUnit = ""
        txtUnit = rsTemp("unit")
        txtUnit_KeyPress 13
        Exit Function
    End If
Else
    SqlString = "select * from [EDM_units] where abs(centerx-" & PointsADO.recordset("x") & ")<=radius and abs(centery-" & PointsADO.recordset("y") & ")<=radius"
    Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
    If Not rsTemp.EOF Then
        If Sqr((rsTemp("centerx") - PointsADO.recordset("x")) ^ 2 + (rsTemp("centery") - PointsADO.recordset("y")) ^ 2) <= rsTemp("radius") Then
            If Not IsNull(rsTemp("unit")) Then
                If OriginalUnit <> rsTemp("unit") Then
                    txtUnit.ForeColor = QBColor(12)
                    txtUnit.FontBold = True
                Else
                    txtUnit.ForeColor = 0
                    txtUnit.FontBold = True
                End If
                OriginalUnit = ""
                txtUnit = rsTemp("unit")
                txtUnit_KeyPress 13
                Exit Function
            End If
        End If
    End If
End If

If SelectedUnit <> "" Then
    txtUnit = SelectedUnit
    txtUnit_KeyPress 13
    Exit Function
Else
    SqlString = "Select unit from [EDM_units] where minx=-99999"
    Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenDynaset)
    If Not rsTemp.EOF Then
        frmSelectUnit.txtUnit.Clear
        rsTemp.MoveFirst
        While Not rsTemp.EOF
            frmSelectUnit.txtUnit.AddItem rsTemp("unit")
            rsTemp.MoveNext
        Wend
        frmSelectUnit.txtUnit.ListIndex = 0
        frmSelectUnit.Show 1
        If frmSelectUnit.Cancelling Then
            SelectedUnit = ""
            response = vbCancel
        ElseIf frmSelectUnit.UnitSelected <> "" Then
            If frmSelectUnit.DefaultUnit <> "" Then
                SelectedUnit = frmSelectUnit.DefaultUnit
                mdiMain.mnuSelectedUnit.Caption = "Default unit=" + frmSelectUnit.DefaultUnit
                mdiMain.mnuSelectedUnit.Checked = True
                mdiMain.mnuSelectedUnit.Visible = True
                frmSelectUnit.DefaultUnit = ""
            End If
            txtUnit = frmSelectUnit.UnitSelected
            txtUnit_KeyPress 13
            Exit Function
        Else
            SelectedUnit = ""
            response = vbYes
        End If
    Else
        response = MsgBox("Object in undefined unit.  Define now?", vbYesNoCancel)
    End If
End If

Set rsTemp = Nothing

If response = vbCancel Then
    Cancelling = True
    Exit Function
ElseIf response = vbYes Then
    AddUnits.XY(0) = Int(PointsADO.recordset("x"))
    AddUnits.XY(1) = Int(PointsADO.recordset("y"))
    AddUnits.XY(2) = Int(PointsADO.recordset("x")) + 1
    AddUnits.XY(3) = Int(PointsADO.recordset("y")) + 1
    AddUnits.optType(0) = True
    AddUnits.Editing = True
    Screen.MousePointer = 1
    AddUnits.Show 1
    If Not Cancelling Then
        GoTo Start
    End If
End If

End Function

Public Sub FillUnitFields()

UnitTB.Index = "unitname"
UnitTB.Seek "=", txtUnit
If Not UnitTB.NoMatch Then
    For I = 1 To nUnitFields
        Select Case UCase(Unitfield(I))
            Case "UNIT", "ID", "SUFFIX", "X", "Y", "Z", "VANGLE", "HANGLE", "SLOPED"
            Case "PRISM"
                On Error Resume Next
                Gotit = False
                
                Gotit = LCase(SiteDB.TableDefs("EDM_units").Fields("prism").Name) = "prism"
                On Error GoTo 0
                If Gotit Then
                    If IsNull(UnitTB("PRISM")) Then
                        response = MsgBox("Invalid default prism.  Use " + txtprism + "?", vbYesNo)
                        If response = vbNo Then
                            MsgBox ("Please select default prism for " + txtUnit)
                        Else
                            UnitTB.Edit
                            UnitTB("prism") = txtPoleHT
                            UnitTB.Update
                        End If
                    Else
                        If txtPoleHT <> UnitTB("prism") Then
                            Gotit = False
                            For J = 0 To txtprism.ListCount - 1
                                If PoleHeight(txtprism.ItemData(J)) = UnitTB("prism") Then
                                    txtXYZ(2) = Format(Val(txtXYZ(2)) + txtPoleHT - UnitTB("prism"), "######0.000")
                                    txtPoleHT = UnitTB("prism")
                                    txtprism = txtprism.List(J)
                                    UpdatePointsTable "prism", txtPoleHT, 0, 0
                                    UpdatePointsTable "z", txtXYZ(2), 0, 0
                                    Gotit = True
                                    Exit For
                                End If
                            Next J
                            If Not Gotit Then
                                response = MsgBox("Invalid default prism.  Use " + txtprism + "?", vbYesNo)
                                If response = vbNo Then
                                    MsgBox ("Please select default prism for " + txtUnit)
                                Else
                                    UnitTB.Edit
                                    UnitTB("prism") = txtPoleHT
                                    UnitTB.Update
                                End If
                            End If
                        End If
                    End If
                End If
            Case Else
                For J = 1 To Vars
                    If LCase(VarList(J)) = LCase(Unitfield(I)) Then
                        Select Case UCase(VType(J))
                            Case "TEXT"
                                If Not IsNull(UnitTB(Unitfield(I))) Then
                                    TextBox(J) = UnitTB(Unitfield(I))
                                Else
                                    TextBox(J) = ""
                                End If
                                If CountRecords > 0 Then
                                    UpdatePointsTable VarList(J), TextBox(J), 0, 0
                                End If
                            Case "NUMERIC"
                                If Not IsNull(UnitTB(Unitfield(I))) Then
                                    NumberBox(J) = UnitTB(Unitfield(I))
                                Else
                                    NumberBox(J) = ""
                                End If
                                If CountRecords > 0 Then
                                    UpdatePointsTable VarList(J), NumberBox(J), 0, 0
                                End If
                            Case "MENU"
                                If Not IsNull(UnitTB(Unitfield(I))) Then
                                    MenuBox(J) = UnitTB(Unitfield(I))
                                Else
                                    MenuBox(J) = ""
                                End If
                                If CountRecords > 0 Then
                                    UpdatePointsTable VarList(J), MenuBox(J), 0, 0
                                End If
                                MenuBox_Click Int(J)
                        End Select
                        Exit For
                    End If
                Next J
        End Select
    Next I
End If

End Sub

Public Sub ClearDBfields()

mdiMain.StatusBar.Panels(4) = ""
txtPT = ""
Set SiteDB = Nothing
PointsADO.RecordSource = ""
UnitsADO.RecordSource = ""
Set rsTemp = Nothing

Set UnitTB = Nothing
Set PoleTB = Nothing
Set DatumTB = Nothing
Set cfgTB = Nothing
On Error Resume Next
SiteDBname = ""
SiteDBOpen = False
DBPath = ""
DBName = ""
PointTableName = ""
lblDBWarning.Visible = True
lblPointsWarning.Visible = True
lblPoleWarning.Visible = True
txtCurrentRecord = 0
txtTotalRecords = 0
txtprism.Clear
txtXYZ(0).Clear
txtXYZ(1).Clear
txtXYZ(2).Clear
txtUnit.Clear
txtID.Clear
txtSuffix.Clear
txtUnit.Enabled = False
txtID.Enabled = False
txtprism.Enabled = False
txtXYZ(0).Enabled = False
txtXYZ(1).Enabled = False
txtXYZ(2).Enabled = False
frmMain.txtUnit.Enabled = False
frmMain.txtID.Enabled = False
frmMain.txtprism.Enabled = False

End Sub

Public Sub CheckFields()

On Error GoTo 0
'doevents
lblBlankFields.Visible = False

If txtUnit = "" Then
    lblBlankFields = "Unit not entered."
    lblBlankFields.Visible = True
    Exit Sub
End If


For I = 1 To Vars
    Select Case UCase(VarList(I))
        Case "UNIT", "ID", "SUFFIX", "PRISM", "X", "Y", "Z", "VANGLE", "HANGLE", "SLOPED"
        Case Else
            Select Case VType(I)
                Case "TEXT"
                        If TextBox(I) = "" Then
                            lblBlankFields = "Record contains blank fields"
                            lblBlankFields.Visible = True
                            Exit Sub
                        End If
                Case "MENU"
                        If MenuBox(I) = "" Then
                            lblBlankFields = "Record contains blank fields"
                            lblBlankFields.Visible = True
                            Exit Sub
                        End If
                Case "NUMERIC", "INSTRUMENT"
                        If NumberBox(I) = "" Then
                            lblBlankFields = "Record contains blank fields"
                            lblBlankFields.Visible = True
                            Exit Sub
                        End If
            End Select
    End Select
Next I

End Sub

Public Function CheckStatus()

CheckStatus = False
If PointTableName = "" Or frmMain.lblPointsWarning.Visible = True Then
    MsgBox ("Point table must be opened.")
    CheckStatus = True
ElseIf Not StationInitialized Then
    MsgBox ("Total Station not initialized.  Initialize before recording points")
    CheckStatus = True
'ElseIf Not LimitChecking And txtUnit = "" Then
'    MsgBox ("Select Unit before shooting, or set Auto-Find Unit")
'    CheckStatus = True
ElseIf PoleTB.BOF And PoleTB.EOF Then
    MsgBox ("No prisms defined.  Define before taking a shot")
    CheckStatus = True
ElseIf Not frmMain.theoport.PortOpen And EDMName <> "Simulate" And EDMName <> "Microscribe" And EDMName <> "None" Then
    MsgBox ("Total Station not cabled")
    CheckStatus = True
End If

End Function

Public Sub FindBlankField()

On Error GoTo Boxerror
For I = 1 To Vars
    Select Case UCase(VType(I))
        Case "MENU"
            If Trim(MenuBox(I)) = "" Then
                MenuBox(I) = Space(30)
                MenuBox(I).SelStart = 0
                MenuBox(I).SelLength = 30
                MenuBox(I).SetFocus
                Exit For
            End If
        Case "NUMERIC", "INSTRUMENT"
            If Trim(NumberBox(I)) = "" Then
                NumberBox(I) = Space(30)
                NumberBox(I).SelStart = 0
                NumberBox(I).SelLength = 30
                NumberBox(I).SetFocus
                NumberBox(I).SetFocus
                Exit For
            End If
        Case "TEXT"
            If Trim(TextBox(I)) = "" Then
                TextBox(I) = Space(30)
                TextBox(I).SelStart = 0
                TextBox(I).SelLength = 30
                TextBox(I).SetFocus
                TextBox(I).SetFocus
                Exit For
            End If
    End Select
Continue:
Next I

Exit Sub

Boxerror:
Resume Continue

End Sub

Public Sub AdoAccessOLD(Action As Integer, UnitValue As String, IDvalue As String, Var As String, VarValue As String)

'If mdiMain.mnuServer.Checked = True Then
'    Dim rsTemp As ADODB.Recordset
'    Dim ConnectString As String
'    Dim NothingDone As Boolean
'    Dim rsLock As Recordset
'    Dim rsCheck As Recordset
'
'    Cancelling = False
'    Screen.MousePointer = 11
'    time1 = Timer
'    On Error Resume Next
'    Do
'        Set rsCheck = SiteDB.OpenRecordset("edm_units", dbOpenDynaset)
'    Loop Until Err.Number = 0 Or Timer - time1 > 5
'    If Err.Number <> 0 Then
'        MsgBox ("Unable to access server database.  Check connections and retake this shot.")
'        Screen.MousePointer = 1
'        Cancelling = True
'        Exit Sub
'    End If
'    On Error GoTo 0
'    Set rsCheck = Nothing
'    Set rsLock = SiteDB.OpenRecordset("edm_units", dbOpenForwardOnly, dbDenyRead)
'    ConnectString = "provider=microsoft.jet.oledb.4.0;data source=" + SiteDBname
'    Conn.CursorLocation = adUseServer
'    Conn.Open ConnectString
'    JRO.RefreshCache Conn
'    Conn.BeginTrans
'
'    If Action = GetNextID Then
'            SqlString = "select max(id) from [EDM_units] where unit='" + UnitValue + "' and id<'A'"
'            Set rsTemp = Conn.Execute(SqlString)
'            Conn.CommitTrans
'            JRO.RefreshCache Conn
'            If IsNull(rsTemp(0)) Then
'                IDvalue = PadID("1")
'            Else
'                IDvalue = PadID(Str(Val(rsTemp(0)) + 1))
'            End If
'            rsTemp.Close
'            Conn.Close
'            Set rsTemp = Nothing
'            rsLock.Close
'            Set rsLock = Nothing
'            Screen.MousePointer = 1
'            Exit Sub
'    ElseIf Action = DecID Then
'            SqlString = "select max(id) from [EDM_units] where unit='" + UnitValue + "' and id<'A'"
'            Set rsTemp = Conn.Execute(SqlString)
'            Conn.CommitTrans
'            JRO.RefreshCache Conn
'            If IsNull(rsTemp(0)) Or Val(rsTemp(0)) = 0 Then
'                GoTo CloseAll
'            Else
'                If Val(rsTemp(0)) = Val(IDvalue) Then
'                    IDvalue = PadID(Str(Val(rsTemp(0)) - 1))
'                Else
'                    GoTo CloseAll
'                End If
'            End If
'            SqlString = "Update [edm_units] set id=" + IDvalue + " where unit='" + UnitValue + " '"
'            Conn.BeginTrans
'            Conn.Execute SqlString
'            Conn.CommitTrans
'            JRO.RefreshCache Conn
'            rsTemp.Close
'            Conn.Close
'            Set rsTemp = Nothing
'    ElseIf Action = SetID Then
'            SqlString = "select max(id) from [EDM_units] where unit='" + UnitValue + "' and id<'A'"
'            Set rsTemp = Conn.Execute(SqlString)
'            Conn.CommitTrans
'            JRO.RefreshCache Conn
'            If Val(rsTemp(0)) + 1 = Val(IDvalue) Then
'                IDvalue = PadID(IDvalue)
'            ElseIf Val(rsTemp(0)) < Val(IDvalue) Or Val(rsTemp(0)) > Val(IDvalue) + 1 Then
'                response = MsgBox("This ID value is out of sequence.  Reset Last ID in Unit " + UnitName + " to " & Val(ID) & "?", vbYesNo)
'                If response = vbYes Then
'                    IDvalue = PadID(IDvalue)
'                Else
'                    GoTo CloseAll
'                End If
'            Else
'                Screen.MousePointer = 1
'                MsgBox ("Duplicate ID found")
'                GoTo CloseAll
'            End If
'
'            SqlString = "Update [edm_units] set id=" + IDvalue + " where unit='" + UnitValue + " '"
'            Conn.BeginTrans
'            Conn.Execute SqlString
'            Conn.CommitTrans
'            JRO.RefreshCache Conn
'            rsTemp.Close
'            Conn.Close
'            Set rsTemp = Nothing
'    ElseIf Action = DelRec Then
'        SqlString = "select max(suffix) from [" + PointTableName + "] where unit='" + txtUnit + "' and id='" + txtID + "'"
'        Set rsTemp = Conn.Execute(SqlString)
'        Conn.CommitTrans
'        JRO.RefreshCache Conn
'        If rsTemp(0) > 0 Then
'            response = MsgBox("Delete all " + Str(rsTemp(0) + 1) + " records for " & txtUnit & "-" & Trim(txtID) & "?", vbYesNo)
'            If response = vbYes Then
'                Need2Decrement = True
'                SqlString = "delete * from [" + PointTableName + "] where unit='" + txtUnit + "' and id='" + txtID + "'"
'            Else
'                SqlString = "delete * from [" + PointTableName + "] where unit='" + txtUnit + "' and id='" + txtID + "' and suffix=" + txtSuffix
'            End If
'        Else
'            Need2Decrement = True
'            SqlString = "delete * from [" + PointTableName + "] where unit='" + txtUnit + "' and id='" + txtID + "' and suffix=" + txtSuffix
'        End If
'
'        Conn.BeginTrans
'        Conn.Execute SqlString
'        Conn.CommitTrans
'        JRO.RefreshCache Conn
'        rsTemp.Close
'        Conn.Close
'        Set rsTemp = Nothing
'    ElseIf Action = SetField Then
'    End If
'    timer1 = Timer
'    Do
'    Loop Until Timer - timer1 > 4
'
'    rsLock.Close
'    Set rsLock = Nothing
'    Screen.MousePointer = 1
'    Exit Sub
'
'CloseAll:
'        Conn.Close
'        Set rsTemp = Nothing
'        rsLock.Close
'        Screen.MousePointer = 1
'        Exit Sub
'Else
'    Dim RsTemp2 As Recordset
'    If Action = GetNextID Then
'            SqlString = "select max(id) from [EDM_units] where unit='" + UnitValue + "' and id<'A'"
'            Set RsTemp2 = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
'            If IsNull(RsTemp2(0)) Then
'                IDvalue = PadID("1")
'            Else
'                IDvalue = PadID(Str(Val(RsTemp2(0)) + 1))
'            End If
'            RsTemp2.Close
'            Set RsTemp2 = Nothing
'            Screen.MousePointer = 1
'            Exit Sub
'    ElseIf Action = DecID Then
'            SqlString = "select max(id) from [EDM_units] where unit='" + UnitValue + "' and id<'A'"
'            Set RsTemp2 = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
'            If IsNull(RsTemp2(0)) Or Val(RsTemp2(0)) = 0 Then
'            Else
'                If Val(RsTemp2(0)) = Val(IDvalue) Then
'                    IDvalue = PadID(Str(Val(RsTemp2(0)) - 1))
'                End If
'            End If
'            SqlString = "Update [edm_units] set id=" + IDvalue + " where unit='" + UnitValue + " '"
'            SiteDB.Execute SqlString
'            Set RsTemp2 = Nothing
'    ElseIf Action = SetID Then
'            SqlString = "select max(id) from [EDM_units] where unit='" + UnitValue + "' and id<'A'"
'            Set RsTemp2 = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
'            If Val(RsTemp2(0)) + 1 = Val(IDvalue) Then
'                IDvalue = PadID(IDvalue)
'            ElseIf Val(RsTemp2(0)) < Val(IDvalue) Or Val(RsTemp2(0)) > Val(IDvalue) + 1 Then
'                response = MsgBox("This ID value is out of sequence.  Reset Last ID in Unit " + UnitName + " to " & Val(ID) & "?", vbYesNo)
'                If response = vbYes Then
'                    IDvalue = PadID(IDvalue)
'                End If
'            Else
'                Screen.MousePointer = 1
'                MsgBox ("Duplicate ID found")
'            End If
'            SqlString = "Update [edm_units] set id=" + IDvalue + " where unit='" + UnitValue + " '"
'            SiteDB.Execute SqlString
'            RsTemp2.Close
'            Set RsTemp2 = Nothing
'    ElseIf Action = DelRec Then
'        SqlString = "select max(suffix) from [" + PointTableName + "] where unit='" + txtUnit + "' and id='" + txtID + "'"
'        Set RsTemp2 = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
'        If RsTemp2(0) > 0 Then
'            response = MsgBox("Delete all " + Str(RsTemp2(0) + 1) + " records for " & txtUnit & "-" & Trim(txtID) & "?", vbYesNo)
'            If response = vbYes Then
'                Need2Decrement = True
'                SqlString = "delete * from [" + PointTableName + "] where unit='" + txtUnit + "' and id='" + txtID + "'"
'            Else
'                SqlString = "delete * from [" + PointTableName + "] where unit='" + txtUnit + "' and id='" + txtID + "' and suffix=" + txtSuffix
'            End If
'        Else
'            Need2Decrement = True
'            SqlString = "delete * from [" + PointTableName + "] where unit='" + txtUnit + "' and id='" + txtID + "' and suffix=" + txtSuffix
'        End If
'
'        SiteDB.Execute SqlString
'        Set RsTemp2 = Nothing
'    ElseIf Action = SetField Then
'    End If
'    Exit Sub
'End If

End Sub

Public Sub ADOAccess(Action As Integer, UnitValue As String, IDvalue As String, Var As String, VarValue As String)

Dim currentrecord As Variant

UnitsADO.RecordSource = "edm_units"
UnitsADO.Refresh
UnitsADO.recordset.Requery
If Not UnitTB.BOF Or Not UnitTB.EOF Then
    UnitsADO.recordset.MoveLast
End If

Dim updated As Boolean
updated = False

If Action = GetNextID Then
    
    If Not UnitTB.BOF Or Not UnitTB.EOF Then
        UnitTB.MoveFirst
        Do
            If UnitTB("unit") = UnitValue Then
                UnitTB.Edit
                If Val(UnitTB("ID")) = 0 Then
                    UnitTB("id") = PadID(Str(1))
                Else
                    UnitTB("id") = PadID(Str(Val(UnitTB("id")) + 1))
                End If
                UnitTB.Update
                updated = True
                Exit Do
            End If
            UnitTB.MoveNext
            If UnitTB.EOF Then Exit Do
        Loop
    End If
    
    If updated Then
        IDvalue = PadID(UnitTB("id"))
    Else
        MsgBox ("Error updating units table.  ID could not be udpated.  You may need to add a matching Unit.  Otherwise tell Shannon, update ID by hand, and watch for ID sequence errors.")
        Exit Sub
    End If
    
ElseIf Action = DecID Then
    
    If Not UnitTB.BOF Or Not UnitTB.EOF Then
        UnitTB.MoveFirst
        Do
            If UnitTB("unit") = UnitValue Then
                UnitTB.Edit
                If IsNull(UnitTB("ID")) Then
                    UnitTB("id") = PadID(Str(0))
                Else
                    UnitTB("id") = PadID(Str(Val(UnitTB("id")) - 1))
                End If
                UnitTB.Update
                updated = True
                Exit Do
            End If
            UnitTB.MoveNext
            If UnitTB.EOF Then Exit Do
        Loop
    End If
    
    If updated Then
        IDvalue = PadID(UnitTB("id"))
    Else
        MsgBox ("Error updating units table.  ID could not be udpated.  You may need to add a matching Unit.  Otherwise tell Shannon, update ID by hand, and watch for ID sequence errors.")
        Exit Sub
    End If
    
ElseIf Action = SetID Then
        Dim TempIDValue As String
        Do
            UnitsADO.Refresh
            UnitsADO.recordset.Filter = "[unit]='" + UnitValue + "'"
            On Error Resume Next
            TempIDValue = UnitsADO.recordset("ID")
            If Err = 0 Then
                Exit Do
            Else
                MsgBox ("Error updating units table.  ID could not be udpated.  You may need to add a matching Unit.  Otherwise tell Shannon, update ID by hand, and watch for ID sequence errors.")
                Exit Sub
            End If
        Loop
        On Error GoTo 0
        UnitsADO.recordset.Filter = adFilterNone
        If Val(TempIDValue) + 1 = Val(IDvalue) Then
            IDvalue = PadID(IDvalue)
        ElseIf Val(TempIDValue) < Val(IDvalue) Or Val(TempIDValue) > Val(IDvalue) + 1 Then
            response = MsgBox("This ID value is out of sequence.  Reset Last ID in Unit " + OriginalUnit + " to " & Val(IDvalue) & "?", vbYesNo)
            If response = vbYes Then
                IDvalue = PadID(IDvalue)
            Else
                Exit Sub
            End If
        Else
'            Screen.MousePointer = 1
'            MsgBox ("Duplicate ID found")
'            Exit Sub
        End If
        Do
            UnitsADO.Refresh
            UnitsADO.recordset.Filter = "[unit]='" + UnitValue + "'"
            On Error Resume Next
            If IsNull(UnitsADO.recordset("ID")) Then
                UnitsADO.recordset.Update "id", 0
            Else
                UnitsADO.recordset.Update "id", IDvalue
            End If
            If Err = 0 Then Exit Do
        Loop
        On Error GoTo 0
        UnitsADO.recordset.Filter = adFilterNone

ElseIf Action = DelRec Then
        'PointsADO.Recordset.Filter = "unit='" + UnitValue + "' and id='" + IDvalue + "'"
        'ndels = PointsADO.Recordset.RecordCount
        
        GridLoading = True
        currentrecord = PointsADO.recordset.Bookmark
        
        'first peel through to get number of records to delete
        ndels = 0
        PointsADO.recordset.MoveFirst
        Do Until PointsADO.recordset.EOF
            If PointsADO.recordset("unit") = UnitValue And PointsADO.recordset("id") = IDvalue Then ndels = ndels + 1
            PointsADO.recordset.MoveNext
        Loop
                
        If ndels > 1 Then
            response = MsgBox("Delete all" + Str(ndels) + " records for " & txtUnit & "-" & Trim(txtID) & "?", vbYesNoCancel)
            If response = vbCancel Then
                Exit Sub
            ElseIf response = vbYes Then
                Need2Decrement = True
                PointsADO.recordset.MoveFirst
                Do Until PointsADO.recordset.EOF
                    If PointsADO.recordset("unit") = UnitValue And PointsADO.recordset("id") = IDvalue Then PointsADO.recordset.Delete
                    PointsADO.recordset.MoveNext
                Loop
            Else
                PointsADO.recordset.Bookmark = currentrecord
                PointsADO.recordset.Delete
            End If
        ElseIf ndels > 0 Then
            Need2Decrement = True
            PointsADO.recordset.Bookmark = currentrecord
            PointsADO.recordset.Delete
        Else
            MsgBox ("Record not found")
        End If
        
        If Need2Decrement Then
            UnitsADO.recordset.Requery
            UnitsADO.recordset.Filter = "unit='" + UnitValue + "'"
            If Val(UnitsADO.recordset("id")) > Val(IDvalue) Then
                Need2Decrement = False
            End If
            UnitsADO.recordset.Filter = adFilterNone
        End If
        
        'PointsADO.Recordset.Filter = adFilterNone
        'PointsADO.Recordset.Requery
        'PointsADO.Refresh
        
End If

End Sub

Public Sub FillDependentFields()

DefaultsTB.Seek "=", MasterVal
If Not DefaultsTB.NoMatch Then
    For I = 1 To nDependentVars
        For J = 1 To Vars
            If LCase(VarList(J)) = LCase(DependentVar(I)) Then
                Select Case UCase(VType(J))
                    Case "TEXT"
                        If Not IsNull(DefaultsTB(DependentVar(I))) Then
                            TextBox(J) = DefaultsTB(DependentVar(I))
                        Else
                            TextBox(J) = ""
                        End If
                        If CountRecords > 0 Then
                            UpdatePointsTable VarList(J), TextBox(J), 0, 0
                        End If
                    Case "NUMERIC"
                        If Not IsNull(DefaultsTB(DependentVar(I))) Then
                            NumberBox(J) = DefaultsTB(DependentVar(I))
                        Else
                            NumberBox(J) = ""
                        End If
                        If CountRecords > 0 Then
                            UpdatePointsTable VarList(J), NumberBox(J), 0, 0
                        End If
                    Case "MENU"
                        If Not IsNull(DefaultsTB(DependentVar(I))) Then
                            MenuBox(J) = DefaultsTB(DependentVar(I))
                        Else
                            MenuBox(J) = ""
                        End If
                        If CountRecords > 0 Then
                            UpdatePointsTable VarList(J), MenuBox(J), 0, 0
                        End If
                    End Select
               Exit For
           End If
        Next J
    Next I
Else
    If Trim(MasterVal) <> "" Then
        DefaultsTB.AddNew
        DefaultsTB(MasterVar) = MasterVal
        DefaultsTB.Update
        DefaultsTB.Seek "=", MasterVal
        For I = 1 To Vars
            For J = 1 To nDependentVars
                If LCase(VarList(I)) = LCase(DependentVar(J)) Then
                    DefaultsTB.Edit
                    Select Case UCase(VType(I))
                    Case "TEXT"
                        If Trim(TextBox(I)) = "" Then
                            DefaultsTB(VarList(I)) = " "
                        Else
                            DefaultsTB(VarList(I)) = TextBox(I)
                        End If
                    Case "NUMERIC"
                        If Trim(NumberBox(I)) = "" Then
                            DefaultsTB(VarList(I)) = " "
                        Else
                            DefaultsTB(VarList(I)) = NumberBox(I)
                        End If
                    Case "MENU"
                        If Trim(MenuBox(I)) = "" Then
                            DefaultsTB(VarList(I)) = " "
                        Else
                            DefaultsTB(VarList(I)) = MenuBox(I)
                        End If
                    End Select
                    DefaultsTB.Update
                End If
            Next J
        Next I
    End If
End If
    
End Sub

Public Sub UpdateDependentVar(CurrentVar As String, CurrentVal As Variant)

Cancelling = False
MasterVal = ""
For I = 1 To Vars
    If UCase(VarList(I)) = UCase(MasterVar) Then
        Select Case UCase(VType(I))
            Case "TEXT"
                If Not IsNull(TextBox(I)) Then
                    MasterVal = TextBox(I)
                Else
                    MasterVal = ""
                End If
            Case "NUMERIC"
                If Not IsNull(NumberBox(I)) Then
                    MasterVal = NumberBox(I)
                Else
                    MasterVal = ""
                End If
            Case "MENU"
                If Not IsNull(MenuBox(I)) Then
                    MasterVal = MenuBox(I)
                Else
                    MasterVal = ""
                End If
        End Select
    Exit For
    End If
    
Next I

If MasterVal = "" Then
    MsgBox ("Because " + CurrentVar + " is dependent on the value of " + MasterVar + ", you must first enter value for " + MasterVar)
    Cancelling = True
    Exit Sub
End If

DefaultsTB.Seek "=", MasterVal

If Not DefaultsTB.NoMatch Then
    DefaultsTB.Edit
    If Trim(CurrentVal) = "" Then CurrentVal = " "
    DefaultsTB(CurrentVar) = CurrentVal
    DefaultsTB.Update
End If
    
End Sub

