VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EDM Windows"
   ClientHeight    =   6264
   ClientLeft      =   156
   ClientTop       =   432
   ClientWidth     =   10224
   Icon            =   "frmMain2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6264
   ScaleWidth      =   10224
   Begin VB.TextBox txtPoleHT 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   2550
      Width           =   1155
   End
   Begin VB.ComboBox txtXYZ 
      Height          =   315
      Index           =   2
      Left            =   3000
      TabIndex        =   57
      Text            =   "txtXYZ"
      Top             =   2520
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&X-Shot"
      Height          =   375
      Left            =   8880
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   2700
      Width           =   1185
   End
   Begin VB.CommandButton Button 
      Caption         =   "Button6"
      Height          =   345
      Index           =   6
      Left            =   6870
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   3030
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Button 
      Caption         =   "Button5"
      Height          =   345
      Index           =   5
      Left            =   6870
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2658
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Button 
      Caption         =   "Button4"
      Height          =   345
      Index           =   4
      Left            =   6870
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2286
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Button 
      Caption         =   "Button3"
      Height          =   345
      Index           =   3
      Left            =   6870
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1914
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Button 
      Caption         =   "Button2"
      Height          =   345
      Index           =   2
      Left            =   6870
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1542
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Button 
      Caption         =   "Button1"
      Height          =   345
      Index           =   1
      Left            =   6870
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TxtDB 
      Height          =   315
      Left            =   1170
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   720
      Width           =   3885
   End
   Begin VB.TextBox txtCFG 
      Height          =   315
      Left            =   1170
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   360
      Width           =   3885
   End
   Begin VB.ComboBox txtPT 
      Height          =   315
      Left            =   6390
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   330
      Width           =   2145
   End
   Begin VB.ComboBox txtXYZ 
      Height          =   315
      Index           =   1
      Left            =   3000
      TabIndex        =   5
      Text            =   "txtXYZ"
      Top             =   2160
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   1308
      Left            =   8976
      Picture         =   "frmMain2.frx":000C
      ScaleHeight     =   1284
      ScaleWidth      =   852
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   264
      Width           =   876
   End
   Begin VB.CommandButton cmdPlusShot 
      Caption         =   "&+-Shot"
      Height          =   375
      Left            =   8880
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2250
      Width           =   1185
   End
   Begin VB.TextBox txtTotalRecords 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8010
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   705
      Width           =   495
   End
   Begin VB.TextBox txtCurrentRecord 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   705
      Width           =   495
   End
   Begin VB.ComboBox txtXYZ 
      Height          =   288
      Index           =   0
      Left            =   3000
      TabIndex        =   4
      Text            =   "txtXYZ"
      Top             =   1800
      Width           =   1155
   End
   Begin VB.CommandButton cmdShoot 
      Caption         =   "&New Object"
      Height          =   375
      Left            =   8880
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1185
   End
   Begin VB.ComboBox txtPrism 
      Height          =   288
      Left            =   930
      Sorted          =   -1  'True
      TabIndex        =   3
      Text            =   "txtPrism"
      Top             =   2160
      Width           =   1155
   End
   Begin VB.TextBox txtSlopeD 
      Height          =   285
      Left            =   5490
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "txtSlopeD"
      Top             =   2490
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox txtVangle 
      Height          =   285
      Left            =   5490
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "txtVangle"
      Top             =   2145
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox txtHangle 
      Height          =   285
      Left            =   5490
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "txtHangle"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.ComboBox txtID 
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Text            =   "txtID"
      Top             =   1260
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.ComboBox txtUnit 
      Height          =   315
      Left            =   2010
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "txtUnit"
      Top             =   1260
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtSuffix 
      Height          =   285
      Left            =   6000
      TabIndex        =   2
      Text            =   "txtSuffix"
      Top             =   1260
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.ComboBox MenuBox 
      Height          =   315
      Index           =   0
      Left            =   4530
      TabIndex        =   8
      Text            =   "MenuBox"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TextBox 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3510
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   3510
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8940
      Top             =   7410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm theoport 
      Left            =   8310
      Top             =   7500
      _ExtentX        =   995
      _ExtentY        =   995
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label lblAutoFind 
      AutoSize        =   -1  'True
      Caption         =   "Auto-Find Units set to ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   56
      Top             =   90
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Label lblBlankFields 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Record contains blank fields "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6180
      TabIndex        =   55
      Top             =   90
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "To edit fields, enter a new value, or choose one from the drop-down list,  and then press ENTER."
      Height          =   195
      Left            =   3240
      TabIndex        =   54
      Top             =   5910
      Width           =   6840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "To move between fields, use the TAB key."
      Height          =   195
      Left            =   90
      TabIndex        =   53
      Top             =   5910
      Width           =   3015
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   $"frmMain2.frx":5A96
      Height          =   195
      Left            =   90
      TabIndex        =   52
      Top             =   5640
      Width           =   9750
   End
   Begin VB.Shape Shape4 
      Height          =   825
      Left            =   240
      Top             =   300
      Width           =   8445
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Points Table:"
      Height          =   195
      Left            =   5220
      TabIndex        =   40
      Top             =   435
      Width           =   930
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Database:"
      Height          =   195
      Left            =   390
      TabIndex        =   39
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CFG:"
      Height          =   195
      Left            =   735
      TabIndex        =   38
      Top             =   435
      Width           =   360
   End
   Begin VB.Label Label5 
      Caption         =   "Total Records:"
      Height          =   255
      Left            =   6930
      TabIndex        =   32
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblCurrentRecord 
      AutoSize        =   -1  'True
      Caption         =   "Current Record:"
      Height          =   195
      Left            =   5220
      TabIndex        =   31
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label lblPoleHT 
      Alignment       =   1  'Right Justify
      Caption         =   "Height"
      Height          =   192
      Left            =   324
      TabIndex        =   30
      Top             =   2580
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "EDM Fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   300
      TabIndex        =   25
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      Height          =   465
      Left            =   240
      Top             =   1170
      Width           =   6525
   End
   Begin VB.Shape Shape1 
      Height          =   1272
      Left            =   240
      Top             =   1680
      Width           =   6528
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      Caption         =   "X:"
      Height          =   192
      Left            =   2340
      TabIndex        =   24
      Top             =   1860
      Visible         =   0   'False
      Width           =   564
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      Caption         =   "Y:"
      Height          =   192
      Left            =   2340
      TabIndex        =   23
      Top             =   2220
      Visible         =   0   'False
      Width           =   564
   End
   Begin VB.Label lblZ 
      Alignment       =   1  'Right Justify
      Caption         =   "Z:"
      Height          =   192
      Left            =   2364
      TabIndex        =   22
      Top             =   2580
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblHangle 
      Alignment       =   1  'Right Justify
      Caption         =   "Horizontal Angle:"
      Height          =   195
      Left            =   4215
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblVangle 
      Alignment       =   1  'Right Justify
      Caption         =   "Vertical Angle:"
      Height          =   195
      Left            =   4395
      TabIndex        =   20
      Top             =   2280
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label lblSlopeD 
      Alignment       =   1  'Right Justify
      Caption         =   "Slope Distance:"
      Height          =   195
      Left            =   4290
      TabIndex        =   19
      Top             =   2580
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblPrism 
      Alignment       =   1  'Right Justify
      Caption         =   "Prism: "
      Height          =   192
      Left            =   252
      TabIndex        =   18
      Top             =   2220
      Width           =   636
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Object ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   300
      TabIndex        =   14
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Optional Fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   300
      TabIndex        =   13
      Top             =   3210
      Width           =   1590
   End
   Begin VB.Label lblSuffix 
      Alignment       =   1  'Right Justify
      Caption         =   "Suffix:"
      Height          =   192
      Left            =   5280
      TabIndex        =   12
      Top             =   1296
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblID 
      Alignment       =   1  'Right Justify
      Caption         =   "ID:"
      Height          =   192
      Left            =   3360
      TabIndex        =   11
      Top             =   1296
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblUnit 
      Alignment       =   1  'Right Justify
      Caption         =   "Unit:"
      Height          =   192
      Left            =   1344
      TabIndex        =   10
      Top             =   1296
      Visible         =   0   'False
      Width           =   672
   End
   Begin VB.Label VarLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Label7"
      Height          =   195
      Index           =   0
      Left            =   285
      TabIndex        =   9
      Top             =   3525
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblCFGWarning 
      Alignment       =   2  'Center
      Caption         =   "No CFG File Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   3168
      TabIndex        =   28
      Top             =   24
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Label lblDBWarning 
      Alignment       =   2  'Center
      Caption         =   "No Site Database Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   3168
      TabIndex        =   26
      Top             =   24
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Label lblPoleWarning 
      Alignment       =   2  'Center
      Caption         =   "No Prisms Defined"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   3168
      TabIndex        =   37
      Top             =   24
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Label lblPointsWarning 
      Alignment       =   2  'Center
      Caption         =   "No Points Table Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   3168
      TabIndex        =   27
      Top             =   24
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Label lblEDMWarning 
      Alignment       =   2  'Center
      Caption         =   "No EDM defined"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   3168
      TabIndex        =   44
      Top             =   24
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Label lblStationWarning 
      Alignment       =   2  'Center
      Caption         =   "Station not Initialized"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   3168
      TabIndex        =   59
      Top             =   24
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewCFG 
         Caption         =   "New CFG File"
      End
      Begin VB.Menu mnuOpenCFG 
         Caption         =   "Open CFG File"
      End
      Begin VB.Menu mnuSaveCFGas 
         Caption         =   "Save CFG as ...."
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrinter 
         Caption         =   "Setup Printer"
      End
      Begin VB.Menu space6 
         Caption         =   "-"
      End
      Begin VB.Menu FileList 
         Caption         =   "filelist"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu FileList 
         Caption         =   "filelist"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu FileList 
         Caption         =   "FileList"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu FileList 
         Caption         =   "FileList"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu FileList 
         Caption         =   "FileList"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu space5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditFields 
         Caption         =   "Fields"
      End
      Begin VB.Menu mnuEditPrisms 
         Caption         =   "Prisms"
      End
      Begin VB.Menu mnuEditUnits 
         Caption         =   "Units"
      End
      Begin VB.Menu mnuCreateDatum 
         Caption         =   "Datums"
      End
      Begin VB.Menu mnuButtons 
         Caption         =   "Shot Buttons"
      End
      Begin VB.Menu mnueditspace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteLast 
         Caption         =   "Delete Current Shot"
      End
      Begin VB.Menu DeleteAll 
         Caption         =   "Delete All Points"
      End
   End
   Begin VB.Menu mnuStation 
      Caption         =   "&Station"
      Begin VB.Menu mnuTheodolite 
         Caption         =   "Select Total Station"
      End
      Begin VB.Menu mnuInitialize 
         Caption         =   "Initialize"
      End
      Begin VB.Menu mnuStationStatus 
         Caption         =   "Status"
      End
      Begin VB.Menu mnuStationVerify 
         Caption         =   "Verify"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Plot"
      Begin VB.Menu mnuViewPoints 
         Caption         =   "Points"
      End
      Begin VB.Menu mnuViewDatums 
         Caption         =   "Datums"
      End
      Begin VB.Menu mnuViewUnits 
         Caption         =   "Units"
      End
      Begin VB.Menu mnuViewAll 
         Caption         =   "All"
      End
      Begin VB.Menu plotspace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewHide 
         Caption         =   "Hide Plot"
      End
   End
   Begin VB.Menu mnuDB 
      Caption         =   "&Database"
      Begin VB.Menu mnuNewDB 
         Caption         =   "New Site Database"
      End
      Begin VB.Menu mnuOpenDB 
         Caption         =   "Open Site Database"
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewPointsTB 
         Caption         =   "Select/Create Points Table"
      End
      Begin VB.Menu space15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImportTables 
         Caption         =   "Import Tables from External Database"
      End
      Begin VB.Menu mnuImportCFGfield 
         Caption         =   "Import Fields from CFG file"
      End
      Begin VB.Menu dbspace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvert2Newplot 
         Caption         =   "Convert Database to Newplot Format"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuFindUnit 
         Caption         =   "Auto-Find Unit"
      End
      Begin VB.Menu mnuPrismPrompt 
         Caption         =   "Prompt for Prism"
      End
      Begin VB.Menu mnuPrintShots 
         Caption         =   "Print Shots"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu helpstatus 
         Caption         =   "&Status"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChangingXYZ As Boolean
Dim OrigIndex As Integer
Dim OrigValue As String
Dim Dropping As Boolean
Public OriginalXYZ As Single
Const IDonly = 1
Const Everything = 2
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
txtSlopeD.Visible = False

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
        Case "ID"
            lblID.Visible = True
            lblID = VPrompt(I)
            txtID.Visible = True
        Case "SUFFIX"
            lblSuffix.Visible = True
            lblSuffix = VPrompt(I)
            txtSuffix.Visible = True
        Case "PRISM"
            lblPrism.Visible = True
            lblPrism = VPrompt(I)
            txtPrism.Visible = True
            lblPoleHT.Visible = True
            txtPoleHT.Visible = True
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
            txtSlopeD.Visible = True
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
                    If VarList(I) = "SITENAME" Then TextBox(I) = SiteName
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
                    
            End Select
            Noptionals = Noptionals + 1
            LastOptional = I
            LabelLeft = LabelLeft + 2 * VarLabel(I).Width + 100
            BoxLeft = LabelLeft + VarLabel(I).Width + 50
            If Noptionals Mod 3 = 0 Then
                LabelTop = LabelTop + VarLabel(I).Height + 120
                'LabelTop = VarLabel(0).Top
                LabelLeft = OriginalLeft
                'BoxTop = TextBox(0).Top
                BoxTop = LabelTop
                BoxLeft = LabelLeft + VarLabel(I).Width + 50
            End If
    End Select
Next I
LabelTop = LabelTop + VarLabel(LastOptional).Height + 200
Label8.Top = LabelTop
Label9.Top = Label8.Top + Label8.Height + 10
Label10.Top = Label9.Top
Me.Height = Label10.Top + Label10.Height + 900

Loading = False
End Sub




Private Sub Button_Click(Index As Integer)
    
If CheckStatus() = True Then Exit Sub

IncrementID
txtSuffix = 0
OriginalUnit = txtUnit
OriginalID = txtID
OriginalSuffix = txtSuffix
TestShot
If LimitChecking And txtSuffix = 0 Then
    FindUnit
    If Cancelling Then
        PointsTB.Delete
        PointsTB.MoveLast
    End If
Else
    FillUnitFields
End If
CheckFields
UpdateUnitTable txtUnit, txtID, Everything
If mnuViewPoints.Checked Then
    frmPlot.SetScale
    frmPlot.PlotPoints
End If

For I = 1 To nButtonVars(Index)
    If LCase(VarList(ButtonVars(Index, I, 1))) = "unit" Then
        txtUnit = ButtonVars(Index, I, 2)
        txtUnit_KeyPress 13
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
            PointsTB.Edit
            PointsTB("id") = txtID
            PointsTB("suffix") = 0
            PointsTB.Update
            OriginalID = txtID
            
        End If
    ElseIf LCase(VarList(ButtonVars(Index, I, 1))) = "prism" Then
        Gotit = False
        For J = 0 To txtPrism.ListCount - 1
            If LCase(txtPrism.List(J)) = LCase(ButtonVars(Index, I, 2)) Then
                txtPrism.ListIndex = J
                Gotit = True
                Exit For
            End If
        Next J
        If Not Gotit Then
            MsgBox ("Prism name not found in current prism list")
        Else
            txtPoleHT = PoleHeight(txtPrism.ItemData(txtPrism.ListIndex))
            txtXYZ(2) = Format(Val(txtXYZ(2)) + OriginalPoleHT - PoleHeight(txtPrism.ItemData(txtPrism.ListIndex)), "#######0.000")
            PointsTB.Edit
            PointsTB("prism") = txtPoleHT
            PointsTB("z") = txtXYZ(2)
            PointsTB.Update
        End If
    Else
        Select Case VType(ButtonVars(Index, I, 1))
            Case "TEXT"
                TextBox(ButtonVars(Index, I, 1)).Text = ButtonVars(Index, I, 2)
            Case "MENU"
                MenuBox(ButtonVars(Index, I, 1)) = ButtonVars(Index, I, 2)
            Case "NUMERIC", "INSTRUMENT"
                NumberBox(ButtonVars(Index, I, 1)) = ButtonVars(Index, I, 2)
        End Select
        PointsTB.Edit
        PointsTB(VarList(ButtonVars(Index, I, 1))) = ButtonVars(Index, I, 2)
        PointsTB.Update
    End If
Next I

Picture1.SetFocus

End Sub

Public Sub cmdPlusShot_Click()

If CheckStatus() = True Then Exit Sub
If txtUnit = "" Or txtID = "" Then
    MsgBox ("You cannot continue with an object unless it has a valid Unit and ID.")
    Exit Sub
End If

txtSuffix = Val(txtSuffix) + 1

Take_Shot

If mnuViewPoints.Checked Then
    frmPlot.SetScale
    frmPlot.PlotPoints
End If

Picture1.SetFocus

End Sub

Private Sub cmdShoot_Click()

If CheckStatus() = True Then Exit Sub

IncrementID
txtSuffix = 0
OriginalUnit = txtUnit
OriginalID = txtID
OriginalSuffix = txtSuffix

Take_Shot

If LimitChecking And txtSuffix = 0 Then
    FindUnit
    If Cancelling Then
        PointsTB.Delete
        PointsTB.MoveLast
    End If
Else
    FillUnitFields
End If
CheckFields

UpdateUnitTable txtUnit, txtID, Everything
If mnuViewPoints.Checked Then
    frmPlot.SetScale
    frmPlot.PlotPoints
End If
Picture1.SetFocus

End Sub

Private Sub Command1_Click()

Dim edmpoffset As Single
Dim edmshot As shotdata

If Not StationInitialized Then
    MsgBox ("Total Station not initialized.  Initialize before recording points")
    Exit Sub
ElseIf PoleTB.BOF And PoleTB.EOF Then
    MsgBox ("No prisms defined.  Define before taking a shot")
    Exit Sub
End If

Select Case UCase(EDMName)
Case "SIMULATE"
    Randomize
    X = Int((3) * Rnd) + 998 + Rnd
    y = Int((1) * Rnd) + 1012 + Rnd
    z = Rnd
    edmshot.X = X
    edmshot.y = y
    edmshot.z = z
    edmshot.hangle = 111.505
    edmshot.vangle = 90.202
    edmshot.sloped = Sqr(edmshot.X ^ 2 + edmshot.y ^ 2 + edmshot.z ^ 2)

Case Else
    Call recordpoint(returndata$)
    Call parsenez(returndata$, edmshot, edmpoffset, mesunits$, angleunit$, errorcode)
    If errorcode = 0 Then
        Call vhdtonez(edmshot)
    End If

End Select

If txtPrism.ListIndex <> -1 Then
    edmshot.poleh = PoleHeight(txtPrism.ItemData(txtPrism.ListIndex))
Else
    edmshot.poleh = 0
End If

edmshot.X = CurrentStation.X + edmshot.X
edmshot.y = CurrentStation.y + edmshot.y
edmshot.z = CurrentStation.z + edmshot.z - edmshot.poleh

actuald = Sqr((CurrentStation.X - edmshot.X) ^ 2 + (CurrentStation.y - edmshot.y) ^ 2)

TempString = "X : " & Format(edmshot.X, "######0.000") & Chr(13)
TempString = TempString & "Y : " & Format(edmshot.y, "######0.000") & Chr(13)
TempString = TempString & "Z : " & Format(edmshot.z, "######0.000") & Chr(13)
TempString = TempString & "Horizontal angle : " & Format(edmshot.vangle, "######0.0000") & Chr(13)
TempString = TempString & "Vertical angle : " & Format(edmshot.hangle, "######0.0000") & Chr(13)
TempString = TempString & "Horizontal Distance : " & Format(actuald, "######0.000") & Chr(13)
TempString = TempString & "Slope Distance : " & Format(edmshot.sloped, "######0.000")

MsgBox (TempString)

End Sub

Private Sub DeleteAll_Click()

If PointTableName = "" Then
    MsgBox ("Open points table before performing this operation")
    Exit Sub
End If
If PointsTB.RecordCount = 0 Then
    Exit Sub
End If
response = MsgBox("This will permanently delete all records.  Continue anyway?", vbYesNo)
If response = vbYes Then
    Set PointsTB = Nothing
    SqlString = "delete * from " + PointTableName
    SiteDB.Execute SqlString
    Set PointsTB = SiteDB.OpenRecordset(PointTableName, dbOpenDynaset)
    txtCurrentRecord = 0
    txtTotalRecords = 0
    If mnuViewPoints.Checked Then
        frmPlot.SetScale
        frmPlot.PlotPoints
    End If
End If


    
End Sub

Private Sub FileList_Click(Index As Integer)
CFGName = Filelist(Index).Caption
Cancelling = False
parsecfg A


'put code in here to rewrite cfg if necessary
If Cancelling Then
    Cancelling = False
    FormatVarList
    Exit Sub
End If
LastPath = GetPath(CFGName)

End Sub

Private Sub Form_Load()
'Public variables initialized
BannerHeight = 400
BannerWidth = 150
Me.Show

Loading = True
inifile$ = fixpath(App.Path) + "edm.ini"
txtXYZ(0).Clear
txtXYZ(1).Clear
txtXYZ(2).Clear
txtPrism.Clear
txtXYZ(0).AddItem "Offset East"
txtXYZ(0).AddItem "Offset West"
txtXYZ(1).AddItem "Offset North"
txtXYZ(1).AddItem "Offset South"
txtXYZ(2).AddItem "Offset Up"
txtXYZ(2).AddItem "Offset Down"

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
If PointTableName = "" Then
    lblPointsWarning.Visible = True
    Gotit = False
End If
If Not Gotit Then Exit Sub

If EDMName$ <> "" And comport <> "" And comsettings <> "" Then
    Select Case UCase(EDMName$)
    Case "TOPCON"
        answer = MsgBox("Cable the total station to the computer and communications will be initialized.", vbOKCancel)
        If answer = 1 Then
            Call initcomport(comport, errorcode)
        End If
    Case "WILD"
        Call initcomport(comport, errorcode)
    Case "SOKKIA"
        Call initcomport(comport, errorcode)
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

End Sub


Private Sub Form_Unload(Cancel As Integer)

inifile$ = fixpath(App.Path) + "edm.ini"
Call WriteEDMIni(inifile$)
Dim Inidata(100, 2) As String
Dim IniClass As String
Dim Status As Byte

IniClass = "[EDM]"
Inidata(1, 1) = "Sitename"
Inidata(1, 2) = SiteName
Inidata(2, 1) = "Database"
Inidata(2, 2) = SiteDBname
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
For I = 2 To nUnitFields
    Inidata(8, 2) = Inidata(8, 2) + "," + Unitfield(I)
Next I
Inidata(9, 1) = "Limitchecking"
If LimitChecking Then
    Inidata(9, 2) = "Yes"
Else
    Inidata(9, 2) = "No"
End If


Call WriteIni(CFGName, IniClass, Inidata(), Status)



End Sub


Private Sub Label11_Click()

End Sub


Private Sub HelpStatus_Click()

frmStatus.Show 1

End Sub

Private Sub MenuBox_Click(Index As Integer)
UpdatePointsTable VarList(Index), MenuBox(Index), 1, 1
OrigValue = MenuBox(Index)
CheckFields
End Sub

Private Sub MenuBox_DropDown(Index As Integer)

If PointTableName = "" Then
    MsgBox ("Open point table first")
    MenuBox(Index) = OriginalValue
    Exit Sub
End If
MenuString = VMenu(Index)

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

End Sub

Private Sub MenuBox_GotFocus(Index As Integer)
OrigValue = MenuBox(Index)
If MenuBox(Index) = "" Then
    MenuBox(Index) = Space(30)
End If
MenuBox(Index).SelStart = 0
MenuBox(Index).SelLength = Len(MenuBox(Index))
End Sub

Private Sub MenuBox_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
    Picture1.SetFocus
    
ElseIf KeyAscii = 13 And Trim(MenuBox(Index)) <> "" Then
        If MenuBox(Index) = OrigValue Then
            Picture1.SetFocus
        Else
            MenuBox(Index) = UCase(MenuBox(Index))
            Gotit = False
            For I = 0 To MenuBox(Index).ListCount - 1
                If MenuBox(Index) = MenuBox(Index).List(I) Then
                    Gotit = True
                    Exit For
                End If
            Next I
            If Not Gotit Then
                response = MsgBox("Add " + MenuBox(Index) + " to list of terms for " + VarList(Index) + "?", vbYesNo)
                If response = vbNo Then Exit Sub
                MenuBox(Index).AddItem MenuBox(Index)
                If Len(VMenu(Index)) > 0 Then
                    VMenu(Index) = VMenu(Index) + "," + MenuBox(Index)
                Else
                    VMenu(Index) = MenuBox(Index)
                End If
                Dim Inidata(1, 2) As String
                Dim IniClass As String
                Dim Status As Byte
                IniClass = VarList(Index)
                Inidata(1, 1) = "Menu"
                Inidata(1, 2) = VMenu(Index)
                Call WriteIni(CFGName, IniClass, Inidata(), Status)
            End If
            UpdatePointsTable VarList(Index), MenuBox(Index), 1, 1
            OrigValue = MenuBox(Index)
            CheckFields
        End If
End If
End Sub

Private Sub MenuBox_LostFocus(Index As Integer)
MenuBox(Index) = OrigValue
End Sub


Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuButtons_Click()

If CFGName = "" Then
    MsgBox ("Open or Create CFG before performing this operation")
    Exit Sub
End If
frmButtons.Show 1

End Sub

Private Sub mnuConvert2Newplot_Click()

If CFGName = "" Then
    MsgBox ("Must open valid CFG file before performing this operation")
    Exit Sub
End If
If SiteDBname = "" Then
    MsgBox ("Must open database before performing this operation")
    Exit Sub
End If
Screen.MousePointer = 11
If Not tablematch("context") Then
    CreateContext
End If
If Not tablematch("xyz") Then
    CreateXYZ
End If
MsgBox ("Done")
Screen.MousePointer = 1
mnuConvert2Newplot.Enabled = False
End Sub

Private Sub mnuCreateDatum_Click()
If SiteDBname = "" Then
    MsgBox ("You must open site database before defining datums")
    Exit Sub
End If
frmDatumSheet.Show 1

End Sub

Private Sub mnuDeleteLast_Click()

If PointTableName = "" Then
    MsgBox ("Open points table before performing this operation")
    Exit Sub
End If
If PointsTB.RecordCount = 0 Then
    Exit Sub
End If

DeleteRecord
If mnuViewPoints.Checked Then
    frmPlot.SetScale
    frmPlot.PlotPoints
End If

End Sub

Private Sub mnuEditFields_Click()
If CFGName = "" Then
    MsgBox ("Open or Create CFG before performing this operation")
    Exit Sub
End If
frmEditField.Show 1
End Sub


Private Sub mnuEditPrisms_Click()

If PoleTableName$ <> "" Then
    frmPolesheet.Show 1
    txtPrism.Clear
    If PoleTB.RecordCount = 0 Then
        lblPoleWarning.Visible = True
    Else
        lblPoleWarning.Visible = False
        PoleTB.MoveFirst
        If Not PoleTB.EOF Then
            PoleTB.MoveFirst
            While Not PoleTB.EOF
                nPoleHeights = nPoleHeights + 1
                txtPrism.AddItem PoleTB("pole")
                txtPrism.ItemData(frmMain.txtPrism.NewIndex) = nPoleHeights
                PoleHeight(nPoleHeights) = PoleTB("height")
                PoleTB.MoveNext
            Wend
            txtPrism.ListIndex = 0
        End If
    End If
    
Else
    MsgBox "Open or create a database first.", vbInformation
End If
End Sub

Private Sub mnuEditUnits_Click()

If CFGName = "" Then
    MsgBox ("Open CFG and Database before performing this operation")
    Exit Sub
End If
If SiteDBname$ <> "" Then
    AddUnits.Show 1
    
Else
    MsgBox "Open or create a site first.", vbInformation
End If
End Sub

Private Sub mnuExit_Click()

Dim Cancel As Integer

Call Form_Unload(Cancel)

End

End Sub

Private Sub mnuFindUnit_Click()
If mnuFindUnit.Checked Then
    mnuFindUnit.Checked = False
    LimitChecking = False
    lblAutoFind.Visible = False
Else
    mnuFindUnit.Checked = True
    LimitChecking = True
    lblAutoFind.Visible = True
End If

Dim Inidata(1, 2) As String
Dim IniClass As String
Dim Status As Byte
IniClass = "[EDM]"
Inidata(1, 1) = "Limitchecking"
If LimitChecking Then
    Inidata(1, 2) = "YES"
Else
    Inidata(1, 2) = "NO"
End If

Call WriteIni(CFGName, IniClass, Inidata(), Status)

End Sub

Private Sub mnuImportCFGfield_Click()
frmImportFields.Show

End Sub

Private Sub mnuImportTables_Click()

If SiteDBname = "" Then
    MsgBox ("Open database before performing this operation")
    Exit Sub
End If

On Error Resume Next
frmInputTables.Show
On Error GoTo 0

End Sub

Private Sub mnuInitialize_Click()

If SiteDBname = "" Then
    MsgBox ("Before you can initial the station,  you need to open or create a site database.  Use Database Open or New first.")
    Exit Sub
End If

If lblEDMWarning.Visible = True Then
    MsgBox ("The type of total station has not yet been set.  Use Station Select Total Station first.")
    Exit Sub
End If

If DatumTB.BOF And DatumTB.EOF Then
    MsgBox ("There are no datums defind.  Use Edit Datums first to add new datums and then initialize the station.")
    Exit Sub
End If

frmStationSetup.Show 1

End Sub

Private Sub mnuNewCFG_Click()
If txtCFG <> "" Then
    response = MsgBox("Retain current settings and fields?", vbYesNo)
Else
    response = vbNo
    If response = vbNo Then
        Vars = 6
        For I = 1 To Vars
            VCarry(I) = False
            VMenu(I) = ""
        Next I
        VarList(1) = "UNIT"
        VPrompt(1) = "Unit"
        VLen(1) = 6
        VType(1) = "TEXT"
        VarList(2) = "ID"
        VPrompt(2) = "ID"
        VLen(2) = 5
        VType(2) = "TEXT"
        VarList(3) = "SUFFIX"
        VPrompt(3) = "Suffix"
        VLen(3) = 10
        VType(3) = "NUMERIC"
        VarList(4) = "PRISM"
        VPrompt(4) = "Prism"
        VLen(4) = 10
        VType(4) = "NUMERIC"
        VarList(5) = "X"
        VPrompt(5) = "X"
        VLen(5) = 10
        VType(5) = "NUMERIC"
        VarList(6) = "Y"
        VPrompt(6) = "Y"
        VLen(6) = 10
        VType(6) = "NUMERIC"
        VarList(7) = "Z"
        VPrompt(7) = "Z"
        VLen(7) = 10
        VType(7) = "NUMERIC"
        mnuFindUnit.Checked = False
        lblAutoFind.Visible = False
        LimitChecking = False
        PointTableName = ""
        SiteName = ""
        SiteDBname = ""
        SqidCheck = False
        UnitFieldString = ""
    End If
End If
Set SiteDB = Nothing
Set PointsTB = Nothing

start:
cd.CancelError = True
On Error GoTo cdcancel
cd.Filter = "CFG Files (*.cfg)|*.cfg"
cd.filename = CFGName
cd.ShowSave
If cd.filename = "" Then Exit Sub
A = Dir(cd.filename)
If A <> "" Then
    response = MsgBox("Overwrite existing file?", vbYesNo)
    If response = vbNo Then
        GoTo start
    End If
End If

CFGName = cd.filename



frmEditField.Show 1


cdcancel:
End Sub

Private Sub mnuNewDB_Click()

response = MsgBox("Create new database based on " + CFGName + "?", vbYesNo)
If response = vbNo Then Exit Sub
Loading = True
cd.filename = Left(CFGName, Len(CFGName) - 4)

cd.CancelError = True
On Error GoTo cdcancel

cd.Filter = "Site Files (*.mdb)|*.mdb"

cd.ShowSave

If cd.filename <> "" Then
    
    If Dir$(cd.filename) <> "" Then
        answer = MsgBox(cd.filename + " already exists.  Overwrite?", vbQuestion + vbYesNo)
        If answer = 7 Then Exit Sub
        ClearDBfields
        On Error Resume Next
        
        Kill cd.filename
        If Err <> 0 Then
            MsgBox cd.filename + " could not be erased." + Chr$(13) + "Ensure that it is not already open in another application.", vbInformation + vbOKOnly
            Exit Sub
        End If
    End If
    SiteDBname = cd.filename
    
    Call createsitedb(cd.filename)
    OpenSite SiteDBname
    frmMain.Caption = "EDM" + " - " + parsefilename$(cd.filename)
    TxtDB = LCase(cd.filename)

cdcancel:
End If

End Sub


Public Sub mnuNewPointsTB_Click()
If SiteDBname$ = "" Then
    MsgBox "Open or create a site file before creating a points file.", vbInformation
    Exit Sub
End If
If CFGName = "" Or Vars = 0 Then
    MsgBox ("Open or create a CFG file with fields before creating a points file")
    Exit Sub
End If
frmPointfiles.Show
If Not Cancelling Then
    txtPT = PointTableName
End If



End Sub


Private Sub mnuOpenCFG_Click()
start:
cd.CancelError = True
On Error GoTo cdcancel

cd.Filter = "CFG Files (*.cfg)|*.cfg"
cd.ShowOpen
If cd.filename <> "" Then
    CFGName = cd.filename
    parsecfg A
    LastPath = GetPath(cd.filename)
End If

cdcancel:
End Sub

Private Sub mnuOpenDB_Click()


cd.CancelError = True
On Error GoTo cdcancel

cd.Filter = "Site Files (*.mdb)|*.mdb"
cd.ShowOpen
If cd.filename <> "" Then
    ClearDBfields
    PointTableName = ""
    txtCurrentRecord = 0
    txtTotalRecords = 0
    SiteDBname$ = cd.filename
    Call OpenSite(SiteDBname$)
End If

cdcancel:
End Sub


Private Sub mnuOpenPointsTB_Click()

If SiteDBname$ = "" Then
    MsgBox "Open or create a site file before opening a points file.", vbInformation
    Exit Sub
End If

frmPointfiles.Show 1

End Sub


Private Sub mnuPrinter_Click()

Screen.MousePointer = 11
frmPrinter.Show 1
End Sub













Private Sub mnuPrintShots_Click()

End Sub

Private Sub mnuPrismPrompt_Click()
If mnuPrismPrompt.Checked Then
    mnuPrismPrompt.Checked = False
Else
    mnuPrismPrompt.Checked = True
End If

Dim Inidata(1, 2) As String
Dim IniClass As String
Dim Status As Byte
IniClass = "[EDM]"
Inidata(1, 1) = "PrismPrompt"
If mnuPrismPrompt.Checked Then
    Inidata(1, 2) = "YES"
Else
    Inidata(1, 2) = "NO"
End If

Call WriteIni(CFGName, IniClass, Inidata(), Status)
End Sub

Private Sub mnuSaveCFGas_Click()

start:
cd.CancelError = True
On Error GoTo cdcancel
cd.Filter = "CFG Files (*.cfg)|*.cfg"
cd.filename = CFGName
cd.ShowSave
If cd.filename = "" Then Exit Sub
A = Dir(cd.filename)
If A <> "" Then
    response = MsgBox("Overwrite existing file?", vbYesNo)
    If response = vbNo Then
        GoTo start
    End If
End If

CFGName = cd.filename
txtCFG = CFGName
Open CFGName For Output As 1
Print #1, "[EDM]"
Print #1, "Sitename="; SiteName
Print #1, "Database="; SiteDBname
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
Print #1, "Instrument="; EDMName
Print #1, "COMport"; comport
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

For I = 1 To Vars
    Print #1, "[" + VarList(I) + "]"
    Print #1, "Prompt="; VPrompt(I)
    Print #1, "Length="; VLen(I)
    Print #1, "Type="; VType(I)
    If VType(I) = "MENU" Then
        Print #1, "Menu=" + VMenu(I)
    End If
    If VCarry(I) Then
        Print #1, "Carry=True"
    End If
    Print #1, ""
Next I
Close 1


cdcancel:


End Sub

Private Sub mnuStationStatus_Click()
    
If Not StationInitialized Then
    MsgBox ("Station Not Initialized")
Else
    TempString = "Current station: " + Trim(CurrentStation.Name) + Chr(13)
    TempString = TempString + "   X: " + Format(CurrentStation.X, "#####0.000") + Chr(13)
    TempString = TempString + "   Y: " + Format(CurrentStation.y, "#####0.000") + Chr(13)
    TempString = TempString + "   Z: " + Format(CurrentStation.z, "#####0.000") + Chr(13)
    MsgBox (TempString)
End If
End Sub


Private Sub mnuStationVerify_Click()
If Not StationInitialized Then
    MsgBox ("Station Not Initialized")
Else
    frmStationVerify.Show 1
End If
End Sub

Private Sub mnuTheodolite_Click()
Screen.MousePointer = 11
frmTheodolite.Show

End Sub


Private Sub mnuViewAll_Click()
If SiteDBname = "" Then
    MsgBox ("Open database before performing this operation")
    Exit Sub
End If


If mnuViewAll.Caption = "All" Then
    mnuViewDatums.Checked = True
    mnuViewUnits.Checked = True
    mnuViewPoints.Checked = True
    mnuViewAll.Caption = "None"
    frmPlot.SetScale
    frmPlot.PlotPoints
    frmPlot.Show

Else
    mnuViewDatums.Checked = False
    mnuViewUnits.Checked = False
    mnuViewPoints.Checked = False
    mnuViewAll.Caption = "All"
    Unload frmPlot
End If

End Sub

Public Sub mnuViewDatums_Click()
If SiteDBname = "" Then
    MsgBox ("Open database before performing this operation")
    Exit Sub
End If


If mnuViewDatums.Checked Then
    mnuViewDatums.Checked = False
    If mnuViewPoints.Checked = False And mnuViewUnits.Checked = False Then
        mnuViewAll.Caption = "All"
    Else
        mnuViewAll.Caption = "None"
    End If
    
    If mnuViewPoints.Checked Or mnuViewUnits.Checked Then
        frmPlot.SetScale
        frmPlot.PlotPoints
    Else
        Unload frmPlot
    End If
Else
    mnuViewDatums.Checked = True
    mnuViewAll.Caption = "None"
    frmPlot.Show
    frmPlot.SetScale
    frmPlot.PlotPoints
End If

End Sub

Private Sub mnuViewHide_Click()
If mnuViewHide.Caption = "Hide" Then
    frmPlot.Hide
    mnuViewHide.Caption = "Show"
Else
    frmPlot.Show
    mnuViewHide.Caption = "Hide"
End If

End Sub

Private Sub mnuViewPoints_Click()
If SiteDBname = "" Then
    MsgBox ("Open database before performing this operation")
    Exit Sub
End If

If mnuViewPoints.Checked Then
    mnuViewPoints.Checked = False
    If mnuViewDatums.Checked = False And mnuViewUnits.Checked = False Then
        mnuViewAll.Caption = "All"
    Else
        mnuViewAll.Caption = "None"
    End If
    If mnuViewDatums.Checked Or mnuViewUnits.Checked Then
        frmPlot.SetScale
        frmPlot.PlotPoints
    Else
        Unload frmPlot
    End If
Else
    mnuViewPoints.Checked = True
    mnuViewAll.Caption = "None"
    frmPlot.Show
    frmPlot.SetScale
    frmPlot.PlotPoints
End If
End Sub

Private Sub mnuViewUnits_Click()
If SiteDBname = "" Then
    MsgBox ("Open database before performing this operation")
    Exit Sub
End If


If mnuViewUnits.Checked Then
    mnuViewUnits.Checked = False
    If mnuViewPoints.Checked = False And mnuViewDatums.Checked = False Then
        mnuViewAll.Caption = "All"
    Else
        mnuViewAll.Caption = "None"
    End If
    If mnuViewDatums.Checked Or mnuViewPoints.Checked Then
        frmPlot.SetScale
        frmPlot.PlotPoints
    Else
        Unload frmPlot
    End If
Else
    mnuViewUnits.Checked = True
    mnuViewAll.Caption = "None"
    frmPlot.Show
    frmPlot.SetScale
    frmPlot.PlotPoints
End If

End Sub


Private Sub NumberBox_GotFocus(Index As Integer)
NumberBox(Index).SelStart = 0
NumberBox(Index).SelLength = Len(NumberBox(Index))
OrigValue = NumberBox(Index)

End Sub


Private Sub NumberBox_KeyPress(Index As Integer, KeyAscii As Integer)

If Len(NumberBox(Index)) > VLen(Index) Then
    Exit Sub

ElseIf KeyAscii = 13 Then
    If OrigValue = NumberBox(Index) Then
        Picture1.SetFocus
    Else
        UpdatePointsTable VarList(Index), NumberBox(Index), 1, 1
        CheckFields
    End If
ElseIf KeyAscii = 27 Then
    Picture1.SetFocus

End If
End Sub


Private Sub NumberBox_LostFocus(Index As Integer)
If NumberBox(Index) <> OrigValue Then
    response = MsgBox("Update value of " + VarLabel(Index) + " to " + TextBox(Index), vbYesNo)
    If response = vbYes Then
        NumberBox_KeyPress Index, 13
    Else
        NumberBox(Index) = OrigValue
    End If
End If

End Sub


Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If PointTableName = "" Then Exit Sub
If PointsTB.RecordCount = 0 Then Exit Sub
Loading = True
If KeyCode = vbKeyHome Then
    PointsTB.AbsolutePosition = 0
    ShowValues
ElseIf KeyCode = vbKeyEnd Then
    PointsTB.AbsolutePosition = PointsTB.RecordCount - 1
    ShowValues
ElseIf KeyCode = vbKeyUp Then
    If PointsTB.AbsolutePosition > 1 Then
        PointsTB.AbsolutePosition = PointsTB.AbsolutePosition - 1
    Else
        PointsTB.AbsolutePosition = 0
    End If
    ShowValues
ElseIf KeyCode = vbKeyPageUp Then
    If PointsTB.AbsolutePosition > 10 Then
        PointsTB.AbsolutePosition = PointsTB.AbsolutePosition - 10
    Else
        PointsTB.AbsolutePosition = 0
    End If
    ShowValues
ElseIf KeyCode = vbKeyDown Then
    If PointsTB.AbsolutePosition < PointsTB.RecordCount - 1 Then
        PointsTB.AbsolutePosition = PointsTB.AbsolutePosition + 1
    Else
        PointsTB.AbsolutePosition = PointsTB.RecordCount - 1
    End If
    ShowValues
ElseIf KeyCode = vbKeyPageDown Then
    If PointsTB.AbsolutePosition < PointsTB.RecordCount - 10 Then
        PointsTB.AbsolutePosition = PointsTB.AbsolutePosition + 10
    Else
        PointsTB.AbsolutePosition = PointsTB.RecordCount - 1
    End If
    ShowValues
ElseIf KeyCode = vbKeyDelete Then
    DeleteRecord
ElseIf KeyCode = vbKeyAdd Then
    cmdPlusShot_Click
    ShowValues
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



Private Sub TextBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Len(TextBox(Index)) > VLen(Index) Then Exit Sub

End Sub

Private Sub TextBox_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 And Trim(TextBox(Index)) <> "" Then
    If OrigValue = TextBox(Index) Then
        Picture1.SetFocus
    Else
        UpdatePointsTable VarList(Index), TextBox(Index), 1, 1
        CheckFields
        OrigValue = TextBox(Index)
    End If
ElseIf KeyAscii = 27 Then
    Picture1.SetFocus
End If

End Sub


Private Sub TextBox_LostFocus(Index As Integer)
If TextBox(Index) <> OrigValue Then
    response = MsgBox("Update value of " + VarLabel(Index) + " to " + TextBox(Index), vbYesNo)
    If response = vbYes Then
        TextBox_KeyPress Index, 13
    Else
        TextBox(Index) = OrigValue
    End If
End If
    
    
End Sub




Private Sub txtHangle_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp, vbKeyPageUp, vbKeyPageDown, vbKeyDown, vbKeyHome, vbKeyEnd, vbKeyDelete, vbKeyRight, vbKeyLeft
        Picture1.SetFocus
End Select

End Sub


Private Sub txtID_Click()
Dim NewID As String
If PointTableName = "" Then
    MsgBox ("Open point table first")
    txtID = OriginalID
    Exit Sub
End If
If Not Loading Then
    NewID = txtID
    txtID = OriginalID
    If txtSuffix > 0 Then
        response = MsgBox("Change ID for all points in this series?", vbYesNo)
        If response = vbYes Then
            UpdatePointsTable "id", NewID, 1, 1
            If Val(Trim(txtID)) = 0 Then
                DecrementID OriginalUnit, OriginalID, OriginalSuffix
            End If
        
        Else
            If PointsTB.AbsolutePosition <> PointsTB.RecordCount - 1 Then
                bkmrk = PointsTB.Bookmark
                PointsTB.MoveNext
                If PointsTB("suffix") = txtSuffix + 1 Then
                    MsgBox ("You cannot change ID in the middle of a series")
                    Exit Sub
                End If
                PointsTB.Bookmark = bkmrk
            End If
            UpdatePointsTable "id", NewID, 1, 0
            UpdatePointsTable "suffix", 0, 0, 0
            txtSuffix = 0
        End If
    Else
        If Val(Trim(txtID)) = 0 Then
            DecrementID OriginalUnit, OriginalID, OriginalSuffix
        End If
        UpdatePointsTable "id", NewID, 1, 0
    End If
    DoEvents
    txtID = NewID
    OriginalID = txtID
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
If Val(Trim(txtID)) > 0 Then
    txtID.AddItem hash(5)
Else
    SqlString = "select max(id) from [*units] where unit='" + txtUnit + "' and id<'A'"
    Set RsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
    If IsNull(RsTemp(0)) Then
        txtID.AddItem PadID("1")
    Else
        txtID.AddItem PadID(Str(RsTemp(0) + 1))
    End If
End If

End Sub


Private Sub txtID_GotFocus()
    OriginalUnit = txtUnit
    OriginalID = txtID
    OriginalSuffix = Val(txtSuffix)
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Picture1.SetFocus
End If
End Sub

Private Sub txtID_LostFocus()

txtUnit = OriginalUnit
txtID = OriginalID
txtSuffix = OriginalSuffix

End Sub


Private Sub txtPoleHT_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Picture1.SetFocus
End If
End Sub


Private Sub txtPrism_Click()
If txtPrism.ListIndex > -1 Then
    txtPoleHT = Format(PoleHeight(txtPrism.ItemData(txtPrism.ListIndex)), "####0.000")
    txtPrism_KeyPress 13
End If
End Sub

Private Sub txtPrism_GotFocus()
OriginalPoleHT = Val(txtPoleHT)
OriginalPrismIndex = txtPrism.ListIndex
txtPrism.SelStart = 0
txtPrism.SelLength = Len(txtPrism)

End Sub

Private Sub txtPrism_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Picture1.SetFocus
ElseIf KeyAscii = 13 Then
    If Not Loading And txtCurrentRecord > 0 Then
        If PoleHeight(txtPrism.ItemData(txtPrism.ListIndex)) <> OriginalPoleHT Then
            response = MsgBox("Update Z value from " & txtXYZ(2) & " to " & Format(Val(txtXYZ(2)) + OriginalPoleHT - PoleHeight(txtPrism.ItemData(txtPrism.ListIndex)), "#####0.000") & "?", vbYesNo)
            If response = vbYes Then
                txtPoleHT = Format(PoleHeight(txtPrism.ItemData(txtPrism.ListIndex)), "#####0.000")
                txtXYZ(2) = Format(Val(txtXYZ(2)) + OriginalPoleHT - PoleHeight(txtPrism.ItemData(txtPrism.ListIndex)), "#####0.000")
                UpdatePointsTable "prism", txtPoleHT, 0, 0
                UpdatePointsTable "z", txtXYZ(2), 1, 0
            Else
                txtPrism.ListIndex = OriginalPrismIndex
            End If
        End If
    End If
    txtPoleHT = PoleHeight(txtPrism.ItemData(txtPrism.ListIndex))
    OriginalPoleHT = txtPoleHT
    OriginalPrismIndex = txtPrism.ListIndex

End If
End Sub

Private Sub txtPrism_LostFocus()
txtPrism.SelLength = 0
txtPoleHT = OriginalPoleHT
If txtPrism.ListCount > 0 Then
    txtPrism.ListIndex = OriginalPrismIndex
End If

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


Private Sub txtPT_LostFocus()

If PointTableName = "" Then
    lblPointsWarning.Visible = True
Else
    txtPT = PointTableName
End If
End Sub

Private Sub txtSlopeD_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp, vbKeyPageUp, vbKeyPageDown, vbKeyDown, vbKeyHome, vbKeyEnd, vbKeyDelete, vbKeyRight, vbKeyLeft
        Picture1.SetFocus
End Select

End Sub


Private Sub txtSuffix_GotFocus()
txtSuffix.SelStart = 0
txtSuffix.SelLength = Len(txtSuffix)
End Sub

Public Sub ShowValues()

If PointTableName = "" Then Exit Sub
Loading = True
txtCurrentRecord = PointsTB.AbsolutePosition + 1
txtTotalRecords = PointsTB.RecordCount

If PointsTB.RecordCount > 0 Then
    lblBlankFields.Visible = False
    txtUnit = PointsTB("Unit")
    txtID = PadID(PointsTB("ID"))
    txtSuffix = PointsTB("suffix")
    OriginalID = txtID
    OriginalUnit = txtUnit
    OriginalSuffix = Val(txtSuffix)
    DoEvents
    If txtXYZ(0).Visible Then txtXYZ(0) = Format(PointsTB("x"), "#########0.000")
    If txtXYZ(1).Visible Then txtXYZ(1) = Format(PointsTB("y"), "#########0.000")
    If txtXYZ(2).Visible Then txtXYZ(2) = Format(PointsTB("z"), "#########0.000")
    If txtVangle.Visible Then txtVangle = Format(PointsTB("vangle"), "#########0.0000")
    If txtHangle.Visible Then txtHangle = Format(PointsTB("hangle"), "#########0.0000")
    If txtSlopeD.Visible Then txtSlopeD = Format(PointsTB("sloped"), "#########0.0000")
    txtXYZ(0).Refresh
    If txtPoleHT.Visible Then
        If Not IsNull(PointsTB("Prism")) Then
            OriginalPoleHT = PointsTB("prism")
            txtPoleHT = Format(PointsTB("prism"), "#####0.000")
            For I = 0 To txtPrism.ListCount - 1
                If PoleHeight(txtPrism.ItemData(I)) = PointsTB("prism") Then
                    txtPrism.ListIndex = I
                    Exit For
                End If
            Next I
        Else
            OriginalPoleHT = 0
            txtPoleHT = 0
            txtPrism = ""
            
        End If
    End If
    On Error Resume Next
    For I = 1 To Vars
        
        Select Case VType(I)
            Case "TEXT"
                If IsNull(PointsTB(VarList(I))) Then
                    If VarList(I) = "DATE" Then
                        TextBox(I) = Date
                    ElseIf VarList(I) = "TIME" Then
                        TextBox(I) = Time
                    ElseIf VarList(I) = "SITENAME" Then
                        TextBox(I) = SiteName
                    Else
                        TextBox(I) = ""
                    End If
                Else
                    TextBox(I).Text = PointsTB(VarList(I))
                End If
            Case "MENU"
                If IsNull(PointsTB(VarList(I))) Then
                    MenuBox(I) = ""
                Else
                    MenuBox(I) = PointsTB(VarList(I))
                End If
            Case "NUMERIC", "INSTRUMENT"
                If IsNull(PointsTB(VarList(I))) Then
                    NumberBox(I) = ""
                Else
                    NumberBox(I) = PointsTB(VarList(I))
                End If
        End Select
    Next I
    On Error GoTo 0
    If frmMain.mnuViewPoints.Checked Then
        frmPlot.shpPoint.Left = PointsTB(PlotX) - frmPlot.shpPoint.Width / 2
        frmPlot.shpPoint.Top = PointsTB(PlotY) + frmPlot.shpPoint.Height / 2
    End If
    Loading = False
Else
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
    If txtSlopeD.Visible Then txtSlopeD = ""
    On Error Resume Next
    For I = 1 To Vars
        Select Case VType(I)
            Case "TEXT"
                If VarList(I) = "DATE" Then
                    TextBox(I) = Date
                ElseIf VarList(I) = "TIME" Then
                    TextBox(I) = Time
                ElseIf VarList(I) = "SITENAME" Then
                    TextBox(I) = SiteName
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

On Error Resume Next
Picture1.SetFocus
On Error GoTo 0
End Sub

Private Sub txtSuffix_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Picture1.SetFocus
End If
End Sub

Private Sub txtUnit_Click()
txtUnit_KeyPress 13
End Sub

Private Sub txtUnit_DropDown()
If UnitTB.RecordCount > 0 Then
    txtUnit.Clear
    UnitTB.MoveFirst
    While Not UnitTB.EOF
        txtUnit.AddItem UnitTB("Unit")
        UnitTB.MoveNext
    Wend
End If
        
End Sub


Private Sub txtUnit_GotFocus()
    OriginalUnit = txtUnit
    OriginalID = txtID
    OriginalSuffix = Val(txtSuffix)
End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Picture1.SetFocus
ElseIf KeyAscii = 13 Then
    If Not Loading And txtCurrentRecord > 0 And txtUnit <> OriginalUnit And PointTableName <> "" Then
        If txtSuffix = 0 And PointsTB.AbsolutePosition = PointsTB.RecordCount - 1 Then
            PointsTB.Edit
            PointsTB("unit") = txtUnit
            If Val(Trim(txtID)) > 0 Then
                DecrementID OriginalUnit, OriginalID, OriginalSuffix
                SqlString = "select max(id), min(minx)  from [*units] where unit='" + txtUnit + "' and id<'A'"
                Set RsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
                If IsNull(RsTemp(0)) Then
                    PointsTB("id") = PadID("1")
                Else
                    PointsTB("id") = PadID(Str(RsTemp(0) + 1))
                End If
            End If
            PointsTB.Update
            txtID = PointsTB("id")
            FillUnitFields
            txtID = PointsTB("id")
            OriginalUnit = txtUnit
            OriginalID = txtID
            OriginalSuffix = Val(txtSuffix)
            UpdateUnitTable txtUnit, txtID, IDonly
            If RsTemp(1) = -99999 And mnuFindUnit.Checked Then
                response = MsgBox("This unit has no limits defined.  Turn off Auto-Find Units?", vbYesNo)
                If response = vbYes Then
                    'mnuFindUnit.Checked = False
                    mnuFindUnit_Click
                End If
            End If
            Set RsTemp = Nothing
                
        Else
        End If
    ElseIf txtCurrentRecord = 0 Then
        OriginalUnit = txtUnit
    End If
    On Error Resume Next
    Picture1.SetFocus
    On Error GoTo 0
End If

End Sub


Private Sub txtUnit_LostFocus()
txtUnit = OriginalUnit
txtID = OriginalID
txtSuffix = OriginalSuffix
End Sub


Private Sub txtVangle_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp, vbKeyPageUp, vbKeyPageDown, vbKeyDown, vbKeyHome, vbKeyEnd, vbKeyDelete, vbKeyRight, vbKeyLeft
        Picture1.SetFocus
End Select

End Sub










Public Sub UpdatePointsTable(Variable As String, Value As String, ShowMessage As Byte, Scope As Byte)
Dim CurrentPosition As Long
Dim NRecords As Integer

CurrentPosition = PointsTB.AbsolutePosition

If txtCurrentRecord = 0 Then Exit Sub

If Scope = 0 Then
    PointsTB.Edit
    If UCase(Variable) = "ID" Then
        PointsTB(Variable) = PadID(Value)
        UpdateUnitTable txtUnit, PadID(Value), IDonly
    Else
        PointsTB(Variable) = Value
    End If
    PointsTB.Update
    NRecords = 1
Else

    SqlString = "select [*reccounter] from [" + PointTableName + "] where unit='" + txtUnit + "' and id='" + txtID + "'"
    Set RsTemp = SiteDB.OpenRecordset(SqlString, dbOpenSnapshot)
    If RsTemp.EOF Then Exit Sub
    Set PointsTB = SiteDB.OpenRecordset(PointTableName, dbOpenTable)
    PointsTB.Index = "recordcounter"
    While Not RsTemp.EOF
        NRecords = NRecords + 1
        PointsTB.Seek "=", RsTemp(0)
        PointsTB.Edit
        If UCase(Variable) = "ID" Then
            PointsTB(Variable) = PadID(Value)
            UpdateUnitTable txtUnit, PadID(Value), IDonly
        Else
            PointsTB(Variable) = Value
        End If
        PointsTB.Update
        RsTemp.MoveNext
    Wend
    For I = 1 To nUnitFields
        If UCase(Variable) = UCase(Unitfield(I)) Then
            UpdateUnitTable txtUnit, txtID, Everything
            Exit For
        End If
    Next I
End If
Set PointsTB = SiteDB.OpenRecordset(PointTableName, dbOpenDynaset)
PointsTB.MoveLast
PointsTB.AbsolutePosition = CurrentPosition
If ShowMessage = 1 Then
    MsgBox (NRecords & " shot(s) on " + txtUnit + "-" + Trim(txtID) + " changed")
End If

End Sub

Public Sub DecrementID(Unit As String, ID As String, Suffix As Integer)
Dim MaxID As Long

If Val(ID) = 0 Then Exit Sub
UnitTB.Index = "unitname"
UnitTB.Seek "=", Unit
    
If Not UnitTB.NoMatch Then
    If IsNull(UnitTB("id")) Then
        MaxID = 0
    Else
        MaxID = Val(UnitTB("id"))
    End If
    If Val(ID) = MaxID Then
        UnitTB.Edit
        UnitTB("id") = PadID(Str(Val(ID) - 1))
        UnitTB.Update
    End If
End If
End Sub

Public Sub IncrementID()

SqlString = "select id from [*units] where unit='" + txtUnit + "' and id<'A'"
Set RsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)

If RsTemp.EOF Or IsNull(RsTemp(0)) Then
    txtID = 1
Else
    txtID = RsTemp(0) + 1
End If
txtID = PadID(txtID)

End Sub

Public Sub Take_Shot()

Dim edmpoffset As Single
Dim edmshot As shotdata

Select Case UCase(EDMName)
Case "SIMULATE"
    Randomize
    X = Int((3) * Rnd) + 999 + Rnd
    y = Int((1) * Rnd) + 1013 + Rnd
    z = Rnd
    edmshot.X = X
    edmshot.y = y
    edmshot.z = z
    edmshot.hangle = 111.505
    edmshot.vangle = 98.2525
    edmshot.sloped = Sqr(edmshot.X ^ 2 + edmshot.y ^ 2 + edmshot.z ^ 2)

Case Else
    Call recordpoint(returndata$)
    Call parsenez(returndata$, edmshot, edmpoffset, mesunits$, angleunit$, errorcode)
    If errorcode = 0 Then
        Call vhdtonez(edmshot)
    End If
End Select

If txtPrism.ListIndex <> -1 Then
    edmshot.poleh = PoleHeight(txtPrism.ItemData(txtPrism.ListIndex))
Else
    edmshot.poleh = 0
End If

edmshot.X = CurrentStation.X + edmshot.X
edmshot.y = CurrentStation.y + edmshot.y
edmshot.z = CurrentStation.z + edmshot.z - edmshot.poleh

txtXYZ(0) = Format(edmshot.X, "######0.000")
txtXYZ(1) = Format(edmshot.y, "######0.000")
txtXYZ(2) = Format(edmshot.z, "######0.000")

txtVangle = edmshot.vangle
txtHangle = edmshot.hangle
txtSlopeD = edmshot.sloped

If mnuPrismPrompt.Checked Then
    MsgBox ("Verify Prism.  Currently set to " + txtPrism.List(txtPrism.ListIndex))
End If

PointsTB.AddNew
PointsTB("Unit") = txtUnit
PointsTB("ID") = PadID(txtID)
PointsTB("suffix") = txtSuffix
PointsTB("prism") = edmshot.poleh
If txtXYZ(0).Visible Then PointsTB("x") = txtXYZ(0)
If txtXYZ(1).Visible Then PointsTB("y") = txtXYZ(1)
If txtXYZ(2).Visible Then PointsTB("z") = txtXYZ(2)
If txtVangle.Visible Then PointsTB("vangle") = txtVangle
If txtHangle.Visible Then PointsTB("hangle") = txtHangle
If txtSlopeD.Visible Then PointsTB("sloped") = txtSlopeD

On Error Resume Next
For I = 1 To Vars
    Select Case VType(I)
        Case "TEXT"
            If VCarry(I) Then
                PointsTB(VarList(I)) = TextBox(I).Text
            End If
        Case "MENU"
            If VCarry(I) Then
                PointsTB(VarList(I)) = MenuBox(I)
            End If
        Case "NUMERIC", "INSTRUMENT"
            If VCarry(I) Then
                PointsTB(VarList(I)) = NumberBox(I)
            End If

    End Select
Next I
On Error GoTo 0

PointsTB.Update
PointsTB.MoveLast

txtCurrentRecord = PointsTB.AbsolutePosition + 1
txtTotalRecords = PointsTB.RecordCount

End Sub

Public Sub DeleteRecord()
Dim NDeletions As Integer
Dim PreviousRecord As Integer


response = MsgBox("Warning:  This action will permanently remove records.  Continue anyway", vbYesNo)
If response = vbNo Then Exit Sub

SqlString = "select [*reccounter] from [" + PointTableName + "] where unit='" + txtUnit + "' and id='" + txtID + "' order by [*reccounter]"
Set RsTemp = SiteDB.OpenRecordset(SqlString, dbOpenDynaset)
If RsTemp.EOF Then Exit Sub
RsTemp.MoveFirst

RsTemp.MoveLast

NDeletions = 1
If RsTemp.RecordCount > 1 Then
    response = MsgBox("Delete all " & RsTemp.RecordCount & " records for " & txtUnit & "-" & Trim(txtID) & "?", vbYesNo)
    If response = vbYes Then
        DecrementID txtUnit, txtID, txtSuffix
        NDeletions = RsTemp.RecordCount
        SqlString = "delete * from " + PointTableName + " where unit='" + txtUnit + "' and id='" + txtID + "'"
        SiteDB.Execute SqlString
        Set PointsTB = SiteDB.OpenRecordset(PointTableName, dbOpenDynaset)
    Else
        If txtSuffix = 0 Then
            MsgBox ("You cannot delete the first shot in a series.  You must delete the entire object")
            Exit Sub
        End If
        PointsTB.Delete
    End If
Else
    DecrementID txtUnit, txtID, txtSuffix
    PointsTB.Delete
End If
Set RsTemp = Nothing
MsgBox (NDeletions & " records deleted")
PointsTB.MoveLast
ShowValues
End Sub

Private Sub txtXYZ_Click(Index As Integer)

If PointTableName = "" Then
    MsgBox ("Open point table first")
    txtXYZ(Index) = OriginalXYZ
    Exit Sub
End If

frmOffSet.Caption = txtXYZ(Index)
Set frmOffSet.CallingBox = txtXYZ(Index)
Select Case Index
    Case 0
        frmOffSet.Varname = "X"
    Case 1
        frmOffSet.Varname = "Y"
    Case 2
        frmOffSet.Varname = "Z"
End Select
frmOffSet.Show 1
If txtXYZ(Index) <> OriginalXYZ Then
    Select Case Index
        Case 0
            UpdatePointsTable "x", txtXYZ(Index), 1, 0
        Case 1
            UpdatePointsTable "y", txtXYZ(Index), 1, 0
        Case 2
            UpdatePointsTable "z", txtXYZ(Index), 1, 0
    End Select
    OriginalXYZ = txtXYZ(Index)
End If
ChangingXYZ = True
Picture1.SetFocus
'ShowValues

End Sub

Private Sub txtXYZ_GotFocus(Index As Integer)
On Error Resume Next
OriginalXYZ = txtXYZ(Index)
On Error GoTo 0
End Sub


Private Sub txtXYZ_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
    
    Picture1.SetFocus
End If

End Sub


Private Sub txtXYZ_LostFocus(Index As Integer)
    txtXYZ(Index) = OriginalXYZ
End Sub



Public Sub UpdateUnitTable(UnitName As String, ID As String, Scope)
UnitTB.Index = "UnitName"
UnitTB.Seek "=", UnitName
If Not UnitTB.NoMatch Then
    UnitTB.Edit
    If Val(Trim(ID)) > 0 Then
        UnitTB("id") = PadID(ID)
    End If
    UnitTB("suffix") = txtSuffix
    If Scope = Everything Then
        For I = 1 To nUnitFields
            Select Case UCase(Unitfield(I))
                Case "ID", "UNIT"
                Case "PRISM"
                    UnitTB("PRISM") = txtPoleHT
                Case Else
                    If PointsTB(Unitfield(I)) = "" Then
                        UnitTB(Unitfield(I)) = " "
                    Else
                        UnitTB(Unitfield(I)) = PointsTB(Unitfield(I))
                    End If
            End Select
        Next I
    End If
    UnitTB.Update
End If

End Sub

Public Function FindUnit()
Dim X As Single
Dim y As Single
Dim RsTemp As Recordset

start:
Cancelling = False
SqlString = "select * from [*units] where minx< " & PointsTB("x") & " and maxx>" & PointsTB("x") & " and miny<" & PointsTB("y") & " and maxy>" & PointsTB("y")
Set RsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
If Not RsTemp.EOF Then
    If Not IsNull(RsTemp("unit")) Then
        txtUnit = RsTemp("unit")
        txtUnit_KeyPress 13

        ShowValues
    End If
Else
    response = MsgBox("Object in undefined unit.  Define now?", vbYesNoCancel)
    If response = vbCancel Then
        Cancelling = True
        Exit Function
    ElseIf response = vbYes Then
        AddUnits.XY(0) = Int(txtXYZ(0))
        AddUnits.XY(1) = Int(txtXYZ(1))
        AddUnits.Show 1
        If Not Cancelling Then
            GoTo start
        End If
    End If
End If
End Function

Public Sub SetCurser()

End Sub

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
                Gotit = LCase(SiteDB.TableDefs("*units").Fields("prism").Name) = "prism"
                On Error GoTo 0
                If Gotit Then
                    If IsNull(UnitTB("PRISM")) Then
                        response = MsgBox("Invalid default prism.  Use " + txtPrism + "?", vbYesNo)
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
                            For J = 0 To txtPrism.ListCount - 1
                                If PoleHeight(txtPrism.ItemData(J)) = UnitTB("prism") Then
                                    txtXYZ(2) = Format(Val(txtXYZ(2)) + txtPoleHT - UnitTB("prism"), "######0.000")
                                    txtPoleHT = UnitTB("prism")
                                    txtPrism = txtPrism.List(J)
                                    UpdatePointsTable "prism", txtPoleHT, 0, 0
                                    UpdatePointsTable "z", txtXYZ(2), 0, 0
                                    Gotit = True
                                    Exit For
                                End If
                            Next J
                            If Not Gotit Then
                                response = MsgBox("Invalid default prism.  Use " + txtPrism + "?", vbYesNo)
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
                                If txtCurrentRecord > 0 Then
                                    UpdatePointsTable VarList(J), TextBox(J), 0, 0
                                End If
                            Case "NUMERIC"
                                If Not IsNull(UnitTB(Unitfield(I))) Then
                                    NumberBox(J) = UnitTB(Unitfield(I))
                                Else
                                    NumberBox(J) = ""
                                End If
                                If txtCurrentRecord > 0 Then
                                    UpdatePointsTable VarList(J), NumberBox(J), 0, 0
                                End If
                            Case "MENU"
                                If Not IsNull(UnitTB(Unitfield(I))) Then
                                    MenuBox(J) = UnitTB(Unitfield(I))
                                Else
                                    MenuBox(J) = ""
                                End If
                                If txtCurrentRecord > 0 Then
                                    UpdatePointsTable VarList(J), MenuBox(J), 0, 0
                                End If
                        End Select
                        Exit For
                    End If
                Next J
        End Select
    Next I
End If
    

End Sub

Public Sub ClearDBfields()
TxtDB = ""
txtPT = ""
Set SiteDB = Nothing
Set PointsTB = Nothing
Set UnitTB = Nothing
Set PoleTB = Nothing
Set DatumTB = Nothing
Set cfgTB = Nothing
On Error Resume Next
For I = 1 To 50
    Unload MenuBox(I)
    Unload TextBox(I)
    Unload NumberBox(I)
Next I

lblDBWarning.Visible = True
lblPointsWarning.Visible = True
lblPoleWarning.Visible = True
txtCurrentRecord = 0
txtTotalRecords = 0
txtUnit.Enabled = False
txtID.Enabled = False
txtPrism.Enabled = False



End Sub

Public Sub CheckFields()

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
On Error GoTo 0
End Sub

Public Function CheckStatus()
CheckStatus = False
If PointTableName = "" Then
    MsgBox ("Point table must be opened.")
    CheckStatus = True
ElseIf Not StationInitialized Then
    MsgBox ("Total Station not initialized.  Initialize before recording points")
    CheckStatus = True
ElseIf Not LimitChecking And txtUnit = "" Then
    MsgBox ("Select Unit before shooting, or set Auto-Find Unit")
    CheckStatus = True
ElseIf PoleTB.BOF And PoleTB.EOF Then
    MsgBox ("No prisms defined.  Define before taking a shot")
    CheckStatus = True
End If

End Function
