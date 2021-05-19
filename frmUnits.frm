VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmUnits 
   Caption         =   "Setup Units"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   5325
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   3372
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5292
      _ExtentX        =   9340
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   420
      TabCaption(0)   =   "Switch To"
      TabPicture(0)   =   "frmUnits.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "unitlist"
      Tab(0).Control(1)=   "Label1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Add/Edit"
      TabPicture(1)   =   "frmUnits.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "unitsheet"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "unitdata"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Data unitdata 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   276
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2880
         Visible         =   0   'False
         Width           =   1692
      End
      Begin MSDBGrid.DBGrid unitsheet 
         Bindings        =   "frmUnits.frx":0038
         Height          =   2292
         Left            =   240
         OleObjectBlob   =   "frmUnits.frx":004F
         TabIndex        =   4
         Top             =   720
         Width           =   4812
      End
      Begin VB.ListBox unitlist 
         Height          =   2400
         Left            =   -74760
         TabIndex        =   1
         Top             =   720
         Width           =   4452
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edit what unit?"
         Height          =   192
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   996
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Begin recording points in what unit?"
         Height          =   192
         Left            =   -74760
         TabIndex        =   2
         Top             =   480
         Width           =   2484
      End
   End
End
Attribute VB_Name = "frmUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim exitform As Boolean

Private Sub Form_Activate()

If exitform Then Unload Me

unitsheet.Columns(0).Width = 1000
unitsheet.Columns(1).Width = 800
unitsheet.Columns(2).Width = 800
unitsheet.Columns(3).Width = 800
unitsheet.Columns(4).Width = 800
        
Screen.MousePointer = 1

End Sub

Private Sub Form_Load()

exitform = False

SSTab1.Left = 0
SSTab1.Top = 0
Me.Width = SSTab1.Width + BannerWidth
Me.Height = SSTab1.Height + BannerHeight
SSTab1.Tab = 0


Call fillunitlist

SqlString = "SELECT Unit as [Unit Name], Minx as [Minimum X], Miny as [Minimum Y], Maxx as [Maximum X], Maxy as [Maximum Y] FROM [*UNITS]"
unitdata.DatabaseName = SiteDBname$
Set unitdata.Recordset = UnitTB
unitsheet.ReBind

End Sub

Private Sub fillunitlist()

unitlist.Clear
UnitTB.MoveFirst
If Not UnitTB.EOF Or Not UnitTB.BOF Then
    UnitTB.MoveFirst
    Do
         unitlist.AddItem UnitTB("Unit")
        UnitTB.MoveNext
    Loop Until UnitTB.EOF
End If

End Sub

Private Sub unitsheet_AfterDelete()

Call fillunitlist

End Sub

Private Sub unitsheet_AfterUpdate()

Call fillunitlist

End Sub

