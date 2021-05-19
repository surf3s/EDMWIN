VERSION 5.00
Begin VB.Form frmPointfiles 
   Caption         =   "Open/Create Point File"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   4890
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Delete"
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox tablename 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ListBox tablelist 
      Height          =   1230
      Left            =   120
      TabIndex        =   3
      Top             =   1380
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "(Note:  Table name cannot contain spaces or special characters such as ', "", *, $, etc)"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a points file from the list in the upper box, or  to create a new one, type a name in the lower box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   4680
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPointfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)

Dim A As Integer
Dim flag As Boolean

If Trim(tablename) = "" Then
    Cancelling = True
    Unload Me
    Exit Sub
End If
Cancelling = False

Select Case Index
Case 0
    flag = False
    For A = 0 To SiteDB.TableDefs.Count - 1
        If LCase(SiteDB.TableDefs(A).Name) = LCase(tablename.Text) Then
            flag = True
            Exit For
        End If
    Next A
    PointTableName = tablename.Text
    
    If flag = False Then
        Call CreatePointTB(tablename.Text)
        flag = True
    End If
    If Cancelling Then
        Exit Sub
    End If
    frmMain.lblPointsWarning.Visible = False
    frmMain.txtXYZ(0).Enabled = True
    frmMain.txtXYZ(1).Enabled = True
    frmMain.txtXYZ(2).Enabled = True
    frmMain.txtUnit.Enabled = True
    frmMain.txtID.Enabled = True
    frmMain.txtprism.Enabled = True
    OpenPointsTable

Case 1
    Cancelling = True
    Unload Me
    
Case 2
    response = MsgBox("Permanently delete table " + tablename.Text + "?", vbYesNo)
    If response = vbYes Then
        SiteDB.TableDefs.Delete tablename.Text
        SiteDB.TableDefs.Refresh
        tablename.Text = ""
    End If
    GetPointTables
    For I = 1 To nPointTables
        tablelist.AddItem PointTable(I)
    Next I
    
End Select
Unload Me

End Sub

Private Sub Form_Load()

GetPointTables

For I = 1 To nPointTables
    tablelist.AddItem PointTable(I)
Next I

Call CenterForm(Me)
Call sizecontrols

End Sub

Private Sub Form_Resize()

Call sizecontrols

End Sub

Private Sub sizecontrols()

Command1(0).Left = Me.Width * 0.2 - Command1(0).Width / 2
Command1(1).Left = Me.Width * 0.8 - Command1(1).Width / 2

tablelist.Left = BannerWidth
tablelist.Width = Me.Width - 3 * BannerWidth
tablename.Width = tablelist.Width
tablename.Left = tablelist.Left

End Sub

Private Sub Form_Unload(Cancel As Integer)

Me.Hide
frmMain.Picture1.SetFocus

End Sub

Private Sub tablelist_Click()

tablename.Text = tablelist.Text

End Sub

Private Sub tablelist_DblClick()

tablename.Text = tablelist.Text
Command1_Click 0

End Sub


