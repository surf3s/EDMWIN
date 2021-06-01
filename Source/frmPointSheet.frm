VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPointSheet 
   Caption         =   "Point Data Sheet"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   5820
   Begin VB.Data pointdata 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Width           =   3375
   End
   Begin MSDBGrid.DBGrid pointsheet 
      Bindings        =   "frmPointSheet.frx":0000
      Height          =   2295
      Left            =   480
      OleObjectBlob   =   "frmPointSheet.frx":0018
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
End
Attribute VB_Name = "frmPointSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Me.Left = frmMain.Left
Me.Width = frmMain.Width
Me.Height = frmMain.Height / 2
Me.Top = frmMain.Top + frmMain.Height - Me.Height
Call sizecontrols

Set pointdata.Recordset = PointsTB

pointdata.Caption = PointTableName$

End Sub

Private Sub sizecontrols()

If Me.WindowState <> 1 Then
    pointsheet.Width = Me.Width - BannerWidth
    pointsheet.Height = Me.Height - pointdata.Height - BannerHeight
    pointsheet.Top = 0
    pointsheet.Left = 0
    pointdata.Top = pointsheet.Top + pointsheet.Height
    pointdata.Left = 0
    pointdata.Width = pointsheet.Width
End If

End Sub

Private Sub Form_Resize()

Call sizecontrols

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set pointdata.Recordset = Nothing
End Sub

Private Sub pointsheet_DblClick()

frmEditrecord.Show

End Sub
