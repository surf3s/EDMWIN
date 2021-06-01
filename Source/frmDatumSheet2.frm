VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDatumSheet2 
   Caption         =   "Datum Data Sheet"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Data datumdata 
      Caption         =   "Datums"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid datumsheet 
      Bindings        =   "frmDatumSheet2.frx":0000
      Height          =   1095
      Left            =   240
      OleObjectBlob   =   "frmDatumSheet2.frx":0018
      TabIndex        =   0
      Top             =   1080
      Width           =   3975
   End
End
Attribute VB_Name = "frmDatumSheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub sizecontrols()

If Me.WindowState = 1 Then Exit Sub

datumsheet.Width = Me.Width - BannerWidth
datumsheet.Height = Me.Height - datumdata.Height - BannerHeight
datumsheet.Top = 0
datumsheet.Left = 0
datumdata.Top = datumsheet.Top + datumsheet.Height
datumdata.Left = 0
datumdata.Width = datumsheet.Width

End Sub

Private Sub datumsheet_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case 13
    If datumsheet.Col + 1 = datumdata.Recordset.Fields.Count Then
        datumsheet.Col = 0
    Else
        datumsheet.Col = datumsheet.Col + 1
    End If
Case Else
End Select

End Sub
Private Sub Form_Load()

Me.Left = frmMain.Left
Me.Width = frmMain.Width
Me.Height = frmMain.Height / 2
Me.Top = frmMain.Top + BannerHeight * 2

Call sizecontrols

Set datumdata.Recordset = DatumTB

End Sub

Private Sub Form_Resize()

Call sizecontrols

End Sub
