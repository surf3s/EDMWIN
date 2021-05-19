VERSION 5.00
Begin VB.Form frmFilter 
   Caption         =   "Set Grid Filter"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstVars 
      Height          =   2205
      Left            =   270
      TabIndex        =   3
      Top             =   330
      Width           =   1755
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4110
      TabIndex        =   2
      Top             =   960
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Filter"
      Height          =   375
      Left            =   4110
      TabIndex        =   1
      Top             =   390
      Width           =   1185
   End
   Begin VB.ListBox lstVals 
      Height          =   2205
      Left            =   2220
      TabIndex        =   0
      Top             =   330
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Values"
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
      Index           =   1
      Left            =   2700
      TabIndex        =   5
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fields"
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
      Index           =   0
      Left            =   750
      TabIndex        =   4
      Top             =   60
      Width           =   660
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If lstVars.ListIndex = -1 Or lstVals.ListIndex = -1 Then
    MsgBox ("Select field and value to filter by.")
    Exit Sub
End If
frmDataGrid.Command1.Caption = "Filter:" + LCase(lstVars.List(lstVars.ListIndex)) + "=" + LCase(lstVals.List(lstVals.ListIndex))

End Sub

Private Sub Form_Load()

lstVars.Clear
lstVals.Clear

For Each cfield In frmMain.PointsADO.Recordset.Fields
    If LCase(cfield.Name) <> "recno" Then
        lstVars.AddItem LCase(cfield.Name)
    End If
Next
lstVars.ListIndex = 0

End Sub

Private Sub lstVars_Click()

Dim rsTemp As Recordset

lstVals.Clear
SqlString = "select distinct [" + lstVars.List(lstVars.ListIndex) + "] from [" + PointTableName + "]"
Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
If rsTemp.EOF Then
    MsgBox ("No values currently in this table")
    Exit Sub
End If
While Not rsTemp.EOF
    lstVals.AddItem rsTemp(0)
    rsTemp.MoveNext
Wend
Set rsTemp = Nothing

End Sub


