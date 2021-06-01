VERSION 5.00
Begin VB.Form frmDatumShow 
   Caption         =   "Datums to Show"
   ClientHeight    =   3456
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   3180
   Icon            =   "frmDatumShow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3456
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstDatums 
      Height          =   2640
      Left            =   144
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   720
      Width           =   2844
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Check or uncheck datums to show when ploting datums"
      Height          =   492
      Left            =   192
      TabIndex        =   1
      Top             =   144
      Width           =   2796
   End
End
Attribute VB_Name = "frmDatumShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Loading As Boolean
Dim rstDatums As Recordset

Private Sub Form_Load()
Me.height = 4000
Me.Width = 3400



Dim Tablename As String
Dim FieldName As String
Tablename = "EDM_Datums"
FieldName = "Show"

Set rstDatums = MDImain.MainDB.OpenRecordset("EDM_Datums", dbOpenTable)
rstDatums.Index = "DatumName"
Loading = True
If Not FieldMatch(Tablename, "Show") Then
    Dim TblDef As TableDef
    Dim Fld As Field

    Set TblDef = MDImain.MainDB.TableDefs(Tablename)
    Set Fld = TblDef.CreateField(FieldName, dbBoolean)
    TblDef.Fields.Append Fld
    Set TblDef = Nothing
    Set Fld = Nothing
    sqlstring = "update [EDM_Datums] set show=true"
    MDImain.MainDB.Execute sqlstring
End If
lstDatums.Clear
i = -1
For i = 1 To nDatumPoints
    lstDatums.AddItem DatumName(i)
    If DatumShow(i) Then
        lstDatums.Selected(i - 1) = True
    Else
        lstDatums.Selected(i - 1) = False
    End If
Next i
Loading = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set rstDatums = Nothing

End Sub


Private Sub lstDatums_ItemCheck(Item As Integer)
Dim Value As Boolean

If Loading Then Exit Sub

If lstDatums.Selected(Item) Then
    Value = True
Else
    Value = False
End If

DatumShow(Item + 1) = Value

rstDatums.Seek "=", lstDatums.List(Item)
If Not rstDatums.NoMatch Then
    rstDatums.Edit
    rstDatums("show") = Value
    rstDatums.Update
End If

If PlotFormLoaded Then
    CurrentPF.Form_Paint
End If


End Sub


