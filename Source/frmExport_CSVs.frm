VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmExport_CSVs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export CSVs"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton change_file_path 
      Caption         =   "Change"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3120
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox filepath_txt 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   5175
   End
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton export 
      Caption         =   "Export"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   $"frmExport_CSVs.frx":0000
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "frmExport_CSVs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cancel_Click()

Unload Me

End Sub

Private Sub change_file_path_Click()

cd.ShowSave
If Not cd.CancelError Then
    A = InStr(cd.filename, ".")
    If A <> 0 Then
        filepath_txt.Text = Left(cd.filename, A - 1)
    Else
        filepath_txt.Text = cd.filename
    End If
End If

End Sub

Private Function check_for_null(item) As String

    If IsNull(item) Then
        check_for_null = "0"
    Else
        check_for_null = item
    End If
    
End Function

Private Sub export_Click()

If filepath_txt.Text <> "" Then

    Screen.MousePointer = 11
    
    filename = filepath_txt.Text + "_datums.csv"
    F = FreeFile
    datums = 0
    Open filename For Output As #F
        Write #F, "Name", "X", "Y", "Z"
        DatumTB.MoveFirst
        While Not DatumTB.EOF
            Write #F, DatumTB("Name"), CDbl(Format(check_for_null(DatumTB("X")), "#####0.000")), CDbl(Format(DatumTB("Y"), "#####0.000")), CDbl(Format(DatumTB("Z"), "#####0.000"))
            DatumTB.MoveNext
            datums = datums + 1
        Wend
    Close F
    
    filename = filepath_txt.Text + "_prisms.csv"
    F = FreeFile
    prisms = 0
    Open filename For Output As #F
        Write #F, "Name", "Height", "Offset"
        PoleTB.MoveFirst
        While Not PoleTB.EOF
            Write #F, PoleTB("Name"), CDbl(Format(check_for_null(PoleTB("height")), "#####0.000")), CDbl(Format(check_for_null(PoleTB("offset")), "#####0.000"))
            PoleTB.MoveNext
            prisms = prisms + 1
        Wend
    Close F

    filename = filepath_txt.Text + "_units.csv"
    F = FreeFile
    units = 0
    Open filename For Output As #F
        Write #F, "Unit", "Minx", "Maxx", "Miny", "Maxy", "CenterX", "CenterY", "Radius"
        UnitTB.MoveFirst
        While Not UnitTB.EOF
            Write #F, UnitTB("Unit"), CDbl(Format(check_for_null(UnitTB("minx")), "#####0.000")), CDbl(Format(check_for_null(UnitTB("maxx")), "#####0.000")), CDbl(Format(check_for_null(UnitTB("miny")), "#####0.000")), CDbl(Format(check_for_null(UnitTB("maxy")), "#####0.000")), CDbl(Format(check_for_null(UnitTB("centerx")), "#####0.000")), CDbl(Format(check_for_null(UnitTB("centery")), "#####0.000")), CDbl(Format(check_for_null(UnitTB("radius")), "#####0.000"))
            UnitTB.MoveNext
            units = units + 1
        Wend
    Close F

    filename = filepath_txt.Text + "_points.csv"
    F = FreeFile
    points = 0
    Open filename For Output As #F
        lineout$ = ""
        For Each cfield In frmMain.PointsADO.Recordset.Fields
            If lineout$ <> "" Then
                lineout$ = lineout$ + ","
            End If
            lineout$ = lineout$ + Chr$(34) + UCase(cfield.Name) + Chr$(34)
        Next
        Print #F, lineout$
        frmMain.PointsADO.Recordset.MoveFirst
        While Not frmMain.PointsADO.Recordset.EOF
            lineout$ = ""
            For Each cfield In frmMain.PointsADO.Recordset.Fields
                If lineout$ <> "" Then
                    lineout$ = lineout$ + ","
                End If
                Select Case cfield.Type
                Case 202, 129, 133, 134, 135, 201, 203, 200
                    If Not IsNull(frmMain.PointsADO.Recordset.Fields(cfield.Name)) Then
                        lineout$ = lineout$ + Chr$(34) + frmMain.PointsADO.Recordset.Fields(cfield.Name) + Chr$(34)
                    Else
                        lineout$ = lineout$ + Chr$(34) + Chr$(34)
                    End If
                Case 7
                    If Not IsNull(frmMain.PointsADO.Recordset.Fields(cfield.Name)) Then
                        lineout$ = lineout$ + Chr$(34) + Str(frmMain.PointsADO.Recordset.Fields(cfield.Name)) + Chr$(34)
                    Else
                        lineout$ = lineout$ + Chr$(34) + Chr$(34)
                    End If
                Case 129, 14, 5, 3, 131, 4, 2, 16, 19, 18, 139
                    If Not IsNull(frmMain.PointsADO.Recordset.Fields(cfield.Name)) Then
                        lineout$ = lineout$ + Str(frmMain.PointsADO.Recordset.Fields(cfield.Name))
                    End If
                Case Else
                    Stop
                End Select
            Next
            Print #F, lineout$
            frmMain.PointsADO.Recordset.MoveNext
            points = points + 1
        Wend
    Close F

    Screen.MousePointer = 1
    
    message = Str(datums) + " datums," + Str(prisms) + " prisms," + Str(units) + " units, and" + Str(points) + " points were exported."
    answer = MsgBox(message, vbOKOnly)
    
    Unload Me
    
End If

End Sub
