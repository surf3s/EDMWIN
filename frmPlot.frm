VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPlot 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Plot"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "frmPlot.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2790
      Top             =   1140
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
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
   Begin VB.ComboBox txtGridSize 
      Height          =   315
      ItemData        =   "frmPlot.frx":000C
      Left            =   1140
      List            =   "frmPlot.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   60
      Width           =   1215
   End
   Begin VB.ComboBox txtOverlay 
      Height          =   315
      Left            =   48
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   900
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   228
      Left            =   24
      TabIndex        =   10
      Top             =   5784
      Width           =   4164
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   192
         Left            =   72
         TabIndex        =   16
         Top             =   24
         Width           =   108
      End
      Begin VB.Label txtX 
         AutoSize        =   -1  'True
         Height          =   192
         Left            =   252
         TabIndex        =   15
         Top             =   24
         Width           =   48
      End
      Begin VB.Label lblY 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         Height          =   192
         Left            =   1512
         TabIndex        =   14
         Top             =   24
         Width           =   108
      End
      Begin VB.Label txtY 
         AutoSize        =   -1  'True
         Height          =   192
         Left            =   1692
         TabIndex        =   13
         Top             =   24
         Width           =   48
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         Caption         =   "Unit:"
         Height          =   192
         Left            =   2772
         TabIndex        =   12
         Top             =   24
         Width           =   336
      End
      Begin VB.Label txtUnit 
         AutoSize        =   -1  'True
         Height          =   192
         Left            =   3156
         TabIndex        =   11
         Top             =   24
         Width           =   48
      End
   End
   Begin VB.CheckBox KeepScale 
      Caption         =   "Keep Scale"
      Height          =   225
      Left            =   60
      TabIndex        =   4
      Top             =   648
      Width           =   1188
   End
   Begin VB.OptionButton View 
      Caption         =   "Front"
      Height          =   195
      Index           =   2
      Left            =   1515
      TabIndex        =   2
      Top             =   1236
      Width           =   780
   End
   Begin VB.OptionButton View 
      Caption         =   "Side"
      Height          =   195
      Index           =   1
      Left            =   744
      TabIndex        =   1
      Top             =   1236
      Width           =   750
   End
   Begin VB.OptionButton View 
      Caption         =   "Plan"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   1236
      Value           =   -1  'True
      Width           =   648
   End
   Begin VB.CheckBox Grid 
      Caption         =   "Draw Grid"
      Height          =   225
      Left            =   60
      TabIndex        =   3
      Top             =   90
      Width           =   1095
   End
   Begin VB.CheckBox DatumNames 
      Caption         =   "Label Datums"
      Height          =   225
      Left            =   60
      TabIndex        =   5
      Top             =   372
      Width           =   1335
   End
   Begin VB.PictureBox PlotArea 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   2
      ForeColor       =   &H0000FFFF&
      Height          =   4000
      Left            =   30
      ScaleHeight     =   3945
      ScaleWidth      =   3945
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1728
      Width           =   4000
      Begin VB.Shape shpX 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape shpPoint 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   1650
         Shape           =   3  'Circle
         Top             =   660
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label DatumLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Label1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   990
         TabIndex        =   7
         Top             =   180
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape shpDatum 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   0
         Left            =   420
         Top             =   210
         Visible         =   0   'False
         Width           =   165
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Current Point"
      Height          =   195
      Index           =   2
      Left            =   2880
      TabIndex        =   21
      Top             =   780
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Datum"
      Height          =   195
      Index           =   1
      Left            =   2880
      TabIndex        =   20
      Top             =   450
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "X-shot location"
      Height          =   195
      Index           =   0
      Left            =   2880
      TabIndex        =   19
      Top             =   150
      Width           =   1050
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   2610
      Shape           =   3  'Circle
      Top             =   780
      Width           =   195
   End
   Begin VB.Shape shp3 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   2610
      Top             =   465
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   2610
      Shape           =   3  'Circle
      Top             =   150
      Width           =   195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Overlays"
      Height          =   192
      Left            =   1764
      TabIndex        =   18
      Top             =   948
      Width           =   648
   End
   Begin VB.Label lblScale 
      Caption         =   "Right-mouse click on plot  to return to original scale"
      Height          =   228
      Left            =   96
      TabIndex        =   9
      Top             =   1488
      Visible         =   0   'False
      Width           =   3828
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "m"
      Height          =   195
      Left            =   2400
      TabIndex        =   8
      Top             =   120
      Width           =   120
   End
End
Attribute VB_Name = "frmPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GridWidth As Single
Dim MaxGridWidth As Single
Dim MinX, MaxX, MinY, MaxY, MaxZ, MinZ As Single
Dim MovingLabel As Boolean
Dim CurrentLabel As Integer
Dim Xoffset, Yoffset As Single
Dim StartZoomX, StartZoomY, PrevZoomX, PrevZoomY As Single
Dim OriginalX, OriginalY As Single
Dim PlotMinX, PlotMaxX, PlotMinY, PlotMaxY As Single
Dim ndatums As Integer
Dim Zooming, StartedDrawing  As Boolean
Dim ZoomingOut As Boolean
Dim OPlotPoints(100000, 4) As Single
Dim NPlotPoints As Long

Public Sub PlotPoints()

Dim StartingX As Double
Dim StartingY As Double
data1.Recordset.Requery
On Error Resume Next
For I = 1 To ndatums
    Unload shpDatum(I)
    Unload DatumLabel(I)
Next I
On Error GoTo 0
PlotArea.Cls

If View(0) Then
    PlotX = "X"
    PlotY = "Y"
    PlotMinX = MinX - 0.1 * GridWidth
    PlotMaxX = MaxX + 0.1 * GridWidth
    PlotMinY = MinY - 0.1 * GridWidth
    PlotMaxY = MaxY + 0.1 * GridWidth
ElseIf View(1) Then
    PlotX = "Y"
    PlotY = "Z"
    PlotMinX = MinY - 0.1 * GridWidth
    PlotMaxX = MaxY + 0.1 * GridWidth
    PlotMinY = MinZ - 0.1 * GridWidth
    PlotMaxY = MaxZ + 0.1 * GridWidth
ElseIf View(2) Then
    PlotX = "X"
    PlotY = "Z"
    PlotMinX = MinX - 0.1 * GridWidth
    PlotMaxX = MaxX + 0.1 * GridWidth
    PlotMinY = MinZ - 0.1 * GridWidth
    PlotMaxY = MaxZ + 0.1 * GridWidth
End If
If PlotMaxX - PlotMinX < 0.001 Then
    PlotMinX = PlotMinX - 1
    PlotMaxX = PlotMaxX + 1
End If
If PlotMaxY - PlotMinY < 0.001 Then
    PlotMinY = PlotMinY - 1
    PlotMaxY = PlotMaxY + 1
End If

PlotArea.Scale (PlotMinX, PlotMaxY)-(PlotMaxX, PlotMinY)

lblX = PlotX
lblY = PlotY
If mdiMain.mnuViewPoints.Checked Then
    If CountRecords = 0 Then
        shpPoint.Visible = False
        
    Else
        data1.Recordset.MoveFirst
        While Not data1.Recordset.EOF
            If data1.Recordset("suffix") > 0 Then
                PlotArea.Line -(data1.Recordset(PlotX), data1.Recordset(PlotY)), QBColor(6)
            End If
            PlotArea.Circle (data1.Recordset(PlotX), data1.Recordset(PlotY)), 0.05, QBColor(6)
            data1.Recordset.MoveNext
        Wend
        data1.Recordset.MoveLast
        shpPoint.Visible = True
        shpPoint.Left = data1.Recordset(PlotX) - shpPoint.Width / 2
        shpPoint.Top = data1.Recordset(PlotY) + shpPoint.Height / 2
        Me.Caption = data1.Recordset("Unit") & "-" & data1.Recordset("ID") & "(" & data1.Recordset("Suffix") & ")"
    End If
Else
    shpPoint.Visible = False
End If

If mdiMain.mnuViewDatums.Checked Then
    If DatumTB.RecordCount > 0 Then
        DatumTB.MoveFirst
        While Not DatumTB.EOF
            ndatums = ndatums + 1
            Load shpDatum(ndatums)
            shpDatum(ndatums).Visible = True
            shpDatum(ndatums).Top = DatumTB(PlotY)
            shpDatum(ndatums).Left = DatumTB(PlotX)
            If DatumNames = 1 Then
                Load DatumLabel(ndatums)
                DatumLabel(ndatums) = LCase(DatumTB("Name"))
                DatumLabel(ndatums).Top = shpDatum(ndatums).Top
                DatumLabel(ndatums).Left = shpDatum(ndatums).Left + shpDatum(ndatums).Width
                If DatumLabel(ndatums).Left + DatumLabel(ndatums).Width > PlotMaxX Then
                    DatumLabel(ndatums).Left = shpDatum(ndatums).Left - DatumLabel(ndatums).Width
                End If
                DatumLabel(ndatums).Visible = True
            End If
            DatumTB.MoveNext
        Wend
    End If
End If
If mdiMain.mnuViewUnits.Checked And View(0) Then
    If UnitTB.RecordCount > 0 Then
        UnitTB.MoveFirst
        While Not UnitTB.EOF
            If UnitTB("minx") <> -99999 Then
                    PlotArea.Line (UnitTB("minx"), UnitTB("miny"))-(UnitTB("maxx"), UnitTB("maxy")), QBColor(10), B
            ElseIf Not IsNull(UnitTB("RADIUS")) Then
                    PlotArea.Circle (UnitTB("CENTERX"), UnitTB("CENTERY")), UnitTB("RADIUS"), QBColor(10)
            End If
            UnitTB.MoveNext
        Wend
    End If
End If

If Grid Then
    GridSize = Val(txtGridSize)
    
    PlotArea.DrawStyle = vbDash
    Select Case GridSize
        Case 0.5
            StartingX = Int(PlotMinX) - 0.5
            StartinyX = Int(PlotMinY) - 0.5
        Case 1
            StartingX = Int(PlotMinX)
            StartingY = Int(PlotMinY)
        
        Case 5
            StartingX = Int(PlotMinX)
            StartingY = Int(PlotMinY)
            Do Until StartingX Mod 5 = 0
                StartingX = StartingX + 1
            Loop
            Do Until StartingY Mod 5 = 0
                StartingY = StartingY + 1
            Loop
        Case 10
            StartingX = Int(PlotMinX)
            StartingY = Int(PlotMinY)
            Do Until StartingX Mod 10 = 0
                StartingX = StartingX + 1
            Loop
            Do Until StartingY Mod 10 = 0
                StartingY = StartingY + 1
            Loop
            
        Case 100
            StartingX = Int(PlotMinX)
            StartingY = Int(PlotMinY)
            Do Until StartingX Mod 100 = 0
                StartingX = StartingX + 1
            Loop
            Do Until StartingY Mod 100 = 0
                StartingY = StartingY + 1
            Loop
    End Select
    
    For I = Int(StartingX) To Int(PlotMaxX) Step GridSize
        PlotArea.Line (I, PlotMinY)-(I, PlotMaxY), QBColor(8)
    Next I
    For I = Int(StartingY) To Int(PlotMaxY) Step GridSize
        PlotArea.Line (PlotMinX, I)-(PlotMaxX, I), QBColor(8)
    Next I
End If
PlotArea.DrawStyle = 0
If shpX.Visible Then
    If View(0) Then
        shpX.Left = edmshot.X
        shpX.Top = edmshot.y
    ElseIf View(1) Then
        shpX.Left = edmshot.y
        shpX.Top = edmshot.z
    ElseIf View(2) Then
        shpX.Left = edmshot.X
        shpX.Top = edmshot.z
    End If
End If
PlotOverlay

End Sub


Private Sub DatumLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

CurrentLabel = Index
MovingLabel = True
DatumLabel(Index).Visible = False

Xoffset = X * PlotArea.ScaleWidth / PlotArea.Width
Yoffset = y * PlotArea.ScaleWidth / PlotArea.Width
DatumLabel(Index).Drag

End Sub

Private Sub DatumLabel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)

MovingLabel = False

End Sub

Private Sub DatumNames_Click()

If DatumNames = 1 Then
    If mdiMain.mnuViewDatums.Checked = False Then
        mdiMain.mnuViewDatums_Click
    Else
        PlotPoints
    End If
Else
    PlotPoints
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd
        frmMain.Form_KeyDown KeyCode, 0
        frmMain.Picture1.SetFocus
        Exit Sub
End Select
If Shift = 2 Or Shift = 4 Then
    frmMain.SetFocus
    frmMain.Form_KeyDown KeyCode, Shift
End If

End Sub

Private Sub Form_Load()

Me.Width = mdiMain.Width - frmMain.Width - 200
Me.Left = frmMain.Width + 10
Me.Top = 0
PlotArea.Width = Me.Width - 20
PlotArea.Height = PlotArea.Width
Frame1.Top = PlotArea.Top + PlotArea.Height
Me.Height = Frame1.Top + Frame1.Height + 500
data1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + SiteDBname + ";Persist Security Info=False"
Set data1.Recordset = frmMain.PointsADO.Recordset.Clone

PlotShowing = True
txtGridSize = txtGridSize.List(1)

End Sub

Private Sub Form_Unload(Cancel As Integer)

mdiMain.mnuViewDatums.Checked = False
mdiMain.mnuViewPoints.Checked = False
mdiMain.mnuViewUnits.Checked = False
mdiMain.mnuViewAll.Caption = "&All"
PlotShowing = False

End Sub

Private Sub Grid_Click()

PlotPoints

End Sub

Private Sub PlotArea_DragDrop(Source As Control, X As Single, y As Single)

If MovingLabel Then
    Source.Top = y + Yoffset
    Source.Left = X - Xoffset
    Source.Visible = True
    MovingLabel = False
End If

End Sub

Private Sub PlotArea_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

Dim ClosestOne, Distance As Single
Dim CurrentUnit, CurrentID As String
Dim CurrentSuffix As Integer

If Button = 2 Then
    ZoomingOut = True
    lblScale.Visible = False
    SetScale
    PlotPoints
    Exit Sub
End If

StartZoom:

StartZoomX = X: StartZoomY = y
PrevZoomX = X: PrevZoomY = y
Zooming = True

End Sub

Private Sub PlotArea_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

txtX = Format(X, "########0.000")
txtY = Format(y, "########0.000")
SqlString = "select unit from [EDM_units] where minx< " & X & " and maxx>" & X & " and miny<" & y & " and maxy>" & y
Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
If Not rsTemp.EOF Then
    If Not IsNull(rsTemp("unit")) Then
        txtUnit = rsTemp("unit")
    Else
        txtUnit = ""
    End If
Else
    SqlString = "select UNIT from [EDM_units] where abs(centerx-" & X & ")<=radius and abs(centery-" & y & ")<=radius"
    Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("unit")) Then
            txtUnit = rsTemp("unit")
        Else
            txtUnit = ""
        End If
    Else
        txtUnit = ""
    End If
End If

If Not StartedDrawing And Abs(PrevZoomX - X) < 0.03 * PlotArea.ScaleWidth And Abs(PrevZoomY - y) < 0.03 * Abs(PlotArea.ScaleHeight) Then
    Exit Sub
End If

If (Button And vbLeftButton) > 0 And Zooming Then
    PlotArea.DrawWidth = 1
    PlotArea.DrawMode = 6
    PlotArea.DrawStyle = 2
    StartedDrawing = True
    PlotArea.Line (StartZoomX, StartZoomY)-(PrevZoomX, PrevZoomY), , B
    PlotArea.Line (StartZoomX, StartZoomY)-(X, y), , B
    PrevZoomX = X: PrevZoomY = y
    PlotArea.DrawStyle = 0
    PlotArea.DrawWidth = 1
    PlotArea.DrawMode = 13
End If

Set rsTemp = Nothing

End Sub

Private Sub PlotArea_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

Dim TempMinX, TempMinY, TempMaxX, TempMaxY As Single
Dim RangeX, RangeY As Single

If ZoomingOut Then
    ZoomingOut = False
    Exit Sub
End If
If Not StartedDrawing Then
    SqlString = "select unit, id, suffix, " + PlotX + ", " + PlotY + " from [" + PointTableName + "] where " + PlotX + " < " + PlotX + "+.1 and " + PlotX + " > " + PlotX + "-.1 and " + PlotY + " < " + PlotY + "+.1 and " + PlotY + " > " + PlotY + "-.1"
    Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenDynaset)
    ClosestOne = 32000
    
    If Not rsTemp.EOF Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            Distance = Sqr((rsTemp(PlotX) - X) ^ 2 + (rsTemp(PlotY) - y) ^ 2)
            If Distance < ClosestOne Then
                CurrentUnit = rsTemp("unit")
                CurrentID = rsTemp("id")
                CurrentSuffix = rsTemp("suffix")
                Gotit = True
                ClosestOne = Distance
            End If
            rsTemp.MoveNext
        Loop
        If Not Gotit Then Exit Sub
        data1.Recordset.MoveFirst
        Do While Not data1.Recordset.EOF
            If data1.Recordset("unit") = CurrentUnit And data1.Recordset("id") = CurrentID And data1.Recordset("suffix") = CurrentSuffix Then
                Me.Caption = CurrentUnit & "-" & CurrentID & "(" & CurrentSuffix & ")"
                frmMain.PointsADO.Recordset.Bookmark = data1.Recordset.Bookmark
                frmMain.ShowValues
                shpPoint.Visible = True
                shpPoint.Left = data1.Recordset(PlotX) - shpPoint.Width / 2
                shpPoint.Top = data1.Recordset(PlotY) + shpPoint.Height / 2
                Exit Do
            End If
            data1.Recordset.MoveNext
        Loop
    End If
    PlotArea.SetFocus
    Exit Sub
End If

PlotArea.DrawWidth = 1
PlotArea.DrawMode = 6
PlotArea.DrawStyle = 2

PlotArea.Line (StartZoomX, StartZoomY)-(PrevZoomX, PrevZoomY), , B
StartedDrawing = False

PlotArea.DrawStyle = 0
PlotArea.DrawWidth = 1
PlotArea.DrawMode = 13

If Abs(StartZoomX - X) < 0.1 And Abs(StartZoomY - y) < 0.1 Then Exit Sub

If StartZoomX < X Then
    TempMaxX = X
    TempMinX = StartZoomX
Else
    TempMaxX = StartZoomX
    TempMinX = X
End If
If StartZoomY < y Then
    TempMaxY = y
    TempMinY = StartZoomY
Else
    TempMaxY = StartZoomY
    TempMinY = y
End If
RangeX = TempMaxX - TempMinX
RangeY = TempMaxY - TempMinY

If View(0) Then
    MaxX = TempMaxX
    MinX = TempMinX
    MaxY = TempMaxY
    MinY = TempMinY
ElseIf View(1) Then
    MaxY = TempMaxX
    MinY = TempMinX
    MaxZ = TempMaxY
    MinZ = TempMinY
ElseIf View(2) Then
    MaxX = TempMaxX
    MinX = TempMinX
    MaxZ = TempMaxY
    MinZ = TempMinY
End If

KeepScale.Visible = True

If View(0) Then
    MaxX = TempMaxX
    MinX = TempMinX
    MaxY = TempMaxY
    MinY = TempMinY
ElseIf View(1) Then
    MaxY = TempMaxX
    MinY = TempMinX
    MaxZ = TempMaxY
    MinZ = TempMinY
ElseIf View(2) Then
    MaxX = TempMaxX
    MinX = TempMinX
    MaxZ = TempMaxY
    MinZ = TempMinY
End If

If RangeX > RangeY Then
    GridWidth = RangeX
    Offset = (RangeX - RangeY) / 2
    If View(0) Then
        MinY = MinY - Offset
        MaxY = MaxY + Offset
    Else
        MinZ = MinZ - Offset
        MaxZ = MaxZ + Offset
    End If
Else
    GridWidth = RangeY
    Offset = (RangeY - RangeX) / 2
    If View(0) Then
        MinX = MinX - Offset
        MaxX = MaxX + Offset
    Else
        MinY = MinY - Offset
        MaxY = MaxY + Offset
    End If
End If

lblScale.Visible = True
PlotPoints
Set rsTemp = Nothing

End Sub

Private Sub txtGridSize_Click()

PlotPoints

End Sub

Private Sub txtOverlay_Click()

If txtOverlay <> "None" Then
    NPlotPoints = 0
    SqlString = "Select x, y, z, suffix from [" + txtOverlay + "]"
    Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenForwardOnly)
    While Not rsTemp.EOF
        NPlotPoints = NPlotPoints + 1
        OPlotPoints(NPlotPoints, 1) = rsTemp("X")
        OPlotPoints(NPlotPoints, 2) = rsTemp("y")
        OPlotPoints(NPlotPoints, 3) = rsTemp("z")
        OPlotPoints(NPlotPoints, 4) = rsTemp("suffix")
        rsTemp.MoveNext
    Wend
    Set rsTemp = Nothing
    PlotPoints
Else
    NPlotPoints = 0
    PlotPoints
End If

End Sub

Private Sub txtOverlay_DropDown()

Dim Parity As Byte
Dim A As Boolean, B As Boolean, C As Boolean, d As Boolean

txtOverlay.Clear
txtOverlay.AddItem "None"

SiteDB.TableDefs.Refresh
For Each tdtemp In SiteDB.TableDefs
    A = False
    If Left(tdtemp.Name, 7) <> "WinPlot" And Left(tdtemp.Name, 4) <> "MSys" And Left(tdtemp.Name, 2) <> "$$" And Left(tdtemp.Name, 4) <> "EDM_" Then
        On Error Resume Next
        A = SiteDB.TableDefs(tdtemp.Name).Fields("X").Size > 0
        B = SiteDB.TableDefs(tdtemp.Name).Fields("y").Size > 0
        C = SiteDB.TableDefs(tdtemp.Name).Fields("z").Size > 0
        d = SiteDB.TableDefs(tdtemp.Name).Fields("suffix").Size > 0
        If A And B And C And d Then
            txtOverlay.AddItem LCase(tdtemp.Name)
        End If
    End If
Next
On Error GoTo 0

End Sub

Private Sub View_Click(Index As Integer)

'KeepScale = 0
'KeepScale.Visible = False
'lblScale.Visible = False
SetScale
PlotPoints

End Sub

Public Function MaximumNum(Number1, Number2, Number3)

Dim X As Double
X = Number1
MaximumNum = 1
If Number2 > X Then
    X = Number2
    MaximumNum = 2
End If
If Number3 > X Then
    MaximumNum = 3
End If

End Function

Public Sub SetScale()

Dim RangeX, RangeY, RangeZ As Single
Dim rsTemp As Recordset
Dim Offset As Single


'If ZoomingOut Then
'    RangeX = (MaxX - MinX) * 2
'    RangeY = (MaxY - MinY) * 2
'    RangeZ = (MaxZ - MinZ) * 2
'    GoTo Continue
'End If

If KeepScale = 1 Then Exit Sub
MinX = 320000
MinY = 320000
MaxX = -320000
MaxY = -320000
MinZ = 320000
MaxZ = -320000

If mdiMain.mnuViewPoints.Checked Then
    GoSub GetPointsRange
End If
If mdiMain.mnuViewDatums.Checked Then
    GoSub GetDatumsRange
End If
If mdiMain.mnuViewUnits.Checked And View(0) Then
    GoSub GetUnitsRange
End If
If shpX.Visible Then
    If edmshot.X > MaxX Then MaxX = edmshot.X
    If edmshot.X < MinX Then MinX = edmshot.X
    If edmshot.y > MaxY Then MaxY = edmshot.y
    If edmshot.y < MinY Then MinY = edmshot.y
    If edmshot.z > MaxZ Then MaxZ = edmshot.z
    If edmshot.z < MinZ Then MinZ = edmshot.z
End If
RangeX = MaxX - MinX
RangeY = MaxY - MinY
RangeZ = MaxZ - MinZ

Continue:
Select Case MaximumNum(RangeX, RangeY, RangeZ)
    Case 1
        GridWidth = RangeX
        Offset = (RangeX - RangeY) / 2
        MinY = MinY - Offset
        MaxY = MaxY + Offset
        Offset = (RangeX - RangeZ) / 2
        MinZ = MinZ - Offset
        MaxZ = MaxZ + Offset
    Case 2
        GridWidth = RangeY
        Offset = (RangeY - RangeX) / 2
        MinX = MinX - Offset
        MaxX = MaxX + Offset
        Offset = (RangeY - RangeZ) / 2
        MinZ = MinZ - Offset
        MaxZ = MaxZ + Offset
    Case 3
        GridWidth = RangeZ
        Offset = (RangeZ - RangeX) / 2
        MinX = MinX - Offset
        MaxX = MaxX + Offset
        Offset = (RangeZ - RangeY) / 2
        MinY = MinY - Offset
        MaxY = MaxY + Offset

End Select

'KeepScale = 0
'KeepScale.Visible = False

Exit Sub

GetPointsRange:
    SqlString = "Select max(x) as Maxx, min(x) as minx, max(y) as maxy, min(Y) as miny, max(z) as maxz, min(z) as minz from [" + PointTableName + "]"
    Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenDynaset)
    If rsTemp.EOF Then Return
    rsTemp.MoveFirst
    If rsTemp("maxx") > MaxX Then MaxX = rsTemp("maxx")
    If rsTemp("minx") < MinX Then MinX = rsTemp("minx")
    If rsTemp("maxy") > MaxY Then MaxY = rsTemp("maxy")
    If rsTemp("miny") < MinY Then MinY = rsTemp("miny")
    If rsTemp("maxz") > MaxZ Then MaxZ = rsTemp("maxz")
    If rsTemp("minz") < MinZ Then MinZ = rsTemp("minz")
    Set rsTemp = Nothing
Return

GetDatumsRange:
    SqlString = "Select max(x) as Maxx, min(x) as minx, max(y) as maxy, min(Y) as miny, max(z) as maxz, min(z) as minz from [EDM_datums]"
    Set rsTemp = SiteDB.OpenRecordset(SqlString, dbOpenDynaset)
    If rsTemp.EOF Then Return
    rsTemp.MoveFirst
    If rsTemp("maxx") > MaxX Then MaxX = rsTemp("maxx")
    If rsTemp("minx") < MinX Then MinX = rsTemp("minx")
    If rsTemp("maxy") > MaxY Then MaxY = rsTemp("maxy")
    If rsTemp("miny") < MinY Then MinY = rsTemp("miny")
    If rsTemp("maxz") > MaxZ Then MaxZ = rsTemp("maxz")
    If rsTemp("minz") < MinZ Then MinZ = rsTemp("minz")
    Set rsTemp = Nothing
Return

GetUnitsRange:
    If Not (UnitTB.BOF And UnitTB.EOF) Then
        UnitTB.MoveFirst
        While Not UnitTB.EOF
            If UnitTB("MaxX") = -99999 Then
            Else
                If UnitTB("MaxX") > MaxX Then MaxX = UnitTB("MaxX")
                If UnitTB("MINX") < MinX Then MinX = UnitTB("MINX")
                If UnitTB("MaxY") > MaxY Then MaxY = UnitTB("MaxY")
                If UnitTB("MInY") < MinY Then MinY = UnitTB("MINY")
                If UnitTB("CENTERX") + UnitTB("RADIUS") > MaxX Then MaxX = UnitTB("CENTERx") + UnitTB("RADIUS")
                If UnitTB("CENTERX") - UnitTB("RADIUS") < MinX Then MinX = UnitTB("CENTERx") - UnitTB("RADIUS")
                If UnitTB("CENTERY") + UnitTB("RADIUS") > MaxY Then MaxY = UnitTB("CENTERY") + UnitTB("RADIUS")
                If UnitTB("CENTERY") - UnitTB("RADIUS") < MinY Then MinY = UnitTB("CENTERY") - UnitTB("RADIUS")
            End If
            UnitTB.MoveNext
        Wend
    End If
Return

End Sub

Public Sub PlotOverlay()

Dim PlotX, PlotY As Integer
If View(0) Then
    PlotX = 1
    PlotY = 2
ElseIf View(1) Then
    PlotX = 2
    PlotY = 3
ElseIf View(2) Then
    PlotX = 1
    PlotY = 3
End If

For I = 1 To NPlotPoints
    If OPlotPoints(I, 4) > 0 Then
        PlotArea.Line -(OPlotPoints(I, PlotX), OPlotPoints(I, PlotY)), QBColor(10)
    End If
    PlotArea.Circle (OPlotPoints(I, PlotX), OPlotPoints(I, PlotY)), 0.05, QBColor(10)
Next I

End Sub
