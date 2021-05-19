Attribute VB_Name = "Module1"
Global CallingForm As Form
Global Const XY = 1
Global Const YZ = 2
Global Const XZ = 3
Global DatumFile As String
Global PointFile As String
Global PointData(3, 3) As Single
Global TitleString As String




Public Sub GetOffset()
Dim OffsetX As Single
Dim OffsetY As Single
Dim OffsetZ As Single

OffsetX = Val(CallingForm.Point1(0)) - Val(CallingForm.Datum1(0))
OffsetY = Val(CallingForm.Point1(1)) - Val(CallingForm.Datum1(1))
OffsetZ = Val(CallingForm.Point1(2)) - Val(CallingForm.Datum1(2))

CallingForm.Offset1(0) = Format(Val(CallingForm.Point1(0)) - OffsetX, "0.0000")
CallingForm.Offset1(1) = Format(Val(CallingForm.Point1(1)) - OffsetY, "0.0000")
CallingForm.Offset1(2) = Format(Val(CallingForm.Point1(2)) - OffsetZ, "0.0000")
CallingForm.Offset2(0) = Format(Val(CallingForm.Point2(0)) - OffsetX, "0.0000")
CallingForm.Offset2(1) = Format(Val(CallingForm.Point2(1)) - OffsetY, "0.0000")
CallingForm.Offset2(2) = Format(Val(CallingForm.Point2(2)) - OffsetZ, "0.0000")
CallingForm.Offset3(0) = Format(Val(CallingForm.Point3(0)) - OffsetX, "0.0000")
CallingForm.Offset3(1) = Format(Val(CallingForm.Point3(1)) - OffsetY, "0.0000")
CallingForm.Offset3(2) = Format(Val(CallingForm.Point3(2)) - OffsetZ, "0.0000")
End Sub



Public Sub GetAngle(Action As Integer)
Dim Y1 As Single
Dim Y2 As Single
Dim X1 As Single
Dim X2 As Single
Dim HLineAngle As Single
Dim LineAngle As Single
Dim DatumAngle As Single
Dim Direction As Integer
Dim TooSmall As Boolean
Dim TotalAngle As Single
Dim nTotalAngles As Integer
Dim Xpoint As Single
Dim Ypoint As Single


CallingForm.Angle1(0) = ""
CallingForm.Angle1(1) = ""
CallingForm.Angle1(2) = ""
Select Case Action
    Case XY
        Xpoint = 0
        Ypoint = 1
    Case XZ
        Xpoint = 0
        Ypoint = 2
    Case YZ
        Xpoint = 1
        Ypoint = 2
End Select




    'line 1-2
    Y1 = Val(CallingForm.Datum1(Ypoint))
    Y2 = Val(CallingForm.Datum2(Ypoint))
    X1 = Val(CallingForm.Datum1(Xpoint))
    X2 = Val(CallingForm.Datum2(Xpoint))
    GoSub ConvertAngleFromNorth
    CallingForm.AngleDatum(0) = HLineAngle
    DatumAngle = HLineAngle
    
    TooSmall = False
    Y1 = Val(CallingForm.Point1(Ypoint))
    Y2 = Val(CallingForm.Point2(Ypoint))
    X1 = Val(CallingForm.Point1(Xpoint))
    X2 = Val(CallingForm.Point2(Xpoint))
    GoSub ConvertAngleFromNorth
    If Not TooSmall Then
        CallingForm.AngleLine(0) = HLineAngle
        LineAngle = HLineAngle
        If LineAngle > DatumAngle Then
            CallingForm.Angle1(0) = LineAngle - DatumAngle
            If CallingForm.Angle1(0) < 1 Then CallingForm.Angle1(0) = ""
        Else
            CallingForm.Angle1(0) = (360 - DatumAngle) + LineAngle
            If CallingForm.Angle1(0) > 359 Then CallingForm.Angle1(0) = ""
        End If
    End If
    'Line 1-3
    Y1 = Val(CallingForm.Datum1(Ypoint))
    Y2 = Val(CallingForm.Datum3(Ypoint))
    X1 = Val(CallingForm.Datum1(Xpoint))
    X2 = Val(CallingForm.Datum3(Xpoint))
    GoSub ConvertAngleFromNorth
    DatumAngle = HLineAngle
    CallingForm.AngleDatum(1) = HLineAngle
    
    TooSmall = False
    Y1 = Val(CallingForm.Point1(Ypoint))
    Y2 = Val(CallingForm.Point3(Ypoint))
    X1 = Val(CallingForm.Point1(Xpoint))
    X2 = Val(CallingForm.Point3(Xpoint))
    GoSub ConvertAngleFromNorth
    If Not TooSmall Then
        CallingForm.AngleLine(1) = HLineAngle
        LineAngle = HLineAngle
        If LineAngle > DatumAngle Then
            CallingForm.Angle1(1) = LineAngle - DatumAngle
            If CallingForm.Angle1(1) < 1 Then CallingForm.Angle1(1) = ""
        Else
            CallingForm.Angle1(1) = (360 - DatumAngle) + LineAngle
            If CallingForm.Angle1(1) > 359 Then CallingForm.Angle1(1) = ""
        End If
    End If

    'Line 2-3
    Y1 = Val(CallingForm.Datum2(Ypoint))
    Y2 = Val(CallingForm.Datum3(Ypoint))
    X1 = Val(CallingForm.Datum2(Xpoint))
    X2 = Val(CallingForm.Datum3(Xpoint))
    GoSub ConvertAngleFromNorth
    DatumAngle = HLineAngle
    CallingForm.AngleDatum(2) = HLineAngle
    
    TooSmall = False
    Y1 = Val(CallingForm.Point2(Ypoint))
    Y2 = Val(CallingForm.Point3(Ypoint))
    X1 = Val(CallingForm.Point2(Xpoint))
    X2 = Val(CallingForm.Point3(Xpoint))
    GoSub ConvertAngleFromNorth
    If Not TooSmall Then
        CallingForm.AngleLine(2) = HLineAngle
        LineAngle = HLineAngle
        If LineAngle > DatumAngle Then
            CallingForm.Angle1(2) = LineAngle - DatumAngle
            If CallingForm.Angle1(2) < 1 Then CallingForm.Angle1(2) = ""
        Else
            CallingForm.Angle1(2) = (360 - DatumAngle) + LineAngle
            If CallingForm.Angle1(2) > 359 Then CallingForm.Angle1(2) = ""
        End If
    End If
nTotalAngles = 0
TotalAngle = 0
For i = 0 To 2
    If CallingForm.Angle1(i) <> "" Then
        nTotalAngles = nTotalAngles + 1
        TotalAngle = TotalAngle + Val(CallingForm.Angle1(i))
    End If
Next i

If nTotalAngles > 0 Then
    CallingForm.AverageAngle = Format(TotalAngle / nTotalAngles, "0.0000")
Else
    CallingForm.AverageAngle = Format(0, "0.0000")
End If
Exit Sub

ConvertAngleFromNorth:
    If Abs((Y2 - Y1)) < 0.01 Then
        TooSmall = True
        'Return
    End If
    slopetan = (Y2 - Y1) / (X2 - X1)
    HLineAngle = Atn(slopetan) * 57.2958
    
    If X2 > X1 Then
        HLineAngle = 90 - HLineAngle
    Else
        HLineAngle = 270 - HLineAngle
    End If
Return

End Sub

Public Sub RotatePoints(Action)
Dim Xpoint As Integer
Dim Ypoint As Integer


Select Case Action
    Case XY
        Xpoint = 0
        Ypoint = 1
    Case XZ
        Xpoint = 0
        Ypoint = 2
    Case YZ
        Xpoint = 1
        Ypoint = 2
End Select

Xorigin = CallingForm.Offset1(Xpoint)
Yorigin = CallingForm.Offset1(Ypoint)

Sinangle = Sin(-CallingForm.AverageAngle * 1.74532925199433E-02)
Cosangle = Cos(-CallingForm.AverageAngle * 1.74532925199433E-02)

X = Val(CallingForm.Offset1(Xpoint)) - Xorigin
Y = Val(CallingForm.Offset1(Ypoint)) - Yorigin
Xtemp = X * Cosangle + Y * Sinangle
Ytemp = Y * Cosangle - X * Sinangle
CallingForm.Rotate1(Xpoint) = Format(Xtemp + Xorigin, "0.0000")
CallingForm.Rotate1(Ypoint) = Format(Ytemp + Yorigin, "0.0000")

X = Val(CallingForm.Offset2(Xpoint)) - Xorigin
Y = Val(CallingForm.Offset2(Ypoint)) - Yorigin
Xtemp = X * Cosangle + Y * Sinangle
Ytemp = Y * Cosangle - X * Sinangle
CallingForm.Rotate2(Xpoint) = Format(Xtemp + Xorigin, "0.0000")
CallingForm.Rotate2(Ypoint) = Format(Ytemp + Yorigin, "0.0000")

X = Val(CallingForm.Offset3(Xpoint)) - Xorigin
Y = Val(CallingForm.Offset3(Ypoint)) - Yorigin
Xtemp = X * Cosangle + Y * Sinangle
Ytemp = Y * Cosangle - X * Sinangle
CallingForm.Rotate3(Xpoint) = Format(Xtemp + Xorigin, "0.0000")
CallingForm.Rotate3(Ypoint) = Format(Ytemp + Yorigin, "0.0000")

CallingForm.Diff1(Xpoint) = Format(Val(CallingForm.Rotate1(Xpoint)) - Val(CallingForm.Datum1(Xpoint)), "0.0000")
CallingForm.Diff1(Ypoint) = Format(Val(CallingForm.Rotate1(Ypoint)) - Val(CallingForm.Datum1(Ypoint)), "0.0000")
CallingForm.Diff2(Xpoint) = Format(Val(CallingForm.Rotate2(Xpoint)) - Val(CallingForm.Datum2(Xpoint)), "0.0000")
CallingForm.Diff2(Ypoint) = Format(Val(CallingForm.Rotate2(Ypoint)) - Val(CallingForm.Datum2(Ypoint)), "0.0000")
CallingForm.Diff3(Xpoint) = Format(Val(CallingForm.Rotate3(Xpoint)) - Val(CallingForm.Datum3(Xpoint)), "0.0000")
CallingForm.Diff3(Ypoint) = Format(Val(CallingForm.Rotate3(Ypoint)) - Val(CallingForm.Datum3(Ypoint)), "0.0000")

End Sub

Public Sub GetData(Filename As String)
Open "C:\Dropbox\VB\EDM ADO Windows\" & Filename For Input As 1
For i = 1 To 3
    Line Input #1, cbdata
    If Len(cbdata) = 0 Then
        MsgBox ("Copy data from notepad")
        Exit Sub
    End If
        For J = 1 To Len(cbdata)
            Select Case Asc(Mid(cbdata, J, 1))
                Case 48 To 57, Asc("-"), Asc("."), Asc(","), 13, 10
                
                Case Else
                    MsgBox ("Invalid microscribe data")
                    J = Len(cbdata) + 1
                    i = 4
            End Select
        Next J
    
    Xtemp = InStr(cbdata, ",")
    PointData(i, 0) = Left(cbdata, Xtemp - 1)
    cbdata = Mid(cbdata, Xtemp + 1)
    Xtemp = InStr(cbdata, ",")
    PointData(i, 1) = Left(cbdata, Xtemp - 1)
    cbdata = Mid(cbdata, Xtemp + 1)
    PointData(i, 2) = cbdata
Next i
Close 1


End Sub
