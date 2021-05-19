VERSION 5.00
Begin VB.Form frmGetSlope 
   Caption         =   "test"
   ClientHeight    =   7368
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   9492
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7368
   ScaleWidth      =   9492
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox AverageAngle 
      Height          =   300
      Left            =   8208
      TabIndex        =   85
      Text            =   "Text1"
      Top             =   6336
      Width           =   972
   End
   Begin VB.TextBox Angle1 
      Height          =   300
      Index           =   2
      Left            =   8208
      TabIndex        =   79
      Text            =   "Text3"
      Top             =   5760
      Width           =   972
   End
   Begin VB.TextBox AngleDatum 
      Height          =   348
      Index           =   2
      Left            =   8208
      TabIndex        =   78
      Text            =   "Text1"
      Top             =   4800
      Width           =   924
   End
   Begin VB.TextBox AngleLine 
      Height          =   348
      Index           =   2
      Left            =   8160
      TabIndex        =   77
      Text            =   "Text1"
      Top             =   5280
      Width           =   1020
   End
   Begin VB.TextBox Angle1 
      Height          =   300
      Index           =   1
      Left            =   8208
      TabIndex        =   72
      Text            =   "Text3"
      Top             =   3360
      Width           =   972
   End
   Begin VB.TextBox AngleDatum 
      Height          =   348
      Index           =   1
      Left            =   8208
      TabIndex        =   71
      Text            =   "Text1"
      Top             =   2400
      Width           =   924
   End
   Begin VB.TextBox AngleLine 
      Height          =   348
      Index           =   1
      Left            =   8160
      TabIndex        =   70
      Text            =   "Text1"
      Top             =   2880
      Width           =   1020
   End
   Begin VB.TextBox AngleLine 
      Height          =   348
      Index           =   0
      Left            =   7968
      TabIndex        =   65
      Text            =   "Text1"
      Top             =   768
      Width           =   1020
   End
   Begin VB.TextBox AngleDatum 
      Height          =   348
      Index           =   0
      Left            =   8016
      TabIndex        =   64
      Text            =   "Text1"
      Top             =   288
      Width           =   924
   End
   Begin VB.TextBox Diff3 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   2
      Left            =   3312
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   6864
      Width           =   876
   End
   Begin VB.TextBox Diff3 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   1
      Left            =   1872
      TabIndex        =   62
      Text            =   "Text1"
      Top             =   6912
      Width           =   876
   End
   Begin VB.TextBox Diff3 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   0
      Left            =   528
      TabIndex        =   61
      Text            =   "Text1"
      Top             =   6912
      Width           =   876
   End
   Begin VB.TextBox Diff2 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   2
      Left            =   3264
      TabIndex        =   60
      Text            =   "Text1"
      Top             =   4608
      Width           =   876
   End
   Begin VB.TextBox Diff2 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   1
      Left            =   1776
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   4608
      Width           =   876
   End
   Begin VB.TextBox Diff2 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   0
      Left            =   288
      TabIndex        =   58
      Text            =   "Text1"
      Top             =   4560
      Width           =   876
   End
   Begin VB.TextBox Diff1 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   2
      Left            =   3264
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   2304
      Width           =   876
   End
   Begin VB.TextBox Diff1 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   1
      Left            =   1824
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   2304
      Width           =   876
   End
   Begin VB.TextBox Diff1 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   0
      Left            =   384
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   2304
      Width           =   876
   End
   Begin VB.CommandButton Command7 
      Caption         =   "get average"
      Height          =   396
      Left            =   4896
      TabIndex        =   54
      Top             =   2400
      Width           =   1020
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Paste Datums"
      Height          =   396
      Left            =   4848
      TabIndex        =   53
      Top             =   480
      Width           =   1164
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   396
      Left            =   4896
      TabIndex        =   52
      Top             =   4032
      Width           =   1020
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Paste"
      Height          =   396
      Left            =   4896
      TabIndex        =   51
      Top             =   960
      Width           =   1020
   End
   Begin VB.CommandButton Command3 
      Caption         =   "get angle"
      Height          =   396
      Left            =   4896
      TabIndex        =   50
      Top             =   1968
      Width           =   1020
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Get Offset"
      Height          =   396
      Left            =   4896
      TabIndex        =   49
      Top             =   1440
      Width           =   1020
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Rotate Points"
      Height          =   396
      Left            =   4944
      TabIndex        =   48
      Top             =   2880
      Width           =   1020
   End
   Begin VB.TextBox Rotate3 
      Height          =   288
      Index           =   2
      Left            =   3360
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   6480
      Width           =   876
   End
   Begin VB.TextBox Rotate3 
      Height          =   288
      Index           =   1
      Left            =   1920
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   6480
      Width           =   876
   End
   Begin VB.TextBox Rotate3 
      Height          =   288
      Index           =   0
      Left            =   528
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   6480
      Width           =   876
   End
   Begin VB.TextBox Rotate2 
      Height          =   288
      Index           =   2
      Left            =   3312
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   4224
      Width           =   876
   End
   Begin VB.TextBox Rotate2 
      Height          =   288
      Index           =   1
      Left            =   1776
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   4224
      Width           =   876
   End
   Begin VB.TextBox Rotate2 
      Height          =   288
      Index           =   0
      Left            =   336
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   4224
      Width           =   876
   End
   Begin VB.TextBox Rotate1 
      Height          =   288
      Index           =   2
      Left            =   3312
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   1920
      Width           =   876
   End
   Begin VB.TextBox Rotate1 
      Height          =   288
      Index           =   1
      Left            =   1872
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   1920
      Width           =   876
   End
   Begin VB.TextBox Rotate1 
      Height          =   288
      Index           =   0
      Left            =   384
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   1920
      Width           =   876
   End
   Begin VB.TextBox Offset3 
      Height          =   288
      Index           =   2
      Left            =   3408
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   6096
      Width           =   876
   End
   Begin VB.TextBox Offset3 
      Height          =   288
      Index           =   1
      Left            =   1968
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   6096
      Width           =   876
   End
   Begin VB.TextBox Offset3 
      Height          =   288
      Index           =   0
      Left            =   528
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   6096
      Width           =   876
   End
   Begin VB.TextBox Offset2 
      Height          =   288
      Index           =   2
      Left            =   3312
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   3840
      Width           =   876
   End
   Begin VB.TextBox Offset2 
      Height          =   288
      Index           =   1
      Left            =   1776
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   3840
      Width           =   876
   End
   Begin VB.TextBox Offset2 
      Height          =   288
      Index           =   0
      Left            =   384
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   3840
      Width           =   876
   End
   Begin VB.TextBox Offset1 
      Height          =   288
      Index           =   2
      Left            =   3312
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   1440
      Width           =   876
   End
   Begin VB.TextBox Offset1 
      Height          =   288
      Index           =   1
      Left            =   1872
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   1440
      Width           =   876
   End
   Begin VB.TextBox Offset1 
      Height          =   288
      Index           =   0
      Left            =   384
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   1440
      Width           =   876
   End
   Begin VB.TextBox Point3 
      Height          =   288
      Index           =   2
      Left            =   3384
      TabIndex        =   27
      Top             =   5664
      Width           =   972
   End
   Begin VB.TextBox Point3 
      Height          =   288
      Index           =   1
      Left            =   1920
      TabIndex        =   26
      Text            =   ".163"
      Top             =   5664
      Width           =   972
   End
   Begin VB.TextBox Point3 
      Height          =   288
      Index           =   0
      Left            =   480
      TabIndex        =   25
      Text            =   "-.0885"
      Top             =   5664
      Width           =   972
   End
   Begin VB.TextBox Datum3 
      Height          =   300
      Index           =   2
      Left            =   3408
      TabIndex        =   23
      Text            =   "Text4"
      Top             =   5232
      Width           =   924
   End
   Begin VB.TextBox Datum3 
      Height          =   300
      Index           =   1
      Left            =   1920
      TabIndex        =   22
      Text            =   ".1525"
      Top             =   5232
      Width           =   924
   End
   Begin VB.TextBox Datum3 
      Height          =   300
      Index           =   0
      Left            =   528
      TabIndex        =   21
      Text            =   "0"
      Top             =   5232
      Width           =   924
   End
   Begin VB.TextBox Point2 
      Height          =   288
      Index           =   2
      Left            =   3240
      TabIndex        =   17
      Top             =   3408
      Width           =   972
   End
   Begin VB.TextBox Point2 
      Height          =   288
      Index           =   1
      Left            =   1776
      TabIndex        =   16
      Text            =   ".0441"
      Top             =   3408
      Width           =   972
   End
   Begin VB.TextBox Point2 
      Height          =   288
      Index           =   0
      Left            =   336
      TabIndex        =   15
      Text            =   ".0933"
      Top             =   3408
      Width           =   972
   End
   Begin VB.TextBox Datum2 
      Height          =   300
      Index           =   2
      Left            =   3264
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   2976
      Width           =   924
   End
   Begin VB.TextBox Datum2 
      Height          =   300
      Index           =   1
      Left            =   1776
      TabIndex        =   12
      Text            =   "0"
      Top             =   2976
      Width           =   924
   End
   Begin VB.TextBox Datum2 
      Height          =   300
      Index           =   0
      Left            =   384
      TabIndex        =   11
      Text            =   ".1525"
      Top             =   2976
      Width           =   924
   End
   Begin VB.TextBox Angle1 
      Height          =   300
      Index           =   0
      Left            =   8016
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   1248
      Width           =   972
   End
   Begin VB.TextBox Datum1 
      Height          =   300
      Index           =   2
      Left            =   3264
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   480
      Width           =   924
   End
   Begin VB.TextBox Datum1 
      Height          =   300
      Index           =   1
      Left            =   1776
      TabIndex        =   7
      Text            =   "0"
      Top             =   480
      Width           =   924
   End
   Begin VB.TextBox Datum1 
      Height          =   300
      Index           =   0
      Left            =   384
      TabIndex        =   6
      Text            =   "0"
      Top             =   480
      Width           =   924
   End
   Begin VB.TextBox Point1 
      Height          =   288
      Index           =   0
      Left            =   384
      TabIndex        =   2
      Text            =   "-.0564"
      Top             =   1008
      Width           =   972
   End
   Begin VB.TextBox Point1 
      Height          =   288
      Index           =   1
      Left            =   1824
      TabIndex        =   1
      Text            =   ".0138"
      Top             =   1008
      Width           =   972
   End
   Begin VB.TextBox Point1 
      Height          =   288
      Index           =   2
      Left            =   3288
      TabIndex        =   0
      Top             =   1008
      Width           =   972
   End
   Begin VB.Label Label7 
      Caption         =   "Average Angle"
      Height          =   204
      Left            =   6336
      TabIndex        =   84
      Top             =   6336
      Width           =   1404
   End
   Begin VB.Label Label6 
      Caption         =   "2-3"
      Height          =   252
      Index           =   2
      Left            =   6576
      TabIndex        =   83
      Top             =   4560
      Width           =   1116
   End
   Begin VB.Label Label5 
      Caption         =   "Angle Difference"
      Height          =   204
      Index           =   2
      Left            =   6624
      TabIndex        =   82
      Top             =   5760
      Width           =   1020
   End
   Begin VB.Label Label4 
      Caption         =   "Angle Line"
      Height          =   204
      Index           =   2
      Left            =   6624
      TabIndex        =   81
      Top             =   5376
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Angle Datum "
      Height          =   204
      Index           =   2
      Left            =   6624
      TabIndex        =   80
      Top             =   4896
      Width           =   1020
   End
   Begin VB.Label Label6 
      Caption         =   "1-3"
      Height          =   252
      Index           =   1
      Left            =   6576
      TabIndex        =   76
      Top             =   2160
      Width           =   1116
   End
   Begin VB.Label Label5 
      Caption         =   "Angle Difference"
      Height          =   204
      Index           =   1
      Left            =   6624
      TabIndex        =   75
      Top             =   3360
      Width           =   1020
   End
   Begin VB.Label Label4 
      Caption         =   "Angle Line"
      Height          =   204
      Index           =   1
      Left            =   6624
      TabIndex        =   74
      Top             =   2976
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Angle Datum "
      Height          =   204
      Index           =   1
      Left            =   6624
      TabIndex        =   73
      Top             =   2496
      Width           =   1020
   End
   Begin VB.Label Label6 
      Caption         =   "1-2"
      Height          =   252
      Index           =   0
      Left            =   6384
      TabIndex        =   69
      Top             =   48
      Width           =   1116
   End
   Begin VB.Label Label5 
      Caption         =   "Angle Difference"
      Height          =   204
      Index           =   0
      Left            =   6432
      TabIndex        =   68
      Top             =   1248
      Width           =   1020
   End
   Begin VB.Label Label4 
      Caption         =   "Angle Line"
      Height          =   204
      Index           =   0
      Left            =   6432
      TabIndex        =   67
      Top             =   864
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Angle Datum "
      Height          =   204
      Index           =   0
      Left            =   6432
      TabIndex        =   66
      Top             =   384
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Z:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   7
      Left            =   3144
      TabIndex        =   29
      Top             =   5688
      Width           =   168
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   6
      Left            =   1680
      TabIndex        =   28
      Top             =   5688
      Width           =   180
   End
   Begin VB.Label Label3 
      Caption         =   "Datum3"
      Height          =   204
      Index           =   2
      Left            =   288
      TabIndex        =   24
      Top             =   4944
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Z:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   5
      Left            =   3000
      TabIndex        =   20
      Top             =   3432
      Width           =   168
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   4
      Left            =   1536
      TabIndex        =   19
      Top             =   3432
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   1
      Left            =   96
      TabIndex        =   18
      Top             =   3432
      Width           =   168
   End
   Begin VB.Label Label3 
      Caption         =   "Datum2"
      Height          =   204
      Index           =   1
      Left            =   144
      TabIndex        =   14
      Top             =   2688
      Width           =   1500
   End
   Begin VB.Label Label3 
      Caption         =   "Datum1"
      Height          =   204
      Index           =   0
      Left            =   96
      TabIndex        =   10
      Top             =   192
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   0
      Left            =   144
      TabIndex        =   5
      Top             =   1032
      Width           =   168
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   2
      Left            =   1584
      TabIndex        =   4
      Top             =   1032
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Z:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   3
      Left            =   3048
      TabIndex        =   3
      Top             =   1032
      Width           =   168
   End
End
Attribute VB_Name = "frmGetSlope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Loading As Boolean
Dim PointData(3, 3) As Single

Private Sub Angle2_Change()

End Sub

Private Sub Command1_Click()
Cancelling = True
Unload Me
'frmMain.cmdCancel_Click

End Sub

Private Sub Command2_Click()
GetData "testpoints.txt"

Point1(0) = PointData(1, 0)
Point1(1) = PointData(1, 1)
Point1(2) = PointData(1, 2)
Point2(0) = PointData(2, 0)
Point2(1) = PointData(2, 1)
Point2(2) = PointData(2, 2)
Point3(0) = PointData(3, 0)
Point3(1) = PointData(3, 1)
Point3(2) = PointData(3, 2)

End Sub

Private Sub Command3_Click()
Dim Y1 As Single
Dim Y2 As Single
Dim X1 As Single
Dim X2 As Single
Dim HLineAngle As Single
Dim LineAngle As Single
Dim DatumAngle As Single
Dim Direction As Integer


'line 1-2
Y1 = Val(Datum1(1))
Y2 = Val(Datum2(1))
X1 = Val(Datum1(0))
X2 = Val(Datum2(0))
GoSub ConvertAngleFromNorth
AngleDatum(0) = HLineAngle
DatumAngle = HLineAngle

Y1 = Val(Point1(1))
Y2 = Val(Point2(1))
X1 = Val(Point1(0))
X2 = Val(Point2(0))
GoSub ConvertAngleFromNorth
AngleLine(0) = HLineAngle
LineAngle = HLineAngle
If LineAngle > DatumAngle Then
    Angle1(0) = LineAngle - DatumAngle
Else
    Angle1(0) = (360 - DatumAngle) + LineAngle
End If


'Line 1-3
Y1 = Val(Datum1(1))
Y2 = Val(Datum3(1))
X1 = Val(Datum1(0))
X2 = Val(Datum3(0))
GoSub ConvertAngleFromNorth
DatumAngle = HLineAngle
AngleDatum(1) = HLineAngle

Y1 = Val(Point1(1))
Y2 = Val(Point3(1))
X1 = Val(Point1(0))
X2 = Val(Point3(0))
GoSub ConvertAngleFromNorth
AngleLine(1) = HLineAngle
LineAngle = HLineAngle
If LineAngle > DatumAngle Then
    Angle1(1) = LineAngle - DatumAngle
Else
    Angle1(1) = (360 - DatumAngle) + LineAngle
End If



'Line 2-3
Y1 = Val(Datum2(1))
Y2 = Val(Datum3(1))
X1 = Val(Datum2(0))
X2 = Val(Datum3(0))
GoSub ConvertAngleFromNorth
DatumAngle = HLineAngle
AngleDatum(2) = HLineAngle

Y1 = Val(Point2(1))
Y2 = Val(Point3(1))
X1 = Val(Point2(0))
X2 = Val(Point3(0))
GoSub ConvertAngleFromNorth
AngleLine(2) = HLineAngle
LineAngle = HLineAngle
If LineAngle > DatumAngle Then
    Angle1(2) = LineAngle - DatumAngle
Else
    Angle1(2) = (360 - DatumAngle) + LineAngle
End If
AverageAngle = Format((Val(Angle1(0)) + Val(Angle1(1)) + Val(Angle1(2))) / 3, "#.####")


Exit Sub
ConvertAngleFromNorth:
    slopetan = (Y2 - Y1) / (X2 - X1)
    HLineAngle = Atn(slopetan) * 57.2958
    
    If X2 > X1 Then
        HLineAngle = 90 - HLineAngle
    Else
        HLineAngle = 270 - HLineAngle
    End If
Return

End Sub




Private Sub Command4_Click()
    
Xorigin = Offset1(0)
Yorigin = Offset1(1)
Sinangle = Sin(-AverageAngle * 1.74532925199433E-02)
Cosangle = Cos(-AverageAngle * 1.74532925199433E-02)

X = Val(Offset1(0)) - Xorigin
Y = Val(Offset1(1)) - Yorigin
Xtemp = X * Cosangle + Y * Sinangle
Ytemp = Y * Cosangle - X * Sinangle
Rotate1(0) = Format(Xtemp + Xorigin, "#.#####")
Rotate1(1) = Format(Ytemp + Yorigin, "#.#####")

X = Val(Offset2(0)) - Xorigin
Y = Val(Offset2(1)) - Yorigin
Xtemp = X * Cosangle + Y * Sinangle
Ytemp = Y * Cosangle - X * Sinangle
Rotate2(0) = Format(Xtemp + Xorigin, "#.#####")
Rotate2(1) = Format(Ytemp + Yorigin, "#.#####")

X = Val(Offset3(0)) - Xorigin
Y = Val(Offset3(1)) - Yorigin
Xtemp = X * Cosangle + Y * Sinangle
Ytemp = Y * Cosangle - X * Sinangle
Rotate3(0) = Format(Xtemp + Xorigin, "#.#####")
Rotate3(1) = Format(Ytemp + Yorigin, "#.#####")

Diff1(0) = Format(Val(Rotate1(0)) - Val(Datum1(0)), "#.####")
Diff1(1) = Format(Val(Rotate1(1)) - Val(Datum1(1)), "#.####")
Diff2(0) = Format(Val(Rotate2(0)) - Val(Datum2(0)), "#.####")
Diff2(1) = Format(Val(Rotate2(1)) - Val(Datum2(1)), "#.####")
Diff3(0) = Format(Val(Rotate3(0)) - Val(Datum3(0)), "#.####")
Diff3(1) = Format(Val(Rotate3(1)) - Val(Datum3(1)), "#.####")

End Sub

Private Sub Command5_Click()
GetData "testdatums.txt"

Datum1(0) = PointData(1, 0)
Datum1(1) = PointData(1, 1)
Datum1(2) = PointData(1, 2)
Datum2(0) = PointData(2, 0)
Datum2(1) = PointData(2, 1)
Datum2(2) = PointData(2, 2)
Datum3(0) = PointData(3, 0)
Datum3(1) = PointData(3, 1)
Datum3(2) = PointData(3, 2)
End Sub

Private Sub Command6_Click()
Dim OffsetX As Single
Dim OffsetY As Single

OffsetX = Val(Point1(0)) - Val(Datum1(0))
OffsetY = Val(Point1(1)) - Val(Datum1(1))

Offset1(0) = Format(Val(Point1(0)) - OffsetX, "0.0000")
Offset1(1) = Format(Val(Point1(1)) - OffsetY, "0.0000")
Offset2(0) = Format(Val(Point2(0)) - OffsetX, "0.0000")
Offset2(1) = Format(Val(Point2(1)) - OffsetY, "0.0000")
Offset3(0) = Format(Val(Point3(0)) - OffsetX, "0.0000")
Offset3(1) = Format(Val(Point3(1)) - OffsetY, "0.0000")




End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Text1(0) = Chr(KeyAscii)
    KeyAscii = 0
    Text1(0).SetFocus
    Text1(0).SelStart = 1
    Me.KeyPreview = False
    
End Sub

Private Sub Form_Load()
Screen.MousePointer = 1
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
                    Unload Me
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
