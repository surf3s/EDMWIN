VERSION 5.00
Begin VB.Form frmGetSlopeXZ 
   Caption         =   "Test X-Z"
   ClientHeight    =   7656
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   12684
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7656
   ScaleWidth      =   12684
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox AverageAngle 
      Height          =   300
      Left            =   9600
      TabIndex        =   77
      Text            =   "Text1"
      Top             =   6384
      Width           =   972
   End
   Begin VB.TextBox Angle1 
      Height          =   300
      Index           =   2
      Left            =   9600
      TabIndex        =   71
      Text            =   "Text3"
      Top             =   5808
      Width           =   972
   End
   Begin VB.TextBox AngleDatum 
      Height          =   348
      Index           =   2
      Left            =   9600
      TabIndex        =   70
      Text            =   "Text1"
      Top             =   4848
      Width           =   924
   End
   Begin VB.TextBox AngleLine 
      Height          =   348
      Index           =   2
      Left            =   9552
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   5328
      Width           =   1020
   End
   Begin VB.TextBox Angle1 
      Height          =   300
      Index           =   1
      Left            =   9600
      TabIndex        =   64
      Text            =   "Text3"
      Top             =   3408
      Width           =   972
   End
   Begin VB.TextBox AngleDatum 
      Height          =   348
      Index           =   1
      Left            =   9600
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   2448
      Width           =   924
   End
   Begin VB.TextBox AngleLine 
      Height          =   348
      Index           =   1
      Left            =   9552
      TabIndex        =   62
      Text            =   "Text1"
      Top             =   2928
      Width           =   1020
   End
   Begin VB.TextBox AngleLine 
      Height          =   348
      Index           =   0
      Left            =   9360
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   816
      Width           =   1020
   End
   Begin VB.TextBox AngleDatum 
      Height          =   348
      Index           =   0
      Left            =   9408
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   336
      Width           =   924
   End
   Begin VB.TextBox Diff3 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   2
      Left            =   5112
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   7008
      Width           =   972
   End
   Begin VB.TextBox Diff3 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   1
      Left            =   3648
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   7056
      Width           =   972
   End
   Begin VB.TextBox Diff3 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   0
      Left            =   2352
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   7056
      Width           =   972
   End
   Begin VB.TextBox Diff2 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   2
      Left            =   5112
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   4656
      Width           =   972
   End
   Begin VB.TextBox Diff2 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   1
      Left            =   3648
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   4656
      Width           =   972
   End
   Begin VB.TextBox Diff2 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   0
      Left            =   2352
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   4608
      Width           =   972
   End
   Begin VB.TextBox Diff1 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   2
      Left            =   5112
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   2352
      Width           =   972
   End
   Begin VB.TextBox Diff1 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   1
      Left            =   3648
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   2352
      Width           =   972
   End
   Begin VB.TextBox Diff1 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   0
      Left            =   2352
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   2352
      Width           =   972
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Paste Datums"
      Height          =   396
      Left            =   6240
      TabIndex        =   46
      Top             =   528
      Width           =   1164
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   396
      Left            =   6288
      TabIndex        =   45
      Top             =   4080
      Width           =   1020
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Paste"
      Height          =   396
      Left            =   6288
      TabIndex        =   44
      Top             =   1008
      Width           =   1020
   End
   Begin VB.CommandButton Command3 
      Caption         =   "get angle"
      Height          =   396
      Left            =   6288
      TabIndex        =   43
      Top             =   2016
      Width           =   1020
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Get Offset"
      Height          =   396
      Left            =   6288
      TabIndex        =   42
      Top             =   1488
      Width           =   1020
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Rotate Points"
      Height          =   396
      Left            =   6288
      TabIndex        =   41
      Top             =   2544
      Width           =   1020
   End
   Begin VB.TextBox Rotate3 
      Height          =   288
      Index           =   2
      Left            =   5112
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   6624
      Width           =   972
   End
   Begin VB.TextBox Rotate3 
      Height          =   288
      Index           =   1
      Left            =   3648
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   6624
      Width           =   972
   End
   Begin VB.TextBox Rotate3 
      Height          =   288
      Index           =   0
      Left            =   2352
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   6624
      Width           =   972
   End
   Begin VB.TextBox Rotate2 
      Height          =   288
      Index           =   2
      Left            =   5112
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   4272
      Width           =   972
   End
   Begin VB.TextBox Rotate2 
      Height          =   288
      Index           =   1
      Left            =   3648
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   4272
      Width           =   972
   End
   Begin VB.TextBox Rotate2 
      Height          =   288
      Index           =   0
      Left            =   2352
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   4272
      Width           =   972
   End
   Begin VB.TextBox Rotate1 
      Height          =   288
      Index           =   2
      Left            =   5112
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   1968
      Width           =   972
   End
   Begin VB.TextBox Rotate1 
      Height          =   288
      Index           =   1
      Left            =   3648
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   1968
      Width           =   972
   End
   Begin VB.TextBox Rotate1 
      Height          =   288
      Index           =   0
      Left            =   2352
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   1968
      Width           =   972
   End
   Begin VB.TextBox Offset3 
      Height          =   288
      Index           =   2
      Left            =   5112
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   6240
      Width           =   972
   End
   Begin VB.TextBox Offset3 
      Height          =   288
      Index           =   1
      Left            =   3648
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   6240
      Width           =   972
   End
   Begin VB.TextBox Offset3 
      Height          =   288
      Index           =   0
      Left            =   2352
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   6240
      Width           =   972
   End
   Begin VB.TextBox Offset2 
      Height          =   288
      Index           =   2
      Left            =   5112
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   3888
      Width           =   972
   End
   Begin VB.TextBox Offset2 
      Height          =   288
      Index           =   1
      Left            =   3648
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   3888
      Width           =   972
   End
   Begin VB.TextBox Offset2 
      Height          =   288
      Index           =   0
      Left            =   2352
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   3888
      Width           =   972
   End
   Begin VB.TextBox Offset1 
      Height          =   288
      Index           =   2
      Left            =   5112
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   1488
      Width           =   972
   End
   Begin VB.TextBox Offset1 
      Height          =   288
      Index           =   1
      Left            =   3648
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   1488
      Width           =   972
   End
   Begin VB.TextBox Offset1 
      Height          =   288
      Index           =   0
      Left            =   2352
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   1488
      Width           =   972
   End
   Begin VB.TextBox Point3 
      Height          =   288
      Index           =   2
      Left            =   5112
      TabIndex        =   22
      Top             =   5808
      Width           =   972
   End
   Begin VB.TextBox Point3 
      Height          =   288
      Index           =   1
      Left            =   3648
      TabIndex        =   21
      Text            =   ".163"
      Top             =   5808
      Width           =   972
   End
   Begin VB.TextBox Point3 
      Height          =   288
      Index           =   0
      Left            =   2352
      TabIndex        =   20
      Text            =   "-.0885"
      Top             =   5808
      Width           =   972
   End
   Begin VB.TextBox Datum3 
      Height          =   288
      Index           =   2
      Left            =   5112
      TabIndex        =   18
      Text            =   "Text4"
      Top             =   5376
      Width           =   972
   End
   Begin VB.TextBox Datum3 
      Height          =   288
      Index           =   1
      Left            =   3648
      TabIndex        =   17
      Text            =   ".1525"
      Top             =   5376
      Width           =   972
   End
   Begin VB.TextBox Datum3 
      Height          =   288
      Index           =   0
      Left            =   2352
      TabIndex        =   16
      Text            =   "0"
      Top             =   5376
      Width           =   972
   End
   Begin VB.TextBox Point2 
      Height          =   288
      Index           =   2
      Left            =   5112
      TabIndex        =   15
      Top             =   3456
      Width           =   972
   End
   Begin VB.TextBox Point2 
      Height          =   288
      Index           =   1
      Left            =   3648
      TabIndex        =   14
      Text            =   ".0441"
      Top             =   3456
      Width           =   972
   End
   Begin VB.TextBox Point2 
      Height          =   288
      Index           =   0
      Left            =   2352
      TabIndex        =   13
      Text            =   ".0933"
      Top             =   3456
      Width           =   972
   End
   Begin VB.TextBox Datum2 
      Height          =   288
      Index           =   2
      Left            =   5112
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   3024
      Width           =   972
   End
   Begin VB.TextBox Datum2 
      Height          =   288
      Index           =   1
      Left            =   3648
      TabIndex        =   10
      Text            =   "0"
      Top             =   3024
      Width           =   972
   End
   Begin VB.TextBox Datum2 
      Height          =   288
      Index           =   0
      Left            =   2352
      TabIndex        =   9
      Text            =   ".1525"
      Top             =   3024
      Width           =   972
   End
   Begin VB.TextBox Angle1 
      Height          =   300
      Index           =   0
      Left            =   9408
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   1296
      Width           =   972
   End
   Begin VB.TextBox Datum1 
      Height          =   288
      Index           =   2
      Left            =   5112
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   528
      Width           =   972
   End
   Begin VB.TextBox Datum1 
      Height          =   288
      Index           =   1
      Left            =   3648
      TabIndex        =   5
      Text            =   "0"
      Top             =   528
      Width           =   972
   End
   Begin VB.TextBox Datum1 
      Height          =   288
      Index           =   0
      Left            =   2352
      TabIndex        =   4
      Text            =   "0"
      Top             =   528
      Width           =   972
   End
   Begin VB.TextBox Point1 
      Height          =   288
      Index           =   0
      Left            =   2352
      TabIndex        =   2
      Text            =   "-.0564"
      Top             =   1056
      Width           =   972
   End
   Begin VB.TextBox Point1 
      Height          =   288
      Index           =   1
      Left            =   3648
      TabIndex        =   1
      Text            =   ".0138"
      Top             =   1056
      Width           =   972
   End
   Begin VB.TextBox Point1 
      Height          =   288
      Index           =   2
      Left            =   5112
      TabIndex        =   0
      Top             =   1056
      Width           =   972
   End
   Begin VB.Label Title 
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   2352
      TabIndex        =   92
      Top             =   144
      Width           =   4332
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Offset"
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
      Index           =   17
      Left            =   1344
      TabIndex        =   91
      Top             =   6336
      Width           =   492
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Recorded Point"
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
      Index           =   16
      Left            =   528
      TabIndex        =   90
      Top             =   5952
      Width           =   1308
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Difference"
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
      Index           =   15
      Left            =   972
      TabIndex        =   89
      Top             =   7056
      Width           =   864
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Rotated"
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
      Index           =   14
      Left            =   1164
      TabIndex        =   88
      Top             =   6672
      Width           =   672
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Original Datum"
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
      Index           =   13
      Left            =   600
      TabIndex        =   87
      Top             =   5424
      Width           =   1236
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Offset"
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
      Index           =   12
      Left            =   1344
      TabIndex        =   86
      Top             =   3984
      Width           =   492
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Recorded Point"
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
      Index           =   11
      Left            =   528
      TabIndex        =   85
      Top             =   3600
      Width           =   1308
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Difference"
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
      Index           =   10
      Left            =   972
      TabIndex        =   84
      Top             =   4752
      Width           =   864
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Rotated"
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
      Index           =   9
      Left            =   1164
      TabIndex        =   83
      Top             =   4368
      Width           =   672
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Original Datum"
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
      Index           =   8
      Left            =   600
      TabIndex        =   82
      Top             =   3072
      Width           =   1236
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Offset"
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
      Left            =   1320
      TabIndex        =   81
      Top             =   1488
      Width           =   492
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Recorded Point"
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
      Left            =   504
      TabIndex        =   80
      Top             =   1104
      Width           =   1308
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Difference"
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
      Left            =   948
      TabIndex        =   79
      Top             =   2400
      Width           =   864
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Rotated"
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
      Left            =   1140
      TabIndex        =   78
      Top             =   2016
      Width           =   672
   End
   Begin VB.Label Label7 
      Caption         =   "Average Angle"
      Height          =   204
      Left            =   7728
      TabIndex        =   76
      Top             =   6384
      Width           =   1404
   End
   Begin VB.Label Label6 
      Caption         =   "2-3"
      Height          =   252
      Index           =   2
      Left            =   7968
      TabIndex        =   75
      Top             =   4608
      Width           =   1116
   End
   Begin VB.Label Label5 
      Caption         =   "Angle Difference"
      Height          =   204
      Index           =   2
      Left            =   8016
      TabIndex        =   74
      Top             =   5808
      Width           =   1020
   End
   Begin VB.Label Label4 
      Caption         =   "Angle Line"
      Height          =   204
      Index           =   2
      Left            =   8016
      TabIndex        =   73
      Top             =   5424
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Angle Datum "
      Height          =   204
      Index           =   2
      Left            =   8016
      TabIndex        =   72
      Top             =   4944
      Width           =   1020
   End
   Begin VB.Label Label6 
      Caption         =   "1-3"
      Height          =   252
      Index           =   1
      Left            =   7968
      TabIndex        =   68
      Top             =   2208
      Width           =   1116
   End
   Begin VB.Label Label5 
      Caption         =   "Angle Difference"
      Height          =   204
      Index           =   1
      Left            =   8016
      TabIndex        =   67
      Top             =   3408
      Width           =   1020
   End
   Begin VB.Label Label4 
      Caption         =   "Angle Line"
      Height          =   204
      Index           =   1
      Left            =   8016
      TabIndex        =   66
      Top             =   3024
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Angle Datum "
      Height          =   204
      Index           =   1
      Left            =   8016
      TabIndex        =   65
      Top             =   2544
      Width           =   1020
   End
   Begin VB.Label Label6 
      Caption         =   "1-2"
      Height          =   252
      Index           =   0
      Left            =   7776
      TabIndex        =   61
      Top             =   96
      Width           =   1116
   End
   Begin VB.Label Label5 
      Caption         =   "Angle Difference"
      Height          =   204
      Index           =   0
      Left            =   7824
      TabIndex        =   60
      Top             =   1296
      Width           =   1020
   End
   Begin VB.Label Label4 
      Caption         =   "Angle Line"
      Height          =   204
      Index           =   0
      Left            =   7824
      TabIndex        =   59
      Top             =   912
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Angle Datum "
      Height          =   204
      Index           =   0
      Left            =   7824
      TabIndex        =   58
      Top             =   432
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Datum3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   2
      Left            =   288
      TabIndex        =   19
      Top             =   5040
      Width           =   852
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Datum2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   1
      Left            =   288
      TabIndex        =   12
      Top             =   2688
      Width           =   852
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Datum1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   0
      Left            =   288
      TabIndex        =   8
      Top             =   144
      Width           =   852
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Original Datum"
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
      Left            =   576
      TabIndex        =   3
      Top             =   576
      Width           =   1236
   End
End
Attribute VB_Name = "frmGetSlopeXZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Loading As Boolean


Private Sub Angle2_Change()

End Sub

Private Sub Command1_Click()
Cancelling = True
Unload Me
'frmMain.cmdCancel_Click

End Sub

Private Sub Command2_Click()
GetData PointFile

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
Set CallingForm = Me
GetAngle XZ

End Sub




Private Sub Command4_Click()
Set CallingForm = Me
RotatePoints XZ

End Sub

Private Sub Command5_Click()
GetData DatumFile

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
Set CallingForm = Me
GetOffset






End Sub

Private Sub Command7_Click()

End Sub

Private Sub Form_Load()
Me.Show
Title = TitleString
Screen.MousePointer = 1
Command5_Click
Command2_Click
Command6_Click
Command3_Click
Command4_Click
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
