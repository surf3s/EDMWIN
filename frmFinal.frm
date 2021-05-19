VERSION 5.00
Begin VB.Form frmFinal 
   Caption         =   "Final Points"
   ClientHeight    =   6636
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6732
   LinkTopic       =   "Form2"
   ScaleHeight     =   6636
   ScaleWidth      =   6732
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Load Data"
      Height          =   348
      Left            =   4848
      TabIndex        =   57
      Top             =   384
      Width           =   1164
   End
   Begin VB.TextBox Final3 
      Height          =   288
      Index           =   2
      Left            =   3600
      TabIndex        =   55
      Top             =   5520
      Width           =   972
   End
   Begin VB.TextBox Final3 
      Height          =   288
      Index           =   1
      Left            =   2544
      TabIndex        =   54
      Top             =   5520
      Width           =   972
   End
   Begin VB.TextBox Final3 
      Height          =   288
      Index           =   0
      Left            =   1536
      TabIndex        =   53
      Top             =   5520
      Width           =   972
   End
   Begin VB.TextBox Diff3 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   2
      Left            =   3600
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   5952
      Width           =   972
   End
   Begin VB.TextBox Diff3 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   1
      Left            =   2544
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   5952
      Width           =   972
   End
   Begin VB.TextBox Diff3 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   0
      Left            =   1536
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   5952
      Width           =   972
   End
   Begin VB.TextBox Point3 
      Height          =   288
      Index           =   2
      Left            =   3600
      TabIndex        =   43
      Top             =   5088
      Width           =   972
   End
   Begin VB.TextBox Point3 
      Height          =   288
      Index           =   1
      Left            =   2544
      TabIndex        =   42
      Top             =   5088
      Width           =   972
   End
   Begin VB.TextBox Point3 
      Height          =   288
      Index           =   0
      Left            =   1536
      TabIndex        =   41
      Top             =   5088
      Width           =   972
   End
   Begin VB.TextBox Datum3 
      Height          =   288
      Index           =   2
      Left            =   3600
      TabIndex        =   40
      Top             =   4608
      Width           =   972
   End
   Begin VB.TextBox Datum3 
      Height          =   288
      Index           =   1
      Left            =   2544
      TabIndex        =   39
      Top             =   4608
      Width           =   972
   End
   Begin VB.TextBox Datum3 
      Height          =   288
      Index           =   0
      Left            =   1536
      TabIndex        =   38
      Top             =   4608
      Width           =   972
   End
   Begin VB.TextBox Final2 
      Height          =   288
      Index           =   2
      Left            =   3600
      TabIndex        =   36
      Top             =   3408
      Width           =   972
   End
   Begin VB.TextBox Final2 
      Height          =   288
      Index           =   1
      Left            =   2544
      TabIndex        =   35
      Top             =   3408
      Width           =   972
   End
   Begin VB.TextBox Final2 
      Height          =   288
      Index           =   0
      Left            =   1536
      TabIndex        =   34
      Top             =   3408
      Width           =   972
   End
   Begin VB.TextBox Diff1 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   2
      Left            =   3600
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   3840
      Width           =   972
   End
   Begin VB.TextBox Diff1 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   1
      Left            =   2544
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   3840
      Width           =   972
   End
   Begin VB.TextBox Diff1 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   0
      Left            =   1536
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   3840
      Width           =   972
   End
   Begin VB.TextBox Point2 
      Height          =   288
      Index           =   2
      Left            =   3600
      TabIndex        =   24
      Top             =   2976
      Width           =   972
   End
   Begin VB.TextBox Point2 
      Height          =   288
      Index           =   1
      Left            =   2544
      TabIndex        =   23
      Top             =   2976
      Width           =   972
   End
   Begin VB.TextBox Point2 
      Height          =   288
      Index           =   0
      Left            =   1536
      TabIndex        =   22
      Top             =   2976
      Width           =   972
   End
   Begin VB.TextBox Datum2 
      Height          =   288
      Index           =   2
      Left            =   3600
      TabIndex        =   21
      Top             =   2496
      Width           =   972
   End
   Begin VB.TextBox Datum2 
      Height          =   288
      Index           =   1
      Left            =   2544
      TabIndex        =   20
      Top             =   2496
      Width           =   972
   End
   Begin VB.TextBox Datum2 
      Height          =   288
      Index           =   0
      Left            =   1536
      TabIndex        =   19
      Top             =   2496
      Width           =   972
   End
   Begin VB.TextBox Final1 
      Height          =   288
      Index           =   2
      Left            =   3600
      TabIndex        =   17
      Top             =   1200
      Width           =   972
   End
   Begin VB.TextBox Final1 
      Height          =   288
      Index           =   1
      Left            =   2544
      TabIndex        =   16
      Top             =   1200
      Width           =   972
   End
   Begin VB.TextBox Final1 
      Height          =   288
      Index           =   0
      Left            =   1536
      TabIndex        =   15
      Top             =   1200
      Width           =   972
   End
   Begin VB.TextBox Diff 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   2
      Left            =   3600
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1632
      Width           =   972
   End
   Begin VB.TextBox Diff 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   1
      Left            =   2544
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1632
      Width           =   972
   End
   Begin VB.TextBox Diff 
      ForeColor       =   &H000000FF&
      Height          =   288
      Index           =   0
      Left            =   1536
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1632
      Width           =   972
   End
   Begin VB.TextBox Point1 
      Height          =   288
      Index           =   2
      Left            =   3600
      TabIndex        =   5
      Top             =   768
      Width           =   972
   End
   Begin VB.TextBox Point1 
      Height          =   288
      Index           =   1
      Left            =   2544
      TabIndex        =   4
      Top             =   768
      Width           =   972
   End
   Begin VB.TextBox Point1 
      Height          =   288
      Index           =   0
      Left            =   1536
      TabIndex        =   3
      Top             =   768
      Width           =   972
   End
   Begin VB.TextBox Datum1 
      Height          =   288
      Index           =   2
      Left            =   3600
      TabIndex        =   2
      Top             =   288
      Width           =   972
   End
   Begin VB.TextBox Datum1 
      Height          =   288
      Index           =   1
      Left            =   2544
      TabIndex        =   1
      Top             =   288
      Width           =   972
   End
   Begin VB.TextBox Datum1 
      Height          =   288
      Index           =   0
      Left            =   1536
      TabIndex        =   0
      Top             =   288
      Width           =   972
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
      Index           =   11
      Left            =   588
      TabIndex        =   56
      Top             =   6000
      Width           =   864
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Final"
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
      Left            =   1032
      TabIndex        =   49
      Top             =   5616
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Z"
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
      Left            =   4128
      TabIndex        =   48
      Top             =   4368
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Y"
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
      Left            =   2880
      TabIndex        =   47
      Top             =   4368
      Width           =   132
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "X"
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
      TabIndex        =   46
      Top             =   4368
      Width           =   120
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
      Index           =   9
      Left            =   216
      TabIndex        =   45
      Top             =   4656
      Width           =   1236
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
      Index           =   8
      Left            =   144
      TabIndex        =   44
      Top             =   5184
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
      Index           =   7
      Left            =   588
      TabIndex        =   37
      Top             =   3888
      Width           =   864
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Final"
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
      Left            =   1032
      TabIndex        =   30
      Top             =   3504
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Z"
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
      Left            =   4128
      TabIndex        =   29
      Top             =   2256
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Y"
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
      Left            =   2880
      TabIndex        =   28
      Top             =   2256
      Width           =   132
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "X"
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
      Left            =   1680
      TabIndex        =   27
      Top             =   2256
      Width           =   120
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
      Index           =   5
      Left            =   216
      TabIndex        =   26
      Top             =   2544
      Width           =   1236
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
      Index           =   3
      Left            =   144
      TabIndex        =   25
      Top             =   3072
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
      Index           =   2
      Left            =   588
      TabIndex        =   18
      Top             =   1680
      Width           =   864
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Final"
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
      Left            =   1032
      TabIndex        =   11
      Top             =   1296
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Z"
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
      Left            =   4128
      TabIndex        =   10
      Top             =   48
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Y"
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
      Left            =   2880
      TabIndex        =   9
      Top             =   48
      Width           =   132
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "X"
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
      Left            =   1680
      TabIndex        =   8
      Top             =   48
      Width           =   120
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
      Left            =   216
      TabIndex        =   7
      Top             =   336
      Width           =   1236
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
      Left            =   144
      TabIndex        =   6
      Top             =   864
      Width           =   1308
   End
End
Attribute VB_Name = "frmFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
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


