VERSION 5.00
Begin VB.Form frmGetMSData 
   Caption         =   "Get Microscribe Data"
   ClientHeight    =   1305
   ClientLeft      =   7500
   ClientTop       =   5445
   ClientWidth     =   5700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   5700
   Begin VB.CommandButton Command2 
      Caption         =   "Paste"
      Height          =   252
      Left            =   2112
      TabIndex        =   8
      Top             =   960
      Width           =   1116
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   396
      Left            =   4464
      TabIndex        =   7
      Top             =   96
      Width           =   1020
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   0
      Left            =   336
      TabIndex        =   2
      Top             =   576
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   1
      Left            =   1776
      TabIndex        =   1
      Top             =   576
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   2
      Left            =   3336
      TabIndex        =   0
      Top             =   576
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Press Microscribe pedal to record data point, or press Paste button to copy data from clipboard."
      Height          =   348
      Left            =   192
      TabIndex        =   6
      Top             =   96
      Width           =   4092
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   0
      Left            =   96
      TabIndex        =   5
      Top             =   600
      Width           =   168
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   2
      Left            =   1536
      TabIndex        =   4
      Top             =   600
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Z:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   3
      Left            =   3096
      TabIndex        =   3
      Top             =   600
      Width           =   168
   End
End
Attribute VB_Name = "frmGetMSData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Loading As Boolean

Private Sub Command1_Click()

Cancelling = True
Unload Me
'frmMain.cmdCancel_Click

End Sub

Private Sub Command2_Click()

cbdata = Clipboard.GetText
If Len(cbdata) = 0 Then
    MsgBox ("Copy data from notepad")
    Exit Sub
End If
For I = 1 To Len(cbdata)
    Select Case Asc(Mid(cbdata, I, 1))
        Case 48 To 57, Asc("-"), Asc("."), Asc(","), 13, 10
        
        Case Else
            MsgBox ("Invalid microscribe data")
            I = Len(cbdata) + 1
            Unload Me
            Exit Sub
    End Select
Next I

X = InStr(cbdata, ",")
Text1(0) = Left(cbdata, X - 1)
cbdata = Mid(cbdata, X + 1)
X = InStr(cbdata, ",")
Text1(1) = Left(cbdata, X - 1)
cbdata = Mid(cbdata, X + 1)
Text1(2) = cbdata
edmshot.X = Val(Text1(0)) / 1000 - CurrentStation.X
edmshot.y = Val(Text1(1)) / 1000 - CurrentStation.y
edmshot.z = Val(Text1(2)) / 1000 - CurrentStation.z
Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

Text1(0) = Chr(KeyAscii)
KeyAscii = 0
Text1(0).SetFocus
Text1(0).SelStart = 1
Me.KeyPreview = False
    
End Sub

Private Sub Form_Load()

CenterForm Me

Screen.MousePointer = 1

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

Select Case KeyAscii
    Case 13
        edmshot.X = Val(Text1(0)) / 1000
        edmshot.y = Val(Text1(1)) / 1000
        edmshot.z = Val(Text1(2)) / 1000
        Beep
        Unload Me
    Case 48 To 57, Asc("-"), Asc(".")
    Case 44
        If Index < 2 Then
            KeyAscii = 0
            Text1(Index + 1).SetFocus
            'Text1(0).SelStart = 1
        End If
    Case 8
    Case Else
        KeyAscii = 0
        MsgBox ("Invalid data received from Microscribe")
End Select

End Sub


