VERSION 5.00
Begin VB.Form frmHorizAngle 
   Caption         =   "Set Horizontal Angle"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmHorizAngle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Top             =   1290
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Angle"
      Default         =   -1  'True
      Height          =   285
      Left            =   750
      TabIndex        =   2
      Top             =   1290
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   810
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter Horizontal Angle (DDD.MMSS)"
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
      Left            =   570
      TabIndex        =   0
      Top             =   210
      Width           =   3750
   End
End
Attribute VB_Name = "frmHorizAngle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

deg = 0
min = 0
sec = 0

If Text1 = "" Then
    MsgBox ("Enter horizontal angle")
    Exit Sub
End If
If Text1 = "0" Then Text1 = "0.0000"
If Not IsNumeric(Text1) Then
    MsgBox ("Enter angle as numbers: Deg.minsec")
    Exit Sub
End If
    
If LCase(EDMName) <> "simulate" Then
    Screen.MousePointer = 11
    Call sethortangle(Text1, deg, min, sec)
    Screen.MousePointer = 1
End If
MsgBox ("Angle Set")
Unload Me

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Form_Load()

CenterForm Me

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 8, 48 To 57, Asc(".")
    Case Else
        KeyAscii = 0
End Select

End Sub

Private Sub Text1_LostFocus()

If Trim(Text1) = "" Then Exit Sub

X = InStr(Text1, ".")
If X = 0 Then
    Text1 = Text1 + ".0000"
End If
If Not IsNumeric(Text1) Then GoTo BadAngle
Exit Sub

BadAngle:
MsgBox ("Enter Horizontal Angle as DDD.MMSS")
Text1.SetFocus
    
End Sub


