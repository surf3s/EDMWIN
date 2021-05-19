VERSION 5.00
Begin VB.Form frmOffSet 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5100
   ControlBox      =   0   'False
   Icon            =   "frmOffSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   630
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton OffSetScale 
      Caption         =   "mm"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.OptionButton OffSetScale 
      Caption         =   "cm"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton OffSetScale 
      Caption         =   "m"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Offset by how much:"
      Height          =   252
      Left            =   144
      TabIndex        =   6
      Top             =   384
      Width           =   1476
   End
End
Attribute VB_Name = "frmOffSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CallingBox As Control
Public Varname As String
Public OriginalXYZ As Single

Private Sub Command1_Click()

Dim OffsetValue As Single
If Text1 <> "" And IsNumeric(Text1) Then
    OffsetValue = Val(Text1)
    Cancelling = False
    If OffSetScale(1) Then
        OffsetValue = OffsetValue / 100
    ElseIf OffSetScale(2) Then
        OffsetValue = OffsetValue / 1000
    End If
    ' OffSetScale(0) = True
    If Text1 > 1000 Then
        MsgBox ("Offsetting a point by more than a kilometer does not seem right -- if that's the case, then do it in increments")
        Exit Sub
    End If
    If OffsetValue = 0 Then
        frmMain.OffsetValue = OriginalXYZ + OffsetValue
    Else
        Select Case LCase(Me.Caption)
            Case "offset north", "offset east", "offset up"
                response = MsgBox("Adjust " & Varname & " from " & Format(OriginalXYZ, "####0.000") & " to " & Format(OriginalXYZ + Val(OffsetValue), "####0.000") & "?", vbYesNo)
                If response = vbYes Then
                    frmMain.OffsetValue = OriginalXYZ + Val(OffsetValue)
                Else
                    Text1.SetFocus
                    Exit Sub
                End If
            Case Else
                response = MsgBox("Adjust " & Varname & " from " & Format(OriginalXYZ, "####0.000") & " to " & Format(OriginalXYZ - Val(OffsetValue), "####0.000") & "?", vbYesNo)
                If response = vbYes Then
                    frmMain.OffsetValue = OriginalXYZ - Val(OffsetValue)
                Else
                    Text1.SetFocus
                    Exit Sub
                End If
    
        End Select
    End If
Else
    MsgBox ("Offset value must be numeric")
    Exit Sub
End If

Me.Hide

End Sub

Private Sub Command2_Click()

Cancelling = True
Text1 = 0
Command1_Click
'Me.Hide
'frmMain.txtXYZ(Index) = Format(OriginalXYZ, "####0.000")
'frmMain.Picture1.SetFocus

End Sub

Private Sub Form_Activate()

Text1.SetFocus

End Sub

Private Sub Form_Load()

Me.Height = 1230
Me.Width = 5190
CenterForm Me

End Sub

Private Sub Text1_GotFocus()

Text1.SelStart = 0
Text1.SelLength = Len(Text1)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 8, 48 To 57, Asc("-"), Asc(".")
    Case Else
        KeyAscii = 0
End Select

End Sub


