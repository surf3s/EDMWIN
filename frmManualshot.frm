VERSION 5.00
Begin VB.Form frmManualshot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Record point"
   ClientHeight    =   2490
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton varmenu 
      Caption         =   "M"
      Height          =   375
      Index           =   2
      Left            =   2640
      TabIndex        =   4
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox poleh 
      Alignment       =   1  'Right Justify
      Height          =   372
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   372
      Index           =   1
      Left            =   1800
      TabIndex        =   6
      Top             =   2040
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Accept"
      Height          =   372
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   1092
   End
   Begin VB.TextBox sloped 
      Alignment       =   1  'Right Justify
      Height          =   372
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   972
   End
   Begin VB.TextBox vangle 
      Alignment       =   1  'Right Justify
      Height          =   372
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   972
   End
   Begin VB.TextBox hangle 
      Alignment       =   1  'Right Justify
      Height          =   372
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   972
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Prism height :"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label instruct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "meters"
      Height          =   195
      Index           =   3
      Left            =   3120
      TabIndex        =   13
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label instruct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "meters"
      Height          =   192
      Index           =   2
      Left            =   2640
      TabIndex        =   12
      Top             =   1200
      Width           =   492
   End
   Begin VB.Label instruct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ddd.mmss"
      Height          =   192
      Index           =   1
      Left            =   2640
      TabIndex        =   11
      Top             =   720
      Width           =   756
   End
   Begin VB.Label instruct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ddd.mmss"
      Height          =   192
      Index           =   0
      Left            =   2640
      TabIndex        =   10
      Top             =   240
      Width           =   756
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Slope distance :"
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1332
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vertical angle :"
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Horizontal angle :"
      Height          =   252
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1332
   End
End
Attribute VB_Name = "frmManualshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

Dim edmpoffset As Single

Select Case Index
Case 0
    edmshot.vangle = vangle.Text
    edmshot.hangle = hangle.Text
    edmshot.sloped = sloped.Text
    
    'convert VHD to XYZ
    Call vhdtonez(edmshot)
    
    'offset shot from current station
    edmshot.X = edmshot.X + CurrentStation.X
    edmshot.Y = edmshot.Y + CurrentStation.Y
    edmshot.z = edmshot.z + CurrentStation.z
    
    With frmPointdata
        .X = Format(edmshot.X, "#########0.000")
        .Y = Format(edmshot.Y, "#########0.000")
        .z = Format(edmshot.z, "#########0.000")
        .vangle = Format(edmshot.vangle, "##0.0000")
        .hangle = Format(edmshot.hangle, "##0.0000")
        .sloped = Format(edmshot.sloped, "#########0.000")
        .poleh = Format(edmshot.poleh, "#########0.000")
        .Show 1
     End With
    Call insertpointintotb(edmshot)
    Me.Hide
    Me.Refresh
    frmEditrecord.Show 1
Case 1
Case Else
End Select

Unload Me

End Sub

Private Sub Form_Activate()

If Me.WindowState <> 0 Then hangle.SetFocus

End Sub

Private Sub Form_Load()

Label1.Top = hangle.Top + hangle.Height / 2 - Label1.Height / 2
Label2.Top = vangle.Top + vangle.Height / 2 - Label2.Height / 2
Label3.Top = sloped.Top + sloped.Height / 2 - Label3.Height / 2

instruct(0).Top = hangle.Top + hangle.Height / 2 - instruct(0).Height / 2
instruct(1).Top = vangle.Top + vangle.Height / 2 - instruct(1).Height / 2
instruct(2).Top = sloped.Top + sloped.Height / 2 - instruct(2).Height / 2

vhd = True

If vhd = True Then
    Label1.Caption = "Horizontal angle : "
    Label2.Caption = "Vertical angle : "
    Label3.Caption = "Slope distance : "
    instruct(0).Caption = "ddd.mmss"
    instruct(1).Caption = "ddd.mmss"
    instruct(2).Caption = "ddd.mmss"
Else
    Label1.Caption = "X coordinate : "
    Label2.Caption = "Y coordinate : "
    Label3.Caption = "Z coordinate : "
    instruct(0).Caption = "meters"
    instruct(1).Caption = "meters"
    instruct(2).Caption = "meters"
End If

End Sub

Private Sub hangle_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case 13
    KeyAscii = 0
    vangle.SetFocus
Case Else
End Select

End Sub

Private Sub poleh_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case 13
    KeyAscii = 0
    Command1(0).SetFocus
Case Else
End Select

End Sub

Private Sub vangle_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case 13
    KeyAscii = 0
    sloped.SetFocus
Case Else
End Select

End Sub

Private Sub sloped_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case 13
    KeyAscii = 0
    poleh.SetFocus
Case Else
End Select

End Sub

Private Sub varmenu_Click(Index As Integer)

If Not PoleTB.EOF Or Not PoleTB.BOF Then
    PoleTB.MoveFirst
    Do Until PoleTB.EOF
        frmMenu.menulist.AddItem PoleTB("Name") + " " + Str$(PoleTB("height"))
        PoleTB.MoveNext
    Loop
    frmMenu.Caption = "Pole menu"
    frmMenu.menutitle = "Select from the following :"
    frmMenu.Show 1
    If MenuSelection$ <> "" Then
        A = InStr(MenuSelection$, " ")
        If A <> 0 Then
            poleh.Text = Mid$(MenuSelection$, A + 1)
        End If
    End If
End If

End Sub
