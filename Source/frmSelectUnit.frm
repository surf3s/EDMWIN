VERSION 5.00
Begin VB.Form frmSelectUnit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Unit"
   ClientHeight    =   2505
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkdefault 
      Alignment       =   1  'Right Justify
      Caption         =   "Set as default"
      Height          =   255
      Left            =   375
      TabIndex        =   6
      Top             =   1785
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   2256
      TabIndex        =   5
      Top             =   1824
      Width           =   1284
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Define New"
      Height          =   420
      Left            =   2256
      TabIndex        =   4
      Top             =   1344
      Width           =   1284
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   2256
      TabIndex        =   3
      Top             =   864
      Width           =   1284
   End
   Begin VB.ComboBox txtUnit 
      Height          =   315
      Left            =   192
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1185
      Width           =   1668
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Select Unit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   195
      TabIndex        =   2
      Top             =   975
      Width           =   1620
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Point is outside of defined Units with Limits.  Select Unit from list below (which are Units without Limits), or define new Unit."
      Height          =   564
      Left            =   72
      TabIndex        =   1
      Top             =   96
      Width           =   3564
   End
End
Attribute VB_Name = "frmSelectUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UnitSelected As String
Public Cancelling As Boolean
Public DefaultUnit As String

Private Sub Command1_Click()

Cancelling = False
If txtUnit.ListIndex <> -1 Then
    UnitSelected = txtUnit
    If chkdefault = 1 Then
        DefaultUnit = txtUnit
    Else
        DefaultUnit = ""
    End If
    Unload Me
Else
    UnitSelected = ""
    MsgBox ("Select Unit, Define a new Unit, or Cancel")
    Exit Sub
End If
Screen.MousePointer = 11

End Sub

Private Sub Command2_Click()

Cancelling = False
Unload Me

End Sub

Private Sub Command3_Click()

Cancelling = True
Screen.MousePointer = 11
Unload Me

End Sub

Private Sub Form_Load()

CenterForm Me
UnitSelected = ""
Screen.MousePointer = 1

End Sub


