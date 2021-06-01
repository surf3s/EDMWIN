VERSION 5.00
Begin VB.Form frmPrinter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Printer"
   ClientHeight    =   1320
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   4164
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4164
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Select Printer"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.OptionButton printtype 
      Caption         =   "Do not print"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2292
   End
   Begin VB.OptionButton printtype 
      Caption         =   "Print each recorded point"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2292
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   372
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Accept"
      Height          =   372
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1092
   End
End
Attribute VB_Name = "frmPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)

Select Case Index
Case 0
    PrinterOn = True
Case 1
    PrinterOn = False
Case Else
End Select

Dim IniClass As String
Dim Inidata(2, 2) As String
Dim Status As Byte

IniClass = "[EDM]"
Inidata(1, 1) = "AutoPrint"
Inidata(1, 2) = PrinterOn
Inidata(2, 1) = "Printer"
Inidata(2, 2) = comport
'Call WriteIni(CFGName, IniClass, Inidata(), Status)
Unload Me

End Sub

Private Sub Command2_Click()

frmMain.cd.ShowPrinter

End Sub

Private Sub Form_Load()

Call CenterForm(Me)
Screen.MousePointer = 1

End Sub
