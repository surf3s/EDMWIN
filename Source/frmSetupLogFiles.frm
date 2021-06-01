VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetupLogFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup Log Files"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClearTSLog 
      Caption         =   "Clear file"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox chkTSLog 
      Caption         =   "Log total station communications"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtTSLog 
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   6495
   End
   Begin VB.CommandButton cmdClearGeneralLog 
      Caption         =   "Clear file"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox chkGeneralLog 
      Caption         =   "Log setups and other program options"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox txtGeneralLog 
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   6495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(double click in the box to open a file browser)"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   1395
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(double click in the box to open a file browser)"
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   320
      Width           =   3495
   End
End
Attribute VB_Name = "frmSetupLogFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkTSLog_Click()

If chkTSLog Then
    txtTSLog.Enabled = True
    cmdClearTSLog.Enabled = True
Else
    txtTSLog.Enabled = False
    cmdClearTSLog.Enabled = False
End If

End Sub

Private Sub Form_Load()

chkTSLog.Value = Int(TSLog) * -1
chkGeneralLog.Value = Int(GeneralLog) * -1
txtTSLog.Text = TSLogFile
txtGeneralLog.Text = GeneralLogFile

End Sub

Private Sub Form_Unload(Cancel As Integer)

TSLog = chkTSLog.Value
TSLogFile = txtTSLog.Text
GeneralLog = chkGeneralLog.Value
GeneralLogFile = txtGeneralLog.Text
inifile$ = fixpath(App.Path) + "edm.ini"
Call WriteEDMIni(inifile$)

End Sub

Private Sub txtGeneralLog_DblClick()

On Error Resume Next
cd.CancelError = True
cd.ShowSave
If Err.Number = 0 Then
    txtGeneralLog.Text = cd.filename
End If

End Sub

Private Sub txtTSLog_DblClick()

On Error Resume Next
cd.CancelError = True
cd.ShowSave
If Err.Number = 0 Then
    txtTSLog.Text = cd.filename
End If

End Sub
