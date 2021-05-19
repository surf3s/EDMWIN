VERSION 5.00
Begin VB.Form frmSubUnits 
   Caption         =   "Context-Dependent Defaults"
   ClientHeight    =   3360
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   4740
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   372
      Left            =   3840
      TabIndex        =   7
      Top             =   720
      Width           =   732
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3840
      TabIndex        =   5
      Top             =   1200
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   372
      Left            =   3840
      TabIndex        =   4
      Top             =   240
      Width           =   732
   End
   Begin VB.ComboBox cmbSubUnitMaster 
      Height          =   288
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   480
      Width           =   2412
   End
   Begin VB.ListBox lstSubUnitSlave 
      Height          =   2085
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1080
      Width           =   2412
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   $"frmSubUnits.frx":0000
      Height          =   1332
      Left            =   2520
      TabIndex        =   6
      Top             =   1920
      Width           =   2172
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Master Variable "
      Height          =   192
      Left            =   738
      TabIndex        =   3
      Top             =   240
      Width           =   1176
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Dependent Variables"
      Height          =   192
      Left            =   564
      TabIndex        =   1
      Top             =   840
      Width           =   1536
   End
End
Attribute VB_Name = "frmSubUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbSubUnitMaster_Click()

lstSubUnitSlave.Clear
For I = 1 To Vars
    Select Case UCase(VarList(I))
        Case "UNIT", "ID", "PRISM", "SUFFIX", "X", "Y", "Z", "VANGLE", "HANGLE", "SLOPED", "DATE", "TIME"
        Case UCase(cmbSubUnitMaster.Text)
        Case Else
            lstSubUnitSlave.AddItem VarList(I)
    End Select
Next I
For I = 1 To nDependentVars
    For J = 0 To lstSubUnitSlave.ListCount - 1
        If LCase(lstSubUnitSlave.List(J)) = LCase(DependentVar(I)) Then
            lstSubUnitSlave.Selected(J) = True
            Exit For
        End If
    Next J
Next I

End Sub

Private Sub Command1_Click()

MasterVar = cmbSubUnitMaster
nDependentVars = 0

For I = 0 To lstSubUnitSlave.ListCount - 1
    If lstSubUnitSlave.Selected(I) Then
        For J = 1 To nUnitFields
            If LCase(Unitfield(J)) = LCase(lstSubUnitSlave.List(I)) Then
                response = MsgBox("Make " + lstSubUnitSlave.List(I) + " dependent on " + MasterVar + " instead of Unit?", vbYesNo)
                If response = vbYes Then
                    AddUnits.Visible = False
                    AddUnits.lstUnitFields.Clear
                    For k = 1 To nUnitFields
                        If LCase(Unitfield(k)) <> LCase(lstSubUnitSlave.List(I)) Then
                            AddUnits.lstUnitFields.AddItem Unitfield(k)
                            AddUnits.lstUnitFields.Selected(AddUnits.lstUnitFields.NewIndex) = True
                            For l = 1 To Vars
                                If LCase(Unitfield(k)) = LCase(VarList(l)) Then
                                    AddUnits.lstUnitFields.ItemData(AddUnits.lstUnitFields.NewIndex) = I
                                    Exit For
                                End If
                            Next l
                        End If
                    Next k
                    Unload AddUnits
                End If
            End If
        Next J
    nDependentVars = nDependentVars + 1
        DependentVar(nDependentVars) = lstSubUnitSlave.List(I)
    End If
Next I

If nDependentVars = 0 Then
    response = MsgBox("No dependent variables have been checked. Do you want to clear this function?", vbYesNo)
    If response = vbYes Then
        Command3_Click
    Else
        Exit Sub
    End If
End If

On Error Resume Next
Set DefaultsTB = Nothing
SiteDB.TableDefs.Delete "EDM_Defaults"
On Error GoTo 0

CreateDefaultstb
Unload Me

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Command3_Click()

response = MsgBox("This will clear default values stored in the database.  Proceed anyway?", vbYesNo)
If response = vbYes Then
    MasterVar = ""
    nDependentVars = 0
    On Error Resume Next
    SiteDB.TableDefs.Delete "EDM_Defaults"
    On Error GoTo 0
    cmbSubUnitMaster.ListIndex = 0
    For I = 0 To lstSubUnitSlave.ListCount - 1
        lstSubUnitSlave.Selected(I) = False
    Next I
    frmMain.lblDefaults.Visible = False
    Set DefaultsTB = Nothing
    nDependentVars = 0
    MasterVar = ""
    Unload Me
End If

End Sub

Private Sub Form_Load()

Me.Height = 3900
Me.Width = 4956

cmbSubUnitMaster.Clear
For I = 1 To Vars
    Select Case UCase(VarList(I))
        Case "UNIT", "ID", "PRISM", "SUFFIX", "X", "Y", "Z", "VANGLE", "HANGLE", "SLOPED", "DATE", "TIME"
        Case Else
            cmbSubUnitMaster.AddItem VarList(I)
    End Select
Next I

If nDependentVars > 0 Then
    For I = 0 To cmbSubUnitMaster.ListCount - 1
        If LCase(MasterVar) = LCase(cmbSubUnitMaster.List(I)) Then
            cmbSubUnitMaster.ListIndex = I
            Exit For
        End If
    Next I
Else
    cmbSubUnitMaster.ListIndex = 0
End If

End Sub


