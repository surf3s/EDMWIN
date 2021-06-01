VERSION 5.00
Begin VB.Form frmButtons 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Buttons"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8040
   ControlBox      =   0   'False
   Icon            =   "frmButtons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox txtShortcut 
      Height          =   315
      Left            =   6255
      TabIndex        =   20
      Top             =   3870
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Button"
      Height          =   435
      Left            =   6672
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   600
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   6672
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1110
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   435
      Left            =   6672
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   90
      Width           =   1155
   End
   Begin VB.ComboBox MenuBox 
      Height          =   315
      Left            =   6255
      TabIndex        =   5
      Top             =   4260
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Button 
      Caption         =   "Setup"
      Height          =   345
      Index           =   5
      Left            =   132
      TabIndex        =   13
      Top             =   4204
      Width           =   1185
   End
   Begin VB.TextBox NumberBox 
      Height          =   285
      Left            =   6255
      TabIndex        =   6
      Top             =   4260
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TextBox 
      Height          =   285
      Left            =   6255
      TabIndex        =   7
      Top             =   4260
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton optID 
      Caption         =   "Alpha"
      Height          =   255
      Index           =   1
      Left            =   7065
      TabIndex        =   9
      Top             =   4665
      Visible         =   0   'False
      Width           =   804
   End
   Begin VB.OptionButton optID 
      Caption         =   "Numeric"
      Height          =   255
      Index           =   0
      Left            =   6090
      TabIndex        =   8
      Top             =   4665
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtCaption 
      Height          =   285
      Left            =   6240
      TabIndex        =   4
      Top             =   3528
      Width           =   1335
   End
   Begin VB.ComboBox AddField 
      Height          =   315
      Left            =   6240
      TabIndex        =   1
      Text            =   "AddField"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ComboBox EditField 
      Height          =   315
      Left            =   6240
      TabIndex        =   2
      Text            =   "EditField"
      Top             =   2700
      Width           =   1335
   End
   Begin VB.ComboBox DeleteField 
      Height          =   315
      Left            =   6240
      TabIndex        =   3
      Text            =   "DeleteField"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Button 
      Caption         =   "Setup"
      Height          =   345
      Index           =   1
      Left            =   132
      TabIndex        =   0
      Top             =   2316
      Width           =   1185
   End
   Begin VB.CommandButton Button 
      Caption         =   "Setup"
      Height          =   345
      Index           =   2
      Left            =   132
      TabIndex        =   10
      Top             =   2788
      Width           =   1185
   End
   Begin VB.CommandButton Button 
      Caption         =   "Setup"
      Height          =   345
      Index           =   3
      Left            =   132
      TabIndex        =   11
      Top             =   3260
      Width           =   1185
   End
   Begin VB.CommandButton Button 
      Caption         =   "Setup"
      Height          =   345
      Index           =   4
      Left            =   132
      TabIndex        =   12
      Top             =   3732
      Width           =   1185
   End
   Begin VB.CommandButton Button 
      Caption         =   "Setup"
      Height          =   345
      Index           =   6
      Left            =   132
      TabIndex        =   14
      Top             =   4680
      Width           =   1185
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Shortcut:  ctrl-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4830
      TabIndex        =   28
      Top             =   3930
      Width           =   1245
   End
   Begin VB.Label Label5 
      Caption         =   $"frmButtons.frx":000C
      Height          =   432
      Left            =   96
      TabIndex        =   27
      Top             =   1536
      Width           =   6444
   End
   Begin VB.Label Label4 
      Caption         =   $"frmButtons.frx":00B1
      Height          =   648
      Left            =   96
      TabIndex        =   26
      Top             =   852
      Width           =   6396
   End
   Begin VB.Label Label3 
      Caption         =   $"frmButtons.frx":01C8
      Height          =   432
      Left            =   96
      TabIndex        =   25
      Top             =   384
      Width           =   6372
   End
   Begin VB.Label Label2 
      Caption         =   "Buttons allow you to predefine shot types that will automatically fill in values for fields."
      Height          =   252
      Left            =   96
      TabIndex        =   24
      Top             =   96
      Width           =   6324
   End
   Begin VB.Label lblButton 
      AutoSize        =   -1  'True
      Height          =   192
      Left            =   2616
      TabIndex        =   23
      Top             =   2052
      Width           =   48
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current Button:"
      Height          =   192
      Left            =   1380
      TabIndex        =   22
      Top             =   2052
      Width           =   1068
   End
   Begin VB.Label lblvalue 
      AutoSize        =   -1  'True
      Caption         =   "Value:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4830
      TabIndex        =   21
      Top             =   4335
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "ID Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4830
      TabIndex        =   19
      Top             =   4710
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Caption:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4830
      TabIndex        =   18
      Top             =   3585
      Width           =   720
   End
End
Attribute VB_Name = "frmButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentButton As Integer
Dim Adding As Boolean
Dim Editing As Boolean

Private Sub AddField_Click()

Adding = True
MenuBox.Visible = False
TextBox.Visible = False
NumberBox.Visible = False
lblvalue.Visible = False
lblID.Visible = False
optID(0).Visible = False
optID(1).Visible = False

Grid.Rows = Grid.Rows + 1
Grid.FixedRows = 1
Grid.Row = Grid.Rows - 1
Grid.Col = 1
Grid = AddField
nButtonVars(CurrentButton) = Grid.Rows - 2
ButtonVars(CurrentButton, nButtonVars(CurrentButton), 1) = I

Grid.Col = 2
For I = 1 To Vars
    If LCase(VarList(I)) = LCase(AddField) Then
        If LCase(AddField) = "id" Then
            lblID.Visible = True
            optID(0).Visible = True
            optID(1).Visible = True
        ElseIf LCase(AddField) = "unit" Then
            MenuBox.Clear
            For J = 0 To frmMain.txtUnit.ListCount - 1
                MenuBox.AddItem frmMain.txtUnit.List(J)
            Next J
            MenuBox.Visible = True
            MenuBox.SetFocus
            lblvalue.Visible = True
        ElseIf LCase(AddField) = "prism" Then
            MenuBox.Clear
            For J = 0 To frmMain.txtprism.ListCount - 1
                MenuBox.AddItem frmMain.txtprism.List(J)
            Next J
            MenuBox.Visible = True
            MenuBox.SetFocus
            lblvalue.Visible = True
        Else
            lblvalue.Visible = True
            Select Case LCase(VType(I))
                Case "menu"
                    MenuBox.Visible = True
                    MenuBox.Clear
                        MenuString = LCase(VMenu(I))
                        Gotit = False
                        Do Until Gotit
                            X = InStr(MenuString, ",")
                            If X > 0 Then
                                MenuBox.AddItem Left(MenuString, X - 1)
                                MenuString = Mid(MenuString, X + 1)
                            Else
                                MenuBox.AddItem MenuString
                                Gotit = True
                            End If
                        Loop
                        MenuBox = ""
                        MenuBox.SetFocus
                Case "text"
                    TextBox.Top = MenuBox.Top
                    TextBox = ""
                    TextBox.Visible = True
                    TextBox.SetFocus
                Case "numeric", "instrument"
                    NumberBox.Top = MenuBox.Top
                    NumberBox = ""
                    NumberBox.Visible = True
                    NumberBox.SetFocus
           End Select
        End If
    End If
Next I

End Sub

Private Sub AddField_DropDown()

AddField.Clear
For J = 1 To Vars
    Gotit = False
    For I = 1 To nButtonVars(Index)
        If LCase(VarList(J)) = LCase(VarList(ButtonVars(Index, I, 1))) Then
            Gotit = True
            Exit For
        End If
    Next I
    If Not Gotit Then
        If Not LCase(VarList(J)) = "x" And Not LCase(VarList(J)) = "y" And Not LCase(VarList(J)) = "z" And Not LCase(VarList(J)) = "suffix" And Not LCase(VarList(J)) = "vangle" And Not LCase(VarList(J)) = "hangle" And Not LCase(VarList(J)) = "sloped" And Not LCase(VarList(J)) = "suffix" Then
            AddField.AddItem LCase(VarList(J))
        End If
    End If
Next J

End Sub

Private Sub AddField_LostFocus()

AddField = "Add Field"

End Sub

Public Sub Button_Click(Index As Integer)

Dim I, J As Integer

If Not Loading Then
    SaveButton
End If

Grid.Rows = nButtonVars(Index) + 1
For I = 1 To nButtonVars(Index)
    Grid.Row = I
    Grid.Col = 1
    Grid.Text = LCase(VarList(ButtonVars(Index, I, 1)))
    Grid.Col = 2
    Grid.Text = LCase(ButtonVars(Index, I, 2))
Next I
lblButton = Button(Index).Caption
txtShortcut = ButtonShortCut(Index)
txtShortcut.Refresh
optID(0) = False
optID(1) = False

CurrentButton = Index
If Button(Index).Caption = "Setup" Then
    txtCaption = "Button" + Trim(Str(Index))
Else
    txtCaption = Button(Index).Caption
End If
        
End Sub

Private Sub Command1_Click()

Dim I As Integer
Dim J, X As Integer
Dim IniClass As String
Dim Inidata() As String
Dim Status As Byte
SaveButton
ClearButtonIni

For I = 1 To 6
    If nButtonVars(I) > 0 Then
        Loading = True
        Button_Click I
        IniClass = "[BUTTON" + Trim(Str(I)) + "]"
        X = Grid.Rows
        ReDim Inidata(X + 1, 2) As String
    
        Inidata(1, 1) = "Title"
        Inidata(1, 2) = txtCaption
        Inidata(2, 1) = "Shortcut"
        Inidata(2, 2) = ButtonShortCut(I)
        For J = 1 To Grid.Rows - 1
            Grid.Row = J
            Grid.Col = 1
            Inidata(J + 2, 1) = Grid
            Grid.Col = 2
            Inidata(J + 2, 2) = Grid
        Next J
        Call WriteIni(CFGName, IniClass, Inidata(), Status)
    End If
Next I
Unload Me
parsecfg A

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Command3_Click()

nButtonVars(CurrentButton) = 0
Grid.Rows = 1
txtCaption = ""
txtShortcut.Clear
Button(CurrentButton).Caption = "Setup"
ButtonShortCut(CurrentButton) = ""
AddField.Clear
EditField.Clear
DeleteField.Clear
AddField = "Add Field"
EditField = "Edit Field"
DeleteField = "Delete Field"

End Sub

Private Sub DeleteField_Click()

Dim CurrentField As Integer
Dim TempString1, TempString2 As String

CurrentField = DeleteField.ListIndex + 1
For I = CurrentField To Grid.Rows - 2
    Grid.Row = I + 1
    Grid.Col = 1
    TempString1 = Grid
    Grid.Col = 2
    TempString2 = Grid
    Grid.Row = I
    
    Grid = TempString2
    Grid.Col = 1
    Grid = TempString1
Next I
Grid.Rows = Grid.Rows - 1
DeleteField = "Delete Field"
SaveButton

End Sub

Private Sub DeleteField_DropDown()

DeleteField.Clear
For I = 1 To nButtonVars(CurrentButton)
    DeleteField.AddItem LCase(VarList(ButtonVars(CurrentButton, I, 1)))
Next I

End Sub

Private Sub DeleteField_LostFocus()

DeleteField = "Delete Field"

End Sub

Private Sub EditField_Click()

MenuBox.Visible = False
TextBox.Visible = False
NumberBox.Visible = False
lblvalue.Visible = False
lblID.Visible = False
optID(0).Visible = False
optID(1).Visible = False

Grid.Row = EditField.ListIndex + 1
Grid.Col = 2

If LCase(EditField) = "id" Then
    lblID.Visible = True
    optID(0).Visible = True
    optID(1).Visible = True
    Loading = True
    If LCase(Grid) = "alpha" Then
        optID(1) = True
    Else
        optID(0) = True
    End If
    Loading = False
    Editing = True
    
Else
    For I = 1 To Vars
        If LCase(EditField) = LCase(VarList(I)) Then
            Exit For
        End If
    Next I
    lblvalue.Visible = True
    Select Case LCase(VType(I))
        Case "menu"
            MenuBox.Visible = True
            MenuBox.Clear
                MenuString = LCase(VMenu(I))
                Gotit = False
                Do Until Gotit
                    X = InStr(MenuString, ",")
                    If X > 0 Then
                        MenuBox.AddItem Left(MenuString, X - 1)
                        MenuString = Mid(MenuString, X + 1)
                    Else
                        MenuBox.AddItem MenuString
                        Gotit = True
                    End If
                Loop
                MenuBox = Grid
                MenuBox.SetFocus
        Case "text"
            TextBox.Top = MenuBox.Top
            TextBox = Grid
            TextBox.Visible = True
            TextBox.SetFocus
        Case "numeric", "instrument"
            NumberBox.Top = MenuBox.Top
            NumberBox = Grid
            NumberBox.Visible = True
            NumberBox.SetFocus
   End Select
   Editing = True
End If

End Sub

Private Sub EditField_DropDown()

EditField.Clear
For I = 1 To nButtonVars(CurrentButton)
    EditField.AddItem LCase(VarList(ButtonVars(CurrentButton, I, 1)))
Next I

End Sub

Private Sub EditField_LostFocus()

EditField = "Edit Field"

End Sub

Private Sub Form_Activate()

Button(1).SetFocus

End Sub

Private Sub Form_Load()

Dim I As Integer

Grid.ColWidth(0) = 1
Grid.ColWidth(1) = 1600
Grid.ColWidth(2) = 1600
Grid.Row = 0
Grid.Col = 1
Grid.Text = "Field"
Grid.Col = 2
Grid.Text = "Value"
Grid.Rows = 1
'Grid.SelStartCol = 0
Grid.SelEndCol = 2
Grid.SelStartRow = 0
Grid.SelEndRow = 0

For I = 1 To 6
    If frmMain.Button(I).Visible Then
        Button(I).Caption = ButtonCaption(I)
    End If
Next I
AddField = "Add Field"
EditField = "Edit Field"
DeleteField = "Delete Field"
'Button_Click 1

CenterForm Me
CurrentButton = 1
For I = 1 To 6
    If Button(I).Caption <> "Setup" Then
        Loading = True
        Button_Click I
        Loading = False
        Exit Sub
    End If
Next I

End Sub

Private Sub Form_Unload(Cancel As Integer)

Me.Hide
frmMain.Picture1.SetFocus

End Sub

Private Sub MenuBox_Click()

Grid.Col = 2
Grid = MenuBox
ButtonVars(CurrentButton, nButtonVars(CurrentButton), 2) = Grid

If Adding Then
    AddField.Clear
    AddField = "Add Field"
    Adding = False
    lblvalue.Visible = False
    MenuBox.Visible = False
ElseIf Editing Then
    EditField = "Edit Field"
    Editing = False
    lblvalue.Visible = False
    MenuBox.Visible = False
End If
SaveButton

End Sub

Private Sub MenuBox_KeyPress(KeyAscii As Integer)

If UpperCase Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub MenuBox_LostFocus()

Static Warned As Boolean

If MenuBox = "" And Not Warned Then
    MsgBox ("Enter default value for " + AddField)
    Warned = True
End If
    
End Sub

Private Sub NumberBox_Change()

If Not IsNumeric(NumberBox) Then
    Beep
    Exit Sub
End If
Grid.Col = 2
Grid = MenuBox
If Adding Then
    AddField = "Add Field"
    Adding = False
ElseIf Editing Then
    EditField = "Edit Field"
    Editing = False
End If

End Sub

Private Sub NumberBox_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Adding Then
        AddField = "Add Field"
        Adding = False
        NumberBox.Visible = False
        lblvalue.Visible = False
    ElseIf Editing Then
        EditField = "Edit Field"
        Editing = False
        NumberBox.Visible = False
        lblvalue.Visible = False
    End If
Else
    Select Case KeyAscii
        Case 8, 46, 48 To 57, Asc("-"), Asc(".")
        Case Else
            KeyAscii = 0
    End Select
End If

End Sub

Private Sub NumberBox_LostFocus()

Static Warned As Boolean

If Not NumberBox = "" And Not Warned Then
    MsgBox ("Enter default value for " + AddField)
    Warned = True
Else
    If Adding Then
        AddField = "Add Field"
        Adding = False
        NumberBox.Visible = False
        lblvalue.Visible = False
    ElseIf Editing Then
        EditField = "Edit Field"
        Editing = False
        NumberBox.Visible = False
        lblvalue.Visible = False
    End If
    
End If

End Sub

Private Sub optID_Click(Index As Integer)

Grid.Col = 2
If Index = 0 Then
    Grid.Text = "Numeric"
ElseIf Index = 1 Then
    Grid.Text = "Alpha"
End If

If Adding Then
    AddField = "Add Field"
    Adding = False
    optID(0).Visible = False
    optID(1).Visible = False
    lblID.Visible = False
ElseIf Editing Then
    EditField = "Edit Field"
    Editing = False
    optID(0).Visible = False
    optID(1).Visible = False
    lblID.Visible = False
End If
If Not Loading Then SaveButton

End Sub

Private Sub TextBox_Change()

Grid.Col = 2
Grid = TextBox

End Sub

Private Sub TextBox_KeyPress(KeyAscii As Integer)

If UpperCase Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub TextBox_LostFocus()

Static Warned As Boolean

If Not TextBox = "" And Not Warned Then
    MsgBox ("Enter default value for " + AddField)
    Warned = True
Else
    If Adding Then
        AddField = "Add Field"
        Adding = False
        TextBox.Visible = False
        lblvalue.Visible = False
    ElseIf Editing Then
        EditField = "Edit Field"
        Editing = False
        TextBox.Visible = False
        lblvalue.Visible = False
    End If
    ButtonVars(CurrentButton, nButtonVars(CurrentButton), 2) = Grid
End If

End Sub

Private Sub txtCaption_Change()

If txtCaption <> "" Then
    ButtonCaption(CurrentButton) = txtCaption
    Button(CurrentButton).Caption = txtCaption
    lblButton = txtCaption
End If

End Sub

Public Sub SaveButton()

nButtonVars(CurrentButton) = Grid.Rows - 1
For I = 1 To nButtonVars(CurrentButton)
    Grid.Row = I
    Grid.Col = 1
    For J = 1 To Vars
        If LCase(VarList(J)) = LCase(Grid.Text) Then
            ButtonVars(CurrentButton, I, 1) = J
            Exit For
        End If
    Next J
    Grid.Col = 2
    ButtonVars(CurrentButton, I, 2) = Grid.Text
Next I

ButtonCaption(CurrentButton) = Button(CurrentButton).Caption
If Trim(txtShortcut) = "" Or LCase(txtShortcut) = "none" Then
    ButtonShortCut(CurrentButton) = ""
Else
    ButtonShortCut(CurrentButton) = txtShortcut
End If

End Sub

Private Sub txtCaption_KeyPress(KeyAscii As Integer)

If UpperCase Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub txtShortcut_Click()

ButtonShortCut(CurrentButton) = ""
For I = 1 To 6
    If txtShortcut = ButtonShortCut(I) Then
        MsgBox ("This shortcut conflicts with another one already defined.")
        Exit For
    End If
Next I
SaveButton

End Sub

Private Sub txtShortcut_DropDown()

txtShortcut.Clear
txtShortcut.AddItem "None"
For I = 65 To 90
    If I <> 71 And I <> 80 Then
        txtShortcut.AddItem Chr(I)
    End If
Next I

End Sub

