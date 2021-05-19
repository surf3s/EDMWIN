VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCEupdownload 
   Caption         =   "Transfer between PC and Pocket PC"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   9855
   Begin VB.CommandButton Command2 
      Caption         =   "Initialize communication with Pocket PC"
      Height          =   345
      Left            =   2430
      TabIndex        =   0
      Top             =   570
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   8670
      TabIndex        =   9
      Top             =   240
      Width           =   885
   End
   Begin VB.Timer timClearStatus 
      Interval        =   5000
      Left            =   7200
      Top             =   4440
   End
   Begin VB.Frame Frame2 
      Caption         =   "CFG File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   510
      TabIndex        =   14
      Top             =   2910
      Width           =   8535
      Begin VB.TextBox txtPPCFile 
         Height          =   285
         Left            =   1110
         TabIndex        =   6
         Top             =   750
         Width           =   6645
      End
      Begin VB.TextBox txtDesktopFile 
         Height          =   285
         Left            =   1110
         TabIndex        =   5
         Top             =   390
         Width           =   6645
      End
      Begin VB.CommandButton cmduploadfile 
         Caption         =   "Copy from Pocket PC"
         Height          =   375
         Left            =   5550
         TabIndex        =   8
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton cmdDownloadFile 
         Caption         =   "Copy to Pocket PC"
         Height          =   375
         Left            =   1110
         TabIndex        =   7
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton cmdFileSelect 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   7830
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "PPC File"
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Desktop File"
         Height          =   255
         Left            =   150
         TabIndex        =   16
         Top             =   390
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Points Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   510
      TabIndex        =   11
      Top             =   1020
      Width           =   8535
      Begin VB.CommandButton cmdFileSelect 
         Caption         =   "..."
         Height          =   255
         Index           =   1
         Left            =   7890
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   450
         Width           =   255
      End
      Begin VB.CommandButton cmdDownload 
         Caption         =   "Copy to Pocket PC"
         Height          =   375
         Left            =   1170
         TabIndex        =   3
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Copy from Pocket PC"
         Height          =   375
         Left            =   5610
         TabIndex        =   4
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtDesktopDB 
         Height          =   285
         Left            =   1170
         TabIndex        =   1
         Top             =   450
         Width           =   6615
      End
      Begin VB.TextBox txtPPCDB 
         Height          =   285
         Left            =   1170
         TabIndex        =   2
         Top             =   810
         Width           =   6615
      End
      Begin VB.Label Label2 
         Caption         =   "Desktop DB"
         Height          =   255
         Left            =   210
         TabIndex        =   13
         Top             =   450
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "PPC DB"
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   810
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog cdlgFiles 
      Left            =   90
      Top             =   4380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2370
      TabIndex        =   10
      Top             =   4770
      Width           =   4695
   End
End
Attribute VB_Name = "frmCEupdownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim intRetVal As Long
Dim ChangingPath As Boolean
Dim CommInit As Boolean
Dim inifile As String
Private Sub cmdDownload_Click()
    If Not CommInit Then
        MsgBox ("First Initialize communication with the Pocket PC")
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    'Pretty the process up a bit by playing the filemove
    Label1.Caption = "Downloading Data To Pocket PC"
    Screen.MousePointer = 11
    DoEvents
    'download from the desktop to the local DB
    intRetVal = DESKTOPTODEVICE(txtDesktopDB.Text, _
        "", False, True, txtPPCDB.Text)
    
    
    If intRetVal <> 0 Then
        Label1.Caption = "Error " & intRetVal & " Occurred."
    Else
        Label1.Caption = "Download Successful"
    End If
    Screen.MousePointer = 1
    Exit Sub

ErrHandler:
    MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical
End Sub

Private Sub cmdDownloadFile_Click()
    If Not CommInit Then
        MsgBox ("First Initialize communication with the Pocket PC")
        Exit Sub
    End If
    
    Label1.Caption = "Downloading CFG File To Pocket PC"
    Screen.MousePointer = 11
    DoEvents
    If CopyFileToPocketPC = False Then
        Label1.Caption = "An error occurred transferring the file"
    Else
        Label1.Caption = "File copied successfully."
    End If
    Screen.MousePointer = 1
End Sub

Private Sub cmdFileSelect_Click(Index As Integer)
    cdlgFiles.ShowOpen
    If cdlgFiles.CancelError = False Then
        If Index = 0 Then
            txtDesktopFile = cdlgFiles.filename
        Else
            txtdesktoppc = cdlgFiles.FileTitle
        End If
    End If

End Sub

Private Sub cmdUpload_Click()
    
    If Not CommInit Then
        MsgBox ("First Initialize communication with the Pocket PC")
        Exit Sub
    End If
    
    Label1.Caption = "Uploading Data from Pocket PC"
    Screen.MousePointer = 11
    DoEvents
    'upload from the local DB to the desktop DB
    If Dir(txtDesktopDB) <> "" Then
        Kill txtDesktopDB
    End If
    intRetVal = DEVICETODESKTOP(txtDesktopDB.Text, _
    "", False, True, txtPPCDB.Text)
    If intRetVal <> 0 Then
        Label1.Caption = "An error occurred transferring the data"
    Else
        Label1.Caption = "Upload Successful"
    End If
    Screen.MousePointer = 1
End Sub

Private Sub cmduploadfile_Click()
    If Not CommInit Then
        MsgBox ("First Initialize communication with the Pocket PC")
        Exit Sub
    End If

    
    Label1.Caption = "Uploading CFG File From Pocket PC"
    Screen.MousePointer = 11
    DoEvents
    If CopyFileFromPocketPC = False Then
        Label1.Caption = "An error occurred transferring the file"
    Else
        Label1.Caption = "File copied successfully."
    End If
    Screen.MousePointer = 1
End Sub


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
MsgBox ("Insert Pocket PC into its cradle and make sure that it is ON")
    
intRetVal = CeRapiInit()
If intRetVal <> INIT_SUCCESS Then
    MsgBox "Could not initialize communication with the Pocket PC.  Verify that all cables are connected, that it is properly seated in its cradle, and that its power is on.  Then close this form, re-open it, and try again"
    CeRapiUninit
    Unload Me
End If

CommInit = True
Label1 = "Communication with Pocket PC initialized"
End Sub

Private Sub Form_Load()
Me.Height = 5685
Me.Width = 9975
txtDesktopDB = SiteDBname
txtDesktopFile = CFGName


txtPPCDB = PPPath + DBName
txtPPCFile = PPPath + CFGTitle

CenterForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CeRapiUninit
End Sub

Private Function CopyFileToPocketPC() As Boolean
    Dim A As Long
    'Please note, this example only handles files up to 16384 bytes in size.
    'this could relatively easily be re-written to read buffered chunks of a file
    'and make consecutive writes to the device, or even just set up a huge buffer
    'capable of reading any size file, and do the whole write at once.
    
    'i have left the example buffered for simplicity
    
    'Set up the binary buffer to read the file from the PC
    Dim bytBuffer(16384) As Byte
    
    Dim lngFileHandle As Long
    Dim lngDestinationHandle As Long
    Dim lngNumBytesWritten As Long
    
    Dim filename As String
    
    Dim typFindFileData As CE_FIND_DATA
    
    Dim lngFileSize As Long
    
    'On Error GoTo ErrHandler

    'locate the file, and see if it already exists on the device
    lngFileHandle = CeFindFirstFile(txtPPCFile.Text, typFindFileData)
    
    'if we get -1 then the file wasn't found so we can go ahead and create it
    'otherwise we dont want to overwrite in this demo, so we will cancel the operation.
    
'    If lngFileHandle <> INVALID_HANDLE Then
'        MsgBox txtPPCFile & " already exists. Operation Cancelled."
'        CeFindClose lngFileHandle
'        CopyFileToPocketPC = False
'    End If
    
    
    'create the file on the device, and return the handle to the file
    filename = txtPPCFile.Text
    A = CeDeleteFile(filename)
    lngDestinationHandle = CeCreateFile(filename, GENERIC_WRITE, FILE_SHARE_READ, vbNullString, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    
    'if something went wrong in CeCreateFile we will get -1, and therefore need to abort
    If lngDestinationHandle = INVALID_HANDLE Then
        MsgBox "CeCreateFile Error " & CeGetLastError & " occurred.", vbCritical & vbOKOnly
        CopyFileToPocketPC = False
        Exit Function
    End If
    
    'Call function to read file on the desktop into the byte buffer
    If ReadFileAsBinary(txtDesktopFile.Text, lngFileSize, bytBuffer()) = True Then
        
        'Write contents of Binary buffer to CE file.
        'Return value of 0 = failure, non zero = success
        intRetVal = CeWriteFile(lngDestinationHandle, bytBuffer(0), lngFileSize, lngNumBytesWritten, 0)
        
        If intRetVal = WRITE_ERROR Then
            GoTo ErrHandler
        End If
        
    End If
    
    CeCloseHandle lngDestinationHandle
    
    CopyFileToPocketPC = True

    Exit Function

ErrHandler:

    MsgBox "Error " & CeGetLastError & " Occurred When Copying File " & _
    txtDesktopFile.Text & " To The Device as " & txtPPCFile.Text, vbCritical & vbOKOnly
    
    If lngDestinationHandle Or lngDestinationHandle <> INVALID_HANDLE Then
        CeCloseHandle lngDestinationHandle
    End If
    CopyFileToPocketPC = False

End Function
Private Function CopyFileFromPocketPC() As Boolean

    Dim bytBuffer(16384) As Byte
    
    Dim lngFileHandle As Long
    Dim lngBytesRead As Long
    Dim typFindFileData As CE_FIND_DATA
    
    Dim intFreeFileID As Integer
    
    Dim intWriteLoop As Integer
    
    'locate the file, and see if it already exists on the device
    lngFileHandle = CeFindFirstFile(txtPPCFile.Text, typFindFileData)
    
    If lngFileHandle = INVALID_HANDLE Then
        MsgBox "File " & txtPPCFile.Text & " Not Found. Operation Aborted.", vbOKOnly
        CopyFileFromPocketPC = False
        Exit Function
    End If
    
    'we dont need this handle now that we know the file is there
    CeFindClose lngFileHandle
    
    'i know it seems odd to have to call a function called CreateFile to read a file
    'but what it really refers to is create me a handle to a file of this type.
    lngFileHandle = CeCreateFile(txtPPCFile.Text, GENERIC_READ, FILE_SHARE_READ, vbNullString, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    
    If lngFileHandle = INVALID_HANDLE Then
        MsgBox "Failed to open file " & txtPPCFile.Text
        CopyFileFromPocketPC = False
        Exit Function
    End If
    
    intRetVal = CeReadFile(lngFileHandle, bytBuffer(0), 16384, lngBytesRead, 0)
    
    'if we got a 0 return value from readfile then there is an error
    'so pass it to the error handler
    If intRetVal = READ_ERROR Then
        GoTo ErrHandler
    End If
    
    intFreeFileID = FreeFile
    
    Open txtDesktopFile.Text For Binary As intFreeFileID
    For intWriteLoop = 0 To lngBytesRead
        Put #intFreeFileID, intWriteLoop + 1, bytBuffer(intWriteLoop)
    Next intWriteLoop
        
    Close intFreeFileID
        
    CopyFileFromPocketPC = True

    Exit Function
    
ErrHandler:

    MsgBox "Error " & CeGetLastError & " Occurred When Copying File " & _
    txtPPCFile.Text & " From The Device as " & txtDesktopFile.Text, vbCritical & vbOKOnly
    CopyFileFromPocketPC = False

End Function

Private Sub timClearStatus_Timer()
'    Label1.Caption = ""
End Sub

Private Function ReadFileAsBinary(strSrcFilename As String, lngFileSize As Long, bytBuffer() As Byte) As Boolean

    Dim intFileHandle As Integer
    Dim intSeekPos As Integer
    
    'On Error GoTo ErrHandler
 
    lngFileSize = FileLen(strSrcFilename)

    intFileHandle = FreeFile

    Open strSrcFilename For Binary As intFileHandle

    For intSeekPos = 1 To lngFileSize
        Get #intFileHandle, intSeekPos, bytBuffer(intSeekPos - 1)
    Next intSeekPos
    
    Close intFileHandle

    ReadFileAsBinary = True

    Exit Function

ErrHandler:

    MsgBox "Error reading file " & strSrcFilename, vbCritical & vbOKOnly
    
    ReadFileAsBinary = False

End Function

Private Sub txtDesktopDB_Change()
ChangingPath = True
End Sub


Private Sub txtDesktopDB_GotFocus()
ChangingPath = False
End Sub


Private Sub txtDesktopDB_LostFocus()
If ChangingPath Then
    PCdbPath = txtDesktopDB
    inifile$ = fixpath(App.Path) + "edm.ini"
    Call WriteEDMIni(inifile$)
End If

End Sub

Private Sub txtDesktopFile_Change()
ChangingPath = True
End Sub

Private Sub txtDesktopFile_GotFocus()
ChangingPath = False
End Sub


Private Sub txtDesktopFile_LostFocus()
If ChangingPath Then
    PCcfgPath = txtDesktopFile
    inifile$ = fixpath(App.Path) + "edm.ini"
    Call WriteEDMIni(inifile$)
End If
End Sub

Private Sub txtPPCDB_Change()
ChangingPath = True

End Sub


Private Sub txtPPCDB_GotFocus()
ChangingPath = False
End Sub


Private Sub txtPPCDB_LostFocus()
If ChangingPath Then
    For I = Len(txtPPCDB) To 1 Step -1
        If Mid(txtPPCDB, I, 1) = "\" Then
            PPPath = Left(txtPPCDB, I)
            Exit For
        End If
    Next I
    Dim Inidata(100, 2) As String
    Dim IniClass As String
    Dim Status As Byte
    IniClass = "[EDM]"
    Inidata(1, 1) = "PPPath"
    Inidata(1, 2) = PPPath
    Call WriteIni(CFGName, IniClass, Inidata(), Status)

End If

End Sub

Private Sub txtPPCFile_Change()
ChangingPath = True

End Sub


Private Sub txtPPCFile_GotFocus()
ChangingPath = False
End Sub


Private Sub txtPPCFile_LostFocus()
If ChangingPath Then
    For I = Len(txtPPCFile) To 1 Step -1
        If Mid(txtPPCFile, I, 1) = "\" Then
            PPPath = Left(txtPPCFile, I)
            Exit For
        End If
    Next I
    Dim Inidata(100, 2) As String
    Dim IniClass As String
    Dim Status As Byte
    IniClass = "[EDM]"
    Inidata(1, 1) = "PPPath"
    Inidata(1, 2) = PPPath
    Call WriteIni(CFGName, IniClass, Inidata(), Status)

End If

End Sub


