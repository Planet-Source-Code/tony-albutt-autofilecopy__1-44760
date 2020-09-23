VERSION 5.00
Begin VB.Form frmAutoFileCopy 
   AutoRedraw      =   -1  'True
   Caption         =   "Auto File Copy"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   Icon            =   "frmAutoFileCopy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin AutoFileCopy.cpvProgressBar proSize 
      Height          =   225
      Left            =   1980
      Top             =   4980
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   397
      BarPicture      =   "frmAutoFileCopy.frx":030A
      BarPictureBack  =   "frmAutoFileCopy.frx":5294
      CaptionFormat   =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontColor       =   65535
   End
   Begin VB.ComboBox cboFileTypes 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3300
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   1140
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7800
      TabIndex        =   23
      Top             =   5340
      Width           =   915
   End
   Begin VB.TextBox txtSize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   720
      TabIndex        =   21
      Top             =   4980
      Width           =   1155
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "File Count"
      Height          =   375
      Left            =   4920
      TabIndex        =   20
      Top             =   5340
      Width           =   915
   End
   Begin VB.TextBox txtFiles 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   720
      TabIndex        =   19
      Top             =   4680
      Width           =   1155
   End
   Begin VB.TextBox txtFolders 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   720
      TabIndex        =   18
      Top             =   4380
      Width           =   1155
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   60
      Picture         =   "frmAutoFileCopy.frx":A21E
      ScaleHeight     =   525
      ScaleWidth      =   525
      TabIndex        =   15
      ToolTipText     =   "Auto File Copy"
      Top             =   60
      Width           =   555
   End
   Begin VB.CheckBox optCopySubFolder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Include sub folders"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1980
      TabIndex        =   13
      Top             =   5400
      Width           =   1635
   End
   Begin VB.ComboBox cboUpdateRate 
      Enabled         =   0   'False
      Height          =   315
      Left            =   720
      TabIndex        =   12
      Top             =   5400
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8280
      Top             =   780
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Un-Lock"
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Top             =   5340
      Width           =   915
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   555
      Left            =   7920
      TabIndex        =   10
      Text            =   "08:00"
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdRunNow 
      Caption         =   "Run Now"
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   5340
      Width           =   915
   End
   Begin VB.FileListBox File2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   2565
      Left            =   6840
      TabIndex        =   5
      Top             =   1560
      Width           =   1995
   End
   Begin VB.DirListBox Dir2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   2565
      Left            =   4560
      TabIndex        =   4
      Top             =   1560
      Width           =   2235
   End
   Begin VB.DriveListBox Drive2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   1140
      Width           =   2235
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   2565
      Left            =   2400
      TabIndex        =   2
      Top             =   1560
      Width           =   1995
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2235
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   2235
   End
   Begin AutoFileCopy.cpvProgressBar proFiles 
      Height          =   225
      Left            =   1980
      Top             =   4680
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   397
      BarPicture      =   "frmAutoFileCopy.frx":A528
      BarPictureBack  =   "frmAutoFileCopy.frx":F4B2
      CaptionFormat   =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontColor       =   65535
   End
   Begin AutoFileCopy.cpvProgressBar proFolders 
      Height          =   225
      Left            =   1980
      Top             =   4380
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   397
      BarPicture      =   "frmAutoFileCopy.frx":1443C
      BarPictureBack  =   "frmAutoFileCopy.frx":193C6
      CaptionFormat   =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontColor       =   65535
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   5940
      Width           =   8715
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   3
      Left            =   60
      Top             =   5880
      Width           =   8835
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "File Types"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   2400
      TabIndex        =   25
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   22
      Top             =   4980
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Folders"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   4380
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Files"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " File copy - ONLY  New files will be copied"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   660
      TabIndex        =   8
      Top             =   60
      Width           =   7215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination Folder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   7
      Top             =   660
      Width           =   3435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Source Folder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   660
      Width           =   2715
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   3555
      Index           =   0
      Left            =   60
      Top             =   660
      Width           =   4395
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   3555
      Index           =   1
      Left            =   4500
      Top             =   660
      Width           =   4395
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Index           =   2
      Left            =   60
      Top             =   4260
      Width           =   8835
   End
End
Attribute VB_Name = "frmAutoFileCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private LastDrive1           As Integer
Private LastDrive2           As Integer
Private FileCount            As Long
Private FolderCount          As Long
Private TotalFileSise        As Single
'Private FilesScaned          As Long
'Private FoldersScaned        As Long
Private FilesMoved           As Long
'Private FilesSkiped          As Long
Private SizeCopied           As Long
Private CurrFolderCount      As Long
Private LastRunTime As String
Private RunLocked As Boolean
Private FromFile     As Long
Private ToFile       As Long
Private QuitRequested As Boolean






Private Sub cboFileTypes_Click()

   File1.Pattern = cboFileTypes.Text

End Sub

Private Sub cmdClose_Click()

   Unload Me

End Sub

Private Sub cmdCount_Click()
If RunLocked = True Then Exit Sub
   CountFiles

End Sub

Private Sub cmdEdit_Click()
   If cmdEdit.Caption = "Un-Lock" Then
      If RunLocked = True Then Exit Sub
      cmdEdit.Caption = "Lock"
      RunLocked = True
     Else
      cmdEdit.Caption = "Un-Lock"
      RunLocked = False
   End If
   
   cboFileTypes.Enabled = Not cboFileTypes.Enabled
   Dir1.Enabled = Not Dir1.Enabled
   File1.Enabled = Not File1.Enabled
   Drive1.Enabled = Not Drive1.Enabled
   Dir2.Enabled = Not Dir2.Enabled
   File2.Enabled = Not File2.Enabled
   Drive2.Enabled = Not Drive2.Enabled
   optCopySubFolder.Enabled = Not optCopySubFolder.Enabled
   cboUpdateRate.Enabled = Not cboUpdateRate.Enabled


End Sub

Private Sub cmdRunNow_Click()
   If RunLocked = True Then
      If cmdEdit.Caption = "Lock" Then
         MsgBox "Click Lock to continue", vbCritical
      Else
      
      End If
      Exit Sub
   End If
   StartTheRun
End Sub

Private Sub StartTheRun()
   If RunLocked = True Then Exit Sub
   RunLocked = True
   CountFiles
   proFiles.Value = 0
   SizeCopied = 0
   FilesMoved = 0
   CurrFolderCount = 0
   Me.lblStatus = " Scanning " & FileCount & " files for coping in progress..."
   TotalFileSise = CSng(MoveNewFiles(Dir1.Path, cboFileTypes.Text))
   LastRunTime = Format(Now(), "hh:mm")
   lblStatus.Caption = " Last run @ " & Now() & " -:- " & FileCount & " Files Scaned" & " -:- " & FilesMoved & " Files copied -:- Copy Size = " & proSize.Value & " Bytes"
   RunLocked = False
End Sub

Private Function CopyNewFile(SourceFile As String, _
                              DestFile As String) As Boolean

  Dim Bytearray()  As Byte
  Dim FileSize     As Long
  Dim FromFile     As Long
  Dim ToFile       As Long
  Dim FromFileDate As String
  Dim ToFileDate   As String
  Dim foundPos     As Long
  Dim tPath        As String
DestFile = Replace(DestFile, "\\", "\", 1, -1, vbBinaryCompare)
'DestFile = Replace(DestFile, "(", "", 1, -1, vbBinaryCompare)
'DestFile = Replace(DestFile, ")", "", 1, -1, vbBinaryCompare)
   FromFileDate = GetFileDate(SourceFile)
   ToFileDate = GetFileDate(DestFile)
   ' File not modified
   If FromFileDate <= ToFileDate Then
      CopyNewFile = False
      Exit Function
   End If
   foundPos = InStrRev(DestFile, "\", -1, vbBinaryCompare)
   tPath = Left$(DestFile, foundPos)
   MakeDIR tPath
   FromFile = FreeFile
   Open SourceFile For Binary Access Read As #FromFile
   ToFile = FreeFile
   Open DestFile For Binary Access Write As #ToFile
   FileSize = LOF(1)
   ReDim Bytearray(FileSize)
   Get #FromFile, , Bytearray
   Put #ToFile, , Bytearray
   Close FromFile
   Close ToFile
   CopyNewFile = True

End Function

Private Sub CountFiles()
   Me.lblStatus = " File count in progress..."
   DoEvents
   FileCount = 0
   FolderCount = 0
   TotalFileSise = Size_Of_All_Files_Found_Under(Dir1.Path, cboFileTypes.Text)
   txtSize.Text = Format$(TotalFileSise, "#,##0")
   txtFolders.Text = FolderCount
   txtFiles.Text = FileCount
   proFolders.Max = IIf(FolderCount > 0, FolderCount, 1)
   proFiles.Max = IIf(FileCount > 0, FileCount, 1)
   proSize.Max = IIf(TotalFileSise > 0, TotalFileSise, 1)
   Me.Refresh
   Me.lblStatus = " "
End Sub


Private Sub Dir1_Change()

   File1.Pattern = cboFileTypes.Text
   File1.Path = Dir1.Path                       'Set file path.
   txtFolders.Text = ""
   txtFiles.Text = ""
   txtSize.Text = ""

End Sub

Private Sub Dir2_Change()

   File2.Path = Dir2.Path                       'Set file path.

End Sub

Private Sub Drive1_Change()

  Dim response As Long
  Dim Message  As String

   On Error GoTo Err_Init
   Dir1.Path = Drive1.Drive                     'set directory path1.
Exit_Err_Init:
   LastDrive1 = Drive1.ListIndex                'save current drive

Exit Sub

Err_Init:
   Select Case Err.Number
     Case 68
      Message = "drive is not available"
      response = MsgBox(Message, vbRetryCancel + vbCritical, "Selecting Drive")
      If response = vbRetry Then
         Resume
        Else
         Drive1.Drive = Drive1.List(LastDrive1)  'Set Last Drive
         Resume Next
      End If
     Case Else
      MsgBox Err.Description, vbOKOnly + vbExclamation
      Resume Next
   End Select

End Sub

Private Sub Drive2_Change()

  Dim response As Long
  Dim Message  As String

   On Error GoTo Err_Init
   Dir2.Path = Drive2.Drive                     'set directory path2.
Exit_Err_Init:
   LastDrive2 = Drive2.ListIndex                'save current drive

Exit Sub

Err_Init:
   Select Case Err.Number
     Case 68
      Message = "Drive is not available"
      response = MsgBox(Message, vbRetryCancel + vbCritical, "Selecting Drive")
      If response = vbRetry Then
         Resume
        Else
         Drive2.Drive = Drive2.List(LastDrive2)  'Set Last Drive
         Resume Next
      End If
     Case Else
      MsgBox Err.Description, vbOKOnly + vbExclamation
      Resume Next
   End Select

End Sub

Private Sub Form_Load()
   txtTime.Text = Format$(Now(), "hh:mm")       'Set Time
   LastDrive1 = Drive1.ListIndex                'Save current drive
   LastDrive2 = Drive2.ListIndex                'Save current drive
   LoadcboUpdateRate                            'Load Combo
   LoadcboFileTypes
   txtTime.Locked = True
   txtFiles.Locked = True
   txtFolders.Locked = True
   LoadLastSettings
   'CreateIcon picIcon
   'SetToolTip "Auto File Copy"
End Sub

Private Sub LoadLastSettings()
   Drive1.ListIndex = GetSetting(App.Title, "Defaults", "Drive1Index", 1)
   Drive2.ListIndex = GetSetting(App.Title, "Defaults", "Drive2Index", 1)
   Dir1.Path = GetSetting(App.Title, "Defaults", "Dir1Path", "C:\")
   Dir2.Path = GetSetting(App.Title, "Defaults", "Dir2Path", "C:\")
   cboUpdateRate.ListIndex = GetSetting(App.Title, "Defaults", "cboUpdateRateIndex", 0)
   optCopySubFolder.Value = GetSetting(App.Title, "Defaults", "optCopySubFolderValue", 0)
End Sub

Private Sub SaveLastSettings()
   SaveSetting App.Title, "Defaults", "Drive1Index", Drive1.ListIndex
   SaveSetting App.Title, "Defaults", "Drive2Index", Drive2.ListIndex
   SaveSetting App.Title, "Defaults", "Dir1Path", Dir1.Path
   SaveSetting App.Title, "Defaults", "Dir2Path", Dir2.Path
   SaveSetting App.Title, "Defaults", "cboUpdateRateIndex", cboUpdateRate.ListIndex
   SaveSetting App.Title, "Defaults", "optCopySubFolderValue", optCopySubFolder.Value
End Sub

Private Function GetFileDate(FileName As String) As String

   On Error Resume Next
       GetFileDate = FileDateTime(FileName)

End Function

Private Sub LoadcboFileTypes()

   With cboFileTypes
      .Clear
      .AddItem "*.*"
      .AddItem "*.exe"
      .AddItem "*.txt"
      .AddItem "*.csv"
      .AddItem "*.xls"
      .AddItem "*.doc"
      .AddItem "*.mdb"
      .AddItem "*.rpt"
      .AddItem "*.rep"
      .AddItem "*.rtf"
      .AddItem "*.dot"
      .AddItem "*.pdf"
      .ListIndex = 0
   End With

End Sub

Private Sub LoadcboUpdateRate()

  Dim count As Long

   With cboUpdateRate
      .Clear
      .AddItem "No Auto"
      .AddItem "15 Min"
      .AddItem "30 Min"
      .AddItem "1  hrs"
      .AddItem "2  hrs"
      For count = 1 To 23
         .AddItem Format$(count, "00") & ":00"
      Next count
      .ListIndex = 0
   End With

End Sub

Private Sub MakeDIR(Path As String)

   On Error GoTo Err_Init
   MkDir Path


Exit_Err_Init:
Exit Sub
LastResort:
   CreateSubFolders Path
   
Exit Sub
Err_Init:
   Select Case Err.Number
   Case 0
   Case 75
      Err.Clear
      Resume Exit_Err_Init
   Case 76
      Err.Clear
      Resume LastResort
   Case Else
   
   End Select

End Sub

Private Sub CreateSubFolders(rPath As String)
      Dim FolderList() As String
      Dim count As Integer
      Dim LastPath As String
      FolderList = Split(rPath, "\", -1, vbBinaryCompare)
      LastPath = LastPath & FolderList(0) & "\"
      For count = 1 To UBound(FolderList)
         LastPath = LastPath & FolderList(count) & "\"
         MakeDIR LastPath
      Next count

End Sub

Private Function MoveNewFiles(path_to_search As String, _
                               file_to_search_for As String) As Single

  Dim search_result            As String
  Dim name_of_subdirectory     As String
  Dim directory_names_array()  As String
  Dim number_of_subdirectories As Long
  Dim counter                  As Long
   If QuitRequested = True Then
      RunLocked = False
      Exit Function
   End If
   If Right(path_to_search, 1) <> "\" Then
      path_to_search = path_to_search & "\"
   End If
   'Find out how many subdirectories there are, and load their names into the array.
   number_of_subdirectories = 0
   ReDim directory_names_array(number_of_subdirectories)
   'Get the name  of the first subdirectory.
   name_of_subdirectory = Dir(path_to_search, vbDirectory Or vbHidden)
   Do While Len(name_of_subdirectory) > 0
      'Ignore the current and parent directories.
      If (name_of_subdirectory <> ".") Then
         If (name_of_subdirectory <> "..") Then
            If GetAttr(path_to_search & name_of_subdirectory) And vbDirectory Then
               directory_names_array(number_of_subdirectories) = name_of_subdirectory
               number_of_subdirectories = number_of_subdirectories + 1
               ReDim Preserve directory_names_array(number_of_subdirectories)
            End If
         End If
      End If
      name_of_subdirectory = Dir()  'Get next subdirectory.
   Loop
   'Search through this directory and total the file sizes.
   search_result = Dir(path_to_search & file_to_search_for, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
   Do While Len(search_result) <> 0
      MoveNewFiles = MoveNewFiles + 1
      MoveThisFile path_to_search & search_result
      search_result = Dir()  'Get next file.
   Loop
   'If there are sub-directories...
   If optCopySubFolder.Value Then
      If number_of_subdirectories > 0 Then
         'Have this function we are in call itself!
         For counter = 0 To number_of_subdirectories - 1
            If QuitRequested = True Then
               RunLocked = False
               Exit Function
            End If
            CurrFolderCount = CurrFolderCount + 1
            txtFolders.Text = CurrFolderCount & " of " & FolderCount
            proFolders.Value = CurrFolderCount
            MoveNewFiles = MoveNewFiles + MoveNewFiles(path_to_search & directory_names_array(counter) & "\", file_to_search_for)
         Next counter
      End If
   End If

End Function


Function Size_Of_All_Files_Found_Under(path_to_search As String, file_to_search_for As String)

'##### vote for this routeen at http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=9196&lngWId=1 ###

Dim search_result As String
Dim name_of_subdirectory As String
Dim directory_names_array() As String
Dim number_of_subdirectories As Integer
Dim counter As Integer

'Make sure the search path ends with a backslash.
    If Right(path_to_search, 1) <> "\" Then path_to_search = path_to_search & "\"
    
'Find out how many subdirectories there are, and load their names into the array.
    number_of_subdirectories = 0
    ReDim directory_names_array(number_of_subdirectories)
    
    'Get the name  of the first subdirectory.
    name_of_subdirectory = Dir(path_to_search, vbDirectory Or vbHidden)
    
Do While Len(name_of_subdirectory) > 0
    'Ignore the current and parent directories.
    If (name_of_subdirectory <> ".") And (name_of_subdirectory <> "..") Then
        If GetAttr(path_to_search & name_of_subdirectory) And vbDirectory Then
            directory_names_array(number_of_subdirectories) = name_of_subdirectory
            number_of_subdirectories = number_of_subdirectories + 1
            ReDim Preserve directory_names_array(number_of_subdirectories)
        End If
    End If
    'FileCount = FileCount + 1
    name_of_subdirectory = Dir()  'Get next subdirectory.
Loop

'Search through this directory and total the file sizes.
    search_result = Dir(path_to_search & file_to_search_for, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
    While Len(search_result) <> 0
      FileCount = FileCount + 1
        Size_Of_All_Files_Found_Under = Size_Of_All_Files_Found_Under + FileLen(path_to_search & search_result)
        search_result = Dir()  'Get next file.
    Wend

'If there are sub-directories...
   If optCopySubFolder Then
       If number_of_subdirectories > 0 Then
           'Have this function we are in call itself!
            For counter = 0 To number_of_subdirectories - 1
               FolderCount = FolderCount + 1
              Size_Of_All_Files_Found_Under = Size_Of_All_Files_Found_Under + Size_Of_All_Files_Found_Under(path_to_search & directory_names_array(counter) & "\", file_to_search_for)
            Next counter
       End If
   End If
End Function

Private Sub MoveThisFile(FromFileName As String)

  Dim FileWasCopied As Boolean
  Dim FromFolder    As String
  Dim ToFolder      As String
  Dim foundPos      As Long
  Dim fPath         As String
  Dim fFile         As String
  Dim tPath         As String
  Dim ToFileName    As String
   fPath = IIf(Right$(Dir1, 1) = "\", Dir1, Dir1 & "\")
   fFile = Replace$(FromFileName, fPath, "", 1, -1, vbBinaryCompare)
   foundPos = InStrRev(fPath, "\", Len(fPath) - 1, vbBinaryCompare)
   FromFolder = Right$(fPath, Len(fPath) - foundPos)
   tPath = IIf(Right$(Dir2, 1) = "\", Dir2, Dir2 & "\")
   foundPos = InStrRev(tPath, "\", Len(tPath) - 1, vbBinaryCompare)
   If foundPos = 0 Then
      ToFolder = FromFolder
      ToFileName = tPath & ToFolder & fFile
     Else
      ToFolder = Right$(tPath, foundPos + 1)
      ToFileName = tPath & fFile
   End If
   FileWasCopied = CopyNewFile(FromFileName, ToFileName)
   If FileWasCopied Then
      FilesMoved = FilesMoved + 1
      txtFiles.Text = FilesMoved & " of " & FileCount
      proFiles.Value = FilesMoved
      SizeCopied = SizeCopied + FileLen(FromFileName)
   End If
   proSize.Value = SizeCopied
   DoEvents

End Sub

Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then
      SetToolTip "Auto File Copy"
      MoveToTaksbar Me, picIcon
      Exit Sub
   End If
   If Me.Width <> 9030 Then Me.Width = 9030
   If Me.Height <> 6690 Then Me.Height = 6690
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Dim resp As Integer
   If RunLocked = True Then
      resp = MsgBox("Do you wish to Quit the currect Copy run?", vbYesNo)
      If resp = vbYes Then
         QuitRequested = True
         Cancel = True
      Else
         Cancel = True
         Exit Sub
      
      End If
   End If
   SaveLastSettings
   DeleteIcon picIcon
End Sub

Private Sub picIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

     Call RestoreFromTaskbar(Me, _
         picIcon, Button, Shift, X, Y)

End Sub

Private Sub Timer1_Timer()
   txtTime.Text = Format$(Now(), "hh:mm")
   If QuitRequested = True Then
      RunLocked = False
      Unload Me
   End If
End Sub

Private Sub txtTime_Change()
Dim ttime As String
If RunLocked = True Then Exit Sub

Select Case cboUpdateRate
Case "No Auto"
Case "15 Min"
   ttime = Right(txtTime, 2)
   If ttime Mod 15 = 0 Then
      StartTheRun
   End If
Case "30 Min"
   ttime = Right(txtTime, 2)
   If ttime Mod 30 = 0 Then
      StartTheRun
   End If
Case "1  hrs"
   ttime = Right(txtTime, 2)
   If ttime <> "00" Then Exit Sub
   If Left(txtTime, 2) = Left(txtTime, 2) Then Exit Sub
   StartTheRun
Case "2  hrs"
   ttime = Right(txtTime, 2)
   If ttime <> "00" Then Exit Sub
   ttime = Left(txtTime, 2)
   If ttime Mod 2 <> 0 Then Exit Sub
   If Left(txtTime, 2) = Left(txtTime, 2) Then Exit Sub
   StartTheRun
Case Else
   If cboUpdateRate = txtTime Then
      StartTheRun
   End If
End Select
End Sub
