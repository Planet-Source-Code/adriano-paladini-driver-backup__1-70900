VERSION 5.00
Object = "{DB797681-40E0-11D2-9BD5-0060082AE372}#5.0#0"; "xzip.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drivers BackUp"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10815
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8640
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton cmdInvertSelection 
      Caption         =   "Invert selection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   4680
      Width           =   1695
   End
   Begin DriversBackup.LynxGrid LynxGrid1 
      Height          =   4215
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7435
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdDoBackup 
      Caption         =   "Do BackUp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   2
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "non-Microsoft drivers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.Label lblStatus 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   4755
      Width           =   3180
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   120
      Picture         =   "frmMain.frx":599A
      Top             =   4755
      Width           =   1530
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   120
      Picture         =   "frmMain.frx":6014
      Top             =   4760
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "Type of backup:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   60
      Width           =   1575
   End
   Begin XceedZipLibCtl.XceedZip Zip 
      Left            =   4320
      Top             =   0
      BasePath        =   ""
      CompressionLevel=   6
      EncryptionPassword=   ""
      RequiredFileAttributes=   0
      ExcludedFileAttributes=   24
      FilesToProcess  =   ""
      FilesToExclude  =   ""
      MinDateToProcess=   2
      MaxDateToProcess=   2958465
      MinSizeToProcess=   0
      MaxSizeToProcess=   0
      SplitSize       =   0
      PreservePaths   =   -1  'True
      ProcessSubfolders=   0   'False
      SkipIfExisting  =   0   'False
      SkipIfNotExisting=   0   'False
      SkipIfOlderDate =   0   'False
      SkipIfOlderVersion=   0   'False
      TempFolder      =   ""
      UseTempFile     =   -1  'True
      UnzipToFolder   =   ""
      ZipFilename     =   ""
      SpanMultipleDisks=   2
      ExtraHeaders    =   10
      ZipOpenedFiles  =   0   'False
      BackgroundProcessing=   0   'False
      SfxBinrayModule =   ""
      SfxDefaultPassword=   ""
      SfxDefaultUnzipToFolder=   ""
      SfxExistingFileBehavior=   0
      SfxReadmeFile   =   ""
      SfxExecuteAfter =   ""
      SfxInstallMode  =   0   'False
      SfxProgramGroup =   ""
      SfxProgramGroupItems=   ""
      SfxExtensionsToAssociate=   ""
      SfxIconFilename =   ""
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###############################
'#                             #
'#       Drivers BackUp        #
'#             By              #
'#  adrianopaladini@gmail.com  #
'#                             #
'#        version 1.0          #
'#                             #
'###############################
'
'
'
' To create ZIP and Self-installer i use
' xceedzip component. to use project rename
' "xzip.ckk" to "xzip.dll" (this file is
' not required on compiled exe, because it
' has this file in resource.)
'
'
' The DriverBackup.exe contains in resource
' the icon, sfx, dlls, and setup.exe to make
' ZIP and EXE.
'
'
' To make changes in Self-Installer, use the
' setup.vbp in "self" folder, modify, compile
' and insert "setup.exe" in resource.res, open
' DriverBackUp.vbp and compile the main project.
'
'
'
' Files is resource.res:
'
' ICON.ICO      <- icon used in self-installer
' MSVBVM60.DLL  <- for run setup.exe in any windows
' SETUP.EXE     <- setup program. project in "self" folder
' SFX.BIN       <- binary for create self-installer
' XZIP.DLL      <- xceed dll to create zip and sfx file
'
'
'


'# Variable for drivers information #
Dim CH() As String
Function WhereIsDir(str)
'# function to discover dirs with inf code #
Select Case str
    Case "01"
        cDir = Mid(winDir, 1, 3)
    Case "10"
        cDir = winDir
    Case "11"
        cDir = sysDir
    Case "12"
        cDir = sysDir & "\Drivers"
    Case "17"
        cDir = infDir
    Case "18"
        cDir = winDir & "\Help"
    Case "20"
        cDir = winDir & "\Fonts"
    Case "21"
        cDir = "" 'viewer dir
    Case "23"
        cDir = sysDir & "\spool\drivers\color"
    Case "24"
        cDir = Mid(winDir, 1, 3)
    Case "25"
        cDir = "" 'shared dir
    Case "30"
        cDir = Mid(winDir, 1, 3)
    Case "50"
        cDir = sysDir
    Case "51"
        cDir = sysDir & "\Spool"
    Case "52"
        cDir = sysDir & "\Spool\Drivers"
    Case "53"
        cDir = "" 'user profile dir
    Case "54"
        cDir = "" ' ntldr.exe dir
    Case "55"
        cDir = sysDir & "\spool\prtprocs"
    Case "-1"
        cDir = "" ' absolute path
    Case "66000"
        cDir = sysDir & "\spool\Drivers\w32x86"
    Case "66001"
        cDir = sysDir & "\spool\prtprocs\w32x86"
    Case "66002"
        cDir = "" ' print monitor dir
    Case "66003"
        cDir = sysDir & "\spool\drivers\color"
    Case "66004"
        cDir = sysDir & "\spool\Drivers\w32x86"
    Case Else
        cDir = ""
End Select
WhereIsDir = cDir
End Function

Function SafeDir(str)
'# function to replace special chars to create dirs correctly #
R = str
R = Replace(R, "\", " ")
R = Replace(R, "/", " ")
R = Replace(R, "*", " ")
R = Replace(R, ":", " ")
R = Replace(R, ";", " ")
R = Replace(R, "?", " ")
R = Replace(R, ">", " ")
R = Replace(R, "<", " ")
R = Replace(R, "|", " ")
SafeDir = R
End Function

Private Sub ReadDrivers()
'# sub to read drivers and populate grid #
On Error Resume Next

MousePointer = 11 '# display hourglass cursor while read #
DoEvents

LynxGrid1.Clear '# clear grid #
LynxGrid1.Redraw = False '# prepare grid to not refresh #


'# list all class of drivers installed
n = 0
Z = ListKey(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Class")
For i = 0 To UBound(Z)
    U = ListKey(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Class\" & Z(i))
    For j = 0 To UBound(U)
        n = n + 1
        ReDim Preserve CH(n)
        CH(n) = Z(i) & "\" & U(j)
    Next
Next

'# get all info of each instaled driver #
h = -1
For i = 0 To UBound(CH)
    sInf = CH(i)
    inf0 = ReadKey(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Class\" & sInf, "ProviderName")
    inf1 = ReadKey(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Class\" & Mid(sInf, 1, Len(sInf) - 5), "")
    inf2 = ReadKey(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Class\" & Mid(sInf, 1, Len(sInf) - 5), "Class")
    inf3 = ReadKey(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Class\" & sInf, "DriverDesc")
    inf4 = ReadKey(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Class\" & sInf, "InfPath")
    If inf3 <> "" Then
        '# and populate grid #
        If inf0 <> "" And LCase(inf0) <> "microsoft" Then
            '# if is non-microsoft #
            h = h + 1
            LynxGrid1.AddItem h
            LynxGrid1.CellText(h, 0) = inf3
            LynxGrid1.CellText(h, 1) = inf1
            LynxGrid1.CellText(h, 2) = inf0
            LynxGrid1.CellText(h, 3) = sInf
            LynxGrid1.CellText(h, 4) = inf2
            LynxGrid1.CellText(h, 5) = inf4
        Else
            If Check1.Value = 0 Then
                '# or if is microsoft #
                h = h + 1
                LynxGrid1.AddItem h
                LynxGrid1.CellText(h, 0) = inf3
                LynxGrid1.CellText(h, 1) = inf1
                LynxGrid1.CellText(h, 2) = inf0
                LynxGrid1.CellText(h, 3) = sInf
                LynxGrid1.CellText(h, 4) = inf2
                LynxGrid1.CellText(h, 5) = inf4
            End If
        End If
    End If
Next

LynxGrid1.Redraw = True '# refresh grid #

cmdSelectAll_Click '# call select all function #

DoEvents
MousePointer = 0 '# display default cursor #
End Sub
Private Sub Check1_Click()

ReadDrivers '# call function to read drivers #

End Sub




Private Sub cmdDoBackup_Click()
'# do backup of selected drivers #
Dim inf() As String
Dim SH As New Shell
Dim ShBFF As Folder


'# verify if has selected drivers #
If LynxGrid1.CheckedCount = 0 Then
    MsgBox "Select the desired Drivers to make the BackUp.", vbInformation + vbOKOnly
    Exit Sub
End If

'# open folder select #
Set ShBFF = SH.BrowseForFolder(hWnd, "Select the folder where it will be Backup done.", 1)
If ShBFF Is Nothing Then '# if user cancel #
    Exit Sub
End If


MousePointer = 11 '# display hourglass cursor while read #
DoEvents

'# Disable buttons #
Check1.Enabled = False
Combo1.Enabled = False
cmdDoBackup.Enabled = False
cmdSelectAll.Enabled = False
cmdInvertSelection.Enabled = False
DoEvents
'# Disable buttons #

'# set dirs to variables #
destDir = ShBFF.Items.Item.Path

If Right(destDir, 1) = "\" Then
    destDir = destDir & "DriversBackUp"
Else
    destDir = destDir & "\DriversBackUp"
End If

winDir = Getpath_WINDOWS
sysDir = Getpath_SYSTEM
infDir = winDir & "\inf\"


'# show status
lblStatus = "Searching and saving driver files..."

'# loop al drivers in grid #
n = -1
For i = 0 To LynxGrid1.ItemCount - 1
    '# verify if driver is selected #
    If LynxGrid1.ItemChecked(i) = True Then
    
        n = n + 1
        ReDim Preserve inf(n)
        
        '# create necessaries folders #
        If Dir(destDir, vbDirectory) = "" Then MkDir destDir
        dest = destDir & "\" & LynxGrid1.CellText(i, 4)
        If Dir(dest, vbDirectory) = "" Then MkDir dest
        dest = destDir & "\" & LynxGrid1.CellText(i, 4) & "\" & SafeDir(LynxGrid1.CellText(i, 0))
        If Dir(dest, vbDirectory) = "" Then MkDir dest
        
        
        '# copy inf file of driver #
        FileCopy infDir & LynxGrid1.CellText(i, 5), dest & "\" & LynxGrid1.CellText(i, 5)
        DoEvents
        
        
        '# read this inf and verify if has catalog file
        Dim Cats() As String
        xf = -1
        catfile = ReadFromINI("Version", "CatalogFile", dest & "\" & LynxGrid1.CellText(i, 5), "")
        If catfile <> "" Then '# if has catalog file #
            '# list dirs of catalogs #
            pasta = Dir(sysDir & "\CatRoot\", vbDirectory)
            Do While pasta <> ""
                If pasta <> "." And pasta <> ".." Then
                    xf = xf + 1
                    ReDim Preserve Cats(xf)
                    Cats(xf) = sysDir & "\CatRoot\" & pasta & "\" & catfile
                End If
                pasta = Dir()
            Loop
            '# make backup of catalog #
            For jh = 0 To UBound(Cats)
                If Dir(Cats(jh), vbSystem) <> "" Then
                    FileCopy Cats(jh), dest & "\" & catfile
                End If
            Next
        End If
        
        
        '# read this inf and verify all necessaries files #
        Z = LoadIniSectionKeys("SourceDisksFiles", dest & "\" & LynxGrid1.CellText(i, 5))
        For d = 0 To UBound(Z)
            If Z(d) <> "" Then
                '# verify file extenssion #
                v = Split(Z(d), ".")
                ext = LCase(v(1))
                
                '# verify if it is in special/custom folder #
                Dim customDir As String
                customDir = ReadFromINI("DestinationDirs", "DefaultDestDir", dest & "\" & LynxGrid1.CellText(i, 5), "")
                If customDir <> "" Then '# if it is #
                    cDir = WhereIsDir(customDir) '# discover folder by inf ID #
                    '# copy file from correctly folder #
                    If Dir(cDir & "\" & Z(d)) <> "" Then FileCopy cDir & "\" & Z(d), dest & "\" & Z(d)
                    '# if is printer driver #
                    If cDir = (sysDir & "\spool\Drivers\w32x86") Then
                        '# search for correctly driver if has more tha one printer #
                        For kd = 1 To 19
                            '# copy file from correctly folder #
                            If Dir(cDir & "\" & kd & "\" & Z(d)) <> "" Then FileCopy cDir & "\" & kd & "\" & Z(d), dest & "\" & Z(d)
                        Next
                    End If
                End If
                
                '# copy file from correctly folder #
                If ext = "hlp" Then
                    '# copy file from HELP folder #
                    If Dir(winDir & "\help\" & Z(d)) <> "" Then FileCopy winDir & "\help\" & Z(d), dest & "\" & Z(d)
                ElseIf ext = "sys" Then
                    '# copy file from DRIVERS folder #
                    If Dir(sysDir & "\drivers\" & Z(d)) <> "" Then FileCopy sysDir & "\drivers\" & Z(d), dest & "\" & Z(d)
                Else
                    '# copy file from SYSTEM folder #
                    If Dir(sysDir & "\" & Z(d)) <> "" Then FileCopy sysDir & "\" & Z(d), dest & "\" & Z(d)
                End If
            End If
        Next
    End If
    
    '# show progress #
    Image2.Width = (Image1.Width / (LynxGrid1.ItemCount - 1)) * i
    
Next

DoEvents
MousePointer = 0 '# display default cursor #

'# type of backup #
If Combo1.ListIndex = 0 Then
    '# Do nothing #
End If
If Combo1.ListIndex = 1 Then
    '# show status #
    lblStatus = "Zipping driver files..."
    DoEvents
    '# create ZIP #
    DoZip destDir, destDir & ".zip"
    DoEvents
    '# delete temp folder #
    DeleteDir destDir
End If
If Combo1.ListIndex = 2 Then
    '# show status #
    lblStatus = "Creating Self-extractor..."
    DoEvents
    '# create SFX #
    DoSFX destDir, destDir & ".exe"
    DoEvents
    '# delete temp folder #
    DeleteDir destDir
End If


DoEvents
'# Enable buttons #
Check1.Enabled = True
Combo1.Enabled = True
cmdDoBackup.Enabled = True
cmdSelectAll.Enabled = True
cmdInvertSelection.Enabled = True
'# Enable buttons #

Image2.Width = 0 '# hide image of "virtual" progressbar #

'# show status #
lblStatus = ""

'# show info of end process #
MsgBox "Backup successfully created.", vbInformation + vbOKOnly
End Sub

Sub DeleteDir(strPath)
Dim fso As Object
Dim fldr As Object
Dim fl As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Set fldr = fso.GetFolder(strPath)
For Each fl In fldr.Files
    fl.Delete
Next fl
fldr.Delete
Set fldr = Nothing
Set fl = Nothing
Set fso = Nothing
End Sub

Sub DoZip(strPath, strFile)
On Error Resume Next
'Zip.License "" '# if you have put here your licency
Zip.PreservePaths = True '# to preserve folders in zip
Zip.DeleteZippedFiles = True '# to delete files create on folder
Zip.ProcessSubfolders = True '# to zip sub folders
Zip.BasePath = strPath '# base folder for zip
Zip.FilesToProcess = "*.*" '# files to be zipped
Zip.ZipFilename = strFile '# filename of zip
Zip.CompressionLevel = xclMedium '# compression Medium
ResultCode = Zip.Zip '# finally create zip
If ResultCode <> xerSuccess Then '# verify if occurred an erros
    MsgBox "Unsuccessful. Error on creating ZIP. " & _
    "Description: " & Zip.GetErrorDescription(xvtError, ResultCode)
End If
End Sub
Sub DoSFX(strPath, strFileSfx)
'# save necessaries files to be created SFX from resource #
SaveRes "SETUP.EXE", "CUSTOM", strPath & "\setup.exe", 270336
DoEvents
SaveRes "MSVBVM60.DLL", "CUSTOM", strPath & "\msvbmv60.dll", 1386496
DoEvents
SaveRes "SFX.BIN", "CUSTOM", App.Path & "\sfx.bin", 110602
DoEvents
SaveRes "ICON.ICO", "CUSTOM", App.Path & "\icon.ico", 1078
'# save necessaries files to be created SFX from resource #
DoEvents
'Zip.License "" '# if you have put here your licency
Zip.PreservePaths = True '# to preserve folders in zip
Zip.DeleteZippedFiles = True '# to delete files create on folder
Zip.ProcessSubfolders = True '# to zip sub folders
Zip.BasePath = strPath '# base folder for zip
Zip.FilesToProcess = "*.*" '# files to be zipped
Zip.ZipFilename = strFileSfx '# filename of exe
Zip.CompressionLevel = xclMedium '# compression Medium
Zip.EncryptionPassword = "minhasenhazip" '# insert a password on file
Zip.SfxDefaultPassword = "minhasenhazip" '# say to exe what password use to decompress
Zip.SfxBinaryModule = App.Path & "\sfx.bin" '# select the binary file for exe
Zip.SfxIconFilename = App.Path & "\icon.ico" '# select icon for exe
Zip.SfxExecuteAfter = "%d\setup.exe|%d" '# say to exe, to execute setup.exe from zip after decompress
Zip.SfxDefaultUnzipToFolder = "%t\_sfxDriverSetup"
Zip.SfxInstallMode = True '# mark to be a self-installer
Zip.SfxClearMessages '# to not show any window or message to decompress
ResultCode = Zip.Zip '# finally create zip
If ResultCode <> xerSuccess Then '# verify if occurred an erros
    MsgBox "Unsuccessful. Error on creating SelfExtractor. " & _
    "Description: " & Zip.GetErrorDescription(xvtError, ResultCode)
Else
    '# delete files used to be created SFX #
    If Dir(App.Path & "\sfx.bin") <> "" Then Kill App.Path & "\sfx.bin"
    If Dir(App.Path & "\icon.ico") <> "" Then Kill App.Path & "\icon.ico"
    '# delete files used to be created SFX #
End If
DoEvents
End Sub


Private Sub cmdInvertSelection_Click()

LynxGrid1.Redraw = False '# prepare grid to not refresh #
For i = 0 To LynxGrid1.ItemCount - 1 '# loop all drivers #
    LynxGrid1.ItemChecked(i) = Not LynxGrid1.ItemChecked(i) '# invert selection drivers on grid #
Next
LynxGrid1.Redraw = True '# refresh grid #

End Sub

Private Sub cmdSelectAll_Click()

LynxGrid1.Redraw = False '# prepare grid to not refresh #
For i = 0 To LynxGrid1.ItemCount - 1 '# loop all drivers #
    LynxGrid1.ItemChecked(i) = True '# select drivers on grid #
Next
LynxGrid1.Redraw = True '# refresh grid #

End Sub

Private Sub Form_Load()

'# Prepare LynxGrid #
LynxGrid1.Font.Name = "Arial"
LynxGrid1.Font.Bold = True
LynxGrid1.AddColumn "Description", 5520
LynxGrid1.AddColumn "Class", 3000
LynxGrid1.AddColumn "Provider", 2000
LynxGrid1.AddColumn "ID", 1000
LynxGrid1.AddColumn "Type", 1000
LynxGrid1.AddColumn "Inf", 1000
LynxGrid1.CheckBoxes = True
LynxGrid1.DisplayEllipsis = True
LynxGrid1.ColumnDrag = False
LynxGrid1.ColumnSort = True
LynxGrid1.AllowUserResizing = lgResizeCol
LynxGrid1.Editable = False
LynxGrid1.RowHeightMin = 400
LynxGrid1.ColVisible(3) = False
LynxGrid1.ColVisible(4) = False
LynxGrid1.ColVisible(5) = False
'# Prepare LynxGrid #

'# Prepare Combo #
Combo1.Clear
Combo1.AddItem "Folder with drivers", 0
Combo1.AddItem "Zip with drivers", 1
Combo1.AddItem "Self-Installer", 2
Combo1.ListIndex = 0
'# Prepare Combo #

Image2.Width = 0 '# hide image of "virtual" progressbar #

ReadDrivers '# call function to read drivers #

End Sub
Private Sub Zip_GlobalStatus(ByVal lFilesTotal As Long, ByVal lFilesProcessed As Long, ByVal lFilesSkipped As Long, ByVal nFilesPercent As Integer, ByVal lBytesTotal As Long, ByVal lBytesProcessed As Long, ByVal lBytesSkipped As Long, ByVal nBytesPercent As Integer, ByVal lBytesOutput As Long, ByVal nCompressionRatio As Integer)
On Error Resume Next
iPor = (Image1.Width / 100) * nBytesPercent
If iPor <> Image2.Width Then Image2.Width = iPor
End Sub

Private Sub Zip_Warning(ByVal sFilename As String, ByVal xWarning As XceedZipLibCtl.xcdWarning)
MsgBox xWarning
End Sub
