VERSION 5.00
Begin VB.Form frmSelf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Driver Auto Install"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      TabIndex        =   2
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install selected drivers"
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
      Left            =   4320
      TabIndex        =   1
      Top             =   4200
      Width           =   2655
   End
   Begin DriverAutoInstall.LynxGrid LynxGrid1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7223
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
   Begin VB.Image Image2 
      Height          =   225
      Left            =   2400
      Picture         =   "frmSelf.frx":0000
      Top             =   4280
      Width           =   1530
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   2400
      Picture         =   "frmSelf.frx":067A
      Top             =   4280
      Width           =   1530
   End
End
Attribute VB_Name = "frmSelf"
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


Dim SDir As String
Private Sub cmdInstall_Click()
'# Discovery total of drivers to be installed #
k = -1
total = -1
For i = 0 To LynxGrid1.ItemCount - 1 '# loop all drivers in grid #
    If LynxGrid1.ItemChecked(i) = True Then '# if this driver is selected on grid #
        total = total + 1
    End If
Next
'# Discovery total of drivers to be installed #

If total = -1 Then
    MsgBox "Select drivers do you want to install", vbInformation + vbOKOnly
    Exit Sub
End If

'# disable buttons #
cmdExit.Enabled = False
cmdInstall.Enabled = False
'# disable buttons #

MousePointer = 11 '# show hourglass cursor #

'# this command is a system install for inf drivers #
CMD = Getpath_SYSTEM & "\rundll32.exe setupapi,InstallHinfSection DefaultInstall 132 "


For i = 0 To LynxGrid1.ItemCount - 1 '# loop all drivers in grid #
    If LynxGrid1.ItemChecked(i) = True Then '# if this driver is selected on grid #
        '# discover inf file in folder #
        sInf = LynxGrid1.CellText(i, 2)
        sFile = Dir(sInf & "\*.inf")
        '# finally shell instalation #
        ShellAndWait CMD & sInf & "\" & sFile, vbNormalFocus
        DoEvents
        
        '# show progress #
        k = k + 1
        Image2.Width = (Image1.Width / total) * k
    End If
Next

'# show info #
resp = MsgBox("Drivers installed. You need reboot you system to take effect. Do you want to reboot now?", vbInformation + vbYesNo + vbDefaultButton1)
If resp = vbYes Then
    '# reboot windows #
    ShutDownWindows True, True
    Unload Me
Else
    '# enable buttons #
    cmdExit.Enabled = True
    cmdInstall.Enabled = True
    '# enable buttons #
    
    MousePointer = 0 '# show default cursor #
End If
End Sub

Private Sub cmdExit_Click()
Unload Me '# this function is self explain ;-) #
End Sub

Private Sub Form_Load()
Dim Class() As String

SDir = Command '# get temp dir of drivers #
If SDir = "" Then
    End
End If

'# Prepare LynxGrid #
LynxGrid1.Font.Name = "Arial"
LynxGrid1.Font.Bold = True
LynxGrid1.AddColumn "Class", 1780
LynxGrid1.AddColumn "Description", 5000
LynxGrid1.AddColumn "folder", 1000
LynxGrid1.CheckBoxes = True
LynxGrid1.DisplayEllipsis = True
LynxGrid1.ColumnDrag = False
LynxGrid1.ColumnSort = True
LynxGrid1.AllowUserResizing = lgResizeCol
LynxGrid1.Editable = False
LynxGrid1.RowHeightMin = 400
LynxGrid1.ColVisible(2) = False
'# Prepare LynxGrid #

Image2.Width = 0 '# hide "virtual" progressbar #

LynxGrid1.Clear '# clear grid #
LynxGrid1.Redraw = False '# prepare grid to not refresh #

'# List all class folders #
n = -1
folder = Dir(SDir & "\", vbDirectory)
Do While folder <> ""
    If folder <> "." And folder <> ".." Then
    If (GetAttr(SDir & "\" & folder) And vbDirectory) = vbDirectory Then
        n = n + 1
        ReDim Preserve Class(n)
        Class(n) = folder
    End If
    End If
    folder = Dir()
Loop
'# List all class folders #

'# List all drivers in class folders #
n = -1
For i = 0 To UBound(Class)
    folder = Dir(SDir & "\" & Class(i) & "\", vbDirectory)
    Do While folder <> ""
        If folder <> "." And folder <> ".." Then
        If (GetAttr(SDir & "\" & Class(i) & "\" & folder) And vbDirectory) = vbDirectory Then
            n = n + 1
            LynxGrid1.AddItem n
            LynxGrid1.CellText(n, 0) = Class(i)
            LynxGrid1.CellText(n, 1) = folder
            LynxGrid1.CellText(n, 2) = SDir & "\" & Class(i) & "\" & folder
        End If
        End If
        folder = Dir()
    Loop
Next
'# List all drivers in class folders #


For i = 0 To LynxGrid1.ItemCount - 1 '# loop all drivers #
    LynxGrid1.ItemChecked(i) = True '# select drivers on grid #
Next

LynxGrid1.Redraw = True '# refresh grid #
End Sub
