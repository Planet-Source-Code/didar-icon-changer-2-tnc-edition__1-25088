VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command34 
      Caption         =   "Dial UpNe"
      Height          =   375
      Left            =   1680
      Picture         =   "FrmMain.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "To Change  Dial Up NetWork Icon"
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Printer"
      Height          =   375
      Left            =   600
      Picture         =   "FrmMain.frx":2413
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "To Change  Printer Icon"
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Control"
      Height          =   375
      Left            =   6000
      Picture         =   "FrmMain.frx":395C
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "To Change  Control Panel Icon"
      Top             =   2640
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton Command43 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmMain.frx":4EA5
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton Command42 
         Caption         =   "Inf File"
         Height          =   375
         Left            =   3720
         Picture         =   "FrmMain.frx":65DD
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "To Change  Inf File"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Dll File"
         Height          =   375
         Left            =   2640
         Picture         =   "FrmMain.frx":7B26
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "To Change  DLL File Icon"
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton Command40 
         Caption         =   "Mpeg"
         Height          =   375
         Left            =   5880
         Picture         =   "FrmMain.frx":906F
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "To Change  MPEG File Icon"
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton Command39 
         Caption         =   "SysFile"
         Height          =   375
         Left            =   4920
         Picture         =   "FrmMain.frx":A5B8
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "To Change  System File Icon"
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Command38 
         Caption         =   "ShortCut"
         Height          =   375
         Left            =   3720
         Picture         =   "FrmMain.frx":BB01
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "To Change  Shortcut Icon"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton Command37 
         Caption         =   "Default"
         Height          =   375
         Left            =   2640
         Picture         =   "FrmMain.frx":D04A
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "To Change  *.dat File Icon"
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton Command36 
         Caption         =   "Html"
         Height          =   375
         Left            =   1560
         Picture         =   "FrmMain.frx":E593
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "To Change  HTML Document Icon"
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton Command35 
         Caption         =   "Ini File"
         Height          =   375
         Left            =   480
         Picture         =   "FrmMain.frx":FADC
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "To Change  Ini File Icon"
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Command31"
         Height          =   435
         Left            =   4560
         TabIndex        =   38
         Top             =   5520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1680
         TabIndex        =   37
         Text            =   "Text3"
         Top             =   6000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command30 
         Caption         =   "My Compu"
         Height          =   375
         Left            =   480
         Picture         =   "FrmMain.frx":11025
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "To Change  My Computer Icon"
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Command28"
         Height          =   375
         Left            =   3960
         TabIndex        =   35
         Top             =   5640
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Command29"
         Height          =   495
         Left            =   6000
         TabIndex        =   34
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Info"
         Height          =   375
         Left            =   4920
         Picture         =   "FrmMain.frx":1256E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "About Icon Changer"
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton Command26 
         Caption         =   "E X I T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         Picture         =   "FrmMain.frx":13CA6
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Exit To Windows System"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Diagnostics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         Picture         =   "FrmMain.frx":153DE
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "If No Change Of Icon,Click Here.."
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Config File"
         Height          =   375
         Left            =   4920
         Picture         =   "FrmMain.frx":16B16
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "To Change  Config File Icon"
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Bat File"
         Height          =   375
         Left            =   2640
         Picture         =   "FrmMain.frx":1805F
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "To Change  Batch File Icon"
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Txt File"
         Height          =   375
         Left            =   3720
         Picture         =   "FrmMain.frx":195A8
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "To Change  Text File Icon"
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton Command21 
         Caption         =   "RecyFull"
         Height          =   375
         Left            =   1560
         Picture         =   "FrmMain.frx":1AAF1
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "To Change  Recyclebin Full Icon"
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton Command20 
         Caption         =   "RecyEmpty"
         Height          =   375
         Left            =   480
         Picture         =   "FrmMain.frx":1C03A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "To Change  Empty Recyclebin Icon"
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton Command19 
         Caption         =   "CD Drive"
         Height          =   375
         Left            =   5880
         Picture         =   "FrmMain.frx":1D583
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "To Change  CD-ROM Drive Icon"
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Application"
         Height          =   375
         Left            =   3720
         Picture         =   "FrmMain.frx":1EACC
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "To Change  Application Icon"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Windows"
         Height          =   375
         Left            =   4920
         Picture         =   "FrmMain.frx":20015
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "To Change  Default Icon"
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Drive"
         Height          =   375
         Left            =   2640
         Picture         =   "FrmMain.frx":2155E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "To Change Drive Icon"
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Floopy"
         Height          =   375
         Left            =   1560
         Picture         =   "FrmMain.frx":22AA7
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "To Change The Floopy Drive Icon"
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Printers"
         Height          =   375
         Left            =   5880
         Picture         =   "FrmMain.frx":23FF0
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "To Change StartMenu Setting's Printer Icon"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command13 
         Caption         =   "StartMenu"
         Height          =   375
         Left            =   4920
         Picture         =   "FrmMain.frx":25539
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "To Change StartMenu Program,List Of Programs Icon"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Control"
         Height          =   375
         Left            =   3720
         Picture         =   "FrmMain.frx":26A82
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "To Change Control Icon"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command11 
         Caption         =   "ShutDown"
         Height          =   375
         Left            =   2640
         Picture         =   "FrmMain.frx":27FCB
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "To Change ShutDown Icon"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Suspend"
         Height          =   375
         Left            =   1560
         Picture         =   "FrmMain.frx":29514
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "To Change StartMenu Suspend Icon"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Run"
         Height          =   375
         Left            =   480
         Picture         =   "FrmMain.frx":2AA5D
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "To Change StartMenu Run Icon"
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Help"
         Height          =   375
         Left            =   5880
         Picture         =   "FrmMain.frx":2BFA6
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "To Change StartMenu Help Icon"
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Find"
         Height          =   375
         Left            =   4920
         Picture         =   "FrmMain.frx":2D4EF
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "To Change StartMenu Find Icon"
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Setting"
         Height          =   375
         Left            =   3720
         Picture         =   "FrmMain.frx":2EA38
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "To Change StartMenu Setting Icon"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Documents"
         Height          =   375
         Left            =   2640
         Picture         =   "FrmMain.frx":2FF81
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "To Change Startmenu Document Icon"
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Programs"
         Height          =   375
         Left            =   480
         Picture         =   "FrmMain.frx":314CA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "To Change StartMenu Program Icon"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Folder"
         Height          =   375
         Left            =   1560
         Picture         =   "FrmMain.frx":32A13
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "To Change Folder Icon"
         Top             =   1200
         Width           =   975
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   1320
         Picture         =   "FrmMain.frx":33F5C
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   7
         Top             =   360
         Width           =   540
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   6000
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Apply"
         Height          =   375
         Left            =   6840
         Picture         =   "FrmMain.frx":34826
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Apply For Change"
         Top             =   4800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Text            =   "General System"
         Top             =   5400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Browse"
         Height          =   375
         Left            =   5400
         TabIndex        =   1
         ToolTipText     =   "Select The File For StartUp.."
         Top             =   5520
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComDlg.CommonDialog cmdlg 
         Left            =   3480
         Top             =   5880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.OLE OLE1 
         Class           =   "Package"
         Height          =   375
         Left            =   360
         OleObjectBlob   =   "FrmMain.frx":35D6F
         SourceDoc       =   "C:\WINDOWS\DESKTOP\IconChanger1.0\Refresh.bat"
         TabIndex        =   53
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "General Corporation.Bangladesh   Product Id: 47144-HJ-KNHHH"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   52
         Top             =   4680
         Width           =   4935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Help"
         Height          =   255
         Left            =   6120
         TabIndex        =   49
         Top             =   720
         Width           =   390
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "CopyRight By General Corporation 2000"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "GSI ICON CHANGER 2.0"
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
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Written By Didar"
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Menu MnuMain 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu MnuMainShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu MnuMainHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu MnuMainS1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMainNext 
         Caption         =   "&Next"
      End
      Begin VB.Menu MnuMainBack 
         Caption         =   "&Back"
      End
      Begin VB.Menu MnuMainS2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMainClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

 Const HKEY_LOCAL_MACHINE = &H80000002
  Const HKEY_CLASSES_ROOT = &H80000000

Public Sub SaveString(hKey As Long, StrPath As String, StrValue As String, StrData As String)
   Dim KeyH&
    r = RegCreateKey(hKey, StrPath, KeyH&)
    r = RegSetValueEx(KeyH&, StrValue, 0, 1, ByVal StrData, Len(StrData))
    r = RegCloseKey(KeyH&)
End Sub

Public Sub delString(hKey As Long, StrPath As String)
   Dim KeyH&
    r = RegDeleteKey(hKey, StrPath)
     r = RegCloseKey(KeyH&)
   End Sub




Private Sub Command1_Click()
If Command1.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\shell icons", Text2.Text, Text2.Text
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\shell icons", Text2.Text, (Text1.Text & ",0")
End If

End Sub


Private Sub Command10_Click()
Text2.Text = 25
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command11_Click()
Text2.Text = 27
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command12_Click()
Text2.Text = 35
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command13_Click()
Text2.Text = 36
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command14_Click()
Text2.Text = 37
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command15_Click()
Text2.Text = 6
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command16_Click()
Text2.Text = 8
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command17_Click()
Text2.Text = 0
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command18_Click()
Text2.Text = 2
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command19_Click()
Text2.Text = 11
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
cmdlg.FileName = ""
cmdlg.Filter = "*.ico|*.ico"
cmdlg.ShowOpen
Text1.Text = cmdlg.FileName
End Sub


Private Sub Command20_Click()
Text2.Text = 31
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command21_Click()
Text2.Text = 32
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command22_Click()
Text3.Text = "SOFTWARE\Classes\txtfile\DefaultIcon"
Command28_Click
Command2_Click
Command31_Click
End Sub

Private Sub Command23_Click()
Text3.Text = "SOFTWARE\Classes\batfile\DefaultIcon"
Command28_Click
Command2_Click
Command31_Click
End Sub

Private Sub Command24_Click()
Text2.Text = 61
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub




Private Sub Command25_Click()
On Error Resume Next
Dim a As Integer
a = MsgBox("                                     !!!Diagnostics!!!                                                                                                                                                                   If There Is No Change Of Icon Of Your System Click 'Yes'.", 49, "Diagnostics")
If a = vbCancel Then
Load Me
Else
i = Shell("c:\windows\sys.exe", vbNormalFocus)
End If
End Sub

Private Sub Command26_Click()
Dim a As Integer
a = MsgBox("Do You Really Want To Quit?", 49, "Quit")
If a = vbCancel Then
Load Me
Else
End
End If

End Sub

Private Sub Command27_Click()
Form1.Show
End Sub





Private Sub Command28_Click()
If Command28.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
delString HKEY_LOCAL_MACHINE, Text3.Text
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
delString HKEY_LOCAL_MACHINE, Text3.Text
End If
End Sub

Private Sub Command30_Click()
Text3.Text = "SOFTWARE\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon"
Command28_Click
Command2_Click
Command31_Click
End Sub

Private Sub Command31_Click()
If Command31.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, Text3.Text, "", ""
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, Text3.Text, "", (Text1.Text & ",0")
End If
End Sub

Private Sub Command32_Click()
Text3.Text = "SOFTWARE\Classes\CLSID\{21EC2020-3AEA-1069-A2DD-08002B30309D}\DefaultIcon"
Command28_Click
Command2_Click
Command31_Click
End Sub

Private Sub Command33_Click()
Text3.Text = "SOFTWARE\Classes\CLSID\{2227A280-3AEA-1069-A2DE-08002B30309D}\DefaultIcon"
Command28_Click
Command2_Click
Command31_Click
End Sub

Private Sub Command34_Click()
Text3.Text = "SOFTWARE\Classes\CLSID\{992CFFA0-F557-101A-88EC-00DD010CCC48}\DefaultIcon"
Command28_Click
Command2_Click
Command31_Click
End Sub

Private Sub Command35_Click()
Text3.Text = "SOFTWARE\Classes\inifile\DefaultIcon"
Command28_Click
Command2_Click
Command31_Click
End Sub

Private Sub Command36_Click()
Text3.Text = "SOFTWARE\Classes\htmlfile\DefaultIcon"
Command28_Click
Command2_Click
Command31_Click
End Sub

Private Sub Command37_Click()
Text2.Text = 1
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command38_Click()
Text2.Text = 29
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command39_Click()
Text3.Text = "SOFTWARE\Classes\sysfile\DefaultIcon"
Command28_Click
Command2_Click
Command31_Click
End Sub

Private Sub Command4_Click()
Text2.Text = 19
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub


Private Sub Command3_Click()
Text2.Text = 3
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command40_Click()
Text3.Text = "SOFTWARE\Classes\MPEGFile\DefaultIcon"
Command28_Click
Command2_Click
Command31_Click
End Sub

Private Sub Command41_Click()
Text3.Text = "SOFTWARE\Classes\dllfile\DefaultIcon"
Command28_Click
Command2_Click
Command31_Click
End Sub

Private Sub Command42_Click()
Text3.Text = "SOFTWARE\Classes\inffile\DefaultIcon"
Command28_Click
Command2_Click
Command31_Click
End Sub

Private Sub Command43_Click()
On Error Resume Next
Dim a As Integer
a = MsgBox("This New Setting Will Take Effect After Restart Your Computer. Do You Want To Restart Now?", 49, "You Need To Restart")
If a = vbCancel Then
Load Me
Else
i = Shell("c:\windows\rundll.exe user.exe,exitwindowsexec", vbNormalFocus)
End If
End Sub

Private Sub Command44_Click()
On Error Resume Next
i = Shell("c:\windows\sys", vbNormalFocus)
End Sub

Private Sub Command5_Click()
Text2.Text = 20
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command6_Click()
Text2.Text = 21
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command7_Click()
Text2.Text = 22
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command8_Click()
Text2.Text = 23
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Command9_Click()
Text2.Text = 24
Command2_Click
If cmdlg.FileName = "" Then
MsgBox "You Must Select An Icon", 16, "Info"
Else
Command1_Click
End If
End Sub

Private Sub Label2_Click()
MsgBox "If There Is No Change Of Icon,Click 'Diagnostics' Or Run 'Refresh Icon'.Don't Click 'Diagnostics' If Everything Is 'OK'.Run 'Icon Color Level' For Show Icon With Full Color.", 32, "It's Important"
End Sub

