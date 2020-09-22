VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FormMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Page an ICQ User..."
   ClientHeight    =   3390
   ClientLeft      =   -10005
   ClientTop       =   1365
   ClientWidth     =   3765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton BtnExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3495
      Begin VB.TextBox TextUIN 
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   0
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Send Message to ICQ UIN:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSWinsockLib.Winsock SockPager 
      Left            =   120
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox TextMessage 
      Height          =   975
      Left            =   120
      MaxLength       =   450
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1800
      Width           =   3495
   End
   Begin VB.CommandButton BtnSend 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox TextSubject 
      Height          =   315
      Left            =   120
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label LabelStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Message:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   600
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cMessage As String
Dim cSubject As String
Private Sub BtnExit_Click()
   End
End Sub
Public Sub BtnSend_Click()
   On Error Resume Next
   
   Dim cSend As String
   Dim cData As String
   
   ' Verify UIN
   If Not IsNumeric(TextUIN.text) Then
      MsgBox "The ICQ UIN not Numeric!", "Error:"
      TextUIN.SetFocus
      Exit Sub
   End If
            
   If Trim(TextMessage.text) = "" Then
      MsgBox "Don't Allow Blank Messages!", "Error:"
      TextMessage.SetFocus
      Exit Sub
   End If

   ' Status
   LabelStatus.Caption = "Connecting..."
   
   ' Close Socket
   SockPager.Close
      
   ' Change the " " for "+"
   cSubject = ChangeSpaces(TextSubject.text)
   cMessage = ChangeSpaces(TextMessage.text)

   ' Fill the String with things such as "From" etc
   cData = "from=" & Form1.server.LocalHostName & "&fromemail=notifypro@notifypro.com&subject=" & cSubject & "&body=" & cMessage & "&to=" & Trim(TextUIN.text) & "&Send=" & """"
   cSend = "POST /scripts/WWPMsg.dll HTTP/1.0" & vbCrLf
   cSend = cSend & "Referer: http://wwp.mirabilis.com" & vbCrLf
   cSend = cSend & "User-Agent: Mozilla/4.06 (Win95; I)" & vbCrLf
   cSend = cSend & "Connection: Keep-Alive" & vbCrLf
   cSend = cSend & "Host: wwp.mirabilis.com:80" & vbCrLf
   cSend = cSend & "Content-type: application/x-www-form-urlencoded" & vbCrLf
   cSend = cSend & "Content-length: " & Len(cData) & vbCrLf
   cSend = cSend & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf & vbCrLf
   cSend = cSend & cData & vbCrLf & vbCrLf & vbCrLf & vbCrLf
   ' Send Message
   SockPager.Tag = cSend
   SockPager.Connect "wwp.mirabilis.com", 80
End Sub

Private Sub Form_Load()
   On Error Resume Next
FormMain.Hide

   ' Close Socket
   SockPager.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next

   ' Close Socket
   SockPager.Close
   
   ' Force Exit
   End
End Sub

Private Sub SockPager_Connect()
   On Error Resume Next
   
   ' Status
   LabelStatus.Caption = "Sending..."
  
   SockPager.SendData SockPager.Tag
End Sub

Private Sub SockPager_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   ' Status
   LabelStatus.Caption = "Error..."
   
   SockPager.Tag = ""
End Sub

Private Sub SockPager_SendComplete()
   ' Status
   LabelStatus.Caption = "Complete!"
   
   SockPager.Tag = ""
End Sub

Private Function ChangeSpaces(cString As String) As String
   On Error Resume Next
  
   Dim cChar As String
   Dim cReturn As String
  
   Dim nLoop As Long
  
   
   cReturn = ""
  
   For nLoop = 1 To Len(cString)
       cChar = Mid(cString, nLoop, 1)
      
       If cChar = " " Then
          cChar = "+"
       End If
      
       cReturn = cReturn + cChar
   Next
  
   ChangeSpaces = cReturn
End Function
