VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ListenFrm 
   Caption         =   "Sapient2003 Net Tools - Listener"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   Icon            =   "ListenFrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "ListenFrm.frx":0442
   ScaleHeight     =   299
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   448
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Listen"
      Default         =   -1  'True
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton optTCP 
      BackColor       =   &H80000012&
      Caption         =   "TCP/IP"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.OptionButton optUDP 
      BackColor       =   &H80000012&
      Caption         =   "UDP"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3000
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   960
      TabIndex        =   7
      Top             =   960
      Width           =   4575
      Begin VB.TextBox port3 
         Height          =   285
         Left            =   480
         MaxLength       =   5
         TabIndex        =   10
         Text            =   "139"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Close 
         Caption         =   "Close"
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "Protocol:"
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
         Left            =   0
         TabIndex        =   9
         Top             =   600
         Width           =   855
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
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Caption         =   "Status"
      ForeColor       =   &H8000000E&
      Height          =   2175
      Left            =   840
      TabIndex        =   5
      Top             =   1800
      Width           =   4815
      Begin VB.TextBox txtStatus 
         BackColor       =   &H80000009&
         ForeColor       =   &H80000007&
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   4335
      End
   End
   Begin MSWinsockLib.Winsock ws1 
      Left            =   3120
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "ListenFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Close_Click()
Unload Me
End Sub

Private Sub cmdConnect_Click()
cmdConnect.Enabled = False
port3.Enabled = False
cmdDisconnect.Enabled = True
txtStatus = ""
If optTCP = True Then
    ws1.Protocol = sckTCPProtocol
End If
If optUDP = True Then
    ws1.Protocol = sckUDPProtocol
End If
On Error GoTo PortIsOpen
ws1.Close
ws1.LocalPort = port3.Text
ws1.Listen
Exit Sub
PortIsOpen:
ws1.Close
If Err.Number = 10048 Then
    txtStatus = "The port " & port3.Text & " is already open."
Else
    txtStatus = "Error: " & Err.Number & vbCrLf & "   " & Err.Description
End If
cmdDisconnect.Enabled = False
port3.Enabled = True
cmdConnect.Enabled = True
End Sub

Private Sub cmdDisconnect_Click()
ws1.Close
cmdDisconnect.Enabled = False
port3.Enabled = True
cmdConnect.Enabled = True
End Sub


Private Sub Form_Load()
optTCP = True
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 4890
Me.Width = 6840
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub ws1_ConnectionRequest(ByVal requestID As Long)
 If ws1.State <> sckClosed Then ws1.Close
 ws1.Accept (requestID)
 txtStatus.Text = "Connection..."
End Sub

Private Sub ws1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
ws1.GetData strData
txtStatus.Text = txtStatus.Text & vbCrLf & " - " & strData
End Sub

Private Sub ws1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
txtStatus = "Winsock Error: " & Number & vbCrLf & "   " & descriptoin
End Sub
