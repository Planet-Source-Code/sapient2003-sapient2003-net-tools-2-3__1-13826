VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Scannerfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Scanner"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   Icon            =   "Scannerfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "Scannerfrm.frx":0442
   ScaleHeight     =   4500
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   2040
      TabIndex        =   19
      Text            =   "Ports Scanned:"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Portn 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   2160
      TabIndex        =   18
      Text            =   "0"
      Top             =   4050
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000009&
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Text            =   "Open Ports:"
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Close 
      Caption         =   "Close"
      Height          =   255
      Left            =   5160
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000009&
      Height          =   195
      Left            =   4800
      TabIndex        =   14
      Text            =   "Ports to scan:"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000009&
      Height          =   195
      Left            =   2400
      TabIndex        =   13
      Text            =   "IP Address:"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000009&
      Height          =   195
      Left            =   5160
      TabIndex        =   12
      Text            =   "Max Connections:"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdClearList 
      Caption         =   "Clear List"
      Height          =   255
      Left            =   5160
      TabIndex        =   17
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame fraMaxConnections 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
      Begin VB.TextBox txtMaxConnections 
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "1"
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdStop 
      Cancel          =   -1  'True
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Default         =   -1  'True
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame fraOpenPorts 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Open Ports"
      Height          =   2295
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   4695
      Begin VB.ListBox lstOpenPorts 
         Height          =   2205
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   4695
      End
   End
   Begin VB.Timer timTimer 
      Interval        =   100
      Left            =   1080
      Top             =   1080
   End
   Begin MSWinsockLib.Winsock wskSocket 
      Index           =   0
      Left            =   1200
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraScanPorts 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
      Begin VB.TextBox txtUpperBound 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "65535"
         Top             =   0
         Width           =   615
      End
      Begin VB.TextBox txtLowerBound 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   0
         MaxLength       =   5
         TabIndex        =   3
         Text            =   "1"
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblTo 
         BackColor       =   &H80000008&
         Caption         =   "To"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame fraRemoteIP 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Text            =   "127.0.0.1"
         Top             =   0
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Scannerfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClearList_Click()
   Me.lstOpenPorts.Clear
End Sub

Private Sub cmdScan_Click()

   Dim intI As Integer
   
   lngNextPort = Val(Me.txtLowerBound)
  
   For intI = 1 To Val(Me.txtMaxConnections)
   
      Load Me.wskSocket(intI)
     
      lngNextPort = lngNextPort + 1
      
      Me.wskSocket(intI).Connect Me.txtIP, lngNextPort
   
   Next intI

 cmdStop.Enabled = True

End Sub

Private Sub cmdStop_Click()

   Dim intI As Integer
   
   For intI = 1 To Val(Me.txtMaxConnections)
   
      Me.wskSocket(intI).Close
 
      Unload Me.wskSocket(intI)
   
   Next intI
   
cmdStop.Enabled = False

End Sub

Private Sub timTimer_Timer()

   Me.Portn.Text = Str(lngNextPort)

End Sub

Private Sub wskSocket_Connect(Index As Integer)

   Me.lstOpenPorts.AddItem "Port: " + Str(Me.wskSocket(Index).RemotePort)
  
   Try_Next_Port (Index)

End Sub

Private Sub wskSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

   Try_Next_Port (Index)

End Sub

Private Sub Try_Next_Port(Index As Integer)

   Me.wskSocket(Index).Close

   If lngNextPort < Val(Me.txtUpperBound) Then
      
      Me.wskSocket(Index).Connect , lngNextPort
      
      lngNextPort = lngNextPort + 1

   Else

      Unload Me.wskSocket(Index)

   End If

End Sub

Private Sub Close_Click()
Unload Me
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 4875
Me.Width = 6795
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub
