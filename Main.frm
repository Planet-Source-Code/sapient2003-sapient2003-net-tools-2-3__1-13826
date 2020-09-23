VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "Sapient2003 Net Tools"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "Main.frx":0442
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   360
      TabIndex        =   21
      Text            =   "Port Scanner"
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   5520
      TabIndex        =   20
      Text            =   "Version 2.3"
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   360
      TabIndex        =   19
      Text            =   "Proccesses"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   360
      TabIndex        =   18
      Text            =   "Updates:"
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Proccess 
      Caption         =   "Proccesses"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Speed 
      Caption         =   "Speed Check"
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Winsck 
      Caption         =   "About"
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton FingerBtn 
      Caption         =   "Finger"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton ListenBtn 
      Caption         =   "Listener"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton WhoisBtn 
      Caption         =   "Whois"
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Source 
      Caption         =   "Get HTML"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Mail 
      Caption         =   "Email"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Trace 
      Caption         =   "TraceRoute"
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Lookup 
      Caption         =   "Host Lookup"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Scanner 
      Caption         =   "Port Scan"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Ping 
      Caption         =   "Ping"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton ConnectBtn 
      Caption         =   "Raw Connect"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox HostTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   960
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3240
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000B&
      BorderWidth     =   2
      X1              =   8
      X2              =   440
      Y1              =   224
      Y2              =   224
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   224
      X2              =   224
      Y1              =   96
      Y2              =   224
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      BorderWidth     =   2
      X1              =   8
      X2              =   440
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Local Host:"
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
      Left            =   3480
      TabIndex        =   15
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label IPLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Local IP:"
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
      Left            =   360
      TabIndex        =   13
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ConnectBtn_Click()
Connect.Show
End Sub

Private Sub Exit_Click()
Unload Me
End
End Sub

Private Sub FingerBtn_Click()
FingerFrm.Show
End Sub

Private Sub Form_DblClick()
Me.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
Dim NameStr As String, SerStr As String
IPTxt.Text = Winsock1.LocalIP
Hosttxt.Text = Winsock1.LocalHostName
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub ListenBtn_Click()
ListenFrm.Show
End Sub

Private Sub ICQc_Click()
ICQCrash.Show
End Sub

Private Sub Lookup_Click()
LookupFrm.Show
End Sub

Private Sub Mail_Click()
EmailFrm.Show
End Sub

Private Sub Ping_Click()
Pingfrm.Show
End Sub

Private Sub Scanner_Click()
Scannerfrm.Show
End Sub

Private Sub Source_Click()
GetHTML.Show
End Sub

Private Sub Speed_Click()
SpeedChk.Show
End Sub

Private Sub Trace_Click()
Tracefrm.Show
End Sub

Private Sub WhoisBtn_Click()
WhoisFrm.Show
End Sub
Private Sub Proccess_Click()
frmProgramCloser.Show
End Sub

Private Sub Winsck_Click()
Aboutfrm.Show 1
End Sub
