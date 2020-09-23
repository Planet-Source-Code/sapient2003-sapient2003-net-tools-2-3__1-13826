VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FingerFrm 
   Caption         =   "Sapient2003 Net Tools - Finger"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   Icon            =   "FingerFrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FingerFrm.frx":0442
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   448
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   1080
      TabIndex        =   7
      Top             =   2280
      Width           =   4455
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   2880
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.TextBox txtWhois 
         Height          =   1335
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   5
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Width           =   4455
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2760
         TabIndex        =   1
         Text            =   "00000"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Close 
         Caption         =   "Close"
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Host 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdFinger 
         Caption         =   "Finger"
         Default         =   -1  'True
         Height          =   255
         Left            =   3480
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Hosttxt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label3 
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
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "Server:"
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
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "User:"
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
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
   End
End
Attribute VB_Name = "FingerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IP As String

Dim iReturn As Long, sLowByte As String, sHighByte As String
Dim sMsg As String, HostLen As Long
Dim Hostent As Hostent, PointerToPointer As Long, ListAddress As Long
Dim WSAdata As WSAdata, DotA As Long, DotAddr As String, ListAddr As Long
Dim MaxUDP As Long, MaxSockets As Long, i As Integer
Dim Description As String, Status As String
' Ping Variables
Dim bReturn As Boolean, hIP As Long
Dim szBuffer As String
Dim Addr As Long
Dim RCode As String
Dim RespondingHost As String
' TRACERT Variables
Dim TraceRT As Boolean
Dim TTL As Integer
' WSock32 Constants
Const WS_VERSION_MAJOR = &H101 \ &H100 And &HFF&
Const WS_VERSION_MINOR = &H101 And &HFF&
Const MIN_SOCKETS_REQD = 0

Private Sub Close_Click()
Unload Me
End Sub

Private Sub cmdFinger_Click()
On Error Resume Next
If Winsock1.State <> sckClosed Then
    Winsock1.Close
End If
txtWhois.Text = ""
vbGetHostByName
Winsock1.Connect IP, txtPort.Text

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 4830
Me.Width = 6840
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
On Error Resume Next
Winsock1.SendData ("/W " & Hosttxt.Text & vbCrLf)

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim dataA
Winsock1.GetData dataA, vbString
txtWhois.Text = txtWhois.Text & dataA '& vbCrLf
Dim counter As Long
counter = 1
start:
   Dim Search, where   ' Declare variables.
   ' Get search string from user.
   Search = Chr$(10)
   where = InStr(counter, txtWhois.Text, Search, vbTextCompare) ' Find string in text.
   'MsgBox Where
   If where Then   ' If found,
      txtWhois.SelStart = where - 1   ' set selection start and
      txtWhois.SelLength = Len(Search)
      txtWhois.SelText = vbCrLf
      counter = where + txtWhois.SelLength + 2 ': 'MsgBox counter
   Else
      Exit Sub  ' Notify user.
   End If

GoTo start
Winsock1.Close
End Sub

Public Sub vbGetHostByName()
    Dim szString As String
    Host = Trim$(Host.Text)
    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))
DoEvents
    If gethostbyname(Host) = SOCKET_ERROR Then
        sMsg = "Winsock Error" & Str$(WSAGetLastError())
        MsgBox sMsg, 0, ""
    Else
DoEvents
        PointerToPointer = gethostbyname(Host) ' Get the pointer to the address of the winsock hostent structure
        CopyMemory Hostent.h_name, ByVal _
        PointerToPointer, Len(Hostent) ' Copy Winsock structure to the VisualBasic structure
        ListAddress = Hostent.h_addr_list ' Get the ListAddress of the Address List
        CopyMemory ListAddr, ByVal ListAddress, 4 ' Copy Winsock structure To the VisualBasic structure
        CopyMemory IPLong7, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        CopyMemory Addr, ByVal ListAddr, 4
        IP = Trim$(CStr(Asc(IPLong7.Byte4)) + "." + CStr(Asc(IPLong7.Byte3)) _
        + "." + CStr(Asc(IPLong7.Byte2)) + "." + CStr(Asc(IPLong7.Byte1)))
    End If
End Sub


