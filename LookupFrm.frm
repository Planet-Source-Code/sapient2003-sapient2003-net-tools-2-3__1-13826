VERSION 5.00
Begin VB.Form LookupFrm 
   Caption         =   "Sapient2003 Net Tools - Host Lookup"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   Icon            =   "LookupFrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "LookupFrm.frx":0442
   ScaleHeight     =   299
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   447
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   600
      TabIndex        =   7
      Top             =   2640
      Width           =   4575
      Begin VB.TextBox Address 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "Host Address:"
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
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1080
      TabIndex        =   6
      Top             =   1920
      Width           =   3375
      Begin VB.CommandButton Close 
         Caption         =   "Close"
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Resolve 
         Caption         =   "Resolve Host"
         Default         =   -1  'True
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   4335
      Begin VB.TextBox Host 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2400
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Hostname or IP Address:"
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
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "LookupFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 4890
Me.Width = 6825
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub Resolve_Click()
On Error Resume Next
Address.Text = ""
If Len(Host.Text) = 0 Then
    vbGetHostName
End If
vbGetHostByName
End Sub



Public Sub vbGetHostByName()
    Dim szString As String
    Host = Trim$(Host.Text)
    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))

    If gethostbyname(Host) = SOCKET_ERROR Then
        sMsg = "Winsock Error" & Str$(WSAGetLastError())
        MsgBox sMsg, 0, ""
    Else
        PointerToPointer = gethostbyname(Host) ' Get the pointer to the address of the winsock hostent structure
        CopyMemory Hostent.h_name, ByVal _
        PointerToPointer, Len(Hostent) ' Copy Winsock structure to the VisualBasic structure
        ListAddress = Hostent.h_addr_list ' Get the ListAddress of the Address List
        CopyMemory ListAddr, ByVal ListAddress, 4 ' Copy Winsock structure To the VisualBasic structure
        CopyMemory IPLong3, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        CopyMemory Addr, ByVal ListAddr, 4
        Address.Text = Trim$(CStr(Asc(IPLong3.Byte4)) + "." + CStr(Asc(IPLong3.Byte3)) _
        + "." + CStr(Asc(IPLong3.Byte2)) + "." + CStr(Asc(IPLong3.Byte1)))
    End If
End Sub


Public Sub vbGetHostName()
    
    Host = String(64, &H0)
    


    If gethostname(Host, HostLen) = SOCKET_ERROR Then
        sMsg = "WSock32 Error" & Str$(WSAGetLastError())
        MsgBox sMsg, 0, ""
    Else
        Host = Left$(Trim$(Host), Len(Trim$(Host)) - 1)
        Host.Text = Host
    End If
End Sub
