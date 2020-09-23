VERSION 5.00
Begin VB.Form Pingfrm 
   Caption         =   "Sapient2003 Net Tools - Ping"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   Icon            =   "PingFrm.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "PingFrm.frx":0442
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmTop 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   4815
      Begin VB.CommandButton Close 
         Caption         =   "Close"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.HScrollBar ScrollPacket 
         Height          =   255
         Left            =   3000
         Max             =   5000
         Min             =   1
         TabIndex        =   3
         Top             =   720
         Value           =   1
         Width           =   1095
      End
      Begin VB.HScrollBar ScrollTimes 
         Height          =   255
         Left            =   3000
         Max             =   500
         Min             =   1
         TabIndex        =   1
         Top             =   360
         Value           =   1
         Width           =   1095
      End
      Begin VB.TextBox Host 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   360
         TabIndex        =   0
         Top             =   330
         Width           =   1455
      End
      Begin VB.CommandButton cmdPing 
         Caption         =   "Ping"
         Default         =   -1  'True
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblPacketSize 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "32"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   2760
         TabIndex        =   12
         Top             =   720
         Width           =   180
      End
      Begin VB.Label lblPacket 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Packet:"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   2160
         TabIndex        =   11
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lblPingTimes 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "1"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   2760
         TabIndex        =   10
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblPings 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Ping(s):"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   2160
         TabIndex        =   9
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lblIpHost 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Ip/Host:"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   840
         TabIndex        =   8
         Top             =   120
         Width           =   585
      End
   End
   Begin VB.Frame frmBottom 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   6135
      Begin VB.TextBox txtStatus 
         Height          =   2295
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   5775
      End
   End
End
Attribute VB_Name = "Pingfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PingTimes As Integer
Dim Speed As Long
Dim IP As String
Dim KeepGoing As Integer
Dim TotalNum As Long
Dim iReturn As Long, sLowByte As String, sHighByte As String
Dim sMsg As String, HostLen As Long
Dim Hostent As Hostent, PointerToPointer As Long, ListAddress As Long
Dim WSAdata As WSAdata, DotA As Long, DotAddr As String, ListAddr As Long
Dim MaxUDP As Long, MaxSockets As Long, i As Integer
Dim Description As String, Status As String
Dim ExitTheFor As Integer
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

Private Sub cmdPing_Click()
    Speed = 0
    PingTimes = 0
    cmdPing.Enabled = False
    ScrollTimes.Enabled = False
    ScrollPacket.Enabled = False
    txtStatus = ""
    szBuffer = Space(Val(lblPacketSize))
    vbWSAStartup
    If Len(Host.Text) = 0 Then
        vbGetHostName
    End If
    vbGetHostByName
    vbIcmpCreateFile
    pIPo2.TTL = Trim$(255)
    '
    For Times = 1 To lblPingTimes
    If ExitTheFor = 1 Then ExitTheFor = 0: Exit For
    vbIcmpSendEcho
    Next
    vbIcmpCloseHandle
    vbWSACleanup
    ScrollTimes.Enabled = True
    ScrollPacket.Enabled = True
    cmdPing.Enabled = True
    On Error GoTo skipit
    Speed = Speed / PingTimes
    txtStatus = txtStatus & vbCrLf & " Average Speed: " & Speed & "."
    txtStatus.SelStart = Len(txtStatus)
    Exit Sub
skipit:
End Sub

Public Sub GetRCode()
RCode = ""
    If pIPe.Status = 0 Then RCode = "Success"
    If pIPe.Status = 11001 Then RCode = "Buffer too Small"
    If pIPe.Status = 11002 Then RCode = "Destination Unreahable"
    If pIPe.Status = 11003 Then RCode = "Dest Host Not Reachable"
    If pIPe.Status = 11004 Then RCode = "Dest Protocol Not Reachable"
    If pIPe.Status = 11005 Then RCode = "Dest Port Not Reachable"
    If pIPe.Status = 11006 Then RCode = "No Resources Available"
    If pIPe.Status = 11007 Then RCode = "Bad Option"
    If pIPe.Status = 11008 Then RCode = "Hardware Error"
    If pIPe.Status = 11009 Then RCode = "Packet too Big"
    If pIPe.Status = 11010 Then RCode = "Reqested Timed Out"
    If pIPe.Status = 11011 Then RCode = "Bad Request"
    If pIPe.Status = 11012 Then RCode = "Bad Route"
    If pIPe.Status = 11014 Then RCode = "TTL Exprd Reassemb"
    If pIPe.Status = 11015 Then RCode = "Parameter Problem"
    If pIPe.Status = 11016 Then RCode = "Source Quench"
    If pIPe.Status = 11017 Then RCode = "Option too Big"
    If pIPe.Status = 11018 Then RCode = "Bad Destination"
    If pIPe.Status = 11019 Then RCode = "Address Deleted"
    If pIPe.Status = 11020 Then RCode = "Spec MTU Change"
    If pIPe.Status = 11021 Then RCode = "MTU Change"
    If pIPe.Status = 11022 Then RCode = "Unload"
    If pIPe.Status = 11050 Then RCode = "General Failure"

    DoEvents

        If RCode <> "" Then
            If RCode = "Success" Then
                Speed = Speed + Val(Trim$(CStr(pIPe2.RoundTripTime)))
                txtStatus.Text = txtStatus.Text + " Reply from " + RespondingHost + ": Bytes = " + Trim$(CStr(pIPe2.DataSize)) + " RTT = " + Trim$(CStr(pIPe2.RoundTripTime)) + "ms TTL = " + Trim$(CStr(pIPe2.Options.TTL)) + vbCrLf
                txtStatus.SelStart = Len(txtStatus)
            Exit Sub
            End If
            KeepGoing = 1
            txtStatus.Text = txtStatus.Text & RCode
        Else
            KeepGoing = 1
            txtStatus.Text = txtStatus.Text & RCode
        End If
        txtStatus.SelStart = Len(txtStatus)
    End Sub


Public Sub vbGetHostByName()
    Dim szString As String
    Host = Trim$(Host.Text)
    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))

    If gethostbyname(Host) = SOCKET_ERROR Then
        sMsg = "Winsock Error" & Str$(WSAGetLastError())
        txtStatus = sMsg
        ExitTheFor = 1
    Else
        PointerToPointer = gethostbyname(Host) ' Get the pointer to the address of the winsock hostent structure
        CopyMemory Hostent.h_name, ByVal _
        PointerToPointer, Len(Hostent) ' Copy Winsock structure to the VisualBasic structure
        ListAddress = Hostent.h_addr_list ' Get the ListAddress of the Address List
        CopyMemory ListAddr, ByVal ListAddress, 4 ' Copy Winsock structure To the VisualBasic structure
        CopyMemory IPLong2, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        CopyMemory Addr, ByVal ListAddr, 4
        IP = Trim$(CStr(Asc(IPLong2.Byte4)) + "." + CStr(Asc(IPLong2.Byte3)) _
        + "." + CStr(Asc(IPLong2.Byte2)) + "." + CStr(Asc(IPLong2.Byte1)))
    End If
End Sub


Public Sub vbGetHostName()
    
    Host = String(64, &H0)
    


    If gethostname(Host, HostLen) = SOCKET_ERROR Then
        sMsg = "WSock32 Error" & Str$(WSAGetLastError())
        txtStatus = sMsg
        ExitTheFor = 1
    Else
        Host = Left$(Trim$(Host), Len(Trim$(Host)) - 1)
        Host.Text = Host
    End If
End Sub


Public Sub vbIcmpSendEcho()
    Dim NbrOfPkts As Integer
    For NbrOfPkts = 1 To Trim$(1)

        DoEvents
            bReturn = IcmpSendEcho(hIP, Addr, szBuffer, Len(szBuffer), pIPo2, pIPe2, Len(pIPe2) + 8, 2700)
            If bReturn Then
                If KeepGoing = 1 Then KeepGoing = 0: Exit For
                PingTimes = PingTimes + 1
                RespondingHost = CStr(pIPe2.Address(0)) + "." + CStr(pIPe2.Address(1)) + "." + CStr(pIPe2.Address(2)) + "." + CStr(pIPe2.Address(3))
                GetRCode
            Else
                txtStatus.Text = txtStatus.Text + " Request Timeout" + vbCrLf
                txtStatus.SelStart = Len(txtStatus)
            End If
        Next NbrOfPkts
    End Sub


Sub vbWSAStartup()
Dim wsdaata As WSAdata
    iReturn = WSAStartup(&H101, WSAdata)


    If iReturn <> 0 Then ' If WSock32 error, then tell me about it
        txtStatus = "WSock32.dll is Not responding!"
        ExitTheFor = 1
    End If


    If LoByte(WSAdata.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAdata.wVersion) = WS_VERSION_MAJOR And HiByte(WSAdata.wVersion) < WS_VERSION_MINOR) Then
        sHighByte = Trim$(Str$(HiByte(WSAdata.wVersion)))
        sLowByte = Trim$(Str$(LoByte(WSAdata.wVersion)))
        sMsg = "WinSock Version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is Not supported "
        txtStatus = sMsg
        ExitTheFor = 1
        End
    End If


    If WSAdata.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
            txtStatus = sMsg
            ExitTheFor = 1
        End
    End If
    
    MaxSockets = WSAdata.iMaxSockets


    If MaxSockets < 0 Then
        MaxSockets = 65536 + MaxSockets
    End If
    MaxUDP = WSAdata.iMaxUdpDg


    If MaxUDP < 0 Then
        MaxUDP = 65536 + MaxUDP
    End If
    
    Description = ""


    For i = 0 To WSADESCRIPTION_LEN
        If WSAdata.szDescription(i) = 0 Then Exit For
        Description = Description + Chr$(WSAdata.szDescription(i))
    Next i
    Status = ""


    For i = 0 To WSASYS_STATUS_LEN
        If WSAdata.szSystemStatus(i) = 0 Then Exit For
        Status = Status + Chr$(WSAdata.szSystemStatus(i))
    Next i
End Sub


Public Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function


Public Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function


Public Sub vbWSACleanup()
    iReturn = WSACleanup()
End Sub


Public Sub vbIcmpCloseHandle()
    bReturn = IcmpCloseHandle(hIP)
End Sub


Public Sub vbIcmpCreateFile()
    hIP = IcmpCreateFile()
End Sub


Private Sub Form_Load()
ScrollPacket.Value = 32
vbWSAStartup
vbWSACleanup
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 4905
Me.Width = 6870
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub ScrollPacket_Change()
lblPacketSize = ScrollPacket.Value
End Sub

Private Sub ScrollTimes_Change()
lblPingTimes = ScrollTimes.Value
End Sub

