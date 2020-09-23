VERSION 5.00
Begin VB.Form Tracefrm 
   BackColor       =   &H80000012&
   Caption         =   "Sapient2003 Net Tools - Trace Route"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   Icon            =   "Tracefrm.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "Tracefrm.frx":0442
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmTop 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   600
      TabIndex        =   4
      Top             =   840
      Width           =   4935
      Begin VB.CommandButton Close 
         Caption         =   "Close"
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton TraceRT2 
         Caption         =   "Trace Route"
         Default         =   -1  'True
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Host 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblIPHost 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "IP/Host:"
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
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblIP 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "IP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   240
         Left            =   3000
         TabIndex        =   7
         Top             =   120
         Width           =   225
      End
      Begin VB.Label IP 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3360
         TabIndex        =   5
         Top             =   120
         Width           =   1245
      End
   End
   Begin VB.Frame frmBottom 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   840
      TabIndex        =   6
      Top             =   1800
      Width           =   4455
      Begin VB.TextBox Response 
         Height          =   2295
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   4455
      End
   End
End
Attribute VB_Name = "Tracefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim TotalNum As Long
Dim KeepGoing As Integer
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

Public Sub GetRCode()
RCode = ""
DoEvents
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
'    If pIPe.Status = 11013 Then RCode = "TTL Exprd In Transit"
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
'    RCode = RCode + " (" + CStr(pIPe.Status) + ")"


    DoEvents

        If RCode <> "" Then
            If RCode = "Reqested Timed Out" Then
                vbWSACleanup
                If TotalNum < 10 Then Response.Text = Response.Text + " # 0" & TotalNum Else Response.Text = Response.Text + " # " & TotalNum
                Response.Text = Response.Text + " " & RCode & vbCrLf
                Response.SelStart = Len(Response)
            Exit Sub
            End If
            If RCode = "Success" Then
                vbWSACleanup
                If TotalNum < 10 Then Response.Text = Response.Text + " # 0" & TotalNum Else Response.Text = Response.Text + " # " & TotalNum
                Response.Text = Response.Text + " " + IP + vbCrLf
                Response.SelStart = Len(Response)
            Exit Sub
            End If
            KeepGoing = 1
            Response.Text = Response.Text & RCode
        Else
            If TTL - 1 < 10 Then Response.Text = Response.Text + " # 0" & TotalNum Else Response.Text = Response.Text + " # " & TotalNum
            Response.Text = Response.Text + " " + RespondingHost + vbCrLf
        End If
        Response.SelStart = Len(Response)
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
        CopyMemory IPLong, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        CopyMemory Addr, ByVal ListAddr, 4
        IP.Caption = Trim$(CStr(Asc(IPLong.Byte4)) + "." + CStr(Asc(IPLong.Byte3)) _
        + "." + CStr(Asc(IPLong.Byte2)) + "." + CStr(Asc(IPLong.Byte1)))
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


Public Sub vbIcmpSendEcho()
    vbWSACleanup
    Dim NbrOfPkts As Integer
    For NbrOfPkts = 1 To Trim$(1)


        DoEvents
        vbWSACleanup
            bReturn = IcmpSendEcho(hIP, Addr, szBuffer, Len(szBuffer), pIPo, pIPe, Len(pIPe) + 8, 2700)
            If bReturn Then
                TotalNum = TotalNum + 1
                RespondingHost = CStr(pIPe.Address(0)) + "." + CStr(pIPe.Address(1)) + "." + CStr(pIPe.Address(2)) + "." + CStr(pIPe.Address(3))
                GetRCode
            Else
                TotalNum = TotalNum + 1
                    GetRCode
                    TTL = TTL + 1
            End If
        Next NbrOfPkts
    End Sub


Sub vbWSAStartup()
Dim wsdaata As WSAdata
    iReturn = WSAStartup(&H101, WSAdata)


    If iReturn <> 0 Then ' If WSock32 error, then tell me about it
        MsgBox "WSock32.dll is Not responding!", 0, ""
    End If


    If LoByte(WSAdata.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAdata.wVersion) = WS_VERSION_MAJOR And HiByte(WSAdata.wVersion) < WS_VERSION_MINOR) Then
        sHighByte = Trim$(Str$(HiByte(WSAdata.wVersion)))
        sLowByte = Trim$(Str$(LoByte(WSAdata.wVersion)))
        sMsg = "WinSock Version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is Not supported "
        MsgBox sMsg
        End
    End If


    If WSAdata.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
            MsgBox sMsg
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

Private Sub Close_Click()
Unload Me
End Sub

Private Sub Form_Load()
vbWSAStartup
vbWSACleanup
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 4920
Me.Width = 6855
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
KeepGoing = 1

End Sub

Private Sub TraceRT2_Click()
TotalNum = 0
Response.Enabled = True
    szBuffer = Space(32)
    Response.Text = ""
    vbWSAStartup


    If Len(Host.Text) = 0 Then
        vbGetHostName
    End If
    vbGetHostByName
    vbIcmpCreateFile
    ' The following determines the TTL of th
    '     e ICMPEcho for TRACE function
    TraceRT = True
    Response.Text = Response.Text + "Tracing Route To " + IP.Caption + ":" + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)

    For TTL = 2 To 255
        If KeepGoing = 1 Then
        KeepGoing = 0
        Exit For
        End If
        pIPo.TTL = TTL
        DoEvents
        vbIcmpSendEcho


        DoEvents

            If RespondingHost = IP.Caption Then
                Response.Text = Response.Text + vbCrLf + "Route Trace has Completed"
                Exit For ' Stop TraceRT
            End If
        Next TTL
        Response.SelStart = Len(Response)
        TraceRT = False
        vbIcmpCloseHandle
        vbWSACleanup
End Sub
