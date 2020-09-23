VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form EmailFrm 
   BackColor       =   &H80000012&
   Caption         =   "Sapient2003 - Send Email"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   Icon            =   "EmailFrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "EmailFrm.frx":0442
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   446
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   255
      Left            =   6120
      TabIndex        =   8
      Top             =   5760
      Width           =   975
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   720
      TabIndex        =   19
      Top             =   3120
      Width           =   5055
      Begin VB.TextBox txtEmailSubject 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtEmailBodyOfMessage 
         Height          =   1575
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000012&
         Caption         =   "Subject:"
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
         Left            =   480
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   5895
      Begin VB.TextBox txtEmailServer 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000012&
         Caption         =   "SMTP Email Server:"
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
         TabIndex        =   18
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   6015
      Begin VB.TextBox ToNametxt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtToEmailAddress 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000012&
         Caption         =   "Name:"
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
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000012&
         Caption         =   "Send To:"
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
         TabIndex        =   15
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   6135
      Begin VB.TextBox txtFromName 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtFromEmailAddress 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "Name:"
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
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Your Email Address:"
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
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Status:"
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1200
      TabIndex        =   9
      Top             =   5520
      Width           =   4815
      Begin VB.Label StatusTxt 
         BackColor       =   &H80000012&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5760
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3600
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "EmailFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Response As String, Reply As Integer, DateNow As String
Dim first As String, Second As String, Third As String
Dim Fourth As String, Fifth As String, Sixth As String
Dim Seventh As String, Eighth As String
Dim start As Single, Tmr As Single



Sub SendEmail(MailServerName As String, FromName As String, FromEmailAddress As String, ToName As String, ToEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String)
          
    Winsock1.LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail pre program start
    
If Winsock1.State = sckClosed Then ' Check to see if socet is closed
    DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
    first = "mail from:" + Chr(32) + FromEmailAddress + vbCrLf ' Get who's sending E-Mail address
    Second = "rcpt to:" + Chr(32) + ToEmailAddress + vbCrLf ' Get who mail is going to
    Third = "Date:" + Chr(32) + DateNow + vbCrLf ' Date when being sent
    Fourth = "From:" + Chr(32) + FromName + vbCrLf ' Who's Sending
    Fifth = "To:" + Chr(32) + ToNametxt + vbCrLf ' Who it going to
    Sixth = "Subject:" + Chr(32) + EmailSubject + vbCrLf ' Subject of E-Mail
    Seventh = EmailBodyOfMessage + vbCrLf ' E-mail message body
    Ninth = "X-Mailer: EBT Reporter v 2.x" + vbCrLf ' What program sent the e-mail, customize this
    Eighth = Fourth + Third + Ninth + Fifth + Sixth  ' Combine for proper SMTP sending

    Winsock1.Protocol = sckTCPProtocol ' Set protocol for sending
    Winsock1.RemoteHost = MailServerName ' Set the server address
    Winsock1.RemotePort = 25 ' Set the SMTP Port
    Winsock1.Connect ' Start connection
    
    WaitFor ("220")
    
    StatusTxt.Caption = "Connecting...."
    StatusTxt.Refresh
    
    Winsock1.SendData ("HELO worldcomputers.com" + vbCrLf)

    WaitFor ("250")

    StatusTxt.Caption = "Connected"
    StatusTxt.Refresh

    Winsock1.SendData (first)

    StatusTxt.Caption = "Sending Message"
    StatusTxt.Refresh

    WaitFor ("250")

    Winsock1.SendData (Second)

    WaitFor ("250")

    Winsock1.SendData ("data" + vbCrLf)
    
    WaitFor ("354")


    Winsock1.SendData (Eighth + vbCrLf)
    Winsock1.SendData (Seventh + vbCrLf)
    Winsock1.SendData ("." + vbCrLf)

    WaitFor ("250")

    Winsock1.SendData ("quit" + vbCrLf)
    
    StatusTxt.Caption = "Disconnecting"
    StatusTxt.Refresh

    WaitFor ("221")

    Winsock1.Close
Else
    MsgBox (Str(Winsock1.State))
End If
   
End Sub
Sub WaitFor(ResponseCode As String)
    start = Timer ' Time event so won't get stuck in loop
    While Len(Response) = 0
        Tmr = start - Timer
        DoEvents ' Let System keep checking for incoming response **IMPORTANT**
        If Tmr > 50 Then ' Time in seconds to wait
            MsgBox "SMTP service error, timed out while waiting for response", 64, MsgTitle
            Exit Sub
        End If
    Wend
    While Left(Response, 3) <> ResponseCode
        DoEvents
        If Tmr > 50 Then
            MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64, MsgTitle
            Exit Sub
        End If
    Wend
Response = "" ' Sent response code to blank **IMPORTANT**
End Sub


Private Sub Command1_Click()
    SendEmail txtEmailServer.Text, txtFromName.Text, txtFromEmailAddress.Text, txtToEmailAddress.Text, txtToEmailAddress.Text, txtEmailSubject.Text, txtEmailBodyOfMessage.Text
    'MsgBox ("Mail Sent")
    StatusTxt.Caption = "Mail Sent"
    StatusTxt.Refresh
    Beep
    
    Close
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 6630
Me.Width = 7335
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Winsock1.GetData Response ' Check for incoming response *IMPORTANT*

End Sub
