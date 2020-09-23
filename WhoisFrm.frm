VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form WhoisFrm 
   Caption         =   "Sapient2003 Net Tools - Whois"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   Icon            =   "WhoisFrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "WhoisFrm.frx":0442
   ScaleHeight     =   299
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox ServerTxt 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000014&
      Height          =   1695
      Left            =   1200
      TabIndex        =   6
      Top             =   2520
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
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   4455
      Begin VB.CommandButton Close 
         Caption         =   "Close"
         Height          =   255
         Left            =   3360
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdWhois 
         Caption         =   "Whois"
         Default         =   -1  'True
         Height          =   255
         Left            =   3360
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Host 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   2055
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
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Domain:"
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
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
   End
End
Attribute VB_Name = "WhoisFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Close_Click()
Unload Me
End Sub

Private Sub cmdWhois_Click()
Winsock1.Close
Dim WhoisStr As String
txtWhois.Text = ""
Winsock1.Connect ServerTxt, 43
End Sub

Private Sub Form_Load()
ServerTxt.AddItem "127.0.0.1"
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 4890
Me.Width = 6855
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
On Error Resume Next
Winsock1.SendData ("whois " & Host.Text & vbCrLf)
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





