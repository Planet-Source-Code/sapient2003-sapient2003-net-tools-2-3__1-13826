VERSION 5.00
Begin VB.Form ICQCrash 
   Caption         =   "ICQ Crasher"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   Picture         =   "ICQCrash.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Close 
      Caption         =   "Close"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Crash 
      Caption         =   "Crash"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox IPA 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   2880
      TabIndex        =   0
      Text            =   "IP Address:"
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "ICQCrash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

