VERSION 5.00
Begin VB.Form help 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   Picture         =   "help.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Text            =   "Help:"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox help1 
      Height          =   2775
      Left            =   960
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   4815
   End
End
Attribute VB_Name = "help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
