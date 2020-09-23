VERSION 5.00
Begin VB.Form Aboutfrm 
   BorderStyle     =   0  'None
   Caption         =   "Sapient2003 Net Tools - About"
   ClientHeight    =   3420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   Icon            =   "Aboutfrm.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Aboutfrm.frx":0442
   ScaleHeight     =   3420
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Aboutfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

