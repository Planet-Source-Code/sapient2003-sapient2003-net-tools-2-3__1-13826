VERSION 5.00
Begin VB.Form frmProgramCloser 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Closer"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmClose.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmClose.frx":0442
   ScaleHeight     =   4485
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Close 
      Caption         =   "Close"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3600
      Top             =   2880
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&End Selected Task"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   3600
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "frmClose.frx":153A
      Left            =   360
      List            =   "frmClose.frx":153C
      TabIndex        =   1
      Top             =   840
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Show Running Tasks"
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   3600
      Width           =   1935
   End
End
Attribute VB_Name = "frmProgramCloser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindow Lib "user32" _
(ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetParent Lib "user32" _
(ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib _
"user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" _
Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal _
lpString As String, ByVal cch As Long) As Long
Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2
Sub LoadTaskList()
Dim CurrWnd As Long
Dim Length As Long
Dim TaskName As String
Dim parent As Long

List1.Clear
CurrWnd = GetWindow(frmProgramCloser.hwnd, GW_HWNDFIRST)

While CurrWnd <> 0
parent = GetParent(CurrWnd)
Length = GetWindowTextLength(CurrWnd)
TaskName = Space$(Length + 1)
Length = GetWindowText(CurrWnd, TaskName, Length + 1)
TaskName = Left$(TaskName, Len(TaskName) - 1)

If Length > 0 Then
    If TaskName <> Me.Caption Then
        If TaskName <> "taskmon" Then
            List1.AddItem TaskName
        End If
    End If
End If
CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
DoEvents

Wend

End Sub

Private Sub Command1_Click()
LoadTaskList
End Sub

Private Sub Command2_Click()
On Error GoTo erlevel
Dim winHwnd As Long
Dim RetVal As Long
winHwnd = FindWindow(vbNullString, List1.Text)
Debug.Print winHwnd
If winHwnd <> 0 Then
RetVal = PostMessage(winHwnd, &H10, 0&, 0&)
If RetVal = 0 Then
MsgBox "Error posting message."
End If
Else: MsgBox List1.Text + " is not open."
End If
erlevel:
LoadTaskList
End Sub

Private Sub Close_Click()
Unload Me
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 4860
Me.Width = 6825
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub Form_Load()
stayontop Me
End Sub

Private Sub Timer1_Timer()
If List1.Text = "" Then
    Command2.Enabled = False
Else
    Command2.Enabled = True
End If
End Sub

