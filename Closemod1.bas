Attribute VB_Name = "Closemodul"
Public Number As Integer
Public dlfile As String
Public password As String
Public register As String
Public ListBox As String
Public combobocs As String
Public winselect As String
Public diceno As Byte
Public dicelet As String
Public diceroll As Integer
Public whatintxt As String
Public Declare Function GetCurrentProcessId _
Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess _
Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess _
Lib "kernel32" (ByVal dwProcessId As Long, _
ByVal dwType As Long) As Long
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0
Public Const WM_CLOSE = &H10
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Declare Function PostMessage Lib "user32" Alias _
"PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long




Public Sub MakeMeService()
Dim pid As Long
Dim reserv As Long

pid = GetCurrentProcessId()
regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub
Public Sub UnMakeMeService()
Dim pid As Long
Dim reserv As Long

pid = GetCurrentProcessId()
regserv = RegisterServiceProcess(pid, _
RSP_UNREGISTER_SERVICE)
End Sub



