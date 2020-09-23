Attribute VB_Name = "Closemodul1"
#If Win32 Then
Private Declare Function ShowOwnedPopups& Lib "user32" (ByVal hwnd As Long, ByVal fShow As Long)
#Else
Private Declare Sub ShowOwnedPopups Lib "User" (ByVal hwnd As Integer, ByVal fShow As Integer)
#End If 'WIN32
#If Win32 Then
Public Declare Function ShowWindow& Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long)
#Else
Public Declare Function ShowWindow% Lib "User" (ByVal hwnd As Integer, ByVal nCmdShow As Integer)
#End If 'WIN32
#If Win32 Then
Public Declare Function CreateMDIWindow& Lib "user32" Alias "CreateMDIWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hInstance As Long, ByVal lParam As Long)
Public Declare Function CreateMenu& Lib "user32" ()
Public Declare Function CreatePopupMenu& Lib "user32" ()
Public Declare Function CreateWindow& Lib "user32" Alias "CreateWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any)
#Else
'Function CreateMDIWindow is not available in the WIN16 API.
Public Declare Function CreateMenu% Lib "User" ()
Public Declare Function CreatePopupMenu% Lib "User" ()
Public Declare Function CreateWindow% Lib "User" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hWndParent As Integer, ByVal hMenu As Integer, ByVal hInstance As Integer, ByVal lpParam As String)
#End If 'WIN32
#If Win32 Then
Public Declare Function SetClassWord& Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long)
#Else
Public Declare Function SetClassWord% Lib "User" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal wNewWord As Integer)
#End If 'WIN32
#If Win32 Then
Public Declare Function DestroyWindow& Lib "user32" (ByVal hwnd As Long)
#Else
Public Declare Function DestroyWindow% Lib "User" (ByVal hwnd As Integer)
#End If 'WIN32
#If Win32 Then
Public Declare Function EnableWindow& Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long)
#Else
Public Declare Function EnableWindow% Lib "User" (ByVal hwnd As Integer, ByVal aBOOL As Integer)
#End If 'WIN32
#If Win32 Then
Public Declare Function SetFocusApi& Lib "user32" Alias "SetFocus" (ByVal hwnd As Long)
#Else
Public Declare Function SetFocusApi% Lib "User" Alias "SetFocus" (ByVal hwnd As Integer)
#End If 'WIN32
#If Win32 Then
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)
#Else
Public Declare Function WritePrivateProfileString% Lib "kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)
Public Declare Function GetPrivateProfileString% Lib "kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String)
#End If
#If Win32 Then
Public Declare Function SetSysColors& Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long)
Public Declare Function GetSysColor& Lib "user32" (ByVal nIndex As Long)
#Else
Public Declare Sub SetSysColors Lib "User" (ByVal nChanges As Integer, lpSysColor As Integer, lpColorValues As Long)
Public Declare Function GetSysColor& Lib "User" (ByVal nIndex As Integer)
#End If
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function gettopwindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
#If Win32 Then
Public Const SND_ASYNC& = &H1
Public Const SND_SYNC& = &H0
#Else
Public Const SND_ASYNC% = &H1
Public Const SND_SYNC% = &H0
#End If
#If Win32 Then
Public Declare Function sndPlaySound& Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long)
#Else
Public Declare Function sndPlaySound% Lib "mmsystem.dll" (ByVal lpszSoundName As String, ByVal uFlags As Integer)
#End If
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_gettext = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3


Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETtext = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
#If Win32 Then
Public Const WS_VISIBLE& = &H10000000
#Else
Public Const WS_VISIBLE& = &H10000000
#End If 'WIN32
#If Win32 Then
Public Const WM_TIMER& = &H113
#Else
Public Const WM_TIMER& = &H113
#End If 'WIN32
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
End Type




Function FindChildByClass(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
room% = firs%
FindChildByClass = room%

End Function

Function FindChildByTitle(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitle = 0

bone:
room% = firs%
FindChildByTitle = room%
End Function

Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
A% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function

Function AOLGetChat()

child = FindChildByClass(childs%, "_AOL_View")

GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)

theview$ = TrimSpace$
AOLGetChat = theview$
End Function


Public Function stayontop(the As Form)
Flag% = SWP_NOMOVE Or SWP_NOSIZE
lSetPos = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Function

Public Function PlayWav(FileName)
Flagss% = SND_ASNC Or SND_SYNC
Play2 = sndPlaySound(FileName, Flagss%)
End Function

Public Function SetText(Text)
lSetText = SendMessageByString(Wind&, WM_SETTEXT, 0, Text)
End Function

Public Function Click(Window%)
lClick& = SendMessage(Window%, WM_LBUTTONDOWN, 0, 0&)
lClick& = SendMessage(Window%, WM_LBUTTONUP, 0, 0&)
End Function

Public Function ChatSend(Text)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = FindChildByClass(firs%, "_AOL_Listbox")
listerb% = FindChildByClass(firs%, "_AOL_Combobox")
If listers% And listere% And listerb% Then GoTo bone

firs% = GetWindow(MDI%, GW_CHILD)
While firs%
firs% = GetWindow(firs%, 2)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = FindChildByClass(firs%, "_AOL_Listbox")
listerb% = FindChildByClass(firs%, "_AOL_Combobox")
If listers% And listere% And listerb% Then GoTo bone
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = FindChildByClass(firs%, "_AOL_Listbox")
listerb% = FindChildByClass(firs%, "_AOL_Combobox")
Listerc% = FindChildByClass(firs%, "_AOL_Glyph")
Listerd% = FindChildByClass(firs%, "_AOL_Static")
If listers% And listere% And listerb% Then GoTo bone Else Exit Function
Wend

Fuck% = listers% And listere% And listerb%
bone:
One1% = GetWindow(listers%, 2)
Two2% = GetWindow(One1%, 2)
Three3% = GetWindow(Two2%, 2)
Four4% = GetWindow(Three3%, 2)
Five5% = GetWindow(Four4%, 2)
six6% = GetWindow(Five5%, 2)
Seve7% = GetWindow(six6%, 2)
Eight8% = GetWindow(Seve7%, 2)

lSnd% = SendMessageByString(six6%, WM_SETTEXT, 0, Text)
'Click (Seve7%)
lEnter% = SendMessageByNum(six6%, WM_CHAR, 13, 0&)
End Function

Sub pause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Function AOLGetUser()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
A% = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = User
End Function

Public Sub AOLIcons()
AOL% = FindWindow("AOL Frame25", vbNullString)
TL1% = FindChildByClass(AOL%, "AOL Toolbar")
TL2% = FindChildByClass(TL1%, "_AOL_Toolbar")
ICO1% = FindChildByClass(TL2%, "_AOL_Icon")
ICO2% = GetWindow(ICO1%, 2)
ICO3% = GetWindow(ICO2%, 2)
ICO4% = GetWindow(ICO3%, 2)
ICO5% = GetWindow(ICO4%, 2)
ICO6% = GetWindow(ICO5%, 2)
ICO7% = GetWindow(ICO6%, 2)
ICO8% = GetWindow(ICO7%, 2)
ICO9% = GetWindow(ICO8%, 2)
ICO10% = GetWindow(ICO9%, 2)
ICO11% = GetWindow(ICO10%, 2)
ICO12% = GetWindow(ICO11%, 2)
ICO13% = GetWindow(ICO12%, 2)
ICO14% = GetWindow(ICO13%, 2)
ICO15% = GetWindow(ICO14%, 2)
ICO16% = GetWindow(ICO15%, 2)
ICO17% = GetWindow(ICO16%, 2)
ICO18% = GetWindow(ICO17%, 2)
ICO19% = GetWindow(ICO18%, 2)
ICO20% = GetWindow(ICO19%, 2)
ICO21% = GetWindow(ICO20%, 2)
ICO22% = GetWindow(ICO21%, 2)
ICO23% = GetWindow(ICO22%, 2)
ICO24% = GetWindow(ICO23%, 2)
ICO25% = GetWindow(ICO24%, 2)
ICO26% = GetWindow(ICO25%, 2)
ICO27% = GetWindow(ICO26%, 2)
End Sub

Public Function IMAnswer(Text)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM1% = FindChildByTitle(MDI%, ">Instant Message From: ")
IM2% = FindChildByClass(IM1%, "RICHCNTL")
IM3% = GetWindow(IM2%, 2)
IM4% = GetWindow(IM3%, 2)
IM5% = GetWindow(IM4%, 2)
IM6% = GetWindow(IM5%, 2)
IM7% = GetWindow(IM6%, 2)
IM8% = GetWindow(IM7%, 2)
IM9% = GetWindow(IM8%, 2)
IM10% = GetWindow(IM9%, 2)
IM11% = GetWindow(IM10%, 2)
IM12% = GetWindow(IM11%, 2)
IM13% = GetWindow(IM12%, 2)
IM14% = GetWindow(IM13%, 2)
IM15% = GetWindow(IM14%, 2)
IM16% = GetWindow(IM15%, 2)
IM17% = GetWindow(IM16%, 2)
SNDTX% = SendMessageByString(IM15%, WM_SETTEXT, 0, Text)
Click (IM16%)
CLOSE1% = SendMessage(IM1%, WM_CLOSE, 0, 0&)
End Function

Sub runmenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

End Sub

Public Sub RunMenuByString(Application, StringSearch)

ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)

For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For GetString = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, GetString)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If

Next GetString

Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub

Function ChatSend2(Text)
AOL = FindWindow("AOL Frame25", vbNullString)
If AOL = 0 Then Exit Function
f = FindChildByClass(AOL, "MDIClient")
b = FindChildByClass(f, "AOL Child")

start:
c = FindChildByClass(b, "RICHCNTL")
If c = 0 Then GoTo nextwnd
d = FindChildByClass(b, "_AOL_Combobox")
If d = 0 Then GoTo nextwnd
e = FindChildByClass(b, "_AOL_Listbox")
If e = 0 Then GoTo nextwnd
findchatroom = b
Exit Function

nextwnd:
b = GetWindow(b, 2)
If b = GetWindow(b, GW_HWNDLAST) Then Exit Function
FNL% = FindChildByClass(b, "RICHCNTL")
If FNL% <> 0 Then MsgBox "Found It!"
One1% = GetWindow(FNL%, 2)
Two2% = GetWindow(One1%, 2)
Three3% = GetWindow(Two2%, 2)
Four4% = GetWindow(Three3%, 2)
Five5% = GetWindow(Four4%, 2)
six6% = GetWindow(Five5%, 2)
Seve7% = GetWindow(six6%, 2)
Eight8% = GetWindow(Seve7%, 2)

lSnd% = SendMessageByString(six6%, WM_SETTEXT, 0, Text)
'Click (Seve7%)
lEnter% = SendMessageByNum(six6%, WM_CHAR, 13, 0&)
GoTo start
End Function

Public Function Click2(hwnd)
lEnter% = SendMessageByNum(hwnd, WM_CHAR, 13, 0&)
End Function
