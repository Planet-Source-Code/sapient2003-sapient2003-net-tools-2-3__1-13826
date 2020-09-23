Attribute VB_Name = "Functions"

Option Explicit
Public Const REG_NONE = 0 ' No value type
Public Const REG_SZ = 1 ' Unicode nul terminated string
Public Const REG_EXPAND_SZ = 2 ' Unicode nul terminated string
Public Const REG_BINARY = 3 ' Free form binary
Public Const REG_DWORD = 4 ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4 ' 32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN = 5 ' 32-bit number
Public Const REG_LINK = 6 ' Symbolic Link (unicode)
Public Const REG_MULTI_SZ = 7 ' Multiple Unicode strings
Public Const REG_RESOURCE_LIST = 8 ' Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9 ' Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10


Public Enum hKeyNames
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_ARENA_TRASHED = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259
Public Const KEY_ALL_ACCESS = &H3F
Public Const REG_OPTION_NON_VOLATILE = 0

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_SYSCOMMAND = &HA1
Public Const WM_MOVE = &O2


Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long


Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long


Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long


Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long


Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long


Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long


Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long


Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long


Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long


Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long


Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long


Private Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String


    Select Case lType
        Case REG_SZ
        sValue = vValue & Chr$(0)
        SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
        lValue = vValue
        SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
    End Select
End Function


Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    On Error GoTo QueryValueExError
    ' Determine the size and type of data to
    '     be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5


    Select Case lType
        ' For strings
        Case REG_SZ, REG_EXPAND_SZ:
        sValue = String(cch, 0)
        lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)


        If lrc = ERROR_NONE Then
            vValue = Left$(sValue, cch - 1)
        Else
            vValue = Empty
        End If
        ' For DWORDS
        Case REG_DWORD:
        lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
        If lrc = ERROR_NONE Then vValue = lValue
        Case Else
        'all other data types not supported
        lrc = -1
    End Select
QueryValueExExit:
QueryValueEx = lrc
Exit Function
QueryValueExError:
Resume QueryValueExExit
End Function


Public Function GetSetting(AppName As String, Section As String, Key As String, Optional default As String, Optional hKeyName As hKeyNames = HKEY_LOCAL_MACHINE, Optional AppNameHeader = "SOFTWARE") As String
    Dim lRetVal As Long 'result of the API functions
    Dim hKey As Long 'handle of opened key
    Dim vValue As Variant 'setting of queried value
    Dim keyString As String

    keyString = ""


    If AppNameHeader <> "" Then
        keyString = keyString + AppNameHeader
    End If


    If AppName <> "" Then


        If keyString <> "" Then
            keyString = keyString & "\"
        End If
        keyString = keyString & AppName
    End If


    If Section <> "" Then


        If keyString <> "" Then
            keyString = keyString & "\"
        End If
        keyString = keyString & Section
    End If
    lRetVal = RegOpenKeyEx(hKeyName, keyString, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = QueryValueEx(hKey, Key, vValue)


    If IsEmpty(vValue) Then
        vValue = default
    End If
    GetSetting = vValue
    RegCloseKey (hKey)
    Exit Function
e_Trap:
    vValue = default
    Exit Function
End Function

Public Function SaveSetting(AppName As String, Section As String, Key As String, Setting As String, Optional hKeyName As hKeyNames = HKEY_LOCAL_MACHINE, Optional AppNameHeader = "SOFTWARE") As Boolean
    Dim lRetVal As Long 'result of the SetValueEx function
    Dim hKey As Long 'handle of open key
    Dim keyString As String
    On Error GoTo e_Trap
    keyString = ""


    If AppNameHeader <> "" Then
        keyString = keyString + AppNameHeader
    End If


    If AppName <> "" Then


        If keyString <> "" Then
            keyString = keyString & "\"
        End If
        keyString = keyString & AppName
    End If


    If Section <> "" Then


        If keyString <> "" Then
            keyString = keyString & "\"
        End If
        keyString = keyString & Section
    End If
    lRetVal = RegCreateKeyEx(hKeyName, keyString, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
    lRetVal = SetValueEx(hKey, Key, REG_SZ, Setting)
    RegCloseKey (hKey)
    SaveSetting = True
    Exit Function
e_Trap:
    SaveSetting = False
    Exit Function
End Function

Public Function DeleteSetting(AppName As String, Optional Section As String, Optional Key As String, Optional hKeyName As hKeyNames = HKEY_LOCAL_MACHINE, Optional AppNameHeader = "SOFTWARE") As Boolean
    Dim hNewKey As Long 'handle to the new key
    Dim lRetVal As Long 'result of the SetValueEx function
    Dim hKey As Long 'handle of open key
    Dim keyString As String
    On Error GoTo e_Trap
    keyString = ""


    If AppNameHeader <> "" Then
        keyString = keyString + AppNameHeader
    End If


    If AppName <> "" Then


        If keyString <> "" Then
            keyString = keyString & "\"
        End If
        keyString = keyString & AppName
    End If


    If Section <> "" Then


        If keyString <> "" Then
            keyString = keyString & "\"
        End If
        keyString = keyString & Section
    End If


    If Key <> "" Then
        lRetVal = RegCreateKeyEx(hKeyName, keyString, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
        lRetVal = RegDeleteValue(hKey, Key)
        RegCloseKey (hKey)
    Else
        lRetVal = RegDeleteKey(hKeyName, keyString)
    End If
    DeleteSetting = True
    Exit Function
e_Trap:
    DeleteSetting = False
    Exit Function
End Function

Public Property Get Environ(variableName As String) As String
    Environ = GetSetting("Session Manager", "Environment", variableName, "", HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control")
End Property


Public Property Let Environ(variableName As String, Setting As String)
    Call SaveSetting("Session Manager", "Environment", variableName, Setting, HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control")
    Call SetEnvironmentVariable(variableName, Setting)
End Property


Public Sub VerifyPath(pathString As String)
    Dim CurrentPath As String
    pathString = Trim(pathString)
    If pathString = "" Then Exit Sub
    CurrentPath = Environ("PATH")


    If Mid(pathString, 1, 1) = ";" Then
        pathString = Mid(pathString, 2)
    End If


    If Mid(pathString, Len(pathString), 1) = ";" Then
        pathString = Mid(pathString, 1, Len(pathString) - 1)
    End If


    If InStr(1, UCase(CurrentPath), UCase(pathString), vbTextCompare) = 0 Then


        If Mid(CurrentPath, Len(CurrentPath), 1) = ";" Then
            Environ("PATH") = CurrentPath & pathString
        Else
            Environ("PATH") = CurrentPath & ";" & pathString
        End If
    End If
End Sub



Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
  Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot& + Len(ReplaceWith$)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, LCase(MyString$), LCase(ToFind$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = NewString$
End Function


Sub FormDrag(TheForm As Form)
    Call ReleaseCapture
    Call SendMessage(TheForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0&)
End Sub
