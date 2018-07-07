Attribute VB_Name = "modMain"
'tombol ctrl+alt+del
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

'declare registry
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const KEY_ALL_ACCESS = &HF003F
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEY = &H8
Public Const KEY_EXECUTE = &H20019
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_READ = &H20019
Public Const KEY_SET_VALUE = &H2
Public Const KEY_WRITE = &H20006
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const REG_DWORD_LITTLE_ENDIAN = 4
Public Const REG_DWORD_BIG_ENDIAN = 5
Public Const REG_EXPAND_SZ = 2
Public Const REG_LINK = 6
Public Const REG_MULTI_SZ = 7
Public Const REG_NONE = 0
Public Const REG_RESOURCE_LIST = 8
Public Const REG_SZ = 1

Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubkey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubkey As String, phkResult As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubkey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubkey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

' Declare functions obtained from the API Text Viewer
Private Declare Function SetWindowPos Lib _
    "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
' Constants for SetWindowPos hWndInsertAfter Parameter
Private Const HWND_TOP = 0
Private Const HWND_BOTTOM = 1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
' Constants for SetWindowPos wFlags Parameter
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40
Public frmpesanload As Boolean
Public Sub SetFormMain()
Dim x As Long
    x = SetWindowPos(FrmMain.hwnd, HWND_TOPMOST, 0, 0, Screen.Width, Screen.Height, SWP_SHOWWINDOW)
    If FrmMain.FrameProfile.Visible = False Then
        FrmMain.FrameProfile.Visible = True
    End If
End Sub
Public Sub Down()
        x = SetWindowPos(FrmMain.hwnd, HWND_NOTOPMOST, 0, 0, Screen.Width, Screen.Height, SWP_SHOWWINDOW)
        Shell ("Shutdown -f -s -t 0")
        End
End Sub
Public Sub RegSaveString(hKey As Long, strPath As String, strValue As String, strData As String)
Dim keyHand As Long
Dim x As Long
    x = RegCreateKey(hKey, strPath, keyHand)
    x = RegSetValueEx(keyHand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    x = RegCloseKey(keyHand)
End Sub
Public Function GetRegistryValue(ByVal hKey As Long, ByVal subkey_name As String) As String
Dim value As String
Dim Length As Long
Dim value_type As Long
    Length = 256
    value = Space$(Length)
    If RegQueryValueEx(hKey, subkey_name, 0&, value_type, ByVal value, Length) <> ERROR_SUCCESS Then
        value = "<Error>"
    Else
        value = Left$(value, Length - 1)
    End If
    GetRegistryValue = value
End Function

Public Sub MatikanCtrlAltDel(bDiasble As Boolean)
Dim x As Long
    x = SystemParametersInfo(97, dDiasble, CStr(1), 0)
End Sub
