Attribute VB_Name = "Billing"

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

Public dbpass As String
Public cnstr As String
Public username As String
Public kunci(1 To ClientMax) As Boolean
Public remotestr1(1 To ClientMax) As String
Public remotestr2(1 To ClientMax) As String

'definisi setingan umum
Public mata_uang As String
Public discount As Long
Public jam_awal As String
Public jam_akhir As String
Public harga_personal As Single
Public harga_member As Single
Public harga_game As Single
Public harga_ketik As Single
Public harga_awal As Single
Public deposit_min As Single
Public password_admin As String
Public clientpopup As Long
Public pathdata As String
Public lang As String

'definisi form pesan loaded
Public FrmPesanLoad As Boolean

'definisi aplikasi
Public ketik1 As String
Public ketik2 As String
Public ketik3 As String
Public game1 As String
Public game2 As String
Public game3 As String

Public Function GetRegistryValue(ByVal hKey As Long, ByVal subkey_name As String) As String
Dim value As String
Dim length As Long
Dim value_type As Long
    length = 256
    value = Space$(length)
    If RegQueryValueEx(hKey, subkey_name, 0&, value_type, ByVal value, length) <> ERROR_SUCCESS Then
        value = "<Error>"
    Else
        value = Left$(value, length - 1)
    End If
    GetRegistryValue = value
End Function

Public Function FExists(OrigFile As String)
Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    FExists = fs.fileexists(OrigFile)
End Function

Public Function DirExists(OrigFile As String)
Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    DirExists = fs.folderexists(OrigFile)
End Function

Public Function EncryptText(strText As String, ByVal strPwd As String)
Dim i As Integer, c As Integer
Dim strBuff As String
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    EncryptText = strBuff
End Function

Public Function DecryptText(strText As String, ByVal strPwd As String)
Dim i As Integer, c As Integer
Dim strBuff As String
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    DecryptText = strBuff
End Function

Public Sub RegSaveString(hKey As Long, strPath As String, strValue As String, strData As String)
Dim keyHand As Long
Dim X As Long
    X = RegCreateKey(hKey, strPath, keyHand)
    X = RegSetValueEx(keyHand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    X = RegCloseKey(keyHand)
End Sub

Sub TulisKey()
Dim subkey As String
    Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\server\" & VersiAplikasi & "\", "path", App.path & "\")
    Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\server\" & VersiAplikasi & "\", "versi", VersiAplikasi)
    Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\server\" & VersiAplikasi & "\", "penulis", Penulis)
    Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\server\" & VersiAplikasi & "\", "email", EmailPenulis)
    Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\server\" & VersiAplikasi & "\", "lang", DefaultLang)
    Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\server\" & VersiAplikasi & "\", "homepage", HomePage)
    pathdata = App.path & "\"
    lang = DefaultLang
    dbpass = DefaultDbPass
    Call RegSaveString(HKEY_LOCAL_MACHINE, "Software\snakebite\server\" & VersiAplikasi & "\", "dbpass", EncryptText(dbpass, "password"))
End Sub

Sub CekKey()
Dim X As Long
    X = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\snakebite\server\" & VersiAplikasi, 0, KEY_READ, hregKey)
    If X <> 0 Then
        Call TulisKey
        If FExists(App.path & "\" & DataFileName) = True Then
            FileCopy App.path & "\" & DataFileName, App.path & "\" & OldFileName
            Kill App.path & "\" & DataFileName
        End If
    Else
        Call AmbilKey
    End If
    X = RegCloseKey(hregKey)
End Sub

Sub AmbilKey()
Dim X As Long
    X = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\snakebite\server\" & VersiAplikasi, 0, KEY_READ, hregKey)
    pathdata = GetRegistryValue(hregKey, "path")
    lang = GetRegistryValue(hregKey, "lang")
    dbpass = DecryptText(GetRegistryValue(hregKey, "dbpass"), "password")
End Sub

Public Sub Main()
    cnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pathdata & DataFileName & ";Persist Security Info=False;Jet OLEDB:Database Password=" & dbpass
End Sub

Public Sub PBcolor(PB As ProgressBar, Backcolor As Long, Forecolor As Long)
    SendMessage PB.hwnd, CCM_SETBKCOLOR, 0, ByVal Backcolor
    SendMessage PB.hwnd, PBM_SETBARCOLOR, 0, ByVal Forecolor
End Sub
