Attribute VB_Name = "modFirewall"
Public Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public sysPID(1 To 5) As Integer
Public expPID As Integer
Public servPID As Integer
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONUP = &H205
Public TrayI As NOTIFYICONDATA
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Function Registry_Read(Key_Path, Key_Name) As Variant
    
    On Error Resume Next
    
    Dim Registry As Object
    
    Set Registry = CreateObject("WScript.Shell")
    'Read Registry key to check for Operating System
    Registry_Read = Registry.regread(Key_Path & Key_Name)
    
End Function

Public Function isWinXp() As Boolean
    
    Dim Operating_System As String
    'Read this keep if Windows 9x
    Operating_System = Registry_Read("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\", "PRODUCTNAME")

    If Operating_System = "" Then
         'Read this key if XP
         Operating_System = Registry_Read("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\", "PRODUCTNAME")

    End If
    
    If UCase(Operating_System) = UCase("microsoft windows xp") Then
        isWinXp = True
    Else
        isWinXp = False
        'Not XP, Cant run program
    End If

End Function
