Attribute VB_Name = "modProc"
Public sysPID(1 To 5) As Integer
Public cSysTime As SYSTEMTIME
Public eSysTime As SYSTEMTIME
Public kSysTime As SYSTEMTIME
Public uSysTime As SYSTEMTIME
Public Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Public Declare Function GetProcessMemoryInfo Lib "PSAPI.DLL" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Public Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Const PROCESS_TERMINATE As Long = (&H1)
Public Const MAX_PATH As Integer = 260
Public Const TH32CS_SNAPHEAPLIST = &H1
Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPTHREAD = &H4
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public infoProcInfo As PROCESSENTRY32

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Public Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type
Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Sub enumProc()
Dim procType As String
    procType = ""
    servProc = 0
    uknProc = 0
    sysProc = 0
    tempName = ""
    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    r = Process32First(hSnapShot, uProcess)
    r = Process32Next(hSnapShot, uProcess)
    Do While r
        ProcessName = Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0))
        If UCase(ProcessName) = UCase("services.exe") Then
                servPID = uProcess.th32ProcessID
            ElseIf UCase(ProcessName) = UCase("explorer.exe") Then
                expPID = uProcess.th32ProcessID
            ElseIf UCase(ProcessName) = UCase("system") Then
                sysPID(1) = uProcess.th32ProcessID
            ElseIf UCase(ProcessName) = UCase("smss.exe") Then
                sysPID(2) = uProcess.th32ProcessID
            ElseIf UCase(ProcessName) = UCase("winlogon.exe") Then
                sysPID(3) = uProcess.th32ProcessID
            ElseIf UCase(ProcessName) = UCase("csrss.exe") Then
                sysPID(4) = uProcess.th32ProcessID
            ElseIf UCase(ProcessName) = UCase("lsass.exe") Then
                sysPID(5) = uProcess.th32ProcessID
        End If
        If tempPID = uProcess.th32ProcessID Then
            tempName = ProcessName
            foundName = 1
        End If
        Call popLstvw(ProcessName, uProcess, getPriority(uProcess.th32ProcessID), getProcMem(uProcess.th32ProcessID), getCTime(uProcess.th32ProcessID), procType)
        
        r = Process32Next(hSnapShot, uProcess)
    Loop
    CloseHandle hSnapShot
End Sub

Public Function getPriority(pid As Long)
    hwnd = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
    pri = GetPriorityClass(hwnd)
    CloseHandle hwnd
    getPriority = pri
End Function

Public Function getProcMem(pid As Long) As PROCESS_MEMORY_COUNTERS
Dim procMem As PROCESS_MEMORY_COUNTERS
    
    hwnd = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
    procMem.cb = LenB(procMem)
    GetProcessMemoryInfo hwnd, procMem, procMem.cb
    CloseHandle hwnd
    getProcMem = procMem
End Function
Public Function getCTime(pid As Long) As String
Dim createTime As FILETIME, exitTime As FILETIME, kernelTime As FILETIME, userTime As FILETIME
Dim test As Integer
    hwnd = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
    GetProcessTimes hwnd, createTime, exitTime, kernelTime, userTime
    Call convertTimes(createTime, exitTime, kernelTime, userTime)
    CloseHandle hwnd
    
    getCTime = parseSysTime(cSysTime)
End Function
Public Sub convertTimes(cTime As FILETIME, eTime As FILETIME, kTime As FILETIME, uTime As FILETIME)
    FileTimeToSystemTime cTime, cSysTime
    FileTimeToSystemTime eTime, eSysTime
    FileTimeToSystemTime kTime, kSysTime
    FileTimeToSystemTime uTime, uSysTime
End Sub
Public Function parseSysTime(cTime As SYSTEMTIME) As String
Dim dayof As Integer, monthof As String, dayS As String
    
    Select Case cTime.wDayOfWeek
    
        Case 0
            dayS = "Sun"
        Case 1
            dayS = "Mon"
        Case 2
            dayS = "Tues"
        Case 3
            dayS = "Wed"
        Case 4
            dayS = "Thur"
        Case 5
            dayS = "Fri"
        Case 6
            dayS = "Sat"
            
    End Select
    
    Select Case cTime.wMonth
    
        Case 1
            monthof = "Jan"
        Case 2
            monthof = "Feb"
        Case 3
            monthof = "Mar"
        Case 4
            monthof = "Apr"
        Case 5
            monthof = "May"
        Case 6
            monthof = "Jun"
        Case 7
            monthof = "Jul"
        Case 8
            monthof = "Aug"
        Case 9
            monthof = "Sept"
        Case 10
            monthof = "Oct"
        Case 11
            monthof = "Nov"
        Case 12
            monthof = "Dec"
            
    End Select
    dayof = cTime.wDay
    
    If cTime.wHour > 5 And cTime.wHour < 18 Then
        If cTime.wHour - 5 = 12 Then
            timeof = "12:" & cTime.wMinute & "pm"
        Else
            timeof = cTime.wHour - 5 & ":" & cTime.wMinute & "am"
        End If
    ElseIf cTime.wHour > 17 And cTime.wHour < 25 Then
        timeof = cTime.wHour - 17 & ":" & cTime.wMinute & "pm"
    ElseIf cTime.wHour > 0 And cTime.wHour < 6 Then
        If cTime.wHour + 7 = 12 Then
            timeof = "12:" & cTime.wMinute & "am"
        Else
            timeof = cTime.wHour + 7 & ":" & cTime.wMinute & "pm"
        End If
    End If
    parseSysTime = dayS & " " & monthof & " " & dayof & ", " & timeof
End Function


Public Sub popLstvw(procName, procArray As PROCESSENTRY32, priority As Long, procMem As PROCESS_MEMORY_COUNTERS, creationDate As String, procType As String)
Dim procArr(1 To 6) As Variant
Dim tmpPri As String

    If procType = "" Then
        procType = "Other"
    End If
    
    Select Case priority
    
    Case 32
        tmpPri = "Normal"
    Case 64
        tmpPri = "Idle"
    Case 128
        tmpPri = "High"
    Case 256
        tmpPri = "RealTime"
    End Select
    
    procArr(1) = procArray.th32ProcessID
    procArr(2) = procArray.cntThreads
    procArr(3) = procArray.th32ParentProcessID
    procArr(4) = procArray.pcPriClassBase
    procArr(5) = procArray.szExeFile
    procArr(6) = tmpPri
    Set lstItem = FrmMain.ListProses.ListItems.Add(, , procName)
        lstItem.SubItems(1) = procArr(1)
        lstItem.SubItems(2) = procArr(2)
        lstItem.SubItems(3) = procArr(3)
        lstItem.SubItems(4) = procArr(4)
        lstItem.SubItems(5) = procArr(5)
End Sub
