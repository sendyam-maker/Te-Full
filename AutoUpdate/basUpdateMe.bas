Attribute VB_Name = "basUpdateMe"
'Memo By Morgan 2012/12/10 智權人員欄已修改
Option Explicit

'觸發 timer 用
Public IsGo As Boolean
Public UpdateCommand As String
Public UpdatePath As String
Public UpdateExe As String
Public tempPath As String
Public UpdateDate As String

Global Const TH32CS_SNAPPROCESS = &H2&
Global Const PROCESS_QUERY_INFORMATION = 1024
Global Const PROCESS_VM_READ = 16
Global Const MAX_PATH = 260
Global Const hNull = 0
Global Const FO_MOVE = &H1
Global Const FO_DELETE = &H3
Global Const FOF_ALLOWUNDO = &H40
Global Const FOF_NOCONFIRMATION = &H10
Global Const FOF_SILENT = &H4
Global Const STARTF_USESHOWWINDOW = &H1
Global Const SW_SHOWNORMAL = &H1
Global Const SW_SHOWDEFAULT = &H10
Global Const OFS_MAXPATHNAME = 128
Global Const OF_READWRITE = &H2

 Public Type PROCESSENTRY32
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long           ' This process
   th32DefaultHeapID As Long
   th32ModuleID As Long            ' Associated exe
   cntThreads As Long
   th32ParentProcessID As Long     ' This process's parent process
   pcPriClassBase As Long          ' Base priority of process threads
   dwFlags As Long
   szExeFile As String * 260       ' MAX_PATH
End Type

Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long           '1 = Windows 95.
                                  '2 = Windows NT

   szCSDVersion As String * 128
End Type

Public Type SHFILEOPSTRUCT
        hWnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Long
        hNameMappings As Long
        lpszProgressTitle As String
End Type
Public Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
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

Public Type TIME_ZONE_INFORMATION
      bias As Long
      StandardName(32) As Integer
      StandardDate As SYSTEMTIME
      StandardBias As Long
      DaylightName(32) As Integer
      DaylightDate As SYSTEMTIME
      DaylightBias As Long
 End Type
 
 Public Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDriectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Public Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As Any) As Long

Public Function CheckRunIs(ProcName As String) As Boolean
CheckRunIs = True
Dim hProcess As Long
Dim cb As Long
Dim cbNeeded As Long
Dim NumElements As Long
Dim ProcessIDs() As Long
Dim cbNeeded2 As Long
Dim NumElements2 As Long
Dim Modules(1 To 200) As Long
Dim lRet As Long
Dim ModuleName As String
Dim nSize As Long
Dim f As Long, sname As String
Dim hSnap As Long, proc As PROCESSENTRY32
Dim i As Long


Select Case getVersion()
Case 1 'Windows 95/98
        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap = hNull Then Exit Function
        proc.dwSize = Len(proc)
        '查詢正在執行的執行檔
        f = Process32First(hSnap, proc)
        Do While f
          sname = StrZToStr(proc.szExeFile)
          '檢查檔案
          'edit by nickc 2006/01/04
          'If InStr(1, UCase(sname), UCase(ProcName)) <> 0 Then
          If InStr(1, UCase(sname), UCase(ProcName)) <> 0 And InStr(1, UCase(sname), UCase("台一國際專利商標-----自動更新")) <> 0 Then
                Exit Function
                '查詢 process code
'                 hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
'                                 Or PROCESS_VM_READ, 0, proc.th32ProcessID)
'                 If hProcess <> 0 Then
'                    '提示關閉
'                    If MsgBox(ProcName & " 正在使用中，請先存檔再按下確定！", vbOKOnly, "警告！") = vbOK Then
'                        '強制關閉
'                         'TerminateProcess hProcess, 0
'                    End If
'                 End If
'                 CloseHandle (hProcess)
          End If
          f = Process32Next(hSnap, proc)
        Loop
Case 2
   cb = 8
   cbNeeded = 96
   Do While cb <= cbNeeded
      cb = cb * 2
      ReDim ProcessIDs(cb / 4) As Long
      lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
   Loop
   NumElements = cbNeeded / 4

   For i = 1 To NumElements
      '查詢 process code
      hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
         Or PROCESS_VM_READ, 0, ProcessIDs(i))
      If hProcess <> 0 Then
          lRet = EnumProcessModules(hProcess, Modules(1), 200, _
                                       cbNeeded2)
          If lRet <> 0 Then
             ModuleName = Space(MAX_PATH)
             nSize = 500
             lRet = GetModuleFileNameExA(hProcess, Modules(1), _
                             ModuleName, nSize)
            If InStr(1, UCase(ModuleName), UCase(ProcName)) <> 0 Then
                Exit Function
'                    If hProcess <> 0 Then
'                       '提示關閉
'                       If MsgBox(ProcName & " 正在使用中，請先存檔再按下確定！", vbOKOnly, "警告！") = vbOK Then
'                           '強制關閉
'                            'TerminateProcess hProcess, 0
'                       End If
'                    End If
'                    CloseHandle (hProcess)
             End If
          End If
      End If
   lRet = CloseHandle(hProcess)
   Next
End Select
CheckRunIs = False
End Function


Public Function getVersion() As Long
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer
   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)
   getVersion = osinfo.dwPlatformId
End Function

Public Function StrZToStr(s As String) As String
   StrZToStr = Left$(s, Len(s) - 1)
End Function

Public Function RecCommand(oStr As String)
Dim tmpCommand As Variant
Dim tmpName As Variant
tmpCommand = Split(oStr, "|")
tmpName = Split(tmpCommand(0), "\")
UpdateExe = tmpName(UBound(tmpName))
UpdatePath = Trim(Replace(tmpCommand(0), UpdateExe, ""))
UpdateDate = Trim(tmpCommand(1))
End Function
