Attribute VB_Name = "basAutoUpdate"
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
Option Explicit

Public pub_HostName As String 'Added by Morgan 2017/9/27
Public pub_LoginUser As String 'Added by Morgan 2017/10/3
Public pub_LocalIP As String 'Added by Morgan 2020/8/27

'Add by Morgan 2008/7/29
Public bolReboot As Boolean '是否要重新開機
'觸發 timer 用
Public IsGo As Boolean
'檢查網路
Public IsNetErr As Boolean
'判斷分所還是北所
Public IsTaipei As Boolean
'檢查是否有公司程式
Public HaveProc As Boolean
'是否有收文程式，有則要啟動 傾聽器
'Public IsStartListen As Boolean
'分所傳資料用
'Public strSendDataClient() As Byte   '傳資料
'Public strGetDataClient() As Byte     '收資料
'Public strSendDataServer1() As Byte   '傳資料
'Public strGetDataServer1() As Byte     '收資料
'Public strSendDataServer2() As Byte   '傳資料
'Public strGetDataServer2() As Byte     '收資料
'Public strSendDataServer3() As Byte   '傳資料
'Public strGetDataServer3() As Byte     '收資料
'Public strSendDataServer4() As Byte   '傳資料
'Public strGetDataServer4() As Byte     '收資料
'Public strSendDataServer5() As Byte   '傳資料
'Public strGetDataServer5() As Byte     '收資料
'Public StrStateClient As String
'Public StrStateServer1 As String
'Public StrStateServer2 As String
'Public StrStateServer3 As String
'Public StrStateServer4 As String
'Public StrStateServer5 As String
'北所 ftp 假 IP
'Modify by Morgan 2010/7/13 改為 本機ip位址的前三組+.253
'Global Const Local_Ftp_ip  As String = "192.168.1.253"
Public Local_Ftp_ip  As String
Global Const Local_Ftp_port  As String = "21"

'北所 ftp 真 IP
'Modify by Morgan 2010/7/13 改為 本機ip位址的前三組+.253
'Global Const External_Ftp_ip  As String = "192.168.1.253" '"211.75.113.67"
Public External_Ftp_ip  As String
Global Const External_Ftp_port  As String = "21"

Public Taipei_Ftp_ip As String

'Modified by Morgan 2015/2/5 Linux 系統的 FTP Server 大小寫有分,目錄分隔符號要用 "/"
'Public Const cRemotePath As String = "\polycom\setup\newproc\"
Public Const cRemotePath As String = "/PolyCOM/Setup/NewProc/"
'end 2015/2/5
'Modified by Morgan 2017/3/17 O12大小寫有分,改與案件連線一樣都大寫以方便測試
'Public Const cAdoConnect As String = "Provider=MSDAORA.1;Password=pgmpwd;User ID=pgmid;Data Source=m51con;Persist Security Info=True"
'Modified by Morgan 2024/11/20 改與案件系統相同否則O19連線會錯
'Public Const cAdoConnect As String = "Provider=MSDAORA.1;Password=PGMPWD;User ID=PGMID;Data Source=m51con;Persist Security Info=True"
Public Const cAdoConnect As String = "Provider=OraOLEDB.Oracle;Password=PGMPWD;User ID=PGMID;Data Source=m51con;Persist Security Info=True"
Public adoConn As New ADODB.Connection 'Added by Morgan 2020/9/10

'系統預設路徑
Public WinPath As String, SysPath As String, tempPath As String
'add by nickc 2008/03/17
Public AppDPath As String, FontPath As String
'準備要更新的檔案總和
Public TaieAllFileSize As Long
'已下載的 size
Public TaieAllFileSize_OK As Long
'狀態跑馬用
Public StateRun As String

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

'程式名稱
'Public Type MyDate
'    oYear       As String * 4
'    oMonth    As String * 2
'    oDay        As String * 2
'    oHour      As String * 2
'    oMinute   As String * 2
'    oSecond  As String * 2
'End Type

Public Type TaieProc
    ProcCName   As String
    ProcName     As String
    LocalPath      As String
    LocalDate   As SYSTEMTIME
    RemotePath   As String
    RemoteDate   As SYSTEMTIME
    IsReg            As Boolean
    oFileSize         As Long  '舊檔案大小
    nFileSize         As Long  '新檔案大小
    IsDownOk     As Boolean
End Type
'總程式數
Public MaxTaieProc As Integer
Public AllTaieProc() As TaieProc
'add by nickc 2008/03/17
Public MaxTaieEUDC As Integer
Public AllTaieEUDC() As TaieProc
Public MaxUpdateTaieEUDC As Integer
'準備要更新的檔案總和
Public TaieEUDCFileSize As Long
'已下載的 size
Public TaieEUDCFileSize_OK As Long


'要更新的檔案
Public UpdateTaieProc() As TaieProc
'add by nickc 2008/03/17
Public UpdateTaieEUDC() As TaieProc

Public UpdateThisProc As TaieProc
Public MaxUpdateTaieProc As Integer
'所有可以執行的程式
Public AllRunProc() As TaieProc
Public MaxAllRunProc As Integer
'強制要下載的
Public NewTaieProc() As TaieProc
Public MaxNewTaieProc As Integer

'Add by Morgan 2011/3/23
Public bolNoEUDC As Boolean '是否更新造字檔

Global Const scUserAgent = "Taie Auto Update"
'API 用
Global Const ERROR_SUCCESS As Long = 0
Global Const WS_VERSION_REQD As Long = &H101
Global Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Global Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Global Const MIN_SOCKETS_REQD As Long = 1
Global Const SOCKET_ERROR As Long = -1
Global Const IP_SUCCESS As Long = 0
Global Const IP_STATUS_BASE As Long = 11000
Global Const IP_BUF_TOO_SMALL As Long = (11000 + 1)
Global Const IP_DEST_NET_UNREACHABLE As Long = (11000 + 2)
Global Const IP_DEST_HOST_UNREACHABLE As Long = (11000 + 3)
Global Const IP_DEST_PROT_UNREACHABLE As Long = (11000 + 4)
Global Const IP_DEST_PORT_UNREACHABLE As Long = (11000 + 5)
Global Const IP_NO_RESOURCES As Long = (11000 + 6)
Global Const IP_BAD_OPTION As Long = (11000 + 7)
Global Const IP_HW_ERROR As Long = (11000 + 8)
Global Const IP_PACKET_TOO_BIG As Long = (11000 + 9)
Global Const IP_REQ_TIMED_OUT As Long = (11000 + 10)
Global Const IP_BAD_REQ As Long = (11000 + 11)
Global Const IP_BAD_ROUTE As Long = (11000 + 12)
Global Const IP_TTL_EXPIRED_TRANSIT As Long = (11000 + 13)
Global Const IP_TTL_EXPIRED_REASSEM As Long = (11000 + 14)
Global Const IP_PARAM_PROBLEM As Long = (11000 + 15)
Global Const IP_SOURCE_QUENCH As Long = (11000 + 16)
Global Const IP_OPTION_TOO_BIG As Long = (11000 + 17)
Global Const IP_BAD_DESTINATION As Long = (11000 + 18)
Global Const IP_ADDR_DELETED As Long = (11000 + 19)
Global Const IP_SPEC_MTU_CHANGE As Long = (11000 + 20)
Global Const IP_MTU_CHANGE As Long = (11000 + 21)
Global Const IP_UNLOAD As Long = (11000 + 22)
Global Const IP_ADDR_ADDED As Long = (11000 + 23)
Global Const IP_GENERAL_FAILURE As Long = (11000 + 50)
Global Const MAX_IP_STATUS As Long = (11000 + 50)
Global Const IP_PENDING As Long = (11000 + 255)
Global Const PING_TIMEOUT As Long = 500
Global Const INADDR_NONE As Long = &HFFFFFFFF
Global Const MAX_WSADescription As Long = 256
Global Const MAX_WSASYSStatus As Long = 128
Global Const SPI_SCREENSAVERRUNNING = 97
Global Const INTERNET_OPEN_TYPE_DIRECT = 1
Global Const INTERNET_SERVICE_FTP = 1
Global Const INTERNET_FLAG_PASSIVE = &H8000000
Global Const ERROR_INTERNET_EXTENDED_ERROR = 12003
Global Const FTP_TRANSFER_TYPE_ASCII = &H1
Global Const FTP_TRANSFER_TYPE_BINARY = &H1
Global Const INTERNET_FLAG_TRANSFER_BINARY = &H2 'Added by Morgan 2015/2/5 Linux 系統的 FTP Server 要用此格式下載 Size 才會符合
Global Const MAX_PATH = 260
Global Const ERROR_NO_MORE_FILES = 18
Global Const OF_READ = &H0
Global Const OF_READWRITE = &H2
Global Const OFS_MAXPATHNAME = 128
Global Const FO_MOVE = &H1
Global Const FO_DELETE = &H3
Global Const FOF_ALLOWUNDO = &H40
Global Const FOF_NOCONFIRMATION = &H10
Global Const FOF_SILENT = &H4
Global Const STARTF_USESHOWWINDOW = &H1
Global Const SW_SHOWNORMAL = &H1
Global Const SW_SHOWDEFAULT = &H10
Global Const SW_HIDE = &H0
Global Const INTERNET_OPTION_CONNECT_TIMEOUT = &H2
'add by nickc 2008/03/17
Global Const CSIDL_PERSONAL = &H5 'My Documents
Global Const CSIDL_FONTS = &H14& '字型
Global Const CSIDL_APPDATA = &H1A  'app data

Public hOpen As Long
Public szFileRemote As String, szDirRemote As String, szFileLocal As String, hConnection As Long

Public Type SHFILEINFO
        hIcon As Long                      '  out: icon
        iIcon As Long          '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH '  out: display name (or path)
        szTypeName As String * 80         '  out: type name
End Type

Global Const SHGFI_DISPLAYNAME = &H200
Global Const SHGFI_EXETYPE = &H2000
Global Const SHGFI_LARGEICON = &H0
Global Const SHGFI_SHELLICONSIZE = &H4
Global Const SHGFI_SMALLICON = &H1
Global Const SHGFI_SYSICONINDEX = &H4000
Global Const SHGFI_TYPENAME = &H400
Global Const ILD_TRANSPARENT = &H1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOSIZE = &H1
Global Const PROCESS_QUERY_INFORMATION = 1024
Global Const PROCESS_VM_READ = 16
Global Const STANDARD_RIGHTS_REQUIRED = &HF0000
Global Const SYNCHRONIZE = &H100000
Global Const PROCESS_ALL_ACCESS = &H1F0FFF
'Global Const TH32CS_SNAPPROCESS = &H2&
Global Const TH32CS_SNAPPROCESS As Long = 2&
Global Const hNull = 0


Public shinfo As SHFILEINFO
Global Const SHGFI_ICON = &H100                         '  get icon
Global Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
   Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
   Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type BY_HANDLE_FILE_INFORMATION
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        dwVolumeSerialNumber As Long
        nFileSizeHigh As Long
        nFileSizeLow As Long
        nNumberOfLinks As Long
        nFileIndexHigh As Long
        nFileIndexLow As Long
End Type
Public Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Public Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Long
    Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Public Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Public Type HOSTENT
  hName As Long
  hAliases As Long
  hAddrType As Integer
  hLen As Integer
  hAddrList As Long
End Type



Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
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

Type STARTUPINFO
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

Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

 Type TIME_ZONE_INFORMATION
      bias As Long
      StandardName(32) As Integer
      StandardDate As SYSTEMTIME
      StandardBias As Long
      DaylightName(32) As Integer
      DaylightDate As SYSTEMTIME
      DaylightBias As Long
 End Type
 
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

'Public Type BY_HANDLE_FILE_INFORMATION
'        dwFileAttributes As Long
'        ftCreationTime As FILETIME
'        ftLastAccessTime As FILETIME
'        ftLastWriteTime As FILETIME
'        dwVolumeSerialNumber As Long
'        nFileSizeHigh As Long
'        nFileSizeLow As Long
'        nNumberOfLinks As Long
'        nFileIndexHigh As Long
'        nFileIndexLow As Long
'End Type

'API 宣告
Public Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Public Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nbytes As Long)
Public Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal s As String) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function InternetReadFileByte Lib "wininet.dll" Alias "InternetReadFile" (ByVal hFile As Long, ByVal AddOfBuffer As Long, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal dwOption As Long, ByRef lpBuffer As Any, ByVal dwBufferLength As Long) As Long
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean
Public Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal sFileName As String, ByVal lAccess As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
'Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal flags As Long) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDriectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As Any) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Public Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'Public Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Const INFINITE = &HFFFFFFFF
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public IsHaveEudc As Boolean
Public IsEudcUsing As Boolean
'add by nickc 2008/03/17
Public IsCntChg As Boolean
Public Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hwndOwner As Long, ByVal nFolder As Integer, ppidl As Long) As Long
Public Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal szPath As String) As Long
'add by nickc 2008/04/08 debug 用
Public DBMode As Boolean
Public debugListItem As Long


'Add by Morgan 2010/7/15
Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long

Public Function GetIpAddrTable()
   Dim Buf(0 To 511) As Byte
   Dim BufSize As Long: BufSize = UBound(Buf) + 1
   Dim rc As Long
   rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
   If rc <> 0 Then Err.Raise vbObjectError, , "GetIpAddrTable failed with return value " & rc
   Dim NrOfEntries As Integer: NrOfEntries = Buf(1) * 256 + Buf(0)
   If NrOfEntries = 0 Then GetIpAddrTable = Array(): Exit Function
   ReDim IpAddrs(0 To NrOfEntries - 1) As String
   Dim i As Integer
   For i = 0 To NrOfEntries - 1
      Dim j As Integer, s As String: s = ""
      For j = 0 To 3: s = s & IIf(j > 0, ".", "") & Buf(4 + i * 24 + j): Next
      IpAddrs(i) = s
   Next
   GetIpAddrTable = IpAddrs
End Function

Public Function GetLocalIP() As String
   Dim IpAddrs
   Dim IpAddr As String
   
   IpAddrs = GetIpAddrTable
   Dim i As Integer
   For i = LBound(IpAddrs) To UBound(IpAddrs)
      'Debug.Print IpAddrs(i)
      'Modified by Morgan 2020/8/27 改192.168.6.開頭(VPN)的優先
      'Modified by Morgan 2020/9/9 +第3碼小的優先
      If Left(IpAddrs(i), 10) = "192.168.6." Then
         IpAddr = IpAddrs(i)
         Exit For
      ElseIf Left(IpAddrs(i), 8) = "192.168." Then
         If IpAddr = "" Or Val(Mid(IpAddrs(i), 9)) < Val(Mid(IpAddr, 9)) Then
            IpAddr = IpAddrs(i)
         End If
      End If
   Next
   GetLocalIP = IpAddr
End Function
   
Public Function GetStatusCode(status As Long) As String

   Dim msg As String
   
   Select Case status
      Case IP_SUCCESS:               msg = "ip success"
      Case INADDR_NONE:              msg = "inet_addr: bad IP format"
      Case IP_BUF_TOO_SMALL:         msg = "ip buf too_small"
      Case IP_DEST_NET_UNREACHABLE:  msg = "ip dest net unreachable"
      Case IP_DEST_HOST_UNREACHABLE: msg = "ip dest host unreachable"
      Case IP_DEST_PROT_UNREACHABLE: msg = "ip dest prot unreachable"
      Case IP_DEST_PORT_UNREACHABLE: msg = "ip dest port unreachable"
      Case IP_NO_RESOURCES:          msg = "ip no resources"
      Case IP_BAD_OPTION:            msg = "ip bad option"
      Case IP_HW_ERROR:              msg = "ip hw_error"
      Case IP_PACKET_TOO_BIG:        msg = "ip packet too_big"
      Case IP_REQ_TIMED_OUT:         msg = "ip req timed out"
      Case IP_BAD_REQ:               msg = "ip bad req"
      Case IP_BAD_ROUTE:             msg = "ip bad route"
      Case IP_TTL_EXPIRED_TRANSIT:   msg = "ip ttl expired transit"
      Case IP_TTL_EXPIRED_REASSEM:   msg = "ip ttl expired reassem"
      Case IP_PARAM_PROBLEM:         msg = "ip param_problem"
      Case IP_SOURCE_QUENCH:         msg = "ip source quench"
      Case IP_OPTION_TOO_BIG:        msg = "ip option too_big"
      Case IP_BAD_DESTINATION:       msg = "ip bad destination"
      Case IP_ADDR_DELETED:          msg = "ip addr deleted"
      Case IP_SPEC_MTU_CHANGE:       msg = "ip spec mtu change"
      Case IP_MTU_CHANGE:            msg = "ip mtu_change"
      Case IP_UNLOAD:                msg = "ip unload"
      Case IP_ADDR_ADDED:            msg = "ip addr added"
      Case IP_GENERAL_FAILURE:       msg = "ip general failure"
      Case IP_PENDING:               msg = "ip pending"
      Case PING_TIMEOUT:             msg = "ping timeout"
      Case Else:                     msg = "unknown  msg returned"
   End Select
   
   GetStatusCode = CStr(status) '& "   [ " & msg & " ]"
   
End Function

Public Function Ping(sAddress As String, sDataToSend As String, ECHO As ICMP_ECHO_REPLY) As Long
   
   Dim hPort As Long
   Dim dwAddress As Long
   
   dwAddress = inet_addr(sAddress)
   
   If dwAddress <> INADDR_NONE Then
   
      hPort = IcmpCreateFile()
      
      If hPort Then
      
         Call IcmpSendEcho(hPort, _
                           dwAddress, _
                           sDataToSend, _
                           Len(sDataToSend), _
                           0, _
                           ECHO, _
                           Len(ECHO), _
                           PING_TIMEOUT)

         Ping = ECHO.status
         Call IcmpCloseHandle(hPort)
      
      End If
      
   Else
         Ping = INADDR_NONE
   End If
  
End Function

'將From移至畫面之中心
Public Sub MoveFormToCenter()
Dim intX As Integer, intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
intX = (Screen.Width - frmAutoUpdate.Width) / 2
intY = (Screen.Height - frmAutoUpdate.Height) / 4
'intY = 2000
frmAutoUpdate.Move intX, intY
'貼圖
   If Dir("c:\pics\background090.jpg") <> "" Then
         frmAutoUpdate.Image1 = LoadPicture("c:\pics\background090.jpg")
         sglWidth = frmAutoUpdate.Image1.Width
         sglHeight = frmAutoUpdate.Image1.Height
         For intX = 0 To Int(frmAutoUpdate.ScaleWidth / sglWidth)
             For intY = 0 To Int(frmAutoUpdate.ScaleHeight / sglHeight)
                 frmAutoUpdate.PaintPicture frmAutoUpdate.Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
             Next
         Next
   End If
End Sub

Sub ScanAllTaieProc()
Dim MyPath As String
Dim MyName As String
Dim SavePath As String
Dim ArrPath As Variant
Dim FileHandle As Long
Dim lpReOpenBuff As OFSTRUCT
Dim ft As SYSTEMTIME
Dim FileInfo As BY_HANDLE_FILE_INFORMATION
Dim oI As Integer
Dim len5 As Long
Dim tZone As TIME_ZONE_INFORMATION
Dim bias As Long
Dim writedate As Date
Dim ft2 As SYSTEMTIME
Dim stOsVer As String
Dim stImePath As String, stImeFile As String
Dim MyName2 As String 'Added by Morgan 2020/8/12

'本支程式
Put2DBList "　　分析本支程式"
UpdateThisProc.LocalPath = App.Path & "\"
UpdateThisProc.ProcName = App.EXEName & ".exe"
UpdateThisProc.ProcCName = "自動更新"
UpdateThisProc.RemotePath = cRemotePath 'dll\"
UpdateThisProc.oFileSize = FileLen(App.Path & "\" & App.EXEName & ".exe")
UpdateThisProc.IsReg = False
FileHandle = OpenFile(UpdateThisProc.LocalPath & UpdateThisProc.ProcName, lpReOpenBuff, OF_READ)
GetFileInformationByHandle FileHandle, FileInfo
CloseHandle FileHandle
Call GetTimeZoneInformation(tZone)
bias = tZone.bias
FileTimeToSystemTime FileInfo.ftLastWriteTime, ft
writedate = CDate(ft.wYear & "/" & ft.wMonth & "/" & ft.wDay & " " & ft.wHour & ":" & ft.wMinute & ":" & ft.wSecond) - TimeSerial(0, bias, 0)
ft2.wYear = Year(writedate)
ft2.wMonth = Month(writedate)
ft2.wDay = Day(writedate)
ft2.wHour = Hour(writedate)
ft2.wMinute = Minute(writedate)
ft2.wSecond = Second(writedate)
ft2.wDayOfWeek = Weekday(writedate)
ft2.wMilliseconds = ft.wMilliseconds
UpdateThisProc.LocalDate = ft2


'掃目錄，一律 te 開頭的目錄，且檔名為 te 開頭的執行檔，
'prj 開頭要註冊，且目錄會依 os 不同而有所不同，
'其餘開頭，不註冊，因為 interface 不能變(Accreport.exe)，
MaxTaieProc = 0
MaxAllRunProc = 0
SavePath = ""
Put2DBList "　　分析所有台一程式"

'Add by Morgan 2010/10/18 只要搜尋本目錄
If LCase(Mid(App.Path, InStrRev(App.Path, "\") + 1, 2)) = "te" Then
   MyPath = App.Path
   SavePath = "\"
Else
'end 2010/10/18
   '*****先掃所有路徑
   MyPath = "c:\program files\"   ' 指定路徑。
   MyName = Dir(MyPath, vbDirectory)   ' 找尋第一個子目錄。
   Do While MyName <> ""   ' 執行迴圈。
      If MyName <> "." And MyName <> ".." Then
         If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory And LCase(Mid(MyName, 1, 2)) = "te" Then
            SavePath = SavePath & MyName & "\:"
         End If
      End If
      MyName = Dir   ' 尋找下一個目錄。
   Loop
   '*****依所有路徑找檔案   執行檔
End If

ArrPath = Split(SavePath, ":")
For oI = 0 To UBound(ArrPath)
    'Added by Morgan 2020/8/12
    '因有時.exe會消失,增加檢查若有.f而沒有.exe時先更名
    MyName = Dir(MyPath & ArrPath(oI) & "te*.exe.f")
    Do While MyName <> ""
      UpdateProgramData , ">" & MyName 'Added by Morgan 2020/9/10
      MyName2 = Left(MyName, Len(MyName) - 2)
      If CheckFileExists(MyPath & ArrPath(oI) & MyName2) = False Then
         Name MyPath & ArrPath(oI) & MyName As MyPath & ArrPath(oI) & MyName2
         UpdateProgramData , ">>" & MyName2 'Added by Morgan 2020/9/10
      End If
      MyName = Dir
    Loop
    'end 2020/8/12
    
    MyName = Dir(MyPath & ArrPath(oI) & "te*.exe")
    Do While MyName <> ""   ' 執行迴圈。
       If MyName <> "." And MyName <> ".." Then
          If (GetAttr(MyPath & ArrPath(oI) & MyName) And vbDirectory) <> vbDirectory And LCase(Mid(MyName, 1, 2)) = "te" And UCase(MyName) <> UCase(App.EXEName & ".exe") Then
               UpdateProgramData , ">" & MyName 'Added by Morgan 2020/9/10
            
                'add by nick 2005/01/26 判斷分所的話要啟動監聽
'                If UCase(MyName) = "TEWRITER" And IsTaipei = False Then IsStartListen = True
                MaxTaieProc = MaxTaieProc + 1
                If UCase(MyName) <> "TEAUTODOWN.EXE" Then
                     MaxAllRunProc = MaxAllRunProc + 1
                End If
                ReDim Preserve AllTaieProc(MaxTaieProc) As TaieProc
                AllTaieProc(MaxTaieProc).ProcName = MyName
                AllTaieProc(MaxTaieProc).ProcCName = OneUcase(Replace(UCase(Mid(MyName, 3)), ".EXE", ""))
                AllTaieProc(MaxTaieProc).LocalPath = MyPath & ArrPath(oI)
                AllTaieProc(MaxTaieProc).RemotePath = cRemotePath ' & Mid(MyName, 3, InStr(1, MyName, ".") - 3) & "\support\"
                AllTaieProc(MaxTaieProc).IsReg = False
                AllTaieProc(MaxTaieProc).oFileSize = FileLen(MyPath & ArrPath(oI) & MyName)
                FileHandle = OpenFile(AllTaieProc(MaxTaieProc).LocalPath & AllTaieProc(MaxTaieProc).ProcName, lpReOpenBuff, OF_READ)
                GetFileInformationByHandle FileHandle, FileInfo
                CloseHandle FileHandle
                Call GetTimeZoneInformation(tZone)
                bias = tZone.bias
                FileTimeToSystemTime FileInfo.ftLastWriteTime, ft
                writedate = CDate(ft.wYear & "/" & ft.wMonth & "/" & ft.wDay & " " & ft.wHour & ":" & ft.wMinute & ":" & ft.wSecond) - TimeSerial(0, bias, 0)
                ft2.wYear = Year(writedate)
                ft2.wMonth = Month(writedate)
                ft2.wDay = Day(writedate)
                ft2.wHour = Hour(writedate)
                ft2.wMinute = Minute(writedate)
                ft2.wSecond = Second(writedate)
                ft2.wDayOfWeek = Weekday(writedate)
                ft2.wMilliseconds = ft.wMilliseconds
                AllTaieProc(MaxTaieProc).LocalDate = ft2
                If UCase(MyName) <> "TEAUTODOWN.EXE" Then
                     ReDim Preserve AllRunProc(MaxAllRunProc) As TaieProc
                     AllRunProc(MaxAllRunProc) = AllTaieProc(MaxTaieProc)
                End If
          End If
       End If
       MyName = Dir   ' 尋找下一個目錄。
    Loop
Next oI
'取得Windows 的目錄
WinPath = String(255, 0)
len5 = GetWindowsDirectory(WinPath, 256)
WinPath = Left(WinPath, InStr(1, WinPath, Chr(0)) - 1) & "\"
'取得Windows System的目錄
SysPath = String(255, 0)
len5 = GetSystemDirectory(SysPath, 256)
SysPath = Left(SysPath, InStr(1, SysPath, Chr(0)) - 1) & "\"
'取得Temp的Directory
tempPath = String(255, 0)
len5 = GetTempPath(256, tempPath)
tempPath = Left(tempPath, len5)
'add by nickc 2008/03/17
AppDPath = String(MAX_PATH, 0)
SHGetSpecialFolderLocation 0, CSIDL_APPDATA, len5
SHGetPathFromIDList len5, AppDPath
AppDPath = Left(AppDPath, InStr(AppDPath, Chr(0)) - 1) & "\"
'Modified by Morgan 2023/4/17 造字檔改放程式路徑
'FontPath = String(MAX_PATH, 0)
'SHGetSpecialFolderLocation 0, CSIDL_FONTS, len5
'SHGetPathFromIDList len5, FontPath
'FontPath = Left(FontPath, InStr(FontPath, Chr(0)) - 1) & "\"
FontPath = App.Path & "\"
'end 2023/4/17

'移除 prj*.exe,prj*.dll 檢查 Memo by Morgan 2017/9/19
'移除 acc*.exe 檢查 Memo by Morgan 2017/9/19

If bolNoEUDC = False Then

   Put2DBList "　　分析造字"
   MaxTaieEUDC = 0
   
   '移除XP以下版本及註冊機碼相關程式 Memo by Morgan 2017/9/19
   
   MyName = Dir(FontPath & "eudc.tte")
   If MyName <> "" Then
      MaxTaieEUDC = MaxTaieEUDC + 1
      ReDim Preserve AllTaieEUDC(MaxTaieEUDC) As TaieProc
      AllTaieEUDC(MaxTaieEUDC).ProcName = MyName
      AllTaieEUDC(MaxTaieEUDC).LocalPath = FontPath
      AllTaieEUDC(MaxTaieEUDC).RemotePath = cRemotePath 'dll\"
      AllTaieEUDC(MaxTaieEUDC).oFileSize = FileLen(FontPath & MyName)
      AllTaieEUDC(MaxTaieEUDC).IsReg = True
      FileHandle = OpenFile(AllTaieEUDC(MaxTaieEUDC).LocalPath & AllTaieEUDC(MaxTaieEUDC).ProcName, lpReOpenBuff, OF_READ)
      GetFileInformationByHandle FileHandle, FileInfo
      CloseHandle FileHandle
      Call GetTimeZoneInformation(tZone)
      bias = tZone.bias
      FileTimeToSystemTime FileInfo.ftLastWriteTime, ft
      writedate = CDate(ft.wYear & "/" & ft.wMonth & "/" & ft.wDay & " " & ft.wHour & ":" & ft.wMinute & ":" & ft.wSecond) - TimeSerial(0, bias, 0)
      ft2.wYear = Year(writedate)
      ft2.wMonth = Month(writedate)
      ft2.wDay = Day(writedate)
      ft2.wHour = Hour(writedate)
      ft2.wMinute = Minute(writedate)
      ft2.wSecond = Second(writedate)
      ft2.wDayOfWeek = Weekday(writedate)
      ft2.wMilliseconds = ft.wMilliseconds
      AllTaieEUDC(MaxTaieEUDC).LocalDate = ft2
   '造字檔不存在時也要下載
   Else
      MyName = "eudc.tte"
      MaxTaieEUDC = MaxTaieEUDC + 1
      ReDim Preserve AllTaieEUDC(MaxTaieEUDC) As TaieProc
      AllTaieEUDC(MaxTaieEUDC).ProcName = MyName
      AllTaieEUDC(MaxTaieEUDC).LocalPath = FontPath
      AllTaieEUDC(MaxTaieEUDC).RemotePath = cRemotePath
   End If
      
   MyName = Dir(FontPath & "eudc.euf")
   If MyName <> "" Then
      MaxTaieEUDC = MaxTaieEUDC + 1
      ReDim Preserve AllTaieEUDC(MaxTaieEUDC) As TaieProc
      AllTaieEUDC(MaxTaieEUDC).ProcName = MyName
      AllTaieEUDC(MaxTaieEUDC).LocalPath = FontPath
      AllTaieEUDC(MaxTaieEUDC).RemotePath = cRemotePath 'dll\"
      AllTaieEUDC(MaxTaieEUDC).oFileSize = FileLen(FontPath & MyName)
      AllTaieEUDC(MaxTaieEUDC).IsReg = True
      FileHandle = OpenFile(AllTaieEUDC(MaxTaieEUDC).LocalPath & AllTaieEUDC(MaxTaieEUDC).ProcName, lpReOpenBuff, OF_READ)
      GetFileInformationByHandle FileHandle, FileInfo
      CloseHandle FileHandle
      Call GetTimeZoneInformation(tZone)
       bias = tZone.bias
       FileTimeToSystemTime FileInfo.ftLastWriteTime, ft
       writedate = CDate(ft.wYear & "/" & ft.wMonth & "/" & ft.wDay & " " & ft.wHour & ":" & ft.wMinute & ":" & ft.wSecond) - TimeSerial(0, bias, 0)
       ft2.wYear = Year(writedate)
       ft2.wMonth = Month(writedate)
       ft2.wDay = Day(writedate)
       ft2.wHour = Hour(writedate)
       ft2.wMinute = Minute(writedate)
       ft2.wSecond = Second(writedate)
       ft2.wDayOfWeek = Weekday(writedate)
       ft2.wMilliseconds = ft.wMilliseconds
       AllTaieEUDC(MaxTaieEUDC).LocalDate = ft2
   Else
      MyName = "eudc.euf"
      MaxTaieEUDC = MaxTaieEUDC + 1
      ReDim Preserve AllTaieEUDC(MaxTaieEUDC) As TaieProc
      AllTaieEUDC(MaxTaieEUDC).ProcName = MyName
      AllTaieEUDC(MaxTaieEUDC).LocalPath = FontPath
      AllTaieEUDC(MaxTaieEUDC).RemotePath = cRemotePath
   End If
         
   'Add by Morgan 2011/2/9 Win7 造字檢查
   'Modify by Morgan 2011/3/16 +XP
   stOsVer = getVersionNo
   stImePath = AppDPath & "Microsoft\IME"
   
   MyName = Dir(stImePath, vbDirectory)
   If MyName = "" Then
      MkDir stImePath
   End If
   
   'XP
   If Val(stOsVer) < 6 Then
      '檢查輸入法目錄是否存在,若不存在則新增目錄
      stImePath = AppDPath & "Microsoft\IME\CHAJEI"
      stImeFile = "CHAJEI.TBL"
   'Win 7
   ElseIf Val(stOsVer) < 6.2 Then
      stImePath = AppDPath & "Microsoft\IME\IMTC10"
      stImeFile = "TCEUDCCJ.TBL"
   'Win 8 以上
   Else
      stImePath = AppDPath & "Microsoft\IME\15.0"
      MyName = Dir(stImePath, vbDirectory)
      If MyName = "" Then
         MkDir stImePath
      End If
      
      stImePath = AppDPath & "Microsoft\IME\15.0\IMETC"
      stImeFile = "TCEUDCCJ.TBL"
   End If
   
   MyName = Dir(stImePath, vbDirectory)
   If MyName = "" Then
      MkDir stImePath
   End If
   stImePath = stImePath & "\"
   
   MyName = Dir(stImePath & stImeFile)
   If MyName <> "" Then
      MaxTaieEUDC = MaxTaieEUDC + 1
      ReDim Preserve AllTaieEUDC(MaxTaieEUDC) As TaieProc
      AllTaieEUDC(MaxTaieEUDC).ProcName = MyName
      AllTaieEUDC(MaxTaieEUDC).LocalPath = stImePath
      AllTaieEUDC(MaxTaieEUDC).RemotePath = cRemotePath 'dll\"
      AllTaieEUDC(MaxTaieEUDC).oFileSize = FileLen(stImePath & MyName)
      AllTaieEUDC(MaxTaieEUDC).IsReg = True
      FileHandle = OpenFile(AllTaieEUDC(MaxTaieEUDC).LocalPath & AllTaieEUDC(MaxTaieEUDC).ProcName, lpReOpenBuff, OF_READ)
      GetFileInformationByHandle FileHandle, FileInfo
      CloseHandle FileHandle
      Call GetTimeZoneInformation(tZone)
      bias = tZone.bias
      FileTimeToSystemTime FileInfo.ftLastWriteTime, ft
      writedate = CDate(ft.wYear & "/" & ft.wMonth & "/" & ft.wDay & " " & ft.wHour & ":" & ft.wMinute & ":" & ft.wSecond) - TimeSerial(0, bias, 0)
      ft2.wYear = Year(writedate)
      ft2.wMonth = Month(writedate)
      ft2.wDay = Day(writedate)
      ft2.wHour = Hour(writedate)
      ft2.wMinute = Minute(writedate)
      ft2.wSecond = Second(writedate)
      ft2.wDayOfWeek = Weekday(writedate)
      ft2.wMilliseconds = ft.wMilliseconds
      AllTaieEUDC(MaxTaieEUDC).LocalDate = ft2
   Else
      MyName = stImeFile
      MaxTaieEUDC = MaxTaieEUDC + 1
      ReDim Preserve AllTaieEUDC(MaxTaieEUDC) As TaieProc
      AllTaieEUDC(MaxTaieEUDC).ProcName = MyName
      AllTaieEUDC(MaxTaieEUDC).LocalPath = stImePath
      AllTaieEUDC(MaxTaieEUDC).RemotePath = cRemotePath
   End If
   
   'XP
   If Val(stOsVer) < 6 Then
      '檢查輸入法目錄是否存在,若不存在則新增目錄
      stImePath = AppDPath & "Microsoft\IME\PHON"
      stImeFile = "PHON.TBL"
   'Win 7
   ElseIf Val(stOsVer) < 6.2 Then
      stImePath = AppDPath & "Microsoft\IME\IMTC10"
      stImeFile = "TCEUDCPH.TBL"
   'Win 8 以上
   Else
      stImePath = AppDPath & "Microsoft\IME\15.0"
      MyName = Dir(stImePath, vbDirectory)
      If MyName = "" Then
         MkDir stImePath
      End If
      stImePath = AppDPath & "Microsoft\IME\15.0\IMETC"
      stImeFile = "TCEUDCPH.TBL"
   End If
   
   MyName = Dir(stImePath, vbDirectory)
   If MyName = "" Then
      MkDir stImePath
   End If
   stImePath = stImePath & "\"
   
   MyName = Dir(stImePath & stImeFile)
   If MyName <> "" Then
      MaxTaieEUDC = MaxTaieEUDC + 1
      ReDim Preserve AllTaieEUDC(MaxTaieEUDC) As TaieProc
      AllTaieEUDC(MaxTaieEUDC).ProcName = MyName
      AllTaieEUDC(MaxTaieEUDC).LocalPath = stImePath
      AllTaieEUDC(MaxTaieEUDC).RemotePath = cRemotePath 'dll\"
      AllTaieEUDC(MaxTaieEUDC).oFileSize = FileLen(stImePath & MyName)
      AllTaieEUDC(MaxTaieEUDC).IsReg = True
      FileHandle = OpenFile(AllTaieEUDC(MaxTaieEUDC).LocalPath & AllTaieEUDC(MaxTaieEUDC).ProcName, lpReOpenBuff, OF_READ)
      GetFileInformationByHandle FileHandle, FileInfo
      CloseHandle FileHandle
      Call GetTimeZoneInformation(tZone)
      bias = tZone.bias
      FileTimeToSystemTime FileInfo.ftLastWriteTime, ft
      writedate = CDate(ft.wYear & "/" & ft.wMonth & "/" & ft.wDay & " " & ft.wHour & ":" & ft.wMinute & ":" & ft.wSecond) - TimeSerial(0, bias, 0)
      ft2.wYear = Year(writedate)
      ft2.wMonth = Month(writedate)
      ft2.wDay = Day(writedate)
      ft2.wHour = Hour(writedate)
      ft2.wMinute = Minute(writedate)
      ft2.wSecond = Second(writedate)
      ft2.wDayOfWeek = Weekday(writedate)
      ft2.wMilliseconds = ft.wMilliseconds
      AllTaieEUDC(MaxTaieEUDC).LocalDate = ft2
   Else
      MyName = stImeFile
      MaxTaieEUDC = MaxTaieEUDC + 1
      ReDim Preserve AllTaieEUDC(MaxTaieEUDC) As TaieProc
      AllTaieEUDC(MaxTaieEUDC).ProcName = MyName
      AllTaieEUDC(MaxTaieEUDC).LocalPath = stImePath
      AllTaieEUDC(MaxTaieEUDC).RemotePath = cRemotePath
   End If
End If

Put2DBList "　　分析 updateme"
'*****依路徑找檔案   UpdateMe檔
MyName = Dir(App.Path & "\updateme.exe")
Do While MyName <> ""   ' 執行迴圈。
   If MyName <> "." And MyName <> ".." Then
        If (GetAttr(App.Path & "\" & MyName) And vbDirectory) <> vbDirectory Then
            MaxTaieProc = MaxTaieProc + 1
            ReDim Preserve AllTaieProc(MaxTaieProc) As TaieProc
            AllTaieProc(MaxTaieProc).ProcName = MyName
            AllTaieProc(MaxTaieProc).LocalPath = App.Path & "\"
            AllTaieProc(MaxTaieProc).RemotePath = cRemotePath 'dll\"
            AllTaieProc(MaxTaieProc).oFileSize = FileLen(App.Path & "\" & MyName)
            AllTaieProc(MaxTaieProc).IsReg = False
            FileHandle = OpenFile(AllTaieProc(MaxTaieProc).LocalPath & AllTaieProc(MaxTaieProc).ProcName, lpReOpenBuff, OF_READ)
            GetFileInformationByHandle FileHandle, FileInfo
            CloseHandle FileHandle
            Call GetTimeZoneInformation(tZone)
             bias = tZone.bias
             FileTimeToSystemTime FileInfo.ftLastWriteTime, ft
             writedate = CDate(ft.wYear & "/" & ft.wMonth & "/" & ft.wDay & " " & ft.wHour & ":" & ft.wMinute & ":" & ft.wSecond) - TimeSerial(0, bias, 0)
             ft2.wYear = Year(writedate)
             ft2.wMonth = Month(writedate)
             ft2.wDay = Day(writedate)
             ft2.wHour = Hour(writedate)
             ft2.wMinute = Minute(writedate)
             ft2.wSecond = Second(writedate)
             ft2.wDayOfWeek = Weekday(writedate)
             ft2.wMilliseconds = ft.wMilliseconds
             AllTaieProc(MaxTaieProc).LocalDate = ft2
        End If
   End If
   MyName = Dir   ' 尋找下一個目錄。
Loop
End Sub

'讀取電腦名稱
Public Function PUB_ReadHostName() As String
   Dim stHostName As String
   Dim dwLength As Integer
   dwLength = 256
   stHostName = String(dwLength, Chr(0))
   gethostname stHostName, Len(stHostName)
   PUB_ReadHostName = Replace(stHostName, Chr(0), "")
End Function
'pChoice:0=紀錄執行檔名,1=紀錄O12安裝起始時間,2=紀錄O12安裝結束時間
Public Function UpdateProgramData(Optional pChoice As Integer = 0, Optional pIsRunExeName As String) As Boolean
   'Dim adoConn As New ADODB.Connection 'Removed by Morgan 2020/9/10 改全域
   Dim stSQL As String, iR As Integer
   Static iSNo As Integer
   Static sTime As String
   
On Error GoTo ErrHand
   
   PUB_OpenConn
   '紀錄執行時間及版本
   If pChoice = 0 Then
      If sTime = "" Then sTime = Format(Now, "yyyy/mm/dd hh:mm:ss")
      '+序號,以免dupe
      iSNo = iSNo + 1
      If iSNo > 99 Then iSNo = 1
   
      stSQL = "insert into PrintStartPoint(PSP01,PSP02,PSP03,PSP06)" & _
         " values('" & pub_HostName & "','AUTOUPDATE','" & sTime & " '||to_char(sysdate,'hh24:mi:ss')||'-" & Format(iSNo, "00") & "','" & App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision & " (" & pub_LoginUser & "@" & pub_LocalIP & ") " & pIsRunExeName & "')"
      adoConn.Execute stSQL, iR

      stSQL = "delete PrintStartPoint" & _
         " where PSP01='" & pub_HostName & "' and PSP02='AUTOUPDATE' and PSP03<to_char(sysdate-7,'yyyy/mm/dd hh24:mi:ss')"
      adoConn.Execute stSQL, iR
      
   '紀錄開始安裝時間
   ElseIf pChoice = 1 Then
      stSQL = "update PrintStartPoint set PSP06=to_char(sysdate,'yyyy/mm/dd hh24:mi:ss')" & _
         " where PSP01='" & pub_HostName & "' and PSP02='O12Client' and PSP03='Install'"
      adoConn.Execute stSQL, iR
      
      If iR = 0 Then
         stSQL = "insert into PrintStartPoint(PSP01,PSP02,PSP03,PSP06)" & _
            " values('" & pub_HostName & "','O12Client','Install',to_char(sysdate,'yyyy/mm/dd hh24:mi:ss'))"
         adoConn.Execute stSQL, iR
      End If
      
   '紀錄安裝完成時間
   ElseIf pChoice = 2 Then
      stSQL = "update PrintStartPoint set PSP06=PSP06||' - '||to_char(sysdate,'yyyy/mm/dd hh24:mi:ss')" & _
         " where PSP01='" & pub_HostName & "' and PSP02='O12Client' and PSP03='Install'"
      adoConn.Execute stSQL, iR
      
      If iR = 0 Then
         stSQL = "insert into PrintStartPoint(PSP01,PSP02,PSP03,PSP06)" & _
            " values('" & pub_HostName & "','O12Client','Install','- '||to_char(sysdate,'yyyy/mm/dd hh24:mi:ss')"
         adoConn.Execute stSQL, iR
      End If
      
   '檢查是否曾經安裝
   ElseIf pChoice = 3 Then
      stSQL = "update PrintStartPoint set PSP04=PSP04" & _
         " where PSP01='" & pub_HostName & "' and PSP02='O12Client' and PSP03='Install'"
      adoConn.Execute stSQL, iR
      If iR = 0 Then GoTo ExitPort
   
   '安裝成功
   ElseIf pChoice = 4 Then
      stSQL = "update PrintStartPoint set PSP04=1" & _
         " where PSP01='" & pub_HostName & "' and PSP02='O12Client' and PSP03='Install'"
      adoConn.Execute stSQL, iR
      
   'Added by Morgan 2018/5/31
   '檢查是否曾經測試 O12 連線
   ElseIf pChoice = 5 Then
      stSQL = "update PrintStartPoint set PSP05=0" & _
         " where PSP01='" & pub_HostName & "' and PSP02='O12Client' and PSP03='Test'"
      adoConn.Execute stSQL, iR
      If iR = 0 Then GoTo ExitPort
   
   '紀錄開始測試時間
   ElseIf pChoice = 6 Then
      stSQL = "update PrintStartPoint set PSP04=NULL,PSP05=1,PSP06=to_char(sysdate,'yyyy/mm/dd hh24:mi:ss')" & _
         " where PSP01='" & pub_HostName & "' and PSP02='O12Client' and PSP03='Test'"
      adoConn.Execute stSQL, iR
      
      If iR = 0 Then
         stSQL = "insert into PrintStartPoint(PSP01,PSP02,PSP03,PSP05,PSP06)" & _
            " values('" & pub_HostName & "','O12Client','Test',1,to_char(sysdate,'yyyy/mm/dd hh24:mi:ss'))"
         adoConn.Execute stSQL, iR
      End If
      
   '紀錄測試完成時間
   ElseIf pChoice = 7 Then
      stSQL = "update PrintStartPoint set PSP05=2,PSP06=PSP06||' - '||to_char(sysdate,'yyyy/mm/dd hh24:mi:ss')" & _
         " where PSP01='" & pub_HostName & "' and PSP02='O12Client' and PSP03='Test'"
      adoConn.Execute stSQL, iR
      
      If iR = 0 Then
         stSQL = "insert into PrintStartPoint(PSP01,PSP02,PSP03,PSP05,PSP06)" & _
            " values('" & pub_HostName & "','O12Client','Test',2,'- '||to_char(sysdate,'yyyy/mm/dd hh24:mi:ss')"
         adoConn.Execute stSQL, iR
      End If
      
   '測試成功
   ElseIf pChoice = 8 Then
      stSQL = "update PrintStartPoint set PSP04=1" & _
         " where PSP01='" & pub_HostName & "' and PSP02='O12Client' and PSP03='Test'"
      adoConn.Execute stSQL, iR
   'end 2018/5/31
   
   'Added by Morgan 2019/6/27
   'app:檢查是否有測試紀錄,沒有則新增
   ElseIf pChoice = 9 Then
      stSQL = "update PrintStartPoint set PSP04=PSP04" & _
         " where PSP01='" & pub_HostName & "' and PSP02='O12ClientUpdate' and PSP03='app'"
      adoConn.Execute stSQL, iR
      
      If iR = 0 Then
         stSQL = "insert into PrintStartPoint(PSP01,PSP02,PSP03,PSP05,PSP06)" & _
            " values('" & pub_HostName & "','O12ClientUpdate','app',1,to_char(sysdate,'yyyy/mm/dd hh24:mi:ss')||' app已存在')"
         adoConn.Execute stSQL, iR
         
         GoTo ExitPort
      End If
      
   'app:連線測試成功
   ElseIf pChoice = 10 Then
      stSQL = "update PrintStartPoint set PSP04=1" & _
         " where PSP01='" & pub_HostName & "' and PSP02='O12ClientUpdate' and PSP03='app'"
      adoConn.Execute stSQL, iR
      
   'instantclient:已安裝,檢查是否有測試紀錄,沒有則新增
   ElseIf pChoice = 11 Then
      stSQL = "update PrintStartPoint set PSP04=PSP04" & _
         " where PSP01='" & pub_HostName & "' and PSP02='O12ClientUpdate' and PSP03='instantclient'"
      adoConn.Execute stSQL, iR
      
      If iR = 0 Then
         stSQL = "insert into PrintStartPoint(PSP01,PSP02,PSP03,PSP05,PSP06)" & _
            " values('" & pub_HostName & "','O12ClientUpdate','instantclient',1,to_char(sysdate,'yyyy/mm/dd hh24:mi:ss')||' oledb已存在')"
         adoConn.Execute stSQL, iR
         
         GoTo ExitPort
      End If
      
   'instantclient:已安裝oledb, 連線測試成功
   ElseIf pChoice = 12 Then
      stSQL = "update PrintStartPoint set PSP04=1" & _
         " where PSP01='" & pub_HostName & "' and PSP02='O12ClientUpdate' and PSP03='instantclient'"
      adoConn.Execute stSQL, iR
      
   'instantclient:未安裝 oledb,檢查是否有安裝紀錄,沒有則新增 oledb 安裝開始時間
   ElseIf pChoice = 13 Then
      stSQL = "update PrintStartPoint set PSP04=PSP04" & _
         " where PSP01='" & pub_HostName & "' and PSP02='O12ClientUpdate' and PSP03='instantclient'"
      adoConn.Execute stSQL, iR
      
      If iR = 0 Then
         stSQL = "insert into PrintStartPoint(PSP01,PSP02,PSP03,PSP05,PSP06)" & _
            " values('" & pub_HostName & "','O12ClientUpdate','instantclient',1,to_char(sysdate,'yyyy/mm/dd hh24:mi:ss'))"
         adoConn.Execute stSQL, iR
         
         GoTo ExitPort
      End If
   'instantclient:紀錄 oledb 安裝完成時間
   ElseIf pChoice = 14 Then
      stSQL = "update PrintStartPoint set PSP05=2,PSP06=PSP06||' - '||to_char(sysdate,'yyyy/mm/dd hh24:mi:ss')" & _
         " where PSP01='" & pub_HostName & "' and PSP02='O12ClientUpdate' and PSP03='instantclient'"
      adoConn.Execute stSQL, iR
      
   End If
   UpdateProgramData = True
   
ErrHand:
   'If Err.Number <> 0 Then MsgBox Err.Description
ExitPort:
   'Set adoConn = Nothing 'Removed by Morgan 2020/9/10 改全域
   PUB_CloseConn 'Added by Morgan 2023/3/23
End Function

'Added by Morgan 2014/4/9
'取得FTP帳號密碼
Private Function GetFTPAccount(pID As String, pPWD As String) As Boolean
   'Dim adoConn As New ADODB.Connection 'Removed by Morgan 2020/9/10 改全域
   Dim adoRst As New ADODB.Recordset
   Dim stAccount As String, iPos As Integer
   
   pID = "74001"
   pPWD = "0418"
   
On Error GoTo ErrHand
   
   PUB_OpenConn 'Added by Morgan 2023/3/24
   
   adoRst.CursorLocation = adUseClient
   If Local_Ftp_ip = Taipei_Ftp_ip Then
      adoRst.Open "select oMan from SetSpecMan  where oCode='FTP_Account'", adoConn, adOpenStatic, adLockReadOnly
   Else
      adoRst.Open "select oMan from SetSpecMan  where oCode='FTP_AccountB'", adoConn, adOpenStatic, adLockReadOnly
   End If
   If adoRst.RecordCount > 0 Then
      stAccount = adoRst.Fields(0)
      iPos = InStr(stAccount, ":")
      If iPos > 0 Then
         pID = Left(stAccount, iPos - 1)
         pPWD = Mid(stAccount, iPos + 1)
      Else
         pPWD = stAccount
      End If
      GetFTPAccount = True
   End If
   adoRst.Close
   
ErrHand:
   'If Err.Number <> 0 Then MsgBox Err.Description
   
   'Set adoConn = Nothing 'Removed by Morgan 2020/9/10 改全域
   Set adoRst = Nothing
   PUB_CloseConn 'Added by Morgan 2023/3/23
End Function

Public Function FTP_Conn() As Boolean
Dim IsTimeOut As Boolean
Dim SeekTimer As Long
Dim stID As String, stPWD As String
   
GetFTPAccount stID, stPWD 'Added by Morgan 2014/4/9


FTP_Conn = True

'Dim SetTimeOutIsOk As Long
    'hOpen = InternetOpen(scUserAgent, INTERNET_FLAG_ASYNC, vbNullString, vbNullString, 0)
    '檢查連線
    
'Removed by Morgna 2024/8/12 呼叫前有先檢查,此處可略過(實際上若用網域名稱有時也會無法連線)
'    Put2DBList "　　檢查 FTP 是否存在"
'    frmAutoUpdate.lblState.Caption = "檢查網路...."
'    If frmAutoUpdate.Winsock1.State <> 0 Then frmAutoUpdate.Winsock1.Close
'    frmAutoUpdate.Winsock1.Connect Local_Ftp_ip, Local_Ftp_port
'    IsTimeOut = False
'    SeekTimer = Timer
'    'Modified by Morgan 2024/8/2
'    'Do While frmAutoUpdate.Winsock1.State = 6 And IsTimeOut = False
'    Do While frmAutoUpdate.Winsock1.State <> 7 And IsTimeOut = False
'    'end 2024/8/2
'        DoEvents
'        If Timer - SeekTimer > 2 Then
'            IsTimeOut = True
'        End If
'    Loop
'    If IsTimeOut = False Then
'        If frmAutoUpdate.Winsock1.State = 7 Then
'            frmAutoUpdate.Winsock1.Close
'        Else
'            frmAutoUpdate.Winsock1.Close
'            IsNetErr = True
'            FTP_Conn = False
'            Exit Function
'        End If
'    Else
'       frmAutoUpdate.Winsock1.Close
'       IsNetErr = True
'       FTP_Conn = False
'       Exit Function
'    End If
'    frmAutoUpdate.Winsock1.Close
    
    
         Put2DBList "　　開啟通道"
          hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
          If hOpen = 0 Then
            ErrorOut Err.LastDllError, "InternetOpen"
            'Added by Morgan 2024/8/12
            IsNetErr = True
            FTP_Conn = False
            Put2DBList "　　開啟通道失敗"
            'end 2024/8/12
          End If
          
      '        SetTimeOutIsOk = InternetSetOption(hOpen, INTERNET_OPTION_CONNECT_TIMEOUT, 8, 4)
              'edit by nickc 2005/04/11 不用 ping 指令，所以改做法
              'hConnection = InternetConnect(hOpen, IIf(IsTaipei, Local_Ftp_ip, External_Ftp_ip), IIf(IsTaipei, Local_Ftp_port, External_Ftp_port), _
              "74001", "74001", INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
              'Debug.Print Timer
              
              Put2DBList "　　開始連線到 FTP"
              
              'Modified by Morgan 2014/4/9
              'hConnection = InternetConnect(hOpen, Local_Ftp_ip, Local_Ftp_port, _
               "74001", "74001", INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
               hConnection = InternetConnect(hOpen, Local_Ftp_ip, Local_Ftp_port, _
                  stID, stPWD, INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
               'end 2014/4/9
               
              'Debug.Print Timer
              '檢查分所還是北所
              If hConnection = 0 Then
                  'ErrorOut Err.LastDllError, "InternetConnect"
      '            hConnection = InternetConnect(hOpen, External_Ftp_ip, External_Ftp_port, _
      '                                "74001", "74001", INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
      '            If hConnection = 0 Then
      'edit by nickc 2005/04/26 薛說不秀
      '                  SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
      '                  MsgBox "請檢查網路狀態！", , "警告！"
      '                  SetWindowPos frmAutoUpdate.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                        IsNetErr = True
                        FTP_Conn = False
                        Put2DBList "　　嘗試連到 FTP 失敗"
      '            Else
      '                  IsTaipei = False
      '            End If
               Else
                    IsTaipei = True
                    Put2DBList "　　FTP 連線成功"
               End If
End Function

Public Function FTP_Disc()
    If hConnection <> 0 Then InternetCloseHandle hConnection
    hConnection = 0
    Put2DBList "　　與 FTP 斷線 成功"
End Function

Public Function ErrorOut(dError As Long, szCallFunction As String)
    Dim dwIntError As Long, dwLength As Long
    Dim strBuffer As String
    If dError = ERROR_INTERNET_EXTENDED_ERROR Then
        InternetGetLastResponseInfo dwIntError, vbNullString, dwLength
        strBuffer = String(dwLength + 1, 0)
        InternetGetLastResponseInfo dwIntError, strBuffer, dwLength
        'MsgBox szCallFunction & " 延伸錯誤代碼: " & dwIntError & " " & strBuffer
        Put2DBList "　　網路錯誤：" & szCallFunction & " 延伸錯誤代碼: " & dwIntError & " " & strBuffer
        frmAutoUpdate.lblState.Caption = szCallFunction & " 延伸錯誤代碼: " & dwIntError & " " & strBuffer
    End If
'    If MsgBox(szCallFunction & " 錯誤代碼: " & dError & _
'        vbCrLf & "Close Connection and Session?", vbYesNo) = vbYes Then
'        If hConnection Then InternetCloseHandle hConnection
'        If hOpen Then InternetCloseHandle hOpen
'        hConnection = 0
'        hOpen = 0
'    End If
End Function

Function DownLoadFileToTemp(oObj As TaieProc) As Boolean
Dim hFile As Long
Dim strCommand As String
Dim bDoLoop As Boolean
Dim sBuffer As Long
Dim lNumberOfBytesRead As Long
Dim byteBuffer(102399) As Byte
Dim ReadyBuffer() As Byte
Dim F1 As Integer
Dim oIjk As Long
Put2DBList "　　切換目錄：" & oObj.RemotePath
rcd oObj.RemotePath
sBuffer = 102400
Put2DBList "　　開啟檔案通道：" & oObj.ProcName
'Modified by Morgan 2018/8/3 +強制重新下載參數(INTERNET_FLAG_RELOAD=&H80000000)
hFile = FtpOpenFile(hConnection, oObj.ProcName, &H80000000, INTERNET_FLAG_TRANSFER_BINARY + &H80000000, 0)
If hFile = 0 Then
    Put2DBList "　　　開啟失敗"
    DownLoadFileToTemp = False
    Exit Function
End If
F1 = FreeFile
Put2DBList "　　　開始下載"
'Open tempPath & oObj.ProcName For Binary As F1
Open oObj.LocalPath & oObj.ProcName & ".f" For Binary As F1
bDoLoop = True
While bDoLoop
    bDoLoop = InternetReadFileByte(hFile, VarPtr(byteBuffer(0)), sBuffer, lNumberOfBytesRead)
    If Not CBool(lNumberOfBytesRead) Then
        bDoLoop = False
    Else
        ReDim ReadyBuffer(lNumberOfBytesRead - 1) As Byte
        If lNumberOfBytesRead <> sBuffer Then
            For oIjk = 0 To (lNumberOfBytesRead - 1)
                ReadyBuffer(oIjk) = byteBuffer(oIjk)
                DoEvents
            Next oIjk
        Else
            ReadyBuffer = byteBuffer
        End If
        TaieAllFileSize_OK = TaieAllFileSize_OK + lNumberOfBytesRead
        If TaieAllFileSize_OK > frmAutoUpdate.ProgressBar1.Max Then
         TaieAllFileSize_OK = frmAutoUpdate.ProgressBar1.Max
        End If
        frmAutoUpdate.ProgressBar1.Value = TaieAllFileSize_OK '/ 1024
        frmAutoUpdate.lbldownState.Caption = Format(TaieAllFileSize_OK / 1024, "###,###,###,###") & "  Ｋ / " & Format(TaieAllFileSize / 1024, "###,###,###,###") & "  Ｋ"
        frmAutoUpdate.lblSpeed.Caption = Format(Trim(TaieAllFileSize_OK / TaieAllFileSize * 100), "###.00") & " ％"
        DoEvents
        Put #F1, , ReadyBuffer
    End If
Wend
Close F1
Put2DBList "　　　下載成功，關閉通道"
InternetCloseHandle hFile
Erase byteBuffer
Erase ReadyBuffer
End Function

'Function CheckUpdateFile(oObj As TaieProc) As Boolean
'    Dim pData As WIN32_FIND_DATA
'    Dim hFind As Long
'    Dim ft As SYSTEMTIME
'    Dim dtFileTime As Date
'    Dim nLastError As Long
'    'Dim bias As Long
'
'    If Len(oObj.RemotePath) > 0 Then rcd (oObj.RemotePath)
'    pData.cFileName = String(MAX_PATH, 0)
'        hFind = FtpFindFirstFile(hConnection, oObj.ProcName, pData, 0, 0)
'    nLastError = Err.LastDllError
'
'    'Call GetTimeZoneInformation(tZone)
'    'bias = tZone.bias
'    If hFind <> 0 Then
'        '比對檔案，若日期及 size 不同則要更新
'        FileTimeToSystemTime pData.ftLastWriteTime, ft
'        'If ft.wYear <> oObj.LocalDate.wYear Or ft.wMonth <> oObj.LocalDate.wMonth Or ft.wDay <> oObj.LocalDate.wDay Or ft.wHour <> oObj.LocalDate.wHour Or ft.wMinute <> oObj.LocalDate.wMinute Or pData.nFileSizeLow <> oObj.oFileSize Then
'        If IsNew(ft, oObj.LocalDate) And pData.nFileSizeLow <> oObj.oFileSize Then
'            MaxUpdateTaieProc = MaxUpdateTaieProc + 1
'            ReDim Preserve UpdateTaieProc(MaxUpdateTaieProc) As TaieProc
'            UpdateTaieProc(MaxUpdateTaieProc).IsReg = oObj.IsReg
'            UpdateTaieProc(MaxUpdateTaieProc).LocalPath = oObj.LocalPath
'            UpdateTaieProc(MaxUpdateTaieProc).ProcCName = oObj.ProcCName
'            UpdateTaieProc(MaxUpdateTaieProc).ProcName = oObj.ProcName
'            UpdateTaieProc(MaxUpdateTaieProc).RemotePath = oObj.RemotePath
'            UpdateTaieProc(MaxUpdateTaieProc).nFileSize = pData.nFileSizeLow
'            UpdateTaieProc(MaxUpdateTaieProc).IsDownOk = False
'            TaieAllFileSize = TaieAllFileSize + pData.nFileSizeLow
'            '不知為何不用算時間差
'            'dtFileTime = DateSerial(ft.wYear, ft.wMonth, ft.wDay) + TimeSerial(ft.wHour, ft.wMinute - bias, ft.wSecond)
'            'dtFileTime = DateSerial(ft.wYear, ft.wMonth, ft.wDay) + TimeSerial(ft.wHour, ft.wMinute, ft.wSecond)
'            UpdateTaieProc(MaxUpdateTaieProc).RemoteDate = ft
'        End If
'    End If
'    InternetCloseHandle hFind
'End Function

Public Sub rcd(pszDir As String)
    If pszDir = "" Then
        'MsgBox "請選擇伺服器端希望變更的目錄！"
        Put2DBList "　　切換目錄錯誤"
        frmAutoUpdate.lblState.Caption = "切換伺服器目錄錯誤..."
        Exit Sub
    Else
        Dim bRet As Boolean
        bRet = FtpSetCurrentDirectory(hConnection, pszDir)
        If bRet = False Then
            ErrorOut Err.LastDllError, "rcd"
        Else
            Put2DBList "　　切換成功"
        End If
    End If
End Sub

Public Function MoveFileToReady() As Boolean
Dim oIjk As Integer
Dim ShFileOP As SHFILEOPSTRUCT
Dim lpct As FILETIME, lplac As FILETIME, lplwr As FILETIME
Dim ofs As OFSTRUCT
Dim hFile As Long
Dim tZone As TIME_ZONE_INFORMATION
Dim bias As Long
Dim writedate As Date
Dim ft As SYSTEMTIME
Dim lngRCode As Long
Dim udtStartupInfo As STARTUPINFO
Dim udtProcessInfo As PROCESS_INFORMATION
Dim strCmd As String
Dim strExc As String
Dim updateFontCount As Integer
MoveFileToReady = True
For oIjk = 1 To MaxUpdateTaieProc
    '檢查是否下載成功
    If UpdateTaieProc(oIjk).IsDownOk = True Then
        '檢查要更新的檔案是否正在執行，若是 dll 或是 Active X EXE 時，要將其他的TE 開頭的程式斷線
        Put2DBList "　　檢查 " & UpdateTaieProc(oIjk).ProcName & "是否正在執行中"
        CheckAllRunIs UpdateTaieProc(oIjk)
        '檢查是否要註冊，Dll 才要
        If UpdateTaieProc(oIjk).IsReg = True And UCase(Right(UpdateTaieProc(oIjk).ProcName, 3)) = "DLL" Then
            '反註冊  改成要等待
'            Shell "command.com /c regsvr32.exe """ & UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & """ /u /s", vbHide
            strExc = "regsvr32.exe"
            strCmd = " """ & UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & """ /u /s"
            udtStartupInfo.cb = Len(udtStartupInfo)
            udtStartupInfo.dwFlags = STARTF_USESHOWWINDOW
            udtStartupInfo.wShowWindow = SW_HIDE
            udtProcessInfo.dwProcessId = 0&
            udtProcessInfo.dwThreadId = 0&
            udtProcessInfo.hProcess = 0&
            udtProcessInfo.hThread = 0&
            lngRCode = CreateProcess(vbNullString, _
                                     strExc & strCmd, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     udtStartupInfo, _
                                     udtProcessInfo)
            If lngRCode = 0& Then
                If DBMode = False Then
                    SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                    MsgBox "Failed to call CreateProcess() function."
                    SetWindowPos frmAutoUpdate.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                Else
                    Put2DBList "　　呼叫 CreateProcess 錯誤"
                End If
            Else
                Do
                    lngRCode = WaitForSingleObject(udtProcessInfo.hProcess, 300)
                    If lngRCode = 0& Then
                        Exit Do
                    End If
                    DoEvents
                Loop
                CloseHandle udtProcessInfo.hProcess
                udtProcessInfo.hProcess = 0&
                CloseHandle udtProcessInfo.hThread
                udtProcessInfo.hThread = 0&
            End If
        End If

        Put2DBList "　　更改日期時間與伺服器相同"
        '更改日期
        Call GetTimeZoneInformation(tZone)
        bias = tZone.bias
        writedate = CDate(UpdateTaieProc(oIjk).RemoteDate.wYear & "/" & UpdateTaieProc(oIjk).RemoteDate.wMonth & "/" & UpdateTaieProc(oIjk).RemoteDate.wDay & " " & UpdateTaieProc(oIjk).RemoteDate.wHour & ":" & UpdateTaieProc(oIjk).RemoteDate.wMinute & ":" & UpdateTaieProc(oIjk).RemoteDate.wSecond) + TimeSerial(0, bias, 0)
        ft.wYear = Year(writedate)
        ft.wMonth = Month(writedate)
        ft.wDay = Day(writedate)
        ft.wHour = Hour(writedate)
        ft.wMinute = Minute(writedate)
        ft.wSecond = Second(writedate)
        ft.wDayOfWeek = Weekday(writedate)
        ft.wMilliseconds = UpdateTaieProc(oIjk).RemoteDate.wMilliseconds
        'Modify by Morgan 2010/10/20 改.f時就要先更新(程式從下面搬上來)
        'hFile = OpenFile(UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName, ofs, OF_READWRITE)
        hFile = OpenFile(UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & ".f", ofs, OF_READWRITE)
        'UpdateTaieProc(oIjk).RemoteDate.wMinute = UpdateTaieProc(oIjk).RemoteDate.wMinute + bias
        Call SystemTimeToFileTime(ft, lplwr)
        '更動hFile的時間，第2個參數改Create DateTime
        '第3個參數改Last Access DateTime
        '第四個參數改Last Modify DateTime
        Call SetFileTime(hFile, ByVal 0, ByVal 0, lplwr) '只更改第三個(最後寫入)時間
        Call CloseHandle(hFile) '關閉檔案
        
        Put2DBList "　　開始更名成 f2"
        updateFontCount = 1
        Do While Dir(UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName) <> ""
            If updateFontCount >= 10 Then
                Put2DBList "　　更名 f2 失敗"
                MoveFileToReady = False
                'add by nickc 2008/03/12 失敗要還原
                ShFileOP.wFunc = FO_MOVE
                ShFileOP.pFrom = Left(UpdateTaieProc(oIjk).LocalPath, Len(UpdateTaieProc(oIjk).LocalPath) - 1) & "\" & UpdateTaieProc(oIjk).ProcName & ".f2" + Chr(0)
                ShFileOP.pTo = UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName
                ShFileOP.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_SILENT
                SHFileOperation ShFileOP
                Exit Function
            End If
            Put2DBList "　　" & updateFontCount
            updateFontCount = updateFontCount + 1
            ShFileOP.wFunc = FO_MOVE
            ShFileOP.pFrom = UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName + Chr(0)
            ShFileOP.pTo = Left(UpdateTaieProc(oIjk).LocalPath, Len(UpdateTaieProc(oIjk).LocalPath) - 1) & "\" & UpdateTaieProc(oIjk).ProcName & ".f2"
            ShFileOP.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_SILENT
            SHFileOperation ShFileOP
        Loop
        
        Put2DBList "　　開始將 f 更新成正確名稱"
        updateFontCount = 1
        'Modified by Morgan 2016/12/19 有發生exe不見情形,加exe檢查
        'Do While Dir(UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & ".f") <> ""
        Do While (Dir(UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & ".f") <> "" Or Dir(UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName) = "")
        'end 2016/12/19
            If updateFontCount >= 10 Then
                Put2DBList "　　將 f 更新成正確名稱失敗"
                MoveFileToReady = False
                'add by nickc 2008/03/12 失敗要還原
                ShFileOP.wFunc = FO_MOVE
                ShFileOP.pFrom = UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName + Chr(0)
                ShFileOP.pTo = UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & ".f"
                ShFileOP.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_SILENT
                SHFileOperation ShFileOP
                Exit Function
            End If
            Put2DBList "　　" & updateFontCount
            updateFontCount = updateFontCount + 1
            ShFileOP.wFunc = FO_MOVE
            ShFileOP.pFrom = UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & ".f" + Chr(0)
            ShFileOP.pTo = UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName
            ShFileOP.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_SILENT
            SHFileOperation ShFileOP
        Loop

        'Added by Morgan 2017/11/22 win7下更名有時也會更新修改時間
        hFile = OpenFile(UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName, ofs, OF_READWRITE)
        Call SystemTimeToFileTime(ft, lplwr)
        Call SetFileTime(hFile, ByVal 0, ByVal 0, lplwr) '只更改第三個(最後寫入)時間
        Call CloseHandle(hFile) '關閉檔案
        'end 2017/11/22
        
        '檢查是否要註冊，Dll 才要
        If UpdateTaieProc(oIjk).IsReg = True And UCase(Right(UpdateTaieProc(oIjk).ProcName, 3)) = "DLL" Then
            '註冊
'            Shell "command.com /c regsvr32.exe """ & UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & """ /s", vbHide
            strExc = "regsvr32.exe"
            strCmd = " """ & UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & """ /s"
            udtStartupInfo.cb = Len(udtStartupInfo)
            udtStartupInfo.dwFlags = STARTF_USESHOWWINDOW
            udtStartupInfo.wShowWindow = SW_HIDE
            udtProcessInfo.dwProcessId = 0&
            udtProcessInfo.dwThreadId = 0&
            udtProcessInfo.hProcess = 0&
            udtProcessInfo.hThread = 0&
            lngRCode = CreateProcess(vbNullString, _
                                     strExc & strCmd, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     udtStartupInfo, _
                                     udtProcessInfo)
            If lngRCode = 0& Then
                If DBMode = False Then
                    SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                    MsgBox "Failed to call CreateProcess() function."
                    SetWindowPos frmAutoUpdate.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                Else
                    Put2DBList "　　呼叫 CreateProcess 錯誤"
                End If
            Else
                Do
                    lngRCode = WaitForSingleObject(udtProcessInfo.hProcess, 300)
                    If lngRCode = 0& Then
                        Exit Do
                    End If
                    DoEvents
                Loop
                CloseHandle udtProcessInfo.hProcess
                udtProcessInfo.hProcess = 0&
                CloseHandle udtProcessInfo.hThread
                udtProcessInfo.hThread = 0&
            End If
        End If
        
        Put2DBList "　　刪除 f2"
        updateFontCount = 1
        Do While Dir(UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & ".f2") <> ""
            If updateFontCount >= 10 Then
                MoveFileToReady = False
                Put2DBList "　　f2 刪除失敗"
                Exit Function
            End If
            Put2DBList "　　" & updateFontCount
            updateFontCount = updateFontCount + 1
            ShFileOP.wFunc = FO_DELETE
            ShFileOP.pFrom = UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & ".f2" + Chr(0)
            ShFileOP.fFlags = FOF_NOCONFIRMATION + FOF_SILENT
            SHFileOperation ShFileOP
        Loop
        
    End If
Next oIjk
End Function

Public Function CheckSize()
Dim oIjk As Integer
Dim ShFileOP As SHFILEOPSTRUCT

For oIjk = 1 To MaxUpdateTaieProc
    '檔案 size 不合
    Put2DBList "　　檢查 " & UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & "是否存在"
    If Dir(UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & ".f") <> "" Then
        Put2DBList "　　檢查 " & UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & "檔案大小"
        If FileLen(UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & ".f") = UpdateTaieProc(oIjk).nFileSize Then
            UpdateTaieProc(oIjk).IsDownOk = True
            Put2DBList "　　檔案正確"
        Else
            UpdateTaieProc(oIjk).IsDownOk = False
            Put2DBList "　　錯誤，大小不符"
            Put2DBList "　　刪除暫存"
            '刪除
            ShFileOP.wFunc = FO_DELETE
            ShFileOP.pFrom = UpdateTaieProc(oIjk).LocalPath & UpdateTaieProc(oIjk).ProcName & ".f" & Chr(0)
            ShFileOP.fFlags = FOF_NOCONFIRMATION + FOF_SILENT
            SHFileOperation ShFileOP
            If DBMode = False Then
                SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
            End If
            Put2DBList "　　重新下載??"
            If MsgBox(UpdateTaieProc(oIjk).ProcCName & "  下載失敗，是否重新下載此檔！", vbYesNo, "錯誤！") = vbYes Then
                TaieAllFileSize = UpdateTaieProc(oIjk).nFileSize
                TaieAllFileSize_OK = 0
                frmAutoUpdate.ProgressBar1.Value = 0
                Put2DBList "　　重新下載"
                DownLoadFileToTemp UpdateTaieProc(oIjk)
                Put2DBList "　　檢查重新下載的大小"
                CheckSize
                Exit For
            End If
            If DBMode = False Then
                SetWindowPos frmAutoUpdate.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
            End If
        End If
    Else
        Put2DBList "　　不存在，重新下載??"
        UpdateTaieProc(oIjk).IsDownOk = False
        If DBMode = False Then
            SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
        End If
        Put2DBList "　　重新下載??"
        If MsgBox(UpdateTaieProc(oIjk).ProcCName & "  下載失敗，是否重新下載此檔！", vbYesNo, "錯誤！") = vbYes Then
            TaieAllFileSize = UpdateTaieProc(oIjk).nFileSize
            TaieAllFileSize_OK = 0
            frmAutoUpdate.ProgressBar1.Value = 0
            Put2DBList "　　重新下載"
            DownLoadFileToTemp UpdateTaieProc(oIjk)
            Put2DBList "　　檢查重新下載的大小"
            CheckSize
            Exit For
        End If
        If DBMode = False Then
            SetWindowPos frmAutoUpdate.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
        End If
    End If
Next oIjk
End Function

Public Function CheckAllUpdateFile() As Boolean
Dim oIjk As Integer
'add by nick 2005/01/07 若執行檔正在使用，不更新，連 dll 也是
Dim IsExeIsRun As Boolean
IsExeIsRun = False
Dim lngTaieAllFileSizeOld As Long 'Added by Morgan 2020/8/27
Dim stRunners As String 'Added by Morgan 2020/9/30

MaxUpdateTaieProc = 0
TaieAllFileSize = 0
For oIjk = 1 To MaxTaieProc
    'add by nick 2005/01/07 若執行檔正在使用，不更新，連 dll 也是
    'Modified by Morgan 2020/9/30
    'If CheckIsRun(AllTaieProc(oIjk).ProcName) = False Then
    If CheckIsRunning(AllTaieProc(oIjk).ProcName, stRunners) = False Then
    
        'add by nickc 2005/12/29 nt 2000 xp 的造字不更新，強制也略過
        'If InStr(1, UCase(AllTaieProc(oIjk).ProcName), "EUDC") = 0 Or getVersion = 1 Then
            If (IsExeIsRun = True And AllTaieProc(oIjk).IsReg = False) Or IsExeIsRun = False Then
                'add by nickc 2008/03/13
                'If InStr(1, UCase(AllTaieProc(oIjk).ProcName), "EUDC") > 0 Then IsHaveEudc = True
                Call CheckUpdateFile2(AllTaieProc(oIjk))
            End If
        'End If
    Else
        UpdateProgramData , "-" & AllTaieProc(oIjk).ProcName & "(" & Left(stRunners, 30) & ")" 'Added by Morgan 2020/8/28 紀錄執行中的程式
        IsExeIsRun = True
    End If
    
    'Added by Morgan 2020/8/27
    If TaieAllFileSize > lngTaieAllFileSizeOld Then
      UpdateProgramData , "+" & AllTaieProc(oIjk).ProcName
    End If
    lngTaieAllFileSizeOld = TaieAllFileSize
    'end 2020/8/27
Next oIjk

frmAutoUpdate.ProgressBar1.Min = 0
frmAutoUpdate.ProgressBar1.Max = Format(IIf(TaieAllFileSize = 0, 102400, TaieAllFileSize), "###########0")
frmAutoUpdate.ProgressBar1.Value = 0
frmAutoUpdate.lbldownState.Caption = "0 Ｋ / " & Format(TaieAllFileSize / 1024, "###,###,###,##0") & "  Ｋ"
End Function

'Added by Morgan 2020/9/11
'跟DB比對檔案是否最新
Public Function CheckAllUpdateFile2() As Boolean
   Dim oIjk As Integer
   Dim stSQL As String, intQ As Integer
   Dim stProcName As String, strProcDate As String
   
   CheckAllUpdateFile2 = True
   For oIjk = 1 To MaxTaieProc
      stProcName = AllTaieProc(oIjk).ProcName
      strProcDate = SysTimeToStr(AllTaieProc(oIjk).LocalDate)
      
      PUB_OpenConn 'Added by Morgan 2023/3/24
      
      '確認程式時間與最新紀錄相差一秒內
      stSQL = "update filelist set fl02=fl02 where upper(fl01)='" & UCase(stProcName) & "' and to_date(fl03,'yyyymmddhh24miss')-1/(24*60*60)<=to_date(" & strProcDate & ",'yyyymmddhh24miss')"
      adoConn.Execute stSQL, intQ
      If intQ = 0 Then
         CheckAllUpdateFile2 = False
         Exit For
      End If
   Next oIjk
   PUB_CloseConn 'Added by Morgan 2023/3/23
End Function

Public Function OneUcase(oStr As String) As String
OneUcase = UCase(Mid(oStr, 1, 1)) & LCase(Mid(oStr, 2))
End Function

Public Function DownLoadAllFileToTemp() As Boolean
Dim oIjk As Integer
TaieAllFileSize_OK = 0
For oIjk = 1 To MaxUpdateTaieProc
    'rcd ("\")
    Call DownLoadFileToTemp(UpdateTaieProc(oIjk))
Next oIjk
End Function

Public Sub Pic_to_Cmd(oObj As CommandButton, oObj2 As TaieProc)
   Dim r As Long
   Dim hImgLarge As Long
   Put2DBList "　　放圖"
   hImgLarge& = SHGetFileInfo(oObj2.LocalPath & oObj2.ProcName, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

   frmAutoUpdate.Picture1.Picture = LoadPicture()
   frmAutoUpdate.Picture1.AutoRedraw = True
 
   r = ImageList_Draw(hImgLarge&, shinfo.iIcon, frmAutoUpdate.Picture1.hDC, 16, 12, ILD_TRANSPARENT)

   oObj.Picture = LoadPicture()
   Set oObj.Picture = frmAutoUpdate.Picture1.Image
   oObj.Caption = oObj2.ProcCName
End Sub

Public Function StrToStr(ByRef Strindex As String, ByRef StrIndex2 As Single) As String
StrToStr = StrConv(MidB(StrConv(Strindex, vbFromUnicode), 1, StrIndex2 * 2), vbUnicode)
End Function

Public Function CheckAllRunIs(oObj As TaieProc) As Boolean

Dim tmpLoop As Integer

If oObj.IsReg = True Then
    For tmpLoop = 1 To MaxAllRunProc
        CheckRunIs AllRunProc(tmpLoop)
    Next tmpLoop
Else
    CheckRunIs oObj
End If
End Function

Public Function CheckRunIs(oObj As TaieProc) As Boolean
Dim hProcess As Long
Dim cb As Long
Dim cbNeeded As Long
Dim NumElements As Long
Dim ProcessIDs() As Long
Dim cbNeeded2 As Long
Dim NumElements2 As Long
Dim Modules(1 To 200) As Long
Dim lret As Long
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
          If InStr(1, UCase(sname), UCase(oObj.ProcName)) <> 0 Then
                '查詢 process code
                 hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
                                 Or PROCESS_VM_READ, 0, proc.th32ProcessID)
                 If hProcess <> 0 Then
                    '提示關閉
                    If DBMode = False Then
                        SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                    End If
                    Put2DBList "　　使用中"
                    If MsgBox(oObj.ProcCName & " 正在使用中，請先存檔再按下確定！", vbOKOnly, "警告！") = vbOK Then
                        Put2DBList "　　強制關閉"
                        '強制關閉
                         TerminateProcess hProcess, 0
                    End If
                    If DBMode = False Then
                        SetWindowPos frmAutoUpdate.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                    End If
                 End If
                 CloseHandle (hProcess)
          End If
          f = Process32Next(hSnap, proc)
        Loop
Case 2
   cb = 8
   cbNeeded = 96
   Do While cb <= cbNeeded
      cb = cb * 2
      ReDim ProcessIDs(cb / 4) As Long
      lret = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
   Loop
   NumElements = cbNeeded / 4

   For i = 1 To NumElements
      '查詢 process code
      hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
         Or PROCESS_VM_READ, 0, ProcessIDs(i))
      If hProcess <> 0 Then
          lret = EnumProcessModules(hProcess, Modules(1), 200, _
                                       cbNeeded2)
          If lret <> 0 Then
             ModuleName = Space(MAX_PATH)
             nSize = 500
             lret = GetModuleFileNameExA(hProcess, Modules(1), _
                             ModuleName, nSize)
            If InStr(1, UCase(ModuleName), UCase(oObj.ProcName)) <> 0 Then
                    If hProcess <> 0 Then
                       '提示關閉
                       If DBMode = False Then
                            SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                       End If
                       Put2DBList "　　使用中"
                       If MsgBox(oObj.ProcCName & " 正在使用中，請先存檔再按下確定！", vbOKOnly, "警告！") = vbOK Then
                           Put2DBList "　　強制關閉"
                           '強制關閉
                            TerminateProcess hProcess, 0
                       End If
                       If DBMode = False Then
                            SetWindowPos frmAutoUpdate.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                       End If
                    End If
                    CloseHandle (hProcess)
             End If
          End If
      End If
   lret = CloseHandle(hProcess)
   Next

End Select
End Function

Public Function CheckIsRun(oFileName As String) As Boolean
Dim hProcess As Long
Dim cb As Long
Dim cbNeeded As Long
Dim NumElements As Long
Dim ProcessIDs() As Long
Dim cbNeeded2 As Long
Dim NumElements2 As Long
Dim Modules(1 To 200) As Long
Dim lret As Long
Dim ModuleName As String
Dim nSize As Long
Dim f As Long, sname As String
Dim hSnap As Long, proc As PROCESSENTRY32
Dim i As Long
CheckIsRun = False

Select Case getVersion()
Case 2   'NT

'Modify by Morgan 2011/5/27 改用新的函數才能讀到其他使用者的程式(多人系統問題)
   CheckIsRun = CheckIsRunning(oFileName)
   Exit Function
   
'edit by nick 2005/01/10 因為其他使用者會抓不到
   cb = 8
   cbNeeded = 96
   Do While cb <= cbNeeded
      cb = cb * 2
      ReDim ProcessIDs(cb / 4) As Long
      lret = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
   Loop
   NumElements = cbNeeded / 4

   For i = 1 To NumElements
      '查詢 process code
      hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
         Or PROCESS_VM_READ, 0, ProcessIDs(i))
      If hProcess <> 0 Then
          lret = EnumProcessModules(hProcess, Modules(1), 200, _
                                       cbNeeded2)
          If lret <> 0 Then
             ModuleName = Space(MAX_PATH)
             nSize = 500
             lret = GetModuleFileNameExA(hProcess, Modules(1), _
                             ModuleName, nSize)
            If InStr(1, UCase(ModuleName), UCase(oFileName)) <> 0 And App.PrevInstance = True Then
                    If hProcess <> 0 Then
                        CheckIsRun = True
                        CloseHandle (hProcess)
                        Exit Function
                    End If
                    CloseHandle (hProcess)
             End If
          End If
      End If
   lret = CloseHandle(hProcess)
   Next
Case Else
    Dim hSnapShot As Long ', lret As Long
    Dim uProcess As PROCESSENTRY32
    
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapShot = 0 Then
        CheckIsRun = False
        Exit Function
    End If
    uProcess.dwSize = Len(uProcess)
    lret = ProcessFirst(hSnapShot, uProcess)
    Dim processID As Long
    GetWindowThreadProcessId frmAutoUpdate.hWnd, processID
    
    Do While lret
        If InStr(1, UCase(uProcess.szExeFile), UCase(oFileName)) <> 0 And uProcess.th32ProcessID <> processID Then
            CheckIsRun = True
            Exit Do
        End If
        lret = ProcessNext(hSnapShot, uProcess)
    Loop
    Call CloseHandle(hSnapShot)
End Select
End Function

Public Function getVersion() As Long
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer
   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)
   getVersion = osinfo.dwPlatformId
End Function

Public Function getVersion2() As Long
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer
   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)
   getVersion2 = Val(Trim(osinfo.dwMajorVersion) & Trim(osinfo.dwPlatformId))
End Function
'Add by Morgan 2011/2/9
Public Function getVersionNo() As String
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer
   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)
   getVersionNo = str$(osinfo.dwMajorVersion) + "." + LTrim(str(osinfo.dwMinorVersion))
End Function

Public Function StrZToStr(s As String) As String
   StrZToStr = Left$(s, Len(s) - 1)
End Function

Function CheckUpdateThisFile() As Boolean
    CheckUpdateThisFile = False
    Dim pData As WIN32_FIND_DATA
    Dim hFind As Long
    Dim ft As SYSTEMTIME
    Dim dtFileTime As Date
    Dim nLastError As Long
    TaieAllFileSize = 0
    If Len(UpdateThisProc.RemotePath) > 0 Then rcd (UpdateThisProc.RemotePath)
    pData.cFileName = String(MAX_PATH, 0)
    hFind = FtpFindFirstFile(hConnection, UpdateThisProc.ProcName, pData, 0, 0)
    nLastError = Err.LastDllError
    
    If hFind <> 0 Then
        '比對檔案，若日期及 size 不同則要更新
        FileTimeToSystemTime pData.ftLastWriteTime, ft
        'If ft.wYear >= UpdateThisProc.LocalDate.wYear Or ft.wMonth >= UpdateThisProc.LocalDate.wMonth Or ft.wDay >= UpdateThisProc.LocalDate.wDay Or ft.wHour >= UpdateThisProc.LocalDate.wHour Or ft.wMinute >= UpdateThisProc.LocalDate.wMinute Or pData.nFileSizeLow <> UpdateThisProc.oFileSize Then
        If IsNew(ft, UpdateThisProc.LocalDate) Then
            UpdateThisProc.nFileSize = pData.nFileSizeLow
            UpdateThisProc.IsDownOk = False
            TaieAllFileSize = pData.nFileSizeLow
            UpdateThisProc.RemoteDate = ft
            frmAutoUpdate.ProgressBar1.Min = 0
            'frmAutoUpdate.ProgressBar1.Max = Format(IIf(TaieAllFileSize = 0, 102400, TaieAllFileSize) / 1024, "###########0")
            frmAutoUpdate.ProgressBar1.Max = Format(IIf(TaieAllFileSize = 0, 102400, TaieAllFileSize), "###########0")
            frmAutoUpdate.ProgressBar1.Value = 0
            frmAutoUpdate.lbldownState.Caption = "0 Ｋ / " & Format(TaieAllFileSize / 1024, "###,###,###,##0") & "  Ｋ"
            CheckUpdateThisFile = True
        End If
    End If
    InternetCloseHandle hFind
End Function

Public Function CheckMeSize()
Dim oIjk As Integer
Dim ShFileOP As SHFILEOPSTRUCT

    '檔案 size 不合
    Put2DBList "　　檢查本支程式檔案大小"
    If FileLen(UpdateThisProc.LocalPath & UpdateThisProc.ProcName & ".f") = UpdateThisProc.nFileSize Then
        UpdateThisProc.IsDownOk = True
        Put2DBList "　　檔案正確"
    Else
        UpdateThisProc.IsDownOk = False
        Put2DBList "　　錯誤，大小不符"
        Put2DBList "　　刪除暫存"
        '刪除
        ShFileOP.wFunc = FO_DELETE
        ShFileOP.pFrom = UpdateThisProc.LocalPath & UpdateThisProc.ProcName & "file" & Chr(0)
        ShFileOP.fFlags = FOF_NOCONFIRMATION + FOF_SILENT
        SHFileOperation ShFileOP
        frmAutoUpdate.lblState.Caption = UpdateThisProc.ProcCName & "  下載失敗，重新下載..."
        TaieAllFileSize = UpdateThisProc.nFileSize
        TaieAllFileSize_OK = 0
        frmAutoUpdate.ProgressBar1.Value = 0
        Put2DBList "　　重新下載"
        DownLoadFileToTemp UpdateThisProc
        Put2DBList "　　檢查重新下載的大小"
        CheckMeSize
    End If
End Function

Public Function UpdateMe()
'再開另一個trared  執行程式
    Dim strExc As String
    Dim lngRCode As Long
    Dim udtStartupInfo As STARTUPINFO
    Dim udtProcessInfo As PROCESS_INFORMATION
    Dim strCmd As String
    
    udtStartupInfo.cb = Len(udtStartupInfo)
    udtStartupInfo.dwFlags = STARTF_USESHOWWINDOW
    udtStartupInfo.wShowWindow = SW_SHOWDEFAULT

    udtProcessInfo.dwProcessId = 0&
    udtProcessInfo.dwThreadId = 0&
    udtProcessInfo.hProcess = 0&
    udtProcessInfo.hThread = 0&
    strExc = UpdateThisProc.LocalPath & "UpdateMe.exe"
    
    'Added by Morgan 2016/12/22 Win7 的 UAC 會擋名稱有 update 的程式
    If Dir(UpdateThisProc.LocalPath & "UpdMe.exe") <> "" Then
      strExc = UpdateThisProc.LocalPath & "UpdMe.exe"
    End If
    'end 2016/12/22
    
    'Modified by Morgan 2015/2/26 NAS FTP Server 大小寫有別
    'strCmd = App.Path & "\" & App.EXEName & ".exe|" & Format(UpdateThisProc.RemoteDate.wYear, "0000") & Format(UpdateThisProc.RemoteDate.wMonth, "00") & Format(UpdateThisProc.RemoteDate.wDay, "00") & Format(UpdateThisProc.RemoteDate.wHour, "00") & Format(UpdateThisProc.RemoteDate.wMinute, "00") & Format(UpdateThisProc.RemoteDate.wSecond, "00") & Trim(UpdateThisProc.RemoteDate.wMilliseconds)
    strCmd = App.Path & "\" & UpdateThisProc.ProcName & "|" & Format(UpdateThisProc.RemoteDate.wYear, "0000") & Format(UpdateThisProc.RemoteDate.wMonth, "00") & Format(UpdateThisProc.RemoteDate.wDay, "00") & Format(UpdateThisProc.RemoteDate.wHour, "00") & Format(UpdateThisProc.RemoteDate.wMinute, "00") & Format(UpdateThisProc.RemoteDate.wSecond, "00") & Trim(UpdateThisProc.RemoteDate.wMilliseconds)
    'end 2015/2/26
    Clipboard.SetText strCmd
    lngRCode = CreateProcess(strExc, _
                             "", _
                             0&, _
                             0&, _
                             0&, _
                             0&, _
                             0&, _
                             0&, _
                             udtStartupInfo, _
                             udtProcessInfo)
    If lngRCode = 0& Then
        'MsgBox "Failed to call CreateProcess() function." '
        frmAutoUpdate.lblState.Caption = "呼叫程式失敗..."
        frmAutoUpdate.CmdEnd.Visible = True
    Else
        CloseHandle udtProcessInfo.hProcess
        udtProcessInfo.hProcess = 0&
        CloseHandle udtProcessInfo.hThread
        udtProcessInfo.hThread = 0&
    End If
'Shell strExc, vbHide
    '解 Ctrl + Alt + Del
SystemParametersInfo SPI_SCREENSAVERRUNNING, False, frmAutoUpdate.pOld, 0

End Function

Public Function IsNew(oDate1 As SYSTEMTIME, oDate2 As SYSTEMTIME) As Boolean
Dim StrDate1 As String
Dim StrDate2 As String
'StrDate1 = Format(oDate1.wYear, "0000") & Format(oDate1.wMonth, "00") & Format(oDate1.wDay, "00") & Format(oDate1.wHour, "00") & Format(oDate1.wMinute, "00") & Format(oDate1.wSecond, "00")
'StrDate2 = Format(oDate2.wYear, "0000") & Format(oDate2.wMonth, "00") & Format(oDate2.wDay, "00") & Format(oDate2.wHour, "00") & Format(oDate2.wMinute, "00") & Format(oDate2.wSecond, "00")
StrDate1 = Format(oDate1.wMonth, "00") & Format(oDate1.wDay, "00") '& Format(oDate1.wHour, "00") & Format(oDate1.wMinute, "00") '& Format(oDate1.wSecond, "00")
StrDate2 = Format(oDate2.wMonth, "00") & Format(oDate2.wDay, "00") '& Format(oDate2.wHour, "00") & Format(oDate2.wMinute, "00") '& Format(oDate2.wSecond, "00")
If StrDate1 <> StrDate2 Then
    IsNew = True
Else
    IsNew = False
End If

End Function

'*********************************************************************************
' 以下為第二版所用 function
'*********************************************************************************
'下載清單
Function DownFileList()
Dim hFile As Long
Dim strCommand As String
Dim bDoLoop As Boolean
Dim sBuffer As Long
Dim lNumberOfBytesRead As Long
Dim byteBuffer(20479) As Byte
Dim ReadyBuffer() As Byte
Dim F1 As Integer
Dim oIjk As Long
Dim ShFileOP As SHFILEOPSTRUCT
Put2DBList "　　刪除舊清單"
'先刪除舊檔
If Dir(App.Path & "\" & "filelist.lst") <> "" Then
    ShFileOP.wFunc = FO_DELETE
    ShFileOP.pFrom = App.Path & "\" & "filelist.lst" & Chr(0)
    ShFileOP.fFlags = FOF_NOCONFIRMATION + FOF_SILENT
    SHFileOperation ShFileOP
End If
Put2DBList "　　切換目錄：" & cRemotePath

rcd cRemotePath

sBuffer = 20480
Put2DBList "　　建立清單連線"
'Modified by Morgan 2018/8/3 +強制重新下載參數(INTERNET_FLAG_RELOAD=&H80000000)
'hFile = FtpOpenFile(hConnection, "filelist.lst", &H80000000, INTERNET_FLAG_TRANSFER_BINARY, 0)
hFile = FtpOpenFile(hConnection, "filelist.lst", &H80000000, INTERNET_FLAG_TRANSFER_BINARY + &H80000000, 0)
If hFile = 0 Then
    UpdateProgramData , "-filelist.lst" 'Added by Morgan 2020/9/9
    Put2DBList "　　建立失敗"
    Exit Function
End If
Put2DBList "　　開始下載清單"
F1 = FreeFile
Open App.Path & "\" & "filelist.lst" For Binary As F1
bDoLoop = True
While bDoLoop
    bDoLoop = InternetReadFileByte(hFile, VarPtr(byteBuffer(0)), sBuffer, lNumberOfBytesRead)
    If Not CBool(lNumberOfBytesRead) Then
        bDoLoop = False
    Else
        ReDim ReadyBuffer(lNumberOfBytesRead - 1) As Byte
        If lNumberOfBytesRead <> sBuffer Then
            For oIjk = 0 To (lNumberOfBytesRead - 1)
                ReadyBuffer(oIjk) = byteBuffer(oIjk)
                DoEvents
            Next oIjk
        Else
            ReadyBuffer = byteBuffer
        End If
        Put #F1, , ReadyBuffer
    End If
Wend
Close F1
Put2DBList "　　下載完成"
InternetCloseHandle hFile
UpdateProgramData , "+filelist.lst" 'Added by Morgan 2020/9/10
Erase byteBuffer
Erase ReadyBuffer
End Function

Function CheckUpdateThisFile2() As Boolean
    CheckUpdateThisFile2 = False

    Dim F1 As Integer
    Dim Ftp_FileData As String
    Dim Ftp_FileName As String
    Dim Ftp_FileSize As String
    Dim Ftp_FileDate As String
    Dim tmpFtp As Variant
    F1 = FreeFile
    TaieAllFileSize = 0
    Open App.Path & "\" & "filelist.lst" For Input As F1
    Do While Not EOF(F1)
        Input #F1, Ftp_FileData
        tmpFtp = Split(Ftp_FileData, "||")
        'Modified by Morgan 2015/2/25 NAS FTP Server 大小寫有分別
        'Ftp_FileName = Trim(UCase(tmpFtp(0)))
        'If Ftp_FileName = UCase(App.EXEName & ".exe") Then
        Ftp_FileName = Trim(tmpFtp(0))
        If UCase(Ftp_FileName) = UCase(App.EXEName & ".exe") Then
        'end 2015/2/25
            Ftp_FileSize = Trim(tmpFtp(1))
            Ftp_FileDate = Trim(tmpFtp(2))
            'If Val(Ftp_FileSize) <> UpdateThisProc.oFileSize Or Ftp_FileDate > SysTimeToStr(UpdateThisProc.LocalDate) Then
            'Modify by Morgan 2010/10/19 允許1秒的誤差值(95會發生)
            'If Ftp_FileDate > SysTimeToStr(UpdateThisProc.LocalDate) Then
            If fnAddSec(Ftp_FileDate, -1) > SysTimeToStr(UpdateThisProc.LocalDate) Then
                UpdateThisProc.ProcName = Ftp_FileName 'Added by Morgan 2015/2/26
                UpdateThisProc.nFileSize = Ftp_FileSize
                UpdateThisProc.IsDownOk = False
                TaieAllFileSize = Val(Ftp_FileSize)
                UpdateThisProc.RemoteDate = StrTimeToSys(Ftp_FileDate)
                frmAutoUpdate.ProgressBar1.Min = 0
                'frmAutoUpdate.ProgressBar1.Max = Format(IIf(TaieAllFileSize = 0, 102400, TaieAllFileSize) / 1024, "###########0")
                frmAutoUpdate.ProgressBar1.Max = Format(IIf(TaieAllFileSize = 0, 102400, TaieAllFileSize), "###########0")
                frmAutoUpdate.ProgressBar1.Value = 0
                frmAutoUpdate.lbldownState.Caption = "0 Ｋ / " & Format(TaieAllFileSize / 1024, "###,###,###,##0") & "  Ｋ"
                CheckUpdateThisFile2 = True
            End If
            Exit Do
        End If
    Loop
    Close #F1
End Function

Function SysTimeToStr(oFt As SYSTEMTIME) As String
Dim tmpDate As Date
If oFt.wYear > 0 Then
   tmpDate = CDate(oFt.wYear & "/" & oFt.wMonth & "/" & oFt.wDay & " " & oFt.wHour & ":" & oFt.wMinute & ":" & oFt.wSecond)
   SysTimeToStr = Format(tmpDate, "YYYYMMDDHHmmss")
End If
End Function

Function StrTimeToSys(oStr As String) As SYSTEMTIME
StrTimeToSys.wYear = Mid(oStr, 1, 4)
StrTimeToSys.wMonth = Mid(oStr, 5, 2)
StrTimeToSys.wDay = Mid(oStr, 7, 2)
StrTimeToSys.wHour = Mid(oStr, 9, 2)
StrTimeToSys.wMinute = Mid(oStr, 11, 2)
StrTimeToSys.wSecond = Mid(oStr, 13, 2)
End Function
Function CheckUpdateFile2(oObj As TaieProc) As Boolean

    Dim F1 As Integer
    Dim Ftp_FileData As String
    Dim Ftp_FileName As String
    Dim Ftp_FileSize As String
    Dim Ftp_FileDate As String
    Dim tmpFtp As Variant
    
    F1 = FreeFile
    Open App.Path & "\" & "filelist.lst" For Input As F1
    Do While Not EOF(F1)
        Input #F1, Ftp_FileData
        tmpFtp = Split(Ftp_FileData, "||")
        'Modified by Morgan 2015/2/25 新 FTP 大小寫有分
        'Ftp_FileName = Trim(UCase(tmpFtp(0)))
        'If Ftp_FileName = UCase(oObj.ProcName) Then
        Ftp_FileName = Trim(tmpFtp(0))
        If UCase(Ftp_FileName) = UCase(oObj.ProcName) Then
        'end 2015/2/25
            Ftp_FileSize = Trim(tmpFtp(1))
            Ftp_FileDate = Trim(tmpFtp(2))
            DoEvents
            Put2DBList "　　比對" & Ftp_FileName
            'If Val(Ftp_FileSize) <> oObj.oFileSize Or Ftp_FileDate > SysTimeToStr(oObj.LocalDate) Then
            'Modify by Morgan 2010/10/19 允許1秒的誤差值(95會發生)
            'If Ftp_FileDate > SysTimeToStr(oObj.LocalDate) Then
            If fnAddSec(Ftp_FileDate, -1) > SysTimeToStr(oObj.LocalDate) Then
                MaxUpdateTaieProc = MaxUpdateTaieProc + 1
                ReDim Preserve UpdateTaieProc(MaxUpdateTaieProc) As TaieProc
                UpdateTaieProc(MaxUpdateTaieProc).IsReg = oObj.IsReg
                UpdateTaieProc(MaxUpdateTaieProc).LocalPath = oObj.LocalPath
                UpdateTaieProc(MaxUpdateTaieProc).ProcCName = oObj.ProcCName
                'Modified by Morgan 2015/2/25 NAS FTP Server 大小寫有分,以清單的檔名為準
                'UpdateTaieProc(MaxUpdateTaieProc).ProcName = oObj.ProcName
                UpdateTaieProc(MaxUpdateTaieProc).ProcName = Ftp_FileName
                'end 2015/2/25
                UpdateTaieProc(MaxUpdateTaieProc).RemotePath = oObj.RemotePath
                UpdateTaieProc(MaxUpdateTaieProc).nFileSize = Val(Ftp_FileSize)
                UpdateTaieProc(MaxUpdateTaieProc).IsDownOk = False
                UpdateTaieProc(MaxUpdateTaieProc).RemoteDate = StrTimeToSys(Ftp_FileDate)
                TaieAllFileSize = TaieAllFileSize + Val(Ftp_FileSize)
            End If
            Exit Do
        End If
    Loop
    Close #F1
End Function

'檢查強制下載的檔案
'FileList.exe||20480||20041014100332||path
Function CheckUpdateNewFile() As Boolean
    CheckUpdateNewFile = False

    Dim F1 As Integer
    Dim Ftp_FileData As String
    Dim Ftp_FileName As String
    Dim Ftp_FileSize As String
    Dim Ftp_FileDate As String
    Dim tmpFtp As Variant
    Dim bolYes As Boolean '是否下載 Added by Morgan 2014/5/9
    Dim Local_FileDate As String 'Added by Morgan 2024/9/4
    
    F1 = FreeFile
    TaieAllFileSize = 0
    MaxNewTaieProc = 0
    Put2DBList "　　檢查強制安裝"
    Open App.Path & "\" & "filelist.lst" For Input As F1
    Do While Not EOF(F1)
        Input #F1, Ftp_FileData
        tmpFtp = Split(Ftp_FileData, "||")
        'Modified by Morgan 2015/2/25 新的 FTP Server 大小寫有分別
        'Ftp_FileName = Replace(Trim(UCase(tmpFtp(0))), "*", "")
        Ftp_FileName = Replace(Trim(tmpFtp(0)), "*", "")
        'add by nickc 2006/01/02 nt/2000/xp 不做
        'If getVersion <> 2 And InStr(1, UCase(Ftp_FileName), "EUDC") = 0 Then
        If InStr(1, UCase(Ftp_FileName), "EUDC") = 0 Then
            bolYes = False 'Added by Morgan 2014/5/9
            '新檔案前面要 * 號判斷，且格式稍有不同
            If Mid(Ftp_FileData, 1, 1) = "*" Then
            
            'Added by Morgan 2014/5/9 +pdf合併程式不存在時要下載
               bolYes = True
            ElseIf UCase(Ftp_FileName) = UCase("pdftk.exe") Or UCase(Ftp_FileName) = UCase("libiconv2.dll") Then
               'Modified by Morgan 2024/9/4 pdf合併程式也要更新
               'If Dir(SysPath & Ftp_FileName) = "" And Dir(App.Path & "\" & Ftp_FileName) = "" Then
               '   bolYes = True
               If Dir(App.Path & "\" & Ftp_FileName) = "" Then
                  bolYes = True
               Else
                  Ftp_FileDate = Trim(tmpFtp(2))
                  Local_FileDate = PUB_GetFileDateTime(App.Path & "\" & Ftp_FileName)
                  If fnAddSec(Ftp_FileDate, -1) > Local_FileDate Then
                     bolYes = True
                  End If
               End If
               'end 2024/9/4
            End If
            If bolYes = True Then
            'end 2014/5/9
            
                MaxNewTaieProc = MaxNewTaieProc + 1
                ReDim Preserve NewTaieProc(MaxNewTaieProc) As TaieProc
                Ftp_FileSize = Trim(tmpFtp(1))
                Ftp_FileDate = Trim(tmpFtp(2))
                NewTaieProc(MaxNewTaieProc).nFileSize = Ftp_FileSize
                NewTaieProc(MaxNewTaieProc).ProcName = Ftp_FileName
                NewTaieProc(MaxNewTaieProc).IsReg = False
                NewTaieProc(MaxNewTaieProc).IsDownOk = False
                NewTaieProc(MaxNewTaieProc).RemotePath = cRemotePath
                TaieAllFileSize = TaieAllFileSize + Val(Ftp_FileSize)
                NewTaieProc(MaxNewTaieProc).RemoteDate = StrTimeToSys(Ftp_FileDate)
                '要放入 client 端的位置，空白表示跟 自動更新程式 同一位置，最後要加 \  若是空白就不用
                '*FileList.exe||20480||20041014100332||C:\Program Files\TEPATENT\
                
               'Modified by Morgan 2014/5/9
               'NewTaieProc(MaxNewTaieProc).LocalPath = IIf(Trim(tmpFtp(3)) = "", App.Path & "\", Trim(tmpFtp(3)))
               If UBound(tmpFtp) >= 3 Then
                  NewTaieProc(MaxNewTaieProc).LocalPath = IIf(Trim(tmpFtp(3)) = "", App.Path & "\", Trim(tmpFtp(3)))
               Else
                  NewTaieProc(MaxNewTaieProc).LocalPath = App.Path & "\"
               End If
               'end 2014/5/9
               
                frmAutoUpdate.ProgressBar1.Min = 0
                'frmAutoUpdate.ProgressBar1.Max = Format(IIf(TaieAllFileSize = 0, 102400, TaieAllFileSize) / 1024, "###########0")
                frmAutoUpdate.ProgressBar1.Max = Format(IIf(TaieAllFileSize = 0, 102400, TaieAllFileSize), "###########0")
                frmAutoUpdate.ProgressBar1.Value = 0
                frmAutoUpdate.lbldownState.Caption = "0 Ｋ / " & Format(TaieAllFileSize / 1024, "###,###,###,##0") & "  Ｋ"
                CheckUpdateNewFile = True
            End If
        End If
    Loop
    Close #F1
End Function

Public Function DownLoadAllNewFileToTemp() As Boolean
Dim oIjk As Integer
TaieAllFileSize_OK = 0
For oIjk = 1 To MaxNewTaieProc
    'rcd ("\")
    Call DownLoadFileToTemp(NewTaieProc(oIjk))
Next oIjk
End Function

Public Function CheckNewFileSize()
Dim oIjk As Integer
Dim ShFileOP As SHFILEOPSTRUCT

For oIjk = 1 To MaxNewTaieProc
    '檔案 size 不合
    Put2DBList "　　檢查 " & NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName & "是否存在"
    If Dir(NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName & ".f") <> "" Then
        Put2DBList "　　檢查 " & NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName & "檔案大小"
        If FileLen(NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName & ".f") = NewTaieProc(oIjk).nFileSize Then
            NewTaieProc(oIjk).IsDownOk = True
            Put2DBList "　　檔案正確"
        Else
            NewTaieProc(oIjk).IsDownOk = False
            Put2DBList "　　錯誤，大小不符"
            Put2DBList "　　刪除暫存"
            '刪除
            ShFileOP.wFunc = FO_DELETE
            ShFileOP.pFrom = NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName & ".f" & Chr(0)
            ShFileOP.fFlags = FOF_NOCONFIRMATION + FOF_SILENT
            SHFileOperation ShFileOP
            If DBMode = False Then
                SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
            End If
            Put2DBList "　　重新下載??"
            If MsgBox(NewTaieProc(oIjk).ProcCName & "  下載失敗，是否重新下載此檔！", vbYesNo, "錯誤！") = vbYes Then
                TaieAllFileSize = NewTaieProc(oIjk).nFileSize
                TaieAllFileSize_OK = 0
                frmAutoUpdate.ProgressBar1.Value = 0
                Put2DBList "　　重新下載"
                DownLoadFileToTemp NewTaieProc(oIjk)
                Put2DBList "　　檢查重新下載的大小"
                CheckSize
                Exit For
            End If
            If DBMode = False Then
                SetWindowPos frmAutoUpdate.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
            End If
        End If
    Else
        Put2DBList "　　不存在，重新下載??"
        NewTaieProc(oIjk).IsDownOk = False
        If DBMode = False Then
            SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
        End If
        Put2DBList "　　重新下載??"
        If MsgBox(NewTaieProc(oIjk).ProcCName & "  下載失敗，是否重新下載此檔！", vbYesNo, "錯誤！") = vbYes Then
            TaieAllFileSize = NewTaieProc(oIjk).nFileSize
            TaieAllFileSize_OK = 0
            frmAutoUpdate.ProgressBar1.Value = 0
            Put2DBList "　　重新下載"
            DownLoadFileToTemp NewTaieProc(oIjk)
            Put2DBList "　　檢查重新下載的大小"
            CheckSize
            Exit For
        End If
        If DBMode = False Then
            SetWindowPos frmAutoUpdate.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
        End If
    End If
Next oIjk
End Function

Public Function MoveNewFileToReady() As Boolean
Dim oIjk As Integer
Dim ShFileOP As SHFILEOPSTRUCT
Dim lpct As FILETIME, lplac As FILETIME, lplwr As FILETIME
Dim ofs As OFSTRUCT
Dim hFile As Long
Dim tZone As TIME_ZONE_INFORMATION
Dim bias As Long
Dim writedate As Date
Dim ft As SYSTEMTIME
Dim lngRCode As Long
Dim udtStartupInfo As STARTUPINFO
Dim udtProcessInfo As PROCESS_INFORMATION
Dim strCmd As String
Dim strExc As String
Dim updateFontCount As Integer
For oIjk = 1 To MaxNewTaieProc
    '檢查是否下載成功
    If NewTaieProc(oIjk).IsDownOk = True Then
        '檢查要更新的檔案是否正在執行，若是 dll 或是 Active X EXE 時，要將其他的TE 開頭的程式斷線
        Put2DBList "　　檢查 " & NewTaieProc(oIjk).ProcName & "是否正在執行中"
        CheckAllRunIs NewTaieProc(oIjk)
        '檢查是否要註冊，Dll 才要
        If NewTaieProc(oIjk).IsReg = True And UCase(Right(NewTaieProc(oIjk).ProcName, 3)) = "DLL" Then
            '反註冊  改成要等待
'            Shell "command.com /c regsvr32.exe """ & NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName & """ /u /s", vbHide
            strExc = "regsvr32.exe"
            strCmd = " """ & NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName & """ /u /s"
            udtStartupInfo.cb = Len(udtStartupInfo)
            udtStartupInfo.dwFlags = STARTF_USESHOWWINDOW
            udtStartupInfo.wShowWindow = SW_HIDE
            udtProcessInfo.dwProcessId = 0&
            udtProcessInfo.dwThreadId = 0&
            udtProcessInfo.hProcess = 0&
            udtProcessInfo.hThread = 0&
            lngRCode = CreateProcess(vbNullString, _
                                     strExc & strCmd, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     udtStartupInfo, _
                                     udtProcessInfo)
            If lngRCode = 0& Then
                If DBMode = False Then
                    SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                    MsgBox "Failed to call CreateProcess() function."
                    SetWindowPos frmAutoUpdate.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                Else
                    Put2DBList "　　呼叫 CreateProcess 失敗"
                End If
            Else
                Do
                    lngRCode = WaitForSingleObject(udtProcessInfo.hProcess, 300)
                    If lngRCode = 0& Then
                        Exit Do
                    End If
                    DoEvents
                Loop
                CloseHandle udtProcessInfo.hProcess
                udtProcessInfo.hProcess = 0&
                CloseHandle udtProcessInfo.hThread
                udtProcessInfo.hThread = 0&
            End If
        End If
        Put2DBList "　　開始更名成 f2"
        updateFontCount = 1
        Do While Dir(NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName) <> ""
            If updateFontCount >= 10 Then
                Put2DBList "　　更名 f2 失敗"
                MoveNewFileToReady = False
                'add by nickc 2008/03/12 失敗要還原
                ShFileOP.wFunc = FO_MOVE
                ShFileOP.pFrom = Left(NewTaieProc(oIjk).LocalPath, Len(NewTaieProc(oIjk).LocalPath) - 1) & "\" & NewTaieProc(oIjk).ProcName & ".f2" + Chr(0)
                ShFileOP.pTo = NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName
                ShFileOP.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_SILENT
                SHFileOperation ShFileOP
                Exit Function
            End If
                Put2DBList "　　" & updateFontCount
                updateFontCount = updateFontCount + 1
                ShFileOP.wFunc = FO_MOVE
                ShFileOP.pFrom = NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName + Chr(0)
                ShFileOP.pTo = Left(NewTaieProc(oIjk).LocalPath, Len(NewTaieProc(oIjk).LocalPath) - 1) & "\" & NewTaieProc(oIjk).ProcName & ".f2"
                ShFileOP.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_SILENT
                SHFileOperation ShFileOP
        Loop
        Put2DBList "　　開始將 f 更新成正確名稱"
        updateFontCount = 1
        'Modified by Morgan 2016/12/22 有發生只剩.f的情形,加判斷執行檔要存在才繼續
        'Do While Dir(NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName & ".f") <> ""
        Do While Dir(NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName & ".f") <> "" Or Dir(NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName) = ""
        'end 2016/12/22
            If updateFontCount >= 10 Then
                Put2DBList "　　將 f 更新成正確名稱失敗"
                MoveNewFileToReady = False
                'add by nickc 2008/03/12 失敗要還原
                ShFileOP.wFunc = FO_MOVE
                ShFileOP.pFrom = NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName + Chr(0)
                ShFileOP.pTo = NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName & ".f"
                ShFileOP.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_SILENT
                SHFileOperation ShFileOP
                Exit Function
            End If
            Put2DBList "　　" & updateFontCount
            updateFontCount = updateFontCount + 1
            ShFileOP.wFunc = FO_MOVE
            ShFileOP.pFrom = NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName & ".f" + Chr(0)
            ShFileOP.pTo = NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName
            ShFileOP.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_SILENT
            SHFileOperation ShFileOP
            DoEvents
        Loop
        Put2DBList "　　刪除 f2"
        updateFontCount = 1
        Do While Dir(NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName & ".f2") <> ""
            If updateFontCount >= 10 Then
                MoveNewFileToReady = False
                Put2DBList "　　f2 刪除失敗"
                Exit Function
            End If
            Put2DBList "　　" & updateFontCount
            updateFontCount = updateFontCount + 1
            ShFileOP.wFunc = FO_DELETE
            ShFileOP.pFrom = NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName & ".f2" + Chr(0)
            ShFileOP.fFlags = FOF_NOCONFIRMATION + FOF_SILENT
            SHFileOperation ShFileOP
        Loop
        Put2DBList "　　更改日期時間與伺服器相同"
        '更改日期
        Call GetTimeZoneInformation(tZone)
        bias = tZone.bias
        writedate = CDate(NewTaieProc(oIjk).RemoteDate.wYear & "/" & NewTaieProc(oIjk).RemoteDate.wMonth & "/" & NewTaieProc(oIjk).RemoteDate.wDay & " " & NewTaieProc(oIjk).RemoteDate.wHour & ":" & NewTaieProc(oIjk).RemoteDate.wMinute & ":" & NewTaieProc(oIjk).RemoteDate.wSecond) + TimeSerial(0, bias, 0)
        ft.wYear = Year(writedate)
        ft.wMonth = Month(writedate)
        ft.wDay = Day(writedate)
        ft.wHour = Hour(writedate)
        ft.wMinute = Minute(writedate)
        ft.wSecond = Second(writedate)
        ft.wDayOfWeek = Weekday(writedate)
        ft.wMilliseconds = NewTaieProc(oIjk).RemoteDate.wMilliseconds
        hFile = OpenFile(NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName, ofs, OF_READWRITE)
        'NewTaieProc(oIjk).RemoteDate.wMinute = NewTaieProc(oIjk).RemoteDate.wMinute + bias
        Call SystemTimeToFileTime(ft, lplwr)
        '更動hFile的時間，第2個參數改Create DateTime
        '第3個參數改Last Access DateTime
        '第四個參數改Last Modify DateTime
        Call SetFileTime(hFile, ByVal 0, ByVal 0, lplwr) '只更改第三個(最後寫入)時間
        Call CloseHandle(hFile) '關閉檔案
        '檢查是否要註冊，Dll 才要
        If NewTaieProc(oIjk).IsReg = True And UCase(Right(NewTaieProc(oIjk).ProcName, 3)) = "DLL" Then
            '註冊
'            Shell "command.com /c regsvr32.exe """ & NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName & """ /s", vbHide
            strExc = "regsvr32.exe"
            strCmd = " """ & NewTaieProc(oIjk).LocalPath & NewTaieProc(oIjk).ProcName & """ /s"
            udtStartupInfo.cb = Len(udtStartupInfo)
            udtStartupInfo.dwFlags = STARTF_USESHOWWINDOW
            udtStartupInfo.wShowWindow = SW_HIDE
            udtProcessInfo.dwProcessId = 0&
            udtProcessInfo.dwThreadId = 0&
            udtProcessInfo.hProcess = 0&
            udtProcessInfo.hThread = 0&
            lngRCode = CreateProcess(vbNullString, _
                                     strExc & strCmd, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     0&, _
                                     udtStartupInfo, _
                                     udtProcessInfo)
            If lngRCode = 0& Then
                If DBMode = False Then
                    SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                    MsgBox "Failed to call CreateProcess() function."
                    SetWindowPos frmAutoUpdate.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                Else
                    Put2DBList "　　呼叫 CreateProcess 失敗"
                End If
            Else
                Do
                    lngRCode = WaitForSingleObject(udtProcessInfo.hProcess, 300)
                    If lngRCode = 0& Then
                        Exit Do
                    End If
                    DoEvents
                Loop
                CloseHandle udtProcessInfo.hProcess
                udtProcessInfo.hProcess = 0&
                CloseHandle udtProcessInfo.hThread
                udtProcessInfo.hThread = 0&
            End If
        End If
    End If
Next oIjk
End Function


'將 byte 轉字串，可以定區間
Public Function RecStr(ByRef strData, oStart As Long, oEnd As Long) As String
Dim i As Long
Dim tmpStr As String
Dim otByte() As Byte
tmpStr = ""
For i = oStart To oEnd
    tmpStr = tmpStr & ChrB(strData(i))
Next i
RecStr = tmpStr
End Function

'將字串轉 byte
Public Sub StrToByte(ByRef strData, oTransStr As String)
Dim i As Long
Dim iLen As Long
Dim otByte() As Byte
ReDim otByte(LenB(oTransStr) - 1) As Byte
iLen = 0
For i = 1 To LenB(oTransStr)
        otByte(i - 1) = AscB(MidB(oTransStr, i, 1))
Next i
strData = otByte
End Sub
'add by nickc 2008/03/17
Function CheckUpdateEUDCFile() As Boolean
    CheckUpdateEUDCFile = False

Dim oIjk As Integer


MaxUpdateTaieEUDC = 0
TaieEUDCFileSize = 0
For oIjk = 1 To MaxTaieEUDC
    Call CheckUpdateEUDCFile2(AllTaieEUDC(oIjk))
Next oIjk
If MaxUpdateTaieEUDC <> 0 Then
    IsHaveEudc = True
    CheckUpdateEUDCFile = True
    frmAutoUpdate.ProgressBar1.Min = 0
    frmAutoUpdate.ProgressBar1.Max = Format(IIf(TaieEUDCFileSize = 0, 102400, TaieEUDCFileSize), "###########0")
    frmAutoUpdate.ProgressBar1.Value = 0
    frmAutoUpdate.lbldownState.Caption = "0 Ｋ / " & Format(TaieEUDCFileSize / 1024, "###,###,###,##0") & "  Ｋ"
End If
End Function
Function CheckUpdateEUDCFile2(oObj As TaieProc) As Boolean

    Dim F1 As Integer
    Dim Ftp_FileData As String
    Dim Ftp_FileName As String
    Dim Ftp_FileSize As String
    Dim Ftp_FileDate As String
    Dim tmpFtp As Variant
    F1 = FreeFile
    Put2DBList "　　檢查造字是否要更新"
    Open App.Path & "\" & "filelist.lst" For Input As F1
    Do While Not EOF(F1)
        Input #F1, Ftp_FileData
        tmpFtp = Split(Ftp_FileData, "||")
        'Modified by Morgan 2015/2/25 NAS FTP Server 大小寫有分別
        'Ftp_FileName = Trim(UCase(tmpFtp(0)))
        'If Ftp_FileName = UCase(oObj.ProcName) Then
        Ftp_FileName = Trim(tmpFtp(0))
        If UCase(Ftp_FileName) = UCase(oObj.ProcName) Then
        'end 2015/2/25
            Ftp_FileSize = Trim(tmpFtp(1))
            Ftp_FileDate = Trim(tmpFtp(2))
            'If Val(Ftp_FileSize) <> oObj.oFileSize Or Ftp_FileDate > SysTimeToStr(oObj.LocalDate) Then
            'Modify by Morgan 2010/10/19 允許1秒的誤差值(95會發生)
            'If Ftp_FileDate > SysTimeToStr(oObj.LocalDate) Then
            If fnAddSec(Ftp_FileDate, -1) > SysTimeToStr(oObj.LocalDate) Then
                MaxUpdateTaieEUDC = MaxUpdateTaieEUDC + 1
                ReDim Preserve UpdateTaieEUDC(MaxUpdateTaieEUDC) As TaieProc
                UpdateTaieEUDC(MaxUpdateTaieEUDC).IsReg = oObj.IsReg
                UpdateTaieEUDC(MaxUpdateTaieEUDC).LocalPath = oObj.LocalPath
                UpdateTaieEUDC(MaxUpdateTaieEUDC).ProcCName = oObj.ProcCName
                'Modified by Morgan 2015/2/25 NAS FTP Server 大小寫有分,以清單的檔名為準
                'UpdateTaieEUDC(MaxUpdateTaieEUDC).ProcName = oObj.ProcName
                UpdateTaieEUDC(MaxUpdateTaieEUDC).ProcName = Ftp_FileName
                'end 2015/2/25
                UpdateTaieEUDC(MaxUpdateTaieEUDC).RemotePath = oObj.RemotePath
                UpdateTaieEUDC(MaxUpdateTaieEUDC).nFileSize = Val(Ftp_FileSize)
                UpdateTaieEUDC(MaxUpdateTaieEUDC).IsDownOk = False
                UpdateTaieEUDC(MaxUpdateTaieEUDC).RemoteDate = StrTimeToSys(Ftp_FileDate)
                TaieEUDCFileSize = TaieEUDCFileSize + Val(Ftp_FileSize)
            End If
            Exit Do
        End If
    Loop
    Close #F1
End Function
Public Function DownLoadEudcNewFileToTemp() As Boolean
Dim oIjk As Integer
TaieEUDCFileSize_OK = 0
For oIjk = 1 To MaxUpdateTaieEUDC
    'rcd ("\")
    Call DownLoadFileToTemp_Eudc(UpdateTaieEUDC(oIjk))
Next oIjk
End Function
Function DownLoadFileToTemp_Eudc(oObj As TaieProc) As Boolean
Dim hFile As Long
Dim strCommand As String
Dim bDoLoop As Boolean
Dim sBuffer As Long
Dim lNumberOfBytesRead As Long
Dim byteBuffer(102399) As Byte
Dim ReadyBuffer() As Byte
Dim F1 As Integer
Dim oIjk As Long
'add by nickc 2008/03/17
On Error GoTo CntChgErr
DownLoadFileToTemp_Eudc = False
Put2DBList "　　切換目錄：" & oObj.RemotePath
rcd oObj.RemotePath
sBuffer = 102400
Put2DBList "　　開啟檔案通道：" & oObj.ProcName
'Modified by Morgan 2018/8/3 +強制重新下載參數(INTERNET_FLAG_RELOAD=&H80000000)
hFile = FtpOpenFile(hConnection, oObj.ProcName, &H80000000, INTERNET_FLAG_TRANSFER_BINARY + &H80000000, 0)
If hFile = 0 Then
    Put2DBList "　　　開啟失敗"
    DownLoadFileToTemp_Eudc = False
    Exit Function
End If
F1 = FreeFile
'Open tempPath & oObj.ProcName For Binary As F1
Open oObj.LocalPath & oObj.ProcName & ".f" For Binary As F1
bDoLoop = True
While bDoLoop
    bDoLoop = InternetReadFileByte(hFile, VarPtr(byteBuffer(0)), sBuffer, lNumberOfBytesRead)
    If Not CBool(lNumberOfBytesRead) Then
        bDoLoop = False
    Else
        ReDim ReadyBuffer(lNumberOfBytesRead - 1) As Byte
        If lNumberOfBytesRead <> sBuffer Then
            For oIjk = 0 To (lNumberOfBytesRead - 1)
                ReadyBuffer(oIjk) = byteBuffer(oIjk)
                DoEvents
            Next oIjk
        Else
            ReadyBuffer = byteBuffer
        End If
        TaieEUDCFileSize_OK = TaieEUDCFileSize_OK + lNumberOfBytesRead
        If TaieEUDCFileSize_OK > frmAutoUpdate.ProgressBar1.Max Then
         TaieEUDCFileSize_OK = frmAutoUpdate.ProgressBar1.Max
        End If
        frmAutoUpdate.ProgressBar1.Value = TaieEUDCFileSize_OK '/ 1024
        frmAutoUpdate.lbldownState.Caption = Format(TaieEUDCFileSize_OK / 1024, "###,###,###,###") & "  Ｋ / " & Format(TaieEUDCFileSize / 1024, "###,###,###,###") & "  Ｋ"
        frmAutoUpdate.lblSpeed.Caption = Format(Trim(TaieEUDCFileSize_OK / TaieEUDCFileSize * 100), "###.00") & " ％"
        DoEvents
        Put #F1, , ReadyBuffer
    End If
Wend
Close F1
Put2DBList "　　　下載成功，關閉通道"
InternetCloseHandle hFile
Erase byteBuffer
Erase ReadyBuffer
DownLoadFileToTemp_Eudc = True
Exit Function
CntChgErr:
    If Err.Number = 75 Then
        InternetCloseHandle hFile
        IsCntChg = True
    End If
    Put2DBList "　　***" & Err.Description
End Function
Public Function CheckEudcFileSize()
Dim oIjk As Integer
Dim ShFileOP As SHFILEOPSTRUCT

For oIjk = 1 To MaxUpdateTaieEUDC
    '檔案 size 不合
    Put2DBList "　　檢查 " & UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName & "是否存在"
    If Dir(UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName & ".f") <> "" Then
        Put2DBList "　　檢查 " & UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName & "檔案大小"
        If FileLen(UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName & ".f") = UpdateTaieEUDC(oIjk).nFileSize Then
            UpdateTaieEUDC(oIjk).IsDownOk = True
            Put2DBList "　　檔案正確"
        Else
            UpdateTaieEUDC(oIjk).IsDownOk = False
            Put2DBList "　　錯誤，大小不符"
            Put2DBList "　　刪除暫存"
            '刪除
            ShFileOP.wFunc = FO_DELETE
            ShFileOP.pFrom = UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName & ".f" & Chr(0)
            ShFileOP.fFlags = FOF_NOCONFIRMATION + FOF_SILENT
            SHFileOperation ShFileOP
            If DBMode = False Then
                SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
            End If
            Put2DBList "　　重新下載??"
            If MsgBox(UpdateTaieEUDC(oIjk).ProcCName & "  下載失敗，是否重新下載此檔！", vbYesNo, "錯誤！") = vbYes Then
                TaieEUDCFileSize = UpdateTaieEUDC(oIjk).nFileSize
                TaieEUDCFileSize_OK = 0
                frmAutoUpdate.ProgressBar1.Value = 0
                Put2DBList "　　重新下載"
                DownLoadFileToTemp_Eudc UpdateTaieEUDC(oIjk)
                Put2DBList "　　檢查重新下載的大小"
                CheckEudcFileSize
                Exit For
            End If
            If DBMode = False Then
                SetWindowPos frmAutoUpdate.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
            End If
        End If
    Else
        Put2DBList "　　不存在，重新下載??"
        UpdateTaieEUDC(oIjk).IsDownOk = False
        If DBMode = False Then
            SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
        End If
        Put2DBList "　　重新下載??"
        If MsgBox(UpdateTaieEUDC(oIjk).ProcCName & "  下載失敗，是否重新下載此檔！", vbYesNo, "錯誤！") = vbYes Then
            TaieEUDCFileSize = UpdateTaieEUDC(oIjk).nFileSize
            TaieEUDCFileSize_OK = 0
            frmAutoUpdate.ProgressBar1.Value = 0
            Put2DBList "　　重新下載"
            DownLoadFileToTemp_Eudc UpdateTaieEUDC(oIjk)
            Put2DBList "　　檢查重新下載的大小"
            CheckEudcFileSize
            Exit For
        End If
        If DBMode = False Then
            SetWindowPos frmAutoUpdate.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
        End If
    End If
Next oIjk
End Function
Public Function MoveEudcFileToReady() As Boolean
Dim oIjk As Integer
Dim ShFileOP As SHFILEOPSTRUCT
Dim lpct As FILETIME, lplac As FILETIME, lplwr As FILETIME
Dim ofs As OFSTRUCT
Dim hFile As Long
Dim tZone As TIME_ZONE_INFORMATION
Dim bias As Long
Dim writedate As Date
Dim ft As SYSTEMTIME
Dim lngRCode As Long
Dim udtStartupInfo As STARTUPINFO
Dim udtProcessInfo As PROCESS_INFORMATION
Dim strCmd As String
Dim strExc As String
Dim updateFontCount As Integer
On Error GoTo 0
On Error GoTo MoveErr

MoveEudcFileToReady = False
For oIjk = 1 To MaxUpdateTaieEUDC
    '檢查是否下載成功
    If UpdateTaieEUDC(oIjk).IsDownOk = True Then
        Put2DBList "　　刪除 f2"
        updateFontCount = 1
        Do While Dir(UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName & ".f2") <> ""
            If updateFontCount >= 10 Then
                MoveEudcFileToReady = False
                IsEudcUsing = True
                Put2DBList "　　f2 刪除失敗"
                Exit Function
            End If
            Put2DBList "　　" & updateFontCount
            updateFontCount = updateFontCount + 1
            ShFileOP.wFunc = FO_DELETE
            ShFileOP.pFrom = UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName & ".f2" + Chr(0)
            ShFileOP.fFlags = FOF_NOCONFIRMATION + FOF_SILENT
            SHFileOperation ShFileOP
        Loop
        Put2DBList "　　開始更名成 f2"
        updateFontCount = 1
        Do While Dir(UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName) <> ""
            If updateFontCount >= 10 Then
                Put2DBList "　　更名 f2 失敗"
                MoveEudcFileToReady = False
                IsEudcUsing = True
                'add by nickc 2008/03/12 失敗要還原
                ShFileOP.wFunc = FO_MOVE
                ShFileOP.pFrom = Left(UpdateTaieEUDC(oIjk).LocalPath, Len(UpdateTaieEUDC(oIjk).LocalPath) - 1) & "\" & UpdateTaieEUDC(oIjk).ProcName & ".f2" + Chr(0)
                ShFileOP.pTo = UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName
                ShFileOP.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_SILENT
                SHFileOperation ShFileOP
                Exit Function
            End If
            Put2DBList "　　" & updateFontCount
            updateFontCount = updateFontCount + 1
            ShFileOP.wFunc = FO_MOVE
            ShFileOP.pFrom = UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName + Chr(0)
            ShFileOP.pTo = Left(UpdateTaieEUDC(oIjk).LocalPath, Len(UpdateTaieEUDC(oIjk).LocalPath) - 1) & "\" & UpdateTaieEUDC(oIjk).ProcName & ".f2"
            ShFileOP.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_SILENT
            SHFileOperation ShFileOP
        Loop
        Put2DBList "　　開始將 f 更新成正確名稱"
        updateFontCount = 1
        Do While Dir(UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName & ".f") <> ""
            If updateFontCount >= 10 Then
                Put2DBList "　　將 f 更新成正確名稱失敗"
                MoveEudcFileToReady = False
                IsEudcUsing = True
                'add by nickc 2008/03/12 失敗要還原
                ShFileOP.wFunc = FO_MOVE
                ShFileOP.pFrom = UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName + Chr(0)
                ShFileOP.pTo = UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName & ".f"
                ShFileOP.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_SILENT
                SHFileOperation ShFileOP
                Exit Function
            End If
            Put2DBList "　　" & updateFontCount
            updateFontCount = updateFontCount + 1
            ShFileOP.wFunc = FO_MOVE
            ShFileOP.pFrom = UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName & ".f" + Chr(0)
            ShFileOP.pTo = UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName
            ShFileOP.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_SILENT
            SHFileOperation ShFileOP
            DoEvents
        Loop
        
         'Add by Morgan 2011/3/17 刪除失敗繼續不必重試下次更新會再刪除
         Put2DBList "　　刪除 f2"
         ShFileOP.wFunc = FO_DELETE
         ShFileOP.pFrom = UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName & ".f2" + Chr(0)
         ShFileOP.fFlags = FOF_NOCONFIRMATION + FOF_SILENT
         SHFileOperation ShFileOP
         'end 2011/3/17

        
        Put2DBList "　　更改日期時間與伺服器相同"
        '更改日期
        Call GetTimeZoneInformation(tZone)
        bias = tZone.bias
        writedate = CDate(UpdateTaieEUDC(oIjk).RemoteDate.wYear & "/" & UpdateTaieEUDC(oIjk).RemoteDate.wMonth & "/" & UpdateTaieEUDC(oIjk).RemoteDate.wDay & " " & UpdateTaieEUDC(oIjk).RemoteDate.wHour & ":" & UpdateTaieEUDC(oIjk).RemoteDate.wMinute & ":" & UpdateTaieEUDC(oIjk).RemoteDate.wSecond) + TimeSerial(0, bias, 0)
        ft.wYear = Year(writedate)
        ft.wMonth = Month(writedate)
        ft.wDay = Day(writedate)
        ft.wHour = Hour(writedate)
        ft.wMinute = Minute(writedate)
        ft.wSecond = Second(writedate)
        ft.wDayOfWeek = Weekday(writedate)
        ft.wMilliseconds = UpdateTaieEUDC(oIjk).RemoteDate.wMilliseconds
        hFile = OpenFile(UpdateTaieEUDC(oIjk).LocalPath & UpdateTaieEUDC(oIjk).ProcName, ofs, OF_READWRITE)
        'NewTaieProc(oIjk).RemoteDate.wMinute = NewTaieProc(oIjk).RemoteDate.wMinute + bias
        Call SystemTimeToFileTime(ft, lplwr)
        '更動hFile的時間，第2個參數改Create DateTime
        '第3個參數改Last Access DateTime
        '第四個參數改Last Modify DateTime
        Call SetFileTime(hFile, ByVal 0, ByVal 0, lplwr) '只更改第三個(最後寫入)時間
        Call CloseHandle(hFile) '關閉檔案
        
         'Add by Morgan 2008/7/29 只有更新TTE檔時才要重開機
         If UCase(UpdateTaieEUDC(oIjk).ProcName) = "EUDC.TTE" Then
              bolReboot = True
         End If
         'end 2008/7/29
    Else
        Put2DBList "　　下載的檔案有問題"
    End If
Next oIjk
MoveEudcFileToReady = True
Exit Function
MoveErr:
    MoveEudcFileToReady = False
    IsEudcUsing = True
    Put2DBList "　　***" & Err.Description
End Function

Public Function p_GetModuleFileName() As String
'宣告變數
Dim str As String

p_GetModuleFileName = ""
str = String(MAX_PATH, "#")
GetModuleFileName App.hInstance, str, MAX_PATH
p_GetModuleFileName = Replace(str, "#", "")

End Function

Public Sub Put2DBList(m_str As String)
If frmAutoUpdate.Visible = True Then
    frmAutoUpdate.DebugList.AddItem m_str, debugListItem
    frmAutoUpdate.DebugList.Selected(debugListItem) = True
    debugListItem = debugListItem + 1
End If

'WLog m_str 'Add by Morgan 2011/2/15 除錯用
End Sub
'Add by Morgan 2010/10/19
'+秒數
Public Function fnAddSec(pDate As String, pSec As Integer) As String
   fnAddSec = Format(DateAdd("s", pSec, Format(pDate, "####/##/## ##:##:##")), "yyyymmddhhmmss")
End Function

'Add by Morgan 2010/10/7
Public Sub WLog(Optional strContent As String)
   Dim F2 As Integer
   
On Error Resume Next

   F2 = FreeFile
   Open App.Path & "\" & App.EXEName & ".log" For Append As F2
   If strContent <> "" Then
      Print #F2, Now & " " & strContent
   Else
      Print #F2, ""
   End If
   Close #F2
End Sub

'Add by Morgan 2011/5/27
'Modified by Morgan 2020/9/30 +pOwners:執行者ID
'Modified by Morgan 2020/10/21 pOwners 改為 ProCount 執行個數
Public Function CheckIsRunning(pProcessName As String, Optional ByRef ProCount As String) As Boolean
   Dim Processes, Process
   Dim stOwner As String 'Added by Morgan 2020/9/30
      
   Set Processes = GetObject("winmgmts:").ExecQuery("select * from Win32_Process where name='" & pProcessName & "'")
   If UCase(pProcessName) = UCase(App.EXEName & ".exe") Then
      If Processes.Count > 1 Then CheckIsRunning = True
   Else
      If Processes.Count > 0 Then CheckIsRunning = True
   End If
   
   'Added by Morgan 2020/9/30
   'Modified by Morgan 2020/10/21 取得執行者會有權限問題，也可能因時間差造成程式已關閉而引發 Not found 錯誤
   'pOwners = ""
   'If Processes.Count > 0 Then
   '   For Each Process In Processes
   '      If Process.getowner(stOwner) = 0 Then
   '         pOwners = pOwners & "," & stOwner
   '      End If
   '   Next
   '   If pOwners <> "" Then pOwners = Mid(pOwners, 2)
   'End If
   ProCount = Processes.Count
   'end 2020/10/21
   'end 2020/9/30
End Function

'Added by Morgan 2020/8/12
Private Function CheckFileExists(pFilePath As String) As Boolean
   Dim fs As Object
   
   On Error GoTo ErrHnd

   Set fs = CreateObject("Scripting.FileSystemObject")
   CheckFileExists = fs.FileExists(pFilePath)
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description
   Set fs = Nothing
End Function
'Added by Morgan 2023/3/24
Public Sub PUB_OpenConn()
   If adoConn.State = adStateClosed Then
      adoConn.ConnectionString = cAdoConnect
      adoConn.Open
   End If
End Sub
'Added by Morgan 2023/3/24
Public Sub PUB_CloseConn()
   If adoConn.State = adStateOpen Then
      adoConn.Close
   End If
End Sub

'Added by Morgan 2024/9/4
'取得檔案的日期時間
Public Function PUB_GetFileDateTime(pFullFilePath As String) As String
   Dim FileHandle As Long
   Dim lpReOpenBuff As OFSTRUCT
   Dim FileInfo As BY_HANDLE_FILE_INFORMATION
   Dim tZone As TIME_ZONE_INFORMATION
   Dim bias As Long
   Dim ft As SYSTEMTIME
   Dim tmpDate As Date
   
   FileHandle = OpenFile(pFullFilePath, lpReOpenBuff, OF_READ)
   GetFileInformationByHandle FileHandle, FileInfo
   CloseHandle FileHandle
   Call GetTimeZoneInformation(tZone)
   bias = tZone.bias
   FileTimeToSystemTime FileInfo.ftLastWriteTime, ft
   tmpDate = CDate(ft.wYear & "/" & ft.wMonth & "/" & ft.wDay & " " & ft.wHour & ":" & ft.wMinute & ":" & ft.wSecond) - TimeSerial(0, bias, 0)
   PUB_GetFileDateTime = Format(tmpDate, "YYYYMMDDHHmmss")
End Function
