Attribute VB_Name = "basFileList"
'Memo By Morgan 2012/12/10 智權人員欄已修改
Option Explicit


Global Const OFS_MAXPATHNAME = 128
Global Const MAX_PATH = 260
Global Const OF_READ = &H0

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

Public Type SHFILEINFO
        hIcon As Long                      '  out: icon
        iIcon As Long          '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH '  out: display name (or path)
        szTypeName As String * 80         '  out: type name
End Type
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
 Type TIME_ZONE_INFORMATION
      bias As Long
      StandardName(32) As Integer
      StandardDate As SYSTEMTIME
      StandardBias As Long
      DaylightName(32) As Integer
      DaylightDate As SYSTEMTIME
      DaylightBias As Long
 End Type
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

'Add by Morgan 2010/9/10
Public cnnConnection As ADODB.Connection

Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long

'連線
Public Function PUB_Connect2DB() As Boolean

On Error GoTo ErrHnd

   If cnnConnection Is Nothing Then
      Set cnnConnection = New ADODB.Connection
   ElseIf cnnConnection.State = adStateOpen Then
      cnnConnection.Close
   End If
   cnnConnection.ConnectionTimeout = 60
   cnnConnection.Provider = "OraOLEDB.Oracle" 'Modified by Morgan 2024/8/16 (補改)要換成這個才能支援Unicode
   cnnConnection.Properties("Data Source").Value = "M51CON"
   cnnConnection.Properties("User ID").Value = "PGMID"
   cnnConnection.Properties("Password").Value = "PGMPWD"
   cnnConnection.Open
   PUB_Connect2DB = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description
End Function

Public Function GetLocalIP() As String
   Dim IpAddrs
   IpAddrs = GetIpAddrTable
   Dim i As Integer
   For i = LBound(IpAddrs) To UBound(IpAddrs)
      If Left(IpAddrs(i), 7) = "192.168" Then
         GetLocalIP = IpAddrs(i)
         Exit For
      End If
   Next
End Function

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
