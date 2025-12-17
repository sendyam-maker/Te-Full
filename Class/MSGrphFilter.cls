VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MSGrphFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'欢迎你下载使用本代码，本份代码由程序太平洋提供下载学习之用
'声明:
'1.本站所有代码的版权归原作者所有，如果你使用了在本站下载的源代码
'  引起的一切纠纷(后果)与本站无关,请您尊重原作者的劳动成果！
'2.若本站在代码上有侵权之处请您与站长联系，站长会及时更正。
'网站：http://www.dapha.net
'论坛：http://www.5ivb.net
'Email:dapha@etang.com
'CopyRight 2001-2005 By dapha.net
'整理时间：2004-1-17 2:31:28

Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GetObjectType Lib "gdi32" (ByVal hGDIObj As Long) As Long
Private Declare Function SetViewportExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal ny As Long, lpSize As Any) As Long
Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function GetMetaFileBitsEx Lib "gdi32" (ByVal hMF As Long, ByVal nSize As Long, lpvData As Any) As Long
Private Declare Function SetMetaFileBitsEx Lib "gdi32" (ByVal nSize As Long, lpData As Any) As Long
Private Declare Function PlayMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hMF As Long) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Const MM_ANISOTROPIC = 8
Private Const OBJ_METAFILE = 9
Private Const IBMFNSIZE = 256

Public Enum MS_FILTERTYPES
    eImport
    eExport
End Enum

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_ALL_ACCESS = &HF003F
Private Const ERROR_SUCCESS = 0&
Private Const MAX_SIZE = 2048
Const sRegImportFilters As String = "Software\Microsoft\Shared Tools\Graphics Filters\Import"
Const sRegExportFilters As String = "Software\Microsoft\Shared Tools\Graphics Filters\Export"

Private Type ALDUS_FILESPEC
    flags As Long
    hFile As Long                   '/* The file handle. */
    '   fullName As String * IBMFNSIZE '/* Device, path, file names. */
    fullName(IBMFNSIZE - 1) As Byte '/* Device, path, file names. */
    filePos As Long                 '/* Our current file posn. */
    Unknown As Long
End Type

'Flags
Const fSlippery = &H1    ' { file may disappear }
Const fWrite = &H2       ' { open for write }
Const fUnnamed = &H4     ' { unnamed }
Const fLinked = &H8      ' { linked to an FS FCB }
Const fMark = &H10       ' { Generic mark bit }

Const CommonFileTypes = &H170000

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type ALDUS_PICTINFO
    hMF As Long
    rec As RECT
    inch As Long
End Type

Dim m_ImportSupported As Boolean
Dim m_ExportSupported As Boolean
Dim m_TransparentBk As Boolean
Dim m_JPEGQuality As Integer
Dim m_ShowProgress As Boolean
Dim m_ImportFilter As String
Dim m_ExportFilter As String

Dim hModule As Long, hMemPref As Long
Private mlngParameters() As Long
Private mbytCode() As Byte
Private mlngAddress As Long
Private mlngCP As Long

Public Property Let ShowProgress(ByVal bShow As Boolean)
    m_ShowProgress = bShow
End Property

Public Property Get ShowProgress() As Boolean
    ShowProgress = m_ShowProgress
End Property

Public Property Let JPEGQuality(ByVal Quality As Integer)
    m_JPEGQuality = Quality
End Property

Public Property Get JPEGQuality() As Integer
    JPEGQuality = m_JPEGQuality
End Property

Public Property Let TransparentBackground(ByVal bTransparent As Boolean)
    m_TransparentBk = bTransparent
End Property

Public Property Get TransparentBackground() As Boolean
    TransparentBackground = m_TransparentBk
End Property

Public Property Get IsImportSupported() As Boolean
    IsImportSupported = m_ImportSupported
End Property

Public Property Get IsExportSupported() As Boolean
    IsExportSupported = m_ExportSupported
End Property

Public Property Get ImportFilterStr() As String
    Dim sTemp As String
    Dim i As Long
    If m_ImportFilter = "" Then
        For i = 0 To UBound(ImportFilters)
            m_ImportFilter = m_ImportFilter & ImportFilters(i).Name & " (*." & ImportFilters(i).Extension & " )" & "|*." & ImportFilters(i).Extension & "|"
            sTemp = sTemp & "*." & ImportFilters(i).Extension & ";"
        Next i
        m_ImportFilter = Left(m_ImportFilter, Len(m_ImportFilter) - 1)
        sTemp = Left(sTemp, Len(sTemp) - 1) & ";*.tiff;*.jpeg"
        m_ImportFilter = "All graphics (" & sTemp & ")" & "|" & sTemp & "|" & m_ImportFilter
    End If
    ImportFilterStr = m_ImportFilter
End Property

Public Property Get ExportFilterStr() As String
    Dim sTemp As String
    Dim i As Long
    If m_ExportFilter = "" Then
        For i = 0 To UBound(ExportFilters)
            m_ExportFilter = m_ExportFilter & ExportFilters(i).Name & " (*." & ExportFilters(i).Extension & " )" & "|*." & ExportFilters(i).Extension & "|"
            IsBMPExportSupported = IsBMPExportSupported Or (ExportFilters(i).Extension = "bmp")
        Next i
        m_ExportFilter = Left(m_ExportFilter, Len(m_ExportFilter) - 1)
        If Not IsBMPExportSupported Then
            m_ExportFilter = "Windows bitmap (*.bmp)|*.bmp|" & m_ExportFilter
        End If
    End If
    ExportFilterStr = m_ExportFilter
End Property

Private Sub ApplySettings(ByVal sExt As String)
    Dim pMem As Long, lTemp As Long, nSize As Long
    If hMemPref = 0 Then Exit Sub
    ' *******hMemPref settings for *.gif,*.jpeg,*.png********
    ' jpg occupies 32 bytes, while *.gif and *.png occupy 36 bytes
    ' Offset   Default_Value          Description
    '   0           2               nPageMakerversion(1-4)
    '   4           2               MetaFileType(1-EMW,2-WMF)
    '                               EMF files(1) are supported by
    '                               nPageMakerVersion > 2 only
    '   6           1               ProgressDialog Template number
    '  10           1               ShowProgressDialog (Boolean)
    '..22      75(jpg), 0(gif,png)  JPEG Quality, or transparent background
    '                               for gif/png
    pMem = GlobalLock(hMemPref)
    nSize = GlobalSize(hMemPref)
    If nSize = 0 Then Exit Sub
    CopyMemory ByVal pMem, 2, 4
    If nSize >= 4 Then CopyMemory ByVal pMem + 4, 2, 1
    lTemp = Abs(m_ShowProgress)
    If pMem >= 6 Then CopyMemory ByVal pMem + 6, lTemp, 4
    If pMem >= 10 Then CopyMemory ByVal pMem + 10, lTemp, 4
    If sExt = "jpg" Then
        lTemp = m_JPEGQuality
        If lTemp Then CopyMemory ByVal pMem + 22, lTemp, 2
    ElseIf (sExt = "gif") Or (sExt = "png") Then
        lTemp = Abs(m_TransparentBk)
        If nSize >= 22 Then CopyMemory ByVal pMem + 22, CInt(lTemp), 2
    End If
    GlobalUnlock hMemPref
End Sub

Private Function GetFilterSettings(ByVal sFilterPath As String) As Boolean
    Dim lpfnGetFilterInfo As Long
    Dim lpfnGetFilterPref As Long
    Dim ret As Long
    hModule = LoadLibrary(sFilterPath)

    lpfnGetFilterInfo = GetProcAddress(hModule, "GetFilterInfo")
    ret = CallApiByPointer(lpfnGetFilterInfo, 2, 0, VarPtr(hMemPref), CommonFileTypes)
    If (ret > 5000) Then 'ret is supported nPageMaker version. Currently 9. Error values started from 5000.
        ClearUp
        Exit Function
    End If
    GetFilterSettings = True
    lpfnGetFilterPref = GetProcAddress(hModule, "GetFilterPref")
    If lpfnGetFilterPref And hMemPref Then
        CallApiByPointer lpfnGetFilterPref, 0, 0, hMemPref, CommonFileTypes
    End If
    FreeLibrary hModule
End Function

Public Function MsoLoadPicture(ByVal sFileName As String, pb As PictureBox) As Boolean
    Dim sExt As String, sFilterPath As String
    Dim ret As Long, hDummy As Long, i As Long
    Dim picWidth As Single, picHeight As Single
    Dim pMem As Long, hMem As Long
    Dim lpfnImportGr As Long
    Dim PictInfo As ALDUS_PICTINFO
    Dim FileSpec As ALDUS_FILESPEC
    Dim sPic As StdPicture
    sExt = GetExtension(sFileName)
    sFilterPath = GetFilter(sExt, True).Path
    If GetFilterSettings(sFilterPath) = False Then Exit Function

    hModule = LoadLibrary(sFilterPath)
    If hModule = 0 Then Exit Function

    lpfnImportGr = GetProcAddress(hModule, "ImportGr")
    If lpfnImportGr = 0 Then lpfnImportGr = GetProcAddress(hModule, "ImportGR")
    If lpfnImportGr = 0 Then ClearUp

    CopyMemory FileSpec.fullName(0), ByVal sFileName, IBMFNSIZE
'    If (sExt = "jpg") Or (sExt = "gif") Or (sExt = "png") Then ApplySettings sExt
    ret = CallApiByPointer(lpfnImportGr, pb.hDC, VarPtr(FileSpec), VarPtr(PictInfo), hMemPref)
    If PictInfo.hMF = 0 Then
        ClearUp
        Exit Function
    End If
    'EMS and WPG filters return hMem value instead of hWMF, ie they return memory pointer to metafile
    'bits instead of Metafile handle. Check this and convert memory block to MetaFile if nessesary
    If GetObjectType(PictInfo.hMF) <> OBJ_METAFILE Then
        hMem = PictInfo.hMF
        pMem = GlobalLock(hMem)
        PictInfo.hMF = SetMetaFileBitsEx(GlobalSize(hMem), ByVal pMem)
        PictInfo.inch = PictInfo.inch * 96 / 72
        GlobalUnlock hMem
        GlobalFree hMem
    End If
    If PictInfo.hMF = 0 Then
        ClearUp
        Exit Function
    End If
    picWidth = (PictInfo.rec.Right - PictInfo.rec.Left) / PictInfo.inch  'image width in inches
    picHeight = (PictInfo.rec.Bottom - PictInfo.rec.Top) / PictInfo.inch 'image height in inches
    'convert inches to pixels
    picWidth = pb.ScaleX(picWidth, vbInches, vbPixels)
    picHeight = pb.ScaleX(picHeight, vbInches, vbPixels)
    pb.Move 0, 0, picWidth * Screen.TwipsPerPixelX, picHeight * Screen.TwipsPerPixelY
    'EMF filter call SetViewportExtEx internally, so second call lead to error.
    SetMapMode pb.hDC, MM_ANISOTROPIC
    If sExt <> "emf" Then
        SetViewportExtEx pb.hDC, picWidth, picHeight, ByVal 0&
    End If
    Call PlayMetaFile(pb.hDC, PictInfo.hMF)
    pb.Refresh
    DeleteMetaFile PictInfo.hMF
    FreeLibrary hModule
End Function

Public Function MsoSavePicture(ByVal sFileName As String, pb As PictureBox) As Boolean
    Dim sExt As String, sFilterPath As String
    Dim ret As Long, hMetaDC As Long, hDummy As Long, i As Long
    Dim lpfnGetFilterInfo As Long
    Dim lpfnExportGr As Long
    'Actually, ExportGr function require forth parameter as ExportCallBack
    'address but VB address of function doesn't work in this case because
    'above function declared as _cdecl. Following trick allow bypass this
    'restriction - we fill stack with dummy values.
    Dim Dummy1 As Long, Dummy2 As Long, Dummy3 As Long, Dummy4 As Long
    Dim PictInfo As ALDUS_PICTINFO
    Dim FileSpec As ALDUS_FILESPEC

    ReDim abFileSpec(LenB(FileSpec) - 1)

    sExt = GetExtension(sFileName)
'    If sExt = "bmp" Then
'        SavePicture pb.Picture, sFileName
'        MsoSavePicture = True
'        Exit Function
'    End If
    hMetaDC = CreateMetaFile(App.Path & "\tmpDstInvert.wmf")
    Call BitBlt(hMetaDC, 0, 0, pb.ScaleWidth, pb.ScaleHeight, pb.hDC, 0, 0, vbDstInvert)
    PictInfo.hMF = CloseMetaFile(hMetaDC)
    DeleteMetaFile PictInfo.hMF
    
    hMetaDC = CreateMetaFile(App.Path & "\tmpMergeCopy.wmf")
    Call BitBlt(hMetaDC, 0, 0, pb.ScaleWidth, pb.ScaleHeight, pb.hDC, 0, 0, vbMergeCopy)
    PictInfo.hMF = CloseMetaFile(hMetaDC)
    DeleteMetaFile PictInfo.hMF
    
    hMetaDC = CreateMetaFile(App.Path & "\tmpMergePaint.wmf")
    Call BitBlt(hMetaDC, 0, 0, pb.ScaleWidth, pb.ScaleHeight, pb.hDC, 0, 0, vbMergePaint)
    PictInfo.hMF = CloseMetaFile(hMetaDC)
    DeleteMetaFile PictInfo.hMF
    
    hMetaDC = CreateMetaFile(App.Path & "\tmpNotSrcCopy.wmf")
    Call BitBlt(hMetaDC, 0, 0, pb.ScaleWidth, pb.ScaleHeight, pb.hDC, 0, 0, vbNotSrcCopy)
    PictInfo.hMF = CloseMetaFile(hMetaDC)
    DeleteMetaFile PictInfo.hMF
    
    hMetaDC = CreateMetaFile(App.Path & "\tmpNotSrcErase.wmf")
    Call BitBlt(hMetaDC, 0, 0, pb.ScaleWidth, pb.ScaleHeight, pb.hDC, 0, 0, vbNotSrcErase)
    PictInfo.hMF = CloseMetaFile(hMetaDC)
    DeleteMetaFile PictInfo.hMF
    
    hMetaDC = CreateMetaFile(App.Path & "\tmpPatCopy.wmf")
    Call BitBlt(hMetaDC, 0, 0, pb.ScaleWidth, pb.ScaleHeight, pb.hDC, 0, 0, vbPatCopy)
    PictInfo.hMF = CloseMetaFile(hMetaDC)
    DeleteMetaFile PictInfo.hMF
    
    hMetaDC = CreateMetaFile(App.Path & "\tmpPatInvert.wmf")
    Call BitBlt(hMetaDC, 0, 0, pb.ScaleWidth, pb.ScaleHeight, pb.hDC, 0, 0, vbPatInvert)
    PictInfo.hMF = CloseMetaFile(hMetaDC)
    DeleteMetaFile PictInfo.hMF
    
    hMetaDC = CreateMetaFile(App.Path & "\tmpPatPaint.wmf")
    Call BitBlt(hMetaDC, 0, 0, pb.ScaleWidth, pb.ScaleHeight, pb.hDC, 0, 0, vbPatPaint)
    PictInfo.hMF = CloseMetaFile(hMetaDC)
    DeleteMetaFile PictInfo.hMF
    
    hMetaDC = CreateMetaFile(App.Path & "\tmpSrcAnd.wmf")
    Call BitBlt(hMetaDC, 0, 0, pb.ScaleWidth, pb.ScaleHeight, pb.hDC, 0, 0, vbSrcAnd)
    PictInfo.hMF = CloseMetaFile(hMetaDC)
    DeleteMetaFile PictInfo.hMF
    
    hMetaDC = CreateMetaFile(App.Path & "\tmpSrcCopy.wmf")
    Call BitBlt(hMetaDC, 0, 0, pb.ScaleWidth, pb.ScaleHeight, pb.hDC, 0, 0, vbSrcCopy)
    PictInfo.hMF = CloseMetaFile(hMetaDC)
    DeleteMetaFile PictInfo.hMF
    
    hMetaDC = CreateMetaFile(App.Path & "\tmpSrcErase.wmf")
    Call BitBlt(hMetaDC, 0, 0, pb.ScaleWidth, pb.ScaleHeight, pb.hDC, 0, 0, vbSrcErase)
    PictInfo.hMF = CloseMetaFile(hMetaDC)
    DeleteMetaFile PictInfo.hMF
    
    hMetaDC = CreateMetaFile(App.Path & "\tmpSrcInvert.wmf")
    Call BitBlt(hMetaDC, 0, 0, pb.ScaleWidth, pb.ScaleHeight, pb.hDC, 0, 0, vbSrcInvert)
    PictInfo.hMF = CloseMetaFile(hMetaDC)
    DeleteMetaFile PictInfo.hMF
    
    hMetaDC = CreateMetaFile(App.Path & "\tmpSrcPaint.wmf")
    Call BitBlt(hMetaDC, 0, 0, pb.ScaleWidth, pb.ScaleHeight, pb.hDC, 0, 0, vbSrcPaint)
    PictInfo.hMF = CloseMetaFile(hMetaDC)
    DeleteMetaFile PictInfo.hMF
    
    
    
'    sFilterPath = GetFilter(sExt, True).Path
'    hModule = LoadLibrary(sFilterPath)
'    If hModule = 0 Then Exit Function
'    Call GetFilterSettings(sFilterPath)
'    '   Check out if ExportGr enabled
'    lpfnExportGr = GetProcAddress(hModule, "ExportGr")
'    If lpfnExportGr = 0 Then
'        ClearUp
'        Exit Function
'    End If
'    'Transform picture(BMP) to WMF format
'    hMetaDC = CreateMetaFile(vbNullString)
'    Call BitBlt(hMetaDC, 0, 0, 0, 0, pb.hDC, 0, 0, vbSrcCopy)
'    PictInfo.hMF = CloseMetaFile(hMetaDC)
'    'Apply some settings(Show_Progress, JPEG_Quality)
'    'If (sExt = "jpg") Or (sExt = "gif") Or (sExt = "png") Then ApplySettings sExt
'    CopyMemory FileSpec.fullName(0), ByVal sFileName, IBMFNSIZE
'    'Fill up the stack with dummy values - see note above. Now callback function
'    'will just re-fill dummies
'    Dummy1 = 1: Dummy2 = 1
'    Dummy3 = 1: Dummy4 = 1
'    ret = CallApiByPointer(lpfnExportGr, VarPtr(FileSpec), VarPtr(PictInfo), hMemPref)
    MsoSavePicture = (ret = 0)
'    DeleteMetaFile PictInfo.hMF
'    GlobalFree hMemPref
'    FreeLibrary hModule
End Function

Private Sub Class_Initialize()
    InitFilters
    If m_ExportSupported Then
        m_JPEGQuality = Asc(GetRegValue(GetFilter("jpg", False).RegPath & "\Options", "Quality"))
        m_TransparentBk = (GetRegValue(GetFilter("gif", False).RegPath & "\Options", "SetMetafileBackgroundTransparent") = "Yes")
        m_ShowProgress = (GetRegValue(GetFilter("gif", False).RegPath & "\Options", "ShowProgressDialog") = "Yes")
    End If
End Sub

Private Sub Class_Terminate()
    ClearUp
End Sub

Private Sub InitFilters()
    Dim nBound As Long
    m_ImportSupported = EnumFilters(ImportFilters, eImport)
    m_ExportSupported = EnumFilters(ExportFilters, eExport)
End Sub

Private Function GetExtension(ByVal sFileName As String) As String
    Dim sExt As String
    sExt = Right(sFileName, 4)
    If Left(sExt, 1) = "." Then sExt = Mid(sExt, 2)
    If sExt = "tiff" Then sExt = "tif"
    If sExt = "jpeg" Then sExt = "jpg"
    GetExtension = LCase(sExt)
End Function

Private Function GetFilter(ByVal sExt As String, Optional bImport As Boolean = True) As FILTERINFO
    Dim i As Long
    If bImport Then
        For i = 0 To UBound(ImportFilters)
            If ImportFilters(i).Extension = sExt Then
                GetFilter = ImportFilters(i)
                Exit For
            End If
        Next i
    Else
        For i = 0 To UBound(ExportFilters)
            'If ExportFilters(i).Extension = sExt Then
            If InStr(1, ExportFilters(i).Extension, sExt) <> 0 Then
                GetFilter = ExportFilters(i)
                Exit For
            End If
        Next i
    End If
End Function

Private Sub ClearUp()
    GlobalFree hMemPref
    FreeLibrary hModule
End Sub

Public Function CallApiByPointer(ByVal LPFN As Long, ParamArray FuncParams()) As Long
    If LPFN = 0 Then Exit Function
    Dim i As Integer
    ReDim mlngParameters(0)
    ReDim mbytCode(0)
    mlngAddress = LPFN
    ReDim mlngParameters(UBound(FuncParams) + 1)
    For i = 1 To UBound(mlngParameters)
        mlngParameters(i) = CLng(FuncParams(i - 1))
    Next i
    CallApiByPointer = CallWindowProc(PrepareCode, 0, 0, 0, 0)
End Function

Public Function GetRegValue(ByVal sKey As String, ByVal sSubKey As String) As String
    Dim hKey As Long, sTemp As String, lSize As Long, lTemp As Long, lType As Long
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKey, 0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then Exit Function
    lType = 1
    lSize = MAX_SIZE
    sTemp = String(lSize, 0)
    If RegQueryValueEx(hKey, sSubKey, 0, lType, ByVal sTemp, lSize) = ERROR_SUCCESS Then
        GetRegValue = TrimNull(sTemp)
    End If
End Function

Private Function EnumFilters(arrFI() As FILTERINFO, Optional eFilterType As MS_FILTERTYPES) As Boolean
    Dim sKey As String, sTemp As String
    Dim hKey As Long, curidx As Long, KeyName As String, KeyValue As String
    Dim i As Long
    If eFilterType = eImport Then
        sKey = sRegImportFilters
    Else
        sKey = sRegExportFilters
    End If
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKey, 0&, KEY_ALL_ACCESS, hKey) Then Exit Function
    Do
        KeyName = Space$(MAX_SIZE)
        KeyValue = Space$(MAX_SIZE)
        If RegEnumKey(hKey, curidx, KeyName, MAX_SIZE) <> ERROR_SUCCESS Then Exit Do
        ReDim Preserve arrFI(curidx)
        KeyName = TrimNull(KeyName)
        arrFI(curidx) = GetFilterInfo(sKey & "\" & KeyName)
        curidx = curidx + 1
    Loop
    RegCloseKey hKey
    EnumFilters = curidx
End Function

Private Function PrepareCode() As Long
    Dim lngX As Long, codeStart As Long
    ReDim mbytCode(18 + 32 + 6 * UBound(mlngParameters))
    codeStart = GetAlignedCodeStart(VarPtr(mbytCode(0)))
    mlngCP = codeStart - VarPtr(mbytCode(0))
    For lngX = 0 To mlngCP - 1
        mbytCode(lngX) = &HCC
    Next
    AddByteToCode &H58 'pop eax
    AddByteToCode &H59 'pop ecx
    AddByteToCode &H59 'pop ecx
    AddByteToCode &H59 'pop ecx
    AddByteToCode &H59 'pop ecx
    AddByteToCode &H50 'push eax
    For lngX = UBound(mlngParameters) To 1 Step -1
        AddByteToCode &H68 'push wwxxyyzz
        AddLongToCode mlngParameters(lngX)
    Next
    AddCallToCode mlngAddress
    AddByteToCode &HC3
    AddByteToCode &HCC
    PrepareCode = codeStart
End Function

Public Function TrimNull(startstr As String) As String
    Dim Pos As Integer
    Pos = InStr(startstr, Chr$(0))
    If Pos Then
        TrimNull = Left$(startstr, Pos - 1)
        Exit Function
    End If
    TrimNull = startstr
End Function

Private Function GetFilterInfo(ByVal sKey As String) As FILTERINFO
    Dim hKey As Long, sTemp As String, lSize As Long
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKey, 0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then Exit Function
    GetFilterInfo.RegPath = sKey
    lSize = MAX_SIZE
    sTemp = String(lSize, 0)
    If RegQueryValueEx(hKey, "Path", 0, 1, ByVal sTemp, lSize) = ERROR_SUCCESS Then
        GetFilterInfo.Path = TrimNull(sTemp)
    End If
    lSize = MAX_SIZE
    sTemp = String(lSize, 0)
    If RegQueryValueEx(hKey, "Extensions", 0, 1, ByVal sTemp, lSize) = ERROR_SUCCESS Then
        GetFilterInfo.Extension = LCase(TrimNull(sTemp))
    End If
    lSize = MAX_SIZE
    sTemp = String(lSize, 0)
    If RegQueryValueEx(hKey, "Name", 0, 1, ByVal sTemp, lSize) = ERROR_SUCCESS Then
        GetFilterInfo.Name = TrimNull(sTemp)
    End If
End Function
Private Function GetAlignedCodeStart(lngAddress As Long) As Long
    GetAlignedCodeStart = lngAddress + (15 - (lngAddress - 1) Mod 16)
    If (15 - (lngAddress - 1) Mod 16) = 0 Then GetAlignedCodeStart = GetAlignedCodeStart + 16
End Function
Private Sub AddCallToCode(lngAddress As Long)
    AddByteToCode &HE8
    AddLongToCode lngAddress - VarPtr(mbytCode(mlngCP)) - 4
End Sub

Private Sub AddLongToCode(lng As Long)
    Dim intX As Integer
    Dim byt(3) As Byte
    CopyMemory byt(0), lng, 4
    For intX = 0 To 3
        AddByteToCode byt(intX)
    Next
End Sub

Private Sub AddByteToCode(byt As Byte)
    mbytCode(mlngCP) = byt
    mlngCP = mlngCP + 1
End Sub
