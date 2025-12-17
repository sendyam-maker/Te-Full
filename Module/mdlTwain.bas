Attribute VB_Name = "mdlTwain"
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/4 日期欄已修改

'*******************************************************************************
'
' Description: VB Module for accessing TWAIN compatible scanner (VB 5, 6)
'
' Author:      Lumir Mik (lmik@seznam.cz)
'
' Version:     1.0
'
' License:     Free to any use. If you change some part of this code, please,
'              mention it here.
'              Receive it as my contribution to free programmer sources
'              in which I found much help and inspiration.
'
' There are 3 public functions in this module:
'
'   1. PopupSelectSourceDialog
'           shows TWAIN dialog for selecting default source for acquisition
'
'   2. TransferWithoutUI
'           transfers one image from TWAIN data source without showing
'           the data source user interface (silent transfer). The programmer
'           can set following attributes of the image:
'               - resolution (DPI)
'               - colour depth - monochromatic, grey, fullcolour
'               - image size and position on the scanner glass
'                       - left, top, right, bottom (in inches).
'           The image is saved into the BMP file.
'
'   3. TransferWithUI
'           transfers one image from TWAIN data source using the data
'           source user interface to set image attributes.
'           The image is saved into the BMP file.
'
'*******************************************************************************

Option Explicit

'-----------------------------
' Declaration for TWAIN_32.DLL
'-----------------------------
Private Declare Function DSM_Entry Lib "TWAIN_32.DLL" _
                                   (ByRef pOrigin As Any, _
                                    ByRef pDest As Any, _
                                    ByVal DG As Long, _
                                    ByVal DAT As Integer, _
                                    ByVal MSG As Integer, _
                                    ByRef pData As Any) As Integer

Private Type TW_VERSION
    MajorNum        As Integer                  ' TW_UINT16
    MinorNum        As Integer                  ' TW_UINT16
    Language        As Integer                  ' TW_UINT16
    Country         As Integer                  ' TW_UINT16
    Info(1 To 34)   As Byte                     ' TW_STR32
End Type

Private Type TW_IDENTITY
    Id                      As Long             ' TW_UINT32
    Version                 As TW_VERSION       ' TW_VERSION
    ProtocolMajor           As Integer          ' TW_UINT16
    ProtocolMinor           As Integer          ' TW_UINT16
    SupportedGroups1        As Integer          ' TW_UINT32
    SupportedGroups2        As Integer
    Manufacturer(1 To 34)   As Byte             ' TW_STR32
    ProductFamily(1 To 34)  As Byte             ' TW_STR32
    ProductName(1 To 34)    As Byte             ' TW_STR32
End Type

Private Type TW_USERINTERFACE
    ShowUI   As Integer                         ' TW_BOOL
    ModalUI  As Integer                         ' TW_BOOL
    hParent  As Long                            ' TW_HANDLE
End Type

Private Type TW_PENDINGXFERS
    Count       As Integer                      ' TW_UINT16
    Reserved1   As Integer                      ' TW_UINT32
    Reserved2   As Integer
End Type

Private Type TW_ONEVALUE
    ItemType As Integer                         ' TW_UINT16
    Item1    As Integer                         ' TW_UINT32
    Item2    As Integer
End Type

Private Type TW_CAPABILITY
    Cap          As Integer                     ' TW_UINT16
    ConType      As Integer                     ' TW_UINT16
    hContainer   As Long                        ' TW_HANDLE
End Type

Private Type TW_FIX32
    Whole   As Integer                          ' TW_INT16
    Frac    As Integer                          ' TW_UINT16
End Type

Private Type TW_FRAME
    Left     As TW_FIX32                        ' TW_FIX32
    Top      As TW_FIX32                        ' TW_FIX32
    Right    As TW_FIX32                        ' TW_FIX32
    Bottom   As TW_FIX32                        ' TW_FIX32
End Type

Private Type TW_IMAGELAYOUT
    Frame            As TW_FRAME                ' TW_FRAME
    DocumentNumber   As Long                    ' TW_UINT32
    PageNumber       As Long                    ' TW_UINT32
    FrameNumber      As Long                    ' TW_UINT32
End Type

Private Type TW_EVENT
    pEvent      As Long                         ' TW_MEMREF
    TWMessage   As Integer                      ' TW_UINT16
End Type

Private Const DG_CONTROL = 1
Private Const DG_IMAGE = 2

Private Const MSG_GET = 1
Private Const MSG_SET = 6
Private Const MSG_XFERREADY = 257
Private Const MSG_CLOSEDSREQ = 258
Private Const MSG_OPENDSM = 769
Private Const MSG_CLOSEDSM = 770
Private Const MSG_OPENDS = 1025
Private Const MSG_CLOSEDS = 1026
Private Const MSG_USERSELECT = 1027
Private Const MSG_DISABLEDS = 1281
Private Const MSG_ENABLEDS = 1282
Private Const MSG_PROCESSEVENT = 1537
Private Const MSG_ENDXFER = 1793
Private Const MSG_GETDEFAULT = 3
Private Const MSG_GETFIRST = 4
Private Const MSG_GETNEXT = 5


Private Const DAT_CAPABILITY = 1
Private Const DAT_EVENT = 2
Private Const DAT_IDENTITY = 3
Private Const DAT_PARENT = 4
Private Const DAT_PENDINGXFERS = 5
Private Const DAT_USERINTERFACE = 9
Private Const DAT_IMAGELAYOUT = 258
Private Const DAT_IMAGENATIVEXFER = 260

Private Const TWRC_SUCCESS = 0
Private Const TWRC_CHECKSTATUS = 2
Private Const TWRC_DSEVENT = 4
Private Const TWRC_NOTDSEVENT = 5
Private Const TWRC_XFERDONE = 6
Private Const TWRC_ENDOFLIST = 7

Private Const TWLG_CZECH = 45
Private Const TWLG_CHINESE_TRADITIONAL = 43


Private Const TWCY_CZECHOSLOVAKIA = 42
Private Const TWCY_TAIWAN = 886

Private Const TWON_PROTOCOLMAJOR = 1
Private Const TWON_ONEVALUE = 5
Private Const TWON_PROTOCOLMINOR = 9


'-------------------------
' Declaration for WIN32API
'-------------------------
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
                               (ByVal pDest As Long, _
                                ByVal pSource As Long, _
                                ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" _
                               (ByVal pDest As Long, _
                                ByVal Length As Long)
Private Declare Function GlobalFree Lib "kernel32.dll" _
                                    (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" _
                                    (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" _
                                      (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" _
                                     (ByVal wFlags As Long, _
                                      ByVal dwBytes As Long) As Long
Private Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" _
                                    (ByRef lpMsg As MSG, _
                                     ByVal hWnd As Long, _
                                     ByVal wMsgFilterMin As Long, _
                                     ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32.dll" _
                                          (ByRef lpMsg As MSG) As Long
Private Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" _
                                         (ByRef lpMsg As MSG) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" _
                                        (ByVal dwExStyle As Long, _
                                         ByVal lpClassName As String, _
                                         ByVal lpWindowName As String, _
                                         ByVal dwStyle As Long, _
                                         ByVal x As Long, _
                                         ByVal y As Long, _
                                         ByVal nWidth As Long, _
                                         ByVal nHeight As Long, _
                                         ByVal hWndParent As Long, _
                                         ByVal hMenu As Long, _
                                         ByVal hInstance As Long, _
                                         ByVal lpParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32.dll" _
                                       (ByVal hWnd As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Private Type BITMAPFILEHEADER
    bfType      As Integer
    bfSize      As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits   As Long
End Type

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type RGBQUAD
    rgbBlue     As Byte
    rgbGreen    As Byte
    rgbRed      As Byte
    rgbReserved As Byte
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MSG
    hWnd    As Long
    message As Long
    wParam  As Long
    lParam  As Long
    time    As Long
    pt      As POINTAPI
End Type

Private Const GHND = 66


'----------------------------
' Declaration for this Module
'----------------------------
Private m_tAppID As TW_IDENTITY
Private m_tSrcID As TW_IDENTITY
Private m_lHndMsgWin As Long

Public Enum TWAIN_MDL_COLOURTYPE
    to_BW = 0                              ' TWPT_BW
    to_GREY = 1                            ' TWPT_GRAY
    to_RGB = 2                             ' TWPT_RGB
End Enum

Private Enum TWAIN_MDL_CAPABILITY
    XFERCOUNT = 1                       ' CAP_XFERCOUNT
    PIXELTYPE = 257                     ' ICAP_PIXELTYPE
    INDICATORS = 4107                   ' CAP_INDICATORS
    UICONTROLLABLE = 4110               ' CAP_UICONTROLLABLE
    PHYSICALWIDTH = 4369                ' ICAP_PSYSICALWIDTH
    PHYSICALHEIGHT = 4370               ' ICAP_PSYSICALHEIGHT
    XRESOLUTION = 4376                  ' ICAP_XRESOLUTION
    YRESOLUTION = 4377                  ' ICAP_YRESOLUTION
    BITDEPTH = 4395                     ' ICAP_BITDEPTH
End Enum

Private Enum TWAIN_MDL_ITEMYPE
    INT16 = 1                           ' TW_INT16      short
    UINT16 = 4                          ' TW_UINT16     unsigned short
    bool = 6                            ' TW_BOOL       unsigned short
    FIX32 = 7                           ' TW_FIX32      structure
End Enum

Public Function TransferWithoutUI(ByVal sngResolution As Single, _
                                  ByVal tColourType As TWAIN_MDL_COLOURTYPE, _
                                  ByVal sngImageLeft As Single, _
                                  ByVal sngImageTop As Single, _
                                  ByVal sngImageRight As Single, _
                                  ByVal sngImageBottom As Single, _
                                  ByVal sBMPFileName As String) As Long

    '----------------------------------------------------------------------------
    ' Function transfers one image from Twain data source without showing
    '   the data source user interface (silent transfer).
    '
    ' Input values
    '   - sngResolution (Single) - resolution of the image in DPI
    '                              (dots per inch)
    '   - tColourType (UDT) - colour depth of the imaged - monochromatic (BW),
    '                         colours of grey (GREY), full colours (COLOUR)
    '   - sngImageLeft, sngImageTop, sngImageRight, sngImageBottom (Single) -
    '       values determine the rectangle on the scanner glass that will
    '       be scanned (default units are inches) - if you set Right and Bottom
    '       values to 0, the module sets maximum values the scanner driver allows
    '       (the bottom right corner of the scanner glass)
    '   - sBMPFileName (String) - the file name of the saved image
    '
    ' Function returns 0 if OK, 1 if an error occurs
    '----------------------------------------------------------------------------
                                   
    Dim lRtn As Long
    Dim lTmp As Long
    Dim blTwainOpen As Boolean
    Dim lhDib As Long
    
    On Local Error GoTo ErrPlace

    '-------------------------------
    ' Open Twain Data Source Manager
    '-------------------------------
    lRtn = OpenTwainDSM()
    If lRtn Then GoTo ErrPlace
    blTwainOpen = True

    '-----------------------
    ' Open Twain Data Source
    '-----------------------
    lRtn = OpenTwainDS()
    If lRtn Then GoTo ErrPlace

    '-----------------------------------------------------------
    ' Set all important attributes of the image and the transfer
    '-----------------------------------------------------------
    
    '----------------------------------------------------------------------
    ' Set image size and position
    ' If sngImageRight or sngImageBottom is 0 put physical width and height
    '   of the scanner into these values
    '----------------------------------------------------------------------
'edit by nickc 2005/10/21
'    If (sngImageRight = 0) Or (sngImageBottom = 0) Then
'        lRtn = TwainGetOneValue(PHYSICALWIDTH, sngImageRight)
'        If lRtn Then GoTo ErrPlace
'        lRtn = TwainGetOneValue(PHYSICALHEIGHT, sngImageBottom)
'        If lRtn Then GoTo ErrPlace
'    End If
'
'    lRtn = SetImageSize(sngImageLeft, sngImageTop, sngImageRight, sngImageBottom)
'    If lRtn Then GoTo ErrPlace
     lRtn = SetMaxImageSize()
     If lRtn Then GoTo ErrPlace
    '-----------------------------------------------
    ' Set the image resolution in DPI - both X and Y
    '-----------------------------------------------
'edit by nickc     2005/10/21
'    lRtn = TwainSetOneValue(XRESOLUTION, FIX32, sngResolution)
'    If lRtn Then GoTo ErrPlace
'
'    lRtn = TwainSetOneValue(YRESOLUTION, FIX32, sngResolution)
'    If lRtn Then GoTo ErrPlace
     lRtn = SetResolution(sngResolution)
     If lRtn Then GoTo ErrPlace
    '--------------------------
    ' Set the image colour type
    '--------------------------
    lRtn = TwainSetOneValue(PIXELTYPE, UINT16, tColourType)
    If lRtn Then GoTo ErrPlace
    
    '----------------------------------------------------------------
    ' If the colour type is fullcolour, set the bitdepth of the image
    '   - 24 bits, 32 bits, ...
    '----------------------------------------------------------------
    If tColourType = to_RGB Then lRtn = TwainSetOneValue(BITDEPTH, UINT16, 24)

    '---------------------------------------------------
    ' Set number of images you want to transfer (just 1)
    '---------------------------------------------------
    lRtn = TwainSetOneValue(XFERCOUNT, INT16, 1)
    If lRtn Then GoTo ErrPlace

    '----------------------------------------------------
    ' TRANSFER the image with UI disabled.
    '   If successful, lhDIB is filled with handle to DIB
    '----------------------------------------------------
    lRtn = TwainTransfer(False, lhDib)
    If lRtn Then GoTo ErrPlace

    '------------------
    ' Close Data Source
    '------------------
    lRtn = CloseTwainDS()
    If lRtn Then GoTo ErrPlace
    
    '--------------------------
    ' Close Data Source Manager
    '--------------------------
    lRtn = CloseTwainDSM()
    If lRtn Then GoTo ErrPlace
    blTwainOpen = False
    
    '----------------------------------
    ' Save DIB handle into the BMP file
    '----------------------------------
    lRtn = SaveDibToFile(lhDib, sBMPFileName)
    If lRtn Then GoTo ErrPlace

    TransferWithoutUI = 0
    Exit Function

ErrPlace:
    If lhDib Then lRtn = GlobalFree(lhDib)
    If blTwainOpen Then lRtn = CloseTwainDS(): lRtn = CloseTwainDSM()
    TransferWithoutUI = 1
End Function

Public Function TransferWithUI(ByVal sBMPFileName As String) As Long
        
    '-------------------------------------------------------------------
    ' Function transfers one image from Twain data source using the data
    '   source user interface to set image attributes.
    '
    ' Input values
    '   - sBMPFileName (String) - the file name of the saved image
    '
    ' Function returns 0 if OK, 1 if an error occurs
    '-------------------------------------------------------------
    
    Dim lRtn As Long
    Dim blTwainOpen As Boolean
    Dim lhDib As Long
    
    On Local Error GoTo ErrPlace

    '-------------------------------
    ' Open Twain Data Source Manager
    '-------------------------------
    lRtn = OpenTwainDSM()
    If lRtn Then GoTo ErrPlace
    blTwainOpen = True

    '-----------------------
    ' Open Twain Data Source
    '-----------------------
    lRtn = OpenTwainDS()
    If lRtn Then GoTo ErrPlace

    '----------------------------------------------------
    ' TRANSFER the image with UI enabled.
    '   If successful, lhDIB is filled with handle to DIB
    '----------------------------------------------------
    lRtn = TwainTransfer(True, lhDib)
    If lRtn Then GoTo ErrPlace

    '------------------
    ' Close Data Source
    '------------------
    lRtn = CloseTwainDS()
    If lRtn Then GoTo ErrPlace
    
    '--------------------------
    ' Close Data Source Manager
    '--------------------------
    lRtn = CloseTwainDSM()
    If lRtn Then GoTo ErrPlace
    blTwainOpen = False
    
    '----------------------------------
    ' Save DIB handle into the BMP file
    '----------------------------------
    lRtn = SaveDibToFile(lhDib, sBMPFileName)
    If lRtn Then GoTo ErrPlace

    TransferWithUI = 0
    Exit Function

ErrPlace:
    If lhDib Then lRtn = GlobalFree(lhDib)
    If blTwainOpen Then lRtn = CloseTwainDS(): lRtn = CloseTwainDSM()
    TransferWithUI = 1
End Function

Public Function PopupSelectSourceDialog() As Long
    
    '------------------------------------------------------------------
    ' Function shows the Twain dialog for selecting default data source
    '
    ' Function returns 0 if OK, 1 if an error occurs
    '------------------------------------------------------------------
    
    Dim iRtn As Integer
    Dim lRtn As Long
    
    On Local Error GoTo ErrPlace
    
    '-------------------------------
    ' Open Twain Data Source Manager
    '-------------------------------
    lRtn = OpenTwainDSM()
    If lRtn Then GoTo ErrPlace

    '----------------------------------------------------
    ' Popup "Select source" dialog
    '   DG_CONTROL, DAT_IDENTITY, MSG_USERSELECT
    '----------------------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, _
                     MSG_USERSELECT, m_tSrcID)
    If iRtn <> TWRC_SUCCESS Then
        lRtn = CloseTwainDSM()
        GoTo ErrPlace
    End If
    
    '--------------------------------
    ' Close Twain Data Source Manager
    '--------------------------------
    lRtn = CloseTwainDSM()
    If lRtn Then GoTo ErrPlace
    
    PopupSelectSourceDialog = 0
    Exit Function
    
ErrPlace:
    PopupSelectSourceDialog = 1
End Function

Private Function OpenTwainDSM() As Long
    
    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace
    
    '----------------------------------------------------
    ' Create window that will receive all TWAIN messages
    ' Message loop can be found in TwainTransfer function
    '----------------------------------------------------
    m_lHndMsgWin = CreateWindowEx(0&, "#32770", "TWAIN_MSG_WINDOW", 0&, _
                                  10&, 10&, 150&, 50&, 0&, 0&, 0&, 0&)
    If m_lHndMsgWin = 0 Then GoTo ErrPlace
    
    '------------------------------------------------------------
    ' Introduce yourself to TWAIN - MajorNum, MinorNum, Language,
    ' Country, Manufacturer, ProductFamily, ProductName, etc.
    '------------------------------------------------------------
    Call ZeroMemory(VarPtr(m_tAppID), Len(m_tAppID))
    With m_tAppID
        .Version.MajorNum = 1
        .Version.Language = TWLG_CHINESE_TRADITIONAL 'TWLG_CZECH
        .Version.Country = TWCY_TAIWAN = 886  'TWCY_CZECHOSLOVAKIA
        .ProtocolMajor = TWON_PROTOCOLMAJOR
        .ProtocolMinor = TWON_PROTOCOLMINOR
        .SupportedGroups1 = DG_CONTROL Or DG_IMAGE
    End With
    
    Call CopyMemory(VarPtr(m_tAppID.Manufacturer(1)), _
                    StrPtr(StrConv("Taie System", vbFromUnicode)), _
                    Len("Taie System"))
    Call CopyMemory(VarPtr(m_tAppID.ProductFamily(1)), _
                    StrPtr(StrConv("Taie System", vbFromUnicode)), _
                    Len("Taie System"))
    Call CopyMemory(VarPtr(m_tAppID.ProductName(1)), _
                    StrPtr(StrConv("Taie System", vbFromUnicode)), _
                    Len("Taie System"))
    
    '--------------------------------------
    ' Open Data Source Manager
    '   DG_CONTROL, DAT_PARENT, MSG_OPENDSM
    '--------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_PARENT, MSG_OPENDSM, _
                     m_lHndMsgWin)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace
    
    OpenTwainDSM = 0
    Exit Function
    
ErrPlace:
    OpenTwainDSM = 1
End Function

Private Function OpenTwainDS() As Long

    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace
    
    '----------------------------------------------------------------------
    ' Open Data Source
    '   DG_CONTROL, DAT_IDENTITY, MSG_OPENDS
    '
    ' The default data source is opened. If you want user to select the new
    '   default one, call public function PopupSelectSourceDialog.
    '----------------------------------------------------------------------
    Call ZeroMemory(VarPtr(m_tSrcID), Len(m_tSrcID))
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, MSG_OPENDS, _
                     m_tSrcID)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace

    OpenTwainDS = 0
    Exit Function
    
ErrPlace:
    OpenTwainDS = 1
End Function

Private Function CloseTwainDS() As Long
    
    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace
    
    '----------------------------------------
    ' Close Data Source
    '   DG_CONTROL, DAT_IDENTITY, MSG_CLOSEDS
    '----------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, _
                     MSG_CLOSEDS, m_tSrcID)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace
    
    CloseTwainDS = 0
    Exit Function

ErrPlace:
    CloseTwainDS = 1
End Function

Private Function CloseTwainDSM() As Long
    
    Dim lRtn As Long
    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace

    '---------------------------------------
    ' Close Data Source Manager
    '   DG_CONTROL, DAT_PARENT, MSG_CLOSEDSM
    '---------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_PARENT, MSG_CLOSEDSM, _
                     m_lHndMsgWin)
    If iRtn <> TWRC_SUCCESS Then
        lRtn = DestroyWindow(m_lHndMsgWin)
        GoTo ErrPlace
    End If

    '---------------------------
    ' Destroy the message window
    '---------------------------
    lRtn = DestroyWindow(m_lHndMsgWin)
    If lRtn = 0 Then GoTo ErrPlace
    
    CloseTwainDSM = 0
    Exit Function

ErrPlace:
    CloseTwainDSM = 1
End Function

Private Function SetImageSize(ByRef sngLeft As Single, _
                              ByRef sngTop As Single, _
                              ByRef sngRight As Single, _
                              ByRef sngBottom As Single) As Long

    Dim tImageLayout As TW_IMAGELAYOUT
    Dim lRtn As Long
    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace
    
    '-------------------------------------------------------------------
    ' Set the size of the image - in default units
    '   DG_IMAGE, DAT_IMAGELAYOUT, MSG_SET
    '
    ' If you do not select any units the INCHES are selected as default.
    ' The values of Single type are converted into TWAIN TW_FIX32.
    '-------------------------------------------------------------------
    lRtn = FloatToFix32(sngLeft, tImageLayout.Frame.Left)
    If lRtn Then GoTo ErrPlace
    
    lRtn = FloatToFix32(sngTop, tImageLayout.Frame.Top)
    If lRtn Then GoTo ErrPlace
    
    lRtn = FloatToFix32(sngRight, tImageLayout.Frame.Right)
    If lRtn Then GoTo ErrPlace
    
    lRtn = FloatToFix32(sngBottom, tImageLayout.Frame.Bottom)
    If lRtn Then GoTo ErrPlace
    
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_IMAGE, DAT_IMAGELAYOUT, MSG_SET, _
                     tImageLayout)
    If (iRtn <> TWRC_SUCCESS) And (iRtn <> TWRC_CHECKSTATUS) Then GoTo ErrPlace
    
    SetImageSize = 0
    Exit Function
    
ErrPlace:
    SetImageSize = 1
End Function

Private Function TwainTransfer(ByRef blShowUI As Boolean, _
                               ByRef lDIBHandle As Long) As Long
                               
    Dim tUI As TW_USERINTERFACE
    Dim tPending As TW_PENDINGXFERS
    Dim lhDib As Long
    Dim tEvent As TW_EVENT
    Dim tMSG As MSG
    Dim lRtn As Long
    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace
    
    '---------------------------------------------
    ' Set tUI.ShowUI to 1 (show UI) or 0 (hide UI)
    '---------------------------------------------
    With tUI
        .ShowUI = IIf(blShowUI = True, 1, 0)
        .ModalUI = 1
        .hParent = m_lHndMsgWin
    End With
    
    '----------------------------------------------
    ' Enable Data Source User Interface
    '   DG_CONTROL, DAT_USERINTERFACE, MSG_ENABLEDS
    '----------------------------------------------
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, _
                     MSG_ENABLEDS, tUI)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace
    
    '-----------------------------------------------------------------
    ' Process events in the message loop
    '   DG_CONTROL, DAT_EVENT, MSG_PROCESSEVENT
    '
    ' There are two messages we are interested in in this message loop
    '   - MSG_XFERREADY - the data source is ready to transfer
    '   - MSG_CLOSEDSREQ - the data source requests to close itself
    '-----------------------------------------------------------------
    While GetMessage(tMSG, 0&, 0&, 0&)
        Call ZeroMemory(VarPtr(tEvent), Len(tEvent))
        tEvent.pEvent = VarPtr(tMSG)
        iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_EVENT, _
                         MSG_PROCESSEVENT, tEvent)
        Select Case tEvent.TWMessage
            Case MSG_XFERREADY
                GoTo MSGGET
            Case MSG_CLOSEDSREQ
                GoTo MSGDISABLEDS
        End Select
        lRtn = TranslateMessage(tMSG)
        lRtn = DispatchMessage(tMSG)
    Wend
    
MSGGET:
    '----------------------------------------------------
    ' Start transfer
    '   DG_IMAGE, DAT_IMAGENATIVEXFER, MSG_GET
    '
    ' If transfer is successful you get the handle to DIB
    '----------------------------------------------------
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_IMAGE, DAT_IMAGENATIVEXFER, _
                     MSG_GET, lhDib)
    If iRtn <> TWRC_XFERDONE Then
        iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_PENDINGXFERS, _
                         MSG_ENDXFER, tPending)
        iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, _
                         MSG_DISABLEDS, tUI)
        GoTo ErrPlace
    End If
    
    '--------------------------------------------
    ' End transfer
    '   DG_CONTROL, DAT_PENDINGXFERS, MSG_ENDXFER
    '--------------------------------------------
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_PENDINGXFERS, _
                     MSG_ENDXFER, tPending)
    If iRtn <> TWRC_SUCCESS Then
        iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, _
                         MSG_DISABLEDS, tUI)
        GoTo ErrPlace
    End If
    
MSGDISABLEDS:
    '-----------------------------------------------
    ' Disable Data Source
    '   DG_CONTROL, DAT_USERINTERFACE, MSG_DISABLEDS
    '-----------------------------------------------
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, _
                     MSG_DISABLEDS, tUI)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace

    lDIBHandle = lhDib
    TwainTransfer = 0
    Exit Function
    
ErrPlace:
    If lhDib Then lRtn = GlobalFree(lhDib)
    lDIBHandle = 0
    TwainTransfer = 1
End Function

Private Function SaveDibToFile(ByRef lhDib As Long, _
                               ByRef sFileName As String) As Long
    
    '---------------------------------------------------------------------------
    ' Function saves the handle to DIB (device independent bitmap) into BMP file
    '---------------------------------------------------------------------------

    Dim tBFH As BITMAPFILEHEADER
    Dim tBIH As BITMAPINFOHEADER
    Dim tRGB As RGBQUAD
    Dim lpDIB As Long
    Dim lDibSize As Long
    Dim bDibits() As Byte
    Dim iFileNum As Integer
    Dim lRtn As Long
    
    On Local Error GoTo ErrPlace
    
    If sFileName = "" Then GoTo ErrPlace
    
    If Dir(sFileName, vbNormal Or vbHidden Or vbSystem) <> "" Then
        Call SetAttr(sFileName, vbNormal)
        Call Kill(sFileName)
    End If
    
    lpDIB = GlobalLock(lhDib)
    If lpDIB = 0 Then GoTo ErrPlace
    
    Call CopyMemory(VarPtr(tBIH), lpDIB, Len(tBIH))
    
    lDibSize = Len(tBIH) + (tBIH.biClrUsed * Len(tRGB)) + _
               (((tBIH.biWidth * tBIH.biBitCount + 31) \ 32) * 4 * tBIH.biHeight)
    ReDim bDibits(1 To lDibSize) As Byte
    Call CopyMemory(VarPtr(bDibits(1)), lpDIB, lDibSize)
    
    lRtn = GlobalUnlock(lhDib)
    lRtn = GlobalFree(lhDib)
    lhDib = 0
    
    With tBFH
        .bfType = 19778     ' "BM"
        .bfSize = Len(tBFH) + lDibSize
        .bfOffBits = Len(tBFH) + Len(tBIH) + (tBIH.biClrUsed * Len(tRGB))
    End With
    iFileNum = FreeFile
    Open sFileName For Binary As #iFileNum
        Put #iFileNum, , tBFH
        Put #iFileNum, , bDibits()
    Close #iFileNum
    
    SaveDibToFile = 0
    Exit Function
    
ErrPlace:
    lRtn = GlobalUnlock(lhDib)
    lRtn = GlobalFree(lhDib)
    lhDib = 0
    SaveDibToFile = 1
End Function

Private Function TwainSetOneValue(ByVal Cap As TWAIN_MDL_CAPABILITY, _
                                  ByVal ItemType As TWAIN_MDL_ITEMYPE, _
                                  ByRef Item As Variant) As Long

    '-----------------------------------------------------------------------
    ' There are four types of containers that TWAIN defines for capabilities
    ' (TW_ONEVALUE, TW_ARRAY, TW_RANGE and TW_ENUMERATION)
    ' This module deals with one of them only - TW_ONEVALUE (single value)
    ' To set some capability you have to fill TW_ONEVALUE fields and use
    '   the triplet DG_CONTROL DAT_CAPABILITY MSG_SET
    ' The macros that convert some data types are used here as well
    '-----------------------------------------------------------------------
    On Local Error GoTo ErrPlace
    
    Dim tCapability As TW_CAPABILITY
    Dim tOneValue As TW_ONEVALUE
    Dim lhOneValue As Long
    Dim lpOneValue As Long
    Dim lRtn As Long
    Dim iRtn As Integer
    Dim tFix32 As TW_FIX32
    Dim iTmp As Integer

    tCapability.ConType = TWON_ONEVALUE
    tCapability.Cap = Cap
    
    tOneValue.ItemType = ItemType
    
    Select Case ItemType
        Case INT16
            tOneValue.Item1 = CInt(Item)
        Case UINT16, bool
            If ToUnsignedShort(CLng(Item), iTmp) Then GoTo ErrPlace
            Call CopyMemory(VarPtr(tOneValue.Item1), VarPtr(iTmp), 2&)
        Case FIX32
            If FloatToFix32(CSng(Item), tFix32) Then GoTo ErrPlace
            Call CopyMemory(VarPtr(tOneValue.Item1), VarPtr(tFix32), 4&)
    End Select

    lhOneValue = GlobalAlloc(GHND, Len(tOneValue))
    lpOneValue = GlobalLock(lhOneValue)
    Call CopyMemory(lpOneValue, VarPtr(tOneValue), Len(tOneValue))
    lRtn = GlobalUnlock(lhOneValue)
    tCapability.hContainer = lhOneValue

    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_CAPABILITY, MSG_SET, _
                     tCapability)
    If iRtn <> TWRC_SUCCESS Then
        lRtn = GlobalFree(lhOneValue)
        GoTo ErrPlace
    End If
    lRtn = GlobalFree(lhOneValue)

    TwainSetOneValue = 0
    Exit Function

ErrPlace:
    TwainSetOneValue = 1
End Function

Private Function TwainGetOneValue(ByVal Cap As TWAIN_MDL_CAPABILITY, _
                                  ByRef Item As Variant) As Long

    '-----------------------------------------------------------------------
    ' There are four types of containers that TWAIN defines for capabilities
    ' (TW_ONEVALUE, TW_ARRAY, TW_RANGE and TW_ENUMERATION)
    ' This module deals with one of them only - TW_ONEVALUE (single value)
    ' To get some capability you have to fill TW_ONEVALUE fields and use
    '   the triplet DG_CONTROL DAT_CAPABILITY MSG_GET
    ' The macros that convert some data types are used here as well
    '-----------------------------------------------------------------------
    
    On Local Error GoTo ErrPlace
    
    Dim tCapability As TW_CAPABILITY
    Dim tOneValue As TW_ONEVALUE
    Dim tFix32 As TW_FIX32
    Dim lpOneValue As Long
    Dim lRtn As Long
    Dim iRtn As Integer

    tCapability.ConType = TWON_ONEVALUE
    tCapability.Cap = Cap
    
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_CAPABILITY, MSG_GET, _
                     tCapability)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace
    
    lpOneValue = GlobalLock(tCapability.hContainer)
    Call CopyMemory(VarPtr(tOneValue), lpOneValue, Len(tOneValue))
    lRtn = GlobalUnlock(tCapability.hContainer)
    lRtn = GlobalFree(tCapability.hContainer)

    Select Case tOneValue.ItemType
        Case INT16
            Item = tOneValue.Item1
        Case UINT16, bool
            Item = FromUnsignedShort(tOneValue.Item1)
        Case FIX32
            Call CopyMemory(VarPtr(tFix32), VarPtr(tOneValue.Item1), 4&)
            Item = Fix32ToFloat(tFix32)
    End Select
    
    TwainGetOneValue = 0
    Exit Function

ErrPlace:
    TwainGetOneValue = 1
End Function

Private Function ToUnsignedShort(ByRef lSrc As Long, _
                                 ByRef iDst As Integer) As Long
    
    '------------------------------------------------------------------------
    ' Sets number ranging from 0 to 65535 into 2-byte VB Integer
    ' (useful for communicating with other dll that uses unsigned data types)
    '
    ' Function returns 0 is OK, 1 if an error occurs
    '------------------------------------------------------------------------
    
    On Local Error GoTo ErrPlace
    
    If (lSrc < 0) Or (lSrc > 65535) Then GoTo ErrPlace
    
    Call CopyMemory(VarPtr(iDst), VarPtr(lSrc), 2&)
    
    ' Another way
    'iDst = IIf(lSrc > 32767, lSrc - 65536, lSrc)
    
    ToUnsignedShort = 0
    Exit Function
    
ErrPlace:
    ToUnsignedShort = 1
End Function

Private Function FromUnsignedShort(ByRef iSrc As Integer) As Long

    '------------------------------------------------------------------------
    ' Gets the 2-byte unsigned number from VB Integer data type
    ' (useful for communicating with other dll that uses unsigned data types)
    '
    ' Function returns unsigned 2-byte value (in VB Long type)
    '------------------------------------------------------------------------
    
    Dim lTmp As Long
    
    Call CopyMemory(VarPtr(lTmp), VarPtr(iSrc), 2&)
    
    ' Another way
    'lTmp = IIf(iSrc < 0, iSrc + 65536, iSrc)
    
    FromUnsignedShort = lTmp

End Function

Private Function ToUnsignedLong(ByRef sngSrc As Single, _
                                ByRef lDst As Long) As Long

    '------------------------------------------------------------------------
    ' Sets number ranging from 0 to 4294967295 into 4-byte VB Long
    ' (useful for communicating with other dll that uses unsigned data types)
    '
    ' Function returns 0 is OK, 1 if an error occurs
    '------------------------------------------------------------------------

    On Local Error GoTo ErrPlace
    
    If (sngSrc < 0) Or (sngSrc > 4294967295#) Then GoTo ErrPlace
    
    lDst = IIf(sngSrc > 2147483647, sngSrc - 4294967296#, sngSrc)
    
    ToUnsignedLong = 0
    Exit Function
    
ErrPlace:
    ToUnsignedLong = 1
End Function

Private Function FromUnsignedLong(ByRef lSrc As Long) As Single

    '------------------------------------------------------------------------
    ' Gets the 4-byte unsigned number from VB Long data type
    ' (useful for communicating with other dll that uses unsigned data types)
    '
    ' Function returns unsigned 4-byte value (in VB Single type)
    '------------------------------------------------------------------------

    Dim sngTmp As Single
    
    sngTmp = IIf(lSrc < 0, lSrc + 4294967296#, lSrc)
    FromUnsignedLong = sngTmp

End Function

Private Function Fix32ToFloat(ByRef tFix32 As TW_FIX32) As Single
        
    '----------------------------------------------------------------
    ' Converts TWAIN TW_FIX32 data structure into VB Single data type
    ' (needed for communicating with TWAIN)
    '
    ' Function returns floating-point number in VB Single data type
    '----------------------------------------------------------------
        
    Dim sngTmp As Single
    
    sngTmp = tFix32.Whole + CSng(FromUnsignedShort(tFix32.Frac) / 65536)
    Fix32ToFloat = sngTmp

End Function

Private Function FloatToFix32(ByRef sngSrc As Single, _
                              ByRef tFix32 As TW_FIX32) As Long
    
    '----------------------------------------------------------------
    ' Converts VB Single data type into TWAIN TW_FIX32 data structure
    ' (needed for communicating with TWAIN)
    '
    ' Function returns 0 is OK, 1 if an error occurs
    '----------------------------------------------------------------
    
    On Local Error GoTo ErrPlace
    
    tFix32.Whole = CInt(Fix(sngSrc))
    Call ToUnsignedShort(CLng(sngSrc * 65536) And 65535, tFix32.Frac)
    FloatToFix32 = 0
    Exit Function

ErrPlace:
    FloatToFix32 = 1
End Function

Private Function SetMaxImageSize() As Integer
    Dim tCapability As TW_CAPABILITY
    Dim tOneValueWidth As TW_ONEVALUE
    Dim tOneValueHeight As TW_ONEVALUE
    Dim lpOneValue As Long
    Dim tImageLayout As TW_IMAGELAYOUT
    Dim lRtn As Long
    
    On Local Error GoTo ErrPlace
    
    '----------------------------------------
    ' Get ICAP_PHYSICALWIDTH into TW_ONEVALUE
    '----------------------------------------
    tCapability.ConType = TWON_ONEVALUE
    tCapability.Cap = PHYSICALWIDTH
    lRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_CAPABILITY, MSG_GET, tCapability)
    If lRtn Then GoTo ErrPlace
    
    lpOneValue = GlobalLock(tCapability.hContainer)
    Call CopyMemory(VarPtr(tOneValueWidth), lpOneValue, Len(tOneValueWidth))
    lRtn = GlobalUnlock(tCapability.hContainer)
    lRtn = GlobalFree(tCapability.hContainer)
    
    '-----------------------------------------
    ' Get ICAP_PHYSICALHEIGHT into TW_ONEVALUE
    '-----------------------------------------
    tCapability.ConType = TWON_ONEVALUE
    tCapability.Cap = PHYSICALHEIGHT
    lRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_CAPABILITY, MSG_GET, tCapability)
    If lRtn Then GoTo ErrPlace
    
    lpOneValue = GlobalLock(tCapability.hContainer)
    Call CopyMemory(VarPtr(tOneValueHeight), lpOneValue, Len(tOneValueHeight))
    lRtn = GlobalUnlock(tCapability.hContainer)
    lRtn = GlobalFree(tCapability.hContainer)
        
    '----------
    ' Set frame
    '----------
    tImageLayout.Frame.Right.Whole = tOneValueWidth.Item1
    tImageLayout.Frame.Right.Frac = tOneValueWidth.Item2
    tImageLayout.Frame.Bottom.Whole = tOneValueHeight.Item1
    tImageLayout.Frame.Bottom.Frac = tOneValueHeight.Item2
    lRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_IMAGE, DAT_IMAGELAYOUT, MSG_SET, tImageLayout)
    If ((lRtn) And (lRtn <> 2)) Then GoTo ErrPlace
    
    SetMaxImageSize = 0
    Exit Function
    
ErrPlace:
    SetMaxImageSize = 1
End Function

Private Function SetResolution(ByVal iRes As Integer) As Integer
    Dim tCapability As TW_CAPABILITY
    Dim tOneValue As TW_ONEVALUE
    Dim lhOneValue As Long
    Dim lpOneValue As Long
    Dim lRtn As Long
    
    On Local Error GoTo ErrPlace
    
    tCapability.ConType = TWON_ONEVALUE
    
    '-----------------------
    ' tCapability.hContainer
    '-----------------------
    tOneValue.ItemType = FIX32
    tOneValue.Item1 = iRes
        
    lhOneValue = GlobalAlloc(GHND, Len(tOneValue))
    lpOneValue = GlobalLock(lhOneValue)
    Call CopyMemory(lpOneValue, VarPtr(tOneValue), Len(tOneValue))
    lRtn = GlobalUnlock(lhOneValue)
    tCapability.hContainer = lhOneValue
    
    '------------
    ' XResolution
    '------------
    tCapability.Cap = XRESOLUTION
    lRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_CAPABILITY, MSG_SET, tCapability)
    If lRtn Then
        lRtn = GlobalFree(lhOneValue)
        GoTo ErrPlace
    End If
    
    '------------
    ' YResolution
    '------------
    tCapability.Cap = YRESOLUTION
    lRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_CAPABILITY, MSG_SET, tCapability)
    If lRtn Then
        lRtn = GlobalFree(lhOneValue)
        GoTo ErrPlace
    End If
    
    lRtn = GlobalFree(lhOneValue)
    
    SetResolution = 0
    
    Exit Function
    
ErrPlace:
    SetResolution = 1
End Function

'add by nickc 2005/12/16 抓預設 twain 裝置
Function GetDefTwainDev() As String

GetDefTwainDev = ""
    '------------------------------------------------------------------
    ' Function shows the Twain dialog for selecting default data source
    '
    ' Function returns 0 if OK, 1 if an error occurs
    '------------------------------------------------------------------
    
    Dim iRtn As Integer
    Dim lRtn As Long
    
    On Local Error GoTo ErrPlace
    
    '-------------------------------
    ' Open Twain Data Source Manager
    '-------------------------------
    lRtn = OpenTwainDSM()
    If lRtn Then GoTo ErrPlace

    '----------------------------------------------------
    ' Popup "Select source" dialog
    '   DG_CONTROL, DAT_IDENTITY, MSG_USERSELECT
    '----------------------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, _
                     MSG_GETDEFAULT, m_tSrcID)
    If iRtn <> TWRC_SUCCESS Then
        lRtn = CloseTwainDSM()
        GoTo ErrPlace
    Else
        GetDefTwainDev = StrConv(MidB(m_tSrcID.ProductName, 1, 26), vbUnicode)
    End If
    '--------------------------------
    ' Close Twain Data Source Manager
    '--------------------------------
    lRtn = CloseTwainDSM()
    If lRtn Then GoTo ErrPlace
    
    Exit Function
    
ErrPlace:
    GetDefTwainDev = ""
End Function

'add by nickc 2005/12/16 檢查有無 twain 設備
Function CheckAnyTwainDev() As Boolean
CheckAnyTwainDev = False
    '------------------------------------------------------------------
    ' Function shows the Twain dialog for selecting default data source
    '
    ' Function returns 0 if OK, 1 if an error occurs
    '------------------------------------------------------------------
    
    Dim iRtn As Integer
    Dim lRtn As Long
    
    On Local Error GoTo ErrPlace
    
    '-------------------------------
    ' Open Twain Data Source Manager
    '-------------------------------
    lRtn = OpenTwainDSM()
    If lRtn Then GoTo ErrPlace

    '----------------------------------------------------
    ' Popup "Select source" dialog
    '   DG_CONTROL, DAT_IDENTITY, MSG_USERSELECT
    '----------------------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, _
                     MSG_GETFIRST, m_tSrcID)
    If iRtn <> TWRC_SUCCESS Then
'        If iRtn = TWRC_ENDOFLIST Then
'            CheckAnyTwainDev = False
'        End If
        lRtn = CloseTwainDSM()
        GoTo ErrPlace
    Else
        If Trim(StrConv(MidB(m_tSrcID.ProductName, 1, 26), vbUnicode)) <> "" Then
            CheckAnyTwainDev = True
        End If
    End If
    '--------------------------------
    ' Close Twain Data Source Manager
    '--------------------------------
    lRtn = CloseTwainDSM()
    If lRtn Then GoTo ErrPlace
    
    Exit Function
    
ErrPlace:
    CheckAnyTwainDev = False
End Function

'add by nickc 2005/12/16 列出所有 twain 裝置
Function EnumAllDev() As String
EnumAllDev = ""
    '------------------------------------------------------------------
    ' Function shows the Twain dialog for selecting default data source
    '
    ' Function returns 0 if OK, 1 if an error occurs
    '------------------------------------------------------------------
    
    Dim iRtn As Integer
    Dim lRtn As Long
    
    On Local Error GoTo ErrPlace
    
    '-------------------------------
    ' Open Twain Data Source Manager
    '-------------------------------
    lRtn = OpenTwainDSM()
    If lRtn Then GoTo ErrPlace

    '----------------------------------------------------
    ' Popup "Select source" dialog
    '   DG_CONTROL, DAT_IDENTITY, MSG_USERSELECT
    '----------------------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, _
                     MSG_GETFIRST, m_tSrcID)
    If iRtn <> TWRC_SUCCESS Then
'        If iRtn = TWRC_ENDOFLIST Then
'            CheckAnyTwainDev = False
'        End If
        lRtn = CloseTwainDSM()
        GoTo ErrPlace
    Else
        EnumAllDev = EnumAllDev & Replace(Trim(StrConv(MidB(m_tSrcID.ProductName, 1, 26), vbUnicode)), Chr(0), "") & vbCrLf
        Dim IsEnd As Boolean
        IsEnd = False
        Do Until IsEnd = True
            iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, _
                             MSG_GETNEXT, m_tSrcID)
            If iRtn <> TWRC_SUCCESS Then
                If iRtn = TWRC_ENDOFLIST Then
                    IsEnd = True
                Else
                    lRtn = CloseTwainDSM()
                    GoTo ErrPlace
                End If
            Else
                EnumAllDev = EnumAllDev & Replace(Trim(StrConv(MidB(m_tSrcID.ProductName, 1, 26), vbUnicode)), Chr(0), "") & vbCrLf
            End If
        Loop
    End If
    '--------------------------------
    ' Close Twain Data Source Manager
    '--------------------------------
    lRtn = CloseTwainDSM()
    If lRtn Then GoTo ErrPlace
    
    Exit Function
    
ErrPlace:
    EnumAllDev = ""
End Function

'add by nickc 2005/12/16 列出所有 twain 裝置 數量
Function GetTwainCounts() As Integer
GetTwainCounts = 0
    '------------------------------------------------------------------
    ' Function shows the Twain dialog for selecting default data source
    '
    ' Function returns 0 if OK, 1 if an error occurs
    '------------------------------------------------------------------
    
    Dim iRtn As Integer
    Dim lRtn As Long
    
    On Local Error GoTo ErrPlace
    
    '-------------------------------
    ' Open Twain Data Source Manager
    '-------------------------------
    lRtn = OpenTwainDSM()
    If lRtn Then GoTo ErrPlace

    '----------------------------------------------------
    ' Popup "Select source" dialog
    '   DG_CONTROL, DAT_IDENTITY, MSG_USERSELECT
    '----------------------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, _
                     MSG_GETFIRST, m_tSrcID)
    If iRtn <> TWRC_SUCCESS Then
'        If iRtn = TWRC_ENDOFLIST Then
'            CheckAnyTwainDev = False
'        End If
        lRtn = CloseTwainDSM()
        GoTo ErrPlace
    Else
        GetTwainCounts = GetTwainCounts + 1
        Dim IsEnd As Boolean
        IsEnd = False
        Do Until IsEnd = True
            iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, _
                             MSG_GETNEXT, m_tSrcID)
            If iRtn <> TWRC_SUCCESS Then
                If iRtn = TWRC_ENDOFLIST Then
                    IsEnd = True
                Else
                    lRtn = CloseTwainDSM()
                    GoTo ErrPlace
                End If
            Else
                GetTwainCounts = GetTwainCounts + 1
            End If
        Loop
    End If
    '--------------------------------
    ' Close Twain Data Source Manager
    '--------------------------------
    lRtn = CloseTwainDSM()
    If lRtn Then GoTo ErrPlace
    
    Exit Function
    
ErrPlace:
    GetTwainCounts = 0
End Function

'Modify by Amy 2025/02/07 intCodeType/stUPMime, stOfficeTag, stEndMime/stMS01,stMSD06
'intCodeType :0-未搜尋/1-Big5/2-UTF8
Public Function GetMime(stAttPath As String, Optional bolAtt As Boolean, Optional bolNonBig5 As Boolean, Optional bolShowErr As Boolean, _
  Optional ByRef bolChkCode As Boolean = False, Optional ByRef intCodeType As Integer, Optional ByRef stUPMime As String, Optional ByRef stOfficeTag As String, Optional ByRef stEndMime As String, _
  Optional ByVal stMS01 As String = "", Optional ByVal stMSD06 As String = "") As String
   Const cBoundaryA As String = "Boundary_A_3435FE2_6617A_AA"
   Const cDASH2 As String = "--"
   Dim sCharset As String, sCTEnc As String
   Dim strLine As String, strPrefix As String
   Dim bStart As Boolean
   Dim stMime As String
   Dim iPos As Integer
   Dim bHeadEnd As Boolean
   Dim bErr As Boolean, strErrLine As String 'Added by Morgan 2018/9/27
   Dim bInHTML As Boolean, bolOutHTML As Boolean 'Added by Morgan 2018/11/28
   Dim fso As New FileSystemObject
   Dim ts As TextStream
   'Add by Amy 2025/02/11
   Dim bolRunBig5 As Boolean, bolRunUTF8 As Boolean, bolOfficeTag As Boolean, bolUPEnd As Boolean, bolENDOpen As Boolean
   Dim stNowData As String, stOrgBefData As String, stBackData As String
   
   If bolNonBig5 = True Then
      If fso.FileExists(stAttPath) Then
         Set ts = fso.OpenTextFile(stAttPath)
         bStart = False
         Do While Not ts.AtEndOfStream
            strLine = ts.ReadLine
            
            'Added by Morgan 2013/1/21
            If bStart = True And bHeadEnd = False Then
               'Removed by Morgan 2023/6/1
'               If InStr(LTrim(UCase(strLine)), UCase("Subject: ")) = 1 Then
'                  GoTo SkipLine1
'               End If
'               If InStr(LTrim(UCase(strLine)), UCase("Date: ")) = 1 Then
'                  GoTo SkipLine1
'               End If
'               If InStr(LTrim(UCase(strLine)), UCase("From: ")) = 1 Then
'                  GoTo SkipLine1
'               End If
'               If InStr(LTrim(UCase(strLine)), UCase("To: ")) = 1 Then
'                  GoTo SkipLine1
'               End If
'               'Added by Morgan 2013/9/4
'               If InStr(LTrim(UCase(strLine)), UCase("cc: ")) = 1 Then
'                  GoTo SkipLine1
'               End If
               'end 2023/6/1
               'end 2013/9/4
               
               If InStr(LTrim(UCase(strLine)), UCase("--")) = 1 Then
                  bHeadEnd = True
               
               'Added by Morgan 2023/6/1
               ElseIf InStr(UCase(strLine), UCase("Content-Type")) > 0 Then
                  bHeadEnd = True
               Else
                  GoTo SkipLine1
               'end 2023/6/1
               End If
            End If
            'end 2013/1/21
            
            If bStart = False Then
               'Modified by Morgan 2014/1/21 Content-Type可能會在前面
               'If InStr(UCase(strLine), UCase("MIME-Version")) > 0 Then
               If InStr(UCase(strLine), UCase("MIME-Version")) > 0 Or InStr(UCase(strLine), UCase("Content-Type")) > 0 Then
                  bStart = True
                  stMime = strLine & vbCrLf
                  
                  'Added by Morgan 2023/6/1
                  If InStr(UCase(strLine), UCase("Content-Type")) > 0 Then
                     bHeadEnd = True
                  End If
                  'end 2023/6/1
               End If
            Else
               stMime = stMime & strLine & vbCrLf
            End If
SkipLine1:
         Loop
         ts.Close
      End If
   Else
   
      sCharset = "charset=" & Chr$(34) & "Big5" & Chr$(34) & vbCrLf
      sCTEnc = "Content-Transfer-Encoding: quoted-printable" & vbCrLf
   
      If bolAtt = False Then
         stMime = "MIME-Version: 1.0" & vbCrLf & _
            "Content-Type: multipart/alternative;" & vbCrLf & _
            vbTab & "boundary=" & Chr$(34) & cBoundaryA & Chr$(34) & vbCrLf & _
            "X-Mailer: Taie" & vbCrLf & vbCrLf & _
            cDASH2 & cBoundaryA & vbCrLf & "Content-Type: text/plain;" & vbCrLf & _
            vbTab & sCharset & sCTEnc & vbCrLf & _
            TextBlurb() & _
            cDASH2 & cBoundaryA & vbCrLf & "Content-Type: text/html;" & vbCrLf & _
            vbTab & sCharset & sCTEnc & vbCrLf
      End If
   
      If fso.FileExists(stAttPath) Then
         Set ts = fso.OpenTextFile(stAttPath)
         bStart = False
         Do While Not ts.AtEndOfStream
            strLine = ts.ReadLine
            
            'Add by Amy 2025/02/11 找office@taie 且為回信之Tag後,切變數 stUPMime/stOfficeTag/stEndMime
            '主旨有[專利電子報]抓eml檔Tag,將客戶回覆mail之主旨,加上客戶編號,以利將客戶信箱設定不寄電子報
            'Modify by Amy 2025/09/02 +if bolENDOpen 未抓完整 officeTag才做
            If bolENDOpen = False Then
               stNowData = strLine
               If bolChkCode = True Then
                  Call GetCutOfficeTag(1, bStart, intCodeType, stNowData, stOrgBefData, bolRunBig5, bolRunUTF8, stMS01, stMSD06, bolOfficeTag, stOfficeTag)
                  'Add by Amy 2025/07/24 專利雙週電子報 No.380 回不需電子報之內文不見,因body tag 被切成bod=(換行)y
                  If InStr(stOfficeTag, "錯誤") > 0 Then
                     strErrLine = stOfficeTag
                     bErr = True
                  End If
               End If
            End If
            'end 2025/09/02
            'end 2025/02/11
            
            'Added by Morgan 2013/1/21
            If bStart = True And bHeadEnd = False Then
               'Removed by Morgan 2023/6/1
'               If InStr(LTrim(UCase(strLine)), UCase("Subject: ")) = 1 Then
'                  GoTo SkipLine
'               End If
'               If InStr(LTrim(UCase(strLine)), UCase("Date: ")) = 1 Then
'                  GoTo SkipLine
'               End If
'               If InStr(LTrim(UCase(strLine)), UCase("From: ")) = 1 Then
'                  GoTo SkipLine
'               End If
'               If InStr(LTrim(UCase(strLine)), UCase("To: ")) = 1 Then
'                  GoTo SkipLine
'               End If
'               'Added by Morgan 2013/9/4
'               If InStr(LTrim(UCase(strLine)), UCase("cc: ")) = 1 Then
'                  GoTo SkipLine1
'               End If
               'end 2023/6/1
               'end 2013/9/4
               
               If InStr(LTrim(UCase(strLine)), UCase("--")) = 1 Then
                  bHeadEnd = True
                  
               'Added by Morgan 2023/6/1
               ElseIf InStr(UCase(strLine), UCase("Content-Type")) > 0 Then
                  bHeadEnd = True
               Else
                  GoTo SkipLine
               'end 2023/6/1
               End If
            End If
            'end 2013/1/21
            
            If bolAtt = False Then
               '圖不必寄
               If InStr(UCase(strLine), UCase("</HTML>")) > 0 Then
                  stMime = stMime & strLine & vbCrLf
                  Exit Do
               End If
   
               If bStart = False Then
                  'If InStr(UCase(strLine), UCase("MIME-Version:")) > 0 Then
                  iPos = InStr(UCase(strLine), UCase("<HTML"))
                  If iPos > 0 Then
                     bStart = True
                     '前面可能會有註解,要忽略
                     stMime = stMime & Mid(strLine, iPos) & vbCrLf
                  End If
               Else
                  'Modified by Morgan 2016/5/4 自訂style名稱的前置符號"."若再行首時會在寄送過程中被消除,故取消前面跳行使"."不是在行首
                  stMime = stMime & strLine & vbCrLf
                  'If Left(strLine, 1) = "." And (Right(stMIME, 3) = ">" & vbCrLf Or Right(stMIME, 3) = "}" & vbCrLf) Then
                  '   stMIME = Left(stMIME, Len(stMIME) - 2) & strLine & vbCrLf
                  'Else
                  '   stMIME = stMIME & strLine & vbCrLf
                  'End If
                  'end 2016/5/4
               End If
            Else
               If bStart = False Then
                  'Modified by Morgan 2014/1/21 Content-Type可能會在前面
                  'If InStr(UCase(strLine), UCase("MIME-Version")) > 0 Then
                  If InStr(UCase(strLine), UCase("MIME-Version")) > 0 Or InStr(UCase(strLine), UCase("Content-Type")) > 0 Then
                     bStart = True
                     stMime = strLine & vbCrLf
                     
                     'Added by Morgan 2023/6/1
                     If InStr(UCase(strLine), UCase("Content-Type")) > 0 Then
                        bHeadEnd = True
                     End If
                     'end 2023/6/1
                  End If
               Else
                  '修正"."開頭的資料行(寄出會被吃掉而導致內容缺.或格式異常)
                  If bInHTML = False Then
                     If InStr(UCase(strLine), UCase("<html ")) = 1 Then
                        bInHTML = True
                     End If
                  ElseIf bolOutHTML = False Then
                     If InStr(UCase(strLine), UCase("</html>")) > 0 Then
                        bolOutHTML = True
                     ElseIf strLine = ".shape {behavior:url(#default#VML);}" Then
                        strLine = " " & strLine
                     ElseIf strLine = ".MsoChpDefault" Then
                        strLine = " " & strLine
                     '字首為 "."
                     ElseIf Left(strLine, 1) = "." Then
                        'Memo by Amy 此處判斷有修改需確認 stOrgBefData 變數是否也要改
                        '前一行結尾為 "=換行"
                        If Right(stMime, 3) = ("=" & vbCrLf) Then
                           '將字首為"."拆成兩行:
                           '第1行:取前一行 "="的前一個字加上".=" 換行
                           '第2行:取此行"."後的文字
                           strLine = Left(Right(stMime, 4), 1) & ".=" & vbCrLf & Mid(strLine, 2)
                           stMime = Left(stMime, Len(stMime) - 4) & "=" & vbCrLf
                        Else
                           Debug.Print strLine
                           strErrLine = strErrLine & strLine & vbCrLf
                           bErr = True
                        End If
                     End If
                  End If
                  stMime = stMime & strLine & vbCrLf
               End If
            End If
SkipLine:
            'Add by Amy 2025/02/11 找office@taie 且為回信之Tag後,切變數stUPMime/stOfficeTag/stEndMime
            '主旨有[專利電子報]
            If bolChkCode = True Then
               'Add by Amy 2025/09/02 +if Office Tag 已完整抓到後,都抓 strLine
               '  避免字首為 "." 寄出會被吃掉而導致格式異常 ex:雙週電子報 no.382
               If stOfficeTag <> "" And bolENDOpen = True Then
                  If Mid(Replace(strLine, vbCrLf, "</BR>"), 2, 7) = ".=</BR>" Then
                     stEndMime = Left(stEndMime, Len(stEndMime) - 4) & "=" & vbCrLf
                  End If
                  stNowData = strLine
               End If
               'end 2025/09/02
               
               '找到回信Tag
               If bolOfficeTag = True Then
                  Call GetCutOfficeTag(2, bStart, intCodeType, stNowData, stOrgBefData, bolRunBig5, bolRunUTF8, stMS01, stMSD06, bolOfficeTag, stOfficeTag, bolUPEnd, stUPMime, bolENDOpen, stEndMime, stMime)
                  'Add by Amy 2025/07/24 專利雙週電子報 No.380 回不需電子報之內文不見,因body tag 被切成bod=(換行)y
                  If InStr(stOfficeTag, "錯誤") > 0 Then
                     strErrLine = stOfficeTag
                     bErr = True
                  End If
               '找到<body Tag
               ElseIf bolRunBig5 = True Then
                  'Modify by Amy 2025/09/02 未抓bolOfficeTag 前,都抓 strLine
                  '  避免字首為 "." 寄出會被吃掉而導致格式異常 ex:雙週電子報 no.382
                  'stOrgBefData = stNowData
                  stOrgBefData = strLine
               End If
            End If
            'end 2025/02/11
         Loop
         ts.Close
      End If
   
      If bolAtt = False Then
         stMime = stMime & cDASH2 & cBoundaryA & cDASH2 & vbCrLf
      End If
   End If
   GetMime = stMime
   If bErr And bolShowErr Then MsgBox "MIME內容有異常，請修正下列內容：" & vbCrLf & vbCrLf & strErrLine, vbExclamation
End Function

'Added by Morgan 2024/8/16
'Unicode文字轉UTF-8Base64編碼
Public Sub PUB_ConvUni2UTF8Base64(ByRef pUniText As String)
   If pUniText <> StrConv(StrConv(pUniText, vbFromUnicode), vbUnicode) Then
      pUniText = "=?utf-8?B?" & ConvertToBase64(pUniText, False, False, True) & "?="
   End If
End Sub

'Add by Amy 2025/01/23 取得Office@taie.com.tw 回信mail Tag
Public Function bolGetRep_Big5(stMargeData As String, ByRef stOrgBefData As String, ByRef stNowData As String, ByRef bolOfficeTag As Boolean, ByRef stReturnData As String, _
  Optional ByRef stMS01 As String = "", Optional ByRef stMSD06 As String = "") As Boolean
   Dim intATagS As Integer, stFixEnterWod As String, stOfficeTagData As String, stSign As String, stTmp(2) As String, stRepTxt(2) As String
   Dim i As Integer, arrTemp
   Dim j As Integer, stChkW(2) As String 'Add by Amy 2025/07/24
   Dim IsNoOk As Boolean 'Add by Amy 2025/09/08
       
   bolGetRep_Big5 = False
   stReturnData = ""
   
   stTmp(0) = stMargeData
   
   '找到 office Tag
   If bolOfficeTag = False And InStr(UCase(stTmp(0)), UCase("<a href")) > 0 And InStr(UCase(stTmp(0)), UCase("mailto:office@")) > 0 Then
      bolOfficeTag = True
      intATagS = InStr(UCase(stTmp(0)), UCase("<a href"))
   End If
   If bolOfficeTag = False Then Exit Function
   
   stRepTxt(1) = ""
   '第1次抓到 (intATagS 才有值)且<a href Tag 未結束
   If intATagS > 0 And InStr(UCase(stTmp(0)), UCase("_blank"">")) = 0 Then
      'Modify by Amy 2025/09/02 +if ,抓到的mailto:office@ 跨2行且前一行有=符號
      '     Ex:No.382 寄出信的"若您不需要此電子報，請按此。"字會變超大
      'Modify by Amy 2025/09/08 No383 加入的(MSD01：),"："編碼後會使主旨部份字變亂碼且前一行長度為76
      'If Len(stOrgBefData) < intATagS And Len(stOrgBefData) <= 75 Then
      If Len(stOrgBefData) < intATagS And Right(stOrgBefData, 1) = "=" Then
         '前行未寫入ts.Line,此行抓<span lang=3D"EN-US"><a href=3D"mailto:office@taie.co
         stOrgBefData = stOrgBefData & ";☆☆☆"
         stTmp(1) = InStr(UCase(stNowData), "<A HREF")
         If Val(stTmp(1)) >= 2 Then
            stTmp(2) = Mid(stNowData, 1, Val(stTmp(1)) - 1)
            '將<a href 前Tag 寫入stOrgBefData 並+換檔符號
            stOrgBefData = stOrgBefData & stTmp(2) & ";***"
         End If
         stNowData = Mid(stNowData, Val(stTmp(1)))
      Else
         stOrgBefData = Mid(stOrgBefData, 1, intATagS - 1)
         stNowData = Replace(stTmp(0), stOrgBefData, "")
         stOrgBefData = stOrgBefData & ";***" '<a href 前Tag +換檔符號
      End If
      'end 2025/09/02
      'Add by Amy 2025/09/08 No383 主旨加入(MSD01：)會變亂碼,改以原碼解析
      If Right(stNowData, 1) = "=" Then
         stNowData = Mid(stNowData, 1, Len(stNowData) - 1) & "☆☆"
      End If
   '已找到 <a href Tag 結束 Tag
   ElseIf InStr(UCase(stTmp(0)), UCase("_blank"">")) > 0 Then
      If Right(stNowData, 1) = "=" Then stSign = "="
      stRepTxt(1) = Mid(stTmp(0), 1, Val(InStr(UCase(stTmp(0)), UCase("_blank"">"))) + 7)
      stRepTxt(0) = Replace(stTmp(0), stRepTxt(1), "") & stSign '剩下的字串
      stRepTxt(1) = stOrgBefData & Replace(stNowData, stRepTxt(0), "") '抓原始資料+結束Tag
      stNowData = stRepTxt(0) & ";***"
   End If
   If stRepTxt(1) <> "" Then
      'Modify by Amy 2025/09/08 No.383 主旨加入(MSD01：),解析再寄出編碼後全型冒號後會變亂碼
      stFixEnterWod = stRepTxt(1)
      
      stRepTxt(0) = "": stRepTxt(1) = "": stRepTxt(2) = ""
      stTmp(1) = stMS01
      If stTmp(1) = "" Then stTmp(1) = "空"
      stTmp(2) = stMSD06
      If stTmp(2) = "" Then stTmp(2) = "空"
'*** Memo by Amy 2025/09/09 此處有修改 frmEDM.BatchMail 取代字也要修改 ***
      stTmp(2) = "%20(SEQ%20" & stTmp(1) & "%20/%20" & stTmp(2) & ")"
'*** End Memo by Amy 2025/09/09 此處有修改 frmEDM.BatchMail 取代字也要修改 ***
      
      stTmp(0) = InStr(stFixEnterWod, "&amp;")
      '找到&amp;字串
      If Val(stTmp(0)) > 0 Then
         stRepTxt(0) = Mid(stFixEnterWod, 1, Val(stTmp(0)) + 4)
         stRepTxt(2) = Replace(stFixEnterWod, stRepTxt(0), "")
         '避免一行超過75個字 (可接受有空白76個字,故保險以75字判斷),故加換行串
         stRepTxt(0) = Replace(stRepTxt(0), "&amp;", vbCrLf & stTmp(2) & "&amp;")
         stOfficeTagData = stRepTxt(0) & vbCrLf & stRepTxt(2) '加入要加的字串組回
         stOfficeTagData = Replace(stOfficeTagData, "☆☆", "=" & vbCrLf) '☆☆換回=用
      '&amp;字串被切
      Else
         stReturnData = "錯誤：不需電子報內容有誤-找不到〔&amp;〕"
         Exit Function
         'Memo by Amy 2025/09/08 目前&amp;字串被切尚未遇到,遇到再測
'         arrTemp = Split(stFixEnterWod, "☆") '2☆保留一個☆換回=用
'         For i = LBound(arrTemp) To UBound(arrTemp)
'            stRepTxt(1) = arrTemp(i) '目前字串
'            stTmp(0) = InStr(Replace(stRepTxt(0) & stRepTxt(1), "☆", ""), "&amp;")
'            If Val(stTmp(0)) > 0 And IsNoOk = False Then
'               IsNoOk = True
'               If Len(stRepTxt(0)) > Val(stTmp(0)) And Right(stRepTxt(0), 1) = "☆" Then
'                  stOfficeTagData = Left(stOfficeTagData, Len(stRepTxt(0)) - Val(stTmp(0))) & "☆"
'                  stRepTxt(2) = Mid(stRepTxt(0), Val(stTmp(0)), Len(stRepTxt(0)) - 1) & stRepTxt(1)
'                  stRepTxt(2) = Mid(stRepTxt(2), 1, InStr(stRepTxt(2), "&amp;") + 5)
'                  stRepTxt(2) = Left(stRepTxt(2), Len(stRepTxt(2)) - 6) & "&amp;" & vbCrLf
'                  stRepTxt(1) = Mid(stRepTxt(1), InStr(stRepTxt(1), "&amp;") + 6)
'                  stOfficeTagData = stOfficeTagData & stRepTxt(2) & stRepTxt(1)
'               End If
'            Else
'               stOfficeTagData = stOfficeTagData & stRepTxt(0)
'            End If
'         Next i
'         stOfficeTagData = Replace(stOfficeTagData, "☆", "=") '☆換回=用
      End If
      stReturnData = stOfficeTagData
      bolGetRep_Big5 = True
   End If
End Function

'Add by Amy 2025/02/11 抓到<a href=3D"mailto:office@taie.com.tw?subject=..." target=3D"_blank"> 切變數stUPMime/stOfficeTag/stEndMime
Public Sub GetCutOfficeTag(ByVal intState As Integer, ByVal bStart As Boolean, ByRef intCodeType As Integer, ByRef stNowData As String, ByRef stOrgBefData As String, ByRef bolRunBig5 As Boolean, ByRef bolRunUTF8 As Boolean, Optional ByRef stMS01 As String, Optional ByRef stMSD06 As String, _
   Optional ByRef bolOfficeTag As Boolean, Optional ByRef stOfficeTag As String, Optional ByRef bolUPEnd As Boolean, Optional ByRef stUPMime As String, Optional ByRef bolENDOpen As Boolean, Optional ByRef stEndMime As String, Optional ByVal stMime As String)
Dim stBackData As String, stTmp(1) As String

   '為big5編碼且抓到<HTML Tag
   If intCodeType = 1 And bStart = True Then
      stTmp(0) = ""
      If stOrgBefData <> "" Then
         If Right(stOrgBefData, 1) = "=" Then
            stTmp(0) = stTmp(0) & Mid(stOrgBefData, 1, Len(stOrgBefData) - 1)
         Else
            stTmp(0) = stTmp(0) & stOrgBefData
         End If
         stTmp(0) = Replace(stTmp(0), "☆☆", "") 'Add by Amy 2025/09/08
      End If
      
      If stNowData <> "" Then
         If Right(stNowData, 1) = "=" Then
            stTmp(0) = stTmp(0) & Left(stNowData, Len(stNowData) - 1)
         Else
            stTmp(0) = stTmp(0) & stNowData
         End If
         stTmp(0) = Replace(stTmp(0), "☆☆", "") 'Add by Amy 2025/09/08
      End If
      If bolENDOpen = False And intState = 1 Then
         If bolRunBig5 = True Then
            Call bolGetRep_Big5(stTmp(0), stOrgBefData, stNowData, bolOfficeTag, stBackData, stMS01, stMSD06)
            'Modify by Amy 2025/07/24 +if InStr(stBackData, "錯誤")
            If InStr(stBackData, "錯誤") > 0 Then
               stOfficeTag = stBackData
               Exit Sub
            ElseIf stBackData <> "" Then
               stOfficeTag = stBackData
            End If
         ElseIf bolRunUTF8 = True Then
            'Memo 2025/01/13 utf8 <body 編碼後 code 可能不同無法觸析,故先不做
         End If
      End If
   End If
   
   If intState = 1 Then
'*** 設定 intCodeType / bolRunBig5 ***
      If intCodeType = 0 Then
         If InStr(UCase(stNowData), UCase("?big5?")) > 0 Then
            intCodeType = 1
         '信另存eml為utf8 編碼<body  編碼後的code 可能不同無法觸析,故先不做
         'ex:113104 No.359 (PGJvZHk-<body 編碼) / 1131017 No.360 (DQo8Ym9keSB-<body 編碼)
         ElseIf InStr(UCase(stNowData), UCase("Subject: =?utf-8?")) > 0 Then
            intCodeType = 2
         End If
      End If
      
      If bolRunBig5 = False And bolRunUTF8 = False Then
         'Big5編碼有<body 才開始判斷
         If intCodeType = 1 And InStr(UCase(stNowData), UCase("<body ")) > 0 Then
            bolRunBig5 = True
         'UTF8編碼有<html 才開始判斷 (2025/01/13 目前不會Run)
         ElseIf intCodeType = 2 And InStr(UCase(stNowData), UCase("Content-Type: text/html;")) > 0 Then
            bolRunUTF8 = True
         End If
      End If
   ElseIf intState = 2 Then
'*** 切變數stUPMime/stOfficeTag/stEndMime (找office@taie 且為回信之Tag後) ***
      'Office Tag 已完整抓到後
      If bolENDOpen = True Then
         stEndMime = stEndMime & stNowData & vbCrLf
      '已抓到Office Tag開始寫stEndMime
      ElseIf bolOfficeTag = True Then
         'stUPMime[已]結束
         If bolUPEnd = True Then
            If Right(stNowData, 4) = ";***" Then
               bolENDOpen = True
               stTmp(1) = Mid(stNowData, 1, Len(stNowData) - 4)
               stEndMime = stEndMime & stTmp(1) & vbCrLf
               stNowData = ""
            End If
         'stUPMime[未]結束
         ElseIf bolUPEnd = False Then
            'A Tag 已抓到,改寫至stEndMime
            If Right(stOrgBefData, 4) = ";***" Then
               bolUPEnd = True
               'Modify by Amy 2025/09/02 +if  Ex:No.382 寄出信的"若您不需要此電子報，請按此。"字會變超大
               If InStr(stOrgBefData, ";☆☆☆") > 0 Then
                  '抓到的<a href Tag為前一行資料(有長度換行=符號)+第2行部份資料
                  stOrgBefData = Replace(stOrgBefData, ";☆☆☆", vbCrLf)
               End If
               stTmp(1) = Mid(stOrgBefData, 1, Len(stOrgBefData) - 4)
               stUPMime = Mid(stMime, 1, Val(InStrRev(stMime, stTmp(1))) + (Len(stTmp(1)) - 1)) & vbCrLf
               
               stOrgBefData = ""
            End If
            'Office Tag 已抓到完整字串
            If Right(stNowData, 4) = ";***" Then
               bolENDOpen = True
               stEndMime = stEndMime & Mid(stNowData, 1, Len(stNowData) - 4) & vbCrLf
            End If
         End If
         If bolENDOpen = False Then
            If stOrgBefData <> "" Then
               'Add by Amy 2025/09/08
               If Right(stNowData, 1) = "=" Then
                  stNowData = Mid(stNowData, 1, Len(stNowData) - 1) & "☆☆"
               End If
               'end 2025/09/08
               stOrgBefData = stOrgBefData & stNowData '找到<a href 未找到結束tag
            Else
               stOrgBefData = stNowData
            End If
         End If
      End If
      
   End If
End Sub
