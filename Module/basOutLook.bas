Attribute VB_Name = "basOutLook"
'右下角圖示用 Added by Morgan 2012/3/30
Public Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4

Private mlngID As Long
Private mcolNID As Collection
Private Declare Function Shell_NotifyIconA Lib "SHELL32.DLL" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
'end 2012/3/30


'Added by Morgan 2012/3/30
'右下角圖示用
Public Function AddToSystemTray(ByVal hWnd As Long, _
                                ByVal vlngCallbackMessage As Long, _
                                ByVal vipdIcon As IPictureDisp, _
                                ByVal vstrTip As String) As Long

    mlngID = mlngID + 1
   
    Dim nidTemp As NOTIFYICONDATA
   
    With nidTemp
        .cbSize = Len(nidTemp)
        .hWnd = hWnd
        .uID = mlngID
        .uFlags = NIF_MESSAGE + NIF_ICON + NIF_TIP
        .uCallbackMessage = vlngCallbackMessage
        .hIcon = CLng(vipdIcon)
        .szTip = vstrTip & vbNullChar
    End With

    If mcolNID Is Nothing Then Set mcolNID = New Collection

    mcolNID.add hWnd, CStr(mlngID)

    Shell_NotifyIconA NIM_ADD, nidTemp
   
    AddToSystemTray = mlngID

End Function

Public Sub DeleteFromSystemTray(ByVal vlngID As Long)

Dim nidTemp As NOTIFYICONDATA

With nidTemp
.cbSize = Len(nidTemp)
.hWnd = mcolNID(CStr(vlngID))
.uID = vlngID
.uFlags = NIF_MESSAGE + NIF_ICON + NIF_TIP
End With

Shell_NotifyIconA NIM_DELETE, nidTemp

End Sub
'end 2012/3/30

'Add By Sindy 2022/10/4 因為Account有引用 basFlow 會連帶需要引用到 Service1
'但因接洽單電子收文就會呼叫到一些案件系統函數, 所以才建此虛函數
Public Function PUB_AutoRecvCRLMain(strSys As String, strCRL01 As String) As Boolean
End Function

