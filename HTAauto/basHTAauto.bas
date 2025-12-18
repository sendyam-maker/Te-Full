Attribute VB_Name = "basHTAauto"
Option Explicit

'右下角圖示用
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

'Public Const WM_MOUSEMOVE = &H200
'Public Const WM_LBUTTONDBLCLK = &H203
'Public Const WM_LBUTTONDOWN = &H201
''Public Const WM_LBUTTONUP = &H202
'Public Const WM_MBUTTONDBLCLK = &H209
'Public Const WM_MBUTTONDOWN = &H207
'Public Const WM_MBUTTONUP = &H208
'Public Const WM_RBUTTONDBLCLK = &H206
'Public Const WM_RBUTTONDOWN = &H204
'Public Const WM_RBUTTONUP = &H205

Private mlngID As Long
Private mcolNID As Collection
Private Declare Function Shell_NotifyIconA Lib "SHELL32.DLL" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
'---------

'Public cnnConnection As New ADODB.Connection


'連線
Public Function fConnect() As Boolean
            
On Error GoTo ErrHand
   
   If cnnConnection.State = adStateClosed Then
      Forms(0).StatusBar1.Panels(1).Text = "連線中..."
      'Modified by Morgan 2022/3/23
      'cnnConnection.ConnectionString = "Provider=MSDAORA.1;Password=PGMPWD;User ID=PGMID;Data Source=m51con;Persist Security Info=True"
      cnnConnection.ConnectionTimeout = 60
      cnnConnection.Provider = IIf(strProvider <> "", strProvider, cProvider)
      cnnConnection.Properties("Data Source").Value = "m51con"
      cnnConnection.Properties("User ID").Value = UserName
      cnnConnection.Properties("Password").Value = Password
      'end 2022/3/23
      cnnConnection.Open
      Forms(0).StatusBar1.Panels(1).Text = "已連線..."
      Forms(0).Caption = Forms(0).Caption & PUB_GetDbTerminal
   End If
   fConnect = True
   
ErrHand:

   If Err.Number <> 0 Then
      Forms(0).StatusBar1.Panels(1).Text = "連線失敗..."
      WLog Err.Description 'Add By Sindy 2015/8/14
      MsgBox Err.Description
   End If

End Function

''讀資料庫電腦名稱
'Public Function PUB_GetDbTerminal() As String
'   Dim strSql As String, adoRst As New ADODB.Recordset
'On Error GoTo ErrHnd
'   strSql = "select TERMINAL FROM V$SESSION where SID=1"
'   With adoRst
'      .CursorLocation = adUseClient
'      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
'      If .RecordCount > 0 Then
'         PUB_GetDbTerminal = "(" & .Fields(0) & ")"
'      End If
'   End With
'ErrHnd:
'   If Err.Number <> 0 Then
'      MsgBox Err.Description, vbCritical
'   End If
'   Set adoRst = Nothing
'End Function

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

'Removed by Morgan 2014/3/10 整合同名函數到 basQuery
''*************************************************
''  電腦自動給號
''
''*************************************************
'Public Function AutoNo(InputItem As String, InputLength As Integer) As String
'Dim adoaccnum As New ADODB.Recordset
'Dim strItem As String, strYes As String
'
'
''911106 NICK '911106 nick 避免相同連線作做2次 transation
'Dim BolTransOk As Boolean
'BolTransOk = True
'On Error GoTo TransErr
'
'   adoTaie.BeginTrans
'   adoTaie.Execute "update autonumber set au03 = au03 where au01 = '" & InputItem & "'"
'   If Len(InputItem) > 1 Then
'      strItem = Mid(InputItem, 2, 1)
'   Else
'      strItem = InputItem
'   End If
'   adoaccnum.CursorLocation = adUseClient
'   adoaccnum.Open "select * from autonumber where au01 = '" & InputItem & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccnum.RecordCount = 0 Then
'      If InputItem = "E" Then
'         AutoNo = strItem & Mid(ACDate(strSrvDate(1)), 1, 3) & ZeroBeforeNo("2000", InputLength)
'      Else
'         If InputItem = "U" Or InputItem = "V" Then
'            AutoNo = strItem & Mid(ACDate(strSrvDate(1)), 1, 3) & ZeroBeforeNo("0", InputLength)
'         Else
'            'Modify By Sindy 2010/8/17 比對自動編號年度
'            'Modify by Morgan 2011/1/5 +K
'            If InputItem = "A" Or InputItem = "B" Or InputItem = "C" Or _
'               InputItem = "D" Or InputItem = "DP" Or InputItem = "K" Then
'               AutoNo = strItem & CompAutoNumberYear(Val(Mid(ACDate(strSrvDate(1)), 1, 3))) & ZeroBeforeNo("0", InputLength)
'            '2010/8/17 End
'            Else
'               AutoNo = strItem & Val(Mid(ACDate(strSrvDate(1)), 1, 3)) & ZeroBeforeNo("0", InputLength)
'            End If
'         End If
'      End If
'   Else
'      If adoaccnum.Fields("au02").Value <> Val(Mid(strSrvDate(1), 1, 4)) Then
'         If InputItem = "E" Then
'            AutoNo = strItem & Mid(ACDate(strSrvDate(1)), 1, 3) & ZeroBeforeNo("2000", InputLength)
'         Else
'            If InputItem = "U" Or InputItem = "V" Then
'               AutoNo = strItem & Mid(ACDate(strSrvDate(1)), 1, 3) & ZeroBeforeNo("0", InputLength)
'            Else
'               'Modify By Sindy 2010/8/17 比對自動編號年度
'               'Modify by Morgan 2011/1/5 +K
'               If InputItem = "A" Or InputItem = "B" Or InputItem = "C" Or _
'                  InputItem = "D" Or InputItem = "DP" Or InputItem = "K" Then
'                  AutoNo = strItem & CompAutoNumberYear(Val(Mid(ACDate(strSrvDate(1)), 1, 3))) & ZeroBeforeNo("0", InputLength)
'               '2010/8/17 End
'               Else
'                  AutoNo = strItem & Val(Mid(ACDate(strSrvDate(1)), 1, 3)) & ZeroBeforeNo("0", InputLength)
'               End If
'            End If
'         End If
'      Else
'         If InputItem = "U" Or InputItem = "V" Then
'            AutoNo = strItem & Mid(ACDate(strSrvDate(1)), 1, 3) & ZeroBeforeNo(str(adoaccnum.Fields("au03").Value), InputLength)
'         Else
'            'Modify By Sindy 2010/8/17 比對自動編號年度
'            'Modify by Morgan 2011/1/5 +K
'            If InputItem = "A" Or InputItem = "B" Or InputItem = "C" Or _
'               InputItem = "D" Or InputItem = "DP" Or InputItem = "K" Then
'               AutoNo = strItem & CompAutoNumberYear(Val(Mid(ACDate(strSrvDate(1)), 1, 3))) & ZeroBeforeNo(str(adoaccnum.Fields("au03").Value), InputLength)
'            '2010/8/17 End
'            Else
'               AutoNo = strItem & Val(Mid(ACDate(strSrvDate(1)), 1, 3)) & ZeroBeforeNo(str(adoaccnum.Fields("au03").Value), InputLength)
'            End If
'         End If
'      End If
'   End If
'   If Len(InputItem) = 1 Then
'      If InputItem = "U" Or InputItem = "V" Then
'         strYes = SaveAutoNo(InputItem, Mid(AutoNo, 5, InputLength))
'      Else
'         'Modify By Sindy 2010/8/17
'         'Modify by Morgan 2011/1/5 +K
'         If InputItem = "A" Or InputItem = "B" Or InputItem = "C" Or _
'            InputItem = "D" Or InputItem = "DP" Or InputItem = "K" Or _
'            Val(Mid(ACDate(strSrvDate(1)), 1, 3)) <= 99 Then
'            strYes = SaveAutoNo(InputItem, Mid(AutoNo, 4, InputLength))
'         '2010/8/17 End
'         Else
'            strYes = SaveAutoNo(InputItem, Mid(AutoNo, 5, InputLength))
'         End If
'      End If
'   End If
'   adoaccnum.Close
'   If BolTransOk Then
'        adoTaie.CommitTrans
'   End If
''911106 nick 避免相同連線作做2次 transation
'   Exit Function
'TransErr:
'   If Err.Number = -2147168237 Then
'      BolTransOk = False
'      Resume Next
'   End If
'End Function
'
''*************************************************
''  電腦給號存檔
''
''*************************************************
'Public Function SaveAutoNo(InputItem, InputNo As String) As String
'Dim adoaccnum As New ADODB.Recordset
'   adoaccnum.CursorLocation = adUseClient
'   adoaccnum.Open "select * from autonumber where au01 = '" & InputItem & "'", cnnConnection, adOpenDynamic, adLockBatchOptimistic
'   If adoaccnum.RecordCount = 0 Then
'      adoaccnum.AddNew
'      adoaccnum.Fields("au01").Value = InputItem
'   End If
'   adoaccnum.Fields("au02").Value = Mid(strSrvDate(1), 1, 4)
'   adoaccnum.Fields("au03").Value = InputNo
'   adoaccnum.UpdateBatch
'   adoaccnum.Close
'   SaveAutoNo = "Y"
'End Function
