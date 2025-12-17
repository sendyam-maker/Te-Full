VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCopyCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

' 宣告欄位內容結構
Private Type DBFIELDITEM
   fiName As String
   fiData As String
   fiType As Integer
End Type

Dim m_FieldList() As DBFIELDITEM
Dim m_FieldCount As Integer

' 刪除所有欄位
Private Function ClearField()
   If m_FieldCount > 0 Then
      Erase m_FieldList
   End If
   m_FieldCount = 0
End Function

' 尋找欄位
Private Function FindField(ByVal strFieldName As String) As Integer
   Dim nIndex As Integer
   FindField = -1
   For nIndex = 0 To m_FieldCount - 1
      If m_FieldList(nIndex).fiName = strFieldName Then
         FindField = nIndex
         Exit For
      End If
   Next nIndex
End Function

' 檢查欄位是否已存在
Private Function IsFieldExist(ByVal strFieldName) As Boolean
   Dim nIndex As Integer
   
   IsFieldExist = False
   If FindField(strFieldName) >= 0 Then
      IsFieldExist = True
   End If
End Function

' 新增一個欄位
Private Function AddField(ByVal strFieldName As String, ByVal nType As Integer) As Boolean
   AddField = False
   If IsFieldExist(strFieldName) = False Then
      m_FieldCount = m_FieldCount + 1
      ReDim Preserve m_FieldList(m_FieldCount)
      m_FieldList(m_FieldCount - 1).fiName = strFieldName
      m_FieldList(m_FieldCount - 1).fiType = nType
      m_FieldList(m_FieldCount - 1).fiData = Empty
      AddField = True
   End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 設定欄位的內容
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetFieldData(ByVal strFieldName As String, ByVal strData As String) As Boolean
   Dim nIndex As Integer
   SetFieldData = False
   nIndex = FindField(strFieldName)
   If nIndex >= 0 Then
      m_FieldList(nIndex).fiData = strData
   End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 初始化
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OnInitial()
   ' 先清除所有欄位
   ClearField
   ' 設定欄位
   AddField "CP01", 0: AddField "CP02", 0: AddField "CP03", 0: AddField "CP04", 0: AddField "CP05", 1
   AddField "CP06", 1: AddField "CP07", 1: AddField "CP08", 0: AddField "CP09", 0: AddField "CP10", 0
   AddField "CP11", 0: AddField "CP12", 0: AddField "CP13", 0: AddField "CP14", 0: AddField "CP15", 1
   AddField "CP16", 1: AddField "CP17", 1: AddField "CP18", 1: AddField "CP19", 1: AddField "CP20", 0
   AddField "CP21", 0: AddField "CP22", 0: AddField "CP23", 0: AddField "CP24", 0: AddField "CP25", 1
   AddField "CP26", 0: AddField "CP27", 1: AddField "CP28", 0: AddField "CP29", 0: AddField "CP30", 0
   AddField "CP31", 0: AddField "CP32", 0: AddField "CP33", 1: AddField "CP34", 1: AddField "CP35", 0
   AddField "CP36", 0: AddField "CP37", 0: AddField "CP38", 0: AddField "CP39", 0: AddField "CP40", 0
   AddField "CP41", 0: AddField "CP42", 0: AddField "CP43", 0: AddField "CP44", 0: AddField "CP45", 0
   AddField "CP46", 1: AddField "CP47", 1: AddField "CP48", 1: AddField "CP49", 0: AddField "CP50", 0
   AddField "CP51", 0: AddField "CP52", 0: AddField "CP53", 1: AddField "CP54", 1: AddField "CP55", 0
   AddField "CP56", 0: AddField "CP57", 1: AddField "CP58", 0: AddField "CP59", 0: AddField "CP60", 0
   AddField "CP61", 0: AddField "CP62", 0: AddField "CP63", 0: AddField "CP64", 0
   AddField "CP71", 0: AddField "CP72", 0: AddField "CP73", 1: AddField "CP74", 1: AddField "CP75", 1
   AddField "CP76", 1: AddField "CP77", 1: AddField "CP78", 1: AddField "CP79", 1
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 結束
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OnTerminate()
   ClearField
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 寫入案件進度檔中
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function WriteCaseProgress()
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim strTemp As String
   
   bFirst = True
   strSql = "INSERT INTO CaseProgress ("
   For nIndex = 0 To m_FieldCount - 1
      strTemp = m_FieldList(nIndex).fiName
      If strTemp <> Empty Then
         If bFirst = True Then
            strSql = strSql & strTemp
            bFirst = False
         Else
            strSql = strSql & "," & strTemp
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To m_FieldCount - 1
      strTemp = Empty
      If m_FieldList(nIndex).fiType = 0 Then
         strTemp = "'" & ChgSQL(m_FieldList(nIndex).fiData) & "'"
      Else
         If IsEmptyText(m_FieldList(nIndex).fiData) = True Then
            strTemp = "NULL"
         Else
            strTemp = m_FieldList(nIndex).fiData
         End If
      End If
      If bFirst = True Then
         strSql = strSql & strTemp
         bFirst = False
      Else
         strSql = strSql & "," & strTemp
      End If
   Next nIndex
   strSql = strSql & ")"
   ' 執行SQL指令
   cnnConnection.Execute strSql
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 拷貝母案到子案
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function CopyCaseProgress(ByVal strCP09 As String) As Boolean
'Modify By Sindy 2012/7/4 +strCP10,strCP06
Public Function CopyCaseProgress(ByVal strCP09 As String, Optional TFCountry As String = "", Optional strCP10 As String = "", Optional strCP06 As String = "") As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim fldList() As DBFIELDITEM
   Dim fldCount As Integer
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTemp As String
   Dim nIndex As Integer
   Dim strFieldName As String
   
   CopyCaseProgress = False
   
   ' 先讀取案件進度檔
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & strCP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查此筆記錄是否存在
   If rsTmp.RecordCount <= 0 Then
      rsTmp.Close
      GoTo EXITSUB
   End If
   ' 子案不予處理
   ' 檢查是否為母案
   If rsTmp.Fields("CP03") <> "0" Then
      rsTmp.Close
      GoTo EXITSUB
   End If
   If rsTmp.Fields("CP04") <> "00" Then
      rsTmp.Close
      GoTo EXITSUB
   End If
   ' 設定帳號前兩欄
   If IsNull(rsTmp.Fields("CP01")) = False Then
      strTM01 = rsTmp.Fields("CP01")
   End If
   If IsNull(rsTmp.Fields("CP02")) = False Then
      strTM02 = rsTmp.Fields("CP02")
   End If
   
   ' 初始化欄位串列
   OnInitial
   ' 讀取欄位內容
   For nIndex = 1 To 79
      strFieldName = "CP" & Format(nIndex, "00")
      If IsNull(rsTmp.Fields(strFieldName)) = True Then
         SetFieldData strFieldName, Empty
      Else
         ' 設定欄位內的資料
         SetFieldData strFieldName, rsTmp.Fields(strFieldName)
      End If
   Next nIndex
   ' 關閉 RecordSet
   rsTmp.Close
   
   ' 針對以下特殊欄位設定值
   ' 費用
   SetFieldData "CP16", Empty
   ' 規費
   SetFieldData "CP17", Empty
   ' 點數
   SetFieldData "CP18", Empty
   ' 後金
   SetFieldData "CP19", Empty
   ' 90.12.19 modify by louis
   SetFieldData "CP33", Empty
   SetFieldData "CP34", Empty
   SetFieldData "CP60", Empty
   SetFieldData "CP61", Empty
   SetFieldData "CP62", Empty
   SetFieldData "CP63", Empty
   SetFieldData "CP73", Empty
   SetFieldData "CP74", Empty
   SetFieldData "CP75", Empty
   SetFieldData "CP76", Empty
   SetFieldData "CP77", Empty
   SetFieldData "CP78", Empty
   SetFieldData "CP79", Empty
   ' 是否向客戶收款
   SetFieldData "CP20", "N"
   ' 是否算案件數
   SetFieldData "CP26", "N"
   ' 是否開電腦收據
   SetFieldData "CP32", "N"
   ' 90.12.19 modify by louis
   SetFieldData "CP21", "Y"
   
   ' 取得所有待新增的商標基本檔串列
   '2006/10/11 MODIFY BY SONIA 領土延伸之案號發文時,只針對該領土延伸之指定國家子案產生
   '                           母案延展時則針對所有指定國家包含領土延伸之指定國家子案
   'strSQL = "SELECT TM01, TM02, TM03, TM04 FROM TradeMark " & _
   '         "WHERE TM01 = '" & strTM01 & "' AND " & _
   '               "substr(TM02,1,5) = substr('" & strTM02 & "',1,5) AND " & _
   '               "TM03 <> '0' AND " & _
   '               "TM04 <> '00' "
   'Modify By Sindy 2012/7/4 +strCP10
   If strCP10 = "105" Then
      strSql = "SELECT TM01, TM02, TM03, TM04 FROM TradeMark,NextProgress " & _
               "WHERE NP02 = '" & strTM01 & "' AND " & _
                     "substr(NP03,1,5) = '" & Left(strTM02, 5) & "' " & _
                     "AND NP07='" & strCP10 & "' AND NP06='Y' AND NP08=" & DBDATE(strCP06) & _
                     " AND TM01=NP02 AND TM02=NP03 AND TM03=NP04 AND TM04=NP05 "
   Else
   '2012/7/4 End
      If TFCountry <> "" Then
         strSql = "SELECT TM01, TM02, TM03, TM04 FROM TradeMark " & _
                  "WHERE TM01 = '" & strTM01 & "' AND " & _
                        "substr(TM02,1,5) = substr('" & strTM02 & "',1,5) AND " & _
                        "TM03 <> '0' AND " & _
                        "TM04 <> '00' "
      Else
         strSql = "SELECT TM01, TM02, TM03, TM04 FROM TradeMark " & _
                  "WHERE TM01 = '" & strTM01 & "' AND " & _
                        "substr(TM02,1,6) = substr('" & strTM02 & "',1,6) AND " & _
                        "TM03 <> '0' AND " & _
                        "TM04 <> '00' "
      End If
      '2006/10/11 END
      'add by nickc 2006/06/02
      If TFCountry <> "" Then
           strSql = strSql & " and tm10 in (" & TFCountry & ")"
      End If
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         ' 設定本所案號
         SetFieldData "CP01", rsTmp.Fields("TM01")
         SetFieldData "CP02", rsTmp.Fields("TM02")
         SetFieldData "CP03", rsTmp.Fields("TM03")
         SetFieldData "CP04", rsTmp.Fields("TM04")
         ' 設定總收文號
         strTemp = Empty
         strTemp = AutoNo("B", 6)
         SetFieldData "CP09", strTemp
         ' 寫檔
         WriteCaseProgress
         ' 下一筆
         rsTmp.MoveNext
      Loop
   End If
   ' 設定產生成功
   CopyCaseProgress = True
   ' 結束作業
   OnTerminate
EXITSUB:
   Set rsTmp = Nothing
End Function
