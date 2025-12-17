VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCopyPA"
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

Dim m_ExtraFieldList() As DBFIELDITEM
Dim m_ExtraFieldCount As Integer

Dim m_PA01 As String
Dim m_PA02 As String
Dim m_PA03 As String
Dim m_PA04 As String
Dim m_PA01_D As String
Dim m_PA02_D As String
Dim m_PA03_D As String
Dim m_PA04_D As String

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

Private Sub ClearExtraField()
   If m_ExtraFieldCount > 0 Then
      Erase m_ExtraFieldList
   End If
   m_ExtraFieldCount = 0
End Sub

' 新增一個欄位
Private Sub SetExtraFieldData(ByVal strFieldName As String, ByVal strFieldData)
   Dim bFind As Boolean
   Dim nIndex As Integer
   bFind = False
   For nIndex = 0 To m_ExtraFieldCount - 1
      If m_ExtraFieldList(nIndex).fiName = strFieldName Then
         m_ExtraFieldList(nIndex).fiData = strFieldData
         bFind = True
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_ExtraFieldList(m_ExtraFieldCount + 1)
      m_ExtraFieldList(m_ExtraFieldCount).fiName = strFieldName
      m_ExtraFieldList(m_ExtraFieldCount).fiData = strFieldData
      m_ExtraFieldCount = m_ExtraFieldCount + 1
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 初始化
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OnInitial()
   Dim nIndex As Integer
   Dim nType As Integer
   Dim strField As String
   
   ' 先清除所有欄位
   ClearField
   ' 設定欄位
   For nIndex = 1 To 132
      If nIndex < 10 Then
         strField = "PA" & "0" & CStr(nIndex)
      Else
         strField = "PA" & CStr(nIndex)
      End If
      nType = 0
      Select Case nIndex
         Case 10, 12, 14, 20, 21, 23, 24, 25, 49, 50, 58, 93, 94, 96, 97:
            nType = 1
         Case Else:
            nType = 0
      End Select
      AddField strField, nType
   Next nIndex
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 結束
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OnTerminate()
   ClearField
   ClearExtraField
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 寫入案件進度檔中
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function WritePatent()
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim strTemp As String
   
   bFirst = True
   strSql = "INSERT INTO PATENT ("
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

Public Function SetSrc(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String)
   m_PA01 = strPA01
   m_PA02 = strPA02
   m_PA03 = strPA03
   m_PA04 = strPA04
End Function

Public Function SetDes(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String)
   m_PA01_D = strPA01
   m_PA02_D = strPA02
   m_PA03_D = strPA03
   m_PA04_D = strPA04
End Function

Public Sub SetExtraField(ByVal strField As String, ByVal strData As String)
   SetExtraFieldData strField, strData
End Sub

Private Function IsRecordExist(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   IsRecordExist = False
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & strPA01 & "' AND " & _
                  "PA02 = '" & strPA02 & "' AND " & _
                  "PA03 = '" & strPA03 & "' AND " & _
                  "PA04 = '" & strPA04 & "' "
   ' 檢查此筆記錄是否存在
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Public Function CopyPatent()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim fldList() As DBFIELDITEM
   Dim fldCount As Integer
   Dim strFieldName As String
   Dim nIndex As Integer
   
   ' 先讀取專利基本檔
   If IsRecordExist(m_PA01, m_PA02, m_PA03, m_PA04) = False Then
      GoTo EXITSUB
   End If
   
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & m_PA01 & "' AND " & _
                  "PA02 = '" & m_PA02 & "' AND " & _
                  "PA03 = '" & m_PA03 & "' AND " & _
                  "PA04 = '" & m_PA04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查此筆記錄是否存在
   If rsTmp.RecordCount <= 0 Then
      rsTmp.Close
      GoTo EXITSUB
   End If
   
   ' 初始化欄位串列
   OnInitial
   ' 讀取欄位內容
   For nIndex = 1 To 132
      If nIndex < 10 Then
         strFieldName = "PA" & "0" & CStr(nIndex)
      Else
         strFieldName = "PA" & CStr(nIndex)
      End If
      If IsNull(rsTmp.Fields(strFieldName)) = True Then
         SetFieldData strFieldName, Empty
      Else
         ' 設定欄位內的資料
         SetFieldData strFieldName, rsTmp.Fields(strFieldName)
      End If
   Next nIndex
   ' 關閉 RecordSet
   rsTmp.Close
   
   ' 變更本所案號
   SetFieldData "PA01", m_PA01_D
   SetFieldData "PA02", m_PA02_D
   SetFieldData "PA03", m_PA03_D
   SetFieldData "PA04", m_PA04_D
   ' 變更所指定的欄位
   For nIndex = 0 To m_ExtraFieldCount - 1
      SetFieldData m_ExtraFieldList(nIndex).fiName, m_ExtraFieldList(nIndex).fiData
   Next nIndex
   
   ' 若原始檔案已存在則刪除
   If IsRecordExist(m_PA01_D, m_PA02_D, m_PA03_D, m_PA04_D) = True Then
      strSql = "DELETE FROM PATENT " & _
               "WHERE PA01 = '" & m_PA01_D & "' AND " & _
                     "PA02 = '" & m_PA02_D & "' AND " & _
                     "PA03 = '" & m_PA03_D & "' AND " & _
                     "PA04 = '" & m_PA04_D & "' "
      cnnConnection.Execute strSql
   End If
   
   ' 寫檔
   WritePatent
   
   ' 結束作業
   OnTerminate
   
   ' 設定產生成功
   CopyPatent = True
EXITSUB:
End Function



