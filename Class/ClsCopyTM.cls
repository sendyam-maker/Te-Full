VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCopyTM"
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

Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
Dim m_TM01_D As String
Dim m_TM02_D As String
Dim m_TM03_D As String
Dim m_TM04_D As String

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
   ' 先清除所有欄位
   ClearField
   ' 設定欄位
   AddField "TM01", 0: AddField "TM02", 0: AddField "TM03", 0: AddField "TM04", 0: AddField "TM05", 0
   AddField "TM06", 0: AddField "TM07", 0: AddField "TM08", 0: AddField "TM09", 0: AddField "TM10", 0
   AddField "TM11", 1: AddField "TM12", 0: AddField "TM13", 1: AddField "TM14", 1: AddField "TM15", 0
   AddField "TM16", 0: AddField "TM17", 0: AddField "TM18", 0: AddField "TM19", 0: AddField "TM20", 1
   AddField "TM21", 1: AddField "TM22", 1: AddField "TM23", 0: AddField "TM24", 0: AddField "TM25", 0
   AddField "TM26", 0: AddField "TM27", 0: AddField "TM28", 0: AddField "TM29", 0: AddField "TM30", 1
   AddField "TM31", 0: AddField "TM32", 0: AddField "TM33", 0: AddField "TM34", 0: AddField "TM35", 0
   AddField "TM36", 1: AddField "TM37", 1: AddField "TM38", 0: AddField "TM39", 0: AddField "TM40", 0
   AddField "TM41", 0: AddField "TM42", 0: AddField "TM43", 0: AddField "TM44", 0: AddField "TM45", 0
   AddField "TM46", 0: AddField "TM47", 0: AddField "TM48", 0: AddField "TM49", 0: AddField "TM50", 0
   AddField "TM51", 0: AddField "TM52", 0: AddField "TM53", 0: AddField "TM54", 0: AddField "TM55", 0
   AddField "TM56", 0: AddField "TM57", 0: AddField "TM58", 0: AddField "TM59", 0: AddField "TM60", 1
   AddField "TM61", 1: AddField "TM62", 0: AddField "TM63", 1: AddField "TM64", 1: AddField "TM65", 0
   AddField "TM66", 0: AddField "TM67", 0: AddField "TM68", 1
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
Private Function WriteTradeMark()
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim strTemp As String
   
   bFirst = True
   strSql = "INSERT INTO TradeMark ("
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
         'edit by nickc 2005/08/08 修正遺漏逗點
         'strTemp = "'" & m_FieldList(nIndex).fiData & "'"
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

Public Function SetSrc(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String)
   m_TM01 = strTM01
   m_TM02 = strTM02
   m_TM03 = strTM03
   m_TM04 = strTM04
End Function

Public Function SetDes(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String)
   m_TM01_D = strTM01
   m_TM02_D = strTM02
   m_TM03_D = strTM03
   m_TM04_D = strTM04
End Function

Public Sub SetExtraField(ByVal strField As String, ByVal strData As String)
   SetExtraFieldData strField, strData
End Sub

Private Function IsRecordExist(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   IsRecordExist = False
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & strTM01 & "' AND " & _
                  "TM02 = '" & strTM02 & "' AND " & _
                  "TM03 = '" & strTM03 & "' AND " & _
                  "TM04 = '" & strTM04 & "' "
   ' 檢查此筆記錄是否存在
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Public Function CopyTradeMark()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim fldList() As DBFIELDITEM
   Dim fldCount As Integer
   Dim strFieldName As String
   Dim nIndex As Integer
   
   ' 先讀取商標基本檔
   If IsRecordExist(m_TM01, m_TM02, m_TM03, m_TM04) = False Then
      GoTo EXITSUB
   End If
   
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
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
   'Modified by Morgan 2022/6/7 要抓最新的欄位數
   'For nIndex = 1 To 68
   For nIndex = 69 To rsTmp.Fields.Count
      AddField rsTmp.Fields(nIndex - 1).Name, 0 '固定用文字型態
   Next
   For nIndex = 1 To m_FieldCount
      Select Case nIndex
      Case 59, 60, 61, 62, 63, 64, 57, 73, 74, 75 '特定欄位(Create, Update, 銷卷...)不設定
         strFieldName = "TM" & Format(nIndex, "00")
         SetFieldData strFieldName, Empty
      Case Else
   'end2022/6/7
   
         strFieldName = "TM" & Format(nIndex, "00")
         If IsNull(rsTmp.Fields(strFieldName)) = True Then
            SetFieldData strFieldName, Empty
         Else
            ' 設定欄位內的資料
            SetFieldData strFieldName, rsTmp.Fields(strFieldName)
         End If
         
      End Select 'Added by Morgan 2022/6/7
   Next nIndex
   ' 關閉 RecordSet
   rsTmp.Close
   
   ' 變更本所案號
   SetFieldData "TM01", m_TM01_D
   SetFieldData "TM02", m_TM02_D
   SetFieldData "TM03", m_TM03_D
   SetFieldData "TM04", m_TM04_D
   ' 變更所指定的欄位
   For nIndex = 0 To m_ExtraFieldCount - 1
      SetFieldData m_ExtraFieldList(nIndex).fiName, m_ExtraFieldList(nIndex).fiData
   Next nIndex
   
   ' 若原始檔案已存在則刪除
   If IsRecordExist(m_TM01_D, m_TM02_D, m_TM03_D, m_TM04_D) = True Then
      strSql = "DELETE FROM TradeMark " & _
               "WHERE TM01 = '" & m_TM01_D & "' AND " & _
                     "TM02 = '" & m_TM02_D & "' AND " & _
                     "TM03 = '" & m_TM03_D & "' AND " & _
                     "TM04 = '" & m_TM04_D & "' "
      cnnConnection.Execute strSql
   End If
   
   ' 寫檔
   WriteTradeMark
   
   ' 結束作業
   OnTerminate
   
   ' 設定產生成功
   CopyTradeMark = True
EXITSUB:
End Function

