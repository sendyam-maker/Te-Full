VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPrtAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印地址條程式
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim m_DataList() As String
Dim m_DataListCount As Integer
Dim m_Language As Integer
Dim m_Copys As Integer
Dim m_NoMagazine As Boolean
Dim m_PrinterName As String

Private Sub Class_Initialize()
   m_DataListCount = 0
   m_Language = 1
   m_Copys = 1
   m_NoMagazine = False
   m_PrinterName = Empty
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 清除資料
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Clear()
   If m_DataListCount > 0 Then
      Erase m_DataList
   End If
   m_DataListCount = 0
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 新增要列印的資料
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddData(ByVal strData As String)
   Dim bFind As Boolean
   Dim nIndex As Integer
   bFind = False
   For nIndex = 0 To m_DataListCount - 1
      If m_DataList(nIndex) = strData Then
         bFind = True
         Exit For
      End If
   Next nIndex
   If Not bFind Then
      ReDim Preserve m_DataList(m_DataListCount + 1)
      m_DataList(m_DataListCount) = strData
      m_DataListCount = m_DataListCount + 1
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Input : nLanguage = 0 : 中文
'                   = 1 : 英文
'                   = 2 : 日文
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetLanguage(ByVal nLanguage As Integer)
   m_Language = nLanguage
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Input : nCopys ==> 列印的份數
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetCopys(ByVal nCopys As Integer)
   m_Copys = nCopys
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 設定印表機
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetPrinter(ByVal strPrinter As String)
   m_PrinterName = strPrinter
End Sub

Private Sub SetDefaultPrinter(ByVal strPrinterName As String)
   Dim Prn As Printer
   '搜尋 Printer
   For Each Prn In Printers
      If Prn.DeviceName = strPrinterName Then
         Set Printer = Prn
         Exit For
      End If
   Next
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印地址條
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintAddress() As Boolean
   Dim i As Integer
   Dim Page As Integer
   Dim iPrint As Integer
   Dim IntF As Integer
   Dim PriType As Integer
   Dim j As Integer
   Dim nRow As Integer
   Dim nPos As Integer
   Dim strCond As String
   Dim strOldPrinterName As String
   Dim Prn As Printer
   
   PrintAddress = False
   
   strOldPrinterName = Printer.DeviceName
   If Not IsEmptyText(m_PrinterName) Then
      SetDefaultPrinter m_PrinterName
   End If
   
   For nPos = 0 To m_DataListCount - 1
      iPrint = nPos + 1
      Select Case Mid(m_DataList(nPos), 1, 1)
         ' 申請人
         Case "X":
            strCond = ChgCustomer(m_DataList(nPos))
            If m_NoMagazine Then
               strCond = strCond & " AND (CU32<>'N' OR CU32 IS NULL)"
            End If
            ' 列印的語言
            Select Case m_Language
               Case 1  '7
                  strExc(0) = "SELECT CU30,SUBSTR(CU23,1,20),SUBSTR(CU23,21,15)," & _
                     "SUBSTR(CU04,1,20),SUBSTR(CU04,21,20),CU08,CU01||CU02 FROM CUSTOMER " & _
                     "WHERE " & strCond
                  Page = 1
               Case 2  '11
                  strExc(0) = "SELECT CU05,CU88,CU89,CU90,NA04,CU24," & _
                     "CU25,CU26,CU27,CU28,CU01||CU02 FROM CUSTOMER,NATION WHERE " & strCond & " AND CU87=NA01(+) "
                  Page = 2
               Case 3  '6
                  strExc(0) = "SELECT SUBSTR(CU29,1,20),SUBSTR(CU29,21,15)," & _
                     "SUBSTR(CU06,1,20),SUBSTR(CU06,21,20),CU08,CU01||CU02 FROM CUSTOMER " & _
                     "WHERE " & strCond
                  Page = 3
            End Select
         Case "Y":
            strCond = ChgFagent(m_DataList(nPos))
            If m_NoMagazine Then
               strCond = strCond & " AND (FA24<>'N' OR FA24 IS NULL)"
            End If
            ' 列印的語言
            Select Case m_Language
               Case 1  '5
                  strExc(0) = "SELECT SUBSTR(FA17,1,20),SUBSTR(FA17,21,15)," & _
                     "SUBSTR(FA04,1,20),SUBSTR(FA04,21,20),FA01||FA02 FROM FAGENT " & _
                     "WHERE " & strCond
                  Page = 4
               Case 2  '11
                  strExc(0) = "SELECT FA05,FA63,FA64,FA65,NA04," & _
                     "FA18,FA19,FA20,FA21,FA22,FA01||FA02 FROM FAGENT,NATION WHERE " & strCond & " AND FA55=NA01(+)"
                  Page = 2
               Case 3  '5
                  strExc(0) = "SELECT SUBSTR(FA23,1,20),SUBSTR(FA23,21,15)," & _
                     "SUBSTR(FA06,1,20),SUBSTR(FA06,21,20),FA01||FA02 FROM FAGENT " & _
                     "WHERE " & strCond
                  Page = 4
            End Select
         Case Else
            strExc(0) = "SELECT OR05,SUBSTR(OR06,1,15),SUBSTR(OR06,16,15),OR02,OR01 " & _
               "FROM ORGANIZATION WHERE OR01='" & m_DataList(nPos) & "'"
            Page = 4
      End Select
      intI = 0
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI <> 1 Then Exit Function
      Select Case Page
         Case 1
            IntF = 7
         Case 2
            IntF = 11
         Case 3
            IntF = 6
         Case 4
            IntF = 5
      End Select
      If Page <> 2 Then
         Printer.Font.Size = 12
         Printer.Height = 2200
         Printer.Width = 10000
      Else
         Printer.Font.Size = 10
         Printer.Height = 2900
         Printer.Width = 10000
      End If
      For j = 1 To Val(m_Copys)
         RsTemp.MoveFirst
         With RsTemp
            Select Case Page
               Case 2
                  Do While Not .EOF
                     nRow = 0
                     For i = 0 To IntF - 1
                        Printer.CurrentX = 1000
                        If IsNull(.Fields(i)) = False Then
                           If IsEmptyText(.Fields(i)) = False Then
                              ' 語文為英文時不空行
                              If m_Language = 2 Then
                                 Printer.CurrentY = nRow * 220
                              Else
                                 Printer.CurrentY = i * 220
                              End If
                              nRow = nRow + 1
                           End If
                        End If
                        
                        If IsNull(.Fields(i)) = False Then
                           If IsEmptyText(.Fields(i)) = False Then
                              Printer.Print .Fields(i)
                           End If
                        End If
                     Next
                     Printer.CurrentX = 4200
                     If m_Language = "2" Then
                        Printer.CurrentY = (nRow - 1) * 220
                     Else
                        Printer.CurrentY = (i - 1) * 220
                     End If
                     Printer.Print Format(iPrint, "000000")
                     'iPrint = iPrint + 1
                     Printer.NewPage
                     .MoveNext
                  Loop
               Case Else
                  Do While Not .EOF
                     nRow = 0
                     For i = 0 To IntF - 1
                        Printer.CurrentX = 1000
                        If IsNull(.Fields(i)) = False Then
                           If IsEmptyText(.Fields(i)) = False Then
                              ' 語文為英文時不空行
                              If m_Language = 2 Then
                                 Printer.CurrentY = nRow * 250
                              Else
                                 Printer.CurrentY = i * 250
                              End If
                              nRow = nRow + 1
                           End If
                        End If
                        
                        If IsNull(.Fields(i)) = False Then
                           If IsEmptyText(.Fields(i)) = False Then
                              Printer.Print .Fields(i)
                           End If
                        End If
                     Next
                     Printer.CurrentX = 5000
                     If m_Language = 2 Then
                        Printer.CurrentY = (nRow - 1) * 250
                     Else
                        Printer.CurrentY = (i - 1) * 250
                     End If
                     Printer.Print Format(iPrint, "000000")
                     'iPrint = iPrint + 1
                     Printer.NewPage
                     .MoveNext
                  Loop
            End Select
         End With
      Next
   Next
   Printer.EndDoc
   
   PrintAddress = True
   
   ' 還原之前的預設印表機
   If strOldPrinterName <> m_PrinterName Then
      SetDefaultPrinter strOldPrinterName
   End If
   
   Exit Function

ERRSECTION:
   ' 還原之前的預設印表機
   If strOldPrinterName <> m_PrinterName Then
      SetDefaultPrinter strOldPrinterName
   End If
End Function

