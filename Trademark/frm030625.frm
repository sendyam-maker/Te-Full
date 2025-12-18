VERSION 5.00
Begin VB.Form frm030625 
   BorderStyle     =   1  '單線固定
   Caption         =   "同業台灣各區類別數比較"
   ClientHeight    =   3720
   ClientLeft      =   2790
   ClientTop       =   3945
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4650
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   1
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   1
      Top             =   720
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   345
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   2955
      Width           =   2175
   End
   Begin VB.TextBox txt1 
      Height          =   345
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtDate 
      Height          =   345
      Index           =   0
      Left            =   1890
      MaxLength       =   5
      TabIndex        =   0
      Top             =   720
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3750
      TabIndex        =   5
      Top             =   90
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Excel(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2700
      TabIndex        =   4
      Top             =   90
      Width           =   990
   End
   Begin VB.Label Label3 
      Caption         =   "PS：Excel儲存於"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   720
      TabIndex        =   14
      Top             =   3405
      Width           =   3195
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   3090
      Y1              =   880
      Y2              =   880
   End
   Begin VB.Label Label4 
      Caption         =   "6."
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   13
      Top             =   3000
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "5."
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   12
      Top             =   2565
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "4.台灣國際"
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "3.理律法律"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   10
      Top             =   1995
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "2.聖島國際"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   9
      Top             =   1725
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "1.台一國際"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "事務所或代理人名稱："
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "公報年月："
      Height          =   210
      Left            =   930
      TabIndex        =   6
      Top             =   810
      Width           =   1005
   End
End
Attribute VB_Name = "frm030625"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Lydia 2015/12/14 同業台灣各區類別數比較
Option Explicit

Private Sub cmdOK_Click(Index As Integer)
Dim Cancel As Boolean
   
   Select Case Index
      Case 0

         For intI = 0 To 1
             If Trim(txtDate(intI)) = "" Then
                MsgBox IIf(intI = 0, "起始", "截止") & "公報年月不可空白！", vbInformation, "輸入錯誤！"
                txtDate(intI).SetFocus
                Exit Sub
             End If
             txtDate_Validate intI, Cancel
             If Cancel = True Then
                txtDate(intI).SetFocus
             End If
         Next
         If Val(txtDate(1)) < Val(txtDate(0)) Then
            MsgBox "截止年月必須大於起始年月！", vbInformation, "輸入錯誤！"
            txtDate(1).SetFocus
            Exit Sub
         End If
         Screen.MousePointer = vbHourglass
         If StrMenu = False Then
         End If
         Screen.MousePointer = vbDefault
         
      Case 1
         Unload Me
   End Select
End Sub
Private Function StrMenu() As Boolean
Dim inR1 As Integer, inXr As Integer
Dim m_rs As New ADODB.Recordset
Dim strVol1_S As String, strVol2_S As String
Dim strVol1_E As String, strVol2_E As String
Dim xlsSalesPoint As New Excel.Application
Dim wks625 As New Worksheet
Dim tmpArr As Variant
Dim strTemp As String
Dim strPath As String, strTempFile As String
Dim cRange  As Integer
Dim cCol(1 To 2) As Integer '表-起始欄位
Dim xRows(1 To 2) As Integer '表-目前列位置
Dim strR(1 To 2) As Integer  '表-起始列位置
Dim strAd(1 To 2) As String  '表-欄位置
Dim strAt(1 To 2) As String '表-抬頭
Dim strAr(1 To 2) As String '事務所-列位置

   StrMenu = False

   strExc(1) = ""
     
   'Added By Lydia 2022/01/13 查詢印表記錄檔欄位
   ClearQueryLog (Me.Name)
   pub_QL05 = pub_QL05 & ";" & Label1.Caption & txtDate(0) & "-" & txtDate(1)
   pub_QL05 = pub_QL05 & ";" & Label4(0).Caption & Label4(1).Caption & "、" & Label4(2).Caption & "、" & Label4(3).Caption & "、" & Label4(4).Caption
   If txt1(0) <> "" Then pub_QL05 = pub_QL05 & "、5." & txt1(0)
   If txt1(1) <> "" Then pub_QL05 = pub_QL05 & "、6." & txt1(1)
  'end 2022/01/13
  
   Call Pub_ChgDateToTMBM07(txtDate(0), strVol1_S, strVol2_S)
   Call Pub_ChgDateToTMBM07(txtDate(1), strVol1_E, strVol2_E)
      '自訂同業5,6
      If txt1(0) <> "" Then
           'Modified by Lydia 2018/04/25 'T'=TA01 -> 'T'=TA01(+)
           'Memo by Lydia 2021/01/11 這段修改應該源自於2018年幫"葉特助抓商標公報統計資料->調整商標公報國外地區的代理人(TMBM06)空白也要抓到資料"；
           'Modified by Lydia 2021/01/11 'T'=TA01(+) -> 'T'=TA01 ;  遇到針對出名代理人(事務所)會造成SQL執行時間過長，所以對出名代理人(事務所)的查詢一定要抓到TMBM06。
           strExc(1) = strExc(1) & " union SELECT '50' ord1,TMBM01,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,DECODE(SUBSTR(NA02,1,1),'A',SUBSTR(NA02,1,3),SUBSTR(NA02,1,1)) NA00,TMBM08" & _
                    " FROM TMBULLETIN, TAGENT, NATION" & _
                    " WHERE TMBM05=NA03(+) AND length(na01)=3 AND TMBM06=TA03(+) AND 'T'=TA01 and tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
                    " and ta04 in (" & Pub_GetTA04(txt1(0)) & ") "
      End If
      If txt1(1) <> "" Then
           'Modified by Lydia 2018/04/25 'T'=TA01 -> 'T'=TA01(+)
           'Modified by Lydia 2021/01/11 'T'=TA01(+) -> 'T'=TA01 ;
           strExc(1) = strExc(1) & " union SELECT '60' ord1,TMBM01,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,DECODE(SUBSTR(NA02,1,1),'A',SUBSTR(NA02,1,3),SUBSTR(NA02,1,1)) NA00,TMBM08" & _
                    " FROM TMBULLETIN, TAGENT, NATION" & _
                    " WHERE TMBM05=NA03(+) AND length(na01)=3 AND TMBM06=TA03(+) AND 'T'=TA01 and tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
                    " and ta04 in (" & Pub_GetTA04(txt1(1)) & ") "
      End If
      
    'Modified by Lydia 2018/04/25 'T'=TA01 -> 'T'=TA01(+)
    'Modified by Lydia 2021/01/11 'T'=TA01(+) -> 'T'=TA01 ;
   strExc(1) = "SELECT '10' ORD1,TMBM01,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,DECODE(SUBSTR(NA02,1,1),'A',SUBSTR(NA02,1,3),SUBSTR(NA02,1,1)) NA00,TMBM08 " & _
               "From TMBULLETIN, TAGENT, NATION,trademark,caseprogress " & _
               "WHERE TMBM05=NA03(+) AND length(na01)=3 AND TMBM06=TA03(+) AND 'T'=TA01 and tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
               " and tmbm06 ='林晉章' and tmbm01=tm15(+) and tm15 is not null and tm16='1' and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp10='101' " & _
               "union SELECT decode(ta04,'聖島國際','20','理律法律','30','台灣國際','40') ord1 " & _
               ",TMBM01,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,DECODE(SUBSTR(NA02,1,1),'A',SUBSTR(NA02,1,3),SUBSTR(NA02,1,1)) NA00,TMBM08 " & _
               "FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05=NA03(+) AND length(na01)=3 AND TMBM06=TA03(+) AND 'T'=TA01 and tmbm07>=" & strVol1_S & " And tmbm07<=" & strVol2_E & _
               " and ta04 in ('聖島國際','理律法律','台灣國際') " & strExc(1)
   
   strSql = "select ord1,substr(na00,1,1) ord2,TA04,NA00,sum(counting(tmbm08)) vc from (" & _
             strExc(1) & ") group by ord1,substr(na00,1,1),ta04,na00 order by 2,1,4 "
      
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      StrMenu = True
      InsertQueryLog (m_rs.RecordCount)  'Added by Lydia 2022/01/13
      cCol(1) = Asc("a")
      cCol(2) = Asc("c")
      
      'NA02=A11->北區,A12-桃竹苗,A21-中區,A22-彰投,A31-南區,A41-高區,A51-花東,國內合計
      strAt(1) = "北區,桃竹苗,中區,彰投,南區,高區,花東,國內"
      strAt(2) = "大陸,國外,全所"
      strAd(1) = "A1101,A1202,A2103,A2204,A3105,A4106,A5107,AXX08"
      strAd(2) = "BXX01,CXX02,TXX03"
      strAr(1) = "10R01,20R02,30R03,40R04,50R05,60R06"
      strAr(2) = strAr(1)
      strExc(1) = ""
      m_rs.MoveFirst
      xRows(1) = 1
      strTempFile = Me.Caption & txtDate(0) & "至" & txtDate(1) & "-" & ACDate(ServerDate) & ServerTime & MsgText(43)
      strPath = strExcelPath & strTempFile
      
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = "" Then
         MkDir strExcelPath
      End If
      If Dir(strPath) <> "" Then
         Kill strPath
      End If
      
      xlsSalesPoint.SheetsInNewWorkbook = 3 'Modify by Amy 2021/06/21 Added by Lydia 2019/03/13 預設工作表數量
      xlsSalesPoint.Workbooks.add
      Set wks625 = xlsSalesPoint.Worksheets(1)
      wks625.PageSetup.Orientation = xlPortrait '直印
      '抬頭
       wks625.PageSetup.PrintTitleRows = "$1:$3"

       For intI = cCol(1) To cCol(1) + Val(Right(strAd(1), 2))
          wks625.Columns(Chr(intI) & ":" & Chr(intI)).ColumnWidth = 8
       Next

       wks625.Range("a1").Value = Me.Caption & " " & Mid(ChangeTStringToTDateString(txtDate(0) & "01"), 1, 6) & "至" & Mid(ChangeTStringToTDateString(txtDate(1) & "01"), 1, 6)
       wks625.Range("a1:" & Chr(cCol(1) + Val(Right(strAd(1), 2))) & "1").MergeCells = True

       wks625.Range("a2").Value = "(以類計)"
       wks625.Range("a2:" & Chr(cCol(1) + Val(Right(strAd(1), 2))) & "2").MergeCells = True
       
       xRows(1) = 3
       wks625.Range(Chr(cCol(1)) & xRows(1)).Value = "事務所"
       tmpArr = Empty
       tmpArr = Split(strAt(1), ",")
       For intI = 0 To UBound(tmpArr)
           If tmpArr(intI) <> "" Then
              wks625.Range(Chr(cCol(1) + intI + 1) & xRows(1)).Value = tmpArr(intI)
           End If
       Next intI
       
      wks625.Range("a1:" & Chr(cCol(1) + Val(Right(strAd(1), 2))) & xRows(1)).HorizontalAlignment = xlCenter
      wks625.Range("a1:" & Chr(cCol(1) + Val(Right(strAd(1), 2))) & xRows(1)).VerticalAlignment = xlBottom
      
      xRows(1) = xRows(1): strR(1) = xRows(1)
       
      Do While Not m_rs.EOF
         '上半部=國內
         If m_rs.Fields("ord2") = "A" Then
           '抓資料的明細欄位置
            inR1 = InStr(strAd(1), Trim(m_rs.Fields("na00")))
            If inR1 > 0 Then
                If strTemp <> m_rs.Fields("ord1") Then
                   xRows(1) = xRows(1) + 1
                End If

                inXr = cCol(1) + Val(Mid(strAd(1), inR1 + 3, 2))
                inR1 = strR(1) + Val(Mid(strAr(1), InStr(strAr(1), Trim(m_rs.Fields("ord1"))) + 3, 2))
                '事務所 抬頭
                If wks625.Range(Chr(cCol(1)) & inR1).Value = "" Then
                   wks625.Range(Chr(cCol(1)) & inR1).Value = Trim(m_rs.Fields("ta04"))
                   wks625.Range(Chr(cCol(1) + Val(Right(strAd(1), 2))) & inR1).Formula = "=SUM(" & Chr(cCol(1) + 1) & inR1 & ":" & Chr(cCol(1) + Val(Right(strAd(1), 2)) - 1) & inR1 & ")"
                ElseIf Trim(m_rs.Fields("ta04")) <> "" Then
                       strExc(6) = Trim(wks625.Range(Chr(cCol(1)) & inR1).Value)
                       If InStr(strExc(6), Trim(m_rs.Fields("ta04"))) = 0 Then
                          wks625.Range(Chr(cCol(1)) & inR1).Value = strExc(6) & "、" & Trim(m_rs.Fields("ta04"))
                       End If
                End If

                wks625.Range(Chr(inXr) & inR1).Value = Trim(m_rs.Fields("vc"))
            End If
            
         '下半部=大陸,國外,全所
         Else
            If xRows(2) = Empty Or xRows(2) = 0 Then
               xRows(2) = xRows(1) + 2
               tmpArr = Empty
               tmpArr = Split(strAt(2), ",")
               wks625.Range(Chr(cCol(2)) & xRows(2)).Value = "事務所"
               For intI = 0 To UBound(tmpArr)
                If tmpArr(intI) <> "" Then
                   wks625.Range(Chr(cCol(2) + intI + 1) & xRows(2)).Value = tmpArr(intI)
                End If
               Next intI
               
               wks625.Range("a" & xRows(2) & ":" & Chr(cCol(2) + Val(Right(strAd(2), 2))) & xRows(2)).HorizontalAlignment = xlCenter
               wks625.Range("a" & xRows(2) & ":" & Chr(cCol(2) + Val(Right(strAd(2), 2))) & xRows(2)).VerticalAlignment = xlBottom
               
               xRows(2) = xRows(2): strR(2) = xRows(2)
               strTemp = ""
            End If
            
           '抓資料的明細欄位置
            inR1 = InStr(strAd(2), Trim(m_rs.Fields("ord2") & "XX"))
            If inR1 > 0 Then
                If strTemp <> m_rs.Fields("ord1") Then
                   xRows(2) = xRows(2) + 1
                End If
                
                inXr = cCol(2) + Val(Mid(strAd(2), inR1 + 3, 2))
                inR1 = strR(2) + Val(Mid(strAr(2), InStr(strAr(2), Trim(m_rs.Fields("ord1"))) + 3, 2))
                '事務所 抬頭
                If wks625.Range(Chr(cCol(2)) & inR1).Value = "" Then
                   wks625.Range(Chr(cCol(2)) & inR1).Value = Trim(m_rs.Fields("ta04"))
                   wks625.Range(Chr(cCol(2) + Val(Right(strAd(2), 2))) & inR1).Formula = "=SUM(" & Chr(cCol(2) + 1) & inR1 & ":" & Chr(cCol(2) + Val(Right(strAd(2), 2)) - 1) & inR1 & ")" & _
                                                                    "+VLOOKUP(" & Chr(cCol(2)) & inR1 & "," & Chr(cCol(1)) & strR(1) & ":" & Chr(cCol(1) + Val(Right(strAd(1), 2))) & xRows(1) & "," & (cCol(1) + Val(Right(strAd(1), 2))) - Asc("a") + 1 & " ,FALSE)"
                ElseIf Trim(m_rs.Fields("ta04")) <> "" Then
                       strExc(6) = Trim(wks625.Range(Chr(cCol(2)) & inR1).Value)
                       If InStr(strExc(6), Trim(m_rs.Fields("ta04"))) = 0 Then
                          wks625.Range(Chr(cCol(2)) & inR1).Value = strExc(6) & "、" & Trim(m_rs.Fields("ta04"))
                       End If
                End If

                wks625.Range(Chr(inXr) & inR1).Value = Trim(m_rs.Fields("vc"))
            End If

         End If
         
         strTemp = m_rs.Fields("ord1")
         m_rs.MoveNext
      Loop

        '判斷若版本2007以上改變存格式
        If Val(xlsSalesPoint.Version) < 12 Then
            xlsSalesPoint.Workbooks(1).SaveAs FileName:=strPath, FileFormat:=-4143
        Else
            xlsSalesPoint.Workbooks(1).SaveAs FileName:=strPath, FileFormat:=56
        End If
        xlsSalesPoint.Workbooks.Close
        xlsSalesPoint.Quit
        'Modify by Amy 2021/06/21 原:strPath 改中文字顯示
        MsgBox "檔案已產生！" & vbCrLf & "檔案存於 " & strExcelPathN & " " & strTempFile, vbInformation
        Exit Function
   Else
      MsgBox "查詢無資料！", vbExclamation + vbOKOnly
      InsertQueryLog (0) 'Added by Lydia 2022/01/13
      Exit Function
   End If

End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   
   txtDate(0) = Left(strSrvDate(2), 5)
   txtDate(1) = Left(strSrvDate(2), 5)
   Label3.Caption = Label3 & strExcelPathN 'Modify by Amy 2021/06/21
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set frm030625 = Nothing
End Sub


Private Sub txtDate_GotFocus(Index As Integer)
   InverseTextBox txtDate(Index)
   CloseIme
End Sub

Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
Dim rsQuery As ADODB.Recordset
Dim stSQL As String, intR As Integer
Dim strVol1 As String, strVol2 As String
Dim intCnt As Integer, intC2 As Integer
   
   If txtDate(Index) <> "" Then
      If ChkDate(txtDate(Index) & "01") = False Then
          txtDate_GotFocus Index
          Cancel = True
          Exit Sub
      End If
      '檢查資料是否已存在,每月有2期公報
      Call Pub_ChgDateToTMBM07(txtDate(Index), strVol1, strVol2)
      stSQL = "select tmbm07 from tmbulletin where tmbm07>='" & strVol1 & "' and tmbm07<='" & strVol2 & "' group by tmbm07"
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         intCnt = Val("" & rsQuery.Fields(0))
      End If
      intC2 = rsQuery.RecordCount
      rsQuery.Close
      
      If intCnt = 0 Then
         MsgBox txtDate(Index) & "此月份尚無公報資料!!"
         txtDate_GotFocus (Index)
         Cancel = True
         Exit Sub
      ElseIf intC2 < 2 Then
         MsgBox txtDate(Index) & "此月份公報資料尚不足!!"
         txtDate_GotFocus (Index)
         Cancel = True
         Exit Sub
      End If
   End If

End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
    If txt1(Index) <> "" Then
       If InStr(Label4(1).Caption & ";" & Label4(2).Caption & ";" & Label4(3).Caption & ";" & Label4(4).Caption & ";", txt1(Index)) > 0 Then
          MsgBox "請勿輸入1-4的事務所名稱!", vbCritical
          txt1(Index).SetFocus
          txt1_GotFocus Index
          Cancel = True
       End If
    End If
End Sub
Private Sub txt1_GotFocus(Index As Integer)
    TextInverse txt1(Index)
    OpenIme
End Sub

