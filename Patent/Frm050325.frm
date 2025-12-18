VERSION 5.00
Begin VB.Form frm050325 
   BorderStyle     =   1  '單線固定
   Caption         =   "美國發明退公開費報表/指示信"
   ClientHeight    =   3690
   ClientLeft      =   3045
   ClientTop       =   1515
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4110
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   3192
      TabIndex        =   12
      Top             =   36
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   2412
      TabIndex        =   11
      Top             =   36
      Width           =   756
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   300
      Left            =   5235
      MaxLength       =   1
      TabIndex        =   13
      Top             =   5325
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   180
      TabIndex        =   15
      Top             =   2160
      Width           =   3840
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   0
         Left            =   1215
         MaxLength       =   3
         TabIndex        =   6
         Top             =   690
         Width           =   465
      End
      Begin VB.CheckBox Check3 
         Caption         =   "單號"
         Height          =   255
         Left            =   1350
         TabIndex        =   21
         Top             =   270
         Width           =   870
      End
      Begin VB.CheckBox Check4 
         Caption         =   "雙號"
         Height          =   255
         Left            =   2430
         TabIndex        =   20
         Top             =   270
         Width           =   960
      End
      Begin VB.OptionButton Option2 
         Caption         =   "案號類別："
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   4
         Top             =   270
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton Option2 
         Caption         =   "本所案號："
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   690
         Width           =   1260
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   4
         Left            =   1845
         MaxLength       =   1
         TabIndex        =   10
         Text            =   "Y"
         Top             =   1080
         Width           =   330
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   3
         Left            =   3075
         MaxLength       =   2
         TabIndex        =   9
         Top             =   690
         Width           =   375
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   2
         Left            =   2715
         MaxLength       =   1
         TabIndex        =   8
         Top             =   690
         Width           =   255
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   1
         Left            =   1785
         MaxLength       =   6
         TabIndex        =   7
         Top             =   690
         Width           =   810
      End
      Begin VB.Line Line6 
         X1              =   3159
         X2              =   1395
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "是否修改指示信：          ( Y:Word )"
         Height          =   180
         Left            =   405
         TabIndex        =   18
         Top             =   1110
         Width           =   2670
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   180
      TabIndex        =   14
      Top             =   570
      Width           =   3840
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   6
         Left            =   1125
         MaxLength       =   1
         TabIndex        =   25
         Text            =   "2"
         Top             =   900
         Width           =   255
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   5
         Left            =   855
         MaxLength       =   1
         TabIndex        =   23
         Text            =   "1"
         Top             =   525
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "雙號"
         Height          =   255
         Left            =   2385
         TabIndex        =   2
         Top             =   210
         Width           =   960
      End
      Begin VB.CheckBox Check1 
         Caption         =   "單號"
         Height          =   255
         Left            =   1215
         TabIndex        =   1
         Top             =   210
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "輸出方式：  　   ( 1.螢幕 2.印表機 )"
         Height          =   180
         Index           =   9
         Left            =   180
         TabIndex        =   24
         Top             =   930
         Width           =   3090
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "退費別：       ( 1:未退, 2:已退, 空白:全部 )"
         Height          =   180
         Left            =   180
         TabIndex        =   22
         Top             =   570
         Width           =   3180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "案號類別："
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   210
         Width           =   900
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "指示信"
      Height          =   255
      Index           =   1
      Left            =   45
      TabIndex        =   3
      Top             =   1920
      Width           =   1170
   End
   Begin VB.OptionButton Option1 
      Caption         =   "報表"
      Height          =   195
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   360
      Value           =   -1  'True
      Width           =   960
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "(1. 管制表 2. 定稿)"
      Height          =   180
      Left            =   5955
      TabIndex        =   17
      Top             =   5325
      Width           =   1425
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "列印格式:"
      Height          =   180
      Left            =   4170
      TabIndex        =   16
      Top             =   5340
      Width           =   765
   End
End
Attribute VB_Name = "frm050325"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
'Create by Morgan 2007/12/25
Option Explicit

Dim PLeft() As Integer, PColName() As String, strTemp() As String
Dim iPrint As Integer, iPage As Integer
Dim m_iStartX As Integer, m_iStartY As Integer
Dim m_iPageHeight As Integer, m_iLineHeight As Integer, m_iMargin As Integer
Dim m_RptType As String

Private Sub Process()
Dim stVTB As String, stCon As String, stLetter As String, stCon1 As String
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/01/22 清除查詢印表記錄檔欄位
   
   '報表
   If Option1(0).Value = True Then
      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption 'Add By Sindy 2010/01/22
      If Check1.Value + Check2.Value = 0 Then
         MsgBox "案號類別至少需勾選一種！"
         Exit Sub
         
      ElseIf txt2(6) = "" Then
         MsgBox "請輸入輸出方式！"
         txt2(6).SetFocus
         Exit Sub
         
      '只勾一種
      ElseIf Check1.Value + Check2.Value = 1 Then
         '單
         If Check1.Value = 1 Then
            stCon = " and mod(pa02,2)=1"
            pub_QL05 = pub_QL05 & ";" & Label1(0) & Check1.Caption 'Add By Sindy 2010/01/22
         '雙
         Else
            stCon = " and mod(pa02,2)=0"
            pub_QL05 = pub_QL05 & ";" & Label1(0) & Check2.Caption 'Add By Sindy 2010/01/22
         End If
      End If
   '指示信
   Else
      pub_QL05 = pub_QL05 & ";" & Option1(1).Caption 'Add By Sindy 2010/01/22
      '指定案號
      If Option2(1).Value = True Then
         If Trim(txt2(0) & txt2(1)) = "" Then
            MsgBox "請輸入本所案號！"
            txt2(0).SetFocus
            Exit Sub
         Else
            txt2(2) = Val(txt2(2))
            txt2(3) = Format(Val(txt2(3)), "00")
            stCon = " and pa01='" & txt2(0) & "' and pa02='" & txt2(1) & "' and pa03='" & txt2(2) & "' and pa04='" & txt2(3) & "'"
            pub_QL05 = pub_QL05 & ";" & Option2(1).Caption & txt2(0) & "-" & txt2(1) & "-" & txt2(2) & "-" & txt2(3) 'Add By Sindy 2010/01/22
         End If
         If txt2(4) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label9 & txt2(4) 'Add By Sindy 2010/01/22
         End If
      ElseIf Check3.Value + Check4.Value = 0 Then
         MsgBox "案號類別至少需勾選一種！"
         Exit Sub
      '只勾一種
      ElseIf Check3.Value + Check4.Value = 1 Then
         '單
         If Check3.Value = 1 Then
            stCon = " and mod(pa02,2)=1"
            pub_QL05 = pub_QL05 & ";" & Option2(0).Caption & Check3.Caption 'Add By Sindy 2010/01/22
         '雙
         Else
            stCon = " and mod(pa02,2)=0"
            pub_QL05 = pub_QL05 & ";" & Option2(0).Caption & Check4.Caption 'Add By Sindy 2010/01/22
         End If
      End If
   End If
   
   stVTB = "select pa01,pa02,pa03,pa04,pa12,pa14,cp09,cp13,cp44,cp60,cp45,pa22" & _
      " from patent,caseprogress where pa01='CFP' and pa09='101' and pa08='1'" & _
      " and instr(pa15,'B1')>0 and pa13 is null and cp01(+)=pa01 and cp02(+)=pa02" & _
      " and cp03(+)=pa03 and cp04(+)=pa04 and (cp10='601' or cp10='217') and cp61 is not null and cp27>0" & stCon
      
   'Add by Morgan 2008/2/14 加判斷帳單有含公開費的才要
   stVTB = stVTB & " and exists(select * from acc151 where axf02=cp09 and axf16='Y')"

   '2010/3/1 ADD BY SONIA CFP-012544,11893美國專利局回覆已超過二年請求時效不可退費
   'Modify by Morgan 2010/6/1 +012813
   stVTB = stVTB & " AND PA02<>'012544' AND PA02<>'011893' AND PA02<>'012813'"
   '2010/3/1 END
   
   '2011/5/27 add by sonia Y48916已被撤照不再列印
   stVTB = stVTB & " AND CP44<>'Y48916000'"
   '2011/5/27 end
   stVTB = stVTB & " AND CP02<>'012692'"    '2011/11/21禧佩說王副總指示已銷卷且已計算結餘,代理人也沒回覆,不再列印
   
   If Option1(0).Value = True Then
      Select Case txt2(5).Text
         Case "1"
            stCon1 = " AND Y.cp09 is null"
         Case "2"
            stCon1 = " AND Y.cp09 is NOT null"
      End Select
      pub_QL05 = pub_QL05 & ";" & Label2 & txt2(5) 'Add By Sindy 2010/01/22
      '1. 代理人 2. 案號 3. 智權人員 4. 收據抬頭 5. 公告日(發證日) 6. 公開日
'      '查詢
'      If txt2(6) = "1" Then
'         strExc(0) = "select RTRIM(CP44||' '||FA05||' '||FA63||' '||FA64||' '||FA65) C1" & _
'            ",DECODE(Y.CP09,NULL,'','Y') C2" & _
'            ",PA01||'-'||PA02||DECODE(PA03||PA04,'000','','-'||PA03||'-'||PA04) C3" & _
'            ",CP13||' '||ST02 C4,A0K04 C5,SQLDATET(PA14) C6,SQLDATET(PA12) C7"
'      '列印
'      Else
'         strExc(0) = "select mod(pa02,2) Srt,FA05,FA63,FA64,FA65,ST02,A0K04,X.*,DECODE(Y.CP09,NULL,' ','*') PAY"
'      End If
      '查詢用欄位
      strExc(0) = "select RTRIM(CP44||' '||FA05||' '||FA63||' '||FA64||' '||FA65) C1" & _
            ",DECODE(Y.CP09,NULL,'','Y') C2" & _
            ",PA01||'-'||PA02||DECODE(PA03||PA04,'000','','-'||PA03||'-'||PA04) C3" & _
            ",CP13||' '||ST02 C4,A0K04 C5,SQLDATET(PA14) C6,SQLDATET(PA12) C7"
      '列印用欄位
      strExc(0) = strExc(0) & ",mod(pa02,2) Srt,FA05,FA63,FA64,FA65,ST02,A0K04,X.*,DECODE(Y.CP09,NULL,' ','*') PAY,cp45,pa22"
      
      '2008/1/22 modify by sonia 取消AX207>0的過濾條件,因CFP-015003未付給代理人故代理人不會退費,但已退客戶者也要剔除
      strExc(0) = strExc(0) & " from (" & stVTB & ") X,(select distinct C1.cp09" & _
         " from (" & stVTB & ") C1,acc021,ACC1P0,ACC190,ACC161,CASEPROGRESS C2" & _
         " where ax214(+)=pa01||pa02||pa03||pa04 and ax205(+)='220106' and instr(ax212,'退公開費')>0" & _
         " AND A1P22(+)=AX202 AND A1P17(+)=AX214 AND A1908(+)=substr(A1P04, 1, Length(A1P04) - 9)" & _
         " AND AXG01(+)=A1902 and (AXG04>100 OR AXG01 IS NULL)" & _
         " AND C2.cp09(+)=axg02 and (C2.cp10 IN ('601','217') OR C2.CP10 IS NULL)" & _
         ") Y,FAGENT,STAFF,acc0k0" & _
         " where Y.cp09(+)=X.cp09" & stCon1 & _
         " AND FA01(+)=SUBSTR(CP44,1,8) AND FA02(+)=SUBSTR(CP44,9) AND ST01(+)=CP13" & _
         " and a0k01(+)=cp60"
         
      If txt2(6) = "1" Then
         strExc(0) = strExc(0) & " ORDER BY 1,3,4"
      Else
         strExc(0) = strExc(0) & " ORDER BY Srt,cp44,pa01,pa02,pa03,pa04"
      End If
      pub_QL05 = pub_QL05 & ";" & Label1(9) & txt2(6) 'Add By Sindy 2010/01/22
   Else
      '2008/1/22 modify by sonia 取消AX207>0的過濾條件,因CFP-015003未付給代理人故代理人不會退費,但已退客戶者也要剔除
      strExc(0) = "select mod(pa02,2) Srt,X.*" & _
         " from (" & stVTB & ") X,(select distinct C1.cp09" & _
         " from (" & stVTB & ") C1,acc021,ACC1P0,ACC190,ACC161,CASEPROGRESS C2" & _
         " where ax214(+)=pa01||pa02||pa03||pa04 and ax205(+)='220106' and instr(ax212,'退公開費')>0" & _
         " AND A1P22(+)=AX202 AND A1P17(+)=AX214 AND A1908(+)=substr(A1P04, 1, Length(A1P04) - 9)" & _
         " AND AXG01(+)=A1902 and (AXG04>100 OR AXG01 IS NULL)" & _
         " AND C2.cp09(+)=axg02 and (C2.cp10='601' or C2.cp10='217' OR C2.CP10 IS NULL)) Y" & _
         " where Y.cp09(+)=X.cp09 AND Y.cp09 is null" & _
         " ORDER BY Srt,cp44,pa01,pa02,pa03,pa04"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.RecordCount > 0 Then
         InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2010/01/22
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/01/22
      End If
      If Option1(0).Value = True Then
         If txt2(6) = "1" Then
            SetGrid RsTemp
         Else
            DoPrint RsTemp
         End If
      Else
         With RsTemp
         If Option2(1).Value = True Then
            StartLetter "01", .Fields("cp09"), "99"
            If txt2(4) = "Y" Then
               NowPrint .Fields("cp09"), "01", "99", True, strUserNum, , , True, stLetter
               NowPrint .Fields("cp09"), "09", "09", True, strUserNum, , stLetter
            Else
               NowPrint .Fields("cp09"), "01", "99", False, strUserNum, , , , , 1
               NowPrint .Fields("cp09"), "09", "09", False, strUserNum
            End If
         Else
            .MoveFirst
            Do While Not .EOF
               StartLetter "01", .Fields("cp09"), "99"
               NowPrint .Fields("cp09"), "01", "99", False, strUserNum, , , , , 1
               NowPrint .Fields("cp09"), "09", "09", False, strUserNum
               .MoveNext
            Loop
         End If
         End With
      End If
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/01/22
   End If
End Sub

Private Function StartLetter(ET01 As String, ET02 As String, ET03 As String) As Boolean
   EndLetter ET01, ET02, ET03, strUserNum
   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','傳真頁數','2')"
   cnnConnection.Execute strSql
End Function

Sub GetPleft()
   Dim ii As Integer
   Printer.Orientation = 2
   m_iStartX = 500
   m_iStartY = 500
   m_iPageHeight = Printer.ScaleHeight
   m_iLineHeight = 300
   m_iMargin = 500
   
   ReDim PLeft(1 To 9)
   ReDim PColName(1 To 8)
   ReDim strTemp(1 To 8)
   
   '代理人 彼所案號 證書號 本所案號 智權人員 收據抬頭 公告日(發證日) 公開日
   ii = 1
   PLeft(ii) = 500
   PColName(ii) = "代理人"
   PLeft(ii + 1) = PLeft(ii) + 3000
   
   ii = ii + 1
   PColName(ii) = "彼所案號"
   PLeft(ii + 1) = PLeft(ii) + 3100
   
   ii = ii + 1
   PColName(ii) = "證書號"
   PLeft(ii + 1) = PLeft(ii) + 2100
   
   ii = ii + 1
   PColName(ii) = "本所案號"
   PLeft(ii + 1) = PLeft(ii) + 1600
   
   ii = ii + 1
   PColName(ii) = "智權人員"
   PLeft(ii + 1) = PLeft(ii) + 1500
   
   ii = ii + 1
   PColName(ii) = "收據抬頭"
   PLeft(ii + 1) = PLeft(ii) + 2200
   
   ii = ii + 1
   PColName(ii) = "公告日"
   PLeft(ii + 1) = PLeft(ii) + 1000
   
   ii = ii + 1
   PColName(ii) = "公開日"
   PLeft(ii + 1) = PLeft(ii) + 1000
End Sub

Public Function DoPrint(ByRef p_Rst As ADODB.Recordset) As Boolean
   Dim iRecs As Integer, iLstType As Integer
   GetPleft
On Error GoTo ErrHnd
   With p_Rst
      .MoveFirst
      iLstType = .Fields("Srt")
      If iLstType = 0 Then
         m_RptType = "雙號"
      Else
         m_RptType = "單號"
      End If
      iPage = 0
      iRecs = 0
      PrintPageHeader
      PrintPageHeader1
      Do While Not .EOF
         If iLstType <> .Fields("Srt") Then
            Call PrintReportFooter(iRecs)
            iRecs = 0
            iLstType = .Fields("Srt")
            If iLstType = 0 Then
               m_RptType = "雙號"
            Else
               m_RptType = "單號"
            End If
            Printer.NewPage
            PrintPageHeader
            PrintPageHeader1
         End If
         iRecs = iRecs + 1
         strTemp(1) = CutWords(.Fields("cp44") & " " & .Fields("fa05") & .Fields("fa63") & .Fields("fa64") & .Fields("fa65"), 1)
         strTemp(2) = "" & .Fields("cp45")
         strTemp(3) = "" & .Fields("pa22")
         strTemp(4) = CutWords(.Fields("PAY") & .Fields("pa01") & "-" & .Fields("pa02") & IIf(.Fields("pa03") & .Fields("pa04") = "000", "", "-" & .Fields("pa03") & "-" & .Fields("pa04")), 4)
         strTemp(5) = CutWords(.Fields("cp13") & " " & .Fields("st02"), 5)
         strTemp(6) = CutWords("" & .Fields("a0k04"), 6)
         strTemp(7) = ChangeWStringToTDateString("" & .Fields("pa14"))
         strTemp(8) = ChangeWStringToTDateString("" & .Fields("pa12"))
         PrintDetail
         .MoveNext
      Loop
      Call PrintReportFooter(iRecs)
      Printer.EndDoc
   End With
   DoPrint = True
   Exit Function

ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Function

Sub PrintDetail()
   Dim iCol As Integer
   PrintNewLine
   For iCol = LBound(strTemp) To UBound(strTemp)
      Printer.CurrentX = PLeft(iCol)
      Printer.CurrentY = iPrint
      Printer.Print strTemp(iCol)
   Next
End Sub

'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)
   PrintNewLine 3
   DrawLine
   PrintNewLine
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "合計： " & iRecCount & " 筆"
   PrintMemo
End Sub

Private Function CutWords(p_Data As String, p_ID As Integer) As String
   Dim iLen As Integer, stNew As String, stNew1 As String
   iLen = Len(p_Data)
   For intI = 1 To iLen
      stNew1 = stNew & Mid(p_Data, intI, 1)
      If PLeft(p_ID) + Printer.TextWidth(stNew1) > PLeft(p_ID + 1) - 200 Then
         Exit For
      Else
         stNew = stNew1
      End If
   Next
   CutWords = stNew
End Function

Sub PrintPageHeader()
   Dim strTmp As String
   
   strTmp = "美國發明可退公開費報表"
   iPrint = m_iStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   iPrint = iPrint + 500
   
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   PrintNewLine
   
   strTmp = "案號類別：" & m_RptType
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = Printer.ScaleWidth - m_iMargin - 2500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   PrintNewLine
   
   iPage = iPage + 1
   Printer.CurrentX = Printer.ScaleWidth - m_iMargin - 2500
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
   PrintNewLine
   
   Printer.Font.Size = 10
End Sub

Private Sub PrintPageHeader1()
   PrintNewLine
   For intI = 1 To UBound(PColName)
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      Printer.Print PColName(intI)
   Next
   PrintNewLine
   DrawLine
End Sub

Private Sub DrawLine()
   Printer.DrawWidth = 5
   Printer.Line (PLeft(LBound(PLeft)), iPrint)-(PLeft(UBound(PLeft)), iPrint)
   iPrint = iPrint - m_iLineHeight / 2
End Sub

Private Sub PrintNewLine(Optional ByVal p_iExtraLines As Integer = 2)

   iPrint = iPrint + m_iLineHeight
   If iPrint >= (m_iPageHeight - m_iMargin - p_iExtraLines * m_iLineHeight) Then
      DrawLine
      PrintMemo
      Printer.NewPage
      PrintPageHeader
      PrintPageHeader1
      iPrint = iPrint + m_iLineHeight
   End If
   
End Sub

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 1
         Screen.MousePointer = vbHourglass
         Process
         Screen.MousePointer = vbDefault
      Case 2
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Option1_Click 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm050325 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   Dim bolEnable1 As Boolean
   If Index = 0 Then
      bolEnable1 = True
   Else
      bolEnable1 = False
   End If
   Frame1.Enabled = bolEnable1
   Frame2.Enabled = Not bolEnable1
End Sub

Private Sub Option2_Click(Index As Integer)
   Dim bolEnable As Boolean
   If Index = 0 Then
      bolEnable = False
   Else
      bolEnable = True
   End If
   Check3.Enabled = Not bolEnable
   Check4.Enabled = Not bolEnable
   txt2(0).Enabled = bolEnable
   txt2(1).Enabled = bolEnable
   txt2(2).Enabled = bolEnable
   txt2(3).Enabled = bolEnable
   txt2(4).Enabled = bolEnable
End Sub

Private Sub txt2_GotFocus(Index As Integer)
   TextInverse txt2(Index)
End Sub

Private Sub txt2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 5, 6
         If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub SetGrid(p_Rst As ADODB.Recordset)
   With frm050325_1
      .Show
      .grdDataList.Visible = False
      Set .grdDataList.Recordset = p_Rst.Clone
      .grdDataList.FormatString = "代理人　　　　　　　　|已退|本所案號　　　|智權人員　　　|收據抬頭　　　　　　|公告日　　|公開日　　"
      For intI = 0 To .grdDataList.Cols - 1
         Select Case intI
            '置中
            Case 1, 5, 6
               .grdDataList.ColAlignment(intI) = 4
            '靠左
            Case 0, 2, 3, 4
               .grdDataList.ColAlignment(intI) = 1
            Case Else
               .grdDataList.ColWidth(intI) = 0
         End Select
      Next
      .grdDataList.Visible = True
   End With
   Me.Hide
End Sub

Private Sub PrintMemo()
   Dim iSize As Integer
   iSize = Printer.Font.Size
   Printer.Font.Size = 10
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = m_iPageHeight - m_iMargin - Printer.TextHeight("註")
   Printer.Print "註：1.本所案號欄位有加 * 號表示已退費"
   Printer.Font.Size = iSize
End Sub
