VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040114_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "非台灣案發文後暫緩"
   ClientHeight    =   5124
   ClientLeft      =   792
   ClientTop       =   1068
   ClientWidth     =   7920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5124
   ScaleWidth      =   7920
   Begin VB.CommandButton cmkok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   2
      Left            =   6996
      TabIndex        =   34
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmkok 
      Caption         =   "回前畫面(&U)"
      Height          =   405
      Index           =   1
      Left            =   5772
      TabIndex        =   32
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmkok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   4944
      TabIndex        =   31
      Top             =   70
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   60
      TabIndex        =   17
      Top             =   2910
      Width           =   7785
      Begin VB.TextBox Text8 
         Height          =   270
         Left            =   4680
         MaxLength       =   7
         TabIndex        =   26
         Top             =   1170
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   4680
         MaxLength       =   7
         TabIndex        =   23
         Top             =   570
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   1470
         MaxLength       =   1
         TabIndex        =   27
         Top             =   1470
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   25
         Top             =   1170
         Width           =   885
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1170
         MaxLength       =   7
         TabIndex        =   22
         Top             =   570
         Width           =   1215
      End
      Begin MSForms.TextBox Text3 
         Height          =   300
         Left            =   1170
         TabIndex        =   24
         Top             =   870
         Width           =   6555
         VariousPropertyBits=   671107099
         Size            =   "11562;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label12 
         Caption         =   "管制回覆委任代理人："
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   270
         Width           =   1875
      End
      Begin VB.Label Label29 
         Caption         =   "承辦期限："
         Height          =   255
         Left            =   3720
         TabIndex        =   30
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "法定期限："
         Height          =   255
         Left            =   3720
         TabIndex        =   29
         Top             =   570
         Width           =   975
      End
      Begin MSForms.Label Label26 
         Height          =   195
         Left            =   2100
         TabIndex        =   28
         Top             =   1200
         Width           =   1035
         VariousPropertyBits=   27
         Size            =   "1826;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label25 
         Caption         =   "是否算案件數：                 (N：不算)"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1470
         Width           =   3015
      End
      Begin VB.Label Label24 
         Caption         =   "承辦人："
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1170
         Width           =   735
      End
      Begin VB.Label Label23 
         Caption         =   "進度備註："
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   870
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "本所期限："
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   570
         Width           =   1095
      End
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1200
      TabIndex        =   33
      Top             =   1080
      Width           =   6645
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "11721;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "總收文號："
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   1200
      TabIndex        =   35
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   360
      Width           =   975
   End
   Begin MSForms.Label Label20 
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   2520
      Width           =   6645
      VariousPropertyBits=   27
      Size            =   "11721;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label19 
      Caption         =   "代理人："
      Height          =   255
      Left            =   330
      TabIndex        =   13
      Top             =   2520
      Width           =   765
   End
   Begin VB.Label Label18 
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label17 
      Caption         =   "申請國家："
      Height          =   255
      Left            =   3990
      TabIndex        =   11
      Top             =   2160
      Width           =   915
   End
   Begin MSForms.Label Label16 
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   1800
      Width           =   1935
      VariousPropertyBits=   27
      Size            =   "3413;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label15 
      Caption         =   "智權人員："
      Height          =   255
      Left            =   3990
      TabIndex        =   9
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label Label14 
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label13 
      Caption         =   "專利號數："
      Height          =   255
      Left            =   3990
      TabIndex        =   7
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label11 
      Caption         =   "承辦人："
      Height          =   255
      Left            =   330
      TabIndex        =   6
      Top             =   2160
      Width           =   765
   End
   Begin MSForms.Label Label10 
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
      VariousPropertyBits=   27
      Size            =   "2143;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      Caption         =   "案件性質："
      Height          =   255
      Left            =   150
      TabIndex        =   4
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "申請人："
      Height          =   255
      Left            =   330
      TabIndex        =   2
      Top             =   1440
      Width           =   765
   End
   Begin MSForms.Label Label6 
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   6645
      VariousPropertyBits=   27
      Size            =   "11721;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frm040114_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/22 改成Form2.0 (Text3,Combo1,Label6,Label16,Label10,Label20,Label26)
'Create By Sindy 2014/7/22 參考:代理人通知修正(frm04010508_3)改寫
Option Explicit

Dim autonum As String
Dim m_PA09 As String '申請國家
Dim cp() As String
Dim m_blnFirstShow As Boolean '判斷畫面是否第一次顯示
'若承辦人是王協理且未發文則要發EMail通知
Dim stCP09 As String, stCP14 As String ', stCP27 As String
Dim m_936CP09 As String '901內部收文之總收文號
Dim m_936CP12 As String '936內部收文之業務區
Dim m_936CP13 As String '936內部收文之智權人員
Dim stCP12 As String, stCP13 As String '最新收文智權人員,業務區
Dim m_bolUpdCP46 As Boolean '是否更新原程序的收達日
Dim m_bolFMP As Boolean
Dim strCP07 As String '相關號法限
Dim str995NP09 As String '指定提申日或最終提申 加入最終提申日P-108231
Dim m_CP44 As String, m_CP45 As String 'Add By Sindy 2015/1/16
Dim m_Subject As String 'Added by Morgan 2016/6/16


Private Sub cmkok_Click(Index As Integer)
   Select Case Index
      Case 0
         If Text2.Text = "" Then MsgBox "本所期限不可為空值", vbInformation: Text2.SetFocus: Exit Sub
         '不可大於相關號法限
         If strCP07 <> "" And Val(Text2) > Val(TransDate(strCP07, 1)) Then
            MsgBox "本所期限不可大於相關號法定期限【" & TransDate(strCP07, 1) & "】!!", vbInformation: Text2.SetFocus: Exit Sub
         End If
         '比照法限控制
         If Text2 <> "" Then
            If str995NP09 <> "" And Val(Text2) > Val(TransDate(str995NP09, 1)) Then
               MsgBox "本所期限不可大於指定提申或最終提申期限【" & TransDate(str995NP09, 1) & "】!!", vbInformation: Text2.SetFocus: Exit Sub
            End If
         End If
         If Text7 <> "" Then
            '不可大於相關號法限
            If strCP07 <> "" And Val(Text7) > Val(TransDate(strCP07, 1)) Then
               MsgBox "法定期限不可大於相關號法定期限【" & TransDate(strCP07, 1) & "】!!", vbInformation: Text7.SetFocus: Exit Sub
            End If
            If str995NP09 <> "" And Val(Text7) > Val(TransDate(str995NP09, 1)) Then
               MsgBox "法定期限不可大於指定提申或最終提申期限【" & TransDate(str995NP09, 1) & "】!!", vbInformation: Text7.SetFocus: Exit Sub
            End If
            If Val(Text7.Text) < Val(Text2.Text) Then MsgBox "法定期限不可小於本所期限", vbInformation: Text7.SetFocus: Exit Sub
         End If
         If Text8.Enabled = True Then
            If Text8.Text = "" Then MsgBox "承辦期限不可為空值", vbInformation: Text8.SetFocus: Exit Sub
            If Text8.Text > Text2.Text Then MsgBox "承辦期限不可大於本所期限", vbInformation: Text8.SetFocus: Exit Sub
         End If
         '考慮程序新人
         If PUB_GetST05(Text4.Text) = "75" And Text5.Text <> "N" Then MsgBox "當承辦人為程序人員時必需為N": Text5.SetFocus: Exit Sub
         If Text5.Text <> "" And Text5.Text <> "N" Then MsgBox "輸入的數值不正確", vbInformation: Text5.SetFocus: Exit Sub
         
         '檢查本所期限
         With Me.Text2
            If .Text <> "" Then
               If Val(.Text) + 19110000 < strSrvDate(1) Then
                  MsgBox "本所期限不可小於系統日!!!"
                  .SetFocus
                  .SelStart = 0
                  .SelLength = Len(.Text)
                  Exit Sub
               End If
            End If
         End With
         
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
        
         Screen.MousePointer = vbHourglass
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         '若承辦人是王協理且未發文則要發EMail通知
         'Modify by Amy 2024/07/16 原:71011(王副總) 改李柏翰經理
         If stCP14 = "99050" Then
            Call PUB_SendMail(strUserNum, "99050", stCP09, "分案通知")
         End If
         
         '產生指示信
         EndLetter "02", autonum, "01", strUserNum
'         strExc(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            "VALUES ('02','" & autonum & "','01','" & strUserNum & "','前案發文日','" & Text1 & "')"
'         If Not ClsLawExecSQL(1, strExc) Then
'            MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
'            Exit Sub
'         End If
      
         'Modified by Morgan 2016/6/16
         '指示信電子化
         'NowPrint autonum, "02", "01", False, strUserNum, 0
         If Left(Pub_StrUserSt03, 1) = "F" Then
            NowPrint autonum, "02", "01", False, strUserNum
         Else
            NowPrint autonum, "02", "01", True, strUserNum, , , , , , , , , , , , , autonum
            frm1105_1.m_RecNo = autonum
            frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & ".950.DATA.PDF"
            frm1105_1.m_Subject = m_Subject
            frm1105_1.Show
         End If
         'end 2016/6/16
      
'         If m_bolFMP Then
'            g_PrtForm001.PrintCForm autonum
'         End If
'         If m_936CP09 <> "" Then
'            g_PrtForm001.PrintCForm m_936CP09
'         End If
         
         frm040114_1.Show
         frm040114_1.Clear
         Unload Me
      Case 1
         frm040114_1.Show
         Unload Me
      Case 2
         Unload frm040114_1
         Unload Me
   End Select
End Sub

Private Function FormSave() As Boolean
   
On Error GoTo ErrorHandler
   
   cnnConnection.BeginTrans
   
'   'FMP不上發文日
'   If m_bolFMP Then
'      stCP27 = "NULL"
'      stCP14 = Text4
'   Else
'      stCP27 = strSrvDate(1)
      stCP14 = strUserNum
'   End If
   
   '新增B類收文
   autonum = AutoNo("B", 6)
   'Modify By Sindy 2015/1/16 加存CP44,CP45
   strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05," & _
      "cp09,cp10,cp11,cp12,cp13," & _
      "cp14,cp20,cp26,cp32,cp27,cp43,cp48,cp64,cp44,cp45) values " & _
      "('" & strKey1 & "','" & StrKey2 & "','" & strKey3 & "','" & strKey4 & "'," & strSrvDate(1) & _
      ",'" & autonum & "','950','90','" & stCP12 & "','" & stCP13 & "'" & _
      ",'" & stCP14 & "','N','N','N'," & strSrvDate(1) & ",'" & strKey5 & "'," & CNULL(DBDATE(Text8), True) & "," & CNULL(ChgSQL(Text3.Text)) & _
      "," & CNULL(m_CP44) & "," & CNULL(m_CP45) & ")"
   cnnConnection.Execute strSql, intI
   
   '更新下一程序收達期限
   cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & strKey5 & "' AND NP06 IS NULL AND NP07='997'"
   
   '刪除所點選的收文號(將案件暫緩的案件)之下一程序的998.提申期限
   strSql = "delete from nextprogress where np01='" & strKey5 & "' and np07='998' and (np06 is null or np06='N')"
   cnnConnection.Execute strSql, intI
   
   '回覆委任代理人
   m_936CP09 = AutoNo("B", 6)
   m_936CP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
   m_936CP12 = GetSalesArea(m_936CP13)
   'Modify By Sindy 2015/1/16 加存CP44,CP45
   strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
      "CP09,CP10,CP11,CP12,CP13," & _
      "CP14,CP20,CP26,CP32,CP43,cp48,cp44,cp45) VALUES " & _
      "('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & "," & CNULL(DBDATE(Text2), True) & "," & CNULL(DBDATE(Text7), True) & _
      ",'" & m_936CP09 & "','936','90'," & CNULL(m_936CP12) & "," & CNULL(m_936CP13) & _
      ",'" & Text4 & "','N','N','N','" & autonum & "'," & CNULL(DBDATE(Text8), True) & "," & CNULL(m_CP44) & "," & CNULL(m_CP45) & ") "
   cnnConnection.Execute strSql, intI
   
   If Text8 = "" Then
      strSql = "update engineerprogress set ep06=" & strSrvDate(1) & " where ep02='" & m_936CP09 & "'"
      cnnConnection.Execute strSql, intI
   End If
   
   'Added by Morgan 2016/6/16
   strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
   m_Subject = "煩請　貴公司暫緩 " & strExc(1) & " 案之提申，並回覆收達。"
   If ExistCheck("AppForm", "AF01", autonum, "", False) = False Then
      'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
      strExc(2) = PUB_GetLetterJudgeNew("2", cp(1), "950", m_PA09, cp(10))
      PUB_AddAppForm autonum, True, strExc(2), m_Subject  '不轉檔,自行判發
   End If
   'end 2016/6/16
   
   cnnConnection.CommitTrans
   Screen.MousePointer = vbDefault
   FormSave = True
   Exit Function
   
ErrorHandler:
   cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
End Function

Private Sub Form_Activate()
   If m_blnFirstShow = False Then
      Exit Sub
   Else
      m_blnFirstShow = False
   End If
   
   cp(1) = strKey1
   cp(2) = StrKey2
   cp(3) = strKey3
   cp(4) = strKey4
   
   stCP13 = PUB_GetAKindSalesNo(strKey1, StrKey2, strKey3, strKey4)
   stCP12 = GetSalesArea(stCP13)
   
   '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
   strExc(1) = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人"
   
   'Modify By Sindy 2015/1/16 +,cp44,cp45
   If strKey1 = "P" Then
      strExc(0) = "SELECT PA22,cu04,CP10,CP14,CP45,CP13,PA09," & strExc(1) & ", CP06,CP07,CP64,CP48,CP46,CP12,CP07,cp27,cp44,cp45" & _
                  " FROM CASEPROGRESS,PATENT,fagent,customer,SystemKind " & _
                  " WHERE CP09='" & strKey5 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and substr(cp44,1,8)=fa01(+) and substr(cp44,9,1)=fa02(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) AND CP01=SK01(+) "
   ElseIf strKey1 = "PS" Then
      strExc(0) = "SELECT '',cu04,CP10,CP14,CP45,CP13,SP09," & strExc(1) & ", CP06,CP07,CP64,CP48,CP46,CP12,CP07,cp27,cp44,cp45" & _
                  " FROM CASEPROGRESS,SERVICEPRACTICE,fagent,customer,SystemKind " & _
                  " WHERE CP09='" & strKey5 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) and substr(cp44,1,8)=fa01(+) and substr(cp44,9,1)=fa02(+) and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) AND CP01=SK01(+) "
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         strCP07 = "" & .Fields("cp07")
         '是否更新原程序的收達日
         If IsNull(.Fields(12).Value) Then
            m_bolUpdCP46 = True
         Else
            m_bolUpdCP46 = False
         End If
         '專利號數
         Label14.Caption = "" & .Fields(0).Value
         '本所案號
         Label2.Caption = strKey1 + "-" + StrKey2 + "-" + strKey3 + "-" + strKey4
         '申請人
         Label6.Caption = "" & .Fields(1).Value
         '案件性質
         If Not IsNull(.Fields(2)) And Not IsNull(.Fields(6).Value) Then
            Label8.Caption = casenum(.Fields(2).Value, .Fields(6).Value)
         End If
         '承辦人
         If Not IsNull(.Fields(3)) Then
            Label10.Caption = GetStaffName(.Fields(3).Value, True)
         End If
         '智權人員
         If Not IsNull(.Fields(5)) Then
            Label16.Caption = GetStaffName(.Fields(5).Value, True)
         End If
         '申請國家
         If Not IsNull(.Fields(6)) Then
            Label18.Caption = GetNationName(.Fields(6).Value)
         End If
         '取得申請國家代號
         m_PA09 = "" & .Fields(6).Value
         '預設是否算案件數為"N"
         If m_PA09 <> 台灣國家代號 Then
             Me.Text5.Text = "N"
         End If
         '代理人
         Label20.Caption = "" & .Fields(7).Value
         m_CP44 = "" & .Fields("CP44").Value 'Add By Sindy 2015/1/16
         m_CP45 = "" & .Fields("CP45").Value 'Add By Sindy 2015/1/16
         
         If Left(stCP12, 1) = "F" Then
            m_bolFMP = True
         Else
            m_bolFMP = False
         End If
         
         If m_bolFMP Then
            'Modified by Morgan 2017/10/11 FMP預設承辦人比照FCP
            'Text4.Text = PUB_GetFmpCP14(cp)
            Text4.Text = PUB_GetFCPPromoterNo(strKey5, "936", "" & .Fields("CP14"))
            'end 2017/10/11
         Else
            If GetStaffDepartment("" & .Fields(3)) <> "P12" Then
               Text4.Text = .Fields(3).Value
            Else
               '點申請程序時要檢查若有國內案帶國內案的承辦人 --郭
               If InStr(CaseMapIn, "" & .Fields("CP10")) > 0 Then
                  Text4.Text = PUB_GetInCaseCP14(strKey1, StrKey2, strKey3, strKey4)
               End If
               If Text4.Text = "" Then
                  Text4.Text = .Fields(3).Value
               End If
            End If
         End If
         
         '7個日曆天
         If m_bolFMP Then
            strExc(1) = CompDate(2, 7, strSrvDate(1))
         '3個工作天
         Else
            strExc(1) = Pub_GetHandleDay(strKey1, m_PA09, "902")
         End If
      
         If strExc(1) <> "" Then
            Text8 = TransDate(strExc(1), 1)
            '不可大於相關號法限
            If strCP07 <> "" And Val(Text8) > Val(TransDate(strCP07, 1)) Then
               Text8 = TransDate(strCP07, 1)
            End If
            
            '指定提申日
            '加入最終提申日P-108231(996)
            strExc(0) = "select np09 from nextprogress where np01='" & strKey5 & "' and np06 is null and (np07='995' or np07='996')"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               str995NP09 = RsTemp.Fields(0)
               If DBDATE(Text8) > str995NP09 Then
                  Text8 = TransDate(str995NP09, 1)
               End If
            End If
                        
            '不必再預設法定期限--郭
            Text2.Text = Text8.Text
            
            '新規則承辦期限隔日凌晨算
            If Not PUB_IfSetCP48(strKey5) Then
               Text8.Text = ""
               Text8.Enabled = False
            End If
         End If
      End With
   End If
   
   If Text4.Text <> "" Then
      Label26.Caption = GetStaffName(Text4.Text, True)
   End If
   If IsNull(strKey6) Then
      Combo1.AddItem "中: ", 0
      Combo1.Text = "中: "
   Else
      Combo1.AddItem "中: " + strKey6, 0
      Combo1.Text = "中: " + strKey6
   End If
   If IsNull(strKey7) Then
      Combo1.AddItem "英: ", 1
   Else
      Combo1.AddItem "英: " + strKey7, 1
   End If
   If IsNull(strKey8) Then
      Combo1.AddItem "日: ", 2
   Else
      Combo1.AddItem "日: " + strKey8, 2
   End If
   Label3.Caption = strKey5
End Sub

'取案件性質名
Private Function casenum(NUM As String, CON As String)
   If CON = "000" Then
      casenum = GetCaseTypeName(strKey1, NUM, 0)
   Else
      casenum = GetCaseTypeName(strKey1, NUM, 1)
   End If
End Function

Private Sub Form_Initialize()
   ReDim cp(TF_CP) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   m_blnFirstShow = True
   'Add by Amy 2014/09/17 承辦人期限隱藏
   Label29.Visible = False
   Text8.Enabled = False
   Text8.Visible = False
   'end 2014/09/17
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Set frm040114_2 = Nothing 'Removed by Morgan 2021/12/22 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2.Text <> "" Then
      If CheckIsTaiwanDate(Text2.Text) = False Then
         Text2.SetFocus
         Text2_GotFocus
         Cancel = True
         Exit Sub
      Else
         Cancel = False
      End If
   End If
   '若有輸入本所期限, 則不可小於系統日期
   If Me.Text2.Text <> "" Then
      If Val(Me.Text2.Text) + 19110000 < strSrvDate(1) Then
          MsgBox "本所期限不可小於系統日期!!!", vbExclamation
          Text2.SetFocus
          Text2_GotFocus
          Cancel = True
      '若本所期限非工作天則直接調整至最近的工作天
      Else
          Me.Text2.Text = TransDate(PUB_GetWorkDay1(Me.Text2.Text, True), 1)
      End If
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   Dim strTempName As String
   
   If Text4.Text <> "" Then
      '需判斷員工是否離職
      Cancel = Not ClsPDGetStaff(Text4.Text, strTempName)
      Label26 = strTempName
      If Cancel = False Then
         If PUB_GetST05(strUserNum) = "75" Then
            Text5.Text = "N"
         End If
      End If
   End If
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5.Text <> "" And Text5.Text <> "N" Then
      MsgBox "輸入只能空白或 N"
      Text5.SetFocus
      Text5_GotFocus
      Cancel = True
   End If
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If Text7 <> "" Then
      If CheckIsTaiwanDate(Text7.Text) = False Then
         Text7.SetFocus
         Text7_GotFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Text8.Text <> "" Then
      If CheckIsTaiwanDate(Text8.Text) = False Then
         Text8_GotFocus
         Cancel = True
      Else
         '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
         If Len(Me.Text2.Text) > 0 And Len(Me.Text8.Text) > 0 Then
            If Val(Me.Text2.Text) < Val(Me.Text8.Text) Then
               MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
               Text8_GotFocus
               Cancel = True
               Exit Sub
            End If
         End If
      End If
   End If
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
   
   TxtValidate = False
   'Added by Morgan 2021/12/22 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/22
   
   If Me.Text2.Enabled = True Then
      Cancel = False
      Text2_Validate Cancel
      If Cancel = True Then
         Me.Text2.SetFocus
         Text2_GotFocus
         Exit Function
      End If
   End If
   
   If Me.Text4.Enabled = True Then
      Cancel = False
      Text4_Validate Cancel
      If Cancel = True Then
         Me.Text4.SetFocus
         Text4_GotFocus
         Exit Function
      End If
   End If
   
   If Me.Text5.Enabled = True Then
      Cancel = False
      Text5_Validate Cancel
      If Cancel = True Then
         Me.Text5.SetFocus
         Text5_GotFocus
         Exit Function
      End If
   End If
   
   If Me.Text7.Enabled = True Then
      Cancel = False
      Text7_Validate Cancel
      If Cancel = True Then
         Me.Text7.SetFocus
         Text7_GotFocus
         Exit Function
      End If
   End If
   
   If Text8.Enabled = True Then
      Cancel = False
      Text8_Validate Cancel
      If Cancel = True Then
         Me.Text8.SetFocus
         Text8_GotFocus
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function
