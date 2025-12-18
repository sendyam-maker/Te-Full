VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010508_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人通知修正"
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
      Height          =   400
      Index           =   2
      Left            =   6996
      TabIndex        =   37
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmkok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   5772
      TabIndex        =   35
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmkok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4944
      TabIndex        =   34
      Top             =   70
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   60
      TabIndex        =   17
      Top             =   2910
      Width           =   7695
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   30
         Top             =   1710
         Width           =   3285
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Left            =   4680
         MaxLength       =   7
         TabIndex        =   28
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   4680
         MaxLength       =   7
         TabIndex        =   25
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   29
         Top             =   1380
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   1170
         TabIndex        =   27
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   24
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   23
         Top             =   180
         Width           =   1575
      End
      Begin MSForms.TextBox Text3 
         Height          =   300
         Left            =   1200
         TabIndex        =   26
         Top             =   780
         Width           =   5535
         VariousPropertyBits=   671107099
         Size            =   "9763;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         Caption         =   "彼所案號："
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "承辦期限："
         Height          =   255
         Left            =   3720
         TabIndex        =   33
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "法定期限："
         Height          =   255
         Left            =   3720
         TabIndex        =   32
         Top             =   480
         Width           =   975
      End
      Begin MSForms.Label Label26 
         Height          =   255
         Left            =   2400
         TabIndex        =   31
         Top             =   1080
         Width           =   735
         VariousPropertyBits=   27
         Size            =   "1296;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label25 
         Caption         =   "是否算案件數：                 (N：不算)"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1380
         Width           =   3015
      End
      Begin VB.Label Label24 
         Caption         =   "承辦人："
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label23 
         Caption         =   "進度備註："
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "本所期限："
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "來函收文日："
         Height          =   252
         Left            =   120
         TabIndex        =   18
         Top             =   180
         Width           =   1092
      End
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1200
      TabIndex        =   36
      Top             =   1080
      Width           =   5535
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "9763;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   975
   End
   Begin MSForms.Label Label20 
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   2520
      Width           =   5895
      VariousPropertyBits=   27
      Size            =   "10398;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label19 
      Caption         =   "代理人："
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   735
   End
   Begin MSForms.Label Label18 
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   2160
      Width           =   1815
      VariousPropertyBits=   27
      Size            =   "3201;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      Caption         =   "申請國家："
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   2160
      Width           =   975
   End
   Begin MSForms.Label Label16 
      Height          =   255
      Left            =   4920
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
      Left            =   4080
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
      Left            =   3960
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "承辦人："
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin MSForms.Label Label10 
      Height          =   255
      Left            =   960
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
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   975
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
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   735
   End
   Begin MSForms.Label Label6 
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   6735
      VariousPropertyBits=   27
      Size            =   "11880;450"
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
Attribute VB_Name = "frm04010508_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (Combo1,Label6,Label16,Label10,Label18,Label20,Label26,Text3)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim autonum As String
Dim m_strSaleNo As String '智權人員
Dim strProgressNo As String '下一程序之序號
Dim m_PA09 As String '申請國家
'edit by nickc 2007/02/02
'Dim cp(T_CP) As String
Dim cp() As String

Dim strAutoNumber As String
Dim m_blnFirstShow As Boolean '判斷畫面是否第一次顯示
'若承辦人是王協理且未發文則要發EMail通知
Dim stCP09 As String, stCP14 As String, stCP27 As String
'Add by Morgan 2006/6/26
'Modify by Morgan 2009/12/21 改為936
'Dim m_901CP09 As String '901內部收文之總收文號
'Dim m_901CP12 As String '901內部收文之業務區
'Dim m_901CP13 As String '901內部收文之智權人員
Dim m_936CP09 As String '901內部收文之總收文號
Dim m_936CP12 As String '936內部收文之業務區
Dim m_936CP13 As String '936內部收文之智權人員
'end 2009/12/21
Dim stCP12 As String, stCP13 As String '最新收文智權人員,業務區
Dim m_bolUpdCP46 As Boolean '是否更新原程序的收達日
Dim m_bolFMP As Boolean 'Add by Morgan 2009/12/21
Dim strCP07 As String '相關號法限 Add by Morgan 2010/10/26
Dim str995NP09 As String '指定提申日或最終提申日　Added by Morgan 2012/9/7  modify by sonia 2014/5/8 加入最終提申日P-108231
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2016/10/7 END
Dim m_AltrMsgPath As String 'Added by Morgan 2023/9/6
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/10/31 是否為寰華案

'Add By Sindy 2022/7/1
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmkok_Click(Index As Integer)
Dim mRCno As String, mCCno As String, oSubject As String, oContext As String    'Add by Lydia 2014/10/16 FMP案會列印 C類接洽單, 請同時E-MAIL給畫面上之承辦人, 副本發給該員之工程師組別主管.

   Select Case Index
      Case 0
                     
         If Text1.Text = "" Then MsgBox "來函收文日不可為空值", vbInformation: Text1.SetFocus: Exit Sub
         'If Text6.Text = "" Then MsgBox "下一程序不可為空值", vbInformation: Text6.SetFocus: Exit Sub 'Remove by Morgan 2009/12/21
         If Text2.Text = "" Then MsgBox "本所期限不可為空值", vbInformation: Text2.SetFocus: Exit Sub
         'Add by Morgan 2010/10/26 不可大於相關號法限
         If strCP07 <> "" And Val(Text2) > Val(TransDate(strCP07, 1)) Then
            MsgBox "本所期限不可大於相關號法定期限【" & TransDate(strCP07, 1) & "】!!", vbInformation: Text2.SetFocus: Exit Sub
         End If
         
         'Added by Morgan 2014/2/19
         '比照法限控制
         If Text2 <> "" Then
            If str995NP09 <> "" And Val(Text2) > Val(TransDate(str995NP09, 1)) Then
               MsgBox "本所期限不可大於指定提申或最終提申期限【" & TransDate(str995NP09, 1) & "】!!", vbInformation: Text2.SetFocus: Exit Sub
            End If
         End If
         'end 2014/2/19
            
         'If Text7.Text = "" Then MsgBox "法定期限不可為空值", vbInformation: Text7.SetFocus: Exit Sub 'Remove by Morgan 2009/12/21
         If Text7 <> "" Then
            'Add by Morgan 2010/10/26 不可大於相關號法限
            If strCP07 <> "" And Val(Text7) > Val(TransDate(strCP07, 1)) Then
               MsgBox "法定期限不可大於相關號法定期限【" & TransDate(strCP07, 1) & "】!!", vbInformation: Text7.SetFocus: Exit Sub
            End If
            'Added by Morgan 2012/9/7
            If str995NP09 <> "" And Val(Text7) > Val(TransDate(str995NP09, 1)) Then
               MsgBox "法定期限不可大於指定提申或最終提申期限【" & TransDate(str995NP09, 1) & "】!!", vbInformation: Text7.SetFocus: Exit Sub
            End If
            'end 2012/9/7
            
            If Val(Text7.Text) < Val(Text2.Text) Then MsgBox "法定期限不可小於本所期限", vbInformation: Text7.SetFocus: Exit Sub
         End If
         If Text8.Enabled = True Then
            If Text8.Text = "" Then MsgBox "承辦期限不可為空值", vbInformation: Text8.SetFocus: Exit Sub
            If Text8.Text > Text2.Text Then MsgBox "承辦期限不可大於本所期限", vbInformation: Text8.SetFocus: Exit Sub
         End If
         'Modified by Morgan 2013/10/23 考慮程序新人
         'If (Text4.Text = "81002" Or Text4.Text = "73017") And Text5.Text <> "N" Then MsgBox "當承辦人為81002和73017時必需為N": Text5.SetFocus: Exit Sub
         If PUB_GetST05(Text4.Text) = "75" And Text5.Text <> "N" Then MsgBox "當承辦人為程序人員時必需為N": Text5.SetFocus: Exit Sub
         'end 2013/10/23
         
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
         
         'Add By Sindy 2020/7/20
         If m_strIR01 <> "" Then
            '下載信件檔
            If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, "", , , True) = False Then
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            'Add By Sindy 2022/7/21
            If Left(Pub_StrUserSt03, 2) = "F2" Then
               If PUB_ChkFileOpening2(m_PrevForm.m_strFullFileName, "後續才能一併歸卷！") = True Then
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            '2022/7/21 END
         End If
         '2020/7/20 END
         
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         Mail2Eng 'Added by Morgan 2023/9/6
         
         'Add by Morgan 2004/2/18
         '若承辦人是王協理且未發文則要發EMail通知
         'Modify by Amy 2024/07/16 原:71011(王副總) 改李柏翰經理
         If stCP14 = "99050" Then
             Call PUB_SendMail(strUserNum, "99050", stCP09, "分案通知")
         End If
         
'Modify by Morgan 2009/12/21 改一律內部收文936
'         '有輸入下一程序
'         If IsNull(Text6.Text) = False Then
'            '加印新增B類收文號
'            g_PrtForm001.PrintForm strProgressNo, strKey1, strKey2, strKey3, strKey4, cp(9)
'            If Left(cp(12), 1) = "F" Then
'               bol901 = True
'               g_PrtForm001.PrintForm strProgressNo, cp(1), cp(2), cp(3), cp(4), m_901CP09
'               bol901 = False
'            End If
'         End If
         If m_bolFMP Then
            'g_PrtForm001.PrintCForm autonum 'Removed by Morgan 2022/10/6 取消--品薇

            'Add by Lydia 2014/10/16 FMP案會列印 C類接洽單, 請同時E-MAIL給畫面上之承辦人, 副本發給該員之工程師組別主管.
            'Modified by Lydia 2020/08/24 改用模組
            'strExc(0) = "SELECT ST01,ST04,decode(ST16,'1','T','2','R','3','S','4','T1','') mst16 FROM STAFF WHERE ST01='" & Trim(Text4.Text) & "' "
            'intI = 1
            'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            'If intI = 1 Then
            '  strExc(0) = "" & RsTemp.Fields("mst16")
            '  mCCno = Pub_GetSpecMan(strExc(0))
            '  mRCno = RsTemp.Fields("ST01")
            '  If mCCno = mRCno Then mCCno = "" '承辦人已是主管則不必再發副本
            'End If
            mRCno = Trim(Text4.Text)
            mCCno = PUB_GetFCPEngSup(mRCno)
            If mCCno = mRCno Then mCCno = ""
            'end 2020/08/24
            strExc(0) = "SELECT NVL(PA05,NVL(PA06,PA07)) pa05,nvl(FA05||' '||FA63,'') as faname1, nvl(FA04,'') as faname2, nvl(FA06,'') as faname3,CP48,NP23 " & _
                        "FROM PATENT,FAGENT,caseprogress,nextprogress WHERE substr(PA75,1,8)=FA01(+) And substr(PA75,9,1)=FA02(+) and CP01(+)=PA01 And CP02(+)=PA02 And CP03(+)=PA03 And CP04(+)=PA04 " & _
                        "And CP09 = '" & autonum & "' and np01(+)=cp09 and np06(+) is null and PA01='" & cp(1) & "' and PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strExc(0) = "" & RsTemp.Fields("PA05")
               strExc(1) = "" & RsTemp.Fields("faname1")
               strExc(2) = "" & RsTemp.Fields("faname2")
               strExc(3) = "" & RsTemp.Fields("faname3")
               strExc(4) = "" & RsTemp.Fields("CP48") '承辦期限
               strExc(5) = "" & RsTemp.Fields("NP23") '約定期限
            End If
            
            If Len(strExc(1)) > 0 Then '代理人名稱(英->中->日)
               strExc(1) = "代理人　：" & strExc(1)
            ElseIf Len(strExc(2)) > 0 Then
               strExc(1) = "代理人　：" & strExc(2)
            Else
               strExc(1) = "代理人　：" & strExc(3)
            End If
            
            oSubject = cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)
            '發E-Mail通知承辦人
            oContext = "本所案號：" & oSubject & "　　" & vbTab & vbTab & "來函收文日：" & ChangeTStringToTDateString(Trim(Text1)) & vbCrLf & _
                       "專利名稱：" & strExc(0) & vbCrLf & _
                          strExc(1) & vbCrLf & _
                       "承辦人　：" & Trim(Label26) & vbCrLf & _
                       "本所期限：" & ChangeTStringToTDateString(Trim(Text2)) & "　　　　" & vbTab & vbTab & "法定期限：" & ChangeTStringToTDateString(Trim(Text7)) & vbCrLf & _
                       "承辦期限：" & IIf(Len(strExc(4)) > 0, ChangeWStringToTDateString(strExc(4)), "　　　　") & "　　　　" & vbTab & vbTab & "來函性質：代理人通知修正" & vbCrLf & _
                       "約定期限：" & ChangeWStringToTDateString(strExc(5)) & vbCrLf
                       
             oSubject = oSubject & "　收文-代理人通知修正，請自行去調卷處取卷，謝謝！"
             
            PUB_SendMail strUserNum, mRCno, "", oSubject, oContext, "", "", , , , mCCno, "", "", ""
            'end Lydia 2014/10/16
            
         End If
         
         'Added by Lydia 2023/10/04 FMP案待客戶最終指示相關控管
         If m_bolFMP2 = True Then 'Added by Lydia 2023/10/31 只限寰華案要通知
            If PUB_ChkFMP970mail("1", cp(1), cp(2), cp(3), cp(4)) = True Then
            End If
         End If 'end 2023/10/31
         'end 2023/10/04
         
         'Modified by Morgan 2016/7/21 非臺灣案電子化,非FMP不必印B類接洽單
         'If m_936CP09 <> "" Then
         If m_936CP09 <> "" And m_bolFMP Then
            'g_PrtForm001.PrintCForm m_936CP09 'Removed by Morgan 2022/10/6 取消--品薇
         End If
'end 2009/12/21

         strKey1 = "1"
         StrKey2 = ""
         strKey3 = ""
         strKey4 = ""
         strKey5 = ""
         strKey6 = ""
         strKey7 = ""
         strKey8 = ""
         
         'Add By Sindy 2016/10/7
         If Me.m_strIR01 <> "" Then
            Unload frm04010508_1
            Unload frm04010508_2
            Unload Me
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         Else
         '2016/10/7 END
            Unload frm04010508_2
            frm04010508_1.Show
            Unload Me
         End If
      Case 1
         frm04010508_2.Show
         Unload Me
      Case 2
         Unload frm04010508_2
         Unload frm04010508_1
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
   
   If m_blnFirstShow = False Then
      Exit Sub
   Else
      m_blnFirstShow = False
   End If
   
   'Add by Morgan 2009/12/21
   cp(1) = strKey1
   cp(2) = StrKey2
   cp(3) = strKey3
   cp(4) = strKey4
   'end 2009/12/21

   stCP13 = PUB_GetAKindSalesNo(strKey1, StrKey2, strKey3, strKey4)
   stCP12 = GetSalesArea(stCP13)
   
   '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
   strExc(1) = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人"
   
   If strKey1 = "P" Then
      strExc(0) = "SELECT PA22,cu04,CP10,CP14,CP45,CP13,PA09," & strExc(1) & ", CP06,CP07,CP64,CP48,CP46,CP12,CP07 " & _
                  " FROM CASEPROGRESS,PATENT,fagent,customer,SystemKind " & _
                  " WHERE CP09='" & strKey5 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and substr(cp44,1,8)=fa01(+) and substr(cp44,9,1)=fa02(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) AND CP01=SK01(+) "
   ElseIf strKey1 = "PS" Then
      strExc(0) = "SELECT '',cu04,CP10,CP14,CP45,CP13,SP09," & strExc(1) & ", CP06,CP07,CP64,CP48,CP46,CP12,CP07 " & _
                  " FROM CASEPROGRESS,SERVICEPRACTICE,fagent,customer,SystemKind " & _
                  " WHERE CP09='" & strKey5 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) and substr(cp44,1,8)=fa01(+) and substr(cp44,9,1)=fa02(+) and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) AND CP01=SK01(+) "
   End If
   
   m_strSaleNo = ""
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         strCP07 = "" & .Fields("cp07") 'Add by Morgan 2010/10/26
         '是否更新原程序的收達日
         If IsNull(.Fields(12).Value) Then
            m_bolUpdCP46 = True
         Else
            m_bolUpdCP46 = False
         End If
         '專利號數
         Label14.Caption = "" & .Fields(0).Value
         '來函收文日
         'Modify By Sindy 2018/1/2
         If m_strIR01 <> "" Then
            Text1.Text = frm04010508_1.m_RDate
         Else
         '2018/1/2 END
            Text1.Text = strSrvDate(2)
         End If
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
         '彼所案號
         Me.Text9.Text = "" & .Fields(4).Value
         '智權人員
         If Not IsNull(.Fields(5)) Then
            m_strSaleNo = .Fields(5)
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
         
         'Add by Morgan 2009/12/21
         'Modified by Lydia 2023/10/31 判斷FMP案和寰華案
         'If Left(stCP12, 1) = "F" Then
         '   m_bolFMP = True
         'Else
         '   m_bolFMP = False
         'End If
         'end 2009/12/21
         If Left(stCP12, 1) = "F" And m_PA09 <> 台灣國家代號 Then
            m_bolFMP = True
         Else
            m_bolFMP = False
         End If
         m_bolFMP2 = False
         If m_bolFMP = True Then '判斷寰華案
            m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, strKey1, StrKey2, strKey3, strKey4)
         End If
         'end 2023/10/31
         
         'Add by Morgan 2006/8/1 國外部收文承辦人預設黃得峻78063 -- 郭
         If m_bolFMP Then
            '2008/12/3 MODIFY BY SONIA  依FC代理人國籍抓預設承辦人
            'Text4.Text = "85030"         '2008/2/5 MODIFY BY SONIA 78063離職改85030阮威立--郭
            'Modify by Morgan 2009/12/21 改用新規則
            'Text4.Text = PUB_GetFMCASECP14(strKey1, strKey2, strKey3, strKey4)
            'Modified by Morgan 2017/10/11 FMP預設承辦人比照FCP
            'Text4.Text = PUB_GetFmpCP14(cp)
            Text4.Text = PUB_GetFCPPromoterNo(strKey5, "1224", "" & .Fields("cp14"))
            'end 2017/10/11
         Else
            'Modify by Morgan 2007/7/17 加判斷承辦是程序才要 -- 郭
            If GetStaffDepartment("" & .Fields(3)) <> "P12" Then
               Text4.Text = .Fields(3).Value
            Else
            'end 2007/7/17
               'Add by Morgan 2006/8/2 點申請程序時要檢查若有國內案帶國內案的承辦人 --郭
               If InStr(CaseMapIn, "" & .Fields("CP10")) > 0 Then
                  Text4.Text = PUB_GetInCaseCP14(strKey1, StrKey2, strKey3, strKey4)
               End If
               If Text4.Text = "" Then
                  Text4.Text = .Fields(3).Value
               End If
            End If
         End If
         
         'Add by Morgan 2010/1/4
         '7個日曆天
         If m_bolFMP Then
            strExc(1) = CompDate(2, 7, strSrvDate(1))
         '3個工作天
         Else
            strExc(1) = Pub_GetHandleDay(strKey1, m_PA09, "902")
         End If
      
         If strExc(1) <> "" Then
            Text8 = TransDate(strExc(1), 1)
            'Add by Morgan 2010/10/26 不可大於相關號法限
            If strCP07 <> "" And Val(Text8) > Val(TransDate(strCP07, 1)) Then
               Text8 = TransDate(strCP07, 1)
            End If
            
            'Added by Morgan 2012/9/7
            '指定提申日
            'modify by sonia 2014/5/8 加入最終提申日P-108231(996)
            strExc(0) = "select np09 from nextprogress where np01='" & strKey5 & "' and np06 is null and (np07='995' or np07='996')"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               str995NP09 = RsTemp.Fields(0)
               If DBDATE(Text8) > str995NP09 Then
                  Text8 = TransDate(str995NP09, 1)
               End If
            End If
            'end 2012/9/7
                        
            'Text7.Text = Text8.Text 'Removed by Morgan 2013/6/26 不必再預設法定期限--郭
            Text2.Text = Text8.Text
            
            'Add by Morgan 2010/9/29 新規則承辦期限隔日凌晨算
            If Not PUB_IfSetCP48(strKey5) Then
               Text8.Text = ""
               Text8.Enabled = False
            End If
            'end 2010/9/29
            
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
   'Add by Amy 2014/09/18 承辦人期限隱藏
    Label29.Visible = False
    Text8.Enabled = False
    Text8.Visible = False
    'end 2014/09/18

End Sub

Private Sub Form_Initialize()
    'add by nickc 2007/02/02
    ReDim cp(TF_CP) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   m_blnFirstShow = True
   
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010508_2.m_strIR01
   m_strIR02 = frm04010508_2.m_strIR02
   m_strIR03 = frm04010508_2.m_strIR03
   m_strIR04 = frm04010508_2.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2022/7/1
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   '2022/7/1 END
   
   'Set frm04010508_3 = Nothing 'Removed by Morgan 2021/12/20 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 <> "" Then
      If CheckIsTaiwanDate(Text1.Text) = False Then
         Cancel = True
      Else
         Cancel = False
      End If
      'edit by nickc 2007/09/27
      'If Text1.Text > ChangeWDateStringToTString(Date) Then
      If Text1.Text > strSrvDate(2) Then
         MsgBox "輸入的日期不可大於系統日"
         Cancel = True
      Else
         Cancel = False
      End If
      If Cancel = True Then
         Text1.SetFocus
         Text1_GotFocus
      End If
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

'Remove by Morgan 2009/12/21 不預設
'Private Sub Text2_LostFocus()
'    If m_PA09 <> 台灣國家代號 Then
'        Text7.Text = Text2.Text
'    End If
'End Sub

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
   'Add by Morgan 2004/7/30
   Dim strTempName As String
   
   If Text4.Text <> "" Then
      'Modify by Morgan 2004/7/30
      '需判斷員工是否離職
      'edit by nickc 2007/02/02 不用 dll 了
      'Cancel = Not objPublicData.GetStaff(Text4.Text, strTempName)
      Cancel = Not ClsPDGetStaff(Text4.Text, strTempName)
      Label26 = strTempName
      If Cancel = False Then
         'Modified by Morgan 2013/10/23 考慮程序新人
         'If Text4.Text = "81002" Or Text4.Text = "73017" Then
         If PUB_GetST05(strUserNum) = "75" Then
         'end 2013/10/23
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

'Remove by Morgan 2009/12/21
'Private Sub Text6_GotFocus()
'   TextInverse Text6
'End Sub
'
'Private Sub Text6_LostFocus()
'   If Text6.Text <> "" Then
'      strExc(1) = Pub_GetHandleDay(strKey1, m_PA09, Text6.Text)
'      '2008/12/3 add by sonia FMP案改為7個日曆天
'      If m_bolFMP Then
'         strExc(1) = CompDate(2, 7, strSrvDate(1))
'      End If
'      '2008/12/3 end
'
'      If strExc(1) <> "" Then
'         Text8 = TransDate(strExc(1), 1)
'         Text7.Text = Text8.Text
'         Text2.Text = Text8.Text
'      End If
'   End If
'End Sub
'
'Private Sub Text6_Validate(Cancel As Boolean)
'   If Len(Me.Text6.Text) > 0 Then
'      If Len(Me.Text6.Text) <> 3 Then
'         MsgBox "下一程序欄位值必須為三碼 !", vbInformation
'         Text6.SetFocus
'         Text6_GotFocus
'         Cancel = True
'      Else
'         Label30 = casenum(Text6.Text, m_PA09)
'         If Label30 = "" Then
'            MsgBox "下一程序輸入錯誤", vbInformation
'            Text6.SetFocus
'            Text6_GotFocus
'            Cancel = True
'         End If
'      End If
'   End If
'End Sub
'end2009/12/21

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

'取案件性質名
Private Function casenum(NUM As String, CON As String)
   If CON = "000" Then
      casenum = GetCaseTypeName(strKey1, NUM, 0)
   Else
      casenum = GetCaseTypeName(strKey1, NUM, 1)
   End If
End Function

Private Function FormSave() As Boolean
   Dim stCP14 As String, stCP27 As String 'Add by Morgan 2009/12/21
   
On Error GoTo ErrorHandler
   
   cnnConnection.BeginTrans
   
   'Add by Morgan 2009/12/21 FMP不上發文日
   If m_bolFMP Then
      stCP27 = "NULL"
      stCP14 = Text4
   Else
      stCP27 = strSrvDate(1)
      stCP14 = strUserNum
   End If
   'end 2009/12/21
   
   '新增C類收文
   autonum = AutoNo("C", 6)
   'Modify by Morgan 2009/12/21 案件性質改用代理人通知修正1224(原為通知修正1201)
   'Modified by Morgan 2012/5/25 +CP119
   strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp20,cp26," & _
      "cp32,cp27,cp43,cp48,cp64,cp119) values('" & strKey1 & "','" & StrKey2 & "','" & strKey3 & "','" & strKey4 & "'" & _
      "," & CNULL(DBDATE(Text1), True) & "," & CNULL(DBDATE(Text2), True) & "," & CNULL(DBDATE(Text7), True) & _
      ",'" & autonum & "','1224','" & stCP12 & "','" & stCP13 & "','" & stCP14 & "','N','N','N'," & stCP27 & _
      ",'" & strKey5 & "'," & CNULL(DBDATE(Text8), True) & "," & CNULL(ChgSQL(Text3.Text)) & "," & CNULL(DBDATE(Text1), True) & ")"
   cnnConnection.Execute strSql, intI
   
   strSql = ""
   '原程序上代理人已收達
   If m_bolUpdCP46 = True Then
      strSql = ",CP46 = " & ChangeTStringToWString(Text1)
   End If
   '更新彼所案號
   cnnConnection.Execute "UPDATE CASEPROGRESS SET CP45 ='" & ChgSQL(Text9.Text) & "'" & strSql & " WHERE CP09='" & strKey5 & "'"
   '2008/9/3 ADD BY SONIA 更新相同本所案號之相同代理人的彼所案號，若是彼所案號空的話
   'Modify by Morgan 2009/5/7 欄位抓錯了
   'cnnConnection.Execute "update caseprogress set cp45=" & CNULL(ChgSQL(Text4)) & " where cp09 in (select cp09 from caseprogress where cp45 is null and " & ChgCaseprogress(strKey1 & strKey2 & strKey3 & strKey4) & " and cp09<'C' AND cp44 in (select cp44 from caseprogress where cp09='" & strKey5 & "' ))"
   'Modified by Morgan 2012/2/15 取消 cp09<'C' 條件(C類也會有發文作業,有代理人就要更新彼號,資料才會一致)
   cnnConnection.Execute "update caseprogress set cp45=" & CNULL(ChgSQL(Text9)) & " where cp09 in (select cp09 from caseprogress where rtrim(cp45) is null and " & ChgCaseprogress(strKey1 & StrKey2 & strKey3 & strKey4) & " AND cp44 in (select cp44 from caseprogress where cp09='" & strKey5 & "' ))"
   
   '2011/5/12 add by sonia 更新下一程序收達期限
   cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP06 ='Y' WHERE NP01 ='" & strKey5 & "' AND NP06 IS NULL AND NP07='997' "
   '2011/5/12 END
   
   'Added by Morgan 2016/6/6
   If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
      PUB_AddLetterProgress autonum, 1, False
   End If
   'end 2016/6/6

'Modify by Morgan 2009/12/21 改一律內部收文回覆委任代理人936
'   '有輸入下一程序
'   If IsNull(Text6.Text) = False Then
'      '自動上續辦'Y'
'      strProgressNo = GetNextProgressNo
'      strSQL = "insert into nextprogress(NP01,NP02,NP03,NP04,NP05,NP06,NP07,NP08,NP09,NP10,NP15,NP22) " & _
'         "values('" & autonum & "','" & strKey1 & "','" & strKey2 & "','" & strKey3 & "','" & strKey4 & "','Y','" & _
'         Text6.Text & "','" & ChangeTStringToWString(Text2.Text) & "','" & ChangeTStringToWString(Text7.Text) & _
'         "','" & PUB_GetAKindSalesNo(strKey1, strKey2, strKey3, strKey4) & "','" & Text3.Text & "','" & strProgressNo & "')"
'
'      cnnConnection.Execute strSQL
'
'      '新增B類收文
'      cp(9) = AutoNo("B", 6)
'      cp(1) = strKey1
'      cp(2) = strKey2
'      cp(3) = strKey3
'      cp(4) = strKey4
'      cp(5) = strSrvDate(1)
'      cp(6) = ChangeTStringToWString(Text2.Text)
'      cp(7) = ChangeTStringToWString(Text7.Text)
'      cp(10) = Me.Text6.Text
'      cp(11) = "90"
'      cp(12) = stCP12
'      cp(13) = stCP13
'      cp(14) = Me.Text4.Text '承辦人
'      '是否算案件數
'      cp(26) = Me.Text5.Text
'      cp(43) = autonum
'      cp(48) = ChangeTStringToWString(Text8.Text)
'      If PUB_AddNewCaseProgress(cp) = False Then GoTo ErrorHandler
'
'      'Add by Morgan 2004/2/18
'      '若承辦人是王協理且未發文則要發EMail通知
'      stCP09 = cp(9)
'      stCP14 = cp(14)
'
'      'Add by Morgan 2006/6/26
'      '國外部收文若有期限則自動內部收文901告知代理人,承辦人固定為78063黃得峻並列印內部收文接洽單
'      If m_bolFMP Then
'         m_901CP09 = AutoNo("B", 6)
'         '2008/12/2 modify by sonia 改FMP控管方式
'         'm_901CP13 = PUB_GetFCPSalesNo(cp(1), cp(2), cp(3), cp(4))
'         m_901CP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
'         '2008/12/2 END
'         m_901CP12 = GetSalesArea(m_901CP13)
'
''2008/12/3 modify by sonia FMP案改為畫面之本所,法定,承辦期限,即同上面B類的期限
''         'Modify by Morgan 2007/11/2 設定已改為3日，此處則固定抓7天--郭
''         'strExc(1) = GetWorkDays(strKey1, m_PA09, "901")
''         'If strExc(1) = Empty Then strExc(1) = 7
''         strExc(1) = 7
''         'Add by Morgan 2008/5/26 若來函期限超過(含)3個月則告代的承辦期限為14天--阮威立
''         If Val(strExc(1)) < 14 Then
''            If DBDATE(Text7) >= CompDate(1, 3, strSrvDate(1)) Then
''               strExc(1) = 14
''            End If
''         End If
''         'end 2008/5/26
''         'end 2007/11/2
''         'Modify by Morgan 2006/8/4 不必抓工作天--郭
''         'strExc(0) = CompWorkDay(Val(strExc(1)), strSrvDate(1), 0)
''         strExc(0) = CompDate(2, Val(strExc(1)), strSrvDate(1))
''2008/12/3 END
'
'         '2008/12/3 MODIFY BY SONIA 依FC代理人國籍抓預設承辦人,期限改為畫面之本所,法定,承辦期限,即同上面B類的期限
'         'strSQL = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
'            "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) VALUES " & _
'            "('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & "," & strExc(0) & "," & strExc(0) & _
'            ",'" & m_901CP09 & "','" & 告知代理人 & "','90'," & CNULL(m_901CP12) & "," & CNULL(m_901CP13) & _
'            ",'85030','N','N','N','" & autonum & "'," & strExc(0) & ") "    '2008/2/5 MODIFY BY SONIA 78063離職改85030阮威立--郭
'         strSQL = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
'            "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) VALUES " & _
'            "('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & "," & ChangeTStringToWString(Text2.Text) & "," & ChangeTStringToWString(Text7.Text) & _
'            ",'" & m_901CP09 & "','" & 告知代理人 & "','90'," & CNULL(m_901CP12) & "," & CNULL(m_901CP13) & _
'            "," & CNULL(PUB_GetFMCASECP14(cp(1), cp(2), cp(3), cp(4))) & ",'N','N','N','" & autonum & "'," & ChangeTStringToWString(Text8.Text) & ") "
'         '2008/12/3 END
'         cnnConnection.Execute strSQL
'      End If
'   End If

   m_936CP09 = AutoNo("B", 6)
   'Modified by Morgan 2021/1/28
   'm_936CP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
   'm_936CP12 = GetSalesArea(m_936CP13)
   m_936CP13 = stCP13
   m_936CP12 = stCP12
   'end 2021/1/28
   strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
      "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,cp48) VALUES " & _
      "('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & "," & CNULL(DBDATE(Text2), True) & "," & CNULL(DBDATE(Text7), True) & _
      ",'" & m_936CP09 & "','936','90'," & CNULL(m_936CP12) & "," & CNULL(m_936CP13) & _
      ",'" & Text4 & "','N','N','N','" & autonum & "'," & CNULL(DBDATE(Text8), True) & ") "
   cnnConnection.Execute strSql, intI
   
   'Add by Morgan 2010/9/30
   If Text8 = "" Then
      strSql = "update engineerprogress set ep06=" & strSrvDate(1) & " where ep02='" & m_936CP09 & "'"
      cnnConnection.Execute strSql, intI
   End If
   
   'Add by Sindy 2016/10/7
   If m_strIR01 <> "" Then
      'Add by Sindy 2017/10/23 通知修正信件要歸卷
      '下載信件檔,上傳卷宗區
      'Modify By Sindy 2017/12/22 m_936CP09 == > autonum
      'If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, m_936CP09, "ALTR") = False Then
      'Modified by Morgan 2023/9/6 +m_MsgPath
      If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, autonum, "ALTR", m_AltrMsgPath) = False Then
      '2017/12/22 END
         GoTo ErrorHandler
      End If
      '2017/10/23 END
      
      'Added by Morgan 2023/9/6
      
      'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", autonum, "")
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm04010508_1", IIf(Pub_StrUserSt03 = "F22", autonum, "")
   End If
   '2016/10/7 END
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrorHandler:
   cnnConnection.RollbackTrans
    
End Function

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
   Dim ii As Integer
   Dim Cancel As Boolean
   
   TxtValidate = False
   
   'Added by Morgan 2021/12/20 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/20
   
   If Me.Text1.Enabled = True Then
      Cancel = False
      Text1_Validate Cancel
      If Cancel = True Then
         Me.Text1.SetFocus
         Text1_GotFocus
         Exit Function
      End If
   End If
   
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
   
'Remove by Morgan 2009/12/21
'   If Me.Text6.Enabled = True Then
'      Cancel = False
'      Text6_Validate Cancel
'      If Cancel = True Then
'         Me.Text6.SetFocus
'         Text6_GotFocus
'         Exit Function
'      End If
'   End If
'end 2009/12/21
   
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

'Added by Morgan 2023/9/6
'敏莉說她們不會輸代理人來函，不必排除
Private Sub Mail2Eng()
   Dim stSub As String
   
   If Text4.Text <> "" And m_AltrMsgPath <> "" Then
      stSub = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
      'Modified by Morgan 2024/8/16 cp(9) -> strKey5
      strExc(0) = "select np09 from nextprogress where np01='" & strKey5 & "' and np06 is null and np07='995'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         stSub = "TOP URGENT! " & stSub & "指定於" & ChangeWStringToTDateString(RsTemp(0)) & "提交，現代理人來函通知修正，請詳見附件內容並續行後續回覆。"
      Else
         stSub = stSub & "代理人來函通知修正，請詳見附件內容並續行後續回覆。"
      End If
      PUB_SendMail strUserNum, Text4.Text, "", stSub, "如旨", , m_AltrMsgPath
   End If
End Sub
