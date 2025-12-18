VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140118 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外部人員離職修改資料"
   ClientHeight    =   5730
   ClientLeft      =   660
   ClientTop       =   640
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8950
   Begin VB.CommandButton Command1 
      Caption         =   "檢查資料(&C)"
      Height          =   345
      Index           =   3
      Left            =   4380
      TabIndex        =   1
      Top             =   155
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   1
      Left            =   7092
      TabIndex        =   3
      Top             =   155
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "執行(&O)"
      Height          =   345
      Index           =   0
      Left            =   6210
      TabIndex        =   2
      Top             =   155
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1050
      MaxLength       =   6
      TabIndex        =   0
      Top             =   210
      Width           =   795
   End
   Begin MSForms.Label LblST55 
      Height          =   240
      Left            =   1440
      TabIndex        =   11
      Top             =   1410
      Width           =   1710
      Size            =   "3016;423"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblST54 
      Height          =   240
      Left            =   1440
      TabIndex        =   9
      Top             =   1140
      Width           =   1710
      Size            =   "3016;423"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblST53 
      Height          =   240
      Left            =   1440
      TabIndex        =   7
      Top             =   870
      Width           =   1710
      Size            =   "3016;423"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblST52 
      Height          =   240
      Left            =   1440
      TabIndex        =   5
      Top             =   600
      Width           =   1710
      Size            =   "3016;423"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox List1 
      Height          =   3200
      Left            =   72
      TabIndex        =   20
      Top             =   2450
      Width           =   8800
      VariousPropertyBits=   746586139
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "15522;5644"
      MatchEntry      =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblST02 
      Height          =   300
      Left            =   1920
      TabIndex        =   19
      Top             =   240
      Width           =   1188
      Size            =   "2096;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LblNote 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   1230
      TabIndex        =   18
      Top             =   2070
      Width           =   7695
   End
   Begin VB.Label LblST16 
      Height          =   240
      Left            =   4860
      TabIndex        =   17
      Top             =   870
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國外部組別："
      Height          =   240
      Index           =   7
      Left            =   3750
      TabIndex        =   16
      Top             =   870
      Width           =   1080
   End
   Begin VB.Label LblST03 
      Height          =   240
      Left            =   4860
      TabIndex        =   15
      Top             =   600
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部　　門："
      Height          =   240
      Index           =   6
      Left            =   3750
      TabIndex        =   14
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第二級管制人："
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   13
      Top             =   600
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   240
      Index           =   0
      Left            =   150
      TabIndex        =   12
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第五級管制人："
      Height          =   180
      Index           =   5
      Left            =   150
      TabIndex        =   10
      Top             =   1410
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第四級管制人："
      Height          =   180
      Index           =   4
      Left            =   150
      TabIndex        =   8
      Top             =   1140
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第三級管制人："
      Height          =   180
      Index           =   3
      Left            =   150
      TabIndex        =   6
      Top             =   870
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "資料顯示區："
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   4
      Top             =   2130
      Width           =   1080
   End
End
Attribute VB_Name = "frm140118"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/13 Form2.0已修改(LblST02,List1,LblST52~LblST55)
'Create By Sindy 2018/4/13
Option Explicit

'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bUpdate = IsUserHasRightOfFunction("frm140118", strEdit, False)
   MoveFormToCenter Me
   If m_bUpdate Then
      Command1(0).Visible = True
      Command1(3).Visible = True
   Else
      Command1(0).Visible = False
      Command1(3).Visible = False
   End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim rsTmp As New ADODB.Recordset
Dim strText As String
Dim strUpdDate As String
Dim strSQLCmd As String
Dim strNP02 As String, strNP03 As String, strNP04 As String, strNP05 As String
Dim strNP01 As String, strNP22 As String
Dim strCP13 As String
   
On Error GoTo ErrHand
   Select Case Index
      Case 0 '確定
         If LblST02.Caption = "" Then Exit Sub
         If LblST03.Caption = "" Then
            MsgBox "無部門資料，不可執行！", vbInformation
            Exit Sub
         End If
         '該員工為F1外商人員時,要檢查國外部組別
         If Left(LblST03, 2) = "F1" Then
            If Len(LblST16.Caption) <= 2 Then
               MsgBox "國外部組別不正確，不可執行！", vbInformation
               Exit Sub
            End If
         End If
         
         If MsgBox("確定要執行修改資料嗎？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
         
         strUpdDate = Left(strSrvDate(2), 3) & "年" & Val(Mid(strSrvDate(2), 4, 2)) & "月" & Val(Right(strSrvDate(2), 2)) & "日"
         
         Screen.MousePointer = vbHourglass
         
         On Error GoTo ErrorHandler
         cnnConnection.BeginTrans
         
         '--更改資料:
         '--該員工為F2外專承辦組(F23):
         If Left(LblST03, 3) = "F23" Then
            '--1.國家檔:
            '--所有該員控管的國家資料都要修改,若不同國家改不同人控管則要逐一國家更新
            '  update nation set na51='新FCP承辦業務員' where na51='離職人員員工編號';
            
            '--2.再更新該員工名下的客戶資料, 改業務員為該國籍的FCP承辦業務員(na51)
            '--cu12也要跟著cu13改(以免新舊承辦業務員不同部門),原cu13改到cu129(若cu129有值則加到後面), cu79也要加註(備註加在前面)
            strExc(10) = "update customer set" & _
                         " cu12=(select st15 from nation,staff where cu10=na01 and na51=st01)," & _
                         " cu13=(select na51 from nation where cu10=na01)," & _
                         " CU129=DECODE(CU129,NULL,CU13,CU129||','||CU13)," & _
                         " CU79=(SELECT DECODE(CU79,NULL,'" & strUpdDate & "整批改業務員,原為'||CU13||ST02||'改至開發人員欄'," & _
                                    "'" & strUpdDate & "整批改業務員,原為'||CU13||ST02||'改至開發人員欄;'||CU79)" & _
                         " FROM STAFF WHERE CU13=ST01) where cu13='" & Text1(0).Text & "'"
            cnnConnection.Execute strExc(10)
            '--3.下一程序檔:參考PUB_GetFCPSalesNo方式更新
            '--依系統類別,先看有沒有不是專利或專利服務的案件,再加語法
            '  select np02,count(*) from NEXTPROGRESS where NP06 IS NULL AND NP10='離職人員員工編號' group by np02;
            
            '--專利
            'Modify By Sindy 2019/4/16 摩根:202補文件不加註異動備註,以免產出資料時一併抓出去了
            strExc(10) = "update nextprogress set" & _
                         " np10=(select nvl(fn.na51,cn.na51) from patent,fagent,customer,nation fn,nation cn" & _
                         " where np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+)" & _
                         " and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and fa10=fn.na01(+)" & _
                         " and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and cu10=cn.na01(+))," & _
                         " NP15=(SELECT decode(np07,'202',NP15,DECODE(NP15,NULL,'" & strUpdDate & "整批改業務員,原為'||NP10||ST02," & _
                         " NP15||';" & strUpdDate & "整批改業務員,原為'||NP10||ST02||';')) FROM STAFF WHERE NP10=ST01)" & _
                         " where np06 is null and (np02,np03,np04,np05,np22) in" & _
                         " (select np02,np03,np04,np05,np22 from NEXTPROGRESS,patent where NP06 IS NULL AND NP10='" & Text1(0).Text & "' and np02 IN ('P','CFP','FCP')" & _
                         " and np07 not IN (" & PAnp07NotIn & ")" & _
                         " and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and pa57 is null)"
            cnnConnection.Execute strExc(10)
            '--專利服務
            strExc(10) = "update nextprogress set" & _
                         " np10=(select nvl(fn.na51,cn.na51) from servicepractice,fagent,customer,nation fn,nation cn" & _
                         " where np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+)" & _
                         " and substr(sp26,1,8)=fa01(+) and substr(sp26,9,1)=fa02(+) and fa10=fn.na01(+)" & _
                         " and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) and cu10=cn.na01(+))," & _
                         " NP15=(SELECT DECODE(NP15,NULL,'" & strUpdDate & "整批改業務員,原為'||NP10||ST02," & _
                         " NP15||';" & strUpdDate & "整批改業務員,原為'||NP10||ST02||';') FROM STAFF WHERE NP10=ST01)" & _
                         " where np06 is null and (np02,np03,np04,np05,np22) in" & _
                         " (select np02,np03,np04,np05,np22 from NEXTPROGRESS,servicepractice where NP06 IS NULL AND NP10='" & Text1(0).Text & "' and np02 IN ('PS','CPS','FG')" & _
                         " and np07 not IN (" & PAnp07NotIn & ")" & _
                         " and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+) and sp15 is null)"
            cnnConnection.Execute strExc(10)
            
         '--該員工為F2外專程序組(F22):
         ElseIf Left(LblST03, 3) = "F22" Then
            '--1.國家檔:
            '--所有該員控管的國家資料都要修改,若不同國家改不同人控管則要逐一國家更新
            '  update nation set na16='新FCP程序管制人' where na16='離職人員員工編號';

            '--2.下一程序檔:參考PUB_GetFCPSalesNo方式更新
            '--依系統類別,先看有沒有不是專利或專利服務的案件,再加語法
            '  select np02,count(*) from NEXTPROGRESS where NP06 IS NULL AND NP10='離職人員員工編號' group by np02;
            
            '--專利
            'Modify By Sindy 2019/4/16 摩根:202補文件不加註異動備註,以免產出資料時一併抓出去了
            strExc(10) = "update nextprogress set" & _
                         " np10=(select nvl(fn.na16,cn.na16) from patent,fagent,customer,nation fn,nation cn" & _
                         "      where np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+)" & _
                         "      and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and fa10=fn.na01(+)" & _
                         "      and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and cu10=cn.na01(+))," & _
                         " NP15=(SELECT decode(np07,'202',NP15,DECODE(NP15,NULL,'" & strUpdDate & "整批改管制人,原為'||NP10||ST02," & _
                         "      NP15||';" & strUpdDate & "整批改管制人,原為'||NP10||ST02||';')) FROM STAFF WHERE NP10=ST01)" & _
                         " where np06 is null and (np02,np03,np04,np05,np22) in" & _
                         "     (select np02,np03,np04,np05,np22 from NEXTPROGRESS,patent where NP06 IS NULL AND NP10='" & Text1(0).Text & "' and np02 IN ('P','CFP','FCP')" & _
                         "     and np07 IN (" & PAnp07NotIn & ")" & _
                         "     and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and pa57 is null)"
            cnnConnection.Execute strExc(10)
            '--專利服務
            strExc(10) = "update nextprogress set" & _
                         " np10=(select nvl(fn.na16,cn.na16) from servicepractice,fagent,customer,nation fn,nation cn" & _
                         "      where np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+)" & _
                         "      and substr(sp26,1,8)=fa01(+) and substr(sp26,9,1)=fa02(+) and fa10=fn.na01(+)" & _
                         "      and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) and cu10=cn.na01(+))," & _
                         " NP15=(SELECT DECODE(NP15,NULL,'" & strUpdDate & "整批改管制人,原為'||NP10||ST02," & _
                         "      NP15||';" & strUpdDate & "整批改管制人,原為'||NP10||ST02||';') FROM STAFF WHERE NP10=ST01)" & _
                         " where np06 is null and (np02,np03,np04,np05,np22) in" & _
                         "     (select np02,np03,np04,np05,np22 from NEXTPROGRESS,servicepractice where NP06 IS NULL AND NP10='" & Text1(0).Text & "' and np02 IN ('PS','CPS','FG')" & _
                         "     and np07 IN (" & PAnp07NotIn & ")" & _
                         "     and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+) and sp15 is null)"
            cnnConnection.Execute strExc(10)
         
         '--該員工為F1外商人員:
         ElseIf Left(LblST03, 2) = "F1" Then
            '--1.更新該員工名下的客戶資料, 改業務員為該業務的第二級主管, 第二級離職抓第三級, 依此類推...
            '--cu12也要跟著cu13改(以免新舊承辦業務員不同部門),原cu13改到cu129(若cu129有值則加到後面), cu79也要加註(備註加在前面)
            
            '客戶檔:
            '離職人員為英文組：
            '客戶國籍為001~008者，維持原來更新為離職人員的主管(st52->st53->st54)
            '    國籍非台灣者，改依客戶國籍之國家檔設定之FCT承辦智權人員(NA55)
            '--1.國家檔:
            '--所有該員控管的國家資料都要修改,若不同國家改不同人控管則要逐一國家更新
            '  update nation set na55='新FCT承辦智權人員' where na55='離職人員員工編號';
            
            '離職人員為日文組：維持原來更新為離職人員的主管(st52->st53->st54)
            strSQLCmd = ""
            If InStr(LblST16.Caption, "英") > 0 Then
               '--1.國家檔:
               '--所有該員控管的國家資料都要修改,若不同國家改不同人控管則要逐一國家更新
               '  update nation set NA55='新FCT承辦智權人員' where NA55='離職人員員工編號';
               
               '國籍非台灣者，改依客戶國籍之國家檔設定之FCT承辦智權人員(NA55)
               strExc(10) = "update customer set" & _
                            " cu12=(select st15 from nation,staff where cu10=na01 and NA55=st01)," & _
                            " cu13=(select NA55 from nation where cu10=na01)," & _
                            " CU129=DECODE(CU129,NULL,CU13,CU129||','||CU13)," & _
                            " CU79=(SELECT DECODE(CU79,NULL,'" & strUpdDate & "整批改業務員,原為'||CU13||ST02||'改至開發人員欄'," & _
                                       "'" & strUpdDate & "整批改業務員,原為'||CU13||ST02||'改至開發人員欄;'||CU79)" & _
                            " FROM STAFF WHERE CU13=ST01)" & _
                            " where cu13='" & Text1(0).Text & "' and not (cu10>='000' and cu10<='008')"
               cnnConnection.Execute strExc(10)
               
               strSQLCmd = " and cu10>='000' and cu10<='008'" '台灣客戶
            End If
            strExc(10) = "UPDATE CUSTOMER SET" & _
                       " CU12=(SELECT DECODE(S2.ST04,'1',S2.ST15,DECODE(S3.ST04,'1',S3.ST15,DECODE(S4.ST04,'1',S4.ST15,S5.ST15)))" & _
                       " FROM STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5" & _
                       " WHERE CU13=S1.ST01(+) AND S1.ST52=S2.ST01(+) AND S1.ST53=S3.ST01(+) AND S1.ST54=S4.ST01(+) AND S1.ST55=S5.ST01(+))," & _
                       " CU13=(SELECT DECODE(S2.ST04,'1',S1.ST52,DECODE(S3.ST04,'1',S1.ST53,DECODE(S4.ST04,'1',S1.ST54,S1.ST55)))" & _
                       " FROM STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5" & _
                       " WHERE CU13=S1.ST01(+) AND S1.ST52=S2.ST01(+) AND S1.ST53=S3.ST01(+) AND S1.ST54=S4.ST01(+) AND S1.ST55=S5.ST01(+))," & _
                       " CU129=DECODE(CU129,NULL,CU13,CU129||','||CU13)," & _
                       " CU79=(SELECT DECODE(CU79,NULL,'" & strUpdDate & "整批改業務員,原為'||CU13||ST02||'改至開發人員欄','" & strUpdDate & "整批改業務員,原為'||CU13||ST02||'改至開發人員欄;'||CU79)" & _
                       " FROM STAFF WHERE CU13=ST01)" & _
                       " where cu13='" & Text1(0).Text & "'" & strSQLCmd
            cnnConnection.Execute strExc(10)
            
            '--2.下一程序檔:參考PUB_GetFCTSalesNo方式更新
            '--依系統類別,先看有沒有不是商標或商標服務的案件,再加語法
            '  select np02,count(*) from NEXTPROGRESS where NP06 IS NULL AND NP10='離職人員員工編號' group by np02;
            
            'CFT案之專業部管制的期限strNpSqlOfNoSalesDuty不必改 , 國家檔的Trigger會一併改
            '--CFT案專業部管制的期限,不可更新為同案之其他在職業務,也不可依最後A類收文業務更新,直接更新為其第二級主管,第二級主管離職則為第三級主管,以此類推.....
'            strExc(10) = "update NEXTPROGRESS set" & _
'                         " np10=(SELECT DECODE(S2.ST04,'1',S1.ST52,DECODE(S3.ST04,'1',S1.ST53,DECODE(S4.ST04,'1',S1.ST54,S1.ST55)))" & _
'                         "      FROM STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5" & _
'                         "      WHERE np10=S1.ST01(+) AND S1.ST52=S2.ST01(+) AND S1.ST53=S3.ST01(+) AND S1.ST54=S4.ST01(+) AND S1.ST55=S5.ST01(+))," & _
'                         " NP15=(SELECT NP15||';" & strUpdDate & "整批改業務員,原為'||NP10||ST02 FROM STAFF WHERE NP10=ST01)" & _
'                         " where (np02,np03,np04,np05,np22) in" & _
'                         " (select np02,np03,np04,np05,np22 from nextprogress,trademark" & _
'                         " where NP06 IS NULL AND NP10='" & Text1(0).Text & "' AND NP02='CFT'" & _
'                         " and np07 IN ("& CFTnp07NotIn &")" & _
'                         " and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) and tm29 is null)"
'            cnnConnection.Execute strExc(10)
            
            'CFT非專業部管制的期限
            strExc(0) = "SELECT NP01,NP02,NP03,NP04,NP05,NP22 from NEXTPROGRESS,trademark where NP06 IS NULL AND NP10='" & Text1(0).Text & "' AND NP02='CFT'" & _
                        " and np07 NOT IN (" & TMnp07NotIn & ")" & _
                        " and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) and tm29 is null"
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               rsTmp.MoveFirst
               Do While Not rsTmp.EOF
                  strNP01 = rsTmp.Fields("NP01")
                  strNP02 = rsTmp.Fields("NP02")
                  strNP03 = rsTmp.Fields("NP03")
                  strNP04 = rsTmp.Fields("NP04")
                  strNP05 = rsTmp.Fields("NP05")
                  strNP22 = rsTmp.Fields("NP22")
                  strCP13 = PUB_GetAKindSalesNo(strNP02, strNP03, strNP04, strNP05)
                  strExc(10) = "update NEXTPROGRESS set" & _
                               " np10='" & strCP13 & "'," & _
                               " NP15=(SELECT DECODE(NP15,NULL,'" & strUpdDate & "整批改業務員,原為'||NP10||ST02," & _
                               "      NP15||';" & strUpdDate & "整批改業務員,原為'||NP10||ST02||';') FROM STAFF WHERE NP10=ST01)" & _
                               " where np01='" & strNP01 & "'" & _
                               " and np02='" & strNP02 & "'" & _
                               " and np03='" & strNP03 & "'" & _
                               " and np04='" & strNP04 & "'" & _
                               " and np05='" & strNP05 & "'" & _
                               " and np22=" & strNP22
                  cnnConnection.Execute strExc(10)
                  rsTmp.MoveNext
               Loop
            End If
            
            'FCT,T,TF案改為依案件之FCT承辦智權人員規則PUB_GetFCTSalesNo
            '其他系統類別案件, 等碰到再考慮 !
            strExc(0) = "SELECT NP01,NP02,NP03,NP04,NP05,NP22 from NEXTPROGRESS,trademark where NP06 IS NULL AND NP10='" & Text1(0).Text & "' AND NP02 IN ('FCT','T','TF')" & _
                        " and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) and tm29 is null"
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               rsTmp.MoveFirst
               Do While Not rsTmp.EOF
                  strNP01 = rsTmp.Fields("NP01")
                  strNP02 = rsTmp.Fields("NP02")
                  strNP03 = rsTmp.Fields("NP03")
                  strNP04 = rsTmp.Fields("NP04")
                  strNP05 = rsTmp.Fields("NP05")
                  strNP22 = rsTmp.Fields("NP22")
                  strCP13 = PUB_GetFCTSalesNo(strNP02, strNP03, strNP04, strNP05)
                  strExc(10) = "update NEXTPROGRESS set" & _
                               " np10='" & strCP13 & "'," & _
                               " NP15=(SELECT DECODE(NP15,NULL,'" & strUpdDate & "整批改業務員,原為'||NP10||ST02," & _
                               "      NP15||';" & strUpdDate & "整批改業務員,原為'||NP10||ST02||';') FROM STAFF WHERE NP10=ST01)" & _
                               " where np01='" & strNP01 & "'" & _
                               " and np02='" & strNP02 & "'" & _
                               " and np03='" & strNP03 & "'" & _
                               " and np04='" & strNP04 & "'" & _
                               " and np05='" & strNP05 & "'" & _
                               " and np22=" & strNP22
                  cnnConnection.Execute strExc(10)
                  rsTmp.MoveNext
               Loop
            End If
            
         Else
'--該員工為F3,F4投資法務人員:要看案件情形
'--若有案件,可參考F1外商人員方式更新資料
'--若無案件則
'1.以申請人名稱找名稱近似
'2.以申請人名稱找傳票摘要
'  SELECT * FROM ACC021 WHERE INSTR(AX212,'名稱代表值')>0 ORDER BY AX202,AX203
'
'--2017/5/12廖宗岳經理退休(SQL僅參考用)
'update customer set
'    cu12='F31',
'    cu13='99015',
'    CU129=DECODE(CU129,NULL,CU13,CU129||','||CU13),
'    CU79=(SELECT DECODE(CU79,NULL,'" & strUpdDate & "整批改業務員,原為'||CU13||ST02||'改至開發人員欄',
'                                  '" & strUpdDate & "整批改業務員,原為'||CU13||ST02||'改至開發人員欄;'||CU79)
'          FROM STAFF WHERE CU13=ST01)
'  where cu13='離職人員員工編號';
            MsgBox "無程式需執行！", vbInformation
         End If
         
         cnnConnection.CommitTrans
         
         Screen.MousePointer = vbDefault
         Call Command1_Click(3)
         MsgBox "資料修改完畢！", vbInformation
         
      Case 1 '結束
         Unload frm140118
         Set frm140118 = Nothing
         
      Case 3 '檢查資料
         If LblST02.Caption = "" Then Exit Sub
         Command1(0).Enabled = True
         List1.Clear
         If Left(LblST03, 2) = "F1" Then 'F1外商人員:NA55.FCT承辦業務員
            '找出該員工名下的客戶編號, 名稱, 國籍, 該國籍的NA55, 客戶備註
            strExc(0) = "select cu01||cu02,Decode(CU04,Null,Decode(CU05,Null,CU06,CU05|| ' '||CU88||' '||CU89||' '||CU90),CU04),cu10,na03,NA55,s2.st02,cu79" & _
                        " from customer,nation,staff s2" & _
                        " where cu13='" & Text1(0).Text & "' and cu10=na01(+) and NA55=s2.st01(+)" & _
                        " ORDER BY NA55,CU10,1"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            List1.AddItem "■找出該員工名下的客戶編號, 名稱, 國籍, 該國籍的NA55, 客戶備註"
         Else
            '找出該員工名下的客戶編號, 名稱, 國籍, 該國籍的na51, 客戶備註
            strExc(0) = "select cu01||cu02,Decode(CU04,Null,Decode(CU05,Null,CU06,CU05|| ' '||CU88||' '||CU89||' '||CU90),CU04),cu10,na03,na51,s2.st02,cu79" & _
                        " from customer,nation,staff s2" & _
                        " where cu13='" & Text1(0).Text & "' and cu10=na01(+) and na51=s2.st01(+)" & _
                        " ORDER BY NA51,CU10,1"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            List1.AddItem "■找出該員工名下的客戶編號, 名稱, 國籍, 該國籍的na51, 客戶備註"
         End If
         List1.AddItem convForm("客戶編號", 10) & _
                       " " & convForm("名稱", 50) & _
                       " " & convForm("國籍", 30) & _
                       " " & convForm(IIf(Left(LblST03, 2) = "F1", "FCT承辦業務員", "FCP承辦業務員"), 20) & _
                       " " & convForm("客戶備註", 62)
         List1.AddItem String(10, "=") & _
                       " ==========================" & _
                       " ===============" & _
                       " " & String(20, "=") & _
                       " " & String(60, "=")
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               strText = convForm("" & RsTemp.Fields(0), 10)
               strText = strText & " " & convForm("" & RsTemp.Fields(1), 30)
               strText = strText & " " & convForm("" & RsTemp.Fields(2) & RsTemp.Fields(3), 20)
               strText = strText & " " & convForm("" & RsTemp.Fields(4) & RsTemp.Fields(5), 16)
               strText = strText & " " & convForm("" & RsTemp.Fields(6), 60)
               List1.AddItem strText
               RsTemp.MoveNext
            Loop
            List1.AddItem "共 " & RsTemp.RecordCount & " 筆"
         Else
            List1.AddItem "（無）"
         End If
         List1.AddItem ""
         '找出該掛在該員工名下未續辦的下一程序資料,帶出本所案號, 下一程序, 本所期限, 法定期限, np22, 備註
         strExc(0) = "select NP02||'-'||NP03||'-'||NP04||'-'||NP05,NP07,NP08,NP09,NP22,NP15 from nextprogress,trademark" & _
                     " where np10='" & Text1(0).Text & "' and np06 is null and np02 in ('T','TF','CFT','FCT')" & _
                     " and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) and tm29 is null union" & _
                     " select NP02||'-'||NP03||'-'||NP04||'-'||NP05,NP07,NP08,NP09,NP22,NP15 from nextprogress,patent" & _
                     " where np10='" & Text1(0).Text & "' and np06 is null and np02 in ('P','CFP','FCP')" & _
                     " and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and pa57 is null union" & _
                     " select NP02||'-'||NP03||'-'||NP04||'-'||NP05,NP07,NP08,NP09,NP22,NP15 from nextprogress,lawcase" & _
                     " where np10='" & Text1(0).Text & "' and np06 is null and np02 in ('L','LIN','ACS','CFL','FCL')" & _
                     " and np02=lc01(+) and np03=lc02(+) and np04=lc03(+) and np05=lc04(+) and lc08 is null union" & _
                     " select NP02||'-'||NP03||'-'||NP04||'-'||NP05,NP07,NP08,NP09,NP22,NP15 from nextprogress,hirecase" & _
                     " where np10='" & Text1(0).Text & "' and np06 is null and np02='LA'" & _
                     " and np02=hc01(+) and np03=hc02(+) and np04=hc03(+) and np05=hc04(+) and hc09 is null union" & _
                     " select NP02||'-'||NP03||'-'||NP04||'-'||NP05,NP07,NP08,NP09,NP22,NP15 from nextprogress,servicepractice" & _
                     " where np10='" & Text1(0).Text & "' and np06 is null and np02 not in ('T','TF','CFT','FCT','P','CFP','FCP','L','LIN','ACS','CFL','FCL','LA')" & _
                     " and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+) and sp15 is null" & _
                     " ORDER BY 1"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         List1.AddItem "■找出該掛在該員工名下未續辦的下一程序資料,帶出本所案號, 下一程序, 本所期限, 法定期限, np22, 備註"
         List1.AddItem convForm("本所案號", 23) & _
                       " " & convForm("下一程序", 13) & _
                       " " & convForm("本所期限", 11) & _
                       " " & convForm("法定期限", 13) & _
                       " " & convForm("np22", 16) & _
                       " " & convForm("備註", 60)
         List1.AddItem String(16, "=") & _
                       " " & String(10, "=") & _
                       " " & String(10, "=") & _
                       " " & String(10, "=") & _
                       " " & String(10, "=") & _
                       " " & String(60, "=")
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               strText = convForm("" & RsTemp.Fields(0), 18)
               strText = strText & " " & convForm("" & RsTemp.Fields(1), 18)
               strText = strText & " " & convForm("" & RsTemp.Fields(2), 11)
               strText = strText & " " & convForm("" & RsTemp.Fields(3), 12)
               strText = strText & " " & convForm("" & RsTemp.Fields(4), 10)
               strText = strText & " " & convForm("" & RsTemp.Fields(5), 60)
               List1.AddItem strText
               RsTemp.MoveNext
            Loop
            List1.AddItem "共 " & RsTemp.RecordCount & " 筆"
         Else
            List1.AddItem "（無）"
         End If
         List1.AddItem ""
         '檢查下一程序檔的所有系統類別,外商及外法有些系統類別需人工更新
         strExc(0) = "select np02,np10,COUNT(*) from (" & _
                     " select nextprogress.* from NEXTPROGRESS,trademark" & _
                     " where NP06 IS NULL AND NP10='" & Text1(0).Text & "' and np02 in ('T','TF','CFT','FCT')" & _
                     " and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) and tm29 is null union" & _
                     " select nextprogress.* from NEXTPROGRESS,patent" & _
                     " where NP06 IS NULL AND NP10='" & Text1(0).Text & "' and np02 in ('P','CFP','FCP')" & _
                     " and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and pa57 is null union" & _
                     " select nextprogress.* from NEXTPROGRESS,lawcase" & _
                     " where NP06 IS NULL AND NP10='" & Text1(0).Text & "' and np02 in ('L','LIN','ACS','CFL','FCL')" & _
                     " and np02=lc01(+) and np03=lc02(+) and np04=lc03(+) and np05=lc04(+) and lc08 is null union" & _
                     " select nextprogress.* from NEXTPROGRESS,hirecase" & _
                     " where NP06 IS NULL AND NP10='" & Text1(0).Text & "' and np02='LA'" & _
                     " and np02=hc01(+) and np03=hc02(+) and np04=hc03(+) and np05=hc04(+) and hc09 is null union" & _
                     " select nextprogress.* from NEXTPROGRESS,servicepractice" & _
                     " where NP06 IS NULL AND NP10='" & Text1(0).Text & "' and np02 not in ('T','TF','CFT','FCT','P','CFP','FCP','L','LIN','ACS','CFL','FCL','LA')" & _
                     " and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+) and sp15 is null)" & _
                     " group by np02,np10"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         List1.AddItem "■檢查下一程序檔的所有系統類別,外商及外法有些系統類別需人工更新"
         List1.AddItem convForm("np02", 12) & _
                       " " & convForm("np10", 12) & _
                       " " & convForm("COUNT(*)", 12)
         List1.AddItem String(12, "=") & _
                       " " & String(12, "=") & _
                       " " & String(12, "=")
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               strText = convForm("" & RsTemp.Fields(0), 12)
               strText = strText & " " & convForm("" & RsTemp.Fields(1), 12)
               strText = strText & " " & convForm("" & RsTemp.Fields(2), 12)
               List1.AddItem strText
               RsTemp.MoveNext
            Loop
            List1.AddItem "共 " & RsTemp.RecordCount & " 筆"
         Else
            List1.AddItem "（無）"
         End If
         List1.AddItem ""
         '先檢查下一程序檔無智權人員的資料筆數,因為發現下面更新語法可能有問題
         strExc(0) = "SELECT NP02,COUNT(*) FROM NEXTPROGRESS WHERE NP10 IS NULL AND NP17>=20030201 GROUP BY NP02"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         List1.AddItem "■先檢查下一程序檔無智權人員的資料筆數,因為發現下面更新語法可能有問題"
         List1.AddItem convForm("np02", 12) & _
                       " " & convForm("np10", 12)
         List1.AddItem String(12, "=") & _
                       " " & String(12, "=")
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               strText = convForm("" & RsTemp.Fields(0), 12)
               strText = strText & " " & convForm("" & RsTemp.Fields(1), 12)
               List1.AddItem strText
               RsTemp.MoveNext
            Loop
            List1.AddItem "共 " & RsTemp.RecordCount & " 筆"
         Else
            List1.AddItem "（無）"
         End If
         'If List1.ListCount > 0 Then SetListScroll List1   'cancel by sonia 2021/12/13
         List1.AddItem ""   'add by sonia 2021/12/13否則最後一行看不到
   End Select
   
   Set rsTmp = Nothing
   Exit Sub
   
ErrHand:
   MsgBox "錯誤 : " & Err.Description, vbInformation
   Set rsTmp = Nothing
   Exit Sub
   
ErrorHandler:
   cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox "修改資料失敗，請洽系統管理員 !", vbCritical
   Set rsTmp = Nothing
End Sub

Private Sub ClearAll()
   LblST02.Caption = ""
   LblST52.Caption = ""
   LblST53.Caption = ""
   LblST54.Caption = ""
   LblST55.Caption = ""
   LblST03.Caption = ""
   LblST16.Caption = ""
   List1.Clear
   LblNote.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140118 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   If Index = 0 Then
      CloseIme
   Else
      OpenIme
   End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Text1(Index) = "" Then Exit Sub
   Select Case Index
      Case 0
         Command1(0).Enabled = False
         Command1(3).Enabled = False
         Call ClearAll
         strExc(0) = "SELECT * FROM staff,acc090 WHERE st01='" & Text1(0).Text & "' and A0901(+)=st03"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            LblST02.Caption = RsTemp.Fields("st02")
            LblST52.Caption = "" & RsTemp.Fields("st52") & GetPrjSalesNM("" & RsTemp.Fields("st52"))
            LblST53.Caption = "" & RsTemp.Fields("st53") & GetPrjSalesNM("" & RsTemp.Fields("st53"))
            LblST54.Caption = "" & RsTemp.Fields("st54") & GetPrjSalesNM("" & RsTemp.Fields("st54"))
            LblST55.Caption = "" & RsTemp.Fields("st55") & GetPrjSalesNM("" & RsTemp.Fields("st55"))
            LblST03.Caption = "" & RsTemp.Fields("st03") & "" & RsTemp.Fields("A0902")
            LblST16.Caption = "" & RsTemp.Fields("st16")
            Select Case Left("" & RsTemp.Fields("st03"), 2)
               Case "F1" '外商
                  If LblST16.Caption = "2" Then
                     LblST16.Caption = LblST16.Caption & " 英"
                  ElseIf LblST16.Caption = "4" Then
                     LblST16.Caption = LblST16.Caption & " 日"
                  End If
                  LblNote.Caption = "若有非商標的案件，要再加語法！"
               Case "F2" '外專
                  If LblST16.Caption = "1" Then
                     LblST16.Caption = LblST16.Caption & " 電"
                  ElseIf LblST16.Caption = "2" Then
                     LblST16.Caption = LblST16.Caption & " 化"
                  ElseIf LblST16.Caption = "3" Then
                     LblST16.Caption = LblST16.Caption & " 日"
                  ElseIf LblST16.Caption = "4" Then
                     LblST16.Caption = LblST16.Caption & " 機"
                  End If
                  LblNote.Caption = "若有非專利或專利服務的案件，要再加語法！"
               Case "F3" '法務
                  If LblST16.Caption = "1" Then
                     LblST16.Caption = LblST16.Caption & " 英"
                  ElseIf LblST16.Caption = "2" Then
                     LblST16.Caption = LblST16.Caption & " 日"
                  End If
                  LblNote.Caption = "語法尚未撰寫！"
            End Select
            If "" & RsTemp.Fields("st04") = "1" Then
               MsgBox "此員工並未離職不可更新資料！", vbCritical
               Text1(0).SetFocus
            Else
               Command1(3).Enabled = True
            End If
         Else
            MsgBox "查無此員工編號！", vbCritical
            Text1(0).SetFocus
         End If
   End Select
   If Cancel = True Then TextInverse Text1(Index)
   If Cancel = False Then CloseIme
End Sub

'cancel by sonia 2021/12/13
'Private Sub SetListScroll(oList As ListBox)
'   Dim ii As Integer
'   Dim lWnow As Long, lWmax As Long
'
'   lWmax = 0
'   For ii = 0 To oList.ListCount - 1
'      lWnow = TextWidth(oList.List(ii) & " ")
'      If lWnow > lWmax Then
'         lWmax = lWnow
'      End If
'   Next
'
'   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
'   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
'End Sub
