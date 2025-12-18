VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060122 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專翻譯分案作業"
   ClientHeight    =   5676
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   10608
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5676
   ScaleWidth      =   10608
   Begin VB.CheckBox Check2 
      Caption         =   "含已經過主管確認"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   653
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含未提申之案件"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8520
      TabIndex        =   13
      Top             =   5360
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "主管確認(&S)"
      Height          =   375
      Index           =   6
      Left            =   6360
      TabIndex        =   12
      Top             =   480
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "E-Mail(&E)"
      Height          =   375
      Index           =   5
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1620
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   8520
      TabIndex        =   5
      Top             =   120
      Width           =   900
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm060122.frx":0000
      Left            =   8010
      List            =   "frm060122.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   630
      Width           =   2355
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "新案建檔(&U)"
      Height          =   375
      Index           =   3
      Left            =   6360
      TabIndex        =   3
      Top             =   120
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "分案(&A)"
      Height          =   375
      Index           =   4
      Left            =   7560
      TabIndex        =   4
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "卷宗區(&C)"
      Height          =   375
      Index           =   2
      Left            =   1810
      TabIndex        =   1
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Top             =   120
      Width           =   900
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4065
      Left            =   45
      TabIndex        =   8
      Top             =   1080
      Width           =   10515
      _ExtentX        =   18542
      _ExtentY        =   7176
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      BackColorBkg    =   16772048
      AllowUserResizing=   3
      FormatString    =   $"frm060122.frx":0004
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   15
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Left            =   1080
      TabIndex        =   15
      Top             =   630
      Width           =   2010
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3545;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblSales 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   690
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "符號說明：▲急件翻譯已立案 ●未提申先翻譯 ◎一案兩請之發明案 ◆新型案 ＊上班譯 ⊕所內譯"
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   5370
      Width           =   7605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "顏色符號說明："
      Height          =   180
      Left            =   6720
      TabIndex        =   9
      Top             =   690
      Width           =   1260
   End
End
Attribute VB_Name = "frm060122"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/24 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB、Combo2
'Created by Lydia 2018/05/15外專翻譯分案作業
Option Explicit
'Added by Lydia 2018/08/15 設定可使用表單
Public Tmpfrm060102 As Form '新案建檔
Public Tmpfrm060504 As Form '案件命名追蹤
Public Tmpfrm060101_1 As Form '外專分案
Public Tmpfrm060101_3 As Form '外專FMP分案
Public PubRole As String '1-分案, 2-認翻譯
'end 2018/08/15
Public cmdState As Integer
Public bolNextDone As Boolean '下一畫面作業完成
Public nKeyNo As String '下一畫面輸入的承辦人
Private Const cFixed As Integer = 7 '固定欄位 'Modified by Lydia 2018/08/29 6=>7
Private Const 認領啟用日 = "20190801" 'Added by Lydia 2018/09/27 若要啟用,請修改為上線日'Modified by Lydia 2019/08/14 配合翻譯費新費率,直接上線

Dim mType As String '翻譯狀態: 0-急件翻譯(未立案), 1-急件翻譯(已立案), 2-一般翻譯, 3-未提申先翻譯, 4-英文參考本
Dim mTransKind As String 'Y:只能上班翻譯 (Y後面+逗號,再加上有折扣或固定報價)
Dim m_bUpdate As Boolean
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_Grp As String, m_GrpMan As String, m_GrpManList As String  '工程師組別和主管
Dim m2FileName As String '工作確認單範本

'Grid欄位設定
Dim colC1CP09 As Integer    '新案翻譯收文號
Dim colC2CP09 As Integer    '新案收文號
Dim colCase As Integer        '案號/追蹤號

Dim colTcn15 As Integer     '急件翻譯人員
Dim colTFA04 As Integer 'Added by Lydia 2018/09/25 認領人員
Dim colTF29 As Integer       '待比對
Dim colTF30 As Integer       '待英文本
Dim colTF26 As Integer       '交稿期限
Dim colC1CP48 As Integer    'Added by Lydia 承辦期限
Dim colTF32 As Integer       '只交Claim期限
Dim colDc As Integer         '有折扣
Dim colFc As Integer          '固定報價
Dim colTF34 As Integer          'Added by Lydia 2018/08/24 暫不翻譯
Dim colTF19 As Integer 'Added by Lydia 2018/09/25 相似度
Dim colFAno As Integer, colAppNo1 As Integer 'Added by Lydia 2023/04/19 代理人編號(FANO),申請人1(APPNO1)
'Modified by Lydia 2025/06/05 更改名稱
'Dim m_strBASF As String 'Added by Lydia 2023/04/19 BASF集團的X編號
Dim m_str所內譯 As String
Dim m_str所內譯例外 As String 'Added by Lydia 2025/07/01
Dim m_AttachPath As String 'Added  by Lydia 2023/08/18 附件宣告區

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Function QueryData(Optional ByRef bolM As Boolean = True) As Boolean
'Memo by Lydia 2019/08/15 如果更新工程師看案件(非急件)的語法,請一併更新frmAutoBatchDay.strMenu98
Dim rsRead As New ADODB.Recordset
Dim strCon1 As String
Dim strCon2 As String
Dim strCon1_1 As String 'Added by Lydia 2018/09/25

QueryData = False
   
On Error GoTo ErrorHand2:
    
    Call SetGrd(True) '清空
       
    '工程師只能看該組案件
    'Modified by Lydia 2018/09/25
    'If m_Grp <> "" And m_bUpdate = False Then
    If PubRole = "2" Then
        strCon1 = " and pa150=" & CNULL(m_Grp)
        strCon1 = strCon1 & " and nvl(tf30,'N') <> 'Y' " 'Added by Lydia 2018/08/17 排除待英文本
        strCon2 = " and 0 =1 "
        'Added by Lydia 2018/09/25  排除主管已確認的認領
        If Check2.Value = vbUnchecked Then
            'If InStr(m_GrpManList, Trim(Left(Combo2.Text, 6))) = 0 Then 'Remove by Lydia 2018/10/02 以勾選項為準
                 strCon1 = strCon1 & " and tfa06 is null "
            'End If 'end 2018/10/02
        End If
        'Added by Lydia 2019/08/23 不顯示"待比對TF29,待英文本TF30,暫不翻譯TF34"
        strCon1 = strCon1 & " and TF29 is null and TF34 is null and substr(nvl(TF30,'B'),1,1)='B' "
        
    'Added by Lydia 2018/09/25 分案
    Else
         strCon1_1 = strCon1_1 & " and tfa06 is not null " '顯示有主管確認的認領人員
    'end 2018/09/25
    End If
    
    'Added by Lydia 2018/08/29 增加含未提申之案件,將and c2.cp27||tf31||tcn14 is not null抽出
    If Check1.Value = 0 Then
        strCon1 = strCon1 & " and c2.cp27||tf31||tcn14 is not null "
    End If
    
    '排序
    strExc(1) = "decode(tcn14,null, " & _
                                "decode(tf31,null, " & _
                                        "decode(b1.cm01||b2.cm05,null,decode(pa08,'2','13','20'),decode(pa08,'1','12','15')) " & _
                                ",'11') " & _
                      ",'10') ord1, "
    '符號
    strExc(2) = "decode(tcn14,null,'','▲')||decode(tf31,null,'','●')||decode(b1.cm01||b2.cm05,null,'',decode(pa08,'1','◎',''))||decode(pa08,'2','◆','')||"
    '一般翻譯:已提申(新申請案已發文)或勾選未提申先翻譯(TF31)
    'FCP案抓新案翻譯未設承辦人; FMP案和寰華組在內專分案輸入工程師組別,同時設定該案未分案之新案翻譯,製作中說,檢視中說,901告知代理人的承辦人為該組管制人
    'Modified by Lydia 2018/08/21 改命名作業之其他認翻譯人員 '其他-'||tct28 =>decode(substr(upper(tct28),-1),'A',substr(tct28,1,length(tct28)-1)||'-下班','B',substr(tct28,1,length(tct28)-1)||'-上班',tct28)
    'Modified by Lydia 2018/08/27 +暫不翻譯TF34
    'Modified by Lydia 2018/08/29 顯示+只交Claims期限tf32t
    'Modified by Lydia 2018/08/29 增加含未提申之案件,將and c2.cp27||tf31||tcn14 is not null抽出
    'Modified by Lydia 2018/09/25 改抓工程師認翻譯人員
'    strSql = "select ' ' V, " & strExc(1) & strExc(2) & _
'                "c1.cp01||'-'||c1.cp02||DECODE(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04) AS caseno,sqldatet(pa10) pa10t,sqldatet(nvl(tf26,c1.cp48)) tf26t,sqldatet(tf32) tf32t" & _
'                ",nvl(s2.st02,decode(tct27,'1','舜禹','2','捷恩凱','3','迅達','4',decode(substr(upper(tct28),-1),'A',substr(tct28,1,length(tct28)-1)||'-下班','B',substr(tct28,1,length(tct28)-1)||'-上班',tct28),'A',s1.st02||'-下班','B',s1.st02||'-上班',tct27)) as transman" & _
'                ",tf33,decode(pa49||pa50||fa25||fa26||cu36||cu37,null,'','Y') as dcprice,decode(x01||y01,null,'','Y') as fcprice" & _
'                ",tf20,decode(tf19,0,null,tf19) tf19,s1.st02 as tct10n,decode(substr(tf30,1,1),'B','',tf30) tf30t,tf29" & _
'                ",decode(pa150,'1','" & PUB_GetFCPGrpName("1") & "','2','" & PUB_GetFCPGrpName("2") & "','3','" & PUB_GetFCPGrpName("3") & "','4','" & PUB_GetFCPGrpName("4") & "',pa150) grpname" & _
'                ",decode(tct25,'1','生醫','2','化學','3','化工','4','材料','5','電子','6','機械','7','電機','8','其他',tct25) tct25n " & _
'                ",nvl(pa05,nvl(pa06,pa07)) casename,pa150,tf01 as c1_cp09,c2.cp09 as c2_cp09,tf26,tct25,tct27,tct28,tcn15,tcn14,tf30,tf31,tf32,tf34 " & _
'                "from TransFee,CaseProgress c1,CaseProgress c2 ,TransCaseTitle,staff s1,patent" & _
'                ",customer,fagent,TrackingCaseName,staff s2" & _
'                ",(select aal04 as x01 from addressa4list where aal01='FCPtct' and substr(aal04,1,1)='X') vtb1" & _
'                ",(select aal04 as y01 from addressa4list where aal01='FCPtct' and substr(aal04,1,1)='Y') vtb2 " & _
'                ", casemap b1 ,casemap b2 " & _
'                "where tf01=c1.cp09(+) and c1.cp10 ='201' and c1.cp158=0 and c1.cp159=0 and pa57 is null " & strCon1 & _
'                "and (c1.cp01||c1.cp14='FCP' or (c1.cp01='P' and instr('" & m_GrpManList & "',c1.cp14)>0)) " & _
'                "and c1.cp01=pa01(+) and c1.cp02=pa02(+) and c1.cp03=pa03(+) and c1.cp04=pa04(+) " & _
'                "and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+) and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) and c2.cp31='Y' " & _
'                "and c2.cp09=tct01(+) and tct10=s1.st01(+) " & _
'                "and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
'                "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) " & _
'                "and pa26=x01(+) and pa75=y01(+) and tf01=tcn14(+) and tcn15=s2.st01(+) " & _
'                "and b1.cm01(+)=c1.cp01 and b1.cm02(+)=c1.cp02 and b1.cm03(+)=c1.cp03 and b1.cm04(+)=c1.cp04 and b1.cm10(+)='3' " & _
'                "and b2.cm05(+)=c1.cp01 and b2.cm06(+)=c1.cp02 and b2.cm07(+)=c1.cp03 and b2.cm08(+)=c1.cp04 and b2.cm10(+)='3' "
    'Modified by Lydia 2018/11/20 已立案之急件翻譯在本所案號後,加上原本的急件翻譯號(命名追蹤號) =>decode(tcn01,null,'','('||tcn01||')')
    'Modified by Lydia 2019/01/17 組別及類別 2欄位往前移至相似度之後,命名人員之前。
    'strSql = "select ' ' V, " & strExc(1) & strExc(2) & _
                "c1.cp01||'-'||c1.cp02||DECODE(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04)||decode(tcn01,null,'','('||tcn01||')') AS caseno,sqldatet(pa10) pa10t,sqldatet(nvl(tf26,c1.cp48)) tf26t,sqldatet(tf32) tf32t" & _
                ",nvl(s2.st02,tfa04n) transman" & _
                ",tf33,decode(pa49||pa50||fa25||fa26||cu36||cu37,null,'','Y') as dcprice,null as fcprice" & _
                ",tf20,decode(tf19,0,null,tf19) tf19,s1.st02 as tct10n,decode(substr(tf30,1,1),'B','',tf30) tf30t,tf29" & _
                ",decode(pa150,'1','" & PUB_GetFCPGrpName("1") & "','2','" & PUB_GetFCPGrpName("2") & "','3','" & PUB_GetFCPGrpName("3") & "','4','" & PUB_GetFCPGrpName("4") & "',pa150) grpname" & _
                ",decode(tct25,'1','生醫','2','化學','3','化工','4','材料','5','電子','6','機械','7','電機','8','其他',tct25) tct25n " & _
                ",nvl(pa05,nvl(pa06,pa07)) casename,pa150,tf01 as c1_cp09,c2.cp09 as c2_cp09,tf26,tct25,tcn15,tcn14,tf30,tf31,tf32,tf34,tfa04,tfa06,c1.cp48 as c1_cp48 " & _
                "from TransFee,CaseProgress c1,CaseProgress c2 ,TransCaseTitle,staff s1,patent" & _
                ",customer,fagent,TrackingCaseName,staff s2" & _
                ",(select tfa01,s3.st02||decode(tfa05,'A','-下班','B','-上班','') as tfa04n,tfa04,tfa06 from transfeeassign,staff s3 where tfa04=s3.st01(+) " & strCon1_1 & ") vtb3 " & _
                ", casemap b1 ,casemap b2 " & _
                "where tf01=c1.cp09(+) and c1.cp10 ='201' and c1.cp158=0 and c1.cp159=0 and pa57 is null " & strCon1 & _
                "and (c1.cp01||c1.cp14='FCP' or (c1.cp01='P' and instr('" & m_GrpManList & "',c1.cp14)>0)) " & _
                "and c1.cp01=pa01(+) and c1.cp02=pa02(+) and c1.cp03=pa03(+) and c1.cp04=pa04(+) " & _
                "and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+) and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) and c2.cp31='Y' " & _
                "and c2.cp09=tct01(+) and tct10=s1.st01(+) " & _
                "and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
                "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and tf01=tfa01(+) and tf01=tcn14(+) and tcn15=s2.st01(+) " & _
                "and b1.cm01(+)=c1.cp01 and b1.cm02(+)=c1.cp02 and b1.cm03(+)=c1.cp03 and b1.cm04(+)=c1.cp04 and b1.cm10(+)='3' " & _
                "and b2.cm05(+)=c1.cp01 and b2.cm06(+)=c1.cp02 and b2.cm07(+)=c1.cp03 and b2.cm08(+)=c1.cp04 and b2.cm10(+)='3' "
    'Modified by Lydia 2019/08/26 +翻譯特殊指示TF36
    'Modfied by Lydia 2023/04/19 +代理人編號(FANO),申請人1(APPNO1)
    strSql = "select ' ' V, " & strExc(1) & strExc(2) & _
                "c1.cp01||'-'||c1.cp02||DECODE(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04)||decode(tcn01,null,'','('||tcn01||')') AS caseno,sqldatet(pa10) pa10t,sqldatet(nvl(tf26,c1.cp48)) tf26t,sqldatet(tf32) tf32t" & _
                ",nvl(s2.st02,tfa04n) transman,tf33,decode(pa49||pa50||fa25||fa26||cu36||cu37,null,'','Y') as dcprice,null as fcprice,tf20,decode(tf19,0,null,tf19) tf19" & _
                ",decode(pa150,'1','" & PUB_GetFCPGrpName("1") & "','2','" & PUB_GetFCPGrpName("2") & "','3','" & PUB_GetFCPGrpName("3") & "','4','" & PUB_GetFCPGrpName("4") & "',pa150) grpname" & _
                ",decode(tct25,'1','生醫','2','化學','3','化工','4','材料','5','電子','6','機械','7','電機','8','其他',tct25) tct25n " & _
                ",s1.st02 as tct10n,decode(substr(tf30,1,1),'B','',tf30) tf30t,tf29" & _
                ",nvl(pa05,nvl(pa06,pa07)) casename,TF36,pa150,tf01 as c1_cp09,c2.cp09 as c2_cp09,tf26,tct25,tcn15,tcn14,tf30,tf31,tf32,tf34,tfa04,tfa06,c1.cp48 as c1_cp48,pa75 as FANO,pa26 as APPNO1 " & _
                "from TransFee,CaseProgress c1,CaseProgress c2 ,TransCaseTitle,staff s1,patent" & _
                ",customer,fagent,TrackingCaseName,staff s2" & _
                ",(select tfa01,s3.st02||decode(tfa05,'A','-下班','B','-上班','') as tfa04n,tfa04,tfa06 from transfeeassign,staff s3 where tfa04=s3.st01(+) " & strCon1_1 & ") vtb3 " & _
                ", casemap b1 ,casemap b2 " & _
                "where tf01=c1.cp09(+) and c1.cp10 ='201' and c1.cp158=0 and c1.cp159=0 and pa57 is null " & strCon1 & _
                "and (c1.cp01||c1.cp14='FCP' or (c1.cp01='P' and instr('" & m_GrpManList & "',c1.cp14)>0)) " & _
                "and c1.cp01=pa01(+) and c1.cp02=pa02(+) and c1.cp03=pa03(+) and c1.cp04=pa04(+) " & _
                "and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+) and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) and c2.cp31='Y' " & _
                "and c2.cp09=tct01(+) and tct10=s1.st01(+) " & _
                "and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
                "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and tf01=tfa01(+) and tf01=tcn14(+) and tcn15=s2.st01(+) " & _
                "and b1.cm01(+)=c1.cp01 and b1.cm02(+)=c1.cp02 and b1.cm03(+)=c1.cp03 and b1.cm04(+)=c1.cp04 and b1.cm10(+)='3' " & _
                "and b2.cm05(+)=c1.cp01 and b2.cm06(+)=c1.cp02 and b2.cm07(+)=c1.cp03 and b2.cm08(+)=c1.cp04 and b2.cm10(+)='3' "
    '急件翻譯(未立案,從Tracking_no而來)
    'Modified by Lydia 2018/08/27 +暫不翻譯TF34
    'Modified by Lydia 2018/08/29 顯示+只交Claims期限tf32t
    'Modified by Lydia 2018/09/25工程師認翻譯人員
'    strSql = strSql & "union all select ' ' V,'01' ord1,tcn14 as caseno,null as pa10t,sqldatet(tf26) tf26t,sqldatet(tf32) tf32t" & _
'                ",s1.st02 as transman,null as tf33,null as dcprice, null as fcprice,null as tf20, null as tf19,null as tct10n" & _
'                ", null as tf30t,null as tf29,null as grpname,null as tct25n,null as casename,null as pa150," & _
'                "null as c1_cp09,null as c2_cp09,tf26,null as tct25,null as tct27,null as tct28,tcn15,tcn14,tf30,tf31,tf32,TF34" & _
'                " from TrackingCaseName,TransFee,staff s1 where nvl(tcn14,'N') <'A'  and tcn14=tf01(+) and tcn15=s1.st01(+) " & strCon2
    'Modified by Lydia 2019/01/17 組別及類別 2欄位往前移至相似度之後,命名人員之前。
    'strSql = strSql & "union all select ' ' V,'01' ord1,tcn14 as caseno,null as pa10t,sqldatet(tf26) tf26t,sqldatet(tf32) tf32t" & _
                ",s1.st02 as transman,null as tf33,null as dcprice, null as fcprice,null as tf20, null as tf19,null as tct10n" & _
                ", null as tf30t,null as tf29,null as grpname,null as tct25n,null as casename,null as pa150" & _
                ",null as c1_cp09,null as c2_cp09,tf26,null as tct25,tcn15,tcn14,tf30,tf31,tf32,TF34,null as tfa04,null as tfa06,null as c1_cp48 " & _
                " from TrackingCaseName,TransFee,staff s1 where nvl(tcn14,'N') <'A'  and tcn14=tf01(+) and tcn15=s1.st01(+) " & strCon2
    'Modified by Lydia 2019/08/26 +翻譯特殊指示TF36
    'Modfied by Lydia 2023/04/19 +代理人編號(FANO),申請人1(APPNO1)
    strSql = strSql & "union all select ' ' V,'01' ord1,tcn14 as caseno,null as pa10t,sqldatet(tf26) tf26t,sqldatet(tf32) tf32t" & _
                ",s1.st02 as transman,null as tf33,null as dcprice, null as fcprice,null as tf20, null as tf19,null as grpname,null as tct25n" & _
                ",null as tct10n, null as tf30t,null as tf29,null as casename,null as TF36,null as pa150" & _
                ",null as c1_cp09,null as c2_cp09,tf26,null as tct25,tcn15,tcn14,tf30,tf31,tf32,TF34,null as tfa04,null as tfa06,null as c1_cp48,'' as FANO,'' as APPNO1 " & _
                " from TrackingCaseName,TransFee,staff s1 where nvl(tcn14,'N') <'A'  and tcn14=tf01(+) and tcn15=s1.st01(+) " & strCon2
    'Added by Lydia 2024/03/08 因應內專支援機械組OA (含P案) 收文: 927其他翻譯，承辦人為程序人員->案件帶入至 外專分案作業，同原程式分案
    strSql = strSql & " union select ' ' v,'01' ord1,'＃'||c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04) as caseno" & _
            " ,sqldatet(pa10) as pa10t,sqldatet(c1.cp48) as tf26t, null as tf32t,null as transman, null as tf33" & _
            " ,decode(pa49||pa50||fa25||fa26||cu36||cu37,null,'','Y') as dcprice,null as fcprice,null as tf20, null as tf19" & _
            " ,cst16(pa150) grpname,null as tct25n, null as tct10n,null as tf30t, null as tf29,nvl(pa05,nvl(pa06,pa07)) casename" & _
            " ,null as tf36,pa150,c1.cp09 as c1_cp09,c2.cp09 as c2_cp09, null as tf26,null as tct25,null as tcn15, null as tcn14" & _
            " ,null as tf30,null as tf31,null as tf32,null as tf34,null as tfa04,null as tfa06,c1.cp48 as c1_cp48,pa75 as fano,pa26 as appno1" & _
            " from caseprogress c1,caseprogress c2 ,staff,patent,fagent,customer" & _
            " where c1.cp01 in ('FCP','P') and c1.cp158=0 and c1.cp159=0 and c1.cp10='927' and c1.cp14=st01(+)  and st03='F22'" & _
            " and c1.cp43=c2.cp09(+) and substr(c2.cp09,1,1)='C' and c1.cp01=pa01(+) and c1.cp02=pa02(+) and c1.cp03=pa03(+) and c1.cp04=pa04(+)" & _
            " and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) "
    '優先處理:急件翻譯TCN14->未提申先翻譯TF31->一案兩請之發明案->新型案PA08->一般案(by 申請日)
    strSql = strSql & "order by ord1,pa10t "
    
    intI = 1
    Set rsRead = ClsLawReadRstMsg(intI, strSql)
    If intI = 1 Then
         QueryData = True
         MSHFlexGrid1.FixedCols = 0
         Set MSHFlexGrid1.Recordset = rsRead
         Call SetGrd
         MSHFlexGrid1.FixedCols = cFixed
    Else
          If bolM = True Then MsgBox "查無資料!!"
    End If
    
    Set rsRead = Nothing
    Exit Function

ErrorHand2:
   If Err.Number > 0 Then
      MsgBox Err.Description
      Exit Function
   End If
   
End Function

Private Sub cmdok_Click(Index As Integer)
   Me.Tag = "": mType = "": mTransKind = ""
   bolNextDone = False
   nKeyNo = ""
   cmdState = Index
   PubShowNextData
End Sub

Private Sub Form_Load()

    m_Grp = PUB_GetStaffST16(strUserNum)
    If m_Grp = "" And m_bUpdate = False Then m_Grp = "1" '預設為電子組
    'Modified by Lydia 2019/01/19
    'Select Case m_Grp
    '     Case "1": m_GrpMan = Pub_GetSpecMan("T")  '電子
    '     Case "2": m_GrpMan = Pub_GetSpecMan("R")  '化學
    '     Case "3": m_GrpMan = Pub_GetSpecMan("S")  '日文
    '     Case "4": m_GrpMan = Pub_GetSpecMan("T1")  '機械
    'End Select
    m_GrpMan = Pub_GetFCPGrpMan(m_Grp)
    'end 2019/01/19
      
    m_GrpManList = Pub_GetSt16Man(True)
    'Modified by Lydia 2025/06/05 更改名稱
    'm_str所內譯  = Pub_GetSpecMan("外專翻譯分案-BASF") & ","  'Added by Lydia 2023/04/19
    m_str所內譯 = Pub_GetSpecMan("外專翻譯分案-所內譯") & ","
    m_str所內譯例外 = Pub_GetSpecMan("外專翻譯分案-所內譯例外") & "," 'Added by Lydia 2025/07/01
    
    '是否為翻譯分案作業人員
    strExc(5) = Pub_GetSpecMan("外專對外翻聯絡人員")
    Combo2.Clear
    Combo2.AddItem strUserNum & " " & strUserName
    Combo2.ListIndex = 0
    cmdOK(6).Top = cmdOK(3).Top
    cmdOK(6).Visible = False
    Check2.Visible = False  'Added by Lydia 2018/09/25
    'Modified by Lydia 2018/08/15 判斷進入方式(1-分案, 2-認翻譯)
    'If Pub_StrUserSt03 = "M51" Or InStr(strExc(5), strUserNum) > 0 Then
    'Modified by Lydia 2018/10/01
    'If PubRole = "1" And (Pub_StrUserSt03 = "M51" Or InStr(strExc(5), strUserNum) > 0) Then
    If PubRole = "1" Then
         If Pub_StrUserSt03 = "M51" Or InStr(strExc(5), strUserNum) > 0 Then
    'end 2018/10/01
           'Added by Lydia 2023/08/18
           m_AttachPath = App.path & "\" & strUserNum
           Call Pub_ChkExcelPath(m_AttachPath)
           m_AttachPath = m_AttachPath & "\FCmail2"
           Call Pub_ChkExcelPath(m_AttachPath)
           PUB_KillAnyFile m_AttachPath
           'end 2023/08/18
           m_bUpdate = True
           cmdOK(5).Caption = "E-Mail(&E)"
           lblSales.Visible = False
           Combo2.Visible = False
           '下載工作通知單
           m2FileName = "外專翻譯_案件工作確認單樣本.doc"
           'Modified by Lydia 2023/08/18 + , , m_AttachPath
           Call PUB_GetSampleFile(m2FileName, "M51-000299-0-01", , m_AttachPath)
           'Added by Lydia 2019/08/06 下載郵件範本
           'Modified by Lydia 2023/08/18 + , , m_AttachPath
           Call PUB_GetSampleFile("$$TOT-000F22-0-01.oft", "TOT-000F22-0-01", , m_AttachPath)
         'Added by Lydia 2019/12/23 開放程序人員只供查詢功能
         ElseIf Pub_StrUserSt03 = "F22" Then
            m_bUpdate = False
            lblSales.Visible = False
            Combo2.Visible = False
            cmdOK(4).Visible = False
            cmdOK(6).Visible = False
            cmdOK(5).Visible = False
         End If 'end 2018/10/01
    Else
           Me.Caption = "外專翻譯分案-認翻譯"
           m_bUpdate = False
           cmdOK(5).Caption = "認領(&E)"
           'Added by Lydia 2019/08/21 隱藏顏色說明
           Label3.Visible = False
           Combo1.Visible = False
           
           'Added by Lydia 2018/08/15 先提供工程師看列表
           If 認領啟用日 >= strSrvDate(1) Then
                lblSales.Visible = False
                Combo2.Visible = False
                cmdOK(5).Visible = False
           Else  '有認領功能
                'Modified by Lydia 2021/04/14 外專翻譯承辦及核稿期限控管：日文組改由2位副理(99037簡偉倫、94012林軒吉)為最後確認主管(現為王協理)
                'If InStr(m_GrpManList, strUserNum) > 0 Or Pub_StrUserSt03 = "M51" Then '主管確認
                If InStr(m_GrpManList & ",99037,94012", strUserNum) > 0 Or Pub_StrUserSt03 = "M51" Then '主管確認
                    cmdOK(6).Visible = True
                    Check2.Visible = True
                End If
                lblSales.Visible = True
                Combo2.Visible = True
           End If
           'end 2018/08/15
           Check2.Visible = True 'Added by Lydia 2019/05/08 在開放認領前, 先提供查詢已主管確認的功能
    End If

    If m_bUpdate = False Then
       cmdOK(3).Visible = False
       cmdOK(4).Visible = False
    End If

   'Added by Lydia 2018/09/27 日文組不開放認領
   'Remove by Lydia 2019/06/19 by Lydia 2019/06/19 Sharon表示沒有這件事的印象
   'If m_Grp = "3" Then
   '     cmdOK(5).Visible = False
   'End If
   ''end 2018/09/27
   
   MoveFormToCenter Me
      
   Combo1.Clear
   '符號加在管制日期
   Combo1.AddItem "紅色：表示待比對"
   Combo1.AddItem "橘色：表示待英文本"
   Combo1.AddItem "黃色：表示暫不翻譯"   'Added by Lydia 2018/08/24
   Combo1.ListIndex = 0
   
   Call cmdok_Click(0)

End Sub

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   Dim lngColor As Long
   Dim strTmp1 As String 'Added by Lydia 2018/08/24
   Dim dbRate As Double 'Added by Lydia 2021/07/29
   
   'Modified by Lydia 2018/08/24 +TF34
   'Modified by Lydia 2018/08/29 +只交Claims期限
   'Modified by Lydia 2018/09/25 改抓工程師認翻譯人員
   ' arrGridHeadText = Array("V", "ORD1", "案號/追蹤號", "申請日", "交稿期限", "交Claims", "認領人員", "不得延", "有折扣", "固定報價", "相似案號", "相似度", "命名人員", "待英文本", "待比對", "組別", "類別", "案件名稱", _
                                             "PA150", "C1_CP09", "C2_CP09", "TF26", "TCT25", "TCT27", "TCT28", "TCN15", "TCN14", "TF30", "TF31", "TF32", "TF34")
    'Modified by Lydia 2019/01/17 組別及類別 2欄位往前移至相似度之後,命名人員之前。
    'arrGridHeadText = Array("V", "ORD1", "案號/追蹤號", "申請日", "交稿期限", "交Claims", "認領人員", "不得延", "有折扣", "固定報價", "相似案號", "相似度", "命名人員", "待英文本", "待比對", "組別", "類別", "案件名稱", _
                                             "PA150", "C1_CP09", "C2_CP09", "TF26", "TCT25", "TCN15", "TCN14", "TF30", "TF31", "TF32", "TF34", "TFA04", "TFA06", "C1CP48")
    'Modified by Lydia 2019/08/26 +翻譯特殊指示TF36
    'Modified by Lydia 2023/04/19 +代理人編號(FANO),申請人1(APPNO1)
    arrGridHeadText = Array("V", "ORD1", "案號/追蹤號", "申請日", "交稿期限", "交Claims", "認領人員", "不得延", "有折扣", "固定報價", "相似案號", "相似度", "組別", "類別", "命名人員", "待英文本", "待比對", "案件名稱", "翻譯特殊指示", _
                                             "PA150", "C1_CP09", "C2_CP09", "TF26", "TCT25", "TCN15", "TCN14", "TF30", "TF31", "TF32", "TF34", "TFA04", "TFA06", "C1CP48", "FANO", "APPNO1")
   
    If m_bUpdate = True Then
         'Modified by Lydia 2018/08/29
         'arrGridHeadWidth = Array(300, 0, 1400, 840, 840, 840, 640, 640, 840, 1200, 640, 840, 840, 640, 800, 600, 2000)
         'Modified by Lydia 2019/01/17 組別及類別 2欄位往前移至相似度之後,命名人員之前。
         'arrGridHeadWidth = Array(300, 0, 1400, 840, 840, 840, 840, 640, 640, 840, 1200, 640, 840, 840, 640, 800, 600, 2000)
         'arrGridHeadText = Array("V", "ORD1", "案號/追蹤號", "申請日", "交稿期限", "交Claims", "認領人員", "不得延", "有折扣", "固定報價", "相似案號", "相似度", "命名人員", "待英文本", "待比對", "組別", "類別", "案件名稱"
         'Modified by Lydia 2019/08/26 +翻譯特殊指示TF36
         'arrGridHeadWidth = Array(300, 0, 1400, 840, 840, 840, 840, 640, 640, 840, 1200, 640, 800, 600, 840, 840, 640, 2000)
         arrGridHeadWidth = Array(300, 0, 1400, 840, 840, 840, 840, 640, 640, 840, 1200, 640, 800, 600, 840, 840, 640, 2000, 2000)
    Else
         'Modified by Lydia 2018/08/29 顯示申請日和只交Claims期限
         'Modified by Lydia 2019/01/17 組別及類別 2欄位往前移至相似度之後,命名人員之前。
         'arrGridHeadWidth = Array(300, 0, 1400, 840, 840, 840, 840, 640, 640, 840, 1200, 640, 0, 0, 0, 0, 600, 2000)
         'Modified by Lydia 2019/08/26 +翻譯特殊指示TF36
         'arrGridHeadWidth = Array(300, 0, 1400, 840, 840, 840, 840, 640, 640, 840, 1200, 640, 0, 600, 0, 0, 0, 2000)
         arrGridHeadWidth = Array(300, 0, 1400, 840, 840, 840, 840, 640, 640, 840, 1200, 640, 0, 600, 0, 0, 0, 2000, 2000)
    End If
   
   
   MSHFlexGrid1.Visible = False
   MSHFlexGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
         MSHFlexGrid1.Clear
         MSHFlexGrid1.Rows = 2
   End If
       
    For iRow = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.row = 0
       MSHFlexGrid1.col = iRow
       MSHFlexGrid1.Text = arrGridHeadText(iRow)
       If iRow <= UBound(arrGridHeadWidth) Then
            MSHFlexGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
       Else '案件名稱以後的欄位
            MSHFlexGrid1.ColWidth(iRow) = 0
       End If
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
    Next

    If colC1CP09 = 0 Then
         colC1CP09 = PUB_MGridGetId("C1_CP09", MSHFlexGrid1)
         colC2CP09 = PUB_MGridGetId("C2_CP09", MSHFlexGrid1)
         colTcn15 = PUB_MGridGetId("TCN15", MSHFlexGrid1)
         colCase = PUB_MGridGetId("案號/追蹤號", MSHFlexGrid1)
         colTF29 = PUB_MGridGetId("待比對", MSHFlexGrid1)
         colTF30 = PUB_MGridGetId("TF30", MSHFlexGrid1)
         colTF26 = PUB_MGridGetId("TF26", MSHFlexGrid1)
         colTF32 = PUB_MGridGetId("TF32", MSHFlexGrid1)
         colDc = PUB_MGridGetId("有折扣", MSHFlexGrid1)
         colFc = PUB_MGridGetId("固定報價", MSHFlexGrid1)
         colTF34 = PUB_MGridGetId("TF34", MSHFlexGrid1) 'Added by Lydia 2018/08/24 暫不翻譯
          'Added by Lydia 2018/09/25
         colTFA04 = PUB_MGridGetId("TFA04", MSHFlexGrid1)
         colC1CP48 = PUB_MGridGetId("C1CP48", MSHFlexGrid1)
         colTF19 = PUB_MGridGetId("相似度", MSHFlexGrid1)
         'end 2018/09/25
         'Added by Lydia 2023/04/19
         colFAno = PUB_MGridGetId("FANO", MSHFlexGrid1)
         colAppNo1 = PUB_MGridGetId("APPNO1", MSHFlexGrid1)
         'end 2023/04/19
    End If

   For intI = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.row = intI
         '待比對 -> 紅色
         If "" & MSHFlexGrid1.TextMatrix(intI, colTF29) = "Y" Then
               lngColor = &HFF&
         '待英文本 -> 橘色
        ElseIf "" & MSHFlexGrid1.TextMatrix(intI, colTF30) = "Y" Then
               lngColor = &H80FF&
         'Added by Lydia 2018/08/24 暫不翻譯-> 黃色
        ElseIf "" & MSHFlexGrid1.TextMatrix(intI, colTF34) = "Y" Then
               lngColor = &HFFFF&
        'end 2018/08/24
        Else
               lngColor = &H80000005
        End If
        For iRow = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.col = iRow
           MSHFlexGrid1.CellBackColor = lngColor
           'Modified by Lydia 2019/01/17 組別及類別 2欄位往前移至相似度之後,命名人員之前。
           'If InStr("00,07,08,09,10,12,13", Format(iRow, "00")) > 0 Then  '置中
           If InStr("00,07,08,09,11,15,16", Format(iRow, "00")) > 0 Then
                MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
                'Added by Lydia 2018/08/27 固定報價改在SetGrd逐筆判斷
                If iRow = colFc And "" & MSHFlexGrid1.TextMatrix(intI, iRow) = "" And "" & MSHFlexGrid1.TextMatrix(intI, colCase) <> "" Then
                      'Modified by Lydia 2018/11/20
                      'strExc(1) = Pub_RplStr(MSHFlexGrid1.TextMatrix(intI, colCase))  '清掉符號
                      strExc(1) = MSHFlexGrid1.TextMatrix(intI, colCase)
                      If InStr(strExc(1), "(") > 0 Then strExc(1) = Mid(strExc(1), 1, InStr(strExc(1), "(") - 1)
                      strExc(1) = Pub_RplStr(strExc(1))
                      'end 2018/11/20
                      If InStr(strExc(1), "-") > 0 Then
                           If InStrRev(strExc(1), "-") < 6 Then strExc(1) = strExc(1) & "-0-00"
                      End If
                      strExc(1) = Replace(strExc(1), "-", "")
                      strTmp1 = Pub_GetPa62Flag(strExc(1))
                      MSHFlexGrid1.Text = strTmp1
                End If
                'end 2018/08/27
           End If
           'Added by Lydia 2021/07/29 翻譯費折扣率＞30%客戶，在外專翻譯分案作業畫面的本所案號帶入符號
           If iRow = colCase And (Left("" & MSHFlexGrid1.TextMatrix(intI, colCase), 1) = "F" Or Left("" & MSHFlexGrid1.TextMatrix(intI, colCase), 1) = "P") Then
              strExc(0) = MSHFlexGrid1.TextMatrix(intI, colCase)
              If InStr(strExc(0), "(") > 0 Then strExc(0) = Mid(strExc(0), 1, InStr(strExc(0), "(") - 1)
              strExc(0) = Pub_RplStr(strExc(0))
              If InStr(strExc(0), "-") > 0 Then
                   If InStrRev(strExc(0), "-") < 6 Then strExc(0) = strExc(0) & "-0-00"
              End If
              strExc(0) = Replace(strExc(0), "-", "")
              Call ChgCaseNo(strExc(0), strExc)
              'Added by Lydia 2023/04/19 BASF集團公司為申請人的所有專利案件相關翻譯事宜（201新案翻譯/927其他翻譯）皆須由本所工程師翻譯/處理，不得委外。
              'Memo by Lydia 2025/06/05  「m_strBASF」改為「m_str所內譯」
              'Modified by Lydia 2025/07/01 增加例外案件設定InStr(m_str所內譯例外, Replace(strExc(0), "-", "")) = 0 And
              If InStr(m_str所內譯例外, Replace(strExc(0), "-", "")) = 0 And (InStr(m_str所內譯, "" & MSHFlexGrid1.TextMatrix(intI, colAppNo1)) > 0 Or InStr(m_str所內譯, "" & MSHFlexGrid1.TextMatrix(intI, colFAno)) > 0) Then
                  strExc(5) = MSHFlexGrid1.Text
                  MSHFlexGrid1.Text = Replace(strExc(5), strExc(1), "⊕" & strExc(1))
              Else
              'end 2023/04/19
                 If Val(strExc(2)) >= 6 Then
                    If PUB_GetTransFeeRate(strExc(1), strExc(2), strExc(3), strExc(4)) > 30 Then
                        strExc(5) = MSHFlexGrid1.Text
                        MSHFlexGrid1.Text = Replace(strExc(5), strExc(1), "＊" & strExc(1))
                    End If
                 End If
              End If 'Added by Lydia 2023/04/19
           End If
           'end 2021/07/29
        Next iRow
   Next intI
   MSHFlexGrid1.Visible = True
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Lydia 2018/09/25
   Set frm060122 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
Dim intRow As Integer
Dim lngColor As Long
   With MSHFlexGrid1
       If .MouseRow > 0 Then
          intRow = .MouseRow
          .row = intRow
          .col = cFixed
          lngColor = .CellBackColor
          GridClick MSHFlexGrid1, intRow, 0, 0, cFixed, "V", lngColor
       End If
   End With
End Sub

Private Sub MSHFlexGrid1_DblClick()
    'Modified by Lydia 2019/12/23 +權限管制
    'If cmdOK(3).Visible = True Then
    If cmdOK(4).Visible = True And m_bUpdate = True Then
        Call cmdok_Click(4) '分案
    Else
        Call cmdok_Click(1)  '基本檔
    End If
End Sub

Public Sub PubShowNextData()
Dim inX As Integer, inY As Integer
Dim Str01 As String  'Memo by Lydia 2018/09/25 本所案號
Dim lngColor As Long
Dim mUserNo As String '預設翻譯人員
Dim m_CP09 As String, m_CP48 As String 'Added by Lydia 2018/09/25 新案翻譯收文號,承辦期限

On Error GoTo ErrHand01
    
    '查詢
    If cmdState = 0 Then
         If bolNextDone = True And Me.Tag <> "" And mType <> "" Then  '下一畫面作業完成,呼叫Outlook
              Call ProcEmail(mType, Me.Tag, mTransKind, nKeyNo)
         End If
         Call QueryData
         Exit Sub
    End If
    
    For inX = 1 To MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.row = inX
       MSHFlexGrid1.col = 0
       If Trim(MSHFlexGrid1.Text) = "V" Then
           MSHFlexGrid1.Text = ""
           MSHFlexGrid1.col = 0
           MSHFlexGrid1.CellBackColor = MSHFlexGrid1.BackColor
           MSHFlexGrid1.col = cFixed + 1
           lngColor = MSHFlexGrid1.CellBackColor
           For inY = 1 To cFixed
               MSHFlexGrid1.col = inY
               MSHFlexGrid1.CellBackColor = lngColor
           Next inY
           '本所案號
           Str01 = Trim(MSHFlexGrid1.TextMatrix(inX, colCase))
           m_CP09 = Trim(MSHFlexGrid1.TextMatrix(inX, colC1CP09))   'Added by Lydia 2018/09/25 中說收文號
           mTransKind = "" 'Added by Lydia 2024/03/11
           If cmdState = 1 Or cmdState = 2 Then  '共同查詢之畫面
              'Added by Lydia 2024/05/06 增加對執行”基本資料”和”卷宗區”的限閱案件控制
              If Pub_StrUserSt03 <> "M51" Then
                 If PUB_ChkCufaByCaseNo(strUserNum, Me.Name, Replace(Str01, "-", ""), "1") = False Then
                    Exit For
                 End If
              End If
              'end 2024/05/06
              If fnSaveParentForm(Me) = False Then
                  Me.Enabled = True
                  Exit Sub
              End If
           End If
           '急件
           If Trim(Val(Str01)) = Str01 Or InStr(Str01, "▲") > 0 Then
                If Val(Left(Str01, 1)) > 0 Then
                    mType = "0"
                Else
                    mType = "1"
                End If
           Else  '一般翻譯
                If InStr(Str01, "●") > 0 Then
                      mType = "3" '未提申先翻譯
                ElseIf "" & MSHFlexGrid1.TextMatrix(inX, colTF30) <> "" And "" & MSHFlexGrid1.TextMatrix(inX, colTF30) <> "Y" Then
                      mType = "4" '英文參考本
                Else
                      mType = "2"  '一般翻譯
                End If
           End If
           If InStr(Str01, "(") > 0 Then Str01 = Mid(Str01, 1, InStr(Str01, "(") - 1) 'Added by Lydia 2018/11/20
           Str01 = Pub_RplStr(Str01) '清掉符號
           If InStr(Str01, "-") > 0 Then
               If InStrRev(Str01, "-") < 6 Then Str01 = Str01 & "-0-00"
           End If
           
           '是否只能上班翻譯
           'Modified by Lydia 2018/09/25 +相似度
           'If "" & MSHFlexGrid1.TextMatrix(inX, colDc) & MSHFlexGrid1.TextMatrix(inX, colFc) <> "" Then
           'Modified by Lydia 2019/08/13 自2019年8月15日起實施，亦即自當日起交稿案件一律以調整後費率計算; 並且取消折扣案件之限制
           'If "" & MSHFlexGrid1.TextMatrix(inX, colDc) & MSHFlexGrid1.TextMatrix(inX, colFc) & MSHFlexGrid1.TextMatrix(inX, colTF19) <> "" Then
           If strSrvDate(1) < "20190815" And "" & MSHFlexGrid1.TextMatrix(inX, colDc) & MSHFlexGrid1.TextMatrix(inX, colFc) & MSHFlexGrid1.TextMatrix(inX, colTF19) <> "" Then
                mTransKind = "Y,"
                '列出說明有折扣或固定報價
                If "" & MSHFlexGrid1.TextMatrix(inX, colDc) <> "" Then
                     mTransKind = mTransKind & "有折扣,"
                End If
                If "" & MSHFlexGrid1.TextMatrix(inX, colFc) <> "" Then
                     mTransKind = mTransKind & "固定報價,"
                End If
                'Added by Lydia 2018/09/25 有相似度
                If Val("" & MSHFlexGrid1.TextMatrix(inX, colTF19)) > 0 Then
                     mTransKind = mTransKind & "相似度,"
                End If
                'end 2018/09/25
           End If
          
           If mType = 0 Or mType = 1 Then 'Added by Lydia 2018/09/25 判斷欲認領人員的來源
               mUserNo = "" & MSHFlexGrid1.TextMatrix(inX, colTcn15)
            'Added by Lydia 2018/09/25
           Else
               mUserNo = "" & MSHFlexGrid1.TextMatrix(inX, colTFA04)
           End If
           'end 2018/09/25
           
           If Replace(Str01, "-", "") <> "" Then
                Select Case cmdState
                    Case 1  '基本檔
                         If mType = "0" Then
                                MsgBox "急件翻譯尚未立案 !!", vbCritical
                                fnCloseAllFrm100
                         Else
                                frm100101_3.Show
                                frm100101_3.Tag = Str01
                                frm100101_3.StrMenu
                         End If
                         
                    Case 2   '卷宗區
                         If mType = "0" Then
                                MsgBox "急件翻譯尚未立案 !!", vbCritical
                                fnCloseAllFrm100
                         Else
                                frm100101_L.m_strKey = Str01
                                frm100101_L.SetParent Me
                                If frm100101_L.QueryData = True Then
                                   frm100101_L.Show
                                   Me.Hide
                                End If
                         End If
                         
                    Case 3 '新案建檔
                         If mType = "0" Then
                                MsgBox "急件翻譯尚未立案 !!", vbCritical
                         'Modified by Lydia 2018/08/15 改成傳變數
                         'Else
                         '       Call frm060102.SetParent(Me, Replace(Str01, "-", ""))
                         '       frm060102.Show
                         ElseIf TypeName(Tmpfrm060102) <> "Nothing" Then
                                Call Tmpfrm060102.SetParent(Me, Replace(Str01, "-", ""))
                                Tmpfrm060102.Show
                                Me.Hide
                         End If
                    Case 4 '分案
                         If mType = "0" Then
                                'Modified by Lydia 2018/08/15 改成傳變數
                                'Call frm060504.SetParent(Me, Trim(MSHFlexGrid1.TextMatrix(inX, 2)))
                                If TypeName(Tmpfrm060504) <> "Nothing" Then
                                    Call Tmpfrm060504.SetParent(Me, Trim(MSHFlexGrid1.TextMatrix(inX, 2)))
                                    Tmpfrm060504.Show
                                    Me.Hide
                                End If
                         Else
                                'Added by Lydia 2021/08/27 發生前一畫面資料保留; ex.FCP-065514有指定日期會發email,因為先分案造成後續案都比照辦理
                                If PUB_CheckFormExist("frm060101_1") Then
                                    MsgBox "請先關閉〔外專分案〕畫面！"
                                    Exit Sub
                                End If
                                If PUB_CheckFormExist("frm060101_3") Then
                                    MsgBox "請先關閉〔外專分案-FMP案〕畫面！"
                                    Exit Sub
                                End If
                                'end 2021/08/27
                                '排除
                                If "" & Trim(MSHFlexGrid1.TextMatrix(inX, colTF29)) = "Y" Then
                                      MsgBox Str01 & "尚待比對 !", vbCritical
                                      Exit Sub
                                End If
                                If "" & Trim(MSHFlexGrid1.TextMatrix(inX, colTF30)) = "Y" Then
                                      MsgBox Str01 & "尚待英文本 !", vbCritical
                                      Exit Sub
                                End If
                                'Added by Lydia 2018/08/24
                                If "" & Trim(MSHFlexGrid1.TextMatrix(inX, colTF34)) = "Y" Then
                                      MsgBox Str01 & "暫不翻譯 !", vbCritical
                                      Exit Sub
                                End If
                                'Added by Lydia 2023/06/06 針對特定Y編號+X82908000彈提醒
                                If Mid(Trim("" & MSHFlexGrid1.TextMatrix(inX, colAppNo1)), 1, 8) = "X8290800" And InStr("Y55451000,Y55450000,Y45268000,Y55456000", Mid(Trim("" & MSHFlexGrid1.TextMatrix(inX, colFAno)), 1, 8)) > 0 Then
                                    '「以外」的內容是指工程師和承辦要另外提供相關文件
                                    MsgBox "此客戶只需翻譯常見內容「以外」的內容，務必將翻譯對照文件一併交譯者。"
                                End If
                                'end 2023/06/06
                                
                                'Modified by Lydia 2018/09/25
                                'Me.Tag = Trim(MSHFlexGrid1.TextMatrix(inX, colC1CP09))
                                Me.Tag = m_CP09
                                'Modified by Lydia 2018/08/15 改成傳變數
'                                If Left(Str01, 1) = "P" Then 'FMP案
'                                     Call frm060101_3.SetParent(Me, 1, Replace(Str01, "-", ""), Trim(MSHFlexGrid1.TextMatrix(inX, colC1CP09)), mTransKind)
'                                     frm060101_3.Show
'                                ElseIf Left(Str01, 1) = "F" Then 'FCP案
'                                     Call frm060101_1.SetParent(Me, 1, Replace(Str01, "-", ""), Trim(MSHFlexGrid1.TextMatrix(inX, colC1CP09)), mTransKind)
'                                     frm060101_1.Show
'                                End If
'                                Me.Hide
                                'Added by Lydia 2024/03/08 因應內專支援機械組OA (含P案) 收文: 927其他翻譯
                                If Left(Trim(MSHFlexGrid1.TextMatrix(inX, colCase)), 1) = "＃" Then
                                    mTransKind = "＃"
                                End If
                                'end 2024/03/08
                                If Left(Str01, 1) = "P" And TypeName(Tmpfrm060101_3) <> "Nothing" Then  'FMP案
                                     'Modified by Lydia 2018/09/25 MSHFlexGrid1.TextMatrix(inX, colC1CP09)=>m_cp09
                                     'Modified by Lydia 2022/07/05 +交稿期限
                                     'Call Tmpfrm060101_3.SetParent(Me, 1, Replace(Str01, "-", ""), m_CP09, mTransKind)
                                     Call Tmpfrm060101_3.SetParent(Me, 1, Replace(Str01, "-", ""), m_CP09, mTransKind, Trim(MSHFlexGrid1.TextMatrix(inX, colTF26)))
                                     Tmpfrm060101_3.Show
                                     Me.Hide
                                ElseIf Left(Str01, 1) = "F" And TypeName(Tmpfrm060101_1) <> "Nothing" Then  'FCP案
                                     'Modified by Lydia 2018/09/12 +交稿期限
                                     'Call Tmpfrm060101_1.SetParent(Me, 1, Replace(Str01, "-", ""), Trim(MSHFlexGrid1.TextMatrix(inX, colC1CP09)), mTransKind)
                                      'Modified by Lydia 2018/09/25 MSHFlexGrid1.TextMatrix(inX, colC1CP09)=>m_cp09
                                     Call Tmpfrm060101_1.SetParent(Me, 1, Replace(Str01, "-", ""), m_CP09, mTransKind, Trim(MSHFlexGrid1.TextMatrix(inX, colTF26)))
                                     Tmpfrm060101_1.Show
                                     Me.Hide
                                End If
                                'end 2018/08/15
                         End If

                    Case 5 'E-Mail
                         '是否為工程師認領Email
                         If m_bUpdate = False Then
                              mType = ""
                         'Remove by Lydia 2018/09/25 與工程師認翻譯作業分成不同地方點選
                         'ElseIf Pub_StrUserSt03 = "M51" Then
                         '     If MsgBox("是否為工程師認領Email ? ", vbYesNo + vbDefaultButton2) = vbYes Then
                         '          mType = ""
                        '      End If
                         'end 2018/09/25
                         End If
                         '急件翻譯立案後,改抓收文號
                         If mType = "1" Then
                              'Modified by Lydia 2018/09/25
                              'Str01 = Trim(MSHFlexGrid1.TextMatrix(inX, colC1CP09))
                              Str01 = m_CP09
                         End If
                         'Modified by Lydia 2018/09/25 傳中說收文號m_cp09
                         Call ProcEmail(mType, Str01, mTransKind, mUserNo, Trim(MSHFlexGrid1.TextMatrix(inX, colTF26)), Trim(MSHFlexGrid1.TextMatrix(inX, colTF32)), m_CP09, Trim(MSHFlexGrid1.TextMatrix(inX, colC1CP48)))
                         'Added by Lydia 2018/09/25 重整Grid
                         If mType = "" Then
                             Call QueryData
                         End If
                         'end 2018/09/25
                    Case 6 'Added by Lydia 2018/09/28 工程師認翻譯-主管確認
                         Call frm090906.SetParent(Me, Replace(Str01, "-", ""), m_CP09, Trim(Left(Combo2.Text, 6)), m_Grp, mTransKind)
                         frm090906.Show
                         Me.Hide
                    'end 2018/09/28
                End Select
           End If
           Exit For
       End If
    Next inX
    Me.Enabled = True

    Exit Sub

ErrHand01:
    If Err.Number <> 0 Then
         MsgBox Err.Description
         Resume Next
    End If

End Sub

'Modified by Lydia 2018/09/25 +iCP09,iCP48
Private Function ProcEmail(ByVal iType As String, ByVal iKeyNo As String, ByVal iTKind As String, Optional ByVal iUserNo As String, Optional ByVal iLdate As String, Optional ByVal iLdate2 As String, Optional ByVal iCp09 As String = "", Optional ByVal iCp48 As String = "")
'iKeyNo　本所案號/追蹤號
'iTKind   是否只能上班翻譯(=Y)
'iUserNo 預設翻譯人員
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim pbolDone As Boolean
Dim strSavePath As String
'Added by Lydia 2018/09/25
Dim intJ As Integer, strKind As String
Dim strCase(1 To 4) As String
Dim dbTfRate As Double, bolIsHigher As Boolean  'Added by Lydia 2021/07/29 判斷翻譯費折扣率＞30%
Dim strAdd As String 'Added by Lydia 2022/04/19

     'Added by Lydia 2018/09/25 工程師認領通知信
    If iType = "" Then
        'Modified by Lydia 2019/06/19 Sharon表示沒有這件事的印象
        'If m_Grp = "3" Then
        '     MsgBox "日文組只能在命名作業階段認領翻譯 !", vbInformation
        '     Exit Function
        'ElseIf iUserNo <> "" Then
        If iUserNo <> "" Then
        'end 2019/06/19
             If iUserNo = Trim(Left(Combo2.Text, 6)) Then
                 MsgBox "已提出認領翻譯，尚未經過主管確認 !", vbInformation
             Else
                 'Modified by Lydia 2019/06/19
                 'MsgBox "已有他人認領翻譯，若欲認領請向工程師主管申請 !", vbInformation
                 MsgBox "已有他人認領翻譯 !", vbInformation
             End If
             Exit Function
        ElseIf m_Grp = "" Then
             MsgBox "非工程師不可認領翻譯 !", vbInformation
             Exit Function
        Else
             'Remove by Lydia 2019/09/12 取消下班翻譯控制
             'Remark by Lydia 2019/11/25 還原; 只取消"不能下班翻譯"的控制
             'intJ = MsgBox(IIf(iTKind <> "", "※本案不能下班翻譯" & vbCrLf & vbCrLf, "") & "是否上班翻譯？" & vbCrLf & "選擇""是""：上班翻譯" & vbCrLf & "選擇""否""：下班翻譯" & vbCrLf & "選擇""取消""：不認領", vbInformation + vbYesNoCancel + vbDefaultButton1, "選擇上班/下班翻譯")
             'Modifeid by Lydia 2021/07/30 翻譯費折扣率＞30%客戶只能上班譯; 可排除特別客戶
             strExc(1) = ""
             Call ChgCaseNo(Replace(iKeyNo, "-", ""), strCase)
             dbTfRate = PUB_GetTransFeeRate(strCase(1), strCase(2), strCase(3), strCase(4), , bolIsHigher, True)
             '控制翻譯費折扣率＞30%客戶案件之承辦人只能為所內人員上班譯編號。
             If dbTfRate > 30 Then
                 strExc(1) = "B"
             ElseIf bolIsHigher = True Then  '折扣率＞30%但是例外控制的客戶
                  '不受限
             End If
             'Modified by Lydia 2022/04/19 (例外認領)因為常有例外案件，所以比照過去email溝通模式(和Sharon討論)
             'If strExc(1) = "B" Then
             '     intJ = MsgBox(iKeyNo & "只能上班翻譯，是否認領翻譯？" & vbCrLf & "選擇""是""：上班翻譯" & vbCrLf & "選擇""否""：不認領", vbInformation + vbYesNo + vbDefaultButton1, "認領翻譯")
             '     If intJ = 7 Then 'No
             '          Exit Function
             '     Else
             '          strKind = "B上班"
             '     End If
             'Else
             'end 2021/07/30
             '     intJ = MsgBox("是否上班翻譯？" & vbCrLf & "選擇""是""：上班翻譯" & vbCrLf & "選擇""否""：下班翻譯" & vbCrLf & "選擇""取消""：不認領", vbInformation + vbYesNoCancel + vbDefaultButton1, "選擇上班/下班翻譯")
             strAdd = ""
             If strExc(1) = "B" Then
                 strAdd = "※本案不能下班翻譯"
                 mTransKind = "Y,折扣率＞30%"
             End If
             intJ = MsgBox(IIf(strAdd <> "", strAdd & vbCrLf, "") & "是否上班翻譯？" & vbCrLf & "選擇""是""：上班翻譯" & vbCrLf & "選擇""否""：下班翻譯" & vbCrLf & "選擇""取消""：不認領", vbInformation + vbYesNoCancel + vbDefaultButton1, "選擇上班/下班翻譯")
             'end 2022/04/19
                  If intJ = 2 Then 'Cancel
                     Exit Function
                  ElseIf intJ = 7 Then 'No
                       strKind = "A下班"
                  Else
                       strKind = "B上班"
                  End If
             'End If 'Added by Lydia 2021/07/30 'Mark by Lydia 2022/04/19
        End If

        If strKind <> "" Then
             'Added by Lydia 2021/04/14 外專翻譯承辦及核稿期限控管：
             '工程師認領翻譯時，查詢該認領人員，新案翻譯未上完稿日案件,請彈提醒: 尚未完稿案件FCPxxxx , 承辦期限
             strExc(4) = Pub_GetEngEP09List(Trim(Left(Combo2.Text, 6)))
             If strExc(4) <> "" Then
                 MsgBox "尚未完稿案件：" & strExc(4), vbCritical
             End If
             'end 2021/04/14
             'Call ChgCaseNo(Replace(iKeyNo, "-", ""), strCase) 'Remove by Lydia 2021/07/30 移到前方
             pbolDone = False
             'Added by Lydia 2019/06/19 認下班翻譯時，若是該案為不可下班翻譯之案件，會自動開啟系統Email介面，提供工程師輸入原因(Email主旨後面加上呈報主管)。
             'Remove by Lydia 2019/09/12 取消下班翻譯控制
             If iTKind <> "" And Left(strKind, 1) = "A" Then
ReEmail:
                  frm880019.txtSubject = strCase(1) & "-" & strCase(2) & IIf(strCase(3) & strCase(4) <> "000", "-" & strCase(3) & "-" & strCase(4), "") & "欲認領新案翻譯人員(" & Trim(Mid(Combo2.Text, 7)) & IIf(strKind <> "", "-" & Mid(strKind, 2), "") & ") ，呈報主管"
                  frm880019.txtContent = "※本案不能下班翻譯：" & Mid(iTKind, 3, Len(iTKind) - 3) & vbCrLf & vbCrLf & _
                                                     "欲承接下班翻譯，請呈報主管"
                  frm880019.txtReceiver = m_GrpMan
                  frm880019.cmdAttach.Visible = False
                  frm880019.SetParent Me
                  frm880019.Show vbModal
                  pbolDone = frm880019.m_bolDone '是否傳送成功
                  Unload frm880019
                  If pbolDone = False Then
                      If MsgBox("送信失敗，是否重新Email？ ", vbCritical + vbYesNo + vbDefaultButton2, "呈報主管") = vbYes Then
                          GoTo ReEmail
                      Else
                          Exit Function
                      End If
                  End If
             End If
             'end 2019/06/19
             'end 2019/09/12
             cnnConnection.BeginTrans
                  'Added by Lydia 2019/06/19 先刪檔
                  strSql = "delete from transfeeassign where tfa01='" & iCp09 & "' and tfa04=" & CNULL(Trim(Left(Combo2.Text, 6)))
                  cnnConnection.Execute strSql
                  'end 2019/06/19
                  strSql = "insert into transfeeassign (tfa01,tfa02,tfa03,tfa04,tfa05) select '" & iCp09 & "', to_char(sysdate, 'YYYYMMDD') , to_char(sysdate, 'HH24MISS'), '" & Trim(Left(Combo2.Text, 6)) & "'," & CNULL(Left(strKind, 1)) & " from dual "
                  cnnConnection.Execute strSql
                  'Remove by Lydia 2019/09/12 取消下班翻譯控制
                  'Remark by Lydia 2019/11/25 還原; 只取消"不能下班翻譯"的控制
                  'If Not (iTKind <> "" And Left(strKind, 1) = "A") Then 'Added by Lydia 2019/06/19 排除認下班翻譯時，若是該案為不可下班翻譯之案件; 因為前面已發過Email
                      strExc(1) = "": strExc(2) = "": strExc(6) = ""
                      '通知工程師主管，CC副本給自己
                      strExc(1) = strCase(1) & "-" & strCase(2) & IIf(strCase(3) & strCase(4) <> "000", "-" & strCase(3) & "-" & strCase(4), "") & "欲認領新案翻譯人員(" & Trim(Mid(Combo2.Text, 7)) & IIf(strKind <> "", "-" & Mid(strKind, 2), "") & ") "
                      '提醒只能上班翻譯
                      'Remove by Lydia 2019/11/25
                      'If iTKind <> "" And Left(strKind, 1) = "A" Then
                      '      strExc(1) = strExc(1) & "，本案不能下班翻譯"
                      '      strExc(2) = "※本案不能下班翻譯：" & Mid(iTKind, 3, Len(iTKind) - 4) & vbCrLf
                      'End If
                      'end 2019/11/25
                      'Added by Lydia 2022/04/19 (例外認領)因為常有例外案件，所以比照過去email溝通模式(和Sharon討論)
                      If strAdd <> "" And Left(strKind, 1) = "A" Then
                            strExc(1) = strExc(1) & strAdd
                            strExc(2) = strAdd & "：折扣率＞30%" & vbCrLf & "欲承接下班翻譯，請呈報主管。" & vbCrLf
                      End If
                      'end 2022/04/19
                      strExc(2) = strExc(2) & "若主管欲確認本案之翻譯人員，請到""外專翻譯分案-認翻譯""進行主管確認。"
                      strExc(3) = Trim(Left(Combo2.Text, 6))
                      If strExc(3) <> strUserNum Then strExc(3) = strExc(3) & ";" & strUserNum '若代人認領
                      'Added by Lydia 2021/04/15 日文組認領通知email：收件者改為工程師之第二級主管(副理)，並CC給王協理+本人。
                      If m_Grp = "3" Then
                            strExc(6) = PUB_GetFCPEngSup(strExc(3), True)
                            strExc(5) = Replace(Replace(strExc(6), m_GrpMan, ""), ";;", "") '收件人
                            If Len(strExc(5)) >= 5 Then
                                strExc(6) = ";" & m_GrpMan 'CC
                            Else
                                strExc(5) = m_GrpMan
                                strExc(6) = ""
                            End If
                            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                 " values( '" & strUserNum & "','" & strExc(5) & "',to_char(sysdate,'yyyymmdd')" & _
                                 ",to_char(sysdate,'hh24miss'),'" & strExc(1) & "','" & strExc(2) & "'," & CNULL(strExc(3) & strExc(6)) & ")"
                      Else  '英文組認領通知email:收件者為各組主管 , CC給本人
                      'end 2021/04/15
                            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                 " values( '" & strUserNum & "','" & m_GrpMan & "',to_char(sysdate,'yyyymmdd')" & _
                                 ",to_char(sysdate,'hh24miss'),'" & strExc(1) & "','" & strExc(2) & "'," & CNULL(strExc(3)) & ")"
                      End If 'Added by Lydia 2021/04/15
                      cnnConnection.Execute strSql
                  'End If 'end 2019/06/19
                 
             cnnConnection.CommitTrans
        End If
    'end 2018/09/25
    Else
         'Modified by Lydia 2023/08/18 改資料夾
         'strSavePath = App.path & "\FCmail2"
         'If Dir(strSavePath, vbDirectory) = "" Then
         '     MkDir strSavePath
         'Else
         '     If Dir(strSavePath & "\*.*") <> "" Then
         '          Kill strSavePath & "\*.*"
         '     End If
         'End If
         strSavePath = m_AttachPath
         Call Pub_ChkExcelPath(strSavePath)
         'end 2023/08/18
         
         Select Case iType
             Case "0" '0.急件翻譯確認信(未立案)
                    strExc(1) = IIf(iUserNo <> "", "  目前預設為" & iUserNo & " " & GetStaffName(iUserNo), "")
JumpReInput:
                    'Modified by Lydia 2025/03/13 改用模組取得
                    'strExc(2) = UCase(InputBox("請選擇國外譯者或輸入員工編號：" & strExc(1) & vbCrLf & "(1: 舜禹 2: 捷恩凱 3: 迅達" & ")", "急件翻譯確認信(未立案)", IIf(iUserNo <> "", SetF51Order(iUserNo), "")))
                    strExc(2) = UCase(InputBox("請選擇國外譯者或輸入員工編號：" & strExc(1) & vbCrLf & "(" & Pub_SetF51Order("F", "3") & "), 急件翻譯確認信(未立案)", IIf(iUserNo <> "", SetF51Order(iUserNo), "")))
                    If strExc(2) = "" Then
                        If MsgBox("未選擇國外譯者或輸入員工編號，是否繼續發Email？", vbYesNo + vbDefaultButton1) = vbYes Then
                             GoTo JumpReInput
                        End If
                    Else
                        'Modified by Lydia 2025/03/13 改用模組取得
                        'If InStr("1,2,3", strExc(2)) = 0 Or (Val(strExc(2)) = 0 And GetStaffName(strExc(2)) = "") Then
                        If InStr(Pub_SetF51Order("F", "3"), strExc(2)) = 0 Or (Val(strExc(2)) = 0 And GetStaffName(strExc(2)) = "") Then
                            If MsgBox("輸入非員工編號，是否繼續發Email？", vbYesNo + vbDefaultButton1) = vbYes Then
                                 GoTo JumpReInput
                            End If
                        End If
                    End If
                    If strExc(2) <> "" Then
                        'Modified by Lydia 2023/08/18  App.path=>m_AttachPath
                        Call PUB_Translate_SendMail(iType, strSavePath, m_AttachPath & "\" & m2FileName, iKeyNo, IIf(Val(strExc(2)) > 0, SetF51Order(strExc(2)), strExc(2)))
                    End If
             Case Else
                    '1.急件翻譯(已立案) 第2次發信, 2.一般翻譯 3.未提申先翻譯 4.英文參考本
                    If iUserNo = "" Then
                         MsgBox "尚未有承辦人(翻譯人員) !", vbCritical
                    Else
                         'Modified by Lydia 2023/08/18  App.path=>m_AttachPath
                         Call PUB_Translate_SendMail(iType, strSavePath, m_AttachPath & "\" & m2FileName, iKeyNo, iUserNo)
                    End If
         End Select
    End If
End Function

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow MSHFlexGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MSHFlexGrid1.col = nCol
   MSHFlexGrid1.row = nRow
   If Me.MSHFlexGrid1.row < 1 And Me.MSHFlexGrid1.Text <> "V" Then
      If InStr("相似度", Me.MSHFlexGrid1.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.MSHFlexGrid1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.MSHFlexGrid1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.MSHFlexGrid1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.MSHFlexGrid1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

'國外譯者的順序和員工編號
Private Function SetF51Order(ByVal pNo As String) As String
Dim strMid As String

    SetF51Order = ""
    If pNo = "" Then Exit Function
    
    If Val(pNo) > 0 Then
        strMid = Pub_GetTct27ID("", pNo, "")
    Else
        'Modified by Lydia 2025/03/13 改用模組取得
        'Select Case pNo
        '      Case 外翻_舜禹: strMid = "1"
        '      Case 外翻_捷恩凱: strMid = "2"
        '      Case 外翻_迅達: strMid = "3"
        '      Case Else: strMid = "9"
        'End Select
        strMid = Pub_SetF51Order("", pNo)
        If strMid <> pNo Then strMid = ""
        'end 2025/03/13
    End If
    SetF51Order = strMid
End Function

'Added by Lydia 2018/10/01
Private Sub Check1_Click()
    Call cmdok_Click(0)
End Sub

Private Sub Check2_Click()
    Call cmdok_Click(0)
End Sub
