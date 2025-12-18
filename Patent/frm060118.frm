VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060118 
   BorderStyle     =   1  '單線固定
   Caption         =   "程序大項工作整批發文"
   ClientHeight    =   5856
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9384
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5856
   ScaleWidth      =   9384
   Begin VB.CommandButton CmdPass 
      BackColor       =   &H00C0FFFF&
      Caption         =   "假發文(&B)"
      Height          =   400
      Left            =   3690
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   480
      Width           =   1185
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1260
      TabIndex        =   15
      Text            =   "Combo2"
      Top             =   420
      Width           =   2265
   End
   Begin VB.TextBox Txt1 
      Height          =   270
      Index           =   0
      Left            =   4530
      MaxLength       =   7
      TabIndex        =   8
      Top             =   60
      Width           =   900
   End
   Begin VB.ComboBox cboPrinter 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5505
      TabIndex        =   6
      Top             =   5505
      Width           =   3675
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印清單(&P)"
      Height          =   400
      Left            =   5595
      TabIndex        =   5
      Top             =   60
      Width           =   1185
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "發文(&S)"
      Height          =   400
      Left            =   7696
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6838
      TabIndex        =   2
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8550
      TabIndex        =   0
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4395
      Left            =   30
      TabIndex        =   1
      Top             =   1050
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   7747
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "V |E化情形 |核准函發文日 |承辦期限 |定稿日 |管制人 |承辦人 |本所案號 |備註"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1260
      TabIndex        =   12
      Top             =   30
      Width           =   2265
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3995;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "符號說明：●代表銷卷＊代表閉卷"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5160
      TabIndex        =   17
      Top             =   780
      Width           =   3240
   End
   Begin VB.Label Label1 
      Caption         =   "工作項目："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   14
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   13
      Top             =   60
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "按標題V：全選或全取消"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   1
      Left            =   60
      TabIndex        =   11
      Top             =   780
      Width           =   2655
   End
   Begin VB.Label lblCnt2 
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   2130
      TabIndex        =   10
      Top             =   5550
      Width           =   1710
   End
   Begin VB.Label Label2 
      Caption         =   "定稿日期："
      Height          =   180
      Index           =   1
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   7
      Top             =   5535
      Width           =   975
   End
   Begin VB.Label lblCnt 
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   330
      TabIndex        =   4
      Top             =   5550
      Width           =   1710
   End
End
Attribute VB_Name = "frm060118"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/29 Form2.0已修改
'Create by Sindy 2017/1/12
'Memo by Lydia 2019/05/31 原本「告准函整批發文」，更名為「程序大項工作整批發文」
Option Explicit

Dim bolBarShow As Boolean
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim iPage As Integer, iPrint As Integer, PLeft() As Integer
Dim m_iCols As Integer, m_iPrtCols As Integer
Private Const ciFontSize = 12, ciTitleFontSize = 22
Private Const ciStartX = 500, ciColGap = 250, ciStartY = 500
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim strPrinter As String
Dim m_adoRst As ADODB.Recordset
Dim m_stSort As String '排序方式
Dim m_bolSelAll As Boolean 'Add By Sindy 2017/3/1 是否全選或全取消
'Added by Lydia 2019/05/31 要處理的工作大項種類(1.告准函1917、2.專利證書1603、3.公開公報1229、4.專利權消滅1604、5.通知年費逾期1605
'Modified by Lydia 2019/08/16 + 6.期限通知-年費
Dim iKind As String  '1~5 => 2019/08/16 1~6　'2025/06/05 改為1~7
Dim iPty(1 To 7) As String '案件性質 '2019/08/16 5=>6  '2025/06/05 改為6=>7
Dim colCp09 As Integer  'C1總收文號(工作項目)
Dim colCaseNo As Integer 'Added by Lydia 2021/08/30 本所案號
Dim colC2CP14 As Integer 'Added by Lydia 2021/08/30 承辦工程師
Dim colC2CP148, colNA51 As Integer, colPA75 As Integer, colPA26 As Integer  'Added by Lydia 2021/09/27 專利.是否有檢索、NA51(FCP承辦管制),PA75,PA26~PA30


Private Sub SetRst2Grid()
   Set grdDataList.Recordset = m_adoRst
End Sub

Private Sub cmdPrint_Click()
   If grdDataList.Rows - 1 > 0 Then
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      PUB_RestorePrinter cboPrinter.Text
      DoPrint
      PUB_RestorePrinter strPrinter
      Me.Enabled = True
      Screen.MousePointer = vbDefault
   End If
End Sub

Public Sub cmdQuery_Click()
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   doQuery
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSend_Click()
   Screen.MousePointer = vbHourglass
   'Modified by Lydia 2019/06/17 處理方式 : 1
   'doBatch
   Call doBatch("1")
   Screen.MousePointer = vbDefault
End Sub

'Modified by Lydia 2019/06/17 +處理方式pType : 1-發文, 2-假發文
'Private Sub doBatch()
Private Sub doBatch(ByVal pType As String)
Dim iRow As Integer, ii As Integer
Dim bolSelRow As Boolean
Dim strCP09 As String
Dim bolUpdate As Boolean 'Added by Lydia 2019/05/31
Dim intQ As Integer, rsQuery As New ADODB.Recordset 'Added by Lydia 2021/08/30
Dim bolPlus1919 As Boolean 'Added by Morgan 2022/8/17
Dim bolHave926 As Boolean 'Added by Morgan 2024/2/5
Dim strCase(1 To 4) As String 'Added by Lydia 2025/08/22

On Error GoTo ErrHnd
   
   bolSelRow = False
   For iRow = 1 To grdDataList.Rows - 1
      If grdDataList.TextMatrix(iRow, 0) = "V" Then
         bolSelRow = True
      End If
   Next iRow
   If bolSelRow = False Then
      MsgBox "請選取欲發文的資料！", vbExclamation
      Exit Sub
   End If
   
cnnConnection.BeginTrans 'Added by Lydia 2019/05/31
   With grdDataList
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, 0) = "V" And .RowHeight(iRow) > 0 Then
            'Modified by Lydia 2019/05/31 改成變數
            'strCP09 = .TextMatrix(iRow, 9) '告准總收文號
            strCP09 = .TextMatrix(iRow, colCp09) 'C1總收文號
            If strCP09 <> "" Then
               .row = iRow
               grdDataList.col = 0
               grdDataList.Text = ""
               For ii = 0 To grdDataList.Cols - 1
                  grdDataList.col = ii
                  grdDataList.CellBackColor = grdDataList.BackColor
               Next
               .RowHeight(iRow) = 0
               'Modified by Lydia 2019/05/31 將承辦人變更為操作者
               'strSql = "update caseprogress set cp27=" & strSrvDate(1) & _
                        " where cp09='" & strCP09 & "'"
               'Modified by Lydia 2019/06/17 判斷處理方式
               'strSql = "update caseprogress set cp27=" & strSrvDate(1) & ", cp14=" & CNULL(strUserNum) & _
                        " where cp09='" & strCP09 & "'"
               'end 2019/05/31
               strSql = "update caseprogress set cp27=" & IIf(pType = "1", strSrvDate(1), "19221111") & ", cp14=" & CNULL(strUserNum) & _
                        " where cp09='" & strCP09 & "'"
               If bolUpdate = False Then bolUpdate = True 'Added by Lydia 2019/11/21 可能有問題
               cnnConnection.Execute strSql, intI
               
               'Added by Lydia 2025/08/22 取得本所案號
               strExc(0) = "" & .TextMatrix(iRow, colCaseNo)
               If Len(strExc(0)) < 10 Then
                  strExc(0) = strExc(0) & "-0-00"
               End If
               Call ChgCaseNo(Replace(strExc(0), "-", ""), strCase)
               'end 2025/08/22
               
               'Added by Lydia 2021/08/30 告准函發文時，由系統發email【請進行一次核對】給工程師及其主管
               If iKind = 1 And pType = 1 Then '排除假發文
                   If PUB_GetST03("" & .TextMatrix(iRow, colC2CP14)) = "F21" Then
                        'Added by Morgan 2022/8/17 是否有未發文之1919非屬相同創作
                        bolPlus1919 = False
                        strSql = "update caseprogress a set cp64=cp64 where cp09='" & strCP09 & "'" & _
                           " and exists(select * from caseprogress b where cp01=a.cp01 and cp02=a.cp02 and cp03=a.cp03 and cp04=a.cp04" & _
                           " and cp10='1919' and cp158=0 and cp159=0)"
                        cnnConnection.Execute strSql, intI
                        If intI = 1 Then
                           bolPlus1919 = True
                        End If
                        'end 2022/8/17
                        
                        'Added by Morgan 2024/2/5
                        bolHave926 = False
                        strSql = "update caseprogress a set cp64=cp64 where cp09='" & strCP09 & "'" & _
                           " and exists(select * from caseprogress b where cp01=a.cp01 and cp02=a.cp02 and cp03=a.cp03 and cp04=a.cp04" & _
                           " and cp10='926' and cp159=0)"
                        cnnConnection.Execute strSql, intI
                        If intI = 1 Then
                           bolHave926 = True
                        End If
                        'end 2024/2/5
                        
                        strExc(0) = "": strExc(1) = ""
                        strExc(0) = "" & .TextMatrix(iRow, colC2CP14)
                        'Added by Morgan 2024/5/8
                        '若承辦工程師為內專工程師，請協助改寄對接的外專工程師主任
                        If Mid(strExc(0), 4, 1) = "9" Then
                           strExc(0) = PUB_GetFCPProSup(strExc(0))
                        End If
                        'end 2024/5/8
                        strExc(1) = PUB_GetFCPEngSup(strExc(0), True) 'CC: 工程師主管(副理)
                        strExc(5) = strExc(1) 'Added by Lydia 2024/06/25
                        strExc(1) = strExc(1) & IIf(strExc(1) <> "", ";backup", "backup")  '+backup匯入卷宗區
                        'Added by Lydia 2023/09/19 若承辦工程師離職，寄送對象改為副理並CC:backup，內文加註
                        strExc(9) = ""
                        'Modified by Lydia 2024/03/11 承辦工程師已離職，【核對已准專利】進度承辦人掛工程師主管（副理）
                        'If GetStaffName(strExc(0)) = "" Then
                        '   strExc(9) = vbCrLf & vbCrLf & "原承辦工程師為：" & GetStaffName(strExc(0), True) & vbCrLf & _
                        '               "請重新分案 , 謝謝"
                        '   strExc(0) = Replace(strExc(1), ";backup", "")
                        '   strExc(1) = "backup"
                        'End If
                        'end 2023/09/19
                        If strExc(5) = "" & .TextMatrix(iRow, colC2CP14) Or Mid("" & .TextMatrix(iRow, colC2CP14), 4, 1) = "9" Or GetStaffName("" & .TextMatrix(iRow, colC2CP14)) = "" Then 'Added by Lydia 2024/06/25 排除已重新分案的狀況;ex.FCP-70724在6/21已分案
                           strSql = "select b.cp09,b.cp64 from caseprogress a, caseprogress b where a.cp09='" & strCP09 & "'" & _
                                    "  and a.cp43=b.cp43(+) and b.cp10='926' and b.cp159=0 and instr(b.cp64,'原承辦工程師為：') > 0 "
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                           If intI = 1 Then
                              strExc(9) = vbCrLf & vbCrLf & Mid("" & RsTemp.Fields("cp64"), InStr("" & RsTemp.Fields("cp64"), "原承辦工程師為："), Len("原承辦工程師為：") + 4) & vbCrLf & _
                                          "請重新分案 , 謝謝"
                              strExc(0) = Replace(strExc(1), ";backup", "")
                              strExc(1) = "backup"
                           End If
                        End If  'Added by Lydia 2024/06/25
                        'end 2024/03/11
                        
                        'Modified by Morgan 2024/2/5
                        ''Added by Morgan 2022/8/17
                        'If bolPlus1919 Then
                        '   strExc(2) = "【請進行一次核對+非屬相同創作告代】Our Ref: " & .TextMatrix(iRow, colCaseNo) & "[INCOM.1917]"
                        '   strExc(3) = "1. 您好，本案已告准，請進行一次核對，核對完成請通知主管銷核對已准專利之承辦期限，若無核對已准專利(不二核)則不須銷期限，謝謝。" & vbCrLf & _
                        '               "2. 函中提到「本案係於申請時聲明同時申請發明及新型專利，惟經審查認為，該聲明主張因非屬相同創作，不適用專利法第32條規定。」，已收文非屬相同創作，請報告客戶。"
                        'Else
                        ''end 2022/8/17
                        '   strExc(2) = "【請進行一次核對】Our Ref: " & .TextMatrix(iRow, colCaseNo) & "[INCOM.1917]"
                        '   strExc(3) = "您好，本案已告准，請進行一次核對，核對完成請通知主管銷核對已准專利之承辦期限，若無核對已准專利(不二核)則不須銷期限，謝謝。"
                        'End If 'Added by Morgan 2022/8/17
                        If bolHave926 Then
                           strExc(2) = "【請至工作進度資料維護啟動歷程進行一次核對" & IIf(bolPlus1919, "+非屬相同創作告代", "") & "】Our Ref: " & .TextMatrix(iRow, colCaseNo) & "[INCOM.1917]"
                           strExc(3) = "您好，本案已告准，請至工作進度資料維護啟動歷程進行一次核對，交主管核判，謝謝。"
                        Else
                           strExc(2) = "【請進行一次核對" & IIf(bolPlus1919, "+非屬相同創作告代", "") & "】Our Ref: " & .TextMatrix(iRow, colCaseNo) & "[INCOM.1917]"
                           strExc(3) = "您好，本案已告准，請進行一次核對，因無核對已准專利(不二核)，不須啟動歷程作業，謝謝。"
                        End If
                        
                        If bolPlus1919 Then
                           strExc(3) = "1. " & strExc(3) & vbCrLf & _
                                       "2. 函中提到「本案係於申請時聲明同時申請發明及新型專利，惟經審查認為，該聲明主張因非屬相同創作，不適用專利法第32條規定。」，已收文非屬相同創作，請報告客戶。"
                        End If
                        'end 2024/2/5
                        
                        'Added by Morgan 2024/3/5 機械組案件主旨都加【機械設計組】--Sharon
                        If .TextMatrix(iRow, PUB_MGridGetId("PA150", grdDataList)) = "4" Then
                           strExc(2) = "【機械設計組】" & strExc(2)
                        End If
                        'end 2024/3/5
                        
                        'Modified by Lydia 2023/09/19 +加註strExc(9)
                        strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                    " values( '" & strUserNum & "','" & strExc(0) & "',to_char(sysdate,'yyyymmdd')" & _
                                    ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "' ,'" & ChgSQL(strExc(3) & strExc(9)) & "' ,'" & strExc(1) & "' )"
                        cnnConnection.Execute strSql
                        Sleep 1000
                        'Added by Lydia 2021/09/27 代理人為Y2776600 MURATA，同時相關收文的”是否有檢索: Y”，請另外發信給承辦及其主管
                        If "" & .TextMatrix(iRow, colNA51) <> "" And "" & .TextMatrix(iRow, colC2CP148) = "Y" And Left("" & .TextMatrix(iRow, colPA75), 8) = "Y2776600" Then
                             strExc(0) = "": strExc(1) = ""
                             strExc(0) = "" & .TextMatrix(iRow, colNA51)  '承辦
                             strExc(1) = PUB_GetFCPProSup(strExc(0))
                             strExc(1) = strExc(1) & IIf(strExc(1) <> "", ";backup", "backup")  '+backup匯入卷宗區
                             strExc(2) = "【請收文告代】Our Ref: " & .TextMatrix(iRow, colCaseNo) & "[INCOM.1917]"
                             strExc(3) = "本案為murata案件，已報告核准，因核准函有檢索報告，請收文告代，由工程師另外發函美代，謝謝。"
                             strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                         " values( '" & strUserNum & "','" & strExc(0) & "',to_char(sysdate,'yyyymmdd')" & _
                                         ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "' ,'" & strExc(3) & "' ,'" & strExc(1) & "' )"
                             cnnConnection.Execute strSql
                             Sleep 1000
                        End If
                        'end 2021/09/27
                   End If
               End If
               'end 2021/08/30
               'Added by Lydia 2025/08/22  專利證書，非假發文
               If iKind = 2 And pType = 1 Then
                  '特定客戶優先二核期限控管; 日代<Y4520400> SOEI、<Y5518900>TOKOSHIE 各項指示、二核相關備註
                  If InStr("Y45204000,Y55189000", Mid(.TextMatrix(iRow, colPA75), 1, 8)) > 0 Then '若有異動請一併修改發文frm060104_e
                     strExc(0) = "select cp14,st04,cp09 from caseprogress,staff where cp01='" & strCase(1) & "' and cp02='" & strCase(2) & "' and cp03='" & strCase(3) & "' and cp04='" & strCase(4) & "' and cp158=0 and cp10='926' and cp14=st01(+) "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        '掛上【核對已准專利】期限：承辦期限=本所期限=指定日期之前'證書發文日起算14個日曆天(遇假日往前至前一工作天)
                        strExc(1) = PUB_GetWorkDay1(CompDate(2, 14, strSrvDate(1)), True)
                        strSql = "Update CaseProgress Set cp48=" & strExc(1) & ", cp06=" & strExc(1) & " where cp09='" & RsTemp.Fields("cp09") & "' "
                        cnnConnection.Execute strSql
                        '自動發一封Email給承辦工程師及程序人員(特殊設定)  , 內容如下
                        strExc(2) = "" & RsTemp.Fields("cp14")
                        If "" & RsTemp.Fields("st04") <> "1" Then
                           Call PUB_GetFCPCP14_F21(strCase, strExc(2))
                        End If
                        If strExc(2) <> "" Then
                           '收件人員: 承辦工程師、江如玉(固定核對公報程序人員)  副本:工程師主管
                           strExc(3) = PUB_GetFCPEngSup(strExc(2))
                           strExc(4) = Pub_GetSpecMan("外專程序-匯入公告本收件者")
                           strExc(5) = "【優先二次核對】" & strCase(1) & "-" & strCase(2) & IIf(strCase(3) & strCase(4) <> "000", "-" & strCase(3) & "-" & strCase(4), "") & "證書已寄出、請優先二次核對"
                           strExc(6) = "TO:" & GetStaffName(strExc(2), True) & vbCrLf & _
                                       String(4, " ") & strCase(1) & "-" & strCase(2) & IIf(strCase(3) & strCase(4) <> "000", "-" & strCase(3) & "-" & strCase(4), "") & "客戶要求於寄證書後2週內二核請款報告，期限: " & ChangeWStringToTDateString(strExc(1)) & "，請優先處理二次核對報告。"
                           strExc(6) = strExc(6) & vbCrLf & vbCrLf & "TO:" & PUB_ReadUserData(strExc(4)) & vbCrLf & _
                                        String(4, " ") & "請優先核對公報速退工程師進行二核。"
                           strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                       " values( '" & strUserNum & "','" & strExc(2) & IIf(strExc(4) <> "", ";" & strExc(4), "") & "',to_char(sysdate,'yyyymmdd')" & _
                                       ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strExc(5)) & "' ,'" & ChgSQL(strExc(6)) & "' ,'" & strExc(3) & "' )"
                           cnnConnection.Execute strSql
                        End If
                     End If
                  End If  'end --- '特定客戶優先二核期限控管; 日代<Y4520400> SOEI、<Y5518900>TOKOSHIE 各項指示、二核相關備註
               End If
               'end 2025/08/22
            End If
         End If
      Next
   End With
cnnConnection.CommitTrans 'Added by Lydia 2019/05/31

   Call doQuery
      
ErrHnd:
   'Added by Lydia 2019/05/31
   Exit Sub
   If bolUpdate = True Then cnnConnection.RollbackTrans
   'end 2019/05
   'Modified by Lydia 2019/11/26 Jessica反應有時候程式會出錯
   'If Err.Number <> 0 Then MsgBox Err.Description
   If Err.Number <> 0 Then MsgBox Err.Description & vbCrLf & "SQL語法：" & strSql
End Sub

Private Sub Form_Activate()
   bolBarShow = Forms(0).StatusBar1.Visible
   Forms(0).StatusBar1.Visible = True
End Sub

Private Sub Form_Deactivate()
   Forms(0).StatusBar1.Visible = bolBarShow
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

'語法內有用組合欄位為條件以控制使用特定index(避掉某些不適當的)
Private Sub doQuery(Optional bolShow As Boolean = True)
Dim strConSql As String
   
   'Added by Lydia 2019/05/31 工作項目
   '1.畫面上增加"工作項目"、"承辦人員"的下拉選項，可以切換個人負責的工作，切換項目會自動啟動查詢。
   '2.工作項目:下拉選單"1.告准函、2.專利證書、3.公開公報、4.專利權消滅、5.通知年費逾期"。
   iKind = Left(Combo2.Text, 1)

   strConSql = ""
   'Modified by Lydia 2019/05/31 告准函才有定稿日
   'If Val(Txt1(0)) > 0 Then
   If Val(txt1(0)) > 0 And iKind = "1" Then
      strConSql = " and C1.cp85=" & DBDATE(txt1(0))
   End If
   
   'Modified by Lydia 2019/05/31 避免查詢無資料,造成點選列判斷位置有誤
   'SetGrid
   Call SetGrid(True)
   
   'Added by Lydia 2019/05/31 指定承辦人
   If Combo1.Text <> "" Then
         strConSql = strConSql & " and c1.cp14='" & Trim(Left(Combo1, 6)) & "' "
   End If
   
    'Modify By Sindy 2018/7/19 + substr(sqldatet(.....),1,10)
    'ADODB.Recordset用Sort方法排序時發生-2147467259 無法在其定義長度是不明
    '或過長的資料行執行 Relate、Compute By、及 Sort 操作的錯誤
    '當在O12上select的欄位是DB的function回傳值時，Recordset看到的資料長度會是32767(超過Sort的限制)，
    '要加 substr 指定長度才不會發生錯誤。
    'Modified by Lydia 2019/05/31 區分工作項目+配合其他案件性質
    'strExc(0) = "Select '' V,substr(nvl(GETEMAILFLAG(c1.cp09),' '),1,1) E化情形,substr(sqldatet(c2.cp27),1,10) 核准函發文日,substr(sqldatet(c1.cp48),1,10) 承辦期限,substr(sqldatet(c1.cp85),1,10) 定稿日,S1.ST02 管制人" & _
                ",S2.ST02 承辦人,c1.CP01||'-'||c1.CP02||DECODE(c1.CP03||c1.CP04,'000','','-'||c1.CP03||'-'||c1.CP04) 本所案號" & _
                ",c1.cp64 備註,c1.cp09 告准總收文號" & _
                " from caseprogress c1,caseprogress c2,staff s1,staff s2,PATENT,FAGENT,Nation" & _
                " where c1.cp01='FCP' and c1.cp10='1917' and c1.cp27||c1.cp57 is null" & _
                " AND PA01(+)=c1.CP01 AND PA02(+)=c1.CP02 AND PA03(+)=c1.CP03 AND PA04(+)=c1.CP04" & _
                " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
                " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & _
                " and na16=s1.st01(+) and c1.cp14=s2.st01(+)" & _
                " and c1.cp43=c2.cp09(+) and C1.cp85 is not null" & strConSql & _
                " order by substr(nvl(GETEMAILFLAG(c1.cp09),' '),1,1),pa01,pa02,pa03,pa04"
   Select Case iKind
        Case "1" '1.告准函1917
          'Modified by Lydia 2019/06/17 本所案號前加註銷卷＊/閉卷●
            'Modified by Lydia 2021/08/30 +承辦工程師C2CP14
            'Modifie by Lydia 2021/09/27 +C2CP148, NA51(FCP承辦管制),pa75,pa26~pa30
            'Modified by Morgan 2024/3/5 +pa150
            strExc(0) = "Select '' V,substr(nvl(GETEMAILFLAG(c1.cp09),' '),1,1) E化情形,substr(sqldatet(c2.cp27),1,10) 核准函發文日,substr(sqldatet(c1.cp05),1,10) 來函收文日,substr(sqldatet(c1.cp48),1,10) 承辦期限,substr(sqldatet(c1.cp85),1,10) 定稿日,S1.ST02 管制人" & _
                        ",S2.ST02 承辦人,decode(pa57,null,'','＊')||decode(pa108,null,'','●')||c1.CP01||'-'||c1.CP02||DECODE(c1.CP03||c1.CP04,'000','','-'||c1.CP03||'-'||c1.CP04) 本所案號" & _
                        ",substr(c1.cp64,1,80) 備註,c1.cp09 C1總收文號,c2.cp14 as C2CP14,c2.cp148 as C2CP148, NA51 , PA75, PA26, PA27, PA28, PA29, PA30, PA150 " & _
                        " from caseprogress c1,caseprogress c2,staff s1,staff s2,PATENT,FAGENT,Nation" & _
                        " where c1.cp01='FCP' and c1.cp10='" & iPty(iKind) & "' and c1.cp158=0 and c1.cp159=0" & _
                        " AND PA01(+)=c1.CP01 AND PA02(+)=c1.CP02 AND PA03(+)=c1.CP03 AND PA04(+)=c1.CP04" & _
                        " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
                        " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & _
                        " and na16=s1.st01(+) and c1.cp14=s2.st01(+)" & _
                        " and c1.cp43=c2.cp09(+) and C1.cp85 is not null" & strConSql & _
                        " order by substr(nvl(GETEMAILFLAG(c1.cp09),' '),1,1),pa01,pa02,pa03,pa04"
        Case Else '其他:2.專利證書1603(有出定稿日才可以發文)、3.公開公報1229、4.專利權消滅1604、5.通知年費逾期1605
            'Added by Lydia 2019/08/16 區分通知期限
            'Modified by Lydia 2021/08/30 +承辦工程師C2CP14
            If InStr(iPty(iKind), "-") > 0 Then
                strExc(1) = Mid(iPty(iKind), 1, InStr(iPty(iKind), "-") - 1) '
                '相關收文號的案件性質
                strExc(2) = Mid(iPty(iKind), InStr(iPty(iKind), "-") + 1)
                strExc(2) = " and np07='" & strExc(2) & "' "
                'Modifie by Lydia 2021/09/27 +C2CP148, NA51(FCP承辦管制),pa75,pa26~pa30
                'Modified by Morgan 2024/3/5 +pa150
                strExc(0) = "Select '' V,' ' E化情形,' ' 核准函發文日,substr(sqldatet(c1.cp05),1,10) 來函收文日,substr(sqldatet(c1.cp48),1,10) 承辦期限,substr(sqldatet(c1.cp85),1,10) 定稿日,S1.ST02 管制人" & _
                            ",S2.ST02 承辦人,decode(pa57,null,'','＊')||decode(pa108,null,'','●')||c1.CP01||'-'||c1.CP02||DECODE(c1.CP03||c1.CP04,'000','','-'||c1.CP03||'-'||c1.CP04) 本所案號" & _
                            ",substr(c1.cp64,1,80) 備註,c1.cp09 C1總收文號,'' as C2CP14,'' AS C2CP148, NA51, PA75, PA26, PA27, PA28, PA29, PA30, PA150 " & _
                            " from caseprogress c1,staff s1,staff s2,PATENT,FAGENT,Nation,nextprogress" & _
                            " where c1.cp01='FCP' and c1.cp10='" & strExc(1) & "' and c1.cp158=0 and c1.cp159=0 and c1.cp43=np01(+) and c1.cp30=np22(+)" & strExc(2) & _
                            " AND PA01(+)=c1.CP01 AND PA02(+)=c1.CP02 AND PA03(+)=c1.CP03 AND PA04(+)=c1.CP04" & _
                            " AND PA75 IS NOT NULL" & _
                            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & _
                            " and na16=s1.st01(+) and c1.cp14=s2.st01(+)" & strConSql & IIf(iKind = "2", " and c1.cp85 is not null ", "") & _
                            " order by c1.cp48,pa01,pa02,pa03,pa04"
            Else
            'end 2019/08/16
                'Modifeid by Lydia 2019/06/17 拿掉PA57||PA108 IS NULL ; 若是已上閉卷的案件，各項大批進度檔發文日請先上"111111"(目前年費逾繳已會先上111111)'
                                                                                                        '，若有特殊案件已閉卷尚需報告，則程序會個案下定稿，到時自動將發文日"111111"拿掉，如此案件又可自大批上發文。
                'Modified by Lydia 2019/06/17 本所案號前加註銷卷＊/閉卷●
                'Modified by Lydia 2021/08/30 +承辦工程師C2CP14
                'Modifie by Lydia 2021/09/27 +C2CP148,NA51(FCP承辦管制),pa75,pa26~pa30
                'Modified by Morgan 2024/3/5 +pa150
                strExc(0) = "Select '' V,' ' E化情形,' ' 核准函發文日,substr(sqldatet(c1.cp05),1,10) 來函收文日,substr(sqldatet(c1.cp48),1,10) 承辦期限,substr(sqldatet(c1.cp85),1,10) 定稿日,S1.ST02 管制人" & _
                            ",S2.ST02 承辦人,decode(pa57,null,'','＊')||decode(pa108,null,'','●')||c1.CP01||'-'||c1.CP02||DECODE(c1.CP03||c1.CP04,'000','','-'||c1.CP03||'-'||c1.CP04) 本所案號" & _
                            ",substr(c1.cp64,1,80) 備註,c1.cp09 C1總收文號,'' as C2CP14, '' AS C2CP148, NA51, PA75, PA26, PA27, PA28, PA29, PA30, PA150 " & _
                            " from caseprogress c1,staff s1,staff s2,PATENT,FAGENT,Nation" & _
                            " where c1.cp01='FCP' and c1.cp10='" & iPty(iKind) & "' and c1.cp158=0 and c1.cp159=0" & _
                            " AND PA01(+)=c1.CP01 AND PA02(+)=c1.CP02 AND PA03(+)=c1.CP03 AND PA04(+)=c1.CP04" & _
                            " AND PA75 IS NOT NULL" & _
                            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & _
                            " and na16=s1.st01(+) and c1.cp14=s2.st01(+)" & strConSql & IIf(iKind = "2", " and c1.cp85 is not null ", "") & _
                            " order by c1.cp48,pa01,pa02,pa03,pa04"
            End If 'end 2019/08/16
   End Select
   'end 2019/05/31
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If RsTemp Is Nothing Then Exit Sub
   'Remove by Lydia 2019/05/31 後面有設來源
   'Set GrdDataList.Recordset = RsTemp
   'RecordShow
   'end 2019/05/31
   lblCnt2.Caption = ""
   If RsTemp.RecordCount = 0 Then
      'Remove by Lydia 2019/05/31 若無資料再設到Grid中,會造成點選位置計算錯誤( Morgan : 最近發現的問題)
      'Set m_adoRst = RsTemp.Clone
      'SetRst2Grid
      'end 2019/05/31
      If bolShow = True Then
         MsgBox "查無資料！", vbInformation
      End If
      LblCnt.Caption = "共 0 筆"
   Else
      Set m_adoRst = RsTemp.Clone
      'Mark by Lydia 2019/05/31
      'm_stSort = "E化情形 asc,本所案號 asc"
      'm_adoRst.Sort = m_stSort
      'end 2019/05/31
      SetRst2Grid
      Call SetGrid 'Added by Lydia 2019/05/31
      LblCnt.Caption = "共 " & RsTemp.RecordCount & " 筆"
      m_blnColOrderAsc = True
   End If
End Sub

'Modified by Lydia 2019/05/31
'Private Sub SetGrid()
Private Sub SetGrid(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modified by Lydia 2019/05/31 區分工作項目
   'arrGridHeadText = Array("V", "E化情形", "核准函發文日", "承辦期限", "定稿日", "管制人", "承辦人", "本所案號", "備註", "告准總收文號")
   'arrGridHeadWidth = Array(200, 750, 1250, 900, 900, 900, 900, 1200, 2000, 0)
   'Modified by Lydia 2021/08/30 +承辦工程師C2CP14
   'Modifie by Lydia 2021/09/27 +C2CP148,NA51(FCP承辦管制),pa75,pa26~pa30
   'Modified by Morgan 2024/3/5 +pa150
   arrGridHeadText = Array("V", "E化情形", "核准函發文日", "來函收文日", "承辦期限", "定 稿 日", "管制人", "承辦人", "本  所  案  號", "備　　註", "C1總收文號", "C2CP14", _
                               "C2CP148", "NA51", "PA75", "PA26", "PA27", "PA28", "PA29", "PA30", "PA150")
   If iKind = "1" Then '告准函
       'Modified by Lydia 2021/08/30 +承辦工程師C2CP14 => +0
       'Modifie by Lydia 2021/09/27 +C2CP148,NA51(FCP承辦管制),pa75,pa26~pa30 => +0
       'Modified by Morgan 2024/3/5 +pa150
       arrGridHeadWidth = Array(240, 750, 1300, 0, 1000, 1000, 1000, 1000, 1500, 2000, 0, 0, _
                                0, 0, 0, 0, 0, 0, 0, 0, 0)
   Else
       arrGridHeadWidth = Array(240, 0, 0, 1100, 1000, 0, 0, 1000, 1500, 2000, 0, 0, _
                                0, 0, 0, 0, 0, 0, 0, 0, 0)
   End If
   'end 2019/05/31
   grdDataList.Visible = False
   grdDataList.Cols = UBound(arrGridHeadText) + 1
   'Modified by Lydia 2019/05/31 避免查詢無資料,造成點選列判斷位置有誤
   'GrdDataList.Rows = 2
   If pReset = True Then
         grdDataList.Clear
         grdDataList.Rows = 2
   End If
   'end 2019/05/31
   
   For iRow = 0 To grdDataList.Cols - 1
      grdDataList.row = 0
      grdDataList.col = iRow
      grdDataList.Text = arrGridHeadText(iRow)
      grdDataList.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grdDataList.CellAlignment = flexAlignCenterCenter
   Next
   
   'Added by Lydia 2019/05/31 預設欄位置
   If m_iCols = 0 Then
      colCp09 = PUB_MGridGetId("C1總收文號", grdDataList)
      m_iCols = UBound(arrGridHeadText)
      'Added by Lydia 2021/08/30
      colCaseNo = PUB_MGridGetId("本  所  案  號", grdDataList)
      colC2CP14 = PUB_MGridGetId("C2CP14", grdDataList)  '承辦工程師
      'end 2021/08/30
      'Added by Lydia 2021/09/27
      colC2CP148 = PUB_MGridGetId("C2CP148", grdDataList) '相關收文號的”是否有檢索”
      colNA51 = PUB_MGridGetId("NA51", grdDataList)  'FCP承辦管制
      colPA75 = PUB_MGridGetId("PA75", grdDataList)
      colPA26 = PUB_MGridGetId("PA26", grdDataList)  'PA27~PA30依順序在PA26的後面
      'end 2021/09/27
   End If
   
   grdDataList.Visible = True
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   PUB_SetPrinter Me.Name, cboPrinter, strPrinter
   
   'Added by Lydia 2019/05/31
   Combo2.Clear
   Combo2.AddItem "1. 告准函"
   iPty(1) = "1917"
   Combo2.AddItem "2. 專利證書"
   iPty(2) = "1603"
   Combo2.AddItem "3. 公開公報"
   iPty(3) = "1229"
   Combo2.AddItem "4. 專利權消滅"
   iPty(4) = "1604"
   Combo2.AddItem "5. 通知年費逾期"
   iPty(5) = "1605"
   'Added by Lydia 2019/08/16
   Combo2.AddItem "6. 期限通知-年費"
   iPty(6) = "1913-605"
   'Added by Lydia 2025/06/05
   Combo2.AddItem "7. 期限通知-實體審查"
   iPty(7) = "1913-416"
   'Memo by Lydia 2025/06/18 增加大項，必須一併修改期限查詢frm060210
   
   Call SetCombo1
   'end 2019/05/31
   
   Call doQuery(False)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Lydia 2021/08/30
   
   Set frm060118 = Nothing
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = grdDataList.Recordset.RecordCount
End Sub

Private Sub grdDataList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim iCol As Integer

   If m_bolSelAll = True Then Exit Sub '為全選或全取消,不須執行下列程式段

   iCol = grdDataList.MouseCol
   'Modify By Sindy 2018/7/19 + And m_adoRst.Fields(iCol).Name <> "備註"
   'ADODB.Recordset用Sort方法排序時發生-2147467259 無法在其定義長度是不明
   '或過長的資料行執行 Relate、Compute By、及 Sort 操作的錯誤
   '當在O12上select的欄位是DB的function回傳值時，Recordset看到的資料長度會是32767(超過Sort的限制)，
   '要加 substr 指定長度才不會發生錯誤。
   If iCol > 0 And grdDataList.MouseRow < 1 And m_adoRst.Fields(iCol).Name <> "備註" Then
      grdDataList.Visible = False
      Set grdDataList.Recordset = Nothing
      If m_blnColOrderAsc = True Then
         m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " desc"
         m_blnColOrderAsc = False
      Else
         m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " asc"
         m_blnColOrderAsc = True
      End If
      SetRst2Grid
      LblCnt.Caption = "共 " & RsTemp.RecordCount & " 筆"
      lblCnt2.Caption = ""
      grdDataList.Visible = True
   End If
End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow grdDataList, x, y, nCol, nRow
   grdDataList.col = nCol
   grdDataList.row = nRow
   
End Sub

Private Sub grdDataList_SelChange()
Dim ii As Integer, jj As Integer
Dim strRows As Integer

   m_bolSelAll = False 'Add By Sindy 2017/3/1 是否全選或全取消
   With grdDataList
      'Add By Sindy 2017/3/1 全選或全取消
      If grdDataList.col = 0 And grdDataList.row = 0 Then
         If .Rows > 1 Then
            m_bolSelAll = True '為全選或全取消
            .Visible = False
            .row = 1
            .col = 0
            '全取消
            If .Text = "V" Then
               For jj = 1 To .Rows - 1
                  .row = jj
                  .col = 0
                  If .Text = "V" Then
                     .Text = ""
                     strRows = Val(Replace(Replace(lblCnt2.Caption, "選取", ""), "筆", ""))
                     If Val(strRows) > 0 Then
                        strRows = strRows - 1
                     Else
                        strRows = 0
                     End If
                     lblCnt2.Caption = "選取 " & strRows & " 筆"
                     For ii = 0 To .Cols - 1
                        .col = ii
                        .CellBackColor = .BackColor
                     Next ii
                  End If
               Next jj
            '全選
            Else
               For jj = 1 To .Rows - 1
                  .row = jj
                  .col = 0
                  If .Text = "" Then
                     .Text = "V"
                     strRows = Val(Replace(Replace(lblCnt2.Caption, "選取", ""), "筆", "")) + 1
                     lblCnt2.Caption = "選取 " & strRows & " 筆"
                     For ii = 0 To .Cols - 1
                        .col = ii
                        .CellBackColor = &HFFC0C0
                     Next ii
                  End If
               Next jj
            End If
            .Visible = True
         End If
      '2017/3/1 END
      ElseIf .MouseRow > 0 Then
         .Visible = False
         .row = .MouseRow
         .col = 0
         If .Text = "V" Then
            .Text = ""
            strRows = Val(Replace(Replace(lblCnt2.Caption, "選取", ""), "筆", ""))
            If Val(strRows) > 0 Then
               strRows = strRows - 1
            Else
               strRows = 0
            End If
            lblCnt2.Caption = "選取 " & strRows & " 筆"
            For ii = 0 To .Cols - 1
               .col = ii
               .CellBackColor = .BackColor
            Next ii
         Else
            .Text = "V"
            strRows = Val(Replace(Replace(lblCnt2.Caption, "選取", ""), "筆", "")) + 1
            lblCnt2.Caption = "選取 " & strRows & " 筆"
            For ii = 0 To .Cols - 1
               .col = ii
               .CellBackColor = &HFFC0C0
            Next ii
         End If
         .Visible = True
      End If
   End With
End Sub

Private Sub DoPrint()
Dim iOrientation As Integer, iRow As Integer, iCol As Integer
Dim strTemp() As String
   
   iOrientation = Printer.Orientation
   If iKind = "1" Or iKind = "2" Then 'Added by Lydia 2019/05/31 區分工作項目
       Printer.Orientation = 2 '橫印: 告准函,專利證書
   'Added by Lydia 2019/05/31
   Else
       Printer.Orientation = 1 '直印
   End If
   'end 2019/05/31
   
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   With grdDataList
      GetPleft
      m_iPrtCols = m_iCols
      ReDim strTemp(1 To m_iPrtCols)
      
      iPage = 1
      PrintPageHeader
      PrintPageHeader1
      For iRow = 1 To .Rows - 1
         For iCol = LBound(strTemp) To UBound(strTemp)
            strTemp(iCol) = .TextMatrix(iRow, iCol)
         Next
         PrintDetail strTemp
      Next
      Call PrintReportFooter(.Rows - 1)
      Printer.EndDoc
      MsgBox "列印完成！"
   End With
   Printer.Orientation = iOrientation
End Sub

Sub GetPleft()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   'm_iCols = 8 'Mark by Lydia 2019/05/31 改在SetGrid設定
   ReDim PLeft(1 To m_iCols)
   PLeft(1) = ciStartX
   For intI = 2 To m_iCols
      If grdDataList.ColWidth(intI - 1) > 0 Then
         PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(grdDataList.TextMatrix(0, intI - 1)) + ciColGap
      Else
         PLeft(intI) = PLeft(intI - 1)
      End If
   Next
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)
   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      'Modified by Lydia 2019/05/31
      'Printer.Print String(130, "-")
      Call PrintLine
      
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If bolSubtotal Then
         PrintPageHeader1
         iPrint = iPrint + lngLineHeight
      End If
   End If
End Sub

Sub PrintDetail(strData() As String)
   Dim iCol As Integer
   Dim strTemp As String 'Added by Lydia 2019/06/17
   
   PrintNewLine
   For iCol = LBound(strData) To UBound(strData)
      If Me.grdDataList.ColWidth(iCol) > 0 Then
         Printer.CurrentX = PLeft(iCol)
         Printer.CurrentY = iPrint
         strTemp = PUB_StringFilter(strData(iCol)) 'Added by Lydia 2019/06/17 去除跳行符號
         'Added by Lydia 2019/05/31 備註限長度
         If iCol = 9 Then
             If iKind = "2" Then
                 'Modified by Lydia 2019/06/17 strData(iCol)=> strtemp
                 Printer.Print convForm(strTemp, 80)
             Else
                 Printer.Print convForm(strTemp, 46)
             End If
         Else
         'end 2019/05/31
              Printer.Print strTemp
         End If 'end 2019/05/31
      End If
   Next
End Sub

Sub PrintPageHeader()
Dim strPTmp As String
   
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = Me.Caption & "清單"
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   PrintNewLine
   'Added by Lydia 2019/05/31
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "工作項目：" & Trim(Mid(Combo2.Text, 3))
   'end 2019/05/31
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   'Modified by Lydia 2019/05/31
   'Printer.Print String(130, "-")
   Call PrintLine
End Sub

Sub PrintPageHeader1()
   Call PrintNewLine(False, 1)
   For intI = 1 To m_iPrtCols
     If Me.grdDataList.ColWidth(intI) > 0 Then
        Printer.CurrentX = PLeft(intI)
        Printer.CurrentY = iPrint
        Printer.Print grdDataList.TextMatrix(0, intI)
     End If
   Next
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   'Modified by Lydia 2019/05/31
   'Printer.Print String(130, "-")
   Call PrintLine
End Sub

'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)
   Call PrintNewLine(True, 1)
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   'Modified by Lydia 2019/05/31
   'Printer.Print String(130, "-")
   Call PrintLine
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "合計： " & iRecCount & " 筆"
   Printer.EndDoc
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Dim iLen As Integer
   
   Select Case Index
      Case 0
         If txt1(Index) <> "" Then
            If CheckIsTaiwanDate(txt1(Index)) = False Then
               GoTo JumpCancel
            End If
         End If
   End Select
   
   If Cancel = False Then
      If txt1(Index).MaxLength > 0 Then
         If Not CheckLengthIsOK(txt1(Index), iLen) Then
            GoTo JumpCancel
         End If
      End If
   End If
   Exit Sub
   
JumpCancel:
   txt1_GotFocus Index
   Cancel = True
End Sub

'Added by Lydia 2019/05/31
Private Sub SetCombo1()
Dim strTmp As String

   Combo1.Clear
   strExc(0) = "select st01,st02 from staff a where st03='F22' and st04='1' order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not RsTemp.EOF
         If .Fields("st01") = strUserNum Then
            Combo1.AddItem .Fields("st01") & " " & .Fields("st02"), 0
            strTmp = .Fields("st01") & " " & .Fields("st02")
         Else
            Combo1.AddItem .Fields("st01") & " " & .Fields("st02")
         End If
      .MoveNext
      Loop
      End With
   End If
   
   If strTmp <> "" Then
      Combo1.ListIndex = 0
   Else
      Combo1.ListIndex = Combo1.ListCount - 1
   End If
   
   '抓操作者有整批未發文的案件性質,預設工作項目種類
   strExc(1) = ""
   strExc(2) = ""
   For intI = 1 To UBound(iPty)
       'Modified by Lydia 2019/08/16 改成組合語法
       'strExc(1) = strExc(1) & ", '" & iPty(intI) & "' , '" & intI & "' "
       'strExc(2) = strExc(2) & "," & iPty(intI)
       If iPty(intI) <> "" Then
           If InStr(iPty(intI), "-") > 0 Then '通知期限
               strExc(1) = strExc(1) & "Union select '" & intI & "' as ord1, cp10,count(*) cnt " & _
                               "from caseprogress,nextprogress where cp01='FCP'  and cp05>=20190501 and cp158=0 and cp159=0 and cp14='" & Trim(Left(Combo1.Text, 6)) & "' and cp10='" & Mid(iPty(intI), 1, InStr(iPty(intI), "-") - 1) & "' " & _
                               "and cp43=np01(+) and cp30=np22(+) and np07='" & Mid(iPty(intI), InStr(iPty(intI), "-") + 1) & "' group by cp10 "
           Else
               strExc(1) = strExc(1) & "Union select '" & intI & "' as ord1, cp10,count(*) cnt " & _
                               "from caseprogress where cp01='FCP'  and cp05>=20190501 and cp158=0 and cp159=0 and cp14='" & Trim(Left(Combo1.Text, 6)) & "' and cp10='" & iPty(intI) & "' " & _
                               "group by cp10 "
           End If
       End If
       'end 2019/08/16
   Next intI
   'Modified by Lydia 2019/08/16
'   strExc(2) = GetAddStr(Mid(strExc(2), 2))
'   strExc(0) = "select decode(cp10 " & strExc(1) & ",'9' ) ord1,cp10 ,count(*) cnt " & _
'                    "from caseprogress where cp01='FCP'  and cp05>=20190501 and cp158=0 and cp159=0 and cp14='" & Trim(Left(Combo1.Text, 6)) & "' and cp10 in (" & strExc(2) & ") " & _
'                    "group by decode(cp10 " & strExc(1) & ",'9' ),cp10 "
   'strExc(0) = strExc(0) & " order by ord1 "
   strExc(0) = "select * from (" & Mid(strExc(1), 6) & ") where cnt > 0 order by ord1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       RsTemp.MoveFirst
       Combo2.ListIndex = Val(RsTemp.Fields("ord1")) - 1
   Else  '無=>預設告准函
       Combo2.ListIndex = 0
   End If
   
End Sub

Private Sub Combo2_Click()
     '切換人員先不啟動查詢
    If (Combo1.Tag <> "" And Combo1.Tag <> Combo1.Text) Or (Combo2.Tag <> "" And Combo2.Tag <> Combo2.Text) Then
        Call doQuery(True)
        txt1(0).Text = ""
    End If
    Combo1.Tag = Combo1.Text
    Combo2.Tag = Combo2.Text

    If Left(Combo2.Text, 1) = "1" Then '告准函才有定稿日期
        Label2(1).Visible = True
        txt1(0).Visible = True
    Else
        Label2(1).Visible = False
        txt1(0).Visible = False
    End If
End Sub

'列印分隔線
Private Sub PrintLine()
   If iKind = "1" Or iKind = "2" Then '橫印: 告准函,專利證書
       Printer.Print String(130, "-")
   Else
       Printer.Print String(92, "-")
   End If
End Sub

'Added by Lydia 2019/06/17 假發文
Private Sub CmdPass_Click()
   If MsgBox("是否確定假發文？", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
       Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Call doBatch("2")
   Screen.MousePointer = vbDefault
End Sub
