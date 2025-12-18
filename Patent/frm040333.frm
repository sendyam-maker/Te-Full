VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm040333 
   BorderStyle     =   1  '單線固定
   Caption         =   "FC案件不請款確認維護"
   ClientHeight    =   5750
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8950
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   2
      Left            =   7110
      TabIndex        =   7
      Top             =   90
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "存檔(&S)"
      Enabled         =   0   'False
      Height          =   345
      Index           =   1
      Left            =   6210
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   90
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   5310
      TabIndex        =   5
      Top             =   90
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2970
      MaxLength       =   2
      TabIndex        =   3
      Top             =   405
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   2
      Top             =   405
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1740
      MaxLength       =   6
      TabIndex        =   1
      Top             =   405
      Width           =   825
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   0
      Top             =   405
      Width           =   525
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4785
      Left            =   60
      TabIndex        =   4
      Top             =   870
      Width           =   8835
      _ExtentX        =   15593
      _ExtentY        =   8431
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   1
      FixedCols       =   0
      AllowUserResizing=   3
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
      _Band(0).Cols   =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "沒有費用支出者, 確認後不再出現！"
      Height          =   180
      Index           =   1
      Left            =   5256
      TabIndex        =   9
      Top             =   576
      Width           =   2796
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1470
      X2              =   3030
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   8
      Top             =   450
      Width           =   900
   End
End
Attribute VB_Name = "frm040333"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/01 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo By Sindy 2010/11/17 日期欄已修改
Option Explicit

Dim m_row As Integer, i As Integer

Private Sub cmdOK_Click(Index As Integer)
Dim intRow As Integer, bolRun As Boolean
   
   Select Case Index
   Case 0 '查詢
            If Trim(txt1(0)) = "" Or Trim(txt1(1)) = "" Then
               MsgBox "本所案號不可以空白！", vbCritical, "操作錯誤！"
               Exit Sub
            End If
            'modify by sonia 2016/11/24 加入P,PS,CFP,CPS (P-109934之補正AA5014329有銷帳單)
            'modify by sonia 2023/5/5 因FCL-010985/ 其他 / 收文日112.3.31 故不限制系統類別，僅LA不可
            'If txt1(0) <> "FCT" And txt1(0) <> "FCP" And txt1(0) <> "FG" And txt1(0) <> "S" And txt1(0) <> "P" And txt1(0) <> "PS" And txt1(0) <> "CFP" And txt1(0) <> "CPS" Then
            '   MsgBox "請輸入系統別為FCT,FCP,FG或S！", vbCritical, "操作錯誤！"
            If txt1(0) = "LA" Then
               MsgBox "系統類別不可為LA！", vbCritical, "操作錯誤！"
            'end 2023/5/5
               Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            GRD1.MousePointer = flexHourglass
            doQuery
            GRD1.MousePointer = flexDefault
            Screen.MousePointer = vbDefault
   Case 1 '存檔
            cnnConnection.BeginTrans
            bolRun = False
            For intRow = 1 To GRD1.Rows - 1
               If GRD1.TextMatrix(intRow, 0) = "V" Then
                  'modify by sonia 2016/12/6 加入判斷 P-109934補正
                  'strSql = "update caseprogress" & _
                                 " set cp16=null,cp17=null,cp18=null,cp20='N',cp64='" & ChangeTStringToTDateString(strSrvDate(2)) & "取消費用'||cp16||'規費'||cp17||'點數'||cp18||'及不請款;'||cp64" & _
                                 " where CP09='" & grd1.TextMatrix(intRow, 8) & "'"
                  '下列cp64的更新,'取消費用','及不請款'的字樣不可改變,否則會造成frm040332錯誤
                  '有收費
                  If Val(GRD1.TextMatrix(intRow, 4)) <> 0 Then
                     strSql = "update caseprogress" & _
                                    " set cp16=null,cp17=null,cp18=null,cp20='N',cp64='" & ChangeTStringToTDateString(strSrvDate(2)) & "取消費用'||cp16||'規費'||cp17||'點數'||cp18||'及不請款;'||cp64" & _
                                    " where CP09='" & GRD1.TextMatrix(intRow, 8) & "'"
                  '未收費
                  'add by sonia 2016/12/9 負點數(作業失誤) FCP-044942
                  ElseIf Val(GRD1.TextMatrix(intRow, 6)) < 0 Then
                     strSql = "update caseprogress" & _
                                    " set cp20='N',cp64='" & ChangeTStringToTDateString(strSrvDate(2)) & "取消費用(負點數但不能請款)及不請款;'||cp64" & _
                                    " where CP09='" & GRD1.TextMatrix(intRow, 8) & "'"
                  'end 2016/12/9
                  Else
                     strSql = "update caseprogress" & _
                                    " set cp20='N',cp64='" & ChangeTStringToTDateString(strSrvDate(2)) & "取消費用(有支出但不能請款)及不請款;'||cp64" & _
                                    " where CP09='" & GRD1.TextMatrix(intRow, 8) & "'"
                  End If
                  'end 2016/12/6
                  Pub_SeekTbLog strSql
                  cnnConnection.Execute strSql
                  bolRun = True
               End If
            Next intRow
            If bolRun = True Then
               MsgBox "不請款確認已完成！", vbExclamation
            End If
            cnnConnection.CommitTrans
            Call doQuery
   Case 2 '結束
            Unload Me
   Case Else
   End Select
End Sub

Sub doQuery()
Dim strCP01 As String, strCP0110 As String
On Error GoTo ErrHnd
   
   cmdOK(1).Enabled = False
   If txt1(2) = "" Then txt1(2) = "0"
   If txt1(3) = "" Then txt1(3) = "00"
   'modify by sonia 2016/11/24 加入P,PS,CFP,CPS (P-109934之補正AA5014329有銷帳單)
   'modify by sonia 2023/5/5 因FCL-010985/ 其他 / 收文日112.3.31 故不限制系統類別
   'If txt1(0) = "FCT" Or txt1(0) = "FCP" Or txt1(0) = "FG" Or txt1(0) = "P" Or txt1(0) = "PS" Or txt1(0) = "CFP" Or txt1(0) = "CPS" Then
   If txt1(0) <> "S" Then
      'cancel by sonia 2023/5/5 因FCL-010985/ 其他 / 收文日112.3.31 故不限制系統類別
      'strCP01 = "'FCT','FCP','FG','P','PS','CFP','CPS'" '系統別
     
      'modify by sonia 2019/8/6 +T303,T705,P404,P911,PS404,PS911,CFP404,CFP911
      'strCP0110 = "'FCT303','FCT705','FCP404','FCP911','FG911'" '系統別+案件性質(延期+補收款)
      strCP0110 = "'FCT303','FCT705','FCP404','FCP911','FG911','T303','T705','P404','P911','PS404','PS911','CFP404','CFP911'" '系統別+案件性質(延期+補收款)
      '指該收文號不請款但有發文規費...cp20||''<==加||''是為了影響index讓查詢速度變快
      'modify by sonia 2016/11/24 欄位+CP64,cp84改為decode(cp84,null,cp61,'0',cp61,cp84),條件cp84>0改為(cp84>0 or cp61 is not null) (P-109934之補正AA5014329有銷帳單)
      'modify by sonia 2016/2/6 加入P,PS,CFP,CPS故案件性質要判斷申請國家
     ' strSql = " select ' ',cp01||'-'||cp02||'-'||cp03||'-'||cp04 CaseNO,(CP05-19110000),SUBSTR(CPM03,1,6),cp16,cp17,cp18,decode(cp84,null,cp61,'0',cp61,cp84) cp84,CP09,(CP27-19110000),CP64" & _
                     " From caseprogress a, casepropertymap" & _
                     " where cp01 in(" & strCP01 & ")" & _
                     " and cp01||cp10 not in(" & strCP0110 & ")" & _
                     " and SUBSTR(cp12,1,1)='F'" & _
                     " and cp20||''='N'" & _
                     " and (cp84>0 or cp61 is not null)" & _
                     " and CP60 IS NULL" & _
                     " and (cp18>=0 or cp18 is null)" & _
                     " and CP01=CPM01(+) AND CP10=CPM02(+)" & _
                     " and CP01='" & txt1(0) & "' and CP02='" & txt1(1) & "' and CP03='" & txt1(2) & "' and CP04='" & txt1(3) & "'"
      'modify by sonia 2016/12/9 取消and (cp18>=0 or cp18 is null)控制 FCP-044942
      'modify by sonia 2019/8/6 配合frm040332, cp84>0改為nvl(cp16,0)+nvl(cp84,0)>0
      'modify by sonia 2023/5/5 因FCL-010985/ 其他 / 收文日112.3.31 故不限制系統類別取消cp01 in(" & strCP01 & ")並加入lawcase,FCL案件LXX部門收文故SUBSTR(cp12,1,1)='F'再加SUBSTR(cp01,1,1)='F'
      'strSql = " select ' ',cp01||'-'||cp02||'-'||cp03||'-'||cp04 CaseNO,(CP05-19110000),SUBSTR(decode(nvl(nvl(pa09,tm10),sp09),'000',CPM03,CPM04),1,6),cp16,cp17,cp18,decode(cp84,null,cp61,'0',cp61,cp84) cp84,CP09,(CP27-19110000),CP64" & _
                     " From caseprogress a, casepropertymap,patent,trademark,servicepractice" & _
                     " where cp01 in(" & strCP01 & ")" & _
                     " and cp01||cp10 not in(" & strCP0110 & ")" & _
                     " and SUBSTR(cp12,1,1)='F'" & _
                     " and cp20||''='N'" & _
                     " and (nvl(cp16,0)+nvl(cp84,0)>0 or cp61 is not null)" & _
                     " and CP60 IS NULL" & _
                     " and CP01=CPM01(+) AND CP10=CPM02(+)" & _
                     " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
                     " and CP01='" & txt1(0) & "' and CP02='" & txt1(1) & "' and CP03='" & txt1(2) & "' and CP04='" & txt1(3) & "'"
      strSql = " select ' ',cp01||'-'||cp02||'-'||cp03||'-'||cp04 CaseNO,(CP05-19110000),SUBSTR(decode(nvl(nvl(pa09,tm10),sp09),'000',CPM03,CPM04),1,6),cp16,cp17,cp18,decode(cp84,null,cp61,'0',cp61,cp84) cp84,CP09,(CP27-19110000),CP64" & _
                     " From caseprogress a, casepropertymap,patent,trademark,servicepractice,lawcase" & _
                     " where cp01||cp10 not in(" & strCP0110 & ")" & _
                     " and (SUBSTR(cp12,1,1)='F' or SUBSTR(cp01,1,1)='F') " & _
                     " and cp20||''='N'" & _
                     " and (nvl(cp16,0)+nvl(cp84,0)>0 or cp61 is not null)" & _
                     " and CP60 IS NULL" & _
                     " and CP01=CPM01(+) AND CP10=CPM02(+)" & _
                     " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)" & _
                     " and CP01='" & txt1(0) & "' and CP02='" & txt1(1) & "' and CP03='" & txt1(2) & "' and CP04='" & txt1(3) & "'"
      
      'Add by Morgan 2010/6/3 排除FCP申請案有收文工程師提申
      strSql = strSql & " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp10='940' and b.cp57 is null and a.cp01='FCP' and a.cp10 in ('101','102','103','105'))"
      
      '指延期案件性質有發文規費但其相關總收文號不請款
      'modify by sonia 2016/11/24 cp84改為to_char(C1.cp84) cp84
      'modify by sonia 2019/8/6 配合frm040332, C1.cp84>0改為nvl(C1.cp16,0)+nvl(C1.cp84,0)>0, 並取消 and C1.cp20||''='N'條件P-116460延期AA8028334
      strSql = strSql & " Union All" & _
                     " select ' ',C1.cp01||'-'||C1.cp02||'-'||C1.cp03||'-'||C1.cp04 CaseNO,(C1.CP05-19110000),SUBSTR(CPM03,1,6),C1.cp16,C1.cp17,C1.cp18,to_char(C1.cp84) cp84,C1.CP09,(C1.CP27-19110000),C1.CP64" & _
                     " from caseprogress C1,caseprogress C2,casepropertymap" & _
                     " where C1.cp01||C1.cp10 in(" & strCP0110 & ")" & _
                     " and SUBSTR(C1.cp12,1,1)='F'" & _
                     " and (nvl(C1.cp16,0)+nvl(C1.cp84,0)>0 OR C1.CP61 IS NOT NULL)" & _
                     " and C1.CP60 IS NULL" & _
                     " and C1.CP43=C2.CP09(+)" & _
                     " and C2.CP60 IS NULL" & _
                     " and C2.cp20='N'" & _
                     " and (C2.cp18>=0 or C2.cp18 is null)" & _
                     " and C1.CP01=CPM01(+) AND C1.CP10=CPM02(+)" & _
                     " and C1.CP01='" & txt1(0) & "' and C1.CP02='" & txt1(1) & "' and C1.CP03='" & txt1(2) & "' and C1.CP04='" & txt1(3) & "'"
   'S(台灣案)
   ElseIf txt1(0) = "S" Then
      '指該收文號不請款但有發文規費...cp20||''<==加||''是為了影響index讓查詢速度變快
      'modify by sonia 2016/12/9 取消and (cp18>=0 or cp18 is null)控制 FCP-044942
      strSql = strSql & _
                     " select ' ',cp01||'-'||cp02||'-'||cp03||'-'||cp04 CaseNO,(CP05-19110000),SUBSTR(CPM03,1,6),cp16,cp17,cp18,cp84,CP09,(CP27-19110000),CP64" & _
                     " From caseprogress, servicepractice, casepropertymap" & _
                     " where cp01='S'" & _
                     " and cp01||cp10 not in('S705')" & _
                     " and SUBSTR(cp12,1,1)='F'" & _
                     " and cp20||''='N'" & _
                     " and cp84>0" & _
                     " and CP60 IS NULL" & _
                     " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)" & _
                     " and sp09='000'" & _
                     " and CP01=CPM01(+) AND CP10=CPM02(+)" & _
                     " and CP01='" & txt1(0) & "' and CP02='" & txt1(1) & "' and CP03='" & txt1(2) & "' and CP04='" & txt1(3) & "'"
      '指延期案件性質有發文規費但其相關總收文號不請款
      strSql = strSql & " Union All" & _
                     " select ' ',C1.cp01||'-'||C1.cp02||'-'||C1.cp03||'-'||C1.cp04 CaseNO,(C1.CP05-19110000),SUBSTR(CPM03,1,6),C1.cp16,C1.cp17,C1.cp18,C1.cp84,C1.CP09,(C1.CP27-19110000),C1.CP64" & _
                     " from caseprogress C1,caseprogress C2, servicepractice, casepropertymap" & _
                     " where C1.cp01||C1.cp10 in('S705')" & _
                     " and SUBSTR(C1.cp12,1,1)='F'" & _
                     " and C1.cp84>0" & _
                     " and C1.CP60 IS NULL and C1.cp20||''='N'" & _
                     " and C1.CP43=C2.CP09(+)" & _
                     " and C2.CP60 IS NULL" & _
                     " and C2.cp20='N'" & _
                     " and (C2.cp18>=0 or C2.cp18 is null)" & _
                     " and C1.cp01=sp01(+) and C1.cp02=sp02(+) and C1.cp03=sp03(+) and C1.cp04=sp04(+)" & _
                     " and sp09='000'" & _
                     " and C1.CP01=CPM01(+) AND C1.CP10=CPM02(+)" & _
                     " and C1.CP01='" & txt1(0) & "' and C1.CP02='" & txt1(1) & "' and C1.CP03='" & txt1(2) & "' and C1.CP04='" & txt1(3) & "'"
   End If
   strSql = strSql & " order by CaseNO"
   CheckOC3
   GRD1.Rows = 2
   GRD1.Clear
   SetDataListWidth
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Set GRD1.Recordset = AdoRecordSet3.Clone
         SetDataListWidth
         GRD1.Visible = True
         cmdOK(1).Enabled = True
      Else
         ShowNoData
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040333 = Nothing
End Sub

Private Sub grd1_SelChange()
Dim m_mouseRow As Integer
GRD1.Visible = False
m_mouseRow = GRD1.MouseRow
GRD1.col = 0
If m_mouseRow <> 0 Then
    If m_row <> 0 Then
        GRD1.row = m_row
         For i = 0 To GRD1.Cols - 1
              GRD1.col = i
              If GRD1.CellBackColor = &HFFC0C0 Then
                GRD1.CellBackColor = &H80000018
                GRD1.TextMatrix(m_row, 0) = ""
              Else
                GRD1.CellBackColor = &HFFC0C0
                GRD1.TextMatrix(m_row, 0) = "V"
              End If
        Next i
    End If
    If m_row <> m_mouseRow Then
        GRD1.row = m_mouseRow
        m_row = m_mouseRow
         For i = 0 To GRD1.Cols - 1
              GRD1.col = i
              If GRD1.CellBackColor = &HFFC0C0 Then
                GRD1.CellBackColor = &H80000018
                GRD1.TextMatrix(m_row, 0) = ""
                m_row = 0
              Else
                GRD1.CellBackColor = &HFFC0C0
                GRD1.TextMatrix(m_row, 0) = "V"
              End If
        Next i
    Else
        m_row = 0
    End If
End If
GRD1.Visible = True
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub SetDataListWidth()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer, m_i As Integer
   
   GRD1.Visible = False
   arrGridHeadText = Array("V", "本所案號", "收文日", "案件性質", "費用" _
             , "規費", "點數", "發文規費", "總收文號", "發文日", "進度備註")
   arrGridHeadWidth = Array(200, 1500, 750, 1000, 800 _
                      , 700, 700, 900, 1000, 750, 6000)
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
'      If iRow > 10 Then
'         grd1.ColWidth(iRow) = 0
'      Else
         GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
'      End If
      GRD1.CellAlignment = flexAlignLeftCenter
   Next
   GRD1.Visible = True
End Sub
