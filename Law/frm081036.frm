VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm081036 
   BorderStyle     =   1  '單線固定
   Caption         =   "TIPS案請款階段分配比例維護作業"
   ClientHeight    =   5280
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7860
   Begin VB.CommandButton Command1 
      Caption         =   "測試用"
      Height          =   300
      Left            =   7968
      TabIndex        =   39
      Top             =   2136
      Width           =   756
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   0
      Left            =   1116
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "ACS"
      Top             =   168
      Width           =   495
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   1
      Left            =   1668
      MaxLength       =   6
      TabIndex        =   1
      Top             =   168
      Width           =   855
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   2
      Left            =   2568
      MaxLength       =   1
      TabIndex        =   2
      Top             =   168
      Width           =   345
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   3
      Left            =   2976
      MaxLength       =   2
      TabIndex        =   3
      Top             =   168
      Width           =   495
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   360
      Left            =   3552
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
      Height          =   2004
      Left            =   48
      TabIndex        =   22
      Top             =   3168
      Width           =   7572
      _ExtentX        =   13356
      _ExtentY        =   3535
      _Version        =   393216
      Cols            =   10
      FormatString    =   "V|請款階段/年度|員工編號|員工姓名|智權人員%|專業點數%|顧服獎金%|協作案號|結算通知日期|主管確認日期"
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&E)"
      Height          =   360
      Left            =   6576
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "人員分配比例維護"
      Height          =   1836
      Left            =   120
      TabIndex        =   21
      Top             =   1224
      Width           =   7404
      Begin VB.TextBox txtData 
         Height          =   300
         Index           =   1
         Left            =   4464
         MaxLength       =   5
         TabIndex        =   12
         Top             =   696
         Width           =   756
      End
      Begin VB.TextBox txtData 
         Height          =   300
         Index           =   0
         Left            =   1056
         MaxLength       =   6
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   696
         Width           =   684
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame4"
         Height          =   732
         Left            =   96
         TabIndex        =   29
         Top             =   1056
         Width           =   7188
         Begin VB.TextBox txtData 
            Height          =   300
            Index           =   2
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   18
            Top             =   384
            Width           =   756
         End
         Begin VB.TextBox txtData 
            Height          =   300
            Index           =   3
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   19
            Top             =   384
            Width           =   756
         End
         Begin VB.TextBox txtSP 
            Height          =   300
            Index           =   3
            Left            =   3204
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   16
            Top             =   8
            Width           =   495
         End
         Begin VB.TextBox txtSP 
            Height          =   300
            Index           =   2
            Left            =   2796
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   15
            Top             =   8
            Width           =   345
         End
         Begin VB.TextBox txtSP 
            Height          =   300
            Index           =   1
            Left            =   1896
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   14
            Top             =   8
            Width           =   855
         End
         Begin VB.TextBox txtSP 
            Height          =   300
            Index           =   0
            Left            =   1344
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   13
            Top             =   8
            Width           =   495
         End
         Begin VB.CommandButton CmdPS 
            Caption         =   "收文號-隱藏"
            Height          =   300
            Left            =   3720
            Style           =   1  '圖片外觀
            TabIndex        =   17
            Top             =   8
            Visible         =   0   'False
            Width           =   1236
         End
         Begin VB.TextBox txtData 
            Height          =   300
            Index           =   4
            Left            =   6408
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   20
            Top             =   384
            Width           =   540
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   1656
            X2              =   3288
            Y1              =   144
            Y2              =   144
         End
         Begin MSForms.Label lblFM2 
            Height          =   252
            Index           =   3
            Left            =   5832
            TabIndex        =   35
            Top             =   32
            Width           =   900
            BackColor       =   12648447
            Size            =   "1587;444"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "專業業務獎金(顧服獎金)比例："
            Height          =   252
            Index           =   11
            Left            =   2136
            TabIndex        =   34
            Top             =   408
            Width           =   2580
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "專業點數比例："
            Height          =   252
            Index           =   12
            Left            =   0
            TabIndex        =   33
            Top             =   408
            Width           =   1308
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "智財協作案號："
            Height          =   252
            Index           =   1
            Left            =   0
            TabIndex        =   32
            Top             =   32
            Width           =   1308
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "收文號："
            Height          =   252
            Index           =   4
            Left            =   5016
            TabIndex        =   31
            Top             =   32
            Width           =   804
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "請款年度："
            Height          =   252
            Index           =   6
            Left            =   5496
            TabIndex        =   30
            Top             =   408
            Width           =   924
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         Height          =   348
         Left            =   240
         TabIndex        =   28
         Top             =   216
         Width           =   3996
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "智權人員"
            Height          =   228
            Index           =   0
            Left            =   192
            TabIndex        =   7
            Top             =   96
            Width           =   1068
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "智財協作"
            Height          =   228
            Index           =   1
            Left            =   2376
            TabIndex        =   8
            Top             =   96
            Width           =   1116
         End
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "刪除-隱藏"
         Height          =   300
         Index           =   1
         Left            =   6288
         Style           =   1  '圖片外觀
         TabIndex        =   10
         Top             =   552
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "存檔"
         Height          =   300
         Index           =   0
         Left            =   6288
         Style           =   1  '圖片外觀
         TabIndex        =   9
         Top             =   192
         Width           =   900
      End
      Begin MSForms.Label lblFM2 
         Height          =   252
         Index           =   2
         Left            =   1776
         TabIndex        =   38
         Top             =   720
         Width           =   900
         BackColor       =   16777215
         Size            =   "1587;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "智權業績獎金比例："
         Height          =   252
         Index           =   10
         Left            =   2832
         TabIndex        =   37
         Top             =   720
         Width           =   1644
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "員工編號："
         Height          =   252
         Index           =   9
         Left            =   96
         TabIndex        =   36
         Top             =   720
         Width           =   924
      End
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   228
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   216
      Width           =   948
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   228
      Index           =   2
      Left            =   120
      TabIndex        =   26
      Top             =   864
      Width           =   948
   End
   Begin VB.Label Label1 
      Caption         =   "當事人1："
      Height          =   228
      Index           =   3
      Left            =   120
      TabIndex        =   25
      Top             =   540
      Width           =   888
   End
   Begin MSForms.Label lblFM2 
      Height          =   264
      Index           =   0
      Left            =   1140
      TabIndex        =   24
      Top             =   540
      Width           =   888
      Size            =   "1566;459"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   264
      Index           =   1
      Left            =   2040
      TabIndex        =   23
      Top             =   540
      Width           =   5532
      Size            =   "9758;459"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1140
      TabIndex        =   6
      Top             =   840
      Width           =   6612
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11668;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1476
      X2              =   3066
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "frm081036"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2025/04/18
Option Explicit
Dim intLastRow As Integer '記錄MGrid1勾選最後一筆
Dim m_ATR(1 To 12) As String
Dim strQuery As String, intQ As Integer
Dim rsQuery As New ADODB.Recordset
Dim oObj As Object
Dim colATR05 As Integer, colATR06 As Integer, colATR07 As Integer, colATR08 As Integer
Dim colATR09 As Integer, colATR10 As Integer
Dim colSCP01 As Integer, colSCP02 As Integer, colSCP03 As Integer, colSCP04 As Integer
Dim colPKno As Integer, m_PKno As String
Dim m_AT112 As String '○○○年已有年度結算
Dim bolUpdATR08 As Boolean '是否可修改智權人員比例(智權業務獎金比例)

'(保留) 按鈕隱藏
Private Sub CmdPS_Click()

   If Trim(m_ATR(5) & m_ATR(6) & m_ATR(7)) <> "" Then
      Exit Sub
   End If
   If txtSP(0) = "" Or Len(txtSP(1)) < 6 Then
      MsgBox "請輸入智財協作案號！", vbExclamation, "檢核資料"
      Exit Sub
   End If
   If Trim(txtSP(2)) = "" Then txtSP(2) = "0"
   If Trim(txtSP(3)) = "" Then txtSP(3) = "00"
   
   '接洽單收文PS及CPS之智財協作967，TT及S之智財協作737，L之智財協作7601，在收文時一定要輸「本案與總號之有關」欄，且必須要有ACS且為TIPS的案件，
   If txtSP(0) = "L" Then
      strExc(0) = "select '' v,cp01||'-'||cp02||decode(cp03||cp04,'000',null,'-'||cp03||'-'||cp04) as caseno,nvl(lc05,nvl(lc06,lc07)) casename, " & _
                  "cp09,decode(lc15,'000',nvl(cpm03,cpm04),nvl(cpm04,cpm03)) cpm0304,substr(sqldatet(cp05),1,10) as cp05t,substr(SqlDateT(Cp27), 1, 10) As cp27t,st02,cp14 " & _
                  "From caseprogress, lawcase, casepropertymap,caserelation1,staff " & _
                  "where cp01='" & txtSP(0) & "' and cp02='" & txtSP(1) & "' and cp03='" & txtSP(2) & "' and cp04='" & txtSP(3) & "' " & _
                  "and cp158 > 0 and cp159=0 and cp10='7601' and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) " & _
                  "and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp09 not in (select atr06 from acs_tips_rate) " & _
                  "and cp01=cr01(+) and cp02=cr02(+) and cp03=cr03(+) and cp04=cr04(+) and cr05='" & txtCase(0) & "' and cr06='" & txtCase(1) & "' and cr07='" & txtCase(2) & "' and cr08='" & txtCase(3) & "' " & _
                  "order by cp27 desc "
   Else
      strExc(0) = "select '' v,cp01||'-'||cp02||decode(cp03||cp04,'000',null,'-'||cp03||'-'||cp04) as caseno,nvl(sp05,nvl(sp06,sp07)) casename, " & _
                  "cp09,decode(sp09,'000',nvl(cpm03,cpm04),nvl(cpm04,cpm03)) cpm0304,substr(sqldatet(cp05),1,10) as cp05t,substr(SqlDateT(Cp27), 1, 10) As cp27t,st02,cp14 " & _
                  "From caseprogress, servicepractice, casepropertymap,caserelation1,staff " & _
                  "where cp01='" & txtSP(0) & "' and cp02='" & txtSP(1) & "' and cp03='" & txtSP(2) & "' and cp04='" & txtSP(3) & "' " & _
                  "and cp158 > 0 and cp159=0 and cp10='" & IIf(txtSP(0) = "PS" Or txtSP(0) = "CPS", "967", "737") & "' and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) " & _
                  "and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp09 not in (select atr06 from acs_tips_rate) " & _
                  "and cp01=cr01(+) and cp02=cr02(+) and cp03=cr03(+) and cp04=cr04(+) and cr05='" & txtCase(0) & "' and cr06='" & txtCase(1) & "' and cr07='" & txtCase(2) & "' and cr08='" & txtCase(3) & "' " & _
                  "order by cp27 desc "
   End If
   Me.Tag = ""
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      MsgBox "智財協作案號不存在關聯的收文號！", vbExclamation, "檢核資料"
      Exit Sub
   Else
      Set frm880012.grdDataList.Recordset = RsTemp
      Set frm880012.fmParent = Me
      frm880012.iTyp = "8"
      frm880012.Show vbModal
      If Me.Tag <> "" Then  '回傳請款年度3碼+CP09+CP14
         txtData(0) = Mid(Me.Tag, 13)
         If ClsPDGetStaff(txtData(0), strExc(1)) = True Then
            txtData(4) = Mid(Me.Tag, 1, 3)
            lblFM2(3) = Mid(Me.Tag, 4, 9)
            lblFM2(2) = strExc(1)
         Else
            lblFM2(2) = ""
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim bolTmp As Boolean, intErr As Integer
   
   intErr = -1
   If Trim(txtCase(0).Text) = "" Or Len(Trim(txtCase(1).Text)) < 6 Then
      MsgBox "請輸入本所案號！", vbExclamation, "檢核資料"
      Exit Sub
   End If
   If m_ATR(1) & m_ATR(2) & m_ATR(3) & m_ATR(4) <> txtCase(0) & txtCase(1) & txtCase(2) & txtCase(3) Then
      MsgBox "輸入本所案號後，請執行查詢功能！", vbExclamation, "檢核資料"
      Exit Sub
   End If
   If "" & m_ATR(12) <> "" Then
      MsgBox "主管已確認，不可變更！", vbExclamation, "檢核資料"
      Exit Sub
   End If
   
   '是否已有年度結算過
   m_AT112 = ""
   strExc(0) = "select at112,at114 from acs_tips_rate1 where at101='" & m_ATR(1) & "' and at102='" & m_ATR(2) & "' and at103='" & m_ATR(3) & "' and at104='" & m_ATR(4) & "' and at105='" & txtData(4) & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '已有年度結算過
   If intI = 1 Then
      If "" & RsTemp.Fields("at112") <> "" Then
         m_AT112 = "" & RsTemp.Fields("at112")
         If "" & RsTemp.Fields("at114") <> "" Then
            MsgBox txtData(4) & "年主管已確認，不可變更！", vbExclamation, "檢核資料"
            Exit Sub
         End If
      End If
   End If
'----------------------------------------------------
   If Index = 1 Then  '刪除:按鈕隱藏
      
      If m_PKno = "" Then
         MsgBox "尚未新增記錄！", vbExclamation, "檢核資料"
         Exit Sub
      Else
         If Val(m_ATR(8)) + Val(m_ATR(9)) + Val(m_ATR(10)) > 0 Then
            MsgBox "已有分配比例，不可刪除記錄！", vbExclamation, "檢核資料"
            Exit Sub
         End If
         If MsgBox("是否要刪除記錄？", vbExclamation + vbYesNo + vbDefaultButton2, "檢核資料") = vbNo Then
            Exit Sub
         End If
         If SaveData("D") = True Then
            QueryData2
         End If
      End If
      
   ElseIf Index = 0 Then '新增/修改=存檔
'----------------------------------------------------
      If Option1(0).Value = False And Option1(1).Value = False Then
         MsgBox "請選擇「智權人員」或「智財協作」進行維護！", vbExclamation, "檢核資料"
         Exit Sub
      End If
      intErr = 0
      If Trim(txtData(0)) = "" Then
         MsgBox "員工編號不可空白！", vbExclamation, "檢核資料"
         GoTo EXITSUB
      End If
      If Option1(0).Value = True Then
         intErr = 1
         If Val(txtData(1)) = 0 Then
            MsgBox "智權業績獎金比例不可空白！", vbExclamation, "檢核資料"
            GoTo EXITSUB
         Else
            Call Txtdata_Validate(0, bolTmp)
            If bolTmp = True Then
               GoTo EXITSUB
            End If
         End If
         m_ATR(5) = "1"
         If m_ATR(6) = "" Then
            '不限制智權人員---拿掉and instr(nvl(atr07,'N'),'" & Trim(txtData(0)) & "')=0
            strSql = "select cp09,atr07 from caseprogress,acs_tips_rate where cp01='" & m_ATR(1) & "' and cp02='" & m_ATR(2) & "' and cp03='" & m_ATR(3) & "' and cp04='" & m_ATR(4) & "' and nvl(cp156,0)=1 and cp159=0 " & _
                     "and cp09=atr06(+) "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If "" & RsTemp.Fields("atr07") <> "" Then
                  MsgBox "已存在智權業績獎金比例，不可重複新增！", vbExclamation, "檢核資料"
                  GoTo EXITSUB
               End If
               m_ATR(6) = "" & RsTemp.Fields("cp09")
            End If
         End If
         m_ATR(8) = Val(txtData(1))
      Else
         If txtSP(0) = "" Or txtSP(1) = "" Or txtSP(2) = "" Or txtSP(3) = "" Or lblFM2(3) = "" Then
            MsgBox "請輸入智財協案號的收文號！", vbExclamation, "檢核資料"
            GoTo EXITSUB
         End If
         
         intErr = 4
         If Val(txtData(4)) = 0 Then
            MsgBox "請款年度不可空白！", vbExclamation, "檢核資料"
            GoTo EXITSUB
         Else
            Call Txtdata_Validate(4, bolTmp)
            If bolTmp = True Then
               GoTo EXITSUB
            End If
         End If
   
         If Val(txtData(2)) = 0 And Val(txtData(3)) = 0 Then
            If MsgBox("專業點數比例和專業業務獎金(顧服獎金)未輸入，是否繼續存檔作業？", vbInformation + vbYesNo + vbDefaultButton2, "檢核資料") = vbNo Then
               intErr = 2
               GoTo EXITSUB
            End If
         Else
            If Val(txtData(2)) > 20 Then
               MsgBox "專業點數比例不可超過20%！", vbExclamation, "檢核資料"
               intErr = 2
               GoTo EXITSUB
            Else
               If Val(txtData(2)) > 10 Then
                  If MsgBox("專業點數比例超過10%，是否繼續存檔作業？", vbInformation + vbYesNo + vbDefaultButton2, "檢核資料") = vbNo Then
                     intErr = 2
                     GoTo EXITSUB
                  End If
               End If
            End If
            If Val(txtData(3)) > 20 Then
               MsgBox "專業業務獎金(顧服獎金)比例不可超過20%！", vbExclamation, "檢核資料"
               intErr = 3
               GoTo EXITSUB
            Else
               If Val(txtData(3)) > 10 Then
                  If MsgBox("專業業務獎金(顧服獎金)比例超過10%，是否繼續存檔作業？", vbInformation + vbYesNo + vbDefaultButton2, "檢核資料") = vbNo Then
                     intErr = 3
                     GoTo EXITSUB
                  End If
               End If
            End If

            strSql = "select nvl(sum(atr09),0) atr09,nvl(sum(atr10),0) atr10 from acs_tips_rate where atr01='" & m_ATR(1) & "' and atr02='" & m_ATR(2) & "' and atr03='" & m_ATR(3) & "' and atr04='" & m_ATR(4) & "' " & _
                     "and atr05>100 and atr05=" & txtData(4).Text & " and atr05||atr06||atr07<>'" & txtData(4).Tag & lblFM2(3).Tag & txtData(0).Tag & "' "
            intI = 1

            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               strExc(9) = Val("" & RsTemp.Fields("atr09"))
               strExc(10) = Val("" & RsTemp.Fields("atr09"))
               If Val(txtData(2)) + Val("" & RsTemp.Fields("atr09")) > 100 Then
                  MsgBox "目前輸入的專業點數比例已超過100%！", vbExclamation, "檢核資料"
                  intErr = 2
                  GoTo EXITSUB
               End If
               If Val(txtData(3)) + Val("" & RsTemp.Fields("atr10")) > 100 Then
                  MsgBox "目前輸入的專業業務獎金(顧服獎金)比例比例已超過100%！", vbExclamation, "檢核資料"
                  intErr = 3
                  GoTo EXITSUB
               End If
            End If
         End If
         m_ATR(5) = Val(txtData(4))
         m_ATR(6) = lblFM2(3)
         m_ATR(9) = Val(txtData(2))
         m_ATR(10) = Val(txtData(3))
      End If
      
      m_ATR(7) = txtData(0)
      If m_ATR(5) = "" Or m_ATR(6) = "" Or m_ATR(7) = "" Then
         For intI = 5 To UBound(m_ATR)
            m_ATR(intI) = ""
         Next
         Exit Sub
      End If
      CmdOK(0).Enabled = False
      'cmdOK(1).Enabled = False '(保留)按鈕隱藏
      Screen.MousePointer = vbHourglass
      If SaveData(IIf(m_PKno = "", "A", "U")) = True Then
         QueryData2
         CmdOK(0).Enabled = True
         'cmdOK(1).Enabled = True '(保留)按鈕隱藏
      End If
      Screen.MousePointer = vbDefault
   End If
   
   Exit Sub
   
EXITSUB:
   If intErr > -1 Then
      txtData(intErr).SetFocus
      Txtdata_GotFocus intErr
   End If
End Sub

Private Function SaveData(ByVal pType As String)
Dim strUpd As String, strCon As String
Dim intB As Integer, strB1 As String
Dim rsBD As New ADODB.Recordset
   
   If pType = "U" Then
      If txtData(4).Text <> txtData(4).Tag Then
         strUpd = strUpd & ", ATR05='" & m_ATR(5) & "'"
      End If
      If lblFM2(3).Tag <> lblFM2(3) Then
         strUpd = strUpd & ", ATR06='" & m_ATR(6) & "'"
      End If
      If txtData(0).Text <> txtData(0).Tag Then
         strUpd = strUpd & ", ATR07='" & m_ATR(7) & "'"
      End If
      If Option1(0).Value = True Then
         If txtData(1).Tag <> txtData(1) Then
            strUpd = strUpd & ", ATR08=" & CNULL(m_ATR(8), True)
         End If
      Else
         If txtData(2).Tag <> txtData(2) Then
            strUpd = strUpd & ", ATR09=" & CNULL(m_ATR(9), True)
         End If
         If txtData(3).Tag <> txtData(3) Then
            strUpd = strUpd & ", ATR10=" & CNULL(m_ATR(10), True)
         End If
      End If
      
      If Option1(0).Value = True Then
         strCon = " and atr05='" & m_ATR(5) & "' and atr06='" & m_ATR(6) & "' and atr07='" & txtData(0).Tag & "' "
      Else
         strCon = " and atr05='" & txtData(4).Tag & "' and atr06='" & lblFM2(3).Tag & "' and atr07='" & txtData(0).Tag & "' "
      End If

      If strUpd <> "" Then
         strUpd = "Update Acs_Tips_Rate Set " & Mid(strUpd, 2) & ", atr13='" & strUserNum & "', atr14=sysdate Where atr01='" & m_ATR(1) & "'and atr02='" & m_ATR(2) & "' and atr03='" & m_ATR(3) & "' and atr04='" & m_ATR(4) & "' " & strCon
      End If
   Else
      '(保留)因為第一次繳款作業和智財協作案件發文都會產生空白的對應記錄,所以沒有新增; 同時拿掉刪除功能
      'If pType = "A" Then
      '   strUpd = "Insert Into Acs_Tips_Rate (ATR01,ATR02,ATR03,ATR04,ATR05,ATR06,ATR07,ATR08,ATR09,ATR10,ATR13,ATR14) " & _
      '            "Values ('" & m_ATR(1) & "','" & m_ATR(2) & "','" & m_ATR(3) & "','" & m_ATR(4) & "','" & m_ATR(5) & "' " & _
      '            ",'" & m_ATR(6) & "','" & m_ATR(7) & "'," & CNULL(m_ATR(8), True) & "," & CNULL(m_ATR(9), True) & "," & CNULL(m_ATR(10), True) & _
      '            ",'" & strUserNum & "', SYSDATE) "
      'Else
      '   strUpd = "Delete from Acs_Tips_Rate where atr01='" & m_ATR(1) & "' and atr02='" & m_ATR(2) & "' and atr03='" & m_ATR(3) & "' and atr04='" & m_ATR(4) & "' and atr05='" & m_ATR(5) & "' and atr06='" & m_ATR(6) & "' and atr07='" & m_ATR(7) & "' "
      'End If
   End If
   
   If strUpd <> "" Then
      cnnConnection.BeginTrans
      
         cnnConnection.Execute strUpd
         '已有年度結算過>>抓○○○年最後一階段款項的收據號碼
         If m_AT112 <> "" Then
            strCon = "select cp09,cp156,cp115,a0j01,a0j13 From caseprogress, acc0j0 " & _
                     "where cp01='" & m_ATR(1) & "' and cp02='" & m_ATR(2) & "' and cp03='" & m_ATR(3) & "' and cp04='" & m_ATR(4) & "' " & _
                     "and cp156||cp09=(select max(cp156||cp09) mno from caseprogress " & _
                     "where cp01='" & m_ATR(1) & "' and cp02='" & m_ATR(2) & "' and cp03='" & m_ATR(3) & "' and cp04='" & m_ATR(4) & "' and nvl(cp156,0) > 0 and cp115='" & m_ATR(5) & "') and cp60=a0j13(+) "
            intB = 1
            Set rsBD = ClsLawReadRstMsg(intB, strCon)
            If intB = 1 Then
               If "" & rsBD.Fields("a0j01") <> "" And "" & rsBD.Fields("a0j13") <> "" Then
                  Call PUB_ProcAcs_Tips_Rate1(False, m_ATR(1), m_ATR(2), m_ATR(3), m_ATR(4), "" & rsBD.Fields("a0j01"), "" & rsBD.Fields("a0j13"))
               End If
            End If
         End If
         '------------
         If pType <> "D" Then
            If m_ATR(5) = "1" Then
               '顧服組主管輸入智權比例後，通知【正本：財務處；副本:智權人員、智權部主管】
               strExc(0) = m_ATR(1) & "-" & m_ATR(2) & IIf(m_ATR(3) & m_ATR(4) = "000", "", "-" & m_ATR(3) & "-" & m_ATR(4))
               strExc(1) = strExc(0) & "智權業務點數比例為100%" & vbCrLf & _
                           strExc(0) & "智權業績獎金比例為" & m_ATR(8) & "%，專業業務獎金(顧服獎金)比例為" & 100 - Val(m_ATR(8)) & "%，後續階段繳款請依前述比例分配。"
               strExc(2) = Pub_GetSpecMan("財務處出納人員")
               If strExc(2) <> "" Then
                   strExc(3) = Pub_GetSpecMan("全所智權部主管")
                   strCon = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09) " & _
                            "values('" & strUserNum & "','" & strExc(2) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss') " & _
                            ",'" & strExc(0) & "第一階段款項已繳款，請依系統輸入智權人員比例分配點數','" & ChgSQL(strExc(1)) & "','" & m_ATR(7) & IIf(strExc(3) <> "", ";" & strExc(3), "") & "') "
                   cnnConnection.Execute strCon
               End If
            'Modfied by Lydia 2025/05/19　改寫法；若智權人員比例後輸，要重新發送智智財協作分配比例Email
            'Else
            End If
               If m_AT112 = "" Then '已年度結算後再修改比例不發Email
                  'Modified by Lydia 2025/05/19 改寫法
                  ''已於系統輸入協作分配比例，通知【正本：承辦人，副本：承辦人主管、顧服組主管】
                  'strExc(0) = txtSP(0) & "-" & txtSP(1) & IIf(txtSP(2) & txtSP(3) = "000", "", "-" & txtSP(2) & "-" & txtSP(3))
                  'strExc(1) = "請智財協作部門確認專業點數比例以及專業業務獎金(顧服獎金)比例" & vbCrLf & _
                  '            strExc(0) & "專業點數比例為" & Val(txtData(2)) & "%，專業業務獎金(顧服獎金)比例為" & Val(txtData(3)) & "%，" & vbCrLf & _
                  '           "當年度專案結束後，將加總各次智財協作占比，進行點數分配，如對本次分配比例有任何建議，請盡速通知顧服組主管。"
                  'strExc(2) = ""
                  'strExc(2) = Pub_GetSpecMan(txtSP(0) & "智財協作之協作部門主管")
                  'strExc(2) = strExc(2) & IIf(strExc(2) <> "", ";", "") & Pub_GetSpecMan("ACS郵件通知主管")
                  'strCon = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09) " & _
                  '         "values('" & strUserNum & "','" & Trim(txtData(0)) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss') " & _
                  '         ",'" & m_ATR(1) & "-" & m_ATR(2) & IIf(m_ATR(3) & m_ATR(4) = "000", "", "-" & m_ATR(3) & "-" & m_ATR(4)) & "<" & strExc(0) & ">已於系統輸入協作分配比例，請確認','" & ChgSQL(strExc(1)) & "','" & strExc(2) & "') "
                  'cnnConnection.Execute strCon
                  strCon = "select atr05,atr06,atr07,atr09,atr10,cp01,cp02,cp03,cp04,b06,fee1,fee2 from acs_tips_rate,caseprogress " & _
                            ",(select cp01 as b01, cp02 as b02, cp03 as b03, cp04 as b04 ,cp115 as b05,nvl(atr08,0)/100 as b06,sum(nvl(a0j09,0))/1000 as fee1,sum(nvl(a0j10,0))/1000 as fee2 " & _
                            "from caseprogress, acc0j0,acs_tips_rate where cp01='" & txtCase(0) & "' and cp02='" & txtCase(1) & "' and cp03='" & txtCase(2) & "' and cp04='" & txtCase(3) & "' and cp60=a0j13(+) and nvl(cp115,0)>100 " & _
                            "and cp01=atr01(+) and cp02=atr02(+) and cp03=atr03(+) and cp04=atr04(+) and '1'=atr05(+) group by cp01,cp02,cp03,cp04,cp115,nvl(atr08,0)) vtb01 " & _
                            "where atr01='" & txtCase(0) & "' and atr02='" & txtCase(1) & "' and atr03='" & txtCase(2) & "' and atr04='" & txtCase(3) & "' and atr05>100 " & _
                            "and atr12 is null and nvl(atr09,0)+nvl(atr10,0) > 0 and atr06=cp09(+) " & _
                            "and atr01=b01(+) and atr02=b02(+) and atr03=b03(+) and atr04=b04(+) and atr05=b05(+) " & _
                            IIf(m_ATR(5) = "1", "", "and atr06='" & lblFM2(3) & "'")
                  strCon = strCon & "order by atr05,cp01,cp05 "
                  intB = 1
                  Set rsBD = ClsLawReadRstMsg(intB, strCon)
                  If intB = 1 Then
                     rsBD.MoveFirst
                     Do While Not rsBD.EOF
                        strExc(0) = rsBD.Fields("cp01") & "-" & rsBD.Fields("cp02") & IIf(rsBD.Fields("cp03") & rsBD.Fields("cp04") = "000", "", "-" & rsBD.Fields("cp03") & "-" & rsBD.Fields("cp04"))
                        '已於系統輸入協作分配比例，通知【正本：承辦人，副本：承辦人主管、顧服組主管】
                        strExc(9) = Round(Val("" & rsBD.Fields("fee1")) * (Val("" & rsBD.Fields("atr09")) / 100), 2)
                        If Val("" & rsBD.Fields("b06")) = 0 Then
                           strExc(10) = "尚未計算"
                        Else
                           strExc(10) = Round(Val("" & rsBD.Fields("fee1")) * (1 - Val("" & rsBD.Fields("b06"))) * (Val("" & rsBD.Fields("atr10")) / 100), 2)
                        End If
                        strExc(1) = "請智財協作部門確認專業點數比例以及專業業務獎金(顧服獎金)比例" & vbCrLf & _
                                    strExc(0) & "專業點數比例為" & Val("" & rsBD.Fields("atr09")) & "%(" & strExc(9) & "點)，專業業務獎金(顧服獎金)比例為" & Val("" & rsBD.Fields("atr10")) & "%(" & strExc(10) & "點)，" & vbCrLf & _
                                   "當年度專案結束後，將加總各次智財協作占比，進行點數分配，如對本次分配比例有任何建議，請盡速通知顧服組主管。"
                        strExc(2) = ""
                        strExc(2) = Pub_GetSpecMan(rsBD.Fields("cp01") & "智財協作之協作部門主管")
                        strExc(2) = strExc(2) & IIf(strExc(2) <> "", ";", "") & Pub_GetSpecMan("ACS郵件通知主管")
                        strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09) " & _
                                 "values('" & strUserNum & "','" & Trim("" & rsBD.Fields("atr07")) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss') " & _
                                 ",'" & m_ATR(1) & "-" & m_ATR(2) & IIf(m_ATR(3) & m_ATR(4) = "000", "", "-" & m_ATR(3) & "-" & m_ATR(4)) & "<" & strExc(0) & ">已於系統輸入協作分配比例，請確認','" & ChgSQL(strExc(1)) & "','" & strExc(2) & "') "
                        cnnConnection.Execute strSql
                        
                        rsBD.MoveNext
                     Loop
                  End If
                  'end 2025/05/19
               End If  'If m_AT112 = "" Then '已年度結算後再修改比例不發Email
            'End If  'Mark by Lydia 2025/05/19
         End If
      cnnConnection.CommitTrans
   End If
   Set rsBD = Nothing
'--------------------------------------
   PUB_SendMailCache
   
   SaveData = True
   Exit Function
   
ErrHandle:
   If strUpd <> "" Then
      cnnConnection.RollbackTrans
   End If
   If Err.Number <> 0 Then
      MsgBox IIf(pType = "A", "新增", IIf(pType = "D", "修改", "刪除")) & "失敗：" & Err.Description
   End If
End Function

Private Sub cmdQuery_Click()
   
   If Len(txtCase(1)) <> 6 Then
      MsgBox "請輸入本所案號！！", vbExclamation
      txtCase(1).SetFocus
      txtCase_GotFocus 1
      Exit Sub
   End If
   If Trim(txtCase(2)) = "" Then txtCase(2) = "0"
   If Trim(txtCase(3)) = "" Then txtCase(3) = "00"
   ClearData1 False

   strQuery = "select lc01,lc02,lc03,lc04,lc05,lc06,lc07,lc11 as custno,nvl(cu04,nvl(cu05,cu06)) custname,cp10 " & _
             "from caseprogress,lawcase,customer " & _
             "where cp01='" & txtCase(0) & "' and cp02='" & txtCase(1) & "' and cp03='" & txtCase(2) & "' and cp04='" & txtCase(3) & "' " & _
             "and cp31='Y' and cp159=0 and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) " & _
             "and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) "
   intQ = 0
   Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 0 Then
       Exit Sub
   Else
       If InStr(ACSforTIPSstep, "'" & rsQuery.Fields("cp10") & "'") = 0 Then
           MsgBox "查無TIPS進度！", vbInformation
           txtCase(1).SetFocus
           txtCase_GotFocus 1
           Exit Sub
       End If
       m_ATR(1) = "" & rsQuery.Fields("lc01")
       m_ATR(2) = "" & rsQuery.Fields("lc02")
       m_ATR(3) = "" & rsQuery.Fields("lc03")
       m_ATR(4) = "" & rsQuery.Fields("lc04")
       intQ = 0
       Combo1.AddItem "中：" & rsQuery.Fields("lc05"), 0
       If "" & rsQuery.Fields("lc05") <> "" And intQ = 0 Then intQ = 1
       Combo1.AddItem "英：" & rsQuery.Fields("lc06"), 1
       If "" & rsQuery.Fields("lc06") <> "" And intQ = 0 Then intQ = 2
       Combo1.AddItem "日：" & rsQuery.Fields("lc07"), 2
       If "" & rsQuery.Fields("lc07") <> "" And intQ = 0 Then intQ = 3
       Combo1.ListIndex = intQ - 1
       lblFM2(0).Caption = "" & rsQuery.Fields("custno")
       lblFM2(1).Caption = "" & rsQuery.Fields("custname")
       
       '第1次繳款先發mail通知設定智權人員比例(智權業務獎金比例)，財務處收款後不可再修改
       strExc(0) = "select * from (select cp09,cp10,a0j13,a0j01,nvl(sum(nvl(a1u04,0)+nvl(a1u07,0)-nvl(a1u08,0)),0) x3,nvl(sum(nvl(a1u05,0)+nvl(a1u09,0)-nvl(a1u10,0)),0) x4" & _
                   " From caseprogress, acc0j0, acc0k0, acc1u0" & _
                   " where cp01='" & m_ATR(1) & "' and cp02='" & m_ATR(2) & "' and cp03='" & m_ATR(3) & "' and cp04='" & m_ATR(4) & "'" & _
                   " and cp09=a0j01(+) and a0j13=a0k01(+) and nvl(a0k09,0)=0 and a0j13=a1u02 and a0j01=a1u03" & _
                   " group by cp09,cp10,a0j13,a0j01) where x3 >0 or x4 >0"
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
          bolUpdATR08 = False
       Else
          bolUpdATR08 = True
       End If
       
       QueryData2
   End If
   TxtLocked False

End Sub
'測試用
Private Sub Command1_Click()
   'Call PUB_ProcAcs_Tips_Rate1(True, "ACS", "000230", "0", "00", "AB3046333", "E11323707")
   'Call PUB_ProcAcs_Tips_Rate1(True, "ACS", "000207", "0", "00", "AB3008308", "E11312427")
   ' Call PUB_ProcAcs_Tips_Rate1(True, "ACS", "000242", "0", "00", "AB4006487", "E11404999") '當年度只有一個請款階段
   Call PUB_ProcAcs_Tips_Rate1(True, "ACS", "000227", "0", "00", "AB3036947", "E11319326")
   PUB_SendMailCache
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   Frame2.BackColor = &HC0FFFF
   Frame3.BackColor = &HC0FFFF
   
   SetGrd1 True
   ClearData1 True
   TxtLocked True

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsQuery = Nothing
   
   Set frm081036 = Nothing
End Sub

Private Sub ClearData1(ByVal bolAll As Boolean)
   If bolAll = True Then
      For Each oObj In txtCase
         If oObj.Index = 0 Then
            oObj.Text = "ACS"
         Else
            oObj.Text = ""
         End If
         oObj.Tag = oObj.Text
      Next
      m_ATR(1) = "": m_ATR(2) = "": m_ATR(3) = "": m_ATR(4) = ""
      bolUpdATR08 = True
   End If
   Combo1.Clear
   For Each oObj In lblFM2
      oObj.Caption = ""
      oObj.Tag = ""
   Next
   intLastRow = 0
   ClearData2
End Sub

Private Sub ClearData2()

   For intI = 5 To UBound(m_ATR)
      m_ATR(intI) = ""
   Next intI
   m_PKno = ""
   m_AT112 = ""
   lblFM2(2).Caption = ""
   lblFM2(3).Caption = ""
   lblFM2(2).Tag = ""
   lblFM2(3).Tag = ""
   Option1(0).Value = 0
   Option1(1).Value = 0
   For Each oObj In txtData
      oObj.Text = ""
      oObj.Tag = ""
   Next
   For Each oObj In txtSP
      oObj.Text = ""
      oObj.Tag = ""
   Next

End Sub

Private Sub TxtLocked(Optional ByVal bLocked As Boolean)
   '主管已確認／尚未查詢
   If Val(m_ATR(12)) > 0 Or bLocked = True Then
      CmdOK(0).Visible = False
      'CmdOK(1).Visible = False: CmdPS.Visible = False '(保留) 按鈕隱藏
      Frame2.Enabled = False
      txtData(0).Locked = True
      txtData(1).Locked = True
      Frame3.Enabled = False
   Else
      If bLocked = False Then
         CmdOK(0).Visible = True
         '(保留) 按鈕隱藏
         'cmdOK(1).Visible = True
         'If Val(m_ATR(5)) = 0 And Option1(0).Value = 0 And Option1(1) = 0 Then
         '   txtData(0).Locked = False
         '   txtData(1).Locked = False
         '   Frame2.Enabled = True
         '   Frame3.Enabled = True
         '   For intI = 0 To 3
         '      txtSP(intI).Locked = False
         '   Next
         'Else
            'CmdPS.Visible = False '(保留) 按鈕隱藏
            txtData(0).Locked = False
            Frame2.Enabled = False
            '(保留) 按鈕隱藏
            'For intI = 0 To 3
            '   txtSP(intI).Locked = True
            'Next
            '智權人員
            If Val(m_ATR(5)) < 100 Then
               '顧服組通知財務處智權人員比例後，就可鎖住；要到財務處收款後不可再修改
               If Val(m_ATR(8)) > 0 And bolUpdATR08 = False Then
                  txtData(0).Locked = True
                  txtData(1).Locked = True
               Else
                  txtData(0).Locked = False
                  txtData(1).Locked = False
               End If
               Frame3.Enabled = False
            '智財協作
            Else
               txtData(1).Locked = True
               Frame3.Enabled = True
            End If
         'End If '(保留) 按鈕隱藏
      End If
   End If
End Sub
Private Sub QueryData2()

    ClearData2
    SetGrd1 True
    TxtLocked False
    
    If m_ATR(1) <> "" And m_ATR(2) <> "" Then
       strQuery = "select '' v,atr05,atr06,atr07,st02,atr08,atr09,atr10,decode(atr01||atr02,c2.cp01||c2.cp02,null,c2.cp01||'-'||c2.cp02||decode(c2.cp03||c2.cp04,'000',null,'-'||c2.cp03||'-'||c2.cp04)) as scaseno, " & _
                 "substr(sqldatet2(at112),1,10) as mdate,substr(sqldatet2(atr12),1,10) as m2date,c2.cp01 as scp01,c2.cp02 as scp02,c2.cp03 as scp03,c2.cp04 as scp04,atr01||atr02||atr03||atr04||atr05||atr06||atr07 as pkno " & _
                 "from acs_tips_rate,acs_tips_rate1 ,staff s1,caseprogress c2 " & _
                 "where atr01='" & m_ATR(1) & "' and atr02='" & m_ATR(2) & "' and atr03='" & m_ATR(3) & "' and atr04='" & m_ATR(4) & "' and atr07=st01(+) and atr06=cp09(+) " & _
                 "and atr01=at101(+) and atr02=at102(+) and atr03=at103(+) and atr04=at104(+) and atr05=at105(+) and atr11=at106(+) "
       '預設：請款階段1-如果智權人員離職改抓最新進度的智權人員
       strQuery = strQuery & "union select '' v,to_char(cp156) cp156,cp09,decode(st04,'1',cp13,substr(v05,9,6)) as cp13,staffname(decode(st04,'1',cp13,substr(v05,9,6))) as st02,null as atr08,null as atr09,null as atr10,null as scaseno,null as atr12t,null as atr14t,cp01,cp02,cp03,cp04,null as pkno " & _
                 "From caseprogress, staff,(select cp01 as v01,cp02 as v02,cp03 as v03,cp04 as v04, max(cp05||cp13) v05 " & _
                  "from caseprogress,staff where cp01='" & m_ATR(1) & "' and cp02='" & m_ATR(2) & "' and cp03='" & m_ATR(3) & "' and cp04='" & m_ATR(4) & "' and cp159=0 and cp09 <'C' and cp13=st01(+) and st04='1' group by cp01,cp02,cp03,cp04) " & _
                 "where cp01='" & m_ATR(1) & "' and cp02='" & m_ATR(2) & "' and cp03='" & m_ATR(3) & "' and cp04='" & m_ATR(4) & "' and nvl(cp156,0)= 1 and cp159=0 and cp13=st01(+) " & _
                 "and cp09 not in (select atr06 from acs_tips_rate where  atr01='" & m_ATR(1) & "' and atr02='" & m_ATR(2) & "' and atr03='" & m_ATR(3) & "' and atr04='" & m_ATR(4) & "') " & _
                 "and cp01=v01(+) and cp02=v02(+) and cp03=v03(+) and cp04=v04(+) "
       
       intQ = 1
       Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
       If intQ = 1 Then
          MGrid1.FixedCols = 0
          Set MGrid1.Recordset = rsQuery
          Call SetGrd1
          MGrid1.FixedCols = 5
       End If
    End If

End Sub

Private Sub SetGrd1(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
 
   arrGridHeadText = Array("V", "階段/年度", "收文號", "員工編號", "員工姓名", _
                   "智權獎金%", "專業點數%", "顧服獎金%", "協作案號", "結算通知日期", _
                   "主管確認日期", "SCP01", "SCP02", "SCP03", "SCP04", _
                   "PKNO")
   arrGridHeadWidth = Array(260, 1000, 0, 1000, 1000, _
                   1000, 1000, 1000, 1000, 1200, _
                   1200, 0, 0, 0, 0, _
                   0)
        
   MGrid1.Visible = False
   MGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
       MGrid1.Clear
       MGrid1.Rows = 2
   End If
       
   For iRow = 0 To MGrid1.Cols - 1
      MGrid1.row = 0
      MGrid1.col = iRow
      MGrid1.Text = arrGridHeadText(iRow)
      MGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MGrid1.CellAlignment = flexAlignCenterCenter
   Next
   If colATR05 = 0 Then
      colATR05 = PUB_MGridGetId("階段/年度", MGrid1)
      colATR06 = PUB_MGridGetId("收文號", MGrid1)
      colATR07 = PUB_MGridGetId("員工編號", MGrid1)
      colATR08 = PUB_MGridGetId("智權獎金%", MGrid1)
      colATR09 = PUB_MGridGetId("專業點數%", MGrid1)
      colATR10 = PUB_MGridGetId("顧服獎金%", MGrid1)
      colSCP01 = PUB_MGridGetId("SCP01", MGrid1)
      colSCP02 = PUB_MGridGetId("SCP02", MGrid1)
      colSCP03 = PUB_MGridGetId("SCP03", MGrid1)
      colSCP04 = PUB_MGridGetId("SCP04", MGrid1)
      colPKno = PUB_MGridGetId("PKNO", MGrid1)
   End If
   
   For intI = 1 To MGrid1.Rows - 1
      MGrid1.row = intI
      For iRow = 1 To MGrid1.Cols - 1
         MGrid1.col = iRow
         '置中
         If InStr("01,05,06,07", Format(iRow, "00")) > 0 Then
            MGrid1.CellAlignment = flexAlignCenterCenter
         End If
      Next iRow
   Next intI
   
   MGrid1.Visible = True
End Sub

Private Sub MGrid1_Click()
Dim intRow As Integer, intCol As Integer
   
   With MGrid1
      If .MouseRow > 0 Then
         intRow = .MouseRow
         intCol = .MouseCol
         .row = intRow
         '----單選
         GridClick MGrid1, intLastRow, 0, 0, , "V"
         intLastRow = intRow
         .col = intCol
         
         ClearData2
         TxtLocked False
         If "" & .TextMatrix(intRow, 0) = "V" And "" & .TextMatrix(intRow, colATR05) <> "" And colATR05 > 0 Then
             If ShowATRdetail("" & .TextMatrix(intRow, colATR05), "" & .TextMatrix(intRow, colATR06), "" & .TextMatrix(intRow, colATR07) _
                , "" & .TextMatrix(intRow, colATR08), "" & .TextMatrix(intRow, colATR09), "" & .TextMatrix(intRow, colATR10) _
                , "" & .TextMatrix(intRow, colSCP01), "" & .TextMatrix(intRow, colSCP02), "" & .TextMatrix(intRow, colSCP03), "" & .TextMatrix(intRow, colSCP04), "" & .TextMatrix(intRow, colPKno)) = False Then
                 MsgBox "目前無此筆記錄，請重新查詢！"
             Else
                TxtLocked False
             End If
         End If
       End If
   End With
End Sub

Private Sub Option1_Click(Index As Integer)
   If "" & m_ATR(12) = "" And Val(m_ATR(5)) = 0 Then
      If Index = 0 Then
         txtData(1).Locked = False
         Frame3.Enabled = False
         '(保留) 按鈕隱藏
         'CmdPS.Visible = False
         'For intI = 0 To 3
         '   txtSP(intI).Locked = True
         'Next
      Else
         txtData(1).Locked = True
         Frame3.Enabled = True
         '(保留) 按鈕隱藏
         'CmdPS.Visible = True
         'For intI = 0 To 3
         '   txtSP(intI).Locked = False
         'Next
      End If
   End If
End Sub

Private Sub txtCase_GotFocus(Index As Integer)
   TextInverse txtCase(Index)
End Sub

Private Sub txtCase_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCase_LostFocus(Index As Integer)
   If Index > 1 And Trim(txtCase(Index)) = "" Then
      If Index = 2 Then
           txtCase(2) = "0"
      ElseIf Index = 3 Then
           txtCase(3) = "00"
      End If
   End If
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
   TextInverse txtData(Index)
End Sub

Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then
      KeyAscii = UpperCase(KeyAscii)
   Else
      KeyAscii = Pub_NumAscii(KeyAscii, True)
   End If
End Sub

Private Function ShowATRdetail(ByVal pATR05 As String, ByVal pATR06 As String, ByVal pATR07 As String, pATR08 As String, pATR09 As String, pATR10 As String, _
        pCP01 As String, pCP02 As String, pCP03 As String, pCP04 As String, pPKno As String) As Boolean
   
   ShowATRdetail = False
   
   If pPKno = "" And pATR05 = "1" Then  '預設:請款階段1
      '如果智權人員離職改抓最新進度的智權人員
      strExc(0) = "select decode(st04,'1',cp13,substr(v05,9,6)) as cp13,staffname(decode(st04,'1',cp13,substr(v05,9,6))) as st02 " & _
                  "from caseprogress,staff,(select cp01 as v01,cp02 as v02,cp03 as v03,cp04 as v04, max(cp05||cp13) v05 " & _
                  "from caseprogress,staff where cp01='" & m_ATR(1) & "' and cp02='" & m_ATR(2) & "' and cp03='" & m_ATR(3) & "' and cp04='" & m_ATR(4) & "' and cp159=0 and cp09 <'C' and cp13=st01(+) and st04='1' group by cp01,cp02,cp03,cp04) " & _
                  "where cp09='" & pATR06 & "' and cp13=st01(+) and cp01=v01(+) and cp02=v02(+) and cp03=v03(+) and cp04=v04(+) "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_PKno = pPKno
         Option1(0).Value = True
         txtData(0) = "" & RsTemp.Fields("cp13")
         lblFM2(2) = "" & RsTemp.Fields("st02")
         Frame2.Enabled = False
      Else
         Exit Function
      End If
   ElseIf pATR05 <> "" And pATR06 <> "" And pATR07 <> "" Then
      strExc(0) = "select atr01,atr02,atr03,atr04,atr05,atr06,atr07,atr08,atr09,atr10,atr11,atr12,cp01,cp02,cp03,cp04,cp13,cp14,st02 " & _
                  "from acs_tips_rate , caseprogress,staff where atr01='" & m_ATR(1) & "' and atr02='" & m_ATR(2) & "' and atr03='" & m_ATR(3) & "' and atr04='" & m_ATR(4) & "' " & _
                  "and atr05='" & pATR05 & "'  and atr06='" & pATR06 & "'  and atr07='" & pATR07 & "' and atr06=cp09(+) and atr07=st01(+) "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         For intQ = 5 To UBound(m_ATR)
            m_ATR(intQ) = "" & RsTemp.Fields("ATR" & Format(intQ, "00"))
         Next intQ
         If Trim(m_ATR(8) & m_ATR(9) & m_ATR(10) & RsTemp.Fields("cp01") & RsTemp.Fields("cp02") & RsTemp.Fields("cp03") & RsTemp.Fields("cp04")) _
               <> Trim(pATR08 & pATR09 & pATR10 & pCP01 & pCP02 & pCP03 & pCP04) Then
               MsgBox "資料已有變更，請重新查詢！", vbCritical
               Exit Function
         End If
         '(保留)人員只是代表部門，離職後會有其他人員接手
         'If "" & m_ATR(12) = "" Then '排除已確認
         '   If Val(m_ATR(5)) < 100 Then
         '      If "" & RsTemp.Fields("atr07") <> "" & RsTemp.Fields("cp13") Then
         '         MsgBox "目前智權人員為" & RsTemp.Fields("cp13") & GetStaffName(RsTemp.Fields("cp13"), True) & "，" & vbCrLf & "與設定不同請洽協作單位！", vbInformation
         '         Exit Function
         '      End If
         '   Else
         '      If "" & RsTemp.Fields("atr07") <> "" & RsTemp.Fields("cp14") Then
         '         MsgBox "智財協作的承辦人為" & RsTemp.Fields("cp14") & GetStaffName(RsTemp.Fields("cp14"), True) & "，" & vbCrLf & "與設定不同請洽協作單位！", vbInformation
         '         Exit Function
         '      End If
         '   End If
         'End If
         'end ----(保留)
         m_PKno = pPKno
         txtData(0) = m_ATR(7)
         lblFM2(2) = "" & RsTemp.Fields("st02")
         If Val(m_ATR(5)) < 100 Then
            Option1(0).Value = True
            txtData(1) = m_ATR(8)
         Else
            Option1(1).Value = True
            txtData(2) = m_ATR(9)
            txtData(3) = m_ATR(10)
            txtData(4) = m_ATR(5)
            lblFM2(3) = m_ATR(6)
            lblFM2(3).Tag = lblFM2(3)
            txtSP(0) = "" & RsTemp.Fields("cp01")
            txtSP(1) = "" & RsTemp.Fields("cp02")
            txtSP(2) = "" & RsTemp.Fields("cp03")
            txtSP(3) = "" & RsTemp.Fields("cp04")
         End If
         For Each oObj In txtData
            oObj.Tag = oObj.Text
         Next
      End If
   End If
   
   ShowATRdetail = True
   Exit Function


End Function

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
Dim strTmp1 As String

   If Index >= 1 And Index <= 3 And Trim(txtData(Index)) = "" Then
      Exit Sub
   End If
   If CmdOK(0).Visible = False Or CmdOK(0).Enabled = False Then
      Exit Sub
   End If
   
   Select Case Index
      Case 0  '員工編號
         If txtData(Index) = "" Then
            MsgBox "員工編號不可空白！", vbCritical, "檢核資料"
            lblFM2(2) = ""
            GoTo EXITSUB
         Else
            If txtData(Index).Tag <> txtData(Index).Text Then
               If ClsPDGetStaff(txtData(Index), strTmp1) = True Then
                  lblFM2(2) = strTmp1
               Else
                  lblFM2(2) = ""
                  GoTo EXITSUB
               End If
            End If
         End If
         txtData(Index).Tag = txtData(Index).Text
      Case 1  '智權業績獎金比例
         If Option1(0).Value = True Then
            If Val(txtData(Index)) <> 35 And Val(txtData(Index)) <> 40 Then
               MsgBox "智權業績獎金比例請輸入35%或40%！", vbCritical, "檢核資料"
               GoTo EXITSUB
            End If
         End If
      Case 2, 3 '專業點數比例、專業業務獎金(顧服獎金)
         If Option1(1).Value = True Then
            If Val(txtData(Index)) > 20 Then
               MsgBox IIf(Index = 2, "專業點數比例", "專業業務獎金(顧服獎金)比例") & "不可超過20%！", vbCritical, "檢核資料"
               GoTo EXITSUB
            End If
            '保留
'            If Val(txtData(Index)) > 10 Then
'               If MsgBox(IIf(Index = 2, "專業點數比例", "專業業務獎金(顧服獎金)比例") & "超過10%，是否繼續輸入？", vbInformation + vbYesNo + vbDefaultButton2, "檢核資料") = vbNo Then
'                  GoTo EXITSUB
'               End If
'            End If
         End If
      Case 4  '請款年度
         If Option1(1).Value = True Then
            strTmp1 = Mid(strSrvDate(2), 1, 3)
            If Val(txtData(Index)) < Val(strTmp1) - 1 Or Val(txtData(Index)) > Val(strTmp1) + 1 Then
               MsgBox "請款年度請輸入" & Val(strTmp1) - 1 & "、" & Val(strTmp1) & "、" & Val(strTmp1) + 1, vbCritical, "檢核資料"
               GoTo EXITSUB
            End If
         End If
   End Select
   Exit Sub
   
EXITSUB:
   Cancel = True
   txtData(Index).SetFocus
   Txtdata_GotFocus Index
End Sub

Private Sub TxtSP_GotFocus(Index As Integer)
   TextInverse txtSP(Index)
End Sub

Private Sub txtSP_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub TxtSP_LostFocus(Index As Integer)
   If Index > 1 And Trim(txtSP(Index)) = "" Then
      If Index = 2 Then
           txtSP(2) = "0"
      ElseIf Index = 3 Then
           txtSP(3) = "00"
      End If
   End If
End Sub
