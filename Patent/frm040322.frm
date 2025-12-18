VERSION 5.00
Begin VB.Form frm040322 
   BorderStyle     =   1  '單線固定
   Caption         =   "其他通知函/聯絡單"
   ClientHeight    =   5820
   ClientLeft      =   1170
   ClientTop       =   3300
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6495
   Begin VB.ComboBox cboNP 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1035
      Style           =   2  '單純下拉式
      TabIndex        =   30
      Top             =   3570
      Width           =   1905
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   1176
      Left            =   1488
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   5
      Text            =   "frm040322.frx":0000
      Top             =   4200
      Width           =   4668
   End
   Begin VB.TextBox Text3 
      Height          =   264
      Left            =   1485
      MaxLength       =   1
      TabIndex        =   6
      Top             =   5415
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Left            =   1032
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1710
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   1
      Left            =   1032
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "P"
      Top             =   504
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   2
      Left            =   1512
      MaxLength       =   6
      TabIndex        =   1
      Top             =   504
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   3
      Left            =   2352
      MaxLength       =   1
      TabIndex        =   2
      Top             =   504
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   4
      Left            =   2592
      MaxLength       =   2
      TabIndex        =   3
      Top             =   504
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   5436
      TabIndex        =   8
      Top             =   20
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   350
      Index           =   0
      Left            =   4608
      TabIndex        =   7
      Top             =   20
      Width           =   800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(7:TW-SUPA期限通知)"
      Height          =   180
      Index           =   14
      Left            =   1515
      TabIndex        =   31
      Top             =   2760
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(5:未收文期限提醒函)"
      Height          =   180
      Index           =   13
      Left            =   1515
      TabIndex        =   29
      Top             =   2220
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "解除期限日："
      Height          =   270
      Index           =   12
      Left            =   90
      TabIndex        =   28
      Top             =   3900
      Width           =   1185
   End
   Begin VB.Label lblCP05 
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1275
      TabIndex        =   27
      Top             =   3900
      Width           =   1620
   End
   Begin VB.Label lblNP08 
      Height          =   270
      Left            =   3990
      TabIndex        =   24
      Top             =   3600
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "下一程序："
      Height          =   270
      Index           =   11
      Left            =   90
      TabIndex        =   26
      Top             =   3600
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "本所期限："
      Height          =   270
      Index           =   10
      Left            =   3015
      TabIndex        =   25
      Top             =   3600
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "聯絡單備註："
      Height          =   270
      Index           =   9
      Left            =   90
      TabIndex        =   23
      Top             =   4200
      Width           =   1185
   End
   Begin VB.Label lblSaleName 
      Height          =   270
      Left            =   4005
      TabIndex        =   22
      Top             =   3300
      Width           =   1620
   End
   Begin VB.Label lblSaleZone 
      Height          =   270
      Left            =   870
      TabIndex        =   21
      Top             =   3300
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   270
      Index           =   8
      Left            =   3015
      TabIndex        =   20
      Top             =   3300
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      Height          =   270
      Index           =   7
      Left            =   90
      TabIndex        =   19
      Top             =   3300
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "是否修改定稿：             (Y : Word)"
      Height          =   270
      Index           =   6
      Left            =   90
      TabIndex        =   18
      Top             =   5415
      Width           =   3450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(6:其他聯絡單)"
      Height          =   180
      Index           =   5
      Left            =   1515
      TabIndex        =   17
      Top             =   2490
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(3:大陸年費聯絡函 4:大陸領證/年費逾期函)"
      Height          =   180
      Index           =   4
      Left            =   1515
      TabIndex        =   16
      Top             =   1965
      Width           =   3360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(2:信件退回聯絡函)"
      Height          =   180
      Index           =   3
      Left            =   1515
      TabIndex        =   15
      Top             =   1710
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "列印種類："
      Height          =   270
      Index           =   2
      Left            =   90
      TabIndex        =   14
      Top             =   1710
      Width           =   1185
   End
   Begin VB.Label lblCaseName 
      Caption         =   "(日)："
      Height          =   270
      Index           =   2
      Left            =   885
      TabIndex        =   13
      Top             =   1395
      Width           =   5415
   End
   Begin VB.Label lblCaseName 
      Caption         =   "(英)："
      Height          =   270
      Index           =   1
      Left            =   885
      TabIndex        =   12
      Top             =   1110
      Width           =   5415
   End
   Begin VB.Label lblCaseName 
      Caption         =   "(中)："
      Height          =   270
      Index           =   0
      Left            =   885
      TabIndex        =   11
      Top             =   810
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱"
      Height          =   270
      Index           =   1
      Left            =   90
      TabIndex        =   10
      Top             =   810
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   276
      Index           =   0
      Left            =   96
      TabIndex        =   9
      Top             =   504
      Width           =   1188
   End
End
Attribute VB_Name = "frm040322"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
Dim intWhere As Integer, strReceiveNo As String
Const ET01 As String = "10"
Dim m_blnTxtValidate As Boolean
Dim m_strCP09 As String '總收文號
'Add By Cheng 2002/12/18
Dim m_strCopy As String '北所一頁一張, 分所一頁二張
'Add By Cheng 2003/04/03
Dim m_PrtOrientation As Integer '列印方向
Dim m_PrtScaleMode As Integer '列印座標單位
Dim m_dblTop As Double '上邊界
Dim m_dblLeft As Double '左邊界
Dim m_dblTitleHeight As Double '表頭高度
Dim m_dblLine As Double '行數
Dim m_dblLineHeight As Double '行高
Dim m_dblBetweenLine As Double '行間空隙
Dim m_dblLineHeight1 As Double '行高
Dim m_dblBetweenLine1 As Double '行間空隙
Dim m_strSQLA As String
Dim m_rsA As New ADODB.Recordset
Dim pa(4) As String 'Add by Morgan 2010/6/23
'Added by Morgan 2012/3/30
Dim m_UsPatent As String '美國發明案號
Dim m_RtnDate As String '客戶回覆期限
Dim m_TwAppNo As String '台灣案申請號

'Add by Morgan 2010/6/23
Private Sub cboNP_Click()
   If cboNP.Tag <> Trim(cboNP.ListIndex) Then
      intI = cboNP.ListIndex
      If intI >= 0 Then
         If cboNP.ItemData(intI) > 0 Then
            Me.lblNP08.Caption = ChangeTStringToTDateString(cboNP.ItemData(intI) - 19110000)
            Text4 = GetNote(Text2.Text)
         End If
      End If
      cboNP.Tag = intI
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
    Dim rsA  As New ADODB.Recordset
    Dim StrSQLa As String
    Dim strDept As String
    Dim strDeadLine As String '時間要求
    Dim ii As Integer
    Dim stET01 As String, stET02 As String, stET03 As String 'Added by Morgan 2012/3/29
    
   Select Case Index
      Case 0 '確定
         ' 設定滑鼠游標為等待狀態
         Screen.MousePointer = vbHourglass
         If Me.Text1(1).Text = "" Then
            MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
            Me.Text1(1).SetFocus
            TextInverse Me.Text1(1)
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         'If Me.Text1(2).Text = ""  Then
         'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
            Text1_Validate 4, False '增加判斷本所案號
         If Me.Text1(2).Text = "" Or m_blnTxtValidate = False Then
            MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
            Me.Text1(2).SetFocus
            TextInverse Me.Text1(2)
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         If Me.Text2.Text = "" Then
            MsgBox "請輸入列印種類!!!", vbExclamation + vbOKOnly
            Me.Text2.SetFocus
            TextInverse Me.Text2
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         'Added by Lydia 2015/04/20 P案每月5日TW-SUPA通知定稿改在"繳年費/實體審查通知函"
         If Me.Text2.Text = "7" And Text1(1).Text = "P" Then
            MsgBox "P案TW-SUPA通知定稿改在繳年費/實體審查通知函!!!", vbExclamation + vbOKOnly
            Me.Text2.SetFocus
            TextInverse Me.Text2
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         
        '判斷列印種類
         Select Case Me.Text2.Text
         
'Removed by Morgan 2016/6/16 配合電子化,移除不再使用的選項(將來若再使用要搭配C類來函)--玲玲
'         '1:遭異可領證通知函
'         Case "1"
'            '列印定稿
'            NowPrint Me.Text1(1).Text & Me.Text1(2).Text & IIf(Me.Text1(3).Text = "", "0", Me.Text1(3).Text) & IIf(Me.Text1(4).Text = "", "00", Me.Text1(4).Text) & "&000", "18", "01", IIf(Me.Text3.Text = "Y", True, False), strUserNum, 0
'end 2016/6/16

         'Memo by Morgan 2016/6/16 P案移到繳年費/實體審查通知,此處只會是CFP案
         Case "7" 'TW-SUPA期限通知 Added by Morgan 2012/3/29
            If pa(1) = "P" Then
               stET01 = "18"
               stET03 = "02"
               strExc(0) = "select pd01||'-'||pd02||decode(pd03||pd04,'000','','-'||pd03||'-'||pd04) UsNo,cp27,pd06 from patent p1,pridate,caseprogress c1" & _
                  " where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "' and pa08='1' and pa09='000' and pa16 is null and pd06(+)=pa11 and pd07(+)=pa09" & _
                  " and exists(select * from patent p2 where p2.pa01=pd01 and p2.pa02=pd02 and p2.pa03=pd03 and p2.pa04=pd04 and p2.pa09='101' and p2.pa08='1')" & _
                  " and cp01(+)=pd01 and cp02(+)=pd02 and cp03(+)=pd03 and cp04(+)=pd04 and cp10='101' and cp27>0 and cp57 is null" & _
                  " AND NOT EXISTS(SELECT * FROM CASEPROGRESS C2 WHERE C2.CP01=P1.PA01 AND C2.CP02=P1.PA02" & _
                  " AND C2.CP03=P1.PA03 AND C2.CP04=P1.PA04 AND C2.CP10='1202')"
            Else
               stET01 = "10"
               stET03 = "30"
               strExc(0) = "select pd01||'-'||pd02||decode(pd03||pd04,'000','','-'||pd03||'-'||pd04) UsNo,cp27,pd06 from patent p1,pridate,caseprogress c1" & _
                  " where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "' and pa08='1' and pa09='101'" & _
                  " and pd01(+)=pa01 and pd02(+)=pa02 and pd03(+)=pa03 and pd04(+)=pa04 and pd07='000'" & _
                  " and cp01(+)=pd01 and cp02(+)=pd02 and cp03(+)=pd03 and cp04(+)=pd04 and cp10='101' and cp27>0 and cp57 is null"
            End If
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m_UsPatent = RsTemp("UsNo")
               m_RtnDate = CompDate(1, 5, RsTemp("cp27"))
               m_TwAppNo = "" & RsTemp("pd06")
               stET02 = pa(1) & pa(2) & pa(3) & pa(4) & "&000"
               StartLetter stET01, stET02, stET03
               NowPrint stET02, stET01, stET03, IIf(Me.Text3.Text = "Y", True, False), strUserNum, 0
            Else
              MsgBox "資料不符!!", vbExclamation
            End If
            
'Removed by Morgan 2016/6/16 配合電子化,移除不再使用的選項(將來若再使用要搭配C類來函)--玲玲
'         'Added by Morgan 2012/7/24
'         '8:變更代理人核准通知函
'         Case "8"
'            NowPrint pa(1) & pa(2) & pa(3) & pa(4) & "&000", "05", "23", IIf(Me.Text3.Text = "Y", True, False), strUserNum, 0
'end 2016/6/16

         Case Else
            StrSQLa = "Select A0902 From Staff,ACC090 WHERE ST03=A0901 AND ST01='" & strUserNum & "'"
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               strDept = "" & rsA.Fields(0).Value
            Else
               strDept = ""
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '時間期限為發文日 + 7 天
            If Text2 <> "5" Then
               strDeadLine = ChangeWDateStringToWString(DateAdd("d", 7, ChangeWStringToWDateString(strSrvDate(1))))
            End If
            'Modify by Morgan 2010/6/24 改抓基本檔就好
            'm_strSQLA = "select A0902 a01,st02 a02," & ChgCaseprogress("", 1) & " a03,pa05 a04,pa06 a05," & _
                "pa07 a06,cu04 a07,FA05||FA63 a08," & SQLDate("PA14", True) & " a09,TPB08 a10 " & _
                "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,CUSTOMER,FAGENT,TPBULLETIN,ACC090 WHERE ST03=A0901 AND cp09='" & m_strCP09 & "' AND " & _
                "CP01=PA01 and CP02=PA02 and CP03=PA03 and CP04=PA04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) and " & _
                "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND " & _
                "PA11=TPB01(+)"
            m_strSQLA = "select '" & lblSaleZone & "' a01,'" & lblSaleName & "' a02" & _
                "," & ChgPatent("", 1) & " a03,pa05 a04,pa06 a05," & _
                "pa07 a06,cu04 a07,FA05||FA63 a08," & SQLDate("PA14", True) & " a09,TPB08 a10 " & _
                "FROM PATENT,CUSTOMER,FAGENT,TPBULLETIN" & _
                " WHERE pa01='" & pa(1) & "' and PA02='" & pa(2) & "' and PA03='" & pa(3) & "' and PA04='" & pa(4) & "'" & _
                " and SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+)" & _
                " and SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)" & _
                " AND PA11=TPB01(+)"
            'end 2010/6/23
            m_rsA.CursorLocation = adUseClient
            m_rsA.Open m_strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
            If m_rsA.RecordCount > 0 Then
                '取得預設印表機設定值
                m_PrtOrientation = Printer.Orientation
                m_PrtScaleMode = Printer.ScaleMode
                '重新設定印表機
                Printer.PaperSize = vbPRPSA4
                Printer.Orientation = vbPRORPortrait
                Printer.ScaleMode = vbCentimeters
                '列印聯絡單
                InitPrtPosition 0.5, 0.5
                PrintContactSheet strDept, strDeadLine
                '若受文者非北所
                If m_strCopy <> "1" Then
                    '列印聯絡單
                    InitPrtPosition 13.5, 0.5
                    PrintContactSheet strDept, strDeadLine
                End If
                Printer.EndDoc
                '還原預設印表機設值
                Printer.Orientation = m_PrtOrientation
                Printer.ScaleMode = m_PrtScaleMode
            End If
            If m_rsA.State <> adStateClosed Then m_rsA.Close
            Set m_rsA = Nothing
         End Select
        ClearForm
         ' 設定滑鼠游標為預設
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   Me.Text4.Text = ""
   cboNP.Clear 'Add by Morgan 2010/6/23
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040322 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   Me.cmdOK(0).Enabled = False
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Select Case Index
Case 4
   If m_blnTxtValidate = False Then
      Me.Text1(1).SetFocus
      m_blnTxtValidate = True
   End If
End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim stCP13 As String
   
Select Case Index
Case 1 '系統類別
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   StrSQLa = "Select * From SystemKind Where SK01='" & Me.Text1(1).Text & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount <= 0 Then
      MsgBox "系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
      Cancel = True
      Me.Text1(1).SetFocus
      TextInverse Me.Text1(1)
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing

Case 4
   m_blnTxtValidate = True
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/30 清除查詢印表記錄檔欄位
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   pa(1) = Me.Text1(1).Text
   pa(2) = Me.Text1(2).Text
   pa(3) = IIf(Me.Text1(3).Text = "", "0", Me.Text1(3).Text)
   pa(4) = IIf(Me.Text1(4).Text = "", "00", Me.Text1(4).Text)
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    If FMP2open = True Then
      If PUB_FMPtoCheck(0, 1, Pub_strUserST05, pa(1), pa(2), pa(3), pa(4)) = False Then
        ClearForm
        m_blnTxtValidate = False
        Me.Text1(1).Text = pa(1)
        Me.Text1(2).Text = pa(2)
        Me.Text1(3).Text = pa(3)
        Me.Text1(4).Text = pa(4)
        Exit Sub
      End If
    End If
    
   pub_QL05 = pub_QL05 & ";" & Label1(0) & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) 'Add By Sindy 2010/11/30
   StrSQLa = "Select * From PATENT Where PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount <= 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/11/30
      MsgBox "資料庫無此案號資料!!!", vbExclamation + vbOKOnly
      Me.lblCaseName(0).Caption = "(中)："
      Me.lblCaseName(1).Caption = "(英)："
      Me.lblCaseName(2).Caption = "(日)："
      Me.lblSaleZone.Caption = ""
      Me.lblSaleName.Caption = ""
      m_blnTxtValidate = False
   Else
      InsertQueryLog (rsA.RecordCount) 'Add By Sindy 2010/11/30
      Me.lblCaseName(0).Caption = "(中)：" & rsA("PA05").Value
      Me.lblCaseName(1).Caption = "(英)：" & rsA("PA06").Value
      Me.lblCaseName(2).Caption = "(日)：" & rsA("PA07").Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   'Add by Morgan 2010/6/23
   Me.lblSaleZone.Caption = ""
   Me.lblSaleName.Caption = ""
   m_strCopy = ""
   stCP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
   If stCP13 <> "" Then
      strExc(0) = "select A0902,ST02,ST06 from staff,acc090 where st01='" & stCP13 & "' and a0901(+)=st15"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Me.lblSaleZone.Caption = "" & RsTemp("A0902")
         Me.lblSaleName.Caption = "" & RsTemp("ST02")
         m_strCopy = "" & RsTemp("ST06")
      End If
   End If
   'end 2010/6/23
End Select
If Cancel Then TextInverse Text1(Index)
End Sub

Private Sub Text2_Change()
'Add by Morgan 2010/6/23
If Text2.Text = "5" Then
   cboNP.Enabled = True
Else
   cboNP.Clear
   cboNP.Tag = ""
   cboNP.Enabled = False
End If
'end 2010/6/23
'Modified by Morgan 2012/3/29 +7
'If Me.Text2.Text = "1" Then
'Modified by Morgan 2012/7/24 +8
If Me.Text2.Text = "1" Or Me.Text2.Text = "7" Or Me.Text2.Text = "8" Then
   Me.Text4.Enabled = False
   Me.Text3.Enabled = True
Else
   Me.Text4.Enabled = True
   Me.Text3.Text = ""
   Me.Text3.Enabled = False
End If
End Sub

Private Sub Text2_GotFocus()
TextInverse Me.Text2
Me.cmdOK(0).Enabled = False
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
'Modified by Morgan 2012/3/29 + 7
'Modified by Morgan 2012/7/24 +8
'Modified by Morgan 2016/6/16 取消1,8
If (KeyAscii < Asc("2") Or KeyAscii > Asc("7")) And KeyAscii <> 8 Then
   KeyAscii = 0
End If
End Sub

Private Sub Text2_LostFocus()
If m_blnTxtValidate = False Then
   Me.Text1(1).SetFocus
   m_blnTxtValidate = True
End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim bolTmp As Boolean
Dim strTempName As String

If Me.Text2.Text = "" Then Exit Sub
m_blnTxtValidate = True
ClearData
m_strCP09 = ""

'若列印種類為"3"或"4", 檢查是否為大陸案
'Modify by Morgan 2010/6/23 +5 要判斷非台灣案
If Me.Text2.Text = "3" Or Me.Text2.Text = "4" Or Me.Text2.Text = "5" Then
   If Text2.Text = "5" Then
      strExc(1) = " And PA09<>'" & 台灣國家代號 & "'"
      strExc(2) = "申請國家不可為台灣!!!"
   Else
      strExc(1) = " And PA09='" & 大陸國家代號 & "'"
      strExc(2) = "申請國家必須為大陸!!!"
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   StrSQLa = "Select * From PATENT Where PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "'" & strExc(1)
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount <= 0 Then
      MsgBox strExc(2), vbExclamation + vbOKOnly
      m_blnTxtValidate = False
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      Exit Sub
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End If

'檢查案件進度檔的資料
'Remove by Morgan 2010/6/23 改案號跳離時用新的規則抓
'StrSQLa = "Select A0902,ST02,CP09,ST06 From CaseProgress,Acc090,Staff Where CP12=A0901(+) AND CP13=ST01(+) AND CP09=(SELECT MAX(CP09) FROM CASEPROGRESS WHERE CP01='" & Me.text1(1).Text & "' AND CP02='" & Me.text1(2).Text & "' AND CP03='" & IIf(Me.text1(3).Text = "", "0", Me.text1(3).Text) & "' AND CP04='" & IIf(Me.text1(4).Text = "", "00", Me.text1(4).Text) & "'" & " And CP09 <'B')"
'rsA.CursorLocation = adUseClient
'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount <= 0 Then
'   MsgBox "資料庫無案件進度資料!!!", vbExclamation + vbOKOnly
'   m_blnTxtValidate = False
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'   Exit Sub
'Else
'   Me.lblSaleZone.Caption = "" & rsA.Fields(0).Value
'   Me.lblSaleName.Caption = "" & rsA.Fields(1).Value
'   m_strCP09 = "" & rsA.Fields(2).Value
'   m_strCopy = "" & rsA.Fields(3).Value
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'end 2010/6/23

 
'若列印種類為"3", 檢查是否有下一程序
If Me.Text2.Text = "3" Then
   StrSQLa = "Select NP07,NP08 From NextProgress Where NP02='" & Me.Text1(1).Text & "' AND NP03='" & Me.Text1(2).Text & "' AND NP04='" & IIf(Me.Text1(3).Text = "", "0", Me.Text1(3).Text) & "' AND NP05='" & IIf(Me.Text1(4).Text = "", "00", Me.Text1(4).Text) & "'" & " And NP07='605' AND NP06 IS NULL "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount <= 0 Then
      MsgBox "無年費期限資料!!!", vbExclamation + vbOKOnly
      m_blnTxtValidate = False
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      Exit Sub
   Else
      '大陸申請案
      bolTmp = True
      If ClsPDGetCaseProperty(Me.Text1(1).Text, rsA.Fields(0).Value, strTempName, bolTmp) Then
         'Modify by Morgan 2010/6/23
         'Me.lblNP07.Caption = "" & strTempName
         Me.cboNP.AddItem strTempName
         cboNP.ListIndex = 0
      End If
      Me.lblNP08.Caption = "" & IIf("" & rsA.Fields(1).Value <> "", ChangeTStringToTDateString(rsA.Fields(1).Value - 19110000), "")
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
'若列印種類為"4", 檢查是否有下一程序
ElseIf Me.Text2.Text = "4" Then
   StrSQLa = "Select NP07,NP08 From NextProgress Where NP02='" & Me.Text1(1).Text & "' AND NP03='" & Me.Text1(2).Text & "' AND NP04='" & IIf(Me.Text1(3).Text = "", "0", Me.Text1(3).Text) & "' AND NP05='" & IIf(Me.Text1(4).Text = "", "00", Me.Text1(4).Text) & "'" & " And (NP07='601' OR NP07='605') AND NP06 IS NULL AND NP08<" & strSrvDate(1) & " Order By NP08"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount <= 0 Then
      MsgBox "無領證或年費逾期的期限資料!!!", vbExclamation + vbOKOnly
      m_blnTxtValidate = False
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      Exit Sub
   Else
      '大陸申請案
      bolTmp = True
      If ClsPDGetCaseProperty(Me.Text1(1).Text, rsA.Fields(0).Value, strTempName, bolTmp) Then
         'Modify by Morgan 2010/6/23
         'Me.lblNP07.Caption = "" & strTempName
         Me.cboNP.AddItem strTempName
         cboNP.ListIndex = 0
      End If
      Me.lblNP08.Caption = "" & IIf("" & rsA.Fields(1).Value <> "", ChangeTStringToTDateString(rsA.Fields(1).Value - 19110000), "")
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
'Add by Morgan 2010/6/23
ElseIf Me.Text2.Text = "5" Then
   strExc(0) = "Select distinct NP07,NP08,cpm04 From NextProgress,casepropertymap Where NP02='" & Me.Text1(1).Text & "' AND NP03='" & Me.Text1(2).Text & "' AND NP04='" & IIf(Me.Text1(3).Text = "", "0", Me.Text1(3).Text) & "' AND NP05='" & IIf(Me.Text1(4).Text = "", "00", Me.Text1(4).Text) & "'" & _
      " AND NP06 IS NULL" & strNpSqlOfNoSalesDuty & " and cpm01(+)=np02 and cpm02(+)=np07 Order By NP08 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Do While Not RsTemp.EOF
         cboNP.AddItem "" & RsTemp("cpm04"), 0
         cboNP.ItemData(0) = RsTemp("np08")
         RsTemp.MoveNext
      Loop
      cboNP.ListIndex = 0
      cboNP_Click
   End If
End If
StrSQLa = "Select CP05 From CaseProgress Where " & ChgCaseprogress(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4).Text) & " And CP10='907' Order By CP06 Desc"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    Me.lblCP05.Caption = ChangeTStringToTDateString(ChangeWStringToTString("" & rsA.Fields(0).Value))
Else
    Me.lblCP05.Caption = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

'Modified by Morgan 2012/7/24 +7,8
If Me.Text2.Text = "1" Or Me.Text2.Text = "7" Or Me.Text2.Text = "8" Then
   Me.Text4.Text = ""
   Me.Text4.Enabled = False
Else
   Me.Text4.Text = ""
   Me.Text4.Enabled = True
   Select Case Me.Text2.Text
   'Modify by Morgan 2010/6/23 +5
   Case 2, 3, 4, 5
      Me.Text4.Text = GetNote(Text2.Text)
   Case Else
      Me.Text4.Text = GetNote()
   End Select
End If
Me.cmdOK(0).Enabled = True
 
End Sub

Private Sub Text3_GotFocus()
TextInverse Me.Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
If KeyAscii <> 89 And KeyAscii <> 8 Then
   KeyAscii = 0
End If
End Sub


'預設聯絡單備註
Private Function GetNote(Optional strKind As String) As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim rsB As New ADODB.Recordset
Dim StrSqlB As String
Dim strPA72 As String
Dim ii As Integer
Dim ArrYear
Dim str0 As String
Dim Str1 As String

str0 = "　"
Str1 = "　"

Select Case strKind
Case "2"
   'Modify by Morgan 2008/9/9 改通用--玲玲
   'GetNote = "　　上述之申請人地址有誤，以致於繳年費通知函遭退件，煩請寫異動表更正之。"
   GetNote = "　　上述之申請人地址有誤，以致於通知函遭退件，煩請寫異動表更正之。"
Case "3"
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   StrSQLa = "Select PA72,PA08,PA09 From PATENT Where PA01='" & Me.Text1(1).Text & "' AND PA02='" & Me.Text1(2).Text & "' AND PA03='" & IIf(Me.Text1(3).Text = "", "0", Me.Text1(3).Text) & "' AND PA04='" & IIf(Me.Text1(4).Text = "", "00", Me.Text1(4).Text) & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      If IsNull(rsA.Fields(0).Value) = False Then
         strPA72 = 0
         ArrYear = Split(rsA.Fields("PA72").Value, ",")
         For ii = LBound(ArrYear) To UBound(ArrYear)
            If Val(ArrYear(ii)) > Val(strPA72) Then
               strPA72 = ArrYear(ii)
            End If
         Next ii
         StrSqlB = "Select NA21, NA23, NA25 FROM NATION WHERE NA01='" & rsA.Fields("PA09").Value & "'"
         rsB.CursorLocation = adUseClient
         rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
         If rsB.RecordCount > 0 Then
            '發明
            If rsA.Fields("PA08").Value = "1" Then
               If IsNull(rsB.Fields("NA21").Value) Then GoTo DoNothing
               ArrYear = Split(rsB.Fields("NA21").Value, ",")
            '新型
            ElseIf rsA.Fields("PA08").Value = "2" Then
               If IsNull(rsB.Fields("NA23").Value) Then GoTo DoNothing
               ArrYear = Split(rsB.Fields("NA23").Value, ",")
            '設計
            Else
               If IsNull(rsB.Fields("NA25").Value) Then GoTo DoNothing
               ArrYear = Split(rsB.Fields("NA25").Value, ",")
            End If
            For ii = LBound(ArrYear) To UBound(ArrYear)
               If Val(ArrYear(ii)) > Val(strPA72) Then
                  str0 = ArrYear(ii)
                  Exit For
               End If
            Next ii
            
         End If
DoNothing:
         If rsB.State <> adStateClosed Then rsB.Close
         Set rsB = Nothing
      End If
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   'Modified by Morgan 2012/1/12 遇 101/02/29 會錯因為是用西元格式檢查,改判斷有沒有值就好
   'If IsDate("" & Me.lblNP08.Caption) Then
   If Me.lblNP08.Caption <> "" Then
      Str1 = DBYEAR(Me.lblNP08.Caption) - 1911 & "年" & DBMONTH(Me.lblNP08.Caption) & "月" & DBDAY(Me.lblNP08.Caption) & "日"
   End If
   '2008/11/28 MODIFY BY SONIA 分所取消卷宗內容
   'GetNote = "　　上述大陸專利案，之前曾通知你及客戶第" & str0 & "年年費繳費期限至" & Str1 & "止，並把卷退給你聯絡客戶，今代理人又來函通知上述事情，故再次提醒你該案之繳費期限將屆。另請你將此通知函放入該案卷內。"
   GetNote = "　　上述大陸專利案，之前曾通知你及客戶第" & str0 & "年年費繳費期限至" & Str1 & "止"
   If m_strCopy = "1" Then
      GetNote = GetNote & "，並把卷退給你聯絡客戶。"
   Else
      GetNote = GetNote & "。"
   End If
   GetNote = GetNote & "今代理人又來函通知上述事情，故再次提醒你該案之繳費期限將屆。"
   If m_strCopy = "1" Then GetNote = GetNote & "另請你將此通知函放入該案卷內。"
   '2008/11/28 END
Case "4"
   'Modified by Morgan 2012/1/12 遇 101/02/29 會錯因為是用西元格式檢查,改判斷有沒有值就好
   'If IsDate("" & Me.lblNP08.Caption) Then
   If Me.lblNP08.Caption <> "" Then
      str0 = DBYEAR(Me.lblNP08.Caption) - 1911 & "年" & DBMONTH(Me.lblNP08.Caption) & "月" & DBDAY(Me.lblNP08.Caption) & "日"
   End If
   If cboNP <> "年費" Then
      'Modify by Morgan 2010/6/23
      'Str1 = Me.lblNP07.Caption
      Str1 = cboNP
   Else
      StrSQLa = "Select PA72,PA08,PA09 From PATENT Where PA01='" & Me.Text1(1).Text & "' AND PA02='" & Me.Text1(2).Text & "' AND PA03='" & IIf(Me.Text1(3).Text = "", "0", Me.Text1(3).Text) & "' AND PA04='" & IIf(Me.Text1(4).Text = "", "00", Me.Text1(4).Text) & "'"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         If IsNull(rsA.Fields(0).Value) = False Then
            strPA72 = 0
            ArrYear = Split(rsA.Fields("PA72").Value, ",")
            For ii = LBound(ArrYear) To UBound(ArrYear)
               If Val(ArrYear(ii)) > Val(strPA72) Then
                  strPA72 = ArrYear(ii)
               End If
            Next ii
            StrSqlB = "Select NA21, NA23, NA25 FROM NATION WHERE NA01='" & rsA.Fields("PA09").Value & "'"
            rsB.CursorLocation = adUseClient
            rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
            If rsB.RecordCount > 0 Then
               '發明
               If rsA.Fields("PA08").Value = "1" Then
                  If IsNull(rsB.Fields("NA21").Value) Then GoTo DoNothing1
                  ArrYear = Split(rsB.Fields("NA21").Value, ",")
               '新型
               ElseIf rsA.Fields("PA08").Value = "2" Then
                  If IsNull(rsB.Fields("NA23").Value) Then GoTo DoNothing1
                  ArrYear = Split(rsB.Fields("NA23").Value, ",")
               '設計
               Else
                  If IsNull(rsB.Fields("NA25").Value) Then GoTo DoNothing1
                  ArrYear = Split(rsB.Fields("NA25").Value, ",")
               End If
               For ii = LBound(ArrYear) To UBound(ArrYear)
                  If Val(ArrYear(ii)) > Val(strPA72) Then
                     Str1 = ArrYear(ii)
                     Exit For
                  End If
               Next ii
               
            End If
DoNothing1:
            If rsB.State <> adStateClosed Then rsB.Close
            Set rsB = Nothing
         End If
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      'Modify by Morgan 2010/6/23
      'Str1 = "第" & Str1 & "年" & Me.lblNP07.Caption
      Str1 = "第" & Str1 & "年" & cboNP
   End If
   'Modify by Morgan 2010/6/23
   'GetNote = "　　上述大陸專利案，因一直未接獲結案單，至今已逾" & Me.lblNP07.Caption & "繳費期限(" & str0 & ")，而本所與大陸代理人間協議是在期限內未接獲本所不委辦的通知即代為提出申請，故本案之" & Str1 & "已產生，且代理人也將帳單寄至本所請款，故請你聯絡客戶並補收文，如欲結案請於結案單上註明費用如何處理。"
   GetNote = "　　上述大陸專利案，因一直未接獲結案單，至今已逾" & cboNP & "繳費期限(" & str0 & ")，而本所與大陸代理人間協議是在期限內未接獲本所不委辦的通知即代為提出申請，故本案之" & Str1 & "已產生，且代理人也將帳單寄至本所請款，故請你聯絡客戶並補收文，如欲結案請於結案單上註明費用如何處理。"
'Add by Morgan 2010/6/23
Case "5"
   'Modified by Morgan 2012/1/12 101/02/29 會錯因為是用西元格式檢查,改判斷有沒有值就好
   'If IsDate("" & Me.lblNP08.Caption) Then
   If Me.lblNP08.Caption <> "" Then
      str0 = DBYEAR(Me.lblNP08.Caption) - 1911 & "年" & DBMONTH(Me.lblNP08.Caption) & "月" & DBDAY(Me.lblNP08.Caption) & "日"
   End If
   GetNote = "　　上述專利案，之前曾通知你及客戶" & cboNP & "期限(" & str0 & ")，今代理人又來函通知前述事情，故再次提醒。"
   '北所
   If m_strCopy = "1" Then
      GetNote = GetNote & "請儘速收文，並將此函存卷。"
   End If
   
Case Else
   '無預設
End Select
End Function

'Add By Cheng 2003/04/03
Private Sub InitPrtPosition(dblTop As Double, dblLeft As Double)
    m_dblTop = dblTop
    m_dblLeft = dblLeft
    m_dblTitleHeight = 0
    m_dblLine = 0
    m_dblLineHeight = 1
    m_dblBetweenLine = 0.2
    m_dblLineHeight1 = 0.6
    m_dblBetweenLine1 = 0.1
End Sub

'Add By Cheng 2003/04/03
Private Sub PrintContactSheet(strDept As String, strDeadLine As String)
Dim dblPrtX As Double
Dim dblPrtY As Double
Dim ii As Integer
Dim jj As Integer
Dim strTxt  As String
Dim intTxtLeng As Integer
    
    Printer.Font.Name = "標楷體"
    Printer.Font.Size = 16
    'Removed by Morgan 2020/3/30
    'dblPrtX = m_dblLeft + (19 - Printer.TextWidth("台一國際專利商標事務所")) / 2
    'dblPrtY = m_dblTop + m_dblBetweenLine + 0
    'Printer.CurrentX = dblPrtX
    'Printer.CurrentY = dblPrtY
    'Printer.Print "台一國際專利商標事務所"
    'end 2020/3/30
    dblPrtX = m_dblLeft + (19 - Printer.TextWidth("簡易聯絡單")) / 2
    dblPrtY = m_dblTop + m_dblBetweenLine + 1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "簡易聯絡單"
        
    m_dblTitleHeight = 2.2
    
    m_dblLine = 0
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 19, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight + 9 * m_dblLineHeight)
    Printer.Line (m_dblLeft + 4.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 4.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight + 3 * m_dblLineHeight)
    Printer.Line (m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight + 9 * m_dblLineHeight)
    Printer.Line (m_dblLeft + 19, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 19, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight + 9 * m_dblLineHeight)
    Printer.Font.Size = 14
    dblPrtX = m_dblLeft + (4.5 - Printer.TextWidth("受文者")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "受文者"
    dblPrtX = m_dblLeft + 4.5 + (4 - Printer.TextWidth("發文者")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "發文者"
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    '受文者部門
    dblPrtX = m_dblLeft + m_dblBetweenLine
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "" & m_rsA.Fields(0).Value
    '受文者
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + (4.5 - Printer.TextWidth("" & m_rsA.Fields(1).Value)) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight - 0.3 * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "" & m_rsA.Fields(1).Value
    '發文者
    dblPrtX = m_dblLeft + 4.5 + (4 - Printer.TextWidth(GetStaffName(strUserNum, True))) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight - (m_dblLineHeight / 2)
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print GetStaffName(strUserNum, True)
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    Printer.Line (m_dblLeft + 2.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 2.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight + 6 * m_dblLineHeight)
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("發文時間")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "發文時間"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth(Mid(strSrvDate(1), 1, 4) - 1911 & "年" & Mid(strSrvDate(1), 5, 2) & "月" & Mid(strSrvDate(1), 7, 2) & "日")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print Mid(strSrvDate(1), 1, 4) - 1911 & "年" & Mid(strSrvDate(1), 5, 2) & "月" & Mid(strSrvDate(1), 7, 2) & "日"
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("答覆")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "答覆"
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("□否 □要")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "□否 □要"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth("用□電話 □口頭  回覆")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight - (m_dblLineHeight / 2)
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "用□電話 □口頭  回覆"
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("　限　　")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "　限　　"
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("時　要求")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight + 0.5 * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "時　要求"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth(Mid(Val(strDeadLine), 1, 4) - 1911 & "年" & Mid(Val(strDeadLine), 5, 2) & "月" & Mid(Val(strDeadLine), 7, 2) & "日")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight + 0.5 * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    '2008/11/28 MODIFY BY SONIA
    'Printer.Print Mid(Val(strDeadLine), 1, 4) - 1911 & "年" & Mid(Val(strDeadLine), 5, 2) & "月" & Mid(Val(strDeadLine), 7, 2) & "日"
    If Text2 = "4" Then
       Printer.Print Mid(Val(strDeadLine), 1, 4) - 1911 & "年" & Mid(Val(strDeadLine), 5, 2) & "月" & Mid(Val(strDeadLine), 7, 2) & "日"
    End If
    '2008/11/28 END
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("　間　　")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "　間　　"
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("發文地點")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "發文地點"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth(strDept)) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strDept
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 19, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)

    Printer.Font.Size = 13
    m_dblLine = 0
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "本所案號：" & m_rsA.Fields(2).Value
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "案件名稱(中)：" & m_rsA.Fields(3).Value
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    'Modify By Sindy 2009/09/04
    'Printer.Print "案件名稱(英)：" & m_rsA.Fields(4).Value
    If Len(Trim(m_rsA.Fields(4).Value)) > 26 Then
      Printer.Print "案件名稱(英)：" & Left(Trim(m_rsA.Fields(4).Value), 26) & "..."
    Else
      Printer.Print "案件名稱(英)：" & Trim(m_rsA.Fields(4).Value)
    End If
    '2009/09/04 End
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "案件名稱(日)：" & m_rsA.Fields(5).Value
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "申請人：" & m_rsA.Fields(6).Value
    m_dblLine = m_dblLine + 1
    If Me.Text4.Text <> "" Then
        strTxt = ""
        intTxtLeng = 0
        For ii = 1 To Len(Me.Text4.Text)
            If Asc(Mid(Me.Text4.Text, ii, 1)) >= 0 And Asc(Mid(Me.Text4.Text, ii, 1)) < 128 Then
                intTxtLeng = intTxtLeng + 1
            Else
                intTxtLeng = intTxtLeng + 2
            End If
            strTxt = strTxt & Mid(Me.Text4.Text, ii, 1)
            If intTxtLeng >= 39 Then
                m_dblLine = m_dblLine + 1
                dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
                dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
                Printer.CurrentX = dblPrtX
                Printer.CurrentY = dblPrtY
                Printer.Print strTxt
                strTxt = ""
                intTxtLeng = 0
            End If
        Next ii
        m_dblLine = m_dblLine + 1
        dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
        dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
        Printer.CurrentX = dblPrtX
        Printer.CurrentY = dblPrtY
        Printer.Print strTxt
        strTxt = ""
        intTxtLeng = 0
    End If
End Sub

'Add By Cheng 2003/04/09
'清除畫面欄位資料
Private Sub ClearForm()
    Me.Text1(2).Text = ""
    Me.Text1(3).Text = ""
    Me.Text1(4).Text = ""
    Me.Text2.Text = ""
    Me.Text3.Text = ""
    Me.lblCaseName(0).Caption = "(中)："
    Me.lblCaseName(1).Caption = "(英)："
    Me.lblCaseName(2).Caption = "(日)："
    Me.lblSaleName.Caption = ""
    Me.lblSaleZone.Caption = ""
    ClearData
    Me.Text1(2).SetFocus
End Sub

Private Sub ClearData()
   Me.Text4.Text = ""
   'Modify by Morgan 2010/6/23
    'Me.lblNP07.Caption = ""
    Me.cboNP.Clear
    Me.cboNP.Tag = ""
    Me.lblNP08.Caption = ""
    Me.lblCP05.Caption = ""
End Sub

'Added by Morgan 2012/3/29
Private Sub StartLetter(ByVal ET01 As String, ET02 As String, ByVal ET03 As String)
   Dim strTxt() As String, ii As Integer
    
   EndLetter ET01, ET02, ET03, strUserNum
   ii = 0
   
   ii = ii + 1
   ReDim Preserve strTxt(ii)
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','美國發明案','" & m_UsPatent & "')"
      
   ii = ii + 1
   ReDim Preserve strTxt(ii)
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','回覆期限','" & m_RtnDate & "')"
      
   ii = ii + 1
   ReDim Preserve strTxt(ii)
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','台灣案申請號','" & m_TwAppNo & "')"
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

