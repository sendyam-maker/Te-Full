VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm110104_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "更換FC代理人作業－多筆"
   ClientHeight    =   5748
   ClientLeft      =   216
   ClientTop       =   720
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9312
   Begin VB.Frame Frame1 
      Caption         =   "設定清單"
      Height          =   570
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5085
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   1
         Top             =   210
         Width           =   4110
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   8
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "案件筆數：共0筆!!!  已選取0筆 !"
      Height          =   4845
      Left            =   60
      TabIndex        =   5
      Top             =   780
      Width           =   9225
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4455
         Left            =   60
         TabIndex        =   0
         Top             =   240
         Width           =   9075
         _ExtentX        =   16002
         _ExtentY        =   7853
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   16772048
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         MergeCells      =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
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
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "全部選取(&A)"
      Height          =   400
      Left            =   5460
      TabIndex        =   2
      Top             =   90
      Width           =   1125
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8160
      TabIndex        =   4
      Top             =   90
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "更換FC代理人(&P)"
      Height          =   400
      Left            =   6615
      TabIndex        =   3
      Top             =   90
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PS：存檔後會清除個案的各項請款折扣數！"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   62
      Left            =   5520
      TabIndex        =   6
      Top             =   600
      Width           =   3420
   End
End
Attribute VB_Name = "frm110104_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/17 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB
'2011/5/30 CREATE BY SONIA
Option Explicit
 
Dim strPrinter As String
Dim m_blnColOrderAsc1 As Boolean '欄位資料由小到大排序
Dim lngSelRow As Long '記錄選取筆數
Dim PLeft(1 To 10) As Integer, Pleft2(1 To 4) As Integer
Dim strTemp(1 To 10) As String
Dim iLine As Integer
Dim strCaseNo As String
Dim strJumpList As String 'Added by Lydia 2022/07/04 FCP和FMP案之一案兩請僅其中一案更代時，彈提醒選擇發email通知的案號

Private Sub cmdBack_Click()
   frm110104_1.Show
   Unload Me
End Sub

Private Sub cmdok_Click()
Dim i As Long, m_i As Integer, intRow As Integer
Dim intQ As Integer
Dim intRe As Integer 'Added by Lydia 2022/07/04
Dim StrCaseList As String 'Added by Lydia 2019/12/19 更換FC代理人的本所案號
    
   strJumpList = "" 'Added by Lydia 2022/07/04
   
   If lngSelRow = 0 Then
      MsgBox "尚未選取欲更換FC代理人的案件!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   '先將選取資料列印清單,再更新資料
   PUB_RestorePrinter Combo1
RePrint:
   intRe = intRe + 1 'Added by Lydia 2022/07/04 記錄清單列印次數
   Printer.Orientation = 2 '1.直印 2.橫印
   iLine = 1: intRow = 0
   strCaseNo = ""
   StrCaseList = "" 'Added by Lydia 2019/12/19
   For i = 1 To MSHFlexGrid1.Rows - 1
      If MSHFlexGrid1.TextMatrix(i, 0) = "V" Then
         'intRow = intRow + 1 'Mark by Lydia 2022/07/04 改在判斷後才合計筆數
         For m_i = 1 To 8
            strTemp(m_i) = ""
         Next m_i
         strTemp(1) = CheckStr(MSHFlexGrid1.TextMatrix(i, 1))
         strTemp(2) = CheckStr(MSHFlexGrid1.TextMatrix(i, 2))
         strTemp(3) = StrToStr(CheckStr(MSHFlexGrid1.TextMatrix(i, 3)), 30)
         If Trim(strTemp(3)) <> Trim(CheckStr(MSHFlexGrid1.TextMatrix(i, 3))) Then
            strTemp(3) = strTemp(3) & ".."
         End If
         strTemp(4) = StrToStr(CheckStr(MSHFlexGrid1.TextMatrix(i, 4)), 30)
         If Trim(strTemp(4)) <> Trim(CheckStr(MSHFlexGrid1.TextMatrix(i, 4))) Then
            strTemp(4) = strTemp(4) & ".."
         End If
         strTemp(5) = CheckStr(MSHFlexGrid1.TextMatrix(i, 5))
         strTemp(6) = CheckStr(MSHFlexGrid1.TextMatrix(i, 6))
         strTemp(7) = CheckStr(MSHFlexGrid1.TextMatrix(i, 7))
         strTemp(8) = StrToStr(CheckStr(MSHFlexGrid1.TextMatrix(i, 8)), 20)
         'Added by Lydia 2022/07/04 FCP和FMP案之一案兩請僅其中一案更代時，彈提醒選擇
         If InStr(strJumpList & ",", MSHFlexGrid1.TextMatrix(i, 9) & MSHFlexGrid1.TextMatrix(i, 10) & MSHFlexGrid1.TextMatrix(i, 11) & MSHFlexGrid1.TextMatrix(i, 12) & ",") = 0 Then  '排除彈提醒選擇發email通知的案號
             strExc(0) = ""
             If intRe = 1 Then
                 If PUB_ChkFCforChange(MSHFlexGrid1.TextMatrix(i, 9), MSHFlexGrid1.TextMatrix(i, 10), MSHFlexGrid1.TextMatrix(i, 11), MSHFlexGrid1.TextMatrix(i, 12)) = False Then
                     strExc(0) = "N"
                     strJumpList = strJumpList & MSHFlexGrid1.TextMatrix(i, 9) & MSHFlexGrid1.TextMatrix(i, 10) & MSHFlexGrid1.TextMatrix(i, 11) & MSHFlexGrid1.TextMatrix(i, 12) & ","
                 End If
             End If
             If strExc(0) <> "N" Then
                 intRow = intRow + 1 '改在判斷後才合計筆數
         'end 2022/07/04
                 'If iLine > 54 Or iLine = 1 Then
                 'Modified by Morgan 2015/9/22
                 'If iLine > 37 Or iLine = 1 Then
                 If (iLine + 3) * 300 > Printer.ScaleHeight Or iLine = 1 Then
                 'end 2015/9/22
                    If strCaseNo <> "" Then Printer.NewPage
                    PrintTitle '列印表頭
                 End If
                 strCaseNo = strTemp(1)
                 StrCaseList = StrCaseList & "," & strTemp(1)
                 PrintDetail
         'Added by Lydia 2022/07/04
             End If
         End If
         'end 2022/07/04 ---- If InStr(strJumpList & ",", strTemp(1) & ",") = 0
      End If
   Next i
   
   If StrCaseList <> "" Then 'Added by Lydia 2022/07/04
      'Modified by Morgan 2015/9/22
      'If iLine > 37 Or iLine = 1 Then
      If (iLine + 3) * 300 > Printer.ScaleHeight Or iLine = 1 Then
      'end 2015/9/22
         If strCaseNo <> "" Then Printer.NewPage
         PrintTitle '列印表頭
      End If
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iLine * 300
      Printer.Print "共計 " & intRow & " 筆"
      Printer.EndDoc
      intQ = MsgBox("請確認清單是否已列印成功，有需要重新列印嗎？" & vbCrLf & "（後續將進行更新FC代理人資料）", vbYesNoCancel + vbDefaultButton1)
      If intQ = vbYes Then
         GoTo RePrint
      ElseIf intQ = vbCancel Then
         PUB_RestorePrinter strPrinter
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
   End If  'Added by Lydia 2022/07/04

   
   PUB_RestorePrinter strPrinter
   Screen.MousePointer = vbHourglass 'Added by Lydia 2025/07/23
   '更新資料
   If StrCaseList <> "" Then 'Added by Lydia 2022/07/04
      Call SaveData
   End If  'Added by Lydia 2022/07/04
   Screen.MousePointer = vbDefault 'Added by Lydia 2025/07/23
   
   'Added by Lydia 2019/12/19
   'Move by Lydia 2025/07/23 從PUB_RestorePrinter上方移下來; 改發給新程序管制人 ---Phoebe
   If StrCaseList <> "" And (Left(Pub_StrUserSt03, 2) = "F2" Or Pub_StrUserSt03 = "M51") Then
       strExc(1) = "Y"
       If Pub_StrUserSt03 = "M51" Then
           If MsgBox("是否發案件清單通知外專承辦和程序人員？", vbInformation + vbYesNo + vbDefaultButton1, "電腦中心") = vbNo Then
               strExc(1) = ""
           End If
       End If
       If strExc(1) = "Y" Then
           StrCaseList = Replace(Pub_RplStr(StrCaseList), "-", "")
           StrCaseList = Mid(StrCaseList, 2)
           'Modified by Lydia 2022/02/24 移到basFunction
           'If Me.PUB_ChgPA75List(StrCaseList) = False Then
           strExc(2) = "智權人員：" & frm110104_1.txtCaseField(4) & " " & frm110104_1.lblSname & vbCrLf & _
                    "變更案件條件：代理人：" & frm110104_1.txtCaseField(1) & " " & frm110104_1.lblAgent & vbCrLf & _
                    "　　　　　　　申請人：" & frm110104_1.txtCaseField(2) & " " & frm110104_1.lblCustomer & vbCrLf & _
                    "新代理人：" & frm110104_1.txtCaseField(3) & " " & frm110104_1.NewAgent & vbCrLf & _
                    "　　　　　　　" & IIf(frm110104_1.Check1.Value = 1, "■", "□") & "含閉卷或銷卷案件　　　　　　" & IIf(frm110104_1.Check4.Value = 1, "■", "□") & "清除案件聯絡人資料" & vbCrLf & _
                    "　　　　　　　" & IIf(frm110104_1.Check2.Value = 1, "■", "□") & "彼所案號清除　　　　　　　　" & IIf(frm110104_1.Check3.Value = 1, "■", "□") & "案件聯絡人同時更改"
           If PUB_ChgPA75List(StrCaseList, "0", "", strSrvDate(1), strExc(2)) = False Then
           'end 2022/02/24
               MsgBox "發案件清單作業失敗，中止更新FC代理人！", vbCritical
               GoTo JumpToExit
           End If
       End If
   End If
   'end 2019/12/19
   
   'Added by Lydia 2022/07/04
   If strJumpList <> "" Then
       PUB_SendMailCache
   End If
   'end 2022/07/04
   
JumpToExit: 'Added by Lydia 2019/12/9
   Screen.MousePointer = vbDefault
   
   'Added by Lydia 2020/02/10
   If Err.Number <> 0 Then
       MsgBox Err.Description, vbCritical + vbOKOnly, "清單列印失敗"
   End If
End Sub

Private Function SaveData() As Boolean
Dim intCaseKind As Integer
Dim i As Long
Dim strCP10 As String, strCP09 As String, strCP110 As String
Dim strMCTF(0) As String, stMsg As String 'Add by Amy 2019/06/26
   
On Error GoTo ErrHand
   'Add by Amy 2019/06/26 取得新代理人之控管智權人員
   strExc(0) = GetCusORFagentData(ChangeCustomerL(frm110104_1.txtCaseField(3)), "FA120", strMCTF())
   
   cnnConnection.BeginTrans
   
   For i = 1 To MSHFlexGrid1.Rows - 1
      'Modified by Lydia 2022/07/04 排除彈提醒選擇發email通知的案號
      'If MSHFlexGrid1.TextMatrix(i, 0) = "V" Then
      If MSHFlexGrid1.TextMatrix(i, 0) = "V" And InStr(strJumpList & ",", MSHFlexGrid1.TextMatrix(i, 9) & MSHFlexGrid1.TextMatrix(i, 10) & MSHFlexGrid1.TextMatrix(i, 11) & MSHFlexGrid1.TextMatrix(i, 12) & ",") = 0 Then  '排除彈提醒選擇發email通知的案號
         If ClsPDGetSystemKind(MSHFlexGrid1.TextMatrix(i, 9), intCaseKind) Then
            'Modified by Morgan 2019/3/7 所有聯絡人欄位加 ChgSQL
            '更新基本檔
            Select Case intCaseKind
               Case 專利
                  strCP10 = "937"
                  'Modified by Sindy 2018/1/24 備註加 ChgSQL(代理人名稱可能有單引號)
                  'Modified by Lydia 2019/12/24 備註加「請留意最新指示及聯絡對象」
                  strSql = "UPDATE PATENT SET PA143=NULL,PA75='" & frm110104_1.txtCaseField(3) & "'" & _
                           IIf(frm110104_1.Check2.Value = 1, ",PA77=null", "") & _
                           IIf(frm110104_1.Check3.Value = 1, ",PA51=" & CNULL(ChgSQL(frm110104_1.Text2(0))) & ",PA52=" & CNULL(ChgSQL(frm110104_1.Text2(1))) & ",PA53=" & CNULL(ChgSQL(frm110104_1.Text2(2))) & ",PA54=" & CNULL(ChgSQL(frm110104_1.Text2(3))) & ",PA55=" & CNULL(ChgSQL(frm110104_1.Text2(4))) & ",PA56=" & CNULL(ChgSQL(frm110104_1.Text2(5))) & ",PA139=" & CNULL(ChgSQL(frm110104_1.Text2(6))), "") & _
                           IIf(frm110104_1.Check4.Value = 1, ",PA51=NULL,PA52=NULL,PA53=NULL,PA54=NULL,PA55=NULL,PA56=NULL,PA139=NULL", "") & _
                           ",PA91='" & ChgSQL(ChangeTStringToTDateString(strSrvDate(2)) & "換FC代理人,請留意最新指示及聯絡對象,原FC代理人" & Left(MSHFlexGrid1.TextMatrix(i, 3), 9) & "/" & Trim(Mid(MSHFlexGrid1.TextMatrix(i, 3), 10))) & ";'||PA91" & _
                           ",PA49=null,PA50=null,PA151=null,PA152=null" & _
                           " WHERE PA01='" & MSHFlexGrid1.TextMatrix(i, 9) & "'" & _
                             " AND PA02='" & MSHFlexGrid1.TextMatrix(i, 10) & "'" & _
                             " AND PA03='" & MSHFlexGrid1.TextMatrix(i, 11) & "'" & _
                             " AND PA04='" & MSHFlexGrid1.TextMatrix(i, 12) & "'"
               Case 商標
                  strCP10 = "726"
                  'Modified by Sindy 2018/1/24 備註加 ChgSQL(代理人名稱可能有單引號)
                  'Modify By Sindy 2025/3/10 + ,TM140=null,TM141=null
                  strSql = "UPDATE TRADEMARK SET TM44='" & frm110104_1.txtCaseField(3) & "'" & _
                           IIf(frm110104_1.Check2.Value = 1, ",TM45=null", "") & _
                           IIf(frm110104_1.Check3.Value = 1, ",TM38=" & CNULL(ChgSQL(frm110104_1.Text2(0))) & ",TM39=" & CNULL(ChgSQL(frm110104_1.Text2(1))) & ",TM40=" & CNULL(ChgSQL(frm110104_1.Text2(2))) & ",TM41=" & CNULL(ChgSQL(frm110104_1.Text2(3))) & ",TM42=" & CNULL(ChgSQL(frm110104_1.Text2(4))) & ",TM43=" & CNULL(ChgSQL(frm110104_1.Text2(5))) & ",TM76=" & CNULL(ChgSQL(frm110104_1.Text2(6))), "") & _
                           IIf(frm110104_1.Check4.Value = 1, ",TM38=NULL,TM39=NULL,TM40=NULL,TM41=NULL,TM42=NULL,TM43=NULL,TM76=NULL", "") & _
                           ",TM58='" & ChgSQL(ChangeTStringToTDateString(strSrvDate(2)) & "換FC代理人,請留意最新指示及聯絡對象,原FC代理人" & Left(MSHFlexGrid1.TextMatrix(i, 3), 9) & "/" & Trim(Mid(MSHFlexGrid1.TextMatrix(i, 3), 10))) & ";'||TM58" & _
                           ",TM36=null,TM37=null,TM140=null,TM141=null" & _
                           " WHERE TM01='" & MSHFlexGrid1.TextMatrix(i, 9) & "'" & _
                             " AND TM02='" & MSHFlexGrid1.TextMatrix(i, 10) & "'" & _
                             " AND TM03='" & MSHFlexGrid1.TextMatrix(i, 11) & "'" & _
                             " AND TM04='" & MSHFlexGrid1.TextMatrix(i, 12) & "'"
                  'Modify By Sindy 2022/7/26 在多筆更換時，是F1X人員操作時才清除部分案件資料
                  If Left(Trim(Pub_StrUserSt03), 2) = "F1" Then
                     cnnConnection.Execute strSql '*****
                     strSql = "UPDATE TRADEMARK SET TM58='" & ChangeTStringToTDateString(strSrvDate(2)) & "整批更代清除部分案件資料;'||TM58" & _
                              ",TM35=null,TM127=null,TM124=null,TM125=null,TM46=null,TM69=null,TM56=null,TM122=null,TM68=null,TM129=null" & _
                              ",TM33=null,TM65=null,TM71=null,TM66=null,TM121=null,TM70=null,TM126=null" & _
                              " WHERE TM01='" & MSHFlexGrid1.TextMatrix(i, 9) & "'" & _
                                " AND TM02='" & MSHFlexGrid1.TextMatrix(i, 10) & "'" & _
                                " AND TM03='" & MSHFlexGrid1.TextMatrix(i, 11) & "'" & _
                                " AND TM04='" & MSHFlexGrid1.TextMatrix(i, 12) & "'"
                  End If
                  '2022/7/26 END
               Case 法務
                  strCP10 = "994"
                  'Modified by Sindy 2018/1/24 備註加 ChgSQL(代理人名稱可能有單引號)
                  strSql = "UPDATE LAWCASE SET LC22='" & frm110104_1.txtCaseField(3) & "'" & _
                           IIf(frm110104_1.Check2.Value = 1, ",LC23=null", "") & _
                           IIf(frm110104_1.Check3.Value = 1, ",LC18=" & CNULL(ChgSQL(frm110104_1.Text2(0))) & ",LC19=" & CNULL(ChgSQL(frm110104_1.Text2(1))) & ",LC20=" & CNULL(ChgSQL(frm110104_1.Text2(2))) & ",LC39=" & CNULL(ChgSQL(frm110104_1.Text2(6))), "") & _
                           IIf(frm110104_1.Check4.Value = 1, ",LC18=NULL,LC19=NULL,LC20=NULL,LC39=NULL", "") & _
                           ",LC27='" & ChgSQL(ChangeTStringToTDateString(strSrvDate(2)) & "換FC代理人,請留意最新指示及聯絡對象,原FC代理人" & Left(MSHFlexGrid1.TextMatrix(i, 3), 9) & "/" & Trim(Mid(MSHFlexGrid1.TextMatrix(i, 3), 10))) & ";'||LC27" & _
                           ",LC24=null" & _
                           " WHERE LC01='" & MSHFlexGrid1.TextMatrix(i, 9) & "'" & _
                             " AND LC02='" & MSHFlexGrid1.TextMatrix(i, 10) & "'" & _
                             " AND LC03='" & MSHFlexGrid1.TextMatrix(i, 11) & "'" & _
                             " AND LC04='" & MSHFlexGrid1.TextMatrix(i, 12) & "'"
               Case Else '服務
                  If MSHFlexGrid1.TextMatrix(i, 9) = "FG" Or _
                     MSHFlexGrid1.TextMatrix(i, 9) = "PS" Or _
                     MSHFlexGrid1.TextMatrix(i, 9) = "CPS" Then
                     strCP10 = "937"
                  Else
                     strCP10 = "726"
                  End If
                  'Modified by Sindy 2018/1/24 備註加 ChgSQL(代理人名稱可能有單引號)
                  strSql = "UPDATE SERVICEPRACTICE SET SP26='" & frm110104_1.txtCaseField(3) & "'" & _
                           IIf(frm110104_1.Check2.Value = 1, ",SP27=null", "") & _
                           IIf(frm110104_1.Check3.Value = 1, ",SP30=" & CNULL(ChgSQL(frm110104_1.Text2(0))) & ",SP75=" & CNULL(ChgSQL(frm110104_1.Text2(3))) & ",SP71=" & CNULL(ChgSQL(frm110104_1.Text2(6))), "") & _
                           IIf(frm110104_1.Check4.Value = 1, ",SP30=NULL,SP75=NULL,SP71=NULL", "") & _
                           ",SP18='" & ChgSQL(ChangeTStringToTDateString(strSrvDate(2)) & "換FC代理人,請留意最新指示及聯絡對象,原FC代理人" & Left(MSHFlexGrid1.TextMatrix(i, 3), 9) & "/" & Trim(Mid(MSHFlexGrid1.TextMatrix(i, 3), 10))) & ";'||SP18" & _
                           ",SP31=null" & _
                           " WHERE SP01='" & MSHFlexGrid1.TextMatrix(i, 9) & "'" & _
                             " AND SP02='" & MSHFlexGrid1.TextMatrix(i, 10) & "'" & _
                             " AND SP03='" & MSHFlexGrid1.TextMatrix(i, 11) & "'" & _
                             " AND SP04='" & MSHFlexGrid1.TextMatrix(i, 12) & "'"
                  'Modify By Sindy 2022/7/26 在多筆更換時，是F1X人員操作時才清除部分案件資料
                  If Left(Trim(Pub_StrUserSt03), 2) = "F1" Then
                     cnnConnection.Execute strSql '*****
                     strSql = "UPDATE SERVICEPRACTICE SET SP18='" & ChangeTStringToTDateString(strSrvDate(2)) & "整批更代清除部分案件資料;'||SP18" & _
                              ",SP29=null,SP84=null,SP81=null,SP82=null,SP33=null,SP67=null,SP37=null,SP80=null,SP83=null" & _
                              " WHERE SP01='" & MSHFlexGrid1.TextMatrix(i, 9) & "'" & _
                                " AND SP02='" & MSHFlexGrid1.TextMatrix(i, 10) & "'" & _
                                " AND SP03='" & MSHFlexGrid1.TextMatrix(i, 11) & "'" & _
                                " AND SP04='" & MSHFlexGrid1.TextMatrix(i, 12) & "'"
                  End If
                  '2022/7/26 END
            End Select
            cnnConnection.Execute strSql '*****
            '新增案件進度檔
            strCP09 = AutoNo("B", 6)
            '取得出名代理人
            strCP110 = ""
'CANCEL BY SONIA 2015/6/17 FCT-024182各式申請書抓最新A,B類發文之CP110會抓到此進度
'            strExc(0) = "select cp110 from caseprogress" & _
'                        " where cp09=(select substr(max(cp27||cp09),9) from caseprogress" & _
'                        " WHERE cp01='" & MSHFlexGrid1.TextMatrix(i, 9) & "' and cp02='" & MSHFlexGrid1.TextMatrix(i, 10) & "' and cp03='" & MSHFlexGrid1.TextMatrix(i, 11) & "' and cp04='" & MSHFlexGrid1.TextMatrix(i, 12) & "'" & _
'                        " and cp09<'C'" & _
'                        " and cp110 is not null and cp27 is not null" & _
'                        " group by cp01,cp02,cp03,cp04)"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               strCP110 = RsTemp.Fields(0)
'            End If
'END 2015/6/17
            'Modified by Morgan 2016/8/22 備註加 ChgSQL(代理人名稱可能有單引號)
            'Modified by Lydia 2019/12/24 備註加「請留意最新指示及聯絡對象」
            strSql = "INSERT INTO CASEPROGRESS(CP09,CP01,CP02,CP03,CP04,CP05,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP27,CP64,cp82,cp83,cp110)" & _
                     " values(" & CNULL(strCP09) & "," & CNULL(MSHFlexGrid1.TextMatrix(i, 9)) & "," & CNULL(MSHFlexGrid1.TextMatrix(i, 10)) & "," & CNULL(MSHFlexGrid1.TextMatrix(i, 11)) & "," & CNULL(MSHFlexGrid1.TextMatrix(i, 12)) & _
                     "," & strSrvDate(1) & "," & CNULL(strCP10) & ",'90'," & CNULL(PUB_GetStaffST15(frm110104_1.txtCaseField(4), 1)) & "," & CNULL(frm110104_1.txtCaseField(4)) & "," & CNULL(strUserNum) & ",'N','N'" & _
                     "," & strSrvDate(1) & ",'" & ChgSQL(ChangeTStringToTDateString(strSrvDate(2)) & "換FC代理人,請留意最新指示及聯絡對象,原FC代理人" & Left(MSHFlexGrid1.TextMatrix(i, 3), 9) & "/" & Trim(Mid(MSHFlexGrid1.TextMatrix(i, 3), 10))) & ";'" & _
                     ",substr(to_char(sysdate,'yyyymmddhh24mmss'),9)," & CNULL(strUserNum) & "," & CNULL(strCP110) & ")"
            cnnConnection.Execute strSql
            
            'Add by Amy 2019/06/26 依案號更新當日AB類收文之 收文MCTF組別
            If intCaseKind <> 專利 Then
                stMsg = "Y"
                If UpdCP161(MSHFlexGrid1.TextMatrix(i, 9) & ";" & MSHFlexGrid1.TextMatrix(i, 10) & ";" & MSHFlexGrid1.TextMatrix(i, 11) & ";" & MSHFlexGrid1.TextMatrix(i, 12), strMCTF(0), stMsg) = False Then GoTo ErrHand
            End If
         End If
      End If
   Next i
   
   cnnConnection.CommitTrans
   cmdOK.Enabled = False
   MsgBox "更新完成!!", vbOKOnly, "執行成功"
   Call cmdBack_Click
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   'Modify by Amy 2019/06/26
'   If Err.Number <> 0 Then
'      MsgBox Err.Description, vbCritical
'   End If
   If stMsg = MsgText(601) And Err.Number <> 0 Then stMsg = Err.Description
   MsgBox stMsg, vbCritical
   'end 2019/06/26
End Function

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 2200
PLeft(3) = 3700
PLeft(4) = 10500
PLeft(5) = 14500
PLeft(6) = 15500
End Sub

Sub PrintTitle()
GetPleft
iLine = 1
Printer.Font.Size = 14
Printer.Font.Underline = False
Printer.FontBold = True
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("更換FC代理人清單") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "更換FC代理人清單"
Printer.FontBold = False
Printer.Font.Size = 10
iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "列印人員：" & strUserName
Printer.CurrentX = 7300
Printer.CurrentY = iLine * 300
Printer.Print "智權人員：" & frm110104_1.txtCaseField(4) & " " & frm110104_1.lblSname
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "變更案件條件：代理人：" & frm110104_1.txtCaseField(1) & " " & frm110104_1.lblAgent
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "　　　　　　　申請人：" & frm110104_1.txtCaseField(2) & " " & frm110104_1.lblCustomer
Printer.CurrentX = 7300
Printer.CurrentY = iLine * 300
Printer.Print "新代理人：" & frm110104_1.txtCaseField(3) & " " & frm110104_1.NewAgent
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "　　　　　　　" & IIf(frm110104_1.Check1.Value = 1, "■", "□") & "含閉卷或銷卷案件　　　　　　" & IIf(frm110104_1.Check4.Value = 1, "■", "□") & "清除案件聯絡人資料"
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "　　　　　　　" & IIf(frm110104_1.Check2.Value = 1, "■", "□") & "彼所案號清除　　　　　　　　" & IIf(frm110104_1.Check3.Value = 1, "■", "□") & "案件聯絡人同時更改"

iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "本所案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "申請號/審定號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "原FC代理人"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "彼所案號"

iLine = iLine + 1
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "申請人1"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "閉卷日"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "北所銷卷日"

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(255, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(1)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(3)
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(5)
   iLine = iLine + 1
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(4)
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(8)
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(6)
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(7)
   iLine = iLine + 1
   'Add By Sindy 2015/8/5
   '操作人員為F2x時加印每一案件之下一程序期限
   If Left(Pub_StrUserSt03, 2) = "F2" Then
      Call PrintDetail_2
   End If
   '2015/8/5 END
End Sub

'Add By Sindy 2015/8/5
Sub PrintDetail_2()
Dim AdoRs As ADODB.Recordset
Dim strTemp2(1 To 4) As String
   
   '下一程序:非程序管制期限
   'Modify by Amy 2017/06/27 拿掉 & strNpSqlOfNoSalesDuty 條件 for 更代完成將下一程序未發文備註印出-葉敏莉
   strSql = "select decode(pa09,'000',cpm03,cpm04),sqldatet(np08),sqldatet(np09),np15" & _
            " From nextprogress,casepropertymap,patent" & _
            " where np02='" & SystemNumber(strCaseNo, 1) & "' and np03='" & SystemNumber(strCaseNo, 2) & "' and np04='" & SystemNumber(strCaseNo, 3) & "' and np05='" & SystemNumber(strCaseNo, 4) & "'" & _
            " and np06 is null" & _
            " and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+)" & _
            " and np02=cpm01(+) and np07=cpm02(+)" & _
            " order by np01 asc"
   intI = 1
   Set AdoRs = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      AdoRs.MoveFirst
      'Modified by Morgan 2015/9/22
      'If iLine > 37 Or iLine = 1 Then
      If (iLine + 3) * 300 > Printer.ScaleHeight Or iLine = 1 Then
      'end 2015/9/22
         Printer.NewPage
         '列印表頭
         PrintTitle
      End If
      PrintTitle_2 '列印下一程序表頭
      Do While Not AdoRs.EOF
         strTemp2(1) = StrToStr("" & AdoRs.Fields(0), 30)
         strTemp2(2) = StrToStr("" & AdoRs.Fields(1), 9)
         strTemp2(3) = StrToStr("" & AdoRs.Fields(2), 9)
         strTemp2(4) = StrToStr("" & AdoRs.Fields(3), 30)
         'Modified by Morgan 2015/9/22
         'If iLine > 37 Or iLine = 1 Then
         If (iLine + 3) * 300 > Printer.ScaleHeight Or iLine = 1 Then
         'end 2015/9/22
            Printer.NewPage
            PrintTitle '列印表頭
            PrintTitle_2 '列印下一程序表頭
         End If
         Printer.CurrentX = Pleft2(1)
         Printer.CurrentY = iLine * 300
         Printer.Print strTemp2(1)
         Printer.CurrentX = Pleft2(2)
         Printer.CurrentY = iLine * 300
         Printer.Print strTemp2(2)
         Printer.CurrentX = Pleft2(3)
         Printer.CurrentY = iLine * 300
         Printer.Print strTemp2(3)
         Printer.CurrentX = Pleft2(4)
         Printer.CurrentY = iLine * 300
         Printer.Print strTemp2(4)
         iLine = iLine + 1
         
         AdoRs.MoveNext
      Loop
   End If
   Set AdoRs = Nothing
End Sub

'Add By Sindy 2015/8/5
Sub GetPleft_2()
Pleft2(1) = 3700
Pleft2(2) = 7500
Pleft2(3) = 9000
Pleft2(4) = 10500
End Sub

'Add By Sindy 2015/8/5
Sub PrintTitle_2()
GetPleft_2
Printer.Font.Size = 10
Printer.CurrentX = 2200
Printer.CurrentY = iLine * 300
Printer.Print "下一程序："
Printer.CurrentX = Pleft2(1)
Printer.CurrentY = iLine * 300
Printer.Print "案件性質"
Printer.CurrentX = Pleft2(2)
Printer.CurrentY = iLine * 300
Printer.Print "本所期限"
Printer.CurrentX = Pleft2(3)
Printer.CurrentY = iLine * 300
Printer.Print "法定期限"
Printer.CurrentX = Pleft2(4)
Printer.CurrentY = iLine * 300
Printer.Print "備　　註"
iLine = iLine + 1
Printer.CurrentX = Pleft2(1)
Printer.CurrentY = iLine * 300
Printer.Print String(203, "-")
iLine = iLine + 1
End Sub

Private Sub cmdSelect_Click()
Dim i As Integer, iCol As Integer
   
   If cmdSelect.Caption = "全部選取(&A)" Then
      With MSHFlexGrid1
         For i = 1 To MSHFlexGrid1.Rows - 1
            .col = 0
            .row = i
            .Text = "V"
            For iCol = 0 To .Cols - 1
              .col = iCol
              .CellBackColor = &HFFC0C0
            Next
         Next
      End With
      lngSelRow = Me.MSHFlexGrid1.Rows - 1
      cmdSelect.Caption = "全部取消(&R)"
   Else
      With MSHFlexGrid1
         For i = 1 To MSHFlexGrid1.Rows - 1
            .col = 0
            .row = i
            .Text = ""
            For iCol = 0 To .Cols - 1
              .col = iCol
              .CellBackColor = QBColor(15)
            Next
         Next
      End With
      lngSelRow = 0
      cmdSelect.Caption = "全部選取(&A)"
   End If
   Me.Frame2.Caption = "案件筆數：共 " & Me.MSHFlexGrid1.Rows - 1 & " 筆!!!　已選取 " & lngSelRow & " 筆 !"
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Set MSHFlexGrid1.Recordset = RsTemp
   Grid
   PUB_SetPrinter Me.Name, Combo1, strPrinter
End Sub

Private Sub Grid()
   With MSHFlexGrid1
      .Visible = False
      .Cols = 13
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "V"
      .col = 1: .ColWidth(1) = 1500: .Text = "本所案號"
      .col = 2: .ColWidth(2) = 1000: .Text = "申請案號/審定號"
      .col = 3: .ColWidth(3) = 1500: .Text = "原FC代理人"
      .col = 4: .ColWidth(4) = 1500: .Text = "申請人1"
      .col = 5: .ColWidth(5) = 1000: .Text = "彼所案號"
      .col = 6: .ColWidth(6) = 800: .Text = "閉卷日"
      .col = 7: .ColWidth(7) = 800: .Text = "北所銷卷日"
      .col = 8: .ColWidth(8) = 2000: .Text = "案件名稱"
      .col = 9: .ColWidth(9) = 0: .Text = "PA01"
      .col = 10: .ColWidth(10) = 0: .Text = "PA02"
      .col = 11: .ColWidth(11) = 0: .Text = "PA03"
      .col = 12: .ColWidth(12) = 0: .Text = "PA04"
      .Visible = True
   End With
   lngSelRow = 0
   Me.Frame2.Caption = "案件筆數：共 " & Me.MSHFlexGrid1.Rows - 1 & " 筆!!!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   
   Set frm110104_2 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
Dim nCol As Integer, nRow As Integer ', iRow As Integer, iCol As Integer
   
   With MSHFlexGrid1
      .Visible = False
      nCol = .MouseCol
      nRow = .MouseRow
      If nRow = 0 Then
         .col = nCol
         If m_blnColOrderAsc1 = False Then '字串降冪
            .Sort = 5 '字串昇冪
            m_blnColOrderAsc1 = True
         Else
            .Sort = 6 '字串降冪
            m_blnColOrderAsc1 = False
         End If
      ElseIf nRow > 0 And .TextMatrix(nRow, 1) <> "" Then
         SelectRow nRow, MSHFlexGrid1
      End If
      .Visible = True
   End With
   Me.Frame2.Caption = "案件筆數：共 " & Me.MSHFlexGrid1.Rows - 1 & " 筆!!!　已選取 " & lngSelRow & " 筆 !"
End Sub

Private Sub SelectRow(ByRef pRow As Integer, ByRef FlexGrid As MSHFlexGrid)
Dim iCol As Integer
   
   With FlexGrid
      If pRow > 0 Then
         .row = pRow
         If Trim(.TextMatrix(pRow, 0)) = "" Then
            .TextMatrix(pRow, 0) = "V"
            lngSelRow = lngSelRow + 1
            For iCol = 0 To .Cols - 1
              .col = iCol
              .CellBackColor = &HFFC0C0
            Next
         Else
            .TextMatrix(pRow, 0) = ""
            lngSelRow = lngSelRow - 1
            For iCol = 0 To .Cols - 1
              .col = iCol
              .CellBackColor = QBColor(15)
            Next
         End If
      End If
   End With
End Sub

