VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc11r0 
   AutoRedraw      =   -1  'True
   Caption         =   "國內應收待處理作業"
   ClientHeight    =   5290
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5290
   ScaleWidth      =   8760
   Begin VB.CommandButton Command1 
      Caption         =   "查詢"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7260
      TabIndex        =   0
      Top             =   60
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid DataGrid1 
      Height          =   4725
      Left            =   60
      TabIndex        =   1
      Top             =   390
      Width           =   8625
      _ExtentX        =   15222
      _ExtentY        =   8326
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "待處理內容|說明"
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "銷貨折讓單未處理 (轉檔)及(上傳)及列印作業"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4995
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc11r0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/08 Form2.0已檢查 (無需修改的物件)
'Create by Sindy 2014/1/2
Option Explicit

Const qST06Name As String = "decode(ST06,'1','北所','2','中所','3','南所','4','高所',st06)"


Private Sub SetDataListWidth()
   Dim iCol As Integer
   
   DataGrid1.row = 0
   DataGrid1.ColAlignment = flexAlignLeftCenter
   
   iCol = 0
   DataGrid1.col = iCol: DataGrid1.Text = "待處理內容"
   DataGrid1.ColWidth(iCol) = 2000
   
   iCol = iCol + 1 '1
   DataGrid1.col = iCol: DataGrid1.Text = "說明"
   DataGrid1.ColWidth(iCol) = 6250
End Sub

Private Sub Command1_Click()
   Call QueryData
   SetDataListWidth
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8880
   Me.Height = 5700
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   Call Command1_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strConTitle = MsgText(601)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc11r0 = Nothing
End Sub

Private Function QueryData()
Dim dblCnt As Double, dblCnt2 As Double
Dim strData As String
Dim strA0K11 As String, strDate As String, strVal As String, i As Integer, sqlST06 As String 'Add By Sidy 2021/6/2
   
   Screen.MousePointer = vbHourglass
   DataGrid1.Clear
   DataGrid1.Rows = 2
   
   '****************************************************
   '發票號碼剩下250號提醒 '2015/8/7 modify by sonia 改剩50號提醒
   '****************************************************
   strSql = "select a4101,a4102,nvl(a4105,0)-nvl(a4110,0) from acc410" & _
            " where (" & Left(strSrvDate(2), 5) & " between a4101 and a4102) and a4109='1'" & _
            " and nvl(a4105,0)-nvl(a4110,0)<=50" & _
            " and a4107 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If RsTemp.RecordCount > 0 Then
         Call SetDrdColData("發票號碼將用罄", "發票號碼尚餘" & RsTemp.Fields(2) & "號", True)
      End If
   End If
   
   '****************************************************
   '收據
   '****************************************************
   'Modify By Sindy 2025/7/31 瑞婷提:國內應收待處理作業之「待列印收據」改為「暫不列印放出收據」
   '檢查"暫不列印放出收據"張數
   'Modify By 2021/6/2 => and a0k11 not in('J','L') and a0k20=st01(+)
   strSql = "select a0k11,st06," & qST06Name & ",count(*) from acc0k0,staff" & _
            " where a0k32='Y' and a0k11 not in('J','L') and a0k20=st01(+)" & _
            " and (a0k19=0 or a0k19 is null)" & _
            " group by a0k11,st06," & qST06Name & " order by a0k11,st06," & qST06Name
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strData = "": strA0K11 = ""
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If strA0K11 = "" Or strA0K11 <> RsTemp.Fields(0) Then
            If strA0K11 <> "" Then Call SetDrdColData("暫不列印放出收據", strData)
            strData = RsTemp.Fields(0) & "公司"
         End If
         strData = strData & "　" & RsTemp.Fields(2) & " " & RsTemp.Fields(3) & "張"
         strA0K11 = RsTemp.Fields(0)
         RsTemp.MoveNext
      Loop
      Call SetDrdColData("暫不列印放出收據", strData)
   End If
   'Add By Sindy 2021/6/2 L法律所
   'Modify By Sindy 2021/7/26 若為案源,以介紹人的所別判斷
   'Modified by Morgan 2025/2/13 1張收據可能會有多個收文號 count(*)->count(distinct a0k01) Ex:E11404645
   strVal = "select a0k11,st06," & qST06Name & " as st06name,count(distinct a0k01) as cnt from acc0k0,staff,caseprogress,lawofficesource" & _
            " where a0k32='Y' and a0k11='L' AND cp60(+)=a0k01 AND Los15(+)=cp162" & _
            " and decode(substr(los04,1,5),NULL,a0k20,substr(los04,1,5))=st01(+)" & _
            " and (a0k19=0 or a0k19 is null)"
   strSql = "select a0k11,st06,st06name,sum(cnt) from(" & _
            strVal & " and exists (" & Replace(strLOSSalesDuty, "\#ST06SQL#\", " and st06 is not null") & ")" & _
            " group by a0k11,st06," & qST06Name & _
            " union " & _
            strVal & " and st06 is not null and not exists (" & Replace(strLOSSalesDuty, "\#ST06SQL#\", "") & ")" & _
            " group by a0k11,st06," & qST06Name & _
            ") group by a0k11,st06,st06name order by a0k11,st06,st06name"
   intI = 1
   strData = ""
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If strData = "" Then strData = RsTemp.Fields(0) & "公司"
         strData = strData & "　" & RsTemp.Fields(2) & " " & RsTemp.Fields(3) & "張"
         RsTemp.MoveNext
      Loop
   End If
   If strData <> "" Then Call SetDrdColData("暫不列印放出收據", strData)
   '2021/6/2 END
   
   '檢查是否有應列印但未列印的收據日期
   'MODIFY BY SONAI 2014/11/18 加大於系統日的資料-瑞婷
   'strSql = "select sqldatet(a0k02),a0k11,count(*) from acc0k0" & _
            " where a0k02>=920201 and a0k02<" & strSrvDate(2) & " and a0k32 is null and a0k10 is null" & _
            " and nvl(a0k09,0)=0 and nvl(a0k19,0)=0 and a0k11<>'J'" & _
            " group by a0k02,a0k11" & _
            " order by a0k02,a0k11"
   'Modify By 2021/6/2 => and a0k11 not in('J','L') and a0k20=st01(+)
   'modify by sonia 2021/7/23 取消非當日的限制
   'strSql = "select sqldatet(a0k02),a0k11,st06," & qST06Name & ",count(*) from acc0k0,staff" & _
            " where a0k02>=920201 and (a0k02<" & strSrvDate(2) & " or a0k02>" & strSrvDate(2) & ") and a0k32 is null and a0k10 is null" & _
            " and nvl(a0k09,0)=0 and nvl(a0k19,0)=0 and a0k11 not in('J','L') and a0k20=st01(+)" & _
            " group by a0k02,a0k11,st06," & qST06Name & _
            " order by a0k02,a0k11,st06," & qST06Name
   'modify by sonia 2022/8/11 取消銷帳條件故a0k10 is null改為(a0k37 is null or a0k37<>'N')
   strSql = "select sqldatet(a0k02),a0k11,st06," & qST06Name & ",count(*) from acc0k0,staff" & _
            " where a0k02>=920201 and a0k32 is null and (a0k37 is null or a0k37<>'N')" & _
            " and nvl(a0k09,0)=0 and nvl(a0k19,0)=0 and a0k11 not in('J','L') and a0k20=st01(+)" & _
            " group by a0k02,a0k11,st06," & qST06Name & _
            " order by a0k02,a0k11,st06," & qST06Name
   'end 2021/7/23
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strData = "": strDate = "": strA0K11 = ""
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If strDate & strA0K11 = "" Or strDate & strA0K11 <> RsTemp.Fields(0) & RsTemp.Fields(1) Then
            'modify by sonia 2021/7/23 取消非當日的限制
            'If strDate & strA0K11 <> "" Then Call SetDrdColData("非今日未列印收據", strData)
            If strDate & strA0K11 <> "" Then Call SetDrdColData("未列印收據", strData)
            'end 2021/7/23
            strData = RsTemp.Fields(0) & "（" & RsTemp.Fields(1) & "公司）"
         End If
         strData = strData & "　" & RsTemp.Fields(3) & " " & RsTemp.Fields(4) & "張"
         strDate = RsTemp.Fields(0)
         strA0K11 = RsTemp.Fields(1)
         RsTemp.MoveNext
      Loop
      'modify by sonia 2021/7/23 取消非當日的限制
      'Call SetDrdColData("非今日未列印收據", strData)
      Call SetDrdColData("未列印收據", strData)
      'end 2021/7/23
   End If
   'Add By Sindy 2021/6/2 L法律所
   'modify by sonia 2021/7/23 取消非當日的限制
   'strVal = "select sqldatet(a0k02) as a0k02,a0k11,st06," & qST06Name & " as st06name,count(*) as cnt from acc0k0,staff" & _
            " where a0k02>=920201 and (a0k02<" & strSrvDate(2) & " or a0k02>" & strSrvDate(2) & ") and a0k32 is null and a0k10 is null" & _
            " and nvl(a0k09,0)=0 and nvl(a0k19,0)=0 and a0k11='L' and a0k20=st01(+)"
   'Modify By Sindy 2021/7/26 若為案源,以介紹人的所別判斷
   'modify by sonia 2022/8/11 取消銷帳條件故a0k10 is null改為(a0k37 is null or a0k37<>'N')
   'Modified by Morgan 2025/2/13 1張收據可能會有多個收文號 count(*)->count(distinct a0k01) Ex:E11404645
   strVal = "select sqldatet(a0k02) as a0k02,a0k11,st06," & qST06Name & " as st06name,count(distinct a0k01) as cnt" & _
            " from acc0k0,staff,caseprogress,lawofficesource" & _
            " where a0k02>=920201 and a0k32 is null and (a0k37 is null or a0k37<>'N') AND cp60(+)=a0k01 AND Los15(+)=cp162" & _
            " and nvl(a0k09,0)=0 and nvl(a0k19,0)=0 and a0k11='L' and decode(substr(los04,1,5),NULL,a0k20,substr(los04,1,5))=st01(+)"
   'end 2021/7/23
   strSql = "select a0k02,a0k11,st06,st06name,sum(cnt) from(" & _
            strVal & " and exists (" & Replace(strLOSSalesDuty, "\#ST06SQL#\", " and st06 is not null") & ")" & _
            " group by a0k02,a0k11,st06," & qST06Name & _
            " union " & _
            strVal & " and st06 is not null and not exists (" & Replace(strLOSSalesDuty, "\#ST06SQL#\", "") & ")" & _
            " group by a0k02,a0k11,st06," & qST06Name & _
            ") group by a0k02,a0k11,st06,st06name order by a0k02,a0k11,st06,st06name"
   intI = 1
   strData = ""
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         'Modify By Sindy 2022/4/29 不同日期產生不同資料列
         If strData = "" Then
            strData = RsTemp.Fields(0) & "（" & RsTemp.Fields(1) & "公司）"
         ElseIf (InStr(strData, RsTemp.Fields(0)) = 0) Then
            Call SetDrdColData("未列印收據", strData)
            strData = RsTemp.Fields(0) & "（" & RsTemp.Fields(1) & "公司）"
         End If
         '2022/4/29 END
         strData = strData & "　" & RsTemp.Fields(3) & " " & RsTemp.Fields(4) & "張"
         RsTemp.MoveNext
      Loop
   End If
   'modify by sonia 2021/7/23 取消非當日的限制
   'If strData <> "" Then Call SetDrdColData("非今日未列印收據", strData)
   If strData <> "" Then Call SetDrdColData("未列印收據", strData)
   'end 2021/7/23
   '2021/6/2 END
   
   '****************************************************
   '請款單
   '****************************************************
   '檢查待列印請款單張數
   'Modify By 2021/6/2 => and a0k20=st01(+)
   strSql = "select a0k11,st06," & qST06Name & ",count(*) from acc0k0,staff" & _
            " where a0k32='Y' and a0k11='J'" & _
            " and (a0k19=0 or a0k19 is null) and a0k20=st01(+)" & _
            " group by a0k11,st06," & qST06Name & " order by a0k11,st06," & qST06Name
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strData = "": strA0K11 = ""
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If strA0K11 = "" Or strA0K11 <> RsTemp.Fields(0) Then
            If strA0K11 <> "" Then Call SetDrdColData("待列印請款單", strData)
            strData = RsTemp.Fields(0) & "公司"
         End If
         strData = strData & "　" & RsTemp.Fields(2) & " " & RsTemp.Fields(3) & "張"
         strA0K11 = RsTemp.Fields(0)
         RsTemp.MoveNext
      Loop
      Call SetDrdColData("待列印請款單", strData)
   End If
   '檢查是否有應列印但未列印的請款單日期
   'MODIFY BY SONAI 2014/11/18 加大於系統日的資料-瑞婷
   'strSql = "select sqldatet(a0k02),a0k11,count(*) from acc0k0" & _
            " where a0k02>=920201 and a0k02<" & strSrvDate(2) & " and a0k32 is null and a0k10 is null" & _
            " and nvl(a0k09,0)=0 and nvl(a0k19,0)=0 and a0k11='J'" & _
            " group by a0k02,a0k11" & _
            " order by a0k02,a0k11"
   'Modify By 2021/6/2 => and a0k20=st01(+)
   'modify by sonia 2021/7/23 取消非當日的限制
   'strSql = "select sqldatet(a0k02),a0k11,st06," & qST06Name & ",count(*) from acc0k0,staff" & _
            " where a0k02>=920201 and (a0k02<" & strSrvDate(2) & " or a0k02>" & strSrvDate(2) & ") and a0k32 is null and a0k10 is null" & _
            " and nvl(a0k09,0)=0 and nvl(a0k19,0)=0 and a0k11='J' and a0k20=st01(+)" & _
            " group by a0k02,a0k11,st06," & qST06Name & _
            " order by a0k02,a0k11,st06," & qST06Name
   'modify by sonia 2022/8/11 取消銷帳條件故a0k10 is null改為(a0k37 is null or a0k37<>'N')
   strSql = "select sqldatet(a0k02),a0k11,st06," & qST06Name & ",count(*) from acc0k0,staff" & _
            " where a0k02>=920201 and a0k32 is null and (a0k37 is null or a0k37<>'N')" & _
            " and nvl(a0k09,0)=0 and nvl(a0k19,0)=0 and a0k11='J' and a0k20=st01(+)" & _
            " group by a0k02,a0k11,st06," & qST06Name & _
            " order by a0k02,a0k11,st06," & qST06Name
   'end 2021/7/23
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strData = "": strDate = "": strA0K11 = ""
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If strDate & strA0K11 = "" Or strDate & strA0K11 <> RsTemp.Fields(0) & RsTemp.Fields(1) Then
            'modify by sonia 2021/7/23 取消非當日的限制
            'If strDate & strA0K11 <> "" Then Call SetDrdColData("非今日未列印請款單", strData)
            If strDate & strA0K11 <> "" Then Call SetDrdColData("未列印請款單", strData)
            'end 2021/7/23
            strData = RsTemp.Fields(0) & "（" & RsTemp.Fields(1) & "公司）"
         End If
         strData = strData & "　" & RsTemp.Fields(3) & " " & RsTemp.Fields(4) & "張"
         strDate = RsTemp.Fields(0)
         strA0K11 = RsTemp.Fields(1)
         RsTemp.MoveNext
      Loop
      'modify by sonia 2021/7/23 取消非當日的限制
      'Call SetDrdColData("非今日未列印請款單", strData)
      Call SetDrdColData("未列印請款單", strData)
      'end 2021/7/23
   End If
   
   '****************************************************
   '未列印發票
   '****************************************************
   strSql = "select sqldatet(a4302),count(*) from acc430" & _
            " Where nvl(a4308, 0) = 0 And a4307 Is Null" & _
            " group by a4302"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strData = ""
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If strData <> "" Then strData = strData & "；"
         strData = strData & RsTemp.Fields(0) & "共" & RsTemp.Fields(1) & "張"
         RsTemp.MoveNext
      Loop
      Call SetDrdColData("未列印發票", strData)
   End If
   '****************************************************
   '銷退折讓單列印
   '****************************************************
   'Modify by Amy 2019/12/16 發票上傳上線後改抓發票上傳日
'   strSql = "select a4301 from acc430" & _
'            " Where nvl(a4310,0)>0 and a4318 is null" & _
'            " Union All" & _
'            " select a0s01 from acc0s0" & _
'            " where a0s27='N' and a0s25 is null"
   strSql = "select a4301 from acc430" & _
            " Where nvl(a4310,0)>=" & TranInvoiceDate & " and nvl(a4324,0)=0 " & _
            " Union All" & _
            " select a0s01 from acc0s0" & _
            " where nvl(a0s03,0)>=" & TranInvoiceDate & " and a0s27='N' and nvl(a0s28,0)=0 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      dblCnt = RsTemp.RecordCount
      If dblCnt > 0 Then
         'Modify by Amy 2019/12/17 原:未列印
         Call SetDrdColData("銷退折讓單未上傳", "尚有" & dblCnt & "筆資料未上傳")
      End If
   End If
   '****************************************************
   '已收款請款單但未開發票
   '****************************************************
   'modify by sonia 2021/3/31 +加判斷非ACS代收代付706(ACS代收代付不可開發票)
   'strSql = "select a0m02 from acc0l0,acc0m0,acc0k0,acc431" & _
            " where a0l02 between " & ChangeWStringToTString(CompDate(1, -2, strSrvDate(2))) & " and " & strSrvDate(2) & _
            " and a0l01=a0m01" & _
            " and a0m02=a0k01" & _
            " and a0k11='J' and a0k37='Y'" & _
            " and a0k01=axc02(+)" & _
            " and axc01 is null"
   strSql = "select a0m02 from acc0l0,acc0m0,acc0k0,acc431,acc0j0,caseprogress" & _
            " where a0l02 between " & ChangeWStringToTString(CompDate(1, -2, strSrvDate(2))) & " and " & strSrvDate(2) & _
            " and a0l01=a0m01" & _
            " and a0m02=a0k01" & _
            " and a0k11='J' and a0k37='Y'" & _
            " and a0k01=axc02(+)" & _
            " and axc01 is null and a0k01=a0j13(+) and a0j01=cp09(+) and cp01||cp10<>'ACS706'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strData = ""
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If strData <> "" Then strData = strData & "，"
         strData = strData & RsTemp.Fields(0)
         RsTemp.MoveNext
      Loop
      Call SetDrdColData("請款單已收款未開發票", "未開立發票之請款單號" & strData)
   End If
   '****************************************************
   '智權人員繳款
   '****************************************************
   dblCnt = 0: dblCnt2 = 0
   strSql = "select count(*) from acc440 where a4416 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If RsTemp.Fields(0) > 0 Then
         dblCnt = RsTemp.Fields(0) '待收款筆數
         strSql = "select count(*) from acc440,staff where a4416 is null and a4401=st01(+) and st06<>'1' and a4413 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               dblCnt2 = RsTemp.Fields(0) '分所出納未確認筆數
            End If
         End If
         Call SetDrdColData("智權人員繳款", "尚有" & dblCnt & "筆資料待收款" & IIf(dblCnt2 > 0, "，其中包含" & dblCnt2 & "筆分所出納未確認", ""))
      End If
   End If
   Screen.MousePointer = vbDefault
   If DataGrid1.Rows = 2 And DataGrid1.TextMatrix(1, 0) = "" Then
      MsgBox "無待處理資料!!"
   End If
End Function

Private Sub SetDrdColData(strCol1 As String, strCol2 As String, Optional bolChangeColor As Boolean = False)
Dim i As Integer
   
   If DataGrid1.TextMatrix(DataGrid1.Rows - 1, 0) <> "" Then
      DataGrid1.AddItem ""
   End If
   DataGrid1.TextMatrix(DataGrid1.Rows - 1, 0) = strCol1
   DataGrid1.TextMatrix(DataGrid1.Rows - 1, 1) = strCol2
   If bolChangeColor = True Then
      DataGrid1.row = DataGrid1.Rows - 1
      For i = 0 To 1
         DataGrid1.col = i
         DataGrid1.CellBackColor = &H8080FF
      Next i
   End If
End Sub
