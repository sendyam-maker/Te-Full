VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071012 
   BorderStyle     =   1  '單線固定
   Caption         =   "開庭通知"
   ClientHeight    =   5925
   ClientLeft      =   105
   ClientTop       =   705
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9345
   Begin VB.TextBox txtcp01 
      Height          =   285
      Left            =   1152
      MaxLength       =   3
      TabIndex        =   0
      Top             =   372
      Width           =   550
   End
   Begin VB.TextBox txtcp02 
      Height          =   285
      Left            =   1776
      MaxLength       =   6
      TabIndex        =   1
      Top             =   372
      Width           =   855
   End
   Begin VB.TextBox txtcp03 
      Height          =   285
      Left            =   2688
      MaxLength       =   1
      TabIndex        =   2
      Top             =   372
      Width           =   255
   End
   Begin VB.TextBox txtcp04 
      Height          =   285
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   3
      Top             =   372
      Width           =   375
   End
   Begin VB.TextBox txtAccept 
      Height          =   285
      Left            =   1152
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1410
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Left            =   3528
      TabIndex        =   4
      Top             =   282
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8328
      TabIndex        =   8
      Top             =   70
      Width           =   770
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7536
      TabIndex        =   7
      Top             =   70
      Width           =   770
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4020
      Left            =   75
      TabIndex        =   6
      Top             =   1770
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   7091
      _Version        =   393216
      Cols            =   14
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
      _Band(0).Cols   =   14
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label lbeCusName 
      Height          =   285
      Left            =   2250
      TabIndex        =   16
      Top             =   1064
      Width           =   6375
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "11245;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   285
      Left            =   1152
      TabIndex        =   15
      Top             =   718
      Width           =   7485
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13203;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   252
      Left            =   192
      TabIndex        =   14
      Top             =   384
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   252
      Left            =   192
      TabIndex        =   13
      Top             =   731
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "當  事  人："
      Height          =   252
      Left            =   192
      TabIndex        =   12
      Top             =   1078
      Width           =   972
   End
   Begin VB.Label lbeNum 
      Height          =   285
      Left            =   1512
      TabIndex        =   11
      Top             =   384
      Width           =   1692
   End
   Begin VB.Label lbeCusNum 
      Height          =   285
      Left            =   1152
      TabIndex        =   10
      Top             =   1064
      Width           =   1035
   End
   Begin VB.Label Label21 
      Caption         =   "收  受  日："
      Height          =   252
      Left            =   192
      TabIndex        =   9
      Top             =   1426
      Width           =   972
   End
End
Attribute VB_Name = "frm071012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ; cboCaseName、lbeCusName、MSHFlexGrid1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim intLastRow As Integer, blnOKtoShow As Boolean, intCols As Integer
Dim blnCom1 As Boolean, blnCom2 As Boolean, blnCom3 As Boolean, blnCom4 As Boolean
Dim m_Lawcase As String

Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim LcTmp As String

   'Modify By Sindy 2011/7/6
   lbeCusNum.Caption = ""
   lbeCusName.Caption = ""
   cboCaseName.Clear
   cboCaseName.Text = ""
   MSHFlexGrid1.Clear
   MSHFlexGrid1.Rows = 2
   'End 2011/7/6
   If QueryDB = False Then
      MsgBox "本所案號不存在!", vbInformation, "開庭通知"
      txtcp02.SetFocus
      TextInverse txtcp02
'      MSHFlexGrid1.Clear
'      MSHFlexGrid1.Rows = 2
      m_Lawcase = False
      Exit Sub
   Else
      m_Lawcase = True
   End If

'   lbeCusNum.Caption = ""
'   lbeCusName.Caption = ""
'   cboCaseName.Clear
'   cboCaseName.Text = ""
'   MSHFlexGrid1.Clear
   If txtcp03 = "" Then txtcp03 = "0"
   If txtcp04 = "" Then txtcp04 = "00"
   LcTmp = txtcp01 + txtcp02 + txtcp03 + txtcp04
   'Modify By Sindy 2009/07/24 增加LIN系統類別
   'modify by sonia 2019/7/29 +ACS系統類別
   If txtcp01 = "L" Or txtcp01 = "FCL" Or txtcp01 = "CFL" Or txtcp01 = "LIN" Or txtcp01 = "ACS" Then
   
'      strExc(0) = "select ' ',SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2)," + _
'         "cp09,decode(lc15,020,cpm04,cpm03),decode(CP13," + _
'         "S1.ST01,S1.ST02),decode(CP14,S2.ST01,S2.ST02),decode(CP29,S3.ST01,S3.ST02)," + _
'         "decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2))," + _
'         "decode(cp71,or01,or02),lc05,lc06,lc07,lc11  from caseprogress, lawcase," + _
'         "STAFF S1,STAFF S2,STAFF S3, CASEPROPERTYMAP,organization where " & ChgCaseprogress(LcTmp) + " and " + _
'         ChgLawcase(LcTmp) + " AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND " + _
'         "cp01=cpm01(+) AND cp10=cpm02(+) and cp71=or01(+) and  substr(cp09,1,1)<>'C' and cp27 is not null"
      '91.11.10 MODIFY BY SONIA
      'strExc(0) = "select ' ',SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2)," + _
      '   "cp09,decode(lc15,020,cpm04,cpm03),decode(CP13," + _
      '   "S1.ST01,S1.ST02),decode(CP14,S2.ST01,S2.ST02),decode(CP29,S3.ST01,S3.ST02)," + _
      '   "decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2))," + _
      '   "decode(cp71,or01,or02),lc05,lc06,lc07,lc11  from caseprogress, lawcase," + _
      '   "STAFF S1,STAFF S2,STAFF S3, CASEPROPERTYMAP,organization where " & ChgCaseprogress(LcTmp) + " and " + _
      '   ChgLawcase(LcTmp) + " AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND " + _
      '   "cp01=cpm01(+) AND cp10=cpm02(+) and cp71=or01(+) and  cp09<'C' and cp27 is not null"
      strExc(0) = "select ' ',SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2)," + _
         "cp09,decode(lc15,020,cpm04,cpm03),decode(CP13," + _
         "S1.ST01,S1.ST02),decode(CP14,S2.ST01,S2.ST02),decode(CP29,S3.ST01,S3.ST02)," + _
         "decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2))," + _
         "decode(cp71,or01,or02),CP64,lc05,lc06,lc07,lc11  from caseprogress, lawcase," + _
         "STAFF S1,STAFF S2,STAFF S3, CASEPROPERTYMAP,organization where " & ChgCaseprogress(LcTmp) + " and " + _
         ChgLawcase(LcTmp) + " AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND " + _
         "cp01=cpm01(+) AND cp10=cpm02(+) and cp71=or01(+) and cp09<'C'"
      '91.11.10 END
   ElseIf txtcp01 = "LA" Then
      '91.11.10 MODIFY BY SONIA
      'strExc(0) = "select ' ',SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2)," + _
      '   "cp09,cpm03,decode(CP13,S1.ST01,S1.ST02),decode(CP14,S2.ST01,S2.ST02)," & _
      '   "decode(CP29,S3.ST01,S3.ST02)," + _
      '   "decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2))," + _
      '   "decode(cp71,or01,or02),hc05,hc06 from caseprogress, hirecase,STAFF S1,STAFF S2,STAFF S3, " + _
      '   "CASEPROPERTYMAP,organization where " & ChgCaseprogress(LcTmp) + " and " & ChgHirecase(LcTmp) + _
      '   " AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) and cp71=or01(+) " + _
      '   "and CP09<'C' and cp27 is not null order by cp05"
      strExc(0) = "select ' ',SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2)," + _
         "cp09,cpm03,decode(CP13,S1.ST01,S1.ST02),decode(CP14,S2.ST01,S2.ST02)," & _
         "decode(CP29,S3.ST01,S3.ST02)," + _
         "decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2))," + _
         "decode(cp71,or01,or02),CP64,hc05,hc06 from caseprogress, hirecase,STAFF S1,STAFF S2,STAFF S3, " + _
         "CASEPROPERTYMAP,organization where " & ChgCaseprogress(LcTmp) + " and " & ChgHirecase(LcTmp) + _
         " AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) and cp71=or01(+) " + _
         "and CP09<'C'"
      '91.11.10 END
   End If
   '2011/5/4 modify by sonia 取消發文日限制,因為有可能為被告案件未發文即有來函
   'strExc(0) = strExc(0) & " AND CP27 IS NOT NULL AND CP10<>'0' Order by cp27 DESC,CP09" '2009/9/9 ADD BY SONIA 加剔除未發文,顧問聘任及排序條件
   strExc(0) = strExc(0) & " AND CP10<>'0' Order by cp27 DESC,CP09" '2009/9/9 ADD BY SONIA 加剔除未發文,顧問聘任及排序條件
   intI = 0
   'edit by nickc 2007/02/07 不用 dll 了
   'Set rsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      PutDataInGrid
      txtAccept.SetFocus
   Else
      lbeCusNum = ""
      lbeCusName = ""
      cboCaseName.Clear
      MSHFlexGrid1.Rows = 2
'      cmdSure.Enabled = False
      txtcp02_GotFocus
   End If
   GridHead
  ' cmdSearch.Enabled = False
   blnCom1 = False
   blnCom2 = False
   blnCom3 = False
   blnCom4 = False
End Sub

Private Sub cmdSure_Click()
 Dim i As Integer, blnChoese As Boolean
   If txtcp01.Text = "" Or txtcp02.Text = "" Then
      txtcp01.SetFocus
      TextInverse txtcp01
      MsgBox "本所案號不可空白!", vbExclamation, "開庭通知"
      Exit Sub
   End If
   
   blnChoese = False
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         .row = i
         .col = 0
         If .Text = "v" Then
            blnChoese = True
            Exit For
         End If
      Next
   End With
   If Not blnChoese Then
      MsgBox "請點選輸入資料", vbCritical
'      cmdSure.Enabled = False
      MSHFlexGrid1.SetFocus
      Exit Sub
   End If
   If txtAccept = "" Then
      txtAccept.SetFocus
      TextInverse txtAccept
      MsgBox "收受日不可空白!", vbCritical
      Exit Sub
   'Else
   '   strExc(0) = txtcp01 & txtcp02 & txtcp03 & txtcp04
   '   If Not objLawDll.ChkMRec(ChangeTStringToWString(txtAccept), strExc(0), strExc(1), strExc(2)) Then
   '   If Right(strExc(0), 3) = "000" Then strExc(0) = Left(strExc(0), Len(strExc(0)) - 3)
   '      MsgBox "本所案號'" + strExc(0) + "'與收受日'" + txtAccept + "'不存在於來函記錄檔中", vbCritical
   '      Exit Sub
   '   End If
   End If

   frm071013.Show
   'cmdSearch.SetFocus
   Me.Hide
End Sub

Private Sub Form_Activate()
   txtcp01.SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
  ' cmdSearch.Enabled = False
'   cmdSure.Enabled = False
   blnCom1 = False
   blnCom2 = False
   blnCom3 = False
   blnCom4 = False
   'txtAccept.Text = Format(Date, "EE") - 1911 & Format(Date, "MM") & Format(Date, "DD")
   txtAccept.Text = ChangeWStringToTString(GetTodayDate)

End Sub

Private Sub GridHead()
   With MSHFlexGrid1
      blnOKtoShow = False
      .Visible = False
      .Cols = 14
      .row = 0
      .col = 0
      .Visible = True
      .MergeCells = flexMergeRestrictRows
      .MergeRow(0) = True
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .col = 1: .ColWidth(1) = 900: .Text = "收文日"
      .col = 2: .ColWidth(2) = 1000: .Text = "收文號"
      .col = 3: .ColWidth(3) = 1000: .Text = "案件性質"
      .col = 4: .ColWidth(4) = 900: .Text = "智權人員"
      .col = 5: .ColWidth(5) = 900: .Text = "承辦人"
      'Modified by Lydia 2015/10/05
      '.col = 6: .ColWidth(6) = 900: .Text = "法務人員"
      .col = 6: .ColWidth(6) = 900: .Text = "協辦人員"
      .col = 7: .ColWidth(7) = 900: .Text = "發文日"
      .col = 8: .ColWidth(8) = 1200: .Text = "法院"
      .col = 9: .ColWidth(9) = 1500: .Text = "進度備註"
      .col = 10: .ColWidth(10) = 0
      .col = 11: .ColWidth(11) = 0
      .col = 12: .ColWidth(12) = 0
      .col = 13: .ColWidth(13) = 0
      .CellAlignment = flexAlignCenterCenter
      intLastRow = 0
      blnOKtoShow = True
      '判斷是否有資料
      If .Rows > 1 Then .row = 1
      .Visible = True
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm071012 = Nothing
End Sub

Private Sub lbeCusNum_Change()
 Dim StrCusName As String
   If Len(lbeCusNum) > 8 Then
     'edit by nickc 2007/02/07 不用 dll 了
     'If objPublicData.GetCustomer(lbeCusNum, StrCusName) Then lbeCusName = StrCusName
     If ClsPDGetCustomer(lbeCusNum, StrCusName) Then lbeCusName = StrCusName
   End If
End Sub

Private Sub MSHFlexGrid1_Click()
'   intCols = MSHFlexGrid1.Cols - 1
'   If Not CheckGridChoese(MSHFlexGrid1, intLastRow, intCols) Then Exit Sub
'   If txtAccept <> "" Then
'      cmdSure.Enabled = True
'      cmdSure.SetFocus
'   End If
 Dim i As Integer
   With MSHFlexGrid1
   intCols = MSHFlexGrid1.Cols - 1
   ShowBar MSHFlexGrid1, intLastRow, intCols

      'intClkRow = .Row
      .col = 0
      ClearGrid
      .row = intLastRow
      If .Text = "v" Then
         .Text = ""
      Else
         .Text = "v"
         cmdSure.SetFocus
      End If
   End With

End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then MSHFlexGrid1_Click
End Sub

Private Sub txtAccept_GotFocus()
   TextInverse txtAccept
End Sub

Private Sub txtAccept_Validate(Cancel As Boolean)
If txtAccept <> "" Then
If CheckIsTaiwanDate(txtAccept) Then
   If Val(GetTaiwanTodayDate) - Val(txtAccept) < 0 Then
       MsgBox "輸入日期大於系統日", vbCritical
       Cancel = True
   Else
       If MSHFlexGrid1.Text <> "" Then
          cmdSure.Enabled = True
          cmdSure.SetFocus
        End If
    End If
Else
   Cancel = True
End If
End If
If Cancel Then TextInverse txtAccept

End Sub

Private Sub txtcp01_GotFocus()
   TextInverse txtcp01
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
  Dim strTit As String
  Dim strMsg As String
  
   txtcp01 = UCase(txtcp01)
   If IsEmptyText(txtcp01) = False Then
      blnCom1 = True
      '2011/5/20 MODIFY BY SONIA
      'If Not IsCorrectSysKindLaw(txtcp01) Then
      If CheckSys(txtcp01) <> "3" And CheckSys(txtcp01) <> "4" Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "系統類別不正確"
         MsgBox strMsg, vbOKOnly, strTit
         Cancel = True
         blnCom1 = False
         TextInverse txtcp01
         Exit Sub
      End If
      
      ' 檢查使用者是否有使用該系統類別的權限
      If IsUserHasRightOfSystem(strUserNum, txtcp01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有使有此系統別的權限"
         MsgBox strMsg, vbOKOnly, strTit
         Cancel = True
         blnCom1 = False
         TextInverse txtcp01
         Exit Sub
      End If
   End If

'   If txtcp01 <> "" Then
'      txtcp01 = UCase(txtcp01)
'      If txtcp01 = "L" Or txtcp01 = "LA" Or txtcp01 = "FCL" Then
'         blnCom1 = True
'      Else
'         DataErrorMessage 1, "系統類別"
'         blnCom1 = False
'         Cancel = True
'      End If
'   End If
'   If Cancel Then TextInverse txtcp01
End Sub

Private Sub txtcp02_GotFocus()
   TextInverse txtcp02
End Sub

Private Sub txtcp02_Validate(Cancel As Boolean)
   If txtcp02 <> "" Then
      blnCom2 = True
      CmdSearch.Enabled = True
   End If
   If Cancel Then TextInverse txtcp02
End Sub
Private Sub ChkCmd()
   If txtcp03 = "" Then blnCom3 = True: blnCom4 = True
   If blnCom1 And blnCom2 And blnCom3 And blnCom4 Then
      CmdSearch.Enabled = True
      CmdSearch.SetFocus
   End If
End Sub

Private Sub txtcp03_GotFocus()
   TextInverse txtcp03
End Sub

Private Sub txtcp03_Validate(Cancel As Boolean)
   If txtcp03 <> "" Then
      blnCom3 = True
      ChkCmd
   End If
   If Cancel Then TextInverse txtcp03
End Sub

Private Sub PutDataInGrid()
 Dim i As Integer, strTempName As String, strCus As String
   With MSHFlexGrid1
      .Visible = False
      If Not (RsTemp.EOF And RsTemp.BOF) Then
         If txtcp01 <> "LA" Then
            strCus = "" & RsTemp.Fields!LC11
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetCustomer(strCus, strTempName) Then lbeCusNum = strCus: lbeCusName = strTempName
            '2011/7/4 MODIFY BY SONIA 有當事人才抓
            'If ClsPDGetCustomer(strCus, strTempName) Then
            If strCus <> "" Then
               If ClsPDGetCustomer(strCus, strTempName) Then lbeCusNum = strCus: lbeCusName = strTempName
            End If
            
            If Not IsNull(RsTemp.Fields!lc05) Then cboCaseName.AddItem "中:" + RsTemp.Fields!lc05
            If Not IsNull(RsTemp.Fields!lc06) Then cboCaseName.AddItem "英:" + RsTemp.Fields!lc06
            If Not IsNull(RsTemp.Fields!lc07) Then cboCaseName.AddItem "日:" + RsTemp.Fields!lc07
         Else
            strCus = RsTemp.Fields!hc05
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetCustomer(strCus, strTempName) Then lbeCusNum = strCus: lbeCusName = strTempName
            If ClsPDGetCustomer(strCus, strTempName) Then lbeCusNum = strCus: lbeCusName = strTempName
            If Not IsNull(RsTemp.Fields!hc06) Then cboCaseName.AddItem "中:" + RsTemp.Fields!hc06
         End If
         cboCaseName.ListIndex = 0
         Set .Recordset = RsTemp
      End If
      .Visible = True
   End With
End Sub

Private Sub txtcp04_GotFocus()
   TextInverse txtcp04
End Sub

Private Sub txtcp04_Validate(Cancel As Boolean)
   If txtcp04 <> "" Then
      blnCom4 = True
      ChkCmd
   End If
   If Cancel Then TextInverse txtcp04 Else ChkCmd
End Sub
Private Function QueryDB() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   Dim bQuery As Boolean
   strTM01 = txtcp01.Text
   strTM02 = txtcp02.Text
   If txtcp03.Text <> Empty Then
      strTM03 = txtcp03.Text
   Else
      strTM03 = "0"
   End If
   If txtcp04.Text <> Empty Then
      strTM04 = txtcp04.Text
   Else
      strTM04 = "00"
   End If
   
   ' 依本所案號讀取基本檔案
   
   Select Case UCase(txtcp01.Text)
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         bQuery = QueryTradeMark(strTM01, strTM02, strTM03, strTM04)
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         bQuery = QueryPatent(strTM01, strTM02, strTM03, strTM04)
      ' 讀取法務基本檔
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS":
         bQuery = QueryLawCase(strTM01, strTM02, strTM03, strTM04)
      ' 讀取顧問案件基本檔
      Case "LA":
         bQuery = QueryHireCase(strTM01, strTM02, strTM03, strTM04)
      ' 讀取服務業務基本檔
      Case Else:
         bQuery = QueryServicePractice(strTM01, strTM02, strTM03, strTM04)
   End Select
   QueryDB = bQuery
End Function
' 讀取商標基本檔
Private Function QueryTradeMark(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryTradeMark = False
   strSql = "SELECT * FROM TRADEMARK " & _
            "WHERE TM01 = '" & strTM01 & "' AND " & _
                  "TM02 = '" & strTM02 & "' AND " & _
                  "TM03 = '" & strTM03 & "' AND " & _
                  "TM04 = '" & strTM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryTradeMark = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cboCaseName.AddItem "中 : " & rsTmp.Fields("TM05")
      End If
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cboCaseName.AddItem "英 : " & rsTmp.Fields("TM06")
      End If
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cboCaseName.AddItem "日 : " & rsTmp.Fields("TM07")
      End If
      ' 顯示商標名稱
      If cboCaseName.ListCount > 0 Then
         cboCaseName.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         lbeCusName.Caption = GetCustomerName(rsTmp.Fields("TM23"), 0)
         lbeCusNum.Caption = rsTmp.Fields("TM23")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取服務業務基本檔
Private Function QueryServicePractice(ByVal strSP01 As String, ByVal strSP02 As String, ByVal strSP03 As String, ByVal strSP04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryServicePractice = False
   strSql = "SELECT * FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & strSP01 & "' AND " & _
                  "SP02 = '" & strSP02 & "' AND " & _
                  "SP03 = '" & strSP03 & "' AND " & _
                  "SP04 = '" & strSP04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryServicePractice = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then
         cboCaseName.AddItem "中 : " & rsTmp.Fields("SP05")
      End If
      If IsNull(rsTmp.Fields("SP06")) = False Then
         cboCaseName.AddItem "英 : " & rsTmp.Fields("SP06")
      End If
      If IsNull(rsTmp.Fields("SP07")) = False Then
         cboCaseName.AddItem "日 : " & rsTmp.Fields("SP07")
      End If
      ' 顯示商標名稱
      If cboCaseName.ListCount > 0 Then
         cboCaseName.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         lbeCusName.Caption = GetCustomerName(rsTmp.Fields("SP08"), 0)
         lbeCusNum.Caption = rsTmp.Fields("SP08")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取專利基本檔
Private Function QueryPatent(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryPatent = False
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & strPA01 & "' AND " & _
                  "PA02 = '" & strPA02 & "' AND " & _
                  "PA03 = '" & strPA03 & "' AND " & _
                  "PA04 = '" & strPA04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryPatent = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("PA05")) = False Then
         cboCaseName.AddItem "中 : " & rsTmp.Fields("PA05")
      End If
      If IsNull(rsTmp.Fields("PA06")) = False Then
         cboCaseName.AddItem "英 : " & rsTmp.Fields("PA06")
      End If
      If IsNull(rsTmp.Fields("PA07")) = False Then
         cboCaseName.AddItem "日 : " & rsTmp.Fields("PA07")
      End If
      ' 顯示商標名稱
      If cboCaseName.ListCount > 0 Then
         cboCaseName.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("PA26")) = False Then
         lbeCusName.Caption = GetCustomerName(rsTmp.Fields("PA26"), 0)
         lbeCusNum.Caption = rsTmp.Fields("PA26")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取法務基本檔
Private Function QueryLawCase(ByVal strLC01 As String, ByVal strLC02 As String, ByVal strLC03 As String, ByVal strLC04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryLawCase = False
   strSql = "SELECT * FROM LAWCASE " & _
            "WHERE LC01 = '" & strLC01 & "' AND " & _
                  "LC02 = '" & strLC02 & "' AND " & _
                  "LC03 = '" & strLC03 & "' AND " & _
                  "LC04 = '" & strLC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryLawCase = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("LC05")) = False Then
         cboCaseName.AddItem "中 : " & rsTmp.Fields("LC05")
      End If
      If IsNull(rsTmp.Fields("LC06")) = False Then
         cboCaseName.AddItem "英 : " & rsTmp.Fields("LC06")
      End If
      If IsNull(rsTmp.Fields("LC07")) = False Then
         cboCaseName.AddItem "日 : " & rsTmp.Fields("LC07")
      End If
      ' 顯示商標名稱
      If cboCaseName.ListCount > 0 Then
         cboCaseName.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("LC11")) = False Then
         lbeCusName.Caption = GetCustomerName(rsTmp.Fields("LC11"), 0)
         lbeCusNum.Caption = rsTmp.Fields("LC11")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取顧問案件基本檔
Private Function QueryHireCase(ByVal strHC01 As String, ByVal strHC02 As String, ByVal strHC03 As String, ByVal strHC04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryHireCase = False
   strSql = "SELECT * FROM HIRECASE " & _
            "WHERE HC01 = '" & strHC01 & "' AND " & _
                  "HC02 = '" & strHC02 & "' AND " & _
                  "HC03 = '" & strHC03 & "' AND " & _
                  "HC04 = '" & strHC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryHireCase = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("HC06")) = False Then
         cboCaseName.AddItem rsTmp.Fields("HC06")
      End If
      ' 顯示商標名稱
      If cboCaseName.ListCount > 0 Then
         cboCaseName.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("HC05")) = False Then
         lbeCusName.Caption = GetCustomerName(rsTmp.Fields("HC05"), 0)
         lbeCusNum.Caption = rsTmp.Fields("HC05")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function


Private Sub ClearGrid()
 Dim i As Integer
   With MSHFlexGrid1
      .Visible = False
      For i = 1 To .Rows - 1
         .col = 0
         .row = i
         .Text = ""
      Next
      .Visible = True
   End With
End Sub

