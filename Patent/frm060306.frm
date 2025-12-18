VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060306 
   BorderStyle     =   1  '單線固定
   Caption         =   "請款通知函"
   ClientHeight    =   5172
   ClientLeft      =   216
   ClientTop       =   960
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5172
   ScaleWidth      =   9312
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   6210
      Style           =   2  '單純下拉式
      TabIndex        =   6
      Top             =   4770
      Width           =   2250
   End
   Begin VB.TextBox Text5 
      Height          =   264
      Left            =   1116
      TabIndex        =   5
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   264
      Left            =   2550
      MaxLength       =   2
      TabIndex        =   3
      Top             =   390
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   264
      Left            =   2310
      MaxLength       =   1
      TabIndex        =   2
      Top             =   390
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Left            =   1470
      MaxLength       =   6
      TabIndex        =   1
      Top             =   390
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   990
      MaxLength       =   3
      TabIndex        =   0
      Top             =   390
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7740
      TabIndex        =   9
      Top             =   30
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8532
      TabIndex        =   10
      Top             =   12
      Width           =   756
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6948
      TabIndex        =   7
      Top             =   12
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3660
      Left            =   30
      TabIndex        =   8
      Top             =   1080
      Width           =   9270
      _ExtentX        =   16341
      _ExtentY        =   6456
      _Version        =   393216
      Cols            =   12
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   990
      TabIndex        =   4
      Top             =   690
      Width           =   8265
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "14579;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "設定地址條印表機 : "
      Height          =   180
      Index           =   1
      Left            =   4620
      TabIndex        =   14
      Top             =   4815
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "請款函日期："
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   13
      Top             =   4815
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   12
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   11
      Top             =   420
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   7710
      Y1              =   1035
      Y2              =   1035
   End
End
Attribute VB_Name = "frm060306"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/15 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim pa(1 To 7) As String
Dim CP10 As String
Dim intWhere As Integer
Dim intLastRow As Integer, intCols As Integer
'Add By Cheng 2003/01/28
Dim m_OriPrinterName As String, SeekPrint As Integer, SeekPrintL As Integer, j As Integer, i As Integer
'Add By Cheng 2003/01/30
Dim m_blnPrtGreenPaper As Boolean '判斷是否列印綠皮貼紙
'Add By Cheng 2003/02/07
Dim m_dblPrintLeft As Double '橫座偏移值
Dim m_dblPrintTop As Double '縱座偏移值
'Add by Morgan 2011/3/15
Dim strPrinter As String
Public m_quy416 As Boolean 'Add By Sindy 2017/3/20
Public m_quyNewCase As Boolean 'Add By Sindy 2022/5/12
Public m_quyAnyCP10 As String 'Add By Sindy 2022/5/17
Public m_FCna01 As String 'FC代理人國籍 Add By Sindy 2017/3/20


'Add By Sindy 2017/3/20
'Private Sub cmdok_Click(Index As Integer)
Public Sub cmdOK_Click(Index As Integer)
'2017/3/20 END
Dim i As Integer, bolChk As Boolean, strTmp As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim stPS As String 'Add by Morgan 2004/10/15 請款函備註預設
Dim stNo As String 'Add by Morgan 2004/12/8 代理人
Dim stNo1 As String 'Added by Morgan 2013/3/25
Dim stNo2 As String, stNo3 As String, stNo4 As String, stNo5 As String    'Added by Lydia 2023/08/01 申請人2~5

   Select Case Index
      Case 1 '確定
         With MSHFlexGrid1
            For i = 1 To .Rows - 1
               If .TextMatrix(i, 0) = "v" Then
                  bolChk = True
                  Me.Tag = .TextMatrix(i, 2)
                  CP10 = .TextMatrix(i, 9)
                  Exit For
               End If
            Next
         End With
         If bolChk = False Then
            MsgBox "請選擇資料 !", vbInformation
            Exit Sub
         End If

         'Add by Morgan 2004/10/15 代理人 Y20412 的所有請款函預設P.S. Our debit note will be separated from the rest of this letter and dealt with separately.
         stNo = GetPrjPeopleNum6(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))
         
         'Modified by Morgan 2013/3/25 +X28186的請款函加判斷代理人Y53495案件才帶備註
         'stPS = PUB_GetDNPS(stNo, pa(1) & pa(2) & pa(3) & pa(4), CP10)
         stNo1 = GetPrjPeopleNum1(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))
         'Added by Lydia 2023/08/01
         stNo2 = GetPrjPeopleNum2(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))
         stNo3 = GetPrjPeopleNum3(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))
         stNo4 = GetPrjPeopleNum4(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))
         stNo5 = GetPrjPeopleNum5(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))
         'end 2023/08/01
         
         'Add by Lydia 2015/02/06 改成表單維護, 共用模組PUB_GetDebitNotePS
'
'         stPS = PUB_GetDNPS(stNo, pa(1) & pa(2) & pa(3) & pa(4), CP10, stNo1)
'         'end 2013/3/25
'
'         'end 2004/101/5
'         'Add by Morgan 2008/11/13 申請人 X28186 的所有請款函也要預設
'         If stPS = "" Then
'            'Modified by Morgan 2013/3/25 申請人改用上面已設定的 stNo1 變數且無需再考慮是否多人申請(原X28186需求已不存在)
'            'stNo = GetPrjPeopleNum1(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))
'            'If GetPrjPeopleNum2(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)) = "" Then
'            '   stPS = PUB_GetDNPS(stNo, pa(1) & pa(2) & pa(3) & pa(4), CP10)
'            'End If
'            stPS = PUB_GetDNPS(stNo1, pa(1) & pa(2) & pa(3) & pa(4), CP10)
'         End If
'         'End 2008/11/13
         'Modified by Lydia 2023/08/01 增加申請人2~5
         'stPS = PUB_GetDebitNotePS(pa(1) & pa(2) & pa(3) & pa(4), CP10, stNo, stNo1)
         'end 2015/02/06
         stPS = PUB_GetDebitNotePS(pa(1) & pa(2) & pa(3) & pa(4), CP10, stNo, stNo1 & "," & stNo2 & "," & stNo3 & "," & stNo4 & "," & stNo5)

         Select Case CP10
            'Modify By Cheng 2003/02/25
            '翻譯(201), 檢視中說(209), 製作中說(210)
'            Case "201" '翻譯
            '93.4.8 modify by sonia 改為新申請案
            'Case "201", "209", "210"
            'Modify by Morgan 2004/10/6 加 "301", "302", "303", "304", "305", "306", "307"
            'Modified by Morgan 2013/4/15
            Case "101", "102", "103", "104", "105", "125", "301", "302", "303", "304", "305", "306", "307"
            '93.4.8 end
               'Add By Cheng 2003/01/30
               m_blnPrtGreenPaper = True
               frm060306_1.m_CP10 = CP10
               frm060306_1.Show
               frm060306_1.Text1(0).Text = stPS 'Add by Morgan 2004/10/15
               Me.Hide
            Case "601", "602", "603", "605" '領證/年費
               frm060306_2.Show
               frm060306_2.Text1(0).Text = stPS 'Add by Morgan 2004/10/15
               Me.Hide
            '讓與, 合併, 繼承, 授權, 變更, 延緩公告, 催審
            Case "701", "702", "703", "704", "401", "413", "411"
               frm060306_3.Show
               frm060306_3.Text1(0).Text = stPS 'Add by Morgan 2004/10/15
               Me.Hide
            'Modified by Lydia 2015/10/05 + 1008
            Case "1001", "1008" '核准
               'MSHFlexGrid1.TextMatrix(i, 7)
               frm060306_4.Show
               frm060306_4.Text1(0).Text = stPS 'Add by Morgan 2004/10/15
               Me.Hide
            Case "907", "913" '不續辦, 閉卷
               frm060306_5.Show
               frm060306_5.Text1(0).Text = stPS 'Add by Morgan 2004/10/15
               Me.Hide
            Case "1604" '專利權消滅
               frm060306_6.Label2(12).Caption = MSHFlexGrid1.TextMatrix(i, 6)
               frm060306_6.Show
               frm060306_6.Text1(0).Text = stPS 'Add by Morgan 2004/10/15
               Me.Hide
            'Modify by Morgan 2004/12/29 加 '加註專利權延長608','提早公開417'
            'Case "416" '實體審查
            'Modified by Morgan 2024/11/18 +447再審查加速審查
            Case "416", "608", "417", "447"
               frm060306_7.Show
               frm060306_7.Text1(0).Text = stPS 'Add by Morgan 2004/10/15
               Me.Hide
            'Add by Morgan 2004/12/29
            Case "425"
               frm060306_8.Show
               frm060306_8.Text1(0).Text = stPS
               Me.Hide
         End Select
'         Me.Hide
      Case 2 '結束
         Me.Enabled = False
         Unload Me
   End Select
End Sub

Public Sub Command1_Click()
Dim i As Integer
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
   MSHFlexGrid1.Clear
   If Text3 = "" Then Text3 = "0"
   If Text4 = "" Then Text4 = "00"
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   pub_QL05 = pub_QL05 & ";" & Label1(0) & Text1 & "-" & Text2 & "-" & Text3 & "-" & Text4 'Add By Sindy 2010/12/7
   If Text1 = "FCP" Then
      strExc(0) = "SELECT PA05,PA06,PA07,PA23 FROM PATENT WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   ElseIf Text1 = "FG" Then
      strExc(0) = "SELECT SP05,SP06,SP07 FROM SERVICEPRACTICE WHERE " & ChgService(pa(1) & pa(2) & pa(3) & pa(4))
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         For i = 0 To 2
            If IsNull(.Fields(i)) = False Then
               pa(i + 5) = .Fields(i)
            Else
               pa(i + 5) = ""
            End If
         Next
         AddCboName Combo1, pa(5), pa(6), pa(7)
      End With
   End If
   
   'Add By Sindy 2017/3/20
   'FC代理人國籍
   m_FCna01 = ""
   strExc(0) = "Select na01,na03" & _
               " From Patent,nation,fagent" & _
               " Where PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "'" & _
               " and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and fa10=na01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_FCna01 = Left(Trim("" & RsTemp.Fields("na01")), 3)
   End If
   
   'Modify by Morgan 2004/6/25 加603
   'Modify by Morgan 2004/10/6 加 '301','302','303','304','305','306','307'
   'Modify by Morgan 2004/12/29 加　'608','417','425'
   'Modified by Morgan 2013/4/15 + 125
   'Modified by Lydia 2015/10/05 + 1008
   strExc(0) = "SELECT ''," & SQLDate("CP05") & ",CP09,CPM03," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗')," & SQLDate("CP25") & ",CP43," & _
      "DECODE(CP10,'704',NVL(CP50,NVL(CP51,CP52)),'705',NVL(CP50,NVL(CP51,CP52)),'706',NVL(CP50,NVL(CP51,CP52)),'701',NVL(CU04,NVL(CU05,CU06)),NVL(CP40,NVL(CP41,CP42)))," & _
      "CP10 from caseprogress,casepropertymap,CUSTOMER where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4))
   'Add By Sindy 2017/3/20 只查實體審查
   If m_quy416 = True Then
      strExc(0) = strExc(0) & " and cp10='416'"
   'Add By Sindy 2022/5/12 只查新申請案
   ElseIf m_quyNewCase = True Then
      strExc(0) = strExc(0) & " and instr('" & NewCasePtyList & "',CP10)>0"
   'Add By Sindy 2022/5/17
   ElseIf m_quyAnyCP10 <> "" Then
      strExc(0) = strExc(0) & " and cp10='" & m_quyAnyCP10 & "'"
   Else
   '2017/3/20 END
      'Modified by Morgan 2024/11/18 +447再審查加速審查
      strExc(0) = strExc(0) & _
      " and cp10 IN ('101','102','103','104','105','125','601','602','603','605','701','702','704','401','413','411','1001','907','913','1604','416','301','302','303','304','305','306','307','608','417','425','1008','447')"
   End If
      strExc(0) = strExc(0) & _
      " and cp01=cpm01(+) and " & _
      "cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) order by cp05,cp09"
   '93.4.8 end
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2010/12/7
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
End Sub

Private Sub Form_Activate()
   'Modify By Sindy 2017/3/24 Mark
' Dim i As Integer
'   With MSHFlexGrid1
'      For i = 1 To .Rows - 1
'        .TextMatrix(i, 0) = ""
'      Next
'   End With

   'Modify By Sindy 2017/3/24
   'Me.Text2.SetFocus
   If Me.Text2.Visible = True And Me.Text2.Enabled = True Then Me.Text2.SetFocus
   '2017/3/24 END
End Sub

Private Sub Form_Load()
'Add By Cheng 2003/02/05
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer
    
    MoveFormToCenter Me
    intWhere = 國外_FC
    'Combo1.ListIndex = 0
    InitGrid 10, MSHFlexGrid1
    GridHead
    Text1 = "FCP"
    Text5 = GetTaiwanTodayDate
    
'Modify by Morgan 2011/3/15 改共用且不要排除預設印表機
   PUB_SetPrinter Me.Name, Combo2, strPrinter
'end 2011/3/15
End Sub

Private Sub Form_Unload(Cancel As Integer)

   'Copy from cmdok_Click by Morgan 2004/10/26
   '列印地址條
   PUB_PrintAddressList strUserNum, Me.Combo2.Text
   '刪除地址條列表資料
   PUB_DeleteAddressList strUserNum
   '初始化序號
   pub_AddressListSN = 0
   '若地址條印表機, 則更新列印設定
   If Me.Combo2.Text <> Me.Combo2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
   '2004/10/26 end
    
   Set frm060306 = Nothing
End Sub

Public Sub Clear()
   Text2 = Empty
   Text3 = Empty
   Text4 = Empty
   Combo1.Clear
   InitGrid 10, MSHFlexGrid1
   GridHead
   Me.Text2.SetFocus
   Command1.Default = True
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK(1).SetFocus
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 <> "FCP" And Text1 <> "FG" Then
      MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
      TextInverse Text1
      Cancel = True
   End If
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1200: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1400: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1200: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1400: .Text = "實際結果"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1400: .Text = "實際結果日期"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 1400: .Text = "相關總收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(8) = 1400: .Text = "相關人"
      For i = 9 To 9
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_LostFocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   If Text5 = "" Then
      MsgBox "請款函日期不得空白，請重新輸入 !", vbCritical
      Text5.SetFocus
   Else
      If CheckIsTaiwanDate(Text5, False) = False Then
         strTit = "檢核資料"
         strMsg = "請款函日期格式不正確"
         Text5.SetFocus
         TextInverse Text5
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
End Sub
