VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090801_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "申請人資料查詢"
   ClientHeight    =   5870
   ClientLeft      =   2800
   ClientTop       =   3720
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5870
   ScaleWidth      =   8950
   Begin VB.CommandButton CmdRelation 
      Caption         =   "關係企業(&R)"
      Height          =   400
      Left            =   5280
      TabIndex        =   4
      Top             =   45
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6682
      TabIndex        =   0
      Top             =   45
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7680
      TabIndex        =   1
      Top             =   45
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5340
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   8895
      _ExtentX        =   15681
      _ExtentY        =   9419
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
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
      _Band(0).Cols   =   16
   End
   Begin MSForms.Label lblName 
      Height          =   255
      Left            =   1140
      TabIndex        =   5
      Top             =   180
      Width           =   5145
      VariousPropertyBits=   27
      Size            =   "9075;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人中文："
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   210
      Width           =   1080
   End
End
Attribute VB_Name = "frm090801_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/14 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB、lblName
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
Dim strSql As String, i As Integer, j As Integer, strTemp As Variant, strTemp1 As Variant, s As Integer
Dim StrTag As String, intK As Integer
'Add By Sindy 2010/4/30
Public m_strCustCode As String
Public m_blnOneRec As Boolean
Public m_strCustChnName As String
Public m_Type As Integer '0.申請人(中文) 1.代理人(中文) 2.申請人(英文) 3.發明人(Add by Lydia 2014/9/22)
                         '4.收據抬頭 Add by Sindy 2023/8/7
Public m_DouChk As Boolean  'Add by Lydia 2014/9/22 判斷是否可多選
Public m_frm0908A As Form 'Add by Lydia 2014/9/22 傳回各表單發明人資料
'2010/4/30 End
Public m_Lang As String 'Added by Morgan 2020/7/2


Private Sub SetDataListWidth()
Me.grdDataList.Cols = 6
Me.grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 200
grdDataList.CellAlignment = flexAlignCenterCenter

If m_Type = 3 Then 'add by Lydia 2014/9/22
   grdDataList.col = 1: grdDataList.Text = "發明人編號"
   grdDataList.ColWidth(1) = 1200
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 2: grdDataList.Text = "申請人ID"
   grdDataList.ColWidth(2) = 1000
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 3: grdDataList.Text = "發明人"
   grdDataList.ColWidth(3) = 1000
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 4: grdDataList.Text = "地址"
   grdDataList.ColWidth(4) = 2600
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 5: grdDataList.Text = "國籍"
   grdDataList.ColWidth(5) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   
   grdDataList.ColWidth(6) = 0 '英文姓名
   grdDataList.ColWidth(7) = 0 '英文地址

'Add By Sindy 2023/8/7 收據抬頭
ElseIf m_Type = 4 Then
   Me.grdDataList.Cols = 10
   grdDataList.col = 1: grdDataList.Text = "智權人員"
   grdDataList.ColWidth(1) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 2: grdDataList.Text = "收據抬頭"
   grdDataList.ColWidth(2) = 1200
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 3: grdDataList.Text = "統一編號"
   grdDataList.ColWidth(3) = 1000
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 4: grdDataList.Text = "營業地址"
   grdDataList.ColWidth(4) = 1500
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 5: grdDataList.Text = "郵寄地址"
   grdDataList.ColWidth(5) = 1500
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 6: grdDataList.Text = "電話"
   grdDataList.ColWidth(6) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 7: grdDataList.Text = "傳真"
   grdDataList.ColWidth(7) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 8: grdDataList.Text = "財務Mail"
   grdDataList.ColWidth(8) = 1500
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 9: grdDataList.Text = "sort"
   grdDataList.ColWidth(9) = 0
   grdDataList.CellAlignment = flexAlignCenterCenter
   '2023/8/7 END
   
Else
   'Modified by Morgan 2020/7/20 智權人員移到第1欄--政興,秀玲
   grdDataList.col = 1: grdDataList.Text = "智權人員"
   grdDataList.ColWidth(1) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 2: grdDataList.Text = "代號"
   grdDataList.ColWidth(2) = 1000 'Modify by Amy 2016/08/05 原:0 因Y45776000/010 北京金信有兩筆名稱相同,編號做不同使用
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 3: grdDataList.Text = "申請人名稱"
   grdDataList.ColWidth(3) = 2600
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 4: grdDataList.Text = "聯絡地址"
   grdDataList.ColWidth(4) = 4600
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 5: grdDataList.Text = "接洽人"
   grdDataList.ColWidth(5) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
End If

End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim ii As Integer

Dim m3Name As String, m3Addr As String, m3Na As String, m3ID As String  'Add by Lydia 2014/9/22 傳回各表單發明人資料
Dim m3NameE As String, m3AddrE As String

m_strCustCode = ""
Select Case Index
Case 0 '確定
    For ii = 1 To Me.grdDataList.Rows - 1
        If Me.grdDataList.TextMatrix(ii, 0) = "V" Then
            m_blnOneRec = True 'Add By Sindy 2010/4/30
            
          'Add by Lydia 2014/9/22 判斷是否可多選
          'Modified by Morgan 2020/7/20 智權人員移到第1欄--政興,秀玲
            If m_DouChk = True Then
              If m_Type <> 3 Then
                  
                  'Added by Morgan 2020/7/20
                  If m_Type = 0 Or m_Type = 2 Then
                     If Len(m_strCustCode) > 0 Then
                        m_strCustCode = m_strCustCode + "," + LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 2)))
                     Else
                       m_strCustCode = Me.grdDataList.TextMatrix(ii, 2)
                     End If
                  Else
                  'end 2020/7/20
                     If Len(m_strCustCode) > 0 Then
                        m_strCustCode = m_strCustCode + "," + LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 1)))
                     Else
                       m_strCustCode = Me.grdDataList.TextMatrix(ii, 1)
                     End If
                  End If 'Added by Morgan 2020/7/20
              Else
                 If Len(m3Name) > 0 Then
                    m_strCustCode = m_strCustCode + "," + LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 1)))
                    m3Name = m3Name + "、" + LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 3)))
                    If Len(LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 6)))) = 0 Then
                    Else
                     m3NameE = m3NameE + ", " + LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 6)))
                    End If
                    If Len(LTrim(RTrim(m3Addr))) = 0 Then
                     m3Addr = "" & LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 4)))
                    End If
                 Else
                    m_strCustCode = Me.grdDataList.TextMatrix(ii, 1)
                    m3Name = LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 3)))
                    m3ID = LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 2)))
                    m3Addr = "" & LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 4)))
                    m3Na = "" & LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 5)))
                    m3NameE = "" & LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 6))) '英文姓名
                    m3AddrE = "" & LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 7))) '英文地址
                 End If
              End If
              
            Else
               'Modified by Morgan 2020/7/20
               'm_strCustCode = Me.grdDataList.TextMatrix(ii, 1)
               If m_Type = 0 Or m_Type = 2 Then
                  m_strCustCode = Me.grdDataList.TextMatrix(ii, 2)
               'Add by Sindy 2023/8/7
               ElseIf m_Type = 4 Then
                  m_strCustCode = ii
                  '2023/8/7 END
               Else
                  m_strCustCode = Me.grdDataList.TextMatrix(ii, 1)
               End If
               'end 2020/7/20
               
               If m_Type = 3 Then
                  m3Name = LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 3)))
                  m3ID = LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 2)))
                  m3Addr = "" & LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 4)))
                  m3Na = "" & LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 5)))
                  m3NameE = "" & LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 6))) '英文姓名
                  m3AddrE = "" & LTrim(RTrim(Me.grdDataList.TextMatrix(ii, 7))) '英文地址
                
                'Add By Sindy 2023/8/7
                ElseIf m_Type = 4 Then
                  m_strCustChnName = grdDataList.TextMatrix(ii, 2)
                  '2023/8/7 END
                  
                'Added by Morgan 2020/7/2
                ElseIf m_Type = 1 Then
                  m_Lang = grdDataList.TextMatrix(ii, 6)
                  m_strCustChnName = grdDataList.TextMatrix(ii, 2)
                'end 2020/7/2
                End If
                Exit For
            End If
        End If
    Next ii
    
    
    'If ii > Me.GrdDataList.Rows - 1 Then  modify by Lidia 2014/9/22
    If Len(m_strCustCode) = 0 Then
      If m_Type = 0 Or m_Type = 2 Then
         MsgBox "請勾選一筆申請人!!!", vbExclamation + vbOKOnly
      Else
        If m_Type = 3 Then
         MsgBox "請勾選一筆發明人!!!", vbExclamation + vbOKOnly
        'Add By Sindy 2023/8/7
        ElseIf m_Type = 4 Then
         MsgBox "請勾選一筆收據抬頭!!!", vbExclamation + vbOKOnly
         '2023/8/7 END
        Else
         '1.代理人(中文)
         MsgBox "請勾選一筆代理人!!!", vbExclamation + vbOKOnly
        End If
      End If
      Exit Sub
    Else
      If m_Type = 3 Then '傳回各表單發明人資料
          If m_frm0908A.Name = "frm210114_1" Then
            m_frm0908A.Txt1(1) = m3Name
            m_frm0908A.Txt1(2) = m3ID
            m_frm0908A.Txt1(3) = m3Addr
            m_frm0908A.Txt1(4) = m3Na
          Else
            m_frm0908A.Txt1(1) = m3Name
            m_frm0908A.Txt1(2) = m3NameE
            m_frm0908A.Txt1(3) = m3Addr
            m_frm0908A.Txt1(4) = m3AddrE
          End If
      End If
    End If
    'Unload Me
    Me.Hide
Case 1 '回前畫面
    Unload Me
Case Else
End Select
End Sub

Private Sub CmdRelation_Click()
'Add by Lydia 2014/9/22 關係企業資料
Dim iX As Integer, m_Xe As String, m_Xe2 As String

'Modified by Morgan 2020/7/20 申請人-智權人員移到第1欄--政興,秀玲
For iX = 1 To Me.grdDataList.Rows - 1
    If Me.grdDataList.TextMatrix(iX, 0) = "V" Then
        If Len(m_Xe) > 0 Then
            If m_Type = 3 Then
               m_Xe = m_Xe + ",'" + Mid(LTrim(RTrim(Me.grdDataList.TextMatrix(iX, 1))), 1, 6) + "'"
            Else
               m_Xe = m_Xe + ",'" + Mid(LTrim(RTrim(Me.grdDataList.TextMatrix(iX, 2))), 1, 6) + "'"
            End If
        Else
            If m_Type = 3 Then
               m_Xe = "'" + Mid(Me.grdDataList.TextMatrix(iX, 1), 1, 6) + "'"
            Else
               m_Xe = "'" + Mid(Me.grdDataList.TextMatrix(iX, 2), 1, 6) + "'"
            End If
        End If
        
        If Len(m_Xe2) > 0 Then
           m_Xe2 = m_Xe2 + ", "
        End If
        If m_Type = 3 Then
           m_Xe2 = m_Xe2 + Trim(Me.grdDataList.TextMatrix(iX, 3))
        Else
           m_Xe2 = m_Xe2 + Trim(Me.grdDataList.TextMatrix(iX, 3))
        End If
        
    End If
         
Next iX
    

    If Len(m_Xe) = 0 Then
      If m_Type = 0 Or m_Type = 2 Then
         MsgBox "請勾選一筆申請人!!!", vbExclamation + vbOKOnly
      Else
        If m_Type = 3 Then
         MsgBox "請勾選一筆發明人!!!", vbExclamation + vbOKOnly
        'Add By Sindy 2023/8/7
        ElseIf m_Type = 4 Then
         MsgBox "請勾選一筆收據抬頭!!!", vbExclamation + vbOKOnly
         '2023/8/7 END
        Else
         '1.代理人(中文)
         MsgBox "請勾選一筆代理人!!!", vbExclamation + vbOKOnly
        End If
      End If
      Exit Sub
    End If
    
If m_Type = 0 Then
   strSql = "Select '' As V, ST02 As 智權人員, CU01||CU02 As 代號, CU04 As 申請人名稱, CU31 As 聯絡地址, nvl(pcc05,CU08) As 接洽人, CU79 AS 備註 From Customer, Staff, potcustcont Where CU02='0' and CU13=ST01(+) " & _
            " And substr(CU01,1,6) in (" & m_Xe & ") and pcc01(+)=cu01 and pcc02(+)=cu127 Order By CU04, CU01 "

ElseIf m_Type = 2 Then
   strSql = "Select '' As V, ST02 As 智權人員, CU01||CU02 As 代號, rtrim(CU05||' '||CU88||' '||CU89||' '||CU90) As 申請人名稱, CU31 As 聯絡地址, nvl(pcc05,CU08) As 接洽人, CU79 AS 備註 From Customer, Staff, potcustcont Where CU02='0' and CU13=ST01(+) " & _
            "And  substr(CU01,1,6) in (" & m_Xe & ") and pcc01(+)=cu01 and pcc02(+)=cu127 Order By CU04, CU01 "

ElseIf m_Type = 3 Then
   'Modified by Lydia 2019/08/05 台灣地區改顯示為中華民國
   'strSql = " SELECT ' ' as V,IN01||'-'||IN02 AS 發明人編號,IN03 AS 發明人ID,IN04 AS 發明人名稱,IN07 As 地址,NA03 AS 國籍 " & _
            " ,IN05 as 英文名,IN08 as 英文地址 " & _
            " FROM INVENTOR,NATION WHERE IN11=NA01(+) AND " & _
            " substr(IN01,1,6) in (" & m_Xe & ") order by IN03,IN02,IN04 "
   strSql = " SELECT ' ' as V,IN01||'-'||IN02 AS 發明人編號,IN03 AS 發明人ID,IN04 AS 發明人名稱,IN07 As 地址,decode(sign((to_number(substr(na01,1,3)) - 10)),1,na03,'中華民國') AS 國籍 " & _
            " ,IN05 as 英文名,IN08 as 英文地址 " & _
            " FROM INVENTOR,NATION WHERE IN11=NA01(+) AND " & _
            " substr(IN01,1,6) in (" & m_Xe & ") order by IN03,IN02,IN04 "
End If

lblName.Caption = m_Xe2 + "(含關係企業）"

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then

Else
    ShowNoData
    Screen.MousePointer = vbDefault
    m_strCustCode = ""

End If
Set grdDataList.Recordset = adoRecordset
CheckOC
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
   'Add by Lydia 2014/9/22
   If m_DouChk = True Then
      CmdRelation.Visible = True
   Else
      CmdRelation.Visible = False
   End If
End Sub

'Add By Sindy 2023/8/7
Public Sub StrMenu2(AdoRs As ADODB.Recordset)
m_blnOneRec = False
m_strCustCode = ""
Set grdDataList.Recordset = AdoRs
End Sub

Function StrMenu() As Boolean
m_blnOneRec = False
m_strCustCode = "" 'Add By Sindy 2010/4/30
StrMenu = False
Screen.MousePointer = vbHourglass
'Modify by Morgan 2008/7/24 接洽人改先用聯絡人編號抓聯絡人檔
'strSQL = "Select '' As V, CU01||CU02 As 代號, CU04 As 申請人名稱, CU31 As 聯絡地址, CU08 As 接洽人, ST02 As 智權人員 From Customer, Staff Where CU13=ST01(+) And CU04 Like '" & ChgSQL(frm090801.m_strCustChnName) & "%' Order By CU04, CU01 "
'Modify By Sindy 2010/6/28 where條件增加判斷CU02='0'及FA02='0'
'Modified by Morgan 2020/7/20 申請人-智權人員移到第1欄--政興,秀玲
If m_Type = 0 Then
   'Modified by Lydia 2019/02/15 +接洽單查詢申請人,增加CU80狀態
   strSql = "Select '' As V, ST02 As 智權人員, CU01||CU02 As 代號, CU04 As 申請人名稱, CU31 As 聯絡地址, nvl(pcc05,CU08) As 接洽人, CU79 AS 備註, CU80 AS 狀態 From Customer, Staff, potcustcont Where CU02='0' and CU13=ST01(+) And CU04 Like '" & ChgSQL(m_strCustChnName) & "%' and pcc01(+)=cu01 and pcc02(+)=cu127 Order By CU04, CU01 "
'Added by Morgan 2013/4/11
ElseIf m_Type = 2 Then
   strSql = "Select '' As V, ST02 As 智權人員, CU01||CU02 As 代號, rtrim(CU05||' '||CU88||' '||CU89||' '||CU90) As 申請人名稱, CU31 As 聯絡地址, nvl(pcc05,CU08) As 接洽人, CU79 AS 備註 From Customer, Staff, potcustcont Where CU02='0' and CU13=ST01(+) And upper(CU05||' '||CU88||' '||CU89||' '||CU90) Like '" & ChgSQL(UCase(m_strCustChnName)) & "%' and pcc01(+)=cu01 and pcc02(+)=cu127 Order By CU04, CU01 "
'end 2013/4/11
ElseIf m_Type = 3 Then '3.發明人(Add by Lydia 2014/9/22)
   'Modified by Lydia 2019/08/05 台灣地區改顯示為中華民國
   'strSql = "SELECT ' ' as V,IN01||'-'||IN02 AS 發明人編號,IN03 AS 發明人ID,IN04 AS 發明人名稱,IN07 As 地址,NA03 AS 國籍,IN05 as 英文名,IN08 as 英文地址 FROM INVENTOR,NATION WHERE IN11=NA01(+) AND instr(IN04,'" & m_strCustChnName & "')>0 order by IN03 "
   strSql = "SELECT ' ' as V,IN01||'-'||IN02 AS 發明人編號,IN03 AS 發明人ID,IN04 AS 發明人名稱,IN07 As 地址,decode(sign((to_number(substr(na01,1,3)) - 10)),1,na03,'中華民國') AS 國籍,IN05 as 英文名,IN08 as 英文地址 FROM INVENTOR,NATION WHERE IN11=NA01(+) AND instr(IN04,'" & m_strCustChnName & "')>0 order by IN03 "
Else
   'Modified by Morgan 2020/7/2 +英日文查詢
   'strSql = "Select '' As V, FA01||FA02 As 代號, FA04 As 代理人名稱, FA17 As 地址, FA07 As 聯絡人, FA29 AS 備註,'中' As 搜尋語文 From Fagent Where FA02='0' and FA04 Like '" & ChgSQL(m_strCustChnName) & "%' Order By FA04, FA01"
   strSql = "Select '' As V, FA01||FA02 As 代號, decode(min(Lng),'1',max(FA04),'2',max(FA05||' '||FA63||' '||FA64||' '||FA65),max(FA06)) As 代理人名稱" & _
      ", decode(max(Lng),'1',max(FA17),'2',max(FA18||' '||FA19||' '||FA20||' '||FA21||' '||FA22),max(FA23)) As 地址" & _
      ", decode(max(lng),'1',max(FA07),'2',max(FA08),max(FA09)) As 聯絡人, max(FA29) AS 備註,decode(max(Lng),'1','中','2','英','3','日') As 搜尋語文" & _
      " From (select fa01 k1,fa02 k2,'1' Lng from Fagent Where FA02='0' and FA04 Like '" & ChgSQL(m_strCustChnName) & "%'" & _
      " union Select fa01,fa02,'2' Lng from Fagent Where FA02='0' and instr(upper(FA05||' '||FA63||' '||FA64||' '||FA65),upper('" & ChgSQL(m_strCustChnName) & "'))>0" & _
      " union Select fa01,fa02,'3' Lng From Fagent Where FA02='0' and FA06 Like '" & ChgSQL(m_strCustChnName) & "%'" & _
      ") X,Fagent where fa01(+)=k1 and fa02(+)=k2 group by fa01,fa02 Order By 代理人名稱, 代號"
   'end 2020/7/2
    
End If

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    StrMenu = True
    If adoRecordset.RecordCount = 1 Then
    
            'Add by Lydia 2014/9/22 傳回各表單發明人資料
        If m_Type = 3 Then
          If m_frm0908A.Name = "frm210114_1" Then
            m_frm0908A.Txt1(1) = "" & LTrim(RTrim(adoRecordset.Fields(3).Value))
            m_frm0908A.Txt1(2) = "" & LTrim(RTrim(adoRecordset.Fields(2).Value))
            m_frm0908A.Txt1(3) = "" & LTrim(RTrim(adoRecordset.Fields(4).Value))
            m_frm0908A.Txt1(4) = "" & LTrim(RTrim(adoRecordset.Fields(5).Value))
          Else
            m_frm0908A.Txt1(1) = "" & LTrim(RTrim(adoRecordset.Fields(3).Value))
            m_frm0908A.Txt1(2) = "" & LTrim(RTrim(adoRecordset.Fields(6).Value))
            m_frm0908A.Txt1(3) = "" & LTrim(RTrim(adoRecordset.Fields(4).Value))
            m_frm0908A.Txt1(4) = "" & LTrim(RTrim(adoRecordset.Fields(7).Value))
          End If
        'Added by Morgan 2020/7/2
        ElseIf m_Type = 1 Then
         m_Lang = "" & adoRecordset("搜尋語文")
        'end 2020/7/2
        End If
        
      If m_DouChk = False Then  'Add by Lydia 2014/9/22
         m_blnOneRec = True
         If m_Type = 0 Or m_Type = 2 Then
             m_strCustCode = "" & adoRecordset.Fields(2).Value
         Else
             m_strCustCode = "" & adoRecordset.Fields(1).Value
         End If

      Else
         m_blnOneRec = False
         If m_Type = 0 Or m_Type = 2 Then
             m_strCustCode = "" & adoRecordset.Fields(2).Value
         Else
          m_strCustCode = "" & adoRecordset.Fields(1).Value
         End If

       End If
    End If
Else

   If m_Type = 0 Then
      strSql = "Select '' As V, ST02 As 智權人員, CU01||CU02 As 代號, CU04 As 申請人名稱, CU31 As 聯絡地址, nvl(pcc05,CU08) As 接洽人, CU79 AS 備註 From Customer, Staff, potcustcont Where CU02='0' and CU13=ST01(+) And CU04 Like '" & ChgSQL(m_strCustChnName) & "%' and pcc01(+)=cu01 and pcc02(+)=cu127 Order By CU04, CU01 "
   'Added by Morgan 2013/4/11
   ElseIf m_Type = 2 Then
      strSql = "Select '' As V, ST02 As 智權人員, CU01||CU02 As 代號, rtrim(CU05||' '||CU88||' '||CU89||' '||CU90) As 申請人名稱, CU31 As 聯絡地址, nvl(pcc05,CU08) As 接洽人, CU79 AS 備註 From Customer, Staff, potcustcont Where CU02='0' and CU13=ST01(+) And upper(CU05||' '||CU88||' '||CU89||' '||CU90) Like '" & ChgSQL(UCase(m_strCustChnName)) & "%' and pcc01(+)=cu01 and pcc02(+)=cu127 Order By CU04, CU01 "
   'end 2013/4/11
   ElseIf m_Type = 3 Then '3.發明人(Add by Lydia 2014/9/22)
      'Modified by Lydia 2019/08/05 台灣地區改顯示為中華民國
      'strSql = "SELECT ' ' as V,IN01||'-'||IN02 AS 發明人編號,IN03 AS 發明人ID,IN04 AS 發明人名稱,IN07 As 地址,NA03 AS 國籍,IN05 as 英文名,IN08 as 英文地址 FROM INVENTOR,NATION WHERE IN11=NA01(+) AND instr(IN04,'" & m_strCustChnName & "')>0 order by IN03 "
      strSql = "SELECT ' ' as V,IN01||'-'||IN02 AS 發明人編號,IN03 AS 發明人ID,IN04 AS 發明人名稱,IN07 As 地址,decode(sign((to_number(substr(na01,1,3)) - 10)),1,na03,'中華民國') AS 國籍,IN05 as 英文名,IN08 as 英文地址 FROM INVENTOR,NATION WHERE IN11=NA01(+) AND instr(IN04,'" & m_strCustChnName & "')>0 order by IN03 "
   Else
      strSql = "Select '' As V, FA01||FA02 As 代號, FA04 As 代理人名稱, FA17 As 中文地址, FA07 As 聯絡人, FA29 AS 備註 From Fagent Where FA02='0' and FA04 Like '" & ChgSQL(m_strCustChnName) & "%' Order By FA04, FA01 "
   End If

    'Modified by Lydia 2017/12/20 區分訊息
    'ShowNoData
    Select Case m_Type
           Case 0, 2: strExc(1) = "客戶檔搜尋不到符合資料!!"
           Case 3: strExc(1) = "發明人檔搜尋不到符合資料!!"
           Case Else: strExc(1) = "代理人檔搜尋不到符合資料!!"
    End Select
    MsgBox strExc(1), , "沒有資料"
    'end 2017/12/20
    Screen.MousePointer = vbDefault
    m_strCustCode = ""
    Exit Function
End If
Set grdDataList.Recordset = adoRecordset
CheckOC
Screen.MousePointer = vbDefault
End Function

Private Sub Form_Unload(Cancel As Integer)
Set frm090801_1 = Nothing
End Sub

Private Sub grdDataList_SelChange()
Dim ii As Double
Dim intRow As Integer
Dim intCol As Integer

grdDataList.Visible = False
grdDataList.row = grdDataList.MouseRow
grdDataList.col = 0
intRow = Me.grdDataList.row
intCol = Me.grdDataList.col
If grdDataList.row <> 0 Then

 'Add by Lydia 2014/9/22 判斷是否可多選
 If m_DouChk = True Then
 
 Else
'只能選擇一筆
    For ii = 1 To Me.grdDataList.Rows - 1
        If Me.grdDataList.TextMatrix(ii, 0) = "V" And ii <> intRow Then
            Me.grdDataList.TextMatrix(ii, 0) = ""
            For i = 0 To grdDataList.Cols - 1
                 grdDataList.col = i
                 grdDataList.CellBackColor = QBColor(15)
            Next i
        End If
    Next ii
End If


    Me.grdDataList.row = intRow
    Me.grdDataList.col = intCol
    If grdDataList.Text = "V" Then
         grdDataList.Text = ""
         For i = 0 To grdDataList.Cols - 1
              grdDataList.col = i
              grdDataList.CellBackColor = QBColor(15)
        Next i
    Else
         grdDataList.Text = "V"
         For i = 0 To grdDataList.Cols - 1
             grdDataList.col = i
             grdDataList.CellBackColor = &HFFC0C0
         Next i
    End If
End If
grdDataList.Visible = True
End Sub

