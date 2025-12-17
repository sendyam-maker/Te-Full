VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc42b0 
   AutoRedraw      =   -1  'True
   Caption         =   "未繳款資料查詢與銀存核對"
   ClientHeight    =   5184
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5184
   ScaleWidth      =   8760
   Begin VB.TextBox Text5 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   5616
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1362
      Width           =   1284
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   3528
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1362
      Width           =   1284
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1362
      Width           =   1284
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   216
      TabIndex        =   13
      Top             =   1362
      Width           =   1284
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   792
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3120
      TabIndex        =   4
      Top             =   792
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Caption         =   "簽收明細"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7440
      TabIndex        =   9
      Top             =   120
      Width           =   1092
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1200
      Style           =   2  '單純下拉式
      TabIndex        =   0
      Top             =   96
      Width           =   3030
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   444
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   2760
      TabIndex        =   2
      Top             =   444
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   3384
      Left            =   120
      TabIndex        =   8
      Top             =   1752
      Width           =   8508
      _ExtentX        =   15007
      _ExtentY        =   5969
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   16
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   5136
      TabIndex        =   19
      Top             =   1416
      Width           =   132
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   3024
      TabIndex        =   18
      Top             =   1416
      Width           =   132
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   1536
      TabIndex        =   17
      Top             =   1416
      Width           =   132
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "【存摺餘額  =   銀存科目   +(未支款或差異金額)+未繳款簽收資料金額】"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   240
      TabIndex        =   12
      Top             =   1152
      Width           =   7296
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2880
      TabIndex        =   11
      Top             =   792
      Width           =   252
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   10
      Top             =   792
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "簽收日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   444
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   888
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2520
      TabIndex        =   6
      Top             =   444
      Width           =   252
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "銀存科目"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   96
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc42b0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/05 Form2.0已修改 GrdDataList
'2014/04/15 Create By Amy
Option Explicit
Dim i As Integer
'Add by Amy 2014/05/08
Dim strFieldName, intFieldWidth, intFieldAlignment

Private Sub Command1_Click()
    Dim j As Integer
    If grdDataList.Rows = 2 And grdDataList.TextMatrix(1, 1) = "" Then
        Exit Sub
    End If
    
     With grdDataList
        .Enabled = False
        For i = 1 To .Rows - 1
            .col = 0
            .row = i
            If .Text = "V" Then
                .col = 0
                .Text = ""
                For j = 0 To .Cols - 1
                    .col = j
                    .CellBackColor = QBColor(15)
                Next j
                If Not IsNull(grdDataList.TextMatrix(i, 14)) Then
                    strExitControl = MsgText(601)
                    tool4_enabled
                    Screen.MousePointer = vbHourglass
                    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
                    Frmacc41e0.txtA2301 = grdDataList.TextMatrix(i, 15) 'Modify by Amy 2014/05/08 增加簽收確認日期欄位
                    Frmacc41e0.Tag = "Frmacc42b0"
                    Frmacc41e0.StrMenu
                    Frmacc41e0.Show
                    Me.Hide
                    Screen.MousePointer = vbDefault
                    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
                    Exit For
                End If
            End If
        Next i
        .Enabled = True
    End With
   
End Sub

Private Sub Form_Activate()
    tool3_enabled
   strFormName = Name
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   '設定Grid欄位名稱/大小/對齊1.flexAlignLeftCenter 4.flexAlignCenterCenter 7.flexAlignRightCenter
   strFieldName = Array("V", "銀存科目", "智權人員", "客戶名稱", "簽收日期", _
                                    "金額", "簽收確認日期", "到期日", "票據號碼", "收票銀行", _
                                    "收票帳號", "", "", "", "", "")
                                    
   intFieldWidth = Array(300, 2700, 1100, 1485, 1050, _
                                    1300, 1530, 825, 1100, 1300, _
                                    1300, 0, 0, 0, 0, 0)
                                    
   intFieldAlignment = Array(1, 1, 1, 1, 7, _
                                           7, 7, 7, 7, 1, _
                                           1, 1, 1, 1, 1, 1)
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8880
   Me.Height = 5535
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
  
  '銀存科目
  'Modify by Amy 2023/05/25 修改項目改抓共用
'   Combo1.Clear
'   Combo1.AddItem MsgText(601)
'   If pub_strUserOffice = "2" Then
'      strExc(0) = "select a0101||' '||a0102 from acc010 where  a0101='1911'"
'   ElseIf pub_strUserOffice = "3" Then
'      strExc(0) = "select a0101||' '||a0102 from acc010 where  a0101='1912'"
'   ElseIf pub_strUserOffice = "4" Then
'      strExc(0) = "select a0101||' '||a0102 from acc010 where  a0101='1913'"
'   Else
'      'Modify by Amy 2014/05/08 +會計科目110209
'      'Modify by Morgan 2016/11/11 +1101 北所現金
'      'Modify by Amy 2020/04/10 +110602/110502
'      strExc(0) = "Select a0101||' '||a0102 C1,decode(a0101,'110202',1,'110207',2,'110303',3,'110223',4,'110208',5,'110208',6,'110204',7,'110205',8,'110301',9,'110302',10,'110209',11,'110602',12,'110502',13,14) Srt" & _
'         " From Acc010 Where  a0101 in ('110202','110204','110205','110207','110208','110209','110223','110602','110502','110301','110302','110303')"
'      strExc(0) = strExc(0) & " union select '1101 北所現金',15 from dual"
'      strExc(0) = strExc(0) & " order by 2"
'   End If
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      Do While Not RsTemp.EOF
'         Combo1.AddItem RsTemp(0)
'         RsTemp.MoveNext
'      Loop
'   End If
   Pub_AccBankTit Combo1, Me.Name
   'end 2025/05/25

   SetDataListWidth
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   
   'Added by Morgan 2025/6/10 簽收日期預設1日至(系統日-1日)--瑞婷
   strExc(2) = CompDate(2, -1, strSrvDate(1))
   strExc(1) = Left(strExc(2), 6) & "01"
   MaskEdBox1 = ChangeWStringToTDateString(strExc(1))
   MaskEdBox2 = ChangeWStringToTDateString(strExc(2))
   'end 2025/6/10
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc42b0 = Nothing
End Sub

Private Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
   
   With grdDataList
      .Visible = False
      If p_bolHeaderOnly = True Then
         .Clear
         .Rows = 2: .Cols = UBound(strFieldName) + 1: .FixedRows = 1
      End If
      .row = 0
      
      'Modify by Amy 2014/05/08 +簽收確認日期(a2321) 欄位設定改寫至Array
      For i = 0 To UBound(strFieldName)
        .col = i: .ColWidth(.col) = intFieldWidth(i): .Text = strFieldName(i)
        .ColAlignment(.col) = intFieldAlignment(i)
        .CellAlignment = flexAlignCenterCenter: .CellFontBold = True 'Head 對齊/粗體
      Next i
      'end 2014/05/08
      
      .Visible = True
   End With
End Sub

Private Sub KeyDefine(KeyCode As Integer)
    Select Case KeyCode
        Case vbKeyF12
            If FormCheck = True Then
                Screen.MousePointer = vbHourglass
                'Add by Amy 2016/08/24
                Text1_LostFocus (0)
                Text1_LostFocus (1)
                QueryTable
                Screen.MousePointer = vbDefault
            End If
    End Select
    
    KeyEnter KeyCode
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Function FormCheck() As Boolean
    If MaskEdBox1.Text = MsgText(29) Then
        MsgBox "請輸入簽收日期起日!!", vbExclamation
        MaskEdBox1.SetFocus
        Exit Function
    ElseIf IsDate(ChangeTStringToWDateString(Replace(Me.MaskEdBox1.Text, "/", ""))) = False Then
        MsgBox "日期輸入錯誤!!", vbExclamation
        MaskEdBox1.SetFocus
        Exit Function
    End If
   
    If MaskEdBox2.Text = MsgText(29) Then
        MsgBox "請輸入簽收日期迄日!!", vbExclamation
        MaskEdBox2.SetFocus
        Exit Function
    ElseIf IsDate(ChangeTStringToWDateString(Replace(Me.MaskEdBox2.Text, "/", ""))) = False Then
        MsgBox "日期輸入錯誤!!", vbExclamation
        MaskEdBox2.SetFocus
        Exit Function
    End If
    
    If Val(Replace(Me.MaskEdBox1.Text, "/", "")) > Val(Replace(Me.MaskEdBox2.Text, "/", "")) Then
        MsgBox "簽收日期迄日不可大於起日"
        MaskEdBox2.SetFocus
        Exit Function
    End If
    
    FormCheck = True
End Function

Private Sub QueryTable()
    Dim strSql As String, strWhere As String, strWhereNot1101 As String, bolCashOnly As Boolean
    
On Error GoTo ErrHnd

   Text3 = ""
   Text4 = ""
   Text5 = ""
   If Combo1 <> "" And Text1(0) & Text1(1) = "" Then
      '科目餘額
      '同科目分類帳查詢，不須加上未傳送資料的收款資料--瑞婷
      strExc(1) = FCDate(MaskEdBox1.Text) '起
      strExc(2) = FCDate(MaskEdBox2.Text) '迄
      
      strExc(3) = CompDate(1, -1, strExc(1)) '上月
      strExc(4) = Val(Left(strExc(3), 4)) - 1911 '年
      strExc(5) = Val(Mid(strExc(3), 5, 2)) '月
      
      '上月餘額
      strExc(0) = "select sum(a0408) from acc040 where a0405='" & Trim(Val(Combo1)) & "' and a0404='TOT' and a0401=" & strExc(4) & " and a0402=" & strExc(5)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '之後的借/貸合計
         strExc(6) = Val("" & RsTemp(0))
         
         strExc(0) = "select sum(ax206) AmtD,sum(ax207) AmtC from acc020,acc021 where a0205>=" & strExc(1) & " and a0205<=" & strExc(2) & _
            " and ax201(+)=a0201 and ax202(+)=a0202 and ax205='" & Trim(Val(Combo1)) & "'"
            
         '"select sum(a1p07),sum(a1p08) from acc1p0 where a1p18>=" & Left(strExc(1), 5) & "01 and a1p05='" & Trim(Val(Combo1)) & "' and a1p22 is null)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(7) = Val(strExc(6)) + RsTemp(0) - RsTemp(1)
         End If
         
         '餘額=上月餘額+借-貸
         Text3 = Format(strExc(7), DDollar2)
      End If
   End If
   'end 2025/6/16
    
    'Added by Morgan 2014/10/8 排除手動更新簽收確認日期者(未輸繳款直接輸收款)--辜
    strWhere = " and (a2321 is null or to_char(a2321,'hh24miss')<>'000000')"
    'end 2014/10/8
    
   'Added by Morgan 2025/6/17
   If pub_strUserOffice <> "1" Then
      If pub_strUserOffice = "2" Then
         strWhere = strWhere & " and A2322 like '1911%'"
      ElseIf pub_strUserOffice = "3" Then
         strSql = strSql & " and A2322 like '1912%'"
      ElseIf pub_strUserOffice = "4" Then
         strSql = strSql & " and A2322 like '1913%'"
      End If
   End If
   'end 2025/6/17
         
    If Trim(Combo1) <> MsgText(601) Then
        'Modified by Morgan 2016/11/11
        'strWhere = strWhere & " And A2322='" & Left(Combo1.Text, InStr(Combo1, " ") - 1) & "' "
        strWhereNot1101 = strWhereNot1101 & " And A2322='" & Left(Combo1.Text, InStr(Combo1, " ") - 1) & "' "
        If Left(Combo1.Text, InStr(Combo1, " ") - 1) = "1101" Then bolCashOnly = True
        'end 2016/11/11
    End If
    'Add by Amy 2016/08/24 +客戶編號
    If Text1(0) <> MsgText(601) Then
        strWhere = strWhere & " And A2304>='" & Text1(0) & "' "
    End If
    If Text1(1) <> MsgText(601) Then
        strWhere = strWhere & " And A2304<='" & Text1(1) & "' "
    End If
    'end 2016/08/24
    
    'Modify by Amy 2014/05/08 +簽收確認日期及銀存科目110209 原抓簽收確認時間為空,改剔除以簽收單號抓a4421或a4427且Length(a4416)>1
    'Modify by Amy 2016/08/24 因抓InStr(a4427|| a4421,a2301)>0會慢所以改寫至TempTB,先抓a4421=a2301再過濾 a4427
'    strSql = "Select '' as V,a2322||a0102 as 銀存科目,st02 as 智權人員,NVL(SubStr(cu04,1,6),Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 客戶名稱,Nvl(to_char(a2302),'') as 簽收日期,to_char(a2318,'999,999,999') as 金額,Decode(a2321,null,'',to_char(a2321, 'yyyymmdd') -19110000) as 簽收確認日期,Nvl(to_char(a2325),'') as 到期日,a2326 as 票據號碼,a2327 as 收票銀行,a2328 as 收票帳號," & _
'                  "a2322,a2305,a2303,a2304,a2301 From Acc230,Acc010,Staff,Customer Where a2322=a0101(+) And a2303=st01(+) And SubStr(a2304,1,8)=cu01(+) And SubStr(a2304,9,1)=cu02(+) And a2322 is not null And a2322 In ('110202','110204','110205','110207','110208','110209','110223','110301','110302','110303') " & _
'                  "And a2302 >=" & FCDate(MaskEdBox1.Text) & " And a2302<=" & FCDate(MaskEdBox2.Text) & strWhere & _
'                  " And Not Exists (Select * From Acc440 Where InStr(a4427|| a4421,a2301)>0 And Length(a4416)>1) "
'    '銀存科目小計
'    strSql = strSql & " Union " & _
'                 "Select '' as V,'                  小  計' as 銀存科目,'' as 智權人員,'' as 客戶名稱,'' as 簽收日期,to_char(sum(a2318),'999,999,999') as 金額,'' as 簽收確認日期,'' as 到期日,'' as 票據號碼,''as 收票銀行,'' as 收票帳號," & _
'                  "a2322||'ZZ' as a2322,'Z' as a2305,'ZZZZZZ' as a2303,'ZZZZZZZZZ' as a2304,'ZZZZZZZZZZ' as a2301 From Acc230 Where a2322 is not null And a2322 In ('110202','110204','110205','110207','110208','110209','110223','110301','110302','110303') " & _
'                  "And a2302 >=" & FCDate(MaskEdBox1.Text) & " And a2302<=" & FCDate(MaskEdBox2.Text) & strWhere & _
'                  " And Not Exists (Select * From Acc440 Where InStr(a4427|| a4421,a2301)>0 And Length(a4416)>1) " & _
'                  "Group by a2322 "
'    '總計
'    strSql = strSql & " Union " & _
'                 "Select '' as V,'                總        計' as 銀存科目,'' as 智權人員,'' as 客戶名稱,'' as 簽收日期,to_char(sum(a2318),'999,999,999') as 金額,'' as 簽收確認日期,'' as 到期日,'' as 票據號碼,''as 收票銀行,'' as 收票帳號," & _
'                  "'ZZZZZZZZZ' as a2322,'Z' as a2305,'ZZZZZZ' as a2303,'ZZZZZZZZZ' as a2304,'ZZZZZZZZZZ' as a2301 From Acc230 Where a2322 is not null And a2322 In ('110202','110204','110205','110207','110208','110209','110223','110301','110302','110303') " & _
'                  "And a2302 >=" & FCDate(MaskEdBox1.Text) & " And a2302<=" & FCDate(MaskEdBox2.Text) & strWhere & _
'                  " And Not Exists (Select * From Acc440 Where InStr(a4427|| a4421,a2301)>0 And Length(a4416)>1) "
'    '2014/05/08 改依 銀存科目+簽收單號 排序 原:銀存科目+所別+智權人+客戶編號+簽收日期+簽收單號 排序
'    strSql = "Select * From(" & strSql & ") Order by a2322,a2301"
    'end 2014/05/08
    strSql = "Delete From Accrpt42b0 Where ID='" & strUserNum & "'"
    cnnConnection.Execute strSql

    '抓取資料寫入暫存檔
    If Not bolCashOnly Then 'Added by Morgan 2016/11/11 指定北所現金科目時不抓
      'Modify by Amy 2020/04/10 +110602/110502
      'Modified by Morgan 2025/6/16 +191102,191202,191302
      strSql = "Insert Into Accrpt42b0 (ID,R001,R002,R003,R004,R005,R006,R007,R008,R009,R010,R011,R012) " & _
                   "Select '" & strUserNum & "',a2301 as 簽收單號,a2305 as 所別,a2322 as 銀存科目,a2303 as 智權人員,a2304 as 客戶編號,Decode(Nvl(a2302,0),0,'',a2302) as 簽收日期" & _
                   ",a2318 as 金額,Decode(a2321,null,'',to_char(a2321, 'yyyymmdd') -19110000) as 簽收確認日期,Decode(Nvl(a2325,0),0,'',a2325) as 到期日" & _
                   ",a2326 as 票據號碼,a2327 as 收票銀行,a2328 as 收票帳號 " & _
                    "From Acc230 Where a2322 is not null And a2322 In ('110202','110204','110205','110207','110208','110209','110223','110301','110302','110303','110602','110502','191102','191202','191302') " & _
                    "And a2302 >=" & FCDate(MaskEdBox1.Text) & " And a2302<=" & FCDate(MaskEdBox2.Text) & strWhere & strWhereNot1101 & _
                    " And Not Exists (Select * From Acc440 Where  a4421=a2301 And Length(a4416)>1) "
      cnnConnection.Execute strSql
   End If
    
    'Added by Morgan 2016/11/11 +北所現金
    If Trim(Combo1) = MsgText(601) Or bolCashOnly Then
      strSql = "Insert Into Accrpt42b0 (ID,R001,R002,R003,R004,R005,R006,R007,R008,R009,R010,R011,R012) " & _
                   "Select '" & strUserNum & "',a2301 as 簽收單號,a2305 as 所別,'1101' as 銀存科目,a2303 as 智權人員,a2304 as 客戶編號,Decode(Nvl(a2302,0),0,'',a2302) as 簽收日期" & _
                   ",a2317 as 金額,Decode(a2321,null,'',to_char(a2321, 'yyyymmdd') -19110000) as 簽收確認日期,'' as 到期日" & _
                   ",a2326 as 票據號碼,a2327 as 收票銀行,a2328 as 收票帳號 " & _
                    " From Acc230 Where a2317>0 and a2305='1'" & _
                    " And a2302 >=" & FCDate(MaskEdBox1.Text) & " And a2302<=" & FCDate(MaskEdBox2.Text) & strWhere & _
                    " And Not Exists (Select * From Acc440 Where  a4421=a2301 And Length(a4416)>1) "
      cnnConnection.Execute strSql
    End If
    'end 2016/11/11
    
    'Added by Morgan 2025/6/16
    If (Trim(Combo1) = MsgText(601) Or Left(Combo1, 6) = "191101") And (pub_strUserOffice = "1" Or pub_strUserOffice = "2") Then
      strSql = "Insert Into Accrpt42b0 (ID,R001,R002,R003,R004,R005,R006,R007,R008,R009,R010,R011,R012) " & _
                   "Select '" & strUserNum & "',a2301 as 簽收單號,a2305 as 所別,'191101' as 銀存科目,a2303 as 智權人員,a2304 as 客戶編號,Decode(Nvl(a2302,0),0,'',a2302) as 簽收日期" & _
                   ",a2317 as 金額,Decode(a2321,null,'',to_char(a2321, 'yyyymmdd') -19110000) as 簽收確認日期,'' as 到期日" & _
                   ",a2326 as 票據號碼,a2327 as 收票銀行,a2328 as 收票帳號 " & _
                    " From Acc230 Where a2317>0 and a2305='2'" & _
                    " And a2302 >=" & FCDate(MaskEdBox1.Text) & " And a2302<=" & FCDate(MaskEdBox2.Text) & strWhere & _
                    " And Not Exists (Select * From Acc440 Where  a4421=a2301 And Length(a4416)>1) "
      cnnConnection.Execute strSql
    End If
    If (Trim(Combo1) = MsgText(601) Or Left(Combo1, 6) = "191201") And (pub_strUserOffice = "1" Or pub_strUserOffice = "3") Then
      strSql = "Insert Into Accrpt42b0 (ID,R001,R002,R003,R004,R005,R006,R007,R008,R009,R010,R011,R012) " & _
                   "Select '" & strUserNum & "',a2301 as 簽收單號,a2305 as 所別,'191201' as 銀存科目,a2303 as 智權人員,a2304 as 客戶編號,Decode(Nvl(a2302,0),0,'',a2302) as 簽收日期" & _
                   ",a2317 as 金額,Decode(a2321,null,'',to_char(a2321, 'yyyymmdd') -19110000) as 簽收確認日期,'' as 到期日" & _
                   ",a2326 as 票據號碼,a2327 as 收票銀行,a2328 as 收票帳號 " & _
                    "From Acc230 Where a2317>0 and a2305='3' " & _
                    "And a2302 >=" & FCDate(MaskEdBox1.Text) & " And a2302<=" & FCDate(MaskEdBox2.Text) & strWhere & _
                    " And Not Exists (Select * From Acc440 Where  a4421=a2301 And Length(a4416)>1) "
      cnnConnection.Execute strSql
    End If
    If (Trim(Combo1) = MsgText(601) Or Left(Combo1, 6) = "191301") And (pub_strUserOffice = "1" Or pub_strUserOffice = "4") Then
      strSql = "Insert Into Accrpt42b0 (ID,R001,R002,R003,R004,R005,R006,R007,R008,R009,R010,R011,R012) " & _
                   "Select '" & strUserNum & "',a2301 as 簽收單號,a2305 as 所別,'191301' as 銀存科目,a2303 as 智權人員,a2304 as 客戶編號,Decode(Nvl(a2302,0),0,'',a2302) as 簽收日期" & _
                   ",a2317 as 金額,Decode(a2321,null,'',to_char(a2321, 'yyyymmdd') -19110000) as 簽收確認日期,'' as 到期日" & _
                   ",a2326 as 票據號碼,a2327 as 收票銀行,a2328 as 收票帳號 " & _
                    "From Acc230 Where a2317>0 and a2305='4' " & _
                    "And a2302 >=" & FCDate(MaskEdBox1.Text) & " And a2302<=" & FCDate(MaskEdBox2.Text) & strWhere & _
                    " And Not Exists (Select * From Acc440 Where  a4421=a2301 And Length(a4416)>1) "
      cnnConnection.Execute strSql, intI
    End If
    'end 2025/6/16
    
    '刪除存在於智權人員繳款記錄資料(a4427多筆)
    strSql = "Delete From Accrpt42b0 Where ID='" & strUserNum & "' And Exists (Select * From Acc440 Where  a4427 is not null and InStr(a4427,R001)>0  And Length(a4416)>1) "
    cnnConnection.Execute strSql, intI

    '銀存科目小計
    strSql = "Insert Into Accrpt42b0 (ID,R001,R003,R007) " & _
                 "Select '" & strUserNum & "','Z' as 簽收單號,R003||'ZZ' as 銀存科目,Sum(R007) as 金額 " & _
                  "From Accrpt42b0 Where ID='" & strUserNum & "' Group by R003 "
     cnnConnection.Execute strSql

    '總計
    strSql = "Insert Into Accrpt42b0 (ID,R001,R003,R007) " & _
                 "Select '" & strUserNum & "','Z' as 簽收單號,'ZZZZZZZZZ' as 銀存科目,Sum(R007) as 金額 " & _
                  "From Accrpt42b0 Where ID='" & strUserNum & "' And InStr(R003,'Z')=0"
      cnnConnection.Execute strSql

    strSql = "Select '' as V,Decode(R003,'ZZZZZZZZZ','                總        計',Decode(SubStr(R003,length(R003)-1,2),'ZZ','                  小  計',R003||a0102)) as 銀存科目" & _
                ",st02 as 智權人員,NVL(SubStr(cu04,1,6),Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 客戶名稱" & _
                ",R006 as 簽收日期,to_char(R007,'999,999,999') as 金額,R008 as 簽收確認日期,R009 as 到期日,R010 as 票據號碼,R011 as 收票銀行,R012 as 收票帳號," & _
                "R003,R002,R004,R005,R001 From Accrpt42b0,Acc010,Staff,Customer " & _
                "Where ID='" & strUserNum & "' And R003=a0101(+) And R004=st01(+) And SubStr(R005,1,8)=cu01(+) And SubStr(R005,9,1)=cu02(+) " & _
                "Order by R003,R001"
    'end 2016/08/24

    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
    If RsTemp.RecordCount = 1 Then
        grdDataList.Clear
        SetDataListWidth True
        MsgBox MsgText(28), , MsgText(5)
        Exit Sub
    End If
    Set grdDataList.Recordset = RsTemp
    SetDataListWidth
   
   'Added by Morgan 2025/6/16
   If Combo1 <> "" And Text1(0) & Text1(1) = "" Then
      '未繳款簽收資料金額
      strExc(0) = "select R007 from Accrpt42b0 Where ID='" & strUserNum & "' and R003='ZZZZZZZZZ'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Text5 = Format(Val("" & RsTemp(0)), DDollar2)
      End If
      SumShow
   End If
   'end 2025/6/16
         
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub GrdDataList_Click()
    If grdDataList.Rows = 2 And grdDataList.TextMatrix(1, 1) = "" Then
        Exit Sub
    End If
    
    With grdDataList
        .Visible = False
        .row = .MouseRow
        .col = 0
        If .row <> 0 Then
            If Trim(.TextMatrix(.row, 1)) <> "小  計" And Trim(.TextMatrix(.row, 1)) <> "總        計" Then
                If .Text = "V" Then
                    .Text = ""
                    For i = 0 To .Cols - 1
                        .col = i
                        .CellBackColor = QBColor(15)
                    Next i
                Else
                    .Text = "V"
                    For i = 0 To .Cols - 1
                        .col = i
                        .CellBackColor = &HFFC0C0
                    Next i
                End If
            End If
        End If
        .Visible = True
    End With
End Sub

'Add by Amy 2016/08/24
Private Sub Text1_GotFocus(Index As Integer)
    TextInverse Text1(Index)
    CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If Text1(Index) = MsgText(601) Then Exit Sub
    
    If Len(Text1(Index)) = 6 Then Text1(Index) = AfterZero(Text1(Index))
    
End Sub
'Added by Morgan 2025/5/28
Private Sub SumShow()
   If Text2 <> "" And Text3 <> "" Then
      Text4 = Format(Val(Format(Text2)) - Val(Format(Text3)) - Val(Format(Text5)), DDollar2)
   End If
End Sub

Private Sub Text2_GotFocus()
   Text2 = Format(Text2)
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 <> "" Then
      Text2 = Format(Text2, DDollar2)
      SumShow
   Else
      Text4 = ""
   End If
End Sub
