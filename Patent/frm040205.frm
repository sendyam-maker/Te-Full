VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040205 
   BorderStyle     =   1  '單線固定
   Caption         =   "FC收款請款點數查詢"
   ClientHeight    =   4155
   ClientLeft      =   3495
   ClientTop       =   3075
   ClientWidth     =   5505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5505
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   1185
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1680
      Width           =   1185
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   2700
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1680
      Width           =   1185
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1185
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1995
      Width           =   1185
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   2700
      MaxLength       =   4
      TabIndex        =   10
      Top             =   2310
      Width           =   1185
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1185
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2310
      Width           =   1185
   End
   Begin VB.CommandButton Cmdok 
      Caption         =   "結束(&X)"
      Height          =   350
      Index           =   1
      Left            =   4500
      TabIndex        =   13
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton Cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   3675
      TabIndex        =   12
      Top             =   20
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1185
      MaxLength       =   1
      TabIndex        =   11
      Top             =   2625
      Width           =   1185
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2700
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1365
      Width           =   1185
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1185
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1365
      Width           =   1185
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2700
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1050
      Width           =   1185
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1185
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1050
      Width           =   1185
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1185
      MaxLength       =   1
      TabIndex        =   1
      Top             =   744
      Width           =   1185
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1185
      TabIndex        =   0
      Top             =   456
      Width           =   4200
   End
   Begin MSForms.Label lbl1 
      Height          =   285
      Left            =   2400
      TabIndex        =   34
      Top             =   1985
      Width           =   1620
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "2857;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "  統計該業務區收文的請款點數, 未扣除分配至其他單位的點數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   19
      Left            =   120
      TabIndex        =   33
      Top             =   3795
      Width           =   5250
   End
   Begin VB.Label Label1 
      Caption         =   "       法務系統類別 : FCL,CFL,L,LA"
      Height          =   195
      Index           =   18
      Left            =   120
      TabIndex        =   32
      Top             =   3585
      Width           =   5220
   End
   Begin VB.Label Label1 
      Caption         =   "       商標系統類別 : FCT,CFT,CFC,S,T,TF,TB,TC,TD,TM,TR,TS,TT"
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   31
      Top             =   3390
      Width           =   5220
   End
   Begin VB.Label Label1 
      Caption         =   "       專利系統類別 : FCP,FG,P,PS,CFP,CPS"
      Height          =   195
      Index           =   11
      Left            =   120
      TabIndex        =   30
      Top             =   3195
      Width           =   5220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "( ALL：全部 )"
      Height          =   180
      Index           =   17
      Left            =   1200
      TabIndex        =   29
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "　外法：F30~F49"
      Height          =   180
      Index           =   16
      Left            =   4080
      TabIndex        =   28
      Top             =   2040
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "　外專：F20~F29"
      Height          =   180
      Index           =   15
      Left            =   4080
      TabIndex        =   27
      Top             =   1800
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "　外商：F10~F19"
      Height          =   180
      Index           =   14
      Left            =   4080
      TabIndex        =   26
      Top             =   1560
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "業務區說明："
      Height          =   180
      Index           =   13
      Left            =   4080
      TabIndex        =   25
      Top             =   1320
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "PS : 國外部查詢, 因有跨部門收文情形, 故系統類別建議用ALL全部"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   5220
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2430
      X2              =   2625
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      Height          =   180
      Index           =   9
      Left            =   380
      TabIndex        =   23
      Top             =   1725
      Width           =   732
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "智權人員："
      Height          =   180
      Index           =   8
      Left            =   210
      TabIndex        =   22
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   7
      Left            =   190
      TabIndex        =   21
      Top             =   2340
      Width           =   912
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2430
      X2              =   2625
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2430
      X2              =   2625
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2430
      X2              =   2625
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "(1.明細 2.總計)"
      Height          =   180
      Index           =   6
      Left            =   2430
      TabIndex        =   20
      Top             =   2685
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "(1.請款 2.收款)"
      Height          =   180
      Index           =   5
      Left            =   2430
      TabIndex        =   19
      Top             =   795
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "代理人國籍："
      Height          =   180
      Index           =   4
      Left            =   30
      TabIndex        =   18
      Top             =   1395
      Width           =   1100
   End
   Begin VB.Label Label1 
      Caption         =   "查詢內容："
      Height          =   180
      Index           =   3
      Left            =   190
      TabIndex        =   17
      Top             =   2655
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "日期："
      Height          =   180
      Index           =   2
      Left            =   560
      TabIndex        =   16
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "查詢別："
      Height          =   180
      Index           =   1
      Left            =   380
      TabIndex        =   15
      Top             =   780
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   190
      TabIndex        =   14
      Top             =   504
      Width           =   912
   End
End
Attribute VB_Name = "frm040205"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; lbl1
'Memo By Sonia 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'2007/11/16 整理 by sonia
Option Explicit

Dim strSql As String, i As Integer, j As Integer, s As Integer
Dim strTemp As Variant, strTemp1 As Variant
'add by nickc 2008/01/18
Dim stST05 As String
Dim strBaseTable As String, strCaseNa As String


Private Sub cmdok_Click(Index As Integer)
'On Error GoTo Checking

Select Case Index
   Case 0 '確定
      If Len(txt1(0)) = 0 Or Len(txt1(1)) = 0 Or Len(txt1(2)) = 0 Or Len(txt1(3)) = 0 Or Len(txt1(6)) = 0 Then
         If Len(txt1(0)) = 0 Then s = MsgBox("系統類別不可空白", , "USER 輸入錯誤"): txt1(0).SetFocus: txt1_GotFocus (0): Exit Sub
         If Len(txt1(1)) = 0 And txt1(6) = "1" Then s = MsgBox("查詢別不可空白", , "USER 輸入錯誤"): txt1(1).SetFocus: txt1_GotFocus (1): Exit Sub
         'Add By Cheng 2002/03/19
         If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
            Me.txt1(2).SetFocus
            txt1_GotFocus 2
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
            Me.txt1(3).SetFocus
            txt1_GotFocus 3
            Exit Sub
         End If
         If Len(txt1(3)) = 0 Then s = MsgBox("日期區間迄不可空白", , "USER 輸入錯誤"): txt1(2).SetFocus: txt1_GotFocus (2): Exit Sub
         If Len(txt1(6)) = 0 Then s = MsgBox("查詢內容不可空白", , "USER 輸入錯誤"): txt1(6).SetFocus: txt1_GotFocus (6): Exit Sub
      End If
'2007/11/16 cancel by sonia 因外商收法務案故取消使用權限控制
'      'Add By Cheng 2003/01/03
'      '檢查系統類別的使用權限
'      If Len(Trim(txt1(0))) <> 0 Then
'         strTemp = Split(GetSystemKindByNick, ",")
'         strTemp1 = Split(txt1(0), ",")
'         For i = 0 To UBound(strTemp1)
'            s = 0
'            For j = 0 To UBound(strTemp)
'               If strTemp1(i) = strTemp(j) Then
'                   s = 1
'               End If
'            Next j
'            If s = 0 Then
'               s = MsgBox(strUserNum + " 沒有 " + strTemp1(i) + " 的使用權限 ", , "USER 權限不足!!!")
'               txt1(0).SetFocus
'               txt1(0).SelStart = 0
'               txt1(0).SelLength = Len(txt1(0))
'               Exit Sub
'            End If
'         Next i
'      End If
'2007/11/16 end
      If Val(Me.txt1(2).Text) > Val(Me.txt1(3).Text) Then MsgBox "日期區間範圍輸入錯誤!!!", vbExclamation + vbOKOnly: txt1(2).SetFocus: txt1_GotFocus (2): Exit Sub
      
      'Add By Sindy 2018/7/23
      If Len(frm040205.txt1(10)) <> 0 And Len(frm040205.txt1(11)) = 0 Then
         s = MsgBox("業務區間迄不可空白", , "USER 輸入錯誤")
         txt1(11).SetFocus
         Exit Sub
      End If
      If Len(frm040205.txt1(11)) <> 0 And Len(frm040205.txt1(10)) = 0 Then frm040205.txt1(11) = ""
      If Len(frm040205.txt1(10)) <> 0 And Len(frm040205.txt1(10)) < 3 Then frm040205.txt1(10) = Left(frm040205.txt1(10) & "00", 3)
      If Len(frm040205.txt1(11)) <> 0 And Len(frm040205.txt1(11)) < 2 Then
         frm040205.txt1(11) = frm040205.txt1(11) & "0Z"
      ElseIf Len(frm040205.txt1(11)) <> 0 And Len(frm040205.txt1(11)) < 3 Then
         frm040205.txt1(11) = frm040205.txt1(11) & "Z"
      End If
      '2018/7/23 END
      
      'add by nickc 2008/01/18 陳經理新增加的規則
      'Modify By Sindy 2009/05/21 80030.洪琬姿及78011.葉易雲不受限制
      If (stST05 >= "21" And stST05 <= "29") And strUserNum <> "80030" And strUserNum <> "78011" And Trim(txt1(9)) = "" Then MsgBox "智權人員不可以空白!!!", vbExclamation + vbOKOnly: txt1(9).SetFocus: Exit Sub
      
      'add by nickc 2008/01/18
      'Modify By Sindy 2009/05/21 80030.洪琬姿及78011.葉易雲不受限制
      If (stST05 = "21" Or stST05 = "26" Or stST05 = "28") And strUserNum <> "80030" And strUserNum <> "78011" Then
           If PUB_GetStaffST16(strUserNum) <> PUB_GetStaffST16(txt1(9)) Then
                MsgBox "僅可以輸入相同組別的智權人員！", vbExclamation, "操作錯誤！"
                txt1(9).SetFocus
                txt1_GotFocus 9
                Exit Sub
           End If
      End If
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/9/28 清除查詢印表記錄檔欄位
      Me.Enabled = False
      
      'Add By Sindy 2018/7/19 案件主檔
      strBaseTable = "": strCaseNa = ""
      strBaseTable = ", (SELECT PA01 VC01,pa02 vc02,pa03 vc03,pa04 vc04,pa26 vc05,pa75 vc06,pa09 vc07,PA57 vc08 FROM PATENT " & _
                    "UNION SELECT tm01,tm02,tm03,tm04,tm23,tm44,tm10,TM29 FROM TRADEMARK " & _
                    "UNION SELECT lc01,lc02,lc03,lc04,lc11,lc22,lc15,LC08 FROM LAWCASE " & _
                    "UNION SELECT hc01,hc02,hc03,hc04,hc05,' ',' ',HC09 FROM HIRECASE " & _
                    "UNION SELECT sp01,sp02,sp03,sp04,sp08,sp26,sp09,SP15 FROM SERVICEPRACTICE ) VT1 "
      If txt1(1) = "1" Then '請款點數
         strCaseNa = " AND CP01=VC01(+) AND CP02=VC02(+) AND CP03=VC03(+) AND CP04=VC04(+)"
      Else                '收款點數
         strCaseNa = " A1K13=VC01(+) AND A1K14=VC02(+) AND A1K15=VC03(+) AND A1K16=VC04(+)"
      End If
'         If txtKind = "1" Then '申請人
'             strBaseTable = strBaseTable & ",CUSTOMER,NATION,CasePropertyMap,SYSTEMKIND "
'             strCaseNa = strCaseNa & " AND SUBSTR(VC05,1,8)=CU01(+) AND SUBSTR(VC05,9,1)=CU02(+) AND SUBSTR(CU10,1,3)=NA01(+)"
'         Else                  '代理人
      strBaseTable = strBaseTable & ",FAGENT,NATION,CasePropertyMap,SYSTEMKIND "
'         End
      strCaseNa = strCaseNa & " AND SUBSTR(VC06,1,8)=FA01(+) AND SUBSTR(VC06,9,1)=FA02(+) AND SUBSTR(FA10,1,3)=NA01(+)"
      strCaseNa = strCaseNa & " AND cp01=sk01(+) AND CP01=CPM01(+) AND CP10=CPM02(+)"
      '2018/7/19 END
      
      'Add By Sindy 2018/7/19 產生收款資料
      If txt1(1) = "2" Then '收款
         Call Select2
      End If
      '2018/7/19 END
      Select Case txt1(6)
         Case "1" '明細
            pub_QL05 = pub_QL05 & ";" & Label1(3) & "明細" 'Add By Sindy 2010/9/28
            Screen.MousePointer = vbHourglass
            frm040205a.Show
            'frm040205a.Hide
            'frm040205a.Tag
            frm040205a.StrMenu
            Screen.MousePointer = vbDefault
            Me.Hide
            'frm040205a.Show
            Do
            DoEvents
            If bolToEndByNick = True Then Unload Me: Exit Sub
            Loop Until Not frm040205a.Visible
            Unload frm040205a
         Case "2" '總計
            pub_QL05 = pub_QL05 & ";" & Label1(3) & "總計" 'Add By Sindy 2010/9/28
            Screen.MousePointer = vbHourglass
            frm040205b.Show
            'frm040205b.Hide
            frm040205b.StrMenu
            Screen.MousePointer = vbDefault
            Me.Hide
            'frm040205b.Show
            Do
            DoEvents
            If bolToEndByNick = True Then
               Unload Me
               Exit Sub
            End If
            Loop Until Not frm040205b.Visible
            Unload frm040205b
         Case Else
      End Select
      Me.Enabled = True
      Me.Show
   Case 1 '結束
      Unload Me
   Case Else '其他
   End Select
   Exit Sub
   
'Checking:
'   Screen.MousePointer = vbDefault
'   MsgBox Err.Description, , MsgText(5)
End Sub


'*************************************************
'Add By Sindy 2018/7/19 數字要同於 Frmacc24c0
'  選擇請款收款統計
'
'*************************************************
Private Function Select2() As Boolean
'Dim ii As Integer
Dim straccSales As String  '2007/11/20 add by sonia
Dim str1P0Sales As String  '2007/11/20 add by sonia
Dim StrSQLa As String      '2007/11/26 add by sonia

Dim strSQL1k0 As String
Dim strSQLCP As String
Dim strSalesArea As String
Dim strAccSystem As String
Dim StrSQL0 As String
Dim strSQL1 As String      '2013/5/30  add by sonia
Dim strSQL2 As String      '2013/5/30  add by sonia
Dim StrSQL3 As String      '2013/5/30  add by sonia
Dim adoacc0y0 As New ADODB.Recordset
Dim adoaccrpt0205 As New ADODB.Recordset
Dim adoaccsum As New ADODB.Recordset
Dim douExchange As Double
   
   Screen.MousePointer = vbHourglass
   cnnConnection.Execute "delete from accrpt0205 where R21201='" & strUserNum & "'"
   StrSQL0 = "": strSQL1 = "": strSQL2 = "": StrSQL3 = "": straccSales = "": str1P0Sales = ""
   strSQL1k0 = "": strSQLCP = ""
   strAccSystem = "": straccSales = ""
   '2007/11/16 modify by sonia
   'If Len(frm040205.txt1(0)) <> 0 Then
   If frm040205.txt1(0) <> "ALL" Then
   '2007/11/16 end
      strSQL1k0 = " and a1k13 in (" & GetAddStr(frm040205.txt1(0)) & ") "
      strAccSystem = " and substr(ax214, 1, Length(ax214) - 9) in (" & GetAddStr(frm040205.txt1(0)) & ") "
   End If
   If Trim(frm040205.txt1(0)) <> "" Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(0) & frm040205.txt1(0) 'Add By Sindy 2010/9/28
   End If
   pub_QL05 = pub_QL05 & ";" & frm040205.Label1(1) & "收款" 'Add By Sindy 2010/9/28
   
   If Len(Trim(frm040205.txt1(2))) <> 0 Or Len(Trim(frm040205.txt1(3))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(2) & Trim(frm040205.txt1(2)) & "-" & Trim(frm040205.txt1(3)) 'Add By Sindy 2010/9/28
      StrSQL0 = " a0y02 >= " & Val(frm040205.txt1(2)) & ""    '2013/5/30 add by sonia
      StrSQL0 = StrSQL0 & " and a0y02 <= " & Val(frm040205.txt1(3)) & ""    '2013/5/30 add by sonia
      strSQL1 = " a0205 >= " & Val(frm040205.txt1(2)) & ""    '2013/5/30 add by sonia
      strSQL1 = strSQL1 & " and a0205 <= " & Val(frm040205.txt1(3)) & ""    '2013/5/30 add by sonia
      strSQL2 = " a1h02 >= " & Val(frm040205.txt1(2)) & ""    '2013/5/30 add by sonia
      strSQL2 = strSQL2 & " and a1h02 <= " & Val(frm040205.txt1(3)) & ""    '2013/5/30 add by sonia
      StrSQL3 = " a1p18 >= " & Val(frm040205.txt1(2)) & ""    '2013/5/30 add by sonia
      StrSQL3 = StrSQL3 & " and a1p18 <= " & Val(frm040205.txt1(3)) & ""    '2013/5/30 add by sonia
   End If
   
   '2007/11/26 add by sonia
   '國籍
   If Len(frm040205.txt1(4)) <> 0 Then
      strSQLCP = strSQLCP & " and fa10>='" & frm040205.txt1(4) & "' "
   End If
   If Len(frm040205.txt1(5)) <> 0 Then
      strSQLCP = strSQLCP & " and fa10<='" & frm040205.txt1(5) & "z' "
   End If
   '2007/11/26 end
   '案件性質
   If Len(Trim(frm040205.txt1(4))) <> 0 Or Len(Trim(frm040205.txt1(5))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(4) & Trim(frm040205.txt1(4)) & "-" & Trim(frm040205.txt1(5))  'Add By Sindy 2010/9/28
   End If
   If Len(frm040205.txt1(7)) <> 0 Then
      strSQLCP = strSQLCP & " and CP10>='" & frm040205.txt1(7) & "' "
   End If
   If Len(frm040205.txt1(8)) <> 0 Then
      strSQLCP = strSQLCP & " and CP10<='" & frm040205.txt1(8) & "' "
   End If
   If Len(Trim(frm040205.txt1(7))) <> 0 Or Len(Trim(frm040205.txt1(8))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(7) & Trim(frm040205.txt1(7)) & "-" & Trim(frm040205.txt1(8))   'Add By Sindy 2010/9/28
   End If
   'Add by Morgan 2003/12/04
   '智權人員
   If Len(frm040205.txt1(9)) <> 0 Then
      strSQLCP = strSQLCP & " and CP13||''='" & frm040205.txt1(9) & "' "
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(8) & Trim(frm040205.txt1(9)) & frm040205.lbl1   'Add By Sindy 2010/9/28
   End If
   '業務區
   If Len(frm040205.txt1(10)) <> 0 Then
      strSalesArea = strSalesArea & " and CP12||''>='" & frm040205.txt1(10) & "' "
   End If
   If Len(frm040205.txt1(11)) <> 0 Then
      strSalesArea = strSalesArea & " and CP12||''<='" & frm040205.txt1(11) & "' "
   End If
   If Len(frm040205.txt1(10)) <> 0 Or Len(frm040205.txt1(11)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm040205.Label1(9) & Trim(frm040205.txt1(10)) & "-" & Trim(frm040205.txt1(11))  'Add By Sindy 2010/9/28
   End If
   'End 2003/12/04
   '2007/11/20 add by sonia 非個人時再加印財務調整傳票
   If frm040205.txt1(9) = "" Then
      Select Case Mid(frm040205.txt1(10), 1, 2)
         Case "F3"
            straccSales = " and ax209='F4101' "
            str1P0Sales = " and a1p16='F4101' "
         Case "F2"
            'modify by sonia 2021/1/20 +F4104,F4105
            straccSales = " and ax209 in ('F4102','F4104','F4105') "
            str1P0Sales = " and a1p16 in ('F4102','F4104','F4105') "
         Case "F1"
            'modify by sonia 2021/1/20 +F4106,F4107
            straccSales = " and ax209 in ('F4103','F4106','F4107') "
            str1P0Sales = " and a1p16 in ('F4103','F4106','F4107') "
      End Select
      If Mid(frm040205.txt1(10), 1, 2) = "F1" And Mid(frm040205.txt1(11), 1, 2) = "F4" Then
            'modify by sonia 2021/1/20 +F4104~F4107
            straccSales = " and ax209 in ('F4101','F4102','F4103','F4104','F4105','F4106','F4107') "
            str1P0Sales = " and a1p16 in ('F4101','F4102','F4103','F4104','F4105','F4106','F4107') "
      End If
   Else
      straccSales = " and ax209='" & frm040205.txt1(9) & "' "
   End If
   '2007/11/20 end
   
   StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as FAname "
   '2007/11/12 modify by sonia 同一請款單計入最後收文之智權人員,另再加印財務調整傳票D096091227,再加舊系統請款單找不到收文記錄者D096090203
   'adoacc0y0.Open "select * from acc0z0, acc0y0, acc1k0, (select distinct cp01, cp13, cp60, CP12 from acc0z0, acc0y0, caseprogress, acc1k0 where a0z01 = a0y01 and a0z02 = cp60 (+) and a0z02 = a1k01 (+) and a0z04 <> 0" & strSalesMan & strSQL & ") new where a0z01 = a0y01 and a0z02 = a1k01 (+) and a0z02 = cp60 (+) and a0z04 <> 0" & strSalesMan & strSQL & " order by cp13", adoTaie, adOpenStatic, adLockReadOnly
   '2009/7/29 MODIFY BY SONIA FCL-10530於98/7/13收款有收文但因拆點數於FCL未出現(因為CP09 is null)故第三段修改
   '2010/9/13 MODIFY BY SONIA 發現第三段因為抓找不到收文記錄者應要加CP60 IS NULL的控制
   'Modify by Morgan 2011/4/11 ax207 改抓加總;另調整語法 cp09 in substr(new.cp,9,9)-->cp09(+)=substr(new.cp,9,9),a0202=ax202-->a0202=ax202(+)
   'modify by sonia 2013/5/30 加抵帳點數Z10200013(102/4/9)第四段,但抵帳之a1p21與國外收款之a1p21存的值不同故抵帳不做new.a0z04=a1p21(+),因抵帳A1P02='K'故第二段的'F'=a1p02(+)改掉,而第三段的'F'=a1p02(+)改為'F'=a1p02
   'adoacc0y0.Open "select cp13,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y04,a0z04,sum(ax207) ax207 from caseprogress, acc1p0, acc021, " & _
                  "(select max(cp05||cp09) cp,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y04,a0z04,a0z02,a0z01 from acc0z0, acc0y0, acc1k0, caseprogress where a0z01(+) = a0y01 " & strSql & " and a0z04 <> 0 " & _
                  "and a0z02=a1k01(+) and a0z02 = cp60 (+) " & strSalesArea & strSystem & " group by a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y04,a0z04,a0z02,a0z01) new " & _
                  "where cp09(+)=substr(new.cp,9,9) " & strSalesMan & " and new.a0z01=a1p04(+) and new.a0z04=a1p21(+) and substr(a1P05,1,1) in ('4','7') " & str1P0Sales & "and new.a1k13||new.a1k14||new.a1k15||new.a1k16=a1p17(+) " & _
                  "and a1p22=ax202(+) and a1p17=ax214(+) and substr(ax205,1,1) in ('4','7') and a1p03=ax203(+) " & _
                  " group by cp13,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y04,a0z04 " & _
                  "union " & _
                  "select ax209 cp13,'' a1k01,0 a1k09,substr(ax214, 1, Length(ax214) - 9) a1k13,substr(ax214, Length(ax214) - 8, 6) a1k14,substr(ax214, Length(ax214) - 2, 1) a1k15,substr(ax214, Length(ax214) - 1, 2) a1k16,0 a1k30,a0205 a0y02,1 a0y04,ax207-ax206 a0z04,ax207-ax206 ax207 " & _
                  "from acc020,acc021,acc1p0 where a0202=ax202(+) " & strSQL1 & " and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & _
                  "and a0202=a1p22(+) and 'F'=a1p02(+) and a1p04 is null " & _
                  "union " & _
                  "select ax209 cp13,'' a1k01,0 a1k09,substr(ax214, 1, Length(ax214) - 9) a1k13,substr(ax214, Length(ax214) - 8, 6) a1k14,substr(ax214, Length(ax214) - 2, 1) a1k15,substr(ax214, Length(ax214) - 1, 2) a1k16,0 a1k30,a0205 a0y02,1 a0y04,ax207-ax206 a0z04,ax207-ax206 ax207 " & _
                  "from acc020,acc021,acc1p0,acc0z0,acc1k0,caseprogress where a0202=ax202(+) " & strSQL1 & " and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & _
                  "and ax202=a1p22(+) and 'F'=a1p02(+) and ax214=a1p17(+) and ax209=a1p16(+) and a1p04=a0z01(+) and a1p21=a0z04(+) and a0z02=a1k01(+) and a1k01=cp60(+) and cp09 is null " & _
                  "order by cp13,a1k01,a1k13,a1k14,a1k15,a1k16", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Lydia 2017/10/18 拿掉 adoacc0y0.Open, 放在strexc(1)
   'modify by sonia 2021/1/20 加傳票公司別條件a0201=ax201(+)及ax201=a1p01(+)
   strExc(1) = "select cp13,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y01,a0y03,a0y04,a0z04,sum(ax207) ax207,cp01,cp10,a1k29" & _
               " from caseprogress, acc1p0, acc021, " & _
               "(select max(cp05||cp09) cp,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y01,a0y03,a0y04,a0z04,a0z02,a0z01,a1k29" & _
               " from acc0z0, acc0y0, acc1k0, caseprogress where " & StrSQL0 & " and a0z01(+) = a0y01 and a0z04 <> 0 " & _
               "and a0z02=a1k01(+) and a0z02 = cp60 (+) " & strSalesArea & strSQL1k0 & " group by a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y01,a0y03,a0y04,a0z04,a0z02,a0z01,a1k29) new " & _
               "where cp09(+)=substr(new.cp,9,9) " & strSQLCP & " and new.a0z01=a1p04(+) and new.a0z04=a1p21(+) and substr(a1P05,1,1) in ('4','7') " & str1P0Sales & "and new.a1k13||new.a1k14||new.a1k15||new.a1k16=a1p17(+) " & _
               "and a1p01=ax201(+) and a1p22=ax202(+) and a1p17=ax214(+) and substr(ax205,1,1) in ('4','7') and a1p03=ax203(+) group by cp13,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y01,a0y03,a0y04,a0z04,cp01,cp10,a1k29 " & _
               "union " & _
               "select ax209 cp13,'' a1k01,0 a1k09,substr(ax214, 1, length(ax214) - 9) a1k13,substr(ax214, length(ax214) - 8, 6) a1k14,substr(ax214, length(ax214) - 2, 1) a1k15,substr(ax214, length(ax214) - 1, 2) a1k16,0 a1k30,a0205 a0y02,'','NTD',1 a0y04,ax207-ax206 a0z04,ax207-ax206 ax207,'' cp01,'' cp10,'' a1k29 " & _
               "from acc020,acc021,(select * from acc1p0 where " & StrSQL3 & " and a1p02 in ('F','K')) where " & strSQL1 & " and a0201=ax201(+) and a0202=ax202(+) and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & _
               " and a0201=a1p01(+) and a0202=a1p22(+) and a1p04 is null " & _
               "union " & _
               "select ax209 cp13,'' a1k01,0 a1k09,substr(ax214, 1, length(ax214) - 9) a1k13,substr(ax214, length(ax214) - 8, 6) a1k14,substr(ax214, length(ax214) - 2, 1) a1k15,substr(ax214, length(ax214) - 1, 2) a1k16,0 a1k30,a0205 a0y02,a1p04,a1p19,1 a0y04,ax207-ax206 a0z04,ax207-ax206 ax207,'' cp01,'' cp10,'' a1k29 " & _
               "from acc020,acc021,acc1p0,acc0z0,acc1k0,caseprogress where " & strSQL1 & " and a0201=ax201(+) and a0202=ax202(+) and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & _
               " and ax201=a1p01(+) and ax202=a1p22(+) and 'F'=a1p02 and ax214=a1p17(+) and ax209=a1p16(+) and a1p04=a0z01(+) and a1p21=a0z04(+) and a0z02=a1k01(+) and a1k01=cp60(+) and cp09 is null " & _
               "union " & _
               "select cp13,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a1h02,a1h01,a1h03,a1h04,a1k08,sum(ax207) ax207,cp01,cp10,a1k29" & _
               " from caseprogress, acc1p0, acc021, " & _
               "(select max(cp05||cp09) cp,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a1h02,a1h01,a1h03,a1h04,a1k08,a1k29" & _
               " from acc1h0, acc1k0, caseprogress where " & strSQL2 & _
               " and a1h01=a1k17(+) and a1k01 = cp60 (+) " & strSalesArea & strSQL1k0 & " group by a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a1h02,a1h01,a1h03,a1h04,a1k08,a1k29) new " & _
               "where cp09(+)=substr(new.cp,9,9) " & strSQLCP & " and new.a1h01=a1p04(+) and substr(a1P05,1,1) in ('4','7') " & str1P0Sales & "and new.a1k13||new.a1k14||new.a1k15||new.a1k16=a1p17(+) " & _
               "and a1p01=ax201(+) and a1p22=ax202(+) and a1p17=ax214(+) and substr(ax205,1,1) in ('4','7') and a1p03=ax203(+)" & _
               " group by cp13,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a1h02,a1h01,a1h03,a1h04,a1k08,cp01,cp10,a1k29" & _
               " order by cp13,a1k01,a1k13,a1k14,a1k15,a1k16"
   'Add By Sindy 2018/7/19 案件主檔/代理人國籍條件
   If strBaseTable <> "" Then
       strExc(1) = UCase(strExc(1))
       strSql = "select a.*,vc08,DECODE(vc07,'000',CPM03,CPM04) AS cpName," & StrSQLa & " from (" & Mid(strExc(1), 1, InStr(strExc(1), "ORDER BY") - 1) & ") a " & strBaseTable & _
                " where " & strCaseNa & strSQLCP & " order by cp13,a1k01,a1k13,a1k14,a1k15,a1k16"
   Else
       strSql = strExc(1)
   End If
   '2018/7/19 END
   adoacc0y0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly 'Added by Lydia 2017/10/18
   If adoacc0y0.RecordCount = 0 Then
'      strCon10 = MsgText(602)
'      adoacc0y0.Close
'      MsgBox MsgText(28), , MsgText(5)
      Exit Function
   End If
   adoaccrpt0205.CursorLocation = adUseClient
   adoaccrpt0205.Open "select * from accrpt0205 where R21201='" & strUserNum & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Do While adoacc0y0.EOF = False
      adoaccrpt0205.AddNew
      'Add by Morgan 2003/12/02
      If IsNull(adoacc0y0.Fields("A1k01").Value) Then
         adoaccrpt0205.Fields("r21213").Value = Null
      Else
         adoaccrpt0205.Fields("r21213").Value = adoacc0y0.Fields("A1k01").Value
      End If
      'End 2003/12/02
      adoaccrpt0205.Fields("r21201").Value = strUserNum
      If IsNull(adoacc0y0.Fields("cp13").Value) Then
         adoaccrpt0205.Fields("r21202").Value = Null
         adoaccrpt0205.Fields("r21211").Value = Null
      Else
         adoaccrpt0205.Fields("r21202").Value = adoacc0y0.Fields("cp13").Value
         adoaccrpt0205.Fields("r21211").Value = StaffQuery(adoacc0y0.Fields("cp13").Value)
      End If
      If IsNull(adoacc0y0.Fields("a1k13").Value) Then
         adoaccrpt0205.Fields("r21203").Value = Null
      Else
         adoaccrpt0205.Fields("r21203").Value = adoacc0y0.Fields("a1k13").Value
         If IsNull(adoacc0y0.Fields("a1k14").Value) = False Then
            adoaccrpt0205.Fields("r21203").Value = adoaccrpt0205.Fields("r21203").Value & "-" & adoacc0y0.Fields("a1k14").Value
         End If
         If IsNull(adoacc0y0.Fields("a1k15").Value) = False Then
            adoaccrpt0205.Fields("r21203").Value = adoaccrpt0205.Fields("r21203").Value & "-" & adoacc0y0.Fields("a1k15").Value
         End If
         If IsNull(adoacc0y0.Fields("a1k16").Value) = False Then
            adoaccrpt0205.Fields("r21203").Value = adoaccrpt0205.Fields("r21203").Value & "-" & adoacc0y0.Fields("a1k16").Value
         End If
      End If
      If IsNull(adoacc0y0.Fields("a0y02").Value) Then
         adoaccrpt0205.Fields("r21204").Value = Null
      Else
         adoaccrpt0205.Fields("r21204").Value = adoacc0y0.Fields("a0y02").Value
      End If
      If IsNull(adoacc0y0.Fields("a0y04").Value) Then
         douExchange = 0
      Else
         douExchange = adoacc0y0.Fields("a0y04").Value
      End If
      If IsNull(adoacc0y0.Fields("a0z04").Value) Then
         adoaccrpt0205.Fields("r21208").Value = 0
      Else
         adoaccrpt0205.Fields("r21208").Value = Val(Format(Val(adoacc0y0.Fields("a0z04").Value) * douExchange, FAmount))
      End If
      '2007/11/8 modify by sonia 已收點數應扣除規費,第一次收款即應扣規費
      'adoaccrpt0205.Fields("r21209").Value = Val(Format(Val(adoaccrpt0205.Fields("r21208").Value) / 1000, FAmount))
      'adoaccrpt0205.Fields("r21212").Value = Val(adoaccrpt0205.Fields("r21208").Value) / 1000
      If adoaccrpt0205.Fields("r21208").Value = Val(adoacc0y0.Fields("A1k30").Value) Then
         adoaccrpt0205.Fields("r21209").Value = Val(Format(Val(Val(adoaccrpt0205.Fields("r21208").Value) - Val(adoacc0y0.Fields("a1k09").Value)) / 1000, FAmount)) '0.00000
      Else
         adoaccrpt0205.Fields("r21209").Value = Val(Format(Val(adoaccrpt0205.Fields("r21208").Value) / 1000, FAmount)) '0.00000
      End If
      '2007/11/8 end
      adoaccsum.CursorLocation = adUseClient
'      adoaccsum.Open "select * from acc1p0 where a1p18 in (select min(a1p18) from acc1p0 where a1p01 = '1' and a1p17 = '" & adoacc0y0.Fields("a1k13").Value & adoacc0y0.Fields("a1k14").Value & adoacc0y0.Fields("a1k15").Value & adoacc0y0.Fields("a1k16").Value & "' and a1p05 = '6130') and a1p01 = '1' and a1p17 = '" & adoacc0y0.Fields("a1k13").Value & adoacc0y0.Fields("a1k14").Value & adoacc0y0.Fields("a1k15").Value & adoacc0y0.Fields("a1k16").Value & "' and a1p05 = '6130'", adoTaie, adOpenStatic, adLockReadOnly
      '2007/11/13 modify by sonia 該請款單有201新案翻譯才抓
      'adoaccsum.Open "select ax206 from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and a0205 in (select min(a0205) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax214 = '" & adoacc0y0.Fields("a1k13").Value & adoacc0y0.Fields("a1k14").Value & adoacc0y0.Fields("a1k15").Value & adoacc0y0.Fields("a1k16").Value & "' and ax205 = '6130') and ax214 = '" & adoacc0y0.Fields("a1k13").Value & adoacc0y0.Fields("a1k14").Value & adoacc0y0.Fields("a1k15").Value & adoacc0y0.Fields("a1k16").Value & "' and ax205 = '6130'", adoTaie, adOpenStatic, adLockReadOnly
      adoaccrpt0205.Fields("r21210").Value = 0
      If Not IsNull(adoacc0y0.Fields("A1k01").Value) Then
         adoaccsum.Open "select ax206 from acc021, acc020, caseprogress where ax201 = a0201 and ax202 = a0202 and a0205 in (select min(a0205) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax214 = '" & adoacc0y0.Fields("a1k13").Value & adoacc0y0.Fields("a1k14").Value & adoacc0y0.Fields("a1k15").Value & adoacc0y0.Fields("a1k16").Value & "' and ax205 = '6130') and ax214 = '" & adoacc0y0.Fields("a1k13").Value & adoacc0y0.Fields("a1k14").Value & adoacc0y0.Fields("a1k15").Value & adoacc0y0.Fields("a1k16").Value & "' and ax205 = '6130' " & _
                        "and '" & adoacc0y0.Fields("A1k01").Value & "'=cp60(+) and (cp10='201' or cp10='927') ", adoTaie, adOpenStatic, adLockReadOnly
         If adoaccsum.RecordCount <> 0 Then
            If Not IsNull(adoaccsum.Fields(0).Value) Then
               adoaccrpt0205.Fields("r21210").Value = adoaccsum.Fields(0).Value
            End If
         End If
         adoaccsum.Close
      End If
      '2007/11/13 add by sonia 再印財務點數
      adoaccrpt0205.Fields("r21212").Value = Val(Format(Val(adoacc0y0.Fields("ax207").Value) / 1000, FAmount)) '0.00000
      '2007/11/13 end
      'Add By Sindy 2018/7/20
      '規費
      If adoaccrpt0205.Fields("r21208").Value = Val(adoacc0y0.Fields("A1k30").Value) Then
         adoaccrpt0205.Fields("R21206").Value = Val(Format(adoacc0y0.Fields("a1k09").Value, FAmount))
      Else
         adoaccrpt0205.Fields("R21206").Value = 0
      End If
      '外幣金額
      If IsNull(adoacc0y0.Fields("a0z04").Value) Then
         adoaccrpt0205.Fields("R21207").Value = 0
      Else
         adoaccrpt0205.Fields("R21207").Value = Val(Format(adoacc0y0.Fields("a0z04").Value, FAmount))
      End If
      '案件性質
      If IsNull(adoacc0y0.Fields("cpName").Value) Then
         adoaccrpt0205.Fields("R21216").Value = Null
      Else
         adoaccrpt0205.Fields("R21216").Value = adoacc0y0.Fields("cpName").Value
      End If
      '單據編號
      If IsNull(adoacc0y0.Fields("A0Y01").Value) Then
         adoaccrpt0205.Fields("R21217").Value = Null
      Else
         adoaccrpt0205.Fields("R21217").Value = adoacc0y0.Fields("A0Y01").Value
      End If
      '幣別
      If IsNull(adoacc0y0.Fields("A0Y03").Value) Then
         adoaccrpt0205.Fields("R21218").Value = Null
      Else
         adoaccrpt0205.Fields("R21218").Value = adoacc0y0.Fields("A0Y03").Value
      End If
      '結清
      If IsNull(adoacc0y0.Fields("A1K29").Value) Then
         adoaccrpt0205.Fields("R21219").Value = Null
      Else
         adoaccrpt0205.Fields("R21219").Value = adoacc0y0.Fields("A1K29").Value
      End If
      '是否閉卷
      If "" & adoacc0y0.Fields("vc08").Value = "Y" Then
         adoaccrpt0205.Fields("R21214").Value = "*"
      Else
         adoaccrpt0205.Fields("R21214").Value = Null
      End If
      '2018/7/20 END
      adoaccrpt0205.UpdateBatch
      adoacc0y0.MoveNext
   Loop
   adoacc0y0.Close
   adoaccrpt0205.Close
   Screen.MousePointer = vbDefault
End Function

Private Sub Form_Load()
    stST05 = PUB_GetST05(strUserNum)
   MoveFormToCenter Me
   bolToEndByNick = False
   txt1(0) = StrStartSystemByNick
   'add by nickc 2007/04/10 陳經理的請作單，當承辦進來時，鎖住智權人員
   'edit by nickc 2008/01/18
   'If PUB_GetST05(strUserNum) = "15" Then
   If stST05 = "15" Then
      txt1(9).Enabled = False
      txt1(9) = strUserNum
      txt1_LostFocus 9
   'add by nickc 2008/01/18
   'Modify By Sindy 2009/05/21 80030.洪琬姿及78011.葉易雲不受限制
   ElseIf (stST05 = "21" Or stST05 = "26" Or stST05 = "28") And strUserNum <> "80030" And strUserNum <> "78011" Then '主管僅鎖業務區
      txt1(10).Enabled = False
      txt1(11).Enabled = False
      txt1(9) = strUserNum
      txt1_LostFocus 9
   'Modify By Sindy 2009/05/21 80030.洪琬姿及78011.葉易雲不受限制
   ElseIf (stST05 >= "21" And stST05 <= "29") And strUserNum <> "80030" And strUserNum <> "78011" Then '非主管鎖業務區智權人員
      txt1(10).Enabled = False
      txt1(11).Enabled = False
      txt1(9).Enabled = False
      txt1(9) = strUserNum
      txt1_LostFocus 9
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040205 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   If Index = 11 Then
      If txt1(Index) = "" Then
         txt1(Index) = txt1(Index - 1)
      End If
   End If
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   'Add by Morgan 2003/12/04
   '智權人員
   If Len(txt1(9)) <> 0 Then
      lbl1.Caption = GetStaffName(txt1(9), True)
   Else
      lbl1.Caption = ""
   End If
   'End 2003/12/04

   'Add By Cheng 2003/01/03
   If Me.txt1(Index).Text = "" Then Exit Sub
   Select Case Index
      Case 0
      '2007/11/16 modify by sonia 因外商收法務案故取消使用權限控制
      '      If Len(Trim(txt1(0))) <> 0 Then
      '            strTemp = Split(GetSystemKindByNick, ",")
      '            strTemp1 = Split(txt1(0), ",")
      '            For i = 0 To UBound(strTemp1)
      '                s = 0
      '                For j = 0 To UBound(strTemp)
      '                    If strTemp1(i) = strTemp(j) Then
      '                        s = 1
      '                    End If
      '                Next j
      '                If s = 0 Then
      '                    s = MsgBox(strUserNum + " 沒有 " + strTemp1(i) + " 的使用權限 ", , "USER 權限不足!!!")
      '                    txt1(0).SetFocus
      '                    txt1(0).SelStart = 0
      '                    txt1(0).SelLength = Len(txt1(0))
      '                    Exit Sub
      '                End If
      '            Next i
      '      End If
            If txt1(0) <> "ALL" Then
               'Add By Sindy 2013/1/15
               txt1(0) = Trim(txt1(0))
               If Right(txt1(0), 1) = "," Then
                  txt1(0) = Left(txt1(0), Len(txt1(0)) - 1)
               End If
               If Left(txt1(0), 1) = "," Then
                  txt1(0) = Right(txt1(0), Len(txt1(0)) - 1)
               End If
               '2013/1/15 End
               strTemp1 = Split(txt1(0), ",")
               For i = 0 To UBound(strTemp1)
                  If CheckSys(strTemp1(i)) = "" Then
                     MsgBox "系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
                     txt1(0).SetFocus
                     txt1(0).SelStart = 0
                     txt1(0).SelLength = Len(txt1(0))
                     Exit Sub
                  End If
               Next i
            End If
      '2007/11/16 end
      Case 1
         Select Case txt1(1)
            Case "1", "2"
            Case Else
               s = MsgBox("查詢別只能 1 或 2 !!", , "USER 輸入錯誤")
               txt1(1).SetFocus
               txt1(1).SelStart = 0
               txt1(1).SelLength = Len(txt1(1))
               Exit Sub
         End Select
      Case 2, 3 '日期起, 迄
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            'add by nickc 2007/04/18 解決 BUG ，會與下面的 msg 互跳
            Exit Sub
            
         End If
         If Index = 3 Then
            If Not nickChgRan(txt1(2), txt1(3), "日期") Then
               txt1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
          End If
      Case 5
         If Not nickChgRan(txt1(4), txt1(5), "國籍") Then
            txt1(4).SetFocus
            txt1_GotFocus 4
            Exit Sub
         End If
      Case 6
         Select Case txt1(6)
            Case "1", "2"
            Case Else
               s = MsgBox("查詢內容只能 1 或 2 !!", , "USER 輸入錯誤")
               txt1(6).SetFocus
               txt1(6).SelStart = 0
               txt1(6).SelLength = Len(txt1(6))
               Exit Sub
         End Select
      Case 8
         If Not nickChgRan(txt1(7), txt1(8), "案件性質") Then
            txt1(7).SetFocus
            txt1_GotFocus 7
            Exit Sub
         End If
          
      'Add by Morgan 2003/12/04
      '業務區
      Case 11
         If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
         End If
      'End 2003/12/04
      
      Case Else
   End Select
   
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
      'add by nickc 2008/01/18
      If Index = 9 And Trim(txt1(9)) <> "" Then
        'Modify By Sindy 2009/05/21 80030.洪琬姿及78011.葉易雲不受限制
        If (stST05 = "21" Or stST05 = "26" Or stST05 = "28") And strUserNum <> "80030" And strUserNum <> "78011" Then
             If PUB_GetStaffST16(strUserNum) <> PUB_GetStaffST16(txt1(9)) Then
                  MsgBox "僅可以輸入相同組別的智權人員！", vbExclamation, "操作錯誤！"
                  txt1(9).SetFocus
                  txt1_GotFocus 9
                  Cancel = True
                  Exit Sub
             End If
        End If
      End If
End Sub
