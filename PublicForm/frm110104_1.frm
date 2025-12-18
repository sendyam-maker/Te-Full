VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm110104_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "更換FC代理人作業"
   ClientHeight    =   5790
   ClientLeft      =   1350
   ClientTop       =   1605
   ClientWidth     =   8535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8535
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   4005
      Left            =   60
      TabIndex        =   26
      Top             =   1080
      Width           =   8445
      Begin VB.CheckBox Check4 
         Caption         =   "清除案件聯絡人資料"
         Height          =   300
         Left            =   3765
         TabIndex        =   7
         Top             =   720
         Width           =   4000
      End
      Begin VB.TextBox txtCaseField 
         Height          =   300
         Index           =   3
         Left            =   1470
         MaxLength       =   9
         TabIndex        =   9
         Top             =   1380
         Width           =   1212
      End
      Begin VB.TextBox txtCaseField 
         Height          =   300
         Index           =   1
         Left            =   1485
         MaxLength       =   9
         TabIndex        =   3
         Top             =   30
         Width           =   1212
      End
      Begin VB.TextBox txtCaseField 
         Height          =   300
         Index           =   2
         Left            =   1485
         MaxLength       =   9
         TabIndex        =   4
         Top             =   345
         Width           =   1212
      End
      Begin VB.CheckBox Check1 
         Caption         =   "含閉卷或銷卷案件"
         Height          =   300
         Left            =   465
         TabIndex        =   5
         Top             =   720
         Value           =   1  '核取
         Width           =   2500
      End
      Begin VB.CheckBox Check2 
         Caption         =   "彼所案號清除(未勾選時會保留)"
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   465
         TabIndex        =   6
         Top             =   990
         Width           =   3000
      End
      Begin VB.CheckBox Check3 
         Caption         =   "案件聯絡人同時更改(請輸入下方聯絡人資料)"
         Height          =   300
         Left            =   3765
         TabIndex        =   8
         Top             =   990
         Width           =   4000
      End
      Begin MSForms.TextBox Text2 
         Height          =   300
         Index           =   2
         Left            =   1470
         TabIndex        =   12
         Top             =   2331
         Width           =   6810
         VariousPropertyBits=   671105051
         Size            =   "12012;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   300
         Index           =   3
         Left            =   1470
         TabIndex        =   13
         Top             =   2648
         Width           =   3700
         VariousPropertyBits=   671105051
         Size            =   "6526;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   300
         Index           =   4
         Left            =   1470
         TabIndex        =   14
         Top             =   2965
         Width           =   6800
         VariousPropertyBits=   671105051
         Size            =   "11994;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   300
         Index           =   5
         Left            =   1470
         TabIndex        =   15
         Top             =   3282
         Width           =   6810
         VariousPropertyBits=   671105051
         Size            =   "12012;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   300
         Index           =   0
         Left            =   1470
         TabIndex        =   10
         Top             =   1697
         Width           =   3700
         VariousPropertyBits=   671105051
         Size            =   "6526;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   300
         Index           =   1
         Left            =   1470
         TabIndex        =   11
         Top             =   2014
         Width           =   6800
         VariousPropertyBits=   671105051
         Size            =   "11994;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   300
         Index           =   6
         Left            =   1470
         TabIndex        =   16
         Top             =   3600
         Width           =   6810
         VariousPropertyBits=   671105051
         Size            =   "12012;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCustomer 
         Height          =   255
         Left            =   2730
         TabIndex        =   40
         Top             =   368
         Width           =   5640
         VariousPropertyBits=   27
         Size            =   "9948;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label NewAgent 
         Height          =   255
         Left            =   2730
         TabIndex        =   39
         Top             =   1403
         Width           =   5640
         VariousPropertyBits=   27
         Size            =   "9948;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblAgent 
         Height          =   255
         Left            =   2730
         TabIndex        =   38
         Top             =   53
         Width           =   5640
         VariousPropertyBits=   27
         Size            =   "9948;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "新FC代理人："
         Height          =   180
         Left            =   120
         TabIndex        =   36
         Top             =   1395
         Width           =   1110
      End
      Begin VB.Label Label4 
         Caption         =   "代理人　："
         Height          =   180
         Left            =   465
         TabIndex        =   35
         Top             =   75
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "申請人　："
         Height          =   180
         Left            =   465
         TabIndex        =   34
         Top             =   405
         Width           =   975
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(中)："
         Height          =   180
         Left            =   120
         TabIndex        =   33
         Top             =   1718
         Width           =   1110
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(英)："
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   2041
         Width           =   1110
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(中)："
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   2687
         Width           =   1110
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(英)："
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   3010
         Width           =   1110
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(日)："
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   3333
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門(日)："
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   3660
         Width           =   1500
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(日)："
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   2364
         Width           =   1110
      End
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "本所案號："
      CausesValidation=   0   'False
      Height          =   180
      Index           =   1
      Left            =   270
      TabIndex        =   17
      Top             =   5310
      Width           =   1200
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "系統類別："
      CausesValidation=   0   'False
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   2
      Top             =   780
      Width           =   1200
   End
   Begin VB.TextBox textTM02 
      Height          =   300
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   19
      Top             =   5235
      Width           =   1092
   End
   Begin VB.TextBox textTM02_2 
      Height          =   300
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5235
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox txtSystem 
      Height          =   300
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   18
      Top             =   5235
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   300
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   21
      Top             =   5235
      Width           =   372
   End
   Begin VB.TextBox textTM04 
      Height          =   300
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   22
      Top             =   5235
      Width           =   732
   End
   Begin VB.TextBox txtCaseField 
      Height          =   300
      Index           =   4
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   0
      Top             =   360
      Width           =   700
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7365
      TabIndex        =   24
      Top             =   96
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6540
      TabIndex        =   23
      Top             =   96
      Width           =   800
   End
   Begin VB.TextBox txtCaseField 
      Height          =   300
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   705
      Width           =   4935
   End
   Begin MSForms.Label lblSName 
      Height          =   255
      Left            =   2340
      TabIndex        =   37
      Top             =   360
      Width           =   1125
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "1984;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   540
      TabIndex        =   25
      Top             =   420
      Width           =   900
   End
End
Attribute VB_Name = "frm110104_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/17 改成Form2.0 ;lblSName、lblAgent、lblCustomer、NewAgent、Text2(index)
'Memo By Sonia 2012/12/17 智權人員欄已修改
'2011/5/30 CREATE BY SONIA
Option Explicit

Dim strSystemKind As String '使用者可使用之系統類別
Dim strTemp1 As Variant, strTemp2 As Variant
Dim strTM01 As String
Dim strTM02 As String
Dim strTM03 As String
Dim strTM04 As String

Private Sub Check3_Click()
   If Check3.Value = 1 Then
      Check4.Value = 0
   End If
End Sub

Private Sub Check4_Click()
   If Check4.Value = 1 Then
      Check3.Value = 0
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim varSaveCursor, i As Integer
Dim Cancel As Boolean
Dim intCaseKind As Integer
Dim AdoRs As ADODB.Recordset
   
   Select Case Index
      Case 0 '確定
         'Added by Lydia 2021/09/17 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
         If PUB_ChkUniText(Me, , True, "TextBox") = False Then
              Exit Sub
         End If
         'end 2021/09/17
         
         varSaveCursor = Screen.MousePointer
         Screen.MousePointer = vbHourglass
         
         '整批
         If optChoose(0).Value = True Then
            If txtCaseField(1) = "" And txtCaseField(2) = "" Then
               MsgBox "代理人和申請人至少輸入一項 !!!", vbExclamation + vbOKOnly
               Screen.MousePointer = varSaveCursor
               txtCaseField(1).SetFocus
               Exit Sub
            'Modify By Sindy 2022/7/26 外商阿蓮提,因有同一申請人委辦二個以上代理人的情形，故請設下列提醒
            Else
               If txtCaseField(1) <> "" And txtCaseField(2) = "" Then
                  If MsgBox("是否指定申請人？", vbInformation + vbYesNo + vbDefaultButton1, "電腦中心") = vbYes Then
                     Screen.MousePointer = varSaveCursor
                     txtCaseField(2).SetFocus
                     Exit Sub
                  End If
               ElseIf txtCaseField(2) <> "" And txtCaseField(1) = "" Then
                  If MsgBox("是否指定代理人？", vbInformation + vbYesNo + vbDefaultButton1, "電腦中心") = vbYes Then
                     Screen.MousePointer = varSaveCursor
                     txtCaseField(1).SetFocus
                     Exit Sub
                  End If
               End If
               '2022/7/26 END
            End If
            
            For i = 0 To 4
               If txtCaseField(i).Enabled Then
                  If CheckKeyIn(i) <> 1 Then
                     txtCaseField(i).SetFocus
                     Call txtCaseField_GotFocus(i)
                     Screen.MousePointer = varSaveCursor
                     Exit Sub
                  End If
               End If
            Next
            If Check3.Value = 1 Then
               If Text2(0).Text = "" And Text2(1).Text = "" And Text2(2).Text = "" Then
                  MsgBox "請輸入聯絡人1 !!!", vbExclamation + vbOKOnly
                  Screen.MousePointer = varSaveCursor
                  Text2(0).SetFocus
                  Exit Sub
               End If
            End If
            If txtCaseField(1) = txtCaseField(3) Then
               MsgBox "代理人和新FC代理人不可相同 !!!", vbExclamation + vbOKOnly
               Screen.MousePointer = varSaveCursor
               txtCaseField(3).SetFocus
               Exit Sub
            End If
         '單筆修改
         Else
            Cancel = False
            Call txtSystem_Validate(Cancel)
            If Cancel = True Then
               Screen.MousePointer = varSaveCursor
               txtSystem.SetFocus
               Exit Sub
            End If
            If txtCaseField(4).Enabled Then
               If CheckKeyIn(4) <> 1 Then
                  txtCaseField(4).SetFocus
                  Call txtCaseField_GotFocus(4)
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
               End If
            End If
            If CheckKeyIn(5) <> 1 Then
               txtSystem.SetFocus
               txtSystem_GotFocus
               Screen.MousePointer = varSaveCursor
               Exit Sub
            End If
         End If
         
         If optChoose(1).Value Then
            '外專人員操作且系統別為P,PS,CFP,CPS時,要過濾案件必須曾有F2字頭收文
            'Modify by Amy 2017/06/27 若下一程序未發文且有備註則彈訊息
            If Left(Pub_StrUserSt03, 2) = "F2" Then
               If (strTM01 = "P" Or strTM01 = "PS" Or strTM01 = "CFP" Or strTM01 = "CPS") Then
                    strSql = "select cp09 from caseprogress where cp01='" & strTM01 & "' and cp02='" & strTM02 & "' and cp03='" & strTM03 & "' and cp04='" & strTM04 & "' and substr(cp12,1,2)='F2'"
                    intI = 1
                    Set AdoRs = ClsLawReadRstMsg(intI, strSql)
                    If intI = 0 Then
                       MsgBox "非外專案件, 無操作權限!!!", vbExclamation + vbOKOnly
                       Screen.MousePointer = varSaveCursor
                       Exit Sub
                    End If
               End If
               If bolNP15NotNull = True Then
                  MsgBox "下一程序有備註,請確認新代是否延用指示!!!"
               End If
            End If
            'end 2017/06/27
            
            If ClsPDGetSystemKind(strTM01, intCaseKind) Then
               If intCaseKind = 專利 Then
                  frm110104_3.SetData 0, strTM01, True
                  frm110104_3.SetData 1, strTM02, False
                  frm110104_3.SetData 2, strTM03, False
                  frm110104_3.SetData 3, strTM04, False
                  Me.Hide
                  frm110104_3.Show
                  frm110104_3.QueryData
               Else
                  frm110104_4.SetData 0, strTM01, True
                  frm110104_4.SetData 1, strTM02, False
                  frm110104_4.SetData 2, strTM03, False
                  frm110104_4.SetData 3, strTM04, False
                  Me.Hide
                  frm110104_4.Show
                  frm110104_4.QueryData
               End If
            Else
               Screen.MousePointer = varSaveCursor
               Exit Sub
            End If
         Else
            GetData
            intI = 0
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               frm110104_2.Show
               Me.Hide
'            Else
'               ShowNoData
            End If
         End If
         Screen.MousePointer = varSaveCursor
      Case 1 '結束
         Unload Me
   End Select
   
   Set AdoRs = Nothing
End Sub

Private Sub GetData()
Dim i As Integer, intCaseKind As Integer
Dim strSQLP As String, strHadP As String
Dim strSQLT As String, strHadT As String
Dim strSQLL As String, strHadL As String
Dim strSQLS As String, strHadS As String
Dim SqlCUname As String, SqlFAname As String
   
   '系統別
   strTemp1 = Split(UCase(txtCaseField(0)), ",")
   For i = 0 To UBound(strTemp1)
      If ClsPDGetSystemKind(CStr(strTemp1(i)), intCaseKind) Then
         Select Case intCaseKind
            Case 專利
               strHadP = strHadP & strTemp1(i) & ","
            Case 商標
               strHadT = strHadT & strTemp1(i) & ","
            Case 法務
               strHadL = strHadL & strTemp1(i) & ","
            Case Else '服務
               strHadS = strHadS & strTemp1(i) & ","
         End Select
      End If
   Next i
   If strHadP <> "" Then
      strHadP = Replace(Left(strHadP, Len(strHadP) - 1), ",", "','")
      strSQLP = " PA01 in('" & strHadP & "')"
   End If
   If strHadT <> "" Then
      strHadT = Replace(Left(strHadT, Len(strHadT) - 1), ",", "','")
      strSQLT = " TM01 in('" & strHadT & "')"
   End If
   If strHadL <> "" Then
      strHadL = Replace(Left(strHadL, Len(strHadL) - 1), ",", "','")
      strSQLL = " LC01 in('" & strHadL & "')"
   End If
   If strHadS <> "" Then
      strHadS = Replace(Left(strHadS, Len(strHadS) - 1), ",", "','")
      strSQLS = " SP01 in('" & strHadS & "')"
   End If
   '代理人
   If txtCaseField(1) <> "" Then
      If strHadP <> "" Then strSQLP = strSQLP & " AND PA75='" & txtCaseField(1) & "'"
      If strHadT <> "" Then strSQLT = strSQLT & " AND TM44='" & txtCaseField(1) & "'"
      If strHadL <> "" Then strSQLL = strSQLL & " AND LC22='" & txtCaseField(1) & "'"
      If strHadS <> "" Then strSQLS = strSQLS & " AND SP26='" & txtCaseField(1) & "'"
   End If
   '申請人
   If txtCaseField(2) <> "" Then
      If strHadP <> "" Then strSQLP = strSQLP & " AND PA26='" & txtCaseField(2) & "'"
      If strHadT <> "" Then strSQLT = strSQLT & " AND TM23='" & txtCaseField(2) & "'"
      If strHadL <> "" Then strSQLL = strSQLL & " AND LC11='" & txtCaseField(2) & "'"
      If strHadS <> "" Then strSQLS = strSQLS & " AND SP08='" & txtCaseField(2) & "'"
   End If
   '不含閉卷或銷卷案件
   If Check1.Value = 0 Then
      If strHadP <> "" Then strSQLP = strSQLP & " AND PA57 IS NULL AND NVL(PA108,0)=0"
      If strHadT <> "" Then strSQLT = strSQLT & " AND TM29 IS NULL AND NVL(TM57,0)=0"
      If strHadL <> "" Then strSQLL = strSQLL & " AND LC08 IS NULL AND NVL(LC34,0)=0"
      If strHadS <> "" Then strSQLS = strSQLS & " AND SP15 IS NULL AND NVL(SP61,0)=0"
   End If
   
   'Add By Sindy 2014/12/2
   '依操作者的部門判斷名稱抓法:
   'FXX者帶英->中->日
   If Left(Pub_StrUserSt03, 1) = "F" Then
      SqlCUname = "DECODE(CU05,NULL,NVL(CU04,CU06),CU05||' '||CU88||' '||CU89||' '||CU90)"
      SqlFAname = "DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)"
   '非FXX者帶中->英->日
   Else
      SqlCUname = "DECODE(CU04,NULL,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06),CU04)"
      SqlFAname = "DECODE(FA04,NULL,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06),FA04)"
   End If
   '2014/12/2 END
   
   '組查詢sql
   strExc(0) = ""
   If strHadP <> "" Then
      strExc(0) = strExc(0) & " union " & _
                  "SELECT ' ',decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●'),PA11,FA01||FA02||" & SqlFAname & ",CU01||CU02||" & SqlCUname & ",PA77,sqldatet(PA58),sqldatet(PA108),nvl(PA05,nvl(PA06,PA07)),PA01,PA02,PA03,PA04 FROM PATENT,fagent,customer WHERE" & strSQLP & _
                  " AND substr(PA75,1,8)=fa01(+) AND substr(PA75,9,1)=fa02(+)" & _
                  " AND substr(PA26,1,8)=cu01(+) AND substr(PA26,9,1)=cu02(+)"
      If Left(Pub_StrUserSt03, 2) = "F2" Then
         strExc(0) = strExc(0) & " AND ((pa01 in('P','CFP') and exists (select cp09 from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and substr(cp12,1,2)='F2')) or pa01 not in('P','CFP'))"
      End If
   End If
   If strHadT <> "" Then
      strExc(0) = strExc(0) & " union " & _
                  "SELECT ' ',decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●'),nvl(TM15,TM12),FA01||FA02||" & SqlFAname & ",CU01||CU02||" & SqlCUname & ",TM45,sqldatet(TM30),sqldatet(TM57),TM05,TM01,TM02,TM03,TM04 FROM TRADEMARK,fagent,customer WHERE" & strSQLT & _
                  " AND substr(TM44,1,8)=fa01(+) AND substr(TM44,9,1)=fa02(+)" & _
                  " AND substr(TM23,1,8)=cu01(+) AND substr(TM23,9,1)=cu02(+)"
   End If
   If strHadL <> "" Then
      strExc(0) = strExc(0) & " union " & _
                  "SELECT ' ',LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●'),'',FA01||FA02||" & SqlFAname & ",CU01||CU02||" & SqlCUname & ",LC23,sqldatet(LC09),sqldatet(LC34),nvl(LC05,nvl(LC06,LC07)),LC01,LC02,LC03,LC04 FROM LAWCASE,fagent,customer WHERE" & strSQLL & _
                  " AND substr(LC22,1,8)=fa01(+) AND substr(LC22,9,1)=fa02(+)" & _
                  " AND substr(LC11,1,8)=cu01(+) AND substr(LC11,9,1)=cu02(+)"
   End If
   If strHadS <> "" Then
      strExc(0) = strExc(0) & " union " & _
                  "SELECT ' ',SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●'),SP11,FA01||FA02||" & SqlFAname & ",CU01||CU02||" & SqlCUname & ",SP27,sqldatet(SP16),sqldatet(SP61),nvl(SP05,nvl(SP06,SP07)),SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE,fagent,customer WHERE" & strSQLS & _
                  " AND substr(SP26,1,8)=fa01(+) AND substr(SP26,9,1)=fa02(+)" & _
                  " AND substr(SP08,1,8)=cu01(+) AND substr(SP08,9,1)=cu02(+)"
      If Left(Pub_StrUserSt03, 2) = "F2" Then
         strExc(0) = strExc(0) & " AND ((sp01 in('PS','CPS') and exists (select cp09 from caseprogress where cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 and substr(cp12,1,2)='F2')) or sp01 not in('PS','CPS'))"
      End If
   End If
   strExc(0) = Mid(strExc(0), 7) & " order by 10,11,12,13"
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   optChoose(0).Value = True
   ClsPDGetGroupSystemKind strGroup, strSystemKind
   If Left(Pub_StrUserSt03, 2) = "F2" Then
      strSystemKind = strSystemKind & ",P,PS,CFP,CPS"
   Else
      strSystemKind = Replace(strSystemKind, ",LA", "")
   End If
   Cleartxt
   
   'Modify By Sindy 2022/7/26 限制F1X人員操作才預設
   If Left(Pub_StrUserSt03, 2) = "F1" Then
      Check2.Value = 1 '彼所案號清除 (未勾選時會保留)
      Check4.Value = 1 '清除案件聯絡人資料
   End If
   '2022/7/26 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm110104_1 = Nothing
End Sub

Private Sub optChoose_Click(Index As Integer)
   txtCaseField(0).Enabled = False
   Frame1.Enabled = False
   txtSystem.Enabled = False
   textTM02.Enabled = False
   textTM02_2.Enabled = False
   textTM03.Enabled = False
   textTM04.Enabled = False

   Select Case Index
      Case 0
         txtCaseField(0).Enabled = True
         Frame1.Enabled = True
         'txtCaseField(0).SetFocus
      Case 1
         txtSystem.Enabled = True
         textTM02.Enabled = True
         textTM02_2.Enabled = True
         textTM03.Enabled = True
         textTM04.Enabled = True
         'txtSystem.SetFocus
   End Select
End Sub

Private Sub Text2_GotFocus(Index As Integer)
   TextInverse Text2(Index)
   Select Case Index
      Case 0, 2, 3, 5, 6
         OpenIme
      Case Else
         CloseIme
   End Select
End Sub

Private Sub Text2_Validate(Index As Integer, Cancel As Boolean)
   'Added by Lydia 2017/06/14 設欄位長度
    Dim iLen As Integer
    Select Case Index
    Case 0, 3 '專利-聯絡人中文
         iLen = 30
    Case 1, 4 '聯絡人英文
         iLen = 35
    Case 2, 5, 6 '聯絡人日文
         iLen = 60
    Case Else
         iLen = Text2(Index).MaxLength
    End Select
    'end 2017/06/14
    
   '檢查中文欄位長度是否過長
   'Modified by Lydia 2017/06/14
   'If CheckLengthIsOK(Text2(Index).Text, Text2(Index).MaxLength) Then
   If CheckLengthIsOK(Text2(Index).Text, iLen) Then
      Cancel = False
   Else
      Cancel = True
   End If
   If Cancel = True Then TextInverse Text2(Index)
End Sub

Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCaseField_Change(Index As Integer)
   Select Case Index
      Case 1
         lblAgent = ""
      Case 2
         lblCustomer = ""
      Case 3
         NewAgent = ""
      Case 4
         If Trim(txtCaseField(4)) = "" Then
            lblSName = ""
         Else
            lblSName = GetPrjSalesNM(Trim(txtCaseField(4)))
         End If
   End Select
End Sub

Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)

   If CheckKeyIn(Index) = -1 Then
      Cancel = True
   End If
   If Cancel Then Call txtCaseField_GotFocus(Index)
End Sub

Private Sub txtSystem_Change()
   Select Case txtSystem
      Case "TF":
         textTM02_2.Visible = True
         textTM02_2.Locked = False
         textTM02_2.TabStop = True
         textTM02.MaxLength = 5
      Case Else
         textTM02_2.Visible = False
         textTM02_2.Locked = True
         textTM02_2.TabStop = False
         textTM02.MaxLength = 6
   End Select
End Sub

Private Sub txtSystem_GotFocus()
   txtSystem.SelStart = 0
   txtSystem.SelLength = Len(txtSystem.Text)
   CloseIme
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSystem_Validate(Cancel As Boolean)
   If txtSystem.Text <> "" Then
      If ChkSysID(txtSystem) = False Then
         Cancel = True
         txtSystem_GotFocus
      End If
   End If
End Sub

Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
   CloseIme
End Sub

Private Sub textTM02_2_GotFocus()
   InverseTextBox textTM02_2
   CloseIme
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
   CloseIme
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
   CloseIme
End Sub

Private Sub txtCaseField_GotFocus(Index As Integer)
   txtCaseField(Index).SelStart = 0
   txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
   Select Case Index
      Case 0, 1, 2, 3, 4
         CloseIme
      Case Else
         OpenIme
   End Select
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strNo As String
   
   CheckKeyIn = -1

   Select Case intIndex
      Case 0 '系統類別
         If txtCaseField(intIndex) <> "" Then
            If ChkSysID(txtCaseField(intIndex)) = True Then
               CheckKeyIn = 1
            End If
         Else
            MsgBox "系統別不可空白!!!", vbExclamation + vbOKOnly
            Exit Function
         End If
      Case 1 '代理人代號
         If txtCaseField(intIndex) <> "" Then
            strNo = txtCaseField(intIndex)
            'Modify By Sindy 2015/8/27 +IIf(optChoose(0).Value = True, txtCaseField(0), txtSystem)
            If GetAgentAndState(strNo, strTemp, , False, True, IIf(optChoose(0).Value = True, txtCaseField(0), txtSystem)) Then
               txtCaseField(intIndex) = ChangeCustomerL(strNo)
               lblAgent.Caption = strTemp
               CheckKeyIn = 1
            End If
         Else
            CheckKeyIn = 1
         End If
      Case 2 '申請人代號
         If txtCaseField(intIndex) <> "" Then
            '檢查該申請人或代理人狀態，若為不再使用則停在原地
            strNo = txtCaseField(intIndex)
            'Modify By Sindy 2015/8/27 +IIf(optChoose(0).Value = True, txtCaseField(0), txtSystem)
            If GetCustomerAndState(strNo, strTemp, , , True, IIf(optChoose(0).Value = True, txtCaseField(0), txtSystem)) Then
               txtCaseField(intIndex) = ChangeCustomerL(strNo)
               lblCustomer.Caption = strTemp
               CheckKeyIn = 1
            End If
         Else
            CheckKeyIn = 1
         End If
      Case 3 '新FC代理人代號
         If txtCaseField(intIndex) <> "" Then
            '檢查該申請人或代理人狀態，若為不再使用則停在原地
            strNo = txtCaseField(intIndex)
            'Modify By Sindy 2015/8/27 +IIf(optChoose(0).Value = True, txtCaseField(0), txtSystem)
            If GetAgentAndState(strNo, strTemp, , , True, IIf(optChoose(0).Value = True, txtCaseField(0), txtSystem)) Then
               txtCaseField(intIndex) = ChangeCustomerL(strNo)
               NewAgent.Caption = strTemp
               CheckKeyIn = 1
            End If
            If CheckKeyIn = 1 Then
               '若輸入9碼且最後一碼不為"0"
               If Len(txtCaseField(intIndex)) = 9 And Right(txtCaseField(intIndex), 1) <> "0" Then
                   MsgBox "此代理人已變更名稱，請使用新名稱之編號收文!!!", vbExclamation + vbOKOnly
                   CheckKeyIn = -1
                   Exit Function
               End If
            End If
         Else
            If optChoose(0).Value = True Then
               MsgBox "請輸入新FC代理人!!!", vbExclamation + vbOKOnly
               Exit Function
            End If
            CheckKeyIn = 1
         End If
      Case 4 '智權人員
         If Len(txtCaseField(intIndex).Text) <= 0 Then
            MsgBox "請輸入智權人員!!!", vbExclamation + vbOKOnly
            Exit Function
         End If
         If Not (Left(txtCaseField(intIndex).Text, 1) >= "6" And Left(txtCaseField(intIndex).Text, 1) < "F" And Mid(txtCaseField(intIndex).Text, 4, 1) <> "9") Then
            MsgBox "智權人員輸入錯誤!!!", vbExclamation + vbOKOnly
            Exit Function
         End If
         If ClsPDGetStaff(txtCaseField(intIndex).Text, strTemp) Then
            lblSName.Caption = strTemp
            If ChkStaffST04(txtCaseField(intIndex)) = True Then
               Exit Function
            Else
               CheckKeyIn = 1
            End If
         End If
      Case 5 '本所案號
         strTM01 = txtSystem
         strTM02 = textTM02
         If txtSystem = "TF" Then: strTM02 = strTM02 & textTM02_2
         If textTM03 = "" Then textTM03 = "0"
         If textTM04 = "" Then textTM04 = "00"
         strTM03 = textTM03
         strTM04 = textTM04
         If ClsPDCheckCaseCodeIsExist(strTM01, strTM02, strTM03, strTM04) Then
            CheckKeyIn = 1
         End If
   End Select
End Function

Public Sub Cleartxt()
   txtCaseField(0) = strSystemKind
   txtCaseField(1) = ""
   txtCaseField(2) = ""
   txtCaseField(3) = ""
   txtCaseField(4) = ""
   Check1.Value = 1
   Check2.Value = 0
   Check3.Value = 0
   Check4.Value = 0
   Text2(0) = ""
   Text2(1) = ""
   Text2(2) = ""
   Text2(3) = ""
   Text2(4) = ""
   Text2(5) = ""
   Text2(6) = ""
   txtSystem = ""
   textTM02 = ""
   textTM02_2 = ""
   textTM03 = ""
   textTM04 = ""
End Sub

Private Function ChkSysID(strSys As String) As Boolean
Dim i, j, s As Integer

   ChkSysID = True
   strTemp1 = Split(UCase(strSystemKind), ",")
   strTemp2 = Split(UCase(strSys), ",")
   For i = 0 To UBound(strTemp2)
      s = 0
      For j = 0 To UBound(strTemp1)
         If strTemp1(j) = strTemp2(i) Then
            s = 1
            Exit For
         End If
      Next j
      If s = 0 Then
         ChkSysID = False
         s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
         Exit Function
      End If
   Next i
End Function

'Add by Amy 2017/06/27 下一程序是否有未發文備註
Private Function bolNP15NotNull() As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    bolNP15NotNull = False
    strQ = "select * From nextprogress " & _
            "Where np02='" & txtSystem & "' and np03='" & textTM02 & "' and np04='" & textTM03 & "' and np05='" & textTM04 & "'" & _
            " And np06 is null And np15 is not null" & _
            " Order by np01 asc"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        bolNP15NotNull = True
    End If
End Function
