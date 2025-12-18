VERSION 5.00
Begin VB.Form frm060321 
   BorderStyle     =   1  '單線固定
   Caption         =   "TNT列印"
   ClientHeight    =   6096
   ClientLeft      =   156
   ClientTop       =   1620
   ClientWidth     =   4524
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6096
   ScaleWidth      =   4524
   Begin VB.Frame Frame3 
      Caption         =   "申請人/代理人/潛在客戶"
      Height          =   2055
      Left            =   240
      TabIndex        =   28
      Top             =   1680
      Width           =   4095
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   1560
         TabIndex        =   14
         Top             =   1635
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1560
         TabIndex        =   13
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtNO 
         Height          =   264
         Index           =   0
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtNO 
         Height          =   264
         Index           =   1
         Left            =   2130
         MaxLength       =   6
         TabIndex        =   9
         Top             =   960
         Width           =   720
      End
      Begin VB.TextBox txtNO 
         Height          =   264
         Index           =   2
         Left            =   2925
         MaxLength       =   1
         TabIndex        =   10
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtNO 
         Height          =   264
         Index           =   3
         Left            =   3270
         MaxLength       =   2
         TabIndex        =   11
         Top             =   960
         Width           =   360
      End
      Begin VB.OptionButton OptKind 
         Caption         =   "非案件說明："
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton OptKind 
         Caption         =   "本所案號："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtCNo 
         Height          =   270
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   6
         Top             =   275
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "聯絡人："
         Height          =   180
         Index           =   3
         Left            =   320
         TabIndex        =   32
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Line Line5 
         X1              =   2055
         X2              =   2115
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line4 
         X1              =   2985
         X2              =   3405
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label4 
         Caption         =   "1111"
         Height          =   180
         Index           =   2
         Left            =   1320
         TabIndex        =   31
         Top             =   615
         Width           =   2595
      End
      Begin VB.Label Label4 
         Caption         =   "名　　稱："
         Height          =   180
         Index           =   1
         Left            =   320
         TabIndex        =   30
         Top             =   620
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "編　　號："
         Height          =   180
         Index           =   0
         Left            =   320
         TabIndex        =   29
         Top             =   320
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "本所案號"
      Height          =   975
      Left            =   240
      TabIndex        =   26
      Top             =   600
      Width           =   4095
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   0
         Left            =   1245
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   1
         Left            =   1815
         MaxLength       =   6
         TabIndex        =   1
         Top             =   240
         Width           =   720
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   2
         Left            =   2610
         MaxLength       =   1
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   3
         Left            =   2955
         MaxLength       =   2
         TabIndex        =   3
         Top             =   240
         Width           =   360
      End
      Begin VB.OptionButton Option1 
         Caption         =   "FC代理人"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   4
         Top             =   600
         Width           =   1365
      End
      Begin VB.OptionButton Option1 
         Caption         =   "CF代理人"
         Height          =   255
         Index           =   1
         Left            =   2085
         TabIndex        =   5
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號："
         Height          =   180
         Left            =   360
         TabIndex        =   27
         Top             =   270
         Width           =   915
      End
      Begin VB.Line Line1 
         X1              =   1740
         X2              =   1800
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line2 
         X1              =   2490
         X2              =   2610
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line3 
         X1              =   2670
         X2              =   3090
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   360
      TabIndex        =   19
      Top             =   3945
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   16
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   180
         Left            =   105
         TabIndex        =   21
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   18
      Top             =   4650
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   20
      Top             =   4950
      Width           =   705
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3390
      TabIndex        =   24
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2595
      TabIndex        =   22
      Top             =   60
      Width           =   756
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "畫面上印表機的X及Y偏移值。"
      ForeColor       =   &H00000080&
      Height          =   180
      Index           =   3
      Left            =   810
      TabIndex        =   25
      Top             =   5790
      Width           =   2400
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "注意：新人第一次使用此作業功能，須設定"
      ForeColor       =   &H00000080&
      Height          =   180
      Index           =   2
      Left            =   270
      TabIndex        =   23
      Top             =   5550
      Width           =   3420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "橫軸偏移值(X)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   17
      Top             =   4710
      Width           =   3240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "縱軸偏移值(Y)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   15
      Top             =   5010
      Width           =   3240
   End
End
Attribute VB_Name = "frm060321"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, StrTmpNick As String, StrTmpNick1 As String, StrTmpNick2 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 9) As String, SeekPrint As Integer, SeekPrintL As Integer
Dim PLeft(0 To 5) As Integer, strTemp1 As Variant, strTemp2 As Variant, strNum(0 To 3) As String, poliu As Integer
'Add By Cheng 2002/02/27
Dim m_dbl_LeftMargin  As Double '橫軸偏移值
Dim m_dbl_TopMargin  As Double '縱軸偏移值
'Add By Cheng 2002/12/24
Dim m_CP09 As String '總收文號


'Add By Cheng 2002/12/24
'取得總收文號
Property Let GetCP09(strCP09 As String)
    m_CP09 = strCP09
End Property

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0 '確定
    
    
  If Option1(0).Value = True Or Option1(1).Value = True Then
     If Len(txt1(0)) = 0 Then
        s = MsgBox("第一欄位不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        Exit Sub
     Else
         strNum(0) = txt1(0)
         If Len(txt1(1)) = 0 Then
             s = MsgBox("第二欄位不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             Exit Sub
         Else
             strNum(1) = txt1(1)
             If Len(txt1(2)) = 0 Then
                 strNum(2) = "0"
             Else
                 strNum(2) = txt1(2)
             End If
             If Len(txt1(3)) = 0 Then
                 strNum(3) = "00"
             Else
                 strNum(3) = txt1(3)
             End If
         End If
     End If
  Else 'Add by Lydia 2014/11/20 增加”申請人/代理人/潛在客戶”選項。
     If OptKind(0).Value = True Or OptKind(1).Value = True Then
        If LTrim(RTrim(txtCNo)) = "" Then
           MsgBox "申請人/代理人/潛在客戶只能為 X、Y 或 R !!", , "USER 輸入錯誤"
           txtCNo_GotFocus
           Exit Sub
        End If
        If OptKind(0).Value = True Then
            If txtNo(2) = "" Then txtNo(2) = "0"
            If txtNo(3) = "" Then txtNo(3) = "00"
            If ClsPDCheckCaseCodeIsExist(txtNo(0), txtNo(1), txtNo(2), txtNo(3)) = False Then
              txtNo(1).SetFocus
              Exit Sub
            End If
        End If
        
        If OptKind(1).Value = True Then
           If Len(Text2) = 0 Then
              'Modified by Lydia 2015/12/16
              'MsgBox "非案件說明不可空白!!", , "USER 輸入錯誤"
              MsgBox "非案件說明會列印在案號的位置,方便日後追蹤,所以不可空白!!", vbInformation
              Text2.SetFocus
              Exit Sub
           End If
        End If
     Else
        MsgBox "請選擇條件範圍 !!", , "USER 輸入錯誤"
        If Len(txt1(0)) > 0 Then
           txt1(0).SetFocus
        Else
           txtCNo.SetFocus
        End If
     End If
  End If
  
      PUB_RestorePrinter Combo1 'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
      Printer.Orientation = 1
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      
      'Modify by Morgan 2004/11/26 改用新規則
      'Process
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/22 清除查詢印表記錄檔欄位
      ProcessNew
   
      'Modify by Morgan 2009/1/20 改成列印完就還原(原來Unload才做而在那之前有其他畫面要列印時會抓錯印表機)
      Set Printer = Printers(SeekPrint)
      Printer.Orientation = SeekPrintL
    
      '初始化收文號
      m_CP09 = ""
      
      bolToEndByNick = True
      Me.Enabled = True
      Screen.MousePointer = vbDefault
     
Case 1 '結束
      '若有變動印表機或偏移值, 則更新列印設定
      If Me.Combo1.Text <> Me.Combo1.Tag Or Me.Text1(0).Text <> Me.Text1(0).Tag Or Me.Text1(1).Text <> Me.Text1(1).Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, Me.Text1(0).Text, Me.Text1(1).Text, Me.Combo1.Text
      End If
      bolToEndByNick = True
      Unload Me
Case Else
End Select
End Sub
'Add by Morgan 2004/11/26 搜尋最近發文的資料
Private Sub SetCP09()

On Error GoTo ErrHnd

   'Modify by Morgan 2010/4/13 排除CFT的B類申請英文證明304
   'Modify By Sindy 2012/4/5 +and CP44 is not null
   strSql = "Select * From CaseProgress Where " & ChgCaseprogress(strNum(0) & strNum(1) & strNum(2) & strNum(3)) & " And CP09 < 'C' AND CP27>0 and not (CP01='CFT' and cp09>'B' and cp10='304') and CP44 is not null Order By CP27 Desc,CP09 Desc "
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         m_CP09 = "" & .Fields("CP09").Value
      Else
         m_CP09 = ""
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Add by Morgan 2004/11/26 改規則
'  FC:聯絡人=>1.基本檔-->2.代理人檔(PA75,TM44)-->3.客戶檔(PA26,TM23)
'     名稱&地址=>1.代理人檔(PA75,TM44)-->2.客戶檔(PA26,TM23)
'  CF:聯絡人=>代理人檔(CP44)
'     名稱&地址=>代理人檔(CP44)
Private Sub ProcessNew()
   
'Add by Lydia 2014/11/20 增加”申請人/代理人/潛在客戶”選項。
If Option1(0).Value = True Or Option1(1).Value = True Then
   If Option1(1).Value = True Then
      pub_QL05 = pub_QL05 & ";" & Option1(1).Caption 'Add By Sindy 2010/10/22
      '若是直接下TNT列印, 非發文時列印TNT
      If m_CP09 = "" Then SetCP09
      If m_CP09 = "" Then
         MsgBox "找不到AB類的發文資料！", vbExclamation
         Exit Sub
      End If
      'Modify by Morgan 2011/5/26 +CU28->CU28||rtrim(' '||cu102),FA22->FA22||rtrim(' '||FA70)(TNT行數不夠英文地址5,6合併)
      strSql = "SELECT NVL(FA08,FA53) C00" & _
         ", FA05||' '||FA63||' '||FA64||' '||FA65 C01" & _
         ", FA18 C02,FA19 C03,FA20 C04,FA21 C05,FA22||rtrim(' '||FA70) C06" & _
         ", NVL(FA12,FA13) C07" & _
         " FROM CASEPROGRESS,FAGENT" & _
         " WHERE CP09='" & m_CP09 & "'" & _
         " AND FA01(+)=SUBSTR(CP44,1,8) AND FA02(+)=SUBSTR(CP44,9,1)"
   Else
      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption 'Add By Sindy 2010/10/22
      'Modify by Morgan 2007/1/10 改判斷系統種類
      'Select Case Me.txt1(0).Text
      Select Case CheckSys(Me.txt1(0))
         
         'Case "FCP"
         Case "1"
            'Modify by Morgan 2006/9/25 補FA22
            'Modify by Morgan 2011/5/26 +CU28->CU28||rtrim(' '||cu102),FA22->FA22||rtrim(' '||FA70)(TNT行數不夠英文地址5,6合併)
            strSql = "SELECT NVL(NVL(NVL(NVL(NVL(PA52,PA55),FA08),FA53),CU59),CU62) C00" & _
               ", DECODE(PA75,NULL,CU05||' '||CU88||' '||CU89||' '||CU90,FA05||' '||FA63||' '||FA64||' '||FA65) C01" & _
               ", DECODE(PA75,NULL,CU24,FA18) C02, DECODE(PA75,NULL,CU25,FA19) C03" & _
               ", DECODE(PA75,NULL,CU26,FA20) C04, DECODE(PA75,NULL,CU27,FA21) C05, DECODE(PA75,NULL,CU28||rtrim(' '||cu102),FA22||rtrim(' '||FA70)) C06" & _
               ", DECODE(PA75,NULL,NVL(CU16,CU17),NVL(FA12,FA13)) C07" & _
               " FROM PATENT,CUSTOMER,FAGENT" & _
               " WHERE PA01='" & strNum(0) & "' AND PA02='" & strNum(1) & "' AND PA03='" & strNum(2) & "' AND PA04='" & strNum(3) & "' " & _
               " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1)" & _
               " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9,1)"
            
         'Case "FG"
         Case "5", "6", "7", "8"
            'Modify by Morgan 2006/10/18
            '聯絡人加部門 SP30-->SP71||' '||SP30, TNT的收件人只有一行
            'Modify by Morgan 2011/5/26 +CU28->CU28||rtrim(' '||cu102),FA22->FA22||rtrim(' '||FA70)(TNT行數不夠英文地址5,6合併)
            strSql = "SELECT DECODE(SP30,NULL,NVL(NVL(NVL(FA08,FA53),CU59),CU62),SP71||' '||SP30) C00" & _
               ", DECODE(SP26,NULL,CU05||' '||CU88||' '||CU89||' '||CU90,FA05||' '||FA63||' '||FA64||' '||FA65) C01" & _
               ", DECODE(SP26,NULL,CU24,FA18) C02, DECODE(SP26,NULL,CU25,FA19) C03" & _
               ", DECODE(SP26,NULL,CU26,FA20) C04, DECODE(SP26,NULL,CU27,FA21) C05, DECODE(SP26,NULL,CU28||rtrim(' '||cu102),FA22||rtrim(' '||FA70)) C06" & _
               ", DECODE(SP26,NULL,NVL(CU16,CU17),NVL(FA12,FA13)) C07" & _
               " FROM SERVICEPRACTICE,CUSTOMER,FAGENT" & _
               " WHERE SP01='" & strNum(0) & "' AND SP02='" & strNum(1) & "' AND SP03='" & strNum(2) & "' AND SP04='" & strNum(3) & "' " & _
               " AND CU01(+)=SUBSTR(SP08,1,8) AND CU02(+)=SUBSTR(SP08,9,1)" & _
               " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9,1)"
            
         'Case "FCT"
         Case "2"
            'Modify by Morgan 2011/5/26 +CU28->CU28||rtrim(' '||cu102),FA22->FA22||rtrim(' '||FA70)(TNT行數不夠英文地址5,6合併)
            strSql = "SELECT NVL(NVL(NVL(NVL(NVL(TM39,TM42),FA08),FA53),CU59),CU62) C00" & _
               ", DECODE(TM44,NULL,CU05||' '||CU88||' '||CU89||' '||CU90,FA05||' '||FA63||' '||FA64||' '||FA65) C01" & _
               ", DECODE(TM44,NULL,CU24,FA18) C02, DECODE(TM44,NULL,CU25,FA19) C03" & _
               ", DECODE(TM44,NULL,CU26,FA20) C04, DECODE(TM44,NULL,CU27,FA21) C05, DECODE(TM44,NULL,CU28||rtrim(' '||cu102),FA22||rtrim(' '||FA70)) C06" & _
               ", DECODE(TM44,NULL,NVL(CU16,CU17),NVL(FA12,FA13)) C07" & _
               " FROM TRADEMARK,CUSTOMER,FAGENT" & _
               " WHERE TM01='" & strNum(0) & "' AND TM02='" & strNum(1) & "' AND TM03='" & strNum(2) & "' AND TM04='" & strNum(3) & "' " & _
               " AND CU01(+)=SUBSTR(TM23,1,8) AND CU02(+)=SUBSTR(TM23,9,1)" & _
               " AND FA01(+)=SUBSTR(TM44,1,8) AND FA02(+)=SUBSTR(TM44,9,1)"
      End Select
   End If
   pub_QL05 = pub_QL05 & ";" & Label1 & strNum(0) & "-" & strNum(1) & "-" & strNum(2) & "-" & strNum(3) 'Add By Sindy 2010/10/22
   
'Add by Lydia 2014/11/20 增加”申請人/代理人/潛在客戶”選項。
ElseIf OptKind(0).Value = True Or OptKind(1).Value = True Then

   strExc(0) = Left(LTrim(RTrim(txtCNo)) & "000000000", 9)
   strExc(1) = Left(txtCNo, 1)

   Select Case strExc(1)
        Case "X"
            pub_QL05 = pub_QL05 & ";申請人:" & LTrim(RTrim(txtCNo))
            'Modified by Lydia 2015/12/16 若無英文資料,改抓中文->日文
            'strSql = "SELECT NVL(CU59,CU62) C00, CU05||' '||CU88||' '||CU89||' '||CU90 C01, CU24 C02,CU25 C03,CU26 C04,CU27 C05," & _
                     "CU28||rtrim(' '||CU102) C06, NVL(CU16,CU17) C07 FROM CUSTOMER WHERE " & _
                     "CU01=SUBSTR('" & strExc(0) & "',1,8) AND CU02(+)=SUBSTR('" & strExc(0) & "',9,1)"
            strSql = "SELECT DECODE(CU59||CU62,NULL,DECODE(CU58||CU61,NULL,NVL(CU60,CU63),NVL(CU58,CU61)) ,NVL(CU59,CU62)) C00, " & _
                     "DECODE(CU05||CU88||CU89||CU90,NULL,NVL(CU04,CU06),CU05||' '||CU88||' '||CU89||' '||CU90) C01, " & _
                     "NVL(CU24,NVL(SUBSTR(CU23,1,18),SUBSTR(CU29,1,18))) C02, " & _
                     "NVL(CU25,NVL(SUBSTR(CU23,19,36),SUBSTR(CU29,19,36))) C03, " & _
                     "NVL(CU26,NVL(SUBSTR(CU23,37,54),SUBSTR(CU29,37,54))) C04, " & _
                     "NVL(CU27,NVL(SUBSTR(CU23,55,72),SUBSTR(CU29,55,72))) C05, " & _
                     "NVL(CU28,NVL(SUBSTR(CU23,73,80),SUBSTR(CU29,73,80)))||rtrim(' '||CU102) C06, " & _
                     "NVL(CU16,CU17) C07 FROM CUSTOMER WHERE " & _
                     "CU01=SUBSTR('" & strExc(0) & "',1,8) AND CU02(+)=SUBSTR('" & strExc(0) & "',9,1)"
        Case "Y"
            pub_QL05 = pub_QL05 & ";代理人:" & LTrim(RTrim(txtCNo))
            'Modified by Lydia 2015/12/16 若無英文資料,改抓中文->日文
            'strSql = "SELECT NVL(FA08,FA53) C00, FA05||' '||FA63||' '||FA64||' '||FA65 C01, FA18 C02,FA19 C03,FA20 C04,FA21 C05," & _
                     "FA22||rtrim(' '||FA70) C06, NVL(FA12,FA13) C07 FROM FAGENT WHERE " & _
                     "FA01=SUBSTR('" & strExc(0) & "',1,8) AND FA02=SUBSTR('" & strExc(0) & "',9,1)"
            'Added by Lydia 2015/12/28 特定代理人指定抓中文
            'Modified by Lydia 2017/10/16 改成共用變數
            'If InStr("Y53541,Y52268", ChangeCustomerS(txtCNo)) > 0 Then
            'Modified by Lydia 2025/03/13 改用模組取得
            'If InStr(外翻Y編號, ChangeCustomerS(txtCNo)) > 0 Then
            If InStr(Pub_SetF51Order("Y", ""), ChangeCustomerS(txtCNo)) > 0 Then
                strSql = "SELECT DECODE(FA07||FA52,NULL,DECODE(FA08||FA53,NULL,NVL(FA09,FA54),NVL(FA08,FA53)) ,NVL(FA07,FA52)) C00," & _
                         "DECODE(FA04,NULL,DECODE(FA05||FA63||FA64||FA65,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65),FA04) C01," & _
                         "DECODE(FA17,NULL,NVL(FA18,SUBSTR(FA23,1,18)),SUBSTR(FA17,1,18)) C02," & _
                         "DECODE(FA17,NULL,NVL(FA19,SUBSTR(FA23,19,36)),SUBSTR(FA17,19,36)) C03," & _
                         "DECODE(FA17,NULL,NVL(FA20,SUBSTR(FA23,37,54)),SUBSTR(FA17,37,54)) C04," & _
                         "DECODE(FA17,NULL,NVL(FA21,SUBSTR(FA23,55,72)),SUBSTR(FA17,55,72)) C05," & _
                         "DECODE(FA17,NULL,NVL(FA22,SUBSTR(FA23,73,80))||rtrim(' '||FA70),SUBSTR(FA17,73,80)) C06," & _
                         "NVL(FA12,FA13) C07 FROM FAGENT WHERE " & _
                         "FA01=SUBSTR('" & strExc(0) & "',1,8) AND FA02=SUBSTR('" & strExc(0) & "',9,1)"
            Else
            'END 2015/12/28
                strSql = "SELECT DECODE(FA08||FA53,NULL,DECODE(FA07||FA52,NULL,NVL(FA09,FA54),NVL(FA07,FA52)) ,NVL(FA08,FA53)) C00, " & _
                         "DECODE(FA05||FA63||FA64||FA65,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65) C01, " & _
                         "NVL(FA18,NVL(SUBSTR(FA17,1,18),SUBSTR(FA23,1,18))) C02, " & _
                         "NVL(FA19,NVL(SUBSTR(FA17,19,36),SUBSTR(FA23,19,36))) C03, " & _
                         "NVL(FA20,NVL(SUBSTR(FA17,37,54),SUBSTR(FA23,37,54))) C04, " & _
                         "NVL(FA21,NVL(SUBSTR(FA17,55,72),SUBSTR(FA23,55,72))) C05, " & _
                         "NVL(FA22,NVL(SUBSTR(FA17,73,80),SUBSTR(FA23,73,80)))||rtrim(' '||FA70) C06, " & _
                         "NVL(FA12,FA13) C07 FROM FAGENT WHERE " & _
                         "FA01=SUBSTR('" & strExc(0) & "',1,8) AND FA02=SUBSTR('" & strExc(0) & "',9,1)"
            End If
        Case "R"
            pub_QL05 = pub_QL05 & ";潛在客戶:" & LTrim(RTrim(txtCNo))
            'Modified by Lydia 2015/12/16 若無英文資料,改抓中文->日文
            'strSql = "SELECT '' C00, PCU03||' '||PCU04||' '||PCU05||' '||PCU06 C01, PCU20 C02,PCU21 C03,PCU22 C04,PCU23 C05," & _
                     "PCU24||rtrim(' '||PCU25) C06, NVL(PCU13,PCU14) C07 FROM POTCUSTOMER WHERE " & _
                     "PCU01=SUBSTR('" & strExc(0) & "',1,8) AND PCU02(+)=SUBSTR('" & strExc(0) & "',9,1) " & _
                     "union SELECT '' C00, POC03 C01, POC10 C02,'' C03,'' C04,'' C05," & _
                     "'' C06, NVL(POC05,POC06) C07 FROM POTCUSTOMER1 WHERE " & _
                     "POC01=SUBSTR('" & strExc(0) & "',1,8) AND POC02(+)=SUBSTR('" & strExc(0) & "',9,1) "
            strSql = "SELECT '' C00, DECODE(PCU03||PCU04||PCU05||PCU06,NULL,NVL(PCU08,PCU07),PCU03||' '||PCU04||' '||PCU05||' '||PCU06) C01," & _
                     "NVL(PCU20,NVL(SUBSTR(PCU27,1,18),SUBSTR(PCU26,1,18))) C02," & _
                     "NVL(PCU21,NVL(SUBSTR(PCU27,19,36),SUBSTR(PCU26,19,36))) C03," & _
                     "NVL(PCU22,NVL(SUBSTR(PCU27,37,54),SUBSTR(PCU26,37,54))) C04," & _
                     "NVL(PCU23,NVL(SUBSTR(PCU27,55,72),SUBSTR(PCU26,55,72))) C05," & _
                     "NVL(PCU24,NVL(SUBSTR(PCU27,73,80),SUBSTR(PCU26,73,80)))||rtrim(' '||PCU25) C06," & _
                     "NVL(PCU13,PCU14) C07 FROM POTCUSTOMER WHERE " & _
                     "PCU01=SUBSTR('" & strExc(0) & "',1,8) AND PCU02(+)=SUBSTR('" & strExc(0) & "',9,1) " & _
                     "union SELECT '' C00, NVL(POC03,DECODE(POC23||POC24||POC25||POC26,NULL,POC27,POC23||' '||POC24||' '||POC25||' '||POC26)) C01," & _
                     "POC10 C02,'' C03,'' C04,'' C05,'' C06, NVL(POC05,POC06) C07 FROM POTCUSTOMER1 WHERE " & _
                     "POC01=SUBSTR('" & strExc(0) & "',1,8) AND POC02(+)=SUBSTR('" & strExc(0) & "',9,1) "
   End Select
   pub_QL05 = pub_QL05 & ";" & strExc(0)
End If

On Error GoTo ErrHnd

   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount = 0 Then
         InsertQueryLog (0) 'Add By Sindy 2010/10/22
         If Option1(0).Value = True Or Option1(1).Value = True Then
            s = MsgBox("此本所案號搜尋不到!!", vbInformation)
         Else
         'Add by Lydia 2014/11/20 增加”申請人/代理人/潛在客戶”選項。
            s = MsgBox("此申請人/代理人/潛在客戶號碼搜尋不到!!", vbInformation)
         End If
      Else
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/22
         '聯絡人
         strTemp(1) = "" & .Fields("C00")
         '英文名稱
         strTemp(2) = "" & .Fields("C01")
         '英文地址
         strTemp(3) = "" & .Fields("C02")
         strTemp(4) = "" & .Fields("C03")
         strTemp(5) = "" & .Fields("C04")
         strTemp(6) = "" & .Fields("C05")
         strTemp(7) = "" & .Fields("C06")
         '電話
         strTemp(9) = "" & (.Fields("C07"))
         'Add by Lydia 2014/11/20 增加”申請人/代理人/潛在客戶”選項。
         If Option1(0).Value = True Or Option1(1).Value = True Then
            strTemp(8) = strNum(0) & "-" & strNum(1) & "-" & strNum(2) & "-" & strNum(3)
         Else
            strTemp(8) = Left(LTrim(RTrim(txtCNo)) & "000000000", 9)
            If OptKind(0).Value = True Then
               strTemp(8) = Trim(txtNo(0)) & "-" & Trim(txtNo(1)) & (IIf(Len(txtNo(2)) > 0, "-" & txtNo(2), "-0")) & (IIf(Len(txtNo(3)) > 0, "-" & txtNo(3), "-00"))
            Else
               strTemp(8) = Trim(Text2)
            End If
            
            If Len(Text3) > 0 Then strTemp(1) = LTrim(RTrim(Text3))
            
         End If
         PrintData
      End If
   End With
   
ErrHnd:

   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Sub Process()
'Add By Cheng 2003/02/14
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

'Modify By Cheng 2002/02/27
'For外專
'Modify By Cheng 2002/12/24
'If InStr(txt1(0).Text, "P") > 0 Then
If Me.txt1(0).Text = "FCP" Or Me.txt1(0).Text = "FG" Then
   'Modify by Morgan 2011/5/26 +CU28->CU28||rtrim(' '||cu102),FA22->FA22||rtrim(' '||FA70)(TNT行數不夠英文地址5,6合併)
   strSql = "SELECT PA52,PA55,FA08,FA53,CU59,CU62,FA18,FA19,FA20,FA21,FA22||rtrim(' '||FA70) FA22,CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,FA12,FA13,CU16,CU17,FA05||' '||FA63||' '||FA64||' '||FA65,CU05||' '||CU88||' '||CU89||' '||CU90 FROM PATENT,CUSTOMER,FAGENT WHERE " & SQLNewFag("PA26", "CU") & " AND " & SQLNewFag("PA75", "FA") & " AND PA01='" & strNum(0) & "' AND PA02='" & strNum(1) & "' AND PA03='" & strNum(2) & "' AND PA04='" & strNum(3) & "' "
   strSql = strSql + " union all select SP30,'',FA08,FA53,CU59,CU62,FA18,FA19,FA20,FA21,FA22||rtrim(' '||FA70) FA22,CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,FA12,FA13,CU16,CU17,FA05||' '||FA63||' '||FA64||' '||FA65,CU05||' '||CU88||' '||CU89||' '||CU90 FROM SERVICEPRACTICE,CUSTOMER,FAGENT WHERE " & SQLNewFag("SP08", "CU") & " AND " & SQLNewFag("SP26", "FA") & " AND SP01='" & strNum(0) & "' AND SP02='" & strNum(1) & "' AND SP03='" & strNum(2) & "' AND SP04='" & strNum(3) & "' "
'Add By Cheng 2002/12/24
'For CFP(CFP的代理人是在案件進度檔)
'ElseIf Me.txt1(0).Text = "CFP" And Me.txt1(0).Text = "CPS" Then
ElseIf Me.txt1(0).Text = "CFP" Or Me.txt1(0).Text = "CPS" Then
    '若是直接下TNT列印, 非發文時列印TNT
    If m_CP09 = "" Then
        '搜尋最近發文的資料
        StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(strNum(0) & strNum(1) & strNum(2) & strNum(3)) & " And CP09 < 'C' And CP27 IS Not Null And Cp57 Is Null  Order By CP27 Desc,CP09 Desc "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            m_CP09 = "" & rsA("CP09").Value
        Else
            m_CP09 = ""
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
   'Modify by Morgan 2011/5/26 +CU28->CU28||rtrim(' '||cu102),FA22->FA22||rtrim(' '||FA70)(TNT行數不夠英文地址5,6合併)
   strSql = "SELECT PA52,PA55,FA08,FA53,CU59,CU62,FA18,FA19,FA20,FA21,FA22||rtrim(' '||FA70) FA22,CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,FA12,FA13,CU16,CU17,FA05||' '||FA63||' '||FA64||' '||FA65,CU05||' '||CU88||' '||CU89||' '||CU90 FROM CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND " & SQLNewFag("CP44", "FA") & " AND CP09='" & m_CP09 & "' AND PA01='" & strNum(0) & "' AND PA02='" & strNum(1) & "' AND PA03='" & strNum(2) & "' AND PA04='" & strNum(3) & "' "
   strSql = strSql + " union all select SP30,'',FA08,FA53,CU59,CU62,FA18,FA19,FA20,FA21,FA22||rtrim(' '||FA70) FA22,CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,FA12,FA13,CU16,CU17,FA05||' '||FA63||' '||FA64||' '||FA65,CU05||' '||CU88||' '||CU89||' '||CU90 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND " & SQLNewFag("CP44", "FA") & " AND CP09='" & m_CP09 & "' AND SP01='" & strNum(0) & "' AND SP02='" & strNum(1) & "' AND SP03='" & strNum(2) & "' AND SP04='" & strNum(3) & "' "
'For外商
Else
   'Modify by Morgan 2011/5/26 +CU28->CU28||rtrim(' '||cu102),FA22->FA22||rtrim(' '||FA70)(TNT行數不夠英文地址5,6合併)
   strSql = "SELECT TM39,TM42,FA08,FA53,CU59,CU62,FA18,FA19,FA20,FA21,FA22||rtrim(' '||FA70) FA22,CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,FA12,FA13,CU16,CU17,FA05||' '||FA63||' '||FA64||' '||FA65,CU05||' '||CU88||' '||CU89||' '||CU90 FROM TRADEMARK,CUSTOMER,FAGENT WHERE " & SQLNewFag("TM23", "CU") & " AND " & SQLNewFag("TM44", "FA") & " AND TM01='" & strNum(0) & "' AND TM02='" & strNum(1) & "' AND TM03='" & strNum(2) & "' AND TM04='" & strNum(3) & "' "
   strSql = strSql + " union all select SP30,'',FA08,FA53,CU59,CU62,FA18,FA19,FA20,FA21,FA22||rtrim(' '||FA70) FA22,CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,FA12,FA13,CU16,CU17,FA05||' '||FA63||' '||FA64||' '||FA65,CU05||' '||CU88||' '||CU89||' '||CU90 FROM SERVICEPRACTICE,CUSTOMER,FAGENT WHERE " & SQLNewFag("SP08", "CU") & " AND " & SQLNewFag("SP26", "FA") & " AND SP01='" & strNum(0) & "' AND SP02='" & strNum(1) & "' AND SP03='" & strNum(2) & "' AND SP04='" & strNum(3) & "' "
End If
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
         .MoveFirst
         '聯絡人
         If Len(CheckStr(.Fields(0))) <> 0 Then
            strTemp(1) = CheckStr(.Fields(0))
         Else
            If Len(CheckStr(.Fields(1))) <> 0 Then
               strTemp(1) = CheckStr(.Fields(1))
            Else
               If Len(CheckStr(.Fields(2))) <> 0 Then
                  strTemp(1) = CheckStr(.Fields(2))
               Else
                  If Len(CheckStr(.Fields(3))) <> 0 Then
                     strTemp(1) = CheckStr(.Fields(3))
                  Else
                     If Len(CheckStr(.Fields(4))) <> 0 Then
                        strTemp(1) = CheckStr(.Fields(4))
                     Else
                        If Len(CheckStr(.Fields(5))) <> 0 Then
                           strTemp(1) = CheckStr(.Fields(5))
                        Else
                           strTemp(1) = ""
                        End If
                     End If
                  End If
               End If
            End If
         End If
         '英文地址
         If Len(CheckStr(.Fields(6))) <> 0 Then
            strTemp(3) = CheckStr(.Fields(6))
            strTemp(4) = CheckStr(.Fields(7))
            strTemp(5) = CheckStr(.Fields(8))
            strTemp(6) = CheckStr(.Fields(9))
            strTemp(7) = CheckStr(.Fields(10))
         Else
            If Len(CheckStr(.Fields(11))) <> 0 Then
               strTemp(3) = CheckStr(.Fields(11))
               strTemp(4) = CheckStr(.Fields(12))
               strTemp(5) = CheckStr(.Fields(13))
               strTemp(6) = CheckStr(.Fields(14))
               strTemp(7) = CheckStr(.Fields(15))
            Else
               strTemp(3) = ""
               strTemp(4) = ""
               strTemp(5) = ""
               strTemp(6) = ""
               strTemp(7) = ""
            End If
         End If
         '電話
         If Len(CheckStr(.Fields(16))) <> 0 Then
            strTemp(9) = CheckStr(.Fields(16))
         Else
            If Len(CheckStr(.Fields(17))) <> 0 Then
               strTemp(9) = CheckStr(.Fields(17))
            Else
               If Len(CheckStr(.Fields(18))) <> 0 Then
                  strTemp(9) = CheckStr(.Fields(18))
               Else
                  If Len(CheckStr(.Fields(19))) <> 0 Then
                     strTemp(9) = CheckStr(.Fields(19))
                  Else
                     strTemp(9) = ""
                  End If
               End If
            End If
         End If
         '英文名稱
         If Len(CheckStr(.Fields(20))) <> 0 Then
            strTemp(2) = CheckStr(.Fields(20))
         Else
            If Len(CheckStr(.Fields(13))) <> 0 Then
               strTemp(2) = CheckStr(.Fields(21))
            Else
               strTemp(2) = ""
            End If
         End If
    End With
Else
    s = MsgBox("此本所案號搜尋不到!!")
    Exit Sub
End If
CheckOC
strTemp(8) = strNum(0) & "-" & strNum(1) & "-" & strNum(2) & "-" & strNum(3)
PrintData
End Sub

Sub PrintData()
Dim strSql As String

poliu = 45
'聯絡人
'strTemp(1)
'英文名稱
'strTemp(2)
'英文地址
'strTemp(3)
'案號
'STRTEMP(4)
'電話
'STRTEMP(5)
'For i = 1 To 4
'   strTemp(i) = strTemp(i) & "中文中文中文中文中文中文中文中文中文中文中文中文中文中文中文中文中文中文中文中文"
'Next i

'edit by nickc 2006/11/21 XP 不能指定高度
'95
If pub_OS = "1" Then
    Printer.Height = 6 * 1440
    Printer.Width = 12096
'NT 須先結束文件,否則紙張不會用喜好設定
Else
   'Modify by Morgan 2008/4/9
   'Printer.Orientation = 1
   'Printer.EndDoc
   Printer.PaperSize = PUB_GetPaperSize(11)
   'end 2008/4/9
End If

strTemp(1) = StrToStr(strTemp(1), 81)
Printer.Font.Name = "細明體"
Printer.FontBold = True
'Modify By Cheng 2003/02/14
'調整字型大小
'Printer.Font.Size = 8
Printer.Font.Size = 10
'Printer.CurrentX = 50
'Printer.CurrentY = 2800
'Printer.Print strTemp(0)
'Modify By Cheng 2003/01/30
'修改列印起點設定
''Add By Cheng 2002/02/27
'm_dbl_LeftMargin = 0: m_dbl_TopMargin = -50
m_dbl_LeftMargin = CDbl(Me.Text1(0).Text) * 576: m_dbl_TopMargin = CDbl(Me.Text1(1).Text) * 576
iPrint = 3590 + m_dbl_TopMargin
'Add By Cheng 2003/01/30
'設定列印區域
'Modify By Cheng 2003/02/14
'Printer.Height = 8850


'Modify By Cheng 2002/02/27
'For j = 1 To 8
For j = 2 To 7
'   If LenB(strTemp(j)) <= poliu Then
'      If Len(strTemp(j)) <> 0 Then
'         Printer.CurrentX = 90
'         Printer.CurrentY = iPrint
'         Printer.Print strTemp(j)
'         iPrint = iPrint + 250
'      End If
'   Else
'       For i = 0 To Int((LenB(strTemp(j)) + LenB(strTemp(j)) - 1) / poliu)
'         If Len(StrConv(MidB(StrConv(strTemp(j), vbFromUnicode), i * poliu + 1, poliu), vbUnicode)) <> 0 Then
'            Printer.CurrentX = 90
'            Printer.CurrentY = iPrint
'            Printer.Print StrConv(MidB(StrConv(strTemp(j), vbFromUnicode), i * poliu + 1, poliu), vbUnicode)
'            iPrint = iPrint + 250
'         End If
'       Next i
'   End If
   'Add By Sindy 2009/07/23
   Printer.Font.Size = 14
   'Add by Lydia 2014/11/20 增加”申請人/代理人/潛在客戶”選項。
   'If j = 2 Then
   If (Option1(0).Value = True Or Option1(1).Value = True) And j = 2 Then
      Printer.CurrentX = 90 + m_dbl_LeftMargin
      Printer.CurrentY = 1150 + m_dbl_TopMargin '1250 + m_dbl_TopMargin
      'Add By Sindy 2009/09/02
      '判斷最近一筆收文A、B類智權人員代號第1碼為F時,P案則列印FCP,反之若為T案則列印FCT
      strSql = "SELECT * FROM CaseProgress WHERE CP01='" & strNum(0) & "' and CP02='" & strNum(1) & "' and CP03='" & strNum(2) & "' and CP04='" & strNum(3) & "' and CP09 < 'C' and not CP13 is null order by CP05 desc "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Left(Trim(RsTemp("CP13")), 1) = "F" And txt1(0) = "P" Then
            Printer.Print "FCP"
         ElseIf Left(Trim(RsTemp("CP13")), 1) = "F" And txt1(0) = "T" Then
            Printer.Print "FCT"
         Else
            Printer.Print txt1(0)
         End If
      Else
         Printer.Print txt1(0)
      End If
   End If
   Printer.Font.Size = 10
   '2009/07/23 End
   If LenB(StrConv(strTemp(j), vbFromUnicode)) > poliu Then
       StrTmpNick = strTemp(j)
       StrTmpNick2 = StrTmpNick
       Do While Len(Trim(StrTmpNick)) <> 0
          StrTmpNick1 = StrToStr(StrTmpNick, poliu / 2)
          Printer.CurrentX = 90 + m_dbl_LeftMargin
          Printer.CurrentY = iPrint
          Printer.Print StrTmpNick1
            'Modify By Cheng 2003/02/14
            '調整Y軸位置
'          iPrint = iPrint + 250
          iPrint = iPrint + 240
          StrTmpNick = Replace(StrTmpNick, StrTmpNick1, "")
          If StrTmpNick = StrTmpNick2 Then
             StrTmpNick = Replace(StrTmpNick, Left(StrTmpNick1, Len(StrTmpNick1) - 1), "")
             StrTmpNick2 = StrTmpNick
          Else
             StrTmpNick2 = StrTmpNick
          End If
       Loop
   Else
      Printer.CurrentX = 90 + m_dbl_LeftMargin
      Printer.CurrentY = iPrint
      Printer.Print strTemp(j)
      'Modify By Cheng 2003/02/14
      '調整Y軸位置
'     iPrint = iPrint + 250
      iPrint = iPrint + 240
   End If
Next j
'Add By Cheng 2002/02/27
'列印ContactName
Printer.CurrentX = 90 + 100 + m_dbl_LeftMargin
Printer.CurrentY = 5030 + m_dbl_TopMargin
Printer.Print strTemp(1)
'列印TelNo
Printer.CurrentX = 3500 + m_dbl_LeftMargin
Printer.CurrentY = 5030 + m_dbl_TopMargin
Printer.Print strTemp(9)
'列印案號
Printer.CurrentX = 5600 + m_dbl_LeftMargin
'Modify By Cheng 2003/02/14
'調整Y軸位置
'Printer.CurrentY = 5850 + m_dbl_TopMargin
Printer.CurrentY = 5900 + m_dbl_TopMargin
Printer.Print strTemp(8)

Printer.EndDoc
ShowPrintOk
End Sub

Sub Process1()
'Modify by Morgan 2011/5/26 +FA70
If Len(CheckStr(adoRecordset.Fields(0))) = 9 Then
    strSql = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,NVL(FA32||FA33||FA34||FA35||FA36,FA18||FA19||FA20||FA21||FA22||FA70),FA12 FROM FAGENT WHERE FA01='" & Mid(CheckStr(adoRecordset.Fields(0)), 1, 8) & "' AND FA02='" & Mid(CheckStr(adoRecordset.Fields(0)), 9, 1) & "' "
Else
    strSql = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,NVL(FA32||FA33||FA34||FA35||FA36,FA18||FA19||FA20||FA21||FA22||FA70),FA12 FROM FAGENT WHERE FA01='" & Mid(CheckStr(adoRecordset.Fields(0)), 1, 8) & "' AND FA02='0' "
End If
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    strTemp(0) = CheckStr(adoRecordset1.Fields(0))
    strTemp(1) = CheckStr(adoRecordset1.Fields(1))
    strTemp(3) = CheckStr(adoRecordset1.Fields(2))
Else
    strTemp(0) = ""
    strTemp(1) = ""
    strTemp(3) = ""
End If
CheckOC2
End Sub

Sub Process2()
If Len(CheckStr(adoRecordset.Fields(1))) = 9 Then
    strSql = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,NVL(CU65||CU66||CU67||CU68||CU69,CU24||CU25||CU26||CU27||CU28),CU16 FROM CUSTOMER WHERE CU01='" & Mid(CheckStr(adoRecordset.Fields(1)), 1, 8) & "' AND CU02='" & Mid(CheckStr(adoRecordset.Fields(1)), 9, 1) & "' "
Else
    strSql = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,NVL(CU65||CU66||CU67||CU68||CU69,CU24||CU25||CU26||CU27||CU28),CU16 FROM CUSTOMER WHERE CU01='" & Mid(CheckStr(adoRecordset.Fields(1)), 1, 8) & "' AND CU02='0' "
End If
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    strTemp(0) = CheckStr(adoRecordset1.Fields(0))
    strTemp(1) = CheckStr(adoRecordset1.Fields(1))
    strTemp(3) = CheckStr(adoRecordset1.Fields(2))
Else
    strTemp(0) = ""
    strTemp(1) = ""
    strTemp(3) = ""
End If
CheckOC2
End Sub

Private Sub Form_Load()

MoveFormToCenter Me

SeekPrintL = Printer.Orientation
PUB_SetPrinter Me.Name, Combo1, , False, SeekPrint    'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除

'Add by Lydia 2014/11/20 增加”申請人/代理人/潛在客戶”選項。
Label4(2).Caption = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm060321 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    'Add By Cheng 2003/01/30
    '反白設定
    TextInverse Me.Text1(Index)
End Sub

Private Sub txt1_GotFocus(Index As Integer)
'Add by Lydia 2014/11/20
If txtCNo.Text <> "" Or (OptKind(0).Value = True Or OptKind(1).Value = True) Then
   ClearInputData 2
End If

txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
'Add by Lydia 2014/11/20
CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
   Case 0
      'Modify By Cheng 2002/03/01
      If App.EXEName = "Patent" Then
        If txt1(Index).Enabled = True Then
        Select Case txt1(0)
           'Modify By Cheng 2003/02/05
           '加系統類別CFP
   '     Case "FCP", "FG", ""
        Case "FCP", "FG", "CFP", ""
        Case Else
             s = MsgBox("本所案號只能 FCP 或 FG 或 CFP !!!", , "USER 輸入錯誤")
             txt1(0).SetFocus
             txt1(0).SelStart = 0
             txt1(0).SelLength = Len(txt1(0))
             Exit Sub
        End Select
        End If
      ElseIf App.EXEName = "Trademark" Then
        If txt1(Index).Enabled = True Then
        Select Case txt1(0)
        Case "CFT", "FCT", ""
        Case Else
             s = MsgBox("本所案號只能 CFT 或 FCT !!", , "USER 輸入錯誤")
             txt1(0).SetFocus
             txt1(0).SelStart = 0
             txt1(0).SelLength = Len(txt1(0))
             Exit Sub
        End Select
        End If
      End If
   Case Else
   End Select

   'Add by Morgan 2007/1/10 預設代理人種類
   Select Case Me.txt1(0).Text
      'FC
      Case "FCP", "FG", "FCT"
         Option1(0).Value = True
      'CF
      Case Else
         Option1(1).Value = True
   End Select
   'end 2007/1/10
End Sub
'Add by Lydia 2014/11/20
Private Sub txtCNo_GotFocus()
If txt1(0).Text <> "" Or (Option1(0).Value = True Or Option1(1).Value = True) Then
   ClearInputData 1
End If

txtCNo.SelStart = 0
txtCNo.SelLength = Len(txtCNo)
CloseIme
End Sub

Private Sub OptKind_Click(Index As Integer)
If Index = 0 Then
   Text2.Text = ""
   Text3.Text = ""
   txtNo(0).SetFocus
Else
   For intI = 0 To 3
       txtNo(intI).Text = ""
   Next intI
   Text2.SetFocus
End If

End Sub
Private Sub txtCNo_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCNo_LostFocus()
Dim clname As String
If LTrim(RTrim(txtCNo)) <> "" Then
   clname = "求值"
   If PUB_GetCustData(Left(LTrim(RTrim(txtCNo)) & "000000000", 9), , clname) = True Then
      Label4(2).Caption = clname
   End If
End If
End Sub

Private Sub txtCNo_Validate(Cancel As Boolean)
If Trim(txtCNo) <> "" Then
   strExc(0) = Left(LTrim(RTrim(txtCNo)), 1)
   If Not (strExc(0) = "X" Or strExc(0) = "Y" Or strExc(0) = "R") Then
      MsgBox "申請人/代理人/潛在客戶只能為 X、Y 或 R !!", , "USER 輸入錯誤"
      txtCNo_GotFocus
   End If
End If
End Sub
Private Sub ClearInputData(cInX As Integer)
Dim rr As Integer

Select Case cInX
       Case 1 '清空本所案號條件範圍
            For rr = 0 To 3
                txt1(rr) = ""
                If rr < 2 Then Option1(rr).Value = False
            Next rr
       
       Case 2 '清空申請人/代理人/潛在客戶範圍
            For rr = 0 To 3
                txtNo(rr) = ""
                If rr < 2 Then OptKind(rr).Value = False
            Next rr
            txtCNo.Text = ""
            Text2.Text = ""
            Text3.Text = ""
End Select

End Sub
Private Sub txtNO_LostFocus(Index As Integer)
   Select Case Index
   Case 0

      If App.EXEName = "Patent" Then
        If txtNo(Index).Enabled = True Then
        Select Case txtNo(0)
        Case "FCP", "FG", "CFP", ""
        Case Else
             s = MsgBox("本所案號只能 FCP 或 FG 或 CFP !!!", , "USER 輸入錯誤")
             txtNo(0).SetFocus
             txtNo(0).SelStart = 0
             txtNo(0).SelLength = Len(txtNo(0))
             Exit Sub
        End Select
        End If
      ElseIf App.EXEName = "Trademark" Then
        If txtNo(Index).Enabled = True Then
        Select Case txtNo(0)
        Case "CFT", "FCT", ""
        Case Else
             s = MsgBox("本所案號只能 CFT 或 FCT !!", , "USER 輸入錯誤")
             txtNo(0).SetFocus
             txtNo(0).SelStart = 0
             txtNo(0).SelLength = Len(txtNo(0))
             Exit Sub
        End Select
        End If
      End If
   Case Else
   End Select
End Sub
Private Sub txtNo_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtNo_GotFocus(Index As Integer)
txtNo(Index).SelStart = 0
txtNo(Index).SelLength = Len(txtNo(Index))
CloseIme
End Sub

'end 'Add by Lydia 2014/11/20


