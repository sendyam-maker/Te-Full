VERSION 5.00
Begin VB.Form frm140412 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外部新客戶/代理人查詢"
   ClientHeight    =   4270
   ClientLeft      =   1290
   ClientTop       =   2960
   ClientWidth     =   4860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4270
   ScaleWidth      =   4860
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   3540
      ScaleHeight     =   460
      ScaleWidth      =   650
      TabIndex        =   17
      Top             =   1140
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1230
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2580
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1230
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1500
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1155
      MaxLength       =   7
      TabIndex        =   0
      Top             =   765
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2280
      MaxLength       =   7
      TabIndex        =   1
      Top             =   765
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1155
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1140
      Width           =   630
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1140
      Width           =   630
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1155
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1845
      Width           =   630
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1845
      Width           =   630
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1230
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2220
      Width           =   240
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2676
      TabIndex        =   9
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3468
      TabIndex        =   10
      Top             =   60
      Width           =   800
   End
   Begin VB.Label Label1 
      Caption         =   "注意：在讀取”新客戶”資料時，因為要過濾以前有收文之案件及客戶資料，大約需要5分鐘的時間，請耐心等候！"
      ForeColor       =   &H00FF0000&
      Height          =   585
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   3630
      Width           =   4440
   End
   Begin VB.Label Label1 
      Caption         =   "PS：新客戶/代理人係指第一次接洽記錄單收文符合查詢條件者"
      ForeColor       =   &H000000C0&
      Height          =   435
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   3150
      Width           =   4440
   End
   Begin VB.Label Label1 
      Caption         =   "輸出方式： 　    ( 1.螢幕 2.印表機 )"
      Height          =   180
      Index           =   9
      Left            =   240
      TabIndex        =   16
      Top             =   2640
      Width           =   3090
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "對　象：         ( 1.代理人 2.客戶 )"
      Height          =   180
      Index           =   7
      Left            =   420
      TabIndex        =   15
      Top             =   1560
      Width           =   2550
   End
   Begin VB.Line Line3 
      X1              =   1695
      X2              =   2040
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Line Line2 
      X1              =   1650
      X2              =   2070
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line1 
      X1              =   2070
      X2              =   2340
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Label1 
      Caption         =   "收文日："
      Height          =   180
      Index           =   2
      Left            =   420
      TabIndex        =   14
      Top             =   810
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "國　籍："
      Height          =   180
      Index           =   3
      Left            =   420
      TabIndex        =   13
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      Height          =   180
      Index           =   5
      Left            =   420
      TabIndex        =   12
      Top             =   1890
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "查詢別：   　  ( 1.明細 2.統計 )"
      Height          =   180
      Index           =   8
      Left            =   420
      TabIndex        =   11
      Top             =   2280
      Width           =   3405
   End
End
Attribute VB_Name = "frm140412"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo by Sindy 2010/9/2 日期欄已修改
Option Explicit

Dim blnClkSure As Boolean, Page As Integer
Dim m_bPrinter As Boolean, m_Device, m_iPages As Integer
Dim m_rs As New ADODB.Recordset
Dim s As Integer
Dim m_i As Integer
Dim PLeft(1 To 7) As Integer
Dim strTemp(1 To 7) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If Len(txt1(0)) = 0 Or Len(txt1(1)) = 0 Then
             s = MsgBox("收文日不可空白!!", , "USER 輸入錯誤")
             If Len(txt1(0)) = 0 Then txt1(0).SetFocus
             If Len(txt1(1)) = 0 Then txt1(1).SetFocus
             Exit Sub
         End If
         If Len(txt1(4)) = 0 Then
             s = MsgBox("對象不可空白!!", , "USER 輸入錯誤")
             txt1(4).SetFocus
             Exit Sub
         End If
         If Len(txt1(7)) = 0 Then
            s = MsgBox("查詢別不可空白!!", , "USER 輸入錯誤")
            txt1(7).SetFocus
            Exit Sub
         End If
         If Len(txt1(8)) = 0 Then
            s = MsgBox("輸出方式不可空白!!", , "USER 輸入錯誤")
            txt1(8).SetFocus
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(0)) = -1 Then
            Me.txt1(0).SetFocus
            txt1_GotFocus 0
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
            Me.txt1(1).SetFocus
            txt1_GotFocus 1
            Exit Sub
         End If
         If Me.txt1(0).Text <> "" And Me.txt1(1).Text <> "" Then
            If Val(Me.txt1(0).Text) > Val(Me.txt1(1).Text) Then
               MsgBox "收文日範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               blnClkSure = True
               Me.txt1(0).SetFocus
               txt1_GotFocus 0
            End If
         End If
         If Me.txt1(2).Text <> "" And Me.txt1(3).Text <> "" Then
            If Val(Me.txt1(2).Text) > Val(Me.txt1(3).Text) Then
               MsgBox "國籍範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               blnClkSure = True
               Me.txt1(2).SetFocus
               txt1_GotFocus 2
            End If
         End If
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         Process
         Me.Enabled = True
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
      Case Else
   End Select
End Sub

Private Sub Process()
Dim strCon As String, strCon2 As String, longCnt As Long
Dim dblTot1 As Double, dblTot2 As Double, dblTot3 As Double
   
On Error GoTo ErrHnd
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2025/8/14 清除查詢印表記錄檔欄位
   strCon = "": strCon2 = ""
   dblTot1 = 0: dblTot2 = 0: dblTot3 = 0
   If txt1(4) = "1" Then '1.代理人
      pub_QL05 = pub_QL05 & ";對象：1.代理人" 'Add By Sindy 2025/8/14
      '業務區
      If txt1(5) <> "" Then
         strCon2 = strCon2 & " and cp12>='" & txt1(5) & "' "
      End If
      If txt1(6) <> "" Then
         strCon2 = strCon2 & " and cp12<='" & txt1(6) & "' "
      End If
      '國籍
      If txt1(2) <> "" Then
         strCon = strCon & " and fa10>='" & txt1(2) & "' "
      End If
      If txt1(3) <> "" Then
         strCon = strCon & " and fa10<='" & txt1(3) & "' "
      End If
   Else '2.客戶
      pub_QL05 = pub_QL05 & ";對象：2.客戶" 'Add By Sindy 2025/8/14
      '業務區
      If txt1(5) <> "" Then
         strCon = strCon & " and cp12>='" & txt1(5) & "' "
      End If
      If txt1(6) <> "" Then
         strCon = strCon & " and cp12<='" & txt1(6) & "' "
      End If
      '國籍
      If txt1(2) <> "" Then
         strCon = strCon & " and cu10>='" & txt1(2) & "' "
      End If
      If txt1(3) <> "" Then
         strCon = strCon & " and cu10<='" & txt1(3) & "' "
      End If
   End If
   'Add By Sindy 2025/8/14
   If Len(txt1(0)) <> 0 Or Len(txt1(1)) <> 0 Then
      pub_QL05 = pub_QL05 & ";收文日：" & txt1(0) & "-" & txt1(1)
   End If
   If Len(txt1(2)) <> 0 Or Len(txt1(3)) <> 0 Then
      pub_QL05 = pub_QL05 & ";國籍：" & txt1(2) & "-" & txt1(3)
   End If
   If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
      pub_QL05 = pub_QL05 & ";業務區：" & txt1(5) & "-" & txt1(6)
   End If
   pub_QL05 = pub_QL05 & ";查詢別：" & IIf(txt1(7) = "1", "1.明細", "2.統計")
   pub_QL05 = pub_QL05 & ";輸出方式：" & IIf(txt1(8) = "1", "1.螢幕", "2.印表機")
   '2025/8/14 END

   Screen.MousePointer = vbHourglass
   If txt1(4) = "1" Then '1.代理人
      strExc(0) = "select cp1.cp139,nvl(fa05,nvl(fa04,fa06)),na03,substr(a0902,1,6),cp2.cp01,sqldatet(substr(cp1.cp,1,8)) " & _
                         "from caseprogress cp2,fagent,acc090,nation,(SELECT CP139,MIN(CP05||' '||CP09) cp FROM caseprogress where cp139 is not null and cp09||''<'B' " & strCon2 & " GROUP by CP139) cp1, " & _
                         "( " & _
                         "select distinct fagno from ( " & _
                         "select pa75 fagno from caseprogress,patent where cp01 in ('FCP','P','CFP') and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' " & strCon2 & " and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and pa75 is not null " & _
                         "Union " & _
                         "select tm44 fagno from caseprogress,trademark where cp01 in ('FCT','T','CFT','TF') and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' " & strCon2 & " and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and tm44 is not null " & _
                         "Union " & _
                         "select lc22 fagno from caseprogress,lawcase where cp01 in ('FCL','LIN','ACS','CFL','L') and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' " & strCon2 & " and cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 and lc22 is not null " & _
                         "Union " & _
                         "select sp26 fagno from caseprogress,servicepractice where cp01 NOT in ('FCP','P','CFP','FCT','T','CFT','TF','FCL','LIN','ACS','CFL','L','LA') and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' " & strCon2 & " and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 and sp26 is not null) " & _
                         ") case " & _
                         "where fagno=cp1.cp139 and cp1.cp>='" & DBDATE(txt1(0)) & "' and cp2.cp09=substr(cp1.cp,10,9) and substr(cp1.cp139,1,8)=fa01(+) and substr(cp1.cp139,9,1)=fa02(+) and cp2.cp12=a0901(+) and fa10=na01(+) " & strCon & _
                         "order by 6,1 "
                         
   ElseIf txt1(4) = "2" Then '2.客戶
      cnnConnection.BeginTrans
      '刪除資料再統計
      strSql = "delete R140412 where ID='" & strUserNum & "'"
      cnnConnection.Execute strSql
      strSql = "delete R140412C where ID='" & strUserNum & "'"
      cnnConnection.Execute strSql
      '新增國外部當月收文之所有客戶
      '2010/10/28 modify by sonia加當月收文移轉讓與的受讓人,否則當月未發文時會無法計入
      '因後面有剔除查名案件故當月收文之查名也不列入,待新申請案才計入
      '再剔除當收移轉人讓與人之案件,只抓受讓人FCT-30887移轉人X65600不抓,受讓人X65601要抓(但此案另收A類回覆代理人故X65600仍會出現)
      strSql = "insert into R140412 (" & _
                  " select distinct substr(cuno,1,8),'" & strUserNum & "' from (" & _
                        " select pa26 cuno from caseprogress,patent,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 in ('FCP','P','CFP') AND CP55 IS NULL and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and pa26 is not null and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+)" & strCon & _
                  " union select pa27 cuno from caseprogress,patent,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 in ('FCP','P','CFP') AND CP55 IS NULL and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and pa27 is not null and substr(pa27,1,8)=cu01(+) and substr(pa27,9,1)=cu02(+)" & strCon & _
                  " union select pa28 cuno from caseprogress,patent,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 in ('FCP','P','CFP') AND CP55 IS NULL and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and pa28 is not null and substr(pa28,1,8)=cu01(+) and substr(pa28,9,1)=cu02(+)" & strCon & _
                  " union select pa29 cuno from caseprogress,patent,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 in ('FCP','P','CFP') AND CP55 IS NULL and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and pa29 is not null and substr(pa29,1,8)=cu01(+) and substr(pa29,9,1)=cu02(+)" & strCon & _
                  " union select pa30 cuno from caseprogress,patent,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 in ('FCP','P','CFP') AND CP55 IS NULL and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and pa30 is not null and substr(pa30,1,8)=cu01(+) and substr(pa30,9,1)=cu02(+)" & strCon & _
                  " union select tm23 cuno from caseprogress,trademark,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 in ('FCT','T','CFT','TF') and cp10<>'001' AND CP55 IS NULL and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and tm23 is not null and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+)" & strCon & _
                  " union select tm78 cuno from caseprogress,trademark,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 in ('FCT','T','CFT','TF') and cp10<>'001' AND CP55 IS NULL and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and tm78 is not null and substr(tm78,1,8)=cu01(+) and substr(tm78,9,1)=cu02(+)" & strCon & _
                  " union select tm79 cuno from caseprogress,trademark,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 in ('FCT','T','CFT','TF') and cp10<>'001' AND CP55 IS NULL and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and tm79 is not null and substr(tm79,1,8)=cu01(+) and substr(tm79,9,1)=cu02(+)" & strCon & _
                  " union select tm80 cuno from caseprogress,trademark,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 in ('FCT','T','CFT','TF') and cp10<>'001' AND CP55 IS NULL and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and tm80 is not null and substr(tm80,1,8)=cu01(+) and substr(tm80,9,1)=cu02(+)" & strCon & _
                  " union select tm81 cuno from caseprogress,trademark,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 in ('FCT','T','CFT','TF') and cp10<>'001' AND CP55 IS NULL and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and tm81 is not null and substr(tm81,1,8)=cu01(+) and substr(tm81,9,1)=cu02(+)" & strCon
      'Modify By Sindy 2011/2/24 增加LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27
      strSql = strSql & _
                  " union select lc11 cuno from caseprogress,lawcase,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 in ('FCL','LIN','ACS','CFL','L') AND CP55 IS NULL and cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 and lc11 is not null and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+)" & strCon & _
                  " union select lc43 cuno from caseprogress,lawcase,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 in ('FCL','LIN','ACS','CFL','L') AND CP55 IS NULL and cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 and lc43 is not null and substr(lc43,1,8)=cu01(+) and substr(lc43,9,1)=cu02(+)" & strCon & _
                  " union select lc44 cuno from caseprogress,lawcase,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 in ('FCL','LIN','ACS','CFL','L') AND CP55 IS NULL and cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 and lc44 is not null and substr(lc44,1,8)=cu01(+) and substr(lc44,9,1)=cu02(+)" & strCon & _
                  " union select lc45 cuno from caseprogress,lawcase,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 in ('FCL','LIN','ACS','CFL','L') AND CP55 IS NULL and cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 and lc45 is not null and substr(lc45,1,8)=cu01(+) and substr(lc45,9,1)=cu02(+)" & strCon & _
                  " union select lc46 cuno from caseprogress,lawcase,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 in ('FCL','LIN','ACS','CFL','L') AND CP55 IS NULL and cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 and lc46 is not null and substr(lc46,1,8)=cu01(+) and substr(lc46,9,1)=cu02(+)" & strCon & _
                  " union select hc05 cuno from caseprogress,hirecase,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 = 'LA' AND CP55 IS NULL and cp01=hc01 and cp02=hc02 and cp03=hc03 and cp04=hc04 and hc05 is not null and substr(hc05,1,8)=cu01(+) and substr(hc05,9,1)=cu02(+)" & strCon & _
                  " union select hc24 cuno from caseprogress,hirecase,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 = 'LA' AND CP55 IS NULL and cp01=hc01 and cp02=hc02 and cp03=hc03 and cp04=hc04 and hc24 is not null and substr(hc24,1,8)=cu01(+) and substr(hc24,9,1)=cu02(+)" & strCon & _
                  " union select hc25 cuno from caseprogress,hirecase,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 = 'LA' AND CP55 IS NULL and cp01=hc01 and cp02=hc02 and cp03=hc03 and cp04=hc04 and hc25 is not null and substr(hc25,1,8)=cu01(+) and substr(hc25,9,1)=cu02(+)" & strCon & _
                  " union select hc26 cuno from caseprogress,hirecase,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 = 'LA' AND CP55 IS NULL and cp01=hc01 and cp02=hc02 and cp03=hc03 and cp04=hc04 and hc26 is not null and substr(hc26,1,8)=cu01(+) and substr(hc26,9,1)=cu02(+)" & strCon & _
                  " union select hc27 cuno from caseprogress,hirecase,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 = 'LA' AND CP55 IS NULL and cp01=hc01 and cp02=hc02 and cp03=hc03 and cp04=hc04 and hc27 is not null and substr(hc27,1,8)=cu01(+) and substr(hc27,9,1)=cu02(+)" & strCon
      '2011/2/24 End
      strSql = strSql & _
                  " union select sp08 cuno from caseprogress,servicepractice,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 NOT in ('FCP','P','CFP','FCT','T','CFT','TF','FCL','LIN','ACS','CFL','L','LA') and cp10<>'001' AND CP55 IS NULL and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 and sp08 is not null and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+)" & strCon & _
                  " union select sp58 cuno from caseprogress,servicepractice,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 NOT in ('FCP','P','CFP','FCT','T','CFT','TF','FCL','LIN','ACS','CFL','L','LA') and cp10<>'001' AND CP55 IS NULL and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 and sp58 is not null and substr(sp58,1,8)=cu01(+) and substr(sp58,9,1)=cu02(+)" & strCon & _
                  " union select sp59 cuno from caseprogress,servicepractice,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 NOT in ('FCP','P','CFP','FCT','T','CFT','TF','FCL','LIN','ACS','CFL','L','LA') and cp10<>'001' AND CP55 IS NULL and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 and sp59 is not null and substr(sp59,1,8)=cu01(+) and substr(sp59,9,1)=cu02(+)" & strCon & _
                  " union select sp65 cuno from caseprogress,servicepractice,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 NOT in ('FCP','P','CFP','FCT','T','CFT','TF','FCL','LIN','ACS','CFL','L','LA') and cp10<>'001' AND CP55 IS NULL and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 and sp65 is not null and substr(sp65,1,8)=cu01(+) and substr(sp65,9,1)=cu02(+)" & strCon & _
                  " union select sp66 cuno from caseprogress,servicepractice,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp01 NOT in ('FCP','P','CFP','FCT','T','CFT','TF','FCL','LIN','ACS','CFL','L','LA') and cp10<>'001' AND CP55 IS NULL and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 and sp66 is not null and substr(sp66,1,8)=cu01(+) and substr(sp66,9,1)=cu02(+)" & strCon & _
                  " union select cp56 cuno from caseprogress,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp56 is not null and substr(cp56,1,8)=cu01(+) and substr(cp56,9,1)=cu02(+)" & strCon & _
                  " union select cp89 cuno from caseprogress,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp89 is not null and substr(cp89,1,8)=cu01(+) and substr(cp89,9,1)=cu02(+)" & strCon & _
                  " union select cp90 cuno from caseprogress,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp90 is not null and substr(cp90,1,8)=cu01(+) and substr(cp90,9,1)=cu02(+)" & strCon & _
                  " union select cp91 cuno from caseprogress,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp91 is not null and substr(cp91,1,8)=cu01(+) and substr(cp91,9,1)=cu02(+)" & strCon & _
                  " union select cp92 cuno from caseprogress,customer where cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " and cp09<'B' and substr(cp12,1,1)='F' and cp92 is not null and substr(cp92,1,8)=cu01(+) and substr(cp92,9,1)=cu02(+)" & strCon & _
                  " ))"
      cnnConnection.Execute strSql
      
      '檢查有無資料
      strSql = "SELECT count(*) FROM R140412 where ID='" & strUserNum & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.Fields(0) > 0 Then
            '刪除以前曾收移轉讓與客戶
            '2010/10/28應同時刪除以前曾收移轉讓與之受讓人客戶
            strSql = "delete R140412 where ID='" & strUserNum & "' AND R1401 in (" & _
                        " select distinct substr(cuno,1,8) from (" & _
                             "  select CP55 cuno from caseprogress where cp05+0<" & DBDATE(txt1(0)) & " and cp09||''<'B' and substr(cp12,1,1)='F' and cp55 is not null" & _
                        " union select CP93 cuno from caseprogress where cp05+0<" & DBDATE(txt1(0)) & " and cp09||''<'B' and substr(cp12,1,1)='F' and cp93 is not null" & _
                        " union select CP94 cuno from caseprogress where cp05+0<" & DBDATE(txt1(0)) & " and cp09||''<'B' and substr(cp12,1,1)='F' and cp94 is not null" & _
                        " union select CP95 cuno from caseprogress where cp05+0<" & DBDATE(txt1(0)) & " and cp09||''<'B' and substr(cp12,1,1)='F' and cp95 is not null" & _
                        " union select CP96 cuno from caseprogress where cp05+0<" & DBDATE(txt1(0)) & " and cp09||''<'B' and substr(cp12,1,1)='F' and cp96 is not null" & _
                        " union select cp56 cuno from caseprogress where cp05+0<" & DBDATE(txt1(0)) & " and cp09||''<'B' and substr(cp12,1,1)='F' and cp56 is not null" & _
                        " union select cp89 cuno from caseprogress where cp05+0<" & DBDATE(txt1(0)) & " and cp09||''<'B' and substr(cp12,1,1)='F' and cp89 is not null" & _
                        " union select cp90 cuno from caseprogress where cp05+0<" & DBDATE(txt1(0)) & " and cp09||''<'B' and substr(cp12,1,1)='F' and cp90 is not null" & _
                        " union select cp91 cuno from caseprogress where cp05+0<" & DBDATE(txt1(0)) & " and cp09||''<'B' and substr(cp12,1,1)='F' and cp91 is not null" & _
                        " union select cp92 cuno from caseprogress where cp05+0<" & DBDATE(txt1(0)) & " and cp09||''<'B' and substr(cp12,1,1)='F' and cp92 is not null" & _
                        " ))"
            cnnConnection.Execute strSql
            '新增以前有收文之案件及客戶
            'Modify By Sindy 2010/10/8 增加讀取本所案號
            'Modify By Sindy 2011/2/24 增加LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27
            strSql = "insert into R140412C (" & _
                        " select pa26,pa27,pa28,pa29,pa30,'" & strUserNum & "',pa01,pa02,pa03,pa04 from caseprogress," & _
                        " (select substr(pa26,1,8) as pa26,substr(pa27,1,8) as pa27,substr(pa28,1,8) as pa28,substr(pa29,1,8) as pa29,substr(pa30,1,8) as pa30,pa01,pa02,pa03,pa04 from r140412,patent where id='" & strUserNum & "' and (r1401=substr(pa26,1,8) or r1401=substr(pa27,1,8) or r1401=substr(pa28,1,8) or r1401=substr(pa29,1,8) or r1401=substr(pa30,1,8))) case" & _
                        " Where pa01=cp01 And pa02=cp02 And pa03=cp03 And pa04=cp04 and cp05+0<" & DBDATE(txt1(0)) & " and cp09||''<'B'" & _
                        " Union" & _
                        " select tm23,tm78,tm79,tm80,tm81,'" & strUserNum & "',tm01,tm02,tm03,tm04 from caseprogress," & _
                        " (select substr(tm23,1,8) as tm23,substr(tm78,1,8) as tm78,substr(tm79,1,8) as tm79,substr(tm80,1,8) as tm80,substr(tm81,1,8) as tm81,tm01,tm02,tm03,tm04 from r140412,trademark where id='" & strUserNum & "' and (r1401=substr(tm23,1,8) or r1401=substr(tm78,1,8) or r1401=substr(tm79,1,8) or r1401=substr(tm80,1,8) or r1401=substr(tm81,1,8))) case" & _
                        " Where tm01=cp01 And tm02=cp02 And tm03=cp03 And tm04=cp04 and cp05+0<" & DBDATE(txt1(0)) & " and cp09||''<'B' and cp10<>'001'" & _
                        " Union" & _
                        " select lc11,LC43,LC44,LC45,LC46,'" & strUserNum & "',lc01,lc02,lc03,lc04 from caseprogress," & _
                        " (select substr(lc11,1,8) as lc11,substr(lc43,1,8) as lc43,substr(lc44,1,8) as lc44,substr(lc45,1,8) as lc45,substr(lc46,1,8) as lc46,lc01,lc02,lc03,lc04 from r140412,lawcase where id='" & strUserNum & "' and (r1401=substr(lc11,1,8) or r1401=substr(lc43,1,8) or r1401=substr(lc44,1,8) or r1401=substr(lc45,1,8) or r1401=substr(lc46,1,8))) case" & _
                        " Where lc01=cp01 And lc02=cp02 And lc03=cp03 And lc04=cp04 and cp05+0<" & DBDATE(txt1(0)) & " and cp09||''<'B'" & _
                        " Union" & _
                        " select hc05,HC24,HC25,HC26,HC27,'" & strUserNum & "',hc01,hc02,hc03,hc04 from caseprogress," & _
                        " (select substr(hc05,1,8) as hc05,substr(hc24,1,8) as hc24,substr(hc25,1,8) as hc25,substr(hc26,1,8) as hc26,substr(hc27,1,8) as hc27,hc01,hc02,hc03,hc04 from r140412,hirecase where id='" & strUserNum & "' and (r1401=substr(hc05,1,8) or r1401=substr(hc24,1,8) or r1401=substr(hc25,1,8) or r1401=substr(hc26,1,8) or r1401=substr(hc27,1,8))) case" & _
                        " Where hc01=cp01 And hc02=cp02 And hc03=cp03 And hc04=cp04 and cp05+0<" & DBDATE(txt1(0)) & " and cp09||''<'B'" & _
                        " Union" & _
                        " select sp08,sp58,sp59,sp65,sp66,'" & strUserNum & "',sp01,sp02,sp03,sp04 from caseprogress," & _
                        " (select substr(sp08,1,8) as sp08,substr(sp58,1,8) as sp58,substr(sp59,1,8) as sp59,substr(sp65,1,8) as sp65,substr(sp66,1,8) as sp66,sp01,sp02,sp03,sp04 from r140412,servicepractice where id='" & strUserNum & "' and (r1401=substr(sp08,1,8) or r1401=substr(sp58,1,8) or r1401=substr(sp59,1,8) or r1401=substr(sp65,1,8) or r1401=substr(sp66,1,8))) case" & _
                        " Where sp01=cp01 And sp02=cp02 And sp03=cp03 And sp04=cp04 and cp05+0<" & DBDATE(txt1(0)) & " and cp09||''<'B'" & _
                        " )"
            cnnConnection.Execute strSql
            '刪除收文之移轉讓與受讓人的客戶資料
            'Modify By Sindy 2010/10/8 增加判斷本所案號也要相同才會剔除
            strSql = "delete r140412c where ID='" & strUserNum & "' AND (RC01||RC02||RC03||RC04||RC05||RC06||RC07||RC08||RC09) IN (" & _
                        " SELECT DISTINCT SUBSTR(CP56,1,8)||SUBSTR(CP89,1,8)||SUBSTR(CP90,1,8)||SUBSTR(CP91,1,8)||SUBSTR(CP92,1,8)||CP01||CP02||CP03||CP04" & _
                        " From CASEPROGRESS" & _
                        " WHERE CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " AND CP56 IS NOT NULL)"
            cnnConnection.Execute strSql
            '刪除以前曾收文案件之客戶
            strSql = "delete R140412 where ID='" & strUserNum & "' AND R1401 in (select distinct RC01 from R140412C where ID='" & strUserNum & "')"
            cnnConnection.Execute strSql
            strSql = "delete R140412 where ID='" & strUserNum & "' AND R1401 in (select distinct RC02 from R140412C where ID='" & strUserNum & "')"
            cnnConnection.Execute strSql
            strSql = "delete R140412 where ID='" & strUserNum & "' AND R1401 in (select distinct RC03 from R140412C where ID='" & strUserNum & "')"
            cnnConnection.Execute strSql
            strSql = "delete R140412 where ID='" & strUserNum & "' AND R1401 in (select distinct RC04 from R140412C where ID='" & strUserNum & "')"
            cnnConnection.Execute strSql
            strSql = "delete R140412 where ID='" & strUserNum & "' AND R1401 in (select distinct RC05 from R140412C where ID='" & strUserNum & "')"
            cnnConnection.Execute strSql
         End If
      End If
      cnnConnection.CommitTrans
      
      '2010/10/28 modify by sonia 加R140412直接串CP的受讓人,否則當月未發文之受讓人因串不到基本檔不會出現
      strExc(0) = "SELECT cuno,nvl(cu05,nvl(substr(cu04,1,14),substr(cu06,1,14))),na03,substr(a0902,1,6),cp01,sqldatet(substr(cpp,1,8)) from caseprogress,customer,acc090,nation,(SELECT R1401||'0' as cuno,MIN(cp) as cpp FROM (" & _
                              " select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,patent where ID='" & strUserNum & "' AND R1401=substr(pa26,1,8) and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,patent where ID='" & strUserNum & "' AND R1401=substr(pa27,1,8) and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,patent where ID='" & strUserNum & "' AND R1401=substr(pa28,1,8) and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,patent where ID='" & strUserNum & "' AND R1401=substr(pa29,1,8) and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,patent where ID='" & strUserNum & "' AND R1401=substr(pa30,1,8) and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,trademark where ID='" & strUserNum & "' AND R1401=substr(tm23,1,8) and tm01=cp01 and tm02=cp02 and tm03=cp03 and tm04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,trademark where ID='" & strUserNum & "' AND R1401=substr(tm78,1,8) and tm01=cp01 and tm02=cp02 and tm03=cp03 and tm04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,trademark where ID='" & strUserNum & "' AND R1401=substr(tm79,1,8) and tm01=cp01 and tm02=cp02 and tm03=cp03 and tm04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,trademark where ID='" & strUserNum & "' AND R1401=substr(tm80,1,8) and tm01=cp01 and tm02=cp02 and tm03=cp03 and tm04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,trademark where ID='" & strUserNum & "' AND R1401=substr(tm81,1,8) and tm01=cp01 and tm02=cp02 and tm03=cp03 and tm04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401"
      'Modify By Sindy 2011/2/24 增加LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27
      strExc(0) = strExc(0) & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,lawcase where ID='" & strUserNum & "' AND R1401=substr(lc11,1,8) and lc01=cp01 and lc02=cp02 and lc03=cp03 and lc04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,lawcase where ID='" & strUserNum & "' AND R1401=substr(lc43,1,8) and lc01=cp01 and lc02=cp02 and lc03=cp03 and lc04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,lawcase where ID='" & strUserNum & "' AND R1401=substr(lc44,1,8) and lc01=cp01 and lc02=cp02 and lc03=cp03 and lc04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,lawcase where ID='" & strUserNum & "' AND R1401=substr(lc45,1,8) and lc01=cp01 and lc02=cp02 and lc03=cp03 and lc04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,lawcase where ID='" & strUserNum & "' AND R1401=substr(lc46,1,8) and lc01=cp01 and lc02=cp02 and lc03=cp03 and lc04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,hirecase where ID='" & strUserNum & "' AND R1401=substr(hc05,1,8) and hc01=cp01 and hc02=cp02 and hc03=cp03 and hc04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,hirecase where ID='" & strUserNum & "' AND R1401=substr(hc24,1,8) and hc01=cp01 and hc02=cp02 and hc03=cp03 and hc04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,hirecase where ID='" & strUserNum & "' AND R1401=substr(hc25,1,8) and hc01=cp01 and hc02=cp02 and hc03=cp03 and hc04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,hirecase where ID='" & strUserNum & "' AND R1401=substr(hc26,1,8) and hc01=cp01 and hc02=cp02 and hc03=cp03 and hc04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,hirecase where ID='" & strUserNum & "' AND R1401=substr(hc27,1,8) and hc01=cp01 and hc02=cp02 and hc03=cp03 and hc04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401"
      '2011/2/24 End
      strExc(0) = strExc(0) & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,servicepractice where ID='" & strUserNum & "' AND R1401=substr(sp08,1,8) and sp01=cp01 and sp02=cp02 and sp03=cp03 and sp04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,servicepractice where ID='" & strUserNum & "' AND R1401=substr(sp58,1,8) and sp01=cp01 and sp02=cp02 and sp03=cp03 and sp04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,servicepractice where ID='" & strUserNum & "' AND R1401=substr(sp59,1,8) and sp01=cp01 and sp02=cp02 and sp03=cp03 and sp04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,servicepractice where ID='" & strUserNum & "' AND R1401=substr(sp65,1,8) and sp01=cp01 and sp02=cp02 and sp03=cp03 and sp04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress,servicepractice where ID='" & strUserNum & "' AND R1401=substr(sp66,1,8) and sp01=cp01 and sp02=cp02 and sp03=cp03 and sp04=cp04 and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress where ID='" & strUserNum & "' AND R1401=substr(cp56,1,8) and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress where ID='" & strUserNum & "' AND R1401=substr(cp89,1,8) and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress where ID='" & strUserNum & "' AND R1401=substr(cp90,1,8) and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress where ID='" & strUserNum & "' AND R1401=substr(cp91,1,8) and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " union select R1401,MIN(CP05||' '||CP09) cp from R140412,caseprogress where ID='" & strUserNum & "' AND R1401=substr(cp92,1,8) and CP09<'B' and substr(cp12,1,1)='F' and cp05>=" & DBDATE(txt1(0)) & " and cp05<=" & DBDATE(txt1(1)) & " group by R1401" & _
                        " ) GROUP by R1401) case Where cp09 = substr(cpp, 10, 9) and substr(cuno,1,8)=cu01(+) and substr(cuno,9,1)=cu02(+)" & _
                        " and cp12=a0901(+) and cu10=na01(+) order by 6,1"
   End If
   intI = 1
   Set m_rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If pub_QL04 <> "" Then InsertQueryLog (m_rs.RecordCount) 'Add By Sindy 2025/8/14
      Page = 0: m_iPages = 0: longCnt = 0
      '輸出方式
      If txt1(8) = "1" Then '1.螢幕
         m_bPrinter = False
         Set m_Device = Picture1
         m_Device.AutoRedraw = True
         m_Device.Width = 11500
         m_Device.Height = 16500
         DelPic
      Else '2.印表機
         m_bPrinter = True
         Set m_Device = Printer
         m_Device.Orientation = 1 '1.直印 2.橫印
      End If
      
      With m_rs
         .MoveFirst
         iLine = 1
         strType = ""
         Do While Not .EOF
            For m_i = 1 To 7
                strTemp(m_i) = ""
            Next m_i
            longCnt = longCnt + 1 '筆數
            strTemp(1) = CheckStr(m_rs.Fields(0))
            strTemp(2) = Left(CheckStr(m_rs.Fields(1)) & "                    ", 20)
            strTemp(3) = Left(CheckStr(m_rs.Fields(2)) & "      ", 6)
            strTemp(4) = Left(CheckStr(m_rs.Fields(3)) & "      ", 6)
            strTemp(5) = CheckStr(m_rs.Fields(4))
            strTemp(6) = CheckStr(m_rs.Fields(5))
            '外專
            If Left(Trim(strTemp(4)), 2) = "外專" Then
               dblTot1 = dblTot1 + 1
            '外商
            ElseIf Left(Trim(strTemp(4)), 2) = "外商" Then
               dblTot2 = dblTot2 + 1
            '投資法務
            Else
               dblTot3 = dblTot3 + 1
            End If
            If (iLine > 52 And txt1(7) = "1") Or iLine = 1 Then
               iLine = 1
               PrintTitle '列印表頭
            End If
            If txt1(7) = "1" Then '1.明細
               PrintDetail
            End If
            strType = CheckStr(m_rs.Fields(0))
            .MoveNext
         Loop
         '合計
         If txt1(7) = "1" Then '1.明細
            m_Device.CurrentX = PLeft(1)
            m_Device.CurrentY = iLine * 300
            m_Device.Print String(140, "-")
            iLine = iLine + 1
         End If
         m_Device.CurrentX = PLeft(2)
         m_Device.CurrentY = iLine * 300
         m_Device.Print "外專  " & dblTot1 & "  筆　　外商  " & dblTot2 & "  筆　　投資法務  " & dblTot3 & "  筆 "
         m_Device.CurrentX = PLeft(5)
         m_Device.CurrentY = iLine * 300
         m_Device.Print "合計  " & longCnt & "  筆"
         iLine = iLine + 1
      End With
      If m_bPrinter = True Then
         m_Device.EndDoc
         ShowPrintOk
      Else
         If m_iPages > 0 Then
            SetPic m_iPages
            frm140412_1.m_ImageW = m_Device.Width
            frm140412_1.m_ImageH = m_Device.Height
            frm140412_1.m_iPages = m_iPages
            frm140412_1.Show
         End If
      End If
   Else
      If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/14
      MsgBox "無可列印資料！"
   End If
   Set m_rs = Nothing
   Set m_Device = Nothing
ErrHnd:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub DelPic()
   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_*.tmp"
   If Dir(strPicFileName) <> "" Then
      Kill strPicFileName
   End If
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
End Sub

Private Sub SetPic(idx As Integer)
   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_" & idx & ".tmp"
   
   SavePicture Picture1.Image, strPicFileName
   '要用覆蓋的否則會錯誤--VB Bug
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 2000
PLeft(3) = 5500
PLeft(4) = 7000
PLeft(5) = 8800
PLeft(6) = 10000
End Sub

Sub PrintTitle()
GetPleft

Page = Page + 1
m_iPages = m_iPages + 1

If m_iPages > 1 Then
   If m_bPrinter = False Then
      SetPic m_iPages - 1
   ElseIf Page > 1 Then
      m_Device.NewPage
   End If
End If

m_Device.Font.Size = 12
m_Device.Font.Underline = False
m_Device.FontBold = False

If txt1(4) = "1" Then '1.代理人
   m_Device.CurrentX = m_Device.ScaleWidth / 2 - (m_Device.TextWidth("國外部新代理人") / 2)
   m_Device.CurrentY = iLine * 300
   m_Device.Print "國外部新代理人"
Else
   m_Device.CurrentX = m_Device.ScaleWidth / 2 - (m_Device.TextWidth("國外部新客戶") / 2)
   m_Device.CurrentY = iLine * 300
   m_Device.Print "國外部新客戶"
End If
m_Device.CurrentX = m_Device.ScaleWidth - m_Device.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
m_Device.CurrentY = 600
m_Device.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
m_Device.CurrentX = PLeft(1)
m_Device.CurrentY = 900
m_Device.Print "列印人員：" & strUserName
m_Device.CurrentX = m_Device.ScaleWidth / 2 - (m_Device.TextWidth("收文日：" & ChangeTStringToTDateString(txt1(0)) & " - " & ChangeTStringToTDateString(txt1(1))) / 2)
m_Device.CurrentY = 900
m_Device.Print "收文日：" & ChangeTStringToTDateString(txt1(0)) & " - " & ChangeTStringToTDateString(txt1(1))
m_Device.CurrentX = m_Device.ScaleWidth - m_Device.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
m_Device.CurrentY = 900
m_Device.Print "頁　　次：" & str(Page)
iLine = 5
m_Device.CurrentX = PLeft(1)
m_Device.CurrentY = iLine * 300
If txt1(4) = "1" Then '1.代理人
   m_Device.Print "代理人編號"
Else
   m_Device.Print "客戶編號"
End If
m_Device.CurrentX = PLeft(2)
m_Device.CurrentY = iLine * 300
m_Device.Print "名稱"
m_Device.CurrentX = PLeft(3)
m_Device.CurrentY = iLine * 300
m_Device.Print "國籍"
m_Device.CurrentX = PLeft(4)
m_Device.CurrentY = iLine * 300
m_Device.Print "業務區"
m_Device.CurrentX = PLeft(5)
m_Device.CurrentY = iLine * 300
m_Device.Print "系統類別"
m_Device.CurrentX = PLeft(6)
m_Device.CurrentY = iLine * 300
m_Device.Print "最早收文日"
iLine = iLine + 1
m_Device.CurrentX = PLeft(1)
m_Device.CurrentY = iLine * 300
m_Device.Print String(140, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
For m_j = 1 To 6
   m_Device.CurrentX = PLeft(m_j)
   m_Device.CurrentY = iLine * 300
   m_Device.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   '收文日
   txt1(0) = strSrvDate(2)
   txt1(1) = strSrvDate(2)
   '對象--預設代理人
   txt1(4) = "1"
   '查詢別--預設明細
   txt1(7) = "1"
   '輸出方式--預設印表機
   txt1(8) = "2"
   '業務區
   If Mid(Trim(Pub_StrUserSt03), 1, 1) = "F" Then '國外部人員
      txt1(5) = "F"
      txt1(6) = "F99"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140412 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 4, 7, 8 '4.對象 7.查詢別 8.輸出方式
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
      Case 1, 3, 6 '1.收文日 3.國籍 6.業務區
         If blnClkSure = False Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               txt1(Index - 1).SetFocus
               txt1_GotFocus (Index - 1)
               Exit Sub
            End If
         Else
            blnClkSure = False
         End If
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1 '收文日
         If txt1(Index) = "" Then Exit Sub
         Cancel = Not ChkDate(txt1(Index))
         If Cancel Then TextInverse txt1(Index)
         
      Case 4 '對象
         Select Case Val(txt1(Index))
           Case 1, 2
           Case Else
              s = MsgBox("對象只能 1 或 2 !!", , "USER 輸入錯誤")
              Cancel = True
         End Select
         
      Case 7 '查詢別
         Select Case Val(txt1(Index))
           Case 1, 2
           Case Else
              s = MsgBox("查詢別只能 1 或 2 !!", , "USER 輸入錯誤")
              Cancel = True
         End Select
         
      Case 8 '輸出方式
         Select Case Val(txt1(Index))
           Case 1, 2
           Case Else
              s = MsgBox("輸出方式只能 1 或 2 !!", , "USER 輸入錯誤")
              Cancel = True
         End Select
         
   End Select
   If Cancel Then TextInverse txt1(Index)
End Sub
