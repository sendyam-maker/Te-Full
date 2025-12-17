VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc44w0 
   AutoRedraw      =   -1  'True
   Caption         =   "代填繳款書客戶明細"
   ClientHeight    =   3024
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5508
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3024
   ScaleWidth      =   5508
   Begin VB.CommandButton CmdMemo 
      Caption         =   "說明"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      TabIndex        =   15
      Top             =   516
      Width           =   600
   End
   Begin VB.CheckBox chkCon 
      Caption         =   "不含單筆收款扣繳合計未達2000元 (Excel條件)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   50
      TabIndex        =   14
      Top             =   1850
      Width           =   5000
   End
   Begin VB.CheckBox chkA4228 
      Caption         =   "客戶有勾選每筆代繳"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   1764
      TabIndex        =   12
      Top             =   1248
      Width           =   2625
   End
   Begin VB.CheckBox chkA4228 
      Caption         =   "單筆收據稅額超過2000元"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   1
      Left            =   1764
      TabIndex        =   11
      Top             =   1500
      Width           =   2625
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2880
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   2136
      Width           =   1755
   End
   Begin VB.ComboBox Combo3 
      Height          =   276
      Left            =   1764
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   2640
      Width           =   3450
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel(&E)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   870
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   2136
      Width           =   1755
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   336
      Left            =   1764
      TabIndex        =   0
      Top             =   516
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   593
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   336
      Left            =   3216
      TabIndex        =   1
      Top             =   516
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   593
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   336
      Left            =   1764
      TabIndex        =   5
      Top             =   60
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   593
      _Version        =   393216
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   336
      Left            =   3216
      TabIndex        =   6
      Top             =   60
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   593
      _Version        =   393216
      BackColor       =   -2147483633
      AllowPrompt     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      Caption         =   "可複選，未勾選即為全部代繳名單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   288
      Left            =   1560
      TabIndex        =   13
      Top             =   960
      Width           =   3204
   End
   Begin VB.Label Label4 
      Caption         =   "客戶代填方式："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   288
      Left            =   50
      TabIndex        =   10
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "地址條印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   50
      TabIndex        =   9
      Top             =   2700
      Width           =   1488
   End
   Begin VB.Label Label2 
      Caption         =   "收款日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   50
      TabIndex        =   8
      Top             =   516
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "上次收款日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   50
      TabIndex        =   7
      Top             =   60
      Width           =   1500
   End
   Begin VB.Line Line2 
      X1              =   2940
      X2              =   3150
      Y1              =   216
      Y2              =   216
   End
   Begin VB.Line Line1 
      X1              =   2940
      X2              =   3150
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   120
      Top             =   876
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc44w0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/03/01 改成Form2.0 ; 地址條改成Excel列印
'Create By Sindy 2016/11/9
Option Explicit

Dim adoquery As New ADODB.Recordset
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim intCounter As Integer
Dim lngPageNo As Long '頁數
Dim m_strT01 As String, m_strT15 As String, m_strT19 As String
Dim strPrinter As String
Dim intTitleR As String 'Add by Amy 2025/02/20

'Add by Amy 2025/04/01 +說明-瑞婷
Private Sub CmdMemo_Click()
   Frmacc44w2.Show
End Sub

Private Sub Command1_Click()
   Call Process
End Sub

Private Sub Process()
Dim rsA As New ADODB.Recordset
Dim strT01 As String, strT02 As String, strT05 As String
Dim strT06 As String, strT14 As String, strT15 As String
Dim m_CU01 As String, m_CU02 As String
Dim m_CU30 As String, m_CU31 As String, m_CU16 As String
Dim m_CU159 As String, m_CU168 As String
Dim m_CU169 As String, m_CU170 As String
Dim m_CU171 As String, m_CU11 As String
Dim strT21 As String, strT25 As String
Dim strT26 As String, strT16 As String
Dim strT20 As String, strT19 As String
Dim m_CU181 As String
Dim strTempAddressList As String 'Added by Lydia 2022/03/01
Dim hLocalFile As Long, strCmd As String, stOldComp As String, stMsg As String 'Add by Amy 2025/02/20

On Error GoTo ErrHnd
   
   If FormCheck = False Then
'      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If

   Screen.MousePointer = vbHourglass
   
   m_strT01 = "": m_strT15 = "": m_strT19 = ""
   
   '產生暫存檔資料
   '**********************************************
   'ACCTMP44w0:
   '**********************************************
   'T01:客戶編號 key
   'T02:收據編號 key
   'T05:Me.Name  key
   'T06:收款單號 key
   'T07:票據號碼 key
   'T14:strUserName key
   'T15:收據抬頭
   '**********************************************
   adoTaie.Execute "delete from ACCTMP44w0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
   
   '收款收據資料
   'Modify by Amy 2025/02/20 +strCmd/T23:a0k11(收據公司別)
   'Modify by Amy 2025/03/07 避免程式跑很久,調語法-秀玲
   '              加條件:a0k05='2',可扣繳的收據 / A0k01=A1v02(+) and a1v02 is not null,沒有ACC1V0的資料就代表不需要扣繳 / A1v06=0,先過濾就不用再跑迴圈刪除
'   strCmd = "insert into ACCTMP44w0(T01,T02,T05,T06,T07,T14,T15,T08,T23)" & _
'                   " select distinct a0k03,a0k01,'" & Me.Name & "',a0l01,nvl(a1p09,'X'),'" & strUserNum & "',a0k04 T15,a1p12,a0k11" & _
'                   " From acc0l0, acc0m0, acc0k0, Acc1p0" & _
'                   " Where a0l02 >=" & ACDate(DBDATE(MaskEdBox1)) & " And a0l02 <=" & ACDate(DBDATE(MaskEdBox2)) & _
'                   " and a0l01=a0m01(+) and nvl(a0m06,0)=0" & _
'                   " and a0m02=a0k01(+) and a0k37='Y' and a0k11<>'J'" & _
'                   " and a0l01=a1p04(+)" 'and a1p09 is not null
   strCmd = "insert into ACCTMP44w0(T01,T02,T05,T06,T07,T14,T15,T08,T23)" & _
                   " select distinct a0k03,a0k01,'" & Me.Name & "',a0l01,nvl(a1p09,'X'),'" & strUserNum & "',a0k04 T15,a1p12,a0k11" & _
                   " From acc0l0, acc0m0, acc0k0, Acc1p0, acc1v0" & _
                   " Where a0l02 >=" & ACDate(DBDATE(MaskEdBox1)) & " And a0l02 <=" & ACDate(DBDATE(MaskEdBox2)) & _
                   " and a0l01=a0m01(+) and nvl(a0m06,0)=0" & _
                   " and a0m02=a0k01(+) and a0k37='Y' and a0k11<>'J' and a0k05='2' And A0k01=A1v02(+) and a1v02 is not null And A1v06=0" & _
                   " and a0l01=a1p04(+)" 'and a1p09 is not null
   adoTaie.Execute strCmd
   'end 2025/02/20
'      " select distinct a0k03,a0k01,'" & Me.Name & "',a0l01,'" & strUserNum & "',a0k04 T15,' '" & _
'      " From acc0l0,acc0m0,acc0k0" & _
'      " Where a0l02 >=" & ACDate(DBDATE(MaskEdBox1)) & " And a0l02 <=" & ACDate(DBDATE(MaskEdBox2)) & _
'      " and a0l01=a0m01(+) and nvl(a0m06,0)=0" & _
'      " and a0m02=a0k01(+) and a0k37='Y' and a0k11<>'J'"
   '逐筆依收據抬頭讀取繳款書資料
   strExc(0) = "select T15 from ACCTMP44w0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
               " group by T15"
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsA.MoveFirst
      Do While rsA.EOF = False
         'Add By Sindy 2020/10/6 + m_CU16
         If GetTitleCustData(rsA.Fields("T15"), "", "", m_CU01, m_CU02, _
                            "", "", "", "", "", "", m_CU16, _
                            "", "", "", "", "", m_CU30, m_CU31, _
                            m_CU159, "", "", , m_CU168, m_CU169, m_CU170, m_CU171, m_CU11, _
                            , , , , , m_CU181) = True Then
         End If
         
         'Modify by Amy 2025/02/20 每月代填繳款書 欄位原m_CU168<> "Y",改存智慧所/法律所
         If m_CU168 = "" Then '非每月提醒代填繳款書,則刪除
            adoTaie.Execute "delete from ACCTMP44w0 where T05='" & Me.Name & "' and T14='" & strUserNum & "' and T15='" & ChgSQL(rsA.Fields("T15")) & "'"
         Else
            '更新資料:
            'T21:地址 IIf(InStr(m_CU31, m_CU30) > 0, m_CU31, m_CU30 & m_CU31)
            'T25:繳款書寄件處 m_CU169
            'T26:會計備註   m_CU159
            'T16:收件人     m_CU171
            'T20:繳款書地址 m_CU170
            'T19:統一編號   m_CU11
            'T24:繳款書代填方式 m_CU181 :1.每筆代繳2.單筆收據稅額超過2000元
            'T17:電話 Add By Sindy 2020/10/6
            'T18:每月代填繳款書公司別 'Add by Amy 2025/02/20
            adoTaie.Execute "update ACCTMP44w0 set" & _
                            " T21='" & IIf(InStr(m_CU31, m_CU30) > 0, m_CU31, m_CU30 & m_CU31) & "'" & _
                            ",T25='" & m_CU169 & "'" & _
                            ",T26='" & ChgSQL(m_CU159) & "'" & _
                            ",T16='" & m_CU171 & "'" & _
                            ",T20='" & m_CU170 & "'" & _
                            ",T19='" & m_CU11 & "'" & _
                            ",T24='" & m_CU181 & "'" & _
                            ",T17='" & m_CU16 & "'" & _
                            ",T18='" & m_CU168 & "'" & _
                            " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and T15='" & ChgSQL(rsA.Fields("T15")) & "'"
         End If
         
         rsA.MoveNext
      Loop
   End If
   
   'Mark by Amy 2025/03/07 避免程式跑很久,調語法後不用再跑迴圈刪除-秀玲
'   '逐筆依收款單號讀取票據資料
'   strExc(0) = "select t02 from ACCTMP44w0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
'               " group by T02"
'   intI = 1
'   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      rsA.MoveFirst
'      Do While rsA.EOF = False
'         'Modify By Sindy 2017/1/5
'         '踢除已有扣繳金額
'         strExc(0) = "select a1v01" & _
'                     " from acc1v0" & _
'                     " where a1v02='" & rsA.Fields("T02") & "' and a1v06>0"
'         intI = 1
'         Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            adoTaie.Execute "delete from ACCTMP44w0 where T05='" & Me.Name & "' and T14='" & strUserNum & "' and T02='" & rsA.Fields("T02") & "'"
''         Else
''         '2017/1/5 END
''            strT01 = "" & rsA.Fields("T01")
''            strT02 = "" & rsA.Fields("T02")
''            strT05 = "" & rsA.Fields("T05")
''            strT06 = "" & rsA.Fields("T06")
''            strT14 = "" & rsA.Fields("T14")
''            strT15 = "" & rsA.Fields("T15")
''            strT21 = "" & rsA.Fields("T21")
''            strT25 = "" & rsA.Fields("T25")
''            strT26 = "" & rsA.Fields("T26")
''            strT16 = "" & rsA.Fields("T16")
''            strT20 = "" & rsA.Fields("T20")
''            strT19 = "" & rsA.Fields("T19")
''
''            'T07:票據號碼
''            'T08:到期日期
''            strExc(0) = "select a1p09,a1p12" & _
''                        " from acc1p0" & _
''                        " where a1p04='" & strT06 & "' and a1p09 is not null"
''            intI = 1
''            Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
''            If intI = 1 Then
''               adoquery.MoveFirst
''               adoTaie.Execute "update ACCTMP44w0 set" & _
''                               " T07='" & adoquery.Fields("a1p09") & "'" & _
''                               ",T08=" & "" & adoquery.Fields("a1p12") & _
''                               " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and T15='" & ChgSQL(rsA.Fields("T15")) & "'" & _
''                               " and T06 = '" & strT06 & "'"
''               adoquery.MoveNext
''               'Modify By Sindy 2017/8/2 Mark
'''               Do While Not adoquery.EOF
'''                  adoTaie.Execute "insert into ACCTMP44w0(T01,T02,T05,T06,T07,T08,T14,T15,T21,T25,T26,T16,T20,T19)" & _
'''                                  " values('" & strT01 & "','" & strT02 & "','" & strT05 & "'" & _
'''                                  ",'" & strT06 & "','" & adoquery.Fields("a1p09") & "'," & _
'''                                  adoquery.Fields("a1p12") & ",'" & strT14 & "','" & ChgSQL(strT15) & _
'''                                  "','" & strT21 & "','" & strT25 & "','" & strT26 & "','" & strT16 & _
'''                                  "','" & strT20 & "','" & strT19 & "')"
'''                  adoquery.MoveNext
'''               Loop
''            End If
'         End If
'         rsA.MoveNext
'      Loop
'   End If
'end 2025/03/07
   
   'Add by Amy 2025/02/20 +代填方式
   strCmd = "Delete From ACCTMP44w0 where T05='" & Me.Name & "' and T14='" & strUserNum & "' "
   If chkA4228(0).Value = 1 And chkA4228(1).Value = 1 Then
      adoTaie.Execute strCmd & " and T24 is Null "
   Else
      '勾選只出現[每筆代繳]
      If chkA4228(0).Value = 1 Then
         adoTaie.Execute strCmd & " and (T24 is Null or T24='2') "
      End If
      '勾選只出現[單筆收據稅額超過2000元]
      If chkA4228(1).Value = 1 Then
         adoTaie.Execute strCmd & " and (T24 is Null or T24='1') "
      End If
   End If
   'end 2025/02/20
   'Modify by Amy 2025/03/28 畫面原「含單筆收款扣繳合計未達2000元,[不勾]才加過濾條件」改不含
   'Modify by Amy 2025/03/04 +含單筆收款扣繳合計未達2000元,[不勾]才加過濾條件
   '               若取消勾選則,資料只顯示同一收款單號之扣款金額2000元(含)以上之資料(不管收據公司別)-秀玲
   'If chkCon.Value = 0 Then
   'end 2025/03/28
   If chkCon.Value = 1 Then
      'Modify by Amy 2025/03/28 +And (T24 is Null or T24='2')  因[每筆代繳],不因畫面勾選「不含單筆收款扣繳合計未達2000元 (Excel條件)」影響
      strExc(9) = " And T06 in (Select t06 From acc1v0,(Select Distinct T06,T02 From Acctmp44w0 Where T05='" & Me.Name & "' and T14='" & strUserNum & "' And (T24 is Null or T24='2') ) " & _
                                                   "Where T02=a1v02(+) Group by T06 Having Sum(a1v04)<2000 ) "
      adoTaie.Execute strCmd & strExc(9)
   End If
   
   lngPageNo = 0
   'Add By Sindy 2020/10/6 + ,T17
   'Modify by Amy 2025/02/20 +T18/T23,依[收據公司別]產生2個檔-瑞婷
   strExc(0) = "select T24,T01,T15,T19,T20,T16,T07,T08,T26,T17,T18,Replace(T23,'2','1') as Comp" & _
               " from ACCTMP44w0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
               " group by Replace(T23,'2','1'),T24,T01,T15,T19,T20,T16,T07,T08,T26,T17,T18" & _
               " order by Replace(T23,'2','1'),T15,T01"
   intI = 1
   Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoquery
         .MoveFirst
         Set xlsAnnuity = New Excel.Application
         'Mark by Amy 2025/02/20  依[收據公司別]產生2個檔,設定Excel 及 頁首 往下搬
'         Call SetExcelWorksheets
'         PrintHead_Excel intCounter '頁首
         'end 2025/02/20
         Do While Not .EOF
            'Add by Amy 2025/02/20 依[收據公司別]產生2個檔,拿掉換頁-瑞婷
            If stOldComp <> .Fields("Comp") Then
               If stOldComp <> MsgText(601) Then
                  Call SaveExcel(stOldComp, False) '存檔
                  stMsg = stMsg & "," & stOldComp
               End If
               Call SetExcelWorksheets
               PrintHead_Excel intCounter '頁首
               intTitleR = intCounter
               lngPageNo = 0
            End If
            
'            '第2頁切頁有誤 +  And intCounter <> 48 判斷
'            If (lngPageNo = 1 And intCounter Mod 32 = 0) Or _
'               (lngPageNo <> 1 And intCounter Mod 32 = 0 And intCounter <> 32) Then
'               '換頁
'               intCounter = intCounter + 1
'               wksAnnuity.Range("A" & intCounter).Select
'               wksAnnuity.HPageBreaks.add Before:=wksAnnuity.Application.ActiveCell
'               PrintHead_Excel intCounter '頁首
'            End If
            'end 2025/02/20
            '明細資料
            PrintData_Excel adoquery, intCounter
            stOldComp = .Fields("Comp")
            .MoveNext
         Loop
         Call SaveExcel(stOldComp, True) 'Add by Amy 2025/02/20 存檔
         stMsg = stMsg & "," & stOldComp
      End With
   Else
      Screen.MousePointer = vbDefault
      MsgBox "無資料可供列印！"
      adoquery.Close
      Set adoquery = Nothing
      Exit Sub
   End If
   
   'Mark by Amy 2025/02/20 改成存檔,不顯示畫面上
'   xlsAnnuity.Visible = True
'   xlsAnnuity.WindowState = wdWindowStateMaximize

   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   
   '列印地址條
   If MsgBox("是否要列印地址條？" & vbCrLf & _
         "若要印，請放地址條貼紙於選取的印表機!!", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
      strExc(0) = "select T15,T20,T16 from ACCTMP44w0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                  " group by T15,T20,T16" & _
                  " order by T15"
      intI = 1
      Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         PUB_SetOsDefaultPrinter Combo3 'Added by Lydia 2022/03/01 切換Word/Excel印表機
         PUB_RestorePrinter Combo3
         With adoquery
            .MoveFirst
            Do While Not .EOF
               'Modified by Lydia 2022/03/01 傳入多張地址條的內容；用|區隔不同張地址條，同一張地址條用$區隔地址和收件人
               'Call PUB_PrintAccAddress(.Fields("T20"), .Fields("T16"))
               If "" & .Fields("T20") & .Fields("T16") <> "" Then strTempAddressList = strTempAddressList & Trim(.Fields("T20")) & "$" & Trim(.Fields("T16")) & "|"
               DoEvents
               .MoveNext
            Loop
         End With
         'Added by Lydia 2022/03/01 改用Execl列印地址條
         If strTempAddressList <> "" Then
             If PUB_XlsAccAddress(strTempAddressList) = False Then
                 MsgBox "列印失敗！", vbCritical
             End If
         End If
         'end 2022/03/01
         PUB_SetOsDefaultPrinter strPrinter 'Added by Lydia 2022/03/01 切換Word/Excel印表機
         PUB_RestorePrinter strPrinter
         Screen.MousePointer = vbDefault
         MsgBox "列印完畢!!", vbInformation
      End If
   End If
   
   'Add by Amy 2025/02/20 改成存檔,不顯示畫面上
   If stMsg <> MsgText(601) Then
      stMsg = Replace(Replace(Mid(stMsg, 2), "1", "客戶代填繳書明細表-智慧所"), "L", "客戶代填繳書明細表-法律所")
      stMsg = "已產生檔案如下：" & vbCrLf & _
                        Replace(stMsg, ",", vbCrLf) & vbCrLf & vbCrLf
   End If
   If MsgBox(stMsg & "是否開啟資料夾？", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
      ShellExecute hLocalFile, "explore", strExcelPath, vbNullString, vbNullString, 1
   End If
   
   Screen.MousePointer = vbDefault

   PUB_SaveLastDate Me.Name, "MaskEdBox3", ChangeTDateStringToTString(MaskEdBox1)
   PUB_SaveLastDate Me.Name, "MaskEdBox4", ChangeTDateStringToTString(MaskEdBox2)
   MaskEdBox3.Text = ChangeTStringToTDateString(PUB_GetLastDate(Me.Name, "MaskEdBox3"))
   MaskEdBox4.Text = ChangeTStringToTDateString(PUB_GetLastDate(Me.Name, "MaskEdBox4"))
   MaskEdBox1.Mask = ""
   MaskEdBox2.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox2.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   
   Set rsA = Nothing
   Set adoquery = Nothing
   Exit Sub

ErrHnd:
   Screen.MousePointer = vbDefault
   Set rsA = Nothing
   Set adoquery = Nothing
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing

   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub SetExcelWorksheets()
   xlsAnnuity.Visible = True
   'Modify by Amy 2025/02/20 改回預設
   'xlsAnnuity.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsAnnuity.SheetsInNewWorkbook = 3
   'end 2025/02/20
   xlsAnnuity.Workbooks.add
   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   wksAnnuity.PageSetup.Orientation = xlLandscape '橫印
   'wksAnnuity.PageSetup.Orientation = wdOrientLandscape '直印
   wksAnnuity.PageSetup.LeftMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.RightMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.TopMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.BottomMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.HeaderMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.FooterMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   '設定各欄位長度
   'Modify by Amy 2025/03/04 原:10-文字全顯示
   wksAnnuity.Columns("A:A").ColumnWidth = 13 '代填方式
   wksAnnuity.Columns("B:B").ColumnWidth = 25 '收據抬頭
   wksAnnuity.Columns("C:C").ColumnWidth = 9 '統一編號
   'Modify by Amy 2025/02/20 增加欄寬 原11
   wksAnnuity.Columns("D:D").ColumnWidth = 18 '電話
   wksAnnuity.Columns("E:E").ColumnWidth = 10 '票號
   wksAnnuity.Columns("F:F").ColumnWidth = 10 '票據到期日
   'Modify by Amy 2025/02/20 +同意書,縮小欄寬-繳款書地址 原30/收件人 原25 /會計備註 原14
   '              瑞婷原想隱藏欄位,但因抓Acc1p0資料若為收票資料會出現多筆,造成1筆資料正常,另1筆只有--- 很怪 ex:1公司D114011257
   wksAnnuity.Columns("G:G").ColumnWidth = 15 '繳款書地址
   'Modify by Amy 2025/03/04 原:15 (讓「代填方式」文字全顯示
   wksAnnuity.Columns("H:H").ColumnWidth = 14 '收件人
   wksAnnuity.Columns("I:I").ColumnWidth = 14 '會計備註
   wksAnnuity.Columns("J:J").ColumnWidth = 7.5 '同意書
   'end 2025/02/20
   
   wksAnnuity.Range("C:C").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("F:F").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("G:G").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   
   intCounter = 1
End Sub

'表頭
Private Sub PrintHead_Excel(ByRef iRow As Integer)
Dim i As Integer, strTemp As String
  
   lngPageNo = lngPageNo + 1
   With wksAnnuity
      .Range("E" & iRow).Value = "代填繳款書客戶明細"
      '選取,儲存格合併,置中,粗體字
      strTemp = "A" & iRow & ":I" & iRow
      .Range(strTemp).Select
      With .Application.Selection
          .HorizontalAlignment = xlGeneral
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      With .Application.Selection
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      .Application.Selection.Font.Bold = True

      iRow = iRow + 1
      .Range("A" & iRow).Value = "列印人：" & strUserName
      .Range("D" & iRow).Value = "收款日期：" & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
      .Range("G" & iRow).Value = "列印日期："
      .Range("H" & iRow).Value = Format(strSrvDate(2), "###/##/##")
      iRow = iRow + 1
      .Range("G" & iRow).Value = "頁數："
      .Range("H" & iRow).Value = lngPageNo
      strTemp = "D" & iRow - 1 & ":D" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlCenter '置中
      End With
      strTemp = "F" & iRow - 1 & ":F" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlRight '靠右
      End With
      strTemp = "G" & iRow & ":H" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlLeft '靠左
      End With
      'Add by Amy 2025/02/20 +同意書
      .Range("J" & iRow).Value = "X:無檔案"
      .Range("J" & iRow).Font.Size = 9
      .Range("J" & iRow).HorizontalAlignment = xlCenter
      'end 2025/02/20
      
      iRow = iRow + 1
      .Range("A" & iRow).Value = "代填方式" '"客戶編號"
      .Range("B" & iRow).Value = "收據抬頭"
      .Range("C" & iRow).Value = "統一編號"
'      .Range("D" & iRow).Value = "繳款書地址"
'      .Range("E" & iRow).Value = "收件人"
'      .Range("F" & iRow).Value = "票號"
'      .Range("G" & iRow).Value = "票據到期日"
      .Range("D" & iRow).Value = "電話" 'Add By Sindy 2020/10/6
      .Range("E" & iRow).Value = "票號"
      .Range("F" & iRow).Value = "票據到期日"
      'Memo by Amy 2025/02/20 縮小欄寬-繳款書地址/收件人/會計備註(於 SetExcelWorksheets 設定欄位)
      .Range("G" & iRow).Value = "繳款書地址"
      .Range("H" & iRow).Value = "收件人"
      .Range("I" & iRow).Value = "會計備註"
      
      'Add by Amy 2025/02/20 +同意書
      .Range("J" & iRow).Value = "同意書"
      strTemp = "A" & iRow & ":J" & iRow
      'end 2025/02/20
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlCenter '置中
      End With
'      With .Application.Selection.Borders(xlEdgeLeft)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlEdgeTop)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
      With .Application.Selection.Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
'      With .Application.Selection.Borders(xlEdgeRight)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlInsideVertical)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
   End With
End Sub

Private Sub PrintData_Excel(p_Rst As ADODB.Recordset, ByRef iRow As Integer)
Dim strTemp As String
   
   iRow = iRow + 1
   With wksAnnuity
      If m_strT01 = "" & p_Rst.Fields("T01") And _
         m_strT15 = "" & p_Rst.Fields("T15") And _
         m_strT19 = "" & p_Rst.Fields("T19") Then
         .Range("A" & iRow).Value = "---"
         .Range("B" & iRow).Value = "---"
         .Range("C" & iRow).Value = "---"
         .Range("D" & iRow).Value = "---" 'Add By Sindy 2020/10/6
         .Range("G" & iRow).Value = "---"
         .Range("H" & iRow).Value = "---"
      Else
         'Modify By Sindy 2019/12/18
         '.Range("A" & iRow).Value = "" & p_Rst.Fields("T01") '客戶編號
         If p_Rst.Fields("T24") = "1" Then '代填方式
            .Range("A" & iRow).Value = "每筆代繳"
         ElseIf p_Rst.Fields("T24") = "2" Then
            'Modify by Amy 2025/03/04 字太長縮減,讓文字可全顯示
            '.Range("A" & iRow).Value = "單筆收據稅額超過2000元"
            .Range("A" & iRow).Value = "單筆超過2000元"
         Else
            .Range("A" & iRow).Value = "" & p_Rst.Fields("T24")
         End If
         '2019/12/18 END
          .Range("A" & iRow).Font.Size = 10 'Add by Amy 2025/03/04 字完整顯示(不要換行)-瑞婷
          
         .Range("B" & iRow).Value = "" & p_Rst.Fields("T15") '收據抬頭
         .Range("C" & iRow).Value = "" & p_Rst.Fields("T19") '統一編號
         .Range("D" & iRow).Value = "" & p_Rst.Fields("T17") '電話 Add By Sindy 2020/10/6
         .Range("G" & iRow).Value = "" & p_Rst.Fields("T20") '繳款書地址
         .Range("H" & iRow).Value = "" & p_Rst.Fields("T16") '收件人
      End If
      .Range("E" & iRow).Value = IIf("" & p_Rst.Fields("T07") = "X", "", "" & p_Rst.Fields("T07")) '票號
      .Range("F" & iRow).Value = ChangeTStringToTDateString("" & p_Rst.Fields("T08")) '票據到期日
      .Range("I" & iRow).Value = "" & p_Rst.Fields("T26") '會計備註
      'Add by Amy 2025/02/20 每月代填繳款書公司別=目前[收據公司別]時,判斷[無]同意書顯示X
      .Range("J" & iRow).Value = " " '避免會計備註資料過長,看不出是否有同意書
      If "" & p_Rst.Fields("T18") <> MsgText(601) And InStr("" & p_Rst.Fields("T18"), "" & p_Rst.Fields("Comp")) > 0 Then
         If ChkWithholdingTaxConsent(0, Me.Name, "" & p_Rst.Fields("Comp"), "" & p_Rst.Fields("T15")) = False Then
            .Range("J" & iRow).Value = "X"
         End If
      End If
      'end 2025/02/20
      'Modify by Amy 2025/02/20 原:I
      .Range("A" & iRow & ":J" & iRow).Select
      .Application.Selection.VerticalAlignment = xlTop '靠上
      .Range("I" & iRow).Select
      '.Application.Selection.WrapText = True '自動換行'Mark by Amy 2025/02/20 先拿掉換行-瑞婷
      
'      strTemp = "A" & iRow & ":I" & iRow
'      .Range(strTemp).Select
'      With .Application.Selection.Borders(xlEdgeLeft)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlEdgeTop)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlEdgeBottom)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlEdgeRight)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlInsideVertical)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
   End With
   m_strT01 = "" & p_Rst.Fields("T01")
   m_strT15 = "" & p_Rst.Fields("T15")
   m_strT19 = "" & p_Rst.Fields("T19")
End Sub

'Add By Sindy 2017/6/19
Private Sub Command2_Click()
   Dim stCon As String 'Add by Amy 2025/03/04 畫面條件
   
   'Add by Amy 2025/03/04 畫面條件 (備註:都沒勾=全部=不需過濾cu181條件)
   If chkA4228(0).Value = 1 And chkA4228(1).Value = 1 Then
      stCon = stCon & "And cu181 is not null "
   ElseIf chkA4228(0).Value = 1 Then
      stCon = stCon & "And cu181='1' "
   ElseIf chkA4228(1).Value = 1 Then
      stCon = stCon & "And cu181='2' "
   End If
   Frmacc44w1.stPreCon = stCon
   'end 2025/03/04
   Frmacc44w1.Show
   Me.Hide
End Sub

Private Sub Form_Activate()
   strFormName = Name
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single

   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5625
   Me.Height = 3468 'Modify by Amy 2025/03/04 原:3240
   '改單線固定(調整大小不用再設定)
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
      For intY = 0 To Int(ScaleHeight / sglHeight)
         PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
      Next
   Next

   '起日預設為上月1日
   strDate = CompDate(1, -1, strSrvDate(1))
   strDate = Left(strDate, 6) & "01"
   MaskEdBox1.Text = CFDate(ACDate(strDate))
   MaskEdBox1.Mask = DFormat
   '止日預設為上月底
   strDate = GetMonthStdDay(Left(strDate, 6), 1, True)
   MaskEdBox2.Text = CFDate(ACDate(strDate))
   MaskEdBox2.Mask = DFormat
   
   '上次發放日期
   If PUB_GetLastDate(Me.Name, "MaskEdBox3") <> "" Then
      MaskEdBox3.Text = ChangeTStringToTDateString(PUB_GetLastDate(Me.Name, "MaskEdBox3"))
   End If
   If PUB_GetLastDate(Me.Name, "MaskEdBox4") <> "" Then
      MaskEdBox4.Text = ChangeTStringToTDateString(PUB_GetLastDate(Me.Name, "MaskEdBox4"))
   End If
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = DFormat
   
   PUB_SetPrinter Me.Name, Combo3, strPrinter
   'Modify by Amy 2025/04/01 改代填方式不預勾,預勾「不含單筆收款扣繳合計未達2000元 (Excel條件)」-瑞婷
   'Add by Amy 2025/02/20 +代填方式,都預設勾選-瑞婷
   chkA4228(0).Value = 0
   chkA4228(1).Value = 0
   chkCon.Value = 1
   'end 2025/04/01
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc44w0 = Nothing
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   '日期檢查
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox "收款起始日期格式錯誤！", vbExclamation
      FormCheck = False
      MaskEdBox1.SetFocus
      Exit Function
   End If

   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox "收款迄止日期格式錯誤！", vbExclamation
      FormCheck = False
      MaskEdBox2.SetFocus
      Exit Function
   End If
   FormCheck = True
End Function

'Add by Amy 2025/02/20 Excel 存檔
Private Sub SaveExcel(ByVal stCmp As String, bolLast As Boolean)
   Dim strFileN As String
   
   strFileN = "智慧所"
   If stCmp = "L" Then strFileN = "法律所"
   strFileN = "客戶代填繳書明細表-" & strFileN & ACDate(ServerDate) & ServerTime
   
   With wksAnnuity
      .PageSetup.PaperSize = 9 'A4
      .PageSetup.Orientation = xlLandscape '橫印
      .PageSetup.Zoom = 100
      .PageSetup.LeftMargin = 0 '邊界
      .PageSetup.RightMargin = 0
      .PageSetup.TopMargin = xlsAnnuity.InchesToPoints(0.4)
      .PageSetup.BottomMargin = xlsAnnuity.InchesToPoints(0.4)
      .PageSetup.PrintTitleRows = "$1:$" & intTitleR '標題列
      .PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
   End With
   
   '判斷版本
   If Val(xlsAnnuity.Version) < 12 Then
      xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN & MsgText(43), FileFormat:=-4143
   Else
      xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN & ".xlsx", FileFormat:=51
   End If
    xlsAnnuity.Workbooks.Close
    If bolLast = True Then xlsAnnuity.Quit
End Sub

