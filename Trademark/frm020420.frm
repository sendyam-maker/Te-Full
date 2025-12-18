VERSION 5.00
Begin VB.Form frm020420 
   BorderStyle     =   1  '單線固定
   Caption         =   "MCT收發文件數及點數統計"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   6345
   Begin VB.TextBox Text2 
      Height          =   264
      Left            =   2085
      MaxLength       =   1
      TabIndex        =   2
      Top             =   360
      Width           =   492
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   540
      Left            =   30
      TabIndex        =   8
      Top             =   960
      Width           =   6000
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   9
         Top             =   180
         Width           =   5000
      End
      Begin VB.Label Label4 
         Caption         =   "印表機"
         Height          =   180
         Left            =   105
         TabIndex        =   10
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Index           =   0
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   0
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Index           =   1
      Left            =   2265
      MaxLength       =   5
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5430
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   4590
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   30
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   1185
      TabIndex        =   3
      Top             =   630
      Width           =   3975
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "不計件之案件是否統計：             （Ｙ：統計）"
      Height          =   180
      Left            =   60
      TabIndex        =   14
      Top             =   420
      Width           =   3645
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(ALL：全部)"
      Height          =   180
      Left            =   5190
      TabIndex        =   13
      Top             =   690
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Enabled         =   0   'False
      Height          =   180
      Index           =   3
      Left            =   75
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   1560
      X2              =   2310
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "月"
      Height          =   180
      Index           =   0
      Left            =   2805
      TabIndex        =   11
      Top             =   1050
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收發文年月："
      Height          =   180
      Left            =   75
      TabIndex        =   7
      Top             =   120
      Width           =   1080
   End
   Begin VB.Line Line5 
      X1              =   2115
      X2              =   2385
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   60
      TabIndex        =   6
      Top             =   690
      Width           =   900
   End
End
Attribute VB_Name = "frm020420"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Amy 2019/05/28
Option Explicit

Dim RsQ As New ADODB.Recordset
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim iLine As Integer, iPage As Integer
Dim intFieldH As Integer, intTxtH As Integer, intMaxField As Integer '欄位高/字高/MCTF0x顯示個數(A4橫印)
Dim intQ As Integer
Dim strQ As String, strOldN As String, strAllMCTF As String, strAllMCTFNo As String
Dim PLeft() As Integer
Dim arrMCTF() As String, strMCTF0X() As String
Dim bolLessMax As Boolean '小於A4可放欄位
'Add by Amy 2019/08/06
Dim bolSetABClass As Boolean, bolPrint As Boolean '區分AB類/列印
Dim strFieldN '欄位名稱
Dim intCounter As Integer, intTitleR As Integer, intField As Integer
Dim xlsFileName As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Function doQuery() As Boolean
    Dim strWhere(3) As String, strSys As String, strField As String
    Dim strBase(2) As String
    Dim strAllTM01 As String, strField1 As String 'Add by Amy 2019/08/06
    
On Error GoTo ErrHnd
    
    doQuery = False
    'Add by Amy 2019/08/06 取得商標基本檔 tm01
    strQ = "Select Distinct tm01 From TradeMark"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            strAllTM01 = strAllTM01 & "," & RsQ.Fields(0)
            RsQ.MoveNext
        Loop
    End If
    If strAllTM01 <> MsgText(601) Then strAllTM01 = "'" & Replace(Mid(strAllTM01, 2), ",", "','") & "'"
    'end 2019/08/06
    
    strQ = "Delete From R020420 Where ID='" & strUserNum & "' "
    cnnConnection.Execute strQ
  
    If Text1 <> MsgText(601) Then
        strSys = IIf(UCase(Text1) <> "ALL", Text1, GetAllSysKind(Text1))
    End If
    If txtDate(0) <> MsgText(601) Then
        strWhere(0) = strWhere(0) & " And cp05>=" & Val(txtDate(0) & "00") + 19110000
        strWhere(1) = strWhere(1) & " And R003>=" & Val(txtDate(0) & "00") + 19110000
        strWhere(2) = strWhere(2) & " And cp27>=" & Val(txtDate(0) & "00") + 19110000
        strWhere(3) = strWhere(3) & " And R004>=" & Val(txtDate(0) & "00") + 19110000
    End If
    If txtDate(1) <> MsgText(601) Then
        strWhere(0) = strWhere(0) & " And cp05<=" & Val(txtDate(1) & "00") + 19110031
        strWhere(1) = strWhere(1) & " And R003<=" & Val(txtDate(1) & "00") + 19110031
        strWhere(2) = strWhere(2) & " And cp27<=" & Val(txtDate(1) & "00") + 19110031
        strWhere(3) = strWhere(3) & " And R004<=" & Val(txtDate(1) & "00") + 19110031
    End If
    
    '不計件之案件是否統計
    If Len(Trim(Text2)) = 0 Then
         strWhere(0) = strWhere(0) & " And cp26 is null "
         strWhere(2) = strWhere(2) & " And cp26 is null "
    End If
    strWhere(0) = strWhere(0) & " And cp159=0 And cp09< 'C' And SubStr(cp161,1,4)='MCTF' "
    strWhere(2) = strWhere(2) & " And cp09< 'C' And SubStr(cp161,1,4)='MCTF' "
                        
   
    strField = "cp09,cp10,cp05,cp27,Nvl(cp16,0)-Nvl(cp17,0) as cp18,cp60,cp01,cp02,cp03,cp04,cp13,cp161"
    'Modify by Amy 2019/08/06 增加商品類別
    '商標
    strBase(0) = "Select " & strField & ",tm44 as Ag,tm23 as App,tm10 as Nation,counting(tm09) as tm09 From CaseProgress,TradeMark,Fagent F1,Fagent F2,Customer " & _
                    "Where cp01=tm01(+) And cp02=tm02(+) And cp03=tm03(+) And cp04=tm04(+) And tm44 is not null " & _
                    "And SubStr(tm44,1,8)=F1.FA01(+) And SubStr(tm44,9,1)=F1.FA02(+) And SubStr(cp44,1,8)=F2.FA01(+) And SubStr(cp44,9,1)=F2.FA02(+) " & _
                    "And SubStr(tm23,1,8)=cu01(+) And Decode(SubStr(tm23,9,1),Null,'0',SubStr(tm23,9,1))=cu02(+) " & _
                    "And cp01 IN (" & SQLGrpStr(strSys, 2) & ") "
    '法務
    strBase(1) = "Select " & strField & ",lc22 as Ag,lc11 as App,lc15 as Nation,0 as tm09 From CaseProgress,LawCase,Fagent F1,Fagent F2,Customer  " & _
                    "Where cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) And lc22 is not null " & _
                    "And SubStr(lc22,1,8)=F1.FA01(+) And SubStr(lc22,9,1)=F1.FA02(+) And SubStr(cp44,1,8)=F2.FA01(+) And SubStr(cp44,9,1)=F2.FA02(+) " & _
                    "And SubStr(lc11,1,8)=cu01(+) And Decode(SubStr(lc11,9,1),Null,'0',SubStr(lc11,9,1))=cu02(+) " & _
                    "And cp01 IN (" & SQLGrpStr(strSys, 3) & ") "
    '服務
    strBase(2) = "Select " & strField & ",sp26 as Ag,sp08 as App,sp09 as Nation,counting(sp73) as tm09 From CaseProgress,ServicePractice,Fagent F1,Fagent F2,Customer " & _
                    "Where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) And sp26 is not null " & _
                    "And SubStr(sp26,1,8)=F1.FA01(+) And SubStr(sp26,9,1)=F1.FA02(+) And SubStr(sp26,1,8)=F2.FA01(+) And SubStr(cp44,9,1)=F2.FA02(+) " & _
                    "And SubStr(sp08,1,8)=cu01(+) And Decode(SubStr(sp08,9,1),Null,'0',SubStr(sp08,9,1))=cu02(+) " & _
                    "And cp01 IN (" & SQLGrpStr(strSys, 5) & ") "
    
    'Modify by Amy 2019/07/23 收發文資料拆開Insert 否則發文資料若有「不續辦」且收發文同一天時,收文件數會被計算
    '收文資料
    strQ = "Select  '" & strUserNum & "',cp09,cp10,cp05,Nvl(cp18,0),Nvl(a1u07,0),cp60,cp01,cp02,cp03,cp04,cp13,Ag,App,Nation,cp161,Nvl(cpm03,cpm04)||'-'||sk02,'1',tm09 " & _
                "From (" & strBase(0) & strWhere(0) & " Union " & strBase(1) & strWhere(0) & " Union " & strBase(2) & strWhere(0) & " ),CaseProPertyMap,SystemKind, " & _
                            "(Select a1u03,Sum(a1u07) a1u07 From Caseprogress,Acc1u0 Where cp09=a1u03(+) " & strWhere(0) & " Group By a1u03)" & _
                "Where cp01=cpm01(+) And cp10=cpm02(+) And cp01=sk01(+) And cp09=a1u03(+)"
    strQ = "Insert Into R020420 (ID,R001,R002,R003,R005,R006,R007,R008,R009,R010,R011,R012,R013,R014,R015,R016,R017,R018,R019) " & strQ
    cnnConnection.Execute strQ
    
    '發文資料
    strQ = "Select  '" & strUserNum & "',cp09,cp10,cp27,Nvl(cp18,0),Nvl(a1u07,0),cp60,cp01,cp02,cp03,cp04,cp13,Ag,App,Nation,cp161,Nvl(cpm03,cpm04)||'-'||sk02,'1',tm09 " & _
                "From (" & strBase(0) & strWhere(2) & " Union " & strBase(1) & strWhere(2) & " Union " & strBase(2) & strWhere(2) & " ),CaseProPertyMap,SystemKind, " & _
                            "(Select a1u03,Sum(a1u07) a1u07 From Caseprogress,Acc1u0 Where cp09=a1u03(+) " & strWhere(2) & " Group By a1u03)" & _
                "Where cp01=cpm01(+) And cp10=cpm02(+) And cp01=sk01(+) And cp09=a1u03(+)"
    strQ = "Insert Into R020420 (ID,R001,R002,R004,R005,R006,R007,R008,R009,R010,R011,R012,R013,R014,R015,R016,R017,R018,R019) " & strQ
    cnnConnection.Execute strQ
    'end 2019/07/23
    'end 2019/08/06
    
    '更新顯示案件性質名稱+sk02欄位說明第一個字
    strQ = "Update R020420 set R017=Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(r017,'-2',''),'-6',''),'-1','-專'),'-5','-專'),'-3','-法'),'-7','-法'),'-4','-顧'),'-8','-顧') " & _
                "Where ID='" & strUserNum & "' And R018='1' "
    cnnConnection.Execute strQ
    
    'Add by Amy 2019/08/06 類別數計算,商標-計算案件性質為101及001,服務業務-計算案件性質為001
    strQ = "Update R020420 set R019=0 Where ID='" & strUserNum & "' And R018='1' " & _
                "And ( (R002 Not In ('101','001') And R008 In (" & strAllTM01 & ")) Or (R002 <>'001' And R008 Not In (" & strAllTM01 & ")) ) "
    cnnConnection.Execute strQ

    'Modify by Amy 2019/08/06 原區分A、B類顯示,改為不分
    strField = "": strField1 = ""
    If bolSetABClass = True Then
        strField = ",R001"
        strField1 = ",SubStr(R001,1,1)"
    End If
    '收文 件數/點數(Union 新增一筆類別數資料)
    strQ = "Insert Into R020420 (ID,R018,R016" & strField & ",R002,R005,R006,R017,R003) " & _
                "Select '" & strUserNum & "','2',R016" & strField1 & ",R002,Count(R003) CNT,Round(Sum(Nvl(r005,0)-Nvl(r006,0))/1000,3) Point,R017,'111111' " & _
                "From R020420 Where ID='" & strUserNum & "' And R018='1' " & strWhere(1) & " " & _
                "Group by R002,R016,R017 " & strField1 & _
    "Union Select '" & strUserNum & "','2',R016" & strField1 & ",R002||'ZZ',Sum(Nvl(R019,0)) CNT,Round(Sum(Nvl(r005,0)-Nvl(r006,0))/1000,3) Point,R017||'(類)','111111' " & _
                "From R020420 Where ID='" & strUserNum & "' And R018='1' And R002 In ('101','001') " & strWhere(1) & " " & _
                "Group by R002,R016,R017 " & strField1
    cnnConnection.Execute strQ
    
    '發文 件數/點數(Union 新增一筆類別數資料)
    strQ = "Insert Into R020420 (ID,R018,R016" & strField & ",R002,R005,R006,R017,R004) " & _
                "Select '" & strUserNum & "','2',R016" & strField1 & ",R002,Count(R004) CNT,Round(Sum(Nvl(r005,0)-Nvl(r006,0))/1000,3) Point,R017,'111111' " & _
                "From R020420 Where ID='" & strUserNum & "' And R018='1' " & strWhere(3) & " " & _
                "Group by R002,R016,R017 " & strField1 & _
    "Union Select '" & strUserNum & "','2',R016" & strField1 & ",R002||'ZZ',Sum(Nvl(R019,0)) CNT,Round(Sum(Nvl(r005,0)-Nvl(r006,0))/1000,3) Point,R017||'(類)','111111' " & _
                "From R020420 Where ID='" & strUserNum & "' And R018='1' And R002 In ('101','001') " & strWhere(3) & " " & _
                "Group by R002,R016,R017 " & strField1
    cnnConnection.Execute strQ
        
    '收文合計MCTFZZ(橫向)
    strQ = "Insert Into R020420 (ID,R018,R016" & strField & ",R002,R005,R006,R017,R003) " & _
                "Select '" & strUserNum & "','2','MCTFZZ'" & strField1 & ",R002,Count(R003) CNT,Round(Sum(Nvl(r005,0)-Nvl(r006,0))/1000,3) Point,R017,'111111' " & _
                "From R020420 Where ID='" & strUserNum & "' And R018='1' " & strWhere(1) & " " & _
                "Group by R002,R017 " & strField1 & _
    "Union Select '" & strUserNum & "','2','MCTFZZ'" & strField1 & ",R002,Sum(Nvl(R005,0)) CNT,Sum(Nvl(r006,0)) Point,R017,'111111' " & _
                "From R020420 Where ID='" & strUserNum & "' And R018='2' And R002 In ('101ZZ','001ZZ') And R004 is null " & _
                "Group by R002,R017 " & strField1
    cnnConnection.Execute strQ
    
    '發文合計MCTFZZ(橫向)
    strQ = "Insert Into R020420 (ID,R018,R016" & strField & ",R002,R005,R006,R017,R004) " & _
                "Select '" & strUserNum & "','2','MCTFZZ'" & strField1 & ",R002,Count(R004) CNT,Round(Sum(Nvl(r005,0)-Nvl(r006,0))/1000,3) Point,R017,'111111' " & _
                "From R020420 Where ID='" & strUserNum & "' And R018='1' " & strWhere(3) & " " & _
                "Group by R002,R017 " & strField1 & _
    "Union Select '" & strUserNum & "','2','MCTFZZ'" & strField1 & ",R002,Sum(Nvl(R005,0)) CNT,Sum(Nvl(r006,0)) Point,R017,'111111' " & _
                "From R020420 Where ID='" & strUserNum & "' And R018='2' And R002 In ('101ZZ','001ZZ') And R003 is null " & _
                "Group by R002,R017 " & strField1
    cnnConnection.Execute strQ
    
    'AB類小計(以類別數加總,點數也只計算類別數那筆,才不會重覆計算)
    If bolSetABClass = True Then
        '收文
        strQ = "Insert Into R020420 (ID,R018,R016,R001,R002,R005,R006,R017,R003) " & _
                    "Select '" & strUserNum & "','2',R016,R001||'Z','Z',Sum(R005),Sum(R006),'小計','111111' From R020420 " & _
                    "Where ID='" & strUserNum & "' And R018='2' And R003='111111' And R002 Not In ('101','001') Group by r001,r016 "
        cnnConnection.Execute strQ
    
        '發文
        strQ = "Insert Into R020420 (ID,R018,R016,R001,R002,R005,R006,R017,R004) " & _
                    "Select '" & strUserNum & "','2',R016,R001||'Z','Z',Sum(R005),Sum(R006),'小計','111111' From R020420 " & _
                    "Where ID='" & strUserNum & "' And R018='2' And R004='111111' And R002 Not In ('101','001') Group by r001,r016 "
        cnnConnection.Execute strQ
        
        strField = ",R001": strField1 = ",'ZZ'"
    End If
   
    '總計-收文(以類別數加總,點數也只計算類別數那筆,才不會重覆計算；AB類需排除小計)
    strQ = "Insert Into R020420 (ID,R018,R016" & strField & ",R002,R005,R006,R017,R003) " & _
                "Select '" & strUserNum & "','2',R016" & strField1 & ",'ZZ',Sum(R005),Sum(R006),'總計','111111' From R020420 " & _
                "Where ID='" & strUserNum & "' And R018='2' And R003='111111' And R002 Not In ('101','001') And R002<>'Z' " & _
                "Group by R016 "
    cnnConnection.Execute strQ
    
    '總計-發文(以類別數加總,點數也只計算類別數那筆,才不會重覆計算；AB類需排除小計)
    strQ = "Insert Into R020420 (ID,R018,R016" & strField & ",R002,R005,R006,R017,R004) " & _
                "Select '" & strUserNum & "','2',R016" & strField1 & ",'ZZ',Sum(R005),Sum(R006),'總計','111111' From R020420 " & _
                "Where ID='" & strUserNum & "' And R018='2' And R004='111111' And R002 Not In ('101','001') And R002<>'Z' " & _
                "Group by R016 "
    cnnConnection.Execute strQ
    'end 2019/08/06
    
    doQuery = True
    Exit Function

ErrHnd:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "錯誤 : " & Err.Description, vbCritical
    End If
End Function

Private Sub PrintA4(ByVal intStart As Integer, ByVal intEnd As Integer)
    'Memo by Amy 因107年一整年點數 9萬多點,故數值部份 5位整數+2位小數,MCTF0X只能顯示4筆(不含合計)
    iLine = 1:  strOldN = ""
    Do While RsQ.EOF = False
        If iLine > 35 Or iLine = 1 Then
            If iPage <> 0 Then Printer.NewPage
            iLine = 1
            Call PrintTitle(intStart, intEnd) '列印表頭
        End If
        Call PrintDetail(intStart, intEnd)
        strOldN = "" & RsQ.Fields("cp10N")
        RsQ.MoveNext
    Loop
    'Line-橫(最後)
    iLine = iLine - 1
    Printer.CurrentX = PLeft(LBound(PLeft))
    Printer.CurrentY = iLine * intFieldH - 10
    Printer.Line (PLeft(LBound(PLeft)), iLine * intFieldH - 10)-(PLeft(UBound(PLeft)), iLine * intFieldH - 10)
    
End Sub

'顯示下一區類別
Private Sub PrintClass(ByVal stClass As String)
    Printer.Font.Size = 11
    'Line-直線(最左)
    Printer.CurrentX = 0
    Printer.CurrentY = iLine * intFieldH
    Printer.Line (0, iLine * intFieldH)-(0, (iLine + 1) * intFieldH)
    '資料
    Printer.CurrentX = 10 + PLeft(LBound(PLeft)) + (PLeft(LBound(PLeft) + 1) - PLeft(LBound(PLeft))) / 2 - (Printer.TextWidth(stClass) / 2)
    Printer.CurrentY = iLine * intFieldH
    Printer.Print stClass
    'Line-直線(最右)
    Printer.CurrentX = PLeft(UBound(PLeft))
    Printer.CurrentY = iLine * intFieldH
    Printer.Line (PLeft(UBound(PLeft)), iLine * intFieldH)-(PLeft(UBound(PLeft)), (iLine + 1) * intFieldH)
    'Line-橫(件數/點數 下方)
    Printer.CurrentX = PLeft(LBound(PLeft))
    Printer.CurrentY = iLine * intFieldH
    Printer.Line (PLeft(LBound(PLeft)), (iLine + 1) * intFieldH)-(PLeft(UBound(PLeft)), (iLine + 1) * intFieldH)
    iLine = iLine + 1
End Sub

Private Sub PrintTitle(ByVal intStart As Integer, ByVal intEnd As Integer)
    Dim intCount1 As Integer, intCount2 As Integer
    Dim stTmp As String, stTmp2 As String
    Dim ii As Integer, intX As Integer
    
    intFieldH = 300: intTxtH = 305
    iPage = iPage + 1
    Printer.Font.Size = 18
    Printer.Font.Underline = False
    Printer.FontBold = False
    
    stTmp = "MCT 收發文之件數及點數統計"
    Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(stTmp) / 2)
    Printer.CurrentY = iLine * intFieldH
    Printer.Print stTmp
    iLine = iLine + 2
    
    Printer.Font.Size = 12
    stTmp2 = Val(txtDate(0)) + 191100
    stTmp = Val(Mid(stTmp2, 1, 4)) - 1911 & "年" & Val(Mid(stTmp2, 5, 2)) & "月"
    stTmp2 = Val(txtDate(1)) + 191100
    stTmp = stTmp & "~" & Val(Mid(stTmp2, 1, 4)) - 1911 & "年" & Val(Mid(stTmp2, 5, 2)) & "月"
    stTmp = "日期：" & stTmp
    Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(stTmp) / 2)
    Printer.CurrentY = iLine * intFieldH
    Printer.Print stTmp
    iLine = iLine + 1
    
    stTmp = "列印人員：" & strUserName
    Printer.CurrentX = PLeft(0)
    Printer.CurrentY = iLine * intFieldH
    Printer.Print stTmp
 
    stTmp = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(stTmp) - 500
    Printer.CurrentY = iLine * intFieldH
    Printer.Print stTmp
    iLine = iLine + 2
    
    '*** 表格-抬頭 ***
    'Line-橫(最上方)
    Printer.CurrentX = PLeft(LBound(PLeft))
    Printer.CurrentY = iLine * intFieldH
    Printer.Line (PLeft(LBound(PLeft)), iLine * intFieldH)-(PLeft(UBound(PLeft)), iLine * intFieldH)
    '直線(最左)
    Printer.CurrentX = 0
    Printer.CurrentY = iLine * intFieldH
    Printer.Line (0, iLine * intFieldH)-(0, (iLine + 3) * intFieldH)
  
    'MCTF0X
    For ii = intStart To intEnd
        stTmp = arrMCTF(ii)
        stTmp = Replace(Replace(stTmp, "'", ""), "MCTFZZ", "合　計")
        
        intX = ii '起始位置
        If intStart > intMaxField Then
            intX = ii Mod intMaxField
        End If
        If ii > 1 Then
        '(件數+點數)*2欄 =4
            intX = 4 * (intX - 1) + 1
        End If
        '                   起位置     +(    迄位置          -  起位置      )/2  - 字寬
        intX = 10 + PLeft(intX) + (PLeft(intX + 4) - PLeft(intX)) / 2 - (Printer.TextWidth(stTmp) / 2)
        
        Printer.CurrentX = intX
        Printer.CurrentY = iLine * intTxtH
        Printer.Print stTmp
    Next ii
     iLine = iLine + 1
    
    'Line-橫(MCTF0 下方)
    Printer.CurrentX = PLeft(LBound(PLeft))
    Printer.CurrentY = iLine * intFieldH
    Printer.Line (PLeft(LBound(PLeft) + 1), iLine * intFieldH)-(PLeft(UBound(PLeft)), iLine * intFieldH)
    
    '收文/發文/件數/點數
    intCount1 = 1: intCount2 = 1
    For ii = LBound(PLeft) To UBound(PLeft)
        '直線 (最左)
        If ii = 0 Then
            Printer.CurrentX = PLeft(ii)
            Printer.CurrentY = iLine * intFieldH
            Printer.Line (PLeft(ii), iLine * intFieldH)-(PLeft(ii), iLine * intFieldH)
        '直線 i=1,5,9...(畫2列)
        ElseIf ii Mod 4 = 1 Then
            Printer.CurrentX = PLeft(ii)
            Printer.CurrentY = (iLine - 1) * intFieldH
            Printer.Line (PLeft(ii), (iLine - 1) * intFieldH)-(PLeft(ii), iLine * intFieldH)
        End If
        '收文/發文
        If ii Mod 2 = 1 Then
            If ii < UBound(PLeft) Then
                stTmp = "收文"
                If intCount1 Mod 2 = 0 Then stTmp = "發文"
                                             '        起位置         +(   迄位置            -    起位置       )/2 -  字寬
                Printer.CurrentX = 10 + Val(PLeft(ii)) + (Val(PLeft(ii + 2)) - Val(PLeft(ii))) / 2 - (Printer.TextWidth(stTmp) / 2)
                Printer.CurrentY = iLine * intTxtH
                Printer.Print stTmp
                intCount1 = intCount1 + 1
            End If
            '直線 i=1,3,5...
            Printer.CurrentX = PLeft(ii)
            Printer.CurrentY = iLine * intFieldH
            Printer.Line (PLeft(ii), iLine * intFieldH)-(PLeft(ii), (iLine + 1) * intFieldH)
        End If
        If ii > 0 And ii < UBound(PLeft) Then
            stTmp = "件數"
            If intCount2 Mod 2 = 0 Then stTmp = "點數"
                                             '        起位置     +(   迄位置            -    起位置       )/2 -  字寬
            Printer.CurrentX = 10 + Val(PLeft(ii)) + (Val(PLeft(ii + 1)) - Val(PLeft(ii))) / 2 - (Printer.TextWidth(stTmp) / 2)
            Printer.CurrentY = (iLine + 1) * intTxtH
            Printer.Print stTmp
            intCount2 = intCount2 + 1
            '直線 (件數點數中間)
            Printer.CurrentX = PLeft(ii)
            Printer.CurrentY = (iLine + 1) * intFieldH
            Printer.Line (PLeft(ii), (iLine + 1) * intFieldH)-(PLeft(ii), (iLine + 2) * intFieldH)
        End If
        
        If ii = UBound(PLeft) Then
            'Line-橫(收文/發文 下方)
            Printer.CurrentX = PLeft(LBound(PLeft))
            Printer.CurrentY = (iLine + 1) * intFieldH
            Printer.Line (PLeft(LBound(PLeft) + 1), (iLine + 1) * intFieldH)-(PLeft(ii), (iLine + 1) * intFieldH)
            'Line-橫(件數/點數 下方)
            Printer.CurrentX = PLeft(LBound(PLeft))
            Printer.CurrentY = (iLine + 2) * intFieldH
            Printer.Line (PLeft(LBound(PLeft)), (iLine + 2) * intFieldH)-(PLeft(ii), (iLine + 2) * intFieldH)
            '直線 (件數點數 最右)
            Printer.CurrentX = PLeft(ii)
            Printer.CurrentY = (iLine + 1) * intFieldH
            Printer.Line (PLeft(ii), (iLine + 1) * intFieldH)-(PLeft(ii), (iLine + 2) * intFieldH)
        End If
    Next ii
    iLine = iLine + 2
    '*** End 表格-抬頭 ***
End Sub

Private Sub PrintDetail(ByVal intStart As Integer, ByVal intEnd As Integer)
    Dim intX As Integer, ii As Integer
    Dim stTmp As String
    
    Printer.Font.Size = 11
    If (strOldN = MsgText(601) Or strOldN = "小計") Then
        '顯示下一區類別
        If "" & RsQ.Fields("cp10N") = "總計" Then
            Call PrintClass("")
        'Modfiy by Amy 2019/08/06 + bolSetABClass = True,若仍需區分AB類用,目前不使用
        ElseIf bolSetABClass = True Then
            Call PrintClass(RsQ.Fields("cp09") & " 　類　")
        End If
    End If
    
    If iLine > 35 Then
        Printer.NewPage
        iLine = 1
        Call PrintTitle(intStart, intEnd) '列印表頭
    End If
  
    For ii = LBound(PLeft) To UBound(PLeft)
        Printer.CurrentX = PLeft(ii)
        Printer.CurrentY = iLine * intFieldH
        Printer.Line (PLeft(ii), iLine * intFieldH)-(PLeft(ii), (iLine + 1) * intFieldH)
        If ii <> UBound(PLeft) Then
            stTmp = "" & RsQ.Fields(ii)
            '案件性質名稱
            If UCase(RsQ.Fields(ii).Name) = "CP10N" Then
                stTmp = IIf(PLeft(1) = 1500, PUB_StrToStr_byVal(stTmp, 18), stTmp)
                If stTmp = "小計" Then
                    intX = PLeft(ii) + (PLeft(ii + 1) - PLeft(ii)) / 2 - Printer.TextWidth(stTmp)
                Else
                    intX = PLeft(ii) + 20
                End If
            'MCTGX數值
            Else
                intX = PLeft(ii + 1) - Printer.TextWidth(stTmp) - 20
            End If
        End If
        If ii <> UBound(PLeft) Then
            Printer.CurrentX = intX
            Printer.CurrentY = iLine * intFieldH
            Printer.Print stTmp
        End If
    Next
    'Line-橫(資料 下方)
    Printer.CurrentX = PLeft(LBound(PLeft))
    Printer.CurrentY = iLine * intFieldH
    Printer.Line (PLeft(LBound(PLeft)), (iLine + 1) * intFieldH)-(PLeft(UBound(PLeft)), (iLine + 1) * intFieldH)
    iLine = iLine + 1
End Sub

Private Sub GetPleft()
    Dim ii As Integer, intWidth As Integer
    
    For ii = LBound(PLeft) To UBound(PLeft)
        If ii = 0 Then
            PLeft(ii) = 0
        ElseIf ii = 1 Then
            PLeft(ii) = 1500
            If intMaxField < 6 Then PLeft(ii) = 2200
            intWidth = (16630 - PLeft(1) - PLeft(0)) \ (intMaxField) * 4
        Else
            PLeft(ii) = PLeft(ii - 1) + intWidth
        End If
    Next ii
End Sub

Private Sub GetMCTF0X()
    'Add by Amy 2019/08/06
    Dim ii As Integer, j As Integer, stTmp As String
    
    strQ = "Select Distinct R016 From R020420 Where ID='" & strUserNum & "' And R018='2' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            strAllMCTF = strAllMCTF & ",'" & RsQ.Fields("R016") & "'"
            RsQ.MoveNext
        Loop
    End If
    'Add by Amy 2019/08/06
    If strAllMCTF <> MsgText(601) Then
        arrMCTF = Split(strAllMCTF, ",")
        '設定Excel  欄位名稱
        If bolPrint = False Then
            ReDim strFieldN(UBound(arrMCTF) * 4)
            
            For ii = LBound(arrMCTF) To UBound(arrMCTF)
                If ii = LBound(arrMCTF) Then
                    strFieldN(ii) = "案件性質"
                Else
                    For j = 1 To 4
                        stTmp = "M" & Right(Replace(arrMCTF(ii), "'", ""), 2)
                        If j Mod 4 = 1 Or j Mod 4 = 2 Then
                            stTmp = stTmp & "收文"
                        Else
                            stTmp = stTmp & "發文"
                        End If
                        If j Mod 2 = 0 Then
                            stTmp = stTmp & "點數"
                        Else
                            stTmp = stTmp & "件數"
                        End If
                        strFieldN(4 * (ii - 1) + j) = stTmp
                    Next j
                End If
            Next ii
        End If
    End If
    'end 2019/08/06
End Sub

Private Sub PrintData()
    Dim strQAll As String, strField As String, strQWhere As String
    Dim i As Integer, j As Integer, intEnd As Integer
    Dim strTmp As String 'Add by Amy 2019/08/06
    
    'Modfy by Amy 2019/08/06  加產生Excel
    If bolPrint = True Then
        If intMaxField > UBound(arrMCTF) Then
            intMaxField = UBound(arrMCTF)
            bolLessMax = True
        End If
    Else
        intMaxField = UBound(arrMCTF)
        intEnd = UBound(arrMCTF)
    End If
    'end 2019/08/06
   
    For j = LBound(arrMCTF) + 1 To UBound(arrMCTF) Step intMaxField
        strField = "": strQAll = "": strQWhere = ""
        'Modify by Amy 2019/08/06 加產生Excel,故印紙本才設定
        If bolPrint = True Then
            intEnd = j + intMaxField - 1
            If bolLessMax = True Then
                '小於A4可顯示欄位
                If j = LBound(arrMCTF) + 1 Then
                    ReDim PLeft(1 + (intMaxField) * 4)
                    GetPleft
                     intEnd = intMaxField
                End If
            '欄位大於A4 顯示
            ElseIf intEnd > UBound(arrMCTF) Then
                 intEnd = UBound(arrMCTF)
            End If
            
            If j = LBound(arrMCTF) + 1 Then
                iPage = 0
                Printer.Orientation = 2 '2.橫印
            End If
        End If
        'end 2019/08/06
       
        For i = j To intEnd
            'Modify by Amy 2019/08/06 加產生Excel,故印紙本才設定
            If bolPrint = True Then
                '第一頁及最後一頁判斷是否重設邊界
                'Modify by Amy 2019/07/23 原:(j = i Or j = UBound(arrMCTF) * intMaxField)
                If bolLessMax = False And j = i Then
                    ReDim PLeft(1 + (intEnd - i + 1) * 4)
                    GetPleft
                End If
            End If
            
            strTmp = Replace(Replace(arrMCTF(i), "'", ""), "MCTF", "")

            strField = strField & ",Round(Nvl(M" & strTmp & "a.R005,0),1) M" & strTmp & "a1,Round(Nvl(M" & strTmp & "a.R006,0),1) M" & strTmp & "a2" & _
                                            ",Round(Nvl(M" & strTmp & "b.R005,0),1) M" & strTmp & "b1,Round(Nvl(M" & strTmp & "b.R006,0),1) M" & strTmp & "b2"
            strQAll = strQAll & "," & _
                    "(Select " & IIf(bolSetABClass = True, "R001,", "") & "R002,R005,R006,R017 From R020420 " & _
                    "Where ID='" & strUserNum & "' And R018='2' And R016=" & arrMCTF(i) & " And r003='111111' ) M" & strTmp & "a"
            strQAll = strQAll & "," & _
                        "(Select " & IIf(bolSetABClass = True, "R001,", "") & "R002,R005,R006,R017 From R020420 " & _
                        "Where ID='" & strUserNum & "' And R018='2' And R016=" & arrMCTF(i) & " And r004='111111' ) M" & strTmp & "b"
            strQWhere = strQWhere & IIf(bolSetABClass = True, "And cp09=M" & strTmp & "a.r001(+)", "") & "And cp10=M" & strTmp & "a.r002(+) And cp10N=M" & strTmp & "a.r017(+) " & _
                                                       IIf(bolSetABClass = True, "And cp09=M" & strTmp & "b.r001(+)", "") & "And cp10=M" & strTmp & "b.r002(+) And cp10N=M" & strTmp & "b.r017(+) "
            'end 2019/08/06
        Next i
        '資料顯示
        'Modify by Amy 2019/08/06 加產生Excel
        strQ = "Select cp10N" & strField & IIf(bolSetABClass = True, ",cp09", "") & ",cp10 From " & _
                   "(Select Distinct " & IIf(bolSetABClass = True, "R001 cp09,", "") & "R002 cp10,R017 cp10N " & _
                    "From R020420 Where ID='" & strUserNum & "' And R018='2' ) " & _
                    strQAll & " Where " & Mid(strQWhere, 5) & " Order by " & IIf(bolSetABClass = True, "cp09,", "") & "cp10 "
        intQ = 1
        Set RsQ = ClsLawReadRstMsg(intQ, strQ)
        If intQ = 1 Then
            'Modify by Amy 2019/08/06 加產生Excel
            If bolPrint = True Then
                Call PrintA4(j, intEnd)
                iPage = iPage + 1
            'Excel
            ElseIf ExcelSave = True Then
                'Modify by Amy 2021/06/21 原:strExcelPath 改中文字顯示
                MsgBox "檔案已產生！" & vbCrLf & _
                "檔案存於 " & strExcelPathN & " " & xlsFileName
            End If
        End If
    Next j
    'Modify by Amy 2019/08/06 加產生Excel,故印紙本才設定
    If bolPrint = True Then
        Printer.EndDoc
        ShowPrintOk
    End If
    'end 2019/08/06
End Sub

Private Sub cmdOK_Click()
    Dim strMsg As String
   
    Screen.MousePointer = vbHourglass
    If FormCheck = True Then
        bolLessMax = False
        If doQuery = True Then
            ''Modify by Amy 2019/08/06 加產生Excel,故印紙本才設定
            If bolPrint = True Then
                '取數字長度
                strQ = "Select Max(Length(R005)) From (" & _
                             "Select Round(R005,2) R005 From R020420 Where ID='" & strUserNum & "' And R018='2' " & _
                "Union Select Round(R006,2) R005 From R020420  Where ID='" & strUserNum & "' And R018='2' )"
                intQ = 1
                Set RsQ = ClsLawReadRstMsg(intQ, strQ)
                If intQ = 1 Then intMaxField = Val("" & RsQ.Fields(0))
                
                If intMaxField = 0 Then
                    '判斷是否無資料
                    strQ = "Select * From R020420 Where ID='" & strUserNum & "' And R018='1' "
                    intQ = 1
                    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
                    If intQ = 0 Then
                        strMsg = "無資料可列印"
                    Else
                        strMsg = "取數字長度有誤，請洽電腦中心"
                    End If
                    MsgBox strMsg, vbExclamation
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Else
                    '數字長度(含小數點)=5,A4橫向可印 5個MCTF欄加一欄「合計」(共6個),遞減
                    'ex:數字長度 6,可印5個MCTF+1個合計
                    intMaxField = 6 + (6 - intMaxField)
                End If
            End If
            'end 2019/08/06
            
            '組MCTF字串
            strAllMCTF = ""
            Call GetMCTF0X
            '列印資料
            Call PrintData
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dim strTp As String
    
    MoveFormToCenter Me
    SeekPrintL = Printer.Orientation
    PUB_SetPrinter Me.Name, Combo1, , , SeekPrint
       
    Text1 = "ALL"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
    Set Printer = Printers(SeekPrint)
    Printer.Orientation = SeekPrintL
    Set frm020420 = Nothing
End Sub

Private Function FormCheck() As Boolean
    Dim bolCancel As Boolean
    
    If txtDate(0) = "" Then
        MsgBox "請輸入補充資料日期(起)！", vbExclamation
        txtDate(0).SetFocus
        txtDate_GotFocus (0)
        Exit Function
    End If
    If txtDate(1) = "" Then
        MsgBox "請輸入補充資料日期(迄)！", vbExclamation
        txtDate(1).SetFocus
        txtDate_GotFocus (1)
        Exit Function
    End If
    If txtDate(0) <> "" And txtDate(1) <> "" Then
        Call txtDate_Validate(0, bolCancel)
        If bolCancel = True Then
            Exit Function
        End If
      
        Call txtDate_Validate(1, bolCancel)
        If bolCancel = True Then
            Exit Function
        End If
    End If
    
    FormCheck = True
End Function

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 89 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
    TextInverse txtDate(Index)
    CloseIme
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
    If txtDate(Index) = MsgText(601) Then Exit Sub
    
    If ChkDate(txtDate(Index) & "01") = False Then
        txtDate(Index).SetFocus
        txtDate_GotFocus Index
        Cancel = True
        Exit Sub
    End If
    
    If Index = 1 Then
        If RunNick2(txtDate(0), txtDate(1)) = True Then
            txtDate(Index).SetFocus
            txtDate_GotFocus Index
            Cancel = True
            Exit Sub
         End If
    End If
End Sub

'Add by Amy 2019/08/06
Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = 1 To UBound(strFieldN)
       If UCase(strFieldN(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

Private Function ExcelSave() As Boolean
    Dim xlsSalesPoint As New Excel.Application
    Dim Wks As New Worksheet
    Dim strNotSum As String, strSum(1 To 4) As String
    Dim i As Integer, intSum As Integer
    Dim stTP As String, stF As String
    
    ExcelSave = False
    
    xlsFileName = Me.Caption & ACDate(ServerDate) & ServerTime & MsgText(43)
    If Dir(strExcelPath & xlsFileName) = MsgText(601) Then
       If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
          MkDir strExcelPath
       End If
    Else
       Kill strExcelPath & xlsFileName
    End If
    
    intField = 65: intCounter = 1: intTitleR = 1
    xlsSalesPoint.SheetsInNewWorkbook = 3 '改設定(選項->一般->包括的工作表份數)
    xlsSalesPoint.Workbooks.add
    Set Wks = xlsSalesPoint.Worksheets(1)
    xlsSalesPoint.Visible = False
    Wks.PageSetup.PaperSize = 9 'A4
    Wks.PageSetup.Orientation = xlLandscape '橫印
    Wks.PageSetup.LeftMargin = xlsSalesPoint.InchesToPoints(0.5)
    Wks.PageSetup.RightMargin = xlsSalesPoint.InchesToPoints(0.5)
    Wks.PageSetup.TopMargin = xlsSalesPoint.InchesToPoints(0.2)
    Wks.PageSetup.BottomMargin = xlsSalesPoint.InchesToPoints(0.2)
            
    Call SetTitle(Wks)
    
    'Modify by Amy 2021/12/15 因MCTF編號增加,造成欄位抓錯(因Excel欄位只到Z)
    Do While RsQ.EOF = False
        Wks.Range(GetFieldStr(LBound(strFieldN), intField) & intCounter & ":" & GetFieldStr(UBound(strFieldN), intField) & intCounter).Font.Size = 12
        '查名及申請只算(類)類別的部分
        If "" & RsQ.Fields("CP10") = "001" Or "" & RsQ.Fields("CP10") = "101" Then
            strNotSum = strNotSum & "," & intCounter
        End If
        stTP = ""
        For i = LBound(strFieldN) To UBound(strFieldN)
            If i = GetValue("案件性質") Then
                stTP = "" & RsQ.Fields("CP10N")
            '總計
            ElseIf "" & RsQ.Fields("CP10") = "ZZ" Then
                stTP = "=Sum(" & GetFieldStr(i, intField) & intTitleR + 1 & ":" & GetFieldStr(i, intField) & intCounter - 1 & ")"
                If strNotSum <> MsgText(601) Then
                    stTP = stTP & "-Sum(" & Mid(Replace(strNotSum, ",", "," & GetFieldStr(i, intField)), 2) & ")"
                End If
            'MCTFZZ(橫向合計)
            ElseIf Left(strFieldN(i), 3) = "MZZ" Then
                '收文
                If InStr(strFieldN(i), "收文") > 0 Then
                    stTP = strSum(1)
                    If Right(strFieldN(i), 2) = "點數" Then stTP = strSum(2)
                '發文
                Else
                    stTP = strSum(3)
                    If Right(strFieldN(i), 2) = "點數" Then stTP = strSum(4)
                End If
                stTP = "=Sum(" & Mid(stTP, 2) & ")"
            '資料
            ElseIf Left(RsQ.Fields(i).Name, 3) = Left(strFieldN(i), 3) Then
                stF = Left(strFieldN(i), 3)
                If Mid(strFieldN(i), 4, 2) = "收文" Then
                    stF = stF & "A"
                Else
                    stF = stF & "B"
                End If
                If Right(strFieldN(i), 2) = "點數" Then
                    stF = stF & "2"
                Else
                    stF = stF & "1"
                End If
                stTP = Val("" & RsQ.Fields(stF))
                 If Mid(strFieldN(i), 4, 2) = "收文" Then
                    If Right(strFieldN(i), 2) = "點數" Then
                        strSum(2) = strSum(2) & "," & GetFieldStr(i, intField) & intCounter
                    Else
                        strSum(1) = strSum(1) & "," & GetFieldStr(i, intField) & intCounter
                    End If
                '發文
                Else
                    If Right(strFieldN(i), 2) = "點數" Then
                        strSum(4) = strSum(4) & "," & GetFieldStr(i, intField) & intCounter
                    Else
                        strSum(3) = strSum(3) & "," & GetFieldStr(i, intField) & intCounter
                    End If
                End If
            End If
            Wks.Range(GetFieldStr(i, intField) & intCounter).Value = stTP
            If i <> GetValue("案件性質") Then
                stTP = "#,##0_ "
                If Right(strFieldN(i), 2) = "點數" Then
                    Wks.Range(GetFieldStr(i, intField) & intCounter).NumberFormatLocal = "#,##0.0_ "
                End If
            End If
            If Left(strFieldN(i), 3) = "MZZ" And i = UBound(strFieldN) Then
                strSum(1) = "": strSum(2) = "": strSum(3) = "": strSum(4) = ""
            End If
        Next i
        intCounter = intCounter + 1
        RsQ.MoveNext
    Loop
    
    Call SetTitle(Wks, True)
    '畫框
    Wks.Range(GetFieldStr(LBound(strFieldN), intField) & intTitleR - 2 & ":" & GetFieldStr(UBound(strFieldN), intField) & intCounter - 1).Select
    xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    '合計
    Wks.Range(GetFieldStr(LBound(strFieldN), intField) & intCounter - 1 & ":" & GetFieldStr(UBound(strFieldN), intField) & intCounter - 1).Select
    'end 2021/12/15
    xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlDouble
    xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlDouble
    '設定-表頭保留
    Wks.PageSetup.PrintTitleRows = "$1:$" & intTitleR
    Wks.PageSetup.PrintTitleColumns = "$A:$A"
            
    '判斷若版本2007以上改變存格式
    If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsSalesPoint.Workbooks.Close
    xlsSalesPoint.Quit
    ExcelSave = True
    Exit Function
    
ErrHand:
    MsgBox Err.Description, , MsgText(5)
    If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsSalesPoint.Workbooks.Close
    xlsSalesPoint.Quit
    Set xlsSalesPoint = Nothing
End Function

Private Sub SetTitle(ByRef Wks As Worksheet, Optional ByVal bolLast As Boolean = False)
    Dim i As Integer, intVal As Integer, stTmp As String
    
    'Modify by Amy 2021/12/15 因MCTF編號增加,造成欄位抓錯(因Excel欄位只到Z)
    With Wks
        If bolLast = False Then
            .Range(GetFieldStr(LBound(strFieldN), intField) & intCounter).Font.Size = 18
            .Range(GetFieldStr(LBound(strFieldN), intField) & intCounter).Value = Me.Caption
            .Range(GetFieldStr(LBound(strFieldN), intField) & intCounter & ":" & GetFieldStr(UBound(strFieldN), intField) & intCounter).HorizontalAlignment = xlCenter
            .Range(GetFieldStr(LBound(strFieldN), intField) & intCounter & ":" & GetFieldStr(UBound(strFieldN), intField) & intCounter).MergeCells = True
            intCounter = intCounter + 1
            .Range(GetFieldStr(LBound(strFieldN), intField) & intCounter).Font.Size = 12
            .Range(GetFieldStr(LBound(strFieldN), intField) & intCounter).Value = Label2 & txtDate(0) & "~" & txtDate(1)
            .Range(GetFieldStr(LBound(strFieldN), intField) & intCounter & ":" & Chr(intField + 4) & intCounter).MergeCells = True
            intCounter = intCounter + 1
            .Range(GetFieldStr(LBound(strFieldN), intField) & intCounter).Value = "列印人員:" & StaffQuery(strUserNum)
            .Range(GetFieldStr(LBound(strFieldN), intField) & intCounter & ":" & Chr(intField + 4) & intCounter).MergeCells = True
            intCounter = intCounter + 1
            .Range(GetFieldStr(LBound(strFieldN), intField) & intCounter).Value = "列印日期:" & CFDate(ACDate(ServerDate))
            .Range(GetFieldStr(LBound(strFieldN), intField) & intCounter & ":" & Chr(intField + 4) & intCounter).MergeCells = True
            intCounter = intCounter + 1
        
            For i = LBound(strFieldN) To UBound(strFieldN)
                If i = LBound(strFieldN) Then
                    .Range(GetFieldStr(LBound(strFieldN), intField) & intCounter & ":" & GetFieldStr(LBound(strFieldN), intField) & intCounter + 2).MergeCells = True
                    .Range(GetFieldStr(LBound(strFieldN), intField) & intCounter).ColumnWidth = 14
                Else
                    'MCTF0X
                    If i Mod 4 = 1 Then
                        stTmp = Replace(arrMCTF(i / 4 + 1), "'", "")
                        .Range(GetFieldStr(i, intField) & intCounter).Value = stTmp
                        .Range(GetFieldStr(i, intField) & intCounter & ":" & GetFieldStr(i + 3, intField) & intCounter).MergeCells = True
                        .Range(GetFieldStr(i, intField) & intCounter & ":" & GetFieldStr(i + 3, intField) & intCounter).HorizontalAlignment = xlCenter
                    End If
                    '收文/發文
                    If i Mod 4 = 1 Or i Mod 4 = 3 Then
                        stTmp = "收文"
                        If i Mod 4 = 3 Then stTmp = "發文"
                        .Range(GetFieldStr(i, intField) & intCounter + 1).Value = stTmp
                        .Range(GetFieldStr(i, intField) & intCounter + 1 & ":" & GetFieldStr(i + 1, intField) & intCounter + 1).MergeCells = True
                        .Range(GetFieldStr(i, intField) & intCounter + 1 & ":" & GetFieldStr(i + 1, intField) & intCounter + 1).HorizontalAlignment = xlCenter
                    End If
                    '件數/點數
                     .Range(GetFieldStr(i, intField) & intCounter + 2).Value = strFieldN(i)
                     .Range(GetFieldStr(i, intField) & intCounter + 2).ColumnWidth = 8
                End If
            Next i
            intTitleR = intCounter + 2
            intCounter = intCounter + 3
        '最後設定
        Else
            For i = LBound(strFieldN) + 1 To UBound(strFieldN)
                .Range(GetFieldStr(i, intField) & intTitleR).Value = Right(.Range(GetFieldStr(i, intField) & intTitleR).Value, 2)
                .Range(GetFieldStr(i, intField) & intTitleR).HorizontalAlignment = xlCenter
                If i = UBound(strFieldN) - 3 Then
                    .Range(GetFieldStr(i, intField) & intTitleR - 2).Value = "合　計"
                End If
            Next i
        End If
    End With
    'end 2021/12/15
End Sub
'end 2019/08/06

    
    
