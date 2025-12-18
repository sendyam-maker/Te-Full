VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm050310 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文點數明細表"
   ClientHeight    =   1500
   ClientLeft      =   5325
   ClientTop       =   4560
   ClientWidth     =   3195
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3195
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4200
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3585
      Top             =   3135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   2220
      TabIndex        =   8
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   1425
      TabIndex        =   7
      Top             =   60
      Width           =   756
   End
   Begin VB.TextBox Text4 
      Height          =   264
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1140
      Width           =   315
   End
   Begin VB.TextBox Text3 
      Height          =   264
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   5
      Top             =   840
      Width           =   915
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Left            =   960
      MaxLength       =   7
      TabIndex        =   4
      Top             =   840
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   960
      TabIndex        =   3
      Top             =   540
      Width           =   2115
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(Y列印)"
      Height          =   180
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   600
   End
   Begin VB.Line Line1 
      X1              =   1890
      X2              =   2130
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "是否列印明細："
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "發文日期："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   900
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "系統類別："
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   900
   End
End
Attribute VB_Name = "frm050310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit
Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String, Fee As Long, pnt As Integer, httlx As Integer, rcno As Integer, pbar As String, n As Integer, sp As String
Dim SeekPrint As Integer, SeekPrintL As Integer, pst As String, iLine As Integer, tpage As Integer, hbar As String, prcno As String, systype As String, httl As String, det As String, d1234 As String
Dim Atmp(1 To 9) As String
'Add By Cheng 2002/09/16
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub Command1_Click()

   'Add By Cheng 2002/09/16
   blnClkSure = False
   
     Printer.Orientation = 2
     DoEvents

    If Text1.Text = "" Then
        MsgBox "系統類別為必要輸入欄位", vbOKOnly, "注意"
        Text1.SetFocus
        Exit Sub
    End If
    
   'Add By Cheng 2002/03/20
   If PUB_CheckKeyInDate(Me.Text2) = -1 Then
      Me.Text2.SetFocus
      Text2_GotFocus
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Text3) = -1 Then
      Me.Text3.SetFocus
      Text3_GotFocus
      Exit Sub
   End If
    
    If Text2.Text = "" And Text3 = "" Then
        MsgBox "發文日期必須輸入 起始日期 或 終止日期", vbOKCancel + vbExclamation, "注意"
        Text2.SetFocus
        Exit Sub
    Else
        If Text2.Text <> "" And Text3.Text <> "" Then
            If Val(Text2.Text) > Val(Text3.Text) Then
                MsgBox "發文日期之起訖範圍錯誤 !!", vbOKOnly + vbExclamation, "注意"
                blnClkSure = True
                Text2.SetFocus
                Text2_GotFocus
                Exit Sub
            End If
        End If
    End If
    Screen.MousePointer = vbHourglass
    ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/3 清除查詢印表記錄檔欄位
    Command1.Enabled = False

    Dim fst As Boolean
    Dim lst As Boolean
        
    strSQL1 = ""
    strSQL2 = ""
    StrSQL3 = ""
    StrSQL4 = ""
    strSQL5 = ""
   If Len(Trim(Text3.Text)) <> 0 Then
      strSQL1 = strSQL1 + " and cp27<=" & Val(ChangeTStringToWString(Text3.Text)) & ""
   'Add By Cheng 2002/03/20
   Else
      If Len(Trim(Text2.Text)) <> 0 Then
         strSQL1 = strSQL1 + " and cp27<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & ""
      End If
   End If
   If Len(Trim(Text2.Text)) <> 0 Then
       strSQL1 = strSQL1 + " and cp27>=" & Val(ChangeTStringToWString(Text2.Text)) & ""
   End If
   If Len(Trim(Text2.Text)) <> 0 Or Len(Trim(Text3.Text)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label2(0) & Text2 & "-" & Text3 'Add By Sindy 2010/12/3
   End If
   
   strSQL2 = strSQL1
   StrSQL3 = strSQL1
   StrSQL4 = strSQL1
   strSQL5 = strSQL1
    If Len(Trim(Text1.Text)) <> 0 Then
      strSQL1 = strSQL1 & " and cp01 in (" & SQLGrpStr(Text1.Text, 1) & ") "
      strSQL2 = strSQL2 & " and cp01 in (" & SQLGrpStr(Text1.Text, 2) & ") "
      StrSQL3 = StrSQL3 & " and cp01 in (" & SQLGrpStr(Text1.Text, 3) & ") "
      StrSQL4 = StrSQL4 & " and cp01 in (" & SQLGrpStr(Text1.Text, 4) & ") "
      strSQL5 = strSQL5 & " and cp01 in (" & SQLGrpStr(Text1.Text, 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1 & Text1 'Add By Sindy 2010/12/3
    End If
    If Text4.Text = "Y" Then     ' 列印明細
      pub_QL05 = pub_QL05 & ";" & Left(Label3, 7) & Text4 'Add By Sindy 2010/12/3
    End If
    'Modify by Morgan 2010/8/12 百年蟲 SQLDate("cp27") --> substrb(' '||sqldatet(cp27),-9)
    strSql = "select substrb(' '||sqldatet(cp27),-9),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),ptm03,nvl(na03,na04),nvl(decode(pa09,'000',cpm03,cpm04),cp10), nvl(st02,cp14), cp16, cp18,'" & strUserNum & "',cp09 from caseprogress,patent,patenttrademarkmap,nation,casepropertymap,staff    where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp14=st01(+) and '1'=ptm01(+) and pa08=ptm02(+) and pa09=na01(+) " & strSQL1
    strSql = strSql & " union all select substrb(' '||sqldatet(cp27),-9),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(tm05,nvl(tm06,tm07)),decode(tm10,'000',ptm03,ptm04),nvl(na03,na04),nvl(decode(tm10,'000',cpm03,cpm04),cp10), nvl(st02,cp14), cp16, cp18,'" & strUserNum & "',cp09 from caseprogress,trademark,patenttrademarkmap,nation,casepropertymap,staff where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp14=st01(+) and '2'=ptm01(+) and tm08=ptm02(+) and tm10=na01(+) " & strSQL2
    strSql = strSql & " union all select substrb(' '||sqldatet(cp27),-9),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(lc05,nvl(lc06,lc07)),''                            ,nvl(na03,na04),nvl(decode(lc15,'000',cpm03,cpm04),cp10), nvl(st02,cp14), cp16, cp18,'" & strUserNum & "',cp09 from caseprogress,lawcase,nation,casepropertymap,staff                      where cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp14=st01(+) and lc15=na01(+) " & StrSQL3
    strSql = strSql & " union all select substrb(' '||sqldatet(cp27),-9),cp01||'-'||cp02||'-'||cp03||'-'||cp04,hc06                    ,''                            ,nvl(na03,na04),nvl(cpm03,cp10)                         , nvl(st02,cp14), cp16, cp18,'" & strUserNum & "',cp09 from caseprogress,hirecase,nation,casepropertymap,staff                     where cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp14=st01(+) and '000'=na01(+) " & StrSQL4
    strSql = strSql & " union all select substrb(' '||sqldatet(cp27),-9),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(sp05,nvl(sp06,sp07)),''                            ,nvl(na03,na04),nvl(decode(sp09,'000',cpm03,cpm04),cp10), nvl(st02,cp14), cp16, cp18,'" & strUserNum & "',cp09 from caseprogress,servicepractice,nation,casepropertymap,staff              where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp14=st01(+) and sp09=na01(+) " & strSQL5
    cnnConnection.Execute "delete from r050310 where id='" & strUserNum & "' "
    cnnConnection.Execute "insert into r050310 " & strSql
    CheckOC
    strSql = "select * from r050310 where id='" & strUserNum & "' order by 2,1 "
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    pst = ""
    If adoRecordset.RecordCount > 0 Then
        InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/3
        adoRecordset.MoveFirst
        If Text4.Text = "Y" Then     ' 列印明細
            Atmp(1) = "發文日　　"
            Atmp(2) = "本所案號　　　　"
            Atmp(3) = "案件名稱　　　　　　　　　　　　　　　"
            Atmp(4) = "專利種類　　"
            Atmp(5) = "申請國家　　"
            Atmp(6) = "案件性質　　"
            Atmp(7) = "承辦人　"
            Atmp(8) = "費　用　　　"
            Atmp(9) = "點　數　　　"
            iLine = 0
            fst = True
            lst = False
            tpage = 0
            Printer.Font.Name = "細明體"
            Do Until adoRecordset.EOF
                If SystemNumber(CheckStr(adoRecordset.Fields(1)), 1) <> pst Then
                    If fst = True Then
                        fst = False
                    Else
                        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(hbar)) / 2
                        Printer.Print hbar
                        
                        prcno = ""
                        prcno = prcno + Space(110) + "合計:" + Space(8 - Len(Format(Trim(str(Fee)), "###,###,##0"))) & Format(Trim(str(Fee)), "###,###,##0") + " " + Space(8 - Len(Format(Trim(str(pnt)), "###,###,##0.0"))) + Format(Trim(str(pnt)), "###,###,##0.0")
                        Printer.CurrentX = httlx
                        Printer.Print prcno
                        
                        prcno = ""
                        prcno = prcno + Space(110) + "合計筆數:" + Trim(str(rcno)) + "筆"
                        Printer.CurrentX = httlx
                        Printer.Print prcno
                                                    
                        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(hbar)) / 2
                        Printer.Print hbar
                        
                       Printer.NewPage
                    End If
                    rcno = 0    ' 合計筆數
                    Fee = 0     ' 費用
                    pnt = 0   ' 點數
                    tpage = tpage + 1
                    iLine = 0
                    pst = SystemNumber(CheckStr(adoRecordset.Fields(1)), 1)
                    
                    
                    Printer.Font.Size = 22
                    Printer.Font.Underline = True
                    Printer.Font.Name = "細明體"
                    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("發文點數明細表")) / 2
                    Printer.Print "發文點數明細表"
                    Printer.Font.Underline = False
                    Printer.Font.Size = 12
                                                            
                    pbar = ""
                    pbar = pbar + "發文日期:" + Format(ChangeTStringToTDateString(Text2.Text) & " ", "@@@@@@@@@@") & " - " & ChangeTStringToTDateString(Text3.Text)
                    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(pbar)) / 2
                    Printer.Print pbar
                    
                    systype = ""
                    systype = systype + "系統別:" + SystemNumber(CheckStr(adoRecordset.Fields(1)), 1)
                    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(systype)) / 2
                    Printer.Print systype
                    Printer.Print "列印人:" + GetPrjSalesNM(strUserNum)
                    
                    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期:" + ChangeTStringToTDateString(GetTaiwanTodayDate))
                    Printer.Print "列印日期:" + ChangeTStringToTDateString(GetTaiwanTodayDate)
                    
                    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期:" + ChangeTStringToTDateString(GetTaiwanTodayDate))
                    Printer.Print "頁　　次: " + Trim(str(tpage))
                    Printer.Font.Size = 12
                    'Printer.CurrentY = 1700
                    hbar = ""
                    For n = 1 To 150
                        hbar = hbar + "-"
                    Next
                    Printer.CurrentX = 0
                    'Printer.CurrentY = 1939
                    Printer.Print hbar
                                        
                    httl = ""
                    For n = 1 To 9
                        httl = httl & Atmp(n)
                    Next n
                    httlx = 0
                    Printer.CurrentX = httlx
                    'Printer.CurrentY = 2178
                    Printer.Print httl
                    
                    Printer.CurrentX = 0
                    'Printer.CurrentY = 2417
                    Printer.Print hbar
                End If
                 iLine = iLine + 1
                 
                rcno = rcno + 1
                Fee = Fee + Val(CheckStr(adoRecordset.Fields(7)))
                pnt = pnt + Val(CheckStr(adoRecordset.Fields(8)))
                             
                det = ""
                det = det + CheckStr(adoRecordset.Fields(0)) + " "
                Printer.CurrentX = 0
                Printer.CurrentY = (iLine - 1) * 239 + 2417
                Printer.Print det
                d1234 = ""
                d1234 = d1234 + CheckStr(adoRecordset.Fields(1))
                Printer.CurrentX = Printer.TextWidth(Atmp(1))
                Printer.CurrentY = (iLine - 1) * 239 + 2417
                Printer.Print d1234
                'sp = ""
                'For N = 1 To (16 - LenB(StrConv(d1234, vbFromUnicode)))
                '    sp = sp + " "
                'Next
                'det = det + d1234 + sp & " "
                
                ' 案件名稱
                If Len(CheckStr(adoRecordset.Fields(2))) <> 0 Then
                    Printer.CurrentX = Printer.TextWidth(Atmp(1) & Atmp(2))
                    Printer.CurrentY = (iLine - 1) * 239 + 2417
                    Printer.Print StrConv(MidB(StrConv(CheckStr(adoRecordset.Fields(2)), vbFromUnicode), 1, LenB(Atmp(3)) - 1), vbUnicode)
                End If
                
                ' 專利種類
                If Len(CheckStr(adoRecordset.Fields(3))) <> 0 Then
                    Printer.CurrentX = Printer.TextWidth(Atmp(1) & Atmp(2) & Atmp(3))
                    Printer.CurrentY = (iLine - 1) * 239 + 2417
                    Printer.Print StrConv(MidB(StrConv(CheckStr(adoRecordset.Fields(3)), vbFromUnicode), 1, LenB(Atmp(4)) - 1), vbUnicode)
                End If
                
                ' 申請國家
                If Len(CheckStr(adoRecordset.Fields(4))) <> 0 Then
                    Printer.CurrentX = Printer.TextWidth(Atmp(1) & Atmp(2) & Atmp(3) & Atmp(4))
                    Printer.CurrentY = (iLine - 1) * 239 + 2417
                    Printer.Print StrConv(MidB(StrConv(CheckStr(adoRecordset.Fields(4)), vbFromUnicode), 1, LenB(Atmp(5)) - 1), vbUnicode)
                End If
                
                ' 案件性質
                If Len(CheckStr(adoRecordset.Fields(5))) <> 0 Then
                    Printer.CurrentX = Printer.TextWidth(Atmp(1) & Atmp(2) & Atmp(3) & Atmp(4) & Atmp(5))
                    Printer.CurrentY = (iLine - 1) * 239 + 2417
                    Printer.Print StrConv(MidB(StrConv(CheckStr(adoRecordset.Fields(5)), vbFromUnicode), 1, LenB(Atmp(6)) - 1), vbUnicode)
                End If
                
                ' 承辦人
                If Len(CheckStr(adoRecordset.Fields(6))) <> 0 Then
                    Printer.CurrentX = Printer.TextWidth(Atmp(1) & Atmp(2) & Atmp(3) & Atmp(4) & Atmp(5) & Atmp(6))
                    Printer.CurrentY = (iLine - 1) * 239 + 2417
                    Printer.Print CheckStr(adoRecordset.Fields(6))
                End If
                
                ' 費用
                If Len(CheckStr(adoRecordset.Fields(7))) <> 0 Then
                    Printer.CurrentX = Printer.TextWidth(Atmp(1) & Atmp(2) & Atmp(3) & Atmp(4) & Atmp(5) & Atmp(6) & Atmp(7) & Atmp(8)) - 500 - Printer.TextWidth(Format(CheckStr(adoRecordset.Fields(7)), "###,###,##0"))
                    Printer.CurrentY = (iLine - 1) * 239 + 2417
                    Printer.Print Format(CheckStr(adoRecordset.Fields(7)), "###,###,##0")
                End If
                
                ' 點數
                If Len(CheckStr(adoRecordset.Fields(8))) <> 0 Then
                    Printer.CurrentX = Printer.TextWidth(Atmp(1) & Atmp(2) & Atmp(3) & Atmp(4) & Atmp(5) & Atmp(6) & Atmp(7) & Atmp(8) & Atmp(9)) - 500 - Printer.TextWidth(Format(CheckStr(adoRecordset.Fields(8)), "###,###,##0.0"))
                    Printer.CurrentY = (iLine - 1) * 239 + 2417
                    Printer.Print Format(CheckStr(adoRecordset.Fields(8)), "###,###,##0.0")
                End If
                'Printer.CurrentX = httlx
                'Printer.Print det
                'Printer.CurrentY = Printer.CurrentY + 300
                If iLine >= 32 Then
                     Printer.NewPage
                     tpage = tpage + 1
                    Printer.Font.Size = 22
                    Printer.Font.Underline = True
                    Printer.Font.Name = "細明體"
                    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("發文點數明細表")) / 2
                    Printer.Print "發文點數明細表"
                    Printer.Font.Underline = False
                    
                    Printer.Font.Size = 12
                                                            
                    pbar = ""
                    pbar = pbar + "發文日期:" + Format(ChangeTStringToTDateString(Text2.Text) & " ", "@@@@@@@@@@") & " - " & ChangeTStringToTDateString(Text3.Text)
                    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(pbar)) / 2
                    Printer.Print pbar
                    
                    systype = ""
                    systype = systype + "系統別:" + SystemNumber(CheckStr(adoRecordset.Fields(1)), 1)
                    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(systype)) / 2
                    Printer.Print systype
                    Printer.Print "列印人:" + GetPrjSalesNM(strUserNum)
                    
                    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期:" + ChangeTStringToTDateString(GetTaiwanTodayDate))
                    Printer.Print "列印日期:" + ChangeTStringToTDateString(GetTaiwanTodayDate)
                    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期:" + ChangeTStringToTDateString(GetTaiwanTodayDate))
                    Printer.Print "頁　　次: " + Trim(str(tpage))
                    
                    hbar = ""
                    For n = 1 To 150
                        hbar = hbar + "-"
                    Next
                    Printer.CurrentX = 0
                    Printer.Print hbar
                                        
                    httl = ""
                    For n = 1 To 9
                        httl = httl & Atmp(n)
                    Next n
                     httlx = 0
                    Printer.CurrentX = httlx
                    Printer.Print httl
                    
                    Printer.CurrentX = 0
                    Printer.Print hbar
                    
                    iLine = 0
                  End If
                adoRecordset.MoveNext
                If adoRecordset.EOF Then
                    lst = True
                End If
            Loop
            
            If lst = True Then
                Printer.CurrentX = 0
                Printer.Print hbar
                iLine = Printer.CurrentY
                prcno = "合計:" + Format(Trim(str(Fee)), "###,###,##0") + " "
                Printer.CurrentX = Printer.TextWidth(Atmp(1) & Atmp(2) & Atmp(3) & Atmp(4) & Atmp(5) & Atmp(6) & Atmp(7) & Atmp(8)) - 500 - Printer.TextWidth(prcno)
                Printer.CurrentY = iLine + 239
                Printer.Print prcno
                prcno = Format(Trim(str(pnt)), "###,###,##0.0")
                Printer.CurrentX = Printer.TextWidth(Atmp(1) & Atmp(2) & Atmp(3) & Atmp(4) & Atmp(5) & Atmp(6) & Atmp(7) & Atmp(8) & Atmp(9)) - 500 - Printer.TextWidth(prcno)
                Printer.CurrentY = iLine + 239
                Printer.Print prcno
                
                prcno = "合計:" + Format(Trim(str(Fee)), "###,###,##0") + " "
                Printer.CurrentX = Printer.TextWidth(Atmp(1) & Atmp(2) & Atmp(3) & Atmp(4) & Atmp(5) & Atmp(6) & Atmp(7) & Atmp(8)) - 500 - Printer.TextWidth(prcno)
                Printer.CurrentY = iLine + 239 + 239
                Printer.Print "合計筆數:" + Trim(str(rcno)) + "筆"
                Printer.CurrentX = 0
                Printer.Print hbar
            End If
            Printer.EndDoc
        Else                                            ' 不列印明細
            'CheckOC
            'strSQL = "select * from r050310 where id='" & strUserNum & "' order by 2 "
            'adoRecordset.CursorLocation = adUseClient
            'adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
            fst = True
            lst = False
            tpage = 0
            
            
            Do Until adoRecordset.EOF
                If SystemNumber(CheckStr(adoRecordset.Fields(1)), 1) <> pst Then
                    If fst = True Then
                        fst = False
                    Else
                        prcno = ""
                        'prcno = prcno + "合計筆數:" + Trim(Str(rcno)) + " 費用:" + Format(Trim(Str(Fee)), "###,###,##0") + " 點數:" + Format(Trim(Str(pnt)), "###,###,##0.0")
                        prcno = prcno + "合計筆數:" + Trim(str(rcno)) + "       費用:" + Format(Trim(str(Fee)), "###,###,##0") + "         點數:" + Format(Trim(str(pnt)), "###,###,##0.0")
                        
                        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(prcno)) / 2
                        Printer.Print prcno
                            
                        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(hbar)) / 2
                        Printer.Print hbar
                        Printer.NewPage
                    End If
                    rcno = 0    ' 合計筆數
                    Fee = 0     ' 費用
                    pnt = 0   ' 點數
                    tpage = tpage + 1
                
                    pst = SystemNumber(CheckStr(adoRecordset.Fields(1)), 1)
                    
                    'Printer.Orientation = 2
                    Printer.Font.Size = 22
                    
                    Printer.Font.Underline = True
                    Printer.Font.Name = "細明體"
                    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("發文點數明細表")) / 2
                    Printer.Print "發文點數明細表"
                    Printer.Font.Underline = False
                                        
                    Printer.Font.Size = 12
                                        
                    pbar = ""
                    pbar = pbar + "發文日期:" + Format(ChangeTStringToTDateString(Text2.Text) & " ", "@@@@@@@@@@") & " - " & ChangeTStringToTDateString(Text3.Text)
                    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(pbar)) / 2
                    Printer.Print pbar
                    
                    systype = ""
                    systype = systype + "系統別:" + SystemNumber(CheckStr(adoRecordset.Fields(1)), 1)
                    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(systype)) / 2
                    Printer.Print systype
                    Printer.Print "列印人:" + strUserNum
                    
                    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期:" + ChangeTStringToTDateString(GetTaiwanTodayDate))
                    Printer.Print "列印日期:" + ChangeTStringToTDateString(GetTaiwanTodayDate)
                    
                    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期:" + ChangeTStringToTDateString(GetTaiwanTodayDate))
                    Printer.Print "頁　　次: " + Trim(str(tpage))
                    
                    hbar = ""
                    For n = 1 To 150
                        hbar = hbar + "-"
                    Next
                    Printer.CurrentX = 0
                    Printer.Print hbar
                    httl = ""
                    For n = 1 To 9
                        httl = httl & Atmp(n)
                    Next n
                    httlx = 0
                    Printer.CurrentX = httlx
                    Printer.Print httl
                    
                    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(hbar)) / 2
                    Printer.Print hbar
                    
                End If
                    
                rcno = rcno + 1
                Fee = Fee + Val(CheckStr(adoRecordset.Fields(7)))
                pnt = pnt + Val(CheckStr(adoRecordset.Fields(8)))
                
                adoRecordset.MoveNext
                If adoRecordset.EOF Then
                    lst = True
                End If
            Loop
            
            If lst = True Then
                prcno = ""
                prcno = prcno + "合計筆數:" + Trim(str(rcno)) + "       費用:" + Format(Trim(str(Fee)), "###,###,##0") + "         點數:" + Format(Trim(str(pnt)), "###,###,##0.0")
                Printer.CurrentX = 0
                Printer.Print prcno
                            
                Printer.CurrentX = 0
                Printer.Print hbar
            End If
            Printer.EndDoc
        End If
    Else
        'MsgBox "沒有符合條件之紀錄 可供列印", vbOKOnly + vbExclamation, "注意"
        InsertQueryLog (0) 'Add By Sindy 2010/12/3
        ShowNoData
        Screen.MousePointer = vbDefault
        Command1.Enabled = True
        Exit Sub
    End If
    CheckOC
    Command1.Enabled = True
    ShowPrintOk
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHandler:
    Command1.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    Text1.Text = GetSystemKindByNick
'strSQL = Printer.DeviceName
'SeekPrintL = Printer.Orientation
'For i = 0 To Printers.Count - 1
'    Set Printer = Printers(i)
'    If Printer.DeviceName <> strSQL Then
'        Combo1.AddItem Printer.DeviceName, j
'        j = j + 1
'    End If
'    If Printer.DeviceName = strSQL Then
'        SeekPrint = i
'    End If
'Next i
'Combo1.Text = Combo1.List(0)

End Sub


Private Sub Form_Unload(Cancel As Integer)
'Set Printer = Printers(SeekPrint)
'Printer.Orientation = SeekPrintL

 Set frm050310 = Nothing
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Text1.Text <> "" Then
        'Text1.Text = StrConv(Text1.Text, vbUpperCase)
        Dim ttlstr, relstr, i, j
        
        ttlstr = Split(Text1.Text, ",")
        relstr = Split(GetSystemKindByNick, ",")
        
        ' i 迴圈代表Text1.text中的所有系統類別
        ' j 迴圈代表該使用者實際所能查詢的系統類別
        ' 並且作使用者實際所能查詢類別與所輸入類別是否相符
        Dim Cl As Boolean, ClStr As String
        Cl = True
        
        For i = 0 To UBound(ttlstr)
            ClStr = ttlstr(i)
            For j = 0 To UBound(relstr)
                If ttlstr(i) = relstr(j) Then
                    Cl = False
                End If
            Next j
            If Cl = True Then
                MsgBox ttlstr + "非屬於您所能列印之權限報表範圍", vbOKOnly + vbExclamation, "注意"
                Exit Sub
            End If
        Next i
    End If
End Sub


Private Sub Text2_GotFocus()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    If Text2.Text <> "" Then
        If CheckIsTaiwanDate(Text2.Text) = True Then
            Cancel = False
        Else
            Text2.SelStart = 0
            Text2.SelLength = Len(Text2.Text)
            Cancel = True
        End If
    End If
End Sub

Private Sub Text3_GotFocus()
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_LostFocus()
   'Modify By Cheng 2002/09/16
   If blnClkSure = False Then
      If RunNick(Text2, Text3) Then
         Text2.SetFocus
         Text2_GotFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
    If Text3.Text <> "" Then
        If CheckIsTaiwanDate(Text3.Text) = True Then
            Cancel = False
        Else
            Text3.SelStart = 0
            Text3.SelLength = Len(Text3.Text)
            Cancel = True
        End If
    End If
End Sub

Private Sub Text4_GotFocus()
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
        Case 8
        Case 89
        Case Else
            KeyAscii = 0
    End Select
End Sub
