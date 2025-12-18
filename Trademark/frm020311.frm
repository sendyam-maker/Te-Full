VERSION 5.00
Begin VB.Form frm020311 
   BorderStyle     =   1  '單線固定
   Caption         =   "智慧局註冊費通知函列印"
   ClientHeight    =   3360
   ClientLeft      =   1650
   ClientTop       =   1530
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4770
   Begin VB.CheckBox Check1 
      Caption         =   "僅產生PDF電子檔清單，寄給智權人員"
      ForeColor       =   &H00004000&
      Height          =   220
      Left            =   240
      TabIndex        =   14
      Top             =   2100
      Width           =   3700
   End
   Begin VB.FileListBox File1 
      Height          =   420
      Left            =   60
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.OptionButton radio 
      Caption         =   "延展(定稿＋清單)"
      Height          =   252
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   1050
      Width           =   3240
   End
   Begin VB.OptionButton radio 
      Caption         =   "繳納註冊費(定稿＋清單)"
      Height          =   252
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   3240
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   330
      Index           =   0
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   330
      Index           =   1
      Left            =   3870
      TabIndex        =   5
      Top             =   120
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   260
      Left            =   1560
      Style           =   2  '單純下拉式
      TabIndex        =   3
      Top             =   2450
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.ListBox List1 
      Height          =   580
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.OptionButton radio 
      Caption         =   "第二期註冊費(定稿＋清單)"
      Height          =   252
      Index           =   2
      Left            =   1440
      TabIndex        =   1
      Top             =   270
      Visible         =   0   'False
      Width           =   3240
   End
   Begin VB.Label Label6 
      Caption         =   "清單產生PDF電子檔，直接寄給智權人員！"
      ForeColor       =   &H00C00000&
      Height          =   230
      Left            =   240
      TabIndex        =   13
      Top             =   1470
      Width           =   4310
   End
   Begin VB.Label Label5 
      Caption         =   "PS：因依輸入順序排序且同一代理人或申請人只印一份，　　所以地址條會有空號！"
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   80
      TabIndex        =   10
      Top             =   2840
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "定稿清單印表機："
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   1740
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "通知函性質： "
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "地址條　印表機："
      Height          =   180
      Left            =   80
      TabIndex        =   7
      Top             =   2490
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1700
      TabIndex        =   6
      Top             =   1740
      Width           =   2850
   End
End
Attribute VB_Name = "frm020311"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
'2007/9/10 add by sonia
Option Explicit

Dim intWhere As Integer, strReceiveNo As String, PLeft(0 To 7) As Integer
' 預設印表機
Dim m_DefaultPrinter As String
Dim strSql As String, i As Integer, j As Integer, s As Integer
Dim iPrint As Integer, Page As Integer, strTemp(0 To 20) As String
'Dim m_strTM23Nation As String '申請人國籍              '2011/12/26 cancel by sonia 改用tm77
Dim m_strSales As String '智權人員
Dim m_CP09 As String, m_CP13 As String
Dim m_Print As String   '2011/12/26 add by sonia
'Dim boleFileSave As Boolean
Dim m_TM01 As String, m_TM02 As String, m_TM03 As String, m_TM04 As String
Dim m_Appl As String, m_LimitDt As String, m_CaseNo As String 'Add By Sindy 2013/11/4
Dim m_CP10 As String, m_TM23 As String, m_TM44 As String 'Add By Sindy 2019/11/1
Dim m_AttachPath As String
Dim m_TM45 As String, m_TM05 As String, m_TM09 As String, m_TM15 As String 'Add By Sindy 2020/2/6
Dim m_strIsClose As String 'Add By Sindy 2020/7/30


Private Sub cmdok_Click(Index As Integer)
Dim Cancel As Boolean
   
   Select Case Index
      Case 0
         Screen.MousePointer = vbHourglass
         DoEvents
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/15 清除查詢印表記錄檔欄位
'         If radio(0).Value = True Then '繳納註冊費
            Progress
'         Else '通知延展
'            Progress1717
'         End If
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
Dim PrinterIndex As Integer
   
   MoveFormToCenter Me
   
   PUB_SetPrinter Me.Name, Combo1, m_DefaultPrinter, False 'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
   ' 暫存預設印表機
   Label3.Caption = m_DefaultPrinter
   
   'Modify by Amy 2022/03/09 地址條已不使用-桂英
   Me.Height = 2835
'   '刪除地址條列表資料
'   PUB_DeleteAddressList strUserNum
   'end 2022/03/09
   '初始化序號
   pub_AddressListSN = 0
   
   'Add By Sindy 2019/11/1
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum
   '檢查是否有安裝PDFCreator
   PrinterIndex = -1
   For i = 0 To Printers.Count - 1
    If UCase(Printers(i).DeviceName) = UCase$("PDFCreator") Then
     PrinterIndex = i
     Exit For
    End If
   Next i
   If PrinterIndex < 0 Then
      MsgBox "請通知電腦中心安裝PDFCreator !!!"
      Exit Sub
   End If
   '2019/11/1 END
   
   'Add By Sindy 2024/9/3 暫檢查中
   frm020311.Tag = frm020311.Caption
   Check1.Value = 0
   If Pub_StrUserSt03 = "M51" Or strSrvDate(1) = 20241016 Then
      Check1.Visible = True
   Else
      Check1.Visible = False
   End If
   '2024/9/3 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Prn As Printer
   
   'Mark by Amy 2022/03/09 地址條已不使用-桂英
'   'DoEvents
'   '一個代理人或申請人印一個名條
'   PUB_PrintAddressList strUserNum, Combo1.Text, True
'   '刪除地址條列表資料
'   PUB_DeleteAddressList strUserNum
   'end 2022/03/09
   '初始化序號
   pub_AddressListSN = 0
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   Set frm020311 = Nothing
End Sub

Sub Progress()
Dim StrStaff As String
'Dim StrSaGp As String
Dim rsTmp1 As New ADODB.Recordset
Dim strTM23 As String
Dim bolExists As Boolean
Dim intRow As Integer
Dim strUpdCP09 As String, strUpdTM01 As String, strUpdTM02 As String, strUpdTM03 As String, strUpdTM04 As String 'Add By Sindy 2020/2/4
Dim strCuName As String, strMailSubject As String, StrMailContent As String 'Add By Sindy 2020/2/21
Dim varTemp As Variant, ii As Integer 'Add By Sindy 2020/2/24
Dim strSQLCon As String 'Add By Sindy 2024/9/3
   
On Error GoTo ErrorHandler

   cnnConnection.Execute "delete from r020311 where id='" & strUserNum & "'"
   
   'Modify By Sindy 2024/9/3
   If Check1.Value = 1 Then
      strSQLCon = " and cp27=" & strSrvDate(1)
   Else
      strSQLCon = " and cp27 is null"
   End If
   '2024/9/3 END
   
   If rsTmp1.State = 1 Then rsTmp1.Close
'   boleFileSave = False 'Add By Sindy 2012/1/16
   'Modify By Sindy +,tm01,tm02,tm03,tm04
   If radio(0).Value = True Then      '繳納註冊費
      pub_QL05 = pub_QL05 & ";" & Label1 & radio(0).Caption 'Add By Sindy 2010/10/15
      'Modify By Sindy 2010/11/10 增加CP13
      '2012/12/20 MODIFY BY Sindy 第一期註冊費715改為717註冊費,通知繳納第一期註冊費1715改為1720通知繳納註冊費
      'Modify By Sindy 2013/11/4 +order by
      'Modify By Sindy 2013/11/6 sqldatet(np09) --> sqldatet(max(np09))
      'Modify By Sindy 2013/11/6 +group by cp01,cp02,cp03,cp04,cu104,cu04,tm44,tm23,tm09,tm05,tm12,cp09,cu10,cp13,tm77,tm53,tm01,tm02,tm03,tm04,cp64,cp65,tm26
      'Modify By Sindy 2019/11/1 + ,cp10,tm23 as T_tm23,tm44 as T_tm44,tm45
      'Modify By Sindy 2020/7/30 + ,decode(nvl(tm57,0),0,decode(tm29,'Y','Y',decode(np06,'N','Y','N')),'Y') isclose
      strSql = "select cp01,cp02,cp03,cp04,NVL(cu104,cu04) as CUName,nvl(tm44,tm23) tm23,tm09,tm05,tm12 as tm15,sqldatet(max(np09)) as np09,cp09,cu10,cp13,nvl(tm77,tm53) tm77,tm01,tm02,tm03,tm04,cp64,cp10,tm23 as T_tm23,tm44 as T_tm44,tm45,decode(nvl(tm57,0),0,decode(tm29,'Y','Y',decode(np06,'N','Y','N')),'Y') isclose " & _
               "From Caseprogress,Trademark,Customer,Nextprogress " & _
               " where cp01='T' and cp10= '1720'" & strSQLCon & " and cp57 is null " & _
               " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
               " and cp01=np02(+) and cp02=np03(+) and cp03=np04(+) and cp04=np05(+) and np07 in ('715','717') " & _
               " and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) " & _
               " group by cp01,cp02,cp03,cp04,cu104,cu04,tm44,tm23,tm09,tm05,tm12,cp09,cu10,cp13,tm77,tm53,tm01,tm02,tm03,tm04,cp64,cp65,tm26,cp10,tm23,tm44,tm45,decode(nvl(tm57,0),0,decode(tm29,'Y','Y',decode(np06,'N','Y','N')),'Y') " & _
               " order by cp65,cp09,tm26,cp01,cp02,cp03,cp04 "
               '" order by cp65,tm23,cp09,cp01,cp02,cp03,cp04 "
               '" order by tm23,cp01,cp02,cp03,cp04 "
   'Modify By Sindy 2019/10/29 Mark
'   ElseIf radio(1).Value = True Then  '第二期註冊費
'      pub_QL05 = pub_QL05 & ";" & Label1 & radio(1).Caption 'Add By Sindy 2010/10/15
'      'Modify By Sindy 2010/11/10 增加CP13
'      'Modify By Sindy 2013/11/4 +order by
'      strSql = "select cp01,cp02,cp03,cp04,NVL(cu104,cu04) as CUName,nvl(tm44,tm23) tm23,tm09,tm05,tm15,sqldatet(np09) as np09,cp09,cu10,cp13,nvl(tm77,tm53) tm77,tm01,tm02,tm03,tm04,cp64 From Caseprogress,Trademark,Customer,Nextprogress " & _
'               " where cp01='T' and cp10= '1716' and cp27 is null and cp57 is null " & _
'               " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
'               " and cp01=np02(+) and cp02=np03(+) and cp03=np04(+) and cp04=np05(+) and '716'=np07(+) " & _
'               " and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) " & _
'               " order by cp65,cp09,tm26,cp01,cp02,cp03,cp04 "
   Else '延展
      pub_QL05 = pub_QL05 & ";" & Label1 & radio(1).Caption 'Add By Sindy 2010/10/15
      'Modify By Sindy 2010/11/10 增加CP13
      'Modify By Sindy 2013/11/4 +order by
      'Modify By Sindy 2013/11/6 sqldatet(np09) --> sqldatet(max(np09))
      'Modify By Sindy 2013/11/6 +group by cp01,cp02,cp03,cp04,cu104,cu04,tm44,tm23,tm09,tm05,tm15,cp09,cu10,cp13,tm77,tm53,tm01,tm02,tm03,tm04,cp64,cp65,tm26
      'Modify By Sindy 2019/11/1 + ,cp10,tm23 as T_tm23,tm44 as T_tm44,tm45
      'Modify By Sindy 2020/7/30 + ,decode(nvl(tm57,0),0,decode(tm29,'Y','Y',decode(np06,'N','Y','N')),'Y') isclose
      strSql = "select cp01,cp02,cp03,cp04,NVL(cu104,cu04) as CUName,nvl(tm44,tm23) tm23,tm09,tm05,tm15,sqldatet(max(np09)) as np09,cp09,cu10,cp13,nvl(tm77,tm53) tm77,tm01,tm02,tm03,tm04,cp64,cp10,tm23 as T_tm23,tm44 as T_tm44,tm45,decode(nvl(tm57,0),0,decode(tm29,'Y','Y',decode(np06,'N','Y','N')),'Y') isclose " & _
               "From Caseprogress,Trademark,Customer,Nextprogress " & _
               " where cp01='T' and cp10= '1717'" & strSQLCon & " and cp57 is null " & _
               " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
               " and cp01=np02(+) and cp02=np03(+) and cp03=np04(+) and cp04=np05(+) and '102'=np07(+) " & _
               " and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) " & _
               " group by cp01,cp02,cp03,cp04,cu104,cu04,tm44,tm23,tm09,tm05,tm15,cp09,cu10,cp13,tm77,tm53,tm01,tm02,tm03,tm04,cp64,cp65,tm26,cp10,tm23,tm44,tm45,decode(nvl(tm57,0),0,decode(tm29,'Y','Y',decode(np06,'N','Y','N')),'Y') " & _
               " order by cp65,cp09,tm26,cp01,cp02,cp03,cp04 "
               '" order by cp65,tm23,cp09,cp01,cp02,cp03,cp04 "
               '" order by tm23,cp01,cp02,cp03,cp04 "
   End If
   'strSql = strSql & " order by cp65,cp09,tm26,cp01,cp02,cp03,cp04 "
   rsTmp1.CursorLocation = adUseClient
   rsTmp1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp1.RecordCount <> 0 Then
      InsertQueryLog (rsTmp1.RecordCount) 'Add By Sindy 2010/10/15
      
      With rsTmp1
         'Modify By Sindy 2024/9/3
         If Check1.Value = 1 Then
            'Modify By Sindy 2013/11/5 統一後面再更新資料
            .MoveFirst
            Do While Not .EOF
               '抓智權人員
               StrStaff = PUB_GetAKindSalesNo("" & .Fields("cp01"), "" & .Fields("cp02"), "" & .Fields("cp03"), "" & .Fields("cp04"))
               m_CP09 = .Fields("CP09")
               '產生暫存檔
               cnnConnection.Execute "insert into r020311 (r01,r02,r03,r04,r05,r06,r07,r08,r09,r10,r11,id) values ('" & "" & .Fields("cp01") & "','" & "" & .Fields("cp02") & "','" & "" & .Fields("cp03") & "','" & "" & .Fields("cp04") & "','" & "" & .Fields("tm09") & "','" & "" & ChgSQL(.Fields("tm05")) & "','" & "" & .Fields("tm15") & "','" & ChgSQL("" & .Fields("CUName")) & " ','" & "" & .Fields("np09") & "','" & StrStaff & "','" & "" & .Fields("tm23") & "','" & strUserNum & "' ) "
               .MoveNext
            Loop
            
            '印智權人員案件清單
            StrMenu
            MsgBox "執行完畢！", vbInformation
            
            Exit Sub
         End If
         '2024/9/3 END
         
         .MoveFirst
         'Add By Sindy 2013/11/4
         strTM23 = ""
         List1.Clear '記錄已出過定稿的客戶編號
         intRow = 0
         '2013/11/4 END
         Do While Not .EOF
            intRow = intRow + 1 'Add By Sindy 2013/11/5
            'Add By Sindy 2012/1/17
            m_TM01 = .Fields("tm01")
            m_TM02 = .Fields("tm02")
            m_TM03 = .Fields("tm03")
            m_TM04 = .Fields("tm04")
            '2012/1/17 End
            
            '2011/12/26 modify by sonia T-160531 改用tm77
'            '取得申請人國籍
'            If Not IsNull(.Fields("CU10")) Then
'               m_strTM23Nation = .Fields("cu10")
'            Else
'               m_strTM23Nation = "001"
'            End If
                        
            'Modify By Sindy 2020/2/4
'            m_Print = CheckStr(.Fields("TM77"))
'            '2011/12/26 end
            If IsNull(.Fields("TM77")) = False Then
               m_Print = CheckStr(.Fields("TM77"))
            Else
               m_Print = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
            End If
            '2020/2/4 END
            m_strIsClose = CheckStr(.Fields("isclose")) 'Add By Sindy 2020/7/30 是否已結案
            
            m_CP09 = .Fields("CP09")
            m_CP13 = .Fields("CP13") 'Add By Sindy 2010/11/10
            'Add By Sindy 2019/11/1
            m_CP10 = .Fields("CP10")
            m_TM23 = "" & .Fields("T_tm23")
            m_TM44 = "" & .Fields("T_tm44")
            strCuName = "" & .Fields("CuName")
            strMailSubject = ""
            StrMailContent = ""
            '2019/11/1 END
            
            'Modify By Sindy 2019/10/29 Mark
'            '第二期註冊費才印定稿
'            If radio(1).Value = True Then
'               PrintLetter
'               'Modify by Morgan 2008/10/24 第二期註冊費預定97.11.17起改用開窗信封不必印地址條
'               If Val(strSrvDate(1)) < 20081117 Then
'                  pub_AddressListSN = pub_AddressListSN + 1
'                  PUB_AddNewAddressList strUserNum, .Fields("cp01").Value, .Fields("cp02").Value, .Fields("cp03").Value, .Fields("cp04").Value, "" & pub_AddressListSN, "0"
'               End If
'            Else
'               '加入地址條
'               pub_AddressListSN = pub_AddressListSN + 1
'               PUB_AddNewAddressList strUserNum, .Fields("cp01").Value, .Fields("cp02").Value, .Fields("cp03").Value, .Fields("cp04").Value, "" & pub_AddressListSN, "0"
               'Add By Sindy 2013/11/4
               strTM23 = .Fields("tm23")
               '檢查此客戶是否已出過定稿
               bolExists = False
               For i = 0 To List1.ListCount - 1
                  'Modify By Sindy 2020/2/21 若為Y編號需要再比對X編號是否相同
                  If Left(strTM23, 1) = "Y" Then
                     If InStr(List1.List(i), m_TM23) > 0 And InStr(List1.List(i), m_TM44) > 0 Then
                        'Add By Sindy 2020/7/30 增加檢查結案狀況
                        'Modify By Sindy 2025/7/17 mark
'                        If m_CP10 = "1717" Then '通知延展
                        '2025/7/17 END
                           If InStr(List1.List(i), ":" & m_strIsClose & ":") > 0 Then
                              bolExists = True
                              Exit For
                           End If
'                        Else
'                        '2020/7/30 END
'                           bolExists = True
'                           Exit For
'                        End If
                     End If
                  Else
                     'If strTM23 = List1.List(i) Then
                     If InStr(List1.List(i), strTM23) > 0 Then
                  '2020/2/21 END
                        'Add By Sindy 2020/7/30 增加檢查結案狀況
                        'Modify By Sindy 2025/7/17 mark
'                        If m_CP10 = "1717" Then '通知延展
                        '2025/7/17 END
                           If InStr(List1.List(i), ":" & m_strIsClose & ":") > 0 Then
                              bolExists = True
                              Exit For
                           End If
'                        Else
'                        '2020/7/30 END
'                           bolExists = True
'                           Exit For
'                        End If
                     End If
                  End If
               Next i
               If bolExists = False Then
                  '通知函最後繳費日:
                  If InStr(.Fields("cp64"), "通知函最後繳費日:") > 0 Then
                     m_LimitDt = Mid(.Fields("cp64"), InStr(InStr(.Fields("cp64"), "通知函最後繳費日:"), .Fields("cp64"), ":") + 1, 7)
                     m_LimitDt = Mid(m_LimitDt, 1, 3) & "年" & Mid(m_LimitDt, 4, 2) & "月"
                  End If
                  
                  '組多個案號及號數
                  'Modify By Sindy 2020/2/21 + , m_TM44, m_CP09, m_TM01, m_TM02, m_TM03, m_TM04, strCuName, strMailSubject, StrMailContent
                  'Modify By Sindy 2020/7/30 + m_strIsClose
                  Call GetTextNo(rsTmp1, m_TM23, m_TM44, m_CP09, m_TM01, m_TM02, m_TM03, m_TM04, strCuName, strMailSubject, StrMailContent, m_strIsClose)
                  
                  'Add By Sindy 2019/11/1
                  If strSrvDate(1) >= T商標電子化啟用日 Then
                     '新增信函進度
                     '傳FC代理人(tm44)
                     '傳是否大宗發文(pbolBulk=True)
                     'Modify By Sindy 2025/4/23 1720(通知繳納註冊費)件數不多,取消為大宗 + IIf(radio(0).Value = True, False, True)
                     PUB_AddLetterProgress m_CP09, 0, True, "", True, m_TM23, m_CP10, m_TM44, , , IIf(radio(0).Value = True, False, True)
                     '更新EMail寄送主旨,EMail寄送內文(後面會再更新一次,依本所案號順序)
                     If strMailSubject <> "" Then
                        strSql = "update letterprogress set lp44='" & strMailSubject & "',lp45='" & StrMailContent & "'" & _
                                 " where lp01='" & m_CP09 & "'"
                        cnnConnection.Execute strSql
                     End If
                  End If
                  '2019/11/1 END
                  PrintLetter
                  .MoveFirst
                  For i = 1 To intRow - 1
                     .MoveNext
                  Next i
               'Add By Sindy 2020/2/4 合併定稿,信函還是要記錄資訊
               Else
                  '定稿合併收文號
                  varTemp = Split(List1.List(i), ":")
                  For ii = 0 To UBound(varTemp)
                     'If ii = 2 Then
                     If ii = 3 Then
                        strUpdCP09 = varTemp(ii)
                     'ElseIf ii = 3 Then
                     ElseIf ii = 4 Then
                        strExc(10) = varTemp(ii)
                        strUpdTM01 = SystemNumber(strExc(10), 1)
                        strUpdTM02 = SystemNumber(strExc(10), 2)
                        strUpdTM03 = SystemNumber(strExc(10), 3)
                        strUpdTM04 = SystemNumber(strExc(10), 4)
                     End If
                  Next ii
                  '新增信函進度
                  'Modify By Sindy 2025/4/23 1720(通知繳納註冊費)件數不多,取消為大宗 + IIf(radio(0).Value = True, False, True)
                  PUB_AddLetterProgress m_CP09, 0, False, "", False, m_TM23, m_CP10, m_TM44, , , IIf(radio(0).Value = True, False, True)
                  strExc(0) = "已併入" & IIf(radio(0).Value = True, "通知繳納註冊費", "通知延展") & "通知函(" & IIf(strUpdTM03 & strUpdTM04 = "000", strUpdTM01 & "-" & strUpdTM02, strUpdTM01 & "-" & strUpdTM02 & "-" & strUpdTM03 & "-" & strUpdTM04) & ":" & strUpdCP09 & ")告知客戶;"
                  '更新客戶函上N.不通知
                  strSql = "update letterprogress set lp10='N',lp06='" & strUserNum & "',lp07=to_char(sysdate,'yyyymmdd'),lp12='" & strExc(0) & "'||lp12,lp42='" & strUpdCP09 & "'" & _
                           " where lp01='" & m_CP09 & "'"
                  cnnConnection.Execute strSql
               '2020/2/4 END
               End If
               '2013/11/4 END
'            End If
'            cnnConnection.Execute "insert into r020311 (r01,r02,r03,r04,r05,r06,r07,r08,r09,r10,r11,id) values ('" & "" & .Fields("cp01") & "','" & "" & .Fields("cp02") & "','" & "" & .Fields("cp03") & "','" & "" & .Fields("cp04") & "','" & "" & .Fields("tm09") & "','" & "" & ChgSQL(.Fields("tm05")) & "','" & "" & .Fields("tm15") & "','" & ChgSQL("" & .Fields("CUName")) & " ','" & "" & .Fields("np09") & "','" & StrStaff & "','" & "" & .Fields("tm23") & "','" & strUserNum & "' ) "
'            '印過上發文日
'            cnnConnection.Execute "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_CP09 & "'"
            .MoveNext
         Loop
         
         'Modify By Sindy 2013/11/5 統一後面再更新資料
         .MoveFirst
         Do While Not .EOF
            '抓智權人員
            StrStaff = PUB_GetAKindSalesNo("" & .Fields("cp01"), "" & .Fields("cp02"), "" & .Fields("cp03"), "" & .Fields("cp04"))
            '業務區
            'StrSaGp = PUB_GetStaffST15(StrStaff, "1")
            m_CP09 = .Fields("CP09")
            cnnConnection.Execute "insert into r020311 (r01,r02,r03,r04,r05,r06,r07,r08,r09,r10,r11,id) values ('" & "" & .Fields("cp01") & "','" & "" & .Fields("cp02") & "','" & "" & .Fields("cp03") & "','" & "" & .Fields("cp04") & "','" & "" & .Fields("tm09") & "','" & "" & ChgSQL(.Fields("tm05")) & "','" & "" & .Fields("tm15") & "','" & ChgSQL("" & .Fields("CUName")) & " ','" & "" & .Fields("np09") & "','" & StrStaff & "','" & "" & .Fields("tm23") & "','" & strUserNum & "' ) "
            '印過上發文日
            cnnConnection.Execute "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_CP09 & "'"
            PUB_UpdateLP03 m_CP09 'Add By Sindy 2020/2/6
            .MoveNext
         Loop
      End With
      
      'Add By Sindy 2020/10/26 最後再組合多案的EMail內容,必須依本所案號從小到大排序 ex:T-172505,T-172506,T-172507
      If radio(0).Value = True Then '繳納註冊費
         strSql = "SELECT lp01,lp44 from caseprogress,letterprogress" & _
                  " where cp01='T' and cp10='1720' and cp27=" & strSrvDate(1) & " and cp57 is null" & _
                  " and cp09=lp01 and lp44 is not null"
      Else
         '延展
         strSql = "SELECT lp01,lp44 from caseprogress,letterprogress" & _
                  " where cp01='T' and cp10='1717' and cp27=" & strSrvDate(1) & " and cp57 is null" & _
                  " and cp09=lp01 and lp44 is not null"
      End If
      intI = 1
      If rsTmp1.State = 1 Then rsTmp1.Close
      Set rsTmp1 = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         rsTmp1.MoveFirst
         Do While Not rsTmp1.EOF
            strMailSubject = "": StrMailContent = ""
            Call GetTextNo_2(rsTmp1.Fields("lp44").Value, IIf(radio(0).Value = True, "1720", "1717"), strMailSubject, StrMailContent)
            If strMailSubject <> "" Then
               '更新EMail寄送主旨,EMail寄送內文
               strSql = "update letterprogress set lp44='" & strMailSubject & "',lp45='" & StrMailContent & "'" & _
                        " where lp01='" & rsTmp1.Fields("lp01").Value & "'"
               cnnConnection.Execute strSql
            End If
            rsTmp1.MoveNext
         Loop
      End If
      '2020/10/26 END
      
      '印智權人員案件清單
      frm020311.Caption = frm020311.Tag & "(StrMenu)" 'Add By Sindy 2024/9/3
      StrMenu
      
'      'Add By Sindy 2012/1/16
'      If boleFileSave = True Then
'         MsgBox "列印結束，電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
'      Else
'      '2012/1/16 End
'         'MsgBox "列印結束 !", vbInformation
         MsgBox "執行完畢！", vbInformation
'      End If
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/10/15
      MsgBox "無符合條件之資料可列印！", vbInformation
   End If
   
   Exit Sub
   
ErrorHandler:
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   MsgBox "(" & Err.Number & ")" & Err.Description
End Sub

'Add By Sindy 2020/10/26
Private Sub GetTextNo_2(strLP44 As String, strCP10 As String, _
   ByRef strMailSubject As String, ByRef StrMailContent As String)

Dim AdoRs As New ADODB.Recordset
Dim strTM01 As String, strTM02 As String, strTM03 As String, strTM04 As String
Dim strCuName As String
   
   If Left(strLP44, 1) = "," Then
      strLP44 = Mid(strLP44, 2)
   End If
   strLP44 = "'" & Replace(strLP44, ",", "','") & "'"
   strSql = "SELECT trademark.*,NVL(cu104,cu04) as CUName from caseprogress,trademark,Customer" & _
            " where cp01='T' and cp10='" & strCP10 & "' and cp27=" & strSrvDate(1) & " and cp57 is null" & _
            " and cp09 in(" & strLP44 & ")" & _
            " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
            " and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+)" & _
            " order by tm01,tm02,tm03,tm04"
   intI = 1
   Set AdoRs = ClsLawReadRstMsg(intI, strSql)
   If intI = 0 Then Exit Sub
   
   m_Appl = "": m_CaseNo = ""
   m_TM45 = "": m_TM05 = "": m_TM09 = "": m_TM15 = ""
   strCuName = ""
   AdoRs.MoveFirst
   Do While Not AdoRs.EOF
      strTM01 = AdoRs.Fields("tm01")
      strTM02 = AdoRs.Fields("tm02")
      strTM03 = AdoRs.Fields("tm03")
      strTM04 = AdoRs.Fields("tm04")
      If strTM03 & strTM04 = "000" Then
         m_CaseNo = m_CaseNo & "," & strTM01 & "-" & strTM02
      Else
         m_CaseNo = m_CaseNo & "," & strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04
      End If
      If Trim("" & AdoRs.Fields("tm45")) <> "" Then
         m_TM45 = m_TM45 & "," & ChgSQL("" & AdoRs.Fields("tm45"))
      End If
      If Trim("" & AdoRs.Fields("tm05")) <> "" Then
         m_TM05 = m_TM05 & "," & ChgSQL("" & AdoRs.Fields("tm05"))
      End If
      If Trim("" & AdoRs.Fields("tm09")) <> "" Then
         m_TM09 = m_TM09 & ",第" & "" & AdoRs.Fields("tm09") & "類"
      End If
      If Trim("" & AdoRs.Fields("tm15")) <> "" Then
         m_TM15 = m_TM15 & "," & "" & AdoRs.Fields("tm15")
      End If
      m_Appl = m_Appl & "和第" & AdoRs.Fields("tm15") & "號「" & ChgSQL(AdoRs.Fields("tm05")) & "」"
      strCuName = "" & AdoRs.Fields("CUName")
      AdoRs.MoveNext
   Loop
   
   If m_TM45 <> "" Then m_TM45 = Right(m_TM45, Len(m_TM45) - 1)
   If m_TM05 <> "" Then m_TM05 = Right(m_TM05, Len(m_TM05) - 1)
   If m_TM09 <> "" Then m_TM09 = Right(m_TM09, Len(m_TM09) - 1)
   If m_TM15 <> "" Then m_TM15 = Right(m_TM15, Len(m_TM15) - 1)
   
   m_Appl = Right(m_Appl, Len(m_Appl) - 1)
   m_CaseNo = Right(m_CaseNo, Len(m_CaseNo) - 1)

   '多筆件時,增加回傳主旨和內文,儲存於信函進度裡,方便使用者寄信使用
   If InStr(m_CaseNo, ",") > 0 Then
      strMailSubject = m_CaseNo
      StrMailContent = "貴方卷號：" & m_TM45 & vbCrLf & _
      "我方案號：" & m_CaseNo & vbCrLf & vbCrLf & _
      "申請人：" & strCuName & vbCrLf & _
      "商標：" & m_TM05 & vbCrLf & _
      "類別：" & m_TM09 & vbCrLf & _
      "註冊號：" & m_TM15 & vbCrLf
   Else
      strMailSubject = ""
      StrMailContent = ""
   End If
End Sub

'Add By Sindy 2013/11/5 組多個案號及號數
'Modify By Sindy 2020/2/21 + , strTM44 As String, m_CP09 As String, _
   m_TM01 As String, m_TM02 As String, m_TM03 As String, m_TM04 As String, strCuName As String, _
   ByRef strMailSubject As String, ByRef StrMailContent As String
'Modify By Sindy 2020/7/30 + , strIsClose As String
Private Sub GetTextNo(AdoRs As ADODB.Recordset, strTM23 As String, strTM44 As String, m_CP09 As String, _
   m_TM01 As String, m_TM02 As String, m_TM03 As String, m_TM04 As String, strCuName As String, _
   ByRef strMailSubject As String, ByRef StrMailContent As String, strIsClose As String)

Dim strTM01 As String, strTM02 As String, strTM03 As String, strTM04 As String
Dim bolOK As Boolean
Dim strRecvNo As String
   
   m_Appl = "": m_CaseNo = "": strRecvNo = ""
   m_TM45 = "": m_TM05 = "": m_TM09 = "": m_TM15 = "" 'Add By Sindy 2020/2/6
   AdoRs.MoveFirst
   Do While Not AdoRs.EOF
      'Modify By Sindy 2020/2/21
      'If strTM23 = AdoRs.Fields("tm23") Then
      If (Left(AdoRs.Fields("tm23"), 1) = "X" And strTM23 = AdoRs.Fields("tm23")) Or _
         (Left(AdoRs.Fields("tm23"), 1) = "Y" And strTM23 = AdoRs.Fields("T_tm23") And strTM44 = AdoRs.Fields("T_tm44")) Then
      '2020/2/21 END
         
         'Add By Sindy 2020/7/30 增加檢查結案狀況
         bolOK = True
         'Modify By Sindy 2025/7/17 mark
'         If AdoRs.Fields("cp10") = "1717" Then '通知延展
         '2025/7/17 END
            If AdoRs.Fields("isclose") <> strIsClose Then
               bolOK = False
            End If
'         End If
         If bolOK = True Then
         '2020/7/30 END
            strTM01 = AdoRs.Fields("tm01")
            strTM02 = AdoRs.Fields("tm02")
            strTM03 = AdoRs.Fields("tm03")
            strTM04 = AdoRs.Fields("tm04")
            If strTM03 & strTM04 = "000" Then
               m_CaseNo = m_CaseNo & "," & strTM01 & "-" & strTM02
            Else
               m_CaseNo = m_CaseNo & "," & strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04
            End If
            
            strRecvNo = strRecvNo & "," & AdoRs.Fields("cp09") 'Add By Sindy 2020/10/26
            
            'Add By Sindy 2020/2/6
            If Trim("" & AdoRs.Fields("tm45")) <> "" Then
               m_TM45 = m_TM45 & "," & ChgSQL("" & AdoRs.Fields("tm45"))
            End If
            If Trim("" & AdoRs.Fields("tm05")) <> "" Then
               m_TM05 = m_TM05 & "," & ChgSQL("" & AdoRs.Fields("tm05"))
            End If
            If Trim("" & AdoRs.Fields("tm09")) <> "" Then
               m_TM09 = m_TM09 & ",第" & "" & AdoRs.Fields("tm09") & "類"
            End If
            If Trim("" & AdoRs.Fields("tm15")) <> "" Then
               m_TM15 = m_TM15 & "," & "" & AdoRs.Fields("tm15")
            End If
            '2020/2/6 END
            'Modify By Sindy 2013/12/5 +案件名稱
            'm_Appl = m_Appl & "和第" & adoRs.Fields("tm15") & "號"
            m_Appl = m_Appl & "和第" & AdoRs.Fields("tm15") & "號「" & ChgSQL(AdoRs.Fields("tm05")) & "」"
         End If
      End If
      AdoRs.MoveNext
   Loop
   'Add By Sindy 2020/2/6
   If m_TM45 <> "" Then m_TM45 = Right(m_TM45, Len(m_TM45) - 1)
   If m_TM05 <> "" Then m_TM05 = Right(m_TM05, Len(m_TM05) - 1)
   If m_TM09 <> "" Then m_TM09 = Right(m_TM09, Len(m_TM09) - 1)
   If m_TM15 <> "" Then m_TM15 = Right(m_TM15, Len(m_TM15) - 1)
   '2020/2/6 END
   m_Appl = Right(m_Appl, Len(m_Appl) - 1)
   m_CaseNo = Right(m_CaseNo, Len(m_CaseNo) - 1)
   strRecvNo = Right(strRecvNo, Len(strRecvNo) - 1)
   
   'Add By Sindy 2020/2/21 當多筆件時,增加回傳主旨和內文,儲存於信函進度裡,方便使用者寄信使用
   If InStr(m_CaseNo, ",") > 0 Then
      'strMailSubject = m_CaseNo
      strMailSubject = strRecvNo 'Add By Sindy 2020/10/26
      StrMailContent = "貴方卷號：" & m_TM45 & vbCrLf & _
      "我方案號：" & m_CaseNo & vbCrLf & vbCrLf & _
      "申請人：" & strCuName & vbCrLf & _
      "商標：" & m_TM05 & vbCrLf & _
      "類別：" & m_TM09 & vbCrLf & _
      "註冊號：" & m_TM15 & vbCrLf
   Else
      strMailSubject = ""
      StrMailContent = ""
   End If
   
   'Modify By Sindy 2020/2/21 + & ":" & strTM44 & ":" & m_CP09 & ":" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   'Modify By Sindy 2020/7/30 + & ":" & strIsClose
   List1.AddItem strTM23 & ":" & strIsClose & ":" & strTM44 & ":" & m_CP09 & ":" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
End Sub

'Modify By Sindy 2019/11/1
'產生PDF清單,直接寄給智權人員
Sub StrMenu()
Dim stFileName As String
Dim dblFCnt As Double
Dim strTo As String
Dim strProName As String 'Add By Sindy 2019/12/19
Dim strChkST01 As String 'Add By Sindy 2024/12/27
   
On Error GoTo RunErr 'Add By Sindy 2024/11/1
   
   'Add By Sindy 2019/11/1
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   Else
'      ChDir App.path 'Add By Sindy 2020/2/18 釋放資料夾權限
'      If Dir(m_AttachPath & "\.") <> "" Then
'         Kill m_AttachPath & "\*.*"
'         DoEvents
'      End If
      Call PUB_KillTempFile(Mid(Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum, 2) & "\*.*") '去掉\
   End If
   '2019/11/1 END
   'Add By Sindy 2019/12/19
   If radio(0).Value = True Then
      strProName = "智慧局繳納註冊費通知"
   Else
      strProName = "智慧局延展通知"
   End If
   '2019/12/19 END
   
   GetPleft
   'Modify By Sindy 2019/11/1 + ,r10
   strSql = "select nvl(st02,r10),r01,r02,r03,r04,r05,r06,r07,r08,r09,tm29,r10 from r020311,staff,trademark where r10=st01(+) " & _
            "and id='" & strUserNum & "' and r01=tm01(+) and r02=tm02(+) and r03=tm03(+) and r04=tm04(+) order by ST15,r10,r11,r01,r02,r03,r04 "
   CheckOC
   Page = 0
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      With adoRecordset
         .MoveFirst
         m_strSales = ""
         'Add By Sindy 2019/11/1
         '產生PDF
         'Load frmPDF
         frmPDF.Show
         stFileName = CheckStr(.Fields(11)) & "-" & strProName & strSrvDate(2) '& "-" & CheckStr(.Fields(0))
         strChkST01 = "," & CheckStr(.Fields(11))  'Add By Sindy 2024/12/27
         frmPDF.StartProcess m_AttachPath, stFileName
         '2019/11/1 END
         Printer.PaperSize = 9
         Do While .EOF = False
            For i = 0 To 11 '10
               strTemp(i) = CheckStr(.Fields(i))
            Next i
            If m_strSales <> strTemp(0) Then
               'Add By Sindy 2019/11/1
               If m_strSales <> "" Then
                  Printer.EndDoc
                  frmPDF.EndtProcess
                  Unload frmPDF
                  '產生PDF
                  'Load frmPDF
                  frmPDF.Show
                  stFileName = CheckStr(.Fields(11)) & "-" & strProName & strSrvDate(2) '& "-" & strTemp(0)
                  strChkST01 = strChkST01 & "," & CheckStr(.Fields(11)) 'Add By Sindy 2024/12/27
                  frmPDF.StartProcess m_AttachPath, stFileName
                  Page = 0
               End If
               '2019/11/1 END
               m_strSales = strTemp(0)
               Page = Page + 1
               If Page <> 1 Then
                 Printer.NewPage
               End If
               PrintTitle
            End If
            strTemp(5) = StrToStr(strTemp(5), 7.5)
            strTemp(6) = StrToStr(strTemp(6), 20)
            strTemp(7) = StrToStr(strTemp(7), 7)
            strTemp(8) = StrToStr(strTemp(8), 12)
            strTemp(9) = StrToStr(strTemp(9), 8)
            PrintDatil
            If iPrint >= 10000 Then
               Page = Page + 1
               Printer.NewPage
               PrintTitle
            End If
            .MoveNext
         Loop
      End With
   Else
      CheckOC
      Exit Sub
   End If
   CheckOC
   Printer.EndDoc
   'Add By Sindy 2019/11/1
   frmPDF.EndtProcess
   Unload frmPDF
   '寄Mail
   File1.path = m_AttachPath & "\"
   File1.Refresh
   If File1.ListCount > 0 Then
      For dblFCnt = 0 To File1.ListCount - 1
         If InStr(UCase(Trim(File1.List(dblFCnt))), strProName) > 0 And _
            UCase(Right(Trim(File1.List(dblFCnt)), 4)) = ".PDF" Then
            strTo = Left(Trim(File1.List(dblFCnt)), InStr(Trim(File1.List(dblFCnt)), "-") - 1)
            strChkST01 = Replace(strChkST01, "," & strTo, "")  'Add By Sindy 2024/12/27
            PUB_SendMail strUserNum, strTo, "", strProName, "請參附件！", , m_AttachPath & "\" & File1.List(dblFCnt), , , , , , , , True
            Kill m_AttachPath & "\" & File1.List(dblFCnt)
         End If
      Next dblFCnt
   End If
   '2019/11/1 END
   
'Add By Sindy 2024/11/1
'   'Add By Sindy 2024/12/27
'   If strChkST01 <> "" Then
'      If CheckIsPersonRest("97038", strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = False Then
'         PUB_SendMail strUserNum, "97038", "", "[智慧局註冊費通知函列印-偵查用(StrMenu)]", _
'            m_AttachPath & "\" & stFileName & ".pdf " & vbCrLf & vbCrLf & strChkST01 & " 無檔案可寄送!!", , , , , , , , , , True, False
'      End If
'   Else
'   '2024/12/27 END
'      If CheckIsPersonRest("97038", strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = False Then
'         PUB_SendMail strUserNum, "97038", "", "[智慧局註冊費通知函列印-偵查用(StrMenu)]", _
'            m_AttachPath & "\" & stFileName & ".pdf " & vbCrLf & vbCrLf & "strChkST01=''; 寄送完畢!!", , , , , , , , , , True, False
'      End If
'   End If

   Exit Sub
RunErr:
   If Err.Number <> 0 Then
'      If CheckIsPersonRest("97038", strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = False Then
'         PUB_SendMail strUserNum, "97038", "", "[智慧局註冊費通知函列印-偵查用(StrMenu)]", _
'            m_AttachPath & "\" & stFileName & ".pdf" & vbCrLf & Err.Number & Err.Description, , , , , , , , , , True, False
'      End If
      If InStr(Err.Description, "無法預期的錯誤") > 0 Then
         If Dir(m_AttachPath & "\" & stFileName & ".pdf") <> "" Then
            Resume Next
         Else
            MsgBox Err.Description, , MsgText(5) & " (StrMenu)"
         End If
      Else
         MsgBox Err.Description, , MsgText(5) & " (StrMenu)"
      End If
   End If
   '2024/11/1 END
End Sub

Sub PrintTitle()
   GetPleft
   iPrint = 500
   
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 5700
   Printer.CurrentY = iPrint
   If radio(0).Value = True Then
      'Modify By Sindy 2012/12/20
      'Printer.Print "智慧局繳納第一期註冊費通知"
      Printer.Print "智慧局繳納註冊費通知"
      '2012/12/20 End
   Else
      'Modify By Sindy 2019/10/29 Mark
'      If radio(1).Value = True Then
'         Printer.Print "智慧局繳納第二期註冊費通知"
'      Else
         Printer.Print "智　慧　局　延　展　通　知"
'      End If
   End If
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   iPrint = iPrint + 500
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "智權人員：" & m_strSales
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(Page)
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.Font.Size = 10
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "申請案號/審定號"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "商品類別"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "申請人"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "法定期限"
   iPrint = iPrint + 300
   Printer.Font.Size = 12
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.Font.Size = 10
End Sub

Sub PrintDatil()
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   If strTemp(10) = "" Then
      Printer.Print strTemp(1) & "-" & strTemp(2) & "-" & strTemp(3) & "-" & strTemp(4)
   Else
      Printer.Print strTemp(1) & "-" & strTemp(2) & "-" & strTemp(3) & "-" & strTemp(4) & "*"
   End If
   
   For i = 5 To 9
      Printer.CurrentX = PLeft(i - 3)
      Printer.CurrentY = iPrint
      Select Case i
         Case 5  '申請案號/審定號
            Printer.Print strTemp(7)
         Case 7  '商品類別
            Printer.Print strTemp(5)
         Case Else
            Printer.Print strTemp(i)
      End Select
   Next i
   iPrint = iPrint + 300
End Sub

Sub GetPleft()
   Erase PLeft
   PLeft(1) = 500
   PLeft(2) = 2000
   PLeft(3) = 3700
   PLeft(4) = 8980
   PLeft(5) = 10900
   PLeft(6) = 14020
   PLeft(7) = 4660
End Sub

' 列印定稿
Private Sub PrintLetter()
'Add By Sindy 2012/1/16
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/16 End
Dim dblTMKindCnt As Double, varTmp As Variant 'Add By Sindy 2020/1/31
   
   'Add By Sindy 2020/1/31
   '取得商品類別數
   strSql = "SELECT tm09 from trademark " & _
            "where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "'" & _
            " and tm03='" & m_TM03 & "' and tm04='" & m_TM04 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   dblTMKindCnt = 1
   If intI = 1 Then
      If "" & RsTemp.Fields("TM09").Value <> "" Then
         If InStr(RsTemp.Fields("TM09").Value, ",") > 0 Then
            varTmp = Split(RsTemp.Fields("TM09").Value, ",")
            dblTMKindCnt = UBound(varTmp) + 1
         End If
      End If
   End If
   '2020/1/31 END
   
   'Add By Sindy 2012/1/16
   ET01 = "15"
   ET02 = m_CP09
   bolEdit = False
   '2012/1/16 End
   
   '2011/12/26 modify by sonia 改判斷畫面上定稿語文,原用申請人國籍m_strTM23Nation< "010"
   ' 台->各國
   'Modify By Sindy 2013/11/4 + Or (radio(0).Value = True Or radio(2).Value = True)
   'Modify By Sindy 2020/1/31 繳納註冊費 及台->台的通知延展
   'If m_Print = "1" Or (radio(0).Value = True Or radio(1).Value = True) Then
   If (m_Print = "1" And radio(1).Value = True) Or radio(0).Value = True Then
   '2020/1/31 END
      ' 清除定稿例外欄位檔原有資料
      If radio(0).Value = True Then
         'Modify By Sindy 2025/7/17 +ET03 = "01"
         If m_strIsClose = "N" Then
            ET03 = "00" 'Modify By Sindy 2012/1/16
         Else
            ET03 = "01" '台->台通知繳納註冊費(已結案)
         End If
         '2025/7/17 END
      'Add By Sindy 2020/7/30
      ElseIf radio(1).Value = True Then
         If m_strIsClose = "N" Then
            ET03 = "00" '台->台通知延展(未結案)
         Else
            ET03 = "02" '台->台通知延展(已結案)
         End If
      End If
      '2020/7/30 END
      EndLetter ET01, m_CP09, ET03, strUserNum
      
'      NowPrint m_CP09, "15", "00", False, strUserNum, 0, "", False, "", 2
      iCopy = 1 'Modify By Sindy 2013/12/9 原印2份改印1份
      'Add By Sindy 2013/11/4
      'Modify By Sindy 2025/7/17 mark
      'If radio(0).Value = True Or radio(1).Value = True Then
      '2025/7/17 END
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "'," & _
                  "'多個案號','" & m_CaseNo & "')"
         cnnConnection.Execute strSql
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "'," & _
                  "'多筆註冊號數','" & m_Appl & "')"
         cnnConnection.Execute strSql
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "'," & _
                  "'期限年月','" & m_LimitDt & "')"
         cnnConnection.Execute strSql
         'Add By Sindy 2020/7/30
         If InStr(m_CaseNo, ",") = 0 Then
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "'," & _
                  "'單筆案號要印','♀')"
            cnnConnection.Execute strSql
         End If
         '2020/7/30 END
      'End If
      '2013/11/4 END
   '外->台的通知延展
   ElseIf m_Print = "2" Then
      ' 清除定稿例外欄位檔原有資料
      ET03 = "01" 'Modify By Sindy 2012/1/16
      EndLetter ET01, m_CP09, ET03, strUserNum
      'Add By Sindy 2020/1/31
'      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               "VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "'," & _
'               "'費用','" & Format(1800 + (1500 * (dblTMKindCnt - 1)), "##,##0") & "')"
'      cnnConnection.Execute strSql
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "'," & _
               "'多個案號','" & m_CaseNo & "')"
      cnnConnection.Execute strSql
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "'," & _
               "'多筆彼所案號','" & m_TM45 & "')"
      cnnConnection.Execute strSql
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "'," & _
               "'多筆商標名稱','" & m_TM05 & "')"
      cnnConnection.Execute strSql
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "'," & _
               "'多筆類別','" & m_TM09 & "')"
      cnnConnection.Execute strSql
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "'," & _
               "'多筆商標號數','" & m_TM15 & "')"
      cnnConnection.Execute strSql
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "'," & _
               "'多筆註冊號數','" & m_Appl & "')"
      cnnConnection.Execute strSql
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "'," & _
               "'期限年月','" & m_LimitDt & "')"
      cnnConnection.Execute strSql
      '2020/1/31 END
      
      'Modify by Sindy 2021/1/26 Mark,電子化了
'      'Modify By Sindy 2010/11/10
'      If m_CP13 = "96029" Or m_CP13 = "96030" Then
'         iCopy = 3
'      '2010/11/10 End
'      Else
'         iCopy = 2
'      End If
   End If
   
   'Add By Sindy 2012/1/16
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper)
      If bolEmail Then
         'Modify By Sindy 2021/4/29 份數改1
'         '判斷是否EMail同時寄紙本
'         If Not bolPlusPaper Then
            iCopy = 1
'         End If
         'Modify By Sindy 2019/10/18 + 信函收文號
         'NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True, , , , , IIf(strSrvDate(1) >= T商標電子化啟用日, m_CP09, "")
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , IIf(strSrvDate(1) >= T商標電子化啟用日, m_CP09, "")
'         boleFileSave = True
'         MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
      Else
         'Modify By Sindy 2019/10/18 + 信函收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , IIf(strSrvDate(1) >= T商標電子化啟用日, m_CP09, "")
      End If
   End If
   '2012/1/16 End
End Sub
