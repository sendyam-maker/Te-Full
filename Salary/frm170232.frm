VERSION 5.00
Begin VB.Form frm170232 
   BorderStyle     =   1  '單線固定
   Caption         =   "個人勞退自提明細表"
   ClientHeight    =   2544
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2544
   ScaleWidth      =   4740
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   2820
      MaxLength       =   5
      TabIndex        =   2
      Top             =   1080
      Width           =   780
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   1830
      MaxLength       =   5
      TabIndex        =   1
      Top             =   1080
      Width           =   780
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1830
      MaxLength       =   6
      TabIndex        =   0
      Top             =   735
      Width           =   780
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   3
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   7
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2580
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3660
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin VB.Label lblDsp 
      AutoSize        =   -1  'True
      Caption         =   "台一員工"
      Height          =   180
      Index           =   1
      Left            =   2760
      TabIndex        =   10
      Top             =   780
      Width           =   720
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2550
      X2              =   2865
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "薪資年月："
      Height          =   180
      Index           =   2
      Left            =   840
      TabIndex        =   9
      Top             =   1140
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   840
      TabIndex        =   8
      Top             =   780
      Width           =   900
   End
End
Attribute VB_Name = "frm170232"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2009/7/6 add by sonia
Option Explicit
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblAmt1 As Double, dblAmt2 As Double        '小計

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If txt1(1) = "" Then
            If MsgBox("是否確定要列印所有員工資料!!", vbYesNo + vbDefaultButton2) = vbNo Then
               txt1(1).SetFocus
               Exit Sub
            End If
         End If
'2011/7/28 CANCEL BY SONIA 辜說不必限制
'         If Len(txt1(3)) = 5 Then
'            If Left(Trim(txt1(3)), 3) <> Left(Trim(txt1(4)), 3) Then
'               MsgBox "起迄薪資年度不同！", vbInformation, "操作錯誤！"
'               txt1(3).SetFocus
'               Exit Sub
'            End If
'         Else
'            If Left(Trim(txt1(3)), 2) <> Left(Trim(txt1(4)), 2) Then
'               MsgBox "起迄薪資年度不同！", vbInformation, "操作錯誤！"
'               txt1(3).SetFocus
'               Exit Sub
'            End If
'         End If
'2011/7/28 END
         If RunNick(Val(txt1(3)), Val(txt1(4))) Then
            txt1(3).SetFocus
            Exit Sub
         End If
         
         Screen.MousePointer = vbHourglass
         StrMenu
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

Sub StrMenu()
Dim strYM As String

   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   
   m_StrSQL = ""
   
   If txt1(1) <> "" Then '員工編號起
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'm_StrSQL = m_StrSQL & " and replace(SM01,'A','0') ='" & Trim(txt1(1)) & "' "
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      m_StrSQL = m_StrSQL & " and substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) ='" & Trim(txt1(1)) & "' "
   End If
   If txt1(3) <> "" Then '薪資年月起
      m_StrSQL = m_StrSQL & " and SM02 >=" & Val(txt1(3)) + 191100
   End If
   If txt1(4) <> "" Then '薪資年月起
      m_StrSQL = m_StrSQL & " and SM02 <=" & Val(txt1(4)) + 191100
   End If
   
   '每月薪資資料
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   m_str = "SELECT sm37 公司別,a0802 公司名稱,SM01||' '||SUBSTR(ST02,1,3) 姓名,st26 身份證字號, SM02-191100 薪資年月," & _
           "TO_CHAR(NVL(SM16,0),'9G999G999G999') 勞退自提,TO_CHAR(NVL(SM30,0),'9G999G999G999') 勞退公司提撥 " & _
           "FROM SALARYMONTH,STAFF,acc080 " & _
           "WHERE substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=ST01(+) " & m_StrSQL & "" & _
           "and nvl(sm16,0)+nvl(sm30,0)>0 and sm37=a0801(+) order by sm37,sm01,sm02"
   
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         
         iLine = 1
         strType = "" '切頁條件
         dblAmt1 = 0: dblAmt2 = 0
         
         Do While Not m_rs.EOF
             
            For m_i = 1 To 10
               strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields(0))  '公司別
            strTemp(2) = CheckStr(m_rs.Fields(1))  '公司名稱
            strTemp(3) = CheckStr(m_rs.Fields(2))  '員工編號+姓名
            strTemp(4) = CheckStr(m_rs.Fields(3))  '身分證字號
            strTemp(5) = CheckStr(m_rs.Fields(4))  '薪資年月
            strTemp(6) = CheckStr(m_rs.Fields(5))  '勞退自提
            strTemp(7) = CheckStr(m_rs.Fields(6))  '勞退公司提撥
            
            If iLine > 50 Or iLine = 1 Or strType <> strTemp(1) + strTemp(3) Then
                     
               If (strType <> "" And strType <> strTemp(1) + strTemp(3)) Then
                  PrintEnd '小計
               End If
               
               If strType <> "" Then Printer.NewPage
               iLine = 1
               PrintTitle '列印表頭
            End If
            
            PrintDetail '列印表中
            
            strType = strTemp(1) + strTemp(3) '依公司別+員工編號跳頁
            
            dblAmt1 = dblAmt1 + strTemp(6)        '小計
            dblAmt2 = dblAmt2 + strTemp(7)        '小計
            m_rs.MoveNext
         Loop
          
         '列印表尾
         PrintEnd '小計
         
      End With
   Else
      MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub PrintTitle()
   GetPleft
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("個人勞退自提明細表") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "個人勞退自提明細表"
   
   iLine = iLine + 1
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "頁　　次：" & Printer.Page
   
   iLine = iLine + 2
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTemp(2)) / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)  '公司名稱
   
   iLine = iLine + 2
   Printer.CurrentX = 3000
   Printer.CurrentY = iLine * 300
   Printer.Print "姓名：" & strTemp(3)
   Printer.CurrentX = 6500
   Printer.CurrentY = iLine * 300
   Printer.Print "身份證字號：" & strTemp(4)
   
   iLine = iLine + 2
   Printer.CurrentX = PLeft(1) - Printer.TextWidth("薪資年月")
   Printer.CurrentY = iLine * 300
   Printer.Print "薪資年月"
   Printer.CurrentX = PLeft(2) - Printer.TextWidth("勞退自提")
   Printer.CurrentY = iLine * 300
   Printer.Print "勞退自提"
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("勞退公司提撥")
   Printer.CurrentY = iLine * 300
   Printer.Print "勞退公司提撥"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
End Sub

Sub PrintEnd()
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "合　計："
   Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(dblAmt1, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt1, "##,###,###")
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmt2, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt2, "##,###,###")
   
   dblAmt1 = 0
   dblAmt2 = 0
End Sub

Sub GetPleft()
   PLeft(1) = 4500
   PLeft(2) = 6500
   PLeft(3) = 8500
End Sub

Sub PrintDetail()
   Printer.CurrentX = PLeft(1) - Printer.TextWidth(Format(strTemp(5), "###/##")) - 300
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(5), "###/##")
   Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(strTemp(6), "#,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(6), "#,###,##0")
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(strTemp(7), "#,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(7), "#,###,##0")
   
   iLine = iLine + 1
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   strSystemKind = GetSystemKindByNick
   strSql = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      Combo1.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = strSql Then
         SeekPrint = i
      End If
   Next i
   
   lblDsp(1) = ""
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170232 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 1
         KeyAscii = UpperCase(KeyAscii)
      Case 3, 4
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 1
         If txt1(Index) <> "" Then
            lblDsp(1) = ""
            If ChkStaffID(txt1(Index)) = True Then
               Cancel = True
            End If
            If ClsPDGetStaffN(txt1(Index), strExc(1)) = False Then
               Cancel = True
            Else
               lblDsp(1) = strExc(1)
            End If
         Else
            lblDsp(1) = ""
         End If
   End Select

   If Cancel = True Then TextInverse txt1(Index)
End Sub
