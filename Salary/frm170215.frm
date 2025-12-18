VERSION 5.00
Begin VB.Form frm170215 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工差旅房租技術/證照津貼明細表"
   ClientHeight    =   2952
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5064
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2952
   ScaleWidth      =   5064
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1530
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1780
      Width           =   765
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   0
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   0
      Top             =   975
      Width           =   435
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   1
      Left            =   1530
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1380
      Width           =   300
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   0
      TabIndex        =   8
      Top             =   2280
      Width           =   5000
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   4
         Top             =   180
         Width           =   4200
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   9
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2940
      TabIndex        =   5
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   4020
      TabIndex        =   6
      Top             =   90
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2475
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1780
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "次　　數：          (1：1~4月 2：5~8月 3：9~12月)"
      Height          =   180
      Index           =   2
      Left            =   600
      TabIndex        =   11
      Top             =   1420
      Width           =   3810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "薪資年度："
      Height          =   180
      Index           =   0
      Left            =   600
      TabIndex        =   10
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   600
      TabIndex        =   7
      Top             =   1820
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   1980
      X2              =   2640
      Y1              =   1900
      Y2              =   1900
   End
End
Attribute VB_Name = "frm170215"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2009/1/22 add by sonia
Option Explicit
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQLM As String
Dim m_StrSQLS As String
Dim m_month As String
Dim m_i As Integer
Dim PLeft(1 To 20) As Integer
Dim strTemp(1 To 20) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblAmtS1 As Double, dblAmtS2 As Double, dblAmtS3 As Double, dblAmtS4 As Double, dblAmtS5 As Double, dblAmtS6 As Double
Dim dblAmtT1 As Double, dblAmtT2 As Double, dblAmtT3 As Double, dblAmtT4 As Double, dblAmtT5 As Double, dblAmtT6 As Double

Private Sub cmdok_Click(Index As Integer)

   Select Case Index
      Case 0
         If txt1(0) = "" Then
            MsgBox "薪資年度不可空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
         End If
         If txt1(1) = "" Then
            MsgBox "次數不可空白！", vbInformation, "操作錯誤！"
            txt1(1).SetFocus
            Exit Sub
         End If
         If RunNick(txt1(2), txt1(3)) Then
            txt1(2).SetFocus
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
Dim strSql As String

   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   
   m_StrSQLM = "": m_StrSQLS = ""
   If txt1(0) <> "" Then '薪資－年
      m_StrSQLM = m_StrSQLM & " and SUBSTR(SM02,1,4)='" & Val(txt1(0)) + 1911 & "' "
      'm_StrSQLS = m_StrSQLS & " and SUBSTR(SM02,1,4)='" & Val(txt1(0)) + 1911 & "' "
   End If
   If txt1(1) <> "" Then '次數
      Select Case txt1(1)
         Case 1
            m_StrSQLM = m_StrSQLM & " and SUBSTR(SM02,5,2)>='01' and SUBSTR(SM02,5,2)<='04' "
            'm_StrSQLS = m_StrSQLS & " and SUBSTR(SM02,5,2)>='01' and SUBSTR(SM02,5,2)<='04' "
            m_month = "01-04"
         Case 2
            m_StrSQLM = m_StrSQLM & " and SUBSTR(SM02,5,2)>='05' and SUBSTR(SM02,5,2)<='08' "
            'm_StrSQLS = m_StrSQLS & " and SUBSTR(SM02,5,2)>='05' and SUBSTR(SM02,5,2)<='08' "
            m_month = "05-08"
         Case 3
            m_StrSQLM = m_StrSQLM & " and SUBSTR(SM02,5,2)>='09' and SUBSTR(SM02,5,2)<='12' "
            'm_StrSQLS = m_StrSQLS & " and SUBSTR(SM02,5,2)>='09' and SUBSTR(SM02,5,2)<='12' "
            m_month = "09-12"
      End Select
   End If
   If txt1(2) <> "" Then '員工編號起
      m_StrSQLM = m_StrSQLM & " and SM01 >='" & Trim(txt1(2)) & "' "
      'm_StrSQLS = m_StrSQLS & " and SM01 >='" & Trim(txt1(2)) & "' "
   End If
   If txt1(3) <> "" Then '員工編號迄
      m_StrSQLM = m_StrSQLM & " and SM01 <='" & Trim(txt1(3)) & "' "
      'm_StrSQLS = m_StrSQLS & " and SM01 <='" & Trim(txt1(3)) & "' "
   End If

   '2009/5/21 modify by sonia加技術津貼
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   m_str = "select 公司別,員工編號,ST02 姓名,sum(第一個月+第二個月+第三個月+第四個月),sum(第一個月),sum(第二個月),sum(第三個月),sum(第四個月),sum(房租技術津貼合計) from staff,( " & _
           "SELECT SM37 公司別, SM01 員工編號, DECODE(TO_CHAR(SUBSTR(SM02,5,2))-(" & txt1(1) - 1 & ")*4,1,NVL(SM08,0),0) 第一個月, DECODE(TO_CHAR(SUBSTR(SM02,5,2))-(" & txt1(1) - 1 & ")*4,2,NVL(SM08,0),0) 第二個月, " & _
           "DECODE(TO_CHAR(SUBSTR(SM02,5,2))-(" & txt1(1) - 1 & ")*4,3,NVL(SM08,0),0) 第三個月, DECODE(TO_CHAR(SUBSTR(SM02,5,2))-(" & txt1(1) - 1 & ")*4,4,NVL(SM08,0),0) 第四個月, 0 房租技術津貼合計 FROM SALARYMONTH WHERE NVL(SM08,0)>0 " & m_StrSQLM & _
           " UNION SELECT SM37 公司別, SM01 員工編號, 0 第一個月, 0 第二個月, 0 第三個月, 0 第四個月, SUM(nvl(SM06,0)+nvl(SM09,0)) 房租技術津貼合計 FROM SALARYMONTH WHERE nvl(SM06,0)+NVL(SM09,0)>0 " & m_StrSQLM & " GROUP BY SM37,SM01 " & _
           " ) where substr(員工編號,1,2)||replace(substr(員工編號,3,1),'A','0')||substr(員工編號,4)=ST01(+) group by 公司別,員工編號,ST02 ORDER BY 公司別,員工編號"
   
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         
         '預設值
         iLine = 1
         strType = "" '切頁條件
         dblAmtS1 = 0: dblAmtS2 = 0: dblAmtS3 = 0: dblAmtS4 = 0: dblAmtS5 = 0: dblAmtS6 = 0
         dblAmtT1 = 0: dblAmtT2 = 0: dblAmtT3 = 0: dblAmtT4 = 0: dblAmtT5 = 0: dblAmtT6 = 0
         
         Do While Not m_rs.EOF
             
            For m_i = 1 To 9
                strTemp(m_i) = ""
            Next m_i
             
            strTemp(1) = CheckStr(m_rs.Fields(0))   '公司別
            strTemp(2) = CheckStr(m_rs.Fields(1))   '員工編號
            strTemp(3) = CheckStr(m_rs.Fields(2))   '姓名
            strTemp(4) = CheckStr(m_rs.Fields(4))   '第一個月
            strTemp(5) = CheckStr(m_rs.Fields(5))   '第二個月
            strTemp(6) = CheckStr(m_rs.Fields(6))   '第三個月
            strTemp(7) = CheckStr(m_rs.Fields(7))   '第四個月
            strTemp(8) = CheckStr(m_rs.Fields(8))   '房租技術津貼合計
            '2009/9/10 MODIFY BY SONIA 辜要求差旅津貼合計改為三項津貼合計
            'strTemp(9) = CheckStr(m_rs.Fields(3))   '差旅津貼合計
            strTemp(9) = CheckStr(m_rs.Fields(3) + m_rs.Fields(8)) '三項津貼合計
            
            If iLine > 50 Or iLine = 1 Or strType <> strTemp(1) Then

               If strType <> "" And strType <> strTemp(1) Then
                  PrintEnd '小計
               End If

               If iLine <> 1 Then Printer.NewPage
               iLine = 1
               PrintTitle '列印表頭
            End If
            
            PrintDetail '列印表中
            
            strType = strTemp(1) '依公司別跳頁
            
            '小計及合計
            dblAmtS1 = dblAmtS1 + strTemp(9)
            dblAmtS2 = dblAmtS2 + strTemp(4)
            dblAmtS3 = dblAmtS3 + strTemp(5)
            dblAmtS4 = dblAmtS4 + strTemp(6)
            dblAmtS5 = dblAmtS5 + strTemp(7)
            dblAmtS6 = dblAmtS6 + strTemp(8)
            dblAmtT1 = dblAmtT1 + strTemp(9)
            dblAmtT2 = dblAmtT2 + strTemp(4)
            dblAmtT3 = dblAmtT3 + strTemp(5)
            dblAmtT4 = dblAmtT4 + strTemp(6)
            dblAmtT5 = dblAmtT5 + strTemp(7)
            dblAmtT6 = dblAmtT6 + strTemp(8)
               
            m_rs.MoveNext
         Loop
          
         '列印表尾
         PrintEnd    '小計
         PrintTotal  '合計
           
      End With
   Else
      MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   Printer.EndDoc
   ShowPrintOk

End Sub

Sub PrintEnd()
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(141, "-")

   iLine = iLine + 1
   Printer.CurrentX = 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "小　計："
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmtS1, "###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtS1, "###,###,###")
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblAmtS2, "###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtS2, "###,###,###")
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblAmtS3, "###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtS3, "###,###,###")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(dblAmtS4, "###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtS4, "###,###,###")
   Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(dblAmtS5, "###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtS5, "###,###,###")
   Printer.CurrentX = PLeft(8) - Printer.TextWidth(Format(dblAmtS6, "###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtS6, "###,###,###")

   dblAmtS1 = 0: dblAmtS2 = 0: dblAmtS3 = 0: dblAmtS4 = 0: dblAmtS5 = 0: dblAmtS6 = 0

End Sub

Sub PrintTotal()
   
   iLine = iLine + 2
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(141, "-")

   iLine = iLine + 1
   Printer.CurrentX = 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "合　計："
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmtT1, "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtT1, "#,###,###,###")
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblAmtT2, "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtT2, "#,###,###,###")
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblAmtT3, "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtT3, "#,###,###,###")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(dblAmtT4, "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtT4, "#,###,###,###")
   Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(dblAmtT5, "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtT5, "#,###,###,###")
   Printer.CurrentX = PLeft(8) - Printer.TextWidth(Format(dblAmtT6, "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtT6, "#,###,###,###")
   
   dblAmtT1 = 0: dblAmtT2 = 0: dblAmtT3 = 0: dblAmtT4 = 0: dblAmtT5 = 0: dblAmtT6 = 0
   
End Sub
Sub PrintTitle()

   GetPleft
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("員工差旅房租津貼明細表") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "員工差旅房租技術/證照津貼明細表"
   
   iLine = iLine + 1
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
   Printer.CurrentY = iLine * 300
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 4600
   Printer.CurrentY = iLine * 300
   Printer.Print "薪資年月：" & Val(txt1(0)) & " 年 " & m_month & " 月"
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
   Printer.CurrentY = iLine * 300
   Printer.Print "頁　　次：" & Printer.Page
   
   iLine = iLine + 2
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "公司別：" & strTemp(1) & "　" & CompNameQuery(strTemp(1))

   iLine = iLine + 2
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("本次合計")
   Printer.CurrentY = iLine * 300
   Printer.Print "三項津貼"
   Printer.CurrentX = 4400
   Printer.CurrentY = iLine * 300
   Printer.Print "本　　　次　　　差　　　旅　　　津　　　貼"
   Printer.CurrentX = PLeft(8) - Printer.TextWidth("　本次合計") - 120
   Printer.CurrentY = iLine * 300
   Printer.Print "房租技術/證照津貼"
   
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1) - Printer.TextWidth("編號")
   Printer.CurrentY = iLine * 300
   Printer.Print "編號"
   Printer.CurrentX = PLeft(2) - Printer.TextWidth("姓　名")
   Printer.CurrentY = iLine * 300
   Printer.Print "姓　名"
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("本次合計")
   Printer.CurrentY = iLine * 300
   Printer.Print "本次合計"
   Printer.CurrentX = PLeft(4) - Printer.TextWidth("00 月")
   Printer.CurrentY = iLine * 300
   Printer.Print "  " & Format(1 + (Val(txt1(1)) - 1) * 4, "##") & "月"
   Printer.CurrentX = PLeft(5) - Printer.TextWidth("00 月")
   Printer.CurrentY = iLine * 300
   Printer.Print "  " & Format(2 + (Val(txt1(1)) - 1) * 4, "##") & "月"
   Printer.CurrentX = PLeft(6) - Printer.TextWidth("00 月")
   Printer.CurrentY = iLine * 300
   Printer.Print "  " & Format(3 + (Val(txt1(1)) - 1) * 4, "##") & "月"
   Printer.CurrentX = PLeft(7) - Printer.TextWidth("00 月")
   Printer.CurrentY = iLine * 300
   Printer.Print "  " & Format(4 + (Val(txt1(1)) - 1) * 4, "##") & "月"
   Printer.CurrentX = PLeft(8) - Printer.TextWidth("本次合計")
   Printer.CurrentY = iLine * 300
   Printer.Print "本次合計"
  
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(141, "-")
   
   iLine = iLine + 1
End Sub

Sub GetPleft()
   PLeft(1) = 1000
   PLeft(2) = 2000
   PLeft(3) = 3500
   PLeft(4) = 5000
   PLeft(5) = 6500
   PLeft(6) = 8000
   PLeft(7) = 9500
   PLeft(8) = 11000
End Sub

Sub PrintDetail()
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)
   Printer.CurrentX = PLeft(2) - Printer.TextWidth(strTemp(3))
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(3)
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(strTemp(9), "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(9), "#,###,###,###")
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(strTemp(4), "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(4), "#,###,###,###")
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(5), "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(5), "#,###,###,###")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(6), "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(6), "#,###,###,###")
   Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(strTemp(7), "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(7), "#,###,###,###")
   Printer.CurrentX = PLeft(8) - Printer.TextWidth(Format(strTemp(8), "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(8), "#,###,###,###")
   
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
   
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170215 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 2, 3
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "0101") = False Then
               Cancel = True
            End If
         End If
      Case 1
         If txt1(Index) <> "" Then
            If Val(txt1(Index)) < 1 Or Val(txt1(Index)) > 3 Then
               MsgBox "次數錯誤！", vbInformation, "操作錯誤！"
               Cancel = True
            End If
         End If
      Case 2, 3
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 3 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Cancel = True
            End If
         End If
      Case Else
   End Select
   
   If Cancel = True Then TextInverse txt1(Index)
      
End Sub
