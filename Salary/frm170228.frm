VERSION 5.00
Begin VB.Form frm170228 
   BorderStyle     =   1  '單線固定
   Caption         =   "其他各類所得資料檢核表"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5070
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   3135
      MaxLength       =   12
      TabIndex        =   2
      Top             =   1410
      Width           =   1300
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1485
      MaxLength       =   12
      TabIndex        =   1
      Top             =   1410
      Width           =   1300
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1485
      MaxLength       =   3
      TabIndex        =   0
      Top             =   960
      Width           =   435
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   20
      TabIndex        =   6
      Top             =   2160
      Width           =   5000
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   3
         Top             =   180
         Width           =   4200
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
      Left            =   2700
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3780
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin VB.Line Line3 
      X1              =   2730
      X2              =   3130
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "所得人代號："
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "所得年度： "
      Height          =   180
      Left            =   555
      TabIndex        =   8
      Top             =   990
      Width           =   945
   End
End
Attribute VB_Name = "frm170228"
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
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 20) As Integer
Dim strTemp(1 To 20) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblAmtS1 As Double, dblAmtS2 As Double
Dim dblAmtT1 As Double, dblAmtT2 As Double
Dim dblCntS As Double, dblCntT As Double

Private Sub cmdok_Click(Index As Integer)

   Select Case Index
      Case 0
         If txt1(0) = "" Then
            MsgBox "所得年度不可空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
         End If
         If RunNick(txt1(1), txt1(2)) Then
            txt1(1).SetFocus
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

   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   
   m_StrSQL = ""
   If txt1(0) <> "" Then '所得年度
      m_StrSQL = m_StrSQL & " and OID01=" & Val(txt1(0)) + 1911
   End If
   If txt1(1) <> "" Then '所得人代號起
      m_StrSQL = m_StrSQL & " and OID02 >='" & Trim(txt1(1)) & "' "
   End If
   If txt1(2) <> "" Then '所得人代號迄
      m_StrSQL = m_StrSQL & " and OID02 <='" & Trim(txt1(2)) & "' "
   End If

   m_str = "SELECT OID03 公司別, OID02 所得人代號,SUBSTR(NVL(OI04,ST02),1,4) 名稱,OID04 格式,OID05||'~'||OID06 起迄月份, " & _
           "NVL(OID08,0) 所得總額,NVL(OID09,0) 扣繳稅額, OID07 共用欄位 " & _
           "FROM OTHERINCOMEDATA,STAFF,OTHERINCOMER WHERE OID02=ST01(+) AND OID02=OI01(+) " & m_StrSQL & _
           " ORDER BY 公司別,格式,所得人代號"
   
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         
         '預設值
         iLine = 1
         strType = "" '切頁條件
         dblAmtS1 = 0: dblAmtS2 = 0: dblCntS = 0
         dblAmtT1 = 0: dblAmtT2 = 0: dblCntT = 0
         
         Do While Not m_rs.EOF
             
            For m_i = 1 To 8
                strTemp(m_i) = ""
            Next m_i
             
            strTemp(1) = CheckStr(m_rs.Fields(0))   '公司別
            strTemp(2) = CheckStr(m_rs.Fields(1))   '所得人代號
            strTemp(3) = CheckStr(m_rs.Fields(2))   '名稱
            strTemp(4) = CheckStr(m_rs.Fields(3))   '格式
            strTemp(5) = CheckStr(m_rs.Fields(4))   '起迄月份
            strTemp(6) = CheckStr(m_rs.Fields(5))   '所得總額
            strTemp(7) = CheckStr(m_rs.Fields(6))   '扣繳稅額
            strTemp(8) = CheckStr(m_rs.Fields(7))   '共用欄位
            
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
            dblCntS = dblCntS + 1
            dblAmtS1 = dblAmtS1 + strTemp(6)
            dblAmtS2 = dblAmtS2 + strTemp(7)
            dblCntT = dblCntT + 1
            dblAmtT1 = dblAmtT1 + strTemp(6)
            dblAmtT2 = dblAmtT2 + strTemp(7)
               
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
   Printer.Print String(130, "-")

   iLine = iLine + 1
   Printer.CurrentX = 3000
   Printer.CurrentY = iLine * 300
   Printer.Print "小　計："
   Printer.CurrentX = 5000 - Printer.TextWidth(dblCntS & "人")
   Printer.CurrentY = iLine * 300
   Printer.Print dblCntS & "人"
   Printer.CurrentX = PLeft(1) - Printer.TextWidth(Format(dblAmtS1, "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtS1, "#,###,###,###")
   Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(dblAmtS2, "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtS2, "#,###,###,###")

   dblAmtS1 = 0: dblAmtS2 = 0: dblCntS = 0

End Sub

Sub PrintTotal()
   
   iLine = iLine + 2
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(130, "-")

   iLine = iLine + 1
   Printer.CurrentX = 3000
   Printer.CurrentY = iLine * 300
   Printer.Print "合　計："
   Printer.CurrentX = 5000 - Printer.TextWidth(dblCntT & "人")
   Printer.CurrentY = iLine * 300
   Printer.Print dblCntT & "人"
   Printer.CurrentX = PLeft(1) - Printer.TextWidth(Format(dblAmtT1, "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtT1, "#,###,###,###")
   Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(dblAmtT2, "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtT2, "#,###,###,###")
   
   dblAmtT1 = 0: dblAmtT2 = 0: dblCntT = 0
   
End Sub

Sub PrintTitle()

   GetPleft
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("其他各類所得資料檢核表") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "其他各類所得資料檢核表"
   
   iLine = iLine + 1
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 5100
   Printer.CurrentY = iLine * 300
   Printer.Print "所得年度：" & Val(txt1(0))
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "頁　　次：" & Printer.Page
   
   iLine = iLine + 2
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "公司別：" & strTemp(1) & "　" & CompNameQuery(strTemp(1))

   iLine = iLine + 2
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "所得人代號"
   Printer.CurrentX = 2000
   Printer.CurrentY = iLine * 300
   Printer.Print "名　稱"
   Printer.CurrentX = 3150
   Printer.CurrentY = iLine * 300
   Printer.Print "格式"
   Printer.CurrentX = 4000
   Printer.CurrentY = iLine * 300
   Printer.Print "起迄月份"
   Printer.CurrentX = PLeft(1) - Printer.TextWidth("所得總額")
   Printer.CurrentY = iLine * 300
   Printer.Print "所得總額"
   Printer.CurrentX = PLeft(2) - Printer.TextWidth("扣繳稅額")
   Printer.CurrentY = iLine * 300
   Printer.Print "扣繳稅額"
   Printer.CurrentX = 8500
   Printer.CurrentY = iLine * 300
   Printer.Print "共用欄位"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(130, "-")
   
   iLine = iLine + 1
End Sub

Sub GetPleft()
   PLeft(1) = 6500
   PLeft(2) = 8000
End Sub

Sub PrintDetail()
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)
   Printer.CurrentX = 2000
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(3)
   Printer.CurrentX = 3300
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(4)
   Printer.CurrentX = 4300
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(5)
   Printer.CurrentX = PLeft(1) - Printer.TextWidth(Format(strTemp(6), "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(6), "#,###,###,###")
   Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(strTemp(7), "#,###,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(7), "#,###,###,###")
   Printer.CurrentX = 8500
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(8)
   
   iLine = iLine + 1
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSQL As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   strSystemKind = GetSystemKindByNick
   strSQL = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      Combo1.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = strSQL Then
         SeekPrint = i
      End If
   Next i
   
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170228 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 1, 2
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
      Case 1, 2
         If Index = 1 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 2 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Cancel = True
            End If
         End If
      Case Else
   End Select
   
   If Cancel = True Then TextInverse txt1(Index)
      
End Sub


