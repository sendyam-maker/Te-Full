VERSION 5.00
Begin VB.Form frm170205 
   BorderStyle     =   1  '單線固定
   Caption         =   "地紙條列印"
   ClientHeight    =   3468
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4332
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3468
   ScaleWidth      =   4332
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "Y"
      Top             =   1860
      Width           =   270
   End
   Begin VB.TextBox txtSet 
      Height          =   270
      Index           =   2
      Left            =   1530
      TabIndex        =   14
      Top             =   2850
      Width           =   705
   End
   Begin VB.TextBox txtSet 
      Height          =   270
      Index           =   1
      Left            =   1530
      TabIndex        =   13
      Top             =   2580
      Width           =   705
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   855
      Style           =   2  '單純下拉式
      TabIndex        =   12
      Top             =   2160
      Width           =   3150
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   855
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "1"
      Top             =   1530
      Width           =   270
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "Y"
      Top             =   1200
      Width           =   270
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   2925
      MaxLength       =   12
      TabIndex        =   1
      Text            =   "123456789012"
      Top             =   570
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1485
      MaxLength       =   12
      TabIndex        =   0
      Text            =   "123456789012"
      Top             =   570
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "96"
      Top             =   900
      Width           =   435
   End
   Begin VB.CommandButton cmdeXit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   3105
      TabIndex        =   7
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Left            =   2025
      TabIndex        =   6
      Top             =   90
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否含其他所得人：        ( Y: 是 )"
      Height          =   180
      Index           =   4
      Left            =   135
      TabIndex        =   18
      Top             =   1875
      Width           =   2580
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "縱軸偏移值(Y)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   5
      Left            =   135
      TabIndex        =   17
      Top             =   2910
      Width           =   3240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "橫軸偏移值(X)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   4
      Left            =   135
      TabIndex        =   16
      Top             =   2610
      Width           =   3240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      Height          =   255
      Index           =   3
      Left            =   135
      TabIndex        =   15
      Top             =   2250
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "眷屬別：       (  1: 父親, 2: 母親 )"
      Height          =   180
      Index           =   3
      Left            =   135
      TabIndex        =   11
      Top             =   1575
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "只印外翻人員：         ( Y: 是 )"
      Height          =   180
      Index           =   2
      Left            =   135
      TabIndex        =   10
      Top             =   1245
      Width           =   2265
   End
   Begin VB.Line Line2 
      X1              =   2310
      X2              =   2970
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "列印對象代碼："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   9
      Top             =   600
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "離職年度："
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   8
      Top             =   930
      Width           =   900
   End
End
Attribute VB_Name = "frm170205"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Modified by Morgan 2024/1/31 改新部門
'Create by Morgan 2009/1/19
Option Explicit

Dim m_Actived As Boolean
Dim m_DefaultPrinter As String
Dim m_PageNo As Integer
Public m_bolBeCalled As Boolean
Public m_bolRegAddr As Boolean 'Add by Morgan 2011/2/17 是否印戶籍地址
Public m_bolVPrint As Boolean 'Add by Morgan 2011/2/17 是否直印

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   Screen.MousePointer = vbHourglass
   If TxtValidate = True Then
      Me.Enabled = False
      If cmbPrinter <> Printer.DeviceName Then
         PUB_RestorePrinter cmbPrinter
      End If
      m_PageNo = 0
      PrintSheet
      '若印表機變動, 則更新列印設定
      If cmbPrinter.Tag <> cmbPrinter Or txtSet(1) <> txtSet(1).Tag Or txtSet(2) <> txtSet(2).Tag Then
          PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, Val(txtSet(1)), Val(txtSet(2)), Me.cmbPrinter.Text
      End If
      If Printer.DeviceName <> m_DefaultPrinter Then
         PUB_RestorePrinter m_DefaultPrinter
      End If
      Me.Enabled = True
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
'   If m_Actived = False Then
'      FormReset
'      Text1(0).SetFocus
'      Text1_GotFocus 0
'      m_Actived = True
'   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter, m_DefaultPrinter, , , Me.txtSet(1), Me.txtSet(2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170205 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
   Case 2 '離職年度
      If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
         KeyAscii = 0
         Beep
      End If
   
   Case 3, 5 '外翻,其他所得人
      If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
         KeyAscii = 0
         Beep
      End If
   Case 4 '眷屬別
      If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
         KeyAscii = 0
         Beep
      End If
      
   Case Else
   
   End Select
End Sub

Public Sub FormReset()
   Dim oText As TextBox
   For Each oText In Text1
      oText.Text = ""
   Next
End Sub

Private Function TxtValidate() As Boolean
   TxtValidate = True
   If Trim(Text1(2)) & Trim(Text1(3)) & Trim(Text1(4)) = "" Then
      If Trim(Text1(0)) = "" Then
         TxtValidate = False
         MsgBox "請輸入員工編號起!"
         Exit Function
      End If
      If Text1(1) = "" Then
         TxtValidate = False
         MsgBox "請輸入員工編號迄!"
         Exit Function
      End If
   End If
   
End Function

Public Sub PrintSheet()
   Dim stCon As String, stCon1 As String
   Dim strFontSize As String, strFontName As String
   Dim Xo As Integer, Yo As Integer, xi As Long, yi As Long
   Dim lLineHeight As Long
      
   stCon = ""
   '所得人代碼
   If Text1(0) <> "" Then
      stCon = stCon & " and st01>='" & Text1(0) & "'"
      stCon1 = stCon1 & " and oi01>='" & Text1(0) & "'"
   End If
   If Text1(1) <> "" Then
      stCon = stCon & " and st01<='" & Text1(1) & "'"
      stCon1 = stCon1 & " and oi01<='" & Text1(1) & "'"
   End If
   '離職年度
   If Text1(2) <> "" Then
      stCon = stCon & " and substr(st51,1,4)=" & (Val(Text1(2)) + 1911)
   '非被呼叫狀態時只抓在職員工
   ElseIf m_bolBeCalled = False Then
      stCon = stCon & " and st04='1'"
   End If
   '外翻人員
   If Text1(3) <> "" Then
      stCon = stCon & " and st03='F51'"
   End If
      
   '父親、母親(排除已歿)
   If Text1(4) <> "" Then
      'Modify by Morgan 2009/7/7 sr12 改放 'Y'
      strExc(0) = "select nvl(st93,st03) st03,st01,sr04 name,sr10 zipc,sr11 addr from staff,Staff_Relation" & _
         " where sr01(+)=st01 and sr03='" & Text1(4) & "' and sr13 is null" & stCon
   Else
      'Modify by Morgan 2011/2/17
      'strExc(0) = "select st03,st01,st02 name,st33 zipc,st08 addr from staff where 1=1" & stCon
      If m_bolRegAddr Then
         strExc(0) = "select nvl(st93,st03),st01,st02 name,st36 zipc,st34 addr from staff where 1=1" & stCon
      Else
         strExc(0) = "select nvl(st93,st03),st01,st02 name,st33 zipc,st08 addr from staff where 1=1" & stCon
      End If
      
      '其他所得人
      If Text1(5) = "Y" Then
          strExc(0) = strExc(0) & _
            " union select '',oi01,oi04 name,'' zipc,oi05 addr from OtherIncomer where 1=1" & stCon1
      End If
   End If
   
   '依照部門排序
   strExc(0) = strExc(0) & " order by 1,2,3"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      MsgBox "無可列印資料!"
   ElseIf intI = 1 Then
   
      strFontSize = Printer.FontSize
      strFontName = Printer.FontName
      
      Printer.EndDoc
      Printer.PaperSize = PUB_GetPaperSize(2) '地址條
      
      'Add by Morgan 2011/2/17
      'Printer.Font = "細明體"
      If m_bolVPrint Then
         Printer.Font = "@細明體"
      Else
         Printer.Font = "細明體"
      End If
      
      Printer.Font.Size = 12
      lLineHeight = Printer.TextHeight("址") + 50
      
      Xo = Val(txtSet(1)) * 567
      Yo = Val(txtSet(2)) * 567
      
      With RsTemp
      Do While Not .EOF
         m_PageNo = m_PageNo + 1
         If .AbsolutePosition > 1 Then
            Printer.NewPage
            'Add by Morgan 2011/2/17
            If m_bolVPrint Then
               Printer.Font = "@細明體"
            Else
               Printer.Font = "細明體"
            End If
            'end 2011/2/17
         End If
         '郵遞區號
         yi = Yo + 300
         strExc(1) = "" & .Fields("zipc")
         If strExc(1) <> "" Then
            xi = Xo + 200
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
         End If
         
         '地址
         yi = yi + lLineHeight
         strExc(1) = Trim("" & .Fields("addr"))
         If strExc(1) <> "" Then
            xi = Xo + 200
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Pub_SmartPrint strExc(1), xi, yi, , lLineHeight
         End If
         
         '收件人
         yi = yi + 2 * lLineHeight
         
         Select Case Text1(4)
            Case "1"
               strExc(1) = "" & .Fields("name") & "　　　　　先生　鈞啟"
            Case "2"
               strExc(1) = "" & .Fields("name") & "　　　　　女士　鈞啟"
            Case Else
               strExc(1) = "" & .Fields("name") & "　　　　　君　鈞啟"
         End Select
         If strExc(1) <> "" Then
            xi = Xo + 200
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Pub_SmartPrint strExc(1), xi, yi, , lLineHeight
         End If
                  
         '員工編號+頁次
         Printer.Font = "細明體"
         
         yi = yi + 2 * lLineHeight
         xi = Xo + 200
         strExc(1) = "" & .Fields("st01")
         strExc(1) = strExc(1) & String(36 - GetTextLength(strExc(1)) - 6, " ") & Format(m_PageNo, String(6, "0"))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         .MoveNext
      Loop
      End With
      
      Printer.EndDoc
      Printer.FontSize = strFontSize
      Printer.FontName = strFontName
      
      'Add By Sindy 2022/3/11 改用Execl列印地址條
      Dim strTempAddressList As String
      If strTempAddressList <> "" Then
         strTempAddressList = "７１７臺南市仁德區保安里民生路１５-１號$聯豐生物科技股份有限公司|"
         If PUB_XlsAccAddress(strTempAddressList) = False Then
             MsgBox "列印失敗！", vbCritical
         End If
      End If
      'end 2022/03/01
      
      If m_bolBeCalled = False Then
         MsgBox "列印結束 !"
      End If
   End If
End Sub


