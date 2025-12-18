VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210120 
   BorderStyle     =   1  '單線固定
   Caption         =   "新舊客戶收款貢獻度分析"
   ClientHeight    =   6310
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6310
   ScaleWidth      =   9310
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1320
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7155
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1320
      Width           =   300
   End
   Begin VB.TextBox txtSales 
      Height          =   285
      Left            =   1125
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1020
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea1 
      Height          =   285
      Left            =   2340
      MaxLength       =   3
      TabIndex        =   2
      Top             =   735
      Width           =   915
   End
   Begin VB.TextBox txtZone 
      Height          =   285
      Left            =   1125
      MaxLength       =   1
      TabIndex        =   0
      Top             =   450
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea 
      Height          =   285
      Left            =   1125
      MaxLength       =   3
      TabIndex        =   1
      Top             =   735
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   1245
      Left            =   240
      TabIndex        =   12
      Top             =   4140
      Width           =   3105
      _ExtentX        =   5486
      _ExtentY        =   2205
      _Version        =   393216
      BackColor       =   13820671
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7440
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6540
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   400
      Left            =   5610
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   15
      Width           =   800
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Index           =   0
      Left            =   1125
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1320
      Width           =   915
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Index           =   1
      Left            =   2200
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1320
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd2 
      Height          =   1185
      Left            =   4560
      TabIndex        =   19
      Top             =   60
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8273
      _ExtentY        =   2099
      _Version        =   393216
      BackColor       =   13820671
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   2160
      TabIndex        =   20
      Top             =   1050
      Width           =   1650
      VariousPropertyBits=   27
      Size            =   "2910;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "含結餘及保留：        (Y:含 N.不含) "
      Height          =   180
      Left            =   3285
      TabIndex        =   18
      Top             =   1365
      Width           =   2700
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "列印內容：           (1.明細 2.區合計) "
      Height          =   180
      Left            =   6240
      TabIndex        =   17
      Top             =   1365
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Left            =   90
      TabIndex        =   16
      Top             =   780
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "所別："
      Height          =   180
      Left            =   90
      TabIndex        =   15
      Top             =   495
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   90
      TabIndex        =   14
      Top             =   1065
      Width           =   900
   End
   Begin VB.Label lblZone 
      AutoSize        =   -1  'True
      Caption         =   "（1北所 2中所 3南所 4高所）"
      Height          =   180
      Left            =   2160
      TabIndex        =   13
      Top             =   495
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Line Line2 
      X1              =   2070
      X2              =   2340
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Label Label2 
      Caption         =   "收款日期："
      Height          =   180
      Left            =   90
      TabIndex        =   11
      Top             =   1365
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   1980
      X2              =   2250
      Y1              =   1455
      Y2              =   1455
   End
End
Attribute VB_Name = "frm210120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/05 改成Form2.0 ; Grd1改字型=新細明體-ExtB、Grd2改字型=新細明體-ExtB、lblSalesName ; Printer列印未改
'Memo by Lydia 2019/07/01 表單名稱:智權部新舊客戶收款貢獻度分析=>新舊客戶收款貢獻度分析
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Dim m_stdDay As String
Dim m_endDay As String
Dim stST05 As String, stST15 As String

Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim m_strListPer As String 'Add By Sindy 2020/7/1


Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim m_i As Integer
Dim m_j As Integer
Dim m_width As Long
Dim m_height As Long
Dim m_posY As Long
Dim m_line As Integer
Dim m_std As Integer
Dim m_end As Integer
   If Grd1.Rows < 2 Then
      MsgBox "沒有待列印資料!!", vbCritical, "發生錯誤！"
      Exit Sub
   End If
   If Grd1.Rows = 2 And Grd1.TextMatrix(1, 0) = "" Then
      MsgBox "沒有待列印資料!!", vbCritical, "發生錯誤！"
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Grd1.MousePointer = flexHourglass
   'edit by nickc 2008/05/08  鎖定副總、阿蓮74028 才可以印直式
   'modify by sonia 2019/4/12阿蓮調職改成莊敏惠73017
   'modify by sonia 2019/5/15再改屬於北所業務助理人員才可印直式
   'If strUserNum = "68006" Or strUserNum = "73017" Or Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
   'Modified by Morgan 2019/12/30 杜主秘說加開簡協理69005可看全所
   'Modified by Lydia 2022/05/03 簡協理69005改為抓系統特殊設定「全所智權部主管」
   'If strUserNum = "68006" Or strUserNum = "69005" Or InStr(Pub_GetSpecMan("北所業務助理人員"), strUserNum) > 0 Or Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
   If strUserNum = "68006" Or InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0 Or InStr(Pub_GetSpecMan("北所業務助理人員"), strUserNum) > 0 Or Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
   Else
       GoTo H2
       Exit Sub
   End If
   m_line = 55
   Printer.Orientation = 1
   Printer.Font.Size = 18
   Printer.Font.Underline = True
   Printer.FontBold = True
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較") / 2)
   Printer.CurrentY = 300
   Printer.Print ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較"
   Printer.Font.Size = 10
   Printer.Font.Underline = False
   Printer.FontBold = False
   '2010/7/29 add by sonia
   If Text2 = "N" Then
      Printer.CurrentX = 9500
      Printer.CurrentY = 400
      Printer.Print " 不含結餘及保留"
   End If
   '2010/7/29 end
   m_posY = 600
   m_height = Printer.ScaleHeight / 65
   m_width = Printer.ScaleWidth / 14
   Dim m_i_Pok As Integer 'Add By Sindy 2010/8/2
   With Grd1
       For m_i = 0 To .Rows - 1
         'Add By Sindy 2010/8/2
         If ((.TextMatrix(m_i, 11) <> "" Or .TextMatrix(m_i, 0) = "國內部") And Text1 = "2") Or _
           Text1 = "1" Then
           If Text1 = "2" Then
               If m_i = 0 Then
                   m_i_Pok = 0
               Else
                   m_i_Pok = m_i_Pok + 1
               End If
           Else
               m_i_Pok = m_i
           End If
           '2010/8/2 End
           'Modify By Sindy 2010/8/2 把計算X軸及Y軸的m_i變數值改為m_i_Pok
           If (m_i_Pok Mod m_line) = 0 And m_i_Pok <> 0 Then
               For m_j = 0 To .Cols - 1
                   Select Case m_j
                   Case 2, 4, 5, 7, 8, 11
                       Printer.CurrentX = m_width * (m_j + 2) - Printer.TextWidth(.TextMatrix(m_i, m_j)) - 80
                       Printer.CurrentY = m_height * ((m_line) + 1) + m_posY
                       Printer.Print .TextMatrix(m_i, m_j)
                   Case 3, 6, 9
                       Printer.CurrentX = m_width * (m_j + 2) - Printer.TextWidth(Format(.TextMatrix(m_i, m_j), "0.00")) - 80
                       Printer.CurrentY = m_height * ((m_line) + 1) + m_posY
                       Printer.Print Format(.TextMatrix(m_i, m_j), "0.00")
                   Case 10
                       If .TextMatrix(m_i, m_j) = "所佔國內" Then
                           Printer.CurrentX = m_width * (m_j + 1)
                           Printer.CurrentY = m_height * ((m_line) + 1) + m_posY
                           Printer.Print .TextMatrix(m_i, m_j)
                       Else
                           Printer.CurrentX = m_width * (m_j + 2) - Printer.TextWidth(.TextMatrix(m_i, m_j)) - 80
                           Printer.CurrentY = m_height * ((m_line) + 1) + m_posY
                           Printer.Print .TextMatrix(m_i, m_j)
                       End If
                   Case Else
                       Printer.CurrentX = m_width * (m_j + 1)
                       Printer.CurrentY = m_height * ((m_line) + 1) + m_posY
                       Printer.Print .TextMatrix(m_i, m_j)
                   End Select
                   '直線
                   Printer.Line (m_width * (m_j + 1), (m_height * 1) - 50 + m_posY)-(m_width * (m_j + 1), (m_height * ((m_line) + 2)) - 50 + m_posY)
               Next m_j
               '橫線
               Printer.Line (m_width, (m_height * (m_line + 1)) - 50 + m_posY)-(m_width * (.Cols + 1), (m_height * (m_line + 1)) - 50 + m_posY)
               Printer.Line (m_width, (m_height * (m_line + 2)) - 50 + m_posY)-(m_width * (.Cols + 1), (m_height * (m_line + 2)) - 50 + m_posY)
               '直線
               Printer.Line (m_width * (.Cols + 1), (m_height * 1) - 50 + m_posY)-(m_width * (.Cols + 1), (m_height * (m_line + 2)) - 50 + m_posY)
               Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("第 " & Format(Printer.Page, "#0") & " 頁")
               Printer.CurrentY = (m_height * (m_line + 3)) - 50 + m_posY
               Printer.Print "第 " & Format(Printer.Page, "#0") & " 頁"
               Printer.NewPage
               Printer.Font.Size = 18
               Printer.Font.Underline = True
               Printer.FontBold = True
               Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較") / 2)
               Printer.CurrentY = 300
               Printer.Print ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較"
               Printer.Font.Size = 10
               Printer.Font.Underline = False
               Printer.FontBold = False
               '2010/7/29 add by sonia
               If Text2 = "N" Then
                  Printer.CurrentX = 9500
                  Printer.CurrentY = 400
                  Printer.Print " 不含結餘及保留"
               End If
               '2010/7/29 end
               m_posY = 600
           End If
           For m_j = 0 To .Cols - 1
               If (m_i_Pok Mod m_line) = 0 Then
                   Printer.CurrentX = m_width * (m_j + 1)
                   Printer.CurrentY = m_height * ((m_i_Pok Mod m_line) + 1) + m_posY
                   Printer.Print StrToStr(.TextMatrix((m_i_Pok Mod m_line), m_j), 4)
               Else
                   Select Case m_j
                   Case 2, 4, 5, 7, 8, 11
                       Printer.CurrentX = m_width * (m_j + 2) - Printer.TextWidth(.TextMatrix(m_i, m_j)) - 80
                       Printer.CurrentY = m_height * ((m_i_Pok Mod m_line) + 1) + m_posY
                       Printer.Print .TextMatrix(m_i, m_j)
                   Case 3, 6, 9
                       Printer.CurrentX = m_width * (m_j + 2) - Printer.TextWidth(Format(.TextMatrix(m_i, m_j), "0.00")) - 80
                       Printer.CurrentY = m_height * ((m_i_Pok Mod m_line) + 1) + m_posY
                       Printer.Print Format(.TextMatrix(m_i, m_j), "0.00")
                   Case 10
                       If .TextMatrix(m_i, m_j) = "所佔國內" Then
                           Printer.CurrentX = m_width * (m_j + 1)
                           Printer.CurrentY = m_height * ((m_i_Pok Mod m_line) + 1) + m_posY
                           Printer.Print .TextMatrix(m_i, m_j)
                       Else
                           Printer.CurrentX = m_width * (m_j + 2) - Printer.TextWidth(.TextMatrix(m_i, m_j)) - 80
                           Printer.CurrentY = m_height * ((m_i_Pok Mod m_line) + 1) + m_posY
                           Printer.Print .TextMatrix(m_i, m_j)
                       End If
                   Case Else
                       Printer.CurrentX = m_width * (m_j + 1)
                       Printer.CurrentY = m_height * ((m_i_Pok Mod m_line) + 1) + m_posY
                       Printer.Print .TextMatrix(m_i, m_j)
                   End Select
               End If
               '直線
               Printer.Line (m_width * (m_j + 1), (m_height * 1) - 50 + m_posY)-(m_width * (m_j + 1), (m_height * ((m_i_Pok Mod m_line) + 2)) - 50 + m_posY)
           Next m_j
           '橫線
           Printer.Line (m_width, m_height * ((m_i_Pok Mod m_line) + 1) - 50 + m_posY)-(m_width * (.Cols + 1), m_height * ((m_i_Pok Mod m_line) + 1) - 50 + m_posY)
         End If
       Next m_i
       If m_i_Pok Mod m_line = 0 Then
           Printer.Line (m_width, (m_height * ((m_line) + 1)) - 50 + m_posY)-(m_width * (.Cols + 1), (m_height * ((m_line) + 1)) - 50 + m_posY)
           Printer.Line (m_width * (.Cols + 1), (m_height * 1) - 50 + m_posY)-(m_width * (.Cols + 1), (m_height * ((m_line) + 1)) - 50 + m_posY)
           Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("第 " & Format(Printer.Page, "#0") & " 頁")
           Printer.CurrentY = (m_height * (m_line + 3)) - 50 + m_posY
           Printer.Print "第 " & Format(Printer.Page, "#0") & " 頁"
       Else
           'Modify By Sindy 2010/8/2 把計算X軸及Y軸的m_i變數值改為m_i_Pok 且原+1改+2
           Printer.Line (m_width, (m_height * ((m_i_Pok Mod m_line) + 2)) - 50 + m_posY)-(m_width * (.Cols + 1), (m_height * ((m_i_Pok Mod m_line) + 2)) - 50 + m_posY)
           Printer.Line (m_width * (.Cols + 1), (m_height * 1) - 50 + m_posY)-(m_width * (.Cols + 1), (m_height * ((m_i_Pok Mod m_line) + 2)) - 50 + m_posY)
           Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("第 " & Format(Printer.Page, "#0") & " 頁")
           Printer.CurrentY = (m_height * (m_line + 3)) - 50 + m_posY
           Printer.Print "第 " & Format(Printer.Page, "#0") & " 頁"
       End If
   End With
   Printer.EndDoc
   'Add By Sindy 2010/9/23
   '新增一張直印報表：新客戶數收文系統別分析
   m_line = 55
   Printer.Orientation = 1
   Printer.Font.Size = 18
   Printer.Font.Underline = True
   Printer.FontBold = True
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新客戶數收文系統別分析") / 2)
   Printer.CurrentY = 300
   Printer.Print ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新客戶數收文系統別分析"
   Printer.Font.Size = 10
   Printer.Font.Underline = False
   Printer.FontBold = False
   '2010/7/29 add by sonia
   If Text2 = "N" Then
      Printer.CurrentX = 9500
      Printer.CurrentY = 400
      Printer.Print " 不含結餘及保留"
   End If
   '2010/7/29 end
   m_posY = 600
   m_height = Printer.ScaleHeight / 65
   m_width = Printer.ScaleWidth / 14
   'Dim m_i_Pok As Integer 'Add By Sindy 2010/8/2
   m_i_Pok = 0
   With Grd2
       For m_i = 0 To .Rows - 1
         'Add By Sindy 2010/8/2
         If m_i = 0 Or ((.TextMatrix(m_i, 1) = "北一區" Or _
               .TextMatrix(m_i, 1) = "北三區" Or _
               .TextMatrix(m_i, 1) = "北四區" Or _
               .TextMatrix(m_i, 1) = "北五區" Or _
               .TextMatrix(m_i, 1) = "北所合計" Or _
               .TextMatrix(m_i, 1) = "中一區" Or _
               .TextMatrix(m_i, 1) = "中二區" Or _
               .TextMatrix(m_i, 1) = "中三區" Or _
               .TextMatrix(m_i, 1) = "中區其他" Or _
               .TextMatrix(m_i, 1) = "中所合計" Or _
               .TextMatrix(m_i, 1) = "南所合計" Or _
               .TextMatrix(m_i, 1) = "高所合計" Or _
               .TextMatrix(m_i, 0) = "其他" Or _
               .TextMatrix(m_i, 0) = "國內部") And Text1 = "2") Or Text1 = "1" Then
           If Text1 = "2" Then
               If m_i = 0 Then
                   m_i_Pok = 0
               Else
                   m_i_Pok = m_i_Pok + 1
               End If
           Else
               m_i_Pok = m_i
           End If
           '2010/8/2 End
           'Modify By Sindy 2010/8/2 把計算X軸及Y軸的m_i變數值改為m_i_Pok
           If (m_i_Pok Mod m_line) = 0 And m_i_Pok <> 0 Then
               For m_j = 0 To .Cols - 1
                   Select Case m_j
                   Case 2, 3, 4, 5, 6, 7, 8, 9
                       Printer.CurrentX = m_width * (m_j + 2) - Printer.TextWidth(.TextMatrix(m_i, m_j)) - 80
                       Printer.CurrentY = m_height * ((m_line) + 1) + m_posY
                       Printer.Print .TextMatrix(m_i, m_j)
                   Case Else
                       Printer.CurrentX = m_width * (m_j + 1)
                       Printer.CurrentY = m_height * ((m_line) + 1) + m_posY
                       Printer.Print .TextMatrix(m_i, m_j)
                   End Select
                   '直線
                   Printer.Line (m_width * (m_j + 1), (m_height * 1) - 50 + m_posY)-(m_width * (m_j + 1), (m_height * ((m_line) + 2)) - 50 + m_posY)
               Next m_j
               '橫線
               Printer.Line (m_width, (m_height * (m_line + 1)) - 50 + m_posY)-(m_width * (.Cols + 1), (m_height * (m_line + 1)) - 50 + m_posY)
               Printer.Line (m_width, (m_height * (m_line + 2)) - 50 + m_posY)-(m_width * (.Cols + 1), (m_height * (m_line + 2)) - 50 + m_posY)
               '直線
               Printer.Line (m_width * (.Cols + 1), (m_height * 1) - 50 + m_posY)-(m_width * (.Cols + 1), (m_height * (m_line + 2)) - 50 + m_posY)
               Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("第 " & Format(Printer.Page, "#0") & " 頁")
               Printer.CurrentY = (m_height * (m_line + 3)) - 50 + m_posY
               Printer.Print "第 " & Format(Printer.Page, "#0") & " 頁"
               Printer.NewPage
               Printer.Font.Size = 18
               Printer.Font.Underline = True
               Printer.FontBold = True
               Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較") / 2)
               Printer.CurrentY = 300
               Printer.Print ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較"
               Printer.Font.Size = 10
               Printer.Font.Underline = False
               Printer.FontBold = False
               '2010/7/29 add by sonia
               If Text2 = "N" Then
                  Printer.CurrentX = 9500
                  Printer.CurrentY = 400
                  Printer.Print " 不含結餘及保留"
               End If
               '2010/7/29 end
               m_posY = 600
           End If
           For m_j = 0 To .Cols - 1
               If (m_i_Pok Mod m_line) = 0 Then
                   Printer.CurrentX = m_width * (m_j + 1)
                   Printer.CurrentY = m_height * ((m_i_Pok Mod m_line) + 1) + m_posY
                   If .TextMatrix((m_i_Pok Mod m_line), m_j) = "P" Or _
                      .TextMatrix((m_i_Pok Mod m_line), m_j) = "T" Or _
                      .TextMatrix((m_i_Pok Mod m_line), m_j) = "L" Then
                      Printer.Print "       " & StrToStr(.TextMatrix((m_i_Pok Mod m_line), m_j), 4)
                   ElseIf .TextMatrix((m_i_Pok Mod m_line), m_j) = "CFP" Or _
                      .TextMatrix((m_i_Pok Mod m_line), m_j) = "CFT" Or _
                      .TextMatrix((m_i_Pok Mod m_line), m_j) = "CFL" Or _
                      .TextMatrix((m_i_Pok Mod m_line), m_j) = "FCP" Then
                      Printer.Print "     " & StrToStr(.TextMatrix((m_i_Pok Mod m_line), m_j), 4)
                   Else
                      Printer.Print StrToStr(.TextMatrix((m_i_Pok Mod m_line), m_j), 4)
                   End If
               Else
                   Select Case m_j
                   Case 2, 3, 4, 5, 6, 7, 8, 9
                       Printer.CurrentX = m_width * (m_j + 2) - Printer.TextWidth(.TextMatrix(m_i, m_j)) - 80
                       Printer.CurrentY = m_height * ((m_i_Pok Mod m_line) + 1) + m_posY
                       Printer.Print .TextMatrix(m_i, m_j)
                   Case Else
                       Printer.CurrentX = m_width * (m_j + 1)
                       Printer.CurrentY = m_height * ((m_i_Pok Mod m_line) + 1) + m_posY
                       Printer.Print .TextMatrix(m_i, m_j)
                   End Select
               End If
               '直線
               Printer.Line (m_width * (m_j + 1), (m_height * 1) - 50 + m_posY)-(m_width * (m_j + 1), (m_height * ((m_i_Pok Mod m_line) + 2)) - 50 + m_posY)
           Next m_j
           '橫線
           Printer.Line (m_width, m_height * ((m_i_Pok Mod m_line) + 1) - 50 + m_posY)-(m_width * (.Cols + 1), m_height * ((m_i_Pok Mod m_line) + 1) - 50 + m_posY)
         End If
       Next m_i
       If m_i_Pok Mod m_line = 0 Then
           Printer.Line (m_width, (m_height * ((m_line) + 1)) - 50 + m_posY)-(m_width * (.Cols + 1), (m_height * ((m_line) + 1)) - 50 + m_posY)
           Printer.Line (m_width * (.Cols + 1), (m_height * 1) - 50 + m_posY)-(m_width * (.Cols + 1), (m_height * ((m_line) + 1)) - 50 + m_posY)
           Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("第 " & Format(Printer.Page, "#0") & " 頁")
           Printer.CurrentY = (m_height * (m_line + 3)) - 50 + m_posY
           Printer.Print "第 " & Format(Printer.Page, "#0") & " 頁"
       Else
           'Modify By Sindy 2010/8/2 把計算X軸及Y軸的m_i變數值改為m_i_Pok 且原+1改+2
           Printer.Line (m_width, (m_height * ((m_i_Pok Mod m_line) + 2)) - 50 + m_posY)-(m_width * (.Cols + 1), (m_height * ((m_i_Pok Mod m_line) + 2)) - 50 + m_posY)
           Printer.Line (m_width * (.Cols + 1), (m_height * 1) - 50 + m_posY)-(m_width * (.Cols + 1), (m_height * ((m_i_Pok Mod m_line) + 2)) - 50 + m_posY)
           Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("第 " & Format(Printer.Page, "#0") & " 頁")
           Printer.CurrentY = (m_height * (m_line + 3)) - 50 + m_posY
           Printer.Print "第 " & Format(Printer.Page, "#0") & " 頁"
       End If
   End With
   Printer.EndDoc
   '2010/9/23 End
H2:
   '橫印
   m_line = 22
   Printer.Orientation = 2
   Printer.Font.Size = 18
   Printer.Font.Underline = True
   Printer.FontBold = True
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較") / 2)
   Printer.CurrentY = 300
   Printer.Print ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較"
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   '2010/7/29 add by sonia
   If Text2 = "N" Then
      Printer.CurrentX = 13250
      Printer.CurrentY = 400
      Printer.Print " 不含結餘及保留"
   End If
   '2010/7/29 end
   m_height = Printer.ScaleHeight / 28
   m_width = Printer.ScaleWidth / 13
   Printer.CurrentX = m_width
   Printer.CurrentY = 800
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1500
   Printer.CurrentY = 800
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   m_posY = 900
   
   If Text1 = "1" Then '明細
         '分三段
         m_std = 1
         '北所合計
         For m_i = m_std To Grd1.Rows - 1
             If Grd1.TextMatrix(m_i, 1) = "北所合計" Then
                 m_end = m_i
                 Exit For
             End If
         Next m_i
         For m_j = 1 To Grd1.Cols - 1
                 Printer.CurrentX = m_width * (m_j)
                 Printer.CurrentY = m_height * (1) + m_posY
                 Printer.Print StrToStr(Grd1.TextMatrix(0, m_j), 4)
             '直線
             Printer.Line (m_width * (m_j), (m_height * 1) - 50 + m_posY)-(m_width * (m_j), (m_height * (2)) - 50 + m_posY)
         Next m_j
         '橫線
         Printer.Line (m_width, m_height * (1) - 50 + m_posY)-(m_width * (Grd1.Cols), m_height * (1) - 50 + m_posY)
         
         PrintBig m_std, m_end, m_line, m_height, m_width, m_posY, "1"
         m_std = m_end + 1
         
         '中所合計
         For m_i = m_std To Grd1.Rows - 1
             If Grd1.TextMatrix(m_i, 1) = "中所合計" Then
                 m_end = m_i
                 Exit For
             End If
         Next m_i
         If m_std <> m_end + 1 Then
            Printer.NewPage
            Printer.Font.Size = 18
            Printer.Font.Underline = True
            Printer.FontBold = True
            Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較") / 2)
            Printer.CurrentY = 300
            Printer.Print ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較"
            Printer.Font.Size = 12
            Printer.Font.Underline = False
            Printer.FontBold = False
            '2010/7/29 add by sonia
            If Text2 = "N" Then
               Printer.CurrentX = 13250
               Printer.CurrentY = 400
               Printer.Print " 不含結餘及保留"
            End If
            '2010/7/29 end
            Printer.CurrentX = m_width
            Printer.CurrentY = 800
            Printer.Print "列印人：" & strUserName
            Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1500
            Printer.CurrentY = 800
            Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
            
            For m_j = 1 To Grd1.Cols - 1
                    Printer.CurrentX = m_width * (m_j)
                    Printer.CurrentY = m_height * (1) + m_posY
                    Printer.Print StrToStr(Grd1.TextMatrix(0, m_j), 4)
                '直線
                Printer.Line (m_width * (m_j), (m_height * 1) - 50 + m_posY)-(m_width * (m_j), (m_height * (2)) - 50 + m_posY)
            Next m_j
            '橫線
            Printer.Line (m_width, m_height * (1) - 50 + m_posY)-(m_width * (Grd1.Cols), m_height * (1) - 50 + m_posY)
            
            PrintBig m_std, m_end, m_line, m_height, m_width, m_posY, "2"
            m_std = m_end + 1
         End If
         
         If m_std < Grd1.Rows - 1 Then
            Printer.NewPage
            Printer.Font.Size = 18
            Printer.Font.Underline = True
            Printer.FontBold = True
            Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較") / 2)
            Printer.CurrentY = 300
            Printer.Print ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較"
            Printer.Font.Size = 12
            Printer.Font.Underline = False
            Printer.FontBold = False
            '2010/7/29 add by sonia
            If Text2 = "N" Then
               Printer.CurrentX = 13250
               Printer.CurrentY = 400
               Printer.Print " 不含結餘及保留"
            End If
            '2010/7/29 end
            Printer.CurrentX = m_width
            Printer.CurrentY = 800
            Printer.Print "列印人：" & strUserName
            Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1500
            Printer.CurrentY = 800
            Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
            
            '剩下的
            m_end = Grd1.Rows - 1
            For m_j = 1 To Grd1.Cols - 1
                    Printer.CurrentX = m_width * (m_j)
                    Printer.CurrentY = m_height * (1) + m_posY
                    Printer.Print StrToStr(Grd1.TextMatrix(0, m_j), 4)
                '直線
                Printer.Line (m_width * (m_j), (m_height * 1) - 50 + m_posY)-(m_width * (m_j), (m_height * (2)) - 50 + m_posY)
            Next m_j
            '橫線
            Printer.Line (m_width, m_height * (1) - 50 + m_posY)-(m_width * (Grd1.Cols), m_height * (1) - 50 + m_posY)
            
            PrintBig m_std, m_end, m_line, m_height, m_width, m_posY, "3"
            m_std = m_end + 1
         End If
         
   'Add By Sindy 2010/8/2
   ElseIf Text1 = "2" Then '區合計
         m_std = 1
         m_end = Grd1.Rows - 1
         For m_j = 1 To Grd1.Cols - 1
                 Printer.CurrentX = m_width * (m_j)
                 Printer.CurrentY = m_height * (1) + m_posY
                 Printer.Print StrToStr(Grd1.TextMatrix(0, m_j), 4)
             '直線
             Printer.Line (m_width * (m_j), (m_height * 1) - 50 + m_posY)-(m_width * (m_j), (m_height * (2)) - 50 + m_posY)
         Next m_j
         '橫線
         Printer.Line (m_width, m_height * (1) - 50 + m_posY)-(m_width * (Grd1.Cols), m_height * (1) - 50 + m_posY)
         PrintBig m_std, m_end, m_line, m_height, m_width, m_posY, "1"
   End If
   '2010/8/2 End
   Printer.EndDoc
   Grd1.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Sub

Sub PrintBig(m_std As Integer, m_end As Integer, m_line As Integer, m_height As Long, m_width As Long, m_posY As Long, strArea As String)
Dim m_i As Integer
Dim m_j As Integer
Dim m_i_Pok As Integer 'Add By Sindy 2010/8/2

With Grd1
    For m_i = m_std To m_end
      'Add By Sindy 2010/8/2
      If ((.TextMatrix(m_i, 11) <> "" Or .TextMatrix(m_i, 0) = "國內部") And Text1 = "2") Or _
        Text1 = "1" Then
        If Text1 = "2" Then
            If m_i_Pok = 0 Then
              m_i_Pok = m_std
            Else
              If strArea = "1" Then '北所
                 m_i_Pok = m_i_Pok + 1
              Else
                 m_i_Pok = m_i_Pok + 1 + m_std
              End If
            End If
        Else
            m_i_Pok = m_i
        End If
        '2010/8/2 End
        'Modify By Sindy 2010/8/2 把計算X軸及Y軸的m_i變數值改為m_i_Pok
        If ((m_i_Pok - m_std + 1) Mod m_line) = 0 And m_i_Pok <> m_std And m_i_Pok <> 0 And m_std <> 0 Then
            For m_j = 1 To .Cols - 1
                Select Case m_j
                Case 2, 4, 5, 7, 8, 11
                    Printer.CurrentX = m_width * (m_j + 1) - Printer.TextWidth(.TextMatrix(m_i, m_j)) - 80
                    Printer.CurrentY = m_height * ((m_line) + 1) + m_posY
                    Printer.Print .TextMatrix(m_i, m_j)
                Case 3, 6, 9
                    Printer.CurrentX = m_width * (m_j + 1) - Printer.TextWidth(Format(.TextMatrix(m_i, m_j), "0.00")) - 80
                    Printer.CurrentY = m_height * ((m_line) + 1) + m_posY
                    Printer.Print Format(.TextMatrix(m_i, m_j), "0.00")
                Case 10
                    If .TextMatrix(m_i, m_j) = "所佔國內" Then
                        'Modify By Sindy 2010/8/2
                        'Printer.CurrentX = m_width * (m_j + 1)
                        Printer.CurrentX = m_width * (m_j)
                        Printer.CurrentY = m_height * ((m_line) + 1) + m_posY
                        Printer.Print .TextMatrix(m_i, m_j)
                    Else
                        'Modify By Sindy 2010/8/2
                        'Printer.CurrentX = m_width * (m_j) - Printer.TextWidth(.TextMatrix(m_i, m_j)) - 80
                        Printer.CurrentX = m_width * (m_j + 1) - Printer.TextWidth(.TextMatrix(m_i, m_j)) - 80
                        Printer.CurrentY = m_height * ((m_line) + 1) + m_posY
                        Printer.Print .TextMatrix(m_i, m_j)
                    End If
                Case Else
                    Printer.CurrentX = m_width * (m_j)
                    Printer.CurrentY = m_height * ((m_line) + 1) + m_posY
                    Printer.Print .TextMatrix(m_i, m_j)
                End Select
                '直線
                Printer.Line (m_width * (m_j), (m_height * (1)) - 50 + m_posY)-(m_width * (m_j), (m_height * ((m_line) + 2)) - 50 + m_posY)
            Next m_j
            '橫線
            Printer.Line (m_width, (m_height * (m_line + 1)) - 50 + m_posY)-(m_width * (.Cols), (m_height * (m_line + 1)) - 50 + m_posY)
            Printer.Line (m_width, (m_height * (m_line + 2)) - 50 + m_posY)-(m_width * (.Cols), (m_height * (m_line + 2)) - 50 + m_posY)
            '直線
            Printer.Line (m_width * (.Cols), (m_height * (1)) - 50 + m_posY)-(m_width * (.Cols), (m_height * (m_line + 2)) - 50 + m_posY)
            If m_i_Pok < m_end Then 'Modify By Sindy 2010/8/2 增加if
               Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("第 " & Format(Printer.Page, "#0") & " 頁")
               Printer.CurrentY = (m_height * (m_line + 3)) - 50 + m_posY
               Printer.Print "第 " & Format(Printer.Page, "#0") & " 頁"
               Printer.NewPage
               Printer.Font.Size = 18
               Printer.Font.Underline = True
               Printer.FontBold = True
               Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較") / 2)
               Printer.CurrentY = 300
               Printer.Print ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較"
               Printer.Font.Size = 12
               Printer.Font.Underline = False
               Printer.FontBold = False
               m_posY = 600
            End If
        End If
        For m_j = 1 To .Cols - 1
            If ((m_i_Pok - m_std + 1) Mod m_line) = 0 Then
                Printer.CurrentX = m_width * (m_j)
                Printer.CurrentY = m_height * (((m_i_Pok - m_std + 1) Mod m_line) + 1) + m_posY
                Printer.Print StrToStr(.TextMatrix(((m_i - m_std + 1) Mod m_line), m_j), 4)
            Else
                Select Case m_j
                Case 2, 4, 5, 7, 8, 11
                    Printer.CurrentX = m_width * (m_j + 1) - Printer.TextWidth(.TextMatrix(m_i, m_j)) - 80
                    Printer.CurrentY = m_height * (((m_i_Pok - m_std + 1) Mod m_line) + 1) + m_posY
                    Printer.Print .TextMatrix(m_i, m_j)
                Case 3, 6, 9
                    Printer.CurrentX = m_width * (m_j + 1) - Printer.TextWidth(Format(.TextMatrix(m_i, m_j), "0.00")) - 80
                    Printer.CurrentY = m_height * (((m_i_Pok - m_std + 1) Mod m_line) + 1) + m_posY
                    Printer.Print Format(.TextMatrix(m_i, m_j), "0.00")
                Case 10
                    If .TextMatrix(m_i, m_j) = "所佔國內" Then
                        Printer.CurrentX = m_width * (m_j)
                        Printer.CurrentY = m_height * (((m_i_Pok - m_std + 1) Mod m_line) + 1) + m_posY
                        Printer.Print .TextMatrix(m_i, m_j)
                    Else
                        Printer.CurrentX = m_width * (m_j + 1) - Printer.TextWidth(.TextMatrix(m_i, m_j)) - 80
                        Printer.CurrentY = m_height * (((m_i_Pok - m_std + 1) Mod m_line) + 1) + m_posY
                        Printer.Print .TextMatrix(m_i, m_j)
                    End If
                Case 1
                    Printer.CurrentX = m_width * (m_j)
                    Printer.CurrentY = m_height * (((m_i_Pok - m_std + 1) Mod m_line) + 1) + m_posY
                    If Trim(.TextMatrix(m_i, m_j)) = "" Then
                        Printer.Print .TextMatrix(m_i, m_j - 1)
                    Else
                        Printer.Print .TextMatrix(m_i, m_j)
                    End If
                Case Else
                    Printer.CurrentX = m_width * (m_j)
                    Printer.CurrentY = m_height * (((m_i_Pok - m_std + 1) Mod m_line) + 1) + m_posY
                    Printer.Print .TextMatrix(m_i, m_j)
                End Select
            End If
            '直線
            Printer.Line (m_width * (m_j), (m_height * (1)) - 50 + m_posY)-(m_width * (m_j), (m_height * (((m_i_Pok - m_std + 1) Mod m_line) + 2)) - 50 + m_posY)
        Next m_j
        '橫線
        Printer.Line (m_width, m_height * (((m_i_Pok - m_std + 1) Mod m_line) + 1) - 50 + m_posY)-(m_width * (.Cols), m_height * (((m_i_Pok - m_std + 1) Mod m_line) + 1) - 50 + m_posY)
      End If
    Next m_i
    
    If m_i_Pok Mod m_line = 0 Then
        Printer.Line (m_width, (m_height * ((m_line) + 1)) - 50 + m_posY)-(m_width * (.Cols), (m_height * ((m_line) + 1)) - 50 + m_posY)
        Printer.Line (m_width * (.Cols), (m_height * (1)) - 50 + m_posY)-(m_width * (.Cols), (m_height * ((m_line) + 1)) - 50 + m_posY)
        Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("第 " & Format(Printer.Page, "#0") & " 頁")
        Printer.CurrentY = (m_height * (m_line + 3)) - 50 + m_posY
        Printer.Print "第 " & Format(Printer.Page, "#0") & " 頁"
    Else
        'Modify By Sindy 2010/8/2 把計算X軸及Y軸的m_i變數值改為m_i_Pok 且原+1改+2
        Printer.Line (m_width, (m_height * (((m_i_Pok - m_std + 1) Mod m_line) + 2)) - 50 + m_posY)-(m_width * (.Cols), (m_height * (((m_i_Pok - m_std + 1) Mod m_line) + 2)) - 50 + m_posY)
        Printer.Line (m_width * (.Cols), (m_height * (1)) - 50 + m_posY)-(m_width * (.Cols), (m_height * (((m_i_Pok - m_std + 1) Mod m_line) + 2)) - 50 + m_posY)
        Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("第 " & Format(Printer.Page, "#0") & " 頁")
        Printer.CurrentY = (m_height * (m_line + 3)) - 50 + m_posY
        Printer.Print "第 " & Format(Printer.Page, "#0") & " 頁"
    End If
End With
End Sub

Private Sub cmdSearch_Click()
Dim Cancel As Boolean
Dim intErrCol As Integer

   If txtCloseDate(0) = "" Then
       MsgBox "日期不可以空白！", vbInformation, "操作錯誤！"
       txtCloseDate(0).SetFocus
       Exit Sub
   End If
   If txtCloseDate(1) = "" Then
       MsgBox "日期不可以空白！", vbInformation, "操作錯誤！"
       txtCloseDate(1).SetFocus
       Exit Sub
   End If
   '2015/4/29 add by sonia
   If Val(txtCloseDate(1)) < Val(txtCloseDate(0)) Then
      MsgBox "收款起迄日範圍錯誤！", vbExclamation
      txtCloseDate(0).SetFocus
      Exit Sub
   End If
   'end 2015/4/29

   'Modify By Sindy 2009/05/14
   Call txtSales_Validate(Cancel)
   If Cancel = True Then
      txtSales.SetFocus
      Exit Sub
   '2009/05/14 End
   End If
   
   'Modify By Sindy 2020/7/29 檢查部門欄位
   'Modify By Sindy 2025/8/11 +, txtZone
   If PUB_ChkFormSalesDept(strUserNum, txtSales, txtSalesArea, txtSalesArea1, intErrCol, txtZone) = False Then
      If intErrCol = 0 Then
         txtSales.SetFocus
         txtSales_GotFocus
      ElseIf intErrCol = 1 Then
         txtSalesArea.SetFocus
         txtSalesArea_GotFocus
      Else
         txtSalesArea1.SetFocus
         txtSalesArea1_GotFocus
      End If
      Exit Sub
   End If
   
'   Else
'      '林永生71003檢查業務區範圍
'      If strUserNum = "71003" Then
'         If txtSalesArea < "S2" Or txtSalesArea > "S29" Then
'            MsgBox "業務區起始條件錯誤！只可查中所業務區", vbExclamation
'            txtSalesArea.SetFocus
'            txtSalesArea_GotFocus
'            Exit Sub
'         End If
'         If txtSalesArea1 < "S2" Or txtSalesArea1 > "S29" Then
'            MsgBox "業務區迄止條件錯誤！只可查中所業務區", vbExclamation
'            txtSalesArea1.SetFocus
'            txtSalesArea1_GotFocus
'            Exit Sub
'         End If
'      End If
'      '簡金泉69005檢查業務區範圍
''Removed by Morgan 2019/12/30 杜主秘說加開簡協理69005可看全所
''      If strUserNum = "69005" Then
''         If txtSalesArea < "S1" Or txtSalesArea > "S19" Then
''            MsgBox "業務區起始條件錯誤！只可查北所業務區", vbExclamation
''            txtSalesArea.SetFocus
''            txtSalesArea_GotFocus
''            Exit Sub
''         End If
''         If txtSalesArea1 < "S1" Or txtSalesArea1 > "S19" Then
''            MsgBox "業務區迄止條件錯誤！只可查北所業務區", vbExclamation
''            txtSalesArea1.SetFocus
''            txtSalesArea1_GotFocus
''            Exit Sub
''         End If
''      End If
''end 2019/12/30
'
'      'add by sonia 2016/12/21 柄佑82026可看中所全部或自已
'      If strUserNum = "82026" Then
'         If txtSalesArea <> "P11" Or txtSalesArea1 <> "P11" Then
'            If txtSalesArea < "S2" Or txtSalesArea > "S29" Then
'               MsgBox "業務區起始條件錯誤！只可查中所業務區", vbExclamation
'               txtSalesArea.SetFocus
'               txtSalesArea_GotFocus
'               Exit Sub
'            End If
'            If txtSalesArea1 < "S2" Or txtSalesArea1 > "S29" Then
'               MsgBox "業務區迄止條件錯誤！只可查中所業務區", vbExclamation
'               txtSalesArea1.SetFocus
'               txtSalesArea1_GotFocus
'               Exit Sub
'            End If
'         Else
'            If Trim(txtSales) <> strUserNum Then
'               MsgBox "查專利處工程師時只可查自己的資料", vbExclamation
'               txtSales.SetFocus
'               txtSales_GotFocus
'               Exit Sub
'            End If
'         End If
'      End If
'      'end 2016/12/21
'   End If
'
'   If txtSalesArea > txtSalesArea1 Then
'      MsgBox "業務區範圍條件錯誤！", vbExclamation
'      txtSalesArea.SetFocus
'      txtSalesArea_GotFocus
'      Exit Sub
'   End If
'
'   '加入外商主管  可以輸入相同組別的
'   If (stST05 = "21" Or stST05 = "26" Or stST05 = "28") Then
'       If Trim(txtSales) = "" Then
'           MsgBox "智權人員不可以空白！", vbExclamation, "操作錯誤！"
'           txtSales.SetFocus
'           txtSales_GotFocus
'           Exit Sub
'       End If
'       If PUB_GetStaffST16(strUserNum) <> PUB_GetStaffST16(txtSales) Then
'           MsgBox "僅可以輸入相同組別的智權人員！", vbExclamation, "操作錯誤！"
'           txtSales.SetFocus
'           txtSales_GotFocus
'           Exit Sub
'       End If
'   End If

   Screen.MousePointer = vbHourglass
   Grd1.MousePointer = flexHourglass
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/23 清除查詢印表記錄檔欄位
   StrMenu
   Grd1.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
   End Sub

Sub StrMenu()
Dim m_str As String
Dim m_rs As New ADODB.Recordset
Dim m_gr1 As String
Dim m_gr2 As String
Dim m_newc As Double, m_newp As Double, m_oldc As Double, m_oldp As Double, m_perc As Double, m_perp As Double
Dim m_Anewc As Double, m_Anewp As Double, m_Aoldc As Double, m_Aoldp As Double, m_Aperc As Double, m_Aperp As Double
Dim m_std As Integer
Dim m_end As Integer
Dim m_i As Integer
Dim m_j As Integer
Dim m_seekst03 As String
Dim m_seekst06 As String
Dim m_tmp As Variant
Dim m_tmp2 As Variant
Dim stConST As String
Dim STCONSTAREA As String    '2010/7/30 ADD BY SONIA
Dim m_stdArea As String
Dim m_endArea As String
   
   m_stdDay = txtCloseDate(0)
   m_endDay = txtCloseDate(1)
   pub_QL05 = pub_QL05 & ";" & Label2 & txtCloseDate(0) & "-" & txtCloseDate(1) 'Add By Sindy 2010/12/23
   'm_stdArea = "S"
   'm_endArea = "S99"
    stConST = ""
'cancel by sonia 2014/6/9
'    If strUserNum = "79037" Then
'       stConST = stConST & " and st06 = '" & pub_strUserOffice & "'"
'    End If
'end 2014/6/9
'   '2005/9/8 ENDsystemkind = "CFT,FCT,S,CFC"
'   '2005/9/12 ADD BY SONIA 陳經理查詢所有智權人員要控制系統類別
'   If strUserNum = "68005" And txtSales <> "68005" Then
'      systemkind = "CFT,FCT,S,CFC"
'   End If
'   '2005/9/12 END
   
   '區別
   'Modify By Sindy 2009/05/12
   '若為帶人主管權限時,查詢之智權人員編號非本人時,不限制區別
   If (Trim(txtSales) <> "" And PUB_GetST05(strUserNum) = "SA" And txtSales.Enabled = True And txtSales <> strUserNum) Then
      '不限制區別
   Else
      If txtSalesArea <> "" Then
         m_stdArea = txtSalesArea
         '2010/7/30 ADD BY SONIA
         STCONSTAREA = STCONSTAREA & " and st15>='" & m_stdArea & "' "
      Else
         STCONSTAREA = STCONSTAREA & " and st15>='S' "
         '2010/7/30 END
      End If
      If txtSalesArea1 <> "" Then
         m_endArea = txtSalesArea1
         '2010/7/30 ADD BY SONIA
         STCONSTAREA = STCONSTAREA & " and st15<='" & m_endArea & "' "
      Else
         STCONSTAREA = STCONSTAREA & " and st15<='S99' "
         '2010/7/30 END
      End If
      pub_QL05 = pub_QL05 & ";" & Label1 & txtSalesArea & "-" & txtSalesArea1 'Add By Sindy 2010/12/23
   End If
   
   '智權人員
   If txtSales <> "" Then
        stConST = stConST & " and AX209 in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1, txtZone) & ") "
        pub_QL05 = pub_QL05 & ";" & Label4 & txtSales & lblSalesName 'Add By Sindy 2010/12/23
   End If
   
   If Text2 <> "" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label5, 7) & Text2 & "(Y:含 N.不含)" 'Add By Sindy 2010/12/23
   End If
   If Text1 = "1" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label6, 5) & "1.明細" 'Add By Sindy 2010/12/23
   Else
      pub_QL05 = pub_QL05 & ";" & Left(Label6, 5) & "2.區合計" 'Add By Sindy 2010/12/23
   End If
   
'抓Grd1資料  語法 秀玲2008/03/07 mail
'2014/1/21 MODIFY BY SONIA 取消 a0201='1' 條件
m_str = "select AA.st15,AA.AX209,AA.st02,BB.newc,AA.newp,BB.oldc,AA.oldp,a0902,decode(st.st06,'1','北所','2','中所','3','南所','4','高所','其他') st06 from ( "
m_str = m_str & " SELECT ST15,';',AX209,';',ST02,';',SUM(DECODE(TAG,'1',POINT,0)) NEWP,';',SUM(DECODE(TAG,'2',POINT,0)) OLDP FROM ( "
m_str = m_str & " select distinct decode(substr(x.mindAtE,1,6), '','2',substr(cu14,1,6),'1','2') tag,st15,ax209,st02,ax202, ax203,ax208,ax214, to_char(round((ax207-ax206)/1000,2),'99999.00') Point,ax212"
m_str = m_str & " from acc020, acc021, staff st,acc1p0,customer,"
m_str = m_str & " (select AX208 CUNO,NVL(min(cp05),19110000) mindate from acc1u0,acc020,acc021,acc1p0,caseprogress"
m_str = m_str & " where ax201(+) = a0201  and ax202(+) = a0202 and ax209 Is Not Null " & stConST
'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
m_str = m_str & " and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and a0205 >=" & m_stdDay & " and a0205 <= " & m_endDay & ""
'modify by sonia 2015/4/20 將此程式所有'2010/7/29 ADD BY SONIA 加的控制,其中INSTR(AX212,'結餘')=0都改為INSTR(AX213||' ','結餘')=0, 科目'4193'改為'4194'
If Text2 = "N" Then m_str = m_str & " AND AX205 NOT IN ('4191','4192','4194') AND INSTR(AX213||' ','結餘')=0" '2010/7/29 ADD BY SONIA
m_str = m_str & " and ax202=a1p22(+) and ax203=a1p03(+) and a1p04=a1u01(+) and a1u03=cp09(+)"
 m_str = m_str & " group by aX208) x"
m_str = m_str & " where ax201(+) = a0201  and ax202(+) = a0202 and st01(+)=ax209 and ax209 Is Not Null " & stConST
'2010/7/30 MODIFY BY SONIA
'm_str = m_str & " and (substr(ax205, 1, 2) = '41' Or ax205 = '7121') and a0205 >= " & m_stdDay & " and a0205 <= " & m_endDay & " and st15>='" & m_stdArea & "' and st15<='" & m_endArea & "' "
'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
m_str = m_str & " and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and a0205 >= " & m_stdDay & " and a0205 <= " & m_endDay & STCONSTAREA
'2010/7/30 END
If Text2 = "N" Then m_str = m_str & " AND AX205 NOT IN ('4191','4192','4194') AND INSTR(AX213||' ','結餘')=0" '2010/7/29 ADD BY SONIA
m_str = m_str & " and substr(ax208,1,8)=cu01(+) and substr(ax208,9,1)=cu02(+)"
m_str = m_str & " and ax202=a1p22(+) and ax203=a1p03(+)"
m_str = m_str & " and ax208=x.CUNO(+)) "
m_str = m_str & " GROUP BY ST15,AX209,ST02 union"
m_str = m_str & " SELECT '國內其他',';',' ',';','國內其他',';',SUM(NEW),';',SUM(OLD) FROM ("
m_str = m_str & " SELECT ST15,AX209,ST02,SUM(DECODE(TAG,'1',POINT,0)) NEW,SUM(DECODE(TAG,'2',POINT,0)) OLD FROM ("
m_str = m_str & " select distinct decode(substr(x.mindAtE,1,6), '','2',substr(cu14,1,6),'1','2') tag,st15,ax209,st02,ax202, ax203,ax208,ax214, to_char(round((ax207-ax206)/1000,2),'99999.00') Point,ax212"
m_str = m_str & " from acc020, acc021, staff,acc1p0,customer,"
m_str = m_str & " (select AX208 CUNO,NVL(min(cp05),19110000) mindate from acc1u0,acc020,acc021,acc1p0,caseprogress"
m_str = m_str & " where ax201(+) = a0201  and ax202(+) = a0202 and ax209 Is Not Null "
'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
m_str = m_str & " and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and a0205 >= " & m_stdDay & " and a0205 <= " & m_endDay & " "
If Text2 = "N" Then m_str = m_str & " AND AX205 NOT IN ('4191','4192','4194') AND INSTR(AX213||' ','結餘')=0" '2010/7/29 ADD BY SONIA
m_str = m_str & " and ax202=a1p22(+) and ax203=a1p03(+) and a1p04=a1u01(+) and a1u03=cp09(+)"
 m_str = m_str & " group by aX208) x"
m_str = m_str & " where ax201(+) = a0201  and ax202(+) = a0202 and st01(+)=ax209 and ax209 Is Not Null "
'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
m_str = m_str & " and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and a0205 >= " & m_stdDay & " and a0205 <= " & m_endDay & " and SUBSTR(st15,1,1)<>'S'"
If Text2 = "N" Then m_str = m_str & " AND AX205 NOT IN ('4191','4192','4194') AND INSTR(AX213||' ','結餘')=0" '2010/7/29 ADD BY SONIA
m_str = m_str & " and substr(st15,1,1)<>'F' and substr(ax208,1,8)=cu01(+) and substr(ax208,9,1)=cu02(+)"
m_str = m_str & " and ax202=a1p22(+) and ax203=a1p03(+)"
m_str = m_str & " and ax208=x.CUNO(+)) "
m_str = m_str & " GROUP BY ST15,AX209,ST02)) AA,("
m_str = m_str & " SELECT ST15,';',AX209,';',ST02,';',SUM(DECODE(TAG,'1',CC)) NewC,';',SUM(DECODE(TAG,'2',CC)) OldC FROM ("
m_str = m_str & " SELECT ST15,AX209,ST02,TAG,COUNT(DISTINCT AX208) CC FROM ("
m_str = m_str & " select distinct decode(substr(x.mindAtE,1,6), '','2',substr(cu14,1,6),'1','2') tag,x.mindAtE-19110000,cu14-19110000,st15,ax209,st02,ax202, ax203,ax208,ax214, to_char(round((ax207-ax206)/1000,2),'99999.00') Point,ax212"
m_str = m_str & " from acc020, acc021, staff,acc1p0,customer,"
m_str = m_str & " (select AX208 CUNO,NVL(min(cp05),19110000) mindate from acc1u0,acc020,acc021,acc1p0,caseprogress"
m_str = m_str & " where ax201(+) = a0201  and ax202(+) = a0202 and ax209 Is Not Null " & stConST
'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
m_str = m_str & " and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and a0205 >= " & m_stdDay & " and a0205 <= " & m_endDay & " "
If Text2 = "N" Then m_str = m_str & " AND AX205 NOT IN ('4191','4192','4194') AND INSTR(AX213||' ','結餘')=0" '2010/7/29 ADD BY SONIA
m_str = m_str & " and ax202=a1p22(+) and ax203=a1p03(+) and a1p04=a1u01(+) and a1u03=cp09(+)"
 m_str = m_str & " group by aX208) x"
m_str = m_str & " where ax201(+) = a0201  and ax202(+) = a0202 and st01(+)=ax209 and ax209 Is Not Null " & stConST
'2010/7/30 MODIFY BY SONIA
'm_str = m_str & " and (substr(ax205, 1, 2) = '41' Or ax205 = '7121') and a0205 >= " & m_stdDay & " and a0205 <= " & m_endDay & " and st15>='" & m_stdArea & "' and st15<='" & m_endArea & "' "
'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
m_str = m_str & " and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and a0205 >= " & m_stdDay & " and a0205 <= " & m_endDay & STCONSTAREA
'2010/7/30 END
If Text2 = "N" Then m_str = m_str & " AND AX205 NOT IN ('4191','4192','4194') AND INSTR(AX213||' ','結餘')=0" '2010/7/29 ADD BY SONIA
m_str = m_str & " and substr(ax208,1,8)=cu01(+) and substr(ax208,9,1)=cu02(+)"
m_str = m_str & " and ax202=a1p22(+) and ax203=a1p03(+)"
m_str = m_str & " and ax208=x.CUNO(+))"
m_str = m_str & " GROUP BY ST15,AX209,ST02,TAG)"
m_str = m_str & " GROUP BY ST15,AX209,ST02 union"
m_str = m_str & " SELECT '國內其他',';',' ',';','國內其他',';',SUM(DECODE(TAG,'1',CC)),';',SUM(DECODE(TAG,'2',CC)) FROM ("
m_str = m_str & " SELECT TAG,COUNT(DISTINCT AX208) CC FROM ("
m_str = m_str & " select distinct decode(substr(x.mindAtE,1,6), '','2',substr(cu14,1,6),'1','2') tag,x.mindAtE-19110000,cu14-19110000,st15,ax209,st02,ax202, ax203,ax208,ax214, to_char(round((ax207-ax206)/1000,2),'99999.00') Point,ax212"
m_str = m_str & " from acc020, acc021, staff,acc1p0,customer,"
m_str = m_str & " (select AX208 CUNO,NVL(min(cp05),19110000) mindate from acc1u0,acc020,acc021,acc1p0,caseprogress"
m_str = m_str & " where ax201(+) = a0201  and ax202(+) = a0202 and ax209 Is Not Null "
'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
m_str = m_str & " and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and a0205 >= " & m_stdDay & " and a0205 <= " & m_endDay & " "
If Text2 = "N" Then m_str = m_str & " AND AX205 NOT IN ('4191','4192','4194') AND INSTR(AX213||' ','結餘')=0" '2010/7/29 ADD BY SONIA
m_str = m_str & " and ax202=a1p22(+) and ax203=a1p03(+) and a1p04=a1u01(+) and a1u03=cp09(+)"
m_str = m_str & " group by aX208) x"
m_str = m_str & " where ax201(+) = a0201  and ax202(+) = a0202 and st01(+)=ax209 and ax209 Is Not Null "
'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
m_str = m_str & " and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and a0205 >= " & m_stdDay & " and a0205 <= " & m_endDay & " and SUBSTR(st15,1,1)<>'S'"
If Text2 = "N" Then m_str = m_str & " AND AX205 NOT IN ('4191','4192','4194') AND INSTR(AX213||' ','結餘')=0" '2010/7/29 ADD BY SONIA
m_str = m_str & " and substr(st15,1,1)<>'F' and substr(ax208,1,8)=cu01(+) and substr(ax208,9,1)=cu02(+)"
m_str = m_str & " and ax202=a1p22(+) and ax203=a1p03(+)"
m_str = m_str & " and ax208=x.CUNO(+))"
m_str = m_str & " GROUP BY ST15,AX209,ST02,TAG))BB,acc090,staff st  where AA.st15=BB.st15 and AA.AX209=BB.AX209 and AA.st02=BB.st02 and AA.st15=a0901(+) and AA.AX209=st.st01(+)  "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    InsertQueryLog (m_rs.RecordCount) 'Add By Sindy 2010/12/23
    Grd1.Visible = False
    Grd1.Clear
    Grd1.Rows = 2
    SetGrd
    With m_rs
'塞資料
        .MoveFirst
        m_gr1 = ""
        m_gr2 = ""
        m_seekst03 = ""
        m_seekst06 = ""
        Do While Not .EOF
            If Grd1.TextMatrix(Grd1.Rows - 1, 0) <> "" Or Grd1.TextMatrix(Grd1.Rows - 1, 1) <> "" Or Grd1.TextMatrix(Grd1.Rows - 1, 2) <> "" Then
                Grd1.Rows = Grd1.Rows + 1
            End If
            If m_gr2 <> CheckStr(m_rs.Fields("a0902")) And m_gr2 <> "" And m_gr2 <> "台南所" And m_gr2 <> "高雄所" Then
                Grd1.TextMatrix(Grd1.Rows - 1, 1) = m_gr2
                m_gr2 = CheckStr(m_rs.Fields("a0902"))
                Grd1.Rows = Grd1.Rows + 1
            End If
            If m_gr1 <> CheckStr(m_rs.Fields("st06")) Then
                If m_gr1 <> "" Then
                    Grd1.TextMatrix(Grd1.Rows - 1, 1) = m_gr1 & "合計"
                    'If m_stdArea = "S" And m_endArea = "S99" Then
                        Grd1.Rows = Grd1.Rows + 1
                        If m_gr1 = "高所" Then
                            Grd1.TextMatrix(Grd1.Rows - 1, 1) = "智權部"
                            Grd1.Rows = Grd1.Rows + 1
                        End If
                    'End If
                End If
                '2010/7/30 MODIFY BY SONIA
                'If m_stdArea = "S" And m_endArea = "S99" Then
                If (m_stdArea = "S" And m_endArea = "S99") Or (m_stdArea = "" And m_endArea = "") Then
                '2010/7/30 END
                    Grd1.TextMatrix(Grd1.Rows - 1, 0) = CheckStr(m_rs.Fields("st06"))
                    m_gr1 = CheckStr(m_rs.Fields("st06"))
                    m_gr2 = CheckStr(m_rs.Fields("a0902"))
                ElseIf CheckStr(m_rs.Fields("st06")) <> "其他" Then
                    Grd1.TextMatrix(Grd1.Rows - 1, 0) = CheckStr(m_rs.Fields("st06"))
                    m_gr1 = CheckStr(m_rs.Fields("st06"))
                    m_gr2 = CheckStr(m_rs.Fields("a0902"))
                End If
            End If
            If CheckStr(m_rs.Fields("st06")) <> "其他" Then
                Grd1.TextMatrix(Grd1.Rows - 1, 1) = CheckStr(m_rs.Fields("st02"))
                Grd1.TextMatrix(Grd1.Rows - 1, 2) = CheckStr(m_rs.Fields("newc"))
                Grd1.TextMatrix(Grd1.Rows - 1, 3) = CheckStr(m_rs.Fields("newp"))
                Grd1.TextMatrix(Grd1.Rows - 1, 5) = CheckStr(m_rs.Fields("oldc"))
                Grd1.TextMatrix(Grd1.Rows - 1, 6) = CheckStr(m_rs.Fields("oldp"))
                Grd1.TextMatrix(Grd1.Rows - 1, 8) = Val(Grd1.TextMatrix(Grd1.Rows - 1, 2)) + Val(Grd1.TextMatrix(Grd1.Rows - 1, 5))
                Grd1.TextMatrix(Grd1.Rows - 1, 9) = Val(Grd1.TextMatrix(Grd1.Rows - 1, 3)) + Val(Grd1.TextMatrix(Grd1.Rows - 1, 6))
                If Val(Grd1.TextMatrix(Grd1.Rows - 1, 3)) <> 0 And Val(Grd1.TextMatrix(Grd1.Rows - 1, 9)) <> 0 Then
                    Grd1.TextMatrix(Grd1.Rows - 1, 4) = Format((Val(Grd1.TextMatrix(Grd1.Rows - 1, 3)) / Val(Grd1.TextMatrix(Grd1.Rows - 1, 9))) * 100, "0.00") & "%"
                Else
                    Grd1.TextMatrix(Grd1.Rows - 1, 4) = "0%"
                End If
                If Val(Grd1.TextMatrix(Grd1.Rows - 1, 6)) <> 0 And Val(Grd1.TextMatrix(Grd1.Rows - 1, 9)) <> 0 Then
                    Grd1.TextMatrix(Grd1.Rows - 1, 7) = Format((Val(Grd1.TextMatrix(Grd1.Rows - 1, 6)) / Val(Grd1.TextMatrix(Grd1.Rows - 1, 9))) * 100, "0.00") & "%"
                Else
                    Grd1.TextMatrix(Grd1.Rows - 1, 7) = "0%"
                End If
            '2010/7/30 ADD BY SONIA
            ElseIf (m_stdArea = "S" And m_endArea = "S99") Or (m_stdArea = "" And m_endArea = "") Then
                Grd1.TextMatrix(Grd1.Rows - 1, 1) = CheckStr(m_rs.Fields("st02"))
                Grd1.TextMatrix(Grd1.Rows - 1, 2) = CheckStr(m_rs.Fields("newc"))
                Grd1.TextMatrix(Grd1.Rows - 1, 3) = CheckStr(m_rs.Fields("newp"))
                Grd1.TextMatrix(Grd1.Rows - 1, 5) = CheckStr(m_rs.Fields("oldc"))
                Grd1.TextMatrix(Grd1.Rows - 1, 6) = CheckStr(m_rs.Fields("oldp"))
                Grd1.TextMatrix(Grd1.Rows - 1, 8) = Val(Grd1.TextMatrix(Grd1.Rows - 1, 2)) + Val(Grd1.TextMatrix(Grd1.Rows - 1, 5))
                Grd1.TextMatrix(Grd1.Rows - 1, 9) = Val(Grd1.TextMatrix(Grd1.Rows - 1, 3)) + Val(Grd1.TextMatrix(Grd1.Rows - 1, 6))
                If Val(Grd1.TextMatrix(Grd1.Rows - 1, 3)) <> 0 And Val(Grd1.TextMatrix(Grd1.Rows - 1, 9)) <> 0 Then
                    Grd1.TextMatrix(Grd1.Rows - 1, 4) = Format((Val(Grd1.TextMatrix(Grd1.Rows - 1, 3)) / Val(Grd1.TextMatrix(Grd1.Rows - 1, 9))) * 100, "0.00") & "%"
                Else
                    Grd1.TextMatrix(Grd1.Rows - 1, 4) = "0%"
                End If
                If Val(Grd1.TextMatrix(Grd1.Rows - 1, 6)) <> 0 And Val(Grd1.TextMatrix(Grd1.Rows - 1, 9)) <> 0 Then
                    Grd1.TextMatrix(Grd1.Rows - 1, 7) = Format((Val(Grd1.TextMatrix(Grd1.Rows - 1, 6)) / Val(Grd1.TextMatrix(Grd1.Rows - 1, 9))) * 100, "0.00") & "%"
                Else
                    Grd1.TextMatrix(Grd1.Rows - 1, 7) = "0%"
                End If
            End If
            .MoveNext
        Loop
        '2010/7/30 MODIFY BY SONIA
        'If m_stdArea = "S" And m_endArea = "S99" Then
        If (m_stdArea = "S" And m_endArea = "S99") Or (m_stdArea = "" And m_endArea = "") Then
        '2010/7/30 END
            Grd1.Rows = Grd1.Rows + 1
            Grd1.TextMatrix(Grd1.Rows - 1, 0) = "國內部"
        End If
    End With
    If Grd1.TextMatrix(Grd1.Rows - 1, 1) = "" And Grd1.TextMatrix(Grd1.Rows - 1, 0) = "" Then
        Grd1.Rows = Grd1.Rows - 1
    End If
'統計及計算
    m_std = 1
    m_newc = 0
    m_newp = 0
    m_oldc = 0
    m_oldp = 0
    m_perc = 0
    m_perp = 0
    m_Anewc = 0
    m_Anewp = 0
    m_Aoldc = 0
    m_Aoldp = 0
    m_Aperc = 0
    m_Aperp = 0
    With Grd1
        '北一區
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "北一區" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "北一區" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If Val(.TextMatrix(m_end, 3)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 4) = "0.00%"
                End If
                If Val(.TextMatrix(m_end, 6)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 7) = "0.00%"
                End If
            Next m_i
            For m_i = m_std To m_end - 1
                If Val(.TextMatrix(m_i, 9)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_i, 10) = Format((Val(.TextMatrix(m_i, 9)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                     .TextMatrix(m_i, 10) = "0.00%"
                End If
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc = m_newc + Val(.TextMatrix(m_end, 2))
            m_newp = m_newp + Val(.TextMatrix(m_end, 3))
            m_oldc = m_oldc + Val(.TextMatrix(m_end, 5))
            m_oldp = m_oldp + Val(.TextMatrix(m_end, 6))
            m_perc = m_perc + Val(.TextMatrix(m_end, 8))
            m_perp = m_perp + Val(.TextMatrix(m_end, 9))
            m_Anewc = m_Anewc + Val(.TextMatrix(m_end, 2))
            m_Anewp = m_Anewp + Val(.TextMatrix(m_end, 3))
            m_Aoldc = m_Aoldc + Val(.TextMatrix(m_end, 5))
            m_Aoldp = m_Aoldp + Val(.TextMatrix(m_end, 6))
            m_Aperc = m_Aperc + Val(.TextMatrix(m_end, 8))
            m_Aperp = m_Aperp + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '北三區
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "北三區" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "北三區" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If Val(.TextMatrix(m_end, 3)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 4) = "0.00%"
                End If
                If Val(.TextMatrix(m_end, 6)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 7) = "0.00%"
                End If
            Next m_i
            For m_i = m_std To m_end - 1
                If Val(.TextMatrix(m_i, 9)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_i, 10) = Format((Val(.TextMatrix(m_i, 9)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                     .TextMatrix(m_i, 10) = "0.00%"
                End If
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc = m_newc + Val(.TextMatrix(m_end, 2))
            m_newp = m_newp + Val(.TextMatrix(m_end, 3))
            m_oldc = m_oldc + Val(.TextMatrix(m_end, 5))
            m_oldp = m_oldp + Val(.TextMatrix(m_end, 6))
            m_perc = m_perc + Val(.TextMatrix(m_end, 8))
            m_perp = m_perp + Val(.TextMatrix(m_end, 9))
            m_Anewc = m_Anewc + Val(.TextMatrix(m_end, 2))
            m_Anewp = m_Anewp + Val(.TextMatrix(m_end, 3))
            m_Aoldc = m_Aoldc + Val(.TextMatrix(m_end, 5))
            m_Aoldp = m_Aoldp + Val(.TextMatrix(m_end, 6))
            m_Aperc = m_Aperc + Val(.TextMatrix(m_end, 8))
            m_Aperp = m_Aperp + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '北四區
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "北四區" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "北四區" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If Val(.TextMatrix(m_end, 3)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 4) = "0.00%"
                End If
                If Val(.TextMatrix(m_end, 6)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 7) = "0.00%"
                End If
            Next m_i
            For m_i = m_std To m_end - 1
                If Val(.TextMatrix(m_i, 9)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_i, 10) = Format((Val(.TextMatrix(m_i, 9)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                     .TextMatrix(m_i, 10) = "0.00%"
                End If
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc = m_newc + Val(.TextMatrix(m_end, 2))
            m_newp = m_newp + Val(.TextMatrix(m_end, 3))
            m_oldc = m_oldc + Val(.TextMatrix(m_end, 5))
            m_oldp = m_oldp + Val(.TextMatrix(m_end, 6))
            m_perc = m_perc + Val(.TextMatrix(m_end, 8))
            m_perp = m_perp + Val(.TextMatrix(m_end, 9))
            m_Anewc = m_Anewc + Val(.TextMatrix(m_end, 2))
            m_Anewp = m_Anewp + Val(.TextMatrix(m_end, 3))
            m_Aoldc = m_Aoldc + Val(.TextMatrix(m_end, 5))
            m_Aoldp = m_Aoldp + Val(.TextMatrix(m_end, 6))
            m_Aperc = m_Aperc + Val(.TextMatrix(m_end, 8))
            m_Aperp = m_Aperp + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '北五區
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "北五區" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "北五區" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If Val(.TextMatrix(m_end, 3)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 4) = "0.00%"
                End If
                If Val(.TextMatrix(m_end, 6)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 7) = "0.00%"
                End If
            Next m_i
            For m_i = m_std To m_end - 1
                If Val(.TextMatrix(m_i, 9)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_i, 10) = Format((Val(.TextMatrix(m_i, 9)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                     .TextMatrix(m_i, 10) = "0.00%"
                End If
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc = m_newc + Val(.TextMatrix(m_end, 2))
            m_newp = m_newp + Val(.TextMatrix(m_end, 3))
            m_oldc = m_oldc + Val(.TextMatrix(m_end, 5))
            m_oldp = m_oldp + Val(.TextMatrix(m_end, 6))
            m_perc = m_perc + Val(.TextMatrix(m_end, 8))
            m_perp = m_perp + Val(.TextMatrix(m_end, 9))
            m_Anewc = m_Anewc + Val(.TextMatrix(m_end, 2))
            m_Anewp = m_Anewp + Val(.TextMatrix(m_end, 3))
            m_Aoldc = m_Aoldc + Val(.TextMatrix(m_end, 5))
            m_Aoldp = m_Aoldp + Val(.TextMatrix(m_end, 6))
            m_Aperc = m_Aperc + Val(.TextMatrix(m_end, 8))
            m_Aperp = m_Aperp + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '北所合計
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "北所合計" Then
                m_end = m_i
                m_seekst06 = m_seekst06 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "北所合計" Then
            .TextMatrix(m_end, 2) = m_newc
            .TextMatrix(m_end, 3) = m_newp
            .TextMatrix(m_end, 5) = m_oldc
            .TextMatrix(m_end, 6) = m_oldp
            .TextMatrix(m_end, 8) = m_perc
            .TextMatrix(m_end, 9) = m_perp
            If Val(.TextMatrix(m_end, 3)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
            Else
                .TextMatrix(m_end, 4) = "0.00%"
            End If
            If Val(.TextMatrix(m_end, 6)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
            Else
                .TextMatrix(m_end, 7) = "0.00%"
            End If
            For m_i = 0 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &HFF80FF   '&HFF00FF
            Next m_i
            m_std = m_end + 1
        End If
        m_newc = 0: m_newp = 0: m_oldc = 0: m_oldp = 0: m_perc = 0: m_perp = 0
        '中一區
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "中一區" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "中一區" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If Val(.TextMatrix(m_end, 3)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 4) = "0.00%"
                End If
                If Val(.TextMatrix(m_end, 6)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 7) = "0.00%"
                End If
            Next m_i
            For m_i = m_std To m_end - 1
                If Val(.TextMatrix(m_i, 9)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_i, 10) = Format((Val(.TextMatrix(m_i, 9)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                     .TextMatrix(m_i, 10) = "0.00%"
                End If
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc = m_newc + Val(.TextMatrix(m_end, 2))
            m_newp = m_newp + Val(.TextMatrix(m_end, 3))
            m_oldc = m_oldc + Val(.TextMatrix(m_end, 5))
            m_oldp = m_oldp + Val(.TextMatrix(m_end, 6))
            m_perc = m_perc + Val(.TextMatrix(m_end, 8))
            m_perp = m_perp + Val(.TextMatrix(m_end, 9))
            m_Anewc = m_Anewc + Val(.TextMatrix(m_end, 2))
            m_Anewp = m_Anewp + Val(.TextMatrix(m_end, 3))
            m_Aoldc = m_Aoldc + Val(.TextMatrix(m_end, 5))
            m_Aoldp = m_Aoldp + Val(.TextMatrix(m_end, 6))
            m_Aperc = m_Aperc + Val(.TextMatrix(m_end, 8))
            m_Aperp = m_Aperp + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '中二區
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "中二區" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "中二區" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If Val(.TextMatrix(m_end, 3)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 4) = "0.00%"
                End If
                If Val(.TextMatrix(m_end, 6)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 7) = "0.00%"
                End If
            Next m_i
            For m_i = m_std To m_end - 1
                If Val(.TextMatrix(m_i, 9)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_i, 10) = Format((Val(.TextMatrix(m_i, 9)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                     .TextMatrix(m_i, 10) = "0.00%"
                End If
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc = m_newc + Val(.TextMatrix(m_end, 2))
            m_newp = m_newp + Val(.TextMatrix(m_end, 3))
            m_oldc = m_oldc + Val(.TextMatrix(m_end, 5))
            m_oldp = m_oldp + Val(.TextMatrix(m_end, 6))
            m_perc = m_perc + Val(.TextMatrix(m_end, 8))
            m_perp = m_perp + Val(.TextMatrix(m_end, 9))
            m_Anewc = m_Anewc + Val(.TextMatrix(m_end, 2))
            m_Anewp = m_Anewp + Val(.TextMatrix(m_end, 3))
            m_Aoldc = m_Aoldc + Val(.TextMatrix(m_end, 5))
            m_Aoldp = m_Aoldp + Val(.TextMatrix(m_end, 6))
            m_Aperc = m_Aperc + Val(.TextMatrix(m_end, 8))
            m_Aperp = m_Aperp + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '中三區
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "中三區" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "中三區" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If Val(.TextMatrix(m_end, 3)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 4) = "0.00%"
                End If
                If Val(.TextMatrix(m_end, 6)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 7) = "0.00%"
                End If
            Next m_i
            For m_i = m_std To m_end - 1
                If Val(.TextMatrix(m_i, 9)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_i, 10) = Format((Val(.TextMatrix(m_i, 9)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                     .TextMatrix(m_i, 10) = "0.00%"
                End If
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc = m_newc + Val(.TextMatrix(m_end, 2))
            m_newp = m_newp + Val(.TextMatrix(m_end, 3))
            m_oldc = m_oldc + Val(.TextMatrix(m_end, 5))
            m_oldp = m_oldp + Val(.TextMatrix(m_end, 6))
            m_perc = m_perc + Val(.TextMatrix(m_end, 8))
            m_perp = m_perp + Val(.TextMatrix(m_end, 9))
            m_Anewc = m_Anewc + Val(.TextMatrix(m_end, 2))
            m_Anewp = m_Anewp + Val(.TextMatrix(m_end, 3))
            m_Aoldc = m_Aoldc + Val(.TextMatrix(m_end, 5))
            m_Aoldp = m_Aoldp + Val(.TextMatrix(m_end, 6))
            m_Aperc = m_Aperc + Val(.TextMatrix(m_end, 8))
            m_Aperp = m_Aperp + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '2010/7/30 ADD BY SONIA
        '中區其他
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "中區其他" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "中區其他" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If Val(.TextMatrix(m_end, 3)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 4) = "0.00%"
                End If
                If Val(.TextMatrix(m_end, 6)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 7) = "0.00%"
                End If
            Next m_i
            For m_i = m_std To m_end - 1
                If Val(.TextMatrix(m_i, 9)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_i, 10) = Format((Val(.TextMatrix(m_i, 9)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                     .TextMatrix(m_i, 10) = "0.00%"
                End If
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc = m_newc + Val(.TextMatrix(m_end, 2))
            m_newp = m_newp + Val(.TextMatrix(m_end, 3))
            m_oldc = m_oldc + Val(.TextMatrix(m_end, 5))
            m_oldp = m_oldp + Val(.TextMatrix(m_end, 6))
            m_perc = m_perc + Val(.TextMatrix(m_end, 8))
            m_perp = m_perp + Val(.TextMatrix(m_end, 9))
            m_Anewc = m_Anewc + Val(.TextMatrix(m_end, 2))
            m_Anewp = m_Anewp + Val(.TextMatrix(m_end, 3))
            m_Aoldc = m_Aoldc + Val(.TextMatrix(m_end, 5))
            m_Aoldp = m_Aoldp + Val(.TextMatrix(m_end, 6))
            m_Aperc = m_Aperc + Val(.TextMatrix(m_end, 8))
            m_Aperp = m_Aperp + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '2010/7/30 END
        '中所合計
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "中所合計" Then
                m_end = m_i
                m_seekst06 = m_seekst06 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "中所合計" Then
            .TextMatrix(m_end, 2) = m_newc
            .TextMatrix(m_end, 3) = m_newp
            .TextMatrix(m_end, 5) = m_oldc
            .TextMatrix(m_end, 6) = m_oldp
            .TextMatrix(m_end, 8) = m_perc
            .TextMatrix(m_end, 9) = m_perp
            If Val(.TextMatrix(m_end, 3)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
            Else
                .TextMatrix(m_end, 4) = "0.00%"
            End If
            If Val(.TextMatrix(m_end, 6)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
            Else
                .TextMatrix(m_end, 7) = "0.00%"
            End If
            For m_i = 0 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &HFF80FF   '&HFF00FF
            Next m_i
            m_std = m_end + 1
        End If
        m_newc = 0: m_newp = 0: m_oldc = 0: m_oldp = 0: m_perc = 0: m_perp = 0
        '南所合計
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "南所合計" Then
                m_end = m_i
                m_seekst06 = m_seekst06 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "南所合計" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If Val(.TextMatrix(m_end, 3)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 4) = "0.00%"
                End If
                If Val(.TextMatrix(m_end, 6)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 7) = "0.00%"
                End If
            Next m_i
            For m_i = m_std To m_end - 1
                If Val(.TextMatrix(m_i, 9)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_i, 10) = Format((Val(.TextMatrix(m_i, 9)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                     .TextMatrix(m_i, 10) = "0.00%"
                End If
            Next m_i
            For m_i = 0 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &HFF80FF   '&HFF00FF
            Next m_i
            m_Anewc = m_Anewc + Val(.TextMatrix(m_end, 2))
            m_Anewp = m_Anewp + Val(.TextMatrix(m_end, 3))
            m_Aoldc = m_Aoldc + Val(.TextMatrix(m_end, 5))
            m_Aoldp = m_Aoldp + Val(.TextMatrix(m_end, 6))
            m_Aperc = m_Aperc + Val(.TextMatrix(m_end, 8))
            m_Aperp = m_Aperp + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '高所合計
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "高所合計" Then
                m_end = m_i
                m_seekst06 = m_seekst06 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "高所合計" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If Val(.TextMatrix(m_end, 3)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 4) = "0.00%"
                End If
                If Val(.TextMatrix(m_end, 6)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                    .TextMatrix(m_end, 7) = "0.00%"
                End If
            Next m_i
            For m_i = m_std To m_end - 1
                If Val(.TextMatrix(m_i, 9)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                    .TextMatrix(m_i, 10) = Format((Val(.TextMatrix(m_i, 9)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                Else
                     .TextMatrix(m_i, 10) = "0.00%"
                End If
            Next m_i
            For m_i = 0 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &HFF80FF   '&HFF00FF
            Next m_i
            m_Anewc = m_Anewc + Val(.TextMatrix(m_end, 2))
            m_Anewp = m_Anewp + Val(.TextMatrix(m_end, 3))
            m_Aoldc = m_Aoldc + Val(.TextMatrix(m_end, 5))
            m_Aoldp = m_Aoldp + Val(.TextMatrix(m_end, 6))
            m_Aperc = m_Aperc + Val(.TextMatrix(m_end, 8))
            m_Aperp = m_Aperp + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '智權部
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "智權部" Then
                m_end = m_i
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "智權部" Then
            .TextMatrix(m_end, 2) = m_Anewc
            .TextMatrix(m_end, 3) = m_Anewp
            .TextMatrix(m_end, 5) = m_Aoldc
            .TextMatrix(m_end, 6) = m_Aoldp
            .TextMatrix(m_end, 8) = m_Aperc
            .TextMatrix(m_end, 9) = m_Aperp
            If Val(.TextMatrix(m_end, 3)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
            Else
                .TextMatrix(m_end, 4) = "0.00%"
            End If
            If Val(.TextMatrix(m_end, 6)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
            Else
                .TextMatrix(m_end, 7) = "0.00%"
            End If
            For m_i = 0 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &HFFFF80
            Next m_i
            m_Anewc = m_Anewc + Val(.TextMatrix(m_end, 2))
            m_Anewp = m_Anewp + Val(.TextMatrix(m_end, 3))
            m_Aoldc = m_Aoldc + Val(.TextMatrix(m_end, 5))
            m_Aoldp = m_Aoldp + Val(.TextMatrix(m_end, 6))
            m_Aperc = m_Aperc + Val(.TextMatrix(m_end, 8))
            m_Aperp = m_Aperp + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '其他
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 0) = "其他" Then
                m_end = m_i
                m_seekst06 = m_seekst06 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 0) = "其他" Then
            .TextMatrix(m_end, 1) = ""
            If Val(.TextMatrix(m_end, 3)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
            Else
                .TextMatrix(m_end, 4) = "0.00%"
            End If
            If Val(.TextMatrix(m_end, 6)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
            Else
                .TextMatrix(m_end, 7) = "0.00%"
            End If
            m_std = m_end + 1
        End If
        '國內部
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 0) = "國內部" Then
                m_end = m_i
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 0) = "國內部" Then
            .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end - 2, 2)) + Val(.TextMatrix(m_end - 1, 2))
            .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end - 2, 3)) + Val(.TextMatrix(m_end - 1, 3))
            .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end - 2, 5)) + Val(.TextMatrix(m_end - 1, 5))
            .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end - 2, 6)) + Val(.TextMatrix(m_end - 1, 6))
            .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end - 2, 8)) + Val(.TextMatrix(m_end - 1, 8))
            .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end - 2, 9)) + Val(.TextMatrix(m_end - 1, 9))
            If Val(.TextMatrix(m_end, 3)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
            Else
                .TextMatrix(m_end, 4) = "0.00%"
            End If
            If Val(.TextMatrix(m_end, 6)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
            Else
                .TextMatrix(m_end, 7) = "0.00%"
            End If
            For m_i = 0 To 11
                .row = m_end
                .col = m_i
                .CellBackColor = &H80C0FF
            Next m_i
        End If
        '2010/7/30 MODIFY BY SONIA
        'If m_stdArea = "S" And m_endArea = "S99" Then
        If (m_stdArea = "S" And m_endArea = "S99") Or (m_stdArea = "" And m_endArea = "") Then
        '2010/7/30 END
            'm_Std = m_End + 1
            m_tmp = Split(m_seekst06, ",")
            m_tmp2 = Split(m_seekst03, ",")
            For m_i = 0 To UBound(m_tmp)
                If Val(m_tmp(m_i)) <> 0 Then
                    For m_j = 0 To UBound(m_tmp2)
                        If Val(m_tmp2(m_j)) < Val(m_tmp(m_i)) And Val(m_tmp2(m_j)) <> 0 Then
                            If Val(.TextMatrix(m_tmp2(m_j), 9)) <> 0 And Val(.TextMatrix(m_tmp(m_i), 9)) <> 0 Then
                                .TextMatrix(m_tmp2(m_j), 11) = Format((Val(.TextMatrix(m_tmp2(m_j), 9)) / Val(.TextMatrix(m_tmp(m_i), 9))) * 100, "0.00") & "%"
                            Else
                                .TextMatrix(m_tmp2(m_j), 11) = "0.00%"
                            End If
                            m_tmp2(m_j) = 0
                        End If
                    Next m_j
                    .TextMatrix(m_tmp(m_i), 10) = "所佔國內"
                    If Val(.TextMatrix(m_tmp(m_i), 9)) <> 0 And Val(.TextMatrix(m_end, 9)) <> 0 Then
                        .TextMatrix(m_tmp(m_i), 11) = Format((Val(.TextMatrix(m_tmp(m_i), 9)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
                    Else
                        .TextMatrix(m_tmp(m_i), 11) = "0.00%"
                    End If
                End If
            Next m_i
        End If
    End With
    Grd1.Visible = True
Else
    InsertQueryLog (0) 'Add By Sindy 2010/12/23
    ShowNoData
    txtCloseDate(0).SetFocus
    Exit Sub
End If

'Add By Sindy 2010/9/23
Call ReadGrd2(stConST, STCONSTAREA, m_stdArea, m_endArea)
End Sub

'Add By Sindy 2010/9/23 抓Grd2資料, 新客戶數收文系統別分析
Private Sub ReadGrd2(stConST As String, STCONSTAREA As String, m_stdArea As String, m_endArea As String)
Dim m_str As String
Dim m_rs As New ADODB.Recordset
Dim m_gr1 As String
Dim m_gr2 As String
Dim m_newc_p As Double, m_newc_cfp As Double, m_newc_t As Double, m_newc_cft As Double
Dim m_newc_l As Double, m_newc_cfl As Double, m_newc_fcp As Double, m_newc_all As Double
Dim m_allnewc_p As Double, m_allnewc_cfp As Double, m_allnewc_t As Double, m_allnewc_cft As Double
Dim m_allnewc_l As Double, m_allnewc_cfl As Double, m_allnewc_fcp As Double, m_allnewc_all
Dim m_std As Integer
Dim m_end As Integer
Dim m_i As Integer
Dim m_j As Integer
Dim m_seekst03 As String
Dim m_seekst06 As String
Dim m_tmp As Variant
Dim m_tmp2 As Variant
Dim i_col As Integer
Dim m_st02 As String

'2014/1/21 MODIFY BY SONIA 取消 a0201='1' 條件
m_str = "select BB.st15,BB.AX209,BB.st02,BB.sysid,BB.newc,BB.oldc,a0902,decode(st.st06,'1','北所','2','中所','3','南所','4','高所','其他') st06 from ( "
m_str = m_str & " SELECT ST15,';',AX209,';',ST02,';',sysid,';',SUM(DECODE(TAG,'1',CC)) NewC,';',SUM(DECODE(TAG,'2',CC)) OldC FROM ("
m_str = m_str & " SELECT ST15,AX209,ST02,TAG,sysid,COUNT(DISTINCT AX208) CC FROM ("
m_str = m_str & " select distinct decode(substr(x.mindAtE,1,6), '','2',substr(cu14,1,6),'1','2') tag,x.mindAtE,cu14,st15,ax209,st02,ax202, ax203,ax208,ax214, to_char(round((ax207-ax206)/1000,2),'99999.00') Point,ax212,cp2.cp01 sysid"
m_str = m_str & " from acc020, acc021, staff,acc1p0,customer,caseprogress cp2,"
m_str = m_str & " (select AX208 CUNO,NVL(min(cp05||cp09),19110000) mindate from acc1u0,acc020,acc021,acc1p0,caseprogress"
m_str = m_str & " where ax201(+) = a0201  and ax202(+) = a0202 and ax209 Is Not Null " & stConST
'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
m_str = m_str & " and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and a0205 >= " & m_stdDay & " and a0205 <= " & m_endDay & " "
If Text2 = "N" Then m_str = m_str & " AND AX205 NOT IN ('4191','4192','4194') AND INSTR(AX213||' ','結餘')=0" '2010/7/29 ADD BY SONIA
m_str = m_str & " and ax202=a1p22(+) and ax203=a1p03(+) and a1p04=a1u01(+) and a1u03=cp09(+)"
 m_str = m_str & " group by aX208) x"
m_str = m_str & " where ax201(+) = a0201  and ax202(+) = a0202 and st01(+)=ax209 and ax209 Is Not Null " & stConST
'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
m_str = m_str & " and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and a0205 >= " & m_stdDay & " and a0205 <= " & m_endDay & STCONSTAREA
If Text2 = "N" Then m_str = m_str & " AND AX205 NOT IN ('4191','4192','4194') AND INSTR(AX213||' ','結餘')=0" '2010/7/29 ADD BY SONIA
m_str = m_str & " and substr(ax208,1,8)=cu01(+) and substr(ax208,9,1)=cu02(+)"
m_str = m_str & " and ax202=a1p22(+) and ax203=a1p03(+)"
m_str = m_str & " and ax208=x.CUNO(+) and cp2.cp09(+)=substr(x.mindate,9,9))"
m_str = m_str & " GROUP BY ST15,AX209,ST02,TAG,sysid) where TAG='1'"
m_str = m_str & " GROUP BY ST15,AX209,ST02,sysid union"
m_str = m_str & " SELECT '國內其他',';',' ',';','國內其他',';',sysid,';',SUM(DECODE(TAG,'1',CC)) NewC,';',SUM(DECODE(TAG,'2',CC)) OldC FROM ("
m_str = m_str & " SELECT TAG,sysid,COUNT(DISTINCT AX208) CC FROM ("
m_str = m_str & " select distinct decode(substr(x.mindAtE,1,6), '','2',substr(cu14,1,6),'1','2') tag,x.mindAtE,cu14,st15,ax209,st02,ax202, ax203,ax208,ax214, to_char(round((ax207-ax206)/1000,2),'99999.00') Point,ax212,cp2.cp01 sysid"
m_str = m_str & " from acc020, acc021, staff,acc1p0,customer,caseprogress cp2,"
m_str = m_str & " (select AX208 CUNO,NVL(min(cp05||cp09),19110000) mindate from acc1u0,acc020,acc021,acc1p0,caseprogress"
m_str = m_str & " where ax201(+) = a0201  and ax202(+) = a0202 and ax209 Is Not Null "
'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
m_str = m_str & " and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and a0205 >= " & m_stdDay & " and a0205 <= " & m_endDay & " "
If Text2 = "N" Then m_str = m_str & " AND AX205 NOT IN ('4191','4192','4194') AND INSTR(AX213||' ','結餘')=0" '2010/7/29 ADD BY SONIA
m_str = m_str & " and ax202=a1p22(+) and ax203=a1p03(+) and a1p04=a1u01(+) and a1u03=cp09(+)"
m_str = m_str & " group by aX208) x"
m_str = m_str & " where ax201(+) = a0201  and ax202(+) = a0202 and st01(+)=ax209 and ax209 Is Not Null "
'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
m_str = m_str & " and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and a0205 >= " & m_stdDay & " and a0205 <= " & m_endDay & " and SUBSTR(st15,1,1)<>'S'"
If Text2 = "N" Then m_str = m_str & " AND AX205 NOT IN ('4191','4192','4194') AND INSTR(AX213||' ','結餘')=0" '2010/7/29 ADD BY SONIA
m_str = m_str & " and substr(st15,1,1)<>'F' and substr(ax208,1,8)=cu01(+) and substr(ax208,9,1)=cu02(+)"
m_str = m_str & " and ax202=a1p22(+) and ax203=a1p03(+)"
m_str = m_str & " and ax208=x.CUNO(+) and cp2.cp09(+)=substr(x.mindate,9,9))"
m_str = m_str & " GROUP BY ST15,AX209,ST02,TAG,sysid) where TAG='1' GROUP BY sysid) BB,acc090,staff st  where BB.st15=a0901(+) and BB.AX209=st.st01(+) "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    Grd2.Visible = False
    Grd2.Clear
    Grd2.Rows = 2
    SetGrd2
    With m_rs
'塞資料
        .MoveFirst
        m_gr1 = ""
        m_gr2 = ""
        m_seekst03 = ""
        m_seekst06 = ""
        m_st02 = ""
        Do While Not .EOF
            If Grd2.TextMatrix(Grd2.Rows - 1, 0) <> "" Or Grd2.TextMatrix(Grd2.Rows - 1, 1) <> "" Or Grd2.TextMatrix(Grd2.Rows - 1, 2) <> "" Then
                Grd2.Rows = Grd2.Rows + 1
            End If
            If m_gr2 <> CheckStr(m_rs.Fields("a0902")) And m_gr2 <> "" And m_gr2 <> "台南所" And m_gr2 <> "高雄所" Then
                Grd2.TextMatrix(Grd2.Rows - 1, 1) = m_gr2
                m_gr2 = CheckStr(m_rs.Fields("a0902"))
                Grd2.Rows = Grd2.Rows + 1
            End If
            If m_gr1 <> CheckStr(m_rs.Fields("st06")) Then
                If m_gr1 <> "" Then
                    Grd2.TextMatrix(Grd2.Rows - 1, 1) = m_gr1 & "合計"
                    Grd2.Rows = Grd2.Rows + 1
                    If m_gr1 = "高所" Then
                        Grd2.TextMatrix(Grd2.Rows - 1, 1) = "智權部"
                        Grd2.Rows = Grd2.Rows + 1
                    End If
                End If
                If (m_stdArea = "S" And m_endArea = "S99") Or (m_stdArea = "" And m_endArea = "") Then
                    Grd2.TextMatrix(Grd2.Rows - 1, 0) = CheckStr(m_rs.Fields("st06"))
                    m_gr1 = CheckStr(m_rs.Fields("st06"))
                    m_gr2 = CheckStr(m_rs.Fields("a0902"))
                ElseIf CheckStr(m_rs.Fields("st06")) <> "其他" Then
                    Grd2.TextMatrix(Grd2.Rows - 1, 0) = CheckStr(m_rs.Fields("st06"))
                    m_gr1 = CheckStr(m_rs.Fields("st06"))
                    m_gr2 = CheckStr(m_rs.Fields("a0902"))
                End If
            End If
                If CheckStr(m_rs.Fields("sysid")) = "P" Or _
                   CheckStr(m_rs.Fields("sysid")) = "PS" Then
                   i_col = 2 'P
                ElseIf CheckStr(m_rs.Fields("sysid")) = "CFP" Or _
                          CheckStr(m_rs.Fields("sysid")) = "CPS" Then
                   i_col = 3 'CFP
                ElseIf CheckStr(m_rs.Fields("sysid")) = "CFT" Or _
                          CheckStr(m_rs.Fields("sysid")) = "CFC" Or _
                          CheckStr(m_rs.Fields("sysid")) = "S" Or _
                          CheckStr(m_rs.Fields("sysid")) = "FCT" Then
                   i_col = 5 'CFT
                ElseIf CheckStr(m_rs.Fields("sysid")) = "L" Or _
                          CheckStr(m_rs.Fields("sysid")) = "LA" Then
                   i_col = 6 'L
                ElseIf CheckStr(m_rs.Fields("sysid")) = "CFL" Or _
                          CheckStr(m_rs.Fields("sysid")) = "FCL" Or _
                          CheckStr(m_rs.Fields("sysid")) = "LIN" Then
                   i_col = 7 'CFL
                ElseIf CheckStr(m_rs.Fields("sysid")) = "FCP" Or _
                          CheckStr(m_rs.Fields("sysid")) = "FG" Then
                   i_col = 8 'FCP
                Else
                   i_col = 4 'T
                End If
                If m_st02 <> CheckStr(m_rs.Fields("st02")) Then
                  If CheckStr(m_rs.Fields("st02")) <> "國內其他" Then
                     Grd2.TextMatrix(Grd2.Rows - 1, 1) = CheckStr(m_rs.Fields("st02"))
                  End If
                  Grd2.TextMatrix(Grd2.Rows - 1, i_col) = Val(Grd2.TextMatrix(Grd2.Rows - 1, i_col)) + Val(CheckStr(m_rs.Fields("newc")))
               Else
                  Grd2.TextMatrix(Grd2.Rows - 2, i_col) = Val(Grd2.TextMatrix(Grd2.Rows - 2, i_col)) + Val(CheckStr(m_rs.Fields("newc")))
               End If
                m_st02 = CheckStr(m_rs.Fields("st02"))
            .MoveNext
        Loop
        If (m_stdArea = "S" And m_endArea = "S99") Or (m_stdArea = "" And m_endArea = "") Then
            'grd2.Rows = grd2.Rows + 1
            Grd2.TextMatrix(Grd2.Rows - 1, 0) = "國內部"
        End If
    End With
    If Grd2.TextMatrix(Grd2.Rows - 1, 1) = "" And Grd2.TextMatrix(Grd2.Rows - 1, 0) = "" Then
        Grd2.Rows = Grd2.Rows - 1
    End If
'統計及計算
    m_std = 1
    m_newc_p = 0
    m_newc_cfp = 0
    m_newc_t = 0
    m_newc_cft = 0
    m_newc_l = 0
    m_newc_cfl = 0
    m_newc_fcp = 0
    m_newc_all = 0
    m_allnewc_p = 0
    m_allnewc_cfp = 0
    m_allnewc_t = 0
    m_allnewc_cft = 0
    m_allnewc_l = 0
    m_allnewc_cfl = 0
    m_allnewc_fcp = 0
    m_allnewc_all = 0
    With Grd2
        '北一區
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "北一區" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "北一區" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_i, 9) = Val(.TextMatrix(m_i, 2)) + Val(.TextMatrix(m_i, 3)) + _
                                                  Val(.TextMatrix(m_i, 4)) + Val(.TextMatrix(m_i, 5)) + _
                                                  Val(.TextMatrix(m_i, 6)) + Val(.TextMatrix(m_i, 7)) + _
                                                  Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                If .TextMatrix(m_end, 2) = 0 Then .TextMatrix(m_end, 2) = ""
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                If .TextMatrix(m_end, 3) = 0 Then .TextMatrix(m_end, 3) = ""
                .TextMatrix(m_end, 4) = Val(.TextMatrix(m_end, 4)) + Val(.TextMatrix(m_i, 4))
                If .TextMatrix(m_end, 4) = 0 Then .TextMatrix(m_end, 4) = ""
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                If .TextMatrix(m_end, 5) = 0 Then .TextMatrix(m_end, 5) = ""
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                If .TextMatrix(m_end, 6) = 0 Then .TextMatrix(m_end, 6) = ""
                .TextMatrix(m_end, 7) = Val(.TextMatrix(m_end, 7)) + Val(.TextMatrix(m_i, 7))
                If .TextMatrix(m_end, 7) = 0 Then .TextMatrix(m_end, 7) = ""
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                If .TextMatrix(m_end, 8) = 0 Then .TextMatrix(m_end, 8) = ""
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If .TextMatrix(m_end, 9) = 0 Then .TextMatrix(m_end, 9) = ""
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc_p = m_newc_p + Val(.TextMatrix(m_end, 2))
            m_newc_cfp = m_newc_cfp + Val(.TextMatrix(m_end, 3))
            m_newc_t = m_newc_t + Val(.TextMatrix(m_end, 4))
            m_newc_cft = m_newc_cft + Val(.TextMatrix(m_end, 5))
            m_newc_l = m_newc_l + Val(.TextMatrix(m_end, 6))
            m_newc_cfl = m_newc_cfl + Val(.TextMatrix(m_end, 7))
            m_newc_fcp = m_newc_fcp + Val(.TextMatrix(m_end, 8))
            m_newc_all = m_newc_all + Val(.TextMatrix(m_end, 9))
            m_allnewc_p = m_allnewc_p + Val(.TextMatrix(m_end, 2))
            m_allnewc_cfp = m_allnewc_cfp + Val(.TextMatrix(m_end, 3))
            m_allnewc_t = m_allnewc_t + Val(.TextMatrix(m_end, 4))
            m_allnewc_cft = m_allnewc_cft + Val(.TextMatrix(m_end, 5))
            m_allnewc_l = m_allnewc_l + Val(.TextMatrix(m_end, 6))
            m_allnewc_cfl = m_allnewc_cfl + Val(.TextMatrix(m_end, 7))
            m_allnewc_fcp = m_allnewc_fcp + Val(.TextMatrix(m_end, 8))
            m_allnewc_all = m_allnewc_all + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '北三區
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "北三區" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "北三區" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_i, 9) = Val(.TextMatrix(m_i, 2)) + Val(.TextMatrix(m_i, 3)) + _
                                                  Val(.TextMatrix(m_i, 4)) + Val(.TextMatrix(m_i, 5)) + _
                                                  Val(.TextMatrix(m_i, 6)) + Val(.TextMatrix(m_i, 7)) + _
                                                  Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                If .TextMatrix(m_end, 2) = 0 Then .TextMatrix(m_end, 2) = ""
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                If .TextMatrix(m_end, 3) = 0 Then .TextMatrix(m_end, 3) = ""
                .TextMatrix(m_end, 4) = Val(.TextMatrix(m_end, 4)) + Val(.TextMatrix(m_i, 4))
                If .TextMatrix(m_end, 4) = 0 Then .TextMatrix(m_end, 4) = ""
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                If .TextMatrix(m_end, 5) = 0 Then .TextMatrix(m_end, 5) = ""
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                If .TextMatrix(m_end, 6) = 0 Then .TextMatrix(m_end, 6) = ""
                .TextMatrix(m_end, 7) = Val(.TextMatrix(m_end, 7)) + Val(.TextMatrix(m_i, 7))
                If .TextMatrix(m_end, 7) = 0 Then .TextMatrix(m_end, 7) = ""
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                If .TextMatrix(m_end, 8) = 0 Then .TextMatrix(m_end, 8) = ""
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If .TextMatrix(m_end, 9) = 0 Then .TextMatrix(m_end, 9) = ""
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc_p = m_newc_p + Val(.TextMatrix(m_end, 2))
            m_newc_cfp = m_newc_cfp + Val(.TextMatrix(m_end, 3))
            m_newc_t = m_newc_t + Val(.TextMatrix(m_end, 4))
            m_newc_cft = m_newc_cft + Val(.TextMatrix(m_end, 5))
            m_newc_l = m_newc_l + Val(.TextMatrix(m_end, 6))
            m_newc_cfl = m_newc_cfl + Val(.TextMatrix(m_end, 7))
            m_newc_fcp = m_newc_fcp + Val(.TextMatrix(m_end, 8))
            m_newc_all = m_newc_all + Val(.TextMatrix(m_end, 9))
            m_allnewc_p = m_allnewc_p + Val(.TextMatrix(m_end, 2))
            m_allnewc_cfp = m_allnewc_cfp + Val(.TextMatrix(m_end, 3))
            m_allnewc_t = m_allnewc_t + Val(.TextMatrix(m_end, 4))
            m_allnewc_cft = m_allnewc_cft + Val(.TextMatrix(m_end, 5))
            m_allnewc_l = m_allnewc_l + Val(.TextMatrix(m_end, 6))
            m_allnewc_cfl = m_allnewc_cfl + Val(.TextMatrix(m_end, 7))
            m_allnewc_fcp = m_allnewc_fcp + Val(.TextMatrix(m_end, 8))
            m_allnewc_all = m_allnewc_all + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '北四區
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "北四區" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "北四區" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_i, 9) = Val(.TextMatrix(m_i, 2)) + Val(.TextMatrix(m_i, 3)) + _
                                                  Val(.TextMatrix(m_i, 4)) + Val(.TextMatrix(m_i, 5)) + _
                                                  Val(.TextMatrix(m_i, 6)) + Val(.TextMatrix(m_i, 7)) + _
                                                  Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                If .TextMatrix(m_end, 2) = 0 Then .TextMatrix(m_end, 2) = ""
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                If .TextMatrix(m_end, 3) = 0 Then .TextMatrix(m_end, 3) = ""
                .TextMatrix(m_end, 4) = Val(.TextMatrix(m_end, 4)) + Val(.TextMatrix(m_i, 4))
                If .TextMatrix(m_end, 4) = 0 Then .TextMatrix(m_end, 4) = ""
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                If .TextMatrix(m_end, 5) = 0 Then .TextMatrix(m_end, 5) = ""
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                If .TextMatrix(m_end, 6) = 0 Then .TextMatrix(m_end, 6) = ""
                .TextMatrix(m_end, 7) = Val(.TextMatrix(m_end, 7)) + Val(.TextMatrix(m_i, 7))
                If .TextMatrix(m_end, 7) = 0 Then .TextMatrix(m_end, 7) = ""
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                If .TextMatrix(m_end, 8) = 0 Then .TextMatrix(m_end, 8) = ""
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If .TextMatrix(m_end, 9) = 0 Then .TextMatrix(m_end, 9) = ""
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc_p = m_newc_p + Val(.TextMatrix(m_end, 2))
            m_newc_cfp = m_newc_cfp + Val(.TextMatrix(m_end, 3))
            m_newc_t = m_newc_t + Val(.TextMatrix(m_end, 4))
            m_newc_cft = m_newc_cft + Val(.TextMatrix(m_end, 5))
            m_newc_l = m_newc_l + Val(.TextMatrix(m_end, 6))
            m_newc_cfl = m_newc_cfl + Val(.TextMatrix(m_end, 7))
            m_newc_fcp = m_newc_fcp + Val(.TextMatrix(m_end, 8))
            m_newc_all = m_newc_all + Val(.TextMatrix(m_end, 9))
            m_allnewc_p = m_allnewc_p + Val(.TextMatrix(m_end, 2))
            m_allnewc_cfp = m_allnewc_cfp + Val(.TextMatrix(m_end, 3))
            m_allnewc_t = m_allnewc_t + Val(.TextMatrix(m_end, 4))
            m_allnewc_cft = m_allnewc_cft + Val(.TextMatrix(m_end, 5))
            m_allnewc_l = m_allnewc_l + Val(.TextMatrix(m_end, 6))
            m_allnewc_cfl = m_allnewc_cfl + Val(.TextMatrix(m_end, 7))
            m_allnewc_fcp = m_allnewc_fcp + Val(.TextMatrix(m_end, 8))
            m_allnewc_all = m_allnewc_all + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '北五區
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "北五區" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "北五區" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_i, 9) = Val(.TextMatrix(m_i, 2)) + Val(.TextMatrix(m_i, 3)) + _
                                                  Val(.TextMatrix(m_i, 4)) + Val(.TextMatrix(m_i, 5)) + _
                                                  Val(.TextMatrix(m_i, 6)) + Val(.TextMatrix(m_i, 7)) + _
                                                  Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                If .TextMatrix(m_end, 2) = 0 Then .TextMatrix(m_end, 2) = ""
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                If .TextMatrix(m_end, 3) = 0 Then .TextMatrix(m_end, 3) = ""
                .TextMatrix(m_end, 4) = Val(.TextMatrix(m_end, 4)) + Val(.TextMatrix(m_i, 4))
                If .TextMatrix(m_end, 4) = 0 Then .TextMatrix(m_end, 4) = ""
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                If .TextMatrix(m_end, 5) = 0 Then .TextMatrix(m_end, 5) = ""
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                If .TextMatrix(m_end, 6) = 0 Then .TextMatrix(m_end, 6) = ""
                .TextMatrix(m_end, 7) = Val(.TextMatrix(m_end, 7)) + Val(.TextMatrix(m_i, 7))
                If .TextMatrix(m_end, 7) = 0 Then .TextMatrix(m_end, 7) = ""
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                If .TextMatrix(m_end, 8) = 0 Then .TextMatrix(m_end, 8) = ""
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If .TextMatrix(m_end, 9) = 0 Then .TextMatrix(m_end, 9) = ""
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc_p = m_newc_p + Val(.TextMatrix(m_end, 2))
            m_newc_cfp = m_newc_cfp + Val(.TextMatrix(m_end, 3))
            m_newc_t = m_newc_t + Val(.TextMatrix(m_end, 4))
            m_newc_cft = m_newc_cft + Val(.TextMatrix(m_end, 5))
            m_newc_l = m_newc_l + Val(.TextMatrix(m_end, 6))
            m_newc_cfl = m_newc_cfl + Val(.TextMatrix(m_end, 7))
            m_newc_fcp = m_newc_fcp + Val(.TextMatrix(m_end, 8))
            m_newc_all = m_newc_all + Val(.TextMatrix(m_end, 9))
            m_allnewc_p = m_allnewc_p + Val(.TextMatrix(m_end, 2))
            m_allnewc_cfp = m_allnewc_cfp + Val(.TextMatrix(m_end, 3))
            m_allnewc_t = m_allnewc_t + Val(.TextMatrix(m_end, 4))
            m_allnewc_cft = m_allnewc_cft + Val(.TextMatrix(m_end, 5))
            m_allnewc_l = m_allnewc_l + Val(.TextMatrix(m_end, 6))
            m_allnewc_cfl = m_allnewc_cfl + Val(.TextMatrix(m_end, 7))
            m_allnewc_fcp = m_allnewc_fcp + Val(.TextMatrix(m_end, 8))
            m_allnewc_all = m_allnewc_all + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '北所合計
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "北所合計" Then
                m_end = m_i
                m_seekst06 = m_seekst06 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "北所合計" Then
            .TextMatrix(m_end, 2) = m_newc_p
            If .TextMatrix(m_end, 2) = 0 Then .TextMatrix(m_end, 2) = ""
            .TextMatrix(m_end, 3) = m_newc_cfp
            If .TextMatrix(m_end, 3) = 0 Then .TextMatrix(m_end, 3) = ""
            .TextMatrix(m_end, 4) = m_newc_t
            If .TextMatrix(m_end, 4) = 0 Then .TextMatrix(m_end, 4) = ""
            .TextMatrix(m_end, 5) = m_newc_cft
            If .TextMatrix(m_end, 5) = 0 Then .TextMatrix(m_end, 5) = ""
            .TextMatrix(m_end, 6) = m_newc_l
            If .TextMatrix(m_end, 6) = 0 Then .TextMatrix(m_end, 6) = ""
            .TextMatrix(m_end, 7) = m_newc_cfl
            If .TextMatrix(m_end, 7) = 0 Then .TextMatrix(m_end, 7) = ""
            .TextMatrix(m_end, 8) = m_newc_fcp
            If .TextMatrix(m_end, 8) = 0 Then .TextMatrix(m_end, 8) = ""
            .TextMatrix(m_end, 9) = m_newc_all
            If .TextMatrix(m_end, 9) = 0 Then .TextMatrix(m_end, 9) = ""
            For m_i = 0 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &HFF80FF   '&HFF00FF
            Next m_i
            m_std = m_end + 1
        End If
        m_newc_p = 0: m_newc_cfp = 0: m_newc_t = 0: m_newc_cft = 0: m_newc_l = 0
        m_newc_cfl = 0: m_newc_fcp = 0: m_newc_all = 0
        '中一區
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "中一區" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "中一區" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_i, 9) = Val(.TextMatrix(m_i, 2)) + Val(.TextMatrix(m_i, 3)) + _
                                                  Val(.TextMatrix(m_i, 4)) + Val(.TextMatrix(m_i, 5)) + _
                                                  Val(.TextMatrix(m_i, 6)) + Val(.TextMatrix(m_i, 7)) + _
                                                  Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                If .TextMatrix(m_end, 2) = 0 Then .TextMatrix(m_end, 2) = ""
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                If .TextMatrix(m_end, 3) = 0 Then .TextMatrix(m_end, 3) = ""
                .TextMatrix(m_end, 4) = Val(.TextMatrix(m_end, 4)) + Val(.TextMatrix(m_i, 4))
                If .TextMatrix(m_end, 4) = 0 Then .TextMatrix(m_end, 4) = ""
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                If .TextMatrix(m_end, 5) = 0 Then .TextMatrix(m_end, 5) = ""
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                If .TextMatrix(m_end, 6) = 0 Then .TextMatrix(m_end, 6) = ""
                .TextMatrix(m_end, 7) = Val(.TextMatrix(m_end, 7)) + Val(.TextMatrix(m_i, 7))
                If .TextMatrix(m_end, 7) = 0 Then .TextMatrix(m_end, 7) = ""
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                If .TextMatrix(m_end, 8) = 0 Then .TextMatrix(m_end, 8) = ""
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If .TextMatrix(m_end, 9) = 0 Then .TextMatrix(m_end, 9) = ""
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc_p = m_newc_p + Val(.TextMatrix(m_end, 2))
            m_newc_cfp = m_newc_cfp + Val(.TextMatrix(m_end, 3))
            m_newc_t = m_newc_t + Val(.TextMatrix(m_end, 4))
            m_newc_cft = m_newc_cft + Val(.TextMatrix(m_end, 5))
            m_newc_l = m_newc_l + Val(.TextMatrix(m_end, 6))
            m_newc_cfl = m_newc_cfl + Val(.TextMatrix(m_end, 7))
            m_newc_fcp = m_newc_fcp + Val(.TextMatrix(m_end, 8))
            m_newc_all = m_newc_all + Val(.TextMatrix(m_end, 9))
            m_allnewc_p = m_allnewc_p + Val(.TextMatrix(m_end, 2))
            m_allnewc_cfp = m_allnewc_cfp + Val(.TextMatrix(m_end, 3))
            m_allnewc_t = m_allnewc_t + Val(.TextMatrix(m_end, 4))
            m_allnewc_cft = m_allnewc_cft + Val(.TextMatrix(m_end, 5))
            m_allnewc_l = m_allnewc_l + Val(.TextMatrix(m_end, 6))
            m_allnewc_cfl = m_allnewc_cfl + Val(.TextMatrix(m_end, 7))
            m_allnewc_fcp = m_allnewc_fcp + Val(.TextMatrix(m_end, 8))
            m_allnewc_all = m_allnewc_all + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '中二區
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "中二區" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "中二區" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_i, 9) = Val(.TextMatrix(m_i, 2)) + Val(.TextMatrix(m_i, 3)) + _
                                                  Val(.TextMatrix(m_i, 4)) + Val(.TextMatrix(m_i, 5)) + _
                                                  Val(.TextMatrix(m_i, 6)) + Val(.TextMatrix(m_i, 7)) + _
                                                  Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                If .TextMatrix(m_end, 2) = 0 Then .TextMatrix(m_end, 2) = ""
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                If .TextMatrix(m_end, 3) = 0 Then .TextMatrix(m_end, 3) = ""
                .TextMatrix(m_end, 4) = Val(.TextMatrix(m_end, 4)) + Val(.TextMatrix(m_i, 4))
                If .TextMatrix(m_end, 4) = 0 Then .TextMatrix(m_end, 4) = ""
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                If .TextMatrix(m_end, 5) = 0 Then .TextMatrix(m_end, 5) = ""
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                If .TextMatrix(m_end, 6) = 0 Then .TextMatrix(m_end, 6) = ""
                .TextMatrix(m_end, 7) = Val(.TextMatrix(m_end, 7)) + Val(.TextMatrix(m_i, 7))
                If .TextMatrix(m_end, 7) = 0 Then .TextMatrix(m_end, 7) = ""
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                If .TextMatrix(m_end, 8) = 0 Then .TextMatrix(m_end, 8) = ""
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If .TextMatrix(m_end, 9) = 0 Then .TextMatrix(m_end, 9) = ""
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc_p = m_newc_p + Val(.TextMatrix(m_end, 2))
            m_newc_cfp = m_newc_cfp + Val(.TextMatrix(m_end, 3))
            m_newc_t = m_newc_t + Val(.TextMatrix(m_end, 4))
            m_newc_cft = m_newc_cft + Val(.TextMatrix(m_end, 5))
            m_newc_l = m_newc_l + Val(.TextMatrix(m_end, 6))
            m_newc_cfl = m_newc_cfl + Val(.TextMatrix(m_end, 7))
            m_newc_fcp = m_newc_fcp + Val(.TextMatrix(m_end, 8))
            m_newc_all = m_newc_all + Val(.TextMatrix(m_end, 9))
            m_allnewc_p = m_allnewc_p + Val(.TextMatrix(m_end, 2))
            m_allnewc_cfp = m_allnewc_cfp + Val(.TextMatrix(m_end, 3))
            m_allnewc_t = m_allnewc_t + Val(.TextMatrix(m_end, 4))
            m_allnewc_cft = m_allnewc_cft + Val(.TextMatrix(m_end, 5))
            m_allnewc_l = m_allnewc_l + Val(.TextMatrix(m_end, 6))
            m_allnewc_cfl = m_allnewc_cfl + Val(.TextMatrix(m_end, 7))
            m_allnewc_fcp = m_allnewc_fcp + Val(.TextMatrix(m_end, 8))
            m_allnewc_all = m_allnewc_all + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '中三區
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "中三區" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "中三區" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_i, 9) = Val(.TextMatrix(m_i, 2)) + Val(.TextMatrix(m_i, 3)) + _
                                                  Val(.TextMatrix(m_i, 4)) + Val(.TextMatrix(m_i, 5)) + _
                                                  Val(.TextMatrix(m_i, 6)) + Val(.TextMatrix(m_i, 7)) + _
                                                  Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                If .TextMatrix(m_end, 2) = 0 Then .TextMatrix(m_end, 2) = ""
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                If .TextMatrix(m_end, 3) = 0 Then .TextMatrix(m_end, 3) = ""
                .TextMatrix(m_end, 4) = Val(.TextMatrix(m_end, 4)) + Val(.TextMatrix(m_i, 4))
                If .TextMatrix(m_end, 4) = 0 Then .TextMatrix(m_end, 4) = ""
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                If .TextMatrix(m_end, 5) = 0 Then .TextMatrix(m_end, 5) = ""
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                If .TextMatrix(m_end, 6) = 0 Then .TextMatrix(m_end, 6) = ""
                .TextMatrix(m_end, 7) = Val(.TextMatrix(m_end, 7)) + Val(.TextMatrix(m_i, 7))
                If .TextMatrix(m_end, 7) = 0 Then .TextMatrix(m_end, 7) = ""
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                If .TextMatrix(m_end, 8) = 0 Then .TextMatrix(m_end, 8) = ""
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If .TextMatrix(m_end, 9) = 0 Then .TextMatrix(m_end, 9) = ""
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc_p = m_newc_p + Val(.TextMatrix(m_end, 2))
            m_newc_cfp = m_newc_cfp + Val(.TextMatrix(m_end, 3))
            m_newc_t = m_newc_t + Val(.TextMatrix(m_end, 4))
            m_newc_cft = m_newc_cft + Val(.TextMatrix(m_end, 5))
            m_newc_l = m_newc_l + Val(.TextMatrix(m_end, 6))
            m_newc_cfl = m_newc_cfl + Val(.TextMatrix(m_end, 7))
            m_newc_fcp = m_newc_fcp + Val(.TextMatrix(m_end, 8))
            m_newc_all = m_newc_all + Val(.TextMatrix(m_end, 9))
            m_allnewc_p = m_allnewc_p + Val(.TextMatrix(m_end, 2))
            m_allnewc_cfp = m_allnewc_cfp + Val(.TextMatrix(m_end, 3))
            m_allnewc_t = m_allnewc_t + Val(.TextMatrix(m_end, 4))
            m_allnewc_cft = m_allnewc_cft + Val(.TextMatrix(m_end, 5))
            m_allnewc_l = m_allnewc_l + Val(.TextMatrix(m_end, 6))
            m_allnewc_cfl = m_allnewc_cfl + Val(.TextMatrix(m_end, 7))
            m_allnewc_fcp = m_allnewc_fcp + Val(.TextMatrix(m_end, 8))
            m_allnewc_all = m_allnewc_all + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '中區其他
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "中區其他" Then
                m_end = m_i
                m_seekst03 = m_seekst03 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "中區其他" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_i, 9) = Val(.TextMatrix(m_i, 2)) + Val(.TextMatrix(m_i, 3)) + _
                                                  Val(.TextMatrix(m_i, 4)) + Val(.TextMatrix(m_i, 5)) + _
                                                  Val(.TextMatrix(m_i, 6)) + Val(.TextMatrix(m_i, 7)) + _
                                                  Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                If .TextMatrix(m_end, 2) = 0 Then .TextMatrix(m_end, 2) = ""
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                If .TextMatrix(m_end, 3) = 0 Then .TextMatrix(m_end, 3) = ""
                .TextMatrix(m_end, 4) = Val(.TextMatrix(m_end, 4)) + Val(.TextMatrix(m_i, 4))
                If .TextMatrix(m_end, 4) = 0 Then .TextMatrix(m_end, 4) = ""
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                If .TextMatrix(m_end, 5) = 0 Then .TextMatrix(m_end, 5) = ""
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                If .TextMatrix(m_end, 6) = 0 Then .TextMatrix(m_end, 6) = ""
                .TextMatrix(m_end, 7) = Val(.TextMatrix(m_end, 7)) + Val(.TextMatrix(m_i, 7))
                If .TextMatrix(m_end, 7) = 0 Then .TextMatrix(m_end, 7) = ""
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                If .TextMatrix(m_end, 8) = 0 Then .TextMatrix(m_end, 8) = ""
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If .TextMatrix(m_end, 9) = 0 Then .TextMatrix(m_end, 9) = ""
            Next m_i
            For m_i = 1 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80FF80
            Next m_i
            m_newc_p = m_newc_p + Val(.TextMatrix(m_end, 2))
            m_newc_cfp = m_newc_cfp + Val(.TextMatrix(m_end, 3))
            m_newc_t = m_newc_t + Val(.TextMatrix(m_end, 4))
            m_newc_cft = m_newc_cft + Val(.TextMatrix(m_end, 5))
            m_newc_l = m_newc_l + Val(.TextMatrix(m_end, 6))
            m_newc_cfl = m_newc_cfl + Val(.TextMatrix(m_end, 7))
            m_newc_fcp = m_newc_fcp + Val(.TextMatrix(m_end, 8))
            m_newc_all = m_newc_all + Val(.TextMatrix(m_end, 9))
            m_allnewc_p = m_allnewc_p + Val(.TextMatrix(m_end, 2))
            m_allnewc_cfp = m_allnewc_cfp + Val(.TextMatrix(m_end, 3))
            m_allnewc_t = m_allnewc_t + Val(.TextMatrix(m_end, 4))
            m_allnewc_cft = m_allnewc_cft + Val(.TextMatrix(m_end, 5))
            m_allnewc_l = m_allnewc_l + Val(.TextMatrix(m_end, 6))
            m_allnewc_cfl = m_allnewc_cfl + Val(.TextMatrix(m_end, 7))
            m_allnewc_fcp = m_allnewc_fcp + Val(.TextMatrix(m_end, 8))
            m_allnewc_all = m_allnewc_all + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '中所合計
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "中所合計" Then
                m_end = m_i
                m_seekst06 = m_seekst06 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "中所合計" Then
            .TextMatrix(m_end, 2) = m_newc_p
            If .TextMatrix(m_end, 2) = 0 Then .TextMatrix(m_end, 2) = ""
            .TextMatrix(m_end, 3) = m_newc_cfp
            If .TextMatrix(m_end, 3) = 0 Then .TextMatrix(m_end, 3) = ""
            .TextMatrix(m_end, 4) = m_newc_t
            If .TextMatrix(m_end, 4) = 0 Then .TextMatrix(m_end, 4) = ""
            .TextMatrix(m_end, 5) = m_newc_cft
            If .TextMatrix(m_end, 5) = 0 Then .TextMatrix(m_end, 5) = ""
            .TextMatrix(m_end, 6) = m_newc_l
            If .TextMatrix(m_end, 6) = 0 Then .TextMatrix(m_end, 6) = ""
            .TextMatrix(m_end, 7) = m_newc_cfl
            If .TextMatrix(m_end, 7) = 0 Then .TextMatrix(m_end, 7) = ""
            .TextMatrix(m_end, 8) = m_newc_fcp
            If .TextMatrix(m_end, 8) = 0 Then .TextMatrix(m_end, 8) = ""
            .TextMatrix(m_end, 9) = m_newc_all
            If .TextMatrix(m_end, 9) = 0 Then .TextMatrix(m_end, 9) = ""
            For m_i = 0 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &HFF80FF   '&HFF00FF
            Next m_i
            m_std = m_end + 1
        End If
        m_newc_p = 0: m_newc_cfp = 0: m_newc_t = 0: m_newc_cft = 0: m_newc_l = 0
        m_newc_cfl = 0: m_newc_fcp = 0: m_newc_all = 0
        '南所合計
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "南所合計" Then
                m_end = m_i
                m_seekst06 = m_seekst06 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "南所合計" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_i, 9) = Val(.TextMatrix(m_i, 2)) + Val(.TextMatrix(m_i, 3)) + _
                                                  Val(.TextMatrix(m_i, 4)) + Val(.TextMatrix(m_i, 5)) + _
                                                  Val(.TextMatrix(m_i, 6)) + Val(.TextMatrix(m_i, 7)) + _
                                                  Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                If .TextMatrix(m_end, 2) = 0 Then .TextMatrix(m_end, 2) = ""
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                If .TextMatrix(m_end, 3) = 0 Then .TextMatrix(m_end, 3) = ""
                .TextMatrix(m_end, 4) = Val(.TextMatrix(m_end, 4)) + Val(.TextMatrix(m_i, 4))
                If .TextMatrix(m_end, 4) = 0 Then .TextMatrix(m_end, 4) = ""
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                If .TextMatrix(m_end, 5) = 0 Then .TextMatrix(m_end, 5) = ""
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                If .TextMatrix(m_end, 6) = 0 Then .TextMatrix(m_end, 6) = ""
                .TextMatrix(m_end, 7) = Val(.TextMatrix(m_end, 7)) + Val(.TextMatrix(m_i, 7))
                If .TextMatrix(m_end, 7) = 0 Then .TextMatrix(m_end, 7) = ""
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                If .TextMatrix(m_end, 8) = 0 Then .TextMatrix(m_end, 8) = ""
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If .TextMatrix(m_end, 9) = 0 Then .TextMatrix(m_end, 9) = ""
            Next m_i
            For m_i = 0 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &HFF80FF   '&HFF00FF
            Next m_i
            m_allnewc_p = m_allnewc_p + Val(.TextMatrix(m_end, 2))
            m_allnewc_cfp = m_allnewc_cfp + Val(.TextMatrix(m_end, 3))
            m_allnewc_t = m_allnewc_t + Val(.TextMatrix(m_end, 4))
            m_allnewc_cft = m_allnewc_cft + Val(.TextMatrix(m_end, 5))
            m_allnewc_l = m_allnewc_l + Val(.TextMatrix(m_end, 6))
            m_allnewc_cfl = m_allnewc_cfl + Val(.TextMatrix(m_end, 7))
            m_allnewc_fcp = m_allnewc_fcp + Val(.TextMatrix(m_end, 8))
            m_allnewc_all = m_allnewc_all + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '高所合計
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "高所合計" Then
                m_end = m_i
                m_seekst06 = m_seekst06 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "高所合計" Then
            For m_i = m_std To m_end - 1
                .TextMatrix(m_i, 9) = Val(.TextMatrix(m_i, 2)) + Val(.TextMatrix(m_i, 3)) + _
                                                  Val(.TextMatrix(m_i, 4)) + Val(.TextMatrix(m_i, 5)) + _
                                                  Val(.TextMatrix(m_i, 6)) + Val(.TextMatrix(m_i, 7)) + _
                                                  Val(.TextMatrix(m_i, 8))
                .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end, 2)) + Val(.TextMatrix(m_i, 2))
                If .TextMatrix(m_end, 2) = 0 Then .TextMatrix(m_end, 2) = ""
                .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end, 3)) + Val(.TextMatrix(m_i, 3))
                If .TextMatrix(m_end, 3) = 0 Then .TextMatrix(m_end, 3) = ""
                .TextMatrix(m_end, 4) = Val(.TextMatrix(m_end, 4)) + Val(.TextMatrix(m_i, 4))
                If .TextMatrix(m_end, 4) = 0 Then .TextMatrix(m_end, 4) = ""
                .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end, 5)) + Val(.TextMatrix(m_i, 5))
                If .TextMatrix(m_end, 5) = 0 Then .TextMatrix(m_end, 5) = ""
                .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end, 6)) + Val(.TextMatrix(m_i, 6))
                If .TextMatrix(m_end, 6) = 0 Then .TextMatrix(m_end, 6) = ""
                .TextMatrix(m_end, 7) = Val(.TextMatrix(m_end, 7)) + Val(.TextMatrix(m_i, 7))
                If .TextMatrix(m_end, 7) = 0 Then .TextMatrix(m_end, 7) = ""
                .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end, 8)) + Val(.TextMatrix(m_i, 8))
                If .TextMatrix(m_end, 8) = 0 Then .TextMatrix(m_end, 8) = ""
                .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end, 9)) + Val(.TextMatrix(m_i, 9))
                If .TextMatrix(m_end, 9) = 0 Then .TextMatrix(m_end, 9) = ""
            Next m_i
            For m_i = 0 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &HFF80FF   '&HFF00FF
            Next m_i
            m_allnewc_p = m_allnewc_p + Val(.TextMatrix(m_end, 2))
            m_allnewc_cfp = m_allnewc_cfp + Val(.TextMatrix(m_end, 3))
            m_allnewc_t = m_allnewc_t + Val(.TextMatrix(m_end, 4))
            m_allnewc_cft = m_allnewc_cft + Val(.TextMatrix(m_end, 5))
            m_allnewc_l = m_allnewc_l + Val(.TextMatrix(m_end, 6))
            m_allnewc_cfl = m_allnewc_cfl + Val(.TextMatrix(m_end, 7))
            m_allnewc_fcp = m_allnewc_fcp + Val(.TextMatrix(m_end, 8))
            m_allnewc_all = m_allnewc_all + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '智權部
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 1) = "智權部" Then
                m_end = m_i
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 1) = "智權部" Then
            .TextMatrix(m_end, 2) = m_allnewc_p
            If .TextMatrix(m_end, 2) = 0 Then .TextMatrix(m_end, 2) = ""
            .TextMatrix(m_end, 3) = m_allnewc_cfp
            If .TextMatrix(m_end, 3) = 0 Then .TextMatrix(m_end, 3) = ""
            .TextMatrix(m_end, 4) = m_allnewc_t
            If .TextMatrix(m_end, 4) = 0 Then .TextMatrix(m_end, 4) = ""
            .TextMatrix(m_end, 5) = m_allnewc_cft
            If .TextMatrix(m_end, 5) = 0 Then .TextMatrix(m_end, 5) = ""
            .TextMatrix(m_end, 6) = m_allnewc_l
            If .TextMatrix(m_end, 6) = 0 Then .TextMatrix(m_end, 6) = ""
            .TextMatrix(m_end, 7) = m_allnewc_cfl
            If .TextMatrix(m_end, 7) = 0 Then .TextMatrix(m_end, 7) = ""
            .TextMatrix(m_end, 8) = m_allnewc_fcp
            If .TextMatrix(m_end, 8) = 0 Then .TextMatrix(m_end, 8) = ""
            .TextMatrix(m_end, 9) = m_allnewc_all
            If .TextMatrix(m_end, 9) = 0 Then .TextMatrix(m_end, 9) = ""
            For m_i = 0 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &HFFFF80
            Next m_i
            m_allnewc_p = m_allnewc_p + Val(.TextMatrix(m_end, 2))
            m_allnewc_cfp = m_allnewc_cfp + Val(.TextMatrix(m_end, 3))
            m_allnewc_t = m_allnewc_t + Val(.TextMatrix(m_end, 4))
            m_allnewc_cft = m_allnewc_cft + Val(.TextMatrix(m_end, 5))
            m_allnewc_l = m_allnewc_l + Val(.TextMatrix(m_end, 6))
            m_allnewc_cfl = m_allnewc_cfl + Val(.TextMatrix(m_end, 7))
            m_allnewc_fcp = m_allnewc_fcp + Val(.TextMatrix(m_end, 8))
            m_allnewc_all = m_allnewc_all + Val(.TextMatrix(m_end, 9))
            m_std = m_end + 1
        End If
        '其他
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 0) = "其他" Then
                m_end = m_i
                m_seekst06 = m_seekst06 & m_end & ","
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 0) = "其他" Then
            .TextMatrix(m_end, 1) = ""
            .TextMatrix(m_i, 9) = Val(.TextMatrix(m_i, 2)) + Val(.TextMatrix(m_i, 3)) + _
                                              Val(.TextMatrix(m_i, 4)) + Val(.TextMatrix(m_i, 5)) + _
                                              Val(.TextMatrix(m_i, 6)) + Val(.TextMatrix(m_i, 7)) + _
                                              Val(.TextMatrix(m_i, 8))
            m_std = m_end + 1
        End If
        '國內部
        For m_i = m_std To .Rows - 1
            If .TextMatrix(m_i, 0) = "國內部" Then
                m_end = m_i
                Exit For
            End If
        Next m_i
        If .TextMatrix(m_end, 0) = "國內部" Then
            .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end - 2, 2)) + Val(.TextMatrix(m_end - 1, 2))
            If .TextMatrix(m_end, 2) = 0 Then .TextMatrix(m_end, 2) = ""
            .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end - 2, 3)) + Val(.TextMatrix(m_end - 1, 3))
            If .TextMatrix(m_end, 3) = 0 Then .TextMatrix(m_end, 3) = ""
            .TextMatrix(m_end, 4) = Val(.TextMatrix(m_end - 2, 4)) + Val(.TextMatrix(m_end - 1, 4))
            If .TextMatrix(m_end, 4) = 0 Then .TextMatrix(m_end, 4) = ""
            .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end - 2, 5)) + Val(.TextMatrix(m_end - 1, 5))
            If .TextMatrix(m_end, 5) = 0 Then .TextMatrix(m_end, 5) = ""
            .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end - 2, 6)) + Val(.TextMatrix(m_end - 1, 6))
            If .TextMatrix(m_end, 6) = 0 Then .TextMatrix(m_end, 6) = ""
            .TextMatrix(m_end, 7) = Val(.TextMatrix(m_end - 2, 7)) + Val(.TextMatrix(m_end - 1, 7))
            If .TextMatrix(m_end, 7) = 0 Then .TextMatrix(m_end, 7) = ""
            .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end - 2, 8)) + Val(.TextMatrix(m_end - 1, 8))
            If .TextMatrix(m_end, 8) = 0 Then .TextMatrix(m_end, 8) = ""
            .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end - 2, 9)) + Val(.TextMatrix(m_end - 1, 9))
            If .TextMatrix(m_end, 9) = 0 Then .TextMatrix(m_end, 9) = ""
            For m_i = 0 To 9
                .row = m_end
                .col = m_i
                .CellBackColor = &H80C0FF
            Next m_i
        End If
    End With
    Grd2.Visible = False
'Else
'    ShowNoData
'    txtCloseDate(0).SetFocus
End If
End Sub

Private Sub Form_Load()
   Me.Width = mdiMain.ScaleWidth
   Me.Height = mdiMain.ScaleHeight
   MoveFormToCenter Me
   
   cmdExit.Left = Me.Width - 200 - cmdExit.Width
   cmdExit.Top = 30
   cmdSearch.Left = cmdExit.Left - 200 - cmdSearch.Width
   cmdSearch.Top = 30
   cmdPrint.Left = cmdSearch.Left - 200 - cmdPrint.Width
   cmdPrint.Top = 30
   
   Grd1.Top = 1650 'edit by nickc 2008/04/21 加權限控制讓智權人員同仁可以使用 780
   Grd1.Left = 60
   Grd1.Width = Me.Width - (Grd1.Left * 4)
   Grd1.Height = Me.Height - Grd1.Top - 400
   
   stST15 = PUB_GetStaffST15(strUserNum, 1)
   stST05 = PUB_GetST05(strUserNum)
   
   'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
   Call PUB_SetFormSaleDept(strUserNum, txtZone, txtSalesArea, txtSalesArea1, txtSales, bolSpecMan, strSpecCode)
   
'   txtZone.Enabled = False
'   txtSalesArea.Enabled = False
'   txtSalesArea1.Enabled = False
'   txtSales.Enabled = False
'   Select Case strUserNum
'      '外商陳經理可看全所CFT,FCT,S,CFC
'      Case "68005"
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         txtSalesArea = stST15
'         txtSalesArea1 = stST15
''cancel by sonia 2014/6/9
''      '蔣律師可看中所全部
''      Case "79037"
''         txtZone = pub_strUserOffice
''         txtSalesArea.Enabled = True
''         txtSalesArea1.Enabled = True
''         txtSales.Enabled = True
''         txtSalesArea = "S2"
''         txtSalesArea1 = "S29"
''end 2014/6/9
'      '小真,杜副總可看全部
'      'modify by sonia 2014/6/9 +美珍77027
'      Case "65001", "68006", "77027"
'         txtZone.Enabled = True
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         '副總預設所有智權人員
'         If strUserNum = "68006" Then
'            txtSalesArea = "S"
'            txtSalesArea1 = "S99"
'         End If
'      '杜燕文,劉大愛可看S31
'      'modify by sonia 2016/8/22劉大愛78007改為蘇嫄媛79053
'      Case "74018", "79053"
'         txtZone = pub_strUserOffice
'         txtSalesArea = "S31"
'         txtSalesArea1 = "S31"
'         txtSales.Enabled = True
'      '王協理可看專利處
'      Case "71011"
'         txtSalesArea = "P10"
'         txtSalesArea1 = "P19"
'         txtSales.Enabled = True
'      '葉經理可看商標處
'      'modify by sonia 2016/2/24 +69008
'      Case "67002", "69008"
'         txtSalesArea = "P20"
'         txtSalesArea1 = "P29"
'         txtSales.Enabled = True
'      'add by sonia 2016/12/21 柄佑可看中所全部但業務區仍預設自已部門
'      Case "82026"
'         txtZone = pub_strUserOffice
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         txtSalesArea = stST15
'         txtSalesArea1 = stST15
'         txtSales = strUserNum
'      'end 2016/12/21
'      Case Else
'         Select Case stST05
'            '電腦中心,財務,總經理看全部
'            '2015/7/28 MODIFY BY SONIA +主任秘書(等級08)可看全部
'            Case "00", "01", "08"
'               txtZone.Enabled = True
'               txtSalesArea.Enabled = True
'               txtSalesArea1.Enabled = True
'               txtSales.Enabled = True
'            '各區主管
'            Case "SM"
'               txtZone = pub_strUserOffice
'               txtSalesArea.Locked = True
'               txtSalesArea1.Locked = True
'               '原羅文旭72009可兼看中一區,94/7/1只可看S22
'               '71003可看中所全部,但預設S23
'               If strUserNum = "71003" Then
'                  txtSalesArea = "S23"
'                  txtSalesArea1 = "S23"
'                  txtSalesArea.Locked = False
'                  txtSalesArea1.Locked = False
'                  txtSalesArea.Enabled = True
'                  txtSalesArea1.Enabled = True
'               Else
'               '簡協理可看北所全部但預設S15
'               If strUserNum = "69005" Then
'                  txtZone.Enabled = True 'Added by Morgan 2019/12/30 杜主秘說加開簡協理69005可看全所(預設S15)
'                  txtSalesArea = "S15"
'                  txtSalesArea1 = "S15"
'                  txtSalesArea.Locked = False
'                  txtSalesArea1.Locked = False
'                  txtSalesArea.Enabled = True
'                  txtSalesArea1.Enabled = True
'               Else
'                  txtSalesArea = stST15
'                  txtSalesArea1 = stST15
'               End If
'               End If
'               txtSales.Enabled = True
'            '外商主管  王宗珮、洪琬姿、葉易雲
'            Case "21", "26", "28"
'               txtZone = pub_strUserOffice
'               txtSalesArea = stST15
'               txtSalesArea1 = stST15
'               txtSales = strUserNum
'               txtSales.Enabled = True
'            '其他只能看自己
'            Case Else
'               txtZone = pub_strUserOffice
'               txtSalesArea = stST15
'               txtSalesArea1 = stST15
'               txtSales = strUserNum
'               'Added by Lydia 2017/07/25 多使用者權限,則增加部門範圍
'               strExc(1) = PUB_GetSalesList(strUserNum, , , , , strExc(2), strExc(3))
'               If strExc(3) <> "" And strExc(3) > txtSalesArea1 Then
'                  txtSalesArea1 = strExc(3)
'               End If
'               'end 2017/07/25
'         End Select
'   End Select
'
'   'Add By Sindy 2009/05/12
'   '若操作人員的ST05=SA,此類人員稱之為帶人主管,開放智權人員欄位可輸入
'   If PUB_GetST05Limits(strUserNum) = True Then
'      txtSales.Enabled = True
'   End If
'
'   txtSales = strUserNum
'
'   'Add By Sindy 2020/7/1 記錄原操作人可以查詢的業務區及所別
'   txtZone.Tag = txtZone
'   txtSalesArea.Tag = txtSalesArea
'   txtSalesArea1.Tag = txtSalesArea1
'   '2020/7/1 END

   Text1 = "1" '2010/7/29 ADD BY SONIA 加欄位且預設列印個人明細
   Text2 = "Y" '2010/7/29 ADD BY SONIA 加欄位且預設列印個人明細
   SetGrd
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm210120 = Nothing
End Sub

Private Sub SetGrd()
Grd1.Cols = 12
Grd1.row = 0
Grd1.col = 0: Grd1.Text = "所別"
Grd1.ColWidth(0) = 900
Grd1.CellAlignment = flexAlignCenterCenter
Grd1.col = 1: Grd1.Text = "智權同仁"
Grd1.ColWidth(1) = 900
Grd1.CellAlignment = flexAlignCenterCenter
Grd1.col = 2: Grd1.Text = "新客戶數"
Grd1.ColWidth(2) = 900
Grd1.CellAlignment = flexAlignCenterCenter
Grd1.col = 3: Grd1.Text = "新客戶點數"
Grd1.ColWidth(3) = 900
Grd1.CellAlignment = flexAlignCenterCenter
Grd1.col = 4: Grd1.Text = "(新)所佔比率"
Grd1.ColWidth(4) = 900
Grd1.CellAlignment = flexAlignCenterCenter
Grd1.col = 5: Grd1.Text = "舊客戶數"
Grd1.ColWidth(5) = 900
Grd1.CellAlignment = flexAlignCenterCenter
Grd1.col = 6: Grd1.Text = "舊客戶點數"
Grd1.ColWidth(6) = 900
Grd1.CellAlignment = flexAlignCenterCenter
Grd1.col = 7: Grd1.Text = "(舊)所佔比率"
Grd1.ColWidth(7) = 900
Grd1.CellAlignment = flexAlignCenterCenter
Grd1.col = 8: Grd1.Text = "客戶小計" 'edit by nickc 2008/04/02 副總說要改 "個人客戶數"
Grd1.ColWidth(8) = 900
Grd1.CellAlignment = flexAlignCenterCenter
Grd1.col = 9: Grd1.Text = "點數小計" 'edit by nickc 2008/04/02 副總說要改 "個人點數"
Grd1.ColWidth(9) = 900
Grd1.CellAlignment = flexAlignCenterCenter
Grd1.col = 10: Grd1.Text = "個人佔區"
Grd1.ColWidth(10) = 900
Grd1.CellAlignment = flexAlignCenterCenter
Grd1.col = 11: Grd1.Text = "區佔該所"
Grd1.ColWidth(11) = 900
Grd1.CellAlignment = flexAlignCenterCenter
End Sub

Private Sub SetGrd2()
Grd2.Cols = 10
Grd2.row = 0
Grd2.col = 0: Grd2.Text = "所別"
Grd2.ColWidth(0) = 900
Grd2.CellAlignment = flexAlignCenterCenter
Grd2.col = 1: Grd2.Text = "智權同仁"
Grd2.ColWidth(1) = 900
Grd2.CellAlignment = flexAlignCenterCenter
Grd2.col = 2: Grd2.Text = "P"
Grd2.ColWidth(2) = 900
Grd2.CellAlignment = flexAlignCenterCenter
Grd2.col = 3: Grd2.Text = "CFP"
Grd2.ColWidth(3) = 900
Grd2.CellAlignment = flexAlignCenterCenter
Grd2.col = 4: Grd2.Text = "T"
Grd2.ColWidth(4) = 900
Grd2.CellAlignment = flexAlignCenterCenter
Grd2.col = 5: Grd2.Text = "CFT"
Grd2.ColWidth(5) = 900
Grd2.CellAlignment = flexAlignCenterCenter
Grd2.col = 6: Grd2.Text = "L"
Grd2.ColWidth(6) = 900
Grd2.CellAlignment = flexAlignCenterCenter
Grd2.col = 7: Grd2.Text = "CFL"
Grd2.ColWidth(7) = 900
Grd2.CellAlignment = flexAlignCenterCenter
Grd2.col = 8: Grd2.Text = "FCP"
Grd2.ColWidth(8) = 900
Grd2.CellAlignment = flexAlignCenterCenter
Grd2.col = 9: Grd2.Text = "合計"
Grd2.ColWidth(9) = 900
Grd2.CellAlignment = flexAlignCenterCenter
End Sub

'2010/7/29 ADD BY SONIA
Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
      KeyAscii = 0
   End If
End Sub
Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = "" Then
      MsgBox "請輸入列印內容！", vbCritical
      Cancel = True
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> 89 And KeyAscii <> 78 Then
      KeyAscii = 0
   End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 = "" Then
      MsgBox "請輸入是否含結餘及保留！", vbCritical
      Cancel = True
   End If
End Sub
'2010/7/29 END

Private Sub txtSales_Change()
   If Len(txtSales) > 4 Then
      lblSalesName = GetStaffName(txtSales)
   Else
      lblSalesName = ""
   End If
End Sub

Private Sub txtCloseDate_GotFocus(Index As Integer)
   TextInverse txtCloseDate(Index)
   CloseIme
End Sub

Private Sub txtCloseDate_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 9 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtCloseDate_Validate(Index As Integer, Cancel As Boolean)
   If txtCloseDate(Index) <> "" Then
      If ChkDate(txtCloseDate(Index)) = False Then
         Cancel = True
         txtCloseDate(Index).SetFocus
         txtCloseDate_GotFocus Index
         Exit Sub
      End If
      If Index = 1 Then
        If RunNick2(txtCloseDate(0), txtCloseDate(1)) = True Then
           txtCloseDate(Index).SetFocus
           txtCloseDate_GotFocus Index
           Cancel = True
           Exit Sub
        End If
     End If
   End If
End Sub

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   CloseIme
End Sub

'Add By Sindy 2010/11/26
Private Sub txtSales_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_Validate(Cancel As Boolean)
   'Modify By Sindy 2024/9/18 智權部:使用智權人員ID查詢權限; 改為共用函數
   If PUB_txtSales_Limit(txtSales, m_strListPer, txtZone, txtSalesArea, txtSalesArea1, _
                         bolSpecMan, strSpecCode, lblSalesName) = False Then
      If txtSales.Visible = True Then 'Added by Lydia 2021/05/20 排除隱藏
        txtSales.SetFocus
        txtSales_GotFocus
      End If 'Added by Lydia 2021/05/20
      Cancel = True
      Exit Sub
   End If
   '2024/9/18 END
   
'   'Add By Sindy 2015/6/26 若有異動智權人員,需重新查詢業務區和所別
'   'modify by sonia 2016/6/15 加入帶人主管條件
'   'If txtSalesArea.Enabled = True Then 'Modify By Sindy 2016/5/5 + if
'   If txtSalesArea.Enabled = True Or PUB_GetST05Limits(strUserNum) = True Then
'      If txtSales.Text <> "" And txtSales.Text <> txtSales.Tag Then
'         txtZone = PUB_GetST06(Trim(txtSales))
'         txtSalesArea = PUB_GetStaffST15(Trim(txtSales), "1")
'         txtSalesArea1 = PUB_GetStaffST15(Trim(txtSales), "1")
'      End If
'   Else
'      'Add By Sindy 2016/5/6 還原(原操作人)可以查詢的業務區及所別
'      txtZone = txtZone.Tag
'      txtSalesArea = txtSalesArea.Tag
'      txtSalesArea1 = txtSalesArea1.Tag
'      '2016/5/6 END
'   End If
'
''   'add by sonia 2016/6/7 S29
''   If Len(txtSales) <= 4 Then
''      txtSales.Text = Mid(txtSales.Text & "  ", 1, 5)
''   End If
''   'end 2016/6/7
'
'   'Add by Amy 2017/01/12 +MCTF人員
'   'Remove by Lydia 2017/07/21 併入PUB_GetSalesList
'   'strMCTF = GetMCTF0XCode(txtSales)
'
'   txtSales.Tag = txtSales.Text
'
'   'add by sonia 2016/12/21 取消智權人員編號時,無跨所別權限者重新預設所別
'   If Trim(txtSales) = "" And txtZone.Enabled = False Then
'      txtZone = pub_strUserOffice
'   End If
'   'end 2016/12/21
End Sub

Private Sub txtSalesArea_GotFocus()
   TextInverse txtSalesArea
   CloseIme
End Sub

Private Sub txtSalesArea_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_GotFocus()
   TextInverse txtSalesArea1
   CloseIme
End Sub

Private Sub txtSalesArea1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_Validate(Cancel As Boolean)
If Trim(txtSalesArea1) <> "" Then
   If RunNick(txtSalesArea, txtSalesArea1) = True Then
      Cancel = True
      Exit Sub
   End If
End If
End Sub

Private Sub txtZone_GotFocus()
   TextInverse txtZone
   CloseIme
End Sub

Private Sub txtZone_KeyPress(KeyAscii As Integer)
   If (KeyAscii < Asc("1") Or KeyAscii > Asc("4")) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

' 抓離職智權人員及虛建智權人員
'Remove by Lydia 2017/07/24 改成共用模組,已不使用
'Function GetNotInOfficeAndFalseStaff(oStr As String, oStr2 As String) As String
'GetNotInOfficeAndFalseStaff = ""
'Dim rsTmp2 As New ADODB.Recordset
'Dim sqlTmp2 As String
'sqlTmp2 = "select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='2' "
'sqlTmp2 = sqlTmp2 & "union select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='1' and st01<'6' "
'Select Case strUserNum
'   Case "71011"  '王協理
'      'edit by nickc 2008/04/24
'      'sqlTmp2 = sqlTmp2 & "union select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='1' and st01='96030' "
'      sqlTmp2 = sqlTmp2 & "union select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='1' and st01 in ('96031','96032') "
'
'   Case "67002" '葉經理
'      'edit by nickc 2008/04/24
'      'sqlTmp2 = sqlTmp2 & "union select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='1' and st01='96029' "
'      sqlTmp2 = sqlTmp2 & "union select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='1' and st01 in ('96029','96030') "
'   Case Else
'End Select
'Set rsTmp2 = New ADODB.Recordset
'With rsTmp2
'    .CursorLocation = adUseClient
'    .Open sqlTmp2, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 Then
'        .MoveFirst
'        Do While Not .EOF
'            GetNotInOfficeAndFalseStaff = GetNotInOfficeAndFalseStaff & "'" & CheckStr(.Fields(0)) & "',"
'            .MoveNext
'        Loop
'    End If
'End With
'Set rsTmp2 = Nothing
'End Function
'end 2017/07/24
