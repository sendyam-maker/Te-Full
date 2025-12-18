VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040301 
   BorderStyle     =   1  '單線固定
   Caption         =   "公告通知函"
   ClientHeight    =   5796
   ClientLeft      =   9900
   ClientTop       =   960
   ClientWidth     =   6360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5796
   ScaleWidth      =   6360
   Begin VB.OptionButton Option1 
      Caption         =   "Printer.TrackDefault = True"
      Height          =   405
      Index           =   1
      Left            =   4890
      TabIndex        =   34
      Top             =   5220
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Printer.TrackDefault = False"
      Height          =   405
      Index           =   0
      Left            =   4890
      TabIndex        =   33
      Top             =   4740
      Width           =   1395
   End
   Begin VB.TextBox txtByte 
      Height          =   315
      Left            =   1680
      TabIndex        =   28
      Text            =   "30000"
      Top             =   4710
      Width           =   975
   End
   Begin VB.TextBox txtMinSec 
      Height          =   315
      Left            =   1680
      TabIndex        =   27
      Text            =   "5"
      Top             =   5070
      Width           =   555
   End
   Begin VB.TextBox txtMaxSec 
      Height          =   315
      Left            =   4260
      TabIndex        =   26
      Text            =   "45"
      Top             =   5070
      Width           =   555
   End
   Begin VB.TextBox txtFirstAdd 
      Height          =   315
      Left            =   4260
      TabIndex        =   25
      Text            =   "3"
      Top             =   4710
      Width           =   555
   End
   Begin VB.ComboBox cmbPrinter2 
      Height          =   300
      Left            =   1860
      TabIndex        =   12
      Top             =   2790
      Width           =   4395
   End
   Begin VB.TextBox txtPDFPath 
      Height          =   315
      Left            =   1860
      TabIndex        =   11
      Text            =   "C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe"
      Top             =   2400
      Width           =   4395
   End
   Begin VB.ListBox List1 
      Height          =   948
      ItemData        =   "frm040301.frx":0000
      Left            =   60
      List            =   "frm040301.frx":0007
      TabIndex        =   15
      Top             =   3600
      Width           =   6225
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   10
      Left            =   3240
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1104
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   9
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1104
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   8
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1404
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   7
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   2
      Top             =   804
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   0
      Left            =   3240
      MaxLength       =   7
      TabIndex        =   3
      Top             =   804
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   5
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   0
      Top             =   504
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   1
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "P"
      Top             =   1728
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   2
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1728
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   3
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1728
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   4
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   10
      Top             =   1728
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   6
      Left            =   3240
      MaxLength       =   7
      TabIndex        =   1
      Top             =   504
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   5244
      TabIndex        =   14
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   4356
      TabIndex        =   13
      Top             =   120
      Width           =   800
   End
   Begin VB.Label Label4 
      Caption         =   "程序人員："
      Height          =   240
      Left            =   96
      TabIndex        =   36
      Top             =   168
      Visible         =   0   'False
      Width           =   972
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1440
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   2088
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3678;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "幾Byte算1秒："
      Height          =   180
      Index           =   5
      Left            =   510
      TabIndex        =   32
      Top             =   4710
      Width           =   1140
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "1個檔至少幾秒："
      Height          =   180
      Left            =   300
      TabIndex        =   31
      Top             =   5070
      Width           =   1350
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "1個檔最多幾秒："
      Height          =   180
      Left            =   2880
      TabIndex        =   30
      Top             =   5070
      Width           =   1350
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "第1個檔多加幾秒："
      Height          =   180
      Left            =   2700
      TabIndex        =   29
      Top             =   4710
      Width           =   1530
   End
   Begin VB.Label Label6 
      Caption         =   "列印公報PDF印表機："
      Height          =   180
      Left            =   90
      TabIndex        =   24
      Top             =   2850
      Width           =   1755
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PDF執行檔路徑："
      Height          =   180
      Left            =   90
      TabIndex        =   23
      Top             =   2460
      Width           =   1560
   End
   Begin VB.Label Label2 
      Caption         =   "備註：台灣案只印無證書案件"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   90
      TabIndex        =   22
      Top             =   3270
      Width           =   3885
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2760
      X2              =   2880
      Y1              =   1224
      Y2              =   1224
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公告號/證書號："
      Height          =   180
      Index           =   4
      Left            =   96
      TabIndex        =   21
      Top             =   1128
      Width           =   1308
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   3
      Left            =   96
      TabIndex        =   20
      Top             =   1788
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公告日："
      Height          =   180
      Index           =   2
      Left            =   96
      TabIndex        =   19
      Top             =   528
      Width           =   720
   End
   Begin VB.Label lbl 
      Height          =   180
      Index           =   0
      Left            =   2016
      TabIndex        =   18
      Top             =   1464
      Width           =   2172
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利種類："
      Height          =   180
      Index           =   0
      Left            =   96
      TabIndex        =   17
      Top             =   1464
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2760
      X2              =   2880
      Y1              =   924
      Y2              =   924
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2760
      X2              =   2880
      Y1              =   624
      Y2              =   624
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Index           =   1
      Left            =   96
      TabIndex        =   16
      Top             =   840
      Width           =   900
   End
End
Attribute VB_Name = "frm040301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modified by Morgan 2025/1/15 增加程序人員選單並刪除不再使用的物件及部分舊程式碼
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim intWhere As Integer, strReceiveNo As String
Const ET01 As String = "10"
'Modify By Sindy 2014/9/3
Dim m_DefaultPrinter As String
'Dim m_DefaultPrinter2 As String
Dim strPrinter As String
'2014/9/3 END
'Add By Cheng 2002/09/10
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
'Add By Cheng 2003/01/14
Dim m_strSQLA As String '地址條列印時使用
'Add By Cheng 2003/04/03
Dim m_PrtOrientation As Integer '列印方向
Dim m_PrtScaleMode As Integer '列印座標單位
Dim m_dblTop As Double '上邊界
Dim m_dblLeft As Double '左邊界
Dim m_dblTitleHeight As Double '表頭高度
Dim m_dblLine As Double '行數
Dim m_dblLineHeight As Double '行高
Dim m_dblBetweenLine As Double '行間空隙
Dim m_dblLineHeight1 As Double '行高
Dim m_dblBetweenLine1 As Double '行間空隙
Dim m_strSQLB As String
Dim m_rsA As New ADODB.Recordset
Dim iPage As Integer, iCnt As Integer, iPrint As Long
'Dim SeekPrint As Integer
'Add By Sindy 2011/12/22
Dim strTPB04 As String, strTPB05 As String
Dim i As Integer, j As Integer
Dim strTime As String
Dim FF2 As Integer
'2011/12/22 End
Dim m_bolELetter As Boolean 'Added by Morgan 2014/6/19 是否有存電子信函
Dim m_AttachPath As String 'Added by Morgan 2021/6/24 公報PDF暫存路徑
'Added by Morgan 2025/1/15
Dim rsQuery As ADODB.Recordset
Dim mSeqNo As String, stVTBX As String
'end 2025/1/15

Private Sub PrintFooter()
      iPrint = iPrint + 300
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print String(200, "-")
      iPrint = iPrint + 300
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print "待續..."
End Sub

'Add by Morgan 2006/10/13
'延緩公告申請與公告日不符清單
Private Sub Print412Head(ByRef iPage As Integer, ByRef iCnt As Integer, ByRef iPrint As Long)

On Error GoTo ErrHnd

   If iPage = 0 Then
      Printer.Orientation = 1  '直印
      Printer.Font.Name = "細明體"
   Else
      iPrint = iPrint + 300
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print String(200, "-")
      Printer.NewPage
   End If
   
   iCnt = 1
   iPage = iPage + 1
   iPrint = 500
   Printer.CurrentX = 3000
   Printer.CurrentY = iPrint
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.Print "延緩公告申請與公告日不符清單"
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   iPrint = iPrint + 800
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 8500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "公告日：" & Text1(5).Text & " － " & Text1(6).Text
   Printer.CurrentX = 8500
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(iPage)
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "　　　　        　　　　     延緩公告  預定                領證      延緩公告  准予延緩"
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "本所案號        申請案號     發文日    公告日    公告日    發文日    月數/日期 公告來函日"
     
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   
   Exit Sub
        
ErrHnd:

   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Add by Morgan 2004/9/20 '列印可承辦技術報告清單
'iPage 頁次,iCnt 筆次, iPrint 目前列印位置
Private Sub Print421Head(ByRef iPage As Integer, ByRef iCnt As Integer, ByRef iPrint As Long)

On Error GoTo ErrHnd

   If iPage = 0 Then
      Printer.Orientation = 1  '直印
      Printer.Font.Name = "細明體"
   Else
      iPrint = iPrint + 300
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print String(200, "-")
      Printer.NewPage
   End If
   
   iCnt = 1
   iPage = iPage + 1
   iPrint = 500
   Printer.CurrentX = 3700
   Printer.CurrentY = iPrint
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.Print "可承辦技術報告清單"
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   iPrint = iPrint + 800
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 8500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "公告日：" & Text1(5).Text & " － " & Text1(6).Text
   Printer.CurrentX = 8500
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(iPage)
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = 2100
   Printer.CurrentY = iPrint
   Printer.Print "收文日"
   Printer.CurrentX = 3700
   Printer.CurrentY = iPrint
   Printer.Print "申請案號"
   Printer.CurrentX = 5300
   Printer.CurrentY = iPrint
   Printer.Print "公告日"
   Printer.CurrentX = 6900
   Printer.CurrentY = iPrint
   Printer.Print "承辦人"
     
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
        
   Exit Sub
   
ErrHnd:

   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Add by Morgan 2004/9/20 '列印可承辦技術報告清單
Private Sub Print421List(ByRef rstResult As ADODB.Recordset)
   
On Error GoTo ErrHnd
   
   iPage = 0: iCnt = 0: iPrint = 0
'   Print #FF2, "***With rstResult"
   With rstResult
      Print421Head iPage, iCnt, iPrint
      .MoveFirst
'      Print #FF2, "***Do While Not .EOF"
      Do While Not .EOF
         iCnt = iCnt + 1
         If iCnt > 35 Then
            PrintFooter
            Print421Head iPage, iCnt, iPrint
         End If
         iPrint = iPrint + 300
         Printer.CurrentX = 500
         Printer.CurrentY = iPrint
         Printer.Print .Fields("X01")  '"本所案號"
         Printer.CurrentX = 2100
         Printer.CurrentY = iPrint
         Printer.Print ChangeTStringToTDateString(ChangeWStringToTString("" & .Fields("CP05")))  '"收文日"
         Printer.CurrentX = 3700
         Printer.CurrentY = iPrint
         Printer.Print "" & .Fields("PA11") '"申請案號"
         Printer.CurrentX = 5300
         Printer.CurrentY = iPrint
         Printer.Print ChangeTStringToTDateString(ChangeWStringToTString("" & .Fields("PA14"))) '"公告日"
         Printer.CurrentX = 6900
         Printer.CurrentY = iPrint
         Printer.Print "" & .Fields("ST02") '"承辦人"
         .MoveNext
      Loop
      iPrint = iPrint + 300
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print String(200, "-")
      iPrint = iPrint + 300
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print "筆數：共" & Format(.RecordCount, "#,##0") & "筆!!!"
'      Print #FF2, "***Printer.EndDoc Star"
      Printer.EndDoc
'      Print #FF2, "***Printer.EndDoc End"
   End With
   
'   Print #FF2, "***Exit Sub"
   Exit Sub
      
ErrHnd:
'      Print #FF2, "***Print421List ErrHnd"
'      Print #FF2, "***Err.NUMBER : " & Err.NUMBER
'      Print #FF2, "***Err.Description : " & Err.Description
      If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Add by Morgan 2004/9/20 '延緩公告申請與公告日不符清單
Private Sub Print412List(ByRef rstResult As ADODB.Recordset)
   
   Dim strContent As String
   
On Error GoTo ErrHnd
   
   iPage = 0: iCnt = 0: iPrint = 0
   With rstResult
      Print412Head iPage, iCnt, iPrint
      .MoveFirst
      Do While Not .EOF
         iCnt = iCnt + 1
         If iCnt > 35 Then
            PrintFooter
            Print421Head iPage, iCnt, iPrint
         End If
         iPrint = iPrint + 300
         Printer.CurrentX = 500
         Printer.CurrentY = iPrint
         strContent = Empty
         strContent = strContent & convForm(.Fields("C1"), 15) '本所案號
         strContent = strContent & " " & convForm(.Fields("C2"), 12) '申請案號
         strContent = strContent & " " & convForm(Format(.Fields("C3"), "@@@/@@/@@"), 9) '延緩公告發文日
         strContent = strContent & " " & convForm(Format(.Fields("C4"), "@@@/@@/@@"), 9) '預定公告日
         strContent = strContent & " " & convForm(Format(.Fields("C8"), "@@@/@@/@@"), 9) '實際公告日
         strContent = strContent & " " & convForm(Format(.Fields("C5"), "@@@/@@/@@"), 9) '領證發文日
         strContent = strContent & " " & convForm(.Fields("C6"), 9) '延緩公告月數/日期
         strContent = strContent & " " & convForm(Format(.Fields("C7"), "@@@/@@/@@"), 9) '准予延緩公告來函日
         Printer.Print strContent
         .MoveNext
      Loop
      iPrint = iPrint + 300
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print String(200, "-")
      iPrint = iPrint + 300
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print "筆數：共" & Format(.RecordCount, "#,##0") & "筆!!!"
      Printer.EndDoc
   End With
   
   Exit Sub
      
ErrHnd:

      If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Add by Morgan 2006/10/13 '延緩公告申請與公告日不符資料
Private Sub Get412List()

   Dim stVTB0 As String
   Dim stCon As String
   
On Error GoTo ErrHnd
   
   stCon = " and PA14 BETWEEN " & DBDATE(Text1(5).Text) & " AND " & DBDATE(Text1(6).Text)
   'Modified by Morgan 2013/2/22 GetPrePA14+預估公告天數參數(原固定用30天與目前系統不符,為避免不同步函數改寫傳參數)
   '應公告但未公告資料
   stVTB0 = " select PA01,PA02,PA03,PA04,PA11,b.cp27 cp27b,b.cp71,c.cp27 cp27c,a.cp05,pa14" & _
      ",decode(length(b.cp71),1,GetPrePA14(c.cp27,b.cp71," & 預估公告天數 & "),b.cp71+19110000) pa14x" & _
      " from caseprogress a, patent, caseprogress b, caseprogress c" & _
      " where a.cp01='P' and a.cp10='1906' and a.cp05>to_char(sysdate-180,'yyyymmdd')" & _
      " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04 and pa09='000' and pa14 is null" & _
      " and b.cp09(+)=a.cp43 and b.cp10='412'" & _
      " and c.cp01(+)=a.cp01 and c.cp02(+)=a.cp02 and c.cp03(+)=a.cp03 and c.cp04(+)=a.cp04 and c.cp10='601'"
   
   '不應公告但公告資料
   stVTB0 = stVTB0 & " Union All" & _
      " select PA01,PA02,PA03,PA04,PA11,b.cp27 cp27b,b.cp71,c.cp27 cp27c,a.cp05,pa14" & _
      ",decode(length(b.cp71),1,GetPrePA14(c.cp27,b.cp71," & 預估公告天數 & "),b.cp71+19110000) pa14x" & _
      " from caseprogress a, patent, caseprogress b, caseprogress c" & _
      " where a.cp01='P' and a.cp10='1906' and a.cp05>to_char(sysdate-180,'yyyymmdd')" & _
      " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04 and pa09='000'" & stCon & _
      " and b.cp09(+)=a.cp43 and b.cp10='412'" & _
      " and c.cp01(+)=a.cp01 and c.cp02(+)=a.cp02 and c.cp03(+)=a.cp03 and c.cp04(+)=a.cp04 and c.cp10='601'"
      
   'Modified by Morgan 2025/1/15 +PID
   strSql = "select PA01||'-'||PA02||'-'||PA03||'-'||PA04 C1, PA11 C2, cp27b-19110000 C3, pa14x-19110000 C4, cp27c-19110000 C5, cp71 C6, cp05-19110000 C7, pa14-19110000 C8,'' PID" & _
      " from (" & stVTB0 & ") X where pa14x<>pa14 or (pa14 is null and pa14x BETWEEN " & DBDATE(Text1(5).Text) & " AND " & DBDATE(Text1(6).Text) & ") order by 1"
      
   CheckOC3
   
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   
   'Added by Morgan 2025/1/15
   If AdoRecordSet3.RecordCount > 0 And strSrvDate(1) >= P業務區劃分啟用日 And Combo1 <> "" Then
      Set rsQuery = PUB_CreateRecordset(AdoRecordSet3, , , 300, Me.Name, mSeqNo)
      With rsQuery
         .MoveFirst
         Do While Not .EOF
            .Fields("PID") = PUB_GetPHandler(.Fields("C1"))
            .MoveNext
         Loop
         .UpdateBatch
         
         stVTBX = "select R001 as " & .Fields(0).Name
         For intI = 2 To .Fields.Count
            stVTBX = stVTBX & ", R" & Format(intI, "000") & " as " & .Fields(intI - 1).Name
         Next
         stVTBX = stVTBX & " from Rdatafactory Where Id='" & strUserNum & "' And Formname='" & Me.Name & "'  And Seqno='" & mSeqNo & "'"
      End With
      strSql = "Select X.* From (" & stVTBX & ") X where PID='" & Left(Combo1, 5) & "' ORDER BY 1"
      intI = 1
      Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
   End If
   'end 2025/1/15
   
   
   With AdoRecordSet3
      '若有資料
      If .RecordCount > 0 Then
         If MsgBox("準備列印【延緩公告申請與公告日不符清單】，請更換紙張！", vbExclamation + vbOKCancel) = vbOK Then Print412List AdoRecordSet3
      Else
         MsgBox "無待列印【延緩公告申請與公告日不符清單】！", vbCritical
      End If
   End With
   CheckOC3
   
   Exit Sub

ErrHnd:

      If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Add by Morgan 2004/9/20 '讀取可承辦技術報告清單資料
Private Sub Get421List()

On Error GoTo ErrHnd

   'Modified by Morgan 2014/10/2 收文日改2014/10/2以後(之前已確認無資料)
   'Modified by Morgan 2025/1/15 +PID
   strSql = "SELECT PA01||'-'||PA02||'-'||PA03||'-'||PA04 X01,CP05,PA11,PA14,ST02,'' PID " & _
      " FROM CASEPROGRESS A, PATENT B, STAFF C" & _
      " WHERE CP01='P' AND CP05>20140700 AND CP27 IS NULL AND CP10='421' AND CP57 IS NULL" & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL" & _
      " AND PA09='000' AND PA14 IS NOT NULL" & _
      " AND ST01(+)=CP14" & _
      " ORDER BY 1,2"
      
   CheckOC3
   
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   
   'Added by Morgan 2025/1/15
   If AdoRecordSet3.RecordCount > 0 And strSrvDate(1) >= P業務區劃分啟用日 And Combo1 <> "" Then
      Set rsQuery = PUB_CreateRecordset(AdoRecordSet3, , , 300, Me.Name, mSeqNo)
      With rsQuery
         .MoveFirst
         Do While Not .EOF
            .Fields("PID") = PUB_GetPHandler(.Fields("X01"))
            .MoveNext
         Loop
         .UpdateBatch
         
         stVTBX = "select R001 as " & .Fields(0).Name
         For intI = 2 To .Fields.Count
            stVTBX = stVTBX & ", R" & Format(intI, "000") & " as " & .Fields(intI - 1).Name
         Next
         stVTBX = stVTBX & " from Rdatafactory Where Id='" & strUserNum & "' And Formname='" & Me.Name & "'  And Seqno='" & mSeqNo & "'"
      End With
      strSql = "Select X.* From (" & stVTBX & ") X where PID='" & Left(Combo1, 5) & "' ORDER BY 1,2"
      intI = 1
      Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
   End If
   'end 2025/1/15
   
   With AdoRecordSet3
      '若有資料
      If .RecordCount > 0 Then
         If MsgBox("準備列印【技術報告可發文清單】，請更換紙張！", vbExclamation + vbOKCancel) = vbOK Then Print421List AdoRecordSet3
      Else
         MsgBox "無待列印【技術報告可發文清單】！", vbCritical
      End If
   End With

   CheckOC3
   Exit Sub
   
ErrHnd:
'      Print #FF2, "***ErrHnd"
      If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdok_Click(Index As Integer)
   'edit by nickc 2007/02/06 不用 dll 了
   'Dim objPrintDllPublic As clsPrintPublic
   Dim strTmp As String, rsTemp1 As New ADODB.Recordset, rsTemp2 As New ADODB.Recordset
   Dim strTxt(1 To 2) As String
   Dim strNation As String
   Dim strPrintKind As String '1:依公告日範圍列印, 2:依本所案號+公告日
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim blnPrtContSheet As Boolean '是否列印聯絡單
   'Add by Morgan 2004/7/16
   Dim stPA14 As String '公告日
   Dim stPA08 As String '專利種類
   Dim stET03 As String '定稿處理狀況
   Dim pa(1 To 4) As String   '本所案號
   'Add By Sindy 2011/12/21
   Dim bolPrintPDF As Boolean
   Dim int_Copys As Integer
   '2011/12/21 End
   Dim stPA26 As String, strCP09 As String 'Add By Sindy 2014/6/18
   Dim stPA75 As String 'Added by Morgan 2014/7/22
   Dim strLP26 As String 'Added by Morgan 2016/1/11
   Dim intLetterCount As Integer 'Added by Morgan 2017/6/2
   
On Error GoTo ErrHnd
   
   blnClkSure = False
   Select Case Index
      Case 0 '確定
         ' 設定滑鼠游標為等待狀態
         Screen.MousePointer = vbHourglass
        '若未輸入本所案號才要檢查公告日, 及國家
         If Me.Text1(1).Text = "" Or Me.Text1(2).Text = "" Then
            If IsEmptyText(Text1(5)) = True Then
               MsgBox "請輸入公告起始日!", vbOKOnly + vbCritical, "檢核資料"
               Text1(5).SetFocus
               ' 設定滑鼠游標為預設值
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            If IsEmptyText(Text1(6)) = True Then
               MsgBox "請輸入公告結束日!", vbOKOnly + vbCritical, "檢核資料"
               Text1(6).SetFocus
               ' 設定滑鼠游標為預設值
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            '檢查公告日
            If PUB_CheckKeyInDate(Me.Text1(5)) = -1 Then
               Me.Text1(5).SetFocus
               Text1_GotFocus 5
               ' 設定滑鼠游標為預設值
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.Text1(6)) = -1 Then
               Me.Text1(6).SetFocus
               Text1_GotFocus 6
               ' 設定滑鼠游標為預設值
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            '檢查公告日範圍
            If Me.Text1(5).Text <> "" And Me.Text1(6).Text <> "" Then
               If Val(Me.Text1(5).Text) > Val(Me.Text1(6).Text) Then
                  MsgBox "公告日範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.Text1(5).SetFocus
                  Text1_GotFocus 5
                  ' 設定滑鼠游標為預設值
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            '檢查國家
            If Me.Text1(7).Text = "" Then
                MsgBox "請輸入申請國家!!!", vbExclamation + vbOKOnly
                Me.Text1(7).SetFocus
                Text1_GotFocus 7
                  ' 設定滑鼠游標為預設值
                  Screen.MousePointer = vbDefault
                Exit Sub
            End If
            '檢查國家
            If Me.Text1(0).Text = "" Then
                MsgBox "請輸入申請國家!!!", vbExclamation + vbOKOnly
                Me.Text1(0).SetFocus
                Text1_GotFocus 0
                  ' 設定滑鼠游標為預設值
                  Screen.MousePointer = vbDefault
                Exit Sub
            End If
            If Me.Text1(7).Text <> "" And Me.Text1(0).Text <> "" Then
               If Me.Text1(7).Text > Me.Text1(0).Text Then
                  MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.Text1(7).SetFocus
                  Text1_GotFocus 7
                  ' 設定滑鼠游標為預設值
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            'Add by Morgan 2004/2/12
            If Me.Text1(9).Text <> "" And Me.Text1(10).Text <> "" Then
               If Me.Text1(9).Text > Me.Text1(10).Text Then
                  MsgBox "公告號範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.Text1(9).SetFocus
                  Text1_GotFocus 9
                  ' 設定滑鼠游標為預設值
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            
            'Add By Sindy 2011/12/21
            'Removed by Morgan 2021/6/24 公報改抓卷宗區，不再往pat3讀取避免當機沒開的情形
            'If Text1(7).Text = "000" And Text1(0).Text = "000" Then
            '   If GetFilePath(DBDATE(Text1(5))) = False Then
            '      Me.txtPath2.SetFocus
            '      ' 設定滑鼠游標為預設值
            '      Screen.MousePointer = vbDefault
            '      Exit Sub
            '   End If
            'End If
            'end 2021/6/24
            '2011/12/21 End
        End If
        'Add By Sindy 2011/12/21
'        If FF2 > 0 Then Close #FF2
'        FF2 = FreeFile
'        Open txtPath2 & "\專利公告通知函" & strTPB04 & "卷" & strTPB05 & "期" & "偵測程式執行狀況.txt" For Output As FF2
        List1.Clear
        'Modify By Sindy 2014/9/3
        '系統印表機
        PUB_RestorePrinter cmbPrinter2
        '設定控制台預設印表機
        PUB_SetOsDefaultPrinter cmbPrinter2
        '2014/9/3 END
''        Print #FF2, "***List1.Clear"
'        If cmbPrinter2.ListIndex >= 0 Then
''            Set Printer = Printers(cmbPrinter2.ListIndex)
''            Print #FF2, "***Set Printer = Printers(cmbPrinter2.ListIndex)"
''            Printer.EndDoc
'            '設定控制台預設印表機
''            If Option1(0).Value = True Then
''               '會出現錯誤
''               'Err.NUMBER : -2147417848
''               'Err.Description : Automation 錯誤用戶端中斷了已啟動物件的連線。
''               PUB_SetOsDefaultPrinter Printers(cmbPrinter2.ListIndex).DeviceName
''            ElseIf Option1(1).Value = True Then
'               SetOsDefaultPrinter Printers(cmbPrinter2.ListIndex).DeviceName
''            End If
'        End If
        '2011/12/21 End
        
        ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/29 清除查詢印表記錄檔欄位
         '若有輸入本所案號
         If Me.Text1(1).Text <> "" And Me.Text1(2).Text <> "" Then
            strPrintKind = "2"
         '若未輸入本所案號
         Else
            strPrintKind = "1"
         End If
         
         'Add by Morgan 2006/10/13
         If strPrintKind = "1" And Text1(7) = "000" Then Get412List
         
         'Add by Morgan 2004/9/20
'         Print #FF2, "***If strPrintKind = "; 1; " And Text1(7) = "; 0; " Then Get421List"
         If strPrintKind = "1" And Text1(7) = "000" Then Get421List
         
         '依是否有輸入本所案號做不同選擇
'         Print #FF2, "***Select Case strPrintKind"
         Select Case strPrintKind
         Case "1" '未輸入本所案號
            '檢查專利種類
            If Len("" & Me.Text1(8).Text) > 0 Then
               If CheckPKindExist("" & Me.Text1(8).Text) = False Then
                  Me.Text1(8).SetFocus
                  Text1_GotFocus 8
                  ' 設定滑鼠游標為預設值
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               pub_QL05 = pub_QL05 & ";" & Label1(0) & Text1(8) & Lbl(0) 'Add By Sindy 2010/11/29
            End If
            strTmp = ""
            If Text1(7).Text <> "" And Text1(0).Text <> "" Then
               strTmp = " AND PA09 BETWEEN '" & Text1(7).Text & "' AND '" & Text1(0).Text & "'"
               'Add By Sindy 2011/12/21
               If Text1(7).Text = "000" And Text1(0).Text = "000" Then
                  strTmp = strTmp & " AND PA23='1'"
               End If
               '2011/12/21 End
            ElseIf Text1(7).Text = "" And Text1(0).Text <> "" Then
               strTmp = " AND PA09 <='" & Text1(0).Text & "'"
            ElseIf Text1(7).Text <> "" And Text1(0).Text = "" Then
               strTmp = " AND PA09 >='" & Text1(7).Text & "'"
            End If
            If Text1(7).Text <> "" Or Text1(0).Text <> "" Then
               pub_QL05 = pub_QL05 & ";" & Label1(1) & Text1(7) & "-" & Text1(0) 'Add By Sindy 2010/11/29
            End If
            'Add by Morgan 2004/2/12
            '加公告號起訖
            If Me.Text1(9).Text <> "" Then strTmp = strTmp & " AND PA15>='" & Me.Text1(9).Text & "'"
            If Me.Text1(10).Text <> "" Then strTmp = strTmp & " AND PA15<='" & Me.Text1(10).Text & "'"
            If Me.Text1(9).Text <> "" Or Me.Text1(10).Text <> "" Then
               pub_QL05 = pub_QL05 & ";" & Label1(4) & Text1(9) & "-" & Text1(10) 'Add By Sindy 2010/11/29
            End If
            
            '公告日
            If Me.Text1(5).Text <> "" Or Me.Text1(6).Text <> "" Then
               pub_QL05 = pub_QL05 & ";" & Label1(2) & Text1(5) & "-" & Text1(6) 'Add By Sindy 2010/11/29
            End If
            
            'Modify by Morgan 2004/7/16
            '查詢欄位加公告日,專利種類(PA14,PA08)
            'Modify By Sindy 2014/6/18 +,PA26
            'Modified by Morgan 2025/1/15 +PA22,PID,運算欄位補別名(寫暫存重讀要用)
            strExc(0) = "SELECT PA01||PA02||PA03||PA04 C01,CU12,DECODE(TPB07,Null,Null,'01','1','0') C02," & _
               "PA26,PA85,CU64,FA31,PA09,PA01,PA02,PA03,PA04,PA75,PA14,PA08,PA11,PA22,'' PID FROM PATENT,TPBULLETIN,CUSTOMER,FAGENT WHERE " & _
               "PA14 BETWEEN " & TransDate(Text1(5).Text, 2) & " AND " & TransDate(Text1(6).Text, 2) & strTmp & _
               " AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPB01(+) AND " & _
               "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) " & _
               IIf(Len(Me.Text1(8).Text) > 0, " AND PA08='" & Me.Text1(8).Text & "'", " ")
            '限制系統類別
            strExc(0) = strExc(0) & " And  PA01 ='P' "
            
            '若設定申請國家為台灣時, 以公告號由大到小排序
            If Text1(7).Text = 台灣國家代號 Then
               'Added by Morgan 2025/8/6
               '排除公告公報已有信函進度的案件以避免重複執行
               strExc(0) = strExc(0) & " and not exists(select * from caseprogress,letterprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='1228' and lp01(+)=cp09 and lp01 is not null)"
               '排除輸入證書的案件
               strExc(0) = strExc(0) & " and not exists(select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='1603')"
               'end 2025/8/6
            
               'Modify by Morgan 2004/8/12 排序改專利種類、證書號數
               'strExc(0) = strExc(0) & " ORDER BY PA15 DESC "
               strExc(0) = strExc(0) & " ORDER BY PA08, PA22 DESC"
            End If
            '列印地址條時使用
            m_strSQLA = strExc(0)
            intI = 1
            Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            'Added by Morgan 2025/1/15
            If intI = 1 And strSrvDate(1) >= P業務區劃分啟用日 And Combo1 <> "" Then
               Combo1.Tag = ""
               Set rsQuery = PUB_CreateRecordset(rsTemp1, , , 300, Me.Name, mSeqNo)
               With rsQuery
                  .MoveFirst
                  Do While Not .EOF
                     .Fields("PID") = PUB_GetPHandler(.Fields("PA01") & "-" & .Fields("PA02") & "-" & .Fields("PA03") & "-" & .Fields("PA04"))
                     .MoveNext
                  Loop
                  .UpdateBatch
                  
                  stVTBX = "select R001 as " & .Fields(0).Name
                  For intI = 2 To .Fields.Count
                     stVTBX = stVTBX & ", R" & Format(intI, "000") & " as " & .Fields(intI - 1).Name
                  Next
                  stVTBX = stVTBX & " from Rdatafactory Where Id='" & strUserNum & "' And Formname='" & Me.Name & "'  And Seqno='" & mSeqNo & "'"
               End With
               strSql = "Select X.* From (" & stVTBX & ") X where PID='" & Left(Combo1, 5) & "'"
               
               'Added by Morgan 2025/8/6
               If Text1(7).Text = 台灣國家代號 Then
                  '先檢查是否有本所領證但尚未發證的案件(主語法已有排除有證書的案件),以免過早執行而產生不需要的通知函
                  strExc(0) = strSql & " and exists(select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='601' and cp27>19221111)"
                  intI = 1
                  Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If MsgBox("有本所領證但尚未發證的案件，請確認證書皆已輸入，是否仍要繼續？", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                     End If
                  End If
               End If
               'end 2025/8/6
            
               strSql = strSql & " ORDER BY PA08, PA22 DESC"
               intI = 1
               Set rsTemp1 = ClsLawReadRstMsg(intI, strSql)
               Combo1.Tag = Combo1
            End If
            'end 2025/1/15
            
            If intI = 1 Then
               With rsTemp1
                  InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/29
                  intLetterCount = 0 'Added by Morgan 2017/6/2
                  Do While Not .EOF
                     'Add By Sindy 2011/12/21
                     bolPrintPDF = False
                     '2011/12/21 End
                     
                     'Add by Morgan 2004/7/16
                     stPA14 = "" & .Fields("PA14")
                     stPA08 = "" & .Fields("PA08")
                     pa(1) = "" & .Fields("PA01")
                     pa(2) = "" & .Fields("PA02")
                     pa(3) = "" & .Fields("PA03")
                     pa(4) = "" & .Fields("PA04")
                     'END
                     stPA26 = "" & .Fields("PA26") 'Add By Sindy 2014/6/18
                     stPA75 = "" & .Fields("PA75") 'Added by Morgan 2014/7/22
                     strNation = "000"
                     If IsNull(.Fields("PA09")) = False Then
                        strNation = .Fields("PA09")
                     End If
                     
                     'Add By Sindy 2011/12/21
                     '申請國家為000-000時,以公告日抓基本檔卷宗性質為'申請'者
                     '若該案號進度檔無'1603.專利證書'程序時,則印公告通知函定稿外加該筆PDF
                     If Text1(7).Text = "000" And Text1(0).Text = "000" Then
                        'Removed by Morgan 2025/8/6 改上面語法直接排除
                        'intI = 1
                        'strExc(0) = "SELECT CP09 FROM CASEPROGRESS WHERE " & ChgCaseprogress(.Fields(0)) & " AND CP10='1603' ORDER BY CP05 DESC,CP09 DESC"
                        'Set rsTemp2 = ClsLawReadRstMsg(intI, strExc(0))
                        'If intI = 1 Then
                        '   If rsTemp2.RecordCount > 0 Then GoTo ReadNext1
                        'End If
                        'end 2025/8/6
                        bolPrintPDF = True
                     End If
                     '2011/12/21 End
                     
                     m_bolELetter = False 'Added by Morgan 2014/6/19
                     
                     'Add By Sindy 2014/6/18
                     strCP09 = ""
                     'Modified by Morgan 2019/9/4 若已經跑過則不要再跑(Ex:P121198,重複產生LP導致發文室已發文又再出現)
                     strSql = "SELECT cp09,lp01 FROM caseprogress,letterprogress " & _
                              "WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' " & _
                               " AND CP10 = '1228' and lp01(+)=cp09"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        'Modified by Morgan 2020/5/11
                        'If Not IsNull(RsTemp.Fields("cp09")) Then GoTo ReadNext1 'Added by Morgan 2019/9/4
                        If Not IsNull(RsTemp.Fields("lp01")) Then GoTo ReadNext1 'Added by Morgan 2019/9/4
                        'end 2020/5/11
                        strCP09 = RsTemp.Fields("cp09")
                        'Added by Morgan 2025/2/13 +更新發文人員日期時間,因程序分區管制後全E化客戶函要各自EMail(處理人員=發文人員)
                        cnnConnection.Execute "update caseprogress set cp27=" & strSrvDate(1) & ",cp82=to_char(sysdate,'hh24miss'),cp83='" & strUserNum & "' where cp09='" & strCP09 & "' and cp127 is null", intI
                        'end 2025/2/13
                        'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
                        'Modified by Morgan 2015/12/2 要傳非掛號
                        'Modified by Morgan 2016/1/11 +strLP26
                        Call PUB_AddLetterProgress(RsTemp.Fields("cp09"), 1, True, "", False, stPA26, "1228", stPA75, , strLP26)
                        m_bolELetter = True 'Added by Morgan 2014/6/19
                        
                        'Modified by Morgan 2022/2/14 全E化也不要印
                        'If m_bolELetter And strLP26 = "Y" Then bolPrintPDF = False 'Added by Morgan 2016/1/11 e化不印公報
                        If m_bolELetter And strLP26 <> "" Then bolPrintPDF = False
                        'end 2022/2/14
                     End If
                     '2014/6/18 END
                     
                     If IsNull(.Fields(0)) = False Then
                        If rsA.State <> adStateClosed Then rsA.Close
                        Set rsA = Nothing
                        StrSQLa = "Select CP09 From CaseProgress WHERE " & ChgCaseprogress("" & .Fields(0).Value) & " AND CP09 <'B' And CP05 IS NOT NULL AND CP09 IS NOT NULL ORDER BY CP05 DESC, CP09 DESC "
                        rsA.CursorLocation = adUseClient
                        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                        If rsA.RecordCount > 0 Then
                            '(依申請國家不同)
                            If strNation < "010" Then
                               stET03 = "01"
                               If Val(stPA14) >= 20040701 Then
                                 stET03 = "06"
                                 '新型未收技術報告加註
                                 If stPA08 = "2" Then
                                    If PUB_ChkCPExist(pa, "421") = False Then
                                       stET03 = "07"
                                    End If
                                 End If
                               End If
                               'Modify By Sindy 2014/6/18
                               'NowPrint rsA.Fields(0).Value, ET01, stET03, False, strUserNum, 0
                               NowPrint rsA.Fields(0).Value, ET01, stET03, False, strUserNum, 0, , , , , , , , , , , , strCP09
                               '2014/6/18 END
                            Else
                               'Modify By Sindy 2014/6/18
                               'NowPrint rsA.Fields(0).Value, ET01, "02", False, strUserNum, 0
                               NowPrint rsA.Fields(0).Value, ET01, "02", False, strUserNum, 0, , , , , , , , , , , , strCP09
                               '2014/6/18 END
                            End If
                            intLetterCount = intLetterCount + 1
                        End If
                        If rsA.State <> adStateClosed Then rsA.Close
                        Set rsA = Nothing
                     End If
                     'Add By Sindy 2011/12/21
                     If bolPrintPDF = True Then
                        Call GetPDFCopys(pa(1), pa(2), pa(3), pa(4), "" & .Fields("PA11"), int_Copys)
                     End If
                     '2011/12/21 End
                     
                     intI = 1
                     strExc(0) = "SELECT CP09 FROM CASEPROGRESS WHERE " & ChgCaseprogress(.Fields(0)) & " AND CP09<'C' ORDER BY CP05 DESC,CP09 DESC"
                     Set rsTemp2 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                     If intI = 1 And "" & .Fields(2) = "0" Then
                            If blnPrtContSheet = False Then MsgBox "準備列印聯絡單，請更換紙張!!!", vbExclamation + vbOKOnly: blnPrtContSheet = True
                            m_strSQLB = "select st03 a01,st02 a02," & ChgCaseprogress("", 1) & " a03,pa05 a04,pa06 a05," & _
                                                "pa07 a06,cu04 a07,FA04 a08," & SQLDate("PA14", True) & " a09,TPB08 a10 " & _
                                                "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,CUSTOMER,FAGENT,TPBULLETIN WHERE cp09='" & rsTemp2.Fields(0).Value & "' AND " & _
                                                "CP01=PA01 and CP02=PA02 and CP03=PA03 and CP04=PA04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) and " & _
                                                "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND " & _
                                                "PA11=TPB01(+)"
                            m_rsA.CursorLocation = adUseClient
                            m_rsA.Open m_strSQLB, cnnConnection, adOpenStatic, adLockReadOnly
                            If m_rsA.RecordCount > 0 Then
                                '取得預設印表機設定值
                                m_PrtOrientation = Printer.Orientation
                                m_PrtScaleMode = Printer.ScaleMode
                                '重新設定印表機
                                Printer.PaperSize = vbPRPSA4
                                Printer.Orientation = vbPRORPortrait
                                Printer.ScaleMode = vbCentimeters
                                '列印聯絡單
                                InitPrtPosition 0.5, 0.5
                                PrintContactSheet
                                Printer.EndDoc
                                '還原預設印表機設值
                                Printer.Orientation = m_PrtOrientation
                                Printer.ScaleMode = m_PrtScaleMode
                            End If
                            If m_rsA.State <> adStateClosed Then m_rsA.Close
                            Set m_rsA = Nothing
                     End If
ReadNext1:
                     .MoveNext
                  Loop
               End With
               
'Remove by Morgan 2008/7/18 改開窗定稿紙不必再印地址條
'               If MsgBox("是否列印地址條???", vbYesNo + vbInformation, "列印地址條") = vbYes Then
'                    PrintAddress strPrintKind
'               End If
               
               'Add By Sindy 2011/12/21 列印PDF
               If List1.ListCount > 0 Then
                  Call PrintPDF
                  MsgBox "列印結束 ! (列印PDF花費時間：" & strTime & "  " & time() & ")", vbInformation
                  
               'Added by Morgan 2017/6/2 沒有印定稿不要彈列印結束否則會誤以為有定稿--蕭茹曣
               ElseIf intLetterCount = 0 Then
                  MsgBox "無符合條件之資料可列印 !", vbInformation
               'end 2017/6/2
               Else
               '2011/12/21 End
                  MsgBox "列印結束 !", vbInformation
               End If
            Else
               InsertQueryLog (0) 'Add By Sindy 2010/11/29
               MsgBox "無符合條件之資料可列印 !", vbInformation
            End If
         
         Case "2" '有輸入本所案號
            If IsEmptyText(Text1(2)) = True Then
               MsgBox "請輸入本所案號!", vbOKOnly + vbCritical, "檢核資料"
               Text1(2).SetFocus
               Exit Sub
            End If
            strTmp = Text1(1) & Text1(2)
            pub_QL05 = pub_QL05 & ";" & Label1(3) & Text1(1) & "-" & Text1(2) 'Add By Sindy 2010/11/29
            If Text1(3).Text = "" Then
               strTmp = strTmp & "0"
               pub_QL05 = pub_QL05 & "-0" 'Add By Sindy 2010/11/29
            Else
               strTmp = strTmp & Text1(3).Text
               pub_QL05 = pub_QL05 & "-" & Text1(3) 'Add By Sindy 2010/11/29
            End If
            If Text1(4).Text = "" Then
               strTmp = strTmp & "00"
               pub_QL05 = pub_QL05 & "-00" 'Add By Sindy 2010/11/29
            Else
               strTmp = strTmp & Text1(4).Text
               pub_QL05 = pub_QL05 & "-" & Text1(4) 'Add By Sindy 2010/11/29
            End If
            
            If Me.Text1(5).Text <> "" Or Me.Text1(6).Text <> "" Then
               pub_QL05 = pub_QL05 & ";" & Label1(2) & Text1(5) & "-" & Text1(6) 'Add By Sindy 2010/11/29
            End If
            If Text1(7).Text <> "" Or Text1(0).Text <> "" Then
               pub_QL05 = pub_QL05 & ";" & Label1(1) & Text1(7) & "-" & Text1(0) 'Add By Sindy 2010/11/29
            End If
            
            'Modify by Morgan 2004/7/16
            '查詢欄位加公告日,專利種類(PA14,PA08)
            'Modify By Sindy 2014/6/18 +,PA26
            strExc(0) = "SELECT PA01||PA02||PA03||PA04,CU12,DECODE(TPB07,Null,Null,'01','1','0')," & _
               "NVL(PA26,''),PA85,CU64,FA31,PA01,PA02,PA03,PA04,PA09,PA01,PA02,PA03,PA04,PA75,PA14,PA08,PA23,PA11,PA26 FROM PATENT,TPBULLETIN,CUSTOMER,FAGENT WHERE " & ChgPatent(strTmp) & _
               " AND PA11=TPB01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND " & _
               "SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) " & IIf(Me.Text1(5).Text <> "", " AND PA14=" & DBDATE(Me.Text1(5).Text) & " ", "")
            '限制系統類別
            strExc(0) = strExc(0) & " And  PA01 ='P' "
            '若設定申請國家為台灣時, 以公告號由大到小排序
            If Text1(7).Text = 台灣國家代號 Then
               strExc(0) = strExc(0) & " ORDER BY PA15 DESC "
            End If
            '列印地址條時使用
            m_strSQLA = strExc(0)
            intI = 1
            Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               With rsTemp1
                  'Add By Sindy 2011/12/21
                  bolPrintPDF = False
                  '2011/12/21 End
                  
                  InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/29
                  'Add by Morgan 2004/7/16
                  stPA14 = "" & .Fields("PA14")
                  stPA08 = "" & .Fields("PA08")
                  pa(1) = "" & .Fields("PA01")
                  pa(2) = "" & .Fields("PA02")
                  pa(3) = "" & .Fields("PA03")
                  pa(4) = "" & .Fields("PA04")
                  'END
                  stPA26 = "" & .Fields("PA26") 'Add By Sindy 2014/6/18
                  strNation = "000"
                  If IsNull(rsTemp1.Fields("PA09")) = False Then
                     strNation = rsTemp1.Fields("PA09")
                  End If
                  
                  'Add By Sindy 2011/12/21
                  '申請國家為000-000時,以公告日抓基本檔卷宗性質為'申請'者
                  '若該案號進度檔無'1603.專利證書'程序時,則印公告通知函定稿外加該筆PDF
                  If strNation = "000" And "" & .Fields("PA23") = "1" Then
                     intI = 1
                     strExc(0) = "SELECT CP09 FROM CASEPROGRESS WHERE " & ChgCaseprogress(.Fields(0)) & " AND CP10='1603' ORDER BY CP05 DESC,CP09 DESC"
                     Set rsTemp2 = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        If rsTemp2.RecordCount > 0 Then GoTo ReadNext2
                     End If
                     If "" & .Fields("PA14") > "" Then
                        'Modified by Morgan 2021/6/24 公報PDF改抓卷宗區，不再往pat3讀取避免當機沒開的情形
                        'If GetFilePath(.Fields("PA14")) = False Then
                        '   Me.txtPath2.SetFocus
                        If PUB_GetGazettePDF(pa(1), pa(2), pa(3), pa(4)) = False Then
                           MsgBox .Fields(0) & "案公告公報卷宗區PDF檔讀取失敗！", vbExclamation
                        'end 2021/6/24
                           ' 設定滑鼠游標為預設值
                           Screen.MousePointer = vbDefault
                           Exit Sub
                        End If
                        bolPrintPDF = True
                     End If
                  End If
                  '2011/12/21 End
                  
                  m_bolELetter = False 'Added by Morgan 2014/6/19
                  
                  'Add By Sindy 2014/6/18
                  strCP09 = ""
                  'Modified by Morgan 2019/9/4 若已經跑過則不要再跑(Ex:P121198,重複產生LP導致發文室已發文又再出現)
                  strSql = "SELECT cp09,lp01 FROM caseprogress,letterprogress " & _
                           "WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' " & _
                            " AND CP10 = '1228' and lp01(+)=cp09"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     If Not IsNull(RsTemp.Fields("lp01")) Then GoTo ReadNext2 'Added by Morgan 2019/9/4
                     strCP09 = RsTemp.Fields("cp09")
                     
                     'Added by Morgan 2025/2/13 +更新發文人員日期時間,因程序分區管制後全E化客戶函要各自EMail(處理人員=發文人員)
                     cnnConnection.Execute "update caseprogress set cp27=" & strSrvDate(1) & ",cp82=to_char(sysdate,'hh24miss'),cp83='" & strUserNum & "' where cp09='" & strCP09 & "' and cp127 is null", intI
                     'end 2025/2/13
                        
                     'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
                     'Modified by Morgan 2015/12/2 要傳非掛號
                     'Modified by Morgan 2016/1/11 +strLP26
                     Call PUB_AddLetterProgress(RsTemp.Fields("cp09"), 1, True, "", False, stPA26, "1228", stPA75, , strLP26)
                     m_bolELetter = True 'Added by Morgan 2014/6/19
                     
                     'Modified by Morgan 2022/2/14 全E化也不要印
                     'If m_bolELetter And strLP26 = "Y" Then bolPrintPDF = False 'Added by Morgan 2016/1/11 e化不印公報
                     If m_bolELetter And strLP26 <> "" Then bolPrintPDF = False 'Added by Morgan 2016/1/11 e化不印公報
                     'end 2022/2/14
                  End If
                  '2014/6/18 END
                  
                  If IsNull(.Fields(0)) = False Then
                     If rsA.State <> adStateClosed Then rsA.Close
                     Set rsA = Nothing
                     StrSQLa = "Select CP09 From CaseProgress WHERE " & ChgCaseprogress("" & .Fields(0).Value) & " AND CP09<'B' And CP05 IS NOT NULL AND CP09 IS NOT NULL ORDER BY CP05 DESC, CP09 DESC "
                     rsA.CursorLocation = adUseClient
                     rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     If rsA.RecordCount > 0 Then
                         '(依申請國家不同)
                         If strNation < "010" Then
                            stET03 = "01"
                            If Val(stPA14) >= 20040701 Then
                              stET03 = "06"
                              '新型未收技術報告加註
                              If stPA08 = "2" Then
                                  If PUB_ChkCPExist(pa, "421") = False Then
                                    stET03 = "07"
                                  End If
                               End If
                            End If
                            'Modify By Sindy 2014/6/18
                            'NowPrint rsA.Fields(0).Value, ET01, stET03, False, strUserNum, 0
                            NowPrint rsA.Fields(0).Value, ET01, stET03, False, strUserNum, 0, , , , , , , , , , , , strCP09
                            '2014/6/18 END
                         Else
                            'Modify By Sindy 2014/6/18
                            'NowPrint rsA.Fields(0).Value, ET01, "02", False, strUserNum, 0
                            NowPrint rsA.Fields(0).Value, ET01, "02", False, strUserNum, 0, , , , , , , , , , , , strCP09
                            '2014/6/18 END
                         End If
                     End If
                     If rsA.State <> adStateClosed Then rsA.Close
                     Set rsA = Nothing
                  End If
                  'Add By Sindy 2011/12/21
                  If bolPrintPDF = True Then
                     Call GetPDFCopys(pa(1), pa(2), pa(3), pa(4), "" & .Fields("PA11"), int_Copys)
                  End If
                  '2011/12/21 End
                  
                  intI = 1
                  strExc(0) = "SELECT CP09 FROM CASEPROGRESS " & _
                              "WHERE CP01 = '" & .Fields("PA01") & "' AND " & _
                                    "CP02 = '" & .Fields("PA02") & "' AND " & _
                                    "CP03 = '" & .Fields("PA03") & "' AND " & _
                                    "CP04 = '" & .Fields("PA04") & "' AND " & _
                                    " CP09<'C' " & _
                              "ORDER BY CP05 Desc,CP09 DESC"
                  Set rsTemp2 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 And "" & .Fields(2) = "0" Then
                    If blnPrtContSheet = False Then MsgBox "準備列印聯絡單，請更換紙張!!!", vbExclamation + vbOKOnly: blnPrtContSheet = True
                        m_strSQLB = "select st03 a01,st02 a02," & ChgCaseprogress("", 1) & " a03,pa05 a04,pa06 a05," & _
                                            "pa07 a06,cu04 a07,FA04 a08," & SQLDate("PA14", True) & " a09,TPB08 a10 " & _
                                            "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,CUSTOMER,FAGENT,TPBULLETIN WHERE cp09='" & rsTemp2.Fields(0).Value & "' AND " & _
                                            "CP01=PA01 and CP02=PA02 and CP03=PA03 and CP04=PA04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) and " & _
                                            "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND " & _
                                            "PA11=TPB01(+)"
                        m_rsA.CursorLocation = adUseClient
                        m_rsA.Open m_strSQLB, cnnConnection, adOpenStatic, adLockReadOnly
                        If m_rsA.RecordCount > 0 Then
                            '取得預設印表機設定值
                            m_PrtOrientation = Printer.Orientation
                            m_PrtScaleMode = Printer.ScaleMode
                            '重新設定印表機
                            Printer.PaperSize = vbPRPSA4
                            Printer.Orientation = vbPRORPortrait
                            Printer.ScaleMode = vbCentimeters
                            '列印聯絡單
                            InitPrtPosition 0.5, 0.5
                            PrintContactSheet
                            Printer.EndDoc
                            '還原預設印表機設值
                            Printer.Orientation = m_PrtOrientation
                            Printer.ScaleMode = m_PrtScaleMode
                        End If
                        If m_rsA.State <> adStateClosed Then m_rsA.Close
                        Set m_rsA = Nothing
                  End If
               End With
               
'Remove by Morgan 2008/7/18 改開窗定稿紙不必再印地址條
'               If MsgBox("是否列印地址條???", vbYesNo + vbInformation, "列印地址條") = vbYes Then
'                   PrintAddress strPrintKind
'               End If
               
               'Add By Sindy 2011/12/21 列印PDF
               If List1.ListCount > 0 Then
                  Call PrintPDF
                  MsgBox "列印結束 ! (列印PDF花費時間：" & strTime & "  " & time() & ")", vbInformation
               Else
               '2011/12/21 End
                  MsgBox "列印結束 !", vbInformation
               End If
            Else
ReadNext2:
               InsertQueryLog (0) 'Add By Sindy 2010/11/29
               MsgBox "無符合條件之資料可列印 !", vbInformation
            End If
         End Select
         
         ' 設定滑鼠游標為預設
         Screen.MousePointer = vbDefault
         
         'Modify By Sindy 2014/9/3
         '還原系統中預設印表機
         PUB_RestorePrinter m_DefaultPrinter
         '還原控制台預設印表機
         PUB_SetOsDefaultPrinter strPrinter
         '2014/9/3 END
      Case 1 '結束
         Unload Me
   End Select
   
'   Close FF2
   Exit Sub
   
ErrHnd:
   Screen.MousePointer = vbDefault ' 設定滑鼠游標為預設
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   
   'Modify By Sindy 2014/9/3
   '還原系統中預設印表機
   PUB_RestorePrinter m_DefaultPrinter
   '還原控制台預設印表機
   PUB_SetOsDefaultPrinter strPrinter
   '2014/9/3 END
End Sub

'Add By Sindy 2011/12/22
Private Sub GetPDFCopys(strPA01 As String, strPA02 As String, strPA03 As String, strPA04 As String, StrPA11 As String, ByRef int_Copys As Integer)
Dim strFileName As String
   
   int_Copys = 0
   
   'Added by Morgan 2014/6/19
   '103/7/1 起有存電子信函的只要印 1 份
   If Val(strSrvDate(1)) >= 20140701 And m_bolELetter = True Then
      int_Copys = 1
   Else
   'end 2014/6/19
   
      '由員工檔取得列印份數 (北部的員工印2份, 其它地區的員工印3份)
      'Modified by Morgan 2014/5/27 +特殊設定A7所有編號視為北所人員
      'strExc(0) = "SELECT ST06 FROM STAFF WHERE ST01='" & PUB_GetAKindSalesNo(strPA01, strPA02, strPA03, strPA04) & "' "
      strExc(0) = "SELECT DECODE(instr(';'||replace(oMan,',',';')||';',';'||ST01||';'),0,ST06,'1') FROM STAFF, SetSpecMan WHERE ST01='" & PUB_GetAKindSalesNo(strPA01, strPA02, strPA03, strPA04) & "' and ocode(+)='A7'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      int_Copys = 3
      If intI = 1 Then
         If RsTemp.Fields(0).Value = "1" Then
            int_Copys = 2
         Else
            int_Copys = 3
         End If
      End If
      
   End If 'Added by Morgan 2014/6/19
   
   'Modify By Sindy 2013/1/4
   'strFileName = txtPath2 & "\img_1\isu0" & strTPB04 & "0" & strTPB05 & "\" & StrPA11 & "-P01.pdf"
   'Modified by Morgan 2021/6/24 公報改抓卷宗區，不再往pat3讀取避免當機沒開的情形
   'strFileName = txtPath2 & "\img_1\isu0" & strTPB04 & "0" & strTPB05 & "\" & StrPA11 & ".pdf"
   If PUB_GetGazettePDF(strPA01, strPA02, strPA03, strPA04, True, m_AttachPath, strFileName) = False Then
      strFileName = ""
   End If
   'end 2021/6/24
   '2013/1/4 End
   
   If strFileName <> "" Then
      List1.AddItem strFileName & " " & int_Copys
   End If
End Sub

'Add By Sindy 2011/12/22
'Removed by Morgan 2021/6/24 公報改抓卷宗區(原程式已刪除)
'Private Function GetFilePath(strDate As String) As Boolean
'End Sub

'Add By Sindy 2011/12/22
Private Sub PrintPDF()
Dim i As Integer, k As Integer
'Modified by Morgan 2021/6/25
'Dim strTemp As Variant
Dim strTemp(1) As String
'end 2021/6/25
Dim RetVal, intFileCnt As Integer
Dim ff1 As Integer
Dim MySize, dblSec As Double, dblCntSec As Double
'Add By Sindy 2014/9/3
Dim process_id As Long
Dim process_handle_PDF As Long
'2014/9/3 END
   
   strTime = time()
   intFileCnt = 0
   
   'Add By Sindy 2014/9/3
   '因為第 2 個以後開啟的 Reader 才會印完後自動關閉,所以固定先開一個空的程式,全部印完後再關閉
   process_id = Shell(txtPDFPath, vbHide)
   process_handle_PDF = OpenProcess(PROCESS_TERMINATE, 0, process_id)
   '2014/9/3 END
   
'   '設定控制台預設印表機
'   If cmbPrinter2.ListIndex >= 0 Then
'      PUB_SetOsDefaultPrinter Printers(cmbPrinter2.ListIndex).DeviceName
'   End If
   
   If ff1 > 0 Then Close #ff1
   ff1 = FreeFile
   'Modified by Morgan 2021/6/24
   'Open txtPath2 & "\專利公告通知函" & strTPB04 & "卷" & strTPB05 & "期" & "列印PDF時間資訊.txt" For Output As ff1
   Open m_AttachPath & "\專利公告通知函" & strTPB04 & "卷" & strTPB05 & "期" & "列印PDF時間資訊.txt" For Output As ff1
   'end 2021/6/24
   
   For i = 0 To List1.ListCount - 1
      'Modified by Morgan 2021/6/25
      'strTemp = Split(List1.List(i), " ")
      intI = InStrRev(List1.List(i), " ")
      strTemp(0) = Left(List1.List(i), intI - 1)
      strTemp(1) = Mid(List1.List(i), intI + 1)
      'end 2021/6/25
      For k = 0 To Val(strTemp(1)) - 1 '列印份數
         intFileCnt = intFileCnt + 1
   '      AcroPDF1.src = List1.List(i)
   '      AcroPDF1.LoadFile (List1.List(i))
   '      DoEvents
   '      Sleep 5000
   '      AcroPDF1.printAll
   '      DoEvents
   '      Sleep 5000
         'C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe
         'RetVal = SHELL(txtPDFPath & " /p /h /t " & List1.List(i), vbHide) '/p /h /t
         
         'Modified by Morgan 2017/5/15
         'RetVal = SHELL(txtPDFPath & " /p /h " & strTemp(0), vbHide) '/p /h /t
         'MySize = FileLen(strTemp(0))   '傳回檔案長度 (以 Byte 為單位)
         PUB_PrintOnePdf txtPDFPath, " /n /t """ & strTemp(0) & """ """ & cmbPrinter2 & """"
         'end 2017/5/15
         
         'DoEvents
         'If i = 0 And k = 0 Then
'         If CDbl(MySize) <= CDbl(512000) Then
'            Sleep 5000
'         ElseIf CDbl(MySize) > CDbl(512000) And CDbl(MySize) < CDbl(1048576) Then
'            Sleep 8000
'         Else
'            Sleep 10000
'         End If

'Modified by Morgan 2017/5/15
'         '依檔案大小決定秒數
'         If Val(txtByte) = 0 Then
'            MsgBox "[幾Byte算1秒]此欄位不可空白 !", vbInformation
'            txtByte.SetFocus
'            Close #ff1
'            Exit Sub
'         End If
'         dblCntSec = Round(MySize / CDbl(txtByte), 0)
'         If dblCntSec <= 0 Then
'            dblSec = (CDbl(txtMinSec) * 1000)
'         Else
'            dblSec = (dblCntSec * 1000)
'         End If
'         '開第1個檔案時多加幾秒
'         If Val(txtFirstAdd) > 0 Then
'            If i = 0 And k = 0 Then dblSec = dblSec + (CDbl(txtFirstAdd) * 1000)
'         End If
'         '至少幾秒
'         If Val(txtMinSec) > 0 Then
'            If dblSec < (CDbl(txtMinSec) * 1000) Then
'               dblSec = (CDbl(txtMinSec) * 1000)
'            End If
'         End If
'         '最多幾秒
'         If Val(txtMaxSec) > 0 Then
'            If dblSec > (CDbl(txtMaxSec) * 1000) Then
'               dblSec = (CDbl(txtMaxSec) * 1000)
'            End If
'         End If
'         Sleep dblSec
'
'         If k = 0 Then
'            Print #ff1, Left(i + 1 & "     ", 5) & List1.List(i) & " " & MySize & " " & dblCntSec & " " & dblSec
'         End If
         If k = 0 Then
            MySize = FileLen(strTemp(0))   '傳回檔案長度 (以 Byte 為單位)
            Print #ff1, Left(i + 1 & "     ", 5) & List1.List(i) & " " & MySize
         End If
'end 2017/5/15
      Next k
   Next i
   
'   '還原控制台預設印表機
'   If cmbPrinter2.ListIndex >= 0 Then
'      PUB_SetOsDefaultPrinter m_DefaultPrinter
'   End If
   
   Print #ff1, "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   Print #ff1, "列印時間：" & strTime & "  " & time()
   Print #ff1, "檔案數量：" & intFileCnt
   Close ff1
   
   'Add By Sindy 2014/9/3
   TerminateProcess process_handle_PDF, 0&
   CloseHandle process_handle_PDF
   DoEvents
   '2014/9/3 END
End Sub

Private Sub Form_Load()
'Dim SeekPrintL As Integer
'Dim i As Integer, j As Integer
   
   MoveFormToCenter Me
   intWhere = 國內
   
   'Add By Sindy 2011/12/22
   'Modify By Sindy 2014/9/3
   PUB_SetPrinter Me.Name, cmbPrinter2, m_DefaultPrinter
   strPrinter = PUB_GetOsDefaultPrinter '抓控制台目前預設的印表機
   '2014/9/3 END
   
   List1.Clear
   If Pub_StrUserSt03 = "M51" Then
      'txtPDFPath = "C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe"
      Me.Height = 6120
   Else
      'C:\Program Files\Adobe\Acrobat 8.0\Acrobat\Acrobat.exe
      'C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe
      'txtPDFPath = "C:\Program Files\Adobe\Acrobat 8.0\Acrobat\Acrobat.exe"
      Me.Height = 3945
   End If
   '2011/12/22 End
   txtPDFPath = PUB_SetFileAssociation 'Add By Sindy 2014/9/3
   
   'Added by Morgan 2025/1/15
   If strSrvDate(1) >= P業務區劃分啟用日 Then
      Combo1.Visible = True
      Label4.Visible = True
      Call SetPatentP12Combo(Combo1, "P", Label4)
   End If
   'end 2025/1/15

   'Added by Morgan 2021/6/24 公報PDF暫存路徑
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   'end 2021/6/24
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
'   'Add By Sindy 2011/12/22
'   If m_DefaultPrinter <> m_DefaultPrinter2 Then
'      PUB_SetOsDefaultPrinter m_DefaultPrinter2
'   End If
   '2011/12/22 End
   'Modify By Sindy 2014/9/3
   '若印表機變動, 則更新列印設定
   If Me.cmbPrinter2.Text <> Me.cmbPrinter2.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter2.Name, "0", "0", Me.cmbPrinter2.Text
   End If
   '2014/9/3 END
   
   Set rsQuery = Nothing 'Added by Morgan 2025/1/15
   
   Set frm040301 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'910624 add by Sieg
Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
      Case 6 '公告日(迄)
         'Modify By Cheng 2002/09/10
         If blnClkSure = False Then
            If Text1(5).Text <> "" Then
               If Not ChkRange(Text1(5), Text1(6), "公告日") Then
                  Text1(5).SetFocus
                  TextInverse Text1(5)
               End If
            End If
         Else
            blnClkSure = False
         End If
      Case 0 '申請國家(迄)
         'Modify By Cheng 2002/09/10
         If blnClkSure = False Then
            If Text1(7).Text <> "" Then
               If Text1(7) > Text1(0) Then
                  MsgBox "申請國家範圍不正確，請重新輸入 !", vbCritical
                  Text1(7).SetFocus
                  TextInverse Text1(7)
               End If
            End If
         Else
            blnClkSure = False
         End If
     'Add by Morgan 2004/2/12
     '公告號起訖
     Case 10
        If blnClkSure = False Then
            If Text1(9).Text <> "" Then
               If Not ChkRange(Text1(9), Text1(10), "公告號") Then
                  Text1(9).SetFocus
                  TextInverse Text1(9)
               End If
            End If
         Else
            blnClkSure = False
         End If
         
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 1 '系統類別
         If IsEmptyText(Me.Text1(Index).Text) = False Then
            If Text1(Index).Text <> "P" Then
               MsgBox "系統別必需為 P，請重新輸入 !", vbCritical
               Cancel = True
            End If
         End If
      Case 5, 6 '公告日起迄
         If IsEmptyText(Text1(Index)) = False Then
            Cancel = Not ChkDate(Text1(Index))
         End If
      Case 8 '專利種類
         If Len("" & Me.Text1(8).Text) > 0 Then
            If CheckPKindExist("" & Me.Text1(Index).Text) = False Then
               Cancel = True
            End If
         End If
   End Select
   If Cancel Then TextInverse Text1(Index)
End Sub

'Add By Cheng 2002/05/14
Private Function CheckPKindExist(strText As String) As Boolean
Dim Rs As New ADODB.Recordset
CheckPKindExist = True
If Rs.State <> adStateClosed Then Rs.Close
Set Rs = Nothing
Rs.CursorLocation = adUseClient
Rs.Open "Select * From PATENTTRADEMARKMAP WHERE PTM01='1' AND PTM02='" & strText & "'", cnnConnection, adOpenStatic, adLockReadOnly
If Rs.RecordCount <= 0 Then
   Me.Lbl(0).Caption = ""
   MsgBox "無此專利種類!!!", vbExclamation + vbOKOnly
   CheckPKindExist = False
Else
   If Me.Text1(7).Text = "020" And Me.Text1(0).Text = "020" Then
      Me.Lbl(0).Caption = "" & Rs("PTM04").Value
   Else
      Me.Lbl(0).Caption = "" & Rs("PTM03").Value
   End If
End If
If Rs.State <> adStateClosed Then Rs.Close
Set Rs = Nothing
End Function

'Add By Cheng 2003/04/03
Private Sub InitPrtPosition(dblTop As Double, dblLeft As Double)
    m_dblTop = dblTop
    m_dblLeft = dblLeft
    m_dblTitleHeight = 0
    m_dblLine = 0
    m_dblLineHeight = 1
    m_dblBetweenLine = 0.2
    m_dblLineHeight1 = 0.6
    m_dblBetweenLine1 = 0.1
End Sub

'Add By Cheng 2003/04/03
Private Sub PrintContactSheet()
Dim dblPrtX As Double
Dim dblPrtY As Double
Dim ii As Integer
Dim jj As Integer
Dim strTxt  As String
Dim intTxtLeng As Integer
    
    Printer.Font.Name = "標楷體"
    Printer.Font.Size = 16
    'Removed by Morgan 2020/3/30 取消
    'dblPrtX = m_dblLeft + (19 - Printer.TextWidth("台一國際專利商標事務所")) / 2
    'dblPrtY = m_dblTop + m_dblBetweenLine + 0
    'Printer.CurrentX = dblPrtX
    'Printer.CurrentY = dblPrtY
    'Printer.Print "台一國際專利商標事務所"
    'end 2020/3/30
    dblPrtX = m_dblLeft + (19 - Printer.TextWidth("簡易聯絡單")) / 2
    dblPrtY = m_dblTop + m_dblBetweenLine + 1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "簡易聯絡單"
        
    m_dblTitleHeight = 2.2
    
    m_dblLine = 0
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 19, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight + 9 * m_dblLineHeight)
    Printer.Line (m_dblLeft + 4.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 4.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight + 3 * m_dblLineHeight)
    Printer.Line (m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight + 9 * m_dblLineHeight)
    Printer.Line (m_dblLeft + 19, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 19, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight + 9 * m_dblLineHeight)
    Printer.Font.Size = 14
    dblPrtX = m_dblLeft + (4.5 - Printer.TextWidth("受文者")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "受文者"
    dblPrtX = m_dblLeft + 4.5 + (4 - Printer.TextWidth("發文者")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "發文者"
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    '受文者部門
    dblPrtX = m_dblLeft + m_dblBetweenLine
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "" & m_rsA.Fields("A01").Value
    '受文者
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + (4.5 - Printer.TextWidth("" & m_rsA.Fields(1).Value)) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight - 0.3 * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "" & m_rsA.Fields("A02").Value
    '發文者
    dblPrtX = m_dblLeft + 4.5 + (4 - Printer.TextWidth(GetStaffName(strUserNum, True))) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight - (m_dblLineHeight / 2)
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print GetStaffName(strUserNum, True)
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    Printer.Line (m_dblLeft + 2.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 2.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight + 6 * m_dblLineHeight)
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("發文時間")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "發文時間"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth(Mid(ServerDate, 1, 4) - 1911 & "年" & Mid(ServerDate, 5, 2) & "月" & Mid(ServerDate, 7, 2) & "日  " & Format(Left(Right("000000" & ServerTime, 6), 4), "##:##"))) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print Mid(ServerDate, 1, 4) - 1911 & "年" & Mid(ServerDate, 5, 2) & "月" & Mid(ServerDate, 7, 2) & "日  " & Format(Left(Right("000000" & ServerTime, 6), 4), "##:##")
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("答覆")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "答覆"
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("□否 □要")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "□否 □要"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth("用□電話 □口頭  回覆")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight - (m_dblLineHeight / 2)
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "用□電話 □口頭  回覆"
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("　限　　")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "　限　　"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth("            AM")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "            AM"
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("時　要求")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight + 0.5 * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "時　要求"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth("    月    日    ")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight + 0.5 * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "    月    日    "
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("　間　　")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "　間　　"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth("            PM")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "            PM"
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("發文地點")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "發文地點"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth("北所")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "北所"
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 19, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)

    Printer.Font.Size = 13
    m_dblLine = 0
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "本所案號：" & m_rsA.Fields("A03").Value
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "案件名稱(中)：" & m_rsA.Fields("A04").Value
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    'Modify By Sindy 2009/09/04
    'Printer.Print "案件名稱(英)：" & m_rsA.Fields("A05").Value
    If Len(Trim(m_rsA.Fields("A05").Value)) > 26 Then
      Printer.Print "案件名稱(英)：" & Left(Trim(m_rsA.Fields("A05").Value), 26) & "..."
    Else
      Printer.Print "案件名稱(英)：" & Trim(m_rsA.Fields("A05").Value)
    End If
    '2009/09/04 End
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "案件名稱(日)：" & m_rsA.Fields("A06").Value
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "申請人：" & m_rsA.Fields("A07").Value
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "代理人：" & m_rsA.Fields("A08").Value
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "公告日期：" & m_rsA.Fields("A09").Value
    m_dblLine = m_dblLine + 1
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "此案件原為本所代理，現已公告，但已變更代理人"
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "為：" & m_rsA.Fields("A10").Value
End Sub

'Add by Morgan 2006/10/13
'限定字串長度
'Remove by Lydia 2018/08/24 與basQuery重複
'Private Function convForm(ByVal p_InStr As String, ByVal p_Num As Integer, Optional ByVal p_Char As String = " ") As String
'   convForm = StrConv(LeftB(StrConv(p_InStr & String(p_Num, p_Char), vbFromUnicode), p_Num), vbUnicode)
'End Function
