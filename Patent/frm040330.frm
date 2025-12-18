VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm040330 
   BorderStyle     =   1  '單線固定
   Caption         =   "結餘單列印"
   ClientHeight    =   3520
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   7620
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3520
   ScaleWidth      =   7620
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1512
      TabIndex        =   2
      Top             =   1104
      Width           =   500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   2
      Left            =   225
      TabIndex        =   1
      Top             =   1152
      Width           =   1250
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   2250
      MaxLength       =   15
      TabIndex        =   8
      Top             =   2290
      Width           =   1620
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   5700
      TabIndex        =   16
      Text            =   "ALL"
      Top             =   3795
      Width           =   1290
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束"
      Height          =   390
      Index           =   1
      Left            =   3465
      TabIndex        =   10
      Top             =   165
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   390
      Index           =   0
      Left            =   2550
      TabIndex        =   9
      Top             =   165
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "重印"
      Height          =   180
      Index           =   1
      Left            =   225
      TabIndex        =   6
      Top             =   1810
      Width           =   1250
   End
   Begin VB.OptionButton Option1 
      Caption         =   "整批列印"
      Height          =   180
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   780
      Value           =   -1  'True
      Width           =   1250
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   390
      MaxLength       =   15
      TabIndex        =   7
      Top             =   2290
      Width           =   1620
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1104
      Width           =   810
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   3072
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "0"
      Top             =   1104
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   4
      Left            =   3432
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "00"
      Top             =   1104
      Width           =   390
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   90
      TabIndex        =   12
      Top             =   2780
      Width           =   4000
      Begin VB.ComboBox Combo1 
         Height          =   260
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   11
         Top             =   180
         Width           =   3100
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   2
         Left            =   105
         TabIndex        =   13
         Top             =   255
         Width           =   765
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
      Height          =   810
      Left            =   2745
      TabIndex        =   17
      Top             =   5445
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1446
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "PS：法律所在1公司收款案件請自行輸"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   4
      Left            =   4410
      TabIndex        =   33
      Top             =   3030
      Width           =   2970
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "結餘傳票，並將傳票號碼給電腦中心。"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   3
      Left            =   4410
      TabIndex        =   32
      Top             =   3240
      Width           =   3060
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "尚無法結算明細表固定印成PDF"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   1560
      TabIndex        =   31
      Top             =   780
      Width           =   2655
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TS  ==>     26；"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   9
      Left            =   4425
      TabIndex        =   30
      Top             =   1980
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TR  ==>     64； TT  ==>     7；"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   8
      Left            =   4425
      TabIndex        =   29
      Top             =   1785
      Width           =   2880
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TD  ==>    130； TM  ==>    25；"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   7
      Left            =   4425
      TabIndex        =   28
      Top             =   1575
      Width           =   2880
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TC  ==>  10025； TB  ==>   115；"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   4425
      TabIndex        =   27
      Top             =   1380
      Width           =   2880
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "T   ==> 126722； TF  ==>   450；"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   4425
      TabIndex        =   26
      Top             =   1185
      Width           =   2880
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CFC ==>    683； CFL ==> 10408；"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   4425
      TabIndex        =   25
      Top             =   975
      Width           =   2880
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CFT ==>   8295； S   ==>     1；"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   4425
      TabIndex        =   24
      Top             =   780
      Width           =   2880
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "P   ==>  63583； PS  ==>    14；"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   4425
      TabIndex        =   23
      Top             =   585
      Width           =   2880
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CFP ==>  12986； CPS ==>     1；"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   4425
      TabIndex        =   22
      Top             =   375
      Width           =   2880
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "      3.專利其他可結餘日早於6個月"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   4425
      TabIndex        =   21
      Top             =   2745
      Width           =   2970
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "      2.專利 EPC案可結餘日早於1年"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   4425
      TabIndex        =   20
      Top             =   2535
      Width           =   2970
   End
   Begin VB.Label Label5 
      Caption         =   "整批列印需花費一段時間請耐心等候！"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   225
      TabIndex        =   19
      Top             =   180
      Width           =   2030
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "條件：1.商標可結餘日早於3個月"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   4425
      TabIndex        =   18
      Top             =   2310
      Width           =   2610
   End
   Begin VB.Line Line1 
      X1              =   1740
      X2              =   2580
      Y1              =   2410
      Y2              =   2410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "新系統案號："
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   4425
      TabIndex        =   15
      Top             =   180
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "結餘單號："
      Height          =   180
      Index           =   1
      Left            =   410
      TabIndex        =   14
      Top             =   2060
      Width           =   1010
   End
   Begin VB.Line Line2 
      Index           =   0
      Visible         =   0   'False
      X1              =   1968
      X2              =   3800
      Y1              =   1224
      Y2              =   1224
   End
End
Attribute VB_Name = "frm040330"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2011/3/10 整理 by sonia
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit
Dim SeekTemp1 As String, SeekTemp2 As String, SeekPrint As Integer, SeekPrintL As Integer, SeekTempPrint As String
Dim strSql As String, i As Integer, j As Integer, strTemp1 As Variant, strTemp2 As Variant, s As Integer
Dim SeekStr As String, Page As Integer
Dim Seeki1 As Integer
Dim strTemp(0 To 20) As String, DayTemp As String
Dim MaxDay As String, IsPrintok As Boolean
Dim m_CP13St01 As String
Dim m_CP13St02 As String
Dim rs940629 As New ADODB.Recordset
'add by nickc 2005/09/21
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_Title2 As String
Dim PLeft(0 To 11) As Integer
Dim iPrint As Integer
Dim m_Remark As String  '2009/11/24 add by sonia
Dim IsOldSystem As Boolean
Dim strSQLCP As String  '2011/6/1 ADD BY SONIA
Dim strSQLnp As String  '2011/6/1 ADD BY SONIA
Dim strPrinter As String 'add 2016/5/19 改印表機選項的寫法
Dim NA69Emp As String    'add by sonia 2023/7/27
Dim NPsql As String      'add by sonia 2025/5/23
Dim tmpY As Integer      'add by sonia 2025/6/27  從各段抽出來

Private Sub cmdok_Click(Index As Integer)
Dim tRS As New ADODB.Recordset
Dim StrSQLa As String
   
   Select Case Index
      Case 0
         If Option1(0).Value = True Then
         ElseIf Option1(1).Value = True Then
            If Trim(txt1(0)) = "" Or Trim(txt1(5)) = "" Then
               MsgBox "結餘單號不可空白！", vbCritical, "輸入錯誤！"
               txt1(0).SetFocus
               Exit Sub
            End If
         'add by sonia 2020/2/14
         ElseIf Option1(2).Value = True Then
            If Trim(txt1(1)) = "" Then
               MsgBox "本所案號不可空白！", vbCritical, "輸入錯誤！"
               txt1(1).SetFocus
               Exit Sub
            Else
               Set tRS = New ADODB.Recordset
               tRS.CursorLocation = adUseClient
               tRS.Open "select * from caseprogress where cp01='" & txt1(1) & "' and cp02='" & txt1(2) & "' and cp03='" & IIf(Trim(txt1(3)) = "", "0", txt1(3)) & "' and cp04='" & IIf(Trim(txt1(4)) = "", "00", txt1(4)) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
               If tRS.RecordCount = 0 Then
                  MsgBox "無此本所案號！", vbCritical, "輸入錯誤！"
                  txt1(2).SetFocus
                  Exit Sub
               End If
            End If
         'end 2020/2/14
         End If
'         If Combo1.ListIndex >= SeekPrint Then
'            j = Combo1.ListIndex + 1
'         Else
'            j = Combo1.ListIndex
'         End If
'         'A4
'         'modify by sonia 2016/5/19 第一筆之Printer.EndDoc會錯誤但找不出原因故先改寫
'         '設定使用者所選擇的印表機成預設印表機
'         For Each m_Prn In Printers
'            If m_Prn.DeviceName = Combo1.Text Then
'               Set Printer = m_Prn
'               Exit For
'            End If
'         Next
'         'end 2016/5/19
'
         Screen.MousePointer = vbHourglass
         PUB_RestorePrinter Combo1    'add by sonia 2016/11/3
         IsPrintok = False
         If Option1(0).Value = True Then
               PrintByAll
         'add by sonia 2020/2/14 +加單筆本所案號條件
         ElseIf Option1(2).Value = True Then
               PrintByAll
               txt1(2).SetFocus
         'end 2020/2/14
         Else
               'edit by nickc 2005/12/06
               'PrintByOneOld txt1(1), txt1(2), IIf(Trim(txt1(3)) = "", "0", txt1(3)), IIf(Trim(txt1(4)) = "", "00", txt1(4))
               'edit by nickc 2006/06/02 改成可以印多筆
               'PrintByOneOld txt1(1)
               Dim tmpRss As New ADODB.Recordset
               Set tmpRss = New ADODB.Recordset
               strSql = "select A240002 from ACC240 where A240002>='" & txt1(0) & "' and A240002<='" & txt1(5) & "' and A240003 is null order by A240002"
               If tmpRss.State = 1 Then tmpRss.Close
               tmpRss.CursorLocation = adUseClient
               tmpRss.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If tmpRss.RecordCount <> 0 Then
                    tmpRss.MoveFirst
                    Do While Not tmpRss.EOF
                        PrintByOneOld CheckStr(tmpRss.Fields("A240002"))
                        tmpRss.MoveNext
                    Loop
               End If
               txt1(0) = ""
               txt1(5) = ""
'               txt1(2) = ""
'               txt1(3) = ""
'               txt1(4) = ""
               txt1(0).SetFocus
         End If
         If IsPrintok = True Then
            ShowPrintOk
         End If
         PUB_RestorePrinter strPrinter   'add by sonia 2016/10/3 先還原
         Screen.MousePointer = vbDefault
         
      Case 1
         Unload Me
      Case Else
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   PUB_SetPrinter Me.Name, Combo1, strPrinter
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   Set frm040330 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
         txt1(0).Enabled = False
         txt1(5).Enabled = False
         txt1(1).Enabled = False
         txt1(2).Enabled = False
         txt1(3).Enabled = False
         txt1(4).Enabled = False
      Case 1
         txt1(0).Enabled = True
         txt1(5).Enabled = True
         txt1(1).Enabled = False
         txt1(2).Enabled = False
         txt1(3).Enabled = False
         txt1(4).Enabled = False
         txt1(0).SetFocus
      Case 2
         txt1(1).Enabled = True
         txt1(2).Enabled = True
         txt1(3).Enabled = True
         txt1(4).Enabled = True
         txt1(0).Enabled = False
         txt1(5).Enabled = False
         txt1(1).SetFocus
      Case Else
   End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1, 3, 5
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   If Index = 1 Then txt1(5) = txt1(0)
End Sub

'cancel by sonia 2020/2/14
'Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
'   If Trim(txt1(Index)) = "" Then Exit Sub
'   Select Case Index
'      Case 0
'   End Select
'End Sub
'end 2020/2/14

'單筆列印
'edit by nickc 2006/02/27
Function PrintByOne(oCP01 As String, oCP02 As String, oCP03 As String, oCP04 As String, oType As String) As String
Dim returnRec As Long
Dim TmpR030006 As String
Dim tRS As New ADODB.Recordset
Dim tRS2 As New ADODB.Recordset
Dim NewAcc020021 As String
Dim StrSQLa As String
Dim strSQL1 As String
Dim StrSQL3 As String
'Dim TmpRule As String
'add by sonia 2016/6/2
Dim strCompNo As String
Dim strCompName As String

   PrintByOne = ""
   On Error GoTo RollBackData
   'add by nickc 2005/09/21
   m_CP01 = oCP01
   m_CP02 = oCP02
   m_CP03 = oCP03
   m_CP04 = oCP04
   strSQL1 = ""
   strSQLCP = "": strSQLnp = ""
   StrSQL3 = ""
   
   '判斷新案還是舊案 2011/3/10 移至共用
   IsOldSystem = Judgecase(oCP01, oCP02)
   
   '檢查本所案號是否存在
   Set tRS = New ADODB.Recordset
   With tRS
      tRS.CursorLocation = adUseClient
      tRS.Open "select * from caseprogress where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' ", cnnConnection, adOpenStatic, adLockReadOnly
      If tRS.RecordCount = 0 Then
         'edit by nickc 2006/02/27 狀態 1 此本所案號不存在
         PrintByOne = "1"
         Exit Function
      End If
   End With
   
   'CFP 前 7 碼相同 7-3(CFP)=4
   'TF 前 6 碼相同 6-2(TF)=4
   If oCP01 = "TF" Then
      strSQLCP = " and cp02='" & oCP02 & "' "
      '2011/6/1 add by sonia
      strSQLnp = " and np03='" & oCP02 & "' "
      oCP03 = "0": oCP04 = "00"
      '2011/6/1 end
   ElseIf oCP01 = "CFP" Then
      strSQLCP = " and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' "
      '2011/6/1 add by sonia
      strSQLnp = " and np03='" & oCP02 & "' and np04='" & oCP03 & "' "
      oCP04 = "00"
      '2011/6/1 end
   Else
      strSQLCP = " and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' "
      strSQLnp = " and np03='" & oCP02 & "' and np04='" & oCP03 & "' and np05='" & oCP04 & "' "  '2011/6/1 add by sonia
   End If
   
   'add by sonia 2025/7/8 把檢查從下面移上來，否則可結餘日後有收費但未收未付的案件會每個月重覆跑
      'add by sonia 2023/3/24 若可結餘日之後還有有收費且已開收據之進度則上cp146可結餘刪除日期,否則無法一併計算 CFT-022371
      Set tRS2 = New ADODB.Recordset
      'modify by sonia 2023/11/29 cp05>=max109改為cp05>max109，否則CFT-022256不會計算
      'modify by sonia 2025/5/23 Max(Cp109) Max109改檢查工作檔r040320_t1之R030006，否則CFP-027825不會計算
      'StrSQLa = "select * from caseprogress,(Select Max(Cp109) Max109 From Caseprogress Where cp01='" & oCP01 & "' " & strSQLCP & " And Cp109 Is Not Null), " & _
                "(SELECT a1u03,sum(a1u07) a1u07,sum(a1u09) a1u09 from ACC1U0 where A1U03 in (select CP09 from CASEPROGRESS where cp01='" & oCP01 & "' " & strSQLCP & " and CP16>0 ) group by a1u03) " & _
                "Where cp01='" & oCP01 & "' " & strSQLCP & " and cp05>max109 and cp09=a1u03(+) and nvl(cp16,0)-nvl(a1u07,0)-nvl(a1u09,0)>0 and cp60 is not null "
      'modify by sonia 2025/7/8 要區分TF案語法,取消CP03=R030003(+)，否則TF-000830不會檢查到
      'StrSQLa = "select * from caseprogress,r040320_t1, " & _
                "(SELECT a1u03,sum(a1u07) a1u07,sum(a1u09) a1u09 from ACC1U0 where A1U03 in (select CP09 from CASEPROGRESS where cp01='" & oCP01 & "' " & strSQLCP & " and CP16>0 ) group by a1u03) " & _
                "Where cp01='" & oCP01 & "' " & strSQLCP & " and CP01=R030001(+) and CP02=R030002(+) and CP03=R030003(+) and CP05>R030006 and cp09=a1u03(+) and nvl(cp16,0)-nvl(a1u07,0)-nvl(a1u09,0)>0 and cp60 is not null "
      If oCP01 = "TF" Then
         StrSQLa = "select * from caseprogress,r040320_t1, " & _
                   "(SELECT a1u03,sum(a1u07) a1u07,sum(a1u09) a1u09 from ACC1U0 where A1U03 in (select CP09 from CASEPROGRESS where cp01='" & oCP01 & "' " & strSQLCP & " and CP16>0 ) group by a1u03) " & _
                   "Where cp01='" & oCP01 & "' " & strSQLCP & " and CP01=R030001(+) and CP02=R030002(+) and CP05>R030006 and cp09=a1u03(+) and nvl(cp16,0)-nvl(a1u07,0)-nvl(a1u09,0)>0 and cp60 is not null "
      Else
         StrSQLa = "select * from caseprogress,r040320_t1, " & _
                   "(SELECT a1u03,sum(a1u07) a1u07,sum(a1u09) a1u09 from ACC1U0 where A1U03 in (select CP09 from CASEPROGRESS where cp01='" & oCP01 & "' " & strSQLCP & " and CP16>0 ) group by a1u03) " & _
                   "Where cp01='" & oCP01 & "' " & strSQLCP & " and CP01=R030001(+) and CP02=R030002(+) and CP03=R030003(+) and CP05>R030006 and cp09=a1u03(+) and nvl(cp16,0)-nvl(a1u07,0)-nvl(a1u09,0)>0 and cp60 is not null "
      End If
      'end 2025/7/8
      tRS2.CursorLocation = adUseClient
      tRS2.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If tRS2.RecordCount > 0 Then
         PrintByOne = "5"
         ClearData
         'cnnConnection.RollbackTrans   'cancel by sonia 2025/7/8 從下面移上來所以不用此句
         cnnConnection.BeginTrans
         'modify by sonia 2025/6/10 取消cp146 is null條件(否則P-066627會一直出現)
         'cnnConnection.Execute "update caseprogress set cp146='" & strSrvDate(1) & "' where cp01='" & oCP01 & "' " & strSQLCP & " and cp59 is null and cp146 is null and cp109 is not null "
         cnnConnection.Execute "update caseprogress set cp146='" & strSrvDate(1) & "' where cp01='" & oCP01 & "' " & strSQLCP & " and cp59 is null and cp109 is not null "
         cnnConnection.Execute "delete r040320_t1 where id='" & strUserNum & "' and r030001='" & oCP01 & "' and r030002='" & oCP02 & "' and r030003='" & oCP03 & "'"
         cnnConnection.CommitTrans
         Exit Function
      End If
      'end 2023/3/24
   'end 2025/7/8
   
   '檢查未付款未收款
   'add by nickc 2006/04/19 加入作廢不管
   'modify by sonia 2018/9/11 R107090100 CFT-018925 結匯中案件也不能算結餘
   StrSQLa = "select * from ("
   StrSQLa = StrSQLa & " select cp61 from caseprogress where cp01='" & oCP01 & "'  " & strSQLCP & " "
   StrSQLa = StrSQLa & " union select cp62 from caseprogress where cp01='" & oCP01 & "' " & strSQLCP & " "
   StrSQLa = StrSQLa & " union select cp63 from caseprogress where cp01='" & oCP01 & "' " & strSQLCP & " "
   StrSQLa = StrSQLa & " union select cp87 from caseprogress where cp01='" & oCP01 & "' " & strSQLCP & " "
   StrSQLa = StrSQLa & " union select cp88 from caseprogress where cp01='" & oCP01 & "' " & strSQLCP & "  ) AA,acc190,ACC150 where AA.cp61 is not null and AA.cp61=A1902(+) and AA.cp61=A1501(+) "
   'modify by sonia 2018/9/11 R107090100 CFT-018925 結匯中案件也不能算結餘
   'StrSQLa = StrSQLa & " and A1902 is null AND A1512 IS NULL and a1507 is null "
   StrSQLa = StrSQLa & " and (A1902 is null or a1908 is null) AND A1512 IS NULL and a1507 is null "
   Set tRS = New ADODB.Recordset
   With tRS
      .CursorLocation = adUseClient
      .Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         'If oType = "CP" Then   '2012/7/30 CANCEL BY SONIA CFT-013550(NP)
            'edit by nickc 2006/02/27 此案號還有帳單未付款
            PrintByOne = "2"
         'End If
         Exit Function
      End If
      'add by sonia 2013/5/17 add by sonia 抵帳單尚未抵帳() CFT-014314
      StrSQLa = "select acc161.* from acc161,acc160 where axg03='" & oCP01 & oCP02 & oCP03 & oCP04 & "' and axg01=a1601(+) and a1607 is null "
      If .State = 1 Then .Close
      .CursorLocation = adUseClient
      .Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         PrintByOne = "2"
         Exit Function
      End If
      '2013/5/17 end
      'add by nickc 2005/08/02
      '檢查國內未收款
      StrSQLa = "select cp60,cp79 from caseprogress where cp01='" & oCP01 & "' " & strSQLCP & " and cp60 is not null and substr(cp60,1,1)='E' and cp79>0 "
      'Set tRS = New ADODB.Recordset
      If .State = 1 Then .Close
      .CursorLocation = adUseClient
      .Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         'If oType = "CP" Then   '2012/7/30 CANCEL BY SONIA CFT-013550(NP)
            'edit by nickc 2006/02/27 此案號還有國內帳款未收
            PrintByOne = "3"
         'End If
         Exit Function
      End If
      '檢查國外未收款
      'edit by nickc 2007/03/12 秀玲說，若是 a1k29='Y' 就不管 acc0z0
      'StrSQLa = "select cp60,a1k19,acc0z0.* from caseprogress,acc0z0,acc1k0 where cp01='" & oCP01 & "' " & strSQL2 & "  and cp60=a0z02(+) and cp60=a1k01(+) and cp60 is not null and substr(cp60,1,1)='X' and (a0z02 is null or a1k29<>'Y' or a1k29 is null)"
      '2012/7/30 modify by sonia 銷帳不管,A1k25
      StrSQLa = "select cp60,a1k19,acc0z0.* from caseprogress,acc0z0,acc1k0 where cp01='" & oCP01 & "' " & strSQLCP & "  and cp60=a0z02(+) and cp60=a1k01(+) and cp60 is not null and substr(cp60,1,1)='X' and a1k25 is null and ((a0z02 is null and a1k29<>'Y') or a1k29<>'Y' or a1k29 is null)"
      If .State = 1 Then .Close
      .CursorLocation = adUseClient
      .Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         'If oType = "CP" Then   '2012/7/30 CANCEL BY SONIA CFT-013550(NP)
            'edit by nickc 2006/02/27 此案號還有國外帳款未收
            PrintByOne = "4"
         'End If
         Exit Function
      End If
      
      '2011/3/11 add by sonia 若前有未結算結餘單則作廢,此次全部重算
cnnConnection.BeginTrans

'cancel by sonia 2025/7/8 移到未付款未收款之前先做
'      'add by sonia 2023/3/24 若可結餘日之後還有有收費且已開收據之進度則上cp146可結餘刪除日期,否則無法一併計算 CFT-022371
'      Set tRS2 = New ADODB.Recordset
'      'modify by sonia 2023/11/29 cp05>=max109改為cp05>max109，否則CFT-022256不會計算
'      'modify by sonia 2025/5/23 Max(Cp109) Max109改檢查工作檔r040320_t1之R030006，否則CFP-027825不會計算
'      'StrSQLa = "select * from caseprogress,(Select Max(Cp109) Max109 From Caseprogress Where cp01='" & oCP01 & "' " & strSQLCP & " And Cp109 Is Not Null), " & _
'                "(SELECT a1u03,sum(a1u07) a1u07,sum(a1u09) a1u09 from ACC1U0 where A1U03 in (select CP09 from CASEPROGRESS where cp01='" & oCP01 & "' " & strSQLCP & " and CP16>0 ) group by a1u03) " & _
'                "Where cp01='" & oCP01 & "' " & strSQLCP & " and cp05>max109 and cp09=a1u03(+) and nvl(cp16,0)-nvl(a1u07,0)-nvl(a1u09,0)>0 and cp60 is not null "
'      'modify by sonia 2025/7/8 要區分TF案語法,取消CP03=R030003(+)，否則TF-000830不會檢查到
'      'StrSQLa = "select * from caseprogress,r040320_t1, " & _
'                "(SELECT a1u03,sum(a1u07) a1u07,sum(a1u09) a1u09 from ACC1U0 where A1U03 in (select CP09 from CASEPROGRESS where cp01='" & oCP01 & "' " & strSQLCP & " and CP16>0 ) group by a1u03) " & _
'                "Where cp01='" & oCP01 & "' " & strSQLCP & " and CP01=R030001(+) and CP02=R030002(+) and CP03=R030003(+) and CP05>R030006 and cp09=a1u03(+) and nvl(cp16,0)-nvl(a1u07,0)-nvl(a1u09,0)>0 and cp60 is not null "
'      If oCP01 = "TF" Then
'         StrSQLa = "select * from caseprogress,r040320_t1, " & _
'                   "(SELECT a1u03,sum(a1u07) a1u07,sum(a1u09) a1u09 from ACC1U0 where A1U03 in (select CP09 from CASEPROGRESS where cp01='" & oCP01 & "' " & strSQLCP & " and CP16>0 ) group by a1u03) " & _
'                   "Where cp01='" & oCP01 & "' " & strSQLCP & " and CP01=R030001(+) and CP02=R030002(+) and CP05>R030006 and cp09=a1u03(+) and nvl(cp16,0)-nvl(a1u07,0)-nvl(a1u09,0)>0 and cp60 is not null "
'      Else
'         StrSQLa = "select * from caseprogress,r040320_t1, " & _
'                   "(SELECT a1u03,sum(a1u07) a1u07,sum(a1u09) a1u09 from ACC1U0 where A1U03 in (select CP09 from CASEPROGRESS where cp01='" & oCP01 & "' " & strSQLCP & " and CP16>0 ) group by a1u03) " & _
'                   "Where cp01='" & oCP01 & "' " & strSQLCP & " and CP01=R030001(+) and CP02=R030002(+) and CP03=R030003(+) and CP05>R030006 and cp09=a1u03(+) and nvl(cp16,0)-nvl(a1u07,0)-nvl(a1u09,0)>0 and cp60 is not null "
'      End If
'      'end 2025/7/8
'      tRS2.CursorLocation = adUseClient
'      tRS2.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'      If tRS2.RecordCount > 0 Then
'         PrintByOne = "5"
'         ClearData
'         cnnConnection.RollbackTrans
'         cnnConnection.BeginTrans
'         'modify by sonia 2025/6/10 取消cp146 is null條件(否則P-066627會一直出現)
'         'cnnConnection.Execute "update caseprogress set cp146='" & strSrvDate(1) & "' where cp01='" & oCP01 & "' " & strSQLCP & " and cp59 is null and cp146 is null and cp109 is not null "
'         cnnConnection.Execute "update caseprogress set cp146='" & strSrvDate(1) & "' where cp01='" & oCP01 & "' " & strSQLCP & " and cp59 is null and cp109 is not null "
'         cnnConnection.Execute "delete r040320_t1 where id='" & strUserNum & "' and r030001='" & oCP01 & "' and r030002='" & oCP02 & "' and r030003='" & oCP03 & "'"
'         cnnConnection.CommitTrans
'         Exit Function
'      End If
'      'end 2023/3/24
'end 2025/7/8
      
      If oCP01 = "CFP" Then
         strSql = "SELECT a240002 FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and (a240003 is null or a240003=0) and (a240015 is null or a240015=0) "
      ElseIf oCP01 = "TF" Then
         strSql = "SELECT a240002 FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and (a240003 is null or a240003=0) and (a240015 is null or a240015=0) "
      Else
         strSql = "SELECT a240002 FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and A240008='" & oCP04 & "'  and (a240003 is null or a240003=0) and (a240015 is null or a240015=0) "
      End If
      CheckOC2
      Set tRS = New ADODB.Recordset
      tRS.CursorLocation = adUseClient
      tRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      Do While Not tRS.EOF
         'modify by sonia 2025/6/18 作廢人改為QPGMR以利區別是程式作廢的
         'cnnConnection.Execute "update ACC240 set a240003=" & strSrvDate(2) & ",A240016='" & strUserNum & "' where a240002='" & tRS.Fields(0) & "' "
         cnnConnection.Execute "update ACC240 set a240003=" & strSrvDate(2) & ",A240016='QPGMR' where a240002='" & tRS.Fields(0) & "' "
         cnnConnection.Execute "update caseprogress set cp59=null where cp59='" & tRS.Fields(0) & "' "
         tRS.MoveNext
      Loop
      '2011/3/11 end
      
      '抓最近結餘日期
      MaxDay = ""
      'Modified by Morgan 2017/10/6 若無資料時Win7的電腦可能會發生錯誤,語法+having MAX(a240001)>0
      If oCP01 = "CFP" Then
         '2010/3/26 MODIFY BY SONIA 僅EPC的子案與母案合併,接續案不可合併,集體設計暫時也不合併
         'strSQL = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and (a240003 is null or a240003=0) "
         strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and (a240003 is null or a240003=0) having MAX(a240001)>0"
      ElseIf oCP01 = "TF" Then
         '2010/3/26 MODIFY BY SONIA 馬德里案母案與延土延伸分開算
         'strSQL = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and (a240003 is null or a240003=0) "
         strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and (a240003 is null or a240003=0) having MAX(a240001)>0"
      Else
         strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and A240008='" & oCP04 & "'  and (a240003 is null or a240003=0) having MAX(a240001)>0"
      End If
      'end 2017/10/6
      CheckOC2
      Set tRS = New ADODB.Recordset
      tRS.CursorLocation = adUseClient
      tRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If Not tRS.EOF And Not tRS.BOF Then
         MaxDay = CheckStr(tRS.Fields(0))
      End If
      
      '檢查是否可以結餘
      'CFP 前 7 碼相同 7-3(CFP)=4
      'TF 前 6 碼相同 6-2(TF)=4
      '2012/7/12 modify by sonia 加可結餘刪除日期,cp146 is null條件
      If oCP01 = "TF" Then
         'edit by nickc 2005/08/01
         'strSQL1 = " and C1.cp01='" & oCP01 & "' and c1.cp02='" & oCP02 & "'  "
         '2010/11/30 modify by sonia 加c2.cp59 is null
         'modify by sonia 2025/4/28 c2.cp146 is null改為(c2.cp146 is null or c2.cp109>c2.cp146)
         strSQL1 = " and ax214>='" & oCP01 & oCP02 & "000' and ax214<='" & oCP01 & oCP02 & "ZZZ' and C2.cp01='" & oCP01 & "' and c2.cp02='" & oCP02 & "' and c2.cp59 is null and (c2.cp146 is null or c2.cp109>c2.cp146) and (c2.cp60 is not null or nvl(axf04,0)>0) " '& IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
         StrSQL3 = " and ax214>='" & oCP01 & oCP02 & "000' and ax214<='" & oCP01 & oCP02 & "ZZZ'  " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
      ElseIf oCP01 = "CFP" Then
         'edit by nickc 2005/08/01
         'strSQL1 = " and C1.cp01='" & oCP01 & "' and c1.cp02='" & oCP02 & "' and c1.cp03='" & oCP03 & "' "
         '2010/11/30 modify by sonia 加c2.cp59 is null
         'modify by sonia 2025/4/28 c2.cp146 is null改為(c2.cp146 is null or c2.cp109>c2.cp146)
         strSQL1 = " and ax214>='" & oCP01 & oCP02 & oCP03 & "00' and ax214<='" & oCP01 & oCP02 & oCP03 & "ZZ' and C2.cp01='" & oCP01 & "' and c2.cp02='" & oCP02 & "' and c2.cp03='" & oCP03 & "' and c2.cp59 is null and (c2.cp146 is null or c2.cp109>c2.cp146) and (c2.cp60 is not null or nvl(axf04,0)>0) " '& IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
         StrSQL3 = " and ax214>='" & oCP01 & oCP02 & oCP03 & "00' and ax214<='" & oCP01 & oCP02 & oCP03 & "ZZ'  " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
      Else
         'edit by nickc 2005/08/01
         'strSQL1 = " and C1.cp01='" & oCP01 & "' and c1.cp02='" & oCP02 & "' and c1.cp03='" & oCP03 & "' and c1.cp04='" & oCP04 & "' "
         '2010/11/30 modify by sonia 加c2.cp59 is null
         'modify by sonia 2025/4/28 c2.cp146 is null改為(c2.cp146 is null or c2.cp109>c2.cp146)
         strSQL1 = " and ax214='" & oCP01 & oCP02 & oCP03 & oCP04 & "' and C2.cp01='" & oCP01 & "' and c2.cp02='" & oCP02 & "' and c2.cp03='" & oCP03 & "' and c2.cp04='" & oCP04 & "' and c2.cp59 is null and (c2.cp146 is null or c2.cp109>c2.cp146) and (c2.cp60 is not null or nvl(axf04,0)>0) " '& IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
         StrSQL3 = " and ax214='" & oCP01 & oCP02 & oCP03 & oCP04 & "'  " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
      End If
      
      'edit by nickc 2007/11/16 浮動準備金,結餘時才可以清 0
      'DayTemp = GetFloatPrepareCase(oCP01, oCP02, oCP03, oCP04)
      DayTemp = GetFloatPrepareCase(oCP01, oCP02, oCP03, oCP04, True)
      'add by sonia 2016/6/2 抓出名公司
      strCompNo = ""
      strCompName = GetSpecialComp(oCP01, oCP02, oCP03, oCP04, strCompNo, 6)
      'end 2016/6/2
      
      '2010/11/26 MODIFY BY SONIA 剔除結餘傳票,否則第二次以上的結餘會抓到
      'modify by sonia 2016/6/2 只抓該出名公司的傳票,其他公司寫入acc242 (CFT-017099)
      'NewAcc020021 = "(select ax202,ax214,newa1,newa2,newa3,newa4 from (" & _
                    " select ax202,ax212,ax214,sum(A1) as newA1,sum(A2) as newA2,sum(A3) as newA3,sum(A4) as newA4 from (" & _
                    " select ax202,ax212,ax214, " & _
                    " (DECODE(substr(ax205,1,1),'4',nvl(ax207,0)-nvl(ax206,0),decode(substr(ax205,1,4),'2201',decode(instr(ax212,'退費'),0,nvl(ax207,0),0),0))) as A1, " & _
                    "  (decode(substr(ax205,1,1),'4',nvl(ax207,0)-nvl(ax206,0),0)) as A2, " & _
                    " (decode(substr(ax205,1,4),'2201',decode(ax206,0,0,nvl(ax206,0)))) as A3, " & _
                    " (decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)),0))) as A4 " & _
                    " From acc020, acc021 where  ax201=a0201(+) and ax202=a0202(+) AND INSTR(AX212,'結餘')=0 " & StrSQL3 & ") NewTable " & _
                    " group by NewTable.ax202,NewTable.ax212,NewTable.ax214) NewTable2) NewTable3 "
      'modify by sonia 2020/4/16 +L公司
      NewAcc020021 = "(select ax201,ax202,ax214,newa1,newa2,newa3,newa4 from (" & _
                    " select ax201,ax202,ax212,ax214,sum(A1) as newA1,sum(A2) as newA2,sum(A3) as newA3,sum(A4) as newA4 from (" & _
                    " select ax201,ax202,ax212,ax214, " & _
                    " (DECODE(substr(ax205,1,1),'4',nvl(ax207,0)-nvl(ax206,0),decode(substr(ax205,1,4),'2201',decode(instr(ax212,'退費'),0,nvl(ax207,0),0),0))) as A1, " & _
                    "  (decode(substr(ax205,1,1),'4',nvl(ax207,0)-nvl(ax206,0),0)) as A2, " & _
                    " (decode(substr(ax205,1,4),'2201',decode(ax206,0,0,nvl(ax206,0)))) as A3, " & _
                    " (decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)),0))) as A4 " & _
                    " From acc020, acc021 where a0201=decode('" & strCompNo & "','J','J','L','L','1') and  ax201=a0201(+) and ax202=a0202(+) AND INSTR(AX212,'結餘')=0 " & StrSQL3 & ") NewTable " & _
                    " group by NewTable.ax201,NewTable.ax202,NewTable.ax212,NewTable.ax214) NewTable2) NewTable3 "
      
      cnnConnection.Execute "delete from r040320_t where id='" & strUserNum & "' "

'2011/4/6 MODIFY BY SONIA 不分以a1p04的第一碼判斷國內外收款且不可取消案號的(+),否則結匯的資料會抓不到
      '2011/3/16 modify by sonia 以a1p04的第一碼判斷國內外收款,取消案號的(+)
'      '國內收款 F
'      StrSQLa = "SELECT distinct C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04,NVL(PA05,NVL(PA06,PA07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),PA26),c1.CP09,NVL(ST02,c1.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,PATENT,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C1.CP01=PA01 AND C1.CP02=PA02 AND C1.CP03=PA03 AND C1.CP04=PA04 AND c1.CP59 IS NULL AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTr(PA26,9,1)) = CU02(+) AND c1.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL and pa09=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and substr(a1p04,1,1)='F' and a1p04=a1u01(+) and a1u03=c1.cp09(+)    AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null)) "
'      StrSQLa = StrSQLa & " union  select C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04,NVL(tm05,NVL(tm06,tm07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),tm23),c1.CP09,NVL(ST02,c1.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,trademark,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C1.CP01=TM01 AND C1.CP02=TM02 AND C1.CP03=TM03 AND C1.CP04=TM04 AND c1.CP59 IS NULL AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),'','0',SUBSTr(tm23,9,1)) = CU02(+) AND c1.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL and tm10=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and substr(a1p04,1,1)='F' and a1p04=a1u01(+) and a1u03=c1.cp09(+)     AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
'      StrSQLa = StrSQLa & " union  select C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04,NVL(sp05,NVL(sp06,sp07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),sp08),c1.CP09,NVL(ST02,c1.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,servicepractice,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C1.CP01=SP01 AND C1.CP02=SP02 AND C1.CP03=SP03 AND C1.CP04=SP04 AND c1.CP59 IS NULL AND SUBSTR(sp08,1,8)=CU01(+) AND DECODE(SUBSTR(sp08,9,1),'','0',SUBSTr(sp08,9,1)) = CU02(+) AND c1.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL   and sp09=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and substr(a1p04,1,1)='F' and a1p04=a1u01(+) and a1u03=c1.cp09(+)     AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
'      StrSQLa = StrSQLa & " union  select C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04,NVL(LC05,NVL(LC06,LC07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),LC11),c1.CP09,NVL(ST02,c1.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,LAWCASE,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C1.CP01=LC01 AND C1.CP02=LC02 AND C1.CP03=LC03 AND C1.CP04=LC04 AND c1.CP59 IS NULL AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),'','0',SUBSTr(LC11,9,1)) = CU02(+) AND c1.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL   and lc15=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and substr(a1p04,1,1)='F' and a1p04=a1u01(+) and a1u03=c1.cp09(+)     AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
'      '國外收款 M
'      StrSQLa = StrSQLa & " union  select C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04,NVL(PA05,NVL(PA06,PA07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),PA26),c1.CP09,NVL(ST02,c1.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,PATENT,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc0z0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C1.CP01=PA01 AND C1.CP02=PA02 AND C1.CP03=PA03 AND C1.CP04=PA04 AND c1.CP59 IS NULL AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTr(PA26,9,1)) = CU02(+) AND c1.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL and pa09=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and substr(a1p04,1,1)='M' and a1p04=a0z01(+) and a0z02=c1.cp60(+)     AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
'      StrSQLa = StrSQLa & " union  select C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04,NVL(tm05,NVL(tm06,tm07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),tm23),c1.CP09,NVL(ST02,c1.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,trademark,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc0z0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C1.CP01=TM01 AND C1.CP02=TM02 AND C1.CP03=TM03 AND C1.CP04=TM04 AND c1.CP59 IS NULL AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),'','0',SUBSTr(tm23,9,1)) = CU02(+) AND c1.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL and tm10=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and substr(a1p04,1,1)='M' and a1p04=a0z01(+) and a0z02=c1.cp60(+)      AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
'      StrSQLa = StrSQLa & " union  select C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04,NVL(sp05,NVL(sp06,sp07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),sp08),c1.CP09,NVL(ST02,c1.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,servicepractice,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc0z0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C1.CP01=SP01 AND C1.CP02=SP02 AND C1.CP03=SP03 AND C1.CP04=SP04 AND c1.CP59 IS NULL AND SUBSTR(sp08,1,8)=CU01(+) AND DECODE(SUBSTR(sp08,9,1),'','0',SUBSTr(sp08,9,1)) = CU02(+) AND c1.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL   and sp09=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and substr(a1p04,1,1)='M' and a1p04=a0z01(+) and a0z02=c1.cp60(+)     AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
'      StrSQLa = StrSQLa & " union  select C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04,NVL(LC05,NVL(LC06,LC07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),LC11),c1.CP09,NVL(ST02,c1.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,LAWCASE,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc0z0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C1.CP01=LC01 AND C1.CP02=LC02 AND C1.CP03=LC03 AND C1.CP04=LC04 AND c1.CP59 IS NULL AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),'','0',SUBSTr(LC11,9,1)) = CU02(+) AND c1.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL   and lc15=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and substr(a1p04,1,1)='M' and a1p04=a0z01(+) and a0z02=c1.cp60(+)      AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
      '2011/12/21 modify by sonia 因原寫法會抓不到C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04及C1.CP09 導致檢查大陸商標是否只有註冊證會抓不到而跳離開,故改C1為C2
'      StrSQLa = "SELECT distinct C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04,NVL(PA05,NVL(PA06,PA07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),PA26),c1.CP09,NVL(ST02,c1.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,PATENT,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C1.CP01=PA01(+) AND C1.CP02=PA02(+) AND C1.CP03=PA03(+) AND C1.CP04=PA04(+) AND c1.CP59 IS NULL AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTr(PA26,9,1)) = CU02(+) AND c1.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL and pa09=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null)) "
'      StrSQLa = StrSQLa & " union  select C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04,NVL(tm05,NVL(tm06,tm07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),tm23),c1.CP09,NVL(ST02,c1.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,trademark,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND c1.CP59 IS NULL AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),'','0',SUBSTr(tm23,9,1)) = CU02(+) AND c1.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL and tm10=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
'      StrSQLa = StrSQLa & " union  select C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04,NVL(sp05,NVL(sp06,sp07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),sp08),c1.CP09,NVL(ST02,c1.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,servicepractice,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C1.CP01=SP01(+) AND C1.CP02=SP02(+) AND C1.CP03=SP03(+) AND C1.CP04=SP04(+) AND c1.CP59 IS NULL AND SUBSTR(sp08,1,8)=CU01(+) AND DECODE(SUBSTR(sp08,9,1),'','0',SUBSTr(sp08,9,1)) = CU02(+) AND c1.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL   and sp09=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
'      StrSQLa = StrSQLa & " union  select C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04,NVL(LC05,NVL(LC06,LC07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),LC11),c1.CP09,NVL(ST02,c1.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,LAWCASE,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C1.CP01=LC01(+) AND C1.CP02=LC02(+) AND C1.CP03=LC03(+) AND C1.CP04=LC04(+) AND c1.CP59 IS NULL AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),'','0',SUBSTr(LC11,9,1)) = CU02(+) AND c1.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL   and lc15=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
      '2012/7/12 modify by sonia 加可結餘刪除日期,cp146 is null條件
'2012/7/30 modify by sonia 取消AND A1507 IS NULL , CFP-014718(A98006589)帳單
'      StrSQLa = "SELECT distinct C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04,NVL(PA05,NVL(PA06,PA07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),PA26),c2.CP09,NVL(ST02,c2.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,PATENT,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C2.CP01=PA01(+) AND C2.CP02=PA02(+) AND C2.CP03=PA03(+) AND C2.CP04=PA04(+) AND c1.CP59 IS NULL and c1.cp146 is null AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTr(PA26,9,1)) = CU02(+) AND c2.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL and pa09=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null)) "
'      StrSQLa = StrSQLa & " union  select C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04,NVL(tm05,NVL(tm06,tm07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),tm23),c2.CP09,NVL(ST02,c2.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,trademark,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C2.CP01=TM01(+) AND C2.CP02=TM02(+) AND C2.CP03=TM03(+) AND C2.CP04=TM04(+) AND c1.CP59 IS NULL and c1.cp146 is null AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),'','0',SUBSTr(tm23,9,1)) = CU02(+) AND c2.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL and tm10=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
'      StrSQLa = StrSQLa & " union  select C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04,NVL(sp05,NVL(sp06,sp07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),sp08),c2.CP09,NVL(ST02,c2.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,servicepractice,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C2.CP01=SP01(+) AND C2.CP02=SP02(+) AND C2.CP03=SP03(+) AND C2.CP04=SP04(+) AND c1.CP59 IS NULL and c1.cp146 is null AND SUBSTR(sp08,1,8)=CU01(+) AND DECODE(SUBSTR(sp08,9,1),'','0',SUBSTr(sp08,9,1)) = CU02(+) AND c2.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL   and sp09=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
'      StrSQLa = StrSQLa & " union  select C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04,NVL(LC05,NVL(LC06,LC07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),LC11),c2.CP09,NVL(ST02,c2.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
'                    " FROM CASEPROGRESS C1,LAWCASE,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
'                    " WHERE C2.CP01=LC01(+) AND C2.CP02=LC02(+) AND C2.CP03=LC03(+) AND C2.CP04=LC04(+) AND c1.CP59 IS NULL and c1.cp146 is null AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),'','0',SUBSTr(LC11,9,1)) = CU02(+) AND c2.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) AND A1507 IS NULL   and lc15=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
      'modify by sonia 2023/8/10 依系統類別拆語法，不必全串在一起
      'StrSQLa = "SELECT distinct C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04,NVL(PA05,NVL(PA06,PA07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),PA26),c2.CP09,NVL(ST02,c2.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
      '              " FROM CASEPROGRESS C1,PATENT,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
      '              " WHERE C2.CP01=PA01(+) AND C2.CP02=PA02(+) AND C2.CP03=PA03(+) AND C2.CP04=PA04(+) AND c1.CP59 IS NULL and c1.cp146 is null AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTr(PA26,9,1)) = CU02(+) AND c2.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) and pa09=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null)) "
      'StrSQLa = StrSQLa & " union  select C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04,NVL(tm05,NVL(tm06,tm07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),tm23),c2.CP09,NVL(ST02,c2.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
      '              " FROM CASEPROGRESS C1,trademark,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
      '              " WHERE C2.CP01=TM01(+) AND C2.CP02=TM02(+) AND C2.CP03=TM03(+) AND C2.CP04=TM04(+) AND c1.CP59 IS NULL and c1.cp146 is null AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),'','0',SUBSTr(tm23,9,1)) = CU02(+) AND c2.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) and tm10=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
      'StrSQLa = StrSQLa & " union  select C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04,NVL(sp05,NVL(sp06,sp07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),sp08),c2.CP09,NVL(ST02,c2.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
      '              " FROM CASEPROGRESS C1,servicepractice,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
      '              " WHERE C2.CP01=SP01(+) AND C2.CP02=SP02(+) AND C2.CP03=SP03(+) AND C2.CP04=SP04(+) AND c1.CP59 IS NULL and c1.cp146 is null AND SUBSTR(sp08,1,8)=CU01(+) AND DECODE(SUBSTR(sp08,9,1),'','0',SUBSTr(sp08,9,1)) = CU02(+) AND c2.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) and sp09=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
      'StrSQLa = StrSQLa & " union  select C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04,NVL(LC05,NVL(LC06,LC07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),LC11),c2.CP09,NVL(ST02,c2.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
      '              " FROM CASEPROGRESS C1,LAWCASE,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
      '              " WHERE C2.CP01=LC01(+) AND C2.CP02=LC02(+) AND C2.CP03=LC03(+) AND C2.CP04=LC04(+) AND c1.CP59 IS NULL and c1.cp146 is null AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),'','0',SUBSTr(LC11,9,1)) = CU02(+) AND c2.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) and lc15=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+)   and ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
      Select Case oCP01
         Case "CFP", "P"
            'modify by sonia 2025/4/28 c1.cp146 is null改為(c1.cp146 is null or c1.cp109>c1.cp146)
            StrSQLa = "SELECT distinct C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04,NVL(PA05,NVL(PA06,PA07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),PA26),c2.CP09,NVL(ST02,c2.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
                       " FROM CASEPROGRESS C1,PATENT,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
                       " WHERE C2.CP01=PA01(+) AND C2.CP02=PA02(+) AND C2.CP03=PA03(+) AND C2.CP04=PA04(+) AND c1.CP59 IS NULL and (c1.cp146 is null or c1.cp109>c1.cp146) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTr(PA26,9,1)) = CU02(+) AND c2.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) and pa09=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+) and ax201=a1p01(+) and ax202=a1p22(+) and 'A'=a1p02(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null)) "
         Case "T", "TF", "CFT"
            'modify by sonia 2025/4/28 c1.cp146 is null改為(c1.cp146 is null or c1.cp109>c1.cp146)
            StrSQLa = "select distinct C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04,NVL(tm05,NVL(tm06,tm07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),tm23),c2.CP09,NVL(ST02,c2.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
                       " FROM CASEPROGRESS C1,trademark,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
                       " WHERE C2.CP01=TM01(+) AND C2.CP02=TM02(+) AND C2.CP03=TM03(+) AND C2.CP04=TM04(+) AND c1.CP59 IS NULL and (c1.cp146 is null or c1.cp109>c1.cp146) AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),'','0',SUBSTr(tm23,9,1)) = CU02(+) AND c2.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) and tm10=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+) and ax201=a1p01(+) and ax202=a1p22(+) and 'A'=a1p02(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
         Case "L", "CFL"
            'modify by sonia 2025/4/28 c1.cp146 is null改為(c1.cp146 is null or c1.cp109>c1.cp146)
            StrSQLa = "select distinct C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04,NVL(LC05,NVL(LC06,LC07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),LC11),c2.CP09,NVL(ST02,c2.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
                       " FROM CASEPROGRESS C1,LAWCASE,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
                       " WHERE C2.CP01=LC01(+) AND C2.CP02=LC02(+) AND C2.CP03=LC03(+) AND C2.CP04=LC04(+) AND c1.CP59 IS NULL and (c1.cp146 is null or c1.cp109>c1.cp146) AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),'','0',SUBSTr(LC11,9,1)) = CU02(+) AND c2.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) and lc15=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+) and ax201=a1p01(+) and ax202=a1p22(+) and 'A'=a1p02(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
         Case Else
            'modify by sonia 2025/4/28 c1.cp146 is null改為(c1.cp146 is null or c1.cp109>c1.cp146)
             StrSQLa = "select distinct C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04,NVL(sp05,NVL(sp06,sp07)),NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),sp08),c2.CP09,NVL(ST02,c2.CP13),NVL(a1.CPM04,c1.CP10)," & SQLDate("c1.CP27") & ",newa1,newa2,A1505,AXF04,nvl(fa04,decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),newa3," & DayTemp & ",0,na03,ax202,'" & strUserNum & "',NVL(A2.CPM04,c2.CP10),newa4   " & _
                       " FROM CASEPROGRESS C1,servicepractice,CUSTOMER,STAFF,CASEPROPERTYMAP A1,ACC150,ACC151,fagent,nation," & NewAcc020021 & ",acc1p0,acc1u0,CASEPROGRESS c2,CASEPROPERTYMAP A2 " & _
                       " WHERE C2.CP01=SP01(+) AND C2.CP02=SP02(+) AND C2.CP03=SP03(+) AND C2.CP04=SP04(+) AND c1.CP59 IS NULL and (c1.cp146 is null or c1.cp109>c1.cp146) AND SUBSTR(sp08,1,8)=CU01(+) AND DECODE(SUBSTR(sp08,9,1),'','0',SUBSTr(sp08,9,1)) = CU02(+) AND c2.CP13=ST01(+) AND c1.CP01=a1.CPM01(+) AND c1.CP10=a1.CPM02(+) AND c2.CP09=AXF02(+) AND AXF01=A1501(+) and sp09=na01(+) and substr(c1.cp44,1,8)=fa01(+) and substr(c1.cp44,9,1)=fa02(+) and ax201=a1p01(+) and ax202=a1p22(+) and 'A'=a1p02(+) and a1p04=a1u01(+) and a1u03=c1.cp09(+) AND c2.CP01=A2.CPM01(+) AND c2.CP10=A2.CPM02(+) " & strSQL1 & " and ax214=a1p17(+) and (ax214=c1.cp01||c1.cp02||c1.cp03||c1.cp04 or (c1.cp01 is null))"
      End Select
      'end 2023/8/10
'2012/7/30 END
      '2011/12/21 end
'2011/4/6 END
      cnnConnection.Execute "insert into r040320_t " & StrSQLa, returnRec
      If returnRec = 0 Then
         If oType = "CP" Then
            PrintByOne = "5"  '此案總帳傳票資料已平衡,不需再算結餘,為免以後重覆檢查故上cp146
            ClearData
         '2012/11/29 add by sonia P-055364
            cnnConnection.RollbackTrans   '2011/3/14 add by sonia
            cnnConnection.BeginTrans
            'modify by sonia 2016/2/24 取消and cp27 is not null條件
            'modify by sonia 2025/6/10 取消cp146 is null條件(否則P-066627會一直出現)
            'cnnConnection.Execute "update caseprogress set cp146='" & strSrvDate(1) & "' where cp01='" & oCP01 & "' " & strSQLCP & " and cp59 is null and cp146 is null and cp109 is not null "
            cnnConnection.Execute "update caseprogress set cp146='" & strSrvDate(1) & "' where cp01='" & oCP01 & "' " & strSQLCP & " and cp59 is null and cp109 is not null "
            cnnConnection.Execute "delete r040320_t1 where id='" & strUserNum & "' and r030001='" & oCP01 & "' and r030002='" & oCP02 & "' and r030003='" & oCP03 & "'"
            cnnConnection.CommitTrans
         Else
            cnnConnection.RollbackTrans   '2011/3/14 add by sonia
            'add by sonia 2025/5/29
            cnnConnection.BeginTrans
            cnnConnection.Execute "update nextprogress set np25=19221111 where np02='" & oCP01 & "' " & strSQLnp & _
                         " and np06 Is Null and NVL(NP25,0)=0 " & NPsql & _
                         " and ((NP02||'' in ('P','CFP','CPS','PS') and NP09<=TO_NUMBER(TO_CHAR(ADD_MONTHS(sysdate,-6),'YYYYMMDD'))) " & _
                         "  or (NP02||'' in ('T','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT') and NP09<=TO_NUMBER(TO_CHAR(ADD_MONTHS(sysdate,-3),'YYYYMMDD'))))", intI
            cnnConnection.Execute "delete r040320_t1 where id='" & strUserNum & "' and r030001='" & oCP01 & "' and r030002='" & oCP02 & "' and r030003='" & oCP03 & "'"
            cnnConnection.CommitTrans
            'end 2025/5/29
         '2012/11/29 end
         End If
         Exit Function
      '2011/3/17 ADD BY SONIA 婧瑄說大陸商標若只有註冊證
      '2011/11/1 MODIFY BY SONIA 舊系統資料無法由收款傳票抓回CASEPROGESS資料,所以R030001及R030004會存空值,故舊系統案號不做
      'ElseIf oCP01 = "T" Then
      ElseIf oCP01 = "T" And IsOldSystem = False Then
         StrSQLa = "select * from r040320_t where R030001='" & oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04 & "' and R030004<'C' "
         Set tRS2 = New ADODB.Recordset
         tRS2.CursorLocation = adUseClient
         tRS2.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If tRS2.RecordCount = 0 Then
            PrintByOne = "5"
            ClearData
            cnnConnection.RollbackTrans
            '2012/7/12 add by sonia 上cp146可結餘刪除日期
            cnnConnection.BeginTrans
            'modify by sonia 2016/2/24 取消and cp27 is not null條件
            'modify by sonia 2025/6/10 取消cp146 is null條件(否則P-066627會一直出現)
            'cnnConnection.Execute "update caseprogress set cp146='" & strSrvDate(1) & "' where cp01='" & oCP01 & "' " & strSQLCP & " and cp59 is null and cp146 is null and cp109 is not null "
            cnnConnection.Execute "update caseprogress set cp146='" & strSrvDate(1) & "' where cp01='" & oCP01 & "' " & strSQLCP & " and cp59 is null and cp109 is not null "
            cnnConnection.Execute "delete r040320_t1 where id='" & strUserNum & "' and r030001='" & oCP01 & "' and r030002='" & oCP02 & "' and r030003='" & oCP03 & "'"
            cnnConnection.CommitTrans
            '2012/7/12 end
           Exit Function
         End If
      '2011/3/17 END
      End If
   End With

''重新運算明細資料，以免金額重複會不見  使用 strsql3
Dim tmpA1p04 As String
Dim SeekY As Integer
   'edit by nickc 2006/03/01 補有些會串不到的情況 沒有 4 開頭的科目為結匯資料
   'StrSQLa = "select distinct r030017 from  r040320_t where id='" & strUserNum & "'  "
   '2010/11/30 modify by sonia 剔除前次結餘傳票及結餘結算傳票
   'StrSQLa = "select distinct a1p04 from acc1p0 where a1p22 in (select distinct ax202 from acc021 where substr(ax205,1,4)='2201' and ax206>0 and ax214='" & oCP01 & oCP02 & oCP03 & oCP04 & "') and A1p17='" & oCP01 & oCP02 & oCP03 & oCP04 & "'  "
   StrSQLa = "select distinct a1p04 from acc1p0 where a1p22 in (select distinct ax202 from acc021,acc020 where substr(ax205,1,4)='2201' and ax206>0 and ax214='" & oCP01 & oCP02 & oCP03 & oCP04 & "' and instr(ax212,'結餘')=0 and ax201=a0201(+) and ax202=a0202(+)" & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ") & ") and A1p17='" & oCP01 & oCP02 & oCP03 & oCP04 & "' "
   Set tRS = New ADODB.Recordset
   tRS.CursorLocation = adUseClient
   tRS.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   TmpR030006 = ""
   If tRS.RecordCount <> 0 Then
      tRS.MoveFirst
      Do While Not tRS.EOF
         tmpA1p04 = CheckStr(tRS.Fields(0))
         SeekY = 0
         '搜尋最後一個 Y
         Do While InStr(1, UCase(tmpA1p04), "Y") <> 0
             SeekY = SeekY + InStr(1, UCase(tmpA1p04), "Y")
             tmpA1p04 = Mid(tmpA1p04, InStr(1, UCase(tmpA1p04), "Y") + 1)
         Loop
         If SeekY <> 0 Then
             tmpA1p04 = Mid(CheckStr(tRS.Fields(0)), 1, SeekY - 1)
         End If
         Set tRS2 = New ADODB.Recordset
         If Mid(UCase(tmpA1p04), 1, 1) <> "Z" Then
             '2010/12/29 MODIFY BY SONIA 依申請國家抓案件性質
             'StrSQLa = "select cpm03,a1903,axf04 from acc190,acc151,caseprogress,casepropertymap where a1908='" & tmpA1p04 & "' and a1902=axf01(+) and axf03='" & oCP01 & oCP02 & oCP03 & oCP04 & "' and axf02=cp09(+) and cp01=cpm01(+) and cp10=cpm02(+) "
             StrSQLa = "select decode('" & GetPrjNation1(oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04) & "','000',cpm03,cpm04),a1903,axf04 from acc190,acc151,caseprogress,casepropertymap where a1908='" & tmpA1p04 & "' and a1902=axf01(+) and axf03='" & oCP01 & oCP02 & oCP03 & oCP04 & "' and axf02=cp09(+) and cp01=cpm01(+) and cp10=cpm02(+) "
             tRS2.CursorLocation = adUseClient
             tRS2.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
             If tRS2.RecordCount <> 0 Then
                Do While Not tRS2.EOF
                   If CheckStr(tRS2.Fields(0)) <> "" Then
                      TmpR030006 = TmpR030006 & CheckStr(tRS2.Fields(0)) & " " & CheckStr(tRS2.Fields(1)) & " " & Format(CheckStr(tRS2.Fields(2)), "0.00") & " ;"
                   End If
                   tRS2.MoveNext
                Loop
             End If
         Else
             '2010/12/29 MODIFY BY SONIA 依申請國家抓案件性質
             'StrSQLa = "select cpm03,a1505,axf04 from acc150,acc151,caseprogress,casepropertymap where a1512='" & tmpA1p04 & "' and a1501=axf01(+) and axf03='" & oCP01 & oCP02 & oCP03 & oCP04 & "' and axf02=cp09(+) and cp01=cpm01(+) and cp10=cpm02(+) "
             StrSQLa = "select decode('" & GetPrjNation1(oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04) & "','000',cpm03,cpm04),a1505,axf04 from acc150,acc151,caseprogress,casepropertymap where a1512='" & tmpA1p04 & "' and a1501=axf01(+) and axf03='" & oCP01 & oCP02 & oCP03 & oCP04 & "' and axf02=cp09(+) and cp01=cpm01(+) and cp10=cpm02(+) "
             tRS2.CursorLocation = adUseClient
             tRS2.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
             If tRS2.RecordCount <> 0 Then
                Do While Not tRS2.EOF
                   If CheckStr(tRS2.Fields(0)) <> "" Then
                      TmpR030006 = TmpR030006 & CheckStr(tRS2.Fields(0)) & " " & CheckStr(tRS2.Fields(1)) & " " & Format(CheckStr(tRS2.Fields(2)), "0.00") & " ;"
                   End If
                   tRS2.MoveNext
                Loop
             End If
         End If
         tRS.MoveNext
      Loop
     If TmpR030006 <> "" Then
        'modify by sonia 2016/5/19 R105050157會出現插入的欄位值過大
        'cnnConnection.Execute "update r040320_t set r030006='" & ChgSQL(TmpR030006) & "' where id='" & strUserNum & "'     "
        cnnConnection.Execute "update r040320_t set r030006='" & ChgSQL(Left(TmpR030006, 800)) & "' where id='" & strUserNum & "'     "
     End If
   End If

   'add by nickc 2005/06/06 抓智權人員
   StrSQLa = "SELECT MAX(to_char(CP05)||CP09) FROM CASEPROGRESS WHERE CP01='" & oCP01 & "' AND CP02='" & oCP02 & "' AND CP03='" & IIf(Len(oCP03) = 0, "0", oCP03) & "' AND CP04='" & IIf(Len(oCP04) = 0, "00", oCP04) & "' AND CP09<'C' "
   Set tRS = New ADODB.Recordset
   tRS.CursorLocation = adUseClient
   tRS.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If tRS.RecordCount <> 0 Then
      strSQL1 = Right(CheckStr(tRS.Fields(0)), 9)
   End If
   '**************** 將業務區改成抓案件進度檔   91.08.15  nick
   strSql = "SELECT cp13,NVL(NVL(A0902,A0903),cp12),NVL(ST02,CP13),ST04 FROM CASEPROGRESS,STAFF,ACC090 WHERE CP13=ST01(+) AND cp12=A0901(+) AND CP09='" & strSQL1 & "' "
   Set tRS = New ADODB.Recordset
   tRS.CursorLocation = adUseClient
   tRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If tRS.RecordCount <> 0 Then
      m_CP13St01 = CheckStr(tRS.Fields(0))
      m_CP13St02 = CheckStr(tRS.Fields(2))
   Else
      m_CP13St01 = ""
      m_CP13St02 = ""
   End If
      
'cnnConnection.BeginTrans   '2011/3/14 移到上面

   CheckOC
   '結餘存檔
   For i = 2 To 20
      strTemp(i) = ""
   Next i
   '取得編號
   'edit by nickc 2006/03/07
   strSql = "SELECT R030001,R030002,R030003,R030004,R030005,R030006,R030007,R030008,R030009,R030010,R030011,R030016,R030012,R030013,R030014,R030015,r030016,r030017,r030018 FROM R040320_t WHERE ID='" & strUserNum & "' and rownum <2 "
   Set rs940629 = New ADODB.Recordset
   With rs940629
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         strTemp(1) = AccAutoNo("R", 4, Mid(strSrvDate(1), 1, 4) - 1911, Mid(strSrvDate(1), 5, 2))
         '更新資料庫
         '2012/7/12 modify by sonia 加可結餘刪除日期,cp146 is null條件
         If oType = "CP" Then
            '只更新 CP 依照結餘當天，所有未上結餘的，都要上
            'edit by nickc 2005/07/27
            'cnnConnection.Execute "update caseprogress set cp59='*' where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp59 is null " ' and cp109=(select min(cp109) from caseprogress where cp59 is null and cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp109 is not null )  "
            '2011/6/1 modify by sonia
            'cnnConnection.Execute "update caseprogress set cp59='" & strTemp(1) & "' where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp59 is null " ' and cp109=(select min(cp109) from caseprogress where cp59 is null and cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp109 is not null )  "
            '2011/11/7 modify by sonia 加cp27 is not null條件TF-000350
            '2011/11/9 MODIFY BY SONIA 所有CP59 IS NULL & CP109 IS NOT NULL 的進度都要更新結餘單號CFT-007977(R100060132)
            'cnnConnection.Execute "update caseprogress set cp59='" & strTemp(1) & "' where cp01='" & oCP01 & "' " & strSQLCP & " and cp59 is null and cp109=(select min(cp109) from caseprogress where cp59 is null and cp01='" & oCP01 & "' " & strSQLCP & " and cp109 is not null and cp27 is not null) "
            'modify by sonia 2016/2/24 取消and cp27 is not null條件 CFP-013827(若未發文者不更新,則下次再跑時又會抓到此收文號則每月會作廢前次又重新產生)
            'modify by sonia 2024/9/30 取消and cp146 is null條件
            'cnnConnection.Execute "update caseprogress set cp59='" & strTemp(1) & "' where cp01='" & oCP01 & "' " & strSQLCP & " and cp59 is null and cp146 is null and cp109 is not null "
            cnnConnection.Execute "update caseprogress set cp59='" & strTemp(1) & "' where cp01='" & oCP01 & "' " & strSQLCP & " and cp59 is null and cp109 is not null "
         Else
            '從 np 來的全部都更新，不管法定日期和可結餘日期
            'edit by nickc 2005/07/27
            'cnnConnection.Execute "update caseprogress set cp59='*' where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp59 is null "
            '2011/6/1 modify by sonia
            'cnnConnection.Execute "update caseprogress set cp59='" & strTemp(1) & "' where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp59 is null "
            '2011/11/7 modify by sonia 加cp27 is not null條件TF-000350
            'modify by sonia 2016/2/24 取消and cp27 is not null條件 CFP-013827(若未發文者不更新,則下次再跑時又會抓到此收文號則每月會作廢前次又重新產生)
            'modify by sonia 2024/9/30 取消and cp146 is null條件
            'cnnConnection.Execute "update caseprogress set cp59='" & strTemp(1) & "' where cp01='" & oCP01 & "' " & strSQLCP & " and cp59 is null and cp146 is null "
            cnnConnection.Execute "update caseprogress set cp59='" & strTemp(1) & "' where cp01='" & oCP01 & "' " & strSQLCP & " and cp59 is null "
'2010/3/23 CANCEL BY SONIA
'            Select Case oCP01
'            Case "P", "FCP", "CFP", "FG", "CPS", "PS"
'                     TmpRule = " and np07<>1204 and np07 <>1503 and np07<>1601 and np07<>411 and np07<>997 and np07<>998 and np07<>999 "
'            Case "T", "FCT", "CFT", "CFC", "S", "TB", "TC", "TD", "TE", "TF", "TM", "TR", "TS", "TT"
'                     TmpRule = " and np07<>1403 and np07 <>305 and np07<>997 and np07<>998 "
'            Case Else
'            End Select
'2010/3/23 END
            '2008/7/8 MODIFY BY SONIA 未到期者不可更新 CFP-015531
            'cnnConnection.Execute "update nextprogress set np06='N',np11=to_number(to_char(sysdate,'YYYYMMDD')),np12='79' where np02='" & oCP01 & "' and np03='" & oCP02 & "' and np04='" & oCP03 & "' and np05='" & oCP04 & "' and np06 is null " & TmpRule
            '2010/3/23 MODIFY BY SONIA 剔除專業部控管的下一程序改以strNpSqlOfNoSalesDuty控制
            'cnnConnection.Execute "update nextprogress set np06='N',np11=to_number(to_char(sysdate,'YYYYMMDD')),np12='79' where np02='" & oCP01 & "' and np03='" & oCP02 & "' and np04='" & oCP03 & "' and np05='" & oCP04 & "' and np06 is null AND np09<=to_number(to_char(sysdate,'YYYYMMDD')) " & TmpRule
            '2011/6/1 modify by sonia
            'cnnConnection.Execute "update nextprogress set np06='N',np11=to_number(to_char(sysdate,'YYYYMMDD')),np12='79' where np02='" & oCP01 & "' and np03='" & oCP02 & "' and np04='" & oCP03 & "' and np05='" & oCP04 & "' and np06 is null AND np09<=to_number(to_char(sysdate,'YYYYMMDD')) " & strNpSqlOfNoSalesDuty
            'modify by sonia 2025/5/13 改更新新欄位NP25，並非strNpSqlOfNoSalesDuty而是挑案件性質故取消strNpSqlOfNoSalesDuty改用NPsql
            'cnnConnection.Execute "update nextprogress set np06='N',np11=to_number(to_char(sysdate,'YYYYMMDD')),np12='79' where np02='" & oCP01 & "' " & strSQLnp & " and np06 is null AND np09<=to_number(to_char(sysdate,'YYYYMMDD')) " & strNpSqlOfNoSalesDuty
            cnnConnection.Execute "update nextprogress set np25=to_number(to_char(sysdate,'YYYYMMDD')) where np02='" & oCP01 & "' " & strSQLnp & " and np06 is null and nvl(np25,0)=0 AND np09<=to_number(to_char(sysdate,'YYYYMMDD')) " & NPsql
         End If
         strSql = AccSaveAutoNo("R", Right(strTemp(1), 4), Mid(strSrvDate(1), 1, 4) - 1911, Mid(strSrvDate(1), 5, 2))
         
         strTemp(0) = strUserName
         '2011/2/24 modify by sonia
         'strTemp(2) = Format(ChangeTStringToTDateString(strSrvDate(2)), "YY")
         strTemp(2) = strSrvDate(2) \ 10000
         '2011/2/24 end
         strTemp(3) = Format(ChangeTStringToTDateString(strSrvDate(2)), "mm")
         strTemp(4) = Format(ChangeTStringToTDateString(strSrvDate(2)), "dd")
         strTemp(5) = IIf(CheckStr(.Fields(0)) = "---", oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04, CheckStr(.Fields(0)))
         strTemp(6) = GetCustomerName(GetPrjPeopleNum1(oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04))     'CheckStr(.Fields(2))   'edit by nickc 2006/04/12 修改
         strTemp(7) = m_CP13St01 'CheckStr(.Fields(4))
         strTemp(8) = GetPrjNation(oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04) ' CheckStr(.Fields(16))
         'Modify by Morgan 2010/6/21代理人名稱前面加編號
         'strTemp(9) = GetFAgentName(GetPrjFagentNumByCPNot001(oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04)) 'CheckStr(.Fields(12))
         strTemp(9) = GetPrjFagentNumByCPNot001(oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04)
         strTemp(9) = Trim(strTemp(9) & " " & GetFAgentName(strTemp(9)))
         'end 2010/6/21
         strTemp(10) = CheckStr(.Fields(5))
      Else
         cnnConnection.RollbackTrans
         Exit Function
      End If
      CheckOC
      
      cnnConnection.Execute "insert into Acc240 (A240001,A240002,A240003,A240004,A240005,A240006,A240007,A240008,A240009,A240010,A240011,A240012,A240013,A240014,A240015) values (" & strTemp(2) & strTemp(3) & strTemp(4) & ",'" & strTemp(1) & "',null,'" & strUserNum & "','" & oCP01 & "','" & oCP02 & "','" & oCP03 & "','" & oCP04 & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "'," & IIf(IsOldSystem = True, "'Y'", "null") & ",null) "
      '2010/11/26 MODIFY BY SONIA 語法錯誤,少ACC020
      'StrSQLa = "select sum(decode(substr(ax205,1,4),'2201',decode(ax206,0,0,nvl(ax206,0)))) as A3,ax202 from acc021 where 1=1 " & StrSQL3 & " group by ax202 "
      StrSQLa = "select sum(decode(substr(ax205,1,4),'2201',decode(ax206,0,0,nvl(ax206,0)))) as A3,ax202 from acc021,ACC020 where 1=1 AND ax201=a0201(+) and ax202=a0202(+) AND INSTR(AX212,'結餘')=0 " & StrSQL3 & " group by ax202 "
   
      strSql = "select sum(r030008),sum(r030009),r030017 from (SELECT distinct r030018,R030008,R030009,r030017 FROM R040320_t WHERE ID='" & strUserNum & "' and r030008>0) A group by r030017 order by r030017"
   
      If .State = 1 Then .Close
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         .MoveFirst
         Do While Not .EOF
            strTemp(13) = CheckStr(.Fields(2))
            strSql = "select sum(newa1),sum(newa2) from " & NewAcc020021 & ",(" & StrSQLa & ") B where NewTable3.ax202=B.ax202 and NewTable3.ax202='" & strTemp(13) & "' "
            CheckOC3
            AdoRecordSet3.CursorLocation = adUseClient
            AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If AdoRecordSet3.RecordCount <> 0 Then
                AdoRecordSet3.MoveFirst
                strTemp(11) = CheckStr(AdoRecordSet3.Fields(0))
                strTemp(12) = CheckStr(AdoRecordSet3.Fields(1))
            Else
                strTemp(11) = "0"
                strTemp(12) = "0"
            End If
            If Val(strTemp(11)) + Val(strTemp(12)) > 0 Then
                'add by nickc 2006/03/07 重新抓按件性質
                strSql = "select distinct cpm04 from " & NewAcc020021 & ",acc1p0,acc1u0,caseprogress,casepropertymap where ax201=a1p01(+) and ax202=a1p22(+) and 'A'=a1p02(+) and a1p04=a1u01(+) and a1u03=cp09(+) and a1p17='" & oCP01 & oCP02 & oCP03 & oCP04 & "' and cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and a1p06<>'TOT' and ax202='" & strTemp(13) & "' and cp01=cpm01(+) and cp10=cpm02(+)  "
                CheckOC3
                strTemp(13) = ""
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If AdoRecordSet3.RecordCount <> 0 Then
                    AdoRecordSet3.MoveFirst
                    Do While Not AdoRecordSet3.EOF
                        If CheckStr(AdoRecordSet3.Fields(0)) <> "" Then
                            strTemp(13) = strTemp(13) & CheckStr(AdoRecordSet3.Fields(0)) & ";"
                        End If
                        AdoRecordSet3.MoveNext
                    Loop
                End If
                cnnConnection.Execute "insert into acc241 select '" & strTemp(1) & "',nvl(max(A241002),0) +1,'" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "',null,null from Acc241 where A241001='" & strTemp(1) & "'"
            End If
            .MoveNext
         Loop
      End If
      '2010/11/26 MODIFY BY SONIA 語法錯誤,少ACC020
      'StrSQLa = "select sum(decode(substr(ax205,1,4),'2201',decode(ax206,0,0,nvl(ax206,0)))) as A3,ax202 from acc021 where 1=1 " & StrSQL3 & " group by ax202 "
      StrSQLa = "select sum(decode(substr(ax205,1,4),'2201',decode(ax206,0,0,nvl(ax206,0)))) as A3,ax202 from acc021,ACC020 where 1=1 AND ax201=a0201(+) and ax202=a0202(+) AND INSTR(AX212,'結餘')=0 " & StrSQL3 & " group by ax202 "
   
      '合計
      CheckOC
      strSql = "select SUM(newa1),sum(newa2),sum(newa3),sum(e)/count(*),sum(newa1)-sum(newa2)-sum(newa3)-(sum(e)/count(*))+sum(f),sum(f) from (select sum(r030008) as b,sum(r030009) as c,sum(r030013) as d,sum(r030014)/count(*) as e,sum(r030019) as f,r030017 from (SELECT distinct R030008,R030009,r030014,r030013,r030019,r030017 FROM R040320_t WHERE ID='" & strUserNum & "' ) A group by r030017 ) C,(" & StrSQLa & ") B," & NewAcc020021 & " where C.R030017=B.ax202 and C.R030017=NewTable3.ax202 "
      If .State = 1 Then .Close
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         strTemp(11) = CheckStr(.Fields(0))
         strTemp(12) = CheckStr(.Fields(1))
         '2012/7/9 modify by sonia 加format (R100120568)
         strTemp(13) = Format(CheckStr(.Fields(2)), FAmount)
         strTemp(14) = Format(CheckStr(.Fields(3)), FAmount)
         strTemp(16) = CheckStr(.Fields(5))
   
         '2011/10/27 modify by sonia 退費要加回來CFT-013515
         'If Val(strTemp(11)) - Val(strTemp(12)) - Val(strTemp(13)) < Val(strTemp(14)) Then
         '      strTemp(14) = Trim(Val(strTemp(11)) - Val(strTemp(12)) - Val(strTemp(13)))
         If Val(strTemp(11)) - Val(strTemp(12)) - Val(strTemp(13)) + Val(strTemp(16)) < Val(strTemp(14)) Then
            '2012/7/9 modify by sonia 加format (R100120568)
            strTemp(14) = Format(Trim(Val(strTemp(11)) - Val(strTemp(12)) - Val(strTemp(13))) + Val(strTemp(16)), FAmount)
         '2011/10/27 END
               strTemp(15) = 0
         Else
               strTemp(15) = CheckStr(.Fields(4))
         End If
      End If
      cnnConnection.Execute "   insert into acc241 select '" & strTemp(1) & "',998,'" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "' from dual "
      If Val(strTemp(16)) <> 0 Then
         cnnConnection.Execute "insert into acc241 select '" & strTemp(1) & "',999,null                 ,null                 ,'" & strTemp(16) & "',null,null from dual "
      End If
      '2012/7/27 add by sonia
      cnnConnection.Execute "delete r040320_t1 where id='" & strUserNum & "' and r030001='" & oCP01 & "' and r030002='" & oCP02 & "' and r030003='" & oCP03 & "'"
      '2012/7/27 end
      'add by sonia 2016/6/6 新增非出名公司傳票資料至ACC242,以利印出結餘單提醒財務注意
      'modify by sonia 2020/4/16 +L公司
      'cnnConnection.Execute "   insert into acc242 select '" & strTemp(1) & "',ax201,ax202,ax203 from acc020, acc021 where a0201=decode('" & strCompNo & "','J','1','J') and  ax201=a0201(+) and ax202=a0202(+) AND INSTR(AX212,'結餘')=0 " & StrSQL3
      cnnConnection.Execute "   insert into acc242 select '" & strTemp(1) & "',ax201,ax202,ax203 from acc020, acc021 where a0201<>'" & strCompNo & "' and ax201=a0201(+) and ax202=a0202(+) AND INSTR(AX212,'結餘')=0 " & StrSQL3
      
   End With
   CheckOC
   cnnConnection.CommitTrans
   If IsOldSystem = True Then
      PrintData2_SpaceOld strTemp(1)
   Else
      PrintData2_1Old strTemp(1)
   End If
   If IsPrintok = False Then
      IsPrintok = True
   End If
   'Screen.MousePointer = vbDefault
   Exit Function
RollBackData:
   cnnConnection.RollbackTrans
   'edit by nickc 2006/02/27
   MsgBox Err.Description, vbCritical, "錯誤語法發生！"
   PrintByOne = "6"
   Resume Next
End Function

'整批列印
Sub PrintByAll()
Dim StrSQLa As String
Dim rs5 As New ADODB.Recordset
Dim rsQuery As ADODB.Recordset  'add by sonia 2020/3/5
   
   cnnConnection.BeginTrans  '2012/7/11 add by sonia
   'add by sonia 2024/12/6 刪除工作檔中已離職人員的資料
   cnnConnection.Execute "Delete R040320_T Where Id In (Select Distinct Id From R040320_T,Staff Where Id=St01(+) And St04<>'1')"
   cnnConnection.Execute "Delete R040330 Where Id In (Select Distinct Id From R040330,Staff Where Id=St01(+) And St04<>'1')"
   'end 2024/12/6
   '先塞暫存 若 cp  有，就做 cp  ，若沒有 cp 則不管可結餘日期一律結
   cnnConnection.Execute "DELETE FROM R040320_T1 WHERE ID='" & strUserNum & "' "
   cnnConnection.Execute "DELETE FROM r040330 WHERE ID='" & strUserNum & "' "  'add by sonia 2025/4/18
   
   'add by sonia 2020/2/14 +本所案號條件
   If Option1(2).Value = True Then
      If txt1(1) = "TF" Then
         'modify by sonia 2024/9/30 cp146 is null改為(cp146 is null or cp109>cp146)
         'modify by sonia 2025/7/8 馬德里案已於2010/3/26母案與延土延伸分開算，但此處沒改到
         'StrSQLa = "select cp01,cp02,cp03,cp04,'CP','" & strUserNum & "',max(cp109) cp109" & _
            " from caseprogress where cp59 is null and (cp146 is null or cp109>cp146) and (((nvl(cp16,0)-nvl(cp77,0)>0) and substr(cp60,1,1)='E') or substr(cp60,1,1)='X' or cp61 is not null) " & _
            " and cp01='" & txt1(1) & "' and cp02 like '" & Left(txt1(2), 5) & "%' " & _
            " and cp109+0<=to_number(to_char(add_months(sysdate,-3),'YYYYMMDD'))" & _
            " and (cp01||''     in ('T','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT')" & _
            " or  (cp01||'' not in ('T','FCT','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT')" & _
            " and cp109+0<=to_number(to_char(add_months(sysdate,-6),'YYYYMMDD')))) group by cp01,cp02,cp03,cp04"
         StrSQLa = "select cp01,cp02,cp03,cp04,'CP','" & strUserNum & "',max(cp109) cp109" & _
            " from caseprogress where cp59 is null and (cp146 is null or cp109>cp146) and (((nvl(cp16,0)-nvl(cp77,0)>0) and substr(cp60,1,1)='E') or substr(cp60,1,1)='X' or cp61 is not null) " & _
            " and cp01='" & txt1(1) & "' and cp02='" & txt1(2) & "' " & _
            " and cp109+0<=to_number(to_char(add_months(sysdate,-3),'YYYYMMDD'))" & _
            " and (cp01||''     in ('T','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT')" & _
            " or  (cp01||'' not in ('T','FCT','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT')" & _
            " and cp109+0<=to_number(to_char(add_months(sysdate,-6),'YYYYMMDD')))) group by cp01,cp02,cp03,cp04"
      ElseIf txt1(1) = "CFP" Then
         'modify by sonia 2024/9/30 cp146 is null改為(cp146 is null or cp109>cp146)
         StrSQLa = "select cp01,cp02,cp03,cp04,'CP','" & strUserNum & "',max(cp109) cp109" & _
            " from caseprogress where cp59 is null and (cp146 is null or cp109>cp146) and (((nvl(cp16,0)-nvl(cp77,0)>0) and substr(cp60,1,1)='E') or substr(cp60,1,1)='X' or cp61 is not null) " & _
            " and cp01='" & txt1(1) & "' and cp02='" & txt1(2) & "' and cp03='" & IIf(Trim(txt1(3)) = "", "0", txt1(3)) & "' " & _
            " and cp109+0<=to_number(to_char(add_months(sysdate,-3),'YYYYMMDD'))" & _
            " and (cp01||''     in ('T','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT')" & _
            " or  (cp01||'' not in ('T','FCT','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT')" & _
            " and cp109+0<=to_number(to_char(add_months(sysdate,-6),'YYYYMMDD')))) group by cp01,cp02,cp03,cp04"
      Else
         'modify by sonia 2024/9/30 cp146 is null改為(cp146 is null or cp109>cp146)
         StrSQLa = "select cp01,cp02,cp03,cp04,'CP','" & strUserNum & "',max(cp109) cp109" & _
            " from caseprogress where cp59 is null and (cp146 is null or cp109>cp146) and (((nvl(cp16,0)-nvl(cp77,0)>0) and substr(cp60,1,1)='E') or substr(cp60,1,1)='X' or cp61 is not null) " & _
            " and cp01='" & txt1(1) & "' and cp02='" & txt1(2) & "' and cp03='" & IIf(Trim(txt1(3)) = "", "0", txt1(3)) & "' and cp04='" & IIf(Trim(txt1(4)) = "", "00", txt1(4)) & "' " & _
            " and cp109+0<=to_number(to_char(add_months(sysdate,-3),'YYYYMMDD'))" & _
            " and (cp01||''     in ('T','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT')" & _
            " or  (cp01||'' not in ('T','FCT','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT')" & _
            " and cp109+0<=to_number(to_char(add_months(sysdate,-6),'YYYYMMDD')))) group by cp01,cp02,cp03,cp04"
      End If
   Else
   'end 2020/2/14
      'Modify by Morgan 2009/10/23 改寫法
      'StrSQLa = "select distinct cp01,cp02,cp03,cp04,'CP','" & strUserNum & "' from caseprogress where cp59 is null and (((not (cp01||''='T' or cp01||''='FCT' or cp01||''='CFT' or cp01||''='CFC' or cp01||''='S' or cp01||''='TB' or cp01||''='TC' or cp01||''='TD' or cp01||''='TF' or cp01||''='TM' or cp01||''='TR' or cp01||''='TS' or cp01||''='TT')) " & _
                        " and cp109<=to_number(to_char(add_months(sysdate,-6),'YYYYMMDD'))) or (cp01||''='T' or cp01||''='FCT' or cp01||''='CFT' or cp01||''='CFC' or cp01||''='S' or cp01||''='TB' or cp01||''='TC' or cp01||''='TD' or cp01||''='TF' or cp01||''='TM' or cp01||''='TR' or cp01||''='TS' or cp01||''='TT') " & _
                        " and cp109<=to_number(to_char(add_months(sysdate,-3),'YYYYMMDD'))) "
      '2010/11/23 modify by sonia 加入nvl(cp16,0)-nvl(cp77,0)>0條件
      '2012/7/9 modify by sonia 改為(nvl(cp16,0)-nvl(cp77,0)>0 or cp60 is not null) P-091307
      '2012/7/9 modify by sonia 婧瑄說有收據或請款單或帳單都要算結餘
      'StrSQLa = "select cp01,cp02,cp03,cp04,'CP','" & strUserNum & "',max(cp109) cp109" & _
         " from caseprogress where cp59 is null and (nvl(cp16,0)-nvl(cp77,0)>0 or cp60 is not null) " & _
         " and cp109+0<=to_number(to_char(add_months(sysdate,-3),'YYYYMMDD'))" & _
         " and (cp01||''     in ('T','FCT','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT')" & _
         " or  (cp01||'' not in ('T','FCT','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT')" & _
         " and cp109+0<=to_number(to_char(add_months(sysdate,-6),'YYYYMMDD')))) group by cp01,cp02,cp03,cp04"
      '2012/7/12 modify by sonia 加可結餘刪除日期,cp146 is null條件
      'modify by sonia 2024/9/30 cp146 is null改為(cp146 is null or cp109>cp146)
      StrSQLa = "select cp01,cp02,cp03,cp04,'CP','" & strUserNum & "',max(cp109) cp109" & _
         " from caseprogress where cp59 is null and (cp146 is null or cp109>cp146) and (((nvl(cp16,0)-nvl(cp77,0)>0) and substr(cp60,1,1)='E') or substr(cp60,1,1)='X' or cp61 is not null) " & _
         " and cp109+0<=to_number(to_char(add_months(sysdate,-3),'YYYYMMDD'))" & _
         " and (cp01||''     in ('T','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT')" & _
         " or  (cp01||'' not in ('T','FCT','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT')" & _
         " and cp109+0<=to_number(to_char(add_months(sysdate,-6),'YYYYMMDD')))) group by cp01,cp02,cp03,cp04"
                        
      'modify by sonia 2021/11/15 商標下一程序剔除專業部管制案件性質1101,1403,1701,1711,201,203,303,306,310,312,313,612,706,709,710,723,994及TM之1602
      'modify by sonia 2025/5/13 改用新欄位NP25，2024/9 先取消下一程序的(杜協理說不可更新NP06,NP11,NP12)，並把共用NP條件語法抽出來，後面更新NP25可共用(也取消FCP,FG,另NP取消商標之303延期、1101通知申請案號及1711通知使用宣誓)
'      StrSQLa = StrSQLa & " union select np02 as cp01,np03 as cp02,np04 as cp03,np05 as cp04,'NP','" & strUserNum & "',np09" & _
'         " from (select np02,np03,np04,np05,max(np09) as np09 from nextprogress Where np06 Is Null" & _
'         " and ((np02='CFP' and np03>='014505') or (np02='P' and np03>='067711') or (np02='CFT' and np03>='008899')" & _
'         " or (np02='CFC' and np03>='000683') or (np02='T' and np03>='132277') or (np02='TC' and np03>='010025')" & _
'         " or (np02='TD' and np03>='000130') or (np02='TR' and np03>='000064') or (np02='TS' and np03>='000026')" & _
'         " or (np02='CPS' and np03>='000001') or (np02='PS' and np03>='000014') or (np02='S' and np03>='000001')" & _
'         " or (np02='CFL' and np03>='010408') or (np02='TF' and np03>='000450') or (np02='TB' and np03>='000115')" & _
'         " or (np02='TM' and np03>='000025') or (np02='TT' and np03>='000007'))" & _
'         " and ((np02||'' in ('FCP','FG','PS') and np07+0 not in (1204,1503,1601,411,997,998,999))" & _
'         " or (np02||'' in ('T','FCT','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT') and np07+0 not in(1403,305,997,998,1101,1403,1701,1602))" & _
'         " or (np02||'' in ('P','CFP','CPS') and np07 in (107,111,119,204,205,208,416,427,501,502,503,601,605,606,607,804))" & _
'         " ) group by np02,np03,np04,np05) NewNP" & _
'         " where ( np02||'' in ('P','FCP','CFP','FG','CPS','PS') and np09<=to_number(to_char(add_months(sysdate,-6),'YYYYMMDD')))" & _
'         " or (np02||'' in ('T','FCT','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT')" & _
'         " and np09<=to_number(to_char(add_months(sysdate,-3),'YYYYMMDD')))"
      NPsql = " and ((np02||'' in ('P','PS','CFP','CPS') and np07 in (107,111,119,204,205,208,416,427,501,502,503,601,605,606,607,804)) " & _
              " or (np02||'' in ('T','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT') and np07+0 not in (1403,305,997,998,1101,1701,1602,1711,201,203,303,306,310,312,313,612,706,709,710,723,994)) )"
      StrSQLa = StrSQLa & " union select np02 as cp01,np03 as cp02,np04 as cp03,np05 as cp04,'NP','" & strUserNum & "',np09" & _
         " from (select np02,np03,np04,np05,max(np09) as np09 from nextprogress Where nvl(np25,0)=0 and np06 Is Null" & _
         " and ((np02='CFP' and np03>='014505') or (np02='P' and np03>='067711') or (np02='CFT' and np03>='008899')" & _
         " or (np02='CFC' and np03>='000683') or (np02='T' and np03>='132277') or (np02='TC' and np03>='010025')" & _
         " or (np02='TD' and np03>='000130') or (np02='TR' and np03>='000064') or (np02='TS' and np03>='000026')" & _
         " or (np02='CPS' and np03>='000001') or (np02='PS' and np03>='000014') or (np02='S' and np03>='000001')" & _
         " or (np02='CFL' and np03>='010408') or (np02='TF' and np03>='000450') or (np02='TB' and np03>='000115')" & _
         " or (np02='TM' and np03>='000025') or (np02='TT' and np03>='000007'))" & NPsql & " group by np02,np03,np04,np05) NewNP" & _
         " where ( np02||'' in ('P','CFP','CPS','PS') and np09<=to_number(to_char(add_months(sysdate,-6),'YYYYMMDD')))" & _
         " or (np02||'' in ('T','CFT','CFC','S','TB','TC','TD','TF','TM','TR','TS','TT')" & _
         " and np09<=to_number(to_char(add_months(sysdate,-3),'YYYYMMDD')))"
   End If   'add by sonia 2020/2/14
   
   cnnConnection.Execute "insert into r040320_t1(R030001,R030002,R030003,R030004,R030005,id,R030006) (" & StrSQLa & ") ", intI
   
   '剔除重複的，並以 cp 為主
   'add by sonia 2025/5/15 更新NP25=19221111以免下次又抓到浪費時間
   cnnConnection.Execute "update nextprogress set np25=19221111 where (np02,np03,np04,np05,np22) in (select np02,np03,np04,np05,np22 from r040320_t1 t1,r040320_t1 t2,nextprogress " & _
                         "where t1.id='" & strUserNum & "' and t1.R030005='CP' and t2.id='" & strUserNum & "' and t2.R030005='NP' and t1.R030001=t2.R030001(+) and t1.R030002=t2.R030002(+) and t1.R030003=t2.R030003(+) and t1.R030004=t2.R030004(+) and t2.r030001 is not null " & _
                         "and t1.R030001=np02(+) and t1.R030002=np03(+) and t1.R030003=np04(+) and t1.R030004=np05(+) and np01 is not null and np06 Is Null and NVL(NP25,0)=0 " & NPsql & ")", intI
   'end 2025/5/15
   'Modify by Morgan 2009/10/23
   'cnnConnection.Execute "delete from r040320_t1 where id='" & strUserNum & "' and r030001||r030002||r030003||r030004||r030005 in (select distinct r030001||r030002||r030003||r030004||'NP' from r040320_T1 where id='" & strUserNum & "' group by r030001||r030002||r030003||r030004 having count(*) >=2 )"
   cnnConnection.Execute "delete from r040320_t1 a where ID='" & strUserNum & "' and R030005='NP' and exists(select * from r040320_t1 b where b.ID='" & strUserNum & "' and b.R030001=a.R030001 and b.R030002=a.R030002 and b.R030003=a.R030003 and b.R030004=a.R030004 and b.R030005='CP')", intI
   '刪除CFP的EPC案且可結餘日期為1年以內的
   cnnConnection.Execute "delete from r040320_t1 a where ID='" & strUserNum & "' and R030006>to_number(to_char(add_months(sysdate,-12),'YYYYMMDD')) and exists( select * from patent where pa01=R030001 and pa02=R030002 and pa03=R030003 and pa04=R030004 and pa09='221')", intI
   
   '剔除台灣案
   cnnConnection.Execute "delete from r040320_t1 where id='" & strUserNum & "' and r030001||r030002||r030003||r030004 in (select distinct pa01||pa02||pa03||pa04 from patent,r040320_t1 where pa09='000' and r030001=pa01(+) and r030002=pa02(+) and r030003=pa03(+) and r030004=pa04(+) )"
   cnnConnection.Execute "delete from r040320_t1 where id='" & strUserNum & "' and r030001||r030002||r030003||r030004 in (select distinct tm01||tm02||tm03||tm04 from trademark,r040320_t1 where tm10='000' and r030001=tm01(+) and r030002=tm02(+) and r030003=tm03(+) and r030004=tm04(+))"
   cnnConnection.Execute "delete from r040320_t1 where id='" & strUserNum & "' and r030001||r030002||r030003||r030004 in (select distinct sp01||sp02||sp03||sp04 from servicepractice,r040320_t1 where sp09='000' and r030001=sp01(+) and r030002=sp02(+) and r030003=sp03(+) and r030004=sp04(+))"
   cnnConnection.Execute "delete from r040320_t1 where id='" & strUserNum & "' and r030001||r030002||r030003||r030004 in (select distinct lc01||lc02||lc03||lc04 from lawcase,r040320_t1 where lc15='000' and r030001=lc01(+) and r030002=lc02(+) and r030003=lc03(+) and r030004=lc04(+))"
   '剔除閉卷   僅限 由 np 來
   'add by sonia 2025/6/6 更新NP25=19221111以免下次又抓到浪費時間
   cnnConnection.Execute "update nextprogress set np25=19221111 where (np02,np03,np04,np05,np22) in (select np02,np03,np04,np05,np22 from R040320_T1,NEXTPROGRESS,PATENT,TRADEMARK,SERVICEPRACTICE,LAWCASE " & _
                         "where id='" & strUserNum & "' and R030005='NP' and R030001=PA01(+) and R030002=PA02(+) and R030003=PA03(+) and R030004=PA04(+) and R030001=TM01(+) and R030002=TM02(+) and R030003=TM03(+) and R030004=TM04(+) " & _
                         "and R030001=SP01(+) and R030002=SP02(+) and R030003=SP03(+) and R030004=SP04(+) and R030001=lc01(+) and R030002=lc02(+) and R030003=lc03(+) and R030004=lc04(+) and pa57||tm29||sp15||lc08 is not null " & _
                         "and R030001=np02(+) and R030002=np03(+) and R030003=np04(+) and R030004=np05(+) and np01 is not null and np06 Is Null and NVL(NP25,0)=0 " & NPsql & ")", intI
   'end 2025/6/6
   cnnConnection.Execute "delete from r040320_t1 where id='" & strUserNum & "' and r030001||r030002||r030003||r030004||r030005 in (select distinct pa01||pa02||pa03||pa04||'NP' from patent,r040320_t1 where pa57='Y' and r030001=pa01(+) and r030002=pa02(+) and r030003=pa03(+) and r030004=pa04(+) and r030005='NP' )"
   cnnConnection.Execute "delete from r040320_t1 where id='" & strUserNum & "' and r030001||r030002||r030003||r030004||r030005 in (select distinct tm01||tm02||tm03||tm04||'NP' from trademark,r040320_t1 where tm29='Y' and r030001=tm01(+) and r030002=tm02(+) and r030003=tm03(+) and r030004=tm04(+) and r030005='NP')"
   cnnConnection.Execute "delete from r040320_t1 where id='" & strUserNum & "' and r030001||r030002||r030003||r030004||r030005 in (select distinct sp01||sp02||sp03||sp04||'NP' from servicepractice,r040320_t1 where sp15='Y' and r030001=sp01(+) and r030002=sp02(+) and r030003=sp03(+) and r030004=sp04(+) and r030005='NP')"
   cnnConnection.Execute "delete from r040320_t1 where id='" & strUserNum & "' and r030001||r030002||r030003||r030004||r030005 in (select distinct lc01||lc02||lc03||lc04||'NP' from lawcase,r040320_t1 where lc08='Y' and r030001=lc01(+) and r030002=lc02(+) and r030003=lc03(+) and r030004=lc04(+) and r030005='NP')"
   'add by nickc 2005/09/29 刪除 np 的 非 P,CFP,CPS 的有任一筆 cp05>np09 的不結餘
   'add by sonia 2025/6/10 更新NP25=19221111以免下次又抓到浪費時間
   cnnConnection.Execute "update nextprogress set np25=19221111 where (np02,np03,np04,np05,np22) in (select np02,np03,np04,np05,np22 from r040320_t1,nextprogress " & _
                         "where id='" & strUserNum & "' and r030001||r030002||r030003||r030004||r030005 in " & _
                         "(select distinct cp01||cp02||cp03||cp04||'NP' from caseprogress,r040320_t1 where r030001=cp01(+) and r030002=cp02(+) and r030003=cp03(+) and r030004=cp04(+) and r030005='NP' and cp05>r030006 and nvl(cp16,0)>0) " & _
                         "and R030001=np02(+) and R030002=np03(+) and R030003=np04(+) and R030004=np05(+) and np01 is not null and np06 Is Null and NVL(NP25,0)=0 " & NPsql & ")", intI
   'end 2025/6/10
   'modify by sonia 2022/11/4 改cp05>np09為cp05>r030006(CFT-014933)
   'cnnConnection.Execute "delete from r040320_t1 where id='" & strUserNum & "' and r030001||r030002||r030003||r030004||r030005 in (select distinct cp01||cp02||cp03||cp04||'NP' from caseprogress,r040320_t1,nextprogress where r030001=cp01(+) and r030002=cp02(+) and r030003=cp03(+) and r030004=cp04(+) and r030005='NP' and cp01=np02(+) and cp02=np03(+) and cp03=np04(+) and cp04=np05(+)  and cp05>np09 )"
   'modify by sonia 2025/5/29 再加NVL(CP16,0)>0否則C類來函也會被剔除(CFP-028066)
   'modify by sonia 2025/6/10 不用抓nextprogress
   'cnnConnection.Execute "delete from r040320_t1 where id='" & strUserNum & "' and r030001||r030002||r030003||r030004||r030005 in (select distinct cp01||cp02||cp03||cp04||'NP' from caseprogress,r040320_t1,nextprogress where r030001=cp01(+) and r030002=cp02(+) and r030003=cp03(+) and r030004=cp04(+) and r030005='NP' and cp01=np02(+) and cp02=np03(+) and cp03=np04(+) and cp04=np05(+) and cp05>r030006 and nvl(cp16,0)>0 )"
   cnnConnection.Execute "delete from r040320_t1 where id='" & strUserNum & "' and r030001||r030002||r030003||r030004||r030005 in (select distinct cp01||cp02||cp03||cp04||'NP' from caseprogress,r040320_t1 where r030001=cp01(+) and r030002=cp02(+) and r030003=cp03(+) and r030004=cp04(+) and r030005='NP' and cp05>r030006 and nvl(cp16,0)>0 )"
                     
   '2011/6/22 add by sonia 馬德里案母子案同時存在則刪除子案;若只有子案則改為母案案號
   cnnConnection.Execute "delete r040320_t1 where (r030001,r030002,id) in (select r030001,r030002,id from r040320_t1 where r030001='TF' and r030003='0' and id='" & strUserNum & "') and r030003<>'0'"
   cnnConnection.Execute "update r040320_t1 set r030003='0',r030004='00' where r030001='TF' and id='" & strUserNum & "'"
   '2011/6/22 end
   '2011/10/4 add by sonia EPC母子案同時存在則刪除子案;若只有子案則改為母案案號
   cnnConnection.Execute "delete r040320_t1 where (r030001,r030002,r030003,id) in (select r030001,r030002,r030003,id from r040320_t1 where r030001='CFP' and r030004='00' and id='" & strUserNum & "') and r030004<>'00'"
   cnnConnection.Execute "update r040320_t1 set r030004='00' where r030001='CFP' and id='" & strUserNum & "' and r030004<>'00'"
   '2011/10/4 end
   
'cancel by sonia 2022/9/19 婧瑄郵件通知：專利國內部及智權部都同意不再管控
'   'Add by Morgan 2010/6/21
'   '剔除美國發明案之公開費尚未退客戶之案件
'   '2011/10/31 MODIFY BY SONIA 退客戶應加AX206>0條件(CFP-022597,022386)
'   strSql = "delete from r040320_t1 a where ID='" & strUserNum & "' and R030001='CFP'" & _
'      " and exists(select * from patent" & _
'      " Where PA01 = R030001 And pa02 = R030002 And pa03 = R030003 And pa04 = R030004" & _
'      " and pa09='101' and pa08='1' and instr(pa15,'B1')>0 and pa13 is null)" & _
'      " and exists(select * from caseprogress" & _
'      " Where CP01 = R030001 And cp02 = R030002 And cp03 = R030003 And cp04 = R030004" & _
'      " and (cp10='601' or cp10='217') and cp61 is not null and cp27>0" & _
'      " and exists(select * from acc151 where axf02=cp09 and axf16='Y'))" & _
'      " and not exists(select * from acc021,ACC1P0,ACC190,ACC161,CASEPROGRESS" & _
'      " where ax214=R030001||R030002||R030003||R030004 and ax205='220106' and instr(ax212,'退公開費')>0 and ax206>0 " & _
'      " AND A1P22(+)=AX202 AND A1P17(+)=AX214 AND A1908(+)=substr(A1P04, 1, Length(A1P04) - 9)" & _
'      " AND AXG01(+)=A1902 and (AXG04>100 OR AXG01 IS NULL)" & _
'      " AND cp09(+)=axg02 and (cp10 IN ('601','217') OR CP10 IS NULL))"
'   cnnConnection.Execute strSql, intI
'   'end 2010/6/21
'end 2022/9/19

   cnnConnection.CommitTrans  '2012/7/11 ADD BY SONIA 否則後面ROLLBACK則會回復
   
   '2011/10/4 MODIFY BY SONIA加DISTINCT 否則會重覆 CFP-014589-0-09及CFP-014589-0-10
   'StrSQLa = "select r030001 as cp01,r030002 as cp02,r030003 as cp03,r030004 as cp04,r030005 as type from r040320_t1 where id='" & strUserNum & "' order by 1,2,3,4 "
   'modify by sonia 2017/5/10 郭雅娟要求CFP依系統類別+管制人+申請國家排序
   'StrSQLa = "select DISTINCT r030001 as cp01,r030002 as cp02,r030003 as cp03,r030004 as cp04,r030005 as type from r040320_t1 where id='" & strUserNum & "' order by 1,2,3,4 "
   'modify by sonia 2020/3/5 因CFP管制人109/4/1以後改業務區劃分故加CASENO
   'modify by sonia 2020/9/10 調整欄位順序(PID前移,放後面會跟其他程式的大欄位共用導致無法排序)
   StrSQLa = "select cp01,cp02,cp03,cp04,decode(cp01,'CFP',decode(mod(cp02,2),1,na73,na74),null) PID,type,pa09,cp01||cp02||cp03||cp04 CASENO from nation, " & _
             "(select DISTINCT r030001 as cp01,r030002 as cp02,r030003 as cp03,r030004 as cp04,r030005 as type,nvl(nvl(nvl(pa09,tm10),sp09),lc15) pa09 " & _
             "   from r040320_t1,patent,trademark,servicepractice,lawcase where id='" & strUserNum & "' and r030001=pa01(+) and r030002=pa02(+) and r030003=pa03(+) and r030004=pa04(+) " & _
             "    and r030001=tm01(+) and r030002=tm02(+) and r030003=tm03(+) and r030004=tm04(+) " & _
             "    and r030001=sp01(+) and r030002=sp02(+) and r030003=sp03(+) and r030004=sp04(+) " & _
             "    and r030001=lc01(+) and r030002=lc02(+) and r030003=lc03(+) and r030004=lc04(+) " & _
             ") where pa09=na01(+) order by 1,7,6,2,3,4 "
   'StrSQLa = "select DISTINCT r030001 as cp01,r030002 as cp02,r030003 as cp03,r030004 as cp04,r030005 as type from r040320_t1 where id='" & strUserNum & "' and r030001||R030002='CFT018925' order by 1,2,3,4 "
   'modify by sonia 2020/3/5 CFP管制人 109/4/1以後改業務區劃分
   'Set rs5 = New ADODB.Recordset
   'rs5.CursorLocation = adUseClient
   'rs5.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   Set rsQuery = New ADODB.Recordset
   rsQuery.CursorLocation = adUseClient
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, StrSQLa)
   If strSrvDate(1) >= CFP業務區劃分啟用日 Then
      If intI = 1 Then
         Set rs5 = PUB_CreateRecordset(rsQuery, , , 300, Me.Name)
         With rs5
         .MoveFirst
         Do While Not .EOF
            If .Fields("CP01") = "CFP" Or .Fields("CP01") = "CPS" Then
               .Fields("PID") = PUB_GetCFPHandler(.Fields("CASENO"))
            End If
            .MoveNext
         Loop
         .Sort = "cp01,PID,pa09,cp02,cp03,cp04"
         End With
      Else
         Set rs5 = rsQuery
      End If
   Else
      Set rs5 = rsQuery
   End If
   'end 2020/3/5

Dim tmpState As String
   
   If rs5.RecordCount <> 0 Then
      cnnConnection.Execute "delete from r040330 where id='" & strUserNum & "' "
      rs5.MoveFirst
      Do While Not rs5.EOF
         'If CheckStr(rs5.Fields("cp01")) = "TF" Then
         '   Debug.Print CheckStr(rs5.Fields("cp01")) & CheckStr(rs5.Fields("cp02")) & CheckStr(rs5.Fields("cp03")) & CheckStr(rs5.Fields("cp04")) & " " & CheckStr(rs5.Fields("PID"))
         'End If
         tmpState = PrintByOne(CheckStr(rs5.Fields("cp01")), CheckStr(rs5.Fields("cp02")), CheckStr(rs5.Fields("cp03")), CheckStr(rs5.Fields("cp04")), CheckStr(rs5.Fields("type")))
         
         If tmpState <> "" Then
'            Select Case tmpState
'               Case "1"   'MsgBox "此本所案號不存在！", vbCritical, "資料錯誤！"
'               Case "2"   'MsgBox "此本所案號還有帳單未付款"
'               Case "3"   'MsgBox "此本所案號還有國內帳款未收"
'               Case "4"   'MsgBox "此本所案號還有國外帳款未收"
'               Case "5"   'ShowNoData
'               Case "6"   'MsgBox Err.Description, vbCritical, "錯誤發生！"
'               Case Else
'            End Select
            If tmpState >= "2" And tmpState <= "4" Then
               '2011/11/9 modify by sonia cp的帳單抓不到acc150的案號不會寫入CFP009115(U09203890),故改存CP的資料
               'cnnConnection.Execute "insert into r040330 (A1514,A1501,A1504,A1502,A1505,AXF04,AXF02,A1503,A1516,ID,state) (select A1514,A1501,A1504,A1502,A1505,AXF04,AXF02,A1503,A1516,'" & strUserNum & "','2' from Acc150, Acc151 " & _
                                     " where A1501=AXF01(+) and axf02||a1501 in (select cp09||cp61 from (select cp09,cp61 from caseprogress where cp01='" & CheckStr(rs5.Fields("cp01")) & "'  " & strSQLCP & " and cp59 is null  union select cp09,cp62 from caseprogress where cp01='" & CheckStr(rs5.Fields("cp01")) & "' " & strSQLCP & " and cp59 is null  union select cp09,cp63 from caseprogress where cp01='" & CheckStr(rs5.Fields("cp01")) & "' " & strSQLCP & " and cp59 is null  union select cp09,cp87 from caseprogress where cp01='" & CheckStr(rs5.Fields("cp01")) & "' " & strSQLCP & " and cp59 is null  union select cp09,cp88 from caseprogress where cp01='" & CheckStr(rs5.Fields("cp01")) & "' " & strSQLCP & " and cp59 is null  ) AA,acc190,ACC150 where AA.cp61 is not null and AA.cp61=A1902(+) and AA.cp61=A1501(+) and A1902 is null AND A1512 IS NULL and a1507 is null )) "
               
               cnnConnection.BeginTrans   '2012/7/12 add by sonia
               '2012/7/12 modify by sonia 加可結餘刪除日期,cp146 is null條件
               'modify by sonia 2025/4/28 cp146 is null改為(cp146 is null or cp109>cp146)
               cnnConnection.Execute "insert into r040330 (A1514,A1501,A1504,A1502,A1505,AXF04,AXF02,A1503,A1516,ID,state) (select A1514,CP61,A1504,A1502,A1505,AXF04,CP09,A1503,A1516,'" & strUserNum & "','2' from Acc150, Acc151, " & _
                                     " (select cp09,cp61 from (select cp09,cp61 from caseprogress where cp01='" & CheckStr(rs5.Fields("cp01")) & "'  " & strSQLCP & " and cp59 is null and (cp146 is null or cp109>cp146) union select cp09,cp62 from caseprogress where cp01='" & CheckStr(rs5.Fields("cp01")) & "' " & strSQLCP & " and cp59 is null and (cp146 is null or cp109>cp146) " & _
                                     " union select cp09,cp63 from caseprogress where cp01='" & CheckStr(rs5.Fields("cp01")) & "' " & strSQLCP & " and cp59 is null and (cp146 is null or cp109>cp146) union select cp09,cp87 from caseprogress where cp01='" & CheckStr(rs5.Fields("cp01")) & "' " & strSQLCP & " and cp59 is null and (cp146 is null or cp109>cp146) union select cp09,cp88 from caseprogress where cp01='" & CheckStr(rs5.Fields("cp01")) & "' " & strSQLCP & " and cp59 is null and (cp146 is null or cp109>cp146)) AA,acc190,ACC150 where AA.cp61 is not null and AA.cp61=A1902(+) and AA.cp61=A1501(+) and A1902 is null AND A1512 IS NULL and a1507 is null ) BB where AXF01=A1501(+) and BB.cp09=axf02(+) and BB.cp61=axf01(+) ) "
               cnnConnection.Execute "insert into r040330 (A1514,A1501,A1504,A1502,A1505,AXF04,AXF02,A1503,A1516,ID,state) (select A0k24,A0k01,'',A0k02,'NTD',cp79,cp09,A0k03,a0k26,'" & strUserNum & "','3' from Acc0k0,caseprogress where cp01='" & CheckStr(rs5.Fields("cp01")) & "' " & strSQLCP & " and cp59 is null and (cp146 is null or cp109>cp146) and cp60 is not null and substr(cp60,1,1)='E' and cp79>0 and cp60=A0k01(+)  ) "
               '2012/7/30 modify by sonia 銷帳不管a1k25
               cnnConnection.Execute "insert into r040330 (A1514,A1501,A1504,A1502,A1505,AXF04,AXF02,A1503,A1516,ID,state) (select A1k19,A1k01,'',A1k02,'NTD',nvl(a1k11,0) - nvl(a1k06,0) - nvl(A1k30,0),cp09,A1k03,A1k21,'" & strUserNum & "','4' from Acc1K0,caseprogress,acc0z0 where cp01='" & CheckStr(rs5.Fields("cp01")) & "' " & strSQLCP & "  and cp60=a0z02(+) and cp60=a1k01(+) and cp60 is not null and substr(cp60,1,1)='X' and a1k25 is null and (a0z02 is null or a1k29<>'Y' or a1k29 is null) and cp59 is null and (cp146 is null or cp109>cp146)) "
               '2012/7/12 ADD BY SONIA 未收未付從暫存檔刪除
               cnnConnection.Execute "delete r040320_t1 where id='" & strUserNum & "' and r030001='" & CheckStr(rs5.Fields("cp01")) & "' and r030002='" & CheckStr(rs5.Fields("cp02")) & "' and r030003='" & CheckStr(rs5.Fields("cp03")) & "'"
               cnnConnection.CommitTrans
               '2012/7/12 end
            End If
         End If
         rs5.MoveNext
      Loop
   'add by sonia 2020/2/15
      txt1(2) = ""
      txt1(3) = "0"
      txt1(4) = "00"
   Else
      ShowNoData
   End If
   Set rs5 = Nothing

'add by nickc 2006/02/27 列印
Dim StrSQLBy040330 As String
   
   StrSQLBy040330 = "select " & SqlDateT("aa.A1514") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,decode(aa.state,'2','帳單未付','3','國內帳未收','國外帳未收'),'',"
   StrSQLBy040330 = StrSQLBy040330 & " decode(substr(aa.A1501,1,1),'U',aa.A1501,'*'||aa.A1501) as A1501,aa.A1504," & SqlDateT("aa.A1502") & ",aa.A1505,aa.AXF04,aa.A1503,decode(pa09,'013',fa04,'020',fa04,fa05||' '||fa63||' '||fa64||' '||fa65),st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as oSort "
   StrSQLBy040330 = StrSQLBy040330 & " From r040330 aa, Caseprogress, patent, nation, casepropertymap, staff, fagent "
   StrSQLBy040330 = StrSQLBy040330 & " where aa.AXF02=cp09(+) and id='" & strUserNum & "' "
   StrSQLBy040330 = StrSQLBy040330 & "  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09=na01(+) "
   StrSQLBy040330 = StrSQLBy040330 & " and cp01=cpm01(+) and cp10=cpm02(+) and substr(aa.A1503,1,8)=fa01(+) and substr(aa.A1503,9,1)=fa02(+) "
   StrSQLBy040330 = StrSQLBy040330 & " and aa.A1516=st01(+) and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 1) & ")  "
   StrSQLBy040330 = StrSQLBy040330 & " union "
   
   StrSQLBy040330 = StrSQLBy040330 & " select " & SqlDateT("aa.A1514") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,decode(aa.state,'2','帳單未付','3','國內帳未收','國外帳未收'),'',"
   StrSQLBy040330 = StrSQLBy040330 & " decode(substr(aa.A1501,1,1),'U',aa.A1501,'*'||aa.A1501) as A1501,aa.A1504," & SqlDateT("aa.A1502") & ",aa.A1505,aa.AXF04,aa.A1503,decode(tm10,'013',fa04,'020',fa04,fa05||' '||fa63||' '||fa64||' '||fa65),st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as oSort  "
   StrSQLBy040330 = StrSQLBy040330 & " From r040330 aa, Caseprogress, trademark, nation, casepropertymap, staff, fagent "
   StrSQLBy040330 = StrSQLBy040330 & " where aa.AXF02=cp09(+) and id='" & strUserNum & "' "
   StrSQLBy040330 = StrSQLBy040330 & "  and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and tm10=na01(+) "
   StrSQLBy040330 = StrSQLBy040330 & " and cp01=cpm01(+) and cp10=cpm02(+) and substr(aa.A1503,1,8)=fa01(+) and substr(aa.A1503,9,1)=fa02(+) "
   StrSQLBy040330 = StrSQLBy040330 & " and aa.A1516=st01(+) and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 2) & ")  "
   StrSQLBy040330 = StrSQLBy040330 & " union "
   
   StrSQLBy040330 = StrSQLBy040330 & " select " & SqlDateT("aa.A1514") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,decode(aa.state,'2','帳單未付','3','國內帳未收','國外帳未收'),'',"
   StrSQLBy040330 = StrSQLBy040330 & " decode(substr(aa.A1501,1,1),'U',aa.A1501,'*'||aa.A1501) as A1501,aa.A1504," & SqlDateT("aa.A1502") & ",aa.A1505,aa.AXF04,aa.A1503,decode(lc15,'013',fa04,'020',fa04,fa05||' '||fa63||' '||fa64||' '||fa65),st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as oSort  "
   StrSQLBy040330 = StrSQLBy040330 & " From r040330 aa, Caseprogress, lawcase, nation, casepropertymap, staff, fagent "
   StrSQLBy040330 = StrSQLBy040330 & " where aa.AXF02=cp09(+) and id='" & strUserNum & "' "
   StrSQLBy040330 = StrSQLBy040330 & "  and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and lc15=na01(+) "
   StrSQLBy040330 = StrSQLBy040330 & " and cp01=cpm01(+) and cp10=cpm02(+) and substr(aa.A1503,1,8)=fa01(+) and substr(aa.A1503,9,1)=fa02(+) "
   StrSQLBy040330 = StrSQLBy040330 & " and aa.A1516=st01(+) and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 3) & ")  "
   StrSQLBy040330 = StrSQLBy040330 & " union "
   
   StrSQLBy040330 = StrSQLBy040330 & " select " & SqlDateT("aa.A1514") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,decode(aa.state,'2','帳單未付','3','國內帳未收','國外帳未收'),'',"
   StrSQLBy040330 = StrSQLBy040330 & " decode(substr(aa.A1501,1,1),'U',aa.A1501,'*'||aa.A1501) as A1501,aa.A1504," & SqlDateT("aa.A1502") & ",aa.A1505,aa.AXF04,aa.A1503,decode(sp09,'013',fa04,'020',fa04,fa05||' '||fa63||' '||fa64||' '||fa65),st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as oSort  "
   StrSQLBy040330 = StrSQLBy040330 & " From r040330 aa, Caseprogress, servicepractice, nation, casepropertymap, staff, fagent "
   StrSQLBy040330 = StrSQLBy040330 & " where aa.AXF02=cp09(+) and id='" & strUserNum & "' "
   StrSQLBy040330 = StrSQLBy040330 & "  and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and sp09=na01(+) "
   StrSQLBy040330 = StrSQLBy040330 & " and cp01=cpm01(+) and cp10=cpm02(+) and substr(aa.A1503,1,8)=fa01(+) and substr(aa.A1503,9,1)=fa02(+) "
   StrSQLBy040330 = StrSQLBy040330 & " and aa.A1516=st01(+) and cp01 in (" & SQLGrpStr(GetAllSysKind(Text1), 5) & ")  "
   StrSQLBy040330 = StrSQLBy040330 & " order by oSort "
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open StrSQLBy040330, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
          Set grd2.Recordset = AdoRecordSet3
          PUB_RestorePrinter "PDF reDirect v2"   'add by sonia 2017/12/5瑞婷說此表固定產生PDF
          PrintData
          PUB_RestorePrinter strPrinter          'add by sonia 2017/12/5還原
      End If
   End With
   CheckOC3
   
   'add by sonia 2025/6/27 加印當日自動作廢之結餘單
   StrSQLBy040330 = "select '作廢日期 '||sqldatet(a240003+19110000)||'     結餘單號 '||a240002||'    本所案號 '||a240005||'-'||a240006||'-'||a240007||'-'||a240008 from acc240 " & _
                    "where a240003=" & strSrvDate(2) & " order by A240002"
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open StrSQLBy040330, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         PUB_RestorePrinter "PDF reDirect v2"
         tmpY = 1200
         Printer.CurrentX = 50
         Printer.CurrentY = tmpY
         Printer.Print "    今日自動作廢結餘單："
         tmpY = tmpY + 300
         PrintData1
         PUB_RestorePrinter strPrinter
      End If
   End With
   CheckOC3
   
   '加印J公司銷項稅額為0之結餘單清單
   StrSQLBy040330 = "select '結餘單號 '||a240002||'    本所案號 '||a240005||'-'||a240006||'-'||a240007||'-'||a240008 from " & _
                    "(SELECT a240002,a240005,a240006,a240007,a240008 FROM ACC431,ACC430," & _
                    "(SELECT DISTINCT A0J13,a240002,a240005,a240006,a240007,a240008 FROM acc240,CASEPROGRESS,ACC0J0 WHERE a240001=" & strSrvDate(2) & " and a240002=CP59(+) AND CP09=A0J01(+)) " & _
                    "WHERE A0J13=AXC02(+) AND AXC01=A4301(+) and axc01 is not null group by a240002,a240005,a240006,a240007,a240008 having SUM(A4305)=0) order by a240002"
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open StrSQLBy040330, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         PUB_RestorePrinter "PDF reDirect v2"
         tmpY = 1200
         Printer.CurrentX = 50
         Printer.CurrentY = tmpY
         Printer.Print "    J 公司銷項稅額為０之結餘單："
         tmpY = tmpY + 300
         PrintData1
         PUB_RestorePrinter strPrinter
      End If
   End With
   CheckOC3
   'end 2025/6/27
End Sub

Sub ClearData()
   txt1(0).Text = ""
   txt1(1).Text = ""
   txt1(2).Text = ""
   txt1(3).Text = ""
   txt1(4).Text = ""
   txt1(5).Text = ""
End Sub

'****************
'列印結餘單--重印
'****************
Sub PrintData2_1Old(oA240002 As String)
Dim iVarCp10  As Variant
Dim iCountCp10 As Integer
'add by sonia 2014/6/10
Dim strCompName As String
Dim strCompNo As String
'end 2014/6/10

   For i = 0 To 20
       If i <> 1 Then
           strTemp(i) = ""
       End If
   Next i
   '取得編號
   strSql = "SELECT * FROM Acc240,Acc241 WHERE A240002=A241001(+) and A240002='" & oA240002 & "' order by A241002 "
   CheckOC
   SeekTemp1 = "  "    '本所案號
   SeekTemp2 = "   "    '收文號
   Page = 1
   
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         strTemp(0) = strUserName
         strTemp(1) = CheckStr(.Fields("A240002"))
   'edit by nickc 2005/07/21 製表日=結餘日期，原製表日欄位改為作廢日
   '      strTemp(2) = Format(ChangeTStringToTDateString(CheckStr(.Fields("A240003"))), "YY")
   '      strTemp(3) = Format(ChangeTStringToTDateString(CheckStr(.Fields("A240003"))), "mm")
   '      strTemp(4) = Format(ChangeTStringToTDateString(CheckStr(.Fields("A240003"))), "dd")
         '2011/2/24 add by sonia
         'strTemp(2) = Format(ChangeTStringToTDateString(CheckStr(.Fields("A240001"))), "YY")
         strTemp(2) = CheckStr(.Fields("A240001")) \ 10000
         '2011/2/24 end
         strTemp(3) = Format(ChangeTStringToTDateString(CheckStr(.Fields("A240001"))), "mm")
         strTemp(4) = Format(ChangeTStringToTDateString(CheckStr(.Fields("A240001"))), "dd")
         strTemp(5) = CheckStr(.Fields("A240005")) & "-" & CheckStr(.Fields("A240006")) & "-" & CheckStr(.Fields("A240007")) & "-" & CheckStr(.Fields("A240008"))
         'add by nickc 2005/09/21
         m_CP01 = CheckStr(.Fields("A240005"))
         m_CP02 = CheckStr(.Fields("A240006"))
         m_CP03 = CheckStr(.Fields("A240007"))
         m_CP04 = CheckStr(.Fields("A240008"))
         
         strTemp(6) = CheckStr(.Fields("A240009"))
         strTemp(7) = CheckStr(.Fields("A240010"))
         strTemp(8) = CheckStr(.Fields("A240011"))
         strTemp(9) = CheckStr(.Fields("A240012"))
      End If
   
      '2009/11/23 modify by sonia 抓案件備註
      'strSQL = "select cp109 from caseprogress where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "' and cp59='" & strTemp(1) & "' "
      Select Case m_CP01
         Case "T", "CFT", "TF"
            strSql = "select cp109,tm58 as remark from caseprogress,trademark where tm01='" & m_CP01 & "' and tm02='" & m_CP02 & "' and tm03='" & m_CP03 & "' and tm04='" & m_CP04 & "' and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and '" & strTemp(1) & "'=cp59(+) "
         Case "P", "CFP"
            strSql = "select cp109,pa91 as remark from caseprogress,patent where pa01='" & m_CP01 & "' and pa02='" & m_CP02 & "' and pa03='" & m_CP03 & "' and pa04='" & m_CP04 & "' and pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+) and '" & strTemp(1) & "'=cp59(+) "
         Case "L", "CFL", "FCL", "LIN"
            strSql = "select cp109,lc27 as remark from caseprogress,lawcase where lc01='" & m_CP01 & "' and lc02='" & m_CP02 & "' and lc03='" & m_CP03 & "' and lc04='" & m_CP04 & "' and lc01=cp01(+) and lc02=cp02(+) and lc03=cp03(+) and lc04=cp04(+) and '" & strTemp(1) & "'=cp59(+) "
         Case "LA"
            strSql = "select cp109,hc12 as remark from caseprogress,hirecase where hc01='" & m_CP01 & "' and hc02='" & m_CP02 & "' and hc03='" & m_CP03 & "' and hc04='" & m_CP04 & "' and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) and '" & strTemp(1) & "'=cp59(+) "
         Case Else
            strSql = "select cp109,sp18 as remark from caseprogress,servicepractice where sp01='" & m_CP01 & "' and sp02='" & m_CP02 & "' and sp03='" & m_CP03 & "' and sp04='" & m_CP04 & "' and sp01=cp01(+) and sp02=cp02(+) and sp03=cp03(+) and sp04=cp04(+) and '" & strTemp(1) & "'=cp59(+) "
      End Select
      '2009/11/23 end
      m_Title2 = "(逾期未處理)"
      m_Remark = ""
      CheckOC3
      AdoRecordSet3.CursorLocation = adUseClient
      AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If AdoRecordSet3.RecordCount <> 0 Then
         AdoRecordSet3.MoveFirst
         If Not IsNull(AdoRecordSet3.Fields("cp109")) Then m_Title2 = ""
         If Not IsNull(AdoRecordSet3.Fields("remark")) Then m_Remark = AdoRecordSet3.Fields("remark")  '2009/11/24 add by sonia
      End If
      
      'Added by Lydia 2019/07/09 因為案件備註過長,所以只抓"合併計算結餘"含前後各30個字
      If m_Remark <> "" And InStr(m_Remark, "合併計算結餘") Then
          i = InStr(m_Remark, "合併計算結餘")
          If i <= 30 Then
               m_Remark = Mid(m_Remark, 1, i + Len("合併計算結餘") + 30)
          Else
               m_Remark = Mid(m_Remark, i - 30, Len("合併計算結餘") + 60)
          End If
          m_Remark = PUB_StringFilter(m_Remark) '去掉字串裡的跳行符號
      Else
          m_Remark = ""
      End If
      
      Printer.Orientation = 1
      If InStr(Printer.DeviceName, "1200") > 0 Then
         Printer.ScaleWidth = Printer.Width
         Printer.ScaleHeight = Printer.Height
      End If
      
      Printer.Font.Size = 22
      Printer.Font.Name = "細明體"
      Printer.Font.Bold = True
      Printer.Font.Underline = True
      Printer.CurrentX = 5500 - (Printer.TextWidth("CF 案件結餘轉帳通知單") / 2)
      Printer.CurrentY = 500
      Printer.Print "CF 案件結餘轉帳通知單"
      Printer.Font.Size = 12
      Printer.Font.Bold = False
      Printer.Font.Underline = False
   
      Printer.CurrentX = 0
      Printer.CurrentY = 1000
      Printer.Print "製表人：" & GetPrjSalesNM(strUserNum)
      
      Printer.CurrentX = 5500 - (Printer.TextWidth(m_Title2) / 2)
      Printer.CurrentY = 1000
      Printer.Print m_Title2
      
      Printer.CurrentX = 11500 - (Printer.TextWidth("製表日期： " & ChgSQL(strTemp(2)) & " 年 " & ChgSQL(strTemp(3)) & " 月 " & ChgSQL(strTemp(4)) & " 日 "))
      Printer.CurrentY = 1000
      Printer.Print "申請國家：" & strTemp(8)
      Printer.CurrentX = 0
      Printer.CurrentY = 1300
      Printer.Print "結餘單號：" & strTemp(1)
      'add by sonia 2014/6/10 智權公司案件要加印
      strCompNo = ""
      strCompName = GetSpecialComp(m_CP01, m_CP02, m_CP03, m_CP04, strCompNo, 6)
      'modify by sonia 2020/4/16 +L公司
      If strCompNo = "J" Or strCompNo = "L" Then
         Printer.Font.Size = 22
         Printer.Font.Bold = True
         Printer.CurrentX = 4300
         Printer.CurrentY = 1300
         If strCompNo = "J" Then
            Printer.Print "(智權)"
         Else
            Printer.Print "(法律)"
         End If
         Printer.Font.Size = 12
         Printer.Font.Bold = False
      End If
      'end 2014/6/10
      Printer.CurrentX = 11500 - (Printer.TextWidth("製表日期： " & ChgSQL(strTemp(2)) & " 年 " & ChgSQL(strTemp(3)) & " 月 " & ChgSQL(strTemp(4)) & " 日 "))
      Printer.CurrentY = 1300
      Printer.Print "製表日期：" & ChgSQL(strTemp(2)) & " 年 " & ChgSQL(strTemp(3)) & " 月 " & ChgSQL(strTemp(4)) & " 日 "
      
      Printer.Font.Size = 22
      Printer.Font.Bold = True
      For i = 0 To 10
      '下外框
         Printer.Line (400 + i, 13800 + i)-(11000 + i, 13800 + i)
         Printer.Line (400 + i, 16100 + i)-(11000 + i, 16100 + i)
         Printer.Line (400 + i, 13800 + i)-(400 + i, 16100 + i)
         Printer.Line (11000 + i, 13800 + i)-(11000 + i, 16100 + i)
      '下內
         'modify by sonia 2023/2/16 取消 業務部副總經理
         'Printer.Line (1100 + i, 13800 + i)-(1100 + i, 16100 + i)
         'Printer.Line (2500 + i, 13800 + i)-(2500 + i, 16100 + i)
         'Printer.Line (3200 + i, 13800 + i)-(3200 + i, 16100 + i)
         'Printer.Line (4600 + i, 13800 + i)-(4600 + i, 16100 + i)
         'Printer.Line (5300 + i, 13800 + i)-(5300 + i, 16100 + i)
         'Printer.Line (6700 + i, 13800 + i)-(6700 + i, 16100 + i)
         'Printer.Line (7400 + i, 13800 + i)-(7400 + i, 16100 + i)
         'Printer.Line (8800 + i, 13800 + i)-(8800 + i, 16100 + i)
         'Printer.Line (9500 + i, 13800 + i)-(9500 + i, 16100 + i)
         Printer.Line (1100 + i, 13800 + i)-(1100 + i, 16100 + i)
         Printer.Line (3050 + i, 13800 + i)-(3050 + i, 16100 + i)
         Printer.Line (3750 + i, 13800 + i)-(3750 + i, 16100 + i)
         Printer.Line (5700 + i, 13800 + i)-(5700 + i, 16100 + i)
         Printer.Line (6400 + i, 13800 + i)-(6400 + i, 16100 + i)
         Printer.Line (8350 + i, 13800 + i)-(8350 + i, 16100 + i)
         Printer.Line (9050 + i, 13800 + i)-(9050 + i, 16100 + i)
         'end 2023/2/16
      ''上內
      '   Printer.Line (1400 + i, 2200 + i)-(4100 + i, 2200 + i)
      '   Printer.Line (5200 + i, 2200 + i)-(8100 + i, 2200 + i)
      '   Printer.Line (9300 + i, 2200 + i)-(11500 + i, 2200 + i)
      Next i
   
      Printer.Font.Size = 12
      Printer.CurrentX = 300
      Printer.CurrentY = 1900
      Printer.Print "本所案號：" & strTemp(5)
      
      Printer.CurrentX = 4300
      Printer.CurrentY = 1900
      Printer.Print "申請人：" & strTemp(6)
      Printer.CurrentX = 8300
      Printer.CurrentY = 1900
      'Printer.Print "智權人員：" & GetPrjSalesNM(strTemp(7))   'cancel by sonia 2022/6/30
      
      '2009/11/23 add by sonia
      'Modified by Lydia 2019/07/09 改成兩行
      'Printer.CurrentX = 300
      'Printer.CurrentY = 2200
      'Printer.Print "案件備註：" & m_Remark
      '2009/11/23 end
      Pub_SmartPrint "案件備註：" & m_Remark, 300, 2200, 220
      
      Printer.CurrentX = 400 + GetPrPosX(700, "總")
      Printer.CurrentY = 13900
      Printer.Print "總"
      Printer.CurrentX = 400 + GetPrPosX(700, "經")
      Printer.CurrentY = 14800
      Printer.Print "經"
      Printer.CurrentX = 400 + GetPrPosX(700, "理")
      Printer.CurrentY = 15700
      Printer.Print "理"
      
      'modify by sonia 2023/2/16 取消 業務部副總經理,以下調整位置
      Printer.CurrentX = 3050 + GetPrPosX(700, "國")    '2023/2/16 2500改為3050
      Printer.CurrentY = 13900
      Printer.Print "國"
      Printer.CurrentX = 3050 + GetPrPosX(700, "外")
      Printer.CurrentY = 14200
      Printer.Print "外"
      Printer.CurrentX = 3050 + GetPrPosX(700, "部")
      Printer.CurrentY = 14500
      Printer.Print "部"
      Printer.CurrentX = 3050 + GetPrPosX(700, "副")
      Printer.CurrentY = 14800
      Printer.Print "副"
      Printer.CurrentX = 3050 + GetPrPosX(700, "總")
      Printer.CurrentY = 15100
      Printer.Print "總"
      Printer.CurrentX = 3050 + GetPrPosX(700, "經")
      Printer.CurrentY = 15400
      Printer.Print "經"
      Printer.CurrentX = 3050 + GetPrPosX(700, "理")
      Printer.CurrentY = 15700
      Printer.Print "理"
      
      'cancel by sonia 2023/2/16 取消 業務部副總經理
      'Printer.CurrentX = 4600 + GetPrPosX(700, "業")
      'Printer.CurrentY = 13900
      'Printer.Print "業"
      'Printer.CurrentX = 4600 + GetPrPosX(700, "務")
      'Printer.CurrentY = 14200
      'Printer.Print "務"
      'Printer.CurrentX = 4600 + GetPrPosX(700, "部")
      'Printer.CurrentY = 14500
      'Printer.Print "部"
      'Printer.CurrentX = 4600 + GetPrPosX(700, "副")
      'Printer.CurrentY = 14800
      'Printer.Print "副"
      'Printer.CurrentX = 4600 + GetPrPosX(700, "總")
      'Printer.CurrentY = 15100
      'Printer.Print "總"
      'Printer.CurrentX = 4600 + GetPrPosX(700, "經")
      'Printer.CurrentY = 15400
      'Printer.Print "經"
      'Printer.CurrentX = 4600 + GetPrPosX(700, "理")
      'Printer.CurrentY = 15700
      'Printer.Print "理"
      'end 2023/2/16
      
      Printer.CurrentX = 5700 + GetPrPosX(700, "國")    '2023/2/16 6700改為5700
      Printer.CurrentY = 13900
      Printer.Print "國"
      Printer.CurrentX = 5700 + GetPrPosX(700, "外")
      Printer.CurrentY = 14200
      Printer.Print "外"
      Printer.CurrentX = 5700 + GetPrPosX(700, "部")
      Printer.CurrentY = 14500
      Printer.Print "部"
      Printer.CurrentX = 5700 + GetPrPosX(700, "單")
      Printer.CurrentY = 14800
      Printer.Print "單"
      Printer.CurrentX = 5700 + GetPrPosX(700, "位")
      Printer.CurrentY = 15100
      Printer.Print "位"
      Printer.CurrentX = 5700 + GetPrPosX(700, "主")
      Printer.CurrentY = 15400
      Printer.Print "主"
      Printer.CurrentX = 5700 + GetPrPosX(700, "管")
      Printer.CurrentY = 15700
      Printer.Print "管"
      
      Printer.CurrentX = 8350 + GetPrPosX(700, "國")    '2023/2/16 8800改為8350
      Printer.CurrentY = 13900
      Printer.Print "國"
      Printer.CurrentX = 8350 + GetPrPosX(700, "外")
      Printer.CurrentY = 14200
      Printer.Print "外"
      Printer.CurrentX = 8350 + GetPrPosX(700, "部")
      Printer.CurrentY = 14500
      Printer.Print "部"
      Printer.CurrentX = 8350 + GetPrPosX(700, "承")
      Printer.CurrentY = 15100
      Printer.Print "承"
      Printer.CurrentX = 8350 + GetPrPosX(700, "辦")
      Printer.CurrentY = 15400
      Printer.Print "辦"
      Printer.CurrentX = 8350 + GetPrPosX(700, "人")
      Printer.CurrentY = 15700
      Printer.Print "人"
   
      tmpY = 2600
      Printer.Font.Bold = False
      Printer.CurrentX = 300
      Printer.CurrentY = tmpY
      Printer.Print "代理人："
      Printer.CurrentX = 300 + Printer.TextWidth("代理人：") + 200
      Printer.CurrentY = tmpY
      Printer.Print strTemp(9)
      tmpY = tmpY + 300
      Printer.CurrentX = 300
      Printer.CurrentY = tmpY
      Printer.Print "代理人請款金額："
      strTemp(10) = CheckStr(.Fields("A240013"))
      If Right(strTemp(10), 1) = ";" Then
         strTemp(10) = Mid(strTemp(10), 1, Len(strTemp(10)) - 1)
      End If
      iVarCp10 = Split(strTemp(10), ";")
      Dim oPrintStr As String
      oPrintStr = ""
      For iCountCp10 = 0 To UBound(iVarCp10) Step 3
         oPrintStr = ""
         Printer.CurrentX = 300 + Printer.TextWidth("代理人請款金額：") + 200
         Printer.CurrentY = tmpY
         If (iCountCp10) <= (UBound(iVarCp10) + 1) Then
            If (UBound(iVarCp10) + 1) - (iCountCp10 + 1) >= 0 Then
               oPrintStr = oPrintStr & iVarCp10(iCountCp10) & ";"
            End If
         End If
         If (iCountCp10 + 1) <= (UBound(iVarCp10) + 1) Then
            If (UBound(iVarCp10) + 1) - (iCountCp10 + 2) >= 0 Then
               oPrintStr = oPrintStr & iVarCp10(iCountCp10 + 1) & ";"
            End If
         End If
         If (iCountCp10 + 2) <= (UBound(iVarCp10) + 1) Then
            If (UBound(iVarCp10) + 1) - (iCountCp10 + 3) >= 0 Then
               oPrintStr = oPrintStr & iVarCp10(iCountCp10 + 2) & ";"
            End If
         End If
         Printer.Print oPrintStr
         Printer.Font.Bold = True
         Printer.Font.Bold = False
         tmpY = tmpY + 300
      Next iCountCp10
   
      'add by sonia 2023/7/27 CFT、CFC及S案加印承辦人
      If m_CP01 = "CFT" Or m_CP01 = "CFC" Or m_CP01 = "S" Then
         Call GetNA69("", GetPrjNation1(strTemp(5)), strTemp(7), NA69Emp, m_CP01, m_CP02, m_CP03, m_CP04)
         Printer.CurrentX = 300
         Printer.CurrentY = tmpY + 600
         Printer.Print "承辦人：" & GetStaffName(NA69Emp)
      End If
      'end 2023/7/27
      
      Printer.Font.Bold = True
      'Printer.Line (1200 + i, tmpY + 250 + i)-(11500 + i, tmpY + 250 + i)
      tmpY = tmpY + 600
      Printer.CurrentX = 1300 + GetPrPosX(1500, "實際收款金額")
      Printer.CurrentY = tmpY + 300
      Printer.Print "實際收款金額"
      Printer.CurrentX = 2850 + GetPrPosX(400, "減")
      Printer.CurrentY = tmpY + 300
      Printer.Print "減"
      Printer.CurrentX = 3300 + GetPrPosX(1500, "已作收入金額")
      Printer.CurrentY = tmpY + 300
      Printer.Print "已作收入金額"
      Printer.CurrentX = 4850 + GetPrPosX(400, "減")
      Printer.CurrentY = tmpY + 300
      Printer.Print "減"
      Printer.CurrentX = 5300 + GetPrPosX(1500, "實際支出費用")
      Printer.CurrentY = tmpY + 300
      Printer.Print "實際支出費用"
      Printer.CurrentX = 6850 + GetPrPosX(600, "等於")
      Printer.CurrentY = tmpY + 300
      Printer.Print "等於"
      Printer.CurrentX = 7500 + GetPrPosX(1500, "結 轉 應 付")
      Printer.CurrentY = tmpY
      Printer.Print "結 轉 應 付"
      Printer.CurrentX = 7500 + GetPrPosX(1500, "規費安全基金")
      Printer.CurrentY = tmpY + 300
      Printer.Print "規費安全基金"
      Printer.CurrentX = 9050 + GetPrPosX(400, "加")
      Printer.CurrentY = tmpY + 300
      Printer.Print "加"
      Printer.CurrentX = 9500 + GetPrPosX(1500, "結轉收入金額")
      Printer.CurrentY = tmpY + 300
      Printer.Print "結轉收入金額"
      tmpY = tmpY + 600
      Printer.Font.Bold = False
      Printer.CurrentX = 0
      Printer.CurrentY = tmpY
      Printer.Print String(104, "-")
      tmpY = tmpY + 300
      strSql = "SELECT * FROM Acc240,Acc241 WHERE A240002=A241001(+) and A240002='" & oA240002 & "' and A241002<998 order by A241002 "
      CheckOC
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
           .MoveFirst
           Do While Not .EOF
               strTemp(11) = CheckStr(.Fields("A241003"))
               strTemp(12) = CheckStr(.Fields("A241004"))
               strTemp(13) = CheckStr(.Fields("A241005"))
               Printer.CurrentX = 1300 + GetPrPosX(1500, "實際收款金額") + Printer.TextWidth("實際收款金額") - Printer.TextWidth(Format(strTemp(11), "###,###,###,##0.00"))
               Printer.CurrentY = tmpY
               Printer.Print Format(strTemp(11), "###,###,###,##0.00")
               Printer.CurrentX = 3300 + GetPrPosX(1500, "已作收入金額") + Printer.TextWidth("已作收入金額") - Printer.TextWidth(Format(strTemp(12), "###,###,###,##0.00"))
               Printer.CurrentY = tmpY
               Printer.Print Format(strTemp(12), "###,###,###,##0.00")
               Printer.CurrentX = 5300 + GetPrPosX(1500, "實際支出費用")
               Printer.CurrentY = tmpY
               Printer.Print strTemp(13)
               tmpY = tmpY + 300
               .MoveNext
           Loop
      End If
      Printer.CurrentX = 0
      Printer.CurrentY = tmpY
      Printer.Print String(104, "-")
      tmpY = tmpY + 300
      Printer.CurrentX = 2850 + GetPrPosX(400, "-")
      Printer.CurrentY = tmpY
      Printer.Print "-"
      Printer.CurrentX = 4850 + GetPrPosX(400, "-")
      Printer.CurrentY = tmpY
      Printer.Print "-"
      Printer.CurrentX = 6850 + GetPrPosX(600, "=")
      Printer.CurrentY = tmpY
      Printer.Print "="
      Printer.CurrentX = 9050 + GetPrPosX(400, "+")
      Printer.CurrentY = tmpY
      Printer.Print "+"
   
      strSql = "SELECT * FROM Acc240,Acc241 WHERE A240002=A241001(+) and A240002='" & oA240002 & "' and A241002=998 order by A241002 "
      CheckOC
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
   '合計
         .MoveFirst
         strTemp(11) = CheckStr(.Fields("A241003"))
         strTemp(12) = CheckStr(.Fields("A241004"))
         strTemp(13) = CheckStr(.Fields("A241005"))
         strTemp(14) = CheckStr(.Fields("A241006"))
         strTemp(15) = CheckStr(.Fields("A241007"))
      End If
      Printer.CurrentX = 0
      Printer.CurrentY = tmpY
      Printer.Print "合計及計算："
      Printer.CurrentX = 1300 + GetPrPosX(1500, "實際收款金額") + Printer.TextWidth("實際收款金額") - Printer.TextWidth(Format(strTemp(11), "###,###,###,##0.00"))
      Printer.CurrentY = tmpY
      Printer.Print Format(strTemp(11), "###,###,###,##0.00")
      Printer.CurrentX = 3300 + GetPrPosX(1500, "已作收入金額") + Printer.TextWidth("已作收入金額") - Printer.TextWidth(Format(strTemp(12), "###,###,###,##0.00"))
      Printer.CurrentY = tmpY
      Printer.Print Format(strTemp(12), "###,###,###,##0.00")
      Printer.CurrentX = 5300 + GetPrPosX(1500, "實際支出費用") + Printer.TextWidth("實際支出費用") - Printer.TextWidth(Format(strTemp(13), "###,###,###,##0.00"))
      Printer.CurrentY = tmpY
      Printer.Print Format(strTemp(13), "###,###,###,##0.00")
      Printer.CurrentX = 7500 + GetPrPosX(1500, "規費安全基金") + Printer.TextWidth("規費安全基金") - Printer.TextWidth(Format(strTemp(14), "###,###,###,##0.00"))
      Printer.CurrentY = tmpY
      Printer.Print Format(strTemp(14), "###,###,###,##0.00")
      Printer.CurrentX = 9500 + GetPrPosX(1500, "結轉收入金額") + Printer.TextWidth("結轉收入金額") - Printer.TextWidth(Format(strTemp(15), "###,###,###,##0.00"))
      Printer.CurrentY = tmpY
      Printer.Print Format(strTemp(15), "###,###,###,##0.00")
      strSql = "SELECT * FROM Acc240,Acc241 WHERE A240002=A241001(+) and A240002='" & oA240002 & "' and A241002=999 order by A241002 "
      CheckOC
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
   '退費
         .MoveFirst
         strTemp(16) = CheckStr(.Fields("A241005"))
         tmpY = tmpY + 300
         Printer.CurrentX = 4850 + GetPrPosX(400, "退費：")
         Printer.CurrentY = tmpY
         Printer.Print "退費："
         Printer.CurrentX = 5300 + GetPrPosX(1500, "實際支出費用") + Printer.TextWidth("實際支出費用") - Printer.TextWidth(Format(strTemp(16), "###,###,###,##0.00"))
         Printer.CurrentY = tmpY
         Printer.Print Format(strTemp(16), "###,###,###,##0.00")
      End If
   
   'add by sonia 2018/9/11 J公司加印銷項稅額
   If strCompNo = "J" Then
      'modify by sonia 2021/7/19 CFT-022069(R110070180)一發票多收文號會重覆計算
      'strSql = "SELECT SUM(A4305) FROM CASEPROGRESS,ACC0J0,ACC431,ACC430 WHERE CP59='" & oA240002 & "' AND CP09=A0J01(+) AND A0J13=AXC02(+) AND AXC01=A4301(+) "
      strSql = "SELECT SUM(A4305) FROM ACC431,ACC430,(SELECT DISTINCT A0J13 FROM CASEPROGRESS,ACC0J0 WHERE CP59='" & oA240002 & "' AND CP09=A0J01(+)) WHERE A0J13=AXC02(+) AND AXC01=A4301(+) "
      CheckOC
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      'If Val(CheckStr(.Fields(0))) > 0 Then   'cancel by sonia 2025/6/27 無銷項稅額也要印但最後列清單
         strTemp(11) = CheckStr(.Fields(0))
         tmpY = tmpY + 300
         Printer.CurrentX = 0
         Printer.CurrentY = tmpY
         Printer.Print "銷項稅額："
         Printer.CurrentX = 1300 + GetPrPosX(1500, "實際收款金額") + Printer.TextWidth("實際收款金額") - Printer.TextWidth(Format(strTemp(11), "###,###,###,##0.00"))
         Printer.CurrentY = tmpY
         Printer.Print Format(strTemp(11), "###,###,###,##0.00")
      'End If    'cancel by sonia 2025/6/27
      
   End If
   'end 2018/9/11
   
   'add by sonia 2016/6/6 加印非出名公司之其他公司傳票資料
   strSql = "SELECT A242002||' 公司  傳票號碼 '||A242003||' 項次 '||A242004||'   借方'||TO_CHAR(AX206,'9,999,999.99')||'  貸方'||TO_CHAR(AX207,'9,999,999.99') FROM Acc242,ACC021 " & _
            "WHERE A242001='" & oA240002 & "' AND A242002=AX201(+) AND A242003=AX202(+) AND A242004=AX203(+) ORDER BY A242002,A242003,A242004 "
   CheckOC
   .CursorLocation = adUseClient
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      tmpY = tmpY + 1200
      Printer.CurrentX = 10
      Printer.CurrentY = tmpY
      Printer.Print "其他公司帳款："
      Do While Not .EOF
         tmpY = tmpY + 300
         Printer.CurrentX = 10
         Printer.CurrentY = tmpY
         Printer.Print CheckStr(.Fields(0))
         .MoveNext
      Loop
   End If
   'end 2016/6/6
   End With
   
   'CheckOC
   IsPrintok = True
   Printer.EndDoc
   'add by nickc 2008/04/08
   If Printer.ScaleMode = 0 Then Printer.ScaleMode = 1

End Sub

'****************
'列印結餘單 空白--重印
'****************
Sub PrintData2_SpaceOld(oA240002 As String)
Dim iVarCp10  As Variant
Dim iCountCp10 As Integer
'add by sonia 2022/11/3
Dim strCompName As String
Dim strCompNo As String
'end 2022/11/3
   
   For i = 0 To 20
       If i <> 1 Then
           strTemp(i) = ""
       End If
   Next i
   '取得編號
   strSql = "SELECT * FROM Acc240,Acc241 WHERE A240002=A241001(+) and A240002='" & oA240002 & "' and a240003 is null order by A241002 "
   CheckOC
   SeekTemp1 = "  "    '本所案號
   SeekTemp2 = "   "    '收文號
   Page = 1
   
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         strTemp(0) = strUserName
         strTemp(1) = CheckStr(.Fields("A240002"))
   'edit by nickc 2005/07/21 製表日=結餘日期，原製表日欄位改為作廢日
   '      strTemp(2) = Format(ChangeTStringToTDateString(CheckStr(.Fields("A240003"))), "YY")
   '      strTemp(3) = Format(ChangeTStringToTDateString(CheckStr(.Fields("A240003"))), "mm")
   '      strTemp(4) = Format(ChangeTStringToTDateString(CheckStr(.Fields("A240003"))), "dd")
         '2011/2/24 modify by sonia
         'strTemp(2) = Format(ChangeTStringToTDateString(CheckStr(.Fields("A240001"))), "YY")
         strTemp(2) = CheckStr(.Fields("A240001")) \ 10000
         '2011/2/24 end
         strTemp(3) = Format(ChangeTStringToTDateString(CheckStr(.Fields("A240001"))), "mm")
         strTemp(4) = Format(ChangeTStringToTDateString(CheckStr(.Fields("A240001"))), "dd")
         strTemp(5) = CheckStr(.Fields("A240005")) & "-" & CheckStr(.Fields("A240006")) & "-" & CheckStr(.Fields("A240007")) & "-" & CheckStr(.Fields("A240008"))
         'add by nickc 2005/09/21
         m_CP01 = CheckStr(.Fields("A240005"))
         m_CP02 = CheckStr(.Fields("A240006"))
         m_CP03 = CheckStr(.Fields("A240007"))
         m_CP04 = CheckStr(.Fields("A240008"))
         
         strTemp(6) = CheckStr(.Fields("A240009"))
         strTemp(7) = CheckStr(.Fields("A240010"))
         strTemp(8) = CheckStr(.Fields("A240011"))
         strTemp(9) = CheckStr(.Fields("A240012"))
      End If
   
      '2009/11/23 modify by sonia 抓案件備註
      'strSQL = "select cp109 from caseprogress where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "' and cp59='" & strTemp(1) & "' "
      Select Case m_CP01
         Case "T", "CFT", "TF"
            strSql = "select cp109,tm58 as remark from caseprogress,trademark where tm01='" & m_CP01 & "' and tm02='" & m_CP02 & "' and tm03='" & m_CP03 & "' and tm04='" & m_CP04 & "' and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and '" & strTemp(1) & "'=cp59(+) "
         Case "P", "CFP"
            strSql = "select cp109,pa91 as remark from caseprogress,patent where pa01='" & m_CP01 & "' and pa02='" & m_CP02 & "' and pa03='" & m_CP03 & "' and pa04='" & m_CP04 & "' and pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+) and '" & strTemp(1) & "'=cp59(+) "
         Case "L", "CFL", "FCL", "LIN"
            strSql = "select cp109,lc27 as remark from caseprogress,lawcase where lc01='" & m_CP01 & "' and lc02='" & m_CP02 & "' and lc03='" & m_CP03 & "' and lc04='" & m_CP04 & "' and lc01=cp01(+) and lc02=cp02(+) and lc03=cp03(+) and lc04=cp04(+) and '" & strTemp(1) & "'=cp59(+) "
         Case "LA"
            strSql = "select cp109,hc12 as remark from caseprogress,hirecase where hc01='" & m_CP01 & "' and hc02='" & m_CP02 & "' and hc03='" & m_CP03 & "' and hc04='" & m_CP04 & "' and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) and '" & strTemp(1) & "'=cp59(+) "
         Case Else
            strSql = "select cp109,sp18 as remark from caseprogress,servicepractice where sp01='" & m_CP01 & "' and sp02='" & m_CP02 & "' and sp03='" & m_CP03 & "' and sp04='" & m_CP04 & "' and sp01=cp01(+) and sp02=cp02(+) and sp03=cp03(+) and sp04=cp04(+) and '" & strTemp(1) & "'=cp59(+) "
      End Select
      '2009/11/23 end
      m_Title2 = "(逾期未處理)"
      m_Remark = ""
      CheckOC3
      AdoRecordSet3.CursorLocation = adUseClient
      AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If AdoRecordSet3.RecordCount <> 0 Then
         AdoRecordSet3.MoveFirst
         If Not IsNull(AdoRecordSet3.Fields("cp109")) Then m_Title2 = ""
         If Not IsNull(AdoRecordSet3.Fields("remark")) Then m_Remark = AdoRecordSet3.Fields("remark")  '2009/11/24 add by sonia
      End If
      
      'Added by Lydia 2019/07/09 因為案件備註過長,所以只抓"合併計算結餘"含前後各30個字
      If m_Remark <> "" And InStr(m_Remark, "合併計算結餘") Then
          i = InStr(m_Remark, "合併計算結餘")
          If i <= 30 Then
               m_Remark = Mid(m_Remark, 1, i + Len("合併計算結餘") + 30)
          Else
               m_Remark = Mid(m_Remark, i - 30, Len("合併計算結餘") + 60)
          End If
          m_Remark = PUB_StringFilter(m_Remark) '去掉字串裡的跳行符號
      Else
          m_Remark = ""
      End If
      
      Printer.Orientation = 1
'      If InStr(Printer.DeviceName, "1200") > 0 Then
'           Printer.ScaleWidth = Printer.Width
'           Printer.ScaleHeight = Printer.Height
'      End If
      Printer.Font.Size = 22
      Printer.Font.Name = "細明體"
      Printer.Font.Bold = True
      Printer.Font.Underline = True
      Printer.CurrentX = 5500 - (Printer.TextWidth("CF 案件結餘轉帳通知單") / 2)
      Printer.CurrentY = 500
      Printer.Print "CF 案件結餘轉帳通知單"
      Printer.Font.Size = 12
      Printer.Font.Bold = False
      Printer.Font.Underline = False
   
      Printer.CurrentX = 0
      Printer.CurrentY = 1000
      Printer.Print "製表人：" & GetPrjSalesNM(strUserNum)
      
      Printer.CurrentX = 5500 - (Printer.TextWidth(m_Title2) / 2)
      Printer.CurrentY = 1000
      Printer.Print m_Title2
      
      Printer.CurrentX = 11500 - (Printer.TextWidth("製表日期： " & ChgSQL(strTemp(2)) & " 年 " & ChgSQL(strTemp(3)) & " 月 " & ChgSQL(strTemp(4)) & " 日 "))
      Printer.CurrentY = 1000
      Printer.Print "申請國家：" & strTemp(8)
      Printer.CurrentX = 0
      Printer.CurrentY = 1300
      Printer.Print "結餘單號：" & strTemp(1)
      'add by sonia 2022/11/3 特殊出名公司案件要加印
      strCompNo = ""
      strCompName = GetSpecialComp(m_CP01, m_CP02, m_CP03, m_CP04, strCompNo, 6)
      If strCompNo = "J" Or strCompNo = "L" Then
         Printer.Font.Size = 22
         Printer.Font.Bold = True
         Printer.CurrentX = 4300
         Printer.CurrentY = 1300
         If strCompNo = "J" Then
            Printer.Print "(智權)"
         Else
            Printer.Print "(法律)"
         End If
         Printer.Font.Size = 12
         Printer.Font.Bold = False
      End If
      'end 2022/11/3
      Printer.CurrentX = 11500 - (Printer.TextWidth("製表日期： " & ChgSQL(strTemp(2)) & " 年 " & ChgSQL(strTemp(3)) & " 月 " & ChgSQL(strTemp(4)) & " 日 "))
      Printer.CurrentY = 1300
      Printer.Print "製表日期：" & ChgSQL(strTemp(2)) & " 年 " & ChgSQL(strTemp(3)) & " 月 " & ChgSQL(strTemp(4)) & " 日 "
      
      Printer.Font.Size = 22
      Printer.Font.Bold = True
      For i = 0 To 10
      '下外框
         Printer.Line (400 + i, 13800 + i)-(11000 + i, 13800 + i)
         Printer.Line (400 + i, 16100 + i)-(11000 + i, 16100 + i)
         Printer.Line (400 + i, 13800 + i)-(400 + i, 16100 + i)
         Printer.Line (11000 + i, 13800 + i)-(11000 + i, 16100 + i)
      '下內
         'modify by sonia 2023/2/16 取消 業務部副總經理
         'Printer.Line (1100 + i, 13800 + i)-(1100 + i, 16100 + i)
         'Printer.Line (2500 + i, 13800 + i)-(2500 + i, 16100 + i)
         'Printer.Line (3200 + i, 13800 + i)-(3200 + i, 16100 + i)
         'Printer.Line (4600 + i, 13800 + i)-(4600 + i, 16100 + i)
         'Printer.Line (5300 + i, 13800 + i)-(5300 + i, 16100 + i)
         'Printer.Line (6700 + i, 13800 + i)-(6700 + i, 16100 + i)
         'Printer.Line (7400 + i, 13800 + i)-(7400 + i, 16100 + i)
         'Printer.Line (8800 + i, 13800 + i)-(8800 + i, 16100 + i)
         'Printer.Line (9500 + i, 13800 + i)-(9500 + i, 16100 + i)
         Printer.Line (1100 + i, 13800 + i)-(1100 + i, 16100 + i)
         Printer.Line (3050 + i, 13800 + i)-(3050 + i, 16100 + i)
         Printer.Line (3750 + i, 13800 + i)-(3750 + i, 16100 + i)
         Printer.Line (5700 + i, 13800 + i)-(5700 + i, 16100 + i)
         Printer.Line (6400 + i, 13800 + i)-(6400 + i, 16100 + i)
         Printer.Line (8350 + i, 13800 + i)-(8350 + i, 16100 + i)
         Printer.Line (9050 + i, 13800 + i)-(9050 + i, 16100 + i)
         'end 2023/2/16
      ''上內
      '   Printer.Line (1400 + i, 2200 + i)-(4100 + i, 2200 + i)
      '   Printer.Line (5200 + i, 2200 + i)-(8100 + i, 2200 + i)
      '   Printer.Line (9300 + i, 2200 + i)-(11500 + i, 2200 + i)
      Next i
      
      Printer.Font.Size = 12
      Printer.CurrentX = 300
      Printer.CurrentY = 1900
      Printer.Print "本所案號：" & strTemp(5)
      
      Printer.CurrentX = 4300
      Printer.CurrentY = 1900
      Printer.Print "申請人：" & strTemp(6)
      Printer.CurrentX = 8300
      Printer.CurrentY = 1900
      'Printer.Print "智權人員：" & GetPrjSalesNM(strTemp(7))  'cancel by sonia 2022/6/30
      
      '2009/11/23 add by sonia
      'Modified by Lydia 2019/07/09 改成兩行
      'Printer.CurrentX = 300
      'Printer.CurrentY = 2200
      'Printer.Print "案件備註：" & m_Remark
      '2009/11/23 end
      Pub_SmartPrint "案件備註：" & m_Remark, 300, 2200, 220
      
      Printer.CurrentX = 400 + GetPrPosX(700, "總")
      Printer.CurrentY = 13900
      Printer.Print "總"
      Printer.CurrentX = 400 + GetPrPosX(700, "經")
      Printer.CurrentY = 14800
      Printer.Print "經"
      Printer.CurrentX = 400 + GetPrPosX(700, "理")
      Printer.CurrentY = 15700
      Printer.Print "理"
      
      'modify by sonia 2023/2/16 取消 業務部副總經理,以下調整位置
      Printer.CurrentX = 3050 + GetPrPosX(700, "國")    '2023/2/16 2500改為3050
      Printer.CurrentY = 13900
      Printer.Print "國"
      Printer.CurrentX = 3050 + GetPrPosX(700, "外")
      Printer.CurrentY = 14200
      Printer.Print "外"
      Printer.CurrentX = 3050 + GetPrPosX(700, "部")
      Printer.CurrentY = 14500
      Printer.Print "部"
      Printer.CurrentX = 3050 + GetPrPosX(700, "副")
      Printer.CurrentY = 14800
      Printer.Print "副"
      Printer.CurrentX = 3050 + GetPrPosX(700, "總")
      Printer.CurrentY = 15100
      Printer.Print "總"
      Printer.CurrentX = 3050 + GetPrPosX(700, "經")
      Printer.CurrentY = 15400
      Printer.Print "經"
      Printer.CurrentX = 3050 + GetPrPosX(700, "理")
      Printer.CurrentY = 15700
      Printer.Print "理"
      
      'cancel by sonia 2023/2/16 取消 業務部副總經理
      'Printer.CurrentX = 4600 + GetPrPosX(700, "業")
      'Printer.CurrentY = 13900
      'Printer.Print "業"
      'Printer.CurrentX = 4600 + GetPrPosX(700, "務")
      'Printer.CurrentY = 14200
      'Printer.Print "務"
      'Printer.CurrentX = 4600 + GetPrPosX(700, "部")
      'Printer.CurrentY = 14500
      'Printer.Print "部"
      'Printer.CurrentX = 4600 + GetPrPosX(700, "副")
      'Printer.CurrentY = 14800
      'Printer.Print "副"
      'Printer.CurrentX = 4600 + GetPrPosX(700, "總")
      'Printer.CurrentY = 15100
      'Printer.Print "總"
      'Printer.CurrentX = 4600 + GetPrPosX(700, "經")
      'Printer.CurrentY = 15400
      'Printer.Print "經"
      'Printer.CurrentX = 4600 + GetPrPosX(700, "理")
      'Printer.CurrentY = 15700
      'Printer.Print "理"
      'end 2023/2/16
      
      Printer.CurrentX = 5700 + GetPrPosX(700, "國")    '2023/2/16 6700改為5700
      Printer.CurrentY = 13900
      Printer.Print "國"
      Printer.CurrentX = 5700 + GetPrPosX(700, "外")
      Printer.CurrentY = 14200
      Printer.Print "外"
      Printer.CurrentX = 5700 + GetPrPosX(700, "部")
      Printer.CurrentY = 14500
      Printer.Print "部"
      Printer.CurrentX = 5700 + GetPrPosX(700, "單")
      Printer.CurrentY = 14800
      Printer.Print "單"
      Printer.CurrentX = 5700 + GetPrPosX(700, "位")
      Printer.CurrentY = 15100
      Printer.Print "位"
      Printer.CurrentX = 5700 + GetPrPosX(700, "主")
      Printer.CurrentY = 15400
      Printer.Print "主"
      Printer.CurrentX = 5700 + GetPrPosX(700, "管")
      Printer.CurrentY = 15700
      Printer.Print "管"
      
      Printer.CurrentX = 8350 + GetPrPosX(700, "國")    '2023/2/16 8800改為8350
      Printer.CurrentY = 13900
      Printer.Print "國"
      Printer.CurrentX = 8350 + GetPrPosX(700, "外")
      Printer.CurrentY = 14200
      Printer.Print "外"
      Printer.CurrentX = 8350 + GetPrPosX(700, "部")
      Printer.CurrentY = 14500
      Printer.Print "部"
      Printer.CurrentX = 8350 + GetPrPosX(700, "承")
      Printer.CurrentY = 15100
      Printer.Print "承"
      Printer.CurrentX = 8350 + GetPrPosX(700, "辦")
      Printer.CurrentY = 15400
      Printer.Print "辦"
      Printer.CurrentX = 8350 + GetPrPosX(700, "人")
      Printer.CurrentY = 15700
      Printer.Print "人"
   
      Printer.Font.Bold = False
      tmpY = 2600
      Printer.Font.Bold = False
      Printer.CurrentX = 300
      Printer.CurrentY = tmpY
      Printer.Print "代理人："
      Printer.CurrentX = 300 + Printer.TextWidth("代理人：") + 200
      Printer.CurrentY = tmpY
      Printer.Print strTemp(9)
      tmpY = tmpY + 300
      Printer.CurrentX = 300
      Printer.CurrentY = tmpY
      Printer.Print "代理人請款金額："
      strTemp(10) = CheckStr(.Fields("A240013"))
      If Right(strTemp(10), 1) = ";" Then
         strTemp(10) = Mid(strTemp(10), 1, Len(strTemp(10)) - 1)
      End If
      iVarCp10 = Split(strTemp(10), ";")
      Dim oPrintStr As String
      oPrintStr = ""
      For iCountCp10 = 0 To UBound(iVarCp10) Step 3
         oPrintStr = ""
         Printer.CurrentX = 300 + Printer.TextWidth("代理人請款金額：") + 200
         Printer.CurrentY = tmpY
         If (iCountCp10) <= (UBound(iVarCp10) + 1) Then
            If (UBound(iVarCp10) + 1) - (iCountCp10 + 1) >= 0 Then
               oPrintStr = oPrintStr & iVarCp10(iCountCp10) & ";"
            End If
         End If
         If (iCountCp10 + 1) <= (UBound(iVarCp10) + 1) Then
            If (UBound(iVarCp10) + 1) - (iCountCp10 + 2) >= 0 Then
               oPrintStr = oPrintStr & iVarCp10(iCountCp10 + 1) & ";"
            End If
         End If
         If (iCountCp10 + 2) <= (UBound(iVarCp10) + 1) Then
            If (UBound(iVarCp10) + 1) - (iCountCp10 + 3) >= 0 Then
               oPrintStr = oPrintStr & iVarCp10(iCountCp10 + 2) & ";"
            End If
         End If
         Printer.Print oPrintStr
         Printer.Font.Bold = True
         Printer.Font.Bold = False
         tmpY = tmpY + 300
      Next iCountCp10
   
      'add by sonia 2023/7/27 CFT、CFC及S案加印承辦人
      If m_CP01 = "CFT" Or m_CP01 = "CFC" Or m_CP01 = "S" Then
         Call GetNA69("", GetPrjNation1(strTemp(5)), strTemp(7), NA69Emp, m_CP01, m_CP02, m_CP03, m_CP04)
         Printer.CurrentX = 300
         Printer.CurrentY = tmpY + 600
         Printer.Print "承辦人：" & GetStaffName(NA69Emp)
      End If
      'end 2023/7/27
      
      Printer.Font.Bold = True
      'Printer.Line (1200 + i, tmpY + 250 + i)-(11500 + i, tmpY + 250 + i)
      tmpY = tmpY + 600
      Printer.CurrentX = 1300 + GetPrPosX(1500, "實際收款金額")
      Printer.CurrentY = tmpY + 300
      Printer.Print "實際收款金額"
      Printer.CurrentX = 2850 + GetPrPosX(400, "減")
      Printer.CurrentY = tmpY + 300
      Printer.Print "減"
      Printer.CurrentX = 3300 + GetPrPosX(1500, "已作收入金額")
      Printer.CurrentY = tmpY + 300
      Printer.Print "已作收入金額"
      Printer.CurrentX = 4850 + GetPrPosX(400, "減")
      Printer.CurrentY = tmpY + 300
      Printer.Print "減"
      Printer.CurrentX = 5300 + GetPrPosX(1500, "實際支出費用")
      Printer.CurrentY = tmpY + 300
      Printer.Print "實際支出費用"
      Printer.CurrentX = 6850 + GetPrPosX(600, "等於")
      Printer.CurrentY = tmpY + 300
      Printer.Print "等於"
      Printer.CurrentX = 7500 + GetPrPosX(1500, "結 轉 應 付")
      Printer.CurrentY = tmpY
      Printer.Print "結 轉 應 付"
      Printer.CurrentX = 7500 + GetPrPosX(1500, "規費安全基金")
      Printer.CurrentY = tmpY + 300
      Printer.Print "規費安全基金"
      Printer.CurrentX = 9050 + GetPrPosX(400, "加")
      Printer.CurrentY = tmpY + 300
      Printer.Print "加"
      Printer.CurrentX = 9500 + GetPrPosX(1500, "結轉收入金額")
      Printer.CurrentY = tmpY + 300
      Printer.Print "結轉收入金額"
      tmpY = tmpY + 600
      Printer.Font.Bold = False
      Printer.CurrentX = 0
      Printer.CurrentY = tmpY
      Printer.Print String(104, "-")
   End With
   CheckOC
   IsPrintok = True
   Printer.EndDoc
   'add by nickc 2008/04/08
   If Printer.ScaleMode = 0 Then Printer.ScaleMode = 1
End Sub

Sub PrintByOneOld(oA240002 As String)
Dim returnRec As Long
Dim TmpR030006 As String
Dim tRS As New ADODB.Recordset
Dim tRS2 As New ADODB.Recordset
Dim ThisKey As String
Dim oCP01 As String
Dim oCP02 As String
   
   strSql = "SELECT a240002,a240005,a240006 FROM ACC240 WHERE a240002='" & oA240002 & "' "
   Set tRS = New ADODB.Recordset
   If tRS.State = 1 Then tRS.Close
   tRS.CursorLocation = adUseClient
   tRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If tRS.RecordCount <> 0 Then
      ThisKey = CheckStr(tRS.Fields(0))
      oCP01 = CheckStr(tRS.Fields(1))
      oCP02 = CheckStr(tRS.Fields(2))
   Else
       MsgBox "無此結餘單號！", vbCritical, "資料錯誤！"
       Exit Sub
   End If
   
   '判斷新案還是舊案 2011/3/10 移至共用
   IsOldSystem = Judgecase(oCP01, oCP02)
   
   If IsOldSystem = True Then
      PrintData2_SpaceOld ThisKey
   Else
      PrintData2_1Old ThisKey
   End If
   'Screen.MousePointer = vbDefault
End Sub

'add by nickc 2006/02/27
'報表列印
Private Sub PrintData()
   
   Dim ii As Integer
   Dim SeekPrintKind As String
   Printer.Orientation = 2
   strTemp(0) = "'"
   strTemp(4) = ""
   strTemp(5) = ""
   strTemp(6) = ""
   strTemp(10) = ""
   strTemp(11) = ""
   GetPleft
   With grd2
      Page = 1
      SeekPrintKind = SystemNumber(.TextMatrix(1, 12), 1)

      PrintTitle SeekPrintKind
      For ii = 1 To .Rows - 1
         If SeekPrintKind <> SystemNumber(.TextMatrix(ii, 12), 1) Then
            Printer.Font.Size = 12
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            Printer.Print String(200, "-")
            Printer.Font.Size = 10
            Printer.NewPage
            Page = Page + 1
            SeekPrintKind = SystemNumber(.TextMatrix(ii, 12), 1)
            PrintTitle SeekPrintKind
         End If
         Printer.Font.Size = 10
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print IIf(strTemp(4) = .TextMatrix(ii, 4) And strTemp(0) = .TextMatrix(ii, 0), "", .TextMatrix(ii, 0))
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 1)
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iPrint
         'edit by nickc 2006/03/01
         'Printer.Print StrToStr(.TextMatrix(ii, 2), 4)
         'Printer.CurrentX = PLeft(3)
         'Printer.CurrentY = iPrint
         'Printer.Print StrToStr(.TextMatrix(ii, 3), 6)
         Printer.Print StrToStr(.TextMatrix(ii, 2), 10)
         Printer.CurrentX = PLeft(4)
         Printer.CurrentY = iPrint
         Printer.Print IIf(strTemp(4) = .TextMatrix(ii, 4), "", .TextMatrix(ii, 4))
         Printer.CurrentX = PLeft(5)
         Printer.CurrentY = iPrint
         Printer.Print IIf(strTemp(4) = .TextMatrix(ii, 4) And strTemp(5) = .TextMatrix(ii, 5), "", .TextMatrix(ii, 5))
         Printer.CurrentX = PLeft(6)
         Printer.CurrentY = iPrint
         Printer.Print IIf(strTemp(4) = .TextMatrix(ii, 4) And strTemp(6) = .TextMatrix(ii, 6), "", .TextMatrix(ii, 6))
         Printer.CurrentX = PLeft(7)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 7)
         Printer.CurrentX = PLeft(8) + 500 - Printer.TextWidth(Format(.TextMatrix(ii, 8), "0.0"))
         Printer.CurrentY = iPrint
         Printer.Print Format(.TextMatrix(ii, 8), "0.0")
         Printer.CurrentX = PLeft(9)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(ii, 9)
         Printer.CurrentX = PLeft(10)
         Printer.CurrentY = iPrint
         Printer.Print IIf(strTemp(4) = .TextMatrix(ii, 4) And strTemp(10) = StrToStr(.TextMatrix(ii, 10), 15), "", StrToStr(.TextMatrix(ii, 10), 15))
         Printer.CurrentX = PLeft(11)
         Printer.CurrentY = iPrint
         Printer.Print IIf(strTemp(4) = .TextMatrix(ii, 4) And strTemp(11) = .TextMatrix(ii, 11), "", .TextMatrix(ii, 11))
         If strTemp(4) <> .TextMatrix(ii, 4) Then
            strTemp(0) = .TextMatrix(ii, 0)
            strTemp(4) = .TextMatrix(ii, 4)
            strTemp(5) = .TextMatrix(ii, 5)
            strTemp(6) = .TextMatrix(ii, 6)
            strTemp(10) = StrToStr(.TextMatrix(ii, 10), 15)
            strTemp(11) = .TextMatrix(ii, 11)
         End If
         iPrint = iPrint + 300
         If iPrint > 10000 And ii <> .Rows - 1 Then
            If SeekPrintKind = SystemNumber(.TextMatrix(ii + 1, 12), 1) Then
               Printer.Font.Size = 12
               Printer.CurrentX = 500
               Printer.CurrentY = iPrint
               Printer.Print String(200, "-")
               Printer.NewPage
               Page = Page + 1
               PrintTitle SeekPrintKind
            End If
         End If
      Next ii
   End With
   Printer.Font.Size = 12
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   Printer.EndDoc
   '2011/11/15 MODIFY BY SONIA 否則會顯示二次列印完成
   'ShowPrintOk
   s = MsgBox("結餘案件無法結算明細表列印完成!!", , "列印成功")
End Sub

'add by nickc 2006/02/27
Sub GetPleft()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = PLeft(0) + 1000
   PLeft(2) = PLeft(1) + 1600
   PLeft(3) = PLeft(2) + 950
   PLeft(4) = PLeft(3) + 1300 - 300
   PLeft(5) = PLeft(4) + 1150
   PLeft(6) = PLeft(5) + 2000
   PLeft(7) = PLeft(6) + 1000
   PLeft(8) = PLeft(7) + 800
   PLeft(9) = PLeft(8) + 650
   PLeft(10) = PLeft(9) + 1200
   PLeft(11) = PLeft(10) + 2800 + 300
End Sub

'add by nickc 2006/02/27
Sub PrintTitle(oClass As String)
   GetPleft
   
   iPrint = 500
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6000
   Printer.CurrentY = iPrint
   'edit by nickc 2006/03/01 婧瑄修改
   'Printer.Print "結餘案件尚有未結匯明細表"
   Printer.Print "結餘案件尚無法結算明細表"

   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & GetPrjSalesNM(strUserNum)
   Printer.CurrentX = 7500
   Printer.CurrentY = iPrint
   Printer.Print "系統別：" & oClass
   Printer.CurrentX = 13500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")

   iPrint = iPrint + 300
   Printer.CurrentX = 13500
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(Page)

   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.Font.Size = 10
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "輸入日期"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   'edit by nickc 2006/03/01 婧瑄更改
   'Printer.Print "申請國家"
   'Printer.CurrentX = PLeft(3)
   'Printer.CurrentY = iPrint
   'Printer.Print "案件性質"
   Printer.Print "無法結案原因"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "帳單編號"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "代理人 D/N No."
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "帳單日期"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "幣別"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iPrint
   Printer.Print "金額"
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iPrint
   Printer.Print "代理人編號"
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = iPrint
   Printer.Print "代理人"
   Printer.CurrentX = PLeft(11)
   Printer.CurrentY = iPrint
   Printer.Print "輸入人員"
   iPrint = iPrint + 300
   Printer.Font.Size = 12
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   Printer.Font.Size = 10
   iPrint = iPrint + 300
   
End Sub

'抓代理人代碼 從 CP 來 add by nickc 2006/04/26
Public Function GetPrjFagentNumByCPNot001(ByRef Strindex As String) As String            '抓代理人代碼
   If UCase(Left(Strindex, 1)) = "N" Then
       Strindex = Right(Strindex, Len(Strindex) - 1)
   End If
   'edit by nickc
   'strSQL = "SELECT cp44 FROM  caseprogress WHERE cp09 in (select max(cp09) from caseprogress where cp01='" & SystemNumber(Strindex, 1) & "'  AND cp02='" & SystemNumber(Strindex, 2) & "' AND cp03='" & SystemNumber(Strindex, 3) & "' AND cp04='" & SystemNumber(Strindex, 4) & "' and cp44 is not null and cp10>'001' ) "
   strSql = "SELECT cp44 FROM  caseprogress WHERE cp09 in (select max(cp09) from caseprogress where cp01='" & SystemNumber(Strindex, 1) & "'  AND cp02='" & SystemNumber(Strindex, 2) & "' AND cp03='" & SystemNumber(Strindex, 3) & "' AND cp04='" & SystemNumber(Strindex, 4) & "' and cp44 is not null and cp10>'001' and cp27 in (select max(cp27) from caseprogress where cp01='" & SystemNumber(Strindex, 1) & "'  AND cp02='" & SystemNumber(Strindex, 2) & "' AND cp03='" & SystemNumber(Strindex, 3) & "' AND cp04='" & SystemNumber(Strindex, 4) & "' and cp44 is not null and cp10>'001' )) "
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
       If Not IsNull(AdoRecordSet3.Fields(0)) Then
           GetPrjFagentNumByCPNot001 = AdoRecordSet3.Fields(0)
       Else
           GetPrjFagentNumByCPNot001 = ""
       End If
   Else
       GetPrjFagentNumByCPNot001 = ""
   End If
   CheckOC3
End Function

'add by sonia 2025/6/27 加印當天自動作廢之結餘單、J公司銷項稅額為0為結餘單清單
Private Sub PrintData1()
   
   Do While Not AdoRecordSet3.EOF
      tmpY = tmpY + 300
      Printer.CurrentX = 50
      Printer.CurrentY = tmpY
      Printer.Print "    " & CheckStr(AdoRecordSet3.Fields(0))
      AdoRecordSet3.MoveNext
   Loop
   Printer.EndDoc
End Sub


