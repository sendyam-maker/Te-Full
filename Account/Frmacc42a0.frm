VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc42a0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "專業點數分析查詢與列印"
   ClientHeight    =   4320
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   5130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5130
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1185
      MaxLength       =   1
      TabIndex        =   12
      Top             =   3300
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1185
      MaxLength       =   1
      TabIndex        =   11
      Top             =   3570
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1185
      Style           =   2  '單純下拉式
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.ComboBox CboCmp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1215
      TabIndex        =   9
      Top             =   150
      Width           =   3430
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   300
      TabIndex        =   8
      Top             =   510
      Width           =   4350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Excel"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1380
      TabIndex        =   3
      Top             =   2760
      Width           =   2070
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3300
      TabIndex        =   2
      Top             =   3330
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtAccNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1275
      TabIndex        =   1
      Top             =   1140
      Width           =   1035
   End
   Begin VB.TextBox txtAccName 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2325
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1140
      Width           =   1932
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1275
      TabIndex        =   0
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "label7"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1100
      Left            =   210
      TabIndex        =   18
      Top             =   1560
      Width           =   9495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "顯示方式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   17
      Top             =   3330
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1:查詢 2:報表)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1590
      TabIndex        =   16
      Top             =   3330
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "分析內容"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   15
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1:明細 2:統計)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1590
      TabIndex        =   14
      Top             =   3600
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "印表機"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   13
      Top             =   3990
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   6
      Top             =   1140
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   3145
      Left            =   120
      Top             =   90
      Width           =   4905
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "傳票年月"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc42a0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/05 Form2.0已修改 (無需修改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

Dim lngThisMonth As Long '當月
Dim lngYear As Long '年
Dim lngMonth As Long '月

Dim lngDate(8) As Long
Dim iRowHeight As Integer
Dim iTopMargin As Integer
Dim iLeftMargin As Integer
Dim iTBWidth As Integer
Dim iTBHeight As Integer 'Mark by Lydia 2016/02/22 設為抬頭列數
Dim iRows As Integer
Dim iColWidth(0 To 1) As Integer
Dim strCaption As String '報表名稱

Dim m_DefaultPrinter As String, m_Prn As Printer
'欄位座標
Dim PLeft(0 To 20) As Integer
Dim iCharWidth As Integer '一個字元寬度
Dim iXPos As Integer, iYPos As Integer '現在列印的X,Y軸位置
'Add By Sindy 2014/1/23
Dim m_bolExcel As Boolean
Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt424 As New Worksheet
Dim m_lngRow As Long
Dim m_intPage As Integer
'2014/1/23 END
'Add by Amy 2015/06/16
Dim ii As Integer
Dim strField(), intWidth()   '欄位名稱/大小
Dim intField As Integer, StartRow As Integer '起始欄位/起始列
'Dim DstrTitle() As String '變動欄位名稱
Dim IsOpenXls As Boolean 'Excel是否開啟
Dim rsNew As New ADODB.Recordset
'Added by Lydia 2016/02/22
Private Const cRows As Integer = 52 '固定表格資料列
Dim nowR As Integer '現在列印的記錄位置
'Add by Amy 2018/03/12
Dim lngLastMonth As Long '從Process搬過來 上月
Dim bolArrive As Boolean '是否產生專業達成點數
Dim bolSheet2 As Boolean '是否產生表二
Dim intTitleR As Integer
Dim stCon021 As String, stCon040 As String  'Moidfy by Amy 2019/12/19 從Process搬出來
Dim strCmp As String, strCmpN As String 'Add by Amy 2020/05/14

Private Sub cmdok_Click()
   m_bolExcel = False 'Add By Sindy 2014/1/23
   Screen.MousePointer = vbHourglass
   If DataCheck = True Then
      If txtData = "1" Then
         Process1 '明細
      Else
         Process '統計
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Combo2_Click()
    If Combo2 = "專業達成點數分佈情況(當月實際達成)" Then
        txtData = "2"
        txtData.Enabled = False
        txtOutput = "1"
        txtOutput.Enabled = False
        cmdOK.Enabled = False
    Else
        txtData.Enabled = True
        txtOutput.Enabled = True
        cmdOK.Enabled = True
    End If
End Sub

'Add By Sindy 2014/1/23
Private Sub Command1_Click()
   IsOpenXls = False 'Add by Amy 2015/06/16
   m_bolExcel = True
   'Add by Amy 2018/03/12 +bolArrive
   bolArrive = False
   strCmp = "": strCmpN = "" 'Add by Amy 2020/05/14
   Screen.MousePointer = vbHourglass
   If DataCheck = True Then
      'Modify by Amy 2019/12/19 原:Process(原程式一直重抓傳票檔很慢,故改寫法) 10809 411101 因分CCP及MCP 造成兩筆資料都更新排除結餘及轉撥
      If Combo2 = "專業達成點數分佈情況(當月實際達成)" Then
            bolArrive = True '比較三年
            Call Process2
      'Add by Amy 2020/05/14
      ElseIf Combo2 = "國家別點數分析表" Then
            bolArrive = True '比較三年
            Call Process2(True)
            Call Process3
      'Mark by Amy 2020/06/05 不使用-婧瑄
'      ElseIf txtData = "1" Then
'         Process1 '明細
'      Else
'         Process '統計
      End If
   End If
   bolArrive = False
   'end 2018/03/12
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   strFormName = Name
End Sub

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, 5130, 4005 'Modify by Amy 2015/06/17 高原:3525
   '畫面初值設定
   initForm
   'Mark by Amy 2020/06/05 不使用
'   '設定印表機
'   m_DefaultPrinter = Printer.DeviceName
'   For Each m_Prn In Printers
'      cmbPrinter.AddItem m_Prn.DeviceName
'      '2013/8/16 add by sonia 瑞婷要預設1200印表機
'      If InStr(m_Prn.DeviceName, "1200") > 0 Then
'         cmbPrinter = m_Prn.DeviceName
'      End If
'      '2013/8/16 end
'   Next

   'Modfiy  by Amy 2022/01/13 +說明,把條件往下搬
   'Add by Amy 2020/06/05 拿掉 專業點數分析表  調畫面
'   Me.Height = 2800
'   Shape1.Height = 1800
'   Command1.Top = 1950
'   Command1.Left = 1140
'   Command1.Width = 3000
   'end 2020/06/05
   Me.Height = 3630
   Label7.Caption = "專業點數分析表會產三個工作表" & vbCrLf & _
                              "　　　＊＊＊　 較耗時,請耐心等候 　＊＊＊" & vbCrLf & _
                              "工作表一：會計科目餘額資料(實績+結餘)" & vbCrLf & _
                              "不含結餘及轉撥點數：工作表一扣除結餘及轉撥點數" & vbCrLf & _
                              "　　　　　　　　　　(只顯示實績)" & vbCrLf & _
                              "工作表三：專業達成點數表(秘書用+比較三年)"

   'end 2022/01/13
   'Add by Amy 2020/05/14 公司別下拉
   CboCmp.Clear
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, True, False, False, , 1)
   'end 2020/05/14
   'cmbPrinter.ListIndex = 0
   'Add by Amy 2015/06/16 +下拉 可選專業達成點數分佈情況
   'Combo2.AddItem "專業點數分析表"'Mark by Amy 2020/06/05 不使用-婧瑄
   Combo2.AddItem "專業達成點數分佈情況(當月實際達成)"
   Combo2.AddItem "國家別點數分析表" 'Add by Amy 2020/05/14
   'Modify by Amy 2018/03/12 改預設
   Combo2 = "專業達成點數分佈情況(當月實際達成)"
   Call Combo2_Click
   '2018/03/12
End Sub

Private Sub initForm()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   txtOutput.Text = "2"   '2013/8/16改預設
   txtData.Text = "2"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   '還原印表機
   For Each m_Prn In Printers
      If m_Prn.DeviceName = m_DefaultPrinter Then
         Set Printer = m_Prn
         Exit For
      End If
   Next
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc42a0 = Nothing
End Sub

Private Function DataCheck() As Boolean

   Dim stMsg As String
   
   stMsg = "傳票年月輸入錯誤！"
   If InStr(Me.MaskEdBox1, "_") > 0 Then
      MsgBox stMsg
      MaskEdBox1.SetFocus
      Exit Function
   End If
   stMsg = "請輸入顯示方式！"
   If Me.txtOutput = "" Then
      MsgBox stMsg
      txtOutput.SetFocus
      Exit Function
   End If
   stMsg = "請輸入分析內容！"
   If Me.txtData = "" Then
      MsgBox stMsg
      txtData.SetFocus
      Exit Function
   End If
   
   lngThisMonth = Val(Replace(MaskEdBox1, "/", ""))
   lngYear = lngThisMonth \ 100
   lngMonth = lngThisMonth Mod 100
   
   'Modify by Amy 2020/06/05 專業點數分析表不使用-婧瑄
'   If txtData = "1" Then
'      If txtAccNo = "" Then
'         MsgBox "分析內容輸入1(明細)時會計科目不可空白！"
'         txtAccNo.SetFocus
'         Exit Function
'      End If
'      strCaption = lngYear & "年度" & lngMonth & "月專業點數分析表(明細) -- " & txtAccNo & txtAccName
'   Else
      'Modify by Amy 2020/05/14
      If Combo2 = "專業達成點數分佈情況(當月實際達成)" Or Combo2 = "國家別點數分析表" Then
        strCaption = lngYear & "年度" & lngMonth & "月" & Combo2
      Else
        strCaption = lngYear & "年度" & lngMonth & "月專業點數分析表"
      End If
'   End If
   
   DataCheck = True
   
End Function

'Mark by Amy 2020/05/14 公司別改下拉
''Add By Sindy 2014/1/22
'Private Sub Text3_GotFocus()
'   TextInverse Text3
'End Sub
'Private Sub Text3_KeyPress(KeyAscii As Integer)
''   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
''      KeyAscii = 0
''   End If
'End Sub
'end 2020/05/14

'Add by Amy 2020/05/14
Private Sub CboCmp_GotFocus()
    TextInverse CboCmp
End Sub

Private Sub CboCmp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboCmp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(CboCmp) = MsgText(601) Then Exit Sub
    
    strCmp = CboCmp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If

    If InStr(GetBookKeepCmp & 組合作帳公司 & ",", strCmp) = 0 Then
        MsgBox Label1 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/05/14

Private Sub txtAccNo_Change()
   If txtAccNo <> "" Then
      txtAccName = A0102Query(txtAccNo)
   Else
      txtAccName = ""
   End If
End Sub

Private Sub txtAccNo_GotFocus()
   TextInverse txtAccNo
End Sub

Private Sub Txtdata_GotFocus()
   TextInverse txtData
End Sub

Private Sub txtdata_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
        KeyAscii = 0
   End If
End Sub

Private Sub txtOutput_GotFocus()
   TextInverse txtOutput
End Sub

Private Sub txtOutput_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
        KeyAscii = 0
   End If
End Sub

'Mark by Amy 2020/06/05 不使用-婧瑄
Private Sub Process1()
'    Dim strWhere As String 'Add by Amy 2020/05/14
'
'On Error GoTo ErrHnd
'    'Add by Amy 2020/05/14 公司別改下拉
'    strCmp = CboCmp
'    If Trim(strCmp) <> "" Then
'        If InStr(strCmp, "　") > 0 Then
'            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
'        End If
'        If InStr(strCmp, "+") > 0 Then
'            strWhere = " And a0201 In ('" & Replace(strCmp, "+", "','") & "') "
'        Else
'            strWhere = " And a0201='" & strCmp & "' "
'        End If
'    End If
'    strCmpN = GetAccReportCmpN(CboCmp, , True)
'    'end 2020/05/14
'
'   'Modify By Sindy 2014/1/22 +IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", ""))
'   'Modify by Amy 2020/05/14 公司別改抓變數
'   strSql = "select  a0205,ax202,ax204, ax206, ax207, ax212, ax208, ax209, ax214, ax213" & _
'      " from acc020, acc021 where a0205>=" & lngThisMonth & "01 and a0205<=" & lngThisMonth & "31" & strWhere & _
'      " and ax201(+)=a0201 and ax202(+)=a0202 and ax205='" & txtAccNo & "' and ax209 is not null order by ax202, ax203"
'
'   intI = 1
'   'edit by nickc 2007/02/07 不用 dll 了
'   'Set RsTemp = objLawDll.ReadRstMsg(intI, strSQL)
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      If txtOutput = "1" Then
'         Load Frmacc42a2
'         With Frmacc42a2
'            .DataGrid1.Caption = strCaption
'            Set .Adodc1.Recordset = RsTemp.Clone
'            Set .DataGrid1.DataSource = .Adodc1
'            .Hide
'            .Show
'         End With
'      Else
'         'Modify By Sindy 2014/1/23
'         If m_bolExcel = True Then
'            ExcelSave1 RsTemp.Clone, strCmpN 'Modify by Amy 2020/05/14 +strCmpN
'         Else
'         '2014/1/23 END
'            doPrint1 RsTemp.Clone, strCmpN 'Modify by Amy 2020/05/14 +strCmpN
'         End If
'      End If
'   Else
'      MsgBox "無符合資料！"
'   End If
'
'ErrHnd:
'   If Err.Number <> 0 Then
'      MsgBox Err.Description
'   End If
End Sub
   
'Modify By Sindy 2014/1/22 +IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", ""))
'                          +IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", ""))
'Modify by Amy 2015/06/16 拿掉(ByRef p_Rst As ADODB.Recordset)
'Mark by Amy 2020/06/05 不使用-婧瑄
Private Sub Process()
'   'Memo by Amy 2019/12/19 此處有改需確認Process2是否也改
'On Error GoTo ErrHnd
'
'   Dim stVTB As String, stVTB1 As String
'   Dim lngBefLastMonth As Long '上上月
'   'Dim stCon021 As String, stCon040 As String 'Mark by Amy 2019/12/19 搬至全域
'   Dim stCFT As String, stCFP As String 'Add by Amy 2015/06/16
'   Dim StrSQLa As String                'add by sonia 2016/2/18
'   'Dim strArrive(2) As String 'Mark by Amy 2019/12/19 不使用 'Add by Amy 2018/03/12 三年比較
'
'   stCon021 = "": stCon040 = ""
'   If lngMonth = 1 Then
'      lngLastMonth = lngThisMonth - 100 + 11
'   Else
'      lngLastMonth = lngThisMonth - 1
'   End If
'
'   If lngMonth = 2 Then
'      lngBefLastMonth = lngThisMonth - 100 + 10
'   Else
'      'modify by sonia 2016/2/19
'      'lngBefLastMonth = lngThisMonth - 2
'      lngBefLastMonth = lngLastMonth - 1
'   End If
'
'   If txtAccNo <> "" Then
'      stCon021 = " and ax205='" & txtAccNo & "'"
'      stCon040 = " and a0405='" & txtAccNo & "'"
'   End If
'
'   'Add by Amy 2020/05/14 公司別改下拉
'   strCmp = CboCmp
'   If Trim(strCmp) <> "" Then
'        If InStr(strCmp, "　") > 0 Then
'            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
'        End If
'        If InStr(strCmp, "+") > 0 Then
'            stCon021 = stCon021 & " And ax201 In ('" & Replace(strCmp, "+", "','") & "') "
'            stCon040 = stCon040 & " And a0403 In ('" & Replace(strCmp, "+", "','") & "') "
'        Else
'            stCon021 = " and ax201='" & strCmp & "'"
'            stCon040 = " and a0403='" & strCmp & "'"
'        End If
'   End If
'   'end 2020/05/14
'
'   'Modify by Amy 2019/02/18 CCT=餘額-MCT-MFCT(有代理人) 改寫至funciton--婧瑄
'   '沒有本所號的歸到最後一句 or ax214 is null
''   stVTB = "select ax205, sum(decode(floor(a0205/100)," & lngThisMonth & ",ax206)) Sd1" & _
''      ", sum(decode(floor(a0205/100)," & lngThisMonth & ",ax207)) Sc1" & _
''      ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2" & _
''      ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2" & _
''      ", sum(decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax206)) Sd3" & _
''      ", sum(decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax207)) Sc3" & _
''      ", sum(decode(floor(a0205/10000)," & lngYear & ",ax206)) Sd4" & _
''      ", sum(decode(floor(a0205/10000)," & lngYear & ",ax207)) Sc4" & _
''      ", sum(decode(floor(a0205/10000)," & lngYear - 1 & ",ax206)) Sd5" & _
''      ", sum(decode(floor(a0205/10000)," & lngYear - 1 & ",ax207)) Sc5" & strArrive(0)
''  stVTB = stVTB & " from ( select a0205, ax205, ax206, ax207 from acc020, acc021, trademark, fagent, customer" & _
''      " where ( (a0205>=" & lngYear - 1 & "0101 and a0205<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (a0205>=" & lngYear - 2 & "0101 and a0205<=" & lngThisMonth - 200 & "31)", "") & _
''               "Or (a0205>=" & lngYear & "0101 and a0205<=" & lngThisMonth & "31) )" & IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", "")) & _
''      " and ax201(+)=a0201 and ax202(+)=a0202 and ax205 IN ('410101','410104') and ax209 is not null" & stCon021 & _
''      " and tm01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and tm02(+)=substr(lpad(ax214,12,' '),4,6)" & _
''      " and tm03(+)=substr(lpad(ax214,12,' '),10,1) and tm04(+)=substr(lpad(ax214,12,' '),11,2) and tm01 is not null" & _
''      " and fa01(+)=substr(tm44,1,8) and fa02(+)=substr(tm44,9,1)" & _
''      " and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9,1) and nvl(fa10,cu10)>'010'" & _
''      " Union All" & _
''      " select a0205, ax205, ax206, ax207 from acc020, acc021, servicepractice, fagent, customer" & _
''      " where ( (a0205>=" & lngYear - 1 & "0101 and a0205<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, "Or (a0205>=" & lngYear - 2 & "0101 and a0205<=" & lngThisMonth - 200 & "31)", "") & _
''              "Or (a0205>=" & lngYear & "0101 and a0205<=" & lngThisMonth & "31) )" & IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", "")) & _
''      " and ax201(+)=a0201 and ax202(+)=a0202 and ax205 IN ('410101','410104') and ax209 is not null" & stCon021 & _
''      " and sp01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and sp02(+)=substr(lpad(ax214,12,' '),4,6)" & _
''      " and sp03(+)=substr(lpad(ax214,12,' '),10,1) and sp04(+)=substr(lpad(ax214,12,' '),11,2) and (sp01 is not null or ax214 is null)" & _
''      " and fa01(+)=substr(sp26,1,8) and fa02(+)=substr(sp26,9,1)" & _
''      " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9,1) and nvl(fa10,cu10)>'010' ) x GROUP BY ax205"
'   'Modify by Amy 2019/12/19 +"'410101','410104'"
'   stVTB = GetCCT(0, "'410101','410104'", False, stCon021)
'   'end 2019/02/18
'
'   '餘額
'   'Modify by Amy 2019/12/19 改寫至function
''   stVTB1 = "select a0405, sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
''      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
''      ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
''      ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
''      ", sum(decode(a0401," & lngYear - 1 & ",a0408)) net5" & strArrive(1) & _
''      " from acc040 where a0405 IN ('410101','410104')" & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
''      " and (a0401=" & lngYear - 1 & IIf(bolArrive = True, " Or a0401= " & lngYear - 2, "") & " or a0401=" & lngYear & ") and a0402<=" & lngMonth & _
''      " and a0404='TOT' group by a0405"
'   stVTB1 = GetACC040("'410101','410104'", stCon040)
'
'
'   'Modify by Amy 2015/06/16 +strUserNum for 原使用離線資料集物件,改寫入暫存檔
'   'Modify by Amy 2018/03/12 a0102||'-'||'台灣 拿掉台灣字樣
'   strSql = "select '" & strUserNum & "',DECODE(a0101,'410101','110','410104','140') RID, a0101, a0102 C00" & _
'      ", net1-nvl(decode(a0103,'1',(Sd1-Sc1),(Sc1-Sd1)),0) C01, net2-nvl(decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)),0) C02" & _
'      ", net3-nvl(decode(a0103,'1',(Sd3-Sc3),(Sc3-Sd3)),0) C03, net4-nvl(decode(a0103,'1',(Sd4-Sc4),(Sc4-Sd4)),0) C04" & _
'      ", net5-nvl(decode(a0103,'1',(Sd5-Sc5),(Sc5-Sd5)),0) C05" & _
'      " from acc010, (" & stVTB1 & ") w, (" & stVTB & ") y where a0101 IN ('410101','410104')" & _
'      " and a0405(+)=a0101 and ax205(+)=a0101"
'
'   'Modify by Amy 2019/02/18 MCT改抓有代理人且客戶國籍為大陸,增加MFCT 原:stVTB
'   'Modify by Amy 2019/05/02 會計名稱全型英文字部分改半型 ex:ＭＣＴ
'   'Modify by Amy 2019/12/19 +"'410101','410104'"
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "',DECODE(a0101,'410101','111','410104','141') RID , a0101, a0102||'-'||'MCT'" & _
'      ", decode(a0103,'1',(Sd1-Sc1),(Sc1-Sd1)) C01, decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)) C02" & _
'      ", decode(a0103,'1',(Sd3-Sc3),(Sc3-Sd3)) C03, decode(a0103,'1',(Sd4-Sc4),(Sc4-Sd4)) C04" & _
'      ", decode(a0103,'1',(Sd5-Sc5),(Sc5-Sd5)) C05" & _
'      " from acc010, (" & GetCCT(1, "'410101','410104'", False, stCon021) & " ) y where a0101 IN ('410101','410104') and ax205(+)=a0101"
'   'MFCT 有代理人且客戶國籍為非大陸
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "',DECODE(a0101,'410101','112','410104','142') RID , a0101, a0102||'-'||'MFCT'" & _
'      ", decode(a0103,'1',(Sd1-Sc1),(Sc1-Sd1)) C01, decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)) C02" & _
'      ", decode(a0103,'1',(Sd3-Sc3),(Sc3-Sd3)) C03, decode(a0103,'1',(Sd4-Sc4),(Sc4-Sd4)) C04" & _
'      ", decode(a0103,'1',(Sd5-Sc5),(Sc5-Sd5)) C05" & _
'      " from acc010, (" & GetCCT(2, "'410101','410104'", False, stCon021) & " ) y where a0101 IN ('410101','410104') and ax205(+)=a0101"
'    'end 2019/12/19
'    'end 2019/05/02
'
'   'Modify by Amy 2018/03/12 +bolArrive = False 專業達成點數分佈情況不需列上個月資料
'   'add by sonia 2016/2/18 跑一月資料要加讀前一年12月資料,因為上面語法不含
'   If lngMonth = 1 And bolArrive = False Then
'      'Modify by Amy 2019/02/18 MCT有代理人且客戶國籍是大陸,加MFCT-婧瑄
''      stVTB = "select ax205, sum(decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2" & _
''         ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2" & _
''         " from ( select a0205, ax205, ax206, ax207 from acc020, acc021, trademark, fagent, customer" & _
''         " where a0205>=" & lngLastMonth & "01 and a0205<=" & lngLastMonth & "31" & IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", "")) & _
''         " and ax201(+)=a0201 and ax202(+)=a0202 and ax205 IN ('410101','410104') and ax209 is not null" & stCon021 & _
''         " and tm01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and tm02(+)=substr(lpad(ax214,12,' '),4,6)" & _
''         " and tm03(+)=substr(lpad(ax214,12,' '),10,1) and tm04(+)=substr(lpad(ax214,12,' '),11,2) and tm01 is not null" & _
''         " and fa01(+)=substr(tm44,1,8) and fa02(+)=substr(tm44,9,1)" & _
''         " and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9,1) and nvl(fa10,cu10)>'010'" & _
''         " Union All" & _
''         " select a0205, ax205, ax206, ax207 from acc020, acc021, servicepractice, fagent, customer" & _
''         " where a0205>=" & lngLastMonth & "01 and a0205<=" & lngLastMonth & "31" & IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", "")) & _
''         " and ax201(+)=a0201 and ax202(+)=a0202 and ax205 IN ('410101','410104') and ax209 is not null" & stCon021 & _
''         " and sp01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and sp02(+)=substr(lpad(ax214,12,' '),4,6)" & _
''         " and sp03(+)=substr(lpad(ax214,12,' '),10,1) and sp04(+)=substr(lpad(ax214,12,' '),11,2) and (sp01 is not null or ax214 is null)" & _
''         " and fa01(+)=substr(sp26,1,8) and fa02(+)=substr(sp26,9,1)" & _
''         " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9,1) and nvl(fa10,cu10)>'010' ) x GROUP BY ax205"
'      stVTB = GetCCT(0, "'410101','410104'", True, stCon021)
'      'end 2019/02/18
'      'Modify by Amy 2020/05/14 公司別判斷改至stCon040 判斷 原:IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", ""))
'      stVTB1 = "select a0405, sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
'         " from acc040 where a0405 IN ('410101','410104')" & stCon040 & _
'         " and a0401=" & lngYear - 1 & " and a0402=12 " & _
'         " and a0404='TOT' group by a0405"
'
'      'Modify by Amy 2018/03/12 a0102||'-'||'台灣 拿掉台灣字樣
'      StrSQLa = " select '" & strUserNum & "',DECODE(a0101,'410101','110','410104','140') RID, a0101, a0102 C00" & _
'         ", net2-nvl(decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)),0) C02" & _
'         " from acc010, (" & stVTB1 & ") w, (" & stVTB & ") y where a0101 IN ('410101','410104')" & _
'         " and a0405(+)=a0101 and ax205(+)=a0101"
'
'      'Modify by Amy 2019/02/18 MCT改抓有代理人且客戶國籍為大陸,增加MFCT 原:stVTB
'      'Modify by Amy 2019/05/02 會計名稱全型英文字部分改半型 ex:ＭＣＴ
'      StrSQLa = StrSQLa & " Union All" & _
'         " select '" & strUserNum & "',DECODE(a0101,'410101','111','410104','141') RID , a0101, a0102||'-'||'MCT'" & _
'         ", decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)) C02" & _
'         " from acc010, (" & GetCCT(1, "'410101','410104'", True, stCon021) & " ) y where a0101 IN ('410101','410104') and ax205(+)=a0101"
'      'MFCT 有代理人且客戶國籍為非大陸
'      StrSQLa = StrSQLa & " Union All" & _
'         " select '" & strUserNum & "',DECODE(a0101,'410101','112','410104','142') RID , a0101, a0102||'-'||'MFCT'" & _
'         ", decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)) C02" & _
'         " from acc010, (" & GetCCT(2, "'410101','410104'", True, stCon021) & " ) y where a0101 IN ('410101','410104') and ax205(+)=a0101"
'      'end 2019/05/02
'   End If
''end 2016/2/18
'
'   'modify by sonia 2016/2/18 +410110
'   'Modfiy by Amy 2019/12/19 改至function
''   stVTB1 = "select a0405, sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
''      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
''      ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
''      ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
''      ", sum(decode(a0401," & lngYear - 1 & ",a0408)) net5" & strArrive(1) & _
''      " from acc040 where a0405 IN ('410102','410103','410109','410105','410106','410107','410108','417202','410110')" & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
''      " and (a0401=" & lngYear - 1 & IIf(bolArrive = True, " Or a0401=" & lngYear - 2, "") & " or a0401=" & lngYear & ") and a0402<=" & lngMonth & _
''      " and a0404='TOT' group by a0405"
'   stVTB1 = GetACC040("'410102','410103','410109','410105','410106','410107','410108','417202','410110'", stCon040)
'
'   'Modify by Morgan 2010/6/3 +410109
'   'Modify by Morgan 2011/1/11 a0102||decode(a0101,'410103','(含FMT)')-->a0102
'   'modify by sonia 2016/2/18 +410110
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "',decode(a0101,'410102','12','410103','13','410109','131','410105','15','410106','16','410107','17','410108','18','417202','19','410110','181')" & _
'      " , a0101, a0102, net1, net2, net3, net4, net5" & IIf(bolArrive = True, ",net6,net7", "") & _
'      " from acc010, (" & stVTB1 & ") w where a0101 in ('410102','410103','410109','410105','410106','410107','410108','417202','410110')" & _
'      " and a0405(+)=a0101"
'
''Modify by Amy 2018/03/12 +bolArrive = False 專業達成點數分佈情況不需列上個月資料
''add by sonia 2016/2/18 跑一月資料要加讀前一年12月資料,因為上面語法不含
'   If lngMonth = 1 And bolArrive = False Then
'      'Modify by Amy 2020/05/14 公司別判斷改至stCon040 判斷 原:IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", ""))
'      stVTB1 = "select a0405, sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
'         " from acc040 where a0405 IN ('410102','410103','410109','410105','410106','410107','410108','417202','410110')" & stCon040 & _
'         " and a0401=" & lngYear - 1 & " and a0402=12 " & _
'         " and a0404='TOT' group by a0405"
'
'      StrSQLa = StrSQLa & " Union All" & _
'         " select '" & strUserNum & "',decode(a0101,'410102','12','410103','13','410109','131','410105','15','410106','16','410107','17','410108','18','417202','19','410110','181')" & _
'         " , a0101, a0102, net2" & _
'         " from acc010, (" & stVTB1 & ") w where a0101 in ('410102','410103','410109','410105','410106','410107','410108','417202','410110')" & _
'         " and a0405(+)=a0101"
'   End If
''end 2016/2/18
'
'   strSql = strSql & " Union All select '" & strUserNum & "','1z', null, '商標國內專業合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
'
'   'Modify by Morgan 2011/2/10 +012韓國也要單獨統計
'   '沒有本所號的歸到最後一句
'   'Modify by Amy 2020/05/14 公司別判斷改至stCon021 原: IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", ""))
'   stVTB = "select ax205, RID, sum(decode(floor(a0205/100)," & lngThisMonth & ",ax206)) Sd1" & _
'      ", sum(decode(floor(a0205/100)," & lngThisMonth & ",ax207)) Sc1" & _
'      ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2" & _
'      ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2" & _
'      ", sum(decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax206)) Sd3" & _
'      ", sum(decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax207)) Sc3" & _
'      ", sum(decode(floor(a0205/10000)," & lngYear & ",ax206)) Sd4" & _
'      ", sum(decode(floor(a0205/10000)," & lngYear & ",ax207)) Sc4" & _
'      ", sum(decode(floor(a0205/10000)," & lngYear - 1 & ",ax206)) Sd5" & _
'      ", sum(decode(floor(a0205/10000)," & lngYear - 1 & ",ax207)) Sc5" & _
'      " from (select a0205, ax205, decode(substr(nvl(fa10,cu10),1,3),'101','21','011','22','012','23',decode(substr(nvl(fa10,cu10),1,1),'2','24','25')) RID" & _
'      " , ax206, ax207 from acc020, acc021, trademark, fagent, customer" & _
'      " where ((a0205>=" & lngYear - 1 & "0101 and a0205<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (a0205>=" & lngYear - 2 & "0101 and a0205<=" & lngThisMonth - 200 & "31) ", "") & _
'              "or (a0205>=" & lngYear & "0101 and a0205<=" & lngThisMonth & "31))" & _
'      " and ax201(+)=a0201 and ax202(+)=a0202 and ax205='417201'" & stCon021 & _
'      " and tm01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and tm02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'      " and tm03(+)=substr(lpad(ax214,12,' '),10,1) and tm04(+)=substr(lpad(ax214,12,' '),11,2) and tm01 is not null" & _
'      " and fa01(+)=substr(tm44,1,8) and fa02(+)=substr(tm44,9,1)" & _
'      " and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9,1)"
'   stVTB = stVTB & " Union All" & _
'      " select a0205, ax205, decode(substr(nvl(fa10,cu10),1,3),'101','21','011','22','012','23',decode(substr(nvl(fa10,cu10),1,1),'2','24','25')) RID" & _
'      " , ax206, ax207 from acc020, acc021, servicepractice, fagent, customer" & _
'      " where ((a0205>=" & lngYear - 1 & "0101 and a0205<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (a0205>=" & lngYear - 2 & "0101 and a0205<=" & lngThisMonth - 200 & "31) ", "") & _
'              "or (a0205>=" & lngYear & "0101 and a0205<=" & lngThisMonth & "31))" & _
'      " and ax201(+)=a0201 and ax202(+)=a0202 and ax205='417201'" & stCon021 & _
'      " and sp01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and sp02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'      " and sp03(+)=substr(lpad(ax214,12,' '),10,1) and sp04(+)=substr(lpad(ax214,12,' '),11,2) and sp01 is not null" & _
'      " and fa01(+)=substr(sp26,1,8) and fa02(+)=substr(sp26,9,1)" & _
'      " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9,1)"
'
'   '沒有本所號的歸到最後一句 or ax214 is null
'   '2010/10/4 add by sonia 加入法務案件 D099090625 FCL010445000
'   stVTB = stVTB & " Union All" & _
'      " select a0205, ax205, decode(substr(nvl(fa10,cu10),1,3),'101','21','011','22','012','23',decode(substr(nvl(fa10,cu10),1,1),'2','24','25')) RID" & _
'      " , ax206, ax207 from acc020, acc021, lawcase, fagent, customer" & _
'      " where ((a0205>=" & lngYear - 1 & "0101 and a0205<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (a0205>=" & lngYear - 2 & "0101 and a0205<=" & lngThisMonth - 200 & "31) ", "") & _
'               "or (a0205>=" & lngYear & "0101 and a0205<=" & lngThisMonth & "31))" & _
'      " and ax201(+)=a0201 and ax202(+)=a0202 and ax205='417201'" & stCon021 & _
'      " and lc01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and lc02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'      " and lc03(+)=substr(lpad(ax214,12,' '),10,1) and lc04(+)=substr(lpad(ax214,12,' '),11,2) and (lc01 is not null or ax214 is null)" & _
'      " and fa01(+)=substr(lc22,1,8) and fa02(+)=substr(lc22,9,1)" & _
'      " and cu01(+)=substr(lc11,1,8) and cu02(+)=substr(lc11,9,1)" & _
'      " Union All select 0,'417201','21',0,0 from dual" & _
'      " Union All select 0,'417201','22',0,0 from dual" & _
'      " Union All select 0,'417201','23',0,0 from dual" & _
'      " Union All select 0,'417201','24',0,0 from dual" & _
'      " Union All select 0,'417201','25',0,0 from dual" & _
'      ") x GROUP BY ax205, RID"
'   'end 2020/05/14
'
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "',RID, a0101" & _
'      ", a0102||'-'||DECODE(RID,'21','美國','22','日本','23','韓國','24','歐洲','其他') C00" & _
'      ", decode(a0103,'1',(Sd1-Sc1),(Sc1-Sd1)) C01, decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)) C02" & _
'      ", decode(a0103,'1',(Sd3-Sc3),(Sc3-Sd3)) C03, decode(a0103,'1',(Sd4-Sc4),(Sc4-Sd4)) C04" & _
'      ", decode(a0103,'1',(Sd5-Sc5),(Sc5-Sd5)) C05" & _
'      " from acc010, (" & stVTB & ") y where a0101='417201' and ax205(+)=a0101"
'
''Modify by Amy 2018/03/12 +bolArrive = False 專業達成點數分佈情況不需列上個月資料
''add by sonia 2016/2/18 跑一月資料要加讀前一年12月資料,因為上面語法不含
'   If lngMonth = 1 And bolArrive = False Then
'      'Modify by Amy 2020/05/14 公司別判斷改至stCon021 原:IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", ""))
'      stVTB = "select ax205, RID, sum(decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2" & _
'         ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2" & _
'         " from (select a0205, ax205, decode(substr(nvl(fa10,cu10),1,3),'101','21','011','22','012','23',decode(substr(nvl(fa10,cu10),1,1),'2','24','25')) RID" & _
'         " , ax206, ax207 from acc020, acc021, trademark, fagent, customer" & _
'         " where a0205>=" & lngLastMonth & "01 and a0205<=" & lngLastMonth & "31" & _
'         " and ax201(+)=a0201 and ax202(+)=a0202 and ax205='417201'" & stCon021 & _
'         " and tm01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and tm02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'         " and tm03(+)=substr(lpad(ax214,12,' '),10,1) and tm04(+)=substr(lpad(ax214,12,' '),11,2) and tm01 is not null" & _
'         " and fa01(+)=substr(tm44,1,8) and fa02(+)=substr(tm44,9,1)" & _
'         " and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9,1)"
'      stVTB = stVTB & " Union All" & _
'         " select a0205, ax205, decode(substr(nvl(fa10,cu10),1,3),'101','21','011','22','012','23',decode(substr(nvl(fa10,cu10),1,1),'2','24','25')) RID" & _
'         " , ax206, ax207 from acc020, acc021, servicepractice, fagent, customer" & _
'         " where a0205>=" & lngLastMonth & "01 and a0205<=" & lngLastMonth & "31" & _
'         " and ax201(+)=a0201 and ax202(+)=a0202 and ax205='417201'" & stCon021 & _
'         " and sp01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and sp02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'         " and sp03(+)=substr(lpad(ax214,12,' '),10,1) and sp04(+)=substr(lpad(ax214,12,' '),11,2) and sp01 is not null" & _
'         " and fa01(+)=substr(sp26,1,8) and fa02(+)=substr(sp26,9,1)" & _
'         " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9,1)"
'
'      stVTB = stVTB & " Union All" & _
'         " select a0205, ax205, decode(substr(nvl(fa10,cu10),1,3),'101','21','011','22','012','23',decode(substr(nvl(fa10,cu10),1,1),'2','24','25')) RID" & _
'         " , ax206, ax207 from acc020, acc021, lawcase, fagent, customer" & _
'         " where a0205>=" & lngLastMonth & "01 and a0205<=" & lngLastMonth & "31" & _
'         " and ax201(+)=a0201 and ax202(+)=a0202 and ax205='417201'" & stCon021 & _
'         " and lc01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and lc02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'         " and lc03(+)=substr(lpad(ax214,12,' '),10,1) and lc04(+)=substr(lpad(ax214,12,' '),11,2) and (lc01 is not null or ax214 is null)" & _
'         " and fa01(+)=substr(lc22,1,8) and fa02(+)=substr(lc22,9,1)" & _
'         " and cu01(+)=substr(lc11,1,8) and cu02(+)=substr(lc11,9,1)" & _
'         " Union All select 0,'417201','21',0,0 from dual" & _
'         " Union All select 0,'417201','22',0,0 from dual" & _
'         " Union All select 0,'417201','23',0,0 from dual" & _
'         " Union All select 0,'417201','24',0,0 from dual" & _
'         " Union All select 0,'417201','25',0,0 from dual" & _
'         ") x GROUP BY ax205, RID"
'      'end 2020/05/14
'
'      StrSQLa = StrSQLa & " Union All" & _
'         " select '" & strUserNum & "',RID, a0101" & _
'         ", a0102||'-'||DECODE(RID,'21','美國','22','日本','23','韓國','24','歐洲','其他') C00" & _
'         ", decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)) C02" & _
'         " from acc010, (" & stVTB & ") y where a0101='417201' and ax205(+)=a0101"
'   End If
''end 2016/2/18
'
'   'add by sonia 2016/2/18 +FCT收入-法務417203
'   'Modify by Amy 2019/12/19 改寫至function
''   stVTB1 = "select a0405, sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
''      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
''      ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
''      ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
''      ", sum(decode(a0401," & lngYear - 1 & ",a0408)) net5" & strArrive(1) & _
''      " from acc040 where a0405 in ('417203')" & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
''      " and (a0401=" & lngYear - 1 & IIf(bolArrive = True, " Or a0401=" & lngYear - 2, "") & " or a0401=" & lngYear & ") and a0402<=" & lngMonth & _
''      " and a0404='TOT' group by a0405"
'   stVTB1 = GetACC040("'417203'", stCon040)
'
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "',decode(a0101,'417203','26')" & _
'      " , a0101, a0102, net1, net2, net3, net4, net5" & IIf(bolArrive = True, ",net6,net7", "") & _
'      " from acc010, (" & stVTB1 & ") w where a0101 in ('417203')" & _
'      " and a0405(+)=a0101"
'   'end 2016/2/18
'
''Modify by Amy 2018/03/12 +bolArrive = False 專業達成點數分佈情況不需列上個月資料
''add by sonia 2016/2/18 跑一月資料要加讀前一年12月資料,因為上面語法不含
'   If lngMonth = 1 And bolArrive = False Then
'      'Modify by Amy 2020/05/14 公司別判斷改至stCon040 原:IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", ""))
'      stVTB1 = "select a0405, sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
'         " from acc040 where a0405 in ('417203')" & stCon040 & _
'         " and a0401=" & lngYear - 1 & " and a0402=12 " & _
'         " and a0404='TOT' group by a0405"
'      'end 2020/05/14
'
'      StrSQLa = StrSQLa & " Union All" & _
'         " select '" & strUserNum & "',decode(a0101,'417203','26')" & _
'         " , a0101, a0102, net2" & _
'         " from acc010, (" & stVTB1 & ") w where a0101 in ('417203')" & _
'         " and a0405(+)=a0101"
'   End If
''end 2016/2/18
'
'   strSql = strSql & " Union All select '" & strUserNum & "','2z' , null, 'FCT收入合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
'
'   'Add by Amy 2015/06/16 調整 CFT(4121)
'   'modify by sonia 2016/2/18 CFT4121拆為412101CFT收入及412102CFT收入-法務
'   'stCFT = "select a0405, sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
'      ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
'      ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
'      ", sum(decode(a0401," & lngYear - 1 & ",a0408)) net5" & _
'      " from acc040 where a0405 ='4121'" & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
'      " and (a0401=" & lngYear - 1 & " or a0401=" & lngYear & ") and a0402<=" & lngMonth & _
'      " and a0404='TOT' group by a0405"
'
'   '   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "','2z1' RID, a0101, a0102, net1, net2, net3, net4,net5" & _
'      " from acc010, (" & stCFT & ") w where a0101='4121' and a0405(+)=a0101 "
'   'Modify by Amy 2019/12/19 改寫至function
''   stCFT = "select a0405, sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
''      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
''      ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
''      ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
''      ", sum(decode(a0401," & lngYear - 1 & ",a0408)) net5" & strArrive(1) & _
''      " from acc040 where a0405 in ('412101','412102')" & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
''      " and (a0401=" & lngYear - 1 & IIf(bolArrive = True, " Or a0401=" & lngYear - 2, "") & " or a0401=" & lngYear & ") and a0402<=" & lngMonth & _
''      " and a0404='TOT' group by a0405"
'    stCFT = GetACC040("'412101','412102'", stCon040)
'
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "',decode(a0101,'412101','31','412102','32') RID, a0101, a0102, net1, net2, net3, net4,net5" & IIf(bolArrive = True, ",net6,net7", "") & _
'      " from acc010, (" & stCFT & ") w where a0101 in ('412101','412102') and a0405(+)=a0101 "
'    'end 2015/06/16
'
''Modify by Amy 2018/03/12 +bolArrive = False 專業達成點數分佈情況不需列上個月資料
''add by sonia 2016/2/19 跑一月資料要加讀前一年12月資料,因為上面語法不含
'   If lngMonth = 1 And bolArrive = False Then
'      'Modify by Amy 2020/05/14 公司別判斷改至stCon040 原:IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", ""))
'      stCFT = "select a0405, sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
'         " from acc040 where a0405 in ('412101','412102')" & stCon040 & _
'         " and a0401=" & lngYear - 1 & " and a0402=12 " & _
'         " and a0404='TOT' group by a0405"
'
'      StrSQLa = StrSQLa & " Union All" & _
'         " select '" & strUserNum & "',decode(a0101,'412101','31','412102','32') RID, a0101, a0102, net2" & _
'         " from acc010, (" & stCFT & ") w where a0101 in ('412101','412102') and a0405(+)=a0101 "
'   End If
''end 2016/2/19
'
'   'add by sonia 2016/2/18 +CFT收入合計
'   strSql = strSql & " Union All select '" & strUserNum & "','3z', null, 'CFT收入合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
'   'end 2016/2/18
'
'   strSql = strSql & " Union All select '" & strUserNum & "','3zz', null, '商標達成總計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
'
'   '沒有本所號的歸到最後一句  or ax214 is null
'   'Modify by Amy 2019/02/18 非MCP=餘額-MCP-MFCP(有代理人) 改寫至funciton--婧瑄
''   stVTB = "select ax205, sum(decode(floor(a0205/100)," & lngThisMonth & ",ax206)) Sd1" & _
''      ", sum(decode(floor(a0205/100)," & lngThisMonth & ",ax207)) Sc1" & _
''      ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2" & _
''      ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2" & _
''      ", sum(decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax206)) Sd3" & _
''      ", sum(decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax207)) Sc3" & _
''      ", sum(decode(floor(a0205/10000)," & lngYear & ",ax206)) Sd4" & _
''      ", sum(decode(floor(a0205/10000)," & lngYear & ",ax207)) Sc4" & _
''      ", sum(decode(floor(a0205/10000)," & lngYear - 1 & ",ax206)) Sd5" & _
''      ", sum(decode(floor(a0205/10000)," & lngYear - 1 & ",ax207)) Sc5" & strArrive(0)
'' stVTB = stVTB & " from ( select a0205, ax205, ax206, ax207 from acc020, acc021, patent, fagent, customer" & _
''      " where ( (a0205>=" & lngYear - 1 & "0101 and a0205<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, "Or (a0205>=" & lngYear - 2 & "0101 and a0205<=" & lngThisMonth - 200 & "31) ", "") & _
''                "or (a0205>=" & lngYear & "0101 and a0205<=" & lngThisMonth & "31) )" & IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", "")) & _
''      " and ax201(+)=a0201 and ax202(+)=a0202 and ax205='411101' and ax209 is not null" & stCon021 & _
''      " and pa01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and pa02(+)=substr(lpad(ax214,12,' '),4,6)" & _
''      " and pa03(+)=substr(lpad(ax214,12,' '),10,1) and pa04(+)=substr(lpad(ax214,12,' '),11,2) and pa01 is not null" & _
''      " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9,1)" & _
''      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1) and nvl(fa10,cu10)>'010'" & _
''      " Union All" & _
''      " select a0205, ax205, ax206, ax207 from acc020, acc021, servicepractice, fagent, customer" & _
''      " where ( (a0205>=" & lngYear - 1 & "0101 and a0205<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, "Or (a0205>=" & lngYear - 2 & "0101 and a0205<=" & lngThisMonth - 200 & "31) ", "") & _
''               "or (a0205>=" & lngYear & "0101 and a0205<=" & lngThisMonth & "31) )" & IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", "")) & _
''      " and ax201(+)=a0201 and ax202(+)=a0202 and ax205='411101' and ax209 is not null" & stCon021 & _
''      " and sp01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and sp02(+)=substr(lpad(ax214,12,' '),4,6)" & _
''      " and sp03(+)=substr(lpad(ax214,12,' '),10,1) and sp04(+)=substr(lpad(ax214,12,' '),11,2) and ( sp01 is not null or ax214 is null)" & _
''      " and fa01(+)=substr(sp26,1,8) and fa02(+)=substr(sp26,9,1)" & _
''      " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9,1) and nvl(fa10,cu10)>'010' ) x GROUP BY ax205"
'   stVTB = GetCCP(0, False, stCon021)
'   'end 2019/02/18
'
'   'Modify by Amy 2019/12/19 改寫至function
''   stVTB1 = "select a0405, sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
''      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
''      ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
''      ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
''      ", sum(decode(a0401," & lngYear - 1 & ",a0408)) net5" & strArrive(1) & _
''      " from acc040 where a0405='411101'" & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
''      " and (a0401=" & lngYear - 1 & IIf(bolArrive = True, " Or a0401=" & lngYear - 2, "") & " or a0401=" & lngYear & ") and a0402<=" & lngMonth & _
''      " and a0404='TOT' group by a0405"
'   stVTB1 = GetACC040("'411101'", stCon040)
'
'   'Modify by Amy 2018/03/12 a0102||'-'||'台灣 拿掉台灣字樣
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "','410' RID, a0101, a0102 NAME" & _
'      ", net1-nvl(decode(a0103,'1',(Sd1-Sc1),(Sc1-Sd1)),0) C01, net2-nvl(decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)),0) C02" & _
'      ", net3-nvl(decode(a0103,'1',(Sd3-Sc3),(Sc3-Sd3)),0) C03, net4-nvl(decode(a0103,'1',(Sd4-Sc4),(Sc4-Sd4)),0) C04" & _
'      ", net5-nvl(decode(a0103,'1',(Sd5-Sc5),(Sc5-Sd5)),0) C05" & _
'      " from acc010, (" & stVTB1 & ") w, (" & stVTB & ") y where a0101='411101'" & _
'      " and a0405(+)=a0101 and ax205(+)=a0405"
'
'   'Modify by Amy 2019/02/18 MCP改抓有代理人且客戶國籍為大陸,增加MFCP 原:stVTB
'   'Modify by Amy 2019/05/02 會計名稱全型英文字部分改半型 ex:ＭＣＰ
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "','411' RID, a0101, a0102||'-'||'MCP'" & _
'      ", decode(a0103,'1',(Sd1-Sc1),(Sc1-Sd1)) C01, decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)) C02" & _
'      ", decode(a0103,'1',(Sd3-Sc3),(Sc3-Sd3)) C03, decode(a0103,'1',(Sd4-Sc4),(Sc4-Sd4)) C04" & _
'      ", decode(a0103,'1',(Sd5-Sc5),(Sc5-Sd5)) C05" & _
'      " from acc010, ( " & GetCCP(1, False, stCon021) & " ) y where a0101='411101' and ax205(+)=a0101"
'   'MFCP 有代理人且客戶國籍為非大陸
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "','412' RID, a0101, a0102||'-'||'MFCP'" & _
'      ", decode(a0103,'1',(Sd1-Sc1),(Sc1-Sd1)) C01, decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)) C02" & _
'      ", decode(a0103,'1',(Sd3-Sc3),(Sc3-Sd3)) C03, decode(a0103,'1',(Sd4-Sc4),(Sc4-Sd4)) C04" & _
'      ", decode(a0103,'1',(Sd5-Sc5),(Sc5-Sd5)) C05" & _
'      " from acc010, ( " & GetCCP(2, False, stCon021) & " ) y where a0101='411101' and ax205(+)=a0101"
'   'end 2019/05/02
'
''Modify by Amy 2018/03/12 +bolArrive = False 專業達成點數分佈情況不需列上個月資料
''add by sonia 2016/2/19 跑一月資料要加讀前一年12月資料,因為上面語法不含
'   If lngMonth = 1 And bolArrive = False Then
'      'Modify by Amy 2019/02/18 MCP有代理人且客戶國籍是大陸,加MFCP-婧瑄
''      stVTB = "select ax205, sum(decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2" & _
''         ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2" & _
''         " from ( select a0205, ax205, ax206, ax207 from acc020, acc021, patent, fagent, customer" & _
''         " where a0205>=" & lngLastMonth & "01 and a0205<=" & lngLastMonth & "31" & IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", "")) & _
''         " and ax201(+)=a0201 and ax202(+)=a0202 and ax205='411101' and ax209 is not null" & stCon021 & _
''         " and pa01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and pa02(+)=substr(lpad(ax214,12,' '),4,6)" & _
''         " and pa03(+)=substr(lpad(ax214,12,' '),10,1) and pa04(+)=substr(lpad(ax214,12,' '),11,2) and pa01 is not null" & _
''         " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9,1)" & _
''         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1) and nvl(fa10,cu10)>'010'" & _
''         " Union All " & _
''         " select a0205, ax205, ax206, ax207 from acc020, acc021, servicepractice, fagent, customer" & _
''         " where a0205>=" & lngLastMonth & "01 and a0205<=" & lngLastMonth & "31" & IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", "")) & _
''         " and ax201(+)=a0201 and ax202(+)=a0202 and ax205='411101' and ax209 is not null" & stCon021 & _
''         " and sp01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and sp02(+)=substr(lpad(ax214,12,' '),4,6)" & _
''         " and sp03(+)=substr(lpad(ax214,12,' '),10,1) and sp04(+)=substr(lpad(ax214,12,' '),11,2) and ( sp01 is not null or ax214 is null)" & _
''         " and fa01(+)=substr(sp26,1,8) and fa02(+)=substr(sp26,9,1)" & _
''         " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9,1) and nvl(fa10,cu10)>'010' ) x GROUP BY ax205"
'      stVTB = GetCCP(0, True, stCon021)
'      'end 2019/02/18
'
'      'Modify by Amy 2020/05/14 公司別判斷改至stCon040 原:IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", ""))
'      stVTB1 = "select a0405, sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
'         ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
'         ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
'         ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
'         ", sum(decode(a0401," & lngYear - 1 & ",a0408)) net5" & _
'         " from acc040 where a0405='411101'" & stCon040 & _
'         " and a0401=" & lngYear - 1 & " and a0402=12 " & _
'         " and a0404='TOT' group by a0405"
'
'      'Modify by Amy 2018/03/12 a0102||'-'||'台灣 拿掉台灣字樣
'      StrSQLa = StrSQLa & " Union All " & _
'         " select '" & strUserNum & "','410' RID, a0101, a0102 NAME, net2-nvl(decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)),0) C02" & _
'         " from acc010, (" & stVTB1 & ") w, (" & stVTB & ") y where a0101='411101'" & _
'         " and a0405(+)=a0101 and ax205(+)=a0405"
'
'      'Modify by Amy 2019/02/18 MCT改抓有代理人且客戶國籍為大陸,增加MFCT 原:stVTB
'      'Modify by Amy 2019/05/02 會計名稱全型英文字部分改半型 ex:ＭＣＰ
'      StrSQLa = StrSQLa & " Union All " & _
'         " select '" & strUserNum & "','411' RID, a0101, a0102||'-'||'MCP', decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)) C02" & _
'         " from acc010, ( " & GetCCP(1, True, stCon021) & " ) y where a0101='411101' and ax205(+)=a0101"
'      'MFCP 有代理人且客戶國籍為非大陸
'      StrSQLa = StrSQLa & " Union All " & _
'         " select '" & strUserNum & "','411' RID, a0101, a0102||'-'||'MFCP', decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)) C02" & _
'         " from acc010, ( " & GetCCP(2, True, stCon021) & " ) y where a0101='411101' and ax205(+)=a0101"
'      'end 2019/05/02
'   End If
''end 2016/2/19
'
'   'Modify by Morgan 2010/1/29 +411106
'   'modify by sonia 2016/2/18 +411107
'   'Modify by Amy 2019/12/19 改寫至function
''   stVTB1 = "select a0405, sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
''      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
''      ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
''      ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
''      ", sum(decode(a0401," & lngYear - 1 & ",a0408)) net5" & strArrive(1) & _
''      " from acc040 where a0405 IN ('411102','411103','411104','411105','411106','411107')" & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
''      " and (a0401=" & lngYear - 1 & IIf(bolArrive = True, " Or a0401=" & lngYear - 2, "") & " or a0401=" & lngYear & ") and a0402<=" & lngMonth & _
''      " and a0404='TOT' group by a0405"
'   stVTB1 = GetACC040("'411102','411103','411104','411105','411106','411107'", stCon040)
'
'   'modify by sonia 2016/2/18 +411107
'   strSql = strSql & " Union All " & _
'      " select '" & strUserNum & "',decode(a0101,'411102','42','411103','43','411104','44','411105','45','411106','46','411107','461')" & _
'      " , a0101, a0102, net1, net2, net3, net4, net5" & IIf(bolArrive = True, ",net6,net7", "") & _
'      " from acc010, (" & stVTB1 & ") where a0101 in ('411102','411103','411104','411105','411106','411107')" & _
'      " and a0405(+)=a0101"
'
''Modify by Amy 2018/03/12 +bolArrive = False 專業達成點數分佈情況不需列上個月資料
''add by sonia 2016/2/19 跑一月資料要加讀前一年12月資料,因為上面語法不含
'   If lngMonth = 1 And bolArrive = False Then
'      'Modify by Amy 2020/05/14 公司別判斷改至stCon040 原:IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", ""))
'      stVTB1 = "select a0405, sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
'         " from acc040 where a0405 IN ('411102','411103','411104','411105','411106','411107')" & stCon040 & _
'         " and a0401=" & lngYear - 1 & " and a0402=12 " & _
'         " and a0404='TOT' group by a0405"
'
'      StrSQLa = StrSQLa & " Union All " & _
'         " select '" & strUserNum & "',decode(a0101,'411102','42','411103','43','411104','44','411105','45','411106','46','411107','461')" & _
'         " , a0101, a0102, net2" & _
'         " from acc010, (" & stVTB1 & ") where a0101 in ('411102','411103','411104','411105','411106','411107')" & _
'         " and a0405(+)=a0101"
'   End If
''end 2016/2/19
'
'   strSql = strSql & " Union All select '" & strUserNum & "','4z', null, '專利國內專業合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
'
'   'Modify by Morgan 2011/2/10 +012韓國也要單獨統計
'   '2009/4/28 modify by sonia ax205='4171'改為substr(ax205,1,4)='4171'
'   'Modify by Morgan 2010/1/29 FMP單獨一列抓417102,其他則抓417101
'   'modify by sonia 2016/2/18 +417103 FCP收入-法務也單獨一列,故將417102也移出至下一句,substr(ax205,1,4)='4171'改回ax205='417101'
'   'modify by sonia 2016/8/1 因加417104,417105,417109但全併入417101計算,故將欄位之ax205改為DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205),條件的ax205='417101'改為DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205)='417101'
'   'Modify by Amy 2020/05/14 公司別判斷改至stCon021 原:IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", ""))
'   stVTB = " select ax205, RID, sum(decode(floor(a0205/100)," & lngThisMonth & ",ax206)) Sd1" & _
'      ", sum(decode(floor(a0205/100)," & lngThisMonth & ",ax207)) Sc1" & _
'      ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2" & _
'      ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2" & _
'      ", sum(decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax206)) Sd3" & _
'      ", sum(decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax207)) Sc3" & _
'      ", sum(decode(floor(a0205/10000)," & lngYear & ",ax206)) Sd4" & _
'      ", sum(decode(floor(a0205/10000)," & lngYear & ",ax207)) Sc4" & _
'      ", sum(decode(floor(a0205/10000)," & lngYear - 1 & ",ax206)) Sd5" & _
'      ", sum(decode(floor(a0205/10000)," & lngYear - 1 & ",ax207)) Sc5" & _
'      " from ( select a0205, DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205) ax205, decode(substr(nvl(fa10,cu10),1,3),'101','51','011','52','012','53',decode(substr(nvl(fa10,cu10),1,1),'2','54','55')) RID" & _
'      " , ax206, ax207 from acc020, acc021, patent, fagent, customer" & _
'      " where ((a0205>=" & lngYear - 1 & "0101 and a0205<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (a0205>=" & lngYear - 2 & "0101 and a0205<=" & lngThisMonth - 200 & "31) ", "") & _
'              "or (a0205>=" & lngYear & "0101 and a0205<=" & lngThisMonth & "31))" & _
'      " and ax201(+)=a0201 and ax202(+)=a0202 and DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205)='417101' and ax209 is not null" & stCon021 & _
'      " and pa01(+)=ltrim(substr(lpad(ax214,12,' '),1,3))" & _
'      " and pa02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'      " and pa03(+)=substr(lpad(ax214,12,' '),10,1)" & _
'      " and pa04(+)=substr(lpad(ax214,12,' '),11,2) and pa01 is not null" & _
'      " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9,1)" & _
'      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)"
'   stVTB = stVTB & " Union All " & _
'      " select a0205, DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205) ax205, decode(substr(nvl(fa10,cu10),1,3),'101','51','011','52','012','53',decode(substr(nvl(fa10,cu10),1,1),'2','54','55')) RID" & _
'      " , ax206, ax207 from acc020, acc021, servicepractice, fagent, customer" & _
'      "　where ((a0205>=" & lngYear - 1 & "0101 and a0205<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (a0205>=" & lngYear - 2 & "0101 and a0205<=" & lngThisMonth - 200 & "31)  ", "") & _
'                  "or (a0205>=" & lngYear & "0101 and a0205<=" & lngThisMonth & "31))" & _
'      " and ax201(+)=a0201 and ax202(+)=a0202 and DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205)='417101' and ax209 is not null" & stCon021 & _
'      " and sp01(+)=ltrim(substr(lpad(ax214,12,' '),1,3))" & _
'      " and sp02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'      " and sp03(+)=substr(lpad(ax214,12,' '),10,1)" & _
'      " and sp04(+)=substr(lpad(ax214,12,' '),11,2) and sp01 is not null" & _
'      " and fa01(+)=substr(sp26,1,8) and fa02(+)=substr(sp26,9,1)" & _
'      " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9,1)"
'
'   '沒有本所號的歸到最後一句  or ax214 is null
'   'modify by sonia 2016/8/1 因加417104,417105,417109但全併入417101計算,故將欄位之ax205改為DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205),條件的ax205='417101'改為DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205)='417101'
'   stVTB = stVTB & " Union All " & _
'      " select a0205, DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205) ax205, decode(substr(nvl(fa10,cu10),1,3),'101','51','011','52','012','53',decode(substr(nvl(fa10,cu10),1,1),'2','54','55')) RID" & _
'      " , ax206, ax207 from acc020, acc021, lawcase, fagent, customer" & _
'      "　where ((a0205>=" & lngYear - 1 & "0101 and a0205<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, "Or (a0205>=" & lngYear - 2 & "0101 and a0205<=" & lngThisMonth - 200 & "31) ", "") & _
'                 "or (a0205>=" & lngYear & "0101 and a0205<=" & lngThisMonth & "31))" & _
'      " and ax201(+)=a0201 and ax202(+)=a0202 and DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205)='417101' and ax209 is not null" & stCon021 & _
'      " and lc01(+)=ltrim(substr(lpad(ax214,12,' '),1,3))" & _
'      " and lc02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'      " and lc03(+)=substr(lpad(ax214,12,' '),10,1)" & _
'      " and lc04(+)=substr(lpad(ax214,12,' '),11,2) and (lc01 is not null or ax214 is null)" & _
'      " and fa01(+)=substr(lc22,1,8) and fa02(+)=substr(lc22,9,1)" & _
'      " and cu01(+)=substr(lc11,1,8) and cu02(+)=substr(lc11,9,1)"
'   'end 2020/05/14
'   stVTB = stVTB & _
'      " Union All select 0,'417101','51',0,0 from dual" & _
'      " Union All select 0,'417101','52',0,0 from dual" & _
'      " Union All select 0,'417101','53',0,0 from dual" & _
'      " Union All select 0,'417101','54',0,0 from dual" & _
'      " Union All select 0,'417101','55',0,0 from dual" & _
'      " ) x GROUP BY ax205, RID"
'
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "',RID, a0101" & _
'      " , a0102||'-'||DECODE(RID,'51','美國','52','日本','53','韓國','54','歐洲','55','其他') C00" & _
'      ", decode(a0103,'1',(Sd1-Sc1),(Sc1-Sd1)) C01, decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)) C02" & _
'      ", decode(a0103,'1',(Sd3-Sc3),(Sc3-Sd3)) C03, decode(a0103,'1',(Sd4-Sc4),(Sc4-Sd4)) C04" & _
'      ", decode(a0103,'1',(Sd5-Sc5),(Sc5-Sd5)) C05" & _
'      " from acc010, (" & stVTB & ") y where a0101='417101' and ax205(+)=a0101"
'
''Modify by Amy 2018/03/12 +bolArrive = False 專業達成點數分佈情況不需列上個月資料
''add by sonia 2016/2/19 跑一月資料要加讀前一年12月資料,因為上面語法不含
'   If lngMonth = 1 And bolArrive = False Then
'      'modify by sonia 2016/8/1 因加417104,417105,417109但全併入417101計算,故將欄位之ax205改為DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205),條件的ax205='417101'改為DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205)='417101'
'      'Modify by Amy 2020/05/14 公司別判斷改至stCon021 原:IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", ""))
'      stVTB = " select DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205) ax205, RID, sum(decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2" & _
'         ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2" & _
'         " from ( select a0205, ax205, decode(substr(nvl(fa10,cu10),1,3),'101','51','011','52','012','53',decode(substr(nvl(fa10,cu10),1,1),'2','54','55')) RID" & _
'         " , ax206, ax207 from acc020, acc021, patent, fagent, customer" & _
'         " where a0205>=" & lngLastMonth & "01 and a0205<=" & lngLastMonth & "31" & _
'         " and ax201(+)=a0201 and ax202(+)=a0202 and DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205)='417101' and ax209 is not null" & stCon021 & _
'         " and pa01(+)=ltrim(substr(lpad(ax214,12,' '),1,3))" & _
'         " and pa02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'         " and pa03(+)=substr(lpad(ax214,12,' '),10,1)" & _
'         " and pa04(+)=substr(lpad(ax214,12,' '),11,2) and pa01 is not null" & _
'         " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9,1)" & _
'         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)"
'      'modify by sonia 2016/8/1 因加417104,417105,417109但全併入417101計算,故將欄位之ax205改為DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205),條件的ax205='417101'改為DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205)='417101'
'      stVTB = stVTB & " Union All" & _
'         " select a0205, DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205) ax205, decode(substr(nvl(fa10,cu10),1,3),'101','51','011','52','012','53',decode(substr(nvl(fa10,cu10),1,1),'2','54','55')) RID" & _
'         " , ax206, ax207 from acc020, acc021, servicepractice, fagent, customer" & _
'         "　where a0205>=" & lngLastMonth & "01 and a0205<=" & lngLastMonth & "31" & _
'         " and ax201(+)=a0201 and ax202(+)=a0202 and DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205)='417101' and ax209 is not null" & stCon021 & _
'         " and sp01(+)=ltrim(substr(lpad(ax214,12,' '),1,3))" & _
'         " and sp02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'         " and sp03(+)=substr(lpad(ax214,12,' '),10,1)" & _
'         " and sp04(+)=substr(lpad(ax214,12,' '),11,2) and sp01 is not null" & _
'         " and fa01(+)=substr(sp26,1,8) and fa02(+)=substr(sp26,9,1)" & _
'         " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9,1)"
'
'      '沒有本所號的歸到最後一句  or ax214 is null
'      'modify by sonia 2016/8/1 因加417104,417105,417109但全併入417101計算,故將欄位之ax205改為DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205),條件的ax205='417101'改為DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205)='417101'
'      'moidfy by sonia 2017/2/8 因2016/8/1的修改,但未改GROUP BY條件,將GROUP BY ax205, RID改為
'      stVTB = stVTB & " Union All" & _
'         " select a0205, DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205) ax205, decode(substr(nvl(fa10,cu10),1,3),'101','51','011','52','012','53',decode(substr(nvl(fa10,cu10),1,1),'2','54','55')) RID" & _
'         " , ax206, ax207 from acc020, acc021, lawcase, fagent, customer" & _
'         "　where a0205>=" & lngLastMonth & "01 and a0205<=" & lngLastMonth & "31" & _
'         " and ax201(+)=a0201 and ax202(+)=a0202 and DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205)='417101' and ax209 is not null" & stCon021 & _
'         " and lc01(+)=ltrim(substr(lpad(ax214,12,' '),1,3))" & _
'         " and lc02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'         " and lc03(+)=substr(lpad(ax214,12,' '),10,1)" & _
'         " and lc04(+)=substr(lpad(ax214,12,' '),11,2) and (lc01 is not null or ax214 is null)" & _
'         " and fa01(+)=substr(lc22,1,8) and fa02(+)=substr(lc22,9,1)" & _
'         " and cu01(+)=substr(lc11,1,8) and cu02(+)=substr(lc11,9,1)"
'      stVTB = stVTB & _
'         " Union All select 0,'417101','51',0,0 from dual" & _
'         " Union All select 0,'417101','52',0,0 from dual" & _
'         " Union All select 0,'417101','53',0,0 from dual" & _
'         " Union All select 0,'417101','54',0,0 from dual" & _
'         " Union All select 0,'417101','55',0,0 from dual" & _
'         " ) x GROUP BY DECODE(AX205,'417104','417101','417105','417101','417109','417101',AX205), RID"
'      'end 2020/05/14
'
'      StrSQLa = StrSQLa & " Union All" & _
'         " select '" & strUserNum & "',RID, a0101" & _
'         " , a0102||'-'||DECODE(RID,'51','美國','52','日本','53','韓國','54','歐洲','55','其他') C00" & _
'         ", decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)) C02" & _
'         " from acc010, (" & stVTB & ") y where a0101='417101' and ax205(+)=a0101"
'   End If
''end 2016/2/19
'
'   'add by sonia 2016/2/18 +FCP收入-FMP417102及FCP收入-法務417103
''   stVTB1 = "select a0405, sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
''      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
''      ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
''      ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
''      ", sum(decode(a0401," & lngYear - 1 & ",a0408)) net5" & strArrive(1) & _
''      " from acc040 where a0405 in ('417102','417103')" & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
''      " and (a0401=" & lngYear - 1 & IIf(bolArrive = True, " Or a0401=" & lngYear - 2, "") & " or a0401=" & lngYear & ") and a0402<=" & lngMonth & _
''      " and a0404='TOT' group by a0405"
'   stVTB1 = GetACC040("'417102','417103'", stCon040)
'
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "',decode(a0101,'417102','56','417103','57')" & _
'      " , a0101, a0102, net1, net2, net3, net4, net5" & IIf(bolArrive = True, ",net6,net7", "") & _
'      " from acc010, (" & stVTB1 & ") w where a0101 in ('417102','417103')" & _
'      " and a0405(+)=a0101"
'   'end 2016/2/18
'
''Modify by Amy 2018/03/12 +bolArrive = False 專業達成點數分佈情況不需列上個月資料
''add by sonia 2016/2/19 跑一月資料要加讀前一年12月資料,因為上面語法不含
'   If lngMonth = 1 And bolArrive = False Then
'      'Modify by Amy 2020/05/14 公司別判斷改至stCon040 原:IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", ""))
'      stVTB1 = "select a0405, sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
'         " from acc040 where a0405 in ('417102','417103')" & stCon040 & _
'         " and a0401=" & lngYear - 1 & " and a0402=12 " & _
'         " and a0404='TOT' group by a0405"
'
'      StrSQLa = StrSQLa & " Union All" & _
'         " select '" & strUserNum & "',decode(a0101,'417102','56','417103','57')" & _
'         " , a0101, a0102, net2" & _
'         " from acc010, (" & stVTB1 & ") w where a0101 in ('417102','417103')" & _
'         " and a0405(+)=a0101"
'   End If
''end 2016/2/19
'
'   strSql = strSql & " Union All select '" & strUserNum & "','5z', null, 'FCP收入合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
'
'   'Add by Amy 2015/06/16 CFP收入(4131)位置
'   '加註 by sonia 2016/1/25 a0405先不改為substr(a0405,1,4),因為413102暫無數字
'   'Modify by Amy 2018/05/21 10506有413102科目
'   'Modify by Amy 2018/06/12 4131主科目不抓,因4131/413101都是CFP收入出現2筆 原:substr(a0405,1,4)='4131'
'   'Modify by Amy 2019/12/19 改寫至function
''   stCFP = "select a0405, sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
''      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
''      ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
''      ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
''      ", sum(decode(a0401," & lngYear - 1 & ",a0408)) net5" & strArrive(1) & _
''      " from acc040 where a0405 in ('413101','413102') " & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
''      " and (a0401=" & lngYear - 1 & IIf(bolArrive = True, " Or a0401=" & lngYear - 2, "") & " or a0401=" & lngYear & ") and a0402<=" & lngMonth & _
''      " and a0404='TOT' group by a0405"
'   stCFP = GetACC040("'413101','413102'", stCon040)
'
'   'Modify by Amy 2016/04/08 原:6z 造成公式會錯
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "','6z1' RID, a0101, a0102, net1, net2, net3, net4,net5" & IIf(bolArrive = True, ",net6,net7", "") & _
'      " from acc010, (" & stCFP & ") w where a0405 in ('413101','413102') and a0405(+)=a0101 "
'    'end 2018/06/12
'    'end 2018/05/21
'    'end 2015/06/16
'
''Modify by Amy 2018/03/12 +bolArrive = False 專業達成點數分佈情況不需列上個月資料
''add by sonia 2016/2/19 跑一月資料要加讀前一年12月資料,因為上面語法不含
''Modify by Amy 2018/06/12 4131主科目不抓,因4131/413101都是CFP收入出現2筆 原:a0101='413101'
'   If lngMonth = 1 And bolArrive = False Then
'      'Modify by Amy 2020/05/14 公司別判斷改至stCon040 原:IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", ""))
'      stCFP = "select a0405, sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
'         " from acc040 where a0405 in ('413101','413102') " & stCon040 & _
'         " and a0401=" & lngYear - 1 & " and a0402=12 " & _
'         " and a0404='TOT' group by a0405"
'
'      StrSQLa = StrSQLa & " Union All" & _
'         " select '" & strUserNum & "','6z' RID, a0101, a0102, net2" & _
'         " from acc010, (" & stCFP & ") w where a0405 in ('413101','413102') and a0405(+)=a0101 "
'   End If
''end 2018/06/12
''end 2016/2/19
'
'   strSql = strSql & " Union All select '" & strUserNum & "','6zz', null, '專利達成總計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
'
'   'Modify by Amy 2014/06/16 因CFT(4121)及CFP(4131)收入位置顯示調整,拿掉語法中RID 51及52
'   'Modify by Morgan 2010/1/29 取消 418101,418102
'   'Modify by Amy 2019/08/13 增加420101
'   'Modify by Amy 2019/12/19 改寫至function
''   stVTB1 = "select a0405, sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
''      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
''      ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
''      ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
''      ", sum(decode(a0401," & lngYear - 1 & ",a0408)) net5" & strArrive(1) & _
''      " from acc040 where a0405 IN ('414101','414102','415101','415102','416101','416102','420101')" & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
''      " and (a0401=" & lngYear - 1 & IIf(bolArrive = True, " Or a0401=" & lngYear - 2, "") & " or a0401=" & lngYear & ") and a0402<=" & lngMonth & _
''      " and a0404='TOT' group by a0405"
'   stVTB1 = GetACC040("'414101','414102','415101','415102','416101','416102','420101'", stCon040)
'
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "',decode(a0101,'414101','73','414102','74','415101','75','415102','76','416101','77','416102','78','420101','79')" & _
'      " , a0101, a0102, net1, net2, net3, net4,net5" & IIf(bolArrive = True, ",net6,net7", "") & _
'      " from acc010, (" & stVTB1 & ") w where a0101 in ('414101','414102','415101','415102','416101','416102','420101')" & _
'      " and a0405(+)=a0101 "
'    'end 2015/06/16
'    'end 2019/08/13
'
''Modify by Amy 2018/03/12 +bolArrive = False 專業達成點數分佈情況不需列上個月資料
''add by sonia 2016/2/19 跑一月資料要加讀前一年12月資料,因為上面語法不含
''Modify by Amy 2019/08/13 增加420101
'   If lngMonth = 1 And bolArrive = False Then
'       'Modify by Amy 2020/05/14 公司別判斷改至stCon040 原:IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", ""))
'       stVTB1 = "select a0405, sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
'          " from acc040 where a0405 IN ('414101','414102','415101','415102','416101','416102','420101')" & stCon040 & _
'          " and a0401=" & lngYear - 1 & " and a0402=12 " & _
'          " and a0404='TOT' group by a0405"
'
'       StrSQLa = StrSQLa & " Union All" & _
'          " select '" & strUserNum & "',decode(a0101,'414101','73','414102','74','415101','75','415102','76','416101','77','416102','78','420101','79')" & _
'          " , a0101, a0102, net2" & _
'          " from acc010, (" & stVTB1 & ") w where a0101 in ('414101','414102','415101','415102','416101','416102','420101')" & _
'          " and a0405(+)=a0101 "
'   End If
''end 2016/2/19
''end 2019/08/13
'
'   'Modify by Amy 2020/05/14 公司別判斷改至stCon021 原: IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", ""))
'   stVTB = "select ax205, sum(decode(floor(a0205/100)," & lngThisMonth & ",ax206)) Sd1" & _
'      ", sum(decode(floor(a0205/100)," & lngThisMonth & ",ax207)) Sc1" & _
'      ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2" & _
'      ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2" & _
'      ", sum(decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax206)) Sd3" & _
'      ", sum(decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax207)) Sc3" & _
'      ", sum(decode(floor(a0205/10000)," & lngYear & ",ax206)) Sd4" & _
'      ", sum(decode(floor(a0205/10000)," & lngYear & ",ax207)) Sc4" & _
'      ", sum(decode(floor(a0205/10000)," & lngYear - 1 & ",ax206)) Sd5" & _
'      ", sum(decode(floor(a0205/10000)," & lngYear - 1 & ",ax207)) Sc5" & _
'      " from acc020, acc021 where ( (a0205>=" & lngYear - 1 & "0101 and a0205<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, "Or (a0205>=" & lngYear - 2 & "0101 and a0205<=" & lngThisMonth - 200 & "31) ", "") & _
'          "or (a0205>=" & lngYear & "0101 and a0205<=" & lngThisMonth & "31) )" & _
'      " and ax201(+)=a0201 and ax202(+)=a0202 and ax205='7121' and ax209 is not null" & stCon021 & " group by ax205"
'
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "','7b', a0101, '其他收入'" & _
'      ", decode(a0103,'1',(Sd1-Sc1),(Sc1-Sd1)) C01, decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)) C02" & _
'      ", decode(a0103,'1',(Sd3-Sc3),(Sc3-Sd3)) C03, decode(a0103,'1',(Sd4-Sc4),(Sc4-Sd4)) C04" & _
'      ", decode(a0103,'1',(Sd5-Sc5),(Sc5-Sd5)) C05" & _
'      " from acc010, (" & stVTB & ") x where a0101='7121' and ax205(+)=a0101"
'
''Modify by Amy 2018/03/12 +bolArrive = False 專業達成點數分佈情況不需列上個月資料
''add by sonia 2016/2/19 跑一月資料要加讀前一年12月資料,因為上面語法不含
'   If lngMonth = 1 And bolArrive = False Then
'      'Modify by Amy 2020/05/14 公司別判斷改至stCon021 原:IIf(Text3 = "2", " and a0201='J'", IIf(Text3 = "1", " and a0201='1'", ""))
'      stVTB = "select ax205, sum(decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2" & _
'         ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2" & _
'         ", sum(decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax206)) Sd3" & _
'         " from acc020, acc021 where a0205>=" & lngLastMonth & "01 and a0205<=" & lngLastMonth & "31" & _
'         " and ax201(+)=a0201 and ax202(+)=a0202 and ax205='7121' and ax209 is not null" & stCon021 & " group by ax205"
'
'      StrSQLa = StrSQLa & " Union All" & _
'         " select '" & strUserNum & "','7b', a0101, '其他收入'" & _
'         ", decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)) C02" & _
'         " from acc010, (" & stVTB & ") x where a0101='7121' and ax205(+)=a0101"
'   End If
''end 2016/2/19
'
'   strSql = strSql & " Union All select '" & strUserNum & "','7zt1', null, '專業達成總計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
'
''+if 專業達成點數分佈情況不需列
''Modify by Amy 2020/05/14 公司別判斷改至stCon040 原:IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", ""))
'If bolArrive = False Then
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "','81', null, '專利收入上月保留'" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net1" & _
'      ", sum(decode(a0401*100+a0402," & lngBefLastMonth & ",a0408)) net2" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth - 100 & ",a0408)) net3" & _
'      ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 1) + 12 & ",a0408)) net4" & _
'      ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 2) + 12 & ",a0408)) net5" & _
'      " from acc040 where a0405='4191'" & stCon040 & _
'      " and ((a0401=" & lngLastMonth \ 100 & " and a0402=" & lngLastMonth Mod 100 & ")" & _
'      " or (a0401=" & lngBefLastMonth \ 100 & " and a0402=" & lngBefLastMonth Mod 100 & ")" & _
'      " or (a0401=" & lngLastMonth \ 100 - 1 & " and a0402=" & lngLastMonth Mod 100 & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402=12)" & _
'      " or (a0401=" & lngYear - 2 & " and a0402=12))" & _
'      " and a0404='P'"
'
''add by sonia 2016/2/19 跑一月資料要特別,否則上月數字會錯
'   If lngMonth = 1 Then
'      strSql = strSql & " Union All" & _
'         " select '" & strUserNum & "','82', null, '專利收入本月保留'" & _
'         ", sum(decode(a0401*100+a0402," & lngLastMonth & ",-1*a0408," & lngThisMonth & ",a0408)) net1" & _
'         ", sum(decode(a0401*100+a0402," & lngBefLastMonth & ",-1*a0408," & lngLastMonth & ",a0408)) net2" & _
'         ", sum(decode(a0401*100+a0402," & lngLastMonth - 100 & ",-1*a0408," & lngThisMonth - 100 & ",a0408)) net3" & _
'         ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 1) + 12 & ",-1*a0408,decode(a0401," & lngYear & ",a0408))) net4" & _
'         ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 2) + 12 & ",-1*a0408,decode(a0401," & lngYear - 1 & ",decode(sign(" & lngMonth & "-a0402+1),1,a0408)))) net5" & _
'         " from acc040 where a0405='4191'" & stCon040 & _
'         " and ((a0401=" & lngYear & " and a0402<=" & lngMonth & ")" & _
'         " or (a0401=" & lngYear - 1 & " and a0402<=12)" & _
'         " or (a0401=" & lngYear - 2 & " and a0402=12))" & _
'         " and a0404='P'"
'   Else
''end 2016/2/19
'      strSql = strSql & " Union All" & _
'         " select '" & strUserNum & "','82', null, '專利收入本月保留'" & _
'         ", sum(decode(a0401*100+a0402," & lngLastMonth & ",-1*a0408," & lngThisMonth & ",a0408)) net1" & _
'         ", sum(decode(a0401*100+a0402," & lngBefLastMonth & ",-1*a0408," & lngLastMonth & ",a0408)) net2" & _
'         ", sum(decode(a0401*100+a0402," & lngLastMonth - 100 & ",-1*a0408," & lngThisMonth - 100 & ",a0408)) net3" & _
'         ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 1) + 12 & ",-1*a0408,decode(a0401," & lngYear & ",a0408))) net4" & _
'         ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 2) + 12 & ",-1*a0408,decode(a0401," & lngYear - 1 & ",decode(sign(" & lngMonth & "-a0402+1),1,a0408)))) net5" & _
'         " from acc040 where a0405='4191'" & stCon040 & _
'         " and ((a0401=" & lngYear & " and a0402<=" & lngMonth & ")" & _
'         " or (a0401=" & lngYear - 1 & " and a0402<=" & lngMonth & ")" & _
'         " or (a0401=" & lngYear - 1 & " and a0402=12)" & _
'         " or (a0401=" & lngYear - 2 & " and a0402=12))" & _
'         " and a0404='P'"
'   End If
'
'   'Modify by Morgan 2011/2/10 保留合併
'   'strSql = strSql & " Union All" & _
'      " select '63', null, 'FCP上月保留'" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net1" & _
'      ", sum(decode(a0401*100+a0402," & lngBefLastMonth & ",a0408)) net2" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth - 100 & ",a0408)) net3" & _
'      ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 1) + 12 & ",a0408)) net4" & _
'      ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 2) + 12 & ",a0408)) net5" & _
'      " from acc040 where a0405='4192'" & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
'      " and ((a0401=" & lngLastMonth \ 100 & " and a0402=" & lngLastMonth Mod 100 & ")" & _
'      " or (a0401=" & lngBefLastMonth \ 100 & " and a0402=" & lngBefLastMonth Mod 100 & ")" & _
'      " or (a0401=" & lngLastMonth \ 100 - 1 & " and a0402=" & lngLastMonth Mod 100 & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402=12)" & _
'      " or (a0401=" & lngYear - 2 & " and a0402=12))" & _
'      " and a0404='FCP'"
'
'   'strSql = strSql & " Union All" & _
'      " select '64', null, 'FCP本月保留'" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",-1*a0408," & lngThisMonth & ",a0408)) net1" & _
'      ", sum(decode(a0401*100+a0402," & lngBefLastMonth & ",-1*a0408," & lngLastMonth & ",a0408)) net2" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth - 100 & ",-1*a0408," & lngThisMonth - 100 & ",a0408)) net3" & _
'      ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 1) + 12 & ",-1*a0408,decode(a0401," & lngYear & ",a0408))) net4" & _
'      ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 2) + 12 & ",-1*a0408,decode(a0401," & lngYear - 1 & ",decode(sign(" & lngMonth & "-a0402+1),1,a0408)))) net5" & _
'      " from acc040 where a0405='4192'" & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
'      " and ((a0401=" & lngYear & " and a0402<=" & lngMonth & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402<=" & lngMonth & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402=12)" & _
'      " or (a0401=" & lngYear - 2 & " and a0402=12))" & _
'      " and a0404='FCP'"
'
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "','84', null, 'FCP保留'" & _
'      ", sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
'      ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
'      ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
'      ", sum(decode(a0401," & lngYear - 1 & ",decode(sign(" & lngMonth & "-a0402+1),1,a0408))) net5" & _
'      " from acc040 where a0405='4192'" & stCon040 & _
'      " and ((a0401=" & lngYear & " and a0402<=" & lngMonth & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402<=" & lngMonth & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402=12))" & _
'      " and a0404='FCP'"
'
'   'Modify by Morgan 2011/2/10 保留合併
'   'strSql = strSql & " Union All" & _
'      " select '65', null, 'FCT上月保留'" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net1" & _
'      ", sum(decode(a0401*100+a0402," & lngBefLastMonth & ",a0408)) net2" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth - 100 & ",a0408)) net3" & _
'      ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 1) + 12 & ",a0408)) net4" & _
'      ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 2) + 12 & ",a0408)) net5" & _
'      " from acc040 where a0405='4192'" & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
'      " and ((a0401=" & lngLastMonth \ 100 & " and a0402=" & lngLastMonth Mod 100 & ")" & _
'      " or (a0401=" & lngBefLastMonth \ 100 & " and a0402=" & lngBefLastMonth Mod 100 & ")" & _
'      " or (a0401=" & lngLastMonth \ 100 - 1 & " and a0402=" & lngLastMonth Mod 100 & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402=12)" & _
'      " or (a0401=" & lngYear - 2 & " and a0402=12))" & _
'      " and a0404='FCT'"
'
'   'strSql = strSql & " Union All" & _
'      " select '66', null, 'FCT本月保留'" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",-1*a0408," & lngThisMonth & ",a0408)) net1" & _
'      ", sum(decode(a0401*100+a0402," & lngBefLastMonth & ",-1*a0408," & lngLastMonth & ",a0408)) net2" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth - 100 & ",-1*a0408," & lngThisMonth - 100 & ",a0408)) net3" & _
'      ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 1) + 12 & ",-1*a0408,decode(a0401," & lngYear & ",a0408))) net4" & _
'      ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 2) + 12 & ",-1*a0408,decode(a0401," & lngYear - 1 & ",decode(sign(" & lngMonth & "-a0402+1),1,a0408)))) net5" & _
'      " from acc040 where a0405='4192'" & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
'      " and ((a0401=" & lngYear & " and a0402<=" & lngMonth & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402<=" & lngMonth & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402=12)" & _
'      " or (a0401=" & lngYear - 2 & " and a0402=12))" & _
'      " and a0404='FCT'"
'
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "','86', null, 'FCT保留'" & _
'      ", sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
'      ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
'      ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
'      ", sum(decode(a0401," & lngYear - 1 & ",decode(sign(" & lngMonth & "-a0402+1),1,a0408))) net5" & _
'      " from acc040 where a0405='4192'" & stCon040 & _
'      " and ((a0401=" & lngYear & " and a0402<=" & lngMonth & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402<=" & lngMonth & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402=12))" & _
'      " and a0404='FCT'"
'
'   'Modify by Morgan 2011/2/10 保留合併
'   'strSql = strSql & " Union All" & _
'      " select '67', null, 'FCL上月保留'" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net1" & _
'      ", sum(decode(a0401*100+a0402," & lngBefLastMonth & ",a0408)) net2" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth - 100 & ",a0408)) net3" & _
'      ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 1) + 12 & ",a0408)) net4" & _
'      ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 2) + 12 & ",a0408)) net5" & _
'      " from acc040 where a0405='4192'" & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
'      " and ((a0401=" & lngLastMonth \ 100 & " and a0402=" & lngLastMonth Mod 100 & ")" & _
'      " or (a0401=" & lngBefLastMonth \ 100 & " and a0402=" & lngBefLastMonth Mod 100 & ")" & _
'      " or (a0401=" & lngLastMonth \ 100 - 1 & " and a0402=" & lngLastMonth Mod 100 & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402=12)" & _
'      " or (a0401=" & lngYear - 2 & " and a0402=12))" & _
'      " and a0404='FCL'"
'
'   'strSql = strSql & " Union All" & _
'      " select '68', null, 'FCL本月保留'" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",-1*a0408," & lngThisMonth & ",a0408)) net1" & _
'      ", sum(decode(a0401*100+a0402," & lngBefLastMonth & ",-1*a0408," & lngLastMonth & ",a0408)) net2" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth - 100 & ",-1*a0408," & lngThisMonth - 100 & ",a0408)) net3" & _
'      ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 1) + 12 & ",-1*a0408,decode(a0401," & lngYear & ",a0408))) net4" & _
'      ", sum(decode(a0401*100+a0402," & 100 * (lngYear - 2) + 12 & ",-1*a0408,decode(a0401," & lngYear - 1 & ",decode(sign(" & lngMonth & "-a0402+1),1,a0408)))) net5" & _
'      " from acc040 where a0405='4192'" & IIf(Text3 = "2", " and a0403='J'", IIf(Text3 = "1", " and a0403='1'", "")) & stCon040 & _
'      " and ((a0401=" & lngYear & " and a0402<=" & lngMonth & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402<=" & lngMonth & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402=12)" & _
'      " or (a0401=" & lngYear - 2 & " and a0402=12))" & _
'      " and a0404='FCL'"
'
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "','88', null, 'FCL本月保留'" & _
'      ", sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
'      ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
'      ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
'      ", sum(decode(a0401," & lngYear - 1 & ",decode(sign(" & lngMonth & "-a0402+1),1,a0408))) net5" & _
'      " from acc040 where a0405='4192'" & stCon040 & _
'      " and ((a0401=" & lngYear & " and a0402<=" & lngMonth & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402<=" & lngMonth & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402=12))" & _
'      " and a0404='FCL'"
'
'   'end 2011/2/10
'
'   '2015/2/4 ADD BY SONIA 2015年新增4194結餘保留科目,只抓部門TOT
'   strSql = strSql & " Union All" & _
'      " select '" & strUserNum & "','89', null, '收入－結餘保留'" & _
'      ", sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
'      ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
'      ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
'      ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
'      ", sum(decode(a0401," & lngYear - 1 & ",decode(sign(" & lngMonth & "-a0402+1),1,a0408))) net5" & _
'      " from acc040 where a0405='4194'" & stCon040 & _
'      " and ((a0401=" & lngYear & " and a0402<=" & lngMonth & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402<=" & lngMonth & ")" & _
'      " or (a0401=" & lngYear - 1 & " and a0402=12))" & _
'      " and a0404='TOT'"
'   '2015/2/4 END
'
'   strSql = strSql & " Union All select '" & strUserNum & "','8zt2', null, '全所合計', 0, 0, 0, 0, 0 from dual"
'   'strSql = strSql & " order by 1" 'Mark by Amy 2015/06/16
'End If 'end 專業達成點數分佈情況不需列
''end 2020/05/14
'
'   'Modfiy by Amy 2019/12/19 原程式改至SaveTempTB 及 UpdTempTB (專業達成點數分佈情況改寫至 Process2 )
'   Call SaveTempTB(strSql)
'
'   'Modify by Amy 2018/03/12 +bolArrive = False 專業達成點數分佈情況不需列上個月資料
'   If lngMonth = 1 And bolArrive = False Then
'      strSql = "Insert Into Accrpt4201 (ID,r4201,r4202,r4203,r4205) " & StrSQLa
'      cnnConnection.Execute strSql
'      'Modify by Amy 2019/05/14 與 frmacc44r0-ProduceData_Dept 共用 暫存檔,加R4211區分
'      strSql = "update Accrpt420 A set A.r4205=(select b.r4205 from Accrpt4201 B where A.ID=B.ID and A.r4201=B.r4201 and A.r4202=B.r4202 and A.r4203=B.r4203)" & _
'               " where (A.ID,A.r4201,A.r4202,A.r4203) in (select A.ID,A.r4201,A.r4202,A.r4203 from Accrpt420 A,Accrpt4201 B where A.ID='" & strUserNum & "' And R4211 is null  " & _
'                                                                                      "and A.ID=B.ID and A.r4201=B.r4201 and A.r4202=B.r4202 and A.r4203=B.r4203) " & _
'               "And A.ID='" & strUserNum & "' And A.R4211 is null  "
'      cnnConnection.Execute strSql
'   End If
'   'end 2016/2/18
'   Call UpdTempTB
'
'   '讀取暫存檔資料
'    strSql = "Select r4201 as RID,r4202 as A0101,r4203 as C00,r4204 as C01,r4205 as C02,r4206 as C03,r4207 as C04,r4208 as C05 " & _
'                     "From Accrpt420 Where ID='" & strUserNum & "' And R4211 is null  Order by r4201"
'    'end 2019/05/14
'    'end 2019/12/19
''end 2018/03/12
''end 2015/06/16
'
'   intI = 1
'   'edit by nickc 2007/02/07 不用 dll 了
'   'Set RsTemp = objLawDll.ReadRstMsg(intI, strSQL)
'   'Modify by Amy 2015/06/16 改使用暫存檔
'   Set rsNew = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      OutPutData 'RsTemp
'   Else
'      MsgBox "無符合資料！"
'   End If
'   'end 2015/06/16
'ErrHnd:
'   If Err.Number <> 0 Then
'      MsgBox Err.Description
'   End If
End Sub
'設定使用者所選擇的印表機成預設印表機
Private Sub SetPrinter()
   For Each m_Prn In Printers
      If m_Prn.DeviceName = cmbPrinter.Text Then
         Set Printer = m_Prn
         Exit For
      End If
   Next
End Sub
'還原印表機
Private Sub RestorePrinter()
   For Each m_Prn In Printers
      If m_Prn.DeviceName = m_DefaultPrinter Then
         Set Printer = m_Prn
         Exit For
      End If
   Next
End Sub
'列印明細
'Modify by Amy 2020/05/14 +strCmpN 公司名稱
Public Sub doPrint1(ByRef p_Rst As ADODB.Recordset, Optional ByVal strCmpN As String)
   Dim strTemp As String, ii As Integer
   
   iTopMargin = 5
   iLeftMargin = 5
   iRowHeight = 5
   
   SetPrinter
   Printer.PaperSize = vbPRPSA4 '9
   Printer.ScaleMode = vbMillimeters '6 公厘
   Printer.Orientation = vbPRORLandscape '2 橫印
   Printer.Font = "細明體"
   iCharWidth = Printer.TextWidth("A")
   GetPleft
   Call PrintHead(strCmpN) 'Modify by Amy 2020/05/14
   With p_Rst
      .MoveFirst
      Do While Not .EOF
         
         iYPos = iYPos + iRowHeight
         If iYPos + 2 * iRowHeight > Printer.ScaleHeight Then
            iXPos = PLeft(0)
            Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
            Printer.Print String((Printer.ScaleWidth - 2 * iLeftMargin) / Printer.TextWidth("-"), "-")
            Printer.NewPage
            Call PrintHead(strCmpN) 'Modify by Amy 2020/05/14
            iYPos = iYPos + iRowHeight
         End If
         For ii = 0 To .Fields.Count - 1
            Select Case ii
               Case 0
                  strTemp = ChangeTStringToTDateString("" & .Fields(ii))
                  iXPos = PLeft(ii)
               Case 3, 4 '金額
                  strTemp = Format("" & .Fields(ii), "#,##0")
                  iXPos = PLeft(ii + 1) - Printer.TextWidth(strTemp) - 2 * iCharWidth
               Case Else
                  strTemp = "" & .Fields(ii)
                  iXPos = PLeft(ii)
            End Select
            Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
            Printer.Print strTemp
         Next
         .MoveNext
      Loop
      iYPos = iYPos + iRowHeight
      iXPos = PLeft(0)
      Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
      Printer.Print String((Printer.ScaleWidth - 2 * iLeftMargin) / Printer.TextWidth("-"), "-")
   End With
   Printer.EndDoc
   MsgBox "列印完成！"
End Sub
'列印統計
Public Sub DoPrint(ByRef p_Rst As ADODB.Recordset)
Dim strTmp As String, ii As Integer, iRow As Integer
Dim strCmpN As String
   
   iRowHeight = 5
   iTopMargin = 5
   iLeftMargin = 5
   iColWidth(0) = 50
   iColWidth(1) = 28
   iTBWidth = iColWidth(0) + iColWidth(1) * 5
   'Modified by Lydia 2016/02/22
   'iRows = p_Rst.RecordCount + 3
   'iTBHeight = iRowHeight * iRows
   iTBHeight = 3
   nowR = 0
   Call GetTBRows(p_Rst, nowR, iRows)
   
   SetPrinter
   Printer.PaperSize = vbPRPSA4 '9
   Printer.ScaleMode = vbMillimeters '6 公厘
   Printer.Orientation = vbPRORPortrait '1 直印
   'Modified by Lydia 2016/02/22
   'printTable p_Rst
   printTableNew p_Rst
   Printer.FontBold = True
   Printer.FontSize = 15
   Printer.CurrentY = iTopMargin
   Printer.CurrentX = iLeftMargin + iTBWidth / 2 - Printer.TextWidth(Me.Caption) / 2
   Printer.Print strCaption
   
   Printer.FontSize = 12
   
   iRow = 1
   'Add By Sindy 2014/1/22
   'Modify by Amy 2020/05/14 公司別改下拉 原:& IIf(Text3 = "2", "智權", IIf(Text3 = "1", "台一", "台一　專利商標/智權"))
   strCmpN = GetAccReportCmpN(Trim(CboCmp), , True)
   strExc(0) = "公司別：" & strCmpN
   'end 2020/05/14
   Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
   Printer.CurrentX = iLeftMargin + 1
   Printer.Print strExc(0)
   '2014/1/22 END
   strExc(0) = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
   Printer.CurrentX = iLeftMargin + iTBWidth - Printer.TextWidth(strExc(0)) - 1
   Printer.Print strExc(0)
   
   iRow = 2
   'Added by Lydia 2016/02/22
   ii = iLeftMargin + iTBWidth - Printer.TextWidth(strExc(0)) - 1
   strExc(0) = "頁　　次：" & Printer.Page
   Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
   Printer.CurrentX = ii
   Printer.Print strExc(0)
   iRow = 3
   'end 2016/02/22
   strExc(0) = "當月收入"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
   Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
   Printer.Print strExc(0)
   
   strExc(0) = "上月收入"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
   Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 1 + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
   Printer.Print strExc(0)
   
   strExc(0) = "去年同期"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
   Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 2 + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
   Printer.Print strExc(0)
   
   strExc(0) = lngYear & "年1-" & lngMonth & "月"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
   Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 3 + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
   Printer.Print strExc(0)
   
   strExc(0) = lngYear - 1 & "年1-" & lngMonth & "月"
   Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
   Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 4 + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
   Printer.Print strExc(0)
   
   Printer.FontBold = False
   Printer.FontSize = 12
   
   p_Rst.MoveFirst
   Do While Not p_Rst.EOF
      'add by sonia 2016/2/19
      'If iRow > 55 Then
      'Modified by Lydia 2016/02/22
      If iRow >= cRows + 3 Then
        Printer.NewPage
        'Added by Lydia 2016/02/22
        Call GetTBRows(p_Rst, nowR, iRows)
        printTableNew p_Rst  '畫格線
        'end 2016/02/22
        Printer.FontBold = True
        Printer.FontSize = 15
        Printer.CurrentY = iTopMargin
        Printer.CurrentX = iLeftMargin + iTBWidth / 2 - Printer.TextWidth(Me.Caption) / 2
        Printer.Print strCaption
        
        Printer.FontSize = 12
        
        iRow = 1
        'Add By Sindy 2014/1/22
        'Modify by Amy 2020/05/14 公司別改下拉 原:IIf(Text3 = "2", "智權", IIf(Text3 = "1", "台一", "台一　專利商標/智權"))
        strExc(0) = "公司別：" & strCmpN
        Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
        Printer.CurrentX = iLeftMargin + 1
        Printer.Print strExc(0)
        '2014/1/22 END
        strExc(0) = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
        Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
        Printer.CurrentX = iLeftMargin + iTBWidth - Printer.TextWidth(strExc(0)) - 1
        Printer.Print strExc(0)
        
        iRow = 2
        'Added by Lydia 2016/02/22
        ii = iLeftMargin + iTBWidth - Printer.TextWidth(strExc(0)) - 1
        strExc(0) = "頁　　次：" & Printer.Page
        Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
        Printer.CurrentX = ii
        Printer.Print strExc(0)
        iRow = 3
        'end 2016/02/22
        strExc(0) = "當月收入"
        Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
        Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
        Printer.Print strExc(0)
        
        strExc(0) = "上月收入"
        Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
        Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 1 + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
        Printer.Print strExc(0)
        
        strExc(0) = "去年同期"
        Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
        Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 2 + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
        Printer.Print strExc(0)
        
        strExc(0) = lngYear & "年1-" & lngMonth & "月"
        Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
        Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 3 + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
        Printer.Print strExc(0)
        
        strExc(0) = lngYear - 1 & "年1-" & lngMonth & "月"
        Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
        Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * 4 + iColWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
        Printer.Print strExc(0)
        
        Printer.FontBold = False
        Printer.FontSize = 12
      
      End If
      'end 2016/2/19
      iRow = iRow + 1
      strExc(0) = "" & p_Rst.Fields("C00")
      Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
      Printer.CurrentX = iLeftMargin + 1
      Printer.Print strExc(0)
      '上月應收
      For ii = 1 To 5
         strExc(0) = Format(Val("" & p_Rst.Fields("C0" & ii)), FDollar) 'Modify by Amy 2015/06/16 原:DDollar
         Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
         Printer.CurrentX = iLeftMargin + iColWidth(0) + iColWidth(1) * ii - Printer.TextWidth(strExc(0)) - 2
         Printer.Print strExc(0)
      Next
      p_Rst.MoveNext
      nowR = p_Rst.AbsolutePosition 'Added by Lydia 2016/02/22
   Loop
   'Add by Amy 2018/03/12
   iRow = iRow + 3
   Printer.CurrentY = iTopMargin + iRowHeight * iRow - 1
   Printer.CurrentX = iLeftMargin + 1
   Printer.FontBold = True
   Printer.Print "*因四捨五入加總,故紙本與Excel合計會有些許誤差"
   'end 2018/03/12
   Printer.EndDoc
   MsgBox "列印完成！"
End Sub

Private Sub OutPutData()
'Mark by Amy 2015/06/16
'   Dim ADF
'   Dim rsNew As ADODB.Recordset
'   Dim iField As Integer
'   Dim ColInfo()
   Dim dblSubTot1(1 To 5) As Double, dblSubTot2(1 To 5) As Double, dblTot(1 To 5) As Double

On Error GoTo flgErr
  
   'Mark byAmy 2015/06/16
'   Set ADF = CreateObject("RDSServer.DataFactory")
'   With p_Rst
'      .MoveFirst
'      ReDim ColInfo(.Fields.Count - 1)
'      For iField = 0 To UBound(ColInfo)
'         ColInfo(iField) = Array(.Fields(iField).Name, CInt(129), CInt(2000), True)
'      Next
'      Set rsNew = ADF.CreateRecordset(ColInfo)
'      Do While Not .EOF
'         '複製原來資料
'         rsNew.AddNew
'         For iField = 0 To UBound(ColInfo)
'            rsNew.Fields(iField) = .Fields(iField)
'         Next
'         '全所
'         If Right(rsNew.Fields("RID"), 3) = "zt2" Then
'            For iField = 1 To 5
'               rsNew.Fields("C0" & iField) = dblTot(iField) + dblSubTot1(iField)
'            Next
'         ElseIf Right(rsNew.Fields("RID"), 3) = "zt1" Then
'            For iField = 1 To 5
'               dblTot(iField) = dblTot(iField) + dblSubTot1(iField)
'               rsNew.Fields("C0" & iField) = dblTot(iField)
'               dblSubTot1(iField) = 0
'               dblSubTot2(iField) = 0
'            Next
'         ElseIf Right(rsNew.Fields("RID"), 2) = "zz" Then
'            For iField = 1 To 5
'               rsNew.Fields("C0" & iField) = dblSubTot2(iField)
'               dblTot(iField) = dblTot(iField) + dblSubTot2(iField)
'               dblSubTot1(iField) = 0
'               dblSubTot2(iField) = 0
'            Next
'         ElseIf Right(rsNew.Fields("RID"), 1) = "z" Then
'            For iField = 1 To 5
'               rsNew.Fields("C0" & iField) = dblSubTot1(iField)
'               dblSubTot2(iField) = dblSubTot2(iField) + dblSubTot1(iField)
'               dblSubTot1(iField) = 0
'            Next
'         Else
'            For iField = 1 To 5
'               dblSubTot1(iField) = dblSubTot1(iField) + Val("" & rsNew.Fields("C0" & iField))
'            Next
'         End If
'         .MoveNext
'      Loop
'      rsNew.UPDATE
'   End With
   
   'Mark by Amy 2019/12/19 專業達成點數分佈情況(當月實際達成)
'   If Combo2 = "專業達成點數分佈情況(當月實際達成)" Then
'        ExcelSave2
'   Else
        If txtData = "2" Then Call UpdateTotal
        If txtOutput = "1" Then
           Load Frmacc42a1
           With Frmacc42a1
              .DataGrid1.Caption = strCaption
              .DataGrid1.Columns(4).Caption = lngYear & "年1-" & lngMonth & "月"
              .DataGrid1.Columns(5).Caption = lngYear - 1 & "年1-" & lngMonth & "月"
              Set .Adodc1.Recordset = rsNew.Clone
              Set .DataGrid1.DataSource = .Adodc1
              .Hide
              .Show
           End With
        Else
             'Modify By Sindy 2014/1/23
             If m_bolExcel = True Then
                ExcelSave rsNew.Clone
             Else
             '2014/1/23 END
                DoPrint rsNew.Clone
             End If
        End If
'   End If
   'end 2015/06/16
   
flgErr:
   Set rsNew = Nothing
   'Set ADF = Nothing
   If Err.Number <> 0 Then
         'Add byAmy 2015/06/16 避免Excel 開啟未關,造成再次run當掉
        If IsOpenXls = True Then
             'Modify by Amy 2016/06/23 +判斷版本
             If Val(xlsSalesPoint.Version) < 12 Then
                xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
             Else
                xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
             End If
             'end 2016/06/23
             xlsSalesPoint.Workbooks.Close
             xlsSalesPoint.Quit
             Set wksaccrpt424 = Nothing
             Set xlsSalesPoint = Nothing
        End If
        MsgBox Err.Description, vbCritical
   End If
End Sub

'Remark by Lydia 2016/02/22
'Private Sub printTable(ByRef p_Rst As ADODB.Recordset)
'   Dim ii As Integer
'
'   Printer.DrawWidth = 6
'   Printer.Line (iLeftMargin, iTopMargin - 1)-(iLeftMargin + iTBWidth, iTopMargin - 1)
'
'   '橫線
'   Printer.DrawWidth = 3
'   Printer.Line (iLeftMargin, iTopMargin - 1 + iRowHeight * 3)-(iLeftMargin + iTBWidth, iTopMargin - 1 + iRowHeight * 3)
'
'   'For ii = 3 To 55
'      'Printer.Line (iLeftMargin, iTopMargin - 1 + iRowHeight * ii)-(iLeftMargin + iTBWidth, iTopMargin - 1 + iRowHeight * ii)
'   'Next
'   p_Rst.MoveFirst
'   ii = 2
'   Do While Not p_Rst.EOF
'      ii = ii + 1
'      Printer.DrawWidth = 3
'      If Mid(p_Rst.Fields("RID"), 2, 1) = "z" Then
'         Printer.FillStyle = 0
'         Printer.FillColor = &HDDDDDD
'         Printer.Line (iLeftMargin, iTopMargin - 1 + iRowHeight * ii)-(iLeftMargin + iTBWidth, iTopMargin - 1 + iRowHeight * (ii + 1)), , B
'      Else
'         Printer.Line (iLeftMargin, iTopMargin - 1 + iRowHeight * ii)-(iLeftMargin + iTBWidth, iTopMargin - 1 + iRowHeight * ii)
'      End If
'      p_Rst.MoveNext
'   Loop
'
'   '直線
'   Printer.DrawWidth = 6
'   Printer.Line (iLeftMargin, iTopMargin - 1)-(iLeftMargin, iTopMargin - 1 + iRowHeight * iRows)
'
'   Printer.DrawWidth = 3
'   Printer.Line (iLeftMargin + iColWidth(0), iTopMargin - 1 + iRowHeight * 3)-(iLeftMargin + iColWidth(0), iTopMargin - 1 + iRowHeight * iRows)
'   For ii = 1 To 5
'      Printer.Line (iLeftMargin + iColWidth(0) + iColWidth(1) * ii, iTopMargin - 1 + iRowHeight * 3)-(iLeftMargin + iColWidth(0) + iColWidth(1) * ii, iTopMargin - 1 + iRowHeight * iRows)
'   Next
'
'   Printer.DrawWidth = 6
'   Printer.Line (iLeftMargin + iColWidth(0) + iColWidth(1) * 5, iTopMargin - 1)-(iLeftMargin + iColWidth(0) + iColWidth(1) * 5, iTopMargin - 1 + iRowHeight * iRows)
'
'End Sub

Private Sub printTableNew(ByRef p_Rst As ADODB.Recordset)
   Dim ii As Integer
   
   Printer.DrawWidth = 6
   Printer.Line (iLeftMargin, iTopMargin - 1)-(iLeftMargin + iTBWidth, iTopMargin - 1)
   
   '橫線
   Printer.DrawWidth = 3
   Printer.Line (iLeftMargin, iTopMargin - 1 + iRowHeight * iTBHeight)-(iLeftMargin + iTBWidth, iTopMargin - 1 + iRowHeight * iTBHeight)

   ii = 3
   Do While Not p_Rst.EOF
      ii = ii + 1
      Printer.DrawWidth = 3
      If Mid(p_Rst.Fields("RID"), 2, 1) = "z" Then
         Printer.FillStyle = 0
         Printer.FillColor = &HDDDDDD
         Printer.Line (iLeftMargin, iTopMargin - 1 + iRowHeight * ii)-(iLeftMargin + iTBWidth, iTopMargin - 1 + iRowHeight * (ii + 1)), , B
      Else
         Printer.Line (iLeftMargin, iTopMargin - 1 + iRowHeight * ii)-(iLeftMargin + iTBWidth, iTopMargin - 1 + iRowHeight * ii)
      End If
      
      If ii >= cRows + 3 Then Exit Do
      p_Rst.MoveNext
   Loop
   '第二頁~
   If nowR > 0 Then
      p_Rst.MoveFirst
      p_Rst.Move nowR - 1
   End If
   '直線
   Printer.DrawWidth = 6
   Printer.Line (iLeftMargin, iTopMargin - 1)-(iLeftMargin, iTopMargin - 1 + iRowHeight * iRows)
   
   Printer.DrawWidth = 3
   Printer.Line (iLeftMargin + iColWidth(0), iTopMargin - 1 + iRowHeight * iTBHeight)-(iLeftMargin + iColWidth(0), iTopMargin - 1 + iRowHeight * iRows)
   For ii = 1 To 5
      Printer.Line (iLeftMargin + iColWidth(0) + iColWidth(1) * ii, iTopMargin - 1 + iRowHeight * iTBHeight)-(iLeftMargin + iColWidth(0) + iColWidth(1) * ii, iTopMargin - 1 + iRowHeight * iRows)
   Next
   
   Printer.DrawWidth = 6
   Printer.Line (iLeftMargin + iColWidth(0) + iColWidth(1) * 5, iTopMargin - 1)-(iLeftMargin + iColWidth(0) + iColWidth(1) * 5, iTopMargin - 1 + iRowHeight * iRows)
   
End Sub
Sub GetPleft()

   Erase PLeft
   '傳票日期
   PLeft(0) = 4 * iCharWidth
   '1 傳票號碼
   PLeft(1) = PLeft(0) + 9 * iCharWidth
   '部門別
   PLeft(2) = PLeft(1) + 10 * iCharWidth
   '借方金額
   PLeft(3) = PLeft(2) + 6 * iCharWidth
   '貸方金額
   PLeft(4) = PLeft(3) + 15 * iCharWidth
   '摘要
   PLeft(5) = PLeft(4) + 15 * iCharWidth
   '對沖代號(客)
   PLeft(6) = PLeft(5) + 30 * iCharWidth
   '對沖代號(業)
   PLeft(7) = PLeft(6) + 12 * iCharWidth
   '對沖代號(所)
   PLeft(8) = PLeft(7) + 12 * iCharWidth
   '對沖代號(他)
   PLeft(9) = PLeft(8) + 12 * iCharWidth
   
   PLeft(10) = PLeft(9) + 12 * iCharWidth
End Sub

'表頭
'Modify by Amy 2020/05/14 +stCmpN 公司名稱
Private Sub PrintHead(ByVal stCmpN As String)

   Dim strTemp As String
   
   Printer.Font.Size = 20
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   
   strTemp = strCaption
   iXPos = Printer.ScaleWidth / 2 - Printer.TextWidth(strTemp) / 2
   iYPos = iTopMargin
   Printer.CurrentX = iXPos: Printer.CurrentY = iYPos
   Printer.Print strTemp
   
   Printer.Font.Size = 10
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   strTemp = "列印人：" & GetStaffName(strUserNum)
   iYPos = iYPos + 2 * iRowHeight
   iXPos = PLeft(0)
   Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
   Printer.Print strTemp

   strTemp = "列印日期：" & CFDate(strSrvDate(2))
   iXPos = Printer.ScaleWidth - iLeftMargin - Printer.TextWidth(strTemp)
   Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
   Printer.Print strTemp
   
   'Add By Sindy 2014/1/22
   'Modify by Amy 2020/05/14 原:IIf(Text3 = "2", "智權", IIf(Text3 = "1", "台一", "台一　專利商標/智權")
   strTemp = "公司別：" & stCmpN
   iYPos = iYPos + iRowHeight
   iXPos = PLeft(0)
   Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
   Printer.Print strTemp
   '2014/1/22 END
   
   strTemp = "頁　　次：" & Printer.Page
   'iYPos = iYPos + iRowHeight
   'iXPos = Printer.ScaleWidth - iLeftMargin - Printer.TextWidth(strTemp)
   Printer.CurrentY = iYPos: Printer.CurrentX = 258
   Printer.Print strTemp
   
   iYPos = iYPos + iRowHeight
   iXPos = PLeft(0)
   Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
   Printer.Print String((Printer.ScaleWidth - 2 * iLeftMargin) / Printer.TextWidth("-"), "-")
      
   iYPos = iYPos + iRowHeight
   iXPos = PLeft(0)
   Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
   Printer.Print "傳票日期"
   
   iXPos = PLeft(1)
   Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
   Printer.Print "傳票號碼"
   
   iXPos = PLeft(2)
   Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
   Printer.Print "部門別"
   
   iXPos = PLeft(3)
   Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
   Printer.Print "借方金額"
   
   iXPos = PLeft(4)
   Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
   Printer.Print "貸方金額"
   
   iXPos = PLeft(5)
   Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
   Printer.Print "摘要"
   
   iXPos = PLeft(6)
   Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
   Printer.Print "對沖代號(客)"
   
   iXPos = PLeft(7)
   Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
   Printer.Print "對沖代號(業)"
   
   iXPos = PLeft(8)
   Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
   Printer.Print "對沖代號(所)"
   
   iXPos = PLeft(9)
   Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
   Printer.Print "對沖代號(他)"
   
   iYPos = iYPos + iRowHeight
   iXPos = PLeft(0)
   Printer.CurrentY = iYPos: Printer.CurrentX = iXPos
   Printer.Print String((Printer.ScaleWidth - 2 * iLeftMargin) / Printer.TextWidth("-"), "-")
End Sub

'*************************************************
'  轉成Excel檔案-專業點數分析統計
'
'*************************************************
'Mark by Amy 2020/06/05 不使用-婧瑄
Private Sub ExcelSave(ByRef p_Rst As ADODB.Recordset)
'Dim ii As Integer
''Add by Amy 2018/03/12 動態產生欄名
'ReDim strField(0 To 5)
'ReDim intWidth(0 To 5)
'Dim strzzR As String, str7zt1R As String, str8zt2R As String '達成總計/專業達成總計/全所合計 加總列
'Dim intStartR As Integer, bolSum As Boolean '起始列(合計用)/ 是加總欄位
'
'   If Dir(strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
'      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'         MkDir strExcelPath
'      End If
'   Else
'      Kill strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43)
'   End If
'
'   SetPrinter
'
'   m_intPage = 0: m_lngRow = 0: intField = 65
'   xlsSalesPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
'   xlsSalesPoint.Workbooks.add
'   Set wksaccrpt424 = xlsSalesPoint.Worksheets(1)
'   xlsSalesPoint.Visible = True
'   'wksaccrpt424.PageSetup.Orientation = xlLandscape '橫印
'   wksaccrpt424.PageSetup.Orientation = wdOrientLandscape '直印
'   wksaccrpt424.PageSetup.LeftMargin = 28.34
'   wksaccrpt424.PageSetup.RightMargin = 28.34
'   wksaccrpt424.PageSetup.TopMargin = 42.51
'   wksaccrpt424.PageSetup.BottomMargin = 42.51
'   wksaccrpt424.PageSetup.HeaderMargin = 28.34
'   wksaccrpt424.PageSetup.FooterMargin = 28.34
'   'end 2018/03/12
'   Call ExcelHead '頁首
'
'   'Modify by Amy 2018/03/12 合計改為公式
'   intStartR = m_lngRow + 1
'   p_Rst.MoveFirst
'   Do While Not p_Rst.EOF
'      m_lngRow = m_lngRow + 1: bolSum = False
'      If InStr("" & p_Rst.Fields("RID"), "z") > 0 And "" & p_Rst.Fields("RID") <> "6z1" Then bolSum = True
'
'      For ii = 0 To 5
'         strExc(0) = ""
'         '科目名稱
'         If ii = 0 Then
'            strExc(0) = "" & p_Rst.Fields("C00")
'         '加總欄
'         ElseIf bolSum = True Then
'             '合計(Xz:Sum)
'             If Len("" & p_Rst.Fields("RID")) = 2 Then
'                strExc(0) = "=Sum(" & Chr(Asc("a") + ii) & intStartR & ":" & Chr(Asc("a") + ii) & m_lngRow - 1 & ")"
'            '商標/專利達成總計(Xzz:Xz加總)
'            ElseIf Right("" & p_Rst.Fields("RID"), 2) = "zz" Then
'                If ii = 1 Then
'                   strExc(0) = Mid(strzzR, 2)
'                Else
'                    strExc(0) = Replace(Mid(strzzR, 2), Chr(Asc("a") + GetValue("當月收入")), Chr(Asc("a") + ii))
'                End If
'                strExc(0) = "=Sum( " & strExc(0) & ")"
'            '專業達成總計(7zt1:Xzz加總+Sum)
'            ElseIf "" & p_Rst.Fields("RID") = "7zt1" Then
'                If ii = 1 Then
'                   strExc(0) = Mid(str7zt1R, 2)
'                Else
'                    strExc(0) = Replace(Mid(str7zt1R, 2), Chr(Asc("a") + GetValue("當月收入")), Chr(Asc("a") + ii))
'                End If
'                strExc(0) = "=Sum( " & strExc(0) & ")"
'                 If m_lngRow > intStartR Then
'                    strExc(0) = strExc(0) & "+Sum(" & Chr(Asc("a") + ii) & intStartR & ":" & Chr(Asc("a") + ii) & m_lngRow - 1 & ")"
'                End If
'            '全所合計(8zt2:Xzz+Sum)
'            Else
'                If ii = 1 Then
'                   strExc(0) = Mid(str8zt2R, 2)
'                Else
'                    strExc(0) = Replace(Mid(str8zt2R, 2), Chr(Asc("a") + GetValue("當月收入")), Chr(Asc("a") + ii))
'                End If
'                strExc(0) = "=Sum( " & strExc(0) & ")"
'                If m_lngRow > intStartR Then
'                    strExc(0) = strExc(0) & "+Sum(" & Chr(Asc("a") + ii) & intStartR & ":" & Chr(Asc("a") + ii) & m_lngRow - 1 & ")"
'                End If
'            End If
'
'         Else
'            'Modified by Lydia 2014/12/18 excel資料的數字為0時以 0顯示之(DDollar => DDollar2)
'            strExc(0) = Format(Val("" & p_Rst.Fields("C0" & ii)), DDollar2)
'         End If
'         wksaccrpt424.Range(Chr(Asc("a") + ii) & m_lngRow).Value = strExc(0)
'      Next
'      If InStr("" & p_Rst.Fields("RID"), "z") > 0 And "" & p_Rst.Fields("RID") <> "8zt2" Then
'        intStartR = m_lngRow + 1
'        '記錄合計列
'        If Len("" & p_Rst.Fields("RID")) = 2 Or "" & p_Rst.Fields("RID") = "6z1" Then
'            strzzR = strzzR & "+" & Chr(Asc("a") + GetValue("當月收入")) & m_lngRow
'        '記錄達成總計列
'        ElseIf Right("" & p_Rst.Fields("RID"), 2) = "zz" Then
'            strzzR = ""
'            str7zt1R = str7zt1R & "+" & Chr(Asc("a") + GetValue("當月收入")) & m_lngRow
'        '記錄專業達成總計列7zt1
'        Else
'            str8zt2R = str8zt2R & "+" & Chr(Asc("a") + GetValue("當月收入")) & m_lngRow
'        End If
'
'      End If
''      wksaccrpt424.Range("B" & m_lngRow & ":F" & m_lngRow).Select
''      wksaccrpt424.Application.Selection.NumberFormatLocal = "0.00_ ;[紅色]-0.00 "
'      p_Rst.MoveNext
'   Loop
'   wksaccrpt424.Range("a" & m_lngRow + 2).Value = "*因四捨五入加總,故Excel與紙本合計會有些許誤差"
'   wksaccrpt424.Range("a" & m_lngRow + 2).Font.Color = &HFF& '紅字
'   'end 2018/03/12
'
'   'Modify by Amy2016/05/05 判斷若版本2007以上改變存格式
'   If Val(xlsSalesPoint.Version) < 12 Then
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
'   Else
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
'   End If
'   'end 2016/05/05
'   xlsSalesPoint.Workbooks.Close
'   xlsSalesPoint.Quit
'   Set wksaccrpt424 = Nothing
'   Set xlsSalesPoint = Nothing
'   StatusClear
'   MsgBox "Excel檔產生完畢！（" & strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43) & "）"
End Sub

Private Sub ExcelHead()
   m_intPage = m_intPage + 1
   m_lngRow = m_lngRow + 1
   wksaccrpt424.Range("a" & m_lngRow).Value = strCaption
   wksaccrpt424.Range("a" & m_lngRow & ":f" & m_lngRow).Select
   With xlsSalesPoint.Selection
       .HorizontalAlignment = xlCenter
       .VerticalAlignment = xlBottom
       .WrapText = False
       .Orientation = 0
       .AddIndent = False
       .ShrinkToFit = False
       .MergeCells = True
   End With
   wksaccrpt424.Application.Selection.Font.Bold = True
   wksaccrpt424.Application.Selection.Font.Size = 16
   
   m_lngRow = m_lngRow + 1: ii = 0
   'Modify by Amy 2018/03/12 改自動產生欄名
   wksaccrpt424.Range("a" & m_lngRow).Value = "　": strField(ii) = "　": intWidth(ii) = 20: ii = ii + 1
   wksaccrpt424.Range("b" & m_lngRow).Value = "當月收入": strField(ii) = "當月收入": intWidth(ii) = 14: ii = ii + 1
   wksaccrpt424.Range("c" & m_lngRow).Value = "上月收入": strField(ii) = "上月收入": intWidth(ii) = 14: ii = ii + 1
   wksaccrpt424.Range("d" & m_lngRow).Value = "去年同期": strField(ii) = "去年同期": intWidth(ii) = 14: ii = ii + 1
   wksaccrpt424.Range("e" & m_lngRow).Value = lngYear & "年1-" & lngMonth & "月": strField(ii) = lngYear & "年1-" & lngMonth & "月": intWidth(ii) = 14: ii = ii + 1
   wksaccrpt424.Range("f" & m_lngRow).Value = lngYear - 1 & "年1-" & lngMonth & "月": strField(ii) = lngYear - 1 & "年1-" & lngMonth & "月": intWidth(ii) = 14: ii = ii + 1
   For ii = LBound(intWidth) To UBound(intWidth)
        wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).ColumnWidth = intWidth(ii)
   Next ii
   wksaccrpt424.Range("A" & m_lngRow & ":f" & m_lngRow).Select
   'end 2018/03/12
   wksaccrpt424.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
   wksaccrpt424.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
   wksaccrpt424.Application.Selection.Borders(xlEdgeLeft).LineStyle = xlNone
   wksaccrpt424.Application.Selection.Borders(xlEdgeTop).LineStyle = xlNone
   With wksaccrpt424.Application.Selection.Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .Weight = xlThin
      .ColorIndex = xlAutomatic
   End With
   wksaccrpt424.Application.Selection.Borders(xlEdgeRight).LineStyle = xlNone
   wksaccrpt424.Application.Selection.Borders(xlInsideVertical).LineStyle = xlNone
End Sub

'*************************************************
'  轉成Excel檔案-專業點數分析明細
'
'*************************************************
Private Sub ExcelSave1(ByRef p_Rst As ADODB.Recordset, ByVal strCmpN As String)
Dim ii As Integer
Dim strTemp(1 To 10) As String
   
   If Dir(strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43)
   End If
   
   SetPrinter
   
   m_intPage = 0: m_lngRow = 0
   xlsSalesPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsSalesPoint.Workbooks.add
   Set wksaccrpt424 = xlsSalesPoint.Worksheets(1)
   'xlsSalesPoint.Visible = True
   wksaccrpt424.PageSetup.Orientation = xlLandscape '橫印
   'wksaccrpt424.PageSetup.Orientation = wdOrientLandscape '直印
   wksaccrpt424.PageSetup.LeftMargin = 28.34
   wksaccrpt424.PageSetup.RightMargin = 28.34
   wksaccrpt424.PageSetup.TopMargin = 42.51
   wksaccrpt424.PageSetup.BottomMargin = 42.51
   wksaccrpt424.PageSetup.HeaderMargin = 28.34
   wksaccrpt424.PageSetup.FooterMargin = 28.34
   wksaccrpt424.Columns("a:a").ColumnWidth = 10
   wksaccrpt424.Columns("b:b").ColumnWidth = 11
   wksaccrpt424.Columns("c:c").ColumnWidth = 7
   wksaccrpt424.Columns("d:d").ColumnWidth = 10
   wksaccrpt424.Columns("e:e").ColumnWidth = 10
   wksaccrpt424.Columns("f:f").ColumnWidth = 30
   wksaccrpt424.Columns("g:g").ColumnWidth = 12
   wksaccrpt424.Columns("h:h").ColumnWidth = 12
   wksaccrpt424.Range("g:j").Select
   xlsSalesPoint.Selection.NumberFormatLocal = "@"
   wksaccrpt424.Columns("i:i").ColumnWidth = 12
   wksaccrpt424.Columns("j:j").ColumnWidth = 12
   Call ExcelHead1(strCmpN) 'Modify by Amy 2020/05/14 頁首
   
   With p_Rst
      .MoveFirst
      Do While Not .EOF
         If m_lngRow Mod 32 = 0 Then
            '換頁
            wksaccrpt424.Range("A" & (m_lngRow + 1)).Select
            wksaccrpt424.HPageBreaks.add Before:=wksaccrpt424.Application.ActiveCell
            Call ExcelHead1(strCmpN) 'Modify by Amy 2020/05/14 頁首
         End If
         '清空變數值
         For ii = 1 To 10
            strTemp(ii) = ""
         Next ii
         '讀取欄位值
         For ii = 0 To .Fields.Count - 1
            Select Case ii
               Case 0
                  strTemp(ii + 1) = ChangeTStringToTDateString("" & .Fields(ii))
               Case 3, 4 '金額
                  strTemp(ii + 1) = Format("" & .Fields(ii), "#,##0")
               Case Else
                  strTemp(ii + 1) = "" & .Fields(ii)
            End Select
         Next ii
         '存放欄位
         m_lngRow = m_lngRow + 1
         wksaccrpt424.Range("a" & m_lngRow).Value = strTemp(1)
         wksaccrpt424.Range("b" & m_lngRow).Value = strTemp(2)
         wksaccrpt424.Range("c" & m_lngRow).Value = strTemp(3)
         wksaccrpt424.Range("d" & m_lngRow).Value = strTemp(4)
         wksaccrpt424.Range("e" & m_lngRow).Value = strTemp(5)
         wksaccrpt424.Range("f" & m_lngRow).Value = strTemp(6)
         wksaccrpt424.Range("g" & m_lngRow).Value = strTemp(7)
         wksaccrpt424.Range("h" & m_lngRow).Value = strTemp(8)
         wksaccrpt424.Range("i" & m_lngRow).Value = strTemp(9)
         wksaccrpt424.Range("j" & m_lngRow).Value = strTemp(10)
         .MoveNext
      Loop
   End With
   
   'Modify by Amy2016/05/05 判斷若版本2007以上改變存格式
   If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
   End If
   'end 2016/05/05
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   Set wksaccrpt424 = Nothing
   Set xlsSalesPoint = Nothing
   StatusClear
   MsgBox "Excel檔產生完畢！（" & strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43) & "）"
End Sub

'Modify by Amy 2020/05/14 +strCmpN 公司名稱
Private Sub ExcelHead1(ByVal strCmpN As String)
   m_intPage = m_intPage + 1
   m_lngRow = m_lngRow + 1
   wksaccrpt424.Range("a" & m_lngRow).Value = strCaption
   wksaccrpt424.Range("a" & m_lngRow & ":j" & m_lngRow).Select
   With xlsSalesPoint.Selection
       .HorizontalAlignment = xlCenter
       .VerticalAlignment = xlBottom
       .WrapText = False
       .Orientation = 0
       .AddIndent = False
       .ShrinkToFit = False
       .MergeCells = True
   End With
   wksaccrpt424.Application.Selection.Font.Bold = True
   wksaccrpt424.Application.Selection.Font.Size = 16
   
   m_lngRow = m_lngRow + 1
   wksaccrpt424.Range("a" & m_lngRow).Value = "列印人：" & GetStaffName(strUserNum)
   wksaccrpt424.Range("i" & m_lngRow).Value = "列印日期：" & CFDate(strSrvDate(2))
   m_lngRow = m_lngRow + 1
   'Modify by Amy 2020/05/14 公司別改抓變數 原:IIf(Text3 = "2", "智權", IIf(Text3 = "1", "台一", "台一　專利商標/智權"))
   wksaccrpt424.Range("a" & m_lngRow).Value = "公司別：" & strCmpN
   wksaccrpt424.Range("i" & m_lngRow).Value = "頁　　次：" & m_intPage
   m_lngRow = m_lngRow + 1
   wksaccrpt424.Range("a" & m_lngRow).Value = "傳票日期"
   wksaccrpt424.Range("b" & m_lngRow).Value = "傳票號碼"
   wksaccrpt424.Range("c" & m_lngRow).Value = "部門別"
   wksaccrpt424.Range("d" & m_lngRow).Value = "借方金額"
   wksaccrpt424.Range("e" & m_lngRow).Value = "貸方金額"
   wksaccrpt424.Range("f" & m_lngRow).Value = "摘要"
   wksaccrpt424.Range("g" & m_lngRow).Value = "對沖代號(客)"
   wksaccrpt424.Range("h" & m_lngRow).Value = "對沖代號(業)"
   wksaccrpt424.Range("i" & m_lngRow).Value = "對沖代號(所)"
   wksaccrpt424.Range("j" & m_lngRow).Value = "對沖代號(他)"
   wksaccrpt424.Range("A" & m_lngRow & ":J" & m_lngRow).Select
   wksaccrpt424.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
   wksaccrpt424.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
   wksaccrpt424.Application.Selection.Borders(xlEdgeLeft).LineStyle = xlNone
   wksaccrpt424.Application.Selection.Borders(xlEdgeTop).LineStyle = xlNone
   With wksaccrpt424.Application.Selection.Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .Weight = xlThin
      .ColorIndex = xlAutomatic
   End With
   wksaccrpt424.Application.Selection.Borders(xlEdgeRight).LineStyle = xlNone
   wksaccrpt424.Application.Selection.Borders(xlInsideVertical).LineStyle = xlNone
End Sub

'Mark by Amy 2019/12/19 'Add by Amy 2015/06/16
'產生Excel檔案-專業達成點數分佈情況(當月實際達成)
Private Sub ExcelSave2_Old(ByRef p_Rst As ADODB.Recordset)
'    Dim strTp(2) As String
'    Dim bolFormaula As Boolean
'    Dim strReplace, strReplace1
'    Dim strTotal(1 To 8) As String, AllTotal(1 To 8) As String
'    ReDim strField(0 To 8)
'    ReDim intWidth(0 To 8)
'    ReDim DstrTitle(1)
'
'    If Dir(strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
'       If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'          MkDir strExcelPath
'       End If
'    Else
'       Kill strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43)
'    End If
'
'    SetPrinter
'
'    xlsSalesPoint.Workbooks.add
'    Set wksaccrpt424 = xlsSalesPoint.Worksheets(1)
'    IsOpenXls = True
'    xlsSalesPoint.Visible = False
'    wksaccrpt424.PageSetup.Orientation = wdOrientLandscape '直印
'    wksaccrpt424.PageSetup.LeftMargin = 28.34
'    wksaccrpt424.PageSetup.RightMargin = 28.34
'    wksaccrpt424.PageSetup.TopMargin = 42.51
'    wksaccrpt424.PageSetup.BottomMargin = 42.51
'    wksaccrpt424.PageSetup.HeaderMargin = 28.34
'    wksaccrpt424.PageSetup.FooterMargin = 28.34
'
'    ii = 0: m_lngRow = 1: intField = 65
'
'    Call ExcelHead2(False) '頁首
'
'    m_lngRow = m_lngRow + 1: StartRow = m_lngRow: p_Rst.MoveFirst
'    Do While Not p_Rst.EOF
'        For ii = 0 To UBound(strField)
'            If ii <> GetValue("") And ii <> GetValue("同期增減比率") And ii <> GetValue("所佔比率") And ii <> GetValue("所佔比率1") And ii <> GetValue("同期累計比較") _
'              And InStr(p_Rst.Fields("RID"), "z") > 0 And InStr(p_Rst.Fields("RID"), "z1") = 0 Then
'                Select Case p_Rst.Fields("RID")
'                    'modify by sonia 2016/2/16 調整各項之RID
'                    Case "1z", "2z", "3z", "4z", "5z" '合計
'                        strTp(0) = "Sum(" & Chr(intField + ii) & StartRow & ":" & Chr(intField + ii) & m_lngRow - 1 & ")"
'                        strTotal(ii) = strTotal(ii) & Chr(intField + ii) & m_lngRow & ","
'                    Case "3zz", "6zz" '達成總計
'                        strTp(0) = "Sum(" & Left(strTotal(ii), Len(strTotal(ii)) - 1) & ")"
'                        AllTotal(ii) = AllTotal(ii) & Chr(intField + ii) & m_lngRow & ","
'                        strTotal(ii) = ""
'                    Case "7zt1" '專業達成總計
'                        strTp(0) = "Sum(" & AllTotal(ii) & "Sum(" & Chr(intField + ii) & StartRow & ":" & Chr(intField + ii) & m_lngRow - 1 & ")" & ")"
'                End Select
'                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Formula = "=" & strTp(0)
'                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = "#,##0.00"
'            Else
'                bolFormaula = False
'                Select Case ii
'                    Case GetValue("")
'                        strTp(0) = p_Rst.Fields("C00")
'                        strTp(1) = ""
'                    Case GetValue("當月收入")
'                        strTp(0) = Val("" & p_Rst.Fields("C01"))
'                        strTp(1) = "#,##0.00"
'                    Case GetValue("去年同期")
'                        strTp(0) = Val("" & p_Rst.Fields("C03"))
'                        strTp(1) = "#,##0.00"
'                    Case GetValue("同期增減比率")
'                        strTp(0) = "(" & Chr(intField + GetValue("當月收入")) & m_lngRow & "/" & Chr(intField + GetValue("去年同期")) & m_lngRow & ")-1"
'                        strTp(1) = "0.00%"
'                        strTp(2) = Chr(intField + GetValue("去年同期")) & m_lngRow
'                        bolFormaula = True
'                    Case GetValue(DstrTitle(0))
'                        strTp(0) = Val("" & p_Rst.Fields("C04"))
'                        strTp(1) = "#,##0.00"
'                        strTp(2) = ""
'                    Case GetValue("所佔比率")
'                        '資料未跑完無法知道最後一筆位置,故先預設表名欄位,之後再取代
'                        strTp(0) = Chr(intField + GetValue(DstrTitle(0))) & m_lngRow & "/" & Chr(intField + GetValue(DstrTitle(0))) & "$1"
'                        strTp(1) = "0.00%"
'                        strTp(2) = Chr(intField + GetValue(DstrTitle(0))) & "$1"
'                        bolFormaula = True
'                    Case GetValue(DstrTitle(1))
'                        strTp(0) = Val("" & p_Rst.Fields("C05"))
'                        strTp(1) = "#,##0.00"
'                        strTp(2) = ""
'                    Case GetValue("所佔比率1")
'                        '資料未跑完無法知道最後一筆位置,故先預設表名欄位,之後再取代
'                        strTp(0) = Chr(intField + GetValue(DstrTitle(1))) & m_lngRow & "/" & Chr(intField + GetValue(DstrTitle(1))) & "$2"
'                        strTp(1) = "0.00%"
'                         strTp(2) = Chr(intField + GetValue(DstrTitle(1))) & "$2"
'                        bolFormaula = True
'                    Case GetValue("同期累計比較")
'                        strTp(0) = "(" & Chr(intField + GetValue(DstrTitle(0))) & m_lngRow & "/" & Chr(intField + GetValue(DstrTitle(1))) & m_lngRow & ")-1"
'                        strTp(1) = "0.00%"
'                        strTp(2) = Chr(intField + GetValue(DstrTitle(1))) & m_lngRow
'                        bolFormaula = True
'                End Select
'                If bolFormaula = True Then
'                    If p_Rst.Fields("RID") = "7zt1" And (ii = GetValue("所佔比率") Or ii = GetValue("所佔比率1")) Then
'                        '婧瑄:此兩個欄位不需顯示
'                    Else
'                        wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Formula = PUB_ChkExcelZero(2, strTp(2), strTp(0))
'                    End If
'                Else
'                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = strTp(0)
'                End If
'                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = strTp(1)
'                If ii <> GetValue("") And InStr(p_Rst.Fields("RID"), "z1") > 0 Then strTotal(ii) = strTotal(ii) & Chr(intField + ii) & m_lngRow & ","
'            End If
'        Next ii
'         m_lngRow = m_lngRow + 1
'         If InStr(p_Rst.Fields("RID"), "z") > 0 And InStr(p_Rst.Fields("RID"), "z1") = 0 Then StartRow = m_lngRow
'         If p_Rst.Fields("RID") = "7zt1" Then m_lngRow = m_lngRow - 1: Exit Do
'        p_Rst.MoveNext
'    Loop
'    '要取代的文字
'    strReplace = Array(Chr(intField + GetValue(DstrTitle(0))) & "$1", Chr(intField + GetValue(DstrTitle(1))) & "$2", "所佔比率1")
'    '取代成
'    strReplace1 = Array("$" & m_lngRow, "$" & m_lngRow, "所佔比率")
'    For ii = 0 To UBound(strReplace)
'        If ii = UBound(strReplace) Then
'            strTp(2) = Chr(intField + GetValue("" & strReplace(ii)))
'        Else
'            strReplace1(ii) = "$" & Left(strReplace(ii), 1) & strReplace1(ii)
'            strTp(2) = Chr(Asc(Left(strReplace(ii), 1)) + 1)
'        End If
'       wksaccrpt424.Columns(strTp(2)).Replace what:=strReplace(ii), Replacement:=strReplace1(ii), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
'    Next ii
'
'    '框線
'    wksaccrpt424.Range(Chr(intField) & "1:" & Chr(intField + UBound(strField)) & m_lngRow).Select
'     xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'     xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
'     xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
'     xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
'     xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
'     xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
'     xlsSalesPoint.Selection.Font.Size = 9
'
'    'Mark by Amy 2016/04/08 取消分組小計
''    wksaccrpt424.Range(Chr(intField) & "1:" & Chr(intField + UBound(strField)) & m_lngRow).Select
''    xlsSalesPoint.Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(5, 6, 7), Replace:=True, PageBreaks:=False, SummaryBelowData:=True
''    xlsSalesPoint.ActiveSheet.Outline.ShowLevels RowLevels:=2
'
'   'Modify by Amy2016/05/05 判斷若版本2007以上改變存格式
'   If Val(xlsSalesPoint.Version) < 12 Then
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
'   Else
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
'   End If
'   'end 2016/05/05
'   xlsSalesPoint.Workbooks.Close
'   xlsSalesPoint.Quit
'   Set wksaccrpt424 = Nothing
'   Set xlsSalesPoint = Nothing
'
'   StatusClear
'   MsgBox "Excel檔產生完畢！（" & strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43) & "）"
End Sub

'Mark by Amy 2018/03/12 改新格式-婧瑄
Private Sub ExcelHead2_Old(ByVal bolReportTitle As Boolean)
'    If bolReportTitle = True Then
'        wksaccrpt424.Range(Chr(intField) & m_lngRow).Value = lngYear & "年" & lngMonth & "月份各" & Combo2
'
'        wksaccrpt424.Range(Chr(intField) & m_lngRow & ":" & Chr(intField + UBound(strField)) & m_lngRow).Select
'        With xlsSalesPoint.Selection
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .ShrinkToFit = False
'            .MergeCells = True
'        End With
'        wksaccrpt424.Application.Selection.Font.Bold = True
'        wksaccrpt424.Application.Selection.Font.Size = 16
'
'    m_lngRow = m_lngRow + 1
'    wksaccrpt424.Range(Chr(intField) & m_lngRow).Value = "公司別：" & IIf(Text3 = "2", "智權", IIf(Text3 = "1", "台一", "台一　專利商標/智權"))
'    wksaccrpt424.Range(Chr(intField + UBound(strField) - 2) & m_lngRow).Value = "列印日期：" & CFDate(strSrvDate(2))
'    m_lngRow = m_lngRow + 1
'    End If
'
'    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = "": strField(ii) = "": intWidth(ii) = 16: ii = ii + 1
'    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = "當月收入": strField(ii) = "當月收入": intWidth(ii) = 9: ii = ii + 1
'    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = "去年同期": strField(ii) = "去年同期": intWidth(ii) = 9: ii = ii + 1
'    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = "同期增減比率": strField(ii) = "同期增減比率": intWidth(ii) = 8.5: ii = ii + 1
'    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = lngYear & "年1-" & lngMonth & "月": strField(ii) = lngYear & "年1-" & lngMonth & "月": intWidth(ii) = 9.5: DstrTitle(0) = strField(ii): ii = ii + 1
'
'    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = "所佔比率": strField(ii) = "所佔比率": intWidth(ii) = 7: ii = ii + 1
'    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = lngYear - 1 & "年1-" & lngMonth & "月": strField(ii) = lngYear - 1 & "年1-" & lngMonth & "月": intWidth(ii) = 9.5: DstrTitle(1) = strField(ii): ii = ii + 1
'    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = "所佔比率1": strField(ii) = "所佔比率1": intWidth(ii) = 7: ii = ii + 1
'    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = "同期累計比較": strField(ii) = "同期累計比較": intWidth(ii) = 9.5: ii = ii + 1
'    For ii = 0 To UBound(intWidth)
'        wksaccrpt424.Columns(Chr(intField + ii)).ColumnWidth = intWidth(ii)
'    Next ii
End Sub

'Add by Amy 2019/12/19 專業達成點數分佈情況(當月實際達成)
'比較三年(傳票資料寫入暫存TB,否則一直在抓傳票資料會很慢)
Private Sub ExcelSave2(ByVal strCmpN As String)
    Const strSpecAcc As String = "410101,410104,417201,411101,417101,417104,417105,417109"
    Dim strTotal(1 To 10) As String, AllTotal(1 To 10) As String
    Dim strTp(1) As String
    Dim strWkName As String
    Dim intXlsSheet As Integer, intQ As Integer, intA As Integer
    Dim bolFormaula As Boolean, Is417101 As Boolean
    Dim RsQ As ADODB.Recordset, rsA As ADODB.Recordset '結餘點數/更新抓資料用
    Dim strBP As String, strUpd As String, strA As String, strF(2) As String, strWhere As String
    Dim strUpdAcc As String
    Dim strCmp As String 'Add by Amy 2020/05/14

On Error GoTo ErrHand
    If Dir(strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
       If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
          MkDir strExcelPath
       End If
    Else
       Kill strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43)
    End If

    SetPrinter
    intXlsSheet = 1: intField = 65: bolSheet2 = False
    xlsSalesPoint.SheetsInNewWorkbook = 3
    xlsSalesPoint.Workbooks.add
    
NextSheet:
    If strWkName = MsgText(601) Then strWkName = Left(xlsSalesPoint.Worksheets(1).Name, Len(xlsSalesPoint.Worksheets(1).Name) - 1)
    Set wksaccrpt424 = xlsSalesPoint.Worksheets(strWkName & intXlsSheet)
    wksaccrpt424.Activate
    IsOpenXls = True
    'xlsSalesPoint.Visible = True
    wksaccrpt424.PageSetup.PaperSize = 9 'A4
    wksaccrpt424.PageSetup.Orientation = xlLandscape '橫印
    wksaccrpt424.PageSetup.LeftMargin = 28.34
    wksaccrpt424.PageSetup.RightMargin = 28.34
    wksaccrpt424.PageSetup.TopMargin = 42.51
    wksaccrpt424.PageSetup.BottomMargin = 42.51
    wksaccrpt424.PageSetup.HeaderMargin = 28.34
    wksaccrpt424.PageSetup.FooterMargin = 28.34

    m_lngRow = 1
    Call ExcelHead2(strCmpN) 'Add by Amy 2020/05/14

    intTitleR = m_lngRow - 1: StartRow = m_lngRow
    '產生表二,重抓暫存檔資料
    If bolSheet2 = True Then
        '與 frmacc44r0-ProduceData_Dept 共用 暫存檔,加R4211區分
        strSql = "Select r4203 as C00,Round(r4204,0) as C01,Round(r4206,0) as C02,Round(R4210,0) as C03,'' as C04,'' as C05" & _
                        ",Round(r4207,0) as C06,Round(r4208,0) as C07,Round(R4209,0) as C08,'' as C09,'' as C10, r4201 as RID,r4202 as A0101 " & _
                    "From Accrpt420 Where ID='" & strUserNum & "' And R4211 is null Order by r4201"
        intQ = 1
        Set rsNew = ClsLawReadRstMsg(intQ, strSql)
    End If
    rsNew.MoveFirst
    Do While rsNew.EOF = False
        For ii = 0 To UBound(strField)
            '比率
            If ii > GetValue("") And InStr(strField(ii), "vs") > 0 Then
                If InStr(strField(ii), "年") > 0 Then
                    strTp(0) = Chr(intField + GetValue(Mid(strField(ii), 1, InStr(strField(ii), "年vs") - 1) & "年1-" & lngMonth & "月")) & m_lngRow & "/" & _
                                    Chr(intField + GetValue(Mid(strField(ii), InStr(strField(ii), "年vs") + 4) & "年1-" & lngMonth & "月")) & m_lngRow & "-1"
                Else
                    strTp(0) = Chr(intField + GetValue(Mid(strField(ii), 1, InStr(strField(ii), "vs") - 2) & "." & lngMonth)) & m_lngRow & "/" & _
                                    Chr(intField + GetValue(Mid(strField(ii), InStr(strField(ii), "vs") + 3) & "." & lngMonth)) & m_lngRow & "-1"
                End If
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Formula = "=" & strTp(0)
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = "0.00%;[紅色]-0.00%"
                'Modify by Amy 2021/01/21 +CFP合計,拿掉 And "" & rsNew.Fields("RID") <> "6z1"
                If InStr(rsNew.Fields("RID"), "z") > 0 Then
                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Font.Bold = True
                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.ColorIndex = 40 '設置儲存格填充色(膚)
                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.tintandshade = 0.5 '設深淺
                End If
            '合計
            'Modify by Amy 2021/01/21 +CFP合計,拿掉 And "" & rsNew.Fields("RID") <> "6z1"
            ElseIf ii > GetValue("") And InStr(rsNew.Fields("RID"), "z") > 0 Then
                Select Case rsNew.Fields("RID")
                    'Modify by Amy 2021/01/21 +61z
                    Case "1z", "2z", "3z", "4z", "5z", "61z" '合計
                        strTp(0) = "Sum(" & Chr(intField + ii) & StartRow & ":" & Chr(intField + ii) & m_lngRow - 1 & ")"
                        strTotal(ii) = strTotal(ii) & Chr(intField + ii) & m_lngRow & ","
                    Case "3zz", "6zz" '達成總計
                        strTp(0) = "Sum(" & Left(strTotal(ii), Len(strTotal(ii)) - 1) & ")"
                        AllTotal(ii) = AllTotal(ii) & Chr(intField + ii) & m_lngRow & ","
                        strTotal(ii) = ""
                    'Modify by Amy 2020/07/09 因加法務收入-其他 原:7zt1
                    Case "9zt1" '專業達成總計
                        strTp(0) = "Sum(" & AllTotal(ii) & "Sum(" & Chr(intField + ii) & StartRow & ":" & Chr(intField + ii) & m_lngRow - 1 & ")" & ")"
                End Select
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Formula = "=" & strTp(0)
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = "#,##0"
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Font.Bold = True
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.ColorIndex = 40 '設置儲存格填充色(膚)
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.tintandshade = 0.5 '設深淺
            '內容
            Else
                'Mark by Amy 2021/01/21 +CFP合計 原資料RID=6z1改6
                'If ii > GetValue("") And "" & rsNew.Fields("RID") = "6z1" Then strTotal(ii) = strTotal(ii) & Chr(intField + ii) & m_lngRow & ","
                strTp(0) = "" & rsNew.Fields(ii)
                strTp(1) = "#,##0"
                If ii = GetValue("") Then
                    strTp(1) = ""
                Else
                    strTp(0) = Val(strTp(0))
                End If
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = strTp(0)
                'Modify by Amy  +CFP合計 原資料RID=6z1改6,拿掉 And InStr(rsNew.Fields("RID"), "z1") = 0
                If ii = GetValue("" & strField(0)) And InStr(rsNew.Fields("RID"), "z") > 0 Then
                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Font.Bold = True
                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.ColorIndex = 40 '設置儲存格填充色(膚)
                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.tintandshade = 0.5 '設深淺
                Else
                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = strTp(1)
                End If
            End If
        Next ii
        m_lngRow = m_lngRow + 1
        'Modify by Amy  +CFP合計 原資料RID=6z1改6,拿掉 And InStr(rsNew.Fields("RID"), "z1") = 0
        If InStr(rsNew.Fields("RID"), "z") > 0 Then StartRow = m_lngRow
        rsNew.MoveNext
    Loop
    'Memo by Amy 工作表一之「專業達成總計 」= 智權點數實績與結餘分析表 之 全所 「當月實績」+「當月結餘」(frmacc44j0)
    
    If bolSheet2 = False Then
        strF(0) = "Sum(decode(floor(a0205/100)," & lngThisMonth & ",ax206)) Sd1, Sum(decode(floor(a0205/100)," & lngThisMonth & ",ax207)) Sc1" & _
                    ", Sum(Decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax206)) Sd3, Sum(Decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax207)) Sc3" & _
                    ", Sum(Decode(floor(a0205/100)," & lngThisMonth - 200 & ",ax206)) Sd7,Sum(Decode(floor(a0205/100)," & lngThisMonth - 200 & ",ax207)) Sc7" & _
                    ", Sum(Decode(floor(a0205/10000)," & lngYear & ",ax206)) Sd4, Sum(Decode(floor(a0205/10000)," & lngYear & ",ax207)) Sc4" & _
                    ", Sum(Decode(floor(a0205/10000)," & lngYear - 1 & ",ax206)) Sd5, Sum(Decode(floor(a0205/10000)," & lngYear - 1 & ",ax207)) Sc5" & _
                    ", Sum(Decode(floor(a0205/10000)," & lngYear - 2 & ",ax206)) Sd6,Sum(decode(floor(a0205/10000)," & lngYear - 2 & ",ax207)) Sc6"
        
        '未使用strF(1)
        strF(1) = "Sum(Decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1, Sum(Decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
                    ", Sum(Decode(a0401*100+a0402," & lngThisMonth - 200 & ",a0408)) net7, Sum(Decode(a0401," & lngYear & ",a0408)) net4" & _
                    ", Sum(Decode(a0401," & lngYear - 1 & ",a0408)) net5, Sum(Decode(a0401," & lngYear - 2 & ",a0408)) net6"
                        
        strF(2) = "Nvl(Decode(a0103,'1',(Sd1-Sc1),(Sc1-Sd1)),0) C01, Nvl(Decode(a0103,'1',(Sd3-Sc3),(Sc3-Sd3)),0) C03" & _
                    ", Nvl(Decode(a0103,'1',(Sd7-Sc7),(Sc7-Sd7)),0) C07, Nvl(Decode(a0103,'1',(Sd4-Sc4),(Sc4-Sd4)),0) C04" & _
                    ", Nvl(Decode(a0103,'1',(Sd5-Sc5),(Sc5-Sd5)),0) C05, Nvl(Decode(a0103,'1',(Sd6-Sc6),(Sc6-Sd6)),0) C06"
        
        'Modify by Amy 2020/05/14 公司別改下拉
'        If Text3 = "2" Then
'            strWhere = " And R001='J' "
'        ElseIf Text3 = "1" Then
'            strWhere = " And R001='1' "
'        End If
        If stCon021 <> MsgText(601) Then
            strWhere = Replace(Replace(stCon021, " ax205 ", " R004 "), " ax201 ", " R001 ")
        End If
        'end 2020/05/14
        
        strWhere = strWhere & _
             "And (  (R003>=" & lngYear & "0101 And R003<=" & lngYear & IIf(lngMonth < 10, "0" & lngMonth, lngMonth) & "31) " & _
                  " Or (R003>=" & lngYear - 1 & "0101 And R003<=" & lngYear - 1 & IIf(lngMonth < 10, "0" & lngMonth, lngMonth) & "31) " & _
                  " Or (R003>=" & lngYear - 2 & "0101 And R003<=" & lngYear - 2 & IIf(lngMonth < 10, "0" & lngMonth, lngMonth) & "31) " & _
                     ")"
       
        '結餘點數
        'Modify by Amy 2020/05/14  公司別改下拉
        strBP = "Select a0102 as C00,Round(Sum(Decode(Substr(R003+19110000,1,6)," & lngThisMonth & "+191100,Nvl(R006,0)-Nvl(R005,0))),0) as C01, " & _
                        "Round(Sum(Decode(Substr(R003+19110000,1,6)," & lngThisMonth - 100 & "+191100,Nvl(R006,0)-Nvl(R005,0))),0) as C02, " & _
                        "Round(Sum(Decode(Substr(R003+19110000,1,6)," & lngThisMonth - 200 & "+191100,Nvl(R006,0)-Nvl(R005,0))),0) as C03,'' as C04,'' as C05, " & _
                        "Round(Sum(Decode(Substr(R003+19110000,1,4)," & lngYear & "+1911,Nvl(R006,0)-Nvl(R005,0))),0) as C06, " & _
                        "Round(Sum(Decode(Substr(R003+19110000,1,4)," & lngYear - 1 & "+1911,Nvl(R006,0)-Nvl(R005,0))),0) as C07, " & _
                        "Round(Sum(Decode(Substr(R003+19110000,1,4)," & lngYear - 2 & "+1911,Nvl(R006,0)-Nvl(R005,0))),0) as C08,'' as C09,'' as C10, R004,1 Sort " & _
                    "From Accrpt4202,acc010 " & _
                        "Where ID='" & strUserNum & "' And R004=a0101(+) " & strWhere & _
                        " And SubStr(R004, 1, 1) = '4' And Not( R004='4191' or R004='4192' or R004='4194') And InStr(R008||' ','結餘')>0 And InStr(R007,'轉撥')=0  " & _
                    "Group by R004,a0102 "
        '排除轉撥及結餘轉撥(因10601 D106012475 轉撥資料造成與frmacc44j0 當月實績不合)-婧瑄
        strBP = strBP & _
        " Union Select a0102 as C00,Round(Sum(Decode(Substr(R003+19110000,1,6)," & lngThisMonth & "+191100,Nvl(R006,0)-Nvl(R005,0))),0) as C01, " & _
                        "Round(Sum(Decode(Substr(R003+19110000,1,6)," & lngThisMonth - 100 & "+191100,Nvl(R006,0)-Nvl(R005,0))),0) as C02, " & _
                        "Round(Sum(Decode(Substr(R003+19110000,1,6)," & lngThisMonth - 200 & "+191100,Nvl(R006,0)-Nvl(R005,0))),0) as C03,'' as C04,'' as C05, " & _
                        "Round(Sum(Decode(Substr(R003+19110000,1,4)," & lngYear & "+1911,Nvl(R006,0)-Nvl(R005,0))),0) as C06, " & _
                        "Round(Sum(Decode(Substr(R003+19110000,1,4)," & lngYear - 1 & "+1911,Nvl(R006,0)-Nvl(R005,0))),0) as C07, " & _
                        "Round(Sum(Decode(Substr(R003+19110000,1,4)," & lngYear - 2 & "+1911,Nvl(R006,0)-Nvl(R005,0))),0) as C08,'' as C09,'' as C10, R004,2 Sort " & _
                    "From Accrpt4202,acc010 " & _
                        "Where ID='" & strUserNum & "' And R004=a0101(+) " & strWhere & _
                        " And SubStr(R004, 1, 1) = '4' And Not( R004='4191' or R004='4192' or R004='4194') And InStr(R007,'轉撥')>0 " & _
                    "Group by R004,a0102 "
        'end 2020/05/14
        
        strBP = "Select * From (" & strBP & ") Order by Sort,R004"
        intQ = 1
        Set RsQ = ClsLawReadRstMsg(intQ, strBP)
        If intQ = 1 Then
            m_lngRow = m_lngRow + 1
            wksaccrpt424.Range(Chr(intField) & m_lngRow).Value = "結餘點數："
            bolSheet2 = True: m_lngRow = m_lngRow + 1: StartRow = m_lngRow
            Do While Not RsQ.EOF
                'Sheet1 單純只列結餘(有轉撥字樣都不列)
                If RsQ.Fields("Sort") = 1 Then
                    For ii = 0 To UBound(strField)
                        bolFormaula = False: strTp(1) = "#,##0"
                        If ii > GetValue("") And InStr(strField(ii), "vs") > 0 Then
                            bolFormaula = True
                            If InStr(strField(ii), "年") > 0 Then
                                strTp(0) = "=" & Chr(intField + GetValue(Mid(strField(ii), 1, InStr(strField(ii), "年vs") - 1) & "年1-" & lngMonth & "月")) & m_lngRow & "/" & _
                                                         Chr(intField + GetValue(Mid(strField(ii), InStr(strField(ii), "年vs") + 4) & "年1-" & lngMonth & "月")) & m_lngRow & "-1"
                            Else
                                strTp(0) = "=" & Chr(intField + GetValue(Mid(strField(ii), 1, InStr(strField(ii), "vs") - 2) & "." & lngMonth)) & m_lngRow & "/" & _
                                                         Chr(intField + GetValue(Mid(strField(ii), InStr(strField(ii), "vs") + 3) & "." & lngMonth)) & m_lngRow & "-1"
                                                
                            End If
                            strTp(1) = "0.00%;[紅色]-0.00%"
                         '合計
                        ElseIf ii > GetValue("") And InStr(RsQ.Fields(11), "z") > 0 Then
                            '若尚未有期末資料公式會錯
                            If StartRow = m_lngRow Then
                                strTp(0) = "0"
                            Else
                                strTp(0) = "=Sum(" & Chr(intField + ii) & StartRow & ":" & Chr(intField + ii) & m_lngRow - 1 & ")"
                            End If
                        '內容
                        Else
                            strTp(0) = "" & RsQ.Fields(ii)
                        End If
                        If ii > GetValue("") And InStr(RsQ.Fields(11), "z") = 0 Then strTp(0) = Val(strTp(0))
                        If bolFormaula = True Then
                            wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Formula = strTp(0)
                        Else
                            wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = strTp(0)
                        End If
                        wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = strTp(1)
                        If InStr(RsQ.Fields(11), "z") > 0 Then
                            wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Font.Bold = True
                            wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Font.Color = &HFF& '紅字'設置儲存格填充色(膚)
                            wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.tintandshade = 0.5 '設深淺
                        End If
                    Next ii
                End If
                
                '更新結餘資料for 產生表二(結餘及轉撥都剔除,D105102692/D106063241 有輸錯後調整的結餘轉撥都需於表2剔除)
                strUpd = "": strA = ""
                '已更新過之特殊拆資料之會計科目只更新一次(因已含扣除結餘及轉撥資料)
                If InStr(strUpdAcc, "" & RsQ.Fields("R004")) = 0 Then
                    Select Case "" & RsQ.Fields("R004")
                        Case "410101" '商標收入-CCT
                            strA = "Select '110' RID,a0101,a0102 C00," & strF(2) & _
                                      " From Acc010, (" & GetCCT(3, "'" & RsQ.Fields("R004") & "'", False, stCon021, "ax205," & strF(0), True, True) & ") a" & _
                                      " Where a0101 ='" & RsQ.Fields("R004") & "' and ax205(+)=a0101" & _
                            " Union Select '111' RID,a0101,Replace(a0102||'-'||'MCT','CCT-') C00," & strF(2) & _
                                      " From Acc010, (" & GetCCT(1, "'" & RsQ.Fields("R004") & "'", False, stCon021, "ax205," & strF(0), True, True) & ") y " & _
                                      " Where a0101 ='" & RsQ.Fields("R004") & "' and ax205(+)=a0101" & _
                            " Union Select '112' RID,a0101,Replace(a0102||'-'||'MFCT','CCT-') C00," & strF(2) & _
                                      " From Acc010, (" & GetCCT(2, "'" & RsQ.Fields("R004") & "'", False, stCon021, "ax205," & strF(0), True, True) & ") y " & _
                                      " Where a0101 ='" & RsQ.Fields("R004") & "' and ax205(+)=a0101"
                        Case "410104" '商標收入-CCT爭議
                            strA = "Select '140' RID,a0101,a0102 C00," & strF(2) & _
                                      " From Acc010, (" & GetCCT(3, "'" & RsQ.Fields("R004") & "'", False, stCon021, "ax205," & strF(0), True, True) & ") y " & _
                                      " Where a0101 ='" & RsQ.Fields("R004") & "' and ax205(+)=a0101" & _
                            " Union Select '141' RID,a0101,Replace(a0102||'-'||'MCT','CCT-') C00," & strF(2) & _
                                      " From Acc010, (" & GetCCT(1, "'" & RsQ.Fields("R004") & "'", False, stCon021, "ax205," & strF(0), True, True) & ") y " & _
                                      " Where a0101 ='" & RsQ.Fields("R004") & "' and ax205(+)=a0101" & _
                            " Union Select '142' RID,a0101,Replace(a0102||'-'||'MFCT','CCT-') C00," & strF(2) & _
                                      " From Acc010, (" & GetCCT(2, "'" & RsQ.Fields("R004") & "'", False, stCon021, "ax205," & strF(0), True, True) & ") y " & _
                                      " Where a0101 ='" & RsQ.Fields("R004") & "' and ax205(+)=a0101"
                        Case "417201" 'FCT收入-國家
                            'Modify by Amy 2020/05/14 公司別改下拉 原:IIf(Text3 = "2", " and R001='J'", IIf(Text3 = "1", " and R001='1'", ""))
                            strA = "Select ax205, RID,a0101, " & strF(2) & _
                                      " From acc010,(select R004 as ax205,Decode(substr(nvl(R014,R012),1,3),'101','21','011','22','012','23',Decode(substr(nvl(R014,R012),1,1),'2','24','25')) RID," & _
                                                            Replace(Replace(Replace(strF(0), "a0205", "R003"), "ax206", "R005"), "ax207", "R006") & _
                                                " From Accrpt4202" & _
                                                " Where ID='" & strUserNum & "' And ((R003>=" & lngYear - 1 & "0101 and R003<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (R003>=" & lngYear - 2 & "0101 and R003<=" & lngThisMonth - 200 & "31) ", "") & _
                                                    "or (R003>=" & lngYear & "0101 and R003<=" & lngThisMonth & "31))" & strWhere & _
                                                    " and R004='417201' and (Instr(R007,'轉撥')>0 Or InStr(R008||' ','結餘')>0 ) " & _
                                                    " and ( ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 2) & ") " & _
                                                       " Or  ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 5) & ") " & _
                                                       " Or  (ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 3) & ") Or R009 is null)    )" & _
                                                " Group by R004,Decode(substr(nvl(R014,R012),1,3),'101','21','011','22','012','23',Decode(substr(nvl(R014,R012),1,1),'2','24','25'))  ) x " & _
                                        " Where a0101 ='" & RsQ.Fields("R004") & "' and ax205(+)=a0101"
                        Case "411101" '專利收入-CCP
                            strA = " Select '410' RID, a0101, a0102 C00," & strF(2) & _
                                      " From acc010, (" & GetCCP(3, False, stCon021, "ax205," & strF(0), True, True) & ") y " & _
                                      " Where a0101='411101' and ax205(+)=a0101" & _
                           " Union Select '411' RID,a0101,Replace(a0102||'-'||'MCP','CCP-') C00," & strF(2) & _
                                      " From acc010, ( " & GetCCP(1, False, stCon021, "ax205," & strF(0), True, True) & " ) y " & _
                                      " Where a0101='411101' and ax205(+)=a0101" & _
                            " Union Select '412' RID,a0101,Replace(a0102||'-'||'MFCP','CCP-') C00," & strF(2) & _
                                      " From acc010, ( " & GetCCP(2, False, stCon021, "ax205," & strF(0), True, True) & " ) y " & _
                                      " Where a0101='411101' and ax205(+)=a0101"
                        Case "417101", "417104", "417105", "417109" 'FCP收入-申請-國家
                            If Is417101 = False Then
                                'Modify by Amy 2020/05/14 公司別改下拉 原:IIf(Text3 = "2", " and R001='J'", IIf(Text3 = "1", " and R001='1'", ""))
                                strA = "Select ax205, RID,a0101," & strF(2) & _
                                          " From acc010,(select Decode(R004,'417104','417101','417105','417101','417109','417101',R004) as ax205,decode(substr(nvl(R014,R012),1,3),'101','51','011','52','012','53',decode(substr(nvl(R014,R012),1,1),'2','54','55')) RID," & _
                                                                Replace(Replace(Replace(strF(0), "a0205", "R003"), "ax206", "R005"), "ax207", "R006") & _
                                                    " From Accrpt4202" & _
                                                    " Where ID='" & strUserNum & "' And ((R003>=" & lngYear - 1 & "0101 and R003<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (R003>=" & lngYear - 2 & "0101 and R003<=" & lngThisMonth - 200 & "31) ", "") & _
                                                        "or (R003>=" & lngYear & "0101 and R003<=" & lngThisMonth & "31))" & _
                                                        " and DECODE(R004,'417104','417101','417105','417101','417109','417101',R004)='417101' and R010 is not null and (Instr(R007,'轉撥')>0 Or InStr(R008||' ','結餘')>0 ) " & strWhere & _
                                                        " and ( ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 1) & ") " & _
                                                           " Or  ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 5) & ") " & _
                                                           " Or  (ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 3) & ") Or R009 is null)    )" & _
                                                    " Group by Decode(R004,'417104','417101','417105','417101','417109','417101',R004),Decode(substr(nvl(R014,R012),1,3),'101','51','011','52','012','53',decode(substr(nvl(R014,R012),1,1),'2','54','55'))  ) x " & _
                                            " Where a0101 ='" & RsQ.Fields("R004") & "' and ax205(+)=a0101"
                            End If
                            Is417101 = True
                    End Select
                End If
                
                If strA <> MsgText(601) Then
                    strA = "Select * From (" & strA & ") Where c01<>0 or c03<>0 or c07<>0  or c04<>0 or c05<>0 or c06<>0"
                    intA = 1
                    Set rsA = ClsLawReadRstMsg(intA, strA)
                    If intA = 1 Then
                        rsA.MoveFirst
                        Do While Not rsA.EOF
                            strUpd = "Update Accrpt420 set R4204=Round(R4204,0)-(" & Val("" & rsA.Fields("C01")) & "),R4206=Round(R4206,0)-(" & Val("" & rsA.Fields("C03")) & ")" & _
                                        ",R4210=Round(R4210,0)-(" & Val("" & rsA.Fields("C07")) & "),R4207=Round(R4207,0)-(" & Val("" & rsA.Fields("C04")) & ")" & _
                                        ",R4208=Round(R4208,0)-(" & Val("" & rsA.Fields("C05")) & "),R4209=Round(R4209,0)-(" & Val("" & rsA.Fields("C06")) & ")" & _
                                  " Where ID='" & strUserNum & "' And R4211 is null And R4201='" & rsA.Fields("RID") & "' And R4202='" & RsQ.Fields("R004") & "' "
                            adoTaie.Execute strUpd
                            rsA.MoveNext
                        Loop
                        strUpdAcc = strUpdAcc & "," & RsQ.Fields("R004")
                    End If
                ElseIf InStr(strSpecAcc, "" & RsQ.Fields("R004")) = 0 Then
                    strUpd = " And R4203='" & RsQ.Fields("C00") & "' "
                    '與 frmacc44r0-ProduceData_Dept 共用 暫存檔,加R4211區分
                    strUpd = "Update Accrpt420 set R4204=Round(R4204,0)-(" & Val("" & RsQ.Fields(1)) & "),R4206=Round(R4206,0)-(" & Val("" & RsQ.Fields(2)) & ")" & _
                                        ",R4210=Round(R4210,0)-(" & Val("" & RsQ.Fields(3)) & "),R4207=Round(R4207,0)-(" & Val("" & RsQ.Fields(6)) & ")" & _
                                        ",R4208=Round(R4208,0)-(" & Val("" & RsQ.Fields(7)) & "),R4209=Round(R4209,0)-(" & Val("" & RsQ.Fields(8)) & ")" & _
                                  " Where ID='" & strUserNum & "' And R4211 is null And R4202='" & RsQ.Fields("R004") & "' " & strUpd
                    adoTaie.Execute strUpd
                End If
                
                If RsQ.Fields("Sort") = 1 Then m_lngRow = m_lngRow + 1
                RsQ.MoveNext
            Loop
            'Memo by Amy 「結餘點數不含結餘轉撥及實績轉撥」工作表之「專業達成總計 」=智權點數實績與結餘分析表 之 全所 「當月實績」(frmacc44j0)
            '備註
            wksaccrpt424.Range(Chr(intField) & m_lngRow + 2).Value = "結餘點數不含結餘轉撥及實績轉撥"
            wksaccrpt424.Range(Chr(intField) & m_lngRow + 2).Font.Color = &HFF& '紅字
            'end 2018/05/21
            Call ExcelHead2(strCmpN, True) 'Modify by Amy 2020/05/14
            intXlsSheet = intXlsSheet + 1
            Erase strTotal: Erase AllTotal
            GoTo NextSheet
        End If
    End If 'bolsheet2
    
    If bolSheet2 = True Then wksaccrpt424.Name = "不含結餘及轉撥點數"
    Call ExcelHead2(strCmpN, True) 'Modify by Amy 2020/05/14
    'Add by Amy 2021/08/24 +專業達成點數(比較三年)
    Call ProPointData
    intXlsSheet = intXlsSheet + 1
    Set wksaccrpt424 = xlsSalesPoint.Worksheets(strWkName & intXlsSheet)
    wksaccrpt424.Activate
    Call ExcelSave4
    'end 2021/08/24

    '框線
    If Val(xlsSalesPoint.Version) < 12 Then
         xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
    Else
         xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
    End If
    xlsSalesPoint.Workbooks.Close
    xlsSalesPoint.Quit
    Set wksaccrpt424 = Nothing
    Set xlsSalesPoint = Nothing

    StatusClear
    bolSheet2 = False
    MsgBox "Excel檔產生完畢！（" & strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43) & "）"
    Exit Sub
    
ErrHand:
    MsgBox Err.Description, , MsgText(5)
    If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
   End If
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   Set wksaccrpt424 = Nothing
   Set xlsSalesPoint = Nothing
   bolSheet2 = False
End Sub

'Add by Amy 2018/03/12 專業達成點數分佈情況(當月實際達成) 改新格式比較三年-婧瑄
Private Sub ExcelSave2_Old2()
'    Dim strTotal(1 To 10) As String, AllTotal(1 To 10) As String
'    Dim strTp(1) As String
'    Dim strWkName As String
'    Dim intXlsSheet As Integer
'    Dim bolFormaula As Boolean
'    Dim RsQ As ADODB.Recordset '結餘點數
'    Dim strBP As String, strUpd As String, intQ As Integer
'
'On Error GoTo ErrHand
'    If Dir(strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
'       If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'          MkDir strExcelPath
'       End If
'    Else
'       Kill strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43)
'    End If
'
'    SetPrinter
'    intXlsSheet = 1: intField = 65: bolSheet2 = False
'    xlsSalesPoint.SheetsInNewWorkbook = 2 'Added by Lydia 2019/03/13 預設工作表數量
'    xlsSalesPoint.Workbooks.add
'
'NextSheet:
'    If strWkName = MsgText(601) Then strWkName = Left(xlsSalesPoint.Worksheets(1).Name, Len(xlsSalesPoint.Worksheets(1).Name) - 1)
'    Set wksaccrpt424 = xlsSalesPoint.Worksheets(strWkName & intXlsSheet)
'    wksaccrpt424.Activate
'    IsOpenXls = True
'    'xlsSalesPoint.Visible = True
'    wksaccrpt424.PageSetup.PaperSize = 9 'A4
'    wksaccrpt424.PageSetup.Orientation = xlLandscape '橫印
'    wksaccrpt424.PageSetup.LeftMargin = 28.34
'    wksaccrpt424.PageSetup.RightMargin = 28.34
'    wksaccrpt424.PageSetup.TopMargin = 42.51
'    wksaccrpt424.PageSetup.BottomMargin = 42.51
'    wksaccrpt424.PageSetup.HeaderMargin = 28.34
'    wksaccrpt424.PageSetup.FooterMargin = 28.34
'
'    m_lngRow = 1
'    Call ExcelHead2
'
'    intTitleR = m_lngRow - 1: StartRow = m_lngRow
'    '產生表二,重抓暫存檔資料
'    If bolSheet2 = True Then
'        'Modify by Amy 2019/05/14 與 frmacc44r0-ProduceData_Dept 共用 暫存檔,加R4211區分
'        strSql = "Select r4203 as C00,Round(r4204,0) as C01,Round(r4206,0) as C02,Round(R4210,0) as C03,'' as C04,'' as C05" & _
'                        ",Round(r4207,0) as C06,Round(r4208,0) as C07,Round(R4209,0) as C08,'' as C09,'' as C10, r4201 as RID,r4202 as A0101 " & _
'                    "From Accrpt420 Where ID='" & strUserNum & "' And R4211 is null Order by r4201"
'        intQ = 1
'        Set rsNew = ClsLawReadRstMsg(intQ, strSql)
'    End If
'    rsNew.MoveFirst
'    Do While rsNew.EOF = False
'        For ii = 0 To UBound(strField)
'            '比率
'            If ii > GetValue("") And InStr(strField(ii), "vs") > 0 Then
'                If InStr(strField(ii), "年") > 0 Then
'                    strTp(0) = Chr(intField + GetValue(Mid(strField(ii), 1, InStr(strField(ii), "年vs") - 1) & "年1-" & lngMonth & "月")) & m_lngRow & "/" & _
'                                    Chr(intField + GetValue(Mid(strField(ii), InStr(strField(ii), "年vs") + 4) & "年1-" & lngMonth & "月")) & m_lngRow & "-1"
'                Else
'                    strTp(0) = Chr(intField + GetValue(Mid(strField(ii), 1, InStr(strField(ii), "vs") - 2) & "." & lngMonth)) & m_lngRow & "/" & _
'                                    Chr(intField + GetValue(Mid(strField(ii), InStr(strField(ii), "vs") + 3) & "." & lngMonth)) & m_lngRow & "-1"
'                End If
'                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Formula = "=" & strTp(0)
'                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = "0.00%;[紅色]-0.00%"
'                If InStr(rsNew.Fields("RID"), "z") > 0 And "" & rsNew.Fields("RID") <> "6z1" Then
'                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Font.Bold = True
'                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.ColorIndex = 40 '設置儲存格填充色(膚)
'                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.tintandshade = 0.5 '設深淺
'                End If
'            '合計
'            ElseIf ii > GetValue("") And InStr(rsNew.Fields("RID"), "z") > 0 And "" & rsNew.Fields("RID") <> "6z1" Then
'                Select Case rsNew.Fields("RID")
'                    Case "1z", "2z", "3z", "4z", "5z" '合計
'                        strTp(0) = "Sum(" & Chr(intField + ii) & StartRow & ":" & Chr(intField + ii) & m_lngRow - 1 & ")"
'                        strTotal(ii) = strTotal(ii) & Chr(intField + ii) & m_lngRow & ","
'                    Case "3zz", "6zz" '達成總計
'                        strTp(0) = "Sum(" & Left(strTotal(ii), Len(strTotal(ii)) - 1) & ")"
'                        AllTotal(ii) = AllTotal(ii) & Chr(intField + ii) & m_lngRow & ","
'                        strTotal(ii) = ""
'                    Case "7zt1" '專業達成總計
'                        strTp(0) = "Sum(" & AllTotal(ii) & "Sum(" & Chr(intField + ii) & StartRow & ":" & Chr(intField + ii) & m_lngRow - 1 & ")" & ")"
'                End Select
'                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Formula = "=" & strTp(0)
'                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = "#,##0"
'                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Font.Bold = True
'                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.ColorIndex = 40 '設置儲存格填充色(膚)
'                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.tintandshade = 0.5 '設深淺
'            '內容
'            Else
'                If ii > GetValue("") And "" & rsNew.Fields("RID") = "6z1" Then strTotal(ii) = strTotal(ii) & Chr(intField + ii) & m_lngRow & ","
'                strTp(0) = "" & rsNew.Fields(ii)
'                strTp(1) = "#,##0"
'                If ii = GetValue("") Then
'                    strTp(1) = ""
'                Else
'                    strTp(0) = Val(strTp(0))
'                End If
'                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = strTp(0)
'                If ii = GetValue(strField(0)) And InStr(rsNew.Fields("RID"), "z") > 0 And InStr(rsNew.Fields("RID"), "z1") = 0 Then
'                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Font.Bold = True
'                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.ColorIndex = 40 '設置儲存格填充色(膚)
'                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.tintandshade = 0.5 '設深淺
'                Else
'                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = strTp(1)
'                End If
'            End If
'        Next ii
'        m_lngRow = m_lngRow + 1
'        If InStr(rsNew.Fields("RID"), "z") > 0 And InStr(rsNew.Fields("RID"), "z1") = 0 Then StartRow = m_lngRow
'        rsNew.MoveNext
'    Loop
'
'    If bolSheet2 = False Then
'    '結餘點數
'    'Modify by Amy 2018/05/21 拿掉科目And ax205 in('410103','411103','412101','413101') -婧瑄
'    'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:SubStr(ax205, 1, 2) = '41'
'    strBP = "Select a0102 as C00,Round(Sum(Decode(Substr(a0205+19110000,1,6)," & lngThisMonth & "+191100,Nvl(ax207,0)-Nvl(ax206,0))),0) as C01, " & _
'                    "Round(Sum(Decode(Substr(a0205+19110000,1,6)," & lngThisMonth - 100 & "+191100,Nvl(ax207,0)-Nvl(ax206,0))),0) as C02, " & _
'                    "Round(Sum(Decode(Substr(a0205+19110000,1,6)," & lngThisMonth - 200 & "+191100,Nvl(ax207,0)-Nvl(ax206,0))),0) as C03,'' as C04,'' as C05, " & _
'                    "Round(Sum(Decode(Substr(a0205+19110000,1,4)," & lngYear & "+1911,Nvl(ax207,0)-Nvl(ax206,0))),0) as C06, " & _
'                    "Round(Sum(Decode(Substr(a0205+19110000,1,4)," & lngYear - 1 & "+1911,Nvl(ax207,0)-Nvl(ax206,0))),0) as C07, " & _
'                    "Round(Sum(Decode(Substr(a0205+19110000,1,4)," & lngYear - 2 & "+1911,Nvl(ax207,0)-Nvl(ax206,0))),0) as C08,'' as C09,'' as C10, ax205,1 Sort " & _
'                "From acc020, acc021,acc010 " & _
'                    "Where ax201 = a0201(+)  And ax202 = a0202(+) " & IIf(Text3 = "2", " and ax201='J'", IIf(Text3 = "1", " and ax201='1'", "")) & _
'                    " And  ax205=a0101(+) " & IIf(txtAccNo <> "", "And ax205='" & txtAccNo & "'", "") & _
'                    " And ((a0205>=" & lngYear & "0101 And a0205<=" & lngYear & IIf(lngMonth < 10, "0" & lngMonth, lngMonth) & "31) " & _
'                    " Or (a0205>=" & lngYear - 1 & "0101 And a0205<=" & lngYear - 1 & IIf(lngMonth < 10, "0" & lngMonth, lngMonth) & "31)" & _
'                    " Or (a0205>=" & lngYear - 2 & "0101 And a0205<=" & lngYear - 2 & IIf(lngMonth < 10, "0" & lngMonth, lngMonth) & "31) )" & _
'                    " And SubStr(ax205, 1, 1) = '4' And Not( ax205='4191' or ax205='4192' or ax205='4194') And InStr(ax213||' ','結餘')>0 And InStr(ax212,'轉撥')=0  " & _
'                "Group by ax205,a0102 "
'    '排除轉撥及結餘轉撥(因10601 D106012475 轉撥資料造成與frmacc44j0 當月實績不合)-婧瑄
'    strBP = strBP & _
'    " Union Select a0102 as C00,Round(Sum(Decode(Substr(a0205+19110000,1,6)," & lngThisMonth & "+191100,Nvl(ax207,0)-Nvl(ax206,0))),0) as C01, " & _
'                    "Round(Sum(Decode(Substr(a0205+19110000,1,6)," & lngThisMonth - 100 & "+191100,Nvl(ax207,0)-Nvl(ax206,0))),0) as C02, " & _
'                    "Round(Sum(Decode(Substr(a0205+19110000,1,6)," & lngThisMonth - 200 & "+191100,Nvl(ax207,0)-Nvl(ax206,0))),0) as C03,'' as C04,'' as C05, " & _
'                    "Round(Sum(Decode(Substr(a0205+19110000,1,4)," & lngYear & "+1911,Nvl(ax207,0)-Nvl(ax206,0))),0) as C06, " & _
'                    "Round(Sum(Decode(Substr(a0205+19110000,1,4)," & lngYear - 1 & "+1911,Nvl(ax207,0)-Nvl(ax206,0))),0) as C07, " & _
'                    "Round(Sum(Decode(Substr(a0205+19110000,1,4)," & lngYear - 2 & "+1911,Nvl(ax207,0)-Nvl(ax206,0))),0) as C08,'' as C09,'' as C10, ax205,2 Sort " & _
'                "From acc020, acc021,acc010 " & _
'                    "Where ax201 = a0201(+)  And ax202 = a0202(+) " & IIf(Text3 = "2", " and ax201='J'", IIf(Text3 = "1", " and ax201='1'", "")) & _
'                    " And  ax205=a0101(+) " & IIf(txtAccNo <> "", "And ax205='" & txtAccNo & "'", "") & _
'                    " And ((a0205>=" & lngYear & "0101 And a0205<=" & lngYear & IIf(lngMonth < 10, "0" & lngMonth, lngMonth) & "31) " & _
'                    " Or (a0205>=" & lngYear - 1 & "0101 And a0205<=" & lngYear - 1 & IIf(lngMonth < 10, "0" & lngMonth, lngMonth) & "31)" & _
'                    " Or (a0205>=" & lngYear - 2 & "0101 And a0205<=" & lngYear - 2 & IIf(lngMonth < 10, "0" & lngMonth, lngMonth) & "31) )" & _
'                    " And SubStr(ax205, 1, 1) = '4' And Not( ax205='4191' or ax205='4192' or ax205='4194') And InStr(ax212,'轉撥')>0 " & _
'                "Group by ax205,a0102 "
'    'end 2019/08/01
'    strBP = "Select * From (" & strBP & ") Order by Sort,ax205"
'    intQ = 1
'    Set RsQ = ClsLawReadRstMsg(intQ, strBP)
'    If intQ = 1 Then
'        m_lngRow = m_lngRow + 1
'        wksaccrpt424.Range(Chr(intField) & m_lngRow).Value = "結餘點數："
'        bolSheet2 = True: m_lngRow = m_lngRow + 1: StartRow = m_lngRow
'        Do While Not RsQ.EOF
'            'Sheet1 單純只列結餘(有轉撥字樣都不列)
'            If RsQ.Fields("Sort") = 1 Then
'                For ii = 0 To UBound(strField)
'                    bolFormaula = False: strTp(1) = "#,##0"
'                    If ii > GetValue("") And InStr(strField(ii), "vs") > 0 Then
'                        bolFormaula = True
'                        If InStr(strField(ii), "年") > 0 Then
'                            strTp(0) = "=" & Chr(intField + GetValue(Mid(strField(ii), 1, InStr(strField(ii), "年vs") - 1) & "年1-" & lngMonth & "月")) & m_lngRow & "/" & _
'                                                     Chr(intField + GetValue(Mid(strField(ii), InStr(strField(ii), "年vs") + 4) & "年1-" & lngMonth & "月")) & m_lngRow & "-1"
'                        Else
'                            strTp(0) = "=" & Chr(intField + GetValue(Mid(strField(ii), 1, InStr(strField(ii), "vs") - 2) & "." & lngMonth)) & m_lngRow & "/" & _
'                                                     Chr(intField + GetValue(Mid(strField(ii), InStr(strField(ii), "vs") + 3) & "." & lngMonth)) & m_lngRow & "-1"
'
'                        End If
'                        strTp(1) = "0.00%;[紅色]-0.00%"
'                     '合計
'                    ElseIf ii > GetValue("") And InStr(RsQ.Fields(11), "z") > 0 Then
'                        'Add by Amy 2018/04/03 +if 若尚未有期末資料公式會錯
'                        If StartRow = m_lngRow Then
'                            strTp(0) = "0"
'                        Else
'                            strTp(0) = "=Sum(" & Chr(intField + ii) & StartRow & ":" & Chr(intField + ii) & m_lngRow - 1 & ")"
'                        End If
'                    '內容
'                    Else
'                        strTp(0) = "" & RsQ.Fields(ii)
'                    End If
'                    If ii > GetValue("") And InStr(RsQ.Fields(11), "z") = 0 Then strTp(0) = Val(strTp(0))
'                    If bolFormaula = True Then
'                        wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Formula = strTp(0)
'                    Else
'                        wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = strTp(0)
'                    End If
'                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = strTp(1)
'                    If InStr(RsQ.Fields(11), "z") > 0 Then
'                        wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Font.Bold = True
'                        wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Font.Color = &HFF& '紅字'設置儲存格填充色(膚)
'                        wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.tintandshade = 0.5 '設深淺
'                    End If
'                Next ii
'            End If
'
'            '更新結餘資料for 產生表二(結餘及轉撥都剔除,D105102692/D106063241 有輸錯後調整的結餘轉撥都需於表2剔除)
'            strUpd = ""
'            If RsQ.Fields("Sort") = 2 Then
'                strUpd = " And R4203='" & RsQ.Fields("C00") & "' "
'            End If
'            'Modify by Amy 2019/05/14 與 frmacc44r0-ProduceData_Dept 共用 暫存檔,加R4211區分
'            strUpd = "Update Accrpt420 set R4204=Round(R4204,0)-(" & Val("" & RsQ.Fields(1)) & "),R4206=Round(R4206,0)-(" & Val("" & RsQ.Fields(2)) & ")" & _
'                                ",R4210=Round(R4210,0)-(" & Val("" & RsQ.Fields(3)) & "),R4207=Round(R4207,0)-(" & Val("" & RsQ.Fields(6)) & ")" & _
'                                ",R4208=Round(R4208,0)-(" & Val("" & RsQ.Fields(7)) & "),R4209=Round(R4209,0)-(" & Val("" & RsQ.Fields(8)) & ")" & _
'                          " Where ID='" & strUserNum & "' And R4211 is null And R4202='" & RsQ.Fields("ax205") & "' " & strUpd
'            adoTaie.Execute strUpd
'
'            If RsQ.Fields("Sort") = 1 Then m_lngRow = m_lngRow + 1
'            RsQ.MoveNext
'        Loop
'        '備註
'        wksaccrpt424.Range(Chr(intField) & m_lngRow + 2).Value = "結餘點數不含結餘轉撥及實績轉撥"
'        wksaccrpt424.Range(Chr(intField) & m_lngRow + 2).Font.Color = &HFF& '紅字
'        'end 2018/05/21
'        Call ExcelHead2(True)
'        intXlsSheet = intXlsSheet + 1
'        Erase strTotal: Erase AllTotal
'        GoTo NextSheet
'    End If
'    End If 'bolsheet2
'
'    If bolSheet2 = True Then wksaccrpt424.Name = "不含結餘及轉撥點數"
'    Call ExcelHead2(True)
'
'    '框線
'    If Val(xlsSalesPoint.Version) < 12 Then
'         xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
'    Else
'         xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
'    End If
'    xlsSalesPoint.Workbooks.Close
'    xlsSalesPoint.Quit
'    Set wksaccrpt424 = Nothing
'    Set xlsSalesPoint = Nothing
'
'    StatusClear
'    bolSheet2 = False
'    MsgBox "Excel檔產生完畢！（" & strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43) & "）"
'    Exit Sub
'
'ErrHand:
'    MsgBox Err.Description, , MsgText(5)
'    If Val(xlsSalesPoint.Version) < 12 Then
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
'   Else
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
'   End If
'   xlsSalesPoint.Workbooks.Close
'   xlsSalesPoint.Quit
'   Set wksaccrpt424 = Nothing
'   Set xlsSalesPoint = Nothing
'   bolSheet2 = False
End Sub

'Modify by Amy 2020/05/14 +strCmpN 公司名稱
Private Sub ExcelHead2(strCmpN As String, Optional bolLast As Boolean = False)
    Dim stTp(1) As String
    
    If bolLast = False Then
        If bolSheet2 = False Then
            'Modify by Amy 2020/06/05 if 國家別點數分析表增加各國佔比欄
            If Combo2 = "國家別點數分析表" Then
                ReDim strField(0 To 13)
                ReDim intWidth(0 To 13)
            Else
                ReDim strField(0 To 10)
                ReDim intWidth(0 To 10)
            End If
            
            strField(0) = "": intWidth(0) = 15
            For ii = 1 To 3
                strField(ii) = Val(lngYear) + 1 - ii & "." & lngMonth
                intWidth(ii) = 12
                strField(ii + 5) = Val(lngYear) + 1 - ii & "年1-" & lngMonth & "月"
                intWidth(ii + 5) = 12
                'Add by Amy 2020/06/05 國家別點數分析表增加各國佔比欄
                If Combo2 = "國家別點數分析表" Then
                    strField(ii + 10) = Val(lngYear) + 1 - ii & "年1-" & lngMonth & "月各國佔比"
                    intWidth(ii + 10) = 12
                End If
            Next ii
            For ii = 4 To 5
                '以基準年(lngYear)比較
                strField(ii) = Val(lngYear) & " vs " & Val(lngYear) + 3 - ii
                intWidth(ii) = 10
                strField(ii + 5) = Val(lngYear) & "年vs " & Val(lngYear) + 3 - ii
                intWidth(ii + 5) = 10
            Next ii
        End If
        
        wksaccrpt424.Range(Chr(intField) & m_lngRow).Value = lngYear & "年" & lngMonth & "月份各" & Combo2 & IIf(bolSheet2 = True, "-不含結餘點數", "")
        wksaccrpt424.Range(Chr(intField) & m_lngRow & ":" & Chr(intField + UBound(strField)) & m_lngRow).Select
        With xlsSalesPoint.Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = True
        End With
        wksaccrpt424.Application.Selection.Font.Bold = True
        wksaccrpt424.Application.Selection.Font.Size = 14
        m_lngRow = m_lngRow + 1
        
        'Modify by Amy 2020/05/14 公司別改抓變數 原: IIf(Text3 = "2", "智權", IIf(Text3 = "1", "台一", "台一　專利商標/智權"))
        wksaccrpt424.Range(Chr(intField) & m_lngRow).Value = "公司別：" & strCmpN
        wksaccrpt424.Range(Chr(intField + UBound(strField) - 2) & m_lngRow).Value = "列印日期：" & CFDate(strSrvDate(2))
        m_lngRow = m_lngRow + 1
   
        For ii = LBound(strField) To UBound(strField)
            wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = "@"
            wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = strField(ii)
            wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).HorizontalAlignment = xlCenter
            wksaccrpt424.Columns(Chr(intField + ii)).ColumnWidth = intWidth(ii)
        Next ii
        m_lngRow = m_lngRow + 1
    'bolLast=true
    Else
        For ii = GetValue("" & strField(9)) To UBound(strField)
            wksaccrpt424.Range(Chr(intField + ii) & intTitleR).Value = Replace(strField(ii), "年vs", " vs")
            'Add by Amy 2020/06/05 國家別點數分析表增加各國佔比欄
            If Combo2 = "國家別點數分析表" Then
                If InStr(strField(ii), "各國佔比") > 0 Then
                    wksaccrpt424.Range(Chr(intField + ii) & intTitleR).Value = Replace(strField(ii), "各國佔比", " ")
                    wksaccrpt424.Range(Chr(intField + ii) & intTitleR + 1).Value = "各國佔比"
                    wksaccrpt424.Range(Chr(intField + ii) & intTitleR + 1).HorizontalAlignment = xlCenter
                End If
            End If
            'end 2020/06/05
        Next ii
        
        wksaccrpt424.Range(Chr(intField) & intTitleR & ":" & Chr(intField + UBound(strField)) & m_lngRow - 1).Select
        xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Font.Size = 11
        wksaccrpt424.PageSetup.PrintTitleRows = "$1:$" & intTitleR '表頭保留列
        wksaccrpt424.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
    End If
End Sub

Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = 1 To UBound(strField)
       If UCase(strField(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

Private Sub UpdateTotal()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strUpd As String
    Dim intQ As Integer, ii As Integer
    Dim Sum5(1 To 5) As Double, TPSum(1 To 5) As Double
    
    '商標國內專業合計/FCT收入合計/專利國內專業合計/FCP收入合計/*專業達成總計/*全所合計
    'Modify by Amy 2019/05/14 與 frmacc44r0-ProduceData_Dept 共用 暫存檔,加R4211區分
    strQ = "Select SubStr(r4201,1,1) as RID,Sum(Nvl(r4204,0)) as C01,Sum(Nvl(r4205,0)) as C02,Sum(Nvl(r4206,0)) as C03,Sum(Nvl(r4207,0)) as C04,Sum(Nvl(r4208,0)) as C05 " & _
               "From Accrpt420 Where ID='" & strUserNum & "' And R4211 is null And SubStr(r4201,2,1)<>'z' and SubStr(r4201,2,2)<>'z1' Group by SubStr(r4201,1,1)"
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        With RsQ
            Do While .EOF = False
                'Modify by Amy 2019/05/14 與 frmacc44r0-ProduceData_Dept 共用 暫存檔,加R4211區分
                strUpd = "Update Accrpt420 Set r4204=" & .Fields("C01") & ",r4205=" & .Fields("C02") & ",r4206=" & .Fields("C03") & ",r4207=" & .Fields("C04") & ",r4208=" & .Fields("C05") & _
                              " Where ID='" & strUserNum & "' And R4211 is null And SubStr(r4201,1,2)='" & .Fields("RID") & "z' And SubStr(r4201,2,2)<>'z1' And SubStr(r4201,2,2)<>'zz'"
                cnnConnection.Execute strUpd
                If Left(.Fields("RID"), 1) = "7" Then
                    For ii = 1 To 5
                        Sum5(ii) = Sum5(ii) + Val(RsQ.Fields("C0" & ii))
                    Next ii
                End If
                .MoveNext
            Loop
        End With
    End If
    RsQ.Close
    '商標/專利達成總計
    'Modify by Amy 2019/05/14 與 frmacc44r0-ProduceData_Dept 共用 暫存檔,加R4211區分
     strQ = "Select Decode(Sign(SubStr(r4201,1,1)-4),-1,'3','6') as RID,Sum(Nvl(r4204,0)) as C01,Sum(Nvl(r4205,0)) as C02,Sum(Nvl(r4206,0)) as C03,Sum(Nvl(r4207,0)) as C04,Sum(Nvl(r4208,0)) as C05 " & _
               "From Accrpt420 Where ID='" & strUserNum & "' And R4211 is null And SubStr(r4201,2,2)<>'zz' And SubStr(r4201,2,1)='z' And SubStr(r4201,1,1)<7 Group by Decode(Sign(SubStr(r4201,1,1)-4),-1,'3','6')"

    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        With RsQ
            Do While .EOF = False
                'Modify by Amy 2019/05/14 與 frmacc44r0-ProduceData_Dept 共用 暫存檔,加R4211區分
                strUpd = "Update Accrpt420 Set r4204=" & .Fields("C01") & ",r4205=" & .Fields("C02") & ",r4206=" & .Fields("C03") & ",r4207=" & .Fields("C04") & ",r4208=" & .Fields("C05") & _
                                      " Where ID='" & strUserNum & "' And R4211 is null And SubStr(r4201,1,3)='" & .Fields("RID") & "zz'"
                cnnConnection.Execute strUpd
                For ii = 1 To 5
                    TPSum(ii) = TPSum(ii) + Val(RsQ.Fields("C0" & ii))
                Next ii
                .MoveNext
            Loop
        End With
    End If
    RsQ.Close
    'Modify by Amy 2019/05/14 與 frmacc44r0-ProduceData_Dept 共用 暫存檔,加R4211區分
    '專業達成總計
     strUpd = "Update Accrpt420 Set r4204=r4204+" & TPSum(1) & ",r4205=r4205+" & TPSum(2) & ",r4206=r4206+" & TPSum(3) & ",r4207=r4207+" & TPSum(4) & ",r4208=r4208+" & TPSum(5) & _
                    " Where ID='" & strUserNum & "' And R4211 is null And SubStr(r4201,1,3)='7zt' "
    cnnConnection.Execute strUpd
    '全所合計
     strUpd = "Update Accrpt420 Set r4204=r4204+" & Sum5(1) & "+" & TPSum(1) & ",r4205=r4205+" & Sum5(2) & "+" & TPSum(2) & ",r4206=r4206+" & Sum5(3) & "+" & TPSum(3) & ",r4207=r4207+" & Sum5(4) & "+" & TPSum(4) & _
                    ",r4208=r4208+" & Sum5(5) & "+" & TPSum(5) & " Where ID='" & strUserNum & "' And R4211 is null And SubStr(r4201,1,3)='8zt' "
    cnnConnection.Execute strUpd
    
    '重抓資料
    strQ = "Select r4201 as RID,r4202 as A0101,r4203 as C00,r4204 as C01,r4205 as C02,r4206 as C03,r4207 as C04,r4208 as C05 " & _
                "From Accrpt420 Where ID='" & strUserNum & "' And R4211 is null Order by r4201"
    'end 2019/05/14
    Set rsNew = ClsLawReadRstMsg(intI, strQ)
End Sub
'end 2015/06/16

'Added by Lydia 2016/02/22 格線高度
Private Sub GetTBRows(ByRef pRst As ADODB.Recordset, ByRef intN As Integer, ByRef intR As Integer)
Dim iTBr As Integer
   
  iTBr = iTBHeight + 1 '加抬頭列
  '第一頁
  If intN < 1 Then
     If pRst.RecordCount > cRows Then
        intR = cRows + iTBr
     Else
        intR = pRst.RecordCount + iTBr
     End If
  '第二頁~
  Else
     If pRst.RecordCount - intN + 1 > cRows Then
        intR = cRows + iTBr
     Else
        intR = pRst.RecordCount - intN + 1 + iTBr
     End If
  End If
End Sub

'Add by Amy 2019/02/18 取得商標收入410101/410104
'intChoose:0-有代理人(MCT+MFCT)/1-MCT(有代理人且客戶國籍為大陸)/2-MFCT(有代理人且客戶國籍非大陸)/3-沒代理人
'bolBefMon:跑一月資料要加讀前一年12月資料
'Modify by Amy 2019/12/19 +bolTmpTB-抓暫存檔/IsSurTran-只抓結餘及轉撥 參數
Private Function GetCCT(ByVal intChoose As Integer, ByVal stAx205 As String, ByVal bolBefMon As Boolean, stCon021 As String, _
                            Optional ByVal stField As String = "", Optional ByVal bolTmpTB As Boolean = False, Optional ByVal IsSurTran As Boolean = False) As String
    Dim stQ As String, stTmp(1) As String, stTmp2 As String
    Dim stF As String 'Add by Amy 2019/12/19
    
    'Modify by Amy 2019/12/19
    If bolTmpTB = False Then
        stF = "a0205"
        If intChoose = 3 Then
            stTmp(0) = "and tm44 is null "
            stTmp(1) = "and sp26 is null "
        Else
            stTmp(0) = "and tm44 is not null "
            stTmp(1) = "and sp26 is not null "
        End If
    Else
        stF = "R003"
        If intChoose = 3 Then
            stTmp(0) = "and R013 is null "
            stTmp(1) = "and R013 is null "
        Else
            stTmp(0) = "and R013 is not null "
            stTmp(1) = "and R013 is not null "
        End If
    End If
    
    '用Nvl(cu10,fa10) 因TS001663 無申請人
    If intChoose = 1 Then
        If bolTmpTB = False Then
            stTmp(0) = stTmp(0) & "and Nvl(cu10,fa10)='020' "
            stTmp(1) = stTmp(1) & "and Nvl(cu10,fa10)='020' "
        Else
            stTmp(0) = stTmp(0) & "and Nvl(R012,R014)='020' "
            stTmp(1) = stTmp(1) & "and Nvl(R012,R014)='020' "
        End If
    ElseIf intChoose = 2 Then
        If bolTmpTB = False Then
            stTmp(0) = stTmp(0) & "and Nvl(cu10,fa10)<>'020' "
            stTmp(1) = stTmp(1) & "and Nvl(cu10,fa10)<>'020' "
        Else
            stTmp(0) = stTmp(0) & "and Nvl(R012,R014)<>'020' "
            stTmp(1) = stTmp(1) & "and Nvl(R012,R014)<>'020' "
        End If
    End If
    If bolTmpTB = False Then
        stTmp(0) = "and fa01(+)=substr(tm44,1,8) and fa02(+)=substr(tm44,9,1)" & _
                          "and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9,1) " & stTmp(0)
        stTmp(1) = "and fa01(+)=substr(sp26,1,8) and fa02(+)=substr(sp26,9,1)" & _
                          "and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9,1) " & stTmp(1)
    End If
    
    'Modify by Amy 2019/12/19
    stTmp2 = " And ((" & stF & ">=" & lngYear - 1 & "0101 and " & stF & "<=" & lngThisMonth - 100 & "31) " & _
    IIf(bolArrive = True, "Or (" & stF & ">=" & lngYear - 2 & "0101 and " & stF & "<=" & lngThisMonth - 200 & "31)", "") & _
                                     "Or (" & stF & ">=" & lngYear & "0101 and " & stF & "<=" & lngThisMonth & "31) ) "
    '跑一月資料要加讀前一年12月資料
    If bolBefMon = True Then
        stTmp2 = " And " & stF & ">=" & lngLastMonth & "01 and " & stF & "<=" & lngLastMonth & "31 "
    End If
    'Add by Amy 2019/12/19 IsSurTran只抓結餘和轉撥
    If IsSurTran = True Then
        If bolTmpTB = False Then
            stTmp2 = stTmp2 & " And (Instr(ax212,'轉撥')>0 Or InStr(ax213||' ','結餘')>0 ) "
        Else
            stTmp2 = stTmp2 & " And (Instr(R007,'轉撥')>0 Or InStr(R008||' ','結餘')>0 ) "
        End If
    End If
        
    If stField = MsgText(601) Then
        stQ = "select ax205, sum(decode(floor(a0205/100)," & lngThisMonth & ",ax206)) Sd1" & _
                                     ", sum(decode(floor(a0205/100)," & lngThisMonth & ",ax207)) Sc1" & _
                                     ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2" & _
                                     ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2" & _
                                     ", sum(decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax206)) Sd3" & _
                                     ", sum(decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax207)) Sc3" & _
                                     ", sum(decode(floor(a0205/10000)," & lngYear & ",ax206)) Sd4" & _
                                     ", sum(decode(floor(a0205/10000)," & lngYear & ",ax207)) Sc4" & _
                                     ", sum(decode(floor(a0205/10000)," & lngYear - 1 & ",ax206)) Sd5" & _
                                     ", sum(decode(floor(a0205/10000)," & lngYear - 1 & ",ax207)) Sc5"
        If bolBefMon = True Then
            stQ = "select ax205, sum(decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2" & _
                                         ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2"
        End If
    Else
        stQ = "Select " & stField
    End If
    'end 2019/12/19
    
    'Modify by Amy 2020/05/14 拿掉公司別由外層傳入
    If bolTmpTB = False Then
        GetCCT = "select a0205, ax205, ax206, ax207 from acc020, acc021, trademark, fagent, customer" & _
            " where ax201(+)=a0201 and ax202(+)=a0202 and ax205 IN (" & stAx205 & ") and ax209 is not null" & _
              stCon021 & stTmp2 & stTmp(0) & _
            " and tm01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and tm02(+)=substr(lpad(ax214,12,' '),4,6)" & _
            " and tm03(+)=substr(lpad(ax214,12,' '),10,1) and tm04(+)=substr(lpad(ax214,12,' '),11,2) and tm01 is not null" & _
            " Union All" & _
            " select a0205, ax205, ax206, ax207 from acc020, acc021, servicepractice, fagent, customer" & _
            " where  ax201(+)=a0201 and ax202(+)=a0202 and ax205 IN (" & stAx205 & ") and ax209 is not null" & _
              stCon021 & stTmp2 & stTmp(1) & _
            " and sp01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and sp02(+)=substr(lpad(ax214,12,' '),4,6)" & _
            " and sp03(+)=substr(lpad(ax214,12,' '),10,1) and sp04(+)=substr(lpad(ax214,12,' '),11,2) and (sp01 is not null or ax214 is null)"
    '暫存檔
    Else
         GetCCT = "Select R003 as a0205,R004 as ax205, R005 as ax206,R006 as ax207 From Accrpt4202" & _
            " Where ID='" & strUserNum & "' And R004 IN (" & stAx205 & ") and R010 is not null " & _
            " and ( (ltrim(substr(lpad(R009,12,' '),1,3)) in(" & SQLGrpStr(GetAllSysKind(, "ALL"), 2) & ") and R009 is not null) " & _
                 "Or (ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 5) & ") or R009 is null)  ) " & _
             Replace(Replace(stCon021, " ax205 ", " R004 "), " ax201 ", " R001 ") & stTmp2 & stTmp(0)
    End If
    'end 2020/05/14
        
    GetCCT = stQ & " From (" & GetCCT & ") x Group By ax205"
    'end 2019/12/19
End Function

'intChoose:0-有代理人(MCP+MFCP)/1-MCP(有代理人且客戶國籍為大陸)/2-MFCP(有代理人且客戶國籍非大陸)/3-沒有代理人
'bolBefMon:跑一月資料要加讀前一年12月資料
'Modify by Amy 2019/12/19 +bolTmpTB 抓暫存檔/IsSurTran-只抓結餘及轉撥 參數
Private Function GetCCP(ByVal intChoose As Integer, ByVal bolBefMon As Boolean, stCon021 As String, _
                            Optional ByVal stField As String = "", Optional ByVal bolTmpTB As Boolean = False, Optional ByVal IsSurTran As Boolean = False) As String
    Dim stQ As String, stTmp(1) As String, stTmp2 As String
    Dim stF As String 'Add by Amy 2019/12/19
    
    'Modify by Amy 2019/12/19
    If bolTmpTB = False Then
        stF = "a0205"
        If intChoose = 3 Then
            stTmp(0) = "and pa75 is null "
            stTmp(1) = "and sp26 is null "
        Else
            stTmp(0) = "and pa75 is not null "
            stTmp(1) = "and sp26 is not null "
        End If
    Else
        stF = "R003"
        If intChoose = 3 Then
            stTmp(0) = "and R013 is null "
            stTmp(1) = "and R013 is null "
        Else
            stTmp(0) = "and R013 is not null "
            stTmp(1) = "and R013 is not null "
        End If
    End If
    
    '用Nvl(cu10,fa10) 因TS001663 無申請人
    If intChoose = 1 Then
        If bolTmpTB = False Then
            stTmp(0) = stTmp(0) & "and Nvl(cu10,fa10)='020' "
            stTmp(1) = stTmp(1) & "and Nvl(cu10,fa10)='020'"
        Else
            stTmp(0) = stTmp(0) & "and Nvl(R012,R014)='020' "
            stTmp(1) = stTmp(1) & "and Nvl(R012,R014)='020' "
        End If
    ElseIf intChoose = 2 Then
        If bolTmpTB = False Then
            stTmp(0) = stTmp(0) & "and Nvl(cu10,fa10)<>'020' "
            stTmp(1) = stTmp(1) & "and Nvl(cu10,fa10)<>'020' "
        Else
            stTmp(0) = stTmp(0) & "and Nvl(R012,R014)<>'020' "
            stTmp(1) = stTmp(1) & "and Nvl(R012,R014)<>'020' "
        End If
    End If
    
    If bolTmpTB = False Then
        stTmp(0) = "and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9,1) " & _
                          "and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1) " & stTmp(0)
        stTmp(1) = "and fa01(+)=substr(sp26,1,8) and fa02(+)=substr(sp26,9,1) " & _
                          "and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9,1) " & stTmp(1)
    End If
      
    stTmp2 = " And ((" & stF & ">=" & lngYear - 1 & "0101 and " & stF & "<=" & lngThisMonth - 100 & "31) " & _
    IIf(bolArrive = True, "Or (" & stF & ">=" & lngYear - 2 & "0101 and " & stF & "<=" & lngThisMonth - 200 & "31) ", "") & _
                                      "Or (" & stF & ">=" & lngYear & "0101 and " & stF & "<=" & lngThisMonth & "31)  )"
    '跑一月資料要加讀前一年12月資料
    If bolBefMon = True Then
        stTmp2 = " And " & stF & ">=" & lngLastMonth & "01 and " & stF & "<=" & lngLastMonth & "31 "
    End If
    'Add by Amy 2019/12/19 IsSurTran只抓結餘和轉撥
    If IsSurTran = True Then
        If bolTmpTB = False Then
            stTmp2 = stTmp2 & " And (Instr(ax212,'轉撥')>0 Or InStr(ax213||' ','結餘')>0 ) "
        Else
            stTmp2 = stTmp2 & " And (Instr(R007,'轉撥')>0 Or InStr(R008||' ','結餘')>0 ) "
        End If
    End If
    
    If stField = MsgText(601) Then
        stQ = "select ax205, sum(decode(floor(a0205/100)," & lngThisMonth & ",ax206)) Sd1" & _
                                     ", sum(decode(floor(a0205/100)," & lngThisMonth & ",ax207)) Sc1" & _
                                     ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2" & _
                                     ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2" & _
                                     ", sum(decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax206)) Sd3" & _
                                     ", sum(decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax207)) Sc3" & _
                                     ", sum(decode(floor(a0205/10000)," & lngYear & ",ax206)) Sd4" & _
                                     ", sum(decode(floor(a0205/10000)," & lngYear & ",ax207)) Sc4" & _
                                     ", sum(decode(floor(a0205/10000)," & lngYear - 1 & ",ax206)) Sd5" & _
                                     ", sum(decode(floor(a0205/10000)," & lngYear - 1 & ",ax207)) Sc5"
        If bolBefMon = True Then
            stQ = "select ax205, sum(decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2" & _
                                         ", sum(decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2"
        End If
    Else
        stQ = "Select " & stField
    End If
                                        
    'Modify by Amy 2020/05/14 拿掉公司別,改至外層判斷
    If bolTmpTB = False Then
        GetCCP = "Select a0205, ax205, ax206, ax207 from acc020, acc021, patent, fagent, customer" & _
            " Where ax201(+)=a0201 and ax202(+)=a0202 and ax205='411101' and ax209 is not null" & _
             stCon021 & stTmp2 & stTmp(0) & _
            " and pa01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and pa02(+)=substr(lpad(ax214,12,' '),4,6)" & _
            " and pa03(+)=substr(lpad(ax214,12,' '),10,1) and pa04(+)=substr(lpad(ax214,12,' '),11,2) and pa01 is not null" & _
            " Union All" & _
            " Select a0205, ax205, ax206, ax207 from acc020, acc021, servicepractice, fagent, customer" & _
            " Where  ax201(+)=a0201 and ax202(+)=a0202 and ax205='411101' and ax209 is not null" & _
             stCon021 & stTmp2 & stTmp(1) & _
            " and sp01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) and sp02(+)=substr(lpad(ax214,12,' '),4,6)" & _
            " and sp03(+)=substr(lpad(ax214,12,' '),10,1) and sp04(+)=substr(lpad(ax214,12,' '),11,2) and ( sp01 is not null or ax214 is null)"
    '暫存檔
    Else
        GetCCP = "select R003 as a0205,R004 as ax205, R005 as ax206,R006 as ax207 From Accrpt4202" & _
            " Where ID='" & strUserNum & "' And R004='411101' and R010 is not null" & _
            " and ( (ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 1) & ") and R009 is not null) " & _
                 "Or (ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 5) & ") Or R009 is null)  ) " & _
             Replace(Replace(stCon021, " ax205 ", " R004 "), " ax201 ", " ax201 ") & stTmp2 & stTmp(0)
    End If
    'end 2020/05/14
     
    GetCCP = stQ & " From (" & GetCCP & ") x Group By ax205"
    'end 2019/12/19
End Function

'Add by Amy 2019/12/19 專業達成點數分佈情況,寫傳票資料至暫存檔
'Memo by Amy 此處有改需確認Process3/專業點數明細表/專業單位實績點數分析表 /專業達成點數表-秘書 (frmacc44r0)是否也改
'Moidfy by Amy 2020/05/14 +bolVoucherOnly 只抓傳票資料
Private Sub Process2(Optional ByVal bolVoucherOnly = False)
    Dim RsQ As New ADODB.Recordset
    Dim i As Integer, intQ As Integer
    Dim stQ As String, strSql As String, stSQL As String, StrSQLa As String
    Dim stVTB As String, stVTB1 As String
    Dim stCFT As String, stCFP As String, stTB As String, stWhere As String
    Dim strArrive(2) As String
    Dim strInsF As String, strInsF1 As String 'Add by Amy 2020/05/14
    Dim strItemAcc As String, strAllAcc As String, strLAcc As String 'Add by Amy 2020/07/09 大項中包含之會計科目/所有已顯示之會計科目/法務會計科目
    Dim strF(1) As String 'Add by Amy 2021/08/23
    
On Error GoTo ErrHnd
        
    stSQL = "Delete From Accrpt4202 Where ID='" & strUserNum & "' "
    cnnConnection.Execute stSQL
    
    stSQL = "": stCon021 = "": stCon040 = ""
    'Modify by Amy 2020/05/14 公司別改下拉
'    If Text3 <> "" Then
'        If Text3 = "2" Then
'            stSQL = stSQL & " And a0201='J' "
'        Else
'            stSQL = stSQL & " And a0201='1' "
'        End If
'    End If
    strCmp = CboCmp
    If Trim(strCmp) <> "" Then
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
        If InStr(strCmp, "+") > 0 Then
            stSQL = stSQL & " And a0201 In ('" & Replace(strCmp, "+", "','") & "') "
            'Add by Amy 2020/06/17 +公司別
            stCon040 = stCon040 & " And a0403 In ('" & Replace(strCmp, "+", "','") & "') "
        Else
            stSQL = stSQL & " And a0201 = '" & strCmp & "' "
            'Add by Amy 2020/06/17 +公司別
            stCon040 = stCon040 & " And a0403 = '" & strCmp & "' "
        End If
    End If
    strCmpN = GetAccReportCmpN(strCmp, , True)
    If txtAccNo <> "" Then
        stSQL = stSQL & " And ax205 = '" & txtAccNo & "' "
        stCon021 = " and R004='" & txtAccNo & "'"
        stCon040 = " and a0405='" & txtAccNo & "'"
    End If
    strInsF = ",cu10"
    If bolVoucherOnly = True Then
        strInsF = ",Na01"
        strInsF1 = ",Decode(pa01,null,Decode(tm01,null,Decode(lc01,null,Nvl(sp09,''),lc15),tm10),pa09) as Na01 "
        'Modify by Amy 2020/10/08 避免會計新增相關會計編號導致資料未列到,原抓會計科目6碼,FCT/FCP 改抓智權人員,其他改抓4碼會計科目
        'stSQL = stSQL & " And ax205 In ('417201','417203','417102','417103','417104','417105','417109','417101','412101','412102','413101','413102') "
        'modify by sonia 2021/1/20 +F4104~F4107
        stSQL = stSQL & " And (SubStr(ax205,1,4) In ('4121','4131') Or ax209 In ('F4102','F4103','F4104','F4105','F4106','F4107'))"
        stSQL = stSQL & " And InStr(ax213||' ','結餘')=0 And InStr(ax212,'轉撥')=0 "
    End If
    'end 2020/05/14
    
    '傳票資料寫入暫存檔
    'Modify by Amy 2021/08/23 +ax204
    If bolVoucherOnly = False Then strF(0) = ",R017": strF(1) = ",ax204"
    
    'Modify by Amy 2020/10/08 國家別FCT/FCP改以智權角度抓資料,故加對沖代號-客(R015)及傳票項次
    stSQL = "Insert Into Accrpt4202 (ID,R001,R002,R003,R004,R005,R006,R007,R008,R009,R010,R011,R012,R013,R014,R015,R016" & strF(0) & ") " & _
                 "Select '" & strUserNum & "',ax201,ax202,a0205, ax205, ax206, ax207,ax212,ax213,ax214,ax209,CuNo" & strInsF & ",FaNo,fa10,ax208,ax203" & strF(1) & " " & _
                 "From Customer,Fagent,(" & _
                        "Select ax201,ax202,ax203,a0205, ax205, ax206, ax207,ax208,ax209,ax212,ax213,ax214" & strF(1) & " " & _
                                 ",Decode(pa01,null,Decode(tm01,null,Decode(lc01,null,Nvl(sp08,''),lc11),tm23),pa26) as CuNo " & _
                                 ",Decode(pa01,null,Decode(tm01,null,Decode(lc01,null,Nvl(sp26,''),lc22),tm44),pa75) as FaNo " & strInsF1 & _
                        "From (Select ax201,ax202,ax203,a0205, ax205, ax206, ax207,ax208,ax209,ax212,ax213,ax214" & strF(1) & " " & _
                                "From Acc020,Acc021 " & _
                                "Where a0201=ax201(+) And a0202=ax202(+) " & stSQL & _
                                " And (SubStr(ax205, 1, 1) = '4' Or (ax205='7121' And ax209 is not null)) And Not( ax205='4191' or ax205='4192' or ax205='4194') " & _
                                " And ( (a0205>=" & lngYear - 1 & "0101 and a0205<=" & lngThisMonth - 100 & "31) " & _
                                      "Or (a0205>=" & lngYear - 2 & "0101 and a0205<=" & lngThisMonth - 200 & "31) " & _
                                      "Or (a0205>=" & lngYear & "0101 and a0205<=" & lngThisMonth & "31)     ) " & _
                           "),Patent,TradeMark,LawCase,ServicePractice " & _
                        "Where pa01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) And pa02(+)=substr(lpad(ax214,12,' '),4,6)" & _
                           " And pa03(+)=substr(lpad(ax214,12,' '),10,1) And pa04(+)=substr(lpad(ax214,12,' '),11,2)" & _
                           " And tm01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) And tm02(+)=substr(lpad(ax214,12,' '),4,6)" & _
                           " And tm03(+)=substr(lpad(ax214,12,' '),10,1) And tm04(+)=substr(lpad(ax214,12,' '),11,2)" & _
                           " And lc01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) And lc02(+)=substr(lpad(ax214,12,' '),4,6)" & _
                           " And lc03(+)=substr(lpad(ax214,12,' '),10,1) And lc04(+)=substr(lpad(ax214,12,' '),11,2)" & _
                           " And sp01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) And sp02(+)=substr(lpad(ax214,12,' '),4,6)" & _
                           " And sp03(+)=substr(lpad(ax214,12,' '),10,1) And sp04(+)=substr(lpad(ax214,12,' '),11,2)" & _
                 ") Where fa01(+)=substr(FaNo,1,8) And fa02(+)=substr(FaNo,9,1) " & _
                      "And cu01(+)=substr(CuNo,1,8) And cu02(+)=substr(CuNo,9,1) "

    cnnConnection.Execute stSQL
    
    If bolVoucherOnly = True Then Exit Sub 'Add by Amy 2020/05/14

    '設定欄位
    strArrive(0) = "Sum(decode(floor(a0205/100)," & lngThisMonth & ",ax206)) Sd1, Sum(decode(floor(a0205/100)," & lngThisMonth & ",ax207)) Sc1" & _
                        ", Sum(Decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2, Sum(Decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2" & _
                        ", Sum(Decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax206)) Sd3, Sum(Decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax207)) Sc3" & _
                        ", Sum(Decode(floor(a0205/10000)," & lngYear & ",ax206)) Sd4, Sum(Decode(floor(a0205/10000)," & lngYear & ",ax207)) Sc4" & _
                        ", Sum(Decode(floor(a0205/10000)," & lngYear - 1 & ",ax206)) Sd5, Sum(Decode(floor(a0205/10000)," & lngYear - 1 & ",ax207)) Sc5" & _
                        ", Sum(Decode(floor(a0205/10000)," & lngYear - 2 & ",ax206)) Sd6,Sum(decode(floor(a0205/10000)," & lngYear - 2 & ",ax207)) Sc6" & _
                        ", Sum(Decode(floor(a0205/100)," & lngThisMonth - 200 & ",ax206)) Sd7,Sum(Decode(floor(a0205/100)," & lngThisMonth - 200 & ",ax207)) Sc7"
    
    strArrive(1) = "Sum(Decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1, Sum(Decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
                        ", Sum(Decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3, Sum(Decode(a0401," & lngYear & ",a0408)) net4" & _
                        ", Sum(Decode(a0401," & lngYear - 1 & ",a0408)) net5, Sum(Decode(a0401," & lngYear - 2 & ",a0408)) net6" & _
                        ", Sum(Decode(a0401*100+a0402," & lngThisMonth - 200 & ",a0408)) net7"
                        
    strArrive(2) = "net1-nvl(decode(a0103,'1',(Sd1-Sc1),(Sc1-Sd1)),0) C01, net2-nvl(decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)),0) C02" & _
                        ", net3-nvl(decode(a0103,'1',(Sd3-Sc3),(Sc3-Sd3)),0) C03, net4-nvl(decode(a0103,'1',(Sd4-Sc4),(Sc4-Sd4)),0) C04" & _
                        ", net5-nvl(decode(a0103,'1',(Sd5-Sc5),(Sc5-Sd5)),0) C05, net6-Nvl(Decode(a0103,'1',(Sd6-Sc6),(Sc6-Sd6)),0) C06" & _
                        ", net7-nvl(decode(a0103,'1',(Sd7-Sc7),(Sc7-Sd7)),0) C07"
   
 'Memo 沒有本所號的歸到最後一句 or ax214 is null
 '           RID有變動ExcelSave2也需調整
 
 '*** 商標收入 ***
    'Add by Amy 2020/07/10 抓取 商標收入 所有會計科目(417202 FCT-爭議 列於商標收入大項中)
    strItemAcc = GetAccList("'4101'") & ",'417202'"
    strAllAcc = strItemAcc
    'end 2020/07/10
    
    'MCT+MFCT
    stVTB = GetCCT(0, "'410101','410104'", False, stCon021, "ax205," & strArrive(0), True)
   
    '餘額
    stVTB1 = GetACC040("'410101','410104'", stCon040, "a0405," & strArrive(1))
 
    '主要顯示:CCT/CCT爭議 (CCT=餘額-MCT-MFCT(有代理人))
    strSql = "Select '" & strUserNum & "',DECODE(a0101,'410101','110','410104','140') RID, a0101, a0102 C00," & strArrive(2) & _
      " From acc010, (" & stVTB1 & ") w, (" & stVTB & ") y where a0101 IN ('410101','410104')" & _
      " and a0405(+)=a0101 and ax205(+)=a0101"
  
   'MCT有代理人且客戶國籍為大陸
    strSql = strSql & " Union All" & _
      " Select '" & strUserNum & "',DECODE(a0101,'410101','111','410104','141') RID , a0101, a0102||'-'||'MCT'," & _
      Replace(Replace(Replace(Replace(Replace(Replace(Replace(strArrive(2), "net1-", ""), "net2-", ""), "net3-", ""), "net4-", ""), "net5-", ""), "net6-", ""), "net7-", "") & _
      " From acc010, (" & GetCCT(1, "'410101','410104'", False, stCon021, "ax205," & strArrive(0), True) & " ) y where a0101 IN ('410101','410104') and ax205(+)=a0101"
    'MFCT 有代理人且客戶國籍為非大陸
    strSql = strSql & " Union All" & _
      " Select '" & strUserNum & "',DECODE(a0101,'410101','112','410104','142') RID , a0101, a0102||'-'||'MFCT'," & _
     Replace(Replace(Replace(Replace(Replace(Replace(Replace(strArrive(2), "net1-", ""), "net2-", ""), "net3-", ""), "net4-", ""), "net5-", ""), "net6-", ""), "net7-", "") & _
      " From acc010, (" & GetCCT(2, "'410101','410104'", False, stCon021, "ax205," & strArrive(0), True) & " ) y where a0101 IN ('410101','410104') and ax205(+)=a0101"
        
    Call ReplaceNo("'410101','410104'", strItemAcc) 'Add by Amy 2020/07/10

    'CMT/CCT條碼申請/CCT監視系統/CCT馬德里/CCT網址申請/FMT/CCT法務/FCT收入-爭議
    stVTB1 = GetACC040("'410102','410103','410109','410105','410106','410107','410108','417202','410110'", stCon040, "a0405," & strArrive(1))
      
   '主要顯示:CMT/CCT條碼申請/CCT監視系統/CCT馬德里/CCT網址申請/FMT/CCT法務/FCT收入-爭議
    strSql = strSql & " Union All" & _
      " Select '" & strUserNum & "',decode(a0101,'410102','12','410103','13','410109','131','410105','15','410106','16','410107','17','410108','18','417202','19T','410110','181')" & _
      " , a0101, a0102, net1, net2, net3, net4, net5" & IIf(bolArrive = True, ",net6,net7", "") & _
      " From acc010, (" & stVTB1 & ") w where a0101 in ('410102','410103','410109','410105','410106','410107','410108','417202','410110')" & _
      " and a0405(+)=a0101"
      
      Call ReplaceNo("'410102','410103','410109','410105','410106','410107','410108','417202','410110'", strItemAcc) 'Add by Amy 2020/07/10
      
      'Add by Amy 2020/07/10 過濾尚未歸入之會計科目 顯示 1x+會計編號後2碼
      If strItemAcc <> MsgText(601) Then
            stVTB1 = GetACC040(Mid(strItemAcc, 2), stCon040, "a0405," & strArrive(1))
            strSql = strSql & " Union All" & _
                        " Select '" & strUserNum & "','1x'||Substr(a0101,4,2) RID, a0101, a0102, net1, net2, net3, net4, net5" & IIf(bolArrive = True, ",net6,net7", "") & _
                        " From acc010, (" & stVTB1 & ") w where a0101 in (" & Mid(strItemAcc, 2) & ")" & _
                        " and a0405(+)=a0101"
      End If
      'end 2020/07/10
    
    '商標國內專業合計
    strSql = strSql & " Union All select '" & strUserNum & "','1z', null, '商標國內專業合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
    
    'FCT收入(含法務、服務業務案)
    'Add by Amy 2020/07/10 抓取 FCT收入 所有會計科目(剔除 417202 FCT-爭議 因列於商標收入大項中)
    strItemAcc = Replace(GetAccList("'4172'"), ",'417202'", "")
    strAllAcc = strAllAcc & strItemAcc
    'end 2020/07/10
    
    'Modify by Amy 2020/05/14 改不顯示國家
'    stVTB = "Select ax205, RID, " & strArrive(0) & _
'      " From (select R003 as a0205,R004 as ax205,Decode(substr(nvl(R014,R012),1,3),'101','21','011','22','012','23',Decode(substr(nvl(R014,R012),1,1),'2','24','25')) RID,R005 as ax206,R006 as ax207" & _
'                " From Accrpt4202" & _
'                " Where ID='" & strUserNum & "' And ((R003>=" & lngYear - 1 & "0101 and R003<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (R003>=" & lngYear - 2 & "0101 and R003<=" & lngThisMonth - 200 & "31) ", "") & _
'                    "or (R003>=" & lngYear & "0101 and R003<=" & lngThisMonth & "31))" & _
'                    " and R004='417201' " & Replace(Replace(stCon021, " ax205 ", " R004 "), " ax201 ", " R001 ") & _
'                    " and ( ((ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 2) & ") " & _
'                          " Or  ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 5) & ") ) and R009 is not null) " & _
'                       " Or  (ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 3) & ") Or R009 is null) " & _
'                            " )"
'    stVTB = stVTB & _
'      " Union All select 0,'417201','21',0,0 from dual" & _
'      " Union All select 0,'417201','22',0,0 from dual" & _
'      " Union All select 0,'417201','23',0,0 from dual" & _
'      " Union All select 0,'417201','24',0,0 from dual" & _
'      " Union All select 0,'417201','25',0,0 from dual" & _
'      ") x Group by ax205, RID"
    stVTB = "Select ax205, RID, " & strArrive(0) & _
      " From (select R003 as a0205,R004 as ax205,'21' RID,R005 as ax206,R006 as ax207" & _
                " From Accrpt4202" & _
                " Where ID='" & strUserNum & "' And ((R003>=" & lngYear - 1 & "0101 and R003<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (R003>=" & lngYear - 2 & "0101 and R003<=" & lngThisMonth - 200 & "31) ", "") & _
                    "or (R003>=" & lngYear & "0101 and R003<=" & lngThisMonth & "31))" & _
                    " and R004='417201' " & Replace(Replace(stCon021, " ax205 ", " R004 "), " ax201 ", " R001 ") & _
                    " and ( ((ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 2) & ") " & _
                          " Or  ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 5) & ") ) and R009 is not null) " & _
                       " Or  (ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 3) & ") Or R009 is null) " & _
                            " )"
    stVTB = stVTB & _
      " Union All select 0,'417201','21',0,0 from dual" & _
      ") x Group by ax205, RID"
    
    '主要顯示:FCT收入-國家
    'Modify by Amy 2020/05/14 改不顯示國家 原:, a0102||'-'||DECODE(RID,'21','美國','22','日本','23','韓國','24','歐洲','其他') C00
    strSql = strSql & " Union All" & _
      " select '" & strUserNum & "',RID, a0101, a0102 C00," & _
     Replace(Replace(Replace(Replace(Replace(Replace(Replace(strArrive(2), "net1-", ""), "net2-", ""), "net3-", ""), "net4-", ""), "net5-", ""), "net6-", ""), "net7-", "") & _
      " from acc010, (" & stVTB & ") y where a0101='417201' and ax205(+)=a0101"
      
    Call ReplaceNo("'417201'", strItemAcc) 'Add by Amy 2020/07/10
      
    'FCT收入-法務 417203
    'Modify by Amy 2021/08/23 +417202 FCT部門
    stVTB1 = GetACC040("'417203','417202'", stCon040, "a0405," & strArrive(1))

    strSql = strSql & " Union All" & _
      " select '" & strUserNum & "',Decode(a0101,'417203','26','417202','2xFCT') RID, a0101, a0102, net1, net2, net3, net4, net5,net6,net7" & _
      " from acc010, (" & stVTB1 & ") w where a0101 in ('417203','417202')" & _
      " and a0405(+)=a0101"
    'end 2021/08/23
      
    Call ReplaceNo("'417203'", strItemAcc) 'Add by Amy 2020/07/10
    
    'Add by Amy 2020/07/10 過濾尚未歸入之會計科目 顯示 2x+會計編號後2碼
    If strItemAcc <> MsgText(601) Then
        stVTB1 = GetACC040(Mid(strItemAcc, 2), stCon040, "a0405," & strArrive(1))
        strSql = strSql & " Union All" & _
                    " Select '" & strUserNum & "','2x'||Substr(a0101,4,2) RID, a0101, a0102, net1, net2, net3, net4, net5" & IIf(bolArrive = True, ",net6,net7", "") & _
                    " From acc010, (" & stVTB1 & ") w where a0101 in (" & Mid(strItemAcc, 2) & ")" & _
                    " and a0405(+)=a0101"
    End If
    'end 2020/07/10
      
    'FCT收入合計
    strSql = strSql & " Union All select '" & strUserNum & "','2z' , null, 'FCT收入合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
    
    'Add by Amy 2020/07/10 抓取 CFT收入 所有會計科目
    strItemAcc = GetAccList("'4121'")
    strAllAcc = strAllAcc & strItemAcc
    'end 2020/07/10
    
    'CFT收入412101及CFT收入-法務412102
    stCFT = GetACC040("'412101','412102'", stCon040, "a0405," & strArrive(1))
    
    '主要顯示:CFT收入/收入-法務
    strSql = strSql & " Union All" & _
      " select '" & strUserNum & "',decode(a0101,'412101','31','412102','32') RID, a0101, a0102, net1, net2, net3, net4,net5,net6,net7" & _
      " from acc010, (" & stCFT & ") w where a0101 in ('412101','412102') and a0405(+)=a0101 "
    
    Call ReplaceNo("'412101','412102'", strItemAcc) 'Add by Amy 2020/07/10
    
    'Add by Amy 2020/07/10 過濾尚未歸入之會計科目 顯示 3x+會計編號後2碼
    If strItemAcc <> MsgText(601) Then
        stVTB1 = GetACC040(Mid(strItemAcc, 2), stCon040, "a0405," & strArrive(1))
        strSql = strSql & " Union All" & _
                    " Select '" & strUserNum & "','3x'||Substr(a0101,4,2) RID, a0101, a0102, net1, net2, net3, net4, net5" & IIf(bolArrive = True, ",net6,net7", "") & _
                    " From acc010, (" & stVTB1 & ") w where a0101 in (" & Mid(strItemAcc, 2) & ")" & _
                    " and a0405(+)=a0101"
    End If
    'end 2020/07/10
    
    'CFT收入合計
    strSql = strSql & " Union All select '" & strUserNum & "','3z', null, 'CFT收入合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
    
    '商標達成總計
    strSql = strSql & " Union All select '" & strUserNum & "','3zz', null, '商標達成總計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
'*** End 商標收入 ***

'*** 專利收入 ***
    'Add by Amy 2020/07/10 抓取 專利收入 所有會計科目
    strItemAcc = GetAccList("'4111'")
    strAllAcc = strAllAcc & strItemAcc
    'end 2020/07/10

    'MCP+MFCP
    stVTB = GetCCP(0, False, stCon021, "ax205," & strArrive(0), True)
    
    '餘額
    stVTB1 = GetACC040("'411101'", stCon040, "a0405," & strArrive(1))
    
    '主要顯示:CCP/MCP (CCP=餘額-MCP-MFCP(有代理人))
    strSql = strSql & " Union All" & _
      " Select '" & strUserNum & "','410' RID, a0101, a0102 NAME," & strArrive(2) & _
      " from acc010, (" & stVTB1 & ") w, (" & stVTB & ") y where a0101='411101'" & _
      " and a0405(+)=a0101 and ax205(+)=a0405"
      
    'MCP有代理人且客戶國籍為大陸
    strSql = strSql & " Union All" & _
      " Select '" & strUserNum & "','411' RID, a0101, a0102||'-'||'MCP'," & _
      Replace(Replace(Replace(Replace(Replace(Replace(Replace(strArrive(2), "net1-", ""), "net2-", ""), "net3-", ""), "net4-", ""), "net5-", ""), "net6-", ""), "net7-", "") & _
      " From acc010, ( " & GetCCP(1, False, stCon021, "ax205," & strArrive(0), True) & " ) y where a0101='411101' and ax205(+)=a0101"
    
    'MFCP 有代理人且客戶國籍為非大陸
    strSql = strSql & " Union All" & _
      " Select '" & strUserNum & "','412' RID, a0101, a0102||'-'||'MFCP'," & _
      Replace(Replace(Replace(Replace(Replace(Replace(Replace(strArrive(2), "net1-", ""), "net2-", ""), "net3-", ""), "net4-", ""), "net5-", ""), "net6-", ""), "net7-", "") & _
      " From acc010, ( " & GetCCP(2, False, stCon021, "ax205," & strArrive(0), True) & " ) y where a0101='411101' and ax205(+)=a0101"
    
    Call ReplaceNo("'411101'", strItemAcc) 'Add by Amy 2020/07/10
    
    'CCP顧問/CMP/爭議/領證及年費/FMP/CCP-法務
    stVTB1 = GetACC040("'411102','411103','411104','411105','411106','411107'", stCon040, "a0405," & strArrive(1))
      
    '主要顯示:CCP顧問/CMP/爭議/領證及年費/FMP/CCP-法務
    strSql = strSql & " Union All " & _
      " Select '" & strUserNum & "',decode(a0101,'411102','42','411103','43','411104','44','411105','45','411106','46','411107','461') RID" & _
      " , a0101, a0102, net1, net2, net3, net4, net5,net6,net7" & _
      " From acc010, (" & stVTB1 & ") where a0101 in ('411102','411103','411104','411105','411106','411107')" & _
      " and a0405(+)=a0101"
      
    Call ReplaceNo("'411102','411103','411104','411105','411106','411107'", strItemAcc) 'Add by Amy 2020/07/10
    
    'Add by Amy 2020/07/10 過濾尚未歸入之會計科目 顯示 4x+會計編號後2碼
    If strItemAcc <> MsgText(601) Then
        stVTB1 = GetACC040(Mid(strItemAcc, 2), stCon040, "a0405," & strArrive(1))
        strSql = strSql & " Union All" & _
                    " Select '" & strUserNum & "','4x'||Substr(a0101,4,2) RID, a0101, a0102, net1, net2, net3, net4, net5" & IIf(bolArrive = True, ",net6,net7", "") & _
                    " From acc010, (" & stVTB1 & ") w where a0101 in (" & Mid(strItemAcc, 2) & ")" & _
                    " and a0405(+)=a0101"
    End If
    'end 2020/07/10
      
    '專利國內專業合計
    strSql = strSql & " Union All select '" & strUserNum & "','4z', null, '專利國內專業合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"

    'Add by Amy 2020/07/10 抓取 FCP收入 所有會計科目
    strItemAcc = GetAccList("'4171'")
    strAllAcc = strAllAcc & strItemAcc
    'end 2020/07/10
    
    'FCP收入(含法務、服務業務案)-'417104','417101','417105','417101','417109'顯示於 417101
    'Modify by Amy 2020/05/14 改不顯示國家
'    stVTB = "Select ax205, RID, " & strArrive(0) & _
'      " From (select R003 as a0205,Decode(R004,'417104','417101','417105','417101','417109','417101',R004) as ax205,decode(substr(nvl(R014,R012),1,3),'101','51','011','52','012','53',decode(substr(nvl(R014,R012),1,1),'2','54','55')) RID,R005 as ax206,R006 as ax207" & _
'                " From Accrpt4202" & _
'                " Where ID='" & strUserNum & "' And ((R003>=" & lngYear - 1 & "0101 and R003<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (R003>=" & lngYear - 2 & "0101 and R003<=" & lngThisMonth - 200 & "31) ", "") & _
'                    "or (R003>=" & lngYear & "0101 and R003<=" & lngThisMonth & "31)) " & Replace(Replace(stCon021, " ax205 ", " R004 "), " ax201 ", " R001 ") & _
'                    " and DECODE(R004,'417104','417101','417105','417101','417109','417101',R004)='417101' " & _
'                    " and ( ((ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 1) & ") " & _
'                          " Or  ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 5) & ")) and R009 is not  null) " & _
'                       " Or  (ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 3) & ") Or R009 is null) " & _
'                            " )"
'     stVTB = stVTB & _
'      " Union All select 0,'417101','51',0,0 from dual" & _
'      " Union All select 0,'417101','52',0,0 from dual" & _
'      " Union All select 0,'417101','53',0,0 from dual" & _
'      " Union All select 0,'417101','54',0,0 from dual" & _
'      " Union All select 0,'417101','55',0,0 from dual" & _
'      " ) x Group by ax205, RID "
       'Modify by Amy 2020/06/04 拆開顯示 原:Decode(R004,'417104','417101','417105','417101','417109','417101',R004) as ax205,'51' RID
       stVTB = "Select ax205, RID, " & strArrive(0) & _
      " From (select R003 as a0205,R004 as ax205,Decode(R004,'417101','51','417104','52','417105','53','54') RID,R005 as ax206,R006 as ax207" & _
                " From Accrpt4202" & _
                " Where ID='" & strUserNum & "' And ((R003>=" & lngYear - 1 & "0101 and R003<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (R003>=" & lngYear - 2 & "0101 and R003<=" & lngThisMonth - 200 & "31) ", "") & _
                    "or (R003>=" & lngYear & "0101 and R003<=" & lngThisMonth & "31)) " & Replace(Replace(stCon021, " ax205 ", " R004 "), " ax201 ", " R001 ") & _
                    " and R004 In('417101','417104','417105','417109') " & _
                    " and ( ((ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 1) & ") " & _
                          " Or  ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 5) & ")) and R009 is not  null) " & _
                       " Or  (ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 3) & ") Or R009 is null) " & _
                            " )"
     stVTB = stVTB & _
      " Union All select 0,'417101','51',0,0 from dual" & _
      " Union All select 0,'417104','52',0,0 from dual" & _
      " Union All select 0,'417105','53',0,0 from dual" & _
      " Union All select 0,'417109','54',0,0 from dual" & _
      " ) x Group by ax205, RID "
      
    '主要顯示:FCP收入-國家
    'Modify by Amy 2020/05/14 改不顯示國家 原:, a0102||'-'||DECODE(RID,'51','美國','52','日本','53','韓國','54','歐洲','55','其他') C00
    strSql = strSql & " Union All" & _
      " Select '" & strUserNum & "',Decode(a0101,'417101','51','417104','52','417105','53','417109','54'), a0101, a0102 C00," & _
      Replace(Replace(Replace(Replace(Replace(Replace(Replace(strArrive(2), "net1-", ""), "net2-", ""), "net3-", ""), "net4-", ""), "net5-", ""), "net6-", ""), "net7-", "") & _
      " From acc010, (" & stVTB & ") y where a0101 In ('417101','417104','417105','417109') and ax205(+)=a0101"
    'end 2020/06/04
    
    Call ReplaceNo("'417101','417104','417105','417109'", strItemAcc) 'Add by Amy 2020/07/10
      
    'FCP收入-FMP/法務
    stVTB1 = GetACC040("'417102','417103'", stCon040, "a0405," & strArrive(1))
      
    '主要顯示:FCP收入-FMP/法務
    strSql = strSql & " Union All" & _
      " select '" & strUserNum & "',decode(a0101,'417102','56','417103','57') RID" & _
      " , a0101, a0102, net1, net2, net3, net4, net5,net6,net7" & _
      " from acc010, (" & stVTB1 & ") w where a0101 in ('417102','417103')" & _
      " and a0405(+)=a0101"
      
    Call ReplaceNo("'417102','417103'", strItemAcc) 'Add by Amy 2020/07/10
    
    'Add by Amy 2020/07/10 過濾尚未歸入之會計科目 顯示 5x+會計編號後2碼
    If strItemAcc <> MsgText(601) Then
        stVTB1 = GetACC040(Mid(strItemAcc, 2), stCon040, "a0405," & strArrive(1))
        strSql = strSql & " Union All" & _
                    " Select '" & strUserNum & "','5x'||Substr(a0101,4,2) RID, a0101, a0102, net1, net2, net3, net4, net5" & IIf(bolArrive = True, ",net6,net7", "") & _
                    " From acc010, (" & stVTB1 & ") w where a0101 in (" & Mid(strItemAcc, 2) & ")" & _
                    " and a0405(+)=a0101"
    End If
    'end 2020/07/10
      
    'FCP收入合計
    strSql = strSql & " Union All select '" & strUserNum & "','5z', null, 'FCP收入合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"

    'Add by Amy 2020/07/10 抓取 CFP收入 所有會計科目
    strItemAcc = GetAccList("'4131'")
    strAllAcc = strAllAcc & strItemAcc
    'end 2020/07/10
    
    'CFP收入/法務
    stCFP = GetACC040("'413101','413102'", stCon040, "a0405," & strArrive(1))
    
    '主要顯示:CFP收入/法務
    'Modify by Amy 2021/01/21 原:RID=6z1
    strSql = strSql & " Union All" & _
      " select '" & strUserNum & "','6' RID, a0101, a0102, net1, net2, net3, net4,net5,net6,net7" & _
      " from acc010, (" & stCFP & ") w where a0405 in ('413101','413102') and a0405(+)=a0101 "
    'CFP收入合計 'Add by Amy 2021/01/21
    strSql = strSql & " Union All select '" & strUserNum & "','61z', null, 'CFP收入合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
    
    Call ReplaceNo("'413101','413102'", strItemAcc) 'Add by Amy 2020/07/10
    
    'Add by Amy 2020/07/10 過濾尚未歸入之會計科目 顯示顧 6x+會計編號後2碼
    If strItemAcc <> MsgText(601) Then
        stVTB1 = GetACC040(Mid(strItemAcc, 2), stCon040, "a0405," & strArrive(1))
        strSql = strSql & " Union All" & _
                    " Select '" & strUserNum & "','6x'||Substr(a0101,4,2) RID, a0101, a0102, net1, net2, net3, net4, net5" & IIf(bolArrive = True, ",net6,net7", "") & _
                    " From acc010, (" & stVTB1 & ") w where a0101 in (" & Mid(strItemAcc, 2) & ")" & _
                    " and a0405(+)=a0101"
    End If
    'end 2020/07/10
    
    '專利達成總計
    strSql = strSql & " Union All select '" & strUserNum & "','6zz', null, '專利達成總計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"

'*** End 專利收入 ***

'*** 其他收入 ***
    
    '法務收入(414101)/法務收入顧問(414102)/著作權收入(415101)/著作權收入-爭議(415102)/FCL收入(416101)/CFL收入(416102)
    'Modify by Amy 2020/07/09 +414109 法務收入-其他
    stVTB1 = GetACC040("'414101','414102','414109','415101','415102','416101','416102','420101'", stCon040, "a0405," & strArrive(1))
    
    '主要顯示:法務收入/法務收入顧問/著作權收入/著作權收入-爭議/FCL收入/CFL收入
    strSql = strSql & " Union All" & _
      " select '" & strUserNum & "',decode(a0101,'414101','73','414102','74','414109','75','415101','76','415102','77','416101','78','416102','79','420101','80')" & _
      " , a0101, a0102, net1, net2, net3, net4,net5,net6,net7" & _
      " from acc010, (" & stVTB1 & ") w where a0101 in ('414101','414102','414109','415101','415102','416101','416102','420101')" & _
      " and a0405(+)=a0101 "
    'end 2020/07/09
    
    strLAcc = "'414101','414102','414109','415101','415102','416101','416102','420101'"
    
    'Add by Amy 2020/07/10 4字頭其他未列示,顯示顧 7x+會計編號後2碼
    strItemAcc = GetAccList(strLAcc & strAllAcc, True)
    If strItemAcc <> MsgText(601) Then
        stVTB1 = GetACC040(Mid(strItemAcc, 2), stCon040, "a0405," & strArrive(1))
        strSql = strSql & " Union All" & _
                    " Select '" & strUserNum & "','7x'||Substr(a0101,4,2) RID, a0101, a0102, net1, net2, net3, net4, net5" & IIf(bolArrive = True, ",net6,net7", "") & _
                    " From acc010, (" & stVTB1 & ") w where a0101 in (" & Mid(strItemAcc, 2) & ")" & _
                    " and a0405(+)=a0101"
    End If
    'end 2020/07/10
      
    '其他收入(7121)
    'Modify by Amy 2020/05/14 公司別判斷改至stCon021 原:IIf(Text3 = "2", " and R001='J'", IIf(Text3 = "1", " and R001='1'", ""))
    'Modify by Amy 2020/07/09 因加法務收入-其他 RID已排至80,原:'7b' as RID
    stVTB = "Select ax205, RID, " & strArrive(0) & _
     " From (select R003 as a0205,R004 as ax205,'9b' as RID,R005 as ax206,R006 as ax207" & _
                " From Accrpt4202" & _
                " Where ID='" & strUserNum & "' And ((R003>=" & lngYear - 1 & "0101 and R003<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (R003>=" & lngYear - 2 & "0101 and R003<=" & lngThisMonth - 200 & "31) ", "") & _
                    "or (R003>=" & lngYear & "0101 and R003<=" & lngThisMonth & "31))" & _
                    " and R004='7121' and R010 is not null " & Replace(Replace(stCon021, " ax205 ", " R004 "), " ax201 ", " R001 ") & _
     " ) Group by ax205, RID"
     
    '主要顯示:其他收入
    strSql = strSql & " Union All" & _
      " select '" & strUserNum & "','9b', a0101, '其他收入'," & _
      Replace(Replace(Replace(Replace(Replace(Replace(Replace(strArrive(2), "net1-", ""), "net2-", ""), "net3-", ""), "net4-", ""), "net5-", ""), "net6-", ""), "net7-", "") & _
      " from acc010, (" & stVTB & ") x where a0101='7121' and ax205(+)=a0101"
    'end 2020/07/09
'*** End 其他收入 ***
     
'專業達成總計
'Modify by Amy 2020/07/09 因加法務收入-其他 RID已排至80,原:'7zt1'
strSql = strSql & " Union All select '" & strUserNum & "','9zt1', null, '專業達成總計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"

'更新暫存檔資料
Call SaveTempTB(strSql)
Call UpdTempTB

'讀取暫存檔資料
strSql = "Select r4203 as C00,Round(r4204,0) as C01,Round(r4206,0) as C02,Round(R4210,0) as C03,'' as C04,'' as C05" & _
            ",Round(r4207,0) as C06,Round(r4208,0) as C07,Round(R4209,0) as C08,'' as C09,'' as C10, r4201 as RID,r4202 as A0101 " & _
            "From Accrpt420 Where ID='" & strUserNum & "' And R4211 is null  Order by r4201"

intI = 1
Set rsNew = ClsLawReadRstMsg(intI, strSql)
If intI = 1 Then
    Call ExcelSave2(strCmpN)  'Modify by Amy 2020/05/14 +strCmpN
Else
    MsgBox "無符合資料！"
End If
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Function GetACC040(ByVal stA0405 As String, ByVal stCon040 As String, Optional ByVal stField As String = "") As String
    Dim stWhere As String 'Add by Amy 2021/08/23
    
    If stField = MsgText(601) Then
        GetACC040 = "Select a0405, sum(decode(a0401*100+a0402," & lngThisMonth & ",a0408)) net1" & _
                                                    ", sum(decode(a0401*100+a0402," & lngLastMonth & ",a0408)) net2" & _
                                                    ", sum(decode(a0401*100+a0402," & lngThisMonth - 100 & ",a0408)) net3" & _
                                                    ", sum(decode(a0401," & lngYear & ",a0408)) net4" & _
                                                    ", sum(decode(a0401," & lngYear - 1 & ",a0408)) net5"
    Else
        GetACC040 = "Select " & stField
    End If
    'Modify by Amy 2021/08/23 +stWhere
    stWhere = "a0405 In(" & stA0405 & " ) And a0404='TOT' "
    '417202 FCT-爭議需判斷部門
    If InStr(stA0405, "417202") > 0 Then
        stWhere = "a0405 In(" & Replace(stA0405, ",'417202'", "") & ") And a0404='TOT' "
        '商標T
        If InStr(stA0405, "410102") > 0 Then
            stWhere = "((" & stWhere & ") Or (a0405='417202' And a0404='T')) "
        'FCT
        Else
            stWhere = "((" & stWhere & ") Or (a0405='417202' And a0404='FCT')) "
        End If
    End If
    'Modify by Amy 2020/05/14 拿掉公司別,改至外層判斷
    'Modify by Amy 2020/11/06 bug 原:a0402=
    GetACC040 = GetACC040 & _
                        " From Acc040 Where " & stWhere & stCon040 & _
                        " and (a0401=" & lngYear - 1 & IIf(bolArrive = True, " Or a0401= " & lngYear - 2, "") & " or a0401=" & lngYear & ") and a0402<=" & lngMonth & _
                        " Group by a0405"
    'end 2021/08/23
End Function

Private Sub SaveTempTB(ByVal strSql As String)
    'Modify by Amy 2019/05/14 與 frmacc44r0-ProduceData_Dept 共用 暫存檔,加R4211區分
    cnnConnection.Execute "Delete From Accrpt420 Where ID='" & strUserNum & "' And R4211 is null "
    strSql = "Insert Into Accrpt420 (ID,r4201,r4202,r4203,r4204,r4205,r4206,r4207,r4208" & IIf(bolArrive = True, ",R4209,R4210", "") & ") " & strSql
    cnnConnection.Execute strSql
    
    'add by sonia 2016/2/18
    cnnConnection.Execute "Delete From Accrpt4201 Where ID='" & strUserNum & "' "
End Sub

Private Sub UpdTempTB()
    Dim strSql As String
    
    'Add by Amy 2020/05/14 +if
    If Combo2 <> "國家別點數分析表" Then
        'Add by Amy 2019/02/18 拿掉「CCT-/CCP-」文字-婧瑄  ex:商標收入-CCT-MCT
        'Modify by Amy 2019/05/02 會計名稱全型英文字部分改半型
        strSql = "Update Accrpt420 set r4203=Replace(Replace(R4203,'CCT-MCT','MCT'),'CCT-MFCT','MFCT') " & _
                     "Where ID='" & strUserNum & "' And (Instr(R4203,'CCT-MCT')>0  or Instr(R4203,'CCT-MFCT')>0 ) And R4211 is null "
        cnnConnection.Execute strSql
        '拿掉「CCT爭議-/CCP爭議-」文字-婧瑄  ex:商標收入-CCT爭議-MCT
        strSql = "Update Accrpt420 set r4203=Replace(Replace(R4203,'CCT爭議-MCT','MCT爭議'),'CCT爭議-MFCT','MFCT爭議') " & _
                     "Where ID='" & strUserNum & "' And (Instr(R4203,'CCT爭議-MCT')>0  or Instr(R4203,'CCT爭議-MFCT')>0 ) And R4211 is null "
        cnnConnection.Execute strSql
        strSql = "Update Accrpt420 set r4203=Replace(Replace(R4203,'CCP-MCP','MCP'),'CCP-MFCP','MFCP') " & _
                     "Where ID='" & strUserNum & "' And (Instr(R4203,'CCP-MCP')>0  or Instr(R4203,'CCP-MFCP')>0 ) And R4211 is null "
        cnnConnection.Execute strSql
         strSql = "Update Accrpt420 set r4203=Replace(R4203,'CCP爭議-MCP','MCP爭議') " & _
                     "Where ID='" & strUserNum & "' And Instr(R4203,'CCP爭議-MCP')>0  And R4211 is null "
        cnnConnection.Execute strSql
        'end 2019/02/18
    End If
    'end 2020/05/14
    
    '更新資料值-避免報表與Excel 值不同合計可能誤差,故資料先四捨五入到小數2位-婧瑄
    'Modify by Amy 2019/05/14 與 frmacc44r0-ProduceData_Dept 共用 暫存檔,加R4211區分
    strSql = "Update Accrpt420 Set r4204=Round(r4204,2),r4205=Round(r4205,2),r4206=Round(r4206,2),r4207=Round(r4207,2),r4208=Round(r4208,2),r4209=Round(r4209,2),R4210=Round(R4210,2) " & _
                 "Where ID='" & strUserNum & "' And R4211 is null "
    cnnConnection.Execute strSql
    'end 2019/05/14
   
End Sub
'end 2019/12/19

'Add by Amy 2020/05/14
'Memo 有修改看Progress2 /專業點數明細表/專業單位實績點數分析表 /專業達成點數表-秘書 是否要改
Private Sub Process3()
    
    Dim RsQ As New ADODB.Recordset
    Dim i As Integer, intQ As Integer
    Dim strSql As String, stVTB As String
    Dim strArrive(2) As String
    
On Error GoTo ErrHnd
    strArrive(0) = "Sum(decode(floor(a0205/100)," & lngThisMonth & ",ax206)) Sd1, Sum(decode(floor(a0205/100)," & lngThisMonth & ",ax207)) Sc1" & _
                        ", Sum(Decode(floor(a0205/100)," & lngLastMonth & ",ax206)) Sd2, Sum(Decode(floor(a0205/100)," & lngLastMonth & ",ax207)) Sc2" & _
                        ", Sum(Decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax206)) Sd3, Sum(Decode(floor(a0205/100)," & lngThisMonth - 100 & ",ax207)) Sc3" & _
                        ", Sum(Decode(floor(a0205/10000)," & lngYear & ",ax206)) Sd4, Sum(Decode(floor(a0205/10000)," & lngYear & ",ax207)) Sc4" & _
                        ", Sum(Decode(floor(a0205/10000)," & lngYear - 1 & ",ax206)) Sd5, Sum(Decode(floor(a0205/10000)," & lngYear - 1 & ",ax207)) Sc5" & _
                        ", Sum(Decode(floor(a0205/10000)," & lngYear - 2 & ",ax206)) Sd6,Sum(decode(floor(a0205/10000)," & lngYear - 2 & ",ax207)) Sc6" & _
                        ", Sum(Decode(floor(a0205/100)," & lngThisMonth - 200 & ",ax206)) Sd7,Sum(Decode(floor(a0205/100)," & lngThisMonth - 200 & ",ax207)) Sc7"
    
    strArrive(2) = "nvl(decode(a0103,'1',(Sd1-Sc1),(Sc1-Sd1)),0) C01, nvl(decode(a0103,'1',(Sd2-Sc2),(Sc2-Sd2)),0) C02" & _
                        ", nvl(decode(a0103,'1',(Sd3-Sc3),(Sc3-Sd3)),0) C03, nvl(decode(a0103,'1',(Sd4-Sc4),(Sc4-Sd4)),0) C04" & _
                        ", nvl(decode(a0103,'1',(Sd5-Sc5),(Sc5-Sd5)),0) C05, Nvl(Decode(a0103,'1',(Sd6-Sc6),(Sc6-Sd6)),0) C06" & _
                        ", nvl(decode(a0103,'1',(Sd7-Sc7),(Sc7-Sd7)),0) C07"
    
    'Modify by Amy 2020/10/08 FCT / FCP 改以智權人員為基礎抓資料,無本所案號之國籍更新修改
    '無本所案號之國籍更新
    Call UpdNoCaseNoAgNation
'    '*** FCT (以FC代理人國籍分)***
'    'Modify by Amy 2020/06/05 +417203 法務,若417201之FC代理人無國籍,一同併入417203法務及其他
'    stVTB = "Select ax205, RID, " & strArrive(0) & _
'      " From (Select R003 as a0205,Decode(R014,null,'417203',R004) as ax205" & _
'                ",Decode(R004,'417203','191',Decode(SubStr(R014,1,3),'011','101','101','102','231','103','205','104','012','105',Decode(SubStr(Na02,1,1),'A','107','B','107',Decode(Na02,'C20','106','C00','107','C10','108','C30','109','C40','110','191')))) RID" & _
'                ",R005 as ax206,R006 as ax207" & _
'                " From Accrpt4202,Nation" & _
'                " Where ID='" & strUserNum & "' And ((R003>=" & lngYear - 1 & "0101 and R003<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (R003>=" & lngYear - 2 & "0101 and R003<=" & lngThisMonth - 200 & "31) ", "") & _
'                    "or (R003>=" & lngYear & "0101 and R003<=" & lngThisMonth & "31))" & _
'                    " and R004 In('417201','417203') And R014=Na01(+) " & _
'                    " and ( ((ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 2) & ") " & _
'                          " Or  ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 5) & ") ) and R009 is not null) " & _
'                       " Or  (ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 3) & ") Or R009 is null) " & _
'                            " )"
'    stVTB = stVTB & _
'      " Union All Select 0,'417201','101',0,0 From dual" & _
'      " Union All Select 0,'417201','102',0,0 From dual" & _
'      " Union All Select 0,'417201','103',0,0 From dual" & _
'      " Union All Select 0,'417201','104',0,0 From dual" & _
'      " Union All Select 0,'417201','105',0,0 From dual" & _
'      " Union All Select 0,'417201','106',0,0 From dual" & _
'      " Union All Select 0,'417201','107',0,0 From dual" & _
'      " Union All Select 0,'417201','108',0,0 From dual" & _
'      " Union All Select 0,'417201','109',0,0 From dual" & _
'      " Union All Select 0,'417201','110',0,0 From dual" & _
'      " Union All Select 0,'417203','191',0,0 From dual" & _
'      ") x Group by ax205, RID"
'
'    strSql = strSql & " Select '" & strUserNum & "','100', 'FCT', '', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
'    strSql = strSql & " Union " & _
'      " Select '" & strUserNum & "',RID,a0101,Decode(RID,'101','日本','102','美國','103','德國','104','瑞士','105','韓國','106','歐洲(不含德瑞)','107','亞洲(不含日韓)','108','美洲(不含美國)','109','非洲','110','大洋洲','法務及其他') C00," & strArrive(2) & _
'      " From acc010, (" & stVTB & ") y Where a0101 In('417201','417203') and ax205(+)=a0101"
'    'end 2020/06/05
'
'    'Modify by Amy 2020/06/16 合計需含法務及其他 原:RID='18z'
'    strSql = strSql & " Union All select '" & strUserNum & "','1z', null, 'FCT合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
'    '*** End FCT合計 ***
'
'    '*** FCP (以FC代理人國籍分)***
'    'Modify by Amy 2020/06/05 +417102 FMP/417103 法務,若417101之FC代理人無國籍、法務,一同併入417102 FMP及其他
'    'Modiy by Amy 2020/07/09 417102 FMP 獨立列一行,因10901-06 列於法務及其他值過大-婧瑄 原:Decode(R014,null,'417102',Decode(R004,'417104','417101','417105','417101','417109','417101','417103','417102',R004)) as ax205
'    stVTB = "Select ax205, RID, " & strArrive(0) & _
'      " From (Select R003 as a0205,Decode(R014,null,'417103',Decode(R004,'417104','417101','417105','417101','417109','417101',R004)) as ax205" & _
'                ",Decode(R004,'417103','291','417102','211',Decode(SubStr(R014,1,3),'011','201','101','202','231','203','205','204','012','205',Decode(SubStr(Na02,1,1),'A','207','B','207',Decode(Na02,'C20','206','C00','207','C10','208','C30','209','C40','210','291')))) RID" & _
'                ",R005 as ax206,R006 as ax207" & _
'                " From Accrpt4202,Nation" & _
'                " Where ID='" & strUserNum & "' And ((R003>=" & lngYear - 1 & "0101 and R003<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (R003>=" & lngYear - 2 & "0101 and R003<=" & lngThisMonth - 200 & "31) ", "") & _
'                    "or (R003>=" & lngYear & "0101 and R003<=" & lngThisMonth & "31)) " & Replace(Replace(stCon021, " ax205 ", " R004 "), " ax201 ", " R001 ") & _
'                    " and Decode(R004,'417104','417101','417105','417101','417109','417101',R004) In('417101','417102','417103') And R014=Na01(+) " & _
'                    " and ( ((ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 1) & ") " & _
'                          " Or  ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 5) & ")) and R009 is not  null) " & _
'                       " Or  (ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 3) & ") Or R009 is null) " & _
'                            " )"
'     stVTB = stVTB & _
'      " Union All Select 0,'417101','201',0,0 From dual" & _
'      " Union All Select 0,'417101','202',0,0 From dual" & _
'      " Union All Select 0,'417101','203',0,0 From dual" & _
'      " Union All Select 0,'417101','204',0,0 From dual" & _
'      " Union All Select 0,'417101','205',0,0 From dual" & _
'      " Union All Select 0,'417101','206',0,0 From dual" & _
'      " Union All Select 0,'417101','207',0,0 From dual" & _
'      " Union All Select 0,'417101','208',0,0 From dual" & _
'      " Union All Select 0,'417101','209',0,0 From dual" & _
'      " Union All Select 0,'417101','210',0,0 From dual" & _
'      " Union All Select 0,'417102','211',0,0 From dual" & _
'      " Union All Select 0,'417103','291',0,0 From dual" & _
'      " ) x Group by ax205, RID "
'
'    strSql = strSql & " Union Select '" & strUserNum & "','200', 'FCP & FMP', '', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
'    strSql = strSql & " Union " & _
'      " Select '" & strUserNum & "',RID,a0101,Decode(RID,'201','日本','202','美國','203','德國','204','瑞士','205','韓國','206','歐洲(不含德瑞)','207','亞洲(不含日韓)','208','美洲(不含美國)','209','非洲','210','大洋洲','211','FMP','法務及其他') C00," & strArrive(2) & _
'      " From acc010, (" & stVTB & ") y Where a0101 In('417101','417102','417103') and ax205(+)=a0101"
'    'end 2020/07/09
'    'end 2020/06/05
'
'    'Modify by Amy 2020/06/16 合計需含法務及其他 原:RID='28z'
'    strSql = strSql & " Union All select '" & strUserNum & "','2z', null, 'FCP合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
'    '*** End FCP合計 ***
    strSql = GetSalesBase(strArrive(0), strArrive(2))
    'end 2020/10/08
  
    '*** CFT (以申請國家分)***
    'Modify by Amy 2020/06/05 +412102 法務,若417101之無案件申請國家,一同併入412102 法務及其他
    'Modify by Amy 2020/10/08 原:R004 In('412101','412102') /a0101 In('412101','412102')  and ax205(+)=a0101避免有會計科目未顯示,故抓會計科目4碼
    stVTB = "Select ax205, RID, " & strArrive(0) & _
      " From (Select a0205,Decode(RID,'391','412102','412101') as ax205,RID,ax206,ax207 " & _
                "From (Select R003 as a0205,Decode(R012,null,'412102',R004) as ax205," & _
                    "Decode(R004,'412102','391',Decode(SubStr(R012,1,3),'101','301','011','302','239','303','018','304','042','305',Decode(SubStr(Na02,1,1),'A','307','B','307',Decode(Na02,'C20','306','C00','307','C10','308','C30','309','C40','310','391')))) RID" & _
                    ",R005 as ax206,R006 as ax207" & _
                    " From Accrpt4202,Nation" & _
                    " Where ID='" & strUserNum & "' And ((R003>=" & lngYear - 1 & "0101 and R003<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (R003>=" & lngYear - 2 & "0101 and R003<=" & lngThisMonth - 200 & "31) ", "") & _
                        "or (R003>=" & lngYear & "0101 and R003<=" & lngThisMonth & "31))" & _
                        " and SubStr(R004,1,4) In('4121') And R012=Na01(+) " & _
                        " and ( ((ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 2) & ") " & _
                              " Or  ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 5) & ") ) and R009 is not null) " & _
                           " Or  (ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 3) & ") Or R009 is null) " & _
                                " ) ) "
    stVTB = stVTB & _
      " Union All Select 0,'412101','301',0,0 From dual" & _
      " Union All Select 0,'412101','302',0,0 From dual" & _
      " Union All Select 0,'412101','303',0,0 From dual" & _
      " Union All Select 0,'412101','304',0,0 From dual" & _
      " Union All Select 0,'412101','305',0,0 From dual" & _
      " Union All Select 0,'412101','306',0,0 From dual" & _
      " Union All Select 0,'412101','307',0,0 From dual" & _
      " Union All Select 0,'412101','308',0,0 From dual" & _
      " Union All Select 0,'412101','309',0,0 From dual" & _
      " Union All Select 0,'412101','310',0,0 From dual" & _
      " Union All Select 0,'412102','391',0,0 From dual" & _
      ") x Group by ax205, RID"
      
    strSql = strSql & " Union Select '" & strUserNum & "','300', 'CFT', '', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
    strSql = strSql & " Union " & _
      " Select '" & strUserNum & "',RID,a0101,Decode(RID,'301','美國','302','日本','303','歐盟','304','馬來西亞','305','越南','306','歐洲(不含歐盟)','307','亞洲(不含日馬越)','308','美洲(不含美國)','309','非洲','310','大洋洲','法務及其他') C00," & strArrive(2) & _
      " From acc010, (" & stVTB & ") y Where SubStr(a0101,1,4) In('4121')  and ax205=a0101(+)"
    'end 2020/10/05
    'end 2020/06/05
    
    'Modify by Amy 2020/06/16 合計需含法務及其他 原:RID='38z'
    strSql = strSql & " Union All select '" & strUserNum & "','3z', null, 'CFT合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
    '*** End CFT 合計***
   
    '*** CFP (以申請國家分)***
    'Modify by Amy 2020/10/08 原:R004 In('413101','413102') /a0101 In('413101','413102') and ax205(+)=a0101 避免有會計科目未顯示,故抓會計科目4碼
     stVTB = "Select ax205, RID, " & strArrive(0) & _
      " From (Select a0205,Decode(RID,'491','413102','413101') as ax205,RID,ax206,ax207 " & _
                "From (Select R003 as a0205,Decode(R012,null,'413102',R004) as ax205" & _
                    ",Decode(R004,'413102','491',Decode(SubStr(R012,1,3),'101','401','011','402','221','403','231','404',Decode(SubStr(Na02,1,1),'A','406','B','406',Decode(Na02,'C20','405','C00','406','C10','407','C30','408','C40','409','491')))) RID" & _
                    ",R005 as ax206,R006 as ax207" & _
                    " From Accrpt4202,Nation" & _
                    " Where ID='" & strUserNum & "' And ((R003>=" & lngYear - 1 & "0101 and R003<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (R003>=" & lngYear - 2 & "0101 and R003<=" & lngThisMonth - 200 & "31) ", "") & _
                        "or (R003>=" & lngYear & "0101 and R003<=" & lngThisMonth & "31))" & _
                        " and SubStr(R004,1,4) In('4131') And R012=Na01(+) " & _
                        " and ( ((ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 1) & ") " & _
                              " Or  ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 5) & ") ) and R009 is not null) " & _
                           " Or  (ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 3) & ") Or R009 is null) " & _
                                " )  )"
    stVTB = stVTB & _
      " Union All Select 0,'413101','401',0,0 From dual" & _
      " Union All Select 0,'413101','402',0,0 From dual" & _
      " Union All Select 0,'413101','403',0,0 From dual" & _
      " Union All Select 0,'413101','404',0,0 From dual" & _
      " Union All Select 0,'413101','405',0,0 From dual" & _
      " Union All Select 0,'413101','406',0,0 From dual" & _
      " Union All Select 0,'413101','407',0,0 From dual" & _
      " Union All Select 0,'413101','408',0,0 From dual" & _
      " Union All Select 0,'413101','409',0,0 From dual" & _
      " Union All Select 0,'413102','491',0,0 From dual" & _
      ") x Group by ax205, RID"
      
    strSql = strSql & " Union Select '" & strUserNum & "','400', 'CFP', '', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
    strSql = strSql & " Union " & _
      " Select '" & strUserNum & "',RID,a0101,Decode(RID,'401','美國','402','日本','403','EPC','404','德國','405','歐洲(不含EPC/德)','406','亞洲(不含日本)','407','美洲(不含美國)','408','非洲','409','大洋洲','法務及其他') C00," & strArrive(2) & _
      " From acc010, (" & stVTB & ") y Where SubStr(a0101,1,4) In('4131') and ax205=a0101(+)"
    'end 2020/10/08
    'end 2020/06/05
    
    'Modify by Amy 2020/06/16 合計需含法務及其他 原:RID='48z'
    strSql = strSql & " Union All select '" & strUserNum & "','4z', null, 'CFP合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
    '*** End CFP 合計***
    
    '更新暫存檔資料
    Call SaveTempTB(strSql)
    Call UpdTempTB

    '讀取暫存檔資料
    'Modify by Amy 2020/07/09 FCP->FCP & FMP
    strSql = "Select Decode(r4201,'100','FCT','200','FCP & FMP','300','CFT','400','CFP',r4203) as C00,Round(r4204,0) as C01,Round(r4206,0) as C02,Round(R4210,0) as C03,'' as C04,'' as C05" & _
                ",Round(r4207,0) as C06,Round(r4208,0) as C07,Round(R4209,0) as C08,'' as C09,'' as C10,'' as C11,'' as C12,'' as C13,r4201 as RID,r4202 as A0101 " & _
                "From Accrpt420 Where ID='" & strUserNum & "' And R4211 is null  Order by r4201"
    
    intI = 1
    Set rsNew = ClsLawReadRstMsg(intI, strSql)
    If intI = 1 Then
        Call ExcelSave3(strCmpN)
    Else
        MsgBox "無符合資料！"
    End If
   
ErrHnd:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
    
End Sub

'Add by Amy 2020/10/08
'更新無本所案號有對沖代號(客)之資料
Private Sub UpdNoCaseNoAgNation()
    Dim stUpd As String
    
    'FCT/FCP 原以 本所案號抓代理人國籍,增加若無本所案號且有對沖代號(客)且為Y編號者更新「代理人國籍」欄
    'modify by sonia 2021/1/20 +F4104~F4107
    stUpd = "Update Accrpt4202 Set R014=(Select Nvl(fa10,cu10) From Fagent,Customer " & _
                                                                    "Where SubStr(R015,1,8)=fa01(+) And SubStr(R015,9,1)=fa02(+) And fa01 is not null " & _
                                                                    "And SubStr(R015,1,8)=cu01(+) And SubStr(R015,9,1)=cu02(+) And cu01 is not null) " & _
                "Where ID='" & strUserNum & "' And R010 in ('F4102','F4103','F4104','F4105','F4106','F4107') And R009 is null "
     cnnConnection.Execute stUpd
End Sub

'以業務角度抓資料
Private Function GetSalesBase(ByVal strArrive0 As String, ByVal strArrive2 As String) As String
    Dim strSql As String, stVTB As String
    
    '*** FCT (以FC代理人或對沖其他(客)國籍分)***
    '抓傳票智權人員F4103資料,除417203 法務/410109 FMT,其他以會計科目前4碼 4172抓本所案號之FC代理人,若無國籍改抓有「對沖代號(客)」之國籍,若都無國籍資料併入417203法務及其他
    'modify by sonia 2021/1/20 +F4106,F4107
    stVTB = "Select ax205, RID, " & strArrive0 & _
      " From (Select R003 as a0205,Decode(R014,null,'417203',R004) as ax205" & _
                ",Decode(R004,'417203','191','410109','111',Decode(SubStr(R014,1,3),'011','101','101','102','231','103','205','104','012','105',Decode(SubStr(Na02,1,1),'A','107','B','107',Decode(Na02,'C20','106','C00','107','C10','108','C30','109','C40','110','191')))) RID" & _
                ",R005 as ax206,R006 as ax207" & _
                " From Accrpt4202,Nation" & _
                " Where ID='" & strUserNum & "' And ((R003>=" & lngYear - 1 & "0101 and R003<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (R003>=" & lngYear - 2 & "0101 and R003<=" & lngThisMonth - 200 & "31) ", "") & _
                    "or (R003>=" & lngYear & "0101 and R003<=" & lngThisMonth & "31))" & _
                    " and R010 In('F4103','F4106','F4107') And R014=Na01(+) " & _
                    " and ( ((ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 2) & ") " & _
                          " Or  ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 5) & ") ) and R009 is not null) " & _
                          " Or  (ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 3) & ") Or R009 is null)  ) "
    stVTB = stVTB & _
                " Union All Select 0,'417201','101',0,0 From dual" & _
                " Union All Select 0,'417201','102',0,0 From dual" & _
                " Union All Select 0,'417201','103',0,0 From dual" & _
                " Union All Select 0,'417201','104',0,0 From dual" & _
                " Union All Select 0,'417201','105',0,0 From dual" & _
                " Union All Select 0,'417201','106',0,0 From dual" & _
                " Union All Select 0,'417201','107',0,0 From dual" & _
                " Union All Select 0,'417201','108',0,0 From dual" & _
                " Union All Select 0,'417201','109',0,0 From dual" & _
                " Union All Select 0,'417201','110',0,0 From dual" & _
                " Union All Select 0,'410109','111',0,0 From dual" & _
                " Union All Select 0,'417203','191',0,0 From dual" & _
                 ") x Group by ax205, RID"
    stVTB = " Select RID,a0101," & strArrive2 & " From acc010, (" & stVTB & ") y Where ax205=a0101(+) "

    strSql = strSql & " Select '" & strUserNum & "','100', 'FCT', '', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
    strSql = strSql & " Union " & _
            " Select '" & strUserNum & "',RID,Decode(RID,'191','417203','111','410109','417201') as ax205,Decode(RID,'101','日本','102','美國','103','德國','104','瑞士','105','韓國','106','歐洲(不含德瑞)','107','亞洲(不含日韓)','108','美洲(不含美國)','109','非洲','110','大洋洲','111','FMT','法務及其他') as C00" & _
                          ",Sum(C01) as C01 ,Sum(C02) as C02 ,Sum(C03) as C03, Sum(C04) as C04,Sum(C05) as C05 ,Sum(C06) as C06 ,Sum(C07) as C07 " & _
            " From (" & stVTB & ") Group by RID "
    '合計
    strSql = strSql & " Union All select '" & strUserNum & "','1z', null, 'FCT合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
    '*** End FCT合計 ***

    '*** FCP (以FC代理人或對沖其他(客)國籍分)***
    '抓傳票智權人員F4102資料,417103 法務/417102 FMP,其他以會計科目前4碼 4171抓本所案號之FC代理人,若無國籍改抓有「對沖代號(客)」之國籍,若都無國籍資料併入417103法務及其他
    'Modify by Amy 2021/01/15 411106 (專利-FMP)原顯示於每個國家中,改顯示於FMP
    'modify by sonia 2021/1/20 +F4104,F4105
    stVTB = "Select ax205, RID, " & strArrive0 & _
      " From (Select R003 as a0205,Decode(R014,null,'417103',R004) as ax205" & _
                ",Decode(R004,'417103','291','417102','211','411106','211',Decode(SubStr(R014,1,3),'011','201','101','202','231','203','205','204','012','205',Decode(SubStr(Na02,1,1),'A','207','B','207',Decode(Na02,'C20','206','C00','207','C10','208','C30','209','C40','210','291')))) RID" & _
                ",R005 as ax206,R006 as ax207" & _
                " From Accrpt4202,Nation" & _
                " Where ID='" & strUserNum & "' And ((R003>=" & lngYear - 1 & "0101 and R003<=" & lngThisMonth - 100 & "31) " & IIf(bolArrive = True, " Or (R003>=" & lngYear - 2 & "0101 and R003<=" & lngThisMonth - 200 & "31) ", "") & _
                    "or (R003>=" & lngYear & "0101 and R003<=" & lngThisMonth & "31)) " & Replace(Replace(stCon021, " ax205 ", " R004 "), " ax201 ", " R001 ") & _
                    " and R010 In('F4102','F4104','F4105') And R014=Na01(+) " & _
                    " and ( ((ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 1) & ") " & _
                          " Or  ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 5) & ")) and R009 is not  null) " & _
                       " Or  (ltrim(substr(lpad(R009,12,' '),1,3)) in (" & SQLGrpStr(GetAllSysKind(, "ALL"), 3) & ") Or R009 is null)   ) "
    stVTB = stVTB & _
                " Union All Select 0,'417101','201',0,0 From dual" & _
                " Union All Select 0,'417101','202',0,0 From dual" & _
                " Union All Select 0,'417101','203',0,0 From dual" & _
                " Union All Select 0,'417101','204',0,0 From dual" & _
                " Union All Select 0,'417101','205',0,0 From dual" & _
                " Union All Select 0,'417101','206',0,0 From dual" & _
                " Union All Select 0,'417101','207',0,0 From dual" & _
                " Union All Select 0,'417101','208',0,0 From dual" & _
                " Union All Select 0,'417101','209',0,0 From dual" & _
                " Union All Select 0,'417101','210',0,0 From dual" & _
                " Union All Select 0,'417102','211',0,0 From dual" & _
                " Union All Select 0,'417103','291',0,0 From dual" & _
          " ) x Group by ax205, RID "
    stVTB = " Select RID,a0101," & strArrive2 & " From acc010, (" & stVTB & ") y Where ax205=a0101(+) "

    strSql = strSql & " Union Select '" & strUserNum & "','200', 'FCP & FMP', '', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
   'Modify by Amy 2021/01/15 411106 (專利-FMP)原顯示於每個國家中,改顯示於FMP
    strSql = strSql & " Union " & _
      " Select '" & strUserNum & "',RID,Decode(RID,'291','417103','211','411106','211','417102','417101') as ax205,Decode(RID,'201','日本','202','美國','203','德國','204','瑞士','205','韓國','206','歐洲(不含德瑞)','207','亞洲(不含日韓)','208','美洲(不含美國)','209','非洲','210','大洋洲','211','FMP','法務及其他') C00" & _
                    ",Sum(C01) as C01 ,Sum(C02) as C02 ,Sum(C03) as C03, Sum(C04) as C04,Sum(C05) as C05 ,Sum(C06) as C06 ,Sum(C07) as C07 " & _
      " From (" & stVTB & ") Group by RID "
    '合計
    strSql = strSql & " Union All select '" & strUserNum & "','2z', null, 'FCP合計', 0, 0, 0, 0, 0" & IIf(bolArrive = True, ",0,0", "") & " from dual"
    '*** End FCP合計 ***
    GetSalesBase = strSql
    
End Function

'Add by Amy 2020/05/14 國家別點數分析表
Private Sub ExcelSave3(ByVal strCmpN As String)
    Dim strTotal(1 To 10) As String, AllTotal(1 To 10) As String
    Dim strTp(1) As String
    Dim strWkName As String
    Dim intXlsSheet As Integer, intQ As Integer, intA As Integer
    Dim bolFormaula As Boolean
    Dim strBP As String, strUpd As String, strA As String, strF(2) As String, strWhere As String
    Dim strCmp As String
    'Add by Amy 2020/06/05
    Dim strOldRID As String, jj As Integer

On Error GoTo ErrHand
    If Dir(strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
       If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
          MkDir strExcelPath
       End If
    Else
       Kill strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43)
    End If

    SetPrinter
    intXlsSheet = 1: intField = 65
    xlsSalesPoint.SheetsInNewWorkbook = 3
    xlsSalesPoint.Workbooks.add

    If strWkName = MsgText(601) Then strWkName = Left(xlsSalesPoint.Worksheets(1).Name, Len(xlsSalesPoint.Worksheets(1).Name) - 1)
    Set wksaccrpt424 = xlsSalesPoint.Worksheets(strWkName & intXlsSheet)
    wksaccrpt424.Activate
    IsOpenXls = True
    'xlsSalesPoint.Visible = True
    wksaccrpt424.PageSetup.PaperSize = 9 'A4
    wksaccrpt424.PageSetup.Orientation = xlLandscape '橫印
    wksaccrpt424.PageSetup.LeftMargin = 28.34
    wksaccrpt424.PageSetup.RightMargin = 28.34
    wksaccrpt424.PageSetup.TopMargin = 42.51
    wksaccrpt424.PageSetup.BottomMargin = 42.51
    wksaccrpt424.PageSetup.HeaderMargin = 28.34
    wksaccrpt424.PageSetup.FooterMargin = 28.34

    m_lngRow = 1
    Call ExcelHead2(strCmpN)

    intTitleR = m_lngRow - 1: StartRow = m_lngRow
   
    rsNew.MoveFirst
    Do While rsNew.EOF = False
        If Left(strOldRID, 1) <> Left("" & rsNew.Fields("RID"), 1) Then StartRow = m_lngRow
        'Modify by Amy 2020/06/05 加各國佔比,但資料於合計時再計算 原:UBound(strField)
        For ii = 0 To UBound(strField)
            '大項
            If Right(rsNew.Fields("RID"), 2) = "00" Then
                If ii = GetValue("") Then
                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = rsNew.Fields(ii)
                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Font.Bold = True
                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).HorizontalAlignment = xlCenter
                End If
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.ColorIndex = 40 '設置儲存格填充色(膚)
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Interior.tintandshade = 0.5 '設深淺
            'Add by Amy 2020/06/05
            ElseIf ii > GetValue("") And InStr(rsNew.Fields("RID"), "z") > 0 And InStr(strField(ii), "各國佔比") > 0 Then
                For jj = StartRow + 1 To m_lngRow - 1
                    strTp(0) = Chr(intField + GetValue(Replace(strField(ii), "各國佔比", ""))) & jj & "/$" & Chr(intField + GetValue(Replace(strField(ii), "各國佔比", ""))) & "$" & m_lngRow
                    wksaccrpt424.Range(Chr(intField + ii) & jj).Formula = "=" & strTp(0)
                    wksaccrpt424.Range(Chr(intField + ii) & jj).NumberFormatLocal = "0.00%;[紅色]-0.00%"
                    wksaccrpt424.Range(Chr(intField + ii) & jj).Font.Bold = True
                Next jj
            '年度 vs年度 比率
            ElseIf ii > GetValue("") And InStr(strField(ii), "vs") > 0 Then
                If InStr(strField(ii), "年") > 0 Then
                    strTp(0) = Chr(intField + GetValue(Mid(strField(ii), 1, InStr(strField(ii), "年vs") - 1) & "年1-" & lngMonth & "月")) & m_lngRow & "/" & _
                                    Chr(intField + GetValue(Mid(strField(ii), InStr(strField(ii), "年vs") + 4) & "年1-" & lngMonth & "月")) & m_lngRow & "-1"
                Else
                    strTp(0) = Chr(intField + GetValue(Mid(strField(ii), 1, InStr(strField(ii), "vs") - 2) & "." & lngMonth)) & m_lngRow & "/" & _
                                    Chr(intField + GetValue(Mid(strField(ii), InStr(strField(ii), "vs") + 3) & "." & lngMonth)) & m_lngRow & "-1"
                End If
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Formula = "=" & strTp(0)
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = "0.00%;[紅色]-0.00%"
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Font.Bold = True
            '合計
            ElseIf ii > GetValue("") And InStr(rsNew.Fields("RID"), "z") > 0 Then
                strTp(0) = "Sum(" & Chr(intField + ii) & StartRow & ":" & Chr(intField + ii) & m_lngRow - 1 & ")"
                strTotal(ii) = strTotal(ii) & Chr(intField + ii) & m_lngRow & ","
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Formula = "=" & strTp(0)
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = "#,##0"
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Font.Bold = True
            '內容
            Else
                strTp(0) = "" & rsNew.Fields(ii)
                strTp(1) = "#,##0"
                If ii = GetValue("") Then
                    strTp(1) = ""
                Else
                    strTp(0) = Val(strTp(0))
                End If
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = strTp(0)
                If ii = GetValue("" & strField(0)) And InStr(rsNew.Fields("RID"), "z") > 0 Then
                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Font.Bold = True
                Else
                    wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = strTp(1)
                End If
            End If
        Next ii
        m_lngRow = m_lngRow + 1
        'Add by Amy 2020/06/05 因增加法務/其他/FMP ,但加總只有國家別,改記錄RID
        strOldRID = "" & rsNew.Fields("RID")
        rsNew.MoveNext
    Loop
  
    wksaccrpt424.Range(Chr(intField) & m_lngRow).Value = "此報表 FCT/FCP 以「業務」角度分析，可能與「專業點數」報表有差異"  'Add by Amy 2020/10/08
    Call ExcelHead2(strCmpN, True)
    'Add by Amy 2020/10/08 產生無本所案號之傳票資料
    m_lngRow = 1
    Call Sheet2(strWkName, intXlsSheet)
    '設回國家點數分析表
    Set wksaccrpt424 = xlsSalesPoint.Worksheets(strWkName & "1")
    wksaccrpt424.Activate
    'end 2020/10/08
    
    '框線
    If Val(xlsSalesPoint.Version) < 12 Then
         xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
    Else
         xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
    End If
    xlsSalesPoint.Workbooks.Close
    xlsSalesPoint.Quit
    Set wksaccrpt424 = Nothing
    Set xlsSalesPoint = Nothing

    StatusClear
    bolSheet2 = False
    MsgBox "Excel檔產生完畢！（" & strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43) & "）"
    Exit Sub
    
ErrHand:
    MsgBox Err.Description, , MsgText(5)
    If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
   End If
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   Set wksaccrpt424 = Nothing
   Set xlsSalesPoint = Nothing
End Sub

'Add by Amy 2020/10/08 產生無本所案號之傳票資料
Private Sub Sheet2(strWkName As String, intXlsSheet As Integer)
    Dim strQ As String, strTmp As String, strFormat As String
    Dim intQ As Integer, i As Integer
    
    'FCT/FCP
    'modify by sonia 2021/1/20 +F4104~F4107
    strQ = "Select Decode(R010,'F4103',1,'F4106',1,'F4107',1,2) Sort,R001,R002,R004,R005,R006,R016 From Accrpt4202 a Where ID='" & strUserNum & "' And R009 is null And R014 is null And R010 in('F4103','F4102','F4107','F4106','F4105','F4104') And (R004<>'417203' or R004<>'417103') "
    'CFT/CFP
    strQ = strQ & "Union " & _
              "Select Decode(SubStr(R004,1,4),'4121',3,4) Sort,R001,R002,R004,R005,R006,R016 From Accrpt4202 a Where ID='" & strUserNum & "' And R009 is null And R012 is null And SubStr(R004,1,4) in('4121','4131') And (R004<>'412102' or R004<>'413102') "
    strQ = strQ & " Order by Sort,R001,R002 Desc,R016 "
    intQ = 1
    If rsNew.State <> adStateClosed Then rsNew.Close
    Set rsNew = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        '設定Excel
        intXlsSheet = intXlsSheet + 1
        Set wksaccrpt424 = xlsSalesPoint.Worksheets(strWkName & intXlsSheet)
        wksaccrpt424.Activate
        Call ExcelHead_Voucher
        
        rsNew.MoveFirst
        Do While rsNew.EOF = False
            For i = LBound(strField) To UBound(strField)
                strFormat = ""
                Select Case strField(i)
                    Case "公司別"
                        strTmp = "" & rsNew.Fields("R001")
                    Case "傳票號碼"
                        strTmp = "" & rsNew.Fields("R002")
                    Case "項次"
                        strTmp = "" & rsNew.Fields("R016")
                        strFormat = "@"
                    Case "會計科目"
                        strTmp = "" & rsNew.Fields("R004")
                    Case "借方"
                        strTmp = "" & rsNew.Fields("R005")
                        strFormat = "#,##0"
                    Case "貸方"
                        strTmp = "" & rsNew.Fields("R006")
                        strFormat = "#,##0"
                End Select
                wksaccrpt424.Range(Chr(intField + i) & m_lngRow).HorizontalAlignment = xlCenter
                If strFormat <> MsgText(601) Then
                    wksaccrpt424.Range(Chr(intField + i) & m_lngRow).NumberFormatLocal = strFormat
                    If strFormat = "#,##0" Then
                        wksaccrpt424.Range(Chr(intField + i) & m_lngRow).HorizontalAlignment = xlRight
                    End If
                End If
                wksaccrpt424.Range(Chr(intField + i) & m_lngRow).Value = strTmp
            Next i
            m_lngRow = m_lngRow + 1
            rsNew.MoveNext
        Loop
        wksaccrpt424.Name = "有疑慮之傳票"
        wksaccrpt424.PageSetup.PaperSize = 9 'A4
        wksaccrpt424.PageSetup.Orientation = wdOrientLandscape '直印
        wksaccrpt424.PageSetup.LeftMargin = xlsSalesPoint.InchesToPoints(0.5)
        wksaccrpt424.PageSetup.RightMargin = xlsSalesPoint.InchesToPoints(0.5)
        wksaccrpt424.PageSetup.TopMargin = xlsSalesPoint.InchesToPoints(0.3)
        wksaccrpt424.PageSetup.BottomMargin = xlsSalesPoint.InchesToPoints(0.3)
        wksaccrpt424.PageSetup.HeaderMargin = xlsSalesPoint.InchesToPoints(0.5)
        wksaccrpt424.PageSetup.FooterMargin = xlsSalesPoint.InchesToPoints(0.5)
        wksaccrpt424.PageSetup.Zoom = 100 '縮放比例
        
        wksaccrpt424.Range(Chr(intField) & "1:" & Chr(intField + UBound(strField)) & m_lngRow - 1).Select
        xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Font.Size = 11
        wksaccrpt424.PageSetup.PrintTitleRows = "$1:$1" '表頭保留列
        wksaccrpt424.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
    End If
End Sub

'傳票資料表格欄位
Private Sub ExcelHead_Voucher()
    ReDim strField(5)
    ReDim intWidth(5)
    
    strField = Array("公司別", "傳票號碼", "項次", "會計科目", "借方", "貸方")
    intWidth = Array(7, 13, 5, 10, 13, 13)
        
    For ii = LBound(strField) To UBound(strField)
        wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = strField(ii)
        wksaccrpt424.Columns(Chr(intField + ii)).ColumnWidth = intWidth(ii)
        wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).HorizontalAlignment = xlCenter
    Next ii
    m_lngRow = m_lngRow + 1
End Sub
'end 2020/10/08

'Add by Amy 2020/07/09
'傳入大項會計科目,取得其所有會計科目 or 反取 (4字頭中排除stAcc 之會計科目)
Private Function GetAccList(ByVal stAcc As String, Optional ByVal bolReverse As Boolean = False) As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    If bolReverse = True Then
        strQ = " And SubStr(R004,1,1)='4' And R004 not in (" & stAcc & ")"
    Else
        strQ = "And SubStr(R004,1,4) in (" & stAcc & ")"
    End If
    
    strQ = "Select Distinct R004 From Accrpt4202 Where ID='" & strUserNum & "' " & strQ
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        Do While RsQ.EOF = False
            GetAccList = GetAccList & ",'" & RsQ.Fields("R004") & "'"
            RsQ.MoveNext
        Loop
    End If
End Function

'傳入需Replace會計編號及回傳變數
Private Sub ReplaceNo(ByVal stAccNo As String, ByRef stItemAcc As String)
    Dim i As Integer
    Dim arrTp
    
    arrTp = Split(stAccNo, ",")
    For i = LBound(arrTp) To UBound(arrTp)
        stItemAcc = Replace(stItemAcc, "," & arrTp(i), "")
    Next i
End Sub
'end 2020/07/09

'Add by Amy 2021/08/23 參考 frm44r0-秘書 增加當月、去年、前年及三年累計數字
'＊＊＊　注意！！！　此處有修改需確認 frmacc44r0-秘書 專業達成點數表 是否要改　＊＊＊
'專業達成點數(比較三年)-資料
Private Sub ProPointData()
    Dim rsA As New ADODB.Recordset
    Dim intA As Integer, i As Integer, bolPeriod As Boolean
    Dim stField_Fix As String, stField As String, stWhere As String, stGroup As String, st040Sql As String, st040Where As String, stZeroAccNo As String
    Dim stA As String, stA2 As String, stDateS As String, stDateE As String
    Dim stA3 As String, stField3 As String, stField3_Fix As String
    Dim bolHasLaw As Boolean 'Add by Amy 2020/02/17 11101月起法務科目全改為 490102
    
    '刪除暫存檔
    cnnConnection.Execute "Delete From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' "
    '*** 依傳票會計科目新增餘額值 ***
    stField_Fix = "'" & strUserNum & "','" & Me.Name & "' "
    stField = ",Decode(R4202||SubStr(R4201,3),'417202T','417202T','417202FCT','417202FCT',R4202) "
    stWhere = " And ID='" & strUserNum & "' And R4211 is null And R4202 is not null "
    stGroup = " Group by Decode(R4202||SubStr(R4201,3),'417202T','417202T','417202FCT','417202FCT',R4202)  "
    
    stField3_Fix = "Decode(R004||R017,'417202T','417202T','417202FCT','417202FCT',R004)"
    stField3 = ",SubStr(R003+19110000,1,6)-191100"
    stA3 = "Select " & stField3_Fix & " as R004" & stField3 & " as IYM,Sum(R006-R005) as IAmt From Accrpt4202 Where ID='" & strUserNum & "' And R001='L' "
    stA2 = "Insert Into AccRpt44r0 (ID,FormN,R002,R004,R010) "
    
    '當年當月
    stA = "Select " & stField_Fix & ",AccNo,Nvl(Amt,0)-Nvl(IAmt,0),YM From " & _
          "(Select " & stField_Fix & stField & " as AccNo,Sum(Nvl(R4204,0)) as Amt,'" & lngYear & Format(lngMonth, "00") & "' as YM From Accrpt420 Where 1=1 " & stWhere & stGroup & ")" & _
          ",(" & stA3 & " And " & Mid(stField3, 2) & "='" & lngYear & Format(lngMonth, "00") & "' Group by " & stField3_Fix & Replace(stField3, " as IYM", "") & ") " & _
          "Where AccNo=R004(+) and YM=IYM(+) "
    cnnConnection.Execute stA2 & stA
    '去年當月
    stA = "Select " & stField_Fix & ",AccNo,Nvl(Amt,0)-Nvl(IAmt,0),YM From " & _
          "(Select " & stField_Fix & stField & " as AccNo,Sum(Nvl(R4206,0)) as Amt,'" & lngYear - 1 & Format(lngMonth, "00") & "' as YM From Accrpt420 Where 1=1 " & stWhere & stGroup & ")" & _
          ",(" & stA3 & " And " & Mid(stField3, 2) & "='" & lngYear - 1 & Format(lngMonth, "00") & "' Group by " & stField3_Fix & Replace(stField3, " as IYM", "") & ") " & _
          "Where AccNo=R004(+) and YM=IYM(+) "
    cnnConnection.Execute stA2 & stA
    '前年當月
    stA = "Select " & stField_Fix & ",AccNo,Nvl(Amt,0)-Nvl(IAmt,0),YM From " & _
          "(Select " & stField_Fix & stField & " as AccNo,Sum(Nvl(R4210,0)) as Amt,'" & lngYear - 2 & Format(lngMonth, "00") & "' as YM From Accrpt420 Where 1=1 " & stWhere & stGroup & ")" & _
          ",(" & stA3 & " And " & Mid(stField3, 2) & "='" & lngYear - 2 & Format(lngMonth, "00") & "' Group by " & stField3_Fix & Replace(stField3, " as IYM", "") & ") " & _
          "Where AccNo=R004(+) and YM=IYM(+) "
    cnnConnection.Execute stA2 & stA
    '當年1月~畫面當月
    stField3 = ",SubStr(R003+19110000,1,4)-1911"
    stA3 = "Select " & stField3_Fix & " as R004" & stField3 & "||'13' as IYM,Sum(R006-R005) as IAmt From Accrpt4202 Where ID='" & strUserNum & "' And R001='L' "

    stA = "Select " & stField_Fix & ",AccNo,Nvl(Amt,0)-Nvl(IAmt,0),YM From " & _
          "(Select " & stField_Fix & stField & " as AccNo,Sum(Nvl(R4207,0)) as Amt,'" & lngYear & "13' as YM From Accrpt420 Where 1=1 " & stWhere & stGroup & ")" & _
          ",(" & stA3 & " And " & Mid(stField3, 2) & "='" & lngYear & "' Group by " & stField3_Fix & Replace(stField3, " as IYM", "") & ") " & _
          "Where AccNo=R004(+) and YM=IYM(+) "
    cnnConnection.Execute stA2 & stA
    '去年1月~去年畫面當月
    stA = "Select " & stField_Fix & ",AccNo,Nvl(Amt,0)-Nvl(IAmt,0),YM From " & _
          "(Select " & stField_Fix & stField & " as AccNo,Sum(Nvl(R4208,0)) as Amt,'" & lngYear - 1 & "13' as YM From Accrpt420 Where 1=1 " & stWhere & stGroup & ")" & _
          ",(" & stA3 & " And " & Mid(stField3, 2) & "='" & lngYear - 1 & "' Group by " & stField3_Fix & Replace(stField3, " as IYM", "") & ") " & _
          "Where AccNo=R004(+) and YM=IYM(+) "
    cnnConnection.Execute stA2 & stA
    '前年1月~前年畫面當月
    stA = "Select " & stField_Fix & ",AccNo,Nvl(Amt,0)-Nvl(IAmt,0),YM From " & _
          "(Select " & stField_Fix & stField & " as AccNo,Sum(Nvl(R4209,0)) as Amt,'" & lngYear - 2 & "13' as YM From Accrpt420 Where 1=1 " & stWhere & stGroup & ")" & _
          ",(" & stA3 & " And " & Mid(stField3, 2) & "='" & lngYear - 2 & "' Group by " & stField3_Fix & Replace(stField3, " as IYM", "") & ") " & _
          "Where AccNo=R004(+) and YM=IYM(+) "
    cnnConnection.Execute stA2 & stA
    '*** End 依傳票會計科目新增餘額值 ***
    
    '抓取結餘或轉撥資料(含實績轉撥)
    Call ProPoint(Me.Name, lngYear, lngMonth)
    'Add by Amy 2022/02/17 11101月起法務科目全改為 490102,因比較三年故判斷有法務資料才顯示
    bolHasLaw = ChkHasData("And (R002 In('411107','413102','417103','410110','412102','417203') Or SubStr(R002,1,4) In('4141','4161','4181')) ")
  
    '新增 會計科目
    stA2 = "Select Distinct R010 From Accrpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' "
    'Modify by Amy 2022/02/17 11101月起法務科目全改為 490102,因比較三年故判斷有法務資料才顯示
    stA = "Select '4101' as AccNo From Dual " & _
    "Union Select '4111' as AccNo From Dual " & _
    "Union Select '4121' as AccNo From Dual "
    If bolHasLaw = True Then stA = stA & "Union Select '4141' as AccNo From Dual "
    'end 2022/02/17
    stA = "Insert Into AccRpt44r0 (ID,FormN,R002,R010) " & _
             "Select '" & strUserNum & "','" & Me.Name & "',AccNo,R010 From (" & stA & "),(" & stA2 & ") "
    cnnConnection.Execute stA
    
    '更新F41XX
    stDateS = lngYear & Format(lngMonth, "00")
    stDateE = stDateS
    bolPeriod = True
    For i = 1 To 3
        Select Case i
            '畫面當年1月~畫面當月
            Case 1
                stDateS = lngYear & "01"
                stDateE = lngYear & Format(lngMonth, "00")
            Case 2
                stDateS = lngYear - 1 & "01"
                stDateE = lngYear - 1 & Format(lngMonth, "00")
            Case 3
                stDateS = lngYear - 2 & "01"
                stDateE = lngYear - 2 & Format(lngMonth, "00")
        End Select
        Call InsF41XX(Me.Name, stDateS, stDateE, bolPeriod)
    Next i
    
    '*** 更新最後抓取的資料值 ***
    If strCmp <> MsgText(601) Then
        If InStr(strCmp, ",") > 0 Then
            st040Where = " And a0403 In ('" & Replace(strCmp, "+", "','") & "') "
        Else
            st040Where = " And a0403='" & strCmp & "' "
        End If
    End If
    
    stA2 = "Select Sum(Nvl(R004,0))-Sum(Nvl(R008,0)) From Accrpt44r0 D Where R003 is Null "
    stWhere = " And ID='" & strUserNum & "' And FormN='" & Me.Name & "' "
    
    '專利國內部-P
    i = 1 '顯示順序
    stA = "Update AccRpt44r0 A Set R001='" & i & "',R003='專利國內部-P',R004=(" & stA2 & stWhere & " And SubStr(R002,1,4)='4111' And R002<>'411107' And A.R010=D.R010 Group by R010) " & _
             "Where 1=1 " & stWhere & " And R002='4111' "
    adoTaie.Execute stA
    
    '專利國內部-CFP
    i = i + 1
    stA = "Update AccRpt44r0 A Set R001='" & i & "',R003='專利國內部-CFP',R004=(" & stA2 & stWhere & " And R002='413101' And A.R010=D.R010 Group by R010) " & _
             "Where 1=1 " & stWhere & " And R002='413101' "
    adoTaie.Execute stA
    
    '專利國外部-FCP
    i = i + 1
    stZeroAccNo = stZeroAccNo & ",F4104"
    'Modify by Amy 2022/03/28 原:,R003=R003||' - FCP':拿掉FCP文字-與秘書報表一致,已通知婧瑄
    stA = "Update AccRpt44r0 Set R001='" & i & "' Where 1=1 " & stWhere & " And R002='F4104' "
    adoTaie.Execute stA
    
    '專利日本部-FCP
    i = i + 1
    stZeroAccNo = stZeroAccNo & ",F4105"
    'Modify by Amy 2022/03/28 原:,R003=R003||' - FCP':拿掉FCP文字-與秘書報表一致,已通知婧瑄
    stA = "Update AccRpt44r0 Set R001='" & i & "' Where 1=1 " & stWhere & " And R002='F4105' "
    adoTaie.Execute stA
    
    '商標部-T
    i = i + 1
    'Memo 410105/410108 11001月前獨自列示-秘書報表,目前以新格式列示 410105/410108->商標收入
    stA = "Update AccRpt44r0 A Set R001='" & i & "',R003='商標部-T',R004=(" & stA2 & stWhere & " And ((SubStr(R002,1,4)='4101' And R002<>'410110') Or R002='417202T') And A.R010=D.R010 Group by R010) " & _
             "Where 1=1 " & stWhere & " And R002='4101' "
    adoTaie.Execute stA
    
    '商標部-CFT
    i = i + 1
    stA = "Update AccRpt44r0 A Set R001='" & i & "',R003='商標部-CFT',R004=(" & stA2 & stWhere & " And SubStr(R002,1,4)='4121' And R002<>'412102' And A.R010=D.R010 Group by R010) " & _
             "Where 1=1 " & stWhere & " And R002='4121' "
    adoTaie.Execute stA
    
   '商標部-FCT英文組
    i = i + 1
    stZeroAccNo = stZeroAccNo & ",F4106"
    stA = "Update AccRpt44r0 Set R001='" & i & "',R003='商標部 - '||R003 Where 1=1 " & stWhere & " And R002='F4106' "
    adoTaie.Execute stA
    
    '商標部-FCT日文組
    i = i + 1
    stZeroAccNo = stZeroAccNo & ",F4107"
    stA = "Update AccRpt44r0 Set R001='" & i & "',R003='商標部 - '||R003 Where 1=1 " & stWhere & " And R002='F4107' "
    adoTaie.Execute stA
    
    '商標部-著作權
    i = i + 1
    'Memo 415102 11001月前獨自列示-秘書報表,目前以新格式列示 415102->著作權
    stA = "Update AccRpt44r0 A Set R001='" & i & "',R003='商標部-著作權',R004=(" & stA2 & stWhere & " And SubStr(R002,1,4)='4151' And A.R010=D.R010 Group by R010) " & _
             "Where 1=1 " & stWhere & " And R002='415101' "
    adoTaie.Execute stA
    
   '創新業務部-ACS
    i = i + 1
    'Modify by Amy 2022/03/28 原:R003='創新業務部-ACS':拿掉ACS文字-與秘書報表一致,已通知婧瑄
    stA = "Update AccRpt44r0 A Set R001='" & i & "',R003='創新業務部',R004=(" & stA2 & stWhere & " And R002='420101' And A.R010=D.R010 Group by R010) " & _
             "Where 1=1 " & stWhere & " And R002='420101' "
    adoTaie.Execute stA
    
    'Modify by Amy 2022/02/17 +if 11101月起法務科目全改為 490102,因比較三年故判斷有法務資料才顯示
    If bolHasLaw = True Then
        '法務-P
        i = i + 1
        stZeroAccNo = stZeroAccNo & ",411107"
        stA = "Update AccRpt44r0 A Set R001='" & i & "',R003='法務-P',R004=(" & stA2 & stWhere & " And R002='411107' And A.R010=D.R010 Group by R010) " & _
                 "Where 1=1 " & stWhere & " And R002='411107' "
        adoTaie.Execute stA
        
        '法務-CFP
        i = i + 1
        stZeroAccNo = stZeroAccNo & ",413102"
        stA = "Update AccRpt44r0 A Set R001='" & i & "',R003='法務-CFP',R004=(" & stA2 & stWhere & " And R002='413102' And A.R010=D.R010 Group by R010) " & _
                 "Where 1=1 " & stWhere & " And R002='413102' "
        adoTaie.Execute stA
        
        '法務-FCP
        i = i + 1
        stZeroAccNo = stZeroAccNo & ",417103"
        stA = "Update AccRpt44r0 A Set R001='" & i & "',R003='法務-FCP',R004=(" & stA2 & stWhere & " And R002='417103' And A.R010=D.R010 Group by R010) " & _
                 "Where 1=1 " & stWhere & " And R002='417103' "
        adoTaie.Execute stA
        
       '法務-T
        i = i + 1
        stZeroAccNo = stZeroAccNo & ",410110"
        stA = "Update AccRpt44r0 A Set R001='" & i & "',R003='法務-T',R004=(" & stA2 & stWhere & " And R002='410110' And A.R010=D.R010 Group by R010) " & _
                 "Where 1=1 " & stWhere & " And R002='410110' "
        adoTaie.Execute stA
        
        '法務-CFT
        i = i + 1
        stZeroAccNo = stZeroAccNo & ",412102"
        stA = "Update AccRpt44r0 A Set R001='" & i & "',R003='法務-CFT',R004=(" & stA2 & stWhere & " And R002='412102' And A.R010=D.R010 Group by R010) " & _
                 "Where 1=1 " & stWhere & " And R002='412102' "
        adoTaie.Execute stA
        
        '法務-FCT
        i = i + 1
        stZeroAccNo = stZeroAccNo & ",417203"
        stA = "Update AccRpt44r0 A Set R001='" & i & "',R003='法務-FCT',R004=(" & stA2 & stWhere & " And R002='417203' And A.R010=D.R010 Group by R010) " & _
                 "Where 1=1 " & stWhere & " And R002='417203' "
        adoTaie.Execute stA
        
        '一般法務
        i = i + 1
        stZeroAccNo = stZeroAccNo & ",4141"
        stA = "Update AccRpt44r0 A Set R001='" & i & "',R003='法務-一般法務',R004=(" & stA2 & stWhere & " And SubStr(R002,1,4) In ('4141','4161','4181') And A.R010=D.R010 Group by R010) " & _
                 "Where 1=1 " & stWhere & " And R002='4141' "
        adoTaie.Execute stA
    End If
    
    'Add by Amy 2022/02/17 +490102 11101月起法務科目全改為 490102
    If ChkHasData(" And R002='490102' ") = True Then
        i = i + 1
        stA = "Update AccRpt44r0 A Set R001='" & i & "',R003='專業其他收入',R004=(" & stA2 & stWhere & " And R002='490102' And A.R010=D.R010 Group by R010) " & _
                 "Where 1=1 " & stWhere & " And R002='490102' "
        adoTaie.Execute stA
    End If
    
    'Modify by Amy 2022/02/17 11101月起改為 490102,因比較三年故判斷有資料才顯示
    If ChkHasData(" And R002='7121' ") = True Then
        '其他 (7121)
        i = i + 1
        'Memo 410105/410108/415102 11001月前獨自列示-秘書報表,目前以新格式列示 410105/410108->商標收入,415102->著作權
        stA = "Update AccRpt44r0 A Set R001='" & i & "',R003='其他',R004=(" & stA2 & stWhere & " And R002='7121' And A.R010=D.R010 Group by R010) " & _
                 "Where 1=1 " & stWhere & " And R002='7121' "
        adoTaie.Execute stA
    End If
    
    '110年前 F41XX/法務相關科目 值設 Null(由財務自行抓美珍報表填入,因格式不同法務資料抓法太複雜)
    stA = Replace(Mid(stZeroAccNo, 2), ",", "','")
    stA = "Update AccRpt44r0 A Set R004=0 Where 1=1 " & stWhere & " And R002 In ('" & stA & "') And SubStr(R010,1,Length(R010)-2)<110"
    adoTaie.Execute stA
    '*** End 更新最後抓取的資料值 ***
    
    '讀取資料
    'Memo by Amy 2021/10/18 安全基金(490101) 實績為0,此報表只顯示實績,故不需出現-秀玲
    stA2 = "Select Distinct R001 as Sort,R002 as AccNo,R003 as AccName From Accrpt44r0 Where R001 is not null " & stWhere
    stA = "Select " & GetField & " From Accrpt44r0 Where R001 is not null " & stWhere & " Group by R002 "
    stA = "Select AccName,M1,M2,M3,Y1,Y2,Y3 From (" & stA & "),(" & stA2 & ") " & _
              "Where AccNo=R002(+) Order by to_Number(Sort)"
    intA = 1
    Set rsNew = ClsLawReadRstMsg(intA, stA)
   
End Sub

'專業達成點數(實績點數比較三年)-於專業達成點數分佈情況(當月實際達成)產生工作表3
Private Sub ExcelSave4()
    Dim strTemp As String
    
    If rsNew.RecordCount = 0 Then Exit Sub
    
    m_lngRow = 1
    Call ExcelHead4
    
    rsNew.MoveFirst
    Do While rsNew.EOF = False
        
        For ii = LBound(strField) To UBound(strField)
            Select Case ii
                Case GetValue("單位")
                    strTemp = "" & rsNew.Fields("AccName")
                Case GetValue(lngYear & "." & lngMonth)
                    strTemp = "" & rsNew.Fields("M1")
                Case GetValue(lngYear - 1 & "." & lngMonth)
                    strTemp = "" & rsNew.Fields("M2")
                Case GetValue(lngYear - 2 & "." & lngMonth)
                    strTemp = "" & rsNew.Fields("M3")
                Case GetValue(lngYear & "年1-" & lngMonth & "月")
                    strTemp = "" & rsNew.Fields("Y1")
                Case GetValue(lngYear - 1 & "年1-" & lngMonth & "月")
                    strTemp = "" & rsNew.Fields("Y2")
                Case GetValue(lngYear - 2 & "年1-" & lngMonth & "月")
                    strTemp = "" & rsNew.Fields("Y3")
            End Select
            wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = strTemp
            If ii = GetValue("單位") Then
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).HorizontalAlignment = xlLeft
            Else
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).HorizontalAlignment = xlRight
                wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = "#,##0.00"
            End If
        Next ii
        m_lngRow = m_lngRow + 1
        rsNew.MoveNext
    Loop
    '全所
    For ii = LBound(strField) To UBound(strField)
        strTemp = "=Sum(" & Chr(intField + ii) & intTitleR + 1 & ":" & Chr(intField + ii) & m_lngRow - 1 & ")"
        If ii = GetValue("單位") Then
            strTemp = "全所"
            wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).HorizontalAlignment = xlLeft
        Else
            wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).HorizontalAlignment = xlRight
            wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).NumberFormatLocal = "#,##0.00"
        End If
        wksaccrpt424.Range(Chr(intField + ii) & m_lngRow).Value = strTemp
    Next ii
    '設定
    wksaccrpt424.Range(Chr(intField) & intTitleR & ":" & Chr(intField + UBound(strField)) & m_lngRow).Select
    xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Font.Size = 12
    m_lngRow = m_lngRow + 2
    wksaccrpt424.Range(Chr(intField) & m_lngRow).Font.Size = 11
    wksaccrpt424.Range(Chr(intField) & m_lngRow).Font.ColorIndex = 3
    wksaccrpt424.Range(Chr(intField) & m_lngRow).Value = "*11001月前專利國外部 - FCP/專利日本部 - FCP、商標部 - FCT英文組/日文組、法務相關因計算困難，故顯示0"
      
    wksaccrpt424.PageSetup.PaperSize = 9 'A4
    wksaccrpt424.PageSetup.Orientation = xlLandscape '橫印
    wksaccrpt424.PageSetup.LeftMargin = 28.34
    wksaccrpt424.PageSetup.RightMargin = 28.34
    wksaccrpt424.PageSetup.TopMargin = 42.51
    wksaccrpt424.PageSetup.BottomMargin = 42.51
    wksaccrpt424.PageSetup.HeaderMargin = 28.34
    wksaccrpt424.PageSetup.FooterMargin = 28.34
    wksaccrpt424.PageSetup.PrintTitleRows = "$1:$" & intTitleR '表頭保留列
    wksaccrpt424.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
End Sub

'專業達成點數(實績點數比較三年)-欄位 for ExcelSave4
Private Function GetField() As String
    Dim ii As Integer, stTmp(1) As String
    
    ReDim strField(0 To 6)
    ReDim intWidth(0 To 6)

    strField(0) = "單位": intWidth(0) = 25
    For ii = 1 To 3
        strField(ii) = Val(lngYear) + 1 - ii & "." & lngMonth
        intWidth(ii) = 15
        strField(ii + 3) = Val(lngYear) + 1 - ii & "年1-" & lngMonth & "月"
        intWidth(ii + 3) = 15
        
        'DB 語法
        stTmp(0) = stTmp(0) & ",Round(Sum(Decode(R010," & Val(lngYear) + 1 - ii & Format(lngMonth, "00") & ",Nvl(R004,0)))/1000,2) as M" & ii
        stTmp(1) = stTmp(1) & ",Round(Sum(Decode(R010," & Val(lngYear) + 1 - ii & "13,Nvl(R004,0)))/1000,2) as Y" & ii
    Next ii
    GetField = "R002" & stTmp(0) & stTmp(1)
End Function

Private Sub ExcelHead4()
    Dim i As Integer
    
    wksaccrpt424.Range(Chr(intField) & m_lngRow).Value = lngYear & "年" & lngMonth & "月專業達成點數 (實績點數比較三年)"
    wksaccrpt424.Range(Chr(intField) & m_lngRow & ":" & Chr(intField + UBound(strField)) & m_lngRow).Select
    With xlsSalesPoint.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = True
    End With
    wksaccrpt424.Application.Selection.Font.Bold = True
    wksaccrpt424.Application.Selection.Font.Size = 14
    m_lngRow = m_lngRow + 1
    
    For i = LBound(strField) To UBound(strField)
        'Add by Amy 2021/11/10 儲存格為110.10 會為數字格式,顯示會變成110.1
        wksaccrpt424.Range(Chr(intField + i) & m_lngRow).NumberFormatLocal = "@"
        wksaccrpt424.Range(Chr(intField + i) & m_lngRow).Value = strField(i)
        wksaccrpt424.Range(Chr(intField + i) & m_lngRow).HorizontalAlignment = xlCenter
        wksaccrpt424.Columns(Chr(intField + i)).ColumnWidth = intWidth(i)
    Next i
    intTitleR = m_lngRow
    m_lngRow = m_lngRow + 1
End Sub

'Add by Amy 2022/02/17 是否有傳票資料
Private Function ChkHasData(ByVal stWhere As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    ChkHasData = False
    strQ = "Select * From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R001 is  null " & _
               stWhere
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        ChkHasData = True
    End If
    Set RsQ = Nothing
End Function
