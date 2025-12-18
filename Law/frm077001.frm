VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm077001 
   BorderStyle     =   1  '單線固定
   Caption         =   "顧問記錄"
   ClientHeight    =   5760
   ClientLeft      =   1068
   ClientTop       =   2688
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8952
   Begin VB.TextBox textCP113 
      Height          =   300
      Left            =   1170
      MaxLength       =   5
      TabIndex        =   6
      Top             =   2486
      Width           =   732
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1170
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1778
      Width           =   2205
   End
   Begin VB.TextBox textCP43 
      Height          =   300
      Left            =   5730
      MaxLength       =   9
      TabIndex        =   11
      Top             =   3195
      Width           =   2172
   End
   Begin VB.TextBox textCP26 
      Height          =   300
      Left            =   1380
      MaxLength       =   20
      TabIndex        =   10
      Top             =   3195
      Width           =   372
   End
   Begin VB.TextBox textCP06 
      Height          =   300
      Left            =   1170
      MaxLength       =   7
      TabIndex        =   8
      Top             =   2840
      Width           =   1215
   End
   Begin VB.TextBox textCP07 
      Height          =   300
      Left            =   5760
      MaxLength       =   7
      TabIndex        =   9
      Top             =   2840
      Width           =   1095
   End
   Begin VB.TextBox textCP21 
      Height          =   300
      Left            =   5760
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2486
      Width           =   372
   End
   Begin VB.TextBox textCP29 
      Height          =   300
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   5
      Top             =   2132
      Width           =   732
   End
   Begin VB.TextBox textCP14 
      Height          =   300
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   4
      Top             =   2132
      Width           =   732
   End
   Begin VB.TextBox textCP10 
      Height          =   300
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1424
      Width           =   732
   End
   Begin VB.TextBox textCP05 
      Height          =   300
      Left            =   1170
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1424
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8100
      TabIndex        =   22
      Top             =   30
      Width           =   800
   End
   Begin VB.TextBox textHCKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   488
      Width           =   1755
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7275
      TabIndex        =   19
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdCaseProgress 
      Caption         =   "案件進度(&C)"
      Height          =   400
      Left            =   6045
      TabIndex        =   17
      Top             =   30
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   915
      Left            =   1050
      TabIndex        =   13
      Top             =   4290
      Width           =   7755
      _ExtentX        =   13674
      _ExtentY        =   1609
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Left            =   5760
      TabIndex        =   3
      Top             =   1778
      Width           =   2205
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3889;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   60
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   5310
      Width           =   8775
      VariousPropertyBits=   671105055
      BackColor       =   16777215
      Size            =   "15478;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   660
      Left            =   1050
      TabIndex        =   12
      Top             =   3540
      Width           =   7335
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "12938;1164"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP29_2 
      Height          =   300
      Left            =   6555
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2132
      Width           =   2295
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4048;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14_2 
      Height          =   300
      Left            =   1980
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2132
      Width           =   2175
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "3836;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP10_2 
      Height          =   300
      Left            =   6555
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1424
      Width           =   2295
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4048;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textHC09_2 
      Height          =   300
      Left            =   2805
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   488
      Width           =   1515
      VariousPropertyBits=   671105055
      ForeColor       =   255
      MaxLength       =   20
      Size            =   "2672;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "（0： 存檔自動上發文日）"
      Height          =   225
      Index           =   4
      Left            =   1980
      TabIndex        =   48
      Top             =   2524
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "工作時數："
      Height          =   225
      Index           =   20
      Left            =   90
      TabIndex        =   47
      Top             =   2524
      Width           =   975
   End
   Begin MSForms.Label lblCase 
      Height          =   255
      Index           =   6
      Left            =   6690
      TabIndex        =   46
      Top             =   1115
      Width           =   555
      BackColor       =   16777152
      VariousPropertyBits=   27
      Caption         =   "Case(6)"
      Size            =   "979;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblCase 
      Height          =   255
      Index           =   5
      Left            =   5760
      TabIndex        =   45
      Top             =   1115
      Width           =   555
      BackColor       =   16777152
      VariousPropertyBits=   27
      Caption         =   "Case(5)"
      Size            =   "979;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblCase 
      Height          =   255
      Index           =   4
      Left            =   2430
      TabIndex        =   44
      Top             =   1115
      Width           =   735
      BackColor       =   16777152
      VariousPropertyBits=   27
      Caption         =   "Case(4)"
      Size            =   "1296;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblCase 
      Height          =   255
      Index           =   3
      Left            =   1350
      TabIndex        =   43
      Top             =   1115
      Width           =   735
      BackColor       =   16777152
      VariousPropertyBits=   27
      Caption         =   "Case(3)"
      Size            =   "1296;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblCase 
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   42
      Top             =   806
      Width           =   7635
      BackColor       =   16777152
      VariousPropertyBits=   27
      Caption         =   "Case(2)"
      Size            =   "13467;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCase 
      Height          =   255
      Index           =   1
      Left            =   6300
      TabIndex        =   41
      Top             =   488
      Width           =   555
      BackColor       =   16777152
      VariousPropertyBits=   27
      Caption         =   "Case(1)"
      Size            =   "979;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblCase 
      Height          =   255
      Index           =   0
      Left            =   5430
      TabIndex        =   40
      Top             =   488
      Width           =   555
      BackColor       =   16777152
      VariousPropertyBits=   27
      Caption         =   "Case(0)"
      Size            =   "979;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin VB.Label Label40 
      Caption         =   "本案期限："
      Height          =   225
      Left            =   90
      TabIndex        =   39
      Top             =   4350
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "進度備註："
      Height          =   225
      Index           =   14
      Left            =   90
      TabIndex        =   38
      Top             =   3570
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "相關總收文號："
      Height          =   225
      Index           =   15
      Left            =   4440
      TabIndex        =   37
      Top             =   3233
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "是否算案件數：　　　(N:不算)"
      Height          =   225
      Index           =   13
      Left            =   60
      TabIndex        =   36
      Top             =   3233
      Width           =   2595
   End
   Begin VB.Label Label1 
      Caption         =   "本所期限："
      Height          =   225
      Index           =   12
      Left            =   90
      TabIndex        =   35
      Top             =   2878
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "法定期限："
      Height          =   225
      Index           =   17
      Left            =   4470
      TabIndex        =   34
      Top             =   2878
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "是否取締案：　　　   (Y:取締)"
      Height          =   225
      Index           =   18
      Left            =   4470
      TabIndex        =   33
      Top             =   2524
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "業  務  區："
      Height          =   225
      Index           =   10
      Left            =   90
      TabIndex        =   32
      Top             =   1816
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "諮詢人員："
      Height          =   225
      Index           =   1
      Left            =   4470
      TabIndex        =   31
      Top             =   1816
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "協辦人員："
      Height          =   225
      Index           =   19
      Left            =   4470
      TabIndex        =   30
      Top             =   2115
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "承  辦  人："
      Height          =   225
      Index           =   3
      Left            =   90
      TabIndex        =   28
      Top             =   2170
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   225
      Index           =   2
      Left            =   4470
      TabIndex        =   26
      Top             =   1462
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "收  文  日："
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   24
      Top             =   1462
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "本次期間累計：　　　   次　　　　時"
      Height          =   225
      Index           =   9
      Left            =   4470
      TabIndex        =   23
      Top             =   1115
      Width           =   3345
   End
   Begin VB.Label Label1 
      Caption         =   "本案累計：　　　   次　　　　時"
      Height          =   225
      Index           =   8
      Left            =   4470
      TabIndex        =   21
      Top             =   508
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "本次聘任時間：　　  　　－"
      Height          =   225
      Index           =   7
      Left            =   90
      TabIndex        =   20
      Top             =   1115
      Width           =   3405
   End
   Begin VB.Label Label1 
      Caption         =   "當  事  人："
      Height          =   225
      Index           =   5
      Left            =   90
      TabIndex        =   18
      Top             =   825
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   220
      Index           =   6
      Left            =   90
      TabIndex        =   16
      Top             =   510
      Width           =   900
   End
End
Attribute VB_Name = "frm077001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/11 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB 、texCUID、textCP64、lblCase(index)、textHC09_2、textCP10_2、textCP14_2、textCP29_2、Combo2
'Create by Lydia 2020/04/20 顧問記錄(LA-999999)
Option Explicit

'本所案號
Dim m_HC01 As String
Dim m_HC02 As String
Dim m_HC03 As String
Dim m_HC04 As String
'是否閉卷
Dim m_HC09 As String

Dim m_CPKeyList() As String
Dim m_CPKeyCount As Integer
'收文日
Dim m_CP05 As String
'收文號
Dim m_CP09 As String
'案件性質
Dim m_CP10 As String
'相關總收文號
Dim m_CP43 As String
'本次聘任期間：收文號、聘任期間起迄
Dim m_Key09 As String, m_Key53 As String, m_Key54 As String

'宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type

'儲存案件進度檔檔案欄位的串列
Dim m_CPList() As FIELDITEM
Dim m_CPCount As Integer

Dim intLastRow As Integer '記錄勾選最後一筆

Dim strTmpQ As String
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset
'Added by Lydia 2020/05/27
Dim m_EditMode As Integer '操作模式: 0-查詢, 1-新增
Public m_PrevForm As Form '前一畫面
Dim m_ShowList As Variant '前一畫面傳入多筆收文號
Dim xShow As Integer, maxShow As Integer '目前第X筆收文號/最大數字

Public Sub SetData(ByVal strData As String, ByVal nType As Integer, ByVal bClear As Boolean)
   If bClear Then
      m_HC01 = Empty
      m_HC02 = Empty
      m_HC03 = Empty
      m_HC04 = Empty
      m_CP10 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      Case 0: m_HC01 = strData
      Case 1: m_HC02 = strData
      Case 2: m_HC03 = strData & String(1 - Len(strData), "0")
      Case 3: m_HC04 = strData & String(2 - Len(strData), "0")
      Case 4: m_CP10 = strData
      Case 6:
         m_CP43 = strData
         textCP43 = m_CP43
      Case 7:
         m_CP09 = strData
   End Select
End Sub

Private Sub ClearAll(ByVal bolReset As Boolean)
Dim oLbl As Control

   If bolReset = True Then
        '設定控制項的背景顏色
        textHCKey.BackColor = &H8000000F
        textHC09_2.BackColor = &H8000000F
        textCP10_2.BackColor = &H8000000F
        textCP14_2.BackColor = &H8000000F
        textCP29_2.BackColor = &H8000000F
        For Each oLbl In lblCase
           oLbl.BackColor = &H8000000F
        Next
   End If
   
   textHCKey = Empty
   textHC09_2 = Empty
   textCP05 = Empty
   textCP06 = Empty
   textCP07 = Empty
   textCP10 = Empty
   textCP10_2 = Empty
   textCP14 = Empty
   textCP14_2 = Empty
   textCP21 = Empty
   textCP26 = Empty
   textCP29 = Empty
   textCP29_2 = Empty
   textCP43 = Empty
   textCP64 = Empty
   textCP113 = Empty
   For Each oLbl In lblCase
      oLbl.Caption = Empty
   Next
   m_Key09 = Empty
   m_Key53 = Empty
   m_Key54 = Empty
   If bolReset = False Then
       Combo1.ListIndex = 0
       Combo2.ListIndex = 0
   End If
End Sub

Private Sub cmdCaseProgress_Click()
   frm010012_03.SetData 0, m_HC01, True
   frm010012_03.SetData 1, m_HC02, False
   frm010012_03.SetData 2, m_HC03, False
   frm010012_03.SetData 3, m_HC04, False
   frm010012_03.SetData 4, m_CP09, False
   frm010012_03.SetParent Me
   Me.Hide
   frm010012_03.Show
   frm010012_03.QueryData
End Sub

Private Sub cmdExit_Click()
   'Added by Lydia 2020/05/27 顯示下一筆
   If m_EditMode = 0 And m_CP09 <> "" And xShow < maxShow Then
        xShow = xShow + 1
        m_CP09 = m_ShowList(xShow)
        QueryData
   Else
   'end 2020/05/27
        Unload Me
   End If 'Added by Lydia 2020/05/27
End Sub

Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid = True Then
      If ValidateInput() = False Then
         Exit Sub
      End If
      '所有內部收文, 若有輸入本所期限或法定期限者, 檢查期限不可小於系統日
      'Modified by Lyddia 2023/11/08 傳入必需欄位
      'If PUB_CheckCP0607(0, textCP06, textCP07) = False Then Exit Sub
      If PUB_CheckCP0607(0, textCP06, textCP07, "", "000", m_HC01, textCP10) = False Then Exit Sub
      
      If Val(textCP113) = 0 Then
          If MsgBox("工作時數為0存檔自動上發文日，是否繼續？", vbYesNo + vbDefaultButton1) = vbNo Then
              textCP113.SetFocus
              textCP113_GotFocus
              Exit Sub
          End If
      End If
      
      cmdOK.Enabled = False 'Added by Lydia 2020/05/11 控制只能按一次; 5/8有滑鼠按二次產生二道進度
      
      '設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      '儲存資料
      OnUpdateField

      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      '設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      Unload Me
      
      '保留: 可連續輸入
      'Call ClearAll(False)
      'Call QueryData
      'SetInputEntry
   End If
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
Dim intP As Integer

   If Combo1.Tag <> Combo1.Text Then
        intP = GetCmbVal(Combo1, Trim(Combo1.Text), strExc(1))
        Combo1.ListIndex = intP
        
        Combo2.Text = ""
        Call SetCombo2(IIf(intP = 0, "", Trim(Left(Combo1.Text, 4))))
   End If
   Combo1.Tag = Combo1.Text
End Sub

'Modified by Lydia 2022/02/11 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub Combo2_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
Dim intP As Integer

   If Trim(Combo2.Text) <> "" And Combo2.Tag <> Combo2.Text Then
        intP = GetCmbVal(Combo2, Trim(Combo2.Text), strExc(1))
        If intP = 0 Then '已有業務區再輸入之諮詢人員不在清單，重設業務區和諮詢人員
            If ChkExistStaff(Combo2.Text, strExc(4), strExc(5), strExc(6)) = True Then
                Combo1.Text = strExc(6)
                Call Combo1_Validate(False)  '業務區
                Combo2.Text = convForm(strExc(4), 6) & " " & strExc(5)
            Else
                Combo2.SetFocus
                Combo2.Tag = Combo2.Text
                Exit Sub
            End If
        Else
             Combo2.ListIndex = intP
        End If
   End If
   Combo2.Tag = Combo2.Text
   
   '未有業務區再輸入之諮詢人員，重設業務區和諮詢人員
   If Trim(Combo1.Text) = "" And Trim(Combo2.Text) <> "" Then
        'Modified by Lydia 2020/05/14 改ST15
        'strExc(1) = PUB_GetST03(Trim(Left(Combo2.Text, 6)))
        strExc(1) = GetST15(Trim(Left(Combo2.Text, 6)))
        strExc(2) = Combo2.Text
        If strExc(1) <> "" Then
            Combo1.Text = strExc(1)
            Call Combo1_Validate(False)  '業務區
            Combo2.Text = strExc(2)
        End If
   End If
End Sub

'畫面被載入時
Private Sub Form_Load()
   MoveFormToCenter Me
   
   Call ClearAll(True)
   '預設諮詢人員(CP12,CP13)下拉選單
   Call SetCombo1
   Call SetCombo2("")
   
   'Added by Lydia 2020/05/27
   textCUID.BackColor = &H8000000F
   textCUID = Empty
   '操作模式: 0-查詢, 1-新增
   If m_CP09 <> "" Then
       m_EditMode = 0
       cmdOK.Visible = False  '隱藏「確定」
       cmdExit.Caption = "下一筆"
       m_ShowList = Empty
       m_ShowList = Split(m_CP09, ",")
       m_CP09 = m_ShowList(0)
       xShow = 0
       maxShow = UBound(m_ShowList)
   Else
       m_EditMode = 1
   End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Added by Lydia 2020/05/27
   If TypeName(m_PrevForm) <> "Nothing" Then
       m_PrevForm.Show
   End If
   'end 2020/05/27
   
   Set rsQuery = Nothing
   Set frm077001 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
Dim intRow As Integer

   If MSHFlexGrid1.Rows < 2 Then Exit Sub
   
   With MSHFlexGrid1
       If .MouseRow > 0 Then
          intRow = .MouseRow
          GridClick MSHFlexGrid1, intLastRow, 0, 1 '傳入最後一筆勾選，複選
          
          If MSHFlexGrid1.TextMatrix(intRow, 0) = "v" Then
             '若本所期限沒值時，以勾的該筆代 智權人員、本所期限，法定期限，備註，相關總收文號 到上方
                If textCP06.Text = "" Then
                   textCP06 = ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(intRow, 2))
                   textCP07 = ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(intRow, 3))
                   textCP43 = MSHFlexGrid1.TextMatrix(intRow, 8)
                   textCP64 = textCP64 & IIf(MSHFlexGrid1.TextMatrix(intRow, 6) <> "", ";" & MSHFlexGrid1.TextMatrix(intRow, 6), "")
                End If
                If Trim(Combo2.Text) = "" Then
                    'Modified by Lydia 2020/05/14 改ST15
                    'strExc(1) = PUB_GetST03(MSHFlexGrid1.TextMatrix(intRow, 11))
                    strExc(1) = GetST15(MSHFlexGrid1.TextMatrix(intRow, 11))
                    If strExc(1) <> "" Then
                        Combo1.Text = strExc(1)
                        Call Combo1_Validate(False)  '業務區
                        Combo2.Text = MSHFlexGrid1.TextMatrix(intRow, 11)
                        Call Combo2_Validate(False)  '諮詢人員
                    End If
                End If
          End If
       End If
   End With
  
End Sub

'清除案件進度檔檔案欄位串列
Private Sub ClearCPFieldList()
   If m_CPCount > 0 Then
      Erase m_CPList
   End If
   m_CPCount = 0
End Sub

'設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiOldData = strFieldData
         m_CPList(nPos).fiNewData = strFieldData
         m_CPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_CPList(m_CPCount + 1)
      m_CPList(m_CPCount).fiName = strFieldName
      m_CPList(m_CPCount).fiOldData = strFieldData
      m_CPList(m_CPCount).fiNewData = strFieldData
      m_CPList(m_CPCount).fiType = nFieldType
      m_CPCount = m_CPCount + 1
   End If
End Sub

'設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

'更新欄位的內容
Private Sub OnUpdateField()
   SetCPFieldNewData "CP01", m_HC01
   SetCPFieldNewData "CP02", m_HC02
   SetCPFieldNewData "CP03", m_HC03
   SetCPFieldNewData "CP04", m_HC04
   '收文日
   If IsEmptyText(textCP05) = False Then
      SetCPFieldNewData "CP05", DBDATE(textCP05)
   Else
      SetCPFieldNewData "CP05", Empty
   End If
   '本所期限
   If IsEmptyText(textCP06) = False Then
      SetCPFieldNewData "CP06", DBDATE(textCP06)
   Else
      SetCPFieldNewData "CP06", Empty
   End If
   '法定期限
   If IsEmptyText(textCP07) = False Then
      SetCPFieldNewData "CP07", DBDATE(textCP07)
   Else
      SetCPFieldNewData "CP07", Empty
   End If
   '收文號
   If m_CP09 = "" Then
       m_CP09 = AutoNo("B", 6)
   End If
   SetCPFieldNewData "CP09", m_CP09
   '案件性質
   SetCPFieldNewData "CP10", textCP10
   '業務區
   'Modified by Lydia 2020/05/14 改ST15
   'SetCPFieldNewData "CP12", PUB_GetST03(Trim(Left(Combo2.Text, 6)))
   SetCPFieldNewData "CP12", GetST15(Trim(Left(Combo2.Text, 6)))
   '智權人員=諮詢人員
   SetCPFieldNewData "CP13", Trim(Left(Combo2.Text, 6))
   '承辦人員
   SetCPFieldNewData "CP14", textCP14
   
   SetCPFieldNewData "CP11", "04" '04-同仁介紹
   SetCPFieldNewData "CP20", "N"
   SetCPFieldNewData "CP32", "N"
   
   '是否取締案
   SetCPFieldNewData "CP21", textCP21
   '是否算案件數
   SetCPFieldNewData "CP26", textCP26
   '協辦人員
   If Not IsEmptyText(textCP29) Then
      SetCPFieldNewData "CP29", textCP29
   Else
      SetCPFieldNewData "CP29", Empty
   End If
   '相關總收文號
   SetCPFieldNewData "CP43", textCP43
   '案件進度
   SetCPFieldNewData "CP64", ChgSQL(textCP64)
   '工作時數
   SetCPFieldNewData "CP113", Val(textCP113)
   
End Sub

Private Function OnSaveData() As Boolean
Dim strSql As String
Dim intP As Integer

On Error GoTo ErrorHandler

OnSaveData = True
cnnConnection.BeginTrans

   '新增-案件進度檔
   If SaveNewCaseProgress = False Then GoTo ErrorHandler
    
   '工作時數欄可不輸入,在存檔時同時上發文日
   If Val(textCP113) = 0 Then
      strSql = "UPDATE CASEPROGRESS SET CP27=" & strSrvDate(1) & _
               "WHERE CP09 = '" & m_CP09 & "'"
      cnnConnection.Execute strSql
   End If

    '更新下一程序
    With MSHFlexGrid1
       For intP = 1 To .Rows - 1
          If .TextMatrix(intP, 0) = "v" And "" & .TextMatrix(intP, 8) <> "" Then
             strExc(1) = .TextMatrix(intP, 8)
             strExc(2) = .TextMatrix(intP, 9)
             strExc(3) = .TextMatrix(intP, 10)
             strSql = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & strExc(1) & "' AND " & _
                "NP07=" & strExc(2) & " AND NP22=" & strExc(3)
              cnnConnection.Execute strSql
          End If
       Next intP
    End With
 
cnnConnection.CommitTrans
Exit Function

ErrorHandler:
    OnSaveData = False
    cnnConnection.RollbackTrans

End Function

'新增案件進度檔
Private Function SaveNewCaseProgress() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   
On Error GoTo ErrorHandler
SaveNewCaseProgress = True

   strSql = "INSERT INTO CaseProgress ("
   For nIndex = 0 To m_CPCount - 1
      If Not IsEmptyText(m_CPList(nIndex).fiNewData) Then
         If nIndex <> 0 Then strSql = strSql & ","
         strSql = strSql & m_CPList(nIndex).fiName
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   For nIndex = 0 To m_CPCount - 1
      If Not IsEmptyText(m_CPList(nIndex).fiNewData) Then
         If nIndex <> 0 Then strSql = strSql & ","
         If m_CPList(nIndex).fiType = 0 Then
            strSql = strSql & "'" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
         Else
            strSql = strSql & m_CPList(nIndex).fiNewData
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   
   cnnConnection.Execute strSql
Exit Function
ErrorHandler:
    SaveNewCaseProgress = False
End Function

'讀取資料庫
Public Sub QueryData()
   
   textHC09_2 = Empty
   m_HC09 = Empty
   
   '先清除案件進度檔欄位串列
   ClearCPFieldList
   
   m_CP05 = strSrvDate(2)
   
   textCP05 = m_CP05
   textCP10 = m_CP10
   textCP10_Validate False
   
   If Not (m_EditMode = 0 And m_CP09 <> "" And xShow > 0) Then 'Added by Lydia 2020/05/27 查詢：只在第一筆抓固定資料
        '本所案號
        textHCKey = m_HC01 & "-" & m_HC02 & IIf(m_HC03 <> "0", "-" & m_HC03, "") & IIf(m_HC04 <> "00", "-" & m_HC04, "")
           
        '顧問基本檔
        strTmpQ = "SELECT HIRECASE.*,NVL(CU04,NVL(CU05,CU06)) CNAME " & _
                    "FROM HIRECASE,CUSTOMER " & _
                    "WHERE HC01 = '" & m_HC01 & "' AND HC02 = '" & m_HC02 & "' AND HC03 = '" & m_HC03 & "' AND HC04 = '" & m_HC04 & "'" & _
                    "AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) "
        intQ = 1
        Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
        If intQ = 1 Then
            With rsQuery
                 lblCase(2).Caption = "" & .Fields("CNAME")  '當事人1
                 m_HC09 = "" & .Fields("HC09")
            End With
        End If
        
        '是否閉卷
        If m_HC09 = "Y" Then
           textHC09_2 = "本案已閉卷"
        Else
           textHC09_2 = Empty
        End If
        
        strExc(1) = "": strExc(2) = ""
        If GetTimes_LA999999("1", strExc(1), strExc(2)) = True Then
            lblCase(0).Caption = Val(strExc(1))  '本案-累計次數
            lblCase(1).Caption = Val(strExc(2))  '本案-累計時數
        End If
        
        '本次聘任期間
        m_Key09 = "": m_Key53 = "": m_Key54 = ""
        strExc(1) = "": strExc(2) = ""
        If GetTimes_LA999999("2", strExc(1), strExc(2), m_Key09, m_Key53, m_Key54) = False Then
            Unload Me
            Exit Sub
        Else
            lblCase(3).Caption = TransDate(m_Key53, 1)    '本次聘任期間(起)
            lblCase(4).Caption = TransDate(m_Key54, 1)    '本次聘任期間(迄)
            lblCase(5).Caption = Val(strExc(1))  '本次聘任期間-累計次數
            lblCase(6).Caption = Val(strExc(2))  '本次聘任期間-累計時數
        End If
        Set rsQuery = Nothing
   End If 'Added by Lydia 2020/05/27 查詢：只在第一筆抓固定資料
   
   '取得案件進度檔的欄位
   QueryCaseProgressWithNewCP
   
   If Not (m_EditMode = 0 And m_CP09 <> "" And xShow > 0) Then 'Added by Lydia 2020/05/27 查詢：只在第一筆抓固定資料
        '更新本案期限的資料
        UpdateGrdList m_HC01, m_HC02, m_HC03, m_HC04
   End If 'Added by Lydia 2020/05/27
   
   '設定輸入的位置
   SetInputEntry

End Sub

'取得案件進度檔的欄位內容
Private Sub QueryCaseProgressWithNewCP()
'Added by Lydia 2020/05/27
Dim strA1 As String, intA As Integer
Dim rsAD As New ADODB.Recordset

   'Added by Lydia 2020/05/27 查詢
   If m_EditMode = 0 And m_CP09 <> "" Then
       strA1 = "SELECT * FROM CASEPROGRESS WHERE CP09='" & m_CP09 & "' "
       intA = 1
       Set rsAD = ClsLawReadRstMsg(intI, strA1)
       If intA = 1 Then
            SetCPFieldOldData "CP01", m_HC01, 0
            SetCPFieldOldData "CP02", m_HC02, 0
            SetCPFieldOldData "CP03", m_HC03, 0
            SetCPFieldOldData "CP04", m_HC04, 0
            '收文日
            textCP05 = TransDate("" & rsAD.Fields("CP05"), 1)
            SetCPFieldOldData "CP05", "" & rsAD.Fields("CP05"), 1
            '本所期限
            textCP06 = TransDate("" & rsAD.Fields("CP06"), 1)
            SetCPFieldOldData "CP06", "" & rsAD.Fields("CP06"), 1
            '法定期限
            textCP07 = TransDate("" & rsAD.Fields("CP07"), 1)
            SetCPFieldOldData "CP07", "" & rsAD.Fields("CP07"), 1
            '收文號
            SetCPFieldOldData "CP09", m_CP09, 0
            '案件性質
            textCP10 = "" & rsAD.Fields("CP10")
            Call textCP10_Validate(False)
            SetCPFieldOldData "CP10", "" & rsAD.Fields("CP10"), 0
            '業務區
            If "" & rsAD.Fields("CP12") <> "" Then
                 Combo1.Text = "" & rsAD.Fields("CP12")
                 Call Combo1_Validate(False)
            End If
            SetCPFieldOldData "CP12", "" & rsAD.Fields("CP12"), 0
            '智權人員=諮詢人員
            If "" & rsAD.Fields("CP13") <> "" Then
                 strExc(1) = GetStaffName("" & rsAD.Fields("CP13"))
                 Call Combo1_Validate(False)
                 Combo2.Text = convForm(rsAD.Fields("CP13"), 6) & " " & strExc(1)
            End If
            SetCPFieldOldData "CP13", "" & rsAD.Fields("CP13"), 0
            '承辦人員
            textCP14 = "" & rsAD.Fields("CP14")
            Call textCP14_Validate(False)
            SetCPFieldOldData "CP14", "" & rsAD.Fields("CP14"), 0
            '是否取締案
            textCP21 = "" & rsAD.Fields("CP21")
            SetCPFieldOldData "CP21", "" & rsAD.Fields("CP21"), 0
            '是否算案件數
            textCP26 = "" & rsAD.Fields("CP26")
            SetCPFieldOldData "CP26", "" & rsAD.Fields("CP26"), 0
            '協辦人員
            textCP29 = "" & rsAD.Fields("CP29")
            Call textCP29_Validate(False)
            SetCPFieldOldData "CP29", "" & rsAD.Fields("CP29"), 0
            '相關總收文號
            textCP43 = "" & rsAD.Fields("CP43")
            SetCPFieldOldData "CP43", "" & rsAD.Fields("CP43"), 0
            '案件進度
            textCP64 = "" & rsAD.Fields("CP64")
            SetCPFieldOldData "CP64", "" & rsAD.Fields("CP64"), 0
            '工作時數
            textCP113 = "" & rsAD.Fields("CP113")
            SetCPFieldOldData "CP113", Val("" & rsAD.Fields("CP113")), 1
            '費用
            If "" & rsAD.Fields("CP16") = "" Then
                SetCPFieldOldData "CP16", Empty, 1
            Else
                SetCPFieldOldData "CP16", Val("" & rsAD.Fields("CP16")), 1
            End If
            '規費
            If "" & rsAD.Fields("CP17") = "" Then
                SetCPFieldOldData "CP17", Empty, 1
            Else
                SetCPFieldOldData "CP17", Val("" & rsAD.Fields("CP17")), 1
            End If
            '點數
            If "" & rsAD.Fields("CP18") = "" Then
                SetCPFieldOldData "CP18", Empty, 1
            Else
                SetCPFieldOldData "CP18", Val("" & rsAD.Fields("CP18")), 1
            End If
            '後金
            If "" & rsAD.Fields("CP19") = "" Then
                SetCPFieldOldData "CP19", Empty, 1
            Else
                SetCPFieldOldData "CP19", Val("" & rsAD.Fields("CP19")), 1
            End If
            
            SetCPFieldOldData "CP11", "" & rsAD.Fields("CP11"), 0
            SetCPFieldOldData "CP20", "" & rsAD.Fields("CP20"), 0
            SetCPFieldOldData "CP32", "" & rsAD.Fields("CP32"), 0
            
            Call UpdateCUID(rsAD)
       End If
       Set rsAD = Nothing
       Call TxtLocked(True) 'Added by Lydia 2020/05/27 鎖定畫面欄位
   Else
   'end 2020/05/27
       SetCPFieldOldData "CP01", Empty, 0
       SetCPFieldOldData "CP02", Empty, 0
       SetCPFieldOldData "CP03", Empty, 0
       SetCPFieldOldData "CP04", Empty, 0
       '收文日
       SetCPFieldOldData "CP05", Empty, 1
       '本所期限
       SetCPFieldOldData "CP06", Empty, 1
       '法定期限
       SetCPFieldOldData "CP07", Empty, 1
       '收文號
       SetCPFieldOldData "CP09", Empty, 0
       '案件性質
       SetCPFieldOldData "CP10", Empty, 0
    
       '業務區
       SetCPFieldOldData "CP12", Empty, 0
       '智權人員
       SetCPFieldOldData "CP13", Empty, 0
       '承辦人員
       SetCPFieldOldData "CP14", Empty, 0
       '費用
       SetCPFieldOldData "CP16", Empty, 1
       '規費
       SetCPFieldOldData "CP17", Empty, 1
       '點數
       SetCPFieldOldData "CP18", 0, 1
       '後金
       SetCPFieldOldData "CP19", Empty, 1
       '是否取締案
       SetCPFieldOldData "CP21", Empty, 0
       '是否算案件數
       SetCPFieldOldData "CP26", Empty, 0
       '協辦人員
       SetCPFieldOldData "CP29", Empty, 0
       '相關總收文號
       SetCPFieldOldData "CP43", Empty, 0
       '進度備註
       SetCPFieldOldData "CP64", Empty, 0
       '工作時數
       SetCPFieldOldData "CP113", 0, 1
       
       '因為會有些值沒有先定義，所以會沒有更新
       SetCPFieldOldData "CP11", Empty, 0
       SetCPFieldOldData "CP20", Empty, 0
       SetCPFieldOldData "CP32", Empty, 0
       Call TxtLocked(False) 'Added by Lydia 2020/05/27 不鎖定畫面欄位
   End If 'Added by Lydia 2020/05/27
End Sub

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer, iR As Integer

   arrGridHeadText = Array("v", "下一程序", "本所期限", "法定期限", "機關文號", "相關人", "備註", "解除期限日", "收文號", "下一程序代號", "序號", "智權人員")
   arrGridHeadWidth = Array(300, 1000, 900, 900, 1000, 1000, 1000, 1200, 0, 0, 0, 0)
   
   MSHFlexGrid1.Visible = False
   MSHFlexGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
         MSHFlexGrid1.Clear
         MSHFlexGrid1.Rows = 2
   End If
       
    For iRow = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.row = 0
       MSHFlexGrid1.col = iRow
       MSHFlexGrid1.Text = arrGridHeadText(iRow)
       MSHFlexGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
    Next

   MSHFlexGrid1.Visible = True
End Sub

Private Sub UpdateGrdList(ByVal strHC01 As String, ByVal strHC02 As String, ByVal strHC03 As String, ByVal strHC04 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
      
   Call SetGrd(True)
   
   strSql = "select ''as v,cpm03,sqldatet(np08) np08,sqldatet(np09) np09,np13,np14,np15,sqldatet(np11) np11,np01,np07,np22,np10 " & _
               "From nextprogress, Casepropertymap " & _
               "Where np02='" & m_HC01 & "' AND np03='" & m_HC02 & "' AND np04='0' AND np05='00' AND Nvl(np06,'N')='N'" & _
               "and np02=cpm01(+) and np07=cpm02(+) order by np08,np09,np22 "
   intQ = 1
   Set rsTmp = ClsLawReadRstMsg(intQ, strSql)
   If intQ = 1 Then
         MSHFlexGrid1.FixedCols = 0
         Set MSHFlexGrid1.Recordset = rsTmp
         Call SetGrd
   End If
   
   Set rsTmp = Nothing
End Sub

'設定開始輸入的欄位
Private Sub SetInputEntry()
   textCP05.SetFocus
   textCP05_GotFocus
End Sub

'收文日
Private Sub textCP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP05) = False Then
      If CheckIsTaiwanDate(textCP05, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "收文日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05.SetFocus
         textCP05_GotFocus
      End If
   End If
End Sub

'本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      If CheckIsTaiwanDate(textCP06, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "本所期限的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)

         textCP06.SetFocus
         textCP06_GotFocus
      Else
         '若本所期限非工作天則直接調整至最近的工作天
         textCP06 = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
      End If
   End If
End Sub

'法定期限
Private Sub textCP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      If CheckIsTaiwanDate(textCP07, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "法定期限的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07.SetFocus
         textCP07_GotFocus
      End If
   End If
End Sub

'案件性質
Private Sub textCP10_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   textCP10_2 = Empty
   Cancel = False
   If IsEmptyText(textCP10) = False Then
      '取得國內的案件性質名稱
      textCP10_2 = GetCaseTypeName(m_HC01, textCP10, 0)
      If IsEmptyText(textCP10_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP10.SetFocus
         textCP10_GotFocus
      End If
      If textCP10 = "0" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "不可輸入＜顧問聘任＞"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP10.SetFocus
         textCP10_GotFocus
      End If
   End If
End Sub

Private Sub textCP113_GotFocus()
  TextInverse textCP113
End Sub

Private Sub textCP113_Validate(Cancel As Boolean)
   If textCP113.Text <> "" Then
      If Trim(Val(textCP113.Text)) <> textCP113.Text Then
          MsgBox "請輸入數字！", vbOKOnly, "資料檢核"
          Cancel = True
          textCP113.SetFocus
          textCP113_GotFocus
      End If
   End If
   
End Sub

Private Sub textCP14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP21_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'是否取締案
Private Sub textCP21_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP21) = False Then
      Select Case textCP21
         Case " ", "Y":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP21.SetFocus
            textCP21_GotFocus
      End Select
   End If
End Sub

Private Sub textCP26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'是否算案件數
Private Sub textCP26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP26) = False Then
      Select Case textCP26
         Case " ", "N":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26.SetFocus
            textCP26_GotFocus
      End Select
   End If
End Sub

Private Sub textCP29_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'協辦人員
Private Sub textCP29_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCP29_2 = Empty
   If Not IsEmptyText(textCP29) Then
      textCP29_2 = GetStaffName(textCP29, False)
      If IsEmptyText(textCP29_2) Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "協辦人員代碼<" & textCP29 & ">不存在或未在職"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP29.SetFocus
         textCP29_GotFocus
      End If
   End If
End Sub

Private Sub textCP43_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註內容太長"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64.SetFocus
      textCP64_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub


'相關總收文號
Private Sub textCP43_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP43) = False Then
      If textCP43 = m_CP09 Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "相關總收文號不可為本身之收文號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP43.SetFocus
         textCP43_GotFocus
         GoTo EXITSUB
      End If
      
      strSql = "SELECT * FROM CaseProgress " & _
               "WHERE CP01 = '" & m_HC01 & "' AND " & _
                     "CP02 = '" & m_HC02 & "' AND " & _
                     "CP03 = '" & m_HC03 & "' AND " & _
                     "CP04 = '" & m_HC04 & "' AND " & _
                     "CP09 = '" & textCP43 & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
         rsTmp.Close
         Cancel = True
         strTit = "資料檢核"
         strMsg = "相關總收文號資料不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP43.SetFocus
         textCP43_GotFocus
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
EXITSUB:
   Set rsTmp = Nothing
End Sub

'承辦人
Private Sub textCP14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   
   Cancel = False
   textCP14_2 = Empty
   If IsEmptyText(textCP14) = False Then
      textCP14_2 = GetStaffName(textCP14)
      If IsEmptyText(textCP14_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "承辦人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP14.SetFocus
         textCP14_GotFocus
      End If
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   
   CheckDataValid = False
   
   If IsEmptyText(textCP05) = True Then
      strTit = "檢核資料"
      strMsg = "收文日不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP05.SetFocus
      GoTo EXITSUB
   End If
   
   If IsEmptyText(textCP10) = True Then
      strTit = "檢核資料"
      strMsg = "案件性質不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP10.SetFocus
      GoTo EXITSUB
   End If
   
   If IsEmptyText(Combo1) = True Then
      strTit = "檢核資料"
      strMsg = "業務區不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Combo1.SetFocus
      GoTo EXITSUB
   End If
   
   If IsEmptyText(Combo2) = True Then
      strTit = "檢核資料"
      strMsg = "諮詢人員不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Combo2.SetFocus
      GoTo EXITSUB
   End If
   
   If IsEmptyText(textCP14) = True And IsEmptyText(textCP29) = True Then
      strTit = "檢核資料"
      strMsg = "承辦人和協辦人員不可同時空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP14.SetFocus
      GoTo EXITSUB
   End If
    'Added by Lydia 2022/02/11 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
    End If

   CheckDataValid = True
EXITSUB:
End Function

Private Sub textCP05_GotFocus()
   TextInverse textCP05
End Sub

Private Sub textCP06_GotFocus()
   TextInverse textCP06
End Sub

Private Sub textCP07_GotFocus()
   TextInverse textCP07
End Sub

Private Sub textCP10_GotFocus()
   TextInverse textCP10
End Sub

Private Sub textCP14_GotFocus()
   TextInverse textCP14
End Sub

Private Sub textCP21_GotFocus()
   TextInverse textCP21
End Sub

Private Sub textCP26_GotFocus()
   TextInverse textCP26
End Sub

Private Sub textCP29_GotFocus()
   TextInverse textCP29
End Sub

Private Sub textCP43_GotFocus()
   TextInverse textCP43
End Sub

Private Sub textCP64_GotFocus()
   TextInverse textCP64
End Sub

'確認使用者所輸入的都完全正確
Private Function ValidateInput() As Boolean
   Dim Cancel As Boolean

   ValidateInput = False
   
   If textCP05.Enabled = True Then
      Cancel = False
      textCP05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP06.Enabled = True Then
      Cancel = False
      textCP06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP07.Enabled = True Then
      Cancel = False
      textCP07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP10.Enabled = True Then
      Cancel = False
      textCP10_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP14.Enabled = True Then
      Cancel = False
      textCP14_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If

   If textCP21.Enabled = True Then
      Cancel = False
      textCP21_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP26.Enabled = True Then
      Cancel = False
      textCP26_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP29.Enabled = True Then
      Cancel = False
      textCP29_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
      
   If textCP43.Enabled = True Then
      Cancel = False
      textCP43_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Combo2.Enabled = True Then
       If ChkExistStaff(Combo2.Text, strExc(4), strExc(5), strExc(6)) = True Then
            If Combo2.Text <> convForm(strExc(4), 6) & " " & strExc(5) Then
                Combo2.Text = convForm(strExc(4), 6) & " " & strExc(5)
                Combo2_Validate (False)
            End If
       Else
            Combo2.SetFocus
            Exit Function
       End If
   End If
   
   ValidateInput = True
End Function

'預設諮詢人員部門下拉選單
Private Sub SetCombo1()
   
   '除電腦中心及人事處外,其他人只能看到有在職員工的部門(王副總提需求江總同意)
   'Modified by Lydia 2020/04/29  排除F5X部門
   If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M21" Then
        strTmpQ = "SELECT A0901,A0902 FROM ACC090 WHERE A0904<>'Y' AND A0901 NOT LIKE 'F5%' "
   Else
        'Modified by Lydia 2024/11/08 DISTINCT ST03=> DISTINCT ST15
        strTmpQ = "SELECT A0901,A0902 FROM ACC090 WHERE A0904<>'Y' AND A0901 NOT LIKE 'F5%' " & _
                         "AND A0901<>'P29' AND A0901 IN (SELECT DISTINCT ST15 FROM STAFF WHERE ST04='1' AND ST01>'6' AND SUBSTR(ST01,1,1)<'G' AND SUBSTR(ST01,4,1)<>'9') "
   End If
   strTmpQ = strTmpQ & "ORDER BY A0901 DESC"
   
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
   If intQ = 1 Then
       Combo1.Clear
       With rsQuery
            .MoveFirst
            Do While Not .EOF
                If "" & .Fields("A0901") <> "" And "" & .Fields("A0902") <> "" Then
                    Combo1.AddItem "" & .Fields("A0901") & " " & .Fields("A0902"), 0
                End If
                .MoveNext
            Loop
       End With
       Combo1.AddItem String(10, " "), 0
       Combo1.ListIndex = 0
       Combo1.Tag = Combo1.Text
   End If
   Set rsQuery = Nothing
   
End Sub

'預設諮詢人員(CP13)下拉選單
Private Sub SetCombo2(ByVal pA0901 As String)
'pA0901：業務區下拉選單之選取
   
   'Modified by Lydia 2020/04/29 排除F開頭人員
   strTmpQ = "SELECT ST01, ST02 FROM STAFF WHERE ST04='1' AND ST01>'6' AND SUBSTR(ST01,1,1)<'G' AND SUBSTR(ST01,4,1)<>'9' AND ST01 NOT LIKE 'F%' "
   'Modified by Lydia 2020/05/14 改ST15;  ex.杜燕文
   'Memo by Lydia 2020/05/14 最初設計抓ST15,又因為諮詢人員非業務身份,在上線前改為ST03; 目前又改為ST15
   'If Trim(pA0901) <> "" Then strTmpQ = strTmpQ & "AND ST03='" & pA0901 & "' "
   If Trim(pA0901) <> "" Then strTmpQ = strTmpQ & "AND ST15='" & pA0901 & "' "
   strTmpQ = strTmpQ & "ORDER BY ST01 DESC "
   
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
   Combo2.Clear
   If intQ = 1 Then
       With rsQuery
            .MoveFirst
            Do While Not .EOF
                If "" & .Fields("ST01") <> "" And "" & .Fields("ST02") <> "" Then
                    Combo2.AddItem convForm("" & .Fields("ST01"), 6) & " " & .Fields("ST02"), 0
                End If
                .MoveNext
            Loop
       End With
       Combo2.AddItem String(10, " "), 0
       Combo2.ListIndex = 0
       Combo2.Tag = Combo2.Text
   End If
   Set rsQuery = Nothing
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'尋找下拉選單的Index
'Modified by Lydia 2022/02/11 ComboBox => Control
Private Function GetCmbVal(ByRef pCmb As Control, ByVal pKey As String, Optional ByRef pCmbTxt As String) As Integer
Dim intB As Integer

    pCmbTxt = ""
    GetCmbVal = 0
    For intB = 0 To pCmb.ListCount - 1
         If InStr(pCmb.List(intB), pKey) > 0 Then
             GetCmbVal = intB
             pCmbTxt = pCmb.List(intB)
             Exit Function
         End If
    Next intB
End Function

'抓本案(LA-999999)累計次數及時數
Private Function GetTimes_LA999999(ByVal pKind As String, ByRef pTimes As String, ByRef pHours As String, Optional ByRef pCP09 As String, Optional ByRef pCP53 As String, Optional ByRef pCP54 As String) As Boolean
'pKind: 1-本案的累計;  2-本次聘任期間的累計
'pTimes, pHours: 累計次數及時數
Dim strA As String, intA As Integer
Dim rsAD As New ADODB.Recordset

     GetTimes_LA999999 = False
     pTimes = "0"
     pHours = "0"
     
     If pKind = "2" Then '本次聘任期間
        If pCP09 = "" Then
            '未指定=>抓未取消收文之顧問聘任0且收文日最大筆的cp53及cp54, 若無此區間不可輸入新資料
            strA = "SELECT MAX(CP09) CP09,CP53,CP54 FROM CASEPROGRESS " & _
                      "WHERE CP01='LA' AND CP02='999999' AND CP03='0' AND CP04='00' AND CP09<'C' AND CP10='0' AND CP159=0 AND CP53||CP54=( " & _
                      "SELECT MAX(CP53||CP54) MAXR1 FROM CASEPROGRESS WHERE CP01='LA' AND CP02='999999' AND CP03='0' AND CP04='00' AND CP09<'C' AND CP10='0' AND CP159=0 " & _
                      ") AND CP53>0 AND CP54>0 GROUP BY CP53,CP54 "
        Else
            strA = "SELECT CP09, CP53, CP54 FROM CASEPROGRESS WHERE CP01='LA' AND CP02='999999' AND CP03='0' AND CP04='00' AND CP09<'C' AND CP10='0' AND CP159=0 " & _
                      "AND CP09 ='" & pCP09 & "' AND CP53>0 AND CP54>0 "
        End If
        intA = 1
        Set rsAD = ClsLawReadRstMsg(intA, strA)
        If intA = 0 Then
             MsgBox "本案無可用之顧問聘任區間！", vbExclamation + vbOKOnly, "LA-999999累計次數及時數"
             GoTo EXITSUB
        Else
             pCP09 = "" & rsAD.Fields("CP09")
             pCP53 = "" & rsAD.Fields("CP53")
             pCP54 = "" & rsAD.Fields("CP54")
        End If
     End If
     
     '本案累計次數及時數為所有進度(除案件性質顧問聘任0以外)未取消收文且未收費CP16=0的次數及總工作時數
     strA = "SELECT COUNT(CP09) C1,NVL(SUM(CP113),0) C2 FROM CASEPROGRESS WHERE CP01='LA' AND CP02='999999' AND CP03='0' AND CP04='00' AND CP09<'C' AND CP10<>'0'"
     strA = strA & " AND CP159=0 AND NVL(CP16,0)=0 "
     If pKind = "2" Then
         '本次期間累計之次數及工作時數，只抓收文日在本次聘任期間
         strA = strA & "AND CP05 BETWEEN " & pCP53 & " AND " & pCP54
     End If
     intA = 1
     Set rsAD = ClsLawReadRstMsg(intA, strA)
     If intA = 1 Then
         pTimes = "" & rsAD.Fields("C1")
         pHours = "" & rsAD.Fields("C2")
     End If
     
     GetTimes_LA999999 = True
     
EXITSUB:
     Set rsAD = Nothing
End Function

Private Function ChkExistStaff(ByVal pKeyStr As String, Optional ByRef pST01 As String, Optional ByRef pST02 As String, Optional ByRef pST03 As String) As Boolean
'pKeyStr: 傳入字串(員工編號、名稱、員工編號+名稱)
'pST01~pST03 : 回傳員工編號、名稱和所屬部門
Dim strB1 As String, intB As Integer
Dim rsBD As New ADODB.Recordset

    ChkExistStaff = False
    pST01 = "": pST02 = "": pST03 = ""
    
    'Modified by Lydia 2020/05/14 + ST15
    strB1 = "SELECT ST01, ST02, ST03, ST15 FROM STAFF WHERE ST04='1' AND ST01>'6' AND SUBSTR(ST01,1,1)<'G' AND SUBSTR(ST01,4,1)<>'9' " & _
                     "AND INSTR(RPAD(ST01,6,' ')||' '||ST02,'" & Trim(UCase(pKeyStr)) & "')>0 "
    strB1 = strB1 & "ORDER BY ST01 "
    intB = 1
    Set rsBD = ClsLawReadRstMsg(intB, strB1)
    If intB = 1 Then
        ChkExistStaff = True
        pST01 = "" & rsBD.Fields("ST01")
        pST02 = "" & rsBD.Fields("ST02")
        'Modified by Lydia 2020/05/14 改ST15; ex.杜燕文
        'pST03 = "" & rsBD.Fields("ST03")
        pST03 = "" & rsBD.Fields("ST15")
    Else
        MsgBox "查無此人員！", vbCritical
    End If
    
    Set rsBD = Nothing

End Function

'Added by Lydia 2020/05/27 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   If IsNull(rsSrcTmp.Fields("CP65")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP65")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("CP65"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CP66")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP66")) = False Then
         strTemp = ChangeWStringToTString(rsSrcTmp.Fields("CP66"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CP67")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP67")) = False Then
         strTemp = rsSrcTmp.Fields("CP67")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CP68")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP68")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("CP68"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CP69")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP69")) = False Then
         strTemp = ChangeWStringToTString(rsSrcTmp.Fields("CP69"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CP70")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP70")) = False Then
         strTemp = rsSrcTmp.Fields("CP70")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

'Added by Lydia 2020/05/27 控制畫面欄位是否可修改
Private Sub TxtLocked(ByVal bLocked As Boolean)

    textCP05.Locked = bLocked
    textCP06.Locked = bLocked
    textCP07.Locked = bLocked
    textCP10.Locked = bLocked
    Combo1.Locked = bLocked
    Combo2.Locked = bLocked
    textCP14.Locked = bLocked
    textCP21.Locked = bLocked
    textCP26.Locked = bLocked
    textCP29.Locked = bLocked
    textCP43.Locked = bLocked
    textCP64.Locked = bLocked
    textCP113.Locked = bLocked
End Sub
