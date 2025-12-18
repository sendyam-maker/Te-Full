VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040159 
   BorderStyle     =   1  '單線固定
   Caption         =   "申請人/代理人/案件各項指示資料維護"
   ClientHeight    =   6370
   ClientLeft      =   420
   ClientTop       =   4420
   ClientWidth     =   9160
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6370
   ScaleWidth      =   9160
   Begin VB.CommandButton cmdConfirm 
      BackColor       =   &H00FFFFC0&
      Caption         =   "完成確認"
      Height          =   330
      Left            =   7800
      Style           =   1  '圖片外觀
      TabIndex        =   39
      Top             =   120
      Width           =   1155
   End
   Begin VB.CommandButton CmdHelp 
      BackColor       =   &H008080FF&
      Caption         =   "？"
      Height          =   300
      Left            =   4050
      Style           =   1  '圖片外觀
      TabIndex        =   37
      Top             =   4140
      Width           =   260
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "T商標"
      Height          =   255
      Index           =   1
      Left            =   6810
      TabIndex        =   35
      Top             =   945
      Value           =   1  '核取
      Width           =   855
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "P專利"
      Height          =   255
      Index           =   0
      Left            =   5910
      TabIndex        =   34
      Top             =   945
      Value           =   1  '核取
      Width           =   855
   End
   Begin VB.CommandButton CmdWord 
      BackColor       =   &H00C0FFFF&
      Caption         =   "清單Word"
      Height          =   330
      Left            =   7800
      Style           =   1  '圖片外觀
      TabIndex        =   31
      Top             =   900
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   8400
      ScaleHeight     =   460
      ScaleWidth      =   650
      TabIndex        =   30
      Top             =   810
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.CommandButton cmdDgrid 
      Caption         =   "加入"
      Height          =   285
      Index           =   2
      Left            =   7455
      TabIndex        =   14
      Top             =   4155
      Width           =   735
   End
   Begin VB.CommandButton cmdDgrid 
      Caption         =   "刪除"
      Height          =   285
      Index           =   3
      Left            =   8220
      TabIndex        =   15
      Top             =   4155
      Width           =   735
   End
   Begin VB.CommandButton cmdDgrid 
      Caption         =   "新增"
      Height          =   285
      Index           =   1
      Left            =   6690
      TabIndex        =   9
      Top             =   4155
      Width           =   735
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Left            =   810
      TabIndex        =   10
      Text            =   "Combo3"
      Top             =   4147
      Width           =   3135
   End
   Begin VB.TextBox tNo 
      Height          =   300
      Index           =   4
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "tNo"
      Top             =   1650
      Width           =   495
   End
   Begin VB.TextBox tNo 
      Height          =   300
      Index           =   3
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "tNo"
      Top             =   1650
      Width           =   375
   End
   Begin VB.TextBox tNo 
      Height          =   300
      Index           =   2
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   5
      Text            =   "tNo"
      Top             =   1650
      Width           =   735
   End
   Begin VB.TextBox tNo 
      Height          =   300
      Index           =   1
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "tNo"
      Top             =   1650
      Width           =   495
   End
   Begin VB.TextBox tNo 
      Height          =   300
      Index           =   0
      Left            =   2250
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "tNo"
      Top             =   1290
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   3
      Top             =   1650
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "申請人/代理人編號："
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   1313
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含失效舊指示"
      Height          =   255
      Left            =   3270
      TabIndex        =   24
      Top             =   930
      Width           =   1485
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8415
      Top             =   30
      _ExtentX        =   988
      _ExtentY        =   988
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040159.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040159.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040159.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040159.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040159.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040159.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040159.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040159.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040159.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040159.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040159.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9160
      _ExtentX        =   16157
      _ExtentY        =   917
      ButtonWidth     =   1076
      ButtonHeight    =   882
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSGrid1 
      Bindings        =   "frm12040159.frx":20F4
      Height          =   1635
      Left            =   120
      TabIndex        =   29
      Top             =   2355
      Width           =   8835
      _ExtentX        =   15575
      _ExtentY        =   2875
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   5
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "v|分類|有效|記錄日期|內容"
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
      _Band(0).Cols   =   5
   End
   Begin MSForms.TextBox txtData 
      Height          =   300
      Index           =   13
      Left            =   150
      TabIndex        =   42
      Top             =   5460
      Visible         =   0   'False
      Width           =   615
      VariousPropertyBits=   671105051
      Size            =   "1085;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   1425
      Index           =   6
      Left            =   840
      TabIndex        =   13
      Top             =   4860
      Width           =   7935
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "13996;2514"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   300
      Index           =   5
      Left            =   5370
      TabIndex        =   11
      Top             =   4147
      Width           =   375
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   300
      Index           =   4
      Left            =   1170
      TabIndex        =   12
      Top             =   4500
      Width           =   975
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1931;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   300
      Index           =   2
      Left            =   8160
      TabIndex        =   17
      Top             =   1290
      Visible         =   0   'False
      Width           =   975
      VariousPropertyBits=   671105051
      MaxLength       =   12
      Size            =   "1931;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   300
      Index           =   1
      Left            =   810
      TabIndex        =   23
      Top             =   915
      Width           =   375
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   300
      Index           =   0
      Left            =   7440
      TabIndex        =   16
      Top             =   810
      Visible         =   0   'False
      Width           =   975
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1931;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtData 
      Height          =   300
      Index           =   3
      Left            =   2280
      TabIndex        =   18
      Top             =   4020
      Visible         =   0   'False
      Width           =   375
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1931;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Left            =   4020
      TabIndex        =   8
      Top             =   1650
      Width           =   4815
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "8493;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   3240
      TabIndex        =   2
      Top             =   1290
      Width           =   4815
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "8493;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCUID 
      Height          =   255
      Left            =   2220
      TabIndex        =   41
      Top             =   4523
      Width           =   6420
      Size            =   "11324;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblConfirm 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "完成確認："
      Height          =   180
      Left            =   5010
      TabIndex        =   40
      Tag             =   "完成確認："
      Top             =   690
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "(N：失效) "
      Height          =   225
      Left            =   5790
      TabIndex        =   38
      Top             =   4185
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "P.S.欄位設定之指示不在此處顯示，請在清單Word中查詢。"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   210
      TabIndex        =   36
      Top             =   690
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "使用部門："
      Height          =   225
      Index           =   0
      Left            =   5010
      TabIndex        =   33
      Top             =   960
      Width           =   915
   End
   Begin VB.Label lblTitle 
      Caption         =   "lblTitle"
      Height          =   345
      Index           =   1
      Left            =   1320
      TabIndex        =   32
      Top             =   1995
      Width           =   7665
   End
   Begin VB.Label Label1 
      Caption         =   "分類："
      Height          =   220
      Index           =   2
      Left            =   210
      TabIndex        =   25
      Top             =   4187
      Width           =   585
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   3840
      Y1              =   1785
      Y2              =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "內容："
      Height          =   255
      Index           =   6
      Left            =   210
      TabIndex        =   28
      Top             =   4920
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "記錄日期："
      Height          =   220
      Index           =   5
      Left            =   210
      TabIndex        =   27
      Top             =   4540
      Width           =   945
   End
   Begin VB.Label lblCancel 
      Caption         =   "失效註記："
      Height          =   225
      Left            =   4410
      TabIndex        =   26
      Top             =   4185
      Width           =   945
   End
   Begin VB.Label lblTitle 
      Caption         =   "分類第一碼："
      Height          =   345
      Index           =   0
      Left            =   150
      TabIndex        =   22
      Top             =   2000
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "(1代理人 2申請人 3案件)"
      Height          =   255
      Index           =   4
      Left            =   1230
      TabIndex        =   21
      Top             =   945
      Width           =   2025
   End
   Begin VB.Label Label1 
      Caption         =   "屬性："
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   20
      Top             =   945
      Width           =   585
   End
End
Attribute VB_Name = "frm12040159"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/23 改成Form2.0 ;Combo1、Combo2、MSGrid1改字型=新細明體-ExtB、txtData(index)、lblCUID
'Create by Lydia 2016/11/10 申請人/代理人/案件各項指示資料維護
Option Explicit

Dim sType As String '操作模式: E編輯 Q查詢
Dim SNo As String '上一層的申請人/代理人/案件Key值
Dim mPrevForm As Form
Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bolUpd As Boolean 'Added by Lydia 2020/06/02 是否可修改(DB)
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
'Modified by Lydia 2021/09/23 As TextBox=>As Control
Dim oText As Control

Dim rsAssign As New ADODB.Recordset
Dim rsAssignOld As New ADODB.Recordset
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_iGrdEditMode As Integer 'Grid狀態： 1->按下新增按鈕
Dim intLastRow As Integer '點選列
Dim stFormName As String
Dim m_Status As String '代理人(fa69)/申請人(cu80)狀態
Dim m_ITS03List As String '分類種類
Dim m_PKey As String '資料索引值
Dim colCUID As Integer 'CreateID在grid的起始位置
Dim colPKey As Integer
Dim colITS01 As Integer
Dim colITS02 As Integer
Dim colORD1 As Integer

Dim strTitle(0 To 4) As String '清單Word輸出: 表首資料

Dim nORD1 As Integer '目前ORD1值 99:已存在於DB的記錄
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset
Dim intSpecRow As Integer '新增->加入後的記錄
Dim strKind As String '分類第1碼: 改用Table控制
Dim strUpdITS05 As String '另外執行的SQL語法: 設失效指示

Dim m_WordLeft As Long, m_WordTop As Long 'Word開啟位置
Dim bVisible As Boolean 'Word是否顯示
Dim m_SystemKind As String '使用者可用的系統別
Dim m_StrUserST03 As String '使用者的部門
Dim maxSeq As String 'Added by Lydia 2020/09/14 執行序號
Dim m_pub_QL05 As String 'Add By Sindy 2025/8/27 只記錄於此Form


Public Sub SetParent(ByVal iTyp As String, ByVal iNo As String, ByRef Oldfm As Form)
    sType = iTyp
    '區分屬性
    SNo = Pub_GetITS01Type(iNo) & iNo
    Set mPrevForm = Me
End Sub

Private Sub Check1_Click()
   '當有資料並且功能列非動作時，重抓資料
   If txtData(1) <> "" And txtData(2) <> "" And TBar1.Buttons(11).Enabled = False And TBar1.Buttons(12).Enabled = False Then
      If ShowRecord Then
      End If
   End If
End Sub

Private Sub cmdDgrid_Click(Index As Integer)
   
   Select Case Index
      Case 1 '新增
         If m_iGrdEditMode = 1 And txtData(3) & txtData(4) & txtData(5) <> "" Then
            If MsgBox("確定要清除欄位？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
         End If
         cmdDgrid(Index).Enabled = False
         '指定Grid的Row，改變顏色(還原底色)
         If intSpecRow > 0 Then
            Call SetSpecColor(intSpecRow, False)
         End If
         'end 2020/05/11
         ClearInData '清除明細
         txtData(5).Text = ""
         txtData(4).Text = strSrvDate(2)
         Combo3.Text = ""
         Combo3.Tag = ""
         Combo3.SetFocus
         m_iGrdEditMode = 1
         UpdateCUID 0
         cmdDgrid(Index).Enabled = True
         
      Case 2 '加入
         '指定Grid的Row，改變顏色(還原底色)
         If intSpecRow > 0 Then
            Call SetSpecColor(intSpecRow, False)
         End If

         If ChkTxtValAdd = True Then
            cmdDgrid(Index).Enabled = False
            UpdateITS "U"
            ClearInData '清除明細
            cmdDgrid(1).SetFocus '移到新增
            cmdDgrid(Index).Enabled = True
         End If
            
      Case 3 '刪除
        If rsAssign.RecordCount > 0 And txtData(1) <> "" And txtData(2) <> "" And txtData(3) <> "" And txtData(4) <> "" And intLastRow > 0 Then
            If m_PKey <> Trim(MSGrid1.TextMatrix(intLastRow, colPKey)) Then
               MsgBox "請點選欲刪除的資料列！", vbCritical + vbOKOnly
            Else
                '指定Grid的Row，改變顏色(還原底色)
                If intSpecRow > 0 Then
                   Call SetSpecColor(intSpecRow, False)
                End If
                cmdDgrid(Index).Enabled = False
                strExc(1) = ""
                With rsAssign
                   If .RecordCount > 0 Then
                      .MoveFirst
                      Do While Not .EOF
                         '檢查若同一分類有仍有效的指示 ; 排除要刪除的記錄
                         'Modified by Lydia 2022/11/25 +ITS13複製對象編號
                         If .Fields("ITS03") = Trim(txtData(3)) And "" & .Fields("ITS13") = Trim(txtData(13)) And DBDATE(.Fields("ITS04")) <> DBDATE(Trim(txtData(4))) Then
                            If "" & .Fields("ITS05") = "" Then
                                strExc(1) = ""
                            Else
                                strExc(1) = "N"
                            End If
                         End If
                         .MoveNext
                      Loop
                   End If
                End With
               If strExc(1) = "N" Then
                   MsgBox "此分類指示目前皆為失效指示", vbInformation + vbOKOnly
               
               '畫面沒有含失效指示，另外查DB; 因為在未確定(存檔)前，修改指示為失效或直接新增失效指示都會直接顯示在Grid=>Recordset，所以只要另外查DB記錄
               ElseIf Check1.Value = 0 Or Chk2(0).Value = 1 Or Chk2(1).Value = 1 Then
                    'Modified by Lydia 2022/11/25 +ITS13複製對象編號
                    strExc(0) = "select count(*) cnt1,sum(decode(its05,'N',1,0)) cnt2 from instructions where its01='" & txtData(1) & "' and its03='" & Trim(txtData(3)) & "' and its13='" & Trim(txtData(13)) & "' and its04<>'" & DBDATE(txtData(4)) & "' "
                    If txtData(1) = "3" Then '本所案號
                        strExc(0) = strExc(0) & " and its02='" & tNo(1) & tNo(2) & tNo(3) & tNo(4) & "' "
                    Else  '代理人/申請人
                        strExc(0) = strExc(0) & " and its02='" & tNo(0) & "' "
                    End If
                    intQ = 1
                    Set rsQuery = ClsLawReadRstMsg(intQ, strExc(0))
                    If intQ = 1 Then
                        If Val("" & rsQuery.Fields("cnt1")) > 0 And Val("" & rsQuery.Fields("cnt1")) = Val("" & rsQuery.Fields("cnt2")) Then
                             MsgBox "此分類指示目前皆為失效指示", vbInformation + vbOKOnly
                        End If
                    End If
               End If
               UpdateITS "D"
               ClearInData '清除明細
               cmdDgrid(Index).Enabled = True
            End If
        Else
            MsgBox "無資料可刪除！", vbCritical + vbOKOnly
        End If
   End Select
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo3_LostFocus()
  
  If m_iGrdEditMode = 1 Then
       txtData(3).Text = Trim(Mid(Combo3.Text, 1, 3))
  Else '已建立的資料分類不可修改
     If txtData(3).Tag <> "" And Combo3.Tag <> "" Then
        Combo3.ListIndex = Combo3.Tag
     Else
        If cmdDgrid(2).Enabled = True Then
           txtData(3).Text = Trim(Mid(Combo3.Text, 1, 3))
        Else
           Combo3.Text = ""
        End If
     End If
  End If
  
End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
    '增加輸入後,自動檢查
    If txtData(3).Tag <> "" And Combo3.Tag <> "" Then  '已建立的資料分類不可修改
       Combo3.ListIndex = Combo3.Tag
    Else
       If cmdDgrid(2).Enabled = True Then
           If txtData(3).Tag = "" And Combo3.Text <> Combo3.List(Val(Combo3.Tag)) Then
                For intQ = 0 To Combo3.ListCount - 1
                     If InStr(Combo3.List(intQ), Combo3.Text) > 0 Then
                         Combo3.ListIndex = intQ
                         Exit Sub
                     End If
                Next intQ
           End If
       End If
    End If
End Sub

Private Sub MSGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow MSGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MSGrid1.col = nCol
   MSGrid1.row = nRow

   If Me.MSGrid1.row < 1 And Me.MSGrid1.Text <> "V" Then
      If Me.MSGrid1.Text = "記錄日期" Then
         If m_blnColOrderAsc = True Then
            Me.MSGrid1.Sort = 3  '數值昇冪
            
            m_blnColOrderAsc = False
         Else
            Me.MSGrid1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.MSGrid1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.MSGrid1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
   
End Sub

Private Sub MSGrid1_Click()
   GridClick MSGrid1, intLastRow, 0
   
   If intLastRow > 0 Then
      ReadGrid
   End If
   
   '指定Grid的Row，改變顏色(還原底色)
   If intSpecRow > 0 Then
      Call SetSpecColor(intSpecRow, False)
   End If
End Sub

Private Sub SetCombo3()

   Combo3.Clear
   m_ITS03List = ""
   
   intQ = 1
   '排除有基本檔對應欄位IT11
   strExc(0) = "select IT01||IT02 NO,IT03 from insttype where IT11 is null order by 1 "
   Set rsQuery = ClsLawReadRstMsg(intQ, strExc(0))
   If intQ = 1 Then
      With rsQuery
        .MoveFirst
        Do While Not .EOF
           '不限長度
           'Combo3.AddItem .Fields("NO") & " " & Trim(convForm("" & .Fields("IT03"), 20))
           Combo3.AddItem .Fields("NO") & " " & Trim("" & .Fields("IT03"))
           m_ITS03List = m_ITS03List & .Fields("NO") & ","
           .MoveNext
        Loop
      End With
   End If
   Combo3.Text = ""
End Sub

Private Sub GetCombo3(ByVal TS01 As String)
Dim jj As Integer

   If TS01 <> "" Then
      For jj = 0 To Combo3.ListCount - 1
         If Mid(Combo3.List(jj), 1, 3) = TS01 Then
            Combo3.ListIndex = jj
            Exit For
         End If
      Next
      
   End If
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction("frm12040159", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040159", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040159", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040159", strFind, False)
   
   If sType = "" And SNo = "" Then sType = "E" '非共同查詢或基本檔維護呼叫=>直接維護
   'Add By Sindy 2025/9/4
   If sType = "E" Then
      ClearQueryLog (Me.Name) '清除查詢印表記錄檔欄位
      pub_QL04 = ""
   End If
   '2025/9/4 END
   
   '共同查詢或基本檔維護呼叫=>不可查詢
    If SNo <> "" Then
       m_bQuery = False
    Else
       m_bQuery = True
    End If
   
   If sType = "E" And m_bUpdate = True Then
      m_bUpdate = True
   Else
      m_bUpdate = False
      cmdDgrid(1).Visible = False
      cmdDgrid(2).Visible = False
      cmdDgrid(3).Visible = False
      '共同查詢：不顯示失效註記的欄位
      lblCancel.Caption = "N：失效"
      lblCancel.ForeColor = vbRed
      Label3.Visible = False
      txtData(5).Visible = False
   End If
   m_bolUpd = m_bUpdate '是否可修改(DB)
   
   Call SetCombo3
   MoveFormToCenter Me
   '分類第1碼: 改用Table控制
   strKind = PUB_GetInType("2")
   lblTitle(1).Caption = strKind

   stFormName = Me.Caption

   '取得使用者可用的系統別
   m_SystemKind = "," & GetSystemKindByNick & ","
      
   '依使用者部預設使用部門
   If Pub_StrUserSt03 = "M51" Then
       strExc(2) = Left(UCase(InputBox("請輸入欲操作的部門代號？" & vbCrLf & "(F2:外專  P1:內專  F1:外商  P2:內商)", "預設使用部門", Pub_StrUserSt03)), 2)
       If strExc(2) <> "" And strExc(2) <> "M51" Then
          If InStr("F1,F2,P1,P2", Left(strExc(2), 2)) = 0 Then
              m_StrUserST03 = Pub_StrUserSt03
          Else
              m_StrUserST03 = strExc(2)
          End If
       Else
           m_StrUserST03 = Pub_StrUserSt03
       End If
   Else
       m_StrUserST03 = Pub_StrUserSt03
   End If
   If InStr("P1,F2", Left(m_StrUserST03, 2)) > 0 Then '內/外專
       Chk2(1).Value = False
   ElseIf InStr("P2,F1", Left(m_StrUserST03, 2)) > 0 Then '內/外商
       Chk2(0).Value = False
   End If
   
    If SNo <> "" Then
      SetCtrlReadOnly True
      QueryData
      UpdateToolbarState
      '拿掉"維護"
      If m_bUpdate = False Then
          Me.Caption = "申請人/代理人/案件各項指示查詢"
      End If
    Else
      OnAction vbKeyF4
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 修改
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 第一筆, 上一筆, 下一筆, 最後一筆
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
         
      Case vbKeyEscape:
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            If m_EditMode <> 0 Then
               OnAction vbKeyF10
            Else
               OnAction KeyCode
            End If
         End If
         
      Case vbKeyReturn
         If m_EditMode <> 0 Then
            '內容欄位可允許Enter(換行)
            If Me.ActiveControl <> txtData(6) Then OnAction vbKeyF9
         Else
            KeyCode = 0 '取消動作
         End If
         
      Case vbKeyInsert
         If cmdDgrid(2).Enabled = True Then
            cmdDgrid_Click 2
         End If
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsAssign = Nothing
   Set rsAssignOld = Nothing
   Set rsQuery = Nothing
   
   Set frm12040159 = Nothing
   
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      ' 修改
      Case 2: OnAction vbKeyF3
      ' 查詢
      Case 4: OnAction vbKeyF4
      ' 第一筆
      Case 6: OnAction vbKeyHome
      ' 前一筆
      Case 7: OnAction vbKeyPageUp
      ' 後一筆
      Case 8: OnAction vbKeyPageDown
      ' 最後一筆
      Case 9: OnAction vbKeyEnd
      ' 確定
      Case 11: OnAction vbKeyF9
      ' 取消
      Case 12: OnAction vbKeyF10
      ' 離開
      Case 14: OnAction vbKeyEscape
   End Select
End Sub

'清除資料：含最上方X,Y,個案
Private Sub ClearMData()
   For Each oText In txtData
      oText.Text = Empty
      oText.Tag = Empty
   Next
   
   For Each oText In tNo
      oText.Text = Empty
   Next
   
   Option1(0).Value = False: Option1(1).Value = False
   Combo1.Clear
   Combo2.Clear
   m_Status = ""
   
   m_iGrdEditMode = 0
   strUpdITS05 = ""
   
End Sub

'清除明細：最下方的輸入欄位
Private Sub ClearInData()
   For Each oText In txtData
      If oText.Index > 2 Then
         oText.Text = Empty
         oText.Tag = Empty
      End If
   Next
   If txtData(6).Locked = False Then
       txtData(4).Locked = False
   End If
   
   Combo3.Text = ""
   Combo3.Tag = ""
   m_PKey = ""
   m_iGrdEditMode = 0
   lblCUID = ""
   If lblCancel.Caption = "N：失效" Then
       lblCancel.Visible = False
   End If
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF3 ' 修改
         If txtData(1) = "" Or txtData(2) = "" Then
            MsgBox "查無資料！", vbCritical + vbOKOnly, "檢核資料"
            Exit Sub
         Else
            If tNo(0) <> "" And Combo1.ListCount = 0 Then
               MsgBox "申請人/代理人編號查無資料！", vbCritical + vbOKOnly, "檢核資料"
               Exit Sub
            End If
            If tNo(1) <> "" And Combo2.ListCount = 0 Then
               MsgBox "本所案號查無資料！", vbCritical + vbOKOnly, "檢核資料"
               Exit Sub
            End If
         End If
         
         m_EditMode = 2
         Call ProcHelp(False)

         SetCtrlReadOnly False
         UpdateToolbarState
         cmdDgrid(1).SetFocus '移到新增
      Case vbKeyF4 ' 查詢
         m_EditMode = 4
         Call ProcHelp(False)
         
         SetCtrlReadOnly True
         ClearMData  '清除資料：含X,Y,個案
         QueryData  '重新查詢：為了清除grid
         UpdateToolbarState
         
      Case vbKeyHome ' 第一筆
         ShowRecord -2
      Case vbKeyPageUp ' 前一筆
         ShowRecord -1
      Case vbKeyPageDown ' 後一筆
         ShowRecord 1
      Case vbKeyEnd ' 最後一筆
         ShowRecord 2
      Case vbKeyF9 ' 確定
         If OnWork = True Then
            UpdateToolbarState
            Call ProcHelp(True)
         Else
            Exit Sub
         End If
         SetCtrlReadOnly True
         
      Case vbKeyF10 ' 取消
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎？", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  m_EditMode = 0
                  Call ProcHelp(True)

                  ShowRecord
                  UpdateToolbarState
               End If
            Case Else
               If m_EditMode = 4 Or txtData(1) = "" Or txtData(2) = "" Then
                  ClearMData '清除資料：含X,Y,個案
               Else
                  ShowRecord
               End If
               m_EditMode = 0
               Call ProcHelp(True)

               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
         
      Case vbKeyEscape ' 離開
         If sType = "Q" Then
            tmpBol = fnCancelNowFormAndShowParentForm(Me)
         Else
            Unload Me
         End If
         
         Exit Sub
   End Select
   
   Select Case m_EditMode
      Case 1
         Me.Caption = stFormName & "(新增)"
      Case 2
         Me.Caption = stFormName & "(修改)"
      Case 4
         Me.Caption = stFormName & "(查詢)"
      Case Else
         Me.Caption = stFormName
   End Select
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)

    cmdDgrid(1).Enabled = Not bLocked
    cmdDgrid(2).Enabled = Not bLocked
    cmdDgrid(3).Enabled = Not bLocked
    
    For Each oText In txtData
       oText.Locked = bLocked
    Next
    
    '查詢:可變更屬性和申請人/代理人/案件
    If m_EditMode = 4 And sType = "E" And SNo = "" Then
       txtData(1).Locked = False
       For Each oText In tNo
          oText.Locked = False
       Next
       Option1(0).Enabled = True
       Option1(1).Enabled = True
    '修改:不可變更屬性和申請人/代理人/案件
    ElseIf m_EditMode = 2 And sType = "E" Then
       txtData(1).Locked = True
       For Each oText In tNo
          oText.Locked = True
       Next
       Option1(0).Enabled = False
       Option1(1).Enabled = False
    Else
       For Each oText In tNo
          oText.Locked = bLocked
       Next
       Option1(0).Enabled = Not bLocked
       Option1(1).Enabled = Not bLocked
    End If

End Sub

'依照權限設定其工具列的按紐狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      Case 0 ' 無任何動作
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate And txtData(1) <> "" And txtData(2) <> "" And sType = "E" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtData(1) <> "" And txtData(2) <> "" And sType = "E" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         '保留：不開放上一筆,下一筆...
'         If m_bQuery And txtData(0) <> "" And sType = "Q" Then
'            TBar1.Buttons(6).Enabled = True
'            TBar1.Buttons(7).Enabled = True
'            TBar1.Buttons(8).Enabled = True
'            TBar1.Buttons(9).Enabled = True
'         Else
'            TBar1.Buttons(6).Enabled = False
'            TBar1.Buttons(7).Enabled = False
'            TBar1.Buttons(8).Enabled = False
'            TBar1.Buttons(9).Enabled = False
'         End If
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
      
      Case 1, 2, 3, 4 '維護
         TBar1.Buttons(1).Enabled = False
         TBar1.Buttons(2).Enabled = False
         TBar1.Buttons(3).Enabled = False
         TBar1.Buttons(4).Enabled = False
         TBar1.Buttons(6).Enabled = False
         TBar1.Buttons(7).Enabled = False
         TBar1.Buttons(8).Enabled = False
         TBar1.Buttons(9).Enabled = False
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
         TBar1.Buttons(14).Enabled = False
   End Select
   
'不允許上下筆
TBar1.Buttons(6).Enabled = False
TBar1.Buttons(7).Enabled = False
TBar1.Buttons(8).Enabled = False
TBar1.Buttons(9).Enabled = False
'不開放自由查詢
'TBar1.Buttons(4).Enabled = False
End Sub

'檢查欄位：確定->修改、查詢按鈕用
Private Function CheckTxtVal() As Boolean
   
   Dim Cancel As Boolean, ii As Integer, jj As Integer

   '查詢
   If m_EditMode = 4 Then
      Cancel = False
      If Option1(0).Value = False And Option1(1).Value = False Then
         MsgBox "申請人/代理人編號或本所案號至少選擇一項！", vbCritical + vbOKOnly, "檢核資料"
         Cancel = True
         Exit Function
      Else
         If Option1(0).Value = True Then
           tNo_Validate 0, Cancel
           If Cancel = True Then
              Exit Function
           End If
           If Combo1.ListCount = 0 Then
              MsgBox "申請人/代理人編號查無資料！", vbCritical + vbOKOnly, "檢核資料"
              Exit Function
           End If
           tNo(0) = Mid(tNo(0) & "00", 1, 8)
           txtData(2) = tNo(0)
         Else
           If tNo(1) = "" Then
              Cancel = True
              MsgBox "請輸入系統別！", vbCritical + vbOKOnly, "檢核資料"
              tNo_GotFocus 1
              Exit Function
           ElseIf Len(tNo(2)) < 6 Then
              Cancel = True
              MsgBox "案號請至少輸入六碼！", vbCritical + vbOKOnly, "檢核資料"
              tNo_GotFocus 2
              Exit Function
           End If
           tNo_Validate 2, Cancel
           If Combo2.ListCount = 0 Then
              MsgBox "本所案號查無資料！", vbCritical + vbOKOnly, "檢核資料"
              Exit Function
           End If
           txtData(2) = tNo(1) & tNo(2) & tNo(3) & tNo(4)
         End If
         Call ReadITS02(txtData(2))
      End If
   End If
      
   If m_EditMode <> 0 And (m_Status <> "" And (InStr(m_Status, "不再使用") > 0 Or InStr(m_Status, "不得代理") > 0)) Then
      MsgBox "狀態欄為:" & m_Status, vbInformation + vbOKOnly, "檢核資料" 'Modified by Lydia 2024/02/2 vbCritical改成vbInformation
      'Exit Function  'Mark by Lydia 2024/02/02 改成彈提醒，但不限制；(X51332 --- from Bobbie)
   End If
   
   If m_iGrdEditMode > 0 Then
      If MsgBox("分類尚未加入,是否加入？", vbYesNo + vbInformation, "檢核資料") = vbYes Then
         Call cmdDgrid_Click(2)
      End If
   End If
   
   CheckTxtVal = True
   
End Function

Private Function ModRecord() As Boolean
Dim stSQL As String
Dim stDiff As String
Dim bolExist As Boolean
Dim jj As Integer
Dim tmpArr As Variant

On Error GoTo ErrHand

   cnnConnection.BeginTrans
   '有資料 => 無資料
   If rsAssign.RecordCount = 0 Then
      If rsAssignOld.RecordCount > 0 Then
         '刪除資料
         stSQL = "DELETE FROM INSTRUCTIONS WHERE ITS01='" & rsAssignOld.Fields("ITS01") & "' and ITS02='" & rsAssignOld.Fields("ITS02") & "' "
         Pub_SeekTbLog stSQL
         cnnConnection.Execute stSQL, intI
      End If
   Else
      '刪除資料(原來的資料在新的資料中找不到的)
      With rsAssignOld
      If .RecordCount > 0 Then
         .MoveFirst
         Do While Not .EOF
            rsAssign.MoveFirst
            bolExist = False
            Do While Not (rsAssign.EOF Or bolExist = True)
               If rsAssign.Fields("PKEY") = .Fields("PKEY") Then
                  bolExist = True
               End If
               rsAssign.MoveNext
            Loop
            
            If bolExist = False Then
                '刪除資料
                'Modified by Lydia 2022/11/25 +ITS13複製對象編號
                stSQL = "DELETE FROM INSTRUCTIONS WHERE ITS01='" & .Fields("ITS01") & "' and ITS02='" & .Fields("ITS02") & "' and ITS03='" & .Fields("ITS03") & "' and ITS13='" & .Fields("ITS13") & "' and ITS04='" & .Fields("ITS04") & "' "
                Pub_SeekTbLog stSQL
                cnnConnection.Execute stSQL, intI
            End If
            .MoveNext
         Loop
      End If
      End With
      
      '新增/變更資料
      With rsAssign
      .MoveFirst
      Do While Not .EOF
         If rsAssignOld.RecordCount = 0 Then
            bolExist = False
         Else
            rsAssignOld.MoveFirst
            bolExist = False
            Do While Not (rsAssignOld.EOF Or bolExist = True)
               stSQL = "": stDiff = ""
               If rsAssignOld.Fields("PKEY") = .Fields("PKEY") Then
                  bolExist = True
                  For jj = 4 To 6 '分類不可修改
                      If "" & rsAssignOld.Fields("ITS" & Format(jj, "00")) <> "" & rsAssign.Fields("ITS" & Format(jj, "00")) Then
                         stDiff = stDiff & ", ITS" & Format(jj, "00") & "=" & CNULL(rsAssign.Fields("ITS" & Format(jj, "00")))
                      End If
                  Next
               End If
               rsAssignOld.MoveNext
            Loop
         End If
         
         '新增資料
         If bolExist = False Then
            'Modified by Lydia 2022/11/25 +ITS13複製對象編號
            stSQL = "INSERT INTO INSTRUCTIONS (ITS01,ITS02,ITS03,ITS04,ITS05,ITS06,ITS07,ITS08,ITS09,ITS13) " & _
                     "VALUES ('" & .Fields("ITS01") & "','" & .Fields("ITS02") & "','" & .Fields("ITS03") & "','" & .Fields("ITS04") & "'," & CNULL(.Fields("ITS05")) & _
                     ",'" & "" & .Fields("ITS06") & "','" & .Fields("ITS07") & "','" & .Fields("ITS08") & "','" & .Fields("ITS09") & "','" & .Fields("ITS13") & "') "
                    
            Pub_SeekTbLog stSQL
            cnnConnection.Execute stSQL, intI
         '變更資料
         ElseIf stDiff <> "" Then
             'Modified by Lydia 2022/11/25 +ITS13複製對象編號
            stSQL = "UPDATE INSTRUCTIONS SET ITS10='" & strUserNum & "', ITS11='" & strSrvDate(1) & "', ITS12='" & IIf("" & .Fields("ITS12") = "", Left(Format(ServerTime, "000000"), 4), .Fields("ITS12")) & "'" & stDiff & _
                    " WHERE ITS01||ITS02||ITS03||ITS04||ITS13='" & .Fields("PKEY") & "' "
            Pub_SeekTbLog stSQL
            cnnConnection.Execute stSQL, intI
         End If
         
         .MoveNext
      Loop
      End With
   End If
   
   '另外執行的SQL語法;
   If strUpdITS05 <> "" Then  '設定失效指示
       tmpArr = Empty
       tmpArr = Split(strUpdITS05, ",")
       For jj = 0 To UBound(tmpArr)
           If Trim(tmpArr(jj)) <> "" Then
               'Modified by Lydia 2022/11/25 +ITS13複製對象編號
               stSQL = "select its01,its02,its03,its04,its13 from instructions where nvl(its05,'Y')<>'N' and ITS01||ITS02||ITS03||ITS04||ITS13='" & Trim(tmpArr(jj)) & "' "
               intQ = 1
               Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
               If intQ = 1 Then
                    'Modified by Lydia 2022/11/25 +ITS13複製對象編號
                    stSQL = "UPDATE INSTRUCTIONS SET ITS05='N',ITS10='" & strUserNum & "', ITS11='" & strSrvDate(1) & "', ITS12='" & Left(Format(ServerTime, "000000"), 4) & "'" & _
                                " WHERE ITS01||ITS02||ITS03||ITS04||ITS13='" & Trim(tmpArr(jj)) & "' "
                    Pub_SeekTbLog stSQL
                    cnnConnection.Execute stSQL, intI
               End If
           End If
       Next jj
   End If
   strUpdITS05 = ""
   
   'Added by Lydia 2020/08/25 增加-各項指示確認記錄檔
   'Modified by Lydia 2022/11/25 +ITS13複製對象編號
   stSQL = "select its02,count(its01||its02||its03||its04||its13) cno1,count (ic01||ic02) as cno2 From instructions, instconfirm " & _
             "where its01='" & txtData(1) & "' and its02='" & txtData(2) & "' and its01=ic01(+) and its02=ic02(+) group by its02 "
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 0 Then
       stSQL = "delete from instconfirm where ic01='" & txtData(1) & "' and ic02='" & txtData(2) & "' "
       cnnConnection.Execute stSQL, intI
   Else
       If Val("" & rsQuery.Fields("cno2")) = 0 Then
           stSQL = "insert into InstConfirm (IC01,IC02) values ('" & txtData(1) & "', '" & txtData(2) & "' ) "
           cnnConnection.Execute stSQL, intI
       End If
   End If
   'end 2020/08/25
   
   cnnConnection.CommitTrans
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Function OnWork() As Boolean
   Select Case m_EditMode
      Case 2: '確定->修改
         '重新檢查欄位有效性
         If CheckTxtVal() = True Then
            If ModRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
               ClearInData '清除明細
            End If
         End If
      
       Case 4: '查詢
         If CheckTxtVal() = True Then
            If txtData(2) = "" Then txtData(2) = CompITS02
            If ShowRecord = True Then
               OnWork = True
               m_EditMode = 0
               ClearInData '清除明細
            End If
         End If
         
   End Select

End Function

' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
Dim stITS01 As String
Dim stITS02 As String
Dim adoRst As New ADODB.Recordset
   
   stITS01 = txtData(1)
   If txtData(2) = "" Then
      stITS02 = CompITS02
   Else
      stITS02 = txtData(2)
   End If
      
   Select Case p_iWay
      Case 0 '當筆
            strExc(0) = "SELECT ITS01||ITS02 Sno FROM INSTRUCTIONS " & _
                        "WHERE  ITS01||ITS02='" & stITS01 & stITS02 & "' "
      Case -2 '首筆
            strExc(0) = "SELECT nvl(min(ITS01||ITS02),'') Sno FROM INSTRUCTIONS " & _
                        "WHERE ITS01||ITS02<'" & stITS01 & stITS02 & "' "
      Case -1 '前筆
            strExc(0) = "SELECT ITS01||ITS02 Sno FROM INSTRUCTIONS " & _
                        "WHERE  ITS01||ITS02<'" & stITS01 & stITS02 & "' order by 1 dsc "
      Case 1 '後筆
            strExc(0) = "SELECT ITS01||ITS02 Sno FROM INSTRUCTIONS " & _
                        "WHERE  ITS01||ITS02>'" & stITS01 & stITS02 & "' order by 1 asc "
      Case 2 '末筆
            strExc(0) = "SELECT nvl(max(ITS01||ITS02),'') Sno FROM INSTRUCTIONS " & _
                        "WHERE ITS01||ITS02>'" & stITS01 & stITS02 & "' "
   End Select
   intQ = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intQ, strExc(0))
   strExc(1) = ""
   If intQ = 1 Then strExc(1) = "" & adoRst.Fields("Sno")
   
   If strExc(1) <> "" Then
      QueryData strExc(1)
      ShowRecord = True
   Else
      If p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
      Else
         QueryData stITS01 & stITS02
         ShowRecord = True
      End If
   End If
   
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing

End Function

Private Sub QueryData(Optional ByVal pNo As String)
Dim strQ As String
Dim pCase(1 To 4) As String

On Error GoTo Checking
   
   If pNo <> "" Then     '查詢
      strQ = " AND ITS01='" & Mid(pNo, 1, 1) & "' AND ITS02='" & Mid(pNo, 2) & "' "
      pub_QL05 = ";屬性：" & Mid(pNo, 1, 1) & ";對象編號：" & Mid(pNo, 2) 'Add By Sindy 2025/8/27
   ElseIf SNo <> "" Then '上一層傳入
      strQ = " AND ITS01='" & Mid(SNo, 1, 1) & "' AND ITS02='" & IIf(InStr("1,2", Mid(SNo, 1, 1)) > 0, Mid(SNo, 2, 8), Mid(SNo, 2)) & "' "
      pub_QL05 = ";屬性：" & Mid(SNo, 1, 1) & ";對象編號：" & IIf(InStr("1,2", Mid(SNo, 1, 1)) > 0, Mid(SNo, 2, 8), Mid(SNo, 2)) 'Add By Sindy 2025/8/27
   Else          '清空grid
      strQ = " AND 0=1 "
      pub_QL05 = "" 'Add By Sindy 2025/8/27
   End If
   
   '限制使用者可操作的系統
   If sType = "E" And m_bolUpd = True Then '還原是否可修改的權限
      m_bUpdate = True
   Else
      m_bUpdate = False
   End If
   If sType = "E" And (pNo <> "" Or SNo <> "") Then
       If pNo <> "" And InStr("X,Y,", Mid(pNo, 2, 1)) = 0 Then
            Call ChgCaseNo(Mid(pNo, 2), pCase)
       ElseIf SNo <> "" And InStr("X,Y,", Mid(SNo, 2, 1)) = 0 Then
            Call ChgCaseNo(Mid(SNo, 2), pCase)
       End If
       If pCase(1) <> "" Then
         If InStr(m_SystemKind, "," & pCase(1) & ",") = 0 Then
            If Left(m_StrUserST03, 2) = "F2" Then  '外專人員：+FMP案
               If PUB_ChkIsFMP(pCase(1), pCase(2), pCase(3), pCase(4)) = False Then
                   MsgBox "無權限！", vbCritical, "檢核資料"
                   strQ = " and 0=1"
                   m_bUpdate = False
               End If
            Else
                   MsgBox "無權限！", vbCritical, "檢核資料"
                   strQ = " and 0=1"
                   m_bUpdate = False
            End If
         End If
       End If
   End If
   
   nORD1 = 99
   'Added by Lydia 2020/08/25 完成確認：預設
   cmdConfirm.Visible = False
   lblConfirm.Caption = lblConfirm.Tag
   
   If Check1.Value = 0 Then
      strQ = strQ & " AND NVL(ITS05,'Y') <> 'N'"
   'Modify By Sindy 2025/8/27
   Else
      pub_QL05 = pub_QL05 & ";含失效舊指示"
   End If
   '2025/8/27 END
   
   strExc(1) = ""
   If Chk2(0).Value = 1 Then strExc(1) = strExc(1) & " OR IT10='P'"
   If Chk2(1).Value = 1 Then strExc(1) = strExc(1) & " OR IT10='T'"
   'Add By Sindy 2025/8/27
   If Chk2(0).Value = 1 Or Chk2(1).Value = 1 Then
      pub_QL05 = pub_QL05 & ";使用部門："
      If Chk2(0).Value = 1 And Chk2(1).Value = 1 Then
         pub_QL05 = pub_QL05 & "P專利,T商標"
      ElseIf Chk2(0).Value = 1 Then
         pub_QL05 = pub_QL05 & "P專利"
      ElseIf Chk2(1).Value = 1 Then
         pub_QL05 = pub_QL05 & "T商標"
      End If
   End If
   '2025/8/27 END
   If strExc(1) <> "" Then strQ = strQ & " AND (" & Mid(strExc(1), 4) & " OR IT10 IS NULL) "
   
   'Added by Lydia 2020/09/14 原本使用Rdatafactory，因為欄位過長影響其他程式的Sort結果，另外開Table。
   maxSeq = "001" '執行序號=1, 視情況是否要改成遞增
   cnnConnection.Execute "delete from R12040159 where id = '" & strUserNum & "' and frname='" & Me.Name & "' and seqno <='" & maxSeq & "' "
   'Modified by Lydia 2022/11/25 +ITS13複製對象編號；ITS06R是GRID受限於長度並且置換換行符號，另外判斷複製指示加註〔111/xx/xx(複製系統日)複製來源：X or Y編號〕
'   strSql = "Insert into R12040159 (ID,FRNAME,SEQNO,PKEY,ORD1,FLAG1,ITS03,IT03,ITS05,ITS04D,ITS06R,ITS01,ITS02,ITS04,ITS06,ITS07,ITS08,ITS09,ITS10,ITS11,ITS12) "
'   strSql = strSql & " SELECT '" & strUserNum & "' as ID, '" & Me.Name & "' as FRNAME, '" & maxSeq & "' as SeqNO, ITS01||ITS02||ITS03||ITS04 PKEY,99 as ORD1," & _
'               " '' as Flag1,ITS03,IT03,ITS05,(ITS04 - 19110000) ITS04D,SUBSTR(REPLACE(ITS06,CHR(13)||CHR(10),' & '),1,500) ITS06R " & _
'               ",ITS01,ITS02,ITS04,ITS06,ITS07,ITS08,ITS09,ITS10,ITS11,ITS12 " & _
'               "FROM INSTRUCTIONS, INSTTYPE WHERE SUBSTR(ITS03,1,1)=IT01(+) AND SUBSTR(ITS03,2,2)=IT02(+) " & strQ
   strSql = "Insert into R12040159 (ID,FRNAME,SEQNO,PKEY,ORD1,FLAG1,ITS03,IT03,ITS05,ITS04D,ITS06R,ITS01,ITS02,ITS04,ITS06,ITS07,ITS08,ITS09,ITS10,ITS11,ITS12,ITS13) "
   strSql = strSql & " SELECT '" & strUserNum & "' as ID, '" & Me.Name & "' as FRNAME, '" & maxSeq & "' as SeqNO, ITS01||ITS02||ITS03||ITS04||ITS13 PKEY,99 as ORD1," & _
               " '' as Flag1,ITS03,IT03,ITS05,(ITS04 - 19110000) ITS04D,DECODE(ITS13,'0',NULL,'〔'||SQLDATET(ITS08)||' 複製來源：'||ITS13||' '||DECODE(FA01,NULL,NVL(CU05,NVL(CU04,CU06)),NVL(FA05,NVL(FA04,FA06)))||'〕')||SUBSTR(REPLACE(ITS06,CHR(13)||CHR(10),' & '),1,500) ITS06R " & _
               ",ITS01,ITS02,ITS04,ITS06,ITS07,ITS08,ITS09,ITS10,ITS11,ITS12,ITS13 " & _
               "FROM INSTRUCTIONS, INSTTYPE,FAGENT,CUSTOMER WHERE SUBSTR(ITS03,1,1)=IT01(+) AND SUBSTR(ITS03,2,2)=IT02(+) AND ITS13=FA01(+) AND '0'=FA02(+) AND ITS13=CU01(+) AND '0'=CU02(+) " & strQ
   cnnConnection.Execute strSql, intQ
   'end 2020/09/14

   'Modified by Lydia 2020/09/10 限ITS06R為100字
   'Modified by Lydia 2020/09/14 原本使用Rdatafactory，因為欄位過長影響其他程式的Sort結果，另外開Table。
   'strExc(0) = "SELECT '' v,ITS03,IT03,ITS05,(ITS04 - 19110000) ITS04D,SUBSTR(REPLACE(ITS06,CHR(13)||CHR(10),' & '),1,100) ITS06R " & _
               ",ITS01,ITS02,ITS04,ITS06,ITS07,ITS08,ITS09,ITS10,ITS11,ITS12,ITS01||ITS02||ITS03||ITS04 PKEY,99 as ORD1 " & _
               "FROM INSTRUCTIONS, INSTTYPE WHERE SUBSTR(ITS03,1,1)=IT01(+) AND SUBSTR(ITS03,2,2)=IT02(+) " & strQ
   'Modified by Lydia 2022/11/25 +ITS13複製對象編號
   strExc(0) = "SELECT FLAG1,ITS03,IT03,ITS05,ITS04D,ITS06R,ITS01,ITS02,ITS04,ITS06,ITS07,ITS08,ITS09,ITS10,ITS11,ITS12,ITS13,PKEY,ORD1 " & _
                 "FROM R12040159 where id='" & strUserNum & "' and frname ='" & Me.Name & "' and seqno='" & maxSeq & "' "
   '分類+記錄日期desc
   strExc(0) = strExc(0) & " ORDER BY ITS03 asc,ITS04 desc"
   
   'Modified by Lydia 2020/09/14 原本使用Rdatafactory，因為欄位過長影響其他程式的Sort結果，另外開Table。
   'intQ = 1
   'Set rsQuery = ClsLawReadRstMsg(intQ, strExc(0))
   '操作資料
   'Set rsAssign = PUB_CreateRecordset(rsQuery, , , , Me.Name)
   '保留原始資料
   'Set rsAssignOld = PUB_CreateRecordset(rsQuery, , , , Me.Name) '2020/09/14
   With rsAssign '操作資料
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
   End With
   With rsAssignOld '保留原始資料
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
   End With
   'end 2020/09/14
   
   Call SetGrd(True) '清空
   
   pub_QL05 = pub_QL05 & "(各項指示)" 'Add By Sindy 2025/8/27
   m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/27 記錄此Form的查詢條件
   If rsAssign.RecordCount > 0 Then
      If pub_QL04 <> "" Then InsertQueryLog (rsAssign.RecordCount) 'Add By Sindy 2025/8/27
      Call SetGrd
      
      txtData(1) = "" & rsAssign.Fields("ITS01")
      txtData(2) = "" & rsAssign.Fields("ITS02")
      'Added by Lydia 2020/08/25 完成確認：預設
      If sType = "E" And m_bUpdate = True And PUB_GetInstConfirm(m_StrUserST03, txtData(2).Text) = False Then
          cmdConfirm.Visible = True
      End If
      'Modified by Lydia 2022/11/25 改成公用模組
      'Call GetICstatus
       lblConfirm.Caption = lblConfirm.Tag & PUB_GetICstatus(txtData(2))
      'end 2020/08/25
       
      ReadGrid
   Else
      If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/27
      ClearMData '清除資料：含X,Y,個案
      
      If SNo <> "" Or pNo <> "" Then
         If SNo <> "" Then
            txtData(2) = IIf(InStr("1,2", Mid(SNo, 1, 1)) > 0, Mid(SNo, 2, 8), Mid(SNo, 2))
         Else
            txtData(2) = IIf(InStr("1,2", Mid(pNo, 1, 1)) > 0, Mid(pNo, 2, 8), Mid(pNo, 2))
         End If
      End If
      'Added by Lydia 2021/01/08 完成確認：開放無指示也可以確認
      If sType = "E" And m_bUpdate = True And txtData(2) <> "" Then
         If PUB_GetInstConfirm(m_StrUserST03, txtData(2).Text) = False Then
             cmdConfirm.Visible = True
         End If
      End If
      'Modified by Lydia 2022/11/25 改成公用模組
      'Call GetICstatus
      lblConfirm.Caption = lblConfirm.Tag & PUB_GetICstatus(txtData(2))
      'end 2021/01/08
   End If
   
   If txtData(2) <> "" Then ReadITS02 (txtData(2))
   
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
   
End Sub

Private Sub UpdateITS(ByVal iKind As String)
Dim m_Rows  As Integer
Dim bFind As Boolean

   If m_PKey <> "" Then m_Rows = 1
   
   With rsAssign
      '-------刪除
      If iKind = "D" Then
        If .RecordCount > 0 And m_Rows > 0 Then
           .MoveFirst
           '移動到要修改的資料
           Do While Not .EOF
              If .Fields("PKEY") = m_PKey Then
                 .Delete
                 Exit Do
              End If
              .MoveNext
           Loop
        End If
      '--------新增、修改
      ElseIf iKind = "U" Then
        If m_iGrdEditMode = 1 Then
          .AddNew
        Else
          If .RecordCount > 0 And m_Rows > 0 Then
             .MoveFirst
             '移動到要修改的資料
             Do While Not .EOF
                If .Fields("PKEY") = m_PKey Then
                   m_Rows = 0
                   bFind = True
                   Exit Do
                End If
                .MoveNext
             Loop
             If m_Rows > 0 Then .AddNew
          Else
             .AddNew
          End If
        End If
        
        .Fields("ITS01") = Trim(txtData(1).Text)
        .Fields("ITS02") = Trim(txtData(2).Text)
        .Fields("ITS03") = Trim(txtData(3).Text)
        .Fields("ITS04") = TransDate(txtData(4).Text, 2)
        .Fields("ITS05") = Trim(txtData(5).Text)
        .Fields("ITS06") = Trim(ChgSQL(txtData(6).Text))
        If Val("" & .Fields("ORD1")) < 99 Then '改用ORD1判斷
           .Fields("ITS07") = strUserNum   'CreateID, Date, Time
           .Fields("ITS08") = strSrvDate(1)
           .Fields("ITS09") = Left(Format(ServerTime, "000000"), 4)
        Else
           .Fields("ITS10") = strUserNum  'UpdateID, Date, Time
           .Fields("ITS11") = strSrvDate(1)
           .Fields("ITS12") = Left(Format(ServerTime, "000000"), 4)
        End If
        'Added by Lydia 2022/11/25 +ITS13複製對象編號
        If Trim(txtData(13).Text) = "" Then
           .Fields("ITS13") = "0"
        Else
           .Fields("ITS13") = Trim(txtData(13).Text)
        End If
        'end 2022/11/25
        
        If m_PKey = "" Then
           'Modified by Lydia 2022/11/25 +ITS13複製對象編號
           .Fields("PKEY") = Trim(txtData(1).Text) & Trim(txtData(2).Text) & Trim(txtData(3).Text) & TransDate(txtData(4).Text, 2) & Trim(txtData(13).Text)
        End If
        .Fields("IT03") = Trim(Mid(Combo3.Text, 5))
        .Fields("ITS04D") = Trim(txtData(4))
        'ITS06R 可能受限於100字元長度(無->有)會出錯程式錯誤
        .Fields("ITS06R") = PUB_StrToStr(Replace(Trim(ChgSQL(txtData(6).Text)), Chr(13) & Chr(10), " & "), 100)
        If bFind = False Then
           .Fields("ORD1") = nORD1 '目前ORD1值=>為了令新加入的記錄,都先放在Grid的最上方
           intSpecRow = nORD1 '指定Grid的Row
        End If
        .UPDATE
      End If
   End With
   
   '更新Grid
   Call SetGrd(True) '清空
   
   'Modified by Lydia 2020/05/11
   'Modified by by Lydia 2020/08/03 影響重置後的Grid抬頭重置,所以修改SetGrd每一次都重新設Recordset為Grid資料來源
'   If bFind = False Then
'       rsAssign.Sort = "ORD1 asc ,ITS03 asc,ITS04 desc "
'   Else
'       rsAssign.Sort = "ITS03 asc, ITS04 desc"
'   End If
'   'end 2020/08/03
   rsAssign.Sort = "ORD1 asc ,ITS03 asc,ITS04 desc "
   
   If rsAssign.RecordCount > 0 Then
       Call SetGrd
       '指定Grid的Row，改變顏色(反白，並且移動到該筆記錄)
       If iKind = "U" And intSpecRow > 0 Then
           Call SetSpecColor(intSpecRow, True)
       End If
   End If
   
   intLastRow = 0
End Sub

Private Sub ReadGrid()
   
   ClearInData '清除明細
   If MSGrid1.TextMatrix(intLastRow, 1) <> "" And intLastRow > 0 Then
      With MSGrid1
          txtData(1).Text = Trim(.TextMatrix(intLastRow, colITS01))
          txtData(2).Text = Trim(.TextMatrix(intLastRow, colITS02))
          txtData(3).Text = Trim(.TextMatrix(intLastRow, 1))
          txtData(3).Tag = txtData(3).Text
          Call GetCombo3(txtData(3).Text)  '分類
          Combo3.Tag = Combo3.ListIndex
          txtData(4) = Trim(.TextMatrix(intLastRow, 4))
          txtData(4).Tag = txtData(4) 'Added by Lydia 2022/11/25
          txtData(5) = Trim(.TextMatrix(intLastRow, 3))
          'txtData(6) = Trim(.TextMatrix(intLastRow, 9)) '有換行符號的ITS06 'Remove by Lydia 2020/09/15
          m_PKey = Trim(.TextMatrix(intLastRow, colPKey))
          'Added by Lydia 2022/11/25 +ITS13複製對象編號
          txtData(13) = Trim(.TextMatrix(intLastRow, 16))
      End With
   End If
   'Added by Lydia 2020/09/15 因為內容長度受限於MSHFlexGrid欄位的ColumnWidth+RowHeight，最多顯示999字元；所以另外抓內容
   If txtData(6) = "" And m_PKey <> "" Then
       With rsAssign
           .MoveFirst
           Do While Not .EOF
               If .Fields("PKEY") = m_PKey Then
                  txtData(6) = Trim(.Fields("ITS06"))
                  Exit Do
               End If
               .MoveNext
           Loop
       End With
   End If
   'end 2020/09/15
   
   '共同查詢模式
   If lblCancel.Caption = "N：失效" Then
       If txtData(5) = "N" Then
           lblCancel.Visible = True
       Else
           lblCancel.Visible = False
       End If
   End If

   m_iGrdEditMode = 0
   
   If txtData(2) <> "" Then ReadITS02 (txtData(2))
   
   UpdateCUID 1
End Sub

'檢查欄位：加入按鈕用
Private Function ChkTxtValAdd() As Boolean
Dim idx As Integer
Dim Cancel As Boolean

   For Each oText In txtData
      If oText.Index > 0 Then
         idx = oText.Index
         Cancel = False
         Txtdata_Validate idx, Cancel
         If Cancel = True Then
            Txtdata_GotFocus idx
            Exit Function
         End If
      End If
   Next
   
  'Added by Lydia 2021/09/23 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
  If PUB_ChkUniText(Me, , True, "TextBox") = False Then
     Exit Function
  End If

On Error GoTo ErrHandle

   If InStr(m_ITS03List, txtData(3)) = 0 Then
      MsgBox "分類錯誤 ！", vbCritical + vbOKOnly, "檢核資料"
      Exit Function
   End If
    
   '目前ORD1值=>為了令新加入的記錄,都先放在Grid的最上方
   If nORD1 < 1 Then
      MsgBox "加入次數過多，請先按""確定""進行存檔！", vbInformation
      Exit Function
   Else
      nORD1 = nORD1 - 1
   End If
   
   If Trim(txtData(13)) = "" Then txtData(13) = "0" 'Added by Lydia 2022/11/13 +ITS13複製對象編號

   strExc(3) = "" '先抓DB
   '檢查DB是否存在同一分類的指示; 因為失效指示有可能不顯示
   If m_PKey = "" Then
        'Modified by Lydia 2022/11/25 +ITS13複製對象編號
        strExc(0) = "SELECT ITS01,ITS02,ITS03,ITS04,(ITS04-19110000) ITS04D,ITS05,REPLACE(ITS06,CHR(13)||CHR(10),' & ') ITS06R,IT03,ITS13 " & _
                          "FROM INSTRUCTIONS,INSTTYPE WHERE ITS01='" & txtData(1) & "' AND ITS03='" & Trim(txtData(3)) & "' AND SUBSTR(ITS03,1,1)=IT01(+) AND SUBSTR(ITS03,2,2)=IT02(+) "
        If txtData(1) = "3" Then '本所案號
            strExc(0) = strExc(0) & " AND ITS02='" & tNo(1) & tNo(2) & tNo(3) & tNo(4) & "' "
        Else  '代理人/申請人
            strExc(0) = strExc(0) & " AND ITS02='" & tNo(0) & "' "
        End If
        strExc(0) = strExc(0) & " order by its04 desc "
        intQ = 1
        Set rsQuery = ClsLawReadRstMsg(intQ, strExc(0))
        If intQ = 1 Then
             rsQuery.MoveFirst
             Do While Not rsQuery.EOF
                'Modified by Lydia 2022/11/25 +ITS13複製對象編號
                If "" & rsQuery.Fields("ITS04") = DBDATE(Trim(txtData(4))) And "" & rsQuery.Fields("ITS13") = Trim(txtData(13)) Then
                    If "" & rsQuery.Fields("ITS05") = "N" And Check1.Value = 0 Then
                       MsgBox "同一記錄日期已有失效指示！", vbCritical
                    Else
                       MsgBox "同一分類的記錄日期不可相同！", vbCritical
                    End If
                    Exit Function
                ElseIf "" & rsQuery.Fields("ITS05") <> "N" Then
                   '清單內容=PKEY+"<@>" + 分類ITS03 + 分類說明IT03 + 記錄日期ITS04 + 內容ITS06R (以||區隔各筆資料)
                   'Modified by Lydia 2022/11/25 +ITS13複製對象編號
                   If InStr(strUpdITS05 & ",", rsQuery.Fields("ITS01") & rsQuery.Fields("ITS02") & rsQuery.Fields("ITS03") & rsQuery.Fields("ITS04") & rsQuery.Fields("ITS13")) = 0 Then  '排除已被設定失效
                        strExc(3) = strExc(3) & rsQuery.Fields("ITS01") & rsQuery.Fields("ITS02") & rsQuery.Fields("ITS03") & rsQuery.Fields("ITS04") & rsQuery.Fields("ITS13") & "<@>" & convForm("" & rsQuery.Fields("ITS03"), 4) & "　" & convForm("" & rsQuery.Fields("IT03"), 8) & _
                                            "　" & rsQuery.Fields("ITS04D") & "　" & Mid("" & rsQuery.Fields("ITS06R"), 1, 100) & "||"
                   End If
                End If
                rsQuery.MoveNext
             Loop
        End If
   End If
   'end 2020/05/06
   
    With rsAssign
       If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
             '因為同一分類有效指示可以有一筆以上,所以日期不可相同
             'Modified by Lydia 2022/11/25 +ITS13複製對象編號
             If .Fields("ITS03") = Trim(txtData(3)) And .Fields("ITS13") = Trim(txtData(13)) And .Fields("ITS04D") = Trim(txtData(4)) And .Fields("PKEY") <> m_PKey Then
                 MsgBox "同一分類的記錄日期不可相同！", vbCritical
                 Exit Function
             End If
             
             '檢查若同一分類有仍有效的指示
             If Trim(txtData(5)) <> "N" Then
                If .Fields("ITS03") = Trim(txtData(3)) And .Fields("PKEY") <> m_PKey And "" & .Fields("ITS05") <> "N" Then
                   '清單內容=PKey+"<@>" + 分類ITS03 + 分類說明IT03 + 記錄日期ITS04 + 內容ITS06R (以||區隔各筆資料)
                   If strExc(3) & strUpdITS05 = "" Or (strExc(3) & "," & strUpdITS05 <> "" And InStr(strExc(3) & "," & strUpdITS05, .Fields("PKEY")) = 0) Then  'Added by Lydia 2020/05/13 +判斷DB是否有抓到
                      strExc(3) = strExc(3) & .Fields("PKEY") & "<@>" & convForm("" & .Fields("ITS03"), 4) & "　" & convForm("" & .Fields("IT03"), 8) & "　" & convForm("" & .Fields("ITS04D"), 8) & "　" & Mid("" & .Fields("ITS06R"), 1, 100) & "||"
                   End If
                End If
             End If
             .MoveNext
          Loop

       End If
    End With
   
   '各項指示有效清單
   If strExc(3) <> "" Then
       frm880004.iStiu = 5
       frm880004.m_TempList = strExc(3)
       Set frm880004.mPreForm = Me
       Me.Enabled = False
       frm880004.Show vbModal
       Me.Enabled = True
       
       '有勾選項目,比對目前資料集做修改
       strExc(1) = Me.Tag
       If strExc(1) <> "" Then
          strUpdITS05 = strUpdITS05 & "," & strExc(1)
          With rsAssign
             .MoveFirst
             Do While Not .EOF
                If InStr(strExc(1), Trim("" & .Fields("PKEY"))) > 0 And "" & .Fields("PKEY") <> "" Then
                  .Fields("ITS05") = "N"
                End If
                .MoveNext
             Loop
          End With
       End If
       Me.Tag = ""
   End If
   
   ChkTxtValAdd = True

ErrHandle:
   If Err.Number <> 0 Then
       MsgBox Err.Description, vbCritical, "檢核加入資料"
   End If
End Function

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByVal actType As Integer)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If actType = 0 Then
      strCName = GetStaffName(strUserNum, True)
      strCDate = Format(strSrvDate(2), "###/##/##")
      strCTime = ""
   ElseIf intLastRow > 0 Then
      With MSGrid1
        If IsNull(.TextMatrix(intLastRow, colCUID)) = False Then
           If IsEmptyText(.TextMatrix(intLastRow, colCUID)) = False Then
              strCName = GetStaffName(.TextMatrix(intLastRow, colCUID), True)
           End If
        End If
        If IsNull(.TextMatrix(intLastRow, colCUID + 1)) = False Then
           If IsEmptyText(.TextMatrix(intLastRow, colCUID + 1)) = False Then
              strTemp = TAIWANDATE(.TextMatrix(intLastRow, colCUID + 1))
              strCDate = Format(strTemp, "###/##/##")
           End If
        End If
        If IsNull(.TextMatrix(intLastRow, colCUID + 2)) = False Then
           If IsEmptyText(.TextMatrix(intLastRow, colCUID + 2)) = False Then
              strTemp = .TextMatrix(intLastRow, colCUID + 2)
              strCTime = Format(strTemp, "00:00")
           End If
        End If
        If IsNull(.TextMatrix(intLastRow, colCUID + 3)) = False Then
           If IsEmptyText(.TextMatrix(intLastRow, colCUID + 3)) = False Then
              strUName = GetStaffName(.TextMatrix(intLastRow, colCUID + 3), True)
           End If
        End If
        If IsNull(.TextMatrix(intLastRow, colCUID + 4)) = False Then
           If IsEmptyText(.TextMatrix(intLastRow, colCUID + 4)) = False Then
              strTemp = TAIWANDATE(.TextMatrix(intLastRow, colCUID + 4))
              strUDate = Format(strTemp, "###/##/##")
           End If
        End If
        If IsNull(.TextMatrix(intLastRow, colCUID + 5)) = False Then
           If IsEmptyText(.TextMatrix(intLastRow, colCUID + 5)) = False Then
              strTemp = .TextMatrix(intLastRow, colCUID + 5)
              strUTime = Format(strTemp, "00:00")
           End If
        End If
      End With
   End If
   ' 設定CUID中的文字
   lblCUID.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & " " & vbTab & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Sub tNo_GotFocus(Index As Integer)
   TextInverse tNo(Index)
   
   If m_EditMode = 4 Then
      If Index = 0 Then
         Option1(0).Value = True
      ElseIf Index = 1 Then
         Option1(1).Value = True
      End If
   End If
End Sub

Private Sub tNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub tNo_Validate(Index As Integer, Cancel As Boolean)
   If m_EditMode = 4 Then
      Select Case Index
          Case 0
              If tNo(Index) <> "" Then
                 Option1(0).Value = True
                 If Left(tNo(0), 1) <> "X" And Left(tNo(0), 1) <> "Y" Then
                    MsgBox "申請人/代理人編號必須為X/Y開頭！", vbCritical + vbOKOnly, "檢核資料"
                    tNo(0).SetFocus
                    tNo_GotFocus 0
                    Cancel = True
                    Exit Sub
                 Else
                    txtData(1) = Pub_GetITS01Type(tNo(0))
                 End If
                 If Len(tNo(0)) < 6 Then
                    MsgBox "申請人/代理人編號請至少輸入六碼！", vbCritical + vbOKOnly, "檢核資料"
                    tNo(0).SetFocus
                    tNo_GotFocus 0
                    Cancel = True
                    Exit Sub
                 End If
                 tNo(0) = Mid(tNo(0) & "00", 1, 8)
                 txtData(2) = tNo(0)
                 Call ReadITS02(tNo(0), False)
              End If
          Case 1
              If tNo(Index) <> "" Then
                 Option1(1).Value = True
                 txtData(1) = Pub_GetITS01Type(tNo(Index))
              End If
          Case 2
              If tNo(Index) <> "" Then
                 Option1(1).Value = True
                 txtData(1) = Pub_GetITS01Type(tNo(1))
                 If tNo(1) = "" Then
                    MsgBox "請輸入系統別！", vbCritical + vbOKOnly, "檢核資料"
                    tNo(1).SetFocus
                    tNo_GotFocus 1
                    Cancel = True
                    Exit Sub
                 ElseIf Len(tNo(2)) < 6 Then
                    MsgBox "案號請至少輸入六碼！", vbCritical + vbOKOnly, "檢核資料"
                    tNo(Index).SetFocus
                    tNo_GotFocus Index
                    Cancel = True
                    Exit Sub
                 End If
                 tNo(3) = Mid(tNo(3).Text & "0", 1, 1)
                 tNo(4) = Mid(tNo(4).Text & "00", 1, 2)
                 txtData(2) = tNo(1) & tNo(2) & tNo(3) & tNo(4)
                 Call ReadITS02(txtData(2), False)
              End If
          Case 4
              If tNo(1) <> "" And tNo(2) <> "" And tNo(3) <> "" And tNo(4) <> "00" Then
                 txtData(2) = tNo(1) & tNo(2) & tNo(3) & tNo(4)
                 Call ReadITS02(txtData(2), False)
              End If
      End Select
   End If
   
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
   TextInverse txtData(Index)
   CloseIme
End Sub

'Modified by Lydia 2021/09/23 改成Form 2.0
'Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index <> 6 Then KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Lydia 2021/09/23 Form 2.0的TextBox增加右鍵選單功能
Private Sub txtData_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtData(Index)
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)

   Select Case Index
       Case 1
            'Modified by Lydia 2021/09/24 + 排除空白 => txtData(Index).Text <> ""
            If txtData(Index).Text <> "" And txtData(Index).Text <> "1" And txtData(Index).Text <> "2" And txtData(Index).Text <> "3" Then
               MsgBox "請輸入1-3！", vbCritical + vbOKOnly
               GoTo JumpErrorInput
            End If
       Case 2, 3
            If txtData(Index) = "" Then
                MsgBox IIf(Index = 2, "申請人/代理人編號或本所案號", "分類") & "不可空白！"
                Cancel = True
            End If
       Case 4 '記錄日期
            If m_EditMode = 1 Or m_EditMode = 2 Then
                If txtData(Index) = "" Then
                    MsgBox "記錄日期不可空白！"
                    GoTo JumpErrorInput
                Else
                   If ChkDate(txtData(Index)) Then
                      'Added by Lydia 2022/11/25 +ITS13複製對象編號
                      If Len(txtData(13)) > 1 And txtData(4).Tag <> txtData(4).Text Then
                         MsgBox "複製指示的記錄日期不可變更！"
                         GoTo JumpErrorInput
                      End If
                      'end 2022/11/25
                      
                      '關聯企業的複製指示之記錄日期為系統日期+1工作天
                      'Modified by Lydia 2020/09/02 以西元年月判斷
                      'If txtData(Index) > TransDate(CompWorkDay(2, strSrvDate(1)), 1) Then
                      'Modified by Lydia 2022/11/25 改為系統日
                      'If TransDate(txtData(Index), 2) > CompWorkDay(2, strSrvDate(1)) Then
                      '   MsgBox "記錄日期不可大於系統日+1工作天！"
                      If TransDate(txtData(Index), 2) > strSrvDate(1) Then
                         MsgBox "記錄日期不可大於系統日！"
                         GoTo JumpErrorInput
                      End If
                   Else
                      GoTo JumpErrorInput
                   End If
                End If
            End If
       Case 5 '失效註記
            If m_EditMode = 1 Or m_EditMode = 2 Then
              If Trim(txtData(Index).Text) <> "" And txtData(Index).Text <> "N" Then
                 MsgBox "失效註記請輸入N ！"
                 GoTo JumpErrorInput
              End If
            End If
       Case 6 '內容
            'Modified by Lydia 2022/12/26
            'If Not CheckLengthIsOK(txtData(Index), 2000) Then
            If Not CheckLengthIsOK(txtData(Index), 3000) Then
               GoTo JumpErrorInput
            End If
   End Select
   
   Exit Sub
   
JumpErrorInput:
Cancel = True
txtData(Index).SetFocus
Txtdata_GotFocus Index
End Sub

Private Function CompITS02() As String
   If Option1(0).Value = True And tNo(0) <> "" Then
      CompITS02 = Mid(tNo(0) & "00000000", 1, 8)
   End If
   If Option1(1).Value = True And tNo(1) <> "" And tNo(2) <> "" Then
      CompITS02 = Trim(tNo(1)) & Mid(tNo(2) & "000000", 1, 6) & Mid(tNo(3) & "0", 1, 1) & Mid(tNo(4) & "0", 1, 2)
   End If
End Function

Private Sub ReadITS02(ByVal TT01 As String, Optional ByVal bReset As Boolean = True)
Dim strA1 As String

   If TT01 = "" Then Exit Sub
   
   If bReset = True Then
      For Each oText In tNo
         oText.Text = Empty
      Next
      Option1(0).Value = False: Option1(1).Value = False
   End If
   
   Combo1.Clear: Combo2.Clear
   m_Status = ""
   
   Erase strTitle
   
   ''代理人(fa69)/申請人(cu80)狀態
   txtData(1) = Pub_GetITS01Type(TT01)
   Select Case Mid(TT01, 1, 1)
      Case "Y"  '代理人
           Option1(0).Value = True
           strA1 = "select fa01 No,fa04 name1,fa05||fa63||fa64||fa65 name2,fa06 name3,fa69 istatus,substr(na01,1,3) na01,na03" & _
                   " from fagent,nation where fa01 = '" & Mid(TT01 & "0000000", 1, 8) & "' and fa02='0' and fa10=na01(+)"
      Case "X"  '申請人
           Option1(0).Value = True
           strA1 = "select cu01 No,cu04 name1,cu05||cu88||cu89||cu90 name2,cu06 name3,cu80 istatus,substr(na01,1,3) na01,na03" & _
                   " from customer,nation where cu01 = '" & Mid(TT01 & "0000000", 1, 8) & "' and cu02='0' and cu10=na01(+) "
      Case Else '案件
           Option1(1).Value = True
           Call ChgCaseNo(TT01, strExc)
           If ClsPDCheckCaseCodeIsExist(strExc(1), strExc(2), strExc(3), strExc(4), strExc(5), strExc(6), strExc(7), , strExc(8)) Then
              tNo(1) = strExc(1)
              tNo(2) = strExc(2)
              tNo(3) = strExc(3)
              tNo(4) = strExc(4)
              Combo2.AddItem "中: " & strExc(5)
              Combo2.AddItem "英: " & strExc(6)
              'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
              Combo2.AddItem "外: " & strExc(7)
              Combo2.ListIndex = 0
              '清單表首
              strTitle(0) = strExc(1) & "-" & strExc(2) & "-" & strExc(3) & "-" & strExc(4)
              strTitle(1) = PUB_GetNationName(strExc(8))
              strTitle(2) = strExc(5)
              strTitle(3) = strExc(6)
              strTitle(4) = strExc(7)

           Else
              Exit Sub
           End If
   End Select
   
   If strA1 <> "" Then
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, strA1)
      If intQ = 1 Then
         Option1(0).Value = True
         With rsQuery
            m_Status = "" & .Fields("istatus")
            tNo(0).Text = "" & .Fields("No")
            If Left(m_StrUserST03, 1) = "F" Then
                Combo1.AddItem "英: " & .Fields("name2")
                Combo1.AddItem "中: " & .Fields("name1")
                Combo1.AddItem "日: " & .Fields("name3")
            Else
                Combo1.AddItem "中: " & .Fields("name1")
                Combo1.AddItem "英: " & .Fields("name2")
                Combo1.AddItem "日: " & .Fields("name3")
            End If
            Combo1.ListIndex = 0
            '清單表首
            strTitle(0) = "" & .Fields("No")
            strTitle(1) = "" & .Fields("na03")
            strTitle(2) = "" & .Fields("name1")
            strTitle(3) = "" & .Fields("name2")
            strTitle(4) = "" & .Fields("name3")

         End With
      End If
   End If
   
End Sub

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
 Dim arrGridHeadText, arrGridHeadWidth
 Dim iRow As Integer
      
   arrGridHeadText = Array("v", "分類", "分類說明", "失效", "記錄日期", "內　　容", "ITS01", "ITS02", "ITS04", "ITS06", "ITS07", "ITS08", "ITS09", "ITS10", "ITS11", "ITS12", "ITS13", "PKEY", "ORD1")
   arrGridHeadWidth = Array(200, 500, 1060, 500, 840, 6000, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
      
   MSGrid1.Visible = False
   MSGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then '重置Grid
         Set MSGrid1.Recordset = Nothing
         MSGrid1.Clear
         MSGrid1.Rows = 2
   Else
         Set MSGrid1.Recordset = rsAssign '設Recordset為Grid資料來源
   End If
   For iRow = 0 To MSGrid1.Cols - 1
      MSGrid1.row = 0
      MSGrid1.col = iRow
      MSGrid1.Text = arrGridHeadText(iRow)
      MSGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MSGrid1.CellAlignment = flexAlignCenterCenter
   Next
   
   '取得位置
   If colITS01 = 0 Then
        colITS01 = PUB_MGridGetId("ITS01", MSGrid1)
        colITS02 = PUB_MGridGetId("ITS02", MSGrid1)
        colCUID = PUB_MGridGetId("ITS07", MSGrid1)
        colPKey = PUB_MGridGetId("PKEY", MSGrid1)
        colORD1 = PUB_MGridGetId("ORD1", MSGrid1)
   End If
   
   MSGrid1.Visible = True
End Sub

'指定Grid的Row，改變顏色
'原因：加入記錄到Grid後，先依類別+記錄日期排序，再移動Grid的Row到新加入的那一筆，
          '但是只改變底色不設v也不帶入下方欄位
Private Sub SetSpecColor(ByVal idx As Integer, ByVal bolSpec As Boolean)
'idx: 指定nORD1
'bolSpec: True=反白，false=還原底色
Dim intJ As Integer, intK As Integer

    For intJ = 1 To MSGrid1.Rows - 1
        If Val(MSGrid1.TextMatrix(intJ, colORD1)) = idx Then
            MSGrid1.row = intJ
            For intK = 0 To MSGrid1.Cols - 1
                MSGrid1.col = intK
                If bolSpec = True Then '反白
                    MSGrid1.CellBackColor = &HFFC0C0
                Else
                    MSGrid1.CellBackColor = MSGrid1.BackColor
                End If
            Next intK
            If bolSpec = True Then
                MSGrid1.TopRow = intJ
            Else
                intSpecRow = 0
            End If
        End If
    Next intJ
End Sub

Private Sub Chk2_Click(Index As Integer)
   '當有資料並且功能列非動作時，重抓資料
   If txtData(1) <> "" And txtData(2) <> "" And TBar1.Buttons(11).Enabled = False And TBar1.Buttons(12).Enabled = False Then
      If ShowRecord Then
      End If
   End If
End Sub

'清單Word檔產生
Private Sub cmdWord_Click()
   If m_EditMode = 2 Then
      MsgBox "修改中不可執行清單！"
   ElseIf Trim(txtData(1)) = "" Or Trim(txtData(2)) = "" Then
      MsgBox "無清單！"
   Else
      'Add By Sindy 2025/8/27
      pub_QL05 = m_pub_QL05 & "(清單Word)"
      If pub_QL04 <> "" Then InsertQueryLog (MSGrid1.Rows - 1)
      '2025/8/27 End
      
      'Memo by Lydia 2022/11/25 改成共用模組
      strExc(0) = ""
      If Chk2(0).Value = 1 Then strExc(0) = strExc(0) & ",P"
      If Chk2(1).Value = 1 Then strExc(0) = strExc(0) & ",T"
      If PUB_GetITStoList(Me.Name, txtData(1), Trim(txtData(2)), False, IIf(Check1.Value = False, False, True), strExc(0)) = True Then
      End If
   End If
End Sub


'Added by Lydia 2020/08/13 各項指示分類查詢
Private Sub cmdHelp_Click()
    frm140415_1.Show
End Sub

'Added by Lydia 2020/08/13 開啟／關閉功能鈕
Private Sub ProcHelp(ByVal bEnabled As Boolean)
    If bEnabled = False Then
         cmdWord.Enabled = False
         Check1.Enabled = False
         Chk2(0).Enabled = False
         Chk2(1).Enabled = False
         '目前不限制，保留
         ''關閉各項指示分類查詢
         'If PUB_CheckFormExist("frm140415_1") Then
         '   Unload frm140415_1
         'End If
         'CmdHelp.Enabled = False
         cmdConfirm.Enabled = False
    Else
         cmdWord.Enabled = True
         Check1.Enabled = True
         Chk2(0).Enabled = True
         Chk2(1).Enabled = True
         'CmdHelp.Enabled = True '目前不限制，保留
         cmdConfirm.Enabled = True
    End If
End Sub

'Added by Lydia 2020/08/25 依使用者部門進行完成確認
Private Sub cmdConfirm_Click()
'------------------------------
'IC03,IC04 確認人員1+日期1(FCP)         => ST03=F2X
'IC05,IC06 確認人員2+日期2(P,CFP)      => ST03=P1X
'IC07,IC08 確認人員3+日期3(FCT,CFT) => ST03=F1X
'IC09,IC10 確認人員4+日期4(T)             =>ST03=P2X
'------------------------------

    strExc(1) = "": strExc(3) = IIf(lblConfirm.Caption <> lblConfirm.Tag, "、", "")
    Select Case Left(m_StrUserST03, 2)
         Case "F2"
              strExc(1) = ", IC03='" & strUserNum & "', IC04='" & strSrvDate(1) & "' "
         Case "P1"
              strExc(1) = ", IC05='" & strUserNum & "', IC06='" & strSrvDate(1) & "' "
         Case "F1"
              strExc(1) = ", IC07='" & strUserNum & "', IC08='" & strSrvDate(1) & "' "
         Case "P2"
              strExc(1) = ", IC09='" & strUserNum & "', IC10='" & strSrvDate(1) & "' "
         Case Else
              MsgBox "不屬於可確認部門！", vbInformation
    End Select
    If strExc(1) <> "" And txtData(1) <> "" And txtData(2) <> "" Then
         'Added by Lydia 2020/09/02 再次確認,避免誤按
         If MsgBox("是否完成確認？" & vbCrLf & "完成確認：撰寫信函不會再顯示原備註", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
             Exit Sub
         End If
         'edn 2020/09/02
         strSql = "Update InstConfirm Set " & Mid(strExc(1), 2) & " where IC01='" & txtData(1) & "' and IC02='" & txtData(2) & "' "
         cnnConnection.Execute strSql, intI
         'Added by Lydia 2020/12/04 若是整批新增各項指示，需要補確認記錄; ex.FCP-58306
         If intI = 0 Then
              strExc(2) = "insert into InstConfirm (IC01,IC02) values ('" & txtData(1) & "', '" & txtData(2) & "' ) "
              cnnConnection.Execute strExc(2), intI
              cnnConnection.Execute strSql, intI
         End If
         'end 2020/12/04
         cmdConfirm.Visible = False
         'Modified by Lydia 2022/11/25 改成公用模組
         'lblConfirm.Caption = lblConfirm.Caption & strExc(3)
         'Call GetICstatus
         lblConfirm.Caption = lblConfirm.Tag & PUB_GetICstatus(txtData(2))
         'end 2022/11/25
    End If
End Sub

'Modified by Lydia 2022/11/25 改成共用模組
'Private Function GetICstatus() As String
'      strExc(0) = "Select Decode(Ic04,Null,Null,'、外專')||Decode(Ic06,Null,Null,'、內專')||Decode(Ic08,Null,Null,'、外商')||Decode(Ic10,Null,Null,'、內商') as ChkType " & _
'                       "From Instconfirm where IC02='" & txtData(2) & "' "
'      intQ = 1
'      Set rsQuery = ClsLawReadRstMsg(intQ, strExc(0))
'      If intQ = 1 Then
'          If "" & rsQuery.Fields("ChkType") <> "" Then
'               lblConfirm.Caption = lblConfirm.Tag & Mid("" & rsQuery.Fields("ChkType"), 2)
'          End If
'      End If
'End Function

