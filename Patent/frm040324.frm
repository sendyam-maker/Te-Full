VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm040324 
   BorderStyle     =   1  '單線固定
   Caption         =   "年費逾期補繳通知函"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   945
   ClientWidth     =   7110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7110
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1425
      TabIndex        =   7
      Top             =   3900
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Height          =   792
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   6885
      Begin VB.CommandButton Command1 
         Caption         =   "尋找(&F)"
         Default         =   -1  'True
         Height          =   375
         Left            =   5190
         TabIndex        =   1
         Top             =   180
         Width           =   800
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "申請案號："
         Height          =   285
         Left            =   150
         TabIndex        =   6
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   5280
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6105
      TabIndex        =   4
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2235
      Left            =   135
      TabIndex        =   2
      Top             =   1500
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   3942
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label2 
      Caption         =   "地址條印表機："
      Height          =   180
      Index           =   1
      Left            =   105
      TabIndex        =   8
      Top             =   3915
      Width           =   1260
   End
End
Attribute VB_Name = "frm040324"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (MSHFlexGrid1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim intLastRow As Integer, intWhere As Integer
'Remove by Morgan 2008/8/13 改開窗定稿
'Dim m_OriPrinterName As String, SeekPrint As Integer, SeekPrintL As Integer, j As Integer, i As Integer
'Added by Lydia 2015/07/20 共用表單-實審請求期限屆滿前通知函
Public iKind As Integer '1-年費逾期補繳通知函,2-實審請求期限屆滿前通知函
'Add By Sindy 2017/12/29
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String, m_AppNo As String
Dim m_Done As Boolean
'2017/12/29 END


Public Sub Clear()
   Text7.Text = ""
   InitGrid 9, MSHFlexGrid1
   GridHead
   Me.Text7.SetFocus
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         FormConfirm
      Case 2 '結束
        Unload Me
   End Select
End Sub

Public Sub Command1_Click()
   intI = 0
   If Text7 = "" Then MsgBox "申請案號不得空白，請重新輸入 !", vbCritical: Exit Sub
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/30 清除查詢印表記錄檔欄位
   pub_QL05 = pub_QL05 & ";" & Label1 & Text7 'Add By Sindy 2010/11/30
     'Added by Lydia 2016/01/04 + 別名f0
      strExc(0) = "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
         "'' as RName,'',pa01,pa02,pa03,pa04,'' from patent f0 where PA01='P' AND " & _
         "pa11='" & Text7 & "' "
      'Added by Lydia 2015/07/20 限大陸案
      If iKind = 2 Then strExc(0) = strExc(0) & "and pa09='020' "
      'Added by Lydia 2016/01/04 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件(P)，但非此類案件時外專程序人員不可操作。
      strExc(0) = strExc(0) & FMP2openSQL
      strExc(0) = Replace(strExc(0), "f0.CP", "f0.PA")
      'end 2016/01/04
      
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2010/11/30
      If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
      Me.Tag = "0"
   GridHead
   If MSHFlexGrid1.Rows = 2 Then
      OnlyOneRec MSHFlexGrid1, 8
      FormConfirm
   End If
End Sub

Private Sub Form_Activate()
   'Added by Lydia 2015/07/20
   If iKind = 0 Then iKind = 1
   If iKind = 2 Then Me.Caption = "實審請求期限屆滿前通知函"
   
   'Added by Sindy 2017/12/29
   If m_strIR01 <> "" And m_Done = False Then
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/29 END
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   InitGrid 9, MSHFlexGrid1
   GridHead
   
   'Added by Lydia 2016/01/04 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm040324 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 8
   cmdOK(0).SetFocus
End Sub

' 確認鈕
Private Sub FormConfirm()
Dim bolChk As Boolean, i As Integer, j As Integer, strTmp(1 To 2) As String
    With MSHFlexGrid1
        .col = 8
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 8) = "v" Then
                bolChk = True
                For j = 1 To 4
                    strExc(j) = .TextMatrix(i, j + 3)
                Next
                strExc(5) = "1"
                Exit For
            End If
        Next
    End With
    If bolChk = False Then
        MsgBox "請選擇資料 !", vbInformation
        Exit Sub
    End If
    If Me.MSHFlexGrid1.Rows = 2 Then
        MSHFlexGrid1_Click
    End If
    
   'Add By Sindy 2017/12/29
   If m_strIR01 <> "" Then
      If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> strExc(1) & strExc(2) & strExc(3) & strExc(4) Then
         MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
         Exit Sub
      End If
   End If
   '2017/12/29 END
   
    'Added by Lydia 2015/07/20
    If iKind = 1 Then
        '年費逾期補繳通知函
        Load frm040324_1
        frm040324_1.Text1(0).Text = Me.Text7.Text
        frm040324_1.Label1(0).Caption = strExc(1) & "-" & strExc(2) & "-" & strExc(3) & "-" & strExc(4)
        'Add By Sindy 2017/12/29
        frm040324_1.m_strIR01 = m_strIR01
        frm040324_1.m_strIR02 = m_strIR02
        frm040324_1.m_strIR03 = m_strIR03
        frm040324_1.m_strIR04 = m_strIR04
        '2017/12/29 END
        frm040324_1.Show
    Else
       '實審請求期限屆滿前通知函
        frm040324_2.m_NowNo = strExc(1) & "-" & strExc(2) & "-" & strExc(3) & "-" & strExc(4)
        frm040324_2.m_NowPA11 = Me.Text7.Text
        Load frm040324_2
        'Added by Lydia 2015/08/12 判斷是否有資料
        If frm040324_2.m_bolRead = False Then
           Call frm040324_2.cmdExit_Click
           Exit Sub
        End If
        'end 2015/08/12
        'Add By Sindy 2017/12/29
        frm040324_2.m_strIR01 = m_strIR01
        frm040324_2.m_strIR02 = m_strIR02
        frm040324_2.m_strIR03 = m_strIR03
        frm040324_2.m_strIR04 = m_strIR04
        '2017/12/29 END
        frm040324_2.Show
    End If
    Me.Hide
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 1500: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 4000: .Text = "專利名稱"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1500: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      For i = 3 To 8
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub Text7_GotFocus()
   InverseTextBox Text7
End Sub
