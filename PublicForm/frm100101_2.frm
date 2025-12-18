VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件資料及案件進度查詢"
   ClientHeight    =   5748
   ClientLeft      =   1812
   ClientTop       =   2592
   ClientWidth     =   9468
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9468
   Begin VB.CommandButton cmdOK 
      Caption         =   "行事曆"
      Height          =   345
      Index           =   18
      Left            =   24
      TabIndex        =   39
      Top             =   696
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "管制備註"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   17
      Left            =   60
      Style           =   1  '圖片外觀
      TabIndex        =   36
      Top             =   0
      Width           =   860
   End
   Begin VB.CommandButton cmdCustAtt 
      Caption         =   "客戶附件"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   2556
      Style           =   1  '圖片外觀
      TabIndex        =   33
      Top             =   375
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "繳費單"
      Height          =   345
      Index           =   16
      Left            =   60
      TabIndex        =   32
      Top             =   375
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "承辦歷程(聯絡)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   15
      Left            =   6456
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   375
      Width           =   1425
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "原始檔"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   14
      Left            =   876
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   375
      Width           =   840
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "卷宗區"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   13
      Left            =   1716
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   375
      Width           =   840
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "工時統計"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   12
      Left            =   3396
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   375
      Width           =   840
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "不含程序(&Q)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   11
      Left            =   4236
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   375
      Width           =   1104
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "不含未收費"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   10
      Left            =   5364
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   375
      Width           =   1068
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "相關卷號"
      Height          =   345
      Index           =   9
      Left            =   1770
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   860
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "分割案"
      Height          =   345
      Index           =   8
      Left            =   900
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   860
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "變更事項"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   7
      Left            =   2640
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   0
      Width           =   860
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "承辦進度"
      Height          =   345
      Index           =   3
      Left            =   6120
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   0
      Width           =   860
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "繪圖進度"
      Height          =   345
      Index           =   4
      Left            =   6990
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   0
      Width           =   860
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Bindings        =   "frm100101_2.frx":0000
      Height          =   4104
      Left            =   24
      TabIndex        =   0
      Top             =   1608
      Width           =   9288
      _ExtentX        =   16383
      _ExtentY        =   7260
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
      _Band(0).Cols   =   18
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "不含來函"
      Height          =   345
      Index           =   0
      Left            =   3510
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   0
      Width           =   860
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   5
      Left            =   7860
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   0
      Width           =   860
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度"
      Height          =   345
      Index           =   2
      Left            =   5250
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   0
      Width           =   860
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "含來函"
      Height          =   345
      Index           =   1
      Left            =   4380
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   0
      Width           =   860
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   6
      Left            =   8715
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "目前智權人員："
      Height          =   180
      Left            =   3510
      TabIndex        =   27
      Top             =   1395
      Width           =   1260
   End
   Begin MSForms.Label Label15 
      Height          =   195
      Left            =   990
      TabIndex        =   38
      Top             =   1380
      Width           =   2445
      VariousPropertyBits=   27
      Size            =   "4313;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label16 
      Height          =   255
      Left            =   4800
      TabIndex        =   37
      Top             =   1380
      Width           =   945
      VariousPropertyBits=   27
      Size            =   "1667;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   330
      Left            =   960
      TabIndex        =   1
      Top             =   1020
      Width           =   7635
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13462;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNp605 
      Caption         =   "lblNp605"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   8440
      TabIndex        =   35
      Top             =   795
      Width           =   1095
   End
   Begin VB.Label lblCaseMap2 
      AutoSize        =   -1  'True
      Caption         =   "lblCaseMap2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   8040
      TabIndex        =   34
      Top             =   570
      Width           =   1410
   End
   Begin VB.Label lblCMboth 
      Caption         =   "lblCMboth"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   8670
      TabIndex        =   31
      Top             =   1005
      Width           =   945
   End
   Begin VB.Label lblCaseMap 
      AutoSize        =   -1  'True
      Caption         =   "lblCaseMap"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8040
      TabIndex        =   30
      Top             =   360
      Width           =   1380
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "註：可點選欄位標題做資料排序"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   6075
      TabIndex        =   29
      Top             =   1395
      Width           =   2520
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   60
      TabIndex        =   28
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      Caption         =   "lblCancel"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   8670
      TabIndex        =   26
      Top             =   1410
      Width           =   795
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "分所號："
      Height          =   180
      Left            =   240
      TabIndex        =   25
      Top             =   1395
      Width           =   720
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   180
      Left            =   6720
      TabIndex        =   24
      Top             =   795
      Width           =   1635
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   180
      Left            =   3645
      TabIndex        =   21
      Top             =   795
      Width           =   1875
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   180
      Left            =   990
      TabIndex        =   20
      Top             =   795
      Width           =   1665
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "審定證書號："
      Height          =   180
      Left            =   5640
      TabIndex        =   23
      Top             =   795
      Width           =   1080
   End
   Begin VB.Label lblClose 
      AutoSize        =   -1  'True
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   8670
      TabIndex        =   22
      Top             =   1230
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Left            =   2745
      TabIndex        =   19
      Top             =   795
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   60
      TabIndex        =   18
      Top             =   795
      Width           =   900
   End
End
Attribute VB_Name = "frm100101_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/05/24 Form2.0已修改: Combo1、Label16; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
'edit by nickc 2005/10/06 重整
Option Explicit

Dim i As Integer, j As Integer
Dim StrTag As String
Dim intK As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'add by nickc 2006/07/10 收款情形
Dim IntTemp1 As Long
Dim IntTemp2 As Long
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim mFixs As Integer 'Added by Lydia 2018/03/22 固定欄位
Dim strCon As String 'Added by Lydia 2019/05/30 依國別抓案件性質欄位
Dim pa(5) As String 'Add By Sindy 2019/9/17
Dim m_FixNo As Integer 'Add By Sindy 2019/9/17 修法次數
Dim strFeeType As String, strYF15 As String, strPA09 As String 'Add By Sindy 2019/9/17
Public bolEmpFlow As Boolean 'Add By Sindy 2020/9/25 從歷程維護作業指過來
Public m_CKind As String '是否含C類來函 N:不含 Modify By Sindy 2021/4/21 不要為了排除C類而寫2個函數 StrMenu,StrMenu1(Mark)
Public m_PrevForm As Form 'Added by Lydia 2025/09/10

'Added by Lydia 2025/09/10
Public Sub SetParent(ByRef pFrm As Form, Optional ByVal pCaseNo As String)
   
   Set m_PrevForm = pFrm
End Sub

Private Sub SetDataListWidth()
'add by nickc 2005/10/06
Dim o_strSQL As String
Dim o_Str01 As String
Dim ii As Integer 'Added by Morgan 2024/7/10

If Left(Me.Tag, 1) = "N" Then
   o_strSQL = Right(Me.Tag, Len(Me.Tag) - 1)
Else
   o_strSQL = Me.Tag
End If
o_Str01 = SystemNumber(o_strSQL, 1)
'非法務案件依照原本欄位格式
'Modify By Sindy 2009/07/24 增加LIN系統類別
'modify by sonia 2019/7/30 +ACS系統類別
'Modify by Amy 2020/11/13 原:And o_Str01 <> "ACS"
If (o_Str01 <> "L" And o_Str01 <> "FCL" And o_Str01 <> "CFL" And _
   o_Str01 <> "LA" And o_Str01 <> "LIN") Or o_Str01 = "ACS" Then
         '加欄位:相關總收文號, 本所期限, 法定期限
         'grdDataList.Cols = 29
         'Add by Lydia 2014/11/14 進度檔結果欄後加入下一期限及下一期限日期
         'Modified by Morgan 2024/7/10 改控制後面欄位不顯示就好(有欄位不顯示但要拿來控制該列是否顯示 Ex:承辦人部門(pdept))
         'grdDataList.Cols = 32 '31
         If grdDataList.Cols < 32 Then grdDataList.Cols = 32
         For ii = 32 To grdDataList.Cols - 1
            grdDataList.ColWidth(ii) = 0
         Next
         'end 2024/7/10
         
         grdDataList.row = 0
         grdDataList.col = 0: grdDataList.Text = "V"
         grdDataList.ColWidth(0) = 180
         grdDataList.col = 1: grdDataList.Text = "收文日"
         grdDataList.ColWidth(1) = 788 '原795
         grdDataList.CellAlignment = flexAlignRightCenter
         grdDataList.col = 2: grdDataList.Text = "總收文號"
         grdDataList.ColWidth(2) = 938 '原960
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 3: grdDataList.Text = "案件性質"
         grdDataList.ColWidth(3) = 950 '原800
         grdDataList.CellAlignment = flexAlignLeftCenter
         
         grdDataList.col = 4: grdDataList.Text = "相關收文號"
         grdDataList.ColWidth(4) = 938 '原1150
         grdDataList.CellAlignment = flexAlignLeftCenter
         
         grdDataList.col = 5: grdDataList.Text = "承辦人"
         grdDataList.ColWidth(5) = 593
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 6: grdDataList.Text = "智權人員"
         grdDataList.ColWidth(6) = 593
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 7: grdDataList.Text = "本所期限"
         grdDataList.ColWidth(7) = 788 '原780
         grdDataList.CellAlignment = flexAlignRightCenter
         grdDataList.col = 8: grdDataList.Text = "法定期限"
         grdDataList.ColWidth(8) = 788 '原780
         grdDataList.CellAlignment = flexAlignRightCenter
         'modify by sonia 2019/7/24 發文日改專業發文日
         grdDataList.col = 9: grdDataList.Text = "專業發文日"
         grdDataList.ColWidth(9) = 788 '原780
         grdDataList.CellAlignment = flexAlignLeftCenter
         'Add By Cheng 2002/11/11
         
      'Modify By Sindy 2021/4/29 外專部門加看約定期限(取代"取消收文日"的位置,取消收文日改放在結果之後)
      If Left(Pub_StrUserSt03, 2) = "F2" Or Pub_StrUserSt03 = "M51" Then
         grdDataList.col = 10: grdDataList.Text = "約定期限"
         If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
            grdDataList.ColWidth(10) = 788 '原995
         Else
            grdDataList.ColWidth(10) = 0
         End If
         grdDataList.CellAlignment = flexAlignLeftCenter
         
         '是否出名與代理人欄位對調
         grdDataList.col = 11: grdDataList.Text = "代理人"
         If bolFNation = False Then
             grdDataList.ColWidth(11) = 0
         Else
             grdDataList.ColWidth(11) = 590
         End If
         grdDataList.col = 12: grdDataList.Text = "結果"
         grdDataList.ColWidth(12) = 500
         grdDataList.CellAlignment = flexAlignLeftCenter
         
         grdDataList.col = 13: grdDataList.Text = "取消收文"
         grdDataList.ColWidth(13) = 788 '原995
         grdDataList.CellAlignment = flexAlignLeftCenter
      Else
      '2021/4/29 END
         grdDataList.col = 10: grdDataList.Text = "取消收文"
         grdDataList.ColWidth(10) = 788 '原995
         grdDataList.CellAlignment = flexAlignLeftCenter
         
         '是否出名與代理人欄位對調
         grdDataList.col = 11: grdDataList.Text = "代理人"
         If bolFNation = False Then
             grdDataList.ColWidth(11) = 0
         Else
             grdDataList.ColWidth(11) = 590
         End If
         grdDataList.col = 12: grdDataList.Text = "結果"
         grdDataList.ColWidth(12) = 500
         grdDataList.CellAlignment = flexAlignLeftCenter
         
         'Add By Sindy 2021/4/29
         grdDataList.col = 13: grdDataList.Text = "(暫無)"
         grdDataList.ColWidth(13) = 0
         grdDataList.CellAlignment = flexAlignLeftCenter
         '2021/4/29 END
      End If
      
         grdDataList.col = 14: grdDataList.Text = "繳費年度"
         grdDataList.ColWidth(14) = 800
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 15: grdDataList.Text = "是否雙倍"
         grdDataList.ColWidth(15) = 800
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 16: grdDataList.Text = "相關人"
         grdDataList.ColWidth(16) = 600
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 17: grdDataList.Text = "是否出名"
         grdDataList.ColWidth(17) = 800
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 18: grdDataList.Text = "進度備註"
         grdDataList.ColWidth(18) = 800
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 19
         grdDataList.ColWidth(19) = 0
         grdDataList.col = 20
         grdDataList.ColWidth(20) = 0
         grdDataList.col = 21
         grdDataList.ColWidth(21) = 0
         grdDataList.col = 22
         grdDataList.ColWidth(22) = 0
         grdDataList.col = 23
         grdDataList.ColWidth(23) = 0
         grdDataList.col = 24
         grdDataList.ColWidth(24) = 0
         grdDataList.col = 25
         grdDataList.ColWidth(25) = 0
         'add by nickc 2006/07/10
         grdDataList.col = 26: grdDataList.Text = "收款情形" 'CP60
         grdDataList.ColWidth(26) = 1000
         grdDataList.CellAlignment = flexAlignLeftCenter
         'Add By Sindy 2011/6/23
         grdDataList.col = 27
         grdDataList.ColWidth(27) = 0
         grdDataList.col = 28
         grdDataList.ColWidth(28) = 0
         grdDataList.col = 29
         grdDataList.ColWidth(29) = 0
         'Add by Lydia 2014/11/14 進度檔結果欄後加入下一程序及下一程序本所期限
         grdDataList.col = 30: grdDataList.Text = "下一程序"
         grdDataList.ColWidth(30) = 800
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 31: grdDataList.Text = "下一程序所限"
         grdDataList.ColWidth(31) = 1200
         grdDataList.CellAlignment = flexAlignLeftCenter
         
         
Else '法務格式
         '2005/11/29 ADD BY SONIA
         Label1.Visible = False
         Label11.Visible = False
         '2005/11/29 END
         grdDataList.Cols = 32 '31
         grdDataList.row = 0
         grdDataList.col = 0: grdDataList.Text = "V"
         grdDataList.ColWidth(0) = 200
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 1: grdDataList.Text = "收文日"
         grdDataList.ColWidth(1) = 795
         grdDataList.CellAlignment = flexAlignRightCenter
         grdDataList.col = 2: grdDataList.Text = "總收文號"
         grdDataList.ColWidth(2) = 960
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 3: grdDataList.Text = "備註主題(案件性質)"
         grdDataList.ColWidth(3) = 1700
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 4: grdDataList.Text = "相對人"  '2005/12/19 改相關人為相對人
         grdDataList.ColWidth(4) = 600
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 5: grdDataList.Text = "相關總收文號"
         '2005/12/20 MODIFY BY SONIA 暫不顯示
         'GrdDataList.ColWidth(5) = 1150
         grdDataList.ColWidth(5) = 0
         grdDataList.CellAlignment = flexAlignLeftCenter
         'Modified by Lydia 2015/10/05
'         GrdDataList.col = 6: GrdDataList.Text = "承辦律師"
         grdDataList.col = 6: grdDataList.Text = "承辦人"
         grdDataList.ColWidth(6) = 800
         grdDataList.CellAlignment = flexAlignLeftCenter
         'Modified by Lydia 2015/10/05
         'GrdDataList.col = 7: GrdDataList.Text = "承辦法務"
         grdDataList.col = 7: grdDataList.Text = "協辦人員"
         grdDataList.ColWidth(7) = 800
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 8: grdDataList.Text = "智權人員"
         grdDataList.ColWidth(8) = 600
         grdDataList.CellAlignment = flexAlignLeftCenter
         'Added by Lydia 2020/07/15 法律所案源收文：增加案源之介紹人
         grdDataList.col = 9: grdDataList.Text = "介紹人"
         If strSrvDate(1) < 法律所案源收文啟用日 Then
            grdDataList.ColWidth(9) = 0
         Else
            grdDataList.ColWidth(9) = 600
         End If
         grdDataList.CellAlignment = flexAlignLeftCenter
         'end 2020/07/15
         'Modified by Lydia 2020/07/15 index+1
         grdDataList.col = 10: grdDataList.Text = "點數"
         grdDataList.ColWidth(10) = 600
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 11: grdDataList.Text = "本所期限"
         grdDataList.ColWidth(11) = 795
         grdDataList.CellAlignment = flexAlignRightCenter
         grdDataList.col = 12: grdDataList.Text = "法定期限"
         grdDataList.ColWidth(12) = 795
         grdDataList.CellAlignment = flexAlignRightCenter
         grdDataList.col = 13: grdDataList.Text = "會稿日"
         grdDataList.ColWidth(13) = 795
         grdDataList.CellAlignment = flexAlignRightCenter
         grdDataList.col = 14: grdDataList.Text = "發文日"
         grdDataList.ColWidth(14) = 795
         grdDataList.CellAlignment = flexAlignRightCenter
         grdDataList.col = 15: grdDataList.Text = "回執日"
         grdDataList.ColWidth(15) = 795
         grdDataList.CellAlignment = flexAlignRightCenter
         grdDataList.col = 16: grdDataList.Text = "取消收文日"
         grdDataList.ColWidth(16) = 980
         grdDataList.CellAlignment = flexAlignRightCenter
         '2005/12/19 MODIFY BY SONIA
         'GrdDataList.Col = 16: GrdDataList.Text = "機關名稱"
         'GrdDataList.ColWidth(16) = 800
         grdDataList.col = 17: grdDataList.Text = "代理人"
         grdDataList.ColWidth(17) = 1200
         grdDataList.CellAlignment = flexAlignLeftCenter
         '2005/12/19 END
         'end 'Modified by Lydia 2020/07/15 index+1
         'Memo by Lydia 2020/07/15 拿掉代理人和LC15 之間的null
         grdDataList.col = 18
         grdDataList.ColWidth(18) = 0
         grdDataList.col = 19
         grdDataList.ColWidth(19) = 0
         grdDataList.col = 20
         grdDataList.ColWidth(20) = 0
         grdDataList.col = 21
         grdDataList.ColWidth(21) = 0
         grdDataList.col = 22
         grdDataList.ColWidth(22) = 0
         grdDataList.col = 23
         grdDataList.ColWidth(23) = 0
         grdDataList.col = 24
         grdDataList.ColWidth(24) = 0
         grdDataList.col = 25
         grdDataList.ColWidth(25) = 0
         'add by nickc 2006/07/10
         grdDataList.col = 26
         grdDataList.ColWidth(26) = 0
         'Add By Sindy 2011/6/23
         grdDataList.col = 27
         grdDataList.ColWidth(27) = 0
         grdDataList.col = 28
         grdDataList.ColWidth(28) = 0
      'Add By Sindy 2021/4/29
         grdDataList.col = 29
         grdDataList.ColWidth(29) = 0
      '2021/4/29 END
         'Add by Lydia 2014/11/14 進度檔結果欄後加入下一程序及下一程序本所期限
         grdDataList.col = 30: grdDataList.Text = "下一程序"
         grdDataList.ColWidth(30) = 800
         grdDataList.CellAlignment = flexAlignLeftCenter
         grdDataList.col = 31: grdDataList.Text = "下一程序所限"
         grdDataList.ColWidth(31) = 1200
         grdDataList.CellAlignment = flexAlignLeftCenter
End If
End Sub

'Add By Sindy 2011/2/25
Private Sub SetDataListWidth_new()
Dim o_strSQL As String
Dim o_Str01 As String
If Left(Me.Tag, 1) = "N" Then
   o_strSQL = Right(Me.Tag, Len(Me.Tag) - 1)
Else
   o_strSQL = Me.Tag
End If
o_Str01 = SystemNumber(o_strSQL, 1)
'非法務案件依照原本欄位格式
'modify by sonia 2019/7/30 +ACS系統類別
'Modify by Amy 2020/11/13 原:And o_Str01 <> "ACS"
If (o_Str01 <> "L" And o_Str01 <> "FCL" And o_Str01 <> "CFL" And _
   o_Str01 <> "LA" And o_Str01 <> "LIN") Or o_Str01 = "ACS" Then
   grdDataList.ColAlignment(1) = flexAlignRightCenter
   grdDataList.ColAlignment(7) = flexAlignRightCenter
   grdDataList.ColAlignment(8) = flexAlignRightCenter
   grdDataList.ColAlignment(9) = flexAlignRightCenter
Else
   grdDataList.ColAlignment(1) = flexAlignRightCenter
   grdDataList.ColAlignment(10) = flexAlignRightCenter
   grdDataList.ColAlignment(11) = flexAlignRightCenter
   grdDataList.ColAlignment(12) = flexAlignRightCenter
   grdDataList.ColAlignment(13) = flexAlignRightCenter
   grdDataList.ColAlignment(14) = flexAlignRightCenter
   grdDataList.ColAlignment(15) = flexAlignRightCenter
End If
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
Dim strCP01 As String 'Add By Sindy 2012/5/21
Dim strCP02 As String, strCP03 As String, strCP04 As String 'Add By Sindy 2018/12/14
Dim iColX As Integer 'Added by Morgan 2024/7/10

   'Add By Sindy 2018/12/14
   If Trim(Label3.Caption) <> "" Then
      strCP01 = SystemNumber(Label3.Caption, 1)
      strCP02 = SystemNumber(Label3.Caption, 2)
      strCP03 = SystemNumber(Label3.Caption, 3)
      strCP04 = SystemNumber(Label3.Caption, 4)
   End If
   '2018/12/14 END
   Select Case cmdState
   Case 0
        'Modify By Sindy 2021/4/21 不要為了排除C類而寫2個函數 StrMenu,StrMenu1(Mark)
        'Call StrMenu1
        Me.m_CKind = "N"
        Call StrMenu
        '2021/4/21 END
        m_blnColOrderAsc = True    'Add By Sindy 2011/6/23
   Case 1
        Me.m_CKind = "" 'Add By Sindy 2021/4/21 不要為了排除C類而寫2個函數 StrMenu,StrMenu1(Mark)
        Call StrMenu
        m_blnColOrderAsc = True    'Add By Sindy 2011/6/23
   Case 2
        Me.Enabled = False
        StrTag = ""
        For i = 1 To grdDataList.Rows - 1
        grdDataList.col = 0
        grdDataList.row = i
        If Trim(grdDataList.Text) = "V" Then
           grdDataList.col = 0
           grdDataList.Text = ""
           For j = 0 To grdDataList.Cols - 1
               'Modified by Lydia 2018/03/22 固定欄位不變色
               'If j > 5 Or Label1.Visible = True Then
               If j > mFixs Or mFixs = 0 Then
                   grdDataList.col = j
                   grdDataList.CellBackColor = QBColor(15)
               End If
           Next j
            grdDataList.col = 2
            If Not IsNull(grdDataList.Text) Then
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               Screen.MousePointer = vbHourglass
               frm100101_C.Show
               frm100101_C.Tag = Label3.Caption + "=" + Pub_RplStr(grdDataList.Text)
               frm100101_C.StrMenu
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
            End If
        End If
        Next i
        Me.Enabled = True
   Case 3
        Me.Enabled = False
        StrTag = ""
        For i = 1 To grdDataList.Rows - 1
        grdDataList.col = 0
        grdDataList.row = i
        If Trim(grdDataList.Text) = "V" Then
           grdDataList.col = 0
           grdDataList.Text = ""
           For j = 0 To grdDataList.Cols - 1
              'Modified by Lydia 2018/03/22 固定欄位不變色
              'If j > 5 Or Label1.Visible = True Then
              If j > mFixs Or mFixs = 0 Then
                   grdDataList.col = j
                   grdDataList.CellBackColor = QBColor(15)
              End If
           Next j
            grdDataList.col = 2
            If Not IsNull(grdDataList.Text) Then
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               Screen.MousePointer = vbHourglass
               'Modify By Sindy 2012/5/21 +if,frm100101_K
               strCP01 = GetCaseProData(Trim(Pub_RplStr(grdDataList.Text)), "CP01")
               If strCP01 = "P" Or strCP01 = "PS" Or strCP01 = "FG" Or _
                  strCP01 = "FCP" Or strCP01 = "CFP" Or strCP01 = "CPS" Or _
                  Val(strSrvDate(1)) < Val(TMdebateStarDT) Then  '專利處工作進度
                  frm100101_F.Show
                  frm100101_F.Process Pub_RplStr(grdDataList.Text)
               Else
                  frm100101_K.Show
                  frm100101_K.Process Pub_RplStr(grdDataList.Text)
               End If
               '2012/5/21 End
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
            End If
        End If
        Next i
        Me.Enabled = True
   Case 4
        Me.Enabled = False
        StrTag = ""
        For i = 1 To grdDataList.Rows - 1
        grdDataList.col = 0
        grdDataList.row = i
        If Trim(grdDataList.Text) = "V" Then
           grdDataList.col = 0
           grdDataList.Text = ""
           For j = 0 To grdDataList.Cols - 1
               'Modified by Lydia 2018/03/22 固定欄位不變色
               'If j > 5 Or Label1.Visible = True Then
               If j > mFixs Or mFixs = 0 Then
                   grdDataList.col = j
                   grdDataList.CellBackColor = QBColor(15)
               End If
           Next j
            grdDataList.col = 2
            If Not IsNull(grdDataList.Text) Then
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               Screen.MousePointer = vbHourglass
               frm100101_g.Show
               frm100101_g.Process Pub_RplStr(grdDataList.Text)
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
            End If
        End If
        Next i
        Me.Enabled = True
   Case 5
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Case 6
        fnCloseAllFrm100
   'add by nick 2004/10/14 變更事項
   Case 7
        Me.Enabled = False
        StrTag = ""
        For i = 1 To grdDataList.Rows - 1
        grdDataList.col = 0
        grdDataList.row = i
        If Trim(grdDataList.Text) = "V" Then
           grdDataList.col = 0
           grdDataList.Text = ""
           For j = 0 To grdDataList.Cols - 1
               'Modified by Lydia 2018/03/22 固定欄位不變色
               'If j > 5 Or Label1.Visible = True Then
               If j > mFixs Or mFixs = 0 Then
                   grdDataList.col = j
                   grdDataList.CellBackColor = QBColor(15)
               End If
           Next j
            grdDataList.col = 2
            If Not IsNull(grdDataList.Text) Then
               Screen.MousePointer = vbHourglass
               frm050706.Show
               frm050706.Hide
               frm050706.m_bDelete = False
               frm050706.m_bInsert = False
               frm050706.m_bUpdate = False
               frm050706.IsCall = True
               frm050706.textCE01.Text = Pub_RplStr(grdDataList.Text)
               If frm050706.QueryRecord = True Then
                   frm050706.Show
                   If fnSaveParentForm(Me) = False Then
                       Me.Enabled = True
                       Exit Sub
                   End If
                   Screen.MousePointer = vbDefault
                   Me.Enabled = True
                   Exit Sub
               'add by nickc 2005/06/09
               Else
                   MsgBox "無變更事項！", vbCritical, "警告！"
               End If
               Screen.MousePointer = vbDefault
            End If
        End If
        Next i
        Me.Enabled = True
   'add by nick 2004/09/15
   '相關卷號
   Case 9
        cmdState = -1
        Me.Enabled = False
        If fnSaveParentForm(Me) = False Then
           Me.Enabled = True
           Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        frm100108_3.Show
        frm100108_3.Tag = Label3.Caption
        frm100108_3.Caption = "相關卷號"
        frm100108_3.StrMenu2
        Screen.MousePointer = vbDefault
        Me.Enabled = True
   '分割案
   Case 8
        cmdState = -1
        Me.Enabled = False
        If fnSaveParentForm(Me) = False Then
           Me.Enabled = True
           Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        frm100108_4.Show
        frm100108_4.frm100108_txt_7 = "3"
        frm100108_4.SetDataListWidth
        frm100108_4.Tag = Label3.Caption
        frm100108_4.Caption = "分割案"
        frm100108_4.StrMenu1
        Screen.MousePointer = vbDefault
        Me.Enabled = True
   'Add By Sindy 2011/6/23
   '不含未收費
   Case 10
         Me.Enabled = False
         Screen.MousePointer = vbHourglass
         'Added by Morgan 2024/7/10
         If cmdOK(10).Caption = "不含未收費" Then
            cmdOK(10).Caption = "含未收費"
            intI = 0
            cmdOK(11).Enabled = False
         Else
            cmdOK(10).Caption = "不含未收費"
            intI = 255
            cmdOK(11).Enabled = True
         End If
         'end 2024/7/10
         For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 29 '28
            grdDataList.row = i
            If Val(grdDataList.Text) = 0 Then
               'Modified by Morgan 2024/7/10
               'grdDataList.RowHeight(i) = 0
               grdDataList.RowHeight(i) = intI
               'end 2024/7/10
            End If
         Next i
         
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         
   'Add By Sindy 2011/6/23
   '含未收費
   'Modified by Morgan 2024/7/10 空間不夠，原功能與上面功能合併
   'Case 11
   '      Me.Enabled = False
   '      Screen.MousePointer = vbHourglass
   '      For i = 1 To grdDataList.Rows - 1
   '         grdDataList.col = 29 '28
   '         grdDataList.row = i
   '         If Val(grdDataList.Text) = 0 Then
   '            grdDataList.RowHeight(i) = 255
   '         End If
   '      Next i
   '      Screen.MousePointer = vbDefault
   '      Me.Enabled = True
   '不含程序
   Case 11
         Me.Enabled = False
         Screen.MousePointer = vbHourglass
         If cmdOK(11).Caption = "不含程序(&Q)" Then
            cmdOK(11).Caption = "含程序(&Q)"
            intI = 0
            cmdOK(10).Enabled = False
         Else
            cmdOK(11).Caption = "不含程序(&Q)"
            intI = 255
            cmdOK(10).Enabled = True
         End If
         
         iColX = -1
         For i = 1 To grdDataList.Cols - 1
            If UCase(grdDataList.TextMatrix(0, i)) = "PDEPT" Then
               iColX = i
               Exit For
            End If
         Next
         If iColX <> -1 Then
            For i = 1 To grdDataList.Rows - 1
               grdDataList.col = iColX '28
               grdDataList.row = i
               If intI = 0 Then
                  If InStr("F12,F22,P12,P22", grdDataList.Text) > 0 Then
                     grdDataList.RowHeight(i) = intI
                  End If
               Else
                  grdDataList.RowHeight(i) = intI
               End If
            Next i
         End If
         Screen.MousePointer = vbDefault
         Me.Enabled = True
   'end 2024/7/10
   
   'Add By Sindy 2011/6/23
   '工時統計
   Case 12
         cmdState = -1
         Me.Enabled = False
         If fnSaveParentForm(Me) = False Then
            Me.Enabled = True
            Exit Sub
         End If
         Screen.MousePointer = vbHourglass
         frm100101_i.Show
         frm100101_i.Tag = Label3.Caption
         frm100101_i.txtCode(1) = SystemNumber(Label3.Caption, 1)
         frm100101_i.txtCode(2) = SystemNumber(Label3.Caption, 2)
         frm100101_i.txtCode(3) = SystemNumber(Label3.Caption, 3)
         frm100101_i.txtCode(4) = SystemNumber(Label3.Caption, 4)
         frm100101_i.QueryData
         Screen.MousePointer = vbDefault
         Me.Enabled = True
   'Add By Sindy 2013/6/11
   '卷宗區
   Case 13
         Me.Enabled = False
         StrTag = ""
         For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 0
               grdDataList.Text = ""
               For j = 0 To grdDataList.Cols - 1
                  'Modified by Lydia 2018/03/22 固定欄位不變色
                  'If j > 5 Or Label1.Visible = True Then
                  If j > mFixs Or mFixs = 0 Then
                      grdDataList.col = j
                      grdDataList.CellBackColor = QBColor(15)
                  End If
               Next j
               grdDataList.col = 2
               If Not IsNull(grdDataList.Text) Then
                  StrTag = StrTag & Trim(grdDataList.Text) & ","
               End If
            End If
         Next i
         If StrTag <> "" Then
            'Add By Sindy 2025/6/9
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            '2025/6/9 END
            StrTag = Left(StrTag, Len(StrTag) - 1)
            Screen.MousePointer = vbHourglass
            frm100101_L.m_strKey = StrTag '多筆總收文號
            'frm100101_L.Hide
            frm100101_L.SetParent Me
            If frm100101_L.QueryData = True Then
               frm100101_L.Show
               Me.Hide
            End If
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
         Me.Enabled = True
   '原始檔
   Case 14
         Me.Enabled = False
         StrTag = ""
         For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 0
               grdDataList.Text = ""
               For j = 0 To grdDataList.Cols - 1
                  'Modified by Lydia 2018/03/22 固定欄位不變色
                  'If j > 5 Or Label1.Visible = True Then
                  If j > mFixs Or mFixs = 0 Then
                      grdDataList.col = j
                      grdDataList.CellBackColor = QBColor(15)
                  End If
               Next j
               grdDataList.col = 2
               If Not IsNull(grdDataList.Text) Then
                  StrTag = StrTag & Trim(grdDataList.Text) & ","
               End If
            End If
         Next i
         If StrTag <> "" Then
            'Add By Sindy 2025/6/9
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            '2025/6/9 END
            StrTag = Left(StrTag, Len(StrTag) - 1)
            Screen.MousePointer = vbHourglass
            frm100101_M.m_strKey = StrTag '多筆總收文號
            'frm100101_M.Hide
            frm100101_M.SetParent Me
            If frm100101_M.QueryData = True Then
               frm100101_M.Show
               Me.Hide
            End If
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
         Me.Enabled = True
   '2013/6/11 End
   'Modify By Sindy 2013/8/16
   Case 15 '聯絡事項-承辦歷程
         Me.Enabled = False
         StrTag = ""
         For i = 1 To grdDataList.Rows - 1
         grdDataList.col = 0
         grdDataList.row = i
         frm090202_2.ShowNextData = False
         If Trim(grdDataList.Text) = "V" Then
            grdDataList.col = 0
            grdDataList.Text = ""
            For j = 0 To grdDataList.Cols - 1
               'Modified by Lydia 2018/03/22 固定欄位不變色
               'If j > 5 Or Label1.Visible = True Then
               If j > mFixs Or mFixs = 0 Then
                  grdDataList.col = j
                  grdDataList.CellBackColor = QBColor(15)
               End If
            Next j
            grdDataList.col = 2
            If Not IsNull(grdDataList.Text) Then
               'Add By Sindy 2020/8/7
               '檢查表單是否已開啟，若是，則關閉
               If PUB_ChkFormIsClose("frm090202_2") = False Then
                  Me.Enabled = True
                  Exit Sub
               End If
               '2020/8/7 END
              
               Screen.MousePointer = vbHourglass
   '            If Trim(grdDataList.TextMatrix(i, 9)) = "" And _
   '               Trim(grdDataList.TextMatrix(i, 10)) = "" Then '未發文未取消收文才可新增聯絡
                  frm090202_2.Hide
                  frm090202_2.m_EEP01 = grdDataList.Text '總收文號
                  frm090202_2.intReceiveKind = 99 '聯絡
                  frm090202_2.SetParent Me
                  frm090202_2.Caption = frm090202_2.Caption ' & " - " & GrdDataList.TextMatrix(i, 3)
                  frm090202_2.cmdOK(0).Visible = False
                  frm090202_2.cmdOK(1).Visible = False
                  frm090202_2.Cmd1(0).Visible = False
                  If frm090202_2.QueryData = True Then
                     frm090202_2.ShowNextData = True
                     'Modify By Sindy 2014/7/10 Mark:薛經理說系統不用自動啟動聯絡流程,由使用者自行操作
   '                  If Trim(grdDataList.TextMatrix(i, 9)) = "" And _
   '                     Trim(grdDataList.TextMatrix(i, 10)) = "" Then
   '                     frm090202_2.cmdAdd_Click
   '                  End If
                     frm090202_2.Show
                     Me.Hide
                     Screen.MousePointer = vbDefault
                     Me.Enabled = True
                     Exit Sub
                  End If
   '            Else
   '               If Trim(grdDataList.TextMatrix(i, 9)) <> "" Then
   '                  MsgBox "已發文，不可新增聯絡！"
   '               ElseIf Trim(grdDataList.TextMatrix(i, 10)) <> "" Then
   '                  MsgBox "已取消收文，不可新增聯絡！"
   '               End If
   '            End If
               Screen.MousePointer = vbDefault
            End If
         End If
         Next i
         Me.Enabled = True
   '2013/8/16 End
   Case 16 '繳費單
      'Add By Sindy 2018/12/14 註冊費的繳費單
      Call PUB_PrintTFeeForm(strCP01, strCP02, strCP03, strCP04)
      '2018/12/14 END
   'Added by Lydia 2021/05/19
   Case 17  '管制備註
        Me.Enabled = False
        StrTag = ""
        strExc(1) = "": strExc(2) = "": strExc(3) = ""
        For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            'Memo by Lydia 2021/05/28 保留單筆的寫法
'            If Trim(grdDataList.Text) = "V" Then
'               grdDataList.col = 0
'               grdDataList.Text = ""
'               For j = 0 To grdDataList.Cols - 1
'                   If j > mFixs Or mFixs = 0 Then
'                       grdDataList.col = j
'                       grdDataList.CellBackColor = QBColor(15)
'                   End If
'               Next j
'                grdDataList.col = 2
'                If Not IsNull(grdDataList.Text) Then
'                   StrTag = grdDataList.Text
'                   Screen.MousePointer = vbHourglass
'                   Call frm100123_3.SetParent(Me, StrTag, "A")
'                   frm100123_3.Show
'                   Screen.MousePointer = vbDefault
'                   Me.Enabled = True
'                   Exit Sub
'                End If
'            End If
            'end 2021/05/28
            If Trim(grdDataList.Text) = "V" Then
                strExc(1) = strExc(1) & "," & Format(i, "000")
                strExc(2) = strExc(2) & "," & Trim("" & grdDataList.TextMatrix(i, 2)) '總收文號
                strExc(3) = strExc(3) & "," & "A"
            End If
        Next i
        If strExc(1) <> "" Then
            strExc(1) = Mid(strExc(1), 2)
            strExc(2) = Mid(strExc(2), 2)
            strExc(3) = Mid(strExc(3), 2)
            Call frm100123_3.SetParent(Me, strExc(1), strExc(2), strExc(3))
            frm100123_3.Show
            Me.Hide
            Screen.MousePointer = vbDefault
        End If
        Me.Enabled = True
   'Added by Lydia 2025/09/10
   Case 18 '國外部行事曆
        If CheckUse("frm060209", strExec) = True Then
           'Added by Lydia 2025/09/10 傳入本所案號
           If PUB_CheckFormExist("frm060209") Then
              MsgBox "請先關閉〔行事曆提醒通知〕！", vbCritical + vbOKOnly
              Exit Sub
           End If
           Me.Enabled = False
           Screen.MousePointer = vbHourglass
           Call frm060209.SetParent(Me, strCP01 & strCP02 & strCP03 & strCP04)
           frm060209.Show
           Me.Hide
           Screen.MousePointer = vbDefault
           Me.Enabled = True
        End If
   Case Else
   End Select
End Sub

Private Sub cmdCustAtt_Click()
   Dim ii As Integer, jj As Integer
   Dim strCP09 As String, strCP10 As String
   Dim stFileName As String, stFileDescs As String
   Dim stSavePath As String
   
   Screen.MousePointer = vbHourglass
   
   stSavePath = App.path & "\CustLetter"
   If Dir(stSavePath, vbDirectory) = "" Then
      MkDir stSavePath
   Else
      PUB_KillAttach stSavePath
   End If
   
   With grdDataList
   For ii = 1 To .Rows - 1
      If Trim(.TextMatrix(ii, 0)) = "V" Then
         Screen.MousePointer = vbHourglass
         strCP09 = Pub_RplStr(.TextMatrix(ii, 2))
         If strCP09 <> "" Then
            strExc(0) = "select cp10 from letterprogress,caseprogress where lp01='" & strCP09 & "' and lp10='Y' and cp09(+)=lp01"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strCP10 = RsTemp("cp10")
               If PUB_GetAttachFile4Cust(strCP09, stFileName, stSavePath, True, strCP10, stFileDescs) = True Then
                  frm100101_2_1.stFileName = stFileName
                  frm100101_2_1.stFileDescs = stFileDescs
                  frm100101_2_1.stSavePath = stSavePath
                  Screen.MousePointer = vbDefault
                  frm100101_2_1.Show vbModal
               'Modified by Morgan 2020/9/21
               'Else
                  'MsgBox "客戶附件下載失敗！", vbCritical
               ElseIf stFileName = "" Then
                  MsgBox "無客戶附件！", vbExclamation
               'end 2020/9/21
                  'end 2020/9/21
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            Else
               MsgBox "本程序並無客戶函電子化記錄！", vbExclamation
            End If
         End If
         '反白還原
         .row = ii
         .TextMatrix(ii, 0) = ""
         For jj = 0 To .Cols - 1
             If jj >= .FixedCols Then
                 .col = jj
                 .CellBackColor = QBColor(15)
             End If
         Next jj
      End If
   Next ii
   End With
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim nFrm As Form
   
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
   Exit Sub
End Sub

Sub StrMenu()
Dim strSql  As String
Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String
Dim i As Integer, j As Integer
'Add By Cheng 2002/07/08
Dim StrSQLa As String
'add by Toni 20080926 控制跨部門權限訊息
Dim strTit As String
Dim strMsg As String
Dim nResponse
'end by Toni 20080926
Dim rsA As New ADODB.Recordset   '2009/8/19 ADD BY SONIA
Dim strWhereSql As String 'Add By Sindy 2021/4/21
Dim pYYMM As String 'Added by Lydia 2022/09/06 限制收文年月

Me.Enabled = False
StrSQLa = "": strWhereSql = ""
Str01 = ""
Str02 = ""
Str03 = ""
Str04 = ""
Label3.Caption = frm100101_2.Tag
If Left(Me.Tag, 1) = "N" Then
   strSql = Right(Me.Tag, Len(Me.Tag) - 1)
Else
   strSql = Me.Tag
End If
'Modify By Sindy 2025/7/29 + PUB_CheckQL05
'Modify By Sindy 2025/8/7 frm100101_2之StrMenu寫入QueryLog時，不必串接前畫面之查詢條件，只要留本所案號：CFT-XXXXXX(所有進度)
'pub_QL05 = pub_QL05 & ";本所案號：" & strSql 'Add By Sindy 2010/11/16
pub_QL05 = ";本所案號：" & strSql & "(所有進度)"
'2025/8/7 END
Str01 = SystemNumber(strSql, 1)
Str02 = SystemNumber(strSql, 2)
Str03 = SystemNumber(strSql, 3)
Str04 = SystemNumber(strSql, 4)

' 使用者沒有權限
'add by Toni 20080926 控制共同查詢是否有跨部門查詢案件明細權限
'2008/10/2 modify by sonia
'If IsUserHasRightOfSystem(strUserNum, Str01) = False Then
'   If IsUserHasRightOfFunction("frm100101_1", strCrossDept, False) = False Then
'      strTit = "檢核資料"
'      strMsg = "您沒有使用該系統類別的權限"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      tmpBol = fnCancelNowFormAndShowParentForm(Me)
'      Exit Sub
'   End If
'End If
If CheckSR09(strUserNum, Str01, "Y", , Str01, Str02, Str03, Str04) = False Then
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
'2008/10/2 end
'End 20080926

'Added by Lydia 2019/05/30 依國別抓案件性質欄位
Call ClsPDCheckCaseCodeIsExist(Str01, Str02, Str03, Str04, , , , , strPA09)
If strPA09 <= "010" Then
    strCon = "CPM03"
Else
    strCon = "CPM04"
End If
'end 2019/05/30

'Modify By Sindy 2021/4/21 不要為了排除C類而寫2個函數 StrMenu,StrMenu1(Mark)
If Me.m_CKind = "N" Then
   strWhereSql = " AND CP09<'C'"
End If
'2021/4/21 END

'Added by Lydia 2022/09/06 當以本所案號查詢時，若條件為TT-999999則點進度時再增加輸入收文年月對話框(預設當月，取消表示全部)以減少等待時間。
    pYYMM = ""
    'Modified by Lydia 2023/01/11 +LA-999999
    'If Str01 & Str02 = "TT999999" Then
    If InStr("TT999999,LA999999", Str01 & Str02) > 0 Then
JumpToReInput:
        pYYMM = InputBox("請輸入收文年月以減少等待時間，不輸入年月或按取消表示查詢全部資料。", "輸入收文年月", Left(strSrvDate(2), 5))
        If pYYMM <> "" Then
           If Val(Left(pYYMM, 5)) > Val(Left(strSrvDate(2), 5)) Then
               MsgBox "收文年月不可大於系統年月！", vbInformation
               GoTo JumpToReInput
           End If
           strWhereSql = strWhereSql & " AND CP05>=" & Val(pYYMM & "01") + 19110000 & " AND CP05<=" & Val(pYYMM & "31") + 19110000
           Me.Caption = Me.Caption & "-收文年月：" & pYYMM
        End If
    End If
'end 2022/09/06

'Add By Sindy 2018/12/14 檢查是否有註冊費的繳費單
cmdOK(6).Tag = Me.Caption
Me.Caption = Me.Caption & " 連線至" & strTFeeForm
cmdOK(16).Visible = False
'Modify By Sindy 2018/12/18 增加判斷智權部不顯示
If (Str01 = "T" Or Str01 = "FCT") And Left(Pub_StrUserSt03, 1) <> "S" Then
   If PUB_PrintTFeeForm(Str01, Str02, Str03, Str04, , , False) = True Then
      cmdOK(16).Visible = True
   End If
End If
Me.Caption = cmdOK(6).Tag
'2018/12/14 END

'Added by Lydia 2025/09/10 國外部行事曆
If (Str01 = "FCP" Or Str01 = "P") And (Pub_StrUserSt03 = "M51" Or Left(Pub_StrUserSt03, 1) = "F") Then
   cmdOK(18).Visible = True
Else
   cmdOK(18).Visible = False
End If
'end 2025/09/10

'Add By Sindy 2013/7/1
'Modify By Sindy 2015/10/21 + And Str01 <> "PS" And Str01 <> "CPS"
'Modify By Sindy 2018/7/24 + Not (Left(Str01, 1) = "T" Or Str01 = "FCT")
'Modify By Sindy 2021/11/4 + Str01 <> "ACS"
'Modify By Sindy 2023/9/28 mark
'If (Str01 <> "P" And Str01 <> "CFP" And Str01 <> "PS" And Str01 <> "CPS") And _
'   Pub_StrUserSt03 <> "M51" And _
'   Not (Left(Str01, 1) = "T" Or Str01 = "FCT") And _
'   Str01 <> "ACS" Then
'   'cmdOK(13).Visible = False 'Modify By Sindy 2015/3/9 Mark:不分系統別都可以查看
'   'cmdOK(14).Visible = False 'Modify By Sindy 2017/8/16 Mark
'   cmdOK(15).Visible = False
'End If
'2013/7/1 End

'2010/7/30 CANCEL BY SONIA 因內外商欲合併,故取消此控制
''2009/8/19 add by sonia FCT無爭議程序之案件內商人員不可查詢(該案有內商承辦人者為FCT爭議案)
'If Str01 = "FCT" And Mid(PUB_GetST03(strUserNum), 1, 2) = "P2" Then
'   StrSQLa = "Select * From CASEPROGRESS,STAFF Where CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' AND CP14=ST01(+) AND SUBSTR(ST03,1,2)='P2' "
'   rsA.CursorLocation = adUseClient
'   rsA.Open StrSQLa, cnnConnection, adOpenStatic
'   If rsA.RecordCount = 0 Then
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'      strTit = "檢核資料"
'      strMsg = "非FCT爭議案，您沒有使用該案號資料的權限"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      tmpBol = fnCancelNowFormAndShowParentForm(Me)
'      Exit Sub
'   Else
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'   End If
'End If
''2009/8/19 END
''2009/9/8 add by sonia T非台灣案非外商收文之案件,外商人員不可查詢
'If Str01 = "T" And Mid(PUB_GetST03(strUserNum), 1, 2) = "F1" Then
'   StrSQLa = "Select * From TRADEMARK Where TM01='" & Str01 & "' AND TM02='" & Str02 & "' AND TM03='" & Str03 & "' AND TM04='" & Str04 & "' "
'   rsA.CursorLocation = adUseClient
'   rsA.Open StrSQLa, cnnConnection, adOpenStatic
'   If rsA.RecordCount = 0 Then
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'      strTit = "檢核資料"
'      strMsg = "無此商標資料"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      tmpBol = fnCancelNowFormAndShowParentForm(Me)
'      Exit Sub
'   Else
'      If rsA.Fields("TM10") <> "000" Then '非台灣案才要控管外商人員權限
'         If rsA.State <> adStateClosed Then rsA.Close
'         Set rsA = Nothing
'         StrSQLa = "Select * From CASEPROGRESS Where CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' AND SUBSTR(CP12,1,2)='F1' "
'         rsA.CursorLocation = adUseClient
'         rsA.Open StrSQLa, cnnConnection, adOpenStatic
'         If rsA.RecordCount = 0 Then
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'            strTit = "檢核資料"
'            strMsg = "非外商收文之大陸商標案，您沒有使用該案號資料的權限"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            tmpBol = fnCancelNowFormAndShowParentForm(Me)
'            Exit Sub
'         Else
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'         End If
'      Else
'         If rsA.State <> adStateClosed Then rsA.Close
'         Set rsA = Nothing
'      End If
'   End If
'End If
''2009/9/8 END
''2010/1/22 add by sonia TM非外商收文之案件,外商人員不可查詢
'If Str01 = "TM" And Mid(PUB_GetST03(strUserNum), 1, 2) = "F1" Then
'   StrSQLa = "Select * From SERVICEPRACTICE Where SP01='" & Str01 & "' AND SP02='" & Str02 & "' AND SP03='" & Str03 & "' AND SP04='" & Str04 & "' "
'   rsA.CursorLocation = adUseClient
'   rsA.Open StrSQLa, cnnConnection, adOpenStatic
'   If rsA.RecordCount = 0 Then
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'      strTit = "檢核資料"
'      strMsg = "無此監視系統資料"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      tmpBol = fnCancelNowFormAndShowParentForm(Me)
'      Exit Sub
'   Else
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'      StrSQLa = "Select * From CASEPROGRESS Where CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' AND SUBSTR(CP12,1,2)='F1' "
'      rsA.CursorLocation = adUseClient
'      rsA.Open StrSQLa, cnnConnection, adOpenStatic
'      If rsA.RecordCount = 0 Then
'         If rsA.State <> adStateClosed Then rsA.Close
'         Set rsA = Nothing
'         strTit = "檢核資料"
'         strMsg = "非外商收文之監視系統案，您沒有使用該案號資料的權限"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         tmpBol = fnCancelNowFormAndShowParentForm(Me)
'         Exit Sub
'      Else
'         If rsA.State <> adStateClosed Then rsA.Close
'         Set rsA = Nothing
'      End If
'   End If
'End If
''2010/1/22 END
'2010/7/30 END

'add by nickc 2006/07/10 下面各句加入 cp60 要檢查收款沒
'Modify By Sindy 2011/6/23 +CP16
'Modify by Morgan 2011/8/12 收據的收款情形改判斷CP79
'Add by Lydia 2014/11/14 進度檔結果欄後加入下一期限及下一期限日期
             '將下一程序的查詢結果視做虛擬表格與原search做left join, strNpSqlOfNoSalesDuty剔除下一程序為程序管控之案件性質語法
Dim midSql As String
'Modified by Lydia 2015/01/15 同一(最小)所限可能會有兩筆下一程序,取最大序號
'midSql = "(select npf.np01,npf.np07,cpmf.cpm03 as 下一程序性質 ,npf.np08 from nextprogress npf,casepropertymap cpmf " & _
     " where npf.np02='" & Str01 & "' AND npf.np03='" & Str02 & "' AND npf.np04='" & Str03 & "' AND npf.np05='" & Str04 & "' " & _
     " and npf.np02=cpmf.cpm01(+) and npf.np07=cpmf.cpm02(+) " & strNpSqlOfNoSalesDuty & _
     " and np08 in (select min(n2f.np08) D_date from nextprogress n2f " & _
     " where n2f.np02='" & Str01 & "' AND n2f.np03='" & Str02 & "' AND n2f.np04='" & Str03 & "' AND n2f.np05='" & Str04 & "' " & strNpSqlOfNoSalesDuty & _
     " group by n2f.np01) group by npf.np01,npf.np07,cpmf.cpm03,npf.np08 ) nm " 'npf.group=同一性質同一所限可能有多筆(明細不同)
'Modified by Lydia 2019/05/30 重整語法:先抓同一收文號的最小所限(有排除特定性質,ex.CFP-014272的C92000488的公開999排除)
                                     '再抓符合最小所限的最大np22 ; (起因有人反應FCT-042306只出現一筆下一程序性質)
'midSql = "(select * from (select (2000000000-npf.np22) dnp22,npf.np01,npf.np07,cpmf.cpm03 as 下一程序性質 ,npf.np08 from nextprogress npf,casepropertymap cpmf " & _
     " where npf.np02='" & Str01 & "' AND npf.np03='" & Str02 & "' AND npf.np04='" & Str03 & "' AND npf.np05='" & Str04 & "' " & _
     " and npf.np02=cpmf.cpm01(+) and npf.np07=cpmf.cpm02(+) " & strNpSqlOfNoSalesDuty & _
     " and np08 in (select min(n2f.np08) D_date from nextprogress n2f " & _
     " where n2f.np02='" & Str01 & "' AND n2f.np03='" & Str02 & "' AND n2f.np04='" & Str03 & "' AND n2f.np05='" & Str04 & "' " & strNpSqlOfNoSalesDuty & _
     " group by n2f.np01) group by (2000000000-npf.np22),npf.np01,npf.np07,cpmf.cpm03,npf.np08 ) d1 where rownum=1) nm "
'Modified by Lydia 2019/06/10 參考CFP-030994的主張國際優先權
'midSql = "(select n2.np01,n2.np08,max(n2.np22) np22,c2." & strCon & " as 下一程序性質 " & _
              "from nextprogress n2,casepropertymap c2 " & _
              "where (n2.np01,n2.np08) in ( " & _
                    "select np01,min(np08) as mdate from nextprogress where np02='" & Str01 & "' AND np03='" & Str02 & "' AND np04='" & Str03 & "' AND np05='" & Str04 & "' " & _
                         strNpSqlOfNoSalesDuty & " group by np01) " & _
              "and n2.np02=c2.cpm01(+) and n2.np07=c2.cpm02(+) " & Replace(strNpSqlOfNoSalesDuty, "np", "n2.np") & _
              "group by n2.np01,n2.np08,c2." & strCon & ") nm"
'Modify By Sindy 2021/4/29 + ,np23
midSql = "(select np01,np08,np22,np23," & strCon & " as 下一程序性質 from nextprogress,casepropertymap " & _
              "where (np01,np08,np22) in (select n2.np01,n2.np08,max(n2.np22) np22 " & _
              "from nextprogress n2  " & _
              "where (n2.np01,n2.np08) in ( " & _
                    "select np01,min(np08) as mdate from nextprogress where np02='" & Str01 & "' AND np03='" & Str02 & "' AND np04='" & Str03 & "' AND np05='" & Str04 & "' " & _
                         strNpSqlOfNoSalesDuty & " group by np01) " & Replace(strNpSqlOfNoSalesDuty, "np", "n2.np") & _
              "group by n2.np01,n2.np08) and np02=cpm01(+) and np07=cpm02(+)) nm "

'Modified by Morgan 2024/7/10 +承辦人部門(s1.ST03 pdept), 控制是否排除程序用
Select Case Pub_RplStr(Str01)
Case "CFP", "FCP", "P"   '專利
      '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
      StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
      'Modified by Lydia 2014/11/14 進度檔結果欄後加入下一期限及下一期限日期
'      strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,PA73||'-'||PA72 as 繳費年度,PA74 as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, patent.PA05,PATENT.PA06,PATENT.PA07,PATENT.PA11,PATENT.PA09,PA57,PA22,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60,pa47," & IIf(pub_strUserOffice = "1", "pa108", "pa136") & " as IsCancel,cp16 " & _
'               " From caseprogress,staff s1,staff s2,patent,CUSTOMER,CASEPROPERTYMAP,FAGENT,SYSTEMKIND " & _
'               " Where caseprogress.CP01=patent.PA01(+) and caseprogress.CP02=patent.PA02(+) and caseprogress.CP03=patent.PA03(+) and caseprogress.CP04=patent.PA04(+) and CP01='" & Str01 & "' and CP02='" & Str02 & "' and CP03='" & Str03 & "' and CP04='" & Str04 & "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1) = FA02(+) AND CP01=SK01(+) "
      'strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,PA73||'-'||PA72 as 繳費年度,PA74 as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, patent.PA05,PATENT.PA06,PATENT.PA07,PATENT.PA11,PATENT.PA09,PA57,PA22,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60,pa47," & IIf(pub_strUserOffice = "1", "pa108", "pa136") & " as IsCancel,cp16 " & _
               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff s1,staff s2,patent,CUSTOMER,CASEPROPERTYMAP,FAGENT,SYSTEMKIND, " & _
               midSql & " Where caseprogress.CP01=patent.PA01(+) and caseprogress.CP02=patent.PA02(+) and caseprogress.CP03=patent.PA03(+) and caseprogress.CP04=patent.PA04(+) and CP01='" & Str01 & "' and CP02='" & Str02 & "' and CP03='" & Str03 & "' and CP04='" & Str04 & "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1) = FA02(+) AND CP01=SK01(+) " & _
               " and  caseprogress.CP09=nm.np01(+) "
      'Modified by Lydia 2015/02/11 專利權消滅(1604)的案件性質改顯示CP64
      'modify by sonia 2015/10/30 非台灣案CP64之專利權消滅改消滅
      'strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,decode(cp01||cp10,'P1604',decode(sign(instr(cp64,'專利權消滅')),1,substr(cp64,instr(cp64,'專利權消滅'),15),NVL(DECODE(PA09,'000',CPM03,CPM04),CP10)),'CFP1604',decode(sign(instr(cp64,'專利權消滅')),1,substr(cp64,instr(cp64,'專利權消滅'),15),NVL(DECODE(PA09,'000',CPM03,CPM04),CP10)),NVL(DECODE(PA09,'000',CPM03,CPM04),CP10))  as 案件性質," & _
               "NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,PA73||'-'||PA72 as 繳費年度,PA74 as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, patent.PA05,PATENT.PA06,PATENT.PA07,PATENT.PA11,PATENT.PA09,PA57,PA22,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60,pa47," & IIf(pub_strUserOffice = "1", "pa108", "pa136") & " as IsCancel,cp16 " & _
               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff s1,staff s2,patent,CUSTOMER,CASEPROPERTYMAP,FAGENT,SYSTEMKIND, " & _
               midSql & " Where caseprogress.CP01=patent.PA01(+) and caseprogress.CP02=patent.PA02(+) and caseprogress.CP03=patent.PA03(+) and caseprogress.CP04=patent.PA04(+) and CP01='" & Str01 & "' and CP02='" & Str02 & "' and CP03='" & Str03 & "' and CP04='" & Str04 & "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1) = FA02(+) AND CP01=SK01(+) " & _
               " and  caseprogress.CP09=nm.np01(+) "
      'modify by sonia 2019/5/8 收款情形加銷帳CFP-025274及退費T-089878
      'Modify By Sindy 2021/4/29 外專部門加看約定期限(取代"取消收文日"的位置,取消收文日改放在結果之後)
      ' SQLDATET2(CP57) as 取消收文日 => decode('" & Left(Pub_StrUserSt03, 2) & "','F2',SQLDATET2(NP23),SQLDATET2(CP57)) as " & IIf(Left(Pub_StrUserSt03, 2)="F2", "約定期限", "取消收文日")
      'Modified by Lydia 2021/05/24 限制欄位長度
      'strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,decode(cp01||cp10,'P1604',decode(sign(instr(cp64,'專利權消滅')),1,substr(cp64,instr(cp64,'專利權消滅'),15),decode(sign(instr(cp64,'消滅')),1,substr(cp64,instr(cp64,'消滅'),12),NVL(DECODE(PA09,'000',CPM03,CPM04),CP10))),'CFP1604',decode(sign(instr(cp64,'消滅')),1,substr(cp64,instr(cp64,'消滅'),12),NVL(DECODE(PA09,'000',CPM03,CPM04),CP10)),NVL(DECODE(PA09,'000',CPM03,CPM04),CP10)) as 案件性質," & _
               "NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,decode('" & Left(Pub_StrUserSt03, 2) & "','F2',SQLDATET2(NP23),SQLDATET2(CP57)) as " & IIf(Left(Pub_StrUserSt03, 2) = "F2", "約定期限", "取消收文日") & "," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,SQLDATET2(CP57) as 取消收文日,PA73||'-'||PA72 as 繳費年度" & _
               ",PA74 as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, patent.PA05,PATENT.PA06,PATENT.PA07,PATENT.PA11,PATENT.PA09,PA57,PA22,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,pa47," & IIf(pub_strUserOffice = "1", "pa108", "pa136") & " as IsCancel,cp16 " & _
               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限,cp10,pa08 From caseprogress,staff s1,staff s2,patent,CUSTOMER,CASEPROPERTYMAP,FAGENT,SYSTEMKIND, " & _
               midSql & " Where caseprogress.CP01=patent.PA01(+) and caseprogress.CP02=patent.PA02(+) and caseprogress.CP03=patent.PA03(+) and caseprogress.CP04=patent.PA04(+) and CP01='" & Str01 & "' and CP02='" & Str02 & "' and CP03='" & Str03 & "' and CP04='" & Str04 & "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1) = FA02(+) AND CP01=SK01(+)" & _
               " and caseprogress.CP09=nm.np01(+)" & strWhereSql
      'end 2015/10/30
      'modify by sonia 2024/8/2 實際結果加'3'部分勝
      strSql = "select ' ' AS V,substr(SQLDATET2(CP05),1,10) as 收文日,CP09 as 總收文號,decode(cp01||cp10,'P1604',decode(sign(instr(cp64,'專利權消滅')),1,substr(cp64,instr(cp64,'專利權消滅'),15),decode(sign(instr(cp64,'消滅')),1,substr(cp64,instr(cp64,'消滅'),12),NVL(DECODE(PA09,'000',CPM03,CPM04),CP10))),'CFP1604',decode(sign(instr(cp64,'消滅')),1,substr(cp64,instr(cp64,'消滅'),12),NVL(DECODE(PA09,'000',CPM03,CPM04),CP10)),NVL(DECODE(PA09,'000',CPM03,CPM04),CP10)) as 案件性質," & _
               "NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員,substr(SQLDATET2(CP06),1,10) as 本所期限,substr(SQLDATET2(CP07),1,10) as 法定期限,substr(SQLDATET2(CP27),1,10) as 發文日," & IIf((Left(Pub_StrUserSt03, 2) = "F2" Or Pub_StrUserSt03 = "M51"), "decode(CP10,'605','','601','',SQLDATET2(NP23)) as 約定期限", "SQLDATET2(CP57) as 取消收文日") & "," & StrSQLa & "decode(CP24,'1','准','2','駁','3','部分勝',CP24) as 結果,substr(SQLDATET2(CP57),1,10) as 取消收文日CP57,PA73||'-'||PA72 as 繳費年度" & _
               ",PA74 as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人,CP22 as 是否出名,substr(CP64,1,500) as 進度備註, patent.PA05,PATENT.PA06,PATENT.PA07,PATENT.PA11,PATENT.PA09,PA57,PA22,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,pa47," & IIf(pub_strUserOffice = "1", "pa108", "pa136") & " as IsCancel,cp16 " & _
               " ,nm.下一程序性質 ,substr(SQLDATET2(nm.np08),1,10) as 下一程序所限,cp10,pa08,s1.ST03 pdept From caseprogress,staff s1,staff s2,patent,CUSTOMER,CASEPROPERTYMAP,FAGENT,SYSTEMKIND, " & _
               midSql & " Where caseprogress.CP01=patent.PA01(+) and caseprogress.CP02=patent.PA02(+) and caseprogress.CP03=patent.PA03(+) and caseprogress.CP04=patent.PA04(+) and CP01='" & Str01 & "' and CP02='" & Str02 & "' and CP03='" & Str03 & "' and CP04='" & Str04 & "' " & _
               "AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1) = FA02(+) AND CP01=SK01(+)" & _
               " and caseprogress.CP09=nm.np01(+)" & strWhereSql
Case "CFT", "FCT", "T", "TF"   '商標
'是否出名與代理人對調
      '引進是否閉卷欄
      '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
      StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
      'Modified by Lydia 2014/11/14 進度檔結果欄後加入下一期限及下一期限日期
'      strSql = "Select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,NVL(DECODE(tm10,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,' ' as 繳費年度,' ' as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, trademark.TM05,TRADEMARK.TM06,TRADEMARK.TM07,TRADEMARK.TM12,TRADEMARK.TM10,TM29,TM15,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60,tm34," & IIf(pub_strUserOffice = "1", "tm57", "tm73") & " as IsCancel,cp16 " & _
'               " From caseprogress,staff S1,STAFF S2,TradeMark,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind " & _
'               " Where caseprogress.CP01=TRADEMARK.TM01(+) and caseprogress.CP02=TRADEMARK.TM02(+) and caseprogress.CP03=TRADEMARK.TM03(+) and caseprogress.CP04=TRADEMARK.TM04(+) AND CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
'      strSql = strSql + " AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) "
      'modify by sonia 2019/5/8 收款情形加銷帳CFP-025274及退費T-089878
      'Modify By Sindy 2021/4/29 外專部門加看約定期限(取代"取消收文日"的位置,取消收文日改放在結果之後)
      'Modified by Lydia 2021/05/24 限制欄位長度
'      strSql = "Select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,NVL(DECODE(tm10,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,SQLDATET2(CP57) as 取消收文日,' ' as 繳費年度" & _
               ",' ' as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, trademark.TM05,TRADEMARK.TM06,TRADEMARK.TM07,TRADEMARK.TM12,TRADEMARK.TM10,TM29,TM15,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,tm34," & IIf(pub_strUserOffice = "1", "tm57", "tm73") & " as IsCancel,cp16 " & _
               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,TradeMark,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind, " & _
               midSql & " Where caseprogress.CP01=TRADEMARK.TM01(+) and caseprogress.CP02=TRADEMARK.TM02(+) and caseprogress.CP03=TRADEMARK.TM03(+) and caseprogress.CP04=TRADEMARK.TM04(+) AND CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
      'modify by sonia 2024/8/2 實際結果加'3'部分勝
      strSql = "Select ' ' AS V,substr(SQLDATET2(CP05),1,10) as 收文日,CP09 as 總收文號,NVL(DECODE(tm10,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員," & _
               "substr(SQLDATET2(CP06),1,10) as 本所期限,substr(SQLDATET2(CP07),1,10) as 法定期限,substr(SQLDATET2(CP27),1,10) as 發文日,substr(SQLDATET2(CP57),1,10) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁','3','部分勝',CP24) as 結果,substr(SQLDATET2(CP57),1,10) as 取消收文日CP57,' ' as 繳費年度" & _
               ",' ' as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人,CP22 as 是否出名,substr(CP64,1,500) as 進度備註, trademark.TM05,TRADEMARK.TM06,TRADEMARK.TM07,TRADEMARK.TM12,TRADEMARK.TM10,TM29,TM15,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,tm34," & IIf(pub_strUserOffice = "1", "tm57", "tm73") & " as IsCancel,cp16 " & _
               " ,nm.下一程序性質 ,substr(SQLDATET2(nm.np08),1,10) as 下一程序所限,s1.ST03 pdept From caseprogress,staff S1,STAFF S2,TradeMark,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind, " & _
               midSql & " Where caseprogress.CP01=TRADEMARK.TM01(+) and caseprogress.CP02=TRADEMARK.TM02(+) and caseprogress.CP03=TRADEMARK.TM03(+) and caseprogress.CP04=TRADEMARK.TM04(+) AND CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' " & _
               "AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
      strSql = strSql + " and caseprogress.CP09=nm.np01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+)" & _
               strWhereSql
'Add by Amy 2020/11/13 獨立出來,抓資料的方式同專利/商標
Case "ACS"
      StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
      'Modify By Sindy 2021/4/29 外專部門加看約定期限(取代"取消收文日"的位置,取消收文日改放在結果之後)
      'Modified by Lydia 2021/05/24 限制欄位長度
'      strSql = "Select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,NVL(DECODE(LC15,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,SQLDATET2(CP57) as 取消收文日,' ' as 繳費年度" & _
               ",' ' as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, lawcase.LC05,LAWCASE.LC06,LAWCASE.LC07,'' as vt22,LAWCASE.LC15,LC08,'' as vt25,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,lc16," & IIf(pub_strUserOffice = "1", "lc34", "lc36") & " as IsCancel,cp16 " & _
               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,LawCase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind, " & _
               midSql & " Where caseprogress.CP01=LawCase.LC01(+) and caseprogress.CP02=LawCase.LC02(+) and caseprogress.CP03=LawCase.LC03(+) and caseprogress.CP04=LawCase.LC04(+) AND CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
      'modify by sonia 2024/8/2 實際結果加'3'部分勝
      strSql = "Select ' ' AS V,substr(SQLDATET2(CP05),1,10) as 收文日,CP09 as 總收文號,NVL(DECODE(LC15,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員," & _
               "substr(SQLDATET2(CP06),1,10) as 本所期限,substr(SQLDATET2(CP07),1,10) as 法定期限,substr(SQLDATET2(CP27),1,10) as 發文日,substr(SQLDATET2(CP57),1,10) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁','3','部分勝',CP24) as 結果,substr(SQLDATET2(CP57),1,10) as 取消收文日CP57,' ' as 繳費年度" & _
               ",' ' as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人,CP22 as 是否出名,substr(CP64,1,500) as 進度備註, lawcase.LC05,LAWCASE.LC06,LAWCASE.LC07,'' as vt22,LAWCASE.LC15,LC08,'' as vt25,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,lc16," & IIf(pub_strUserOffice = "1", "lc34", "lc36") & " as IsCancel,cp16 " & _
               " ,nm.下一程序性質 ,substr(SQLDATET2(nm.np08),1,10) as 下一程序所限,s1.ST03 pdept From caseprogress,staff S1,STAFF S2,LawCase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind, " & _
               midSql & " Where caseprogress.CP01=LawCase.LC01(+) and caseprogress.CP02=LawCase.LC02(+) and caseprogress.CP03=LawCase.LC03(+) and caseprogress.CP04=LawCase.LC04(+) AND CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' " & _
               "AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
      strSql = strSql + " and caseprogress.CP09=nm.np01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+)" & _
               strWhereSql
'edit by nickc 2005/10/06 法務及顧問更改格式
'2005/12/19 MODIFY BY SONIA 取消機關名稱欄改抓CF代理人,依立卷問題3需求調整欄位位置
'Modify By Sindy 2009/07/24 增加LIN系統類別
'2009/9/9 MODIFY BY SONIA 法務進度備註欄同時帶案件性質
'modify by sonia 2019/7/30 +ACS系統類別
'Modify by Amy 2020/11/13 ACS另外獨立寫,抓資料的方式同專利/商標
Case "CFL", "FCL", "L", "LIN"    '法務
'      '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
      StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
      'Modified by Lydia 2014/11/14 進度檔結果欄後加入下一期限及下一期限日期
'      strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')' as 進度備註,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦律師,decode(CP29,S3.ST01,S3.ST02) as 承辦法務,NVL(S2.ST02,CP13) as 智權人員,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,'' as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & " '',lawcase.LC05,LAWCASE.LC06,LAWCASE.LC07,'',LAWCASE.LC15,LC08,'',decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60,lc16," & IIf(pub_strUserOffice = "1", "lc34", "lc36") & " as IsCancel,cp16 " & _
'               " From caseprogress,staff S1,STAFF S2,staff s3,lawcase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization " & _
'               " Where caseprogress.CP01=lawcase.LC01(+) and caseprogress.CP02=lawcase.LC02(+) and caseprogress.CP03=lawcase.LC03(+) and caseprogress.CP04=lawcase.LC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
'      strSql = strSql + " AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) and cp71=or01(+) "
      'Modified by Lydia 2015/10/05
      'Modified by Lydia 2016/05/30 備註顯示回執退件日/回執未回郵局送達日
      'strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')' as 進度備註,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,decode(CP29,S3.ST01,S3.ST02) as 協辦人員,NVL(S2.ST02,CP13) as 智權人員,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,'' as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & " '',lawcase.LC05,LAWCASE.LC06,LAWCASE.LC07,'',LAWCASE.LC15,LC08,'',decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60,lc16," & IIf(pub_strUserOffice = "1", "lc34", "lc36") & " as IsCancel,cp16 " & _
               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,staff s3,lawcase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization, " & _
               midSql & " Where caseprogress.CP01=lawcase.LC01(+) and caseprogress.CP02=lawcase.LC02(+) and caseprogress.CP03=lawcase.LC03(+) and caseprogress.CP04=lawcase.LC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
      'Modify by Amy 2017/12/26 +串ep07 原:'' as 會稿日
      'modify by sonia 2019/5/8 收款情形加銷帳CFP-025274及退費T-089878
      'Modified by Lydia 2020/07/15 法律所案源收文：增加案源之介紹人
'      strSql = " select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號," & _
'               "decode(sign(instr('3,4',sk02)),1,decode(cp46,19221111,'回執退件日:'||sqldatet(cp47)||';',19220101,'回執未回郵局送達日:'||sqldatet(cp47)||';',''),'')||CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')' as 進度備註," & _
'               "NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,decode(CP29,S3.ST01,S3.ST02) as 協辦人員,NVL(S2.ST02,CP13) as 智權人員,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(ep07) as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & " '',lawcase.LC05,LAWCASE.LC06,LAWCASE.LC07,'',LAWCASE.LC15,LC08,'',decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,lc16," & IIf(pub_strUserOffice = "1", "lc34", "lc36") & " as IsCancel,cp16 " & _
'               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,staff s3,lawcase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization,EngineerProgress, " & _
'               midSql & " Where caseprogress.CP01=lawcase.LC01(+) and caseprogress.CP02=lawcase.LC02(+) and caseprogress.CP03=lawcase.LC03(+) and caseprogress.CP04=lawcase.LC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
'      strSql = strSql + " and caseprogress.CP09=nm.np01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) and cp71=or01(+) And caseprogress.CP09=ep02(+) "
'      'end 2017/12/26
      'Modified by Lydia 2021/05/24 限制欄位長度
      'strSql = " select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號," & _
               "decode(sign(instr('3,4',sk02)),1,decode(cp46,19221111,'回執退件日:'||sqldatet(cp47)||';',19220101,'回執未回郵局送達日:'||sqldatet(cp47)||';',''),'')||CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')' as 進度備註," & _
               "NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,decode(CP29,S3.ST01,S3.ST02) as 協辦人員,NVL(S2.ST02,CP13) as 智權人員,DECODE(LOS15,NULL,NULL,GETSTAFFNAMELIST(LOS04)) 介紹人,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(ep07) as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & _
               StrSQLa & "'' as 暫無, lawcase.LC05,LAWCASE.LC06,LAWCASE.LC07,'' as vt22,LAWCASE.LC15,LC08,'' as vt25,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,lc16," & IIf(pub_strUserOffice = "1", "lc34", "lc36") & " as IsCancel,cp16 " & _
               ",nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,staff s3,lawcase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization,EngineerProgress,LawOfficeSource, " & _
               midSql & " Where caseprogress.CP01=lawcase.LC01(+) and caseprogress.CP02=lawcase.LC02(+) and caseprogress.CP03=lawcase.LC03(+) and caseprogress.CP04=lawcase.LC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
      'modify by sonia 2022/9/22 配合L-888888只能看同區收文資料改語法
      strSql = " select ' ' AS V,substr(SQLDATET2(CP05),1,10) as 收文日,CP09 as 總收文號," & _
               "substr(decode(sign(instr('3,4',sk02)),1,decode(cp46,19221111,'回執退件日:'||sqldatet(cp47)||';',19220101,'回執未回郵局送達日:'||sqldatet(cp47)||';',''),'')||CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',1,500) as 進度備註," & _
               "NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,decode(CP29,S3.ST01,S3.ST02) as 協辦人員,NVL(S2.ST02,CP13) as 智權人員,DECODE(LOS15,NULL,NULL,GETSTAFFNAMELIST(LOS04)) 介紹人," & _
               "cp18 as 點數,substr(SQLDATET2(CP06),1,10) as 本所期限,substr(SQLDATET2(CP07),1,10) as 法定期限,substr(SQLDATET2(ep07),1,10) as 會稿日,substr(SQLDATET2(CP27),1,10) as 發文日," & SQLDate("CP46") & " as 回執日,substr(SQLDATET2(CP57),1,10) as 取消收文日," & _
               StrSQLa & "'' as 暫無, lawcase.LC05,LAWCASE.LC06,LAWCASE.LC07,'' as vt22,LAWCASE.LC15,LC08,'' as vt25,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,lc16," & IIf(pub_strUserOffice = "1", "lc34", "lc36") & " as IsCancel,cp16 " & _
               ",nm.下一程序性質 ,substr(SQLDATET2(nm.np08),1,10) as 下一程序所限,s1.ST03 pdept From caseprogress,staff S1,STAFF S2,staff s3,staff s4,lawcase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,Acc090,organization,EngineerProgress,LawOfficeSource, " & _
               midSql & " Where caseprogress.CP01=lawcase.LC01(+) and caseprogress.CP02=lawcase.LC02(+) and caseprogress.CP03=lawcase.LC03(+) and caseprogress.CP04=lawcase.LC04(+) AND CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' " & _
               "AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) and cp29=s3.st01(+) and los04=S4.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
      strSql = strSql + " and caseprogress.CP09=nm.np01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) AND S4.ST15=A0901(+) and cp71=or01(+) And caseprogress.CP09=ep02(+) and caseprogress.CP162=LOS15(+)" & _
               strWhereSql
      'end 2020/07/15
      'add by sonia 2022/9/22
      '以操作人員之st15再抓ACC090之A0911，再以進度CP12再抓ACC090之A0911，若與操作人員帶出之A0911相同才可顯示進度。卷宗區也是。
      'st05 in ('00',’01’)人員不受上述限制,系統特殊設定「全所智權部主管」的人員也不限制
      If (InStr("00,01", Pub_strUserST05) = 0 And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) = 0) _
         And Str01 = "L" And Str02 = "888888" Then
         strSql = strSql + " AND A0911='" & Pub_StrUserSt15 & "' "
      End If
      'end 2022/9/22
'2009/9/9 MODIFY BY SONIA 顧問進度備註欄同時帶案件性質
Case "LA"                      '顧問
      StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
      'Modified by Lydia 2014/11/14 進度檔結果欄後加入下一期限及下一期限日期
'      strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,decode(cp10,'0',sqldatet(cp53)||'--'||sqldatet(cp54),CP64||'('||NVL(CPM03,CP10)||')') as 進度備註,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦律師,decode(CP29,S3.ST01,S3.ST02) as 承辦法務,NVL(S2.ST02,CP13) as 智權人員,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,'' as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "'',hirecase.HC06,'','','','',HC09,'',decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60,hc07," & IIf(pub_strUserOffice = "1", "hc19", "hc20") & " as IsCancel,cp16 " & _
'               " From caseprogress,staff S1,STAFF S2,staff s3,hirecase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization " & _
'               " Where caseprogress.CP01=hirecase.HC01(+) and caseprogress.CP02=hirecase.HC02(+) and caseprogress.CP03=hirecase.HC03(+) and caseprogress.CP04=hirecase.HC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
'      strSql = strSql + " AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) and cp71=or01(+) "
     'Modified by Lydia 2015/10/05
     'Modified by Lydia 2016/05/30 備註顯示回執退件日/回執未回郵局送達日
      'strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,decode(cp10,'0',sqldatet(cp53)||'--'||sqldatet(cp54),CP64||'('||NVL(CPM03,CP10)||')') as 進度備註,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,decode(CP29,S3.ST01,S3.ST02) as 協辦人員,NVL(S2.ST02,CP13) as 智權人員,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,'' as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "'',hirecase.HC06,'','','','',HC09,'',decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60,hc07," & IIf(pub_strUserOffice = "1", "hc19", "hc20") & " as IsCancel,cp16 " & _
               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,staff s3,hirecase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization, " & _
               midSql & " Where caseprogress.CP01=hirecase.HC01(+) and caseprogress.CP02=hirecase.HC02(+) and caseprogress.CP03=hirecase.HC03(+) and caseprogress.CP04=hirecase.HC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
      'Modify by Amy 2017/12/26 +串ep07 原:'' as 會稿日
      'modify by sonia 2019/5/8 收款情形加銷帳CFP-025274及退費T-089878
      'Modified by Lydia 2020/07/15 法律所案源收文：增加案源之介紹人
      'strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號," & _
               "decode(sign(instr('3,4',sk02)),1,decode(cp46,19221111,'回執退件日:'||sqldatet(cp47)||';',19220101,'回執未回郵局送達日:'||sqldatet(cp47)||';',''),'')||decode(cp10,'0',sqldatet(cp53)||'--'||sqldatet(cp54),CP64||'('||NVL(CPM03,CP10)||')') as 進度備註," & _
               "NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,decode(CP29,S3.ST01,S3.ST02) as 協辦人員,NVL(S2.ST02,CP13) as 智權人員,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(ep07)  as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "'',hirecase.HC06,'','','','',HC09,'',decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,hc07," & IIf(pub_strUserOffice = "1", "hc19", "hc20") & " as IsCancel,cp16 " & _
               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,staff s3,hirecase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization,EngineerProgress, " & _
               midSql & " Where caseprogress.CP01=hirecase.HC01(+) and caseprogress.CP02=hirecase.HC02(+) and caseprogress.CP03=hirecase.HC03(+) and caseprogress.CP04=hirecase.HC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
      'strSql = strSql + " and caseprogress.CP09=nm.np01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) and cp71=or01(+) And caseprogress.CP09=ep02(+) "
      'Modified by Lydia 2021/05/24 限制欄位長度
      'strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號," & _
               "decode(sign(instr('3,4',sk02)),1,decode(cp46,19221111,'回執退件日:'||sqldatet(cp47)||';',19220101,'回執未回郵局送達日:'||sqldatet(cp47)||';',''),'')||decode(cp10,'0',sqldatet(cp53)||'--'||sqldatet(cp54),CP64||'('||NVL(CPM03,CP10)||')') as 進度備註," & _
               "NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,decode(CP29,S3.ST01,S3.ST02) as 協辦人員,NVL(S2.ST02,CP13) as 智權人員,DECODE(LOS15,NULL,NULL,GETSTAFFNAMELIST(LOS04)) 介紹人,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(ep07)  as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & _
               StrSQLa & "'' as 暫無, hirecase.HC06,'' as vt20,'' as vt21,'' as vt22,'' as vt23,HC09,'' as vt25,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,hc07," & IIf(pub_strUserOffice = "1", "hc19", "hc20") & " as IsCancel,cp16 " & _
               ",nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,staff s3,hirecase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization,EngineerProgress,LawOfficeSource, " & _
               midSql & " Where caseprogress.CP01=hirecase.HC01(+) and caseprogress.CP02=hirecase.HC02(+) and caseprogress.CP03=hirecase.HC03(+) and caseprogress.CP04=hirecase.HC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
      strSql = "select ' ' AS V,substr(SQLDATET2(CP05),1,10) as 收文日,CP09 as 總收文號," & _
               "substr(decode(sign(instr('3,4',sk02)),1,decode(cp46,19221111,'回執退件日:'||sqldatet(cp47)||';',19220101,'回執未回郵局送達日:'||sqldatet(cp47)||';',''),'')||decode(cp10,'0',sqldatet(cp53)||'--'||sqldatet(cp54),CP64||'('||NVL(CPM03,CP10)||')'),1,50) as 進度備註," & _
               "NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,decode(CP29,S3.ST01,S3.ST02) as 協辦人員,NVL(S2.ST02,CP13) as 智權人員,DECODE(LOS15,NULL,NULL,GETSTAFFNAMELIST(LOS04)) 介紹人,cp18 as 點數," & _
               "substr(SQLDATET2(CP06),1,10) as 本所期限,substr(SQLDATET2(CP07),1,10) as 法定期限,substr(SQLDATET2(ep07),1,10)  as 會稿日,substr(SQLDATET2(CP27),1,10) as 發文日," & SQLDate("CP46") & " as 回執日,substr(SQLDATET2(CP57),1,10) as 取消收文日," & _
               StrSQLa & "'' as 暫無, hirecase.HC06,'' as vt20,'' as vt21,'' as vt22,'' as vt23,HC09,'' as vt25,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,hc07," & IIf(pub_strUserOffice = "1", "hc19", "hc20") & " as IsCancel,cp16 " & _
               ",nm.下一程序性質 ,substr(SQLDATET2(nm.np08),1,10) as 下一程序所限,s1.ST03 pdept From caseprogress,staff S1,STAFF S2,staff s3,hirecase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization,EngineerProgress,LawOfficeSource, " & _
               midSql & " Where caseprogress.CP01=hirecase.HC01(+) and caseprogress.CP02=hirecase.HC02(+) and caseprogress.CP03=hirecase.HC03(+) and caseprogress.CP04=hirecase.HC04(+) AND CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' " & _
               "AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
      strSql = strSql + " and caseprogress.CP09=nm.np01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) and cp71=or01(+) And caseprogress.CP09=ep02(+) and caseprogress.CP162=LOS15(+)" & _
               strWhereSql
      'end 2020/07/15
Case Else                  '服務
'是否出名與代理人對調
      '引進是否閉卷欄
      '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
      StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
      'Modified by Lydia 2014/11/14 進度檔結果欄後加入下一期限及下一期限日期
'      strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,NVL(DECODE(sP09,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,' ' as 繳費年度,' ' as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, servicepractice.SP05,SERVICEPRACTICE.SP06,SERVICEPRACTICE.SP07,SERVICEPRACTICE.SP11,SERVICEPRACTICE.SP09,SP15,SP14,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60,sp28," & IIf(pub_strUserOffice = "1", "sp61", "sp68") & " as IsCancel,cp16 " & _
'               " From caseprogress,staff S1,STAFF S2,servicepractice,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind " & _
'               " Where caseprogress.CP01=servicepractice.SP01(+) and caseprogress.CP02=servicepractice.SP02(+) and caseprogress.CP03=servicepractice.SP03(+) and caseprogress.CP04=servicepractice.SP04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
'      strSql = strSql + " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) "
      'Modified by Lydia 2015/10/05
      'modify by sonia 2019/5/8 收款情形加銷帳CFP-025274及退費T-089878
      'Modify By Sindy 2021/4/29 外專部門加看約定期限(取代"取消收文日"的位置,取消收文日改放在結果之後)
      ' SQLDATET2(CP57) as 取消收文日 => decode('" & Left(Pub_StrUserSt03, 2) & "','F2',SQLDATET2(NP23),SQLDATET2(CP57)) as " & IIf(Left(Pub_StrUserSt03, 2)="F2", "約定期限", "取消收文日")
      'Modified by Lydia 2021/05/24 限制欄位長度
      'strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,NVL(DECODE(sP09,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,decode('" & Left(Pub_StrUserSt03, 2) & "','F2',SQLDATET2(NP23),SQLDATET2(CP57)) as " & IIf(Left(Pub_StrUserSt03, 2) = "F2", "約定期限", "取消收文日") & "," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,SQLDATET2(CP57) as 取消收文日,' ' as 繳費年度" & _
               ",' ' as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, servicepractice.SP05,SERVICEPRACTICE.SP06,SERVICEPRACTICE.SP07,SERVICEPRACTICE.SP11,SERVICEPRACTICE.SP09,SP15,SP14,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,sp28," & IIf(pub_strUserOffice = "1", "sp61", "sp68") & " as IsCancel,cp16 " & _
               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,servicepractice,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,Acc090, " & _
               midSql & " Where caseprogress.CP01=servicepractice.SP01(+) and caseprogress.CP02=servicepractice.SP02(+) and caseprogress.CP03=servicepractice.SP03(+) and caseprogress.CP04=servicepractice.SP04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
      'modify by sonia 2024/8/2 實際結果加'3'部分勝
      strSql = "select ' ' AS V,substr(SQLDATET2(CP05),1,10) as 收文日,CP09 as 總收文號,NVL(DECODE(sP09,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員," & _
               "substr(SQLDATET2(CP06),1,10) as 本所期限,substr(SQLDATET2(CP07),1,10) as 法定期限,substr(SQLDATET2(CP27),1,10) as 發文日," & IIf((Left(Pub_StrUserSt03, 2) = "F2" Or Pub_StrUserSt03 = "M51"), "decode(CP10,'605','','601','',SQLDATET2(NP23)) as 約定期限", "SQLDATET2(CP57) as 取消收文日") & _
               "," & StrSQLa & "decode(CP24,'1','准','2','駁','3','部分勝',CP24) as 結果,substr(SQLDATET2(CP57),1,10) as 取消收文日CP57,' ' as 繳費年度" & _
               ",' ' as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人,CP22 as 是否出名," & _
               "substr(CP64,1,500) as 進度備註, servicepractice.SP05,SERVICEPRACTICE.SP06,SERVICEPRACTICE.SP07,SERVICEPRACTICE.SP11,SERVICEPRACTICE.SP09,SP15,SP14,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,sp28," & IIf(pub_strUserOffice = "1", "sp61", "sp68") & " as IsCancel,cp16 " & _
               " ,nm.下一程序性質 ,substr(SQLDATET2(nm.np08),1,10) as 下一程序所限,s1.ST03 pdept From caseprogress,staff S1,STAFF S2,servicepractice,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,Acc090, " & _
               midSql & " Where caseprogress.CP01=servicepractice.SP01(+) and caseprogress.CP02=servicepractice.SP02(+) and caseprogress.CP03=servicepractice.SP03(+) and caseprogress.CP04=servicepractice.SP04(+) AND CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' " & _
               "AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
      strSql = strSql + " and caseprogress.CP09=nm.np01(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) AND CP12=A0901(+)" & _
               strWhereSql
      'Add By Sindy 2020/5/12
      '以操作人員之st15再抓ACC090之A0911，再以進度CP12再抓ACC090之A0911，若與操作人員帶出之A0911相同才可顯示進度。卷宗區也是。
      'st05 in ('00',’01’)人員不受上述限制
      'Modify By Sindy 2022/5/23 再加入系統特殊設定「全所智權部主管」的人員也不限制。
      If (InStr("00,01", Pub_strUserST05) = 0 And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) = 0) _
         And Str01 = "TT" And Str02 = "999999" Then
         strSql = strSql + " AND A0911=(select a90.A0911 from acc090 a90 where a90.a0901='" & Pub_StrUserSt15 & "') "
      End If
      '2020/5/12 END
End Select
'2009/12/2 modify by sonia 同日C類來函同時產生B類收文,應收顯示C類再顯示B類 FCT-029810,故再加收文號A->C->B
'strSQL = strSQL + " ORDER BY " & SQLDate("CP05") & ",CP66,CP67,CP09"  '2009/2/20 MODIFY BY SONIA 排序原為收文日+收文號改為收文日+CREATE DATE+CREATE TIME+收文號
'Modify by Morgan 2011/1/3 修正日期排序百年問題
'strSql = strSql + " ORDER BY " & SQLDate("CP05") & ",CP66,CP67,DECODE(SUBSTR(CP09,1,1),'A','1','C','2','3'),CP09"
'Modify By Sindy 2014/5/27
'strSql = strSql + " ORDER BY SQLDatet2(CP05),CP66,CP67,DECODE(SUBSTR(CP09,1,1),'A','1','C','2','3'),CP09"
strSql = strSql + " ORDER BY SQLDatet2(CP05) DESC, CP66 DESC, CP67 DESC, CP09 DESC"
'2014/5/27 END
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly

If adoRecordset.RecordCount <> 0 Then
   If pub_QL04 <> "" Then Call InsertQueryLog(adoRecordset.RecordCount) 'Add By Sindy 2010/11/16
Else
   If pub_QL04 <> "" Then Call InsertQueryLog(0) 'Add By Sindy 2010/11/16
   Label4.Caption = ""
   Me.Label12.Caption = ""
   Me.lblClose.Caption = ""
   'add by nickc 2006/08/25
   lblCancel.Caption = ""
   lblCaseMap.Caption = "" 'Added by Lydia 2015/11/03
   lblCaseMap2.Caption = "" 'Added by Lydia 2019/11/28
   lblCMboth.Caption = "" 'Added by Lydia 2016/06/14
   lblNp605.Caption = "" 'Added by Lydia 2020/01/06
   cmdOK(0).Enabled = False
   cmdOK(1).Enabled = False
   cmdOK(2).Enabled = False
   cmdOK(3).Enabled = False
   cmdOK(4).Enabled = False
   Me.Enabled = True
   'Added by Lydia 2016/04/12  系統類別為 TS或S，且該案件備註欄(SP18)內含'轉入商標：'...字樣時，改顯示 "此案進度已轉入商標：XXX-XXXXX"，例 S-000307, S-000327, TS-001248
   If Str01 = "TS" Or Str01 = "S" Then
      strSql = "select SP18 from servicepractice where sp01='" & Str01 & "' and sp02='" & Str02 & "' and sp03='" & Str03 & "' and sp04='" & Str04 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If InStr(Trim("" & RsTemp(0)), "轉入商標") > 0 Then
            MsgBox "此案進度已" & Mid(Trim(RsTemp(0)), InStr(Trim(RsTemp(0)), "轉入商標")), , "查詢資料"
         Else
            intI = 0
         End If
      End If
      
      If intI = 0 Then ShowNoData
   Else
   'end 2016/04/12
      ShowNoData
   End If
   Screen.MousePointer = vbDefault
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If

grdDataList.Visible = False 'Add by Morgan 2007/10/17
If adoRecordset.RecordCount <> 0 Then
    grdDataList.row = 1
    
    'Modify By Sindy 2011/6/23
    AddCboName Combo1, "" & adoRecordset.Fields(19).Value, "" & adoRecordset.Fields(20).Value, "" & adoRecordset.Fields(21).Value
    
    If IsNull(adoRecordset.Fields(22)) Then
        Label4.Caption = ""
    Else
        Label4.Caption = adoRecordset.Fields(22)
    End If
    If IsNull(adoRecordset.Fields(25)) Then
        Me.Label12.Caption = ""
    Else
        Me.Label12.Caption = adoRecordset.Fields(25)
    End If
    
    If IsNull(adoRecordset.Fields(24)) Then
        Me.lblClose.Caption = ""
    Else
        Me.lblClose.Caption = "已閉卷"
    End If
    
    If IsNull(adoRecordset.Fields(27)) Then
        Me.Label15.Caption = ""
    Else
        Me.Label15.Caption = adoRecordset.Fields(27)
    End If
    'add by nickc 2006/08/25
    If IsNull(adoRecordset.Fields("IsCancel")) Then
        Me.lblCancel.Caption = ""
    Else
        Me.lblCancel.Caption = "已銷卷"
    End If
    
    'Add By Sindy 2010/01/19 增加顯示智權人員
    If Str01 = "FCP" Or Str01 = "FG" Then
      Me.Label16.Caption = GetPrjSalesNM(PUB_GetFCPSalesNo(Str01, Str02, Str03, Str04))
      'Added by Morgan 2022/4/28 外專案件分所號改為顯示目前程序人員--陳亭妙
      Label14.Visible = False
      Label15.Caption = "目前程序人員：" & GetPrjSalesNM(PUB_GetFCPHandler(Str01, Str02, Str03, Str04))
      'end 2022/4/28
    ElseIf Str01 = "FCL" Or Str01 = "LIN" Then
      Me.Label16.Caption = GetPrjSalesNM(PUB_GetFCLSalesNo(Str01, Str02, Str03, Str04))
    'Modify By Sindy 2010/01/21
    'ElseIf Str01 = "FCT" Or Str01 = "S" Then
    ElseIf Str01 = "FCT" Then
      Me.Label16.Caption = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, Str02, Str03, Str04))
    ElseIf Str01 = "S" Then
      If adoRecordset.Fields("SP09") = "000" Then
         Me.Label16.Caption = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, Str02, Str03, Str04))
      Else
         Me.Label16.Caption = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, Str02, Str03, Str04))
      End If
    '2010/01/21 End
    Else
      Me.Label16.Caption = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, Str02, Str03, Str04))
    End If
    '2010/01/19 End
    

End If
grdDataList.Rows = adoRecordset.RecordCount + 1
'add by nickc 2008/03/05
grdDataList.FixedCols = 0

Set grdDataList.Recordset = adoRecordset
SetDataListWidth
'分析字串  因為繳費年度與是否雙倍為長字串     只有專利有
If Str01 = "CFP" Or Str01 = "FCP" Or Str01 = "P" Then
    '宣告字串
    Dim StrStr(4) As String, StrStr1 As Variant, StrStr2 As Variant, StrStr3 As Variant, StrInt1 As Integer
    Dim StrIntT1 As Integer, StrIntT2 As Integer, StrIntT3 As Integer
    
    '2019/9/17 Add By Sindy 抓修法次數
    pa(1) = Str01
    pa(2) = Str02
    pa(3) = Str03
    pa(4) = Str04
    GetMoneyDate adoRecordset.Fields("pa08"), strPA09, pa, strExc(0), strExc(1), , , m_FixNo
    '2019/9/17 END
   
    '檢查資料筆數
    For i = 1 To adoRecordset.RecordCount      'GRDDATALIST.ROWS
        grdDataList.row = i
        'Modify by Morgan 2007/10/17 判斷有相關總收文號才做較快
        'Me.grdDataList.TextMatrix(i, 3) = Me.grdDataList.TextMatrix(i, 3) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(i, 2), "1")
        'If Me.GrdDataList.TextMatrix(i, 4) <> "" Then
           Me.grdDataList.TextMatrix(i, 3) = Me.grdDataList.TextMatrix(i, 3) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(i, 2), "1")
        'End If
        grdDataList.col = 9
        If IsNull(grdDataList.Text) Then
            StrStr(0) = ""
        Else
            StrStr(0) = ChangeTStringToWString(ChangeTDateStringToTString(grdDataList.Text))
        End If
        grdDataList.col = 14 '13
        If IsNull(grdDataList.Text) Then
            StrStr(1) = ""
            StrStr(2) = ""
        Else
            StrStr(1) = Left$(grdDataList.Text, InStr(1, grdDataList.Text, "-") - 1)
            StrStr(2) = Right$(grdDataList.Text, Len(grdDataList.Text) - (InStr(1, grdDataList.Text, "-")))
            StrStr1 = Split(StrStr(1), ",")
            StrStr2 = Split(StrStr(2), ",")
        End If
        grdDataList.col = 15 '14
        If IsNull(grdDataList.Text) Then
            StrStr(3) = ""
        Else
            StrStr(3) = grdDataList.Text
            StrStr3 = Split(StrStr(3), ",")
        End If
        StrIntT1 = UBound(StrStr1)
        StrIntT2 = UBound(StrStr2)
        StrIntT3 = UBound(StrStr3)
        grdDataList.col = 14 '13
        grdDataList.Text = ""
        grdDataList.col = 15 '14
        grdDataList.Text = ""
        'Modified by Morgan 2024/9/30 +控制年費性質(要移動指標,否則永遠是第1筆)
        'If StrStr(0) <> "" Then
        adoRecordset.MoveFirst
        adoRecordset.Move i - 1
        'modify by sonia 2025/4/22 +601領證P-126916
        If StrStr(0) <> "" And (adoRecordset.Fields("cp10").Value = "601" Or adoRecordset.Fields("cp10").Value = "605" Or adoRecordset.Fields("cp10").Value = "606" Or adoRecordset.Fields("cp10").Value = "607") Then
        'end 2024/9/30
            For j = 0 To StrIntT1
                If StrStr1(j) = StrStr(0) Then
                    'Add By Sindy 2019/9/17
                    If adoRecordset.Fields("cp10").Value = "606" Or _
                        adoRecordset.Fields("cp10").Value = "607" Then
                        'strTemp1 = Split(strArr(i), ",") '繳費年度
                        strFeeType = PUB_GetNa20Na22Na24(strPA09, adoRecordset.Fields("pa08").Value)
                        strYF15 = PUB_GetYF15(strPA09, adoRecordset.Fields("pa08").Value, "Y000000" & m_FixNo, strFeeType, CDbl(StrStr2(j)))
                        grdDataList.col = 14 '13
                        grdDataList.Text = strYF15
                    '2019/9/17 END
                    Else
                        grdDataList.col = 14 '13
                        grdDataList.Text = grdDataList.Text + StrStr2(j)
                    End If
                    '是否雙倍欄位若第一年非雙倍時, 此欄位是無資料的
                    If UBound(StrStr3) >= 0 Then
                        grdDataList.col = 15 '14
                        grdDataList.Text = grdDataList.Text + StrStr3(j)
                    End If
                End If
            Next j
        Else
            grdDataList.col = 14 '13
            grdDataList.Text = ""
            grdDataList.col = 15 '14
            grdDataList.Text = ""
        End If
        '收款情形
        IntTemp1 = 0
        IntTemp2 = 0
        Me.grdDataList.col = 26 '25
        If Not IsNull(grdDataList.Text) Then
            '2009/12/8 modify by sonia 加請款單
            'strSQL = "select A0k06,A0K07,'','',A0K17,A0K18 FROM ACC0K0 WHERE A0K01='" & grdDataList.Text & "'"
            'Modify by Morgan 2011/8/12 收據的收款情形改判斷CP79
            'If Mid(grdDataList.Text, 1, 1) = "E" Then
            '   strSql = "select A0k06,A0K07,'','',A0K17,A0K18 FROM ACC0K0 WHERE A0K01='" & grdDataList.Text & "'"
            'Else
            If Mid(grdDataList.Text, 1, 1) = "X" Then
            'end 2011/8/15
               'modify by sonia 2017/4/18 +A1K25銷帳單號FCP-48310之AA6012573更正X10604961已銷帳
               strSql = "select A1k11,0,'','',decode(a1k29,'Y',a1k11,nvl(A1K30,0)),0,A1K25 FROM ACC1K0 WHERE A1K01='" & grdDataList.Text & "'"
               
            'End If 'Remove by Morgan 2011/8/15
            '2009/12/8 end
               CheckOC2
               adoRecordset1.CursorLocation = adUseClient
               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                   If Not IsNull(adoRecordset1.Fields(0)) Then
                       IntTemp1 = IntTemp1 + adoRecordset1.Fields(0)
                   End If
                   If Not IsNull(adoRecordset1.Fields(1)) Then
                       IntTemp1 = IntTemp1 + adoRecordset1.Fields(1)
                   End If
                   If Not IsNull(adoRecordset1.Fields(4)) Then
                       IntTemp2 = IntTemp2 + adoRecordset1.Fields(4)
                   End If
                   If Not IsNull(adoRecordset1.Fields(5)) Then
                       IntTemp2 = IntTemp2 + adoRecordset1.Fields(5)
                   End If
                   If IntTemp1 = IntTemp2 Then
                        grdDataList.Text = "收回"
                   Else
                        If IntTemp2 = 0 Then
                            grdDataList.Text = "未收"
                            If "" & adoRecordset1.Fields(6) <> "" Then grdDataList.Text = "銷帳"    'add by sonia 2017/4/18 +A1K25銷帳單號FCP-48310之AA6012573更正X10604961已銷帳
                        Else
                            If IntTemp1 > IntTemp2 Then
                                grdDataList.Text = "部分收回"
                            End If
                        End If
                    End If
               Else
                    'grdDataList.Text = "查無此收據編號"   '2010/3/18 CANCEL BY SONIA
               End If
            End If 'Add by Morgan 2011/8/15
        End If
    Next i
Else
    For i = 1 To Me.grdDataList.Rows - 1
        grdDataList.row = i
        Me.grdDataList.TextMatrix(i, 3) = Me.grdDataList.TextMatrix(i, 3) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(i, 2), "1")
        '收款情形
        IntTemp1 = 0
        IntTemp2 = 0
        Me.grdDataList.col = 26 '25
        If Not IsNull(grdDataList.Text) And grdDataList.Text <> "" Then
            '2009/12/8 modify by sonia 加請款單
            'strSQL = "select A0k06,A0K07,'','',A0K17,A0K18 FROM ACC0K0 WHERE A0K01='" & grdDataList.Text & "'"
            'Modify by Morgan 2011/8/18 收據的收款情形改判斷CP79
            'If Mid(grdDataList.Text, 1, 1) = "E" Then
            '   strSql = "select A0k06,A0K07,'','',A0K17,A0K18 FROM ACC0K0 WHERE A0K01='" & grdDataList.Text & "'"
            'Else
            If Mid(grdDataList.Text, 1, 1) = "X" Then
            'end 2011/8/18
               'modify by sonia 2017/4/18 +A1K25銷帳單號FCP-48310之AA6012573更正X10604961已銷帳
               strSql = "select A1k11,0,'','',decode(a1k29,'Y',a1k11,nvl(A1K30,0)),0,A1K25 FROM ACC1K0 WHERE A1K01='" & grdDataList.Text & "'"
               
            'End If 'Remove by Morgan 2011/8/18
            '2009/12/8 end
               CheckOC2
               adoRecordset1.CursorLocation = adUseClient
               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                   If Not IsNull(adoRecordset1.Fields(0)) Then
                       IntTemp1 = IntTemp1 + adoRecordset1.Fields(0)
                   End If
                   If Not IsNull(adoRecordset1.Fields(1)) Then
                       IntTemp1 = IntTemp1 + adoRecordset1.Fields(1)
                   End If
                   If Not IsNull(adoRecordset1.Fields(4)) Then
                       IntTemp2 = IntTemp2 + adoRecordset1.Fields(4)
                   End If
                   If Not IsNull(adoRecordset1.Fields(5)) Then
                       IntTemp2 = IntTemp2 + adoRecordset1.Fields(5)
                   End If
                   If IntTemp1 = IntTemp2 Then
                        grdDataList.Text = "收回"
                   Else
                        If IntTemp2 = 0 Then
                            grdDataList.Text = "未收"
                            If "" & adoRecordset1.Fields(6) <> "" Then grdDataList.Text = "銷帳"    'add by sonia 2017/4/18 +A1K25銷帳單號FCP-48310之AA6012573更正X10604961已銷帳
                        Else
                            If IntTemp1 > IntTemp2 Then
                                grdDataList.Text = "部分收回"
                            End If
                        End If
                    End If
               Else
                    'grdDataList.Text = "查無此收據編號"   '2010/3/18 CANCEL BY SONIA
               End If
            End If 'Add by Morgan 2011/8/18
        End If
    Next i
End If
CheckOC
intK = grdDataList.Rows
mFixs = 0 'Added by Lydia 2018/03/22
'2005/12/19 ADD BY SONIA
'Modify By Sindy 2009/07/24 增加LIN系統類別
'modify by sonia 2019/7/30 +ACS系統類別
'Modify by Amy 2020/11/13 原:And Str01 <> "ACS"
If (Str01 <> "L" And Str01 <> "FCL" And Str01 <> "CFL" And Str01 <> "LA" And Str01 <> "LIN") Or Str01 = "ACS" Then
   'Modifie by Lydia 2018/03/16 固定從收文日到承辦人共5個欄位 (by 陳金蓮)
   'grdDataList.FixedCols = 0
   grdDataList.FixedCols = 6
   mFixs = 5 'Added by Lydia 2018/03/22
Else
   grdDataList.FixedCols = 5
   mFixs = 4 'Added by Lydia 2018/03/22
End If
'2005/12/19 END

'若僅有一筆資料
If Me.grdDataList.Rows = 2 Then
   '直接勾選
   'grdDataList.Visible = False 'Remove by Morgan 2007/10/17 在外層控制
   grdDataList.col = 0
   grdDataList.row = 1
   grdDataList.Text = "V"
   For i = 0 To grdDataList.Cols - 1
       grdDataList.col = i
       grdDataList.CellBackColor = &HFFC0C0
   Next i
   'grdDataList.Visible = True 'Remove by Morgan 2007/10/17 在外層控制
End If

''Add by Morgan 2011/1/3 日期欄位靠右
'grdDataList.ColAlignment(1) = flexAlignRightCenter
'grdDataList.ColAlignment(6) = flexAlignRightCenter
'grdDataList.ColAlignment(7) = flexAlignRightCenter
'grdDataList.ColAlignment(8) = flexAlignRightCenter
'grdDataList.ColAlignment(9) = flexAlignRightCenter
''end 2011/1/3
Call SetDataListWidth_new 'Modify By Sindy 2011/2/25

grdDataList.Visible = True 'Add by Morgan 2007/10/17

'add by nickc 2005/05/30  檢查有無分割或相關卷號
cmdOK(8).Visible = ChkDataBy308(Label3.Caption)
cmdOK(9).Visible = ChkDataByCR(Label3.Caption)

'Added by Lydia 2015/11/03　顯示一案兩請，擬制喪失新穎性案件
lblCaseMap.Caption = ""
lblCaseMap2.Caption = "" 'Added by Lydia 2019/11/28
If PUB_GetRefCaseChk(Str01, Str02, Str03, Str04, "CASEMAP", "3") = True Then
   lblCaseMap.Caption = "一案兩請"
End If
If PUB_GetRefCaseChk(Str01, Str02, Str03, Str04, "CASEMAP", "6") = True Then
   'Modified by Lydia 2019/11/28 P-123733有一案兩請和擬制喪失新穎性案件
   'lblCaseMap.Caption = "擬制喪失新穎性案件"
   lblCaseMap2.Caption = "擬制喪失新穎性案件"
End If
'end 2015/11/03

'Added by Lydia 2016/06/14 +台灣大陸案件提示
lblCMboth.Caption = ""
strExc(9) = GetPrjNation1(Label3.Caption)
If (Str01 = "P" Or Str01 = "FCP") And strExc(9) = 台灣國家代號 Then
   If PUB_GetRefCaseChk(Str01, Str02, Str03, Str04, "CASEMAP", "0", "A", 大陸國家代號) Then
      lblCMboth.Caption = "有大陸案"
   End If
ElseIf Str01 = "P" And strExc(9) = 大陸國家代號 Then
   If PUB_GetRefCaseChk(Str01, Str02, Str03, Str04, "CASEMAP", "0", "A", 台灣國家代號) Then
      lblCMboth.Caption = "有台灣案"
   End If
End If
'end 2016/06/14

lblNp605.Caption = GetNp605State(Str01, Str02, Str03, Str04) 'Added by Lydia 2020/01/06 專利案件：年費不續辦

Me.Enabled = True
Me.SetFocus
End Sub

'Sub StrMenu1()
'Dim strSql  As String
'Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String
'Dim i As Integer, j As Integer
'Dim StrSQLa As String
'
''DoEvents Add By Sindy 2019/1/4 Mark,因為會和視窗的function(MenuForFormControl)有ErrCode互影響
'Me.Enabled = False
'Str01 = ""
'Str02 = ""
'Str03 = ""
'Str04 = ""
'Label3.Caption = frm100101_2.Tag
'If Left(Me.Tag, 1) = "N" Then
'   strSql = Right(Me.Tag, Len(Me.Tag) - 1)
'Else
'   strSql = Me.Tag
'End If
'Str01 = SystemNumber(strSql, 1)
'Str02 = SystemNumber(strSql, 2)
'Str03 = SystemNumber(strSql, 3)
'Str04 = SystemNumber(strSql, 4)
'
''Added by Lydia 2019/06/10 依國別抓案件性質欄位
'Call ClsPDCheckCaseCodeIsExist(Str01, Str02, Str03, Str04, , , , , strPA09)
'If strPA09 <= "010" Then
'    strCon = "CPM03"
'Else
'    strCon = "CPM04"
'End If
''end 2019/06/10
'
''Add By Sindy 2013/7/1
''Modify By Sindy 2015/10/21 + And Str01 <> "PS" And Str01 <> "CPS"
''Modify By Sindy 2018/7/24 + Not (Left(Str01, 1) = "T" Or Str01 = "FCT")
'If (Str01 <> "P" And Str01 <> "CFP" And Str01 <> "PS" And Str01 <> "CPS") And _
'   Pub_StrUserSt03 <> "M51" And _
'   Not (Left(Str01, 1) = "T" Or Str01 = "FCT") Then
'   'cmdOK(13).Visible = False 'Modify By Sindy 2015/3/9 Mark:不分系統別都可以查看
'   'cmdOK(14).Visible = False 'Modify By Sindy 2017/8/16 Mark
'   cmdOK(15).Visible = False
'End If
''2013/7/1 End
'
''add by nickc 2006/07/10 下面各句加入 cp60
''Modify By Sindy 2011/6/23 +CP16
''Modify by Morgan 2011/8/15 收據的收款情形改判斷CP79
''Add by Lydia 2014/11/14 進度檔結果欄後加入下一期限及下一期限日期
'Dim midSql As String
''Modified by Lydia 2019/06/10 重整語法:先抓同一收文號的最小所限(有排除特定性質,ex.CFP-014272的C92000488的公開999排除)
'                                     '再抓符合最小所限的最大np22 ; (起因有人反應FCT-042306只出現一筆下一程序性質)
''midSql = "(select npf.np01,npf.np07,cpmf.cpm03 as 下一程序性質 ,npf.np08 from nextprogress npf,casepropertymap cpmf " & _
'     " where npf.np02='" & Str01 & "' AND npf.np03='" & Str02 & "' AND npf.np04='" & Str03 & "' AND npf.np05='" & Str04 & "' " & _
'     " and npf.np02=cpmf.cpm01(+) and npf.np07=cpmf.cpm02(+) " & strNpSqlOfNoSalesDuty & _
'     " and np08 in (select min(n2f.np08) D_date from nextprogress n2f " & _
'     " where n2f.np02='" & Str01 & "' AND n2f.np03='" & Str02 & "' AND n2f.np04='" & Str03 & "' AND n2f.np05='" & Str04 & "' " & strNpSqlOfNoSalesDuty & _
'     " group by n2f.np01)) nm "
'midSql = "(select np01,np08,np22," & strCon & " as 下一程序性質 from nextprogress,casepropertymap " & _
'              "where (np01,np08,np22) in (select n2.np01,n2.np08,max(n2.np22) np22 " & _
'              "from nextprogress n2  " & _
'              "where (n2.np01,n2.np08) in ( " & _
'                    "select np01,min(np08) as mdate from nextprogress where np02='" & Str01 & "' AND np03='" & Str02 & "' AND np04='" & Str03 & "' AND np05='" & Str04 & "' " & _
'                         strNpSqlOfNoSalesDuty & " group by np01) " & Replace(strNpSqlOfNoSalesDuty, "np", "n2.np") & _
'              "group by n2.np01,n2.np08) and np02=cpm01(+) and np07=cpm02(+)) nm "
'
'Select Case Pub_RplStr(Str01)
'Case "CFP", "FCP", "P"   '專利
''是否出名與代理人對調
'      '引進是否閉卷欄
'      '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'      StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
'    'Modify By Cheng 2002/11/11
''      strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,Nvl(DECODE(PA09,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,S1.ST02 as 承辦人,S2.ST02 as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,PA73||'-'||PA72 as 繳費年度,PA74 as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, patent.PA05,PATENT.PA06,PATENT.PA07,PATENT.PA11,PATENT.PA09,PA57,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60," & IIf(pub_strUserOffice = "1", "pa108", "pa136") & " as IsCancel,cp16 " & _
''               " From caseprogress,staff s1,staff s2,patent,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind " & _
''               " Where caseprogress.CP01=patent.PA01(+) and caseprogress.CP02=patent.PA02(+) and caseprogress.CP03=patent.PA03(+) and caseprogress.CP04=patent.PA04(+) and CP01='" & Str01 & "' and CP02='" & Str02 & "' and CP03='" & Str03 & "' and CP04='" & Str04 & "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1) = FA02(+) AND CP09<'C' AND CP01=SK01(+) "
''Add by Lydia 2014/11/14 進度檔結果欄後加入下一期限及下一期限日期
'      'modify by sonia 2019/5/8 收款情形加銷帳CFP-025274及退費T-089878
'      'Modified by Lydia 2019/06/10 比照含來函,增加PA22,PA47
'      strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,Nvl(DECODE(PA09,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,S1.ST02 as 承辦人,S2.ST02 as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,PA73||'-'||PA72 as 繳費年度,PA74 as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, patent.PA05,PATENT.PA06,PATENT.PA07,PATENT.PA11,PATENT.PA09,PA57,PA22,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,PA47 ," & IIf(pub_strUserOffice = "1", "pa108", "pa136") & " as IsCancel,cp16 " & _
'               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff s1,staff s2,patent,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind, " & _
'               midSql & " Where caseprogress.CP01=patent.PA01(+) and caseprogress.CP02=patent.PA02(+) and caseprogress.CP03=patent.PA03(+) and caseprogress.CP04=patent.PA04(+) and CP01='" & Str01 & "' and CP02='" & Str02 & "' and CP03='" & Str03 & "' and CP04='" & Str04 & "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1) = FA02(+) AND CP09<'C' AND CP01=SK01(+) " & _
'              " and caseprogress.CP09=nm.np01(+) "
'Case "CFT", "FCT", "T", "TF"   '商標
''是否出名與代理人對調
'      '引進是否閉卷欄
'      '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'      StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
''      strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,Nvl(DECODE(tm10,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,S1.ST02 as 承辦人,S2.ST02 as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,' ' as 繳費年度,' ' as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, trademark.TM05,TRADEMARK.TM06,TRADEMARK.TM07,TRADEMARK.TM12,TRADEMARK.TM10,TM29,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60," & IIf(pub_strUserOffice = "1", "tm57", "tm73") & " as IsCancel,cp16 " & _
''               " From caseprogress,staff S1,STAFF S2,tradEmark,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind " & _
''               " Where caseprogress.CP01=TRADEMARK.TM01(+) and caseprogress.CP02=TRADEMARK.TM02(+) and caseprogress.CP03=TRADEMARK.TM03(+) and caseprogress.CP04=TRADEMARK.TM04(+) AND CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP09<'C' "
''      strSql = strSql + " AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) "
''Add by Lydia 2014/11/14 進度檔結果欄後加入下一期限及下一期限日期
'      'modify by sonia 2019/5/8 收款情形加銷帳CFP-025274及退費T-089878
'      'Modified by Lydia 2019/06/10 比照含來函,增加TM15,TM34
'      strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,Nvl(DECODE(tm10,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,S1.ST02 as 承辦人,S2.ST02 as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,' ' as 繳費年度,' ' as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, trademark.TM05,TRADEMARK.TM06,TRADEMARK.TM07,TRADEMARK.TM12,TRADEMARK.TM10,TM29,TM15, decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,TM34, " & IIf(pub_strUserOffice = "1", "tm57", "tm73") & " as IsCancel,cp16 " & _
'               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,tradEmark,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind, " & _
'               midSql & " Where caseprogress.CP01=TRADEMARK.TM01(+) and caseprogress.CP02=TRADEMARK.TM02(+) and caseprogress.CP03=TRADEMARK.TM03(+) and caseprogress.CP04=TRADEMARK.TM04(+) AND CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP09<'C' "
'      strSql = strSql + " and caseprogress.CP09=nm.np01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) "
'
''Add by Amy 2020/11/13 獨立出來,抓資料的方式同專利/商標
'Case "ACS"
'      StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
'
'      strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,Nvl(DECODE(LC15,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,S1.ST02 as 承辦人,S2.ST02 as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,' ' as 繳費年度,' ' as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, LawCase.LC05,LawCase.LC06,LawCase.LC07,'',LawCase.LC15,LC08,'', decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,lc16," & IIf(pub_strUserOffice = "1", "lc34", "lc36") & " as IsCancel,cp16 " & _
'               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,LawCase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind, " & _
'               midSql & " Where caseprogress.CP01=LawCase.LC01(+) and caseprogress.CP02=LawCase.LC02(+) and caseprogress.CP03=LawCase.LC03(+) and caseprogress.CP04=LawCase.LC04(+) AND CP01='" & Str01 & "' AND CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP09<'C' "
'      strSql = strSql + " and caseprogress.CP09=nm.np01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) "
'
''Modify By Sindy 2009/07/24 增加LIN系統類別
''modify by sonia 2019/7/30 +ACS系統類別
''Modify by Amy 2020/11/13 ACS另外獨立寫,抓資料的方式同專利/商標
'Case "CFL", "FCL", "L", "LIN"   '法務
''是否出名與代理人對調
'      '引進是否閉卷欄
'      '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'      StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
''      strSQL = "select ' ' AS V," & SQLDate("CP05") & " as 收文日,CP09 as 總收文號,Nvl(DECODE(lc15,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,S1.ST02 as 承辦人,S2.ST02 as 智權人員," & SQLDate("cp06") & " as 本所期限," & SQLDate("cp07") & " as 法定期限," & SQLDate("cp27") & " as 發文日," & SQLDate("cp57") & " as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,' ' as 繳費年度,' ' as 是否雙倍,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, lawcase.LC05,LAWCASE.LC06,LAWCASE.LC07,'無',LAWCASE.LC15,LC08 " & _
''               " From caseprogress,staff S1,STAFF S2,lawcase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SYSTEMKIND " & _
''               " Where caseprogress.CP01=lawcase.LC01(+) and caseprogress.CP02=lawcase.LC02(+) and caseprogress.CP03=lawcase.LC03(+) and caseprogress.CP04=lawcase.LC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP09<'C'  "
''      strSQL = strSQL + " AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) ORDER BY " & SQLDate("CP05") & ",CP09 "
''edit by nickc 2008/03/05 修正欄位
''      strSQL = "select ' ' AS V," & SQLDate("CP05") & " as 收文日,CP09 as 總收文號,CP64 as 進度備註, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦律師,decode(CP29,S3.ST01,S3.ST02) as 承辦法務,NVL(S2.ST02,CP13) as 智權人員,cp18 as 點數," & SQLDate("CP06") & " as 本所期限," & SQLDate("CP07") & " as 法定期限,'' as 會稿日," & SQLDate("CP27") & " as 發文日," & SQLDate("CP46") & " as 回執日,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))))) as 相關人,or02 as 機關名稱, '',lawcase.LC05,LAWCASE.LC06,LAWCASE.LC07,'無',LAWCASE.LC15,LC08,'',cp60," & IIf(pub_strUserOffice = "1", "lc34", "lc36") & " as IsCancel " & _
''               " From caseprogress,staff S1,STAFF S2,staff s3,lawcase,CUSTOMER,CASEPROPERTYMAP,SystemKind,organization " & _
''               " Where caseprogress.CP01=lawcase.LC01(+) and caseprogress.CP02=lawcase.LC02(+) and caseprogress.CP03=lawcase.LC03(+) and caseprogress.CP04=lawcase.LC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+)  AND CP09<'C' "
''2009/9/9 MODIFY BY SONIA 法務無進度備註時帶案件性質
''      strSql = " select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')' as 進度備註,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦律師,decode(CP29,S3.ST01,S3.ST02) as 承辦法務,NVL(S2.ST02,CP13) as 智權人員,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,'' as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & " '',lawcase.LC05,LAWCASE.LC06,LAWCASE.LC07,'',LAWCASE.LC15,LC08,'',decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60,lc16," & IIf(pub_strUserOffice = "1", "lc34", "lc36") & " as IsCancel,cp16 " & _
''               " From caseprogress,staff S1,STAFF S2,staff s3,lawcase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization " & _
''               " Where caseprogress.CP01=lawcase.LC01(+) and caseprogress.CP02=lawcase.LC02(+) and caseprogress.CP03=lawcase.LC03(+) and caseprogress.CP04=lawcase.LC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+)  AND CP09<'C' AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
''      strSql = strSql + " AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) and cp71=or01(+) "
''Add by Lydia 2014/11/14 進度檔結果欄後加入下一期限及下一期限日期
''Modified by Lydia 2015/10/05
'      'Modified by Lydia 2016/05/30 備註顯示回執退件日/回執未回郵局送達日
'      'strSql = " select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')' as 進度備註,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,decode(CP29,S3.ST01,S3.ST02) as 協辦人員,NVL(S2.ST02,CP13) as 智權人員,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,'' as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & " '',lawcase.LC05,LAWCASE.LC06,LAWCASE.LC07,'',LAWCASE.LC15,LC08,'',decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60,lc16," & IIf(pub_strUserOffice = "1", "lc34", "lc36") & " as IsCancel,cp16 " & _
'               "  ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,staff s3,lawcase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization, " & _
'               midSql & " Where caseprogress.CP01=lawcase.LC01(+) and caseprogress.CP02=lawcase.LC02(+) and caseprogress.CP03=lawcase.LC03(+) and caseprogress.CP04=lawcase.LC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+)  AND CP09<'C' AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
'      'modify by sonia 2019/5/8 收款情形加銷帳CFP-025274及退費T-089878
'      'Modified by Lydia 2020/07/15 法律所案源收文：增加案源之介紹人
'      'strSql = " select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號," & _
'               "decode(sign(instr('3,4',sk02)),1,decode(cp46,19221111,'回執退件日:'||sqldatet(cp47)||';',19220101,'回執未回郵局送達日:'||sqldatet(cp47)||';',''),'')||CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')' as 進度備註," & _
'               "NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,decode(CP29,S3.ST01,S3.ST02) as 協辦人員,NVL(S2.ST02,CP13) as 智權人員,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,'' as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & " '',lawcase.LC05,LAWCASE.LC06,LAWCASE.LC07,'',LAWCASE.LC15,LC08,'',decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,lc16," & IIf(pub_strUserOffice = "1", "lc34", "lc36") & " as IsCancel,cp16 " & _
'               "  ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,staff s3,lawcase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization, " & _
'               midSql & " Where caseprogress.CP01=lawcase.LC01(+) and caseprogress.CP02=lawcase.LC02(+) and caseprogress.CP03=lawcase.LC03(+) and caseprogress.CP04=lawcase.LC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+)  AND CP09<'C' AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
'      'strSql = strSql + " and caseprogress.CP09=nm.np01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) and cp71=or01(+) "
'      strSql = " select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號," & _
'               "decode(sign(instr('3,4',sk02)),1,decode(cp46,19221111,'回執退件日:'||sqldatet(cp47)||';',19220101,'回執未回郵局送達日:'||sqldatet(cp47)||';',''),'')||CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')' as 進度備註," & _
'               "NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,decode(CP29,S3.ST01,S3.ST02) as 協辦人員,NVL(S2.ST02,CP13) as 智權人員,DECODE(LOS15,NULL,NULL,GETSTAFFNAMELIST(LOS04)) 介紹人,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,'' as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & _
'               StrSQLa & " lawcase.LC05,LAWCASE.LC06,LAWCASE.LC07,'' as vt22,LAWCASE.LC15,LC08,'' as vt25,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,lc16," & IIf(pub_strUserOffice = "1", "lc34", "lc36") & " as IsCancel,cp16 " & _
'               "  ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,staff s3,lawcase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization, LawOfficeSource," & _
'               midSql & " Where caseprogress.CP01=lawcase.LC01(+) and caseprogress.CP02=lawcase.LC02(+) and caseprogress.CP03=lawcase.LC03(+) and caseprogress.CP04=lawcase.LC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+)  AND CP09<'C' AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) "
'      strSql = strSql + " and caseprogress.CP09=nm.np01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) and cp71=or01(+) and CP162=LOS15(+) "
'      'end 2020/07/15
''2009/9/9 MODIFY BY SONIA 顧問進度備註欄同時帶案件性質
'Case "LA"                      '顧問
''是否出名與代理人對調
'      '引進是否閉卷欄
'      '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'      StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
''      strSQL = "select ' ' AS V," & SQLDate("CP05") & " as 收文日,CP09 as 總收文號,Nvl(CPM03,CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,S1.ST02 as 承辦人,S2.ST02 as 智權人員," & SQLDate("cp06") & " as 本所期限," & SQLDate("cp07") & " as 法定期限," & SQLDate("cp27") & " as 發文日," & SQLDate("cp57") & " as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,' ' as 繳費年度,' ' as 是否雙倍,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, hirecase.HC06,'','','','',HC09 " & _
''               " From caseprogress,staff S1,STAFF S2,hirecase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SYSTEMKIND " & _
''               " Where caseprogress.CP01=hirecase.HC01(+) and caseprogress.CP02=hirecase.HC02(+) and caseprogress.CP03=hirecase.HC03 and caseprogress.CP04=hirecase.HC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP09<'C'  "
''      strSQL = strSQL + " AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) ORDER BY " & SQLDate("CP05") & ",CP09 "
''edit by nickc 2008/03/05 修正
''      strSQL = "select ' ' AS V," & SQLDate("CP05") & " as 收文日,CP09 as 總收文號,decode(cp10,'0',sqldatet(cp53)||'--'||sqldatet(cp54),CP64) as 進度備註, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦律師,decode(CP29,S3.ST01,S3.ST02) as 承辦法務,NVL(S2.ST02,CP13) as 智權人員,cp18 as 點數," & SQLDate("CP06") & " as 本所期限," & SQLDate("CP07") & " as 法定期限,'' as 會稿日," & SQLDate("CP27") & " as 發文日," & SQLDate("CP46") & " as 回執日,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))))) as 相關人,or02 as 機關名稱,'',hirecase.HC06,'','','','',HC09,'',cp60," & IIf(pub_strUserOffice = "1", "hc19", "hc20") & " as IsCancel " & _
'               " From caseprogress,staff S1,STAFF S2,staff s3,hirecase,CUSTOMER,CASEPROPERTYMAP,SystemKind,organization " & _
'               " Where caseprogress.CP01=hirecase.HC01(+) and caseprogress.CP02=hirecase.HC02(+) and caseprogress.CP03=hirecase.HC03(+) and caseprogress.CP04=hirecase.HC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+)  AND CP09<'C'  "
''      strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,decode(cp10,'0',sqldatet(cp53)||'--'||sqldatet(cp54),CP64||'('||NVL(CPM03,CP10)||')') as 進度備註,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦律師,decode(CP29,S3.ST01,S3.ST02) as 承辦法務,NVL(S2.ST02,CP13) as 智權人員,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,'' as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "'',hirecase.HC06,'','','','',HC09,'',decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60,hc07," & IIf(pub_strUserOffice = "1", "hc19", "hc20") & " as IsCancel,cp16 " & _
''               " From caseprogress,staff S1,STAFF S2,staff s3,hirecase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization " & _
''               " Where caseprogress.CP01=hirecase.HC01(+) and caseprogress.CP02=hirecase.HC02(+) and caseprogress.CP03=hirecase.HC03(+) and caseprogress.CP04=hirecase.HC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+)  AND CP09<'C'  AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+)  "
''      strSql = strSql + " AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) and cp71=or01(+) "
''Add by Lydia 2014/11/14 進度檔結果欄後加入下一期限及下一期限日期
''Modified by Lydia 2015/10/05
'      'Modified by Lydia 2016/05/30 備註顯示回執退件日/回執未回郵局送達日
'      'strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,decode(cp10,'0',sqldatet(cp53)||'--'||sqldatet(cp54),CP64||'('||NVL(CPM03,CP10)||')') as 進度備註,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,decode(CP29,S3.ST01,S3.ST02) as 協辦人員,NVL(S2.ST02,CP13) as 智權人員,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,'' as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "'',hirecase.HC06,'','','','',HC09,'',decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60,hc07," & IIf(pub_strUserOffice = "1", "hc19", "hc20") & " as IsCancel,cp16 " & _
'               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,staff s3,hirecase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization, " & _
'               midSql & " Where caseprogress.CP01=hirecase.HC01(+) and caseprogress.CP02=hirecase.HC02(+) and caseprogress.CP03=hirecase.HC03(+) and caseprogress.CP04=hirecase.HC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+)  AND CP09<'C'  AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+)  "
'      'modify by sonia 2019/5/8 收款情形加銷帳CFP-025274及退費T-089878
'      'Modified by Lydia 2020/07/15 法律所案源收文：增加案源之介紹人
'      'strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號," & _
'               "decode(sign(instr('3,4',sk02)),1,decode(cp46,19221111,'回執退件日:'||sqldatet(cp47)||';',19220101,'回執未回郵局送達日:'||sqldatet(cp47)||';',''),'')||decode(cp10,'0',sqldatet(cp53)||'--'||sqldatet(cp54),CP64||'('||NVL(CPM03,CP10)||')') as 進度備註," & _
'               "NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,decode(CP29,S3.ST01,S3.ST02) as 協辦人員,NVL(S2.ST02,CP13) as 智權人員,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,'' as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "'',hirecase.HC06,'','','','',HC09,'',decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,hc07," & IIf(pub_strUserOffice = "1", "hc19", "hc20") & " as IsCancel,cp16 " & _
'               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,staff s3,hirecase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization, " & _
'               midSql & " Where caseprogress.CP01=hirecase.HC01(+) and caseprogress.CP02=hirecase.HC02(+) and caseprogress.CP03=hirecase.HC03(+) and caseprogress.CP04=hirecase.HC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+)  AND CP09<'C'  AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+)  "
'      'strSql = strSql + " and caseprogress.CP09=nm.np01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) and cp71=or01(+) "
'      strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號," & _
'               "decode(sign(instr('3,4',sk02)),1,decode(cp46,19221111,'回執退件日:'||sqldatet(cp47)||';',19220101,'回執未回郵局送達日:'||sqldatet(cp47)||';',''),'')||decode(cp10,'0',sqldatet(cp53)||'--'||sqldatet(cp54),CP64||'('||NVL(CPM03,CP10)||')') as 進度備註," & _
'               "NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))) as 相關人, NVL(CP43,'') as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,decode(CP29,S3.ST01,S3.ST02) as 協辦人員,NVL(S2.ST02,CP13) as 智權人員,DECODE(LOS15,NULL,NULL,GETSTAFFNAMELIST(LOS04)) 介紹人,cp18 as 點數,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,'' as 會稿日,SQLDATET2(CP27) as 發文日," & SQLDate("CP46") & " as 回執日,SQLDATET2(CP57) as 取消收文日," & _
'               StrSQLa & " hirecase.HC06,'' as vt20,'' as vt21,'' as vt22,'' as vt23,HC09,'' as vt25,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,hc07," & IIf(pub_strUserOffice = "1", "hc19", "hc20") & " as IsCancel,cp16 " & _
'               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,staff s3,hirecase,CUSTOMER,CASEPROPERTYMAP,FAGENT,SystemKind,organization, LawOfficeSource," & _
'               midSql & " Where caseprogress.CP01=hirecase.HC01(+) and caseprogress.CP02=hirecase.HC02(+) and caseprogress.CP03=hirecase.HC03(+) and caseprogress.CP04=hirecase.HC04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  and cp29=s3.st01(+) AND CP01=CPM01(+) AND CP10=CPM02(+)  AND CP09<'C'  AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+)  "
'      strSql = strSql + " and caseprogress.CP09=nm.np01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) and cp71=or01(+) and CP162=LOS15(+) "
'      'end 2020/07/15
'Case Else                  '服務
''是否出名與代理人對調
'      '引進是否閉卷欄
'      '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'      StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
''      strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,Nvl(DECODE(sp09,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,S1.ST02 as 承辦人,S2.ST02 as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,' ' as 繳費年度,' ' as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, servicepractice.SP05,SERVICEPRACTICE.SP06,SERVICEPRACTICE.SP07,SERVICEPRACTICE.SP11,SERVICEPRACTICE.SP09,SP15,decode(substr(cp60,1,1),'E',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')),cp60) cp60," & IIf(pub_strUserOffice = "1", "sp61", "sp68") & " as IsCancel,cp16 " & _
''               " From caseprogress,staff S1,STAFF S2,servicepractice,CUSTOMER,CASEPROPERTYMAP,FAGENT,SYSTEMKIND " & _
''               " Where caseprogress.CP01=servicepractice.SP01(+) and caseprogress.CP02=servicepractice.SP02(+) and caseprogress.CP03=servicepractice.SP03(+) and caseprogress.CP04=servicepractice.SP04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP09<'C'  "
''      strSql = strSql + " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) "
''Add by Lydia 2014/11/14 進度檔結果欄後加入下一期限及下一期限日期
'      'modify by sonia 2019/5/8 收款情形加銷帳CFP-025274及退費T-089878
'      'Modified by Lydia 2019/06/10 比照含來函,增加SP14,SP28
'      strSql = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,Nvl(DECODE(sp09,'000',CPM03,CPM04),CP10) as 案件性質,NVL(CP43,'') as 相關總收文號,S1.ST02 as 承辦人,S2.ST02 as 智權人員,SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日," & StrSQLa & "decode(CP24,'1','准','2','駁') as 結果,' ' as 繳費年度,' ' as 是否雙倍,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))))) as 相關人,CP22 as 是否出名,CP64 as 進度備註, servicepractice.SP05,SERVICEPRACTICE.SP06,SERVICEPRACTICE.SP07,SERVICEPRACTICE.SP11,SERVICEPRACTICE.SP09,SP15,SP14,decode(substr(cp60,1,1),'E',decode(cp78,cp16,'退費',decode(cp77,cp16,'銷帳',decode(cp79,0,'收回',decode(sign(cp75),1,'部分收回','未收')))),cp60) cp60,SP28," & IIf(pub_strUserOffice = "1", "sp61", "sp68") & " as IsCancel,cp16 " & _
'               " ,nm.下一程序性質 ,SQLDATET2(nm.np08) as 下一程序所限 From caseprogress,staff S1,STAFF S2,servicepractice,CUSTOMER,CASEPROPERTYMAP,FAGENT,SYSTEMKIND,Acc090, " & _
'               midSql & " Where caseprogress.CP01=servicepractice.SP01(+) and caseprogress.CP02=servicepractice.SP02(+) and caseprogress.CP03=servicepractice.SP03(+) and caseprogress.CP04=servicepractice.SP04(+) AND CP01='" + Str01 + "' AND CP02='" + Str02 + "' AND CP03='" + Str03 + "' AND CP04='" + Str04 + "' AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP09<'C'  "
'      strSql = strSql + " and caseprogress.CP09=nm.np01(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(CP56,1,8)=CUSTOMER.CU01(+) AND SUBSTR(CP56,9,1)=CUSTOMER.CU02(+) AND CP01=SK01(+) AND CP12=A0901(+) "
'      'Add By Sindy 2020/5/12
'      '以操作人員之st15再抓ACC090之A0911，再以進度CP12再抓ACC090之A0911，若與操作人員帶出之A0911相同才可顯示進度。卷宗區也是。
'      'st05 in ('00',’01’)人員不受上述限制
'      If InStr("00,01", Pub_strUserST05) = 0 And Str01 = "TT" And Str02 = "999999" Then
'         strSql = strSql + " AND A0911=(select a90.A0911 from acc090 a90 where a90.a0901='" & Pub_StrUserSt15 & "') "
'      End If
'      '2020/5/12 END
'End Select
''2009/12/2 modify by sonia 同日C類來函同時產生B類收文,應收顯示C類再顯示B類 FCT-029810,故再加收文號A->C->B
''strSQL = strSQL + " ORDER BY " & SQLDate("CP05") & ",CP66,CP67,CP09"  '2009/2/20 MODIFY BY SONIA 排序原為收文日+收文號改為收文日+CREATE DATE+CREATE TIME+收文號
''Modify by Morgan 2011/1/3 修正日期排序百年問題
''strSql = strSql + " ORDER BY " & SQLDate("CP05") & ",CP66,CP67,DECODE(SUBSTR(CP09,1,1),'A','1','C','2','3'),CP09"
''Modify By Sindy 2014/5/27
''strSql = strSql + " ORDER BY SQLDatet2(CP05),CP66,CP67,DECODE(SUBSTR(CP09,1,1),'A','1','C','2','3'),CP09"
'strSql = strSql + " ORDER BY SQLDatet2(CP05) DESC, CP66 DESC, CP67 DESC, CP09 DESC"
''2014/5/27 END
'CheckOC
'adoRecordset.CursorLocation = adUseClient
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset.RecordCount = 0 Then
'    Label4.Caption = ""
'    Me.lblClose.Caption = ""
'    'add by nickc 2006/08/25
'    lblCancel = ""
'    lblCaseMap.Caption = "" 'Added by Lydia 2015/11/03
'    lblCaseMap2.Caption = "" 'Added by Lydia 2019/11/28
'    lblCMboth.Caption = "" 'Added by Lydia 2016/06/14
'    lblNp605.Caption = "" 'Added by Lydia 2020/01/06
'    cmdOK(0).Enabled = False
'    cmdOK(1).Enabled = False
'    cmdOK(2).Enabled = False
'    cmdOK(3).Enabled = False
'    cmdOK(4).Enabled = False
'    Me.Enabled = True
'   ShowNoData
'   Screen.MousePointer = vbDefault
'     tmpBol = fnCancelNowFormAndShowParentForm(Me)
'             Exit Sub
'End If
'grdDataList.Visible = False 'Add by Morgan 2007/10/17
'If adoRecordset.RecordCount <> 0 Then
'    grdDataList.row = 1
'
'    'Modify By Sindy 2011/6/23
'    AddCboName Combo1, "" & adoRecordset.Fields(18).Value, "" & adoRecordset.Fields(19).Value, "" & adoRecordset.Fields(20).Value
'
'    If IsNull(adoRecordset.Fields(21)) Then
'        Label4.Caption = ""
'    Else
'        Label4.Caption = adoRecordset.Fields(21)
'    End If
'    If IsNull(adoRecordset.Fields(23)) Then
'        Me.lblClose.Caption = ""
'    Else
'        Me.lblClose.Caption = "已閉卷"
'    End If
'    'add by nickc 2006/08/25
'    If IsNull(adoRecordset.Fields("IsCancel")) Then
'        Me.lblCancel.Caption = ""
'    Else
'        Me.lblCancel.Caption = "已銷卷"
'    End If
'End If
'grdDataList.Rows = adoRecordset.RecordCount + 1
''add  by nickc 2008/03/05
'grdDataList.FixedCols = 0
'
'Set grdDataList.Recordset = adoRecordset
'SetDataListWidth
''分析字串  因為繳費年度與是否雙倍為長字串     只有專利有
'If Str01 = "CFP" Or Str01 = "FCP" Or Str01 = "P" Then
'    '宣告字串
'    Dim StrStr(4) As String, StrStr1 As Variant, StrStr2 As Variant, StrStr3 As Variant, StrInt1 As Integer
'    Dim StrIntT1 As Integer, StrIntT2 As Integer, StrIntT3 As Integer
'    '檢查資料筆數
'    For i = 1 To adoRecordset.RecordCount      'GRDDATALIST.ROWS
'        grdDataList.row = i
'        'Add by Morgan 2007/10/17
'        'If Me.GrdDataList.TextMatrix(i, 4) <> "" Then
'           Me.grdDataList.TextMatrix(i, 3) = Me.grdDataList.TextMatrix(i, 3) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(i, 2), "1")
'        'End If
'        'end 2007/10/17
'        grdDataList.col = 9
'        If IsNull(grdDataList.Text) Then
'            StrStr(0) = grdDataList.Text
'        Else
'            StrStr(0) = grdDataList.Text
'        End If
'        grdDataList.col = 13
'        If IsNull(grdDataList.Text) Then
'            StrStr(1) = ""
'            StrStr(2) = ""
'        Else
'            StrStr(1) = Left$(grdDataList.Text, InStr(1, grdDataList.Text, "-") - 1)
'            StrStr(2) = Right$(grdDataList.Text, Len(grdDataList.Text) - (InStr(1, grdDataList.Text, "-")))
'            StrStr1 = Split(StrStr(1), ",")
'            StrStr2 = Split(StrStr(2), ",")
'        End If
'        grdDataList.col = 14
'        If IsNull(grdDataList.Text) Then
'            StrStr(3) = ""
'        Else
'            StrStr(3) = grdDataList.Text
'            StrStr3 = Split(StrStr(3), ",")
'        End If
'        StrIntT1 = UBound(StrStr1)
'        StrIntT2 = UBound(StrStr2)
'        StrIntT3 = UBound(StrStr3)
'        grdDataList.col = 13
'        grdDataList.Text = ""
'        grdDataList.col = 14
'        grdDataList.Text = ""
'        For j = 0 To StrIntT1
'            If StrStr1(j) = StrStr(0) Then
'                grdDataList.col = 13
'                grdDataList.Text = grdDataList.Text + StrStr2(j) + ","
'                grdDataList.col = 14
'                grdDataList.Text = grdDataList.Text + StrStr3(j) + ","
'            End If
'        Next j
'        '收款情形
'        IntTemp1 = 0
'        IntTemp2 = 0
'        Me.grdDataList.col = 25
'        If Not IsNull(grdDataList.Text) And grdDataList.Text <> "" Then
'            '2009/12/8 modify by sonia 加請款單
'            'strSQL = "select A0k06,A0K07,'','',A0K17,A0K18 FROM ACC0K0 WHERE A0K01='" & grdDataList.Text & "'"
'            'Modify by Morgan 2011/8/12 收據的收款情形改判斷CP79
'            'If Mid(grdDataList.Text, 1, 1) = "E" Then
'            '   strSql = "select A0k06,A0K07,'','',A0K17,A0K18 FROM ACC0K0 WHERE A0K01='" & grdDataList.Text & "'"
'            'Else
'            If Mid(grdDataList.Text, 1, 1) = "X" Then
'            'end 2011/8/15
'               'modify by sonia 2017/4/18 +A1K25銷帳單號FCP-48310之AA6012573更正X10604961已銷帳
'               strSql = "select A1k11,0,'','',decode(a1k29,'Y',a1k11,nvl(A1K30,0)),0,A1K25 FROM ACC1K0 WHERE A1K01='" & grdDataList.Text & "'"
'
'            'End If 'Remove by Morgan 2011/8/15
'            '2009/12/8 end
'               CheckOC2
'               adoRecordset1.CursorLocation = adUseClient
'               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'               If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                   If Not IsNull(adoRecordset1.Fields(0)) Then
'                       IntTemp1 = IntTemp1 + adoRecordset1.Fields(0)
'                   End If
'                   If Not IsNull(adoRecordset1.Fields(1)) Then
'                       IntTemp1 = IntTemp1 + adoRecordset1.Fields(1)
'                   End If
'                   If Not IsNull(adoRecordset1.Fields(4)) Then
'                       IntTemp2 = IntTemp2 + adoRecordset1.Fields(4)
'                   End If
'                   If Not IsNull(adoRecordset1.Fields(5)) Then
'                       IntTemp2 = IntTemp2 + adoRecordset1.Fields(5)
'                   End If
'                   If IntTemp1 = IntTemp2 Then
'                        grdDataList.Text = "收回"
'                   Else
'                        If IntTemp2 = 0 Then
'                            grdDataList.Text = "未收"
'                            If "" & adoRecordset1.Fields(6) <> "" Then grdDataList.Text = "銷帳"    'add by sonia 2017/4/18 +A1K25銷帳單號FCP-48310之AA6012573更正X10604961已銷帳
'                        Else
'                            If IntTemp1 > IntTemp2 Then
'                                grdDataList.Text = "部分收回"
'                            End If
'                        End If
'                    End If
'               Else
'                    'grdDataList.Text = "查無此收據編號"   '2010/3/18 CANCEL BY SONIA
'               End If
'            End If 'Add by Morgan 2011/8/15
'        End If
'    Next i
'Else
'    For i = 1 To Me.grdDataList.Rows - 1
'        grdDataList.row = i
'        'add by nickc 2008/03/05
'        Me.grdDataList.TextMatrix(i, 3) = Me.grdDataList.TextMatrix(i, 3) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(i, 2), "1")
'
'        '收款情形
'        IntTemp1 = 0
'        IntTemp2 = 0
'        Me.grdDataList.col = 25
'        If Not IsNull(grdDataList.Text) Then
'            '2009/12/8 modify by sonia 加請款單
'            'strSQL = "select A0k06,A0K07,'','',A0K17,A0K18 FROM ACC0K0 WHERE A0K01='" & grdDataList.Text & "'"
'            'Modify by Morgan 2011/8/18 收據的收款情形改判斷CP79
'            'If Mid(grdDataList.Text, 1, 1) = "E" Then
'            '   strSql = "select A0k06,A0K07,'','',A0K17,A0K18 FROM ACC0K0 WHERE A0K01='" & grdDataList.Text & "'"
'            'Else
'            If Mid(grdDataList.Text, 1, 1) = "X" Then
'            'end 2011/8/18
'               'modify by sonia 2017/4/18 +A1K25銷帳單號FCP-48310之AA6012573更正X10604961已銷帳
'               strSql = "select A1k11,0,'','',decode(a1k29,'Y',a1k11,nvl(A1K30,0)),0,A1K25 FROM ACC1K0 WHERE A1K01='" & grdDataList.Text & "'"
'            'End If'Remove by Morgan 2011/8/18
'            '2009/12/8 end
'               CheckOC2
'               adoRecordset1.CursorLocation = adUseClient
'               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'               If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                   If Not IsNull(adoRecordset1.Fields(0)) Then
'                       IntTemp1 = IntTemp1 + adoRecordset1.Fields(0)
'                   End If
'                   If Not IsNull(adoRecordset1.Fields(1)) Then
'                       IntTemp1 = IntTemp1 + adoRecordset1.Fields(1)
'                   End If
'                   If Not IsNull(adoRecordset1.Fields(4)) Then
'                       IntTemp2 = IntTemp2 + adoRecordset1.Fields(4)
'                   End If
'                   If Not IsNull(adoRecordset1.Fields(5)) Then
'                       IntTemp2 = IntTemp2 + adoRecordset1.Fields(5)
'                   End If
'                   If IntTemp1 = IntTemp2 Then
'                        grdDataList.Text = "收回"
'                   Else
'                        If IntTemp2 = 0 Then
'                            grdDataList.Text = "未收"
'                            If "" & adoRecordset1.Fields(6) <> "" Then grdDataList.Text = "銷帳"    'add by sonia 2017/4/18 +A1K25銷帳單號FCP-48310之AA6012573更正X10604961已銷帳
'                        Else
'                            If IntTemp1 > IntTemp2 Then
'                                grdDataList.Text = "部分收回"
'                            End If
'                        End If
'                    End If
'               Else
'                    'grdDataList.Text = "查無此收據編號"   '2010/3/18 CANCEL BY SONIA
'               End If
'            End If 'Add by Morgan 2011/8/18
'        End If
'    Next i
'End If
'CheckOC
'mFixs = 0 'Added by Lydia 2018/03/22
''2005/12/19 ADD BY SONIA
''Modify By Sindy 2009/07/24 增加LIN系統類別
''modify by sonia 2019/7/30 +ACS系統類別
''Modify by Amy 2020/11/13 原:And Str01 <> "ACS"
'If (Str01 <> "L" And Str01 <> "FCL" And Str01 <> "CFL" And Str01 <> "LA" And Str01 <> "LIN") Or Str01 = "ACS" Then
'   'Modifie by Lydia 2018/03/16 固定從收文日到承辦人共5個欄位 (by 陳金蓮)
'   'grdDataList.FixedCols = 0
'   grdDataList.FixedCols = 6
'   mFixs = 5 'Added by Lydia 2018/03/22
'Else
'   grdDataList.FixedCols = 5
'   mFixs = 4 'Added by Lydia 2018/03/22
'End If
''2005/12/19 END
'
''若僅有一筆資料
'If Me.grdDataList.Rows = 2 Then
'   '直接勾選
'   'grdDataList.Visible = False 'Remove by Morgan 2007/10/17 在外層控制
'   grdDataList.col = 0
'   grdDataList.row = 1
'   grdDataList.Text = "V"
'   For i = 0 To grdDataList.Cols - 1
'       grdDataList.col = i
'       grdDataList.CellBackColor = &HFFC0C0
'   Next i
'   'grdDataList.Visible = True 'Remove by Morgan 2007/10/17 在外層控制
'End If
'
'''Add by Morgan 2011/1/3 日期欄位靠右
''grdDataList.ColAlignment(1) = flexAlignRightCenter
''grdDataList.ColAlignment(6) = flexAlignRightCenter
''grdDataList.ColAlignment(7) = flexAlignRightCenter
''grdDataList.ColAlignment(8) = flexAlignRightCenter
''grdDataList.ColAlignment(9) = flexAlignRightCenter
'''end 2011/1/3
'Call SetDataListWidth_new 'Modify By Sindy 2011/2/25
'
''Added by Lydia 2015/11/03　顯示一案兩請，擬制喪失新穎性案件
'lblCaseMap.Caption = ""
'lblCaseMap2.Caption = "" 'Added by Lydia 2019/11/28
'If PUB_GetRefCaseChk(Str01, Str02, Str03, Str04, "CASEMAP", "3") = True Then
'   lblCaseMap.Caption = "一案兩請"
'End If
'If PUB_GetRefCaseChk(Str01, Str02, Str03, Str04, "CASEMAP", "6") = True Then
'      'Modified by Lydia 2019/11/28 P-123733有一案兩請和擬制喪失新穎性案件
'      'If lblCaseMap.Caption <> "" Then lblCaseMap.Caption = lblCaseMap.Caption & "   "
'      'lblCaseMap.Caption = lblCaseMap.Caption & "擬制喪失新穎性案件"
'      lblCaseMap2.Caption = "擬制喪失新穎性案件"
'End If
''end 2015/11/03
'
''Added by Lydia 2016/06/14 +台灣大陸案件提示
'lblCMboth.Caption = ""
'strExc(9) = GetPrjNation1(Label3.Caption)
'If (Str01 = "P" Or Str01 = "FCP") And strExc(9) = 台灣國家代號 Then
'   If PUB_GetRefCaseChk(Str01, Str02, Str03, Str04, "CASEMAP", "0", "A", 大陸國家代號) Then
'      lblCMboth.Caption = "有大陸案"
'   End If
'ElseIf Str01 = "P" And strExc(9) = 大陸國家代號 Then
'   If PUB_GetRefCaseChk(Str01, Str02, Str03, Str04, "CASEMAP", "0", "A", 台灣國家代號) Then
'      lblCMboth.Caption = "有台灣案"
'   End If
'End If
''end 2016/06/14
'
'lblNp605.Caption = GetNp605State(Str01, Str02, Str03, Str04) 'Added by Lydia 2020/01/06 專利案件：年費不續辦
'
'grdDataList.Visible = True 'Add by Morgan 2007/10/17
'Me.Enabled = True
'End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   cmdState = -1
   'Add By Sindy 2013/6/11
   'If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
      'cmdOK(13).Visible = True '卷宗區 'Modify By Sindy 2015/3/9 Mark:不分系統別都可以查看
      'Modify By Sindy 2017/8/16
      'If Pub_StrUserSt03 = "M51" Or Left(Pub_StrUserSt03, 2) = "P1" Then
      '   cmdOK(14).Visible = True '原始檔
      If Left(Pub_StrUserSt03, 1) = "S" Then
         cmdOK(14).Visible = False
      Else
         cmdOK(14).Visible = True '原始檔
      End If
      cmdOK(15).Visible = True
   'Else
   '   cmdOK(13).Visible = False
   '   cmdOK(14).Visible = False
   '   cmdOK(15).Visible = False
   'End If
   '2013/6/11 End
   
   'Added by Morgan 2019/6/20
   If (Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "P12") Then
      cmdCustAtt.Visible = True
      cmdOK(18).Top = cmdCustAtt.Top 'Added by Lydia 2025/09/10
   Else
      cmdCustAtt.Visible = False
      'Added by Lydia 2025/09/10
      cmdOK(18).Top = cmdCustAtt.Top
      cmdOK(18).Left = cmdCustAtt.Left
      'end 2025/09/10
   End If
   'end 2019/6/20
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Added by Lydia 2025/09/10
    If TypeName(m_PrevForm) = "frm060209" Then
       '避免重覆呼叫，同時關閉"行事曆+案件進度查詢"
       'Call m_PrevForm.FormEXIT
       Call m_PrevForm.Show
       Debug.Print Format(ServerTime, "000000") & ":Show"
    'Else
    'end 2025/09/10
    End If
       Set frm100101_2 = Nothing
    'End If '1234
End Sub

Private Sub grdDataList_SelChange()

grdDataList.Visible = False
grdDataList.col = 0
grdDataList.row = grdDataList.MouseRow
If grdDataList.MouseRow <> 0 Then
If grdDataList.Text = "V" Then
     grdDataList.Text = ""
     For i = 0 To grdDataList.Cols - 1
          'Modified by Lydia 2018/03/22 固定欄位不變色
          'If i > 5 Or Label1.Visible = True Then
          If i > mFixs Or mFixs = 0 Then
               grdDataList.col = i
               grdDataList.CellBackColor = QBColor(15)
          End If
    Next i
Else
     grdDataList.Text = "V"
     For i = 0 To grdDataList.Cols - 1
        'Modified by Lydia 2018/03/22 固定欄位不變色
        'If i > 5 Or Label1.Visible = True Then
        If i > mFixs Or mFixs = 0 Then
             grdDataList.col = i
             grdDataList.CellBackColor = &HFFC0C0
        End If
     Next i
End If
End If
grdDataList.Visible = True
End Sub

'Add By Sindy 2011/6/23
'取得正確的 row & col
Public Sub getGrdColRow(ByRef oObj As MSHFlexGrid, ByVal x As Single, ByVal y As Single, ByRef col As Long, ByRef row As Long)
Dim nIndex As Integer
col = 0: row = 0
For nIndex = 0 To oObj.Rows - 1
    If y > oObj.RowHeight(nIndex) Then
        row = row + 1
        y = y - oObj.RowHeight(nIndex)
    ElseIf y > 0 Then
        row = row + 1
        Exit For
    End If
Next nIndex
For nIndex = 0 To oObj.Cols - 1
    If x > oObj.ColWidth(nIndex) Then
        col = col + 1
        x = x - oObj.ColWidth(nIndex)
    ElseIf x > 0 Then
        col = col + 1
        Exit For
    End If
Next nIndex
col = col - 1 + IIf(oObj.LeftCol <> oObj.FixedCols And oObj.LeftCol <> 0, oObj.LeftCol - oObj.FixedCols, 0)
row = row - 1 + IIf(oObj.TopRow <> oObj.FixedRows And oObj.TopRow <> 0, oObj.TopRow - oObj.FixedRows, 0)

If col > oObj.Cols - 1 Then col = oObj.Cols - 1
If row > oObj.Rows - 1 Then row = oObj.Rows - 1
End Sub

'Add By Sindy 2011/6/23
Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow grdDataList, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   grdDataList.col = nCol
   grdDataList.row = nRow
   If Me.grdDataList.row < 1 And Me.grdDataList.Text <> "V" Then
      If Me.grdDataList.Text = "點數" Then
         If m_blnColOrderAsc = True Then
            Me.grdDataList.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.grdDataList.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.grdDataList.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.grdDataList.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

'Added by Lydia 2020/01/06 若專利案件的年費期限為不續辦，則在畫面右上角加註「年費不續辦」。
Private Function GetNp605State(ByVal mPA01 As String, ByVal mPA02 As String, ByVal mPA03 As String, ByVal mPA04 As String) As String

    GetNp605State = ""

    If mPA01 = "CFP" Or mPA01 = "FCP" Or mPA01 = "P" Then
       'Modified by Morgan 2024/6/11 只要判斷年費就好,這樣與訊息一致才不會誤解,如有需要再增加其他性質,且香港案分階段管控原規則會顯示錯誤狀態 Ex:P-130612--Anny
       'strExc(0) = "SELECT NP06 FROM NEXTPROGRESS WHERE " & ChgNextProgress(mPA01 & mPA02 & mPA03 & mPA04) & " AND (NP09||NP22) IN " & _
                        "(SELECT MAX(NP09||NP22) FROM NEXTPROGRESS WHERE " & ChgNextProgress(mPA01 & mPA02 & mPA03 & mPA04) & _
                        " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "') And NP09 IS NOT NULL)"
       strExc(0) = "SELECT NP06 FROM NEXTPROGRESS WHERE " & ChgNextProgress(mPA01 & mPA02 & mPA03 & mPA04) & " AND (NP09||NP22) IN " & _
                        "(SELECT MAX(NP09||NP22) FROM NEXTPROGRESS WHERE " & ChgNextProgress(mPA01 & mPA02 & mPA03 & mPA04) & _
                        " AND NP07='" & 年費 & "' And NP09 IS NOT NULL)"
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
            If "" & RsTemp.Fields("NP06") = "N" Then
                GetNp605State = "年費不續辦"
            End If
       End If
    End If
End Function

'Added by Lydia 2021/05/28 可選多筆，當顯示資料後將前一畫面的勾選項取消。
Public Sub UpdateShowFlag(ByVal pRow As Integer)
Dim intR As Integer
       
    grdDataList.TextMatrix(pRow, 0) = ""
    grdDataList.row = pRow
    For intR = 0 To grdDataList.Cols - 1
        If intR > mFixs Or mFixs = 0 Then
            grdDataList.col = intR
            grdDataList.CellBackColor = QBColor(15) '還原顏色
        End If
    Next intR
End Sub

