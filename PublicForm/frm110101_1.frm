VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm110101_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "解除期限"
   ClientHeight    =   5772
   ClientLeft      =   72
   ClientTop       =   948
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   9324
   Begin VB.TextBox txtNumber2 
      Height          =   285
      Left            =   1044
      TabIndex        =   8
      Top             =   720
      Width           =   3900
   End
   Begin VB.TextBox txtNumber1 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   1050
      Width           =   3165
   End
   Begin VB.Frame fraElse 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1764
      TabIndex        =   15
      Top             =   396
      Width           =   2532
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   0
         Left            =   0
         MaxLength       =   6
         TabIndex        =   1
         Top             =   15
         Width           =   1212
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   1
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   2
         Top             =   15
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   2
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   3
         Top             =   15
         Width           =   492
      End
   End
   Begin VB.TextBox txtSystem 
      Height          =   264
      Left            =   1044
      MaxLength       =   3
      TabIndex        =   0
      Top             =   408
      Width           =   732
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "尋找(&F)"
      Height          =   400
      Index           =   2
      Left            =   4392
      TabIndex        =   12
      Top             =   300
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7488
      TabIndex        =   13
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8388
      TabIndex        =   14
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3756
      Left            =   72
      TabIndex        =   11
      Top             =   1968
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   6625
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
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
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fraTF 
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1764
      TabIndex        =   16
      Top             =   408
      Visible         =   0   'False
      Width           =   2412
      Begin VB.TextBox txtTFCode 
         Height          =   264
         Index           =   3
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   7
         Top             =   0
         Width           =   492
      End
      Begin VB.TextBox txtTFCode 
         Height          =   264
         Index           =   2
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   6
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtTFCode 
         Height          =   264
         Index           =   1
         Left            =   960
         MaxLength       =   1
         TabIndex        =   5
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtTFCode 
         Height          =   264
         Index           =   0
         Left            =   0
         MaxLength       =   5
         TabIndex        =   4
         Top             =   0
         Width           =   852
      End
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1050
      TabIndex        =   10
      Top             =   1380
      Width           =   8175
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "14420;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Left            =   84
      TabIndex        =   22
      Top             =   444
      Width           =   972
   End
   Begin VB.Label Label5 
      Caption         =   "申請案號："
      Height          =   180
      Left            =   75
      TabIndex        =   21
      Top             =   735
      Width           =   975
   End
   Begin VB.Label lblNumber 
      Caption         =   "審定號數/證書號數："
      Height          =   180
      Left            =   75
      TabIndex        =   20
      Top             =   1095
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "申請人："
      Height          =   180
      Left            =   75
      TabIndex        =   19
      Top             =   1695
      Width           =   885
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   0
      Left            =   75
      TabIndex        =   18
      Top             =   1380
      Width           =   900
   End
   Begin MSForms.Label lblAgent 
      Height          =   225
      Left            =   1050
      TabIndex        =   17
      Top             =   1710
      Width           =   8175
      Size            =   "14420;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm110101_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/5/5 改成Form2.0 (cboCaseName,lblAgent,grdDataList)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'sonia 2010/8/19 日期欄已修改
Option Explicit

'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
'edit by nickc 2007/02/06 不用 dll 了
'Dim obj011 As Object
Dim strNation As String
Private Const intField As Integer = 13 'Added by Lydia 2020/01/09 Grid欄位數


Private Function CheckCPOk() As Boolean
   'Modify By Sindy 2020/1/8
   'Modify By Sindy 2023/6/26 + , IIf(grdDataList.TextMatrix(grdDataList.row, 12) = "Y", True, False) : bolIsFMP=True.是
   CheckCPOk = CheckFlowCloseOk(txtSystem.Text, txtCode(0), Right("0" & txtCode(1), 1), Right("00" & txtCode(2), 2), _
               grdDataList.TextMatrix(grdDataList.row, 9), IIf(grdDataList.TextMatrix(grdDataList.row, 12) = "Y", True, False))
   
'   'Add by Morgan 2010/8/31
'   '大陸的標準專利紀錄請求除外
'   If txtSystem.Text = "P" And grdDataList.TextMatrix(grdDataList.row, 9) = "110" Then
'      CheckCPOk = True
'      Exit Function
'   End If
'
'   'Add by Morgan 2007/1/22
'   If txtSystem.Text = "P" Or txtSystem.Text = "PS" Or txtSystem.Text = "CFP" Or txtSystem.Text = "CPS" Then
'     strExc(0) = "select cp09 from caseprogress where cp01='" & txtSystem.Text & "' and cp02='" & txtCode(0) & "' and cp03='" & Right("0" & txtCode(1), 1) & "' and cp04='" & Right("00" & txtCode(2), 2) & "' and cp27 is null and cp57 is null and cp09>'B'"
'     intI = 1
'     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'     If intI = 1 Then
'         '有B,C類收文未發文
'         MsgBox "有B,C類收文未發文不可解除期限！"
'     Else
'         CheckCPOk = True
'     End If
'   Else
'      CheckCPOk = True
'   End If
'   'end 2007/1/122
End Function

Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer, bolRt As Boolean
Dim stTi06 As String 'Add by Amy 2024/09/06

Select Case Index
             Case 0
                        'Add by Morgan 2010/7/15
                        If txtNumber1.Locked = False And txtNumber1 <> "" And txtNumber1 <> txtNumber1.Tag Then
                           MsgBox "證書號數錯誤，請重新輸入！"
                           txtNumber1_GotFocus
                           txtNumber1.SetFocus
                           Exit Sub
                        ElseIf txtNumber2.Locked = False And txtNumber2 <> "" And txtNumber2 <> txtNumber2.Tag Then
                           MsgBox "申請案號錯誤，請重新輸入！"
                           txtNumber2_GotFocus
                           txtNumber2.SetFocus
                           Exit Sub
                        ElseIf txtNumber1.Locked = False And txtNumber1 & txtNumber2 = "" Then
                        'Add by Lydia 2014/10/14 FMP案(補正203和第一階段請求110)可不必輸入申請案號和證書,預設不印指示信和不詢問是否要作結餘
                         If Not (txtSystem.Text = "P" And grdDataList.TextMatrix(grdDataList.row, 12) = "Y" And (grdDataList.TextMatrix(grdDataList.row, 9) = "110" Or grdDataList.TextMatrix(grdDataList.row, 9) = "203")) Then
                           MsgBox "請輸入申請案號或證書號數！"
                           txtNumber2.SetFocus
                           Exit Sub
                         End If
                        End If
                       'end 2010/7/15
                       
                       'Add by Amy 2024/09/06 T延展/第二期註冊費,退回智權/取消延展結案(ti06=Y),不可由此再結案
                       stTi06 = ""
                       If (txtSystem.Text = "T" Or txtSystem.Text = "TF") And (grdDataList.TextMatrix(grdDataList.row, 9) = "102" Or grdDataList.TextMatrix(grdDataList.row, 9) = "716") Then
                           stTi06 = "Ti06"
                           Call ChkT102Inform(grdDataList.TextMatrix(grdDataList.row, 0), grdDataList.TextMatrix(grdDataList.row, 10), stTi06)
                           If stTi06 = "Y" Then
                              MsgBox "此程序已退回智權,不可由此再結案！"
                              Exit Sub
                           End If
                       End If
                        
                        'Modify by Morgan 2007/1/23 加檢查是否有B,C類收文未發文
                        If CheckCPOk Then
                          'Add by Lydia 2014/10/14 FMP案
                           Set frm110101_2.mPrev01 = Me
                           frm110101_2.Show
                           Me.Hide
                        End If
             Case 1
                        Unload Me
             Case 2
                        If txtSystem = 馬德里案 Then
                           bolRt = CheckKeyIn1(3)
                        Else
                           bolRt = CheckKeyIn2(2)
                        End If
                        If bolRt Then
                           GetRelieveDeadlineCaseData
                        End If
End Select
End Sub

Private Sub GetRelieveDeadlineCaseData(Optional bolBlank As Boolean = False)
'TF為馬德里案，另外判斷
If bolBlank = True Then
   GetRelieveDeadlineData "", "", "", ""
Else
   If txtSystem = 馬德里案 Then
      GetRelieveDeadlineData txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3))
   Else
      GetRelieveDeadlineData txtSystem, txtCode(0), _
         IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))
   End If
End If
End Sub

Private Sub GetRelieveDeadlineData(ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String)
Dim varSaveCursor

varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
Set grdDataList.Recordset = ReadRelieveDeadlineRst(strCode1, strCode2, strCode3, strCode4)

SetGridHead
SetDataListVision grdDataList

intLastRow = 0
If grdDataList.Rows > 1 Then
   'Modified by Lydia 2020/01/09 改變數
   'ShowBar GrdDataList, intLastRow, 11
   ShowBar grdDataList, intLastRow, intField - 1
   cmdOK(0).Enabled = True
   cmdOK(0).Default = True
   If grdDataList.Rows = 2 Then
      grdDataList.row = 1
      grdDataList.col = 0
      If Len(grdDataList.Text) <> 0 Then
         cmdOK_Click (0)
      End If
   End If
Else
   cmdOK(0).Enabled = False
   cmdOK(2).Default = True
End If
Screen.MousePointer = varSaveCursor
End Sub

Private Sub SetGridHead()
  With grdDataList
       .row = 0
       .Cols = 13 'Add by Lydia 2014/10/14 12->13
       .col = 0
       .ColWidth(0) = 1000
       .col = 1
       .ColWidth(1) = 1000
       .col = 2
       .ColWidth(2) = 1200
       .col = 3
       .ColWidth(3) = 900
       .col = 4
       .ColWidth(4) = 900
       .col = 5
       .ColWidth(5) = 700
       .col = 6
       .ColWidth(6) = 700
       .col = 7
       .ColWidth(7) = 1000
       .col = 8
       .ColWidth(8) = 1000
       .col = 9
       .ColWidth(9) = 0
       .col = 10
       .ColWidth(10) = 0
       .col = 11
       .ColWidth(11) = 1500
       'Add by Lydia 2014/10/14 FMP案(補正203和第一階段請求110)可不必輸入申請案號和證書,預設不印指示信和不詢問是否要作結餘
       .col = 12
       .ColWidth(12) = 0
  End With
End Sub

Private Sub SetDataListWidth()
Dim varGridWidth() As Variant
'Add by Lydia 2014/10/14 FMP案(補正203和第一階段請求110)可不必輸入申請案號和證書,預設不印指示信和不詢問是否要作結餘
'varGridWidth = Array(1000, 1000, 1200, 900, 900, 700, 700, 1000, 1000, 0, 0, 1500)
varGridWidth = Array(1000, 1000, 1200, 900, 900, 700, 700, 1000, 1000, 0, 0, 1500, 0)
SetGridDataListWidth grdDataList, varGridWidth()
blnOKtoShow = True
End Sub

Private Sub Form_Activate()
   'GetRelieveDeadlineCaseData
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
SetDataListWidth
'edit by nickc 2007/02/06 不用 dll 了
'If obj011 Is Nothing Then
'   Set obj011 = CreateObject("prjTaieDll011.cls011")
'   Set obj011.Connection = cnnConnection
'End If
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm110101_1 = Nothing
End Sub

Private Sub grdDataList_DblClick()
cmdOK_Click 0
End Sub

Private Sub txtNumber1_GotFocus()
   TextInverse txtNumber1
End Sub

Private Sub txtNumber2_GotFocus()
   TextInverse txtNumber2
   CloseIme
End Sub

Private Sub txtSystem_Change()
If txtSystem.Text = 馬德里案 Then
   fraTF.Visible = True
   fraElse.Visible = False
Else
   fraTF.Visible = False
   fraElse.Visible = True
End If
If cboCaseName.ListCount > 0 Then
   txtNumber1 = "" 'Add By Sindy 2018/1/18
   txtNumber2 = "" 'Add By Sindy 2018/1/18
   cboCaseName.Clear
End If
If grdDataList.Rows > 1 Then GetRelieveDeadlineCaseData True
End Sub

Private Sub txtSystem_GotFocus()
txtSystem.SelStart = 0
txtSystem.SelLength = Len(txtSystem.Text)
CloseIme
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSystem_Validate(Cancel As Boolean)
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.GetGroupCase(txtSystem, strGroup) = False Then
'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件(P)，但非此類案件時外專程序人員不可操作。
If Not (FMP2open = True And (txtSystem.Text = "P" Or txtSystem.Text = "PS")) Then
    If ClsPDGetGroupCase(txtSystem, strGroup) = False Then
       ShowMsg MsgText(1056)
       Cancel = True
       txtSystem_GotFocus
    End If
End If
End Sub

Private Sub txtCode_Change(Index As Integer)
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
If grdDataList.Rows > 1 Then GetRelieveDeadlineCaseData True
End Sub

Private Sub txtTFCode_Change(Index As Integer)
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
If grdDataList.Rows > 1 Then GetRelieveDeadlineCaseData True
End Sub

Private Sub txtTFCode_GotFocus(Index As Integer)
txtTFCode(Index).SelStart = 0
txtTFCode(Index).SelLength = Len(txtTFCode(Index).Text)
End Sub

Private Sub txtTFCode_Validate(Index As Integer, Cancel As Boolean)
CheckKeyIn1 (Index)
End Sub

Private Function CheckKeyIn1(ByRef intIndex As Integer) As Boolean
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
Dim strCustomer As String, strNumber1 As String, strNumber2 As String

If Len(txtTFCode(intIndex)) > 0 And Len(txtTFCode(intIndex)) < txtTFCode(intIndex).MaxLength Then
   ShowMsg MsgText(33)
ElseIf intIndex = 3 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) Then
'edit by nickc 2008/05/16 加抓國家
'   If ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) Then
      strNation = ""
  'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   Dim mOKchk As Boolean
   If FMP2open = False Then
      mOKchk = ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation, strNumber1, strNumber2)
   Else
      mOKchk = PUB_FMPtoCheck(0, 1, Pub_strUserST05, txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)))
      If mOKchk = True Then '借由另一模組取值
         mOKchk = ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation, strNumber1, strNumber2)
      End If
   End If
    If mOKchk = True Then
   'end 'Add by Lydia 2014/10/31
 '  If ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation, strNumber1, strNumber2) Then
      SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
      'Modify by Morgan 2010/7/15 FCP,P,CFP 要輸入申請號或證書號
      'lblNumber1 = strNumber1
      'lblNumber2 = strNumber2
      txtNumber1.Tag = strNumber1
      txtNumber2.Tag = strNumber2
      If txtSystem = "FCP" Or txtSystem = "P" Or txtSystem = "CFP" Then
         txtNumber1.Locked = False
         txtNumber2.Locked = False
         txtNumber1 = ""
         txtNumber2 = ""
      Else
         txtNumber1.Locked = True
         txtNumber2.Locked = True
         txtNumber1 = strNumber1
         txtNumber2 = strNumber2
      End If
      'end 2010/7/15
      lblAgent = strCustomer
      CheckKeyIn1 = True
   End If
Else
   CheckKeyIn1 = True
End If
End Function

Private Sub txtCode_GotFocus(Index As Integer)
txtCode(Index).SelStart = 0
txtCode(Index).SelLength = Len(txtCode(Index).Text)
CloseIme
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
CheckKeyIn2 (Index)
End Sub

Private Function CheckKeyIn2(ByRef intIndex As Integer) As Boolean
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
Dim strCustomer As String, strNumber1 As String, strNumber2 As String

If Len(txtCode(intIndex)) > 0 And Len(txtCode(intIndex)) < txtCode(intIndex).MaxLength Then
   ShowMsg MsgText(33)
ElseIf intIndex = 2 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) Then
'edit by nickc 2008/05/16 加抓國家
'   If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) Then
   strNation = ""
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   Dim mOKchk As Boolean
   If FMP2open = False Then
      mOKchk = ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation, strNumber1, strNumber2)
   Else
      mOKchk = PUB_FMPtoCheck(0, 1, Pub_strUserST05, txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
      If mOKchk = True Then '借由另一模組取值
         mOKchk = ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation, strNumber1, strNumber2)
      End If
   End If
   If mOKchk = True Then
   'end 'Add by Lydia 2014/10/31
   'If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, strNation, strNumber1, strNumber2) Then
      SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
      'Modify by Morgan 2010/7/15 FCP,P,CFP 要輸入申請號或證書號
      'lblNumber1 = strNumber1
      'lblNumber2 = strNumber2
      txtNumber1.Tag = strNumber1
      txtNumber2.Tag = strNumber2
      If txtSystem = "FCP" Or txtSystem = "P" Or txtSystem = "CFP" Then
         txtNumber1.Locked = False
         txtNumber2.Locked = False
      Else
         txtNumber1.Locked = True
         txtNumber2.Locked = True
         txtNumber1 = strNumber1
         txtNumber2 = strNumber2
      End If
      'end 2010/7/15
      
      lblAgent = strCustomer
      CheckKeyIn2 = True
   End If
Else
   CheckKeyIn2 = True
End If
End Function

Private Sub grdDataList_GotFocus()
GridGotFocus grdDataList
End Sub

Private Sub grdDataList_LostFocus()
GridLostFocus grdDataList
End Sub

Private Sub grdDataList_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then grdDataList_DblClick
End Sub
Private Sub grdDataList_RowColChange()
If intLastRow <> grdDataList.row Then
   If blnOKtoShow Then
      blnOKtoShow = False
      'Modified by Lydia 2020/01/09 改變數
      'ShowBar GrdDataList, intLastRow, 8
      ShowBar grdDataList, intLastRow, intField - 1
      blnOKtoShow = True
   End If
End If
End Sub

'讀取解除期限資料
Private Function ReadRelieveDeadlineRst(ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String) As ADODB.Recordset
  Dim strSql As String
  Dim rsRecordset As New ADODB.Recordset
  'edit by nickc 2007/02/06 不用 dll 了
  'Dim objPublicData As Object
  'edit by nickc 2007/02/06 不用 dll 了
  'Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
    'Modify By Cheng 2002/12/13
    '以收文日由大至小排序
'  strSQL = "select NP01 總收文號,substr(cp05,1,4)-1911||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) 來函收文日,decode(np07," + CNULL(大陸國家代號) + ",cpm04,cpm03) 下一程序,decode(np08,null,'',substr(np08,1,4)-1911||'/'||substr(np08,5,2)||'/'||substr(np08,7,2)) 本所期限,decode(np09,null,'',substr(np09,1,4)-1911||'/'||substr(np09,5,2)||'/'||substr(np09,7,2)) 法定期限,staff.st02 承辦人,staff1.st02 智權人員,np13 機關文號,np14 相關人,np07,NP22 " & _
'           "from caseprogress,nextprogress,casepropertymap,staff,staff staff1 where np02=cpm01(+) and np07=cpm02(+) and cp14=staff.st01(+) and np10=staff1.st01(+) and NP02=CP01 and np03=cp02 and np04=cp03 and np05=cp04 and np06 is null and NP02=" + CNULL(strCode1) + " and NP03=" + CNULL(strCode2) + " and NP04=" + CNULL(strCode3) + " and NP05=" + CNULL(strCode4) + " and np01=CP09(+)"
  'Modify by Morgan 2007/4/10 抓進度檔時不必串本所案號否則大陸香港或一案兩請會無法解除期限 P-069874
  'strSQL = "select NP01 總收文號,substr(cp05,1,4)-1911||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) 來函收文日,decode(np07," + CNULL(大陸國家代號) + ",cpm04,cpm03) 下一程序,decode(np08,null,'',substr(np08,1,4)-1911||'/'||substr(np08,5,2)||'/'||substr(np08,7,2)) 本所期限,decode(np09,null,'',substr(np09,1,4)-1911||'/'||substr(np09,5,2)||'/'||substr(np09,7,2)) 法定期限,staff.st02 承辦人,staff1.st02 智權人員,np13 機關文號,np14 相關人,np07,NP22,NP15 備註 " & _
           "from caseprogress,nextprogress,casepropertymap,staff,staff staff1 where np02=cpm01(+) and np07=cpm02(+) and cp14=staff.st01(+) and np10=staff1.st01(+) and NP02=CP01 and np03=cp02 and np04=cp03 and np05=cp04 and np06 is null and NP02=" + CNULL(strCode1) + " and NP03=" + CNULL(strCode2) + " and NP04=" + CNULL(strCode3) + " and NP05=" + CNULL(strCode4) + " and np01=CP09(+) Order By CP05 Desc "
  'edit by nickc 2008/05/16 案件性質判斷，原先寫錯了，應該用國家
  'strSQL = "select NP01 總收文號,substr(cp05,1,4)-1911||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) 來函收文日,decode(np07," + CNULL(大陸國家代號) + ",cpm04,cpm03) 下一程序,decode(np08,null,'',substr(np08,1,4)-1911||'/'||substr(np08,5,2)||'/'||substr(np08,7,2)) 本所期限,decode(np09,null,'',substr(np09,1,4)-1911||'/'||substr(np09,5,2)||'/'||substr(np09,7,2)) 法定期限,staff.st02 承辦人,staff1.st02 智權人員,np13 機關文號,np14 相關人,np07,NP22,NP15 備註 " & _
           "from caseprogress,nextprogress,casepropertymap,staff,staff staff1 where np02=cpm01(+) and np07=cpm02(+) and cp14=staff.st01(+) and np10=staff1.st01(+) and np06 is null and NP02=" + CNULL(strCode1) + " and NP03=" + CNULL(strCode2) + " and NP04=" + CNULL(strCode3) + " and NP05=" + CNULL(strCode4) + " and np01=CP09(+) Order By CP05 Desc "
'Add by Lydia 2014/10/14 FMP案(補正203和第一階段請求110)可不必輸入申請案號和證書,預設不印指示信和不詢問是否要作結餘
'   strSql = "select NP01 總收文號,substr(cp05,1,4)-1911||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) 來函收文日,decode(" & CNULL(strNation) & "," + CNULL(大陸國家代號) + ",cpm04,cpm03) 下一程序,decode(np08,null,'',substr(np08,1,4)-1911||'/'||substr(np08,5,2)||'/'||substr(np08,7,2)) 本所期限,decode(np09,null,'',substr(np09,1,4)-1911||'/'||substr(np09,5,2)||'/'||substr(np09,7,2)) 法定期限,staff.st02 承辦人,staff1.st02 智權人員,np13 機關文號,np14 相關人,np07,NP22,NP15 備註 " & _
 '          "from caseprogress,nextprogress,casepropertymap,staff,staff staff1 where np02=cpm01(+) and np07=cpm02(+) and cp14=staff.st01(+) and np10=staff1.st01(+) and np06 is null and NP02=" + CNULL(strCode1) + " and NP03=" + CNULL(strCode2) + " and NP04=" + CNULL(strCode3) + " and NP05=" + CNULL(strCode4) + " and np01=CP09(+) Order By CP05 Desc "
   'modify by sonia 下一程序名稱以申請國家台灣或非台灣來判斷
   'strSql = " select NP01 總收文號,substr(cp05,1,4)-1911||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) 來函收文日,decode(" & CNULL(strNation) & "," + CNULL(大陸國家代號) + ",cpm04,cpm03) 下一程序,decode(np08,null,'',substr(np08,1,4)-1911||'/'||substr(np08,5,2)||'/'||substr(np08,7,2)) 本所期限,decode(np09,null,'',substr(np09,1,4)-1911||'/'||substr(np09,5,2)||'/'||substr(np09,7,2)) 法定期限,staff.st02 承辦人,staff1.st02 智權人員,np13 機關文號,np14 相關人,np07,NP22,NP15 備註,decode(cp01||substr(cp12,1,1),'PF','Y','N') as FMP案 " &
   strSql = " select NP01 總收文號,substr(cp05,1,4)-1911||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) 來函收文日,decode(" & CNULL(strNation) & ",'000',cpm03,cpm04) 下一程序,decode(np08,null,'',substr(np08,1,4)-1911||'/'||substr(np08,5,2)||'/'||substr(np08,7,2)) 本所期限,decode(np09,null,'',substr(np09,1,4)-1911||'/'||substr(np09,5,2)||'/'||substr(np09,7,2)) 法定期限,staff.st02 承辦人,staff1.st02 智權人員,np13 機關文號,np14 相關人,np07,NP22,NP15 備註,decode(cp01||substr(cp12,1,1),'PF','Y','N') as FMP案 " & _
           " from caseprogress,nextprogress,casepropertymap,staff,staff staff1 " & _
           " where np02=cpm01(+) and np07=cpm02(+) and cp14=staff.st01(+) and np10=staff1.st01(+) and np06 is null " & _
           " and NP02=" + CNULL(strCode1) + " and NP03=" + CNULL(strCode2) + " and NP04=" + CNULL(strCode3) + " and NP05=" + CNULL(strCode4) + " and np01=CP09(+) " & _
           " Order By CP05 Desc "
  'end 2007/4/10
  'edit by nickc 2007/02/06 不用 dll 了
  'Set ReadRelieveDeadlineRst = objPublicData.ReadRst(strSQL)
  'Set objPublicData = Nothing
   Set ReadRelieveDeadlineRst = ClsPDReadRst(strSql)

End Function

Public Sub Cleartxt()
txtCode(0) = ""
txtCode(1) = ""
txtSystem = ""
txtCode(2) = ""
cboCaseName.Clear
'Modify by Morgan 2010/7/15
'lblNumber1.Caption = ""
'lblNumber2.Caption = ""
txtNumber1 = ""
txtNumber1.Tag = "" 'Add By Sindy 2018/1/18
txtNumber2 = ""
txtNumber2.Tag = "" 'Add By Sindy 2018/1/18
'end 2010/7/15
lblAgent.Caption = ""
txtSystem.SetFocus
End Sub
