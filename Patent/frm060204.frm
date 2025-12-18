VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060204 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外部專利處期限通知"
   ClientHeight    =   6690
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   10150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10150
   Begin VB.TextBox TextSys 
      Height          =   285
      Left            =   1410
      MaxLength       =   3
      TabIndex        =   2
      Top             =   810
      Width           =   410
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   9270
      TabIndex        =   5
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   6735
      TabIndex        =   1
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   2
      Left            =   7545
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   400
      Index           =   3
      Left            =   8460
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "所有未發文(&A)"
      Height          =   400
      Index           =   1
      Left            =   1195
      TabIndex        =   7
      Top             =   60
      Width           =   1290
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "隱藏白色(&H)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   45
      TabIndex        =   6
      Top             =   60
      Width           =   1140
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "事件說明(&H)"
      Height          =   400
      Left            =   2495
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   60
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "工作進度資料維護(&W)"
      Height          =   400
      Index           =   0
      Left            =   3615
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   60
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "發 E-Mail(S)"
      Height          =   400
      Index           =   1
      Left            =   4860
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   60
      Width           =   1095
   End
   Begin VB.TextBox txtUsernum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1125
      MaxLength       =   6
      TabIndex        =   0
      Top             =   510
      Width           =   915
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   5310
      Style           =   2  '單純下拉式
      TabIndex        =   12
      Top             =   510
      Width           =   4785
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5270
      Left            =   20
      TabIndex        =   11
      Top             =   1170
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   9296
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "V|本所期限|法定期限|約定期限|承辦期限|核稿期限|管制人|承辦人|事件　|本所案號　　　|案件性質|備註　　　　|案件名稱　　　"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   13
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "　　　　　* C類來函未發文"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   6870
      TabIndex        =   22
      Top             =   960
      Width           =   2220
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "符號說明：▲機械設計組"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   6860
      TabIndex        =   21
      Top             =   780
      Width           =   1980
   End
   Begin VB.Label Label5 
      Caption         =   "按 Tab鍵 過濾資料"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   3390
      TabIndex        =   20
      Top             =   870
      Width           =   1620
   End
   Begin VB.Label LblCnt 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   5100
      TabIndex        =   19
      Top             =   870
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "僅顯示系統別：          (空白則顯示全部)"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   140
      TabIndex        =   18
      Top             =   860
      Width           =   3140
   End
   Begin MSForms.Label lblUserName 
      Height          =   255
      Left            =   2100
      TabIndex        =   17
      Top             =   540
      Width           =   1710
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "3016;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblTotCnt 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   7680
      TabIndex        =   16
      Top             =   6480
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PS：已發文未請款案件的請款期限顯示於本所期限欄位"
      Height          =   180
      Left            =   105
      TabIndex        =   15
      Top             =   6480
      Width           =   4320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   135
      TabIndex        =   14
      Top             =   555
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "顏色說明："
      Height          =   180
      Left            =   4410
      TabIndex        =   13
      Top             =   570
      Width           =   900
   End
End
Attribute VB_Name = "frm060204"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/24 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lblUserName
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
'Add by Morgan 2008/6/25
Option Explicit

Dim bolBarShow As Boolean
Public cmdState As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_stSort As String '排序方式
Dim m_adoRst As ADODB.Recordset
Dim bLvlX As Boolean, bLvl4 As Boolean, bLvl5 As Boolean
Dim stNumList1(1 To 5) As String
Dim bLvlChm As Boolean, bLvlJpn As Boolean, bLvlEls As Boolean 'Add by Morgan 2009/9/7 未分案管制人by各組
Dim bLvlMot As Boolean   '2011/11/30 add by sonia 機械組未分案管制人
Dim bolShowAll As Boolean 'Added by Lydia 2021/08/23
Dim bLvO1 As Boolean 'Added by Lydia 2022/12/20 國外部期限通知未分案管制人
Dim colPA150 As Integer, colCaseNo As Integer 'Added by Lydia 2024/02/29
Dim m_stST16 As String 'Added by Lydia 2024/03/06 工程師組別


Private Sub cmdHelp_Click()
   frm060204_1.Show vbModal
End Sub

Private Sub cmdHide_Click()
   SetRst2Grid
   SetColor cmdHide.Tag
   'Add By Sindy 2023/8/31
   textSys.Tag = "": LblCnt.Visible = False '還原預設值
   Call TextSys_LostFocus
   '2023/8/31 END
End Sub

Private Sub SetRst2Grid()
   grdDataList.FixedCols = 0
   Set grdDataList.Recordset = m_adoRst
   grdDataList.FixedCols = 3
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData(Optional bolRefresh As Boolean = False)
   Dim i As Integer, StrTag As String, lngColor As Long, ii As Integer
   Dim StrToMail(1 To 6) As String 'Added by Lydia 2017/02/13
   Dim strMailCont As String 'Added by Lydia 2021/11/22 email預設內文
   
On Error GoTo ErrorHandler
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
         bolRefresh = False 'Add by Morgan 2008/9/22
         grdDataList.col = 0
         grdDataList.Text = ""
         grdDataList.CellBackColor = grdDataList.BackColor
         grdDataList.col = 3
         lngColor = grdDataList.CellBackColor
         For ii = 1 To 2
            grdDataList.col = ii
            grdDataList.CellBackColor = lngColor
         Next
   
         'Add by Amy 20130703
         'Modify By Sindy 2021/4/22 改判斷新案(命名追蹤)
         'If grdDataList.TextMatrix(i, 16) = "H" Then Exit For
         If grdDataList.TextMatrix(i, 16) = "新案" Then Exit For
         
         Dim Str01 As String
         
         StrTag = grdDataList.TextMatrix(i, 11)
         If Left(Right(StrTag, 7), 1) = "-" Then
            StrTag = StrTag & "-0-00"
         ElseIf Left(Right(StrTag, 2), 1) = "-" Then
            StrTag = StrTag & "-00"
         End If
         
         If Left(StrTag, 1) < "A" Or Left(StrTag, 1) > "Z" Then
            StrTag = Right(StrTag, Len(StrTag) - 1)
         End If
         Str01 = SystemNumber(StrTag, 1)
         If fnSaveParentForm(Me) = False Then
            Exit For
         End If
         
         'Added by Lydia 2024/01/12 增加對執行”基本資料”和”案件進度查詢”的限閱案件控制
         If cmdState = 0 Or cmdState = 1 Or cmdState = 2 Or cmdState = 3 Then
            'Modified by Lydia 2024/04/10 +pType = "1"
            If PUB_ChkCufaByCaseNo(strUserNum, Me.Name, Replace(StrTag, "-", ""), "1") = False Then
               Exit For
            End If
         End If
         'end 2024/01/12
         
         Me.Show
         Select Case cmdState
            'Add by Morgan 2008/9/22
            Case 0 '工作進度資料維護
               '達承辦或達核稿才可輸
               If grdDataList.TextMatrix(i, 16) = "B" Or grdDataList.TextMatrix(i, 16) = "C" Then
                  With frm090901
                     .txtCaseNo(1) = SystemNumber(StrTag, 1)
                     .txtCaseNo(2) = SystemNumber(StrTag, 2)
                     .txtCaseNo(3) = SystemNumber(StrTag, 3)
                     .txtCaseNo(4) = SystemNumber(StrTag, 4)
                     .Tag = grdDataList.TextMatrix(i, 28)
                     .CallFormName = Me.Name
                  End With
                  If frm090901.SetGrid(False) = False Then
                     Unload frm090901
                  End If
               Else
                  MsgBox "事件為達承辦或達核稿才可執行本功能！"
               End If
               
            'Added by Lydia 2017/02/13
            Case 1 '發E-Mail
                Me.Enabled = False
                Me.Hide
                '本所案號
                StrToMail(1) = StrTag
                '案件名稱
                StrToMail(2) = grdDataList.TextMatrix(i, 14)
                '收文日
                StrToMail(3) = ""
                '案件性質名稱
                StrToMail(4) = grdDataList.TextMatrix(i, 12)
                '法限
                StrToMail(5) = grdDataList.TextMatrix(i, 2)
                '所限
                StrToMail(6) = grdDataList.TextMatrix(i, 1)
                '內文
                'Modified by Lydia 2021/11/22 strExc(1) => strMailCont
                strMailCont = "           本所案號：" + StrToMail(1) + vbCrLf + vbCrLf + _
                            "           案件名稱：" + StrToMail(2) + vbCrLf + vbCrLf + _
                            "           案件性質：" + StrToMail(4) + vbCrLf + vbCrLf + _
                            "           本所期限：" + StrToMail(6) + "           法定期限：" + StrToMail(5) + vbCrLf + vbCrLf
                StrTag = ""
                '(NA16管制人16,R06承辦人17,R07智權人員18)
                '管制人
                StrTag = StrTag & grdDataList.TextMatrix(i, 18) & "-"
                '智權人員
                StrTag = StrTag & grdDataList.TextMatrix(i, 20) & "-"
                '承辦人
                'Added by Lydia 2017/02/24 新案翻譯發email的承辦人要去抓"核稿人"(若核稿人為所內工程師F外翻編號，請轉為FCP所內編號)，若核稿人為空白，則發e-mail對象的承辦人為空白。
                'Modified by Lydia 2021/04/29 若為新案翻譯【達核稿】時則承辦人要去抓"核稿人"(若核稿人為所內工程師F外翻編號，請轉為FCP所內編號)，若核稿人為空白，則發e-mail對象的承辦人為空白；其餘狀況皆預設為收文之承辦人。
                'If "" & grdDataList.TextMatrix(i, 25) = "201" Then
                'Modified by Lydia 2021/05/18 修改新案翻譯發email的承辦人，以是否輸入完稿日為準；
'                If "" & grdDataList.TextMatrix(i, 25) = "201" And "" & grdDataList.TextMatrix(i, 15) = "C" Then
'                   strExc(2) = PUB_GetEP04id("" & grdDataList.TextMatrix(i, 27), True)
'                   StrTag = StrTag & strExc(2) & "-"
'                Else
'                   StrTag = StrTag & grdDataList.TextMatrix(i, 18) & "-"
'                End If
'                'end 2017/02/24
'
'                'Added by Lydia 2021/04/29 新案翻譯【未交稿】改變預設內文
'                If "" & grdDataList.TextMatrix(i, 25) = "201" And "" & grdDataList.TextMatrix(i, 15) = "G" Then
'                     frm100106_4.strTypeMemo = "本案交稿期限為" & IIf("" & grdDataList.TextMatrix(i, 3) <> "", "" & grdDataList.TextMatrix(i, 3), "") & "，請速完成交稿，以利後續作業，謝謝。"
'                End If
'                'end 2021/04/29
                 'Memo by Lydia 2021/05/18 修改新案翻譯發email的承辦人，以是否輸入完稿日為準；
                        '1.新案翻譯未輸入完稿日
                        '1.1 預設收件人為承辦人：承辦人為所內員工(含下班翻譯)
                        '1.2 若未指定承辦人、國外翻譯社或是所外譯者時，email收件者改Sharon(翻譯分案人員)；
                        '2.新案翻譯已輸入完稿日
                        '2.1 預設收件人為核稿人；
                        '2.2 若未指定核稿人，email收件者改為各組工程師主管。
                strExc(10) = ""
                If "" & grdDataList.TextMatrix(i, 26) = "201" And "" & grdDataList.TextMatrix(i, 28) <> "" Then
                    strExc(1) = PUB_GetEP04id("" & grdDataList.TextMatrix(i, 28), True, strExc(2), strExc(3))
                    If strExc(3) = "" Then  '新案翻譯未輸入完稿日
                        If strExc(2) = "" Or Left(strExc(2), 1) = "F" Then
                            StrTag = StrTag & Pub_GetSpecMan("外專對外翻聯絡人員") & "-"
                        Else
                            StrTag = StrTag & strExc(2) & "-"
                        End If
                        strExc(10) = "本案交稿期限為" & IIf("" & grdDataList.TextMatrix(i, 4) <> "", "" & grdDataList.TextMatrix(i, 4), "") & "，請速完成交稿，以利後續作業，謝謝。"
                    Else  '新案翻譯已輸入完稿日
                        If strExc(1) <> "" Then
                            StrTag = StrTag & strExc(1) & "-"
                        Else
                            Call ChgCaseNo(Replace(StrToMail(1), "-", ""), strExc)
                            strExc(0) = "select oman from setspecman where ocode=(" & _
                                              "select decode(Pa150,'1','T','2','R','3','S','4','T1','外專對外翻聯絡人員') Ocode From Patent Where pa01='" & strExc(1) & "' And pa02='" & strExc(2) & "' And pa03='" & strExc(3) & "' And pa04='" & strExc(4) & "' ) "
                            intI = 1
                            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                            If intI = 1 Then
                                StrTag = StrTag & RsTemp.Fields("oman") & "-"
                            Else
                                StrTag = StrTag & "-"
                            End If
                        End If
                    End If
                'Added by Lydia 2021/05/20 非新案翻譯
                Else
                    StrTag = StrTag & grdDataList.TextMatrix(i, 19) & "-"
                'end 2021/05/20
                End If
                frm100106_4.strTypeMemo = strExc(10)
                'end 2021/05/18
                                
                Call frm100106_4.SetParent(Me, StrToMail(5)) 'Added by Lydia 2020/03/11 傳入前一畫面和法定期限
                'Modified by Lydia 2021/11/22  strExc(1) => strMailCont
                frm100106_4.txt1(1) = strMailCont
                Screen.MousePointer = vbHourglass
                frm100106_4.Show
                '狀態+表單名稱
                frm100106_4.strFRname = grdDataList.TextMatrix(i, 17) + "-" & Me.Name
                frm100106_4.strCaseNo = StrToMail(1) 'Added by Lydia 2020/05/18 傳入本所案號
                frm100106_4.Tag = StrTag
                frm100106_4.StrMenu
                Screen.MousePointer = vbDefault
                Me.Enabled = True
                Exit Sub
            'end 2017/02/13
            
            Case 2 '案件基本資料
               Select Case Pub_RplStr(Str01)
                  Case "CFP", "FCP", "P"   '專利
                     frm100101_3.Show
                     frm100101_3.Tag = StrTag
                     frm100101_3.StrMenu
                     
                  Case "FG"
                     frm100101_B.Show
                     frm100101_B.Tag = StrTag
                     frm100101_B.StrMenu
               End Select
               
            Case 3 '案件進度
               frm100101_2.Show
               frm100101_2.Tag = StrTag
               frm100101_2.StrMenu
         End Select
         Exit For
      End If
   Next i
   'Add by Morgan 2008/9/22
   If bolRefresh = True Then
      cmdQuery_Click 0
   End If
   
ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox "(" & Err.Number & ")" & Err.Description
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
End Sub

Public Sub cmdQuery_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   If pub_bolInformCheck = False Then
      If MsgBox("是否確定要查詢？", vbYesNo + vbDefaultButton2) = vbNo Then
         GoTo SubOut
      End If
   End If
   
   Me.Enabled = False
   doQuery Index
   'Add By Sindy 2023/8/31
   textSys.Text = "": textSys.Tag = "": LblCnt.Visible = False '還原預設值
   Call TextSys_LostFocus
   '2023/8/31 END
   Me.Enabled = True

SubOut:
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   bolBarShow = Forms(0).StatusBar1.Visible
   Forms(0).StatusBar1.Visible = True
End Sub

Private Sub Form_Deactivate()
   Forms(0).StatusBar1.Visible = bolBarShow
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Function GetNumList(p_UID As String, Optional iLevel As Integer) As String
   Dim stRtn As String, stSQL As String
   
   stSQL = "select ''''||st01||'''' from staff"
   Select Case iLevel
      Case 2
         stSQL = stSQL & " where ST52='" & p_UID & "'"
      Case 3
         stSQL = stSQL & " where ST53='" & p_UID & "'"
      Case 4
         stSQL = stSQL & " where ST54='" & p_UID & "'"
      Case 5
         stSQL = stSQL & " where ST55='" & p_UID & "'"
      Case Else
         stSQL = stSQL & " where ST52='" & p_UID & "'"
   End Select
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      stRtn = RsTemp.GetString(adClipString, , , ",")
      stRtn = Left(stRtn, Len(stRtn) - 1)
   End If
   GetNumList = stRtn
End Function

'語法內有用組合欄位為條件以控制使用特定index(避掉某些不適當的)
'Modify By Sindy 2023/10/27 EP33要回歸用在英文核完日,改抓EP39.核稿完成日
Private Sub doQuery(idx As Integer)
   Dim stVTB As String, stDate1 As String, stDate2 As String, stDate3 As String, stDate4 As String, stDate7 As String
   Dim stSQL As String, stCon As String, stConCP14 As String, stConEP04 As String, stConNA16 As String
   Dim stConNA51 As String, stConNP10 As String
   Dim stConCP06 As String, stConCP48 As String, stConEP08 As String
   Dim stConCP142 As String 'Add By Sindy 2021/4/20
   Dim stNumList As String, stDept As String
   Dim ii As Integer, jj As Integer, stIdList
   Dim stUserID As String
   Dim stCP01 As String
   Dim stConPAGrp As String, stConSPGrp As String 'Add by Morgan 2009/9/7
   Dim stConNA51P As String, stConSP26 As String 'Add by Lydia 2014/11/14 FCP承辦區域特殊狀況之智權人員劃分方式
   Dim stConNA16L As String, stConNA51L 'Add By Sindy 2015/10/22
   Dim rsA As ADODB.Recordset 'Add By Sindy 2017/1/16
   Dim strCP09 As String, strCP60 As String, strA1K01 As String, strA1K19 As String, strA1K20 As String
   Dim strDST01 As String 'Add By Sindy 2017/1/18
   Dim stOrdCon As String 'Added by Lydia 2018/02/08
   Dim strCPM1933_Col As String, strCPM1933_Where As String  'Add By Sindy 2020/8/7
   Dim strColR15_M As String, strColR15_N As String, strColR11_M As String, strColR11_N As String 'Add By Sindy 2024/12/31
   
   'Add by Morgan 2009/3/26
   stCP01 = " and cp01 in ('P','PS','FCP','FG','CFP','CPS')"
   
   If lblUserName = "" Then
      MsgBox "員工編號錯誤！", vbExclamation
      If txtUsernum.Enabled = True Then
         txtUsernum_GotFocus
         txtUsernum.SetFocus
      End If
      Exit Sub
   'Added by Morgan 2015/10/5
   ElseIf Pub_StrUserSt03 = "F22" And txtUsernum <> strUserNum Then
      If PUB_GetST03(txtUsernum) <> Pub_StrUserSt03 Then
         MsgBox "員工編號錯誤！", vbExclamation, "權限不足"
         Exit Sub
      End If
   'end 2015/10/5
   End If
   
   stUserID = txtUsernum
   If stUserID = strUserNum Then
      stDept = Pub_StrUserSt15
   Else
      stDept = GetST15(stUserID)
   End If
   
   'Added by Lydia 2024/03/06 非機械設計組人員才顯示符號
   m_stST16 = PUB_GetStaffST16(stUserID)
   If m_stST16 <> "4" Then
      LblNote.Visible = True
   Else
      LblNote.Visible = False
   End If
   'end 2024/03/06
   
   stNumList = PUB_GetMapID(stUserID, 0)
   If stNumList <> "" Then
      stNumList = "'" & stNumList & "','" & stUserID & "'"
   Else
      stNumList = "'" & stUserID & "'"
   End If
   stNumList1(1) = stNumList
   
   For ii = 2 To 5
      stNumList1(ii) = GetNumList(stUserID, ii)
      If stNumList1(ii) <> "" Then
         stNumList = stNumList & "," & stNumList1(ii)
      End If
   Next
   
   stIdList = Split(stNumList, ",")
   'Add by Morgan 2009/12/18 去除重複編號
   stNumList = stIdList(LBound(stIdList))
   For ii = LBound(stIdList) + 1 To UBound(stIdList)
      For jj = LBound(stIdList) To ii - 1
         If stIdList(jj) = stIdList(ii) Then
            Exit For
         End If
      Next
      If ii = jj Then
         stNumList = stNumList & "," & stIdList(ii)
      End If
   Next
   stIdList = Split(stNumList, ",")
   'end 2009/12/18
   
   stDate1 = strSrvDate(1) - 10000
   stDate2 = CompWorkDay(4, strSrvDate(1))
   stDate3 = CompWorkDay(7, strSrvDate(1))
   'stDate3指本所期限前五個工作天的條件,但因當天不算所以為系統日+6天
   '又因6天後若為星期五,則星期六日的期限也要在當天出現,所以為系統日+7天且判斷條件為CP06<
   
   stDate7 = CompWorkDay(8, strSrvDate(1), 1) 'Add By Sindy 2022/3/15 系統日前7個工作天(不含當日)
   
   stConCP06 = " AND CP06>=" & stDate1 & " AND CP06< " & stDate2
   stConCP48 = " AND CP48>=" & stDate1 & " AND CP48< " & stDate2
   stConEP08 = " AND EP08>=" & stDate1 & " AND EP08< " & stDate2
   stConCP142 = " AND CP142>=" & stDate1 & " AND CP142< " & stDate2 'Add By Sindy 2021/4/20
   
   bLvlX = CheckLevel(stUserID, "M") '未交稿,已完稿無核稿管制人
   bLvl4 = CheckLevel(stUserID, "N") '第四級管制人(+FCP,FG未分案將到期) :[有關期限] 國外部專利處非外專承辦或未分案將到期管制人(含逾期)
   bLvl5 = CheckLevel(stUserID, "O") '第五級管制人(+FCP,FG未分案已逾期) :國外部專利處非外專承辦或未分案已逾期管制人
   bLvO1 = CheckLevel(stUserID, "O1") 'Added by Lydia 2022/12/20 國外部期限通知未分案管制人
   
   'Add by Morgan 2009/9/7 未分案改分組管制
   bLvlChm = CheckLevel(stUserID, "R") '未分案化學組管制人
   bLvlJpn = CheckLevel(stUserID, "S") '未分案日文組管制人
   '2011/11/30 modify by sonia 取消德文且機電拆電子電機及機械設計
   'bLvlEls = CheckLevel(stUserID, "T") '未分案機電德文其他組管制人
   bLvlMot = CheckLevel(stUserID, "T1") '未分案機械設計組管制人
   bLvlEls = CheckLevel(stUserID, "T") '未分案電子電機其他組管制人
   '2011/11/30 end
   
   'Add by Lydia 2014/11/14 FCP承辦區域特殊狀況之智權人員劃分方式
   '代理人Y51333010=Pub_GetSpecMan("北京銀龍FCP案承辦業務") ,NA51 = decode(pa75,'Y51333010','" & midStr & "',na51)
   Dim midStr As String
   'Modified by Lydia 2016/02/03改成回傳case句
   'midStr = Pub_GetSpecMan("北京銀龍FCP案承辦業務")
   midStr = Pub_GetSpecFCP
   
   If InStr(stNumList, ",") > 0 Then
      stConCP14 = " AND CP14 in (" & stNumList & ") "
      stConEP04 = " AND EP04 in (" & stNumList & ") "
      'Modified by Lydia 2022/11/03 區分FMP案
      'stConNA16 = " AND NA16 in (" & stNumList & ") "
      stConNA16 = " AND ((CP01 in ('FCP','FG') and NA16 in (" & stNumList & ")) or (CP01 not in ('FCP','FG') and nvl(NA79,NA16) in (" & stNumList & "))) "
      stConNA16L = " AND (n1.NA16 in (" & stNumList & ") OR n2.NA16 in (" & stNumList & "))" 'Add By Sindy 2015/10/22
      stConNA51 = " AND NA51 in (" & stNumList & ") "
      stConNP10 = " AND NP10 in (" & stNumList & ") "
      'Add by Lydia 2014/11/14 FCP承辦區域特殊狀況之智權人員劃分方式
      'Modified by Lydia 2016/02/03  +Y51817040日文承辦
'      stConNA51P = " AND decode(pa75,'Y51333010','" & midStr & "',na51) in (" & stNumList & ") "
'      stConSP26 = " AND decode(sp26,'Y51333010','" & midStr & "',na51) in (" & stNumList & ") "
'      stConNA51L = " AND (decode(LC22,'Y51333010','" & midStr & "',n1.na51) in (" & stNumList & ")" & _
'                        " OR decode(LC22,'Y51333010','" & midStr & "',n2.na51) in (" & stNumList & "))" 'Add By Sindy 2015/10/22
      stConNA51P = " AND decode(pa75," & midStr & ",na51) in (" & stNumList & ") "
      stConSP26 = " AND decode(sp26," & midStr & ",na51) in (" & stNumList & ") "
      stConNA51L = " AND (decode(LC22," & midStr & ",n1.na51) in (" & stNumList & ")" & _
                        " OR decode(LC22," & midStr & ",n2.na51) in (" & stNumList & "))"
   Else
      stConCP14 = " AND CP14=" & stNumList
      stConEP04 = " AND EP04=" & stNumList
      'Modified by Lydia 2022/11/03 區分FMP案
      'stConNA16 = " AND NA16=" & stNumList
      stConNA16 = " AND ((CP01 in ('FCP','FG') and NA16 =" & stNumList & " ) or (CP01 not in ('FCP','FG') and nvl(NA79,NA16) =" & stNumList & " )) "
      stConNA16L = " AND (n1.NA16=" & stNumList & " OR n2.NA16=" & stNumList & ")" 'Add By Sindy 2015/10/22
      stConNA51 = " AND NA51=" & stNumList
      stConNP10 = " AND NP10=" & stNumList
      'Add by Lydia 2014/11/14 FCP承辦區域特殊狀況之智權人員劃分方式
      'Modified by Lydia 2016/02/03  +Y51817040日文承辦
'      stConNA51P = " AND decode(pa75,'Y51333010','" & midStr & "',na51) =" & stNumList
'      stConSP26 = " AND decode(sp26,'Y51333010','" & midStr & "',na51) =" & stNumList
'      stConNA51L = " AND (decode(LC22,'Y51333010','" & midStr & "',n1.na51) =" & stNumList & _
'                        " OR decode(LC22,'Y51333010','" & midStr & "',n2.na51) =" & stNumList & ")" 'Add By Sindy 2015/10/22
      stConNA51P = " AND decode(pa75," & midStr & ",na51) =" & stNumList
      stConSP26 = " AND decode(sp26," & midStr & ",na51) =" & stNumList
      stConNA51L = " AND (decode(LC22," & midStr & ",n1.na51) =" & stNumList & _
                        " OR decode(LC22," & midStr & ",n2.na51) =" & stNumList & ")"
   End If
   
   'Add by Morgan 2009/10/22
   '清除暫存檔
   stSQL = "delete R060204 where R01='" & strUserNum & "'"
   cnnConnection.Execute stSQL, intI

   'Modify By Sindy 2017/1/18 + 'I','准未請款','J','分割建議','K','通知告准'
   'Modified by Lydia 2017/11/29 + 'L','待命名期限'
   'Modified by Lydia 2018/02/08  + 'M','待處理'
   '代碼1(R02):A=達本所,B=達承辦,C=達核稿,D=未收文,E=未發文,F=未請款,G=未交稿,H=早收文(Added by Lydia 2015/09/09)
   '           I=准未請款,J=分割建議,K=通知告准,L=待命名期限,M=待處理,N=達指定,(英文的)O=未分案,P=達約定
   '代碼2(R03):(數字的)0=未分案,1=承辦人,2=管制人,3=智權人員,4=核稿人,8=無核稿人 9=未交稿
   
   'R01:UserID,R02:代碼1,R03:代碼2,R04:收文號,R05:國籍,R06:承辦,R07:業務,R08:所限,R09:法限
   ',R10:辦限,R11:核限,R12:NP22,R13:備註,R14:指定日,R15:約定期限
   
   'Modify by Morgan 2009/10/22 改寫暫存檔以便過濾達所限又達承辦或核稿期限的資料(只要顯示達所限)
   '已收文未發文,2個工作天後達本所期限者(不含當日) --承辦人-A1(本所期限,承辦人)
   'Modified by Lydia 2016/09/14 CP14||CP57||CP27 => CP158=0 AND CP159=0 AND CP14
   'Modified by Lydia 2016/09/22 +/*+ INDEX(CASEPROGRESS IDXCP15815914) */
   'Modified by Lydia 2018/02/08 客戶提供文件1920 =>'M','待處理'
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
      " SELECT '" & strUserNum & "',DECODE(CP10,'1920','M','A') EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
      " From CASEPROGRESS,PATENT,FAGENT" & _
      " WHERE CP158=0 AND CP159=0 AND CP14 IN(" & stNumList & ")" & stConCP06 & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
   cnnConnection.Execute stSQL, intI
   
   'Add By Sindy 2015/10/22 +法務:達本所
   'Modified by Lydia 2016/09/14 CP14||CP57||CP27 => CP158=0 AND CP159=0 AND CP14
   'Modified by Lydia 2016/09/22 +/*+ INDEX(CASEPROGRESS IDXCP15815914) */
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
      " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,nvl(FA10,cu10),CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
      " From CASEPROGRESS,LawCase,FAGENT,Customer" & _
      " WHERE CP158=0 AND CP159=0 AND CP14 IN(" & stNumList & ")" & stConCP06 & _
      " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC08||LC34 IS NULL AND LC01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(LC22,1,8) AND FA02(+)=SUBSTR(LC22,9)" & _
      " AND CU01(+)=SUBSTR(LC11,1,8) AND CU02(+)=SUBSTR(LC11,9)"
   cnnConnection.Execute stSQL, intI
   
   'Modified by Lydia 2016/09/14 CP14||CP57||CP27 => CP158=0 AND CP159=0 AND CP14
   'Modified by Lydia 2016/09/22 +/*+ INDEX(CASEPROGRESS IDXCP15815914) */
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
      " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
      " From CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
      " WHERE CP158=0 AND CP159=0 AND CP14 IN(" & stNumList & ")" & stConCP06 & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
   cnnConnection.Execute stSQL, intI
   
   'Add By Sindy 2021/4/21 達指定
   'Modify By Sindy 2024/1/30 + AND CP10<>'201':FCP-070662(201) 此SQL排除201下列會抓
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
      " SELECT '" & strUserNum & "','N' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
      " From CASEPROGRESS,PATENT,FAGENT" & _
      " WHERE CP158=0 AND CP159=0 AND CP10<>'201' AND CP14 IN(" & stNumList & ")" & stConCP142 & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
   cnnConnection.Execute stSQL, intI
   'Add By Sindy 2023/9/19 新案翻譯的達指定
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
      " SELECT '" & strUserNum & "','N' EV1,decode(EP04,null,'0','1') EV2,CP09,FA10,decode(EP04,null,decode(SIM01,null,CP14,SIM01),EP04),CP13,CP06,CP07,CP48,EP08,0,cp142" & _
      " FROM CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT,STAFF_IDMAP,STAFF S1,STAFF S2" & _
      " WHERE CP10='201' AND CP158=0 AND CP159=0 and substr(CP12,1,1)='F'" & stConCP142 & stCP01 & _
      " AND EP02(+)=CP09" & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
      " AND S1.ST01(+)=CP14 AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01" & _
      " AND decode(EP04,null,decode(SIM01,null,CP14,SIM01),EP04) in (" & stNumList & ")"
   cnnConnection.Execute stSQL, intI
   '2023/9/19 END
      
   'Add By Sindy 2021/4/21 達指定
   'Modify By Sindy 2024/1/30 + AND CP10<>'201':FCP-070662(201) 此SQL排除201下列會抓
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
      " SELECT '" & strUserNum & "','N' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
      " From CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
      " WHERE CP158=0 AND CP159=0 AND CP10<>'201' AND CP14 IN(" & stNumList & ")" & stConCP142 & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
   cnnConnection.Execute stSQL, intI
   'Add By Sindy 2023/9/19 新案翻譯的達指定
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
      " SELECT '" & strUserNum & "','N' EV1,decode(EP04,null,'0','1') EV2,CP09,FA10,decode(EP04,null,decode(SIM01,null,CP14,SIM01),EP04),CP13,CP06,CP07,CP48,EP08,0,cp142" & _
      " FROM CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT,STAFF_IDMAP,STAFF S1,STAFF S2" & _
      " WHERE CP10='201' AND CP158=0 AND CP159=0 and substr(CP12,1,1)='F'" & stConCP142 & stCP01 & _
      " AND EP02(+)=CP09" & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)" & _
      " AND S1.ST01(+)=CP14 AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01" & _
      " AND decode(EP04,null,decode(SIM01,null,CP14,SIM01),EP04) in (" & stNumList & ")"
   cnnConnection.Execute stSQL, intI
   '2023/9/19 END
      
   '已收文未發文,2個工作天後達承辦期限者(不含當日) --承辦人-B1(承辦期限,承辦人)
   'Modified by Lydia 2016/09/14 CP14||CP57||CP27 => CP158=0 AND CP159=0 AND CP14
   'Modified by Lydia 2016/09/22 +/*+ INDEX(CASEPROGRESS IDXCP15815914) */
   'Modify By Sindy 2024/1/3 取消 AND EP09 IS NULL
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
      " SELECT '" & strUserNum & "','B' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
      " From CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT" & _
      " WHERE CP158=0 AND CP159=0 AND CP14 IN(" & stNumList & ")" & stConCP48 & _
      " AND EP02(+)=CP09" & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
   cnnConnection.Execute stSQL, intI

   'Modified by Lydia 2016/09/14 CP14||CP57||CP27 => CP158=0 AND CP159=0 AND CP14
   'Modified by Lydia 2016/09/22 +/*+ INDEX(CASEPROGRESS IDXCP15815914) */
   'Modify By Sindy 2024/1/3 取消 AND EP09 IS NULL
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
      " SELECT '" & strUserNum & "','B' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
      " From CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT" & _
      " WHERE CP158=0 AND CP159=0 AND CP14 IN(" & stNumList & ")" & stConCP48 & _
      " AND EP02(+)=CP09" & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
   cnnConnection.Execute stSQL, intI

   '所有未發文--承辦人-E(未發文)
   If idx = 1 Then
      For ii = LBound(stIdList) To UBound(stIdList)
         'Modified by Lydia 2016/09/14 CP57||CP27 is null => CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS,PATENT,FAGENT" & _
            " WHERE CP05>20030000 AND CP14=" & stIdList(ii) & " and CP158=0 AND CP159=0" & _
            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
         cnnConnection.Execute stSQL, intI
         
         'Add By Sindy 2015/10/22 +法務:未發文
         'Modified by Lydia 2016/09/14 CP57||CP27 is null => CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,nvl(FA10,cu10),CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS,Lawcase,FAGENT,Customer" & _
            " WHERE CP05>20030000 AND CP14=" & stIdList(ii) & " and CP158=0 AND CP159=0" & _
            " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC08||LC34 IS NULL AND LC01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(LC22,1,8) AND FA02(+)=SUBSTR(LC22,9)" & _
            " AND CU01(+)=SUBSTR(LC11,1,8) AND CU02(+)=SUBSTR(LC11,9)"
         cnnConnection.Execute stSQL, intI
         '2015/10/22 END
         
         'Modified by Lydia 2016/09/14 CP57||CP27 is null => CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
            " WHERE CP05>20030000 AND CP14=" & stIdList(ii) & " AND CP158=0 AND CP159=0" & _
            " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
         cnnConnection.Execute stSQL, intI
      Next ii
   End If
   
   '已發文未請款--承辦人
   For ii = LBound(stIdList) To UBound(stIdList)
      'Modify By Sindy 2020/8/7 為配合日本部及專利國外部，於主管會議中經各組主管同意，
      '工程師實施無卷請款，為控管避免請款延遲太久，系統未請款彈跳期限通知，
      '工程師請款部分（參附件標黃色）請改為5個工作天
''      strCPM1933_Col = "TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD')"
''      strCPM1933_Where = "TO_CHAR(ADD_MONTHS(SYSDATE,-1*DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD')"
'      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD'),WORKDAYADD(cpm33+1,CP27))"
'      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
'      'Modify by Morgan 2010/4/12 排除工程師提申940
'      'Modified by Lydia 2016/09/14 CP20||CP57||CP60 IS NULL AND CP14 => CP27>20150914 AND CP14='...' AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'
'      'Modified by Morgan 2025/8/26 +FCP、FMP（包含寰華案）之B類（204）補正,不判斷收文金額(B類會沒有) --敏莉
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,CP09,FA10,CP14,CP13" & _
'         ",NULL,NULL," & strCPM1933_Col & " C06,NULL,0" & _
'         " FROM CASEPROGRESS,PATENT,FAGENT,CASEPROPERTYMAP" & _
'         " WHERE CP27>" & stDate1 & " AND CP14=" & stIdList(ii) & " AND CP159=0 AND NVL(CP20||CP60,'0')='0'" & _
'         " And ((CP09<'B' AND CP16>0) or cp10='204') AND CP10 NOT IN ('403','404','411','418','419','901','902','907','908','913','940')" & _
'         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
'         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP27<=" & strCPM1933_Where
'      cnnConnection.Execute stSQL, intI
      'Modify By Sindy 2025/10/9 修改未請款程式
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(decode(CP01,'FCP',CP27,CP47),'YYYYMMDD'),DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD'),WORKDAYADD(cpm33+1,decode(CP01,'FCP',CP27,CP47)))"
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,CP09,FA10,CP14,CP13" & _
         ",NULL,NULL," & strCPM1933_Col & " C06,NULL,0" & _
         " FROM CASEPROGRESS,PATENT,FAGENT,CASEPROPERTYMAP" & _
         " WHERE CP27>" & stDate1 & " AND CP14=" & stIdList(ii) & " AND CP159=0 AND NVL(CP20||CP60,'0')='0'" & _
         " And ((CP09<'B' AND CP16>0) or cp10='204') AND CP10 NOT IN ('403','404','411','418','419','901','902','907','908','913','940')" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)='FCP' AND CPM02(+)=CP10" & _
         " AND (CP01='FCP' or (CP01<>'FCP' AND CP47 is not null))"
      cnnConnection.Execute stSQL, intI
      '2025/10/9 END
      
      'Modify By Sindy 2020/8/7 為配合日本部及專利國外部，於主管會議中經各組主管同意，
      '工程師實施無卷請款，為控管避免請款延遲太久，系統未請款彈跳期限通知，
      '工程師請款部分（參附件標黃色）請改為5個工作天
''      strCPM1933_Col = "TO_CHAR(ADD_MONTHS(TO_DATE(C2.CP27,'YYYYMMDD'),CPM19),'YYYYMMDD')"
''      strCPM1933_Where = "TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD')"
'      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(C2.CP27,'YYYYMMDD'),CPM19),'YYYYMMDD'),WORKDAYADD(cpm33+1,C2.CP27))"
'      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
'      'Add by Morgan 2010/4/12 工程師提申=940 要抓新申請案的設定
'      'Modified by Lydia 2016/09/14 CP20||CP57||CP60 IS NULL AND CP14 => CP27>20150914 AND CP14='...' AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,C2.CP09,FA10,C2.CP14,C2.CP13" & _
'         ",NULL,NULL," & strCPM1933_Col & " C06,NULL,0" & _
'         " FROM CASEPROGRESS C1,CASEPROGRESS C2,PATENT,FAGENT,CASEPROPERTYMAP" & _
'         " WHERE C1.CP27>" & stDate1 & " AND C1.CP14=" & stIdList(ii) & " AND C1.CP159=0 AND C1.CP16>0 AND C1.CP20||C1.CP60 IS NULL" & _
'         " And C1.CP09<'B' AND C1.CP10='940'" & _
'         " AND PA01(+)=C1.CP01 AND PA02(+)=C1.CP02 AND PA03(+)=C1.CP03 AND PA04(+)=C1.CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
'         " AND C2.CP01(+)=C1.CP01 AND C2.CP02(+)=C1.CP02 AND C2.CP03(+)=C1.CP03 AND C2.CP04(+)=C1.CP04 AND C2.CP10 IN ('101','102','103','105','125') AND C2.CP159=0" & _
'         " AND CPM01(+)=C2.CP01 AND CPM02(+)=C2.CP10 AND C1.CP27<=" & strCPM1933_Where
'      'end 2010/10/14
'      cnnConnection.Execute stSQL, intI
      'Modify By Sindy 2025/10/9 修改未請款程式
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(decode(C2.CP01,'FCP',C2.CP27,C2.CP47),'YYYYMMDD'),CPM19),'YYYYMMDD'),WORKDAYADD(cpm33+1,decode(C2.CP01,'FCP',C2.CP27,C2.CP47)))"
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,C2.CP09,FA10,C2.CP14,C2.CP13" & _
         ",NULL,NULL," & strCPM1933_Col & " C06,NULL,0" & _
         " FROM CASEPROGRESS C1,CASEPROGRESS C2,PATENT,FAGENT,CASEPROPERTYMAP" & _
         " WHERE C1.CP27>" & stDate1 & " AND C1.CP14=" & stIdList(ii) & " AND C1.CP159=0 AND C1.CP16>0 AND C1.CP20||C1.CP60 IS NULL" & _
         " And C1.CP09<'B' AND C1.CP10='940'" & _
         " AND PA01(+)=C1.CP01 AND PA02(+)=C1.CP02 AND PA03(+)=C1.CP03 AND PA04(+)=C1.CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND C2.CP01(+)=C1.CP01 AND C2.CP02(+)=C1.CP02 AND C2.CP03(+)=C1.CP03 AND C2.CP04(+)=C1.CP04 AND C2.CP10 IN ('101','102','103','105','125') AND C2.CP159=0" & _
         " AND CPM01(+)='FCP' AND CPM02(+)=C2.CP10" & _
         " AND (C2.CP01='FCP' or (C2.CP01<>'FCP' AND C2.CP47 is not null))"
      cnnConnection.Execute stSQL, intI
      '2025/10/9 END
      
      'Modify By Sindy 2020/8/7 為配合日本部及專利國外部，於主管會議中經各組主管同意，
      '工程師實施無卷請款，為控管避免請款延遲太久，系統未請款彈跳期限通知，
      '工程師請款部分（參附件標黃色）請改為5個工作天
'      strCPM1933_Col = "TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),nvl(CPM19,0)),'YYYYMMDD')"
'      strCPM1933_Where = "TO_CHAR(ADD_MONTHS(SYSDATE,-1*nvl(CPM19,0)),'YYYYMMDD')"
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),nvl(CPM19,0)),'YYYYMMDD'),WORKDAYADD(cpm33+1,CP27))"
      'Modify By Sindy 2025/10/9 取消 strCPM1933_Where
'      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*nvl(CPM19,0)),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
      'Add By Sindy 2015/10/22 +法務:未請款
      'Modified by Lydia 2016/09/14 CP20||CP57||CP60 IS NULL AND CP14 => CP27>20150914 AND CP14='...' AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,CP09,nvl(FA10,cu10),CP14,CP13" & _
         ",NULL,NULL," & strCPM1933_Col & " C06,NULL,0" & _
         " FROM CASEPROGRESS,Lawcase,FAGENT,CASEPROPERTYMAP,Customer" & _
         " WHERE CP27>" & stDate1 & " AND CP14=" & stIdList(ii) & " AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'" & _
         " And CP09<'B'" & _
         " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC08||LC34 IS NULL AND LC01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(LC22,1,8) AND FA02(+)=SUBSTR(LC22,9)" & _
         " AND CU01(+)=SUBSTR(LC11,1,8) AND CU02(+)=SUBSTR(LC11,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10" 'AND CP27<=" & strCPM1933_Where
      cnnConnection.Execute stSQL, intI
      '2015/10/22 END
      
      'Modify By Sindy 2020/8/7 為配合日本部及專利國外部，於主管會議中經各組主管同意，
      '工程師實施無卷請款，為控管避免請款延遲太久，系統未請款彈跳期限通知，
      '工程師請款部分（參附件標黃色）請改為5個工作天
''      strCPM1933_Col = "TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),CPM19),'YYYYMMDD')"
''      strCPM1933_Where = "TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD')"
'      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),CPM19),'YYYYMMDD'),WORKDAYADD(cpm33+1,CP27))"
'      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
'      'Modified by Lydia 2016/09/14 CP20||CP57||CP60 IS NULL AND CP14 => CP27>20150914 AND CP14='...' AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,CP09,FA10,CP14,CP13" & _
'         ",NULL,NULL," & strCPM1933_Col & " CP06,NULL,0" & _
'         " FROM CASEPROGRESS,SERVICEPRACTICE,FAGENT,CASEPROPERTYMAP" & _
'         " WHERE CP27>" & stDate1 & " AND CP14=" & stIdList(ii) & " AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'" & _
'         " And CP09<'B' AND CP10 NOT IN ('403','404','411','418','419','901','902','907','908','913')" & _
'         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)" & _
'         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP27<=" & strCPM1933_Where
'      cnnConnection.Execute stSQL, intI
      'Modify By Sindy 2025/10/9 修改未請款程式
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(decode(CP01,'FG',CP27,CP47),'YYYYMMDD'),CPM19),'YYYYMMDD'),WORKDAYADD(cpm33+1,decode(CP01,'FG',CP27,CP47)))"
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,CP09,FA10,CP14,CP13" & _
         ",NULL,NULL," & strCPM1933_Col & " CP06,NULL,0" & _
         " FROM CASEPROGRESS,SERVICEPRACTICE,FAGENT,CASEPROPERTYMAP" & _
         " WHERE CP27>" & stDate1 & " AND CP14=" & stIdList(ii) & " AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'" & _
         " And CP09<'B' AND CP10 NOT IN ('403','404','411','418','419','901','902','907','908','913')" & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)" & _
         " AND CPM01(+)='FG' AND CPM02(+)=CP10" & _
         " AND (CP01='FG' or (CP01<>'FG' AND CP47 is not null))"
      cnnConnection.Execute stSQL, intI
      '2025/10/9 END
   Next
   
   '非外專工程師承辦案件(部門非 F2,F5,F8 字頭的) : (N:[有關期限] 國外部專利處非外專承辦或未分案將到期管制人(含逾期))
   If bLvl4 = True Then
      '已收文未發文,2個工作天後達本所期限者(不含當日) --承辦人-A1(本所期限,承辦人)
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' AND CP05+0>20030000 => CP01='FCP' AND CP05+0>20030000 AND CP158=0 AND CP159=0
      'Modified by Lydia 2018/02/12  客戶提供文件1920=>'M','待處理'
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "',DECODE(CP10,'1920','M','A') EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,PATENT,FAGENT" & _
         " WHERE CP01='FCP' AND CP05+0>20030000 AND CP158=0 AND CP159=0" & stConCP06 & _
         " AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
      cnnConnection.Execute stSQL, intI
      
      '已收文未發文,2個工作天後達承辦期限者(不含當日) --承辦人-B1(承辦期限,承辦人)
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' AND CP05+0>20030000 => CP01='FCP' AND CP05+0>20030000 AND CP158=0 AND CP159=0
      'Modify By Sindy 2024/1/3 取消 AND EP09 IS NULL
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','B' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
         " From CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT" & _
         " WHERE CP01='FCP' AND CP05+0>20030000 AND CP158=0 AND CP159=0" & stConCP48 & _
         " AND EP02(+)=CP09" & _
         " AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
      cnnConnection.Execute stSQL, intI
      
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FG' AND CP05+0>20030000 => CP01='FG' AND CP05+0>20030000 AND CP158=0 AND CP159=0
      'Modify By Sindy 2024/1/3 取消 AND EP09 IS NULL
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','B' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
         " From CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT" & _
         " WHERE CP01='FG' AND CP05+0>20030000 AND CP158=0 AND CP159=0" & stConCP48 & _
         " AND EP02(+)=CP09" & _
         " AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
      cnnConnection.Execute stSQL, intI
      
      '所有未發文--承辦人-E(未發文)
      If idx = 1 Then
         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' AND CP05+0>20030000 => CP01='FCP' AND CP05+0>20030000 AND CP158=0 AND CP159=0
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS,PATENT,FAGENT" & _
            " WHERE CP01='FCP' AND CP05+0>20030000 AND CP158=0 AND CP159=0" & _
            " AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
         cnnConnection.Execute stSQL, intI
         
         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FG' AND CP05>20030000 AND CP14 IS NOT NULL => CP01='FG' AND CP05+0>20030000 AND CP14 IS NOT NULL AND CP158=0 AND CP159=0
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
            " WHERE CP01='FG' AND CP05+0>20030000 AND CP14 IS NOT NULL AND CP158=0 AND CP159=0" & _
            " AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
            " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
         cnnConnection.Execute stSQL, intI
      End If
      
      'Modify By Sindy 2020/8/7 為配合日本部及專利國外部，於主管會議中經各組主管同意，
      '工程師實施無卷請款，為控管避免請款延遲太久，系統未請款彈跳期限通知，
      '工程師請款部分（參附件標黃色）請改為5個工作天
'      strCPM1933_Col = "TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD')"
'      strCPM1933_Where = "TO_CHAR(ADD_MONTHS(SYSDATE,-1*DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD')"
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD'),WORKDAYADD(cpm33+1,CP27))"
      'Modify By Sindy 2025/10/9 取消 strCPM1933_Where
'      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
      '已發文未請款--承辦人
      'Modify by Morgan 2010/4/12 排除工程師提申940
      'Modified by Lydia 2016/09/14
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,CP09,FA10,CP14,CP13" & _
         ",NULL,NULL," & strCPM1933_Col & " C06,NULL,0" & _
         " FROM CASEPROGRESS,PATENT,FAGENT,CASEPROPERTYMAP" & _
         " WHERE CP01='FCP' AND CP27>" & stDate1 & " And CP09||''<'B' AND CP10 NOT IN ('403','404','411','418','419','901','902','907','908','913','940')" & _
         " AND NVL(CP20||CP60,'0')='0' AND CP16>0 AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10" '" AND CP27<=" & strCPM1933_Where
      cnnConnection.Execute stSQL, intI
      
      'Modify By Sindy 2020/8/7 為配合日本部及專利國外部，於主管會議中經各組主管同意，
      '工程師實施無卷請款，為控管避免請款延遲太久，系統未請款彈跳期限通知，
      '工程師請款部分（參附件標黃色）請改為5個工作天
'      strCPM1933_Col = "TO_CHAR(ADD_MONTHS(TO_DATE(C2.CP27,'YYYYMMDD'),CPM19),'YYYYMMDD')"
'      strCPM1933_Where = "TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD')"
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(C2.CP27,'YYYYMMDD'),CPM19),'YYYYMMDD'),WORKDAYADD(cpm33+1,C2.CP27))"
      'Modify By Sindy 2025/10/9 取消 strCPM1933_Where
'      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
      'Add by Morgan 2010/4/12 工程師提申940要抓新申請案的設定
      'Modify by Morgan 2010/10/14 工程師提申940由新申請案的承辦人管制且帶新申請案的資料
      'Modified by Lydia 2016/09/14
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,C2.CP09,FA10,C2.CP14,C2.CP13" & _
         ",NULL,NULL," & strCPM1933_Col & " C06,NULL,0" & _
         " FROM CASEPROGRESS C1,CASEPROGRESS C2,PATENT,FAGENT,CASEPROPERTYMAP" & _
         " WHERE C1.CP01='FCP' AND C1.CP27>" & stDate1 & " And C1.CP09||''<'B' AND C1.CP10='940'" & _
         " AND NVL(C1.CP20||C1.CP60,'0')='0' AND C1.CP16>0 AND EXISTS(SELECT * FROM STAFF WHERE ST01=C1.CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND PA01(+)=C1.CP01 AND PA02(+)=C1.CP02 AND PA03(+)=C1.CP03 AND PA04(+)=C1.CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND C2.CP01(+)=C1.CP01 AND C2.CP02(+)=C1.CP02 AND C2.CP03(+)=C1.CP03 AND C2.CP04(+)=C1.CP04 AND C2.CP10 IN ('101','102','103','105','125') AND C2.CP159=0" & _
         " AND CPM01(+)=C2.CP01 AND CPM02(+)=C2.CP10" '" AND C1.CP27<=" & strCPM1933_Where
      'end 2010/10/14
      cnnConnection.Execute stSQL, intI
      
      'Modify By Sindy 2020/8/7 為配合日本部及專利國外部，於主管會議中經各組主管同意，
      '工程師實施無卷請款，為控管避免請款延遲太久，系統未請款彈跳期限通知，
      '工程師請款部分（參附件標黃色）請改為5個工作天
'      strCPM1933_Col = "TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),CPM19),'YYYYMMDD')"
'      strCPM1933_Where = "TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD')"
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),CPM19),'YYYYMMDD'),WORKDAYADD(cpm33+1,CP27))"
      'Modify By Sindy 2025/10/9 取消 strCPM1933_Where
'      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
      'Modified by Lydia 2016/09/14
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,' ' EV2,CP09,FA10,CP14,CP13" & _
         "," & strCPM1933_Col & " CP06,NULL,NULL,NULL,0" & _
         " FROM CASEPROGRESS,SERVICEPRACTICE,FAGENT,CASEPROPERTYMAP" & _
         " WHERE CP01='FG' AND CP27>" & stDate1 & " AND CP16>0 AND CP14 IS NOT NULL And CP09||''<'B'" & _
         " AND CP10 NOT IN ('403','404','411','418','419','901','902','907','908','913')" & _
         " AND NVL(CP20||CP60,'0')='0' AND CP16>0 AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10" '" AND CP27<=" & strCPM1933_Where
      cnnConnection.Execute stSQL, intI
   End If
   'Add By Sindy 2025/10/13 F=未請款 R03:0=未分案,1=承辦人,2=管制人; 程序人員達R10:未請款承辦期限才顯示
   stSQL = "DELETE R060204 WHERE R01='" & strUserNum & "' AND R02='F' AND R03='1' AND R10>" & strSrvDate(1) & _
           " AND R04 in(select cp09 from caseprogress,staff WHERE CP09=R04 and CP14=ST01(+) and ST03='F22')"
   cnnConnection.Execute stSQL, intI
   
   '未分案-0
   'Modify by Morgan 2009/9/7 未分案改分組管制
   'If bLvl4 = True Or bLvl5 = True Then
   'Modified by Lydia 2022/12/20 +bLvO1
   '工程師
   If bLvlChm Or bLvlJpn Or bLvlMot Or bLvlEls Or bLvl5 Or bLvO1 Then
      stConPAGrp = "": stConSPGrp = ""
      'Modify By Sindy 2021/4/22 排除案件性質是412延緩公告417提早公開
      'Modified by Lydia 2022/12/20 國外部期限通知未分案管制人
      'If bLvlChm Then
      If bLvO1 = True Or bLvl5 = True Then
         If bLvlChm Or bLvlJpn Or bLvlMot Or bLvlEls Then
             stConPAGrp = " and (PA150='" & PUB_GetStaffST16(stUserID) & "' or PA150 IS NULL) and cp10 not in('412','417')"
             stConSPGrp = " and (SP79='" & PUB_GetStaffST16(stUserID) & "' or SP79 is null) and cp10 not in('412','417')"
         Else
             stConPAGrp = " and PA150 IS NULL and cp10 not in('412','417')"
             stConSPGrp = " and SP79 is null and cp10 not in('412','417')"
         End If
      ElseIf bLvlChm Then
      'end 2022/12/20
         stConPAGrp = " and PA150='2' and cp10 not in('412','417')"
         stConSPGrp = " and SP79='2' and cp10 not in('412','417')"
      ElseIf bLvlJpn Then
         'Modified by Lydia 2021/02/25 未分組也未分案給王協理看
         'stConPAGrp = " and PA150='3'"
         'stConSPGrp = " and SP79='3'"
         'Modified by Lydia 2022/12/20 改回日文組
         'stConPAGrp = " and (PA150='3' or PA150 IS NULL) and cp10 not in('412','417')"
         'stConSPGrp = " and (SP79='3' or SP79 is null) and cp10 not in('412','417')"
         'end 2021/02/25
         stConPAGrp = " and PA150='3' and cp10 not in('412','417')"
         stConSPGrp = " and SP79='3' and cp10 not in('412','417')"
      '2011/11/30 modify by sonia 取消德文且機電拆電子電機及機械設計
      'ElseIf bLvlEls Then
      '   stConPAGrp = " and ((PA150<>'2' AND PA150<>'3') OR PA150 IS NULL)"
      '   stConSPGrp = " and ((SP79<>'2' and SP79<>'3') or SP79 is null)"
      ElseIf bLvlMot Then
         stConPAGrp = " and PA150='4' and cp10 not in('412','417')"
         stConSPGrp = " and SP79='4' and cp10 not in('412','417')"
      ElseIf bLvlEls Then
         'Modified by Lydia 2021/02/25 最早以前外專工程師主管是阮威立85030為電子電機組現在改成日文組
         'stConPAGrp = " and ((PA150>='2' AND PA150<='4') OR PA150 IS NULL)"
         'stConSPGrp = " and ((SP79>='2' and SP79<='4') or SP79 is null)"
         stConPAGrp = " and PA150='1' and cp10 not in('412','417')"
         stConSPGrp = " and SP79='1' and cp10 not in('412','417')"
      '2011/11/30 end
      End If
      
      '已完稿無核稿人,2個工作天後達核稿期限者(不含當日)-C8(核稿期限,無核稿人)
      'Add By Sindy 2016/8/1 Sharon:新案翻譯,達核槁未Key核稿人,請改為彈工程師主管
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' AND CP10='201' => CP01='FCP' AND CP10='201' AND CP158=0 AND CP159=0
      'Modify By Sindy 2023/8/31 CP01='FCP' => substr(cp12,1,1)='F' + & stCP01 &
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','C' EV1,'8' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT" & _
         " WHERE substr(cp12,1,1)='F' AND CP10='201' AND CP158=0 AND CP159=0 AND CP05>" & stDate1 & stCP01 & _
         " AND EP02(+)=CP09 AND EP04 IS NULL AND EP09>0 AND EP08>=" & stDate1 & " AND EP08<" & stDate2 & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & stConPAGrp
      cnnConnection.Execute stSQL, intI
      
      'Add By Sindy 2016/8/1 Sharon:新案翻譯,達核槁未Key核稿人,請改為彈工程師主管
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FG' AND CP10='201' => CP01='FG' AND CP10='201' AND CP158=0 AND CP159=0
      'Modify By Sindy 2023/8/31 CP01='FG' => substr(cp12,1,1)='F' + & stCP01 &
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','C' EV1,'8' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT" & _
         " WHERE substr(cp12,1,1)='F' AND CP10='201' AND CP158=0 AND CP159=0 AND CP05>" & stDate1 & stCP01 & _
         " AND EP02(+)=CP09 AND EP04 IS NULL AND EP09>0 AND EP08>=" & stDate1 & " AND EP08<" & stDate2 & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)" & stConSPGrp
      cnnConnection.Execute stSQL, intI
      
      '已收文未發,2個工作天後達本所期限者(不含當日)-A0(本所期限,未分案)
      'Modified by Lydia 2016/09/14  CP01||CP14||CP27||CP57='FCP' => CP01||CP14='FCP' AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
      'Modified by Lydia 2018/02/12 客戶提供文件1920 =>'M','待處理'
      'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "',DECODE(CP10,'1920','M','A') EV1,'0' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,PATENT,FAGENT" & _
         " WHERE substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & stConCP06 & stCP01 & _
         " AND PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & stConPAGrp & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
      cnnConnection.Execute stSQL, intI
      'Add By Sindy 2021/4/21 達指定
      'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','N' EV1,'0' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,PATENT,FAGENT" & _
         " WHERE substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & stConCP142 & stCP01 & _
         " AND PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & stConPAGrp & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
      cnnConnection.Execute stSQL, intI
      
      'Modified by Lydia 2016/09/14 CP01||CP14||CP27||CP57='FG' =>  CP01||CP14='FG' AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
      'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS, SERVICEPRACTICE, FAGENT" & _
         " WHERE CP05>20030000 AND substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & stConCP06 & stCP01 & _
         " AND SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & stConSPGrp & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
      cnnConnection.Execute stSQL, intI
      'Add By Sindy 2021/4/21 達指定
      'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','N' EV1,'0' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS, SERVICEPRACTICE, FAGENT" & _
         " WHERE CP05>20030000 AND substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & stConCP142 & stCP01 & _
         " AND SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & stConSPGrp & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
      cnnConnection.Execute stSQL, intI
      
      '已收文未發文,2個工作天後達承辦期限者(不含當日)-B0(承辦期限,未分案)
      'Modified by Lydia 2016/09/14  CP01||CP14||CP27||CP57='FCP' => CP01||CP14='FCP' AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
      'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
      'Modify By Sindy 2024/1/3 取消 AND EP09 IS NULL
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','B' EV1,'0' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
         " From CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT" & _
         " WHERE substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & stConCP48 & stCP01 & _
         " AND EP02(+)=CP09" & _
         " AND PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & stConPAGrp & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
      cnnConnection.Execute stSQL, intI
      
      'Modified by Lydia 2016/09/14 CP01||CP14||CP27||CP57='FG' =>  CP01||CP14='FG' AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
      'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','B' EV1,'0' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
         " From CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
         " WHERE substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & stConCP48 & stCP01 & _
         " AND SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & stConSPGrp & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
      cnnConnection.Execute stSQL, intI
      
      '未分案,所有未發文-E(未發文)
      If idx = 1 Then
         'Modified by Lydia 2016/09/14  CP01||CP14||CP27||CP57='FCP' => CP01||CP14='FCP' AND CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
         'Modified by Lydia 2018/02/12  客戶提供文件1920=>'M','待處理'
         'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "',DECODE(CP10,'1920','M','E') EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS,PATENT,FAGENT" & _
            " WHERE CP05>20030000 AND substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & stCP01 & _
            " AND PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & stConPAGrp & _
            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
         cnnConnection.Execute stSQL, intI
         
         'Modified by Lydia 2016/09/14 CP01||CP14||CP27||CP57='FG' =>  CP01||CP14='FG' AND CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
         'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
            " WHERE substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & stCP01 & _
            " AND SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & stConSPGrp & _
            " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
         cnnConnection.Execute stSQL, intI
      End If
   End If
      
   If bLvlX = True Then
      'Modify by Morgan 2008/10/9 未交稿加判斷無核稿期限的(會有例外狀況需核完稿才給翻譯費故完稿日會先拿掉,如巨京)
      
      '未交稿,2個工作天後達承辦期限者(不含當日)-G9(未交稿)
      'Modify by Morgan 2008/10/21 改當日到期的
      'Modify By Sindy 2016/7/28 + 有關翻譯之未交稿期限(下班翻改彈所內編號,外翻仍維持彈系統特殊人員"M") : AND not exists (select ts1.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')
      'Modified by Lydia 2016/09/22 CP10||CP27||CP57='201' => CP10='201' AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
      'Modify By Sindy 2021/4/15 cp14==>decode(SIM01,null,CP14,SIM01)
      '   所內翻譯也要列出來 mark:AND not exists (select ts1.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','G' EV1,'9' EV2,CP09,FA10,decode(SIM01,null,CP14,SIM01),CP13,CP06,CP07,CP48,NULL,0" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT,STAFF_IDMAP,STAFF S1,STAFF S2" & _
         " WHERE CP10='201' AND CP158=0 AND CP159=0 and substr(CP12,1,1)='F' AND CP48>" & stDate1 & " AND CP48<=" & strSrvDate(1) & stCP01 & _
         " AND EP02(+)=CP09 AND EP09 IS NULL AND EP08 IS NULL" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND S1.ST01(+)=CP14 AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01" & _
         " AND SUBSTRB(S1.ST15,1,1)='F' AND (S2.ST15 IS NULL OR SUBSTRB(S2.ST15,1,1)='F')" '& _
         " AND not exists (select ts1.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')"
      cnnConnection.Execute stSQL, intI
      
      'Modify By Sindy 2016/7/28 + 有關翻譯之未交稿期限(下班翻改彈所內編號,外翻仍維持彈系統特殊人員"M")
      'Modified by Lydia 2016/09/22 CP10||CP27||CP57='201' => CP10='201' AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
      'Modify By Sindy 2021/4/15 cp14==>decode(SIM01,null,CP14,SIM01)
      '   所內翻譯也要列出來 mark:AND not exists (select ts1.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','G' EV1,'9' EV2,CP09,FA10,decode(SIM01,null,CP14,SIM01),CP13,CP06,CP07,CP48,NULL,0" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT,STAFF_IDMAP,STAFF S1,STAFF S2" & _
         " WHERE CP10='201' AND CP158=0 AND CP159=0 and substr(CP12,1,1)='F' AND CP48>" & stDate1 & " AND CP48<=" & strSrvDate(1) & stCP01 & _
         " AND EP02(+)=CP09 AND EP09 IS NULL AND EP08 IS NULL" & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)" & _
         " AND S1.ST01(+)=CP14 AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01" & _
         " AND SUBSTRB(S1.ST15,1,1)='F' AND (S2.ST15 IS NULL OR SUBSTRB(S2.ST15,1,1)='F')" '& _
         " AND not exists (select ts1.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')"
      cnnConnection.Execute stSQL, intI
      
      'Modify By Sindy 2021/4/22 O.未分案,案件性質是412延緩公告417提早公開(沒有本所期限,法定期限,承辦期限)
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','O' EV1,'0' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,PATENT,FAGENT" & _
         " WHERE CP01||CP14='FCP' AND CP158=0 AND CP159=0" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) and cp10 in('412','417')"
      cnnConnection.Execute stSQL, intI
      
      '未交稿,所有未發文-E(只抓有承辦期限的)
      If idx = 1 Then
         'Modified by Lydia 2016/09/22 CP10||CP27||CP57='201' => CP10='201' AND CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
         'Modify By Sindy 2021/4/15 cp14==>decode(SIM01,null,CP14,SIM01)
         'Modify By Sindy 2025/10/13 出現重覆:+ " AND not exists(SELECT R01,R02,R03,R04,R12 FROM R060204 WHERE R01='" & strUserNum & "' and R02='E' and R03=' ' and R04=CP09 and R12=0)"
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,decode(SIM01,null,CP14,SIM01),CP13,CP06,CP07,CP48,NULL,0" & _
            " FROM CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT,STAFF_IDMAP,STAFF S1,STAFF S2" & _
            " WHERE CP10='201' AND CP158=0 AND CP159=0 and substr(CP12,1,1)='F' AND CP48>" & stDate1 & " AND CP48>=" & stDate2 & stCP01 & _
            " AND EP02(+)=CP09 AND EP09 IS NULL AND EP08 IS NULL" & _
            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
            " AND S1.ST01(+)=CP14 AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01" & _
            " AND SUBSTRB(S1.ST15,1,1)='F' AND (S2.ST15 IS NULL OR SUBSTRB(S2.ST15,1,1)='F')" & _
            " AND not exists(SELECT R01,R02,R03,R04,R12 FROM R060204 WHERE R01='" & strUserNum & "' and R02='E' and R03=' ' and R04=CP09 and R12=0)"
         cnnConnection.Execute stSQL, intI
         
         'Modified by Lydia 2016/09/22 CP10||CP27||CP57='201' => CP10='201' AND CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
         'Modify By Sindy 2021/4/15 cp14==>decode(SIM01,null,CP14,SIM01)
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,decode(SIM01,null,CP14,SIM01),CP13,CP06,CP07,CP48,NULL,0" & _
            " FROM CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT,STAFF_IDMAP,STAFF S1,STAFF S2" & _
            " WHERE CP10='201' AND CP158=0 AND CP159=0 and substr(CP12,1,1)='F' AND CP48>" & stDate1 & " AND CP48>=" & stDate2 & stCP01 & _
            " AND EP02(+)=CP09 AND EP09 IS NULL AND EP08 IS NULL" & _
            " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)" & _
            " AND S1.ST01(+)=CP14 AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01" & _
            " AND SUBSTRB(S1.ST15,1,1)='F' AND (S2.ST15 IS NULL OR SUBSTRB(S2.ST15,1,1)='F')"
         cnnConnection.Execute stSQL, intI
      End If
      
      '已完稿無核稿人,所有未發文-E(只抓有核稿期限)
      If idx = 1 Then
         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' AND CP10='201' => CP01='FCP' AND CP10='201' AND CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
         'Modify By Sindy 2023/8/31 CP01='FCP' => substr(cp12,1,1)='F'
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,NULL,0" & _
            " FROM CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT" & _
            " WHERE substr(cp12,1,1)='F' AND CP10='201' AND CP158=0 AND CP159=0 AND CP05>" & stDate1 & stCP01 & _
            " AND EP02(+)=CP09 AND EP04 IS NULL AND EP09>0 AND EP08>=" & stDate2 & _
            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
         cnnConnection.Execute stSQL, intI
         
         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FG' AND CP10='201' => CP01='FG' AND CP10='201' AND CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
         'Modify By Sindy 2023/8/31 CP01='FG' => substr(cp12,1,1)='F'
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,NULL,0" & _
            " FROM CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT" & _
            " WHERE substr(cp12,1,1)='F' AND CP10='201' AND CP158=0 AND CP159=0 AND CP05>" & stDate1 & stCP01 & _
            " AND EP02(+)=CP09 AND EP04 IS NULL AND EP09>0 AND EP08>=" & stDate2 & _
            " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
         cnnConnection.Execute stSQL, intI
      End If
   End If
                  
   '程序
   If stDept = "F22" Then
      '已收文未發文且 2個工作天 後達本所期限者(不含當日) --管制人-A2(本所期限,管制人)
      'Modify by Morgan 2008/11/19 +年費另外
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10 not in('605','926','945') => CP01='FCP' and CP10 not in('605','926','945') AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/23 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
      'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人) ('605','926','945')-->('926','945')
      'Modify By Sindy 2023/8/31 CP01='FCP' => substr(cp12,1,1)='F' + & stCP01 &
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,PATENT,FAGENT,Nation" & _
         " Where substr(cp12,1,1)='F' and CP10 not in('926','945') AND CP158=0 AND CP159=0" & stConCP06 & stCP01 & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
      cnnConnection.Execute stSQL, intI
      'Add By Sindy 2021/4/21 達指定
      'Modify By Sindy 2023/8/31 CP01='FCP' => substr(cp12,1,1)='F' + & stCP01 &
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','N' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,PATENT,FAGENT,Nation" & _
         " Where substr(cp12,1,1)='F' and CP10 not in('926','945') AND CP158=0 AND CP159=0" & stConCP142 & stCP01 & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
      cnnConnection.Execute stSQL, intI
      
      'Add By Sindy 2021/3/17 + 核對已准專利已發文未請款->請彈程序
      'Modify By Sindy 2023/8/31 CP01='FCP' => substr(cp12,1,1)='F' + & stCP01 &
'      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),CPM19),'YYYYMMDD'),WORKDAYADD(cpm33+1,CP27))"
'      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','F' EV1,'2' EV2,CP09,FA10,CP14,CP13" & _
'         ",NULL,NULL," & strCPM1933_Col & " C06,NULL,0" & _
'         " FROM CASEPROGRESS,PATENT,FAGENT,Nation,CASEPROPERTYMAP" & _
'         " WHERE CP27>" & stDate1 & " AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'" & stCP01 & _
'         " And substr(cp12,1,1)='F' And CP10='926'" & _
'         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16 & _
'         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP27<=" & strCPM1933_Where
'      cnnConnection.Execute stSQL, intI
      '2021/3/17 END
      'Modify By Sindy 2025/10/9 修改未請款程式
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(decode(CP01,'FCP',CP27,CP47),'YYYYMMDD'),CPM19),'YYYYMMDD'),WORKDAYADD(cpm33+1,decode(CP01,'FCP',CP27,CP47)))"
      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'2' EV2,CP09,FA10,CP14,CP13" & _
         ",NULL,NULL," & strCPM1933_Col & " C06,NULL,0" & _
         " FROM CASEPROGRESS,PATENT,FAGENT,Nation,CASEPROPERTYMAP" & _
         " WHERE CP27>" & stDate1 & " AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'" & stCP01 & _
         " And substr(cp12,1,1)='F' And CP10='926'" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16 & _
         " AND CPM01(+)='FCP' AND CPM02(+)=CP10" & _
         " AND (CP01='FCP' or (CP01<>'FCP' AND CP47 is not null))" & _
         " AND nvl(CP47,CP27)<=" & strCPM1933_Where
      cnnConnection.Execute stSQL, intI
      '2025/10/9 END
      
      'Add By Sindy 2015/10/22 +法務:達本所,僅限收文業務區CP12=F2字頭
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57 in('FCL','CFL','LIN') => CP01 in('FCL','CFL','LIN') AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/23 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,nvl(FA10,cu10),CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,Lawcase,FAGENT,Nation n1,Customer,Nation n2" & _
         " Where CP01 in('FCL','CFL','LIN') AND CP158=0 AND CP159=0" & stConCP06 & _
         " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC08||LC34 IS NULL" & _
         " AND FA01(+)=SUBSTR(LC22,1,8) AND FA02(+)=SUBSTR(LC22,9) AND n1.NA01(+)=FA10" & _
         " AND CU01(+)=SUBSTR(LC11,1,8) AND CU02(+)=SUBSTR(LC11,9) AND n2.NA01(+)=CU10" & _
         " AND (CP14 IS NULL OR CP14<>n1.NA16 OR CP14<>n2.NA16)" & stConNA16L & _
         " AND substr(CP12,1,2)='F2'"
      cnnConnection.Execute stSQL, intI
      '2015/10/22 END
      
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FG' => CP01='FG' AND CP158=0 AND CP159=0
      'Modify By Sindy 2023/8/31 CP01='FCP' => substr(cp12,1,1)='F' + & stCP01 &
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS, SERVICEPRACTICE, FAGENT, Nation" & _
         " Where substr(cp12,1,1)='F' AND CP158=0 AND CP159=0" & stConCP06 & stCP01 & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP26 IS NOT NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
      cnnConnection.Execute stSQL, intI
      'Add By Sindy 2021/4/21 達指定
      'Modify By Sindy 2023/8/31 CP01='FCP' => substr(cp12,1,1)='F' + & stCP01 &
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','N' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS, SERVICEPRACTICE, FAGENT, Nation" & _
         " Where substr(cp12,1,1)='F' AND CP158=0 AND CP159=0" & stConCP142 & stCP01 & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP26 IS NOT NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
      cnnConnection.Execute stSQL, intI
      
      '所有未發文--管制人-E(未發文)
      If idx = 1 Then
         'Modify by Morgan 2008/11/19 +年費管制人另外有規則
         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10 not in('605','926','945') => CP01='FCP' and CP10 not in('605','926','945') AND CP158=0 AND CP159=0
         'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人) ('605','926','945')-->('926','945')
         'Modify By Sindy 2023/8/31 CP01='FCP' => substr(cp12,1,1)='F' + & stCP01 &
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS, PATENT, FAGENT, Nation" & _
            " Where substr(cp12,1,1)='F' and CP10 not in('926','945') AND CP158=0 AND CP159=0" & stCP01 & _
            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND PA01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
         'Added by Lydia 2018/02/12 排除重複項目
         stSQL = stSQL & " AND CP09 NOT IN (SELECT R04 FROM R060204 WHERE R01='" & strUserNum & "' AND R02='E' AND R03 = ' ' AND R12=0) "
         cnnConnection.Execute stSQL, intI
         
         'Add By Sindy 2015/10/22 +法務:未發文,僅限收文業務區CP12=F2字頭
         'Modified by Lydia 2016/09/14 CP01||CP27||CP57 in('FCL','CFL','LIN') => CP01 in('FCL','CFL','LIN') AND CP158=0 AND CP159=0
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,nvl(FA10,cu10),CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS, Lawcase, FAGENT, Nation n1,Customer, Nation n2" & _
            " Where CP01 in('FCL','CFL','LIN') AND CP158=0 AND CP159=0" & _
            " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC08||LC34 IS NULL" & _
            " AND FA01(+)=SUBSTR(LC22,1,8) AND FA02(+)=SUBSTR(LC22,9) AND n1.NA01(+)=FA10" & _
            " AND CU01(+)=SUBSTR(LC11,1,8) AND CU02(+)=SUBSTR(LC11,9) AND n2.NA01(+)=CU10" & _
            " AND (CP14 IS NULL OR CP14<>n1.NA16 OR CP14<>n2.NA16)" & stConNA16L & _
            " AND substr(CP12,1,2)='F2'"
         'Added by Lydia 2018/02/12 排除重複項目
         stSQL = stSQL & " AND CP09 NOT IN (SELECT R04 FROM R060204 WHERE R01='" & strUserNum & "' AND R02='E' AND R03 = ' ' AND R12=0) "
         cnnConnection.Execute stSQL, intI
         '2015/10/22 END
         
         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FG' => CP01='FG' AND CP158=0 AND CP159=0
         'Modify By Sindy 2023/8/31 CP01='FCP' => substr(cp12,1,1)='F' + & stCP01 &
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS, SERVICEPRACTICE, FAGENT, Nation" & _
            " Where substr(cp12,1,1)='F' AND CP158=0 AND CP159=0" & stCP01 & _
            " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP26 IS NOT NULL AND SP01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
         'Added by Lydia 2018/02/12 排除重複項目
         stSQL = stSQL & " AND CP09 NOT IN (SELECT R04 FROM R060204 WHERE R01='" & strUserNum & "' AND R02='E' AND R03 = ' ' AND R12=0) "
         cnnConnection.Execute stSQL, intI
      End If
   
      '未收文且 2個工作天 後達本所期限者(不含當日) --管制人-D2(未收文,管制人)
      'Modify by Morgan 2008/11/19 +年費另外
      'Modify by Morgan 2009/7/13 +995,996
      'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
      'Modified by Lydia 2014/11/14 NA51 = decode(pa75,'Y51333010','" & midStr & "',na51)
      'Modified by Lydia 2016/02/03 +Y51817040
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807) */
      'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人) ('605','926','945')-->('926','945')
      'Modify By Sindy 2021/4/28 + ,R15
      'Modified by Lydia 2022/11/03 區分FMP案; stConNA16=> Replace(UCase(stConNA16), "CP01", "NP02")
      'Modify By Sindy 2023/8/31 +FMP 拿掉NP02||NP06='FCP'=>NP06 IS NULL
      '                               +,CASEPROGRESS +AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F'" & stCP01
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,FA10,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS, PATENT, FAGENT, Nation,CASEPROGRESS" & _
         " WHERE NP06 IS NULL and np07 not in('926','945')" & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate2 & _
         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F'" & stCP01 & _
         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & Replace(UCase(stConNA16), "CP01", "NP02") & _
         " AND NOT EXISTS(SELECT * FROM CASEPROGRESS WHERE CP09=NP01 AND CP10='928' AND NP07='202')"
      cnnConnection.Execute stSQL, intI
     
      'Add By Sindy 2015/10/22 +法務:程序組的未收文,以管制人角度顯示,僅限下一程序智權人員之ST15=F2字頭
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807) */
      'Modify By Sindy 2021/4/28 + ,R15
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(FA10,cu10),NULL,decode(FA10,null,n2.na51,decode(LC22," & midStr & ",n1.na51)) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS, Lawcase, FAGENT, Nation n1, Staff, Customer, Nation n2" & _
         " WHERE NP02||NP06 in('FCL','CFL','LIN')" & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate2 & _
         " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05 AND LC08||LC34 IS NULL" & _
         " AND FA01(+)=SUBSTR(LC22,1,8) AND FA02(+)=SUBSTR(LC22,9) AND n1.NA01(+)=FA10" & _
         " AND CU01(+)=SUBSTR(LC11,1,8) AND CU02(+)=SUBSTR(LC11,9) AND n2.NA01(+)=cu10" & stConNA16L & _
         " AND NP10=ST01(+) AND substr(ST15,1,2)='F2'"
      cnnConnection.Execute stSQL, intI
      '2015/10/22 END
      
      'end 2008/11/19
      'Modify by Morgan 2009/7/13 +995,996
      'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
      'Modified by Lydia 2014/11/14 NA51 = decode(sp26,'Y51333010','" & midStr & "',na51)
      'Modified by Lydia 2016/02/03
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807)*/
      'Modify By Sindy 2021/4/28 + ,R15
      'Modified by Lydia 2022/11/03 區分FMP案; stConNA16=> Replace(UCase(stConNA16), "CP01", "NP02")
      'Modify By Sindy 2023/8/31 +FMP 拿掉NP02||NP06='FCP'=>NP06 IS NULL
      '                               +,CASEPROGRESS +AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F'" & stCP01
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,FA10,NULL,decode(sp26," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS, SERVICEPRACTICE, FAGENT, Nation,CASEPROGRESS" & _
         " WHERE NP06 IS NULL" & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate2 & _
         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F'" & stCP01 & _
         " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP15||SP61 IS NULL AND SP26 IS NOT NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10" & Replace(UCase(stConNA16), "CP01", "NP02")
      cnnConnection.Execute stSQL, intI
      'end       'Modified by Lydia 2014/11/14 NA51 = decode(pa75,'Y51333010','" & midStr & "',na51)
   End If
   
   '工程師
   If (stDept <> "F22" And stDept <> "F23") Or bLvl4 = True Or bLvl5 = True Then
      '未交稿加判斷無核稿期限的(會有例外狀況需核完稿才給翻譯費故完稿日會先拿掉,如巨京)
      '未交稿,2個工作天後達承辦期限者(不含當日)-G9(未交稿)
      '改當日到期的
      'Add By Sindy 2016/7/28 + 有關翻譯之未交稿期限(下班翻改彈所內編號,外翻仍維持彈系統特殊人員"M"
      'Modified by Lydia 2016/09/22 CP10||CP27||CP57='201' => CP10='201' AND CP158=0 AND CP159=0
      'Modified by Lydia 2019/10/04 改成判斷畫面的員工編號 strUserNum=>strUserId
      'Modify By Sindy 2021/4/15 cp14==>decode(SIM01,null,CP14,SIM01)
      '   新案翻譯承辦人為所內工程師(上班譯-員編,下班譯-F編號)，請彈承辦工程師及其主管、Sharon 原:AND '" & stUserID & "' in (select ts2.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')
      '   原:SUBSTRB(S2.ST15,1,1)='F' ==> S1.ST15='F52'
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','G' EV1,'9' EV2,CP09,FA10,decode(SIM01,null,CP14,SIM01),CP13,CP06,CP07,CP48,NULL,0" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT,STAFF_IDMAP,STAFF S1,STAFF S2" & _
         " WHERE CP10='201' AND CP158=0 AND CP159=0 and substr(CP12,1,1)='F' AND CP48>" & stDate1 & " AND CP48<=" & strSrvDate(1) & stCP01 & _
         " AND EP02(+)=CP09 AND EP09 IS NULL AND EP08 IS NULL" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND S1.ST01(+)=CP14 AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01" & _
         " AND SUBSTRB(S1.ST15,1,1)='F' AND (S2.ST15 IS NULL OR S1.ST15='F52')" & _
         " AND EXISTS(select ts2.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')" & _
         " AND decode(SIM01,null,CP14,SIM01) in (" & stNumList & ")"
      cnnConnection.Execute stSQL, intI
      'Add By Sindy 2016/7/28 + 有關翻譯之未交稿期限(下班翻改彈所內編號,外翻仍維持彈系統特殊人員"M")
      'Modified by Lydia 2016/09/22 CP10||CP27||CP57='201' => CP10='201' AND CP158=0 AND CP159=0
      'Modified by Lydia 2019/10/04 改成判斷畫面的員工編號 strUserNum=>strUserId
      'Modify By Sindy 2021/4/15 cp14==>decode(SIM01,null,CP14,SIM01)
      '   新案翻譯承辦人為所內工程師(上班譯-員編,下班譯-F編號)，請彈承辦工程師及其主管、Sharon 原:AND '" & stUserID & "' in (select ts2.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')
      '   原:SUBSTRB(S2.ST15,1,1)='F' ==> S1.ST15='F52'
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','G' EV1,'9' EV2,CP09,FA10,decode(SIM01,null,CP14,SIM01),CP13,CP06,CP07,CP48,NULL,0" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT,STAFF_IDMAP,STAFF S1,STAFF S2" & _
         " WHERE CP10='201' AND CP158=0 AND CP159=0 and substr(CP12,1,1)='F' AND CP48>" & stDate1 & " AND CP48<=" & strSrvDate(1) & stCP01 & _
         " AND EP02(+)=CP09 AND EP09 IS NULL AND EP08 IS NULL" & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)" & _
         " AND S1.ST01(+)=CP14 AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01" & _
         " AND SUBSTRB(S1.ST15,1,1)='F' AND (S2.ST15 IS NULL OR S1.ST15='F52')" & _
         " AND EXISTS(select ts2.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')" & _
         " AND decode(SIM01,null,CP14,SIM01) in (" & stNumList & ")"
      cnnConnection.Execute stSQL, intI

      'Modify by Morgan 2008/10/9 未核稿不必判斷完稿日
      '未核稿且 2個工作天 後達核稿期限者(不含當日) --核稿人-C4(核稿期限,核稿人)
      'Modified by Lydia 2016/09/14 CP27||CP57 IS NULL => CP158=0 AND CP159=0
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','C' EV1,'4' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
         " FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,FAGENT" & _
         " WHERE " & IIf(strSrvDate(1) >= FCP核完日改用EP39, "EP39 IS NULL", "EP33 IS NULL") & stConEP08 & stConEP04 & _
         " AND CP09(+)=EP02 AND CP158=0 AND CP159=0 AND CP10='201'" & stCP01 & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
      cnnConnection.Execute stSQL, intI
         
      'Modified by Lydia 2016/09/14 CP27||CP57 IS NULL => CP158=0 AND CP159=0
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','C' EV1,'4' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
         " FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
         " WHERE " & IIf(strSrvDate(1) >= FCP核完日改用EP39, "EP39 IS NULL", "EP33 IS NULL") & stConEP08 & stConEP04 & _
         " AND CP09(+)=EP02 AND CP158=0 AND CP159=0 AND CP10='201'" & stCP01 & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
      cnnConnection.Execute stSQL, intI

      '未核稿,所有未發文--核稿人-E(未發文)
      If idx = 1 Then
         'Modified by Lydia 2016/09/14 CP27||CP57 IS NULL => CP158=0 AND CP159=0
         'Modified by Lydia 2018/11/29 FCP-59635的新案翻譯未發文,A0022為承辦人主管同時為核稿人,造成重複主鍵,設R03="4"核稿人(' ' EV2=>'4' EV2)
         'Modify By Sindy 2024/1/30 取消: " & IIf(strSrvDate(1) >= FCP核完日改用EP39, "EP39 IS NULL", "EP33 IS NULL") & " and
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,'4' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
            " FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,FAGENT" & _
            " WHERE EP08>=" & stDate2 & stConEP04 & _
            " AND CP09(+)=EP02 AND CP158=0 AND CP159=0 AND CP10='201'" & stCP01 & _
            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND EP04<>CP14"
         cnnConnection.Execute stSQL, intI
         
         'Modified by Lydia 2016/09/14 CP27||CP57 IS NULL => CP158=0 AND CP159=0
         'Modified by Lydia 2018/11/29 FCP-59635的新案翻譯未發文,A0022為承辦人主管同時為核稿人,造成重複主鍵,設R03="4"核稿人(' ' EV2=>'4' EV2)
         'Modify By Sindy 2024/1/30 取消: " & IIf(strSrvDate(1) >= FCP核完日改用EP39, "EP39 IS NULL", "EP33 IS NULL") & " and
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,'4' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
            " FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
            " WHERE EP08>=" & stDate2 & stConEP04 & _
            " AND CP09(+)=EP02 AND CP158=0 AND CP159=0 AND CP10='201'" & stCP01 & _
            " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND EP04<>CP14"
         cnnConnection.Execute stSQL, intI
      End If
   End If
   
   '業務(FCP,FG抓NA51,其他抓NP10)
   If stDept = "F23" Or bLvl4 = True Or bLvl5 = True Then
      '未收文且 5個工作天 後達本所期限者(不含當日) --智權人員-D3(未收文,智權人員)
      'Modify by Morgan 2008/11/19 +年費另外
      'Modify by Morgan 2009/7/13 +995,996
      'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
      'Modified by Lydia 2014/11/14 NA51 = decode(pa75,'Y51333010','" & midStr & "',na51), stConNA51->stConNA51P
      'Modified by Lydia 2016/02/03
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807) */
      'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人)  取消 and np07<>'605' 條件
      'Modify By Sindy 2021/4/28 + ,R15
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,FA10,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS, PATENT, FAGENT, Nation" & _
         " WHERE NP02||NP06='FCP'" & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate3 & _
         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA51P & _
         " AND NOT (NP07='202' AND EXISTS(SELECT * FROM CASEPROGRESS WHERE CP09=NP01 AND CP10='928'))"
      cnnConnection.Execute stSQL, intI
      
      'Add By Sindy 2015/10/22 +法務:承辦業務組的未收文,僅限下一程序智權人員之ST15=F2字頭
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807) */
      'Modify By Sindy 2021/4/28 + ,R15
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(FA10,cu10),NULL,decode(FA10,null,n2.na51,decode(LC22," & midStr & ",n1.na51)) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS, LawCase, FAGENT, Nation n1, Staff,Customer, Nation n2" & _
         " WHERE NP02||NP06 in('FCL','CFL','LIN')" & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate3 & _
         " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05 AND LC08||LC34 IS NULL" & _
         " AND FA01(+)=SUBSTR(LC22,1,8) AND FA02(+)=SUBSTR(LC22,9) AND n1.NA01(+)=FA10" & _
         " AND CU01(+)=SUBSTR(LC11,1,8) AND CU02(+)=SUBSTR(LC11,9) AND n2.NA01(+)=cu10" & stConNA51L & _
         " AND NP10=ST01(+) AND substr(ST15,1,2)='F2'"
      cnnConnection.Execute stSQL, intI
      '2015/10/22 END
      'end 2016/09/22
      
      'Add by Morgan 2010/7/29 P,CFP也改抓國家檔
      'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
      'Modified by Lydia 2014/11/14
      'Modified by Lydia 2016/02/03
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807) */
      'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人)  取消 and np07<>'605' 條件
      'Modify By Sindy 2021/4/28 + ,R15
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,FA10,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS,CASEPROGRESS, PATENT, FAGENT, Nation" & _
         " WHERE NP02||NP06 in ('P','CFP')" & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate3 & _
         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F'" & _
         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA51P & _
         " AND NOT (NP07='202' AND EXISTS(SELECT * FROM CASEPROGRESS WHERE CP09=NP01 AND CP10='928'))"
      cnnConnection.Execute stSQL, intI

      'Modify by Morgan 2009/7/13 +995,996
      'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
      'Modified by Lydia 2014/11/14 NA51 = decode(sp26,'Y51333010','" & midStr & "',na51),stConNA51->stconsp26
      'Modified by Lydia 2016/02/03
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807) */
      'Modify By Sindy 2021/4/28 + ,R15
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,FA10,NULL,decode(sp26," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS, SERVICEPRACTICE, FAGENT, Nation" & _
         " WHERE NP02||NP06='FG'" & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate3 & _
         " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP15||SP61 IS NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10" & stConSP26
      cnnConnection.Execute stSQL, intI
         
      'Add by Morgan 2010/7/29 PS,CFS也改抓國家檔
      'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
      'Modified by Lydia 2016/02/03
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807) */
      'Modify By Sindy 2021/4/28 + ,R15
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,FA10,NULL,decode(sp26," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS,CASEPROGRESS, SERVICEPRACTICE, FAGENT, Nation" & _
         " WHERE NP02||NP06 IN ('PS,CPS')" & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate3 & _
         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F'" & _
         " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP15||SP61 IS NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10" & stConSP26
      cnnConnection.Execute stSQL, intI
      'end       'Modified by Lydia 2014/11/14 NA51 = decode(pa75,'Y51333010','" & midStr & "',na51)
      
      '未收文且 5個工作天 後達本所期限者(不含當日) --智權人員(非FCP,FG)-45(未收文,智權人員)
      'Memo by Morgan 2008/11/20 非國外部專利案會指定承辦人管制，暫不改--David
      'Modify by Morgan 2009/7/13 +995,996
      'Remove by Morgan 2010/7/29改抓國家檔,移到上面
      'Modify by Morgan 2009/7/13 +995,996
      'Remove by Morgan 2010/7/29改抓國家檔,移到上面
      '寄中說949已收文未發文且 2個工作天 後達本所期限者(不含當日) --智權人員-A2(本所期限,智權人員)
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10='949' => CP01='FCP' and CP10='949' AND CP158=0 AND CP159=0
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,PATENT,FAGENT,Nation" & _
         " Where CP01='FCP' and CP10='949' AND CP158=0 AND CP159=0" & stConCP06 & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA51
      cnnConnection.Execute stSQL, intI
   End If
   
   'Add By Sindy 2022/3/15
   '(1)通知對象為承辦及程序
   If stDept = "F22" Or stDept = "F23" Or bLvl4 = True Or bLvl5 = True Then
      '111.03.03-Sharon - FCP
      '以上104年請作單將已發文未請款未彈期限通知排除是非常危險,故需將此設定調整
      '以下案件性質已發文未請款需彈期限通知:
      '(2)發文日+7個工作天未請款案件
      '(3)排除已上"不請款"
      '201新案翻譯、209檢視中說、210製作中說、401變更、403更改、416實體審查、601領證及繳年費、605年費、701讓與、702合併、917超頁、超項費
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'" & IIf(stDept = "F22", "2", "1") & "' EV2,CP09,FA10,CP14,CP13" & _
         ",NULL,NULL,workdayadd(8,cp27) C06,NULL,0" & _
         " FROM CASEPROGRESS,PATENT,FAGENT,CASEPROPERTYMAP,Nation" & _
         " WHERE CP01='FCP' AND CP27>" & stDate1 & "-10000 AND CP20||CP57||CP60 IS NULL AND SUBSTR(CP12,1,1)='F'" & _
         " AND CP10 in('201','209','210','401','403','416','601','605','701','702','917')" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
         " AND CP27<=" & stDate7 & _
         " AND NA01(+)=FA10" & IIf(stDept = "F22", " AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16, stConNA51P) & _
         " AND not exists(SELECT R01,R02,R03,R04,R12 FROM R060204 WHERE R01='" & strUserNum & "' and R02='F' and R03='" & IIf(stDept = "F22", "2", "1") & "' and R04=CP09 and R12=0)"
      cnnConnection.Execute stSQL, intI
      'end 2014/8/27
      '(4)235核對中說格式：此性質原本就已設為"不請款"，請判斷235核對中說格式發文日+7個工作天，
      '提申那道(101發明申請、102新型申請) 未請款案件，則需彈期限通知
      'Modified by Lydia 2022/11/03 區分FMP案; stConNA16=> Replace(UCase(stConNA16), "CP01", "C1.CP01")
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'" & IIf(stDept = "F22", "2", "1") & "' EV2,c1.CP09,FA10,c1.CP14,c1.CP13" & _
         ",NULL,NULL,workdayadd(8,c1.cp27) C06,NULL,0" & _
         " FROM CASEPROGRESS c1,PATENT,FAGENT,CASEPROPERTYMAP,Nation,CASEPROGRESS c2" & _
         " WHERE c1.CP01='FCP' AND c1.CP27>" & stDate1 & "-10000 AND c1.CP57||c1.CP60 IS NULL AND SUBSTR(c1.CP12,1,1)='F'" & _
         " AND c1.CP10='235'" & _
         " AND PA01(+)=c1.CP01 AND PA02(+)=c1.CP02 AND PA03(+)=c1.CP03 AND PA04(+)=c1.CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=c1.CP01 AND CPM02(+)=c1.CP10" & _
         " AND c1.CP27<=" & stDate7 & _
         " AND NA01(+)=FA10" & IIf(stDept = "F22", " AND (c1.CP14 IS NULL OR c1.CP14<>NA16)" & Replace(UCase(stConNA16), "CP01", "C1.CP01"), stConNA51P) & _
         " AND c2.cp01=c1.cp01 AND c2.cp02=c1.cp02 AND c2.cp03=c1.cp03 AND c2.cp04=c1.cp04" & _
         " AND c2.cp10 in('101','102') AND c2.cp09=c1.cp43 AND c2.CP27 is not null AND c2.CP20||c2.CP57||c2.CP60 IS NULL" & _
         " AND not exists(SELECT R01,R02,R03,R04,R12 FROM R060204 WHERE R01='" & strUserNum & "' and R02='F' and R03='" & IIf(stDept = "F22", "2", "1") & "' and R04=c1.CP09 and R12=0)"
      cnnConnection.Execute stSQL, intI
      '110.03.16-增加FMP案的控管 - P,CFP
      '(1)性質：401變更、403更改、416實體審查、701讓與、702合併
      '以上Key已提申+7個工作天未請款案件
      'Modified by Lydia 2022/11/03  區分FMP案 CP14<>NA16=>NVL(NA79,NA16)
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'" & IIf(stDept = "F22", "2", "1") & "' EV2,CP09,FA10,CP14,CP13" & _
         ",NULL,NULL,workdayadd(8,CP47) C06,NULL,0" & _
         " FROM CASEPROGRESS,PATENT,FAGENT,CASEPROPERTYMAP,Nation" & _
         " WHERE CP01 in('P','CFP') AND CP27>" & stDate1 & "-10000 AND CP20||CP57||CP60 IS NULL AND SUBSTR(CP12,1,1)='F'" & _
         " AND CP10 in('401','403','416','701','702')" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP47 is not null" & _
         " AND CP47<=" & stDate7 & _
         " AND NA01(+)=FA10" & IIf(stDept = "F22", " AND (CP14 IS NULL OR CP14<>NVL(NA79,NA16))" & stConNA16, stConNA51P) & _
         " AND not exists(SELECT R01,R02,R03,R04,R12 FROM R060204 WHERE R01='" & strUserNum & "' and R02='F' and R03='" & IIf(stDept = "F22", "2", "1") & "' and R04=CP09 and R12=0)"
      cnnConnection.Execute stSQL, intI
      'Modify By Sindy 2022/3/22
      '性質：601領證及繳年費、605年費
      '要區分代理人為Y53374北京寰華知識產權代理有限公司,則維持上列(1)的控管
       'Modified by Lydia 2022/11/03  區分FMP案 CP14<>NA16=>NVL(NA79,NA16)
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'" & IIf(stDept = "F22", "2", "1") & "' EV2,CP09,FA10,CP14,CP13" & _
         ",NULL,NULL,workdayadd(8,CP47) C06,NULL,0" & _
         " FROM CASEPROGRESS,PATENT,FAGENT,CASEPROPERTYMAP,Nation" & _
         " WHERE CP01 in('P','CFP') AND CP27>" & stDate1 & "-10000 AND CP20||CP57||CP60 IS NULL AND SUBSTR(CP12,1,1)='F'" & _
         " AND CP10 in('601','605') AND cp44='Y53374000'" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP47 is not null" & _
         " AND CP47<=" & stDate7 & _
         " AND NA01(+)=FA10" & IIf(stDept = "F22", " AND (CP14 IS NULL OR CP14<>NVL(NA79,NA16))" & stConNA16, stConNA51P) & _
         " AND not exists(SELECT R01,R02,R03,R04,R12 FROM R060204 WHERE R01='" & strUserNum & "' and R02='F' and R03='" & IIf(stDept = "F22", "2", "1") & "' and R04=CP09 and R12=0)"
      cnnConnection.Execute stSQL, intI
      '非寰華案,就以1909已提申的發文日+7個工作天控管
      'Modified by Lydia 2022/11/03  區分FMP案 CP14<>NA16=>NVL(NA79,NA16), stConNA16=> Replace(UCase(stConNA16), "CP01", "C1.CP01")
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'" & IIf(stDept = "F22", "2", "1") & "' EV2,c1.CP09,FA10,c1.CP14,c1.CP13" & _
         ",NULL,NULL,workdayadd(8,c2.cp27),NULL,0" & _
         " FROM CASEPROGRESS c1,PATENT,FAGENT,CASEPROPERTYMAP,Nation,CASEPROGRESS c2" & _
         " WHERE c1.CP01 in('P','CFP') AND c1.CP27>" & stDate1 & "-10000 AND c1.CP20||c1.CP57||c1.CP60 IS NULL AND SUBSTR(c1.CP12,1,1)='F'" & _
         " AND c1.CP10 in('601','605') AND c1.cp44<>'Y53374000'" & _
         " AND PA01(+)=c1.CP01 AND PA02(+)=c1.CP02 AND PA03(+)=c1.CP03 AND PA04(+)=c1.CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=c1.CP01 AND CPM02(+)=c1.CP10 AND c1.CP47 is not null" & _
         " AND c2.cp01=c1.cp01 AND c2.cp02=c1.cp02 AND c2.cp03=c1.cp03 AND c2.cp04=c1.cp04" & _
         " AND c2.cp10='1909' AND c2.cp43=c1.cp09 AND c2.CP27 IS NOT NULL" & _
         " AND c2.CP27<=" & stDate7 & _
         " AND NA01(+)=FA10" & IIf(stDept = "F22", " AND (c1.CP14 IS NULL OR c1.CP14<>NVL(NA79,NA16))" & Replace(UCase(stConNA16), "CP01", "C1.CP01"), stConNA51P) & _
         " AND not exists(SELECT R01,R02,R03,R04,R12 FROM R060204 WHERE R01='" & strUserNum & "' and R02='F' and R03='" & IIf(stDept = "F22", "2", "1") & "' and R04=c1.CP09 and R12=0)"
      cnnConnection.Execute stSQL, intI
      '(2)通知申請案號(1101)發文日+7個工作天,
      '提申那道(101發明申請、102新型申請、103設計申請) 未請款案件，則需彈期限通知
      'Modified by Lydia 2022/11/03  區分FMP案 CP14<>NA16=>NVL(NA79,NA16), stConNA16=> Replace(UCase(stConNA16), "CP01", "C1.CP01")
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'" & IIf(stDept = "F22", "2", "1") & "' EV2,c1.CP09,FA10,c1.CP14,c1.CP13" & _
         ",NULL,NULL,workdayadd(8,c1.cp27) C06,NULL,0" & _
         " FROM CASEPROGRESS c1,PATENT,FAGENT,CASEPROPERTYMAP,Nation,CASEPROGRESS c2" & _
         " WHERE c1.CP01 in('P','CFP') AND c1.CP27>" & stDate1 & "-10000 AND c1.CP57||c1.CP60 IS NULL AND SUBSTR(c1.CP12,1,1)='F'" & _
         " AND c1.CP10='1101'" & _
         " AND PA01(+)=c1.CP01 AND PA02(+)=c1.CP02 AND PA03(+)=c1.CP03 AND PA04(+)=c1.CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=c1.CP01 AND CPM02(+)=c1.CP10" & _
         " AND c1.CP27<=" & stDate7 & _
         " AND NA01(+)=FA10" & IIf(stDept = "F22", " AND (c1.CP14 IS NULL OR c1.CP14<>NVL(NA79,NA16))" & Replace(UCase(stConNA16), "CP01", "C1.CP01"), stConNA51P) & _
         " AND c2.cp01=c1.cp01 AND c2.cp02=c1.cp02 AND c2.cp03=c1.cp03 AND c2.cp04=c1.cp04" & _
         " AND c2.cp10 in('101','102','103') AND c2.cp09=c1.cp43 AND c2.CP27 is not null AND c2.CP20||c2.CP57||c2.CP60 IS NULL" & _
         " AND not exists(SELECT R01,R02,R03,R04,R12 FROM R060204 WHERE R01='" & strUserNum & "' and R02='F' and R03='" & IIf(stDept = "F22", "2", "1") & "' and R04=c1.CP09 and R12=0)"
      cnnConnection.Execute stSQL, intI
   End If
   '2022/3/15 END
   
   'Added by Lydia 2015/09/09 +早收文
   '通知工程師和程序
   If stDept = "F21" Or stDept = "F22" Then
       strExc(2) = CompWorkDay(4, strSrvDate(1)) '系統日+3個工作天
       strExc(3) = CompDate(2, 14, strSrvDate(1)) '系統日+14個日歷天
       strExc(4) = CompDate(2, -14, strSrvDate(1)) '系統日-14個日歷天
      '1.  若FMP之香港標準記錄請求(110)或澳門發明進度資料，若本所期限尚未達彈跳條件時，則再判斷若收文日<系統日－14個日曆天且本所期限<=系統日＋14個日曆天；
            'Modified by Lydia 2016/09/22 CP57||CP27 IS NULL => CP158=0 AND CP159=0
            strExc(5) = " SELECT '" & strUserNum & "','H' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
                    " From CASEPROGRESS,PATENT,FAGENT,CASEMAP" & _
                    " Where (CP01='P' OR CP01='CFP') AND CP158=0 AND CP159=0 AND SUBSTR(CP12,1,1)='F' and CP14 IN (" & stNumList & ")" & _
                    " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
                    " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
                    " AND CM01(+)=CP01 AND CM02(+)=CP02 AND CM03(+)=CP03 AND CM04(+)=CP04 AND CM10='4' AND CP10='110'" & _
                    " AND CP06 >= " & strExc(2) & " and cp05<" & strExc(4) & " and cp06 <= " & strExc(3)
            'Added by Lydia 2016/09/22
            stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & strExc(5)
            cnnConnection.Execute stSQL, intI
            'end 2016/09/22
            
            'Modified by Lydia 2016/09/22 拿掉Union
            strExc(5) = " SELECT '" & strUserNum & "','H' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
                    " From CASEPROGRESS,PATENT,FAGENT,CASEMAP" & _
                    " Where (CP01='P' OR CP01='CFP') AND CP158=0 AND CP159=0 and CP14 IN (" & stNumList & ")" & _
                    " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
                    " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
                    " AND CM01(+)=CP01 AND CM02(+)=CP02 AND CM03(+)=CP03 AND CM04(+)=CP04 AND CM10='5' AND CP10 in (" & CaseMapIn & ")" & _
                    " AND CP06 >= " & strExc(2) & " and cp05<" & strExc(4) & " and cp06 <= " & strExc(3)
            stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & strExc(5)
            cnnConnection.Execute stSQL, intI
      '2.若FMP或FCP之實審(416)、年費(605)進度資料，若本所期限尚未達彈跳條件時，則再判斷若收文日<系統日－14個日曆天且本所期限<=系統日＋14個日曆天；
             'Modified by Lydia 2016/09/22 CP57||CP27 IS NULL =>CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
             strExc(5) = " SELECT '" & strUserNum & "','H' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
                    " From CASEPROGRESS,PATENT,FAGENT" & _
                    " Where (CP01='P' OR CP01='CFP') AND CP158=0 AND CP159=0 AND SUBSTR(CP12,1,1)='F' and CP14 IN (" & stNumList & ")" & _
                    " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
                    " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
                    " AND CP10 in ('416','605') AND CP06 >= " & strExc(2) & " and cp05<" & strExc(4) & " and cp06 <= " & strExc(3)
             'Added by Lydia 2016/09/22
             stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & strExc(5)
             cnnConnection.Execute stSQL, intI
             'end 2016/09/22
             'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' => CP01='FCP' AND CP158=0 AND CP159=0
             'Modified by Lydia 2016/09/22 拿掉Union
             strExc(5) = " SELECT '" & strUserNum & "','H' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
                    " From CASEPROGRESS,PATENT,FAGENT" & _
                    " Where CP01='FCP' AND CP158=0 AND CP159=0 AND CP10 in ('416','605') and CP14 IN (" & stNumList & ") AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
                    " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
                    " AND CP06 >= " & strExc(2) & " and cp05<" & strExc(4) & " and cp06 <= " & strExc(3)
             stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & strExc(5)
             cnnConnection.Execute stSQL, intI
      '3.若FMP或FCP之分割(307)進度資料，若本所期限尚未達彈跳條件時，則再判斷若收文日<系統日－14個日曆天且本所期限<=系統日＋14個日曆；
             strExc(5) = " SELECT '" & strUserNum & "','H' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
                    " From CASEPROGRESS,PATENT,FAGENT" & _
                    " Where (CP01='P' OR CP01='CFP') AND CP57||CP27 IS NULL AND SUBSTR(CP12,1,1)='F' and CP14 IN (" & stNumList & ")" & _
                    " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
                    " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) " & _
                    " AND CP10 ='307' AND CP06 >= " & strExc(2) & " and cp05<" & strExc(4) & " and cp06 <= " & strExc(3)
             'Added by Lydia 2016/09/22
             stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & strExc(5)
             cnnConnection.Execute stSQL, intI
             'end 2016/09/22
             'Modified by Lydia 2016/09/22 strExc(5) & " Union SELECT => " SELECT
             strExc(5) = " SELECT '" & strUserNum & "','H' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
                    " From CASEPROGRESS,PATENT,FAGENT" & _
                    " Where CP01||CP27||CP57='FCP' and CP14 IN (" & stNumList & ") AND CP10 ='307' AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
                    " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
                    " AND CP06 >= " & strExc(2) & " and cp05<" & strExc(4) & " and cp06 <= " & strExc(3)
             stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & strExc(5)
             cnnConnection.Execute stSQL, intI
   End If
   'end 2015/09/09
   
   'Add By Sindy 2017/1/16 I.准未請款/J.分割建議/K.通知告准
   '通知告准未發文的核准函(非收文日)之次日起
   strExc(0) = "SELECT C2.cp01 C2_CP01,C2.cp02 C2_CP02,C2.cp03 C2_CP03,C2.cp04 C2_CP04,C2.cp09 C2_CP09,C2.CP66 C2_CP66,C2.CP67 C2_CP67,C2.CP13 C2_CP13,C2.CP14 C2_CP14,Na16,fa10,pa85,C2.cp27 C2_CP27,PA162" & _
               ",nvl(s1.st52,'') s1_ST52,nvl(s2.st52,'') s2_ST52,nvl(s3.st52,'') s3_ST52" & _
               " From CASEPROGRESS C1,CASEPROGRESS C2,PATENT,FAGENT,Nation,staff s1,staff s2,staff s3" & _
               " Where C1.CP01='FCP' and C1.CP10='1917' AND C1.CP158=0 AND C1.CP159=0" & _
               " AND C1.CP43=C2.CP09(+) AND C2.CP10='1001'" & _
               " AND PA01(+)=C2.CP01 AND PA02(+)=C2.CP02 AND PA03(+)=C2.CP03 AND PA04(+)=C2.CP04" & _
               " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
               " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
               " AND NA01(+)=FA10" & _
               " AND C2.cp13=s1.st01(+) AND C2.CP14=s2.st01(+) AND Na16=s3.st01(+)" & _
               " order by 1,2,3,4"
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If rsA.RecordCount > 0 Then
      rsA.MoveFirst
      Do While Not rsA.EOF
         'Modify By Sindy 2017/3/15 FCP-54734排除不向客戶收款(AND cp20 is null)
         'Modify By Sindy 2017/4/14 + 修正204或主動修正203
         'Modify By Sindy 2017/4/20 + 補充說明206擇一申復239
         strSql = "SELECT cp09,cp60,a1k01,a1k02,a1k19,a1k20,DST01 FROM caseprogress,acc1k0,DivsugText" & _
                  " WHERE cp01='" & rsA.Fields("C2_cp01") & "'" & _
                  " AND cp02='" & rsA.Fields("C2_cp02") & "'" & _
                  " AND cp03='" & rsA.Fields("C2_cp03") & "'" & _
                  " AND cp04='" & rsA.Fields("C2_cp04") & "'" & _
                  " AND substr(cp09,1,1)='A' AND cp10 in('205','107','204','203','206','239') AND cp20 is null" & _
                  " AND cp60=a1k01(+)" & _
                  " AND cp01=DST01(+)" & _
                  " AND cp02=DST02(+)" & _
                  " AND cp03=DST03(+)" & _
                  " AND cp04=DST04(+)" & _
                  " order by A1K02 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strCP09 = RsTemp.Fields("cp09")
            strCP60 = "" & RsTemp.Fields("cp60")
            strA1K01 = "" & RsTemp.Fields("a1k01")
            strA1K19 = "" & RsTemp.Fields("a1k19") '請款單輸入日期
            If strA1K19 <> "" Then strA1K19 = DBDATE(strA1K19)
            strA1K20 = "" & RsTemp.Fields("a1k20") '請款單輸入時間
            If strA1K20 <> "" Then strA1K20 = Format(strA1K20, "000000")
            strDST01 = "" & RsTemp.Fields("DST01")
            '未請款
            If strCP60 = "" Then
               'A.該案若有A類申復205或再審107或修正204或主動修正203或補充說明206或擇一申復239未請款(無CP60)時，增加事件為"I准未請款"之提醒
               '核准函輸入日期（非收文日）之次日起
               '提醒工程師、各區承辦、各區程序，及三人主管，直至A類申復或再審或修正或主動修正或補充說明或擇一申復請款
               If (stDept = "F21" Or stDept = "F22" Or stDept = "F23") And _
                   (InStr(stNumList, rsA.Fields("C2_CP13")) > 0 Or (InStr(stNumList, "" & rsA.Fields("s1_ST52")) > 0 And "" & rsA.Fields("s1_ST52") <> "") Or _
                   InStr(stNumList, rsA.Fields("C2_CP14")) > 0 Or (InStr(stNumList, "" & rsA.Fields("s2_ST52")) > 0 And "" & rsA.Fields("s2_ST52") <> "") Or _
                   ((InStr(stNumList, "" & rsA.Fields("Na16")) > 0 And "" & rsA.Fields("Na16") <> "") Or (InStr(stNumList, "" & rsA.Fields("s3_ST52")) > 0 And "" & rsA.Fields("s3_ST52") <> ""))) And _
                  rsA.Fields("C2_CP66") < strSrvDate(1) Then
                  
                  stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12) values(" & _
                  "'" & strUserNum & "','I','1'," & _
                  CNULL(rsA.Fields("C2_CP09")) & "," & _
                  CNULL(rsA.Fields("fa10")) & "," & _
                  CNULL(rsA.Fields("C2_CP14")) & "," & _
                  CNULL(rsA.Fields("C2_CP13")) & "," & _
                  "NULL,NULL,NULL,NULL,0)"
                  cnnConnection.Execute stSQL, intI
               End If
               '核准函未上發文日且案件之PA162有註記應另函通知初審核准後分割者，
               '若案件無「J分割建議」(以本所案號讀取分割建議定稿文字檔之DST05)
               '且非日文定稿(依定稿語文規則)時，增加事件為"分割建議"之提醒
               '核准函輸入日期（非收文日）之次日起
               '提醒工程師及其主管，直至A類申復或再審請款
               'Modifiedby Morgan 2022/10/11 取消日文定稿限制
               If "" & rsA.Fields("C2_CP27") = "" And rsA.Fields("PA162") = "Y" Then
                  If stDept = "F21" And _
                     (InStr(stNumList, rsA.Fields("C2_CP14")) > 0 Or (InStr(stNumList, "" & rsA.Fields("s2_ST52")) > 0 And "" & rsA.Fields("s2_ST52") <> "")) And _
                     rsA.Fields("C2_CP66") < strSrvDate(1) And strDST01 = "" Then
                     
                     stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12) values(" & _
                     "'" & strUserNum & "','J','1'," & _
                     CNULL(rsA.Fields("C2_CP09")) & "," & _
                     CNULL(rsA.Fields("fa10")) & "," & _
                     CNULL(rsA.Fields("C2_CP14")) & "," & _
                     CNULL(rsA.Fields("C2_CP13")) & "," & _
                     "NULL,NULL,NULL,NULL,0)"
                     cnnConnection.Execute stSQL, intI
                  End If
               End If
            '已請款
            Else
               'C.上述A點已請款案件，增加事件為"K通知告准"之提醒
               '提醒各區程序及其主管，直至該案"通知告准" D類進度上發文日
               If stDept = "F22" And _
                  ((InStr(stNumList, "" & rsA.Fields("Na16")) > 0 And "" & rsA.Fields("Na16") <> "") Or (InStr(stNumList, "" & rsA.Fields("s3_ST52")) > 0 And "" & rsA.Fields("s3_ST52") <> "")) Then
                  'Modify By Sindy 2017/2/18 +R13.備註
                  '106/2/18：敏莉提"優先請款"只顯示於請款輸入日期時間大於核准函輸入日期時間的資料
                  If strA1K19 & strA1K20 > rsA.Fields("C2_CP66") & Format(rsA.Fields("C2_CP67") & "00", "000000") Then
                     stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13) values(" & _
                     "'" & strUserNum & "','K','2'," & _
                     CNULL(rsA.Fields("C2_CP09")) & "," & _
                     CNULL(rsA.Fields("fa10")) & "," & _
                     CNULL(rsA.Fields("C2_CP14")) & "," & _
                     CNULL(rsA.Fields("C2_CP13")) & "," & _
                     "NULL,NULL,NULL,NULL,0,'優先請款;')"
                     cnnConnection.Execute stSQL, intI
                  End If
               End If
            End If
         End If
         rsA.MoveNext
      Loop
   End If
   rsA.Close
   '2017/1/16 END
   
   'Added by Lydia 2017/11/29 FCP案件命名電子化:外專工程師及業務承辦組增加"L-待命名期限"
   If strSrvDate(1) >= FCP案件命名啟用日 Then
        If bLvlChm = True Or bLvlJpn = True Or bLvlMot = True Or bLvlEls = True Then  '未分工程師組別
            If bLvlEls = True Then
               strExc(2) = "1"
            ElseIf bLvlChm = True Then
               strExc(2) = "2"
            ElseIf bLvlJpn = True Then
               strExc(2) = "3"
            Else
               strExc(2) = "4"
            End If
            'Add By Sindy 2021/4/22 待命名未分工程師組別,彈給五級主管看
            stSQL = ""
            If bLvl5 = True Then
            '2021/4/22 END
               stSQL = "SELECT '" & strUserNum & "','L' EV1,' ' EV2,CP09,FA10,TCT10 AS CP14,CP13,NVL(TCT02,WORKDAYADD(2,CP66)) AS CP06,CP07,CP48,NULL,0 " & _
                       "FROM TRANSCASETITLE, CASEPROGRESS, PATENT, FAGENT " & _
                       "WHERE NVL(TCT04,'N')='N' AND TCT01=CP09(+) " & _
                       "AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
                       "AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) "
            End If
            If stSQL <> "" Then stSQL = stSQL & " UNION ALL "
            stSQL = stSQL & "SELECT '" & strUserNum & "','L' EV1,' ' EV2,CP09,FA10,TCT10 AS CP14,CP13,NVL(TCT02,WORKDAYADD(2,CP66)) AS CP06,CP07,CP48,NULL,0 " & _
                    "FROM TRANSCASETITLE, CASEPROGRESS, PATENT, FAGENT " & _
                    "WHERE NVL(TCT10,'N')='N' AND TCT01=CP09(+) " & _
                    "AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
                    "AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND PA150=" & CNULL(strExc(2))
            stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12) " & stSQL
            cnnConnection.Execute stSQL, intI
        End If
        
        strExc(2) = CompWorkDay(3, strSrvDate(1), 1) '系統日-2工作天
        If stDept = "F21" Then  '外專工程師
            stSQL = "SELECT '" & strUserNum & "','L' EV1,' ' EV2,CP09,FA10,TCT10 AS CP14,CP13,NVL(TCT02,WORKDAYADD(2,CP66)) AS CP06,CP07,CP48,NULL,0 " & _
                    "FROM TRANSCASETITLE, CASEPROGRESS, PATENT, FAGENT " & _
                    "WHERE TCT10='" & stUserID & "' AND NVL(TCT05,0)=0 AND TCT01=CP09(+) " & _
                    "AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
                    "AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) "
            '2級只看自己及部屬資料
            stSQL = stSQL & "UNION ALL SELECT '" & strUserNum & "','L' EV1,' ' EV2,CP09,FA10,TCT10 AS CP14,CP13,NVL(TCT02,WORKDAYADD(2,CP66)) AS CP06,CP07,CP48,NULL,0 " & _
                              "FROM TRANSCASETITLE, CASEPROGRESS, PATENT, FAGENT,STAFF " & _
                              "WHERE NVL(TCT05,0)=0 AND TCT10=ST01 AND ST52='" & stUserID & "' AND TCT01=CP09(+) " & _
                              "AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
                              "AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) "
            If Trim(stNumList1(3)) <> "" Then
               '3級以上逾期2天資料
               stSQL = stSQL & "UNION ALL SELECT '" & strUserNum & "','L' EV1,' ' EV2,CP09,FA10,TCT10 AS CP14,CP13,NVL(TCT02,WORKDAYADD(2,CP66)) AS CP06,CP07,CP48,NULL,0 " & _
                                 "FROM TRANSCASETITLE, CASEPROGRESS, PATENT, FAGENT,STAFF " & _
                                 "WHERE NVL(TCT05,0)=0 AND TCT10=ST01 AND (ST53='" & stUserID & "' OR ST54='" & stUserID & "' OR ST55='" & stUserID & "')  AND TCT01=CP09(+) " & _
                                 "AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
                                 "AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) " & _
                                 "AND NVL(TCT02,WORKDAYADD(2,CP66))<=" & strExc(2)
            End If
            stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12) " & stSQL
            cnnConnection.Execute stSQL, intI
        ElseIf stDept = "F23" Then '業務承辦組
            '2級只看自己及部屬資料
            stSQL = "SELECT '" & strUserNum & "','L' EV1,' ' EV2,CP09,FA10,TCT10 AS CP14,CP13,NVL(TCT02,WORKDAYADD(2,CP66)) AS CP06,CP07,CP48,NULL,0 " & _
                    "FROM TRANSCASETITLE, CASEPROGRESS, PATENT, FAGENT, NATION " & _
                    "WHERE NVL(TCT05,0)=0 AND TCT01=CP09(+) " & _
                    "AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
                    "AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) " & _
                    "AND FA10=NA01(+) AND decode(pa75," & midStr & ",na51) IN (" & stNumList1(1) & IIf(Trim(stNumList1(2)) <> "", "," & stNumList1(2), "") & ") "
            If Trim(stNumList1(3)) <> "" Then
                '3級以上逾期2天資料
                stSQL = stSQL & "UNION SELECT '" & strUserNum & "','L' EV1,' ' EV2,CP09,FA10,TCT10 AS CP14,CP13,NVL(TCT02,WORKDAYADD(2,CP66)) AS CP06,CP07,CP48,NULL,0 " & _
                        "FROM TRANSCASETITLE, CASEPROGRESS, PATENT, FAGENT, NATION " & _
                        "WHERE NVL(TCT05,0)=0 AND TCT01=CP09(+) " & _
                        "AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
                        "AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) " & _
                        "AND FA10=NA01(+) AND decode(pa75," & midStr & ",na51) IN (" & stNumList & ") " & _
                        "AND NVL(TCT02,WORKDAYADD(2,CP66))<=" & strExc(2)
            End If
            stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12) " & stSQL
            cnnConnection.Execute stSQL, intI
        End If
   End If
   'end 2017/11/29
   
   'Add By Sindy 2021/4/22 未請款的本所期限為承辦期限(管控日期)再＋5個工作天
   stSQL = "UPDATE R060204 SET R08=WORKDAYADD(6,R10) WHERE R01='" & strUserNum & "' AND R02='F'"
   cnnConnection.Execute stSQL, intI
   '2021/4/22 END
   
   'Add By Sindy 2021/8/9 達核稿,因進度裡承辦人不是工程師,例如F5588.舜禹翻譯就不會判斷到達本所,增加其判斷
   'R06:承辦,R07:業務,R08:所限,R09:法限,R10:辦限,R11:核限
   stSQL = "UPDATE R060204 SET R02='A'" & _
            " WHERE R01||R04||R02 IN(" & _
            "SELECT R01||R04||R02 FROM R060204,caseprogress WHERE cp09(+)=R04" & _
            " AND r06<>cp14 AND r01='" & strUserNum & "' AND R02='C'" & _
            Replace(UCase(stConCP06), "CP06", "R08") & ")"
   cnnConnection.Execute stSQL, intI
   '2021/8/9 END
   
   'Add By Sindy 2015/10/22
   '程序組的期限彈跳排除:
   'modiby by sonia 此段原在下面2019/5/22搬上來, 排除C類來函(承辦人為工程師之審查意見,核駁等)逾法定期限之案件之前,但FCP-050839過期未發文否則下一程序也不會出現
   If stDept = "F22" Then
      '在前頭sql裡就過濾掉了
'      '排除926.核對已准專利,945.電話連絡單
'      stSQL = "DELETE R060204 WHERE R01='" & strUserNum & "'" & _
'         " AND EXISTS(SELECT * FROM caseprogress WHERE cp09=R04 AND cp10 in('926','945'))"
'      cnnConnection.Execute stSQL, intI
'      '排除R02='F' and cp27>0 已發文未請款
'      stSQL = "DELETE R060204 WHERE R01='" & strUserNum & "' AND R02='F'" & _
'         " AND EXISTS(SELECT * FROM caseprogress WHERE cp09=R04 AND cp27>0)"
'      cnnConnection.Execute stSQL, intI
      '排除C類來函(承辦人為工程師之審查意見,核駁等)逾法定期限之案件
      stSQL = "DELETE R060204 WHERE R01='" & strUserNum & "' AND substr(R04,1,1)='C'" & _
                                          " AND R06 is not null" & _
                                          " AND R09 is not null AND R09<" & strSrvDate(1) & _
         " AND EXISTS(SELECT * FROM staff WHERE R06=ST01 AND ST15='F21')"
      cnnConnection.Execute stSQL, intI
   End If
   '2015/10/22 END
      
   '刪除E.未發文
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02='E'" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02<>'E' AND R2.R02<>'D')"
   cnnConnection.Execute stSQL, intI
   
   '已達本所'A'時,則刪除達承辦'B'及達核稿'C'
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02 IN ('B','C')" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='A')"
   cnnConnection.Execute stSQL, intI
   
   'Add By Sindy 2023/9/19
   '已達指定'N'時,則刪除達核稿'C'
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02='C'" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='N')"
   cnnConnection.Execute stSQL, intI
   '2023/9/19 END
   
   '*****
   'Modify By Sindy 2017/1/13 事件為'B.達承辦'、'A.達本所'、'E.未發文'抓未發文資料的語法，都要剔除"通知告准" (1917)
   '*****
   'R04:總收文號
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02 IN ('A','B','E')" & _
      " AND EXISTS(SELECT * FROM R060204 R2,caseprogress WHERE R2.R01=R1.R01 AND R2.R02=R1.R02 AND R2.R04=R1.R04 AND R2.R04=cp09(+) and cp158=0 and cp10='1917')"
   cnnConnection.Execute stSQL, intI
   '2017/1/13 END
   
   'Added by Lydia 2019/05/31 程序大項工作整批發文: 排除整批發文的案件性質
   If stDept = "F22" Then
        stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' " & _
           " AND EXISTS(SELECT * FROM R060204 R2,caseprogress WHERE R2.R01=R1.R01 AND R2.R02=R1.R02 AND R2.R04=R1.R04 AND R2.R04=cp09(+) and cp158=0 and cp158=0 and cp10 in ('1603','1229','1604','1605') )"
        cnnConnection.Execute stSQL, intI
        'Added by Lydia 2021/08/26 刪除'A.達本所','N=達指定'同時為承辦人和管制人時，保留管制人資料; ex.FCP065559(AB0035398)、FCP065558(AB0035395)因為承辦人和管制人都屬於Phoebe的下屬，造成後面更新備註語法錯誤
        stSQL = "DELETE From R060204 R1 Where R1.R01='" & strUserNum & "' And R1.R02='A' And R03='1' And Exists(" & _
                     "Select * From R060204 R2 Where R2.R01=R1.R01 And R2.R04=R1.R04  And R1.R02='A' And R03='2' ) "
        cnnConnection.Execute stSQL, intI
        stSQL = "DELETE From R060204 R1 Where R1.R01='" & strUserNum & "' And R1.R02='N' And R03='1' And Exists(" & _
                     "Select * From R060204 R2 Where R2.R01=R1.R01 And R2.R04=R1.R04  And R1.R02='N' And R03='2' ) "
        cnnConnection.Execute stSQL, intI
        'end 2021/08/26
   End If
   
   'Added by Lydia 2020/11/19 因為達本所和客戶提供文件會有重複記錄的情況,所以刪除非客戶提供文件=M; ex.Sharon看到Gill的客戶文件是紅字
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02<>'M'" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='M')"
   cnnConnection.Execute stSQL, intI
   'end 2020/11/19
   
   'Add By Sindy 2017/9/11
   '若同案有 'A達本所'期限,其他的就不要再顯示
   'Modify By Sindy 2021/4/21 N.達指定要另外判斷
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02 not in('A','N')" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='A')"
   cnnConnection.Execute stSQL, intI
   '更新備註
   stSQL = "UPDATE R060204 R1 SET R13='達指定;'||(SELECT R13 FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='N')" & _
      " WHERE R1.R01='" & strUserNum & "' AND R1.R02='A'" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='N' AND R2.R14>0 AND R2.R08>0 AND R2.R08>=R2.R14)" & _
      " AND R1.R08>0 AND R1.R14>0"
   cnnConnection.Execute stSQL, intI
   stSQL = "UPDATE R060204 R1 SET R13='達本所;'||R13" & _
      " WHERE R1.R01='" & strUserNum & "' AND R1.R02='N'" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='A' AND R2.R14>0 AND R2.R08>0 AND R2.R08<R2.R14)" & _
      " AND R1.R08>0 AND R1.R14>0"
   cnnConnection.Execute stSQL, intI
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02='N'" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='A' AND R2.R14>0 AND R2.R08>0 AND R2.R08>=R2.R14)"
   cnnConnection.Execute stSQL, intI
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02='A'" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='N' AND R2.R14>0 AND R2.R08>0 AND R2.R08<R2.R14)"
   cnnConnection.Execute stSQL, intI
   '2017/9/11 END
   'Add By Sindy 2021/12/27 同一個本所案號B.達承辦N.達指定同時出現時,達指定優先
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02='B'" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='N')"
   cnnConnection.Execute stSQL, intI
   '2021/12/27 END
   
   'Add By Sindy 2024/1/4
   '刪除內翻人員的達承辦'B'時,若已有完稿日者
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02='B' AND substr(R06,1,1)='F'" & _
      " AND EXISTS(SELECT * FROM R060204 R2,caseprogress,ENGINEERPROGRESS" & _
                  " WHERE R2.R01=R1.R01 AND R2.R02=R1.R02 AND R2.R04=R1.R04" & _
                  " AND R2.R04=cp09(+) and cp10 not in ('" & 翻譯 & "') AND R2.R04=ep02(+) AND EP09 is not null)"
   cnnConnection.Execute stSQL, intI
   '2024/1/4 END
   
   'Added by Lydia 2021/08/23 工程師所有未發文查詢，全部顯示; 86019做「所有未發文」的查詢，其中FCP-65378的實審AB0032149同時存在早收文和未發文，而早收文負責查看人員=FCP程序管制在SetColor有判斷負責人本人或第二級才看
   bolShowAll = False
   If stDept = "F21" And idx = 1 Then
       bolShowAll = True
   End If
   'end 2021/08/23
   
   'Added by Lydia 2018/02/09 設sort排序
   '外專程序要求: "未請款"放在表單中倒數第1;客戶提供文件(待處理)放在倒數第2
   'Modify By Sindy 2023/8/31 +,'' as CodeUserd
   If stDept = "F22" Then
       stOrdCon = ",decode(R02,'I',0,'J',0,'K',0,'F',9,'M',5,1) sort,'' as CodeUserd "
   Else
       stOrdCon = ",decode(R02,'I',0,'J',0,'K',0,1) sort,'' as CodeUserd "
   End If
   'end 2018/02/08
   
   'Added by Lydia 2015/09/09 + H=早收文
   'Add By Sindy 2015/10... +,NVL(DL.Cnt,0) Cnt,CP43,R12
   'Add By Sindy 2015/11/26 抓延期次數
   'Modify By Sindy 2016/3/11 + FCP-46153 104/12/14尚未收申復
   '未延期:增加判斷若做CP43去檢查是否有延期時,還要過濾掉CP09不可為C類 ==> and cp09<'C'
   stVTB = "select DL01,sum(DLCnt) Cnt from" & _
           " (select DL01,count(*) DLCnt from R060204,datelimit" & _
           " where R01='" & strUserNum & "' and R04=DL01 group by DL01" & _
           " Union all " & _
           " select cp09 DL01,count(*) DLCnt from R060204,caseprogress,datelimit" & _
           " where R01='" & strUserNum & "' and R04=cp09(+) and cp43 is not null and cp43=DL01 and cp09<'C' group by cp09)" & _
           " group by DL01"
   '2015/11/26 END
   
   'Add By Sindy 2024/12/31
   '若身份為F21工程師人員，約定期限欄位改為【指定送件日】並帶出指定送件日期 之前/當天/之後
   '若身份為F22程序(非工程師)人員，核稿期限欄位改為【指定送件日】並帶出指定送件日期 之前/當天/之後
   If PUB_GetST03(txtUsernum) = "F21" Then
      '主檔
      strColR15_M = ",NVL(lpad(SQLDateT(cp142),9,' '),'2')||decode(cp164,'1','當天','2','之前','3','之後','') 指定送件日"
      '下一程序
      strColR15_N = ",'' 指定送件日"
   Else
      'F22,F23
      strColR15_M = ",NVL(lpad(SQLDateT(R15),9,' '),'2') 約定期限"
      strColR15_N = ",NVL(lpad(SQLDateT(R15),9,' '),'2') 約定期限"
   End If
   If PUB_GetST03(txtUsernum) = "F22" Then
      strColR11_M = ",NVL(lpad(SQLDateT(cp142),9,' '),'2')||decode(cp164,'1','當天','2','之前','3','之後','') 指定送件日"
      strColR11_N = ",'' 指定送件日"
   Else
      'F21,F23
      strColR11_M = ",NVL(lpad(SQLDateT(R11),9,' '),'2') 核稿期限"
      strColR11_N = ",NVL(lpad(SQLDateT(R11),9,' '),'2') 核稿期限"
   End If
   '2024/12/31 END
   
   'Modify By Sindy 2015/12/15 + CP142
   'Modify By Sindy 2017/1/18 + ,'I','准未請款','J','分割建議','K','通知告准'
   '                            ,decode(R02,'I',0,'J',0,'K',0,1) sort
   'Modified by Lydia 2017/11/29 + 'L','待命名期限'
   'Modified by Lydia 2018/02/08  + 'M','待處理' ; 改排序decode(R02,'I',0,'J',0,'K',0,1) sort=> stOrdCon
   'Add By Sindy 2018/5/31 FCP-047800資料重覆 + R03 ==> '' as R03
   'Modified by Lydia 2018/11/13 因為SetColor無法判斷到本人和2級主管,所以'' as R03還原抓R03
   'Modify By Sindy 2021/4/28 + ,'' 約定期限
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非臺灣000抓CPM04; DECODE(INSTR('020,013,044',PA09),0,CPM03,CPM04) => DECODE(PA09,'000',CPM03,CPM04)
   'Modified by Lydia 2022/11/03 改判斷非臺灣000改抓FMP案管制人
   'Modified by lydia 2024/02/29 +PA150
   'Modify By Sindy 2024/5/28 +,CP176
   'Modify By Sindy 2024/6/24 +達本所,若為C類來函未發文前頭加*表示
   strSql = "SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      IIf(PUB_GetST03(txtUsernum) = "F21", strColR15_M, ",'' 約定期限") & _
      ",NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限" & strColR11_M & ",decode(cp158,0,decode(cp118,null,'','Y'),'') 電,DECODE(PA09,'000',S1.ST02,NVL(S4.ST02,S1.ST02)) 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(SUBSTR(CP09,1,1)||CP27||R02,'C'||'A','*')||DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號" & _
      ",DECODE(PA09,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||CP64 案件備註,NVL(PA05,NVL(PA06,PA07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,DECODE(PA09,'000',NA16,NVL(NA79,NA16)) NA16 ,R06,R07,CP01,CP02,CP03,CP04,'' 未收款,CP10,CP27,CP09,decode(CP01,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,CP43,R12,CP142,CP164,CP176 " & stOrdCon & _
      ",PA150 FROM R060204,CASEPROGRESS,NATION,PATENT,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,(" & stVTB & ") DL, STAFF S4" & _
      " WHERE R01='" & strUserNum & "' AND R02<>'D' AND DL.DL01(+)=R04 AND CP09(+)=R04 AND NA01(+)=R05" & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND S4.ST01(+)=NA79"

   'Add By Sindy 2015/10/22 +法務進度
   'Modified by Lydia 2018/02/08  + 'M','待處理' ;  改排序decode(R02,'I',0,'J',0,'K',0,1) sort=> stOrdCon
   'Modified by Sindy 2019/10/4 因為SetColor無法判斷到本人和2級主管,所以'' as R03還原抓R03
   'Modify By Sindy 2021/4/28 + ,'' 約定期限
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非臺灣000抓CPM04; DECODE(INSTR('020,013,044',LC15),0,CPM03,CPM04) => DECODE(LC15,'000',CPM03,CPM04)
   'Modified by Lydia 2024/02/29 +'' as PA150
   'Modify By Sindy 2024/6/24 +達本所,若為C類來函未發文前頭加*表示
   strSql = strSql & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      IIf(PUB_GetST03(txtUsernum) = "F21", strColR15_M, ",'' 約定期限") & _
      ",NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限" & strColR11_M & ",decode(cp158,0,decode(cp118,null,'','Y'),'') 電,S1.ST02 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(SUBSTR(CP09,1,1)||CP27||R02,'C'||'A','*')||DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號" & _
      ",DECODE(LC15,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||CP64 案件備註,NVL(LC05,NVL(LC06,LC07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,NA16,R06,R07,CP01,CP02,CP03,CP04,'' 未收款,CP10,CP27,CP09,decode(CP01,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,CP43,R12,CP142,CP164,CP176 " & stOrdCon & _
      ",'' as PA150 FROM R060204,CASEPROGRESS,NATION,Lawcase,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,(" & stVTB & ") DL" & _
      " WHERE R01='" & strUserNum & "' AND R02<>'D' AND DL.DL01(+)=R04 AND CP09(+)=R04 AND NA01(+)=R05" & _
      " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=CP01 AND CPM02(+)=CP10"
   '2015/10/22 END
   
   'Modified by Lydia 2018/02/08  + 'M','待處理'  ;  改排序decode(R02,'I',0,'J',0,'K',0,1) sort=> stOrdCon
   'Modified by Sindy 2019/10/4 因為SetColor無法判斷到本人和2級主管,所以'' as R03還原抓R03
   'Modify By Sindy 2021/4/28 + ,'' 約定期限
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非臺灣000抓CPM04; DECODE(INSTR('020,013,044',SP09),0,CPM03,CPM04) => DECODE(SP09,'000',CPM03,CPM04)
   'Modified by Lydia 2022/11/03 改判斷非臺灣000改抓FMP案管制人
   'Modified by Lydia 2024/02/29 +SP79 as PA150
   'Modify By Sindy 2024/6/24 +達本所,若為C類來函未發文前頭加*表示
   strSql = strSql & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      IIf(PUB_GetST03(txtUsernum) = "F21", strColR15_M, ",'' 約定期限") & _
      ",NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限" & strColR11_M & ",decode(cp158,0,decode(cp118,null,'','Y'),'') 電,DECODE(SP09,'000',S1.ST02,NVL(S4.ST02,S1.ST02)) 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(SUBSTR(CP09,1,1)||CP27||R02,'C'||'A','*')||DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號" & _
      ",DECODE(SP09,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||CP64 案件備註,NVL(SP05,NVL(SP06,SP07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,DECODE(SP09,'000',NA16,NVL(NA79,NA16)) NA16,R06,R07,CP01,CP02,CP03,CP04,'' 未收款,CP10,CP27,CP09,decode(CP01,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,CP43,R12,CP142,CP164,CP176 " & stOrdCon & _
      ",SP79 as PA150 FROM R060204,CASEPROGRESS,NATION,SERVICEPRACTICE,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,(" & stVTB & ") DL,STAFF S4" & _
      " WHERE R01='" & strUserNum & "' AND R02<>'D' AND DL.DL01(+)=R04 AND CP09(+)=R04 AND NA01(+)=R05" & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND S4.ST01(+)=NA79"
   'Modified by Lydia 2018/02/08  + 'M','待處理' ;  改排序decode(R02,'I',0,'J',0,'K',0,1) sort=> stOrdCon
   'modify by sonia 2019/5/22 未收文期限若為C類來函且C類未發文則於事件'未收文'前加*
   'Modified by Sindy 2019/10/4 因為SetColor無法判斷到本人和2級主管,所以'' as R03還原抓R03
   'Modify By Sindy 2021/4/28 + ,NVL(lpad(SQLDateT(R15),10,' '),'2') 約定期限
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非臺灣000抓CPM04; DECODE(INSTR('020,013,044',PA09),0,CPM03,CPM04) => DECODE(PA09,'000',CPM03,CPM04)
   'Modified by Lydia 2022/11/03 改判斷非臺灣000改抓FMP案管制人
   'Modified by Lydia 2024/02/29 +PA150
   strSql = strSql & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      strColR15_N & _
      ",NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限" & strColR11_N & ",' ' 電,DECODE(PA09,'000',S1.ST02,NVL(S4.ST02,S1.ST02)) 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(SUBSTR(NP01,1,1)||CP27,'C','*')||DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號" & _
      ",DECODE(PA09,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||NP15 案件備註,NVL(PA05,NVL(PA06,PA07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,DECODE(PA09,'000',NA16,NVL(NA79,NA16)) NA16,R06,R07,NP02,NP03,NP04,NP05,'' 未收款,NP07,0+'' CP27,np01 CP09,decode(NP02,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,'' CP43,R12,0 CP142,CP164,CP176 " & stOrdCon & _
      ",PA150 FROM R060204,NEXTPROGRESS,NATION,PATENT,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CASEPROGRESS,(" & stVTB & ") DL, STAFF S4" & _
      " WHERE R01='" & strUserNum & "' AND R02='D' AND DL.DL01(+)=R04 AND NP01(+)=R04 AND NP22(+)=R12 AND NA01(+)=R05" & _
      " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=NP02 AND CPM02(+)=NP07 AND NP01=CP09(+) AND S4.ST01(+)=NA79"
   'Add By Sindy 2015/10/22 +法務下一程序
   'Modified by Lydia 2018/02/08  + 'M','待處理' ;  改排序decode(R02,'I',0,'J',0,'K',0,1) sort=> stOrdCon
   'modify by sonia 2019/5/22 未收文期限若為C類來函且C類未發文則於事件'未收文'前加*
   'Modified by Sindy 2019/10/4 因為SetColor無法判斷到本人和2級主管,所以'' as R03還原抓R03
   'Modify By Sindy 2021/4/28 + ,NVL(lpad(SQLDateT(R15),10,' '),'2') 約定期限
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非臺灣000抓CPM04; DECODE(INSTR('020,013,044',LC15),0,CPM03,CPM04) => DECODE(LC15,'000',CPM03,CPM04)
   'Modified by Lydia 2024/02/29 +'' as PA150
   strSql = strSql & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      strColR15_N & _
      ",NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限" & strColR11_N & ",' ' 電,S1.ST02 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(SUBSTR(NP01,1,1)||CP27,'C','*')||DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號" & _
      ",DECODE(LC15,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||NP15 案件備註,NVL(LC05,NVL(LC06,LC07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,NA16,R06,R07,NP02,NP03,NP04,NP05,'' 未收款,NP07,0+'' CP27,np01 CP09,decode(NP02,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,'' CP43,R12,0 CP142,CP164,CP176 " & stOrdCon & _
      ",'' as PA150 FROM R060204,NEXTPROGRESS,NATION,Lawcase,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CASEPROGRESS,(" & stVTB & ") DL" & _
      " WHERE R01='" & strUserNum & "' AND R02='D' AND DL.DL01(+)=R04 AND NP01(+)=R04 AND NP22(+)=R12 AND NA01(+)=R05" & _
      " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05 AND LC01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=NP02 AND CPM02(+)=NP07 AND NP01=CP09(+)"
   '2015/10/22 END
   
   'Modified by Lydia 2018/02/08  + 'M','待處理' ;  改排序decode(R02,'I',0,'J',0,'K',0,1) sort=> stOrdCon
   'modify by sonia 2019/5/22 未收文期限若為C類來函且C類未發文則於事件'未收文'前加*
   'Modified by Sindy 2019/10/4 因為SetColor無法判斷到本人和2級主管,所以'' as R03還原抓R03
   'Modify By Sindy 2021/4/28 + ,NVL(lpad(SQLDateT(R15),10,' '),'2') 約定期限
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非臺灣000抓CPM04; DECODE(INSTR('020,013,044',SP09),0,CPM03,CPM04) => DECODE(SP09,'000',CPM03,CPM04)
   'Modified by Lydia 2022/11/03 改判斷非臺灣000改抓FMP案管制人
   'Modified by Lydia 2024/02/29 +SP79 as PA150
   strSql = strSql & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      strColR15_N & _
      ",NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限" & strColR11_N & ",' ' 電,DECODE(SP09,'000',S1.ST02,NVL(S4.ST02,S1.ST02)) 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(SUBSTR(NP01,1,1)||CP27,'C','*')||DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號" & _
      ",DECODE(SP09,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||NP15 案件備註,NVL(SP05,NVL(SP06,SP07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,DECODE(SP09,'000',NA16,NVL(NA79,NA16)) NA16,R06,R07,NP02,NP03,NP04,NP05,'' 未收款,NP07,0+'' CP27,np01 CP09,decode(NP02,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,'' CP43,R12,0 CP142,CP164,CP176 " & stOrdCon & _
      ",SP79 as PA150 FROM R060204,NEXTPROGRESS,NATION,SERVICEPRACTICE,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CASEPROGRESS,(" & stVTB & ") DL,STAFF S4" & _
      " WHERE R01='" & strUserNum & "' AND R02='D' AND DL.DL01(+)=R04 AND NP01(+)=R04 AND NP22(+)=R12 AND NA01(+)=R05" & _
      " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=NP02 AND CPM02(+)=NP07 AND NP01=CP09(+) AND S4.ST01(+)=NA79"
   'Add by Amy 2013/07/03 增加案件命名追蹤
   'Add by Amy 2013/07/02 '查詢人部門為F23或bLvl5(最高主管68009)抓命名追蹤自輸入日期2個工作日或期限前2個工作天
   If GetST15(txtUsernum) = "F23" Or bLvl5 = True Then
      Dim Manage As String
      stDate2 = PUB_GetWorkDayAfterSysDate(CDbl(strSrvDate(1)), 2) + 19110000
      stDate3 = PUB_GetWorkDayAfterSysDate(CDbl(strSrvDate(1)), -2) + 19110000
      '抓個人建的或是管制人自輸入日期2個工作日或期限前2個工作天資料
      'Modify By Sindy 2017/2/16 + ,1 sort
      'Modify By Sindy 2021/4/28 + ,'' 約定期限
      'Modify By Sindy 2023/8/31 +,'' as CodeUserd
      'Modified by Lydia 2024/02/29 +'' as PA150
      strSql = strSql & _
                        " Union Select '' V,LPAD(SQLDateT(TCN02),9,' ') 本所期限,LPAD(SQLDateT(TCN02),9,' ')  法定期限,'' 約定期限,'' 承辦期限, '' 核稿期限,' ' 電,S1.ST02 管制人, '' 承辦人,S1.ST02 智權人員, '新案' 事件,''||TCN01 本所案號, '新案' 案件性質,TCN04 案件備註, '' 案件名稱, '' 代理人國籍, 'H', '1', '', '', '', '', '', '', '','' 未收款, NULL, 0+'' CP27, '' CP09, 0 Srt1,0 Cnt,'' CP43,0 R12,0 CP142,'' CP164,'' CP176,1 sort,'' as CodeUserd " & _
                        ",'' as PA150 From TrackingCaseName,Staff S1 Where  S1.ST01(+)=TCN03  And TCN05 is null And (TCN03='" & txtUsernum & "' OR TCN06='" & txtUsernum & "' ) " & _
                        "And ( TCN02 <=" & stDate2 & " OR TCN07 <= " & stDate3 & ")"
      Manage = CheckManage
      If Len(Manage) > 5 Then
         '登入者為2級且為3級以上主管
         '2級只看自己及部屬資料
           'Modify By Sindy 2017/2/16 + ,1 sort
           'Modify By Sindy 2023/8/31 +,'' as CodeUserd
           'Modified by Lydia 2024/02/29 +'' as PA150
           strSql = strSql & _
                           " Union Select '' V,LPAD(SQLDateT(TCN02),9,' ') 本所期限,LPAD(SQLDateT(TCN02),9,' ')  法定期限,'' 約定期限,'' 承辦期限, '' 核稿期限,' ' 電,S1.ST02 管制人, '' 承辦人,S1.ST02 智權人員, '新案' 事件,''||TCN01 本所案號, '新案' 案件性質,TCN04 案件備註, '' 案件名稱, '' 代理人國籍, 'H', '1', '', '', '', '', '', '', '','' 未收款, NULL, 0+'' CP27, '' CP09, 0 Srt1,0 Cnt,'' CP43,0 R12,0 CP142,'' CP164,'' CP176,1 sort,'' as CodeUserd " & _
                           ",'' as PA150 From TrackingCaseName,STAFF S1 Where S1.ST01(+)=TCN03  And TCN05 is null And ST52='" & txtUsernum & "' And ( TCN02 <=" & stDate2 & " OR TCN07 <= " & stDate3 & ")"
           '3級以上逾期資料
           'Modified by Lydia 2024/02/29 +'' as PA150
           strSql = strSql & _
                           " Union Select '' V,LPAD(SQLDateT(TCN02),9,' ') 本所期限,LPAD(SQLDateT(TCN02),9,' ')  法定期限,'' 約定期限,'' 承辦期限, '' 核稿期限,' ' 電,S1.ST02 管制人, '' 承辦人,S1.ST02 智權人員, '新案' 事件,''||TCN01 本所案號, '新案' 案件性質,TCN04 案件備註, '' 案件名稱, '' 代理人國籍, 'H', '1', '', '', '', '', '', '', '','' 未收款, NULL, 0+'' CP27, '' CP09, 0 Srt1,0 Cnt,'' CP43,0 R12,0 CP142,'' CP164,'' CP176,1 sort,'' as CodeUserd " & _
                           ",'' as PA150 From TrackingCaseName,STAFF S1 Where S1.ST01(+)=TCN03  And TCN05 is null And (ST53='" & txtUsernum & "' OR ST54='" & txtUsernum & "' OR ST55='" & txtUsernum & "' ) And TCN02 <=" & strSrvDate(1)
   
      ElseIf Manage = "ST52" Then
          '只是2級主管
          'Modify By Sindy 2017/2/16 + ,1 sort
          'Modify By Sindy 2023/8/31 +,'' as CodeUserd
          'Modified by Lydia 2024/02/29 +'' as PA150
          strSql = strSql & _
                            " Union Select '' V,LPAD(SQLDateT(TCN02),9,' ') 本所期限,LPAD(SQLDateT(TCN02),9,' ')  法定期限,'' 約定期限,'' 承辦期限, '' 核稿期限,' ' 電,S1.ST02 管制人, '' 承辦人,S1.ST02 智權人員, '新案' 事件,''||TCN01 本所案號, '新案' 案件性質,TCN04 案件備註, '' 案件名稱, '' 代理人國籍, 'H', '1', '', '', '', '', '', '', '','' 未收款, NULL, 0+'' CP27, '' CP09, 0 Srt1,0 Cnt,'' CP43,0 R12,0 CP142,'' CP164,'' CP176,1 sort,'' as CodeUserd " & _
                            ",'' as PA150 From TrackingCaseName,STAFF S1 Where S1.ST01(+)=TCN03  And TCN05 is null And ST52='" & txtUsernum & "'  And ( TCN02 <=" & stDate2 & " OR TCN07 <= " & stDate3 & ")"
    
      ElseIf Manage = "ST5X" Then
           '只是3級以上主管
           'Modify By Sindy 2017/2/16 + ,1 sort
           'Modify By Sindy 2023/8/31 +,'' as CodeUserd
           'Modified by Lydia 2024/02/29 +'' as PA150
           strSql = strSql & _
                            " Union Select '' V,LPAD(SQLDateT(TCN02),9,' ') 本所期限,LPAD(SQLDateT(TCN02),9,' ')  法定期限,'' 約定期限,'' 承辦期限, '' 核稿期限,' ' 電,S1.ST02 管制人, '' 承辦人,S1.ST02 智權人員, '新案' 事件,''||TCN01 本所案號, '新案' 案件性質,TCN04 案件備註, '' 案件名稱, '' 代理人國籍, 'H', '1', '', '', '', '', '', '', '','' 未收款, NULL, 0+'' CP27, '' CP09, 0 Srt1,0 Cnt,'' CP43,0 R12,0 CP142,'' CP164,'' CP176,1 sort,'' as CodeUserd " & _
                            ",'' as PA150 From TrackingCaseName,STAFF S1 Where S1.ST01(+)=TCN03  And TCN05 is null And (ST53='" & txtUsernum & "' OR ST54='" & txtUsernum & "' OR ST55='" & txtUsernum & "' ) And TCN02 <=" & strSrvDate(1)
      End If
   End If
   '2013/07/02 END
   
   'Added by Lydia 2019/05/31 指定排序
   If stDept = "F22" Then
       strSql = strSql & " order by sort asc, Srt1 asc, 本所期限 asc,管制人 asc,代理人國籍 asc,本所案號 asc"
   ElseIf stDept = "F23" Then
       strSql = strSql & " order by sort asc, 本所期限 asc,智權人員 asc,代理人國籍 asc,本所案號 asc"
   Else
       strSql = strSql & " order by sort asc, 本所期限 asc,承辦人 asc,本所案號 asc"
   End If
   'end 2019/05/31
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If RsTemp Is Nothing Then Exit Sub
   If RsTemp.RecordCount = 0 Then
      Set m_adoRst = RsTemp.Clone
      SetRst2Grid
      MsgBox "查無資料！", vbInformation
      cmdHide.Enabled = False
      LblTotCnt.Caption = "共 0 筆" 'Add By Sindy 2009/10/07
   Else
      'Modify by Amy 2014/06/05 +FormName
      Set m_adoRst = PUB_CreateRecordset(RsTemp, , , 300, Me.Name)
      '更新未收文未收款&未收款本所期限
      SetXRecord
      'Remove by Lydia 2019/05/31 點選本所案號排序會出Sort錯誤
'      Select Case stDept
'         'Modify By Sindy 2017/1/19 + sort asc
'         Case "F22" '程序
'            'Modified by Morgan 2012/5/28 +系統別FMP的排前面
'            m_stSort = "sort asc, Srt1 asc, 本所期限 asc,管制人 asc,代理人國籍 asc,本所案號 asc"
'         Case "F23" '業務
'            m_stSort = "sort asc, 本所期限 asc,智權人員 asc,代理人國籍 asc,本所案號 asc"
'         'F21,F81
'         Case Else
'            m_stSort = "sort asc, 本所期限 asc,承辦人 asc,本所案號 asc"
'      End Select
'      m_adoRst.Sort = m_stSort
      'end 2019/05/31
      SetRst2Grid
      SetGrid
      RecordShow
      
      SetColor
      cmdHide.Enabled = True
      m_blnColOrderAsc = True
   End If
   
   Set rsA = Nothing
End Sub

Private Sub SetGrid()
   With grdDataList
      .Visible = False
      .FontFixed.Size = 8
      .Font.Size = 9
      '                0 1         2         3         4         5         6   7       8       9         10      11          12         13           14
      .FormatString = "V|本所期限 |法定期限 |約定期限 |承辦期限 |核稿期限 |電 |管制人 |承辦人 |智權人員 |事件　 |本所案號　　|案件性質 |備註　　　　|案件名稱　　　"
      For intI = 0 To .Cols - 1
         .ColAlignment(intI) = 0
         'Modify By Sindy 2024/12/31【電】此欄位刪除 + Or intI = 6
         If (intI > 14 And intI < 25) Or intI > 25 Or intI = 6 Then
            .ColWidth(intI) = 0
         End If
      Next
      'Add By Sindy 2024/12/31
      '若身份為工程師人員，約定期限欄位改為【指定送件日】並帶出指定送件日期 之前/當天/之後
      '若身份為程序(非工程師)人員，核稿期限欄位改為【指定送件日】並帶出指定送件日期 之前/當天/之後
      .ColWidth(9) = 580
      If PUB_GetST03(txtUsernum) = "F21" Then
         .TextMatrix(0, 3) = "指定送件日"
         .ColWidth(3) = 900
      ElseIf PUB_GetST03(txtUsernum) = "F22" Then
         .TextMatrix(0, 5) = "指定送件日"
         .ColWidth(5) = 900
      End If
      '2024/12/31 END
      'Added by Lydia 2024/02/29
      If colPA150 = 0 Then
         colPA150 = PUB_MGridGetId("PA150", grdDataList)
         colCaseNo = PUB_MGridGetId("本所案號", grdDataList)
      End If
      .ColWidth(colCaseNo) = 1100
      'end 2024/02/29
      .ColWidth(25) = 700
      .ColAlignment(25) = flexAlignRightTop
      .ColAlignment(1) = flexAlignRightTop
      .ColAlignment(2) = flexAlignRightTop
      .ColAlignment(3) = flexAlignLeftTop
      .ColAlignment(4) = flexAlignRightTop
      .ColAlignment(5) = flexAlignLeftTop
      .Visible = True
   End With
End Sub

Private Sub SetColor(Optional sHide As String = "N")
   Dim lngToday As Long, lngCP06 As Long, lngCP48 As Long, lngEP08 As Long, stType As String
   Dim ii As Integer, jj As Integer, dblCnt As Double
   Dim lngCP142 As Long, lngCP07 As Long, lngNP23 As Long
   Dim stTypeNote As String 'Add By Sindy 2021/7/8
   Dim strCP164 As String 'Add By Sindy 2021/11/5
   
   dblCnt = 0 'Add By Sindy 2009/10/07
   With grdDataList
   If .Rows > 1 Then
      .Visible = False
      ChgEmptyDate False
      lngToday = Val(strSrvDate(1))
      For ii = 1 To .Rows - 1
         .RowHeight(ii) = 255
         lngCP06 = Val(DBDATE(Trim(Replace(.TextMatrix(ii, 1), "/", ""))))
         lngCP07 = Val(DBDATE(Trim(Replace(.TextMatrix(ii, 2), "/", ""))))
         lngNP23 = Val(DBDATE(Trim(Replace(.TextMatrix(ii, 3), "/", "")))) 'Add By Sindy 2021/4/28
         lngCP48 = Val(DBDATE(Trim(Replace(.TextMatrix(ii, 4), "/", ""))))
         lngEP08 = Val(DBDATE(Trim(Replace(.TextMatrix(ii, 5), "/", ""))))
         lngCP142 = Val(DBDATE(Trim(Replace(.TextMatrix(ii, 33), "/", "")))) 'Add By Sindy 2021/4/21
         strCP164 = Trim(.TextMatrix(ii, 34)) 'Add By Sindy 2021/11/5
         
         stType = .TextMatrix(ii, 16)
         stTypeNote = .TextMatrix(ii, 13) 'Add By Sindy 2021/7/8
         .row = ii
         '固定欄位變回白色
         For jj = 0 To 2
            .col = jj
            .CellBackColor = &HFFFFFF
            .CellAlignment = flexAlignRightTop
            .CellFontSize = 9
         Next
         
         'Added by Lydia 2021/11/01 C類來函若有指定期限用淺藍色; 核駁及審查意見通知函備註設定新增：客戶C類來函承辦天數，
         '除更新承辦期限= (官方發文日+客戶C類來函承辦天數)，一併更新指定送件日期（方式=之前CP164）。
         If lngCP142 > 0 And Left(.TextMatrix(ii, 28), 1) = "C" Then
            For jj = 1 To .Cols - 1
               .col = jj
               '淺藍色
               .CellBackColor = &HFFFFC0
            Next
         'end 2021/11/01
         '逾管控期限
         'Modify By Amy 20130703 +stType="H" 命名追蹤 (Modify By Sindy 2021/4/22 Or stType = "H" del)
         'If ((stType = "A" Or stType = "D" Or stType = "F") And lngCP06 > 0 And lngCP06 < lngToday) Or
         'Modified by Lydia 2017/11/29 + stType = "L" 待命名
         'Modify By Sindy 2021/4/21 顏色總調整
         'A.達本所,N達指定:當天或逾期顯示紅色
         'Modified by Lydia 2021/11/01 更改為ElseIf
         'Modify By Sindy 2021/11/5 + And strCP164 <> "3": 以正常的規則走，即達指定送件期限，事件：達指定，彈跳顏色為紅色，請排除●之後
         ElseIf ((stType = "A" Or InStr(stTypeNote, "達本所") > 0) And (lngCP06 > 0 And lngCP06 <= lngToday)) Or _
            ((stType = "N" Or InStr(stTypeNote, "達指定") > 0) And (lngCP142 > 0 And lngCP142 <= lngToday) And strCP164 <> "3") Then
            '.TextMatrix(ii, 11) = "*" & .TextMatrix(ii, 11)
            For jj = 1 To .Cols - 1
               .col = jj
               '橘色：達本所(無法限)當天或逾期
               If stType = "A" And lngCP07 = 0 Then
                  '橘色
                  .CellBackColor = &H80FF&
               Else
                  '紅
                  .CellBackColor = &HFF&
               End If
            Next
         'A.達本所,N達指定:非當天並且非逾期顯示黃色
         'Modify By Sindy 2021/11/5 + And strCP164 <> "3": 以正常的規則走，即達指定送件期限，事件：達指定，彈跳顏色為紅色，請排除●之後
         ElseIf (stType = "A" And lngCP06 > 0 And lngCP06 > lngToday) Or _
            (stType = "N" And lngCP142 > 0 And lngCP142 > lngToday And strCP164 <> "3") Then
            '.TextMatrix(ii, 11) = "*" & .TextMatrix(ii, 11)
            For jj = 1 To .Cols - 1
               .col = jj
               '黃
               .CellBackColor = &HFFFF&
            Next
         'Add By Sindy 2021/4/21 P.達約定顯示黃色
         ElseIf stType = "P" Then
            For jj = 1 To .Cols - 1
               .col = jj
               '黃
               .CellBackColor = &HFFFF&
            Next
         '2021/4/21 END
         'Add By Sindy 2021/4/21 B.達承辦,C.達核稿顯示綠色
         ElseIf stType = "B" Or stType = "C" Then
            For jj = 1 To .Cols - 1
               .col = jj
               '綠
               .CellBackColor = &HC000&
            Next
         '2021/4/21 END
         '其它(D.未收文,F.未請款,L.待命名期限,G.未交稿)逾期顯示紫色
         'Modify By Sindy 2021/11/15 淑華:未收文當天改為紫色
         ElseIf ((stType = "F" Or stType = "L") And lngCP06 > 0 And lngCP06 < lngToday) Or _
            ((stType = "D") And lngNP23 > 0 And (lngNP23 = lngToday Or lngNP23 < lngToday)) Or _
            ((stType = "G") And lngCP48 > 0 And lngCP48 < lngToday) Then
            '.TextMatrix(ii, 11) = "*" & .TextMatrix(ii, 11)
            For jj = 1 To .Cols - 1
               .col = jj
               '紫色
               .CellBackColor = &HE600E6
            Next
         '當日期限
         'ElseIf ((stType = "A" Or stType = "D" Or stType = "F") And lngCP06 = lngToday) Or
         'Modified by Lydia 2017/12/13 +  stType = "L" 待命名
         '其它(L.待命名期限,G.未交稿)當天顯示綠色
         'F.未請款到承辦期限當天及逾期顯示綠色,直到逾本所期限跳紫色
         ElseIf ((stType = "L") And lngCP06 > 0 And lngCP06 = lngToday) Or _
            ((stType = "G") And lngCP48 > 0 And lngCP48 = lngToday) Or _
            ((stType = "F") And lngCP48 > 0 And lngCP48 <= lngToday) Then
            '.TextMatrix(ii, 11) = "v" & .TextMatrix(ii, 11)
            For jj = 1 To .Cols - 1
               .col = jj
               '綠
               .CellBackColor = &HC000&
            Next
         'Add By Sindy 2017/1/13 淡橘色：表示I.准未請款,J.分割建議,K.通知告准顯示橘色
         ElseIf stType = "I" Or stType = "J" Or stType = "K" Then
            For jj = 1 To .Cols - 1
               .col = jj
               '淡橘色
               .CellBackColor = &H80C0FF
            Next
         '2017/1/13 END
         'Added by Lydia 2018/02/08 M.待處理顯示粉紅色
         ElseIf stType = "M" Then
            For jj = 1 To .Cols - 1
               .col = jj
               '粉紅色
               .CellBackColor = &HC0C0FF
            Next
         'end 2018/02/08
         '未分案
         ElseIf .TextMatrix(ii, 17) = "0" Then
            '第五級不看
            If bLvl5 = True Then
               .RowHeight(ii) = 0
            Else
               For jj = 1 To .Cols - 1
                  .col = jj
                  '黃
                  .CellBackColor = &HFFFF&
               Next
            End If
         ElseIf sHide <> "N" Then
            .RowHeight(ii) = 0
         Else
            strExc(1) = .TextMatrix(ii, 17)
            Select Case strExc(1)
               '承辦人,核稿人
               Case "1", "4"
                  strExc(2) = .TextMatrix(ii, 19)
               Case "2" '管制人
                  strExc(2) = .TextMatrix(ii, 18)
               Case "3" '智權人員
                  strExc(2) = .TextMatrix(ii, 20)
               Case Else
                  strExc(2) = ""
            End Select
            'Added by Lydia 2021/08/23 工程師所有未發文查詢，全部顯示
            If bolShowAll = True Then
                strExc(2) = ""
            End If
            'end 2021/08/23
            
            If strExc(2) <> "" Then
               '本人或第二級才看
               If InStr(stNumList1(1) & "," & stNumList1(2), strExc(2)) = 0 Then
                  .RowHeight(ii) = 0
               End If
            End If
         End If
         'Add By Sindy 2009/10/07
         If .RowHeight(ii) > 0 Then
            dblCnt = dblCnt + 1
            .TextMatrix(ii, 36) = "顯示列" 'Add By Sindy 2023/8/31
         Else
            .TextMatrix(ii, 36) = "" 'Add By Sindy 2023/8/31
         End If
         '2009/10/07 End
         'Added by Lydia 2024/02/29 外專機械設計組人員異動調整程式：機械案在本所案號前增加▲符號
         'Modified by Lydia 2024/03/06 非機械設計組人員才顯示符號
         If "" & .TextMatrix(ii, colPA150) = "4" And m_stST16 <> "4" Then
            .TextMatrix(ii, colCaseNo) = "▲" & .TextMatrix(ii, colCaseNo)
         End If
         'end 2024/02/29
      Next
      .TopRow = 1
      .Visible = True
   End If
   End With
   If sHide = "N" Then
      cmdHide.Tag = "Y"
      cmdHide.Caption = "隱藏白色(&H)"
   Else
      cmdHide.Tag = "N"
      cmdHide.Caption = "顯示白色(&S)"
   End If
   LblTotCnt.Caption = "共 " & dblCnt & " 筆" 'Add By Sindy 2009/10/07
End Sub
Private Sub Form_Load()
   MoveFormToCenter Me
   Combo1.Clear
'   Combo1.AddItem "紅色(*)：表示逾管控期限"
'   Combo1.AddItem "綠色(v)：表示當日期限"
'   Combo1.AddItem "黃色(#)：表示未分案"
'   Combo1.AddItem "橘色：表示告准" 'Add By Sindy 2017/1/13
'   Combo1.AddItem "粉紅色：表示待處理" 'Added by Lydia 2018/02/08
   Combo1.AddItem "紅色：達本所(有法限),達指定當天或逾期"
   Combo1.AddItem "橘色：達本所(無法限)當天或逾期"
   Combo1.AddItem "黃色：達本所,達指定前2天"
   Combo1.AddItem "黃色：未分案"
   'Modify By Sindy 2021/11/15 淑華:未收文當天改為紫色
'   Combo1.AddItem "紫色：未收文,未請款,待命名,未交稿逾期"
'   Combo1.AddItem "綠色：未收文,未請款,待命名,未交稿當天"
   Combo1.AddItem "紫色：未收文(含當天),未請款,待命名,未交稿逾期"
   Combo1.AddItem "綠色：未請款,待命名,未交稿當天"
   '2021/11/15 END
   Combo1.AddItem "綠色：達承辦,達核稿"
   Combo1.AddItem "淺藍色：客戶C類來函指定送件日" 'Added by Lydia 2021/11/01
   Combo1.AddItem "淡橘色：准未請款,分割建議,通知告准"
   Combo1.AddItem "粉紅色：待處理"
   Combo1.ListIndex = 0
   txtUsernum = strUserNum
   'Modified by Morgan 2015/10/5
   '改外專程序可看該組其他人員資料
   Select Case Pub_StrUserSt03
   Case "M51", "F22"
      txtUsernum.Enabled = True
   End Select
   'end 2015/10/5
   'Added by Lydia 2023/05/10 開放輸入「員工編號」欄：總經理
   If InStr("01,08,", Pub_strUserST05 & ",") > 0 Then
      txtUsernum.Enabled = True
   End If
   'end 2023/05/10
   
   'Modify by Morgan 2011/4/21 從 Unload 移來(因畫面沒離開時沒寫Log會造成逾時重新登入後重複執行)
   PUB_AddExcuteLog Me.Name
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   DestroyToolTip '清除物件
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Memo by Lydia 2019/11/04 國外部自動通知順序: FMP案frm060206=> 國外部期限frm060204=> 外商frm030301=> 外法frm072005=>國外部行事曆frm060209

   If Not bolUnloading Then 'Add by Morgan 2011/3/11
      Dim strSql As String, bolRun As Boolean
      
      'Add By Sindy 2009/08/28
      '電腦中心除外
      If Pub_StrUserSt03 <> "M51" Then
         '商標
         bolRun = False
         'Modified by Lydia 2016/09/14 cp27||cp57 is null => NVL(CP27,0)=0 AND NVL(CP57,0)=0
         strSql = "select sum(aa) from ( " & _
                        "select count(*) as aa from caseprogress " & _
                        "where cp01 in ('FCT','CFT','CFC','S') " & _
                        "and NVL(CP27,0)=0 AND NVL(CP57,0)=0 " & _
                        "and (cp06>0 or cp48>0) " & _
                        "and cp13='" & strUserNum & "' " & _
                        "Union All " & _
                        "select count(*) as aa from caseprogress " & _
                        "where cp01 in ('FCT','CFT','CFC','S') " & _
                        "and NVL(CP27,0)=0 AND NVL(CP57,0)=0 " & _
                        "and (cp06>0 or cp48>0) " & _
                        "and cp14='" & strUserNum & "')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               bolRun = True
            End If
         End If
         If CheckUse("frm030301", strExec, False) = True Or bolRun = True Then
            'Modify By Sindy 2022/12/14 外商系統的期限彈跳提醒自動執行功能，請改為早上及下午各一次
            'strSql = "select * from executelog where el01='frm030301' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
            'Modify By Sindy 2025/9/3 琬姿副理在反應期限會一直啟動,故再調整一下判斷
            If ServerTime >= 130000 Then
               strSql = "select * from executelog where el01='frm030301' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
            Else
               strSql = "select * from executelog where el01='frm030301' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
            End If
            '2025/9/3 END
            'strSQL = "select * from executelog where el01='frm030301' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
            'strSql = "select * from executelog where el01='frm030301' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI <> 1 Then
               pub_bolInformCheck = True 'Add By Sindy 2009/09/21
               Load frm030301
               frm030301.cmdQuery(0).Value = True
               Exit Sub
            End If
         End If
         '法務
         bolRun = False
         'Modified by Lydia 2016/09/14 cp27||cp57 is null => NVL(CP27,0)=0 AND NVL(CP57,0)=0
         strSql = "select count(*) from caseprogress " & _
                        "where cp01 in ('CFL','FCL','LIN','L','LA') " & _
                        "and NVL(CP27,0)=0 AND NVL(CP57,0)=0 and cp06>0 " & _
                        "and cp14='" & strUserNum & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               bolRun = True
            End If
         End If
         'If strGroup = "F1" Or strGroup = "F2" Or strGroup = "D4" Or bolRun = True Then
            If CheckUse("frm072005", strExec, False) = True Or bolRun = True Then
               strSql = "select * from executelog where el01='frm072005' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI <> 1 Then
                  pub_bolInformCheck = True 'Add By Sindy 2009/09/21
                  Load frm072005
                  frm072005.cmdQuery(0).Value = True
                  Exit Sub
               End If
            End If
         'End If
         
         'Added by Lydia 2015/12/30 國外部行事曆通知(每天早上和下午自動執行時才run)
         'Modified by Lydia 2016/01/27 判斷不完整
         'If pub_bolInformCheck And Left(Pub_StrUserSt03, 2) = "F2" Then
         '   If CheckUse("frm060209", strExec) = True Then
         '      Load frm060209
        '       Exit Sub
         '   End If
         'End If
         If Left(Pub_StrUserSt03, 2) = "F2" Then
            If CheckUse("frm060209", strExec, False) = True Then
                strSql = "select * from executelog where el01='frm060209' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                If intI <> 1 Then
                   pub_bolInformCheck = True
                   Load frm060209
                   pub_bolInformCheck = False
                   Exit Sub
                End If
            End If
         End If
         'end 2016/01/27
      End If
      
      MenuEnabled
   End If
   
   DestroyToolTip '清除物件
   Set frm060204 = Nothing
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = grdDataList.Recordset.RecordCount
End Sub

Private Sub grdDataList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
   
   If grdDataList.MouseRow <> 0 And _
      (grdDataList.MouseCol = 12 Or grdDataList.MouseCol = 13 Or grdDataList.MouseCol = 14) Then
      If iRow <> grdDataList.MouseRow Or iCol <> grdDataList.MouseCol Then
         If grdDataList.TextMatrix(grdDataList.MouseRow, grdDataList.MouseCol) <> "" Then
            CreateToolTip GetHWndForToolTip(grdDataList), grdDataList.TextMatrix(grdDataList.MouseRow, grdDataList.MouseCol)
            iRow = grdDataList.MouseRow
            iCol = grdDataList.MouseCol
         End If
      End If
   End If
End Sub

'Modify By Sindy 2020/5/18 Mark,敏莉:在國外部專利處期限通知畫面，當我們點選依XXXX排序時(例: 以本所期限排序)，案件性質是"客戶提供文件"及事件是"未請款"的案件，依然保持在頁面的最下方不動，以方便檢查智慧局期限
Private Sub grdDataList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim iCol As Integer
    iCol = grdDataList.MouseCol
    If grdDataList.MouseRow < 1 Then
      'Added by Morgan 2024/12/12 若斷線重連資料集會是關閉狀態無法再操作
      If m_adoRst.State = adStateClosed Then
         MsgBox "連線逾時無法排序，請重新查詢！", vbExclamation
         Exit Sub
      End If
      'end 2024/12/12
      grdDataList.Visible = False
      ChgEmptyDate True
      Set grdDataList.Recordset = Nothing
      If m_blnColOrderAsc = True Then
         m_adoRst.Sort = "sort asc," & m_adoRst.Fields(iCol).Name & " desc"
         m_blnColOrderAsc = False
      Else
         m_adoRst.Sort = "sort asc," & m_adoRst.Fields(iCol).Name & " asc"
         m_blnColOrderAsc = True
      End If
      SetRst2Grid
      SetGrid
      SetColor
      'Add By Sindy 2023/8/31
      textSys.Tag = "": LblCnt.Visible = False '還原預設值
      Call TextSys_LostFocus
      '2023/8/31 END
      grdDataList.Visible = True
    End If
End Sub

'Modify By Sindy 2020/5/18 Mark,敏莉:在國外部專利處期限通知畫面，當我們點選依XXXX排序時(例: 以本所期限排序)，案件性質是"客戶提供文件"及事件是"未請款"的案件，依然保持在頁面的最下方不動，以方便檢查智慧局期限
''Added by Lydia 2019/05/31
'Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Dim nCol As Long, nRow As Long
'   getGrdColRow grdDataList, x, y, nCol, nRow
'   If nCol < 0 Or nRow < 0 Then Exit Sub
'   grdDataList.col = nCol
'   grdDataList.row = nRow
'   If Me.grdDataList.row < 1 And Me.grdDataList.Text <> "V" Then
'      '全部都是文字(保留數值排序)
'      'If InStr("公 告 日,申請案號", Me.grdDataList.Text) > 0 Then
'      '   If m_blnColOrderAsc = True Then
'      '      Me.grdDataList.Sort = 3  '數值昇冪
'      '      m_blnColOrderAsc = False
'      '   Else
'      '      Me.grdDataList.Sort = 4 '數值降冪
'      '      m_blnColOrderAsc = True
'      '   End If
'      'Else
'         If m_blnColOrderAsc = True Then
'            Me.grdDataList.Sort = 5 '字串昇冪
'            m_blnColOrderAsc = False
'         Else
'            Me.grdDataList.Sort = 6 '字串降冪
'            m_blnColOrderAsc = True
'         End If
'      'End If
'   End If
'End Sub

Private Sub ChgEmptyDate(Optional p_bolBeforeSort As Boolean)
   Dim ii As Integer, jj As Integer
   With grdDataList
   If .Rows > 1 Then
      For ii = 1 To .Rows - 1
         For jj = 1 To 5 '4
            If p_bolBeforeSort Then
               If .TextMatrix(ii, jj) = "" Then
                  .TextMatrix(ii, jj) = "2"
               End If
            Else
               If .TextMatrix(ii, jj) = "2" Then
                  .TextMatrix(ii, jj) = ""
               End If
            End If
         Next
      Next
   End If
   End With
End Sub

Private Sub grdDataList_SelChange()
   Dim ii As Integer, lngColor As Long
   With grdDataList
      If .MouseRow > 0 Then
         .Visible = False
         .row = .MouseRow
         .col = 0
         If .Text = "V" Then
            .Text = ""
            .col = 0
            .CellBackColor = .BackColor
            .col = 3
            lngColor = .CellBackColor
            For ii = 1 To 2
               .col = ii
               .CellBackColor = lngColor
            Next
         Else
            .Text = "V"
            For ii = 0 To 2
               .col = ii
               .CellBackColor = &HFFC0C0
            Next
         End If
         .Visible = True
      End If
   End With
End Sub

'Add By Sindy 2023/8/31
Private Sub textSys_GotFocus()
   TextInverse textSys
   CloseIme
End Sub
Private Sub textSys_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub TextSys_LostFocus()
Dim dblCnt As Double
Dim ii As Integer, intCurrRow As Integer
Dim strTit As String, strMsg As String
   
   '檢查使用者是否有使用該系統類別的權限
   If textSys.Text <> "" Then
      If IsUserHasRightOfSystem(strUserNum, textSys.Text) = False _
         And InStr("P,CFP,CPS,PS,", textSys.Text & ",") = 0 Then
         strTit = "資料檢核"
         strMsg = "無此系統別 或 無使用此系統別的權限！"
         MsgBox strMsg, vbOKOnly, strTit
         textSys_GotFocus
         Exit Sub
      End If
   End If
   
   '過濾系統別
   If textSys.Tag <> textSys.Text Then
      dblCnt = 0
      With grdDataList
      If .Rows > 1 Then
         .Visible = False
         For ii = 1 To .Rows - 1
            '***** 恢復原狀況 *****
            If .TextMatrix(ii, 36) = "" Then
               .RowHeight(ii) = 0
            Else
               .RowHeight(ii) = 255
            End If
            '**********************
            
            If .RowHeight(ii) > 0 Then
               If textSys.Text <> "" And textSys.Text <> Trim(.TextMatrix(ii, 21)) Then
                  .RowHeight(ii) = 0
               Else
                  dblCnt = dblCnt + 1
                  If dblCnt = 1 Then intCurrRow = ii
               End If
            End If
         Next ii
         If intCurrRow > 0 Then
            .TopRow = intCurrRow '1
         End If
         .Visible = True
      End If
      End With
      If textSys.Text <> "" Then
         LblCnt.Visible = True
         LblCnt.Caption = "有 " & dblCnt & " 筆"
         If dblCnt = 0 Then
            MsgBox "無 " & textSys.Text & " 系統別的資料！", vbInformation
         End If
      Else
         LblCnt.Visible = False
         LblCnt.Caption = ""
      End If
   End If
   textSys.Tag = textSys.Text
End Sub
'2023/8/31 END

Private Sub txtUsernum_Change()
   If Len(txtUsernum) >= 5 Then
      lblUserName = GetStaffName(txtUsernum, True)
   Else
      lblUserName = ""
   End If
End Sub

Private Sub txtUsernum_GotFocus()
   TextInverse txtUsernum
End Sub

Private Function SetXRecord()
Dim iRow As Integer
Dim strNote As String
   
   With m_adoRst
      .MoveFirst
      Do While Not .EOF
         'Modify By Sindy 2015/12/28 寫成共用函數
         '                                    cp01              cp09              cp10              cp43              cnt               np22              cp142              cp164            cp176
         strNote = PUB_GetFCPAddQuyNotes("" & .Fields(21), "" & .Fields(28), "" & .Fields(26), "" & .Fields(31), "" & .Fields(30), "" & .Fields(32), "" & .Fields(33), "" & .Fields(34), "" & .Fields(35))
         If strNote <> "" Then
            'Modify By Sindy 2021/4/21
            If Left(.Fields(13), 3) = "達指定" Or Left(.Fields(13), 3) = "達本所" Then
               .Fields(13) = Left(.Fields(13), 4) & strNote & Mid(.Fields(13), 5)
            Else
            '2021/4/21 END
               .Fields(13) = strNote & .Fields(13)
            End If
         End If
         '2015/12/28 END
         'Add By Sindy 2017/1/18
         'If .Fields(10) = "准未請款" Or .Fields(10) = "通知告准" Then
         If .Fields(10) = "准未請款" Then
            .Fields(13) = "優先請款;" & .Fields(13)
         End If
         '2017/1/18 END
         
'         'Add By Sindy 2015/10/22
'         '201.新案翻譯,新案202.補文件,205.申復,501.訴願等案性質不能延期二次,延期過一次,備註"不得延期"
'         If .Fields(20) = "FCP" Then
'            If (.Fields(25) = "201" Or .Fields(25) = "202" Or .Fields(25) = "205" Or .Fields(25) = "501") _
'               And Val(.Fields(29)) >= 1 Then
'               If .Fields(25) = "202" And "" & .Fields(30) <> "" Then
'                  '檢查是否為新案補文件
'                  strExc(0) = "select cp09,cp10 from caseprogress" & _
'                     " where cp09='" & .Fields(30) & "' and cp10 in(" & NewCasePtyList & ")"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     .Fields(12) = "不得延期;" & .Fields(12)
'                  End If
'               Else
'                  .Fields(12) = "不得延期;" & .Fields(12)
'               End If
'            '107.再審不能延期三次,延期過二次,備註"不得延期"
'            ElseIf .Fields(25) = "107" And Val("" & .Fields(29)) >= 2 Then
'               .Fields(12) = "不得延期;" & .Fields(12)
'            '503.行政訴訟不能延期,備註"不得延期"
'            ElseIf .Fields(24) = "503" Then
'               .Fields(12) = "不得延期;" & .Fields(12)
'            End If
'         '2015/10/22 END
'         'Add By Sindy 2015/10/29 205.陳述意見
'         ElseIf .Fields(20) = "P" And .Fields(25) = "205" Then
'            If Val(.Fields(31)) > 0 Then 'R12=NP22:下一程序
'               strExc(0) = "select np01,np09,cp07,to_char(add_months(to_date(cp07,'YYYYMMDD'),2),'YYYYMMDD')" & _
'                           " From nextprogress,caseprogress" & _
'                           " where np01='" & .Fields(27) & "' and np22=" & .Fields(31) & _
'                           " and np01=cp09(+)" & _
'                           " and cp07 is not null"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  If Val(RsTemp(1)) >= Val(RsTemp(3)) Then
'                     .Fields(12) = "不得延期;" & .Fields(12)
'                  End If
'               End If
'            Else '進度檔
'               strExc(0) = "select c1.cp43,c1.cp07,c2.cp07,to_char(add_months(to_date(c2.cp07,'YYYYMMDD'),2),'YYYYMMDD')" & _
'                           " from caseprogress c1,caseprogress c2" & _
'                           " where c1.cp09='" & .Fields(27) & "'" & _
'                           " and substr(c1.cp43,1,1)='C'" & _
'                           " and c1.cp43=c2.cp09(+)" & _
'                           " and c2.cp07 is not null"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  If Val(RsTemp(1)) >= Val(RsTemp(3)) Then
'                     .Fields(12) = "不得延期;" & .Fields(12)
'                  End If
'               End If
'            End If
'         End If
'         '2015/10/29 END
'         'Add By Sindy 2015/12/15
'         If Val("" & .Fields(32)) > 0 Then 'CP142:指定送件日期
'            .Fields(12) = "當天;" & .Fields(12)
'         End If
'         '2015/12/15 END
         
         If .Fields(16) = "D" Then
            strExc(0) = "select nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)),0) from acc1k0" & _
               " where a1k13 = '" & .Fields(21) & "' and a1k14 = '" & .Fields(22) & "' and a1k15 = '" & .Fields(23) & "' and a1k16 = '" & .Fields(24) & "' and (a1k29 is null or a1k29 = '')"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp(0) > 0 Then
                  .Fields(25) = Format(RsTemp(0), "#,###")
               End If
            End If
         '主動修正或實審若有新案已發文未請款則請款期限同新案
         'Modify by Morgan 2008/10/9 改所有非新申請案的都檢查若新案未請款則不顯示
         'Modify by Morgan 2010/4/13 +工程師提申940,105
         ElseIf .Fields(16) = "F" And InStr("101,102,103,105,940", .Fields(26)) = 0 Then
            'Modify by Morgan 2009/12/3 +判斷新案承辦是國外部同仁否則會沒有人管
            'Modified by Morgan 2024/12/20 +發文日>19221111及要請款的條件 Ex:FCP-066689主動修正,實審 --Winfrey
            strExc(0) = "select 1 from caseprogress,staff" & _
               " where cp01 = '" & .Fields(21) & "' and cp02 = '" & .Fields(22) & "'" & _
               " and cp03 = '" & .Fields(23) & "' and cp04 = '" & .Fields(24) & "'" & _
               " AND CP10 IN ('101','102','103','105','125','940') and cp60 is null and st01(+)=cp14 and substr(st03,1,1)='F' and cp27>19221111 and cp20 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               .Delete
            End If
         End If
         .MoveNext
      Loop
      .UpdateBatch
   End With
End Function

'Add By Sindy 2010/11/26
Private Sub txtUsernum_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

'Add by Amy 20130703 確認要查詢的人是否為主管級
Private Function CheckManage() As String
  CheckManage = ""
  strExc(0) = "Select distinct 'ST52' Manage From Staff WHERE  ST52='" & txtUsernum & "'  " & _
                    "Union Select distinct 'ST5X' Manage From Staff WHERE  ST55='" & txtUsernum & "' Or ST53='" & txtUsernum & "' Or ST54='" & txtUsernum & "'  "
     
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
        With RsTemp
         Do While Not .EOF
            CheckManage = CheckManage & RsTemp.Fields("Manage") & ","
            .MoveNext
         Loop
      End With
      CheckManage = Left(CheckManage, Len(CheckManage) - 1)
   End If
End Function

'Modify By Sindy 2023/8/31 留舊函數,若有需要查看歷史過程可查看
'語法內有用組合欄位為條件以控制使用特定index(避掉某些不適當的)
Private Sub doQuery_Old(idx As Integer)
   Dim stVTB As String, stDate1 As String, stDate2 As String, stDate3 As String, stDate4 As String, stDate7 As String
   Dim stSQL As String, stCon As String, stConCP14 As String, stConEP04 As String, stConNA16 As String
   Dim stConNA51 As String, stConNP10 As String
   Dim stConCP06 As String, stConCP48 As String, stConEP08 As String
   Dim stConCP142 As String 'Add By Sindy 2021/4/20
   Dim stNumList As String, stDept As String
   Dim ii As Integer, jj As Integer, stIdList
   Dim stUserID As String
   Dim stCP01 As String
   Dim stConPAGrp As String, stConSPGrp As String 'Add by Morgan 2009/9/7
   Dim stConNA51P As String, stConSP26 As String 'Add by Lydia 2014/11/14 FCP承辦區域特殊狀況之智權人員劃分方式
   Dim stConNA16L As String, stConNA51L 'Add By Sindy 2015/10/22
   Dim rsA As ADODB.Recordset 'Add By Sindy 2017/1/16
   Dim strCP09 As String, strCP60 As String, strA1K01 As String, strA1K19 As String, strA1K20 As String
   Dim strDST01 As String 'Add By Sindy 2017/1/18
   Dim stOrdCon As String 'Added by Lydia 2018/02/08
   Dim strCPM1933_Col As String, strCPM1933_Where As String 'Add By Sindy 2020/8/7
   
   'Add by Morgan 2009/3/26
   stCP01 = " and cp01 in ('P','PS','FCP','FG','CFP','CPS')"
   
   If lblUserName = "" Then
      MsgBox "員工編號錯誤！", vbExclamation
      If txtUsernum.Enabled = True Then
         txtUsernum_GotFocus
         txtUsernum.SetFocus
      End If
      Exit Sub
   'Added by Morgan 2015/10/5
   ElseIf Pub_StrUserSt03 = "F22" And txtUsernum <> strUserNum Then
      If PUB_GetST03(txtUsernum) <> Pub_StrUserSt03 Then
         MsgBox "員工編號錯誤！", vbExclamation, "權限不足"
         Exit Sub
      End If
   'end 2015/10/5
   End If
   
   stUserID = txtUsernum
   If stUserID = strUserNum Then
      stDept = Pub_StrUserSt15
   Else
      stDept = GetST15(stUserID)
   End If
   
   stNumList = PUB_GetMapID(stUserID, 0)
   If stNumList <> "" Then
      stNumList = "'" & stNumList & "','" & stUserID & "'"
   Else
      stNumList = "'" & stUserID & "'"
   End If
   stNumList1(1) = stNumList
   
   For ii = 2 To 5
      stNumList1(ii) = GetNumList(stUserID, ii)
      If stNumList1(ii) <> "" Then
         stNumList = stNumList & "," & stNumList1(ii)
      End If
   Next
   
   stIdList = Split(stNumList, ",")
   'Add by Morgan 2009/12/18 去除重複編號
   stNumList = stIdList(LBound(stIdList))
   For ii = LBound(stIdList) + 1 To UBound(stIdList)
      For jj = LBound(stIdList) To ii - 1
         If stIdList(jj) = stIdList(ii) Then
            Exit For
         End If
      Next
      If ii = jj Then
         stNumList = stNumList & "," & stIdList(ii)
      End If
   Next
   stIdList = Split(stNumList, ",")
   'end 2009/12/18
   
   stDate1 = strSrvDate(1) - 10000
   stDate2 = CompWorkDay(4, strSrvDate(1))
   stDate3 = CompWorkDay(7, strSrvDate(1))
   'stDate3指本所期限前五個工作天的條件,但因當天不算所以為系統日+6天
   '又因6天後若為星期五,則星期六日的期限也要在當天出現,所以為系統日+7天且判斷條件為CP06<
   
   stDate7 = CompWorkDay(8, strSrvDate(1), 1) 'Add By Sindy 2022/3/15 系統日前7個工作天(不含當日)
   
   stConCP06 = " AND CP06>=" & stDate1 & " AND CP06< " & stDate2
   stConCP48 = " AND CP48>=" & stDate1 & " AND CP48< " & stDate2
   stConEP08 = " AND EP08>=" & stDate1 & " AND EP08< " & stDate2
   stConCP142 = " AND CP142>=" & stDate1 & " AND CP142< " & stDate2 'Add By Sindy 2021/4/20
   
   bLvlX = CheckLevel(stUserID, "M") '未交稿,已完稿無核稿管制人
   bLvl4 = CheckLevel(stUserID, "N") '第四級管制人(+FCP,FG未分案將到期) :[有關期限] 國外部專利處非外專承辦或未分案將到期管制人(含逾期)
   bLvl5 = CheckLevel(stUserID, "O") '第五級管制人(+FCP,FG未分案已逾期) :國外部專利處非外專承辦或未分案已逾期管制人
   bLvO1 = CheckLevel(stUserID, "O1") 'Added by Lydia 2022/12/20 國外部期限通知未分案管制人
   
   'Add by Morgan 2009/9/7 未分案改分組管制
   bLvlChm = CheckLevel(stUserID, "R") '未分案化學組管制人
   bLvlJpn = CheckLevel(stUserID, "S") '未分案日文組管制人
   '2011/11/30 modify by sonia 取消德文且機電拆電子電機及機械設計
   'bLvlEls = CheckLevel(stUserID, "T") '未分案機電德文其他組管制人
   bLvlMot = CheckLevel(stUserID, "T1") '未分案機械設計組管制人
   bLvlEls = CheckLevel(stUserID, "T") '未分案電子電機其他組管制人
   '2011/11/30 end
   
   'Add by Lydia 2014/11/14 FCP承辦區域特殊狀況之智權人員劃分方式
   '代理人Y51333010=Pub_GetSpecMan("北京銀龍FCP案承辦業務") ,NA51 = decode(pa75,'Y51333010','" & midStr & "',na51)
   Dim midStr As String
   'Modified by Lydia 2016/02/03改成回傳case句
   'midStr = Pub_GetSpecMan("北京銀龍FCP案承辦業務")
   midStr = Pub_GetSpecFCP
   
   If InStr(stNumList, ",") > 0 Then
      stConCP14 = " AND CP14 in (" & stNumList & ") "
      stConEP04 = " AND EP04 in (" & stNumList & ") "
      'Modified by Lydia 2022/11/03 區分FMP案
      'stConNA16 = " AND NA16 in (" & stNumList & ") "
      stConNA16 = " AND ((CP01 in ('FCP','FG') and NA16 in (" & stNumList & ")) or (CP01 not in ('FCP','FG') and nvl(NA79,NA16) in (" & stNumList & "))) "
      stConNA16L = " AND (n1.NA16 in (" & stNumList & ") OR n2.NA16 in (" & stNumList & "))" 'Add By Sindy 2015/10/22
      stConNA51 = " AND NA51 in (" & stNumList & ") "
      stConNP10 = " AND NP10 in (" & stNumList & ") "
      'Add by Lydia 2014/11/14 FCP承辦區域特殊狀況之智權人員劃分方式
      'Modified by Lydia 2016/02/03  +Y51817040日文承辦
'      stConNA51P = " AND decode(pa75,'Y51333010','" & midStr & "',na51) in (" & stNumList & ") "
'      stConSP26 = " AND decode(sp26,'Y51333010','" & midStr & "',na51) in (" & stNumList & ") "
'      stConNA51L = " AND (decode(LC22,'Y51333010','" & midStr & "',n1.na51) in (" & stNumList & ")" & _
'                        " OR decode(LC22,'Y51333010','" & midStr & "',n2.na51) in (" & stNumList & "))" 'Add By Sindy 2015/10/22
      stConNA51P = " AND decode(pa75," & midStr & ",na51) in (" & stNumList & ") "
      stConSP26 = " AND decode(sp26," & midStr & ",na51) in (" & stNumList & ") "
      stConNA51L = " AND (decode(LC22," & midStr & ",n1.na51) in (" & stNumList & ")" & _
                        " OR decode(LC22," & midStr & ",n2.na51) in (" & stNumList & "))"
   Else
      stConCP14 = " AND CP14=" & stNumList
      stConEP04 = " AND EP04=" & stNumList
      'Modified by Lydia 2022/11/03 區分FMP案
      'stConNA16 = " AND NA16=" & stNumList
      stConNA16 = " AND ((CP01 in ('FCP','FG') and NA16 =" & stNumList & " ) or (CP01 not in ('FCP','FG') and nvl(NA79,NA16) =" & stNumList & " )) "
      stConNA16L = " AND (n1.NA16=" & stNumList & " OR n2.NA16=" & stNumList & ")" 'Add By Sindy 2015/10/22
      stConNA51 = " AND NA51=" & stNumList
      stConNP10 = " AND NP10=" & stNumList
      'Add by Lydia 2014/11/14 FCP承辦區域特殊狀況之智權人員劃分方式
      'Modified by Lydia 2016/02/03  +Y51817040日文承辦
'      stConNA51P = " AND decode(pa75,'Y51333010','" & midStr & "',na51) =" & stNumList
'      stConSP26 = " AND decode(sp26,'Y51333010','" & midStr & "',na51) =" & stNumList
'      stConNA51L = " AND (decode(LC22,'Y51333010','" & midStr & "',n1.na51) =" & stNumList & _
'                        " OR decode(LC22,'Y51333010','" & midStr & "',n2.na51) =" & stNumList & ")" 'Add By Sindy 2015/10/22
      stConNA51P = " AND decode(pa75," & midStr & ",na51) =" & stNumList
      stConSP26 = " AND decode(sp26," & midStr & ",na51) =" & stNumList
      stConNA51L = " AND (decode(LC22," & midStr & ",n1.na51) =" & stNumList & _
                        " OR decode(LC22," & midStr & ",n2.na51) =" & stNumList & ")"
   End If
   
   'Add by Morgan 2009/10/22
   '清除暫存檔
   stSQL = "delete R060204 where R01='" & strUserNum & "'"
   cnnConnection.Execute stSQL, intI

   'Modify By Sindy 2017/1/18 + 'I','准未請款','J','分割建議','K','通知告准'
   'Modified by Lydia 2017/11/29 + 'L','待命名期限'
   'Modified by Lydia 2018/02/08  + 'M','待處理'
   '代碼1(R02):A=達本所,B=達承辦,C=達核稿,D=未收文,E=未發文,F=未請款,G=未交稿,H=早收文(Added by Lydia 2015/09/09)
   '           I=准未請款,J=分割建議,K=通知告准,L=待命名期限,M=待處理,N=達指定,(英文的)O=未分案,P=達約定
   '代碼2(R03):(數字的)0=未分案,1=承辦人,2=管制人,3=智權人員,4=核稿人,8=無核稿人 9=未交稿
   'R01:UserID,R02:代碼1,R03:代碼2,R04:收文號,R05:國籍,R06:承辦,R07:業務,R08:所限,R09:法限
   ',R10:辦限,R11:核限,R12:NP22,R13:備註,R14:指定日,R15:約定期限
   
   'Modify by Morgan 2009/10/22 改寫暫存檔以便過濾達所限又達承辦或核稿期限的資料(只要顯示達所限)
   '已收文未發文,2個工作天後達本所期限者(不含當日) --承辦人-A1(本所期限,承辦人)
   'Modified by Lydia 2016/09/14 CP14||CP57||CP27 => CP158=0 AND CP159=0 AND CP14
   'Modified by Lydia 2016/09/22 +/*+ INDEX(CASEPROGRESS IDXCP15815914) */
   'Modified by Lydia 2018/02/08 客戶提供文件1920 =>'M','待處理'
   'stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
      " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815914) */ '" & strUserNum & "','A' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
      " From CASEPROGRESS,PATENT,FAGENT" & _
      " WHERE CP158=0 AND CP159=0 AND CP14 IN(" & stNumList & ")" & stConCP06 & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
      " SELECT '" & strUserNum & "',DECODE(CP10,'1920','M','A') EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
      " From CASEPROGRESS,PATENT,FAGENT" & _
      " WHERE CP158=0 AND CP159=0 AND CP14 IN(" & stNumList & ")" & stConCP06 & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
   cnnConnection.Execute stSQL, intI
   'Add By Sindy 2021/4/21 達指定
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
      " SELECT '" & strUserNum & "','N' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
      " From CASEPROGRESS,PATENT,FAGENT" & _
      " WHERE CP158=0 AND CP159=0 AND CP14 IN(" & stNumList & ")" & stConCP142 & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
   cnnConnection.Execute stSQL, intI
   
   'Add By Sindy 2015/10/22 +法務:達本所
   'Modified by Lydia 2016/09/14 CP14||CP57||CP27 => CP158=0 AND CP159=0 AND CP14
   'Modified by Lydia 2016/09/22 +/*+ INDEX(CASEPROGRESS IDXCP15815914) */
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
      " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,nvl(FA10,cu10),CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
      " From CASEPROGRESS,LawCase,FAGENT,Customer" & _
      " WHERE CP158=0 AND CP159=0 AND CP14 IN(" & stNumList & ")" & stConCP06 & _
      " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC08||LC34 IS NULL AND LC01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(LC22,1,8) AND FA02(+)=SUBSTR(LC22,9)" & _
      " AND CU01(+)=SUBSTR(LC11,1,8) AND CU02(+)=SUBSTR(LC11,9)"
   cnnConnection.Execute stSQL, intI
   
   'Modified by Lydia 2016/09/14 CP14||CP57||CP27 => CP158=0 AND CP159=0 AND CP14
   'Modified by Lydia 2016/09/22 +/*+ INDEX(CASEPROGRESS IDXCP15815914) */
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
      " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
      " From CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
      " WHERE CP158=0 AND CP159=0 AND CP14 IN(" & stNumList & ")" & stConCP06 & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
   cnnConnection.Execute stSQL, intI
   'Add By Sindy 2021/4/21 達指定
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
      " SELECT '" & strUserNum & "','N' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
      " From CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
      " WHERE CP158=0 AND CP159=0 AND CP14 IN(" & stNumList & ")" & stConCP142 & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
   cnnConnection.Execute stSQL, intI
   
   '已收文未發文,2個工作天後達承辦期限者(不含當日) --承辦人-B1(承辦期限,承辦人)
   'Modified by Lydia 2016/09/14 CP14||CP57||CP27 => CP158=0 AND CP159=0 AND CP14
   'Modified by Lydia 2016/09/22 +/*+ INDEX(CASEPROGRESS IDXCP15815914) */
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
      " SELECT '" & strUserNum & "','B' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
      " From CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT" & _
      " WHERE CP158=0 AND CP159=0 AND CP14 IN(" & stNumList & ")" & stConCP48 & _
      " AND EP02(+)=CP09 AND EP09 IS NULL" & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
   cnnConnection.Execute stSQL, intI

   'Modified by Lydia 2016/09/14 CP14||CP57||CP27 => CP158=0 AND CP159=0 AND CP14
   'Modified by Lydia 2016/09/22 +/*+ INDEX(CASEPROGRESS IDXCP15815914) */
   stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
      " SELECT '" & strUserNum & "','B' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
      " From CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT" & _
      " WHERE CP158=0 AND CP159=0 AND CP14 IN(" & stNumList & ")" & stConCP48 & _
      " AND EP02(+)=CP09 AND EP09 IS NULL" & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
      " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
   cnnConnection.Execute stSQL, intI

   '所有未發文--承辦人-E(未發文)
   If idx = 1 Then
      For ii = LBound(stIdList) To UBound(stIdList)
         'Modified by Lydia 2016/09/14 CP57||CP27 is null => CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS,PATENT,FAGENT" & _
            " WHERE CP05>20030000 AND CP14=" & stIdList(ii) & " and CP158=0 AND CP159=0" & _
            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
         cnnConnection.Execute stSQL, intI
         
         'Add By Sindy 2015/10/22 +法務:未發文
         'Modified by Lydia 2016/09/14 CP57||CP27 is null => CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,nvl(FA10,cu10),CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS,Lawcase,FAGENT,Customer" & _
            " WHERE CP05>20030000 AND CP14=" & stIdList(ii) & " and CP158=0 AND CP159=0" & _
            " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC08||LC34 IS NULL AND LC01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(LC22,1,8) AND FA02(+)=SUBSTR(LC22,9)" & _
            " AND CU01(+)=SUBSTR(LC11,1,8) AND CU02(+)=SUBSTR(LC11,9)"
         cnnConnection.Execute stSQL, intI
         '2015/10/22 END
         
         'Modified by Lydia 2016/09/14 CP57||CP27 is null => CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815914) */
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
            " WHERE CP05>20030000 AND CP14=" & stIdList(ii) & " AND CP158=0 AND CP159=0" & _
            " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
         cnnConnection.Execute stSQL, intI
      Next ii
   End If
   
   '已發文未請款--承辦人
   For ii = LBound(stIdList) To UBound(stIdList)
      'Modify By Sindy 2020/8/7 為配合日本部及專利國外部，於主管會議中經各組主管同意，
      '工程師實施無卷請款，為控管避免請款延遲太久，系統未請款彈跳期限通知，
      '工程師請款部分（參附件標黃色）請改為5個工作天
'      strCPM1933_Col = "TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD')"
'      strCPM1933_Where = "TO_CHAR(ADD_MONTHS(SYSDATE,-1*DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD')"
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD'),WORKDAYADD(cpm33+1,CP27))"
      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
      'Modify by Morgan 2010/4/12 排除工程師提申940
      'Modified by Lydia 2016/09/14 CP20||CP57||CP60 IS NULL AND CP14 => CP27>20150914 AND CP14='...' AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,CP09,FA10,CP14,CP13" & _
         ",NULL,NULL," & strCPM1933_Col & " C06,NULL,0" & _
         " FROM CASEPROGRESS,PATENT,FAGENT,CASEPROPERTYMAP" & _
         " WHERE CP27>" & stDate1 & " AND CP14=" & stIdList(ii) & " AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'" & _
         " And CP09<'B' AND CP10 NOT IN ('403','404','411','418','419','901','902','907','908','913','940')" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP27<=" & strCPM1933_Where
      cnnConnection.Execute stSQL, intI
      
      'Modify By Sindy 2020/8/7 為配合日本部及專利國外部，於主管會議中經各組主管同意，
      '工程師實施無卷請款，為控管避免請款延遲太久，系統未請款彈跳期限通知，
      '工程師請款部分（參附件標黃色）請改為5個工作天
'      strCPM1933_Col = "TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),nvl(CPM19,0)),'YYYYMMDD')"
'      strCPM1933_Where = "TO_CHAR(ADD_MONTHS(SYSDATE,-1*nvl(CPM19,0)),'YYYYMMDD')"
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),nvl(CPM19,0)),'YYYYMMDD'),WORKDAYADD(cpm33+1,CP27))"
      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*nvl(CPM19,0)),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
      'Add By Sindy 2015/10/22 +法務:未請款
      'Modified by Lydia 2016/09/14 CP20||CP57||CP60 IS NULL AND CP14 => CP27>20150914 AND CP14='...' AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,CP09,nvl(FA10,cu10),CP14,CP13" & _
         ",NULL,NULL," & strCPM1933_Col & " C06,NULL,0" & _
         " FROM CASEPROGRESS,Lawcase,FAGENT,CASEPROPERTYMAP,Customer" & _
         " WHERE CP27>" & stDate1 & " AND CP14=" & stIdList(ii) & " AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'" & _
         " And CP09<'B'" & _
         " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC08||LC34 IS NULL AND LC01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(LC22,1,8) AND FA02(+)=SUBSTR(LC22,9)" & _
         " AND CU01(+)=SUBSTR(LC11,1,8) AND CU02(+)=SUBSTR(LC11,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP27<=" & strCPM1933_Where
      cnnConnection.Execute stSQL, intI
      '2015/10/22 END
      
      'Modify By Sindy 2020/8/7 為配合日本部及專利國外部，於主管會議中經各組主管同意，
      '工程師實施無卷請款，為控管避免請款延遲太久，系統未請款彈跳期限通知，
      '工程師請款部分（參附件標黃色）請改為5個工作天
'      strCPM1933_Col = "TO_CHAR(ADD_MONTHS(TO_DATE(C2.CP27,'YYYYMMDD'),CPM19),'YYYYMMDD')"
'      strCPM1933_Where = "TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD')"
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(C2.CP27,'YYYYMMDD'),CPM19),'YYYYMMDD'),WORKDAYADD(cpm33+1,C2.CP27))"
      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
      'Add by Morgan 2010/4/12 工程師提申=940 要抓新申請案的設定
      'Modified by Lydia 2016/09/14 CP20||CP57||CP60 IS NULL AND CP14 => CP27>20150914 AND CP14='...' AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,C2.CP09,FA10,C2.CP14,C2.CP13" & _
         ",NULL,NULL," & strCPM1933_Col & " C06,NULL,0" & _
         " FROM CASEPROGRESS C1,CASEPROGRESS C2,PATENT,FAGENT,CASEPROPERTYMAP" & _
         " WHERE C1.CP27>" & stDate1 & " AND C1.CP14=" & stIdList(ii) & " AND C1.CP159=0 AND C1.CP16>0 AND C1.CP20||C1.CP60 IS NULL" & _
         " And C1.CP09<'B' AND C1.CP10='940'" & _
         " AND PA01(+)=C1.CP01 AND PA02(+)=C1.CP02 AND PA03(+)=C1.CP03 AND PA04(+)=C1.CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND C2.CP01(+)=C1.CP01 AND C2.CP02(+)=C1.CP02 AND C2.CP03(+)=C1.CP03 AND C2.CP04(+)=C1.CP04 AND C2.CP10 IN ('101','102','103','105','125') AND C2.CP159=0" & _
         " AND CPM01(+)=C2.CP01 AND CPM02(+)=C2.CP10 AND C1.CP27<=" & strCPM1933_Where
      'end 2010/10/14
      cnnConnection.Execute stSQL, intI
      
      'Modify By Sindy 2020/8/7 為配合日本部及專利國外部，於主管會議中經各組主管同意，
      '工程師實施無卷請款，為控管避免請款延遲太久，系統未請款彈跳期限通知，
      '工程師請款部分（參附件標黃色）請改為5個工作天
'      strCPM1933_Col = "TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),CPM19),'YYYYMMDD')"
'      strCPM1933_Where = "TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD')"
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),CPM19),'YYYYMMDD'),WORKDAYADD(cpm33+1,CP27))"
      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
      'Modified by Lydia 2016/09/14 CP20||CP57||CP60 IS NULL AND CP14 => CP27>20150914 AND CP14='...' AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,CP09,FA10,CP14,CP13" & _
         ",NULL,NULL," & strCPM1933_Col & " CP06,NULL,0" & _
         " FROM CASEPROGRESS,SERVICEPRACTICE,FAGENT,CASEPROPERTYMAP" & _
         " WHERE CP27>" & stDate1 & " AND CP14=" & stIdList(ii) & " AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'" & _
         " And CP09<'B' AND CP10 NOT IN ('403','404','411','418','419','901','902','907','908','913')" & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP27<=" & strCPM1933_Where
      cnnConnection.Execute stSQL, intI
   Next
   
   '非外專工程師承辦案件(部門非 F2,F5,F8 字頭的) : (N:[有關期限] 國外部專利處非外專承辦或未分案將到期管制人(含逾期))
   If bLvl4 = True Then
      '已收文未發文,2個工作天後達本所期限者(不含當日) --承辦人-A1(本所期限,承辦人)
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' AND CP05+0>20030000 => CP01='FCP' AND CP05+0>20030000 AND CP158=0 AND CP159=0
      'Modified by Lydia 2018/02/12  客戶提供文件1920=>'M','待處理'
      'stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
         " From CASEPROGRESS,PATENT,FAGENT" & _
         " WHERE CP01='FCP' AND CP05+0>20030000 AND CP158=0 AND CP159=0" & stConCP06 & _
         " AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "',DECODE(CP10,'1920','M','A') EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,PATENT,FAGENT" & _
         " WHERE CP01='FCP' AND CP05+0>20030000 AND CP158=0 AND CP159=0" & stConCP06 & _
         " AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
      cnnConnection.Execute stSQL, intI
      
      '已收文未發文,2個工作天後達承辦期限者(不含當日) --承辦人-B1(承辦期限,承辦人)
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' AND CP05+0>20030000 => CP01='FCP' AND CP05+0>20030000 AND CP158=0 AND CP159=0
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','B' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
         " From CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT" & _
         " WHERE CP01='FCP' AND CP05+0>20030000 AND CP158=0 AND CP159=0" & stConCP48 & _
         " AND EP02(+)=CP09 AND EP09 IS NULL" & _
         " AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
      cnnConnection.Execute stSQL, intI
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FG' AND CP05+0>20030000 => CP01='FG' AND CP05+0>20030000 AND CP158=0 AND CP159=0
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','B' EV1,'1' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
         " From CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT" & _
         " WHERE CP01='FG' AND CP05+0>20030000 AND CP158=0 AND CP159=0" & stConCP48 & _
         " AND EP02(+)=CP09 AND EP09 IS NULL" & _
         " AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
      cnnConnection.Execute stSQL, intI
      
      '所有未發文--承辦人-E(未發文)
      If idx = 1 Then
         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' AND CP05+0>20030000 => CP01='FCP' AND CP05+0>20030000 AND CP158=0 AND CP159=0
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS,PATENT,FAGENT" & _
            " WHERE CP01='FCP' AND CP05+0>20030000 AND CP158=0 AND CP159=0" & _
            " AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
         cnnConnection.Execute stSQL, intI
         
         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FG' AND CP05>20030000 AND CP14 IS NOT NULL => CP01='FG' AND CP05+0>20030000 AND CP14 IS NOT NULL AND CP158=0 AND CP159=0
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
            " WHERE CP01='FG' AND CP05+0>20030000 AND CP14 IS NOT NULL AND CP158=0 AND CP159=0" & _
            " AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
            " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
         cnnConnection.Execute stSQL, intI
      End If
      
      'Modify By Sindy 2020/8/7 為配合日本部及專利國外部，於主管會議中經各組主管同意，
      '工程師實施無卷請款，為控管避免請款延遲太久，系統未請款彈跳期限通知，
      '工程師請款部分（參附件標黃色）請改為5個工作天
'      strCPM1933_Col = "TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD')"
'      strCPM1933_Where = "TO_CHAR(ADD_MONTHS(SYSDATE,-1*DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD')"
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD'),WORKDAYADD(cpm33+1,CP27))"
      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
      '已發文未請款--承辦人
      'Modify by Morgan 2010/4/12 排除工程師提申940
      'Modified by Lydia 2016/09/14
      'stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,CP09,FA10,CP14,CP13" & _
         "," & strCPM1933_Col & " C06,NULL,NULL,NULL,0" & _
         " FROM CASEPROGRESS,PATENT,FAGENT,CASEPROPERTYMAP" & _
         " WHERE CP01||CP20||CP60='FCP' AND CP27>" & stDate1 & " AND CP16>0" & _
         " And CP09||''<'B' AND CP57||CP10 NOT IN ('403','404','411','418','419','901','902','907','908','913','940')" & _
         " AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP27<=" & strCPM1933_Where
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,CP09,FA10,CP14,CP13" & _
         ",NULL,NULL," & strCPM1933_Col & " C06,NULL,0" & _
         " FROM CASEPROGRESS,PATENT,FAGENT,CASEPROPERTYMAP" & _
         " WHERE CP01='FCP' AND CP27>" & stDate1 & " And CP09||''<'B' AND CP10 NOT IN ('403','404','411','418','419','901','902','907','908','913','940')" & _
         " AND NVL(CP20||CP60,'0')='0' AND CP16>0 AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP27<=" & strCPM1933_Where
      cnnConnection.Execute stSQL, intI
      
      'Modify By Sindy 2020/8/7 為配合日本部及專利國外部，於主管會議中經各組主管同意，
      '工程師實施無卷請款，為控管避免請款延遲太久，系統未請款彈跳期限通知，
      '工程師請款部分（參附件標黃色）請改為5個工作天
'      strCPM1933_Col = "TO_CHAR(ADD_MONTHS(TO_DATE(C2.CP27,'YYYYMMDD'),CPM19),'YYYYMMDD')"
'      strCPM1933_Where = "TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD')"
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(C2.CP27,'YYYYMMDD'),CPM19),'YYYYMMDD'),WORKDAYADD(cpm33+1,C2.CP27))"
      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
      'Add by Morgan 2010/4/12 工程師提申940要抓新申請案的設定
      'Modify by Morgan 2010/10/14 工程師提申940由新申請案的承辦人管制且帶新申請案的資料
      'Modified by Lydia 2016/09/14
      'stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,C2.CP09,FA10,C2.CP14,C2.CP13" & _
         "," & strCPM1933_Col & " C06,NULL,NULL,NULL,0" & _
         " FROM CASEPROGRESS C1,CASEPROGRESS C2,PATENT,FAGENT,CASEPROPERTYMAP" & _
         " WHERE C1.CP01||C1.CP20||C1.CP60='FCP' AND C1.CP27>" & stDate1 & " AND C1.CP16>0" & _
         " And C1.CP09||''<'B' AND C1.CP10='940'" & _
         " AND EXISTS(SELECT * FROM STAFF WHERE ST01=C1.CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND PA01(+)=C1.CP01 AND PA02(+)=C1.CP02 AND PA03(+)=C1.CP03 AND PA04(+)=C1.CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND C2.CP01(+)=C1.CP01 AND C2.CP02(+)=C1.CP02 AND C2.CP03(+)=C1.CP03 AND C2.CP04(+)=C1.CP04 AND C2.CP10 IN ('101','102','103','105','125') AND C2.CP57 IS NULL" & _
         " AND CPM01(+)=C2.CP01 AND CPM02(+)=C2.CP10 AND C1.CP27<=" & strCPM1933_Where
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,C2.CP09,FA10,C2.CP14,C2.CP13" & _
         ",NULL,NULL," & strCPM1933_Col & " C06,NULL,0" & _
         " FROM CASEPROGRESS C1,CASEPROGRESS C2,PATENT,FAGENT,CASEPROPERTYMAP" & _
         " WHERE C1.CP01='FCP' AND C1.CP27>" & stDate1 & " And C1.CP09||''<'B' AND C1.CP10='940'" & _
         " AND NVL(C1.CP20||C1.CP60,'0')='0' AND C1.CP16>0 AND EXISTS(SELECT * FROM STAFF WHERE ST01=C1.CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND PA01(+)=C1.CP01 AND PA02(+)=C1.CP02 AND PA03(+)=C1.CP03 AND PA04(+)=C1.CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND C2.CP01(+)=C1.CP01 AND C2.CP02(+)=C1.CP02 AND C2.CP03(+)=C1.CP03 AND C2.CP04(+)=C1.CP04 AND C2.CP10 IN ('101','102','103','105','125') AND C2.CP159=0" & _
         " AND CPM01(+)=C2.CP01 AND CPM02(+)=C2.CP10 AND C1.CP27<=" & strCPM1933_Where
      'end 2010/10/14
      cnnConnection.Execute stSQL, intI
      
      'Modify By Sindy 2020/8/7 為配合日本部及專利國外部，於主管會議中經各組主管同意，
      '工程師實施無卷請款，為控管避免請款延遲太久，系統未請款彈跳期限通知，
      '工程師請款部分（參附件標黃色）請改為5個工作天
'      strCPM1933_Col = "TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),CPM19),'YYYYMMDD')"
'      strCPM1933_Where = "TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD')"
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),CPM19),'YYYYMMDD'),WORKDAYADD(cpm33+1,CP27))"
      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
      'Modified by Lydia 2016/09/14
      'stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13" & _
         "," & strCPM1933_Col & " CP06,NULL,NULL,NULL,0" & _
         " FROM CASEPROGRESS,SERVICEPRACTICE,FAGENT,CASEPROPERTYMAP" & _
         " WHERE CP01||CP20||CP60='FG' AND CP27>" & stDate1 & " AND CP16>0 AND CP14 IS NOT NULL" & _
         " And CP09||''<'B' AND CP57||CP10 NOT IN ('403','404','411','418','419','901','902','907','908','913')" & _
         " AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP27<=" & strCPM1933_Where
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,' ' EV2,CP09,FA10,CP14,CP13" & _
         "," & strCPM1933_Col & " CP06,NULL,NULL,NULL,0" & _
         " FROM CASEPROGRESS,SERVICEPRACTICE,FAGENT,CASEPROPERTYMAP" & _
         " WHERE CP01='FG' AND CP27>" & stDate1 & " AND CP16>0 AND CP14 IS NOT NULL And CP09||''<'B'" & _
         " AND CP10 NOT IN ('403','404','411','418','419','901','902','907','908','913')" & _
         " AND NVL(CP20||CP60,'0')='0' AND CP16>0 AND EXISTS(SELECT * FROM STAFF WHERE ST01=CP14 AND SUBSTRB(ST03,1,2)<>'F2' AND SUBSTRB(ST03,1,2)<>'F5' AND SUBSTRB(ST03,1,2)<>'F8')" & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP27<=" & strCPM1933_Where
      cnnConnection.Execute stSQL, intI
   End If
      
   '未分案-0
   'Modify by Morgan 2009/9/7 未分案改分組管制
   'If bLvl4 = True Or bLvl5 = True Then
   'Modified by Lydia 2022/12/20 +bLvO1
   '工程師
   If bLvlChm Or bLvlJpn Or bLvlMot Or bLvlEls Or bLvl5 Or bLvO1 Then
      stConPAGrp = "": stConSPGrp = ""
      'Modify By Sindy 2021/4/22 排除案件性質是412延緩公告417提早公開
      'Modified by Lydia 2022/12/20 國外部期限通知未分案管制人
      'If bLvlChm Then
      If bLvO1 = True Or bLvl5 = True Then
         If bLvlChm Or bLvlJpn Or bLvlMot Or bLvlEls Then
             stConPAGrp = " and (PA150='" & PUB_GetStaffST16(stUserID) & "' or PA150 IS NULL) and cp10 not in('412','417')"
             stConSPGrp = " and (SP79='" & PUB_GetStaffST16(stUserID) & "' or SP79 is null) and cp10 not in('412','417')"
         Else
             stConPAGrp = " and PA150 IS NULL and cp10 not in('412','417')"
             stConSPGrp = " and SP79 is null and cp10 not in('412','417')"
         End If
      ElseIf bLvlChm Then
      'end 2022/12/20
         stConPAGrp = " and PA150='2' and cp10 not in('412','417')"
         stConSPGrp = " and SP79='2' and cp10 not in('412','417')"
      ElseIf bLvlJpn Then
         'Modified by Lydia 2021/02/25 未分組也未分案給王協理看
         'stConPAGrp = " and PA150='3'"
         'stConSPGrp = " and SP79='3'"
         'Modified by Lydia 2022/12/20 改回日文組
         'stConPAGrp = " and (PA150='3' or PA150 IS NULL) and cp10 not in('412','417')"
         'stConSPGrp = " and (SP79='3' or SP79 is null) and cp10 not in('412','417')"
         'end 2021/02/25
         stConPAGrp = " and PA150='3' and cp10 not in('412','417')"
         stConSPGrp = " and SP79='3' and cp10 not in('412','417')"
      '2011/11/30 modify by sonia 取消德文且機電拆電子電機及機械設計
      'ElseIf bLvlEls Then
      '   stConPAGrp = " and ((PA150<>'2' AND PA150<>'3') OR PA150 IS NULL)"
      '   stConSPGrp = " and ((SP79<>'2' and SP79<>'3') or SP79 is null)"
      ElseIf bLvlMot Then
         stConPAGrp = " and PA150='4' and cp10 not in('412','417')"
         stConSPGrp = " and SP79='4' and cp10 not in('412','417')"
      ElseIf bLvlEls Then
         'Modified by Lydia 2021/02/25 最早以前外專工程師主管是阮威立85030為電子電機組現在改成日文組
         'stConPAGrp = " and ((PA150>='2' AND PA150<='4') OR PA150 IS NULL)"
         'stConSPGrp = " and ((SP79>='2' and SP79<='4') or SP79 is null)"
         stConPAGrp = " and PA150='1' and cp10 not in('412','417')"
         stConSPGrp = " and SP79='1' and cp10 not in('412','417')"
      '2011/11/30 end
      End If
      
      '已完稿無核稿人,2個工作天後達核稿期限者(不含當日)-C8(核稿期限,無核稿人)
      'Add By Sindy 2016/8/1 Sharon:新案翻譯,達核槁未Key核稿人,請改為彈工程師主管
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' AND CP10='201' => CP01='FCP' AND CP10='201' AND CP158=0 AND CP159=0
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','C' EV1,'8' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT" & _
         " WHERE CP01='FCP' AND CP10='201' AND CP158=0 AND CP159=0 AND CP05>" & stDate1 & _
         " AND EP02(+)=CP09 AND EP04 IS NULL AND EP09>0 AND EP08>=" & stDate1 & " AND EP08<" & stDate2 & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & stConPAGrp
      cnnConnection.Execute stSQL, intI
      
      'Add By Sindy 2016/8/1 Sharon:新案翻譯,達核槁未Key核稿人,請改為彈工程師主管
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FG' AND CP10='201' => CP01='FG' AND CP10='201' AND CP158=0 AND CP159=0
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','C' EV1,'8' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT" & _
         " WHERE CP01='FG' AND CP10='201' AND CP158=0 AND CP159=0 AND CP05>" & stDate1 & _
         " AND EP02(+)=CP09 AND EP04 IS NULL AND EP09>0 AND EP08>=" & stDate1 & " AND EP08<" & stDate2 & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)" & stConSPGrp
      cnnConnection.Execute stSQL, intI
      
      '已收文未發,2個工作天後達本所期限者(不含當日)-A0(本所期限,未分案)
      'Modified by Lydia 2016/09/14  CP01||CP14||CP27||CP57='FCP' => CP01||CP14='FCP' AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
      'Modified by Lydia 2018/02/12 客戶提供文件1920 =>'M','待處理'
      'stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/ '" & strUserNum & "','A' EV1,'0' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
         " From CASEPROGRESS,PATENT,FAGENT" & _
         " WHERE CP01||CP14='FCP' AND CP158=0 AND CP159=0" & stConCP06 & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL" & stConPAGrp & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
      'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "',DECODE(CP10,'1920','M','A') EV1,'0' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,PATENT,FAGENT" & _
         " WHERE substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & stConCP06 & _
         " AND PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA57||PA108 IS NULL" & stConPAGrp & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
      cnnConnection.Execute stSQL, intI
      'Add By Sindy 2021/4/21 達指定
      'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','N' EV1,'0' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,PATENT,FAGENT" & _
         " WHERE substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & stConCP142 & _
         " AND PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA57||PA108 IS NULL" & stConPAGrp & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
      cnnConnection.Execute stSQL, intI
      
      'Modified by Lydia 2016/09/14 CP01||CP14||CP27||CP57='FG' =>  CP01||CP14='FG' AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
      'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS, SERVICEPRACTICE, FAGENT" & _
         " WHERE CP05>20030000 AND substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & stConCP06 & _
         " AND SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04 AND SP15||SP61 IS NULL" & stConSPGrp & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
      cnnConnection.Execute stSQL, intI
      'Add By Sindy 2021/4/21 達指定
      'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','N' EV1,'0' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS, SERVICEPRACTICE, FAGENT" & _
         " WHERE CP05>20030000 AND substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & stConCP142 & _
         " AND SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04 AND SP15||SP61 IS NULL" & stConSPGrp & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
      cnnConnection.Execute stSQL, intI
      
      '已收文未發文,2個工作天後達承辦期限者(不含當日)-B0(承辦期限,未分案)
      'Modified by Lydia 2016/09/14  CP01||CP14||CP27||CP57='FCP' => CP01||CP14='FCP' AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
      'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','B' EV1,'0' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
         " From CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT" & _
         " WHERE substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & stConCP48 & _
         " AND EP02(+)=CP09 AND EP09 IS NULL" & _
         " AND PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA57||PA108 IS NULL" & stConPAGrp & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
      cnnConnection.Execute stSQL, intI
      
      'Modified by Lydia 2016/09/14 CP01||CP14||CP27||CP57='FG' =>  CP01||CP14='FG' AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
      'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','B' EV1,'0' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
         " From CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
         " WHERE substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & stConCP48 & _
         " AND SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04 AND SP15||SP61 IS NULL" & stConSPGrp & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
      cnnConnection.Execute stSQL, intI
      
      '未分案,所有未發文-E(未發文)
      If idx = 1 Then
         'Modified by Lydia 2016/09/14  CP01||CP14||CP27||CP57='FCP' => CP01||CP14='FCP' AND CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
          'Modified by Lydia 2018/02/12  客戶提供文件1920=>'M','待處理'
         'stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/ '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS,PATENT,FAGENT" & _
            " WHERE CP05>20030000 AND CP01||CP14='FCP' AND CP158=0 AND CP159=0" & _
            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL" & stConPAGrp & _
            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
         'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "',DECODE(CP10,'1920','M','E') EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS,PATENT,FAGENT" & _
            " WHERE CP05>20030000 AND substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & _
            " AND PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA57||PA108 IS NULL" & stConPAGrp & _
            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
         cnnConnection.Execute stSQL, intI
         
         'Modified by Lydia 2016/09/14 CP01||CP14||CP27||CP57='FG' =>  CP01||CP14='FG' AND CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
         'Modify By Sindy 2023/8/31 CP01||CP14='FCP' => substr(cp12,1,1)||CP14='F' 含FMP案也要彈出來
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
            " WHERE substr(cp12,1,1)||CP14='F' AND CP158=0 AND CP159=0" & _
            " AND SP01=CP01 AND SP02=CP02 AND SP03=CP03 AND SP04=CP04 AND SP15||SP61 IS NULL" & stConSPGrp & _
            " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
         cnnConnection.Execute stSQL, intI
      End If
   End If
      
   If bLvlX = True Then
      'Modify by Morgan 2008/10/9 未交稿加判斷無核稿期限的(會有例外狀況需核完稿才給翻譯費故完稿日會先拿掉,如巨京)
      
      '未交稿,2個工作天後達承辦期限者(不含當日)-G9(未交稿)
      'Modify by Morgan 2008/10/21 改當日到期的
      'Modify By Sindy 2016/7/28 + 有關翻譯之未交稿期限(下班翻改彈所內編號,外翻仍維持彈系統特殊人員"M") : AND not exists (select ts1.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')
      'Modified by Lydia 2016/09/22 CP10||CP27||CP57='201' => CP10='201' AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
      'Modify By Sindy 2021/4/15 cp14==>decode(SIM01,null,CP14,SIM01)
      '   所內翻譯也要列出來 mark:AND not exists (select ts1.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','G' EV1,'9' EV2,CP09,FA10,decode(SIM01,null,CP14,SIM01),CP13,CP06,CP07,CP48,NULL,0" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT,STAFF_IDMAP,STAFF S1,STAFF S2" & _
         " WHERE CP10='201' AND CP158=0 AND CP159=0 and substr(CP12,1,1)='F' AND CP48>" & stDate1 & " AND CP48<=" & strSrvDate(1) & stCP01 & _
         " AND EP02(+)=CP09 AND EP09 IS NULL AND EP08 IS NULL" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND S1.ST01(+)=CP14 AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01" & _
         " AND SUBSTRB(S1.ST15,1,1)='F' AND (S2.ST15 IS NULL OR SUBSTRB(S2.ST15,1,1)='F')" '& _
         " AND not exists (select ts1.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')"
      cnnConnection.Execute stSQL, intI
      
      'Modify By Sindy 2016/7/28 + 有關翻譯之未交稿期限(下班翻改彈所內編號,外翻仍維持彈系統特殊人員"M")
      'Modified by Lydia 2016/09/22 CP10||CP27||CP57='201' => CP10='201' AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
      'Modify By Sindy 2021/4/15 cp14==>decode(SIM01,null,CP14,SIM01)
      '   所內翻譯也要列出來 mark:AND not exists (select ts1.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','G' EV1,'9' EV2,CP09,FA10,decode(SIM01,null,CP14,SIM01),CP13,CP06,CP07,CP48,NULL,0" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT,STAFF_IDMAP,STAFF S1,STAFF S2" & _
         " WHERE CP10='201' AND CP158=0 AND CP159=0 and substr(CP12,1,1)='F' AND CP48>" & stDate1 & " AND CP48<=" & strSrvDate(1) & stCP01 & _
         " AND EP02(+)=CP09 AND EP09 IS NULL AND EP08 IS NULL" & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)" & _
         " AND S1.ST01(+)=CP14 AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01" & _
         " AND SUBSTRB(S1.ST15,1,1)='F' AND (S2.ST15 IS NULL OR SUBSTRB(S2.ST15,1,1)='F')" '& _
         " AND not exists (select ts1.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')"
      cnnConnection.Execute stSQL, intI
      
      'Modify By Sindy 2021/4/22 O.未分案,案件性質是412延緩公告417提早公開(沒有本所期限,法定期限,承辦期限)
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','O' EV1,'0' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,PATENT,FAGENT" & _
         " WHERE CP01||CP14='FCP' AND CP158=0 AND CP159=0" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) and cp10 in('412','417')"
      cnnConnection.Execute stSQL, intI
      
      '未交稿,所有未發文-E(只抓有承辦期限的)
      If idx = 1 Then
         'Modified by Lydia 2016/09/22 CP10||CP27||CP57='201' => CP10='201' AND CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
         'Modify By Sindy 2021/4/15 cp14==>decode(SIM01,null,CP14,SIM01)
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,decode(SIM01,null,CP14,SIM01),CP13,CP06,CP07,CP48,NULL,0" & _
            " FROM CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT,STAFF_IDMAP,STAFF S1,STAFF S2" & _
            " WHERE CP10='201' AND CP158=0 AND CP159=0 and substr(CP12,1,1)='F' AND CP48>" & stDate1 & " AND CP48>=" & stDate2 & stCP01 & _
            " AND EP02(+)=CP09 AND EP09 IS NULL AND EP08 IS NULL" & _
            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
            " AND S1.ST01(+)=CP14 AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01" & _
            " AND SUBSTRB(S1.ST15,1,1)='F' AND (S2.ST15 IS NULL OR SUBSTRB(S2.ST15,1,1)='F')"
         cnnConnection.Execute stSQL, intI
         
         'Modified by Lydia 2016/09/22 CP10||CP27||CP57='201' => CP10='201' AND CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
         'Modify By Sindy 2021/4/15 cp14==>decode(SIM01,null,CP14,SIM01)
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,decode(SIM01,null,CP14,SIM01),CP13,CP06,CP07,CP48,NULL,0" & _
            " FROM CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT,STAFF_IDMAP,STAFF S1,STAFF S2" & _
            " WHERE CP10='201' AND CP158=0 AND CP159=0 and substr(CP12,1,1)='F' AND CP48>" & stDate1 & " AND CP48>=" & stDate2 & stCP01 & _
            " AND EP02(+)=CP09 AND EP09 IS NULL AND EP08 IS NULL" & _
            " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)" & _
            " AND S1.ST01(+)=CP14 AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01" & _
            " AND SUBSTRB(S1.ST15,1,1)='F' AND (S2.ST15 IS NULL OR SUBSTRB(S2.ST15,1,1)='F')"
         cnnConnection.Execute stSQL, intI
      End If
      
'      '已完稿無核稿人,2個工作天後達核稿期限者(不含當日)-C8(核稿期限,無核稿人)
'      'Modify By Sindy 2016/8/1 Sharon:新案翻譯,達核槁未Key核稿人,請改為彈工程師主管
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','C' EV1,'8' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
'         " FROM CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT" & _
'         " WHERE CP01||CP27||CP57='FCP' AND CP10='201' AND CP05>" & stDate1 & _
'         " AND EP02(+)=CP09 AND EP04 IS NULL AND EP09>0 AND EP08>=" & stDate1 & " AND EP08<" & stDate2 & _
'         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
'      cnnConnection.Execute stSQL, intI
'
'      'Modify By Sindy 2016/8/1 Sharon:新案翻譯,達核槁未Key核稿人,請改為彈工程師主管
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','C' EV1,'8' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
'         " FROM CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT" & _
'         " WHERE CP01||CP27||CP57='FG' AND CP10='201' AND CP05>" & stDate1 & _
'         " AND EP02(+)=CP09 AND EP04 IS NULL AND EP09>0 AND EP08>=" & stDate1 & " AND EP08<" & stDate2 & _
'         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL" & _
'         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
'      cnnConnection.Execute stSQL, intI
      
      '已完稿無核稿人,所有未發文-E(只抓有核稿期限)
      If idx = 1 Then
         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' AND CP10='201' => CP01='FCP' AND CP10='201' AND CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,NULL,0" & _
            " FROM CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT" & _
            " WHERE CP01='FCP' AND CP10='201' AND CP158=0 AND CP159=0 AND CP05>" & stDate1 & _
            " AND EP02(+)=CP09 AND EP04 IS NULL AND EP09>0 AND EP08>=" & stDate2 & _
            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
         cnnConnection.Execute stSQL, intI
         
         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FG' AND CP10='201' => CP01='FG' AND CP10='201' AND CP158=0 AND CP159=0
         'Modified by Lydia 2016/09/22 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,NULL,0" & _
            " FROM CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT" & _
            " WHERE CP01='FG' AND CP10='201' AND CP158=0 AND CP159=0 AND CP05>" & stDate1 & _
            " AND EP02(+)=CP09 AND EP04 IS NULL AND EP09>0 AND EP08>=" & stDate2 & _
            " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL" & _
            " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
         cnnConnection.Execute stSQL, intI
      End If
   End If
               
   '程序
   If stDept = "F22" Then
      'Modify By Sindy 2015/11/19 Mark
'      'Added by Morgan 2014/8/27
'      '新案翻譯已發文未請款--管制人
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','F' EV1,'1' EV2,CP09,FA10,CP14,CP13" & _
'         ",TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD') C06,NULL,NULL,NULL,0" & _
'         " FROM CASEPROGRESS,PATENT,FAGENT,CASEPROPERTYMAP,Nation" & _
'         " WHERE CP27>" & stDate1 & "-10000 AND CP20||CP57||CP60 IS NULL AND SUBSTR(CP12,1,2)='F2'" & _
'         " And CP09<'B' AND CP10='201'" & _
'         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
'         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP27<=TO_CHAR(ADD_MONTHS(SYSDATE,-1*DECODE(INSTR('2101,2102',CP10||PA08),0,CPM19,4)),'YYYYMMDD')" & _
'         " AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'      cnnConnection.Execute stSQL, intI
'      'end 2014/8/27
   
      '已收文未發文且 2個工作天 後達本所期限者(不含當日) --管制人-A2(本所期限,管制人)
      'Modify by Morgan 2008/11/19 +年費另外
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10 not in('605','926','945') => CP01='FCP' and CP10 not in('605','926','945') AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/23 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
      'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人) ('605','926','945')-->('926','945')
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,PATENT,FAGENT,Nation Where CP01='FCP' and CP10 not in('926','945') AND CP158=0 AND CP159=0" & stConCP06 & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
      cnnConnection.Execute stSQL, intI
      'Add By Sindy 2021/4/21 達指定
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','N' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,PATENT,FAGENT,Nation Where CP01='FCP' and CP10 not in('926','945') AND CP158=0 AND CP159=0" & stConCP142 & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
      cnnConnection.Execute stSQL, intI
      
      'Add By Sindy 2021/3/17 + 核對已准專利已發文未請款->請彈程序
      strCPM1933_Col = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),CPM19),'YYYYMMDD'),WORKDAYADD(cpm33+1,CP27))"
      strCPM1933_Where = "decode(cpm33,null,TO_CHAR(ADD_MONTHS(SYSDATE,-1*CPM19),'YYYYMMDD'),WORKDAYADD(-1*(cpm33+1),to_char(SYSDATE,'YYYYMMDD')))"
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'2' EV2,CP09,FA10,CP14,CP13" & _
         ",NULL,NULL," & strCPM1933_Col & " C06,NULL,0" & _
         " FROM CASEPROGRESS,PATENT,FAGENT,Nation,CASEPROPERTYMAP" & _
         " WHERE CP27>" & stDate1 & " AND CP159=0 AND CP16>0 AND NVL(CP20||CP60,'0')='0'" & _
         " And CP01='FCP' And CP10='926'" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16 & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP27<=" & strCPM1933_Where
      cnnConnection.Execute stSQL, intI
      '2021/3/17 END
      
      'Add By Sindy 2015/10/22 +法務:達本所,僅限收文業務區CP12=F2字頭
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57 in('FCL','CFL','LIN') => CP01 in('FCL','CFL','LIN') AND CP158=0 AND CP159=0
      'Modified by Lydia 2016/09/23 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,nvl(FA10,cu10),CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,Lawcase,FAGENT,Nation n1,Customer,Nation n2 Where CP01 in('FCL','CFL','LIN') AND CP158=0 AND CP159=0" & stConCP06 & _
         " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC08||LC34 IS NULL" & _
         " AND FA01(+)=SUBSTR(LC22,1,8) AND FA02(+)=SUBSTR(LC22,9) AND n1.NA01(+)=FA10" & _
         " AND CU01(+)=SUBSTR(LC11,1,8) AND CU02(+)=SUBSTR(LC11,9) AND n2.NA01(+)=CU10" & _
         " AND (CP14 IS NULL OR CP14<>n1.NA16 OR CP14<>n2.NA16)" & stConNA16L & _
         " AND substr(CP12,1,2)='F2'"
      cnnConnection.Execute stSQL, intI
      '2015/10/22 END
      
'Removed by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人)
'
'      'Add by Morgan 2008/11/19
'      '個案年費代理人
'      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10='605' => CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0
'      'Modified by Lydia 2016/09/23 + /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ '" & strUserNum & "','A' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
'         " From CASEPROGRESS, PATENT, FAGENT, Nation Where CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0" & stConCP06 & _
'         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA76,1,8) AND FA02(+)=SUBSTR(PA76,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'      'Add by Morgan 2009/4/29 年費代理人也會有X編號
'      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10='605' => CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,NA01,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
'         " From CASEPROGRESS, PATENT, CUSTOMER, Nation Where CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0" & stConCP06 & _
'         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND CU01(+)=SUBSTR(PA76,1,8) AND CU02(+)=SUBSTR(PA76,9) AND NA01(+)=CU10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'      '申請人年費代理人
'      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10='605' => CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
'         " From CASEPROGRESS,PATENT,CUSTOMER,FAGENT,Nation Where CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0" & stConCP06 & _
'         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(CU96,1,8) AND FA02(+)=SUBSTR(CU96,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'      'Add by Morgan 2009/4/29 年費代理人也會有X編號
'      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10='605' => CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,NA01,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
'         " From CASEPROGRESS,PATENT,CUSTOMER c1,CUSTOMER c2,Nation Where CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0" & stConCP06 & _
'         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND c1.CU01(+)=SUBSTR(PA26,1,8) AND c1.CU02(+)=SUBSTR(PA26,9) AND c1.CU96 IS NOT NULL" & _
'         " AND c2.CU01(+)=SUBSTR(c1.CU96,1,8) AND c2.CU02(+)=SUBSTR(c1.CU96,9) AND NA01(+)=c2.CU10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'      '代理人
'      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10='605' => CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
'         " From CASEPROGRESS,PATENT,CUSTOMER,FAGENT,Nation Where CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0" & stConCP06 & _
'         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'end 2020/5/12
      
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FG' => CP01='FG' AND CP158=0 AND CP159=0
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS, SERVICEPRACTICE, FAGENT, Nation Where CP01='FG' AND CP158=0 AND CP159=0" & stConCP06 & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP26 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
      cnnConnection.Execute stSQL, intI
      'Add By Sindy 2021/4/21 達指定
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','N' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS, SERVICEPRACTICE, FAGENT, Nation Where CP01='FG' AND CP158=0 AND CP159=0" & stConCP142 & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP26 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
      cnnConnection.Execute stSQL, intI
      
'Removed by Morgan 2012/6/15 併入原FMP案約定期限通知
'      'Added by Morgan 2012/5/28 +FMP
'      '達本所未完稿
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
'         " From CASEPROGRESS,PATENT,FAGENT,Nation,STAFF,engineerprogress Where (CP01='P' OR CP01='CFP') AND CP57||CP27 IS NULL" & _
'         " AND SUBSTR(CP12,1,1)='F' AND ST01(+)=CP14 AND SUBSTR(ST03,1,1)='F' and ep02(+)=cp09 and ep09 is null" & stConCP06 & _
'         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'      '達本所已完稿未核稿
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,NULL,0" & _
'         " From CASEPROGRESS,PATENT,FAGENT,Nation,STAFF,engineerprogress Where (CP01='P' OR CP01='CFP') AND CP57||CP27 IS NULL" & _
'         " AND SUBSTR(CP12,1,1)='F' and cp10='201' AND ST01(+)=EP04 AND SUBSTR(ST03,1,1)='F' and ep02(+)=cp09 and ep09>0 and ep33 is null" & stConCP06 & _
'         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
'         " From CASEPROGRESS, SERVICEPRACTICE, FAGENT, Nation,STAFF,engineerprogress Where (CP01='PS' OR CP01='CPS') AND CP57||CP27 IS NULL" & _
'         " AND SUBSTR(CP12,1,1)='F' AND ST01(+)=CP14 AND SUBSTR(ST03,1,1)='F' and ep02(+)=cp09 and ep09 is null" & stConCP06 & _
'         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP26 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,NULL,0" & _
'         " From CASEPROGRESS, SERVICEPRACTICE, FAGENT, Nation,STAFF,engineerprogress Where (CP01='PS' OR CP01='CPS') AND CP57||CP27 IS NULL" & _
'         " AND SUBSTR(CP12,1,1)='F' and cp10='201' AND ST01(+)=EP04 AND SUBSTR(ST03,1,1)='F' and ep02(+)=cp09 and ep09>0 and ep33 is null" & stConCP06 & _
'         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP26 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'      cnnConnection.Execute stSQL, intI
'      'end 2012/5/28
'end 2012/6/15
            
'Removed by Morgan 2012/5/28 目前不要管制達承辦但保留程式備用
'      'Added by Morgan 2012/5/28
'      'FMP案
'      '已收文未發文,2個工作天後達承辦期限者(不含當日)-B2(承辦期限,管制人)
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','B' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
'         " From CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT,NATION,STAFF" & _
'         " WHERE (CP01='P' OR CP01='CFP') AND CP57||CP27 IS NULL AND SUBSTR(CP12,1,1)='F' AND ST01(+)=CP14 AND SUBSTR(ST03,1,1)='F'" & stConCP48 & _
'         " AND EP02(+)=CP09 AND EP09 IS NULL" & _
'         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','B' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
'         " From CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT,NATION,STAFF" & _
'         " WHERE (CP01='PS' OR CP01='CPS') AND CP57||CP27 IS NULL AND SUBSTR(CP12,1,1)='F' AND ST01(+)=CP14 AND SUBSTR(ST03,1,1)='F'" & stConCP48 & _
'         " AND EP02(+)=CP09 AND EP09 IS NULL" & _
'         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'      '未核稿且 2個工作天 後達核稿期限者(不含當日) -C2(核稿期限,管制人)
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','C' EV1,'2' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
'         " FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,FAGENT,NATION,STAFF" & _
'         " WHERE EP33 IS NULL" & stConEP08 & _
'         " AND CP09(+)=EP02 AND CP27||CP57 IS NULL AND (CP01='P' OR CP01='CFP') AND CP10='201' AND SUBSTR(CP12,1,1)='F' AND ST01(+)=EP04 AND SUBSTR(ST03,1,1)='F'" & _
'         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','C' EV1,'2' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
'         " FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,FAGENT,NATION,STAFF" & _
'         " WHERE EP33 IS NULL" & stConEP08 & _
'         " AND CP09(+)=EP02 AND CP27||CP57 IS NULL AND (CP01='PS' OR CP01='CPS') AND CP10='201' AND SUBSTR(CP12,1,1)='F' AND ST01(+)=EP04 AND SUBSTR(ST03,1,1)='F'" & _
'         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'      cnnConnection.Execute stSQL, intI
'      'end 2012/5/28
      
      '所有未發文--管制人-E(未發文)
      If idx = 1 Then
      
         'Modify by Morgan 2008/11/19 +年費管制人另外有規則
         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10 not in('605','926','945') => CP01='FCP' and CP10 not in('605','926','945') AND CP158=0 AND CP159=0
         'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人) ('605','926','945')-->('926','945')
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS, PATENT, FAGENT, Nation Where CP01='FCP' and CP10 not in('926','945') AND CP158=0 AND CP159=0" & _
            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
         'Added by Lydia 2018/02/12 排除重複項目
         stSQL = stSQL & " AND CP09 NOT IN (SELECT R04 FROM R060204 WHERE R01='" & strUserNum & "' AND R02='E' AND R03 = ' ' AND R12=0) "
         cnnConnection.Execute stSQL, intI
         
         'Add By Sindy 2015/10/22 +法務:未發文,僅限收文業務區CP12=F2字頭
         'Modified by Lydia 2016/09/14 CP01||CP27||CP57 in('FCL','CFL','LIN') => CP01 in('FCL','CFL','LIN') AND CP158=0 AND CP159=0
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,nvl(FA10,cu10),CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS, Lawcase, FAGENT, Nation n1,Customer, Nation n2 Where CP01 in('FCL','CFL','LIN') AND CP158=0 AND CP159=0" & _
            " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC08||LC34 IS NULL" & _
            " AND FA01(+)=SUBSTR(LC22,1,8) AND FA02(+)=SUBSTR(LC22,9) AND n1.NA01(+)=FA10" & _
            " AND CU01(+)=SUBSTR(LC11,1,8) AND CU02(+)=SUBSTR(LC11,9) AND n2.NA01(+)=CU10" & _
            " AND (CP14 IS NULL OR CP14<>n1.NA16 OR CP14<>n2.NA16)" & stConNA16L & _
            " AND substr(CP12,1,2)='F2'"
         'Added by Lydia 2018/02/12 排除重複項目
         stSQL = stSQL & " AND CP09 NOT IN (SELECT R04 FROM R060204 WHERE R01='" & strUserNum & "' AND R02='E' AND R03 = ' ' AND R12=0) "
         cnnConnection.Execute stSQL, intI
         '2015/10/22 END
         
'Removed by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人)
'
'         'Add by Morgan 2008/11/19
'         '個案年費代理人
'         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10='605' => CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0
'         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
'            " From CASEPROGRESS,PATENT,FAGENT,Nation Where CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0" & _
'            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'            " AND FA01(+)=SUBSTR(PA76,1,8) AND FA02(+)=SUBSTR(PA76,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'         'Added by Lydia 2018/02/12 排除重複項目
'         stSQL = stSQL & " AND CP09 NOT IN (SELECT R04 FROM R060204 WHERE R01='" & strUserNum & "' AND R02='E' AND R03 = ' ' AND R12=0) "
'         cnnConnection.Execute stSQL, intI
'
'         'Add by Morgan 2009/4/29 年費代理人也會有X編號
'         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10='605' => CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0
'         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,NA01,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
'            " From CASEPROGRESS,PATENT,CUSTOMER,Nation Where CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0" & _
'            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'            " AND CU01(+)=SUBSTR(PA76,1,8) AND CU02(+)=SUBSTR(PA76,9) AND NA01(+)=CU10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'         'Added by Lydia 2018/02/12 排除重複項目
'         stSQL = stSQL & " AND CP09 NOT IN (SELECT R04 FROM R060204 WHERE R01='" & strUserNum & "' AND R02='E' AND R03 = ' ' AND R12=0) "
'         cnnConnection.Execute stSQL, intI
'
'         '申請人年費代理人
'         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10='605' => CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0
'         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
'            " From CASEPROGRESS,PATENT,CUSTOMER,FAGENT,Nation Where CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0" & _
'            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'            " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NOT NULL" & _
'            " AND FA01(+)=SUBSTR(CU96,1,8) AND FA02(+)=SUBSTR(CU96,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'         'Added by Lydia 2018/02/12 排除重複項目
'         stSQL = stSQL & " AND CP09 NOT IN (SELECT R04 FROM R060204 WHERE R01='" & strUserNum & "' AND R02='E' AND R03 = ' ' AND R12=0) "
'         cnnConnection.Execute stSQL, intI
'
'         'Add by Morgan 2009/4/29 年費代理人也會有X編號
'         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10='605' => CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0
'         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,NA01,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
'            " From CASEPROGRESS,PATENT,CUSTOMER c1,CUSTOMER c2,Nation Where CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0" & _
'            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'            " AND c1.CU01(+)=SUBSTR(PA26,1,8) AND c1.CU02(+)=SUBSTR(PA26,9) AND c1.CU96 IS NOT NULL" & _
'            " AND c2.CU01(+)=SUBSTR(c1.CU96,1,8) AND c2.CU02(+)=SUBSTR(c1.CU96,9) AND NA01(+)=c2.CU10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'         'Added by Lydia 2018/02/12 排除重複項目
'         stSQL = stSQL & " AND CP09 NOT IN (SELECT R04 FROM R060204 WHERE R01='" & strUserNum & "' AND R02='E' AND R03 = ' ' AND R12=0) "
'         cnnConnection.Execute stSQL, intI
'
'         '代理人
'         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10='605' => CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0
'         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
'            " From CASEPROGRESS,PATENT,CUSTOMER,FAGENT,Nation Where CP01='FCP' and CP10='605' AND CP158=0 AND CP159=0" & _
'            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND PA76 IS NULL" & _
'            " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NULL" & _
'            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
'         'Added by Lydia 2018/02/12 排除重複項目
'         stSQL = stSQL & " AND CP09 NOT IN (SELECT R04 FROM R060204 WHERE R01='" & strUserNum & "' AND R02='E' AND R03 = ' ' AND R12=0) "
'         cnnConnection.Execute stSQL, intI
'         'end 2008/11/19
'
'end 2020/5/12
         
         'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FG' => CP01='FG' AND CP158=0 AND CP159=0
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,' ' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
            " From CASEPROGRESS, SERVICEPRACTICE, FAGENT, Nation Where CP01='FG' AND CP158=0 AND CP159=0" & _
            " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP26 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10 AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16
         'Added by Lydia 2018/02/12 排除重複項目
         stSQL = stSQL & " AND CP09 NOT IN (SELECT R04 FROM R060204 WHERE R01='" & strUserNum & "' AND R02='E' AND R03 = ' ' AND R12=0) "
         cnnConnection.Execute stSQL, intI
      End If
   
      '未收文且 2個工作天 後達本所期限者(不含當日) --管制人-D2(未收文,管制人)
      'Modify by Morgan 2008/11/19 +年費另外
      'Modify by Morgan 2009/7/13 +995,996
      'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
      'Modified by Lydia 2014/11/14 NA51 = decode(pa75,'Y51333010','" & midStr & "',na51)
      'Modified by Lydia 2016/02/03 +Y51817040
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807) */
      'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人) ('605','926','945')-->('926','945')
      'Modify By Sindy 2021/4/28 + ,R15
      'Modified by Lydia 2022/11/03 區分FMP案; stConNA16=> Replace(UCase(stConNA16), "CP01", "NP02")
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,FA10,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS, PATENT, FAGENT, Nation" & _
         " WHERE NP02||NP06='FCP' and np07 not in('926','945')" & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate2 & _
         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & Replace(UCase(stConNA16), "CP01", "NP02") & _
         " AND NOT EXISTS(SELECT * FROM CASEPROGRESS WHERE CP09=NP01 AND CP10='928' AND NP07='202')"
      cnnConnection.Execute stSQL, intI
      
      'Add By Sindy 2015/10/22 +法務:程序組的未收文,以管制人角度顯示,僅限下一程序智權人員之ST15=F2字頭
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807) */
      'Modify By Sindy 2021/4/28 + ,R15
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,nvl(FA10,cu10),NULL,decode(FA10,null,n2.na51,decode(LC22," & midStr & ",n1.na51)) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS, Lawcase, FAGENT, Nation n1, Staff, Customer, Nation n2" & _
         " WHERE NP02||NP06 in('FCL','CFL','LIN')" & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate2 & _
         " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05 AND LC08||LC34 IS NULL" & _
         " AND FA01(+)=SUBSTR(LC22,1,8) AND FA02(+)=SUBSTR(LC22,9) AND n1.NA01(+)=FA10" & _
         " AND CU01(+)=SUBSTR(LC11,1,8) AND CU02(+)=SUBSTR(LC11,9) AND n2.NA01(+)=cu10" & stConNA16L & _
         " AND NP10=ST01(+) AND substr(ST15,1,2)='F2'"
      cnnConnection.Execute stSQL, intI
      '2015/10/22 END
      
'Removed by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人)
'
'      'Add by Morgan 2008/11/19
'      '個案年費代理人
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,FA10,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS, PATENT, FAGENT, Nation" & _
'         " WHERE NP02||NP06='FCP' and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA76,1,8) AND FA02(+)=SUBSTR(PA76,9) AND NA01(+)=FA10" & stConNA16
'      stSQL = Replace(stSQL, "decode(pa75,", "decode(pa76,")
'      cnnConnection.Execute stSQL, intI
'
'      'Add by Morgan 2009/4/29 年費代理人也會有X編號
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,NA01,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS, PATENT, Customer, Nation" & _
'         " WHERE NP02||NP06='FCP' and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND CU01(+)=SUBSTR(PA76,1,8) AND CU02(+)=SUBSTR(PA76,9) AND NA01(+)=CU10" & stConNA16
'      stSQL = Replace(stSQL, "decode(pa75,", "decode(pa76,")
'      cnnConnection.Execute stSQL, intI
'
'      '申請人年費代理人
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,FA10,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS,PATENT,CUSTOMER,FAGENT,Nation" & _
'         " WHERE NP02||NP06='FCP' and np07='605'" & _
'         " AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(CU96,1,8) AND FA02(+)=SUBSTR(CU96,9) AND NA01(+)=FA10" & stConNA16
'      stSQL = Replace(stSQL, "decode(pa75,", "decode(cu96,")
'      cnnConnection.Execute stSQL, intI
'
'      'Add by Morgan 2009/4/29 年費代理人也會有X編號
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,NA01,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS,PATENT,CUSTOMER c1,CUSTOMER c2,Nation WHERE NP02||NP06='FCP' and np07='605'" & _
'         " AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND c1.CU01(+)=SUBSTR(PA26,1,8) AND c1.CU02(+)=SUBSTR(PA26,9) AND c1.CU96 IS NOT NULL" & _
'         " AND c2.CU01(+)=SUBSTR(c1.CU96,1,8) AND c2.CU02(+)=SUBSTR(c1.CU96,9) AND NA01(+)=c2.CU10" & stConNA16
'      stSQL = Replace(stSQL, "decode(pa75,", "decode(c1.cu96,")
'      cnnConnection.Execute stSQL, intI
'
'      '代理人
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,FA10,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS,PATENT,CUSTOMER,FAGENT,Nation" & _
'         " WHERE NP02||NP06='FCP' and np07='605'" & _
'         " AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'end 2020/5/12
      
      'end 2008/11/19
      'Modify by Morgan 2009/7/13 +995,996
      'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
      'Modified by Lydia 2014/11/14 NA51 = decode(sp26,'Y51333010','" & midStr & "',na51)
      'Modified by Lydia 2016/02/03
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807)*/
      'Modify By Sindy 2021/4/28 + ,R15
      'Modified by Lydia 2022/11/03 區分FMP案; stConNA16=> Replace(UCase(stConNA16), "CP01", "NP02")
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,FA10,NULL,decode(sp26," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS, SERVICEPRACTICE, FAGENT, Nation" & _
         " WHERE NP02||NP06='FG' " & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate2 & _
         " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP15||SP61 IS NULL AND SP26 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10" & Replace(UCase(stConNA16), "CP01", "NP02")
      cnnConnection.Execute stSQL, intI
      'end       'Modified by Lydia 2014/11/14 NA51 = decode(pa75,'Y51333010','" & midStr & "',na51)
      
'Removed by Morgan 2012/6/15 併入原FMP案約定期限通知
'      'Added by Morgan 2012/5/28 +FMP
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,FA10,NULL,NA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS, PATENT, FAGENT, Nation,STAFF" & _
'         " WHERE NP02||NP06 in ('P','CFP') and st01(+)=NP10 and substr(st15,1,1)='F' and np07<>'605'" & strNpSqlOfNoSalesDuty & _
'         " AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA16 & _
'         " AND NOT EXISTS(SELECT * FROM CASEPROGRESS WHERE CP09=NP01 AND CP10='928' AND NP07='202')"
'      cnnConnection.Execute stSQL, intI
'
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,FA10,NULL,NA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS, PATENT, FAGENT, Nation,STAFF" & _
'         " WHERE NP02||NP06 in ('P','CFP') and st01(+)=NP10 and substr(st15,1,1)='F' and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA76,1,8) AND FA02(+)=SUBSTR(PA76,9) AND NA01(+)=FA10" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,NA01,NULL,NA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS, PATENT, Customer, Nation,STAFF" & _
'         " WHERE NP02||NP06 in ('P','CFP') and st01(+)=NP10 and substr(st15,1,1)='F' and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND CU01(+)=SUBSTR(PA76,1,8) AND CU02(+)=SUBSTR(PA76,9) AND NA01(+)=CU10" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,FA10,NULL,NA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS,PATENT,CUSTOMER,FAGENT,Nation,STAFF" & _
'         " WHERE NP02||NP06 in ('P','CFP') and st01(+)=NP10 and substr(st15,1,1)='F' and np07='605'" & _
'         " AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(CU96,1,8) AND FA02(+)=SUBSTR(CU96,9) AND NA01(+)=FA10" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,NA01,NULL,NA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS,PATENT,CUSTOMER c1,CUSTOMER c2,Nation,STAFF" & _
'         " WHERE NP02||NP06 in ('P','CFP') and st01(+)=NP10 and substr(st15,1,1)='F' and np07='605'" & _
'         " AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND c1.CU01(+)=SUBSTR(PA26,1,8) AND c1.CU02(+)=SUBSTR(PA26,9) AND c1.CU96 IS NOT NULL" & _
'         " AND c2.CU01(+)=SUBSTR(c1.CU96,1,8) AND c2.CU02(+)=SUBSTR(c1.CU96,9) AND NA01(+)=c2.CU10" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,FA10,NULL,NA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS,PATENT,CUSTOMER,FAGENT,Nation,STAFF" & _
'         " WHERE NP02||NP06 in ('P','CFP') and st01(+)=NP10 and substr(st15,1,1)='F' and np07='605'" & _
'         " AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA16
'      cnnConnection.Execute stSQL, intI
'
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'2' EV2,NP01,FA10,NULL,NA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS, SERVICEPRACTICE, FAGENT, Nation,STAFF" & _
'         " WHERE NP02||NP06 in ('PS','CPS') and st01(+)=NP10 and substr(st15,1,1)='F' " & strNpSqlOfNoSalesDuty & _
'         " AND NP08>=" & stDate1 & " AND NP08< " & stDate2 & _
'         " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP15||SP61 IS NULL AND SP26 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10" & stConNA16
'      cnnConnection.Execute stSQL, intI
'      'end 2012/5/28
'end 2012/6/15
   End If
   
   '工程師
   If (stDept <> "F22" And stDept <> "F23") Or bLvl4 = True Or bLvl5 = True Then
      '未交稿加判斷無核稿期限的(會有例外狀況需核完稿才給翻譯費故完稿日會先拿掉,如巨京)
      '未交稿,2個工作天後達承辦期限者(不含當日)-G9(未交稿)
      '改當日到期的
      'Add By Sindy 2016/7/28 + 有關翻譯之未交稿期限(下班翻改彈所內編號,外翻仍維持彈系統特殊人員"M"
      'Modified by Lydia 2016/09/22 CP10||CP27||CP57='201' => CP10='201' AND CP158=0 AND CP159=0
      'Modified by Lydia 2019/10/04 改成判斷畫面的員工編號 strUserNum=>strUserId
      'Modify By Sindy 2021/4/15 cp14==>decode(SIM01,null,CP14,SIM01)
      '   新案翻譯承辦人為所內工程師(上班譯-員編,下班譯-F編號)，請彈承辦工程師及其主管、Sharon 原:AND '" & stUserID & "' in (select ts2.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')
      '   原:SUBSTRB(S2.ST15,1,1)='F' ==> S1.ST15='F52'
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','G' EV1,'9' EV2,CP09,FA10,decode(SIM01,null,CP14,SIM01),CP13,CP06,CP07,CP48,NULL,0" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,PATENT,FAGENT,STAFF_IDMAP,STAFF S1,STAFF S2" & _
         " WHERE CP10='201' AND CP158=0 AND CP159=0 and substr(CP12,1,1)='F' AND CP48>" & stDate1 & " AND CP48<=" & strSrvDate(1) & stCP01 & _
         " AND EP02(+)=CP09 AND EP09 IS NULL AND EP08 IS NULL" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND S1.ST01(+)=CP14 AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01" & _
         " AND SUBSTRB(S1.ST15,1,1)='F' AND (S2.ST15 IS NULL OR S1.ST15='F52')" & _
         " AND EXISTS(select ts2.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')" & _
         " AND decode(SIM01,null,CP14,SIM01) in (" & stNumList & ")"
      cnnConnection.Execute stSQL, intI
      'Add By Sindy 2016/7/28 + 有關翻譯之未交稿期限(下班翻改彈所內編號,外翻仍維持彈系統特殊人員"M")
      'Modified by Lydia 2016/09/22 CP10||CP27||CP57='201' => CP10='201' AND CP158=0 AND CP159=0
      'Modified by Lydia 2019/10/04 改成判斷畫面的員工編號 strUserNum=>strUserId
      'Modify By Sindy 2021/4/15 cp14==>decode(SIM01,null,CP14,SIM01)
      '   新案翻譯承辦人為所內工程師(上班譯-員編,下班譯-F編號)，請彈承辦工程師及其主管、Sharon 原:AND '" & stUserID & "' in (select ts2.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')
      '   原:SUBSTRB(S2.ST15,1,1)='F' ==> S1.ST15='F52'
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','G' EV1,'9' EV2,CP09,FA10,decode(SIM01,null,CP14,SIM01),CP13,CP06,CP07,CP48,NULL,0" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,SERVICEPRACTICE,FAGENT,STAFF_IDMAP,STAFF S1,STAFF S2" & _
         " WHERE CP10='201' AND CP158=0 AND CP159=0 and substr(CP12,1,1)='F' AND CP48>" & stDate1 & " AND CP48<=" & strSrvDate(1) & stCP01 & _
         " AND EP02(+)=CP09 AND EP09 IS NULL AND EP08 IS NULL" & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)" & _
         " AND S1.ST01(+)=CP14 AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01" & _
         " AND SUBSTRB(S1.ST15,1,1)='F' AND (S2.ST15 IS NULL OR S1.ST15='F52')" & _
         " AND EXISTS(select ts2.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1')" & _
         " AND decode(SIM01,null,CP14,SIM01) in (" & stNumList & ")"
      cnnConnection.Execute stSQL, intI

      'Modify by Morgan 2008/10/9 未核稿不必判斷完稿日
      '未核稿且 2個工作天 後達核稿期限者(不含當日) --核稿人-C4(核稿期限,核稿人)
      'Modified by Lydia 2016/09/14 CP27||CP57 IS NULL => CP158=0 AND CP159=0
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','C' EV1,'4' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
         " FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,FAGENT" & _
         " WHERE " & IIf(strSrvDate(1) >= FCP核完日改用EP39, "EP39 IS NULL", "EP33 IS NULL") & stConEP08 & stConEP04 & _
         " AND CP09(+)=EP02 AND CP158=0 AND CP159=0 AND CP10='201'" & stCP01 & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)"
      cnnConnection.Execute stSQL, intI
         
      'Modified by Lydia 2016/09/14 CP27||CP57 IS NULL => CP158=0 AND CP159=0
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','C' EV1,'4' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
         " FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
         " WHERE " & IIf(strSrvDate(1) >= FCP核完日改用EP39, "EP39 IS NULL", "EP33 IS NULL") & stConEP08 & stConEP04 & _
         " AND CP09(+)=EP02 AND CP158=0 AND CP159=0 AND CP10='201'" & stCP01 & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9)"
      cnnConnection.Execute stSQL, intI

      '未核稿,所有未發文--核稿人-E(未發文)
      If idx = 1 Then
         'Modified by Lydia 2016/09/14 CP27||CP57 IS NULL => CP158=0 AND CP159=0
         'Modified by Lydia 2018/11/29 FCP-59635的新案翻譯未發文,A0022為承辦人主管同時為核稿人,造成重複主鍵,設R03="4"核稿人(' ' EV2=>'4' EV2)
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,'4' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
            " FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,FAGENT" & _
            " WHERE " & IIf(strSrvDate(1) >= FCP核完日改用EP39, "EP39 IS NULL", "EP33 IS NULL") & " and EP08>=" & stDate2 & stConEP04 & _
            " AND CP09(+)=EP02 AND CP158=0 AND CP159=0 AND CP10='201'" & stCP01 & _
            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND EP04<>CP14"
         cnnConnection.Execute stSQL, intI
         
         'Modified by Lydia 2016/09/14 CP27||CP57 IS NULL => CP158=0 AND CP159=0
         'Modified by Lydia 2018/11/29 FCP-59635的新案翻譯未發文,A0022為承辦人主管同時為核稿人,造成重複主鍵,設R03="4"核稿人(' ' EV2=>'4' EV2)
         stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
            " SELECT '" & strUserNum & "','E' EV1,'4' EV2,CP09,FA10,EP04,CP13,CP06,CP07,CP48,EP08,0" & _
            " FROM ENGINEERPROGRESS,CASEPROGRESS,SERVICEPRACTICE,FAGENT" & _
            " WHERE " & IIf(strSrvDate(1) >= FCP核完日改用EP39, "EP39 IS NULL", "EP33 IS NULL") & " and EP08>=" & stDate2 & stConEP04 & _
            " AND CP09(+)=EP02 AND CP158=0 AND CP159=0 AND CP10='201'" & stCP01 & _
            " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL" & _
            " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND EP04<>CP14"
         cnnConnection.Execute stSQL, intI
      End If
   End If
   
   '業務(FCP,FG抓NA51,其他抓NP10)
   If stDept = "F23" Or bLvl4 = True Or bLvl5 = True Then
      '未收文且 5個工作天 後達本所期限者(不含當日) --智權人員-D3(未收文,智權人員)
      'Modify by Morgan 2008/11/19 +年費另外
      'Modify by Morgan 2009/7/13 +995,996
      'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
      'Modified by Lydia 2014/11/14 NA51 = decode(pa75,'Y51333010','" & midStr & "',na51), stConNA51->stConNA51P
      'Modified by Lydia 2016/02/03
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807) */
      'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人)  取消 and np07<>'605' 條件
      'Modify By Sindy 2021/4/28 + ,R15
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,FA10,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS, PATENT, FAGENT, Nation" & _
         " WHERE NP02||NP06='FCP'" & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate3 & _
         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA51P & _
         " AND NOT (NP07='202' AND EXISTS(SELECT * FROM CASEPROGRESS WHERE CP09=NP01 AND CP10='928'))"
      cnnConnection.Execute stSQL, intI
      'Add By Sindy 2015/10/22 +法務:承辦業務組的未收文,僅限下一程序智權人員之ST15=F2字頭
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807) */
      'Modify By Sindy 2021/4/28 + ,R15
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,nvl(FA10,cu10),NULL,decode(FA10,null,n2.na51,decode(LC22," & midStr & ",n1.na51)) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS, LawCase, FAGENT, Nation n1, Staff,Customer, Nation n2" & _
         " WHERE NP02||NP06 in('FCL','CFL','LIN')" & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate3 & _
         " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05 AND LC08||LC34 IS NULL" & _
         " AND FA01(+)=SUBSTR(LC22,1,8) AND FA02(+)=SUBSTR(LC22,9) AND n1.NA01(+)=FA10" & _
         " AND CU01(+)=SUBSTR(LC11,1,8) AND CU02(+)=SUBSTR(LC11,9) AND n2.NA01(+)=cu10" & stConNA51L & _
         " AND NP10=ST01(+) AND substr(ST15,1,2)='F2'"
      cnnConnection.Execute stSQL, intI
      '2015/10/22 END
      'end 2016/09/22
      
      'Add by Morgan 2010/7/29 P,CFP也改抓國家檔
      'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
      'Modified by Lydia 2014/11/14
      'Modified by Lydia 2016/02/03
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807) */
      'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人)  取消 and np07<>'605' 條件
      'Modify By Sindy 2021/4/28 + ,R15
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,FA10,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS,CASEPROGRESS, PATENT, FAGENT, Nation" & _
         " WHERE NP02||NP06 in ('P','CFP')" & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate3 & _
         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F'" & _
         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA51P & _
         " AND NOT (NP07='202' AND EXISTS(SELECT * FROM CASEPROGRESS WHERE CP09=NP01 AND CP10='928'))"
      cnnConnection.Execute stSQL, intI
      
'Removed by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人)
'
'      'Add by Morgan 2008/11/20
'      '個案年費代理人
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,FA10,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS, PATENT, FAGENT, Nation" & _
'         " WHERE NP02||NP06='FCP' and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate3 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA76,1,8) AND FA02(+)=SUBSTR(PA76,9) AND NA01(+)=FA10" & stConNA51P
'      stSQL = Replace(stSQL, "decode(pa75,", "decode(pa76,")
'      cnnConnection.Execute stSQL, intI
'
'      'Add by Morgan 2010/7/29 P,CFP也改抓國家檔
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,FA10,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS,CASEPROGRESS, PATENT, FAGENT, Nation" & _
'         " WHERE NP02||NP06 IN ('P','CFP') and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate3 & _
'         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F'" & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(PA76,1,8) AND FA02(+)=SUBSTR(PA76,9) AND NA01(+)=FA10" & stConNA51P
'      stSQL = Replace(stSQL, "decode(pa75,", "decode(pa76,")
'      cnnConnection.Execute stSQL, intI
'
'      'Add by Morgan 2009/4/29 年費代理人也會有X編號
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,NA01,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS, PATENT, CUSTOMER, Nation" & _
'         " WHERE NP02||NP06='FCP' and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate3 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND CU01(+)=SUBSTR(PA76,1,8) AND CU02(+)=SUBSTR(PA76,9) AND NA01(+)=CU10" & stConNA51P
'      stSQL = Replace(stSQL, "decode(pa75,", "decode(pa76,")
'      cnnConnection.Execute stSQL, intI
'
'      'Add by Morgan 2010/7/29 P,CFP也改抓國家檔
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,NA01,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS,CASEPROGRESS, PATENT, CUSTOMER, Nation" & _
'         " WHERE NP02||NP06 IN ('P','CFP') and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate3 & _
'         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F'" & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NOT NULL" & _
'         " AND CU01(+)=SUBSTR(PA76,1,8) AND CU02(+)=SUBSTR(PA76,9) AND NA01(+)=CU10" & stConNA51P
'      stSQL = Replace(stSQL, "decode(pa75,", "decode(pa76,")
'      cnnConnection.Execute stSQL, intI
'
'      '申請人年費代理人
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,FA10,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS, PATENT,CUSTOMER, FAGENT, Nation" & _
'         " WHERE NP02||NP06='FCP' and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate3 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(CU96,1,8) AND FA02(+)=SUBSTR(CU96,9) AND NA01(+)=FA10" & stConNA51P
'      stSQL = Replace(stSQL, "decode(pa75,", "decode(cu96,")
'      cnnConnection.Execute stSQL, intI
'
'      'Add by Morgan 2010/7/29 P,CFP也改抓國家檔
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,FA10,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS,CASEPROGRESS, PATENT,CUSTOMER, FAGENT, Nation" & _
'         " WHERE NP02||NP06 IN ('P','CFP') and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate3 & _
'         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F'" & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NOT NULL" & _
'         " AND FA01(+)=SUBSTR(CU96,1,8) AND FA02(+)=SUBSTR(CU96,9) AND NA01(+)=FA10" & stConNA51P
'      stSQL = Replace(stSQL, "decode(pa75,", "decode(cu96,")
'      cnnConnection.Execute stSQL, intI
'
'      'Add by Morgan 2009/4/29 年費代理人也會有X編號
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,NA01,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS, PATENT,CUSTOMER c1,CUSTOMER c2,Nation" & _
'         " WHERE NP02||NP06='FCP' and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate3 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND c1.CU01(+)=SUBSTR(PA26,1,8) AND c1.CU02(+)=SUBSTR(PA26,9) AND c1.CU96 IS NOT NULL" & _
'         " AND c2.CU01(+)=SUBSTR(c1.CU96,1,8) AND c2.CU02(+)=SUBSTR(c1.CU96,9) AND NA01(+)=c2.CU10" & stConNA51P
'      stSQL = Replace(stSQL, "decode(pa75,", "decode(c1.cu96,")
'      cnnConnection.Execute stSQL, intI
'
'      'Add by Morgan 2010/7/29 P,CFP也改抓國家檔
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,NA01,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS,CASEPROGRESS, PATENT,CUSTOMER c1,CUSTOMER c2,Nation" & _
'         " WHERE NP02||NP06 IN ('P','CFP') and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate3 & _
'         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F'" & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND c1.CU01(+)=SUBSTR(PA26,1,8) AND c1.CU02(+)=SUBSTR(PA26,9) AND c1.CU96 IS NOT NULL" & _
'         " AND c2.CU01(+)=SUBSTR(c1.CU96,1,8) AND c2.CU02(+)=SUBSTR(c1.CU96,9) AND NA01(+)=c2.CU10" & stConNA51P
'      stSQL = Replace(stSQL, "decode(pa75,", "decode(c1.cu96,")
'      cnnConnection.Execute stSQL, intI
'
'      '代理人
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,FA10,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS, PATENT,CUSTOMER, FAGENT, Nation" & _
'         " WHERE NP02||NP06='FCP' and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate3 & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA51P
'      cnnConnection.Execute stSQL, intI
'
'      'Add by Morgan 2010/7/29 P,CFP也改抓國家檔
'      'Modified by Lydia 2014/11/14
'      'Modified by Lydia 2016/02/03
'      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
'         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,FA10,NULL,decode(pa75," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22" & _
'         " From NEXTPROGRESS,CASEPROGRESS, PATENT,CUSTOMER, FAGENT, Nation" & _
'         " WHERE NP02||NP06 IN ('P','CFP') and np07='605' AND NP08>=" & stDate1 & " AND NP08< " & stDate3 & _
'         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F'" & _
'         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57||PA108 IS NULL AND PA76 IS NULL" & _
'         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9) AND CU96 IS NULL" & _
'         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & stConNA51P
'      cnnConnection.Execute stSQL, intI
'
'end 2020/5/12
      
      'Modify by Morgan 2009/7/13 +995,996
      'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
      'Modified by Lydia 2014/11/14 NA51 = decode(sp26,'Y51333010','" & midStr & "',na51),stConNA51->stconsp26
      'Modified by Lydia 2016/02/03
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807) */
      'Modify By Sindy 2021/4/28 + ,R15
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,FA10,NULL,decode(sp26," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS, SERVICEPRACTICE, FAGENT, Nation" & _
         " WHERE NP02||NP06='FG'" & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate3 & _
         " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP15||SP61 IS NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10" & stConSP26
      cnnConnection.Execute stSQL, intI
         
      'Add by Morgan 2010/7/29 PS,CFS也改抓國家檔
      'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
      'Modified by Lydia 2016/02/03
      'Modified by Lydia 2016/09/22 + /*+ INDEX(NEXTPROGRESS IDXNP0807) */
      'Modify By Sindy 2021/4/28 + ,R15
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R15)" & _
         " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,FA10,NULL,decode(sp26," & midStr & ",na51) nNA51,NP08,NP09,NULL,NULL,NP22,NP23" & _
         " From NEXTPROGRESS,CASEPROGRESS, SERVICEPRACTICE, FAGENT, Nation" & _
         " WHERE NP02||NP06 IN ('PS,CPS')" & strNpSqlOfNoSalesDuty & _
         " AND NP23>=" & stDate1 & " AND NP23< " & stDate3 & _
         " AND CP09(+)=NP01 AND SUBSTR(CP12,1,1)='F'" & _
         " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP15||SP61 IS NULL" & _
         " AND FA01(+)=SUBSTR(SP26,1,8) AND FA02(+)=SUBSTR(SP26,9) AND NA01(+)=FA10" & stConSP26
      cnnConnection.Execute stSQL, intI
      'end       'Modified by Lydia 2014/11/14 NA51 = decode(pa75,'Y51333010','" & midStr & "',na51)
      
      '未收文且 5個工作天 後達本所期限者(不含當日) --智權人員(非FCP,FG)-45(未收文,智權人員)
      'Memo by Morgan 2008/11/20 非國外部專利案會指定承辦人管制，暫不改--David
      'Modify by Morgan 2009/7/13 +995,996
      'Remove by Morgan 2010/7/29改抓國家檔,移到上面
      'Modify by Morgan 2009/7/13 +995,996
      'Remove by Morgan 2010/7/29改抓國家檔,移到上面
      
      '寄中說949已收文未發文且 2個工作天 後達本所期限者(不含當日) --智權人員-A2(本所期限,智權人員)
      'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' and CP10='949' => CP01='FCP' and CP10='949' AND CP158=0 AND CP159=0
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R14)" & _
         " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0,cp142" & _
         " From CASEPROGRESS,PATENT,FAGENT,Nation Where CP01='FCP' and CP10='949' AND CP158=0 AND CP159=0" & stConCP06 & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 " & stConNA51
      cnnConnection.Execute stSQL, intI
   End If
   
   'Add By Sindy 2022/3/15
   '(1)通知對象為承辦及程序
   If stDept = "F22" Or stDept = "F23" Or bLvl4 = True Or bLvl5 = True Then
      '111.03.03-Sharon - FCP
      '以上104年請作單將已發文未請款未彈期限通知排除是非常危險,故需將此設定調整
      '以下案件性質已發文未請款需彈期限通知:
      '(2)發文日+7個工作天未請款案件
      '(3)排除已上"不請款"
      '201新案翻譯、209檢視中說、210製作中說、401變更、403更改、416實體審查、601領證及繳年費、605年費、701讓與、702合併、917超頁、超項費
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'" & IIf(stDept = "F22", "2", "1") & "' EV2,CP09,FA10,CP14,CP13" & _
         ",NULL,NULL,workdayadd(8,cp27) C06,NULL,0" & _
         " FROM CASEPROGRESS,PATENT,FAGENT,CASEPROPERTYMAP,Nation" & _
         " WHERE CP01='FCP' AND CP27>" & stDate1 & "-10000 AND CP20||CP57||CP60 IS NULL AND SUBSTR(CP12,1,2)='F2'" & _
         " AND CP10 in('201','209','210','401','403','416','601','605','701','702','917')" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
         " AND CP27<=" & stDate7 & _
         " AND NA01(+)=FA10" & IIf(stDept = "F22", " AND (CP14 IS NULL OR CP14<>NA16)" & stConNA16, stConNA51P) & _
         " AND not exists(SELECT R01,R02,R03,R04,R12 FROM R060204 WHERE R01='" & strUserNum & "' and R02='F' and R03='" & IIf(stDept = "F22", "2", "1") & "' and R04=CP09 and R12=0)"
      cnnConnection.Execute stSQL, intI
      'end 2014/8/27
      '(4)235核對中說格式：此性質原本就已設為"不請款"，請判斷235核對中說格式發文日+7個工作天，
      '提申那道(101發明申請、102新型申請) 未請款案件，則需彈期限通知
      'Modified by Lydia 2022/11/03 區分FMP案; stConNA16=> Replace(UCase(stConNA16), "CP01", "C1.CP01")
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'" & IIf(stDept = "F22", "2", "1") & "' EV2,c1.CP09,FA10,c1.CP14,c1.CP13" & _
         ",NULL,NULL,workdayadd(8,c1.cp27) C06,NULL,0" & _
         " FROM CASEPROGRESS c1,PATENT,FAGENT,CASEPROPERTYMAP,Nation,CASEPROGRESS c2" & _
         " WHERE c1.CP01='FCP' AND c1.CP27>" & stDate1 & "-10000 AND c1.CP57||c1.CP60 IS NULL AND SUBSTR(c1.CP12,1,2)='F2'" & _
         " AND c1.CP10='235'" & _
         " AND PA01(+)=c1.CP01 AND PA02(+)=c1.CP02 AND PA03(+)=c1.CP03 AND PA04(+)=c1.CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=c1.CP01 AND CPM02(+)=c1.CP10" & _
         " AND c1.CP27<=" & stDate7 & _
         " AND NA01(+)=FA10" & IIf(stDept = "F22", " AND (c1.CP14 IS NULL OR c1.CP14<>NA16)" & Replace(UCase(stConNA16), "CP01", "C1.CP01"), stConNA51P) & _
         " AND c2.cp01=c1.cp01 AND c2.cp02=c1.cp02 AND c2.cp03=c1.cp03 AND c2.cp04=c1.cp04" & _
         " AND c2.cp10 in('101','102') AND c2.cp09=c1.cp43 AND c2.CP27 is not null AND c2.CP20||c2.CP57||c2.CP60 IS NULL" & _
         " AND not exists(SELECT R01,R02,R03,R04,R12 FROM R060204 WHERE R01='" & strUserNum & "' and R02='F' and R03='" & IIf(stDept = "F22", "2", "1") & "' and R04=c1.CP09 and R12=0)"
      cnnConnection.Execute stSQL, intI
      '110.03.16-增加FMP案的控管 - P,CFP
      '(1)性質：401變更、403更改、416實體審查、701讓與、702合併
      '以上Key已提申+7個工作天未請款案件
      'Modified by Lydia 2022/11/03  區分FMP案 CP14<>NA16=>NVL(NA79,NA16)
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'" & IIf(stDept = "F22", "2", "1") & "' EV2,CP09,FA10,CP14,CP13" & _
         ",NULL,NULL,workdayadd(8,CP47) C06,NULL,0" & _
         " FROM CASEPROGRESS,PATENT,FAGENT,CASEPROPERTYMAP,Nation" & _
         " WHERE CP01 in('P','CFP') AND CP27>" & stDate1 & "-10000 AND CP20||CP57||CP60 IS NULL AND SUBSTR(CP12,1,2)='F2'" & _
         " AND CP10 in('401','403','416','701','702')" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP47 is not null" & _
         " AND CP47<=" & stDate7 & _
         " AND NA01(+)=FA10" & IIf(stDept = "F22", " AND (CP14 IS NULL OR CP14<>NVL(NA79,NA16))" & stConNA16, stConNA51P) & _
         " AND not exists(SELECT R01,R02,R03,R04,R12 FROM R060204 WHERE R01='" & strUserNum & "' and R02='F' and R03='" & IIf(stDept = "F22", "2", "1") & "' and R04=CP09 and R12=0)"
      cnnConnection.Execute stSQL, intI
      'Modify By Sindy 2022/3/22
      '性質：601領證及繳年費、605年費
      '要區分代理人為Y53374北京寰華知識產權代理有限公司,則維持上列(1)的控管
       'Modified by Lydia 2022/11/03  區分FMP案 CP14<>NA16=>NVL(NA79,NA16)
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'" & IIf(stDept = "F22", "2", "1") & "' EV2,CP09,FA10,CP14,CP13" & _
         ",NULL,NULL,workdayadd(8,CP47) C06,NULL,0" & _
         " FROM CASEPROGRESS,PATENT,FAGENT,CASEPROPERTYMAP,Nation" & _
         " WHERE CP01 in('P','CFP') AND CP27>" & stDate1 & "-10000 AND CP20||CP57||CP60 IS NULL AND SUBSTR(CP12,1,2)='F2'" & _
         " AND CP10 in('601','605') AND cp44='Y53374000'" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND CP47 is not null" & _
         " AND CP47<=" & stDate7 & _
         " AND NA01(+)=FA10" & IIf(stDept = "F22", " AND (CP14 IS NULL OR CP14<>NVL(NA79,NA16))" & stConNA16, stConNA51P) & _
         " AND not exists(SELECT R01,R02,R03,R04,R12 FROM R060204 WHERE R01='" & strUserNum & "' and R02='F' and R03='" & IIf(stDept = "F22", "2", "1") & "' and R04=CP09 and R12=0)"
      cnnConnection.Execute stSQL, intI
      '非寰華案,就以1909已提申的發文日+7個工作天控管
      'Modified by Lydia 2022/11/03  區分FMP案 CP14<>NA16=>NVL(NA79,NA16), stConNA16=> Replace(UCase(stConNA16), "CP01", "C1.CP01")
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'" & IIf(stDept = "F22", "2", "1") & "' EV2,c1.CP09,FA10,c1.CP14,c1.CP13" & _
         ",NULL,NULL,workdayadd(8,c2.cp27),NULL,0" & _
         " FROM CASEPROGRESS c1,PATENT,FAGENT,CASEPROPERTYMAP,Nation,CASEPROGRESS c2" & _
         " WHERE c1.CP01 in('P','CFP') AND c1.CP27>" & stDate1 & "-10000 AND c1.CP20||c1.CP57||c1.CP60 IS NULL AND SUBSTR(c1.CP12,1,2)='F2'" & _
         " AND c1.CP10 in('601','605') AND c1.cp44<>'Y53374000'" & _
         " AND PA01(+)=c1.CP01 AND PA02(+)=c1.CP02 AND PA03(+)=c1.CP03 AND PA04(+)=c1.CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=c1.CP01 AND CPM02(+)=c1.CP10 AND c1.CP47 is not null" & _
         " AND c2.cp01=c1.cp01 AND c2.cp02=c1.cp02 AND c2.cp03=c1.cp03 AND c2.cp04=c1.cp04" & _
         " AND c2.cp10='1909' AND c2.cp43=c1.cp09 AND c2.CP27 IS NOT NULL" & _
         " AND c2.CP27<=" & stDate7 & _
         " AND NA01(+)=FA10" & IIf(stDept = "F22", " AND (c1.CP14 IS NULL OR c1.CP14<>NVL(NA79,NA16))" & Replace(UCase(stConNA16), "CP01", "C1.CP01"), stConNA51P) & _
         " AND not exists(SELECT R01,R02,R03,R04,R12 FROM R060204 WHERE R01='" & strUserNum & "' and R02='F' and R03='" & IIf(stDept = "F22", "2", "1") & "' and R04=c1.CP09 and R12=0)"
      cnnConnection.Execute stSQL, intI
      '(2)通知申請案號(1101)發文日+7個工作天,
      '提申那道(101發明申請、102新型申請、103設計申請) 未請款案件，則需彈期限通知
      'Modified by Lydia 2022/11/03  區分FMP案 CP14<>NA16=>NVL(NA79,NA16), stConNA16=> Replace(UCase(stConNA16), "CP01", "C1.CP01")
      stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & _
         " SELECT '" & strUserNum & "','F' EV1,'" & IIf(stDept = "F22", "2", "1") & "' EV2,c1.CP09,FA10,c1.CP14,c1.CP13" & _
         ",NULL,NULL,workdayadd(8,c1.cp27) C06,NULL,0" & _
         " FROM CASEPROGRESS c1,PATENT,FAGENT,CASEPROPERTYMAP,Nation,CASEPROGRESS c2" & _
         " WHERE c1.CP01 in('P','CFP') AND c1.CP27>" & stDate1 & "-10000 AND c1.CP57||c1.CP60 IS NULL AND SUBSTR(c1.CP12,1,2)='F2'" & _
         " AND c1.CP10='1101'" & _
         " AND PA01(+)=c1.CP01 AND PA02(+)=c1.CP02 AND PA03(+)=c1.CP03 AND PA04(+)=c1.CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
         " AND CPM01(+)=c1.CP01 AND CPM02(+)=c1.CP10" & _
         " AND c1.CP27<=" & stDate7 & _
         " AND NA01(+)=FA10" & IIf(stDept = "F22", " AND (c1.CP14 IS NULL OR c1.CP14<>NVL(NA79,NA16))" & Replace(UCase(stConNA16), "CP01", "C1.CP01"), stConNA51P) & _
         " AND c2.cp01=c1.cp01 AND c2.cp02=c1.cp02 AND c2.cp03=c1.cp03 AND c2.cp04=c1.cp04" & _
         " AND c2.cp10 in('101','102','103') AND c2.cp09=c1.cp43 AND c2.CP27 is not null AND c2.CP20||c2.CP57||c2.CP60 IS NULL" & _
         " AND not exists(SELECT R01,R02,R03,R04,R12 FROM R060204 WHERE R01='" & strUserNum & "' and R02='F' and R03='" & IIf(stDept = "F22", "2", "1") & "' and R04=c1.CP09 and R12=0)"
      cnnConnection.Execute stSQL, intI
   End If
   '2022/3/15 END
   
   'Added by Lydia 2015/09/09 +早收文
   '通知工程師和程序
   If stDept = "F21" Or stDept = "F22" Then
       strExc(2) = CompWorkDay(4, strSrvDate(1)) '系統日+3個工作天
       strExc(3) = CompDate(2, 14, strSrvDate(1)) '系統日+14個日歷天
       strExc(4) = CompDate(2, -14, strSrvDate(1)) '系統日-14個日歷天
      '1.  若FMP之香港標準記錄請求(110)或澳門發明進度資料，若本所期限尚未達彈跳條件時，則再判斷若收文日<系統日－14個日曆天且本所期限<=系統日＋14個日曆天；
            'Modified by Lydia 2016/09/22 CP57||CP27 IS NULL => CP158=0 AND CP159=0
            strExc(5) = " SELECT '" & strUserNum & "','H' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
                    " From CASEPROGRESS,PATENT,FAGENT,CASEMAP" & _
                    " Where (CP01='P' OR CP01='CFP') AND CP158=0 AND CP159=0 AND SUBSTR(CP12,1,1)='F' and CP14 IN (" & stNumList & ")" & _
                    " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
                    " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
                    " AND CM01(+)=CP01 AND CM02(+)=CP02 AND CM03(+)=CP03 AND CM04(+)=CP04 AND CM10='4' AND CP10='110'" & _
                    " AND CP06 >= " & strExc(2) & " and cp05<" & strExc(4) & " and cp06 <= " & strExc(3)
            'Added by Lydia 2016/09/22
            stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & strExc(5)
            cnnConnection.Execute stSQL, intI
            'end 2016/09/22
            
            'Modified by Lydia 2016/09/22 拿掉Union
            'strExc(5) = " SELECT '" & strUserNum & "','H' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
                    " From CASEPROGRESS,PATENT,FAGENT,CASEMAP" & _
                    " Where (CP01='P' OR CP01='CFP') AND CP57||CP27 IS NULL and CP14 IN (" & stNumList & ")" & _
                    " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
                    " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
                    " AND CM01(+)=CP01 AND CM02(+)=CP02 AND CM03(+)=CP03 AND CM04(+)=CP04 AND CM10='5' AND CP10 in (" & CaseMapIn & ")" & _
                    " AND CP06 >= " & strExc(2) & " and cp05<" & strExc(4) & " and cp06 <= " & strExc(3)
            strExc(5) = " SELECT '" & strUserNum & "','H' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
                    " From CASEPROGRESS,PATENT,FAGENT,CASEMAP" & _
                    " Where (CP01='P' OR CP01='CFP') AND CP158=0 AND CP159=0 and CP14 IN (" & stNumList & ")" & _
                    " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
                    " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
                    " AND CM01(+)=CP01 AND CM02(+)=CP02 AND CM03(+)=CP03 AND CM04(+)=CP04 AND CM10='5' AND CP10 in (" & CaseMapIn & ")" & _
                    " AND CP06 >= " & strExc(2) & " and cp05<" & strExc(4) & " and cp06 <= " & strExc(3)
            stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & strExc(5)
            cnnConnection.Execute stSQL, intI
      '2.若FMP或FCP之實審(416)、年費(605)進度資料，若本所期限尚未達彈跳條件時，則再判斷若收文日<系統日－14個日曆天且本所期限<=系統日＋14個日曆天；
             'Modified by Lydia 2016/09/22 CP57||CP27 IS NULL =>CP158=0 AND CP159=0 ; + /*+ INDEX(CASEPROGRESS IDXCP15815905011014)*/
             strExc(5) = " SELECT '" & strUserNum & "','H' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
                    " From CASEPROGRESS,PATENT,FAGENT" & _
                    " Where (CP01='P' OR CP01='CFP') AND CP158=0 AND CP159=0 AND SUBSTR(CP12,1,1)='F' and CP14 IN (" & stNumList & ")" & _
                    " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
                    " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
                    " AND CP10 in ('416','605') AND CP06 >= " & strExc(2) & " and cp05<" & strExc(4) & " and cp06 <= " & strExc(3)
             'Added by Lydia 2016/09/22
             stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & strExc(5)
             cnnConnection.Execute stSQL, intI
             'end 2016/09/22
             'Modified by Lydia 2016/09/14 CP01||CP27||CP57='FCP' => CP01='FCP' AND CP158=0 AND CP159=0
             'Modified by Lydia 2016/09/22 拿掉Union
             'strExc(5) = strExc(5) & " Union SELECT '" & strUserNum & "','H' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
                    " From CASEPROGRESS,PATENT,FAGENT" & _
                    " Where CP01='FCP' AND CP158=0 AND CP159=0 AND CP10 in ('416','605') and CP14 IN (" & stNumList & ") AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
                    " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
                    " AND CP06 >= " & strExc(2) & " and cp05<" & strExc(4) & " and cp06 <= " & strExc(3)
             strExc(5) = " SELECT '" & strUserNum & "','H' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
                    " From CASEPROGRESS,PATENT,FAGENT" & _
                    " Where CP01='FCP' AND CP158=0 AND CP159=0 AND CP10 in ('416','605') and CP14 IN (" & stNumList & ") AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
                    " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
                    " AND CP06 >= " & strExc(2) & " and cp05<" & strExc(4) & " and cp06 <= " & strExc(3)
             stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & strExc(5)
             cnnConnection.Execute stSQL, intI
      '3.若FMP或FCP之分割(307)進度資料，若本所期限尚未達彈跳條件時，則再判斷若收文日<系統日－14個日曆天且本所期限<=系統日＋14個日曆；
             strExc(5) = " SELECT '" & strUserNum & "','H' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
                    " From CASEPROGRESS,PATENT,FAGENT" & _
                    " Where (CP01='P' OR CP01='CFP') AND CP57||CP27 IS NULL AND SUBSTR(CP12,1,1)='F' and CP14 IN (" & stNumList & ")" & _
                    " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
                    " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) " & _
                    " AND CP10 ='307' AND CP06 >= " & strExc(2) & " and cp05<" & strExc(4) & " and cp06 <= " & strExc(3)
             'Added by Lydia 2016/09/22
             stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & strExc(5)
             cnnConnection.Execute stSQL, intI
             'end 2016/09/22
             'Modified by Lydia 2016/09/22 strExc(5) & " Union SELECT => " SELECT
             strExc(5) = " SELECT '" & strUserNum & "','H' EV1,'2' EV2,CP09,FA10,CP14,CP13,CP06,CP07,CP48,NULL,0" & _
                    " From CASEPROGRESS,PATENT,FAGENT" & _
                    " Where CP01||CP27||CP57='FCP' and CP14 IN (" & stNumList & ") AND CP10 ='307' AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
                    " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
                    " AND CP06 >= " & strExc(2) & " and cp05<" & strExc(4) & " and cp06 <= " & strExc(3)
             stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12)" & strExc(5)
             cnnConnection.Execute stSQL, intI
   End If
   'end 2015/09/09
   
   'Add By Sindy 2017/1/16 I.准未請款/J.分割建議/K.通知告准
   '通知告准未發文的核准函(非收文日)之次日起
   strExc(0) = "SELECT C2.cp01 C2_CP01,C2.cp02 C2_CP02,C2.cp03 C2_CP03,C2.cp04 C2_CP04,C2.cp09 C2_CP09,C2.CP66 C2_CP66,C2.CP67 C2_CP67,C2.CP13 C2_CP13,C2.CP14 C2_CP14,Na16,fa10,pa85,C2.cp27 C2_CP27,PA162" & _
               ",nvl(s1.st52,'') s1_ST52,nvl(s2.st52,'') s2_ST52,nvl(s3.st52,'') s3_ST52" & _
               " From CASEPROGRESS C1,CASEPROGRESS C2,PATENT,FAGENT,Nation,staff s1,staff s2,staff s3" & _
               " Where C1.CP01='FCP' and C1.CP10='1917' AND C1.CP158=0 AND C1.CP159=0" & _
               " AND C1.CP43=C2.CP09(+) AND C2.CP10='1001'" & _
               " AND PA01(+)=C2.CP01 AND PA02(+)=C2.CP02 AND PA03(+)=C2.CP03 AND PA04(+)=C2.CP04" & _
               " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
               " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
               " AND NA01(+)=FA10" & _
               " AND C2.cp13=s1.st01(+) AND C2.CP14=s2.st01(+) AND Na16=s3.st01(+)" & _
               " order by 1,2,3,4"
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If rsA.RecordCount > 0 Then
      rsA.MoveFirst
      Do While Not rsA.EOF
         'Modify By Sindy 2017/3/15 FCP-54734排除不向客戶收款(AND cp20 is null)
         'Modify By Sindy 2017/4/14 + 修正204或主動修正203
         'Modify By Sindy 2017/4/20 + 補充說明206擇一申復239
         strSql = "SELECT cp09,cp60,a1k01,a1k02,a1k19,a1k20,DST01 FROM caseprogress,acc1k0,DivsugText" & _
                  " WHERE cp01='" & rsA.Fields("C2_cp01") & "'" & _
                  " AND cp02='" & rsA.Fields("C2_cp02") & "'" & _
                  " AND cp03='" & rsA.Fields("C2_cp03") & "'" & _
                  " AND cp04='" & rsA.Fields("C2_cp04") & "'" & _
                  " AND substr(cp09,1,1)='A' AND cp10 in('205','107','204','203','206','239') AND cp20 is null" & _
                  " AND cp60=a1k01(+)" & _
                  " AND cp01=DST01(+)" & _
                  " AND cp02=DST02(+)" & _
                  " AND cp03=DST03(+)" & _
                  " AND cp04=DST04(+)" & _
                  " order by A1K02 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strCP09 = RsTemp.Fields("cp09")
            strCP60 = "" & RsTemp.Fields("cp60")
            strA1K01 = "" & RsTemp.Fields("a1k01")
            strA1K19 = "" & RsTemp.Fields("a1k19") '請款單輸入日期
            If strA1K19 <> "" Then strA1K19 = DBDATE(strA1K19)
            strA1K20 = "" & RsTemp.Fields("a1k20") '請款單輸入時間
            If strA1K20 <> "" Then strA1K20 = Format(strA1K20, "000000")
            strDST01 = "" & RsTemp.Fields("DST01")
            '未請款
            If strCP60 = "" Then
               'A.該案若有A類申復205或再審107或修正204或主動修正203或補充說明206或擇一申復239未請款(無CP60)時，增加事件為"I准未請款"之提醒
               '核准函輸入日期（非收文日）之次日起
               '提醒工程師、各區承辦、各區程序，及三人主管，直至A類申復或再審或修正或主動修正或補充說明或擇一申復請款
               If (stDept = "F21" Or stDept = "F22" Or stDept = "F23") And _
                   (InStr(stNumList, rsA.Fields("C2_CP13")) > 0 Or (InStr(stNumList, "" & rsA.Fields("s1_ST52")) > 0 And "" & rsA.Fields("s1_ST52") <> "") Or _
                   InStr(stNumList, rsA.Fields("C2_CP14")) > 0 Or (InStr(stNumList, "" & rsA.Fields("s2_ST52")) > 0 And "" & rsA.Fields("s2_ST52") <> "") Or _
                   ((InStr(stNumList, "" & rsA.Fields("Na16")) > 0 And "" & rsA.Fields("Na16") <> "") Or (InStr(stNumList, "" & rsA.Fields("s3_ST52")) > 0 And "" & rsA.Fields("s3_ST52") <> ""))) And _
                  rsA.Fields("C2_CP66") < strSrvDate(1) Then
                  
                  stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12) values(" & _
                  "'" & strUserNum & "','I','1'," & _
                  CNULL(rsA.Fields("C2_CP09")) & "," & _
                  CNULL(rsA.Fields("fa10")) & "," & _
                  CNULL(rsA.Fields("C2_CP14")) & "," & _
                  CNULL(rsA.Fields("C2_CP13")) & "," & _
                  "NULL,NULL,NULL,NULL,0)"
                  cnnConnection.Execute stSQL, intI
               End If
               '核准函未上發文日且案件之PA162有註記應另函通知初審核准後分割者，
               '若案件無「J分割建議」(以本所案號讀取分割建議定稿文字檔之DST05)
               '且非日文定稿(依定稿語文規則)時，增加事件為"分割建議"之提醒
               '核准函輸入日期（非收文日）之次日起
               '提醒工程師及其主管，直至A類申復或再審請款
               'Modifiedby Morgan 2022/10/11 取消日文定稿限制
               If "" & rsA.Fields("C2_CP27") = "" And rsA.Fields("PA162") = "Y" Then
                  If stDept = "F21" And _
                     (InStr(stNumList, rsA.Fields("C2_CP14")) > 0 Or (InStr(stNumList, "" & rsA.Fields("s2_ST52")) > 0 And "" & rsA.Fields("s2_ST52") <> "")) And _
                     rsA.Fields("C2_CP66") < strSrvDate(1) And strDST01 = "" Then
                     
                     stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12) values(" & _
                     "'" & strUserNum & "','J','1'," & _
                     CNULL(rsA.Fields("C2_CP09")) & "," & _
                     CNULL(rsA.Fields("fa10")) & "," & _
                     CNULL(rsA.Fields("C2_CP14")) & "," & _
                     CNULL(rsA.Fields("C2_CP13")) & "," & _
                     "NULL,NULL,NULL,NULL,0)"
                     cnnConnection.Execute stSQL, intI
                  End If
               End If
            '已請款
            Else
               'C.上述A點已請款案件，增加事件為"K通知告准"之提醒
               '提醒各區程序及其主管，直至該案"通知告准" D類進度上發文日
               If stDept = "F22" And _
                  ((InStr(stNumList, "" & rsA.Fields("Na16")) > 0 And "" & rsA.Fields("Na16") <> "") Or (InStr(stNumList, "" & rsA.Fields("s3_ST52")) > 0 And "" & rsA.Fields("s3_ST52") <> "")) Then
                  'Modify By Sindy 2017/2/18 +R13.備註
                  '106/2/18：敏莉提"優先請款"只顯示於請款輸入日期時間大於核准函輸入日期時間的資料
                  If strA1K19 & strA1K20 > rsA.Fields("C2_CP66") & Format(rsA.Fields("C2_CP67") & "00", "000000") Then
                     stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13) values(" & _
                     "'" & strUserNum & "','K','2'," & _
                     CNULL(rsA.Fields("C2_CP09")) & "," & _
                     CNULL(rsA.Fields("fa10")) & "," & _
                     CNULL(rsA.Fields("C2_CP14")) & "," & _
                     CNULL(rsA.Fields("C2_CP13")) & "," & _
                     "NULL,NULL,NULL,NULL,0,'優先請款;')"
                     cnnConnection.Execute stSQL, intI
                  End If
               End If
            End If
         End If
         rsA.MoveNext
      Loop
   End If
   rsA.Close
   '2017/1/16 END
   
   'Added by Lydia 2017/11/29 FCP案件命名電子化:外專工程師及業務承辦組增加"L-待命名期限"
   If strSrvDate(1) >= FCP案件命名啟用日 Then
        If bLvlChm = True Or bLvlJpn = True Or bLvlMot = True Or bLvlEls = True Then  '未分工程師組別
            If bLvlEls = True Then
               strExc(2) = "1"
            ElseIf bLvlChm = True Then
               strExc(2) = "2"
            ElseIf bLvlJpn = True Then
               strExc(2) = "3"
            Else
               strExc(2) = "4"
            End If
            'Add By Sindy 2021/4/22 待命名未分工程師組別,彈給五級主管看
            stSQL = ""
            If bLvl5 = True Then
            '2021/4/22 END
               stSQL = "SELECT '" & strUserNum & "','L' EV1,' ' EV2,CP09,FA10,TCT10 AS CP14,CP13,NVL(TCT02,WORKDAYADD(2,CP66)) AS CP06,CP07,CP48,NULL,0 " & _
                       "FROM TRANSCASETITLE, CASEPROGRESS, PATENT, FAGENT " & _
                       "WHERE NVL(TCT04,'N')='N' AND TCT01=CP09(+) " & _
                       "AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
                       "AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) "
            End If
            If stSQL <> "" Then stSQL = stSQL & " UNION ALL "
            stSQL = stSQL & "SELECT '" & strUserNum & "','L' EV1,' ' EV2,CP09,FA10,TCT10 AS CP14,CP13,NVL(TCT02,WORKDAYADD(2,CP66)) AS CP06,CP07,CP48,NULL,0 " & _
                    "FROM TRANSCASETITLE, CASEPROGRESS, PATENT, FAGENT " & _
                    "WHERE NVL(TCT10,'N')='N' AND TCT01=CP09(+) " & _
                    "AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
                    "AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND PA150=" & CNULL(strExc(2))
            stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12) " & stSQL
            cnnConnection.Execute stSQL, intI
        End If
        
        strExc(2) = CompWorkDay(3, strSrvDate(1), 1) '系統日-2工作天
        If stDept = "F21" Then  '外專工程師
             stSQL = "SELECT '" & strUserNum & "','L' EV1,' ' EV2,CP09,FA10,TCT10 AS CP14,CP13,NVL(TCT02,WORKDAYADD(2,CP66)) AS CP06,CP07,CP48,NULL,0 " & _
                     "FROM TRANSCASETITLE, CASEPROGRESS, PATENT, FAGENT " & _
                     "WHERE TCT10='" & stUserID & "' AND NVL(TCT05,0)=0 AND TCT01=CP09(+) " & _
                     "AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
                     "AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) "
             '2級只看自己及部屬資料
              stSQL = stSQL & "UNION ALL SELECT '" & strUserNum & "','L' EV1,' ' EV2,CP09,FA10,TCT10 AS CP14,CP13,NVL(TCT02,WORKDAYADD(2,CP66)) AS CP06,CP07,CP48,NULL,0 " & _
                                "FROM TRANSCASETITLE, CASEPROGRESS, PATENT, FAGENT,STAFF " & _
                                "WHERE NVL(TCT05,0)=0 AND TCT10=ST01 AND ST52='" & stUserID & "' AND TCT01=CP09(+) " & _
                                "AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
                                "AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) "
              If Trim(stNumList1(3)) <> "" Then
                 '3級以上逾期2天資料
                 stSQL = stSQL & "UNION ALL SELECT '" & strUserNum & "','L' EV1,' ' EV2,CP09,FA10,TCT10 AS CP14,CP13,NVL(TCT02,WORKDAYADD(2,CP66)) AS CP06,CP07,CP48,NULL,0 " & _
                                   "FROM TRANSCASETITLE, CASEPROGRESS, PATENT, FAGENT,STAFF " & _
                                   "WHERE NVL(TCT05,0)=0 AND TCT10=ST01 AND (ST53='" & stUserID & "' OR ST54='" & stUserID & "' OR ST55='" & stUserID & "')  AND TCT01=CP09(+) " & _
                                   "AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
                                   "AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) " & _
                                   "AND NVL(TCT02,WORKDAYADD(2,CP66))<=" & strExc(2)
              End If
            stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12) " & stSQL
            cnnConnection.Execute stSQL, intI
        ElseIf stDept = "F23" Then '業務承辦組
            '2級只看自己及部屬資料
            stSQL = "SELECT '" & strUserNum & "','L' EV1,' ' EV2,CP09,FA10,TCT10 AS CP14,CP13,NVL(TCT02,WORKDAYADD(2,CP66)) AS CP06,CP07,CP48,NULL,0 " & _
                    "FROM TRANSCASETITLE, CASEPROGRESS, PATENT, FAGENT, NATION " & _
                    "WHERE NVL(TCT05,0)=0 AND TCT01=CP09(+) " & _
                    "AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
                    "AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) " & _
                    "AND FA10=NA01(+) AND decode(pa75," & midStr & ",na51) IN (" & stNumList1(1) & IIf(Trim(stNumList1(2)) <> "", "," & stNumList1(2), "") & ") "
            If Trim(stNumList1(3)) <> "" Then
                '3級以上逾期2天資料
                stSQL = stSQL & "UNION SELECT '" & strUserNum & "','L' EV1,' ' EV2,CP09,FA10,TCT10 AS CP14,CP13,NVL(TCT02,WORKDAYADD(2,CP66)) AS CP06,CP07,CP48,NULL,0 " & _
                        "FROM TRANSCASETITLE, CASEPROGRESS, PATENT, FAGENT, NATION " & _
                        "WHERE NVL(TCT05,0)=0 AND TCT01=CP09(+) " & _
                        "AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL " & _
                        "AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) " & _
                        "AND FA10=NA01(+) AND decode(pa75," & midStr & ",na51) IN (" & stNumList & ") " & _
                        "AND NVL(TCT02,WORKDAYADD(2,CP66))<=" & strExc(2)
            End If
            stSQL = "INSERT INTO R060204(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12) " & stSQL
            cnnConnection.Execute stSQL, intI
        End If
   End If
   'end 2017/11/29
   
   'Add By Sindy 2021/4/22 未請款的本所期限為承辦期限(管控日期)再＋5個工作天
   stSQL = "UPDATE R060204 SET R08=WORKDAYADD(6,R10) WHERE R01='" & strUserNum & "' AND R02='F'"
   cnnConnection.Execute stSQL, intI
   '2021/4/22 END
   
   'Add By Sindy 2021/8/9 達核稿,因進度裡承辦人不是工程師,例如F5588.舜禹翻譯就不會判斷到達本所,增加其判斷
   'R06:承辦,R07:業務,R08:所限,R09:法限,R10:辦限,R11:核限
   stSQL = "UPDATE R060204 SET R02='A'" & _
            " WHERE R01||R04||R02 IN(" & _
            "SELECT R01||R04||R02 FROM R060204,caseprogress WHERE cp09(+)=R04" & _
            " AND r06<>cp14 AND r01='" & strUserNum & "' AND R02='C'" & _
            Replace(UCase(stConCP06), "CP06", "R08") & ")"
   cnnConnection.Execute stSQL, intI
   '2021/8/9 END
   
   'Add By Sindy 2015/10/22
   '程序組的期限彈跳排除:
   'modiby by sonia 此段原在下面2019/5/22搬上來, 排除C類來函(承辦人為工程師之審查意見,核駁等)逾法定期限之案件之前,但FCP-050839過期未發文否則下一程序也不會出現
   If stDept = "F22" Then
      '在前頭sql裡就過濾掉了
'      '排除926.核對已准專利,945.電話連絡單
'      stSQL = "DELETE R060204 WHERE R01='" & strUserNum & "'" & _
'         " AND EXISTS(SELECT * FROM caseprogress WHERE cp09=R04 AND cp10 in('926','945'))"
'      cnnConnection.Execute stSQL, intI
'      '排除R02='F' and cp27>0 已發文未請款
'      stSQL = "DELETE R060204 WHERE R01='" & strUserNum & "' AND R02='F'" & _
'         " AND EXISTS(SELECT * FROM caseprogress WHERE cp09=R04 AND cp27>0)"
'      cnnConnection.Execute stSQL, intI
      '排除C類來函(承辦人為工程師之審查意見,核駁等)逾法定期限之案件
      stSQL = "DELETE R060204 WHERE R01='" & strUserNum & "' AND substr(R04,1,1)='C'" & _
                                          " AND R06 is not null" & _
                                          " AND R09 is not null AND R09<" & strSrvDate(1) & _
         " AND EXISTS(SELECT * FROM staff WHERE R06=ST01 AND ST15='F21')"
      cnnConnection.Execute stSQL, intI
   End If
   '2015/10/22 END
      
   '刪除E.未發文
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02='E'" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02<>'E' AND R2.R02<>'D')"
   cnnConnection.Execute stSQL, intI
   
   '刪除(已達本所'A'則刪除達承辦'B'及達核稿'C')
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02 IN ('B','C')" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='A')"
   cnnConnection.Execute stSQL, intI
   
   '*****
   'Modify By Sindy 2017/1/13 事件為'B.達承辦'、'A.達本所'、'E.未發文'抓未發文資料的語法，都要剔除"通知告准" (1917)
   '*****
   'R04:總收文號
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02 IN ('A','B','E')" & _
      " AND EXISTS(SELECT * FROM R060204 R2,caseprogress WHERE R2.R01=R1.R01 AND R2.R02=R1.R02 AND R2.R04=R1.R04 AND R2.R04=cp09(+) and cp158=0 and cp10='1917')"
   cnnConnection.Execute stSQL, intI
   '2017/1/13 END
   
   'Added by Lydia 2019/05/31 程序大項工作整批發文: 排除整批發文的案件性質
   If stDept = "F22" Then
        stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' " & _
           " AND EXISTS(SELECT * FROM R060204 R2,caseprogress WHERE R2.R01=R1.R01 AND R2.R02=R1.R02 AND R2.R04=R1.R04 AND R2.R04=cp09(+) and cp158=0 and cp158=0 and cp10 in ('1603','1229','1604','1605') )"
        cnnConnection.Execute stSQL, intI
        'Added by Lydia 2021/08/26 刪除'A.達本所','N=達指定'同時為承辦人和管制人時，保留管制人資料; ex.FCP065559(AB0035398)、FCP065558(AB0035395)因為承辦人和管制人都屬於Phoebe的下屬，造成後面更新備註語法錯誤
        stSQL = "DELETE From R060204 R1 Where R1.R01='" & strUserNum & "' And R1.R02='A' And R03='1' And Exists(" & _
                     "Select * From R060204 R2 Where R2.R01=R1.R01 And R2.R04=R1.R04  And R1.R02='A' And R03='2' ) "
        cnnConnection.Execute stSQL, intI
        stSQL = "DELETE From R060204 R1 Where R1.R01='" & strUserNum & "' And R1.R02='N' And R03='1' And Exists(" & _
                     "Select * From R060204 R2 Where R2.R01=R1.R01 And R2.R04=R1.R04  And R1.R02='N' And R03='2' ) "
        cnnConnection.Execute stSQL, intI
        'end 2021/08/26
   End If
   
'cancel by sonia 2019/5/22 搬到上面
'   'Add By Sindy 2015/10/22
'   '程序組的期限彈跳排除:
'   If stDept = "F22" Then
'      '在前頭sql裡就過濾掉了
''      '排除926.核對已准專利,945.電話連絡單
''      stSQL = "DELETE R060204 WHERE R01='" & strUserNum & "'" & _
''         " AND EXISTS(SELECT * FROM caseprogress WHERE cp09=R04 AND cp10 in('926','945'))"
''      cnnConnection.Execute stSQL, intI
''      '排除R02='F' and cp27>0 已發文未請款
''      stSQL = "DELETE R060204 WHERE R01='" & strUserNum & "' AND R02='F'" & _
''         " AND EXISTS(SELECT * FROM caseprogress WHERE cp09=R04 AND cp27>0)"
''      cnnConnection.Execute stSQL, intI
'      '排除C類來函(承辦人為工程師之審查意見,核駁等)逾法定期限之案件
'      stSQL = "DELETE R060204 WHERE R01='" & strUserNum & "' AND substr(R04,1,1)='C'" & _
'                                          " AND R06 is not null" & _
'                                          " AND R09 is not null AND R09<" & strSrvDate(1) & _
'         " AND EXISTS(SELECT * FROM staff WHERE R06=ST01 AND ST15='F21')"
'      cnnConnection.Execute stSQL, intI
'   End If
'   '2015/10/22 END
'end 2019/5/22
   
   'Added by Lydia 2020/11/19 因為達本所和客戶提供文件會有重複記錄的情況,所以刪除非客戶提供文件=M; ex.Sharon看到Gill的客戶文件是紅字
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02<>'M'" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='M')"
   cnnConnection.Execute stSQL, intI
   'end 2020/11/19
   
   'Add By Sindy 2017/9/11
   '若同案有 'A達本所'期限,其他的就不要再顯示
   'Modify By Sindy 2021/4/21 N.達指定要另外判斷
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02 not in('A','N')" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='A')"
   cnnConnection.Execute stSQL, intI
   '更新備註
   stSQL = "UPDATE R060204 R1 SET R13='達指定;'||(SELECT R13 FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='N')" & _
      " WHERE R1.R01='" & strUserNum & "' AND R1.R02='A'" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='N' AND R2.R14>0 AND R2.R08>0 AND R2.R08>=R2.R14)" & _
      " AND R1.R08>0 AND R1.R14>0"
   cnnConnection.Execute stSQL, intI
   stSQL = "UPDATE R060204 R1 SET R13='達本所;'||R13" & _
      " WHERE R1.R01='" & strUserNum & "' AND R1.R02='N'" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='A' AND R2.R14>0 AND R2.R08>0 AND R2.R08<R2.R14)" & _
      " AND R1.R08>0 AND R1.R14>0"
   cnnConnection.Execute stSQL, intI
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02='N'" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='A' AND R2.R14>0 AND R2.R08>0 AND R2.R08>=R2.R14)"
   cnnConnection.Execute stSQL, intI
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02='A'" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='N' AND R2.R14>0 AND R2.R08>0 AND R2.R08<R2.R14)"
   cnnConnection.Execute stSQL, intI
   '2017/9/11 END
   'Add By Sindy 2021/12/27 同一個本所案號B.達承辦N.達指定同時出現時,達指定優先
   stSQL = "DELETE R060204 R1 WHERE R01='" & strUserNum & "' AND R02='B'" & _
      " AND EXISTS(SELECT * FROM R060204 R2 WHERE R2.R01=R1.R01 AND R2.R04=R1.R04 AND R2.R02='N')"
   cnnConnection.Execute stSQL, intI
   '2021/12/27 END
   
   'Added by Lydia 2021/08/23 工程師所有未發文查詢，全部顯示; 86019做「所有未發文」的查詢，其中FCP-65378的實審AB0032149同時存在早收文和未發文，而早收文負責查看人員=FCP程序管制在SetColor有判斷負責人本人或第二級才看
   bolShowAll = False
   If stDept = "F21" And idx = 1 Then
       bolShowAll = True
   End If
   'end 2021/08/23
   
   'Added by Lydia 2018/02/09 設sort排序
   '外專程序要求: "未請款"放在表單中倒數第1;客戶提供文件(待處理)放在倒數第2
   If stDept = "F22" Then
       stOrdCon = ",decode(R02,'I',0,'J',0,'K',0,'F',9,'M',5,1) sort "
   Else
       stOrdCon = ",decode(R02,'I',0,'J',0,'K',0,1) sort "
   End If
   'end 2018/02/08
   
   'Added by Lydia 2015/09/09 + H=早收文
   'Add By Sindy 2015/10... +,NVL(DL.Cnt,0) Cnt,CP43,R12
   'Add By Sindy 2015/11/26 抓延期次數
   'Modify By Sindy 2016/3/11 + FCP-46153 104/12/14尚未收申復
   '未延期:增加判斷若做CP43去檢查是否有延期時,還要過濾掉CP09不可為C類 ==> and cp09<'C'
   stVTB = "select DL01,sum(DLCnt) Cnt from" & _
           " (select DL01,count(*) DLCnt from R060204,datelimit" & _
           " where R01='" & strUserNum & "' and R04=DL01 group by DL01" & _
           " Union all " & _
           " select cp09 DL01,count(*) DLCnt from R060204,caseprogress,datelimit" & _
           " where R01='" & strUserNum & "' and R04=cp09(+) and cp43 is not null and cp43=DL01 and cp09<'C' group by cp09)" & _
           " group by DL01"
   '2015/11/26 END
   'Modify By Sindy 2015/12/15 + CP142
   'Modify By Sindy 2017/1/18 + ,'I','准未請款','J','分割建議','K','通知告准'
   '                            ,decode(R02,'I',0,'J',0,'K',0,1) sort
   'Modified by Lydia 2017/11/29 + 'L','待命名期限'
   'Modified by Lydia 2018/02/08  + 'M','待處理' ; 改排序decode(R02,'I',0,'J',0,'K',0,1) sort=> stOrdCon
   'Add By Sindy 2018/5/31 FCP-047800資料重覆 + R03 ==> '' as R03
   'Modified by Lydia 2018/11/13 因為SetColor無法判斷到本人和2級主管,所以'' as R03還原抓R03
   'Modify By Sindy 2021/4/28 + ,'' 約定期限
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非臺灣000抓CPM04; DECODE(INSTR('020,013,044',PA09),0,CPM03,CPM04) => DECODE(PA09,'000',CPM03,CPM04)
   'Modified by Lydia 2022/11/03 改判斷非臺灣000改抓FMP案管制人
   'strSql = "SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",'' 約定期限,NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限,NVL(lpad(SQLDateT(R11),10,' '),'2') 核稿期限,decode(cp158,0,decode(cp118,null,'','Y'),'') 電,S1.ST02 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號" & _
      ",DECODE(PA09,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||CP64 案件備註,NVL(PA05,NVL(PA06,PA07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,NA16 ,R06,R07,CP01,CP02,CP03,CP04,'' 未收款,CP10,CP27,CP09,decode(CP01,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,CP43,R12,CP142,CP164 " & stOrdCon & _
      " FROM R060204,CASEPROGRESS,NATION,PATENT,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,(" & stVTB & ") DL" & _
      " WHERE R01='" & strUserNum & "' AND R02<>'D' AND DL.DL01(+)=R04 AND CP09(+)=R04 AND NA01(+)=R05" & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=CP01 AND CPM02(+)=CP10"
   strSql = "SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",'' 約定期限,NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限,NVL(lpad(SQLDateT(R11),10,' '),'2') 核稿期限,decode(cp158,0,decode(cp118,null,'','Y'),'') 電,DECODE(PA09,'000',S1.ST02,NVL(S4.ST02,S1.ST02)) 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號" & _
      ",DECODE(PA09,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||CP64 案件備註,NVL(PA05,NVL(PA06,PA07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,DECODE(PA09,'000',NA16,NVL(NA79,NA16)) NA16 ,R06,R07,CP01,CP02,CP03,CP04,'' 未收款,CP10,CP27,CP09,decode(CP01,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,CP43,R12,CP142,CP164 " & stOrdCon & _
      " FROM R060204,CASEPROGRESS,NATION,PATENT,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,(" & stVTB & ") DL, STAFF S4" & _
      " WHERE R01='" & strUserNum & "' AND R02<>'D' AND DL.DL01(+)=R04 AND CP09(+)=R04 AND NA01(+)=R05" & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND S4.ST01(+)=NA79"

   'Add By Sindy 2015/10/22 +法務進度
   'Modified by Lydia 2018/02/08  + 'M','待處理' ;  改排序decode(R02,'I',0,'J',0,'K',0,1) sort=> stOrdCon
   'Modified by Sindy 2019/10/4 因為SetColor無法判斷到本人和2級主管,所以'' as R03還原抓R03
   'Modify By Sindy 2021/4/28 + ,'' 約定期限
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非臺灣000抓CPM04; DECODE(INSTR('020,013,044',LC15),0,CPM03,CPM04) => DECODE(LC15,'000',CPM03,CPM04)
   strSql = strSql & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",'' 約定期限,NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限,NVL(lpad(SQLDateT(R11),10,' '),'2') 核稿期限,decode(cp158,0,decode(cp118,null,'','Y'),'') 電,S1.ST02 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號" & _
      ",DECODE(LC15,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||CP64 案件備註,NVL(LC05,NVL(LC06,LC07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,NA16,R06,R07,CP01,CP02,CP03,CP04,'' 未收款,CP10,CP27,CP09,decode(CP01,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,CP43,R12,CP142,CP164 " & stOrdCon & _
      " FROM R060204,CASEPROGRESS,NATION,Lawcase,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,(" & stVTB & ") DL" & _
      " WHERE R01='" & strUserNum & "' AND R02<>'D' AND DL.DL01(+)=R04 AND CP09(+)=R04 AND NA01(+)=R05" & _
      " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=CP01 AND CPM02(+)=CP10"
   '2015/10/22 END
   
   'Modified by Lydia 2018/02/08  + 'M','待處理'  ;  改排序decode(R02,'I',0,'J',0,'K',0,1) sort=> stOrdCon
   'Modified by Sindy 2019/10/4 因為SetColor無法判斷到本人和2級主管,所以'' as R03還原抓R03
   'Modify By Sindy 2021/4/28 + ,'' 約定期限
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非臺灣000抓CPM04; DECODE(INSTR('020,013,044',SP09),0,CPM03,CPM04) => DECODE(SP09,'000',CPM03,CPM04)
   'Modified by Lydia 2022/11/03 改判斷非臺灣000改抓FMP案管制人
   'strSql = strSql & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",'' 約定期限,NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限,NVL(lpad(SQLDateT(R11),10,' '),'2') 核稿期限,decode(cp158,0,decode(cp118,null,'','Y'),'') 電,S1.ST02 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號" & _
      ",DECODE(SP09,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||CP64 案件備註,NVL(SP05,NVL(SP06,SP07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,NA16,R06,R07,CP01,CP02,CP03,CP04,'' 未收款,CP10,CP27,CP09,decode(CP01,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,CP43,R12,CP142,CP164 " & stOrdCon & _
      " FROM R060204,CASEPROGRESS,NATION,SERVICEPRACTICE,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,(" & stVTB & ") DL" & _
      " WHERE R01='" & strUserNum & "' AND R02<>'D' AND DL.DL01(+)=R04 AND CP09(+)=R04 AND NA01(+)=R05" & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=CP01 AND CPM02(+)=CP10"
   strSql = strSql & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",'' 約定期限,NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限,NVL(lpad(SQLDateT(R11),10,' '),'2') 核稿期限,decode(cp158,0,decode(cp118,null,'','Y'),'') 電,DECODE(SP09,'000',S1.ST02,NVL(S4.ST02,S1.ST02)) 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號" & _
      ",DECODE(SP09,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||CP64 案件備註,NVL(SP05,NVL(SP06,SP07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,DECODE(SP09,'000',NA16,NVL(NA79,NA16)) NA16,R06,R07,CP01,CP02,CP03,CP04,'' 未收款,CP10,CP27,CP09,decode(CP01,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,CP43,R12,CP142,CP164 " & stOrdCon & _
      " FROM R060204,CASEPROGRESS,NATION,SERVICEPRACTICE,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,(" & stVTB & ") DL,STAFF S4" & _
      " WHERE R01='" & strUserNum & "' AND R02<>'D' AND DL.DL01(+)=R04 AND CP09(+)=R04 AND NA01(+)=R05" & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND S4.ST01(+)=NA79"
   'Modified by Lydia 2018/02/08  + 'M','待處理' ;  改排序decode(R02,'I',0,'J',0,'K',0,1) sort=> stOrdCon
   'modify by sonia 2019/5/22 未收文期限若為C類來函且C類未發文則於事件'未收文'前加*
   'Modified by Sindy 2019/10/4 因為SetColor無法判斷到本人和2級主管,所以'' as R03還原抓R03
   'Modify By Sindy 2021/4/28 + ,NVL(lpad(SQLDateT(R15),10,' '),'2') 約定期限
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非臺灣000抓CPM04; DECODE(INSTR('020,013,044',PA09),0,CPM03,CPM04) => DECODE(PA09,'000',CPM03,CPM04)
   'Modified by Lydia 2022/11/03 改判斷非臺灣000改抓FMP案管制人
   'strSql = strSql & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",NVL(lpad(SQLDateT(R15),10,' '),'2') 約定期限,NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限,NVL(lpad(SQLDateT(R11),10,' '),'2') 核稿期限,' ' 電,S1.ST02 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(SUBSTR(NP01,1,1)||CP27,'C','*')||DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號" & _
      ",DECODE(PA09,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||NP15 案件備註,NVL(PA05,NVL(PA06,PA07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,NA16,R06,R07,NP02,NP03,NP04,NP05,'' 未收款,NP07,0+'' CP27,np01 CP09,decode(NP02,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,'' CP43,R12,0 CP142,CP164 " & stOrdCon & _
      " FROM R060204,NEXTPROGRESS,NATION,PATENT,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CASEPROGRESS,(" & stVTB & ") DL" & _
      " WHERE R01='" & strUserNum & "' AND R02='D' AND DL.DL01(+)=R04 AND NP01(+)=R04 AND NP22(+)=R12 AND NA01(+)=R05" & _
      " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=NP02 AND CPM02(+)=NP07 AND NP01=CP09(+)"
   strSql = strSql & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",NVL(lpad(SQLDateT(R15),10,' '),'2') 約定期限,NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限,NVL(lpad(SQLDateT(R11),10,' '),'2') 核稿期限,' ' 電,DECODE(PA09,'000',S1.ST02,NVL(S4.ST02,S1.ST02)) 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(SUBSTR(NP01,1,1)||CP27,'C','*')||DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號" & _
      ",DECODE(PA09,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||NP15 案件備註,NVL(PA05,NVL(PA06,PA07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,DECODE(PA09,'000',NA16,NVL(NA79,NA16)) NA16,R06,R07,NP02,NP03,NP04,NP05,'' 未收款,NP07,0+'' CP27,np01 CP09,decode(NP02,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,'' CP43,R12,0 CP142,CP164 " & stOrdCon & _
      " FROM R060204,NEXTPROGRESS,NATION,PATENT,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CASEPROGRESS,(" & stVTB & ") DL, STAFF S4" & _
      " WHERE R01='" & strUserNum & "' AND R02='D' AND DL.DL01(+)=R04 AND NP01(+)=R04 AND NP22(+)=R12 AND NA01(+)=R05" & _
      " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=NP02 AND CPM02(+)=NP07 AND NP01=CP09(+) AND S4.ST01(+)=NA79"
   'Add By Sindy 2015/10/22 +法務下一程序
   'Modified by Lydia 2018/02/08  + 'M','待處理' ;  改排序decode(R02,'I',0,'J',0,'K',0,1) sort=> stOrdCon
   'modify by sonia 2019/5/22 未收文期限若為C類來函且C類未發文則於事件'未收文'前加*
   'Modified by Sindy 2019/10/4 因為SetColor無法判斷到本人和2級主管,所以'' as R03還原抓R03
   'Modify By Sindy 2021/4/28 + ,NVL(lpad(SQLDateT(R15),10,' '),'2') 約定期限
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非臺灣000抓CPM04; DECODE(INSTR('020,013,044',LC15),0,CPM03,CPM04) => DECODE(LC15,'000',CPM03,CPM04)
   strSql = strSql & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",NVL(lpad(SQLDateT(R15),10,' '),'2') 約定期限,NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限,NVL(lpad(SQLDateT(R11),10,' '),'2') 核稿期限,' ' 電,S1.ST02 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(SUBSTR(NP01,1,1)||CP27,'C','*')||DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號" & _
      ",DECODE(LC15,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||NP15 案件備註,NVL(LC05,NVL(LC06,LC07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,NA16,R06,R07,NP02,NP03,NP04,NP05,'' 未收款,NP07,0+'' CP27,np01 CP09,decode(NP02,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,'' CP43,R12,0 CP142,CP164 " & stOrdCon & _
      " FROM R060204,NEXTPROGRESS,NATION,Lawcase,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CASEPROGRESS,(" & stVTB & ") DL" & _
      " WHERE R01='" & strUserNum & "' AND R02='D' AND DL.DL01(+)=R04 AND NP01(+)=R04 AND NP22(+)=R12 AND NA01(+)=R05" & _
      " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05 AND LC01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=NP02 AND CPM02(+)=NP07 AND NP01=CP09(+)"
   '2015/10/22 END
   
   'Modified by Lydia 2018/02/08  + 'M','待處理' ;  改排序decode(R02,'I',0,'J',0,'K',0,1) sort=> stOrdCon
   'modify by sonia 2019/5/22 未收文期限若為C類來函且C類未發文則於事件'未收文'前加*
   'Modified by Sindy 2019/10/4 因為SetColor無法判斷到本人和2級主管,所以'' as R03還原抓R03
   'Modify By Sindy 2021/4/28 + ,NVL(lpad(SQLDateT(R15),10,' '),'2') 約定期限
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非臺灣000抓CPM04; DECODE(INSTR('020,013,044',SP09),0,CPM03,CPM04) => DECODE(SP09,'000',CPM03,CPM04)
   'Modified by Lydia 2022/11/03 改判斷非臺灣000改抓FMP案管制人
   'strSql = strSql & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",NVL(lpad(SQLDateT(R15),10,' '),'2') 約定期限,NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限,NVL(lpad(SQLDateT(R11),10,' '),'2') 核稿期限,' ' 電,S1.ST02 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(SUBSTR(NP01,1,1)||CP27,'C','*')||DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號" & _
      ",DECODE(SP09,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||NP15 案件備註,NVL(SP05,NVL(SP06,SP07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,NA16,R06,R07,NP02,NP03,NP04,NP05,'' 未收款,NP07,0+'' CP27,np01 CP09,decode(NP02,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,'' CP43,R12,0 CP142,CP164 " & stOrdCon & _
      " FROM R060204,NEXTPROGRESS,NATION,SERVICEPRACTICE,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CASEPROGRESS,(" & stVTB & ") DL" & _
      " WHERE R01='" & strUserNum & "' AND R02='D' AND DL.DL01(+)=R04 AND NP01(+)=R04 AND NP22(+)=R12 AND NA01(+)=R05" & _
      " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=NP02 AND CPM02(+)=NP07 AND NP01=CP09(+)"
   strSql = strSql & " UNION" & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",NVL(lpad(SQLDateT(R15),10,' '),'2') 約定期限,NVL(lpad(SQLDateT(R10),10,' '),'2') 承辦期限,NVL(lpad(SQLDateT(R11),10,' '),'2') 核稿期限,' ' 電,DECODE(SP09,'000',S1.ST02,NVL(S4.ST02,S1.ST02)) 管制人,S2.ST02 承辦人" & _
      ",S3.ST02 智權人員,DECODE(SUBSTR(NP01,1,1)||CP27,'C','*')||DECODE(R02,'A','達本所','B','達承辦','C','達核稿','D','未收文','E','未發文','F','未請款','G','未交稿','H','早收文','I','准未請款','J','分割建議','K','通知告准','L','待命名期限','M','待處理','N','達指定','O','未分案','P','達約定') 事件" & _
      ",NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號" & _
      ",DECODE(SP09,'000',CPM03,CPM04) 案件性質,nvl(R13,'')||NP15 案件備註,NVL(SP05,NVL(SP06,SP07)) 案件名稱" & _
      ",R05 代理人國籍,R02,R03,DECODE(SP09,'000',NA16,NVL(NA79,NA16)) NA16,R06,R07,NP02,NP03,NP04,NP05,'' 未收款,NP07,0+'' CP27,np01 CP09,decode(NP02,'FCP',2,'FG',2,1) Srt1,NVL(DL.Cnt,0) Cnt,'' CP43,R12,0 CP142,CP164 " & stOrdCon & _
      " FROM R060204,NEXTPROGRESS,NATION,SERVICEPRACTICE,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CASEPROGRESS,(" & stVTB & ") DL,STAFF S4" & _
      " WHERE R01='" & strUserNum & "' AND R02='D' AND DL.DL01(+)=R04 AND NP01(+)=R04 AND NP22(+)=R12 AND NA01(+)=R05" & _
      " AND SP01(+)=NP02 AND SP02(+)=NP03 AND SP03(+)=NP04 AND SP04(+)=NP05 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=NP02 AND CPM02(+)=NP07 AND NP01=CP09(+) AND S4.ST01(+)=NA79"
   'Add by Amy 2013/07/03 增加案件命名追蹤
   'Add by Amy 2013/07/02 '查詢人部門為F23或bLvl5(最高主管68009)抓命名追蹤自輸入日期2個工作日或期限前2個工作天
   If GetST15(txtUsernum) = "F23" Or bLvl5 = True Then
      Dim Manage As String
      stDate2 = PUB_GetWorkDayAfterSysDate(CDbl(strSrvDate(1)), 2) + 19110000
      stDate3 = PUB_GetWorkDayAfterSysDate(CDbl(strSrvDate(1)), -2) + 19110000
      '抓個人建的或是管制人自輸入日期2個工作日或期限前2個工作天資料
      'Modify By Sindy 2017/2/16 + ,1 sort
      'Modify By Sindy 2021/4/28 + ,'' 約定期限
      strSql = strSql & _
                        " Union Select '' V,LPAD(SQLDateT(TCN02),9,' ') 本所期限,LPAD(SQLDateT(TCN02),9,' ')  法定期限,'' 約定期限,'' 承辦期限, '' 核稿期限,' ' 電,S1.ST02 管制人, '' 承辦人,S1.ST02 智權人員, '新案' 事件,''||TCN01 本所案號, '新案' 案件性質,TCN04 案件備註, '' 案件名稱, '' 代理人國籍, 'H', '1', '', '', '', '', '', '', '','' 未收款, NULL, 0+'' CP27, '' CP09, 0 Srt1,0 Cnt,'' CP43,0 R12,0 CP142,'' CP164,1 sort " & _
                        "From TrackingCaseName,Staff S1 Where  S1.ST01(+)=TCN03  And TCN05 is null And (TCN03='" & txtUsernum & "' OR TCN06='" & txtUsernum & "' ) " & _
                        "And ( TCN02 <=" & stDate2 & " OR TCN07 <= " & stDate3 & ")"
      Manage = CheckManage
      If Len(Manage) > 5 Then
         '登入者為2級且為3級以上主管
         '2級只看自己及部屬資料
           'Modify By Sindy 2017/2/16 + ,1 sort
           strSql = strSql & _
                           " Union Select '' V,LPAD(SQLDateT(TCN02),9,' ') 本所期限,LPAD(SQLDateT(TCN02),9,' ')  法定期限,'' 約定期限,'' 承辦期限, '' 核稿期限,' ' 電,S1.ST02 管制人, '' 承辦人,S1.ST02 智權人員, '新案' 事件,''||TCN01 本所案號, '新案' 案件性質,TCN04 案件備註, '' 案件名稱, '' 代理人國籍, 'H', '1', '', '', '', '', '', '', '','' 未收款, NULL, 0+'' CP27, '' CP09, 0 Srt1,0 Cnt,'' CP43,0 R12,0 CP142,'' CP164,1 sort " & _
                           "From TrackingCaseName,STAFF S1 Where S1.ST01(+)=TCN03  And TCN05 is null And ST52='" & txtUsernum & "' And ( TCN02 <=" & stDate2 & " OR TCN07 <= " & stDate3 & ")"
           '3級以上逾期資料
           strSql = strSql & _
                           " Union Select '' V,LPAD(SQLDateT(TCN02),9,' ') 本所期限,LPAD(SQLDateT(TCN02),9,' ')  法定期限,'' 約定期限,'' 承辦期限, '' 核稿期限,' ' 電,S1.ST02 管制人, '' 承辦人,S1.ST02 智權人員, '新案' 事件,''||TCN01 本所案號, '新案' 案件性質,TCN04 案件備註, '' 案件名稱, '' 代理人國籍, 'H', '1', '', '', '', '', '', '', '','' 未收款, NULL, 0+'' CP27, '' CP09, 0 Srt1,0 Cnt,'' CP43,0 R12,0 CP142,'' CP164,1 sort " & _
                           "From TrackingCaseName,STAFF S1 Where S1.ST01(+)=TCN03  And TCN05 is null And (ST53='" & txtUsernum & "' OR ST54='" & txtUsernum & "' OR ST55='" & txtUsernum & "' ) And TCN02 <=" & strSrvDate(1)
   
      ElseIf Manage = "ST52" Then
          '只是2級主管
          'Modify By Sindy 2017/2/16 + ,1 sort
          strSql = strSql & _
                            " Union Select '' V,LPAD(SQLDateT(TCN02),9,' ') 本所期限,LPAD(SQLDateT(TCN02),9,' ')  法定期限,'' 約定期限,'' 承辦期限, '' 核稿期限,' ' 電,S1.ST02 管制人, '' 承辦人,S1.ST02 智權人員, '新案' 事件,''||TCN01 本所案號, '新案' 案件性質,TCN04 案件備註, '' 案件名稱, '' 代理人國籍, 'H', '1', '', '', '', '', '', '', '','' 未收款, NULL, 0+'' CP27, '' CP09, 0 Srt1,0 Cnt,'' CP43,0 R12,0 CP142,'' CP164,1 sort " & _
                            "From TrackingCaseName,STAFF S1 Where S1.ST01(+)=TCN03  And TCN05 is null And ST52='" & txtUsernum & "'  And ( TCN02 <=" & stDate2 & " OR TCN07 <= " & stDate3 & ")"
    
      ElseIf Manage = "ST5X" Then
           '只是3級以上主管
           'Modify By Sindy 2017/2/16 + ,1 sort
           strSql = strSql & _
                            " Union Select '' V,LPAD(SQLDateT(TCN02),9,' ') 本所期限,LPAD(SQLDateT(TCN02),9,' ')  法定期限,'' 約定期限,'' 承辦期限, '' 核稿期限,' ' 電,S1.ST02 管制人, '' 承辦人,S1.ST02 智權人員, '新案' 事件,''||TCN01 本所案號, '新案' 案件性質,TCN04 案件備註, '' 案件名稱, '' 代理人國籍, 'H', '1', '', '', '', '', '', '', '','' 未收款, NULL, 0+'' CP27, '' CP09, 0 Srt1,0 Cnt,'' CP43,0 R12,0 CP142,'' CP164,1 sort " & _
                            "From TrackingCaseName,STAFF S1 Where S1.ST01(+)=TCN03  And TCN05 is null And (ST53='" & txtUsernum & "' OR ST54='" & txtUsernum & "' OR ST55='" & txtUsernum & "' ) And TCN02 <=" & strSrvDate(1)
      End If
   End If
   '2013/07/02 END
   
   'Added by Lydia 2019/05/31 指定排序
   If stDept = "F22" Then
       strSql = strSql & " order by sort asc, Srt1 asc, 本所期限 asc,管制人 asc,代理人國籍 asc,本所案號 asc"
   ElseIf stDept = "F23" Then
       strSql = strSql & " order by sort asc, 本所期限 asc,智權人員 asc,代理人國籍 asc,本所案號 asc"
   Else
       strSql = strSql & " order by sort asc, 本所期限 asc,承辦人 asc,本所案號 asc"
   End If
   'end 2019/05/31
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If RsTemp Is Nothing Then Exit Sub
   If RsTemp.RecordCount = 0 Then
      Set m_adoRst = RsTemp.Clone
      SetRst2Grid
      MsgBox "查無資料！", vbInformation
      cmdHide.Enabled = False
      LblTotCnt.Caption = "共 0 筆" 'Add By Sindy 2009/10/07
   Else
      'Modify by Amy 2014/06/05 +FormName
      Set m_adoRst = PUB_CreateRecordset(RsTemp, , , 300, Me.Name)
      '更新未收文未收款&未收款本所期限
      SetXRecord
      'Remove by Lydia 2019/05/31 點選本所案號排序會出Sort錯誤
'      Select Case stDept
'         'Modify By Sindy 2017/1/19 + sort asc
'         Case "F22" '程序
'            'Modified by Morgan 2012/5/28 +系統別FMP的排前面
'            m_stSort = "sort asc, Srt1 asc, 本所期限 asc,管制人 asc,代理人國籍 asc,本所案號 asc"
'         Case "F23" '業務
'            m_stSort = "sort asc, 本所期限 asc,智權人員 asc,代理人國籍 asc,本所案號 asc"
'         'F21,F81
'         Case Else
'            m_stSort = "sort asc, 本所期限 asc,承辦人 asc,本所案號 asc"
'      End Select
'      m_adoRst.Sort = m_stSort
      'end 2019/05/31
      SetRst2Grid
      SetGrid
      RecordShow
      
      SetColor
      cmdHide.Enabled = True
      m_blnColOrderAsc = True
   End If
   
   Set rsA = Nothing
End Sub
