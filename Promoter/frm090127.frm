VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090127 
   BorderStyle     =   1  '單線固定
   Caption         =   "查覆區"
   ClientHeight    =   6108
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6108
   ScaleWidth      =   8952
   Begin VB.CommandButton cmdState 
      BackColor       =   &H00FFFFC0&
      Caption         =   "重新分查名人"
      Height          =   280
      Index           =   4
      Left            =   3240
      Style           =   1  '圖片外觀
      TabIndex        =   44
      Top             =   1440
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   12
      Left            =   2600
      MaxLength       =   2
      TabIndex        =   13
      Top             =   1300
      Width           =   400
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   11
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   12
      Top             =   1300
      Width           =   300
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   10
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   11
      Top             =   1300
      Width           =   700
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   9
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   10
      Top             =   1300
      Width           =   500
   End
   Begin VB.CommandButton cmdSendMail 
      BackColor       =   &H00C0FFFF&
      Caption         =   "通知送件"
      Height          =   360
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   1440
      Width           =   900
   End
   Begin VB.CommandButton cmdMaster 
      Caption         =   "查名單輸入"
      Height          =   360
      Index           =   1
      Left            =   4723
      TabIndex        =   15
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton cmdState 
      Caption         =   "已發文"
      Height          =   360
      Index           =   3
      Left            =   6000
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdState 
      Caption         =   "已完成"
      Height          =   360
      Index           =   2
      Left            =   5160
      Style           =   1  '圖片外觀
      TabIndex        =   21
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdState 
      Caption         =   "處理中"
      Height          =   360
      Index           =   1
      Left            =   4320
      Style           =   1  '圖片外觀
      TabIndex        =   20
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdState 
      Caption         =   "未發文"
      Height          =   360
      Index           =   0
      Left            =   3480
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   30
      Width           =   800
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   8
      Left            =   4200
      TabIndex        =   9
      Top             =   996
      Width           =   2535
   End
   Begin VB.PictureBox G_SeekPicColor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Enabled         =   0   'False
      Height          =   300
      Left            =   9240
      ScaleHeight     =   21
      ScaleMode       =   3  '像素
      ScaleWidth      =   21
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox tmpPic 
      Height          =   4455
      Left            =   9120
      ScaleHeight     =   367
      ScaleMode       =   3  '像素
      ScaleWidth      =   295
      TabIndex        =   34
      Top             =   480
      Visible         =   0   'False
      Width           =   3588
      Begin VB.Image tmpImg 
         Height          =   1770
         Left            =   1425
         Stretch         =   -1  'True
         Top             =   1095
         Visible         =   0   'False
         Width           =   1890
      End
   End
   Begin VB.CommandButton cmdTo 
      BackColor       =   &H00C0FFC0&
      Caption         =   "收文"
      Height          =   360
      Left            =   5926
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   1440
      Width           =   900
   End
   Begin VB.CommandButton cmdMaster 
      Caption         =   "查名單"
      Height          =   360
      Index           =   0
      Left            =   6929
      TabIndex        =   17
      Top             =   1440
      Width           =   900
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   6
      Left            =   3120
      MaxLength       =   1
      TabIndex        =   25
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   5
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "0"
      Top             =   694
      Width           =   400
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   4
      Left            =   2280
      MaxLength       =   7
      TabIndex        =   6
      Top             =   694
      Width           =   855
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   3
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   5
      Top             =   694
      Width           =   855
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   2
      Left            =   2280
      MaxLength       =   7
      TabIndex        =   3
      Top             =   392
      Width           =   855
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   1
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   2
      Top             =   392
      Width           =   855
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   0
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   1
      Top             =   90
      Width           =   855
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   360
      Left            =   6840
      TabIndex        =   23
      Top             =   30
      Width           =   880
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "明細"
      Height          =   360
      Left            =   7935
      TabIndex        =   18
      Top             =   1440
      Width           =   900
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7800
      TabIndex        =   24
      Top             =   30
      Width           =   880
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm090127.frx":0000
      Height          =   3555
      Left            =   60
      TabIndex        =   32
      Top             =   2160
      Width           =   8835
      _ExtentX        =   15579
      _ExtentY        =   6265
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|申請編號|申請人ID|申請人|類別|委查單號|組群|中文筆數|英文筆數|圖形筆數|查名人ID|查名人|申請日期|期限日期|查覆日期|已讀(Y)"
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
      _Band(0).Cols   =   16
   End
   Begin MSForms.TextBox textCName 
      Height          =   300
      Left            =   1080
      TabIndex        =   8
      Top             =   990
      Width           =   2055
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "3625;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   90
      Width           =   1815
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "3201;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSname 
      Height          =   255
      Left            =   2040
      TabIndex        =   46
      Top             =   120
      Width           =   855
      Size            =   "1508;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Top             =   390
      Width           =   1635
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2884;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "下拉選擇"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      Caption         =   "結果：黃色表示進行覆核中；紅色表示已覆核，仍與本所近似。"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   390
      TabIndex        =   45
      Top             =   1890
      Width           =   5145
   End
   Begin VB.Label Label12 
      Caption         =   "本所案號："
      Height          =   240
      Left            =   120
      TabIndex        =   43
      Top             =   1330
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "送：送件狀態◎已通知 ●已送件"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   1530
      TabIndex        =   42
      Top             =   5850
      Width           =   2895
   End
   Begin VB.Label lblState 
      AutoSize        =   -1  'True
      Caption         =   "lblState"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   6960
      TabIndex        =   41
      Top             =   480
      Width           =   510
   End
   Begin VB.Label Label6 
      Caption         =   "狀態："
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6360
      TabIndex        =   40
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "(請以"",""或""."" 區隔)"
      Height          =   240
      Left            =   6840
      TabIndex        =   39
      Top             =   1026
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "委查單號："
      Height          =   240
      Left            =   3240
      TabIndex        =   38
      Top             =   1026
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "客戶名稱："
      Height          =   240
      Left            =   120
      TabIndex        =   37
      Top             =   1026
      Width           =   975
   End
   Begin MSForms.TextBox txtUnicode 
      Height          =   255
      Index           =   1
      Left            =   7800
      TabIndex        =   36
      Top             =   600
      Visible         =   0   'False
      Width           =   525
      VariousPropertyBits=   -1400879077
      MaxLength       =   50
      Size            =   "926;450"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "讀：附件已讀"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   90
      TabIndex        =   33
      Top             =   5850
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   2040
      X2              =   2160
      Y1              =   883
      Y2              =   883
   End
   Begin VB.Label Label7 
      Caption         =   "期限日期："
      Height          =   240
      Left            =   120
      TabIndex        =   31
      Top             =   724
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2160
      Y1              =   563
      Y2              =   563
   End
   Begin VB.Label Label5 
      Caption         =   "委查日期："
      Height          =   240
      Left            =   120
      TabIndex        =   30
      Top             =   422
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "查名人："
      Height          =   240
      Left            =   3240
      TabIndex        =   29
      Top             =   422
      Width           =   825
   End
   Begin VB.Label Label16 
      Caption         =   "雙擊選取時，開啟查覆明細"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   400
      TabIndex        =   28
      Top             =   1640
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "委查人："
      Height          =   240
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   825
   End
   Begin VB.Label Label2 
      Caption         =   "查名類別：              0.全部 1.文字 2.圖形"
      Height          =   240
      Left            =   3240
      TabIndex        =   26
      Top             =   724
      Width           =   3375
   End
End
Attribute VB_Name = "frm090127"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/01 改成Form2.0 ; Combo1、Combo2、lblSname、GRD1改字型=新細明體-ExtB、txtField(7)=>textCName
'Create by Lydia 2015/08/05 查名單-查覆區
Option Explicit
'設定可使用表單
Public Tmpfrm090126 As Form
Public Tmpfrm090128 As Form
Public contCusName As String '從接洽單傳申請人(中文名稱) ->客戶名稱
Dim iStiu As Integer  '狀態 :0查詢 1編輯
Dim R_type As String '使用者角色
Dim m_TMQApp As String '查名單申請號
Dim m_TMQNo As String '委查單號(分單)
Dim bolCont As Boolean '從聯絡單來選擇收文
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Dim dblPrevRow As Double
Dim mTQD01s As String '勾選的申請編號
Dim pTMQList As String 'Added by Lydia 2023/08/02 從接洽單傳入原本勾選的委查單(TQD02)
Dim m_AttachPath As String
Dim m_PrevForm As Form
Dim colMno As Integer, colAno As Integer '申請編號,委查單號的位置
Dim colCno As Integer '撤回
Dim colCp09 As Integer '總收文號
Public stKeyUser As String 'Added by Lydia 2016/04/28 從接洽單傳委查人
'Added by Lydia 2016/04/29
Dim colTMQ02 As Integer '委查人
Dim colState As Integer '送件狀態：◎已通知　●已送件
Dim colCP57 As Integer '取消收文
Dim colCase As Integer '本所案號
Dim colCP14 As Integer '承辦人
Dim colTMQ06 As Integer 'Added by Lydia 2016/05/04 期限日期
Dim colTMQ11 As Integer 'Added by Lydia 2016/07/07 查覆日期=查覆完畢
Dim colTMQ23 As Integer 'Added by Lydia 2021/02/17 覆核日期
Dim colTQD0609 As Integer 'Added by Lydia 2021/02/17 結果=TQD09覆核結果 > TQD06查名結果
Dim colShowCP09 As Integer 'Added by Lydia 2016/06/02
Dim strManUser As String 'Added by Lydia 2016/06/03 可看到所有查名人員
Dim mCaseNo(1 To 4) As String, mStiu  As String 'Added by Lydia 2018/09/20 傳入本所案號
Dim stIdList As String 'Added by Lydia 2019/08/12 創新業務組成員可操作清單(WXX部門的人可以操作自已部門所有人的資料,
                                                                                                                        '例W10所有人都可操作W1001，W20所有人都可操作W2001。
'Added by Lydia 2019/12/25 開放特殊設定權限
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim strGrpTmp1 As String, strGrpTmp2 As String 'Added by Lydia 2020/12/04
Dim strManUser2 As String 'Added by  Lydia 2021/11/12 內商查名覆核人員
Dim bolCaseRead As Boolean 'Added by Lydia 2022/01/06 從歷程過來看結果

'Modified by Lydia 2018/09/20 + fCaseNo,fStiu
Public Sub SetParent(ByRef fm As Form, Optional ByVal fCaseNo As String = "", Optional ByVal fStiu As String = "", Optional ByVal fPlist As String = "")
   Set m_PrevForm = fm
   'Modify By Sindy 2022/9/16 + Or m_PrevForm.Name = "frm090801_New"
   If m_PrevForm.Name = "frm090801" Or m_PrevForm.Name = "frm090801_New" Then bolCont = True
   'Added by Lydia 2018/09/20 傳入本所案號
   If fCaseNo <> "" Then
       Call ChgCaseNo(fCaseNo, mCaseNo)
   End If
   mStiu = fStiu
   'end 2018/09/20
    'Added by Lydia 2022/01/06 增加從歷程過來看結果的人員：只限該案件的資料
    bolCaseRead = False
    If mCaseNo(1) <> "" And TypeName(m_PrevForm) <> "Nothing" Then
        If m_PrevForm.Name = "frm090202_2" Then
            bolCaseRead = True
        End If
    End If
    'end 2022/01/06
   pTMQList = fPlist 'Added by Lydia 2023/08/02
End Sub
Public Function IsRolePlay(ByRef defRole As String, Optional ShowMsg As Boolean = True) As Boolean
   Dim tmpY As Integer
   
   IsRolePlay = True
    Select Case defRole
        Case "待查"
              strExc(1) = "select tmqm01 from tmqmember where tmqm01='" & strUserNum & "' "
              tmpY = 1
              Set RsTemp = ClsLawReadRstMsg(tmpY, strExc(1))
              'Modified by Lydia 2016/06/21 嘉雯(84027)不排查名單,但要可檢視
              'If tmpY = 1 Or InStr("67002,69008,84027", strUserNum) > 0 Or Pub_StrUserSt03 = "M51" Then
              'Move by Lydia 2016/08/01 移動
              strManUser = Pub_GetSpecMan("內商查名主管")
              'Added by Lydia 2021/11/12 林嘉雯請假時職代處理
              strExc(1) = GetDutyList(strManUser)
              If strExc(1) <> "" Then strManUser = strManUser & ";" & strExc(1)
              'end 2021/11/12
              
              If tmpY = 1 Or InStr(strManUser, strUserNum) > 0 Or Pub_StrUserSt03 = "M51" Then
                 R_type = "U"
              Else
                 IsRolePlay = False
              End If
        Case "覆核"  '覆核人員-商標處主管
              'Modified by Lydia 2016/03/28 目前是先由分案人員(嘉雯84027)做初步覆核,若需更進一步則交由商標處主管
              'Memo by Lydia 2016/06/27 近似本所案仍要申請,經相關主管同意後,請林經理(或職代)將結果改為"核可"
              'Memo by Lydia 2021/11/15 已通過電話與嘉雯確認，同時擁有覆核和核可權限，但是覆核和核可作業還是會分開確認(確認作業會發email)；若嘉雯請假則職代有相同權限。
              'Modified by Lydia 2021/11/12 林嘉雯請假職代處理
              'If InStr(Pub_GetSpecMan("內商查名覆核人員"), strUserNum) > 0 Or Pub_StrUserSt03 = "M51" Then
              strManUser2 = Pub_GetSpecMan("內商查名覆核人員")
              If InStr(strManUser2, strUserNum) = 0 And Left(Pub_StrUserSt03, 2) = "P2" Then
                  strExc(1) = GetDutyList(strManUser2)
                  If strExc(1) <> "" Then strManUser2 = strManUser2 & ";" & strExc(1)
              End If
              If InStr(strManUser2, strUserNum) > 0 Or Pub_StrUserSt03 = "M51" Then
              'end 2021/11/12
                 R_type = "M"
              Else
                 IsRolePlay = False
              End If
        Case "查覆"
                 R_type = "Q"
        Case "維護"
              If Pub_StrUserSt03 = "M51" Then
                 R_type = "A"
              Else
                 IsRolePlay = False
              End If
    End Select
    If IsRolePlay = False And ShowMsg = True Then MsgBox "無此使用權限...", , "警告!!"

End Function
'查詢明細資料
Private Sub cmdDetail_Click()
Dim tmpList As String
   
   If TypeName(Tmpfrm090128) <> "Nothing" And cmdDetail.Visible = True Then
        For i = 1 To GRD1.Rows - 1
           If GRD1.TextMatrix(i, 0) = "V" Then
              tmpList = tmpList & Trim(GRD1.TextMatrix(i, colAno)) & ","
           End If
        Next i

        If tmpList <> "" Then
           'Added by Lydia 2016/05/10 從接洽單->查覆區->明細
           If TypeName(m_PrevForm) <> "Nothing" Then
              Tmpfrm090128.mbolCall = True
           End If
           Tmpfrm090128.m_NoList = tmpList
           Tmpfrm090128.R_type = R_type
           Tmpfrm090128.iStiu = iStiu
           'Added by Lydia 2016/06/02
           If Trim(GRD1.TextMatrix(1, colCase)) = Trim(txtField(9)) & "-" & Trim(txtField(10)) & IIf(Val(Trim(txtField(11) & txtField(12))) = 0, "", "-" & Left(Trim(txtField(11)) & "0", 1) & Left(Trim(txtField(12)) & "00", 2)) Then
             Tmpfrm090128.ShowCP09 = Trim(GRD1.TextMatrix(1, colShowCP09))
           End If
           'end 2016/06/02
           Tmpfrm090128.SetParent Me
           Tmpfrm090128.m_NoIdx = 0
           Tmpfrm090128.Show
           If Tmpfrm090128.QueryData = True Then
             Me.Hide
           Else
             Unload Tmpfrm090128
           End If
        Else

        End If
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub
Private Sub cmdMaster_Click(Index As Integer)
  If Index = 0 Then
    If TypeName(Tmpfrm090126) <> "Nothing" Then
         For i = 1 To GRD1.Rows - 1
            If GRD1.TextMatrix(i, 0) = "V" Then
               Tmpfrm090126.mApNoList = Trim(GRD1.TextMatrix(i, colMno))
               Tmpfrm090126.SetParent Me
               Tmpfrm090126.Show
               Me.Hide
               Exit For
            End If
         Next i
    End If
  ElseIf Index = 1 Then
    If TypeName(Tmpfrm090126) <> "Nothing" Then
       Tmpfrm090126.SetParent Me
       Tmpfrm090126.Show
       Me.Hide
    End If

  End If
End Sub

Private Sub cmdQuery_Click()
Dim Cancel As Boolean
   If txtField(1) > txtField(2) Then
      MsgBox "起始期間不可大於終止期間!", vbCritical, "輸入錯誤"
      txtField(1).SetFocus
      Exit Sub
   End If
   If txtField(3) > txtField(4) Then
      MsgBox "起始期間不可大於終止期間!", vbCritical, "輸入錯誤"
      txtField(3).SetFocus
      Exit Sub
   End If
   txtField_Validate 6, Cancel
   If Cancel = True Then Exit Sub
   
   m_TMQApp = "" '清空收文判斷
   If QueryData = False Then ShowNoData
End Sub
'Added by Lydia 2016/06/30 重新分派查名人員
Private Sub GetNewTMQ10()
Dim strQ1 As String, strQ2 As String
Dim strQuser As String
Dim RsQ As New ADODB.Recordset
Dim intQ As Integer
Dim bolUpd As Boolean
Dim iCnt As Integer 'Added by Lydia 2017/06/23
'Added by Lydia 2017/11/07
Dim bolUpTmq06 As Boolean
Dim Inputtm As String, InputWDay As String
Dim chkAllStatus As String 'Added by Lydia 2018/05/25 內商查名單分單狀態：若查名中心聯絡開始不分單將狀態改為N，恢復分單將狀態改為Y

   strQ1 = ""
   chkAllStatus = Pub_GetSpecMan("內商查名單分單狀態") 'Added by Lydia 2018/05/25
   
   'Modified by Lydia 2019/05/29 +查名單號
   'If InStr(Combo1.Text, "全部") > 0 Then
   If InStr(Combo1.Text, "全部") > 0 And Trim(txtField(8)) = "" Then
      'Modified by Lydia 2017/06/23
      'MsgBox "請指定需要重新分派的查名人!"
      'Exit Sub
      If MsgBox("未指定重新分派的查名人，要改抓目前所有未分派的查名單？" & vbCrLf & "繼續作業按是，要重選查名人請按否。", vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      'end 2017/06/23
   Else
      strQ1 = strQ1 & " and tmq10='" & Trim(Left("" & Combo1.Text, 6)) & "'"
      'Move by Lydia 2017/06/23 從下面移上來
      If txtField(1) & txtField(2) & txtField(8) = "" Then
         MsgBox "請輸入委查期間或委查單號!"
         Exit Sub
      End If
   End If
   
      '抓委查期間和委查單號
      'Mark by Lydia 2017/06/23
      'If txtField(1) & txtField(2) & txtField(8) = "" Then
      '   MsgBox "請輸入委查期間或委查單號!"
      '   Exit Sub
      'Else
           'Modified by Lydia 2017/01/25 tmq04->tmq05 (收件分發日期)
           If txtField(1) <> "" Then strQ1 = strQ1 & " and tmq05>=" & TransDate(Trim(txtField(1)), 2)
           If txtField(2) <> "" Then strQ1 = strQ1 & " and tmq05<=" & TransDate(Trim(txtField(2)), 2)
           If txtField(8) <> "" Then strQ1 = strQ1 & " and tmq01 in (" & GetAddStr(Replace(txtField(8).Text, ".", ",")) & ")"
           'Added by Lydia 2017/06/23
           If txtField(0) <> "" Then strQ1 = strQ1 & " and tmq02='" & Trim(txtField(0)) & "' "
           '客戶名稱
           If textCName <> "" Then
              strExc(1) = Replace(Trim(textCName), " ", "%")
              strExc(1) = Replace(strExc(1), "%%", "%")
              strQ1 = strQ1 & " and upper(tqa04) like '%" & UCase(strExc(1)) & "%'"
           End If
           'end 2017/06/23
      'End If
        
   'Added by Lydia 2017/06/23 抓目前未分派的查名單
   If InStr(Combo1.Text, "全部") > 0 Then
        strQ1 = Replace(UCase(strQ1), "TMQ05", "TMQ04") '可能沒有日期
        If MsgBox("是否要排除今天下午6點送出的查名單?", vbInformation + vbYesNo) = vbYes Then
           strQ1 = strQ1 & " AND TMQ13<=" & strSrvDate(1) & " AND TMQ14 < 1800 "
        End If
        'Added by Lydia 2019/05/29 指定查名單
        If Trim(txtField(8)) <> "" Then
            strQ2 = "select a.* from trademarkquery a,tmqapp b where tmq18=tqa01(+) and tmq11 is null and tmq01 in (" & GetAddStr(Replace(txtField(8).Text, ".", ",")) & ") order by tmq01"
        Else
        'end 2019/05/29
            strQ2 = "select a.* from trademarkquery a,tmqapp b where tmq18=tqa01(+) and tmq11 is null and tmq10 is null " & strQ1 & " order by tmq01"
        End If
        intQ = 0
        Set RsQ = ClsLawReadRstMsg(intQ, strQ2)
        If intQ = 1 Then
          If MsgBox("預計有 " & RsQ.RecordCount & " 筆查名單要分派，是否繼續？", vbYesNo + vbInformation + vbDefaultButton1) = vbNo Then
             Exit Sub
          End If
          
          'Added by Lydia 2017/11/07 遇到有人員臨時請假,查名單已到期的狀況
          RsQ.MoveFirst
          If "" & RsQ.Fields("tmq04") <> "" And "" & RsQ.Fields("tmq04") <> strSrvDate(1) Then
             If MsgBox("是否重新計算期限日期？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                bolUpTmq06 = True
                Inputtm = Left(Format(ServerTime, "000000"), 4)
             End If
          End If
          'end 2017/11/07
          
          cnnConnection.BeginTrans
            '更新統計人員狀態
            strQ2 = "select tmqm02,nvl(tmqm03,'N') tmqm03,count(*) r1,count(tmqsr17) r2 from tmqmember,tmqsumr " & _
                    "where tmqm01<>tmqm02 and tmqm01=tmqsr01(+)  group by tmqm02,nvl(tmqm03,'N') "
            intQ = 1
            Set RsTemp = ClsLawReadRstMsg(intQ, strQ2)
            If intQ = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  If ("" & RsTemp.Fields("tmqm03") = "Y" And Val("" & RsTemp.Fields("r2")) > 0) Or _
                     ("" & RsTemp.Fields("tmqm03") = "N" And Val("" & RsTemp.Fields("r1")) = Val("" & RsTemp.Fields("r2"))) Then
                     strQ2 = "update tmqsumr set tmqsr17='N' where tmqsr01=" & CNULL(Trim("" & RsTemp.Fields("tmqm02")))
                     cnnConnection.Execute strQ2
                  End If
                  RsTemp.MoveNext
               Loop
            End If
            '有可能是整批變更資料後的重新分發，先更新統計單量
            Call PUB_TMQtake("2", "", , , , , 0, False)
            Call PUB_TMQtake("2", "", , , , , 1, False)
            
            'RsQ.MoveFirst 'Remove by Lydia 2017/11/07
            Do While Not RsQ.EOF
                'Added by Lydia 2017/11/07 視情況重新計算期限
                strQ1 = ""
                If bolUpTmq06 = True Then
                   'Modified by Lydia 2018/05/25 +狀態
                   'InputWDay = PUB_GetNewTmq06(IIf(Val("" & RsQ.Fields("tmq09")) > 0, 2, 1), "" & RsQ.Fields("tmq03"), strSrvDate(1), Inputtm)
                   InputWDay = PUB_GetNewTmq06(IIf(Val("" & RsQ.Fields("tmq09")) > 0, 2, 1), "" & RsQ.Fields("tmq03"), strSrvDate(1), Inputtm)
                   strQ1 = ", tmq06=" & InputWDay & " "
                End If
                'end 2017/11/07
                
                'end 2017/11/06
                '不計算本數，不重新計算期限
                strQuser = PUB_GetTMQUserPos(True, "1", Val("" & RsQ.Fields("tmq07")), Val("" & RsQ.Fields("tmq08")), Val("" & RsQ.Fields("tmq09")), "" & RsQ.Fields("tmq03"))
                If strQuser <> "" Then
                   'Modified by Lydia 2017/11/07 + strQ1
                   strQ2 = "update trademarkquery set tmq10=" & CNULL(strQuser) & ",tmq05=" & strSrvDate(1) & strQ1 & " where tmq01=" & CNULL(RsQ.Fields("tmq01"))
                   cnnConnection.Execute strQ2
                   '委查日期含前2個工作天到當天,重新計算
                   '分發日期改成系統日=當天
                   Call PUB_TMQtake("2", strQuser, , , , , 1, False)
                   iCnt = iCnt + 1
                End If
                RsQ.MoveNext
            Loop
          cnnConnection.CommitTrans
          MsgBox "已分派完 " & iCnt & " 筆查名單!", vbInformation
        End If
   Else
   'end 2017/06/23
        strQ2 = "select tmqsr01,tmqsr17 from tmqsumr where tmqsr01=" & CNULL(Trim(Left("" & Combo1.Text, 6)))
        intQ = 0
        Set RsQ = ClsLawReadRstMsg(intQ, strQ2)
        If intQ = 1 Then
           If "" & RsQ(1) = "" Then
              If MsgBox("查名人狀態為可分派，若要繼續作業會自動將狀態改為不可分派", vbYesNo) = vbYes Then
                 bolUpd = True
              Else
                 Exit Sub
              End If
           End If
           '更新未查覆完畢的查名單
           'Modified by Lydia 2017/06/23
           'strQ2 = "select * from trademarkquery where tmq11 is null" & strQ1 & " order by tmq01"
           strQ2 = "select a.* from trademarkquery a,tmqapp b where tmq18=tqa01(+) and tmq11 is null" & strQ1 & " order by tmq01"
           intQ = 1
           Set RsQ = ClsLawReadRstMsg(intQ, strQ2)
           If intQ = 0 Then
              MsgBox "無委查單可分派!"
              Exit Sub
           Else
              'Added by Lydia 2017/06/23
              If MsgBox("預計有 " & RsQ.RecordCount & " 筆查名單要分派，是否繼續？", vbYesNo + vbInformation + vbDefaultButton1) = vbNo Then
                  Exit Sub
              End If
              'end 2017/06/23
              
              RsQ.MoveFirst
              'Added by Lydia 2017/11/07 遇到有人員臨時請假,查名單已到期的狀況
              If "" & RsQ.Fields("tmq04") <> "" And "" & RsQ.Fields("tmq04") <> strSrvDate(1) Then
                 If MsgBox("是否重新計算期限日期？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                    bolUpTmq06 = True
                 End If
              End If
              'end 2017/11/07
              
              cnnConnection.BeginTrans 'Added by Lydia 2017/06/23
              '更新查名人狀態
              If bolUpd Then
                strQ2 = "update tmqsumr set tmqsr17='N' where tmqsr01=" & CNULL(Trim(Left("" & Combo1.Text, 6)))
                cnnConnection.Execute strQ2
                '組群分派，一人請假全部不分
                strQ2 = "update tmqsumr set tmqsr17='N' where tmqsr01 in (select tmqm02 from tmqmember where tmqm01=" & CNULL(Trim(Left("" & Combo1.Text, 6))) & " and tmqm02 <> " & CNULL(Trim(Left("" & Combo1.Text, 6))) & " and tmqm03='Y') "
                cnnConnection.Execute strQ2
                strQ2 = "update tmqsumr set tmqsr17='N' where tmqsr01 in (select tmqm01 from tmqmember where tmqm02 in (select tmqm02 from tmqmember where tmqm01=" & CNULL(Trim(Left("" & Combo1.Text, 6))) & " and tmqm02 <> " & CNULL(Trim(Left("" & Combo1.Text, 6))) & " and tmqm03='Y')) "
                cnnConnection.Execute strQ2
              End If
              'Added by Lydia 2020/03/17 先將分發日期拿掉,更新統計單量; ex.10張委查單(文字1,圖形)分成兩次重新分發,其中內商程序-79041被分到8張
              strQ2 = "update trademarkquery set tmq05=null,tmq10=null where tmq01 in (select a.tmq01 from trademarkquery a,tmqapp b where tmq18=tqa01(+) and tmq11 is null " & strQ1 & " ) "
              cnnConnection.Execute strQ2, intI
              '整批變更資料後的重新分發 , 先更新統計單量
              Call PUB_TMQtake("2", "", , , , , 0, False)
              Call PUB_TMQtake("2", "", , , , , 1, False)
              'end 2020/03/17

              With RsQ
                Do While Not RsQ.EOF
                    'Added by Lydia 2017/11/07 視情況重新計算期限
                    strQ1 = ""
                    If bolUpTmq06 = True Then
                       InputWDay = PUB_GetNewTmq06(IIf(Val("" & RsQ.Fields("tmq09")) > 0, 2, 1), "" & RsQ.Fields("tmq03"), strSrvDate(1), Inputtm)
                       strQ1 = ",tmq05=" & strSrvDate(1) & ", tmq06=" & InputWDay & " "
                    End If
                    'end 2017/11/07
                    '不計算本數，不重新計算期限
                    strQuser = PUB_GetTMQUserPos(True, "1", Val("" & .Fields("tmq07")), Val("" & .Fields("tmq08")), Val("" & .Fields("tmq09")), "" & .Fields("tmq03"))
                    If strQuser <> "" Then
                       'Modified by Lydia 2017/11/07 +strQ1
                       strQ2 = "update trademarkquery set tmq10=" & CNULL(strQuser) & strQ1 & " where tmq01=" & CNULL(.Fields("tmq01"))
                       cnnConnection.Execute strQ2
                       '委查日期含前2個工作天到當天,重新計算
                       'Modified by Lydia 2017/01/25 tmq04->tmq05 (收件分發日期)
                       'Modified by Lydia 2021/07/06 debug 分發日期改成系統日=當天
                       'If .Fields("tmq05") = strSrvDate(1) Then
                       '   '當日
                       '   'Modified by Lydia 2017/06/23 不包Transaction
                       '   'Call PUB_TMQtake("2", strQuser, , , , , 1)
                       '   'Call PUB_TMQtake("2", .Fields("tmq10"), , , , , 1)
                       '   Call PUB_TMQtake("2", strQuser, , , , , 1, False)
                       '   Call PUB_TMQtake("2", .Fields("tmq10"), , , , , 1, False)
                       '   'end 2017/06/23
                       'Else
                       '   '前2日
                       '   'Modified by Lydia 2017/06/23 不包Transaction
                       '   'Call PUB_TMQtake("2", strQuser, , , , , 0)
                       '   'Call PUB_TMQtake("2", .Fields("tmq10"), , , , , 0)
                       '   Call PUB_TMQtake("2", strQuser, , , , , 0, False)
                       '   Call PUB_TMQtake("2", .Fields("tmq10"), , , , , 0, False)
                       '   'end 2017/06/23
                       'End If
                       Call PUB_TMQtake("2", strQuser, , , , , 1, False)
                       Call PUB_TMQtake("2", .Fields("tmq10"), , , , , 1, False)
                       'end 2021/07/06
                       iCnt = iCnt + 1 'Added by Lydia 2017/06/23
                    End If
                  .MoveNext
                Loop
              End With
              cnnConnection.CommitTrans 'Added by Lydia 2017/06/23
              'Modified by Lydia 2017/06/23
              'MsgBox "重新分派完成，請重新查詢!"
              MsgBox "已分派完 " & iCnt & " 筆查名單!", vbInformation
           End If
        End If
   End If 'end 2017/06/23
   Exit Sub
ErrInfo:

If Err.Number <> 0 Then
   MsgBox Err.Description
   If strQuser <> "" Then cnnConnection.RollbackTrans
End If

End Sub

'Memo by Lydia 2016/10/12 過去紙本作業查名人可參考其他人的查名路徑,業務可參考其他人的查名單;但是為了避免查名人或業務直接參照其他查名單,所以不開放看其他查名單.
'                         若真的有需要開放,請先經過林經理和文雄協商
Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strCon As String, strCon2 As String
'Dim strTQC As String 'Added by Lydia 2016/07/06 改收文進度檔的對照方式 'Mark by Lydia 2025/04/28
Dim strField As String, strQuery As String 'Added by Lydia 2025/04/28

   'Added by Lydia 2016/06/01 +本所案號
   If txtField(9) <> "" And Len(Trim(txtField(10))) <> 6 Then
      MsgBox "請輸入完整的本所案號!", vbOKOnly, "輸入錯誤"
      txtField(10).SetFocus
      QueryData = False
      Exit Function
   End If
   'end 2016/06/01
   
JumpReSearch:
   m_blnColOrderAsc = True
   QueryData = True
   GRD1.Clear
   strCon = "": strCon2 = ""
   'strTQC = "" 'Added by Lydia 2016/07/06 'Mark by Lydia 2025/04/28
   '委查人
   If InStr(Me.Caption, "查覆區") > 0 Then
      If m_TMQApp = "" Then
         '查覆區預設不列出已收文的委查單
         'Modified by Lydia 2016/04/06 查名代號(收文組群)改成另開Table對應
         'Modified by Lydia 2016/04/28 將已發文(送件)案件特別區隔
         'strCon = strCon & " and tmq21 is null"
         'Modified by Lydia 2016/06/17
         'strCon = strCon & " and cp27 is null"
         'Modified by Lydia 2022/01/06 排除從歷程過來看結果+bolCaseRead
         If Trim(txtField(8)) = "" And bolCaseRead = False Then  'Added by Lydia 2017/03/21 只要輸入委查單號,就不限制狀態
            strCon = strCon & " and c1.cp27 is null"
         End If
      End If
      'Modified by Lydia 2022/01/06 +從歷程過來看結果
      'If InStr(Combo2.Text, "全部") > 0 And Pub_StrUserSt03 = "M51" Then
      If InStr(Combo2.Text, "全部") > 0 And (Pub_StrUserSt03 = "M51" Or bolCaseRead = True) Then
      'Added by Lydia 2019/08/12 創新業務組成員可操作清單
      'Modified by Lydia 2020/12/04 debug
      'ElseIf InStr(stIdList, "W") > 0 And Left(Pub_StrUserSt15, 1) = "W" Then
      ElseIf InStr(stIdList, "W") > 0 And Left(strGrpTmp1, 1) = "W" Then
           strCon = strCon & " and tmq02 in (" & stIdList & ")"
      'end 2019/08/12
      Else
         strCon = strCon & " and tmq02='" & Trim(Left("" & Combo2.Text, 6)) & "'"
      End If
   ElseIf txtField(0) <> "" Then
      strCon = strCon & " and tmq02='" & Trim(txtField(0)) & "'"
   End If
   strCon2 = strCon
   
   '委查期間
   If txtField(1) <> "" Then strCon = strCon & " and tmq04>=" & TransDate(Trim(txtField(1)), 2)
   If txtField(2) <> "" Then strCon = strCon & " and tmq04<=" & TransDate(Trim(txtField(2)), 2)
   '期限期間
   If txtField(3) <> "" Then strCon = strCon & " and tmq06>=" & TransDate(Trim(txtField(3)), 2)
   If txtField(4) <> "" Then strCon = strCon & " and tmq06<=" & TransDate(Trim(txtField(4)), 2)
   '類別
   If txtField(5) <> "" And txtField(5) <> "0" Then strCon = strCon & " and tqa06='" & Trim(txtField(5)) & "'"
   strExc(0) = strCon 'Added by Lydia 2016/05/31
   
   '狀態
   'Added by Lydia 2016/03/21加顯示狀態
   lblState = ""
   'Modified by Lydia 2017/03/21 只要輸入委查單號,就不限制狀態
   'If txtField(6) <> "" Then
   If txtField(6) <> "" And Trim(txtField(8)) = "" Then
      Select Case txtField(6)
         Case "0" '全部=未發文
            lblState = "處理中 + 已完成"
         Case "1" '處理中
            lblState = "處理中"
            strCon = strCon & " and tmq11 is null"
         Case "2" '已完成
            lblState = "已完成"
            strCon = strCon & " and tmq11>0 and tqa20 is null"
         Case "3"
            'Modified by Lydia 2016/04/28 將已發文(送件)案件特別區隔
            'lblState = "已收文"
            'If strCon <> "" Then
            '   strCon = Replace(strCon, "tmq21 is null", "not tmq21 is null ")
            'Else
            '   strCon = strCon & " and not tmq21 is null "
            'End If
            lblState = "已發文"
            'Modified by Lydia 2016/05/31
            'If strCon <> "" Then
            If strCon <> "" And InStr(strCon, "c1.cp27 is null") > 0 Then
               strCon = Replace(strCon, "c1.cp27 is null", "c1.cp27 is not null ")
            Else
               strCon = strCon & " and c1.cp27 is not null "
            End If
      End Select
      'Modified by Lydia 2016/04/28 將已發文(送件)案件特別區隔
      'Modified by Lydia 2022/01/06 排除從歷程過來看結果bolCaseRead
      If txtField(6) <> "3" And InStr(strCon, "c1.cp27") = 0 And bolCaseRead = False Then strCon = strCon & " and c1.cp27 is null"
   End If
   '客戶名稱
   If textCName <> "" Then
      strExc(1) = Replace(Trim(textCName), " ", "%")
      strExc(1) = Replace(strExc(1), "%%", "%")
      strCon = strCon & " and upper(tqa04) like '%" & UCase(strExc(1)) & "%'"
      strExc(0) = strExc(0) & " and upper(tqa04) like '%" & UCase(strExc(1)) & "%'" 'Added by Lydia 2016/05/31
   '從接洽單傳申請人(中文名稱) ->客戶名稱
   ElseIf contCusName <> "" And bolCont = True Then
          '從字首4字開始比對到字首2字
          If Len(contCusName) > 4 Then
             contCusName = Mid(contCusName, 1, 4)
          End If
          strCon = strCon & " and tqa04 like '" & contCusName & "%' "
   End If
   '查名人
   If InStr(Combo1.Text, "全部") = 0 Then
      strCon = strCon & " and tmq10='" & Trim(Left(Combo1.Text, 5)) & "'"
      strCon2 = strCon2 & " and tmq10='" & Trim(Left(Combo1.Text, 5)) & "'"
      strExc(0) = strExc(0) & " and tmq10='" & Trim(Left(Combo1.Text, 5)) & "'" 'Added by Lydia 2016/05/31
   ElseIf R_type = "U" And strUserNum <> Trim(Left(Combo1.Text, 5)) Then
           If Pub_StrUserSt03 <> "M51" Then iStiu = 0 '非分派到的查名人，不可修改。電腦中心除外。
   End If
   '委查單號
   If txtField(8) <> "" Then
      txtField(8).Text = Replace(txtField(8).Text, ".", ",")
      strCon = strCon & " and tmq01 in (" & CNULL(Replace(txtField(8).Text, ",", "','")) & ")"
      strExc(0) = strExc(0) & " and tmq01 in (" & CNULL(Replace(txtField(8).Text, ",", "','")) & ")" 'Added by Lydia 2016/05/31
   End If
   'Added by Lydia 2016/06/01 +本所案號
   If txtField(9) <> "" And txtField(10) <> "" Then
      strExc(1) = " and cp01='" & Trim(txtField(9)) & "' and cp02='" & Trim(txtField(10)) & "' and cp03='" & Left(Trim(txtField(11)) & "0", 1) & "' and cp04='" & Left(Trim(txtField(12)) & "00", 2) & "'"
      If txtField(9) = "T" Then
         'Modified by Lydia 2021/11/19 增加737智財協作之T案
         'strExc(1) = strExc(1) & " and cp10='" & TMQ_T案 & "'"
         strExc(1) = strExc(1) & " and instr('" & TMQ_T案 & "', cp10) > 0 "
      ElseIf txtField(9) = "TS" Then
         'Modified by Lydia 2021/11/19
         'strExc(1) = strExc(1) & " and cp10='" & TMQ_TS案 & "'"
         strExc(1) = strExc(1) & " and instr('" & TMQ_TS案 & "', cp10) > 0 "
      End If
      strExc(3) = Replace(strExc(1), "cp", "c1.cp") & " and c1.cp57 is null" 'Added by Lydia 2016/06/02 個案查詢的本所案號改抓輸入的案號
      'Modified by Lydia 2016/07/06 改收文進度檔的對照方式
      'strExc(1) = " and tmq01 in (select tqc03 from tmqcasemap where tqc02 in (select cp09 from caseprogress where cp57 is null" & strExc(1) & ")) "
      'strCon = strCon & strExc(1)
      'strExc(0) = strExc(0) & strExc(1)
      'Modified by Lydia 2025/04/28 整理SQL
      'strTQC = " AND TMQ01=TQC03(+) AND TQC02=C1.CP09(+) "
      'strCon = strCon & strTQC & strExc(3)
      strCon = strCon & " AND TMQ01=TQC03(+) AND TQC02=C1.CP09(+) " & strExc(3)
      'end 2025/04/28
   Else
      'Modified by Lydia 2025/04/28 整理SQL
      'strTQC = " AND TMQ01=TQC03(+) AND TMQ21=TQC02(+) AND TMQ21=c1.CP09(+)"
      'strCon = strCon & strTQC
      strCon = strCon & " AND TMQ01=TQC03(+) AND TMQ21=TQC02(+) AND TMQ21=c1.CP09(+)"
   End If
   'end 2016/06/01
   
   'strCon = strCon & " And nvl(tmq20,'N') = 'N' " 'Added by Lydia 2018/03/20 用TMQ20判斷是否已刪除明細 'Remove by Lydia 2018/03/21 影響速度
   
   If R_type = "M" Then
        'Added by Lydia 2016/07/06 +判斷未輸入本所案號或委查單,限近似本所案
       If txtField(8) & txtField(9) & txtField(10) = "" Then
           'Modified by Lydia 2017/12/13 預設不顯示撤回
           'strCon = strCon & " and v1c2 in ('" & TMQ_近似1 & "','" & TMQ_近似2 & "') "
           strCon = strCon & " and v1c2 in ('" & TMQ_近似1 & "','" & TMQ_近似2 & "') and nvl(tqa20,'N') ='N' "
       End If
  'Added by Lydia 2016/06/01
       'Modified by Lydia 2016/07/06 覆核區若覆核結果非近似本所案,則不顯示
       'strExc(2) = " min(tqd06) " '覆核區的結果只抓查名結果
       strExc(2) = " min(nvl(tqd09,tqd06)) " '=V1C2
   Else
       strExc(2) = " min(nvl(tqd09,tqd06)) " '覆核結果取代查名結果
   'end 2016/06/01
   End If

   Screen.MousePointer = vbHourglass
   'Modified by Lydia 2016/04/28 +通知送件和已發文判斷，案號取代申請號顯示
'    strSql = "select ' ' V,TMQ19,TQA20,TMQ18,TMQ02,S1.ST02 SST1,DECODE(TQA06,'1',DECODE(TQA13,NULL,'(非正常字)',TQA13),'2','(圖形查詢)') 文字1," & _
'             "DECODE(TQA06,'1',DECODE(TQA14,NULL,DECODE(TQA08,NULL,DECODE(TQF05,NULL,'','(非正常字)'),'(非正常字)'),TQA14),'2','') 文字2," & _
'             "TQA04,decode(v1c2," & TMQ_結果查詢 & ") 結果,TMQ03,TMQ10,S2.ST02 SST02,TMQ01,(TMQ04-19110000) TMQ04,(TMQ06-19110000) TMQ06, (TMQ11-19110000) TMQ11,(TMQ23-19110000) 覆核日期,TMQ21 as 總收文號" & _
'             " FROM TMQAPP,trademarkquery,STAFF S1,STAFF S2," & _
'             "(select tqd02 v1c1, min(tqd06) v1c2 from tmqdetail group by tqd02) VT1 " & _
'             ",(SELECT TQF01,TQF03,TQF05 FROM TMQFILE WHERE TQF02||TQF03||TQF04='" & TMQ_附件F02 & TMQ_AkindWord2 & TMQ_附件F04 & "') VT2 " & _
'             "WHERE TQA01=TMQ18(+) and tmq01=v1c1(+) AND TMQ02=S1.ST01(+) AND TMQ10=S2.ST01(+) AND NOT(TMQ03 IS NULL) AND TQA01=TQF01(+) " & strCon
'     'Added by Lydia 2016/03/28 +控制是否已讀取(True:預設顯示未讀, false:不預設顯示未讀資料)
'     If TMQ_CtrRead Then
'        If R_type = "Q" And bolCont = False Then '預設:委查人未讀的查名單會一直出現,除了接洽單收文的情況
'            '附件未讀=>排除已撤回
'            strSql = strSql & " Union select ' ' V,TMQ19,TQA20,TMQ18,TMQ02,S1.ST02 SST1,DECODE(TQA06,'1',DECODE(TQA13,NULL,'(非正常字)',TQA13),'2','(圖形查詢)') 文字1," & _
'                     "DECODE(TQA06,'1',DECODE(TQA14,NULL,DECODE(TQA08,NULL,DECODE(TQF05,NULL,'','(非正常字)'),'(非正常字)'),TQA14),'2','') 文字2," & _
'                     "TQA04,decode(v1c2," & TMQ_結果查詢 & ") 結果,TMQ03,TMQ10,S2.ST02 SST02,TMQ01,(TMQ04-19110000) TMQ04,(TMQ06-19110000) TMQ06, (TMQ11-19110000) TMQ11,(TMQ23-19110000) 覆核日期,TMQ21 as 總收文號" & _
'                     " FROM TMQAPP,trademarkquery,STAFF S1,STAFF S2," & _
'                     "(select tqd02 v1c1, min(tqd06) v1c2 from tmqdetail group by tqd02) VT1 " & _
'                     ",(SELECT TQF01,TQF03,TQF05 FROM TMQFILE WHERE TQF02||TQF03||TQF04='" & TMQ_附件F02 & TMQ_AkindWord2 & TMQ_附件F04 & "') VT2 " & _
'                     "WHERE TQA01=TMQ18(+) and tmq01=v1c1(+) AND TMQ02=S1.ST01(+) AND TMQ10=S2.ST01(+) AND NOT(TMQ03 IS NULL) AND TQA01=TQF01(+) AND TMQ11 > 0 AND TMQ19||TQA20 IS NULL " & strCon2 & _
'                    IIf(txtField(0) <> "", " and tmq02='" & Trim(txtField(0)) & "'", "")
'        End If
'     End If
    'Modified by Lydia 2016/06/01 覆核結果取代查名結果
    'Modified by Lydia 2016/06/02 TMQ_結果查詢改成模組
    'Modified by Lydia 2016/06/02 個案查詢的本所案號改抓輸入的案號,+ShowCP09
    'strSql = "select ' ' V,TMQ19,TQA20,DECODE(c1.CP27,NULL,DECODE(TMQ20,NULL,'','◎'),'●') C01," & _
             "DECODE(c1.CP01,NULL,'',DECODE(c1.CP03||c1.CP04,'000',c1.CP01||'-'||c1.CP02,c1.CP01||'-'||c1.CP02||'-'||c1.CP03||'-'||c1.CP04)) CASENO," & _
             "TMQ18,TMQ02,S1.ST02 SST1,DECODE(TQA06,'1',DECODE(TQA13,NULL,'(非正常字)',TQA13),'2','(圖形查詢)') 文字1," & _
             "DECODE(TQA06,'1',DECODE(TQA14,NULL,DECODE(TQA08,NULL,DECODE(TQF05,NULL,'','(非正常字)'),'(非正常字)'),TQA14),'2','') 文字2," & _
             "TQA04,decode(v1c2," & PUB_GetTMQans("3", True) & ") 結果,TMQ03,TMQ10,S2.ST02 SST02,TMQ01,(TMQ04-19110000) TMQ04,(TMQ06-19110000) TMQ06, (TMQ11-19110000) TMQ11,(TMQ23-19110000) 覆核日期,TMQ21 as 總收文號,(TMQ20-19110000) as TMQ20,c1.CP27,c1.CP57,c1.CP14" & _
             " FROM TMQAPP,trademarkquery,STAFF S1,STAFF S2,caseprogress c1," & _
             "(select tqd02 v1c1, " & strExc(2) & " v1c2 from tmqdetail group by tqd02) VT1 " & _
             ",(SELECT TQF01,TQF03,TQF05 FROM TMQFILE WHERE TQF02||TQF03||TQF04='" & TMQ_附件F02 & TMQ_AkindWord2 & TMQ_附件F04 & "') VT2 " & _
             "WHERE TQA01=TMQ18(+) and tmq01=v1c1(+) AND TMQ02=S1.ST01(+) AND TMQ10=S2.ST01(+) AND NOT(TMQ03 IS NULL) AND TQA01=TQF01(+) AND TMQ21=c1.CP09(+) " & strCon
    'Modified by Lydia 2016/07/06 +TMQCASEMAP, TMQ20改TQC07
    'Modified by Lydia 2017/09/28  電子化前的查名單TMQ18改為0 , 加指定索引/*+ INDEX(TRADEMARKQUERY IDXTMQ18) */ 加速查詢
    'Added by Lydia 2017/11/06 查覆區欄位順序:委查單號,期限日期,委查日期 (因為智權人員常會打電話催查名人員BY嘉雯)
    'Modified by Lydia 2025/04/28 整理SQL，改用strField設定欄位順序，可以不用分別寫
    'If R_type = "Q" Then
    '    strSql = "select /*+ INDEX(TRADEMARKQUERY IDXTMQ18) */ ' ' V,TMQ19,TQA20,DECODE(c1.CP27,NULL,DECODE(TQC07,NULL,'','◎'),'●') C01," & _
    '             "DECODE(c1.CP01,NULL,'',DECODE(c1.CP03||c1.CP04,'000',c1.CP01||'-'||c1.CP02,c1.CP01||'-'||c1.CP02||'-'||c1.CP03||'-'||c1.CP04)) CASENO," & _
    '             "TMQ18,TMQ02,S1.ST02 SST1,DECODE(TQA06,'1',DECODE(TQA13,NULL,'(非正常字)',TQA13),'2','(圖形查詢)') 文字1," & _
    '             "DECODE(TQA06,'1',DECODE(TQA14,NULL,DECODE(TQA08,NULL,DECODE(TQF05,NULL,'','(非正常字)'),'(非正常字)'),TQA14),'2','') 文字2," & _
    '             "TQA04,decode(v1c2," & PUB_GetTMQans("3", True) & ") 結果,TMQ03,TMQ10,S2.ST02 SST02,TMQ01,(TMQ06-19110000) TMQ06,(TMQ04-19110000) TMQ04, (TMQ11-19110000) TMQ11,(TMQ23-19110000) 覆核日期,TMQ21 as 總收文號,(TQC07-19110000) as TMQ20,c1.CP27,c1.CP57,c1.CP14,c1.CP09 as SCP09" & _
    '             " FROM TMQAPP,trademarkquery,STAFF S1,STAFF S2,caseprogress c1,TMQCASEMAP, " & _
    '             "(select tqd02 v1c1, " & strExc(2) & " v1c2 from tmqdetail group by tqd02) VT1 " & _
    '             ",(SELECT TQF01,TQF03,TQF05 FROM TMQFILE WHERE TQF02||TQF03||TQF04='" & TMQ_附件F02 & TMQ_AkindWord2 & TMQ_附件F04 & "') VT2 " & _
    '             "WHERE TQA01=TMQ18(+) and tmq01=v1c1(+) AND TMQ02=S1.ST01(+) AND TMQ10=S2.ST01(+) AND NOT(TMQ03 IS NULL) AND TQA01=TQF01(+) " & strCon
    'Else '其他人員欄位順序:委查單號,委查日期,期限日期
    ''end 2017/11/06
    'end 2025/04/28

        'Modified by Lydia 2025/04/28 改用strField設定欄位順序,
        'Modified by Lydia 2025/04/28 拿掉/*+ INDEX(TRADEMARKQUERY IDXTMQ18) */，反而變快
        'Modified by Lydia 2025/04/28 先存共通語法
        'strSql = "select /*+ INDEX(TRADEMARKQUERY IDXTMQ18) */ ' ' V,TMQ19,TQA20,DECODE(c1.CP27,NULL,DECODE(TQC07,NULL,'','◎'),'●') C01," & _
        '         "DECODE(c1.CP01,NULL,'',DECODE(c1.CP03||c1.CP04,'000',c1.CP01||'-'||c1.CP02,c1.CP01||'-'||c1.CP02||'-'||c1.CP03||'-'||c1.CP04)) CASENO," & _
                 "TMQ18,TMQ02,S1.ST02 SST1,DECODE(TQA06,'1',DECODE(TQA13,NULL,'(非正常字)',TQA13),'2','(圖形查詢)') 文字1," & _
                 "DECODE(TQA06,'1',DECODE(TQA14,NULL,DECODE(TQA08,NULL,DECODE(TQF05,NULL,'','(非正常字)'),'(非正常字)'),TQA14),'2','') 文字2," & _
                 "TQA04,decode(v1c2," & PUB_GetTMQans("3", True) & ") 結果,TMQ03,TMQ10,S2.ST02 SST02,TMQ01,(TMQ04-19110000) TMQ04,(TMQ06-19110000) TMQ06, (TMQ11-19110000) TMQ11,(TMQ23-19110000) 覆核日期,TMQ21 as 總收文號,(TQC07-19110000) as TMQ20,c1.CP27,c1.CP57,c1.CP14,c1.CP09 as SCP09" & _
                 " FROM TMQAPP,trademarkquery,STAFF S1,STAFF S2,caseprogress c1,TMQCASEMAP, " & _
                 "(select tqd02 v1c1, " & strExc(2) & " v1c2 from tmqdetail group by tqd02) VT1 " & _
                 ",(SELECT TQF01,TQF03,TQF05 FROM TMQFILE WHERE TQF02||TQF03||TQF04='" & TMQ_附件F02 & TMQ_AkindWord2 & TMQ_附件F04 & "') VT2 " & _
                 "WHERE TQA01=TMQ18(+) and tmq01=v1c1(+) AND TMQ02=S1.ST01(+) AND TMQ10=S2.ST01(+) AND NOT(TMQ03 IS NULL) AND TQA01=TQF01(+) " & strCon
        If R_type = "Q" Then
           strField = ",(TMQ06-19110000) TMQ06,(TMQ04-19110000) TMQ04, (TMQ11-19110000) TMQ11"
        Else
           strField = ",(TMQ04-19110000) TMQ04,(TMQ06-19110000) TMQ06, (TMQ11-19110000) TMQ11"
        End If
        strQuery = "select ' ' V,TMQ19,TQA20,DECODE(c1.CP27,NULL,DECODE(TQC07,NULL,'','◎'),'●') C01," & _
                 "DECODE(c1.CP01,NULL,'',DECODE(c1.CP03||c1.CP04,'000',c1.CP01||'-'||c1.CP02,c1.CP01||'-'||c1.CP02||'-'||c1.CP03||'-'||c1.CP04)) CASENO," & _
                 "TMQ18,TMQ02,S1.ST02 SST1,DECODE(TQA06,'1',DECODE(TQA13,NULL,'(非正常字)',TQA13),'2','(圖形查詢)') 文字1," & _
                 "DECODE(TQA06,'1',DECODE(TQA14,NULL,DECODE(TQA08,NULL,DECODE(TQF05,NULL,'','(非正常字)'),'(非正常字)'),TQA14),'2','') 文字2," & _
                 "TQA04,decode(v1c2," & PUB_GetTMQans("3", True) & ") 結果,TMQ03,TMQ10,S2.ST02 SST02,TMQ01" & strField & ",(TMQ23-19110000) 覆核日期,TMQ21 as 總收文號,(TQC07-19110000) as TMQ20,c1.CP27,c1.CP57,c1.CP14,c1.CP09 as SCP09" & _
                 " FROM TMQAPP,trademarkquery,STAFF S1,STAFF S2,caseprogress c1,TMQCASEMAP, " & _
                 "(select tqd02 v1c1, " & strExc(2) & " v1c2 from tmqdetail group by tqd02) VT1 " & _
                 ",(SELECT TQF01,TQF03,TQF05 FROM TMQFILE WHERE TQF02||TQF03||TQF04='" & TMQ_附件F02 & TMQ_AkindWord2 & TMQ_附件F04 & "') VT2 " & _
                 "WHERE TQA01=TMQ18(+) and tmq01=v1c1(+) AND TMQ02=S1.ST01(+) AND TMQ10=S2.ST01(+) AND NOT(TMQ03 IS NULL) AND TQA01=TQF01(+) "
        strSql = strQuery & strCon
     'End If 'end 2017/11/06 'end 2025/04/28
     
     'Modified by Lydia 2025/04/28 debug: 待查+查詢非已發文+輸入本所案號>> 因為畫面條件不同，所以另外列出分給查名人的查名單
     'If InStr(Me.Caption, "待查區") > 0 And txtField(6) <> "3" And Trim(txtField(9) & txtField(10)) = "" Then
     '   'Modified by Lydia 2016/07/06 +TMQCASEMAP,+strTQC, TMQ20改TQC07
     '   'Modified by Lydia 2017/09/28  電子化前的查名單TMQ18改為0 , 加指定索引/*+ INDEX(TRADEMARKQUERY IDXTMQ18) */ 加速查詢
     '   strSql = strSql & " Union select /*+ INDEX(TRADEMARKQUERY IDXTMQ18) */ ' ' V,TMQ19,TQA20,DECODE(c1.CP27,NULL,DECODE(TQC07,NULL,'','◎'),'●') C01," & _
     '        "DECODE(c1.CP01,NULL,'',DECODE(c1.CP03||c1.CP04,'000',c1.CP01||'-'||c1.CP02,c1.CP01||'-'||c1.CP02||'-'||c1.CP03||'-'||c1.CP04)) CASENO," & _
     '        "TMQ18,TMQ02,S1.ST02 SST1,DECODE(TQA06,'1',DECODE(TQA13,NULL,'(非正常字)',TQA13),'2','(圖形查詢)') 文字1," & _
     '        "DECODE(TQA06,'1',DECODE(TQA14,NULL,DECODE(TQA08,NULL,DECODE(TQF05,NULL,'','(非正常字)'),'(非正常字)'),TQA14),'2','') 文字2," & _
     '        "TQA04,decode(v1c2," & PUB_GetTMQans("3", True) & ") 結果,TMQ03,TMQ10,S2.ST02 SST02,TMQ01,(TMQ04-19110000) TMQ04,(TMQ06-19110000) TMQ06, (TMQ11-19110000) TMQ11,(TMQ23-19110000) 覆核日期,TMQ21 as 總收文號,(TQC07-19110000) as TMQ20,c1.CP27,c1.CP57,c1.CP14,c1.CP09 as SCP09" & _
     '        " FROM TMQAPP,trademarkquery,STAFF S1,STAFF S2,caseprogress c1,TMQCASEMAP, " & _
     '        "(select tqd02 v1c1, " & strExc(2) & " v1c2 from tmqdetail group by tqd02) VT1 " & _
     '        ",(SELECT TQF01,TQF03,TQF05 FROM TMQFILE WHERE TQF02||TQF03||TQF04='" & TMQ_附件F02 & TMQ_AkindWord2 & TMQ_附件F04 & "') VT2 " & _
     '        "WHERE TQA01=TMQ18(+) and tmq01=v1c1(+) AND TMQ02=S1.ST01(+) AND TMQ10=S2.ST01(+) AND NOT(TMQ03 IS NULL) AND TQA01=TQF01(+) " & strExc(0) & _
     '        strTQC & " and tmq11 is null and c1.CP27 is not null "
     'End If
     If InStr(Me.Caption, "待查區") > 0 And txtField(6) <> "3" And Trim(txtField(9) & txtField(10)) <> "" Then
        strSql = strSql & " Union " & strQuery & " AND TMQ01=TQC03(+) AND TMQ21=TQC02(+) AND TMQ21=c1.CP09(+) and tmq11 is null " & strCon2
     End If
     'end 2025/04/28
     
     'Mark by Lydia 2025/04/28 整理SQL，判斷不用
     ''控制是否已讀取(True:預設顯示未讀, false:不預設顯示未讀資料)
     'If TMQ_CtrRead Then
     '   If R_type = "Q" And bolCont = False Then '預設:委查人未讀的查名單會一直出現,除了接洽單收文的情況
     '       '附件未讀=>排除已撤回
     '       'Modified by Lydia 2016/07/06 +TMQCASEMAP, TMQ20改TQC07
     '       'Modified by Lydia 2017/09/28  電子化前的查名單TMQ18改為0 , 加指定索引/*+ INDEX(TRADEMARKQUERY IDXTMQ18) */ 加速查詢
     '       'Modified by Lydia 2017/11/06 查覆區欄位順序:委查單號,期限日期,委查日期 (因為智權人員常會打電話催查名人員BY嘉雯)
     '       'strSql = strSql & " Union select /*+ INDEX(TRADEMARKQUERY IDXTMQ18) */ ' ' V,TMQ19,TQA20,DECODE(c1.CP27,NULL,DECODE(TQC07,NULL,'','◎'),'●') C01," & _
     '                "DECODE(c1.CP01,NULL,'',DECODE(c1.CP03||c1.CP04,'000',c1.CP01||'-'||c1.CP02,c1.CP01||'-'||c1.CP02||'-'||c1.CP03||'-'||c1.CP04)) CASENO," & _
     '                "TMQ18,TMQ02,S1.ST02 SST1,DECODE(TQA06,'1',DECODE(TQA13,NULL,'(非正常字)',TQA13),'2','(圖形查詢)') 文字1," & _
     '                "DECODE(TQA06,'1',DECODE(TQA14,NULL,DECODE(TQA08,NULL,DECODE(TQF05,NULL,'','(非正常字)'),'(非正常字)'),TQA14),'2','') 文字2," & _
     '                "TQA04,decode(v1c2," & PUB_GetTMQans("3", True) & ") 結果,TMQ03,TMQ10,S2.ST02 SST02,TMQ01,(TMQ04-19110000) TMQ04,(TMQ06-19110000) TMQ06, (TMQ11-19110000) TMQ11,(TMQ23-19110000) 覆核日期,TMQ21 as 總收文號,(TQC07-19110000) as TMQ20,c1.CP27,c1.CP57,c1.CP14,c1.CP09 as SCP09" & _
     '                " FROM TMQAPP,trademarkquery,STAFF S1,STAFF S2,caseprogress c1,TMQCASEMAP," & _
     '                "(select tqd02 v1c1, " & strExc(2) & " v1c2 from tmqdetail group by tqd02) VT1 " & _
      '               ",(SELECT TQF01,TQF03,TQF05 FROM TMQFILE WHERE TQF02||TQF03||TQF04='" & TMQ_附件F02 & TMQ_AkindWord2 & TMQ_附件F04 & "') VT2 " & _
      '               "WHERE TQA01=TMQ18(+) and tmq01=v1c1(+) AND TMQ02=S1.ST01(+) AND TMQ10=S2.ST01(+) AND NOT(TMQ03 IS NULL) AND TQA01=TQF01(+) AND TMQ11 > 0 AND TMQ19||TQA20 IS NULL " & strCon
      '      strSql = strSql & " Union select /*+ INDEX(TRADEMARKQUERY IDXTMQ18) */ ' ' V,TMQ19,TQA20,DECODE(c1.CP27,NULL,DECODE(TQC07,NULL,'','◎'),'●') C01," & _
      '               "DECODE(c1.CP01,NULL,'',DECODE(c1.CP03||c1.CP04,'000',c1.CP01||'-'||c1.CP02,c1.CP01||'-'||c1.CP02||'-'||c1.CP03||'-'||c1.CP04)) CASENO," & _
      '               "TMQ18,TMQ02,S1.ST02 SST1,DECODE(TQA06,'1',DECODE(TQA13,NULL,'(非正常字)',TQA13),'2','(圖形查詢)') 文字1," & _
      '               "DECODE(TQA06,'1',DECODE(TQA14,NULL,DECODE(TQA08,NULL,DECODE(TQF05,NULL,'','(非正常字)'),'(非正常字)'),TQA14),'2','') 文字2," & _
                     "TQA04,decode(v1c2," & PUB_GetTMQans("3", True) & ") 結果,TMQ03,TMQ10,S2.ST02 SST02,TMQ01,(TMQ06-19110000) TMQ06,(TMQ04-19110000) TMQ04, (TMQ11-19110000) TMQ11,(TMQ23-19110000) 覆核日期,TMQ21 as 總收文號,(TQC07-19110000) as TMQ20,c1.CP27,c1.CP57,c1.CP14,c1.CP09 as SCP09" & _
      '               " FROM TMQAPP,trademarkquery,STAFF S1,STAFF S2,caseprogress c1,TMQCASEMAP," & _
      '               "(select tqd02 v1c1, " & strExc(2) & " v1c2 from tmqdetail group by tqd02) VT1 " & _
      '               ",(SELECT TQF01,TQF03,TQF05 FROM TMQFILE WHERE TQF02||TQF03||TQF04='" & TMQ_附件F02 & TMQ_AkindWord2 & TMQ_附件F04 & "') VT2 " & _
      '               "WHERE TQA01=TMQ18(+) and tmq01=v1c1(+) AND TMQ02=S1.ST01(+) AND TMQ10=S2.ST01(+) AND NOT(TMQ03 IS NULL) AND TQA01=TQF01(+) AND TMQ11 > 0 AND TMQ19||TQA20 IS NULL " & strCon
      '  End If
     'End If
     'end 2025/04/28   ----整理SQL，判斷不用
     
  'Modified by Lydia 2016/03/21 原本先依已讀，再依日期排序＝＞改依日期降冪(最新的再上面)
  '   strSql = strSql & " order by TMQ19 desc,TMQ18,TMQ01"
  'Modified by Lydia 2016/04/28 待查區依期限日期
  If InStr(Me.Caption, "待查區") > 0 Then
     'Modified by Lydia 2016/05/31 未完成的置頂
     'strSql = strSql & " order by TMQ06 ,TMQ01 desc"
     strSql = strSql & " order by tmq11 desc, TMQ06 ,TMQ01 desc"
  Else
     strSql = strSql & " order by TMQ18 desc,TMQ01 desc"
  End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If contCusName <> "" And bolCont = True Then
         textCName.Text = contCusName
         contCusName = ""
      End If
      GRD1.FixedCols = 0
      Set GRD1.Recordset = rsTmp
      SetGrd (rsTmp.RecordCount + 1)
      Select Case R_type
          Case "Q", "M", "A"
               j = PUB_MGridGetId("查名人", GRD1)
          Case "U"
               j = PUB_MGridGetId("委查單號", GRD1)
      End Select
      GRD1.FixedCols = j + 1
   Else
      QueryData = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      If contCusName <> "" And bolCont = True Then
        '從字首4字開始比對到字首2字
         If Len(contCusName) > 2 Then
            contCusName = Mid(contCusName, 1, Len(contCusName) - 1)
            GoTo JumpReSearch
         Else
            contCusName = ""
            GoTo JumpReSearch
         End If
      End If
      SetGrd
      Exit Function
   End If

   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

Private Sub cmdState_Click(Index As Integer)
   'Added by Lydia 2016/05/05
   'Modified by Lydia 2016/06/01 + 本所案號
   If Index <> 1 And Trim(IIf(txtField(0).Visible = True, txtField(0), "") & txtField(1) & txtField(2) & txtField(3) & txtField(4) & textCName & txtField(8) & txtField(9) & txtField(10)) = "" Then
      If MsgBox("是否要輸入查詢條件?", vbInformation + vbYesNo) = vbYes Then
         If txtField(0).Visible = True Then
            txtField(0).SetFocus
         Else
            txtField(1).SetFocus
         End If
         Exit Sub
      End If
   End If
   'end 2016/05/05
   
   
   If Index <> 4 Then
      txtField(6) = Index
      Call cmdQuery_Click
   Else 'Added by Lydia 2016/06/30 重新分查名人
      Call GetNewTMQ10
   End If
End Sub

Private Sub Combo2_Click()
   If Combo2.Tag <> Combo2 Then
      Combo2.Tag = Combo2
   End If
End Sub

Private Sub SetCombo2()
   Dim ii As Integer
   Dim pUser As String
   Dim bUserNo As String
   
   'Modified by Lydia 2016/04/28 改用接洽單的智權人員
   'Combo2.AddItem strUserNum & " " & strUserName
   If stKeyUser <> "" Then
      pUser = stKeyUser
   Else
      pUser = strUserNum
   End If
   
   Combo2.Clear
   'Added by Lydia 2022/01/06 增加從歷程過來看結果的人員：只設定清單=全部
   If bolCaseRead = True Then
       Combo2.AddItem "      " & "全部"
       Combo2.ListIndex = 0
       Exit Sub
   End If
   'end 2022/01/06
   
   'Added by Lydia 2017/04/28 葉特助權限比照商標處主管
   'Remove by Lydia 2020/05/05 葉特助:退休
'   If strUserNum = "67002" Then
'      bUserNo = strUserNum
'      pUser = GetDeptMan("P20")
'      If pUser = "" Then
'         pUser = strUserNum
'      Else
'         Combo2.AddItem bUserNo & " " & GetStaffName(bUserNo)
'         strUserNum = pUser
'      End If
'   End If
'   'end 2017/04/28
   'end 2020/05/05
   
   'Modified by Lydia 2022/05/25 設定屬智權人員作業的下拉選單(共用模組)
'   Combo2.AddItem pUser & " " & GetStaffName(pUser)
'   ''檢查當時是否需要為他人職代
'   Call Pub_SetForOthersEmpCombo(pUser, Combo2, False)
'   ''Modified by Lydia 2020/06/08 +增加特殊權限"AREA"
'   Call Pub_SetSAManageEmpCombo(pUser, Combo2, False, , , "AREA") '如果非主管輸入主管代號,下拉清單有可能帶出其下屬,所以在收文做檢查
'
'   ''專利處智權同仁代處理人
'   If InStr(Pub_GetSpecMan("A8"), pUser) > 0 Or InStr(Pub_GetSpecMan("總經理業務工作代理人員"), pUser) > 0 Then
'      If InStr(Pub_GetSpecMan("A8"), pUser) > 0 Then
'        strSql = "select st01,st02 from setSpecMan,staff where ocode='A7' and instr(';'||replace(oMan,',',';')||';',';'||st01||';')>0"
'      Else
'        strSql = "select st01,st02 from setSpecMan,staff where ocode='總經理員工編號' and instr(';'||replace(oMan,',',';')||';',';'||st01||';')>0"
'      End If
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         RsTemp.MoveFirst
'         Do While Not RsTemp.EOF
'            For ii = 0 To Combo2.ListCount - 1
'               If InStr(Combo2.List(ii), RsTemp(0)) = 1 Then
'                  Exit For
'               End If
'            Next
'            If ii = Combo2.ListCount Then
'               Combo2.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
'            End If
'
'            RsTemp.MoveNext
'         Loop
'      End If
'   End If
'   ''帶人主管抓虛建編號
'   strSql = "select st01,st02 from staff where st01<'63001' and instr(';'||st52||';'||st53||';'||st54||';'||st55||';',';" & pUser & ";')>0"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      RsTemp.MoveFirst
'      Do While Not RsTemp.EOF
'         For ii = 0 To Combo2.ListCount - 1
'            If InStr(Combo2.List(ii), RsTemp(0)) = 1 Then
'               Exit For
'            End If
'         Next
'         If ii = Combo2.ListCount Then
'            Combo2.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
'         End If
'
'         RsTemp.MoveNext
'      Loop
'   End If
'
'   If bUserNo <> "" Then strUserNum = bUserNo 'Added by Lydia 2017/04/28 還原使用者編號
   'Modified by Lydia 2022/05/27 +pUserNo
   Call PUB_SetCombo1Sales(Me.Combo2, pUser)
   'end 2022/05/25
   
   If Pub_StrUserSt03 = "M51" And strUserNum = pUser Then
      Combo2.AddItem "      " & "全部"
   End If
   Combo2.ListIndex = 0
End Sub

Private Sub SetCombo1()
Dim mSQL As String
Dim cInX As Integer 'Added by Lydia 2016/04/21

   Combo1.Clear
   
   'Added by Lydia 2016/04/21 +嘉雯負責管理
   'Modified by Lydia 2016/06/03 改成特殊設定內商查名主管
   'If R_type = "U" And (InStr("67002,69008,84027", strUserNum) = 0 And Pub_StrUserSt03 <> "M51") Then
   If R_type = "U" And (strManUser <> "" And InStr(strManUser, strUserNum) = 0 And Pub_StrUserSt03 <> "M51") Then
      mSQL = " and tmqm01='" & strUserNum & "' "
   End If
   
   strSql = " select tmqm01,st02 from tmqmember,staff where tmqm01=st01(+) and st04='1' " & mSQL & " order by 1 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      'Memo by Lydia 2016/04/21 查名人不可看其他人的委查單,若要參考查名路徑請自行詢問其他人by林經理
      If mSQL = "" Then
         Combo1.AddItem "ALL 全部"
      End If
      Do While Not RsTemp.EOF
         Combo1.AddItem Trim(RsTemp(0)) & " " & Trim(RsTemp(1))
         'Added by Lydia 2016/04/21 預設個人的待查區
         'If InStr("67002,69008,84027", strUserNum) > 0 And Trim(RsTemp(0)) = strUserNum And Me.Caption = "查名/待查區" Then
         If InStr(strManUser, strUserNum) > 0 And Trim(RsTemp(0)) = strUserNum And Me.Caption = "查名/待查區" Then
            cInX = Combo1.ListCount - 1
         End If
         'end 2061/04/21
         RsTemp.MoveNext
      Loop
   End If
   'Added by Lydia 2016/04/21
   If cInX > 0 Then
      Combo1.ListIndex = cInX
   Else
      Combo1.ListIndex = 0
   End If
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
  
    m_AttachPath = App.path & "\" & strUserNum
    If Dir(m_AttachPath, vbDirectory) = "" Then
       MkDir m_AttachPath
    End If
   Call PUB_GetTMQans("1", True) 'Added by Lydia 2016/06/02 求近似本所案
   
   txtField(5).Text = "0" '預設類別:全部
   txtField(6).Text = "0" '預設狀況:未發文
   
   Select Case R_type
       Case "U", "M", "A" '待查(查名人),覆核
            iStiu = 1
            Select Case R_type
                 Case "U"
                    txtField(6).Text = "1" ' 預設處理中,不限日期
                     Me.Caption = "查名/待查區"
                     'Move by Lydia 2016/08/01 移動
                     'strManUser = Pub_GetSpecMan("內商查名主管") 'Added by Lydia 2016/06/03
                 Case "M"
                    txtField(3).Text = strSrvDate(2)
                    txtField(4).Text = TransDate(CompWorkDay(5, strSrvDate(1), 0), 1)
                    Me.Caption = "查名/覆核區"
                 Case "A"
                    txtField(6).Text = "1" ' 預設處理中,不限日期
                    Me.Caption = "查名單維護"
                    cmdState(4).Visible = True 'Added by Lydia 2016/06/30
            End Select
            cmdTo.Visible = False
            cmdSendMail.Visible = False 'Added by Lydia 2016/04/29
            SetCombo1
            Combo2.Visible = False: txtField(0).Visible = True: lblSname.Visible = True
       Case "Q"    '查覆(委查人)
            iStiu = 0
            'Added by Lydia 2016/04/28 從接洽單傳入員工代號
            If stKeyUser <> "" Then
               txtField(0).Text = stKeyUser
            Else
               txtField(0).Text = strUserNum
            End If
            txtField_Validate 0, True
            txtField(1).Text = TransDate(CompWorkDay(1, CompDate(1, -1, strSrvDate(1))), 1)
            txtField(2).Text = strSrvDate(2)
            
            'Modified by Lydia 2019/07/01 更名
            'Me.Caption = "查名/查覆區"
            Me.Caption = "商標查名／查覆區"
            cmdTo.Visible = True
            cmdSendMail.Visible = True 'Added by Lydia 2016/04/29
            SetCombo1
            SetCombo2
            Combo2.Visible = True: txtField(0).Visible = False: lblSname.Visible = False
            'Added by Lydia 2022/01/06 增加從歷程過來看結果的人員：不限人員和日期
            If bolCaseRead = True Then
                txtField(0).Text = ""
                txtField(1).Text = ""
                txtField(2).Text = ""
                '隱藏功能按鈕
                cmdSendMail.Visible = False
                cmdMaster(1).Visible = False
                cmdTo.Visible = False
                cmdMaster(0).Visible = False
                cmdState(0).Visible = False
                cmdState(1).Visible = False
                cmdState(2).Visible = False
                cmdState(3).Visible = False
            End If
            'end 2022/01/06
            
            'Added by Lydia 2019/08/12 創新業務組成員可操作清單
            'Modified by Lydia 2020/12/04 debug-影響到全域變數
            'stIdList = PUB_GetSalesList(txtField(0).Text, , , , , Pub_StrUserSt15, Pub_StrUserSt15)
            'If InStr(stIdList, "W") = 0 Or Left(Pub_StrUserSt15, 1) <> "W" Then
            stIdList = PUB_GetSalesList(txtField(0).Text, , , , , strGrpTmp1, strGrpTmp2)
            If InStr(stIdList, "W") = 0 Or Left(strGrpTmp1, 1) <> "W" Then
            'end 2020/12/04
                stIdList = CNULL(txtField(0).Text) '非創新業務組用切換清單的方式
            End If
            'Added by Lydia 2024/03/22 已新增系統特殊設定「智權部可查詢同所查名單人員」。
                                       '請調整程式，屬於該名單內的人員可查詢同所別智權部同仁的查名單資料。---李承翰
            If UCase(TypeName(m_PrevForm)) = "NOTHING" Then
               strExc(1) = Pub_GetSpecMan("智權部可查詢同所查名單人員")
               If InStr(strExc(1) & ",", strUserNum) > 0 Then
                  strExc(0) = "select st01,st02,st15 from staff where st06='" & pub_strUserOffice & "'" & _
                     " and st15 like 'S%' and st04='1' and st01 in (select distinct(tmq02) tmq02 from trademarkquery) order by st15,st01"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     RsTemp.MoveFirst
                     Do While Not RsTemp.EOF
                        Combo2.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
                        RsTemp.MoveNext
                     Loop
                  End If
                  Combo2.ListIndex = 0
               End If
            End If
            'end 2024/03/22
   End Select
   
   'Added by Lydia 2018/09/20 預設傳入案號為條件
   If mCaseNo(1) <> "" Then
       txtField(9) = mCaseNo(1)
       txtField(10) = mCaseNo(2)
       txtField(11) = mCaseNo(3)
       txtField(12) = mCaseNo(4)
       txtField(6) = IIf(mStiu = "", "2", mStiu) '預設查詢-已完成
       'Added by Lydia 2022/01/06 增加從歷程過來看結果的人員：只限該案件的資料
       If bolCaseRead = True Then
            txtField(9).Enabled = False
            txtField(10).Enabled = False
            txtField(11).Enabled = False
            txtField(12).Enabled = False
            txtField(6).Text = "0" ' 不預設狀態
       End If
       'end 2022/01/06
   End If
   'end 2018/09/20
   
   'Added by Lydia 2019/12/25 開放特殊設定權限
    If CheckLevel(strUserNum, "總經理業務工作代理人員") = True Then
       bolSpecMan = True
       strSpecCode = Pub_GetSpecMan("總經理員工編號")
   '開放專利處部份智權同仁資料給彥葶代為處理
   ElseIf CheckLevel(strUserNum, "A8") = True Then
        bolSpecMan = True
        strSpecCode = Pub_GetSpecMan("A7")
   End If
   'end 2019/12/25
   
   QueryData
   
   'Added by Lydia 2018/09/20 清空預設
   mCaseNo(1) = ""
   mCaseNo(2) = ""
   mCaseNo(3) = ""
   mCaseNo(4) = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If TypeName(m_PrevForm) <> "Nothing" Then
      'Modify By Sindy 2022/9/16 + Or m_PrevForm.Name = "frm090801_New"
      If m_PrevForm.Name = "frm090801" Or m_PrevForm.Name = "frm090801_New" Then
         Call cmdTo_Click
         If mTQD01s = "" Then m_PrevForm.Show
      Else
         m_PrevForm.Show
      End If
      contCusName = ""
      Set m_PrevForm = Nothing
    End If
    
    Set frm090127 = Nothing 'Move by Lydia 2019/07/29 從最上方移下來
End Sub

Private Sub SetGrd(Optional ByVal iR As Integer = 2)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   'Rename by Lydia 2016/07/07 改為colTMQ11
   'Dim iTMQ11 As Integer 'Added by Lydia 2016/05/04
      
   'Modified by Lydia 2016/04/06 以TMQ21為已收文之判斷
   'Modified by Lydia 2016/04/28 隱藏撤回
   'Modified by Lydia 2016/04/29 +送(C01送件狀態),通知送件日=TMQ20,CP27,CP57,CP14
   'Modified by Lydia 2016/06/02 +SCP09
   'arrGridHeadText = Array("V", "讀", "撤", "申請號", "TMQ02", "委查人", "文字1", "文字2", "客戶名稱", "結果", "組群", "TMQ10", "查名人", "委查單號", "委查日期", "期限日期", "查覆日期", "覆核日期", "總收文號")
   '隱藏類別,以文字表達     1    2     3     5     6            7       8        9          10       11          12       13       14       15        16         17         18           19          20          21          22          23       24        25     26      27
   'Added by Lydia 2017/11/06 查覆區欄位順序:委查單號,期限日期,委查日期 (因為智權人員常會打電話催查名人員BY嘉雯)
   If R_type = "Q" Then
                               '1     2     3    4     5            6        7        8         9        10       11          12      13      14       15        16          17          18          19          20          21          22           23      24      25     26
       arrGridHeadText = Array("V", "讀", "撤", "送", "本所案號", "申請號", "TMQ02", "委查人", "文字1", "文字2", "客戶名稱", "結果", "組群", "TMQ10", "查名人", "委查單號", "期限日期", "委查日期", "查覆日期", "覆核日期", "總收文號", "通知送件日", "CP27", "CP57", "CP14", "SCP09")
   Else
   'end 2017/11/06
       arrGridHeadText = Array("V", "讀", "撤", "送", "本所案號", "申請號", "TMQ02", "委查人", "文字1", "文字2", "客戶名稱", "結果", "組群", "TMQ10", "查名人", "委查單號", "委查日期", "期限日期", "查覆日期", "覆核日期", "總收文號", "通知送件日", "CP27", "CP57", "CP14", "SCP09")
   End If
   Select Case R_type
        Case "U"
           '隱藏查名人
           'Modified by Lydia 2016/04/21 有查閱全部人員的顯示查名人員
           'Modified by Lydia 2016/04/28 隱藏申請號和委查單號
           'Modified by Lydia 2016/04/29 +送(C01送件狀態),CASENO,TMQ20,CP27,CP57
           'If InStr("67002,69008,84027", strUserNum) > 0 Or Pub_StrUserSt03 = "M51" Then
           If InStr(strManUser, strUserNum) > 0 Or Pub_StrUserSt03 = "M51" Then
                                       '1   2    3  4  5    6  7  8    9    10   11   12   13   14 15   16 17   18   19   20   21 22 23 24 25 26
              arrGridHeadWidth = Array(200, 300, 0, 0, 900, 0, 0, 800, 860, 860, 900, 800, 900, 0, 820, 0, 820, 820, 820, 820, 0, 0, 0, 0, 0, 0)
           Else
              arrGridHeadWidth = Array(200, 300, 0, 0, 900, 0, 0, 800, 1000, 1000, 900, 800, 900, 0, 0, 0, 820, 820, 820, 820, 0, 0, 0, 0, 0, 0)
           End If
           Label9.Caption = "期限日期紅色:當天或過期"
        Case "M", "A"
           arrGridHeadWidth = Array(200, 300, 0, 300, 900, 0, 0, 800, 1000, 1000, 900, 800, 900, 0, 820, 960, 820, 820, 820, 820, 0, 1000, 0, 0, 0, 0)
        'Modified by Lydia 2022/01/06 +""空白
        Case "Q", ""
           'Modified by Lydia 2016/04/29 隱藏期限日期
           'Modified by Lydia 2017/11/06 顯示期限日期
           'arrGridHeadWidth = Array(200, 300, 0, 300, 900, 0, 0, 0, 1000, 1000, 900, 800, 900, 0, 820, 960, 820, 0, 820, 820, 0, 1000, 0, 0, 0, 0)
                                   '1     2   3  4     5   6  7  8   9   10   11   12   13   14  15  16  17    18   19   20   21 22   23 24 25 26
           arrGridHeadWidth = Array(200, 300, 0, 300, 900, 0, 0, 0, 940, 940, 900, 800, 900, 0, 700, 960, 820, 820, 820, 820, 0, 1000, 0, 0, 0, 0)
   End Select
   
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = iR
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      'Remove by Lydia 2016/04/28 取消撤回的顏色
      'If iRow = 2 Then
      '   GRD1.CellForeColor = QBColor(13)
      'Else
      '   GRD1.CellForeColor = QBColor(0)
      'End If
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   
   If colMno = 0 Then 'Added by Lydia 2021/02/17 加判斷
        colMno = PUB_MGridGetId("申請號", GRD1)
        colAno = PUB_MGridGetId("委查單號", GRD1)
        colCno = PUB_MGridGetId("撤", GRD1)
        colCp09 = PUB_MGridGetId("總收文號", GRD1)
        'Added by Lydia 2016/04/29
        colTMQ02 = PUB_MGridGetId("TMQ02", GRD1)
        colState = PUB_MGridGetId("送", GRD1)
        colCP57 = PUB_MGridGetId("CP57", GRD1)
        colCase = PUB_MGridGetId("本所案號", GRD1)
        colCP14 = PUB_MGridGetId("CP14", GRD1)
        'Added by Lydia 2016/05/04
        colTMQ06 = PUB_MGridGetId("期限日期", GRD1)
        colTMQ11 = PUB_MGridGetId("查覆日期", GRD1)
        colShowCP09 = PUB_MGridGetId("SCP09", GRD1) 'Added by Lydia 2016/06/02
        'Added by Lydia 2021/02/17
        colTMQ23 = PUB_MGridGetId("覆核日期", GRD1)
        colTQD0609 = PUB_MGridGetId("結果", GRD1)
        'end 2021/02/17
   End If 'Added by Lydia 2021/02/17
   
   For intI = 1 To iR - 1
     GRD1.row = intI
     For iRow = 0 To colAno
       GRD1.col = iRow
       GRD1.CellBackColor = QBColor(15)
       'Remove by Lydia 2016/04/28 取消撤回的顏色
       'If iRow = colCno Then
       '   GRD1.CellForeColor = QBColor(13)
       'End If
     Next iRow
     'Added by Lydia 2016/05/04 查名期限的當天和過期設底色
     If R_type = "U" And (txtField(6) = "0" Or txtField(6) = "1") And Trim(GRD1.TextMatrix(intI, colTMQ11)) = "" Then
        If Trim(GRD1.TextMatrix(intI, colTMQ06)) <= strSrvDate(2) Then
           GRD1.col = colTMQ06
           GRD1.CellBackColor = &HC0C0FF
        End If
     End If
     'end 2016/05/04
     'Added by Lydia 2021/02/17 優化顯示需進行覆核流程之查名單 , 以利委查同仁瞭解與本所近似查名單之狀態
     '1.查名結果為相同△或近似△時，該查名單之結果以黃色顯示，即表示進行覆核中
     '2.覆核結果為相同△或近似△時，該查名單之結果以紅色顯示，即表示已覆核，仍與本所近似
     If InStr(Trim(GRD1.TextMatrix(intI, colTQD0609)) & ";", "△") > 0 Then
        GRD1.col = colTQD0609
        If Val("" & Trim(GRD1.TextMatrix(intI, colTMQ23))) = 0 Then  '覆核中
           GRD1.CellBackColor = &HFFFF&
        Else    '已覆核
           GRD1.CellBackColor = &HFF&
        End If
     End If
     'end 2021/02/17
   Next intI

   GRD1.Visible = True

End Sub

Private Sub grd1_SelChange()
Dim TmpRow As Integer
TmpRow = GRD1.MouseRow

If TmpRow > 0 Then
   If GRD1.TextMatrix(TmpRow, 0) = "V" Then
      '清空資料
      GRD1.col = 0
      GRD1.row = TmpRow
      GRD1.Text = ""
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         'Modified by Lydia 2016/05/04 查名期限的當天和過期設底色不變
         'GRD1.CellBackColor = QBColor(15)
         'Modified by Lydia 2021/02/17 +優化顯示需進行覆核流程之查名單：設底色不變
         'If Not (R_type = "U" And (txtField(6) = "0" Or txtField(6) = "1") And i = colTMQ06) Then
         If Not ((R_type = "U" And (txtField(6) = "0" Or txtField(6) = "1") And i = colTMQ06) _
               Or (i = colTQD0609 And InStr(Trim(GRD1.TextMatrix(TmpRow, colTQD0609)) & ";", "△") > 0)) Then
            GRD1.CellBackColor = QBColor(15)
         End If
      Next i
   Else
      '目前資料列反白
      GRD1.col = 0
      GRD1.row = TmpRow
      dblPrevRow = GRD1.row

      If GRD1.TextMatrix(GRD1.row, colAno) <> "" Then
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            'Modified by Lydia 2016/05/04 查名期限的當天和過期設底色不變
            'GRD1.CellBackColor = &HFFC0C0
            'Modified by Lydia 2021/02/17 +優化顯示需進行覆核流程之查名單：設底色不變
            'If Not (R_type = "U" And (txtField(6) = "0" Or txtField(6) = "1") And i = colTMQ06) Then
            If Not ((R_type = "U" And (txtField(6) = "0" Or txtField(6) = "1") And i = colTMQ06) _
                  Or (i = colTQD0609 And InStr(Trim(GRD1.TextMatrix(TmpRow, colTQD0609)) & ";", "△") > 0)) Then
               GRD1.CellBackColor = &HFFC0C0
            End If
         Next i
      End If
   End If
End If

End Sub

Private Sub GRD1_DblClick()
   cmdDetail_Click
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
      'Modified by Lydia 2016/04/29
      'If InStr("委查日期,期限日期,查覆日期,覆核日期", Me.GRD1.Text) > 0 Then
      If InStr("委查日期,期限日期,查覆日期,覆核日期,通知送件日", Me.GRD1.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
txtField(Index).SelStart = 0
txtField(Index).SelLength = Len(txtField(Index).Text)
'Mark by Lydia 2016/10/28 受win7輸入法影響,不切換輸入法
'If Index <> 7 Then
'   CloseIme
'End If
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
       'Added by Lydia 2016/06/01 + 9
       Case 0, 8, 9
            KeyAscii = UpperCase(KeyAscii)
       Case 7
       Case Else
            KeyAscii = Pub_NumAscii(KeyAscii)
   End Select
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
If txtField(Index) <> "" Then
    Select Case Index
        Case 0
            If ClsPDGetStaff(txtField(Index).Text, strExc(1)) Then
               lblSname.Caption = strExc(1)
            Else
                GoTo RetErr
            End If
        Case 1, 2
            If CheckIsTaiwanDate(txtField(Index)) = False Then
                GoTo RetErr
            End If
        Case 3, 4
            If CheckIsTaiwanDate(txtField(Index)) = False Then
                GoTo RetErr
            End If
        Case 5
             If txtField(Index) <> "0" And txtField(Index) <> "1" And txtField(Index) <> "2" Then
                MsgBox "請輸入0-2", vbOKOnly, "輸入錯誤"
                GoTo RetErr
             End If
        Case 6
             If txtField(Index) <> "0" And txtField(Index) <> "1" And txtField(Index) <> "2" And txtField(Index) <> "3" Then
                MsgBox "請輸入0-3", vbOKOnly, "輸入錯誤"
                GoTo RetErr
             End If
        Case 8
             If Left(txtField(Index), 1) <> "H" Then
                MsgBox "委查單號輸入錯誤", vbOKOnly, "輸入錯誤"
                GoTo RetErr
             End If
        'Added by Lydia 2016/06/01
        Case 9
            If txtField(Index) <> "T" And txtField(Index) <> "TS" Then
               MsgBox "查名案件為T或TS案!", vbOKOnly, "輸入錯誤"
               GoTo RetErr
            Else
               '預設清空日期條件
               For intI = 1 To 4
                  txtField(intI).Text = ""
               Next
            End If
        Case 10
            If Len(Trim(txtField(Index))) <> 6 Then
               MsgBox "請輸入案號!", vbOKOnly, "輸入錯誤"
               GoTo RetErr
            End If
        'end 2016/06/01
    End Select
Else
    If Index = 0 Then lblSname.Caption = ""
    'Added by Lydia 2016/06/01
    If Index = 9 Then txtField(10).Text = "": txtField(11).Text = "": txtField(12).Text = ""
End If

Exit Sub

RetErr:
    txtField(Index).SetFocus
    Cancel = True
End Sub
Private Sub cmdTo_Click()
Dim tmpList As String
Dim strTmp1 As String 'Added by Lydia 2019/08/12

'Dim idR1 As Integer, idR2 As Integer
   m_TMQApp = ""
   m_TMQNo = ""
   
   'idR1 = PUB_MGridGetId("查覆日期", GRD1)
   'idR2 = PUB_MGridGetId("讀", GRD1)
   
   'Added by Lydia 2016/05/03 控制主管不可收下屬員工
   'Modified by Lydia 2019/08/12 增加創新業務組成員可互相操作
   'If Trim(Mid(Combo2.Text, 1, InStr(Combo2.Text, " "))) <> strUserNum And cmdTo.Visible = True Then
   strTmp1 = Trim(Mid(Combo2.Text, 1, InStr(Combo2.Text, " ")))
   'Modified by Lydia 2019/12/25 開放特殊設定權限
   'If strTmp1 <> strUserNum And InStr(stIdList, strTmp1) = 0 And cmdTo.Visible = True Then
   If strTmp1 <> strUserNum And cmdTo.Visible = True Then
      strExc(1) = "N"
      If InStr(stIdList, strTmp1) > 0 Then
          strExc(1) = ""
      '代理-總經理、A7
      ElseIf bolSpecMan = True And InStr(strSpecCode, strTmp1) > 0 Then
          strExc(1) = ""
      End If
      If strExc(1) = "N" Then
      'end 2019/12/25
         MsgBox "無權限!!", vbCritical
         Exit Sub
      End If
   End If
   
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
'         If TMQ_CtrRead Then '控制是否已讀
'            If GRD1.TextMatrix(i, idR1) = "" Then
'               MsgBox "委查單: " & Trim(GRD1.TextMatrix(i, colAno)) & " 尚未查覆完成!", vbCritical
'               Exit Sub
'            End If
'            If GRD1.TextMatrix(i, idR2) = "" Then
'               MsgBox "請先查看委查單: " & Trim(GRD1.TextMatrix(i, colAno)) & " 的查覆附件!", vbCritical
'               Exit Sub
'            End If
'         End If
         If GRD1.TextMatrix(i, colCno) = "Y" Then
            MsgBox "委查單: " & Trim(GRD1.TextMatrix(i, colAno)) & " 已撤回", vbCritical
            Exit Sub
         End If
         'Modified by Lydia 2016/04/28
         'Added by Lydia 2016/05/03 判斷接洽單傳入的智權人員
         'Modified by Lydia 2019/08/12 增加創新業務組成員可互相操作
         'If Trim(Mid(Combo2.Text, 1, InStr(Combo2.Text, " "))) <> Trim(GRD1.TextMatrix(i, colTMQ02)) Or (stKeyUser <> "" And stKeyUser <> Trim(Mid(Combo2.Text, 1, InStr(Combo2.Text, " ")))) Then
         If InStr(strTmp1 & "," & stIdList, Trim(GRD1.TextMatrix(i, colTMQ02))) = 0 _
             Or (stKeyUser <> "" And stKeyUser <> strTmp1) Then
            MsgBox "委查單: " & Trim(GRD1.TextMatrix(i, colAno)) & " 不是你申請的!", vbCritical
            Exit Sub
         End If
         '判斷是否可重複收文
         If TMQ_ReApp = False And Len(Trim(GRD1.TextMatrix(i, colCp09))) > 0 Then
            MsgBox "委查單: " & Trim(GRD1.TextMatrix(i, colAno)) & " 已收文，不可再次收文!", vbCritical
            Exit Sub
         End If
         If InStr(tmpList, Trim(GRD1.TextMatrix(i, colMno))) = 0 Then
            tmpList = tmpList & Trim(GRD1.TextMatrix(i, colMno)) & ","
         End If
      End If
   Next i

   mTQD01s = tmpList
   
   'Added by Lydia 2023/08/02
   If pTMQList <> "" And mTQD01s = "" Then
      If MsgBox("沒有勾選查名單，是否繼續回存到接洽單？", vbExclamation + vbYesNo + vbDefaultButton2, "接洽單修改") = vbNo Then
         Exit Sub
      End If
   End If
   'end 2023/08/02
   
   PubShowNextData

End Sub
'Added by Lydia 2016/04/29
Private Sub cmdSendMail_Click()
Dim tmpStr As String
Dim strAll As String
Dim strNoList As String
Dim tmpArr As Variant
Dim id1 As Integer, id2 As Integer
Dim strTmp1 As String 'Added by Lydia 2019/08/12

On Error GoTo ErrHand01
   'Modified by Lydia 2019/08/12 增加創新業務組成員可互相操作
   'If Trim(Mid(Combo2.Text, 1, InStr(Combo2.Text, " "))) <> strUserNum And cmdSendMail.Visible = True Then
   strTmp1 = Trim(Mid(Combo2.Text, 1, InStr(Combo2.Text, " ")))
   'Modified by Lydia 2019/12/25 開放特殊設定權限
   'If strTmp1 <> strUserNum And InStr(stIdList, strTmp1) = 0 And cmdSendMail.Visible = True Then
   If strTmp1 <> strUserNum And cmdSendMail.Visible = True Then
      strExc(1) = "N"
      If InStr(stIdList, strTmp1) > 0 Then
          strExc(1) = ""
      '代理-總經理、A7
      ElseIf bolSpecMan = True And InStr(strSpecCode, strTmp1) > 0 Then
          strExc(1) = ""
      End If
      If strExc(1) = "N" Then
      'end 2019/12/25
          MsgBox "無權限!!", vbCritical
          Exit Sub
      End If
   End If
   id1 = PUB_MGridGetId("客戶名稱", GRD1)
   id2 = PUB_MGridGetId("文字1", GRD1)
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         If GRD1.TextMatrix(i, colCno) = "Y" Then
            MsgBox "委查單: " & Trim(GRD1.TextMatrix(i, colAno)) & " 已撤回", vbCritical
            Exit Sub
         End If
         'Modified by Lydia 2019/08/12 增加創新業務組成員可互相操作
         'If Trim(Mid(Combo2.Text, 1, InStr(Combo2.Text, " "))) <> Trim(GRD1.TextMatrix(i, colTMQ02)) Then
         'Modified by Lydia 2019/12/25 開放特殊設定權限
         'If InStr(strTmp1 & "," & stIdList, Trim(GRD1.TextMatrix(i, colTMQ02))) = 0 Then
         If InStr(strTmp1 & "," & stIdList & IIf(bolSpecMan = True, "," & strSpecCode, ""), Trim(GRD1.TextMatrix(i, colTMQ02))) = 0 Then
            MsgBox "委查單: " & Trim(GRD1.TextMatrix(i, colAno)) & " 不是你申請的!", vbCritical
            Exit Sub
         End If
         If Trim(GRD1.TextMatrix(i, colCp09)) = "" Then
            MsgBox "委查單: " & Trim(GRD1.TextMatrix(i, colAno)) & " 未收文", vbCritical
            Exit Sub
         End If
         If Trim(GRD1.TextMatrix(i, colState)) <> "" Then
            MsgBox "本所案號: " & Trim(GRD1.TextMatrix(i, colCase)) & " " & IIf(Trim(GRD1.TextMatrix(i, colState)) = "◎", "已通知送件", "已發文"), vbCritical
            Exit Sub
         End If
         If Trim(GRD1.TextMatrix(i, colCP57)) <> "" Then
            MsgBox "本所案號: " & Trim(GRD1.TextMatrix(i, colCase)) & " 已取消收文", vbCritical
            Exit Sub
         End If
         If Trim(GRD1.TextMatrix(i, colCP14)) = "" Then
            MsgBox "本所案號: " & Trim(GRD1.TextMatrix(i, colCase)) & " 未分案", vbCritical
            Exit Sub
         End If
         'Added by Lydia 2021/03/08 近似本所案不可通知送件
         If InStr(Trim(GRD1.TextMatrix(i, colTQD0609)) & ";", "△") > 0 Then
            MsgBox "本所案號: " & Trim(GRD1.TextMatrix(i, colCase)) & " 尚在進行覆核流程", vbCritical
            Exit Sub
         End If
         'end 2021/03/08
         
         If InStr(strAll, Trim(GRD1.TextMatrix(i, colCase))) = 0 Then
            tmpStr = "" & Trim(GRD1.TextMatrix(i, id2))
            If InStr(tmpStr, "(圖形查詢)") = 0 Then
               tmpStr = "(" & tmpStr & IIf("" & Trim(GRD1.TextMatrix(i, id2 + 1)) <> "", "," & Trim(GRD1.TextMatrix(i, id2 + 1)), "") & ")"
            End If
            '只傳案號
            'strAll = strAll & Trim(GRD1.TextMatrix(i, colCase)) & "「" & Trim(GRD1.TextMatrix(i, id1)) & "」" & tmpStr & "||" & Trim(GRD1.TextMatrix(i, colCP14)) & ","
            strAll = strAll & Trim(GRD1.TextMatrix(i, colCase)) & "||" & Trim(GRD1.TextMatrix(i, colCP14)) & ","
         End If
         'Modified by Lydia 2016/07/06 改抓收文號
         strNoList = strNoList & Trim(GRD1.TextMatrix(i, colCp09)) & ","
      End If
   Next i
   'Added by Lydia 2016/07/07 判斷委查單是否查覆完畢
   If PUB_TMQCheckOver(strNoList) = False Then
      Exit Sub
   End If
   'end 2016/07/07
   
   If strAll <> "" Then
      tmpArr = Empty
      tmpArr = Split(strAll, ",")
      For i = 0 To UBound(tmpArr)
         If InStr(tmpArr(i), "||") > 0 Then
            PUB_SendMail Trim(Mid(Combo2.Text, 1, InStr(Combo2.Text, " "))), Mid(tmpArr(i), InStr(tmpArr(i), "||") + 2), "", Mid(tmpArr(i), 1, InStr(tmpArr(i), "||") - 1) & "案，經智權人員確認，請送件!", vbCrLf & "如主旨"
         End If
      Next i
      
      '同一收文號，只通知一次
      'Memo by Lydia 2016/05/10 同一查名做兩個以上的申請,因為分案會分給同一承辦人,所以不用發兩封信
      'Memo by Lydia 2016/07/07 若有追加查名結果，可再通知
      cnnConnection.BeginTrans
         'Modified by Lydia 2016/07/06 改成TQC07
         'strSql = "UPDATE TRADEMARKQUERY SET TMQ20=" & strSrvDate(1) & " WHERE TMQ21 IN (SELECT TMQ21 FROM TRADEMARKQUERY WHERE TMQ01 IN (" & GetAddStr(strNoList) & ") AND TMQ21 IS NOT NULL)"
         strSql = "UPDATE TMQCASEMAP SET TQC07=" & strSrvDate(1) & " WHERE TQC02 IN (" & GetAddStr(strNoList) & ") AND TQC07 IS NULL "
         cnnConnection.Execute strSql, i
      cnnConnection.CommitTrans
      
      If QueryData = False Then ShowNoData
   End If
   
   Exit Sub
ErrHand01:
   
   MsgBox Err.Description, vbCritical
   If strAll <> "" Then cnnConnection.RollbackTrans
   
End Sub

Public Sub PubShowNextData()
Dim i As Integer, j As Integer
Dim mLoad As Boolean
Dim sPath As String
Dim APKind As String
Dim rsR As New ADODB.Recordset
Dim oRunform As Form 'Add By Sindy 2022/9/16
   
   'Add By Sindy 2022/9/16
   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
      Set oRunform = frm090801_New
   Else
      Set oRunform = frm090801
   End If
   '2022/9/16 END
   
If m_TMQApp <> "" Then
    m_TMQApp = ""
    QueryData '重整資料
'Modified by Lydia 2023/08/02
'ElseIf mTQD01s <> "" Then
Else
   If mTQD01s <> "" Then
'end 2023/08/02
       Me.Enabled = False: mLoad = False
       Screen.MousePointer = vbHourglass
       APKind = "HM" 'm_TMQApp '先代入申請號
       'Modify By Sindy 2022/9/16 frm090801 改用 oRunform
       If bolCont = False Then
          oRunform.SetParent Me
          oRunform.bolExternalCall = True '記錄是外部程式呼叫使用
       End If
       
       oRunform.Show
       
       If bolCont = False Then
           oRunform.Option1(0).Value = True '新案
           oRunform.Text1(6) = "T" '商標案
           Call oRunform.Text1_LostFocus(9)
       End If
   
       mLoad = False
       '以第一個申請編號的查詢內容為主
          strExc(1) = GetAddStr(mTQD01s)
          strExc(0) = "select a.*,tqf03,tqf05 from tmqapp a ,tmqfile f " & _
                      "where tqa01 in (" & strExc(1) & ") and tqa01=tqf01(+) and tqf02(+)='" & TMQ_附件F02 & "' and tqf04(+)='" & TMQ_附件F04 & "'"
          'Modified by Lydia 2016/04/18 以最新單號
          strExc(0) = strExc(0) & " order by tqa01 desc,tqf03 "
          
           intI = 1
           Set rsR = ClsLawReadRstMsg(intI, strExc(0))
           If intI = 1 Then
              rsR.MoveFirst
              APKind = rsR.Fields("tqa01")
              m_TMQNo = rsR.Fields("tqa01") & " " & rsR.Fields("tqa06")
              txtUnicode(1) = ""
              Do While Not rsR.EOF
                 If m_TMQNo <> rsR.Fields("tqa01") & " " & rsR.Fields("tqa06") Then Exit Do
                 
                 If rsR.Fields("tqf03") = TMQ_AkindPic Then '圖形
                    If InStr(UCase(rsR.Fields("tqf05")), "PDF") = 0 Then
                       mLoad = True
                       'Modified by Lydia 2016/06/23
                       'APKind = APKind & "_" & TMQ_AkindPic
                       APKind = APKind & TMQ_附件F02 & TMQ_AkindPic & TMQ_附件F04
                    End If
                 Else
                    '文字
                    If rsR.AbsolutePosition = 1 Then
                       If Len(rsR.Fields("tqa07")) > 0 Then
                          txtUnicode(1) = txtUnicode(1) & rsR.Fields("tqa07") & " "
                       ElseIf Len(rsR.Fields("tqa13")) > 0 Then
                          txtUnicode(1) = txtUnicode(1) & rsR.Fields("tqa13") & " "
                       End If
                       If Len(rsR.Fields("tqa08")) > 0 Then
                          txtUnicode(1) = txtUnicode(1) & rsR.Fields("tqa08") & " "
                       ElseIf Len(rsR.Fields("tqa14")) > 0 Then
                          txtUnicode(1) = txtUnicode(1) & rsR.Fields("tqa14") & " "
                       End If
                    End If
                    
                    If rsR.Fields("tqf03") = TMQ_AkindWord1 And InStr(UCase(rsR.Fields("tqf05")), "PDF") = 0 Then
                          mLoad = True
                          'Modified by Lydia 2016/06/23
                          'APKind = APKind & "_" & TMQ_AkindWord1
                          APKind = APKind & TMQ_附件F02 & TMQ_AkindWord1 & TMQ_附件F04
                    ElseIf rsR.Fields("tqf03") = TMQ_AkindWord2 And InStr(UCase(rsR.Fields("tqf05")), "PDF") = 0 And mLoad = False Then
                          mLoad = True
                          'Modified by Lydia 2016/06/23
                          'APKind = APKind & "_" & TMQ_AkindWord2
                          APKind = APKind & TMQ_附件F02 & TMQ_AkindWord2 & TMQ_附件F04
                    End If
                 End If
                 rsR.MoveNext
              Loop
           End If
           
           If txtUnicode(1) <> "" Then
             oRunform.opt1(0).Value = True
             'oRunform.PicText = txtUnicode(1) 'Mark by Lydia 2024/10/07 商標文字欄位中，勿直接帶入文字，以留空方式讓智權人員填寫---杜協理
           ElseIf mLoad = True Then
             sPath = Dir(m_AttachPath & "\" & APKind & "*.*")
             If sPath = "" Then
                mLoad = KeyFileGet(Left(m_TMQNo, 10), Right(APKind, 1), False, sPath)
             Else
                sPath = m_AttachPath & "\" & sPath
             End If
             If mLoad = True Then
                oRunform.opt1(1).Value = True
                oRunform.optColor(0).Value = True
                Call oRunform.PicToObj(sPath)
             End If
           End If
       
       m_TMQApp = mTQD01s: m_TMQNo = "ALL"
       oRunform.cmdTMQ.Tag = mTQD01s
       oRunform.Combo1(0).Text = "000" & " " & GetPrjNationName("000")
      '設定案件性質
       Call oRunform.Text1_LostFocus(6)
       Call oRunform.QueryTMQ
       'Added by Lydia 2016/07/12 TS案無商標種類
       If oRunform.Text1(6) = "TS" Then
       
       ElseIf oRunform.Text1(6) = "T" Then
          oRunform.Combo6.ListIndex = 0 'Added by Lydia 2016/05/30 接洽單的商標種類
       End If
       oRunform.bolExternalCall = False '還原預設值
       Screen.MousePointer = vbDefault
       Me.Enabled = True
       Me.Hide
   'Added by Lydia 2023/08/02 沒有勾選查名單，繼續回存到接洽單
   Else
       If pTMQList <> "" Then
         If bolCont = False Then
            oRunform.SetParent Me
            oRunform.bolExternalCall = True '記錄是外部程式呼叫使用
         End If
         oRunform.Show
         oRunform.cmdTMQ.Tag = ""
        '設定案件性質
         Call oRunform.Text1_LostFocus(6)
         Call oRunform.QueryTMQ
         oRunform.bolExternalCall = False '還原預設值
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         Me.Hide
       End If
   End If
   'end 2023/08/02
   
End If
End Sub

Private Function KeyFileGet(mTQF01 As String, mTQFkind As String, Optional mLoad As Boolean = True, Optional mPath As String) As Boolean
Dim adoRst As New ADODB.Recordset
Dim outType As String
Dim stTempFile As String
Dim fileN As Integer
Dim bytes() As Byte

On Error GoTo ErrHnd
   
    KeyFileGet = False
    '開啟時,無法刪除,預設下次開啟表單執行刪檔
    If Dir(m_AttachPath & "\HM*.jpg") <> "" Then
        Kill m_AttachPath & "\HM*.jpg"
    End If
    If Dir(m_AttachPath & "\HM*.pdf") <> "" Then
        Kill m_AttachPath & "\HM*.pdf"
    End If
    If Dir(m_AttachPath & "\HM*.JPG") <> "" Then
        Kill m_AttachPath & "\HM*.JPG"
    End If
    If Dir(m_AttachPath & "\HM*.PDF") <> "" Then
        Kill m_AttachPath & "\HM*.PDF"
    End If
    If adoRst.State <> adStateClosed Then adoRst.Close
    Set adoRst = Nothing
    adoRst.CursorLocation = adUseClient
    adoRst.Open "select * from TMQFile where TQF01='" & mTQF01 & "' AND TQF02='" & TMQ_附件F02 & "' AND TQF03='" & mTQFkind & "' AND TQF04='" & TMQ_附件F04 & "'", cnnConnection, adOpenStatic, adLockOptimistic
    If adoRst.RecordCount > 0 Then

       outType = "" & adoRst.Fields("TQF05")
       'Modified by Lydia 2016/06/23
       'stTempFile = m_AttachPath & "\" & mTQF01 & "_" & mTQFkind & "." & LCase(Trim(outType))
       stTempFile = m_AttachPath & "\" & mTQF01 & TMQ_附件F02 & mTQFkind & TMQ_附件F04 & "." & LCase(Trim(outType))
              
       mPath = stTempFile
       
       'Modified by Lydia 2016/06/23 改放在FTP
       'Remove by Lydia 2016/07/07
       'If strSrvDate(1) < TMQFileFTP Then
       '     ReDim bytes(Val(adoRst.Fields("TQF06").Value))
       '     bytes() = adoRst.Fields("TQF07").GetChunk(Val(adoRst.Fields("TQF06").Value))
       '     fileN = FreeFile
       '     Open stTempFile For Binary Access Write As #fileN
       '     Put #fileN, , bytes()
       '     Close #fileN
       'Else
            If PUB_TMQGetAFile(m_AttachPath, stTempFile, mTQF01, TMQ_附件F02, mTQFkind, TMQ_附件F04, outType) = False Then
               MsgBox "無法儲存檔案[ " & stTempFile & " ]！"
               Exit Function
            End If
       'End If
       ''end 2016/06/23
       
       If InStr(UCase(outType), "PDF") = 0 Then
          Set G_SeekPicColor.Picture = pvGetStdPicture(Trim(stTempFile))
          Call Pub_PicToObj(Trim(stTempFile), G_SeekPicColor, tmpPic, tmpImg)
       End If
       
    Else
       Exit Function
    End If
   
    KeyFileGet = True
    Exit Function

ErrHnd:
   MsgBox Err.Description, vbCritical
   
   If fileN > 0 Then Close #fileN
End Function

'Added by Lydia 2021/10/01
Private Sub textCName_GotFocus()
    TextInverse textCName
End Sub

'Added by Lydia 2021/10/01
Private Sub textCName_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Forms(0).PopupMenu2 textCName
End Sub

'Added by Lydia 2021/11/12 取得人員請假的職代
Private Function GetDutyList(ByVal stIdList As String) As String
Dim tmpArr As Variant
Dim inX As Integer
Dim stTmp1 As String
      
    GetDutyList = ""
    If stIdList <> "" Then
        tmpArr = Split(stIdList, ";")
        For inX = 0 To UBound(tmpArr)
            If Trim(tmpArr(inX)) <> "" Then
                stTmp1 = GetCaseDutyAgent(tmpArr(inX), "", False, , True, "A") 'A.指定抓全部職代
                If stTmp1 <> "" Then
                    GetDutyList = GetDutyList & ";" & stTmp1
                End If
            End If
        Next inX
    End If
    If GetDutyList <> "" Then GetDutyList = Mid(GetDutyList, 2)
End Function

