VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090127_New 
   BorderStyle     =   1  '單線固定
   Caption         =   "查覆區(網中)"
   ClientHeight    =   7152
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   9888
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7152
   ScaleWidth      =   9888
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   1236
      Left            =   2952
      TabIndex        =   48
      Top             =   5760
      Width           =   6828
      Begin VB.CommandButton cmdPost 
         Caption         =   "測試傳送"
         Height          =   324
         Left            =   2904
         TabIndex        =   49
         Top             =   216
         Width           =   972
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   900
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   120
         Width           =   2556
         VariousPropertyBits=   -1467987941
         Size            =   "4508;1587"
         Value           =   "全家Family"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   204
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   900
         Index           =   1
         Left            =   4176
         TabIndex        =   51
         Top             =   120
         Width           =   2556
         VariousPropertyBits=   -1467987941
         Size            =   "4508;1587"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   204
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label14 
         Caption         =   "＞＞"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   13.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   276
         Left            =   2952
         TabIndex        =   50
         Top             =   576
         Width           =   852
      End
   End
   Begin VB.CommandButton cmdState 
      BackColor       =   &H00FFFFC0&
      Caption         =   "重新分查名人"
      Height          =   280
      Index           =   4
      Left            =   4200
      Style           =   1  '圖片外觀
      TabIndex        =   44
      Top             =   1416
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
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   1416
      Width           =   900
   End
   Begin VB.CommandButton cmdMaster 
      Caption         =   "查名單輸入"
      Height          =   360
      Index           =   1
      Left            =   5688
      TabIndex        =   16
      Top             =   1416
      Width           =   1100
   End
   Begin VB.CommandButton cmdState 
      Caption         =   "已發文"
      Height          =   360
      Index           =   3
      Left            =   6960
      Style           =   1  '圖片外觀
      TabIndex        =   23
      Top             =   12
      Width           =   800
   End
   Begin VB.CommandButton cmdState 
      Caption         =   "已完成"
      Height          =   360
      Index           =   2
      Left            =   6120
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   12
      Width           =   800
   End
   Begin VB.CommandButton cmdState 
      Caption         =   "處理中"
      Height          =   360
      Index           =   1
      Left            =   5280
      Style           =   1  '圖片外觀
      TabIndex        =   21
      Top             =   12
      Width           =   800
   End
   Begin VB.CommandButton cmdState 
      Caption         =   "未發文"
      Height          =   360
      Index           =   0
      Left            =   4440
      Style           =   1  '圖片外觀
      TabIndex        =   20
      Top             =   12
      Width           =   800
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   8
      Left            =   5160
      TabIndex        =   9
      Top             =   972
      Width           =   2535
   End
   Begin VB.PictureBox G_SeekPicColor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Enabled         =   0   'False
      Height          =   300
      Left            =   10392
      ScaleHeight     =   21
      ScaleMode       =   3  '像素
      ScaleWidth      =   21
      TabIndex        =   35
      Top             =   216
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox tmpPic 
      Height          =   4455
      Left            =   10344
      ScaleHeight     =   367
      ScaleMode       =   3  '像素
      ScaleWidth      =   295
      TabIndex        =   34
      Top             =   672
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
      Left            =   6888
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   1416
      Width           =   900
   End
   Begin VB.CommandButton cmdMaster 
      Caption         =   "查名單"
      Height          =   360
      Index           =   0
      Left            =   7884
      TabIndex        =   18
      Top             =   1416
      Width           =   900
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   6
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   26
      Text            =   "0"
      Top             =   -24
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   5
      Left            =   5160
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "0"
      Top             =   672
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
      Left            =   7800
      TabIndex        =   24
      Top             =   12
      Width           =   880
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "明細"
      Height          =   360
      Left            =   8892
      TabIndex        =   19
      Top             =   1416
      Width           =   900
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8760
      TabIndex        =   25
      Top             =   12
      Width           =   880
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm090127_New.frx":0000
      Height          =   3552
      Left            =   60
      TabIndex        =   33
      Top             =   2160
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   6265
      _Version        =   393216
      Cols            =   17
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|撤|送|本所案號|委查人|文字檢索|圖形檢索|客戶名稱|智權備註|文字結果|圖形結果|類別組群|查名人|查名單|委查日期|送出期限|查覆期限"
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
      _Band(0).Cols   =   17
   End
   Begin MSForms.TextBox textMemo 
      Height          =   300
      Left            =   1080
      TabIndex        =   14
      Top             =   1608
      Width           =   2500
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "4410;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "智權備註："
      Height          =   240
      Left            =   120
      TabIndex        =   47
      Top             =   1608
      Width           =   972
   End
   Begin MSForms.TextBox textCName 
      Height          =   300
      Left            =   1080
      TabIndex        =   8
      Top             =   990
      Width           =   2500
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "4410;529"
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
      Height          =   288
      Left            =   5160
      TabIndex        =   4
      Top             =   372
      Width           =   1632
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
      Height          =   204
      Left            =   4128
      TabIndex        =   45
      Top             =   1944
      Width           =   5148
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
      Height          =   228
      Left            =   168
      TabIndex        =   42
      Top             =   5784
      Width           =   2892
   End
   Begin VB.Label lblState 
      AutoSize        =   -1  'True
      Caption         =   "lblState"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   7920
      TabIndex        =   41
      Top             =   456
      Width           =   516
   End
   Begin VB.Label Label6 
      Caption         =   "狀態："
      ForeColor       =   &H00C00000&
      Height          =   252
      Left            =   7320
      TabIndex        =   40
      Top             =   456
      Width           =   612
   End
   Begin VB.Label Label11 
      Caption         =   "(請以"",""或""."" 區隔)"
      Height          =   240
      Left            =   7800
      TabIndex        =   39
      Top             =   1008
      Width           =   1932
   End
   Begin VB.Label Label10 
      Caption         =   "查名單號："
      Height          =   240
      Left            =   4200
      TabIndex        =   38
      Top             =   1008
      Width           =   972
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
      Height          =   252
      Index           =   1
      Left            =   8760
      TabIndex        =   36
      Top             =   576
      Visible         =   0   'False
      Width           =   528
      VariousPropertyBits=   -1400879077
      MaxLength       =   50
      Size            =   "926;450"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line2 
      X1              =   2040
      X2              =   2160
      Y1              =   883
      Y2              =   883
   End
   Begin VB.Label Label7 
      Caption         =   "查覆期限："
      Height          =   240
      Left            =   120
      TabIndex        =   32
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
      TabIndex        =   31
      Top             =   422
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "查名人："
      Height          =   240
      Left            =   4200
      TabIndex        =   30
      Top             =   396
      Width           =   828
   End
   Begin VB.Label Label16 
      Caption         =   "雙擊選取時，開啟查覆明細"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   96
      TabIndex        =   29
      Top             =   1944
      Width           =   2292
   End
   Begin VB.Label Label1 
      Caption         =   "委查人："
      Height          =   240
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   825
   End
   Begin VB.Label Label2 
      Caption         =   "查名類別：              0.全部 1.文字 2.圖形 3. 文字及圖形"
      Height          =   240
      Left            =   4200
      TabIndex        =   27
      Top             =   696
      Width           =   4452
   End
End
Attribute VB_Name = "frm090127_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2024/09/24  Form2.0 ; Combo2、textCName、textMemo
Option Explicit
'設定可使用表單
Private nfrm090126_New As Form
Private nfrm090128_New As Form

Public contCusName As String '從接洽單傳申請人(中文名稱) ->客戶名稱
Public stKeyUser As String '從接洽單傳委查人
Dim m_TMQApp As String '收文已勾選的單號(接洽單Form_Unload使用)
Dim m_NoList As String '勾選的查名單編號
Dim pTMQList As String '從接洽單傳入原本勾選的委查單(TQD02)

Dim iStiu As Integer  '狀態 :0查詢 1編輯
Dim R_type As String '使用者角色

Dim bolCont As Boolean '從聯絡單來選擇收文
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim intA As Integer
Dim dblPrevRow As Double

Dim m_AttachPath As String
Dim m_PrevForm As Form

Dim strManUser As String '查名中心主管：可看到所有查名人員
Dim strManUser2 As String '內商查名覆核人員
Dim bolCaseRead As Boolean '從歷程過來看結果
Dim mCaseNo(1 To 4) As String, mStiu  As String '傳入本所案號>>增加從歷程過來看結果的人員：只限該案件的資料
Dim stIdList As String '創新業務組成員可操作清單(WXX部門的人可以操作自已部門所有人的資料,例W10所有人都可操作W1001，W20所有人都可操作W2001。

Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim strGrpTmp1 As String, strGrpTmp2 As String

Dim colTMA01 As Integer, colTMA13 As Integer, colTMA08 As Integer, colTMA12 As Integer 'TMA01查名單、TMA13是否撤回、TMA08委查人、TMA12送出期限
Dim colCaseNo As Integer, colCp09 As Integer, colCP57 As Integer, colCP14 As Integer  'colCaseNo=TMA35已收文本所案號(只有查名單輸入可以寫入)、colCp09=TMA34新申請案收文號、取消收文、商申承辦人
Dim colTMA10 As Integer, colTMA11 As Integer, colTMA14 As Integer, colChkType As Integer 'TAM10查名人員、TMA11查覆期限、TMA14查覆日期=查覆完畢、ChkType覆核流程中
Dim colTMA66D As Integer, colAnsWord As Integer, colAnsPic As Integer  'TMA66覆核日期、文字檢索結果、圖形檢索結果
Dim colTMA05D As Integer, colTMA07D As Integer 'TMA05啟動(網中)=送出日期、TMA07網中系統回寫日期
Dim colState As Integer '送件狀態：◎已通知　●已送件

Public Sub SetParent(ByRef fm As Form, Optional ByVal fCaseNo As String = "", Optional ByVal fStiu As String = "", Optional ByVal fPlist As String = "")
   Set m_PrevForm = fm
   If m_PrevForm.Name = "frm090801" Or m_PrevForm.Name = "frm090801_New" Then bolCont = True
   '傳入本所案號
   If fCaseNo <> "" Then
       Call ChgCaseNo(fCaseNo, mCaseNo)
   End If
   mStiu = fStiu
   '增加從歷程過來看結果的人員：只限該案件的資料
   bolCaseRead = False
   If mCaseNo(1) <> "" And TypeName(m_PrevForm) <> "Nothing" Then
       If m_PrevForm.Name = "frm090202_2" Then
           bolCaseRead = True
       End If
   End If
   pTMQList = fPlist
End Sub

Public Function IsRolePlay(ByRef defRole As String, Optional ShowMsg As Boolean = True) As Boolean
   Dim tmpY As Integer
   
   IsRolePlay = True
    Select Case defRole
        Case "待查"
              strExc(1) = "select tmqm01 from tmqmember where tmqm01='" & strUserNum & "' "
              tmpY = 1
              Set RsTemp = ClsLawReadRstMsg(tmpY, strExc(1))
              strManUser = Pub_GetSpecMan("內商查名主管")
              '林嘉雯請假時職代處理
              strExc(1) = GetDutyList(strManUser)
              If strExc(1) <> "" Then strManUser = strManUser & ";" & strExc(1)
              
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
Dim tmpList As String, tmpBol As Boolean, strCP09 As String
   
   If TypeName(nfrm090128_New) <> "Nothing" And cmdDetail.Visible = True Then
        For intA = 1 To GRD1.Rows - 1
           If GRD1.TextMatrix(intA, 0) = "V" Then
              tmpList = tmpList & Trim(GRD1.TextMatrix(intA, colTMA01)) & ","
           End If
        Next intA

        If tmpList <> "" Then
           tmpList = Mid(tmpList, 1, Len(tmpList) - 1)
           '從接洽單->查覆區->明細
           If TypeName(m_PrevForm) <> "Nothing" Then
              tmpBol = True
           End If
           If Trim(GRD1.TextMatrix(1, colCaseNo)) = Trim(txtField(9)) & "-" & Trim(txtField(10)) & IIf(Val(Trim(txtField(11) & txtField(12))) = 0, "", "-" & Left(Trim(txtField(11)) & "0", 1) & Left(Trim(txtField(12)) & "00", 2)) Then
             strCP09 = Trim(GRD1.TextMatrix(1, colCp09))
           End If
           nfrm090128_New.SetParent Me, tmpBol, tmpList, 0, R_type, iStiu, strCP09
           nfrm090128_New.Show
           If nfrm090128_New.QueryData = True Then
             Me.Hide
           Else
             Unload nfrm090128_New
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
    If TypeName(nfrm090126_New) <> "Nothing" Then
         For intA = 1 To GRD1.Rows - 1
            If GRD1.TextMatrix(intA, 0) = "V" Then
               nfrm090126_New.SetParent Me, Trim(GRD1.TextMatrix(intA, colTMA01))
               nfrm090126_New.Show
               Me.Hide
               Exit For
            End If
         Next intA
    End If
  ElseIf Index = 1 Then
    If TypeName(nfrm090126_New) <> "Nothing" Then
       nfrm090126_New.SetParent Me
       nfrm090126_New.Show
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
    
   m_TMQApp = "" '清空判斷
   If QueryData = False Then ShowNoData
End Sub

'****重新分派查名人員****
Private Sub GetNewTMA10()
Dim strQ1 As String, strQ2 As String
Dim strQuser As String
Dim RsQ As New ADODB.Recordset
Dim intQ As Integer
Dim bolUpd As Boolean
Dim iCnt As Integer
Dim bolUpTMA11 As Boolean
Dim Inputtm As String, InputWDay As String
Dim chkAllStatus As String '內商查名單分單狀態：若查名中心聯絡開始不分單將狀態改為N，恢復分單將狀態改為Y

   strQ1 = ""
   chkAllStatus = Pub_GetSpecMan("內商查名單分單狀態")

   If InStr(Combo1.Text, "全部") > 0 And Trim(txtField(8)) = "" Then
      If MsgBox("未指定重新分派的查名人，要改抓目前所有未分派的查名單？" & vbCrLf & "繼續作業按是，要重選查名人請按否。", vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   Else
      strQ1 = strQ1 & " and TMA10='" & Trim(Left("" & Combo1.Text, 6)) & "'"
      If txtField(1) & txtField(2) & txtField(8) = "" Then
         MsgBox "請輸入委查期間或委查單號!"
         Exit Sub
      End If
   End If
   If txtField(1) <> "" Then strQ1 = strQ1 & " and TMA09>=" & TransDate(Trim(txtField(1)), 2)
   If txtField(2) <> "" Then strQ1 = strQ1 & " and TMA09<=" & TransDate(Trim(txtField(2)), 2)
   If txtField(8) <> "" Then strQ1 = strQ1 & " and TMA01 in (" & GetAddStr(Replace(txtField(8).Text, ".", ",")) & ")"
   If txtField(0) <> "" Then strQ1 = strQ1 & " and TMA08='" & Trim(txtField(0)) & "' "
   '客戶名稱
   If textCName <> "" Then
      strExc(1) = Replace(Trim(textCName), " ", "%")
      strExc(1) = Replace(strExc(1), "%%", "%")
      strQ1 = strQ1 & " and upper(TMA18) like '%" & UCase(strExc(1)) & "%'"
   End If

   '抓目前未分派的查名單
   If InStr(Combo1.Text, "全部") > 0 Then
        strQ1 = Replace(UCase(strQ1), "TMA09", "to_char(tma04,'yyyymmdd')") '可能沒有日期
        If MsgBox("是否要排除今天下午6點送出的查名單?", vbInformation + vbYesNo) = vbYes Then
           strQ1 = strQ1 & " AND to_char(tma09,'yyyymmdd')<=" & strSrvDate(1) & " AND substr(to_char(tma04,'HH24MISS'),1,4) < 1800 "
        End If
        
        strQ2 = "select a.*,to_char(tma04,'yyyymmdd') as tma04d,to_char(tma07,'yyyymmdd') as tma07d from tmqappform a where TMA14 is null " & IIf(Trim(txtField(8)) <> "", " and TMA01 in (" & GetAddStr(Replace(txtField(8).Text, ".", ",")) & ") ", " and TMA10 is null " & strQ1) & " order by TMA01"
        intQ = 0
        Set RsQ = ClsLawReadRstMsg(intQ, strQ2)
        If intQ = 1 Then
          If MsgBox("預計有 " & RsQ.RecordCount & " 筆查名單要分派，是否繼續？", vbYesNo + vbInformation + vbDefaultButton1) = vbNo Then
             Exit Sub
          End If

          '遇到有人員臨時請假,查名單已到期的狀況
          RsQ.MoveFirst
          If "" & RsQ.Fields("TMA04D") <> "" And "" & RsQ.Fields("TMA04D") <> strSrvDate(1) Then
             If MsgBox("是否重新計算期限日期？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                bolUpTMA11 = True
                Inputtm = Left(Format(ServerTime, "000000"), 4)
             End If
          End If

          cnnConnection.BeginTrans
            '更新統計人員狀態
            strQ2 = "select tmqm02,nvl(tmqm03,'N') tmqm03,count(*) r1,count(tmasr11) r2 from tmqmember,TmqAppSumR " & _
                    "where tmqm01<>tmqm02 and tmqm01=tmasr01(+)  group by tmqm02,nvl(tmqm03,'N') "
            intQ = 1
            Set RsTemp = ClsLawReadRstMsg(intQ, strQ2)
            If intQ = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  If ("" & RsTemp.Fields("tmqm03") = "Y" And Val("" & RsTemp.Fields("r2")) > 0) Or _
                     ("" & RsTemp.Fields("tmqm03") = "N" And Val("" & RsTemp.Fields("r1")) = Val("" & RsTemp.Fields("r2"))) Then
                     strQ2 = "update TmqAppSumR set tmasr11='N' where tmasr01=" & CNULL(Trim("" & RsTemp.Fields("tmqm02")))
                     cnnConnection.Execute strQ2
                  End If
                  RsTemp.MoveNext
               Loop
            End If
            '有可能是整批變更資料後的重新分發，先更新統計單量
            Call PUB_TMAtoTake("2", "", "", "0", False)
            Call PUB_TMAtoTake("2", "", "", "1", False)
            
            Do While Not RsQ.EOF
                '視情況重新計算期限
                strQ1 = ""
                If bolUpTMA11 = True Then
                   '參考frm090126_New：送出期限TMA12/查覆期限TMA11：團體標章和證明標章仍舊由查名人負責，設定為19221111
                   If "" & RsQ.Fields("tma20") <> "" Then
                     InputWDay = PUB_GetNewTMADate(IIf("" & RsQ.Fields("tma25") = "1", "4", "5"), strSrvDate(1), Inputtm, chkAllStatus)
                     strQ1 = ", TMA12=19221111, TMA11=" & InputWDay
                   Else
                     '只能預設送出期限，查覆期限>>1.TMA07回寫日期Trigger觸發計算 2.遇見不發單時間改成批次
                     InputWDay = PUB_GetNewTMADate("" & RsQ.Fields("tma25"), strSrvDate(1), Inputtm, chkAllStatus)
                     If "" & RsQ.Fields("tma07d") = "" Then   '網中未回寫，只能預設送出期限
                        strQ1 = ", TMA12=" & InputWDay & " "
                     Else    '網中已回寫，重新計算查覆期限
                        strQ1 = ", TMA11=" & InputWDay & " "
                     End If
                   End If
                End If

                strQuser = PUB_GetTMAUserPos("" & RsQ.Fields("tma25"))
                If strQuser <> "" Then
                   strQ2 = "Update TMQAppForm set TMA10=" & CNULL(strQuser) & ",TMA09=to_date(" & CNULL(strSrvDate(1)) & "||' '||" & CNULL(Inputtm) & "||'00','yyyymmdd HH24MISS') " & strQ1 & " where TMA01=" & CNULL(RsQ.Fields("TMA01"))
                   cnnConnection.Execute strQ2
                   '委查日期含前2個工作天到當天,重新計算
                   '分發日期改成系統日=當天
                    Call PUB_TMAtoTake("2", strQuser, "", "1", False)
                    iCnt = iCnt + 1
                End If
                RsQ.MoveNext
            Loop
          cnnConnection.CommitTrans
          MsgBox "已分派完 " & iCnt & " 筆查名單!", vbInformation
        End If
   Else
        strQ2 = "select tmasr01,tmasr11 from TmqAppSumR where tmasr01=" & CNULL(Trim(Left("" & Combo1.Text, 6)))
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
           strQ2 = "select a.*,to_char(tma04,'yyyymmdd') as tma04d,to_char(tma07,'yyyymmdd') as tma07d from tmqappform where TMA14 is null" & strQ1 & " order by TMA01"
           intQ = 1
           Set RsQ = ClsLawReadRstMsg(intQ, strQ2)
           If intQ = 0 Then
              MsgBox "無委查單可分派!"
              Exit Sub
           Else
              If MsgBox("預計有 " & RsQ.RecordCount & " 筆查名單要分派，是否繼續？", vbYesNo + vbInformation + vbDefaultButton1) = vbNo Then
                  Exit Sub
              End If
              RsQ.MoveFirst
              '遇到有人員臨時請假,查名單已到期的狀況
              If "" & RsQ.Fields("TMA04D") <> "" And "" & RsQ.Fields("TMA04D") <> strSrvDate(1) Then
                 If MsgBox("是否重新計算期限日期？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                    bolUpTMA11 = True
                 End If
              End If

              cnnConnection.BeginTrans
              '更新查名人狀態
              If bolUpd Then
                strQ2 = "update TmqAppSumR set tmasr11='N' where tmasr01=" & CNULL(Trim(Left("" & Combo1.Text, 6)))
                cnnConnection.Execute strQ2
                '組群分派，一人請假全部不分
                strQ2 = "update TmqAppSumR set tmasr11='N' where tmasr01 in (select tmqm02 from tmqmember where tmqm01=" & CNULL(Trim(Left("" & Combo1.Text, 6))) & " and tmqm02 <> " & CNULL(Trim(Left("" & Combo1.Text, 6))) & " and tmqm03='Y') "
                cnnConnection.Execute strQ2
                strQ2 = "update TmqAppSumR set tmasr11='N' where tmasr01 in (select tmqm01 from tmqmember where tmqm02 in (select tmqm02 from tmqmember where tmqm01=" & CNULL(Trim(Left("" & Combo1.Text, 6))) & " and tmqm02 <> " & CNULL(Trim(Left("" & Combo1.Text, 6))) & " and tmqm03='Y')) "
                cnnConnection.Execute strQ2
              End If
              '先將分發日期拿掉,更新統計單量; ex.10張委查單(文字1,圖形)分成兩次重新分發,其中內商程序-79041被分到8張
              strQ2 = "update TMQAppForm set TMA09=null,TMA10=null where TMA14 is null " & strQ1
              cnnConnection.Execute strQ2, intI
              '整批變更資料後的重新分發 , 先更新統計單量
              Call PUB_TMAtoTake("2", "", "", "0", False)
              Call PUB_TMAtoTake("2", "", "", "1", False)

              With RsQ
                Do While Not RsQ.EOF
                    'A視情況重新計算期限
                    strQ1 = ""
                    If bolUpTMA11 = True Then
                       '參考frm090126_New：送出期限TMA12/查覆期限TMA11：團體標章和證明標章仍舊由查名人負責，設定為19221111
                       If "" & RsQ.Fields("tma20") <> "" Then
                          InputWDay = PUB_GetNewTMADate(IIf("" & RsQ.Fields("tma25") = "1", "4", "5"), strSrvDate(1), Inputtm, chkAllStatus)
                          strQ1 = ", TMA12=19221111, TMA11=" & InputWDay
                       Else
                          '只能預設送出期限，查覆期限>>1.TMA07回寫日期Trigger觸發計算 2.遇見不發單時間改成批次
                          InputWDay = PUB_GetNewTMADate("" & RsQ.Fields("tma25"), strSrvDate(1), Inputtm, chkAllStatus)
                          If "" & RsQ.Fields("tma07d") = "" Then   '網中未回寫，只能預設送出期限
                             strQ1 = ", TMA12=" & InputWDay & " "
                          Else    '網中已回寫，重新計算查覆期限
                             strQ1 = ", TMA11=" & InputWDay & " "
                          End If
                       End If
                    End If
                    strQuser = PUB_GetTMAUserPos("" & RsQ.Fields("tma25"))
                    If strQuser <> "" Then
                       strQ2 = "Update TMQAppForm set TMA10=" & CNULL(strQuser) & ",TMA09=to_date(" & CNULL(strSrvDate(1)) & "||' '||" & CNULL(Inputtm) & "||'00','yyyymmdd HH24MISS') " & strQ1 & " where TMA01=" & CNULL(.Fields("TMA01"))
                       cnnConnection.Execute strQ2
                       '委查日期含前2個工作天到當天,重新計算
                       '分發日期改成系統日=當天
                       Call PUB_TMAtoTake("2", strQuser, "", "1", False)
                       Call PUB_TMAtoTake("2", .Fields("TMA10"), "", "1", False) '原查名人員
                    End If
                  .MoveNext
                Loop
              End With
              cnnConnection.CommitTrans
              MsgBox "已分派完 " & iCnt & " 筆查名單!", vbInformation
           End If
        End If
   End If
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
Dim strQuery As String, intQ As Integer
Dim strCon As String, strCon2 As String

   If txtField(9) <> "" And Len(Trim(txtField(10))) <> 6 Then
      MsgBox "請輸入完整的本所案號!", vbOKOnly, "輸入錯誤"
      txtField(10).SetFocus
      QueryData = False
      Exit Function
   End If

   Call SetGrd(True)

   m_blnColOrderAsc = True
   QueryData = True
   
JumpReSearch:

   strCon = "": strCon2 = ""
   
   '委查人
   If InStr(Me.Caption, "查覆區") > 0 Then
      If m_TMQApp = "" Then
         '查覆區預設不列出已收文的委查單
         '只要輸入委查單號,就不限制狀態; 排除從歷程過來看結果+bolCaseRead
         If Trim(txtField(8)) = "" And bolCaseRead = False Then
            strCon = strCon & " and c1.cp27 is null"
         End If
      End If
      If InStr(Combo2.Text, "全部") > 0 And (Pub_StrUserSt03 = "M51" Or bolCaseRead = True) Then
      '創新業務組成員可操作清單
      ElseIf InStr(stIdList, "W") > 0 And Left(strGrpTmp1, 1) = "W" Then
         strCon = strCon & " and TMA08 in (" & stIdList & ")"
      Else
         strCon = strCon & " and TMA08='" & Trim(Left("" & Combo2.Text, 6)) & "'"
      End If
   ElseIf txtField(0) <> "" Then
      strCon = strCon & " and TMA08='" & Trim(txtField(0)) & "'"
   End If
   
   'Modified by Lydia 2025/04/15 追蹤網中執行舊查名單的結果(先只看成功的AND TMA43='Y')
   strCon = strCon & " and TO_CHAR(TMA04,'YYYYMMDD')>='20240601' " '排除1120904-1120928期間資料匯入
   If R_type = "U" Then
      strCon = strCon & " and TMA02='1' " '查名人員：內部輸入的查名檢查
   End If
   'Mark by Lydia 2025/04/18 '2025/04/15 追蹤網中執行舊查名單的結果(先只看成功的AND TMA43='Y')
   'strCon = strCon & " and ((to_char(tma07,'yyyymmdd')>=20250101 AND TMA43='Y') or (TMA02='2' AND TMA43='Y' AND TMA39||TMA41 IS NOT NULL) or (TO_CHAR(TMA04,'YYYYMMDD')>='20240601' "
   'If R_type = "U" Then
   '   strCon = strCon & " and TMA02='1')) " '查名人員：內部輸入的查名檢查
   'Else
   '   strCon = strCon & ")) "
   'End If
   'end 2025/04/15
   
   '委查期間
   If txtField(1) <> "" Then strCon = strCon & " and TO_CHAR(TMA04,'YYYYMMDD')>=" & TransDate(Trim(txtField(1)), 2)
   If txtField(2) <> "" Then strCon = strCon & " and TO_CHAR(TMA04,'YYYYMMDD')<=" & TransDate(Trim(txtField(2)), 2)
   '查覆期限期間
   If txtField(3) <> "" Then strCon = strCon & " and TMA11>=" & TransDate(Trim(txtField(3)), 2)
   If txtField(4) <> "" Then strCon = strCon & " and TMA11<=" & TransDate(Trim(txtField(4)), 2)
   '類別
   If txtField(5) <> "" And txtField(5) <> "0" Then
      strCon = strCon & " and TMA25='" & Trim(txtField(5)) & "'"
   End If
   
   '狀態
   lblState = ""
   '只要輸入委查單號,就不限制狀態
   If txtField(6) <> "" And Trim(txtField(8)) = "" Then
      Select Case txtField(6)
         Case "0" '全部=未發文
            lblState = "處理中 + 已完成"
         Case "1" '處理中
            lblState = "處理中"
            strCon = strCon & " and TMA14 is null"
         Case "2" '已完成
            lblState = "已完成"
            strCon = strCon & " and TMA14>0 and TMA13 is null"
         Case "3"
            '將已發文(送件)案件特別區隔
            lblState = "已發文"
            If strCon <> "" And InStr(strCon, "c1.cp27 is null") > 0 Then
               strCon = Replace(strCon, "c1.cp27 is null", "c1.cp27 is not null ")
            Else
               strCon = strCon & " and c1.cp27 is not null "
            End If
      End Select
      '將已發文(送件)案件特別區隔+排除從歷程過來看結果bolCaseRead
      If txtField(6) <> "3" And InStr(strCon, "c1.cp27") = 0 And bolCaseRead = False Then strCon = strCon & " and c1.cp27 is null"
   End If
  
   '客戶名稱
   If textCName <> "" Then
      strExc(1) = Replace(Trim(textCName), " ", "%")
      strExc(1) = Replace(strExc(1), "%%", "%")
      strCon = strCon & " and upper(TMA18) like '%" & UCase(strExc(1)) & "%'"
   '從接洽單傳申請人(中文名稱) ->客戶名稱
   ElseIf contCusName <> "" And bolCont = True Then
      '從字首4字開始比對到字首2字
      If Len(contCusName) > 4 Then
         contCusName = Mid(contCusName, 1, 4)
      End If
      strCon = strCon & " and TMA18 like '" & contCusName & "%' "
   End If
   '查名人
   If InStr(Combo1.Text, "全部") = 0 Then
      strCon = strCon & " and TMA10='" & Trim(Left(Combo1.Text, 5)) & "'"
   ElseIf R_type = "U" And strUserNum <> Trim(Left(Combo1.Text, 5)) Then
           If Pub_StrUserSt03 <> "M51" Then iStiu = 0 '非分派到的查名人，不可修改。電腦中心除外。'1234
   End If
   '委查單號
   If txtField(8) <> "" Then
      txtField(8).Text = Replace(txtField(8).Text, ".", ",")
      strCon = strCon & " and TMA01 in (" & CNULL(Replace(txtField(8).Text, ",", "','")) & ")"
   End If
   strCon2 = strCon
   '本所案號
   If txtField(9) <> "" And txtField(10) <> "" Then
      '個案查詢的本所案號改抓輸入的案號
      strExc(1) = " and c1.cp01='" & Trim(txtField(9)) & "' and c1.cp02='" & Trim(txtField(10)) & "' and c1.cp03='" & Left(Trim(txtField(11)) & "0", 1) & "' and c1.cp04='" & Left(Trim(txtField(12)) & "00", 2) & "'"
      If txtField(9) = "T" Then
         strExc(1) = strExc(1) & " and instr('" & TMQ_T案 & "', c1.cp10) > 0 "
      ElseIf txtField(9) = "TS" Then
         strExc(1) = strExc(1) & " and instr('" & TMQ_TS案 & "', c1.cp10) > 0 "
      End If
      strCon = strCon & " AND TMA01=TQC03(+) AND TQC02=C1.CP09(+) " & strExc(1) & " and c1.cp57 is null"
   Else
      strCon = strCon & " AND TMA01=TQC03(+) AND TMA34=TQC02(+) AND TMA34=c1.CP09(+) "
   End If

   If R_type = "M" Then
      '判斷未輸入本所案號或委查單,限近似本所案
      If txtField(8) & txtField(9) & txtField(10) = "" Then
         '預設不顯示撤回=N; TMA16=查覆結果是否近似, TMA19=標章查覆結果
         strCon = strCon & " and nvl(TMA13,'N') ='N' and TMA66 IS NULL and (NVL(TMA16,'N')='Y' or (nvl(TMA19,'9') in ('" & TMQ_近似1 & "','" & TMQ_近似2 & "'))) "
      End If
   End If

   Screen.MousePointer = vbHourglass

   '共通語法:
   'ChkType>>覆核流程中(TMA16=查覆結果是否近似, TMA19=標章查覆結果)
     'Added by Lydia 2021/02/17 優化顯示需進行覆核流程之查名單 , 以利委查同仁瞭解與本所近似查名單之狀態
     '1.查名結果為相同△或近似△時，該查名單之結果以黃色顯示，即表示進行覆核中=>ChkType=A
     '2.覆核結果為相同△或近似△時，該查名單之結果以紅色顯示，即表示已覆核，仍與本所近似=>ChkType=B 'Memo by Lydia 2024/10/11 TMA67提出修改：Y=是，需進行協商流程TMA69、N=否（需確認客戶關係）、A=已排除近似
     'ChkType=N協商流程結果(TMA69=1~3) 'Memo by Lydia 2024/10/11 TMA69提出修改：1.經上級核可代理、2.經上級核可先提申再補同
   'Modified by Lydia 2025/04/15 增加判斷TMSEARCH
   strQuery = "SELECT TMA13 AS 是否撤回,DECODE(C1.CP27,NULL,DECODE(TQC07,NULL,'','◎'),'●') 送件狀態, " & _
              "DECODE(C1.CP01,NULL,'',DECODE(C1.CP03||C1.CP04,'000',C1.CP01||'-'||C1.CP02,C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04)) 本所案號, " & _
              "NVL(TMA08,TMA03) AS TMA08 ,DECODE(TMA08,NULL,'(網站)',S1.ST02) AS TMA08N,TMA26 AS 文字檢索,DECODE(TMA27,NULL,NULL,'(圖形檢索)') AS 圖形檢索,TMA18 AS 客戶名稱,TMA33 AS 智權備註, " & _
              "DECODE(TMA25,'2',NULL,DECODE(TMA13,'Y','不查',DECODE(TMA19," & PUB_GetTMQans("3", True) & ", TMA39))) AS 文字檢索結果,DECODE(TMA25,'1',NULL,DECODE(TMA13,'Y','不查',DECODE(TMA19," & PUB_GetTMQans("3", True) & ",TMA41))) AS 圖形檢索結果, " & _
              PUB_GetTMAforClass & " AS 類別組群,DECODE(TMA03||TMA10,'TMSEARCH',TMA10) AS TMA10 ,DECODE(TMA03||TMA10,'TMSEARCH','(網站)',S2.ST02) AS TMA10N,TMA01,(TO_CHAR(TMA04,'YYYYMMDD')-19110000) AS 委查日期,(TMA12-19110000) AS 送出期限,(TMA11-19110000) AS 查覆期限,(TMA14-19110000) AS 查覆日期,DECODE(TMA66,NULL,NULL,(TO_CHAR(TMA66,'YYYYMMDD')-19110000)) AS 覆核日期,(TQC07-19110000) AS 通知送件日, " & _
              "DECODE(TMA69,NULL,DECODE(TMA67,'Y','B',DECODE(TMA16,'Y','A',DECODE(TMA19,'" & TMQ_近似1 & "','A','" & TMQ_近似2 & "','A',NULL))),'B')  AS CHKTYPE,C1.CP09,C1.CP57,C1.CP14, " & _
              "DECODE(TMA05,NULL,NULL,(TO_CHAR(TMA05,'YYYYMMDD')-19110000)) AS TMA05D, DECODE(TMA07,NULL,NULL,(TO_CHAR(TMA07,'YYYYMMDD')-19110000)) AS TMA07D " & _
              "FROM TMQAPPFORM, STAFF S1, STAFF S2, TMQCASEMAP, CASEPROGRESS C1 " & _
              "WHERE TMA08=S1.ST01(+) AND TMA10=S2.ST01(+) "
   strSql = strQuery & strCon
   '待查+查詢非已發文+輸入本所案號>> 因為畫面條件不同，所以另外列出分給查名人的查名單
   If InStr(Me.Caption, "待查區") > 0 And txtField(6) <> "3" And Trim(txtField(9) & txtField(10)) <> "" Then
      strSql = strSql & " Union " & strQuery & " AND TMA01=TQC03(+) AND TMA34=TQC02(+) AND TMA34=c1.CP09(+) AND TMA14 IS NULL " & strCon2
   End If
        
   If R_type = "Q" Then '查覆區：委查人   --->'Added by Lydia 2017/11/06 查覆區欄位順序:委查單號,期限日期,委查日期 (因為智權人員常會打電話催查名人員BY嘉雯)
      strExc(0) = "送出期限,查覆期限,委查日期"
   Else
      strExc(0) = "委查日期,送出期限,查覆期限"
   End If
   strSql = "SELECT ' ' V,是否撤回,送件狀態,本所案號,TMA08,TMA08N AS 委查人,文字檢索,圖形檢索,客戶名稱,智權備註,文字檢索結果,圖形檢索結果" & _
            ",類別組群, TMA10,TMA10N AS 查名人,TMA01 AS 查名單," & strExc(0) & " ,查覆日期,覆核日期,通知送件日,CHKTYPE,CP09,CP57,CP14,TMA05D,TMA07D FROM (" & strSql & ") "
            
   '待查區依期限日期
   If InStr(Me.Caption, "待查區") > 0 Then
      '未完成的置頂
      strSql = strSql & " order by 查覆日期 desc,查覆期限 asc,查名單 desc"
   Else
      strSql = strSql & " order by 查名單 desc"
   End If
   
   GRD1.FixedCols = 0
   intQ = 1
   Set rsTmp = ClsLawReadRstMsg(intQ, strSql)
   If intQ = 1 Then
      If contCusName <> "" And bolCont = True Then
         textCName.Text = contCusName
         contCusName = ""
      End If
      Set GRD1.Recordset = rsTmp
      Call SetGrd
      GRD1.FixedCols = 10
   Else
      QueryData = False
      Screen.MousePointer = vbDefault
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
      Exit Function
   End If

   Set rsTmp = Nothing
   Screen.MousePointer = vbDefault

End Function

Private Sub cmdState_Click(Index As Integer)

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

   
   If Index <> 4 Then
      txtField(6) = Index
      Call cmdQuery_Click
   Else '重新分查名人
      Call GetNewTMA10
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
   
   '改用接洽單的智權人員
   If stKeyUser <> "" Then
      pUser = stKeyUser
   Else
      pUser = strUserNum
   End If
   
   Combo2.Clear
   '增加從歷程過來看結果的人員：只設定清單=全部
   If bolCaseRead = True Then
       Combo2.AddItem "      " & "全部"
       Combo2.ListIndex = 0
       Exit Sub
   End If
   '設定屬智權人員作業的下拉選單(共用模組)
   Call PUB_SetCombo1Sales(Me.Combo2, pUser)

   If Pub_StrUserSt03 = "M51" And strUserNum = pUser Then
      Combo2.AddItem "      " & "全部"
   End If
   Combo2.ListIndex = 0
End Sub

Private Sub SetCombo1()
Dim mSQL As String
Dim cInX As Integer

   Combo1.Clear
   
   If R_type = "U" And (strManUser <> "" And InStr(strManUser, strUserNum) = 0 And Pub_StrUserSt03 <> "M51") Then
      mSQL = " and tmqm01='" & strUserNum & "' "
   End If
   
   strSql = " select tmqm01,st02 from tmqmember,staff where tmqm01=st01(+) and st04='1' " & mSQL & " order by 1 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      '查名人不可看其他人的委查單,若要參考查名路徑請自行詢問其他人by林經理
      If mSQL = "" Then
         Combo1.AddItem "ALL 全部"
      End If
      Do While Not RsTemp.EOF
         Combo1.AddItem Trim(RsTemp(0)) & " " & Trim(RsTemp(1))
         '預設個人的待查區
         If InStr(strManUser, strUserNum) > 0 And Trim(RsTemp(0)) = strUserNum And InStr(Me.Caption, "待查區") > 0 Then
            cInX = Combo1.ListCount - 1
         End If
         RsTemp.MoveNext
      Loop
   End If
   If cInX > 0 Then
      Combo1.ListIndex = cInX
   Else
      Combo1.ListIndex = 0
   End If
End Sub

Private Sub Form_Load()

   '先隱藏輸入功能
   'If Pub_StrUserSt03 <> "M51" Then '1234
      Frame1.Visible = False
      Me.Height = 6500  '1234
   'End If '1234
   
   MoveFormToCenter Me
  
    m_AttachPath = App.path & "\" & strUserNum
    If Dir(m_AttachPath, vbDirectory) = "" Then
       MkDir m_AttachPath
    End If
   Call PUB_GetTMQans("1", True) '求近似本所案
   
   txtField(5).Text = "0" '預設類別:全部
   txtField(6).Text = "0" '預設狀況:未發文
   
   Select Case R_type
       Case "U", "M", "A" '待查(查名人),覆核
            iStiu = 1
            Select Case R_type
                 Case "U"
                    txtField(6).Text = "1" ' 預設處理中,不限日期
                     Me.Caption = "查名/待查區(網中)"
                 Case "M"
                    txtField(3).Text = strSrvDate(2)
                    txtField(4).Text = TransDate(CompWorkDay(5, strSrvDate(1), 0), 1)
                    Me.Caption = "查名/覆核區(網中)"
                 Case "A"
                    txtField(6).Text = "1" ' 預設處理中,不限日期
                    Me.Caption = "查名單維護(網中)"
                    cmdState(4).Visible = True
            End Select
            cmdTo.Visible = False
            cmdSendMail.Visible = False
            SetCombo1
            Combo2.Visible = False: txtField(0).Visible = True: lblSname.Visible = True
       Case "Q"    '查覆(委查人)
            iStiu = 0
            '從接洽單傳入員工代號
            If stKeyUser <> "" Then
               txtField(0).Text = stKeyUser
            Else
               txtField(0).Text = strUserNum
            End If
            txtField_Validate 0, True
            txtField(1).Text = TransDate(CompWorkDay(1, CompDate(1, -1, strSrvDate(1))), 1)
            txtField(2).Text = strSrvDate(2)

            Me.Caption = "商標查名／查覆區(網中)"
            cmdTo.Visible = True
            cmdSendMail.Visible = True
            SetCombo1
            SetCombo2
            Combo2.Visible = True: txtField(0).Visible = False: lblSname.Visible = False
            '增加從歷程過來看結果的人員：不限人員和日期
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
            
            '創新業務組成員可操作清單
            stIdList = PUB_GetSalesList(txtField(0).Text, , , , , strGrpTmp1, strGrpTmp2)
            If InStr(stIdList, "W") = 0 Or Left(strGrpTmp1, 1) <> "W" Then
                stIdList = CNULL(txtField(0).Text) '非創新業務組用切換清單的方式
            End If
            '已新增系統特殊設定「智權部可查詢同所查名單人員」。
            '請調整程式，屬於該名單內的人員可查詢同所別智權部同仁的查名單資料。---李承翰
            If UCase(TypeName(m_PrevForm)) = "NOTHING" Then
               strExc(1) = Pub_GetSpecMan("智權部可查詢同所查名單人員")
               If InStr(strExc(1) & ",", strUserNum) > 0 Then
                  strExc(0) = "select st01,st02,st15 from staff where st06='" & pub_strUserOffice & "'" & _
                     " and st15 like 'S%' and st04='1' and st01 in (select distinct(TMA08) TMA08 from TMQAppform where tma02='1') order by st15,st01"
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
   End Select
   
   '預設傳入案號為條件
   If mCaseNo(1) <> "" Then
       txtField(9) = mCaseNo(1)
       txtField(10) = mCaseNo(2)
       txtField(11) = mCaseNo(3)
       txtField(12) = mCaseNo(4)
       txtField(6) = IIf(mStiu = "", "2", mStiu) '預設查詢-已完成
       '增加從歷程過來看結果的人員：只限該案件的資料
       If bolCaseRead = True Then
            txtField(9).Enabled = False
            txtField(10).Enabled = False
            txtField(11).Enabled = False
            txtField(12).Enabled = False
            txtField(6).Text = "0" ' 不預設狀態
       End If
   End If

   '開放特殊設定權限
    If CheckLevel(strUserNum, "總經理業務工作代理人員") = True Then
       bolSpecMan = True
       strSpecCode = Pub_GetSpecMan("總經理員工編號")
   '開放專利處部份智權同仁資料給彥葶代為處理
   ElseIf CheckLevel(strUserNum, "A8") = True Then
        bolSpecMan = True
        strSpecCode = Pub_GetSpecMan("A7")
   End If
   
   QueryData
   
   mCaseNo(1) = ""
   mCaseNo(2) = ""
   mCaseNo(3) = ""
   mCaseNo(4) = ""
   
   '查名單輸入/顯示
   Set nfrm090126_New = Forms(0).GetForm("frm090126_New")
   If Not nfrm090126_New Is Nothing Then
      cmdMaster(0).Visible = True
      cmdMaster(1).Visible = True
   Else
      cmdMaster(0).Visible = False
      cmdMaster(1).Visible = False
   End If
   '查名單明細
   Set nfrm090128_New = Forms(0).GetForm("frm090128_New")
   If Not nfrm090128_New Is Nothing Then
      cmdDetail.Visible = True
   Else
      cmdDetail.Visible = False
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If TypeName(m_PrevForm) <> "Nothing" Then
      If m_PrevForm.Name = "frm090801" Or m_PrevForm.Name = "frm090801_New" Then
         Call cmdTo_Click
         If m_NoList = "" Then m_PrevForm.Show
      Else
         m_PrevForm.Show
      End If
      contCusName = ""
      Set m_PrevForm = Nothing
    End If
    
    Set nfrm090126_New = Nothing
    Set nfrm090128_New = Nothing
    
    Set frm090127_New = Nothing
End Sub

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim strTmpA As String
   Dim iRow As Integer
   
             '1 2, 3   4       5     6      7        8        9        10       11       12       13       14    15     16                            17       18       19                                           20       21       22         23      24   25   26   27     28
   strTmpA = "V,撤,送,本所案號,TMA08,委查人,文字檢索,圖形檢索,客戶名稱,智權備註,文字結果,圖形結果,類別組群,TMA10,查名人,查名單," & IIf(R_type = "Q", "送出期限,查覆期限,委查日期", "委查日期,送出期限,查覆期限") & ",查覆日期,覆核日期,通知送件日,CHKTYPE,CP09,CP57,CP14,TMA05D,TMA07D"
   arrGridHeadText = Split(strTmpA, ",")
   
   Select Case R_type
        Case "U"  '查名/待查區
           '查名人員=隱藏查名人,主管+M51=有查閱全部人員的顯示查名人員
           If InStr(strManUser, strUserNum) > 0 Or Pub_StrUserSt03 = "M51" Then
                                      '1    2     3    4   5   6   7    8    9    10   11   12   13   14 15   16   17   18   19   20   21  22 23,24,25,26,27,28
              arrGridHeadWidth = Array(200, 300, 300, 900, 0, 820, 920, 920, 920, 860, 860, 860, 920, 0, 820, 900, 800, 800, 800, 800, 800, 0, 0, 0, 0, 0, 0, 0)
           Else
                                      '1    2     3    4   5   6   7    8    9    10   11   12   13   14 15 16   17   18   19   20   21   22   23,24,25,26,27,28
              arrGridHeadWidth = Array(200, 300, 300, 900, 0, 820, 920, 920, 920, 860, 860, 860, 920, 0, 0, 900, 800, 800, 800, 800, 800, 0, 0, 0, 0, 0, 0, 0)
           End If
           Label9.Caption = "期限日期紅色:當天或過期"
        Case "M", "A"  '查名/覆核區、查名單維護：顯示通知送件日、查名人
                                    '1    2     3    4   5   6   7    8    9    10   11   12   13   14 15   16   17   18   19   20   21   22   23,24,25,26,27,28
            arrGridHeadWidth = Array(200, 300, 300, 900, 0, 820, 920, 920, 920, 860, 860, 860, 920, 0, 820, 900, 800, 800, 800, 800, 800, 860, 0, 0, 0, 0, 0, 0)
        Case "Q", ""  '查覆(委查人)：顯示通知送件日、查名人
                                    '1    2     3    4   5   6   7    8    9    10   11   12   13   14 15   16   17   18   19   20   21   22   23,24,25,26,27,28
            arrGridHeadWidth = Array(200, 300, 300, 900, 0, 0, 920, 920, 920, 860, 860, 860, 920, 0, 820, 900, 800, 800, 800, 800, 800, 860, 0, 0, 0, 0, 0, 0)
   End Select
   
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
      GRD1.Clear
      GRD1.Rows = 2
   End If
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   If colTMA01 = 0 Then
      colTMA01 = PUB_MGridGetId("查名單", GRD1)
      colTMA13 = PUB_MGridGetId("撤", GRD1)
      colTMA08 = PUB_MGridGetId("TMA08", GRD1)
      colState = PUB_MGridGetId("送", GRD1)
      colTMA10 = PUB_MGridGetId("TMA10", GRD1)
      colTMA12 = PUB_MGridGetId("送出期限", GRD1)
      colTMA11 = PUB_MGridGetId("查覆期限", GRD1)
      colTMA14 = PUB_MGridGetId("查覆日期", GRD1)
      colChkType = PUB_MGridGetId("CHKTYPE", GRD1)  '覆核流程中
      colCaseNo = PUB_MGridGetId("本所案號", GRD1)
      colCp09 = PUB_MGridGetId("CP09", GRD1)
      colCP57 = PUB_MGridGetId("CP57", GRD1)
      colCP14 = PUB_MGridGetId("CP14", GRD1)
      colTMA66D = PUB_MGridGetId("覆核日期", GRD1)
      colAnsWord = PUB_MGridGetId("文字結果", GRD1)
      colAnsPic = PUB_MGridGetId("圖形結果", GRD1)
      colTMA05D = PUB_MGridGetId("TMA05D", GRD1)
      colTMA07D = PUB_MGridGetId("TMA07D", GRD1)
   End If

   For intI = 1 To GRD1.Rows - 1
     GRD1.row = intI
     For iRow = 0 To colTMA01
       GRD1.col = iRow
       GRD1.CellBackColor = QBColor(15)
     Next iRow
     '待查區：查名期限的當天和過期設底色
     If R_type = "U" And (txtField(6) = "0" Or txtField(6) = "1") And Trim(GRD1.TextMatrix(intI, colTMA14)) = "" Then
        If Trim(GRD1.TextMatrix(intI, colTMA11)) <> "" And Trim(GRD1.TextMatrix(intI, colTMA11)) <= strSrvDate(2) Then
           GRD1.col = colTMA11
           GRD1.CellBackColor = &HC0C0FF
        End If
     End If

   'ChkType>>覆核流程中(TMA16=查覆結果是否近似, TMA19=標章查覆結果)
     'Added by Lydia 2021/02/17 優化顯示需進行覆核流程之查名單 , 以利委查同仁瞭解與本所近似查名單之狀態
     '1.查名結果為相同△或近似△時，該查名單之結果以黃色顯示，即表示進行覆核中=>ChkType=A
     '2.覆核結果為相同△或近似△時，該查名單之結果以紅色顯示，即表示已覆核，仍與本所近似=>ChkType=B 'Memo by Lydia 2024/10/11 TMA67提出修改：Y=是，需進行協商流程TMA69、N=否（需確認客戶關係）、A=已排除近似
     'ChkType=N協商流程結果(TMA69=1~3) 'Memo by Lydia 2024/10/11 TMA69提出修改：1.經上級核可代理、2.經上級核可先提申再補同
     If Trim(GRD1.TextMatrix(intI, colChkType)) <> "" Then
        '黃色=覆核中
        If Trim(GRD1.TextMatrix(intI, colChkType)) = "A" Then
           If Trim(GRD1.TextMatrix(intI, colAnsWord)) <> "" Then
              GRD1.col = colAnsWord
              GRD1.CellBackColor = &HFFFF&
           End If
           If Trim(GRD1.TextMatrix(intI, colAnsPic)) <> "" Then
              GRD1.col = colAnsPic
              GRD1.CellBackColor = &HFFFF&
           End If
        '紅色=已覆核，仍與本所近似
        ElseIf Trim(GRD1.TextMatrix(intI, colChkType)) = "B" Then
           If Trim(GRD1.TextMatrix(intI, colAnsWord)) <> "" Then
              GRD1.col = colAnsWord
              GRD1.CellBackColor = &HFF&
           End If
           If Trim(GRD1.TextMatrix(intI, colAnsPic)) <> "" Then
              GRD1.col = colAnsPic
              GRD1.CellBackColor = &HFF&
           End If
        End If
     End If

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
      For intA = 0 To GRD1.Cols - 1
         GRD1.col = intA
         '查名期限的當天和過期設底色不變; 需進行覆核流程之查名單：設底色不變
         If Not ((R_type = "U" And (txtField(6) = "0" Or txtField(6) = "1") And intA = colTMA11) _
               Or ((intA = colAnsWord Or intA = colAnsPic) And Trim(GRD1.TextMatrix(TmpRow, colChkType)) <> "")) Then
            GRD1.CellBackColor = QBColor(15)
         End If
      Next intA
   Else
      '目前資料列反白
      GRD1.col = 0
      GRD1.row = TmpRow
      dblPrevRow = GRD1.row

      If GRD1.TextMatrix(GRD1.row, colTMA01) <> "" Then
         GRD1.Text = "V"
         For intA = 0 To GRD1.Cols - 1
            GRD1.col = intA
            '查名期限的當天和過期設底色不變; 需進行覆核流程之查名單：設底色不變
            If Not ((R_type = "U" And (txtField(6) = "0" Or txtField(6) = "1") And intA = colTMA11) _
                  Or ((intA = colAnsWord Or intA = colAnsPic) And Trim(GRD1.TextMatrix(TmpRow, colChkType)) <> "")) Then
               GRD1.CellBackColor = &HFFC0C0
            End If
         Next intA
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
      If InStr("委查日期,送出期限,查覆期限,查覆日期,覆核日期,通知送件日", Me.GRD1.Text) > 0 Then
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

Private Sub textMemo_GotFocus()
    TextInverse textMemo
End Sub

Private Sub textMemo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Forms(0).PopupMenu2 textMemo
End Sub

Private Sub txtField_GotFocus(Index As Integer)
txtField(Index).SelStart = 0
txtField(Index).SelLength = Len(txtField(Index).Text)

End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
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
        Case 5  '查名類別：文字=1,圖形=2,文字及圖形=3
             If txtField(Index) <> "0" And txtField(Index) <> "1" And txtField(Index) <> "2" And txtField(Index) <> "3" Then
                MsgBox "請輸入0-3", vbOKOnly, "輸入錯誤"
                GoTo RetErr
             End If
        Case 6  '查詢State：未發文=0,處理中=1,已完成=2,已發文=3
             If txtField(Index) <> "0" And txtField(Index) <> "1" And txtField(Index) <> "2" And txtField(Index) <> "3" Then
                MsgBox "請輸入0-3", vbOKOnly, "輸入錯誤"
                GoTo RetErr
             End If
        Case 8
             If Left(txtField(Index), 1) <> "H" Then
                MsgBox "委查單號輸入錯誤", vbOKOnly, "輸入錯誤"
                GoTo RetErr
             End If
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
    End Select
Else
    If Index = 0 Then lblSname.Caption = ""
    If Index = 9 Then txtField(10).Text = "": txtField(11).Text = "": txtField(12).Text = ""
End If

Exit Sub

RetErr:
    txtField(Index).SetFocus
    Cancel = True
End Sub

Private Sub cmdTo_Click()
   
   m_TMQApp = ""
   m_NoList = ""
   
   strExc(0) = ""
   strExc(1) = Trim(Mid(Combo2.Text, 1, InStr(Combo2.Text, " ")))
   '開放特殊設定權限
   If strExc(1) <> strUserNum And cmdTo.Visible = True Then
      strExc(1) = "N"
      If InStr(stIdList, strExc(1)) > 0 Then
          strExc(1) = ""
      '代理-總經理、A7
      ElseIf bolSpecMan = True And InStr(strSpecCode, strExc(1)) > 0 Then
          strExc(1) = ""
      End If
      If strExc(1) = "N" Then
         MsgBox "無權限!!", vbCritical
         Exit Sub
      End If
   End If
   
   For intA = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(intA, 0) = "V" Then
         If GRD1.TextMatrix(intA, colTMA13) = "Y" Then
            MsgBox "查名單: " & Trim(GRD1.TextMatrix(intA, colTMA01)) & " 已撤回", vbCritical
            Exit Sub
         End If
         '判斷接洽單傳入的智權人員; 增加創新業務組成員可互相操作
         If InStr(strExc(1) & "," & stIdList, Trim(GRD1.TextMatrix(intA, colTMA08))) = 0 _
             Or (stKeyUser <> "" And stKeyUser <> strExc(1)) Then
            MsgBox "查名單: " & Trim(GRD1.TextMatrix(intA, colTMA01)) & " 不是你申請的!", vbCritical
            Exit Sub
         End If
         strExc(0) = strExc(0) & IIf(strExc(0) <> "", ",", "") & Trim(GRD1.TextMatrix(intA, colTMA01))
      End If
   Next intA

   m_NoList = strExc(0)

   If pTMQList <> "" And m_NoList = "" Then
      If MsgBox("沒有勾選查名單，是否繼續回存到接洽單？", vbExclamation + vbYesNo + vbDefaultButton2, "接洽單修改") = vbNo Then
         Exit Sub
      End If
   End If
   
   PubShowNextData

End Sub

'通知送件
Private Sub cmdSendMail_Click()
Dim tmpStr As String
Dim strAll As String
Dim strNoList As String
Dim tmpArr As Variant
Dim id1 As Integer, id2 As Integer
Dim strTmp1 As String

On Error GoTo ErrHand01

   '增加創新業務組成員可互相操作
   strTmp1 = Trim(Mid(Combo2.Text, 1, InStr(Combo2.Text, " ")))
   '開放特殊設定權限
   If strTmp1 <> strUserNum And cmdSendMail.Visible = True Then
      strExc(1) = "N"
      If InStr(stIdList, strTmp1) > 0 Then
          strExc(1) = ""
      '代理-總經理、A7
      ElseIf bolSpecMan = True And InStr(strSpecCode, strTmp1) > 0 Then
          strExc(1) = ""
      End If
      If strExc(1) = "N" Then
          MsgBox "無權限!!", vbCritical
          Exit Sub
      End If
   End If
   id1 = PUB_MGridGetId("客戶名稱", GRD1)
   id2 = PUB_MGridGetId("文字1", GRD1)
   For intA = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(intA, 0) = "V" Then
         If GRD1.TextMatrix(intA, colTMA13) = "Y" Then
            MsgBox "查名單: " & Trim(GRD1.TextMatrix(intA, colTMA01)) & " 已撤回", vbCritical
            Exit Sub
         End If
         '增加創新業務組成員可互相操作 ; 開放特殊設定權限
         If InStr(strTmp1 & "," & stIdList & IIf(bolSpecMan = True, "," & strSpecCode, ""), Trim(GRD1.TextMatrix(intA, colTMA08))) = 0 Then
            MsgBox "查名單: " & Trim(GRD1.TextMatrix(intA, colTMA01)) & " 不是你申請的!", vbCritical
            Exit Sub
         End If
         If Trim(GRD1.TextMatrix(intA, colCp09)) = "" Then
            MsgBox "查名單: " & Trim(GRD1.TextMatrix(intA, colTMA01)) & " 未收文", vbCritical
            Exit Sub
         End If
         If Trim(GRD1.TextMatrix(intA, colState)) <> "" Then
            MsgBox "本所案號: " & Trim(GRD1.TextMatrix(intA, colCaseNo)) & " " & IIf(Trim(GRD1.TextMatrix(intA, colState)) = "◎", "已通知送件", "已發文"), vbCritical
            Exit Sub
         End If
         If Trim(GRD1.TextMatrix(intA, colCP57)) <> "" Then
            MsgBox "本所案號: " & Trim(GRD1.TextMatrix(intA, colCaseNo)) & " 已取消收文", vbCritical
            Exit Sub
         End If
         If Trim(GRD1.TextMatrix(intA, colCP14)) = "" Then
            MsgBox "本所案號: " & Trim(GRD1.TextMatrix(intA, colCaseNo)) & " 未分案", vbCritical
            Exit Sub
         End If
         '近似本所案不可通知送件

         If Trim(GRD1.TextMatrix(intA, colChkType)) <> "" Then
            MsgBox "本所案號: " & Trim(GRD1.TextMatrix(intA, colCaseNo)) & " 尚在進行覆核流程", vbCritical
            Exit Sub
         End If

         
         If InStr(strAll, Trim(GRD1.TextMatrix(intA, colCaseNo))) = 0 Then
            tmpStr = "" & Trim(GRD1.TextMatrix(intA, id2))
            If InStr(tmpStr, "(圖形查詢)") = 0 Then
               tmpStr = "(" & tmpStr & IIf("" & Trim(GRD1.TextMatrix(intA, id2 + 1)) <> "", "," & Trim(GRD1.TextMatrix(intA, id2 + 1)), "") & ")"
            End If
            '只傳案號
            strAll = strAll & Trim(GRD1.TextMatrix(intA, colCaseNo)) & "||" & Trim(GRD1.TextMatrix(intA, colCP14)) & ","
         End If
         strNoList = strNoList & Trim(GRD1.TextMatrix(intA, colCp09)) & ","
      End If
   Next intA
   '判斷委查單是否查覆完畢
   If PUB_TMACheckOver(strNoList) = False Then
      Exit Sub
   End If
   
   If strAll <> "" Then
      tmpArr = Empty
      tmpArr = Split(strAll, ",")
      For intA = 0 To UBound(tmpArr)
         If InStr(tmpArr(intA), "||") > 0 Then
            PUB_SendMail Trim(Mid(Combo2.Text, 1, InStr(Combo2.Text, " "))), Mid(tmpArr(intA), InStr(tmpArr(intA), "||") + 2), "", Mid(tmpArr(intA), 1, InStr(tmpArr(intA), "||") - 1) & "案，經智權人員確認，請送件!", vbCrLf & "如主旨"
         End If
      Next intA
      
      '同一收文號，只通知一次
      'Memo by Lydia 2016/05/10 同一查名做兩個以上的申請,因為分案會分給同一承辦人,所以不用發兩封信
      'Memo by Lydia 2016/07/07 若有追加查名結果，可再通知
      cnnConnection.BeginTrans
         strSql = "UPDATE TMQCASEMAP SET TQC07=" & strSrvDate(1) & " WHERE TQC02 IN (" & GetAddStr(strNoList) & ") AND TQC07 IS NULL "
         cnnConnection.Execute strSql, intA
      cnnConnection.CommitTrans
      
      If QueryData = False Then ShowNoData
   End If
   
   Exit Sub
ErrHand01:
   
   MsgBox Err.Description, vbCritical
   If strAll <> "" Then cnnConnection.RollbackTrans
   
End Sub

Public Sub PubShowNextData()
Dim mLoad As Boolean
Dim sPath As String, strGrpNo As String
Dim strAttFile As String
Dim rsR As New ADODB.Recordset
Dim oRunform As Form
   
   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
      Set oRunform = frm090801_New
   Else
      Set oRunform = frm090801
   End If
   If m_TMQApp <> "" Then
       m_TMQApp = ""
       QueryData '重整資料
   Else
      If m_NoList <> "" Then
          Me.Enabled = False: mLoad = False
          Screen.MousePointer = vbHourglass
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

          strExc(0) = "select tma01,tma25,tma26,tmf01,tmf02,tmf03 from tmqappform, tmqappfile where tma01 in (" & GetAddStr(m_NoList) & ") and tma01=tmf01(+) and (tmf02(+)='1' or tmf02(+)='2') and tmf03(+)='" & TMQ_附件F04 & "' "
          '以最新單號優先的查詢內容為主
          strExc(0) = strExc(0) & " order by tma01 desc "
          intI = 1
          Set rsR = ClsLawReadRstMsg(intI, strExc(0))
          If intI = 1 Then
             rsR.MoveFirst
             strGrpNo = rsR.Fields("tma01") & " " & rsR.Fields("TMA25")
             txtUnicode(1) = ""
             Do While Not rsR.EOF
                If strGrpNo <> rsR.Fields("tma01") & " " & rsR.Fields("TMA25") Then Exit Do
                
                If "" & rsR.Fields("tma25") <> "2" Then
                   txtUnicode(1) = "" & rsR.Fields("tma26")
                Else  '圖形
                   mLoad = True
                   strAttFile = "" & rsR.Fields("tmf01") & rsR.Fields("tmf02") & rsR.Fields("tmf03")
                End If
                rsR.MoveNext
             Loop
          End If
   
          If txtUnicode(1) <> "" Then
             oRunform.opt1(0).Value = True
             'oRunform.PicText = txtUnicode(1) 'Mark by Lydia 2024/10/07 商標文字欄位中，勿直接帶入文字，以留空方式讓智權人員填寫---杜協理
          ElseIf mLoad = True Then
             sPath = Dir(m_AttachPath & "\" & strAttFile & "*.*")
             If sPath = "" Then
                mLoad = AttachFileGet(Mid(strGrpNo, 1, InStr(strGrpNo, " ") - 1))
             Else
                sPath = m_AttachPath & "\" & sPath
             End If
             If mLoad = True Then
                oRunform.opt1(1).Value = True
                oRunform.optColor(0).Value = True
                Call oRunform.PicToObj(sPath)
             End If
          End If
   
          m_TMQApp = m_NoList
          oRunform.cmdTMQ.Tag = m_NoList
          oRunform.Combo1(0).Text = "000" & " " & GetPrjNationName("000")
         '設定案件性質
          Call oRunform.Text1_LostFocus(6)
          Call oRunform.QueryTMQ
          'TS案無商標種類
          If oRunform.Text1(6) = "TS" Then
          
          ElseIf oRunform.Text1(6) = "T" Then
             oRunform.Combo6.ListIndex = 0 'A接洽單的商標種類
          End If
          oRunform.bolExternalCall = False '還原預設值
          Screen.MousePointer = vbDefault
          Me.Enabled = True
          Me.Hide
      '沒有勾選查名單，繼續回存到接洽單
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
   End If '----If m_TMQApp <> "" Then
      
End Sub

Private Function AttachFileGet(ByVal mTMF01 As String, Optional ByRef strRfilePath As String = "") As Boolean
Dim outType As String
Dim stTempFile As String

On Error GoTo ErrHnd
   
   AttachFileGet = False
   '開啟時,無法刪除,預設下次開啟表單執行刪檔
   outType = "JPG"   '----網中系統限制圖片只能為JPG檔
   Call PUB_KillTempFile(strUserNum & "\H*." & outType)
   Call PUB_KillTempFile(strUserNum & "\H*." & LCase(outType))
   
   strSql = "select * from TMQAPPFile where TMF01='" & mTMF01 & "' AND TMF02='1' AND TMF03='" & TMQ_附件F04 & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If Right(UCase("" & RsTemp.Fields("TMF10")), 4) <> "." & outType Then
         MsgBox "附件非JPG檔，無法載入圖片！", vbCritical
      Else
         stTempFile = m_AttachPath & "\" & mTMF01 & RsTemp.Fields("TMF02") & RsTemp.Fields("TMF03") & "." & LCase(Trim(outType))
         
         strRfilePath = stTempFile
         If PUB_TMQAppFileGet(m_AttachPath, stTempFile, mTMF01, "" & RsTemp.Fields("TMF02"), "" & RsTemp.Fields("TMF03")) = False Then
            MsgBox "無法儲存檔案[ " & stTempFile & " ]！"
            Exit Function
         End If
          
         '預設載入
         Set G_SeekPicColor.Picture = pvGetStdPicture(Trim(stTempFile))
         '固定PictureBox中的image,載入圖片後調整圖片大小
         Call Pub_PicToObj(Trim(stTempFile), G_SeekPicColor, tmpPic, tmpImg)
         AttachFileGet = True
      End If
   End If
   
   Exit Function

ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub textCName_GotFocus()
    TextInverse textCName
End Sub

Private Sub textCName_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Forms(0).PopupMenu2 textCName
End Sub

'取得人員請假的職代
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

'Added by Lydia 2025/04/?? '1234
Private Sub cmdPost_Click()
On Error GoTo 0
Dim strB01 As String, strB02 As String
Dim strResTxt As String, intRetry As Integer
Dim intS As Integer, intE As Integer
Dim oCNHttp As New WinHttp.WinHttpRequest
Dim arrResTxt

   strB01 = Pub_GetSpecMan("TMSearch拆字功能")
   If strB01 <> "" Then
      cmdPost.Enabled = False
      Debug.Print "P1:" & Format(ServerTime, "000000")
      Set oCNHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
      oCNHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
      oCNHttp.SetTimeouts 10000, 10000, 10000, 10000  'Resolve, Connect, Send and Receive
      If Trim(txtFM2(0)) = "" Then
         txtFM2(0) = "全家Family"
      End If
      strExc(1) = PUB_StringFilter(txtFM2(0))
      Debug.Print "P2:" & Format(ServerTime, "000000")
      oCNHttp.Open "GET", strB01 & strExc(1), False
      Debug.Print "P3:" & Format(ServerTime, "000000")
JumpToRetry:
      oCNHttp.Send
      If oCNHttp.Status = 200 Then
         Debug.Print "P4:" & Format(ServerTime, "000000")
         '用全形｜取代JSON區隔的"
         strResTxt = Replace(Replace(Replace(Replace(Replace(oCNHttp.ResponseText, """" & "," & """", "｜,｜"), "[" & """", "[｜"), """" & "]", "｜]"), "{" & """", "{｜"), """" & "}", "｜}")
         txtFM2(1) = strResTxt
         intS = InStr(strResTxt, "[") + 1
         intE = InStr(strResTxt, "]")
         strB02 = Replace(Mid(strResTxt, intS, intE - intS), "｜", "")
         arrResTxt = Empty
         arrResTxt = Split(strB02, ",")
         For intI = 0 To UBound(arrResTxt)
            Debug.Print intI & ":" & arrResTxt(intI)
         Next
         '支援Unicode
         UniMsgBox "輸入文字：" & strExc(1) & vbCrLf & "網站拆字：" & vbCrLf & strResTxt

      Else
         'If HttpClient.Status = 408 Then  '逾時過期(保留)
         'End If
         Debug.Print "Error" & intRetry & ":" & Format(ServerTime, "000000")
         If intRetry < 3 Then
            Sleep 5000
            intRetry = intRetry + 1
            GoTo JumpToRetry
         Else
            MsgBox "呼叫網站拆字功能失敗！", vbCritical, "台一商標查詢系統"
         End If
      End If
      Set oCNHttp = Nothing
      cmdPost.Enabled = True
   Else
      MsgBox "沒有網址！", vbCritical, "台一商標查詢系統"
   End If

End Sub

