VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010001_1 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   3270
   ClientLeft      =   915
   ClientTop       =   630
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6960
   Begin VB.Frame fraLastCaseCode 
      BorderStyle     =   0  '沒有框線
      Height          =   372
      Left            =   3600
      TabIndex        =   27
      Top             =   780
      Visible         =   0   'False
      Width           =   3132
      Begin VB.Label lblCaseCode 
         Height          =   252
         Left            =   1080
         TabIndex        =   28
         Top             =   0
         Width           =   1932
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號："
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   972
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5940
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5100
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   70
      Width           =   800
   End
   Begin VB.Frame fraCode 
      BorderStyle     =   0  '沒有框線
      Height          =   1845
      Left            =   300
      TabIndex        =   17
      Top             =   1185
      Width           =   6492
      Begin VB.Frame fraElse 
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
         Height          =   492
         Left            =   1920
         TabIndex        =   18
         Top             =   -105
         Width           =   2892
         Begin VB.TextBox txtCode 
            Height          =   300
            Index           =   2
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   4
            Top             =   120
            Width           =   492
         End
         Begin VB.TextBox txtCode 
            Height          =   300
            Index           =   1
            Left            =   1440
            MaxLength       =   1
            TabIndex        =   3
            Top             =   120
            Width           =   372
         End
         Begin VB.TextBox txtCode 
            Height          =   300
            Index           =   0
            Left            =   0
            MaxLength       =   6
            TabIndex        =   2
            Top             =   120
            Width           =   1332
         End
      End
      Begin VB.TextBox textCP05 
         Height          =   300
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   5
         Top             =   480
         Width           =   1164
      End
      Begin VB.Frame fraPatition 
         BorderStyle     =   0  '沒有框線
         Height          =   420
         Left            =   -15
         TabIndex        =   23
         Top             =   1380
         Visible         =   0   'False
         Width           =   6492
         Begin VB.TextBox txtPetition 
            Height          =   300
            Left            =   1710
            MaxLength       =   9
            TabIndex        =   7
            Top             =   0
            Width           =   1332
         End
         Begin VB.Label Label5 
            Caption         =   "移轉、讓與申請人："
            Height          =   330
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   1710
         End
         Begin MSForms.Label lblPetitionName 
            Height          =   255
            Left            =   3150
            TabIndex        =   24
            Top             =   30
            Width           =   3255
            VariousPropertyBits=   27
            Size            =   "5741;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.TextBox txtCaseProperty 
         Height          =   300
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   6
         Top             =   936
         Width           =   732
      End
      Begin VB.TextBox txtSystem 
         Height          =   300
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   1
         Top             =   0
         Width           =   675
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
         Height          =   492
         Left            =   1920
         TabIndex        =   19
         Top             =   -120
         Visible         =   0   'False
         Width           =   2772
         Begin VB.TextBox txtTFCode 
            Height          =   300
            Index           =   3
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   11
            Top             =   120
            Width           =   492
         End
         Begin VB.TextBox txtTFCode 
            Height          =   300
            Index           =   2
            Left            =   1680
            MaxLength       =   1
            TabIndex        =   10
            Top             =   120
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   300
            Index           =   1
            Left            =   1200
            MaxLength       =   1
            TabIndex        =   9
            Top             =   120
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   300
            Index           =   0
            Left            =   0
            MaxLength       =   5
            TabIndex        =   8
            Top             =   120
            Width           =   1092
         End
      End
      Begin VB.Label Label4 
         Caption         =   "收文日："
         Height          =   252
         Left            =   0
         TabIndex        =   30
         Top             =   504
         Width           =   972
      End
      Begin VB.Label lblCasePropertyName 
         Height          =   372
         Left            =   1896
         TabIndex        =   22
         Top             =   912
         Width           =   4524
      End
      Begin VB.Label Label3 
         Caption         =   "案件性質："
         Height          =   252
         Left            =   0
         TabIndex        =   21
         Top             =   936
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "本所案號："
         Height          =   252
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   972
      End
   End
   Begin VB.Frame fraRecieve 
      BorderStyle     =   0  '沒有框線
      Height          =   612
      Left            =   1680
      TabIndex        =   16
      Top             =   660
      Width           =   1812
      Begin VB.TextBox txtRecieveCode 
         Height          =   300
         Index           =   1
         Left            =   720
         MaxLength       =   6
         TabIndex        =   0
         Top             =   120
         Width           =   1092
      End
      Begin VB.TextBox txtRecieveCode 
         Height          =   300
         Index           =   0
         Left            =   384
         MaxLength       =   2
         TabIndex        =   14
         Top             =   120
         Width           =   372
      End
      Begin VB.Label lblReciveCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   252
      End
   End
   Begin VB.Label lblRecieveKind 
      Caption         =   "上一筆之收文號："
      Height          =   252
      Left            =   240
      TabIndex        =   15
      Top             =   780
      Width           =   1452
   End
End
Attribute VB_Name = "frm010001_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/14 Form2.0已修改 lblPetitionName
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
Option Explicit

'intChoose   0:收文   1:內部收文
Public intChoose As Integer
'intSaveMode，To A:-1為錯誤之本所案號，0為重複之本所案號(新增舊案)，1為正確之本所案號(新增新案)
'intCaseKind，1為專利，2為商標，3為法務，4為顧問，5為專利(服)，6為商標(服)，7為法務(服)，8為顧問(服)
'strCaseName，取得案件名稱
Public intSaveMode As Integer, intCaseKind As Integer, strCaseName As String
'strReceiveKind ，A為接洽記錄單，B為政府機關來函
Dim strReceiveKind As String
'intReceiveKind=0為接洽紀錄單;=1為政府來函
'intModifyKind=0為新增;=1為修改;=2為查詢
Public intReceiveKind As Integer, intModifyKind As Integer
Dim adoquery As New ADODB.Recordset
'Add By Cheng 2003/09/08
Public m_blnNewCase As Boolean '判斷是否為新案(無流水號或基本檔無資料)

Private Sub cmdOK_Click(Index As Integer)
Dim bolErr As Boolean, strAutoNumber As String, strTemp As String
Dim strCode0 As String, strCode1 As String, strCode2 As String, strCode3 As String
Dim intRt As Integer

If Index = 1 Then
   'edit by nickc 2007/02/06 不用 dll 了
   'Set obj001 = Nothing
   Unload Me
Else

   '911111 nick 檢查前兩欄不可空白，新增時，內部收文
   '***** start
   If intModifyKind = 0 Then
      If intChoose = 1 Then
        If txtSystem = "TF" Then
            If txtTFCode(0) = "" Then
                 MsgBox "本所案號不可空白！"
                 GoTo EXITSUB
            End If
        Else
            If txtCode(0) = "" Then
                 MsgBox "本所案號不可空白！"
                 GoTo EXITSUB
            End If
        End If
      End If
   End If
   '***** end
   
   ' 91.03.25 modify by louis
   '91.11.10 MODIFY BY SONIA
   'txtCaseProperty_Validate False
   bolErr = False
   txtCaseProperty_Validate bolErr
   If bolErr = True Then
      Exit Sub
   End If
   
   
   ' 90.12.19 modify by louis
   If txtPetition.Visible = True Then
      If IsEmptyText(txtPetition) Then
         MsgBox "請輸入移轉、讓與申請人", vbOKOnly + vbCritical, "檢核資料"
         txtPetition.SetFocus
         Exit Sub
      '911111 nick 檢查申請人是否有輸入
      '***** start
      Else
          bolErr = False
          txtPetition_Validate bolErr
          If bolErr = True Then
             Exit Sub
          End If
      '***** end
      End If
   End If
   
   ' 91.09.16 modify by louis
   'Add By Cheng 2001/12/12
   'If Me.txtSystem.Visible Then
   '   bolErr = False
   '   txtSystem_Validate bolErr
   '   If bolErr Then txtSystem.SetFocus: Exit Sub
   'End If
   If Me.txtSystem.Visible Then
      If CheckEverythingOK() = False Then
         GoTo EXITSUB
      End If
   End If
   
   ' 91.09.04 modify by louis
   If textCP05 = "111111" Then
        'Add By Cheng 2003/11/25
        '檢查流水號是否大於自動編號
        If CheckCaseNo = True Then
            If Me.txtSystem.Text = "TF" Then
                Me.txtTFCode(0).SetFocus
                txtTFCode_GotFocus 0
            Else
                Me.txtCode(0).SetFocus
                txtCode_GotFocus 0
            End If
            GoTo EXITSUB
        End If
        'End
      OnSaveNewCP
      txtSystem = Empty
      txtCode(0) = Empty
      txtCode(1) = Empty
      txtCode(2) = Empty
      textCP05 = Empty
      txtCaseProperty = Empty
      lblCasePropertyName = Empty
      txtPetition = Empty
      lblPetitionName = Empty
      fraRecieve.Enabled = False
      fraCode.Visible = True
      txtSystem.SetFocus
      fraLastCaseCode.Visible = True
      GoTo EXITSUB
   End If
    'Add By Cheng 2003/09/08
    'Begin
    If Me.txtSystem.Text = 馬德里案 Then
        m_blnNewCase = CheckNewCase(Me.txtSystem.Text, Me.txtTFCode(0).Text & Me.txtTFCode(1).Text, Me.txtTFCode(2).Text, Me.txtTFCode(3).Text)
    Else
        m_blnNewCase = CheckNewCase(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text)
    End If
    'End
   ' 91.09.04 modify by louis
   If intChoose <> 0 Then
      '911111 nick 若為內部收文，案號不存再不新增
      If intModifyKind = 0 And CheckExist(0) = True Then
          MsgBox "找不到此本所案號在基本檔之資料"
          txtSystem.SetFocus
          GoTo EXITSUB
      End If
        'Modify By Cheng 2003/03/28
'      OnNextForm
        If OnNextForm = True Then
            Me.Hide
            GoTo EXITSUB
        Else
            GoTo EXITSUB
        End If
   End If
   
   bolErr = False
   txtCaseProperty_Validate bolErr
   If bolErr Then Exit Sub
   If intChoose = 1 And intSaveMode = 1 Then
      ShowMsg MsgText(1013)
      Exit Sub
   End If
 
   If intModifyKind <> 0 Then
      If CheckEverythingOK = False Then
         Exit Sub
      Else
         If txtCaseProperty.Visible = True And lblCasePropertyName = "" Then
            ShowMsg MsgText(1014)
            txtCaseProperty.SetFocus
            txtCaseProperty_GotFocus
            Exit Sub
         Else
            If txtPetition.Visible = True Then
               txtPetition_Validate False
               If lblPetitionName = "" Then
                  ShowMsg MsgText(1015)
                  txtPetition.SetFocus
                  txtPetition_GotFocus
                  Exit Sub
               End If
            End If
         End If
      End If
   End If
   Select Case intReceiveKind
      Case 0
         Select Case intModifyKind
            Case 0
              '911017 nick
              'If txtCode(0) <> "" Then
              If txtCode(0) <> "" Or txtTFCode(0) <> "" Then
                 '911017 nick 將 1 和 2  對調
                 If CheckCaseNo Then   '----1
                    Exit Sub
                 End If
                 If CheckExist(1) Then    '------2
                    Exit Sub
                 End If
              Else
                 intSaveMode = 1
              End If
              If intChoose = 1 Then
                 If txtCode(0).Text <> "" Then
                    adoquery.CursorLocation = adUseClient
                    adoquery.Open "select np06, np07 from nextprogress where np02 = '" & txtSystem & "' and np03 = '" & txtCode(0).Text & "' and np04 = '" & IIf(txtCode(1).Text = "", "0", txtCode(1).Text) & "' and np05 = '" & IIf(txtCode(2).Text = "", "00", txtCode(2).Text) & "'", cnnConnection, adOpenStatic, adLockReadOnly
                    If adoquery.RecordCount < 2 And adoquery.RecordCount > 0 Then
                       If adoquery.Fields("np07").Value <> txtCaseProperty.Text Then
                          If IsNull(adoquery.Fields(0).Value) = True Or adoquery.Fields(0).Value = "" Then
                             ShowMsg "下一程序的人有未收文之資料，請自行處理"
                          End If
                       End If
                    End If
                    adoquery.Close
                  End If
              End If
              'edit by nickc 2007/02/02 不用 dll 了
              'objPublicData.GetSystemKind txtSystem.Text, intCaseKind, strCaseName
              ClsPDGetSystemKind txtSystem.Text, intCaseKind, strCaseName
              Select Case intCaseKind
                  Case 專利
                     If intModifyKind <> 0 Then
                        If CheckKeyInOkay = False Then
                           Exit Sub
                        End If
                     End If
                     frm010005.Caption = Me.Caption + "－" + strCaseName
                     frm010005.txtRecieveCode.Text = strReceiveKind + txtRecieveCode(0).Text
                     frm010005.txtPatent(1) = txtCaseProperty
                     frm010005.lblCaseProperty = lblCasePropertyName
                     If txtPetition.Visible = True Then
                        frm010005.fraPatition.Visible = True
                        frm010005.txtPatent(23) = txtPetition
                        frm010005.lblPetitionName = lblPetitionName
                        
                        'Add By Cheng 2002/01/14
                        Select Case Me.txtCaseProperty.Text
                        Case 合併
                           frm010005.Label29.Caption = "合併申請人："
                        Case 繼承
                           frm010005.Label29.Caption = "繼承申請人："
                        End Select
                        
                     Else
                        frm010005.fraPatition.Visible = False
                     End If
                     frm010005.txtSystem.Text = txtSystem.Text
                     frm010005.txtCode(0).Text = txtCode(0).Text
                     frm010005.txtCode(1).Text = txtCode(1).Text
                     frm010005.txtCode(2).Text = txtCode(2).Text
                     'modify by sonia 90.10.8
                     If frm010005.txtSystem = "FCP" And Me.intChoose = 1 Then
                        frm010005.txtPatent(24) = strUserNum
                     'Add By Cheng 2001/12/12
                     Else
                        frm010005.txtPatent(24) = ""
                     End If
                     
                     frm010005.Show
                  Case 商標
                     If intModifyKind <> 0 Then
                        If CheckKeyInOkay = False Then
                           Exit Sub
                        End If
                     End If
                     frm010004.Caption = Me.Caption + "－" + strCaseName
                     frm010004.txtRecieveCode.Text = strReceiveKind + txtRecieveCode(0).Text
                     frm010004.txtSystem.Text = txtSystem.Text
                     frm010004.txtTrademark(1) = txtCaseProperty
                     frm010004.lblCaseProperty = lblCasePropertyName
                     If txtPetition.Visible = True Then
                        frm010004.fraPatition.Visible = True
                        frm010004.txtTrademark(20) = txtPetition
                        'edit by nickc 2006/11/22
                        'frm010004.lblPetitionName = lblPetitionName
                        frm010004.lblPetitionName(0) = lblPetitionName
                     Else
                        frm010004.fraPatition.Visible = False
                     End If
                     'TF為馬德里案，另外判斷
                     If txtSystem.Text <> 馬德里案 Then
                        frm010004.fraElse.Visible = True
                        frm010004.fraTF.Visible = False
                        frm010004.txtCode(0).Text = txtCode(0).Text
                        frm010004.txtCode(1).Text = txtCode(1).Text
                        frm010004.txtCode(2).Text = txtCode(2).Text
                     Else
                        frm010004.fraElse.Visible = False
                        frm010004.fraTF.Visible = True
                        frm010004.txtTFCode(0).Text = txtTFCode(0).Text
                        frm010004.txtTFCode(1).Text = txtTFCode(1).Text
                        frm010004.txtTFCode(2).Text = txtTFCode(2).Text
                        frm010004.txtTFCode(3).Text = txtTFCode(3).Text
                      End If
                      'Add by Morgan 2003/11/26
                      frm010004.fraTM15.Visible = txtTM15Control()
                      '---End
                      frm010004.Show
                  Case Else
                     If intModifyKind <> 0 Then
                        If CheckKeyInOkay Then
                           Exit Sub
                        End If
                     End If
                     If Me.intCaseKind = 顧問 And txtCaseProperty = 顧問聘任 Then
                        'Added by Lydia 2021/12/13
                        '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                        If PUB_CheckFormExist("frm010006") = False Then
                           Set frm010006 = Nothing
                        End If
                        'end 2021/12/13
                        frm010006.Caption = Me.Caption + "－" + strCaseName
                        frm010006.txtRecieveCode.Text = strReceiveKind + txtRecieveCode(0).Text
                        frm010006.txtAdviser(1) = txtCaseProperty
                        frm010006.lblCaseProperty = lblCasePropertyName
                        frm010006.txtSystem.Text = txtSystem.Text
                        frm010006.txtCode(0).Text = txtCode(0).Text
                        frm010006.txtCode(1).Text = txtCode(1).Text
                        frm010006.txtCode(2).Text = txtCode(2).Text
                        frm010006.Show
                     Else
                        frm010007.Caption = Me.Caption + "－" + strCaseName
                        frm010007.txtRecieveCode.Text = strReceiveKind + txtRecieveCode(0).Text
                        frm010007.txtOther(1) = txtCaseProperty
                        frm010007.lblCaseProperty = lblCasePropertyName
                        frm010007.txtSystem.Text = txtSystem.Text
                        frm010007.txtCode(0).Text = txtCode(0).Text
                        frm010007.txtCode(1).Text = txtCode(1).Text
                        frm010007.txtCode(2).Text = txtCode(2).Text
                        frm010007.intCaseKind = intCaseKind
                        
                        frm010007.Show
                     End If
              End Select
            Case 1, 2
               'edit by nickc 2007/02/02 不用 dll 了
               'intRt = objPublicData.CheckRecieveCode(strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, strCode0, strCode1, strCode2, strCode3)
               intRt = ClsPDCheckRecieveCode(strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, strCode0, strCode1, strCode2, strCode3)
               If intRt <> 0 Then
                   'edit by nickc 2007/02/02 不用 dll 了
                   'If objPublicData.GetSystemKind(strCode0, intCaseKind, strCaseName) = False Then
                   If ClsPDGetSystemKind(strCode0, intCaseKind, strCaseName) = False Then
                      Exit Sub
                   End If
                   Select Case intCaseKind
                     Case 專利
                        frm010005.Caption = Me.Caption + "－" + strCaseName
                        frm010005.txtRecieveCode = strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text
                        frm010005.txtSystem.Text = strCode0
                        frm010005.txtCode(0).Text = strCode1
                        frm010005.txtCode(1).Text = IIf(strCode2 = "0", "", strCode2)
                        frm010005.txtCode(2).Text = IIf(strCode3 = "00", "", strCode3)
                        frm010005.Show
                     Case 商標
                        frm010004.Caption = Me.Caption + "－" + strCaseName
                        frm010004.txtRecieveCode = strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text
                        frm010004.txtSystem.Text = strCode0
                        'TF為馬德里案，另外判斷
                        If strCode0 <> 馬德里案 Then
                           frm010004.fraElse.Visible = True
                           frm010004.fraTF.Visible = False
                           frm010004.txtCode(0).Text = strCode1
                           frm010004.txtCode(1).Text = IIf(strCode2 = "0", "", strCode2)
                           frm010004.txtCode(2).Text = IIf(strCode3 = "00", "", strCode3)
                        Else
                           frm010004.fraElse.Visible = False
                           frm010004.fraTF.Visible = True
                           frm010004.txtTFCode(0).Text = Left(strCode1, 5)
                           frm010004.txtTFCode(1).Text = IIf(Right(strCode1, 1) = "0", "", Right(strCode1, 1))
                           frm010004.txtTFCode(2).Text = IIf(strCode2 = "0", "", strCode2)
                           frm010004.txtTFCode(3).Text = IIf(strCode3 = "00", "", strCode3)
                        End If
                        frm010004.Show
                     Case Else
                        If Me.intCaseKind = 顧問 And intRt = 2 Then
                            'Added by Lydia 2021/12/13
                            '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                            If PUB_CheckFormExist("frm010006") = False Then
                               Set frm010006 = Nothing
                            End If
                            'end 2021/12/13
                           frm010006.Caption = Me.Caption + "－" + strCaseName
                           frm010006.txtRecieveCode = strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text
                           frm010006.txtSystem.Text = strCode0
                           frm010006.txtCode(0).Text = strCode1
                           frm010006.txtCode(1).Text = IIf(strCode2 = "0", "", strCode2)
                           frm010006.txtCode(2).Text = IIf(strCode3 = "00", "", strCode3)
                           frm010006.Show
                        Else
                           frm010007.Caption = Me.Caption + "－" + strCaseName
                           frm010007.txtRecieveCode = strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text
                           frm010007.txtSystem.Text = strCode0
                           frm010007.txtCode(0).Text = strCode1
                           frm010007.txtCode(1).Text = IIf(strCode2 = "0", "", strCode2)
                           frm010007.txtCode(2).Text = IIf(strCode3 = "00", "", strCode3)
                           frm010007.Show
                        End If
                  End Select
               Else
                   bolErr = True
               End If
         End Select
      Case 1
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.CheckRecieveCode(strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, strCode0, strCode1, strCode2, strCode3) <> 0 Then
         If ClsPDCheckRecieveCode(strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, strCode0, strCode1, strCode2, strCode3) <> 0 Then
             frm010002.lblRecieveCode = strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text
             frm010002.Caption = Me.Caption
         Else
             Exit Sub
         End If
   End Select
   If bolErr = False Then
       Me.Hide
   End If
End If
EXITSUB:
End Sub
'由於使用Validate()，以致無法正確跳躍Focus，因此在txtPetition_GotFocus()及
'cmdOK_GotFocus()加入判斷，以得正確之跳躍
Private Sub cmdOK_GotFocus(Index As Integer)
Static boltxtPetition As Boolean

If Index = 0 Then
   If fraPatition.Visible Then
      If boltxtPetition = False Then
         txtPetition.SetFocus
         boltxtPetition = True
      End If
   Else
      boltxtPetition = False
   End If
End If
End Sub
Public Sub ClearForm(ByVal strAuto1 As String, ByVal strAuto2 As String)
Dim i As Integer, bolNewCaseCode As Boolean

If intModifyKind = 0 Then
   txtRecieveCode(1).Text = Mid(strAuto1, 4)
   If strAuto2 <> "" Then
      If txtSystem = 馬德里案 Then
         If txtTFCode(0) <> strAuto2 Then bolNewCaseCode = True
      Else
         If txtCode(0) <> strAuto2 Then bolNewCaseCode = True
      End If
   End If
   If bolNewCaseCode Then
      lblCaseCode = txtSystem + "- " + strAuto2
   Else
      lblCaseCode = txtSystem
      If txtSystem = 馬德里案 Then
         For i = 0 To 3
                If txtTFCode(i) <> "" Then
                   lblCaseCode = lblCaseCode + "- " + txtTFCode(i)
                Else
                   Exit For
                End If
         Next
      Else
         For i = 0 To 2
                If txtCode(i) <> "" Then
                   lblCaseCode = lblCaseCode + "- " + txtCode(i)
                Else
                   Exit For
                End If
         Next
      End If
   End If
End If
txtSystem = ""
For i = 0 To 2
       txtCode(i) = ""
Next
For i = 0 To 3
       txtTFCode(i) = ""
Next
txtCaseProperty = ""
txtPetition = ""
End Sub
Private Sub Form_Activate()
'edit by nickc 2007/02/06 不用 dll 了
'If obj001 Is Nothing Then
'   Set obj001 = CreateObject("prjTaieDll001.cls001")
'   Set obj001.Connection = cnnConnection
   
'End If
'Modify By Sindy 2010/8/17 比對自動編號年度
'txtRecieveCode(0).Text = GetTaiwanThisYear
txtRecieveCode(0).Text = CompAutoNumberYear(GetTaiwanThisYear)
'Add By Cheng 2002/07/17
strReceiveKind = ""
Select Case intReceiveKind
             Case 0
                        If intChoose = 1 Then
                           strReceiveKind = 內部收文
                        Else
                           strReceiveKind = 接洽記錄單
                        End If
                        Select Case intModifyKind
                                     Case 0
                                               '新增：輸入本所案號
                                               fraRecieve.Enabled = False
                                               fraCode.Visible = True
                                               lblRecieveKind = "上一筆之收文號："
                                               txtSystem.SetFocus
                                               fraLastCaseCode.Visible = True
                                     Case 1, 2
                                               '修改，刪除：輸入收文號
                                               lblRecieveKind = "收文號："
                                               fraRecieve.Enabled = True
                                               fraCode.Visible = False
                                               txtRecieveCode(1).SetFocus
                        End Select
             Case 1
                        strReceiveKind = 政府機關來函
                        fraCode.Visible = False
                        '修改，刪除：可輸入收文號
                        fraRecieve.Enabled = True
End Select
lblReciveCode.Caption = strReceiveKind
End Sub
Private Sub Form_Load()
intSaveMode = 0
MoveFormToCenter Me

   ' 91.09.03 modify by louis
   textCP05 = Empty
   If intChoose = 0 Then
      EnableTextBox textCP05, False
      Label4.Visible = False
      textCP05.Visible = False
   Else
      Label4.Visible = True
      textCP05.Visible = True
      EnableTextBox textCP05, True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/07/18
Set frm010001_1 = Nothing
End Sub

Private Sub textCP05_GotFocus()
   InverseTextBox textCP05
End Sub

Private Sub textCP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP05) = False Then
      If textCP05 <> "111111" Then
         Cancel = True
         strMsg = "收文日只可為空白或111111"
         strTit = "收文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05_GotFocus
      End If
   End If
End Sub

Private Sub txtCaseProperty_Change()

   lblCasePropertyName = ""
   
End Sub

'Add by Morgan 2003/11/26
Private Function txtTM15Control() As Boolean

   If (txtCaseProperty = "102") And ((txtTFCode(0).Text = Empty And txtSystem.Text = "TF") Or _
      (txtCode(0).Text = Empty And (txtSystem.Text = "T" Or txtSystem.Text = "FCT"))) Then
      txtTM15Control = True
   Else
      txtTM15Control = False
   End If
   
End Function

Private Sub txtCaseProperty_Validate(Cancel As Boolean)
   Dim strTemp As String
   Dim bolIsChina As Boolean
   
   If txtCaseProperty.Visible = False Then
      Exit Sub
   End If
   If txtCaseProperty <> "" Then
      'Add By Cheng 2002/01/08
      '若系統類別非"L", "CFL", "FCL", "LA"檢查案件性質必須為三碼
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      If Me.txtSystem.Text <> "L" And Me.txtSystem.Text <> "CFL" And _
         Me.txtSystem.Text <> "FCL" And Me.txtSystem.Text <> "LA" And _
         Me.txtSystem.Text <> "LIN" And Me.txtSystem.Text <> "ACS" Then
         If Len(Me.txtCaseProperty.Text) <> 3 Then
            MsgBox "案件性質必須為三碼!!!", vbExclamation
            Cancel = True
            txtCaseProperty_GotFocus
            Me.txtCaseProperty.SetFocus
            Exit Sub
         End If
      End If
      
      If CheckExist(0) = False Then
         If CheckEverythingOK = False Then
            Cancel = True
            txtCaseProperty_GotFocus
         End If
      Else
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(txtSystem, txtCaseProperty, strTemp, bolIsChina) Then
         If ClsPDGetCaseProperty(txtSystem, txtCaseProperty, strTemp, bolIsChina) Then
            lblCasePropertyName = strTemp
            '92.2.19 ADD BY SONIA
            If txtCaseProperty = "001" And txtSystem <> "TS" And txtSystem <> "S" Then
               MsgBox "查名案件性質之系統類別只可為 TS(內商) 或 S(FCT,CFT)  !!!", vbExclamation
               Cancel = True
               txtCaseProperty_GotFocus
               Me.txtCaseProperty.SetFocus
            End If
            '92.2.19 END
         Else
            lblCasePropertyName = ""
            Cancel = True
            txtCaseProperty_GotFocus
         End If
      End If
      ' 90.12.19 add by louis
      UpdateCtrlState
   Else
      ShowMsg MsgText(1016)
      Cancel = True
      txtCaseProperty_GotFocus
      Me.txtCaseProperty.SetFocus
   End If
End Sub
Private Sub txtCaseProperty_GotFocus()
txtCaseProperty.SelStart = 0
txtCaseProperty.SelLength = Len(txtCaseProperty)
End Sub

Private Sub txtCode_Change(Index As Integer)
   Call txtTM15Control
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
txtCode(Index).SelStart = 0
txtCode(Index).SelLength = Len(txtCode(Index).Text)
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
   If Len(txtCode(Index)) = txtCode(Index).MaxLength Then
      If CheckExist(0) = False Then
         If CheckEverythingOK = False Then
            If Index = 0 Then
               SendKeys "{TAB}"
               SendKeys "{TAB}"
            Else
               Cancel = True
               txtCode_GotFocus Index
            End If
         End If
      End If
      ' 90.12.19 add by louis
      UpdateCtrlState
   ElseIf Len(txtCode(Index)) <> 0 Then
      ShowMsg MsgText(1017)
      Cancel = True
      txtCode_GotFocus Index
   End If
End Sub
Private Sub txtPetition_Change()
lblPetitionName.Caption = ""
End Sub
Private Sub txtPetition_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtPetition_Validate(Cancel As Boolean)
Dim strTemp As String, intLength As Integer, strPetition As String

If Len(txtPetition) > 0 Then
   strPetition = txtPetition
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.GetCustomer(strPetition, strTemp) Then
   If ClsPDGetCustomer(strPetition, strTemp) Then
      txtPetition = strPetition
      lblPetitionName = strTemp
   Else
      Cancel = True
      txtPetition_GotFocus
   End If
End If
End Sub
Private Sub txtPetition_GotFocus()
'由於使用Validate()，以致無法正確跳躍Focus，因此在txtPetition_GotFocus()及
'cmdOK_GotFocus()加入判斷，以得正確之跳躍
If txtPetition.Visible = False Then
   cmdOK(0).SetFocus
Else
   txtPetition.SelStart = 0
   txtPetition.SelLength = Len(txtPetition)
End If
End Sub
Private Sub txtRecieveCode_GotFocus(Index As Integer)
txtRecieveCode(Index).SelStart = 0
txtRecieveCode(Index).SelLength = Len(txtRecieveCode(Index).Text)
End Sub
Private Sub txtSystem_Change()
   If txtSystem.Text = 馬德里案 Then
      fraTF.Visible = True
      fraElse.Visible = False
      '911111 nick
      txtTFCode(0).Enabled = True
      txtTFCode(1).Enabled = True
      txtTFCode(2).Enabled = True
      txtTFCode(3).Enabled = True
      txtCode(0).Enabled = False
      txtCode(1).Enabled = False
      txtCode(2).Enabled = False
      txtTFCode(0).Text = Empty
      txtTFCode(1).Text = Empty
      txtTFCode(2).Text = Empty
      txtTFCode(3).Text = Empty
      txtTFCode(0).TabIndex = 2
      txtTFCode(1).TabIndex = 3
      txtTFCode(2).TabIndex = 4
      txtTFCode(3).TabIndex = 5
      textCP05.TabIndex = 6
      '911017 nick
      txtCaseProperty.TabIndex = 7
      
   Else
      fraTF.Visible = False
      fraElse.Visible = True
      '911111 nick
      txtCode(0).Enabled = True
      txtCode(1).Enabled = True
      txtCode(2).Enabled = True
      txtTFCode(0).Enabled = False
      txtTFCode(1).Enabled = False
      txtTFCode(2).Enabled = False
      txtTFCode(3).Enabled = False
      txtCode(0).Text = Empty
      txtCode(1).Text = Empty
      txtCode(2).Text = Empty
      txtCode(0).TabIndex = 2
      txtCode(1).TabIndex = 3
      txtCode(2).TabIndex = 4
      textCP05.TabIndex = 5
      '911017 nick
      txtCaseProperty.TabIndex = 6
   
      
   End If
   'Add by Morgan 2003/11/24
   If (txtSystem.Text = "T" Or txtSystem.Text = "TF" Or txtSystem.Text = "FCT") Then
      Call txtTM15Control
   End If
   '---End
End Sub

Private Sub txtSystem_GotFocus()
txtSystem.SelStart = 0
txtSystem.SelLength = Len(txtSystem.Text)
End Sub
Private Sub txtSystem_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSystem_LostFocus()
If txtSystem.Text = 馬德里案 Then
   fraTF.Visible = True
   fraElse.Visible = False
   '911111 nick
   txtTFCode(0).Enabled = True
   txtTFCode(1).Enabled = True
   txtTFCode(2).Enabled = True
   txtTFCode(3).Enabled = True
   txtCode(0).Enabled = False
   txtCode(1).Enabled = False
   txtCode(2).Enabled = False
   txtTFCode(0).TabIndex = 2
   txtTFCode(1).TabIndex = 3
   txtTFCode(2).TabIndex = 4
   txtTFCode(3).TabIndex = 5

   '911017 nick
   txtCaseProperty.TabIndex = 11
Else
   fraTF.Visible = False
   fraElse.Visible = True
   '911111 nick
   txtCode(0).Enabled = True
   txtCode(1).Enabled = True
   txtCode(2).Enabled = True
   txtTFCode(0).Enabled = False
   txtTFCode(1).Enabled = False
   txtTFCode(2).Enabled = False
   txtTFCode(3).Enabled = False
   txtCode(0).TabIndex = 2
   txtCode(1).TabIndex = 3
   txtCode(2).TabIndex = 4
   
   '911017 nick
   txtCaseProperty.TabIndex = 6

   
End If

End Sub

Private Sub txtSystem_Validate(Cancel As Boolean)
   ' 91.09.16 modify by louis (變更Flag值)
   'If objPublicData.GetSystemKind(txtSystem.Text, intCaseKind, strCaseName) = False Then
   '   Cancel = True
   '   txtSystem_GotFocus
   'Else
   '   CheckEverythingOK
   'End If
    'Modify By Cheng 2003/03/28
    '若系統類別欄有顯示才要檢查系統類別
    If Me.txtSystem.Visible = True Then
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GetSystemKind(txtSystem.Text, intCaseKind, strCaseName) = False Then
        If ClsPDGetSystemKind(txtSystem.Text, intCaseKind, strCaseName) = False Then
           Cancel = True
           txtSystem_GotFocus
        End If
    End If
End Sub
Private Sub txtTFCode_GotFocus(Index As Integer)
txtTFCode(Index).SelStart = 0
txtTFCode(Index).SelLength = Len(txtTFCode(Index).Text)
End Sub
Private Sub txtTFCode_Validate(Index As Integer, Cancel As Boolean)
If Index = 1 Then
If Len(txtTFCode(Index)) = txtTFCode(Index).MaxLength Then
   If CheckEverythingOK = False Then
      Cancel = True
      txtTFCode_GotFocus Index
   End If
ElseIf Len(txtTFCode(Index)) <> 0 Then
   ShowMsg MsgText(1017)
   Cancel = True
   txtTFCode_GotFocus Index
End If
End If
End Sub
Private Function CheckKeyInOkay() As Boolean
'TF為馬德里案，另外判斷
If txtSystem = 馬德里案 Then
   'edit by nickc 2007/02/06 不用 dll 了
   'If obj001.CheckTFTextOkay(txtSystem.Text, intSaveMode, txtTFCode(0), txtTFCode(1), txtTFCode(2), txtTFCode(3)) Then
   If Cls001CheckTFTextOkay(txtSystem.Text, intSaveMode, txtTFCode(0), txtTFCode(1), txtTFCode(2), txtTFCode(3)) Then
      CheckKeyInOkay = True
   End If
Else
   'edit by nickc 2007/02/06 不用 dll 了
   'If obj001.CheckTextOkay(intCaseKind, txtSystem.Text, intSaveMode, txtCode(0), txtCode(1), txtCode(2)) Then
   If Cls001CheckTextOkay(intCaseKind, txtSystem.Text, intSaveMode, txtCode(0), txtCode(1), txtCode(2)) Then
      CheckKeyInOkay = True
   End If
End If
End Function
Private Function CheckEverythingOK() As Boolean
Dim strTemp As String, bolIsChina As Boolean
'Add By Cheng 2003/08/28
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

If Len(txtCaseProperty) > 0 Then
   If CheckKeyInOkay Then
      If intCaseKind = 顧問 And intSaveMode = 1 And txtCaseProperty <> 顧問聘任 Then
         ShowMsg MsgText(1018)
         Exit Function
      End If
      If intSaveMode = 0 Then
         If intCaseKind <> 顧問 And intCaseKind <> 法務 Then
            If txtSystem = 馬德里案 Then
               'edit by nickc 2007/02/02 不用 dll 了
               'If objPublicData.GetCaseNation(intCaseKind, txtSystem, txtTFCode(0) + txtTFCode(1), IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strTemp) = False Then
               If ClsPDGetCaseNation(intCaseKind, txtSystem, txtTFCode(0) + txtTFCode(1), IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strTemp) = False Then
                  Exit Function
               End If
            Else
               'edit by nickc 2007/02/02 不用 dll 了
               'If objPublicData.GetCaseNation(intCaseKind, txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strTemp) = False Then
               If ClsPDGetCaseNation(intCaseKind, txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strTemp) = False Then
                  Exit Function
               End If
            End If
         End If
         If strTemp = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
      End If
        'Add By Cheng 2003/08/28
        If (Me.txtSystem.Text = "P" Or Me.txtSystem.Text = "CFP" Or Me.txtSystem.Text = "FCP") And Me.txtCaseProperty.Text = "601" Then
            '是否有A類未取消收文的領證(601)
            StrSQLa = "Select Count(*) From CaseProgress Where " & ChgCaseprogress(Me.txtSystem.Text & Me.txtCode(0).Text & Me.txtCode(1).Text & Me.txtCode(2).Text) & " And CP09<'B' And CP10='601' And CP57 Is Null "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                If rsA.Fields(0).Value > 0 Then
                    If rsA.State <> adStateClosed Then rsA.Close
                    Set rsA = Nothing
                    '是否有C類被異議(理由)(1801)
                    StrSQLa = "Select Count(*) From CaseProgress Where " & ChgCaseprogress(Me.txtSystem.Text & Me.txtCode(0).Text & Me.txtCode(1).Text & Me.txtCode(2).Text) & " And CP09>'C' And CP10='1801' "
                    rsA.CursorLocation = adUseClient
                    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                    If rsA.RecordCount > 0 Then
                        If rsA.Fields(0).Value <= 0 Then
                            MsgBox "本案不可再收<領證>!!!", vbExclamation + vbOKOnly
                            Me.txtCaseProperty.SetFocus
                            txtCaseProperty_GotFocus
                            CheckEverythingOK = False
                            Exit Function
                        End If
                    End If
                End If
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
        End If
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetCaseProperty(txtSystem, txtCaseProperty, strTemp, bolIsChina) Then
      If ClsPDGetCaseProperty(txtSystem, txtCaseProperty, strTemp, bolIsChina) Then
         lblCasePropertyName = strTemp
         '92.9.15 MODIFY BY SONIA
         'If (((txtCaseProperty = 申請 Or txtCaseProperty = 異議 Or txtCaseProperty = 評定 Or txtCaseProperty = 廢止) And intCaseKind Mod 4 = 商標) Or _
         '      ((txtCaseProperty = 發明申請 Or txtCaseProperty = 新型申請 Or txtCaseProperty = 設計申請 Or txtCaseProperty = 追加申請 Or txtCaseProperty = 聯合申請 Or txtCaseProperty = 異議_專 Or txtCaseProperty = 舉發) _
         '      And intCaseKind Mod 4 = 專利)) And intSaveMode <> 1 Then
         If (((txtCaseProperty = 申請 Or txtCaseProperty = 異議 Or txtCaseProperty = 評定 Or txtCaseProperty = 廢止) And intCaseKind Mod 4 = 商標) Or _
               ((txtCaseProperty = 發明申請 Or txtCaseProperty = 新型申請 Or txtCaseProperty = 設計申請 Or txtCaseProperty = 追加申請 Or txtCaseProperty = 聯合申請 Or txtCaseProperty = 異議_專 Or txtCaseProperty = 舉發 Or txtCaseProperty = PCT申請 Or txtCaseProperty = 記錄請求_標準專利 Or txtCaseProperty = 短期專利申請 Or txtCaseProperty = CIP申請 Or txtCaseProperty = CPA申請 Or txtCaseProperty = 再發行) _
               And intCaseKind Mod 4 = 專利)) And intSaveMode <> 1 Then
         '92.9.15 END
            ' 91.09.03 marked by louis
            'ShowMsg MsgText(1019)
            'txtCaseProperty.SetFocus
            'CheckEverythingOK = False
            ' 91.09.03 modify by louis
            If textCP05 <> "111111" Then
               ShowMsg MsgText(1019)
               txtCaseProperty.SetFocus
               CheckEverythingOK = False
            Else
               CheckEverythingOK = True
            End If
         ElseIf txtCaseProperty = 移轉 And intCaseKind Mod 4 = 商標 Then
            fraPatition.Visible = True
            CheckEverythingOK = True
         ElseIf txtCaseProperty = 讓與 And intCaseKind Mod 4 = 專利 Then
            fraPatition.Visible = True
            CheckEverythingOK = True
         'Add By Cheng 2002/01/11
         ElseIf txtCaseProperty = 專利權讓與 And intCaseKind Mod 4 = 專利 Then
            If Me.txtSystem.Text = "P" Then
               fraPatition.Visible = True
               CheckEverythingOK = True
            End If
         Else
            fraPatition.Visible = False
            CheckEverythingOK = True
         End If
      End If
   End If
Else
   lblCasePropertyName = ""
   fraPatition.Visible = False
   CheckEverythingOK = True
End If
End Function


'檢查本所案號是否存在
Public Function CheckExist(intShow As Integer) As Boolean
Dim adocase As New ADODB.Recordset
   adocase.CursorLocation = adUseClient
   Select Case txtSystem
      Case "P", "CFP", "FCP"
         strExc(0) = "select pa01 from patent where pa01 = '" & txtSystem & "' and pa02 = " & CNULL(txtCode(0)) & " and pa03 = '" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and pa04 = '" & IIf(txtCode(2) = "", "00", txtCode(2)) & "'"
      Case "T", "CFT", "FCT"
         strExc(0) = "select tm01 from trademark where tm01 = '" & txtSystem & "' and tm02 = '" & txtCode(0) & "' and tm03 = '" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and tm04 = '" & IIf(txtCode(2) = "", "00", txtCode(2)) & "'"
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS"
         strExc(0) = "select lc01 from lawcase where lc01 = '" & txtSystem & "' and lc02 = '" & txtCode(0) & "' and lc03 = '" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and lc04 = '" & IIf(txtCode(2) = "", "00", txtCode(2)) & "'"
      Case "LA"
         strExc(0) = "select hc01 from hirecase where hc01 = '" & txtSystem & "' and hc02 = '" & txtCode(0) & "' and hc03 = '" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and hc04 = '" & IIf(txtCode(2) = "", "00", txtCode(2)) & "'"
      Case "TF"
         strExc(0) = "select tm01 from trademark where tm01 = '" & txtSystem & "' and tm02 = '" & txtTFCode(0) & txtTFCode(1) & "' and tm03 = '" & IIf(txtTFCode(2) = "", "0", txtTFCode(2)) & "' and tm04 = '" & IIf(txtTFCode(3) = "", "00", txtTFCode(3)) & "'"
      Case Else
         strExc(0) = "select sp01 from servicepractice where sp01 = '" & txtSystem & "' and sp02 = '" & txtCode(0) & "' and sp03 = '" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and sp04 = '" & IIf(txtCode(2) = "", "00", txtCode(2)) & "'"
   End Select
   intI = 1
   Set adocase = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      If intShow = 1 Then
         If MsgBox(MsgText(36), vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
            CheckExist = True
         Else
            CheckExist = False
            intSaveMode = 1
         End If
      Else
         CheckExist = True
      End If
   Else
      CheckExist = False
   End If
   adocase.Close
End Function

'檢查本所案號是否大於目前流水號
Public Function CheckCaseNo() As Boolean
Dim adocase As New ADODB.Recordset
Dim strSql As String
    strSql = "select au03 from autonumber where au01 = '" & txtSystem & "'"
    intI = 1
    'edit by nickc 2007/02/05 不用 dll 了
    'Set RsTemp = objLawDll.ReadRstMsg(intI, strSQL)
    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
    If intI <> 1 Then
'   adocase.CursorLocation = adUseClient
   'adocase.Open "select au03 from autonumber where au01 = '" & txtSystem & "'", cnnConnection, adOpenStatic, adLockReadOnly
'   adocase.Open strSQL, cnnConnection, adOpenDynamic, adLockReadOnly
'   If adocase.RecordCount = 0 Then
         ShowMsg MsgText(37)
      CheckCaseNo = True
   Else
      If Val(txtCode(0)) > RsTemp.Fields(0).Value Then
         '91.12.22 MODIFY BY SONIA
         'ShowMsg MsgText(37)
         'CheckCaseNo = True
         If txtSystem = "FCT" And txtCode(1) = "T" Then
            CheckCaseNo = False
         Else
            ShowMsg MsgText(37)
            CheckCaseNo = True
         End If
         '91.12.22 END
      Else
         CheckCaseNo = False
      End If
   End If
'   adocase.Close
End Function

' 90.12.19 add by louis
Private Function UpdateCtrlState()
   Dim bShow As Boolean
   
   bShow = False
   Select Case txtSystem
      Case "T", "TF", "CFT", "FCT", "TB", "TC"
         Select Case txtCaseProperty
            Case "501":
               bShow = True
            Case Else:
         End Select
      Case "P", "CFP", "FCP"
         Select Case txtCaseProperty
            Case "701":
               bShow = True
               Me.Label5.Caption = "移轉、讓與申請人："
            'Add By Cheng 2002/01/11
            Case 專利權讓與:
               If Me.txtSystem.Text = "P" Then
                  bShow = True
                  Me.Label5.Caption = "專利權讓與申請人："
               End If
            'Add By Cheng 2002/01/14
            Case 合併
               bShow = True
               Me.Label5.Caption = "合併申請人："
            Case 繼承
               If Me.txtSystem.Text = "FCP" Then
                  bShow = True
                  Me.Label5.Caption = "繼承申請人："
               End If
            Case Else:
         End Select
   End Select
   
   If bShow Then
      fraPatition.Visible = True
      txtPetition.TabStop = True
   Else
      fraPatition.Visible = False
      txtPetition.TabStop = False
      '911112 nick
      txtPetition.Text = ""
   End If
End Function

' 新增案件進度資料
Private Sub OnSaveNewCP()
    Dim strSql As String
    ' 本所案號
    Dim strCP01 As String
    Dim strCP02 As String
    Dim strCP03 As String
    Dim strCP04 As String
    ' 收文日
    Dim strCP05 As String
    ' 總收文號
    Dim strCP09 As String
    ' 案件性質
    Dim strCP10 As String
    ' 案件來源代號 (固定90)
    Dim strCP11 As String
    ' 業務區別
    Dim strCP12 As String
    ' 智權人員代號
    Dim strCP13 As String
    ' 承辦人代號
    Dim strCP14 As String
    ' 91.11.10 ADD BY SONIA
    Dim strCP20 As String
    Dim strCP26 As String
    Dim strCP32 As String
    '91.11.10 END
    ' 發文日
    Dim strCP27 As String
    '
    Dim strCP56 As String
    Dim strSalesNo As String '上個接洽記錄單的智權人員
    Dim StrSQLa As String
    Dim rsA As New ADODB.Recordset

    strCP01 = txtSystem
    strCP02 = txtCode(0)
    strCP03 = txtCode(1)
    strCP03 = strCP03 & String(1 - Len(strCP03), "0")
    strCP04 = txtCode(2)
    strCP04 = strCP04 & String(2 - Len(strCP04), "0")

    strCP05 = "19221111"
    strCP09 = AutoNo("B", 6)
    strCP11 = "90"
    'Modify By Cheng 2003/08/13
'    'Modify By Cheng 2003/02/05
'    '智權人員抓A類最大的智權人員
''    strCP12 = GetSalesArea(strUserNum)
''    strCP13 = strUserNum
    strCP12 = ""
    strCP13 = ""
'    strSQLA = "Select * From CaseProgress Where " & ChgCaseprogress(strCP01 & strCP02 & strCP03 & strCP04) & " And  CP09 <'B' Order By CP05 Desc "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'    '若有資料
'    If rsA.RecordCount > 0 Then
'        Do While Not rsA.EOF
'            '若有智權人員
'            If "" & rsA("CP13").Value <> "" Then
'                strCP12 = GetSalesArea("" & rsA("CP13").Value)
'                strCP13 = "" & rsA("CP13").Value
'                Exit Do
'            End If
'            rsA.MoveNext
'        Loop
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
    strCP13 = PUB_GetAKindSalesNo(strCP01, strCP02, strCP03, strCP04)
    strCP12 = GetSalesArea(strCP13)
    strCP14 = strUserNum
    '91.11.10 ADD BY SONIA
    strCP20 = "N"
    strCP26 = "N"
    strCP32 = "N"
    '91.11.10 END
    strCP27 = "19221111"
    strCP56 = Empty
    If txtPetition.Visible = True And IsEmptyText(txtPetition) = False Then
      strCP56 = txtPetition
    End If
    ' 組成SQL語法
    '91.11.10 MODIFY BY SONIA
    'strSQL = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP11,CP12,CP13,CP14,CP27,CP56) " & _
    '                "VALUES ('" & strCP01 & "','" & strCP02 & "','" & strCP03 & "','" & strCP04 & "'," & _
    '                        strCP05 & ",'" & strCP09 & "','" & txtCaseProperty & "','" & strCP11 & "'," & _
    '                        "'" & strCP12 & "','" & strCP13 & "','" & strCP14 & "'," & strCP27 & "," & DBNullString(strCP56) & ") "
    strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP56) " & _
                    "VALUES ('" & strCP01 & "','" & strCP02 & "','" & strCP03 & "','" & strCP04 & "'," & _
                            strCP05 & ",'" & strCP09 & "','" & txtCaseProperty & "','" & strCP11 & "'," & _
                            "'" & strCP12 & "','" & strCP13 & "','" & strCP14 & "','" & strCP20 & "','" & strCP26 & "'," & strCP27 & ",'" & strCP32 & "'," & DBNullString(strCP56) & ") "
    '91.11.10 END
    cnnConnection.Execute strSql

    ' 顯示所新增的案件號碼
    txtRecieveCode(0) = Mid(strCP09, 1, 1)
    txtRecieveCode(1) = Mid(strCP09, 2, Len(strCP09) - 1)
    lblCaseCode = strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04
    
End Sub

'Modify By Cheng 2003/03/28
'Private Sub OnNextForm()
Private Function OnNextForm() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSK02 As String
   Dim strSK03 As String
   ' 本所案號
   Dim strCP01 As String
   Dim strCP02 As String
   Dim strCP03 As String
   Dim strCP04 As String
       
    'Add By Cheng 2003/03/28
    OnNextForm = True
   '911111 nick
   'strCP01 = txtSystem
   'strCP02 = txtCode(0)
   'strCP03 = txtCode(1)
   'strCP03 = strCP03 & String(1 - Len(strCP03), "0")
   'strCP04 = txtCode(2)
   'strCP04 = strCP04 & String(2 - Len(strCP04), "0")
   strCP01 = txtSystem
   strCP02 = IIf(txtSystem <> "TF", txtCode(0), txtTFCode(0) & txtTFCode(1))
   strCP03 = IIf(txtSystem <> "TF", txtCode(1), txtTFCode(2))
   strCP03 = strCP03 & String(1 - Len(strCP03), "0")
   strCP04 = IIf(txtSystem <> "TF", txtCode(2), txtTFCode(3))
   strCP04 = strCP04 & String(2 - Len(strCP04), "0")
   
   '911018 nick 當修改時，是輸入收文號，並沒有本所案號
   '所以要先 select
   '***** start
   If strCP01 = "" And strCP02 = "" And txtRecieveCode(1).Text <> "" Then
        Dim nickstrsql As String
        Dim nick911018rs As New ADODB.Recordset
        Set nick911018rs = New ADODB.Recordset
        nickstrsql = "select cp01,cp02,cp03,cp04 from caseprogress where cp09='" & strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text & "' "
        
        nick911018rs.CursorLocation = adUseClient
        nick911018rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
        If nick911018rs.RecordCount <> 0 Then
             strCP01 = CheckStr(nick911018rs.Fields(0).Value)
             strCP02 = CheckStr(nick911018rs.Fields(1).Value)
             strCP03 = CheckStr(nick911018rs.Fields(2).Value)
             strCP04 = CheckStr(nick911018rs.Fields(3).Value)
        'Add By Cheng 2003/03/28
        '若無資料
        Else
            MsgBox "無此收文號資料!!!", vbExclamation + vbOKOnly
            OnNextForm = False
            nick911018rs.Close
            Set nick911018rs = Nothing
            Exit Function
        End If
   End If
   '*** end
   
   strSql = "SELECT * FROM SYSTEMKIND " & _
            "WHERE SK01 = '" & strCP01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("SK02")) = False Then
         strSK02 = rsTmp.Fields("SK02")
      End If
      If IsNull(rsTmp.Fields("SK03")) = False Then
         strSK03 = rsTmp.Fields("SK03")
      End If
    'Add By Cheng 2003/03/28
    '若無系統類別
    Else
        MsgBox "此資料無相關系統類別!!!", vbExclamation + vbOKOnly
        OnNextForm = False
        rsTmp.Close
        Set rsTmp = Nothing
        Exit Function
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   Select Case strSK02
      ' 專利
      Case "1":
         Select Case strSK03
            ' 內
            Case "0":
               frm010012_04.SetData strCP01, 0, True
               frm010012_04.SetData strCP02, 1, False
               frm010012_04.SetData strCP03, 2, False
               frm010012_04.SetData strCP04, 3, False
               frm010012_04.SetData txtCaseProperty, 4, False
               frm010012_04.SetData txtPetition, 5, False
               '92.03.27
               frm010012_04.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_04.Show
               frm010012_04.QueryData
            ' 外
            Case "1":
               frm010012_05.SetData strCP01, 0, True
               frm010012_05.SetData strCP02, 1, False
               frm010012_05.SetData strCP03, 2, False
               frm010012_05.SetData strCP04, 3, False
               frm010012_05.SetData txtCaseProperty, 4, False
               frm010012_05.SetData txtPetition, 5, False
               '92.03.27
               frm010012_05.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_05.Show
               frm010012_05.QueryData
            ' 外
            Case "2":
               frm010012_06.SetData strCP01, 0, True
               frm010012_06.SetData strCP02, 1, False
               frm010012_06.SetData strCP03, 2, False
               frm010012_06.SetData strCP04, 3, False
               frm010012_06.SetData txtCaseProperty, 4, False
               frm010012_06.SetData txtPetition, 5, False
               '92.03.27
               frm010012_06.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_06.Show
               frm010012_06.QueryData
         End Select
      ' 商標
      Case "2":
         Select Case strSK03
            ' 內
            Case "0":
               frm010012_01.SetData strCP01, 0, True
               frm010012_01.SetData strCP02, 1, False
               frm010012_01.SetData strCP03, 2, False
               frm010012_01.SetData strCP04, 3, False
               frm010012_01.SetData txtCaseProperty, 4, False
               frm010012_01.SetData txtPetition, 5, False
               '92.03.27
               frm010012_01.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_01.Show
               frm010012_01.QueryData
            ' 外
            Case "1", "2":
               frm010012_02.SetData strCP01, 0, True
               frm010012_02.SetData strCP02, 1, False
               frm010012_02.SetData strCP03, 2, False
               frm010012_02.SetData strCP04, 3, False
               frm010012_02.SetData txtCaseProperty, 4, False
               frm010012_02.SetData txtPetition, 5, False
               '92.03.27
               frm010012_02.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_02.Show
               frm010012_02.QueryData
         End Select
      ' 法務
      Case "3":
         Select Case strSK03
            ' 內
            Case "0":
               frm010012_07.SetData strCP01, 0, True
               frm010012_07.SetData strCP02, 1, False
               frm010012_07.SetData strCP03, 2, False
               frm010012_07.SetData strCP04, 3, False
               frm010012_07.SetData txtCaseProperty, 4, False
               frm010012_07.SetData txtPetition, 5, False
               '92.03.27
               frm010012_07.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_07.Show
               frm010012_07.QueryData
            ' 外
            Case "1", "2":
               frm010012_08.SetData strCP01, 0, True
               frm010012_08.SetData strCP02, 1, False
               frm010012_08.SetData strCP03, 2, False
               frm010012_08.SetData strCP04, 3, False
               frm010012_08.SetData txtCaseProperty, 4, False
               frm010012_08.SetData txtPetition, 5, False
               '92.03.27
               frm010012_08.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_08.Show
               frm010012_08.QueryData
         End Select
      ' 顧問
      Case "4":
         frm010012_07.SetData strCP01, 0, True
         frm010012_07.SetData strCP02, 1, False
         frm010012_07.SetData strCP03, 2, False
         frm010012_07.SetData strCP04, 3, False
         frm010012_07.SetData txtCaseProperty, 4, False
         frm010012_07.SetData txtPetition, 5, False
         '92.03.27
         frm010012_07.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
         
         frm010012_07.Show
         frm010012_07.QueryData
      ' 專利
      Case "5"
         Select Case strSK03
            ' 內
            Case "0":
               frm010012_04.SetData strCP01, 0, True
               frm010012_04.SetData strCP02, 1, False
               frm010012_04.SetData strCP03, 2, False
               frm010012_04.SetData strCP04, 3, False
               frm010012_04.SetData txtCaseProperty, 4, False
               frm010012_04.SetData txtPetition, 5, False
               '92.03.27
               frm010012_04.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_04.Show
               frm010012_04.QueryData
            ' 外
            Case "1":
               frm010012_05.SetData strCP01, 0, True
               frm010012_05.SetData strCP02, 1, False
               frm010012_05.SetData strCP03, 2, False
               frm010012_05.SetData strCP04, 3, False
               frm010012_05.SetData txtCaseProperty, 4, False
               frm010012_05.SetData txtPetition, 5, False
               '92.03.27
               frm010012_05.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_05.Show
               frm010012_05.QueryData
            ' 外
            Case "2":
               frm010012_06.SetData strCP01, 0, True
               frm010012_06.SetData strCP02, 1, False
               frm010012_06.SetData strCP03, 2, False
               frm010012_06.SetData strCP04, 3, False
               frm010012_06.SetData txtCaseProperty, 4, False
               frm010012_06.SetData txtPetition, 5, False
               '92.03.27
               frm010012_06.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_06.Show
               frm010012_06.QueryData
         End Select
      ' 商標
      Case "6":
         Select Case strSK03
            ' 內
            Case "0":
               frm010012_01.SetData strCP01, 0, True
               frm010012_01.SetData strCP02, 1, False
               frm010012_01.SetData strCP03, 2, False
               frm010012_01.SetData strCP04, 3, False
               frm010012_01.SetData txtCaseProperty, 4, False
               frm010012_01.SetData txtPetition, 5, False
               '92.03.27
               frm010012_01.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_01.Show
               frm010012_01.QueryData
            ' 外
            Case "1", "2":
               frm010012_02.SetData strCP01, 0, True
               frm010012_02.SetData strCP02, 1, False
               frm010012_02.SetData strCP03, 2, False
               frm010012_02.SetData strCP04, 3, False
               frm010012_02.SetData txtCaseProperty, 4, False
               frm010012_02.SetData txtPetition, 5, False
               '92.03.27
               frm010012_02.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_02.Show
               frm010012_02.QueryData
         End Select
      ' 法務
      Case "7":
         Select Case strSK03
            ' 內
            Case "0":
               frm010012_07.SetData strCP01, 0, True
               frm010012_07.SetData strCP02, 1, False
               frm010012_07.SetData strCP03, 2, False
               frm010012_07.SetData strCP04, 3, False
               frm010012_07.SetData txtCaseProperty, 4, False
               frm010012_07.SetData txtPetition, 5, False
               '92.03.27
               frm010012_07.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_07.Show
               frm010012_07.QueryData
            ' 外
            Case "1", "2":
               frm010012_08.SetData strCP01, 0, True
               frm010012_08.SetData strCP02, 1, False
               frm010012_08.SetData strCP03, 2, False
               frm010012_08.SetData strCP04, 3, False
               frm010012_08.SetData txtCaseProperty, 4, False
               frm010012_08.SetData txtPetition, 5, False
               '92.03.27
               frm010012_08.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_08.Show
               frm010012_08.QueryData
         End Select
      ' 顧問
      Case "8":
         frm010012_07.SetData strCP01, 0, True
         frm010012_07.SetData strCP02, 1, False
         frm010012_07.SetData strCP03, 2, False
         frm010012_07.SetData strCP04, 3, False
         frm010012_07.SetData txtCaseProperty, 4, False
         frm010012_07.SetData txtPetition, 5, False
         '92.03.27
         frm010012_07.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
         
         frm010012_07.Show
         frm010012_07.QueryData
   End Select
End Function

Public Sub SetData(ByVal strData As String, ByVal nType As Integer, ByVal bClear As Boolean)
   If bClear Then
      txtSystem = Empty
      txtCode(0) = Empty
      txtCode(1) = Empty
      txtCode(2) = Empty
      textCP05 = Empty
      txtCaseProperty = Empty
      lblCasePropertyName = Empty
      txtPetition = Empty
      lblPetitionName = Empty
      fraRecieve.Enabled = False
      fraCode.Visible = True
      fraLastCaseCode.Visible = True
   End If
   Select Case nType
      Case 0:
         ' 91.10.15 modify by louis (固定顯示後面六碼即可, 其它自動顯示)
         'txtRecieveCode(0) = Mid(strData, 1, 1)
         'txtRecieveCode(1) = Mid(strData, 2, Len(strData) - 1)
         txtRecieveCode(1) = Right(strData, 6)
      Case 1:
         lblCaseCode = strData
      Case 2:
         lblCaseCode = lblCaseCode & "-" & strData
      Case 3:
         lblCaseCode = lblCaseCode & "-" & strData
      Case 4:
         lblCaseCode = lblCaseCode & "-" & strData
   End Select
End Sub

'Add By Cheng 2003/09/08
Private Function CheckNewCase(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

CheckNewCase = False
'若無流水號
If strCP02 = "" Then
    CheckNewCase = True
'若有流水號
Else
    StrSQLa = "Select PA01 From Patent Where " & ChgPatent(strCP01 & strCP02 & strCP03 & strCP04)
    StrSQLa = StrSQLa & " Union Select TM01 From Trademark Where " & ChgTradeMark(strCP01 & strCP02 & strCP03 & strCP04)
    StrSQLa = StrSQLa & " Union Select LC01 From Lawcase Where " & ChgLawcase(strCP01 & strCP02 & strCP03 & strCP04)
    StrSQLa = StrSQLa & " Union Select HC01 From Hirecase Where " & ChgHirecase(strCP01 & strCP02 & strCP03 & strCP04)
    StrSQLa = StrSQLa & " Union Select SP01 From Servicepractice Where " & ChgService(strCP01 & strCP02 & strCP03 & strCP04)
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    '若基本檔無資料
    If rsA.RecordCount <= 0 Then
        CheckNewCase = True
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End If

End Function

