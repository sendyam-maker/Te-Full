VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010403_5 
   BorderStyle     =   1  '單線固定
   Caption         =   "部份核駁商品資料異動"
   ClientHeight    =   5736
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8952
   Begin VB.CommandButton cmd 
      Caption         =   "比對"
      Height          =   285
      Index           =   0
      Left            =   6360
      TabIndex        =   12
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   1
      Top             =   30
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   6720
      TabIndex        =   0
      Top             =   30
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   1125
      Left            =   60
      TabIndex        =   4
      Top             =   450
      Width           =   8835
      _ExtentX        =   15579
      _ExtentY        =   1990
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      WordWrap        =   -1  'True
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
      _Band(0).Cols   =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1605
      Left            =   60
      TabIndex        =   8
      Top             =   2070
      Width           =   8835
      _ExtentX        =   15579
      _ExtentY        =   2836
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "中文-1"
      TabPicture(0)   =   "frm02010403_5.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txt1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "中文-2"
      TabPicture(1)   =   "frm02010403_5.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt1(1)"
      Tab(1).ControlCount=   1
      Begin MSForms.TextBox txt1 
         Height          =   1200
         Index           =   0
         Left            =   60
         TabIndex        =   10
         Top             =   360
         Width           =   8715
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "15372;2117"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   1200
         Index           =   1
         Left            =   -74940
         TabIndex        =   9
         Top             =   360
         Width           =   8715
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "15372;2117"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
   End
   Begin MSForms.TextBox txtTemp 
      Height          =   825
      Left            =   720
      TabIndex        =   11
      Top             =   3720
      Width           =   8175
      VariousPropertyBits=   -1467989989
      ForeColor       =   8421504
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "14420;1455"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   885
      Left            =   60
      TabIndex        =   6
      Top             =   4830
      Width           =   8835
      VariousPropertyBits=   -1467989989
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "15584;1561"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "註：比對符號、，；。"
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   7410
      TabIndex        =   15
      Top             =   1620
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "存檔時會將與基本檔不符的類別自動刪除!!!"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   3150
      TabIndex        =   14
      Top             =   240
      Width           =   3420
   End
   Begin VB.Label Label3 
      Caption         =   "比對前資料 :"
      Height          =   465
      Left            =   60
      TabIndex        =   13
      Top             =   3750
      Width           =   615
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註 :"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   4590
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   $"frm02010403_5.frx":0038
      Height          =   405
      Left            =   60
      TabIndex        =   5
      Top             =   1620
      Width           =   6195
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   60
      TabIndex        =   2
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "frm02010403_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/28 Form2.0已修改 textCP64/txt1()/txtTemp/grd1
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

Public UpForm As Form
Public TGKey As String
Public AllClass As String
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
Dim m_TM10 As String 'Added by Lydia 2024/11/21
Dim m_PreCP09 As String 'Added by Lydia 2024/11/21
Dim NowRow As Integer
Dim IsSave As Boolean
Public ChkCht As Boolean
Public ChkEng As Boolean
Public ChkJpn As Boolean
Public PubMsg As String
Public strCP64 As String      '進度備註(回傳)
Public bolOK As Boolean     'True: 確定  False: 取消
Dim tmpUpd() As String  'Added by Lydia 2024/11/21 內商大陸之部份核駁商品異動，記錄各類的比對結果

Private Sub cmd_Click(Index As Integer)
Dim strTMGoods As String, strText As String
Dim strTextBefore As String, strTextAfter As String
Dim i As Long, j As Long, m_i As Long, m_j As Long

PUB_FilterFormText Me 'Add by Morgan 2008/6/20 修正畫面所有含跳行符號的文字框
If IsSave = True Then IsSave = False
With grd1
    .row = NowRow
    .col = 2
    .Text = txt1(0)
    .col = 3
    .Text = txt1(1)
'    .col = 4
'    .Text = txt1(2)
'    .col = 5
'     .Text = txt1(3)
'    .col = 6
'    .Text = txt1(4).Text
'    .col = 7
'     .Text = txt1(5).Text
End With

Select Case Index
Case 0 '比對
   If SSTab1.Tab = 0 Then
      'txt1(0) = Replace(txt1(0), " ", "")
      strTMGoods = txt1(0)
   ElseIf SSTab1.Tab = 1 Then
      'txt1(1) = Replace(txt1(1), " ", "")
      strTMGoods = txt1(1)
   End If
   
'[逐字比對]
'   j = 1
'   strText = ""
'   For i = 1 To Len(Trim(txtTemp)) '異動前
'      If j <= Len(Trim(strTMGoods)) Then '異動後
'         If Mid(Trim(txtTemp), i, 1) <> Mid(Trim(strTMGoods), j, 1) Then
'            strText = strText & Mid(Trim(txtTemp), i, 1)
'         Else
'            j = j + 1
'         End If
'      End If
'   Next i
'   If strText <> "" Then
'      textCP64 = textCP64 & strText
'      'txtTemp = strTMGoods
'   End If
   
'[逐、，；。比對]
   strText = ""
   m_i = 1: m_j = 1
   Do While m_i <= Len(Trim(txtTemp)) Or m_j <= Len(Trim(strTMGoods))
      '異動前
      strTextBefore = ""
      For i = m_i To Len(Trim(txtTemp))
         If Mid(Trim(txtTemp), i, 1) = "、" Or Mid(Trim(txtTemp), i, 1) = "，" Or _
            Mid(Trim(txtTemp), i, 1) = "；" Or Mid(Trim(txtTemp), i, 1) = "。" Then
            m_i = i + 1
            If strTextBefore <> "" Then
               Exit For
            End If
         Else
            strTextBefore = strTextBefore & Mid(Trim(txtTemp), i, 1)
            If i = Len(Trim(txtTemp)) Then
               m_i = i + 1
               Exit For
            End If
         End If
      Next i
      '異動後
      strTextAfter = ""
      For j = m_j To Len(Trim(strTMGoods))
         If Mid(Trim(strTMGoods), j, 1) = "、" Or Mid(Trim(strTMGoods), j, 1) = "，" Or _
            Mid(Trim(strTMGoods), j, 1) = "；" Or Mid(Trim(strTMGoods), j, 1) = "。" Then
            m_j = j + 1
            If strTextAfter <> "" Then
               Exit For
            End If
         Else
            strTextAfter = strTextAfter & Mid(Trim(strTMGoods), j, 1)
            If j = Len(Trim(strTMGoods)) Then
               m_j = j + 1
               Exit For
            End If
         End If
      Next j
      Do While Trim(strTextBefore) <> Trim(strTextAfter)
         strText = strText & strTextBefore & Mid(Trim(txtTemp), i, 1)
         '異動前
         strTextBefore = ""
         For i = m_i To Len(Trim(txtTemp))
            If Mid(Trim(txtTemp), i, 1) = "、" Or Mid(Trim(txtTemp), i, 1) = "，" Or _
               Mid(Trim(txtTemp), i, 1) = "；" Or Mid(Trim(txtTemp), i, 1) = "。" Then
               m_i = i + 1
               If strTextBefore <> "" Then
                  Exit For
               End If
            Else
               strTextBefore = strTextBefore & Mid(Trim(txtTemp), i, 1)
               If i = Len(Trim(txtTemp)) Then
                  m_i = i + 1
                  Exit For
               End If
            End If
         Next i
      Loop
   Loop
   If strText <> "" Then
      textCP64 = textCP64 & strText
      txtTemp = strTMGoods
   End If
   tmpUpd(NowRow) = strText 'Added by Lydia 2024/11/21 內商大陸之部份核駁商品異動，記錄各類的比對結果
Case Else
End Select
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim oStrSQL As String
Dim i930922 As Integer
Dim IsHaveAsk As Boolean
Dim m_copytext As String

Select Case Index
Case 0 '存檔
    With grd1
        For i930922 = 1 To .Rows - 1
            ChgData (i930922)
            'If Trim(txt1(0).Text) = "" And Trim(txt1(1).Text) = "" And Trim(txt1(2).Text) = "" And Trim(txt1(3).Text) = "" And Trim(txt1(4).Text) = "" And Trim(txt1(5).Text) = "" Then
            If Trim(txt1(0).Text) = "" And Trim(txt1(1).Text) = "" Then
                 strTit = "資料檢核"
                 'strMsg = "中文、英文、日文最少輸一種！"
                 strMsg = "商品資料不可空白！"
                 nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                 txt1(0).SetFocus
                 Exit Sub
            End If
            If ChkCht = True And Trim(txt1(0).Text) = "" And Trim(txt1(1).Text) = "" Then
                 strTit = "資料檢核"
                 nResponse = MsgBox(PubMsg, vbOKOnly, strTit)
                 txt1(0).SetFocus
                 Exit Sub
            End If
            'Add by Amy 2021/12/28檢查畫面的 TextBox是否含有Unicode文字
            If PUB_ChkUniText(Me, , True, "TextBox") = False Then
                Exit Sub
            End If

'            If ChkEng = True And Trim(txt1(2).Text) = "" And Trim(txt1(3).Text) = "" Then
'                 strTit = "資料檢核"
'                 nResponse = MsgBox(PubMsg, vbOKOnly, strTit)
'                 txt1(2).SetFocus
'                 Exit Sub
'            End If
'            If ChkJpn = True And Trim(txt1(4).Text) = "" And Trim(txt1(5).Text) = "" Then
'                 strTit = "資料檢核"
'                 nResponse = MsgBox(PubMsg, vbOKOnly, strTit)
'                 txt1(4).SetFocus
'                 Exit Sub
'            End If
            'add by nickc 2006/06/13檢查有無問號
            If InStr(1, txt1(0), "?") <> 0 Then IsHaveAsk = True
            If InStr(1, txt1(1), "?") <> 0 Then IsHaveAsk = True
'            If InStr(1, txt1(2), "?") <> 0 Then IsHaveAsk = True
'            If InStr(1, txt1(3), "?") <> 0 Then IsHaveAsk = True
'            If InStr(1, txt1(4), "?") <> 0 Then IsHaveAsk = True
'            If InStr(1, txt1(5), "?") <> 0 Then IsHaveAsk = True
            
            If CheckLengthIsOK(txt1(0), 4000) = False Then SSTab1.Tab = 0: txt1_GotFocus 0: Exit Sub
            If CheckLengthIsOK(txt1(1), 4000) = False Then SSTab1.Tab = 1: txt1_GotFocus 1: Exit Sub
'            If CheckLengthIsOK(txt1(2), 4000) = False Then SSTab1.Tab = 2: txt1_GotFocus 2: Exit Sub
'            If CheckLengthIsOK(txt1(3), 4000) = False Then SSTab1.Tab = 3: txt1_GotFocus 3: Exit Sub
'            If CheckLengthIsOK(txt1(4), 4000) = False Then SSTab1.Tab = 4: txt1_GotFocus 4: Exit Sub
'            If CheckLengthIsOK(txt1(5), 4000) = False Then SSTab1.Tab = 5: txt1_GotFocus 5: Exit Sub
        Next i930922
        'add by nickc 2006/06/13若有問號，問一下要不要繼續
        If IsHaveAsk = True Then
            If MsgBox("輸入的名稱含有問號，" & vbCrLf & "　　　　若是正常請按　是　繼續" & vbCrLf & "　　　　若不正常請按　否　修正！", vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
    'Save
        Dim IsCanUpdate As Boolean
        Me.Enabled = False
        grd1.Visible = False
        Screen.MousePointer = vbHourglass
        grd1.MousePointer = flexHourglass
        On Error GoTo ShowErr
        cnnConnection.BeginTrans
        
        For i930922 = 1 To .Rows - 1
            ChgData (i930922)
            .col = 1
            oStrSQL = "select * from tmgoods where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg05='" & Trim(.Text) & "' "
            CheckOC3
            AdoRecordSet3.CursorLocation = adUseClient
            AdoRecordSet3.Open oStrSQL, cnnConnection, adOpenStatic, adLockReadOnly
            If AdoRecordSet3.RecordCount <> 0 Then
                IsCanUpdate = False
                If CheckStr(AdoRecordSet3.Fields("tg06")) <> txt1(0) Then
                    IsCanUpdate = True
                End If
                If CheckStr(AdoRecordSet3.Fields("tg15")) <> txt1(1) Then
                    IsCanUpdate = True
                End If
'                If CheckStr(AdoRecordSet3.Fields("tg07")) <> txt1(2) Then
'                    IsCanUpdate = True
'                End If
'                If CheckStr(AdoRecordSet3.Fields("tg16")) <> txt1(3) Then
'                    IsCanUpdate = True
'                End If
'                If CheckStr(AdoRecordSet3.Fields("tg08")) <> txt1(4) Then
'                    IsCanUpdate = True
'                End If
'                If CheckStr(AdoRecordSet3.Fields("tg17")) <> txt1(5) Then
'                    IsCanUpdate = True
'                End If
                
                If IsCanUpdate = True Then
                    'oStrSQL = "update Tmgoods set tg06='" & ChgSQL(txt1(0)) & "',tg15='" & ChgSQL(txt1(1)) & "',tg07='" & ChgSQL(txt1(2)) & "',tg16='" & ChgSQL(txt1(3)) & "',tg08='" & ChgSQL(txt1(4)) & "',tg17='" & ChgSQL(txt1(5)) & "',tg12='" & strUserNum & "',tg13=to_number(to_char(sysdate,'YYYYMMDD')),tg14=to_number(to_char(sysdate,'HH24MI')) where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg05='" & Trim(.Text) & "' "
                    oStrSQL = "update Tmgoods set tg06='" & ChgSQL(txt1(0)) & "',tg15='" & ChgSQL(txt1(1)) & "',tg12='" & strUserNum & "',tg13=to_number(to_char(sysdate,'YYYYMMDD')),tg14=to_number(to_char(sysdate,'HH24MI')) where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg05='" & Trim(.Text) & "' "
                    cnnConnection.Execute oStrSQL
                End If
            Else
                'oStrSQL = "insert into Tmgoods (tg01,tg02,tg03,tg04,tg05,tg06,tg15,tg07,tg16,tg08,tg17,tg09,tg10,tg11) values ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','" & Trim(.Text) & "','" & ChgSQL(txt1(0)) & "','" & ChgSQL(txt1(1)) & "','" & ChgSQL(txt1(2)) & "','" & ChgSQL(txt1(3)) & "','" & ChgSQL(txt1(4)) & "','" & ChgSQL(txt1(5)) & "','" & strUserNum & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MI'))) "
                oStrSQL = "insert into Tmgoods (tg01,tg02,tg03,tg04,tg05,tg06,tg15,tg09,tg10,tg11) values ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','" & Trim(.Text) & "','" & ChgSQL(txt1(0)) & "','" & ChgSQL(txt1(1)) & "','" & strUserNum & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MI'))) "
                cnnConnection.Execute oStrSQL
            End If
            
            'Added by Lydia 2024/11/21 內商大陸之部份核駁商品異動，記錄各類的比對結果; 先用相關收文號做為PK,等到來函收文再轉為C類收文號
            If m_TM10 = "020" Then
               oStrSQL = "update Tmgoods set tg06='" & ChgSQL(tmpUpd(i930922)) & "',tg12='" & strUserNum & "',tg13=to_number(to_char(sysdate,'YYYYMMDD')),tg14=to_number(to_char(sysdate,'HH24MI')) where tg01='" & Mid(m_PreCP09, 1, 3) & "' and tg02='" & Mid(m_PreCP09, 4, 6) & "' and tg03='0' and tg04='00' and tg05='" & Trim(.Text) & "' "
               cnnConnection.Execute oStrSQL, intI
               If intI = 0 Then
                  oStrSQL = "insert into Tmgoods (tg01,tg02,tg03,tg04,tg05,tg06,tg15,tg09,tg10,tg11) values ('" & Mid(m_PreCP09, 1, 3) & "','" & Mid(m_PreCP09, 4, 6) & "','0','00','" & Trim(.Text) & "','" & ChgSQL(tmpUpd(i930922)) & "', null,'" & strUserNum & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MI'))) "
                  cnnConnection.Execute oStrSQL
               End If
            End If
            'end 2024/11/21
        Next i930922
        'add by nickc 2006/06/28 刪除其他不是在畫面上的類別資料 秀玲說的
        oStrSQL = "delete from tmgoods where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg05 not in (" & GetAddStr(AllClass) & ") "
        cnnConnection.Execute oStrSQL
        'Added by Lydia 2024/11/21 內商大陸之部份核駁商品異動，記錄各類的比對結果
        If m_TM10 = "020" Then
           oStrSQL = "delete from tmgoods where tg01='" & Mid(m_PreCP09, 1, 3) & "' and tg02='" & Mid(m_PreCP09, 4, 6) & "' and tg03='0' and tg04='00' and tg05 not in (" & GetAddStr(AllClass) & ") "
           cnnConnection.Execute oStrSQL, intI
        End If
        'end 2024/11/21
        bolOK = True
        strCP64 = textCP64
        frm02010403_4.textCP64 = textCP64
ShowErr:
    If Err.Number = 0 Then
        cnnConnection.CommitTrans
        MsgBox "存檔成功！", vbInformation
        UpForm.ChkTG = True
    Else
        cnnConnection.RollbackTrans
        MsgBox "存檔失敗！", vbExclamation
        UpForm.ChkTG = False
        grd1.MousePointer = flexDefault
        Screen.MousePointer = vbDefault
        grd1.Visible = True
        Me.Enabled = True
        Exit Sub
    End If
    grd1.MousePointer = flexDefault
    Screen.MousePointer = vbDefault
    grd1.Visible = True
    Me.Enabled = True
'    If UpForm.Tag = "frm030001_1" Then
'        UpForm.cmdok(3).BackColor = &H8000000F
'    End If
    End With
    Me.Hide
    UpForm.Show
    Unload Me
Case 1 '回前畫面
    bolOK = False
    strCP64 = ""
    If IsSave = False Then
        If MsgBox("尚未存檔，確定離開？", vbOKCancel, "警告！") = vbOK Then
            Me.Hide
            UpForm.Show
            Unload Me
            Exit Sub
        End If
    Else
        Me.Hide
        UpForm.Show
        Unload Me
        Exit Sub
    End If
Case Else
End Select
End Sub

Private Sub SetDataListWidth()
grd1.Cols = 14
grd1.row = 0
grd1.col = 0: grd1.Text = ""
grd1.ColWidth(0) = 250
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 1: grd1.Text = "商品類別"
grd1.ColWidth(1) = 1000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 2: grd1.Text = "中文-1"
grd1.ColWidth(2) = 2000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 3: grd1.Text = "中文-2"
grd1.ColWidth(3) = 2000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 4: grd1.Text = "英文-1"
grd1.ColWidth(4) = 2000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 5: grd1.Text = "英文-2"
grd1.ColWidth(5) = 2000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 6: grd1.Text = "日文-1"
grd1.ColWidth(6) = 2000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 7: grd1.Text = "日文-2"
grd1.ColWidth(7) = 2000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 8: grd1.Text = "建立人員"
grd1.ColWidth(8) = 1000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 9: grd1.Text = "建立日期"
grd1.ColWidth(9) = 1000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 10: grd1.Text = "建立時間"
grd1.ColWidth(10) = 1000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 11: grd1.Text = "修改人員"
grd1.ColWidth(11) = 1000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 12: grd1.Text = "修改日期"
grd1.ColWidth(12) = 1000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 13: grd1.Text = "修改時間"
grd1.ColWidth(13) = 1000
grd1.CellAlignment = flexAlignCenterCenter
End Sub

Public Sub QueryData()
Screen.MousePointer = vbHourglass
grd1.MousePointer = flexHourglass
grd1.Visible = False
grd1.Clear
grd1.Rows = 2
SetDataListWidth
Dim oStrSQL As String
Dim tmpClass As Variant
Dim i930922 As Integer
Dim i930922_1 As Integer
Dim IsFind As Boolean
m_TM01 = SystemNumber(TGKey, 1)
m_TM02 = SystemNumber(TGKey, 2)
m_TM03 = SystemNumber(TGKey, 3)
m_TM04 = SystemNumber(TGKey, 4)
Me.lbl1.Caption = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
'Added by Lydia 2024/11/21
If Me.Tag <> "" Then
   m_TM10 = Mid(Me.Tag, 1, 3)
   m_PreCP09 = Mid(Me.Tag, 4)
End If
'end 2024/11/21

If AllClass = "" Then
   MsgBox "尚未建類別！", , "錯誤！"
   Me.Hide
   UpForm.ChkTG = False
   UpForm.Show
   Unload Me
   Screen.MousePointer = vbDefault
   Exit Sub
End If

'Modify By Sindy 2010/6/23 讀取TG全部商品類別
'oStrSQL = "select '',tg05,tg06,tg15,tg07,tg16,tg08,tg17,s1.st02,sqldatet(tg10),decode(tg11,null,'',substr(to_char(tg11,'0000'),1,3) ||':'||substr(to_char(tg11,'0000'),4,2)),s2.st02,sqldatet(tg13),decode(tg14,null,'',substr(to_char(tg14,'0000'),1,3)||':'||substr(to_char(tg14,'0000'),4,2)) from tmgoods,staff S1,staff S2 where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg05 in (" & GetAddStr(AllClass) & ") and tg09=s1.st01(+) and tg12=s2.st01(+)  order by tg05 "
'Modified by Lydia 2024/11/21 +依類別比對後的結果Upd01
oStrSQL = "select '',tg05,tg06,tg15,tg07,tg16,tg08,tg17,s1.st02,sqldatet(tg10),decode(tg11,null,'',substr(to_char(tg11,'0000'),1,3) ||':'||substr(to_char(tg11,'0000'),4,2)),s2.st02,sqldatet(tg13),decode(tg14,null,'',substr(to_char(tg14,'0000'),1,3)||':'||substr(to_char(tg14,'0000'),4,2)) from tmgoods,staff S1,staff S2 where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg09=s1.st01(+) and tg12=s2.st01(+)  order by tg05 "
'oStrSQL = "select '',tg05,tg06,tg15,tg07,tg16,tg08,tg17,s1.st02 as tg09n,sqldatet(tg10) as tg10,decode(tg11,null,'',substr(to_char(tg11,'0000'),1,3) ||':'||substr(to_char(tg11,'0000'),4,2)) as tg11," & _
          "s2.st02 as tg12n,sqldatet(tg13) as tg13,decode(tg14,null,'',substr(to_char(tg14,'0000'),1,3)||':'||substr(to_char(tg14,'0000'),4,2)) as tg14,'' as upd01 " & _
          "from tmgoods,staff S1,staff S2 where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg09=s1.st01(+) and tg12=s2.st01(+) " & _
          "order by tg05 "
CheckOC3
With AdoRecordSet3
    .CursorLocation = adUseClient
    .Open oStrSQL, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        'edit by nick 2005/01/19 因為長度大於 999 用 set 到 grid 時，會後面被截掉
        'Set grd1.Recordset = AdoRecordSet3
        .MoveFirst
        ReDim tmpUpd(1 To .RecordCount)  'Added by Lydia 2024/11/21
        Do While Not .EOF
            If grd1.Rows = 2 Then
                grd1.row = grd1.Rows - 1
                grd1.col = 1
                If Trim(grd1.Text) <> "" Then
                    grd1.Rows = grd1.Rows + 1
                    grd1.row = grd1.Rows - 1
                End If
            Else
                grd1.Rows = grd1.Rows + 1
                grd1.row = grd1.Rows - 1
            End If
            grd1.col = 1
            grd1.Text = CheckStr(.Fields(1))
            grd1.col = 2
            grd1.Text = CheckStr(.Fields(2))
            grd1.col = 3
            grd1.Text = CheckStr(.Fields(3))
            grd1.col = 4
            grd1.Text = CheckStr(.Fields(4))
            grd1.col = 5
            grd1.Text = CheckStr(.Fields(5))
            grd1.col = 6
            grd1.Text = CheckStr(.Fields(6))
            grd1.col = 7
            grd1.Text = CheckStr(.Fields(7))
            grd1.col = 8
            grd1.Text = CheckStr(.Fields(8))
            grd1.col = 9
            grd1.Text = CheckStr(.Fields(9))
            grd1.col = 10
            grd1.Text = CheckStr(.Fields(10))
            grd1.col = 5
            grd1.Text = CheckStr(.Fields(5))
            grd1.col = 6
            grd1.Text = CheckStr(.Fields(6))
            grd1.col = 7
            grd1.Text = CheckStr(.Fields(7))
            grd1.col = 8
            grd1.Text = CheckStr(.Fields(8))
            grd1.col = 9
            grd1.Text = CheckStr(.Fields(9))
            grd1.col = 10
            grd1.Text = CheckStr(.Fields(10))
            grd1.col = 11
            grd1.Text = CheckStr(.Fields(11))
            grd1.col = 12
            grd1.Text = CheckStr(.Fields(12))
            grd1.col = 13
            grd1.Text = CheckStr(.Fields(13))
            .MoveNext
        Loop
        ChgData (1)
        SetDataListWidth
        With grd1
        If AllClass <> "" Then
            tmpClass = Split(AllClass, ",")
                For i930922 = 0 To UBound(tmpClass)
                    IsFind = False
                    For i930922_1 = 1 To .Rows - 1
                        .row = i930922_1
                        .col = 1
                        If Trim(.Text) = Trim(tmpClass(i930922)) Then
                            IsFind = True
                            Exit For
                        End If
                    Next i930922_1
                    If Trim(tmpClass(i930922)) <> "" And IsFind = False Then
                        .row = 1
                        .col = 1
                        If Trim(.Text) <> "" Then
                            .Rows = .Rows + 1
                            .row = .Rows - 1
                            .col = 1
                            .Text = Trim(tmpClass(i930922))
                        Else
                            .Text = Trim(tmpClass(i930922))
                        End If
                    End If
                Next i930922
        End If
        UpForm.ChkTG = True
        For i930922_1 = 1 To .Rows - 1
            .row = i930922_1
            If Trim(.TextMatrix(i930922_1, 2)) = "" And Trim(.TextMatrix(i930922_1, 3)) = "" Then
                If ChkCht = True Then
                        UpForm.ChkTG = False
                        Exit For
                End If
                If Trim(.TextMatrix(i930922_1, 4)) = "" And Trim(.TextMatrix(i930922_1, 5)) = "" Then
                    If ChkEng = True Then
                        UpForm.ChkTG = False
                        Exit For
                    End If
                    If Trim(.TextMatrix(i930922_1, 6)) = "" And Trim(.TextMatrix(i930922_1, 7)) = "" Then
                        UpForm.ChkTG = False
                        Exit For
                    End If
                End If
            End If
            
        Next i930922_1
        End With
    Else
        ReDim tmpUpd(1 To 1)  'Added by Lydia 2024/11/21
        UpForm.ChkTG = False
        If AllClass = "" Then
            MsgBox "尚未建類別！", , "錯誤！"
            Me.Hide
            UpForm.Show
            Unload Me
            Exit Sub
        End If
        tmpClass = Split(AllClass, ",")
        With grd1
            For i930922 = 0 To UBound(tmpClass)
                If Trim(tmpClass(i930922)) <> "" Then
                    .row = 1
                    .col = 1
                    If Trim(.Text) <> "" Then
                        .Rows = .Rows + 1
                        .row = .Rows - 1
                        .col = 1
                        .Text = tmpClass(i930922)
                    Else
                        .Text = tmpClass(i930922)
                    End If
                End If
            Next i930922
        End With
    End If
End With
ChgData 1
'預設為第一筆資料
txtTemp.Text = txt1(0)

grd1.Visible = True
grd1.MousePointer = flexDefault
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   SSTab1.Tab = 0
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   IsSave = True
   ChkCht = False
   ChkEng = False
   ChkJpn = False
   PubMsg = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm02010403_5 = Nothing
End Sub

Sub ChgData(oIndex As Integer)
Dim i930922 As Integer
If oIndex = 0 Then Exit Sub
With grd1
    .Visible = False
    NowRow = oIndex
    For i930922 = 0 To .Rows - 1
        .row = i930922
        .col = 0
        .Text = ""
    Next i930922
    .row = oIndex
    .col = 0
    .Text = "☆"
    .col = 2
    txt1(0).Text = Replace(.Text, " ", "")
    .col = 3
    txt1(1).Text = Replace(.Text, " ", "")
    'Add by Sindy 2016/2/25
    If SSTab1.Tab = 0 Then
      txtTemp.Text = txt1(0)
    Else
      txtTemp.Text = txt1(1)
    End If
    '2016/2/25 END
'    .col = 4
'    txt1(2).Text = .Text
'    'add by nickc 2008/03/28 加欄位
'    .col = 5
'    txt1(3).Text = .Text
'    .col = 6
'    txt1(4).Text = .Text
'    .col = 7
'    txt1(5).Text = .Text
    .Visible = True
End With
End Sub

Private Sub Grd1_Click()
ChgData (grd1.MouseRow)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim i As Integer
   For i = 1 To grd1.Rows - 1
      If grd1.TextMatrix(1, 0) = "☆" Then
         ChgData (i)
         Exit For
      End If
   Next i
   If SSTab1.Tab = 0 Then
      txtTemp.Text = txt1(0)
   ElseIf SSTab1.Tab = 1 Then
      txtTemp.Text = txt1(1)
   End If
End Sub

'Private Sub SSTab1_DblClick()
'Dim i As Integer
'   For i = 1 To GRD1.Rows - 1
'      If GRD1.TextMatrix(1, 0) = "☆" Then
'         ChgData (i)
'         Exit For
'      End If
'   Next i
'   If SSTab1.Tab = 0 Then
'      txtTemp.Text = Txt1(0)
'   ElseIf SSTab1.Tab = 1 Then
'      txtTemp.Text = Txt1(1)
'   End If
'End Sub

Private Sub txt1_GotFocus(Index As Integer)
InverseTextBox txt1(Index)
'edit by nickc 2007/09/29 避免因為輸入法切換空白，而自動回前畫面
'If Index = 0 Then OpenIme Else CloseIme
If Index = 0 And txt1(Index).Enabled = True Then OpenIme Else CloseIme
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Cancel = False
   If IsEmptyText(txt1(Index)) = False Then
      If CheckLengthIsOK(txt1(Index), 4000) = False Then
         Cancel = True
         txt1_GotFocus Index
      End If
    End If
End Sub
