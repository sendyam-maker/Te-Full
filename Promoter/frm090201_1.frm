VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料維護"
   ClientHeight    =   3270
   ClientLeft      =   915
   ClientTop       =   1545
   ClientWidth     =   6825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6825
   Begin VB.CommandButton cmdok1 
      Height          =   400
      Index           =   3
      Left            =   3552
      Picture         =   "frm090201_1.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   26
      ToolTipText     =   "最後一筆"
      Top             =   72
      Width           =   495
   End
   Begin VB.CommandButton cmdok1 
      Height          =   400
      Index           =   2
      Left            =   3060
      Picture         =   "frm090201_1.frx":030A
      Style           =   1  '圖片外觀
      TabIndex        =   25
      ToolTipText     =   "下一筆"
      Top             =   72
      Width           =   495
   End
   Begin VB.CommandButton cmdok1 
      Height          =   400
      Index           =   1
      Left            =   2568
      Picture         =   "frm090201_1.frx":0614
      Style           =   1  '圖片外觀
      TabIndex        =   24
      ToolTipText     =   "上一筆"
      Top             =   72
      Width           =   495
   End
   Begin VB.CommandButton cmdok1 
      Height          =   400
      Index           =   0
      Left            =   2076
      Picture         =   "frm090201_1.frx":091E
      Style           =   1  '圖片外觀
      TabIndex        =   23
      ToolTipText     =   "第一筆"
      Top             =   72
      Width           =   495
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "輸入(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   5244
      TabIndex        =   20
      Top             =   70
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   6000
      TabIndex        =   18
      Top             =   70
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "E-mail(&S)"
      Height          =   400
      Index           =   0
      Left            =   4344
      TabIndex        =   17
      Top             =   70
      Width           =   900
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      Caption         =   "預定會稿日："
      Height          =   255
      Left            =   3915
      TabIndex        =   30
      Top             =   570
      Width           =   1245
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   10
      Left            =   5220
      TabIndex        =   29
      Top             =   570
      Width           =   1575
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2778;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblClose 
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
      Height          =   255
      Left            =   3210
      TabIndex        =   28
      Top             =   840
      Width           =   945
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   9
      Left            =   5205
      TabIndex        =   22
      Top             =   825
      Width           =   1575
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2778;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label21 
      Alignment       =   1  '靠右對齊
      Caption         =   "完稿日："
      Height          =   255
      Left            =   4350
      TabIndex        =   21
      Top             =   825
      Width           =   795
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   8
      Left            =   1110
      TabIndex        =   19
      Top             =   2970
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3757;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label19 
      Caption         =   "智權人員："
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2970
      Width           =   1065
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   7
      Left            =   1110
      TabIndex        =   15
      Top             =   2701
      Width           =   2415
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4260;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      Caption         =   "法定期限："
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2701
      Width           =   930
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   6
      Left            =   1110
      TabIndex        =   13
      Top             =   2433
      Width           =   2325
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4101;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label15 
      Caption         =   "本所期限："
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2433
      Width           =   945
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   5
      Left            =   1110
      TabIndex        =   11
      Top             =   2165
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3228;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      Caption         =   "核稿人："
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2165
      Width           =   765
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   4
      Left            =   1110
      TabIndex        =   9
      Top             =   1897
      Width           =   5775
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "10186;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      Caption         =   "案件性質："
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1897
      Width           =   960
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   3
      Left            =   1485
      TabIndex        =   7
      Top             =   1629
      Width           =   5250
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "9260;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      Caption         =   "案件日文名稱："
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1629
      Width           =   1335
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   2
      Left            =   1485
      TabIndex        =   5
      Top             =   1361
      Width           =   5250
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "9260;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "案件英文名稱："
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1361
      Width           =   1335
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   1485
      TabIndex        =   3
      Top             =   1093
      Width           =   5250
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "9260;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "案件中文名稱："
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1093
      Width           =   1335
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   1110
      TabIndex        =   1
      Top             =   825
      Width           =   2160
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3810;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   825
      Width           =   945
   End
   Begin VB.Label Label6 
      Caption         =   "此案已屆預定會稿日而尚未會稿!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3555
      TabIndex        =   31
      Top             =   2280
      Width           =   3210
   End
   Begin VB.Label Label2 
      Caption         =   "此案件已完稿超過2天, 未鍵入會稿日,              請E-MAIL 通知核稿人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   3540
      TabIndex        =   27
      Top             =   2190
      Width           =   3210
   End
End
Attribute VB_Name = "frm090201_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo by Lydia 2021/12/28 改成Form2.0 ; lbl1(index)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim pemain As New ADODB.Recordset, s As Integer
Public UserStaff As String
Dim CASENAME(3) As String, i As Integer, j As Integer
'Add By Cheng 2003/06/12
Dim m_blnFirstShow As Boolean '是否第一次顯示
'Add By Cheng 2003/05/08
Dim dblDeadLine As Double
    
Private Sub cmdOK_Click(Index As Integer)
On Error GoTo ErrProcess
Select Case Index
   Case 0 'E-Mail
      If Len(Trim(CheckStr(pemain.Fields(11)))) = 0 Then
            s = MsgBox("核稿人空白!!", , "Mail 收件人錯誤!!")
            cmdok(0).SetFocus
            Exit Sub
      Else
            If Index = 0 Then
            '直接寄mail
               Screen.MousePointer = vbHourglass
               Me.Enabled = False
               'edit by nick 2004/11/11 改通用由 frm880005 發
               'Added by Lydia 2022/05/30
               If "" & pemain.Fields("cp09") <> "" Then
                   PUB_SendMail strUserNum, CheckStr(pemain.Fields(11)), pemain.Fields("cp09"), "催函", vbCrLf + "本所案號： " + lbl1(0) + vbCrLf + "完稿日： " + lbl1(9) + vbCrLf + "案件名稱： " + lbl1(1).Caption + "," + lbl1(2).Caption + "," + lbl1(3).Caption + vbCrLf + "案件性質： " + lbl1(4).Caption + vbCrLf + "本所期限： " + lbl1(6).Caption + vbCrLf + "法定期限： " + lbl1(7).Caption + vbCrLf + "智權人員： " + lbl1(8).Caption + vbCrLf + vbCrLf + "核稿天數超過 請儘速核稿!", ""
               Else
               'end 2022/05/25
                   PUB_SendMail strUserNum, CheckStr(pemain.Fields(11)), "", "催函", vbCrLf + "本所案號： " + lbl1(0) + vbCrLf + "完稿日： " + lbl1(9) + vbCrLf + "案件名稱： " + lbl1(1).Caption + "," + lbl1(2).Caption + "," + lbl1(3).Caption + vbCrLf + "案件性質： " + lbl1(4).Caption + vbCrLf + "本所期限： " + lbl1(6).Caption + vbCrLf + "法定期限： " + lbl1(7).Caption + vbCrLf + "智權人員： " + lbl1(8).Caption + vbCrLf + vbCrLf + "核稿天數超過 請儘速核稿!", ""
               End If 'Added by Lydia 2022/05/30
'               mdiMain.MAPISession1.LogonUI = False
'               mdiMain.MAPISession1.UserName = strUserNum
'               mdiMain.MAPISession1.SignOn
'               mdiMain.MAPIMessages1.SessionID = mdiMain.MAPISession1.SessionID
'               mdiMain.MAPIMessages1.MsgIndex = -1
'               mdiMain.MAPIMessages1.Compose
'               mdiMain.MAPIMessages1.MsgSubject = "催函"
'               mdiMain.MAPIMessages1.MsgNoteText = vbCrLf + "本所案號： " + lbl1(0) + vbCrLf + "完稿日： " + lbl1(9) + vbCrLf + "案件名稱： " + lbl1(1).Caption + "," + lbl1(2).Caption + "," + lbl1(3).Caption + vbCrLf + "案件性質： " + lbl1(4).Caption + vbCrLf + "本所期限： " + lbl1(6).Caption + vbCrLf + "法定期限： " + lbl1(7).Caption + vbCrLf + "智權人員： " + lbl1(8).Caption + vbCrLf + vbCrLf + "核稿天數超過 請儘速核稿!"
'               'mdiMain.MAPIMessages1
'               'mdimain.MAPIMessages1
'               mdiMain.MAPIMessages1.RecipIndex = 0
'               mdiMain.MAPIMessages1.RecipDisplayName = CheckStr(pemain.Fields(11))
'               'mdiMain.MAPIMessages1.RecipDisplayName = "74001"
'               'mdiMain.MAPIMessages1.RecipAddress = CheckStr(pemain.Fields(11)) & "@taie.com.tw"
'               'mdiMain.MAPIMessages1.AttachmentName = CheckStr(pemain.Fields(11))
'               mdiMain.MAPIMessages1.ResolveName
'               mdiMain.MAPIMessages1.Send
'               mdiMain.MAPISession1.SignOff
               Me.Enabled = True
               Screen.MousePointer = vbDefault
               s = MsgBox("寄件成功!!")
               pemain.MoveNext
               If pemain.EOF = True Then
                  frm090201_2.Show
                  frm090201_2.SSTab1.Tab = 0
                  Unload Me
                  Exit Sub
               End If
               pemain.MovePrevious
            End If
       End If
      'frm880005.txtEmail(2).Text = vbCrLf + "本所案號： " + lbl1(0) + vbCrLf + "完稿日： " + lbl1(9) + vbCrLf + "案件名稱： " + lbl1(1).Caption + "," + lbl1(2).Caption + "," + lbl1(3).Caption + vbCrLf + "案件性質： " + lbl1(4).Caption + vbCrLf + "本所期限： " + lbl1(6).Caption + vbCrLf + "法定期限： " + lbl1(7).Caption + vbCrLf + "智權人員： " + lbl1(8).Caption + vbCrLf + vbCrLf + "請儘速核稿!"
      'frm090201_1.Hide
      'frm880005.Show
   Case 1 '輸入
      frm090201_1.Hide
      frm090201_2.Show
      frm090201_2.SSTab1.Tab = 1
      frm090201_2.Process (CheckStr(pemain.Fields(10)))
      Unload Me
   Case 2 '結束
      frm090201_1.Hide
      frm090201_2.Show
      frm090201_2.SSTab1.Tab = 0
      Unload Me
End Select
Exit Sub
ErrProcess:
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   s = MsgBox("寄件失敗!! 或郵件無相對應的 Profile !!")
End Sub

Private Sub cmdok1_Click(Index As Integer)
Select Case Index
Case 0
     pemain.MoveFirst
Case 1
     pemain.MovePrevious
     If pemain.BOF = True Then
         s = MsgBox("已經是第一筆")
         pemain.MoveFirst
     End If
Case 2
     pemain.MoveNext
     If pemain.EOF = True Then
         s = MsgBox("已經是最後一筆")
         pemain.MoveLast
     End If
Case 3
     pemain.MoveLast
Case Else
End Select

For i = 0 To 9
  lbl1(i) = CheckStr(pemain.Fields(i))
Next i

   If pemain.Fields("ep28") > 0 Then
      lbl1(10) = ChangeWStringToTDateString(pemain.Fields("ep28"))
   Else
      lbl1(10) = ""
   End If
   If pemain.Fields("ep09") > 0 And pemain.Fields("ep09") < dblDeadLine Then
      Label2.Visible = True
      Label6.Visible = False
   Else
      Label2.Visible = False
      Label6.Visible = True
   End If

End Sub

Private Sub Form_Activate()
    
    If m_blnFirstShow = False Then Exit Sub
    If m_blnFirstShow = True Then m_blnFirstShow = False
    Screen.MousePointer = vbHourglass
    Me.Hide
    If pemain.State = adStateOpen Then pemain.Close
    strExc(0) = "SELECT ST01 FROM STAFF WHERE ST01='" & strUserNum & "' AND ST04='1' "
    pemain.CursorLocation = adUseClient
    pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
    If pemain.BOF And pemain.EOF Then MsgBox "無此LOGIN人員之資料", vbInformation: Unload Me: Exit Sub
    UserStaff = strUserNum
    pemain.Close
    'Add By Cheng 2003/05/08
    dblDeadLine = PUB_GetWorkDay(-2)
   'Modify By Cheng 2002/04/29
   '若已閉卷, 則在本所案號後加"*"號
    'Modify By Cheng 2003/05/08
'                            strExc(0) = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,PA05,PA06,PA07,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10),nvl(A.ST02,EP04)," & SQLDate("CP06") & "," & SQLDate("CP07") & ",NVL(B.ST02,CP14)," & SQLDate("EP09") & ",cp09,ep04,PA57 FROM CASEPROGRESS,PATENT,ENGINEERPROGRESS,CASEPROPERTYMAP,STAFF A,STAFF B WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and ep02=CP09(+) AND EP09 IS NOT NULL   AND (EP07 IS NULL OR EP07=0) AND (CP57 IS NULL OR CP57=0) and cp01=cpm01(+) and cp10=cpm02(+) AND EP04=A.ST01(+) AND CP13=B.ST01(+) and CP26 IS NULL  and ep05='" & strUserNum & "' AND CP27 IS NULL AND CP01 IN (" & SQLGrpStr("", 1) & ") "
'    strExc(0) = strExc(0) & " UNION all  SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,TM05,TM06,TM07,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),nvl(A.ST02,EP04)," & SQLDate("CP06") & "," & SQLDate("CP07") & ",NVL(B.ST02,CP13)," & SQLDate("EP09") & ",cp09,ep04,TM29 FROM CASEPROGRESS,TRADEMARK,ENGINEERPROGRESS,CASEPROPERTYMAP,STAFF A,STAFF B WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) and ep02=CP09(+) AND EP09 IS NOT NULL   AND (EP07 IS NULL OR EP07=0) AND (CP57 IS NULL OR CP57=0) and cp01=cpm01(+) and cp10=cpm02(+) AND EP04=A.ST01(+) AND CP13=B.ST01(+) and CP26 IS NULL  and ep05='" & strUserNum & "' AND CP27 IS NULL AND CP01 IN (" & SQLGrpStr("", 2) & ") "
'    strExc(0) = strExc(0) & " UNION all  SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,LC05,LC06,LC07,NVL(DECODE(LC15,'000',CPM03,CPM04),CP10),nvl(A.ST02,EP04)," & SQLDate("CP06") & "," & SQLDate("CP07") & ",NVL(B.ST02,CP13)," & SQLDate("EP09") & ",cp09,ep04,LC08 FROM CASEPROGRESS,LAWCASE,ENGINEERPROGRESS,CASEPROPERTYMAP,STAFF A,STAFF B WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and ep02=CP09(+) AND EP09 IS NOT NULL   AND (EP07 IS NULL OR EP07=0) AND (CP57 IS NULL OR CP57=0) and cp01=cpm01(+) and cp10=cpm02(+) AND EP04=A.ST01(+) AND CP13=B.ST01(+) and CP26 IS NULL  and ep05='" & strUserNum & "' AND CP27 IS NULL AND CP01 IN (" & SQLGrpStr("", 3) & ") "
'    strExc(0) = strExc(0) & " UNION all  SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,HC06,' ',' ',nvl(CPM03,cp10),nvl(A.ST02,EP04)," & SQLDate("CP06") & "," & SQLDate("CP07") & ",NVL(B.ST02,CP13)," & SQLDate("EP09") & ",cp09,ep04,HC09 FROM CASEPROGRESS,HIRECASE,ENGINEERPROGRESS,CASEPROPERTYMAP,STAFF A, STAFF B WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) and ep02=CP09(+) AND EP09 IS NOT NULL   AND (EP07 IS NULL OR EP07=0) AND (CP57 IS NULL OR CP57=0) and cp01=cpm01(+) and cp10=cpm02(+) AND EP04=A.ST01(+) AND CP13=B.ST01(+) and CP26 IS NULL  and ep05='" & strUserNum & "' AND CP27 IS NULL AND CP01 IN (" & SQLGrpStr("", 4) & ") "
'    strExc(0) = strExc(0) & " UNION all  SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,SP05,SP06,SP07,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),nvl(A.ST02,EP04)," & SQLDate("CP06") & "," & SQLDate("CP07") & ",NVL(B.ST02,CP13)," & SQLDate("EP09") & ",cp09,ep04,SP15 FROM CASEPROGRESS,SERVICEPRACTICE,ENGINEERPROGRESS,CASEPROPERTYMAP,STAFF A,STAFF B WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) and ep02=CP09(+) AND EP09 IS NOT NULL   AND (EP07 IS NULL OR EP07=0) AND (CP57 IS NULL OR CP57=0) and cp01=cpm01(+) and cp10=cpm02(+) AND EP04=A.ST01(+) AND CP13=B.ST01(+) and CP26 IS NULL  and ep05='" & strUserNum & "' AND CP27 IS NULL AND CP01 IN (" & SQLGrpStr("", 5) & ") "
'Modify by Morgan 2009/7/13 +已屆預定會稿日未會稿的案件(含當天),EP09 IS NOT NULL 的條件拿掉,有<dblDeadLine的條件就好
    'Modify By Sindy 2016/9/5 and cp57 is null and cp27 is null ==> and cp158=0 and cp159=0
                            strExc(0) = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,PA05,PA06,PA07,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10),nvl(A.ST02,EP04)," & SQLDate("CP06") & "," & SQLDate("CP07") & ",NVL(B.ST02,CP14)," & SQLDate("EP09") & ",cp09,ep04,PA57,ep28,ep09 FROM CASEPROGRESS,PATENT,ENGINEERPROGRESS,CASEPROPERTYMAP,STAFF A,STAFF B WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and ep02=CP09(+) AND (EP07 IS NULL OR EP07=0) and cp01=cpm01(+) and cp10=cpm02(+) AND EP04=A.ST01(+) AND CP13=B.ST01(+) and CP26 IS NULL  and ep05='" & strUserNum & "' and cp158=0 and cp159=0 AND CP01 IN (" & SQLGrpStr("", 1) & ") And ( EP09 <" & dblDeadLine & " or EP28<=" & strSrvDate(1) & ")"
    strExc(0) = strExc(0) & " UNION all  SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,TM05,TM06,TM07,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),nvl(A.ST02,EP04)," & SQLDate("CP06") & "," & SQLDate("CP07") & ",NVL(B.ST02,CP13)," & SQLDate("EP09") & ",cp09,ep04,TM29,ep28,ep09 FROM CASEPROGRESS,TRADEMARK,ENGINEERPROGRESS,CASEPROPERTYMAP,STAFF A,STAFF B WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) and ep02=CP09(+) AND (EP07 IS NULL OR EP07=0) and cp01=cpm01(+) and cp10=cpm02(+) AND EP04=A.ST01(+) AND CP13=B.ST01(+) and CP26 IS NULL  and ep05='" & strUserNum & "' and cp158=0 and cp159=0 AND CP01 IN (" & SQLGrpStr("", 2) & ") And ( EP09 <" & dblDeadLine & " or EP28<=" & strSrvDate(1) & ")"
    strExc(0) = strExc(0) & " UNION all  SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,LC05,LC06,LC07,NVL(DECODE(LC15,'000',CPM03,CPM04),CP10),nvl(A.ST02,EP04)," & SQLDate("CP06") & "," & SQLDate("CP07") & ",NVL(B.ST02,CP13)," & SQLDate("EP09") & ",cp09,ep04,LC08,ep28,ep09 FROM CASEPROGRESS,LAWCASE,ENGINEERPROGRESS,CASEPROPERTYMAP,STAFF A,STAFF B WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and ep02=CP09(+) AND (EP07 IS NULL OR EP07=0) and cp01=cpm01(+) and cp10=cpm02(+) AND EP04=A.ST01(+) AND CP13=B.ST01(+) and CP26 IS NULL  and ep05='" & strUserNum & "' and cp158=0 and cp159=0 AND CP01 IN (" & SQLGrpStr("", 3) & ") And ( EP09 <" & dblDeadLine & " or EP28<=" & strSrvDate(1) & ")"
    strExc(0) = strExc(0) & " UNION all  SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,HC06,' ',' ',nvl(CPM03,cp10),nvl(A.ST02,EP04)," & SQLDate("CP06") & "," & SQLDate("CP07") & ",NVL(B.ST02,CP13)," & SQLDate("EP09") & ",cp09,ep04,HC09,ep28,ep09 FROM CASEPROGRESS,HIRECASE,ENGINEERPROGRESS,CASEPROPERTYMAP,STAFF A, STAFF B WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) and ep02=CP09(+) AND (EP07 IS NULL OR EP07=0) and cp01=cpm01(+) and cp10=cpm02(+) AND EP04=A.ST01(+) AND CP13=B.ST01(+) and CP26 IS NULL and ep05='" & strUserNum & "' and cp158=0 and cp159=0 AND CP01 IN (" & SQLGrpStr("", 4) & ") And ( EP09 <" & dblDeadLine & " or EP28<=" & strSrvDate(1) & ")"
    strExc(0) = strExc(0) & " UNION all  SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,SP05,SP06,SP07,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),nvl(A.ST02,EP04)," & SQLDate("CP06") & "," & SQLDate("CP07") & ",NVL(B.ST02,CP13)," & SQLDate("EP09") & ",cp09,ep04,SP15,ep28,ep09 FROM CASEPROGRESS,SERVICEPRACTICE,ENGINEERPROGRESS,CASEPROPERTYMAP,STAFF A,STAFF B WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) and ep02=CP09(+) AND (EP07 IS NULL OR EP07=0) and cp01=cpm01(+) and cp10=cpm02(+) AND EP04=A.ST01(+) AND CP13=B.ST01(+) and CP26 IS NULL  and ep05='" & strUserNum & "' and cp158=0 and cp159=0 AND CP01 IN (" & SQLGrpStr("", 5) & ") And ( EP09 <" & dblDeadLine & " or EP28<=" & strSrvDate(1) & ")"
   'StrSQL6 = ""
   'StrSQL6 = StrSQL6 + " and CP26 IS NULL  and ep05='" & strUserNum & "' AND (CP27 IS NULL OR SUBSTR(CP27,1,6)=" & Mid(GetTodayDate, 1, 6) & ") "

   'CheckOC
   '             strSQL = "SELECT nvl(S1.ST02,ep05),EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 1) & ") "
   'strSQL = strSQL + " UNION all  SELECT nvl(S1.ST02,ep05),EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 2) & ") "
   'strSQL = strSQL + " UNION all  SELECT nvl(S1.ST02,ep05),ep01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 3) & ") "
   'strSQL = strSQL + " UNION all  SELECT nvl(S1.ST02,ep05),EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,HC06,CP26,'',nvl(CPM03,cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
   'strSQL = strSQL + " UNION all  SELECT nvl(S1.ST02,ep05),EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 5) & ") "
   'strSQL = strSQL + " ORDER BY 1,4 "
                
    'strExc(0) = "select distinct cp01||cp02||cp03||cp04 as a,pa05 as b,pa06 as c,pa07 as d,'' as e,'' as f,'' as g,'' as h,'' as i,'' as j from caseprogress,patent where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa01='P' "
    pemain.CursorLocation = adUseClient
    pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
    'pemain.Open strExc(0), cnnconnetion, adOpenStatic, adLockReadOnly
    If pemain.RecordCount = 0 Then
        'MsgBox "資料庫內無資料"
        
        Screen.MousePointer = vbHourglass
        frm090201_2.Show
        If frm090201_2.TextOk = False Then
            Unload frm090201_2
        End If
        frm090201_2.TextOk = False
        Screen.MousePointer = vbDefault
        Unload Me
        Exit Sub
    End If
    'If pemain.RecordCount = 0 Then frm090201_2.Show: Unload Me: Screen.MousePointer = vbDefault: Exit Sub
    'If pemain.BOF And pemain.EOF Then MsgBox "資料庫內無資料": Unload Me: Screen.MousePointer = vbHourglass: frm090201_2.Show: If frm090201_2.TextOk = False Then Unload frm090201_2: Screen.MousePointer = vbDefault: Exit Sub

    pemain.MoveFirst
    
    For i = 0 To 9
      lbl1(i) = CheckStr(pemain.Fields(i))
    Next i
   'Add by Morgan 2009/7/13
   If pemain.Fields("ep28") > 0 Then
      lbl1(10) = ChangeWStringToTDateString(pemain.Fields("ep28"))
   Else
      lbl1(10) = ""
   End If
   If pemain.Fields("ep09") > 0 And pemain.Fields("ep09") < dblDeadLine Then
      Label2.Visible = True
      Label6.Visible = False
   Else
      Label2.Visible = False
      Label6.Visible = True
   End If
   'Add By Cheng 2002/04/29
   If IsNull(pemain.Fields(12).Value) Then
      Me.lblClose.Caption = ""
   Else
      Me.lblClose.Caption = "已閉卷"
   End If
   
    Screen.MousePointer = vbDefault
    Me.Show
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    'Add By Cheng 2002/04/29
    Me.lblClose.Caption = ""
    'Add By Cheng 2003/06/12
    m_blnFirstShow = True
    
    'Added by Lydia 2021/12/28
    Dim oLbl As Object
    For Each oLbl In lbl1
        oLbl.Caption = ""
    Next
    'end 2021/12/28
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm090201_1 = Nothing
End Sub

