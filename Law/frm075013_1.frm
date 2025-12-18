VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm075013_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "出庭費確認維護明細"
   ClientHeight    =   4380
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6528
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6528
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "不領取"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   3456
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   3900
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "領取"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1248
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   3900
      Width           =   1260
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   324
      Left            =   4992
      TabIndex        =   2
      Top             =   96
      Width           =   1020
   End
   Begin VB.TextBox txtDB 
      Height          =   460
      Index           =   1
      Left            =   1656
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   3216
      Width           =   4452
   End
   Begin VB.TextBox txtDB 
      Height          =   460
      Index           =   0
      Left            =   1656
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   2712
      Width           =   4452
   End
   Begin MSForms.Label lblFM2 
      Height          =   276
      Index           =   10
      Left            =   4128
      TabIndex        =   24
      Top             =   2280
      Width           =   996
      Size            =   "1757;487"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   276
      Index           =   9
      Left            =   1104
      TabIndex        =   23
      Top             =   2280
      Width           =   1980
      Size            =   "3492;487"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   276
      Index           =   8
      Left            =   4128
      TabIndex        =   22
      Top             =   1920
      Width           =   1668
      Size            =   "2942;487"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   276
      Index           =   7
      Left            =   1104
      TabIndex        =   21
      Top             =   1920
      Width           =   1980
      Size            =   "3492;487"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "FC代理人："
      Height          =   276
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "當事人1："
      Height          =   276
      Index           =   9
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "律師確認記錄："
      Height          =   276
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   3216
      Width           =   1308
   End
   Begin VB.Label Label1 
      Caption         =   "EMAIL通知記錄："
      Height          =   276
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   2736
      Width           =   1452
   End
   Begin MSForms.Label lblFM2 
      Height          =   276
      Index           =   6
      Left            =   2088
      TabIndex        =   16
      Top             =   1560
      Width           =   4332
      Size            =   "7641;487"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   276
      Index           =   5
      Left            =   1104
      TabIndex        =   15
      Top             =   1560
      Width           =   924
      Size            =   "1630;487"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   276
      Index           =   4
      Left            =   2088
      TabIndex        =   14
      Top             =   1200
      Width           =   4332
      Size            =   "7641;487"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   276
      Index           =   3
      Left            =   1104
      TabIndex        =   13
      Top             =   1200
      Width           =   924
      Size            =   "1630;487"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "出庭費："
      Height          =   276
      Index           =   6
      Left            =   3312
      TabIndex        =   12
      Top             =   2280
      Width           =   756
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   276
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "收文號："
      Height          =   276
      Index           =   4
      Left            =   3312
      TabIndex        =   10
      Top             =   1920
      Width           =   756
   End
   Begin VB.Label Label1 
      Caption         =   "發文日期："
      Height          =   276
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   900
   End
   Begin MSForms.Label lblFM2 
      Height          =   276
      Index           =   2
      Left            =   1104
      TabIndex        =   8
      Top             =   840
      Width           =   5100
      Size            =   "8996;487"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   276
      Index           =   1
      Left            =   4512
      TabIndex        =   7
      Top             =   480
      Width           =   1668
      Size            =   "2942;487"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   276
      Index           =   0
      Left            =   1104
      TabIndex        =   6
      Top             =   480
      Width           =   1668
      Size            =   "2942;487"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   276
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "智慧所案號："
      Height          =   276
      Index           =   1
      Left            =   3312
      TabIndex        =   4
      Top             =   480
      Width           =   1116
   End
   Begin VB.Label Label1 
      Caption         =   "律所案號："
      Height          =   276
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "frm075013_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2024/09/30 (113/11/01上線)
Option Explicit
Dim mPrevForm As Form
Dim m_CP09 As String
Dim m_UserNo As String
Dim m_CL04 As String '確認領取日期
Dim m_CL05 As String '最後不領取日期
Dim intQ As Integer, strQuery As String
Dim rsQuery As New ADODB.Recordset
Dim oObj As Object

Public Sub SetParent(ByVal fm As Form, ByVal CP09 As String, ByVal pUser As String)
   Set mPrevForm = fm
   m_CP09 = CP09
   m_UserNo = pUser
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim strUpdState As String
    
   'Modified by Lydia 2025/04/07 +CL09財務確認律師不領取出庭費日期
   strExc(0) = "select cl06,cl09 from caselawer where cl01='" & lblFM2(8) & "' and cl02='" & m_UserNo & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Val("" & RsTemp.Fields("cl06")) > 0 Then
         MsgBox "財務處已發放，不可變更！", vbExclamation
         Call cmdExit_Click
      End If
      'Added by Lydia 2025/04/07
      If Val("" & RsTemp.Fields("cl09")) > 0 Then
         MsgBox "財務處已確認不領取，不可變更！", vbExclamation
         Call cmdExit_Click
      End If
      'end 2025/04/07
   End If
   
   If Index = 0 Then
      strQuery = "領取"
   Else
      strQuery = "不領取"
   End If
   '今日已做不領取，若要改為領取將會取消今日之不領取記錄，確認要改領取請按是，取消領取請按否 !
   strUpdState = ""
   If m_CL05 = strSrvDate(1) Then
      If Index = 0 Then
         If MsgBox("今日已做不領取，若要改為領取將會取消今日之不領取記錄，" & vbCrLf & "確認要改領取請按是，取消領取請按否 !", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
            strUpdState = "2"
         Else
            Call cmdExit_Click
         End If
      ElseIf Index = 1 Then
         MsgBox "今日已做不領取", vbExclamation + vbOKOnly
         Exit Sub
      End If
   ElseIf m_CL04 = strSrvDate(1) Then
      If Index = 1 Then
         If MsgBox("今日已做領取，若要改為不領取將會取消今日領取記錄，" & vbCrLf & "確認要改不領取請按是，取消請按否 !", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
            strUpdState = "3"
         Else
            Call cmdExit_Click
         End If
      ElseIf Index = 0 Then
         MsgBox "今日已做領取", vbExclamation + vbOKOnly
         Exit Sub
      End If
   Else
      If Val(txtDB(1).Tag) >= 3 Then
         MsgBox "確認記錄超過3次，無法記錄！", vbExclamation, "資料稽核"
         Exit Sub
      End If
      If MsgBox("選擇【" & strQuery & "】出庭費，" & vbCrLf & "選擇「是」繼續存檔，選擇「否」取消存檔。", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
         strUpdState = "1"
      End If
   End If

   If strUpdState <> "" Then
      If strUpdState = "2" Then  '同一日:不領改=>領取
         strUpdState = ChangeWStringToTDateString(strSrvDate(1)) & "不領取"
         txtDB(1) = Replace(Replace(Replace(txtDB(1), strUpdState & ",", ""), "," & strUpdState, ""), strUpdState, "")
         strSql = "Update CaseLawer Set CL05=null,cl08='" & ChgSQL(txtDB(1)) & "' Where CL01='" & m_CP09 & "' and CL02='" & m_UserNo & "'"
         cnnConnection.Execute strSql
      End If
      If strUpdState = "3" Then  '同一日:領改=>不領取
         strUpdState = ChangeWStringToTDateString(strSrvDate(1)) & "領取"
         txtDB(1) = Replace(Replace(Replace(txtDB(1), strUpdState & ",", ""), "," & strUpdState, ""), strUpdState, "")
         strSql = "Update CaseLawer Set CL04=null,cl08='" & ChgSQL(txtDB(1)) & "' Where CL01='" & m_CP09 & "' and CL02='" & m_UserNo & "'"
         cnnConnection.Execute strSql
      End If
      strSql = "Update CaseLawer Set " & IIf(Index = 0, "CL04", "CL05") & "=to_char(sysdate,'yyyymmdd'), cl08=sqldatet(to_char(sysdate,'yyyymmdd'))||'" & strQuery & "'||decode(cl08,null,'',',')||cl08 Where CL01='" & m_CP09 & "' and CL02='" & m_UserNo & "' "
      cnnConnection.Execute strSql
      Call cmdExit_Click
   End If
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   
   If TypeName(mPrevForm) <> "frm075013" Then
      Me.Caption = "出庭費確認維護明細-查詢"
      cmdOK(0).Visible = False
      cmdOK(1).Visible = False
   End If
   
   Call QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If TypeName(mPrevForm) <> "Nothing" Then
      If TypeName(mPrevForm) = "Frmacc42d0" Then
         
      Else
        Call mPrevForm.doQuery(True)
      End If
      mPrevForm.Show
   End If
   
   Set rsQuery = Nothing
   Set frm075013_1 = Nothing
End Sub

Private Sub QueryData()

   For Each oObj In lblFM2
      oObj.Caption = ""
   Next
   For Each oObj In txtDB
      oObj.Text = ""
      oObj.Tag = ""
   Next

   strQuery = "select counting(cl08) cnt, lc01||'-'||lc02||'-'||lc03||'-'||lc04 as caseno," & _
              " decode(c2.cp01,null,null,'TT',null,c2.cp01||'-'||c2.cp02||'-'||c2.cp03||'-'||c2.cp04) as pcase, nvl(lc05,nvl(lc06,lc07)) as casename," & _
              " sqldatet(c1.cp27) as cp27t,c1.cp09,decode(lc15,'000',cpm03,cpm04) as cp10n,cl03,CL04,CL05, cl07,cl08,st01,st02,c1.cp162,los02,los01,c1.cp10," & _
              " lc11 as cno,ltrim(rtrim(nvl(cu04,decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)))) as cname," & _
              " lc22 as fno,ltrim(rtrim(nvl(fa04,decode(fa05,null,fa06,cu05||' '||fa63||' '||fa64||' '||fa65)))) as fname" & _
              " from caselawer,caseprogress c1,lawcase,casepropertymap,staff,lawofficesource, caseprogress c2,customer,fagent" & _
              " where cl01='" & m_CP09 & "' and cl01=c1.cp09(+) and c1.cp159=0 and c1.cp01=lc01(+) and c1.cp02=lc02(+) and c1.cp03=lc03(+) and c1.cp04=lc04(+)" & _
              " and c1.cp01=cpm01(+) and c1.cp10=cpm02(+) and cl02=st01(+) and c1.cp162=los15(+) and los01=c2.cp09(+)" & _
              " and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) and substr(lc22,1,8)=fa01(+) and substr(lc22,9,1)=fa02(+)"
   If m_UserNo <> "" Then
      strQuery = strQuery & " and cl02='" & m_UserNo & "' "
   End If
   
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 1 Then
      lblFM2(0) = "" & rsQuery.Fields("caseno")
      lblFM2(1) = "" & rsQuery.Fields("pcase")
      lblFM2(2) = "" & rsQuery.Fields("casename")
      lblFM2(3) = "" & rsQuery.Fields("cno")
      lblFM2(4) = "" & rsQuery.Fields("cname")
      lblFM2(5) = "" & rsQuery.Fields("fno")
      lblFM2(6) = "" & rsQuery.Fields("fname")
      lblFM2(7) = "" & rsQuery.Fields("cp27t")
      lblFM2(8) = "" & rsQuery.Fields("cp09")
      lblFM2(9) = "" & rsQuery.Fields("cp10n")
      lblFM2(10) = Format("" & rsQuery.Fields("cl03"), "##,##0")
      txtDB(0).Text = "" & rsQuery.Fields("cl07")  'email通知記錄
      txtDB(1).Text = "" & rsQuery.Fields("cl08")  '律師確認記錄
      txtDB(1).Tag = "" & rsQuery.Fields("cnt")
      m_CL04 = "" & rsQuery.Fields("CL04") '確認領取日期
      m_CL05 = "" & rsQuery.Fields("CL05") '最後不領取日期
   End If
End Sub

