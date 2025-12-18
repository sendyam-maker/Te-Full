VERSION 5.00
Begin VB.Form frm880016 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文項目選擇"
   ClientHeight    =   5745
   ClientLeft      =   450
   ClientTop       =   990
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   4680
   Begin VB.TextBox txtSendCnt 
      Height          =   270
      Left            =   1770
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1590
      Width           =   315
   End
   Begin VB.ListBox lstData2 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      ItemData        =   "frm880016.frx":0000
      Left            =   120
      List            =   "frm880016.frx":0007
      Sorted          =   -1  'True
      Style           =   1  '項目包含核取方塊
      TabIndex        =   3
      Top             =   4110
      Width           =   4410
   End
   Begin VB.TextBox txtCP123 
      Height          =   270
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   0
      Top             =   1260
      Width           =   315
   End
   Begin VB.TextBox txtPeriod 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   1155
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   600
      Width           =   1410
   End
   Begin VB.TextBox txtCaseNo 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   1155
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   930
      Width           =   1410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2385
      TabIndex        =   4
      Top             =   75
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3330
      TabIndex        =   5
      Top             =   75
      Width           =   1200
   End
   Begin VB.ListBox lstData 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      ItemData        =   "frm880016.frx":0019
      Left            =   120
      List            =   "frm880016.frx":0020
      Sorted          =   -1  'True
      Style           =   1  '項目包含核取方塊
      TabIndex        =   2
      Top             =   2220
      Width           =   4410
   End
   Begin VB.Label Label5 
      Caption         =   "是否算發文室件數：         (Y:是  N:否)"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1590
      Width           =   3645
   End
   Begin VB.Label Label4 
      Caption         =   "其他-主管機關"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3810
      Width           =   4350
   End
   Begin VB.Label Label3 
      Caption         =   "是否經發文室：         (N:不經發文室)"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1260
      Width           =   3645
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   930
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "主要-主管機關 (只能選擇一個)"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   4350
   End
   Begin VB.Label lblFund 
      Caption         =   "送件時段："
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frm880016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/16 Form2.0已檢查 (無需修改的物件);
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit

Public strCP09 As String        '收文號
Public strCP27 As String        '發文日
Public strCP09s As String       '收文號(回傳)
Public strCP123s As String     '是否經發文室-主管機關(回傳)
Public strCP130s As String     '主管機關名稱(回傳)
Dim strCP01 As String, strCP10 As String, strCPM21 As String
Public bolOK As Boolean       'True: 確定  False: 取消
Public bolIsDefer As Boolean '是否為延期發文 Added by Morgan 2011/11/3
Public bolIsEApp As Boolean '是否電子送件 Added by Morgan 2016/4/29
Public bolIsCaseNum As Boolean '是否算發文室件數 Add by Sindy 2018/8/3


Private Sub cmdOK_Click(Index As Integer)
Dim Cancel As Boolean
   
   '確定
   If Index = 0 Then
      strCP123s = ""
      strCP130s = ""
      
      '檢查資料正確性
      Cancel = False
      txtCP123_Validate Cancel
      If Cancel = True Then
         Exit Sub
      End If
      
      If txtSendCnt.Visible = True Then
         txtSendCnt_Validate Cancel
         If Cancel = True Then
            Exit Sub
         End If
      End If
      'FCT的變更案若經發文室其是否算發文室件數不可空白
      If txtSendCnt.Visible = True And txtCP123.Text <> "N" Then
         If Trim(txtSendCnt.Text) = "" Then
            'Modify By Sindy 2012/8/3
            'MsgBox "FCT變更案若要經發文室則是否算發文室件數不可空白!!!", vbExclamation + vbOKOnly
            MsgBox "FCT" & GetPrjState4(txtCaseNo.Text, strCP10) & "案若要經發文室則是否算發文室件數不可空白!!!", vbExclamation + vbOKOnly
            txtSendCnt.SetFocus
            '2012/8/3 End
            Exit Sub
         End If
      End If
      
      lstData_Validate Cancel
      If Cancel = True Then
         Exit Sub
      End If
      lstData2_Validate Cancel
      If Cancel = True Then
         Exit Sub
      End If
      
      If txtCP123.Text <> "N" Then
         If txtSendCnt.Visible = True Then
            strCP123s = Trim(txtSendCnt.Text)
         Else
            strCP123s = "Y"
         End If
      End If
      'Add By Sindy 2018/8/3 回傳是否算發文室件數
      If txtSendCnt = "Y" Then
         bolIsCaseNum = True
      Else
         bolIsCaseNum = False
      End If
      
      'Modified by Morgan 2016/4/29
      'If txtCP123.Text <> "N" Then
      If txtCP123.Text <> "N" Or bolIsEApp Then
         '主要-主管機關
         For intI = 0 To lstData.ListCount - 1
            If lstData.Selected(intI) = True Then
               If strCP130s = "" Then
                  strCP130s = Trim(lstData.List(intI))
               Else
                  strCP130s = strCP130s & "," & Trim(lstData.List(intI))
               End If
            End If
         Next intI
         '其他-主管機關
         For intI = 0 To lstData2.ListCount - 1
            If lstData2.Selected(intI) = True Then
               If strCP130s = "" Then
                  strCP130s = Trim(lstData2.List(intI))
               Else
                  strCP130s = strCP130s & "," & Trim(lstData2.List(intI))
               End If
            End If
         Next intI
         If strCP130s = "" Then
            MsgBox "請點選主管機關!!!", vbExclamation + vbOKOnly
            Exit Sub
         End If
      End If
      bolOK = True
      
   '回前畫面(取消)
   Else
      bolOK = False
   End If
   Me.Hide
End Sub

Public Function CheckShowList() As Boolean
Dim stSQL As String, strCF10 As String
Dim intIdx As Integer
   
   CheckShowList = False
     
   
   lstData.Clear
   lstData2.Clear
   '未指定發文日時預設當日
   If strCP27 = "" Then
      strCP27 = strSrvDate(1)
   Else
      strCP27 = DBDATE(strCP27)
   End If
   
   txtPeriod = Format(strCP27 - 19110000, "###/##/##")
   
   strCP09s = strCP09
   strCP123s = ""
   strCP130s = ""
   strCP01 = ""
   strCP10 = ""
   strCPM21 = ""
   If strCP09 <> "" Then
      '取得系統別、案件性質
      stSQL = "SELECT * FROM CaseProgress WHERE CP09='" & strCP09 & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
      If intI = 1 Then
         strCP01 = RsTemp("CP01")
         strCP10 = RsTemp("CP10")
         'Modify By Sindy 2012/8/3
         'Me.txtCaseNo = RsTemp("cp01") & "-" & RsTemp("cp02") & IIf(RsTemp("cp03") & RsTemp("cp04") = "000", "", "-" & RsTemp("cp03") & "-" & RsTemp("cp04"))
         Me.txtCaseNo = RsTemp("cp01") & "-" & RsTemp("cp02") & "-" & RsTemp("cp03") & "-" & RsTemp("cp04")
         '2012/8/3 End
      End If
      If strCP01 = "" Or strCP10 = "" Then
         bolOK = True
         Exit Function
      End If
      
      'FCT的變更案才開放須要輸入是否算發文室件數
      'Modify By Sindy 2012/8/3 因商標法修正,FCT的移轉、授轉、再授轉等之申請書已改成與變更案相同,可以一文多案申請
      'If strCP01 = "FCT" And strCP10 = "301" Then
      'Modify By Sindy 2019/9/10 + bolIsEApp = False
      If strCP01 = "FCT" And _
         (strCP10 = "301" Or strCP10 = "501" Or strCP10 = "502" Or strCP10 = "504") And _
         bolIsEApp = False Then
      '2012/8/3 End
         txtSendCnt.Enabled = True
         txtSendCnt.Visible = True
         Label5.Visible = True
      Else
         txtSendCnt.Enabled = False
         txtSendCnt.Visible = False
         Label5.Visible = False
      End If
      
      '取得是否經發文室-主管機關
      stSQL = "SELECT * FROM CasePropertyMap WHERE CPM01='" & strCP01 & "' AND CPM02='" & strCP10 & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
      If intI = 1 Then
         strCPM21 = "" & RsTemp("CPM21")
         
         'Added by Morgan 2011/11/3 若為延期發文時不經發文室的也要詢問
         If strCPM21 = "" And bolIsDefer = True Then
            strCPM21 = "Q"
         End If
         
      End If
      
      'Add by Morgan 2009/10/19
      '專利和商標的訴願固定由智慧局轉呈,不抓申請書的主管機關(經濟部)
      If ((strCP01 = "FCP" Or strCP01 = "P") And strCP10 = "501") Or ((strCP01 = "T" Or strCP01 = "FCT") And strCP10 = "401") Then
         strCF10 = "經濟部智慧財產局"
      Else
      'end 2009/10/19
         strCF10 = ""
         '取得主管機關
         stSQL = "SELECT * FROM CaseFee WHERE CF01='" & strCP01 & "' AND CF02='000' AND CF03='" & strCP10 & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
         If intI = 1 Then
            strCF10 = "" & RsTemp("CF10")
         End If
      End If
      
      'Y:是 (CP130存CF10)
      If strCPM21 = "Y" Then
         
         'Modified by Morgan 2016/4/29 電子送件設不經發文室
         'strCP123s = "Y"
         If Not bolIsEApp Then strCP123s = "Y"
         'end 2016/4/29
         strCP130s = strCF10
         bolOK = True
         Exit Function
         
      'Q:詢問 (彈視窗)
      ElseIf strCPM21 = "Q" Then
         '取得各系統別申請國家為000之主管機關
         stSQL = "SELECT Distinct(CF10) FROM CaseFee WHERE CF01='" & strCP01 & "' AND CF02='000' AND length(CF03)=3 "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               If Not IsNull(RsTemp.Fields(0)) Then
                  intIdx = lstData.ListCount
                  lstData.AddItem RsTemp.Fields(0), intIdx
                  intIdx = lstData2.ListCount
                  lstData2.AddItem RsTemp.Fields(0), intIdx
                  '預設為CF10
                  If Trim(strCF10) <> "" And Trim(RsTemp.Fields(0)) = Trim(strCF10) Then
                     lstData.Selected(intIdx) = True
                  '預設為經濟部智慧財產局
                  ElseIf Trim(strCF10) = "" And Trim(RsTemp.Fields(0)) = "經濟部智慧財產局" Then
                     lstData.Selected(intIdx) = True
                  End If
               End If
               RsTemp.MoveNext
            Loop
            lstData.ListIndex = 0
            lstData2.ListIndex = 0
            If bolIsEApp Then txtCP123 = "N": txtCP123.Enabled = False 'Added by Morgan 2016/4/29
            CheckShowList = True
         Else
            bolOK = True
            Exit Function
         End If
         
      Else
         bolOK = True
         Exit Function
      End If
   End If
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub lstData_Validate(Cancel As Boolean)
Dim intCnt As Integer
   intCnt = 0
   For intI = 0 To lstData.ListCount - 1
      If lstData.Selected(intI) = True Then
         intCnt = intCnt + 1
      End If
   Next intI
   If intCnt > 1 Then
      MsgBox "主要-主管機關只能選擇一個!!!", vbExclamation + vbOKOnly
      lstData.SetFocus
      Cancel = True
      Exit Sub
   ElseIf intCnt = 0 Then
      MsgBox "主要-主管機關不可空白!!!", vbExclamation + vbOKOnly
      lstData.SetFocus
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub lstData2_Validate(Cancel As Boolean)
Dim strMainOrg As String
   
   For intI = 0 To lstData.ListCount - 1
      If lstData.Selected(intI) = True Then
         strMainOrg = Trim(lstData.List(intI))
      End If
   Next intI
   
   For intI = 0 To lstData2.ListCount - 1
      If lstData2.Selected(intI) = True Then
         If strMainOrg = Trim(lstData2.List(intI)) Then
            MsgBox "其他-主管機關不可與主要-主管機關重覆!!!", vbExclamation + vbOKOnly
            lstData2.SetFocus
            Cancel = True
            Exit Sub
         End If
      End If
   Next intI
End Sub

Private Sub txtCP123_GotFocus()
   InverseTextBox txtCP123
End Sub

Private Sub txtCP123_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCP123_Validate(Cancel As Boolean)
   'If txtCP123.Text <> "" Then
      Select Case txtCP123.Text
         Case "", "N"
            'FCT的變更案才開放須要輸入是否算發文室件數
            If txtSendCnt.Visible = True Then
               If txtCP123.Text = "N" Then
                  txtSendCnt.Enabled = False
               Else
                  txtSendCnt.Enabled = True
               End If
            Else
               txtSendCnt.Enabled = False
            End If
         Case Else
            MsgBox "只可輸入空白或N,請重新輸入!!!", vbExclamation + vbOKOnly
            Call txtCP123_GotFocus
            Cancel = True
            Exit Sub
      End Select
   'End If
End Sub

Private Sub txtSendCnt_GotFocus()
   InverseTextBox txtSendCnt
End Sub

Private Sub txtSendCnt_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSendCnt_Validate(Cancel As Boolean)
   If txtSendCnt.Text <> "" Then
      Select Case txtSendCnt.Text
         Case "Y", "N"
         Case Else
            MsgBox "只可輸入Y或N,請重新輸入!!!", vbExclamation + vbOKOnly
            Call txtSendCnt_GotFocus
            Cancel = True
            Exit Sub
      End Select
   End If
End Sub
