VERSION 5.00
Begin VB.Form frm880020 
   BorderStyle     =   1  '單線固定
   Caption         =   "新穎性優惠期公開事實"
   ClientHeight    =   3576
   ClientLeft      =   456
   ClientTop       =   996
   ClientWidth     =   2964
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3576
   ScaleWidth      =   2964
   Begin VB.CommandButton cmdMove 
      Caption         =   "清除(&C)"
      Height          =   400
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   2640
      Width           =   912
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "刪除(&D)"
      Height          =   400
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   2040
      Width           =   912
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "新增(&A)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   1440
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      Top             =   120
      Width           =   1200
   End
   Begin VB.TextBox txtDt 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
   Begin VB.ListBox lstDate 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1668
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "公開日："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   845
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "公開日"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1155
   End
End
Attribute VB_Name = "frm880020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/3 改成Form2.0 (無)
'Add by Lydia 2015/02/02 輸入新穎性優惠期公開事實 (多筆)
Option Explicit
Public strFPD01 As String, strFPD02 As String, strFPD03 As String, strFPD04 As String '案號
Public strLimit1 As String, strLimit2 As String '所限，法限
Public strNation As String, strPA140 As String
Public strPA08 As String 'Added by Morgan 2018/3/16
Public strPA10 As String 'Added by Morgan 2022/8/19
Public m_PrevF As Form
Public m_dbCheck As Boolean, m_bolDouble As Boolean  'Added by Lydia 2015/02/25
Dim m_bolAdd As Boolean 'Added by Lydia 2015/03/16
Dim m_bolFMP As Boolean 'Added by Lydia 2025/09/25

Private Sub cmdok_Click(Index As Integer)
   Dim i As Integer, strTmp As String
   Dim Cancel As Boolean

   'Modified by Lydia 2015/02/25  發文DoubleCheck
   If m_dbCheck = True Then
     CheckRecExist
     'Added by Lydia 2015/03/16 若之前已輸入新穎性優惠期,會無記錄
     If m_bolAdd = True And Not IsNull(strPA140) Then
        Unload Me
     End If
     'end 2015/03/16
     Exit Sub
   End If


   If Index = 0 Then '確定
   
'Removed by Morgan 2018/3/16 應該可以不必考慮尚未新增的日期,且也要能允許清除日期
'    If txtDt <> "" Then
'       txtDt_Validate Cancel
'       If Cancel = True Then
'          Exit Sub
'       End If
'    Else
'       If lstDate.ListCount = 0 Then
'          Unload Me
'          Exit Sub
'       End If
'    End If
'end 2018/3/16
    
On Error GoTo ErrHnd
   
      cnnConnection.BeginTrans
      
      strTmp = "delete from FavPridate where FPD01='" & strFPD01 & "' and FPD02='" & strFPD02 & "' and FPD03='" & strFPD03 & "' and FPD04='" & strFPD04 & "'"
      cnnConnection.Execute strTmp
      
      If lstDate.ListCount = 0 Then
         'Modified by Morgan 2018/3/16
         'strPA140 = txtDt
         'strTmp = "insert into FavPridate values ('" & strFPD01 & "','" & strFPD02 & "','" & strFPD03 & "','" & strFPD04 & "'," & Val(txtDt) + 19110000 & ") "
         'cnnConnection.Execute strTmp
         strPA140 = ""
         'end 2018/3/16
      Else
         strPA140 = "" '重新判斷
         For i = 0 To lstDate.ListCount - 1
            If Trim(lstDate.List(i)) <> "" Then
                strTmp = "insert into FavPridate values ('" & strFPD01 & "','" & strFPD02 & "','" & strFPD03 & "','" & strFPD04 & "'," & Val(lstDate.List(i)) + 19110000 & ") "
                cnnConnection.Execute strTmp
                If Val(lstDate.List(i)) < Val(strPA140) Or strPA140 = "" Then
                   strPA140 = LTrim(RTrim(lstDate.List(i)))
                End If
            End If
         Next i
      End If
      
      cnnConnection.CommitTrans

'Removed by Morgan 2018/3/16 回前畫面應該不必動作
'   Else
'     If lstDate.ListCount > 0 And Trim(txtDt) = "" And Len(strPA140) > 0 Then
'        If Len(strPA140) = 8 Then
'           strPA140 = Val(strPA140) - 19110000
'        End If
'        txtDt.Text = strPA140
'     Else
'        If Len(strPA140) = 0 Then
'           Unload Me
'           Exit Sub
'        End If
'     End If
'   End If
'end 2018/3/16
   
      m_PrevF.txtFavDt.Text = strPA140
      If m_PrevF.Name = "frm040101_1" Or m_PrevF.Name = "frm050101_2" Then
         'Modified by Morgan 2018/3/16
         'If Trim(txtDT) = "" Then txtDT.Text = strPA140
         If strPA140 <> "" Then
            txtDT.Text = strPA140
         'end 2018/3/16
            txtDt_Validate Cancel
            If m_PrevF.Name = "frm040101_1" Then
               If strPA10 = "" Then 'Added by Morgan 2022/8/19 提申後補主張則不必更新期限
                  m_PrevF.Text1(4).Text = strLimit1
                  m_PrevF.Text1(5).Text = strLimit2
               End If
            Else
               m_PrevF.txtCaseField(4).Text = strLimit1
               m_PrevF.txtCaseField(9).Text = strLimit2
            End If
         End If 'Added by Morgan 2018/3/16
      End If
   
   End If 'Added by Morgan 2018/3/16
   
   Unload Me
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description
End Sub

'分析字串並存入ListBox
Private Sub Form_Load()
Dim i As Integer, varCountryTemp As Variant, strTemp As String
MoveFormToCenter Me

'Modifeid by Lydia 2015/02/25  發文只DoubleCheck(重新輸入第2次)
If m_dbCheck = True Then

Else
    strTemp = "select FPD05-19110000 from FavPriDate where FPD01='" & strFPD01 & "' AND  FPD02='" & strFPD02 & "' AND FPD03='" & strFPD03 & "' AND FPD04='" & strFPD04 & "' order by FPD05 "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strTemp)
    If intI = 1 Then
        RsTemp.MoveFirst
        For i = 0 To RsTemp.RecordCount - 1
           lstDate.AddItem RsTemp.Fields(0)
           RsTemp.MoveNext
        Next i
    End If
    If lstDate.ListCount > 0 Then
       lstDate.ListIndex = 0
    End If
End If
   
m_bolAdd = False 'Added by Lydia 2015/03/16

m_bolFMP = PUB_ChkIsFMP(strFPD01, strFPD02, strFPD03, strFPD04) 'Added by Lydia 2025/09/25

End Sub

Private Sub lstDate_DblClick()
   If lstDate.ListCount > 0 Then
      txtDT = lstDate.Text
   End If
End Sub

Private Sub cmdMove_Click(Index As Integer)
   Dim i As Integer, intlastIndex As Integer
   Dim varTemp As String
   Dim Cancel As Boolean
   
If Index = 0 Then '新增
    
   If txtDT.Text = "" Then Exit Sub
   
   txtDt_Validate Cancel
   If Cancel = True Then
      Exit Sub
   End If

   For i = 0 To lstDate.ListCount - 1
       varTemp = LTrim(RTrim(lstDate.List(i)))
       If txtDT = varTemp Then
          Exit For
       End If
   Next
   '沒有重複時新增
   If i = lstDate.ListCount Then
      lstDate.AddItem txtDT
      If lstDate.ListCount = 1 Then lstDate.ListIndex = 0
      
   Else
      ShowMsg MsgText(9199)
   End If
ElseIf Index = 1 Then '刪除
   If lstDate.ListIndex = -1 Then
      ShowMsg MsgText(8006)
   Else
      intlastIndex = lstDate.ListIndex
      lstDate.RemoveItem lstDate.ListIndex
      If lstDate.ListCount <> 0 Then
         If intlastIndex = lstDate.ListCount Then
            lstDate.ListIndex = lstDate.ListCount - 1
         Else
            lstDate.ListIndex = intlastIndex
         End If
      End If
   End If
Else
   txtDT.Text = ""
   lstDate.Clear
End If
txtDT.SetFocus
End Sub

Private Sub txtDt_GotFocus()
   TextInverse txtDT
   CloseIme
End Sub

'法限=新穎性優惠期日期+6個月
'所限=法限-7天(台灣P案-2天)
'Modified by Morgan 2022/8/19 若為申請後補主張則改用申請日判斷--蕭茹曣
Private Sub txtDt_Validate(Cancel As Boolean)
Dim stDate As String, iMonth As Integer
Dim stChkDate As String, stDateName As String 'Added by Morgan 2022/8/18

   'Added by Morgan 2022/8/19
   If strPA10 <> "" Then
      stChkDate = DBDATE(strPA10)
      stDateName = "申請日"
   Else
      stChkDate = strSrvDate(1)
      stDateName = "系統日"
   End If
   'end 2022/8/19
   
   'Added by Morgan 2017/7/6 台灣改12個月--潘韻丞(請作單)
   'frm040101_1 也要同步修改
   If strNation = "000" And strFPD01 = "P" Then
      'Modified by Morgan 2018/3/16 台灣設計6個月 - 玲玲
      If strPA08 = "3" Then
         iMonth = 6
      Else
         iMonth = 12
      End If
      'end 2018/3/26
   Else
      iMonth = 6
   End If
   'end 2017/7/6
   
   Cancel = False
   If txtDT <> "" Then
      If ChkDate(txtDT) Then
         stDate = TransDate(CompDate(1, iMonth, txtDT), 1)
         'Modified by Morgan 2022/8/19
         'If Val(stDate) < Val(strSrvDate(2)) Then
         '   MsgBox "優惠期+" & iMonth & "個月不可小於系統日！"
         If DBDATE(stDate) < stChkDate Then
            MsgBox "優惠期+" & iMonth & "個月不可小於" & stDateName & "！"
         'end 2022/8/19
            txtDT.SetFocus
            Cancel = True
            Exit Sub
         End If
         
         If strLimit2 = "" Or Val(strLimit2) > Val(stDate) Then
            strLimit2 = stDate
            If strNation = "000" And strFPD01 = "P" Then
               'Modified by Lydia 2025/09/25
               'strLimit1 = TransDate(PUB_GetWorkDay1(CompDate(2, -2, stDate), True), 1)
               strLimit1 = TransDate(PUB_GetOurDeadline(stDate), 1)
            Else
               'Added by Lydia 2025/10/29
               If m_bolFMP = False And strFPD01 = "P" And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                   strLimit1 = PUB_GetPOurDeadline(stDate, strNation)
               Else
               'end 2025/10/29
                   strLimit1 = TransDate(PUB_GetWorkDay1(CompDate(2, -7, stDate), True), 1)
               End If 'Added by Lydia 2025/10/29
            End If
            If Val(strLimit1) < Val(strSrvDate(2)) Then
               strLimit1 = strSrvDate(2)
            End If
         End If
      Else
         Cancel = True
      End If
   End If
   
   If Val(txtDT) >= Val(strSrvDate(2)) Then
      MsgBox "優惠期事實發生日期不可大於等於系統日！"
      txtDT.SetFocus
      Cancel = True
   End If
End Sub
'Added by Lydia 2015/02/25 發文做DoubleCheck
Private Sub CheckRecExist()
   Dim tempSql As String, idR As Integer
   
   m_bolDouble = False
   tempSql = "select FPD05-19110000 from FavPriDate where FPD01='" & strFPD01 & "' AND  FPD02='" & strFPD02 & "' AND FPD03='" & strFPD03 & "' AND FPD04='" & strFPD04 & "' order by FPD05 "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, tempSql)
    If intI = 1 Then
        If RsTemp.RecordCount <> lstDate.ListCount Then
           
        Else
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               m_bolDouble = False
               For idR = 0 To lstDate.ListCount - 1
                   If RsTemp.Fields(0) = lstDate.List(idR) Then
                      m_bolDouble = True
                      Exit For
                   End If
               Next idR
               If m_bolDouble = False Then Exit Do
               RsTemp.MoveNext
            Loop
        End If
        If m_bolDouble = False Then
           MsgBox "本次輸入之公開日與前次資料不一致,請查明 !", vbCritical
        Else
           m_PrevF.txtFavDt.Text = strPA140
        End If
        Unload Me
        Exit Sub
    'Added by Lydia 2015/03/16 若之前已輸入新穎性優惠期,會無記錄
    Else
        m_bolAdd = True
        MsgBox "請先在分案作業輸入優惠期資料!!", vbCritical
    End If

End Sub
