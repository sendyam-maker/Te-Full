VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100131_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "程式修改公告"
   ClientHeight    =   5280
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7872
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7872
   Begin VB.CommandButton CmdPaper 
      Caption         =   "附件"
      Height          =   540
      Left            =   60
      TabIndex        =   8
      Top             =   4080
      Width           =   700
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   6720
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   45
      Width           =   1035
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "下一筆(&N)"
      Height          =   400
      Index           =   0
      Left            =   5640
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   45
      Width           =   1035
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   1
      Top             =   690
      Width           =   615
   End
   Begin VB.Frame FrmSysKind 
      Height          =   4455
      Left            =   5040
      TabIndex        =   29
      Top             =   600
      Width           =   2655
      Begin VB.CheckBox Check1 
         Caption         =   "外商"
         Height          =   375
         Index           =   15
         Left            =   1440
         TabIndex        =   43
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "帳務"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "專利處"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "商標處"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   20
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "承辦人"
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   22
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "收文"
         Height          =   375
         Index           =   10
         Left            =   1440
         TabIndex        =   23
         Top             =   2880
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "人事"
         Height          =   375
         Index           =   12
         Left            =   1440
         TabIndex        =   25
         Top             =   3360
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "財務"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "分所出納"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "薪資"
         Height          =   375
         Index           =   3
         Left            =   1440
         TabIndex        =   16
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "共同查詢"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "外專"
         Height          =   375
         Index           =   6
         Left            =   1440
         TabIndex        =   19
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "法務"
         Height          =   375
         Index           =   8
         Left            =   1440
         TabIndex        =   21
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "檔案室"
         Height          =   375
         Index           =   11
         Left            =   240
         TabIndex        =   24
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "電腦中心"
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   26
         Top             =   3975
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "每日.每月批次"
         Height          =   375
         Index           =   14
         Left            =   1440
         TabIndex        =   27
         Top             =   3975
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  '置中對齊
         Caption         =   "公佈系統別"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   45
         Width           =   1095
      End
   End
   Begin VB.Frame FrmYN 
      Height          =   615
      Left            =   1200
      TabIndex        =   28
      Top             =   2100
      Width           =   1455
      Begin VB.OptionButton Option1 
         Caption         =   "是"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "否"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   5
         Top             =   120
         Width           =   615
      End
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   5
      Left            =   3810
      TabIndex        =   41
      Top             =   2250
      Visible         =   0   'False
      Width           =   750
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1323;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   1815
      Index           =   4
      Left            =   840
      TabIndex        =   7
      Top             =   3330
      Width           =   4095
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "7223;3201"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   525
      Index           =   3
      Left            =   840
      TabIndex        =   6
      Top             =   2730
      Width           =   4095
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "7223;926"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   1770
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   1050
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   5
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   690
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "時數："
      Height          =   255
      Index           =   8
      Left            =   3225
      TabIndex        =   42
      Top             =   2280
      Visible         =   0   'False
      Width           =   600
   End
   Begin MSForms.Label Label23 
      Height          =   300
      Left            =   60
      TabIndex        =   30
      Top             =   240
      Width           =   5535
      VariousPropertyBits=   27
      Caption         =   "Create ID:        Date        Time        Update ID:        Date        Time"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   165
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "內容："
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   37
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "(民國年月日)"
      Height          =   255
      Left            =   2085
      TabIndex        =   38
      Top             =   720
      Width           =   1020
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   2
      Left            =   1200
      TabIndex        =   11
      Top             =   1455
      Width           =   2295
      BackColor       =   -2147483638
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "流水號："
      Height          =   255
      Index           =   1
      Left            =   3225
      TabIndex        =   39
      Top             =   720
      Width           =   720
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   1
      Left            =   2100
      TabIndex        =   9
      Top             =   1080
      Width           =   1080
      BackColor       =   -2147483638
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "1905;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "摘要："
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   36
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "是否公布："
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   35
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "上線日期："
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   31
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "請作單日："
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   34
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "需求部門："
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   33
      Top             =   1455
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "需求人員："
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   32
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frm100131_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/15 Form2.0已修改
'Memo by Amy 有修改要確認frm100131_2是否也要改
'2013/03/21 Create by Amy
Option Explicit

Dim RbMain As New ADODB.Recordset, bp As New ADODB.Recordset
Dim ActionEdit As Integer '0:add 1:update 2:query 3:cancel
Dim m_AttachPath As String 'Add By Amy 2013/04/30

Dim i As Integer
'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean

'目前正在顯示
Dim m_CurrKEY(2) As String
Dim oText As Object, oCheck As Object, idx As Integer


Private Sub cmdButton_Click(Index As Integer)
    Select Case Index
        Case 0
            tmpBol = fnCancelNowFormAndShowParentForm(Me)
        Case 1
           fnCloseAllFrm100
    End Select
End Sub

Private Sub CmdPaper_Click()
   Dim hLocalFile As Long
   Dim stFileName As String

   Screen.MousePointer = vbHourglass
  
  stFileName = Text1(0) & Text2
   If GetAttachFile(m_AttachPath, stFileName) = False Then
       Screen.MousePointer = vbDefault
       Exit Sub
   End If
   stFileName = m_AttachPath & "\" & stFileName & ".pdf"
   ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
 '取得使用者執行各項功能的權限
   'm_bInsert = IsUserHasRightOfFunction("frm100131_1", strAdd, False)
   'm_bUpdate = IsUserHasRightOfFunction("frm100131_1", strEdit, False)
   'm_bDelete = IsUserHasRightOfFunction("frm100131_1", strDel, False)
   'Add by Amy 20130624 +路徑
   m_AttachPath = App.path & "\PGMBulletinAttach"
    
   MoveFormToCenter Me
   bolToEndByNick = False
   'Add by Amy 2014/07/16 電腦中心顯示時數
   If Pub_StrUserSt03 = "M51" Then
        Label1(8).Visible = True
        Text1(5).Visible = True
   End If
   'end 2014/07/16
   TxtLock 4
End Sub

Private Sub SetTxtValue()
Dim strTmp As String, m_ibf01 As String, m_ibf02 As String
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   'Added by Lydia 2023/12/27
   If m_CurrKEY(0) >= 新部門啟用日 Then
      strSql = "SELECT BU01, BU03,BU04,BU05,BU14,Decode(NVL(BU15,0),0,null,BU15) As BU15,BU06,BU07,BU02,ST02,ST03,BU08,BU09,BU10,BU11,BU12,BU13,NVL(A0922,A0902) AS CDEPT " & _
               "FROM PGMBulletin,STAFF, ACC090, ACC090NEW " & _
               "WHERE BU03=ST01 And BU01='" & m_CurrKEY(0) & "' And BU02='" & m_CurrKEY(1) & "' AND ST03=A0901(+) AND ST93=A0921(+) ORDER BY BU01,BU03"
   Else
   'end 2023/12/27
   'Modify by Amy 2014/07/16 +BU15
      'Modified by Lydia 2023/12/27 + cdept
      strSql = "SELECT BU01, BU03,BU04,BU05,BU14,Decode(NVL(BU15,0),0,null,BU15) As BU15,BU06,BU07,BU02,ST02,ST03,BU08,BU09,BU10,BU11,BU12,BU13,'' AS CDEPT FROM PGMBulletin,STAFF " & _
                   "WHERE BU03=ST01 And BU01='" & m_CurrKEY(0) & "' And BU02='" & m_CurrKEY(1) & "' ORDER BY BU01,BU03"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
    For Each oText In Text1
      idx = oText.Index
      If IsNull(rsTmp(idx)) Then
         oText = ""
      Else
        Select Case idx
           Case 0, 2
            oText = ChangeWStringToTString(rsTmp(idx))
           Case 1, 3, 4
            oText = rsTmp(idx)
           'Modify by Amy 2014/07/16 +BU15
           Case 5
            oText = Val(rsTmp(idx))
        End Select
      End If
    Next
    
    '是否公佈
    Select Case rsTmp.Fields("BU06")
        Case 0
          Option1(1).Value = True
        Case 1
          Option1(0).Value = True
    End Select
   
   Text2.Text = IIf(rsTmp.Fields("BU02") <= 9, Format(Val(rsTmp.Fields("BU02")), "00"), rsTmp.Fields("BU02")) '流水號
    Label2(1).Caption = rsTmp.Fields("ST02") '需求人員名稱
    'Added by Lydia 2023/12/27
    If DBDATE(Text1(0)) >= 新部門啟用日 Then
       Label2(2).Caption = IIf("" & rsTmp.Fields("cdept") = "", "" & rsTmp.Fields("st03"), "" & rsTmp.Fields("cdept"))
    Else
    'end 2023/12/27
       Label2(2).Caption = IIf(ClsPDGetStaffDeptName(rsTmp.Fields("ST03"), strTmp), strTmp, "") '需求部門名稱
    End If
   RedSystemKind IIf(IsNull(rsTmp.Fields("BU07")), "", rsTmp.Fields("BU07")) '系統別
   End If
   
  'Add By Amy 2013/05/03 Start
   '判斷ImgByteFile是否有資料,沒有請作單鈕設Disabled
   'Modify by Amy 2023/06/09 +if 放99年以前的附件會抓不到
   If Len(Text1(0)) = 7 Then
      m_ibf01 = Left(Text1(0), 3)
   Else
      m_ibf01 = "0" & Left(Text1(0), 2)
   End If
   m_ibf02 = Right(Text1(0), 4) & Text2
   
   strExc(0) = "Select  * From ImgByteFile Where IBF01='" & m_ibf01 & "' And IBF02='" & m_ibf02 & "' And IBF03='0' And IBF04='00'"
   If RsTemp.State <> adStateClosed Then RsTemp.Close
   RsTemp.CursorLocation = adUseClient
   RsTemp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If RsTemp.RecordCount > 0 Then
      CmdPaper.Enabled = True
   Else
      CmdPaper.Enabled = False
   End If
   
   '更新CUID
   UpdateCUID rsTmp
   '2013/05/03 End
End Sub

Private Sub TxtClear()
   Dim txt As Object, Lbl As Object, Chk As Object
   For Each txt In frm100131_1.Text1
      txt.Text = ""
   Next
   Text2.Text = "" '流水號
   For Each Lbl In Label2
      Lbl = ""
   Next
   Option1(0).Value = 1
   For Each Chk In Check1
      Chk.Value = 0
   Next
End Sub

Private Sub TxtLock(ByVal Lt As Integer)
 Dim txt As Object, Chk As Object, i As Integer
   Select Case Lt
      Case 4 '全鎖
        For Each txt In frm100131_1.Text1
            txt.Locked = True
         Next
         Text2.Locked = True
        FrmYN.Enabled = False
        FrmSysKind.Enabled = False
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "Select * From PGMBulletin " & _
                "Where BU01 = '" & strKEY01 & "' AND " & _
                          "BU02 = '" & strKEY02 & "' "
                  
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 查詢記錄
Private Function QueryRecord(ByVal strBU01 As String, ByVal strBU02 As String) As Boolean

   QueryRecord = False
   strBU01 = ChangeTStringToWString(strBU01)
   strBU02 = Val(strBU02)
  
   If IsRecordExist(strBU01, strBU02) = True Then
      m_CurrKEY(0) = strBU01
      m_CurrKEY(1) = strBU02
      QueryRecord = True
      
   Else
      QueryRecord = False
      MsgBox ("無此資料")
      If ActionEdit = 4 Then
         Screen.MousePointer = vbDefault
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
         Exit Function
      End If
   End If
    SetTxtValue
  
End Function

Private Function ChkSystemKind() As String
   Dim returnVal As String

    For Each oCheck In Check1
      idx = oCheck.Index
           
      If oCheck.Value Then
      Select Case idx
        Case 0
          returnVal = "Account"
        Case 1
          returnVal = "Finance"
        Case 2
          returnVal = "Casher"
        Case 3
          returnVal = "Salary"
        Case 4
          returnVal = "Query"
        Case 5
          returnVal = "Patpro"
        Case 6
          returnVal = "Patpro1"
        Case 7
          returnVal = "Trademark"
        Case 8
          returnVal = "Law"
        Case 9
          returnVal = "Promoter"
        Case 10
          returnVal = "Writer"
        Case 11
          returnVal = "File"
        Case 12
          returnVal = "Person"
        Case 13
          returnVal = "Computer"
        Case 14
          returnVal = "AutoBatch"
        'Add by Amy 2018/11/14
        Case 15
          returnVal = "Trademark1"
      End Select
           ChkSystemKind = ChkSystemKind & "," & returnVal
      End If
    Next
    '去掉第一個,
    ChkSystemKind = Mid(ChkSystemKind, 2)
End Function
Private Sub RedSystemKind(ByVal SysKind As String)
  If Trim(SysKind) <> "" Then
     Dim strTmp() As String
     strTmp = Split(SysKind, ",")
     
     For i = 0 To UBound(strTmp)
        Select Case strTmp(i)
          Case "Account"
            Check1(0).Value = 1
          Case "Finance"
            Check1(1).Value = 1
          Case "Casher"
            Check1(2).Value = 1
          Case "Salary"
            Check1(3).Value = 1
          Case "Query"
            Check1(4).Value = 1
          Case "Patpro"
            Check1(5).Value = 1
          Case "Patpro1"
            Check1(6).Value = 1
          Case "Trademark"
           Check1(7).Value = 1
          Case "Law"
            Check1(8).Value = 1
          Case "Promoter"
            Check1(9).Value = 1
          Case "Writer"
            Check1(10).Value = 1
          Case "File"
            Check1(11).Value = 1
          Case "Person"
            Check1(12).Value = 1
          Case "Computer"
            Check1(13).Value = 1
          Case "AutoBatch"
            Check1(14).Value = 1
          'Add by Amy 2018/11/14
          Case "Trademark1"
           Check1(15).Value = 1
      End Select
   Next
  End If
End Sub

Sub StrMenu()
Dim strSql  As String, i As Integer
Dim Str01, Str02 As String
Dim intIndex As Integer

ActionEdit = 4
Str01 = "": Str02 = ""
intIndex = InStr(Me.Tag, ",")
Str01 = ChangeTDateStringToTString(Left(Me.Tag, intIndex - 1))
Str02 = Mid(Me.Tag, intIndex + 1)

QueryRecord Str01, Str02

End Sub

'Add By Amy 2013/04/30 pFileName為上線日(民國年月日+序號)
Private Function GetAttachFile(ByVal m_AttachPath As String, ByVal pFileName As String) As Boolean
   
   Dim stAttPath As String, m_ibf01 As String, m_ibf02 As String
   m_ibf01 = "": m_ibf02 = ""
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte
   
On Error GoTo ErrHnd
   'Modify by Amy 2023/06/09 +if 放99年以前的附件會抓不到
   'Modify By Sindy 2023/6/12 pFileName=年度+6碼流水號
   'If Len(pFileName) = 7 Then
   If Len(pFileName) = 9 Then
   '2023/6/12 END
      m_ibf01 = Left(pFileName, 3)
   Else
      m_ibf01 = "0" & Left(pFileName, 2)
   End If
    m_ibf02 = Right(pFileName, 6)
  
    If Dir(m_AttachPath, vbDirectory) = "" Then
        MkDir m_AttachPath
    End If
    stAttPath = m_AttachPath & "\" & pFileName & ".pdf"
    '檔案已存在時不必重新下載
    If Dir(stAttPath) <> "" Then
        pFileName = stAttPath
        GetAttachFile = True
        Exit Function
    End If
      
   strExc(0) = "Select  * From ImgByteFile Where IBF01='" & m_ibf01 & "' And IBF02='" & m_ibf02 & "' And IBF03='0' And IBF04='00'"
   If RsTemp.State <> adStateClosed Then RsTemp.Close
   RsTemp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If RsTemp.RecordCount > 0 Then
      If Dir(stAttPath) <> "" Then Kill stAttPath
      'Add By Sindy 2017/8/10
'      If "" & RsTemp.Fields("IBF15") <> "" Then
         GetAttachFile = PUB_GetFtpFile(RsTemp.Fields("IBF15"), stAttPath, UCase("ImgByteFile"))
'      Else
'      '2017/8/10 END
'         With RsTemp
'              lngSize = Val(.Fields("IBF13").Value)
'              ReDim bytes(lngSize)
'              If lngSize > 0 Then bytes() = .Fields("IBF14").GetChunk(lngSize)
'         End With
'         iFileNo = FreeFile
'         Open stAttPath For Binary Access Write As #iFileNo
'         If lngSize > 0 Then Put #iFileNo, , bytes()
'         Close #iFileNo
'         GetAttachFile = True
'      End If
      pFileName = stAttPath
   Else
      Close #iFileNo
      MsgBox ("無此請作單資料")
   End If
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   If iFileNo > 0 Then Close #iFileNo
End Function

'Add By Amy 2013/04/30
Private Sub KillAttach()
    Dim strPath As String '防刪到c:\
On Error Resume Next
    strPath = App.path & "\PGMBulletinAttach"
   If Dir(strPath & "\.") <> "" Then
      Kill strPath & "\*.*"
   End If
End Sub

'Add By Amy 2013/05/03 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("BU08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("BU08")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("BU08"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("BU09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("BU09")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("BU09"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("BU10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("BU10")) = False Then
         strTemp = rsSrcTmp.Fields("BU10")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("BU11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("BU11")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("BU11"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("BU12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("BU12")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("BU12"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("BU13")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("BU13")) = False Then
         strTemp = rsSrcTmp.Fields("BU13")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(2, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

Private Sub Form_Unload(Cancel As Integer)
   KillAttach
   Set frm100131_1 = Nothing
End Sub
