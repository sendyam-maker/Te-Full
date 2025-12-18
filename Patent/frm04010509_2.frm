VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010509_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "所外鑑定報告結果"
   ClientHeight    =   5400
   ClientLeft      =   975
   ClientTop       =   990
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7215
   Begin VB.CommandButton cmkok 
      Caption         =   "結束(&X)"
      Default         =   -1  'True
      Height          =   405
      Index           =   2
      Left            =   6324
      TabIndex        =   27
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmkok 
      Caption         =   "回前畫面(&U)"
      Height          =   405
      Index           =   1
      Left            =   5100
      TabIndex        =   26
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmkok 
      Caption         =   "確定(&O)"
      Height          =   405
      Index           =   0
      Left            =   4272
      TabIndex        =   25
      Top             =   70
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   6972
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   22
         Top             =   180
         Width           =   1095
      End
      Begin MSForms.TextBox Text2 
         Height          =   1395
         Left            =   1560
         TabIndex        =   23
         Top             =   540
         Width           =   5295
         VariousPropertyBits=   -1467987941
         Size            =   "9340;2461"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label21 
         Caption         =   "進度備註："
         Height          =   252
         Left            =   120
         TabIndex        =   21
         Top             =   540
         Width           =   972
      End
      Begin VB.Label Label20 
         Caption         =   "來函收文日："
         Height          =   252
         Left            =   120
         TabIndex        =   20
         Top             =   180
         Width           =   1092
      End
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1170
      TabIndex        =   24
      Top             =   1080
      Width           =   5655
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "9975;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   720
   End
   Begin MSForms.Label Label18 
      Height          =   270
      Left            =   1080
      TabIndex        =   17
      Top             =   2880
      Width           =   5955
      VariousPropertyBits=   27
      Size            =   "10504;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   4560
      TabIndex        =   16
      Top             =   2160
      Width           =   900
   End
   Begin MSForms.Label Label16 
      Height          =   255
      Left            =   5565
      TabIndex        =   15
      Top             =   2160
      Width           =   1215
      VariousPropertyBits=   27
      Size            =   "2143;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   4560
      TabIndex        =   14
      Top             =   1800
      Width           =   900
   End
   Begin MSForms.Label Label14 
      Height          =   255
      Left            =   5565
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
      VariousPropertyBits=   27
      Size            =   "2143;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "專利號數："
      Height          =   180
      Left            =   4560
      TabIndex        =   12
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label12 
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號："
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   900
   End
   Begin MSForms.Label Label10 
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
      VariousPropertyBits=   27
      Size            =   "2143;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   720
   End
   Begin MSForms.Label Label8 
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
      VariousPropertyBits=   27
      Size            =   "2143;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   720
   End
   Begin MSForms.Label Label4 
      Height          =   270
      Left            =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   5625
      VariousPropertyBits=   27
      Size            =   "9922;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   900
   End
End
Attribute VB_Name = "frm04010509_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (Text2,Combo1,Label4,Label14,Label8,Label18,Label16,Label10)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String, cp(1 To T_CP) As String
Dim pa() As String, cp() As String

Dim intWhere As Integer
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2016/10/7 END


'Add By Sindy 2022/7/1
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmkok_Click(Index As Integer)
   Select Case Index
       Case 0
         If Text1.Text = "" Then MsgBox "來函收文日不可為空值", vbInformation: Text1.SetFocus: Exit Sub
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         'Add By Sindy 2022/7/1
         If m_strIR01 <> "" And Left(Pub_StrUserSt03, 2) = "F2" Then
            If PUB_ChkFileOpening2(m_PrevForm.m_strFullFileName, "後續才能一併歸卷！") = True Then
               Exit Sub
            End If
         End If
         '2022/7/1 END
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         'Add By Sindy 2016/10/7
         If Me.m_strIR01 <> "" Then
            Unload frm04010509_1
            Unload Me
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         Else
         '2016/10/7 END
            frm04010509_1.Show
            Unload Me
         End If
       Case 1
         frm04010509_1.Show
         Unload Me
      Case 2
         Unload frm04010509_1
         Unload Me
   End Select
End Sub
'存檔
Private Function FormSave() As Boolean
 Dim autonum As String
 Dim strTxt(1 To 10) As String
 
 'Add by Morgan 2004/2/9
 Dim stCP12 As String, stCP13 As String
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
FormSave = True
cnnConnection.BeginTrans

   autonum = AutoNo("C", 6)
   
   strTxt(1) = "UPDATE CASEPROGRESS SET CP24='1' WHERE cp09='" & cp(9) & "'"
   'Add By Cheng 2002/11/06
    cnnConnection.Execute strTxt(1)
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
    
    'Modify by Morgan 2004/2/9
   'strTxt(2) = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13," & _
      "cp14,cp20,CP26,cp27,cp32,cp43,cp64) values('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & _
      TransDate(Text1.Text, 2) & "','" & autonum & "','1903'," & CNULL(cp(12)) & "," & _
      CNULL(PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))) & "," & CNULL(strUserNum) & ",'N','N'," & strSrvDate(1) & ",'N','" & _
      cp(9) & "'," & CNULL(ChgSQL(Text2.Text)) & ")"
      
    stCP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
    stCP12 = GetSalesArea(stCP13)
      
   strTxt(2) = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13," & _
      "cp14,cp20,CP26,cp27,cp32,cp43,cp64) values('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & _
      TransDate(Text1.Text, 2) & "','" & autonum & "','1903'," & CNULL(stCP12) & "," & _
      CNULL(stCP13) & "," & CNULL(strUserNum) & ",'N','N'," & strSrvDate(1) & ",'N','" & _
      cp(9) & "'," & CNULL(ChgSQL(Text2.Text)) & ")"
      
    'Modify end 2004/2/9
      
   'Add By Cheng 2002/11/06
    cnnConnection.Execute strTxt(2)
   strTxt(3) = "update nextprogress set np06='Y' where np01='" & cp(9) & _
      "' and np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & _
      "' and np05='" & pa(4) & "' and np07='411'"
   'Add By Cheng 2002/11/06
    cnnConnection.Execute strTxt(3)
'   FormSave = objLawDll.ExecSQL(3, strTxt)

   'Added by Morgan 2016/6/6
   '只有台灣案會輸入，目前幾乎沒有，先比照其他代理人來函控制
   If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
      PUB_AddLetterProgress autonum, 1, False
   End If
   'end 2016/6/6
   
   'Add by Sindy 2016/10/7
   If m_strIR01 <> "" Then
      'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", autonum, "")
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm04010509_1", IIf(Pub_StrUserSt03 = "F22", autonum, "")
   End If
   '2016/10/7 END
   
   cnnConnection.CommitTrans
   Exit Function
   
ErrorHandler:
   cnnConnection.RollbackTrans
   FormSave = False
End Function

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
ReDim cp(1 To TF_CP) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   
   'Add By Sindy 2018/1/2
   m_strIR01 = frm04010509_1.m_strIR01
   m_strIR02 = frm04010509_1.m_strIR02
   m_strIR03 = frm04010509_1.m_strIR03
   m_strIR04 = frm04010509_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
      Text1 = frm04010509_1.m_RDate
   End If
   '2018/1/2 END
   
   ReadPatent
End Sub

Private Sub ReadPatent()
 Dim rsTemp1 As New ADODB.Recordset, Lbl As LABEL, i As Integer, j As Integer
 Dim strKey(0 To 5) As String
 Dim nIndex As Integer
 'Add By Cheng 2002/07/08
 Dim StrSQLa As String
 
   pa(1) = frm04010509_1.Text1(0)
   pa(2) = frm04010509_1.Text1(1)
   pa(3) = frm04010509_1.Text1(2)
   pa(4) = frm04010509_1.Text1(3)
   cp(9) = frm04010509_1.Tag
   
   Label2.Caption = pa(1) & pa(2) & pa(3) & pa(4)
   
   If pa(1) = "P" Then
      If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
         If pa(26) <> "" Then ChgType 26, pa(26)
         If pa(9) <> "" Then ChgType 9, pa(9)
         Label12 = pa(22)
      End If
   Else
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
         If pa(8) <> "" Then ChgType 8, pa(8)
         If pa(9) <> "" Then ChgType 9, pa(9)
      End If
   End If
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.ReadCaseProgressDatabase(cp(), intWhere) Then
   If ClsPDReadCaseProgressDatabase(cp(), intWhere) Then
      
   End If
   'Modify By Cheng 2002/07/08
   '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'   strExc(0) = "select cpm03,staff1.st02 as cp13,staff.st02 as cp14,CP45," & _
'      "nvl(fa05,nvl(fa04,nvl(fa06,'')))" & _
'      " from caseprogress,casepropertymap,staff,staff staff1,fagent where " & _
'      "cp09='" & cp(9) & "' and cp01=cpm01(+) and cp10=cpm02(+) and " & _
'      "cp14=staff.st01(+) and cp13=staff1.st01(+) and SUBSTR(CP44,1,8)=FA01(+) AND " & _
'      "SUBSTR(CP44,9,1)=FA02(+)"
   StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人"
   strExc(0) = "select cpm03,staff1.st02 as cp13,staff.st02 as cp14,CP45," & _
      StrSQLa & _
      " From caseprogress,casepropertymap,staff,staff staff1,fagent,SystemKind where " & _
      "cp09='" & cp(9) & "' and cp01=cpm01(+) and cp10=cpm02(+) and " & _
      "cp14=staff.st01(+) and cp13=staff1.st01(+) and SUBSTR(CP44,1,8)=FA01(+) AND " & _
      "SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
      If intI = 1 Then
         If Not IsNull(.Fields(0)) Then Label6 = .Fields(0)
         If Not IsNull(.Fields(1)) Then Label14 = .Fields(1)
         If Not IsNull(.Fields(2)) Then Label8 = .Fields(2)
         If Not IsNull(.Fields(3)) Then Label10 = .Fields(3)
         If Not IsNull(.Fields(4)) Then Label18 = .Fields(4)
      End If
   End With
End Sub

Private Sub ChgType(ByVal Index As Integer, strID As String)
   Select Case Index
      Case 9
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(strID, strExc(0)) Then
         If ClsPDGetNation(strID, strExc(0)) Then
            Label16 = strExc(0)
         Else
            Label16 = ""
         End If
      Case 26
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(strID, strExc(0)) Then
         If ClsLawLawGetName(strID, strExc(0)) Then
            Label4 = strExc(0)
         Else
            Label4 = ""
         End If
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010509_2 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 <> "" Then
      If ChkDate(Text1) Then
         'Modify by Morgan 2010/8/11 百年蟲
         'If Text1.Text > strSrvDate(2) Then
         If Val(Text1.Text) > Val(strSrvDate(2)) Then
            MsgBox "輸入日期不可大於系統日", vbInformation
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text1
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text2.IMEMode = 1
   OpenIme
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text2.IMEMode = 2
   CloseIme
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

   'Added by Morgan 2021/12/20 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/20
   
If Me.Text1.Enabled = True Then
   Cancel = False
   Text1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function
