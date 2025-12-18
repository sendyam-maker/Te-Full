VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010305_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-延緩公告"
   ClientHeight    =   3720
   ClientLeft      =   870
   ClientTop       =   1515
   ClientWidth     =   7740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7740
   Begin VB.Frame FraPA174 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   285
      Left            =   1050
      TabIndex        =   32
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   900
         Style           =   1  '圖片外觀
         TabIndex        =   33
         Top             =   -30
         Width           =   800
      End
      Begin VB.Label lblPA174 
         Caption         =   "有特殊字"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   35
         TabIndex        =   34
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1935
      MaxLength       =   1
      TabIndex        =   1
      Top             =   3030
      Width           =   300
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   13
      Top             =   540
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2460
      MaxLength       =   1
      TabIndex        =   12
      Top             =   540
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   11
      Top             =   540
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1065
      MaxLength       =   3
      TabIndex        =   10
      Top             =   540
      Width           =   550
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm06010305_1.frx":0000
      Left            =   1065
      List            =   "frm06010305_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   9
      Top             =   1170
      Width           =   615
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1230
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2670
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   4620
      MaxLength       =   7
      TabIndex        =   2
      Top             =   3030
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5676
      TabIndex        =   4
      Top             =   70
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4848
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6810
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   75
      X2              =   7515
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   75
      X2              =   7515
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書內容          (Y:WORD)"
      Height          =   180
      Index           =   1
      Left            =   225
      TabIndex        =   31
      Top             =   3075
      Width           =   2880
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   7
      Left            =   4065
      TabIndex        =   30
      Top             =   2160
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   6
      Left            =   1065
      TabIndex        =   29
      Top             =   2160
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   5
      Left            =   4065
      TabIndex        =   28
      Top             =   1890
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   4
      Left            =   1065
      TabIndex        =   27
      Top             =   1890
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   3
      Left            =   1710
      TabIndex        =   26
      Top             =   1230
      Width           =   5940
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "10477;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   2
      Left            =   4065
      TabIndex        =   25
      Top             =   900
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   1
      Left            =   1065
      TabIndex        =   24
      Top             =   900
      Width           =   2070
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3651;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   285
      TabIndex        =   23
      Top             =   1230
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   3285
      TabIndex        =   22
      Top             =   900
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   285
      TabIndex        =   21
      Top             =   900
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   285
      TabIndex        =   20
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   285
      TabIndex        =   19
      Top             =   1890
      Width           =   765
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3285
      TabIndex        =   18
      Top             =   1890
      Width           =   765
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   0
      Left            =   4065
      TabIndex        =   17
      Top             =   600
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   105
      TabIndex        =   16
      Top             =   2160
      Width           =   945
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   3285
      TabIndex        =   15
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3285
      TabIndex        =   14
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期:"
      Height          =   180
      Left            =   225
      TabIndex        =   8
      Top             =   2700
      Width           =   945
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "延　緩　至"
      Height          =   180
      Left            =   3615
      TabIndex        =   7
      Top             =   3075
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "公告"
      Height          =   180
      Index           =   0
      Left            =   5700
      TabIndex        =   6
      Top             =   3075
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "frm06010305_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/4 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Public strReceiveNo As String
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String
Dim intWhere As Integer


Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
 Dim strTxt(5) As String
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   strExc(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他公告日'," & CNULL(Text6.Text) & ")"
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(1, strExc) Then
   If Not ClsLawExecSQL(1, strExc) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

'延緩公告日
Private Sub cmdOK_Click(Index As Integer)
 Dim bolChk As Boolean
   Select Case Index
      Case 0
         If Text7 = "Y" Then
            bolChk = True
         Else
            bolChk = False
         End If
         'Added by Lydia 2020/02/17 產生各式申請書時，若基本檔「名稱有特殊字」已勾選，彈訊息提醒，並一併開啟原始檔。
         If (pa(1) = "FCP" Or pa(1) = "P") And pa(174) = "Y" Then
             MsgBox MsgText(1111), vbInformation
             If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = False Then
                 Exit Sub
             End If
         End If
         'end 2020/02/17
         
         StartLetter "01", "00"
         strLetterDate = Text5.Text
         NowPrint strReceiveNo, "01", "00", bolChk, strUserNum
         frm060103_1.Show
         ' 90.08.27 modify by louis (回到原畫面要清除畫面)
         frm060103_1.ClearForm
      Case 1
         frm060103_1.Show
      Case 2
         Unload frm060103_1
   End Select
   Unload Me
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label12(3) = pa(5)
      Case "英"
         Label12(3) = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label12(3) = pa(7)
   End Select
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm060103_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   ReadPatent
   Combo1.ListIndex = 0
   Text5.Text = strSrvDate(2)
   
   FraPA174.BackColor = &H8000000F 'Added by Lydia 2020/02/21
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010305_1 = Nothing
End Sub

Private Sub ReadPatent()
 Dim rsTemp1 As New ADODB.Recordset, Lbl As Object
   For Each Lbl In Label12
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Text5 = pa(10)
      Label12(1) = pa(11)
      Label12(2) = pa(22)
      Label12(3) = pa(5)
   End If
   strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2," & _
      "cp43 from caseprogress,casepropertymap,staff," & _
      "staff staff1 where cp09='" & strReceiveNo & "' AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and " & _
      "cp13=staff1.st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      If Not IsNull(.Fields(0)) Then Label12(0) = .Fields(0)
      If Not IsNull(.Fields(1)) Then Label12(4) = .Fields(1)
      If Not IsNull(.Fields(2)) Then Label12(5) = .Fields(2)
      If Not IsNull(.Fields(3)) Then
         strExc(0) = "SELECT CP05,CP08 FROM CASEPROGRESS WHERE CP09='" & .Fields(3) & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(rsTemp1.Fields(0)) Then Label12(6) = TransDate(rsTemp1.Fields(0), 1)
            If Not IsNull(rsTemp1.Fields(1)) Then Label12(7) = rsTemp1.Fields(1)
         End If
      End If
   End If
   End With
   
   'Added by Lydia 2020/02/21 預設「名稱有特殊字」
   FraPA174.Visible = False
   If pa(1) = "FCP" Or pa(1) = "P" Then
       If pa(174) = "Y" Then
          FraPA174.Visible = True
       End If
   End If
   'end 2020/02/21
   
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

'MODIFY BY SONIA 90.10.6
'Private Sub Text6_GotFocus()
'  TextInverse Text6
'End Sub
'MODIFY BY SONIA 90.10.6
'Private Sub Text6_Validate(Cancel As Boolean)
'   If Text6 = "" Then
'      MsgBox "延緩日期不可空白 !", vbCritical
'      Cancel = True
'   Else
'      If Not ChkDate(Text6) Or Val(Text6) < Val(strSrvDate(2)) Then
'         MsgBox "延緩日期不正確或延緩日期小於系統日，請重新輸入 !", vbCritical
'         Cancel = True
'      End If
'   End If
'End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()

    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
    
End Sub
