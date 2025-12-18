VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060105_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "申請案號輸入"
   ClientHeight    =   5856
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5856
   ScaleWidth      =   9312
   Begin VB.TextBox txtRecDate 
      Height          =   270
      Left            =   4320
      MaxLength       =   1
      TabIndex        =   3
      Top             =   2220
      Width           =   375
   End
   Begin VB.TextBox txtEmail 
      Height          =   270
      Left            =   7620
      MaxLength       =   1
      TabIndex        =   4
      Top             =   2220
      Width           =   375
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Index           =   2
      Left            =   1170
      MaxLength       =   8
      TabIndex        =   6
      Top             =   4890
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定綠皮貼紙印表機"
      Height          =   660
      Left            =   4830
      TabIndex        =   29
      Top             =   4890
      Width           =   4365
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   150
         Style           =   2  '單純下拉式
         TabIndex        =   30
         Top             =   240
         Width           =   4110
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   7152
      TabIndex        =   11
      Top             =   48
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1110
      MaxLength       =   3
      TabIndex        =   18
      Top             =   450
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1590
      MaxLength       =   6
      TabIndex        =   17
      Top             =   450
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   270
      Left            =   2430
      MaxLength       =   1
      TabIndex        =   16
      Top             =   450
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   270
      Left            =   2685
      MaxLength       =   2
      TabIndex        =   15
      Top             =   450
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1110
      TabIndex        =   13
      Top             =   1425
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1170
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1890
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   4980
      TabIndex        =   1
      Top             =   1890
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   1170
      TabIndex        =   2
      Top             =   2190
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6312
      TabIndex        =   10
      Top             =   48
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8376
      TabIndex        =   12
      Top             =   48
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm060105_2.frx":0000
      Left            =   1110
      List            =   "frm060105_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   14
      Top             =   810
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Index           =   0
      Left            =   1170
      MaxLength       =   8
      TabIndex        =   7
      Top             =   5190
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Index           =   1
      Left            =   1170
      MaxLength       =   8
      TabIndex        =   8
      Top             =   5490
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "更新期限"
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   9
      Top             =   5160
      Width           =   972
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2355
      Left            =   120
      TabIndex        =   5
      Top             =   2490
      Width           =   9075
      _ExtentX        =   16002
      _ExtentY        =   4149
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblRecDate 
      AutoSize        =   -1  'True
      Caption         =   "當天報告:             (Y:是)"
      Height          =   180
      Left            =   3450
      TabIndex        =   36
      Top             =   2265
      Width           =   1815
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "Email維護:             (Y:是)"
      Height          =   180
      Left            =   6690
      TabIndex        =   35
      Top             =   2265
      Width           =   1860
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   2
      Left            =   1110
      TabIndex        =   34
      Top             =   1140
      Width           =   8145
      BackColor       =   12632256
      VariousPropertyBits=   27
      Size            =   "14367;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   1
      Left            =   1770
      TabIndex        =   33
      Top             =   810
      Width           =   7455
      BackColor       =   12632256
      VariousPropertyBits=   27
      Size            =   "13150;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   0
      Left            =   4980
      TabIndex        =   32
      Top             =   450
      Width           =   1395
      BackColor       =   12632256
      VariousPropertyBits=   27
      Size            =   "2461;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "約定期限"
      Height          =   180
      Index           =   3
      Left            =   330
      TabIndex        =   31
      Top             =   4920
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   28
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      Height          =   180
      Left            =   4170
      TabIndex        =   27
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   810
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   90
      TabIndex        =   25
      Top             =   1455
      Width           =   945
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "申請日期:"
      Height          =   180
      Left            =   120
      TabIndex        =   24
      Top             =   1890
      Width           =   765
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   4170
      TabIndex        =   23
      Top             =   1920
      Width           =   765
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "三聯單文號:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   2190
      Width           =   945
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "申請人:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   1170
      Width           =   585
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "本所期限"
      Height          =   180
      Index           =   1
      Left            =   330
      TabIndex        =   20
      Top             =   5220
      Width           =   720
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "法定期限"
      Height          =   180
      Index           =   2
      Left            =   330
      TabIndex        =   19
      Top             =   5520
      Width           =   720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   9180
      Y1              =   1815
      Y2              =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   9180
      Y1              =   1785
      Y2              =   1785
   End
End
Attribute VB_Name = "frm060105_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/21 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String 'Memo by Lydia 2020/07/13 收文號(新案)
Dim strKind As String   'Memo by Lydia 2010/07/13 收文號(新案)案件性質
Dim cp(1 To 5) As String 'Memo by Lydia 2020/07/13 1-收文號(新案), 2-業務區(CP12), 3-智權人員(CP13), 4-下一程序(NP01), 5-下一程序(NP02)

'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String

Dim intWhere As Integer
Dim intLastRow As Integer
' 申請日
Dim m_PA10 As String

Dim SeekPrint As Integer, SeekPrintL As Integer
Private Const m_ET01 As String = "16" 'Add by Morgan 2010/11/11
'Added by Lydia 2019/06/25 FCP特定案件性質的電子送件，發文確定後直接跳到"申請案號輸入"畫面，直接key 申請案號。
Public PubOtherCall As String '傳入"確定index+表單名稱+本所案號"
Dim mOtherForm As String  '確定index+表單名稱
Dim m_pAgreeOnDate As String 'Modify By Sindy 2021/4/27
'Add By Sindy 2022/5/11
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2022/5/11 END
Dim strChk As String 'Add By Sindy 2022/5/11
Dim m_Done As Boolean 'Add By Sindy 2022/5/25


'Add By Sindy 2022/5/11
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim strTo As String, strCC As String, strSubject As String, strContent As String, strCP10 As String
   
 Dim strTmp(1 To 2) As String
   Select Case Index
      Case 0
'         If objLawDll.ChkMRec(Text5.Text, pa(1) & pa(2) & pa(3) & pa(4), strTmp(1), strTmp(2)) Then
'            If strTmp(1) <> "" Then
'               If MsgBox("與櫃台之來函收文記錄 ( " & strTmp(1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
'            End If
'         Else
'            If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
'         End If
         ' 90.08.06 modify by louis (申請日期)
         If IsEmptyText(Text6) = True Then
            MsgBox "請輸入申請日期 !", vbCritical + vbYesNo
            Exit Sub
         End If
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         Screen.MousePointer = vbHourglass
         cmdOK(0).Enabled = False 'Add By Sindy 2022/5/20
         If FormSave = False Then
            cmdOK(0).Enabled = True 'Add By Sindy 2022/5/20
            MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
         Else
            'Add by Morgan 2010/11/11
            strUserNum = strFMPNum
            
            'Added by Morgan 2012/11/29 102新法改定稿
            If strSrvDate(1) >= "20130101" Then
               StartLetter2 "16", pa(1) & pa(2) & pa(3) & pa(4) & "&000", "02"
               NowPrint pa(1) & pa(2) & pa(3) & pa(4) & "&000", "16", "02", False, strUserNum, 0
               
'Removed by Morgan 2013/11/1 不再使用,刪除
'            Else
'            'end 2012/11/29
'               StartLetter "16", pa(1) & pa(2) & pa(3) & pa(4) & "&000", "01"
'               NowPrint pa(1) & pa(2) & pa(3) & pa(4) & "&000", "16", "01", False, strUserNum, 0
            End If
            
            strUserNum = strUser1Num
            'end 2010/11/11
            
            'Add By Sindy 2022/5/11
            '在新案通知申請案號按"確定"後，判斷是否有告代(901)or主動修正(203) 無發文日，
            '且進度備註有: 提申後告代 or 提申後主動修正，並且第一次輸入申請案號時，系統自動發mail。
            'Memo by Lydia 2023/05/10 如果主旨或內文有變，請一併查看frm090902_2的email是否要一致
            If strChk = "第一次輸入" And Text7 <> "" Then
               strExc(0) = "SELECT * FROM caseprogress,staff" & _
                           " WHERE cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                           " and ((cp10='901' and instr(cp64,'提申後告代')>0) or (cp10='203' and instr(cp64,'提申後主動修正')>0))" & _
                           " and cp27||cp57 is null" & _
                           " and cp14=st01(+)" & _
                           " order by cp10 desc"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  RsTemp.MoveFirst
                  strTo = "" & RsTemp.Fields("cp14") '工程師
                  strCC = PUB_GetFCPEngSup(strTo) & ";" & strUserNum & ";backup" '工程師之主管;發文人員;backup
                  strCP10 = "" & RsTemp.Fields("cp10")
                  strSubject = "【1.請分案 2.進行"
                  strContent = "今日已提申完畢" & vbCrLf & _
                               "1.主管請分案" & vbCrLf & _
                               "2.工程師請進行以下事項:" & vbCrLf
                  strExc(10) = 0
                  Do While Not RsTemp.EOF
                     strExc(10) = Val(strExc(10)) + 1
                     If Val(strExc(10)) > 1 Then
                        strSubject = strSubject & "、"
                     End If
                     If RsTemp.Fields("cp10") = "901" Then
                        strSubject = strSubject & "提申後告代"
                        strContent = strContent & "  提申後告代　　　　承辦期限：" & ChangeWStringToTDateString("" & RsTemp.Fields("cp48")) & "　本所期限：" & ChangeWStringToTDateString("" & RsTemp.Fields("cp06")) & vbCrLf
                     ElseIf RsTemp.Fields("cp10") = "203" Then
                        strSubject = strSubject & "提申後主動修正"
                        strContent = strContent & "  提申後主動修正　　承辦期限：" & ChangeWStringToTDateString("" & RsTemp.Fields("cp48")) & "　本所期限：" & ChangeWStringToTDateString("" & RsTemp.Fields("cp06")) & vbCrLf
                     End If
                     RsTemp.MoveNext
                  Loop
                  'Modified by Lydia 2022/09/23 請將(Murgitroyd案優先)移至主旨前面，以便承辦可以快速辨認(by Bobbie); 判斷改用模組
                  'strSubject = strSubject & "】Our Ref: " & pa(1) & "-" & pa(2) & " [INCOM." & strCP10 & "]" & IIf(Left(ChangeCustomerL(pa(75)), 8) = "Y2099001", "(Murgitroyd案優先)", "")
                  strSubject = strSubject & "】Our Ref: " & pa(1) & "-" & pa(2) & " [INCOM." & strCP10 & "]"
                  strSubject = PUB_GetSetMailSubF2(pa(75)) & strSubject
                  'end 2022/09/23
                  
                  PUB_SendMail strUserNum, strTo, "", strSubject, strContent, "", "", , , , strCC
               End If
            End If
            '2022/5/11 END
            Screen.MousePointer = vbDefault
            
            'Add By Sindy 2022/5/11
            If txtEmail = "Y" Then
               'Modify By Sindy 2022/5/20 and cp27=" & strSrvDate(1) & "
               strExc(0) = "SELECT cp09 FROM caseprogress" & _
                           " WHERE cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                           " and cp10 in('101','102','103','125','307','308')" & _
                           " and cp57 is null" & _
                           " order by cp27 desc"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  RsTemp.MoveFirst
                  frm060104_k.m_CP09 = RsTemp.Fields("cp09")
                  '2022/5/20 END
                  frm060104_k.m_strRecDate = txtRecDate
                  frm060104_k.Hide
                  frm060104_k.cmdOK(0) = 1
                  Unload frm060104_k
               End If
            End If
            '2022/5/11 END
            
            'Add By Sindy 2022/5/11
            If Me.m_strIR01 <> "" Then
               Unload frm060105_1
               If Not m_PrevForm Is Nothing Then
                  Call m_PrevForm.GoNext
               End If
               Unload Me
            Else
            '2022/5/11 END
               'Modify By Cheng 2002/10/24
   '            PrtGreenPaper
               'Modify By Cheng 2002/12/11
               '保留原輸入的系統類別
   '            frm060105_1.Text1 = ""
               If ExceptReturn = False Then  'Added by Lydia 2019/06/25 從發文來的,跳回發文
                   frm060105_1.Text2 = ""
                   frm060105_1.Text3 = ""
                   frm060105_1.Text4 = ""
                   frm060105_1.Show
                   frm060105_1.Text1.SetFocus
                   'Add By Cheng 2002/12/11
                   frm060105_1.Text2.SetFocus
               End If
               Unload Me
            End If
         End If
         
      Case 1
         If ExceptReturn = False Then  'Added by Lydia 2019/06/25 從發文來的,跳回發文
              Unload frm060105_1
         End If
         Unload Me
      Case 2
         'Add By Sindy 2021/5/7
         If Text9(2) = "" Then
            MsgBox "約定期限不可空白 !", vbCritical
            Text9(2).SetFocus
            Exit Sub
         End If
         '2021/5/7 END
         If Text9(0) = "" Then
            MsgBox "本所期限不可空白 !", vbCritical
            Text9(0).SetFocus
            Exit Sub
         End If
         If Text9(1) = "" Then
            MsgBox "法定期限不可空白 !", vbCritical
            Text9(1).SetFocus
            Exit Sub
         End If
         If cp(4) = "" Or cp(5) = "" Then
            MsgBox "無資料更新 !", vbCritical
            Exit Sub
         End If
         'Modify By Sindy 2021/5/7
         strExc(1) = "UPDATE NEXTPROGRESS SET NP08=" & TransDate(Text9(0), 2) & ",NP09=" & _
            TransDate(Text9(1), 2) & ",NP23=" & TransDate(Text9(2), 2) & " WHERE NP01='" & cp(4) & "' AND NP07=" & 補文件 & " AND NP22=" & cp(5)
         'edit by nickc 2007/02/05 不用 dll 了
         'If Not objLawDll.ExecSQL(1, strExc) Then
         If Not ClsLawExecSQL(1, strExc) Then
            MsgBox "更新下一程序檔失敗，請洽系統管理者 !", vbCritical
         Else
            With MSHFlexGrid1
               'Add By Sindy 2021/5/7
               .col = 2
               .Text = Text9(2)
               '2021/5/7 END
               .col = 3 '2
               .Text = Text9(0)
               .col = 4 '3
               .Text = Text9(1)
            End With
            Text9(2) = "" 'Add By Sindy 2021/5/7
            Text9(0) = ""
            Text9(1) = ""
         End If
      Case 3
         Unload Me
         If ExceptReturn = False Then  'Added by Lydia 2019/06/25 從發文來的,跳回發文
            frm060105_1.Show
         End If
   End Select
End Sub

Private Function FormSave() As Boolean
Dim intStep As Integer
'Add By Cheng 2002/07/04
Dim intMax As Long
Dim i As Integer
'Add By Cheng 2002/10/31
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'Add By Cheng 2002/11/01
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim strMemo416 As String   '2011/10/12 add by sonia
'Add By Sindy 2017/3/3
Dim strPA79 As String, strPA80 As String, strPA81 As String
Dim strPA82 As String, strPA83 As String, strPA84 As String
Dim strPA109 As String, strPA110 As String, strPA111 As String
Dim strPA112 As String, strPA113 As String, strPA114 As String
Dim strPA115 As String, strPA116 As String, strPA117 As String
Dim strPA118 As String, strPA119 As String, strPA120 As String
Dim strPA121 As String, strPA122 As String, strPA123 As String
Dim strPA124 As String, strPA125 As String, strPA126 As String
Dim strPA127 As String, strPA128 As String, strPA129 As String
Dim strPA130 As String, strPA131 As String, strPA132 As String
'2017/3/3 END
   
   '911105 nick transation
   FormSave = True

On Error GoTo CheckingErr
cnnConnection.BeginTrans

   intStep = 1
   '2015/1/20 MODIFY BY SONIA
   'If Left(strKind, 1) = "3" Then
   If Left(strKind, 1) = "3" And strKind <> "307" Then
      strExc(intStep) = "UPDATE CASEPROGRESS SET CP30=" & CNULL(pa(11)) & " WHERE CP09='" & strReceiveNo & "'"
      
      '911105 nick transation
      cnnConnection.Execute strExc(intStep)
      
      intStep = intStep + 1
   End If
   'strExc(intStep) = "UPDATE PATENT SET PA10=" & TransDate(Text6, 2) & ",PA11=" & cnull(chgsql(TEXT7)) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
    'Modify By Cheng 2002/10/31
'   strExc(intStep) = "UPDATE PATENT SET PA10=" & TransDate(Text6, 2) & ",PA11=" & CNULL(ChgSQL(Text7)) & _
'      ",PA08=substr(" & CNULL(ChgSQL(Text7)) & ",3,1) WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   intI = 3
   'Modify by Morgan 2010/8/20 9碼格式則第4碼是專利種類
   If bolNewAppNoFormat Then
      intI = 4
   End If
   
   '2011/1/24 modify by sonia 積體電路不更新專利種類FCP-042803
   If strKind <> "117" Then
      strExc(intStep) = "UPDATE PATENT SET PA10=" & TransDate(Text6, 2) & ",PA11=" & CNULL(ChgSQL(Text7)) & _
         ",PA08=" & IIf(Me.Text7.Text = "", "PA08", "'" & Mid(Replace(Me.Text7.Text, "'", ""), intI, 1) & "' ") & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      cnnConnection.Execute strExc(intStep)
   Else
      strExc(intStep) = "UPDATE PATENT SET PA10=" & TransDate(Text6, 2) & ",PA11=" & CNULL(ChgSQL(Text7)) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      cnnConnection.Execute strExc(intStep)
   End If
      
   intStep = intStep + 1

   'strExc(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10," & _
   '   "CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43) VALUES ('" & pa(1) & "','" & pa(2) & "','" & _
   '   pa(3) & "','" & pa(4) & "'," & CNULL(TransDate(Text5, 2)) & "," & CNULL(Text8) & ",'" & _
   '   AutoNo("C", 6) & "'," & 通知申請案號 & "," & CNULL(cp(2)) & "," & _
   '   CNULL(cp(3)) & ",'" & strUserNum & "','N','N','N'," & strSrvDate(1) & ",'" & cp(1) & "')"
   ' 90.06.27 modify by louis 發文日即原申請日
   '2012/10/2 MODIFY BY SONIA 智權人員抓最新的,業務區抓新智權人員的業務區
   'strExc(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10," & _
      "CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43) VALUES ('" & pa(1) & "','" & pa(2) & "','" & _
      pa(3) & "','" & pa(4) & "'," & CNULL(TransDate(Text5, 2)) & "," & CNULL(Text8) & ",'" & _
      AutoNo("C", 6) & "','" & 通知申請案號 & "'," & CNULL(cp(2)) & "," & _
      CNULL(cp(3)) & ",'" & strUserNum & "','N','N','N'," & CNULL(TransDate(Text5, 2)) & ",'" & cp(1) & "')"
   strExc(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10," & _
      "CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43) VALUES ('" & pa(1) & "','" & pa(2) & "','" & _
      pa(3) & "','" & pa(4) & "'," & CNULL(TransDate(Text5, 2)) & "," & CNULL(Text8) & ",'" & _
      AutoNo("C", 6) & "','" & 通知申請案號 & "'," & CNULL(GetSalesArea(PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)))) & "," & _
      CNULL(PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))) & ",'" & strUserNum & "','N','N','N'," & CNULL(TransDate(Text5, 2)) & ",'" & cp(1) & "')"
      
      '911105 nick transation
      cnnConnection.Execute strExc(intStep)
      
   intStep = intStep + 1
   
   'Add By Cheng 2002/07/04
   intMax = GetNextProgressNo  'edit by nickc 2007/02/02 不用 dll 了  intMax = objPublicData.GetNextProgressNo
   '若專利種類為"發明"且"申請日期">=91/10/26
   If pa(8) = "1" And (DBDATE(Text6.Text) - 19110000) >= 911026 Then
      '若案件性質為"發明申請"(101)
      If strKind = "101" Then
         Dim strTmp1A(0 To 4) As String, strTmpA(1 To 3) As String
         strTmp1A(0) = strReceiveNo
         For i = 1 To 4
            strTmp1A(i) = pa(i)
         Next
         If GetMoneyDate(Val(Mid(strKind, 3, 1) + 3), pa(9), strTmp1A, strTmpA(1), strTmpA(2), strTmpA(3)) = True Then
            '法定期限
            If strTmpA(3) <> "" Then
                'Modify By Cheng 2002/10/29
                '法定期限應為止日加一天
               strTmpA(3) = CompDate(2, 1, strTmpA(3))
               
               'Modified by Morgan 2014/11/20 外專改回舊規則
               ''Added by Morgan 2014/10/29
               'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
               '   strTmpA(2) = PUB_GetOurDeadline(strTmpA(3))
               'Else
               
               'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
               If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
                  'Modify By Sindy 2021/4/27 + m_pAgreeOnDate
                  strTmpA(2) = PUB_GetFCPOurDeadline(strTmpA(3), 4, , m_pAgreeOnDate)
               Else
               'end 2019/7/11
         
                  '本所期限 = 法定期限 - 4天
                  strTmpA(2) = CompDate(2, -4, strTmpA(3))
                  
               End If 'Added by Morgan 2019/7/11
                  
               'End If 'Added by Morgan 2014/10/29
               'end 2014/11/20
               
                If rsB.State <> adStateClosed Then rsB.Close
                Set rsB = Nothing
                StrSqlB = "Select * From CaseProgress Where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " And CP10='" & 實體審查 & "' And CP05 IS NOT NULL And CP57 IS NULL "
                rsB.CursorLocation = adUseClient
                rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
                '若案件進度檔無實體審查的資料(即未收文)
                If rsB.RecordCount <= 0 Then
                    '2011/10/12 add by sonia改用PUB_GetMemo
                     strMemo416 = ""
                     'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
                     'For intI = 26 To 30
                     '   If Not IsNull(pa(intI)) Then
                     '      'Modified by Morgan 2013/9/11 改抓設定檔
                     '      'strMemo416 = PUB_Get416Memo(ChangeCustomerL(pa(intI)), ChangeCustomerL(pa(75)))
                     '      strMemo416 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL(pa(75)), ChangeCustomerL(pa(intI)))
                     '      If strMemo416 <> "" Then Exit For
                     '   End If
                     'Next
                     strMemo416 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL(pa(75)), pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30))
                     'end 2022/08/02
                    '2011/10/12 end
                    'Add/Modify By Cheng 2002/10/31
                    If rsA.State <> adStateClosed Then rsA.Close
                    Set rsA = Nothing
                    'Modified by Morgan 2011/11/17 其他條件也要加否則還是會錯 Ex.NP22=815745(P-086124,FCP-044648)
                    'StrSQLa = "Select NP01,NP07,NP22 From Nextprogress Where NP22= (SELECT MAX(NP22) FROM NEXTPROGRESS WHERE NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP07=" & 實體審查 & " AND NP06 IS NULL )"
                    'Modified by Lydia 2022/08/20 +NP15
                    StrSQLa = "Select NP01,NP07,NP22,NP15 From Nextprogress Where NP22= (SELECT MAX(NP22) FROM NEXTPROGRESS WHERE NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP07=" & 實體審查 & " AND NP06 IS NULL ) and NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP07=" & 實體審查 & " AND NP06 IS NULL"
                    rsA.CursorLocation = adUseClient
                    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                    '若有資料
                    If rsA.RecordCount > 0 Then
'2011/10/12 modify by sonia 改用PUB_GetMemo
'                       '2008/10/13 modify by sonia
'                       'strExc(intStep) = "Update NEXTPROGRESS SET NP08=" & strTmpA(2) & ",NP09=" & strTmpA(3) & " WHERE  NP01='" & rsA.Fields(0).Value & "' AND NP07='" & rsA.Fields(1).Value & "' AND NP22=" & rsA.Fields(2).Value
'                       '2009/2/6 MODIFY BY SONIA 加入X20438000,X21775000
'                       '2009/5/4 modify by sonia 加入Y51304020
'                       'If ChangeCustomerL(pa(26)) = "X55560000" Or ChangeCustomerL(pa(26)) = "X55560010" Or ChangeCustomerL(pa(27)) = "X55560000" Or ChangeCustomerL(pa(27)) = "X55560010" Or ChangeCustomerL(pa(28)) = "X55560000" Or ChangeCustomerL(pa(28)) = "X55560010" Or ChangeCustomerL(pa(29)) = "X55560000" Or ChangeCustomerL(pa(29)) = "X55560010" Or ChangeCustomerL(pa(30)) = "X55560000" Or ChangeCustomerL(pa(30)) = "X55560010" Then
'                       If ChangeCustomerL(pa(26)) = "X55560000" Or ChangeCustomerL(pa(27)) = "X55560000" Or ChangeCustomerL(pa(28)) = "X55560000" Or ChangeCustomerL(pa(29)) = "X55560000" Or ChangeCustomerL(pa(30)) = "X55560000" Or _
'                          ChangeCustomerL(pa(26)) = "X55560010" Or ChangeCustomerL(pa(27)) = "X55560010" Or ChangeCustomerL(pa(28)) = "X55560010" Or ChangeCustomerL(pa(29)) = "X55560010" Or ChangeCustomerL(pa(30)) = "X55560010" Or _
'                          ChangeCustomerL(pa(26)) = "X20438000" Or ChangeCustomerL(pa(27)) = "X20438000" Or ChangeCustomerL(pa(28)) = "X20438000" Or ChangeCustomerL(pa(29)) = "X20438000" Or ChangeCustomerL(pa(30)) = "X20438000" Or _
'                          ChangeCustomerL(pa(26)) = "X21775000" Or ChangeCustomerL(pa(27)) = "X21775000" Or ChangeCustomerL(pa(28)) = "X21775000" Or ChangeCustomerL(pa(29)) = "X21775000" Or ChangeCustomerL(pa(30)) = "X21775000" Or _
'                          ChangeCustomerL(pa(75)) = "Y51304020" Then
'                           strExc(intStep) = "Update NEXTPROGRESS SET NP08=" & strTmpA(2) & ",NP09=" & strTmpA(3) & ",NP15='不寄函逕收文;'||NP15 WHERE NP01='" & rsA.Fields(0).Value & "' AND NP07='" & rsA.Fields(1).Value & "' AND NP22=" & rsA.Fields(2).Value
'                       Else
'                           strExc(intStep) = "Update NEXTPROGRESS SET NP08=" & strTmpA(2) & ",NP09=" & strTmpA(3) & " WHERE  NP01='" & rsA.Fields(0).Value & "' AND NP07='" & rsA.Fields(1).Value & "' AND NP22=" & rsA.Fields(2).Value
'                       End If
'                       '2008/10/13 END
                       '2011/10/12 同時更新智權人員
                       'Modify By Sindy 2021/4/27 + ,NP23=" & CNULL(DBDATE(m_pAgreeOnDate)):約定期限
                       'Modified by Lydia 2022/08/02 判斷備註不存在才加註
                       'strExc(intStep) = "Update NEXTPROGRESS SET NP08=" & Val(strTmpA(2)) & ",NP09=" & Val(strTmpA(3)) & ",NP23=" & CNULL(DBDATE(m_pAgreeOnDate)) & ",NP10='" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "',np15=decode(NP15,null,'" & strMemo416 & "','" & strMemo416 & "'||NP15) WHERE NP01='" & rsA.Fields(0).Value & "' AND NP07='" & rsA.Fields(1).Value & "' AND NP22=" & rsA.Fields(2).Value
                       StrSQLa = ""
                       If InStr(rsA.Fields("np15") & ";", strMemo416) = 0 And strMemo416 <> "" Then
                           StrSQLa = ", NP15='" & ChgSQL(strMemo416) & IIf("" & rsA.Fields("np15") <> "", ";", "") & "'||NP15 "
                       End If
                       strExc(intStep) = "Update NEXTPROGRESS SET NP08=" & Val(strTmpA(2)) & ",NP09=" & Val(strTmpA(3)) & ",NP23=" & CNULL(DBDATE(m_pAgreeOnDate)) & ",NP10='" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "' " & StrSQLa & " WHERE NP01='" & rsA.Fields(0).Value & "' AND NP07='" & rsA.Fields(1).Value & "' AND NP22=" & rsA.Fields(2).Value
                       'end 2022/08/02
'2011/10/12 end
                       
                       '911105 nick transation
                       cnnConnection.Execute strExc(intStep)
                       intStep = intStep + 1
                    '若無資料
                    Else
                        '911105 nick
                        intMax = GetNextProgressNo  'edit by nickc 2007/02/02 不用 dll 了  intMax = objPublicData.GetNextProgressNo
'2011/10/12 modify by sonia 改用PUB_GetMemo
'                        '2008/10/13 modify by sonia
'                        'strExc(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
'                           "NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & _
'                           pa(4) & "'," & 實體審查 & "," & CNULL(strTmpA(2)) & "," & CNULL(strTmpA(3)) & ",'" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & intMax & ")"
'                        'If ChangeCustomerL(pa(26)) = "X55560000" Or ChangeCustomerL(pa(26)) = "X55560010" Or ChangeCustomerL(pa(27)) = "X55560000" Or ChangeCustomerL(pa(27)) = "X55560010" Or ChangeCustomerL(pa(28)) = "X55560000" Or ChangeCustomerL(pa(28)) = "X55560010" Or ChangeCustomerL(pa(29)) = "X55560000" Or ChangeCustomerL(pa(29)) = "X55560010" Or ChangeCustomerL(pa(30)) = "X55560000" Or ChangeCustomerL(pa(30)) = "X55560010" Then
'                        If ChangeCustomerL(pa(26)) = "" Or ChangeCustomerL(pa(27)) = "X55560000" Or ChangeCustomerL(pa(28)) = "X55560000" Or ChangeCustomerL(pa(29)) = "X55560000" Or ChangeCustomerL(pa(30)) = "X55560000" Or _
'                           ChangeCustomerL(pa(26)) = "X55560010" Or ChangeCustomerL(pa(27)) = "X55560010" Or ChangeCustomerL(pa(28)) = "X55560010" Or ChangeCustomerL(pa(29)) = "X55560010" Or ChangeCustomerL(pa(30)) = "X55560010" Or _
'                           ChangeCustomerL(pa(26)) = "X20438000" Or ChangeCustomerL(pa(27)) = "X20438000" Or ChangeCustomerL(pa(28)) = "X20438000" Or ChangeCustomerL(pa(29)) = "X20438000" Or ChangeCustomerL(pa(30)) = "X20438000" Or _
'                           ChangeCustomerL(pa(26)) = "X21775000" Or ChangeCustomerL(pa(27)) = "X21775000" Or ChangeCustomerL(pa(28)) = "X21775000" Or ChangeCustomerL(pa(29)) = "X21775000" Or ChangeCustomerL(pa(30)) = "X21775000" Or _
'                           ChangeCustomerL(pa(75)) = "Y51304020" Then
'                        '2009/2/6 MODIFY BY SONIA 加入XX5556000020438000,X21775000
'                        '2009/5/4 modify by sonia 加入Y51304020
'                           strExc(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
'                              "NP09,NP10,NP15,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & _
'                              pa(4) & "'," & 實體審查 & "," & CNULL(strTmpA(2)) & "," & CNULL(strTmpA(3)) & ",'" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "','不寄函逕收文;'," & intMax & ")"
'                        Else
'                           strExc(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
'                              "NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & _
'                              pa(4) & "'," & 實體審查 & "," & CNULL(strTmpA(2)) & "," & CNULL(strTmpA(3)) & ",'" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & intMax & ")"
'                        End If
'                        '2008/10/13 END
                        'Modify By Sindy 2021/4/27 + ,NP23=" & CNULL(DBDATE(m_pAgreeOnDate)):約定期限
                        strExc(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
                           "NP09,NP10,NP15,NP22,NP23) VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & _
                           pa(4) & "'," & 實體審查 & "," & CNULL(strTmpA(2)) & "," & CNULL(strTmpA(3)) & ",'" & _
                           PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "','" & strMemo416 & "'," & intMax & "," & CNULL(DBDATE(m_pAgreeOnDate)) & ")"
'2011/10/12 end
                           
                        '911105 nick transation
                        cnnConnection.Execute strExc(intStep)
                        intStep = intStep + 1
                        intMax = intMax + 1
                    End If
                    If rsA.State <> adStateClosed Then rsA.Close
                    Set rsA = Nothing
                '若案件進度檔有未發文的實體審查資料
                ElseIf "" & rsB("CP27") = "" Then
                    'Add By Cheng 2002/11/04
                    MsgBox "有實審收文尚未發文!!!", vbExclamation + vbOKOnly
                    'Add By Cheng 2002/11/01
                    '更新案件進度檔未發文實體審查的期限資料
                    strExc(intStep) = "Update CaseProgress Set CP06=" & Val(strTmpA(2)) & ",CP07=" & Val(strTmpA(3)) & " Where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " And CP10='" & 實體審查 & "' And CP27 IS NULL "
                    
                    '911105 nick transation
                    cnnConnection.Execute strExc(intStep)
                    
                    intStep = intStep + 1
                End If
                If rsB.State <> adStateClosed Then rsB.Close
                Set rsB = Nothing
            End If
         End If
      'Remove by Morgan 2004/8/6
      '改由發文控制
'
'      '若案件性質為"改請發明"(301)
'      ElseIf strKind = "301" Then
'        'Modify By Cheng 2002/10/31
''         Dim rsA As New ADODB.Recordset
''         Dim strSQLA As String
'         Dim strNP08 As String '下一程序本所期限
'         Dim strNP09 As String '下一程序法定期限
'         '法定期限 = 畫面上申請日期 + 30 天
'         strNP09 = DBDATE(DateSerial(DBYEAR(Text6.Text), DBMONTH(Text6.Text), DBDAY(Text6.Text) + 30))
'         '本所期限 = 法定期限 - 4 天
'         strNP08 = DBDATE(DateSerial(DBYEAR(strNP09), DBMONTH(strNP09), DBDAY(strNP09) - 4))
'        If rsB.State <> adStateClosed Then rsB.Close
'        Set rsB = Nothing
'        strSQLB = "Select * From CaseProgress Where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " And CP10='" & 實體審查 & "' And CP05 IS NOT NULL And CP57 IS NULL "
'        rsB.CursorLocation = adUseClient
'        rsB.Open strSQLB, cnnConnection, adOpenStatic, adLockReadOnly
'        '若案件進度檔無實體審查的資料(即未收文)
'        If rsB.RecordCount <= 0 Then
'             If rsA.State <> adStateClosed Then rsA.Close
'             Set rsA = Nothing
'    '         strSQLA = "Select NP22 From Nextprogress Where NP22= (SELECT MAX(NP22) FROM NEXTPROGRESS WHERE NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP07=" & 實體審查 & " AND NP06 IS NULL ) "
'             strSQLA = "Select NP01,NP07,NP22 From Nextprogress Where NP22= (SELECT MAX(NP22) FROM NEXTPROGRESS WHERE NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP07=" & 實體審查 & " AND NP06 IS NULL ) "
'             rsA.CursorLocation = adUseClient
'             rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'             '若有資料
'             If rsA.RecordCount > 0 Then
'                strExc(intStep) = "Update NEXTPROGRESS SET NP08=" & strNP08 & ",NP09=" & strNP09 & " WHERE NP01='" & rsA.Fields(0).Value & "' AND NP07='" & rsA.Fields(1).Value & "' AND NP22=" & rsA.Fields(2).Value
'
'                '911105 nick transation
'                cnnConnection.Execute strExc(intStep)
'
'                intStep = intStep + 1
'             '若無資料
'             Else
'                '911105 nick
'                intMax = objPublicData.GetNextProgressNo
'
'                strExc(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
'                   "NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & _
'                   pa(4) & "'," & 實體審查 & "," & CNULL(strNP08) & "," & CNULL(strNP09) & ",'" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & intMax & ")"
'
'                '911105 nick transation
'                cnnConnection.Execute strExc(intStep)
'
'                intStep = intStep + 1
'                intMax = intMax + 1
'             End If
'             If rsA.State <> adStateClosed Then rsA.Close
'             Set rsA = Nothing
'        '若案件進度檔有未發文的實體審查資料
'        ElseIf "" & rsB("CP27").Value = "" Then
'            'Add By Cheng 2002/11/04
'            MsgBox "有實審收文尚未發文!!!", vbExclamation + vbOKOnly
'            'Add By Cheng 2002/11/01
'            '更新案件進度檔未發文實體審查的期限資料
'            strExc(intStep) = "Update CaseProgress Set CP06=" & Val(strNP08) & ",CP07=" & Val(strNP09) & " Where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " And CP10='" & 實體審查 & "' And CP27 IS NULL "
'
'            '911105 nick transation
'            cnnConnection.Execute strExc(intStep)
'
'            intStep = intStep + 1
'
'        End If
'        If rsB.State <> adStateClosed Then rsB.Close
'        Set rsB = Nothing
      End If
   End If
   
   'Modify By Cheng 2002/07/04
'   FormSave = objLawDll.ExecSQL(intStep, strExc)
    '911105 nick
   'FormSave = objLawDll.ExecSQL(intStep - 1, strExc)
   
   
   'Modify By Sindy 2018/5/11 Mark:改到分案作業處理
'   'Add by Morgan 2005/4/21 將母案的優先權資料新增到分割案
'   'Modify by Amy 2014/03/26 +pd08,pd09
'   If strKind = "307" Then
'      strSql = "INSERT INTO PRIDATE A (PD01,PD02,PD03,PD04,PD05,PD06,PD07,PD08,PD09)" & _
'         " select DC01,DC02,DC03,DC04,B.PD05,B.PD06,B.PD07,PD08,PD09 from divisioncase,pridate B" & _
'         " where dc01='" & pa(1) & "' and dc02='" & pa(2) & "' and dc03='" & pa(3) & "' and dc04='" & pa(4) & "'" & _
'         " AND B.pd01=dc05 and B.pd02=dc06 and B.pd03=dc07 and B.pd04=dc08" & _
'         " AND NOT EXISTS(SELECT * FROM PRIDATE C WHERE C.pd01=dc01 and C.pd02=dc02 and C.pd03=dc03 and C.pd04=dc04 and c.pd06=B.pd06 and c.pd07=B.pd07)"
'      cnnConnection.Execute strSql
'      'Add By Sindy 2017/3/3
'      '將母案的發明人資料新增到分割案
'      strSql = "INSERT INTO PatentInventor A (PI01,PI02,PI03,PI04,PI05,PI06)" & _
'         " select DC01,DC02,DC03,DC04,B.PI05,B.PI06 from divisioncase,PatentInventor B" & _
'         " where dc01='" & pa(1) & "' and dc02='" & pa(2) & "' and dc03='" & pa(3) & "' and dc04='" & pa(4) & "'" & _
'         " AND B.PI01=dc05 and B.PI02=dc06 and B.PI03=dc07 and B.PI04=dc08" & _
'         " AND NOT EXISTS(SELECT * FROM PatentInventor C WHERE C.PI01=dc01 and C.PI02=dc02 and C.PI03=dc03 and C.PI04=dc04 and c.PI05=B.PI05)"
'      Pub_SeekTbLog strSql 'Add By Sindy 2017/8/23
'      cnnConnection.Execute strSql
'      '將母案的代表人資料新增到分割案
'      If rsB.State <> adStateClosed Then rsB.Close
'      Set rsB = Nothing
'      '讀取母案的代表人資料
'      StrSqlB = " select * from divisioncase,Patent" & _
'         " where dc01='" & pa(1) & "' and dc02='" & pa(2) & "' and dc03='" & pa(3) & "' and dc04='" & pa(4) & "'" & _
'         " AND PA01=dc05 and PA02=dc06 and PA03=dc07 and PA04=dc08"
'      rsB.CursorLocation = adUseClient
'      rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsB.RecordCount > 0 Then
'         strPA79 = "" & rsB.Fields("pa79")
'         strPA80 = "" & rsB.Fields("pa80")
'         strPA81 = "" & rsB.Fields("pa81")
'         strPA82 = "" & rsB.Fields("pa82")
'         strPA83 = "" & rsB.Fields("pa83")
'         strPA84 = "" & rsB.Fields("pa84")
'         strPA109 = "" & rsB.Fields("pa109")
'         strPA110 = "" & rsB.Fields("pa110")
'         strPA111 = "" & rsB.Fields("pa111")
'         strPA112 = "" & rsB.Fields("pa112")
'         strPA113 = "" & rsB.Fields("pa113")
'         strPA114 = "" & rsB.Fields("pa114")
'         strPA115 = "" & rsB.Fields("pa115")
'         strPA116 = "" & rsB.Fields("pa116")
'         strPA117 = "" & rsB.Fields("pa117")
'         strPA118 = "" & rsB.Fields("pa118")
'         strPA119 = "" & rsB.Fields("pa119")
'         strPA120 = "" & rsB.Fields("pa120")
'         strPA121 = "" & rsB.Fields("pa121")
'         strPA122 = "" & rsB.Fields("pa122")
'         strPA123 = "" & rsB.Fields("pa123")
'         strPA124 = "" & rsB.Fields("pa124")
'         strPA125 = "" & rsB.Fields("pa125")
'         strPA126 = "" & rsB.Fields("pa126")
'         strPA127 = "" & rsB.Fields("pa127")
'         strPA128 = "" & rsB.Fields("pa128")
'         strPA129 = "" & rsB.Fields("pa129")
'         strPA130 = "" & rsB.Fields("pa130")
'         strPA131 = "" & rsB.Fields("pa131")
'         strPA132 = "" & rsB.Fields("pa132")
'      End If
'      strSql = ""
'      If strPA79 <> "" And pa(79) = "" Then strSql = strSql & ",pa79='" & ChgSQL(strPA79) & "'"
'      If strPA80 <> "" And pa(80) = "" Then strSql = strSql & ",pa80='" & ChgSQL(strPA80) & "'"
'      If strPA81 <> "" And pa(81) = "" Then strSql = strSql & ",pa81='" & ChgSQL(strPA81) & "'"
'      If strPA82 <> "" And pa(82) = "" Then strSql = strSql & ",pa82='" & ChgSQL(strPA82) & "'"
'      If strPA83 <> "" And pa(83) = "" Then strSql = strSql & ",pa83='" & ChgSQL(strPA83) & "'"
'      If strPA84 <> "" And pa(84) = "" Then strSql = strSql & ",pa84='" & ChgSQL(strPA84) & "'"
'      If strPA109 <> "" And pa(109) = "" Then strSql = strSql & ",pa109='" & ChgSQL(strPA109) & "'"
'      If strPA110 <> "" And pa(110) = "" Then strSql = strSql & ",pa110='" & ChgSQL(strPA110) & "'"
'      If strPA111 <> "" And pa(111) = "" Then strSql = strSql & ",pa111='" & ChgSQL(strPA111) & "'"
'      If strPA112 <> "" And pa(112) = "" Then strSql = strSql & ",pa112='" & ChgSQL(strPA112) & "'"
'      If strPA113 <> "" And pa(113) = "" Then strSql = strSql & ",pa113='" & ChgSQL(strPA113) & "'"
'      If strPA114 <> "" And pa(114) = "" Then strSql = strSql & ",pa114='" & ChgSQL(strPA114) & "'"
'      If strPA115 <> "" And pa(115) = "" Then strSql = strSql & ",pa115='" & ChgSQL(strPA115) & "'"
'      If strPA116 <> "" And pa(116) = "" Then strSql = strSql & ",pa116='" & ChgSQL(strPA116) & "'"
'      If strPA117 <> "" And pa(117) = "" Then strSql = strSql & ",pa117='" & ChgSQL(strPA117) & "'"
'      If strPA118 <> "" And pa(118) = "" Then strSql = strSql & ",pa118='" & ChgSQL(strPA118) & "'"
'      If strPA119 <> "" And pa(119) = "" Then strSql = strSql & ",pa119='" & ChgSQL(strPA119) & "'"
'      If strPA120 <> "" And pa(120) = "" Then strSql = strSql & ",pa120='" & ChgSQL(strPA120) & "'"
'      If strPA121 <> "" And pa(121) = "" Then strSql = strSql & ",pa121='" & ChgSQL(strPA121) & "'"
'      If strPA122 <> "" And pa(122) = "" Then strSql = strSql & ",pa122='" & ChgSQL(strPA122) & "'"
'      If strPA123 <> "" And pa(123) = "" Then strSql = strSql & ",pa123='" & ChgSQL(strPA123) & "'"
'      If strPA124 <> "" And pa(124) = "" Then strSql = strSql & ",pa124='" & ChgSQL(strPA124) & "'"
'      If strPA125 <> "" And pa(125) = "" Then strSql = strSql & ",pa125='" & ChgSQL(strPA125) & "'"
'      If strPA126 <> "" And pa(126) = "" Then strSql = strSql & ",pa126='" & ChgSQL(strPA126) & "'"
'      If strPA127 <> "" And pa(127) = "" Then strSql = strSql & ",pa127='" & ChgSQL(strPA127) & "'"
'      If strPA128 <> "" And pa(128) = "" Then strSql = strSql & ",pa128='" & ChgSQL(strPA128) & "'"
'      If strPA129 <> "" And pa(129) = "" Then strSql = strSql & ",pa129='" & ChgSQL(strPA129) & "'"
'      If strPA130 <> "" And pa(130) = "" Then strSql = strSql & ",pa130='" & ChgSQL(strPA130) & "'"
'      If strPA131 <> "" And pa(131) = "" Then strSql = strSql & ",pa131='" & ChgSQL(strPA131) & "'"
'      If strPA132 <> "" And pa(132) = "" Then strSql = strSql & ",pa132='" & ChgSQL(strPA132) & "'"
'      If strSql <> "" Then
'         strSql = Mid(strSql, 2)
'         strSql = "update patent set " & strSql & " where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
'         cnnConnection.Execute strSql
'      End If
'      '2017/3/3 END
'   End If
   
   'Add by Sindy 2022/5/11
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm060105_1"
   End If
   '2022/5/11 END
   
   cnnConnection.CommitTrans
'911105 nick
   Exit Function
CheckingErr:
   
   cnnConnection.RollbackTrans
   FormSave = False
   
End Function

Private Sub PrtGreenPaper()
 Dim Prn As Printer
 Dim i As Integer, j As Integer, k As Integer, lPos As Integer
 Dim varTmp As Variant
 Dim iLeft(1 To 3) As Integer
 Dim iTxtWidth As Integer

   For Each Prn In Printers
      If Prn.DeviceName = Combo2.Text Then
         Set Printer = Prn
         Exit For
      End If
   Next
   
   'Modify by Morgan 2008/4/9 9x才能自訂
   'Printer.Height = 2200
   'Printer.Width = 10000
   '9x
   If pub_OS = "1" Then
      Printer.Height = 2200
      Printer.Width = 10000
   'NT
   Else
      Printer.Orientation = 1
      Printer.EndDoc
   End If
   'end 2008/4/9
   
   Printer.Orientation = 1
   Printer.Font.Size = 14
   iLeft(1) = 1900
   iLeft(2) = 2600
   iTxtWidth = 10500 - 2600
   
   Printer.CurrentX = iLeft(1)
   Printer.CurrentY = 200
   Printer.Print "Title : "
   
   If Printer.TextWidth(pa(6)) > iTxtWidth Then
      varTmp = Split(pa(6), " ")
      strExc(0) = ""
      j = 0
      lPos = 0
      For i = 0 To UBound(varTmp)
         strExc(0) = strExc(0) & Format(varTmp(i)) & " "
         If Printer.TextWidth(strExc(0)) > iTxtWidth Then
            strExc(0) = ""
            For k = lPos To i - 1
               strExc(0) = strExc(0) & Format(varTmp(k)) & " "
            Next
            Printer.CurrentX = iLeft(2)
            Printer.CurrentY = 200 + j * 300
            Printer.Print strExc(0)
            strExc(0) = ""
            lPos = i
            j = j + 1
            i = i - 1
         End If
      Next
      
      strExc(0) = ""
      For i = lPos To UBound(varTmp)
         strExc(0) = strExc(0) & Format(varTmp(i)) & " "
      Next
      Printer.CurrentX = iLeft(2)
      Printer.CurrentY = 200 + j * 300
      Printer.Print strExc(0)
   
   Else
      Printer.CurrentX = iLeft(2)
      Printer.CurrentY = 200
      Printer.Print pa(6)
   End If
   
   Printer.CurrentX = iLeft(1)
   Printer.CurrentY = 1700
   Printer.Print "Application No : " & Text7
   
   Printer.CurrentX = iLeft(1)
   Printer.CurrentY = 2100
   Printer.Print "Filing Date : " & ChgEngDate(TransDate(Text6, 2))
   Printer.EndDoc
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(1) = pa(5)
      Case "英"
         Label2(1) = pa(6)
      Case "日"
         Label2(1) = pa(7)
   End Select
End Sub

Public Sub SetData(ByVal strNo As String)
   ChgCaseNo strNo, pa
End Sub

'Add by Morgan 2003/11/24
Private Sub Form_Activate()
   'Add By Sindy 2022/5/25
   If m_Done = False Then
   '2022/5/25 END
      If (Text6.Text <> "") Then Text7.SetFocus
      m_Done = True 'Add By Sindy 2022/5/25
   End If
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
   MoveFormToCenter Me
   intWhere = 國外_FC
   
   strExc(0) = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   j = 0
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      If Printer.DeviceName = strExc(0) Then
         SeekPrint = i
      Else
         Combo2.AddItem Printer.DeviceName, j
         j = j + 1
      End If
   Next i
   If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
   
   'Added by Lydia 2019/06/25 FCP特定案件性質的電子送件，發文確定後直接跳到"申請案號輸入"畫面，直接key 申請案號。
   If PubOtherCall <> "" Then
       mOtherForm = Mid(PubOtherCall, 1, InStr(PubOtherCall, ";") - 1)
       Call ChgCaseNo(Mid(PubOtherCall, InStr(PubOtherCall, ";") + 1), pa)
       PubOtherCall = ""
       Me.Text1.Text = pa(1)
       Me.Text2.Text = pa(2)
       Me.Text3.Text = pa(3)
       Me.Text4.Text = pa(4)
       Me.Text5.Text = strSrvDate(2)
   Else
   'end 2019/06/25
        With frm060105_1
           Text1.Text = .Text1
           Text2.Text = .Text2
           Text3.Text = .Text3
           Text4.Text = .Text4
           Text5.Text = .Text5
        End With
   End If
   
   'Add By Sindy 2021/5/7
   If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
      Text9(2).Visible = True
      Label12(3).Visible = True
   Else
      Text9(2).Visible = False
      Label12(3).Visible = False
   End If
   '2021/5/7 END
   
   InitGrid 10, MSHFlexGrid1 '9
   GridHead
   
   ReadPatent
   
   'Add By Sindy 2022/5/11
   m_strIR01 = frm060105_1.m_strIR01
   m_strIR02 = frm060105_1.m_strIR02
   m_strIR03 = frm060105_1.m_strIR03
   m_strIR04 = frm060105_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2022/5/11 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Printer = Printers(SeekPrint)
   Printer.Orientation = SeekPrintL
   
   'Add By Sindy 2022/5/11
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   '2022/5/11 END
   
   Set frm060105_2 = Nothing
End Sub

Private Sub ReadPatent()
Dim Lbl As Object, txt As TextBox, i As Integer
Dim strCP30 As String 'Added by Lydia 2020/07/13

   For Each Lbl In Label2
      Lbl = ""
   Next
   Combo1.ListIndex = 0
   Select Case pa(1)
      Case "FCP"
         If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetPatentTrademarkKind(專利, pA(8), strExc(1), False, 台灣國家代號) = 1 Then
            If ClsPDGetPatentTrademarkKind(專利, pa(8), strExc(1), False, 台灣國家代號) = 1 Then
               Label2(0) = strExc(1)
            End If
            Label2(1) = pa(5)
            If pa(26) <> "" Then
               'edit by nickc 2007/02/05 不用 dll 了
               'If objLawDll.LawGetName(pa(26), strExc(1)) Then
               If ClsLawLawGetName(pa(26), strExc(1)) Then
                  Label2(2) = strExc(1)
               End If
            End If
            Text6 = pa(10)
            Text7 = pa(11)
         End If
      Case "FG"
         'Modified by Lydia 2022/09/23 +SP26
         strExc(0) = "SELECT SP05,SP06,SP07,SP08,SP10,SP11,SP26 FROM SERVICEPRACTICE WHERE " & ChgService(pa(1) & pa(2) & pa(3) & pa(4))
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
               For i = 0 To 2
                  If Not IsNull(.Fields(i)) Then pa(i + 5) = .Fields(i)
               Next
               If .Fields(3) <> "" Then
                  pa(26) = .Fields(3)
                  'edit by nickc 2007/02/05 不用 dll 了
                  'If objLawDll.LawGetName(pa(26), strExc(1)) Then Label2(2) = strExc(1)
                  If ClsLawLawGetName(pa(26), strExc(1)) Then Label2(2) = strExc(1)
               End If
               If Not IsNull(.Fields(4)) Then Text6 = .Fields(4)
               If Not IsNull(.Fields(5)) Then
                  pa(11) = .Fields(5)
                  Text7 = pa(11)
               End If
               pa(75) = .Fields("SP26") 'Added by Lydia 2022/09/23 FC代理人
            End With
         End If
   End Select
   
   'Add by Morgan 2006/1/2
   Text7.Tag = Text7
   'Add By Sindy 2022/5/11
   strChk = ""
   If Text7.Tag = "" Then
      strChk = "第一次輸入"
      'txtEmail = "Y"
   End If
   txtEmail = "Y" 'Modify By Sindy 2023/5/26 因改請衍生設計之故,決定都預設為Y ex:FCP-68201
   '2022/5/11 END
   
   ' 90.06.27 modify by louis (暫存申請日)
   m_PA10 = pa(10)
   'strExc(0) = "select MAX(CP05),CP09,CP12,CP13,CP10 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
   '   " AND CP10 IN ('101','102','103','104','105','301','302','303','304','305','306') AND CP27 IS NOT NULL GROUP BY CP09,CP12,CP13,CP10"
   'Modify by Morgan 2005/4/21 加分割307
   '2007/8/13 modify by sonia 改請案加已發文條件
   'strExc(0) = "select CP05,CP09,CP12,CP13,CP10 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
   '   " AND ((CP10 IN ('101','102','103','104','105') AND CP27 IS NOT NULL) or (CP10 IN ('301','302','303','304','305','306','307'))) and CP05=( " & _
   '   "select MAX(CP05) FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
   '   " AND ((CP10 IN ('101','102','103','104','105') AND CP27 IS NOT NULL) or (CP10 IN ('301','302','303','304','305','306','307'))))"
   '2011/1/24 MODIFY BY SONIA 加積體電路不檢查FCP-042803
   'Modified by Morgan 2013/1/14 +308,309
   'Modified by Morgan 2016/3/15 +125
   'Modified by Morgan 2020/2/26 分割307及改請衍生設計308可能同時發文，此時抓308
   'strExc(0) = "select CP05,CP09,CP12,CP13,CP10 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " AND ((CP10 IN ('101','102','103','104','105','117','125') AND CP27 IS NOT NULL) or (CP10 IN ('301','302','303','304','305','306','307','308','309') AND CP27 IS NOT NULL)) and CP05=( " & _
      "select MAX(CP05) FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " AND ((CP10 IN ('101','102','103','104','105','117','125') AND CP27 IS NOT NULL) or (CP10 IN ('301','302','303','304','305','306','307','308','309') AND CP27 IS NOT NULL)))"
   'Modified by Lydia 2020/07/13 +衍生設計申請之母案案號CP30
   strExc(0) = "select CP05,CP09,CP12,CP13,CP10,CP30 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " AND ((CP10 IN ('101','102','103','104','105','117','125') AND CP27 IS NOT NULL) or (CP10 IN ('301','302','303','304','305','306','307','308','309') AND CP27 IS NOT NULL)) and CP27=( " & _
      "select MAX(CP27) FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " AND ((CP10 IN ('101','102','103','104','105','117','125') AND CP27 IS NOT NULL) or (CP10 IN ('301','302','303','304','305','306','307','308','309') AND CP27 IS NOT NULL)))" & _
      " order by cp10 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         For i = 1 To 3
            If Not IsNull(.Fields(i)) Then cp(i) = .Fields(i)
         Next
         ' 90.07.24 modify by louis (暫存收文號)
         strReceiveNo = .Fields("CP09")
         strKind = .Fields("CP10")
         strCP30 = "" & .Fields("CP30") 'Added by Lydia 2020/07/13 衍生設計申請之母案案號
      End With
   End If
   
   'Added by Lydia 2020/07/13 FCP 衍生設計: 預設母案申請案號(若本案無申請案號時)
   'Modified by Morgan 2023/4/25 改請衍生設計也要帶母案申請案號
   'If Trim(Text7.Text) = "" And pa(1) = "FCP" And strKind = "125" And strCP30 <> "" Then
   If pa(1) = "FCP" And strCP30 <> "" And ((Trim(Text7.Text) = "" And strKind = "125") Or strKind = "308") Then
   'end 2023/4/25
      Call ChgCaseNo(strCP30, strExc)
      If strExc(1) = "FCP" And Len(strExc(2)) = 6 Then
         strExc(0) = "select pa11 from patent where pa01='" & strExc(1) & "' and pa02='" & strExc(2) & "' and pa03='" & strExc(3) & "' and pa04='" & strExc(4) & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
             If "" & RsTemp.Fields("pa11") <> "" Then
                  Text7.Text = "" & RsTemp.Fields("pa11")
                  Text7.Tag = Text7.Text
             End If
         End If
      End If
   End If
   'end 2020/07/13
   
   'Modify By Sindy 2021/5/7 + ,DECODE(NP23,NULL,'',NP23-19110000)
   strExc(0) = "SELECT '','補文件',DECODE(NP23,NULL,'',NP23-19110000),DECODE(NP08,NULL,'',NP08-19110000),DECODE(NP09,NULL,'',NP09-19110000)," & _
      "NP13,NP14,NP15,NP01,NP22 FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " AND NP07=" & 補文件 & " AND NP06 IS NULL"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   
   'Add by Morgan 2005/4/21
   If Text6 = "" Then
      If strKind = "307" Then
         Text6 = TransDate(PUB_DivAppDate(pa(1), pa(2), pa(3), pa(4), True), 1)
         
      'Added by Morgan 2020/2/26
      ElseIf strKind = "308" And PUB_ChkCPExist(pa, "307", 2) Then
         Text6 = TransDate(PUB_DivAppDate(pa(1), pa(2), pa(3), pa(4), True), 1)
      End If
   End If
   'Modified by Morgan 2013/12/20 +308
   If strKind >= "301" And strKind <= "308" Then
      Text6.Enabled = False
   End If
End Sub

Private Sub MSHFlexGrid1_Click()
   With MSHFlexGrid1
      'Add By Sindy 2021/5/7
      .col = 2
      Text9(2) = .Text
      '2021/5/7 END
      .col = 3 '2
      Text9(0) = .Text
      .col = 4 '3
      Text9(1) = .Text
      .col = 8 '7
      cp(4) = .Text
      .col = 9 '8
      cp(5) = .Text
   End With
   GridClick MSHFlexGrid1, intLastRow, 0
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 = "" Then
      MsgBox "申請日不可空白，請重新輸入 !", vbCritical
      Cancel = True
   Else
      If ChkDate(Text6) Then
         If Val(Text6) > Val(strSrvDate(2)) Then
            MsgBox "申請日不可大於系統日 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
End Sub

Private Sub Text7_GotFocus()
   InverseTextBox Text7
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text7.IMEMode = 2
   CloseIme
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If Text7.Text <> "" Then
      '2005/5/20 ADD BY SONIA
      'Modify by Morgan 2010/8/20 9碼格式則前3碼是申請年
      'If Len(Me.Text7.Text) = 8 And Left("" & Me.Text7.Text, 2) <> Left("" & Me.Text6.Text, 2) And Left(strKind, 1) <> "3" Then
      '   MsgBox "申請案號的前二碼必須為申請年度 !", vbCritical
      'Modified by Morgan 2016/3/15 +衍生設計
      If strKind <> "104" And strKind <> "105" And strKind <> "125" Then
         If bolNewAppNoFormat Then
            strExc(1) = Val(Left(Text7, 3))
            strExc(3) = "三"
         Else
            strExc(1) = Val(Left(Text7, 2))
            strExc(3) = "二"
         End If
         strExc(2) = Trim(Val(Text6) \ 10000)
         If strExc(1) <> strExc(2) And Left(strKind, 1) <> "3" Then
            MsgBox "申請案號的前" & strExc(3) & "碼必須為申請年度 !", vbCritical
         'end 2010/8/20
            Cancel = True
            Exit Sub
         End If
      End If
      '2005/5/20 END
      
      '2005/6/14 MODIFY BY SONIA
      'If Not ChkAppNo(Text7.Text, pa(8), 0) Then
      If strKind <> "117" Then  '2011/1/24 MODIFY BY SONIA 積體電路不檢查FCP-042803
         If Not ChkAppNo(Text7.Text, pa(8), 0, Val(pa(23))) Then
      '2005/6/14 END
            Cancel = True
         End If
      End If    '2011/1/24 end
      
      'Added by Morgan 2012/8/21
      If Cancel = False Then
         If PUB_ChkAppNo(Text7, pa(1), pa(2), pa(9)) = False Then
            Cancel = True
            Exit Sub
         End If
      End If
      'end 2012/8/21
   End If
End Sub

Private Sub Text8_GotFocus()
   InverseTextBox Text8
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .col = 1: .ColWidth(1) = 1500: .Text = "下一程序"
      .CellAlignment = flexAlignCenterCenter
      'Add By Sindy 2021/5/7
      .col = 2: .ColWidth(2) = 1200: .Text = "約定期限"
      .CellAlignment = flexAlignCenterCenter
      '2021/5/7 END
      .col = 3: .ColWidth(3) = 1200: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1200: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1500: .Text = "機關文件"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1500: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 4000: .Text = "備註"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(8) = 0
      .col = 9: .ColWidth(9) = 0
      .Visible = True
   End With
End Sub

Private Sub Text9_GotFocus(Index As Integer)
   TextInverse Text9(Index)
End Sub

Private Sub Text9_LostFocus(Index As Integer)
   If Index = 1 And Text9(0) <> "" Then
      If Val(Text9(0)) > Val(Text9(1)) Then
         MsgBox "範圍錯誤，請重新輸入 !", vbCritical
         Text9(0).SetFocus
      End If
   End If
End Sub

Private Sub Text9_Validate(Index As Integer, Cancel As Boolean)
   If Text9(Index) <> "" Then Cancel = Not ChkDate(Text9(Index))
   If Cancel Then TextInverse Text9(Index)
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

If Me.Text6.Enabled = True Then
   Cancel = False
   Text6_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'Add by Morgan 2004/11/19
If Me.Text7.Enabled = True Then
   Cancel = False
   Text7_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add by Morgan 2006/1/2
If Text7.Tag <> "" And Text7 <> Text7.Tag Then
   If MsgBox("申請案號已更改【" & Text7.Tag & "->" & Text7 & "】,確定要繼續?", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
      Text7_GotFocus
      Text7.SetFocus
      Exit Function
   End If
End If

TxtValidate = True
End Function

'Removed by Morgan 2018/8/27 沒再用標為註解
'Private Sub StartLetter(ByVal ET01 As String, ET02 As String, ByVal ET03 As String)
'   Dim strTxt() As String, i As Integer
'   Dim strDoc As String, strList As String
'
'   EndLetter ET01, ET02, ET03, strUserNum
'
'   i = 0
'   If Text7 <> "" Then
'      If PUB_GetEMailFlag(pa(1) & pa(2) & pa(3) & pa(4)) Then
'         'Added by Morgan 2012/3/15 其他關係企業編號有需要另行告知--陳怡君
'         strExc(1) = ChangeCustomerL(pa(75))
'         If InStr("Y48309000,Y48309010,Y48309020,Y48309030,Y48309040,Y48309050,Y51326000", strExc(1)) > 0 Then
'            strExc(1) = "有申請號-E-1"
'         Else
'         'End 2012/3/15
'            strExc(1) = "有申請號-E"
'         End If
'      Else
'         strExc(1) = "有申請號-非E"
'      End If
'   Else
'      strExc(1) = "無申請號"
'   End If
'   i = i + 1
'   ReDim Preserve strTxt(i)
'   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'      "','" & strExc(1) & "','♀')"
'
'
'   If pa(8) = "1" Then
'      i = i + 1
'      ReDim Preserve strTxt(i)
'      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'         "','發明案','♀')"
'
'      strExc(0) = "select cp27 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "'" & _
'         " and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='416' and cp57 is null"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         i = i + 1
'         ReDim Preserve strTxt(i)
'         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'            "','實審有收文','♀')"
'
'         If RsTemp(0) > 0 Then
'            i = i + 1
'            ReDim Preserve strTxt(i)
'            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'               "','實審已發文','♀')"
'         End If
'      Else
'         strExc(0) = "select np08,np09 from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "'" & _
'            " and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np06||np07='416' order by np08"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            i = i + 1
'            ReDim Preserve strTxt(i)
'            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'               "','實審本所期限','" & RsTemp(0) & "')"
'
'            i = i + 1
'            ReDim Preserve strTxt(i)
'            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'               "','實審法定期限','" & RsTemp(1) & "')"
'         End If
'      End If
'   End If
'   'Modified by Morgan 2018/8/23 改例外欄位名:新式樣案->新式樣不印
'   If pa(8) = "3" Then
'      i = i + 1
'      ReDim Preserve strTxt(i)
'      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'         "','新式樣不印','♀')"
'   End If
'
'   'Modified by Morgan 2012/10/15 +231寄存證明
'   strExc(0) = "select np15,np08 from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "'" & _
'      " and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np06 is null and np07 in (202,231) and instr(np15,'專利申請書')=0" & _
'      " order by np08"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      i = i + 1
'      ReDim Preserve strTxt(i)
'      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'         "','有補件期限','♀')"
'
'      i = i + 1
'      ReDim Preserve strTxt(i)
'      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'         "','補件期限','" & RsTemp.Fields("np08") & "')"
'      strDoc = ""
'      strList = ""
'      With RsTemp
'      Do While Not .EOF
'         strExc(1) = ""
'         If InStr("" & .Fields(0), "委任書") > 0 Then
'            strExc(1) = "委任書"
'
'         ElseIf InStr("" & .Fields(0), "申請權證明") > 0 Then
'            strExc(1) = "申請權證明"
'            If InStr(strList, strExc(1)) = 0 Then
'               strExc(0) = "select * from pridate where pd01='" & pa(1) & "' and pd02='" & pa(2) & "'" & _
'                  " and pd03='" & pa(3) & "' and pd04='" & pa(4) & "' and pd07='101'"
'               intI = 1
'               Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  strExc(1) = "申請權證明-美優"
'               End If
'            End If
'
'         ElseIf InStr("" & .Fields(0), "優先權證明") > 0 Then
'            strExc(1) = "優先權證明"
'            If InStr(strList, strExc(1)) = 0 Then
'               strExc(0) = "select * from pridate where pd01='" & pa(1) & "' and pd02='" & pa(2) & "'" & _
'                  " and pd03='" & pa(3) & "' and pd04='" & pa(4) & "' and pd07='101'"
'               intI = 1
'               Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  i = i + 1
'                  ReDim Preserve strTxt(i)
'                  strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                     "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'                     "','優先權證明-美優','♀')"
'               End If
'            End If
'         ElseIf InStr("" & .Fields(0), "切結書") > 0 Then
'            strExc(1) = "切結書"
'
'         ElseIf InStr("" & .Fields(0), "僱傭契約") > 0 Then
'            strExc(1) = "僱傭契約"
'
'         ElseIf InStr("" & .Fields(0), "美國讓與") > 0 Then
'            strExc(1) = "美國讓與"
'            If InStr(strList, strExc(1)) = 0 Then
'               If InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X28186") > 0 Then
'                  strExc(1) = "美國讓與免公證"
'               End If
'            End If
'         ElseIf InStr("" & .Fields(0), "國內寄存證明") > 0 Then
'            strExc(1) = "國內寄存證明"
'
'         ElseIf InStr("" & .Fields(0), "國外寄存證明") > 0 Then
'            strExc(1) = "國外寄存證明"
'
'         ElseIf InStr("" & .Fields(0), "死亡證明") > 0 Then
'            strExc(1) = "死亡證明"
'
'         ElseIf InStr("" & .Fields(0), "繼承證明") > 0 Then
'            strExc(1) = "繼承證明"
'
'         ElseIf InStr("" & .Fields(0), "法人地位證明") > 0 Then
'            strExc(1) = "法人地位證明"
'
'         ElseIf InStr("" & .Fields(0), "國籍證明") > 0 Then
'            strExc(1) = "國籍證明"
'
'         Else
'            strDoc = strDoc & vbCrLf & .Fields(0)
'         End If
'
'         If strExc(1) <> "" And InStr(strList, strExc(1)) = 0 Then
'            i = i + 1
'            ReDim Preserve strTxt(i)
'            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'               "','" & strExc(1) & "','♀')"
'
'            strList = strList & strExc(1) & ";"
'         End If
'         .MoveNext
'      Loop
'      End With
'      If strDoc <> "" Then
'         i = i + 1
'         ReDim Preserve strTxt(i)
'         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'            "','其他補文件','" & ChgSQL(strDoc) & "')"
'      End If
'   End If
'
'   If Not ClsLawExecSQL(i, strTxt) Then
'      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
'   End If
'End Sub

'Added by Morgan 2012/11/29 複製 StartLetter 來修改
Private Sub StartLetter2(ByVal ET01 As String, ET02 As String, ByVal ET03 As String)
   Dim strTxt() As String, i As Integer
   Dim strDoc As String, strList As String
   'Dim strDateList As String 'Removed by Morgan 2022/4/21 沒用了
   
   EndLetter ET01, ET02, ET03, strUserNum

   i = 0
   If Text7 <> "" Then
      If PUB_GetEMailFlag(pa(1) & pa(2) & pa(3) & pa(4)) Then
         'Added by Morgan 2012/3/15 其他關係企業編號有需要另行告知--陳怡君
         strExc(1) = ChangeCustomerL(pa(75))
         If InStr("Y48309000,Y48309010,Y48309020,Y48309030,Y48309040,Y48309050,Y51326000", strExc(1)) > 0 Then
            strExc(1) = "有申請號-E-1"
         Else
         'End 2012/3/15
            strExc(1) = "有申請號-E"
         End If
      Else
         strExc(1) = "有申請號-非E"
      End If
   Else
      strExc(1) = "無申請號"
   End If
   i = i + 1
   ReDim Preserve strTxt(i)
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','" & strExc(1) & "','♀')"
      
   
   If pa(8) = "1" Then
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','發明案','♀')"
         
      strExc(0) = "select cp27 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "'" & _
         " and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='416' and cp57 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         i = i + 1
         ReDim Preserve strTxt(i)
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','實審有收文','♀')"
            
         If RsTemp(0) > 0 Then
            i = i + 1
            ReDim Preserve strTxt(i)
            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
               "','實審已發文','♀')"
         End If
      Else
         'Modify By Sindy 2021/4/27 ,np23
         strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "'" & _
            " and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np06||np07='416'"
         'Modify By Sindy 2021/7/21
         If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
            strExc(0) = strExc(0) & " order by np23"
         Else
         '2021/7/21 END
            strExc(0) = strExc(0) & " order by np08"
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            i = i + 1
            ReDim Preserve strTxt(i)
            'Modify By Sindy 2021/4/27
            If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','實審約定期限','" & RsTemp(2) & "')"
            Else
            '2021/4/27 END
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','實審本所期限','" & RsTemp(0) & "')"
            End If
            
            i = i + 1
            ReDim Preserve strTxt(i)
            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
               "','實審法定期限','" & RsTemp(1) & "')"
         End If
      End If
   End If
   
   'Modified by Morgan 2018/8/23 改例外欄位名:新式樣案->新式樣不印
   'Removed by Morgan 2023/9/13 改共用例外欄位<設計不印>
   'If pa(8) = "3" Then
   '   i = i + 1
   '   ReDim Preserve strTxt(i)
   '   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
   '      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
   '      "','新式樣不印','♀')"
   'End If
   'end 2023/9/13
   
   'Modified by Morgan 2012/10/15 +231寄存證明
   'Modified by Morgan 2018/10/24 +"基本資料表"也排除 --敏莉
   'Modify By Sindy 2021/7/21 + ,np23 Mark: and instr(np15,'專利申請書')=0 and instr(np15,'基本資料表')=0
   'Modified by Morgan 2022/4/21 +np09
   strExc(0) = "select np15,np08,np23,np09 from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "'" & _
      " and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np06 is null and np07 in (202,231)"
   'Modify By Sindy 2021/7/21
   If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
      strExc(0) = strExc(0) & " order by np23"
   Else
   '2021/7/21 END
      strExc(0) = strExc(0) & " order by np08"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','有補件期限','♀')"
      
      strDoc = ""
      strList = ""
      With RsTemp
      Do While Not .EOF
         strExc(1) = ""
         If InStr("" & .Fields(0), "委任書") > 0 Or InStr("" & .Fields(0), "委任狀") > 0 Then
            strExc(1) = "委任書"
         
         ElseIf InStr("" & .Fields(0), "法人地位證明") > 0 Then
            strExc(1) = "法人地位證明"
            
         ElseIf InStr("" & .Fields(0), "國籍證明") > 0 Then
            strExc(1) = "國籍證明"
            
         ElseIf InStr("" & .Fields(0), "國內寄存證明") > 0 Then
            strExc(1) = "國內寄存證明"
            
         ElseIf InStr("" & .Fields(0), "國外寄存證明") > 0 Then
            strExc(1) = "國外寄存證明"
            
         ElseIf InStr("" & .Fields(0), "優先權證明") > 0 Then
            strExc(1) = "優先權證明"
            
         'Add By Sindy 2021/7/21
         ElseIf InStr("" & .Fields(0), "發明人中譯名") > 0 Then
            strExc(1) = "發明人中譯名"
         ElseIf InStr("" & .Fields(0), "發明人國籍") > 0 Then
            strExc(1) = "發明人國籍"
         ElseIf InStr("" & .Fields(0), "申請人中譯名") > 0 Then
            strExc(1) = "申請人中譯名"
         ElseIf InStr("" & .Fields(0), "申請人國籍") > 0 Then
            strExc(1) = "申請人國籍"
         ElseIf InStr("" & .Fields(0), "代表人") > 0 Then
            strExc(1) = "代表人"
         ElseIf InStr("" & .Fields(0), "英文摘要") > 0 Then
            strExc(1) = "英文摘要"
         ElseIf InStr("" & .Fields(0), "英文參考本") > 0 Then
            strExc(1) = "英說"
         ElseIf InStr("" & .Fields(0), "客戶提供中說") > 0 Then
            strExc(1) = "中說"
            'Added by Morgan 2022/9/26
            If pa(175) = "Y" Then
               i = i + 1
               ReDim Preserve strTxt(i)
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','有序列表要印','♀')"
            End If
            'end 2022/9/26
         '2021/7/21 END
         'Added by Morgan 2021/10/12
         ElseIf InStr("" & .Fields(0), "台籍發明人ID碼") > 0 Then
            strExc(1) = "台籍發明人ID碼"
         'end 2021/10/12
         'Added by Morgan 2022/1/13
         ElseIf InStr("" & .Fields(0), "請求資訊不公開之聲明書") > 0 Then
            strExc(1) = "請求資訊不公開之聲明書"
         'end 2022/1/13
         'Added by Morgan 2022/1/17
         ElseIf InStr("" & .Fields(0), "全部發明人名") > 0 Then
            strExc(1) = "全部發明人名"
         ElseIf InStr("" & .Fields(0), "發明人名有特殊字") > 0 Then
            strExc(1) = "發明人名有特殊字"
         'end 2022/1/17
         Else
            'Modify By Sindy 2021/7/21
            If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
               'Modified by Morgan 2022/4/21 +補件法限
               'strDoc = strDoc & vbCrLf & .Fields(0) & ", 補件期限:" & .Fields("np23")
               strDoc = strDoc & vbCrLf & .Fields(0) & IIf(IsNull(.Fields("np09")), "", ", 法定期限:" & .Fields("np09")) & ", 補件期限:" & .Fields("np23")
            Else
            '2021/7/21 END
               strDoc = strDoc & vbCrLf & .Fields(0) & ", 補件期限:" & .Fields("np08")
            End If
         End If
         
         If strExc(1) <> "" And InStr(strList, strExc(1)) = 0 Then
            strList = strList & strExc(1) & ";"
            
            i = i + 1
            ReDim Preserve strTxt(i)
            'Modify By Sindy 2021/7/21
            If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','" & strExc(1) & "補件期限','" & .Fields("np23") & "')"
               
               'Added by Morgan 2022/4/21 沒有法限(Ex:客戶提供中說)要放空白才會帶底線，不可為空字串否則會顯示控制碼
               i = i + 1
               ReDim Preserve strTxt(i)
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','" & strExc(1) & "法定期限','" & Left(.Fields("np09") & String(8, " "), 8) & "')"
               'end 2022/4/21
               
            Else
            '2021/7/21 END
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','" & strExc(1) & "補件期限','" & .Fields("np08") & "')"
            End If
         End If
         
         'Added by Morgan 2012/12/28
         'Removed by Morgan 2022/4/21 沒用了
         'If strDateList = "" Then
         '   'Modify By Sindy 2021/7/21
         '   If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
         '      strDateList = .Fields("np23")
         '   Else
         '   '2021/7/21 END
         '      strDateList = .Fields("np08")
         '   End If
         'ElseIf InStr(strDateList, .Fields("np08")) = 0 Then
         '   'Modify By Sindy 2021/7/21
         '   If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
         '      strDateList = strDateList & "," & .Fields("np23")
         '   Else
         '   '2021/7/21 END
         '      strDateList = strDateList & "," & .Fields("np08")
         '   End If
         'End If
         'end 2022/4/21
         
         .MoveNext
      Loop
      End With
      If strDoc <> "" Then
         i = i + 1
         ReDim Preserve strTxt(i)
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','其他補文件','" & ChgSQL(strDoc) & "')"
      End If
      
      
      'Added by Morgan 2012/12/28
      'Removed by Morgan 2022/4/21 沒用了
      'If strDateList <> "" Then
      '   i = i + 1
      '   ReDim Preserve strTxt(i)
      '   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      '      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      '      "','補件期限','" & strDateList & "')"
      'End If
      'end 2022/4/21
      'end 2012/12/28
      
   End If

   If Not ClsLawExecSQL(i, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

'Added by Lydia 2019/06/25 跳回前一作業
Private Function ExceptReturn() As Boolean
     ExceptReturn = False
     If mOtherForm = "" Then Exit Function
     
     If InStr(mOtherForm, "frm060104_1") > 0 Then
         frm060104_1.Show
         If Left(mOtherForm, 1) = "0" Then '0-確定; 3-同時發文
             frm060104_1.Clear
         Else
             frm060104_1.ReQuery
         End If
     End If
     ExceptReturn = True
End Function

'Add By Sindy 2022/5/17
Private Sub txtRecDate_GotFocus()
   TextInverse txtRecDate
End Sub
Private Sub txtRecDate_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub
Private Sub txtEmail_GotFocus()
   TextInverse txtEmail
End Sub
Private Sub txtEmail_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      Beep
      KeyAscii = 0
   End If
End Sub
Private Sub txtRecDate_Validate(Cancel As Boolean)
   If txtRecDate.Tag <> txtRecDate.Text Then
      If txtRecDate = "Y" Then
         txtEmail = "Y"
      End If
   End If
   txtRecDate.Tag = txtRecDate.Text
End Sub
'2022/5/17 END
