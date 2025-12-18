VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090908_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "主管核判"
   ClientHeight    =   3432
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   8916
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3432
   ScaleWidth      =   8916
   Begin VB.CommandButton cmdOpen 
      Caption         =   "卷宗區(&C)"
      Height          =   300
      Index           =   1
      Left            =   1440
      TabIndex        =   32
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "原始檔(&P)"
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   30
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   360
      Index           =   1
      Left            =   7485
      TabIndex        =   27
      Top             =   120
      Width           =   1160
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&S)"
      Default         =   -1  'True
      Height          =   360
      Index           =   0
      Left            =   6480
      TabIndex        =   18
      Top             =   120
      Width           =   900
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   2970
      Width           =   2295
   End
   Begin VB.TextBox txtData 
      Height          =   270
      Index           =   2
      Left            =   9000
      TabIndex        =   17
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtData 
      Height          =   270
      Index           =   1
      Left            =   5880
      MaxLength       =   4
      TabIndex        =   1
      Top             =   620
      Width           =   615
   End
   Begin VB.TextBox txtData 
      Height          =   270
      Index           =   0
      Left            =   5040
      MaxLength       =   7
      TabIndex        =   0
      Top             =   620
      Width           =   855
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "急件"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   620
      Width           =   735
   End
   Begin VB.Label lblDTime 
      Caption         =   "Label2"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7110
      TabIndex        =   39
      Top             =   2670
      Width           =   1665
   End
   Begin VB.Label lblData 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   11
      Left            =   7320
      TabIndex        =   38
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSForms.Label lblCM 
      Height          =   255
      Left            =   6150
      TabIndex        =   37
      Top             =   2220
      Width           =   2535
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Caption         =   "命名人員："
      Size            =   "4471;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCMboth 
      Caption         =   "相關案號：(台灣大陸案件提示)"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3000
      TabIndex        =   36
      Top             =   2220
      Width           =   2550
   End
   Begin VB.Label Label1 
      Caption         =   "認領期限："
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   6120
      TabIndex        =   35
      Top             =   2670
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "目前各組認領人員："
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   2670
      Width           =   1815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "核判"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   1320
      TabIndex        =   33
      Top             =   3030
      Width           =   420
   End
   Begin VB.Label Label5 
      Caption         =   "命名人員："
      Height          =   255
      Index           =   1
      Left            =   6120
      TabIndex        =   31
      Top             =   945
      Width           =   975
   End
   Begin VB.Label lblData 
      BackColor       =   &H80000003&
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
      Index           =   16
      Left            =   3960
      TabIndex        =   29
      Top             =   945
      Width           =   975
   End
   Begin VB.Label lblData 
      BackColor       =   &H80000003&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   15
      Left            =   1200
      TabIndex        =   28
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label lblData 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   26
      Top             =   945
      Width           =   855
   End
   Begin VB.Label lblData 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   25
      Top             =   640
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "本所期限："
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   24
      Top             =   945
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "法定期限："
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   23
      Top             =   640
      Width           =   975
   End
   Begin VB.Label lblData 
      BackColor       =   &H80000003&
      Height          =   435
      Index           =   3
      Left            =   1320
      TabIndex        =   22
      Top             =   1920
      Width           =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "2.中說類型："
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   21
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblData 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   20
      Top             =   1620
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "1.專利種類："
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   1620
      Width           =   1095
   End
   Begin VB.Label lblData 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   9
      Left            =   7080
      TabIndex        =   16
      Top             =   1260
      Width           =   855
   End
   Begin VB.Label lblData 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   8
      Left            =   7080
      TabIndex        =   15
      Top             =   960
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   120
      X2              =   8760
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lblData 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   10
      Left            =   7320
      TabIndex        =   14
      Top             =   1620
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "國　　籍："
      Height          =   255
      Index           =   5
      Left            =   6120
      TabIndex        =   13
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   3
      Left            =   6120
      TabIndex        =   12
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label lblData 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   6
      Left            =   3960
      TabIndex        =   11
      Top             =   1920
      Width           =   1485
   End
   Begin VB.Label Label4 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblData 
      BackColor       =   &H80000003&
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   9
      Top             =   1620
      Width           =   1005
   End
   Begin VB.Label Label6 
      Caption         =   "收文日期："
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   8
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label lblData 
      BackColor       =   &H80000003&
      Height          =   350
      Index           =   4
      Left            =   3960
      TabIndex        =   7
      Top             =   1260
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "總收文號："
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   6
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "譯畢期限：                ，請於　　　　　　　　　前譯畢名稱"
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   4
      Top             =   645
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1260
      Width           =   975
   End
End
Attribute VB_Name = "frm090908_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Lydia 2022/03/01 外專新案認領區-主管核判
Option Explicit
Dim m_PrevForm As Form '前一畫面
Dim m_UserNo As String   '傳入員工編號
Dim strCase(1 To 12) As String '1~4本所案號pa01~pa04,5-專利種類pa08,6-申請國家pa09,7-分案組別pa150,8-設計案屬性pa158,9-申請日,10-公告日(PA14),11-目前准/駁(PA16),12-名稱有特殊字(PA174)
Dim m_TCT01 As String  '收文號=PK
Dim m_TCT04 As String  '工程師主管
Dim m_TCT07 As String  '工程師主任
Dim m_TCT10 As String  '命名人員編號
Dim m_TCT27kind As String '欲翻譯此案件者可輸入的選項
Dim m_Receiver As String, m_ReceGrp As String '通知認領人員+組別
Dim n_CP118 As String '新申請案：是否電子送件
Dim m_TCN13 As String 'Added by Lydia 2023/06/14
Dim m_TCN23 As String 'Added by Lydia 2024/01/29

Private Function CheckDiff() As Boolean
    If Trim(Mid(Combo1.Text, 3)) <> "" Then
       CheckDiff = True
    End If
End Function

Private Function SaveDatabase() As Boolean
Dim m_Grp As String, m_GrpMan As String
Dim bolConn As Boolean

On Error GoTo Err01
    
    If Trim(Combo1.Text) <> "" Then
       For intI = 1 To 4
          If InStr(Combo1.Text, PUB_GetFCPGrpName("" & intI)) > 0 Then
              m_Grp = intI
              m_GrpMan = Pub_GetFCPGrpMan("" & intI)
              m_GrpMan = PUB_GetStateForMan(m_GrpMan) '特殊情況之指定職代
              Exit For
          End If
       Next intI
    End If
    If m_Grp <> "" Then
       bolConn = True
       cnnConnection.BeginTrans
         m_Receiver = Replace(Mid(m_Receiver, 2), m_GrpMan, "")
         If Mid(m_Receiver, 1) = ";" Then m_Receiver = Mid(m_Receiver, 2)
         m_Receiver = Replace(m_Receiver, ";;", ";")
         '更新認領人員TFA06~08
         'Modified by Lydia 2024/01/29
         'strSql = "Update TransFeeAssign Set TFA06='" & m_UserNo & "', TFA07=to_char(sysdate,'yyyymmdd'), TFA08=to_char(sysdate,'hh24miss') " & _
                     "Where (tfa01,tfa04,tfa09)=(select tfa01,tfa04,tfa09 from transfeeassign,staff where tfa01='" & m_TCT01 & "' and tfa04=st01(+) and tfa09='1' and st16='" & m_Grp & "')"
         'cnnConnection.Execute strSql
         strSql = "Update TransFeeAssign Set TFA06='" & m_UserNo & "', TFA07=to_char(sysdate,'yyyymmdd'), TFA08=to_char(sysdate,'hh24miss') " & _
                     "Where (tfa01,tfa04,tfa09)=(select tfa01,tfa04,tfa09 from transfeeassign,staff where tfa01='" & m_TCT01 & "' and tfa04=st01(+) and tfa09='" & m_TCN23 & "' and st16='" & m_Grp & "')"
         cnnConnection.Execute strSql, intI
         If intI = 0 Then   '若主管未認領,補記錄; Ex.FCP-071056
            strSql = "Insert Into TransFeeAssign(TFA01,TFA02,TFA03,TFA04,TFA05,TFA06,TFA07,TFA08,TFA09) VALUES ('" & m_TCT01 & "'," & _
                  "to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'), '" & m_GrpMan & "','Y','" & m_UserNo & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & m_TCN23 & "') "
            cnnConnection.Execute strSql, intI
         End If
         'end 2024/01/29
         strSql = "Update TrackingCaseName Set TCN20='" & m_Grp & "' Where TCN05='" & m_TCT01 & "' "
         cnnConnection.Execute strSql
         strExc(1) = PUB_GetTCNmTitle(strCase(1), strCase(2), strCase(3), strCase(4), strCase(9), m_TCN13, "")
         strExc(1) = strExc(1) & "，核判結果：此案為" & Trim(Mid(Combo1, 3)) & "新案，請繼續分案予工程師進行命名，謝謝！"
         strExc(2) = "同主旨"
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                      " values( '" & strUserNum & "','" & m_GrpMan & "',to_char(sysdate,'yyyymmdd')" & _
                      ",to_char(sysdate,'hh24miss'),'" & strExc(1) & "','" & strExc(2) & "','" & m_Receiver & "' )"
         cnnConnection.Execute strSql
                   
         If PUB_UpdateTCNstate("2", strCase(1) & strCase(2) & strCase(3) & strCase(4)) = False Then
             GoTo Err01
         End If
       cnnConnection.CommitTrans
    End If

    SaveDatabase = True
    Exit Function
    
Err01:
If Err.Number <> 0 Then
   If bolConn = True Then
       cnnConnection.RollbackTrans
   End If
   MsgBox Err.Description
End If
End Function

Private Sub cmdOK_Click(Index As Integer)

  Select Case Index
      Case 0 '存檔
         If SaveDatabase = True Then
            GoTo JumpCloseFrm
         Else
            Exit Sub
         End If
      Case 1 '回前畫面
         If cmdOK(0).Enabled = False Then GoTo JumpCloseFrm
         If CheckDiff = True Then
            If MsgBox("你並未存檔，確定離開嗎?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Sub
            Else
                GoTo JumpCloseFrm
            End If
         Else
            GoTo JumpCloseFrm
         End If
  End Select
  Exit Sub
  
JumpCloseFrm:
  Me.Hide
  Unload Me
End Sub

Private Sub cmdOpen_Click(Index As Integer)
   If Index = 0 Then
      If PUB_CheckFormExist("frm100101_M") Then
          MsgBox "請先關閉共同查詢〔原始檔區〕畫面！"
          Exit Sub
      Else
          If cmdOpen(0).Tag = "" Then
              MsgBox strCase(1) & "-" & strCase(2) & "在〔原始檔區〕的English_Vers收文號不存在!", vbInformation
          Else
              Screen.MousePointer = vbHourglass
              frm100101_M.m_strKey = cmdOpen(0).Tag '總收文號
              frm100101_M.SetParent Me
              If frm100101_M.QueryData = True Then
                 frm100101_M.Show
                 Me.Hide
              End If
              Screen.MousePointer = vbDefault
          End If
      End If
   ElseIf Index = 1 Then
      If PUB_CheckFormExist("frm100101_L") Then
          MsgBox "請先關閉共同查詢〔卷宗區〕畫面！"
          Exit Sub
      Else
         Screen.MousePointer = vbHourglass
         frm100101_L.m_strKey = strCase(1) & "-" & strCase(2) & "-" & strCase(3) & "-" & strCase(4)
         frm100101_L.SetParent Me
         If frm100101_L.QueryData = True Then
            frm100101_L.Show
            Me.Hide
         End If
         Screen.MousePointer = vbDefault
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache
   If TypeName(m_PrevForm) = "frm090908" Then
      m_PrevForm.doQuery False
      m_PrevForm.Show
   End If
   Set frm090908_1 = Nothing
End Sub

Public Sub SetParent(ByRef fm As Form, ByVal pCase As String, ByVal pNo As String, ByVal pUser As String)
   Set m_PrevForm = fm
   m_TCT01 = pNo
   m_UserNo = pUser
   Call ChgCaseNo(Replace(pCase, "-", ""), strCase)
End Sub

Private Sub ClearForm(Optional ByVal bolRest As Boolean)
Dim oLbl As LABEL
Dim oTxt As TextBox
   For Each oLbl In lblData
      oLbl.Caption = ""
      If bolRest = True Then oLbl.BackColor = &H8000000F
   Next
   
   Chk1.Value = 0
   Chk1.Tag = ""
   Combo1.Locked = False
   m_Receiver = ""
   m_ReceGrp = ""
   
   For Each oTxt In txtData
      oTxt.Text = ""
      oTxt.Tag = ""
   Next
   
   lblCMboth.Caption = ""
   lblCMboth.Tag = ""
   lblCM.Tag = ""
   lblDTime.Caption = ""
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   Combo1.Clear
   ClearForm True
  
   If ReadData = True Then
      Call SetCombo1
   End If

End Sub

Private Function ReadData() As Boolean
Dim rsRd As New ADODB.Recordset
    
    '改成模組控制,若基本資料顯示有變,要注意frm090902_1,frm090902_2,frm090903_1的欄位
    If PUB_GetTCTread(Me, strCase, m_TCT27kind, n_CP118) = True Then
       ReadData = True
       '顯示相關案
       If lblCMboth.Tag <> "" Then
           Call ChgCaseNo(lblCMboth.Tag, strExc)
           strExc(0) = "select tct10,st02 from caseprogress,transcasetitle,staff where cp01='" & strExc(1) & "' and cp02='" & strExc(2) & "' and cp03='" & strExc(3) & "' and cp04='" & strExc(4) & "' and cp31='Y' and cp09=tct01(+) and tct10=st01(+) "
           intI = 1
           Set rsRd = ClsLawReadRstMsg(intI, strExc(0))
           If intI = 1 Then
                lblCM.Visible = True
                lblCM.Caption = "命名人員：" & rsRd.Fields("tct10") & " " & rsRd.Fields("st02")
                lblCM.Tag = "" & rsRd.Fields("tct10")
           End If
       Else
           lblCM.Visible = False
       End If
       Set rsRd = Nothing
       Call SetCaseTitle(True)
       
       If PUB_ChkCPExist(strCase, cntEnglish_Vers, , strExc(1), , "D") = True Then 'English_Vers992
           cmdOpen(0).Tag = strExc(1)
       End If
    Else
       MsgBox "查無資料 !", vbExclamation
       Unload Me
    End If
    
End Function

'設案件命名欄位
Private Sub SetCaseTitle(ByVal bolCmb As Boolean)
Dim rsA As New ADODB.Recordset
Dim Str01 As String
Dim intA As Integer

    Str01 = "select A.*,B.ST02 " & _
            "FROM TransCaseTitle A,STAFF B WHERE TCT01='" & m_TCT01 & "' AND TCT10=ST01(+) "
    intA = 1
    Set rsA = ClsLawReadRstMsg(intA, Str01)
    If intA = 1 Then
        With rsA
           '急件
           If Val("" & .Fields("TCT02")) > 0 Then
              Chk1.Value = 1
           End If
           '譯畢期限
           If "" & .Fields("TCT02") <> "" Then
              txtData(0).Text = TransDate(.Fields("TCT02"), 1)
           End If
           txtData(0).Tag = txtData(0).Text
           If "" & .Fields("TCT03") <> "" Then
              txtData(1).Text = Format(.Fields("TCT03"), "0000")
           End If
           txtData(1).Tag = txtData(1).Text
           '工程師主管
           m_TCT04 = "" & .Fields("TCT04")
           '工程師主任
           m_TCT07 = "" & .Fields("TCT07")
           '命名人員
           m_TCT10 = Trim("" & .Fields("TCT10"))
           '譯畢期限
           txtData(0).Locked = True: txtData(1).Locked = True
           Chk1.Enabled = False
        End With
    End If
    'Modified by Lydia 2023/06/14 +TCN13
    'Modified by Lydia 2024/01/29 +TCN23
    Str01 = "SELECT DECODE(TCN21,NULL,'',SUBSTR(SQLDATET(TCN21),1,9)||' '||substr(SQLTIME6(TCN22||'00'),1,5)) AS 認領期限,TCN13,TCN23 FROM TrackingCaseName " & _
               "Where TCN05='" & m_TCT01 & "' "
    intA = 1
    Set rsA = ClsLawReadRstMsg(intA, Str01)
    If intA = 1 Then
       lblDTime.Caption = "" & rsA.Fields("認領期限")
       m_TCN13 = "" & rsA.Fields("TCN13") 'Added by Lydia 2023/06/14
       m_TCN23 = "" & rsA.Fields("TCN23") 'Added by Lydia 2024/01/29
    End If
    
    Set rsA = Nothing
End Sub

Private Sub SetCombo1()
Dim rsB As New ADODB.Recordset
Dim intA As Integer, intB As Integer
     
   For intA = 1 To 4
      If intA <> 3 Then '跳過日文組
          strSql = "select tfa05,tfa04,st16,tfa09 from transfeeassign,staff where tfa01='" & m_TCT01 & "' and tfa04=st01(+) and st16='" & intA & "' and tfa09=(" & _
                      "select max(tfa09) state from transfeeassign,staff where tfa01='" & m_TCT01 & "' and tfa04=st01(+) and st16='" & intA & "') "
          intB = 1
          Set rsB = ClsLawReadRstMsg(intB, strSql)
          If intB = 1 Then
               If "" & rsB.Fields("TFA05") = "Y" Then
                   Combo1.AddItem rsB.Fields("TFA05") & " " & PUB_GetFCPGrpName("" & intA)
                   m_ReceGrp = m_ReceGrp & ";" & rsB.Fields("st16")
                   strExc(1) = Pub_GetFCPGrpMan("" & intA)
                   strExc(1) = PUB_GetStateForMan(strExc(1)) '特殊情況之指定職代
                   If strExc(1) <> "" & rsB.Fields("tfa04") Then  '職代
                       m_Receiver = m_Receiver & ";" & strExc(1)
                   End If
                   m_Receiver = m_Receiver & ";" & rsB.Fields("tfa04")
               Else
                   Combo1.AddItem "   " & PUB_GetFCPGrpName("" & intA)
               End If
          Else
               Combo1.AddItem "   " & PUB_GetFCPGrpName("" & intA)
          End If
      End If
   Next intA
   Set rsB = Nothing
   
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
    TextInverse txtData(Index)
End Sub
