VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_g 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "¥~±Mµo¤å-FMP®×"
   ClientHeight    =   4170
   ClientLeft      =   270
   ClientTop       =   960
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8760
   Begin VB.TextBox txtSC01 
      Height          =   270
      Left            =   5760
      MaxLength       =   7
      TabIndex        =   2
      Top             =   2460
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "®×¥ó¶i«×(&C)"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   2475
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox txtCP43 
      Height          =   270
      Left            =   1440
      MaxLength       =   9
      TabIndex        =   0
      Top             =   2460
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   17
      Top             =   450
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   16
      Top             =   450
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2460
      MaxLength       =   1
      TabIndex        =   15
      Top             =   450
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   14
      Top             =   450
      Width           =   375
   End
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   3285
      MaxLength       =   4
      TabIndex        =   4
      Top             =   2910
      Width           =   600
   End
   Begin VB.TextBox txtCP14 
      Height          =   270
      Left            =   4995
      MaxLength       =   9
      TabIndex        =   5
      Top             =   2910
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "¦^«eµe­±(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7404
      TabIndex        =   8
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6576
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox txtCP27 
      Height          =   270
      Left            =   1035
      MaxLength       =   7
      TabIndex        =   3
      Top             =   2910
      Width           =   1095
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1140
      TabIndex        =   9
      Top             =   750
      Width           =   7395
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "13044;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCP64 
      Height          =   765
      Left            =   1035
      TabIndex        =   6
      Top             =   3240
      Width           =   7410
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13070;1349"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "°lÂÜ·|½Zµ²ªG´Á­­:"
      Height          =   180
      Index           =   15
      Left            =   4125
      TabIndex        =   38
      Top             =   2505
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8640
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      X1              =   120
      X2              =   8640
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Label Label1 
      Caption         =   "¬ÛÃöÁ`¦¬¤å¸¹:"
      Height          =   180
      Index           =   14
      Left            =   180
      TabIndex        =   37
      Top             =   2505
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   6
      Left            =   6135
      TabIndex        =   36
      Top             =   2940
      Width           =   2310
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4075;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   3
      Left            =   5040
      TabIndex        =   35
      Top             =   1770
      Width           =   3510
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6191;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   4
      Left            =   1140
      TabIndex        =   34
      Top             =   2100
      Width           =   2610
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4604;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   1140
      TabIndex        =   33
      Top             =   1110
      Width           =   2610
      VariousPropertyBits=   27
      Size            =   "4604;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   5
      Left            =   5040
      TabIndex        =   32
      Top             =   2100
      Width           =   2610
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4604;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "ªk©w´Á­­:"
      Height          =   180
      Index           =   9
      Left            =   4125
      TabIndex        =   31
      Top             =   2100
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "¥»©Ò´Á­­:"
      Height          =   180
      Index           =   8
      Left            =   180
      TabIndex        =   30
      Top             =   2100
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¾÷Ãö¤å¸¹:"
      Height          =   180
      Index           =   7
      Left            =   4125
      TabIndex        =   29
      Top             =   1770
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤é:"
      Height          =   180
      Index           =   3
      Left            =   4140
      TabIndex        =   28
      Top             =   1110
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "±M§Q¦WºÙ:"
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   27
      Top             =   750
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð®×¸¹:"
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   26
      Top             =   1110
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¥»©Ò®×¸¹:"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   25
      Top             =   450
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "®×¥ó©Ê½è:"
      Height          =   180
      Index           =   6
      Left            =   180
      TabIndex        =   24
      Top             =   1770
      Width           =   765
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   2
      Left            =   1140
      TabIndex        =   23
      Top             =   1770
      Width           =   2610
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4604;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   22
      Top             =   1110
      Width           =   2610
      VariousPropertyBits=   27
      Size            =   "4604;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¦¬¤å¸¹:"
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   21
      Top             =   1440
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¦¬¤å¤é:"
      Height          =   180
      Index           =   5
      Left            =   4140
      TabIndex        =   20
      Top             =   1440
      Width           =   585
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   0
      Left            =   1140
      TabIndex        =   19
      Top             =   1440
      Width           =   2610
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4604;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   18
      Top             =   1440
      Width           =   2610
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4604;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¤u§@®É¼Æ:"
      Height          =   180
      Index           =   12
      Left            =   2430
      TabIndex        =   13
      Top             =   2955
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "©Ó¿ì¤H:"
      Height          =   180
      Index           =   11
      Left            =   4275
      TabIndex        =   12
      Top             =   2955
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¶i«×³Æµù:"
      Height          =   180
      Index           =   13
      Left            =   180
      TabIndex        =   11
      Top             =   3270
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "µo¤å¤é:"
      Height          =   180
      Index           =   10
      Left            =   180
      TabIndex        =   10
      Top             =   2955
      Width           =   585
   End
End
Attribute VB_Name = "frm060104_g"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/18 Form2.0¤w­×§ï
'Memo By Morgan 2012/12/10 ´¼Åv¤H­ûÄæ¤w­×§ï
'Created by Morgan 2012/5/16
Option Explicit

Dim pa() As String
Dim intWhere As Integer
Dim m_CP10 As String, m_CP09 As String
Dim m_CP60 As String 'Added by Lydia 2015/02/26
Dim m_CP142 As String 'Add By Sindy 2015/12/17
Dim m_CP164 As String 'Add By Sindy 2021/4/20
Dim bolAddSC As Boolean 'Added by Lydia 2016/01/21 ¬O§_²£¥Í¦æ¨Æ¾äºÞ¨î
Dim mFA10 As String 'Added by Lydia 2019/07/03 ¥N²z¤H°êÄy


Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = °ê¥~_FC
   With frm060104_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      m_CP09 = .Tag
      Label3(0) = m_CP09
   End With
   ReDim pa(TF_PA)
   ReadPatent
   Combo1.ListIndex = -1 '0
   txtCP27 = strSrvDate(2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060104_g = Nothing
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         If TxtValidate = False Then Exit Sub
         
         'Add by Sindy 2021/11/18 ÀË¬dµe­±¤Wªºª«¥ó¬O§_§t¦³Unicode¤å¦r
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If

         If FormSave = False Then
            MsgBox "¦sÀÉ¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical
            Exit Sub
         Else
            'ÀË¬d¥N²z¤HEmail
            PUB_CheckEMail pa(75), pa(144)
            If pa(145) <> "" Then
               PUB_CheckEMail pa(75), pa(145)
            End If
            
            'Add By Sindy 2023/11/9
            If frm060104_1.bolIsEMPFlow = True Then
               frm090202_4.QueryData
            End If
            '2023/11/9 End
            '­Y¦³¥¼µo¤å¸ê®ÆÅã¥ÜÄµ§i
            'Modify By Sindy 2023/11/9
            If PUB_GetCPunIssueDatas("" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)) Then
               frm060104_1.Show
               frm060104_1.ReQuery
            Else
               'Add By Sindy 2023/11/9
               If frm060104_1.bolIsEMPFlow = True Then
                  Unload frm060104_1
               Else
               '2023/11/9 End
                  frm060104_1.Show
                  frm060104_1.Clear
               End If
            End If
         End If
      Case 1
         frm060104_1.Show
   End Select
   Unload Me
End Sub

Private Function TxtValidate() As Boolean
Dim bCancel As Boolean
   
   txtCP27_Validate bCancel
   If bCancel = True Then Exit Function
   txtCP14_Validate bCancel
   If bCancel = True Then Exit Function
   txtCP113_Validate bCancel
   If bCancel = True Then Exit Function
   
   'Add By Sindy 2015/12/17 ÀË¬d¬O§_¦³«ü©w°e¥ó¤é´Á,­Y¦³¤£¥i¤p©ó«ü©w¤é´Á°e¥ó
   If m_CP142 <> "" Then
      'Modify By Sindy 2021/12/3 ²QµØ»¡¤§«á¥i¥H§t·í¤Ñµo¤å
      'If m_CP142 >= strSrvDate(1) Then
      If m_CP142 > strSrvDate(1) Then
      '2021/12/3 END
         'Add By Sindy 2021/4/20
         'Modify By Sindy 2021/10/20 + 3.¤§«á
         If ((m_CP164 = "1" Or m_CP164 = "") And m_CP142 > strSrvDate(1)) Or _
            m_CP164 = "3" Then '1.·í¤Ñ 3.¤§«á
         '2021/4/20 END
            MsgBox "¦³«ü©w°e¥ó¤é´Á¡]" & ChangeWStringToTDateString(m_CP142) & "¡^¡A¤£¥i´£«e°e¥ó!!!"
            Exit Function
         End If
      End If
   End If
   '2015/12/17 END
   'Added by Lydia 2015/12/31 P®×·|½Z924µo¤å¡AÀË¬d¤@©w­n¦³¬ÛÃöÁ`¦¬¤å¸¹
   If m_CP10 = "924" And pa(1) = "P" And txtCP43.Visible = True Then
      If txtCP43 = "" Then
         MsgBox "®×¥ó©Ê½è¬°·|½Z®É¡A¤@©w­n¦³¬ÛÃöÁ`¦¬¤å¸¹!!", vbCritical
         txtCP43.SetFocus
         Exit Function
      End If
      'Modified by Lydia 2016/01/21
      'txtCP43_Validate bCancel
      CheckCP43 False, bCancel
      If bCancel = True Then
         'Added by Lydia 2016/01/21
         txtCP43.SetFocus
         txtCP43_GotFocus
         'end 2016/01/21
         Exit Function
      End If
      '·s¥Ó½Ð®×­n¿é¤J°lÂÜ·|½Zµ²ªG´Á­­
      If txtSC01.Visible = True Then
        txtSC01_Validate bCancel
        If bCancel = True Then
           Exit Function
        End If
      End If
   End If
   'end 2015/12/31
   
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
Dim stUpdate As String 'Added by Lydia 2015/12/31
On Error GoTo CheckingErr
   cnnConnection.BeginTrans

   'Modified by Lydia 2015/12/31
   'strSql = "Update Caseprogress set cp27=" & DBDATE(txtCP27) & ",cp14='" & txtCP14 & "',cp113=" & IIf(txtCP113 = "", "NULL", txtCP113) & ",cp64='" & ChgSQL(txtCP64) & "' where cp09='" & m_CP09 & "'"
   If m_CP10 = "924" And txtCP43.Visible = True Then
      stUpdate = ",cp43=" & CNULL(txtCP43)
   End If
   strSql = "Update Caseprogress set cp27=" & DBDATE(txtCP27) & ",cp14='" & txtCP14 & "',cp113=" & IIf(txtCP113 = "", "NULL", txtCP113) & ",cp64='" & ChgSQL(txtCP64) & "'" & stUpdate & " where cp09='" & m_CP09 & "'"
   cnnConnection.Execute strSql, intI
   
   'Added by Lydia 2015/02/26 ­Y¤w¶}½Ð´Ú³æ«h´«©Ó¿ì¤H©Î®Ö½Z¤H®ÉµoMail³qª¾ÀRªÚ
   If m_CP60 > "X" Then
      'Modified by Lydia 2019/10/17 ¥»©Ò®×¸¹+"-"
      'PUB_PointReAssignInform Text1 & Text2 & Text3 & Text4, m_CP60, txtCP14.Tag, txtCP14.Text
      PUB_PointReAssignInform pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)), m_CP60, txtCP14.Tag, txtCP14.Text
   End If
   
   'Added by Lydia 2015/12/31 P·|½Zµo¤å¦sÀÉ®É¡A¦Û°Ê·s¼W°ê¥~³¡¦æ¨Æ¾ä¸ê®Æ
   If m_CP10 = "924" And txtCP43 <> "" And txtCP43.Visible = True And txtSC01 <> "" And txtSC01.Visible = True Then
      strExc(1) = DBDATE(txtSC01)
      
      '´£¿ô¤H­û1¹w³]¬°¸Ó¥»©Ò®×¸¹¤§FCP©Ó¿ì·~°È­û¡A´£¿ô¤H­û2¹w³]¬°¸Ó®×¤§³Ì«á¤uµ{®v¡F¥i¸Ñ°£¤H­û¹w³]¬°´£¿ô¤H­û1¡F¨Æ¥Ñ='°lÂÜ·|½Zµ²ªG'¡F¥»©Ò®×¸¹¤]­n¦s¡F
      strExc(3) = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))
      strExc(5) = ""
      strExc(0) = "select cp14,st04,st02 from caseprogress,staff where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' " & _
                  "and st01(+)=cp14 and cp57 is null order by cp05 desc,cp09 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.Fields("st04") = "1" Then
            strExc(5) = RsTemp.Fields("cp14")
         End If
      End If
      If strExc(3) <> "" Then
          strExc(4) = "°lÂÜ·|½Zµ²ªG"
          strExc(2) = strExc(3) & IIf(strExc(5) <> "", ",", "") & strExc(5)
          '¥i¸Ñ°£¤H­û¹w³]¬°´£¿ô¤H­û1
          If PUB_AddFCPStaffCalendar(strExc(1), "1", strExc(2), strExc(4), strExc(3), "1", pa(1), pa(2), pa(3), pa(4)) Then
          End If
      End If
   End If
   'end 2015/12/31
   cnnConnection.CommitTrans
   
   FormSave = True
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function


Private Sub ReadPatent()
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   Select Case pa(1)
      Case "P", "CFP"
         If ClsPDReadPatentDatabase(pa(), intWhere) Then
            Label2(0) = pa(11)
            Label2(1) = pa(10)
            AddCboName Combo1, pa(5), pa(6), pa(7)
         End If
      Case "PS", "CPS"
         If PUB_ReadServicePracticeDatabase(pa(), intWhere) Then
            Label2(0) = pa(11)
            Label2(1) = pa(10)
            AddCboName Combo1, pa(5), pa(6), pa(7)
         End If
   End Select
   
   If pa(75) <> "" Then mFA10 = GetPrjNationNumber(ChangeCustomerL(pa(75))) 'Added by Lydia 2019/07/03
   
   'Added by Lydia 2015/02/26
   'Added by Lydia 2015/12/31 +CP43
   strExc(0) = "select cp05,cp10,cpm04,cp08,cp06,cp07,cp14,cp113,cp64,st02,CP60,CP142,CP43,CP164" & _
      " from caseprogress,casepropertymap,staff where cp09='" & Label3(0) & "'" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp14"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
        With RsTemp
           Label3(1) = .Fields("cp05") - 19110000
           Label3(2) = .Fields("cp10") & " " & .Fields("cpm04")
           m_CP10 = .Fields("cp10")
           m_CP142 = "" & .Fields("cp142") 'Add By Sindy 2015/12/17
           m_CP164 = "" & .Fields("CP164") 'Add By Sindy 2021/4/20
           Label3(3) = "" & .Fields("cp08")
           If Not IsNull(.Fields("cp06")) Then
              Label3(4) = .Fields("cp06") - 19110000
           Else
              Label3(4) = "" 'Added by Lydia 2016/01/21
           End If
           If Not IsNull(.Fields("cp07")) Then
              Label3(5) = .Fields("cp07") - 19110000
           Else
              Label3(5) = "" 'Added by Lydia 2016/01/21
           End If
           txtCP14 = "" & .Fields("cp14")
           'modify by sonia 2015/9/21
           'Label3(6) = "" & .Fields("st02")
           txtCP14_Validate False
           'end 2015/9/21
           'Added by Lydia 2015/02/26
           txtCP14.Tag = txtCP14.Text
           If Not IsNull(.Fields("CP60")) Then
              m_CP60 = .Fields("CP60")
           Else
              m_CP60 = ""
           End If
           'end 2015/02/26
           Label3(6) = "" & .Fields("st02")
           txtCP113 = "" & .Fields("cp113")
           txtCP64 = "" & .Fields("cp64")
           'Added by Lydia 2015/12/31
           txtCP43 = "" & .Fields("cp43")
           If m_CP10 = "924" Then
              Label1(14).Visible = True
              txtCP43.Visible = True
              Command2.Visible = True
           End If
           'end 2015/12/31
        End With
   End If
End Sub

Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

Private Sub txtCP113_Validate(Cancel As Boolean)
   If txtCP113 <> "" Then
      If Not IsNumeric(txtCP113) Then
         MsgBox "½Ð¿é¤J¼Æ¦r¡I", vbExclamation
         txtCP113.SetFocus
         txtCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   Cancel = Not PUB_CheckCP113(txtCP113, pa(1), m_CP10, txtCP14)
End Sub

Private Sub txtCP14_Change()
   Label3(6) = ""
End Sub

Private Sub txtCP14_GotFocus()
   TextInverse txtCP14
End Sub

Private Sub txtCP14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCP14_Validate(Cancel As Boolean)
   If txtCP14 = "" Then
      MsgBox "©Ó¿ì¤H¤£¥iªÅ¥Õ !", vbCritical
      Cancel = True
   Else
      'ADD BY SONIA 2015/9/21 ©Ó¿ì¤H¬°¥~±Mµ{§Ç®É,§ï¬°¾Þ§@¤H­û
      txtCP14 = GetFCPUser(txtCP14)
      'END 2015/9/21
      Label3(6) = GetStaffName(txtCP14, True)
   End If
End Sub

Private Sub txtCP27_GotFocus()
   TextInverse txtCP27
End Sub

Private Sub txtCP27_Validate(Cancel As Boolean)
   If txtCP27 = "" Then
      MsgBox "µo¤å¤é¤£¥iªÅ¥Õ !", vbCritical
      Cancel = True
   ElseIf Not ChkDate(txtCP27) Then
      txtCP27_GotFocus
      Cancel = True
   End If
End Sub

'Added by Lydia 2015/12/31 ®×¥ó©Ê½è¬°·|½Z924®É¡AÀË¬d¤@©w­n¦³¬ÛÃöÁ`¦¬¤å¸¹;¼W¥[«ö¶s¨Ñ¨Ï¥ÎªÌ¿ï¾Ü;
Private Sub Command2_Click()
Dim bCancel As Boolean
    Set frm060101_2.fmParent = Me
    frm060101_2.Show
    Me.Hide
End Sub
Private Sub txtSC01_GotFocus()
   TextInverse txtSC01
End Sub

Private Sub txtSC01_Validate(Cancel As Boolean)
   If txtSC01 = "" Then
      MsgBox "°lÂÜ·|½Zµ²ªG´Á­­¤£¥iªÅ¥Õ !", vbCritical
      GoTo JumpCancel
   ElseIf Not ChkDate(txtSC01) Then
      GoTo JumpCancel
   ElseIf txtSC01 < strSrvDate(2) Then
      MsgBox "°lÂÜ·|½Zµ²ªG´Á­­¤£¥i¤p©ó¨t²Î¤é !", vbCritical
      GoTo JumpCancel
   End If
   Exit Sub
   
JumpCancel:
    txtSC01.SetFocus
    txtSC01_GotFocus
    Cancel = True
End Sub
Private Sub txtCP43_GotFocus()
   TextInverse txtCP43
End Sub

Private Sub txtCP43_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'end 2015/12/31

'Added by Lydia 2016/01/21
Private Sub txtCP43_Change()
   If Len(txtCP43) = 9 Then
      CheckCP43 True, False
   End If
End Sub
'Modified by Lydia 2016/01/21 ­ì¥»¬°txtCP43_Validate,§ï¬°§PÂ_¼Ò²Õ
Public Sub CheckCP43(ByVal bolMsg As Boolean, Optional Cancel As Boolean = False)
Dim strA1 As String
Dim rsAD As New ADODB.Recordset
Dim m_CP43cpm As String
   
   bolAddSC = False
   'Modified by Lydia 2019/07/03 §ï¦¨¥u°w¹ï011A¤é¥»°Ï¤§¥N²z¤Hªº·s¥Ó½Ð®×¤§·|½Z
   'If m_CP10 = "924" And Me.txtCP43 <> "" Then
   'Remove by Lydia 2019/11/13 ¨ú®ø¤é¥»¥N²z¤H011(A¦r¥À¶}ÀY)°lÂÜ·|½Z¤§¦æ¨Æ¾ä´Á­­¼u¸õ¤§±±ºÞ
'   If m_CP10 = "924" And Me.txtCP43 <> "" And mFA10 = "011" Then
'        strA1 = "select cp06,cp10,cp27,cp57 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp09='" & Me.txtCP43 & "' "
'        intI = 1
'        Set rsAD = ClsLawReadRstMsg(intI, strA1)
'        If intI = 1 Then
'           m_CP43cpm = rsAD.Fields("cp10")
'           If bolMsg = True Then
'                If InStr(NewCasePtyList, m_CP43cpm) > 0 Then
'                   '¸ß°Ý¬O§_²£¥Í¦æ¨Æ¾äºÞ¨î
'                   'Modified by Lydia 2019/07/03 ¼u¸õ´£¿ô¡u½Ð½T»{¬O§_ºÞ¨î·|½Zµ²ªG´Á­­¡v¤£¦Û°Ê²£¥Í¦æ¨Æ“ï¡A¥Ñ¤H­û¦Û¦æ§PÂ_¡C
'                   'If Me.txtSC01.Visible = False Then
'                   '   If MsgBox("¬O§_²£¥ÍºÞ¨î·|½Zµ²ªG´Á­­?", vbInformation + vbYesNo) = vbYes Then
'                   '      bolAddSC = True
'                   '   End If
'                   'Else
'                   '   bolAddSC = True
'                   'End If
'                   MsgBox "½Ð½T»{¬O§_ºÞ¨î·|½Zµ²ªG´Á­­!!!", vbInformation + vbOKOnly
'                   'end 2019/07/03
'                End If
'                If bolAddSC = True Then
'                   Me.txtSC01.Visible = True: Me.Label1(15).Visible = True
'                Else
'                   Me.txtSC01.Visible = False: Me.Label1(15).Visible = False
'                End If
'           End If
'        Else
'           MsgBox "¬ÛÃöÁ`¦¬¤å¸¹¤£¦s¦b!!", vbCritical
'           Cancel = True
'        End If
'   End If
   'end 2019/11/13
   Exit Sub
End Sub

