VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040326 
   BorderStyle     =   1  '單線固定
   Caption         =   "年費逾期補繳通知函(整批)"
   ClientHeight    =   1848
   ClientLeft      =   120
   ClientTop       =   948
   ClientWidth     =   4236
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1848
   ScaleWidth      =   4236
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   276
      TabIndex        =   4
      Top             =   1188
      Visible         =   0   'False
      Width           =   3756
      _ExtentX        =   6625
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtDate 
      Height          =   315
      Left            =   1260
      MaxLength       =   9
      TabIndex        =   3
      Top             =   672
      Width           =   870
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   2445
      TabIndex        =   0
      Top             =   624
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   3270
      TabIndex        =   1
      Top             =   624
      Width           =   800
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1248
      TabIndex        =   6
      Top             =   168
      Visible         =   0   'False
      Width           =   1968
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3471;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Left            =   264
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblCount 
      Alignment       =   2  '置中對齊
      Caption         =   "( 0 / 0 )"
      Height          =   168
      Left            =   1308
      TabIndex        =   5
      Top             =   1512
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.Label Label1 
      Caption         =   "收文日期："
      Height          =   180
      Left            =   276
      TabIndex        =   2
      Top             =   732
      Width           =   900
   End
End
Attribute VB_Name = "frm040326"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Create by Morgan 2010/3/8
Option Explicit

Private Sub cmdok_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 0
         cmdOK(Index).Enabled = False
         If TxtValidate Then
            Process
         End If
         cmdOK(Index).Enabled = True
      Case 2
         Unload Me
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtDate = strSrvDate(2)
   
   'Added by Morgan 2025/1/15
   If strSrvDate(1) >= P業務區劃分啟用日 Then
      Combo1.Visible = True
      Label4.Visible = True
      Call SetPatentP12Combo(Combo1, "P", Label4)
   End If
   'end 2025/1/15
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040326 = Nothing
End Sub

Private Sub txtDate_GotFocus()
   CloseIme
   TextInverse txtDate
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
   If txtDate = "" Then
      MsgBox "收文日期不可空白！"
      Cancel = True
   ElseIf ChkDate(txtDate) = False Then
      Cancel = True
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   
   txtDate_Validate Cancel
   If Cancel = True Then
      txtDate.SetFocus
      Exit Function
   End If
   TxtValidate = True
End Function

Private Sub Process()
   Dim adoRst As ADODB.Recordset
   Dim stCP09 As String, stCP06 As String, stCP07 As String, stCP64 As String
   Dim ET01 As String, ET03 As String
   Dim iPos As Integer, stNextYear As String, stOldDate As String
   Dim strNextTime As String
   Dim stCon As String 'Added by Morgan 2025/1/16
   
   'Added by Morgan 2025/1/16
   If Combo1.Visible And Combo1 <> "" Then
      stCon = stCon & " and cp14='" & Left(Combo1, 5) & "'"
   End If
   'end 2025/1/16
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/30 清除查詢印表記錄檔欄位
   '只適用台灣
   ET01 = "19"
   'ET03 = "01"
   pub_QL05 = pub_QL05 & ";" & Label1 & txtDate 'Add By Sindy 2010/11/30
   'Modified by Lydia 2015/01/08 + 區別台灣案
'   strExc(0) = "select cp09,cp06,cp07,cp64,pa08,pa09,cf28,pa01,pa02,pa03,pa04,pa26,pa75" & _
'      " from caseprogress,patent,casefee where cp01='P' and cp10='1605' and cp27 is null and cp05=" & DBDATE(txtDate) & _
'      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
'      " and cf01(+)=pa01 and cf02(+)=pa09 and cf03='605' order by cp09"
   'Memo by Morgan 2015/9/1 整批列印定稿有另外控制列印順序(同接洽人一起,另年費/實審通知也有)
   strExc(0) = "select cp09,cp06,cp07,cp64,pa08,pa09,cf28,pa01,pa02,pa03,pa04,pa26,pa75" & _
      " from caseprogress,patent,casefee where cp01='P' and cp10='1605' and cp27 is null and cp05=" & DBDATE(txtDate) & stCon & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and cf01(+)=pa01 and cf02(+)=pa09 and cf03='605' and pa09 = '" & 台灣國家代號 & "' order by cp09"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoRst
      InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/30
      ProgressBar1.Min = 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      ProgressBar1.Visible = True
      lblCount.Visible = True
      DoEvents
      
      .MoveFirst
      Do While Not .EOF
         stCP09 = .Fields("cp09")
         stCP06 = "" & .Fields("cp06")
         stCP07 = "" & .Fields("cp07")
         stCP64 = "" & .Fields("cp64")
         strNextTime = "" & .Fields("cf28")
         
         'Modified by Morgan 2014/6/12
         ''Modified by Morgan 2012/12/6
         ''102新法
         'If strSrvDate(1) >= "20130101" Then
         '   ET03 = "04"
         'Else
         '   ET03 = "01"
         'End If
         ''end 2012/12/6
         '大對台定稿不同
         If PUB_CheckCuNation(.Fields("pa26"), .Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04")) = "1" Then
            ET03 = "05"
         Else
            ET03 = "04"
         End If
         'end 2014/6/12
         
         strSql = "update caseprogress set cp14='" & strUserNum & "',cp27=" & strSrvDate(1) & " where cp09='" & stCP09 & "'"
         cnnConnection.Execute strSql, intI
               
         'Added by Morgan 2014/6/27
         '新增信函進度
         'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
         'Modified by Morgan 2016/11/8 +傳是否大宗發文(pbolBulk=True)
         PUB_AddLetterProgress stCP09, 0, True, "", True, "" & .Fields("pa26"), "1605", "" & .Fields("pa75"), , , True
         'end 2014/6/27
               
         iPos = InStr(stCP64, "未繳年度:")
         If iPos > 0 Then
            stNextYear = Val(Mid(stCP64, iPos + 5))
         Else
            stNextYear = ""
         End If
         
         iPos = InStr(stCP64, "原繳費期限:")
         If iPos > 0 Then
            stOldDate = Val(Mid(stCP64, iPos + 6))
         Else
            stOldDate = ""
         End If
         
         EndLetter ET01, stCP09, ET03, strUserNum
               
         'Added by Morgan 2012/9/19
         'If ET03 = "04" Then 'Removed by Morgan 2014/6/12
         'Modified by Lydia 2015/01/07 採共用模組
'            strExc(0) = "Select YF06,YF07 From PatentYearFee Where YF01='" & .Fields("pa09") & "' AND YF02='" & .Fields("pa08") & "' AND YF03='Y00000001' AND YF04='605' AND YF05=" & Val(stNextYear)
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               '服務費
'               strExc(1) = "" & RsTemp("YF06")
'               '規費
'               strExc(2) = "" & RsTemp("YF07")
            strExc(0) = PUB_GetYF0607(.Fields("pa09"), .Fields("pa08"), .Fields("pa26"), "605", stNextYear, stNextYear, "1", strExc(1), strExc(2))
            If strExc(0) = "0" Then strExc(1) = "": strExc(2) = ""
            
            If Val(strExc(0)) > 0 Then
               If Val(stNextYear) < 7 Then
                  '可減免
                  If PUB_GetCaseDiscStat(.Fields("pa01") & .Fields("pa02") & .Fields("pa03") & .Fields("pa04")) = "Y" Then
                     If Val(stNextYear) > 3 Then
                        strExc(2) = Val(strExc(2)) - 1200
                     Else
                        strExc(2) = Val(strExc(2)) - 800
                     End If
                  End If
               End If
'            End If
            End If
            'end 'Modified by Lydia 2015/01/07 採共用模組
            'Added by Morgan 2013/2/21
            '專利處大對台年費服務費+500 --郭雅娟
            strExc(4) = PUB_GetStaffST15(PUB_GetAKindSalesNo(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04")), "1")
            If Left(strExc(4), 2) = "P1" Then
               strExc(1) = Val(strExc(1)) + 500
            End If
            'end 2013/2/21
            
            'Modified by Morgan 2013/2/21 要加服務費
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & stCP09 & "','" & ET03 & "','" & strUserNum & "','原年費','" & Format(Format(Val(strExc(1)) + Val(strExc(2)), "#,###"), String(6, "@")) & "')"
            cnnConnection.Execute strSql, intI
            
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & stCP09 & "','" & ET03 & "','" & strUserNum & "','年費1','" & Format(Format(Val(strExc(1)) + Val(strExc(2)) * 1.2, "#,###"), String(6, "@")) & "')"
            cnnConnection.Execute strSql, intI
            
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & stCP09 & "','" & ET03 & "','" & strUserNum & "','年費2','" & Format(Format(Val(strExc(1)) + Val(strExc(2)) * 1.4, "#,###"), String(6, "@")) & "')"
            cnnConnection.Execute strSql, intI
            
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & stCP09 & "','" & ET03 & "','" & strUserNum & "','年費3','" & Format(Format(Val(strExc(1)) + Val(strExc(2)) * 1.6, "#,###"), String(6, "@")) & "')"
            cnnConnection.Execute strSql, intI
            
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & stCP09 & "','" & ET03 & "','" & strUserNum & "','年費4','" & Format(Format(Val(strExc(1)) + Val(strExc(2)) * 1.8, "#,###"), String(6, "@")) & "')"
            cnnConnection.Execute strSql, intI
            
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & stCP09 & "','" & ET03 & "','" & strUserNum & "','年費5','" & Format(Format(Val(strExc(1)) + Val(strExc(2)) * 2, "#,###"), String(6, "@")) & "')"
            cnnConnection.Execute strSql, intI
         'End If 'Removed by Morgan 2014/6/12
         
         If stOldDate <> "" Then
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & stCP09 & "','" & ET03 & "','" & strUserNum & "','年費法定期限','" & DBDATE(stOldDate) & "')"
            cnnConnection.Execute strSql, intI
         End If
         
         If stNextYear <> "" Then
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & stCP09 & "','" & ET03 & "','" & strUserNum & "','第幾年至幾年費','" & stNextYear & "')"
            cnnConnection.Execute strSql, intI
         End If

         If stCP06 <> "" Then
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & stCP09 & "','" & ET03 & "','" & strUserNum & "','下次繳年費日','" & DBDATE(stCP06) & "')"
            cnnConnection.Execute strSql, intI
         End If
         
         If strNextTime <> "" Then
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & stCP09 & "','" & ET03 & "','" & strUserNum & "','列印備註','" & strNextTime & "')"
            cnnConnection.Execute strSql, intI
         End If
         
         NowPrint stCP09, ET01, ET03, False, strUserNum, , , , , , , , , , , , , stCP09
         
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblCount.Caption = "( " & ProgressBar1.Value & " / " & ProgressBar1.max & " )"
         DoEvents
         .MoveNext
      Loop
      End With
      MsgBox "定稿已產生！"
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/11/30
      MsgBox "無未發文之年費逾期補繳通知程序！"
   End If
   ProgressBar1.Visible = False
   lblCount.Visible = False
   Set adoRst = Nothing
End Sub
