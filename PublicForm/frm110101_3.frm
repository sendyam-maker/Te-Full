VERSION 5.00
Begin VB.Form frm110101_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "ＦＭＰ解除期限"
   ClientHeight    =   1125
   ClientLeft      =   75
   ClientTop       =   945
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   6930
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   2
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   3
      Top             =   240
      Width           =   492
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   1
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   2
      Top             =   240
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   0
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   1
      Top             =   240
      Width           =   1212
   End
   Begin VB.TextBox txtKind 
      Height          =   264
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtSystem 
      Height          =   264
      Left            =   1050
      MaxLength       =   3
      TabIndex        =   0
      Top             =   240
      Width           =   732
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5040
      TabIndex        =   5
      Top             =   480
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6000
      TabIndex        =   6
      Top             =   480
      Width           =   800
   End
   Begin VB.Label Label2 
      Caption         =   "FMP解除期限(1.主動補正 2.香港第一階段請求)："
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   645
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   285
      Width           =   975
   End
End
Attribute VB_Name = "frm110101_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/5/7 改成Form2.0(無)
 'Add by Lydia 2014/10/14 FMP案(補正203和第一階段請求110)可不必輸入申請案號和證書,預設不印指示信和不詢問是否要作結餘
'ＦＭＰ解除期限 = 內專程序特定解除期限
Public m101_3_NP01 As String, m101_3_NP07 As String, m101_3_NP22 As String

Option Explicit


Private Function CheckCPOk() As Boolean
'檢查是否有補正203和第一階段請求110的記錄

'Added by Lydia 2015/08/26
   If InStr(FMP2openSQL, "not") = 0 Then
      strExc(2) = "該案號屬於非寰華案或"
   Else
      strExc(2) = "該案號屬於寰華案或"
   End If
   
If txtKind = "1" Then
   strExc(0) = " and NP07='203' "
   'Modified by Lydia 2015/08/26
   'strExc(2) = " 該案號無主動補正的期限！！"
   strExc(2) = strExc(2) & "無主動補正的期限！！"
Else
   strExc(0) = " and NP07='110' "
   'Modified by Lydia 2015/08/26
   strExc(2) = strExc(2) & "該案號無香港第一階段請求的期限！！"
End If
'Modified by Lydia 2015/08/26 設別名f0,+寰華案控管
'Modified by Lydia 2018/06/05 修改顯示案件性質
'strExc(1) = "select NP01 總收文號,substr(cp05,1,4)-1911||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) 來函收文日,decode('020','020',cpm04,cpm03) 下一程序, " & _
          "decode(np08,null,'',substr(np08,1,4)-1911||'/'||substr(np08,5,2)||'/'||substr(np08,7,2)) 本所期限, " & _
          "decode(np09,null,'',substr(np09,1,4)-1911||'/'||substr(np09,5,2)||'/'||substr(np09,7,2)) 法定期限,staff.st02 承辦人,staff1.st02 智權人員, " & _
          "np13 機關文號,np14 相關人,NP07,NP22,NP15 備註,decode(cp01||substr(cp12,1,1),'PF','Y','N') as FMP案 " & _
          "from caseprogress f0,nextprogress,casepropertymap,staff,staff staff1 where np02=cpm01(+) and np07=cpm02(+) " & _
          "and cp14=staff.st01(+) and np10=staff1.st01(+) and np06 is null and NP02='" & Trim(txtSystem) & "' " & _
          "and NP03='" & Trim(txtCode(0)) & "' and NP04='" & Trim(txtCode(1)) & "' and NP05='" & Trim(txtCode(2)) & "' and np01=CP09(+) " & strExc(0) & FMP2openSQL
strExc(1) = "select NP01 總收文號,substr(cp05,1,4)-1911||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) 來函收文日,decode(pa09,'000',cpm03,cpm04) 下一程序, " & _
          "decode(np08,null,'',substr(np08,1,4)-1911||'/'||substr(np08,5,2)||'/'||substr(np08,7,2)) 本所期限, " & _
          "decode(np09,null,'',substr(np09,1,4)-1911||'/'||substr(np09,5,2)||'/'||substr(np09,7,2)) 法定期限,staff.st02 承辦人,staff1.st02 智權人員, " & _
          "np13 機關文號,np14 相關人,NP07,NP22,NP15 備註,decode(cp01||substr(cp12,1,1),'PF','Y','N') as FMP案 " & _
          "from caseprogress f0,nextprogress,casepropertymap,staff,staff staff1,patent where np02=cpm01(+) and np07=cpm02(+) " & _
          "and cp14=staff.st01(+) and np10=staff1.st01(+) and np06 is null and NP02='" & Trim(txtSystem) & "' " & _
          "and NP03='" & Trim(txtCode(0)) & "' and NP04='" & Trim(txtCode(1)) & "' and NP05='" & Trim(txtCode(2)) & "' and np01=CP09(+) " & strExc(0) & FMP2openSQL & _
          " and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) "
intI = 1
Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
If intI = 1 Then
    If RsTemp.Fields("FMP案") = "N" Then
      MsgBox "該案號不屬於FMP案！！"
      CheckCPOk = False
    Else
      m101_3_NP01 = RsTemp.Fields("總收文號")
      m101_3_NP07 = RsTemp.Fields("NP07")
      m101_3_NP22 = RsTemp.Fields("NP22")
      CheckCPOk = True
    End If
Else
    MsgBox strExc(2)
    CheckCPOk = False
End If

If CheckCPOk = False Then
   txtCode(0).SetFocus
End If
End Function



Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer, bolRt As Boolean

Select Case Index
        Case 0
             If Len(txtCode(1)) = 0 Then
                txtCode(1) = "0"
             End If
             If Len(txtCode(2)) = 0 Then
                txtCode(2) = "00"
             End If
             If txtKind <> "1" And txtKind <> "2" Then
               MsgBox "請輸入正確條件！"
               txtKind.SetFocus
               Exit Sub
             End If
             If txtSystem <> "P" Or txtCode(0) = "" Then
               MsgBox "請輸入本所案號！！"
               txtKind.SetFocus
               Exit Sub
             End If

            If CheckCPOk Then

               Set frm110101_2.mPrev01 = Me
               
               frm110101_2.Show
               Me.Hide
            End If
        Case 1
            Unload Me

End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me

txtKind = "1"
txtSystem = "P"

'Added by Lydia 2015/08/26 請加入控制區分寰華案及非寰華案,分別由外專或內專人員輸入
FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05, "INVERSE_SQL")
If Pub_StrUserSt03 = "M51" Then
   If MsgBox("電腦中心人員請注意你現在是要看FMP寰華案嗎?", vbYesNo + vbInformation + vbDefaultButton1) = vbYes Then
      FMP2openSQL = Replace(FMP2openSQL, "not", "")
   Else
      MsgBox "現在本程式不可使用FMP寰華案"
   End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)

   Set frm110101_3 = Nothing
End Sub

Private Sub txtkind_GotFocus()
If Len(txtCode(1)) = 0 Then
   txtCode(1) = "0"
End If
If Len(txtCode(2)) = 0 Then
   txtCode(2) = "00"
End If

TextInverse txtKind
End Sub

Private Sub txtSystem_GotFocus()
'txtSystem.SelStart = 0
'txtSystem.SelLength = Len(txtSystem.Text)
TextInverse txtSystem
End Sub
Private Sub txtSystem_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtSystem_Validate(Cancel As Boolean)
'Modified by Lydia 2015/09/03 程式本就限P案,改成直接判斷,毋須開系統給外專程序
'If ClsPDGetGroupCase(txtSystem, strGroup) = False Then
If txtSystem <> "P" Then
   ShowMsg MsgText(1056)
   Cancel = True
   txtSystem_GotFocus
End If
End Sub
'
'Private Sub txtTFCode_GotFocus(Index As Integer)
'txtTFCode(Index).SelStart = 0
'txtTFCode(Index).SelLength = Len(txtTFCode(Index).Text)
'End Sub

Private Sub txtCode_GotFocus(Index As Integer)
 TextInverse txtCode(Index)
End Sub



Public Sub Cleartxt()
txtCode(0) = ""
txtCode(1) = ""
txtSystem = ""
txtCode(2) = ""
txtKind = "1"

txtSystem.SetFocus
End Sub


 'Add by Lydia 2014/10/14 FMP案(補正203和第一階段請求110)可不必輸入申請案號和證書,預設不印指示信和不詢問是否要作結餘
'ＦＭＰ解除期限 .end
