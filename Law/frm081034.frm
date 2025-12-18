VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm081034 
   BorderStyle     =   1  '單線固定
   Caption         =   "TIPS案請款階段設定"
   ClientHeight    =   6084
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7812
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6084
   ScaleWidth      =   7812
   Begin VB.TextBox txtAmt 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      Height          =   270
      Left            =   6000
      MaxLength       =   8
      TabIndex        =   18
      Text            =   "txtAmt"
      Top             =   3864
      Width           =   870
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  '平面
      Height          =   270
      Left            =   4752
      MaxLength       =   3
      TabIndex        =   17
      Text            =   "txtYear"
      Top             =   3816
      Width           =   870
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  '平面
      Height          =   270
      Left            =   3168
      MaxLength       =   1
      TabIndex        =   16
      Text            =   "txtInput"
      Top             =   3792
      Width           =   870
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "存檔"
      Height          =   348
      Left            =   6552
      TabIndex        =   14
      Top             =   2808
      Width           =   1068
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
      Height          =   1344
      Left            =   24
      TabIndex        =   12
      Top             =   1296
      Width           =   7656
      _ExtentX        =   13504
      _ExtentY        =   2371
      _Version        =   393216
      AllowUserResizing=   3
      RowSizingMode   =   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   6768
      TabIndex        =   4
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "查詢(&Q)"
      Height          =   375
      Left            =   3504
      TabIndex        =   3
      Top             =   48
      Width           =   855
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   3
      Left            =   2952
      MaxLength       =   2
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   2
      Left            =   2544
      MaxLength       =   1
      TabIndex        =   1
      Top             =   120
      Width           =   345
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   1
      Left            =   1644
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   0
      Left            =   1092
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "ACS"
      Top             =   120
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid2 
      Height          =   2808
      Left            =   48
      TabIndex        =   13
      Top             =   3168
      Width           =   7656
      _ExtentX        =   13504
      _ExtentY        =   4953
      _Version        =   393216
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      Caption         =   "請在階段輸入1~9，請款年度輸入民國年，並且執行存檔才會寫入資料庫。"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   96
      TabIndex        =   15
      Top             =   2712
      Width           =   3780
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1452
      X2              =   3042
      Y1              =   312
      Y2              =   312
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1116
      TabIndex        =   11
      Top             =   864
      Width           =   6612
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11668;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   300
      Index           =   1
      Left            =   2016
      TabIndex        =   10
      Top             =   504
      Width           =   5532
      Size            =   "9763;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   300
      Index           =   0
      Left            =   1116
      TabIndex        =   9
      Top             =   504
      Width           =   888
      Size            =   "1561;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "當事人1："
      Height          =   228
      Index           =   3
      Left            =   96
      TabIndex        =   7
      Top             =   564
      Width           =   888
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   228
      Index           =   2
      Left            =   96
      TabIndex        =   6
      Top             =   888
      Width           =   948
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   228
      Index           =   0
      Left            =   96
      TabIndex        =   5
      Top             =   168
      Width           =   948
   End
End
Attribute VB_Name = "frm081034"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2024/03/25 Form2.0已修改; MGrid1、MGrid2、Combo1、lblFM2
Option Explicit
Dim intLastRow As Integer '記錄MGrid1勾選最後一筆
Dim intLastRow2 As Integer '記錄MGrid2勾選最後一筆
Dim nRow2 As Integer, nCol2 As Integer 'MGrid2本次點選列數,欄數

Dim m_blnCol2OrderAsc As Boolean '欄位資料由小到大排序
Dim strNowCP09 As String, strNowCP10 As String, strNowCP10name 'TIPS收文號和案件性質
Dim strListCP43 As String, strListAmt As String 'TIPS收文號,收文金額
Dim strTmpQ As String
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset
Dim colCP09_1 As Integer, colCP10_1 As Integer, colCP10_1name As Integer
Dim colCP09_2 As Integer, colCP10_2 As Integer, colCP10_2name As Integer
Dim colCP156_New As Integer, colCP156_Old As Integer
Dim colCP43 As Integer, colCP43cp10 As Integer 'Added by Lydia 2024/04/03
'Added by Lydia 2025/04/09 TIPS分配比例管制：先將請款年度開放給顧服組輸入
Dim colCP115_New As Integer, colCP115_Old As Integer
'Modified by Lydia 2025/05/19 6 => 8
Private Const cntFixed As Integer = 8
Dim passRow As Integer, passCol As Integer
'Added by Lydia 2025/05/19 提前輸入請款金額，並且代入對應的收據號碼
Dim colCP144_New As Integer, colCP144_Old As Integer
Dim colCP60_New As Integer, colCP60_Old As Integer
Dim colCP27 As Integer
Dim strListCP60_New As String '暫存未存檔的「新增設定」收據號
Dim strListCP60_Old As String '暫存未存檔的「刪除設定」收據號
'Dim bolAdd As Boolean 'Added by Lydia 2025/09/02 增加請款階段的範圍 'Mark by Lydia 2025/09/05 全部性質都可以，只有一筆收文可以列入請款階段

Private Sub CmdSave_Click()
Dim strExSql As String, strListCP156 As String
Dim bolConn As Boolean, tmpArr As Variant
Dim intM As Integer, strNum(1 To 9) As String
Dim mSubAmt As String, strMsgTitle As String

   strMsgTitle = "存檔前檢查輸入資料"
   txtInput.Visible = False
   txtYear.Visible = False 'Added by Lydia 2025/04/09
   txtAmt.Visible = False 'Added by Lydia 2025/05/19
   
   If Trim(txtCase(0).Text) = "" Or Len(Trim(txtCase(1).Text)) < 6 Then
      MsgBox "請輸入本所案號！", vbExclamation, strMsgTitle
      Exit Sub
   End If
   If txtCase(0) & txtCase(1) & txtCase(2) & txtCase(3) <> txtCase(0).Tag & txtCase(1).Tag & txtCase(2).Tag & txtCase(3).Tag Then
      MsgBox "輸入本所案號後，請執行查詢功能！", vbExclamation, strMsgTitle
      Exit Sub
   End If

   strListCP156 = ""
   
   For intQ = 1 To MGrid2.Rows - 1
      If "" & MGrid2.TextMatrix(intQ, colCP09_2) <> "" And Trim("" & MGrid2.TextMatrix(intQ, colCP156_New) & MGrid2.TextMatrix(intQ, colCP156_Old)) <> "" Then
         If "" & MGrid2.TextMatrix(intQ, colCP144_New) <> "" Then
            mSubAmt = Val(mSubAmt) + Pub_GetCP144Val(txtCase(0), txtCase(1), txtCase(2), txtCase(3), "2", MGrid2.TextMatrix(intQ, colCP144_New))
         End If
         If Val(mSubAmt) > Val(strListAmt) Then
            MsgBox "請款總金額超過【TIPS" & strNowCP10name & "】費用＝" & Format(strListAmt, "##,##0"), vbExclamation + vbOKOnly, strMsgTitle
            Exit Sub
         End If
         If Val("" & MGrid2.TextMatrix(intQ, colCP156_New)) < 0 Or Val("" & MGrid2.TextMatrix(intQ, colCP156_New)) > 9 Then
            MsgBox "請款階段請輸入1~9 ！", vbExclamation, strMsgTitle
            Exit Sub
         End If
         'Added by Lydia 2025/04/09 已有請款金額欄應限制不可改階段、年度
         'Modified by Lydia 2025/05/19 改用發文日判斷
         'If Val(Pub_GetCP144Val(txtCase(0), txtCase(1), txtCase(2), txtCase(3), "2", MGrid2.TextMatrix(intQ, colCP144_New))) > 0 And ("" & MGrid2.TextMatrix(intQ, colCP156_New) <> "" & MGrid2.TextMatrix(intQ, colCP156_Old) _
         '      Or "" & MGrid2.TextMatrix(intQ, colCP115_New) <> "" & MGrid2.TextMatrix(intQ, colCP115_Old)) Then
         '   '排除:補請款年度
         '   If "" & MGrid2.TextMatrix(intQ, colCP156_New) = "" & MGrid2.TextMatrix(intQ, colCP156_Old) And "" & MGrid2.TextMatrix(intQ, colCP115_New) <> "" And "" & MGrid2.TextMatrix(intQ, colCP115_Old) = "" Then
         '   Else
         '      MsgBox "已有請款金額不可修改請款階段、年度！", vbExclamation,  strMsgTitle
         '      Exit Sub
         '   End If
         'End If
         'end 2025/04/09
         If Trim("" & MGrid2.TextMatrix(intQ, colCP27)) <> "" And ("" & MGrid2.TextMatrix(intQ, colCP156_New) <> "" & MGrid2.TextMatrix(intQ, colCP156_Old) Or "" & MGrid2.TextMatrix(intQ, colCP115_New) <> "" & MGrid2.TextMatrix(intQ, colCP115_Old) Or "" & MGrid2.TextMatrix(intQ, colCP144_New) <> "" & MGrid2.TextMatrix(intQ, colCP144_Old)) Then
             MsgBox "已發文不可修改請款階段、年度！", vbExclamation, strMsgTitle
             Exit Sub
         End If
         'end 2025/05/19
         
         If Val("" & MGrid2.TextMatrix(intQ, colCP156_New)) > 0 Then
            If "" & MGrid2.TextMatrix(intQ, colCP10_2) = "706" Then
               MsgBox "【" & MGrid2.TextMatrix(intQ, colCP10_2name) & "】不可輸入請款階段 ！", vbExclamation, strMsgTitle
               Exit Sub
            End If
            'Added by Lydia 2025/04/09
            If "" & MGrid2.TextMatrix(intQ, colCP144_New) <> "" Then
               If Val("" & MGrid2.TextMatrix(intQ, colCP115_New)) > Val(Mid(strSrvDate(2), 1, 3)) + 1 Then
                  If MsgBox("請款年度超過下一年度，是否繼續存檔？", vbYesNo + vbDefaultButton2 + vbExclamation, strMsgTitle) = vbNo Then
                     Exit Sub
                  End If
               End If
            Else
               If Val("" & MGrid2.TextMatrix(intQ, colCP115_New)) = 0 Then
                  MsgBox "請款年度不可空白 ！", vbExclamation, strMsgTitle
                  Exit Sub
               End If
               If Val("" & MGrid2.TextMatrix(intQ, colCP115_New)) < Mid(strSrvDate(2), 1, 3) Then
                  MsgBox "請款年度不可輸入過去年度 ！", vbExclamation, strMsgTitle
                  Exit Sub
               End If
               If Val("" & MGrid2.TextMatrix(intQ, colCP115_New)) > Val(Mid(strSrvDate(2), 1, 3)) + 1 Then
                  If MsgBox("請款年度超過下一年度，是否繼續存檔？", vbYesNo + vbDefaultButton2 + vbExclamation, strMsgTitle) = vbNo Then
                     Exit Sub
                  End If
               End If
               'Added by Lydia 2025/05/19
               If Val("" & MGrid2.TextMatrix(intQ, colCP144_New)) = 0 Then
                  MsgBox "請款金額不可空白 ！", vbExclamation, strMsgTitle
                  Exit Sub
               End If
               '沒有對應的收據
               If Trim("" & MGrid2.TextMatrix(intQ, colCP60_New)) = "" Then
                  strTmpQ = Pub_ACS_TIPS_GetCp60("2", txtCase(0), txtCase(1), txtCase(2), txtCase(3), "" & MGrid2.TextMatrix(intQ, colCP144_New), strListCP60_Old, strListCP60_Old)
                  MsgBox "請輸入正確的請款金額 ！", vbExclamation, strMsgTitle
                  Exit Sub
               End If
               'end 2025/05/19
            End If
            'end 2025/04/09
            
            'Added by Lydia 2024/04/03
            If "" & MGrid2.TextMatrix(intQ, colCP43) = "" Then
               MsgBox "【" & MGrid2.TextMatrix(intQ, colCP10_2name) & "】請先設定相關總收文 ！", vbExclamation, strMsgTitle
               Exit Sub
            Else
               If InStr(ACSforTIPSstep, "'" & MGrid2.TextMatrix(intQ, colCP43cp10) & "'") = 0 And Mid(MGrid2.TextMatrix(intQ, colCP43), 1, 1) <> "C" Then
                  MsgBox "【" & MGrid2.TextMatrix(intQ, colCP10_2name) & "】請先設定相關總收文為TIP案件性質或官方來函 ！", vbExclamation, strMsgTitle
                  Exit Sub
               End If
            End If
            'end 2024/04/03
            
            If InStr(strListCP156 & ",", MGrid2.TextMatrix(intQ, colCP156_New)) > 0 Then
               MsgBox "請款階段重覆輸入！" & vbCrLf & MGrid2.TextMatrix(intQ, colCP09_2) & " " & MGrid2.TextMatrix(intQ, colCP10_2name), vbExclamation, strMsgTitle
               Exit Sub
            End If
            strListCP156 = strListCP156 & Val("" & MGrid2.TextMatrix(intQ, colCP156_New)) & ","
            If intM < Val("" & MGrid2.TextMatrix(intQ, colCP156_New)) Then
               intM = Val("" & MGrid2.TextMatrix(intQ, colCP156_New))
            End If
            strNum(Val("" & MGrid2.TextMatrix(intQ, colCP156_New))) = "Y"
         'Added by Lydia 2025/04/09
         Else
            If Val("" & MGrid2.TextMatrix(intQ, colCP115_New)) <> 0 Then
               MsgBox "請款年度不可輸入 ！", vbExclamation, strMsgTitle
               Exit Sub
            End If
         'end 2025/04/09
            'Added by Lydia 2025/05/19
            If Val("" & MGrid2.TextMatrix(intQ, colCP144_New)) <> 0 Then
               MsgBox "請款金額不可輸入 ！", vbExclamation, strMsgTitle
               Exit Sub
            End If
            'end 2025/05/19
         End If
         
         'Modified by Lydia 2025/04/09 +請款年度 CP115
         'Modified by Lydia 2025/05/19 +請款金額 CP144
         If (Trim("" & MGrid2.TextMatrix(intQ, colCP156_New)) <> Trim(MGrid2.TextMatrix(intQ, colCP156_Old))) Or (Trim("" & MGrid2.TextMatrix(intQ, colCP115_New)) <> Trim(MGrid2.TextMatrix(intQ, colCP115_Old))) Or (Trim("" & MGrid2.TextMatrix(intQ, colCP144_New)) <> Trim(MGrid2.TextMatrix(intQ, colCP144_Old))) Then
            If strNowCP10 = "1014" And Val("" & MGrid2.TextMatrix(intQ, colCP156_New)) > 1 Then
               If MsgBox("【" & strNowCP10name & "(" & strNowCP10 & ")】請款階段原則為一個，請問是否繼續輸入作業？" & vbCrLf & MGrid2.TextMatrix(intQ, colCP09_2) & "目前輸入階段：" & "" & MGrid2.TextMatrix(intQ, colCP156_New), vbInformation + vbYesNo + vbDefaultButton2, strMsgTitle) = vbNo Then
                  Exit Sub
               End If
            End If
            If Val("" & MGrid2.TextMatrix(intQ, colCP156_New)) = 0 And Val("" & MGrid2.TextMatrix(intQ, colCP144_New)) > 0 Then
               If MsgBox("【" & MGrid2.TextMatrix(intQ, colCP10_2name) & "(" & MGrid2.TextMatrix(intQ, colCP09_2) & ")】已有請款金額，請問是否要取消請款階段？" & vbCrLf & "原本輸入階段：" & "" & MGrid2.TextMatrix(intQ, colCP156_Old), vbInformation + vbYesNo + vbDefaultButton2, strMsgTitle) = vbNo Then
                  Exit Sub
               End If
            End If
            If Trim("" & MGrid2.TextMatrix(intQ, colCP156_New)) = "" And Trim("" & MGrid2.TextMatrix(intQ, colCP156_Old)) <> "" Then
               'Modified by Lydia 2025/04/09 +cp115=null
               'Modified by Lydia 2025/05/19 +cp144=null,cp60=null
               strExSql = strExSql & "Update Caseprogress Set CP156=null,cp115=null,cp144=null,cp60=null Where CP09='" & MGrid2.TextMatrix(intQ, colCP09_2) & "'|"
            Else
               'Modified by Lydia 2025/04/09 +cp115
               'Modified by Lydia 2025/05/19
               'strExSql = strExSql & "Update Caseprogress Set CP156='" & MGrid2.TextMatrix(intQ, colCP156_New) & "',CP115='" & MGrid2.TextMatrix(intQ, colCP115_New) & "' Where CP09='" & MGrid2.TextMatrix(intQ, colCP09_2) & "';"
               If "" & MGrid2.TextMatrix(intQ, colCP144_New) <> "" Then
                  strTmpQ = "'" & Pub_GetCP144Val(txtCase(0), txtCase(1), txtCase(2), txtCase(3), "0", "" & MGrid2.TextMatrix(intQ, colCP144_New)) & "'"
               Else
                  strTmpQ = "null"
               End If
               strExSql = strExSql & "Update Caseprogress Set CP156='" & MGrid2.TextMatrix(intQ, colCP156_New) & "',CP115='" & MGrid2.TextMatrix(intQ, colCP115_New) & "',CP144=" & strTmpQ & ",CP60='" & Left(MGrid2.TextMatrix(intQ, colCP60_New), 9) & "' Where CP09='" & MGrid2.TextMatrix(intQ, colCP09_2) & "'|"
               'end 2025/05/19
            End If
         End If
      End If
   Next intQ
   If intM > 1 Then
     For intQ = intM To 1 Step -1
        If strNum(intQ) <> "Y" Then
           MsgBox "請款階段未輸入" & intQ, vbExclamation, strMsgTitle
           Exit Sub
        End If
     Next intQ
   End If
   
   If strExSql = "" Then
      MsgBox "無資料變更！", vbInformation, strMsgTitle
      Exit Sub
   End If
   cmdSave.Enabled = False
   Screen.MousePointer = vbHourglass
   tmpArr = Split(strExSql, "|")
   For intQ = 0 To UBound(tmpArr)
      If Trim(tmpArr(intQ)) <> "" Then
         If bolConn = False Then
            bolConn = True
            cnnConnection.BeginTrans
         End If
         Pub_SeekTbLog tmpArr(intQ)
         cnnConnection.Execute tmpArr(intQ)
      End If
   Next intQ
   bolConn = False
   cnnConnection.CommitTrans
   Screen.MousePointer = vbDefault
   cmdSave.Enabled = True
   MsgBox "存檔完畢！", vbInformation
   Call doQuery(False)
   Exit Sub
   
ErrHandle:
   Screen.MousePointer = vbDefault
   If bolConn = True Then
      cnnConnection.RollbackTrans
   End If
   If Err.Number <> 0 Then
      MsgBox Err.Description, , "存檔失敗"
   End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    CmdQuery.Default = False
    Call doQuery(True)
End Sub

Private Sub doQuery(ByVal bolMsg As Boolean)

    'bolAdd = False 'Added by Lydia 2025/09/02 'Mark by Lydia 2025/09/05
    
    If Trim(txtCase(0).Text) = "" Or Len(Trim(txtCase(1).Text)) < 6 Then
        MsgBox "請輸入本所案號！", vbExclamation, "檢核資料"
        Exit Sub
    End If
    
    Call ClearForm(False)
    If txtCase(2) = "" Then txtCase(2) = "0"
    If txtCase(3) = "" Then txtCase(3) = "00"
    txtCase(0).Tag = txtCase(0).Text
    txtCase(1).Tag = txtCase(1).Text
    txtCase(2).Tag = txtCase(2).Text
    txtCase(3).Tag = txtCase(3).Text
    
    strTmpQ = "select lc01,lc02,lc03,lc04,lc05,lc06,lc07,lc11 as custno,nvl(cu04,nvl(cu05,cu06)) custname " & _
                    "from lawcase,customer where lc01='" & txtCase(0) & "' and lc02='" & txtCase(1) & "' and lc03='" & txtCase(2) & "' and lc04='" & txtCase(3) & "' and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) "
    intQ = 0
    Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
    If intQ = 0 Then
        Exit Sub
    End If
    
    intQ = 0
    Combo1.AddItem "中：" & rsQuery.Fields("lc05"), 0
    If "" & rsQuery.Fields("lc05") <> "" And intQ = 0 Then intQ = 1
    Combo1.AddItem "英：" & rsQuery.Fields("lc06"), 1
    If "" & rsQuery.Fields("lc06") <> "" And intQ = 0 Then intQ = 2
    Combo1.AddItem "日：" & rsQuery.Fields("lc07"), 2
    If "" & rsQuery.Fields("lc07") <> "" And intQ = 0 Then intQ = 3
    Combo1.ListIndex = intQ - 1
    
    lblFM2(0).Caption = "" & rsQuery.Fields("custno")
    lblFM2(1).Caption = "" & rsQuery.Fields("custname")
    
    Call SetGrd1(True) '清空
    Call SetGrd2(True)
    
    'Modified by Lydia 2025/09/02 'Remark by Lydia 2025/09/05 還原
    strTmpQ = "select cp09,cp16 from caseprogress where cp01='" & txtCase(0) & "' and cp02='" & txtCase(1) & "' and cp03='" & txtCase(2) & "' and cp04='" & txtCase(3) & "'" & _
              " and cp10 in (" & ACSforTIPSstep & ") and cp159=0 order by cp09 "
    'strTmpQ = "select '1' as ord1,cp09,cp10,cp16 from caseprogress where cp01='" & txtCase(0) & "' and cp02='" & txtCase(1) & "' and cp03='" & txtCase(2) & "' and cp04='" & txtCase(3) & "' and cp10 in (" & ACSforTIPSstep & ") and cp159=0 " & _
              "union select '2' as ord1,cp09,cp10,cp16 from caseprogress where cp01='" & txtCase(0) & "' and cp02='" & txtCase(1) & "' and cp03='" & txtCase(2) & "' and cp04='" & txtCase(3) & "' and cp10 in (" & ACSforLetter & ") and cp159=0 and cp31='Y' " & _
              "order by ord1,cp09 "
    intQ = 1
    strListCP43 = ""
    Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
    If intQ = 1 Then
       rsQuery.MoveFirst
       'Added by Lydia 2025/09/02
       'Mark by Lydia 2025/09/05 全部性質都可以，只有一筆收文可以列入請款階段
'       strTmpQ = "" & rsQuery.Fields("ord1")
'       If strTmpQ = "1" And InStr(ACSforLetter, "" & rsQuery.Fields("cp10")) > 0 Then
'          bolAdd = True
'       Else
'          If strTmpQ = "2" Then bolAdd = True
'       End If
       'end 2025/09/02
       'end 2025/09/05
       Do While Not rsQuery.EOF
          'If strTmpQ = "" & rsQuery.Fields("ord1") Then 'Added by Lydia 2025/09/02 'Mark by Lydia 2025/09/05
             strListCP43 = strListCP43 & "," & rsQuery.Fields("cp09")
             strListAmt = Val(strListAmt) + Val("" & rsQuery.Fields("cp16"))
          'End If 'Added by Lydia 2025/09/02
          rsQuery.MoveNext
       Loop
       strListCP43 = GetAddStr(Mid(strListCP43, 2))
    End If
    
    '以本所案號抓出TIPS進度(有收費)列示於Grid中
    strTmpQ = " select '' as v,substr(sqldatet(cp05),1,10) cp05t,cp09,cpm03,s1.st02 as cp13n,s2.st02 as cp14n,cp16,cp18,cp10,substr(sqldatet(cp27),1,10) cp27t" & _
             " from caseprogress,casepropertymap,staff s1,staff s2 where cp01='" & txtCase(0) & "' and cp02='" & txtCase(1) & "' and cp03='" & txtCase(2) & "' and cp04='" & txtCase(3) & "'" & _
             " and cp10 in (" & ACSforTIPSstep & ") and nvl(cp16,0) > 0 and cp159=0 and cp13=s1.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp14=s2.st01(+) "
    strTmpQ = strTmpQ & " order by cp05,cp09"
    intQ = 1
    Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
    If intQ = 1 Then
         MGrid1.FixedCols = 0
         Set MGrid1.Recordset = rsQuery
         Call SetGrd1
         If strListCP43 <> "" Then
            Call doQuery2
         End If
    Else
         'Modified by Lydia 2025/09/02
         'If bolMsg = True Then MsgBox "查無TIPS進度！", vbInformation
         If bolMsg = True Then MsgBox "查無可以設定請款階段進度！", vbInformation
    End If
End Sub

Private Sub Form_Load()

   txtInput.Visible = False
   txtYear.Visible = False 'Added by Lydia 2025/04/09
   txtAmt.Visible = False 'Added by Lydia 2025/05/19
   MoveFormToCenter Me
   Call ClearForm(True)
   Call SetGrd1(True)
   Call SetGrd2(True)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm081034 = Nothing
End Sub

Private Sub ClearForm(ByVal bolResetCase As Boolean)
Dim oObj
    
   If bolResetCase = True Then
      For Each oObj In txtCase
          If oObj.Index > 0 Then
             oObj.Text = ""
          End If
      Next
   End If
   
   For Each oObj In lblFM2
      oObj.Caption = ""
   Next

   Combo1.Clear
   strNowCP09 = ""
   strNowCP10 = ""
   strNowCP10name = ""
   strListCP43 = ""
   txtInput.Text = ""
   txtInput.Visible = False
   'Added by Lydia 2025/04/09
   txtYear.Text = ""
   txtYear.Visible = False
   passRow = 0
   passCol = 0
   'end 2025/04/09
   'Added by Lydia 2025/05/19
   txtAmt.Text = ""
   txtAmt.Visible = False
   strListCP60_New = ""
   strListCP60_Old = ""
   'end 2025/05/19
End Sub

Private Sub MGrid1_Click()

   If "" & MGrid1.TextMatrix(MGrid1.row, 2) <> "" Then
       GridClick MGrid1, intLastRow, 0, 0, , "V"
       txtInput.Visible = False
       txtYear.Visible = False 'Added by Lydia 2025/04/09
       txtAmt.Visible = False 'Added by Lydia 2025/05/19
   End If
End Sub

Private Sub MGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol2 As Long, nRow2 As Long

   getGrdColRow MGrid1, x, y, nCol2, nRow2
   If nCol2 < 0 Or nRow2 < 0 Then Exit Sub
   MGrid1.col = nCol2
   MGrid1.row = nRow2
   If Me.MGrid1.row < 1 And Me.MGrid1.Text <> "V" Then
        If m_blnCol2OrderAsc = True Then
           Me.MGrid1.Sort = 5 '字串昇冪
           m_blnCol2OrderAsc = False
        Else
           Me.MGrid1.Sort = 6 '字串降冪
           m_blnCol2OrderAsc = True
        End If
   End If
End Sub

Private Sub MGrid2_Scroll()
   txtInput.Visible = False
End Sub

Private Sub txtCase_GotFocus(Index As Integer)
    TextInverse txtCase(Index)
    CmdQuery.Default = True
End Sub

Private Sub txtCase_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCase_LostFocus(Index As Integer)
    If Index > 1 And Trim(txtCase(Index)) = "" Then
        If Index = 2 Then
             txtCase(2) = "0"
        ElseIf Index = 3 Then
             txtCase(3) = "00"
        End If
    End If
    CmdQuery.Default = False
End Sub

Private Sub SetGrd1(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
 
   arrGridHeadText = Array("V", "收文日期", "收文號", "案件性質", "智權人員", "承辦人員", "收文費用", "收文點數", "CP10", "發文日期")
   arrGridHeadWidth = Array(260, 900, 1000, 1500, 900, 900, 900, 900, 0, 900)
        
   MGrid1.Visible = False
   MGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
         MGrid1.Clear
         MGrid1.Rows = 2
   End If
       
   For iRow = 0 To MGrid1.Cols - 1
      MGrid1.row = 0
      MGrid1.col = iRow
      MGrid1.Text = arrGridHeadText(iRow)
      MGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MGrid1.CellAlignment = flexAlignCenterCenter
   Next
   If colCP09_1 = 0 Then
      colCP09_1 = PUB_MGridGetId("收文號", MGrid1)
      colCP10_1 = PUB_MGridGetId("CP10", MGrid1)
      colCP10_1name = PUB_MGridGetId("案件性質", MGrid1)
   End If
   
   MGrid1.Visible = True
End Sub

Private Sub doQuery2()
Dim intCnt As Integer 'Added by Lydia 2025/09/02

   txtInput.Visible = False
   txtYear.Visible = False 'Added by Lydia 2025/04/09
   txtAmt.Visible = False 'Added by Lydia 2025/05/19
   'Added by Lydia 2025/05/19
   strListCP60_New = ""
   strListCP60_Old = ""
   'end 2025/05/19
   
   Call SetGrd2(True) '清空

   If strListCP43 = "" Then Exit Sub

    'TIPS進度之相關收文
   'Modified by Lydia 2024/04/03 改成非取消收文非TIPS收文的所有收文
   'strTmpQ = " select '' as v,c2.cp156 as cntcp156 ,substr(sqldatet(c2.cp06),1,10) cp06t,c2.cp09,m2.cpm03||'-'||m1.cpm03 as cp10n, s1.st02 as cp13n,s2.st02 as cp14n" & _
              " ,substr(c2.cp144,10, length(c2.cp144)-10) as cp144,c2.cp64,c2.cp10,c2.cp156,substr(sqldatet(c2.cp27),1,10) as cp27,c2.cp43" & _
              " from caseprogress c1,casepropertymap m1,caseprogress c2, casepropertymap m2 ,staff s1, staff s2" & _
              " where c1.cp01='" & txtCase(0) & "' and c1.cp02='" & txtCase(1) & "' and c1.cp03='" & txtCase(2) & "' and c1.cp04='" & txtCase(3) & "' and c1.cp159=0 and c1.cp10 in (" & ACSforTIPSstep & ")" & _
              " and c1.cp01=m1.cpm01(+) and c1.cp10=m1.cpm02(+) and c1.cp09=c2.cp43(+) and c2.cp09 is not null and c2.cp159=0" & _
              " and c2.cp01=m2.cpm01(+) and c2.cp10=m2.cpm02(+) and c2.cp13=s1.st01(+) and c2.cp14=s2.st01(+) and substr(c2.cp10,1,1) > '1' "
   '3/26 增加C類來函的B類收文
   'strTmpQ = strTmpQ & " Union select '' as v,c2.cp156 as cntcp156 ,substr(sqldatet(c2.cp06),1,10) cp06t,c2.cp09,m2.cpm03||'-'||m1.cpm03 as cp10n, s1.st02 as cp13n,s2.st02 as cp14n" & _
              " ,substr(c2.cp144,10, length(c2.cp144)-10) as cp144,c2.cp64,c2.cp10,c2.cp156,substr(sqldatet(c2.cp27),1,10) as cp27,c2.cp43" & _
              " from caseprogress c1,casepropertymap m1,caseprogress c2, casepropertymap m2 ,staff s1, staff s2" & _
              " where c1.cp01='" & txtCase(0) & "' and c1.cp02='" & txtCase(1) & "' and c1.cp03='" & txtCase(2) & "' and c1.cp04='" & txtCase(3) & "' and c1.cp159=0 and substr(c1.cp09,1,1)='C' " & _
              " and c1.cp01=m1.cpm01(+) and c1.cp10=m1.cpm02(+) and c1.cp09=c2.cp43(+) and c2.cp09 is not null and c2.cp159=0" & _
              " and c2.cp01=m2.cpm01(+) and c2.cp10=m2.cpm02(+) and c2.cp13=s1.st01(+) and c2.cp14=s2.st01(+) and substr(c2.cp10,1,1) > '1' "
   'Modified by Lydia 2025/04/09 +CP115
   'Modified by Lydia 2025/05/19 +CP144N,CP60N,CP60
   'Modified by Lydia 2025/05/28 將發文日期改在CP144前面
   'Modified by Lydia 2025/09/02
   'strTmpQ = " select '' as v,c1.cp156 as cntcp156,c1.cp115 as cp115n,substr(c1.cp144,10, length(c1.cp144)-10) as cp144n,decode(c1.cp60,null,null,c1.cp60||':'||substr(c1.cp144,10, length(c1.cp144)-10)) as cp60n," & _
             " substr(sqldatet(c1.cp06),1,10) cp06t,c1.cp09,m1.cpm03||decode(c2.cp09,null,null,'-'||m2.cpm03) as cp10n,s1.st02 as cp13n,s2.st02 as cp14n,substr(sqldatet(c1.cp27),1,10) as cp27,substr(c1.cp144,10, length(c1.cp144)-10) as cp144," & _
             " c1.cp64,c1.cp10,c1.cp156,c1.cp43,c2.cp10 as cp43cp10,c1.cp115,decode(c1.cp60,null,null,c1.cp60||':'||substr(c1.cp144,10, length(c1.cp144)-10)) as cp60" & _
             " from caseprogress c1,casepropertymap m1,caseprogress c2, casepropertymap m2 ,staff s1, staff s2" & _
             " where c1.cp01='" & txtCase(0) & "' and c1.cp02='" & txtCase(1) & "' and c1.cp03='" & txtCase(2) & "' and c1.cp04='" & txtCase(3) & "'" & _
             " and c1.cp159=0 and c1.cp09 <'C' and c1.cp10 not in (" & ACSforTIPSstep & ") and c1.cp01=m1.cpm01(+) and c1.cp10=m1.cpm02(+)" & _
             " and c1.cp13=s1.st01(+) and c1.cp14=s2.st01(+) and c1.cp43=c2.cp09(+) and c2.cp01=m2.cpm01(+) and c2.cp10=m2.cpm02(+)"
   'end 2024/04/03
JumpToRe:
   'Modified by Lydia 2025/09/15 因為9/5增加性質，包含已設過請款階段的收文，所以調整語法；and c1.cp10 not in (" & ACSforTIPSstep & ")=> and (c1.cp10 not in (" & ACSforTIPSstep & ") or (c1.cp10 in (" & ACSforTIPSstep & ") and nvl(c1.cp156,0) > 0))
   'Modified by Lydia 2025/10/03 調整顯示請款金額
   'strTmpQ = " select '' as v,c1.cp156 as cntcp156,c1.cp115 as cp115n,substr(c1.cp144,10, length(c1.cp144)-10) as cp144n,decode(c1.cp60,null,null,c1.cp60||':'||substr(c1.cp144,10, length(c1.cp144)-10)) as cp60n," & _
             " substr(sqldatet(c1.cp06),1,10) cp06t,c1.cp09,m1.cpm03||decode(c2.cp09,null,null,'-'||m2.cpm03) as cp10n,s1.st02 as cp13n,s2.st02 as cp14n,substr(sqldatet(c1.cp27),1,10) as cp27,substr(c1.cp144,10, length(c1.cp144)-10) as cp144," & _
             " c1.cp64,c1.cp10,c1.cp156," & IIf(intCnt = 0, "c1.cp43", "c1.cp09 as cp43") & "," & IIf(intCnt = 0, "c2.cp10", "c1.cp10") & " as cp43cp10,c1.cp115,decode(c1.cp60,null,null,c1.cp60||':'||substr(c1.cp144,10, length(c1.cp144)-10)) as cp60" & _
             " from caseprogress c1,casepropertymap m1,caseprogress c2, casepropertymap m2 ,staff s1, staff s2" & _
             " where c1.cp01='" & txtCase(0) & "' and c1.cp02='" & txtCase(1) & "' and c1.cp03='" & txtCase(2) & "' and c1.cp04='" & txtCase(3) & "'" & _
             " and c1.cp159=0 and c1.cp09 <'C' " & IIf(intCnt = 0, "and (c1.cp10 not in (" & ACSforTIPSstep & ") or (c1.cp10 in (" & ACSforTIPSstep & ") and nvl(c1.cp156,0) > 0)) ", "and c1.cp09 in (" & strListCP43 & ")") & " and c1.cp01=m1.cpm01(+) and c1.cp10=m1.cpm02(+)" & _
             " and c1.cp13=s1.st01(+) and c1.cp14=s2.st01(+) and c1.cp43=c2.cp09(+) and c2.cp01=m2.cpm01(+) and c2.cp10=m2.cpm02(+)"
   strTmpQ = " select '' as v,c1.cp156 as cntcp156,c1.cp115 as cp115n,decode(nvl(c1.cp156,0),0,substr(c1.cp144,6, LENGTH(c1.cp144)-6) ,substr(c1.cp144,10, LENGTH(c1.cp144)-10)) as cp144n,decode(c1.cp60,null,null,c1.cp60||':'||decode(nvl(c1.cp156,0),0,substr(c1.cp144,6, LENGTH(c1.cp144)-6) ,substr(c1.cp144,10, LENGTH(c1.cp144)-10))) as cp60n," & _
             " substr(sqldatet(c1.cp06),1,10) cp06t,c1.cp09,m1.cpm03||decode(c2.cp09,null,null,'-'||m2.cpm03) as cp10n,s1.st02 as cp13n,s2.st02 as cp14n,substr(sqldatet(c1.cp27),1,10) as cp27,decode(nvl(c1.cp156,0),0,substr(c1.cp144,6, LENGTH(c1.cp144)-6) ,substr(c1.cp144,10, LENGTH(c1.cp144)-10)) AS cp144," & _
             " c1.cp64,c1.cp10,c1.cp156," & IIf(intCnt = 0, "c1.cp43", "c1.cp09 as cp43") & "," & IIf(intCnt = 0, "c2.cp10", "c1.cp10") & " as cp43cp10,c1.cp115,decode(c1.cp60,null,null,c1.cp60||':'||decode(nvl(c1.cp156,0),0,substr(c1.cp144,6, LENGTH(c1.cp144)-6) ,substr(c1.cp144,10, LENGTH(c1.cp144)-10))) as cp60" & _
             " from caseprogress c1,casepropertymap m1,caseprogress c2, casepropertymap m2 ,staff s1, staff s2" & _
             " where c1.cp01='" & txtCase(0) & "' and c1.cp02='" & txtCase(1) & "' and c1.cp03='" & txtCase(2) & "' and c1.cp04='" & txtCase(3) & "'" & _
             " and c1.cp159=0 and c1.cp09 <'C' " & IIf(intCnt = 0, "and (c1.cp10 not in (" & ACSforTIPSstep & ") or (c1.cp10 in (" & ACSforTIPSstep & ") and nvl(c1.cp156,0) > 0)) ", "and c1.cp09 in (" & strListCP43 & ")") & " and c1.cp01=m1.cpm01(+) and c1.cp10=m1.cpm02(+)" & _
             " and c1.cp13=s1.st01(+) and c1.cp14=s2.st01(+) and c1.cp43=c2.cp09(+) and c2.cp01=m2.cpm01(+) and c2.cp10=m2.cpm02(+)"
   'end 2025/09/02
   'Modified by Lydia 2025/04/09
   'strTmpQ = strTmpQ & " order by cp43, 3,4 "
   strTmpQ = strTmpQ & " order by c1.cp43, c1.cp06, c1.cp09 "
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
   If intQ = 1 Then
      MGrid2.FixedCols = 0
      Set MGrid2.Recordset = rsQuery
      Call SetGrd2
      MGrid2.FixedCols = cntFixed 'Added by Lydia 2025/04/09
      'Added by Lydia 2025/05/19
      strExc(0) = "select cp60 from caseprogress where cp01='" & txtCase(0) & "' and cp02='" & txtCase(1) & "' and cp03='" & txtCase(2) & "' and cp04='" & txtCase(3) & "' " & _
                  "and cp159=0 and nvl(cp156,0)>0 and nvl(cp60,'N') <> 'N' order by 1"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strListCP60_Old = RsTemp.GetString(adClipString, , , ",")
      End If
   'Added by Lydia 2025/09/02
   Else
      'Modified by Lydia 2025/09/05 全部性質都可以，只有一筆收文可以列入請款階段
      'If bolAdd = True And intCnt = 0 Then
      'If bolAdd = True And intCnt = 0 Then
      If intCnt = 0 Then
         intCnt = 1
         GoTo JumpToRe
      End If
   'end 2025/09/02
   End If
End Sub

Private Sub SetGrd2(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   'Modified by Lydia 2024/04/03 +CP43CP10
   'Modified by Lydia 2025/04/09 + 年度,CP115,比例
   'Modified by Lydia 2025/05/19
   'arrGridHeadText = Array("V", "階段", "年度", "本所期限", "收文號", "案件性質", "智權人員", "承辦人員", "請款金額"6, "進度備註", "CP10", "CP156", "發文日期", "CP43", "CP43CP10", "CP115")
   'arrGridHeadWidth = Array(300, 500, 600, 900, 1000, 1400, 900, 900, 1000, 1500, 0, 0, 900, 0, 0, 0)
   'Modified by Lydia 2025/05/28 將發文日期改在CP144前面，不顯示CP144
   'arrGridHeadText = Array("V", "階段", "年度", "請款金額", "CP60N", "本所期限", "收文號", "案件性質", "智權人員", "承辦人員", "實際請款金額", "進度備註", "CP10", "CP156", "發文日期", "CP43", "CP43CP10", "CP115", "CP60")
   'arrGridHeadWidth = Array(300, 500, 600, 1000, 0, 900, 1000, 1400, 900, 900, 1200, 1500, 0, 0, 900, 0, 0, 0, 0)
   arrGridHeadText = Array("V", "階段", "年度", "請款金額", "CP60N", "本所期限", "收文號", "案件性質", "智權人員", "承辦人員", "發文日期", "CP144", "進度備註", "CP10", "CP156", "CP43", "CP43CP10", "CP115", "CP60")
   arrGridHeadWidth = Array(300, 500, 600, 1000, 0, 900, 1000, 1400, 900, 900, 900, 0, 1500, 0, 0, 0, 0, 0, 0)
   
   MGrid2.Visible = False
   MGrid2.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
      MGrid2.Clear
      MGrid2.Rows = 2
   End If
       
   For iRow = 0 To MGrid2.Cols - 1
      MGrid2.row = 0
      MGrid2.col = iRow
      MGrid2.Text = arrGridHeadText(iRow)
      MGrid2.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MGrid2.CellAlignment = flexAlignCenterCenter
   Next

   If colCP09_2 = 0 Then
      colCP09_2 = PUB_MGridGetId("收文號", MGrid2)
      colCP10_2 = PUB_MGridGetId("CP10", MGrid2)
      colCP10_2name = PUB_MGridGetId("案件性質", MGrid2)
      colCP156_New = PUB_MGridGetId("階段", MGrid2)
      colCP156_Old = PUB_MGridGetId("CP156", MGrid2)
      colCP144_New = PUB_MGridGetId("請款金額", MGrid2)
      'Added by Lydia 2024/04/03
      colCP43 = PUB_MGridGetId("CP43", MGrid2)
      colCP43cp10 = PUB_MGridGetId("CP43CP10", MGrid2)
      'Added by Lydia 2025/04/09
      colCP115_New = PUB_MGridGetId("年度", MGrid2)
      colCP115_Old = PUB_MGridGetId("CP115", MGrid2)
      'Added by Lydia 2025/05/19
      colCP144_Old = PUB_MGridGetId("CP144", MGrid2)
      colCP60_New = PUB_MGridGetId("CP60N", MGrid2)
      colCP60_Old = PUB_MGridGetId("CP60", MGrid2)
      colCP27 = PUB_MGridGetId("發文日期", MGrid2)
   End If
   
   For iRow = 1 To MGrid2.Rows - 1
      MGrid2.row = iRow
      MGrid2.col = colCP156_New
      '輸入階段=置中
      MGrid2.CellAlignment = flexAlignCenterCenter
      'Added by Lydia 2025/04/09
      MGrid2.row = iRow
      MGrid2.col = colCP115_New
      MGrid2.CellAlignment = flexAlignCenterCenter
      'Added by Lydia 2025/05/19 請款金額靠右
      MGrid2.row = iRow
      MGrid2.col = colCP144_New
      MGrid2.CellAlignment = flexAlignRightCenter
      'end 2025/05/19
      For intI = 0 To cntFixed - 1
         MGrid2.col = intI
         MGrid2.CellBackColor = &H80000005
      Next intI
      'end 2025/04/09
   Next iRow
   
   MGrid2.Visible = True
End Sub

Private Sub SetBox2()
Dim lngLeft As Long, lngTop As Long, ii As Integer
'參考Promoter\frm090220
   With MGrid2
      If .row > 0 And .col = colCP156_New Then
         If .TextMatrix(.row, colCP09_2) <> "" Then
            txtInput.FontName = .CellFontName
            txtInput.FontSize = .CellFontSize
            txtInput.Alignment = .CellAlignment \ 5
            txtInput.Text = .TextMatrix(.row, .col)
            txtInput.Tag = txtInput.Text
            txtInput.Width = .ColWidth(.col)
            txtInput.Height = .RowHeight(.row)
            nRow2 = .row: nCol2 = .col
            txtInput.Visible = True
            txtInput.SetFocus
            TextInverse txtInput
            lngLeft = .Left + 25
            lngTop = .Top + .RowHeight(0) + 25
            For ii = 0 To .col - 1
               lngLeft = lngLeft + .ColWidth(ii)
            Next
            For ii = .TopRow To .row - 1
               lngTop = lngTop + .RowHeight(ii)
            Next
            txtInput.Left = lngLeft: txtInput.Top = lngTop
         End If
      End If
   End With
End Sub

Private Sub MGrid2_Click()
Dim intRow As Integer, intCol As Integer
Dim bolUpdate As Boolean 'Adeed by Lydia 2025/04/09

   With MGrid2
      If .MouseRow > 0 Then
         'Modified by Lydia 2025/04/09
         'intRow = .MouseRow
         'intCol = .MouseCol
         intRow = IIf(passRow > 0, passRow, .MouseRow)
         intCol = IIf(passCol > 0, passCol, .MouseCol)
         'end 2025/04/09
         .row = intRow
         '----單選
         'Modified by Lydid 2025/04/09
         'GridClick MGrid2, intLastRow2, 0, 0, , "V", , colCP156_New
         GridClick MGrid2, intLastRow2, 0, 0, , "V", , Format(colCP156_New, "00") & "," & Format(colCP115_New, "00")
         intLastRow2 = intRow
         .col = intCol
         'Adeed by Lydia 2025/04/09 已有請款金額欄應限制不可改階段、年度
         'Modified by Lydia 2025/05/19 改用發文日判斷
         'If Pub_GetCP144Val(txtCase(0), txtCase(1), txtCase(2), txtCase(3), "2", MGrid2.TextMatrix(intRow, colCP144_New)) <> "" Then
         If Trim("" & MGrid2.TextMatrix(intRow, colCP27)) <> "" Then
            bolUpdate = False
         Else
            bolUpdate = True
         End If
        
         'Modified by Lydia 2025/04/09 +And bolUpdate = True
         If "" & MGrid2.TextMatrix(intRow, 0) = "V" And intCol = colCP156_New And bolUpdate = True Then
             SetBox2
         Else
             txtInput.Visible = False
         End If
         'Added by Lydia 2025/04/09
         If "" & MGrid2.TextMatrix(intRow, 0) = "V" And intCol = colCP115_New And bolUpdate = True Then
             SetBoxYear
         Else
             txtYear.Visible = False
         End If
         'Added by Lydia 2025/05/19
         If "" & MGrid2.TextMatrix(intRow, 0) = "V" And intCol = colCP144_New And bolUpdate = True Then
             SetBoxAmt
         Else
             txtAmt.Visible = False
         End If
         'end 2025/05/19
         passRow = 0
         passCol = 0
         'end 2025/04/09
       End If
   End With
End Sub

Private Sub txtInput_GotFocus()
   TextInverse txtInput
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)

   KeyAscii = UpperCase(KeyAscii)
   'Modified by Lydia 2025/04/09 +KeyAscii = vbKeyReturn
   If KeyAscii = vbKeyReturn Or KeyAscii = 8 Or (KeyAscii >= 49 And KeyAscii <= 57) Then
      If KeyAscii = vbKeyReturn Then
         MGrid2.TextMatrix(nRow2, nCol2) = txtInput.Text
         GoNext
      ElseIf KeyAscii = vbKeyEscape Then
         txtInput = txtInput.Tag
         TextInverse txtInput
      End If
   Else
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub GoNext()
   With MGrid2
      If .row < .Rows - 1 Then
         .row = .row + 1
      Else
         .row = 1
      End If
      'Modified by Lydia 2025/04/09 按Enter往下移
      'SetBox2
      If .row > 1 Then
         passRow = .row
         passCol = colCP156_New
         Call MGrid2_Click
      End If
      'end 2025/04/09
   End With
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
   MGrid2.TextMatrix(nRow2, nCol2) = txtInput.Text
End Sub

'Added by Lydia 2025/04/09
Private Sub txtYear_GotFocus()
   TextInverse txtYear
End Sub

'Added by Lydia 2025/04/09
Private Sub txtYear_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii = vbKeyReturn Or KeyAscii = 8 Or (KeyAscii >= 49 And KeyAscii <= 57) Then
      If KeyAscii = vbKeyReturn Then
         MGrid2.TextMatrix(nRow2, nCol2) = txtYear.Text
         GoNextYear
      ElseIf KeyAscii = vbKeyEscape Then
         txtYear = txtYear.Tag
         TextInverse txtYear
      End If
   Else
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Lydia 2025/04/09
Private Sub txtYear_Validate(Cancel As Boolean)
   MGrid2.TextMatrix(nRow2, nCol2) = txtYear.Text
End Sub

'Added by Lydia 2025/04/09
Private Sub SetBoxYear()
Dim lngLeft As Long, lngTop As Long, ii As Integer
'參考Promoter\frm090220
   With MGrid2
      If .row > 0 And .col = colCP115_New Then
         If .TextMatrix(.row, colCP09_2) <> "" Then
            txtYear.FontName = .CellFontName
            txtYear.FontSize = .CellFontSize
            txtYear.Alignment = .CellAlignment \ 5
            txtYear.Text = .TextMatrix(.row, .col)
            txtYear.Tag = txtYear.Text
            txtYear.Width = .ColWidth(.col)
            txtYear.Height = .RowHeight(.row)
            nRow2 = .row: nCol2 = .col
            txtYear.Visible = True
            txtYear.SetFocus
            TextInverse txtYear
            lngLeft = .Left + 25
            lngTop = .Top + .RowHeight(0) + 25
            For ii = 0 To .col - 1
               lngLeft = lngLeft + .ColWidth(ii)
            Next
            For ii = .TopRow To .row - 1
               lngTop = lngTop + .RowHeight(ii)
            Next
            txtYear.Left = lngLeft: txtYear.Top = lngTop
         End If
      End If
   End With
End Sub

'Added by Lydia 2025/04/09
Private Sub GoNextYear()
   With MGrid2
      If .row < .Rows - 1 Then
         .row = .row + 1
      Else
         .row = 1
      End If
      '按Enter往下移
      If .row > 1 Then
         passRow = .row
         passCol = colCP115_New
         Call MGrid2_Click
      End If
   End With
End Sub

'Added by Lydia 2025/05/19
Private Sub txtAmt_GotFocus()
   TextInverse txtAmt
End Sub

'Added by Lydia 2025/05/19
Private Sub txtAmt_KeyPress(KeyAscii As Integer)

Dim tmpBol As Boolean
   KeyAscii = Pub_NumAscii(KeyAscii)
   If KeyAscii = vbKeyReturn Then
      txtAmt_Validate (tmpBol)
      If tmpBol = False Then
         GoNextAmt
      End If
   ElseIf KeyAscii = vbKeyEscape Then
      txtAmt = txtAmt.Tag
      TextInverse txtAmt
   End If
End Sub

'Added by Lydia 2025/05/19
Private Sub txtAmt_Validate(Cancel As Boolean)
   
   If txtAmt = "" Then
      If "" & MGrid2.TextMatrix(nRow2, colCP60_New) <> "" Then strListCP60_New = Replace(strListCP60_New, Left("" & MGrid2.TextMatrix(nRow2, colCP60_New), 9) & ",", "")
      If "" & MGrid2.TextMatrix(nRow2, colCP60_Old) <> "" Then strListCP60_Old = Replace(strListCP60_Old, Left("" & MGrid2.TextMatrix(nRow2, colCP60_Old), 9) & ",", "")
      
      MGrid2.TextMatrix(nRow2, colCP60_New) = ""
   Else
      If txtAmt <> "" & MGrid2.TextMatrix(nRow2, colCP144_Old) Then
         If Trim("" & MGrid2.TextMatrix(nRow2, colCP60_New)) <> "" And InStr("" & MGrid2.TextMatrix(nRow2, colCP60_New), ":" & txtAmt) > 0 Then
            '不修改
         Else
            If Trim("" & MGrid2.TextMatrix(nRow2, colCP60_New)) <> "" Then
               strListCP60_New = Replace(strListCP60_New, Left("" & MGrid2.TextMatrix(nRow2, colCP60_New), 9) & ",", "")
            End If
            strTmpQ = Pub_ACS_TIPS_GetCp60("2", txtCase(0), txtCase(1), txtCase(2), txtCase(3), txtAmt, strListCP60_Old, strListCP60_New)
            If strTmpQ = "" Then
               MsgBox "請款金額與收據金額不符，或已有其他收文已設定相同請款金額！", vbCritical
               Cancel = True
               Exit Sub
            End If
            '1234
            strListCP60_New = strListCP60_New & strTmpQ & ","
            MGrid2.TextMatrix(nRow2, colCP60_New) = strTmpQ & ":" & txtAmt
         End If
      Else
         If "" & MGrid2.TextMatrix(nRow2, colCP60_New) = "" Then
            strListCP60_New = Replace(strListCP60_New, Left("" & MGrid2.TextMatrix(nRow2, colCP60_New), 9) & ",", "")
         End If
         MGrid2.TextMatrix(nRow2, colCP60_New) = MGrid2.TextMatrix(nRow2, colCP60_Old)
      End If
   End If

   MGrid2.TextMatrix(nRow2, nCol2) = txtAmt.Text
End Sub

'Added by Lydia 2025/05/19
Private Sub SetBoxAmt()
Dim lngLeft As Long, lngTop As Long, ii As Integer
'參考Promoter\frm090220
   With MGrid2
      If .row > 0 And .col = colCP144_New Then
         If .TextMatrix(.row, colCP09_2) <> "" Then
            txtAmt.FontName = .CellFontName
            txtAmt.FontSize = .CellFontSize
            txtAmt.Alignment = .CellAlignment \ 5
            txtAmt.Text = .TextMatrix(.row, .col)
            txtAmt.Tag = txtAmt.Text
            txtAmt.Width = .ColWidth(.col)
            txtAmt.Height = .RowHeight(.row)
            nRow2 = .row: nCol2 = .col
            txtAmt.Visible = True
            txtAmt.SetFocus
            TextInverse txtAmt
            lngLeft = .Left + 25
            lngTop = .Top + .RowHeight(0) + 25
            For ii = 0 To .col - 1
               lngLeft = lngLeft + .ColWidth(ii)
            Next
            For ii = .TopRow To .row - 1
               lngTop = lngTop + .RowHeight(ii)
            Next
            txtAmt.Left = lngLeft: txtAmt.Top = lngTop
         End If
      End If
   End With
End Sub

'Added by Lydia 2025/05/19
Private Sub GoNextAmt()
   With MGrid2
      If .row < .Rows - 1 Then
         .row = .row + 1
      Else
         .row = 1
      End If
      '按Enter往下移
      If .row > 1 Then
         passRow = .row
         passCol = colCP144_New
         Call MGrid2_Click
      End If
   End With
End Sub

