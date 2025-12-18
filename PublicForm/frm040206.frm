VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040206 
   BorderStyle     =   1  '單線固定
   Caption         =   "CF 結餘資料維護"
   ClientHeight    =   5500
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5500
   ScaleWidth      =   9120
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   3330
      MaxLength       =   2
      TabIndex        =   3
      Top             =   435
      Width           =   345
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2955
      MaxLength       =   1
      TabIndex        =   2
      Top             =   435
      Width           =   225
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1185
      MaxLength       =   3
      TabIndex        =   0
      Top             =   435
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1815
      MaxLength       =   5
      TabIndex        =   4
      Top             =   435
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2565
      MaxLength       =   1
      TabIndex        =   5
      Top             =   435
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2925
      MaxLength       =   1
      TabIndex        =   6
      Top             =   435
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   3330
      MaxLength       =   2
      TabIndex        =   7
      Top             =   435
      Width           =   315
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束"
      Height          =   375
      Index           =   3
      Left            =   8280
      TabIndex        =   11
      Top             =   30
      Width           =   780
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "調整結餘日期"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   6810
      TabIndex        =   10
      Top             =   30
      Width           =   1470
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   3230
      Left            =   50
      TabIndex        =   13
      Top             =   1760
      Width           =   9030
      _ExtentX        =   15946
      _ExtentY        =   5697
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
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
      _Band(0).Cols   =   1
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "取消結餘日期"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   5340
      TabIndex        =   9
      Top             =   30
      Width           =   1470
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4530
      TabIndex        =   8
      Top             =   30
      Width           =   810
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   1
      Top             =   435
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "說明：有可結餘刪除日時表示：可結餘日之後有新收文 或是 此案已完成結餘計算。調整後自動刪除"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   200
      Left            =   120
      TabIndex        =   25
      Top             =   5160
      Width           =   8700
   End
   Begin VB.Label Label2 
      Caption         =   "PS：未開收據進度不更新可結餘日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   200
      Left            =   5730
      TabIndex        =   24
      Top             =   480
      Width           =   3300
   End
   Begin MSForms.Label lbl2 
      Height          =   180
      Left            =   1440
      TabIndex        =   23
      Top             =   750
      Width           =   7620
      VariousPropertyBits=   27
      Size            =   "13441;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1020
      TabIndex        =   12
      Top             =   1395
      Width           =   8040
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "14182;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   45
      TabIndex        =   22
      Top             =   480
      Width           =   975
   End
   Begin VB.Line Line4 
      X1              =   1575
      X2              =   3435
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   45
      TabIndex        =   21
      Top             =   735
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "中："
      Height          =   180
      Index           =   2
      Left            =   1020
      TabIndex        =   20
      Top             =   735
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "英："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   1020
      TabIndex        =   19
      Top             =   945
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "日："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   1020
      TabIndex        =   18
      Top             =   1170
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   45
      TabIndex        =   17
      Top             =   1455
      Width           =   780
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   1
      Left            =   1440
      TabIndex        =   16
      Top             =   945
      Width           =   7620
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   2
      Left            =   1440
      TabIndex        =   15
      Top             =   1170
      Width           =   7620
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3765
      TabIndex        =   14
      Top             =   495
      Width           =   990
   End
End
Attribute VB_Name = "frm040206"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/25 Form2.0已修改 lbl2(原:lbl1(0))/grd1
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
'Create By Nick 2005/07/22
Option Explicit

Dim oStrNa01 As String
Dim strSQL1  As String
Dim strSQL2 As String
Dim StrSQL3 As String
Dim StrSQL4 As String, s As String
Dim strMaxCP05 As String 'Added by Lydia 2025/06/11 該案有收費進度之最大收文日

Private Sub cmdok_Click(Index As Integer)
On Error GoTo ErrArea
Dim i As Integer
Dim oStrNewDate As String
Dim DateIsOk As Boolean
Dim D_Cancel As Boolean
   Select Case Index
   Case 0
         If Len(txt1(0)) = 0 Then
             s = MsgBox("本所案號不可空白", , "USER 輸入錯誤")
             If Len(txt1(0)) = 0 Then txt1(0).SetFocus: Exit Sub
         Else
            If txt1(0) = "TF" Then
               If Len(txt1(4)) = 0 Or Len(txt1(5)) = 0 Then
                  s = MsgBox("本所案號不可空白", , "USER 輸入錯誤")
                  txt1(0).SetFocus
                  txt1_GotFocus (0)
                  Exit Sub
               End If
            Else
               If Len(txt1(1)) = 0 Then
                  s = MsgBox("本所案號不可空白", , "USER 輸入錯誤")
                  txt1(0).SetFocus
                  txt1_GotFocus (0)
                  Exit Sub
               End If
            End If
         End If
         D_Cancel = False
         txt1_Validate 0, D_Cancel
         If D_Cancel = False Then
            Me.Enabled = False
            Screen.MousePointer = vbHourglass
            StrMenu
            Screen.MousePointer = vbDefault
            Me.Enabled = True
         End If
   Case 1
         Me.Enabled = False
         Screen.MousePointer = vbHourglass
         cnnConnection.BeginTrans
         For i = 0 To GRD1.Rows - 1
            strSql = "update caseprogress set cp109=null where cp09='" & GRD1.TextMatrix(i, 0) & "' "
            cnnConnection.Execute strSql
            '2011/11/7 add by sonia 子案也要更新
            If txt1(0) = "TF" Or txt1(0) = "CFP" Then
               If txt1(0) = "TF" Then
                  strSql = "update caseprogress set cp109=null where cp09 in (select cp09 from caseprogress where cp43='" & GRD1.TextMatrix(i, 0) & "' AND CP01='" & strSQL1 & "' AND CP02='" & strSQL2 & "' AND CP03<>'" & StrSQL3 & "' AND CP04<>'" & StrSQL4 & "') "
               Else
                  strSql = "update caseprogress set cp109=null where cp09 in (select cp09 from caseprogress where cp43='" & GRD1.TextMatrix(i, 0) & "' AND CP01='" & strSQL1 & "' AND CP02='" & strSQL2 & "' AND CP03='" & StrSQL3 & "' AND CP04<>'" & StrSQL4 & "') "
               End If
               cnnConnection.Execute strSql
            End If
            '2011/11/7 end
         Next i
         cnnConnection.CommitTrans
         StrMenu
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         'add by sonia 2025/4/24
         txt1_LostFocus (0)
         cmdok(0).Default = True
         'end 2025/4/24
   Case 2
         DateIsOk = False
         Do While DateIsOk = False
            oStrNewDate = InputBox("請輸入調整後的民國日期(不含/)！", "輸入日期")
            If Len(Trim(ChangeTStringToWString(oStrNewDate))) = 8 Then
               If ChkWorkDay(ChangeTStringToWString(oStrNewDate)) = False Then
                  MsgBox "請輸入工作日！", vbExclamation, "日期錯誤！"
               Else
                  'Added by Lydia 2025/06/11 不可<=該案有收費進度之最大收文日; ex.P-114643
                  If strMaxCP05 <> "" And oStrNewDate <= strMaxCP05 Then
                     MsgBox "不可小於或等於有收費進度之最大收文日" & strMaxCP05, vbExclamation, "日期錯誤！"
                  Else
                  'end 2025/06/11
                     DateIsOk = True
                  End If
               End If

            ElseIf Trim(oStrNewDate) = "" Then
               Exit Sub
            End If
         Loop
         Me.Enabled = False
         Screen.MousePointer = vbHourglass
         cnnConnection.BeginTrans
         For i = 0 To GRD1.Rows - 1
            'modify by sonia 2022/10/13 有收費但未開收據資料不可更新，否則換出名公司時會有錯
            'strSql = "update caseprogress set cp109=" & ChangeTStringToWString(oStrNewDate) & " where cp09='" & GRD1.TextMatrix(i, 0) & "' "
            'modify by sonia 2023/3/10 同時清除CP146
            strSql = "update caseprogress set cp109=" & ChangeTStringToWString(oStrNewDate) & ",cp146=null where cp09='" & GRD1.TextMatrix(i, 0) & "' AND (nvl(cp16,0)=0 or cp60 is not null) "
            cnnConnection.Execute strSql
            '2011/11/7 add by sonia 子案也要更新
            If txt1(0) = "TF" Or txt1(0) = "CFP" Then
               If txt1(0) = "TF" Then
                  'modify by sonia 2023/3/10 同時清除CP146
                  strSql = "update caseprogress set cp109=" & ChangeTStringToWString(oStrNewDate) & ",cp146=null where cp09 in (select cp09 from caseprogress where cp43='" & GRD1.TextMatrix(i, 0) & "' AND CP01='" & strSQL1 & "' AND CP02='" & strSQL2 & "' AND CP03<>'" & StrSQL3 & "' AND CP04<>'" & StrSQL4 & "') "
               Else
                  'modify by sonia 2023/3/10 同時清除CP146
                  strSql = "update caseprogress set cp109=" & ChangeTStringToWString(oStrNewDate) & ",cp146=null where cp09 in (select cp09 from caseprogress where cp43='" & GRD1.TextMatrix(i, 0) & "' AND CP01='" & strSQL1 & "' AND CP02='" & strSQL2 & "' AND CP03='" & StrSQL3 & "' AND CP04<>'" & StrSQL4 & "') "
               End If
               cnnConnection.Execute strSql
            End If
            '2011/11/7 end
         Next i
         cnnConnection.CommitTrans
         StrMenu
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         'add by sonia 2025/4/24
         txt1_LostFocus (0)
         cmdok(0).Default = True
         'end 2025/4/24
   Case 3
        Unload Me
   Case Else
   End Select
   Exit Sub
ErrArea:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbInformation
End Sub

'查詢
Sub StrMenu()
Dim StrFa As String
Dim i As Integer
   '本所案號
   strSQL1 = ""
   strSQL2 = ""
   StrSQL3 = ""
   StrSQL4 = ""
   
   lbl2.Caption = "" 'Modify by Amy 原:Lbl1(0) = "",改Form2.0
   Lbl1(1) = ""
   Lbl1(2) = ""
   Me.lblClose.Caption = ""
   Combo1.Clear
   GRD1.Rows = 2
   GRD1.Clear
   GRD1.Refresh
   SetGridWidth
   
   If txt1(0) = "TF" Then
      strSQL1 = txt1(0)
      strSQL2 = txt1(4) & txt1(5)
      If Len(Trim(txt1(6))) <> 1 Then
         StrSQL3 = "0"
      Else
         StrSQL3 = txt1(6)
      End If
      If Len(Trim(txt1(7))) <> 2 Then
         StrSQL4 = "00"
      Else
         StrSQL4 = txt1(7)
      End If
   Else
      strSQL1 = txt1(0)
      strSQL2 = txt1(1)
      If Len(Trim(txt1(2))) <> 1 Then
         StrSQL3 = "0"
      Else
         StrSQL3 = txt1(2)
      End If
      If Len(Trim(txt1(3))) <> 2 Then
         StrSQL4 = "00"
      Else
         StrSQL4 = txt1(3)
      End If
   End If
                       strSql = "select pa05,pa06,pa07,'中 '||cu04,'英 '||cu05||' '||cu88||' '||cu89||' '||cu90,'日 '||cu06,'錯誤!! '||PA26,PA57,pa09 from patent,customer where pa01='" & strSQL1 & "' and pa02='" & strSQL2 & "' and pa03='" & StrSQL3 & "' and pa04='" & StrSQL4 & "' and pa09<>'000' and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),'','0',substr(pa26,9,1))=cu02(+) "
   strSql = strSql & " union all select sp05,sp06,sp07,'中 '||cu04,'英 '||cu05||' '||cu88||' '||cu89||' '||cu90,'日 '||cu06,'錯誤!! '||Sp08,SP15,sp09 FROM SERVICEPRACTICE,CUSTOMER WHERE SP01='" & strSQL1 & "' AND SP02='" & strSQL2 & "' AND SP03='" & StrSQL3 & "' AND SP04='" & StrSQL4 & "' AND SP09<>'000' AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) "
   strSql = strSql & " union all select TM05,TM06,TM07,'中 '||cu04,'英 '||cu05||' '||cu88||' '||cu89||' '||cu90,'日 '||cu06,'錯誤!! '||TM23,TM29,tm10 FROM TRADEMARK,CUSTOMER WHERE TM01='" & strSQL1 & "' AND TM02='" & strSQL2 & "' AND TM03='" & StrSQL3 & "' AND TM04='" & StrSQL4 & "' AND TM10<>'000' AND SUBSTR(TM23,1,8)=CU01(+) AND DECODE(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) "
   strSql = strSql & " union all select LC05,LC06,LC07,'中 '||cu04,'英 '||cu05||' '||cu88||' '||cu89||' '||cu90,'日 '||cu06,'錯誤!! '||LC11,LC08,lc15 FROM LAWCASE,CUSTOMER WHERE LC01='" & strSQL1 & "' AND LC02='" & strSQL2 & "' AND LC03='" & StrSQL3 & "' AND LC04='" & StrSQL4 & "' AND LC15<>'000' AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),'','0',SUBSTR(LC11,9,1))=CU02(+) "
   
   CheckOC
   Combo1.Clear
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         lbl2 = CheckStr(.Fields(0)) 'Modify by Amy 原:Lbl1(0) = "",改Form2.0
         Lbl1(1) = CheckStr(.Fields(1))
         Lbl1(2) = CheckStr(.Fields(2))
         If IsNull(.Fields(7).Value) Then
            Me.lblClose.Caption = ""
         Else
            Me.lblClose.Caption = "已閉卷"
         End If
         
         If Len(CheckStr(.Fields(3))) = 0 And Len(CheckStr(.Fields(4))) = 0 And Len(CheckStr(.Fields(5))) = 0 Then
            Combo1.AddItem CheckStr(.Fields(6)), 0
         Else
            If Len(CheckStr(.Fields(3))) = 0 Then
               Combo1.AddItem CheckStr(.Fields(6)), 0
            Else
               Combo1.AddItem CheckStr(.Fields(3)), 0
            End If
            If Len(CheckStr(.Fields(4))) = 0 Then
               Combo1.AddItem CheckStr(.Fields(6)), 1
            Else
               Combo1.AddItem CheckStr(.Fields(4)), 1
            End If
            If Len(CheckStr(.Fields(5))) = 0 Then
               Combo1.AddItem CheckStr(.Fields(6)), 2
            Else
               Combo1.AddItem CheckStr(.Fields(5)), 2
            End If
         End If
         Combo1.Text = Combo1.List(0)
         oStrNa01 = CheckStr(.Fields(7))
      Else
         ShowNoData
         Exit Sub
      End If
      CheckOC
   End With
   '大陸為英中日，其餘為中英日
   If oStrNa01 = "020" Then
      StrFa = "DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)"
   Else
      StrFa = "NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65))"
   End If
   '2011/5/31 add by sonia區分CFP,TF
   'modify by sonia 2022/10/13 加有收費未開收據顯示收據欄為N(decode(nvl(cp16,0),0,null,decode(cp60,null,'N',null)))
   'modify by sonia 2024/12/9 +SQLDateT(cp146)可結餘刪除日期
   If txt1(0) = "TF" Then
      strSql = "select cp09,cpm04,SQLDateT(cp05),SQLDateT(cp27),SQLDateT(cp109),SQLDateT(cp146)," & StrFa & ",decode(nvl(cp16,0),0,null,decode(cp60,null,'N',null)) from caseprogress,casepropertymap,fagent " & _
                " where cp01='" & strSQL1 & "' and cp02='" & strSQL2 & "' and (cp03||cp04='000' or (cp03||cp04<>'000' and substr(cp09,1,1)<>'B')) " & _
                " and cp01=cpm01(+) and cp10=cpm02(+) and substr(cp44,1,8)=fa01(+) and substr(cp44,9,1)=fa02(+) and cp59 is null order by cp05 "
   ElseIf txt1(0) = "CFP" Then
      'modify by sonia 2024/12/26 CFP子案之閉卷或不續辦不顯示故加And Cp01||Cp10 Not In ('CFP907','CFP913')
      strSql = "select cp09,cpm04,SQLDateT(cp05),SQLDateT(cp27),SQLDateT(cp109),SQLDateT(cp146)," & StrFa & ",decode(nvl(cp16,0),0,null,decode(cp60,null,'N',null)) from caseprogress,casepropertymap,fagent " & _
                " where cp01='" & strSQL1 & "' and cp02='" & strSQL2 & "' and cp03='" & StrSQL3 & "' and (cp64 is null or instr(cp64,'子案發文記錄')=0) And Cp01||Cp10 Not In ('CFP907','CFP913') " & _
                " and cp01=cpm01(+) and cp10=cpm02(+) and substr(cp44,1,8)=fa01(+) and substr(cp44,9,1)=fa02(+) and cp59 is null order by cp05 "
   Else
   '2011/5/31 END
      strSql = "select cp09,cpm04,SQLDateT(cp05),SQLDateT(cp27),SQLDateT(cp109),SQLDateT(cp146)," & StrFa & ",decode(nvl(cp16,0),0,null,decode(cp60,null,'N',null)) from caseprogress,casepropertymap,fagent " & _
                " where cp01='" & strSQL1 & "' and cp02='" & strSQL2 & "' and cp03='" & StrSQL3 & "' and cp04='" & StrSQL4 & "' " & _
                " and cp01=cpm01(+) and cp10=cpm02(+) and substr(cp44,1,8)=fa01(+) and substr(cp44,9,1)=fa02(+) and cp59 is null order by cp05 "
   End If   '2011/5/31 ADD BY SONIA
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         Set GRD1.Recordset = AdoRecordSet3
         SetGridWidth
         cmdok(1).Enabled = False
         'ADD BY SONIA 2016/8/31
         For i = 1 To GRD1.Rows - 1
            GRD1.TextMatrix(i, 1) = GRD1.TextMatrix(i, 1) & PUB_GetRelateCasePropertyName(GRD1.TextMatrix(i, 0), "1")
         Next i
         'END 2016/8/31
         For i = 1 To GRD1.Rows - 1
            If GRD1.TextMatrix(i, 4) <> "" Then
               'Modify by Amy 2023/7/20 +if 財務處人員不可按「取消結餘日期」鈕-婉莘
               If Pub_StrUserSt03 <> "M31" Then cmdok(1).Enabled = True
               Exit For
            End If
         Next i
         'Modify by Amy 2023/7/20 +if 財務處人員不可按「調整結餘日期」鈕-婉莘
         If Pub_StrUserSt03 <> "M31" Then cmdok(2).Enabled = True
         'add by sonia 2025/4/24
         cmdok(2).Default = True
         'end 2025/4/24
      Else
         ShowNoData
         cmdok(1).Enabled = False
         cmdok(2).Enabled = False
      End If
   End With
   CheckOC3
   
   'Added by Lydia 2025/06/11
   strMaxCP05 = ""
   strSql = "select max(cp05) mdate from caseprogress where cp159=0 and nvl(cp16,0)>0 and cp01='" & strSQL1 & "' and cp02='" & strSQL2 & "' and cp03='" & StrSQL3 & "' and cp04='" & StrSQL4 & "' "
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         strMaxCP05 = TransDate("" & AdoRecordSet3.Fields("mdate"), 1)
      End If
   End With
   CheckOC3
End Sub

Sub SetGridWidth()
   'Modify By Cheng 2002/02/15
   '在單據編號前加本所案號, 案件性質
   
   With GRD1
       .Cols = 8
       .row = 0
       .col = 0
       .ColWidth(0) = 1300
       .Text = "收文號"
       .col = 1
       .ColWidth(1) = 1800
       .Text = "案件性質"
       .col = 2
       .ColWidth(2) = 800
       .Text = "收文日"
       .col = 3
       .ColWidth(3) = 800
       .Text = "發文日"
       .col = 4
       .ColWidth(4) = 800
       .Text = "可結餘日"
       .col = 5
       .ColWidth(5) = 1200
       .Text = "可結餘刪除日"
       .col = 6
       .ColWidth(6) = 3000
       .Text = "代理人"
       'add by sonia 2022/10/13
       .col = 7
       .ColWidth(7) = 400
       .Text = "收據"
       'end 2022/10/13
   End With
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   lblClose.Caption = ""
   'add by sonia 2023/7/26
   If Pub_StrUserSt03 <> "M31" Then
      Me.Caption = "CF 結餘資料維護"
   Else
      'Modify by Amy 2023/08/14 +(尚未產生結餘單號)字樣-瑞婷
      Me.Caption = "CF 可結餘資料查詢 (尚未產生結餘單號)"
   End If
   'end 2023/7/26
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040206 = Nothing
End Sub


Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   cmdok(0).Default = True   'add by sonia 2025/4/24
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
      Case 0
         If txt1(0) = "TF" Then
            txt1(1).Visible = False
            txt1(2).Visible = False
            txt1(3).Visible = False
            txt1(4).Visible = True
            txt1(5).Visible = True
            txt1(6).Visible = True
            txt1(7).Visible = True
            txt1(4).SetFocus
            '2011/5/31 add by sonia
            txt1(6).Enabled = False
            txt1(7).Enabled = False
            '2011/5/31 end
         Else
            txt1(1).Visible = True
            txt1(2).Visible = True
            txt1(3).Visible = True
            txt1(4).Visible = False
            txt1(5).Visible = False
            txt1(6).Visible = False
            txt1(7).Visible = False
            txt1(1).SetFocus
            '2011/5/31 add by sonia
            If txt1(0) = "CFP" Then
               txt1(3).Enabled = False
            Else
               txt1(3).Enabled = True
            End If
            '2011/5/31 end
         End If
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Dim strTemp As Variant, i As Integer

   Select Case Index
   Case 0

     strTemp = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     For i = 0 To UBound(strTemp)
        If strTemp(i) = txt1(0) Then
            Exit Sub
        End If
     Next i
     s = MsgBox(strUserName & " 沒有 " & txt1(0) & " 的權限 ", , "USER 輸入錯誤")
     Cancel = True
   Case Else
   End Select
End Sub
