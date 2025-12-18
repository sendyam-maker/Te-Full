VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060331 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "外專工程師OA報告處理日數統計"
   ClientHeight    =   3390
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   5685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5685
   Begin VB.CommandButton CmdPrt1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel(&E)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   480
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   2760
      Width           =   4692
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   4215
      ForeColor       =   16711680
      VariousPropertyBits=   8388627
      Caption         =   "P.S. OA報告包含FCP及FMP之審查意見及核駁"
      Size            =   "7435;397"
      BorderColor     =   -2147483644
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4683;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2C 
      Height          =   225
      Index           =   4
      Left            =   4440
      TabIndex        =   17
      Top             =   2160
      Width           =   1300
      VariousPropertyBits=   8388627
      Caption         =   "組別說明："
      Size            =   "2293;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2C 
      Height          =   225
      Index           =   3
      Left            =   4440
      TabIndex        =   16
      Top             =   1920
      Width           =   1300
      VariousPropertyBits=   8388627
      Caption         =   "組別說明："
      Size            =   "2293;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2C 
      Height          =   225
      Index           =   2
      Left            =   4440
      TabIndex        =   15
      Top             =   1680
      Width           =   1300
      VariousPropertyBits=   8388627
      Caption         =   "組別說明："
      Size            =   "2293;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2C 
      Height          =   225
      Index           =   1
      Left            =   4440
      TabIndex        =   14
      Top             =   1440
      Width           =   1300
      VariousPropertyBits=   8388627
      Caption         =   "組別說明："
      Size            =   "2293;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   15
      Left            =   2040
      TabIndex        =   13
      Top             =   1185
      Width           =   255
      VariousPropertyBits=   8388627
      Caption         =   "~"
      Size            =   "450;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   13
      Left            =   2760
      TabIndex        =   12
      Top             =   308
      Width           =   255
      VariousPropertyBits=   8388627
      Caption         =   "~"
      Size            =   "450;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   9
      Left            =   4200
      TabIndex        =   11
      Top             =   1125
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "組別說明："
      Size            =   "2143;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "工程師："
      Size            =   "2143;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1185
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "組　　別："
      Size            =   "2143;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   704
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "系統類別："
      Size            =   "2143;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   4
      Left            =   2400
      TabIndex        =   4
      Top             =   1110
      Width           =   495
      VariousPropertyBits=   679495707
      MaxLength       =   1
      Size            =   "882;635"
      Value           =   "4"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Top             =   1110
      Width           =   500
      VariousPropertyBits=   679495707
      MaxLength       =   1
      Size            =   "882;635"
      Value           =   "1"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Top             =   225
      Width           =   1200
      VariousPropertyBits=   679495707
      MaxLength       =   7
      Size            =   "2117;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   225
      Width           =   1200
      VariousPropertyBits=   679495707
      MaxLength       =   7
      Size            =   "2117;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   645
      Width           =   2895
      VariousPropertyBits=   679495707
      Size            =   "5106;635"
      Value           =   "FCP,P"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   308
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "來函日期："
      Size            =   "2143;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
End
Attribute VB_Name = "frm060331"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/11 Form2.0已檢查 (無需修改的物件)
'Create by Lydia 2018/12/03 外專工程師OA報告處理日數統計
'Memo by Lydia 2018/12/03 使用Form 2.0 (Label、ComboBox和TextBox)
Option Explicit

Dim rsAD As New ADODB.Recordset
Dim strMid As String

Dim oText As MSForms.TextBox

Private Sub CmdPrt1_Click()
Dim stCon As String
Dim intP As Integer

   If FormCheck = False Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   CmdPrt1.Enabled = False
   
   ClearQueryLog (Me.Name) 'Added By Lydia 2021/11/16 清除查詢印表記錄檔欄位
   
   '來函日期
   If txtFM2(0) <> "" Then
       stCon = stCon & " AND CP05>=" & TransDate(txtFM2(0), 2)
   End If
   If txtFM2(1) <> "" Then
       stCon = stCon & " AND CP05<=" & TransDate(txtFM2(1), 2)
   End If
   pub_QL05 = pub_QL05 & ";" & LblFM2(0) & txtFM2(0) & "-" & txtFM2(1)  'Added by Lydia 2021/11/16
   
   '系統類別
   If txtFM2(2) <> "ALL" And txtFM2(2) <> "" Then
      stCon = stCon & " AND CP01 IN (" & GetAddStr(txtFM2(2)) & ")"
   End If
   pub_QL05 = pub_QL05 & ";" & LblFM2(1) & txtFM2(2) 'Added by Lydia 2021/11/16

   '組別
   If txtFM2(3) <> "" Then
      stCon = stCon & " AND ST16 >= '" & txtFM2(3) & "'"
   End If
   If txtFM2(4) <> MsgText(601) Then
      stCon = stCon & " AND ST16 <= '" & txtFM2(4) & "'"
   End If
   pub_QL05 = pub_QL05 & ";" & LblFM2(3) & txtFM2(3) & "-" & txtFM2(4)  'Added by Lydia 2021/11/16
   
   '工程師
   If Combo1.Text <> "" Then
      stCon = stCon & " AND CP14= '" & Trim(Left(Combo1.Text, 6)) & "'"
      pub_QL05 = pub_QL05 & ";" & LblFM2(6) & Combo1.Text  'Added by Lydia 2021/11/16
   End If
   'Memo by Lydia 2019/09/24 特別抓資料: 限Y編號和X編號
   'stCon = stCon & " and pa75 like 'Y34263%' and instr(pa26||','||pa27||','||pa28||','||pa29||','||pa30,'X60873') > 0 "
   'Memo by Lydia 2019/09/24 特別抓資料 +PA77
   strSql = " select st16,cp14,st02  as cp14n, cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseno,decode(pa09,'000',cpm03,cpm04) cp10n,sqldatet(cp05) cp05,sqldatet(cp27) cp27," & _
               " TO_DATE(NVL(cp27,to_char(sysdate,'YYYYMMDD')),'YYYYMMDD')-TO_DATE(cp05,'YYYYMMDD') cdate" & _
               " From caseprogress, staff, patent, casepropertymap" & _
               " where cp10 in ('1002','1202') and cp14=st01(+) and st03='F21' and cp01=cpm01(+) and cp10=cpm02(+) " & stCon & _
               " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp27<>19221111 and cp27>0"
   strMid = strSql '保留
   strSql = strSql & " order by st16,cp14,cp05,cp01,cp02"
    
   If rsAD.State = adStateOpen Then
      rsAD.Close
   End If
   rsAD.CursorLocation = adUseClient

   rsAD.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsAD.RecordCount <> 0 Then
        InsertQueryLog (rsAD.RecordCount) 'Added by Lydia 2021/11/16
        Call ProcExcelSave
   Else
        InsertQueryLog (0) 'Added by Lydia 2021/11/16
        ShowNoData
   End If
   rsAD.Close
   
   '執行完不清除條件
   CmdPrt1.Enabled = True
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me

   LblFM2C(1).Caption = "1." + PUB_GetFCPGrpName("1")
   LblFM2C(2).Caption = "2." + PUB_GetFCPGrpName("2")
   LblFM2C(3).Caption = "3." + PUB_GetFCPGrpName("3")
   LblFM2C(4).Caption = "4." + PUB_GetFCPGrpName("4")
   
   Call txtFM2_Validate(3, False)
   
   If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
       MkDir strExcelPath
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060331 = Nothing
End Sub

'  畫面輸入檢查
Private Function FormCheck() As Boolean
Dim bolTmp As Boolean
   
   FormCheck = False
   For Each oText In txtFM2
       txtFM2_Validate oText.Index, bolTmp
       If bolTmp = True Then
           Exit Function
       End If
   Next
   
   If txtFM2(0) = "" And txtFM2(1) = "" Then
      FormCheck = False
      txtFM2(0).SetFocus
      MsgBox "來函日期不可空白！", , MsgText(5)
      Exit Function
   End If
   If txtFM2(0) > txtFM2(1) Then
        FormCheck = False
        txtFM2(0).SetFocus
        MsgBox "來函日期起值不可大於迄值！", , MsgText(5)
        Exit Function
   End If
     
   If txtFM2(3) <> "" And txtFM2(4) <> "" And txtFM2(3) > txtFM2(4) Then
       MsgBox "組別起值不可大於迄值！", vbCritical
       txtFM2(3).SetFocus
       Call txtFM2_GotFocus(3)
       Exit Function
   End If

   FormCheck = True
End Function

Private Sub txtFM2_GotFocus(Index As Integer)
    TextInverse txtFM2(Index)
End Sub

Private Sub txtFM2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
     If Index = 2 Then
         KeyAscii = UpperCase(KeyAscii)
     Else
         KeyAscii = Pub_NumAscii(KeyAscii)
     End If
End Sub

Private Sub txtFM2_Validate(Index As Integer, Cancel As Boolean)
    
    Select Case Index
        Case 0, 1  '來函日期
           If PUB_CheckKeyInDate(txtFM2(Index)) = -1 Then
              GoTo EXITSUB
           End If
        Case 2     '系統類別
            If txtFM2(Index).Text = "" Then Exit Sub
            txtFM2(Index) = Replace(txtFM2(Index), " ", "")
            If txtFM2(Index) = "ALL" Then txtFM2(Index) = "FCP,P"
            
            If Trim(txtFM2(Index)) = "" Then
                 MsgBox "系統類別不可空白！", vbCritical
                 GoTo EXITSUB
            ElseIf txtFM2(Index) <> "ALL" Then
                If PUB_CheckSKAddCross(strUserNum, Systemkind_g, True, txtFM2(Index)) = False Then
                       GoTo EXITSUB
                End If
            End If
        Case 3, 4 '組別
            If txtFM2(Index).Text = "" Then Exit Sub
            If InStr("1,2,3,4", txtFM2(Index)) = 0 Then
                MsgBox "請輸入1~4 ！", vbCritical
                GoTo EXITSUB
            Else
                Call SetCombo
            End If
            
    End Select
    
    Exit Sub
    
EXITSUB:
    txtFM2(Index).SetFocus
    txtFM2_GotFocus Index
    Cancel = True
End Sub

'依組別設工程師清單
Private Sub SetCombo()

    If txtFM2(3).Tag & txtFM2(4).Tag <> txtFM2(3).Text & txtFM2(4).Text Then
         Combo1.Clear
         '工程師排除總經理的編號(94099)
         'Modified by Lydia 2019/08/02 排除F4102 (FCP年費不續辦)
         'modify by sonia 2021/1/22 再排除F4104,F4105
         strSql = "SELECT ST01,ST02,ST16 FROM STAFF WHERE ST04='1' AND ST03='F21' AND ST01<>'94099' AND ST01<>'F4102' and st01<>'F4104' and st01<>'F4105' " & _
                     "AND ST16>='" & txtFM2(3) & "' AND ST16<='" & txtFM2(4) & "' " & _
                     "ORDER BY ST16,ST01 "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
             RsTemp.MoveFirst
             Do While Not RsTemp.EOF
                  Combo1.AddItem RsTemp.Fields("ST01") & " " & RsTemp.Fields("ST02")
                  RsTemp.MoveNext
             Loop
         End If
    End If
    txtFM2(3).Tag = txtFM2(3).Text
    txtFM2(4).Tag = txtFM2(4).Text
End Sub

'產生Excel檔案
Private Sub ProcExcelSave()
Dim cntPage As Integer 'Excel檔的工作表數量
Dim mSt16 As String '組別
Dim xlsWDay As New Excel.Application
Dim wksWDay As New Worksheet
Dim strFileName As String '檔案名稱
Dim strFileList As String '所有檔案名稱
Dim iRow As Integer

Dim pMan As String '工程師-分工作表
Dim intPage As Integer '工作表編號
Dim strColName As String '欄位名稱
Dim strColW As String    '欄寬
Dim tmpArr1 As Variant, tmpArr2 As Variant
Dim xCols As Integer '行位置
Dim endX As String '最後一欄


On Error GoTo ErrHnd
   'Memo by Lydia 2019/09/24 特別抓資料
   'strColName = "組別,工程師姓名,代理人彼所案號,本所案號,來函性質,來函收文日,發文日,日曆天,平均日曆天"
   'strColW = "4,10,16,16,15,11,9.5,7,11"
   strColName = "本所案號,來函性質,來函收文日,發文日,日曆天,平均日曆天"
   strColW = "16,15,11,9.5,7,11"
   
   tmpArr1 = Split(strColName, ",")
   tmpArr2 = Split(strColW, ",")

    rsAD.MoveFirst
     
    Do While Not rsAD.EOF
        '不同組別分不同excel; 不同工程師分不同工作表
        'Memo by Lydia 2019/09/24 特別抓資料: 全部在同一頁 If mSt16 & pMan <> "" & rsAD.Fields("ST16") & rsAD.Fields("CP14") And mSt16 = "" Then
        If mSt16 & pMan <> "" & rsAD.Fields("ST16") & rsAD.Fields("CP14") Then
            If pMan <> "" Then
               '設平均日曆天
                With wksWDay.Range(endX & "2")
                    .Value = "=AVERAGE(" & Chr(Asc(endX) - 1) & "2:" & Chr(Asc(endX) - 1) & iRow & ")"
                    .NumberFormatLocal = "##,##0.00"
                End With
            End If
            
            '不同組別分不同excel
            If mSt16 <> "" & rsAD.Fields("ST16") Then
                If mSt16 <> "" Then
                    xlsWDay.Sheets(1).Select '選擇工作表
                    '判斷版本
                    If Val(xlsWDay.Version) < 12 Then
                         xlsWDay.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
                    Else
                         xlsWDay.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
                    End If
                    xlsWDay.Workbooks.Close
                    xlsWDay.Quit
                    Set wksWDay = Nothing
                    Set xlsWDay = Nothing
                End If
                
                intI = 1
                strExc(0) = "select count(distinct(cp14)) cnt from(" & strMid & ") where st16='" & rsAD.Fields("ST16") & "' "
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then cntPage = Val(RsTemp.Fields("cnt")) + 1
                
                intPage = 0
                pMan = ""
            End If
        End If
               
        '不同工程師分不同工作表
        'Memo by Lydia 2019/09/24 特別抓資料 全部在同一頁 'If pMan <> "" & rsAD.Fields("CP14") And pMan = "" Then
        If pMan <> "" & rsAD.Fields("CP14") Then
           intPage = intPage + 1
           If intPage = 1 Then
              '不同組別分不同檔案
              strExc(1) = ""
              If txtFM2(0) <> "" Then strExc(1) = TransDate(txtFM2(0), 2) & strExc(1) & "~"
              If txtFM2(1) <> "" Then strExc(1) = strExc(1) & IIf(strExc(1) = "", "~", "") & TransDate(txtFM2(1), 2)
              strFileName = strExcelPath & strExc(1) & "外專" & PUB_GetFCPGrpName("" & rsAD.Fields("ST16")) & "OA來函已發文明細" & MsgText(43)
              strFileList = strFileList & IIf(strFileList <> "", "、", "") & strFileName

              If Dir(strFileName) <> "" Then
                 Kill strFileName
              End If
              xlsWDay.SheetsInNewWorkbook = cntPage
              xlsWDay.Workbooks.add
              xlsWDay.Visible = False '預設不顯示
           Else
               '設平均日曆天
                With wksWDay.Range(endX & "2")
                    .Value = "=AVERAGE(" & Chr(Asc(endX) - 1) & "2:" & Chr(Asc(endX) - 1) & iRow & ")"
                    .NumberFormatLocal = "##,##0.00"
                End With
           End If
    
           Set wksWDay = xlsWDay.Worksheets(intPage)
           xlsWDay.Sheets(intPage).Select '選擇工作表
           xlsWDay.Worksheets(intPage).Name = "" & rsAD.Fields("CP14N") '工作表名稱
           '設定抬頭
           iRow = 1
           xCols = Asc("A")
           For intI = 0 To UBound(tmpArr1)
               If Trim(tmpArr1(intI)) <> "" Then
                   wksWDay.Range(Chr(xCols + intI) & iRow).Value = Trim(tmpArr1(intI))
                   wksWDay.Range(Chr(xCols + intI) & iRow).HorizontalAlignment = xlCenter
                   wksWDay.Range(Chr(xCols + intI) & ":" & Chr(xCols + intI)).ColumnWidth = Val(tmpArr2(intI))
                   If intI > 1 Then '日期-置中
                       wksWDay.Range(Chr(xCols + intI) & ":" & Chr(xCols + intI)).HorizontalAlignment = xlCenter
                   End If
                   endX = Chr(xCols + intI)
               End If
           Next intI
           wksWDay.Range("2:2").Select
           xlsWDay.ActiveWindow.FreezePanes = True '凍結窗格
           wksWDay.Range("A1").Select
           wksWDay.Range("A:" & endX).Font.Size = 11
        End If
        
        iRow = iRow + 1
        xCols = Asc("A")
        'Memo by Lydia 2019/09/24 特別抓資料: 代理人彼所案號
'        '組別
'        With wksWDay.Range(Chr(xCols) & iRow)
'            .Value = "" & rsAD.Fields("ST16")
'            .NumberFormatLocal = "@"
'        End With
'        xCols = xCols + 1
'        '工程師姓名
'        With wksWDay.Range(Chr(xCols) & iRow)
'            .Value = "" & rsAD.Fields("CP14N")
'            .NumberFormatLocal = "@"
'        End With
'        xCols = xCols + 1
'        '代理人彼所案號
'        With wksWDay.Range(Chr(xCols) & iRow)
'            .Value = "" & rsAD.Fields("PA77")
'            .NumberFormatLocal = "@"
'        End With
'        xCols = xCols + 1
        '-------end 1234
        
        '本所案號
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = "" & rsAD.Fields("CASENO")
            .NumberFormatLocal = "@"
        End With
        xCols = xCols + 1
        
        '名稱
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = "" & rsAD.Fields("CP10N")
            .NumberFormatLocal = "@"
        End With
        xCols = xCols + 1
        
        '來函收文日
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = Trim("" & rsAD.Fields("CP05"))
            .NumberFormatLocal = "@"
        End With
        xCols = xCols + 1
        
        '發文日
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = Trim("" & rsAD.Fields("CP27"))
            .NumberFormatLocal = "@"
        End With
        xCols = xCols + 1
        
        '日曆天1
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = Trim("" & rsAD.Fields("CDATE"))
            .NumberFormatLocal = "@"
        End With
        xCols = xCols + 1
        
       mSt16 = "" & rsAD.Fields("ST16")
       pMan = "" & rsAD.Fields("CP14")
JumpNextRec:
       rsAD.MoveNext
    Loop

    '最後一頁-設平均日曆天
    With wksWDay.Range(endX & "2")
        .Value = "=AVERAGE(" & Chr(Asc(endX) - 1) & "2:" & Chr(Asc(endX) - 1) & iRow & ")"
        .NumberFormatLocal = "##,##0.00"
    End With
    xlsWDay.Sheets(1).Select '選擇工作表
   
   '判斷版本
   If Val(xlsWDay.Version) < 12 Then
        xlsWDay.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
        xlsWDay.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If

   xlsWDay.Workbooks.Close
   xlsWDay.Quit
   Set wksWDay = Nothing
   Set xlsWDay = Nothing
   'Modify by Amy 2021/06/22 原:strExcelPath 改中文字顯示
   MsgBox "Excel檔案產生完成！" & vbCrLf & "檔案位置：" & strExcelPathN
   Exit Sub

ErrHnd:

   MsgBox Err.Description
End Sub
