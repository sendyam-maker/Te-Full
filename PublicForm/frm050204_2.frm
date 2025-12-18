VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm050204_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人案件性質統計"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   945
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   3
      Left            =   8430
      TabIndex        =   5
      Top             =   70
      Width           =   840
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "案件明細(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   2
      Left            =   5970
      TabIndex        =   4
      Top             =   70
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   3480
      Left            =   1080
      TabIndex        =   3
      Top             =   6840
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   6138
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      HighLight       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MFG1 
      Height          =   4545
      Left            =   30
      TabIndex        =   2
      Top             =   1140
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8017
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
      _Band(0).Cols   =   13
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "代理人資料(&A)"
      Height          =   400
      Index           =   0
      Left            =   4680
      TabIndex        =   1
      Top             =   70
      Width           =   1300
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   7200
      TabIndex        =   0
      Top             =   70
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "－"
      Height          =   180
      Index           =   2
      Left            =   5520
      TabIndex        =   19
      Top             =   870
      Width           =   345
   End
   Begin VB.Label lblDateTo 
      Alignment       =   2  '置中對齊
      Height          =   180
      Left            =   5910
      TabIndex        =   18
      Top             =   870
      Width           =   825
   End
   Begin VB.Label lblDateFrom 
      Alignment       =   2  '置中對齊
      Height          =   180
      Left            =   4650
      TabIndex        =   17
      Top             =   870
      Width           =   825
   End
   Begin VB.Label lblSysKind 
      Alignment       =   2  '置中對齊
      Height          =   180
      Left            =   4650
      TabIndex        =   16
      Top             =   570
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "－"
      Height          =   180
      Index           =   1
      Left            =   1950
      TabIndex        =   15
      Top             =   870
      Width           =   345
   End
   Begin VB.Label lblFagentTo 
      Alignment       =   2  '置中對齊
      Height          =   180
      Left            =   2340
      TabIndex        =   14
      Top             =   870
      Width           =   585
   End
   Begin VB.Label lblFagentFrom 
      Alignment       =   2  '置中對齊
      Height          =   180
      Left            =   1320
      TabIndex        =   13
      Top             =   870
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "－"
      Height          =   180
      Index           =   0
      Left            =   1950
      TabIndex        =   12
      Top             =   570
      Width           =   345
   End
   Begin VB.Label lblNATo 
      Alignment       =   2  '置中對齊
      Height          =   180
      Left            =   2340
      TabIndex        =   11
      Top             =   570
      Width           =   585
   End
   Begin VB.Label lblNAFrom 
      Alignment       =   2  '置中對齊
      Height          =   180
      Left            =   1320
      TabIndex        =   10
      Top             =   570
      Width           =   585
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      Caption         =   "日期:"
      Height          =   180
      Index           =   0
      Left            =   3450
      TabIndex        =   9
      Top             =   870
      Width           =   1125
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "系統類別:"
      Height          =   180
      Left            =   3450
      TabIndex        =   8
      Top             =   570
      Width           =   1125
   End
   Begin VB.Label Label10 
      Alignment       =   1  '靠右對齊
      Caption         =   "代理人國籍:"
      Height          =   180
      Left            =   150
      TabIndex        =   7
      Top             =   870
      Width           =   1125
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      Caption         =   "申請國家:"
      Height          =   180
      Left            =   150
      TabIndex        =   6
      Top             =   570
      Width           =   1125
   End
End
Attribute VB_Name = "frm050204_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/15 改成Form2.0 ; MFG1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/2 日期欄已修改
Option Explicit

Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, strSQL11 As String, strSQL22 As String, strSQL33 As String
Dim strSql As String, i As Integer, j As Integer, s As Integer, strTemp As Variant, StrTest As String, strTemp1 As Variant, StrTest2 As String
Dim Int01 As Integer, Int02 As Integer, Int03 As Integer, Int04 As Integer, Int05 As Integer, StrSQL6 As String
Dim IntT01 As Integer, IntT02 As Integer, IntT03 As Integer, IntT04 As Integer, IntT05 As Integer
Dim StrR08001 As String
Dim StrR08002 As String
Dim StrR08003 As String
Dim StrR08004 As String
Dim StrR08005 As String
Dim StrR08006 As String
Dim StrR08007 As String
Dim StrR08008 As String
Dim StrR08009 As String
Dim StrR08010 As String
Public cmdState As Integer 'Add by Amy 2018/02/27 記錄按鈕

Private Sub cmdok_Click(Index As Integer)
'Modify by Amy 2018/02/27 按了下列按鈕會造成無法關閉表單,故改寫法
cmdState = Index
PubShowNextData
Exit Sub

  'Mark by Amy 2018/02/27 改寫至PubShowNextData
'  Select Case Index
'  Case 3 '結束
'      Unload Me
'      Unload frm050204_1
'  Case 2 '案件明細
'        Me.Enabled = False
'        For i = 1 To MFG1.Rows - 1
'            MFG1.row = i
'            MFG1.col = 0
'            If MFG1.Text = "V" Then
'                MFG1.col = 1
'                bolToEndByNick = False
'                Screen.MousePointer = vbHourglass
'                frm050204_3.Tag = MFG1.Text ' StrTag  傳代理人代號
'                frm050204_3.grdDataList.Tag = Me.MFG1.TextMatrix(i, 7) ' StrTag  傳案件性質代號
'                frm050204_3.StrMenu
'                Screen.MousePointer = vbDefault
'                Me.Hide
'                frm050204_3.Show
'                Do
'                  DoEvents
'                  If bolToEndByNick = True Then Unload Me: Unload frm050204_1: Exit Sub
'                Loop Until Not frm050204_3.Visible
'                Unload frm050204_3
'                MFG1.col = 0
'                MFG1.Text = ""
'                For j = 0 To MFG1.Cols - 1
'                  MFG1.row = i
'                  MFG1.col = j
'                  MFG1.CellBackColor = QBColor(15)
'                Next j
'            End If
'        Next i
'        Me.Enabled = True
'        Me.Show
'  Case 1 '回前畫面
'    Unload Me
'    frm050204_1.Show
'  Case 0 '代理人資料
'        Me.Enabled = False
'        For i = 1 To MFG1.Rows - 1
'            MFG1.row = i
'            MFG1.col = 0
'            If MFG1.Text = "V" Then
'                MFG1.col = 1
'                bolToEndByNick = False
'                Screen.MousePointer = vbHourglass
'                frm100101_10.Show
'                frm100101_10.Hide
'
'                frm100101_10.Tag = MFG1.Text ' StrTag  傳代理人代號
'                frm100101_10.StrMenu
'                Screen.MousePointer = vbDefault
'                Me.Hide
'                frm100101_10.Show
'                Do
'                DoEvents
'                If bolToEndByNick = True Then Unload Me: Me.Enabled = True: Me.Show: Exit Sub
'                Loop Until Not frm100101_10.Visible
'                Unload frm100101_10
'                MFG1.col = 0
'                MFG1.Text = ""
'                For j = 0 To MFG1.Cols - 1
'                  MFG1.row = i
'                  MFG1.col = j
'                  MFG1.CellBackColor = QBColor(15)
'                Next j
'            End If
'        Next i
'        Me.Enabled = True
'        Me.Show
'  Case Else
'  End Select
End Sub

'Add by Amy 2018/02/27
Public Sub PubShowNextData()
    Select Case cmdState
        '代理人資料/案件明細
        Case 0, 2
            Me.Enabled = False
            For i = 1 To MFG1.Rows - 1
                MFG1.row = i
                MFG1.col = 0
                If MFG1.Text = "V" Then
                    MFG1.col = 0
                    MFG1.Text = ""
                    For j = 0 To MFG1.Cols - 1
                        MFG1.row = i
                        MFG1.col = j
                        MFG1.CellBackColor = QBColor(15)
                    Next j
                    If fnSaveParentForm(Me) = False Then
                        Me.Enabled = True
                        Exit Sub
                    End If
                    MFG1.col = 1
                    Screen.MousePointer = vbHourglass
                    '代理人資料
                    If cmdState = 0 Then
                        frm100101_10.Tag = MFG1.Text ' StrTag  傳代理人代號
                        frm100101_10.StrMenu
                        frm100101_10.Show
                    '案件明細
                    Else
                        frm050204_3.Tag = MFG1.Text ' StrTag  傳代理人代號
                        frm050204_3.grdDataList.Tag = Me.MFG1.TextMatrix(i, 7) ' StrTag  傳案件性質代號
                        frm050204_3.StrMenu
                        frm050204_3.Show
                    End If
                    Screen.MousePointer = vbDefault
                    Me.Enabled = True
                    Exit Sub
                End If
            Next i
            Me.Enabled = True
        '回前畫面
        Case 1
            tmpBol = fnCancelNowFormAndShowParentForm(Me)
        '結束
        Case 3
            fnCloseAllFrm100
            Unload Me 'Added by Lydia 2021/09/15
    End Select
End Sub

Private Sub Form_Load()

    MoveFormToCenter Me
    cmdState = -1 'Add by Amy 2018/02/27
    
    Me.lblNAFrom.Caption = frm050204_1.Txt1(0).Text
    Me.lblNATo.Caption = frm050204_1.Txt1(1).Text
    Me.lblFagentFrom.Caption = frm050204_1.Txt1(2).Text
    Me.lblFagentTo.Caption = frm050204_1.Txt1(3).Text
    Me.lblSysKind.Caption = frm050204_1.Txt1(10).Text
    Me.Label3(0).Caption = IIf(frm050204_1.Txt1(4).Text = "1", "收文日期:", "發文日期:")
    Me.lblDateFrom.Caption = ChangeTStringToTDateString(frm050204_1.Txt1(5).Text)
    Me.lblDateTo.Caption = ChangeTStringToTDateString(frm050204_1.Txt1(6).Text)
    
    MFG1.Rows = 2
    MFG1.Cols = 8
    MFG1.FixedRows = 1
    MFG1.FixedCols = 0
    MFG1.ColWidth(0) = 200
    MFG1.ColWidth(1) = 1200
    MFG1.ColWidth(2) = 1000
    MFG1.ColWidth(3) = 2500
    MFG1.ColWidth(4) = 1500
    MFG1.ColWidth(5) = 1000
    MFG1.ColWidth(6) = 1000
    MFG1.ColWidth(7) = 0
    With MFG1
        .TextMatrix(0, 0) = "V"
        .TextMatrix(0, 1) = "代理人代號"
        .TextMatrix(0, 2) = "代理人國籍"
        .TextMatrix(0, 3) = "代理人"
        .TextMatrix(0, 4) = "案件性質"
        .TextMatrix(0, 5) = "件數"
        .TextMatrix(0, 6) = "代理人小計"
        .TextMatrix(0, 7) = "案件性質代號"
   End With
StrMenu1
End Sub

Sub StrMenu1()
Dim StrSQLa As String
Dim strFagentNo As String '記錄代理人代號
Dim dblCnt As Double '代理人件數小計
Dim ii As Integer '回圈序號

Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R050204_1 where id='" & strUserNum & "' "
'只有專利,商標,法務
'檢查收發文
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL6 = ""
'系統類別
If Len(frm050204_1.Txt1(10)) <> 0 Then
   strSQL1 = strSQL1 + " and cp01 in (" & SQLGrpStr(frm050204_1.Txt1(10), 1) & ") "
   strSQL2 = strSQL2 + " and cp01 in (" & SQLGrpStr(frm050204_1.Txt1(10), 2) & ") "
   StrSQL3 = StrSQL3 + " and cp01 in (" & SQLGrpStr(frm050204_1.Txt1(10), 3) & ") "
   pub_QL05 = pub_QL05 & ";" & frm050204_1.Label7 & frm050204_1.Txt1(10) 'Add By Sindy 2010/01/22
End If
'無取消收文日
strSQL1 = strSQL1 + " AND CP57 IS NULL "
strSQL2 = strSQL2 + " AND CP57 IS NULL "
StrSQL3 = StrSQL3 + " AND CP57 IS NULL "
'收文
If frm050204_1.Txt1(4) = "1" Then
   If Len(Trim(frm050204_1.Txt1(5))) <> 0 Then
      strSQL1 = strSQL1 & " and cp05>=" & Val(ChangeTStringToWString(frm050204_1.Txt1(5))) & " "
      strSQL2 = strSQL2 & " and cp05>=" & Val(ChangeTStringToWString(frm050204_1.Txt1(5))) & " "
      StrSQL3 = StrSQL3 & " and cp05>=" & Val(ChangeTStringToWString(frm050204_1.Txt1(5))) & " "
   End If
   If Len(Trim(frm050204_1.Txt1(6))) <> 0 Then
      strSQL1 = strSQL1 & " and cp05<=" & Val(ChangeTStringToWString(frm050204_1.Txt1(6))) & " "
      strSQL2 = strSQL2 & " and cp05<=" & Val(ChangeTStringToWString(frm050204_1.Txt1(6))) & " "
      StrSQL3 = StrSQL3 & " and cp05<=" & Val(ChangeTStringToWString(frm050204_1.Txt1(6))) & " "
   End If
   pub_QL05 = pub_QL05 & ";" & "收文" & frm050204_1.Label3(0) & frm050204_1.Txt1(5) & "-" & frm050204_1.Txt1(6) 'Add By Sindy 2010/01/22
'發文
Else
   If Len(Trim(frm050204_1.Txt1(5))) <> 0 Then
      strSQL1 = strSQL1 & " and cp27>=" & Val(ChangeTStringToWString(frm050204_1.Txt1(5))) & " "
      strSQL2 = strSQL2 & " and cp27>=" & Val(ChangeTStringToWString(frm050204_1.Txt1(5))) & " "
      StrSQL3 = StrSQL3 & " and cp27>=" & Val(ChangeTStringToWString(frm050204_1.Txt1(5))) & " "
   End If
   If Len(Trim(frm050204_1.Txt1(6))) <> 0 Then
      strSQL1 = strSQL1 & " and cp27<=" & Val(ChangeTStringToWString(frm050204_1.Txt1(6))) & " "
      strSQL2 = strSQL2 & " and cp27<=" & Val(ChangeTStringToWString(frm050204_1.Txt1(6))) & " "
      StrSQL3 = StrSQL3 & " and cp27<=" & Val(ChangeTStringToWString(frm050204_1.Txt1(6))) & " "
   End If
   pub_QL05 = pub_QL05 & ";" & "發文" & frm050204_1.Label3(0) & frm050204_1.Txt1(5) & "-" & frm050204_1.Txt1(6) 'Add By Sindy 2010/01/22
End If
'申請國家
If Len(Trim(frm050204_1.Txt1(0))) <> 0 Then
   strSQL1 = strSQL1 + " and PA09>='" & frm050204_1.Txt1(0) & "' "
   strSQL2 = strSQL2 + " and TM10>='" & frm050204_1.Txt1(0) & "' "
   StrSQL3 = StrSQL3 + " and LC15>='" & frm050204_1.Txt1(0) & "' "
End If
If Len(Trim(frm050204_1.Txt1(1))) <> 0 Then
   strSQL1 = strSQL1 & " and PA09<='" & frm050204_1.Txt1(1) & "' "
   strSQL2 = strSQL2 & " and TM10<='" & frm050204_1.Txt1(1) & "' "
   StrSQL3 = StrSQL3 & " and LC15<='" & frm050204_1.Txt1(1) & "' "
End If
If Len(Trim(frm050204_1.Txt1(0))) <> 0 Or Len(Trim(frm050204_1.Txt1(1))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm050204_1.Label6 & frm050204_1.Txt1(0) & "-" & frm050204_1.Txt1(1) 'Add By Sindy 2010/01/22
End If
'代理人國籍
If Len(Trim(frm050204_1.Txt1(2))) <> 0 Then
   StrSQL6 = StrSQL6 + " and fa10>='" & frm050204_1.Txt1(2) & "' "
End If
If Len(Trim(frm050204_1.Txt1(3))) <> 0 Then
   StrSQL6 = StrSQL6 & " and fa10<='" & frm050204_1.Txt1(3) & "z' "
End If
If Len(Trim(frm050204_1.Txt1(2))) <> 0 Or Len(Trim(frm050204_1.Txt1(3))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm050204_1.Label10 & frm050204_1.Txt1(2) & "-" & frm050204_1.Txt1(3) 'Add By Sindy 2010/01/22
End If
'代理人
If Len(frm050204_1.Txt1(7)) <> 0 Then
    strSQL1 = strSQL1 & " and decode(pa09,'000',pa75,cp44)='" & GetNewFagent(frm050204_1.Txt1(7)) & "' "
    strSQL2 = strSQL2 & " and decode(tm10,'000',tm44,cp44)='" & GetNewFagent(frm050204_1.Txt1(7)) & "' "
    StrSQL3 = StrSQL3 & " and decode(lc15,'000',lc22,cp44)='" & GetNewFagent(frm050204_1.Txt1(7)) & "' "
    pub_QL05 = pub_QL05 & ";" & frm050204_1.Label4 & frm050204_1.Txt1(7) 'Add By Sindy 2010/01/22
End If
'案件性質
If Len(frm050204_1.Txt1(8)) <> 0 Then
    strSQL1 = strSQL1 + " and cp10>='" & frm050204_1.Txt1(8) & "' "
    strSQL2 = strSQL2 + " and cp10>='" & frm050204_1.Txt1(8) & "' "
    StrSQL3 = StrSQL3 + " and cp10>='" & frm050204_1.Txt1(8) & "' "
End If
If Len(frm050204_1.Txt1(9)) <> 0 Then
    strSQL1 = strSQL1 + " and cp10<='" & frm050204_1.Txt1(9) & "' "
    strSQL2 = strSQL2 + " and cp10<='" & frm050204_1.Txt1(9) & "' "
    StrSQL3 = StrSQL3 + " and cp10<='" & frm050204_1.Txt1(9) & "' "
End If
If Len(frm050204_1.Txt1(8)) <> 0 Or Len(frm050204_1.Txt1(9)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm050204_1.Label3(1) & frm050204_1.Txt1(8) & "-" & frm050204_1.Txt1(9) 'Add By Sindy 2010/01/22
End If
StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) ,"
                    strSql = "select fa01||decode(fa02,'','0',fa02),na03," & StrSQLa & "NEW.CPM03,count(*),'" & strUserNum & "',NEW.CP10 from (select substr(decode(pa09,'000',pa75,cp44),1,8) as a,decode(substr(decode(pa09,'000',pa75,cp44),9,1),'','0',substr(decode(pa09,'000',pa75,cp44),9,1)) as b,CP01,CP10,cpm03 from caseprogress, patent   ,CasePropertyMap where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1 & ") new,fagent,nation,SystemKind where new.a=fa01(+) and new.b=fa02(+) AND FA10=NA01(+) AND NEW.CP01=SK01(+) AND FA01 IS NOT NULL " & StrSQL6 & " group by na03," & StrSQLa & "fa01||decode(fa02,'','0',fa02),NEW.CPM03,'" & strUserNum & "',NEW.CP10 "
strSql = strSql + " union all select fa01||decode(fa02,'','0',fa02),na03," & StrSQLa & "NEW.CPM03,count(*),'" & strUserNum & "',NEW.CP10 from (select substr(decode(tm10,'000',tm44,cp44),1,8) as a,decode(substr(decode(tm10,'000',tm44,cp44),9,1),'','0',substr(decode(tm10,'000',tm44,cp44),9,1)) as b,CP01,CP10,cpm03 from caseprogress, trademark,CasePropertyMap where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2 & ") new,fagent,nation,SystemKind where new.a=fa01(+) and new.b=fa02(+) AND FA10=NA01(+) AND NEW.CP01=SK01(+) AND FA01 IS NOT NULL " & StrSQL6 & " group by na03," & StrSQLa & "fa01||decode(fa02,'','0',fa02),NEW.CPM03,'" & strUserNum & "',NEW.CP10 "
strSql = strSql + " union all select fa01||decode(fa02,'','0',fa02),na03," & StrSQLa & "NEW.CPM03,count(*),'" & strUserNum & "',NEW.CP10 from (select substr(decode(lc15,'000',lc22,cp44),1,8) as a,decode(substr(decode(lc15,'000',lc22,cp44),9,1),'','0',substr(decode(lc15,'000',lc22,cp44),9,1)) as b,CP01,CP10,cpm03 from caseprogress, lawcase  ,CasePropertyMap where cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3 & ") new,fagent,nation,SystemKind where new.a=fa01(+) and new.b=fa02(+) AND FA10=NA01(+) AND NEW.CP01=SK01(+) AND FA01 IS NOT NULL " & StrSQL6 & " group by na03," & StrSQLa & "fa01||decode(fa02,'','0',fa02),NEW.CPM03,'" & strUserNum & "',NEW.CP10 "

strSql = "INSERT INTO R050204_1 " & strSql
cnnConnection.Execute strSql
CheckOC
strSql = "SELECT '' AS V,R050204_101 AS 代理人代號, R050204_102 AS 代理人國籍,R050204_103 AS 代理人,R050204_104 AS 案件性質, SUM(R050204_105) AS 件數,'' AS 代理人件數,R050204_106 AS 案件性質代號 FROM R050204_1 WHERE ID='" & strUserNum & "' "
strSql = strSql & " GROUP BY R050204_101, R050204_102, R050204_103 ,R050204_104,R050204_106 ORDER BY R050204_101 "

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/01/22
Else
   InsertQueryLog (0) 'Add By Sindy 2010/01/22
   ShowNoData
   Me.Hide
   frm050204_1.Show
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Set MFG1.Recordset = adoRecordset
Me.MFG1.ColAlignment(0) = flexAlignCenterCenter
Me.MFG1.ColAlignment(1) = flexAlignLeftCenter
Me.MFG1.ColAlignment(2) = flexAlignLeftCenter
Me.MFG1.ColAlignment(3) = flexAlignLeftCenter
Me.MFG1.ColAlignment(4) = flexAlignLeftCenter
Me.MFG1.ColAlignment(5) = flexAlignRightCenter
Me.MFG1.ColAlignment(6) = flexAlignRightCenter
Me.MFG1.ColAlignment(7) = flexAlignRightCenter

'整理代理人件數小計
If adoRecordset.RecordCount > 0 Then
   strFagentNo = Me.MFG1.TextMatrix(1, 1)
   dblCnt = 0
   For ii = 1 To Me.MFG1.Rows - 1
      If Me.MFG1.TextMatrix(ii, 1) = strFagentNo Then
         dblCnt = dblCnt + Val(Me.MFG1.TextMatrix(ii, 5))
      Else
         Me.MFG1.TextMatrix(ii - 1, 6) = dblCnt
         strFagentNo = Me.MFG1.TextMatrix(ii, 1)
         dblCnt = Val(Me.MFG1.TextMatrix(ii, 5))
      End If
   Next ii
   Me.MFG1.TextMatrix(ii - 1, 6) = dblCnt
End If

CheckOC
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/07/18
Set frm050204_2 = Nothing
End Sub

Private Sub MFG1_Click()
MFG1.Visible = False
MFG1.col = 0
MFG1.row = MFG1.MouseRow
If MFG1.MouseRow <> 0 Then
If MFG1.Text = "V" Then
     MFG1.Text = ""
     For i = 0 To MFG1.Cols - 1
          MFG1.col = i
          MFG1.CellBackColor = QBColor(15)
    Next i
Else
     MFG1.Text = "V"
     For i = 0 To MFG1.Cols - 1
         MFG1.col = i
         MFG1.CellBackColor = &HFFC0C0
     Next i

End If
End If
MFG1.Visible = True
End Sub
