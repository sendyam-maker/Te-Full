VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010022 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶特殊紀錄異動"
   ClientHeight    =   5172
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8328
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5172
   ScaleWidth      =   8328
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   570
      Width           =   8085
      _ExtentX        =   14245
      _ExtentY        =   7853
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "目前設定"
      TabPicture(0)   =   "frm010022.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCUid"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label16"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblCusName"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt1(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt1(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chk1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdSearch"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CmdSure"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm010022.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ChkUse"
      Tab(1).Control(1)=   "cmdQuery"
      Tab(1).Control(2)=   "MSHFlexGrid1"
      Tab(1).Control(3)=   "txtData(7)"
      Tab(1).Control(4)=   "txtData(6)"
      Tab(1).Control(5)=   "txtData(5)"
      Tab(1).Control(6)=   "txtData(4)"
      Tab(1).Control(7)=   "txtData(3)"
      Tab(1).Control(8)=   "txtData(2)"
      Tab(1).Control(9)=   "txtData(1)"
      Tab(1).Control(10)=   "txtData(0)"
      Tab(1).Control(11)=   "Label4"
      Tab(1).Control(12)=   "Label7"
      Tab(1).Control(13)=   "Line3"
      Tab(1).Control(14)=   "Line2"
      Tab(1).Control(15)=   "Line1"
      Tab(1).Control(16)=   "Label3(4)"
      Tab(1).Control(17)=   "Label3(3)"
      Tab(1).Control(18)=   "Label3(2)"
      Tab(1).Control(19)=   "Label3(1)"
      Tab(1).Control(20)=   "Label3(0)"
      Tab(1).ControlCount=   21
      Begin VB.CheckBox ChkUse 
         Caption         =   "是否含無效客戶"
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   -72168
         TabIndex        =   32
         Top             =   457
         Width           =   1620
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Height          =   400
         Left            =   -68250
         TabIndex        =   28
         Top             =   390
         Width           =   800
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2445
         Left            =   -74790
         TabIndex        =   27
         Top             =   1860
         Width           =   7695
         _ExtentX        =   13568
         _ExtentY        =   4318
         _Version        =   393216
         Cols            =   8
         AllowUserResizing=   3
         FormatString    =   "V|客戶編號|客戶名稱|異動日期|序號|特殊原因|異動人員|異動說明"
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
         _Band(0).Cols   =   8
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   7
         Left            =   -72810
         MaxLength       =   7
         TabIndex        =   25
         Top             =   1260
         Width           =   1000
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   6
         Left            =   -73920
         MaxLength       =   7
         TabIndex        =   24
         Top             =   1260
         Width           =   1000
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   5
         Left            =   -69090
         MaxLength       =   1
         TabIndex        =   26
         Top             =   1275
         Width           =   600
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   4
         Left            =   -72810
         MaxLength       =   9
         TabIndex        =   22
         Top             =   855
         Width           =   1000
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   3
         Left            =   -73920
         MaxLength       =   9
         TabIndex        =   21
         Top             =   855
         Width           =   1000
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   2
         Left            =   -69090
         MaxLength       =   6
         TabIndex        =   23
         Top             =   855
         Width           =   800
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   1
         Left            =   -73080
         MaxLength       =   3
         TabIndex        =   20
         Top             =   450
         Width           =   600
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   0
         Left            =   -73920
         MaxLength       =   3
         TabIndex        =   19
         Top             =   450
         Width           =   600
      End
      Begin VB.CommandButton CmdSure 
         Caption         =   "存檔(&S)"
         Height          =   400
         Left            =   3180
         TabIndex        =   13
         Top             =   1830
         Width           =   800
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "尋找(&F)"
         Default         =   -1  'True
         Height          =   375
         Left            =   2340
         TabIndex        =   3
         Top             =   810
         Width           =   800
      End
      Begin VB.CheckBox chk1 
         Alignment       =   1  '靠右對齊
         Caption         =   "是否為特殊客戶"
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   2040
         Width           =   1605
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   5
         Top             =   2340
         Width           =   4485
         VariousPropertyBits=   -1467989989
         MaxLength       =   48
         ScrollBars      =   2
         Size            =   "7911;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   0
         Left            =   1230
         TabIndex        =   4
         Top             =   840
         Width           =   1020
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1799;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCusName 
         Height          =   255
         Left            =   1260
         TabIndex        =   31
         Top             =   1530
         Width           =   4725
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "8334;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label4 
         Height          =   255
         Left            =   -68250
         TabIndex        =   30
         Top             =   900
         Width           =   1095
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1931;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label7 
         Caption         =   "P.S.若不輸入任何條件，查詢會帶出全部記錄。"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -70290
         TabIndex        =   29
         Top             =   450
         Width           =   2025
      End
      Begin VB.Line Line3 
         X1              =   -73080
         X2              =   -72540
         Y1              =   1410
         Y2              =   1410
      End
      Begin VB.Line Line2 
         X1              =   -73110
         X2              =   -72570
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line1 
         X1              =   -73500
         X2              =   -72930
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label Label3 
         Caption         =   "目前狀態是特殊客戶：　　　　(Y:是/ N:取消 )"
         Height          =   225
         Index           =   4
         Left            =   -70920
         TabIndex        =   18
         Top             =   1320
         Width           =   3795
      End
      Begin VB.Label Label3 
         Caption         =   "異動日期："
         Height          =   225
         Index           =   3
         Left            =   -74790
         TabIndex        =   17
         Top             =   1305
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "客戶編號："
         Height          =   225
         Index           =   2
         Left            =   -74790
         TabIndex        =   16
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "智權人員："
         Height          =   225
         Index           =   1
         Left            =   -70020
         TabIndex        =   15
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "業務區：　　　　　　　"
         Height          =   225
         Index           =   0
         Left            =   -74790
         TabIndex        =   14
         Top             =   495
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "原因最多只能24個中文字"
         Height          =   180
         Left            =   5730
         TabIndex        =   12
         Top             =   2400
         Width           =   1980
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "客戶名稱："
         Height          =   180
         Left            =   270
         TabIndex        =   11
         Top             =   1530
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "原　　因："
         Height          =   180
         Left            =   270
         TabIndex        =   10
         Top             =   2370
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號："
         Height          =   180
         Left            =   270
         TabIndex        =   9
         Top             =   885
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號："
         Height          =   180
         Left            =   270
         TabIndex        =   8
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lblCUid 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1230
         TabIndex        =   7
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "PS：移除特殊客戶時，除取消Ｖ以外，請於原因欄輸入某人同意，例：X主管同意Y人員提出"
         Height          =   360
         Left            =   4140
         TabIndex        =   6
         Top             =   1860
         Width           =   3780
      End
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6855
      TabIndex        =   0
      Top             =   30
      Width           =   800
   End
End
Attribute VB_Name = "frm010022"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/06 客戶特殊紀錄異動：改成兩個頁籤”目前設定”和”多筆查詢”，有”新增”權限的人員才能看見”目前設定”頁籤。
'Memo By Sindy 2022/2/17 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
'create by nickc 2007/11/13
Option Explicit

Dim SeekCU121 As String
'Added by Lydia 2022/05/06
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Private Const cFixed As Integer = 4 '固定欄位
Dim intLastRow As Integer

Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub cmdQuery_Click()
Dim strQ1 As String, intQ As Integer
Dim rsQuery As New ADODB.Recordset
Dim bolTmp As Boolean
    
    For intQ = 0 To 7
       bolTmp = False
       Call Txtdata_Validate(intQ, bolTmp)
       If bolTmp = True Then
          Exit Sub
       End If
    Next intQ
    
    Screen.MousePointer = vbHourglass
    cmdQuery.Enabled = False
    ClearQueryLog (Me.Name) '清除查詢印表記錄檔欄位
    strQ1 = ""
    '業務區
    If Trim(txtData(0)) <> "" Then
        strQ1 = strQ1 & " and cu12>='" & Trim(txtData(0)) & "' "
    End If
    If Trim(txtData(1)) <> "" Then
        strQ1 = strQ1 & " and cu12<='" & Trim(txtData(1)) & "' "
    End If
    pub_QL05 = pub_QL05 & ";" & Label3(0).Caption & txtData(0) & "-" & txtData(1)
    '智權人員編號
    If Trim(txtData(2)) <> "" Then
        strQ1 = strQ1 & " and cu13='" & Trim(txtData(2)) & "' "
    End If
    pub_QL05 = pub_QL05 & ";" & Label3(1).Caption & txtData(2)
    '客戶編號
    If Trim(txtData(3)) <> "" Then
        strQ1 = strQ1 & " and cl01>='" & Trim(txtData(3)) & "' "
    End If
    If Trim(txtData(4)) <> "" Then
        strQ1 = strQ1 & " and cl01<='" & Trim(txtData(4)) & "' "
    End If
    pub_QL05 = pub_QL05 & ";" & Label3(2).Caption & txtData(3) & "-" & txtData(4)
    '異動日期
    If Trim(txtData(6)) <> "" Then
        strQ1 = strQ1 & " and cl02>='" & DBDATE(Trim(txtData(6))) & "' "
    End If
    If Trim(txtData(7)) <> "" Then
        strQ1 = strQ1 & " and cl02<='" & DBDATE(Trim(txtData(7))) & "' "
    End If
    pub_QL05 = pub_QL05 & ";" & Label3(3).Caption & txtData(6) & "-" & txtData(7)
    'Added by Lydia 2024/05/10 判斷無效客戶
    If ChkUse.Value = False Then
      'Modified by Lydia 2024/05/13 +不再使用
      strQ1 = strQ1 & " and instr(st02,'無效')=0 and instr(cu80||',','不再使用')=0 "
    End If
    pub_QL05 = pub_QL05 & ";" & ChkUse.Caption & ":" & IIf(ChkUse.Value = 1, "Y", "N")
    'end 2024/05/10
    
    '目前狀態是特殊客戶
    If Trim(txtData(5)) <> "" Then
        pub_QL05 = pub_QL05 & ";目前狀態是特殊客戶：" & txtData(5)
        strQ1 = "select '' as V, cl01 as cno,nvl(cu04,nvl(cu05,cu06)) cname,sqldatet(cl02) cdate2,cl03 as sno,cl04 as smemo,st01,st02 as sname,cl06 as smemo2 " & _
                   "From CustSpecialLog, customer, staff " & _
                   "where substr(cl01,1,8)=cu01(+) and substr(cl01,9,1)=cu02(+) and cl05=st01(+) " & _
                   "and (cl01,cl02,cl03) in (select x01,substr(xmax,1,8) x02,substr(xmax,9) x03 from ( " & _
                   "select cl01 as X01,max(cl02||cl03) as xmax From CustSpecialLog, customer, staff " & _
                   "where substr(cl01,1,8)=cu01(+) and substr(cl01,9,1)=cu02(+) and cl05=st01(+) " & strQ1 & _
                   "group by cl01 ) ) and cl06=" & CNULL(IIf(txtData(5) = "Y", "設定成特殊", "移除特殊")) & _
                   " order by cno,cdate2,sno"
    Else
        strQ1 = "select '' as V, cl01 as cno,nvl(cu04,nvl(cu05,cu06)) cname,sqldatet(cl02) cdate2,cl03 as sno,cl04 as smemo,st01,st02 as sname,cl06 as smemo2 " & _
                    "From CustSpecialLog, customer, staff " & _
                    "where substr(cl01,1,8)=cu01(+) and substr(cl01,9,1)=cu02(+) and cl05=st01(+)" & strQ1 & _
                    "order by cno,cdate2,sno "
    End If
    Call SetGrd(True) '清空
    intQ = 1
    Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
    If intQ = 1 Then
         Call InsertQueryLog(rsQuery.RecordCount)
         MSHFlexGrid1.FixedCols = 0
         Set MSHFlexGrid1.Recordset = rsQuery
         Call SetGrd
         MSHFlexGrid1.FixedCols = cFixed
    Else
         Call InsertQueryLog(0)
         ShowNoData
    End If
    
   Set rsQuery = Nothing
   cmdQuery.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
CheckOC3
AdoRecordSet3.CursorLocation = adUseClient
'Modified by Lydia 2024/05/10 判斷無效客戶
'AdoRecordSet3.Open "select * from customer where CU01='" & Mid(ChangeCustomerL(Txt1(0)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(Txt1(0)), 9, 1) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
'Modified by Lydia 2024/05/13 +CU80
strExc(0) = "select cu01,cu02,cu04,cu121,cu13||' '||st02 as cu13n,cu80 from customer customer,staff " & _
            "where CU01='" & Mid(ChangeCustomerL(Txt1(0)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(Txt1(0)), 9, 1) & "' and cu13=st01(+) "
'strExc(0) = strExc(0) & " and instr(st02,'無效') =0 " 'Mark by Lydia 2024/05/13
AdoRecordSet3.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
'end 2024/05/10

If AdoRecordSet3.RecordCount <> 0 Then
    'Added by Lydia 2024/05/13 判斷無效客戶/不再使用cu80彈訊息,不顯示訊息
    If InStr("," & AdoRecordSet3.Fields("cu13n"), "無效") > 0 Then
       MsgBox CheckStr(AdoRecordSet3.Fields("cu01")) & CheckStr(AdoRecordSet3.Fields("cu02")) & vbCrLf & "無效客戶！", vbInformation
       GoTo JumpToClear
    ElseIf InStr("," & AdoRecordSet3.Fields("cu80"), "不再使用") > 0 Then
       MsgBox CheckStr(AdoRecordSet3.Fields("cu01")) & CheckStr(AdoRecordSet3.Fields("cu02")) & vbCrLf & "客戶狀態為不再使用！", vbInformation
       GoTo JumpToClear
    End If
    'end 2024/05/13
    lblCUID = CheckStr(AdoRecordSet3.Fields("cu01")) & CheckStr(AdoRecordSet3.Fields("cu02"))
    lblCusName = CheckStr(AdoRecordSet3.Fields("cu04"))
    Chk1.Value = IIf(CheckStr(AdoRecordSet3.Fields("CU121")) = "Y", vbChecked, vbUnchecked)
    SeekCU121 = CheckStr(AdoRecordSet3.Fields("CU121"))
    Txt1(1).SetFocus
Else
    MsgBox "查無客戶！", vbInformation
JumpToClear: 'Added by Lydia 2024/05/13
    lblCUID = ""
    lblCusName = ""
    Chk1.Value = vbUnchecked
End If
End Sub

Private Sub cmdSure_Click()
Dim IsBegin As Boolean
On Error GoTo 0
On Error GoTo ErrHand
If Txt1(1) = "" Then
    MsgBox "原因不可以空白！", vbExclamation
    Txt1(1).SetFocus
ElseIf lblCUID.Caption = "" Then
    MsgBox "請先查詢要更改的客戶！", vbExclamation
    Txt1(0).SetFocus
Else
    cnnConnection.BeginTrans
    IsBegin = True
    'Modified by Lydia 2022/10/27 有更名前後編號一併更新
    'cnnConnection.Execute "update customer set cu121=" & IIf(chk1.Value = vbChecked, "'Y'", "null") & " where cu01='" & Mid(lblCUid, 1, 8) & "' and cu02='" & Mid(lblCUid, 9, 1) & "' "
    cnnConnection.Execute "update customer set cu121=" & IIf(Chk1.Value = vbChecked, "'Y'", "null") & " where cu01='" & Mid(lblCUID, 1, 8) & "' "
    cnnConnection.Execute "insert into custspeciallog (cl01,cl02,cl03,cl04,cl05,cl06) select '" & lblCUID & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(cl03),0)+1,'" & ChgSQL(Txt1(1)) & "','" & strUserNum & "','" & IIf(SeekCU121 = "Y" And Chk1.Value = vbChecked, "無異動", IIf(SeekCU121 = "" And Chk1.Value = vbUnchecked, "無異動", IIf(SeekCU121 = "Y" And Chk1.Value = vbUnchecked, "移除特殊", "設定成特殊"))) & "' from custspeciallog where cl01='" & lblCUID & "' and cl02=to_number(to_char(sysdate,'YYYYMMDD')) "
    cnnConnection.CommitTrans
    MsgBox "更新紀錄成功！", vbInformation
    Txt1(0) = ""
    Txt1(1) = ""
    lblCUID = ""
    lblCusName = ""
    Chk1.Value = vbUnchecked
    Txt1(0).SetFocus
End If
Exit Sub
ErrHand:
    If IsBegin Then
        cnnConnection.RollbackTrans
    End If
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'Added by Lydia 2022/05/06
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction("frm010022", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm010022", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm010022", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm010022", strFind, False)
   lblCUID.BackColor = &H8000000F
   lblCusName.BackColor = &H8000000F
   lblCusName.Caption = ""
   Label4.Caption = ""
   SSTab1.Tab = 1
   Call SetGrd(True)
   If m_bInsert = False Then
      SSTab1.TabVisible(0) = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm010022 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
Txt1(Index).SelStart = 0
Txt1(Index).SelLength = Len(Txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
If Index = 0 Then
    KeyAscii = UpperCase(KeyAscii)
End If
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 0
Case 1
         If Txt1(Index) <> "" Then
            If CheckLengthIsOK(Txt1(Index), Txt1(Index).MaxLength) = False Then
                Cancel = True
            End If
         End If
Case Else
End Select
End Sub

'Added by Lydia 2022/05/06
Private Sub Txtdata_GotFocus(Index As Integer)
   TextInverse txtData(Index)
End Sub

Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If Index = 5 Then
      If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
         KeyAscii = 0
         Beep
      End If
    End If
End Sub

'add by sonia 2024/1/23
Private Sub Txtdata_LostFocus(Index As Integer)
   If Index = 3 And Trim(txtData(Index)) <> MsgText(601) Then
      txtData(Index) = Left(txtData(Index).Text & String(9, "0"), 9)
      txtData(4).Text = Left(txtData(Index).Text, 8) & "Z"
   End If
End Sub
'end 2024/1/23

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)

    Select Case Index
        Case 0, 1  '業務區
            If txtData(Index) <> "" Then
               If Left(txtData(Index), 1) <> "S" And Left(txtData(Index), 1) <> "P" And Left(txtData(Index), 1) <> "W" Then
                   MsgBox "請輸入S,P,W業務區代號！", vbCritical, "查詢條件"
                   GoTo EXITSUB
               End If
               If Index = 1 And txtData(0) > txtData(1) Then
                   MsgBox "起值不可大於迄值！", vbCritical, "查詢條件"
                   GoTo EXITSUB
               End If
            End If
        Case 2  '智權人員編號
            Label4.Caption = ""
            If txtData(Index) <> "" Then
                Label4 = GetStaffName(txtData(Index), True)
                If Label4.Caption = "" Then
                   MsgBox "請輸入正確的員工編號！", vbCritical, "查詢條件"
                   GoTo EXITSUB
                End If
            End If
        Case 3, 4 '客戶編號
            If txtData(Index) <> "" Then
               If Left(txtData(Index), 1) <> "X" Then
                   MsgBox "請輸入正確的客戶編號！", vbCritical, "查詢條件"
                   GoTo EXITSUB
               End If
               If Index = 4 And txtData(3) > txtData(4) Then
                   MsgBox "起值不可大於迄值！", vbCritical, "查詢條件"
                   GoTo EXITSUB
               End If
            End If
        Case 6, 7
            If txtData(Index) <> "" Then
               If CheckIsTaiwanDate(txtData(Index), False) = False Then
                   MsgBox "請輸入正確的日期格式！", vbCritical, "查詢條件"
                   GoTo EXITSUB
               End If
               If Index = 7 And txtData(6) > txtData(7) Then
                   MsgBox "起值不可大於迄值！", vbCritical, "查詢條件"
                   GoTo EXITSUB
               End If
            End If
    End Select
    
    Exit Sub
    
EXITSUB:
    Cancel = True
    Txtdata_GotFocus Index
    txtData(Index).SetFocus
End Sub

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   arrGridHeadText = Array("V", "客戶編號", "客戶名稱", "異動日期", "序號", "特殊原因", "ST01", "異動人員", "異動說明")
   arrGridHeadWidth = Array(300, 1000, 1400, 900, 720, 1000, 0, 900, 1000)

   
   MSHFlexGrid1.Visible = False
   MSHFlexGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
         MSHFlexGrid1.Clear
         MSHFlexGrid1.Rows = 2
   End If
       
    For iRow = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.row = 0
       MSHFlexGrid1.col = iRow
       MSHFlexGrid1.Text = arrGridHeadText(iRow)
       MSHFlexGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
    Next

   For intI = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.row = intI
        For iRow = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.col = iRow
           MSHFlexGrid1.CellBackColor = &H80000005
           If InStr("03,04", Format(iRow, "00")) > 0 Then  '置中
                MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
           End If
        Next iRow
   Next intI
   MSHFlexGrid1.Visible = True
   
End Sub

Private Sub MSHFlexGrid1_DblClick()
Dim inX As Integer, inY As Integer
Dim lngColor As Long

    If SSTab1.TabVisible(0) = True Then
       For inX = 1 To MSHFlexGrid1.Rows - 1
          MSHFlexGrid1.row = inX
          MSHFlexGrid1.col = 0
          If Trim(MSHFlexGrid1.Text) = "V" Then
              MSHFlexGrid1.Text = ""
              MSHFlexGrid1.col = 0
              MSHFlexGrid1.CellBackColor = MSHFlexGrid1.BackColor
              MSHFlexGrid1.col = cFixed + 1
              lngColor = MSHFlexGrid1.CellBackColor
              For inY = 1 To cFixed
                  MSHFlexGrid1.col = inY
                  MSHFlexGrid1.CellBackColor = lngColor
              Next inY
              Txt1(0).Text = MSHFlexGrid1.TextMatrix(inX, 1)
              SSTab1.Tab = 0
              Call cmdSearch_Click
          End If
       Next inX
    End If
End Sub

Private Sub MSHFlexGrid1_Click()
Dim intRow As Integer
Dim lngColor As Long
   With MSHFlexGrid1
       If .MouseRow > 0 Then
          lngColor = .CellBackColor
          GridClick MSHFlexGrid1, intLastRow, 0, 0, cFixed, "V", lngColor
       End If
   End With
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow MSHFlexGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MSHFlexGrid1.col = nCol
   MSHFlexGrid1.row = nRow
   If Me.MSHFlexGrid1.row < 1 And Me.MSHFlexGrid1.Text <> "V" Then
      If InStr("序號", Me.MSHFlexGrid1.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.MSHFlexGrid1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.MSHFlexGrid1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.MSHFlexGrid1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.MSHFlexGrid1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub
