VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm04010702 
   BorderStyle     =   1  '單線固定
   Caption         =   "委任通知函客戶清單確認"
   ClientHeight    =   5745
   ClientLeft      =   255
   ClientTop       =   990
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9345
   Begin VB.CommandButton cmdOK 
      Caption         =   "全部不寄(&A)"
      Height          =   400
      Index           =   6
      Left            =   5370
      TabIndex        =   22
      Top             =   90
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "全部要寄(&C)"
      Height          =   400
      Index           =   5
      Left            =   6500
      TabIndex        =   21
      Top             =   90
      Width           =   1100
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1680
      TabIndex        =   19
      Top             =   1485
      Width           =   2775
   End
   Begin VB.TextBox txtBatchNo 
      Height          =   270
      Left            =   1305
      MaxLength       =   1
      TabIndex        =   16
      Top             =   210
      Width           =   420
   End
   Begin VB.TextBox txtCustNo 
      Height          =   270
      Index           =   1
      Left            =   1305
      MaxLength       =   9
      TabIndex        =   2
      Top             =   1140
      Width           =   1095
   End
   Begin VB.TextBox txtSalesNo 
      Height          =   270
      Left            =   1305
      MaxLength       =   9
      TabIndex        =   0
      Top             =   525
      Width           =   1095
   End
   Begin VB.TextBox txtCustNo 
      Height          =   270
      Index           =   0
      Left            =   1305
      MaxLength       =   9
      TabIndex        =   1
      Top             =   825
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Word編輯(&W)"
      Height          =   400
      Index           =   3
      Left            =   5175
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "整批列印(&F)"
      Height          =   400
      Index           =   4
      Left            =   6525
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "查詢(&F)"
      Default         =   -1  'True
      Height          =   400
      Index           =   2
      Left            =   7605
      TabIndex        =   3
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確認不寄(&O)"
      Height          =   400
      Index           =   1
      Left            =   7875
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   8415
      TabIndex        =   4
      Top             =   90
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm04010702.frx":0000
      Height          =   3495
      Left            =   90
      TabIndex        =   8
      Top             =   1920
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   20
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "C00"
         Caption         =   "不寄"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "C01"
         Caption         =   "已寄"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "C11"
         Caption         =   "已回"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "C02"
         Caption         =   "業務區"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "C03"
         Caption         =   "業務員"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "C04"
         Caption         =   "客戶代號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "C05"
         Caption         =   "客戶名稱"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "C06"
         Caption         =   "接洽人"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "C07"
         Caption         =   "電話"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "C08"
         Caption         =   "傳真"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "C09"
         Caption         =   "狀態"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "C12"
         Caption         =   "本所案號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2789.858
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1484.787
         EndProperty
         BeginProperty Column11 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8010
      Top             =   1020
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "地址條印表機："
      Height          =   180
      Index           =   1
      Left            =   315
      TabIndex        =   20
      Top             =   1545
      Width           =   1260
   End
   Begin VB.Label lblCount 
      Height          =   180
      Left            =   8055
      TabIndex        =   18
      Top             =   5490
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "筆數:"
      Height          =   180
      Left            =   7470
      TabIndex        =   17
      Top             =   5490
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "批號:                          ( A,B,... )"
      Height          =   180
      Left            =   315
      TabIndex        =   15
      Top             =   240
      Width           =   2250
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "迄:"
      Height          =   180
      Left            =   1035
      TabIndex        =   14
      Top             =   1170
      Width           =   225
   End
   Begin VB.Label lblCustName 
      Height          =   180
      Index           =   1
      Left            =   2475
      TabIndex        =   13
      Top             =   1185
      Width           =   5175
   End
   Begin VB.Label lblSalesName 
      Height          =   180
      Left            =   2475
      TabIndex        =   12
      Top             =   570
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "業務員:"
      Height          =   180
      Index           =   0
      Left            =   315
      TabIndex        =   11
      Top             =   570
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號起:"
      Height          =   180
      Left            =   315
      TabIndex        =   10
      Top             =   870
      Width           =   945
   End
   Begin VB.Label lblCustName 
      Height          =   180
      Index           =   0
      Left            =   2475
      TabIndex        =   9
      Top             =   870
      Width           =   5175
   End
End
Attribute VB_Name = "frm04010702"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan2010/8/11 日期欄已修改
'Add by Morgan 2007/6/30
Option Explicit
Dim bolBatchRight As Boolean '批次列印權限
'儲存印表機設定
Dim m_PrinterIdx As Integer, m_PrinterOrient As Integer

Private Sub cmdok_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 0 '結束
         Unload Me
         
      Case 1 '確定
         FormSave
         
      Case 2 '尋找
         AdodcRefresh
         
      Case 3 'Word編輯
         PrintLetter
         
      Case 4 '整批列印
         PrintLetter True
      '2007/7/10 ADD BY SONIA
      Case 5 '全部取消
         ClearAll
         
      Case 6 '全部選取
         SelectAll
      '2007/7/10 end
   End Select
   Screen.MousePointer = vbDefault
End Sub
'2007/7/10 ADD BY SONIA
Private Sub ClearAll()
Dim iRecs As Integer

   If DataGrid1.Enabled = True Then
      With Adodc1.Recordset
      Screen.MousePointer = vbHourglass
      DataGrid1.Visible = False
         .MoveFirst
         iRecs = 0
         Do While Not .EOF
            .Fields("C00") = "□"
            .MoveNext
         Loop
         .MoveFirst
      End With
      Screen.MousePointer = vbDefault
      DataGrid1.Visible = True
   End If
End Sub

Private Sub SelectAll()
Dim iRecs As Integer

   If DataGrid1.Enabled = True Then
      With Adodc1.Recordset
      Screen.MousePointer = vbHourglass
      DataGrid1.Visible = False
         .MoveFirst
         iRecs = 0
         Do While Not .EOF
            .Fields("C00") = "■"
            .MoveNext
         Loop
         .MoveFirst
     End With
      Screen.MousePointer = vbDefault
      DataGrid1.Visible = True
   End If
End Sub
'2007/7/10 end
Private Sub DataGrid1_DblClick()
   If DataGrid1.row >= 0 And DataGrid1.col = 0 Then
      'Debug.Print DataGrid1.Col
      If Adodc1.Recordset.Fields("C01") = "□" Then
         If "" & Adodc1.Recordset.Fields("C00") = "□" Then
            Adodc1.Recordset.Fields("C00") = "■"
         Else
            Adodc1.Recordset.Fields("C00") = "□"
         End If
         Adodc1.Recordset.UPDATE
      End If
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtBatchNo = "B"
   CmdEnable False
   If CheckUse("frm040702", strPrint, False) = True Then
      bolBatchRight = True
   Else
      bolBatchRight = False
   End If
   ' 暫存預設印表機
   m_PrinterOrient = Printer.Orientation
   SetPrinter m_PrinterIdx, cmbPrinter, Me
   '記錄原設定值
   cmbPrinter.Tag = cmbPrinter.Text
   '初始化序號
   pub_AddressListSN = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   CheckAddressList
   Set frm04010702 = Nothing
End Sub

Private Sub CheckAddressList()
   '列印地址條
   PUB_PrintAddressList strUserNum, cmbPrinter
   '刪除地址條列表資料
   PUB_DeleteAddressList strUserNum
   '初始化序號
   pub_AddressListSN = 0
   '若印表機變動, 則更新列印設定
   If cmbPrinter.Text <> cmbPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, cmbPrinter.Name, "0", "0", cmbPrinter.Text
   End If
   Set Printer = Printers(m_PrinterIdx)
   Printer.Orientation = m_PrinterOrient
End Sub

' 設定印表機
Private Sub SetPrinter(ByRef p_PrinterIdx As Integer, ByRef p_cboPrinter As ComboBox, ByRef p_Form As Form)
   
   p_cboPrinter.Tag = ""
   For intI = 0 To Printers.Count - 1
      p_cboPrinter.AddItem Printers(intI).DeviceName
      If Printers(intI).DeviceName = Printer.DeviceName Then
         p_PrinterIdx = intI
      End If
   Next
   If p_cboPrinter.ListCount > 0 Then: p_cboPrinter.ListIndex = 0
   '設定前次列印印表機
   strExc(0) = "Select * From PrintStartPoint Where PSP01='" & strUserNum & "' And PSP02='" & p_Form.Name & "' And PSP03='" & p_cboPrinter.Name & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If p_cboPrinter.ListCount > 0 Then
           For intI = 0 To p_cboPrinter.ListCount - 1
               If p_cboPrinter.List(intI) = "" & RsTemp("PSP06").Value Then
                   p_cboPrinter.ListIndex = intI
                   Exit For
               End If
           Next
       End If
   End If
End Sub


Private Sub CmdEnable(p_Enable As Boolean)
   cmdOK(1).Enabled = p_Enable
   cmdOK(3).Enabled = p_Enable
   '只有第B批可以整批列印
   If txtBatchNo = "B" Then
      If bolBatchRight = False Then
         cmdOK(4).Enabled = False
      Else
         cmdOK(4).Enabled = p_Enable
      End If
   Else
      cmdOK(4).Enabled = False
   End If
End Sub

Private Sub AdodcRefresh()
   
   Dim strCon As String
    
   CmdEnable False
   

      
   If lblSalesName = "" Then
      txtSalesNo_Validate False
   End If
   If lblCustName(0) = "" Then
      txtCustNo_Validate 0, False
   End If
   If lblCustName(1) = "" Then
      txtCustNo_Validate 1, False
   End If
   
   strCon = ""
   
   Select Case txtBatchNo
   
      Case "B"
         If txtSalesNo <> "" Then
            strCon = strCon & " and LL03='" & txtSalesNo & "'"
         End If
         If txtCustNo(0) <> "" Then
            strCon = strCon & " and LL01>='" & txtCustNo(0) & "'"
         End If
         If txtCustNo(1) <> "" Then
            strCon = strCon & " and LL01<='" & txtCustNo(1) & "'"
         End If
         
         strExc(0) = "select DECODE(LL04,'N','■','□') C00,DECODE(LR01,NULL,'□','■') C01,LL02 C02,ST02 C03,LL01 C04" & _
            ",CU04 C05,CU08 C06,CU16 C07,CU18 C08,CU80 C09,DECODE(LL04,'N','■','□') C10,DECODE(lR08,null,'□','■') C11" & _
            ",LR02||'-'||LR03||'-'||LR04||'-'||LR05 C12 From LinInfoList,Customer,STAFF,LinReasignRec " & _
            " where cu01(+)=substr(LL01,1,8) and cu02(+)=substr(LL01,9,1) AND ST01(+)=LL03 and lr01(+)=LL01" & strCon & _
            " order by LL02,LL03,LL01"
            
      Case Else
         If txtSalesNo <> "" Then
            strCon = strCon & " and CU13='" & txtSalesNo & "'"
         End If
         If txtCustNo(0) <> "" Then
            strCon = strCon & " and LR01>='" & txtCustNo(0) & "'"
         End If
         If txtCustNo(1) <> "" Then
            strCon = strCon & " and LR01<='" & txtCustNo(1) & "'"
         End If
         If txtBatchNo = "" Then
            strCon = strCon & " and LR11 is null"
         Else
            strCon = strCon & " and SUBSTR(LR11,1,1)='A'"
         End If
         strExc(0) = "select '□' C00,DECODE(LR01,NULL,'□','■') C01,CU12 C02,ST02 C03,LR01 C04" & _
            ",CU04 C05,CU08 C06,CU16 C07,CU18 C08,CU80 C09,'□' C10,DECODE(lR08,null,'□','■') C11" & _
            ",LR02||'-'||LR03||'-'||LR04||'-'||LR05 C12 From LinReasignRec,Customer,STAFF " & _
            " where cu01(+)=substr(LR01,1,8) and cu02(+)=substr(LR01,9,1) AND ST01(+)=CU13 " & strCon & _
            " order by CU12,CU13,LR01"
      
   End Select
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Set Adodc1.Recordset = PUB_CreateRecordset(RsTemp)
   If intI = 1 Then
      DataGrid1.Enabled = True
      CmdEnable True
   Else
      DataGrid1.Enabled = False
      MsgBox "查無資料！"
   End If
   lblCount = Adodc1.Recordset.RecordCount
   
End Sub

Private Sub txtBatchNo_GotFocus()
   TextInverse txtBatchNo
End Sub

Private Sub txtBatchNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("A") And KeyAscii <> Asc("B") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCustNo_Change(Index As Integer)
   lblCustName(Index) = ""
End Sub

Private Sub txtCustNo_GotFocus(Index As Integer)
   TextInverse txtCustNo(Index)
End Sub

Private Sub txtCustNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCustNo_Validate(Index As Integer, Cancel As Boolean)
   If txtCustNo(Index) <> "" Then
      If Len(txtCustNo(Index)) = 6 Or Len(txtCustNo(Index)) = 9 Then
         txtCustNo(Index) = Left(txtCustNo(Index) & "000", 9)
         lblCustName(Index) = GetCustomerName(txtCustNo(Index))
         If Index = 0 Then
            txtCustNo(1) = txtCustNo(Index)
            lblCustName(1) = lblCustName(Index)
         End If
      End If
   End If
End Sub

Private Sub txtSalesNo_Change()
   lblSalesName = ""
End Sub

Private Sub txtSalesNo_GotFocus()
   TextInverse txtSalesNo
End Sub

Private Sub txtSalesNo_Validate(Cancel As Boolean)
   If txtSalesNo <> "" Then
      If IsNumeric(txtSalesNo) Then
         lblSalesName = GetStaffName(txtSalesNo, True)
      Else
         setSalesNo
      End If
      If lblSalesName = "" Then
         Cancel = True
      End If
   End If
End Sub

Private Function setSalesNo() As String
   txtSalesNo.Tag = Trim(txtSalesNo)
   strExc(0) = "select st01 from staff where st02='" & txtSalesNo.Tag & "' order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      txtSalesNo = RsTemp(0)
      lblSalesName = txtSalesNo.Tag
   End If
End Function

Private Function FormSave() As Boolean

   Dim stSign As String, iRecs As Integer
   
   If DataGrid1.Enabled = True Then
      With Adodc1.Recordset
         
         cnnConnection.BeginTrans
         
On Error GoTo ErrHnd

         DataGrid1.Visible = False
         
         .MoveFirst
         iRecs = 0
         Do While Not .EOF
            If .Fields("C00") <> .Fields("C10") Then
               If .Fields("C00") = "■" Then
                  stSign = "N"
               Else
                  stSign = ""
               End If
               strSql = "Update LinInfoList set LL04='" & stSign & "', LL05='" & strUserNum & "',LL06=" & strSrvDate(1) & ",LL07=to_char(sysdate,'HH24MISS') where LL01='" & .Fields("C04") & "'"
               cnnConnection.Execute strSql, intI
               iRecs = iRecs + 1
               .Fields("C10") = .Fields("C00")
            End If
            .MoveNext
         Loop
         
         cnnConnection.CommitTrans
         Adodc1.Recordset.UpdateBatch
         .MoveFirst
      End With
   End If
   DataGrid1.Visible = True
   
   FormSave = True
   If iRecs = 0 Then
      MsgBox "本次並無資料需更新！"
   Else
      MsgBox "更新完成，共 " & iRecs & " 筆！"
   End If
         
   Exit Function
   
ErrHnd:
   DataGrid1.Visible = True
   If Err.NUMBER <> 0 Then
      adoTaie.RollbackTrans
      MsgBox Err.Description
   End If
   
End Function

'列印通知函
Private Function InfoLetterPrint(p_CustNo As String, Optional p_Edit As Boolean = False) As Boolean

   Dim stCP09 As String, stSNo As String, stCustNo As String, ii As Integer
   
   strExc(0) = "select CP09,LR11,PA26,PA27,PA28,PA29,PA30 from linreasignrec,CASEPROGRESS,patent where lr01='" & p_CustNo & "'" & _
      " AND CP01(+)=LR02 AND CP02(+)=LR03 AND CP03(+)=LR04 AND CP04(+)=LR05 AND CP09<'C' AND CP27>0" & _
      " AND PA01(+)=LR02 AND PA02(+)=LR03 AND PA03(+)=LR04 AND PA04(+)=LR05 ORDER BY CP27 DESC,CP09 DESC"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      stCP09 = RsTemp.Fields(0)
      stSNo = "" & RsTemp.Fields(1)
      stCustNo = "" & RsTemp.Fields(2)
      For ii = 3 To 6
         If Not IsNull(RsTemp.Fields(ii)) Then
            stCustNo = stCustNo & "," & RsTemp.Fields(ii)
         End If
      Next
      EndLetter "02", stCP09, "98", strUserNum
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('02','" & stCP09 & "','98','" & strUserNum & "','客戶編號','(" & stCustNo & ")')"
      cnnConnection.Execute strSql, intI
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('02','" & stCP09 & "','98','" & strUserNum & "','類別','" & Left(stSNo, 1) & "')"
      cnnConnection.Execute strSql, intI
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('02','" & stCP09 & "','98','" & strUserNum & "','流水號','" & Mid(stSNo, 2) & "')"
      cnnConnection.Execute strSql, intI

      NowPrint stCP09, "02", "98", p_Edit, strUserNum, 0, , , , 1
      
      'Add by Morgan 2007/7/4
      If p_Edit = False Then
         pub_AddressListSN = pub_AddressListSN + 1
         strSql = "Insert Into AddressList (AL01,AL02,AL03,AL04,AL05,AL06) select '" & strUserNum & "',cp01,cp02,cp03,cp04," & pub_AddressListSN & " from caseprogress where cp09='" & stCP09 & "'"
         cnnConnection.Execute strSql, intI
      End If
      'end 2007/7/4
   End If
End Function

'新增通知紀錄
Private Function InfoRecAdd(p_CustNo As String, Optional p_Info As Boolean) As Boolean

   Dim stSNo As String
   
On Error GoTo ErrHnd

   strExc(0) = "select LC01,LC02,LC03,LC04,LC05,LC07,LC08,LC09,LC10,LC11 FROM LINCASE WHERE LC07='" & p_CustNo & "'" & _
      " UNION select LC01,LC02,LC03,LC04,LC05,LC07,LC08,LC09,LC10,LC11 FROM LINCASE WHERE LC08='" & p_CustNo & "'" & _
      " UNION select LC01,LC02,LC03,LC04,LC05,LC07,LC08,LC09,LC10,LC11 FROM LINCASE WHERE LC09='" & p_CustNo & "'" & _
      " UNION select LC01,LC02,LC03,LC04,LC05,LC07,LC08,LC09,LC10,LC11 FROM LINCASE WHERE LC10='" & p_CustNo & "'" & _
      " UNION select LC01,LC02,LC03,LC04,LC05,LC07,LC08,LC09,LC10,LC11 FROM LINCASE WHERE LC11='" & p_CustNo & "'" & _
      " ORDER BY LC07,LC08,LC09,LC10,LC11,LC05 DESC"
      
   intI = 1
   Set adoRecordset1 = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoRecordset1
         'Modify by Morgan 2007/7/16 加檢查若為多申請人案件時只印一次
         strExc(0) = "select LR11 from linreasignrec" & _
            " where lr02='" & .Fields("LC01") & "' and lr03='" & .Fields("LC02") & "' and lr04='" & .Fields("LC03") & "' and lr05='" & .Fields("LC04") & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            stSNo = RsTemp.Fields(0)
            p_Info = False
         Else
            stSNo = GetNextNumber(txtBatchNo)
            p_Info = True
            Debug.Print stSNo
         End If
         strSql = "insert into LinReAsignRec(LR01,LR02,LR03,LR04,LR05,LR06,LR07,LR11)" & _
            " VALUES('" & p_CustNo & "','" & .Fields("LC01") & "' ,'" & .Fields("LC02") & "'" & _
            ",'" & .Fields("LC03") & "','" & .Fields("LC04") & "'," & strSrvDate(1) & ",'" & strUserNum & "','" & stSNo & "')"
         cnnConnection.Execute strSql, intI
         Adodc1.Recordset.Fields("C01") = "■"
         Adodc1.Recordset.Fields("C12") = "" & .Fields("LC01") & "-" & .Fields("LC02") & "-" & .Fields("LC03") & "-" & .Fields("LC04")
         Adodc1.Recordset.UPDATE
      End With
      InfoRecAdd = True
   Else
      MsgBox "無法讀取相關資料！", vbExclamation
   End If
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description
   End If
End Function
Private Sub PrintLetter(Optional p_All As Boolean = False)
   
   Dim stCP09 As String, stCustNo As String, stSNo As String, MsgResult As VbMsgBoxResult
   Dim iCount As Integer, bInfo As Boolean
   
   If DataGrid1.Enabled = True Then
      '單筆
      If p_All = False Then
         If Not (Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF) Then
            '檢查是否不寄
            If Adodc1.Recordset.Fields("C00") = "□" Then
               stCustNo = Adodc1.Recordset.Fields("C04")
               '檢查是否已寄
               If Adodc1.Recordset.Fields("C01") = "□" Then
                  If MsgBox("客戶代碼【" & stCustNo & "】尚未有通知紀錄，是否要新增？", vbYesNo + vbDefaultButton2) = vbNo Then
                     Exit Sub
                  Else
                     '新增通知紀錄
                     If InfoRecAdd(stCustNo) = False Then
                        Exit Sub
                     End If
                  End If
               End If
               '產生通知函
               InfoLetterPrint stCustNo, True
            Else
               MsgBox "本客戶已設定為【不寄】！", vbExclamation
               Exit Sub
            End If
         End If
      '整批
      Else
         '已更新檢查
         If UpdateCheck = False Then
            Exit Sub
         End If
         MsgResult = MsgBox("是否包含【已寄】客戶？", vbYesNo + vbDefaultButton2)
         
         With Adodc1.Recordset
            DataGrid1.Visible = False
            .MoveFirst
            Do While Not .EOF
               '不寄和已回覆的都不印
               If Adodc1.Recordset.Fields("C00") = "□" And Adodc1.Recordset.Fields("C11") = "□" Then
                  stCustNo = Adodc1.Recordset.Fields("C04")
                  bInfo = False
                  '未寄
                  If Adodc1.Recordset.Fields("C01") = "□" Then
                     '新增通知紀錄
                     If InfoRecAdd(stCustNo, bInfo) = False Then
                        If MsgBox("新增【" & stCustNo & "】客戶的通知紀錄失敗，是否要繼續？", vbYesNo + vbDefaultButton1) = vbNo Then
                           Exit Do
                        End If
                     End If
                  
                  '已寄
                  ElseIf MsgResult = vbNo Then
                     GoTo NextCust
                  Else
                     bInfo = True
                  End If

                  If bInfo = True Then
                     '產生通知函
                     InfoLetterPrint stCustNo
                     iCount = iCount + 1
                     Debug.Print iCount
                  End If
                  
               End If
NextCust:
               .MoveNext
            Loop
            .MoveFirst
            DataGrid1.Visible = True
            If iCount > 0 Then
               If MsgBox("共產生 " & iCount & " 筆通知函，是否現在要整批印出？", vbYesNo + vbDefaultButton1) = vbYes Then
                  PUB_BatchPrint "1"
               End If
            End If
         End With
         
      End If
   End If
   
End Sub
'檢查是否資料尚未更新
Private Function UpdateCheck() As Boolean
   Dim iRlt As Integer
   UpdateCheck = True
   With Adodc1.Recordset
      DataGrid1.Visible = False
      .MoveFirst
      Do While Not .EOF
         If .Fields("C00") <> .Fields("C10") Then
            MsgBox "客戶不寄欄位已有更動但尚未更新，請按【確認不寄】更新資料或按【查詢】重新查詢！"
            UpdateCheck = False
            Exit Do
         End If
         .MoveNext
      Loop
      DataGrid1.Visible = True
   End With
   
End Function
Private Function GetNextNumber(p_Lead As String) As String
   strExc(0) = "select max(SUBSTR(lr11,2)) from linreasignrec where substr(lr11,1,1)='" & p_Lead & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetNextNumber = p_Lead & Format(Val("" & RsTemp.Fields(0)) + 1, "000000")
   Else
      GetNextNumber = p_Lead & "000001"
   End If
   
End Function


Private Sub cmdTest_Click()
   Dim stCP09 As String
   Dim adoRst As New ADODB.Recordset
   Dim iCol As Integer, strCusNum As String
   
   
   strExc(0) = "select lc07,lc08,lc09,lc10,lc11,c01" & _
      " from (select lc07,lc08,lc09,lc10,lc11,max(lc05) c01 from lincase where lc06='1' and lc07 is not null" & _
      " group by lc07,lc08,lc09,lc10,lc11) X where not exists(select * from linreasignrec where lr01=lc07)" & _
      " order by lc07"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      pub_AddressListSN = 0
      With adoRst
      Do While Not .EOF
         strExc(0) = "select * from linreasignrec where lr01='" & .Fields(0) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
             pub_AddressListSN = pub_AddressListSN + 1
             stCP09 = Mid(.Fields("c01"), 9, 9)
             strCusNum = .Fields(0)
             For iCol = 1 To 4
                If Not IsNull(.Fields(iCol)) Then
                   strCusNum = strCusNum & "," & .Fields(iCol)
                End If
             Next
'             EndLetter "02", stCP09, "98", strUserNum
'             strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                   "VALUES ('02','" & stCP09 & "','98','" & strUserNum & "','客戶編號','(" & strCusNum & ")')"
'             cnnConnection.Execute strSQL, intI
'             strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                   "VALUES ('02','" & stCP09 & "','98','" & strUserNum & "','類別','B')"
'             cnnConnection.Execute strSQL, intI
'             strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                   "VALUES ('02','" & stCP09 & "','98','" & strUserNum & "','流水號','" & Format(pub_AddressListSN, "000000") & "')"
'             cnnConnection.Execute strSQL, intI
'
'             NowPrint stCP09, "02", "98", False, strUserNum, 0, , , , 1
             
'             strSQL = "Insert Into AddressList (AL01,AL02,AL03,AL04,AL05,AL06) select '" & strUserNum & "',cp01,cp02,cp03,cp04," & pub_AddressListSN & " from caseprogress where cp09='" & stCP09 & "'"
'             cnnConnection.Execute strSQL, intI
             
             strSql = "insert into LinReAsignRec(LR01,LR02,LR03,LR04,LR05,LR06,LR07,LR11) select '" & .Fields(0) & "',cp01,cp02,cp03,cp04," & strSrvDate(1) & ",'" & strUserNum & "','B" & Format(pub_AddressListSN, "000000") & "' from caseprogress where cp09='" & stCP09 & "'"
             cnnConnection.Execute strSql, intI
             
             For iCol = 1 To 4
                If Not IsNull(.Fields(iCol)) Then
                  strExc(0) = "select * from linreasignrec where lr01='" & .Fields(iCol) & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 0 Then
                     strSql = "insert into LinReAsignRec(LR01,LR02,LR03,LR04,LR05,LR06,LR07,LR11) select '" & .Fields(iCol) & "',cp01,cp02,cp03,cp04," & strSrvDate(1) & ",'" & strUserNum & "','B" & Format(pub_AddressListSN, "000000") & "' from caseprogress where cp09='" & stCP09 & "'"
                     cnnConnection.Execute strSql, intI
                  End If
                End If
             Next
             
         End If
         .MoveNext
      Loop
      End With
   End If
End Sub
