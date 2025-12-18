VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040104_i 
   BorderStyle     =   1  '單線固定
   Caption         =   "內專領證年費整批發文"
   ClientHeight    =   3936
   ClientLeft      =   -2676
   ClientTop       =   1572
   ClientWidth     =   9204
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3936
   ScaleWidth      =   9204
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消不發文◇"
      Height          =   564
      Index           =   5
      Left            =   6192
      TabIndex        =   11
      Top             =   396
      Width           =   825
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "電子送件申請書◆"
      Height          =   564
      Index           =   4
      Left            =   5196
      TabIndex        =   10
      Top             =   396
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "進度"
      Height          =   564
      Index           =   0
      Left            =   4500
      TabIndex        =   9
      Top             =   396
      Width           =   675
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印"
      CausesValidation=   0   'False
      Height          =   564
      Index           =   3
      Left            =   7044
      TabIndex        =   8
      Top             =   396
      Width           =   675
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   585
      Top             =   1740
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   572
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "查詢"
      CausesValidation=   0   'False
      Height          =   564
      Index           =   0
      Left            =   3810
      TabIndex        =   6
      Top             =   396
      Width           =   675
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   555
      Left            =   996
      TabIndex        =   2
      Top             =   432
      Width           =   2760
      Begin VB.OptionButton Option2 
         Caption         =   "台灣"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   60
         Value           =   -1  'True
         Width           =   2580
      End
      Begin VB.OptionButton Option2 
         Caption         =   "非台灣(大陸,香港,澳門,PCT...)"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   300
         Width           =   2800
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   564
      Index           =   2
      Left            =   8436
      TabIndex        =   1
      Top             =   396
      Width           =   675
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "發文"
      CausesValidation=   0   'False
      Height          =   564
      Index           =   1
      Left            =   7740
      TabIndex        =   0
      Top             =   396
      Width           =   675
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm040104_i.frx":0000
      Height          =   2832
      Left            =   132
      TabIndex        =   7
      Top             =   1032
      Width           =   8916
      _ExtentX        =   15727
      _ExtentY        =   4995
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "C00"
         Caption         =   "收文日"
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
      BeginProperty Column02 
         DataField       =   "C02"
         Caption         =   "案件性質"
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
         DataField       =   "C03"
         Caption         =   "本所期限"
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
      BeginProperty Column04 
         DataField       =   "C04"
         Caption         =   "送件方式"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "C05"
         Caption         =   "指定送件日"
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
         DataField       =   "C06"
         Caption         =   "繳費年度"
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
         DataField       =   "C07"
         Caption         =   "智權人員"
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
         Size            =   275
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   756.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   875.906
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1031.811
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   852.095
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "承辦人："
      Height          =   240
      Left            =   96
      TabIndex        =   13
      Top             =   108
      Visible         =   0   'False
      Width           =   948
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   288
      Left            =   1032
      TabIndex        =   12
      Top             =   72
      Visible         =   0   'False
      Width           =   1500
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2646;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   240
      Left            =   96
      TabIndex        =   5
      Top             =   504
      Width           =   900
   End
End
Attribute VB_Name = "frm040104_i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/14 改成Form2.0 (DataGrid1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Create by Morgan 2010/12/13
Option Explicit

Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim m_iCols As Integer, m_iLstOption As Integer
Dim m_lstCP09 As String 'Added by Morgan 2020/4/9
Dim pa(4) As String


Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   'Added by Morgan 2022/11/3
   cmdOK(5).Enabled = False
   'Modified by Morgan 2025/1/10
   'If Not Adodc1.Recordset.EOF Then
   If Not Adodc1.Recordset.EOF And Not Adodc1.Recordset.BOF Then
   'end 2025/1/10
      If InStr(Adodc1.Recordset.Fields("C01"), "◇") > 0 Then
         cmdOK(5).Enabled = True
      End If
   End If
End Sub


Private Sub cmdok_Click(Index As Integer)
   Dim bCancel As Boolean
   
   Select Case Index
      Case 0
         Screen.MousePointer = vbHourglass
         ReadData
         Screen.MousePointer = vbDefault
         
      Case 1
         
         'If DoPrint Then 'Removed by Morgan 2025/2/14 改程序改分區負責後每次處理的案件數較少不必再自動印清單--郭/玲玲
         
            'Added by Morgan 2012/6/14
            '切換Word印表機為程式預設印表機
            pub_OsPrinter = PUB_GetOsDefaultPrinter
            PUB_SetOsDefaultPrinter Printer.DeviceName
            PUB_SetWordActivePrinter
            'end 2012/6/14
            doBatch
            PUB_SetOsDefaultPrinter pub_OsPrinter 'Added by Morgan 2012/6/14
            
            If Adodc1.Recordset.RecordCount = 0 Then
               cmdOK(1).Enabled = False
               cmdOK(3).Enabled = False
            End If
            MsgBox "作業結束！"
            
         'End If 'Removed by Morgan 2025/2/14

      Case 2 '離開
         Unload Me
      
      Case 3
         If DoPrint Then MsgBox "列印完成！"
      
      'Added by Morgan 2020/4/7
      Case 4 '申請書
         With Adodc1.Recordset
         If "" & .Fields("cp118") = "" Then
            MsgBox "本程序未設定電子送件！", vbCritical
         Else
            'Added by Morgan 2021/5/6
            bCancel = False
            pa(1) = .Fields("cp01")
            pa(2) = .Fields("cp02")
            pa(3) = .Fields("cp03")
            pa(4) = .Fields("cp04")
            If ChkPrintOnly(pa, strExc(1)) = True Then
               If MsgBox("本案尚有 " & strExc(1) & " 未發文，是否要繼續？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                  bCancel = True
               End If
            End If
            If bCancel = False Then
            'end 2021/5/6
            
               m_lstCP09 = "" & .Fields("cp09")
               frm040103_1.stCP09 = "" & .Fields("cp09")
               frm040103_1.Text1 = "" & .Fields("cp01")
               frm040103_1.Text2 = "" & .Fields("cp02")
               frm040103_1.Text3 = "" & .Fields("cp03")
               frm040103_1.Text4 = "" & .Fields("cp04")
               frm040103_1.Show
               
            End If 'Added by Morgan 2021/5/6
         End If
         End With
         
      'Added by Morgan 2022/11/2
      Case 5
         Adodc1.Recordset.Fields("C01") = Replace(Adodc1.Recordset.Fields("C01"), "◇", "　")
         Adodc1.Recordset.UpdateBatch
   End Select
End Sub

Private Sub Combo1_Click()
   If Combo1.Visible = False Then Exit Sub
   If Combo1.Tag <> Combo1 Then
      cmdOK(0).Value = True
   End If
End Sub

'Added by Morgan 2020/1/7
Private Sub Command1_Click(Index As Integer)
   If Index = 0 Then
      Me.Enabled = False
      If fnSaveParentForm(Me) = False Then
         Me.Enabled = True
         Exit Sub
      End If
      Screen.MousePointer = vbHourglass
      frm100101_2.Show
      'Modified by Morgan 2020/4/24
      'frm100101_2.Tag = Pub_RplStr(Adodc1.Recordset.Fields(1))
      frm100101_2.Tag = Pub_RplStr(Mid(Adodc1.Recordset.Fields(1), 3))
      frm100101_2.cmdOK(5).Visible = False '下一筆按鈕隱藏
      frm100101_2.StrMenu
      Screen.MousePointer = vbDefault
      Me.Enabled = True
   End If
End Sub

Private Sub Form_Activate()
   'Added by Morgan 2020/4/9
   If m_lstCP09 <> "" Then
      cmdOK(0).Value = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   cmdOK(1).Enabled = False
   cmdOK(3).Enabled = False
   cmdOK(5).Enabled = False 'Added by Morgan 2022/11/3
   Command1(0).Enabled = False 'Added by Morgan 2020/1/7
   
   'Added by Morgan 2025/1/9
   If strSrvDate(1) >= P業務區劃分啟用日 Then
      Combo1.Visible = True
      Label1.Visible = True
      Call SetPatentP12Combo(Combo1, "P", Label1)
      cmdOK(0).Value = True
   End If
   'end 2025/1/9
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040104_i = Nothing
End Sub

Private Sub ReadData()
   Dim stCon As String
        
   cmdOK(1).Enabled = False
   cmdOK(3).Enabled = False
   cmdOK(4).Enabled = False 'Added by Morgan 2020/4/7
   cmdOK(5).Enabled = False 'Added by Morgan 2022/11/3
   Command1(0).Enabled = False 'Added by Morgan 2020/1/7
   
   Combo1.Tag = Combo1 'Added by Morgan 2025/1/10
      
   If Me.Option2(0).Value = True Then
      'Modified by Morgan 2020/4/17 台灣改都用電子送件，非自動收文也要列出--玲玲
      'stCon = " and pa09='000' and cp140 is not null"
      stCon = " and pa09='000' and cp14 is not null and cp53 is not null and cp54 is not null"
      'end 2020/4/17
   Else
      'Added by Morgan 2019/12/20 取得排除寰華案語法
      FMP2openSQL = ""
      FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05, "INVERSE_SQL")
      'end 2019/12/20
      
      'Modify by Morgan 2011/8/17
      '改非自動收文也要能整批發,但要控制已分案且有輸入年度且無其他未發文程序
      'stCon = " and pa09<>'000' and cp140 is not null"
      'Modified by Morgan 2012/3/23 +FMP 領證要排除--玲玲
      'Modified by Morgan 2019/12/20 要排除外專寰華案--玲玲
      'Modified by Morgan 2020/5/20 +FMP 領證--韻丞  取消 and not(cp12 like 'F%' and cp10='601')
      'Modified by Morgan 2022/10/2 有其他未發文程序時也要帶出,但預設不整批發
      'stCon = " and pa09<>'000' and cp14 is not null and cp53 is not null and cp54 is not null" & _
         " and not exists(select * from caseprogress c2 where c2.cp01=f0.cp01" & _
         " and c2.cp02=f0.cp02 and c2.cp03=f0.cp03 and c2.cp04=f0.cp04" & _
         " and c2.cp09<>f0.cp09 and c2.cp27||c2.cp57 is null)" & FMP2openSQL
      stCon = " and pa09<>'000' and cp14 is not null and cp53 is not null and cp54 is not null" & FMP2openSQL
         
   End If
   
   'Added by Morgan 2025/1/10
   If Combo1 <> "" Then
      stCon = stCon & " and cp14='" & Left(Combo1, 5) & "'"
   End If
   'end 2025/1/10
   
   'Modified by Morgan 2013/11/19 +香港維持費(Ex.P-103291) --玲玲
   'Modified by Morgan 2018/4/24 +cp158=0 and cp159=0 (使語法改用idex提升效能)
   'Modified by Morgan 2020/4/8 +cp118
   'Modified by Morgan 2020/4/9 +cp160
   'Modified by Morgan 2020/12/17 收款後送件也要考慮開請款單時是否結清(+ or nvl(a1k29,'N')='Y') Ex:P-105153
   'Modified by Morgan 2021/1/13 收款後送件考慮分所已繳款情形改先列出後逐筆呼叫函數剔除 -and (cp79=0 or nvl(a1k29,'N')='Y')
   'Modified by Morgan 2023/8/31 +指定日之前、之後狀況(非台灣案也不必提前列出--玲玲)
   strExc(0) = "select SUBSTR(' '||sqldatet(CP05),-9) as C00" & _
      ",decode(cp140,'','紙','　')||decode(cp160,0,'　','◆')||cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as C01" & _
      ",decode(pa09,'000',cpm03,cpm04) C02" & _
      ",SUBSTR(' '||sqldatet(CP06),-9) as C03" & _
      ",decode(cp141,'1','立即','2','收款後','3','指定日期'||decode(cp164,'2','之前','3','之後')) as C04" & _
      ",SUBSTR(' '||sqldatet(CP142),-9) as C05" & _
      ",cp53||'-'||cp54 as C06" & _
      ",st02 C07,cp01,cp02,cp03,cp04,cp09,cp10,cp53,cp54,cp140,pa09,cp118,cp141,cp06" & _
      " from caseprogress f0,patent,casepropertymap,staff,acc1k0" & _
      " where cp158=0 and cp159=0 and  cp01='P' and cp09<'B' and cp10 in ('601','605','606')" & _
      " and cp27||cp57 is null" & _
      " and (cp06<=" & strSrvDate(1) & " or cp141 is null or cp141='1'" & _
      " or cp141='2' or (cp141='3' and (cp164='2' or (nvl(cp164,'1') in ('1','3') and nvl(cp142,0)<=" & strSrvDate(1) & "))))" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa57||pa108 is null" & stCon & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and st01(+)=cp13 and a1k01(+)=cp60 order by cp05,cp09"
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Amy 2014/06/09 +FormName 改暫存TB
   'Set Adodc1.Recordset = PUB_CreateRecordset(RsTemp, , , 50)
   Set Adodc1.Recordset = PUB_CreateRecordset(RsTemp, , , 50, Me.Name)
   If intI = 0 Then
      MsgBox "無符合資料!!"
   Else
   
      'Added by Morgan 2021/1/13
      '剔除收款後送件但尚有未收款之案件
      DataGrid1.Visible = False
      With Adodc1.Recordset
      .MoveFirst
      Do While Not .EOF
         'Added by Morgan 2022/11/2 有其他未發文程序時預設不整批發
         If Me.Option2(1).Value = True Then
            strExc(0) = "select * from caseprogress where cp01='" & .Fields("cp01") & "' and cp02='" & .Fields("cp02") & "'" & _
               "and cp03='" & .Fields("cp03") & "' and cp04='" & .Fields("cp04") & "' and cp09<>'" & .Fields("cp09") & "' and cp27||cp57 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               .Fields("C01") = Replace(.Fields("C01"), "　P", "◇P")
            End If
         End If
         'end 2022/11/2
         
         'Modified by Morgan 2021/2/20 已達所限不可剔除
         'If .Fields("cp141") = "2" Then
         If .Fields("cp141") = "2" And Not ("" & .Fields("cp06") <= strSrvDate(1)) Then
         'end 2021/2/20
            If PUB_ChkPaidByCP09(.Fields("cp09")) = False Then
               .Delete
            End If
         End If
         
         .MoveNext
      Loop
      .UpdateBatch
      End With
      
      DataGrid1.Visible = True
      If Adodc1.Recordset.RecordCount = 0 Then
         MsgBox "無符合資料!!"
      Else
         Adodc1.Recordset.MoveFirst
         DataGrid1.row = 0 'Added by Morgan 2025/1/10 若不指定，前面有刪除資料時可能會指到第2筆
      'end 2021/1/13
      
         cmdOK(1).Enabled = True
         cmdOK(3).Enabled = True
         Command1(0).Enabled = True 'Added by Morgan 2020/1/7
         If Option2(0).Value = True Then cmdOK(4).Enabled = True 'Added by Morgan 2020/4/7
         
         'Added by Morgan 2020/4/9
         If m_lstCP09 <> "" Then
            Adodc1.Recordset.Find "cp09='" & m_lstCP09 & "'"
            If Adodc1.Recordset.EOF Then
               Adodc1.Recordset.MoveFirst
            End If
         End If
      End If 'Added by Morgan 2021/1/13
      'end 2020/4/9
   End If
   RecordShow
   m_lstCP09 = "" 'Added by Morgan 2020/4/9
End Sub

'Add by Morgan 2011/1/14
'檢查 CP 是否有特定案件性質的收文未發文
Private Function ChkPrintOnly(cp() As String, Optional ByRef pRefCPM As String) As Boolean
   Dim stSQL As String, iR As Integer
   
   stSQL = "select decode(pa09,'000',cpm03,cpm04) cp10C" & _
      " from caseprogress,patent,casepropertymap where cp01='" & cp(1) & "' and cp02='" & cp(2) & "'" & _
      " and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp09<'C' and cp27||cp57 is null" & _
      " and cp10 in ('401','412','421','701','702','703','704','705','706','707','708','919')" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10"
   iR = 1
   Set AdoRecordSet3 = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      ChkPrintOnly = True
      pRefCPM = AdoRecordSet3(0)
      If AdoRecordSet3.RecordCount > 1 Then
         pRefCPM = pRefCPM & "...等相關程序"
      End If
   End If
End Function

'Add by Morgan 2011/1/14
'檢查 NP 是否有相關期限
Private Function CheckRefNPExist(pCP09) As Boolean
   Dim stSQL As String, iR As Integer
   
   stSQL = "select 1 from nextprogress where np01='" & pCP09 & "'"
   iR = 1
   Set AdoRecordSet3 = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      CheckRefNPExist = True
   End If
   
End Function

Private Sub doBatch()
   Dim iRow As Integer
   Dim bSuccess As Boolean
   Dim stNoBatchList As String 'Added by Morgan 2020/3/18
   
   With Adodc1.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         If InStr("" & .Fields("C01"), "◇") = 0 Then 'Added by Morgan 2022/11/2 沒有不發文註記才印
            
            bSuccess = False
            
            pa(1) = .Fields("cp01")
            pa(2) = .Fields("cp02")
            pa(3) = .Fields("cp03")
            pa(4) = .Fields("cp04")
            If CheckRefNPExist(.Fields("cp09")) Then
               'Modified by Morgan 2014/6/9 +本所案號
               MsgBox pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & " 收文號【" & .Fields("cp09") & "】下一程序已有相關期限，不可發文!!!", vbExclamation + vbOKOnly
               
            'Add by Morgan 2011/1/14
            '若有其他相關程序未發文則只列印接洽單並帶出"非整批發文"字樣
            ElseIf ChkPrintOnly(pa, strExc(1)) = True Then
               bSuccess = True
               'Modified by Morgan 2021/5/6
               'stNoBatchList = stNoBatchList & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "(相關程序未發文)" & vbCrLf 'Added by Morgan 2020/3/18
               stNoBatchList = stNoBatchList & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "(" & strExc(1) & "未發文)" & vbCrLf 'Added by Morgan 2020/3/18
            'Added by Morgan 2020/4/8
            'Modified by Morgan 2024/1/30 +台灣(因增加大陸案收文會預設電子送件)
            ElseIf "" & .Fields("cp118") <> "" And .Fields("pa09") = "000" Then
               bSuccess = True
               stNoBatchList = stNoBatchList & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "(電子送件)" & vbCrLf 'Added by Morgan 2020/3/18
            'end 2020/4/8
            Else
               '領證
               If .Fields("cp10") = "601" Then
                  '基本檔核准檢查
                  If PUB_ApproveCheck(.Fields("cp09")) Then
                  
                     'Added by Morgan 2021/12/15
                     '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                     If PUB_CheckFormExist("frm040104_7") = False Then
                        Set frm040104_7 = Nothing
                     End If
                     'end 2021/12/15
                     
                     With frm040104_7
                        .m_bolBeCalled = True
                        .m_CP01 = pa(1)
                        .m_CP02 = pa(2)
                        .m_CP03 = pa(3)
                        .m_CP04 = pa(4)
                        .m_CP09 = Adodc1.Recordset.Fields("cp09")
                        'Remove by Morgan 2011/8/17 改發文預設
                        '.Text7(1) = Val("" & Adodc1.Recordset.Fields("cp54"))
                        .Text9 = strSrvDate(2)
                        .cmdOK(0).Enabled = False  'Added by Morgan 2017/12/6 先鎖住確定鈕以免執行中被觸發(P110168發生下一程序新增了兩筆年費)
                        bSuccess = .Process(0)
                     End With
                     DoEvents
                     Unload frm040104_7
                  End If
                  
               '年費
               'Modified by Morgan 2013/11/19 +香港維持費606(Ex.P-103291) --玲玲
               ElseIf .Fields("cp10") = "605" Or .Fields("cp10") = "606" Then
                  If .Fields("pa09") = "000" Then
                      If PUB_ChkCPExist(pa, 減免退費, 1) = True Then
                         If MsgBox("本案有【減免退費】未發文，若要同時發文請改選【減免退費】發文！確定只發文【年費】？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                            Exit Sub
                         End If
                      End If
                  End If
                  
                  'Added by Morgan 2021/12/15
                  '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                  If PUB_CheckFormExist("frm040104_a") = False Then
                     Set frm040104_a = Nothing
                  End If
                  'end 2021/12/15
                  
                  With frm040104_a
                     .m_bolBeCalled = True
                     .m_CP01 = pa(1)
                     .m_CP02 = pa(2)
                     .m_CP03 = pa(3)
                     .m_CP04 = pa(4)
                     .m_CP09 = Adodc1.Recordset.Fields("cp09")
                     '.Text5(0) = Val(Adodc1.Recordset.Fields("cp53"))
                     '.Text5(1) = Val(Adodc1.Recordset.Fields("cp54"))
                     .Text5(4) = strSrvDate(2)
                     .cmdOK(0).Enabled = False   'Added by Morgan 2017/12/6
                     bSuccess = .Process(0)
                  End With
                  DoEvents
                  Unload frm040104_a
                  
               End If
            End If
         
            If bSuccess Then            '
               'Modify by Morgan 2011/8/17 非自動收文不印接洽單
               'Modified by Morgan 2014/12/17 台灣案發文電子化後不必再印接洽單
               'If Not IsNull(.Fields("cp140")) Then
   'Modified by Morgan 2015/9/8 +非臺灣案(都不用印)
   '            If Not IsNull(.Fields("cp140")) And Not (P台灣案電子化啟用日 <= Val(strSrvDate(1)) And .Fields("pa09") = "000") Then
   '
   '               With frm090801
   '                  .txtPCnt = 1
   '                  .txtPrintType = 2
   '                  .Text5 = Adodc1.Recordset.Fields("cp140")
   '                  .m_blnCallPrint_CRL119 = False 'Add By Sindy 2014/2/7 不必列印特殊收據頁
   '                  .cmdOK_Click 4 '查詢
   '                  .m_blnCallPrint = True
   '                  .m_bolPrintMark = bolMark 'Add by Morgan 2011/1/14
   '                  .cmdOK_Click 0 '列印
   '               End With
   '               DoEvents
   '               Unload frm090801
   '            End If
   'end 2015/9/8
            Else
               Exit Do
            End If
            .Delete
         End If 'Added by Morgan 2022/11/2
         .MoveNext
      Loop
      .UpdateBatch
      
      'Added by Morgan 2020/3/18
      If stNoBatchList <> "" Then
         MsgBox "下列案號請改單筆發文！" & vbCrLf & vbCrLf & stNoBatchList, vbInformation
         cmdOK(0).Value = True
      End If
      'end 2020/3/18
   End If
   End With
End Sub

Private Function DoPrint() As Boolean

   Dim iOrientation As Integer, iRow As Integer, iCol As Integer, iCount As Integer
   Dim strTemp() As String
   
   iOrientation = Printer.Orientation
   'Printer.Orientation = 2
   Printer.Orientation = 1
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   m_iCols = 8
   
   With Adodc1.Recordset
      GetPleft
      ReDim strTemp(1 To m_iCols)
      iPage = 1
      iCount = 0
      PrintPageHeader
      PrintPageHeader1
      .MoveFirst
      Do While Not .EOF
         If InStr("" & .Fields("C01"), "◇") = 0 Then 'Added by Morgan 2022/11/2 沒有不發文註記才印
            iCount = iCount + 1
            For iCol = LBound(strTemp) To UBound(strTemp)
               strTemp(iCol) = "" & .Fields(iCol - 1)
            Next
            PrintDetail strTemp
         End If
         .MoveNext
      Loop
      Call PrintReportFooter(iCount)
      Printer.EndDoc
      DoPrint = True
   End With
   Printer.Orientation = iOrientation
End Function


'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = Adodc1.Recordset.RecordCount
End Sub


Sub GetPleft()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   intI = 8
   ReDim PLeft(1 To intI)
   PLeft(1) = ciStartX
   PLeft(2) = PLeft(1) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(3) = PLeft(2) + Printer.TextWidth(String(6, "　")) + ciColGap
   PLeft(4) = PLeft(3) + Printer.TextWidth(String(6, "　")) + ciColGap
   PLeft(5) = PLeft(4) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(6) = PLeft(5) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(7) = PLeft(6) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(8) = PLeft(7) + Printer.TextWidth(String(4, "　")) + ciColGap
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)

   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      Printer.Print String(125, "-")
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If bolSubtotal Then
         PrintPageHeader1
         iPrint = iPrint + lngLineHeight
      End If
   End If
    
End Sub

Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    PrintNewLine
    For iCol = LBound(strData) To UBound(strData)
        Printer.CurrentX = PLeft(iCol)
        Printer.CurrentY = iPrint
        Printer.Print strData(iCol)
    Next
End Sub

Sub PrintPageHeader()
   Dim strPTmp As String
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = Me.Caption
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   PrintNewLine
   strPTmp = "申請國家："
   Printer.CurrentX = lngPageWidth / 2 - Printer.TextWidth(strPTmp)
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   strPTmp = IIf(Option2(0).Value = True, Option2(0).Caption, Option2(1).Caption)
   Printer.CurrentX = lngPageWidth / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   PrintNewLine
   strPTmp = "預定發文日："
   Printer.CurrentX = lngPageWidth / 2 - Printer.TextWidth(strPTmp)
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   strPTmp = Format(strSrvDate(2), "##/##/##")
   Printer.CurrentX = lngPageWidth / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
    
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   PrintNewLine
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print String(125, "-")
End Sub

Sub PrintPageHeader1()

    Call PrintNewLine(False, 1)
    For intI = 1 To m_iCols
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      Printer.Print DataGrid1.Columns(intI - 1).Caption
    Next
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(125, "-")
    
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

    Call PrintNewLine(True, 1)
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(125, "-")
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "共計： " & iRecCount & " 筆"
    Printer.EndDoc
End Sub

Private Sub Option2_Click(Index As Integer)
   If Index <> m_iLstOption Then
      m_iLstOption = Index
      cmdOK(1).Enabled = False
      cmdOK(3).Enabled = False
      If cmdOK(0).Visible Then cmdOK(0).Value = True 'Added by Morgan 2025/1/10
   End If
End Sub

Public Sub PubShowNextData(ByRef iPt As Integer, ByRef Fgrid As MSHFlexGrid, Optional Index As Integer)
   If iPt = 0 Then Exit Sub
   Me.Enabled = False
   If fnSaveParentForm(Me) = False Then
      Me.Enabled = True
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   frm100101_2.Show
   frm100101_2.Tag = Pub_RplStr(Fgrid.TextMatrix(iPt, 1))
   frm100101_2.cmdOK(5).Visible = False '下一筆按鈕隱藏
   frm100101_2.StrMenu
   Screen.MousePointer = vbDefault
   Me.Enabled = True
End Sub
