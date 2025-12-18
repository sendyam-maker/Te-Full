VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090220 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人電子送件作業"
   ClientHeight    =   5720
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5720
   ScaleWidth      =   8920
   Begin VB.CommandButton Command1 
      Caption         =   "結束"
      Height          =   375
      Left            =   7560
      TabIndex        =   11
      Top             =   90
      Width           =   1230
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4845
      Left            =   90
      TabIndex        =   13
      Top             =   780
      Width           =   8745
      _ExtentX        =   15416
      _ExtentY        =   8537
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "待選取"
      TabPicture(0)   =   "frm090220.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command7"
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(2)=   "Text2"
      Tab(0).Control(3)=   "Text3"
      Tab(0).Control(4)=   "Command2"
      Tab(0).Control(5)=   "Command3"
      Tab(0).Control(6)=   "MSHFlexGrid1"
      Tab(0).Control(7)=   "MaskEdBox2"
      Tab(0).Control(8)=   "MaskEdBox1"
      Tab(0).Control(9)=   "lblCPM"
      Tab(0).Control(10)=   "Label1(1)"
      Tab(0).Control(11)=   "Label1(0)"
      Tab(0).Control(12)=   "Line1"
      Tab(0).Control(13)=   "Line2"
      Tab(0).Control(14)=   "Label1(2)"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "已選取"
      TabPicture(1)   =   "frm090220.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "MSHFlexGrid2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtInput"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "已送件"
      TabPicture(2)   =   "frm090220.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSHFlexGrid3"
      Tab(2).Control(1)=   "Command5"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton Command7 
         Caption         =   "預設資料"
         Height          =   375
         Left            =   -67620
         TabIndex        =   6
         Top             =   390
         Width           =   1230
      End
      Begin VB.CommandButton Command6 
         Caption         =   "取消選取"
         Height          =   375
         Left            =   6120
         TabIndex        =   8
         Top             =   810
         Width           =   1230
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -73605
         MaxLength       =   9
         TabIndex        =   0
         Top             =   480
         Width           =   1005
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -72435
         MaxLength       =   9
         TabIndex        =   1
         Top             =   480
         Width           =   1005
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   -70590
         TabIndex        =   4
         Top             =   870
         Width           =   1005
      End
      Begin VB.CommandButton Command2 
         Caption         =   "全部"
         Height          =   375
         Left            =   -68880
         TabIndex        =   5
         Top             =   390
         Width           =   1230
      End
      Begin VB.CommandButton Command3 
         Caption         =   "送件確認"
         Height          =   375
         Left            =   -67620
         TabIndex        =   7
         Top             =   810
         Width           =   1230
      End
      Begin VB.CommandButton Command4 
         Caption         =   "送件存檔"
         Height          =   375
         Left            =   7380
         TabIndex        =   9
         Top             =   810
         Width           =   1230
      End
      Begin VB.CommandButton Command5 
         Caption         =   "送件取消"
         Height          =   375
         Left            =   -67620
         TabIndex        =   10
         Top             =   810
         Width           =   1230
      End
      Begin VB.TextBox txtInput 
         Appearance      =   0  '平面
         Height          =   270
         Left            =   225
         TabIndex        =   14
         Text            =   "Text3"
         Top             =   630
         Width           =   870
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3465
         Left            =   -74865
         TabIndex        =   15
         Top             =   1260
         Width           =   8475
         _ExtentX        =   14958
         _ExtentY        =   6121
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
         FormatString    =   "勾選|本所案號|商品類別|申請人|智權人員|收文日"
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
         _Band(0).Cols   =   6
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   285
         Left            =   -72795
         TabIndex        =   3
         Top             =   870
         Width           =   1005
         _ExtentX        =   1782
         _ExtentY        =   512
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   285
         Left            =   -73965
         TabIndex        =   2
         Top             =   870
         Width           =   1005
         _ExtentX        =   1782
         _ExtentY        =   512
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3465
         Left            =   135
         TabIndex        =   16
         Top             =   1260
         Width           =   8475
         _ExtentX        =   14958
         _ExtentY        =   6121
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         HighLight       =   0
         SelectionMode   =   1
         FormatString    =   "|本所案號|商品類別|智權人員|收文規費|發文規費"
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
         _Band(0).Cols   =   6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   3465
         Left            =   -74865
         TabIndex        =   17
         Top             =   1260
         Width           =   8475
         _ExtentX        =   14958
         _ExtentY        =   6121
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
         FormatString    =   "勾選|本所案號|商品類別|智權人員|收文規費|發文規費"
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
         _Band(0).Cols   =   6
      End
      Begin MSForms.Label lblCPM 
         Height          =   255
         Left            =   -69570
         TabIndex        =   23
         Top             =   900
         Width           =   1215
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "2143;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人編號："
         Height          =   180
         Index           =   1
         Left            =   -74730
         TabIndex        =   20
         Top             =   525
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "收文日："
         Height          =   180
         Index           =   0
         Left            =   -74730
         TabIndex        =   19
         Top             =   915
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   -72570
         X2              =   -72420
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line2 
         X1              =   -72930
         X2              =   -72780
         Y1              =   1005
         Y2              =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "案件性質："
         Height          =   180
         Index           =   2
         Left            =   -71535
         TabIndex        =   18
         Top             =   915
         Width           =   900
      End
   End
   Begin MSForms.ComboBox cboUser 
      Height          =   315
      Left            =   990
      TabIndex        =   22
      Top             =   420
      Width           =   1860
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "3281;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Left            =   180
      TabIndex        =   21
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "請先做智慧局電子送件後才執行本作業!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   180
      TabIndex        =   12
      Top             =   90
      Width           =   5025
   End
End
Attribute VB_Name = "frm090220"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/23 改成Form2.0 ;cboUser、MSHFlexGrid1改字型=新細明體-ExtB、MSHFlexGrid2改字型=新細明體-ExtB、MSHFlexGrid3改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Created by Morgan 2011/6/14
Option Explicit
Dim iRow As Integer '本次點選列數
Dim iCol As Integer '本次點選欄數
Dim m_UserId As String 'Added by Morgan 2012/3/13

Private Sub cboUser_Click()
   Dim strUser As String
   'Modified by Lydia 2021/12/23
   'strUser = PUB_Num2Id(cboUser.ItemData(cboUser.ListIndex))
   strUser = PUB_GetItemData(cboUser.Tag, cboUser.ListIndex)

   If m_UserId <> strUser Then
      m_UserId = strUser
      RefreshGrid1 False, IIf(SSTab1.Tab = 0, True, False)
      RefreshGrid3
   End If
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command2_Click()
   RefreshGrid1 True
End Sub

Private Sub Command3_Click()
   Dim idx As Integer
   With MSHFlexGrid2
   intI = 1
   Do While intI < MSHFlexGrid1.Rows
      If MSHFlexGrid1.TextMatrix(intI, 0) = "V" Then
         idx = .Rows
         .AddItem "", idx
         .TextMatrix(idx, 1) = MSHFlexGrid1.TextMatrix(intI, 1)
         .TextMatrix(idx, 2) = MSHFlexGrid1.TextMatrix(intI, 2)
         .TextMatrix(idx, 3) = MSHFlexGrid1.TextMatrix(intI, 3)
         .TextMatrix(idx, 4) = MSHFlexGrid1.TextMatrix(intI, 5)
         .TextMatrix(idx, 5) = MSHFlexGrid1.TextMatrix(intI, 7)
         .TextMatrix(idx, 6) = MSHFlexGrid1.TextMatrix(intI, 7)
         .TextMatrix(idx, 7) = MSHFlexGrid1.TextMatrix(intI, 8)
         If MSHFlexGrid1.Rows = 2 Then
            MSHFlexGrid1.AddItem "", MSHFlexGrid1.Rows
         End If
         MSHFlexGrid1.RemoveItem intI
      Else
         intI = intI + 1
      End If
   Loop
   If .Rows > 2 And .TextMatrix(1, 1) = "" Then
      .RemoveItem 1
   End If
   End With
End Sub

Private Sub Command4_Click()
   Dim idx As Integer
   Dim Cancel As Boolean
   
   If txtInput.Visible Then
      txtInput_Validate Cancel
      If Cancel Then txtInput.SetFocus: Exit Sub
   End If
   
   With MSHFlexGrid2
   For idx = .Rows - 1 To 1 Step -1
   If .TextMatrix(idx, 7) <> "" Then
      strSql = "update caseprogress set cp85=" & strSrvDate(1) & ",cp84=" & Val(Format(.TextMatrix(idx, 6))) & _
         ",cp118='Y' where cp09='" & .TextMatrix(idx, 7) & "'"
      cnnConnection.Execute strSql, intI
      If .Rows = 2 Then
         .AddItem "", .Rows
      End If
      .RemoveItem idx
   End If
   Next
   RefreshGrid3
   txtInput.Visible = False
   End With
End Sub

Private Sub Command5_Click()
   Dim idx As Integer
   With MSHFlexGrid3
   For idx = .Rows - 1 To 1 Step -1
   If .TextMatrix(idx, 0) = "V" Then
      strSql = "update caseprogress set cp85=null,cp84=null" & _
         ",cp118=null where cp09='" & .TextMatrix(idx, 7) & "'"
      cnnConnection.Execute strSql, intI
      If .Rows = 2 Then
         .AddItem "", .Rows
      End If
      .RemoveItem idx
   End If
   Next
   End With
End Sub

Private Sub Command6_Click()
   Dim idx As Integer
   With MSHFlexGrid2
   For idx = .Rows - 1 To 1 Step -1
      If .TextMatrix(idx, 0) = "V" Then
         If .Rows = 2 Then
            .AddItem "", .Rows
         End If
         .RemoveItem idx
      End If
   Next
   End With
End Sub

Private Sub Command7_Click()
   RefreshGrid1
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SSTab1.Tab = 0
   SetUser 'Added by Morgan 2012/3/13
   RefreshGrid1
   RefreshGrid2
   RefreshGrid3
   'Modified by Lydia 2021/12/23 改Form 2.0
   'Label1(3).BackColor = Me.BackColor
   lblCPM.BackColor = Me.BackColor
   MaskEdBox1.Mask = "###/##/##"
   MaskEdBox2.Mask = "###/##/##"
   txtInput.Visible = False
End Sub

'Added by Morgan 2012/3/13
Private Sub SetUser()
   m_UserId = strUserNum
   cboUser.Tag = "": strExc(1) = ""   'Added by Lydia 2021/12/23
   strExc(0) = "select st01,st02" & _
      " From staff, staff_Absence" & _
      " WHERE st04='1' and st03='" & Pub_StrUserSt03 & "' and st01<'F'" & _
      " and SA01(+)=ST01 and to_char(sysdate,'yyyymmddhh24mi')" & _
      " between SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4)" & _
      " and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         cboUser.AddItem .Fields("st02"), 0
         'Modified by Lydia 2021/12/23 改成Form 2.0沒有ItemData
         'cboUser.ItemData(0) = PUB_Id2Num(.Fields("st01"))
         strExc(1) = .Fields("st01") & IIf(strExc(1) <> "", ",", "") & strExc(1)
         .MoveNext
      Loop
      End With
   End If
   cboUser.AddItem strUserName, 0
   'Modified by Lydia 2021/12/23
   'cboUser.ItemData(0) = PUB_Id2Num(strUserNum)
   strExc(1) = strUserNum & IIf(strExc(1) <> "", ",", "") & strExc(1)
   cboUser.Tag = strExc(1)
   'end 2021/12/23
   cboUser.ListIndex = 0
End Sub

Private Sub RefreshGrid1(Optional bolAll As Boolean, Optional bolShowMsg As Boolean = True)
   Dim strCon As String, ii As Integer, jj As Integer
      
   strCon = " and cp14='" & m_UserId & "'"
   If bolAll Then
      If Text1 <> "" Then
         strCon = strCon & " and tm23>='" & Text1 & "'"
      End If
      If Text2 <> "" Then
         strCon = strCon & " and tm23<='" & Text2 & "'"
      End If
      If Text3 <> "" Then
         strCon = strCon & " and cp10='" & Text3 & "'"
      End If
      If MaskEdBox1 <> "___/__/__" Then
         strCon = strCon & " and cp05>=" & DBDATE(MaskEdBox1)
      Else
         strCon = strCon & " and cp05>20110000"
      End If
      If MaskEdBox2 <> "___/__/__" Then
         strCon = strCon & " and cp05<=" & DBDATE(MaskEdBox2)
      End If
   Else
      strCon = strCon & " and cp05>20110000 and cp10='101' and cp17<3000"
   End If
   
   'MODIFY BY SONIA 2015/9/9 爭議案也可電子送件,故開放FCT(FCT-032736)
   'Modified by Morgan 2016/1/5 +cp141,cp79,cp06
   'Modified by Morgan 2016/8/29 cp27,cp57改判斷cp158,cp159
   strExc(0) = "select '' C0,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) C1" & _
      ",cpm03 C2,TM09 C3,cu04 C4,st02 C5,sqldatet(cp05) C6,cp17 C7,cp09 C8,cp141,cp79,cp06" & _
      " from caseprogress,casepropertymap,trademark,customer,staff" & _
      " where cp158=0 and cp159=0 and cp01 in ('T','FCT')" & strCon & _
      " and cp85 is null and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm10='000'" & _
      " and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9) and st01(+)=cp13 order by 2,3"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modify by Amy 2014/06/10 +FormName 改暫存TB
      'Set MSHFlexGrid1.Recordset = PUB_CreateRecordset(RsTemp)
      Set MSHFlexGrid1.Recordset = PUB_CreateRecordset(RsTemp, , , , Me.Name)
      '剔除已選取者
      If bolAll Then
         With MSHFlexGrid1
         .Visible = False
         For ii = .Rows - 1 To 1 Step -1
            For jj = 1 To MSHFlexGrid2.Rows - 1
               If .TextMatrix(ii, 8) = MSHFlexGrid2.TextMatrix(jj, 7) Then
                  If .Rows = 2 Then
                     .AddItem "", 2
                  End If
                  .RemoveItem ii
               End If
            Next
         Next
         .Visible = True
         End With
      End If
      SetDataListWidth1
   Else
      MSHFlexGrid1.Clear
      MSHFlexGrid1.Cols = 8
      MSHFlexGrid1.Rows = 2
      SetDataListWidth1
      If bolShowMsg Then MsgBox "無符合資料！"
   End If
   
End Sub

Private Sub RefreshGrid2()
   SetDataListWidth2
End Sub

Private Sub RefreshGrid3()
   'Modified by Morgan 2015/4/17 因有可能要改前一天送件的發文規費, 故改抓2天的資料 Ex.T-179673
   'MODIFY BY SONIA 2015/9/9 爭議案也可電子送件,故開放FCT(FCT-032736)
   'Modified by Morgan 2016/8/29 cp27,cp57改判斷cp158,cp159
   strExc(0) = "select '' C0,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) C1" & _
      ",cpm03 C2,TM09 C3,st02 C4,cp17 C5,cp84 C6,cp09 C7" & _
      " from caseprogress,casepropertymap,trademark,customer,staff" & _
      " where cp158=0 and cp159=0 and cp01 in ('T','FCT') and cp14='" & m_UserId & "'" & _
      " and cp85>=to_char(sysdate-1,'yyyymmdd') and cp118||''='Y' and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9) and st01(+)=cp13"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modify by Amy 2014/06/10 +FormName 改暫存TB
      'Set MSHFlexGrid3.Recordset = PUB_CreateRecordset(RsTemp)
      Set MSHFlexGrid3.Recordset = PUB_CreateRecordset(RsTemp, , , , Me.Name)
      SetDataListWidth3
   Else
      MSHFlexGrid3.Cols = 8
      MSHFlexGrid3.Clear
      MSHFlexGrid3.Rows = 2
      SetDataListWidth3
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090220 = Nothing
End Sub

Private Sub SetDataListWidth1()
   With MSHFlexGrid1
   .row = 0
   .col = 0: .Text = "V"
   .ColWidth(0) = 250
   .CellAlignment = flexAlignLeftCenter
   .col = 1: .Text = "本所案號"
   .ColWidth(1) = 1250
   .CellAlignment = flexAlignLeftCenter
   .col = 2: .Text = "案件性質"
   .ColWidth(2) = 1080
   .CellAlignment = flexAlignLeftCenter
   .col = 3: .Text = "商品類別"
   .ColWidth(3) = 900
   .CellAlignment = flexAlignLeftCenter
   .col = 4: .Text = "申請人"
   .ColWidth(4) = 3105
   .CellAlignment = flexAlignLeftCenter
   .col = 5: .Text = "智權人員"
   .ColWidth(5) = 700
   .CellAlignment = flexAlignLeftCenter
   .col = 6: .Text = "收文日"
   .ColWidth(6) = 840
   .CellAlignment = flexAlignLeftCenter
   For intI = 7 To .Cols - 1
      .ColWidth(intI) = 0
   Next
   End With
End Sub

Private Sub MaskEdBox1_GotFocus()
   MaskEdBoxInverse MaskEdBox1
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1 <> "___/__/__" Then
      strExc(1) = Replace(MaskEdBox1, "/", "")
      If ChkDate(strExc(1)) = False Then
         Cancel = True
         MaskEdBox1.SetFocus
         MaskEdBox1_GotFocus
      End If
   End If
End Sub

Private Sub MaskEdBox2_GotFocus()
   MaskEdBoxInverse MaskEdBox2
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   If MaskEdBox2 <> "___/__/__" Then
      strExc(1) = Replace(MaskEdBox2, "/", "")
      If ChkDate(strExc(1)) = False Then
         Cancel = True
         MaskEdBox2.SetFocus
         MaskEdBox2_GotFocus
      End If
   End If
End Sub

Private Sub MSHFlexGrid1_Click()
   With MSHFlexGrid1
   If .MouseRow = 0 Then
      .col = .MouseCol
      .Sort = 7
   End If
   End With
End Sub

Private Sub MSHFlexGrid1_SelChange()
   Dim ii As Integer
   Screen.MousePointer = vbHourglass
   With MSHFlexGrid1
   .Visible = False
   .row = .MouseRow
   .col = 0
   If .row <> 0 And .TextMatrix(.row, 1) <> "" Then
      If .Text = "V" Then
           .Text = ""
           For ii = 0 To .Cols - 1
                .col = ii
                .CellBackColor = QBColor(15)
          Next ii
      Else
         'Added by Morgan 2016/1/5
         If .TextMatrix(.row, 9) = "2" And Val(.TextMatrix(.row, 10)) > 0 Then
            If PUB_ChkPaidByCP09(.TextMatrix(.row, 8)) = False Then 'Added by Morgan 2016/8/29 出納繳款確認後就可送件
               If "" & .TextMatrix(.row, 11) = "" Or "" & .TextMatrix(.row, 11) > strSrvDate(1) Then
                  .Visible = True
                  MsgBox "此案智權人員欲管控收款後才可送件，暫不可發文！"
                  GoTo EXITSUB
               End If
            End If
         End If
         'end 2016/1/5
               
           .Text = "V"
           For ii = 0 To .Cols - 1
               .col = ii
               .CellBackColor = &HFFC0C0
           Next ii
      End If
   End If
   .Visible = True
EXITSUB:
   End With
   Screen.MousePointer = vbDefault
End Sub

Private Sub SetDataListWidth2()
   With MSHFlexGrid2
   .Cols = 8
   .row = 0
   .col = 0: .Text = "V"
   .ColWidth(0) = 250
   .CellAlignment = flexAlignLeftCenter
   .col = 1: .Text = "本所案號"
   .ColWidth(1) = 1250
   .CellAlignment = flexAlignLeftCenter
   .col = 2: .Text = "案件性質"
   .ColWidth(2) = 1080
   .CellAlignment = flexAlignLeftCenter
   .col = 3: .Text = "商品類別"
   .ColWidth(3) = 900
   .CellAlignment = flexAlignLeftCenter
   .col = 4: .Text = "智權人員"
   .ColWidth(4) = 700
   .CellAlignment = flexAlignLeftCenter
   .col = 5: .Text = "收文規費"
   .ColWidth(5) = 840
   .CellAlignment = flexAlignRightCenter
   .col = 6: .Text = "發文規費"
   .ColWidth(6) = 840
   .CellAlignment = flexAlignRightCenter
   For intI = 7 To .Cols - 1
      .ColWidth(intI) = 0
   Next
   End With
End Sub

Private Sub SetDataListWidth3()
   With MSHFlexGrid3
   .row = 0
   .col = 0: .Text = "V"
   .ColWidth(0) = 250
   .CellAlignment = flexAlignLeftCenter
   .col = 1: .Text = "本所案號"
   .ColWidth(1) = 1250
   .CellAlignment = flexAlignLeftCenter
   .col = 2: .Text = "案件性質"
   .ColWidth(2) = 1080
   .CellAlignment = flexAlignLeftCenter
   .col = 3: .Text = "商品類別"
   .ColWidth(3) = 900
   .CellAlignment = flexAlignLeftCenter
   .col = 4: .Text = "智權人員"
   .ColWidth(4) = 700
   .CellAlignment = flexAlignLeftCenter
   .col = 5: .Text = "收文規費"
   .ColWidth(5) = 840
   .CellAlignment = flexAlignRightCenter
   .col = 6: .Text = "發文規費"
   .ColWidth(6) = 840
   .CellAlignment = flexAlignRightCenter
   For intI = 7 To .Cols - 1
      .ColWidth(intI) = 0
   Next
   End With
End Sub

Private Sub MSHFlexGrid2_Click()
   With MSHFlexGrid2
   .row = .MouseRow
   .col = .MouseCol
   If .row = 0 Then
      '.Sort = 7
   Else
      SetBox
   End If
   End With
End Sub

Private Sub MSHFlexGrid2_SelChange()
   Dim ii As Integer
   If MSHFlexGrid2.MouseCol = 6 Then Exit Sub
   Screen.MousePointer = vbHourglass
   With MSHFlexGrid2
   .row = .MouseRow
   .Visible = False
   .col = 0
   If .row <> 0 And .TextMatrix(.row, 1) <> "" Then
      If .Text = "V" Then
           .Text = ""
           For ii = 0 To .Cols - 1
                .col = ii
                .CellBackColor = QBColor(15)
          Next ii
      Else
           .Text = "V"
           For ii = 0 To .Cols - 1
               .col = ii
               .CellBackColor = &HFFC0C0
           Next ii
      End If
   End If
   .Visible = True
   End With
   Screen.MousePointer = vbDefault
End Sub

Private Sub MSHFlexGrid3_Click()
   With MSHFlexGrid3
   If .MouseRow = 0 Then
      .col = .MouseCol
      .Sort = 7
   End If
   End With
End Sub

Private Sub MSHFlexGrid3_SelChange()
   Dim ii As Integer
   Screen.MousePointer = vbHourglass
   With MSHFlexGrid3
   .row = .MouseRow
   .Visible = False
   .col = 0
   If .row <> 0 And .TextMatrix(.row, 1) <> "" Then
      If .Text = "V" Then
           .Text = ""
           For ii = 0 To .Cols - 1
                .col = ii
                .CellBackColor = QBColor(15)
          Next ii
      Else
           .Text = "V"
           For ii = 0 To .Cols - 1
               .col = ii
               .CellBackColor = &HFFC0C0
           Next ii
      End If
   End If
   .Visible = True
   End With
   Screen.MousePointer = vbDefault
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   'Modified by Lydia 2021/12/23 改Form 2.0
   'Label1(3) = GetCaseTypeName("T", Text3)
   lblCPM.Caption = GetCaseTypeName("T", Text3)
End Sub

Private Sub SetBox()
   
   Dim lngLeft As Long, lngTop As Long, ii As Integer
   
   With MSHFlexGrid2
      If .row > 0 And .col = 6 Then
         If .TextMatrix(.row, 7) <> "" Then
            txtInput.FontName = .CellFontName
            txtInput.FontSize = .CellFontSize
            txtInput.Alignment = .CellAlignment \ 5
            txtInput.Text = .TextMatrix(.row, .col)
            txtInput.Tag = txtInput.Text
            txtInput.Width = .ColWidth(.col)
            txtInput.Height = .RowHeight(.row)
            iRow = .row: iCol = .col
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

Private Sub txtInput_GotFocus()
   CloseIme
   TextInverse txtInput
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
   If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = Asc(".") Or KeyAscii = 8 Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
      KeyAscii = 0
      Beep
   Else
      If KeyAscii = vbKeyReturn Then
         'Modified by Morgan 2012/8/13 取消規費>0的限制(分割案只有第一件有規費)
         'If Val(txtInput) > 0 Then
         '   MSHFlexGrid2.TextMatrix(iRow, iCol) = txtInput.Text
         '   GoNext
         'Else
         '   MsgBox "發文規費必須大於 0 ！"
         'End If
         MSHFlexGrid2.TextMatrix(iRow, iCol) = Val(txtInput.Text)
         GoNext
         'end 2012/8/13
      ElseIf KeyAscii = vbKeyEscape Then
         txtInput = txtInput.Tag
         TextInverse txtInput
      End If
   End If
End Sub

Private Sub GoNext()
   With MSHFlexGrid2
      If .row < .Rows - 1 Then
         .row = .row + 1
      Else
         .row = 1
      End If
      SetBox
   End With
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
   'Modified by Morgan 2012/8/13 取消規費>0的限制(分割案只有第一件有規費)
   'If Val(txtInput) > 0 Then
   '   MSHFlexGrid2.TextMatrix(iRow, iCol) = txtInput.Text
   '   txtInput.Visible = False
   'Else
   '   MsgBox "發文規費必須大於 0 ！"
   '   Cancel = True
   'End If
   MSHFlexGrid2.TextMatrix(iRow, iCol) = Val(txtInput.Text)
   'end 2012/8/13
End Sub
