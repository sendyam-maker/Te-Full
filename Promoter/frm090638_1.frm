VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090638_1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "商標未發文原因註記"
   ClientHeight    =   5325
   ClientLeft      =   1695
   ClientTop       =   3105
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8760
   Begin VB.CommandButton cmdAddSort 
      Caption         =   "新增↓"
      Height          =   285
      Left            =   3330
      TabIndex        =   5
      Top             =   1860
      Width           =   735
   End
   Begin VB.CommandButton cmdRemSort 
      Caption         =   "移除↑"
      Height          =   285
      Left            =   3330
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   2
      Left            =   5040
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   10
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   7845
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "存檔"
      Height          =   345
      Left            =   5865
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   10
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   0
      Left            =   6705
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   10
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   1785
      Left            =   120
      TabIndex        =   9
      Top             =   3420
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   3149
      _Version        =   393216
      Cols            =   12
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
      _Band(0).Cols   =   12
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Left            =   1110
      TabIndex        =   32
      Top             =   3120
      Visible         =   0   'False
      Width           =   6015
      VariousPropertyBits=   671105051
      MaxLength       =   1000
      Size            =   "10610;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboReason 
      Height          =   315
      Left            =   1050
      TabIndex        =   4
      Top             =   1830
      Width           =   2205
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3881;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstSort 
      Height          =   600
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   8445
      VariousPropertyBits=   746586139
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "14905;1057"
      MatchEntry      =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "歷史資料："
      Height          =   180
      Left            =   120
      TabIndex        =   31
      Top             =   3180
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "未發文原因："
      Height          =   180
      Left            =   120
      TabIndex        =   30
      Top             =   2220
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "原因選項："
      Height          =   180
      Left            =   120
      TabIndex        =   29
      Top             =   1897
      Width           =   900
   End
   Begin MSForms.Label Lbl 
      Height          =   255
      Index           =   9
      Left            =   4920
      TabIndex        =   28
      Top             =   1500
      Width           =   1995
      VariousPropertyBits=   27
      Caption         =   "lbl(9)"
      Size            =   "3519;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl 
      Height          =   255
      Index           =   8
      Left            =   1080
      TabIndex        =   27
      Top             =   1500
      Width           =   1995
      VariousPropertyBits=   27
      Caption         =   "lbl(8)"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl 
      Height          =   255
      Index           =   7
      Left            =   4920
      TabIndex        =   26
      Top             =   1230
      Width           =   1995
      VariousPropertyBits=   27
      Caption         =   "lbl(7)"
      Size            =   "3519;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl 
      Height          =   255
      Index           =   6
      Left            =   1080
      TabIndex        =   25
      Top             =   1230
      Width           =   1995
      VariousPropertyBits=   27
      Caption         =   "lbl(6)"
      Size            =   "3519;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl 
      Height          =   255
      Index           =   5
      Left            =   4920
      TabIndex        =   24
      Top             =   960
      Width           =   3705
      VariousPropertyBits=   27
      Caption         =   "lbl(5)"
      Size            =   "6535;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl 
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   23
      Top             =   960
      Width           =   2835
      VariousPropertyBits=   27
      Caption         =   "lbl(4)"
      Size            =   "5001;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl 
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   22
      Top             =   690
      Width           =   1995
      VariousPropertyBits=   27
      Caption         =   "lbl(3)"
      Size            =   "3519;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   21
      Top             =   690
      Width           =   1995
      VariousPropertyBits=   27
      Caption         =   "lbl(2)"
      Size            =   "3519;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl 
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   20
      Top             =   420
      Width           =   1995
      VariousPropertyBits=   27
      Caption         =   "lbl(1)"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　："
      Height          =   180
      Index           =   9
      Left            =   3960
      TabIndex        =   19
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發文日　："
      Height          =   180
      Index           =   7
      Left            =   3960
      TabIndex        =   17
      Top             =   1230
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文日　："
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   1230
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人　："
      Height          =   180
      Index           =   5
      Left            =   3960
      TabIndex        =   15
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   3960
      TabIndex        =   13
      Top             =   690
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   690
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總收文號："
      Height          =   180
      Index           =   1
      Left            =   3960
      TabIndex        =   11
      Top             =   420
      Width           =   900
   End
   Begin MSForms.Label Lbl 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   420
      Width           =   1995
      VariousPropertyBits=   27
      Caption         =   "lbl(0)"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   420
      Width           =   900
   End
End
Attribute VB_Name = "frm090638_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/22 改成Form2.0 ; GrdDataList改字型=新細明體-ExtB、CboReason、lstSort、Text1
'Memo by Lydia 2019/07/01 表單名稱:商標未發文案件原因註記=>商標未發文原因註記
'Create by Amy 2015/09/04
Option Explicit

Public cmdState As Integer
Public BFormPeople As Integer '操作人身份 1:承辦人 2.智權人員 3.共同查詢(ReadyOnly)
Public BFormStatus As String '前畫面狀態
Public m_NC01 As String, m_NC02 As String
Dim m_PrevForm As Form '前畫面
Dim oLabel  As Object, i As Integer
'Add by Amy 2016/01/06
Public intModSet As Integer

Private Sub cmdAddSort_Click()
    If AddList(lstSort, cboReason) = True Then
        'Modified by Lydia 2021/12/22
        'Text1 = ComposeList(lstSort)
        Text1.Text = lstSort.Tag
        cboReason = ""
    End If
    cboReason.SetFocus
End Sub

'Modified by Lydia 2021/12/22 ComboBox, ListBox =>Object
Private Function AddList(oList As Object, oCombo As Object, Optional p_iOpt As Integer = 0) As Boolean
    Dim idx As Integer, bFound As Boolean, stNewItem As String, iNewItemData As Integer
    Dim stSort As String, iPos As Integer
   
    If oCombo.Text = "" Then
        Exit Function
    End If
   
    '若有控制字元時後面為說明文字不抓
    iPos = InStr(oCombo, Chr(1))
    If iPos > 0 Then
        stNewItem = Left(oCombo, iPos - 1)
    Else
        stNewItem = oCombo
    End If
      
    If InStr(stNewItem, ",") > 0 Then
        MsgBox "逗號[,]為系統保留字，請改用其他符號！", vbExclamation
        oCombo.SetFocus
        Exit Function
    End If
    
    If stNewItem <> "" Then
      'Modified by Lydia 2021/12/22 Form 2.0元件沒有ItemData
      'For idx = 0 To oList.ListCount - 1
         'If oList.List(idx) = stNewItem And oList.ItemData(idx) = iNewItemData Then
         If InStr(oList.Tag, stNewItem) > 0 And oList.Tag <> "" Then
      'end 2021/12/22
            MsgBox "資料已存在！"
            AddList = False
            bFound = True
            'Exit For 'Remove by Lydia 2021/12/22
         End If
      'Next 'Remove by Lydia 2021/12/22
      If bFound = False Then
         oList.AddItem stNewItem, 0
         'Remove by Lydia 2021/12/22 Form 2.0元件沒有ItemData
         'If p_iOpt <> 0 Then
         '   oList.ItemData(0) = oCombo.ItemData(oCombo.ListIndex)
         'End If
         oList.Tag = oCombo.Text & IIf(oList.Tag <> "", ",", "") & oList.Tag   '載入原因是從最末端讀取
         '移到下面
         'end 2021/12/22
         AddList = True
      End If
      
   End If
End Function
Private Sub cmdOK_Click(Index As Integer)
    'Add by Amy 2015/10/01 +離開時判斷是否修改過
    If Text1.Tag <> Text1 And CmdSave.Visible = True Then
        If MsgBox("你並未存檔,確定離開嗎？", vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
    End If
    cmdState = Index
    PubShowNextData
End Sub

Private Sub cmdRemSort_Click()
    If RemoveList(lstSort) = True Then
        'Modified by Lydia 2021/12/22
        'Text1 = ComposeList(lstSort)
        Text1.Text = lstSort.Tag
        cboReason.SetFocus
    End If
End Sub

'Modified by Lydia 2021/12/22 ListBox =>Object
Private Function RemoveList(oList As Object) As Boolean
    Dim ii As Integer
    Dim strTmp As String, idx As Integer 'Added by Lydia 2021/12/22
    
    If oList.ListCount > 0 Then
        ii = 0
        'Modified by Lydia 2021/12/22 Form 2.0元件沒有ItemData
'        Do While ii < oList.ListCount
'            If oList.Selected(ii) = True Then
'                RemoveList = True
'                oList.RemoveItem ii
'                ii = ii - 1
'            End If
'            ii = ii + 1
'        Loop
       strTmp = "," & oList.Tag
       For idx = 0 To oList.ListCount - 1
         If ii >= 0 Then
             If oList.Selected(ii) = True Then
                strTmp = Replace(strTmp, "," & PUB_GetItemData(oList.Tag, ii), "")
                RemoveList = True
                oList.RemoveItem ii
                ii = ii - 1
             Else
                ii = ii + 1
             End If
         End If
      Next
      oList.Tag = Mid(strTmp, 2)
      'end 2021/12/22
    End If
    
End Function

Private Sub CmdSave_Click()
    If FormSave = True Then
        cmdState = 2
        PubShowNextData
    End If
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    cmdState = -1
    SetCboReason
    lstSort.Clear
    
    Lbl(1).Caption = m_NC01
    '共同查詢進入(ReadyOnly)
    If BFormPeople = 3 Then
        Label3.Visible = False
        cboReason.Visible = False
        Label2.Visible = False
        cmdAddSort.Visible = False: cmdRemSort.Visible = False
        Text1.Visible = False: lstSort.Visible = False
        Label4.Top = 1680
        grdDataList.Top = 1920
        grdDataList.Height = 2995
        
        CmdSave.Visible = False
        
        If UCase(m_PrevForm.Name) = "FRM100101_C" Then cmdOK(2).Visible = False: cmdOK(2).Left = 5865
        If UCase(m_PrevForm.Name) = "FRM100107_2" Then cmdOK(0).Visible = False: cmdOK(2).Left = 7015
    Else
        '維護權限判斷,自己只能改自己當月資料
        'Memo by Amy 2016/01/07離職人員 部門主管及帶人主管可修改當月資料
        If CheckModLimit(m_NC01, m_NC02) = True Then
            CmdSave.Visible = True
        Else
            cmdOK(2).Left = 5865
        End If
    End If
End Sub

Private Sub SetDataListWidth()
    Dim stTitle, intWidth
        
    stTitle = Array("管制年月", "專業部原因", "填寫人員", "日期時間", _
                           "智權部原因", "填寫人員", "日期時間")
    
    intWidth = Array(800, 1700, 800, 1300, _
                               1700, 800, 1300)
                            
    For i = 0 To UBound(stTitle)
        grdDataList.row = 0
        grdDataList.col = i
        grdDataList.Text = stTitle(i)
        grdDataList.ColWidth(i) = intWidth(i)
        grdDataList.CellAlignment = flexAlignLeftCenter
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_NC01 = MsgText(601)
    m_NC02 = MsgText(601)
    intModSet = 0
    Set m_PrevForm = Nothing
    Set frm090638_1 = Nothing
End Sub

Private Sub SetCboReason()
    cboReason.Clear
    cboReason.AddItem "申請人資料未齊備"
    cboReason.AddItem "商品未確定"
    cboReason.AddItem "缺公證資料"
    cboReason.AddItem "缺委任書"
    cboReason.AddItem "會稿中"
    'add by sonia 2015/11/3 外商人員才加
    If Pub_StrUserSt03 = "M51" Or Left(Pub_StrUserSt03, 2) = "F1" Then
      cboReason.AddItem "卷退業務承辦等待指示"
    End If
    'end 2015/11/3
End Sub

Public Function QueryData() As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    
    QueryData = False
    
    For Each oLabel In Lbl
        If oLabel.Index <> 1 Then
            oLabel.Caption = ""
        End If
    Next
    
    If BFormPeople = 1 Then
        '承辦人進入
        strQ = ",NC03 as Reason"
    Else
        '智權人進入
        strQ = ",NC07 as Reason"
    End If
    
    strQ = "Select Nvl(cu04,Nvl(cu05,cu06)) as Apply,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as CaseNo,CP10,SQLDATET2(CP05) as CP05,SQLDATET2(CP27) as CP27,na03 as Country," & _
                "CaseName,Nvl(Decode(tm10,'000',CPM03,CPM04),cp10) as CPM,S.ST02 as Sales,P.ST02 as PEmp" & strQ & ",CP13,CP14 From " & _
                "(Select nc01,nc03,nc07,cp01,cp02,cp03,cp04,cp05, cp10,cp12,cp13,cp14,cp27,Decode(tm05,null,Decode(sp05,null,Decode( sp06, null,sp07,sp06),sp05),tm05) as CaseName ,nvl(tm10,sp09) as tm10,nvl(tm23,sp08) as tm23 " & _
                "From NotComplete,CaseProgress,TradeMark ,ServicePractice Where NC01='" & Lbl(1) & "' And NC02=" & Left(strSrvDate(1), 6) & " And NC01=CP09(+) " & _
                "And cp01=tm01(+) And cp02=tm02(+) And cp03=tm03(+) And cp04=tm04(+) And cp01=sp01(+) And cp02=sp02(+) And cp03=sp03(+) And cp04=sp04(+) )," & _
                "Staff S,Staff P,Customer,Nation, CasePropertyMap " & _
                "Where CP13=S.ST01(+) And CP14=P.ST01(+) " & _
                "And SubStr(tm23,1,8)=cu01(+) And  SubStr(tm23,9,1)=cu02(+) And tm10=na01(+) " & _
                "And cp01=CPM01(+) And cp10=CPM02(+) "
                
    If RsQ.State <> 0 Then RsQ.Close
    
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        Lbl(0).Caption = "" & RsQ.Fields("CaseNo")
        Lbl(2).Caption = "" & RsQ.Fields("CPM")
        Lbl(3).Caption = "" & RsQ.Fields("Country")
        Lbl(4).Caption = "" & RsQ.Fields("CaseName")
        Lbl(5).Caption = "" & RsQ.Fields("Apply")
        Lbl(6).Caption = "" & RsQ.Fields("CP05")
        Lbl(7).Caption = "" & RsQ.Fields("CP27")
        Lbl(8).Caption = "" & RsQ.Fields("Sales")
        Lbl(9).Caption = "" & RsQ.Fields("PEmp")
        SetList lstSort, "" & RsQ.Fields("Reason")
        'Add by Amy 2015/10/01 +離開時判斷是否修改過
        Text1 = "" & RsQ.Fields("Reason")
        Text1.Tag = "" & RsQ.Fields("Reason")
        
        HistoryData
        QueryData = True
    End If
    
End Function

'Modified by Lydia 2021/12/22 ListBox =>Object
Private Sub SetList(oList As Object, p_stList As String)
    Dim arrID
    oList.Clear
    oList.Tag = "" 'Added by Lydia 2021/12/22
    If p_stList <> "" Then
        arrID = Split(p_stList, ",")
        For intI = UBound(arrID) To LBound(arrID) Step -1
            oList.AddItem arrID(intI), 0
        Next
        oList.Tag = p_stList 'Added by Lydia 2021/12/22
    End If
End Sub
'Modified by Lydia 2021/12/22 ListBox =>Object
Private Function ComposeList(oList As Object, Optional p_iOpt As Integer = 0) As String
  Dim iPos As Integer, stItem As String, strTemp As String
   
   strTemp = ""
   If oList.ListCount > 0 Then
      For intI = 0 To oList.ListCount - 1
         If p_iOpt = 0 Then
            iPos = InStr(oList.List(intI), Chr(1))
            If iPos > 0 Then
               stItem = Left(oList.List(intI), iPos - 1)
            Else
               stItem = oList.List(intI)
            End If
         Else
            'Modified by Lydia 2021/12/22 Form 2.0元件沒有ItemData
            'stItem = Format(oList.ItemData(intI), "00")
            stItem = oList.List(intI)
         End If
         If intI = 0 Then
            strTemp = stItem
         Else
            strTemp = strTemp & "," & stItem
         End If
      Next
   End If
   ComposeList = strTemp
End Function

Private Sub HistoryData()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strWhere
    Dim intQ As Integer
    
    strWhere = ""
    '抓取非當月歷史資料
    'Mark by Amy 2019/11/14 當月資料仍顯示
'    If BFormPeople <> 3 Then
'        strWhere = "And NC02<>" & Left(strSrvDate(1), 6)
'    End If
    
    strQ = "Select SubStr(nc02,1,4)-1911||'/'||SubStr(nc02,5,2) as 管制年月,nc03 as 專業人員原因,P.st02 as 填寫人員,SQLDATET2(nc05)||' '||Decode(length(nc06),3,'0'||SubStr(nc06,1,1)||':'||SubStr(nc06,2),Decode(length(nc06),4,SubStr(nc06,1,2)||':'||SubStr(nc06,3),Nvl(nc06,''))) as 日期時間," & _
                "nc07 as 專業人員原因,S.st02 as 填寫人員,SQLDATET2(nc09)||' '||Decode(length(nc10),3,SubStr(nc10,1,1)||':'||SubStr(nc10,2),Decode(length(nc10),4,SubStr(nc10,1,2)||':'||SubStr(nc10,3),Nvl(nc10,''))) as 日期時間 " & _
                "From NotComplete,Staff P,Staff S " & _
                "Where NC01='" & Lbl(1) & "' And nc04=P.st01(+) And nc08=S.st01(+) " & strWhere & _
                "Order by NC02 Desc"
                
    If RsQ.State <> 0 Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount = 0 Then
        grdDataList.Clear
        grdDataList.Rows = 2
    Else
        Set grdDataList.Recordset = RsQ
    End If
    SetDataListWidth
End Sub

Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Public Sub PubShowNextData()
    Select Case cmdState
        Case 0 '回前畫面
            If UCase(m_PrevForm.Name) = "FRM090638" Then
                m_PrevForm.Show
                Unload Me
            Else
                tmpBol = fnCancelNowFormAndShowParentForm(Me)
            End If
        Case 1 '結束
            fnCloseAllFrm100
        Case 2 '下一筆
            tmpBol = fnCancelNowFormAndShowParentForm(Me)
    End Select
End Sub

'判斷是否有修改權限
Private Function CheckModLimit(ByVal stNC01 As String, ByVal stNC02 As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    Dim intQ As Integer
    
    CheckModLimit = False
    
    If BFormPeople = 1 Then
        '承辦人進入
        strQ = "CP14"
    Else
        '智權人進入
        strQ = "CP13"
    End If
    
    strQ = "Select " & strQ & ",NC02 From NotComplete,CaseProgress Where CP09='" & stNC01 & "' And NC01=CP09 And NC02=" & Val(m_NC02) + 191100
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        'Add by Amy 2016/01/07+特殊設定檔人員可修改資料
        If intModSet = 4 Then
            CheckModLimit = True
        '前畫面可選到離職人員則可修改資料
        ElseIf ChkStaffST04("" & RsQ.Fields(0), False) = True And intModSet > 0 Then
            CheckModLimit = True
        ElseIf "" & RsQ.Fields(0) = strUserNum And Val(RsQ.Fields(1)) = Val(Left(strSrvDate(1), 6)) Then
            CheckModLimit = True
        End If
    End If
End Function

'Add by Amy 2015/10/01
Private Function FormSave() As Boolean
    Dim strUpd As String, strField As String
    
On Error GoTo ErrHand

    FormSave = False
    
    If lstSort.ListCount = 0 Then
        MsgBox "未發文原因不可為空", , MsgText(5)
        cboReason.SetFocus
        Exit Function
    End If
    
    If BFormPeople = 1 Then
        '承辦人進入
        strField = "NC03='" & Text1 & "',NC04='" & strUserNum & "',NC05=to_number(to_char(sysdate,'YYYYMMDD')),NC06=to_number(SubStr(to_char(sysdate,'HH24MIss'),1,4))"
    Else
        '智權人進入
        strField = "NC07='" & Text1 & "',NC08='" & strUserNum & "',NC09=to_number(to_char(sysdate,'YYYYMMDD')),NC10=to_number(SubStr(to_char(sysdate,'HH24MIss'),1,4))"
    End If
    strUpd = "Update NotComplete Set " & strField & " Where NC01='" & Lbl(1) & "' And NC02=" & Left(strSrvDate(1), 6)
    cnnConnection.Execute strUpd
    
    m_PrevForm.Tag = "Save"
    'Add by Amy 2016/03/03 前畫面選3筆後此只能存檔2筆後,跳回前畫面bug修正-陳金蓮
    m_PrevForm.strVCP09 = Replace(m_PrevForm.strVCP09, ";" & Lbl(1), "")
    FormSave = True
    Exit Function
    
ErrHand:
    MsgBox "修改失敗，請洽電腦中心！" & vbCrLf & Err.Description
End Function
