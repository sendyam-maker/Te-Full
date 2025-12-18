VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm880012 
   BorderStyle     =   1  '單線固定
   Caption         =   "建議代理人選單"
   ClientHeight    =   5436
   ClientLeft      =   192
   ClientTop       =   2520
   ClientWidth     =   8748
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5436
   ScaleWidth      =   8748
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&Y)"
      Height          =   400
      Index           =   0
      Left            =   5580
      TabIndex        =   0
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "取消(&N)"
      Height          =   400
      Index           =   1
      Left            =   6570
      TabIndex        =   1
      Top             =   60
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4635
      Left            =   180
      TabIndex        =   2
      Top             =   600
      Width           =   8370
      _ExtentX        =   14774
      _ExtentY        =   8170
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   6
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label lblDate 
      Height          =   255
      Left            =   210
      TabIndex        =   3
      Top             =   180
      Width           =   4215
      VariousPropertyBits=   27
      Size            =   "7435;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm880012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/28 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lblDate
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
'Create by Morgan 2008/2/13
Option Explicit

Public m_iSelRow As Integer
Public fmParent As Form

'Added by Lydia 2015/10/26 與作業失誤資料維護frm050714共用(考慮已載入專案)
'Memo by Lydia 2016/12/30  1=作業失誤資料; 2=免費修正事由; 3=帳款處理情形歷史記錄; 4=優先權('Added by Lydia 2017/05/09 P,CFP的C類官方來函性質「視為未主張1918」)
'Added by Lydia 2022/03/25 5=DHL收件國家代號
'Add by Sindy 2022/7/4 6=往來記錄
'Memo by Morgan 2024/5/31 7=日本申請中案件的代理人
'Memo by Lydia 2025/02/26 8=選擇智財協作案號的收文號 'Memo by Lydia 2025/04/10 新增智財協作案號的收文號功能按鈕隱藏
Public iTyp As String
'---------------------
Dim pCols As Integer
Dim m_blnColOrderAsc As Boolean 'Added by Lydia 2022/03/25 欄位資料由小到大排序

Private Sub cmdok_Click(Index As Integer)
   Dim stRefNo2 As String
   Select Case Index
      Case 0
         If CheckCheck(stRefNo2) = False Then
            MsgBox "請點選一筆資料！"
            Exit Sub
         End If
      Case 1
         stRefNo2 = ""
   End Select
   fmParent.Tag = stRefNo2
   Unload Me
End Sub

Private Function CheckCheck(Optional p_No2 As String) As Boolean
   Dim ii As Integer
   
   p_No2 = "" 'Added by Lydia 2017/05/09
   
   With grdDataList
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 0) = "V" Then
            'Added by Lydia 2015/10/26
            If iTyp = "1" Then
               p_No2 = .TextMatrix(ii, 2)
            'Added by Lydia 2017/05/09 +優先權
            ElseIf iTyp = "4" Then
               '傳出"優先權號|優先權日|優先權國家"
               p_No2 = p_No2 & Trim(.TextMatrix(ii, 2)) & "|" & Trim(.TextMatrix(ii, 1)) & "|" & Trim(.TextMatrix(ii, 6)) & ";"
            'end 2017/05/09
            'Added by Lydia 2022/03/25 +DHL收件國家
            ElseIf iTyp = "5" Then
               p_No2 = .TextMatrix(ii, 1)
            'end 2022/03/25
            'Added by Lydia 2025/02/26 選擇智財協作案號的收文號
            ElseIf iTyp = "8" Then  '請款年度3碼+CP09+CP14
               p_No2 = Left("" & .TextMatrix(ii, 6), 3) & .TextMatrix(ii, 3) & .TextMatrix(ii, 8)
            'end 2025/02/26
            Else
            'end 2015/10/26
               p_No2 = .TextMatrix(ii, 1)
            End If
                       
            If iTyp <> "4" Then   'Added by Lydia 2017/05/09 優先權可多選
               CheckCheck = True
               Exit For
            End If     'end 2017/05/09
         End If
      Next
      'Added by Lydia 2017/05/09
      If p_No2 <> "" Then
         CheckCheck = True
      End If
      'end 2017/05/09
   End With
End Function

Private Sub Form_Activate()
   SetDataListWidth
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm880012 = Nothing
End Sub

Private Sub SetDataListWidth()
Dim ii As Integer

'Added by Lydia 2015/10/26
Select Case iTyp
   Case "1"
        Me.Caption = "選擇總收文號"
        Me.Height = 4080
        Me.Width = 5970
        grdDataList.Top = 480
        grdDataList.Left = 15
        grdDataList.Height = 3180
        grdDataList.Width = 5820
        cmdOK(0).Left = 3900
        cmdOK(0).Top = 45
        cmdOK(1).Top = 45
        cmdOK(1).Left = 4950
        MoveFormToCenter Me
        With grdDataList
           .FormatString = "V|收文日|總收文號|案件性質|承辦人|智權人員|發文日"
           .ColWidth(0) = 250
           .ColAlignment(0) = flexAlignCenterCenter
           .ColWidth(1) = 850
           .ColAlignment(1) = flexAlignLeftCenter
           .ColWidth(2) = 1000
           .ColAlignment(2) = flexAlignLeftCenter
           .ColWidth(3) = 1000
           .ColAlignment(3) = flexAlignLeftCenter
           .ColWidth(4) = 800
           .ColAlignment(4) = flexAlignLeftCenter
           .ColWidth(5) = 800
           .ColAlignment(5) = flexAlignLeftCenter
           .ColWidth(6) = 850
           .ColAlignment(6) = flexAlignLeftCenter
        End With
        pCols = grdDataList.Cols 'Added by Lydia 2016/12/30
         'Add By Sindy 2019/5/8
         For ii = 7 To pCols - 1
            grdDataList.ColWidth(ii) = 0
         Next ii
         '2019/5/8 END
         
   'Added by Lydia 2016/12/30 免費修正事由
   Case "2"
        Me.Caption = "免費修正事由"
        Me.Height = 4080
        Me.Width = 5970
        grdDataList.Top = 480
        grdDataList.Left = 15
        grdDataList.Height = 3180
        grdDataList.Width = 5820
        cmdOK(0).Left = 3900
        cmdOK(0).Top = 45
        cmdOK(1).Top = 45
        cmdOK(1).Left = 4950
        MoveFormToCenter Me
        With grdDataList
           .FormatString = "V|事由|編號"
           .ColWidth(0) = 300
           .ColAlignment(0) = flexAlignCenterCenter
           .ColWidth(1) = 5000
           .ColAlignment(1) = flexAlignLeftCenter
           .ColWidth(2) = 0
           .ColAlignment(2) = flexAlignLeftCenter
        End With
       pCols = grdDataList.Cols
       'end 2016/12/30
       
   'Added by Lydia 2017/01/16 帳款處理情形歷史記錄
   Case "3"
        Me.Caption = "帳款處理情形歷史記錄"
        Me.Height = 4080
        Me.Width = 6970
        grdDataList.Top = 480
        grdDataList.Left = 15
        grdDataList.Height = 3180
        grdDataList.Width = 6820
        cmdOK(0).Visible = False
        cmdOK(1).Top = 45
        cmdOK(1).Left = 4950
        cmdOK(1).Width = 1300
        cmdOK(1).Caption = "回前畫面(&X)"
        MoveFormToCenter Me
        With grdDataList
           .FormatString = "V|編號|名稱|日期|處理情形|修改人員|時間"
           .ColWidth(0) = 0
           .ColAlignment(0) = flexAlignCenterCenter
           .ColWidth(1) = 1000
           .ColAlignment(1) = flexAlignLeftCenter
           .ColWidth(2) = 1200
           .ColAlignment(2) = flexAlignLeftCenter
           .ColWidth(3) = 860
           .ColAlignment(3) = flexAlignLeftCenter
           .ColWidth(4) = 2100
           .ColAlignment(4) = flexAlignLeftCenter
           .ColWidth(5) = 900
           .ColAlignment(5) = flexAlignLeftCenter
           .ColWidth(6) = 720
           .ColAlignment(6) = flexAlignLeftCenter
        End With
       pCols = grdDataList.Cols
       'end 2017/01/16
       
   'Added by Lydia 2017/05/09 優先權資料
   Case "4"
        Me.Caption = "優先權資料"
        Me.Height = 4080
        Me.Width = 5970
        grdDataList.Top = 480
        grdDataList.Left = 15
        grdDataList.Height = 3180
        grdDataList.Width = 5820
        cmdOK(0).Left = 3900
        cmdOK(0).Top = 45
        cmdOK(1).Top = 45
        cmdOK(1).Left = 4950
        MoveFormToCenter Me
        With grdDataList
           .FormatString = "V|優先權日|優先權號|優先權國家|優先權存取碼|本所案號|PD07"
           .ColWidth(0) = 270
           .ColAlignment(0) = flexAlignCenterCenter
           .ColWidth(1) = 870
           .ColAlignment(1) = flexAlignLeftCenter
           .ColWidth(2) = 1300
           .ColAlignment(2) = flexAlignLeftCenter
           .ColWidth(3) = 1060
           .ColAlignment(3) = flexAlignLeftCenter
           .ColWidth(4) = 1280
           .ColAlignment(4) = flexAlignLeftCenter
           .ColWidth(5) = 1300
           .ColAlignment(5) = flexAlignLeftCenter
           .ColWidth(6) = 0
           .ColAlignment(6) = flexAlignLeftCenter
        End With
        pCols = grdDataList.Cols
   'end 2017/05/09
   
   'Added by Lydia 2022/03/25 DHL收件國家代號
   Case "5"
        Me.Caption = "DHL收件國家代號"
        Me.Height = 4080
        Me.Width = 5970
        grdDataList.Top = 480
        grdDataList.Left = 15
        grdDataList.Height = 3180
        grdDataList.Width = 5820
        cmdOK(0).Left = 3900
        cmdOK(0).Top = 45
        cmdOK(1).Top = 45
        cmdOK(1).Left = 4950
        MoveFormToCenter Me
        With grdDataList
           .FormatString = "V|代號|中文國名|英文國名|NA89"
           .ColWidth(0) = 270
           .ColAlignment(0) = flexAlignCenterCenter
           .ColWidth(1) = 870
           .ColAlignment(1) = flexAlignLeftCenter
           .ColWidth(2) = 1200
           .ColAlignment(2) = flexAlignLeftCenter
           .ColWidth(3) = 2500
           .ColAlignment(3) = flexAlignLeftCenter
           .ColWidth(4) = 0
           .ColAlignment(4) = flexAlignLeftCenter
        End With
        pCols = grdDataList.Cols
   'end 2022/03/25
   
   'Add by Sindy 2022/7/4 往來記錄
   Case "6"
        Me.Caption = "選擇往來記錄"
        Me.Height = 4080
        Me.Width = 5970
        grdDataList.Top = 480
        grdDataList.Left = 15
        grdDataList.Height = 3180
        grdDataList.Width = 5820
        cmdOK(0).Left = 3900
        cmdOK(0).Top = 45
        cmdOK(1).Top = 45
        cmdOK(1).Left = 4950
        MoveFormToCenter Me
        With grdDataList
           .FormatString = "V|往來記錄編號|往來日期|主旨|內容|建檔人員|建檔日期"
           .ColWidth(0) = 270
           .ColAlignment(0) = flexAlignCenterCenter
           .ColWidth(1) = 1000
           .ColAlignment(1) = flexAlignLeftCenter
           .ColWidth(2) = 850
           .ColAlignment(2) = flexAlignLeftCenter
           .ColWidth(3) = 1800
           .ColAlignment(3) = flexAlignLeftCenter
           .ColWidth(4) = 1800
           .ColAlignment(4) = flexAlignLeftCenter
           .ColWidth(5) = 850
           .ColAlignment(5) = flexAlignLeftCenter
           .ColWidth(6) = 850
           .ColAlignment(6) = flexAlignLeftCenter
        End With
        pCols = grdDataList.Cols
   
   Case 7 'Added by Morgan 2024/5/31
      Me.Caption = "本案申請人目前於日本申請中案件的代理人"
      Me.Height = 4080
      Me.Width = 8000
      With grdDataList
         .Width = 7550
         .Height = 2900
         .FormatString = "選擇|代理人|名稱"
         .ColWidth(0) = 465
         .ColAlignment(0) = flexAlignCenterCenter
         .ColWidth(1) = 1300
         .ColAlignment(1) = flexAlignLeftCenter
         .ColWidth(2) = 5400
         .ColAlignment(2) = flexAlignLeftCenter
         For ii = 3 To .Cols - 1
            .ColWidth(ii) = 0
         Next
      End With
   'Added by Lydia 2025/02/26　TIPS案請款階段分配比例維護作業：選擇智財協作案號的收文號 'Memo by Lydia 2025/04/10 新增智財協作案號的收文號功能按鈕隱藏
   Case "8"
        Me.Caption = "選擇智財協作案號的收文號"
        MoveFormToCenter Me
        With grdDataList
           .FormatString = "V|協作案號|案件名稱|收文號|案件性質|收文日|發文日|承辦人|CP14"
           .ColWidth(0) = 270
           .ColAlignment(0) = flexAlignCenterCenter
           .ColWidth(1) = 1050
           .ColAlignment(1) = flexAlignLeftCenter
           .ColWidth(2) = 1800
           .ColAlignment(2) = flexAlignLeftCenter
           .ColWidth(3) = 1000
           .ColAlignment(3) = flexAlignLeftCenter
           .ColWidth(4) = 1300
           .ColAlignment(4) = flexAlignLeftCenter
           .ColWidth(5) = 900
           .ColAlignment(5) = flexAlignLeftCenter
           .ColWidth(6) = 900
           .ColAlignment(6) = flexAlignLeftCenter
           .ColWidth(7) = 950
           .ColAlignment(7) = flexAlignLeftCenter
           .ColWidth(8) = 0
           .ColAlignment(8) = flexAlignLeftCenter
        End With
        pCols = grdDataList.Cols
   Case Else
   'end 2015/10/26
        With grdDataList
           'Modified by Lydia 2015/10/26
           '.FormatString = .FormatString
           'Modified by Lydia 2018/01/15 +FC給案
           .FormatString = "選擇|代理人/聯絡人編號|名稱|給案量|建議量|給案率|FC給案"
           .ColWidth(0) = 465
           .ColAlignment(0) = flexAlignCenterCenter
           .ColWidth(1) = 1300
           .ColAlignment(1) = flexAlignLeftCenter
            'Added by Lydia 2018/01/15 CFT案+FC給案
           If fmParent.Name = "frm030001_1" Then
               .ColWidth(2) = 3300
           Else
           'end 2018/01/15
               .ColWidth(2) = 4000
           End If 'end 2018/01/15
           .ColAlignment(2) = flexAlignLeftCenter
           .ColWidth(3) = 750
           .ColAlignment(3) = flexAlignRightCenter
           .ColWidth(4) = 750
           .ColAlignment(4) = flexAlignRightCenter
           .ColWidth(5) = 750
           .ColAlignment(5) = flexAlignRightCenter
           'Added by Lydia 2018/01/15 CFT案+FC給案
           If fmParent.Name = "frm030001_1" Then
               .ColAlignment(6) = flexAlignRightCenter
               .ColWidth(6) = 750
               .ColWidth(7) = 0
           Else
           'end 2018/01/15
               .ColWidth(6) = 0
           End If 'end 2018/01/15
        End With
        
        pCols = grdDataList.Cols 'Added by Lydia 2016/12/30
End Select

'Added by Lydia 2015/10/26
'Remove by Lydia 2016/12/30
'pCols = GrdDataList.Cols

End Sub

Private Sub grdSelected(p_iRow As Integer)
   Dim stCheck As String, lColor As Long, ii As Integer
   With grdDataList
      .row = p_iRow
      .col = 0
      If .Text = "" Then
         .Text = "V"
         m_iSelRow = .row
         lColor = &HFFC0C0
      Else
         .Text = ""
         m_iSelRow = -1
         lColor = &H80000018
      End If
      'Modified by Lydia 2015/10/26
      'For ii = 0 To .Cols - 1
      For ii = 0 To pCols - 1
         .col = ii
         .CellBackColor = lColor
      Next
   End With
End Sub

Private Sub GrdDataList_Click()
   Dim iRow As Integer
   With grdDataList
      If .MouseRow > 0 And .MouseRow < .Rows Then
         .Visible = False
         iRow = .MouseRow
         
         If m_iSelRow > 0 Then
            If iTyp <> "4" Or (iTyp = "4" And m_iSelRow = iRow) Then 'Added by Lydia 2017/05/09 優先權資料可多選
               grdSelected m_iSelRow
            End If 'end 2017/05/09
         End If
        
         If m_iSelRow <> iRow Then
            grdSelected iRow
         End If
         .Visible = True
      End If
   End With
End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   'Added by Lydia 2022/03/25 DHL收件國家代號：增加點選Grid最上方抬頭可排序的功能
   If iTyp = "5" Then
        getGrdColRow grdDataList, x, y, nCol, nRow
        If nCol < 0 Or nRow < 0 Then Exit Sub
        grdDataList.col = nCol
        grdDataList.row = nRow
        If Me.grdDataList.row < 1 And Me.grdDataList.Text <> "V" Then
           If InStr("TEST", Me.grdDataList.Text) > 0 Then '保留數值排序
              If m_blnColOrderAsc = True Then
                 Me.grdDataList.Sort = 3  '數值昇冪
                 m_blnColOrderAsc = False
              Else
                 Me.grdDataList.Sort = 4 '數值降冪
                 m_blnColOrderAsc = True
              End If
           Else
              If m_blnColOrderAsc = True Then
                 Me.grdDataList.Sort = 5 '字串昇冪
                 m_blnColOrderAsc = False
              Else
                 Me.grdDataList.Sort = 6 '字串降冪
                 m_blnColOrderAsc = True
              End If
           End If
        End If
   End If
End Sub

