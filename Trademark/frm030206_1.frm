VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030206_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書"
   ClientHeight    =   5748
   ClientLeft      =   132
   ClientTop       =   2412
   ClientWidth     =   9348
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9348
   Begin VB.CommandButton cmdOK 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   3048
      TabIndex        =   4
      Top             =   528
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "FCT"
      Top             =   576
      Width           =   550
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1512
      MaxLength       =   6
      TabIndex        =   1
      Top             =   576
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2352
      MaxLength       =   1
      TabIndex        =   2
      Top             =   576
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   2592
      MaxLength       =   2
      TabIndex        =   3
      Top             =   576
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   270
      Left            =   6420
      MaxLength       =   7
      TabIndex        =   5
      Top             =   576
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7464
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8316
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1140
      MaxLength       =   1
      TabIndex        =   7
      Top             =   5340
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3600
      Left            =   120
      TabIndex        =   6
      Top             =   1656
      Width           =   9072
      _ExtentX        =   16002
      _ExtentY        =   6350
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   330
      Left            =   960
      TabIndex        =   19
      Top             =   1221
      Width           =   8235
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "14526;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期:"
      Height          =   180
      Left            =   5400
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號"
      Height          =   180
      Left            =   120
      TabIndex        =   16
      Top             =   936
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   180
      Left            =   960
      TabIndex        =   15
      Top             =   930
      Width           =   2730
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "審定號數:"
      Height          =   180
      Left            =   5400
      TabIndex        =   14
      Top             =   930
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      Height          =   180
      Left            =   6420
      TabIndex        =   13
      Top             =   930
      Width           =   1350
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "商標名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   1296
      Width           =   768
   End
   Begin VB.Label Label9 
      Caption         =   "特殊申請書:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5370
      Width           =   975
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "(1.延期 2.電子送件)"
      Height          =   180
      Left            =   1560
      TabIndex        =   10
      Top             =   5400
      Width           =   1515
   End
End
Attribute VB_Name = "frm030206_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/08/04 Form2.0已修改; Combo1、MSHFlexGrid1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Dim tm() As String
Dim intWhere As Integer
Dim intLastRow As Integer
Dim m_CP09 As String, m_CP10 As String, m_CP118 As String 'Added by Lydia 2018/11/14 選取的案件進度資料
Public m_RecSysNo As String 'Added by Lydia 2023/08/08 FCT/CFT申請書

Public Sub cmdOK_Click(Index As Integer)
 Dim i As Integer, bolChk As Boolean
   Select Case Index
      Case 0 '確定
         For i = 1 To MSHFlexGrid1.Rows - 1
            If MSHFlexGrid1.TextMatrix(i, 0) = "v" Then
               bolChk = True
               Me.Tag = MSHFlexGrid1.TextMatrix(i, 2)
               'Modified by Lydia 2018/11/14 選取的案件進度資料
               'tm(10) = MSHFlexGrid1.TextMatrix(i, 8)
               m_CP09 = "" & MSHFlexGrid1.TextMatrix(i, 2)
               m_CP10 = "" & MSHFlexGrid1.TextMatrix(i, 8)
               If "" & MSHFlexGrid1.TextMatrix(i, 9) <> "" Then
                   m_CP118 = "Y"
               Else
                   m_CP118 = ""
               End If
               strExc(1) = "" & MSHFlexGrid1.TextMatrix(i, 3)  'CPM03
               'end 2018/11/14
               Exit For
            End If
         Next
         If bolChk = False Then
            MsgBox "請選擇資料 !", vbInformation
            Exit Sub
         End If
                 
         'Added by Lydia 2022/12/19 補換發註冊證103,註冊商標分割308,註冊證副本314,註冊費717,復權729
         If strSrvDate(1) >= "20230101" And Text6 = "" And InStr("103,308,314,717,729", m_CP10) > 0 Then
            MsgBox "無紙本申請書可產生！", vbExclamation
            Exit Sub
         End If
         
         'end 2022/12/19
         '若有輸入特殊申請書
         If Text6 <> "" Then
            'Added by Lydia 2023/08/08
            If m_RecSysNo = "CFT" Then
               Select Case m_CP10
                  Case "304" '304英文證明書
                     frm030101_19.SetData 0, m_CP09, True, "1"
                     frm030101_19.Show
                     frm030101_19.QueryData
               End Select
            Else
            'end 2023/08/08
               Select Case Text6
                  Case "1" '延期
                     'Modified by Lydia 2018/11/14 tm(10)=> m_CP10
                     If m_CP10 = "201" Or m_CP10 = "202" Then '補正, 申復
                        frm03020601_1.Show
                     End If
                  'Added by Lydia 2018/11/14
                  Case "2" '電子送件
                     'Added by Lydia 2024/06/28 註冊後的分割申請書，智慧局目前只接受紙本申請書; ex.FCT-032966
                     If m_CP10 = "308" And Val(tm(21)) > 0 Then
                         MsgBox "註冊後的分割申請書，智慧局目前只接受紙本申請書！", vbInformation
                         Exit Sub
                     End If
                     'end 2024/06/28
                     
                     'Modified by Lydia 2019/02/26 +補優先權證明208
                     'If m_CP10 = "201" Then '補正
                     'Modified by Lydia 2019/03/26 增加補正電子送件書的案件性質
                     'If m_CP10 = "201" Or m_CP10 = "208" Then
                     'Modified by Lydia 2019/03/29 + 101申請,102延展,103補換發證書
                     'Modified by Lydia 2020/10/16 + 301變更,501移轉
                     'Modified by Lydia 2020/12/31 排除尚未有電子送件申請書的性質
                     'If InStr("101申請,102延展,103補換發證書,201補正,208補優先權證明,202申請意見書,206放棄專用權,210陳述意見書,211檢送同意書,303延期,305催審,310暫緩審理,313減縮商品,706其他,301變更,501移轉", m_CP10) > 0 Then
                     'Mark by Lydia 2021/02/05 開放第三階段：308註冊申請案分割、313註冊指定使用商品服務減縮、304英文證明書、502授權登記、717商標註冊費繳費單、725商標規費退費申請書(代辦退費
                     'If InStr("308註冊申請案分割、313註冊指定使用商品服務減縮、304英文證明書、502授權登記、717商標註冊費繳費單", m_CP10) = 0 Then
                          If m_CP118 <> "Y" Then
                               If MsgBox("此案非電子送件，是否確定產生電子送件申請書？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                                   Exit Sub
                               End If
                          End If
                          'Modified by Lydia 2019/02/21 定稿會因為勾選項而變更內容,所以(補正)電子送件申請書(frm03020605_1)併入紙本的畫面
                          'frm03020605_1.Show
                          'Modified by Lydia 2019/03/26 與阿蓮確認過電子送件和紙本用同一畫面
                          '(外商是先收文延期才做申請書，所以是從延期進度進入,為了區分特殊格式是選電子送件和延期在此註明)
                          'frm03020603_1.Show
                          GoTo JumpToNext
                     'Mark by Lydia 2021/02/05 開放第三階段
                     'Else
                     '     MsgBox strExc(1) & "尚未有電子送件申請書 !", vbInformation
                     '     Exit Sub
                     'End If
                     'end 2021/02/05
                  'end 2018/11/14
               End Select
            End If 'Added by Lydia 2023/08/08
         '若未輸入特殊申請書
         Else
            'Modified by Lydia 2018/11/14 tm(10)=> m_CP10
JumpToNext: 'Added by Lydia 2019/03/26
            Select Case m_CP10
               '申請意見書, 催審, 暫緩審理, 減縮商品, 其他
               'Case "202", "305", "310", "313", "706"
               '催審, 暫緩審理, 減縮商品, 其他
               'Modified by Lydia 2021/02/05 開放第三階段：308註冊申請案分割、313註冊指定使用商品服務減縮、304英文證明書、717商標註冊費繳費單、725商標規費退費申請書(代辦退費)
               'Modified by Lydia 2022/12/19 復權729
               'Modified by Lydia 2023/11/30 +自請撤回(306)、自請拋棄商標權(307)
               'Modified by Lydia 2024/08/06 +309中文證明書
               Case "305", "310", "313", "706", "308", "313", "304", "717", "725", "729", "306", "307", "309"
                   frm03020602_1.Show
               
               '補正, 補優先權證明, 申請意見書
               'Remove by Lydia 2020/12/31 統一在Case Else
               'Case "201", "208", "202"
               '    frm03020603_1.Show
               'end 2020/12/31
               '更正
               Case "302"
                   frm03020604_1.Show
               '延期
               Case "303"
                   frm03020601_1.Show
               'Added by Lydia 2019/03/29 申請, 延展, 補換發證書(紙本+電子送件)
               'Modified by Lydia 2022/12/19 + 314 申請註冊證副本
               Case "101", "102", "103", "314"
                   frm03020605_1.Show '取代原本(補正)電子送件申請書(frm03020605_1)表單編號
               'Added by Lydia 2020/10/16 變更,移轉
               'Modified by Lydia 2021/02/05 +授權205
               Case "301", "501", "502"
                   frm03020606_1.Show
               'end 2020/10/16
               
               '其他案件性質
               Case Else
                  'Modified by Lydia 2020/12/31 指定收文性質以外的A、B類收文，皆可產生補正申請書；若該性質沒有另外設計的畫面，則一律經由"各式申請書-補正,申請意見書"的畫面產生
                  'frm03020602_1.Show
                  'Added by Lydia 2021/01/28 (紙本)在舊畫面
                  If Text6 = "" And InStr("201,208,202", m_CP10) = 0 Then
                      frm03020602_1.Show
                  Else
                  'end 2021/01/28
                      frm03020603_1.Show
                  End If 'Added by Lydia 2021/01/212
            End Select
         End If
         cmdOK(1).SetFocus
         Me.Hide
         
      Case 1 '尋找
         Label4 = ""
         Label6 = ""
         If Text3 = "" Then Text3 = "0"
         If Text4 = "" Then Text4 = "00"
         tm(1) = Text1
         tm(2) = Text2
         tm(3) = Text3
         tm(4) = Text4
         
         'Modified by Lydia 2023/08/08
         'If tm(1) = "FCT" Then
         If tm(1) = m_RecSysNo Then
            If ClsPDReadTrademarkDatabase(tm(), intWhere) Then
               Label6.Caption = tm(15)
               Label4.Caption = tm(12)
               Text5.Text = tm(11)
            End If
         End If
         
         AddCboName Combo1, tm(5), tm(6), tm(7)
         'Modify By Sindy 2013/10/16 +and cp27 is null and cp57 is null
         'Modified by Lydia 2018/11/14 +cp118
         'Modified by Lydia 2023/08/08 +IIf(m_RecSysNo = "CFT", " and cp10='304'", "")
         strExc(0) = "select ''," & SQLDate("CP05") & ",cp09,cpm03,CP43,staff.st02 as st1,staff1.st02 as st2,cp64,cp10,cp118" & _
                      " from caseprogress, casepropertymap,staff,staff staff1" & _
                      " where " & ChgCaseprogress(tm(1) & tm(2) & tm(3) & tm(4)) & _
                      " and cp09<'C' and cp01=cpm01(+) and cp10=cpm02(+)" & _
                      " and cp14=staff.st01(+) and cp13=staff1.st01(+)" & _
                      " and cp27 is null and cp57 is null" & IIf(m_RecSysNo = "CFT", " and cp10='304'", "") & _
                      " order by CP05 desc"
         intI = 0
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
         GridHead
         
         '若只搜尋到一筆時直接勾選
         If Me.MSHFlexGrid1.Rows = 2 Then
            MSHFlexGrid1_Click
         End If
      
      Case 2
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
 Dim i As Integer
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = ""
      Next
   End With
End Sub

Private Sub Form_Initialize()
ReDim tm(1 To TF_TM) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'Added by Lydia 2023/08/08
   If m_RecSysNo = "CFT" Then
      Me.Caption = "CFT申請英文證明書"
      intWhere = 國外_CF
      Label9.Visible = False
      Label10.Visible = False
      Text6.Visible = False
      Text1 = m_RecSysNo
   Else
      Me.Caption = "各式申請書"
      Label9.Visible = True
      Label10.Visible = True
      Text6.Visible = True
   'end 2023/08/08
      intWhere = 國外_FC
   End If 'Added by Lydia 2023/08/08
   'Combo1.ListIndex = 0 'Remove by Lydia 2021/08/04
   Label4 = ""
   Label6 = ""
   'Modified by Lydia 2018/11/14
   'InitGrid 9, MSHFlexGrid1
   InitGrid 10, MSHFlexGrid1
   GridHead
   Text5.Text = strSrvDate(2)
   SendKeys "{Tab}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm030206_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
   'Added by Lydia 2018/11/15 若為電子送件,預設為2
   If MSHFlexGrid1.TextMatrix(intLastRow, 0) = "v" Then
        'Modified by Lydia 2019/02/26 +補優先權證明208
        'If "" & MSHFlexGrid1.TextMatrix(intLastRow, 8) = "201" And "" & MSHFlexGrid1.TextMatrix(intLastRow, 9) <> "" Then
        'Modified by Lydia 2018/03/26
        'If InStr("201,208", "" & MSHFlexGrid1.TextMatrix(intLastRow, 8)) > 0 And "" & MSHFlexGrid1.TextMatrix(intLastRow, 9) <> "" Then
        'Modified by Lydia 2109/03/29 +101申請,102延展,103補換發證書
        'Modified by Lydia 2020/10/16 + 301變更,501移轉
        'Modified by Lydia 2020/12/31 排除尚未有電子送件申請書的性質
        'If InStr("101申請,102延展,103補換發證書,201補正,208補優先權證明,202申請意見書,206放棄專用權,210陳述意見書,211檢送同意書,303延期,305催審,310暫緩審理,313減縮商品,706其他,301變更,501移轉", "" & MSHFlexGrid1.TextMatrix(intLastRow, 8)) > 0 And "" & MSHFlexGrid1.TextMatrix(intLastRow, 9) <> "" Then
        ''end 2019/04/11
        'Modified by Lydia 2021/02/05 開放第三階段
        'If InStr("308註冊申請案分割、313註冊指定使用商品服務減縮、304英文證明書、502授權登記、717商標註冊費繳費單", "" & MSHFlexGrid1.TextMatrix(intLastRow, 8)) = 0 And "" & MSHFlexGrid1.TextMatrix(intLastRow, 9) <> "" Then
        'Modified by Lydia 2023/08/08 CFT只有電子送件
        If "" & MSHFlexGrid1.TextMatrix(intLastRow, 9) <> "" Or Text1 = "CFT" Then
            Text6.Text = "2"
        ElseIf Text6.Text = "2" Then
            Text6.Text = ""
        End If
        'Added by Lydia 2024/06/28 註冊後的分割申請書，智慧局目前只接受紙本申請書; ex.FCT-032966
        If "" & MSHFlexGrid1.TextMatrix(intLastRow, 8) = "308" And Val(tm(21)) > 0 Then
            Text6.Text = ""
        End If
        'end 2024/06/28
   End If
   'end 2018/11/15
   
   cmdOK(0).SetFocus
End Sub

Private Sub Text1_GotFocus()
  TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   'Modifed by Lydia 2023/08/08
   'If Text1 <> "FCT" Then
   If Text1 <> m_RecSysNo Then
      MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
      TextInverse Text1
      Cancel = True
   End If
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1200: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1400: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1200: .Text = "相關總收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1200: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1400: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 1400: .Text = "進度備註"
      .col = 8: .ColWidth(8) = 0
      .col = 9: .ColWidth(9) = 0 'Added by Lydia 2018/11/14 電子送件
      .Visible = True
      If .Rows > 1 Then .row = 1
   End With
End Sub

Private Sub Text2_GotFocus()
  TextInverse Text2
End Sub

Private Sub Text3_GotFocus()
  TextInverse Text3
End Sub

Private Sub Text4_GotFocus()
  TextInverse Text4
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Lydia 2018/11/14 增加2.電子送件
   'If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
   If (KeyAscii > 51 Or KeyAscii < 49) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Public Sub ClearForm()
   '保留原輸入的系統類別
   Text2 = Empty
   Text3 = Empty
   Text4 = Empty
   Text5 = Empty
   Label4 = Empty
   Label6 = Empty
   Text6 = Empty
   Combo1.Clear
   'Modified by Lydia 2018/11/14
   'InitGrid 9, MSHFlexGrid1
   InitGrid 10, MSHFlexGrid1
   GridHead
   Text5.Text = strSrvDate(2)
   Text1.SetFocus
   Me.Text2.SetFocus
End Sub

'讀取案件性質
'Remove by Lydia 2018/11/14
'Private Function GetCP10(p_CP09 As String) As String
'   Dim stSQL As String, iRtn As Integer
'   If p_CP09 <> "" Then
'      stSQL = "select CP10 from caseprogress where cp09='" & p_CP09 & "'"
'      iRtn = 1
'      Set AdoRecordSet3 = ClsLawReadRstMsg(iRtn, stSQL)
'      If iRtn = 1 Then
'         GetCP10 = "" & AdoRecordSet3.Fields(0)
'      End If
'   End If
'End Function
'end 2018/11/14
