VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090904 
   BorderStyle     =   1  '單線固定
   Caption         =   "工程師各式申請書"
   ClientHeight    =   5750
   ClientLeft      =   130
   ClientTop       =   2410
   ClientWidth     =   9360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   9360
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   30
      MaxLength       =   1
      TabIndex        =   17
      Text            =   "3"
      Top             =   30
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   3048
      TabIndex        =   15
      Top             =   528
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "FCP"
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
      TabIndex        =   6
      Top             =   576
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7464
      TabIndex        =   4
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
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4050
      Left            =   120
      TabIndex        =   14
      Top             =   1620
      Width           =   9075
      _ExtentX        =   15998
      _ExtentY        =   7161
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "(3.電子送件)"
      Height          =   180
      Left            =   450
      TabIndex        =   18
      Top             =   60
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   960
      TabIndex        =   16
      Top             =   1260
      Width           =   8070
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "14235;503"
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
      TabIndex        =   13
      Top             =   576
      Width           =   768
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期:"
      Height          =   180
      Left            =   5400
      TabIndex        =   12
      Top             =   576
      Visible         =   0   'False
      Width           =   948
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號"
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   936
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   180
      Left            =   960
      TabIndex        =   10
      Top             =   936
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   5400
      TabIndex        =   9
      Top             =   930
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      Height          =   180
      Left            =   6420
      TabIndex        =   8
      Top             =   936
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   1296
      Width           =   768
   End
End
Attribute VB_Name = "frm090904"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/24 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB、Combo1
'Create By Sindy 2017/11/15
Option Explicit

Dim pa() As String
Dim intWhere As Integer
Dim intLastRow As Integer


Public Sub cmdok_Click(Index As Integer)
Dim i As Integer, bolChk As Boolean
Dim stConCP As String
Dim bolShowForm As Boolean
   
   Select Case Index
      Case 0 '確定
         For i = 1 To MSHFlexGrid1.Rows - 1
            If MSHFlexGrid1.TextMatrix(i, 0) = "v" Then
               bolChk = True
               Me.Tag = MSHFlexGrid1.TextMatrix(i, 2)
               pa(10) = MSHFlexGrid1.TextMatrix(i, 8)
               Exit For
            End If
         Next
         If bolChk = False Then
            MsgBox "請選擇資料 !", vbInformation
            Exit Sub
         Else
            Call PUB_FCPChkCP141(Me.Tag) 'Add By Sindy 2024/4/25 檢查暫不送件,指定送件日
         End If
         
         Select Case pa(10)
            'Add By Sindy 2022/5/3
            Case 改請衍生設計
               frm06010301_1.SetParent Me
               frm06010301_1.Show
            '204修正,203主動修正
            '107 再審申請
            '431 高速審查:修正,審查
            '422 加速審查
            '307 分割
            '408面詢,407請求面詢
            '416 實體審查
            '205 申復
            '433 誤譯訂正
            'Modify By Sindy 2019/10/8 + 敏莉:開放"403更改"工程師也可產生"其他申復事項申請書"
            'Modify By Sindy 2019/10/30 + 敏莉:開放"402更正"工程師可產生"專利更正申請書"
            'Modify By Sindy 2022/4/13 + 敏莉:239擇一申復
            'Modified by Morgan 2022/5/12 +435續行母案再審--陳亭妙
            'Modify By Sindy 2023/8/11 + 敏莉:230提供情報
            'Modify By Sindy 2024/5/28 + 敏莉:202補文件
            'Modify By Sindy 2024/8/20 + 亭妙:421=申請技術報告
            'Modify By Sindy 2025/2/18 修正:補充說明 是工程師操作, 產生專利補正文件申請書
            Case 修正, 主動修正, 107, 431, 422, 分割, 實體審查, 申復, 433, 面詢, 請求面詢, _
                 更改, 更正, 239, 435, 230, 補文件, 421, 補充說明
               frm090904_1.Hide
               Call frm090904_1.ReadData(bolShowForm)
               If bolShowForm = True Then
                  frm090904_1.Show
               Else
                  Exit Sub
               End If
            'Add By Sindy 2023/2/16
            'Modified by Morgan 2024/11/18 +447再審查加速審查
            Case 其他, 447
               frm06010310_1.SetParent Me
               frm06010310_1.Show
               frm06010310_1.Caption = "各式申請書-電子送件-其他"
               '2023/2/16 END
            Case Else
               MsgBox "點選的案件性質目前尚無電子送件申請書！"
               Exit Sub
         End Select
         cmdOK(1).SetFocus
         Me.Hide
      Case 1 '尋找
         Label4 = ""
         Label6 = ""
         If Text3 = "" Then Text3 = "0"
         If Text4 = "" Then Text4 = "00"
         pa(1) = Text1
         pa(2) = Text2
         pa(3) = Text3
         pa(4) = Text4

         If pa(1) = "FCP" Then
            If ClsPDReadPatentDatabase(pa(), intWhere) Then
               Label6.Caption = pa(22)
               Label4.Caption = pa(11)
               Text5.Text = pa(10)
            End If
         ElseIf pa(1) = "FG" Then
            If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then
               Text5.Text = pa(10)
               Label4.Caption = pa(11)
            End If
         End If

         AddCboName Combo1, pa(5), pa(6), pa(7)
         
         If Pub_StrUserSt03 <> "M51" And UCase(pub_DbTerminalName) = UCase(正式資料庫電腦名稱) Then
            stConCP = " AND CP01 IN ('FCP','FG') AND (CP14='" & strUserNum & "' or INSTR(staff.ST52||','||staff.ST53||','||staff.ST54,'" & strUserNum & "')>0)"
         End If
         '未發文未取消收文 : and CP158=0 and CP159=0
         strExc(0) = "select ''," & SQLDate("CP05") & ",cp09,cpm03,CP43,staff.st02 as st1,staff1.st02 as st2," & _
            "cp64,cp10,cp118 from caseprogress, casepropertymap,staff,staff staff1" & _
            " where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
            " and cp09<'C' and CP158=0 and CP159=0" & _
            " and cp01=cpm01(+) and cp10=cpm02(+)" & _
            " and cp14=staff.st01(+) and cp13=staff1.st01(+)" & stConCP
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
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   'Combo1.ListIndex = 0 'Remove by Lydia 2021/09/24 Form 2.0  沒有的屬性
   Label4 = ""
   Label6 = ""
   InitGrid 10, MSHFlexGrid1
   GridHead
   Text5.Text = strSrvDate(2)
   SendKeys "{Tab}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090904 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
'   'Add by Sindy 2018/7/30
'   If intLastRow > 0 Then
'      If MSHFlexGrid1.TextMatrix(intLastRow, 0) = "v" Then
'         If MSHFlexGrid1.TextMatrix(intLastRow, 9) = "Y" Then
'            Text6.Text = "1"
'         Else
'            Text6.Text = ""
'         End If
'      End If
'   End If
'   '2018/7/30 END
   cmdOK(0).SetFocus
End Sub

Private Sub Text1_GotFocus()
  TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 <> "FCP" And Text1 <> "FG" Then
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
      .col = 9: .ColWidth(9) = 0: .Text = "電子送件"
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

Public Sub ClearForm()
   '保留原輸入的系統類別
'   Text1 = Empty
   Text2 = Empty
   Text3 = Empty
   Text4 = Empty
   Text5 = Empty
   Label4 = Empty
   Label6 = Empty
   Combo1.Clear
   InitGrid 10, MSHFlexGrid1
   GridHead
   Text5.Text = strSrvDate(2)
   Text1.SetFocus
   Me.Text2.SetFocus
End Sub
