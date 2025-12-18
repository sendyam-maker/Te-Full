VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100121_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工姓名查詢員工資料"
   ClientHeight    =   3920
   ClientLeft      =   1040
   ClientTop       =   2490
   ClientWidth     =   4920
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3920
   ScaleWidth      =   4920
   Begin VB.CheckBox Check2 
      Caption         =   "在職所內員工"
      Height          =   225
      Left            =   390
      TabIndex        =   8
      Top             =   3180
      Width           =   1845
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1380
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1980
      Width           =   525
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含眷屬姓名做查詢"
      Height          =   225
      Left            =   2820
      TabIndex        =   1
      Top             =   630
      Width           =   1845
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1380
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1170
      Width           =   825
   End
   Begin VB.ComboBox cboDepName 
      Height          =   300
      Left            =   1380
      Style           =   2  '單純下拉式
      TabIndex        =   3
      Top             =   1560
      Width           =   2120
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1860
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2745
      Width           =   492
   End
   Begin VB.CommandButton CmdOk 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3735
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   36
      Width           =   756
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2940
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   36
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2700
      MaxLength       =   7
      TabIndex        =   6
      Top             =   2355
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1380
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2355
      Width           =   800
   End
   Begin MSForms.TextBox txtName 
      Height          =   300
      Left            =   1380
      TabIndex        =   0
      Top             =   585
      Width           =   1400
      VariousPropertyBits=   683687963
      MaxLength       =   10
      Size            =   "2469;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
      Height          =   180
      Left            =   1950
      TabIndex        =   19
      Top             =   2010
      Width           =   2055
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "所　　別："
      Height          =   180
      Left            =   390
      TabIndex        =   18
      Top             =   2010
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "( 可模糊比對 )"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1380
      TabIndex        =   17
      Top             =   900
      Width           =   1110
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   390
      TabIndex        =   16
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "員工姓名若為兩個字者，中間請加一全形空白查詢!!!"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   390
      TabIndex        =   15
      Top             =   3660
      Width           =   4140
   End
   Begin VB.Line Line2 
      X1              =   2385
      X2              =   2505
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否含離職人員：           （Y：含離職）"
      Height          =   180
      Left            =   390
      TabIndex        =   14
      Top             =   2790
      Width           =   3135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "到  職  日："
      Height          =   180
      Left            =   390
      TabIndex        =   13
      Top             =   2385
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "部　　門："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   390
      TabIndex        =   12
      Top             =   1605
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工姓名："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   390
      TabIndex        =   11
      Top             =   630
      Width           =   900
   End
End
Attribute VB_Name = "frm100121_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/22 改成Form2.0(txt1(0)改為txtName)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/20 日期欄已修改
Option Explicit

Dim s As Integer, i As Integer, j As Integer
Dim StrTag As String, strSql As String
Dim m_bln_KeyinValid As Boolean
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer


'92.04.16 nick
Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
          cmdState = -1
          'Modify By Sindy 2012/6/20
          If Check2.Value = 0 Then '若無勾選在職所內員工時,則至少輸入一項查詢條件
          '2012/6/20 End
            'Modify By Sindy 2012/6/14 + And Len(Trim(txt1(5).Text)) = 0
            'modify by sonia 2022/1/22 改Form2.0(txt1(0)改為txtName)
            'If Len(Trim(Me.txt1(0).Text)) = 0 And
            If Len(Trim(Me.txtName.Text)) = 0 And _
              Len(Trim(Me.cboDepName.Text)) = 0 And _
              Len(Trim(txt1(1).Text)) = 0 And _
              Len(Trim(txt1(2).Text)) = 0 And _
              Len(Trim(txt1(3).Text)) = 0 And _
              Len(Trim(txt1(4).Text)) = 0 And _
              Len(Trim(txt1(5).Text)) = 0 Then
               s = MsgBox("請檢查是否有必要條件忘了輸入．．．．", , "輸入條件不足")
               'modify by sonia 2022/1/22 改Form2.0(txt1(0)改為txtName)
               'Me.txt1(0).SetFocus
               Me.txtName.SetFocus
               Exit Sub
            End If
          End If
          'add by sonia 2022/1/22
          If Me.txtName.Text <> "" Then
            ' 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
            If PUB_ChkUniText(Me, , True, "TextBox") = False Then
                Exit Sub
            End If
          End If
          'end 2022/1/22
          If Me.txt1(1).Text <> "" Then
            txt1_LostFocus 1
            If m_bln_KeyinValid = False Then txt1_GotFocus 1: Exit Sub
          End If
          If Me.txt1(2).Text <> "" Then
            txt1_LostFocus 2
            If m_bln_KeyinValid = False Then txt1_GotFocus 2: Exit Sub
          End If
          If Me.txt1(3).Text <> "" Then
            txt1_LostFocus 3
            If m_bln_KeyinValid = False Then txt1_GotFocus 3: Exit Sub
          End If
          'Add By Sindy 2012/6/14
          If Me.txt1(5).Text <> "" Then
            txt1_LostFocus 5
            If m_bln_KeyinValid = False Then txt1_GotFocus 5: Exit Sub
          End If
          '2012/6/14 End
         
          Me.Enabled = False
          If fnSaveParentForm(Me) = False Then
              Me.Enabled = True
              Exit Sub
          End If
            Screen.MousePointer = vbHourglass
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/16 清除查詢印表記錄檔欄位
            frm100121_2.Show
            frm100121_2.StrMenu
            Screen.MousePointer = vbDefault
          Me.Enabled = True
      Case 1
           fnCloseAllFrm100
      Case Else
   End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
   Exit Sub
End Sub

Private Sub Form_Load()
Dim m_bQuery As Boolean
   
   bolToEndByNick = False
   MoveFormToCenter Me
   SetComboData
   '92.04.16 nick
   cmdState = -1
   
   'Add By Sindy 2018/1/26 有”櫃檯每日信件輸入”權限的人才可以使用”含眷屬姓名做查詢”查詢功能
   m_bQuery = IsUserHasRightOfFunction("frm010016", strFind, False)
   If (m_bQuery = True Or Pub_StrUserSt03 = "M21") And Pub_StrUserSt03 <> "M31" Then '不可為財務處,因財務處ST05也為00
      Check1.Visible = True
      Check1.Tag = "T"
   Else
      Check1.Visible = False
      Check1.Tag = "F"
   End If
   '2018/1/26 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100121_2 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   
   'Add By Cheng 2002/01/09
   Select Case Index
      Case 0 '員工姓名
         '   Me.txt1(Index).IMEMode = 1
         'edit by nickc 2007/06/06 切換輸入法改用API
         OpenIme
      'add by sonia 2014/10/29
      Case Else
         CloseIme
      'end 2014/10/29
   End Select
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   'Add By Sindy 2012/6/14
   If Index = 5 Then '所別
      KeyAscii = Pub_NumAscii(KeyAscii)
   Else
   '2012/6/14 End
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   m_bln_KeyinValid = False
   Select Case Index
'cancel by sonia 2022/1/22 改Form2.0(txt1(0)改為txtName)
'      Case 0 '員工姓名
'            'Add By Cheng 2002/01/09
'         '   Me.txt1(Index).IMEMode = 0
'         'edit by nickc 2007/06/06 切換輸入法改用API
'         CloseIme
'end 2022/1/22
      Case 1 '到職日起
            If Me.txt1(Index) <> "" Then
               If CheckIsTaiwanDate(Me.txt1(Index)) = False Then
                  Me.txt1(Index).SetFocus
                  Exit Sub
               End If
            End If
      Case 2 '到職日迄
            If Me.txt1(Index) <> "" Then
               If CheckIsTaiwanDate(Me.txt1(Index)) = False Then
                  Me.txt1(Index).SetFocus
                  Exit Sub
               End If
               'Add By Cheng 2002/06/10
               If Val(Me.txt1(1).Text) > Val(Me.txt1(2).Text) Then
                  MsgBox "到職日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  Me.txt1(1).SetFocus
                  Exit Sub
               End If
            End If
      Case 3 '是否含離職人員
           If InStr(1, "Yy ", txt1(Index)) = 0 Then
               s = MsgBox("請輸入 Y 或空白!!", , "輸入錯誤")
               txt1(Index).SetFocus
               Exit Sub
           End If
      'Add By Sindy 2012/6/14
      Case 5 '所別
            If txt1(Index) <> "" Then
               If Val(txt1(Index)) < 1 Or Val(txt1(Index)) > 5 Then
                  s = MsgBox("請輸入1,2,3,4,5或空白!!", , "輸入錯誤")
                  txt1(Index).SetFocus
                  Exit Sub
               End If
            End If
      Case Else
   End Select
   m_bln_KeyinValid = True
End Sub

Private Sub SetComboData()
'宣告變數
Dim Rs As New ADODB.Recordset

   '2014/2/11 modify by sonia 除電腦中心及人事處外,其他人只能看到有在職員工的部門(王副總提需求江總同意)
   'rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' Order By A0901", _
            cnnConnection, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2023/12/22
   If strSrvDate(1) >= 新部門啟用日 Then
      Call SetST93Combo(cboDepName)
   Else
   '2023/12/22 END
      Me.cboDepName.Clear
      Rs.CursorLocation = adUseClient
      If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M21" Then
         Rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' Order By A0901", _
                  cnnConnection, adOpenStatic, adLockReadOnly
      Else
         Rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' and a0901<>'P29' and a0901 in (select distinct st03 from staff where st04='1' and st01>'6' and substr(st01,1,1)<'G' and substr(st01,4,1)<>'9') Order By A0901", _
                  cnnConnection, adOpenStatic, adLockReadOnly
      End If
      '2014/2/11 end
      Me.cboDepName.AddItem ""
      While Not Rs.EOF
         Me.cboDepName.AddItem Left(Rs.Fields(0).Value & Space(5), 5) & Rs.Fields(1).Value
         Rs.MoveNext
      Wend
      If Rs.State <> adStateClosed Then Rs.Close
      Set Rs = Nothing
   End If
End Sub

'add by sonia 2022/1/22
Private Sub txtName_GotFocus()
   txtName.SelStart = 0
   txtName.SelLength = Len(txtName)
End Sub

Private Sub txtName_LostFocus()
   ' 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
      txtName_GotFocus
      txtName.SetFocus
   End If
End Sub
'end 2022/1/22
