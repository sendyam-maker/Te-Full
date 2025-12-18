VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040132 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工密碼資料維護"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7680
   Begin VB.TextBox Text1 
      Height          =   540
      Index           =   3
      Left            =   2415
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   10
      Top             =   2400
      Width           =   4500
   End
   Begin VB.CheckBox chk 
      Caption         =   "Windows密碼同案件系統密碼"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   1680
      Width           =   3000
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   2415
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   2415
      MaxLength       =   6
      TabIndex        =   0
      Top             =   795
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   2415
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1056
      Top             =   -96
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040132.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040132.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040132.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040132.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040132.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040132.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040132.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040132.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040132.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040132.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040132.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   1085
      ButtonWidth     =   1138
      ButtonHeight    =   1032
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   3210
      TabIndex        =   15
      Top             =   803
      Width           =   1350
      ForeColor       =   192
      Caption         =   "Label2"
      Size            =   "2381;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   2415
      TabIndex        =   14
      Top             =   1410
      Width           =   1350
      Caption         =   "Label2"
      Size            =   "2381;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   2415
      TabIndex        =   13
      Top             =   1110
      Width           =   1350
      Caption         =   "Label2"
      Size            =   "2381;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PS：分所同仁請同時通知該所負責人員修改E-mail密碼！"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   6
      Left            =   840
      TabIndex        =   12
      Top             =   3120
      Width           =   4425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備註"
      Height          =   180
      Index           =   5
      Left            =   840
      TabIndex        =   11
      Top             =   2445
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Windows密碼"
      Height          =   180
      Index           =   4
      Left            =   840
      TabIndex        =   9
      Top             =   2085
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號"
      Height          =   180
      Index           =   0
      Left            =   825
      TabIndex        =   8
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   1
      Left            =   825
      TabIndex        =   7
      Top             =   1147
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工部門別"
      Height          =   180
      Index           =   2
      Left            =   825
      TabIndex        =   6
      Top             =   1447
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件系統密碼"
      Height          =   180
      Index           =   3
      Left            =   825
      TabIndex        =   5
      Top             =   1725
      Width           =   1080
   End
End
Attribute VB_Name = "frm12040132"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/14 改成Form2.0 ; Label2(index)
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Const iFieldTotal = 3 '畫面只顯示4個欄位 'Modify by Amy 2015/12/25 原:2
Dim RcMain As New ADODB.Recordset, cp As New ADODB.Recordset
Dim TmpField(0 To iFieldTotal) As String, ActionEdit As Integer
Dim Bmk As Variant, i As Integer
Dim strSP04 As String  'add by sonia 2016/1/29 記錄薪資查詢密碼

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

'2015/12/22 ADD BY SONIA
Private Sub chk_Click()
   If Me.Chk.Value = vbChecked Then
      Text1(2).Text = Text1(1).Text
   End If
End Sub
'2015/12/22 END

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   Select Case KeyCode
      Case vbKeyF2
         Text1(0).SetFocus
         RcEdit 0
      Case vbKeyF3
         Text1(1).SetFocus
         Text1(0).TabStop = False
         RcEdit 1
      Case vbKeyF5
         RcEdit 2
      Case vbKeyF4
         RcEdit 5
      Case vbKeyHome
         If Not (ActionEdit = 0 Or ActionEdit = 1) Then
            ActionRc 0
         End If
      Case vbKeyPageUp
         If Not (ActionEdit = 0 Or ActionEdit = 1) Then
             ActionRc 1
         End If
      Case vbKeyPageDown
         If Not (ActionEdit = 0 Or ActionEdit = 1) Then
             ActionRc 2
         End If
      Case vbKeyEnd
         If Not (ActionEdit = 0 Or ActionEdit = 1) Then
              ActionRc 3
         End If
      Case vbKeyF9
         Doit
      Case vbKeyF10
         RcEdit 4
      Case vbKeyEscape
         Unload Me
         Set frm12040132 = Nothing
   End Select
End Sub

'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到Private Sub Form_KeyPress(KeyAscii As Integer)
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
      Case vbKeyReturn:
         If ActionEdit <> 3 Then
            KeyAscii = 0
            Doit
         End If
    End Select
End Sub

Private Sub Form_Load()
   m_bInsert = IsUserHasRightOfFunction("frm12040132", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040132", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040132", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040132", strFind, False)
   
   MoveFormToCenter Me
   If RcMain.State = adStateOpen Then RcMain.Close
   '2015/12/22 MODIFY BY SONIA +SP02
   'Modify by Amy 2015/12/25 +SP06
   'modify by sonia 2016/1/29 +SP04
   strExc(0) = "SELECT SP01,SP03,SP02,SP06,SP04 FROM STAFF_PWD ORDER BY 1"
   cp.CursorLocation = adUseClient
   RcMain.CursorType = adOpenDynamic
   RcMain.CursorLocation = adUseClient
   RcMain.LockType = adLockBatchOptimistic
   RcMain.Open strExc(0), cnnConnection
   '移到首筆
   If Not RcMain.BOF Then ActionRc 0
   TxtSitu True
   ActionEdit = 3
   
   If m_bInsert Then
       TBar1.Buttons(1).Enabled = True
   Else
       TBar1.Buttons(1).Enabled = False
   End If
   If m_bUpdate Then
       TBar1.Buttons(2).Enabled = True
   Else
       TBar1.Buttons(2).Enabled = False
   End If
   If m_bDelete Then
       TBar1.Buttons(3).Enabled = True
   Else
       TBar1.Buttons(3).Enabled = False
   End If
End Sub

Private Sub ActionRc(ByVal Sty As Integer)
   TxtLock 2
   If RcMain.EOF And RcMain.BOF Then MsgBox "資料庫內無資料 !", vbInformation: Exit Sub
   With RcMain
      Select Case Sty
         Case 0
            .MoveFirst
         Case 1
            .MovePrevious
            If .BOF Then
               Beep
               DataErrorMessage (6)
               .MoveFirst
            End If
         Case 2
            .MoveNext
            If .EOF Then
               Beep
               DataErrorMessage (7)
               .MoveLast
            End If
         Case 3
            .MoveLast
      End Select
   End With
   SetTxtValue
End Sub

Private Sub SetTxtValue()
Dim i As Integer, j As Integer
   
   For i = 0 To iFieldTotal
      If IsNull(RcMain.Fields(i).Value) Then
         Text1(i).Text = ""
      ElseIf i = 1 Or i = 2 Then
         Text1(i).Text = Encrypt(RcMain.Fields(i).Value, False)
      Else
         Text1(i).Text = RcMain.Fields(i).Value
      End If
   Next i
   
   '2015/12/22 ADD BY SONIA
   If Text1(2).Text = Text1(1).Text Then
      Me.Chk.Value = vbChecked
   Else
      Me.Chk.Value = vbUnchecked
   End If
   '2015/12/22 END
   
   'add by sonia 2016/1/29
   strSP04 = ""
   If "" & RcMain.Fields(4).Value <> "" Then
      strSP04 = Encrypt(RcMain.Fields(4).Value, False)
   End If
   'end 2016/1/29
End Sub

Private Sub RcEdit(Situ As Integer)
Dim i As Integer
   
   Select Case Situ
      Case 0 'add
         TxtSitu False
         ActionEdit = 0
         TxtLock 2
         Me.Chk.Value = vbChecked
         TextInverse Text1(0)
      Case 1 'modi
         TxtSitu False
         ActionEdit = 1
         For i = 0 To iFieldTotal
            If i = 12 Then
            TmpField(i) = ChangeTStringToWString(Text1(i))
            Else
            TmpField(i) = Text1(i).Text
            End If
         Next
         Text1(0).Locked = True
      Case 2 'delete
         If MsgBox("是否要刪除此筆資料?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
            RcMain.Delete
            RcMain.UpdateBatch
            If RcMain.EOF = True Then
               ActionRc 1
            Else
               ActionRc 2
            End If
         End If
      Case 3 'update
         If ActionEdit = 0 Then
            If TxtValidate = False Then Exit Sub
            
            RcMain.AddNew
            If GetVal = False Then Exit Sub
            ActionRc 3
         ElseIf ActionEdit = 1 Then
            If TxtValidate = False Then Exit Sub
            
            If GetVal = False Then Exit Sub
         ElseIf ActionEdit = 2 Then
            RcMain.Find "SP01='" & Text1(0).Text & "'", 0, adSearchForward, 1
            If RcMain.EOF Then
               MsgBox "無此記錄之資料 !", vbCritical
               RcMain.Bookmark = Bmk
            End If
            SetTxtValue
         End If
         TxtSitu True
         ActionEdit = 3
      Case 4 'cancel
         If ActionEdit = 2 Then
            TxtSitu True
            RcMain.Bookmark = Bmk
            SetTxtValue
            ActionEdit = 3
         ElseIf MsgBox("妳並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
            TxtSitu True
            If ActionEdit = 0 Then
               ActionRc 3
            ElseIf ActionEdit = 1 Then
               For i = 0 To iFieldTotal
                  If i = 12 Then
                  Text1(i).Text = ChangeWStringToTString(TmpField(i))
                  Else
                  Text1(i).Text = TmpField(i)
                  End If
               Next
            End If
            ActionEdit = 3
         Else
            Exit Sub
         End If
         
      Case 5 'query
         Bmk = RcMain.Bookmark
         TxtSitu False
         TxtLock 2
         For i = 2 To iFieldTotal
            Text1(i).Locked = True
         Next
         ActionEdit = 2
         Text1(0).SetFocus
   End Select
End Sub

Private Function GetVal() As Boolean
Dim i As Integer

On Error GoTo ErrHand
   For i = 0 To iFieldTotal
      If i = 1 Or i = 2 Then
         RcMain.Fields(i).Value = Encrypt(Text1(i).Text, True)
      Else
         If Text1(i).Text <> "" Then
            RcMain.Fields(i).Value = Text1(i).Text
         Else
            RcMain.Fields(i).Value = Null
         End If
      End If
   Next
   RcMain.UpdateBatch
   GetVal = True
   Exit Function
ErrHand:
   GetVal = False
   RcMain.CancelUpdate
   RcMain.ReQuery
   If Err.Number = -2147217887 Then
      MsgBox "資料錯誤，新增失敗 !", vbInformation
   Else
      MsgBox "錯誤 : " & Err.Description, vbInformation
   End If
End Function

Private Sub TxtLock(ByVal Lt As Integer)
Dim txt As TextBox, i As Integer
   
   Select Case Lt
      Case 0
         For Each txt In frm12040132.Text1
            txt.Locked = True
         Next
      Case 1
         For Each txt In frm12040132.Text1
            txt.Locked = False
         Next
      Case 2
         For Each txt In frm12040132.Text1
            txt.Text = ""
         Next
   End Select
End Sub

Private Sub TxtSitu(ByVal TF As Boolean)
Dim i As Integer, txt As TextBox
   
   If TF = True Then
      TxtLock 0
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = True
         TBar1.Buttons(i + 5).Enabled = True
      Next
      TBar1.Buttons(11).Enabled = False
      TBar1.Buttons(12).Enabled = False
      TBar1.Buttons(14).Enabled = True
   Else
      TxtLock 1
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = False
         TBar1.Buttons(i + 5).Enabled = False
      Next
      TBar1.Buttons(11).Enabled = True
      TBar1.Buttons(12).Enabled = True
      TBar1.Buttons(14).Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040132 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   
   Select Case Button.Index
      Case 1
      '新增
         Text1(0).SetFocus
         RcEdit 0
      Case 2
      '修改
         Text1(0).TabStop = False
         Text1(1).SetFocus
         TextInverse Text1(1)
         RcEdit 1
      Case 3
      '刪除
         RcEdit 2
      Case 4
      '查詢
         RcEdit 5
      Case 6
      '首筆
         ActionRc 0
      Case 7
      '前一
         ActionRc 1
      Case 8
      '後一
         ActionRc 2
      Case 9
      '最後
         ActionRc 3
      Case 11
      '確定
          Doit
      Case 12
      '取消
         RcEdit 4
      Case 14
      '結束
         Unload Me
         Set frm12040132 = Nothing
   End Select
   
   If Button.Index <> 14 And Button.Index <> 1 And Button.Index <> 2 And Button.Index <> 3 And Button.Index <> 4 Then
      If m_bInsert Then
          TBar1.Buttons(1).Enabled = True
      Else
          TBar1.Buttons(1).Enabled = False
      End If
      If m_bUpdate Then
          TBar1.Buttons(2).Enabled = True
      Else
          TBar1.Buttons(2).Enabled = False
      End If
      If m_bDelete Then
          TBar1.Buttons(3).Enabled = True
      Else
          TBar1.Buttons(3).Enabled = False
      End If
   End If

End Sub

Private Sub Text1_Change(Index As Integer)

   Select Case Index
      Case 0
         If cp.State = adStateOpen Then cp.Close
         strExc(1) = "select ST02,A0902, DECODE(ST04,'2','離職',NULL) AS ST04 from staff,ACC090 where st01='" & Text1(0) & "' AND A0901=ST03"
         cp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
         If cp.EOF And cp.BOF Then
            Label2(0).Caption = ""
            Label2(1).Caption = ""
            Label2(2).Caption = ""
         Else
            If IsNull(cp.Fields(0)) Then
                Label2(0).Caption = ""
            Else
                Label2(0).Caption = cp.Fields(0).Value
            End If
            If IsNull(cp.Fields(1)) Then
                Label2(1).Caption = ""
            Else
                Label2(1).Caption = cp.Fields(1).Value
            End If
            If IsNull(cp.Fields(2)) Then
                Label2(2).Caption = ""
            Else
                Label2(2).Caption = "(" & cp.Fields(2).Value & ")"
            End If
         End If
         cp.Close
   End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

   Select Case Index
      Case 0
         KeyAscii = UpperCase(KeyAscii)
         '查詢Enter
         If KeyAscii = 13 And ActionEdit = 2 Then
            RcEdit 3
         End If
      '2015/12/22 ADD BY SONIA
      Case 2
         Me.Chk.Value = vbUnchecked
      '2015/12/22 END
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim strTmp As String, i As Integer

   '2015/12/22 ADD BY SONIA
   Select Case Index
      Case 1
         If Trim(Text1(1).Text) = "" Then
            MsgBox "密碼不可為空白!"
            Cancel = True
         Else
            If Me.Chk.Value = vbChecked Then Text1(2).Text = Text1(1).Text
            'add by sonia 2016/1/29
            If Text1(1).Text = strSP04 Then
               MsgBox "案件密碼不可與薪資密碼相同！"
               Cancel = True
               TextInverse Text1(Index)
            End If
            'end 2016/1/29
         End If
      Case 2
         If Text1(2).Text = Text1(1).Text Then
            Me.Chk.Value = vbChecked
         Else
            Me.Chk.Value = vbUnchecked
         End If
      'Add by Amy 2015/12/25 +備註
      Case 3
        If Trim(Text1(3)) <> MsgText(601) Then
            If Not CheckLengthIsOK(Text1(3), Text1(3).MaxLength) Then
               Cancel = True
            End If
        End If
   End Select
   '2015/12/22 END
   
'   If ActionEdit = 3 Then Exit Sub
'
'   Select Case Index
'
'      Case 1
'        '密碼檢查
'         If Trim(Text1(1).Text) = "" Then
'            MsgBox "密碼不可為空白!"
'            Cancel = True
'         Else
'            Cancel = False
'         End If
'   End Select
'   If Cancel Then TextInverse Text1(Index)
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   For Each objTxt In Text1
      If objTxt.Index <> 0 Then
         If objTxt.Enabled = True Then
            Cancel = False
            Text1_Validate objTxt.Index, Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
      End If
   Next
   
   TxtValidate = True
End Function

Sub Doit()

   If Text1(0) = "" Then MsgBox "員工代號不可為空值", vbInformation: Text1(0).SetFocus: Exit Sub
      
   '新增
   If ActionEdit = 0 Then
      If cp.State = adStateOpen Then cp.Close
      '檢查員工是否存在
      strExc(1) = "select count(*) from staff where ST01='" & Text1(0) & "'"
      cp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
      If cp.Fields(0).Value = "0" Then
         MsgBox "員工編號不存在!", vbInformation
         Text1(0).SetFocus
         TextInverse Text1(0)
         cp.Close
         Exit Sub
      End If
      If cp.State = adStateOpen Then cp.Close
      '檢查密碼資料是否存在
      strExc(1) = "select count(*) from staff_pwd where SP01='" & Text1(0) & "'"
      cp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
      If cp.Fields(0).Value <> "0" Then
         MsgBox "密碼資料已存在，不可再新增!", vbInformation
         Text1(0).SetFocus
         TextInverse Text1(0)
         cp.Close
         Exit Sub
      End If
   End If
   
   RcEdit 3
   Text1(0).TabStop = True
   RcMain.ReQuery
   RcMain.Find "sp01='" & Text1(0) & "'", 0, adSearchForward, 1
End Sub
