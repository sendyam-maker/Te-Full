VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100101_13 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶減免身分查詢"
   ClientHeight    =   4430
   ClientLeft      =   2750
   ClientTop       =   4620
   ClientWidth     =   8210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4430
   ScaleWidth      =   8210
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   3888
      Left            =   96
      TabIndex        =   8
      Top             =   504
      Width           =   8088
      _ExtentX        =   14270
      _ExtentY        =   6862
      _Version        =   393216
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
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "回前畫面"
      Height          =   400
      Index           =   1
      Left            =   7050
      TabIndex        =   7
      Top             =   90
      Width           =   1125
   End
   Begin VB.TextBox txtFn 
      Height          =   270
      Index           =   3
      Left            =   5040
      MaxLength       =   3
      TabIndex        =   4
      Top             =   150
      Width           =   525
   End
   Begin VB.TextBox txtFn 
      Height          =   270
      Index           =   2
      Left            =   4410
      MaxLength       =   3
      TabIndex        =   3
      Top             =   150
      Width           =   525
   End
   Begin VB.TextBox txtFn 
      Height          =   270
      Index           =   1
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   2
      Top             =   150
      Width           =   945
   End
   Begin VB.TextBox txtFn 
      Height          =   270
      Index           =   0
      Left            =   990
      MaxLength       =   8
      TabIndex        =   1
      Top             =   150
      Width           =   945
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6060
      TabIndex        =   0
      Top             =   90
      Width           =   912
   End
   Begin VB.Line Line3 
      X1              =   4710
      X2              =   5280
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家:"
      Height          =   180
      Index           =   5
      Left            =   3600
      TabIndex        =   6
      Top             =   180
      Width           =   765
   End
   Begin VB.Line Line2 
      X1              =   1770
      X2              =   2190
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號:"
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   5
      Top             =   180
      Width           =   765
   End
End
Attribute VB_Name = "frm100101_13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/13 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/12/09 改成Form2.0 ; grdList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
Option Explicit

'使用者權限設定
Dim bolInsert As Boolean
Dim bolUpdate As Boolean
Dim bolDelete As Boolean
Dim bolSelect As Boolean
'列印控制
Dim strSql As String


Sub cmQueryData()
Call cmdQuery_Click(0)
End Sub

Private Sub cmdQuery_Click(Index As Integer)
Select Case Index
Case 0
   If TxtValidate = False Then Exit Sub
   '查詢
   grdList.Rows = 1
   If CheckQueryData = True Then
      Screen.MousePointer = vbHourglass
      grdList.MousePointer = flexHourglass
      If QueryData() = False Then
         MsgBox "無資料", vbOKOnly, "查詢資料"
         txtFn(0).SetFocus
      End If
      grdList.MousePointer = flexDefault
      Screen.MousePointer = vbDefault
   End If
Case 1
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case Else
End Select
End Sub

Private Sub Form_Load()
  
   MoveFormToCenter Me
   Me.Show
   setAuthority
    txtFn(0).Text = ""
    txtFn(1).Text = ""
    txtFn(2).Text = "000"
    txtFn(3).Text = "999"
   Call InitGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100101_13 = Nothing
End Sub

Private Sub txtFn_GotFocus(Index As Integer)
   If txtFn(Index).Locked = False Then
      TextInverse txtFn(Index)
      If txtFn(Index).Locked = False Then
         'edit by nickc 2007/06/06
         'txtFn(Index).IMEMode = 2
         CloseIme
      End If
   End If
End Sub

Private Sub txtFn_KeyPress(Index As Integer, KeyAscii As Integer)
   If txtFn(Index).Locked = False Then
      KeyAscii = UpperCase(KeyAscii)
      Select Case Index
         Case 0, 1
         '只可為文數字
            If Not (KeyAscii = 8 Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 47 And KeyAscii < 58)) Then
               KeyAscii = 0
            End If
         Case 2, 3
         '數字
            If Not (KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
                KeyAscii = 0
            End If
      End Select
   End If
End Sub

Private Sub txtFn_LostFocus(Index As Integer)
      Dim bolCancel As Boolean
      bolCancel = False
      Call txtFn_Validate(Index, bolCancel)
      If bolCancel = True Then
         txtFn(Index).SetFocus
      End If
End Sub

Private Function QueryData() As Boolean
   Dim strSql As String, rsQuery As New ADODB.Recordset
   Dim strCon As String
   
On Error GoTo ErrHand

   strCon = ""
   If txtFn(0) <> "" Then
      strCon = strCon & " AND ad01>='" & Mid(GetNewFagent2(txtFn(0)), 1, 8) & "' "
   End If
   If txtFn(1) <> "" Then
      strCon = strCon & " AND ad01<='" & Mid(GetNewFagent2(txtFn(1)), 1, 8) & "' "
   End If
   pub_QL05 = ";客戶編號：" & Mid(GetNewFagent2(txtFn(0)), 1, 8) & "-" & Mid(GetNewFagent2(txtFn(1)), 1, 8) & "(減免身份)" 'Add By Sindy 2025/8/27
   If txtFn(2) <> "" Then
      strCon = strCon & " AND ad02>='" & txtFn(2) & "'"
   End If
   If txtFn(3) <> "" Then
      strCon = strCon & " AND ad02<='" & txtFn(3) & "'"
   End If
   'Add By Sindy 2025/8/27
   If txtFn(2) <> "" Or txtFn(3) <> "" Then
      pub_QL05 = pub_QL05 & ";申請國家：" & txtFn(2) & "-" & txtFn(3)
   End If
   '2025/8/27 END
   
   'Modify By Sindy 2012/5/24 decode(cu15,'0','個人','1','公司','')==>decode(cu15,'0','個人','1','公司','2','學校','3','特殊機構','')
   'Modified by Morgan 2023/10/6 +日本減免身分
   strSql = "select ad01,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),decode(cu15,'0','個人','1','公司','2','學校','3','特殊機構',''),cu11,na03,decode(ad03,'Y','是','N','否',''),decode(ad02,'011',decode(ad10,'1','中小企業','2','獨資企業','3','小企業','4','新興企業','5','大學','6','個人'),decode(ad10,'1','自然人','2','學校','3','中小企業')),replace(ad11||'-'||ad12||'-'||ad13||'-'||ad14,'---',''),na01,ad03,ad10" & _
            " from ApplicantDiscount, customer , nation " & _
            " where  ad01=cu01(+) and '0'=cu02(+) and ad02=na01(+)  " & strCon & " ORDER BY ad01,ad02"
            
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
InitGrid
   If rsQuery.RecordCount > 0 Then
      If pub_QL04 <> "" Then InsertQueryLog (rsQuery.RecordCount) 'Add By Sindy 2025/8/27
      QueryData = True
      Call UpdateGridList(rsQuery)
   Else
      If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/27
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Exit Function
   
ErrHand:

   MsgBox Err.Description, vbCritical
            
End Function

Private Sub UpdateGridList(ByRef rsTmp As ADODB.Recordset)

   Dim iRow As Integer, iCol As Integer
   grdList.Rows = 1
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      With grdList
         .Rows = .Rows + 1
         iRow = .Rows - 1
         For iCol = 1 To grdList.Cols - 1
            .TextMatrix(iRow, iCol) = "" & rsTmp.Fields(iCol - 1).Value
         Next iCol
      End With
      rsTmp.MoveNext
   Loop
   grdList.FixedRows = 1  'Added by Lydia 2023/10/13
End Sub

Private Sub InitGrid()

   Dim arrGridHeadText, arrGridHeadWidth
   Dim iCol As Integer

   arrGridHeadText = Array("", "客戶編號", "名　　稱", "個人/公司", "ID" _
                     , "申請國家", "可否減免", "身份", "減免證明存卷案號", "", "", "")

   arrGridHeadWidth = Array(300, 1100, 2000, 1000 _
                     , 1100, 1200, 1000, 1300, 1500, 0, 0, 0)

   With grdList
      .row = 0
      .Cols = UBound(arrGridHeadText) + 1
      For iCol = 0 To .Cols - 1
         .col = iCol
         .Text = arrGridHeadText(iCol)
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .CellAlignment = flexAlignCenterCenter
      Next
      .Rows = 1
   End With
   grdList.FixedRows = 0  'Added by Lydia 2023/10/13
   
End Sub

Private Function TxtValidate() As Boolean

   Dim oText As TextBox, bolCancel As Boolean, arrText, oMaskEdBox As MaskEdBox
   
   TxtValidate = False

       For Each oText In txtFn
         If oText.Locked = False Then
            txtFn_Validate oText.Index, bolCancel
            If bolCancel = True Then
               oText.SetFocus
               TextInverse oText
               Exit For
            End If
         End If
      Next
   If bolCancel = False Then TxtValidate = True
End Function

'檢查查詢條件
Private Function CheckQueryData() As Boolean
   Dim bolCancel As Boolean, i As Integer
   
   If txtFn(0).Text = "" Then
        MsgBox "請輸入客戶編號起!!!", vbExclamation + vbOKOnly
        txtFn(0).SetFocus
        Exit Function
   End If
   If txtFn(1).Text = "" Then
        MsgBox "請輸入客戶編號迄!!!", vbExclamation + vbOKOnly
        txtFn(1).SetFocus
        Exit Function
   End If
   
   For i = 0 To 3
      Call txtFn_Validate(i, bolCancel)
      If bolCancel = True Then
         txtFn(i).SetFocus
         Exit Function
      End If
   Next
   CheckQueryData = True
End Function

'使用者權限設定
Private Sub setAuthority()
      bolInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
      bolUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
      bolDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
      bolSelect = IsUserHasRightOfFunction(Me.Name, strFind, False)
End Sub

Private Sub txtFn_Validate(Index As Integer, Cancel As Boolean)
   If txtFn(Index).Locked = False Then
      Select Case Index
         Case 0
         Case 1
               If Mid(UCase(txtFn(0)), 1, 6) <> Mid(UCase(txtFn(1)), 1, 6) And (txtFn(1) < txtFn(0)) And txtFn(1) <> "" Then
                  MsgBox "客戶編號迄值必需大於起值，且前 6 碼要相同！", vbCritical
                  Cancel = True
               End If
         Case 3
            If txtFn(2) <> "" And txtFn(3) < txtFn(2) Then
               MsgBox "申請國家迄值必需大於起值！", vbCritical
               Cancel = True
            End If
      End Select
      If Cancel = True Then txtFn_GotFocus (Index)
   End If
End Sub
