VERSION 5.00
Begin VB.Form frm050206 
   BorderStyle     =   1  '單線固定
   Caption         =   "互惠代理人目標給案未輸入明細表"
   ClientHeight    =   1410
   ClientLeft      =   3045
   ClientTop       =   1515
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4110
   Begin VB.TextBox txt1 
      Height          =   276
      Left            =   1035
      MaxLength       =   3
      TabIndex        =   5
      Top             =   570
      Width           =   600
   End
   Begin VB.TextBox txt2 
      Height          =   276
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   4
      Top             =   930
      Width           =   375
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   3192
      TabIndex        =   1
      Top             =   36
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   2412
      TabIndex        =   0
      Top             =   36
      Width           =   756
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "年度：                    (民國年)"
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   630
      Width           =   2100
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "期間：               ( 1:上半年 2:下半年 )"
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   2820
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "(1. 管制表 2. 定稿)"
      Height          =   180
      Left            =   5955
      TabIndex        =   3
      Top             =   5325
      Width           =   1425
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "列印格式:"
      Height          =   180
      Left            =   4170
      TabIndex        =   2
      Top             =   5340
      Width           =   765
   End
End
Attribute VB_Name = "frm050206"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/14 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
'Create by Morgan 2008/4/15
Option Explicit

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 1
         Screen.MousePointer = vbHourglass
         If TxtValidate Then
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/3 清除查詢印表記錄檔欄位
            Process
         End If
         Screen.MousePointer = vbDefault
      Case 2
         Unload Me
   End Select
End Sub

Private Sub Process()
   pub_QL05 = pub_QL05 & ";" & Left(Label63(2), 3) & txt1 'Add By Sindy 2010/12/3
   If txt2 = "1" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label63(3), 3) & "1:上半年" 'Add By Sindy 2010/12/3
   Else
      pub_QL05 = pub_QL05 & ";" & Left(Label63(3), 3) & "2:下半年" 'Add By Sindy 2010/12/3
   End If
   'modify by sonia 2017/6/5 +FC06='CFP'
   strExc(0) = "select NA03 C0,FC01||DECODE(FC03,NULL,'','-'||FC03) C1" & _
      ",NVL(FA05,NVL(FA06,FA04)) C2,NVL(PCC03,NVL(PCC04,PCC05)) C3" & _
      ",FC07,FC08 From FAGENTCONFIG, fagenttarget, FAGENT, Nation, POTCUSTCONT" & _
      " where FC06='CFP' and fc04=" & txt1 & " and fc05='" & txt2 & "' and ft15(+)=fc15 and ft01 is null" & _
      " and FA01(+)=FC01 AND FA02(+)='0'" & _
      " AND NA01(+)=SUBSTR(FA10,1,3) AND PCC01(+)=FC01 AND PCC02(+)=FC03" & _
      " ORDER BY NA01 ASC,C1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2010/12/3
      With frm050206_1
         .Show
         .lblYear = txt1
         .lblPeriod = txt2
         .grdDataList.Visible = False
         .SetDataListWidth
         Set .grdDataList.Recordset = RsTemp.Clone
         .SetDataListWidth True
         .grdDataList.Visible = True
      End With
      Me.Hide
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/3
      MsgBox "無資料！"
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   If txt1 = "" Then
      MsgBox "年度不可空白！"
      txt1.SetFocus
      Exit Function
   End If
   If txt2 = "" Then
      MsgBox "期間不可空白！"
      txt2.SetFocus
      Exit Function
   End If
   TxtValidate = True
End Function
'
Private Sub Form_Load()
   MoveFormToCenter Me
   txt1 = strSrvDate(2) \ 10000
   txt2 = 1 + ((strSrvDate(2) \ 100) Mod 100) \ 7
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm050206 = Nothing
End Sub

Private Sub txt1_GotFocus()
   TextInverse txt1
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txt2_GotFocus()
   TextInverse txt2
End Sub

Private Sub txt2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Chr(KeyAscii) <> 1 And Chr(KeyAscii) <> 2 Then
      KeyAscii = 0
      Beep
   End If
End Sub
