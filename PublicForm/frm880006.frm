VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm880006 
   BorderStyle     =   1  '單線固定
   Caption         =   "發明人資料"
   ClientHeight    =   5745
   ClientLeft      =   90
   ClientTop       =   990
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9345
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增發明人(&A)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4980
      TabIndex        =   16
      Top             =   50
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "移動順序:"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1710
      TabIndex        =   13
      Top             =   3300
      Width           =   2025
      Begin VB.CommandButton cmdDown 
         Caption         =   "▼"
         Height          =   255
         Left            =   1410
         TabIndex        =   15
         Top             =   90
         Width           =   375
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "▲"
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   90
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "取消"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   6630
      TabIndex        =   6
      Top             =   2790
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   7020
      TabIndex        =   7
      Top             =   50
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7980
      TabIndex        =   8
      Top             =   50
      Width           =   1200
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "全部取消"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   5325
      TabIndex        =   5
      Top             =   2790
      Width           =   912
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "全部選取"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4035
      TabIndex        =   4
      Top             =   2790
      Width           =   912
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "選取"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2730
      TabIndex        =   3
      Top             =   2790
      Width           =   912
   End
   Begin MSForms.TextBox TxtKey 
      Height          =   330
      Left            =   1200
      TabIndex        =   1
      Top             =   2820
      Width           =   1215
      VariousPropertyBits=   671105051
      Size            =   "2143;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstInventorTo 
      Height          =   2040
      Left            =   240
      TabIndex        =   2
      Top             =   3660
      Width           =   8835
      VariousPropertyBits=   746586139
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "15584;3598"
      MatchEntry      =   0
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstInventorFrom 
      Height          =   2040
      Left            =   240
      TabIndex        =   0
      Top             =   690
      Width           =   8835
      VariousPropertyBits=   746586139
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "15584;3598"
      MatchEntry      =   0
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "姓名檢索："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   2895
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "視窗上連續按二下亦可選取資料！"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   90
      TabIndex        =   11
      Top             =   60
      Width           =   3795
   End
   Begin VB.Label Label2 
      Caption         =   "已選區"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   270
      TabIndex        =   10
      Top             =   3300
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "待選區(可複選)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   270
      TabIndex        =   9
      Top             =   330
      Width           =   2145
   End
End
Attribute VB_Name = "frm880006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/15 改成Form2.0 ;  lstInventorFrom、lstInventorTo、TxtKey
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit

Public strPetition As String, strInventorNo As String
Public fmCallForm As Form 'Added by Morgan 2020/2/19

Private Sub ReadData()
Dim i As Integer, varSaveCursor, strInventorNoTemp1() As String, strInventorNoTemp2() As String, strInventorNameTemp() As String
Dim varPetitionTemp As Variant, strPetitionTemp(4) As String, j As Integer
Dim varInventorNo As Variant, varInventorName As Variant
'Add by Morgan 2004/7/27
Dim strInventorNationTemp() As String
Dim bolExists As Boolean 'Add by Sindy 2013/8/1
Dim strInventorID() As String 'Add By Sindy 2016/12/7

On Error GoTo ErrHand

varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
varPetitionTemp = Split(strPetition, ",")
'For i = 0 To 4
For i = 0 To UBound(varPetitionTemp)
       strPetitionTemp(i) = varPetitionTemp(i)
Next
'Modify by Morgan 2004/7/27
'加發明人國籍
'If obj003.ReadInventor(strPetitionTemp(), strInventorNoTemp1(), strInventorNoTemp2(), strInventorNameTemp()) Then
'Modify By Sindy 2016/12/7 + , strInventorID
If PUB_ReadInventor(strPetitionTemp(), strInventorNoTemp1(), strInventorNoTemp2(), strInventorNameTemp(), strInventorNationTemp, strInventorID) Then
   If strInventorNo = "" Then
      For i = 0 To UBound(strInventorNoTemp1)
         'Modify by Morgan 2004/7/27
         'lstInventorFrom.AddItem strInventorNoTemp1(i) + "-" + strInventorNoTemp2(i) + "   " + strInventorNameTemp(i)
         'Modify By Sindy 2016/12/7 + "   " + strInventorID(i)
         lstInventorFrom.AddItem strInventorNoTemp1(i) + "-" + strInventorNoTemp2(i) + "   " + strInventorNameTemp(i) + "   " + strInventorNationTemp(i) + "   " + strInventorID(i)
      Next
   Else
      varInventorNo = Split(strInventorNo, ",")
      'Modify By Sindy 2010/3/8
'      For i = 0 To UBound(strInventorNoTemp1)
'             For j = 0 To UBound(varInventorNo)
'                    If strInventorNoTemp1(i) + strInventorNoTemp2(i) = varInventorNo(j) Then
'                       Exit For
'                    End If
'             Next
'             If j = UBound(varInventorNo) + 1 Then
'                  'Modify by Morgan 2004/7/27
'                  'lstInventorFrom.AddItem strInventorNoTemp1(i) + "-" + strInventorNoTemp2(i) + "   " + strInventorNameTemp(i)
'                  lstInventorFrom.AddItem strInventorNoTemp1(i) + "-" + strInventorNoTemp2(i) + "   " + strInventorNameTemp(i) + "   " + strInventorNationTemp(i)
'             Else
'                  'Modify by Morgan 2004/7/27
'                  'lstInventorTo.AddItem strInventorNoTemp1(i) + "-" + strInventorNoTemp2(i) + "   " + strInventorNameTemp(i)
'                  lstInventorTo.AddItem strInventorNoTemp1(i) + "-" + strInventorNoTemp2(i) + "   " + strInventorNameTemp(i) + "   " + strInventorNationTemp(i)
'             End If
'      Next
      For j = 0 To UBound(varInventorNo)
         If varInventorNo(j) <> "" Then
            bolExists = False 'Add by Sindy 2013/8/1
            For i = 0 To UBound(strInventorNoTemp1)
               If strInventorNoTemp1(i) + strInventorNoTemp2(i) = varInventorNo(j) Then
                  bolExists = True 'Add by Sindy 2013/8/1
                  Exit For
               End If
            Next i
            If bolExists = True Then 'Add by Sindy 2013/8/1 +if
               lstInventorTo.AddItem strInventorNoTemp1(i) + "-" + strInventorNoTemp2(i) + "   " + strInventorNameTemp(i) + "   " + strInventorNationTemp(i) + "   " + strInventorID(i)
            End If
         End If
      Next j
      For i = 0 To UBound(strInventorNoTemp1)
            For j = 0 To UBound(varInventorNo)
               If varInventorNo(j) <> "" Then
                  If strInventorNoTemp1(i) + strInventorNoTemp2(i) = varInventorNo(j) Then
                     Exit For
                  End If
               End If
            Next
            If j = UBound(varInventorNo) + 1 Then
               lstInventorFrom.AddItem strInventorNoTemp1(i) + "-" + strInventorNoTemp2(i) + "   " + strInventorNameTemp(i) + "   " + strInventorNationTemp(i) + "   " + strInventorID(i)
            End If
      Next
      '2010/3/8 End
   End If
End If
CheckClear
Screen.MousePointer = varSaveCursor
Exit Sub
ErrHand:
Screen.MousePointer = varSaveCursor
ErrorMsg
Unload Me
End Sub

Private Sub CheckClear()
If lstInventorFrom.ListCount = 0 Then
   cmdMove(0).Enabled = False
   cmdMove(1).Enabled = False
Else
   cmdMove(0).Enabled = True
   cmdMove(1).Enabled = True
End If
If lstInventorTo.ListCount = 0 Then
   cmdMove(2).Enabled = False
   cmdMove(3).Enabled = False
Else
   cmdMove(2).Enabled = True
   cmdMove(3).Enabled = True
End If
End Sub

'Added by Morgan 2020/2/19
Private Sub cmdAdd_Click()
   Dim oForm As Form
   If PUB_CheckFormExist("frm050709") Then
      Forms(0).GetForm ("")
      MsgBox "客戶發明人資料維護畫面正在使用中，請自行操作！"
      Exit Sub
   End If
   
   Me.Hide
   Set oForm = Forms(0).GetForm("frm050709")
   
   With oForm
   Set .fmCallForm = fmCallForm
   .bAddOnly = True
   .UseDatamaintain vbKeyF2 '新增
   .Text1(0) = strPetition
   .Show
   End With
   fmCallForm.Enabled = False
   
End Sub

Private Sub cmdMove_Click(Index As Integer)
   Select Case Index
      Case 0
         MoveOne lstInventorFrom, lstInventorTo
      Case 1
         MoveAll lstInventorFrom, lstInventorTo
      Case 2
         MoveAll lstInventorTo, lstInventorFrom
      Case 3
         MoveOne lstInventorTo, lstInventorFrom
   End Select
   CheckClear
End Sub

'Modified by Lydia 2022/02/15 ListBox=>Control
Private Sub MoveOne(ByRef lstTempFrom As Control, lstTempTo As Control)
Dim i As Integer

Do
       If lstTempFrom.Selected(i) Then
          lstTempTo.AddItem lstTempFrom.List(i)
          lstTempFrom.RemoveItem i
          i = i - 1
       End If
       i = i + 1
Loop Until i = lstTempFrom.ListCount
End Sub

'Modified by Lydia 2022/02/15 ListBox=>Control
Private Sub MoveAll(ByRef lstTempFrom As Control, lstTempTo As Control)
Dim i As Integer

For i = 0 To lstTempFrom.ListCount - 1
       lstTempFrom.Selected(i) = True
Next
MoveOne lstTempFrom, lstTempTo
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer

If Index = 0 Then
   If lstInventorTo.ListCount = 0 Then
      strInventorNo = ""
   Else
   'ElseIf lstInventorTo.ListCount <= 10 Then
      strInventorNo = ""
      For i = 0 To lstInventorTo.ListCount - 2
             strInventorNo = strInventorNo + Left(lstInventorTo.List(i), 8) + Mid(lstInventorTo.List(i), 10, 2) + ","
      Next
      strInventorNo = strInventorNo + Left(lstInventorTo.List(i), 8) + Mid(lstInventorTo.List(i), 10, 2)
   'Modify By Sindy 2014/11/6 '發明人不鎖10個限制
'   Else
'      'Modify By Sindy 2013/1/31
'      'ShowMsg MsgText(9202)
'      MsgBox "你選擇的發明人超過了10個。" & vbCrLf & _
'             "系統內只能記錄10個發明人，超過部分請程序記錄在案件備註！", vbCritical + vbOKOnly, MsgText(9001)
'      '2013/1/31 End
'      Exit Sub
   End If
End If
Unload Me
End Sub

'Add By Sindy 2016/8/25 向上移
Private Sub cmdUp_Click()
Dim ii As Integer
   
   'Modified by Lydia 2022/02/15 取得Form2.0的ListBox選取數量
   'If lstInventorTo.ListCount > 0 And lstInventorTo.SelCount = 1 Then
   ii = GetListSelCount(lstInventorTo)
   If lstInventorTo.ListCount > 0 And ii = 1 Then
   'end 2022/02/15
      For ii = 0 To lstInventorTo.ListCount - 1
         If lstInventorTo.Selected(ii) = True Then
            If ii = 0 Then
               Exit Sub
            Else
               strExc(0) = ii
               strExc(1) = lstInventorTo.List(ii)
               lstInventorTo.Selected(ii) = False
               Exit For
            End If
         End If
      Next ii
      If strExc(0) > 0 Then
         lstInventorTo.List(strExc(0)) = lstInventorTo.List(strExc(0) - 1)
         lstInventorTo.List(strExc(0) - 1) = strExc(1)
         lstInventorTo.Selected(strExc(0) - 1) = True
      End If
   Else
      MsgBox "欲移動資料項目，請選擇一筆資料！", vbCritical + vbOKOnly, MsgText(9001)
   End If
End Sub

'Add By Sindy 2016/8/25 向下移
Private Sub cmdDown_Click()
Dim ii As Integer
   'Modified by Lydia 2022/02/15 取得Form2.0的ListBox選取數量
   'If lstInventorTo.ListCount > 0 And lstInventorTo.SelCount = 1 Then
   ii = GetListSelCount(lstInventorTo)
   If lstInventorTo.ListCount > 0 And ii = 1 Then
   'end 2022/02/15
      For ii = 0 To lstInventorTo.ListCount - 1
         If lstInventorTo.Selected(ii) = True Then
            If ii = lstInventorTo.ListCount - 1 Then
               Exit Sub
            Else
               strExc(0) = ii
               strExc(1) = lstInventorTo.List(ii)
               lstInventorTo.Selected(ii) = False
               Exit For
            End If
         End If
      Next ii
      If strExc(0) < lstInventorTo.ListCount - 1 Then
         lstInventorTo.List(strExc(0)) = lstInventorTo.List(strExc(0) + 1)
         lstInventorTo.List(strExc(0) + 1) = strExc(1)
         lstInventorTo.Selected(strExc(0) + 1) = True
      End If
   Else
      MsgBox "欲移動資料項目，請選擇一筆資料！", vbCritical + vbOKOnly, MsgText(9001)
   End If
End Sub

Private Sub Form_Activate()
ReadData

'Added by Morgan 2020/2/13
If TypeName(fmCallForm) = "frm040104_3" Then
   If IsUserHasRightOfFunction("frm050709", strAdd, False) Then
      cmdAdd.Visible = True
   Else
      cmdAdd.Visible = False
   End If
End If
'end 2020/2/13

End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
'   Set frm880006 = Nothing
End Sub

'Add By Sindy 2010/3/8
'Modified by Lydia 2022/02/15 改成Form2.0
'Private Sub lstInventorFrom_DblClick()
Private Sub lstInventorFrom_DblClick(Cancel As MSForms.ReturnBoolean)
   MoveOne lstInventorFrom, lstInventorTo
   CheckClear
End Sub

'Add By Sindy 2010/3/8
'Modified by Lydia 2022/02/15 改成Form2.0
'Private Sub lstInventorTo_DblClick()
Private Sub lstInventorTo_DblClick(Cancel As MSForms.ReturnBoolean)
   MoveOne lstInventorTo, lstInventorFrom
   CheckClear
End Sub

'Added by Lydia 2016/12/30 輸入關鍵字後，移動到模糊比對的第一筆
Private Sub TxtKey_LostFocus()
Dim inQ As Integer
Dim tmpKey  As String
Dim bSearch As Boolean

   TxtKey.Text = Trim(PUB_RepToOneSpace(TxtKey.Text))
   
   If TxtKey.Text <> "" And lstInventorFrom.ListCount > 0 Then
      tmpKey = UCase(TxtKey)
      For inQ = 0 To lstInventorFrom.ListCount - 1
         If InStr(UCase(lstInventorFrom.List(inQ)), tmpKey) > 0 Then
            bSearch = True
            Exit For
         End If
      Next inQ
   End If
   
   If bSearch Then
      lstInventorFrom.ListIndex = inQ
      lstInventorFrom.Selected(inQ) = True
      lstInventorFrom.SetFocus
   End If
End Sub

'Added by Lydia 2016/12/30
Private Sub txtkey_GotFocus()
   TextInverse TxtKey
End Sub

'Added by Lydia 2022/02/15 取得Form2.0的ListBox選取數量
Private Function GetListSelCount(ByRef oLBox As Control) As Integer
Dim intX As Integer
    
    GetListSelCount = 0
    If oLBox.ListCount > 0 Then
        For intX = 0 To oLBox.ListCount - 1
            If oLBox.Selected(intX) = True Then
                GetListSelCount = GetListSelCount + 1
            End If
        Next intX
    End If
End Function
