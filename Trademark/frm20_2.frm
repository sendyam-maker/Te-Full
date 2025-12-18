VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm20_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "商品名稱查詢"
   ClientHeight    =   5730
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   9135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9135
   Begin VB.CommandButton cmd 
      Caption         =   "最後筆(&L)"
      Height          =   348
      Index           =   5
      Left            =   5220
      TabIndex        =   12
      Top             =   180
      Visible         =   0   'False
      Width           =   996
   End
   Begin VB.CommandButton cmd 
      Caption         =   "下一筆(&N)"
      Height          =   348
      Index           =   4
      Left            =   4200
      TabIndex        =   11
      Top             =   180
      Visible         =   0   'False
      Width           =   996
   End
   Begin VB.CommandButton cmd 
      Caption         =   "前一筆(&P)"
      Height          =   348
      Index           =   3
      Left            =   3195
      TabIndex        =   10
      Top             =   180
      Visible         =   0   'False
      Width           =   996
   End
   Begin VB.CommandButton cmd 
      Caption         =   "第一筆(&F)"
      Height          =   348
      Index           =   2
      Left            =   2190
      TabIndex        =   9
      Top             =   180
      Visible         =   0   'False
      Width           =   996
   End
   Begin VB.CommandButton cmd 
      Caption         =   "回前畫面(&X)"
      Height          =   348
      Index           =   0
      Left            =   7920
      TabIndex        =   5
      Top             =   15
      Width           =   1170
   End
   Begin VB.CommandButton cmd 
      Caption         =   "確定(&B)"
      Default         =   -1  'True
      Height          =   348
      Index           =   1
      Left            =   6930
      TabIndex        =   4
      Top             =   15
      Width           =   996
   End
   Begin VB.Label lblTMN05 
      Caption         =   "流水號"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   6300
      TabIndex        =   15
      Top             =   420
      Width           =   675
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   4
      Left            =   6960
      TabIndex        =   14
      Top             =   390
      Width           =   1020
      VariousPropertyBits=   671105055
      Size            =   "1799;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   1460
      Index           =   3
      Left            =   45
      TabIndex        =   3
      Top             =   4170
      Width           =   9060
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "15981;2575"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   1460
      Index           =   2
      Left            =   45
      TabIndex        =   2
      Top             =   2430
      Width           =   9060
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "15981;2575"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   1460
      Index           =   1
      Left            =   45
      TabIndex        =   1
      Top             =   690
      Width           =   9060
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "15981;2575"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1065
      TabIndex        =   0
      Top             =   45
      Width           =   540
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "952;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商品日文名稱："
      Height          =   180
      Index           =   3
      Left            =   60
      TabIndex        =   13
      Top             =   3960
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商品英文名稱："
      Height          =   180
      Index           =   2
      Left            =   15
      TabIndex        =   8
      Top             =   2220
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商品中文名稱："
      Height          =   180
      Index           =   1
      Left            =   45
      TabIndex        =   7
      Top             =   465
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國際分類："
      Height          =   180
      Index           =   0
      Left            =   105
      TabIndex        =   6
      Top             =   90
      Width           =   900
   End
End
Attribute VB_Name = "frm20_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/24 改成Form2.0 ; Text1(index)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Public m_strFormStatus As String '1:新增 2:修改 3:快顯
Public m_row As Double '記錄列的位置

Private Sub cmd_Click(Index As Integer)
   Select Case Index
   Case 0 '離開
      Unload Me
   Case 1 '確定
      Select Case Me.m_strFormStatus
      Case "1" '新增
         If TxtValidate = False Then Exit Sub
         If CheckDataRepeat = False Then
            If SaveData Then
               clearControls
               Me.Text1(0).SetFocus
            End If
         End If
      Case "2" '修改
         If TxtValidate = False Then Exit Sub
         If CheckDataRepeat = False Then
            If SaveData Then
               Unload Me
            End If
         End If
      End Select
   Case 2 '第一筆
      If frm20.msgList.Rows > 1 Then
         'edit by nick 2004/10/12
         'DisplayProperty frm20.msgList.TextMatrix(1, 0), frm20.msgList.TextMatrix(1, 1), frm20.msgList.TextMatrix(1, 2)
         'Modified by Lydia 2022/03/15 +流水號 frm20.msgList.TextMatrix(1, 4)
         DisplayProperty frm20.msgList.TextMatrix(1, 1), frm20.msgList.TextMatrix(1, 2), frm20.msgList.TextMatrix(1, 3), frm20.msgList.TextMatrix(1, 4), frm20.msgList.TextMatrix(1, 5)
         Me.m_row = 1
      End If
   Case 3 '前一筆
      If frm20.msgList.Rows > 1 Then
         If Me.m_row - 1 > 0 Then
            Me.m_row = Me.m_row - 1
            'edit by nick 2004/10/12
            'DisplayProperty frm20.msgList.TextMatrix(Me.m_row, 0), frm20.msgList.TextMatrix(Me.m_row, 1), frm20.msgList.TextMatrix(Me.m_row, 2)
            'Modified by Lydia 2022/03/15 +流水號 frm20.msgList.TextMatrix(1, 4)
            DisplayProperty frm20.msgList.TextMatrix(Me.m_row, 1), frm20.msgList.TextMatrix(Me.m_row, 2), frm20.msgList.TextMatrix(Me.m_row, 3), frm20.msgList.TextMatrix(Me.m_row, 4), frm20.msgList.TextMatrix(Me.m_row, 5)
         Else
            MsgBox "已至第一筆!!!", vbExclamation + vbOKOnly
         End If
      End If
   Case 4 '後一筆
      If frm20.msgList.Rows > 1 Then
         If Me.m_row + 1 < frm20.msgList.Rows - 1 Then
            Me.m_row = Me.m_row + 1
            'edit by nick 2004/10/12
            'DisplayProperty frm20.msgList.TextMatrix(Me.m_row, 0), frm20.msgList.TextMatrix(Me.m_row, 1), frm20.msgList.TextMatrix(Me.m_row, 2)
            'Modified by Lydia 2022/03/15 +流水號 frm20.msgList.TextMatrix(1, 4)
            DisplayProperty frm20.msgList.TextMatrix(Me.m_row, 1), frm20.msgList.TextMatrix(Me.m_row, 2), frm20.msgList.TextMatrix(Me.m_row, 3), frm20.msgList.TextMatrix(Me.m_row, 4), frm20.msgList.TextMatrix(Me.m_row, 5)
         Else
            MsgBox "已至最後筆!!!", vbExclamation + vbOKOnly
         End If
      End If
   Case 5 '最後筆
      If frm20.msgList.Rows > 1 Then
         'edit by nick 2004/10/12
         'DisplayProperty frm20.msgList.TextMatrix(frm20.msgList.Rows - 1, 0), frm20.msgList.TextMatrix(frm20.msgList.Rows - 1, 1), frm20.msgList.TextMatrix(frm20.msgList.Rows - 1, 2)
         'Modified by Lydia 2022/03/15 +流水號  frm20.msgList.TextMatrix(frm20.msgList.Rows - 1, 4)
         'DisplayProperty frm20.msgList.TextMatrix(frm20.msgList.Rows - 1, 0), frm20.msgList.TextMatrix(frm20.msgList.Rows - 1, 1), frm20.msgList.TextMatrix(frm20.msgList.Rows - 1, 2), frm20.msgList.TextMatrix(frm20.msgList.Rows - 1, 3)
         DisplayProperty frm20.msgList.TextMatrix(frm20.msgList.Rows - 1, 1), frm20.msgList.TextMatrix(frm20.msgList.Rows - 1, 2), frm20.msgList.TextMatrix(frm20.msgList.Rows - 1, 3), frm20.msgList.TextMatrix(frm20.msgList.Rows - 1, 4), frm20.msgList.TextMatrix(frm20.msgList.Rows - 1, 5)
         Me.m_row = frm20.msgList.Rows - 1
      End If
   End Select
End Sub

Private Function SaveData() As Boolean
Dim strUpdStatus As String '0:none 1:Begin 2:Commit
   
   On Error GoTo ErrorHandler
   SaveData = False
   strUpdStatus = "0"
   cnnConnection.BeginTrans
   strUpdStatus = "1"
   '新增
   If Me.m_strFormStatus = "1" Then
      'Modify By Cheng 2002/09/16
'      strSQL = "INSERT INTO TRADEMARKMERCHANDISENAME VALUES('" & Me.Text1(0).Text & "','" & Me.Text1(1).Text & "','" & Me.Text1(2).Text & "'  )"
      'edit by nick 2004/10/12
      'strSQL = "INSERT INTO TRADEMARKMERCHANDISENAME VALUES('" & Me.Text1(0).Text & "','" & ChgSQL(Me.Text1(1).Text) & "','" & ChgSQL(Me.Text1(2).Text) & "'  )"
      'Modified by Lydia 2022/03/15
      'strSql = "INSERT INTO TRADEMARKMERCHANDISENAME (tmn01,tmn02,tmn03,tmn04) VALUES('" & Me.Text1(0).Text & "','" & ChgSQL(Me.Text1(1).Text) & "','" & ChgSQL(Me.Text1(2).Text) & "' ,'" & ChgSQL(Me.Text1(3).Text) & "'  )"
      strSql = "INSERT INTO TRADEMARKMERCHANDISENAME (tmn01,tmn02,tmn03,tmn04,tmn05) VALUES ('" & Me.Text1(0).Text & "','" & ChgSQL(Me.Text1(1).Text) & "','" & ChgSQL(Me.Text1(2).Text) & "' ,'" & ChgSQL(Me.Text1(3).Text) & "','" & GetMaxNo & "'  )"
      cnnConnection.Execute strSql
   '修改
   ElseIf Me.m_strFormStatus = "2" Then
      'Added by Lydia 2022/03/15 +流水號tmn05
      If Me.Text1(4).Text <> "" Then
        strSql = " tmn05=" & Me.Text1(4)
      Else
      'end 2022/03/15
        strSql = IIf(Me.Text1(0).Tag = "", " TMN01 IS NULL ", " TMN01='" & Me.Text1(0).Tag & "' ")
        strSql = strSql & IIf(Me.Text1(1).Tag = "", " AND TMN02 IS NULL ", " AND TMN02='" & ChgSQL(Me.Text1(1).Tag) & "' ")
        strSql = strSql & IIf(Me.Text1(2).Tag = "", " AND TMN03 IS NULL ", " AND TMN03='" & ChgSQL(Me.Text1(2).Tag) & "' ")
        'add by nick 2004/10/12
        strSql = strSql & IIf(Me.Text1(3).Tag = "", " AND TMN04 IS NULL ", " AND TMN04='" & ChgSQL(Me.Text1(3).Tag) & "' ")
      End If 'Added by Lydia 2022/03/15
      'Modify By Cheng 2002/09/16
'      strSQL = "UPDATE TRADEMARKMERCHANDISENAME SET TMN01='" & Me.Text1(0).Text & "', TMN02='" & Me.Text1(1).Text & "' ,TMN03='" & Me.Text1(2).Text & "'  WHERE " & strSQL
      'edit by nick 2004/10/12
      'strSQL = "UPDATE TRADEMARKMERCHANDISENAME SET TMN01='" & Me.Text1(0).Text & "', TMN02='" & ChgSQL(Me.Text1(1).Text) & "' ,TMN03='" & ChgSQL(Me.Text1(2).Text) & "'  WHERE " & strSQL
      strSql = "UPDATE TRADEMARKMERCHANDISENAME SET TMN01='" & Me.Text1(0).Text & "', TMN02='" & ChgSQL(Me.Text1(1).Text) & "' ,TMN03='" & ChgSQL(Me.Text1(2).Text) & "' ,TMN04='" & ChgSQL(Me.Text1(3).Text) & "'  WHERE " & strSql
      cnnConnection.Execute strSql
   End If
   cnnConnection.CommitTrans
   strUpdStatus = "2"
   SaveData = True
   Exit Function
ErrorHandler:
   If strUpdStatus = "1" Then
      cnnConnection.RollbackTrans
      If Err.Number <> 0 Then MsgBox "(" & Err.Number & ")" & Err.Description, vbExclamation + vbOKOnly, "更新動作失敗"
   End If
End Function

Private Function CheckDataRepeat() As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

On Error GoTo ErrorHandler

StrSQLa = IIf(Me.Text1(0).Text = "", " TMN01 IS NULL ", " TMN01 = '" & Me.Text1(0).Text & "' ")
StrSQLa = StrSQLa & IIf(Me.Text1(1).Text = "", " AND TMN02 IS NULL ", " AND TMN02 = '" & ChgSQL(Me.Text1(1).Text) & "' ")
StrSQLa = StrSQLa & IIf(Me.Text1(2).Text = "", " AND TMN03 IS NULL ", " AND TMN03 = '" & ChgSQL(Me.Text1(2).Text) & "' ")
'edit by nick 2004/10/12
StrSQLa = StrSQLa & IIf(Me.Text1(3).Text = "", " AND TMN04 IS NULL ", " AND TMN04 = '" & ChgSQL(Me.Text1(3).Text) & "' ")
StrSQLa = "Select * From TRADEMARKMERCHANDISENAME WHERE " & StrSQLa
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   CheckDataRepeat = True
   MsgBox "資料重覆, 請重新輸入!!!", vbExclamation + vbOKOnly
   Me.Text1(0).SetFocus
   Text1_GotFocus 0
Else
   CheckDataRepeat = False
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
Exit Function

ErrorHandler:
   MsgBox "(" & Err.Number & ")" & Err.Description
   CheckDataRepeat = True
End Function

Private Sub clearControls()
   Me.Text1(0).Text = Empty
   Me.Text1(1).Text = Empty
   Me.Text1(2).Text = Empty
   'add by nick 2004/10/12
   Me.Text1(3).Text = Empty
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   DisplayProperty
   'Added by Lydia 2022/03/15
   If Pub_StrUserSt03 <> "M51" Then
      lblTMN05.Visible = False
      Text1(4).Visible = False
   End If
   'end 2022/03/15
End Sub
   
'Modified by Lydia 2022/03/15 + 流水號strTMN05
Public Sub DisplayProperty(Optional ByVal strTMN01 As String, Optional ByVal strTMN02 As String, Optional ByVal strTMN03 As String, Optional ByVal strTMN04 As String, Optional ByVal strTMN05 As String)
   
   Select Case Me.m_strFormStatus
   Case "1"
      Me.Caption = "商品名稱查詢--新增"
      Me.cmd(2).Visible = False
      Me.cmd(3).Visible = False
      Me.cmd(4).Visible = False
      Me.cmd(5).Visible = False
   Case "2"
      Me.Caption = "商品名稱查詢--修改"
      Me.Text1(0).Text = "" & strTMN01
      Me.Text1(1).Text = "" & strTMN02
      Me.Text1(2).Text = "" & strTMN03
      'add by nick 2004/10/12
      Me.Text1(3).Text = "" & strTMN04
      Me.Text1(0).Tag = "" & strTMN01
      Me.Text1(1).Tag = "" & strTMN02
      Me.Text1(2).Tag = "" & strTMN03
      'add by nick 2004/10/12
      Me.Text1(3).Tag = "" & strTMN04
      Me.cmd(2).Visible = False
      Me.cmd(3).Visible = False
      Me.cmd(4).Visible = False
      Me.cmd(5).Visible = False
   Case "3"
      Me.Caption = "商品名稱查詢--快顯"
      Me.cmd(1).Visible = False
      Me.Text1(0).Text = "" & strTMN01
      Me.Text1(1).Text = "" & strTMN02
      Me.Text1(2).Text = "" & strTMN03
      'add by nick 2004/10/12
      Me.Text1(3).Text = "" & strTMN04
      Me.cmd(2).Visible = True
      Me.cmd(3).Visible = True
      Me.cmd(4).Visible = True
      Me.cmd(5).Visible = True
   End Select
   
   'Added by Lydia 2022/03/15
   If Me.m_strFormStatus <> "1" Then
       Me.Text1(4).Text = "" & strTMN05
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frm20.Show
   If Me.m_strFormStatus = "1" Or Me.m_strFormStatus = "2" Then
      frm20.QueryData
   End If
   Set frm20_2 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Me.Text1(Index)
   Select Case Index
   Case 1
      If Me.m_strFormStatus <> "3" Then
         'edit by nickc 2007/06/06 切換輸入法改用API
         'Me.Text1(Index).IMEMode = 1
         OpenIme
      End If
   End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
   Case 1
      'edit by nickc 2007/06/06 切換輸入法改用API
      'Me.Text1(Index).IMEMode = 2
      CloseIme
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 1
      'edit by nick 2004/09/15
      'If Len(Me.Text1(Index).Text) > 400 Then
      '   MsgBox "商品中文名稱輸入的資料請勿超過 400 個字!!!", vbExclamation + vbOKOnly
      'edit by nick 2004/09/20
      'If Len(Me.Text1(Index).Text) > 600 Then
      '   MsgBox "商品中文名稱輸入的資料請勿超過 600 個字!!!", vbExclamation + vbOKOnly
      If Len(Me.Text1(Index).Text) > 1000 Then
         MsgBox "商品中文名稱輸入的資料請勿超過 1000 個字!!!", vbExclamation + vbOKOnly
         Cancel = True
      End If
   Case 2
      If Len(Me.Text1(Index).Text) > 700 Then
         MsgBox "商品英文名稱輸入的資料請勿超過700個字!!!", vbExclamation + vbOKOnly
         Cancel = True
      End If
   'add by nick 2004/10/12
   Case 3
      If Len(Me.Text1(Index).Text) > 1000 Then
         MsgBox "商品日文名稱輸入的資料請勿超過1000個字!!!", vbExclamation + vbOKOnly
         Cancel = True
      End If
   End Select
   If Cancel = True Then Text1_GotFocus Index
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   'edit by nick 2004/10/12
   'If Me.Text1(1).Text = "" And Me.Text1(2).Text = ""  Then
   If Me.Text1(1).Text = "" And Me.Text1(2).Text = "" And Me.Text1(3).Text = "" Then
      MsgBox "請至少輸入一項商品名稱!!!", vbExclamation + vbOKOnly
      Me.Text1(1).SetFocus
      Text1_GotFocus 1
      Exit Function
   End If
   For Each objTxt In Me.Text1
      If objTxt.Enabled = True Then
         Cancel = False
         Text1_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
    'Added by Lydia 2021/09/24 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
    End If

   TxtValidate = True
End Function

'Added by Lydia 2022/03/15
Private Function GetMaxNo() As String
Dim strQ As String, intQ As Integer
Dim RsQ As New ADODB.Recordset

    strQ = "select nvl(max(tmn05),0)+1 as mno from trademarkmerchandisename "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        GetMaxNo = "" & RsQ.Fields("mno")
    End If
    Set RsQ = Nothing
End Function
