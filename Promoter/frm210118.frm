VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210118 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶案件整理表紀錄查詢 "
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7845
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Height          =   405
      Index           =   2
      Left            =   5730
      TabIndex        =   13
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txtDate2 
      Height          =   300
      Left            =   2340
      MaxLength       =   7
      TabIndex        =   4
      Top             =   700
      Width           =   915
   End
   Begin VB.TextBox txtDate1 
      Height          =   300
      Left            =   1125
      MaxLength       =   7
      TabIndex        =   3
      Top             =   700
      Width           =   915
   End
   Begin VB.TextBox txtCU1 
      Height          =   300
      Left            =   1125
      MaxLength       =   9
      TabIndex        =   5
      Top             =   1020
      Width           =   915
   End
   Begin VB.TextBox txtSales 
      Height          =   300
      Left            =   1125
      MaxLength       =   6
      TabIndex        =   2
      Top             =   380
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea 
      Height          =   300
      Left            =   1125
      MaxLength       =   3
      TabIndex        =   0
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea1 
      Height          =   300
      Left            =   2340
      MaxLength       =   3
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   3525
      Left            =   30
      TabIndex        =   7
      Top             =   1380
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   6218
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
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
      _Band(0).Cols   =   1
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   405
      Index           =   1
      Left            =   6690
      TabIndex        =   8
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   4800
      TabIndex        =   6
      Top             =   60
      Width           =   915
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Left            =   2100
      TabIndex        =   15
      Top             =   1050
      Width           =   5700
      VariousPropertyBits=   27
      Size            =   "10054;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSalesName 
      Height          =   255
      Left            =   2100
      TabIndex        =   14
      Top             =   420
      Width           =   1590
      VariousPropertyBits=   27
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "日期範圍："
      Height          =   180
      Left            =   90
      TabIndex        =   12
      Top             =   760
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2070
      X2              =   2340
      Y1              =   843
      Y2              =   843
   End
   Begin VB.Line Line2 
      X1              =   2070
      X2              =   2340
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   90
      TabIndex        =   11
      Top             =   440
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請人編號："
      Height          =   180
      Left            =   60
      TabIndex        =   9
      Top             =   1095
      Width           =   1080
   End
End
Attribute VB_Name = "frm210118"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/26 改成Form2.0 ; grd1改字型=新細明體-ExtB、lblSalesName、lbl1
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
        grd1.MousePointer = flexHourglass
        ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/23 清除查詢印表記錄檔欄位
        StrMenu
        grd1.MousePointer = flexDefault
        Screen.MousePointer = vbDefault
        Me.Enabled = True
Case 1
        Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
SetGrd
txtSales = strUserNum
txtSalesArea = PUB_GetStaffST15(strUserNum, 1)
txtSalesArea1 = txtSalesArea
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm210118 = Nothing
End Sub
Private Sub txtCU1_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCU1_Validate(Cancel As Boolean)
If Len(txtCU1) < 6 And Trim(txtCU1) <> "" Then MsgBox "申請人編號錯誤！", vbCritical, "發生錯誤！": Cancel = True: Exit Sub
lbl1.Caption = GetCustomerName(txtCU1, 0)
End Sub

Private Sub txtDate1_GotFocus()
   TextInverse txtDate1
   CloseIme
End Sub

Private Sub txtDate1_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 9 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtDate1_Validate(Cancel As Boolean)
   If txtDate1 <> "" Then
      If ChkDate(txtDate1) = False Then
         Cancel = True
         txtDate1.SetFocus
         txtDate1_GotFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub txtDate2_GotFocus()
   TextInverse txtDate2
   CloseIme
End Sub

Private Sub txtDate2_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 9 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtDate2_Validate(Cancel As Boolean)
   If txtDate2 <> "" Then
        If ChkDate(txtDate2) = False Then
           Cancel = True
           txtDate2.SetFocus
           txtDate2_GotFocus
           Exit Sub
        End If
        If RunNick2(txtDate1, txtDate2) = True Then
           txtDate1.SetFocus
           txtDate1_GotFocus
           Cancel = True
           Exit Sub
        End If
   End If
End Sub

Private Sub txtSales_Change()
   If Len(txtSales) > 4 Then
      lblSalesName = GetStaffName(txtSales)
   Else
      lblSalesName = ""
   End If
End Sub

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   CloseIme
End Sub

'Add By Sindy 2010/11/26
Private Sub txtSales_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea_GotFocus()
   TextInverse txtSalesArea
   CloseIme
End Sub

Private Sub txtSalesArea_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_GotFocus()
   TextInverse txtSalesArea1
   CloseIme
End Sub

Private Sub txtSalesArea1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_Validate(Cancel As Boolean)
If Trim(txtSalesArea1) <> "" Then
   If RunNick(txtSalesArea, txtSalesArea1) = True Then
      Cancel = True
      Exit Sub
   End If
End If
End Sub

Sub StrMenu()
Dim stCon As String
Dim Cancel As Boolean
    
    Cancel = False
    txtCU1_Validate Cancel
    If Cancel = True Then Exit Sub
    txtSalesArea1_Validate Cancel
    If Cancel = True Then Exit Sub
    txtDate1_Validate Cancel
    If Cancel = True Then Exit Sub
    txtDate2_Validate Cancel
    If Cancel = True Then Exit Sub
    
    If txtSalesArea <> "" Then
        stCon = stCon & " and st15>='" & txtSalesArea & "'"
    End If
    If txtSalesArea1 <> "" Then
        stCon = stCon & " and st15<='" & txtSalesArea1 & "'"
    End If
    If txtSalesArea <> "" Or txtSalesArea1 <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1 & txtSalesArea & "-" & txtSalesArea1 'Add By Sindy 2010/12/23
    End If
    If txtSales <> "" Then
        stCon = stCon & " and dl06='" & txtSales & "' "
        pub_QL05 = pub_QL05 & ";" & Label4 & txtSales & lblSalesName 'Add By Sindy 2010/12/23
    End If
    If txtCU1 <> "" Then
        stCon = stCon & " and dl09 like '客戶編號：" & txtCU1 & "%' "
        pub_QL05 = pub_QL05 & ";" & Label8 & txtCU1 & lbl1 'Add By Sindy 2010/12/23
    End If
    If txtDate1 <> "" Then
        stCon = stCon & " and dl07>=" & DBDATE(txtDate1) & " "
    End If
    If txtDate2 <> "" Then
        stCon = stCon & " and dl07<=" & DBDATE(txtDate2) & " "
    End If
    If txtDate1 <> "" Or txtDate2 <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label2 & txtDate1 & "-" & txtDate2 'Add By Sindy 2010/12/23
    End If
    
    strSql = "select st02,sqldatet(dl07),sqltime(dl08),dl09 from dml_log,staff where dl06=st01(+) and instr(dl12,'frm210115')<> 0 " & stCon & " order by dl06,dl07,dl08"
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 Then
        InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/23
        Set grd1.Recordset = adoRecordset
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/23
        ShowNoData
    End If
    CheckOC
    SetGrd
End Sub

Private Sub SetGrd()
grd1.Cols = 4
grd1.row = 0
grd1.col = 0: grd1.Text = "操作人員"
grd1.ColWidth(0) = 900
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 1: grd1.Text = "列印日期"
grd1.ColWidth(1) = 900
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 2: grd1.Text = "列印時間"
grd1.ColWidth(2) = 900
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 3: grd1.Text = "列印條件、備註"
grd1.ColWidth(3) = 6000
grd1.CellAlignment = flexAlignCenterCenter
End Sub
