VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21b0 
   AutoRedraw      =   -1  'True
   Caption         =   "調整付款明細"
   ClientHeight    =   5196
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5196
   ScaleWidth      =   8760
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Frmacc21b0.frx":0000
      Left            =   6600
      List            =   "Frmacc21b0.frx":0007
      TabIndex        =   21
      Top             =   1635
      Width           =   735
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6840
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1020
      Width           =   1572
   End
   Begin VB.CommandButton Command4 
      Caption         =   "轉至"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3720
      TabIndex        =   4
      Top             =   1065
      Width           =   1692
   End
   Begin VB.CommandButton Command3 
      Caption         =   "存檔"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7656
      TabIndex        =   7
      Top             =   1635
      Width           =   876
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Frmacc21b0.frx":000E
      Left            =   3936
      List            =   "Frmacc21b0.frx":0010
      TabIndex        =   5
      Top             =   1635
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "刪除結匯單據"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3720
      TabIndex        =   2
      Top             =   672
      Width           =   1692
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc21b0.frx":0012
      Height          =   2160
      Left            =   240
      TabIndex        =   8
      Top             =   2430
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   3810
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "a1902"
         Caption         =   "單據編號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "a1903"
         Caption         =   "幣別"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "a1904"
         Caption         =   "單據金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "a1907"
         Caption         =   "國內客戶名稱(收據抬頭)"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "a1916"
         Caption         =   "個人/公司"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1332.284
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   624.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1391.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4415.811
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1140.095
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6840
      TabIndex        =   17
      Top             =   240
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      Caption         =   "新增付款單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3720
      TabIndex        =   1
      Top             =   279
      Width           =   1692
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2496
      TabIndex        =   15
      Top             =   4635
      Width           =   1332
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   13
      Top             =   2010
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   11
      Top             =   1635
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   30
      Top             =   2370
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   572
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSForms.TextBox Text4 
      Height          =   330
      Left            =   2910
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2010
      Width           =   5535
      VariousPropertyBits=   671105051
      BackColor       =   16777215
      Size            =   "9763;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5850
      TabIndex        =   20
      Top             =   1673
      Width           =   795
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "其他付款單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5427
      TabIndex        =   19
      Top             =   1065
      Width           =   1425
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "匯款方式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2970
      TabIndex        =   18
      Top             =   1673
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "新付款單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5640
      TabIndex        =   16
      Top             =   279
      Width           =   1212
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   -24
      Top             =   4752
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1350
      Left            =   240
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   450
      TabIndex        =   14
      Top             =   4673
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "代理人 "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   2055
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "付款單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   10
      Top             =   1674
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "單據編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   9
      Top             =   279
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc21b0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/03 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text4
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
'Memo by Lydia 2020/08/31 9/1取消"1公司的國外結匯" (combo1拿掉)
Option Explicit
Public adoacc170 As New ADODB.Recordset
Public adoacc180 As New ADODB.Recordset
Public adoacc190 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adocaseprogress As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

Private Sub Command1_Click()
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   AdodcDelete
   AdodcRefresh
   SumShow
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Command2_Click()
   Acc190NewSave
   DataGrid1.Refresh
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Command3_Click()
   Screen.MousePointer = vbHourglass
   If Text2 = MsgText(601) Then
      Exit Sub
   End If
   '2012/10/12 ADD BY SONIA
   If Combo2 = "" Then
      Screen.MousePointer = vbDefault
      'Modified by Lydia 2017/09/06
      'MsgBox "匯票方式不可為空白...", , MsgText(5)
      MsgBox "匯款方式不可為空白...", , MsgText(5)
      Exit Sub
   End If
   If Combo1 = "" Then
      Screen.MousePointer = vbDefault
      MsgBox "公司別不可為空白...", , MsgText(5)
      Exit Sub
   End If
   'Added by Lydia 2017/09/22 檢查匯款方式
   If InStr("3,5", Left(Combo2.Text, 1)) > 0 And Left(Combo1.Text, 1) = "J" Then
      Screen.MousePointer = vbDefault
      MsgBox Mid(Trim(Combo2.Text), 4) & "的公司別不可為J...", , MsgText(5)
      Exit Sub
   End If
   'end 2017/09/22
   
   '2012/10/12 END
   
   'Added by Lydia 2021/12/03 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Screen.MousePointer = vbDefault
       Exit Sub
   End If
   'end 2021/12/03
   
   '2012/10/12 MODIFY BY SONIA 記錄修改人員,日期,時間
   'adoTaie.Execute "update acc180 set a1810 = " & IIf(Me.Text4.Text <> Me.Text4.Tag, CNULL(Text4), "a1810") & ", a1811 = " & IIf(Combo2 = "", "null", Mid(Combo2, 1, 1)) & " where a1801 = '" & Text2 & "'"
   adoTaie.Execute "update acc180 set a1810 = " & IIf(Me.Text4.Text <> Me.Text4.Tag, CNULL(Text4), "a1810") & ", a1811 = " & IIf(Combo2 = "", "null", Mid(Combo2, 1, 1)) & ",a1807=" & Val(strSrvDate(2)) & ",a1808=" & ServerTime & ",a1809='" & strUserNum & "' where a1801 = '" & Text2 & "'"
   '2012/10/12 END
   '更新公司別
   '2012/10/15 MODIFY BY SONIA 記錄修改人員,日期,時間
   'adoTaie.Execute "update acc190 set a1917 = " & CNULL(Trim(Me.Combo1.Text)) & " where a1901 = '" & Text2 & "'"
   adoTaie.Execute "update acc190 set a1917 = " & CNULL(Trim(Me.Combo1.Text)) & ",a1912=" & Val(strSrvDate(2)) & ",a1913=" & ServerTime & ",a1914='" & strUserNum & "' where a1901 = '" & Text2 & "'"
   '2012/10/15 end
   Screen.MousePointer = vbDefault
   MsgBox MsgText(17), , MsgText(21)
End Sub

Private Sub Command4_Click()
Dim StrSQLa As String
Dim rsA  As New ADODB.Recordset
Dim strCompany As String '公司別
   
   If Text7 = MsgText(601) Or Text1 = MsgText(601) Then
      Exit Sub
   End If
    '檢查公司別, 若不同不可做轉號動作
    StrSQLa = "Select * From ACC190 Where a1902='" & Me.Text1.Text & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        strCompany = "" & rsA("a1917").Value
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        StrSQLa = "Select * From ACC190 Where a1901='" & Me.Text7.Text & "' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            '若公司別不同
            If strCompany <> "" & rsA("a1917").Value Then
                MsgBox "公司別不同，不可做轉號動作!!!", vbExclamation + vbOKOnly
                Me.Text1.SetFocus
                TextInverse Me.Text1
                If rsA.State <> adStateClosed Then rsA.Close
                Set rsA = Nothing
                Exit Sub
            End If
        Else
            MsgBox "查無資料!!!", vbExclamation + vbOKOnly
            Me.Text7.SetFocus
            TextInverse Me.Text7
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Exit Sub
        End If
    Else
        MsgBox "查無資料!!!", vbExclamation + vbOKOnly
        Me.Text1.SetFocus
        TextInverse Me.Text1
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        Exit Sub
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    'ENd
   If adoaccsum.State = adStateOpen Then
      adoaccsum.Close
   End If
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select * from acc180, acc190 where a1801 = a1901 and a1801 = '" & Text7 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount = 0 Then
      MsgBox MsgText(28), , MsgText(5)
      adoaccsum.Close
      Exit Sub
   Else
      If adoquery.State = adStateOpen Then
         adoquery.Close
      End If
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select * from acc170 where a1702 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If adoquery.Fields("a1703").Value <> adoaccsum.Fields("a1903").Value Or adoquery.Fields("a1705").Value <> adoaccsum.Fields("a1803").Value Then
            MsgBox MsgText(201), , MsgText(5)
            adoquery.Close
            adoaccsum.Close
            Exit Sub
         End If
      Else
         MsgBox MsgText(28), , MsgText(5)
         adoquery.Close
         adoaccsum.Close
         Exit Sub
      End If
      adoquery.Close
   End If
   adoaccsum.Close
   Screen.MousePointer = vbHourglass
   '2012/10/15 modify by sonia 記錄修改人員,日期,時間
   'adoTaie.Execute "update acc190 set a1901 = '" & Text7 & "' where a1902 = '" & Text1 & "'"
   'adoTaie.Execute "update acc170 set a1709 = '" & Text7 & "' where a1702 = '" & Text1 & "'"
   adoTaie.Execute "update acc190 set a1901 = '" & Text7 & "',a1912=" & Val(strSrvDate(2)) & ",a1913=" & ServerTime & ",a1914='" & strUserNum & "' where a1902 = '" & Text1 & "'"
   adoTaie.Execute "update acc170 set a1709 = '" & Text7 & "',a1713=" & Val(strSrvDate(2)) & ",a1714=" & ServerTime & ",a1715='" & strUserNum & "' where a1702 = '" & Text1 & "'"
   '2012/10/15 end
   Acc190Save
   Screen.MousePointer = vbDefault
   MsgBox MsgText(17), , MsgText(21)
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Text1 = Adodc1.Recordset.Fields("a1902").Value
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy W8850 H5500
   Me.Width = 8860
   Me.Height = 5640
   'end 2023/08/18
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Combo2.AddItem ComboItem(71)
   Combo2.AddItem ComboItem(72)
   'Added by Lydia 2015/04/17 增加匯款方式3
   Combo2.AddItem ComboItem(77)
   'Added by Lydia 2017/09/06 增加匯款方式4
   Combo2.AddItem ComboItem(78)
   'Added by Lydia 2017/09/22 增加匯款方式5
   Combo2.AddItem ComboItem(79)
   'Added by Lydia 2024/09/03 增加匯款方式6
   Combo2.AddItem ComboItem(80)
   
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   '2012/10/12 add by sonia 刪除沒有ACC190的ACC180(ACC190全部轉出)
   adoTaie.Execute "DELETE ACC180 WHERE A1801 IN (SELECT DISTINCT A1801 FROM ACC180,ACC190 WHERE A1801=A1901(+) AND A1902 IS NULL)"
   '2012/10/12 end
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc21b0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc190 where a1901 = '" & Text2 & "' order by a1902 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc190 where a1901 = '" & Text2 & "' order by a1902 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  儲存資料表(國外付款資料(主檔)及(交易檔))
'
'*************************************************
Private Sub Acc190Save()
Dim strName As String

On Error GoTo Checking
   Text6 = ""
   Text7 = ""
   adoacc170.CursorLocation = adUseClient
   adoacc170.Open "select * from acc170 where a1702 = '" & Text1 & "' and a1709 is null", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc170.RecordCount <> 0 Then
      'Modified by Lydia 2020/01/02 改成彈提醒;
                            '發生原因:U10812083=> 只開啟"結匯資料輸入frmacc2170"並沒有按下產生付款單,接著直接到"調整付款明細frmacc21b0"按下Insser產生付款單,
                            '但是"調整付款明細frmacc21b0"不同於frmacc2170,並無預設"匯款方式",並且設定代理人名稱=代理人編號。
'      Text2 = AutoNo(MsgText(814), 5)
'      If IsNull(adoacc170.Fields("a1705").Value) Then
'         Text3 = MsgText(601)
'      Else
'         Text3 = adoacc170.Fields("a1705").Value
'         '2005/9/8 ADD BY SONIA
'         If IsNull(adoacc170.Fields("a1705").Value) Then
'            Text4 = FagentQuery_1(Text3, 2)
'            '2006/3/10 ADD BY SONIA
'            If Text4 = "" Then
'               Text4 = CustomerQuery(Text3, 2)
'            End If
'            '2006/3/10 END
'         Else
'            Text4 = adoacc170.Fields("a1705").Value
'         End If
'         '2005/9/8 END
'      End If
      MsgBox "尚未產生付款單,請到結匯資料輸入產生付款單！", , MsgText(5)
      adoacc170.Close
      Exit Sub
      'end 2020/01/02
   Else
      Text2 = MsgText(601)
      Text3 = MsgText(601)
      adoacc190.CursorLocation = adUseClient
      adoacc190.Open "select DISTINCT a1901 from acc190 where a1902 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc190.RecordCount <> 0 Then
         adoacc180.CursorLocation = adUseClient
         adoacc180.Open "select * from acc180 where a1801 = '" & adoacc190.Fields("a1901").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoacc180.RecordCount <> 0 Then
            Text2 = adoacc180.Fields("a1801").Value
            If IsNull(adoacc180.Fields("a1803").Value) Then
               Text3 = MsgText(601)
            Else
               Text3 = adoacc180.Fields("a1803").Value
            End If
            If IsNull(adoacc180.Fields("a1810").Value) Then
                'Modify By Cheng 2003/07/01
'               Text4 = FagentQuery(Text3, 2)
               Text4 = FagentQuery_1(Text3, 2)
               '2006/3/10 ADD BY SONIA
               If Text4 = "" Then
                  Text4 = CustomerQuery(Text3, 2)
               End If
               '2006/3/10 END
            Else
               Text4 = adoacc180.Fields("a1810").Value
            End If
            'Add By Cheng 2003/07/04
            '記錄原代理人名稱
            Me.Text4.Tag = Me.Text4.Text
            If IsNull(adoacc180.Fields("a1811").Value) Then
               Combo2 = MsgText(601)
            Else
               Combo2 = Combo2.List(Val(adoacc180.Fields("a1811").Value) - 1)
            End If
            AdodcRefresh
            'Add By Cheng 2003/06/12
            '取得公司別
            Me.Combo1.Text = Geta1917(Me.Text2.Text)
         End If
         adoacc180.Close
         adoacc170.Close
         adoacc190.Close
         Exit Sub
      Else
         MsgBox MsgText(28), , MsgText(5)
         adoacc170.Close
         adoacc190.Close
         Exit Sub
      End If
      AdodcRefresh
      adoacc190.Close
      adoacc170.Close
      Exit Sub
   End If
   adoacc170.Close
   adoacc180.CursorLocation = adUseClient
   adoacc180.Open "select * from acc180 where a1801 = '" & Text2 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc180.RecordCount = 0 Then
      adoacc180.AddNew
      adoacc180.Fields("a1801").Value = Text2
      adoacc180.Fields("a1802").Value = Val(strSrvDate(2))
      If Text3 <> MsgText(601) Then
         adoacc180.Fields("a1803").Value = Text3
      Else
         adoacc180.Fields("a1803").Value = Null
      End If
      adoacc180.Fields("a1804").Value = Val(strSrvDate(2))
      adoacc180.Fields("a1805").Value = ServerTime
      adoacc180.Fields("a1806").Value = strUserNum
      adoacc180.UpdateBatch
   Else
      MsgBox MsgText(9), , MsgText(5)
      adoacc180.Close
      Exit Sub
   End If
   adoacc180.Close
   adoacc170.CursorLocation = adUseClient
   adoacc170.Open "select * from acc170 where a1705 = '" & Text3 & "' and a1709 is null order by a1702 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Do While adoacc170.EOF = False
      adocaseprogress.CursorLocation = adUseClient
      adocaseprogress.Open "select a0k04 from caseprogress, acc0k0 where cp60 = a0k01 and cp61 = '" & adoacc170.Fields("a1702").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocaseprogress.RecordCount <> 0 Then
         If IsNull(adocaseprogress.Fields(0).Value) Then
            strName = ""
         Else
            strName = adocaseprogress.Fields(0).Value
         End If
      Else
         strName = ""
      End If
      adocaseprogress.Close
      '2005/9/8 ADD BY SONIA
      If strName = "" And Text4 <> "" Then strName = Text4
      '2005/9/8 END
        'Modify By Cheng 2003/06/10
        '寫入公司別欄位
'      adoTaie.Execute "insert into acc190 (a1901, a1902, a1903, a1904, a1909, a1910, a1911, a1915, a1907) values ('" & Text2 & "', '" & adoacc170.Fields("a1702").Value & "', '" & adoacc170.Fields("a1703").Value & "', " & _
'                      "" & Val(adoacc170.Fields("a1704").Value) & ", " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "', " & Val(adoacc170.Fields("a1708").Value) & ", '" & strName & "')"
      'Modified by Morgan 2012/1/13 收據抬頭會有單引號
      adoTaie.Execute "insert into acc190 (a1901, a1902, a1903, a1904, a1909, a1910, a1911, a1915, a1907, a1917) values ('" & Text2 & "', '" & adoacc170.Fields("a1702").Value & "', '" & adoacc170.Fields("a1703").Value & "', " & _
                      "" & Val(adoacc170.Fields("a1704").Value) & ", " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', " & Val(adoacc170.Fields("a1708").Value) & ", '" & ChgSQL(strName) & "','" & ChgSQL(GetCompany("" & adoacc170.Fields("a1702").Value)) & "')"
      If Text2 <> MsgText(601) Then
         adoacc170.Fields("a1709").Value = Text2
      Else
         adoacc170.Fields("a1709").Value = Null
      End If
      adoacc170.UpdateBatch
      adoacc170.MoveNext
   Loop
   AdodcRefresh
   adoacc170.Close
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyInsert
         'Add by Morgan 2009/5/21
         If CheckValidate Then
            Acc190Save
            SumShow
            Text1 = ""
            Text6 = ""
            Text1.SetFocus
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
End Sub
'Add by Morgan 2009/5/21
'檢查匯票是否已輸入
Private Function CheckValidate() As Boolean
   'Added by Lydia 2017/09/06 檢查匯款方式
   If Mid(Combo2, 1, 1) = "3" And UCase(Mid(Combo1, 1, 1)) = "J" Then
      MsgBox Combo2.Text & "的公司別不可為J公司 !!"
      CheckValidate = False
      Exit Function
   End If
   'end 2017/09/06
   strExc(0) = "select * from acc170,acc190 where a1702='" & Text1 & "' and a1901(+)=a1709 and a1902(+)=a1702 and a1908 is not null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      CheckValidate = False
      MsgBox "該單據已有匯票號碼，不可再作業！"
   Else
      CheckValidate = True
   End If
   
End Function
'*************************************************
'  計算並顯示合計
'
'*************************************************
Private Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a1904) from acc190 where a1901 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text5 = MsgText(601)
      Else
         Text5 = Format(adoaccsum.Fields(0).Value, FDollar)
      End If
   Else
      Text5 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  刪除資料表(國外付款資料(交易檔))
'
'*************************************************
Private Sub AdodcDelete()
Dim strType As String

On Error GoTo Checking
   If Adodc1.Recordset.RecordCount <> 0 Then
      Select Case Mid(Adodc1.Recordset.Fields("a1902").Value, 1, 1)
         Case MsgText(812)
            strType = "1"
         Case MsgText(813)
            strType = "2"
         Case MsgText(809)
            strType = "3"
      End Select
      'adoTaie.Execute "update acc170 set a1709 = '' where a1702 = '" & Adodc1.Recordset.Fields("a1902").Value & "'"
      adoTaie.Execute "delete from acc170 where a1702 = '" & Adodc1.Recordset.Fields("a1902").Value & "'"
      Adodc1.Recordset.Delete
      Adodc1.Recordset.UpdateBatch
      If Adodc1.Recordset.RecordCount = 0 Then
         adoTaie.Execute "delete from acc180 where a1801 = '" & Text2 & "'"
      End If
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  新增國外付款單資料
'
'*************************************************
Private Sub Acc190NewSave()
Dim strPayNo As String

On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   strPayNo = AutoNo(MsgText(814), 5)
   Text6 = strPayNo
    'Modify By Cheng 2003/08/29
'   adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1804, a1805, a1806) values ('" & Text6 & "', " & Val(ACDate(ServerDate)) & ", " & _
'                   "'" & Text3 & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "')"
   'Modified by Lydia 2015/06/12 +台銀電匯匯紙本判斷
   'adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1804, a1805, a1806, a1811) values ('" & Text6 & "', " & strSrvDate(2) & ", " & _
                   "'" & Text3 & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "','" & GetTermOfPayment("" & Adodc1.Recordset.Fields("a1902").Value, "" & Adodc1.Recordset.Fields("a1903").Value) & "')"
   adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1804, a1805, a1806, a1811) values ('" & Text6 & "', " & strSrvDate(2) & ", " & _
                   "'" & Text3 & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "','" & GetTermOfPayment("" & Adodc1.Recordset.Fields("a1902").Value, "" & Adodc1.Recordset.Fields("a1903").Value, Combo1.Text) & "')"
                   
    'Modify By Cheng 2003/06/10
    '寫入公司別欄位
'   adoTaie.Execute "insert into acc190 (a1901, a1902, a1903, a1904, a1907, a1909, a1910, a1911, a1915) values ('" & Text6 & "', '" & Adodc1.Recordset.Fields("a1902").Value & "', " & _
'                   "'" & Adodc1.Recordset.Fields("a1903").Value & "', " & Val(Adodc1.Recordset.Fields("a1904").Value) & ", '" & Adodc1.Recordset.Fields("a1907").Value & "', " & _
'                   "" & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "', " & Val(Adodc1.Recordset.Fields("a1915").Value) & ")"
   adoTaie.Execute "insert into acc190 (a1901, a1902, a1903, a1904, a1907, a1909, a1910, a1911, a1915, a1917) values ('" & Text6 & "', '" & Adodc1.Recordset.Fields("a1902").Value & "', " & _
                   "'" & Adodc1.Recordset.Fields("a1903").Value & "', " & Val(Adodc1.Recordset.Fields("a1904").Value) & ", '" & Adodc1.Recordset.Fields("a1907").Value & "', " & _
                   "" & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', " & Val(Adodc1.Recordset.Fields("a1915").Value) & ",'" & GetCompany("" & Adodc1.Recordset.Fields("a1902").Value) & "')"
   adoTaie.Execute "update acc170 set a1709 = '" & Text6 & "' where a1702 = '" & Adodc1.Recordset.Fields("a1902").Value & "'"
   Adodc1.Recordset.Delete
   Adodc1.Recordset.UpdateBatch
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

'Modified by Lydia 2021/12/03 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub Text4_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Cheng 2003/06/10
'取得公司別
Private Function GetCompany(strA1902 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   GetCompany = ""
   '2006/11/1 MODIFY BY SONIA 同時抓收據號碼及本所案號, 並加入CP87,CP88,並修改抵帳單抓收據方式, 帳單不能改同抵帳單因為舊帳單無收文號之故
   StrSQLa = "select a0k01,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11 from acc170, caseprogress, acc0k0, systemkind where a1702 = cp61 and cp60 = a0k01 (+) and a1701 = '1' and a1702='" & strA1902 & "' and cp01=sk01 union " & _
             "select a0k01,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11 from acc170, caseprogress, acc0k0, systemkind where a1702 = cp62 and cp60 = a0k01 (+) and a1701 = '1' and a1702='" & strA1902 & "' and cp01=sk01 union " & _
             "select a0k01,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11 from acc170, caseprogress, acc0k0, systemkind where a1702 = cp63 and cp60 = a0k01 (+) and a1701 = '1' and a1702='" & strA1902 & "' and cp01=sk01 union " & _
             "select a0k01,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11 from acc170, caseprogress, acc0k0, systemkind where a1702 = cp87 and cp60 = a0k01 (+) and a1701 = '1' and a1702='" & strA1902 & "' and cp01=sk01 union " & _
             "select a0k01,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11 from acc170, caseprogress, acc0k0, systemkind where a1702 = cp88 and cp60 = a0k01 (+) and a1701 = '1' and a1702='" & strA1902 & "' and cp01=sk01 union " & _
             "select a0k01,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11 from acc170, caseprogress, acc0k0, systemkind where a1702 = cp61 (+) and cp60 = a0k01 (+) and a1701 = '1' and length(a1702) = 10 and (cp61 is null and cp62 is null and cp63 is null and cp87 is null and cp88 is null) and (a1709 is null or a1709 = '') and cp01=sk01 union " & _
             "select a0k01,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11 from acc170, caseprogress, acc0k0, systemkind, acc161, acc160 where a1702 = axg01 and a1702 = a1601 and AXG02=CP09 AND CP60 = a0k01 (+) and a1701 = '2' and a1702='" & strA1902 & "' and cp01=sk01 union " & _
             "select NULL a0k01,NULL caseNo,'2' as a0k11 from acc170 where a1701 = '3' and a1702='" & strA1902 & "' union " & _
             "select NULL a0k01,NULL caseNo,'2' as a0k11 from acc170 where a1701 = '4' and a1702='" & strA1902 & "' "
   '2006/11/1 END
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       '2006/11/1 MODIFY BY SONIA
       'GetCompany = "" & rsA.Fields(0).Value
       GetCompany = GetComp("" & rsA.Fields(0).Value, "" & rsA.Fields(1).Value, "" & rsA.Fields(2).Value)
       '2006/11/1 END
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

'取得公司別
Private Function Geta1917(stra1901) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   StrSQLa = "Select a1917 From acc190 Where a1901='" & stra1901 & "' "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       Geta1917 = "" & rsA.Fields(0).Value
   Else
       Geta1917 = ""
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
End Function
