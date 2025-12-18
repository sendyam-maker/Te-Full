VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm04010505_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "證書號數輸入"
   ClientHeight    =   5745
   ClientLeft      =   210
   ClientTop       =   690
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9300
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   270
      Left            =   4452
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "P"
      Top             =   672
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   4932
      MaxLength       =   6
      TabIndex        =   3
      Top             =   672
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   5772
      MaxLength       =   1
      TabIndex        =   4
      Top             =   672
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   6024
      MaxLength       =   2
      TabIndex        =   5
      Top             =   672
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號:"
      Height          =   228
      Index           =   1
      Left            =   3264
      TabIndex        =   1
      Top             =   720
      Width           =   1140
   End
   Begin VB.OptionButton Option1 
      Caption         =   "申請案號:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   11
      Top             =   672
      Value           =   -1  'True
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8328
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7536
      TabIndex        =   9
      Top             =   72
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1344
      MaxLength       =   20
      TabIndex        =   0
      Top             =   672
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1344
      MaxLength       =   8
      TabIndex        =   7
      Top             =   1260
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm04010505_1.frx":0000
      Height          =   3984
      Left            =   120
      TabIndex        =   8
      Top             =   1644
      Width           =   9012
      _ExtentX        =   15901
      _ExtentY        =   7038
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "soul"
         Caption         =   "本所案號"
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
         DataField       =   "NAME1"
         Caption         =   "專利名稱"
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
         ScrollBars      =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4380.095
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   336
      Left            =   192
      Top             =   2208
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin VB.Frame Frame1 
      Height          =   636
      Left            =   72
      TabIndex        =   13
      Top             =   480
      Width           =   7740
      Begin VB.CommandButton Command1 
         Caption         =   "尋找(&F)"
         Default         =   -1  'True
         Height          =   300
         Left            =   6504
         TabIndex        =   6
         Top             =   192
         Width           =   945
      End
   End
   Begin VB.Label lblBillNo 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   4350
      TabIndex        =   15
      Top             =   1290
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "帳單編號:"
      Height          =   180
      Left            =   3540
      TabIndex        =   14
      Top             =   1290
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   336
      TabIndex        =   12
      Top             =   1296
      Width           =   948
   End
End
Attribute VB_Name = "frm04010505_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (DataGrid1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim pmain As New ADODB.Recordset
Dim pmain1 As New ADODB.Recordset
Dim pSelect As New ADODB.Recordset
Public NUMBER1 As String
Public NUMBER2 As String
Public NUMBER3 As String
Public NUMBER4 As String
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_AppNo As String
Public m_RDate As String
Dim m_Done As Boolean
'2016/10/5 END

'Added by Morgan 2022/12/19
Public m_DocWord As String
Public m_DocNo As String
'end 2022/12/19


Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   
   If Text2 = "" Then
      MsgBox "來函收文日不可空白 !", vbCritical
      Text2.SetFocus
      Exit Function
      
   'Add by Morgan 2009/7/31
   Else
      Text2_Validate Cancel
      If Cancel = True Then
         Text2.SetFocus
         Text2_GotFocus
         Exit Function
      End If
      
   End If
   TxtValidate = True
   
End Function

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
   Case 0
   
   If TxtValidate = False Then Exit Sub

   If pmain.State = adStateClosed Then
      MsgBox "請先搜尋資料後再執行此動作 !", vbInformation
      If Option1(0).Value = True Then
         Text1.SetFocus
      Else
         Text4.SetFocus
      End If
      Exit Sub
   Else
      If pmain.RecordCount < 1 Then
         MsgBox "請選擇資料 !", vbInformation
         Exit Sub
      End If
   End If
   
   If pmain.Fields(6).Value = "000" Then
      If pmain1.State = adStateOpen Then pmain1.Close
      strExc(1) = "SELECT MR12,MR13,MR14,MR15 FROM PATENT,MailRec WHERE MR12='" & pmain.Fields(2).Value & "' AND MR13='" & pmain.Fields(3).Value & "' AND MR14='" & pmain.Fields(4).Value & "' AND MR15='" & pmain.Fields(5).Value & "' AND MR02='" & ChangeTStringToWString(Text2.Text) & "' AND (MR16 IS NULL OR MR16=0) AND PA01=MR12 AND PA02=MR13 AND PA03=MR14 AND PA04= MR15 AND PA09='000' "
      pmain1.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
      If Not pmain1.BOF Then pmain1.MoveFirst
      'Modified by Morgan 2023/1/12 排除電子公文
      If pmain1.EOF And pmain1.BOF Then
      
         If m_DocNo = "" Then 'Added by Morgan 2023/1/12
            If MsgBox("與櫃台之來函收文記錄不符,請確認", vbOKCancel) = vbCancel Then
               'Unload Me
               Exit Sub
            End If
         End If
         
           NUMBER1 = pmain.Fields(2).Value
           NUMBER2 = pmain.Fields(3).Value
           NUMBER3 = pmain.Fields(4).Value
           NUMBER4 = pmain.Fields(5).Value
           'Add By Sindy 2017/12/27
           If m_strIR01 <> "" Then
               If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> NUMBER1 & NUMBER2 & NUMBER3 & NUMBER4 Then
                  MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
                  Exit Sub
               End If
           End If
           '2017/12/27 END
   
            'Added by Morgan 2022/12/19
            frm04010505_2.m_DocNo = m_DocNo
            frm04010505_2.m_DocWord = m_DocWord
            'end 2022/12/19
         
           'Add By Sindy 2016/10/5
           frm04010505_2.m_strIR01 = m_strIR01
           frm04010505_2.m_strIR02 = m_strIR02
           frm04010505_2.m_strIR03 = m_strIR03
           frm04010505_2.m_strIR04 = m_strIR04
           '2016/10/5 END
           frm04010505_2.Show
           frm04010505_1.Hide
         
      Else
      
         NUMBER1 = pmain1.Fields(0).Value
         NUMBER2 = pmain1.Fields(1).Value
         NUMBER3 = pmain1.Fields(2).Value
         NUMBER4 = pmain1.Fields(3).Value
            
         'Add By Sindy 2017/12/27
         If m_strIR01 <> "" Then
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> NUMBER1 & NUMBER2 & NUMBER3 & NUMBER4 Then
               MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
               Exit Sub
            End If
         End If
         '2017/12/27 END
         'Add By Sindy 2016/10/5
         frm04010505_2.m_strIR01 = m_strIR01
         frm04010505_2.m_strIR02 = m_strIR02
         frm04010505_2.m_strIR03 = m_strIR03
         frm04010505_2.m_strIR04 = m_strIR04
         '2016/10/5 END
         frm04010505_2.Show
         frm04010505_1.Hide
      End If
   Else
   
      NUMBER1 = pmain.Fields(2).Value
      NUMBER2 = pmain.Fields(3).Value
      NUMBER3 = pmain.Fields(4).Value
      NUMBER4 = pmain.Fields(5).Value
      
      'Add By Sindy 2017/12/27
      If m_strIR01 <> "" Then
         If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> NUMBER1 & NUMBER2 & NUMBER3 & NUMBER4 Then
            MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
            Exit Sub
         End If
      End If
      '2017/12/27 END
      'Add By Sindy 2016/10/5
      frm04010505_2.m_strIR01 = m_strIR01
      frm04010505_2.m_strIR02 = m_strIR02
      frm04010505_2.m_strIR03 = m_strIR03
      frm04010505_2.m_strIR04 = m_strIR04
      '2016/10/5 END
      frm04010505_2.Show
      frm04010505_1.Hide
   End If
       
       Case 1
          Unload Me
End Select
End Sub

'Private Sub Form_Activate()
'If strKey1 = "1" Then
'Text1.Text = ""
' If pmain.State = adStateOpen Then pmain.Close
'   strExc(0) = "SELECT (PA01||'-'||PA02||'-'||PA03||'-'||PA04) as soul,nvl(PA05,nvl(pa06,pa07)) AS NAME1,PA01,PA02,PA03,PA04,PA09 FROM Patent WHERE PA01='P' AND PA16='1' AND PA09='000' AND PA11='' ORDER BY PA01,PA02,PA03,PA04"
'   pmain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
'   Set Adodc1.Recordset = pmain
'   Adodc1.Recordset.ReQuery
'End If
'strKey1 = ""
'End Sub

Public Sub Clear()
   'Modify By Cheng 2002/06/19
   '選擇申請案號
   If Me.Option1(0).Value Then
        'Modify By Cheng 2002/12/18
        '保留原輸入條件
'      Text1.Text = ""
      If pmain.State = adStateOpen Then
         Text1.SetFocus
         pmain.Close
      End If
      strExc(0) = "SELECT (PA01||'-'||PA02||'-'||PA03||'-'||PA04) as soul,nvl(PA05,nvl(pa06,pa07)) AS NAME1,PA01,PA02,PA03,PA04,PA09 FROM Patent WHERE PA01='P' AND PA16='1' AND PA09='000' AND PA11='' ORDER BY PA01,PA02,PA03,PA04"
      pmain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      Set Adodc1.Recordset = pmain
      Adodc1.Recordset.ReQuery
      strKey1 = ""
      'Add By Cheng 2002/12/18
      If Me.Text1.Enabled Then TextInverse Me.Text1
   '選擇本所案號
   Else
        'Modify By Cheng 2002/12/18
        '保留原輸入條件
'      Me.Text3.Text = "P"
'      Me.Text4.Text = ""
'      Me.Text5.Text = ""
'      Me.Text6.Text = ""
      If pmain.State = adStateOpen Then
         Text4.SetFocus
         pmain.Close
      End If
      strExc(0) = "SELECT (PA01||'-'||PA02||'-'||PA03||'-'||PA04) as soul,nvl(PA05,nvl(pa06,pa07)) AS NAME1,PA01,PA02,PA03,PA04,PA09 FROM Patent WHERE PA01='' AND PA02='' AND PA03='' AND PA04='' AND PA16='1' AND PA09='000' ORDER BY PA01,PA02,PA03,PA04"
      pmain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      Set Adodc1.Recordset = pmain
      Adodc1.Recordset.ReQuery
      strKey1 = ""
      'Add By Cheng 2002/12/18
      If Me.Text4.Enabled Then TextInverse Me.Text4
   End If
End Sub

Private Sub Command1_Click()
 Dim strTmp As String
   If Option1(0).Value = True Then
      If Text1 = "" Then
         MsgBox "請輸入申請案號 !", vbCritical
         Text1_GotFocus
         Exit Sub
      End If
      If pmain.State = adStateOpen Then pmain.Close
      '92.12.31 MODIFY BY SONIA
      '2008/5/12 modify by sonia,非香港澳門需檢查已核准pa16='1',香港澳門不用(自動發證)
      'strExc(0) = "SELECT (PA01||'-'||PA02||'-'||PA03||'-'||PA04) as soul,nvl(PA05,nvl(pa06,pa07)) AS NAME1,PA01,PA02,PA03,PA04,PA09 FROM Patent WHERE PA01='P' AND PA16='1' AND PA11='" & Text1.Text & "' ORDER BY PA01,PA02,PA03,PA04"
'      strExc(0) = "SELECT (PA01||'-'||PA02||'-'||PA03||'-'||PA04) as soul," & _
         "nvl(PA05,nvl(pa06,pa07)) AS NAME1,PA01,PA02,PA03,PA04,PA09 FROM Patent " & _
         "WHERE PA01='P' AND PA16='1' AND PA11='" & Text1.Text & "' AND PA09<>'013' AND PA09<>'044' " & _
         "union " & _
         "SELECT (PA01||'-'||PA02||'-'||PA03||'-'||PA04) as soul," & _
         "nvl(PA05,nvl(pa06,pa07)) AS NAME1,PA01,PA02,PA03,PA04,PA09 FROM Patent " & _
         "WHERE PA11='" & Text1.Text & "' AND (PA09='013' or PA09='044') " & _
         "ORDER BY PA01,PA02,PA03,PA04"
      '92.12.31 END
'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
'設別名f0,+FMP2openSQL
      strExc(0) = "SELECT (PA01||'-'||PA02||'-'||PA03||'-'||PA04) as soul," & _
         "nvl(PA05,nvl(pa06,pa07)) AS NAME1,PA01,PA02,PA03,PA04,PA09 FROM Patent f0 " & _
         "WHERE PA01='P' AND PA16='1' AND PA11='" & Text1.Text & "' AND PA09<>'013' AND PA09<>'044' " & FMP2openSQL
      strExc(0) = Replace(strExc(0), "f0.CP", "f0.PA")
      
      If FMP2open = True Then '先以申請案號判斷使用權限
            strExc(1) = Replace(strExc(0), "PA01='P' AND PA16='1' AND", "")
            If PUB_FMPtoCheck(0, 1, Pub_strUserST05, "CHANGE_SQL", strExc(1)) = False Then
               Text1_GotFocus
               Exit Sub
            End If
      End If
      strExc(0) = strExc(0) & " union " & _
         "SELECT (PA01||'-'||PA02||'-'||PA03||'-'||PA04) as soul," & _
         "nvl(PA05,nvl(pa06,pa07)) AS NAME1,PA01,PA02,PA03,PA04,PA09 FROM Patent f2 " & _
         "WHERE PA11='" & Text1.Text & "' AND (PA09='013' or PA09='044') " & FMP2openSQL & _
         "ORDER BY PA01,PA02,PA03,PA04"
      strExc(0) = Replace(strExc(0), "f0.CP", "f2.PA")
      pmain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      If pmain.BOF And pmain.EOF Then
'         If FMP2open = True Then
'           MsgBox "權限不足 !", vbInformation
'         Else
           MsgBox "資料庫內無資料 !", vbInformation
'         End If
           Text1_GotFocus
      Else
        'Added by Lydia 2019/09/10 寰華案控制輸入帳單、已提申、發證書輸入，P的程序不能輸入並出現訊息告知USER「此案為FCP自行連繫，請交FCP程序處理」。
        If FMP2open = False And Pub_StrUserSt03 <> "M51" Then
             strTmp = "" & pmain.Fields("soul")
             If strTmp <> "" Then
                  Call ChgCaseNo(Replace(strTmp, "-", ""), strExc)
                  If PUB_FMPtoCheck(1, 2, Pub_strUserST05, strExc(1), strExc(2), strExc(3), strExc(4)) = True Then
                       MsgBox "此案為FCP自行連繫，請交FCP程序處理！", vbCritical, "寰華案控制輸入"
                       Set pmain = Nothing
                       Exit Sub
                  End If
             End If
        End If
        'end 2019/09/10
      
         Set Adodc1.Recordset = pmain
         Adodc1.Recordset.ReQuery
         If pmain.RecordCount = 1 Then
            cmdOK_Click 0
         ElseIf pmain.RecordCount > 1 Then
            cmdOK(0).Default = True
         End If
      End If
      
   Else
      If Text4 = "" Then
         MsgBox "請輸入本所案號 !", vbCritical
         Text4_GotFocus
         Exit Sub
      End If
      strTmp = Text3.Text & Text4.Text & Text5.Text & Text6.Text
      If FMP2open = True Then '先以申請案號判斷使用權限
           If PUB_FMPtoCheck(0, 1, Pub_strUserST05, Text3.Text, Text4.Text, IIf(Text5.Text = "", "0", Text5.Text), IIf(Text6.Text = "", "00", Text6.Text)) = False Then
              Text4_GotFocus
              Exit Sub
           End If
      'Added by Lydia 2019/09/10 寰華案控制輸入帳單、已提申、發證書輸入，P的程序不能輸入並出現訊息告知USER「此案為FCP自行連繫，請交FCP程序處理」。
      ElseIf Pub_StrUserSt03 <> "M51" Then
           If PUB_FMPtoCheck(1, 2, Pub_strUserST05, Text3.Text, Text4.Text, IIf(Text5.Text = "", "0", Text5.Text), IIf(Text6.Text = "", "00", Text6.Text)) = True Then
                MsgBox "此案為FCP自行連繫，請交FCP程序處理！", vbCritical, "寰華案控制輸入"
                Exit Sub
           End If
      'end 2019/09/10
      End If
      
      If pmain.State = adStateOpen Then pmain.Close
      If Me.Text3.Text = "P" Then
         strExc(0) = "SELECT PA09 FROM PATENT WHERE " & ChgPatent(strTmp)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If RsTemp.Fields(0) = "000" Then
            MsgBox "本案申請國家為台灣, 請改以申請案號查詢 !", vbExclamation
            Option1(0).Value = True
            
            Exit Sub
         End If
      
         'Modify by Morgan 2007/3/26 澳門和香港一樣
        'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
        '設別名f0,+FMP2openSQL
         strExc(0) = "SELECT (PA01||'-'||PA02||'-'||PA03||'-'||PA04) as soul," & _
            "nvl(PA05,nvl(pa06,pa07)) AS NAME1,PA01,PA02,PA03,PA04,PA09 FROM Patent f0 " & _
            "WHERE " & ChgPatent(strTmp) & " AND PA16='1' AND PA09<>'000' AND PA09<>'013' " & FMP2openSQL
         strExc(0) = Replace(strExc(0), "f0.CP", "f0.PA")
         strExc(0) = strExc(0) & " union " & _
            "SELECT (PA01||'-'||PA02||'-'||PA03||'-'||PA04) as soul," & _
            "nvl(PA05,nvl(pa06,pa07)) AS NAME1,PA01,PA02,PA03,PA04,PA09 FROM Patent f2 " & _
            "WHERE " & ChgPatent(strTmp) & " AND (PA09='013' or PA09='044') " & FMP2openSQL & _
            "ORDER BY PA01,PA02,PA03,PA04"
         strExc(0) = Replace(strExc(0), "f0.CP", "f2.PA")
      Else
        'Add by Lydia 2014/10/31 '設別名f0,+FMP2openSQL
         strExc(0) = "SELECT (SP01||'-'||SP02||'-'||SP03||'-'||SP04) as soul," & _
            "nvl(SP05,nvl(Sp06,Sp07)) AS NAME1,SP01,SP02,SP03,SP04,SP09 FROM " & _
            "SERVICEPRACTICE f0 WHERE " & ChgService(strTmp) & FMP2openSQL & _
            " AND SP09<>'000' ORDER BY SP01,SP02,SP03,SP04"
         strExc(0) = Replace(strExc(0), "f0.CP", "f0.SP")
      End If
      pmain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      If pmain.BOF And pmain.EOF Then
'         If FMP2open = True Then
'           MsgBox "權限不足 !", vbInformation
'         Else
           MsgBox "資料庫內無資料 !", vbInformation
'         End If
         Text4_GotFocus
      Else
         Set Adodc1.Recordset = pmain
         cmdOK_Click 0
      End If
   End If
End Sub

Private Sub Form_Activate()
   'Add by Sindy 2016/10/5
   If m_strIR01 <> "" And m_Done = False Then
      'Option1(0).Value = True
      'Text1.Text = m_AppNo
      Text2.Text = m_RDate
      'Command1.Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   'Add by Sindy 2016/10/5
   ElseIf m_AppNo <> "" And m_Done = False Then
      Option1(0).Value = True
      Text1.Text = m_AppNo
      Text2.Text = m_RDate
      Command1.Value = True
      m_Done = True
   End If
   '2016/10/5 END
End Sub

Private Sub Form_Load()
   'Me.Height = 6132
   'Me.Width = 9396
   MoveFormToCenter Me
   
   'cmdok(0).Enabled = False
   pmain.CursorLocation = adUseClient
   pmain1.CursorLocation = adUseClient
   Text2.Text = GetTaiwanTodayDate
   Option1_Click 0
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010505_1 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0 '申請案號
   Me.Text1.Enabled = True
   Me.Text4.Enabled = False
   Me.Text5.Enabled = False
   Me.Text6.Enabled = False
   Text1.SetFocus
Case 1 '本所案號
   Me.Text1.Enabled = False
   Me.Text4.Enabled = True
   Me.Text5.Enabled = True
   Me.Text6.Enabled = True
   Text4.SetFocus
End Select
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 <> "" Then
      If ChkDate(Text2) Then
         Text2 = TransDate(Text2, 1) 'Add by Morgan 2009/7/31 改可輸西元年但自動轉民國年
         If Val(Text2) > Val(strSrvDate(2)) Then
            MsgBox "來函收文日不可大於系統日 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
End Sub

Private Sub Text4_GotFocus()
TextInverse Me.Text4
End Sub

Private Sub Text5_GotFocus()
TextInverse Me.Text5
End Sub

Private Sub Text6_GotFocus()
TextInverse Me.Text6
End Sub
