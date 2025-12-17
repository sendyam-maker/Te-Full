VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm010029 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文－非主管機關"
   ClientHeight    =   5772
   ClientLeft      =   3780
   ClientTop       =   3648
   ClientWidth     =   8952
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   8952
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消確認"
      Height          =   345
      Index           =   6
      Left            =   7600
      Style           =   1  '圖片外觀
      TabIndex        =   27
      Top             =   390
      Width           =   1260
   End
   Begin VB.CheckBox Check1 
      Caption         =   "大宗發文案件(通知繳年費、通知逾期、通知實審)"
      Height          =   225
      Left            =   660
      TabIndex        =   26
      Top             =   1500
      Width           =   4245
   End
   Begin VB.ComboBox cboZone 
      Height          =   300
      ItemData        =   "frm010029.frx":0000
      Left            =   2235
      List            =   "frm010029.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   2
      Top             =   517
      Width           =   2490
   End
   Begin VB.TextBox txtCP27 
      Height          =   264
      Index           =   1
      Left            =   3555
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1170
      Width           =   1185
   End
   Begin VB.TextBox txtSystem 
      Height          =   264
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1800
      Width           =   732
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   2
      Left            =   4155
      MaxLength       =   2
      TabIndex        =   10
      Top             =   1800
      Width           =   492
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   1
      Left            =   3765
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1800
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   0
      Left            =   2535
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1800
      Width           =   1212
   End
   Begin VB.OptionButton Option1 
      Caption         =   "整批"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   6
      Top             =   1800
      Width           =   1320
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   345
      Index           =   5
      Left            =   7230
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   15
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "基本資料(&B)"
      Height          =   345
      Index           =   4
      Left            =   6150
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   15
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消發文(&C)"
      Height          =   345
      Index           =   3
      Left            =   6135
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   390
      Width           =   1260
   End
   Begin VB.TextBox txtType 
      Height          =   264
      Left            =   930
      MaxLength       =   1
      TabIndex        =   0
      Top             =   180
      Width           =   525
   End
   Begin VB.TextBox txtCP27 
      Height          =   264
      Index           =   0
      Left            =   2190
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1170
      Width           =   1140
   End
   Begin VB.ComboBox cboDept 
      Height          =   300
      ItemData        =   "frm010029.frx":0004
      Left            =   1785
      List            =   "frm010029.frx":0006
      Style           =   2  '單純下拉式
      TabIndex        =   3
      Top             =   840
      Width           =   2940
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定發文(&S)"
      Height          =   345
      Index           =   1
      Left            =   4860
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   390
      Width           =   1260
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   3570
      Left            =   30
      TabIndex        =   17
      Top             =   2160
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   6287
      _Version        =   393216
      Cols            =   10
      FixedCols       =   5
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "V|本所案號|發文號|案件名稱|案件性質|申請人|部門|人員|確認/判發時間|發文室發文時間"
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "重新查詢(&F)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   4890
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   15
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   2
      Left            =   7980
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   15
      Width           =   900
   End
   Begin VB.Label lblColorDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "親送"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   5310
      TabIndex        =   25
      Top             =   1830
      Width           =   390
   End
   Begin VB.Label lblColor 
      Appearance      =   0  '平面
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   5085
      TabIndex        =   24
      Top             =   1830
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "智權所(區)別："
      Height          =   180
      Index           =   3
      Left            =   990
      TabIndex        =   23
      Top             =   570
      Width           =   1200
   End
   Begin VB.Line Line1 
      X1              =   3285
      X2              =   3600
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "共　0　件"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   5
      Left            =   7695
      TabIndex        =   22
      Top             =   1845
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(1.未發文 2.所有資料 3.當日發文)"
      Height          =   180
      Index           =   4
      Left            =   1500
      TabIndex        =   21
      Top             =   225
      Width           =   2595
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "類　別："
      Height          =   180
      Index           =   1
      Left            =   165
      TabIndex        =   20
      Top             =   225
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " 專業部發文日期："
      Height          =   180
      Index           =   0
      Left            =   660
      TabIndex        =   19
      Top             =   1215
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專業部門別："
      Height          =   180
      Index           =   2
      Left            =   660
      TabIndex        =   18
      Top             =   900
      Width           =   1080
   End
End
Attribute VB_Name = "frm010029"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/14 Form2.0已修改 grdDataList
'Created by Morgan 2014/3/27
Option Explicit

Public cmdState As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim bolV As Boolean
Dim m_iRows As Integer
'Added by Lydia 2016/06/03
Dim bolChkDate As Boolean '是否詢問過17:30以後電子發文日期是否為翌日
Dim bolDate2 As Boolean


Private Sub cboDept_Click()
   If Option1(0).Value = False Then Option1(0).Value = True
End Sub

Private Sub Check1_Click()
   If Option1(1).Value = False Then
      cmdOK(0).Value = True
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData()
   Dim iRow As Integer
   Dim Str01 As String
   
   Select Case cmdState
   Case 0
      Screen.MousePointer = vbHourglass
      If TxtValidate = True Then OpenTable1
      Screen.MousePointer = vbDefault
   Case 1
      Screen.MousePointer = vbHourglass
      FormSave
      Screen.MousePointer = vbDefault
   Case 2
      Unload Me
   Case 3
      Screen.MousePointer = vbHourglass
      FormCancel
      PUB_SendMailCache
      Screen.MousePointer = vbDefault
   Case 4 '案件基本資料
      Me.Enabled = False
      With grdDataList
      For iRow = 1 To .Rows - 1
         If Trim(.TextMatrix(iRow, 0)) = "V" Then
            SelectRow iRow
            strExc(1) = GetValue(iRow, "本所案號")
            Str01 = SystemNumber(strExc(1), 1)
            If Mid(UCase(Str01), 1, 1) = "N" Then
               Str01 = Mid(Str01, 2, 3)
            End If
        
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            
            Screen.MousePointer = vbHourglass
            Select Case Pub_RplStr(Str01)
               Case "CFP", "FCP", "P"   '專利
                      frm100101_3.Show
                      frm100101_3.Tag = Pub_RplStr(strExc(1))
                      frm100101_3.StrMenu
                Case "CFT", "FCT", "T", "TF"   '商標
                      frm100101_4.Show
                      frm100101_4.Tag = Pub_RplStr(strExc(1))
                      frm100101_4.StrMenu
                Case "CFL", "FCL", "L"          '法務
                      frm100101_5.Show
                      frm100101_5.Tag = Pub_RplStr(strExc(1))
                      frm100101_5.StrMenu
                Case "LA"            '顧問
                      frm100101_6.Show
                      frm100101_6.Tag = Pub_RplStr(strExc(1))
                      frm100101_6.StrMenu
                Case Else                  '服務
                     Select Case Pub_RplStr(Str01)
                         Case "TB"    '條碼
                            frm100101_7.Show
                            frm100101_7.Tag = Pub_RplStr(strExc(1))
                            frm100101_7.StrMenu
                         Case "TM"
                            frm100101_8.Show
                            frm100101_8.Tag = Pub_RplStr(strExc(1))
                            frm100101_8.StrMenu
                         Case "TD"
                            frm100101_9.Show
                            frm100101_9.Tag = Pub_RplStr(strExc(1))
                            frm100101_9.StrMenu
                         Case "TC", "CFC"
                            frm100101_A.Show
                            frm100101_A.Tag = Pub_RplStr(strExc(1))
                            frm100101_A.StrMenu
                         Case Else
                            frm100101_B.Show
                            frm100101_B.Tag = Pub_RplStr(strExc(1))
                            frm100101_B.StrMenu
                      End Select
            End Select
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
      Next
      End With
      Me.Enabled = True
     
   Case 5 '案件進度
      Me.Enabled = False
      With grdDataList
      For iRow = 1 To .Rows - 1
         If Trim(.TextMatrix(iRow, 0)) = "V" Then
            SelectRow iRow
            strExc(1) = GetValue(iRow, "本所案號")
            Str01 = SystemNumber(strExc(1), 1)
            
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            
            Screen.MousePointer = vbHourglass
            frm100101_2.Show
            frm100101_2.Tag = Pub_RplStr(strExc(1))
            frm100101_2.StrMenu
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
      Next
      End With
      Me.Enabled = True
   'Add by Amy 2017/12/06 取消確認
   Case 6
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      If CancelLP = True Then
        cmdok_Click (0)
      End If
      Me.Enabled = True
      Screen.MousePointer = vbDefault
   End Select
End Sub

Private Function CancelLP() As Boolean
    Dim iRow As Integer, bolV As Boolean, bolOK As Boolean
    Dim strMsg As String
On Error GoTo ErrHnd
   
    bolV = False
    bolOK = True
    For iRow = 1 To grdDataList.Rows - 1
        If grdDataList.TextMatrix(iRow, 0) = "V" Then
            bolV = True
            'Added by Morgan 2022/5/16 發文室已發文不可取消確認(分所案件由電腦中心人員從進度查詢取消)
            If Trim(GetValue(iRow, "發文室發文時間")) <> "" Then
               bolOK = False
               strMsg = "勾選的資料中，有發文室已發文資料，請重新確認！"
               Exit For
            End If
            'end 2022/5/16
            If Trim(GetValue(iRow, "LP07")) = "0" Then
                bolOK = False
                strMsg = "勾選的資料中，未有確認時間，請重新確認！"
                Exit For
            End If
            If Trim(GetValue(iRow, "直寄")) = "直寄" Then
                bolOK = False
                strMsg = "勾選的資料中，有直寄資料，直寄不可取消！"
                Exit For
             End If
        End If
    Next
   
    If bolV = False Then
        MsgBox "請勾選欲發文的資料！", vbExclamation + vbOKOnly
        Exit Function
    End If
    If bolOK = False Then
        MsgBox strMsg, vbExclamation + vbOKOnly
        Exit Function
    End If
  
    For iRow = 1 To grdDataList.Rows - 1
        If grdDataList.TextMatrix(iRow, 0) = "V" Then
            strExc(1) = GetValue(iRow, "cp09")
            'Modified by Morgan 2019/12/17 確認人員改要保留否則寄送確認會看不到
            'strSql = "Update LetterProgress Set LP06=null,LP07=0,LP11=null Where LP01='" & strExc(1) & "'"
            strSql = "Update LetterProgress Set LP07=0,LP11=null Where LP01='" & strExc(1) & "'"
            'end 2019/12/17

            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql, intI
            SelectRow iRow
        End If
    Next
   
    CancelLP = True
    Exit Function
   
ErrHnd:
    grdDataList.Visible = True
    MsgBox Err.Description, vbCritical
   
End Function

Private Function FormCancel() As Boolean
   Dim iRow As Integer, bolV As Boolean, bolOK As Boolean
   Dim strCP131 As String, strCP132 As String, strCNoList As String
   
On Error GoTo ErrHnd
   
   bolV = False
   bolOK = True
   With grdDataList
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         bolV = True
         If Trim(GetValue(iRow, "發文室發文時間")) = "" Then
            bolOK = False
            Exit For
         End If
         strCNoList = IIf(strCNoList <> "", strCNoList & ",", "") & GetValue(iRow, "cp09")
      End If
   Next
   End With
   
   If bolV = False Then
      MsgBox "請勾選欲取消發文的資料！", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If bolOK = False Then
      MsgBox "勾選的資料中，未有發文室發文時間，請重新確認！", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   '開啟取消發文視窗
   frm010029_1.strNoList = strCNoList '收文號清單
   If frm010029_1.CheckShowList Then
      Screen.MousePointer = vbDefault
      frm010029_1.Show vbModal
   End If
   strCP131 = frm010029_1.strCP131
   strCP132 = frm010029_1.strCP132
   bolOK = frm010029_1.bolOK
   Unload frm010029_1
   Set frm010029_1 = Nothing
   
   If bolOK = False Then
      Exit Function
   End If
   
   Screen.MousePointer = vbHourglass
   
   With grdDataList
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         cnnConnection.BeginTrans
On Error GoTo ErrHndT

         strExc(1) = GetValue(iRow, "cp09")
         'Modified by Morgan 2015/1/6
         '因P案發文電子化，發文號可能是申請書發文更新的故不可取消
         'strSql = "update caseprogress set cp28=null,cp127=null,cp128=null,cp131='" & ChgSQL(strCP131) & "',cp132=" & DBDATE(strCP132) & " where cp09='" & strExc(1) & "'"
         strSql = "update caseprogress set cp127=null,cp128=null,cp131='" & ChgSQL(strCP131) & "',cp132=" & DBDATE(strCP132) & " where cp09='" & strExc(1) & "'"
         cnnConnection.Execute strSql, intI
         
         'Added by Morgan 2016/6/15
         '直寄郵件若智權人員已確認時要清除並EMail通知 Ex.P111549
         strExc(0) = "select lp06 from letterprogress where lp01='" & strExc(1) & "' and lp07>0 and lp11='Y'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modified by Morgan 2022/4/25 確認人不可清除
            'strSql = "update letterprogress set lp06=null,lp07=0 where lp01='" & strExc(1) & "'"
            strSql = "update letterprogress set lp07=0 where lp01='" & strExc(1) & "'"
            'end 2022/4/25
            cnnConnection.Execute strSql, intI
            
            strExc(2) = GetValue(iRow, "本所案號") & "的客戶通知函(" & GetValue(iRow, "案件性質") & PUB_GetRelateCasePropertyName(GetValue(iRow, "CP09"), "1") & ")，發文室已取消發文!!"
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                     " values ('" & strUserNum & "','" & RsTemp("lp06") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                     ",'" & strExc(2) & "','如旨')"
            cnnConnection.Execute strSql, intI
         End If
         'end 2016/6/15
         
         '直寄且非自行判發的要一併取消判發日
         'Removed by Morgan 2016/5/27 發文室可能選錯案或改隔天發文取消非專業部抽回，若要重判發改人工通知電腦中心取消判發日
         'strSql = "update letterprogress set lp05=0 where lp01='" & strExc(1) & "' and lp04 is not null"
         'cnnConnection.Execute strSql, intI
         'end 2016/5/27

         cnnConnection.CommitTrans
         
         SetValue iRow, "發文室發文時間", ""
         SetValue iRow, "發文號", ""
         SelectRow iRow
      End If
   Next
   End With
   
   FormCancel = True
   Exit Function
   
ErrHndT:
   cnnConnection.RollbackTrans
   
ErrHnd:
   grdDataList.Visible = True
   MsgBox Err.Description, vbCritical
End Function

Private Function FormSave() As Boolean
   Dim iRow As Integer, bolV As Boolean, bolOK As Boolean
   'Added by Lydia 2016/06/03
   Dim strNewCP127 As String
   Dim strTime As String

On Error GoTo ErrHnd
   
   bolV = False
   bolOK = True
   With grdDataList
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         bolV = True
         If Trim(GetValue(iRow, "發文室發文時間")) <> "" Then
            bolOK = False
            Exit For
         End If
      End If
   Next
   End With
   
   If bolV = False Then
      MsgBox "請勾選欲發文的資料！", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If bolOK = False Then
      MsgBox "勾選的資料中，已有發文室發文時間，請重新確認！", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   'Added by Lydia 2016/06/03 發文室17:30以後電子發文日期為翌日
    strTime = ServerTime
    strNewCP127 = strSrvDate(1)
    If bolChkDate = False And Val(strTime) >= 173000 Then '只詢問一次
       bolChkDate = True
       If MsgBox("下午5:30以後，發文日期是否為次日?", vbInformation + vbYesNo) = vbYes Then bolDate2 = True
    End If
    If bolDate2 Then
       strTime = "000001"
       strNewCP127 = CompWorkDay(2, strNewCP127)
    End If
    'end 2016/06/03
   With grdDataList
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         cnnConnection.BeginTrans
On Error GoTo ErrHndT

         strExc(1) = GetValue(iRow, "cp09")
         'Modified by Lydia 2016/06/03
         'strSql = "update caseprogress set cp28=cp09,cp127='" & strSrvDate(1) & "',cp128=to_char(sysdate,'HH24MISS') where cp09='" & strExc(1) & "'"
         strSql = "update caseprogress set cp28=cp09,cp127='" & strNewCP127 & "',cp128='" & strTime & "' where cp09='" & strExc(1) & "'"
         cnnConnection.Execute strSql, intI
         cnnConnection.CommitTrans
         
         If txtType = "1" Then
            .TextMatrix(iRow, 0) = "X"
            .RowHeight(iRow) = 0
            m_iRows = m_iRows - 1
            Label1(5).Caption = "共　" & m_iRows & "　件"
         Else
            strExc(0) = "select sqldatet(cp127)||' '||sqltime6(cp128) 發文室發文時間,substr(cp28,4) 發文號 from caseprogress where cp09='" & strExc(1) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               SetValue iRow, "發文室發文時間", RsTemp(0)
               'Modified by Morgan 2014/7/2 改查詢時直接帶收文號後6碼
               'SetValue iRow, "發文號", RsTemp(1)
            End If
            SelectRow iRow
         End If
      End If
   Next
   End With
   
   FormSave = True
   Exit Function
   
ErrHndT:
   cnnConnection.RollbackTrans
   
ErrHnd:
   grdDataList.Visible = True
   MsgBox Err.Description, vbCritical
   
End Function

Private Function GetValue(pRow As Integer, pFieldName As String) As String
   Dim iRow As Integer
   With grdDataList
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetValue = .TextMatrix(pRow, iRow)
         Exit For
      End If
   Next
   End With
End Function

Private Function SetValue(pRow As Integer, pFieldName As String, pValue As String) As Boolean
   Dim ii As Integer
   With grdDataList
   For ii = 0 To .Cols - 1
      If UCase(.TextMatrix(0, ii)) = UCase(pFieldName) Then
         .TextMatrix(pRow, ii) = pValue
         SetValue = True
         Exit Function
      End If
   Next
   End With
End Function

Private Sub Form_Activate()
   Static bolDone As Boolean
   If bolDone = False Then
      Screen.MousePointer = vbHourglass
      cmdOK(0).Value = True
      Screen.MousePointer = vbDefault
      bolDone = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   '類別
   txtType.Text = "1"
   '整批
   Option1(1).Value = True
   
   '智權所(區)別
   cboZone.Clear
   cboZone.AddItem "其他", 0
   cboZone.ItemData(0) = 99
   cboZone.AddItem "高所", 0
   cboZone.ItemData(0) = 4
   cboZone.AddItem "南所", 0
   cboZone.ItemData(0) = 3
   cboZone.AddItem "中所", 0
   cboZone.ItemData(0) = 2
   cboZone.AddItem "北五", 0
   cboZone.ItemData(0) = 15
   cboZone.AddItem "北四", 0
   cboZone.ItemData(0) = 14
   cboZone.AddItem "北三", 0
   cboZone.ItemData(0) = 13
   cboZone.AddItem "北一", 0
   cboZone.ItemData(0) = 11
   cboZone.AddItem "全部", 0
   cboZone.ItemData(0) = 0
   cboZone.ListIndex = 0
   
   '部門
   cboDept.Clear
   'Add By Sindy 2018/9/18
   cboDept.AddItem "商標處", 0
   cboDept.ItemData(0) = 2
   '2018/9/18 END
   cboDept.AddItem "內專", 0
   cboDept.ItemData(0) = 1
   cboDept.AddItem "全部", 0
   cboDept.ItemData(0) = 0
   cboDept.ListIndex = 0
   
   'Added by Morgan 2020/5/12 配合中所發文測試,非北所人員操作時預設內專
   If pub_strUserOffice <> "1" Then
      cboDept.ListIndex = 1
   End If
   'end 2020/5/12
   
   '發文日
   'txtCP27(0).Text = TransDate(CompWorkDay(2, strSrvDate(1), 1), 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm010029 = Nothing
End Sub

Private Sub OpenTable1()
   Dim stSQL As String, ii As Integer, jj As Integer, iColId As Integer
   Dim strCon As String
   Dim stSQL2 As String, strVal As String 'Add By Sindy 2018/9/19
   Dim strBase As String, strField As String 'Add by Amy 2020/02/05
   
   SetGrid True
   Label1(5).Caption = "共　0　件"
   
   'Modified by Morgan 2014/5/22 特殊設定 A7 的編號照北所的流程
   'Modified by Morgan 2015/6/29 排除發文人員為QPGMR(E化系統自動發文)
   'Modified by Morgan 2015/9/1 +cp10,接洽人
   'Modified by Morgan 2015/10/8 接洽人欄位+lp31判斷
   'Modified by Morgan 2016/3/3 +cp27
   'Modified by Morgan 2016/4/14
   '客戶編號只需抓前8碼判斷
   'Modified by Morgan 2016/6/1 +非臺灣案案件性質
   'Modified by Morgan 2016/11/15 +接洽人先抓若LP33,LP34(副本會放)
   'Modified by Morgan 2017/4/11 LP33也只能抓前8碼才能比對
   'Modify by Amy 2017/12/07 +LP07
   'Modified by Morgan 2018/10/31 +FC代理人--李佳寶
   'Modified by Morgan 2019/3/14 +改判斷發文日>19221111(原>0) Ex:CFP-030433
   'Modified by Morgan 2019/4/16 +lp28
   'Modified by Morgan 2019/4/30 NP欄位沒用了不必再串NP(也會錯,CFP-029129-0-00)
   'Modified by Morgan 2019/9/24 因改先設確認人員,加判斷有確認日期才帶
   'Mark by Amy 2020/02/05 往下搬,修改
'   stSQL = "select '' V,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04 本所案號,substr(c1.cp09,-6) 發文號" & _
'      ",nvl(pa05,nvl(pa06,pa07)) 案件名稱,Decode(pa09,'000',CPM03,CPM04) 案件性質,NVL(CU04,CU05) 申請人,nvl(a3.a0902,nvl(a2.a0902,a1.A0902)) 部門" & _
'      ",decode(lp07,0,nvl(s2.st02,s1.st02),s3.st02) 人員,decode(LP11,'Y','直寄') 直寄,nvl(fa04,nvl(rtrim(fa05||''||fa63||''||fa64||''||fa65),fa06)) FC代理人" & _
'      ",decode(lp07,0,sqldatet(lp05)||' '||sqltime6(lp17) ,sqldatet(lp07)||' '||sqltime6(lp18) ) 確認時間" & _
'      ",sqldatet(c1.cp127)||' '||sqltime6(c1.cp128) 發文室發文時間,c1.cp09 cp09,lp11,decode(c1.cp10,'990',c2.cp10,c1.cp10) cp10" & _
'      ",nvl(substr(lp33,1,8)||lp34,decode(lp31,'Y',substr(pa75,1,8),substr(pa26,1,8)||nvl(pa149,cu127))) 接洽人" & _
'      ",c1.cp27,lp07,c1.cp01 cp01,c1.cp02 cp02,c1.cp03 cp03,c1.cp04 cp04,to_char(lp28,'yyyymmdd') lp28" & _
'      " From letterprogress,caseprogress c1, SetSpecMan, staff s1,staff s2,staff s3,staff s4,acc090 a1" & _
'      ",acc090 a2,acc090 a3,patent, casepropertymap, customer,caseprogress c2,fagent" & _
'      " WHERE c1.cp27>19221111 and NVL(c1.cp154,' ')<>'QPGMR' and c2.cp09(+)=c1.cp43 and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9)"
   
   '本所案號
   If Option1(1).Value = True Then
     'Moidfy by Amy 2020/02/05 原:stSQL
      strBase = strBase & " and c1.cp01='" & txtSystem & "' and c1.cp02='" & txtCode(0) & "' and c1.cp03='" & txtCode(1) & "' and c1.cp04='" & txtCode(2) & "'"
      strCon = strCon & " and c1.cp01='" & txtSystem & "' and c1.cp02='" & txtCode(0) & "' and c1.cp03='" & txtCode(1) & "' and c1.cp04='" & txtCode(2) & "'" 'Add By Sindy 2018/9/18
   Else
      '當日發文
      If txtType = "3" Then
         'Moidfy by Amy 2020/02/05 原:stSQL
         strBase = strBase & " and c1.cp127=" & strSrvDate(1)
         strCon = strCon & " and c1.cp127=" & strSrvDate(1) 'Add By Sindy 2018/9/18
      End If
      
      '專業部發文日
      If txtCP27(0) <> "" Then
         'Moidfy by Amy 2020/02/05 原:stSQL
         strBase = strBase & " and c1.cp27>=" & DBDATE(txtCP27(0))
         strCon = strCon & " and c1.cp27>=" & DBDATE(txtCP27(0)) 'Add By Sindy 2018/9/18
      End If
      If txtCP27(1) <> "" Then
        'Moidfy by Amy 2020/02/05 原:stSQL
         strBase = strBase & " and c1.cp27<=" & DBDATE(txtCP27(1))
         strCon = strCon & " and c1.cp27<=" & DBDATE(txtCP27(1)) 'Add By Sindy 2018/9/18
      End If
      
      '專業部門
      '內專
      If cboDept.ItemData(cboDept.ListIndex) = 1 Then
         stSQL = stSQL & " and substr(s1.st03,1,2)='P1'"
         strCon = strCon & " and substr(s1.st03,1,2)='P1'"
         strBase = strBase & " and substr(s1.st03,1,2)='P1'" 'Add By Sindy 2025/4/18
      'Add By Sindy 2018/9/18
      ElseIf cboDept.ItemData(cboDept.ListIndex) = 2 Then
         stSQL = stSQL & " and substr(s1.st03,1,2)='P2'"
         strCon = strCon & " and substr(s1.st03,1,2)='P2'"
         strBase = strBase & " and substr(s1.st03,1,2)='P2'" 'Add By Sindy 2025/4/18
      '2018/9/18 END
      End If
      
      '智權所(區)
      If cboZone.ItemData(cboZone.ListIndex) > 0 Then
         '分所
         If cboZone.ItemData(cboZone.ListIndex) < 10 Then
            'Modified by Morgan 2024/5/29 P1004案件照分所方式處理
            stSQL = stSQL & " and s4.st06='" & cboZone.ItemData(cboZone.ListIndex) & "' and (instr(';'||replace(oMan,',',';')||';',';'||c1.cp13||';')=0 or lp06='P1004')"
            strCon = strCon & " and s4.st06='" & cboZone.ItemData(cboZone.ListIndex) & "'" 'Add By Sindy 2018/9/18
         '北所
         ElseIf cboZone.ItemData(cboZone.ListIndex) < 20 Then
            stSQL = stSQL & " and s4.st15>='S" & cboZone.ItemData(cboZone.ListIndex) & "' and s4.st15<='S" & cboZone.ItemData(cboZone.ListIndex) & "9'"
            strCon = strCon & " and s4.st15>='S" & cboZone.ItemData(cboZone.ListIndex) & "' and s4.st15<='S" & cboZone.ItemData(cboZone.ListIndex) & "9'" 'Add By Sindy 2018/9/18
         '其他
         Else
            stSQL = stSQL & " and (instr(';'||replace(oMan,',',';')||';',';'||c1.cp13||';')>0 or not ((s4.st06>'1' and s4.st06<'5') or (s4.st15>='S11' and s4.st15<='S159')))"
            strCon = strCon & " and not((s4.st06>'1' and s4.st06<'5') or (s4.st15>='S11' and s4.st15<='S159'))" 'Add By Sindy 2018/9/18
         End If
      End If
      
      'Added by Morgan 2015/11/10
      'Modified by Morgan 2016/7/7 非臺灣案的年費逾期通知非整批,改判斷有無檢核人員
      'modify by sonia 2016/7/13 整批未檢核時,發文室也會跑到非整批,故改寫法
      'Modified by Morgan 2016/7/21 整批判斷有檢核,非整批的排除期限通知及臺灣的年費逾期通知
      If Check1.Value = vbChecked Then
         'stSQL = stSQL & " and cp10 in ('1913','1605')"
         'stSQL = stSQL & " and lp27 is not null"
         'stSQL = stSQL & " and (cp10='1913' or (cp10='1605' and lp27 is not null))"
         'Modified by Morgan 2016/11/8 考慮單筆催年費不算大宗再+lp32='Y'
         'stSQL = stSQL & " and lp27 is not null"
         'Moidfy by Amy 2020/02/05 原:stSQL
         strBase = strBase & " and lp27 is not null and lp32='Y'"
      Else
         'stSQL = stSQL & " and cp10 not in ('1913','1605')"
         'stSQL = stSQL & " and lp27 is null"
         'stSQL = stSQL & " and lp27 is null and cp10<>'1913'"
         'Modified by Morgan 2016/8/9 通知期限標準專利批准記錄請求也是用1913但非整批,加判斷案件性質為605或416
         'stSQL = stSQL & " and cp10<>'1913' and not (pa09='000' and cp10='1605')"
         'Modified by Morgan 2016/10/19 整批有通知進入國家階段(119),再改排除'111'
         'Modified by Morgan 2016/11/8 考慮單筆催年費不算大宗改判斷 lp32 is null
         'stSQL = stSQL & " and not (cp10='1913' and np07<>'111') and not (pa09='000' and cp10='1605')"
         'Moidfy by Amy 2020/02/05 原:stSQL
         strBase = strBase & " and lp32 is null"
      End If
   End If
      
   '未發文
   If txtType = "1" Then
      'Moidfy by Amy 2020/02/05 原:stSQL
      strBase = strBase & " and lp05>0 and lp10='Y' and lp15='N' and c1.cp09(+)=lp01"
      strCon = strCon & " and nvl(c1.cp127,0)=0" 'Add By Sindy 2018/9/18
   Else
      'Moidfy by Amy 2020/02/05 原:stSQL
      strBase = strBase & " and lp01(+)=c1.cp09 and lp05>0 and lp10='Y'"
      strCon = strCon & " and nvl(c1.cp127,0)>0" 'Add By Sindy 2018/9/18
   End If
   
   'Add by Amy 2020/02/05 從上搬下來改
   'Modified by Morgan 2020/8/3 ServicePractice 移到外層
   'strField = ",Nvl(pa05,Nvl(pa06,pa07)) as PA05,pa09,pa26,pa75,pa149"
   strField = ",Decode(pa01,null,Nvl(sp05,Nvl(sp06,sp07)),Nvl(pa05,Nvl(pa06,pa07))) as PA05,nvl(pa09,sp09) as PA09,nvl(pa26,sp08) pa26,nvl(pa75,sp26) pa75,nvl(pa149,sp78) pa149"
   'end 2020/8/3
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
        strField = ",Decode(pa01,null,Decode(tm01,null,Nvl(sp05,Nvl(sp06,sp07)),Nvl(tm05,Nvl(tm06,tm07))),Nvl(pa05,Nvl(pa06,pa07))) as PA05" & _
                        ",Decode(pa01,null,Nvl(tm10,sp09),pa09) as PA09,Decode(pa01,null,Nvl(tm23,sp08),pa26) as PA26,Decode(pa01,null,Nvl(tm44,sp26),pa75) as PA75" & _
                        ",Decode(pa01,null,Nvl(tm123,sp78),pa149) as PA149"
        'Modified by Morgan 2020/8/3 ServicePractice 移到外層
        'strBase = strBase & " And tm01(+)=C1.Cp01 And tm02(+)=C1.Cp02 And tm03(+)=C1.Cp03 And tm04(+)=C1.Cp04 " & _
                                   " And sp01(+)=C1.Cp01 And sp02(+)=C1.Cp02 And sp03(+)=C1.Cp03 And sp04(+)=C1.Cp04 "
         strBase = strBase & " And tm01(+)=C1.Cp01 And tm02(+)=C1.Cp02 And tm03(+)=C1.Cp03 And tm04(+)=C1.Cp04 "
         'end 2020/8/3
   End If
   'Modified by Morgan 2020/8/3 ServicePractice 移到外層
   strBase = "Select LetterProgress.*,C1.*" & strField & " From LetterProgress,CaseProgress c1, Patent,servicepractice,staff s1 " & _
                  IIf(strSrvDate(1) >= T商標電子化第2階段啟用日, ",Trademark", "") & _
                  " Where nvl(lp11,'1')<>'2' And  c1.cp27>19221111 And NVL(c1.cp154,' ')<>'QPGMR' and s1.st01(+)=c1.cp83 " & strBase & _
                  " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04" & _
                  " and sp01(+)=c1.cp01 and sp02(+)=c1.cp02 and sp03(+)=c1.cp03 and sp04(+)=c1.cp04"
  
   'Modify by Amy 2020/02/06 + stSQL-bug 區別無作用
   'Modify by Amy 2020/03/02 +GetRelateCasePropertyName(c1.cp09, '1')
   'Added by Lydia 2023/12/26
   If strSrvDate(1) >= 新部門啟用日 Then
      stSQL = "select '' V,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04 本所案號,substr(c1.cp09,-6) 發文號" & _
         ",PA05 案件名稱,Decode(pa09,'000',CPM03,CPM04)||GetRelateCasePropertyName(c1.cp09, '1') 案件性質,NVL(CU04,CU05) 申請人,nvl(b3.a0922,nvl(b2.a0922,nvl(b1.a0922,nvl(a3.a0902,nvl(a2.a0902,a1.A0902))))) 部門" & _
         ",decode(lp07,0,nvl(s2.st02,s1.st02),s3.st02) 人員,decode(LP11,'Y','直寄') 直寄,nvl(fa04,nvl(rtrim(fa05||''||fa63||''||fa64||''||fa65),fa06)) FC代理人" & _
         ",decode(lp07,0,sqldatet(lp05)||' '||sqltime6(lp17) ,sqldatet(lp07)||' '||sqltime6(lp18) ) 確認時間" & _
         ",sqldatet(c1.cp127)||' '||sqltime6(c1.cp128) 發文室發文時間,c1.cp09 cp09,lp11,decode(c1.cp10,'990',c2.cp10,c1.cp10) cp10" & _
         ",nvl(substr(lp33,1,8)||lp34,decode(lp31,'Y',substr(pa75,1,8),substr(pa26,1,8)||nvl(pa149,cu127))) 接洽人" & _
         ",c1.cp27,lp07,c1.cp01 cp01,c1.cp02 cp02,c1.cp03 cp03,c1.cp04 cp04,to_char(lp28,'yyyymmdd') lp28" & _
         " From (" & strBase & ") c1, SetSpecMan, staff s1,staff s2,staff s3,staff s4,acc090 a1" & _
         ",acc090 a2,acc090 a3,casepropertymap, customer,caseprogress c2,fagent, acc090new b1, acc090new b2, acc090new b3" & _
         " WHERE c2.cp09(+)=c1.cp43 and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9)" & stSQL
   Else
   'end 2023/12/26
      stSQL = "select '' V,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04 本所案號,substr(c1.cp09,-6) 發文號" & _
         ",PA05 案件名稱,Decode(pa09,'000',CPM03,CPM04)||GetRelateCasePropertyName(c1.cp09, '1') 案件性質,NVL(CU04,CU05) 申請人,nvl(a3.a0902,nvl(a2.a0902,a1.A0902)) 部門" & _
         ",decode(lp07,0,nvl(s2.st02,s1.st02),s3.st02) 人員,decode(LP11,'Y','直寄') 直寄,nvl(fa04,nvl(rtrim(fa05||''||fa63||''||fa64||''||fa65),fa06)) FC代理人" & _
         ",decode(lp07,0,sqldatet(lp05)||' '||sqltime6(lp17) ,sqldatet(lp07)||' '||sqltime6(lp18) ) 確認時間" & _
         ",sqldatet(c1.cp127)||' '||sqltime6(c1.cp128) 發文室發文時間,c1.cp09 cp09,lp11,decode(c1.cp10,'990',c2.cp10,c1.cp10) cp10" & _
         ",nvl(substr(lp33,1,8)||lp34,decode(lp31,'Y',substr(pa75,1,8),substr(pa26,1,8)||nvl(pa149,cu127))) 接洽人" & _
         ",c1.cp27,lp07,c1.cp01 cp01,c1.cp02 cp02,c1.cp03 cp03,c1.cp04 cp04,to_char(lp28,'yyyymmdd') lp28" & _
         " From (" & strBase & ") c1, SetSpecMan, staff s1,staff s2,staff s3,staff s4,acc090 a1" & _
         ",acc090 a2,acc090 a3,casepropertymap, customer,caseprogress c2,fagent" & _
         " WHERE c2.cp09(+)=c1.cp43 and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9)" & stSQL
      'end 2020/02/05
   End If
   
   '已判發且為A.直寄(LP11='Y') 或 B.分所案件 或C.已確認的來函資料。排除不寄(LP11<>'2')的。
   'Modified by Morgan 2019/4/30 NP欄位沒用了不必再串NP(也會錯,CFP-029129-0-00)
   'Modify by Amy 2020/02/05 and nvl(lp11,'1')<>'2' and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04往上搬
   'Added by Lydia 2023/12/26
   If strSrvDate(1) >= 新部門啟用日 Then
      'Modified by Morgan 2024/5/29 P1004案件照分所方式處理
      'Modified by Morgan 2025/5/28 所別應以確認人判斷 s4.st01(+)=c1.cp13-->s4.st01(+)=lp06 Ex:T-192813(AB3036037)
      'Modified by Morgan 2025/6/9 s4.st01(+)=lp06 --> and s4.st01(+)=decode(lp06,'QPGMR',c1.cp13,lp06) Ex:P-132280(CB4030819)
      stSQL = stSQL & " and ocode(+)='A7' and s1.st01(+)=c1.cp83 and s2.st01(+)=lp04" & _
         " and s3.st01(+)=lp06 and s4.st01(+)=decode(lp06,'QPGMR',c1.cp13,lp06) AND a1.A0901(+)=s1.ST03 AND a2.A0901(+)=s2.ST03" & _
         " AND a3.A0901(+)=s3.ST03 and (lp11='Y' or (s4.st06>'1' and s4.st06<'5' and (instr(';'||replace(oMan,',',';')||';',';'||c1.cp13||';')=0 or lp06='P1004')) or lp07>0)" & _
         " and cpm01(+)=c1.cp01 and cpm02(+)=c1.cp10 AND CU01(+)=substr(PA26,1,8)" & _
         " AND CU02(+)=substr(PA26,9) and b1.a0921(+)=s1.st93 and b2.a0921(+)=s2.st93 and b3.a0921(+)=s3.st93"
   Else
   'end 2023/12/26
      stSQL = stSQL & " and ocode(+)='A7' and s1.st01(+)=c1.cp83 and s2.st01(+)=lp04" & _
         " and s3.st01(+)=lp06 and s4.st01(+)=c1.cp13 AND a1.A0901(+)=s1.ST03 AND a2.A0901(+)=s2.ST03" & _
         " AND a3.A0901(+)=s3.ST03 and (lp11='Y' or (s4.st06>'1' and s4.st06<'5' and instr(';'||replace(oMan,',',';')||';',';'||c1.cp13||';')=0) or lp07>0)" & _
         " and cpm01(+)=c1.cp01 and cpm02(+)=c1.cp10 AND CU01(+)=substr(PA26,1,8)" & _
         " AND CU02(+)=substr(PA26,9)"
   End If
'Add By Sindy 2020/2/18
If strSrvDate(1) < T商標電子化第2階段啟用日 Then
'2020/2/18 END
   'Add By Sindy 2018/9/18
   '非國外部收文之C類來函承辦人為內商人員
'   strVal = "select e2.eep01 eep01,max(e2.eep02) eep02 from empelectronprocess e2" & _
'            " where e2.eep01 in(" & _
'            "select e1.eep01 from empelectronprocess e1" & _
'            " where substr(e1.eep01,1,1)='C' and e1.eep04='" & EMP_發文歸檔 & "'" & _
'            ") and e2.eep04='" & EMP_判發 & "' group by e2.eep01"
   strVal = "select e2.eep01 eep01,max(e2.eep02) eep02 from empelectronprocess e2" & _
            " where e2.eep01 in(" & _
            "select e1.eep01 from empelectronprocess e1" & _
            " where substr(e1.eep01,1,1)='C' and e1.eep04='" & EMP_發文歸檔 & "'" & _
            ") and e2.eep04='" & EMP_判發 & "' group by e2.eep01" & _
            " union select e1.eep01,e1.eep02 from empelectronprocess e1" & _
            " where substr(e1.eep01,1,1)='C' and e1.eep04='" & EMP_發文歸檔 & "'" & _
            " and not exists(select e2.eep01 from empelectronprocess e2 where e2.eep01=e1.eep01 and e2.eep04='" & EMP_判發 & "')"
   '依歷程讀取資料
   '20180925才開始經發文室線上勾選(非國外部收文之C類來函承辦人為內商人員, 有走歷程者)
   'Modified by Morgan 2018/9/28 +有期限設直寄
   'Modified by Morgan 2018/10/31+FC代理人--李佳寶
   'Modified by Morgan 2019/4/16 +lp28
   'Added by Lydia 2023/12/26
   If strSrvDate(1) >= 新部門啟用日 Then
      'Modified by Morgan 2025/5/28 所別應以確認人判斷 s4.st01(+)=c1.cp13-->s4.st01(+)=lp06
      stSQL2 = " union select '' V,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04 本所案號" & _
         ",substr(c1.cp09,-6) 發文號,tm05 案件名稱,Decode(tm10,'000',CPM03,CPM04) 案件性質" & _
         ",NVL(CU04,CU05) 申請人,nvl(b2.a0922,nvl(b2.a0922,nvl(a2.a0902,a1.A0902))) 部門" & _
         ",nvl(s1.st02,s2.st02) 人員,decode(c1.cp07,null,'','直寄') 直寄" & _
         ",nvl(fa04,nvl(rtrim(fa05||''||fa63||''||fa64||''||fa65),fa06)) FC代理人" & _
         ",sqldatet(e4.eep06)||' '||sqltime6(e4.eep07) 確認時間" & _
         ",sqldatet(c1.cp127)||' '||sqltime6(c1.cp128) 發文室發文時間" & _
         ",c1.cp09 cp09,'' lp11,decode(c1.cp10,'990',c2.cp10,c1.cp10) cp10" & _
         ",'' 接洽人" & _
         ",c1.cp27,e4.eep06 lp07,c1.cp01 cp01,c1.cp02 cp02,c1.cp03 cp03,c1.cp04 cp04,'' lp28" & _
         " From empelectronprocess e3,empelectronprocess e4,(" & strVal & ") T" & _
         ",caseprogress c1,caseprogress c2,staff s1,staff s2,staff s4,acc090 a1" & _
         ",acc090 a2,trademark,casepropertymap,customer,fagent, acc090new b1, acc090new b2" & _
         " WHERE e3.eep01=t.eep01 and e3.eep04='" & EMP_發文歸檔 & "'" & _
         " and e3.eep01=c1.cp09(+) and c1.cp27>=20180925 and substr(c1.cp12,1,1)<>'F' and c2.cp09(+)=c1.cp43" & _
         " and e4.eep01=t.eep01 and e4.eep02=t.eep02" & _
         " and s1.st01(+)=c1.cp83 and s2.st01(+)=e4.eep03 and s4.st01(+)=lp06" & _
         " and a1.A0901(+)=s1.ST03 AND a2.A0901(+)=s2.ST03" & _
         " and tm01(+)=c1.cp01 and tm02(+)=c1.cp02 and tm03(+)=c1.cp03 and tm04(+)=c1.cp04" & _
         " and cpm01(+)=c1.cp01 and cpm02(+)=c1.cp10" & _
         " and CU01(+)=substr(TM23,1,8) AND CU02(+)=substr(TM23,9)" & strCon & _
         " and fa01(+)=substr(tm44,1,8) and fa02(+)=substr(tm44,9) and b1.a0921(+)=s1.st93 and b2.a0921(+)=s2.st93"
   Else
   'end 2023/12/26
      stSQL2 = " union select '' V,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04 本所案號" & _
         ",substr(c1.cp09,-6) 發文號,tm05 案件名稱,Decode(tm10,'000',CPM03,CPM04) 案件性質" & _
         ",NVL(CU04,CU05) 申請人,nvl(a2.a0902,a1.A0902) 部門" & _
         ",nvl(s1.st02,s2.st02) 人員,decode(c1.cp07,null,'','直寄') 直寄" & _
         ",nvl(fa04,nvl(rtrim(fa05||''||fa63||''||fa64||''||fa65),fa06)) FC代理人" & _
         ",sqldatet(e4.eep06)||' '||sqltime6(e4.eep07) 確認時間" & _
         ",sqldatet(c1.cp127)||' '||sqltime6(c1.cp128) 發文室發文時間" & _
         ",c1.cp09 cp09,'' lp11,decode(c1.cp10,'990',c2.cp10,c1.cp10) cp10" & _
         ",'' 接洽人" & _
         ",c1.cp27,e4.eep06 lp07,c1.cp01 cp01,c1.cp02 cp02,c1.cp03 cp03,c1.cp04 cp04,'' lp28" & _
         " From empelectronprocess e3,empelectronprocess e4,(" & strVal & ") T" & _
         ",caseprogress c1,caseprogress c2,staff s1,staff s2,staff s4,acc090 a1" & _
         ",acc090 a2,trademark,casepropertymap,customer,fagent" & _
         " WHERE e3.eep01=t.eep01 and e3.eep04='" & EMP_發文歸檔 & "'" & _
         " and e3.eep01=c1.cp09(+) and c1.cp27>=20180925 and substr(c1.cp12,1,1)<>'F' and c2.cp09(+)=c1.cp43" & _
         " and e4.eep01=t.eep01 and e4.eep02=t.eep02" & _
         " and s1.st01(+)=c1.cp83 and s2.st01(+)=e4.eep03 and s4.st01(+)=c1.cp13" & _
         " and a1.A0901(+)=s1.ST03 AND a2.A0901(+)=s2.ST03" & _
         " and tm01(+)=c1.cp01 and tm02(+)=c1.cp02 and tm03(+)=c1.cp03 and tm04(+)=c1.cp04" & _
         " and cpm01(+)=c1.cp01 and cpm02(+)=c1.cp10" & _
         " and CU01(+)=substr(TM23,1,8) AND CU02(+)=substr(TM23,9)" & strCon & _
         " and fa01(+)=substr(tm44,1,8) and fa02(+)=substr(tm44,9)"
      '2018/9/18 END
   End If
End If '2020/2/18 + End If
   'stSQL = stSQL & stSQL2 & " order by c1.cp01,c1.cp02,c1.cp03,c1.cp04,c1.cp09"
   stSQL = stSQL & stSQL2 & " order by cp01,cp02,cp03,cp04,cp09"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      With grdDataList
      .Visible = False
      .FixedCols = 0
      Set .Recordset = RsTemp
      m_iRows = RsTemp.RecordCount
      Label1(5).Caption = "共　" & m_iRows & "　件"
      SetGrid
      iColId = GetFieldId("LP11", grdDataList)
      For ii = 1 To .Rows - 1
         '筆數多時速度有點慢,先不抓相關收文號性質
         '.TextMatrix(ii, 3) = .TextMatrix(ii, 3) & PUB_GetRelateCasePropertyName(.TextMatrix(ii, 10), "1")
         '親送要變色
         If .TextMatrix(ii, iColId) = "0" Then
            intI = 1
            For jj = 0 To .FixedCols - 1
               .row = ii
               .col = jj
               .CellBackColor = lblColor(1).BackColor
            Next
         End If
      Next
      .Visible = True
      End With
   Else
      ShowNoData
   End If
End Sub

Private Function GetFieldId(pFieldName As String, ByRef FlexGrid As MSHFlexGrid) As Integer
   Dim iRow As Integer
   With FlexGrid
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetFieldId = iRow
         Exit For
      End If
   Next
   End With
End Function

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer

   arrGridHeadWidth = Array(240, 1200, 650, 1700, 975, 1160, 630, 630, 500, 1500, 1500, 1500)
   iUbound = UBound(arrGridHeadWidth)
   
   With grdDataList
   If pReset = True Then
      .Clear
      .Rows = 2
   End If
   .FixedCols = 5
   'Modified by Morgan 2014/7/2 發文號改到本所案號後以便核對
   'Modified by Morgan 2017/6/5 +直寄
   'Modified by Morgan 2018/10/31 +FC代理人
   .FormatString = "V|本所案號|發文號|案件名稱|案件性質|申請人|部門|人員|直寄|FC代理人|確認/判發時間|發文室發文時間"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .ColAlignment(iCol) = flexAlignLeftCenter
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   End With
End Sub

'Added by Morgan 2015/9/1
'相同接洽人檢查
Private Sub SameContactCheck(pRow As Integer)
   Dim idxCP10 As Integer, idxCont As Integer, idxCP27 As Integer, ii As Integer
   Dim idxCaeNo As Integer, idxDocNo As Integer
   Dim strMsg As String, strCheck As String, bolVisible As Boolean
   Dim idxLP28 As String 'Added by Morgan 2019/4/16
   
   idxCP10 = GetFieldId("cp10", grdDataList)
   With grdDataList
   
   '期限通知1913,通知年費逾期1605
   'Remove by Morgan 2018/10/3 改判斷是否有勾選大宗發文案件
   'If .TextMatrix(pRow, idxCP10) = "1913" Or .TextMatrix(pRow, idxCP10) = "1605" Then
      idxCont = GetFieldId("接洽人", grdDataList)
      idxCaeNo = GetFieldId("本所案號", grdDataList)
      idxDocNo = GetFieldId("發文號", grdDataList)
      idxCP27 = GetFieldId("CP27", grdDataList)
      idxLP28 = GetFieldId("lp28", grdDataList) 'Added by Morgan 2019/4/16
      strMsg = ""
      strCheck = .TextMatrix(pRow, 0)
      For ii = 1 To grdDataList.Rows - 1
         If ii <> pRow And .TextMatrix(ii, 0) = strCheck Then
            'Modified by Morgan 2016/3/3 +判斷同一天發文的
            'Modified by Morgan 2019/4/16 CFP可能不同天發文,改判斷同一天檢核的
            If .TextMatrix(ii, idxCP10) = .TextMatrix(pRow, idxCP10) And .TextMatrix(ii, idxCont) = .TextMatrix(pRow, idxCont) And .TextMatrix(ii, idxLP28) = .TextMatrix(pRow, idxLP28) Then
               strMsg = strMsg & vbCrLf & .TextMatrix(ii, idxCaeNo) & " ( " & .TextMatrix(ii, idxDocNo) & " )"
            End If
         End If
      Next
      If strMsg <> "" Then
         bolVisible = .Visible
         .Visible = True
         If MsgBox("下列信函收件人與本案相同是否一併" & IIf(strCheck = "", "發文", "取消") & "？" & vbCrLf & strMsg, vbQuestion + vbYesNo + vbDefaultButton2, .TextMatrix(pRow, idxCaeNo) & " ( " & .TextMatrix(pRow, idxDocNo) & " )" & "同一信封檢查") = vbYes Then
            For ii = 1 To grdDataList.Rows - 1
               If ii <> pRow And .TextMatrix(ii, 0) = strCheck Then
                  'Modified by Morgan 2016/3/3 +判斷同一天發文的
                  'Modified by Morgan 2024/5/16 CFP可能不同天發文,改判斷同一天檢核的
                  If .TextMatrix(ii, idxCP10) = .TextMatrix(pRow, idxCP10) And .TextMatrix(ii, idxCont) = .TextMatrix(pRow, idxCont) And .TextMatrix(ii, idxLP28) = .TextMatrix(pRow, idxLP28) Then
                     SelectRow ii
                  End If
               End If
            Next
         End If
         .Visible = bolVisible
      End If
   'End If
   End With
End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim nCol As Integer, nRow As Integer
   Dim bolV As Boolean, strItem As String, ii As Integer

   With grdDataList
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow = 0 Then
      .col = nCol
      If m_blnColOrderAsc = False Then '字串降冪
         .Sort = 5 '字串昇冪
         m_blnColOrderAsc = True
      Else
         .Sort = 6 '字串降冪
         m_blnColOrderAsc = False
      End If
   ElseIf nRow > 0 And .TextMatrix(nRow, 1) <> "" Then
      'Modified by Morgan 2018/10/3 改判斷是否有勾選大宗發文案件
      'SameContactCheck nRow 'Added by Morgan 2015/9/1
      If Check1.Value = vbChecked Then SameContactCheck nRow
      'end 2018/10/3
      
      SelectRow nRow
      '控制按鈕為預設值
      bolV = False
      strItem = ""
      For ii = 1 To grdDataList.Rows - 1
         If .TextMatrix(ii, 0) = "V" Then
            bolV = True
            If Trim(GetValue(ii, "發文室發文時間")) = "" Then
               strItem = "1" '發文
            Else
               strItem = "3" '取消發文
            End If
            Exit For
         End If
      Next
      If bolV = False Then
         cmdOK(0).SetFocus
      Else
         If strItem = "1" Then
            cmdOK(1).SetFocus
         ElseIf strItem = "3" Then
            cmdOK(3).SetFocus
         End If
      End If
   
   End If
   .col = nCol
   .Visible = True
   End With
End Sub

Private Sub SelectRow(pRow As Integer)
   Dim iCol As Integer
   With grdDataList
   .row = pRow
   If .TextMatrix(.row, 0) = "V" Then
      .TextMatrix(.row, 0) = ""
      For iCol = .FixedCols To .Cols - 1
        .col = iCol
        .CellBackColor = .BackColor
      Next
   Else
      .TextMatrix(.row, 0) = "V"
      For iCol = .FixedCols To .Cols - 1
        .col = iCol
        .CellBackColor = &HFFC0C0
      Next
   End If
   End With
End Sub

Private Sub Option1_Click(Index As Integer)
If Index = 0 Then
   txtType = "1"
Else
   txtType = "2"
End If
End Sub

Private Sub txtCode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Option1(1).Value = False Then Option1(1).Value = True
End Sub

Private Sub txtCP27_GotFocus(Index As Integer)
   TextInverse txtCP27(Index)
   CloseIme
End Sub

Private Sub txtCP27_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtCP27_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Option1(0).Value = False Then Option1(0).Value = True
End Sub

Private Sub txtCP27_Validate(Index As Integer, Cancel As Boolean)
   If txtCP27(Index) <> "" Then
      If CheckIsTaiwanDate(txtCP27(Index), False) = False Then
         MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
         txtCP27_GotFocus Index
         Cancel = True
      End If
   End If
End Sub

Private Sub txtSystem_GotFocus()
   TextInverse txtSystem
   CloseIme
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSystem_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Option1(1).Value = False Then Option1(1).Value = True
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   If txtType = "" Then
      MsgBox "請輸入類別！", vbInformation
      txtType.SetFocus
      Exit Function
   End If
   If Option1(1).Value = True Then
      If txtSystem = "" Then
         MsgBox "請輸入系統類別！"
         txtSystem.SetFocus
         Exit Function
      End If
      If txtCode(0) = "" Then
         MsgBox "請輸入本所案號！"
         txtCode(0).SetFocus
         Exit Function
      Else
         txtCode(0) = Right(String(6, "0") & txtCode(0), 6)
      End If
      txtCode(1) = Right("0" & txtCode(1), 1)
      txtCode(2) = Right("00" & txtCode(2), 2)
   '所有資料必須輸入發文日區間
   ElseIf txtType = "2" Then
      If txtCP27(0) = "" Then
         MsgBox "請輸入發文日期(起)！"
         txtCP27(0).SetFocus
         Exit Function
      Else
         txtCP27_Validate 0, bCancel
         If bCancel = True Then Exit Function
         
         txtCP27_Validate 1, bCancel
         If bCancel = True Then Exit Function
      End If
   End If
   TxtValidate = True
End Function

Private Sub txtType_GotFocus()
   TextInverse txtType
   CloseIme
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") Then
      KeyAscii = 0
      Beep
   End If
End Sub
