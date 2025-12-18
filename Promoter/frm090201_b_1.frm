VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_b_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員補充資料記錄作業"
   ClientHeight    =   5892
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5892
   ScaleWidth      =   8952
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   375
      Left            =   180
      TabIndex        =   29
      Top             =   2700
      Width           =   5265
      Begin VB.TextBox txtTCD06 
         Height          =   285
         Left            =   1290
         MaxLength       =   7
         TabIndex        =   0
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "補充資料日期："
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   31
         Top             =   30
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "(請輸入日期再按加入記錄按鈕)"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   12
         Left            =   2610
         TabIndex        =   30
         Top             =   60
         Width           =   2460
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Height          =   1185
      Left            =   180
      TabIndex        =   32
      Top             =   2310
      Width           =   8745
      Begin VB.Frame Frame3 
         Height          =   825
         Left            =   4740
         TabIndex        =   33
         Top             =   -60
         Width           =   4005
         Begin VB.CommandButton cmdAdd 
            Caption         =   "<- 加入"
            Height          =   285
            Index           =   1
            Left            =   45
            TabIndex        =   2
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "移除 ->"
            Height          =   285
            Index           =   1
            Left            =   45
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   420
            Width           =   735
         End
         Begin MSForms.TextBox txtText 
            Height          =   300
            Index           =   1
            Left            =   840
            TabIndex        =   1
            Top             =   420
            Width           =   3135
            VariousPropertyBits=   671105051
            Size            =   "5530;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            Caption         =   "逐筆輸入補充資料的內容："
            Height          =   225
            Index           =   7
            Left            =   870
            TabIndex        =   34
            Top             =   180
            Width           =   2760
         End
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Left            =   5760
         TabIndex        =   5
         Top             =   840
         Width           =   2925
         VariousPropertyBits=   671105051
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTCD08 
         Height          =   300
         Left            =   60
         TabIndex        =   41
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   671105055
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstText 
         Height          =   1140
         Index           =   1
         Left            =   1290
         TabIndex        =   4
         Top             =   0
         Width           =   3405
         VariousPropertyBits=   746586139
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "6006;2011"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail備註："
         Height          =   255
         Index           =   13
         Left            =   4740
         TabIndex        =   36
         Top             =   900
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "通知補充資料："
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   35
         Top             =   30
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "加入記錄(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   6480
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   30
      Width           =   1170
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   7710
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   30
      Width           =   1170
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   1995
      Left            =   60
      TabIndex        =   8
      Top             =   3870
      Width           =   8835
      _ExtentX        =   15579
      _ExtentY        =   3514
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label LabCP14 
      Height          =   300
      Left            =   1140
      TabIndex        =   40
      Top             =   1260
      Width           =   2220
      BackColor       =   -2147483644
      VariousPropertyBits=   27
      Size            =   "3916;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LabCP10 
      Height          =   300
      Left            =   1140
      TabIndex        =   39
      Top             =   957
      Width           =   2220
      BackColor       =   -2147483644
      VariousPropertyBits=   27
      Size            =   "3916;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LabTM05 
      Height          =   300
      Left            =   1140
      TabIndex        =   38
      Top             =   600
      Width           =   7440
      BackColor       =   -2147483644
      VariousPropertyBits=   27
      Size            =   "13123;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "備註：在資料列上快速點二下即可查看明細資料"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   14
      Left            =   4470
      TabIndex        =   37
      Top             =   3600
      Width           =   3780
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "齊備日："
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   28
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label LabEP06 
      Height          =   255
      Left            =   1140
      TabIndex        =   27
      Top             =   1920
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "是否急件："
      Height          =   255
      Index           =   3
      Left            =   3510
      TabIndex        =   26
      Top             =   1920
      Width           =   930
   End
   Begin VB.Label LabCP122 
      Height          =   255
      Left            =   4470
      TabIndex        =   25
      Top             =   1920
      Width           =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員補充資料記錄："
      Height          =   225
      Index           =   9
      Left            =   90
      TabIndex        =   24
      Top             =   3600
      Width           =   2010
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   23
      Top             =   270
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   11
      Left            =   210
      TabIndex        =   22
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件性質："
      Height          =   255
      Index           =   2
      Left            =   210
      TabIndex        =   21
      Top             =   957
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "收文日："
      Height          =   255
      Index           =   4
      Left            =   3510
      TabIndex        =   20
      Top             =   957
      Width           =   930
   End
   Begin VB.Label LabCP05 
      Height          =   255
      Left            =   4470
      TabIndex        =   19
      Top             =   957
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "總收文號："
      Height          =   255
      Index           =   6
      Left            =   3510
      TabIndex        =   18
      Top             =   270
      Width           =   930
   End
   Begin VB.Label LabCP09 
      Height          =   255
      Left            =   4470
      TabIndex        =   17
      Top             =   270
      Width           =   1740
   End
   Begin VB.Label LabID 
      Height          =   255
      Left            =   1140
      TabIndex        =   16
      Top             =   270
      Width           =   1740
   End
   Begin VB.Label LabCP48 
      Height          =   255
      Left            =   4470
      TabIndex        =   15
      Top             =   1260
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "承辦期限："
      Height          =   255
      Index           =   10
      Left            =   3510
      TabIndex        =   14
      Top             =   1260
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "承辦人："
      Height          =   255
      Index           =   17
      Left            =   210
      TabIndex        =   13
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label LabCP07 
      Height          =   255
      Left            =   4470
      TabIndex        =   12
      Top             =   1590
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "法定期限："
      Height          =   255
      Index           =   19
      Left            =   3510
      TabIndex        =   11
      Top             =   1590
      Width           =   930
   End
   Begin VB.Label LabCP06 
      Height          =   255
      Left            =   1140
      TabIndex        =   10
      Top             =   1590
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "本所期限："
      Height          =   255
      Index           =   21
      Left            =   210
      TabIndex        =   9
      Top             =   1590
      Width           =   900
   End
End
Attribute VB_Name = "frm090201_b_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/23 改成Form2.0 ; grd1改字型=新細明體-ExtB、LabTM05、LabCP10、LabCP14、lstText(index)、txtText(index)、Text1、textTCD08
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create By Sindy 2012/5/7
Option Explicit

'紀錄作用按鍵
Public cmdState As Integer
Dim m_CP13 As String 'Add By Sindy 2012/10/24
Dim m_CP10 As String 'Added by Lydia 2019/06/03
Dim strCase(1 To 4) As String 'Added by Lydia 2022/07/15 記錄本所案號
 
Private Sub SetDataListWidth()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

arrGridHeadText = Array("操作日期", "操作人員", "操作時間", "輸入日期", "操作動作", "通知補充資料")
arrGridHeadWidth = Array(850, 850, 850, 850, 1300, 4000)
grd1.MergeCells = flexMergeRestrictColumns
grd1.Cols = UBound(arrGridHeadText) + 1
For iRow = 0 To grd1.Cols - 1
   grd1.row = 0
   grd1.col = iRow
   grd1.Text = arrGridHeadText(iRow)
   grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
   grd1.CellAlignment = flexAlignLeftCenter
Next
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim Cancel As Boolean
Dim strTemp As Variant, strText As String, strSubject As String, strContent As String
Dim i As Integer
Dim strCaseNo As String

On Error GoTo ErrHnd

cmdState = Index
Select Case cmdState
Case 1
   '智權人員補充資料
   If Me.Frame1.Visible = True Then
      If txtTCD06 <> "" Then
         Cancel = False
         Call txtTCD06_Validate(Cancel)
         If Cancel = True Then
            txtTCD06.SetFocus
            Exit Sub
         End If
         
         '檢查補充資料日期不可大於系統日
         If Val(DBDATE(txtTCD06)) > strSrvDate(1) Then
            MsgBox "補充資料日期不可大於系統日！", vbExclamation
            txtTCD06.SetFocus
            Exit Sub
         End If
         
         strSql = "SELECT count(*) FROM tmctldate WHERE tcd01='" & Trim(LabCP09.Caption) & "' " & _
                  "AND tcd02='2' " & _
                  "AND tcd07='智權人員補充資料' " & _
                  "AND tcd06=" & DBDATE(txtTCD06)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               MsgBox "補充資料日期重覆！", vbExclamation
               txtTCD06.SetFocus
               Exit Sub
            End If
         End If
         
         cnnConnection.BeginTrans
         strSql = "insert into tmctldate(tcd01,tcd02,tcd03,tcd04,tcd05,tcd06,tcd07)" & _
                  " values('" & Trim(LabCP09.Caption) & "','2','" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & "," & DBDATE(txtTCD06) & ",'智權人員補充資料')"
         cnnConnection.Execute strSql
         cnnConnection.CommitTrans
         
         Me.txtTCD06 = ""
         Call ReadGrid
         MsgBox "存檔完成！", vbExclamation
         'Me.Hide
         Exit Sub
      Else
         MsgBox "請輸入資料！", vbExclamation
         txtTCD06.SetFocus
         Exit Sub
      End If
      
   'Add By Sindy 2012/10/24
   '通知智權人員補充資料
   ElseIf Me.Frame2.Visible = True Then
      If lstText(1).ListCount > 0 Then
         cnnConnection.BeginTrans
         '拿掉齊備日
         strSql = "update engineerprogress set ep06=0,ep36=0 where ep02='" & Trim(LabCP09.Caption) & "'"
         cnnConnection.Execute strSql
         '拿掉承辦期限
         'Modified by Lydia 2022/07/15 TC案之文件齊備日管控
         'If m_CP10 <> "102" Then 'Added by Lydia 2019/06/03 排除延展案
         If ((strCase(1) = "T" Or strCase(1) = "FCT") And m_CP10 <> "102") Or strCase(1) = "TC" Then
             strSql = "update CaseProgress SET CP48=null WHERE CP09='" & Trim(LabCP09.Caption) & "'"
             cnnConnection.Execute strSql
         End If 'end 2019/06/03
         '新增一筆通知補充資料記錄
         strSql = "insert into tmctldate(tcd01,tcd02,tcd03,tcd04,tcd05,tcd07,tcd08)" & _
                  " values('" & Trim(LabCP09.Caption) & "','5','" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & ",'通知補充資料','" & textTCD08.Text & "')"
         cnnConnection.Execute strSql
         cnnConnection.CommitTrans
         
         '寄信通知智權人員
         strTemp = Split(textTCD08, ",")
         For i = 0 To UBound(strTemp)
            If i > 0 Then
               strText = strText + "　　　　　　"
            End If
            strText = strText + strTemp(i) + vbCrLf
         Next i
         If Right(LabID, 5) = "-0-00" Then
            strCaseNo = Left(LabID, Len(LabID) - 5)
         Else
            strCaseNo = LabID
         End If
         'Modified by Lydia 2018/12/10 台灣商標爭議案=>台灣商標案
         'Modified by Lydia 2022/07/15 台灣商標案=>商標著作權案
         strSubject = strCaseNo & " 商標著作權案 (通知補充資料)"
         strContent = "本所案號：" + LabID + vbCrLf + _
                      "案件名稱：" + LabTM05 + vbCrLf + _
                      "案件性質：" + LabCP10 + vbCrLf + _
                      "收文日　：" + LabCP05 + vbCrLf + _
                      "本所期限：" + LabCP06 + vbCrLf + _
                      "法定期限：" + LabCP07 + vbCrLf + _
                      "是否急件：" + LabCP122 + vbCrLf + vbCrLf + _
                      "請補充資料：" + strText + vbCrLf
         If Text1 <> "" Then
            strContent = strContent + "備　　註：" + Text1
         End If
         'Modify By Sindy 2012/11/19 給智權人員的通知補充資料郵件,若原已有齊備日及承辦期限,於郵件最下方加註PS,若原無齊備日及承辦期限者則不必加註
         If Trim(LabCP48) <> "" Or Trim(LabEP06) <> "" Then
            strContent = strContent + vbCrLf + vbCrLf
            'Modified by Lydia 2019/06/03 排除延展案
            'strContent = strContent + "ＰＳ：此郵件同時取消原齊備日(" & Trim(LabEP06) & ")及承辦期限(" & Trim(LabCP48) & ")" + vbCrLf
            strContent = strContent + "ＰＳ：此郵件同時取消原齊備日(" & Trim(LabEP06) & ")"
            'Modified by Lydia 2022/07/15 TC案之文件齊備日管控
            'If m_CP10 <> "102" Then
            If ((strCase(1) = "T" Or strCase(1) = "FCT") And m_CP10 <> "102") Or strCase(1) = "TC" Then
                strContent = strContent + "及承辦期限(" & Trim(LabCP48) & ")" + vbCrLf
            End If
            'end 2019/06/03
         End If
         '2012/11/19 End
         PUB_SendMail strUserNum, m_CP13, "", strSubject, strContent, ""
         
         Me.textTCD08 = ""
         Me.lstText(1).Clear
         Me.Text1 = ""
         
         Call ReadGrid
         MsgBox "存檔完成！", vbExclamation
         Me.Hide
         Exit Sub
      Else
         MsgBox "請輸入資料！", vbExclamation
         txtText(1).SetFocus
         Exit Sub
      End If
   '2012/10/24 End
   End If
Case 0
   Me.Hide
   Exit Sub
Case Else
End Select

ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Sub Form_Load()
bolToEndByNick = False
MoveFormToCenter Me
cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090201_b_1 = Nothing
End Sub

Public Function Process(strText As String) As Boolean
On Error GoTo ErrHnd
   
   Process = False
   
   'Modify By Sindy 2012/10/24 +CP13
   'Modified by Lydia 2018/12/10 開放T台灣案管控文件齊備
   'strSql = "select sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,cpm03 as 案件性質,tm05 as 案件名稱," & _
            "s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(ep06) as 齊備日," & _
            "sqldatet(cp48) as 承辦期限,sqldatet(ep28) as 指定會稿日,ep34 as 是否會稿,sqldatet(ep07) as 會稿日,sqldatet(cp27) as 發文日," & _
            "cp16 As 費用, cp18 As 點數, cp64 As 進度備註, cp09 As 總收文號,cp122,cp13" & _
            " from caseprogress,engineerprogress,trademark,casepropertymap,staff s1,staff s2" & _
            " where cp01 in('T','FCT') and cp10 in (" & TMdebate & ")" & _
            " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
            " and cp09=ep02(+)" & _
            " and tm10='000'" & _
            " and cp01=cpm01(+) and cp10=cpm02(+)" & _
            " and cp13=s1.st01(+)" & _
            " and cp14=s2.st01(+)" & _
            " and cp09='" & strText & "'"
   'Modified by Lydia 2019/06/03 +CP10
   'Modified by Lydia 2022/07/15 T大陸案之齊備日管控: tm10='000' => tm10 in ('000','020') 、cpm03 => decode(tm10,'000',cpm03,cpm04)
   strSql = "select sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,decode(tm10,'000',cpm03,cpm04) as 案件性質,tm05 as 案件名稱," & _
            "s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(ep06) as 齊備日," & _
            "sqldatet(cp48) as 承辦期限,sqldatet(ep28) as 指定會稿日,ep34 as 是否會稿,sqldatet(ep07) as 會稿日,sqldatet(cp27) as 發文日," & _
            "cp16 As 費用, cp18 As 點數, cp64 As 進度備註, cp09 As 總收文號,cp122,cp13,cp10" & _
            " from caseprogress,engineerprogress,trademark,casepropertymap,staff s1,staff s2" & _
            " where cp01 in('T','FCT') " & _
            " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
            " and cp09=ep02(+)" & _
            " and tm10 in ('000','020')" & _
            " and cp01=cpm01(+) and cp10=cpm02(+)" & _
            " and cp13=s1.st01(+)" & _
            " and cp14=s2.st01(+)" & _
            " and cp09='" & strText & "'"
   'Added by Lydia 2022/07/15 TC案之文件齊備日管控: 臺灣、大陸
   strSql = strSql & " Union select sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,decode(sp09,'000',cpm03,cpm04) as 案件性質,sp05 as 案件名稱," & _
            "s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(ep06) as 齊備日," & _
            "sqldatet(cp48) as 承辦期限,sqldatet(ep28) as 指定會稿日,ep34 as 是否會稿,sqldatet(ep07) as 會稿日,sqldatet(cp27) as 發文日," & _
            "cp16 As 費用, cp18 As 點數, cp64 As 進度備註, cp09 As 總收文號,cp122,cp13,cp10" & _
            " from caseprogress,engineerprogress,servicepractice,casepropertymap,staff s1,staff s2" & _
            " where cp01 in('TC') " & _
            " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)" & _
            " and cp09=ep02(+)" & _
            " and sp09 in ('000','020')" & _
            " and cp01=cpm01(+) and cp10=cpm02(+)" & _
            " and cp13=s1.st01(+)" & _
            " and cp14=s2.st01(+)" & _
            " and cp09='" & strText & "'"
   'end 2022/07/15
   CheckOC3
   txtTCD06.Text = ""
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Process = True
         LabID.Caption = "" & .Fields("本所案號")
         LabCP09.Caption = "" & .Fields("總收文號")
         LabTM05.Caption = "" & .Fields("案件名稱")
         LabCP10.Caption = "" & .Fields("案件性質")
         LabCP05.Caption = "" & .Fields("收文日")
         LabCP14.Caption = "" & .Fields("承辦人")
         LabCP48.Caption = "" & .Fields("承辦期限")
         LabCP06.Caption = "" & .Fields("本所期限")
         LabCP07.Caption = "" & .Fields("法定期限")
         LabEP06.Caption = "" & .Fields("齊備日")
         LabCP122.Caption = "" & .Fields("cp122")
         m_CP13 = "" & .Fields("cp13") 'Add By Sindy 2012/10/24
         m_CP10 = "" & .Fields("cp10") 'Added by Lydia 2019/06/03
         'Added by Lydia 2022/07/15
         strExc(1) = Replace("" & .Fields("本所案號"), "-", "")
         Call ChgCaseNo(strExc(1), strCase)
         'end 2022/07/15
      Else
         LabID.Caption = ""
         LabCP09.Caption = ""
         LabTM05.Caption = ""
         LabCP10.Caption = ""
         LabCP05.Caption = ""
         LabCP14.Caption = ""
         LabCP48.Caption = ""
         LabCP06.Caption = ""
         LabCP07.Caption = ""
         LabEP06.Caption = ""
         LabCP122.Caption = ""
         m_CP13 = "" 'Add By Sindy 2012/10/24
         m_CP10 = "" 'Added by Lydia 2019/06/03
         MsgBox "查無資料！", vbExclamation
         Me.Hide
         Exit Function
      End If
   End With
   Call ReadGrid
   Exit Function
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub ReadGrid()
   '智權人員補充資料記錄：以總收文號抓的商標案管制日期異動記錄檔異動類別=2的資料，依異動日期＋異動時間排序顯示
   'Modify By Sindy 2012/10/24 增加TCD02=5通知補充資料
   strSql = "select sqldatet(tcd04) as 操作日期,st02 as 操作人員,sqltime(tcd05) as 操作時間,sqldatet(tcd06) as 輸入日期,tcd07 as 操作動作,tcd08 as 通知補充資料" & _
            " From tmctldate,staff" & _
            " where tcd01='" & LabCP09 & "'" & _
            " and tcd02 in('2','5')" & _
            " and tcd03=st01(+)" & _
            " order by tcd04,tcd05 asc"
   CheckOC3
   grd1.Rows = 2
   grd1.Clear
   SetDataListWidth
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Set grd1.Recordset = AdoRecordSet3.Clone
         SetDataListWidth
      End If
   End With
End Sub

'Add By Sindy 2012/12/18
Private Sub GRD1_DblClick()
Dim i As Integer
   
   grd1.col = 0
   grd1.row = grd1.MouseRow
   grd1.Visible = False
   If grd1.row <> 0 Then
      If DBDATE(grd1.TextMatrix(grd1.row, 0)) <> "" Then
         If frm090201_b_2.StrMenu(Trim(LabCP09.Caption), DBDATE(grd1.TextMatrix(grd1.row, 0)), Format(grd1.TextMatrix(grd1.row, 2), "HHMMSS")) = True Then
            frm090201_b_2.Show vbModal
         Else
            grd1.Visible = True
            ShowNoData
            Exit Sub
         End If
      End If
   End If
   grd1.Visible = True
End Sub

'Add By Sindy 2012/12/18
Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
Dim i As Integer
   
   getGrdColRow grd1, x, y, nCol, nRow
   grd1.col = nCol
   grd1.row = nRow
End Sub

Private Sub txtTCD06_GotFocus()
   InverseTextBox txtTCD06
End Sub

Private Sub txtTCD06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(txtTCD06) = False Then
      If CheckIsTaiwanDate(txtTCD06, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "補充資料日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTCD06_GotFocus
      End If
      'Modified by Lydia 2019/05/02 輸入非工作日，改成往後推工作日; ex.T-217900齊備日輸入5/1(勞動節放假)
      'If ChkWork(ChangeTStringToWString(txtTCD06)) = False Then
      '   Cancel = True
      '   txtTCD06_GotFocus
      'End If
      strExc(1) = CompWorkDay(1, DBDATE(txtTCD06))
      If strExc(1) <> DBDATE(txtTCD06) Then
          MsgBox "輸入之日期不是工作天!!" & vbCrLf & "請輸入" & TransDate(strExc(1), 1), , "日期錯誤"
          txtTCD06.SetFocus
          txtTCD06_GotFocus
          Cancel = True
      End If
      'end 2019/05/02
   End If
End Sub

Private Sub txtTCD06_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

'加入資料
Private Sub cmdAdd_Click(Index As Integer)
   'Add By Sindy 2012/11/30 資料內容不可輸入,符號
   If Trim(txtText(Index)) <> "" Then
      txtText(Index) = Replace(Trim(txtText(Index)), ",", "")
      'Added by Lydia 2021/12/23 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me, , True, "TextBox") = False Then
          Exit Sub
      End If
      'end 2021/12/23
   End If
   '2012/11/30 End
   If Addlst(Index) = True Then
      Select Case Index
         Case 1
            If Trim(textTCD08) <> "" Then textTCD08 = Trim(textTCD08) & ","
            textTCD08 = Trim(textTCD08) & Trim(txtText(Index))
            txtText(Index) = ""
      End Select
   End If
   txtText(Index).SetFocus
End Sub

'移除資料
Private Sub cmdRemove_Click(Index As Integer)
   Removelst Index
   Select Case Index
      Case 1
         textTCD08 = ComposeListX(Index)
   End Select
   txtText(Index).SetFocus
End Sub

Private Function Addlst(p_idx As Integer) As Boolean
   Dim idx As Integer, bFound As Boolean
   If txtText(p_idx) <> "" Then
      For idx = 0 To lstText(p_idx).ListCount - 1
         If Trim(txtText(p_idx).Text) = Trim(lstText(p_idx).List(idx)) Then
            MsgBox "補充資料已存在！"
            txtText(p_idx).SetFocus
            txtText_GotFocus p_idx
            bFound = True
            Addlst = False
            Exit For
         End If
      Next idx
      If bFound = False Then
         Addlst = True
         lstText(p_idx).AddItem Trim(txtText(p_idx)), 0
      End If
   End If

End Function

Private Sub Removelst(p_idx As Integer)
   Dim idx As Integer, ii As Integer
   If lstText(p_idx).ListCount > 0 Then
      ii = 0
      For idx = 0 To lstText(p_idx).ListCount - 1
         'Modified by Lydia 2021/12/23
'         If lstText(p_idx).Selected(ii) = True Then
'            lstText(p_idx).RemoveItem ii
'            ii = ii - 1
'         End If
'         ii = ii + 1
         If ii >= 0 Then
             If lstText(p_idx).Selected(ii) = True Then
                lstText(p_idx).RemoveItem ii
                ii = ii - 1
             Else
                ii = ii + 1
             End If
         End If
         'end 2021/12/23
      Next
   End If
End Sub

Private Function ComposeListX(p_index As Integer) As String
   strExc(1) = ""
   If lstText(p_index).ListCount > 0 Then
      strExc(1) = lstText(p_index).List(0)
      For intI = 1 To lstText(p_index).ListCount - 1
         strExc(1) = strExc(1) & "," & lstText(p_index).List(intI)
      Next
   End If
   ComposeListX = strExc(1)
End Function

Private Sub txtText_GotFocus(Index As Integer)
   TextInverse txtText(Index)
End Sub
