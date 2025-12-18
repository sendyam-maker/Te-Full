VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040152 
   BorderStyle     =   1  '單線固定
   Caption         =   "電子收文接洽單查詢"
   ClientHeight    =   6670
   ClientLeft      =   3780
   ClientTop       =   3700
   ClientWidth     =   8960
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6670
   ScaleWidth      =   8960
   Begin VB.Frame Frame1 
      Caption         =   "電腦中心使用"
      ForeColor       =   &H000000FF&
      Height          =   840
      Left            =   36
      TabIndex        =   28
      Top             =   5796
      Width           =   8868
      Begin VB.TextBox txtCRL65 
         Height          =   264
         Left            =   1428
         MaxLength       =   10
         TabIndex        =   30
         Top             =   216
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "關連表單編號："
         Height          =   252
         Index           =   8
         Left            =   108
         TabIndex        =   31
         Top             =   252
         Width           =   1308
      End
      Begin VB.Label Label1 
         Caption         =   "備註：剔除法律所案源預存的接洽單，電腦中心人員除外。"
         ForeColor       =   &H00FF0000&
         Height          =   228
         Index           =   5
         Left            =   3456
         TabIndex        =   29
         Top             =   36
         Width           =   4908
      End
   End
   Begin VB.TextBox txtCRL55 
      Height          =   264
      Left            =   4380
      TabIndex        =   4
      Top             =   396
      Width           =   1095
   End
   Begin VB.TextBox txtCRL01 
      Height          =   264
      Left            =   4380
      MaxLength       =   10
      TabIndex        =   2
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "基本資料(&B)"
      Height          =   345
      Index           =   4
      Left            =   6540
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   510
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   345
      Index           =   5
      Left            =   7635
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   510
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "是否列印特殊收據頁"
      Height          =   255
      Left            =   3450
      TabIndex        =   10
      Top             =   1056
      Width           =   1995
   End
   Begin VB.TextBox txtCRL02 
      Height          =   264
      Index           =   1
      Left            =   2175
      MaxLength       =   7
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.TextBox txtPrintType 
      Height          =   264
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   9
      Text            =   "1"
      Top             =   1032
      Width           =   240
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "刪除(&D)"
      Height          =   345
      Index           =   3
      Left            =   5610
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   510
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox txtSystem 
      Height          =   300
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   5
      Top             =   732
      Width           =   465
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   0
      Left            =   1710
      MaxLength       =   6
      TabIndex        =   6
      Top             =   732
      Width           =   765
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   1
      Left            =   2550
      MaxLength       =   1
      TabIndex        =   7
      Top             =   732
      Width           =   225
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   2
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   8
      Top             =   732
      Width           =   345
   End
   Begin VB.TextBox txtSales 
      Height          =   264
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   3
      Top             =   396
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Left            =   8340
      Top             =   900
   End
   Begin VB.TextBox txtCRL02 
      Height          =   264
      Index           =   0
      Left            =   1170
      MaxLength       =   7
      TabIndex        =   0
      Top             =   60
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   4170
      Left            =   30
      TabIndex        =   18
      Top             =   1335
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   7338
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
      _Band(0).Cols   =   16
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "查詢(&Q)"
      Height          =   345
      Index           =   0
      Left            =   5610
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   2
      Left            =   7980
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   60
      Width           =   900
   End
   Begin VB.TextBox txtPCnt 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Left            =   7290
      MaxLength       =   1
      TabIndex        =   13
      Text            =   "2"
      Top             =   90
      Width           =   270
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印(&P)　　份"
      Height          =   345
      Index           =   1
      Left            =   6540
      TabIndex        =   12
      Top             =   60
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案源單號："
      Height          =   255
      Index           =   7
      Left            =   3240
      TabIndex        =   27
      Top             =   420
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "接洽單編號："
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   26
      Top             =   90
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "電子收文啟用日="
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   1
      Left            =   420
      TabIndex        =   25
      Top             =   5550
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSForms.TextBox lblSalesName 
      Height          =   300
      Left            =   2055
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   390
      Width           =   1125
      VariousPropertyBits=   671105055
      Size            =   "1984;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line2 
      X1              =   1950
      X2              =   2550
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Label Label1 
      Caption         =   "輸出方式： 　   ( 1.螢幕 2.印表機 )"
      Height          =   180
      Index           =   122
      Left            =   240
      TabIndex        =   23
      Top             =   1110
      Width           =   3090
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "本所案號："
      Height          =   255
      Index           =   6
      Left            =   -30
      TabIndex        =   22
      Top             =   735
      Width           =   1155
   End
   Begin VB.Line Line1 
      X1              =   1470
      X2              =   3180
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Label Label1 
      Caption         =   "共　0　件"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   3
      Left            =   4110
      TabIndex        =   21
      Top             =   750
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "填單日期："
      Height          =   255
      Index           =   0
      Left            =   -30
      TabIndex        =   20
      Top             =   90
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "員工編號："
      Height          =   255
      Index           =   2
      Left            =   -30
      TabIndex        =   19
      Top             =   390
      Width           =   1155
   End
End
Attribute VB_Name = "frm12040152"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/13 Form2.0已修改(lblSalesName,GrdDataList改Font)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
'Memo by Lydia 2019/07/01 表單名稱: 接洽記錄單查詢及列印=>接洽記錄單查詢／列印
'Memo by Lydia 2021/05/18 表單名稱: 接洽記錄單查詢／列印=>自動收文接洽單查詢/列印
'Modify By Sindy 2023/1/6 更名為「電子收文接洽單查詢」
Option Explicit

Dim i As Integer, j As Integer, s As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_intRow As Integer, m_intCol As Integer
Public cmdState As Integer
'Add by Morgan 2011/1/11
Dim m_strAuthorityCon As String '權限控制條件
Dim m_stST05 As String
Dim m_PrevRow As Double 'Add By Sindy 2022/11/22
Dim bolSpecMan As Boolean, strSpecCode As String 'Add By Sindy 2023/1/16


Private Sub SetDataListWidth()
   GrdDataList.row = 0
   GrdDataList.col = 0: GrdDataList.Text = "V"
   GrdDataList.ColWidth(0) = 200
   GrdDataList.CellAlignment = flexAlignLeftCenter
   'Add By Sindy 2022/10/19
   If (Val(Trim(txtCRL02(0).Text)) = 0 And strSrvDate(1) >= 接洽單電子收文啟用日) Or _
      DBDATE(Val(Trim(txtCRL02(0).Text))) >= 接洽單電子收文啟用日 Then
      
      GrdDataList.col = 1: GrdDataList.Text = "狀態"
      GrdDataList.ColWidth(1) = 420
      GrdDataList.CellAlignment = flexAlignLeftCenter
   Else
      GrdDataList.col = 1: GrdDataList.Text = "狀態"
      GrdDataList.ColWidth(1) = 0
      GrdDataList.CellAlignment = flexAlignLeftCenter
   End If
   '2022/10/19 END
   GrdDataList.col = 2: GrdDataList.Text = "接洽單編號"
   GrdDataList.ColWidth(2) = 1000
   GrdDataList.CellAlignment = flexAlignLeftCenter
   GrdDataList.col = 3: GrdDataList.Text = "填單日期"
   GrdDataList.ColWidth(3) = 850
   GrdDataList.CellAlignment = flexAlignLeftCenter
   GrdDataList.col = 4: GrdDataList.Text = "員工編號"
   GrdDataList.ColWidth(4) = 1000
   GrdDataList.CellAlignment = flexAlignLeftCenter
   GrdDataList.col = 5: GrdDataList.Text = "種類"
   GrdDataList.ColWidth(5) = 600
   GrdDataList.CellAlignment = flexAlignLeftCenter
   GrdDataList.col = 6: GrdDataList.Text = "本所案號"
   GrdDataList.ColWidth(6) = 1200
   GrdDataList.CellAlignment = flexAlignLeftCenter
   GrdDataList.col = 7: GrdDataList.Text = "主題"
   GrdDataList.ColWidth(7) = 2000
   GrdDataList.CellAlignment = flexAlignLeftCenter
   'Modify By Sindy 2022/9/5
   If (Val(Trim(txtCRL02(0).Text)) = 0 And strSrvDate(1) >= 接洽單電子收文啟用日) Or _
      DBDATE(Val(Trim(txtCRL02(0).Text))) >= 接洽單電子收文啟用日 Then
      GrdDataList.col = 8: GrdDataList.Text = "案件性質"
      GrdDataList.ColWidth(8) = 3600
      GrdDataList.CellAlignment = flexAlignLeftCenter
      GrdDataList.col = 9: GrdDataList.Text = "案件性質2"
      GrdDataList.ColWidth(9) = 0
      GrdDataList.CellAlignment = flexAlignLeftCenter
      GrdDataList.col = 10: GrdDataList.Text = "案件性質3"
      GrdDataList.ColWidth(10) = 0
      GrdDataList.CellAlignment = flexAlignLeftCenter
      GrdDataList.col = 11: GrdDataList.Text = "案件性質4"
      GrdDataList.ColWidth(11) = 0
      GrdDataList.CellAlignment = flexAlignLeftCenter
      GrdDataList.col = 12: GrdDataList.Text = "CRL08"
      GrdDataList.ColWidth(12) = 0
      GrdDataList.CellAlignment = flexAlignLeftCenter
   Else
   '2022/9/5 END
      GrdDataList.col = 8: GrdDataList.Text = "案件性質1"
      GrdDataList.ColWidth(8) = 900
      GrdDataList.CellAlignment = flexAlignLeftCenter
      GrdDataList.col = 9: GrdDataList.Text = "案件性質2"
      GrdDataList.ColWidth(9) = 900
      GrdDataList.CellAlignment = flexAlignLeftCenter
      GrdDataList.col = 10: GrdDataList.Text = "案件性質3"
      GrdDataList.ColWidth(10) = 900
      GrdDataList.CellAlignment = flexAlignLeftCenter
      GrdDataList.col = 11: GrdDataList.Text = "案件性質4"
      GrdDataList.ColWidth(11) = 900
      GrdDataList.CellAlignment = flexAlignLeftCenter
      GrdDataList.col = 12: GrdDataList.Text = "CRL08"
      GrdDataList.ColWidth(12) = 0
      GrdDataList.CellAlignment = flexAlignLeftCenter
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
   Exit Sub
End Sub

Public Sub PubShowNextData()
Dim bolPrint As Boolean
Dim i As Integer
Dim strCRL19 As String, strF0201 As String
   
   Select Case cmdState
      Case 0 '查詢
         Call SearchData
         
      Case 1 '列印
         bolPrint = False
         'Add By Sindy 2022/12/8
         If DBDATE(Trim(txtCRL02(0).Text)) >= 接洽單電子收文啟用日 Then
            '畫面存在,就先關閉
            If PUB_CheckFormExist("frm090801_Q") = True Then
               Unload frm090801_Q
            End If
         End If
         For i = 1 To GrdDataList.Rows - 1
            If GrdDataList.TextMatrix(i, 0) = "V" Then
               'Add By Sindy 2022/11/22
               '將上一筆把勾拿掉的資料列反白
               If m_PrevRow > 0 And m_PrevRow <= GrdDataList.Rows - 1 Then
                  GrdDataList.row = m_PrevRow
                  GrdDataList.col = 1
                  If GrdDataList.TextMatrix(m_PrevRow, 0) = "" And GrdDataList.CellBackColor = &HFFC0C0 Then
                     For j = 0 To GrdDataList.Cols - 1
                        GrdDataList.col = j
                        GrdDataList.CellBackColor = QBColor(15)
                     Next j
                  End If
               End If
               '2022/11/22 END
               GrdDataList.row = i
               GrdDataList.col = 0
               GrdDataList.Text = "": m_PrevRow = i 'Add By Sindy 2022/11/22
               For j = 0 To 0 'GrdDataList.Cols - 1
                  GrdDataList.col = j
                  GrdDataList.CellBackColor = QBColor(15)
               Next j
               
               'Modify By Sindy 2022/9/5
               '檢查CRL19案件性質是否有資料,若有就跑舊接洽單
               '法律所未電子化,電子化上線一樣維持使用舊接洽單
               strSql = "select crL01,crL19,f0201 from consultrecordlist,flow002 where crl01='" & Trim(GrdDataList.TextMatrix(i, 2)) & "'" & _
                        " and f0201(+)=crl01"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  strCRL19 = "" & RsTemp.Fields("crL19")
                  strF0201 = "" & RsTemp.Fields("f0201")
               End If
'               If strCRL19 <> "" And Pub_StrUserSt03 = "M51" Then
'                  If MsgBox("要看電子接洽單嗎？", vbYesNo + vbCritical + vbDefaultButton1, "詢問") = vbYes Then
'                     strCRL19 = ""
'                  End If
'               End If
               'If DBDATE(Trim(txtCRL02(0).Text)) >= 接洽單電子收文啟用日 And strCRL19 = "" Then
               If DBDATE(Trim(txtCRL02(0).Text)) >= 接洽單電子收文啟用日 And (strCRL19 = "" Or strF0201 <> "") Then
                  '畫面存在,就先關閉
                  If PUB_CheckFormExist("frm090801_Q") = True Then
                     Unload frm090801_Q
                  End If
                  '查詢接洽記錄單
                  frm090801_Q.SetParent Me
                  frm090801_Q.m_blnCallPrint = True
                  frm090801_Q.Text5 = Trim(GrdDataList.TextMatrix(i, 2))
                  Call frm090801_Q.cmdOK_Click(4)
                  'frm090801_Q.ZOrder 1
                  frm090801_Q.Show 'vbModal
                  'Me.Hide
                  Exit Sub
               Else
               '2022/9/5 END
                  bolPrint = True
                  frm090801.txtPCnt = Me.txtPCnt
                  frm090801.txtPrintType = Me.txtPrintType
                  '查詢
                  frm090801.Text5 = Trim(GrdDataList.TextMatrix(i, 2))
                  frm090801.m_blnCallPrint_CRL119 = IIf(Check1.Value = 1, True, False) 'Add By Sindy 2014/2/7 是否列印特殊收據頁
                  Call frm090801.cmdOK_Click(4)
                  '列印
                  frm090801.m_blnCallPrint = True
                  Call frm090801.cmdOK_Click(0)
                  Unload frm090801
               End If
'               frm090801.SetParent Me
'               Me.Hide
'               If txtPrintType = "2" Then '2.印表機
'                  Unload frm090801
'               End If
            End If
         Next i
         If Me.txtPrintType = "2" And bolPrint = True Then
            ShowPrintOk
         End If
         
      Case 2 '結束
         Unload Me
         
      Case 3 '刪除
'         If MsgBox("確定是否要刪除資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
'            cnnConnection.BeginTrans
'            For i = 1 To grdDataList.Rows - 1
'               If grdDataList.TextMatrix(i, 0) = "V" Then
'                  strSql = "delete from consultrecordlist where crl01='" & Trim(grdDataList.TextMatrix(i, 2)) & "' "
'                  cnnConnection.Execute strSql, intI
'                  strSql = "delete from consultrecapp where cra01='" & Trim(grdDataList.TextMatrix(i, 2)) & "' "
'                  cnnConnection.Execute strSql, intI
'                  strSql = "delete from consultrecinv where cri01='" & Trim(grdDataList.TextMatrix(i, 2)) & "' "
'                  cnnConnection.Execute strSql, intI
'
'                  PUB_DelFtpFile2 Trim(grdDataList.TextMatrix(i, 2)), , UCase("ConsultRecImageF") '檔案改放FTP,必須在DB資料刪除前執行
'                  strSql = "delete from consultrecimagef where crif01='" & Trim(grdDataList.TextMatrix(i, 2)) & "' "
'                  cnnConnection.Execute strSql, intI
'               End If
'            Next i
'            cnnConnection.CommitTrans
'            Call SearchData
'         End If
         
      Case 4 '基本資料
         Dim Str01 As String
         Me.Enabled = False
         For i = 1 To GrdDataList.Rows - 1
            GrdDataList.col = 0
            GrdDataList.row = i
            If Trim(GrdDataList.Text) = "V" Then
               GrdDataList.col = 0
               GrdDataList.Text = ""
               For j = 0 To GrdDataList.Cols - 1
                  GrdDataList.col = j
                  GrdDataList.CellBackColor = QBColor(15)
               Next j
               Str01 = SystemNumber(Trim(GrdDataList.TextMatrix(i, 6)), 1)
               If Mid(UCase(Str01), 1, 1) = "N" Then
                   Str01 = Mid(Str01, 2, 3)
               End If
               GrdDataList.col = 12
               If GrdDataList.Text <> "" Then
                  'Modified by Morgan 2016/3/24 排除母層是共同查詢
'                  If UCase(m_PrevForm.Name) <> UCase("frm100101_2") Then
                     fnCloseAllFrm100 'Added by Morgan 2016/2/22
'                  End If
                  'end 2016/3/24
                  If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                  End If
                  Select Case Pub_RplStr(Str01)
                      Case "CFP", "FCP", "P"   '專利
                            Screen.MousePointer = vbHourglass
                            frm100101_3.Show
                            frm100101_3.Tag = Pub_RplStr(Trim(GrdDataList.TextMatrix(i, 6)))
                            frm100101_3.StrMenu
                            Screen.MousePointer = vbDefault
                      Case "CFT", "FCT", "T", "TF"   '商標
                            Screen.MousePointer = vbHourglass
                            frm100101_4.Show
                            frm100101_4.Tag = Pub_RplStr(Trim(GrdDataList.TextMatrix(i, 6)))
                            frm100101_4.StrMenu
                            Screen.MousePointer = vbDefault
                      'Modify By Sindy 2009/07/24 增加LIN系統類別
                      'modify by sonia 2019/7/29 +ACS系統類別
                      Case "CFL", "FCL", "L", "LIN", "ACS"   '法務
                            Screen.MousePointer = vbHourglass
                            frm100101_5.Show
                            frm100101_5.Tag = Pub_RplStr(Trim(GrdDataList.TextMatrix(i, 6)))
                            frm100101_5.StrMenu
                            Screen.MousePointer = vbDefault
                      Case "LA"            '顧問
                            Screen.MousePointer = vbHourglass
                            frm100101_6.Show
                            frm100101_6.Tag = Pub_RplStr(Trim(GrdDataList.TextMatrix(i, 6)))
                            frm100101_6.StrMenu
                            Screen.MousePointer = vbDefault
                      Case Else                  '服務
                           Select Case Pub_RplStr(Trim(GrdDataList.TextMatrix(i, 6)))
                               Case "TB"    '條碼
                                  Screen.MousePointer = vbHourglass
                                  frm100101_7.Show
                                  frm100101_7.Tag = Pub_RplStr(Trim(GrdDataList.TextMatrix(i, 6)))
                                  frm100101_7.StrMenu
                                  Screen.MousePointer = vbDefault
                               Case "TM"
                                  Screen.MousePointer = vbHourglass
                                  frm100101_8.Show
                                  frm100101_8.Tag = Pub_RplStr(Trim(GrdDataList.TextMatrix(i, 6)))
                                  frm100101_8.StrMenu
                                  Screen.MousePointer = vbDefault
                               Case "TD"
                                  Screen.MousePointer = vbHourglass
                                  frm100101_9.Show
                                  frm100101_9.Tag = Pub_RplStr(Trim(GrdDataList.TextMatrix(i, 6)))
                                  frm100101_9.StrMenu
                                  Screen.MousePointer = vbDefault
                               Case "TC", "CFC"
                                  Screen.MousePointer = vbHourglass
                                  frm100101_A.Show
                                  frm100101_A.Tag = Pub_RplStr(Trim(GrdDataList.TextMatrix(i, 6)))
                                  frm100101_A.StrMenu
                                  Screen.MousePointer = vbDefault
                               Case Else
                                  Screen.MousePointer = vbHourglass
                                  frm100101_B.Show
                                  frm100101_B.Tag = Pub_RplStr(Trim(GrdDataList.TextMatrix(i, 6)))
                                  frm100101_B.StrMenu
                                  Screen.MousePointer = vbDefault
                            End Select
                  End Select
               Else
                  MsgBox "無本所案號！", vbInformation
               End If
               Me.Enabled = True
               Exit Sub
            End If
         Next i
         Me.Enabled = True
         
      Case 5 '進度
         Me.Enabled = False
         For i = 1 To GrdDataList.Rows - 1
            GrdDataList.col = 0
            GrdDataList.row = i
            If Trim(GrdDataList.Text) = "V" Then
               GrdDataList.col = 0
               GrdDataList.Text = ""
               For j = 0 To GrdDataList.Cols - 1
                   GrdDataList.col = j
                   GrdDataList.CellBackColor = QBColor(15)
               Next j
               GrdDataList.col = 12
               If GrdDataList.Text <> "" Then
                  'Modified by Morgan 2016/3/24 排除母層是共同查詢
'                  If UCase(m_PrevForm.Name) <> UCase("frm100101_2") Then
                     fnCloseAllFrm100 'Added by Morgan 2016/2/22
'                  End If
                  'end 2016/3/24
                  If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                  End If
                  Screen.MousePointer = vbHourglass
                  frm100101_2.Show
                  frm100101_2.Tag = Trim(GrdDataList.TextMatrix(i, 6))
                  'frm100101_2.cmdOK(6).Visible = False
                  frm100101_2.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               Else
                  MsgBox "無本所案號！", vbInformation
               End If
            End If
         Next i
         Me.Enabled = True
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtCRL02(0).Text = strSrvDate(2) '填單日期(起)
   txtCRL02(1).Text = strSrvDate(2) '填單日期(迄)
   txtSales.Text = strUserNum
   lblSalesName = strUserName
   
   SetDataListWidth
   SetAuthorityCon 'Add by Morgan 2011/1/11
   
   'Add By Sindy 2023/7/17
   If Pub_StrUserSt03 = "M51" Then
      Frame1.Visible = True
   Else
      Me.Height = 6168
      Frame1.Visible = False
   End If
   '2023/7/17 END
   
   txtPrintType = "1"
   cmdOK(1).Caption = "明細(&D)"
   txtPCnt.Visible = False
   cmdOK(1).Enabled = False
'   cmdOK(3).Enabled = False
   cmdOK(0).Default = True
   cmdState = -1
   
   'Add By Sindy 2022/10/20
   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
'      cmdOK(3).Visible = False
      Check1.Visible = False
      Label1(122).Visible = False
      txtPrintType.Visible = False
      cmdOK(4).Visible = True
      cmdOK(5).Visible = True
      GrdDataList.Top = txtPrintType.Top: GrdDataList.Height = GrdDataList.Height + txtPrintType.Height
      Label1(1).Visible = True
      Label1(1).Caption = "電子收文啟用日=" & 接洽單電子收文啟用日
   End If
   lblSalesName.BackColor = &H8000000F
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Sindy 2022/12/15 若接洽單已開需關閉
   If PUB_CheckFormExist("frm090801_Q") = True Then
      Unload frm090801_Q
   End If
   '2022/12/15 END
   Set frm12040152 = Nothing
End Sub

Private Sub GrdDataList_Click()
   GrdDataList.Visible = False
   
   '依點選的欄位做排序
   If GrdDataList.MouseRow = 0 Then
      If GrdDataList.MouseCol <> 0 Then
         m_intRow = GrdDataList.MouseRow
         m_intCol = GrdDataList.MouseCol
         GrdDataList.row = m_intRow
         GrdDataList.col = m_intCol
         Select Case m_intCol
            Case 3, 4
               '數字
               If m_blnColOrderAsc = True Then
                   Me.GrdDataList.Sort = 3 '昇冪
                   m_blnColOrderAsc = False
               Else
                   Me.GrdDataList.Sort = 4 '降冪
                   m_blnColOrderAsc = True
               End If
           Case Else
               '字串
               If m_blnColOrderAsc = True Then
                   Me.GrdDataList.Sort = 5 '昇冪
                   m_blnColOrderAsc = False
               Else
                   Me.GrdDataList.Sort = 6 '降冪
                   m_blnColOrderAsc = True
               End If
         End Select
      End If
   End If
   
   '勾選
   GrdDataList.row = GrdDataList.MouseRow
   GrdDataList.col = 0
   If GrdDataList.row <> 0 Then
      If GrdDataList.Text = "V" Then
         GrdDataList.Text = ""
         For i = 0 To GrdDataList.Cols - 1
            GrdDataList.col = i
            GrdDataList.CellBackColor = QBColor(15)
         Next i
      Else
         GrdDataList.Text = "V"
         For i = 0 To GrdDataList.Cols - 1
            GrdDataList.col = i
            GrdDataList.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
   
   GrdDataList.Visible = True
End Sub

Private Sub GrdDataList_Sort()
   GrdDataList.Visible = False
   '依點選的欄位做排序
   If m_intRow = 0 Then
      If m_intCol <> 0 Then
         GrdDataList.row = m_intRow
         GrdDataList.col = m_intCol
         Select Case m_intCol
            Case 3, 4
               '數字
               If m_blnColOrderAsc = False Then
                   Me.GrdDataList.Sort = 3 '昇冪
               Else
                   Me.GrdDataList.Sort = 4 '降冪
               End If
           Case Else
               '字串
               If m_blnColOrderAsc = False Then
                   Me.GrdDataList.Sort = 5 '昇冪
               Else
                   Me.GrdDataList.Sort = 6 '降冪
               End If
         End Select
      End If
   End If
   GrdDataList.Visible = True
End Sub

Private Sub SearchData()
Dim strWhere As String
   
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   
   strWhere = ""
   If txtSystem <> "" And txtCode(0) <> "" Then
      If txtCode(1) = "" Then txtCode(1) = "0"
      If txtCode(2) = "" Then txtCode(2) = "00"
      strWhere = strWhere & " and CRL07='" & txtSystem & "' and CRL08='" & txtCode(0) & "' and CRL09='" & txtCode(1) & "' and CRL10='" & txtCode(2) & "' "
   'Add By Sindy 2023/1/10
   ElseIf txtSystem <> "" And txtCode(0) = "" Then
      strWhere = strWhere & " and CRL07='" & txtSystem & "' "
   End If
   If txtCRL01 <> "" Then
      strWhere = strWhere & " and CRL01='" & txtCRL01 & "' "
      
      '抓此接洽單的填表日期
      strSql = "select crL01,crL02 from consultrecordlist where crl01='" & txtCRL01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         txtCRL02(0) = TransDate(RsTemp.Fields("crL02"), 1)
         txtCRL02(1) = TransDate(RsTemp.Fields("crL02"), 1)
      End If
   End If
   '2023/1/10 END
   'Add By Sindy 2023/3/3
   If txtCRL55 <> "" Then
      'Modify By Sindy 2024/5/7
      'strWhere = strWhere & " and CRL55='" & txtCRL55 & "' "
      strWhere = strWhere & " and instr(CRL55,'" & txtCRL55 & "')>0 "
      '2024/5/7 END
      
      If txtCRL01 = "" Then
         '抓此接洽單的填表日期
         'Modify By Sindy 2024/5/7
         'strSql = "select crL01,crL02 from consultrecordlist where CRL55='" & txtCRL55 & "' order by crL02 asc"
         strSql = "select crL01,crL02 from consultrecordlist where instr(CRL55,'" & txtCRL55 & "')>0 order by crL02 asc"
         '2024/5/7 END
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            txtCRL02(0) = TransDate(RsTemp.Fields("crL02"), 1)
            txtCRL02(1) = TransDate(strSrvDate(1), 1)
         End If
      End If
   End If
   '2023/3/3 END
   
   'Add By Sindy 2023/7/17
   '電腦中心使用
   If Frame1.Visible = True Then
      '關連表單編號
      If txtCRL65 <> "" Then
         strWhere = strWhere & " and CRL65='" & txtCRL65 & "' "
      End If
   End If
   '2023/7/17 END
   
   'Modify by Morgan 2010/12/20
   If txtSales <> "" Then
      'Modify By Sindy 2022/10/19 智權人員或代操作人員都可以查詢出來
      strWhere = strWhere & " and (CRL03='" & txtSales & "' or CRL78='" & txtSales & "')"
   End If
   
   'Modify By Sindy 2023/3/3 + And txtCRL55 = "":有輸案源單號時可以查詢"法律所案源預存的接洽單"
   If Pub_StrUserSt03 <> "M51" And txtCRL55 = "" Then
      'Added by Morgan 2020/6/22
      '剔除法律所案源預存的接洽單
      strWhere = strWhere & " and not exists(select * from caseprogress,lawofficesource where cp140=crl01 and los10(+)=cp09 and los10 is not null)"
   End If
   
   'Modify By Sindy 2022/9/5
   If DBDATE(Trim(txtCRL02(0).Text)) >= 接洽單電子收文啟用日 Then
      strSql = "select ' ' as V,decode(F0309," & ShowFlow表單狀態中文 & ",F0309),CRL01,SUBSTR(' '||sqldatet(CRL02),-9),ST02,decode(CRL05,'1','國內','大至台'), " & _
                             "CRL07||'-'||CRL08||'-'||CRL09||'-'||CRL10,CRL17, " & _
                             "GetCRCaseNmFee(crl01,'2') 案件性質,'','','',CRL08 " & _
                    "from consultrecordlist,staff,flow003 " & _
                   "Where CRL02>=" & ChangeTStringToWString(txtCRL02(0)) & " and CRL02<=" & ChangeTStringToWString(txtCRL02(1)) & " " & _
                     "and CRL03=ST01(+) " & strWhere & m_strAuthorityCon & _
                     " and CRL01=F0301(+)" & _
                   " order by CRL02,CRL03,CRL07,CRL08,CRL09,CRL10,CRL01 "
   Else
   '2022/9/5 END
      strSql = "select ' ' as V,'' 狀態,CRL01,SUBSTR(' '||sqldatet(CRL02),-9),ST02,decode(CRL05,'1','國內','大至台'), " & _
                             "CRL07||'-'||CRL08||'-'||CRL09||'-'||CRL10,CRL17, " & _
                             "decode(CRL15,'000',cpm1.CPM03,cpm1.CPM04), " & _
                             "decode(CRL15,'000',cpm2.CPM03,cpm2.CPM04), " & _
                             "decode(CRL15,'000',cpm3.CPM03,cpm3.CPM04), " & _
                             "decode(CRL15,'000',cpm4.CPM03,cpm4.CPM04),CRL08 " & _
                    "from consultrecordlist,staff, " & _
                         "casepropertymap cpm1,casepropertymap cpm2,casepropertymap cpm3,casepropertymap cpm4 " & _
                   "Where CRL02>=" & ChangeTStringToWString(txtCRL02(0)) & " and CRL02<=" & ChangeTStringToWString(txtCRL02(1)) & " " & _
                     "and CRL03=ST01(+) " & _
                     "and CRL07=cpm1.CPM01(+) and CRL19=cpm1.CPM02(+) " & _
                     "and CRL07=cpm2.CPM01(+) and CRL24=cpm2.CPM02(+) " & _
                     "and CRL07=cpm3.CPM01(+) and CRL29=cpm3.CPM02(+) " & _
                     "and CRL07=cpm4.CPM01(+) and CRL34=cpm4.CPM02(+) " & strWhere & m_strAuthorityCon & _
                   "order by CRL02,CRL03,CRL07,CRL08,CRL09,CRL10,CRL01 "
   End If
   Screen.MousePointer = vbHourglass
   GrdDataList.Clear
   GrdDataList.Rows = 2
   SetDataListWidth
   GrdDataList.FixedCols = 0
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 Then
       Label1(3).Caption = "共　" & adoRecordset.RecordCount & "　件"
       cmdOK(1).Enabled = True
'       cmdOK(3).Enabled = True
       Set GrdDataList.Recordset = adoRecordset
   Else
       Label1(3).Caption = "共　0　件"
       cmdOK(1).Enabled = False
'       cmdOK(3).Enabled = False
       ShowNoData
       GrdDataList.Clear
   End If
   SetDataListWidth
   'GrdDataList.FixedCols = 4
   CheckOC
   Call GrdDataList_Sort
   
   '若只有一筆資料, 則直接設定為點選此筆資料
   With Me.GrdDataList
      If .Rows = 2 Then
         .row = 1
         .col = 1
         If .Text <> "" Then
           .Visible = False
           .row = 1
           .col = 0
           .Text = "V"
           For i = 0 To .Cols - 1
               .col = i
               .CellBackColor = &HFFC0C0
           Next i
           .Visible = True
         End If
      End If
   End With
   
   Screen.MousePointer = vbDefault
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim s As Integer

TxtValidate = False

'填單日期
If Len(Trim(txtCRL02(0).Text)) = 0 Then
   s = MsgBox("填單日期(起)不可空白", , "輸入條件錯誤")
   txtCRL02(0).SetFocus
   Exit Function
End If
If Len(Trim(txtCRL02(1).Text)) = 0 Then
   s = MsgBox("填單日期(迄)不可空白", , "輸入條件錯誤")
   txtCRL02(1).SetFocus
   Exit Function
End If
'Add By Sindy 2022/9/5 電子收文上線,因資料庫結構不同,不可合併查詢
If DBDATE(Trim(txtCRL02(0).Text)) < 接洽單電子收文啟用日 Then
   If DBDATE(Trim(txtCRL02(1).Text)) >= 接洽單電子收文啟用日 Then
      s = MsgBox("填單日期(迄)錯誤，必須小於" & Val(接洽單電子收文啟用日) - 19110000 & vbCrLf & vbCrLf & _
                 "(電子收文上線，因資料庫結構不同，不可合併查詢)", , "輸入條件錯誤")
      txtCRL02(1).SetFocus
      Exit Function
   End If
End If
'2022/9/5 END

'員工編號
'Add by Morgan 2010/12/20
'Modify By Sindy 2023/1/16 + And Pub_GetSpecMan("全所智權部主管") <> strUserNum And strUserNum <> "71011"
If Pub_StrUserSt03 <> "M51" And txtSales.Enabled = False Then 'And Pub_GetSpecMan("全所智權部主管") <> strUserNum And strUserNum <> "71011" Then
   If Len(Trim(txtSales.Text)) = 0 Then
      s = MsgBox("員工編號不可空白", , "輸入條件錯誤")
      txtSales.SetFocus
      Exit Function
   End If
'Add By Sindy 2023/1/16
ElseIf bolSpecMan = True Then
   If Len(Trim(txtSales.Text)) = 0 Then
      s = MsgBox("員工編號不可空白", , "輸入條件錯誤")
      txtSales.SetFocus
      Exit Function
   End If
   '2023/1/16 END
End If

If Me.txtCRL02(0).Enabled = True Then
   Cancel = False
   Call txtCRL02_Validate(0, Cancel)
   If Cancel = True Then
      txtCRL02(0).SetFocus
      Exit Function
   End If
End If
If Me.txtCRL02(1).Enabled = True Then
   Cancel = False
   Call txtCRL02_Validate(1, Cancel)
   If Cancel = True Then
      txtCRL02(1).SetFocus
      Exit Function
   End If
End If
If Me.txtSales.Enabled = True Then
   Cancel = False
   txtSales_Validate Cancel
   If Cancel = True Then
      txtSales.SetFocus
      Exit Function
   End If
End If

TxtValidate = True
End Function

'Add By Sindy 2022/12/23
Private Sub grdDataList_DblClick()
   cmdState = 1
   PubShowNextData
End Sub

Private Sub txtCRL02_GotFocus(Index As Integer)
   InverseTextBox txtCRL02(Index)
End Sub

Private Sub txtCRL02_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtCRL02_Validate(Index As Integer, Cancel As Boolean)
   If CheckIsTaiwanDate(txtCRL02(Index), False) = False Then
      Cancel = True
      MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
      Call txtCRL02_GotFocus(Index)
      Exit Sub
   End If
   If Index = 0 Then
      If txtCRL02(Index) <> "" And txtCRL02(Index + 1) = "" Then
         txtCRL02(Index + 1) = txtCRL02(Index)
      End If
   ElseIf Index = 1 Then
      If RunNick2(txtCRL02(Index - 1), txtCRL02(Index)) Then
         Call txtCRL02_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub txtSales_GotFocus()
   InverseTextBox txtSales
End Sub

Private Sub txtSales_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_Validate(Cancel As Boolean)
   If txtSales.Text = "" Then
      lblSalesName = ""
   Else
      lblSalesName = GetStaffName(txtSales, True)
      If lblSalesName = "" Then
         MsgBox "員工編號錯誤！查無此員工！", vbInformation
         Call txtSales_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub txtSystem_GotFocus()
   CloseIme
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   CloseIme
   txtCode(Index).SelStart = 0
   txtCode(Index).SelLength = Len(txtCode(Index))
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 1 Then KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPrintType_GotFocus()
   TextInverse txtPrintType
End Sub

Private Sub txtPrintType_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
       KeyAscii = 0
   End If
End Sub

Private Sub txtPrintType_Validate(Cancel As Boolean)
   Select Case Val(txtPrintType)
   Case 1, 2
      If Val(txtPrintType) = 1 Then
         cmdOK(1).Caption = "明細(&D)"
         txtPCnt.Visible = False
      ElseIf Val(txtPrintType) = 2 Then
         cmdOK(1).Caption = "列印(&P)　　份"
         txtPCnt.Visible = True
      End If
   Case Else
      s = MsgBox("輸出方式只能 1 或 2 !!", , "USER 輸入錯誤")
      Cancel = True
   End Select
   If Cancel Then TextInverse txtPrintType
End Sub

'Add by Morgan 2011/1/11 參考frm100123(業務期限資料查詢)設定
Private Sub SetAuthorityCon()
'testing code
'   pub_strUserOffice = PUB_GetST06(strUserNum)
'   Pub_StrUserSt15 = PUB_GetStaffST15(strUserNum, "1")
'   Pub_StrUserSt03 = PUB_GetST03(strUserNum)
   
   'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
   Call PUB_SetFormSaleDept(strUserNum, , , , txtSales, bolSpecMan, strSpecCode, , , , m_strAuthorityCon, , True)
   'Modify By Sindy 2023/1/16 副總可以看全所
   If strUserNum = "71011" Or InStr(UCase(App.EXEName), "WRITER") > 0 Then
      m_strAuthorityCon = ""
      txtSales.Enabled = True
      txtSales.Text = "": lblSalesName = ""
   End If
   'Add By Sindy 2023/5/16
   If txtSales = "P1004" Then
      txtCRL02(0) = "1120101"
   End If
   '2023/5/16 END
   
'   txtSales.Enabled = False
'   Select Case strUserNum
'      '外商陳經理可看全所CFT,FCT,S,CFC
'      Case "68005"
'         m_strAuthorityCon = " and CRL07 in ('CFT','FCT','S','CFC')"
'         txtSales.Enabled = True
'
''cancel by sonia 2014/6/9
''      '蔣律師可看中所全部
''      Case "79037"
''         m_strAuthorityCon = " and st06 = '" & pub_strUserOffice & "'"
''         txtSales.Enabled = True
''end 2014/6/9
'
'      '小真,杜副總可看全部
'      'modify by sonia 2014/6/9 +美珍77027
'      Case "65001", "68006", "77027"
'         txtSales.Enabled = True
'
'      '杜燕文,劉大愛可看S31
'      'modify by sonia 2016/8/22劉大愛78007改為蘇嫄媛79053
'      Case "74018", "79053"
'         m_strAuthorityCon = " and st15 = 'S31'"
'         txtSales.Enabled = True
'
'      '王協理可看專利處
'      Case "71011"
'         m_strAuthorityCon = " and st15>='P10' and st15<='P19'"
'         txtSales.Enabled = True
'
'      '葉經理可看商標處
'      'modify by sonia 2016/2/24 +69008
'      Case "67002", "69008"
'         m_strAuthorityCon = " and st15>='P20' and st15<='P29'"
'         txtSales.Enabled = True
'
'      Case Else
'
'         m_stST05 = PUB_GetST05(strUserNum)
'         Select Case m_stST05
'            '電腦中心,財務,總經理看全部
'            'Modify by Morgan 2011/4/20 內專程序也可看全部(自動收文單筆發文時要列印)
'            '2015/7/28 MODIFY BY SONIA +主任秘書(等級08)可看全部
'            Case "00", "01", "75", "73", "08"
'               txtSales.Enabled = True
'
'            '各區主管
'            Case "SM"
'               '林永生71003可看中所全部智權人員
'               If strUserNum = "71003" Then
'                  m_strAuthorityCon = " and st15 like 'S2%'"
'               '簡協理可看北所全部智權人員
'               ElseIf strUserNum = "69005" Then
'                  'Removed by Morgan 2019/12/31 杜主秘說加開簡協理69005可看全所
'                  'm_strAuthorityCon = " and st15 like 'S1%'"
'                  'end 2019/12/31
'               Else
'                  m_strAuthorityCon = " and st15='" & Pub_StrUserSt15 & "'"
'               End If
'               txtSales.Enabled = True
'
'            '外商主管  王宗珮、洪琬姿、葉易雲 可看同組
'            Case "21", "26", "28"
'               m_strAuthorityCon = " and st15='" & Pub_StrUserSt15 & "' and st16='" & PUB_GetStaffST16(strUserNum) & "'"
'               txtSales.Enabled = True
'
'            '智權人員
'            Case "SA"
'               '帶人主管
'               If PUB_GetST05Limits(strUserNum) Then
'                  m_strAuthorityCon = " and '" & strUserNum & "' in (st01,st52)"
'                  txtSales.Enabled = True
'               Else
'                  txtSales.Enabled = False
'               End If
'
'            '其他只能看自己
'            Case Else
'               txtSales.Enabled = False
'
'         End Select
'   End Select
End Sub
