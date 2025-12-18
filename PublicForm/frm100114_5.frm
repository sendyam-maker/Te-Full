VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100114_5 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶/代理人查詢"
   ClientHeight    =   5112
   ClientLeft      =   1692
   ClientTop       =   3108
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5112
   ScaleWidth      =   8760
   Begin VB.CommandButton cmdOK 
      Caption         =   "代理人資料(&O)"
      Height          =   345
      Index           =   0
      Left            =   4430
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   10
      Width           =   1300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "關係企業(&R)"
      Height          =   345
      Index           =   2
      Left            =   5745
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "聯絡人(&T)"
      Height          =   345
      Index           =   7
      Left            =   6900
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   10
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   10
      Left            =   1095
      TabIndex        =   5
      Top             =   1050
      Width           =   1600
   End
   Begin VB.OptionButton Option1 
      Caption         =   "E-Mail："
      Height          =   180
      Index           =   3
      Left            =   100
      TabIndex        =   4
      Top             =   1110
      Width           =   1000
   End
   Begin VB.Frame Frame2 
      Height          =   360
      Left            =   4590
      TabIndex        =   11
      Top             =   345
      Width           =   2475
      Begin VB.OptionButton Option3 
         Caption         =   "字首比對"
         Height          =   180
         Index           =   0
         Left            =   96
         TabIndex        =   2
         Top             =   144
         Width           =   1125
      End
      Begin VB.OptionButton Option3 
         Caption         =   "模糊比對"
         Height          =   180
         Index           =   1
         Left            =   1230
         TabIndex        =   3
         Top             =   144
         Value           =   -1  'True
         Width           =   1080
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "國籍："
      Height          =   204
      Index           =   2
      Left            =   100
      TabIndex        =   6
      Top             =   1425
      Width           =   900
   End
   Begin VB.OptionButton Option1 
      Caption         =   "名稱："
      Height          =   204
      Index           =   1
      Left            =   100
      TabIndex        =   0
      Top             =   465
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   9
      Left            =   1035
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1380
      Width           =   1092
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   345
      Left            =   3600
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   10
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   4
      Left            =   7900
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   10
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   3300
      Left            =   30
      TabIndex        =   15
      Top             =   1740
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   5821
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   15
   End
   Begin VB.Label Label1 
      Caption         =   "＊：舊的名稱＄：有呆帳"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   7290
      TabIndex        =   22
      Top             =   780
      Width           =   1260
   End
   Begin VB.Label Label11 
      Caption         =   "●：特殊客戶"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   7410
      TabIndex        =   21
      Top             =   1200
      Width           =   1260
   End
   Begin VB.Label Label12 
      Caption         =   "⊕：不得代理"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   7290
      TabIndex        =   20
      Top             =   1410
      Width           =   1260
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   330
      Index           =   0
      Left            =   1035
      TabIndex        =   1
      Top             =   420
      Width           =   3495
      VariousPropertyBits=   679493659
      Size            =   "6165;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "查詢後雙擊選取資料可帶編號回前畫面"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   3500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "輸入名稱之特取部分, 不要取國家,省份,城市,例：不可輸美商..,廣東..,廣州.."
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   870
      Width           =   5805
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "註：紅色資料不可承接案件"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   4050
      TabIndex        =   13
      Top             =   1440
      Width           =   3030
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "模糊比對"
      Height          =   180
      Left            =   2760
      TabIndex        =   12
      Top             =   1110
      Width           =   720
   End
   Begin VB.Label lbl1 
      Height          =   210
      Index           =   1
      Left            =   2640
      TabIndex        =   10
      Top             =   2670
      Width           =   3270
   End
End
Attribute VB_Name = "frm100114_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/06 改成Form2.0 ; GrdDataList改字型=新細明體-ExtB、txt1(1)改成txtFM2(0)
'Create by Amy2015/08/31 查詢後帶回相關欄位用
Option Explicit

Dim s As Long, i As Long, j As Long, strSql As String
Dim StrToGrid As String
Dim strTp(3) As String
Public cmdState As Integer
Public stBackField As String '前畫面回傳欄位

Private Sub Form_Load()
    Dim intX As Integer, intY As Integer
    Dim sglWidth As Single, sglHeight As Single
           
    Me.Icon = LoadPicture(strIcoPath)
    strFormName = Name
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    
    SetDataListWidth
    Option1(2).Value = False
    Option1(3).Value = False
    cmdState = -1
    'Add by Amy 2015/12/04 財務不需出現「聯絡人」按鈕,預設「字首比對」-婉莘
    If InStr(UCase(App.EXEName), "ACCOUNT") > 0 Then
        cmdOK(7).Visible = False
        cmdOK(4).Left = 6900
        Option3(0).Value = True
    End If
    
    'Added by Lydia 2017/12/05 改由啟用日控制
    If strSrvDate(1) >= 國外部關聯企業啟用日 Then cmdOK(2).Caption = "關聯企業(&R)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Tag = MsgText(601)
    stBackField = MsgText(601)
    tool13_enabled
    Frmacc2210.Enabled = True
    Frmacc2210.Show
    Set frm100114_5 = Nothing
End Sub

Private Sub SetDataListWidth()
    grdDataList.row = 0
    grdDataList.col = 0: grdDataList.Text = "V"
    grdDataList.ColWidth(0) = 200
    grdDataList.CellAlignment = flexAlignCenterCenter
    grdDataList.col = 1: grdDataList.Text = "編號"
    grdDataList.ColWidth(1) = 1200
    grdDataList.CellAlignment = flexAlignCenterCenter
    grdDataList.col = 2: grdDataList.Text = "名稱"
    grdDataList.ColWidth(2) = 4000
    grdDataList.CellAlignment = flexAlignCenterCenter
    grdDataList.col = 3: grdDataList.Text = "國籍"
    grdDataList.ColWidth(3) = 1200
    grdDataList.CellAlignment = flexAlignCenterCenter
    grdDataList.col = 4: grdDataList.Text = "智權人員"
    grdDataList.ColWidth(4) = 800
    grdDataList.CellAlignment = flexAlignCenterCenter
   
    grdDataList.col = 5: grdDataList.Text = "狀態"
    grdDataList.ColWidth(5) = 1000
    grdDataList.CellAlignment = flexAlignCenterCenter
    grdDataList.col = 6: grdDataList.Text = "備註"
    grdDataList.ColWidth(6) = 2000
    grdDataList.CellAlignment = flexAlignLeftCenter
    '因查詢服務對造資料需依sp09抓不智權人員資料,故加申請國家
    grdDataList.col = 7: grdDataList.Text = "申請國家"
    grdDataList.ColWidth(7) = 0
    grdDataList.col = 8: grdDataList.Text = "總收文號"
    grdDataList.ColWidth(8) = 0
    grdDataList.col = 9: grdDataList.Text = "案件性質"
    grdDataList.ColWidth(9) = 0
    grdDataList.col = 10: grdDataList.Text = "收文日"
    grdDataList.ColWidth(10) = 0
    
   'Added by Lydia 2017/02/14 關聯企業
   If grdDataList.Cols > 11 Then 'Added by Lydia 2017/12/28
        grdDataList.col = 11: grdDataList.Text = "關聯編號"
        grdDataList.ColWidth(11) = 0
        grdDataList.col = 12: grdDataList.Text = "關聯名稱"
        grdDataList.ColWidth(12) = 0
        grdDataList.col = 13: grdDataList.Text = "關聯關係"
        grdDataList.ColWidth(13) = 0
        grdDataList.col = 14: grdDataList.Text = "關聯說明"
        grdDataList.ColWidth(14) = 0
        grdDataList.FixedCols = 0
   End If 'Added by Lydia 2017/12/28
   'end 2017/02/14
End Sub

Private Sub cmdOK_Click(Index As Integer)
   '紀錄作用按鍵
   cmdState = Index
   PubShowNextData
   Exit Sub
End Sub

Public Sub PubShowNextData()
    Dim strTmp As String
    
    If grdDataList.Rows = 2 And grdDataList.TextMatrix(1, 1) = MsgText(601) And cmdState <> 4 Then Exit Sub
    
    Select Case cmdState
        Case 0 '代理人資料
            Me.Enabled = False
            For i = 1 To grdDataList.Rows - 1
                grdDataList.col = 0
                grdDataList.row = i
                If Trim(grdDataList.Text) = "V" Then
                    grdDataList.col = 0
                    grdDataList.Text = ""
                    grdDataList.col = 1
                    If Right(grdDataList.Text, 1) = "⊕" Then
                        For j = 0 To grdDataList.Cols - 1
                            grdDataList.col = j
                            grdDataList.CellBackColor = &H8080FF
                        Next j
                    Else
                        For j = 0 To grdDataList.Cols - 1
                            If j <> 1 Then
                                grdDataList.col = j
                                grdDataList.CellBackColor = QBColor(15)
                            End If
                        Next j
                    End If
                    grdDataList.col = 1
                    If Not IsNull(grdDataList.Text) Then
                        If fnSaveParentForm(Me) = False Then
                            Me.Enabled = True
                            Exit Sub
                        End If
                        Screen.MousePointer = vbHourglass
                        strExc(1) = Pub_RplStr(grdDataList.Text)
                        Select Case Left(strExc(1), 1)
                            Case "X"
                                If Mid(strExc(1), 10, 1) = "-" Then
                                    strExc(1) = Left(strExc(1), 9)
                                End If
                                frm100101_11.Show
                                frm100101_11.Tag = strExc(1)
                                frm100101_11.StrMenu
                            Case "Y"
                                If Mid(strExc(1), 10, 1) = "-" Then
                                    strExc(1) = Left(strExc(1), 9)
                                End If
                                frm100101_10.Show
                                frm100101_10.Tag = strExc(1)
                                frm100101_10.StrMenu
                        End Select
                        Screen.MousePointer = vbDefault
                        Me.Enabled = True
                        Exit Sub
                    End If
                End If
            Next i
            Me.Enabled = True
        Case 2 '關係企業
            Me.Enabled = False
            strExc(9) = "" 'Added by Lydia 2017/08/18 勾選清單
            'Modified by Lydia 2017/12/05 改由啟用日控制
            If strSrvDate(1) < 國外部關聯企業啟用日 Then
                cnnConnection.Execute "delete from r100114 where id='" & strUserNum & "' "
            End If
            'end 2017/12/05
            For i = 1 To grdDataList.Rows - 1
              grdDataList.col = 0
              grdDataList.row = i
              If Trim(grdDataList.Text) = "V" Then
                  grdDataList.col = 1
                  Screen.MousePointer = vbHourglass
                  'Modified by Lydia 2017/12/05 改由啟用日控制
                  If strSrvDate(1) < 國外部關聯企業啟用日 Then
                      Call StrMenu(Pub_RplStr(grdDataList.Text))
                  Else
                      'Added by Lydia 2017/02/14 抓關聯企業改成模組,暫存R100114_1
                      'Modified by Lydia 2017/08/18 是否清除先前記錄
                      'j = PUB_GetR100114_1(Me.Name, Pub_RplStr(GrdDataList.Text))
                      j = PUB_GetR100114_1(IIf(strExc(9) = "", True, False), Me.Name, Pub_RplStr(grdDataList.Text))
                      strExc(9) = strExc(9) & IIf(strExc(9) <> "", ",", "") & Pub_RplStr(grdDataList.Text)
                      'end 2017/08/18
                  End If
                  'end 2017/12/05
                  
                  cmdOK(2).Enabled = False
                  Screen.MousePointer = vbDefault
              End If
            Next i
            'Modified by Lydia 2017/12/05 改由啟用日控制
            If strSrvDate(1) < 國外部關聯企業啟用日 Then
               Call StrMenu1
            Else
               'Added by Lydia 2017/02/14 抓關聯企業改成模組,暫存R100114_1
               If j > 1 Then Call StrMenu1
            End If
            'end 2017/12/05
            Me.Enabled = True
        Case 4 '結束
            fnCloseAllFrm100
        Case 7 '聯絡人
            Me.Enabled = False
            For i = 1 To grdDataList.Rows - 1
                grdDataList.col = 0
                grdDataList.row = i
                If Trim(grdDataList.Text) = "V" Then
                    grdDataList.col = 0
                    grdDataList.Text = ""
                    grdDataList.col = 1
                    If Right(grdDataList.Text, 1) = "⊕" Then
                        For j = 0 To grdDataList.Cols - 1
                            grdDataList.col = j
                            grdDataList.CellBackColor = &H8080FF
                        Next j
                    Else
                        For j = 0 To grdDataList.Cols - 1
                            If j <> 1 Then
                                grdDataList.col = j
                                grdDataList.CellBackColor = QBColor(15)
                            End If
                        Next j
                    End If
                    If fnSaveParentForm(Me) = False Then
                        Me.Enabled = True
                        Exit Sub
                    End If
                    grdDataList.col = 1
                    Screen.MousePointer = vbHourglass
                    strExc(1) = Pub_RplStr(grdDataList.Text)
                    strExc(2) = "F"
                    If Left(strExc(1), 1) = "X" Then
                        strExc(0) = "select st03 from customer,staff where cu01(+)='" & Left(strExc(1), 8) & "' and cu02(+)='" & Mid(strExc(1), 9, 1) & "' and st01(+)=cu13"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           strExc(2) = "" & RsTemp.Fields(0)
                        End If
                    End If
                    If Left(strExc(2), 1) = "F" Then
                        frm100101_17.Show
                        frm100101_17.Tag = strExc(1)
                        frm100101_17.StrMenu
                    Else
                        frm100101_18.Show
                        frm100101_18.Tag = strExc(1)
                        frm100101_18.StrMenu
                    End If
                    Screen.MousePointer = vbDefault
                    grdDataList.col = 0
                    grdDataList.Text = ""
                    grdDataList.col = 1
                    If Right(grdDataList.Text, 1) = "⊕" Then
                        For j = 0 To grdDataList.Cols - 1
                            grdDataList.col = j
                            grdDataList.CellBackColor = &H8080FF
                        Next j
                    Else
                        For j = 0 To grdDataList.Cols - 1
                            If j <> 1 Then
                                grdDataList.col = j
                                grdDataList.CellBackColor = QBColor(15)
                            End If
                        Next j
                    End If
                    Me.Enabled = True
                    Exit Sub
                End If
            Next i
            Me.Enabled = True
    End Select
End Sub

Private Sub cmdSearch_Click()
    Dim StrSQLa As String
    Dim StrSqlB As String
    Dim strSQLc As String
    Dim strSQLD As String
    Dim strCheckWay As String
    Dim strSQLE As String
    Dim strFields As String 'Added by Lydia 2017/02/14 設定關聯代號欄位
    
    If Option1(1).Value = True Then
        If Len(Trim(txtFM2(0))) = 0 Then
            s = MsgBox("名稱不可空白", , "USER 輸入資料錯誤")
            txtFM2(0).SetFocus
            Exit Sub
        End If
    Else
        If Option1(2).Value = True Then
            If Len(Trim(txt1(9))) = 0 Then
                s = MsgBox("國籍不可空白", , "USER 輸入資料錯誤")
                txt1(9).SetFocus
                Exit Sub
            End If
        End If
    End If
    
    'E-mail
    If Option1(3).Value = True Then
       If Len(Trim(txt1(10))) = 0 Then
           s = MsgBox("條件不可空白", , "輸入條件錯誤")
           txt1(10).SetFocus
           Exit Sub
        End If
    End If
   
    Screen.MousePointer = vbHourglass
    grdDataList.Clear
    grdDataList.Rows = 2
    SetDataListWidth
    strFields = ",'' AS 關聯編號,'' AS 關聯名稱,'' AS 關聯關係,'' AS 關聯說明 " 'Added by Lydia 2017/02/14
   
    '若國籍為"013"或"020"則名稱抓中-->英-->日, 否則抓英-->中-->日
    'Modified by Lydia 2020/08/21
'    StrSQLa = "DECODE(FA10,'013',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),'020',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 名稱,"
'    StrSqlB = "DECODE(cu10,'013',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),'020',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),DECODE(cu05,NULL,NVL(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90)) as 名稱,"
'    strSQLc = "DECODE(instr('013,020',pcu09),0,decode(pcu03,NULL,nvl(pcu08,pcu07),rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06)),NVL(pcu08,DECODE(pcu03,NULL,pcu07,rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06)))) as 名稱,"
'    strSQLD = "DECODE(instr('013,020',poc04),0,decode(poc23,NULL,nvl(poc03,poc27),rtrim(poc23||' '||poc24||' '||poc25||' '||poc26)),NVL(poc03,DECODE(poc23,NULL,poc27,rtrim(poc23||' '||poc24||' '||poc25||' '||poc26)))) as 名稱,"
'    strSQLE = "DECODE(instr('013,020',nt08),0,decode(nt03,NULL,nvl(nt02,nt07),rtrim(nt03||' '||nt04||' '||nt05||' '||nt06)),NVL(nt02,DECODE(nt03,NULL,nt07,rtrim(nt03||' '||nt04||' '||nt05||' '||nt06)))) as 名稱,"
    StrSQLa = "Decode(sign(instr('000,001,002,003,004,005,006,007,008,009,013,020',FA10)),1,NVL(FA04,Decode(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),Decode(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 名稱,"
    StrSqlB = "Decode(sign(instr('000,001,002,003,004,005,006,007,008,009,013,020',CU10)),1,NVL(CU04,Decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),Decode(CU05,NULL,NVL(CU04,CU06),CU05||' '||CU88||' '||CU89||' '||CU90 )) as 名稱,"
    strSQLc = "Decode(Sign(Instr('000,001,002,003,004,005,006,007,008,009,013,020',Pcu09)),1,Nvl(Pcu08,Decode(Pcu03,Null,Pcu07,Pcu03||' '||Pcu04||' '||Pcu05||' '||Pcu06)),Decode(Pcu03,Null,Nvl(Pcu07,Pcu08),Pcu03||' '||Pcu04||' '||Pcu05||' '||Pcu06)) as 名稱,"
    strSQLD = "Decode(Sign(Instr('000,001,002,003,004,005,006,007,008,009,013,020',Poc04)),1,Nvl(Poc03,Decode(Poc23,Null,Poc28,Poc23||' '||Poc24||' '||Poc25||' '||Poc26)),Decode(Poc23,Null,Nvl(Poc03,Poc28),Poc23||' '||Poc24||' '||Poc25||' '||Poc26)) as 名稱,"
    strSQLE = "Decode(Sign(Instr('000,001,002,003,004,005,006,007,008,009,013,020',nt08)),1,Nvl(nt02,Decode(nt03,Null,nt07,nt03||' '||nt04||' '||nt05||' '||nt06)),Decode(nt03,Null,Nvl(nt02,nt07),nt03||' '||nt04||' '||nt05||' '||nt06)) as 名稱,"
    'end 2020/08/21
   
    '以名稱查詢
    If Option1(1).Value = True Then
        pub_QL05 = pub_QL05 & ";" & Option1(1).Caption
        '模糊比對
        If Option3(0).Value = False Then
            strCheckWay = ">0"
            pub_QL05 = pub_QL05 & ";" & Option3(1).Caption
            '字首比對
        Else
            strCheckWay = "=1"
            pub_QL05 = pub_QL05 & ";" & Option3(0).Caption
        End If
    
        strTp(3) = ChgSQL(UCase(Trim(txtFM2(0))))
        '查Fagent 代理人 檔
        'Modified by Lydia 2017/02/14 + strfields
        strSql = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,FA04 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(FA04,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & " ) A WHERE FA01=A.A1 AND fa10=na01(+) "
        'Modified by Lydia 2017/02/14 + strfields
        strSql = strSql & " union all Select '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,FA05||' '||FA63||' '||FA64||' '||FA65 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(upper(FA05||' '||FA63||' '||FA64||' '||FA65),'" & UCase(ChgSQL(Trim(txtFM2(0)))) & "')" & strCheckWay & " ) A WHERE FA01=A.A1 AND fa10=NA01(+) "
        'Modified by Lydia 2017/02/14 + strfields
        strSql = strSql & " union all Select '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,FA06 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(FA06,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & " ) A WHERE FA01=A.A1 AND fa10=na01(+) "

        '查customer 客戶 檔
        'Modified by Lydia 2017/02/14 + strfields
        strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,cu04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM customer,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(cu04,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & " ) A WHERE CU01=A.A1 And cu10=na01(+) AND CU13=ST01(+) "
        'Modified by Lydia 2017/02/14 + strfields
        strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,cu05||' '||cu88||' '||cu89||' '||cu90 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM customer,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(upper(cu05||' '||cu88||' '||cu89||' '||cu90),'" & UCase(ChgSQL(Trim(txtFM2(0)))) & "')" & strCheckWay & " ) A WHERE CU01=A.A1 AND cu10=na01(+) AND CU13=ST01(+) "
        'Modified by Lydia 2017/02/14 + strfields
        strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,cu06 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM customer,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(cu06,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & " ) A WHERE CU01=A.A1 AND cu10=na01(+) AND CU13=ST01(+) "
        
    '以國籍查詢
    ElseIf Option1(2).Value = True Then
        'Modified by Lydia 2017/02/14 + strfields
        strSql = "SELECT ''AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號," & StrSQLa & "NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM FAGENT,NATION WHERE INSTR(FA10, '" & txt1(9) & "') = 1 AND fa10=NA01(+) "
        'Modified by Lydia 2017/02/14 + strfields
        strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號," & StrSqlB & "NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM customer,NATION,Staff WHERE INSTR(CU10, '" & txt1(9) & "') = 1 AND cu10=na01(+) AND CU13=ST01(+) "
        pub_QL05 = pub_QL05 & ";" & Option1(2).Caption & txt1(9)
        
    'E-Mail
    ElseIf Option1(3).Value = True Then
        'Modified by Lydia 2017/02/14 + strfields
        'Modified by Lydia 2024/09/18 +財務副本信箱CU200
        strSql = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','')||decode(cu121,'Y','●','') AS 編號,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM CUSTOMER,NATION,Staff  Where (instr(NLS_Upper(CU20),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0 or instr(NLS_Upper(CU115),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0 or instr(NLS_Upper(CU116),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0  or instr(NLS_Upper(CU117),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0 or instr(NLS_Upper(CU118),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 or instr(NLS_Upper(CU200),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 )  and CU10=NA01(+) AND CU13=ST01(+) "
        'Modified by Lydia 2017/02/14 + strfields
        'Modified by Lydia 2018/07/20 +FA105 財務信箱(CF)
        'strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM fagent,nation Where (instr(NLS_Upper(fa16),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa79),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa80),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa81),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 Or InStr(NLS_Upper(fa82),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 )   and  fa10=na01(+)   "
        'Modified by Lydia 2024/09/18 +財務副本信箱FA134
        strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & _
                   " FROM fagent,nation Where (instr(NLS_Upper(fa16),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa79),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa105),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa80),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa81),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 Or InStr(NLS_Upper(fa82),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 Or InStr(NLS_Upper(FA134),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 )   and  fa10=na01(+)   "
        pub_QL05 = pub_QL05 & ";" & Option1(3).Caption & Trim(txt1(10))
    End If
    CheckOC
    strSql = "select * from (" & strSql & ") X order by upper(名稱),編號 "
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        InsertQueryLog (adoRecordset.RecordCount)
        cmdOK(0).Enabled = True
        cmdOK(2).Enabled = True
        Set grdDataList.Recordset = adoRecordset
    Else
        InsertQueryLog (0)
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Me.grdDataList.Visible = False
    CheckOC
    SetDataListWidth
    
    '變色
    With Me.grdDataList
        If .Rows > 0 Then
            For i = 0 To .Rows - 1
                .row = i
                .col = 1
                If Right(.Text, 1) = "$" Then
                    .CellBackColor = &HFF&
                ElseIf Right(.Text, 1) = "⊕" Then
                    .TextMatrix(i, 9) = .TextMatrix(i, 9) & PUB_GetRelateCasePropertyName(.TextMatrix(i, 8), "1")
                    For j = 0 To .Cols - 1
                        .col = j
                        .CellBackColor = &H8080FF
                    Next j
                End If
            Next i
            
        End If
    End With
    
    '查詢結果僅有一筆資料, 則直接勾選
    If Me.grdDataList.Rows = 2 Then
        grdDataList.col = 1
        grdDataList.row = 1
        If grdDataList.Text <> "" Then
            grdDataList.row = 1
            grdDataList.col = 0
            grdDataList.Text = "V"
            For i = 0 To grdDataList.Cols - 1
                If i <> 0 And (i = 2 And Right(grdDataList.TextMatrix(1, 1), 1) = "⊕") = False Then
                    grdDataList.col = i
                    grdDataList.CellBackColor = &HFFC0C0
                End If
            Next i
        End If
    End If
    grdDataList.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub GrdDataList_Click()
    grdDataList.row = grdDataList.MouseRow
    grdDataList.col = 0
    grdDataList.Visible = False
    If grdDataList.row <> 0 Then
        If grdDataList.Text = "V" Then
            grdDataList.Text = ""
            grdDataList.col = 1
            If Right(grdDataList.Text, 1) = "⊕" Then
                For i = 0 To grdDataList.Cols - 1
                    grdDataList.col = i
                    grdDataList.CellBackColor = &H8080FF
                Next i
            Else
                For i = 0 To grdDataList.Cols - 1
                    If i <> 1 Then
                        grdDataList.col = i
                        grdDataList.CellBackColor = QBColor(15)
                    End If
                Next i
            End If
        Else
            grdDataList.Text = "V"
            For i = 0 To grdDataList.Cols - 1
                If i <> 1 And (i = 2 And Right(grdDataList.TextMatrix(grdDataList.MouseRow, 1), 1) = "⊕") = False Then
                    grdDataList.col = i
                    grdDataList.CellBackColor = &HFFC0C0
                End If
            Next i
        End If
    End If
    grdDataList.Visible = True
End Sub

Private Sub grdDataList_DblClick()
    Dim strBackVal As String
    
    grdDataList.row = grdDataList.MouseRow
    grdDataList.col = 1
    If Me.Tag = MsgText(601) Then Exit Sub
    
    If grdDataList.row <> 0 Then
        strBackVal = Pub_RplStr(grdDataList.Text)
        Select Case UCase(Me.Tag)
            Case "FRMACC2210"
                Frmacc2210.Text1.Tag = strBackVal
                Frmacc2210.Text1 = strBackVal
        End Select
        Unload Me
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 1 '名稱
            txtFM2(0).SetFocus
            txtFM2_GotFocus (0)
        Case 2 'E-mail
            txt1(9).SetFocus
            txt1_GotFocus (9)
        Case 3 '國籍
            txt1(10).SetFocus
            txt1_GotFocus (10)
    End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    If Index = 1 Then
        If Left(Pub_StrUserSt03, 1) = "F" Then
            CloseIme
        Else
            OpenIme
        End If
    Else
        CloseIme
    End If
   
    txt1(Index).SelStart = 0
    txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 1 And Index <> 10 Then
        KeyAscii = UpperCase(KeyAscii)
    End If
End Sub

Private Sub txt1_LostFocus(Index As Integer)
    Select Case Index
      Case 9
            If Len(txt1(9)) <> 0 Then
                strSql = "SELECT NA03 FROM NATION WHERE NA01='" & txt1(9) & "'"
                CheckOC
                adoRecordset.CursorLocation = adUseClient
                adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                    If Not IsNull(adoRecordset.Fields(0)) Then
                        Lbl1(1).Caption = adoRecordset.Fields(0)
                    Else
                        Lbl1(1).Caption = ""
                    End If
                Else
                    Lbl1(1).Caption = ""
                    s = MsgBox("國家輸入錯誤！", , "錯誤！")
                    txt1(Index).SetFocus
                    txt1_GotFocus (Index)
                    Exit Sub
                End If
                CheckOC
            Else
                Lbl1(1).Caption = ""
            End If
      Case Else
   End Select
End Sub

Private Sub txt1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   
   Select Case Index
      'Mark by Lydia 2022/01/06
'      Case 1
'          Option1(1).Value = True
      'end 2022/01/06
      Case 9
          Option1(2).Value = True
      Case 10
          Option1(3).Value = True
      Case Else
   End Select
End Sub

Private Sub GetPleft()
    ReDim PLeft(0 To 7)
    ReDim ColName(1 To 7)
    PLeft(0) = 100
    PLeft(1) = PLeft(0) + 2000: ColName(1) = "本所案號"
    PLeft(2) = PLeft(1) + 2700: ColName(2) = "    名       稱    "
    PLeft(3) = PLeft(2) + 1200: ColName(3) = "智權人員"
    PLeft(4) = PLeft(3) + 1500: ColName(4) = " 狀  態 "
    PLeft(5) = PLeft(4) + 1300: ColName(5) = "總收文號"
    PLeft(6) = PLeft(5) + 1800: ColName(6) = "案件性質"
    PLeft(7) = PLeft(6) + 1200: ColName(7) = "收文日"
End Sub

Sub StrMenu(StrToGrid)
    strSql = "SELECT FA01||FA02||Decode(FA02,'0','','＊'),SUBSTR(DECODE(FA10,'013',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),'020',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)),1,80),NA03 FROM FAGENT,NATION WHERE FA01>='" & Left(StrToGrid, 6) & "00' AND FA01<='" & Left(StrToGrid, 6) & "zz' AND fa10=NA01(+) "
    strSql = strSql & " union all SELECT cu01||cu02||Decode(CU02,'0','','＊'),SUBSTR(DECODE(cu10,'013',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),'020',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),DECODE(cu05,NULL,NVL(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90)),1,80),NA03 FROM customer,NATION WHERE cu01>='" & Left(StrToGrid, 6) & "00' AND cu01<='" & Left(StrToGrid, 6) & "zz' AND cu10=NA01(+) "
    strSql = strSql & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','＊'),NVL(PCU08,DECODE(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)),NA03 FROM PotCustomer,Nation WHERE PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz'   AND NA01(+)=PCU09"
    strSql = strSql & " union  SELECT POC01||POC02||Decode(POC02,'0','','＊'),POC03,NA03 FROM PotCustomer1,Nation WHERE POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz'   AND NA01(+)=POC04"
    '傳入R1時找出相關的X
    strSql = strSql & " union  SELECT cu01||cu02||Decode(CU02,'0','','＊'),SUBSTR(DECODE(cu10,'013',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),'020',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),DECODE(cu05,NULL,NVL(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90)),1,80),NA03 " & _
                                                    "From CUSTOMER, PotCustomer1, Nation " & _
                                               "WHERE CU10=NA01(+) " & _
                                                    "AND POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND CU01>=(substr(POC16,1,6)||'00') AND CU01<=(substr(POC16,1,6)||'zz') " & _
                                                    "AND POC16 is not null "
    '找出R1的關係企業
    strSql = strSql & " union  SELECT POC01||POC02||Decode(POC02,'0','','＊'),POC03,NA03 " & _
                                                    "From PotCustomer1, Nation " & _
                                                "WHERE NA01(+)=POC04 " & _
                                                     "AND POC16>='" & Left(StrToGrid, 6) & "00' AND POC16<='" & Left(StrToGrid, 6) & "zz' " & _
                                                     "AND POC16 is not null "
    '傳入R1時找出相關的R
    strSql = strSql & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','＊'),NVL(PCU08,DECODE(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)),NA03 " & _
                                                    "From PotCustomer, Nation, PotCustomer1 " & _
                                               "WHERE NA01(+)=PCU09 " & _
                                                    "AND POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND PCU47>=(substr(POC16,1,6)||'00') AND PCU47<=(substr(POC16,1,6)||'zz') " & _
                                                    "AND POC16 is not null AND PCU47 is not null "
    '傳入R時找出相關的X
    strSql = strSql & " union  SELECT cu01||cu02||Decode(CU02,'0','','＊'),SUBSTR(DECODE(cu10,'013',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),'020',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),DECODE(cu05,NULL,NVL(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90)),1,80),NA03 " & _
                                                    "From CUSTOMER, PotCustomer, Nation " & _
                                               "WHERE CU10=NA01(+) " & _
                                                    "AND PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND CU01>=(substr(PCU47,1,6)||'00') AND CU01<=(substr(PCU47,1,6)||'zz') " & _
                                                    "AND PCU47 is not null "
    '傳入R時找出相關的Y
    strSql = strSql & " union  SELECT FA01||FA02||Decode(FA02,'0','','＊'),NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),NA03 " & _
                                                    "From Fagent, PotCustomer, Nation " & _
                                                "WHERE NA01(+)=FA10 " & _
                                                     "AND PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                     "AND FA01>=(substr(PCU47,1,6)||'00') AND FA01<=(substr(PCU47,1,6)||'zz') " & _
                                                     "AND PCU47 is not null "
    '找出R的關係企業
    strSql = strSql & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','＊'),NVL(PCU08,DECODE(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)),NA03 " & _
                                                    "From PotCustomer, Nation " & _
                                               "WHERE NA01(+)=PCU09 " & _
                                                    "AND PCU47>='" & Left(StrToGrid, 6) & "00' AND PCU47<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND PCU47 is not null "
    '傳入R時找出相關的R1
    strSql = strSql & " union  SELECT POC01||POC02||Decode(POC02,'0','','＊'),POC03,NA03 " & _
                                                    "From PotCustomer1, Nation, PotCustomer " & _
                                               "WHERE NA01(+)=POC04 " & _
                                                    "AND PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND POC16>=(substr(PCU47,1,6)||'00') AND POC16<=(substr(PCU47,1,6)||'zz') " & _
                                                    "AND PCU47 is not null AND POC16 is not null "
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 Then
        adoRecordset.MoveFirst
        Do While adoRecordset.EOF = False
        
            strSql = "INSERT INTO R100114 values ('"
            If Not IsNull(adoRecordset.Fields(0)) Then
                strSql = strSql + ChgSQL(adoRecordset.Fields(0)) + "','"
            Else
                strSql = strSql + "','"
            End If
            If Not IsNull(adoRecordset.Fields(1)) Then
                strSql = strSql & ChgSQL(adoRecordset.Fields(1)) + "','"
            Else
                strSql = strSql + "','"
            End If
            If Not IsNull(adoRecordset.Fields(2)) Then
                strSql = strSql + ChgSQL(adoRecordset.Fields(2)) + "','" & strUserNum & "')"
            Else
                strSql = strSql + "','" & strUserNum & "')"
            End If
            cnnConnection.Execute strSql
            adoRecordset.MoveNext
        Loop
    Else
    End If
    CheckOC
End Sub

Sub StrMenu1()
    Screen.MousePointer = vbHourglass
    'Added by Lydia 2017/12/05 改由啟用日控制
    If strSrvDate(1) < 國外部關聯企業啟用日 Then
        strSql = "SELECT '' AS V,R07001||decode(cu111,'Y','$','') AS 編號,R07002 AS 名稱,R07003 AS 國籍,ST02 as 智權人員,CU80 AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM R100114,CUSTOMER,Staff where id='" & strUserNum & "' And SUBSTR(R07001,1,1)='X' AND SUBSTR(R07001,1,8)=CU01(+) AND SUBSTR(R07001,9,1)=CU02(+) AND CU13=ST01(+)"
        strSql = strSql & "UNION ALL SELECT '' AS V,R07001||decode(fa77,'Y','$','') AS 編號,R07002 AS 名稱,R07003 AS 國籍,'' as 智權人員,FA69 AS 狀態,FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM R100114,FAGENT where id='" & strUserNum & "' And SUBSTR(R07001,1,1)='Y' AND SUBSTR(R07001,1,8)=FA01(+) AND SUBSTR(R07001,9,1)=FA02(+)"
        strSql = strSql & "UNION ALL SELECT '' AS V,R07001 AS 編號,R07002 AS 名稱,R07003 AS 國籍,ST02 as 智權人員,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM R100114,POTCUSTOMER,Staff where id='" & strUserNum & "' AND SUBSTR(R07001,1,1)='R' AND SUBSTR(R07001,1,8)=PCU01 AND SUBSTR(R07001,9,1)=PCU02 and substr(LTrim(PCU38),1,5)=ST01(+) "
        strSql = strSql & "UNION ALL SELECT '' AS V,R07001 AS 編號,R07002 AS 名稱,R07003 AS 國籍,ST02 as 智權人員,POC14 AS 狀態,POC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM R100114,POTCUSTOMER1,Staff where id='" & strUserNum & "' AND SUBSTR(R07001,1,1)='R' AND SUBSTR(R07001,1,8)=POC01 AND SUBSTR(R07001,9,1)=POC02 and POC13=ST01(+) "
        strSql = strSql & "ORDER BY 編號"
    Else
        'Added by Lydia 2017/02/14 抓關聯企業改成模組,暫存R100114_1
        strSql = "SELECT '' AS V,R11402 AS 編號,R11403 AS 名稱,NVL(NA03,R11405) AS 國籍 ,ST02 AS 智權人員,R11407 AS 狀態,R11408 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日," & _
               "R11409 AS 關聯編號,DECODE(SUBSTR(R11409,1,1),'X',DECODE(SIGN(INSTR('000,001,002,003,004,005,006,007,008,009,013,020',C1.CU10)),0,DECODE(C1.CU05,NULL,NVL(C1.CU04,C1.CU06),C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90),NVL(C1.CU04,DECODE(C1.CU05,NULL,C1.CU06,C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)))," & _
               "'Y',DECODE(SIGN(INSTR('000,001,002,003,004,005,006,007,008,009,013,020',F1.FA10)),0,DECODE(F1.FA05,NULL,NVL(F1.FA04,F1.FA06),F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65),NVL(F1.FA04,DECODE(F1.FA05,NULL,F1.FA06,F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65))),R11409) AS 關聯名稱," & _
               "R11410 AS 關聯關係, R11411 AS 關聯說明 FROM R100114_1,STAFF,NATION,CUSTOMER C1,FAGENT F1 " & _
               "WHERE ID='" & strUserNum & "' AND FORMID='" & UCase(Me.Name) & "' AND R11406=ST01(+) AND R11405=NA01(+) " & _
               "AND SUBSTR(R11409,1,8)=C1.CU01(+) AND '0'=C1.CU02(+) AND SUBSTR(R11409,1,8)=F1.FA01(+) AND '0'=F1.FA02(+) "
        strSql = strSql & "ORDER BY R11401,R11402,R11409"
        'end 2017/02/14
    End If
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly

    If adoRecordset.RecordCount <> 0 Then
        Set grdDataList.Recordset = adoRecordset
        'Modified by Lydia 2017/12/05 改由啟用日控制
        If strSrvDate(1) >= 國外部關聯企業啟用日 Then
            'Added by Lydia 2017/02/14 欄寬調整
            grdDataList.FixedCols = 3 '固定編號和名稱
            Call PUB_SetMSFGridColor(Me.grdDataList, "15") '底色設定為空白
            grdDataList.ColWidth(2) = 1200 '名稱
            grdDataList.ColWidth(3) = 800 '國籍
            grdDataList.ColWidth(6) = 1200 '備註
            grdDataList.ColWidth(11) = 1000 '關聯編號
            grdDataList.ColWidth(12) = 1200 '關聯名稱
            grdDataList.ColWidth(13) = 1200 '關聯關係
            grdDataList.ColWidth(14) = 1200 '關聯說明
            'end 2017/02/14
        End If
        'end 2017/12/05
    End If
    CheckOC
    'SetDataListWidth 'Remove by Lydia 2017/02/14
    If Me.grdDataList.Rows = 2 Then
        grdDataList.row = 1
        grdDataList.col = 1
        If grdDataList.Text <> "" Then
            grdDataList.Visible = False
            grdDataList.row = 1
            grdDataList.col = 0
            grdDataList.Text = "V"
            For i = 0 To grdDataList.Cols - 1
                grdDataList.col = i
                grdDataList.CellBackColor = &HFFC0C0
            Next i
            grdDataList.Visible = True
      End If
    End If
    Screen.MousePointer = vbDefault
End Sub

'Added by Lydia 2021/01/06
Private Sub txtFM2_GotFocus(Index As Integer)
    TextInverse txtFM2(Index)
End Sub

Private Sub txtFM2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtFM2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 0 Then
        Option1(1).Value = True
    End If
End Sub


'end 2022/01/6

