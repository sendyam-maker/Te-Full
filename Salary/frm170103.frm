VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm170103 
   BorderStyle     =   1  '單線固定
   Caption         =   "勞健保保費重新計算"
   ClientHeight    =   1572
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3672
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1572
   ScaleWidth      =   3672
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   930
      Width           =   3465
      _ExtentX        =   6117
      _ExtentY        =   508
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "結束(&X)"
      Height          =   405
      Left            =   2460
      TabIndex        =   3
      Top             =   30
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Left            =   1470
      TabIndex        =   2
      Top             =   30
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1140
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "Y"
      Top             =   540
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "( 0/0 )"
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   1260
      Width           =   3435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否確定：        (Y/N)"
      Height          =   180
      Left            =   270
      TabIndex        =   1
      Top             =   600
      Width           =   1665
   End
End
Attribute VB_Name = "frm170103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/1/26 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by Morgan 2009/1/9
Option Explicit


Private Sub Command1_Click()
   Screen.MousePointer = vbHourglass
   If Text1 = "Y" Then
      Me.Enabled = False
      Recalculate
      Text1 = ""
      Me.Enabled = True
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Text1 = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170103 = Nothing
End Sub

Private Sub Recalculate()
   Dim iR(1 To 15) As String '勞保勞退健保費率資料

   Dim stSQL As String
   Dim adoRst As ADODB.Recordset
   Dim lngSD14 As Long, lngSD15 As Long
   Dim lngInsureSalary As Long '投保薪資
   Dim lngInsureBase As Long '投保金額
   Dim lngInsureBaseH As Long '健保投保金額
   Dim dblInsureRate As Double '投保費率
   Dim dblFreeRate As Double '補助比率
   Dim dblInsureRate2 As Double '就業保險費率
   Dim intShareRate As Integer '負擔比例
   
On Error GoTo ErrHnd1
   
   stSQL = "select * from InsuranceRate"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      For intI = 1 To 15
         iR(intI) = "" & RsTemp.Fields("IR" & Format(intI, "00"))
      Next
   Else
      MsgBox "無法讀取費率資料檔!", vbExclamation
      GoTo XPortal
   End If
   '兼職的也可能會有勞健保
   'Modified by Morgan 2015/9/14 +留職停薪
   'stSQL = "select * from staff,salarydata where st01<'F' and (st51 is null or st51>=" & Left(strSrvDate(1), 6) & "01) and sd01(+)=st01 and sd01 is not null"
   stSQL = "select b.* from staff b where st01<'F' and (st51 is null or st51>=" & Left(strSrvDate(1), 6) & "01)"
   stSQL = stSQL & "union select b.* from (select sc01,max(sc02||sc03) from staff_change group by sc01 having substr(max(sc02||sc03),-2)='04') a,staff b where st01(+)=sc01"
   stSQL = "select * from (" & stSQL & ") x,salarydata where sd01(+)=st01 and sd01 is not null"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      With adoRst
      cnnConnection.BeginTrans
      
On Error GoTo ErrHnd

      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      
      Do While Not .EOF
         '勞保投保薪資
         'Modified by Morgan 2016/3/31 特殊勞保投保薪資會輸0(63001,已退休)
         'lngInsureSalary = Val("" & .Fields("sd12"))
         'If lngInsureSalary = 0 Then
         If Not IsNull(.Fields("sd12")) Then
            lngInsureSalary = Val("" & .Fields("sd12"))
         Else
         'end 2016/3/31
            lngInsureSalary = Val("" & .Fields("sd45"))
         End If
         '勞保等級
         lngInsureBase = PUB_GetInsureBase(lngInsureSalary, "L")
         '勞保費=勞保等級*勞保費率*勞保個人負擔比例
         'Modify by Morgan 2009/6/29
         '98/5/1 起外國人也有失業給付,費率改與本國人同,只有雇主(所長)沒有(未來加65歲以上)
         'If Left(.Fields("st24"), 1) = "F" Then
         'Modify by Morgan 2010/10/26 勞保費率及就業保險費率需個別計算(四捨五入)
         'Modified by Morgan 2013/1/21 改判斷 sd11 勞健保是否以合夥人身分投保
         'If .Fields("st20") = "11" Then
         'Modified by Morgan 2023/6/29 +判斷 sd48勞保是否無就保
         If .Fields("sd11") = "Y" Or .Fields("sd48") = "Y" Then
         'end 2013/1/21
            'dblInsureRate = Val(IR(2))
            dblInsureRate = Val(iR(1))
            dblInsureRate2 = 0 'Add
         
                  'Added by Morgan 2015/1/28
         '超過65歲也沒有就保
         ElseIf PUB_ChkOver65(.Fields("st01")) Then
            dblInsureRate = Val(iR(1))
            dblInsureRate2 = 0
         'end 2015/1/28
         
         Else
            dblInsureRate = Val(iR(1))
            dblInsureRate2 = Val(iR(2)) 'Add
         End If
         
         'lngSD14 = Round(lngInsureBase * dblInsureRate / 100 * Val(IR(3)) / 100, 0)
         lngSD14 = Round(lngInsureBase * dblInsureRate / 100 * Val(iR(3)) / 100, 0) + Round(lngInsureBase * dblInsureRate2 / 100 * Val(iR(3)) / 100, 0)
         'end 2010/10/26
         
         '健保投保薪資
         'Modified by Morgan 2016/3/31 與勞保檢查一致
         'lngInsureSalary = Val("" & .Fields("sd13"))
         'If lngInsureSalary = 0 Then
         If Not IsNull(.Fields("sd13")) Then
            lngInsureSalary = Val("" & .Fields("sd13"))
         Else
         'end 2016/3/31
            lngInsureSalary = Val("" & .Fields("sd45"))
         End If
         dblInsureRate = Val(iR(6))
         
         'Added by Morgan 2013/1/21
         '以合夥人身分投保 100% 個人負擔
         'Modified by Morgan 2020/12/3 排除68099(掛名負責人，只需負擔個人保費)
         If .Fields("sd11") = "Y" And .Fields("sd01") <> "68099" Then
            intShareRate = 100
         Else
            intShareRate = Val(iR(7))
         End If
         'end 2013/1/21
            
         '健保費=健保等級*健保費率*健保個人負擔比例
         'Modify by Morgan 2010/4/15 健保費調整改用共用函數
         'lngInsureBase = PUB_GetInsureBase(lngInsureSalary, "H") '健保等級
         'lngSD15 = Round(lngInsureBase * dblInsureRate / 100 * Val(IR(7)) / 100, 0)
         lngInsureBase = PUB_GetInsureBase(lngInsureSalary, "H", dblFreeRate) '健保等級
         'Modified by Morgan 2013/1/21
         'lngSD15 = PUB_GetHIFee(lngInsureBase, dblInsureRate, Val(IR(7)), dblFreeRate)
         lngSD15 = PUB_GetHIFee(lngInsureBase, dblInsureRate, intShareRate, dblFreeRate)
         lngInsureBaseH = lngInsureBase
         'end 2013/1/21
         'end 2010/4/15
         'Modified by Morgan 2013/1/21 +SD47
         stSQL = "update salarydata set sd14=" & lngSD14 & ",sd15=" & lngSD15 & ",SD47=" & lngInsureBaseH & " where sd01='" & .Fields("sd01") & "'"
         cnnConnection.Execute stSQL, intI
         
         ProgressBar1.Value = ProgressBar1.Value + 1
         Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         .MoveNext
      Loop
      End With
      'Add by Morgan 2009/6/24
      '重算勞健保未更新的薪資異動資料
      stSQL = "select * from SalaryUpdate,SalaryLog,staff,salarydata where su05 is null and su03='3' and sl01(+)=su01 and sl02(+)=su02 and st01(+)=sl01 and sd01(+)=sl01"
      intI = 1
      Set adoRst = ClsLawReadRstMsg(intI, stSQL)
      If intI = 1 Then
         With adoRst
         .MoveFirst
         Do While Not .EOF
            '勞保投保薪資
            'Modified by Morgan 2016/3/31 特殊勞保投保薪資會輸0(63001,已退休)
            'lngInsureSalary = Val("" & .Fields("sl07"))
            'If lngInsureSalary = 0 Then
            If Not IsNull(.Fields("sl07")) Then
               lngInsureSalary = Val("" & .Fields("sl07"))
            Else
            'end 2016/3/31
               'Modify By Sindy 2020/6/24 + SL39 證照津貼
               lngInsureSalary = Val("" & .Fields("sl11")) + Val("" & .Fields("sl12")) + _
                                 Val("" & .Fields("SL39")) + Val("" & .Fields("sl14"))
            End If
            '勞保等級
            lngInsureBase = PUB_GetInsureBase(lngInsureSalary, "L")
            
            '98/5/1 起外國人也有失業給付,費率改與本國人同,只有雇主(所長)沒有(未來加65歲以上)
            'Modify by Morgan 2010/10/26 勞保費率及就業保險費率需個別計算(四捨五入)
            'Modified by Morgan 2013/1/21 改判斷 sd11 勞健保是否以合夥人身分投保
            'If .Fields("st20") = "11" Then
            'Modified by Morgan 2023/6/29 +判斷 sd48勞保是否無就保
            If .Fields("sd11") = "Y" Or .Fields("sd48") = "Y" Then
            'end 2013/1/21
               'dblInsureRate = Val(IR(2))
               dblInsureRate = Val(iR(1))
               dblInsureRate2 = 0 'Add
               
            'Added by Morgan 2015/1/28
            '超過65歲也沒有就保
            ElseIf PUB_ChkOver65(.Fields("sl01")) Then
               dblInsureRate = Val(iR(1))
               dblInsureRate2 = 0
            'end 2015/1/28
         
            Else
               dblInsureRate = Val(iR(1))
               dblInsureRate = Val(iR(1))
               dblInsureRate2 = Val(iR(2)) 'Add
            End If
            '勞保費=勞保等級*勞保費率*勞保個人負擔比例
            'lngSD14 = Round(lngInsureBase * dblInsureRate / 100 * Val(IR(3)) / 100, 0)
            lngSD14 = Round(lngInsureBase * dblInsureRate / 100 * Val(iR(3)) / 100, 0) + Round(lngInsureBase * dblInsureRate2 / 100 * Val(iR(3)) / 100, 0)
            'end 2010/10/26
            
            '健保投保薪資
            'Modified by Morgan 2016/3/31 與勞保檢查一致
            'lngInsureSalary = Val("" & .Fields("sl08"))
            'If lngInsureSalary = 0 Then
            If Not IsNull(.Fields("sl08")) Then
               lngInsureSalary = Val("" & .Fields("sl08"))
            Else
            'end 2016/3/31
               'Modify By Sindy 2020/6/24 + SL39 證照津貼
               lngInsureSalary = Val("" & .Fields("sl11")) + Val("" & .Fields("sl12")) + _
                                 Val("" & .Fields("SL39")) + Val("" & .Fields("sl14"))
            End If
            dblInsureRate = Val(iR(6))
            
            'Added by Morgan 2013/1/21
            '以合夥人身分投保 100% 個人負擔
            If .Fields("sd11") = "Y" Then
               intShareRate = 100
            Else
               intShareRate = Val(iR(7))
            End If
            'end 2013/1/21
         
            '健保費=健保等級*健保費率*健保個人負擔比例
            'Modify by Morgan 2010/4/15 健保費調整改用共用函數
            'lngInsureBase = PUB_GetInsureBase(lngInsureSalary, "H") '健保等級
            'lngSD15 = Round(lngInsureBase * dblInsureRate / 100 * Val(IR(7)) / 100, 0)
            lngInsureBase = PUB_GetInsureBase(lngInsureSalary, "H", dblFreeRate) '健保等級
            
            'Modified by Morgan 2013/1/21
            'lngSD15 = PUB_GetHIFee(lngInsureBase, dblInsureRate, Val(IR(7)), dblFreeRate)
            lngSD15 = PUB_GetHIFee(lngInsureBase, dblInsureRate, intShareRate, dblFreeRate)
            lngInsureBaseH = lngInsureBase
            'end 2010/4/15
            'Modified by Morgan 2013/1/21 +SL38
            stSQL = "update salarylog set sl09=" & lngSD14 & ",sl10=" & lngSD15 & ",sl38=" & lngInsureBaseH & " where sl01='" & .Fields("sl01") & "' and sl02=" & .Fields("sl02")
            cnnConnection.Execute stSQL, intI
         
            .MoveNext
         Loop
         End With
      End If
      '2009/6/24
      
      cnnConnection.CommitTrans
      MsgBox "重新計算完畢！", vbInformation
   Else
      MsgBox "無資料可重新計算！", vbExclamation
   End If
   GoTo XPortal
   
ErrHnd:
   cnnConnection.RollbackTrans
   
ErrHnd1:
   MsgBox Err.Description, vbCritical
   
XPortal:
   Set adoRst = Nothing
   
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub
