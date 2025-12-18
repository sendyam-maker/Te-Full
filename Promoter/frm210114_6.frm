VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210114_6 
   BorderStyle     =   1  '單線固定
   Caption         =   "條碼案件委任契約書"
   ClientHeight    =   4560
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9432
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   9432
   Begin VB.CheckBox ChkSeal 
      Caption         =   "用印"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   5250
      TabIndex        =   43
      Top             =   4170
      Width           =   735
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00C0FFFF&
      Caption         =   "空白列印"
      Height          =   330
      Index           =   5
      Left            =   3720
      Style           =   1  '圖片外觀
      TabIndex        =   42
      Top             =   30
      Width           =   920
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm210114_6.frx":0000
      Left            =   6870
      List            =   "frm210114_6.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   16
      Top             =   4170
      Width           =   2475
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "搜尋委託人(&Q)"
      Height          =   330
      Left            =   7350
      TabIndex        =   40
      Top             =   2550
      Width           =   1365
   End
   Begin VB.CheckBox chk1 
      Caption         =   "申請廠商號碼"
      Height          =   225
      Index           =   3
      Left            =   1020
      TabIndex        =   3
      Top             =   1410
      Width           =   255
   End
   Begin VB.CheckBox chk1 
      Caption         =   "正片測試及陳報商品基本資料明細"
      Height          =   225
      Index           =   2
      Left            =   1020
      TabIndex        =   2
      Top             =   1170
      Width           =   3405
   End
   Begin VB.CheckBox chk1 
      Caption         =   "製作正片"
      Height          =   225
      Index           =   1
      Left            =   1020
      TabIndex        =   1
      Top             =   930
      Width           =   1545
   End
   Begin VB.CheckBox chk1 
      Caption         =   "申請廠商號碼"
      Height          =   225
      Index           =   0
      Left            =   1020
      TabIndex        =   0
      Top             =   690
      Width           =   1545
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "讀取文檔"
      Height          =   330
      Index           =   4
      Left            =   5544
      TabIndex        =   19
      Top             =   30
      Width           =   920
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "儲存文檔"
      Height          =   330
      Index           =   3
      Left            =   4632
      TabIndex        =   18
      Top             =   30
      Width           =   920
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   360
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   3645
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   660
         Style           =   2  '單純下拉式
         TabIndex        =   17
         Top             =   30
         Width           =   2940
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "印表機"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   25
         Top             =   90
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "清空資料"
      Height          =   330
      Index           =   2
      Left            =   6456
      TabIndex        =   20
      Top             =   30
      Width           =   920
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面"
      Height          =   330
      Index           =   1
      Left            =   8460
      TabIndex        =   23
      Top             =   30
      Width           =   920
   End
   Begin VB.TextBox txtPCnt 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Left            =   7890
      MaxLength       =   1
      TabIndex        =   21
      Text            =   "2"
      Top             =   60
      Width           =   270
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   9420
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印       份"
      Height          =   330
      Index           =   0
      Left            =   7368
      TabIndex        =   22
      Top             =   30
      Width           =   1100
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1260
      TabIndex        =   4
      Top             =   1410
      Width           =   7995
      VariousPropertyBits=   671105051
      MaxLength       =   54
      Size            =   "14102;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1590
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "2355;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1950
      TabIndex        =   7
      Top             =   2239
      Width           =   1410
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "2487;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1800
      TabIndex        =   8
      Top             =   2558
      Width           =   5475
      VariousPropertyBits=   671105051
      MaxLength       =   48
      Size            =   "9657;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1800
      TabIndex        =   9
      Top             =   2877
      Width           =   5475
      VariousPropertyBits=   671105051
      MaxLength       =   48
      Size            =   "9657;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   1800
      TabIndex        =   10
      Top             =   3196
      Width           =   5475
      VariousPropertyBits=   671105051
      MaxLength       =   48
      Size            =   "9657;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   1800
      TabIndex        =   11
      Top             =   3515
      Width           =   5490
      VariousPropertyBits=   671105051
      MaxLength       =   22
      Size            =   "9684;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   1800
      TabIndex        =   12
      Top             =   3834
      Width           =   5475
      VariousPropertyBits=   671105051
      MaxLength       =   48
      Size            =   "9657;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   9
      Left            =   1800
      TabIndex        =   13
      Top             =   4155
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1244;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   10
      Left            =   2850
      TabIndex        =   14
      Top             =   4155
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1235;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   11
      Left            =   3930
      TabIndex        =   15
      Top             =   4155
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1235;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   4980
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "2355;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "受任人："
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   6120
      TabIndex        =   41
      Top             =   4200
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "元整。"
      Height          =   180
      Left            =   3390
      TabIndex        =   39
      Top             =   2284
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "委辦範圍："
      Height          =   180
      Left            =   135
      TabIndex        =   38
      Top             =   450
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "共計新台幣"
      Height          =   180
      Index           =   0
      Left            =   1020
      TabIndex        =   37
      Top             =   2284
      Width           =   900
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "甲方：委任人："
      Height          =   180
      Left            =   435
      TabIndex        =   36
      Top             =   2603
      Width           =   1260
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "代表人："
      Height          =   180
      Left            =   975
      TabIndex        =   35
      Top             =   2922
      Width           =   720
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "地　址："
      Height          =   180
      Left            =   975
      TabIndex        =   34
      Top             =   3241
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "電　話："
      Height          =   180
      Left            =   975
      TabIndex        =   33
      Top             =   3560
      Width           =   720
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "乙方：經手人："
      Height          =   180
      Left            =   435
      TabIndex        =   32
      Top             =   3879
      Width           =   1260
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "中　華　民　國　　　　　年　　　　　月　　　　　日"
      Height          =   180
      Left            =   420
      TabIndex        =   31
      Top             =   4200
      Width           =   4500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "委辦費用："
      Height          =   180
      Left            =   120
      TabIndex        =   30
      Top             =   1740
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "新台幣"
      Height          =   180
      Left            =   1020
      TabIndex        =   29
      Top             =   1965
      Width           =   540
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "元整，另代收規費新台幣"
      Height          =   180
      Left            =   2970
      TabIndex        =   28
      Top             =   1965
      Width           =   1980
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "元整，"
      Height          =   180
      Left            =   6330
      TabIndex        =   27
      Top             =   1965
      Width           =   540
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "請輸入數字"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3990
      TabIndex        =   26
      Top             =   2284
      Width           =   900
   End
End
Attribute VB_Name = "frm210114_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/24 改成Form2.0 ; txt1(index)、Printer改成Word列印
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'create by nickc 2007/11/14
Option Explicit

Dim SeekPrint As Integer, SeekPrintL As Integer
Dim iCount As Integer
'Add By Sindy 2010/4/30
Public m_strCustCode As String
Public m_blnOneRec As Boolean
'2010/4/30 End
Dim strNowCustNo As String 'Add by Amy 2016/08/19 客戶編號
Dim iPrintC As Integer 'Added by Lydia 2017/03/28 目前列印第幾份
Dim bolAddSeal As Boolean 'Added by Lydia 2017/03/28 是否用印
Dim d_Left As Double, d_Top As Double 'Added by Lydia 2017/04/25 印表機實際列印的左邊界、右邊界
Dim strPrinter As String 'Added by Lydia 2017/04/28
Dim strDetail As String 'Move by Lydia 2017/05/16 記錄內容(從StrMenu移出來)
Dim strCompSeal As String 'Added by Lydia 2020/03/25 記錄"公司名稱|用印編號",用,區隔
'Added by Lydia 2022/01/24  加入圖片用(Word)
Const msoFalse = 0
Const msoLineSolid = 1
Const msoLineSingle = 1
Const msoTrue = -1
Const msoPictureAutomatic = 1
'end 2022/01/24
Dim m_TempPDF As String 'Added by Lydia 2022/01/24
Dim m_TempFN As String 'Added by Lydia 2022/01/24

'Add By Sindy 2010/4/29
Private Sub cmdFind_Click()
Dim strCmpName As String, strMsg As String 'Add by Amy 2016/08/19

   If Me.txt1(4).Text = "" Then
      MsgBox "請輸入申請人中文名稱的關鍵字!!!", vbExclamation + vbOKOnly
      Me.txt1(4).SetFocus
      Exit Sub
   End If
   
   Set frm090801_1.m_frm0908A = Me   'Add by Lydia 2014/9/22
   frm090801_1.m_DouChk = False
   
   frm090801_1.m_strCustChnName = Me.txt1(4).Text
   frm090801_1.lblName.Caption = Me.txt1(4).Text
   m_blnOneRec = False
   m_strCustCode = ""
   If frm090801_1.StrMenu = True Then
      If frm090801_1.m_blnOneRec = False Then
         frm090801_1.Show vbModal
      End If
      m_blnOneRec = frm090801_1.m_blnOneRec
      m_strCustCode = frm090801_1.m_strCustCode
      Unload frm090801_1
   Else
      Unload frm090801_1
   End If
   Combo2.Tag = "": strNowCustNo = "" 'Add by Amy 2016/08/19
   If m_blnOneRec = True And m_strCustCode <> "" Then
     'Add by Amy 2016/08/19 記錄收據公司別(放於SetCustTxt前避免m_strCustCode被清空)
      strNowCustNo = m_strCustCode
      strCmpName = "Y"
      Combo2.Tag = GetReceiptCmp(Left(strNowCustNo, 8), Mid(strNowCustNo, 9, 1), "LA", "000", False, strCmpName, Me.Name)
      If Combo2.Tag <> MsgText(601) And Combo2 <> MsgText(601) And Combo2.Tag <> frm210114_1.GetComp(Combo2) Then
        strMsg = "您輸入之收據公司別「" & Combo2 & "」與客戶檔設定值「" & strCmpName & "」不同" & vbCrLf & _
                     "是否依客戶檔設定覆蓋您的輸入值？"
        If MsgBox(strMsg, vbYesNo + vbCritical) = vbYes Then
            'Modified by Lydia 2024/08/06
            'Combo2 = strCmpName
            Call Pub_SetCboListIdx(Me.Combo2, strCmpName)
        End If
      ElseIf strCmpName = MsgText(601) Then
        Combo2.ListIndex = 0
      Else
        'Modified by Lydia 2024/08/06
        'Combo2 = strCmpName
        Call Pub_SetCboListIdx(Me.Combo2, strCmpName)
      End If
      'end 2016/08/19
      'Modify by Amy 2021/05/13 +if 讀取文檔要保留原文檔內容
      If Me.ActiveControl.Name = "cmdFind" Then
        Call SetCustTxt(m_strCustCode)
      End If
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim tb As Control
Dim op As OptionButton
Dim fN As Integer
Dim strBuffer As String
Dim AllObj(0 To 17) As String
Dim AllObjV As Variant
'Add by Amy 2016/08/19 目前收據公司別
Dim strNowCmp As String

   Select Case Index
      Case 0
          If Chk1(0).Value = vbUnchecked And Chk1(1).Value = vbUnchecked And Chk1(2).Value = vbUnchecked And Chk1(3).Value = vbUnchecked Then
              MsgBox "委辦範圍至少勾一項！", vbInformation, "錯誤！"
              Chk1(0).SetFocus
              Exit Sub
          End If
          If Chk1(3).Value = vbChecked And txt1(0) = "" Then
              MsgBox "委辦範圍不可空白！", vbInformation, "錯誤！"
              txt1(0).SetFocus
              txt1_GotFocus 0
              Exit Sub
          End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(1)) = "" Then
              MsgBox "服務費不可空白！", vbInformation, "錯誤！"
              txt1(1).SetFocus
              txt1_GotFocus 1
              Exit Sub
          End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(2)) = "" Then
              MsgBox "規費不可空白，若沒規費請輸入 0！", vbInformation, "錯誤！"
              txt1(2).SetFocus
              txt1_GotFocus 2
              Exit Sub
          End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(3)) = "" Then
              MsgBox "費用不可空白！", vbInformation, "錯誤！"
              txt1(3).SetFocus
              txt1_GotFocus 3
              Exit Sub
          End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(8)) = "" Then
              MsgBox "經手人不可空白！", vbInformation, "錯誤！"
              txt1(8).SetFocus
              txt1_GotFocus 8
              Exit Sub
          End If
                
         '2011/10/18 ADD BY SONIA 檢查四縣市地址
         If txt1(6) <> "" Then
           If CheckTaiwanAddr(txt1(6), "000", "甲方委任人地址") = False Then
              txt1(6).SetFocus
              txt1_GotFocus (6)
              Exit Sub
           End If
         End If
         '2011/10/18 END
         'Add by Amy 2016/08/19 +受任人不可為空
         If Combo2 = "" Then
             MsgBox "受任人不可為空白！", vbInformation, "錯誤！"
             Combo2.SetFocus
             Exit Sub
         End If
          '2009/11/13 MODIFY BY SONIA 杜副總提出
      '    If txt1(9) = "" Or txt1(10) = "" Or txt1(11) = "" Then
      '        MsgBox "日期需要正確！", vbInformation, "錯誤！"
      '        txt1(9).SetFocus
      '        txt1_GotFocus 9
      '        Exit Sub
      '    End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(9)) = "" Or Trim(txt1(10)) = "" Or Trim(txt1(11)) = "" Then
             If MsgBox("契約書日期不完整，是否確定？", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
               txt1(9).SetFocus
               txt1_GotFocus 9
               Exit Sub
             End If
          End If
      '2009/11/13 END
          'Added by Lydia 2017/03/28
          If ChkSeal.Value = 1 Then
            'Modified by Lydia 2017/04/27 PDF印表機不需詢問,並且份數改為1份
            If InStr(UCase(Combo1.Text), "PDF") > 0 Then
                txtPCnt = "1"
            Else
               If MsgBox("用印的委任書需選擇彩色印表機，是否已選擇？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
               End If
            End If
            'end 2017/04/27
             bolAddSeal = True
          Else
             bolAddSeal = False
          End If
          'end 2017/03/28
          
          'Modified by Lydia 2017/04/13
'          For iCount = 1 To Val(txtPCnt)
'              Set Printer = Printers(Combo1.ListIndex)
'              Screen.MousePointer = vbHourglass
'              DoEvents
'              StrMenu
'          Next iCount
          'Modified by Lydia 2022/01/25 改成Word直接印
          'Call Print2PDF(False)
          Call runWordProc(False)
          PUB_SetOsDefaultPrinter strPrinter
          'end 2022/01/25
 
          'Add by Amy 2016/08/19 畫面與客戶檔收據公司別不同更新客戶檔
          strNowCmp = frm210114_1.GetComp(Combo2)
          If Combo2.Tag <> strNowCmp Then
             Call UpdReceiptCmp(strNowCustNo, strNowCmp)
          End If
          'end 2016/08/19
          Screen.MousePointer = vbDefault
          Call RunEndProc(True) 'Added by Lydia 2022/01/24 刪除暫存檔
          'Modified by Lydia 2022/01/24 判斷是否有列印
          'ShowPrintOk 'Added by Lydia 2017/04/11
          If m_TempPDF <> "" Then ShowPrintOk
      Case 1
          frm210114.Show
          Unload Me
      Case 2
          For Each tb In txt1
              tb.Text = Empty
          Next
          Chk1(0).Value = vbUnchecked
          Chk1(1).Value = vbUnchecked
          Chk1(2).Value = vbUnchecked
          Chk1(3).Value = vbUnchecked
      Case 3
          cd1.Filter = "Contract Files(*.Con)|*.Con"
          cd1.InitDir = GetMyDocPath
          On Error GoTo DialogCancel
          cd1.CancelError = True
          cd1.ShowSave
          If cd1.FileName <> "" Then
              AllObj(0) = "條碼案件委任契約書"
              For iCount = 1 To 12
                  AllObj(iCount) = txt1(iCount - 1).Text
              Next iCount
              AllObj(13) = Chk1(0).Value
              AllObj(14) = Chk1(1).Value
              AllObj(15) = Chk1(2).Value
              AllObj(16) = Chk1(3).Value
              AllObj(17) = Combo2.Text 'Add By Sindy 2011/3/23
              strBuffer = Join(AllObj, Chr(30))
              strBuffer = StrEncrypt(strBuffer)
              fN = FreeFile
              Open cd1.FileName For Output As fN
              Print #fN, strBuffer
              Close #fN
          End If
          'Add by Amy 2016/08/19 畫面與客戶檔收據公司別不同更新客戶檔
          strNowCmp = frm210114_1.GetComp(Combo2)
          If Combo2 <> MsgText(601) And Combo2.Tag <> strNowCmp Then
             Call UpdReceiptCmp(strNowCustNo, strNowCmp)
          End If
          'end 2016/08/19
      Case 4
          cd1.Filter = "Contract Files(*.Con)|*.Con"
          cd1.InitDir = GetMyDocPath
          On Error GoTo DialogCancel
          cd1.CancelError = True
          cd1.ShowOpen
          If cd1.FileName <> "" Then
              fN = FreeFile
              Open cd1.FileName For Input As fN
              Input #fN, strBuffer
              Close #fN
              strBuffer = StrDecrypt(strBuffer)
              AllObjV = Split(strBuffer, Chr(30))
              If AllObjV(0) = "條碼案件委任契約書" Then
                  cmdOK_Click 2
                  For iCount = 1 To 12
                       txt1(iCount - 1).Text = AllObjV(iCount)
                  Next iCount
                  Chk1(0).Value = AllObjV(13)
                  Chk1(1).Value = AllObjV(14)
                  Chk1(2).Value = AllObjV(15)
                  Chk1(3).Value = AllObjV(16)
                  'Modify by Amy 2016/08/19 避免空值會Error
                  If AllObjV(17) = MsgText(601) Then
                    Combo2.ListIndex = 0
                  Else
                    Combo2.Text = AllObjV(17) 'Add By Sindy 2011/3/23
                  End If
                  'end 2016/08/19
                  'Add By Sindy 2011/1/21 檢查地址欄
                  '委任人地址
                  If txt1(4).Text <> "" And txt1(6).Text <> "" Then
                     If CheckCustomerAddr(1, Trim(txt1(4).Text), Trim(txt1(6).Text), "委任人", True) = False Then
                        txt1(6).SetFocus
                     End If
                  End If
                  '2011/1/21 End
                  'Add by Amy 2016/08/19 讀取收據公司別
                  cmdFind_Click
              Else
                  MsgBox "錯誤格式，此份內容並非 條碼案件委任契約書 格式！", vbExclamation
              End If
          End If
      'Added by Lydia 2017/03/28 空白委任書
      Case 5
          If Trim(Combo2) = "" Then
             MsgBox "受任人不可為空白！", vbInformation, "錯誤！"
             Combo2.SetFocus
             Exit Sub
          End If
          'Modified by Lydia 2017/04/17 文雄表示用印由下方勾選,可直接空白列印
          If ChkSeal.Value = 1 Then
            If (InStr(UCase(Combo1.Text), "BATCH") > 0 Or InStr(UCase(Combo1.Text), "WRITER") > 0 Or InStr(UCase(Combo1.Text), "PDF") > 0) And Pub_StrUserSt03 <> "M51" Then
               MsgBox "空白用印的印表機不可選擇PDF列印！", vbInformation, "錯誤！"
               Combo1.SetFocus
               Exit Sub
            End If
            'Modified by Lydia 2017/04/27 PDF印表機不需詢問,並且份數改為1份
            If InStr(UCase(Combo1.Text), "PDF") > 0 Then
                txtPCnt = "1"
            Else
               If MsgBox("用印的委任書需選擇彩色印表機，是否已選擇？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
               End If
            End If
            'end 2017/04/27
            bolAddSeal = True
          End If
          'end 2017/04/17
          Call cmdOK_Click(2) '清空資料
          'Modified by Lydia 2022/01/25 改成Word直接印
          'Call Print2PDF(True)
          Call runWordProc(True)
          PUB_SetOsDefaultPrinter strPrinter
          'end 2022/01/25
          
          m_strCustCode = ""
          bolAddSeal = False
          Screen.MousePointer = vbDefault
          Call RunEndProc(True) 'Added by Lydia 2022/01/24 刪除暫存檔
          'Modified by Lydia 2022/01/24 判斷是否有列印
          'ShowPrintOk 'Added by Lydia 2017/04/11
          If m_TempPDF <> "" Then ShowPrintOk
      'end 2017/03/28
      Case Else
   End Select
   Exit Sub
DialogCancel:

End Sub

'Add by Morgan 2011/2/24 只要鍵盤有動作就不斷線
Private Sub Form_KeyPress(KeyAscii As Integer)
   If UCase(Forms(0).Name) = "MDIMAIN" Then Forms(0).tmrConnect.Tag = 0
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
   
   PUB_InitForm210114 Forms(0), Me 'Added by Lydia 2017/05/19 委任契約書表單大於主表單，控制主表單放大。
   MoveFormToCenter Me
   'Modified by Lydia 2017/04/28 改用模組
   'strSql = Printer.DeviceName
   'SeekPrintL = Printer.Orientation
   'For i = 0 To Printers.Count - 1
   '    Set Printer = Printers(i)
   '    Combo1.AddItem Printer.DeviceName, j
   '    j = j + 1
   '    If Printer.DeviceName = strSql Then
   '        SeekPrint = i
   '    End If
   'Next i
   'Set Printer = Printers(SeekPrint)
   'Combo1.Text = Combo1.List(SeekPrint)
   PUB_SetPrinter Me.Name, Me.Combo1, strPrinter, , , , , True 'Modified by Morgan 2020/10/30 +只顯示有效的印表機參數

    'Added by Lydia 2017/04/17 先用模組抓所有印表機後,排除特定印表機
    'Remove by Lydia 2017/06/07 改直接列印
    'For i = 0 To Combo1.ListCount - 1
    '   If InStr(UCase(Combo1.List(i)), "PDFCREATOR") > 0 And Trim(Combo1.List(i)) <> "" Then
    '      Combo1.RemoveItem i
    '      'If i = SeekPrint Then Combo1.Text = Combo1.List(0) 'Remove by Lydia 2017/04/28
    '   End If
    'Next
    'end 2017/04/17
    'end 2017/06/07
    
   'Modify by Amy 2016/08/19
   'Combo2.Text = Combo2.List(1) 'Add By Sindy 2011/3/23
   
   'Added by Lydia 2020/03/25 設定公司別下拉選項
   Call PUB_SetCboTofrm210114(Me.Name, Me.Combo2, strCompSeal)
   
   Combo2.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '還原預設印表機
   'Modified by Lydia 2017/04/28 記錄表單的印表機
   'Set Printer = Printers(SeekPrint)
   'Printer.Orientation = SeekPrintL
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   'end 2017/04/28
   
   Call RunEndProc(False) 'Added by Lydia 2022/01/24 刪除暫存檔
   Set frm210114_6 = Nothing
End Sub

'Modified by Lydia 2017/03/28
'Sub StrMenu()
Sub StrMenu(Optional ByVal bolSpace As Boolean = False)
Dim iY As Integer
Dim tmpI As Integer
Dim iStr(1 To 38) As String
Dim tBoxTop As Integer
'Added by Lydia 2017/03/28
Dim tObj As New StdPicture
Dim strSealFile As String '公司章圖檔
Dim strSpaceAmt As String
'end 2017/03/28

   iStr(1) = "條碼案件委任契約書"
   iStr(2) = "委任人(甲方)茲委任受任人(乙方)辦理條碼案件，雙方同意條件如下："
   iStr(3) = "第一條　委辦範圍："
   iStr(4) = "　　　　" & IIf(Chk1(0).Value = 1, "■", "□") & "　申請廠商號碼"
   iStr(5) = "　　　　" & IIf(Chk1(1).Value = 1, "■", "□") & "　製作正片"
   iStr(6) = "　　　　" & IIf(Chk1(2).Value = 1, "■", "□") & "　正片測試及陳報商品基本資料明細"
   iStr(7) = "　　　　" & IIf(Chk1(3).Value = 1, "■　" & StrToStr(txt1(0) & String(54, " "), 27), "□")
   iStr(8) = "第二條　委辦費用："
   strSpaceAmt = String(6, "　") '"　　拾　　萬　　仟　　佰"  'Added by Lydia 2017/03/28
   'Modified by Lydia 2017/03/28
   'iStr(9) = "　　　　新台幣　" & String(LenB(StrConv(ChangeNumber(txt1(1)), vbFromUnicode)), " ") & "　元整，另代收規費新台幣　" & String(LenB(StrConv(ChangeNumber(txt1(2)), vbFromUnicode)), " ") & "　元整，"
   'iStr(10) = "　　　　共計新台幣　" & String(LenB(StrConv(ChangeNumber(txt1(3)), vbFromUnicode)), " ") & "　元整。"
   iStr(9) = "　　　　新台幣　" & IIf(Val(Trim(txt1(1))) = 0, strSpaceAmt, String(LenB(StrConv(ChangeNumber(txt1(1)), vbFromUnicode)), " ")) & "　元整，另代收規費新台幣　" & IIf(Val(Trim(txt1(2))) = 0, strSpaceAmt, String(LenB(StrConv(ChangeNumber(txt1(2)), vbFromUnicode)), " ")) & "　元整，"
   iStr(10) = "　　　　共計新台幣　" & IIf(Val(Trim(txt1(3))) = 0, strSpaceAmt, String(LenB(StrConv(ChangeNumber(txt1(3)), vbFromUnicode)), " ")) & "　元整。"
   'end 2017/03/28
   iStr(11) = "第三條　乙方對於甲方所委辦之案件內容，於辦理中應嚴守秘密不得外洩，"
   iStr(12) = "　　　　否則應對甲方負損害賠償責任。"
   iStr(13) = "第四條　甲方確保所交付予乙方之資料，均無虛偽情事，如因不實致生損"
   iStr(14) = "　　　　害時，概由甲方負責，與乙方無關。"
   iStr(15) = "第五條　乙方於辦理過程中，應隨時將辦理經過儘速通知或交付甲方。但"
   iStr(16) = "　　　　甲方於簽約後變更聯絡處所，未通知乙方，因而聯絡不及致延誤"
   iStr(17) = "　　　　時限者，乙方不負責任。"
   iStr(18) = "第六條　凡經乙方正式通知甲方之任何事項，如甲方未依限答覆致延誤時"
   iStr(19) = "　　　　限，乙方不負責任。經乙方通知甲方繳費而未依限繳納者，亦同。"
   iStr(20) = "第七條　甲方如逕自撤回所委辦程序，或未經乙方同意終止契約時，所約"
   iStr(21) = "　　　　定之費用，仍應全數給付。"
   iStr(22) = "第八條　本約一式二份，經甲方暨乙方之經手人簽字或蓋章後生效，並由"
   iStr(23) = "　　　　雙方各執乙份為憑。"
   iStr(24) = "　"
   iStr(25) = "　"
   iStr(26) = " "
   iStr(27) = "　委任人（甲方）：" & StrToStr(txt1(4) & String(48, " "), 24)
   iStr(28) = "　　　　　代表人：" & StrToStr(txt1(5) & String(48, " "), 24)
   iStr(29) = "　　　　　地　址：" & StrToStr(StrConv(MidB(StrConv(txt1(6), vbFromUnicode), 1, 48), vbUnicode) & String(48, " "), 24)
   iStr(30) = "　　　　　電　話：" & StrToStr(txt1(7) & String(22, " "), 11)
   iStr(31) = "　受任人（乙方）：" & Combo2.Text 'Add By Sindy 2011/3/23 台一國際專利商標事務所"
   iStr(32) = "　　　　　經手人：" & StrToStr(txt1(8) & String(30, " "), 15)
   'Modified by Lydia 2020/04/09 改用模組控制
   'iStr(33) = "　　　　　地　址：台北市長安東路二段一一二號九樓"
   iStr(33) = "　　　　　地　址：" & PUB_SetAddrTofrm210114(Combo2.Text)
   iStr(34) = "　　　　　電　話：（０２）２５０６１０２３（總機）"
   iStr(35) = "　　　　　傳　真：（０２）２５０１１６６６"
   iStr(36) = " "
   iStr(37) = " "
   iStr(38) = "  中    華    民    國 " & String((10 - LenB(StrConv((txt1(9)), vbFromUnicode))) / 2, " ") & txt1(9) & String((10 - LenB(StrConv((txt1(9)), vbFromUnicode))) / 2, " ") & "年" & String((10 - LenB(StrConv((txt1(10)), vbFromUnicode))) / 2, " ") & txt1(10) & String((10 - LenB(StrConv((txt1(10)), vbFromUnicode))) / 2, " ") & "月" & String((10 - LenB(StrConv((txt1(11)), vbFromUnicode))) / 2, " ") & txt1(11) & String((10 - LenB(StrConv((txt1(11)), vbFromUnicode))) / 2, " ") & "日"
   'Added by Lydia 2017/03/28 有用印就記錄列印內容
   If iPrintC = 1 And bolAddSeal = True Then
           strDetail = ""
           For intI = 1 To UBound(iStr)
              If Trim(iStr(intI)) <> "" Then
                If (intI >= 1 And intI <= 8) Or (intI >= 27 And intI <= 32) Or intI = 38 Then
                   If intI = 27 Then strDetail = strDetail & vbCrLf
                   strDetail = strDetail & RTrim(iStr(intI)) & vbCrLf
                ElseIf intI = 9 Then
                   strDetail = strDetail & RTrim("　　　　新台幣　" & IIf(Val(Trim(txt1(1))) = 0, strSpaceAmt, Replace(ChangeNumber(txt1(1)), "元整", "")) & "　元整，另代收規費新台幣　" & IIf(Val(Trim(txt1(2))) = 0, strSpaceAmt, Replace(ChangeNumber(txt1(2)), "元整", "")) & "　元整，") & vbCrLf
                ElseIf intI = 10 Then
                   strDetail = strDetail & RTrim("　　　　共計新台幣　" & IIf(Val(Trim(txt1(3))) = 0, strSpaceAmt, Replace(ChangeNumber(txt1(1)), "元整", "")) & "　元整。") & vbCrLf
                End If
              End If
           Next
        'Modified by Lydia 2017/04/17 空白用印改由勾選項目控制
        'If PUB_AddRecSeal("6", txtPCnt.Text, IIf(ChkSeal.Value = 1, "", "Y"), strDetail, Combo2.Text) Then
        'Remove by Lydia 2017/05/16 用印記錄移到pdf建立
        'If PUB_AddRecSeal("6", txtPCnt.Text, IIf(bolSpace = True, "Y", ""), strDetail, Combo2.Text) Then
        'End If
   End If
   'end 2017/03/28
        
   iY = 0
   Printer.PaperSize = 9

   'add by nickc 2007/05/04
   'edit by nickc 2007/07/12 試著解決第二頁的格子線會不見的問題
   'If iCount = 1 Then
       Printer.Orientation = 1
   'End If
   Printer.FontName = "標楷體"
   Printer.FontSize = 20
   Printer.FontBold = True
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(iStr(1))) / 2
   iY = iY + Printer.TextHeight(iStr(1))
   Printer.CurrentY = iY
   iY = iY + ((Printer.TextHeight(iStr(1)) / 3) * 4)
   Printer.Print iStr(1)
   Printer.FontBold = False
   Printer.FontSize = 14
   Printer.CurrentX = 1000
   iY = iY + Printer.TextHeight(iStr(2))
   Printer.CurrentY = iY
   iY = iY + ((Printer.TextHeight(iStr(2)) / 3) * 4)
   Printer.Print iStr(2)
   Printer.FontSize = 14
   'Added by Lydia 2017/03/28 同步用印
   If bolAddSeal = True Then
      '列印座置抓乙方資料的起始
      'X軸
      strExc(1) = 1000 + (Printer.TextWidth("　") * 28)
      'Y軸
      strExc(2) = iY + ((Printer.TextHeight("　") / 3) * 4) * 30
      'Added by Lydia 2017/04/25 圖片尺寸
      strExc(3) = 1600 'width
      strExc(4) = 1600 'height
      
      'Added by Lydia 2020/03/25 已記錄公司名稱|用印編號
      intI = InStr(strCompSeal, Combo2.Text)
      If intI > 0 Then
         strExc(9) = Mid(strCompSeal, intI + Len(Combo2.Text))
         If InStr(strExc(9), ",") > 0 Then
             strExc(9) = Mid(strExc(9), 2, InStr(strExc(9), ",") - 2)
         Else
             strExc(9) = Mid(strExc(9), 2)
         End If
          If PUB_ReadDB2File(strSealFile, Val(strExc(9))) Then
             Set tObj = pvGetStdPicture(strSealFile)
             Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
          End If
      Else
      'end 2020/03/25
            If InStr(Combo2.Text, "專利法律") > 0 Then
              If PUB_ReadDB2File(strSealFile, 51) Then
                 Set tObj = pvGetStdPicture(strSealFile)
                 'Modified by Lydia 2017/04/25
                 'Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), 1570, 1570
                 Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
              End If
            End If
            If InStr(Combo2.Text, "專利商標") > 0 Then
              If PUB_ReadDB2File(strSealFile, 52) Then
                 Set tObj = pvGetStdPicture(strSealFile)
                 'Modified by Lydia 2017/04/25
                 'Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), 1570, 1570
                 Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
              End If
            End If
      End If 'Added by Lydia 2020/03/25
   End If
   'end 2017/03/28
   For tmpI = 3 To UBound(iStr) - 1
       If iStr(tmpI) <> "" Then
           Printer.CurrentX = 1000
           Printer.CurrentY = iY
           Printer.Print iStr(tmpI)
           If tmpI = 9 Then
               Printer.FontBold = True
               Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 8) - 30
               Printer.CurrentY = iY
               'Modified by Lydia 2017/03/28
               'Printer.Print ChangeNumber(txt1(1))
              ' Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 21) + Printer.TextWidth(ChangeNumber(txt1(1))) - 30
               If Val(Trim(txt1(1))) = 0 Then
                   Printer.Print strSpaceAmt
               Else
                   'Modified by Lydia 2023/08/10 改變數控制
                   'Printer.Print Replace(ChangeNumber(txt1(1)), "元整", "")
                   Printer.Print ChangeNumber(txt1(1), False)
               End If
               'Modified by Lydia 2023/08/10 改變數控制
               'Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 21) + Printer.TextWidth(IIf(Val(Trim(txt1(1))) = 0, strSpaceAmt, ChangeNumber(txt1(1)))) - 30
               Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 21) + Printer.TextWidth(IIf(Val(Trim(txt1(1))) = 0, strSpaceAmt, ChangeNumber(txt1(1), False))) - 30
               'end 2017/03/28
               Printer.CurrentY = iY
               'Modified by Lydia 2017/03/28
               'Printer.Print ChangeNumber(txt1(2))
               If Val(Trim(txt1(2))) = 0 Then
                   Printer.Print strSpaceAmt
               Else
                   'Modified by Lydia 2023/08/10 改變數控制
                   'Printer.Print Replace(ChangeNumber(txt1(2)), "元整", "")
                   Printer.Print ChangeNumber(txt1(2), False)
               End If
               'end 2017/03/28
               Printer.FontBold = False
           ElseIf tmpI = 10 Then
               Printer.FontBold = True
               Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 10) - 30
               Printer.CurrentY = iY
               'Modified by Lydia 2017/03/28
               'Printer.Print ChangeNumber(txt1(3))
               If Val(Trim(txt1(3))) = 0 Then
                   Printer.Print strSpaceAmt
               Else
                   'Modified by Lydia 2023/08/10 改變數控制
                   'Printer.Print Replace(ChangeNumber(txt1(3)), "元整", "")
                   Printer.Print ChangeNumber(txt1(3), False)
               End If
               'end 2017/03/28
               Printer.FontBold = False
           End If
           If tmpI = 26 Then
               Printer.FontSize = 14
           End If
           iY = iY + ((Printer.TextHeight(iStr(tmpI)) / 3) * 4)
   '        '畫線
           Select Case tmpI
           Case 7
                Printer.Line (1000 + (Printer.TextWidth("　") * 6), iY - 50)-(1000 + (Printer.TextWidth("　") * 33), iY - 50)
           Case 9
                'Modified by Lydia 2017/03/28
                'Printer.Line (1000 + (Printer.TextWidth("　") * 7), iY - 50)-(1000 + (Printer.TextWidth("　") * 9) + Printer.TextWidth(ChangeNumber(txt1(1))), iY - 50)
                'Printer.Line (1000 + (Printer.TextWidth("　") * 20) + Printer.TextWidth(ChangeNumber(txt1(1))), iY - 50)-(1000 + (Printer.TextWidth("　") * 22) + Printer.TextWidth(ChangeNumber(txt1(1))) + Printer.TextWidth(ChangeNumber(txt1(2))), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 7), iY - 50)-(1000 + (Printer.TextWidth("　") * 9) + Printer.TextWidth(IIf(Val(Trim(txt1(1))) = 0, strSpaceAmt, ChangeNumber(txt1(1)))), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 20) + Printer.TextWidth(IIf(Val(Trim(txt1(1))) = 0, strSpaceAmt, ChangeNumber(txt1(1)))), iY - 50)-(1000 + (Printer.TextWidth("　") * 22) + Printer.TextWidth(IIf(Val(Trim(txt1(1))) = 0, strSpaceAmt, ChangeNumber(txt1(1)))) + Printer.TextWidth(IIf(Val(Trim(txt1(2))) = 0, strSpaceAmt, ChangeNumber(txt1(2)))), iY - 50)
                'end 2017/03/28
           Case 10
                'Modified by Lydia 2017/03/28
                'Printer.Line (1000 + (Printer.TextWidth("　") * 9), iY - 50)-(1000 + (Printer.TextWidth("　") * 11) + Printer.TextWidth(ChangeNumber(txt1(3))), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 9), iY - 50)-(1000 + (Printer.TextWidth("　") * 11) + Printer.TextWidth(IIf(Val(Trim(txt1(3))) = 0, strSpaceAmt, ChangeNumber(txt1(3)))), iY - 50)
           Case 27, 28, 29, 30, 32
                Printer.Line (1000 + (Printer.TextWidth("　") * 9), iY - 50)-(1000 + (Printer.TextWidth("　") * 33), iY - 50)
           Case Else
           End Select
       End If
   Next tmpI
   iY = iY + ((Printer.TextHeight(iStr(tmpI)) / 3) * 4)
   Printer.FontSize = 16
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(iStr(UBound(iStr)))) / 2
   Printer.CurrentY = iY
   Printer.Print iStr(UBound(iStr))
   Printer.EndDoc
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

'Modified by Lydia 2022/01/24 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
'Add By Sindy 98/02/11
Dim intLen As Integer
   
   If KeyAscii <> 8 Then
      intLen = GetTextLength(txt1(Index))
      intLen = intLen + GetTextLength(Chr(KeyAscii))
      '2014/5/13 modify by sonia
      'If intLen > txt1(Index).MaxLength Then KeyAscii = 0
      If CheckLengthIsOK(txt1(Index).Text & Chr(KeyAscii), txt1(Index).MaxLength) = False Then
         KeyAscii = 0
      End If
      'end 2014/5/13
   End If
   '98/02/11 End
   If Index = 1 Or Index = 2 Or Index = 3 Or Index = 9 Or Index = 10 Or Index = 11 Then
       If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 46 Then
           KeyAscii = 0
       End If
   End If
   '2009/11/13 ADD BY SONIA
   If Index = 9 Or Index = 10 Or Index = 11 Then
      If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
         KeyAscii = 0
      End If
   End If
   '2009/11/13 END
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) <> "" Then
       'Modified by Lydia 2018/04/13
       'txt1(Index).Text = Replace(Replace(txt1(Index).Text, Chr(10), ""), Chr(13), "")
       txt1(Index).Text = PUB_StringFilter(txt1(Index).Text)
       Cancel = False
       If CheckLengthIsOK(txt1(Index).Text, txt1(Index).MaxLength) = False Then
           txt1(Index).SetFocus
           txt1_GotFocus Index
           Cancel = True
           Exit Sub
       End If
       If Index = 1 Or Index = 2 Then
           txt1(3) = Val(txt1(1)) + Val(txt1(2))
       End If
       If Index = 10 Then
           If Val(txt1(Index)) > 12 Or Val(txt1(Index)) < 1 Then
               MsgBox "月份輸入錯誤！", vbExclamation, "操作錯誤！"
               txt1(Index).SetFocus
               txt1_GotFocus Index
               Cancel = True
               Exit Sub
           End If
       ElseIf Index = 11 Then
           If Val(txt1(Index)) > 31 Or Val(txt1(Index)) < 1 Then
               MsgBox "日輸入錯誤！", vbExclamation, "操作錯誤！"
               txt1(Index).SetFocus
               txt1_GotFocus Index
               Cancel = True
               Exit Sub
           End If
       End If
   End If
End Sub

Private Sub txtPCnt_GotFocus()
   txtPCnt.SelStart = 0
   txtPCnt.SelLength = Len(txtPCnt)
End Sub

Private Sub txtPCnt_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 46 Then
       KeyAscii = 0
   End If
End Sub

'Add By Sindy 2010/4/29
Private Function SetCustTxt(strCUCode As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   SetCustTxt = False
   strCUCode = Left(strCUCode & "000000000", 9)
   'Modified by Morgan 2021/5/5
   'StrSQLa = "Select * From Customer,nation,potcustcont Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "' and CU10=na01(+) and pcc01(+)=substr(CU08, 1, 8) And pcc02(+)=substr(CU08, 9, 1) "
   StrSQLa = "Select * From Customer Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "'"
   'end 2021/5/5
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      SetCustTxt = True
      '申請人中文
      Me.txt1(4).Text = "" & rsA("CU04").Value
      'ID No.
'      Me.txt1(6).Text = "" & rsA("CU11").Value
      '申請地址
      Me.txt1(6).Text = "" & rsA("CU23").Value
'      '國籍
'      Me.txt1(8).Text = "" & rsA("NA03").Value
'      '聯絡人地址
'      If "" & rsA("CU08").Value <> "" Then
'         Me.txt1(9).Text = "" & rsA("pcc22").Value
'      Else
'         Me.txt1(9).Text = "" & rsA("CU31").Value
'      End If
      '電話1
      Me.txt1(7).Text = "" & rsA("CU16").Value
'      '傳真1
'      Me.txt1(18).Text = "" & rsA("CU18").Value
      '代表人1中文
      Me.txt1(5).Text = "" & rsA("CU07").Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

'Add by Amy 2016/08/19
Private Sub UpdReceiptCmp(ByVal stNowCustNo As String, ByVal stNowCmp As String)
    Dim strUpd As String
    
     Exit Sub 'Added by Lydia 2022/08/30 受任人下拉預設只剩下台一國際智慧財產事務所，所以不必再更新客戶檔的了。
     
    'Add by Amy 2016/12/30 +同業務區或為MCTF同組人員才可回寫收據公司別
    If ChkSameCuArea(stNowCustNo, strUserNum) = False Then Exit Sub
    
    'Modified by Lydia 2019/04/12 拿掉UpdateID,Date,Time(CU84,CU85,CU86)
    'strUpd = "Update Customer Set CU84='" & strUserNum & "',CU85=to_number(to_char(sysdate,'YYYYMMDD')),CU86=to_number(to_char(sysdate,'HH24MI')),CU164='" & stNowCmp & "' " & _
                  "Where CU01='" & Left(stNowCustNo, 8) & "' And CU02='" & Mid(stNowCustNo, 9, 1) & "' "
    strUpd = "Update Customer Set CU164='" & stNowCmp & "' " & _
                  "Where CU01='" & Left(stNowCustNo, 8) & "' And CU02='" & Mid(stNowCustNo, 9, 1) & "' "
    Pub_SeekTbLog strUpd
    'Modified by Lydia 2019/04/23 觸發Trigger
    'cnnConnection.Execute strUpd
    cnnConnection.Execute "begin user_data.user_enabled:=1; " & strUpd & " ; end; "
End Sub

'Added by Lydia 2017/04/13 列印:先轉PDF,列印後刪檔
Private Sub Print2PDF(ByVal bSpace As Boolean)
Dim strFileName As String
Dim strOldName As String 'Added by Lydia 2017/06/07

'Added by Lydia 2017/04/25 VB印表機實際列印的左邊界、右邊界
Set Printer = Printers(PUB_PrinterIndex(Combo1.Text))
d_Top = Format((Printer.Height - Printer.ScaleHeight) / 2, "0") '直印
d_Left = Format((Printer.Width - Printer.ScaleWidth) / 2, "0")
'end 2017/04/25

strDetail = "" 'Added by Lydia 2017/05/16
strOldName = App.Title 'Added by Lydia 2017/06/07

Screen.MousePointer = vbHourglass
    'Modified by Lydia 2022/01/24 先產生Word檔，後轉成PDF檔逐一列印
'    For iCount = 1 To Val(txtPCnt)
'        iPrintC = iCount
'        'Modified by Lydia 2017/06/06 改用App.Title變更印表機列印文件名稱(執行exe檔有效,VB跑無效)
'        'strFileName = strUserNum & "_條碼_" & IIf(bSpace = False, IIf(Trim(txt1(4)) <> "", Mid(Trim(txt1(4)), 1, 4), Mid(Trim(txt1(5)), 1, 4)), "空白") & iCount & ".pdf"
'        'If Dir(App.path & "\" & strFileName) <> "" Then
'        '   Kill App.path & "\" & strFileName
'        'End If
'        ''轉PDF
'        'frmPDF.Show
'        'frmPDF.StartProcess App.path, strFileName
'        'Call StrMenu(bSpace)
'        'frmPDF.EndtProcess
'        'Unload frmPDF
'        strFileName = strUserNum & "_條碼_" & IIf(bSpace = False, IIf(Trim(txt1(5)) <> "", Mid(Trim(txt1(5)), 1, 4), Mid(Trim(txt1(6)), 1, 4)), "空白") & iCount
'        App.Title = strFileName
'        Call StrMenu(bSpace)
'        'end 2017/06/07
'
'        'Added by Lydia 2017/05/16 用印記錄移到pdf建立
'        If iCount = 1 And strDetail <> "" Then
'           'If Dir(App.path & "\" & strFileName) <> "" Then 'Remove by Lydia 2020/03/16 因為不存檔案所以取消檔案檢查(自2017/06/08~2020/03/16無用印記錄)
'              If PUB_AddRecSeal("6", txtPCnt.Text, IIf(bSpace = True, "Y", ""), strDetail, Combo2.Text) Then
'              End If
'           'End If 'Remove by Lydia 2020/03/16
'        End If
'        'end 2017/05/16
'
'        'Remove by Lydia 2017/06/07
'        ''列印PDF
'        'PUB_PrintPDF App.path & "\" & strFileName, Me.Combo1
'        ''刪除PDF
'        'Kill App.path & "\" & strFileName
'    Next iCount
    Call runWordProc(bSpace)
    If m_TempPDF <> "" Then
        For iCount = 1 To Val(txtPCnt)
            iPrintC = iCount
            strFileName = strUserNum & "_條碼_" & m_TempFN & iCount
            PUB_PrintPDF App.path & "\" & strUserNum & "\" & m_TempPDF, Combo1.Text
            App.Title = strFileName
        Next iCount
    End If
'--------------先產生Word檔，後轉成PDF檔逐一列印

    App.Title = strOldName 'Added by Lydia 2017/06/07

End Sub

'Added by Lydia 2022/01/24 下載Word範本套印
Private Sub runWordProc(ByVal pSpace As Boolean)
Dim iStr(1 To 38) As String    '用印記錄(全文)
Dim strSealFile As String '公司章圖檔
Dim strSpaceAmt As String
Dim strName As String
Dim strText As String
Dim intA As Integer
Dim m_FileName As String, m_TempFileName As String
Dim m_DefPath As String
Dim oShape
Dim oWord

On Error GoTo ErrHand

   '上傳檔案
   'Modified by Lydia 2024/07/22 改用變數
   'intI = SaveImgByteFile("\\" & pub_getspecman("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M51-000300-0-06 智權部委任契約書_條碼.docx", "M51", "000300", "0", "06", "4", "1")

   m_DefPath = App.path & "\" & strUserNum
   'Added by Lydia 2022/01/25
   m_TempPDF = ""
   '變更Word印表機
   PUB_SetOsDefaultPrinter Combo1
   PUB_SetWordActivePrinter
   'end 2022/01/25
   
    strDetail = ""
    
   '下載範本檔: M51-000300-0-06 智權部委任契約書_條碼.docx
   m_TempFN = Pub_RepFileName(IIf(pSpace = False, IIf(Trim(txt1(4)) <> "", Mid(Trim(txt1(4)), 1, 4), Mid(Trim(txt1(5)), 1, 4)), "空白")) 'Move by Lydia 2022/01/25 從m_TempFileName移過來
   'Modified by Lydia 2022/01/25 改成Word直接印，所以範本一開始就先命名好
   'm_FileName = "$$" & Me.Name & ".docx"
   m_FileName = "$$" & strUserNum & "_條碼_" & m_TempFN & ".docx"
   If Dir(m_DefPath & "\" & m_FileName) <> "" Then
      Kill m_DefPath & "\" & m_FileName
   End If
   If PUB_GetSampleFile(m_FileName, "M51-000300-0-06", , m_DefPath) = False Then
        Exit Sub
   End If
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   'Remove by Lydia 2022/01/25 不用改存PDF檔
'   m_TempFileName = "$$" & strUserNum & "_條碼_" & m_TempFN & ".pdf"
'   If Dir(m_DefPath & "\" & m_TempFileName) <> "" Then
'      Kill m_DefPath & "\" & m_TempFileName
'   End If
   'end 2022/01/25
   
   '改成直接用範本檔
   'Q: AddToRecentFiles:=False還是會新增到最近開啟記錄
   g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, False, False, False
  
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For intA = 1 To 14
         strName = "PS" & Format(intA, "000")
         strText = ""
'-------第一條
         If intA = 1 Then
              '委辦範圍
              strText = "　　　　" & IIf(Chk1(0).Value = 1, "■", "□") & "　申請廠商號碼"
              strText = strText & vbCrLf & "　　　　" & IIf(Chk1(1).Value = 1, "■", "□") & "　製作正片"
              strText = strText & vbCrLf & "　　　　" & IIf(Chk1(2).Value = 1, "■", "□") & "　正片測試及陳報商品基本資料明細"
         ElseIf intA = 2 Then
              '委辦範圍: 其他
              strText = "　　　　" & IIf(Chk1(3).Value = 1, "■　" & PUB_StrToStr(txt1(0) & " ", 54), "□")
'-------第二條
         ElseIf intA >= 3 And intA <= 5 Then
                '委辦費用1,2,3
                If Val(Trim(txt1(intA - 2))) = 0 Then
                   strText = String(8, "　")
                Else
                   'Modified by Lydia 2023/08/10 改變數控制
                   'strText = " " & ChangeNumber(txt1(intA - 2))
                   strText = " " & ChangeNumber(txt1(intA - 2), False)
                End If
                strText = Replace(strText, "元整", "") 'Added by Lydia 2022/06/24
                If GetTextLength(strText) <= 16 Then 'Added by Lydia 2022/06/24 判斷超過字元長度不限制
                     'Modified by Lydia 2023/08/10
                     'strText = PUB_StrToStr(Replace(strText, "元整", ""), 16, True)
                     strText = PUB_StrToStr(strText, 16, True)
                End If   'Added by Lydia 2022/06/24
'-------
         ElseIf intA = 6 Then
              '委任人（甲方）
              strText = PUB_StrToStr(txt1(4), 44)
         ElseIf intA = 7 Then
              '委任人（代表人）
              strText = PUB_StrToStr(txt1(5), 44)
         ElseIf intA = 8 Then
              '委任人（地　址）
              strText = PUB_StrToStr(txt1(6), 44)
         ElseIf intA = 9 Then
              '委任人（電　話）
              strText = PUB_StrToStr(txt1(7), 44)
         ElseIf intA = 10 Then
              '受任人（乙方）
              strText = Combo2.Text
         ElseIf intA = 11 Then
              '經手人
              strText = PUB_StrToStr(txt1(8), 30)
         ElseIf intA = 12 Then
              '地　址
              strText = PUB_SetAddrTofrm210114(Combo2.Text)
         ElseIf intA = 13 Then
              strText = "        中    華    民    國 " & String((8 - LenB(StrConv((txt1(9)), vbFromUnicode))) / 2, " ") & txt1(9) & String((8 - LenB(StrConv((txt1(9)), vbFromUnicode))) / 2, " ") & "年" & String((8 - LenB(StrConv((txt1(10)), vbFromUnicode))) / 2, " ") & txt1(10) & String((8 - LenB(StrConv((txt1(10)), vbFromUnicode))) / 2, " ") & "月" & String((8 - LenB(StrConv((txt1(11)), vbFromUnicode))) / 2, " ") & txt1(11) & String((8 - LenB(StrConv((txt1(11)), vbFromUnicode))) / 2, " ") & "日"
         ElseIf intA = 14 Then
              strText = ""
         Else
         End If
         
         If Trim(strName) <> "" Then
            .Selection.Find.ClearFormatting
            .Selection.Find.Text = "|#" & strName & "#|"
            .Selection.Find.Replacement.Text = ""
            .Selection.Find.Forward = True
            .Selection.Find.Wrap = wdFindContinue
            .Selection.Find.Format = False
            .Selection.Find.MatchCase = False
            .Selection.Find.MatchWholeWord = False
            .Selection.Find.MatchWildcards = False
            .Selection.Find.MatchSoundsLike = False
            .Selection.Find.MatchAllWordForms = False
            .Selection.Find.MatchByte = True
            .Selection.Find.Execute
            .Selection.Delete
            If (intA >= 3 And intA <= 5) Then
                '金額要粗體
                .Selection.Font.Bold = True
            End If
            '保留;因為先全部以細明體-ExtB,最後全選改字型;
            If (intA >= 6 And intA <= 9) Or intA = 11 Then
               '有Unicode字需要換字型
               .Selection.Font.Name = "細明體-ExtB"
            End If
            
            If intA = 14 And bolAddSeal = True Then  '公司章: 放在受聘人的儲存格
                strExc(9) = Mid(strCompSeal, InStr(strCompSeal, Combo2))
                If InStr(strExc(9), ",") > 0 Then
                    strExc(9) = Right(Mid(strExc(9), 1, InStr(strExc(9), ",") - 1), 2)
                Else
                    strExc(9) = Right(strExc(9), 2)
                End If
                If PUB_ReadDB2File(m_DefPath & "\$$" & Me.Name & "TempFile", Val(strExc(9))) Then
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=m_DefPath & "\$$" & Me.Name & "TempFile", LinkToFile:=False, SaveWithDocument:=True)
                    '--------設定圖片=文蓋圖(文字在前)
                        oShape.Fill.Visible = msoFalse
                        oShape.Fill.Solid
                        oShape.Fill.Transparency = 0#
                        oShape.Line.Weight = 0.75
                        oShape.Line.DashStyle = msoLineSolid
                        oShape.Line.Style = msoLineSingle
                        oShape.Line.Transparency = 0#
                        oShape.Line.Visible = msoFalse
                        oShape.LockAspectRatio = msoTrue
                        oShape.Rotation = 0#
                        oShape.PictureFormat.Brightness = 0.5
                        oShape.PictureFormat.Contrast = 0.5
                        oShape.PictureFormat.ColorType = msoPictureAutomatic
                        oShape.PictureFormat.CropLeft = 0#
                        oShape.PictureFormat.CropRight = 0#
                        oShape.PictureFormat.CropTop = 0#
                        oShape.PictureFormat.CropBottom = 0#

                        oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                        oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                        oShape.Left = .CentimetersToPoints(8.25)
                        oShape.Top = .CentimetersToPoints(0.3)
                        oShape.LockAnchor = False
                        oShape.LayoutInCell = True
                        oShape.WrapFormat.AllowOverlap = True
                        oShape.WrapFormat.Side = wdWrapBoth
                        oShape.WrapFormat.DistanceTop = .CentimetersToPoints(0)
                        oShape.WrapFormat.DistanceBottom = .CentimetersToPoints(0)
                        oShape.WrapFormat.DistanceLeft = .CentimetersToPoints(0.32)
                        oShape.WrapFormat.DistanceRight = .CentimetersToPoints(0.32)
                        oShape.WrapFormat.Type = 3
                        oShape.ZOrder 5 '文蓋圖(文字在前)
                        '---------------------------
                End If
          
            End If
            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strText

            If (intA >= 3 And intA <= 5) Then
                '金額要粗體 =>還原
                .Selection.Font.Bold = False
            End If
            If (intA >= 6 And intA <= 9) Or intA = 11 Then
               '有Unicode字需要換字型 =>還原
               .Selection.Font.Name = "標楷體"
            End If
         End If
      Next intA
      '因為先全部以細明體-ExtB,最後全選改字型;
      .Selection.WholeStory
      .Selection.Font.Name = "標楷體"
   End With

   '改存成PDF檔
   'Memo by Lydia 2022/01/25  因為受PDF redirect設定灰階列印影響，改成Word直接印
   intA = IIf(Val(txtPCnt) = 0, 1, Val(txtPCnt))
   For intI = 1 To intA
       g_WordAp.PrintOut Background:=False, Range:=4, Item:=0, Copies:=1, Pages:="1", Collate:=True
   Next intI
   
   '保留: 存檔
   'g_WordAp.ActiveDocument.Close wdSaveChanges
   g_WordAp.Quit wdDoNotSaveChanges
   Set g_WordAp = Nothing '避免快速開啟Word,程式出錯
   m_TempPDF = m_FileName 'Added by Lydia 2022/01/25
   
   'Mark by Lydia 2022/01/25 因為受PDF redirect設定灰階列印影響，改成Word直接印
   'If PUB_PrintWord2PDF(g_WordAp, m_DefPath, m_TempFileName, m_TempPDF) = False Then
   '    Exit Sub
   'End If
   'end 2022/01/19
   
If bolAddSeal = True Then  '用印記錄
   strDetail = ""
   iStr(1) = "條碼案件委任契約書"
   iStr(2) = "委任人(甲方)茲委任受任人(乙方)辦理條碼案件，雙方同意條件如下："
   iStr(3) = "第一條　委辦範圍："
   iStr(4) = "　　　　" & IIf(Chk1(0).Value = 1, "■", "□") & "　申請廠商號碼"
   iStr(5) = "　　　　" & IIf(Chk1(1).Value = 1, "■", "□") & "　製作正片"
   iStr(6) = "　　　　" & IIf(Chk1(2).Value = 1, "■", "□") & "　正片測試及陳報商品基本資料明細"
   iStr(7) = "　　　　" & IIf(Chk1(3).Value = 1, "■　" & PUB_StrToStr(txt1(0) & " ", 54), "□")
   iStr(8) = "第二條　委辦費用："
   strSpaceAmt = String(6, "　") '"　　拾　　萬　　仟　　佰"  'Added by Lydia 2017/03/28
   iStr(9) = "　　　　新台幣　" & IIf(Val(Trim(txt1(1))) = 0, strSpaceAmt, String(LenB(StrConv(ChangeNumber(txt1(1)), vbFromUnicode)), " ")) & "　元整，另代收規費新台幣　" & IIf(Val(Trim(txt1(2))) = 0, strSpaceAmt, String(LenB(StrConv(ChangeNumber(txt1(2)), vbFromUnicode)), " ")) & "　元整，"
   iStr(10) = "　　　　共計新台幣　" & IIf(Val(Trim(txt1(3))) = 0, strSpaceAmt, String(LenB(StrConv(ChangeNumber(txt1(3)), vbFromUnicode)), " ")) & "　元整。"
   iStr(11) = "第三條　乙方對於甲方所委辦之案件內容，於辦理中應嚴守秘密不得外洩，"
   iStr(12) = "　　　　否則應對甲方負損害賠償責任。"
   iStr(13) = "第四條　甲方確保所交付予乙方之資料，均無虛偽情事，如因不實致生損"
   iStr(14) = "　　　　害時，概由甲方負責，與乙方無關。"
   iStr(15) = "第五條　乙方於辦理過程中，應隨時將辦理經過儘速通知或交付甲方。但"
   iStr(16) = "　　　　甲方於簽約後變更聯絡處所，未即時通知乙方，因而聯絡不及致"
   iStr(17) = "　　　　延誤時限者，乙方不負責任。"
   iStr(18) = "第六條　凡經乙方正式通知甲方之任何事項，如甲方未依限答覆致延誤時"
   iStr(19) = "　　　　限，乙方不負責任。經乙方通知甲方繳費而未依限繳納者，亦同。"
   iStr(20) = "第七條　甲方如逕自撤回所委辦程序，或未經乙方同意終止契約時，所約"
   iStr(21) = "　　　　定之費用，仍應全數給付。"
   iStr(22) = "第八條　本約一式二份，經甲方暨乙方之經手人簽字或蓋章後生效，並由"
   iStr(23) = "　　　　雙方各執乙份為憑。"
   iStr(24) = "　"
   iStr(25) = "　"
   iStr(26) = " "
   iStr(27) = "　委任人（甲方）：" & PUB_StrToStr(txt1(4), 48)
   iStr(28) = "　　　　　代表人：" & PUB_StrToStr(txt1(5), 48)
   iStr(29) = "　　　　　地　址：" & PUB_StrToStr(txt1(6), 48)
   iStr(30) = "　　　　　電　話：" & PUB_StrToStr(txt1(7), 22)
   iStr(31) = "　受任人（乙方）：" & Combo2.Text
   iStr(32) = "　　　　　經手人：" & PUB_StrToStr(txt1(8), 30)
   iStr(33) = "　　　　　地　址：" & PUB_SetAddrTofrm210114(Combo2.Text)
   iStr(34) = "　　　　　電　話：（０２）２５０６１０２３（總機）"
   iStr(35) = "　　　　　傳　真：（０２）２５０１１６６６"
   iStr(36) = " "
   iStr(37) = " "
   iStr(38) = "  中    華    民    國 " & String((10 - LenB(StrConv((txt1(9)), vbFromUnicode))) / 2, " ") & txt1(9) & String((10 - LenB(StrConv((txt1(9)), vbFromUnicode))) / 2, " ") & "年" & String((10 - LenB(StrConv((txt1(10)), vbFromUnicode))) / 2, " ") & txt1(10) & String((10 - LenB(StrConv((txt1(10)), vbFromUnicode))) / 2, " ") & "月" & String((10 - LenB(StrConv((txt1(11)), vbFromUnicode))) / 2, " ") & txt1(11) & String((10 - LenB(StrConv((txt1(11)), vbFromUnicode))) / 2, " ") & "日"
    For intI = 1 To UBound(iStr)
       If Trim(iStr(intI)) <> "" Then
         If (intI >= 1 And intI <= 8) Or (intI >= 27 And intI <= 32) Or intI = 38 Then
            If intI = 27 Then strDetail = strDetail & vbCrLf
            strDetail = strDetail & RTrim(iStr(intI)) & vbCrLf
         ElseIf intI = 9 Then
            'Modified by Lydia 2023/08/10 改變數控制
            'strDetail = strDetail & RTrim("　　　　新台幣　" & IIf(Val(Trim(txt1(1))) = 0, strSpaceAmt, Replace(ChangeNumber(txt1(1)), "元整", "")) & "　元整，另代收規費新台幣　" & IIf(Val(Trim(txt1(2))) = 0, strSpaceAmt, Replace(ChangeNumber(txt1(2)), "元整", "")) & "　元整，") & vbCrLf
            strDetail = strDetail & RTrim("　　　　新台幣　" & IIf(Val(Trim(txt1(1))) = 0, strSpaceAmt, ChangeNumber(txt1(1), False)) & "　元整，另代收規費新台幣　" & IIf(Val(Trim(txt1(2))) = 0, strSpaceAmt, ChangeNumber(txt1(2), False)) & "　元整，") & vbCrLf
         ElseIf intI = 10 Then
            'Modified by Lydia 2023/08/10 改變數控制
            'strDetail = strDetail & RTrim("　　　　共計新台幣　" & IIf(Val(Trim(txt1(3))) = 0, strSpaceAmt, Replace(ChangeNumber(txt1(1)), "元整", "")) & "　元整。") & vbCrLf
            strDetail = strDetail & RTrim("　　　　共計新台幣　" & IIf(Val(Trim(txt1(3))) = 0, strSpaceAmt, ChangeNumber(txt1(1), False)) & "　元整。") & vbCrLf
         End If
       End If
    Next
    If PUB_AddRecSeal("6", txtPCnt.Text, IIf(pSpace = True, "Y", ""), strDetail, Combo2.Text) Then
    End If
End If
          
   Exit Sub
   
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.Number <> 0 Then
      MsgBox Err.Number & ":" & Err.Description, , "錯誤 "
   End If
   
End Sub

'Added by Lydia 2022/01/24 刪除暫存檔
Private Sub RunEndProc(ByVal bolSleep As Boolean)
   If bolSleep = True Then Sleep 3000
   PUB_KillTempFile (strUserNum & "\$$" & strUserNum & "*_條碼*.*")
   PUB_KillTempFile (strUserNum & "\$$" & Me.Name & "*.*")
    
End Sub

