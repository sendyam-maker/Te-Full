VERSION 5.00
Begin VB.Form Frmacc24l0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "請款單月報表"
   ClientHeight    =   2760
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8928
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   8928
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   2
      Top             =   930
      Width           =   1125
   End
   Begin VB.TextBox Text2 
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
      Height          =   315
      Left            =   1980
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1710
      Width           =   360
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
      Height          =   315
      Left            =   1980
      TabIndex        =   7
      Top             =   1320
      Width           =   2250
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1305
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   570
      Width           =   7500
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "執行(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   2130
      Width           =   6900
   End
   Begin VB.TextBox Text3 
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
      Height          =   315
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   1
      Top             =   930
      Width           =   1125
   End
   Begin VB.Line Line1 
      X1              =   2460
      X2              =   2640
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "是否修改月報表：      (Y:是)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   3
      Left            =   180
      TabIndex        =   9
      Top             =   1776
      Width           =   2880
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Purchase Order："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   1380
      Width           =   1665
   End
   Begin VB.Label Label2 
      Caption         =   "注意:本報表只列未結清資料且除Y45493外僅適用純美金格式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   180
      TabIndex        =   6
      Top             =   150
      Width           =   6315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "請款年月：                                   (Ex. 10301 )"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   180
      TabIndex        =   5
      Top             =   960
      Width           =   4965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "請款對象："
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
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   600
      Width           =   1155
   End
End
Attribute VB_Name = "Frmacc24l0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2014/2/24
Option Explicit

Const cTemplateName As String = "Template.xls"
Const cfmtDollar = "#,##0.00"
Dim m_strSavePath As String '電子檔存放路徑
Dim m_FileName As String
Dim m_RptNo As String
Dim m_strNo As String
Dim iPicNo As Integer, iPicNo2 As Integer '信頭、信尾圖檔代碼

Private Function GetPO(pNo As String) As String
   strExc(0) = "select ld17 from ledes where ld01='" & pNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetPO = "" & RsTemp(0)
   End If
End Function

Private Sub Combo1_Click()
   
   Text1 = ""
   'Combo1_Validate False
   If Left(Combo1, 9) = "Y54225B10" Then
      'Modified by Morgan 2015/1/9
      'Text1 = "8700346419"
      'Modified by Morgan 2016/2/17
      'Text1 = "8700468471"
      'Modified by Morgan 2017/1/3 --Lina
      'Text1 = "8700607777"
      'Modified by Morgan 2018/1/10 --Lina
      'Text1 = "8700711072"
      'Modified by Morgan 2019/1/10 --Lina
      'Text1 = "8700820697"
      'Modified by Morgan 2019/2/1 改抓DB
      'Text1 = "8700915432"
      Text1 = GetPO("Y54225B10")
      'end 2015/1/30
   
   'ElseIf Left(Combo1, 9) = "Y48309070" Then 'Removed by Morgan 2024/8/30 未再使用
      'Modified by Morgan 2015/1/30
      'Text1 = "8700347328"
      'Modified by Morgan 2016/2/19
      'Text1 = "8700472463"
      'Modified by Morgan 2017/1/3 --琬姿
      'Text1 = "8700607777"
      'Modified by Morgan 2018/1/9 --黃咸達
      'Text1 = "8700711072"
      'Modified by Morgan 2019/1/10
      '2019/1/1以後請款換 PO number--Monica
      'Text1 = "8700820697"
      'Modified by Morgan 2019/2/1 改抓DB
      'Text1 = "8700915556"
   '   Text1 = GetPO("Y48309070") 'Removed by Morgan 2024/8/30 未再使用
      'end 2015/1/30
      
   End If
   
   'Added by Morgan 2021/2/22
   'Modified by Morgan 2022/2/7 +Y55666000
   If Left(Combo1, 9) = "Y52418000" Or Left(Combo1, 9) = "Y55666000" Then
      Text4.Enabled = True
   Else
      Text4.Enabled = False
   End If
   'end 2021/2/22
End Sub

Private Sub Combo1_GotFocus()
   CloseIme
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   Dim strNo() As String
   If Combo1.ListIndex = -1 And Combo1.Text <> "" Then
      strNo() = Split(Combo1.Text, " ")
      strNo(0) = Left(strNo(0) & "000", 9)
      If Left(strNo(0), 1) = "Y" Then
         Combo1.Text = strNo(0) & " " & GetPrjName2(strNo(0))
      Else
         Combo1.Text = strNo(0) & " " & GetPrjPeople1(strNo(0), 2)
      End If
      Combo1_Click 'Added by Morgan 2021/9/2 用輸的不會觸發導致 PO 沒帶入
   End If
End Sub

Private Sub Command1_Click()
   Dim ii As Integer, strSystemKind As String
   Dim bCancel As Boolean
   Dim hLocalFile As Long
   Dim stRptNo2 As String, stMsg As String 'Added by Morgan 2023/7/19
   
On Error GoTo ErrHnd

   ClearQueryLog Me.Name
   If Combo1 = "" Then
      MsgBox "請輸入請款對象！", vbExclamation
      Combo1.SetFocus
      Exit Sub
   End If
   
   pub_QL05 = pub_QL05 & ";" & Label1(1) & Combo1
   
   If Text3 = "" Then
      MsgBox "請輸入請款年月！", vbExclamation
      Text3.SetFocus
      Exit Sub
   End If
   Text3_Validate bCancel
   If bCancel = True Then
      Exit Sub
   End If
   
   'Added by Morgan 2014/6/16
   m_FileName = ""
   m_strNo = Left(Combo1, 9)
   Select Case m_strNo
      '外專
      'Modified by Morgan 2020/9/21 +Y51971000
      'Modified by Morgan 2021/3/16 +Y20412010 --Franny
      'Modified by Morgan 2022/8/18 +Y48292000 --Tim
      'Modified by Morgan 2023/3/8 +Y55240 --Kahn
      'Modified by Morgan 2023/7/10 +Y54570000 --Kahn
      'Modified by Morgan 2023/7/10 +Y48279000 --Kahn
      'Modified by Morgan 2023/8/1 +Y51467020 --Franny
      'Modified by Morgan 2025/4/24 +Y34049010 --Anny
      'Modified by Morgan 2025/6/19 +Y22327000 --Lisa
      'Modified by Morgan 2025/6/30 +Y22457000--Tim
      Case "Y54225B10", "Y51971000", "Y20412010", "Y48292000", "Y55240000", "Y54570000", "Y48279000", "Y51467020", "Y34049010", "Y22327000", "Y22457000"
         m_RptNo = "1"
         
         'Removed by Morgan 2018/6/28 改用 LEDES
         'm_FileName = "Patent_excel" & m_strNo & Text3
         'Call PUB_GetSampleFile(cTemplateName, "M51-000098-0-01")
         
         'Added by Morgan 2023/7/19 +原來的XLS檔也要 --Kahn
         If m_strNo = "Y54570000" Then
            stRptNo2 = "7"
            m_FileName = m_strNo & "_" & (Val(Text3) + 191100)
         End If
         'end 2023/7/19
         
      '外商
      'Removed by Morgan 2024/8/30 未再使用
      'Case "Y48309070"
      '   m_RptNo = "2"
      'end 2024/8/30
         
         'Removed by Morgan 2018/6/28 改用 LEDES
         'm_FileName = "Trademark_excel" & m_strNo & Text3
         ''SaveImgByteFile("C:\Template.xls","M51","000098","0","02","4","5") '檔案更新
         'Call PUB_GetSampleFile(cTemplateName, "M51-000098-0-02")
         
      'Modified by Morgan 2016/9/30 Y34412010 改單獨(5)
      'Modified by Morgan 2024/5/23 +Y55973000--Izumi
      Case "Y21431000", "Y53280000", "Y55973000"
         m_RptNo = "3"
         
         'Removed by Morgan 2015/3/2
         'm_FileName = "Patent_excel" & m_strNo & Text3
         'Call PUB_GetSampleFile(cTemplateName, "M51-000098-0-03")
         'end 2015/3/2
         
      'Added by Morgan 2015/3/5
      Case "Y45493000"
         m_RptNo = "4"
      
      'Added by Morgan 2016/9/30
      'Modified by Morgan 2016/11/14 +Y52418000 --Lina
      'Modified by Morgan 2022/2/7 +Y55666 NOVOCURE GMBH --Ryan
      'Modified by Morgan 2022/5/12 -Y55666,改獨立用 9
      Case "Y34412010", "Y52418000"
         m_RptNo = "5"
         m_FileName = m_strNo & "_" & (Val(Text3) + 191100)
      
      'Added by Morgan 2018/11/6
      Case "Y52341000"
         m_RptNo = "6"
      'end 2018/11/6
      
      'Added by Morgan 2019/8/19 +Y54570000 Amkor Technology, Inc. --Lisa
      'Removed by Morgan 2023/7/7 改產生Ledes併入1
      'Case "Y54570000"
      '   m_RptNo = "7"
      '   m_FileName = m_strNo & "_" & (Val(Text3) + 191100)
      
      'Added by Morgan 2020/10/5 +Y22327000 MKS -- Lisa
      'Removed by Morgan 2025/6/19  改產生Ledes併入1
      'Case "Y22327000"
      '   m_RptNo = "8"
      '   m_FileName = "\Summary List of Monthly Invoices (" & Format(ChangeTStringToWDateString(Text3 & "01"), "mmmm yyyy") & ")"
      'end 2025/6/19
      
      'Added by Morgan 2022/5/12 從 5 抽出來
      'Modified by Morgan 2022/6/7 +Y55751 ----Ryan
      'Modified by Morgan 2022/8/5 -Y55751--Franny
      'Modified by Morgan 2023/12/20 +X82995000 --Kahn
      'Modified by Morgan 2025/11/4 -Y55666欄位及格式有變，改用11 --Teddy
      Case "X82995000"
         m_RptNo = "9"
         m_FileName = m_strNo & "_" & (Val(Text3) + 191100)
      'end 2022/5/12
      
      'Added by Morgan 2024/8/30
      Case "Y56066000", "Y55751000"
         m_RptNo = "10"
         
      'Added by Morgan 2024/9/3
      Case "Y34126000"
         m_RptNo = "11"
         
      'Added by Morgan 2025/8/21
      Case "Y56199000"
         m_RptNo = "12"
         m_FileName = "Coupang Invoice List_TaiE_" & (Val(Text3) + 191100)
         
      'Added by Morgan 2025/11/4
      Case "Y55666000"
         m_RptNo = "13"
         m_FileName = m_strNo & "_" & (Val(Text3) + 191100)
         
      Case Else
         m_RptNo = "4"
         'Exit Sub
   End Select
   'end 2014/6/16
   
   Screen.MousePointer = vbHourglass
      
   pub_QL05 = pub_QL05 & ";" & Label1(2) & Text3
   
ReRun:
   
   'Added by Morgan 2018/6/28
   If m_RptNo = "1" Then
      If SaveLEDES = True Then
         'Added by Morgan 2023/7/19
         If stRptNo2 <> "" Then
            stMsg = "LEDES帳單電子檔已存於" & m_strSavePath & vbCrLf & vbCrLf
         Else
         'end 2023/7/1
            If MsgBox("LEDES帳單電子檔已存於" & m_strSavePath & vbCrLf & vbCrLf & "是否要開啟資料夾？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
               ShellExecute hLocalFile, "open", m_strSavePath, vbNullString, vbNullString, 1
            End If
         End If
      End If
   'end 2018/6/28
      
   'Added by Morgan 2018/11/6
   ElseIf m_RptNo = "6" Then
      If SaveLEDES2 = True Then
         If MsgBox("LEDES帳單電子檔已存於" & m_strSavePath & vbCrLf & vbCrLf & "是否要開啟資料夾？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
            ShellExecute hLocalFile, "open", m_strSavePath, vbNullString, vbNullString, 1
         End If
      End If
   'end 2018/11/6
   
   'Added by Morgan 2024/9/2
   ElseIf m_RptNo = "10" Then
      If CopyDN = True Then
         If MsgBox("請款單電子檔" & IIf(Text2 = "", "及月報表", "") & "已存於" & m_strSavePath & vbCrLf & vbCrLf & "是否要開啟資料夾？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
            ShellExecute hLocalFile, "open", m_strSavePath, vbNullString, vbNullString, 1
         End If
      End If
   'end 2024/9/2
   
   'Added by Morgan 2024/9/3
   ElseIf m_RptNo = "11" Then
      If CopyLEDES = True Then
         If MsgBox("LEDES帳單電子檔" & IIf(Text2 = "", "及月報表", "") & "已存於" & m_strSavePath & vbCrLf & vbCrLf & "是否要開啟資料夾？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
            ShellExecute hLocalFile, "open", m_strSavePath, vbNullString, vbNullString, 1
         End If
      End If
   'end 2024/9/3
   
   'Added by Morgan 2015/3/2
   '先正達以外改不要Excel,要PDF
   ElseIf m_RptNo > 2 Then
      strExc(0) = "select a1k01,a1k33 from acc1k0 where nvl(a1k12,0)=0 and a1k25||a1k29 is null and a1k02>=" & Text3 & "01 and a1k02<=" & Text4 & "31 and a1k28='" & m_strNo & "' order by a1k01 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      InsertQueryLog RsTemp.RecordCount
      If intI = 1 Then
         SetOutPutPath
         If m_RptNo = "3" Then
            If PdfSave2 = True Then
               If Text2 <> "Y" Then
                  If MsgBox("PDF 電子檔已存於" & m_strSavePath & vbCrLf & vbCrLf & "是否要開啟資料夾？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                     ShellExecute hLocalFile, "open", m_strSavePath, vbNullString, vbNullString, 1
                  End If
               End If
            End If
            
         'Added by Morgan 2015/3/5
         ElseIf m_RptNo = "4" Then
            
            If RsTemp.Fields("a1k33") = "3" Then
               If PdfSave3 = True Then
                  If Text2 <> "Y" Then
                     If MsgBox("PDF 電子檔已存於" & m_strSavePath & vbCrLf & vbCrLf & "是否要開啟資料夾？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                        ShellExecute hLocalFile, "open", m_strSavePath, vbNullString, vbNullString, 1
                     End If
                  End If
               End If
               
            'Added by Morgan 2015/3/9
            ElseIf RsTemp.Fields("a1k33") = "2" Then
               If PdfSave4 = True Then
                  If Text2 <> "Y" Then
                     If MsgBox("PDF 電子檔已存於" & m_strSavePath & vbCrLf & vbCrLf & "是否要開啟資料夾？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                        ShellExecute hLocalFile, "open", m_strSavePath, vbNullString, vbNullString, 1
                     End If
                  End If
               End If
            'end 2015/3/9
            
            End If
         'end 2015/3/5
      
         'Added by Morgan 2016/9/30
         ElseIf m_RptNo = "5" Then
            If PdfSave2(True) = True Then
               If Text2 <> "Y" Then
                  If MsgBox("PDF, EXCEL 電子檔已存於" & m_strSavePath & vbCrLf & vbCrLf & "是否要開啟資料夾？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                     ShellExecute hLocalFile, "open", m_strSavePath, vbNullString, vbNullString, 1
                  End If
               Else
                  If MsgBox("EXCEL 電子檔已存於" & m_strSavePath & vbCrLf & vbCrLf & "是否要開啟資料夾？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                     ShellExecute hLocalFile, "open", m_strSavePath, vbNullString, vbNullString, 1
                  End If
               End If
            End If
         'end 2016/9/30
         
         'Added by Morgan 2019/8/19
         ElseIf m_RptNo = "7" Then
            If ExcelSave3 = True Then
               If MsgBox(stMsg & "EXCEL 電子檔已存於" & m_strSavePath & "，是否要開啟？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                  strExc(0) = Dir(m_strSavePath & "\" & m_FileName & ".*")
                  If strExc(0) <> "" Then
                     ShellExecute hLocalFile, "open", m_strSavePath & "\" & strExc(0), vbNullString, vbNullString, 1
                  End If
               End If
            End If
         
         'Added by Morgan 2022/5/13
         ElseIf m_RptNo = "9" Then
            stMsg = "" 'Added by Morgan 2023/12/20
            If PdfSave2() = True Then
               'Added by Morgan 2023/12/20
               If m_strNo = "X82995000" Then
                  If SaveLEDES = True Then
                     stMsg = "LEDES帳單電子檔已存於：" & vbCrLf & m_strSavePath & vbCrLf & vbCrLf
                  End If
               End If
               'end 2023/12/20
               
               If Text2 <> "Y" Then
                  If MsgBox("PDF 電子檔已存於" & m_strSavePath & vbCrLf & vbCrLf & "是否要開啟資料夾？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                     ShellExecute hLocalFile, "open", m_strSavePath, vbNullString, vbNullString, 1
                  End If
                  
               'Added by Morgan 2023/12/20
               ElseIf stMsg <> "" Then
                  'Modified by Morgan 2025/6/19
                  'MsgBox stMsg, vbInformation
                  If MsgBox(stMsg & vbCrLf & vbCrLf & "是否要開啟資料夾？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                     ShellExecute hLocalFile, "open", m_strSavePath, vbNullString, vbNullString, 1
                  End If
               'end 2023/12/20
               End If
            End If
            
         'Added by Morgan 2025/8/21
         ElseIf m_RptNo = "12" Then
            If ExcelSave5 = True Then
               If MsgBox(stMsg & "EXCEL 電子檔已存於" & m_strSavePath & "，是否要開啟？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                  strExc(0) = Dir(m_strSavePath & "\" & m_FileName & ".*")
                  If strExc(0) <> "" Then
                     ShellExecute hLocalFile, "open", m_strSavePath & "\" & strExc(0), vbNullString, vbNullString, 1
                  End If
               End If
            End If
         
         'Added by Morgan 2025/11/4
         ElseIf m_RptNo = "13" Then
            If PdfSave6() = True Then
               If Text2 <> "Y" Then
                  If MsgBox("PDF 電子檔已存於" & m_strSavePath & vbCrLf & vbCrLf & "是否要開啟資料夾？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                     ShellExecute hLocalFile, "open", m_strSavePath, vbNullString, vbNullString, 1
                  End If
               End If
            End If
         'end 2025/11/4
         End If
         
      Else
         MsgBox "無符合資料！"
      End If
   End If 'Added by Morgan 2015/3/2
   
   'Added by Morgan 2023/7/19
   If stRptNo2 <> "" Then
      m_RptNo = stRptNo2
      stRptNo2 = ""
      GoTo ReRun
   End If
   'end 2023/7/19
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub SetOutPutPath()
   Select Case Left(Pub_StrUserSt03, 2)
      Case "F1" 'FCT
         m_strSavePath = PUB_GetEFilePath("FCT") & "\Account"
      Case "F2" 'FCP
         m_strSavePath = PUB_GetEFilePath("FCP") & "\Account"
      Case "F3"
         m_strSavePath = PUB_GetEFilePath("FCL") & "\Account"
      Case Else
         m_strSavePath = PUB_Getdesktop
   End Select
   
   'Added by Morgan 2024/11/1
   If Dir(m_strSavePath, vbDirectory) = "" Then
      MkDir m_strSavePath
   End If
   'end 2024/11/1
   
   m_strSavePath = m_strSavePath & "\" & m_strNo
   
   If Dir(m_strSavePath, vbDirectory) = "" Then
      MkDir m_strSavePath
   End If
   
   m_strSavePath = m_strSavePath & "\" & Text3
   If Dir(m_strSavePath, vbDirectory) = "" Then
      MkDir m_strSavePath
   End If
   
End Sub

Private Function PdfSave() As Boolean
   Const cFontSize = 12
   Dim oTable As Word.Table
   Dim oShape As Word.Shape
   Dim dblAFee As Double, dblOFee As Double, dblDFee As Double, dblTFee As Double
   Dim iRow As Integer
   Dim stFileName As String
   Dim stPdfName As String, stFullPath As String
   Dim rsReprot As ADODB.Recordset
   
   
On Error GoTo ErrHnd
   
   '表頭
   If Left(m_strNo, 1) = "Y" Then
      strExc(0) = "select fa05,fa63,fa64,fa65,fa32,fa33,fa34,fa35,fa36,fa18,fa19,fa20,fa21,fa22,fa70,fa17,fa23" & _
         " from fagent where fa01='" & Left(m_strNo, 8) & "' and fa02='" & Mid(m_strNo, 9) & "'"
   Else
      strExc(0) = "select cu05 as fa05,cu88 as fa63,cu89 as fa64,cu90 as fa65,cu65 as fa32, cu66 as fa33, cu67 as fa34" & _
         ", cu68 as fa35, cu69 as fa36, cu24 as fa18, cu25 as fa19, cu26 as fa20, cu27 as fa21, cu28 as fa22,cu102 fa70" & _
         ", cu23 as fa17, cu29 as fa23 from customer where cu01='" & Left(m_strNo, 8) & "' and cu02='" & Mid(m_strNo, 9) & "'"
   End If
   intI = 1
   Set rsReprot = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      MsgBox "表頭資料讀取失敗!!"
      Exit Function
   End If
   
   With rsReprot
   '代理人名稱 strexc(1)
   strExc(1) = "" & .Fields("fa05")
   If Not IsNull(.Fields("fa63")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa63")
   End If
   If Not IsNull(.Fields("fa64")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa64")
   End If
   If Not IsNull(.Fields("fa65")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa65")
   End If
   '代理人POBox/地址
   If Not IsNull(.Fields("fa32")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa32")
      If Not IsNull(.Fields("fa33")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa33")
      End If
      If Not IsNull(.Fields("fa34")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa34")
      End If
      If Not IsNull(.Fields("fa35")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa35")
      End If
      If Not IsNull(.Fields("fa36")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa36")
      End If
   ElseIf Not IsNull(.Fields("fa18")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa18")
      If Not IsNull(.Fields("fa19")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa19")
      End If
      If Not IsNull(.Fields("fa20")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa20")
      End If
      If Not IsNull(.Fields("fa21")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa21")
      End If
      If Not IsNull(.Fields("fa22")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa22")
      End If
      If Not IsNull(.Fields("fa70")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa70")
      End If
   End If
   End With
   
   '項目中文雜費者加總
   strExc(0) = "select a1k01,a1k13,to_char(to_date(a1k02+19110000,'yyyymmdd'),'FMMM/DD/yyyy') dt" & _
      ",tm45||pa77||lc23||sp27 YrRef,a1k13||'-'||a1k14||decode(a1k15||a1k16,'000','','-'||a1k16||'-'||a1k17) OrRef,X.*" & _
      " from (select a1l01,sum(amt)-sum(decode(substrb(a1l04,-2),'99',amt,0))-sum(decode(a1j03,'雜費',amt,0)) AFee" & _
      ",sum(decode(substrb(a1l04,-2),'99',amt,0)) OFee,sum(decode(a1j03,'雜費',amt,0)) DFee,sum(amt) TFee" & _
      " from (select a.a1l01,a.a1l03,a.a1l04" & _
      ",decode( nvl(a.a1l17,0), 0, trunc((a.a1l05-nvl(a.a1l07,0)+nvl(b.a1l05,0)-nvl(b.a1l07,0))/a1k10)" & _
      ",trunc((a.a1l17+nvl(b.a1l17,0))* round(1-nvl(a.a1l07,0)/a.a1l05,2))) Amt,a1j03" & _
      " from acc1k0,acc1l0 a,acc1l0 b,acc1j0" & _
      " where nvl(a1k12,0)=0 and a1k25||a1k29 is null and a1k02>=" & Text3 & "01 and a1k02<=" & Text3 & "31 and a1k28='" & m_strNo & "'" & _
      " and a.a1l01(+)=a1k01 and substr(a.a1l04(+),-2)<>'98' and b.a1l01(+)=a.a1l01 and b.a1l03(+)=a.a1l03 and b.a1l04(+)=a.a1l04||'98'" & _
      " and a1j01(+)=a.a1l03 and a1j02(+)=a.a1l04" & _
      ") group by a1l01) X,acc1k0,trademark,patent,lawcase,servicepractice where a1k01(+)=a1l01" & _
      " and tm01(+)=a1k13 and tm02(+)=a1k14 and tm03(+)=a1k15 and tm04(+)=a1k16" & _
      " and pa01(+)=a1k13 and pa02(+)=a1k14 and pa03(+)=a1k15 and pa04(+)=a1k16" & _
      " and sp01(+)=a1k13 and sp02(+)=a1k14 and sp03(+)=a1k15 and sp04(+)=a1k16" & _
      " and lc01(+)=a1k13 and lc02(+)=a1k14 and lc03(+)=a1k15 and lc04(+)=a1k16" & _
      " order by 1,2"
   intI = 1
   Set rsReprot = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      MsgBox "請款明細資料讀取失敗!!"
      Exit Function
   End If
      
   '請款單號 strexc(2)
   strExc(2) = rsReprot("a1k01") & "/" & Text3
   
   If NewWordDoc = False Then Exit Function
   
   With g_WordAp.Application
      
      .Selection.Font.Name = "Times New Roman"
      .Selection.Font.Size = cFontSize
      
      '版面設定
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.5)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.3)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
      .Selection.PageSetup.FooterDistance = .CentimetersToPoints(3)
      .Selection.PageSetup.CharsLine = 40
      .Selection.PageSetup.LinesPage = 38
      .Selection.Orientation = wdTextOrientationHorizontal
      
      '信頭尾
      If PUB_ReadDB2File(stFileName, iPicNo) = True Then
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
         Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
         oShape.ZOrder 4
         oShape.LockAnchor = True
         oShape.LockAspectRatio = -1
         oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
         oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
         oShape.Left = 0
         oShape.Top = 0
         oShape.Width = .CentimetersToPoints(21)
         oShape.WrapFormat.Type = wdWrapNone
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         
         If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
            .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
            oShape.ZOrder 4
            oShape.LockAnchor = True
            oShape.LockAspectRatio = -1
            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
            oShape.Left = 0
            oShape.Top = .CentimetersToPoints(27)
            oShape.Width = .CentimetersToPoints(21)
            oShape.WrapFormat.Type = wdWrapNone
            .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         End If
         
         .Selection.HomeKey Unit:=wdStory
      End If
      
      
      .Selection.TypeParagraph
      '行距
      With .Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .DisableLineHeightGrid = True
      End With
      
      '新增表格(1*2)
      Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumRows:=1, NumColumns:=2)
      With oTable
         '無邊框
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
      End With
            
      oTable.Select
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop '靠上對齊
      .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
      .Selection.InsertRows 8

      '代理人名稱,POBox/地址
      oTable.Cell(1, 1).Merge oTable.Cell(4, 1)
      oTable.Cell(1, 1).Select
      .Selection.Text = strExc(1)
      
      '月份
      strExc(0) = "Month: " & Format(ChangeTStringToWDateString(Text3 & "01"), "mmmm, yyyy")
      oTable.Cell(1, 2).Select
      .Selection.Text = strExc(0)
      
      '請款單號
      strExc(0) = "Invoice No: " & strExc(2)
      oTable.Cell(2, 2).Select
      .Selection.Text = strExc(0)
      
      'Purchase Order
      strExc(0) = "Purchase Order: " & Text1
      oTable.Cell(3, 2).Select
      .Selection.Text = strExc(0)
      
      oTable.Cell(5, 1).Merge oTable.Cell(5, 2)
      oTable.Cell(5, 1).Select
      .Selection.Cells(1).SetHeight RowHeight:=30, HeightRule:=wdRowHeightAtLeast
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      .Selection.Text = "Collective Invoice"
      
      oTable.Cell(6, 1).Select
      .Selection.SelectRow
      
      With .Selection.Cells
        '有邊框
        .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
        .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
      End With
      
      .Selection.Cells.Split NumRows:=1, NumColumns:=8, MergeBeforeSplit:=True
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      .Selection.Font.Size = 10
      '設定表格高度欄寬
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.7), RulerStyle:=wdAdjustProportional
      .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.4), RulerStyle:=wdAdjustProportional
      .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(4.8), RulerStyle:=wdAdjustProportional
      .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
      .Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(1.7), RulerStyle:=wdAdjustProportional
      .Selection.Cells(6).SetWidth ColumnWidth:=.CentimetersToPoints(1.7), RulerStyle:=wdAdjustProportional
      .Selection.Cells(7).SetWidth ColumnWidth:=.CentimetersToPoints(2.3), RulerStyle:=wdAdjustProportional
      
      .Selection.Cells(1).SetHeight RowHeight:=16, HeightRule:=wdRowHeightAtLeast
      .Selection.InsertRows rsReprot.RecordCount + 1
      
      oTable.Cell(6, 1).Select
      .Selection.SelectRow
      .Selection.Font.Bold = True
      
      oTable.Cell(6, 1).Select
      .Selection.Text = "No."
      oTable.Cell(6, 2).Select
      .Selection.Text = "Invoice Date" & vbCrLf & "<mm/dd/yyy>"
      oTable.Cell(6, 3).Select
      .Selection.Text = "Your Ref"
      oTable.Cell(6, 4).Select
      .Selection.Text = "Our Ref"
      oTable.Cell(6, 5).Select
      .Selection.Text = "Attorney" & vbCrLf & "Fee" & vbCrLf & "(USD)"
      oTable.Cell(6, 6).Select
      .Selection.Text = "Official" & vbCrLf & "Fee" & vbCrLf & "(USD)"
      oTable.Cell(6, 7).Select
      .Selection.Text = "Disbursement" & vbCrLf & "Fee" & vbCrLf & "(USD)"
      oTable.Cell(6, 8).Select
      .Selection.Text = "Total Fee" & vbCrLf & "(USD)"
      .Selection.SelectRow
      .Selection.Cells.Shading.Texture = wdTextureNone
      '.Selection.Cells.Shading.BackgroundPatternColorIndex = wdTurquoise
      .Selection.Cells.Shading.Texture = wdTexture5Percent
      .Selection.Cells(1).SetHeight RowHeight:=36, HeightRule:=wdRowHeightAtLeast
      
      iRow = 6
      Do While Not rsReprot.EOF
         iRow = iRow + 1
         oTable.Cell(iRow, 1).Select
         .Selection.Text = iRow - 6
         oTable.Cell(iRow, 2).Select
         .Selection.Text = "" & rsReprot("dt")
         '彼所案號
         strExc(1) = "" & rsReprot("YrRef")
         If GetXYrRef(rsReprot("a1k13"), rsReprot("A1K01"), strExc(2)) = True Then
            strExc(1) = strExc(2)
         End If
         oTable.Cell(iRow, 3).Select
         .Selection.Text = strExc(1)
         oTable.Cell(iRow, 4).Select
         .Selection.Text = "" & rsReprot("OrRef")
         oTable.Cell(iRow, 5).Select
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Text = Format(Val("" & rsReprot("AFee")), cfmtDollar)
         oTable.Cell(iRow, 6).Select
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Text = Format(Val("" & rsReprot("OFee")), cfmtDollar)
         oTable.Cell(iRow, 7).Select
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Text = Format(Val("" & rsReprot("DFee")), cfmtDollar)
         oTable.Cell(iRow, 8).Select
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Text = Format(Val("" & rsReprot("TFee")), cfmtDollar)
         
         dblAFee = dblAFee + Val("" & rsReprot("AFee"))
         dblOFee = dblOFee + Val("" & rsReprot("OFee"))
         dblDFee = dblDFee + Val("" & rsReprot("DFee"))
         dblTFee = dblTFee + Val("" & rsReprot("TFee"))
         rsReprot.MoveNext
      Loop
      
      iRow = iRow + 1
      oTable.Cell(iRow, 1).Merge oTable.Cell(iRow, 4)
      oTable.Cell(iRow, 1).Select
      .Selection.SelectRow
      .Selection.Font.Bold = True
      
      oTable.Cell(iRow, 1).Select
      .Selection.Text = "Total"
      oTable.Cell(iRow, 2).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.Text = Format(dblAFee, cfmtDollar)
      oTable.Cell(iRow, 3).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.Text = Format(dblOFee, cfmtDollar)
      oTable.Cell(iRow, 4).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.Text = Format(dblDFee, cfmtDollar)
      oTable.Cell(iRow, 5).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.Text = Format(dblTFee, cfmtDollar)
      
      '帳號
      iRow = iRow + 2
      oTable.Cell(iRow, 1).Select
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
      .Selection.Text = vbCrLf & ReportSum(71001) & vbCrLf & ReportSum(72) & vbCrLf & ReportSum(73001) & vbCrLf & ReportSum(85) & vbCrLf & ReportSum(74) & vbCrLf & ReportSum(121) & vbCrLf
      
      '建議電匯提醒
      oTable.Cell(iRow, 2).Select
      .Selection.Cells.Split NumRows:=3, NumColumns:=1, MergeBeforeSplit:=False
      .Selection.Cells(1).SetHeight RowHeight:=28, HeightRule:=wdRowHeightAtLeast
      iRow = iRow + 1
      oTable.Cell(iRow, 2).Select
      .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
      With .Selection.Cells(1)
          With .Borders(wdBorderLeft)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderRight)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderTop)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderBottom)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
      End With
      .Selection.ParagraphFormat.LeftIndent = .CentimetersToPoints(0.2)
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
      .Selection.Cells(1).SetHeight RowHeight:=52, HeightRule:=wdRowHeightAtLeast
      .Selection.Text = "Wire" & vbCrLf & "Transfer" & vbCrLf & "Preferred"
      iRow = iRow + 1
      oTable.Cell(iRow, 2).Select
      .Selection.Cells(1).SetHeight RowHeight:=0, HeightRule:=wdRowHeightAtLeast
      
      '備註
      iRow = iRow + 1
      oTable.Cell(iRow, 1).Select
      .Selection.SelectRow
      .Selection.Font.Bold = True
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
      oTable.Cell(iRow, 1).Select
      .Selection.Text = "PS:"
      oTable.Cell(iRow, 2).Select
      .Selection.Text = "Please return a copy of the invoice(s) or indicate the invoice number(s) paid with remittance"
      
      If Text2 = "Y" Then
         .Activate
      Else
         stPdfName = m_strNo & Text3 & ".pdf"
         
'Modified by Morgan 2022/6/8
'         If Pub_GetPrinterIndex("PDFCreator") < 0 Then
'            MsgBox "請先安裝 PDFCreator 印表機，pdf 轉檔失敗！", vbExclamation
'            Exit Function
'         Else
'            stPdfName = m_strNo & Text3 & ".pdf"
'            pub_OsPrinter = PUB_GetOsDefaultPrinter
'            frmPDF.Show
'            frmPDF.StartProcess m_strSavePath, stPdfName
'            PUB_SetOsDefaultPrinter Printer.DeviceName
'            PUB_SetWordActivePrinter
'            .ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
'            .ActiveDocument.Close wdDoNotSaveChanges
'            .Quit wdDoNotSaveChanges
'            frmPDF.EndtProcess
'            Unload frmPDF
'            PUB_SetOsDefaultPrinter pub_OsPrinter
'         End If
         .ActiveDocument.ExportAsFixedFormat OutputFileName:=m_strSavePath & "\" & stPdfName, ExportFormat:=17, OpenAfterExport:=False
'end 2022/6/8
      End If
   End With
   PdfSave = True
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   Set rsReprot = Nothing
   
End Function

Private Function NewWordDoc() As Boolean
Dim iResumeCnt As Integer
   
On Error GoTo ErrHnd

   Set g_WordAp = New Word.Application
   g_WordAp.Documents.add
   g_WordAp.Visible = True
   'g_WordAp.WindowState = wdWindowStateMaximize
   'g_WordAp.ActiveWindow.ActivePane.View.Zoom.Percentage = 100
   NewWordDoc = True
   
ErrHnd:
   If Err.Number <> 0 Then
      If iResumeCnt > 3 Then
         MsgBox "錯誤 : " & Err.Description, vbCritical
      Else
         iResumeCnt = iResumeCnt + 1
         Select Case Err.Number
            Case 91:
               g_WordAp.Documents.add
               Resume Next
            Case 462:
               Set g_WordAp = New Word.Application
               Resume
            Case Else:
               MsgBox "錯誤 : " & Err.Description, vbCritical
         End Select
      End If
   End If
End Function

Private Sub Form_Load()

Dim stPList As String, stTList As String, stList As String, arrList() As String

PUB_InitForm Me, Me.Width, Me.Height

'Modified by Morgan 2016/7/15 +Y34412000 --Lina
'Modified by Morgan 2016/11/14 +Y52418000 --Lina
'Modified by Morgan 2019/8/19 +Y54570--Lisa
'Modified by Morgan 2020/9/21 +Y51971--Franny
'Modified by Morgan 2020/10/5 +Y22327--Lisa
'Modified by Morgan 2021/3/17 +Y20412010 --Franny
'Modified by Morgan 2022/2/7 +Y55666 NOVOCURE GMBH --Ryan
'Modified by Morgan 2022/6/7 +Y55751 --Ryan
'Modified by Morgan 2022/8/5 -Y55751 --Franny
'Modified by Morgan 2023/3/8 +Y55240 --Kahn
'Modified by Morgan 2023/7/10 +Y48279000 --Kahn
'Modified by Morgan 2023/8/1 +Y51467020 --Franny
'Modified by Morgan 2023/12/20 +X82995000 --Kahn
'Modified by Morgan 2024/5/23 +Y55973000--Izumi
'Modified by Morgan 2024/8/30 +Y56066 Harrity & Harrity LLP 及 Y55751 Birkenstock IP GmbH--Izumi
'Modified by Morgan 2024/9/3 +Y34126000 L'Air liquide --Anny
'Modified by Morgan 2025/4/24 +Y34049010 --Anny
'Modified by Morgan 2025/6/30 +Y22457000--Tim
'Modified by Morgan 2025/8/20 +Y56199000--Kahn
stPList = "Y54225B10,Y21431000,Y53280000,Y45493000,Y34412010,Y52418000,Y54570000,Y51971000,Y22327000,Y20412010,Y55666000,Y55240000,Y48279000(X48279000),Y51467020,X82995000,Y55973000,Y56066000,Y55751000,Y34126000,Y34049010,Y22457000,Y56199000"
'Modified by Morgan 2018/11/6 +Y52341000 --Monica
'Modified by Morgan 2024/8/30 -Y48309070 未再使用
stTList = "Y52341000"
Combo1.Clear
'外專
If Left(Pub_StrUserSt03, 2) = "F2" Then
   stList = stPList
'外商
ElseIf Left(Pub_StrUserSt03, 2) = "F1" Then
   stList = stTList
   
ElseIf Pub_StrUserSt03 = "M51" Then
   stList = stPList & "," & stTList
   stList = stList & ",Y48292000" 'Added by Morgan 2022/8/18 單次客製化使用
End If

arrList = Split(stList, ",")
For intI = LBound(arrList) To UBound(arrList)
   If arrList(intI) <> "" Then
      If Left(arrList(intI), 1) = "Y" Then
         Combo1.AddItem arrList(intI) & " " & GetPrjName2(Left(arrList(intI), 9))
      Else
         Combo1.AddItem arrList(intI) & " " & GetPrjPeople1(Left(arrList(intI), 9), 2)
      End If
   End If
Next
If Combo1.ListCount = 1 Then
   Combo1.ListIndex = 0
Else
   Combo1.ListIndex = -1
End If
'Added by Morgan 2020/3/30
If strSrvDate(1) >= 智慧所更名日 Then
   iPicNo = 68
   iPicNo2 = 69
Else
   iPicNo = 5
   iPicNo2 = 9
End If 'Added by Morgan 2020/3/30
End Sub

Private Sub Form_Unload(Cancel As Integer)
strFormName = MsgText(601)
Set Frmacc24l0 = Nothing
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
   CloseIme
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
   CloseIme
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text3 <> "" Then
      If ChkDate(Text3 & "01") = False Then
         Text3.SetFocus
         Cancel = True
      'Added by Morgan 2021/2/22
      Else
         If Text4.Enabled = False Then
            Text4 = Text3
         Else
            If Text4 = "" Then Text4 = Text3
         End If
      'end 2021/2/22
      End If
   End If
End Sub

Private Function ExcelSave2(pRst As ADODB.Recordset) As Boolean
   Dim xlsReport As New Excel.Application
   Dim wksReport As New Worksheet
   Dim stFullPath As String
   Dim ii As Integer, dblTFee As Double
   Dim bolInvDate As Boolean, stCol As String, iCols As Integer, jj As Integer 'Added by Morgan 2017/5/10
   
On Error GoTo ErrHnd

   If m_strNo = "Y52418000" Then bolInvDate = True 'Added by Morgan 2017/5/10 Y52418000 OMYA調整格式 --Lina
   If m_strNo = "Y55666000" Then bolInvDate = True 'Added by Morgan 2022/2/7
   
   xlsReport.Visible = True
   With pRst
   .MoveFirst
   xlsReport.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsReport.Workbooks.add
   Set wksReport = xlsReport.Worksheets(1)
   
   wksReport.Cells.NumberFormatLocal = "@"
   ii = 1
   wksReport.Range("A" & ii) = "No."
   wksReport.Range("B" & ii) = "Invoice No."
   stCol = "B"
   'Added by Morgan 2017/5/10
   If bolInvDate Then
      stCol = Chr(Asc(stCol) + 1)
      wksReport.Range(stCol & ii) = "Invoice Date"
   End If
   'end 2017/5/10
   stCol = Chr(Asc(stCol) + 1)
   wksReport.Range(stCol & ii) = "Your Ref"
   stCol = Chr(Asc(stCol) + 1)
   wksReport.Range(stCol & ii) = "Our Ref"
   stCol = Chr(Asc(stCol) + 1)
   wksReport.Range(stCol & ii) = "Description"
   stCol = Chr(Asc(stCol) + 1)
   wksReport.Range(stCol & ii) = "Total Fee(USD)"
   
   wksReport.Range("A" & ii, stCol & ii).Font.Bold = True
   
   iCols = Asc(stCol) - Asc("A") + 1 'Added by Morgan 2017/5/10
   Do While Not .EOF
      ii = ii + 1
      wksReport.Range("A" & ii) = ii - 1
      wksReport.Range("A" & ii).NumberFormatLocal = "G/通用格式"
      wksReport.Range("B" & ii) = "" & .Fields("a1k01")
      stCol = "B"
      'Added by Morgan 2017/5/10
      If bolInvDate Then
         stCol = Chr(Asc(stCol) + 1)
         wksReport.Range(stCol & ii) = "" & .Fields("dt")
      End If
      'end 2017/5/10
      
      strExc(1) = "" & .Fields("YrRef")
      If GetXYrRef(.Fields("a1k13"), .Fields("A1K01"), strExc(2)) = True Then
         strExc(1) = strExc(2)
      End If
      stCol = Chr(Asc(stCol) + 1)
      wksReport.Range(stCol & ii) = strExc(1)
      stCol = Chr(Asc(stCol) + 1)
      wksReport.Range(stCol & ii) = "" & .Fields("OrRef")
      stCol = Chr(Asc(stCol) + 1)
      wksReport.Range(stCol & ii) = "" & .Fields("IDesc")
      stCol = Chr(Asc(stCol) + 1)
      wksReport.Range(stCol & ii) = Val("" & .Fields("TFee"))
      wksReport.Range(stCol & ii).NumberFormatLocal = "#,##0.00_ "
      dblTFee = dblTFee + Val("" & .Fields("TFee"))
      .MoveNext
   Loop
   End With
      
   For jj = 0 To iCols - 1
      stCol = Chr(Asc("A") + jj)
      wksReport.Columns(stCol & ":" & stCol).EntireColumn.AutoFit
      If wksReport.Range(stCol & "1").ColumnWidth > 80 Then
         wksReport.Range(stCol & "1").ColumnWidth = 80
      End If
   Next
    
   ii = ii + 1
   stCol = Chr(Asc("A") + iCols - 2)
   wksReport.Range("A" & ii) = "Total"
   wksReport.Range("A" & ii, stCol & ii).Merge
   wksReport.Range("A" & ii, stCol & ii).HorizontalAlignment = xlCenter
   
   stCol = Chr(Asc(stCol) + 1)
   wksReport.Range(stCol & ii) = dblTFee
   wksReport.Range(stCol & ii).NumberFormatLocal = "#,##0.00_ "
   wksReport.Range("A" & ii, stCol & ii).Font.Bold = True

   xlsReport.Range("A1").Select
      
   stFullPath = m_strSavePath & "\" & m_FileName
   If Dir(stFullPath & ".*") <> "" Then
      Kill stFullPath & ".*"
   End If
   xlsReport.Workbooks(1).SaveAs stFullPath
   xlsReport.Workbooks.Close
   xlsReport.Quit
   
   ExcelSave2 = True
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
ExitPoint:
   Set xlsReport = Nothing
   
End Function

'Added by Morgan 2019/8/19
Private Function ExcelSave3() As Boolean
   Dim rsReprot As ADODB.Recordset
   Dim xlsReport As New Excel.Application
   Dim wksReport As New Worksheet
   Dim strFaNo As String, iSheet As Integer
   Dim stFullPath As String
   Dim ii As Integer, dblTFee As Double
   Dim stCol As String, iCols As Integer, jj As Integer
   
On Error GoTo ErrHnd
   
   '表頭
   strExc(0) = "select fa05,fa63,fa64,fa65,fa32,fa33,fa34,fa35,fa36,fa18,fa19,fa20,fa21,fa22,fa70,fa17,fa23" & _
      " from fagent where fa01='" & Left(m_strNo, 8) & "' and fa02='" & Mid(m_strNo, 9) & "'"
   intI = 1
   Set rsReprot = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      MsgBox "表頭資料讀取失敗!!"
      Exit Function
   End If
   
   With rsReprot
   '請款對象名稱 strexc(1)
   strExc(1) = "" & .Fields("fa05")
   If Not IsNull(.Fields("fa63")) Then
      strExc(1) = strExc(1) & " " & .Fields("fa63")
   End If
   If Not IsNull(.Fields("fa64")) Then
      strExc(1) = strExc(1) & " " & .Fields("fa64")
   End If
   If Not IsNull(.Fields("fa65")) Then
      strExc(1) = strExc(1) & " " & .Fields("fa65")
   End If
   '請款對象POBox/地址 strexc(2)
   strExc(2) = ""
   If Not IsNull(.Fields("fa32")) Then
      strExc(2) = .Fields("fa32")
      If Not IsNull(.Fields("fa33")) Then
         strExc(2) = strExc(2) & vbCrLf & .Fields("fa33")
      End If
      If Not IsNull(.Fields("fa34")) Then
         strExc(2) = strExc(2) & vbCrLf & .Fields("fa34")
      End If
      If Not IsNull(.Fields("fa35")) Then
         strExc(2) = strExc(2) & vbCrLf & .Fields("fa35")
      End If
      If Not IsNull(.Fields("fa36")) Then
         strExc(2) = strExc(2) & vbCrLf & .Fields("fa36")
      End If
   ElseIf Not IsNull(.Fields("fa18")) Then
      strExc(2) = .Fields("fa18")
      If Not IsNull(.Fields("fa19")) Then
         strExc(2) = strExc(2) & vbCrLf & .Fields("fa19")
      End If
      If Not IsNull(.Fields("fa20")) Then
         strExc(2) = strExc(2) & vbCrLf & .Fields("fa20")
      End If
      If Not IsNull(.Fields("fa21")) Then
         strExc(2) = strExc(2) & vbCrLf & .Fields("fa21")
      End If
      If Not IsNull(.Fields("fa22")) Then
         strExc(2) = strExc(2) & vbCrLf & .Fields("fa22")
      End If
      If Not IsNull(.Fields("fa70")) Then
         strExc(2) = strExc(2) & vbCrLf & .Fields("fa70")
      End If
   End If
   End With
   
   '明細(要依列印格式計算)
   'Modified by Morgan 2020/5/5 +X71102,X81780
   'Modified by Morgan 2024/11/1 不必再限制申請人條件--Lisa
   '   " and (instr(pa26||pa27||pa28||pa29||pa30||sp08||sp58||sp59||sp65||sp66,'X80691000')>0" & _
      " or instr(pa26||pa27||pa28||pa29||pa30||sp08||sp58||sp59||sp65||sp66,'X80691C10')>0" & _
      " or instr(pa26||pa27||pa28||pa29||pa30||sp08||sp58||sp59||sp65||sp66,'X71102000')>0" & _
      " or instr(pa26||pa27||pa28||pa29||pa30||sp08||sp58||sp59||sp65||sp66,'X81780000')>0)"
   'Modified by Morgan 2025/4/10 輸入幣別a1l16為台幣NTD時也要換算
   strExc(0) = "select a1k03,to_char(to_date(a1k02+19110000,'yyyymmdd'),'FMMM/DD/yyyy') dt,a1k01" & _
      ",a1k13||'-'||a1k14||decode(a1k15||a1k16,'000','','-'||a1k16||'-'||a1k17) OrRef,pa48||sp29 CuRef" & _
      ",pa11||sp11 AppNo,a1k18,AFee,OFee,AFee+OFee TFee,fa05,fa63,fa64,fa65" & _
      " from (select a1l01,a1k33,a1k10" & _
      ",trunc(sum(decode(substr(a1l04,-2),'99',0,decode(a1k33,'3',trunc(decode(NTD,'Y',amt/a1k10,amt)),decode(NTD,'Y',amt/a1k10,amt)) ))) AFee" & _
      ",trunc(sum(decode(substr(a1l04,-2),'99',decode(a1k33,'3',trunc(decode(NTD,'Y',amt/a1k10,amt)),decode(NTD,'Y',amt/a1k10,amt)),0 ))) OFee" & _
      " from (select a.a1l01,a.a1l03,a.a1l04,a1k10,a1k33" & _
      ",decode(nvl(a.a1l17,0),0, a.a1l05-nvl(a.a1l07,0)+nvl(b.a1l05,0)-nvl(b.a1l07,0)" & _
      ",trunc((a.a1l17+nvl(b.a1l17,0))* round(1-nvl(a.a1l07,0)/a.a1l05,2))) Amt" & _
      ",decode(a.a1l16,'NTD','Y',decode(nvl(a.a1l17,0),0,'Y')) NTD" & _
      " from acc1k0,acc1l0 a,acc1l0 b" & _
      " where nvl(a1k12,0)=0 and a1k25||a1k29 is null and a1k02>=" & Text3 & "01 and a1k02<=" & Text3 & "31 and a1k28='" & m_strNo & "'" & _
      " and a.a1l01(+)=a1k01 and substr(a.a1l04(+),-2)<>'98' and b.a1l01(+)=a.a1l01 and b.a1l03(+)=a.a1l03 and b.a1l04(+)=a.a1l04||'98'" & _
      " and a1k13 in ('P','PS','CFP','CPS','FCP','FG')) group by a1l01,a1k33,a1k10,NTD) X,acc1k0,patent,servicepractice,fagent where a1k01(+)=a1l01" & _
      " and pa01(+)=a1k13 and pa02(+)=a1k14 and pa03(+)=a1k15 and pa04(+)=a1k16" & _
      " and sp01(+)=a1k13 and sp02(+)=a1k14 and sp03(+)=a1k15 and sp04(+)=a1k16" & _
      " and fa01(+)=substr(a1k03,1,8) and fa02(+)=substr(a1k03,9) order by 1,2,3"
   intI = 1
   Set rsReprot = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      MsgBox "請款明細資料讀取失敗!!"
      Exit Function
   End If
   
   xlsReport.Visible = True
   
   With rsReprot
   .MoveFirst
   xlsReport.SheetsInNewWorkbook = 1
   xlsReport.Workbooks.add
   iSheet = 0
   strFaNo = ""
   Do While Not .EOF
      
      If strFaNo <> .Fields("a1k03") Then
      
         '合計
         If strFaNo <> "" Then
            ii = ii + 1
            wksReport.Range("C" & ii) = "Total"
            wksReport.Range("J" & ii) = dblTFee
            wksReport.Range("J" & ii).NumberFormatLocal = "#,##0.00_ "
            wksReport.Range("A" & ii, "J" & ii).Font.Bold = True
            '置中
            wksReport.Range("A6", "G" & ii).HorizontalAlignment = xlCenter
            '加邊框
            AddBorder wksReport, "A6", "J" & ii
         End If
         dblTFee = 0
         
         iSheet = iSheet + 1
         If xlsReport.Worksheets.Count >= iSheet Then
            Set wksReport = xlsReport.Worksheets(iSheet)
         Else
            Set wksReport = xlsReport.Worksheets.add(After:=xlsReport.Worksheets(xlsReport.Worksheets.Count))
         End If
         
         '自動放寬有問題改自行設定
         wksReport.Columns("A:A").ColumnWidth = 6
         wksReport.Columns("B:B").ColumnWidth = 25
         wksReport.Columns("C:C").ColumnWidth = 12
         wksReport.Columns("D:D").ColumnWidth = 12
         wksReport.Columns("E:E").ColumnWidth = 20
         wksReport.Columns("F:F").ColumnWidth = 16
         wksReport.Columns("G:G").ColumnWidth = 12
         wksReport.Columns("H:H").ColumnWidth = 12
         wksReport.Columns("I:I").ColumnWidth = 20
         wksReport.Columns("J:J").ColumnWidth = 12
         
         
         strFaNo = .Fields("a1k03")
         
         wksReport.Cells.NumberFormatLocal = "@"
         wksReport.Range("B1") = strExc(1) '請款對象名稱
         wksReport.Range("B2") = strExc(2) '請款對象POX/地址
         'wksReport.Range("B2").WrapText = False
         'wksReport.Columns("B").EntireColumn.AutoFit '自動放寬
         'wksReport.Range("B2").WrapText = True
         'wksReport.Columns("B").EntireColumn.AutoFit '自動放寬
         wksReport.Rows("2:2").EntireRow.AutoFit
         
         '代理人名稱 strexc(3)
         strExc(3) = "" & .Fields("fa05")
         If Not IsNull(.Fields("fa63")) Then
            strExc(3) = strExc(3) & " " & .Fields("fa63")
         End If
         If Not IsNull(.Fields("fa64")) Then
            strExc(3) = strExc(3) & " " & .Fields("fa64")
         End If
         If Not IsNull(.Fields("fa65")) Then
            strExc(3) = strExc(3) & " " & .Fields("fa65")
         End If
         
         wksReport.Name = strExc(3)
         wksReport.Range("B4") = "(US Hub Firm: " & strExc(3) & ")" '代理人名稱
         wksReport.Range("B4").WrapText = False
         wksReport.Range("B1", "B4").Font.Bold = True
         
         '請款日期起迄
         wksReport.Range("I5") = ChangeTStringToWDateString(Text3 & "01") & "~" & ChangeWStringToWDateString(GetLastDay(Text3 & "01"))
         wksReport.Columns("I").EntireColumn.AutoFit '自動放寬
         
         '欄位名稱
         '置中
         With wksReport.Range("A6", "J6")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
         End With
         
         '底色
         With wksReport.Range("A6", "J6").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 14408946
            .tintandshade = 0
            .PatternTintAndShade = 0
         End With
         
         wksReport.Range("A6") = "No."
         wksReport.Range("B6") = "Invoice Date" & vbCrLf & "<m/d/yyyy>"
         wksReport.Range("C6") = "Invoice No."
         wksReport.Range("D6") = "Our Ref."
         wksReport.Range("E6") = "Case No. (If any)"
         wksReport.Range("F6") = "Patent" & vbCrLf & "Application No."
         wksReport.Range("G6") = "Currency"
         wksReport.Range("H6") = "Attorney Fee"
         wksReport.Range("I6") = "Official Fee"
         wksReport.Range("J6") = "Total Fee in USD"
         wksReport.Range("A6", "J6").Font.Bold = True
         wksReport.Range("A6", "J6").WrapText = True
         
         ii = 6
      End If
      ii = ii + 1
      wksReport.Range("A" & ii) = ii - 6
      wksReport.Range("A" & ii).NumberFormatLocal = "G/通用格式"
      wksReport.Range("B" & ii) = "" & .Fields("dt")
      wksReport.Range("B" & ii).HorizontalAlignment = xlCenter
      wksReport.Range("B" & ii).VerticalAlignment = xlCenter
      wksReport.Range("C" & ii) = "" & .Fields("a1k01")
      wksReport.Range("D" & ii) = "" & .Fields("OrRef")
      wksReport.Range("E" & ii) = "" & .Fields("CuRef")
      wksReport.Range("F" & ii) = "" & .Fields("AppNo")
      wksReport.Range("G" & ii) = "" & .Fields("a1k18")
      wksReport.Range("H" & ii) = Val("" & .Fields("AFee"))
      wksReport.Range("I" & ii) = Val("" & .Fields("OFee"))
      wksReport.Range("J" & ii) = Val("" & .Fields("TFee"))
      wksReport.Range("H" & ii, "J" & ii).NumberFormatLocal = "#,##0.00_ "
      dblTFee = dblTFee + Val("" & .Fields("TFee"))
      .MoveNext
   Loop
   End With
    
   ii = ii + 1
   wksReport.Range("C" & ii) = "Total"
   wksReport.Range("J" & ii) = dblTFee
   wksReport.Range("J" & ii).NumberFormatLocal = "#,##0.00_ "
   wksReport.Range("A" & ii, "J" & ii).Font.Bold = True
   '置中
   wksReport.Range("A6", "G" & ii).HorizontalAlignment = xlCenter
   '加邊框
   AddBorder wksReport, "A6", "J" & ii
   '切回第一表
   xlsReport.Sheets(1).Select
   xlsReport.Range("A1").Select
   
   stFullPath = m_strSavePath & "\" & m_FileName
   If Dir(stFullPath & ".*") <> "" Then
      Kill stFullPath & ".*"
   End If
   xlsReport.Workbooks(1).SaveAs stFullPath
   xlsReport.Workbooks.Close
   xlsReport.Quit
   
   ExcelSave3 = True
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
ExitPoint:
   Set xlsReport = Nothing
   
End Function

'Added by Morgan 2025/8/21
'Y56199 Coupang Corp.
Private Function ExcelSave5() As Boolean
   Dim rsReprot As ADODB.Recordset
   Dim xlsReport As New Excel.Application
   Dim wksReport As New Worksheet
   Dim stFullPath As String
   Dim ii As Integer, dblTFee As Double
   Dim bolInvDate As Boolean, stCol As String, iCols As Integer, jj As Integer 'Added by Morgan 2017/5/10
   
On Error GoTo ErrHnd

   'Modified by Morgan 2025/10/3 改欄位值(IDFNo,CuRef)--Kahn
   'IDF No.: 客戶案件案號,Coupang Reference No. : Client Matter ID
   strExc(0) = "select  pa48||sp29 IDFNo,pa159||sp84 CuRef,a1k13||'-'||a1k14||decode(a1k15||a1k16,'000','','-'||a1k16||'-'||a1k17) OrRef" & _
      ",pa11||sp11 AppNo,pa22 PatNo,a1k01 DNo,to_char(to_date(a1k02+19110000,'yyyymmdd'),'YYYY/MM/DD') DNDate,a1k08 Amt" & _
      ",a1l04,a1l05 from acc1k0,patent,servicepractice,acc1l0 a" & _
      " where nvl(a1k12,0)=0 and a1k25||a1k29 is null and a1k02>=" & Text3 & "01 and a1k02<=" & Text3 & "31 and a1k28='" & m_strNo & "'" & _
      " and pa01(+)=a1k13 and pa02(+)=a1k14 and pa03(+)=a1k15 and pa04(+)=a1k16" & _
      " and sp01(+)=a1k13 and sp02(+)=a1k14 and sp03(+)=a1k15 and sp04(+)=a1k16" & _
      " and a1l01(+)=a1k01 and not exists(select * from acc1l0 b where a1l01=a1k01 and a1l02<a.a1l02)"
   intI = 1
   Set rsReprot = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      MsgBox "請款明細資料讀取失敗!!"
      Exit Function
   End If
   
   xlsReport.Visible = True
   
   xlsReport.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsReport.Workbooks.add
   Set wksReport = xlsReport.Worksheets(1)
   
   wksReport.Cells.NumberFormatLocal = "@"
   wksReport.Cells.Font.Name = "Calibri"
   ii = 2
   wksReport.Range("B" & ii) = "No."
   wksReport.Range("C" & ii) = "IDF No."
   wksReport.Range("D" & ii) = "Coupang Reference No."
   wksReport.Range("E" & ii) = "Your Reference No."
   wksReport.Range("F" & ii) = "Application No."
   wksReport.Range("G" & ii) = "Patent No."
   wksReport.Range("H" & ii) = "Short Description"
   wksReport.Range("I" & ii) = "Invoice No."
   wksReport.Range("J" & ii) = "Invoice Date"
   wksReport.Range("K" & ii) = "Total Amount"
   wksReport.Range("B" & ii, "K" & ii).Font.Bold = True
   
   With rsReprot
   .MoveFirst
   Do While Not .EOF
      ii = ii + 1
      wksReport.Range("B" & ii) = ii - 2
      wksReport.Range("C" & ii) = "" & .Fields("IDFNo")
      wksReport.Range("D" & ii) = "" & .Fields("CuRef")
      wksReport.Range("E" & ii) = "" & .Fields("OrRef")
      wksReport.Range("F" & ii) = .Fields("AppNo")
      wksReport.Range("G" & ii) = "" & .Fields("PatNo")
      wksReport.Range("H" & ii) = GetDNItemDesc(.Fields("DNo"), .Fields("a1l04"), .Fields("a1l05"))
      wksReport.Range("I" & ii) = "" & .Fields("DNo")
      wksReport.Range("J" & ii) = "" & .Fields("DNDate")
      wksReport.Range("K" & ii) = .Fields("Amt")
      wksReport.Range("K" & ii).NumberFormatLocal = "#,##0.00_ "
      wksReport.Range("K" & ii).HorizontalAlignment = xlRight
      .MoveNext
   Loop
   End With
      
   wksReport.Columns("B:K").EntireColumn.AutoFit
   For jj = Asc("B") To Asc("K")
      If wksReport.Range(Chr(jj) & "1").ColumnWidth > 80 Then
         wksReport.Range(Chr(jj) & "1").ColumnWidth = 80
      End If
   Next
   
   xlsReport.Range("A1").Select
   
   stFullPath = m_strSavePath & "\" & m_FileName
   If Dir(stFullPath & ".*") <> "" Then
      Kill stFullPath & ".*"
   End If
   xlsReport.Workbooks(1).SaveAs stFullPath
   xlsReport.Workbooks.Close
   xlsReport.Quit
   
   ExcelSave5 = True
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
ExitPoint:
   Set xlsReport = Nothing
   Set rsReprot = Nothing
End Function

Private Function GetXYrRef(pSys As String, pInvNo As String, ByRef XYrRef As String) As Boolean
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stYrRef As String
   
   If pSys = "FCP" Then
      stSQL = "Select PA106 From Patent, CaseProgress Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP60='" & pInvNo & "' And CP10='605' And CP01='FCP' and pa76 is not null"
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   'Add By Sindy 2016/7/18 +檢查商標延展
   ElseIf InStr("T,FCT,CFT,TF", pSys) > 0 Then
      stSQL = "Select TM65 From Trademark, CaseProgress Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP60='" & pInvNo & "' And CP10='102' And CP01 in('T','FCT','CFT','TF') and TM33 is not null"
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   Else
      intR = 0
   End If
   '2016/7/18 END
   '專利年費/商標延展請款
   If intR = 1 Then
      If PUB_GetFCCaseNo(pInvNo, XYrRef, True) = False Then
         'XYrRef = "" & rsQuery("PA106")
         XYrRef = "" & rsQuery.Fields(0) 'Modify By Sindy 2016/7/18
      End If
      GetXYrRef = True
   Else
      GetXYrRef = PUB_GetFCCaseNo(pInvNo, XYrRef)
   End If
   Set rsQuery = Nothing
End Function

'Added by Morgan 2015/3/2
Private Function PdfSave2(Optional pSaveExcel As Boolean = False) As Boolean
   Const cFontSize = 12
   Dim oTable As Word.Table
   Dim oShape As Word.Shape
   Dim dblAFee As Double, dblOFee As Double, dblDFee As Double, dblTFee As Double, dblTFeeNT As Double
   Dim iRow As Integer, iSNo As Integer
   Dim stFileName As String
   Dim stPdfName As String, stFullPath As String
   Dim rsReprot As ADODB.Recordset
   Dim bolInvDate As Boolean, iCol As Integer, iCols As Integer 'Added by Morgan 2017/5/10
   Dim stAddrNo As String 'Added by Morgan 2021/2/22 列印對象
   Dim oWordAp As Word.Application
   Dim stCon0K0 As String
   Dim stDNCurr As String 'Added by Morgan 2022/7/1
   Dim stTitle As String 'Added by Morgan 2022/11/23
   Dim strInvNo As String 'Added by Morgan 2024/1/24
   
On Error GoTo ErrHnd

   stAddrNo = m_strNo
   
   'Added by Morgan 2017/5/10 Y52418000 OMYA調整格式 --Lina
   If m_strNo = "Y52418000" Then
      bolInvDate = True
      stAddrNo = "Y55553000" 'Added by Morgan 2021/2/22 Y52418000之月帳單列印對象固定顯示Y55553000 --Franny
   End If
   'end 2017/5/10
   
   If m_strNo = "Y55666000" Then bolInvDate = True  'Added by Morgan 2022/2/7
   
   '表頭
   If Left(stAddrNo, 1) = "Y" Then
      strExc(0) = "select fa05,fa63,fa64,fa65,fa32,fa33,fa34,fa35,fa36,fa18,fa19,fa20,fa21,fa22,fa70,fa17,fa23" & _
         " from fagent where fa01='" & Left(stAddrNo, 8) & "' and fa02='" & Mid(stAddrNo, 9) & "'"
   Else
      strExc(0) = "select cu05 as fa05,cu88 as fa63,cu89 as fa64,cu90 as fa65,cu65 as fa32, cu66 as fa33, cu67 as fa34" & _
         ", cu68 as fa35, cu69 as fa36, cu24 as fa18, cu25 as fa19, cu26 as fa20, cu27 as fa21, cu28 as fa22,cu102 fa70" & _
         ", cu23 as fa17, cu29 as fa23 from customer where cu01='" & Left(stAddrNo, 8) & "' and cu02='" & Mid(stAddrNo, 9) & "'"
   End If
   intI = 1
   Set rsReprot = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      MsgBox "表頭資料讀取失敗!!"
      Exit Function
   End If
   
   With rsReprot
   '代理人名稱 strexc(1)
   strExc(1) = "" & .Fields("fa05")
   If Not IsNull(.Fields("fa63")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa63")
   End If
   If Not IsNull(.Fields("fa64")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa64")
   End If
   If Not IsNull(.Fields("fa65")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa65")
   End If
   '代理人POBox/地址
   If Not IsNull(.Fields("fa32")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa32")
      If Not IsNull(.Fields("fa33")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa33")
      End If
      If Not IsNull(.Fields("fa34")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa34")
      End If
      If Not IsNull(.Fields("fa35")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa35")
      End If
      If Not IsNull(.Fields("fa36")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa36")
      End If
   ElseIf Not IsNull(.Fields("fa18")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa18")
      If Not IsNull(.Fields("fa19")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa19")
      End If
      If Not IsNull(.Fields("fa20")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa20")
      End If
      If Not IsNull(.Fields("fa21")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa21")
      End If
      If Not IsNull(.Fields("fa22")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa22")
      End If
      If Not IsNull(.Fields("fa70")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa70")
      End If
   End If
   End With
   
   stCon0K0 = " and a1k02>=" & Text3 & "01 and a1k02<=" & Text4 & "31 and a1k28='" & m_strNo & "'"
   
   'Added by Morgan 2022/6/8
   'Removed by Morgan 2022/8/5 -Y55751 --Franny
   'If m_strNo = "Y55751000" Then
   '   bolInvDate = True
   '   strExc(1) = strExc(1) & vbCrLf & "Attn: Mrs. Daniela Denner, Mr. Marvin Petzold" & vbCrLf
   '   strExc(1) = strExc(1) & vbCrLf & "Cost Center: 35007100"
   '   If Left(Pub_StrUserSt03, 2) = "F1" Then
   '      strExc(1) = strExc(1) & vbCrLf & "Internal Order Number: 61671"
   '      stCon0K0 = stCon0K0 & " and a1k13='FCT'"
   '   Else
   '      strExc(1) = strExc(1) & vbCrLf & "Internal Order Number: 61672"
   '      stCon0K0 = stCon0K0 & " and a1k13='FCP'"
   '   End If
   'End If
   'end 2022/8/5
   'end 2022/6/8
   
   'Modified by Morgan 2016/7/15 請款金額抓 a1k08 比較不會有誤差(非純美金)
   'strExc(0) = "select a1k01,a1k13,to_char(to_date(a1k02+19110000,'yyyymmdd'),'FMMM/DD/yyyy') dt" & _
      ",tm45||pa77||lc23||sp27 YrRef,a1k13||'-'||a1k14||decode(a1k15||a1k16,'000','','-'||a1k16||'-'||a1k17) OrRef,X.*" & _
      ",rtrim(decode(a2607,null,X004,a2607||' '||a2608||' '||a2609)) IDesc" & _
      " from (select a1l01,sum(amt)-sum(decode(substrb(a1l04,-2),'99',amt,0))-sum(decode(a1j03,'雜費',amt,0)) AFee" & _
      ",sum(decode(substrb(a1l04,-2),'99',amt,0)) OFee,sum(decode(a1j03,'雜費',amt,0)) DFee,sum(amt) TFee" & _
      ",min(a1k28) X001,min(a1l03) X002,substr(min(a1l02||a1l04),4) X003,substr(min(a1l02||a1j04),4) X004" & _
      " from (select a.a1l01,a.a1l02,a.a1l03,a.a1l04" & _
      ",decode( nvl(a.a1l17,0), 0, trunc((a.a1l05-nvl(a.a1l07,0)+nvl(b.a1l05,0)-nvl(b.a1l07,0))/a1k10)" & _
      ",trunc((a.a1l17+nvl(b.a1l17,0))* round(1-nvl(a.a1l07,0)/a.a1l05,2))) Amt,a1j03" & _
      ",a1k28,rtrim(a1j04||' '||a1j05||' '||a1j06) a1j04" & _
      " from acc1k0,acc1l0 a,acc1l0 b,acc1j0" & _
      " where nvl(a1k12,0)=0 and a1k25||a1k29 is null and a1k02>=" & Text3 & "01 and a1k02<=" & Text3 & "31 and a1k28='" & m_strNo & "'" & _
      " and a.a1l01(+)=a1k01 and substr(a.a1l04(+),-2)<>'98' and b.a1l01(+)=a.a1l01 and b.a1l03(+)=a.a1l03 and b.a1l04(+)=a.a1l04||'98'" & _
      " and a1j01(+)=a.a1l03 and a1j02(+)=a.a1l04" & _
      ") group by a1l01) X,acc1k0,trademark,patent,lawcase,servicepractice,acc260 where a1k01(+)=a1l01" & _
      " and tm01(+)=a1k13 and tm02(+)=a1k14 and tm03(+)=a1k15 and tm04(+)=a1k16" & _
      " and pa01(+)=a1k13 and pa02(+)=a1k14 and pa03(+)=a1k15 and pa04(+)=a1k16" & _
      " and sp01(+)=a1k13 and sp02(+)=a1k14 and sp03(+)=a1k15 and sp04(+)=a1k16" & _
      " and lc01(+)=a1k13 and lc02(+)=a1k14 and lc03(+)=a1k15 and lc04(+)=a1k16" & _
      " and a2601(+)=substr(X001,1,8) and a2602(+)=X002 and a2603(+)=X003" & _
      " order by 1,2"
   'Modidfied by Morgan 2022/7/1 +a1k18
   strExc(0) = "select a1k01,a1k13,to_char(to_date(a1k02+19110000,'yyyymmdd'),'FMMM/DD/yyyy') dt" & _
      ",tm45||pa77||lc23||sp27 YrRef,a1k13||'-'||a1k14||decode(a1k15||a1k16,'000','','-'||a1k16||'-'||a1k17) OrRef,a1k08 TFee" & _
      ",a1k11,a1k18,X.*,rtrim(decode(a2607,null,X004,a2607||' '||a2608||' '||a2609)) IDesc,tm35||pa48||lc17||sp29 CuRef,to_char(to_date(a1k02+19110000,'yyyymmdd'),'YYYY/MM/DD') dt2" & _
      " from (select a1l01,min(a1k28) X001,min(a1l03) X002,substr(min(a1l02||a1l04),4) X003,substr(min(a1l02||a1j04),4) X004" & _
      " from (select a.a1l01,a.a1l02,a.a1l03,a.a1l04,a1j03" & _
      ",a1k28,rtrim(a1j04||' '||a1j05||' '||a1j06) a1j04" & _
      " from acc1k0,acc1l0 a,acc1l0 b,acc1j0" & _
      " where nvl(a1k12,0)=0 and a1k25||a1k29 is null " & stCon0K0 & _
      " and a.a1l01(+)=a1k01 and substr(a.a1l04(+),-2)<>'98' and b.a1l01(+)=a.a1l01 and b.a1l03(+)=a.a1l03 and b.a1l04(+)=a.a1l04||'98'" & _
      " and a1j01(+)=a.a1l03 and a1j02(+)=a.a1l04" & _
      ") group by a1l01) X,acc1k0,trademark,patent,lawcase,servicepractice,acc260 where a1k01(+)=a1l01" & _
      " and tm01(+)=a1k13 and tm02(+)=a1k14 and tm03(+)=a1k15 and tm04(+)=a1k16" & _
      " and pa01(+)=a1k13 and pa02(+)=a1k14 and pa03(+)=a1k15 and pa04(+)=a1k16" & _
      " and sp01(+)=a1k13 and sp02(+)=a1k14 and sp03(+)=a1k15 and sp04(+)=a1k16" & _
      " and lc01(+)=a1k13 and lc02(+)=a1k14 and lc03(+)=a1k15 and lc04(+)=a1k16" & _
      " and a2601(+)=substr(X001,1,8) and a2602(+)=X002 and a2603(+)=X003" & _
      " order by 1,2"
   intI = 1
   Set rsReprot = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      MsgBox "請款明細資料讀取失敗!!"
      Exit Function
   End If
   
   'Modified by Morgan 2022/5/18
   'If NewWordDoc = False Then Exit Function
   'With g_WordAp.Application
   Set oWordAp = New Word.Application
   oWordAp.Visible = True
   oWordAp.Documents.add
   With oWordAp
   'end 2022/5/18
   
      .Selection.Font.Name = "Times New Roman"
      .Selection.Font.Size = cFontSize
      
      '版面設定
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.5)
      'Modified by Morgan 2023/12/20 改和請款單一致
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(4.3)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4)
      'end 2023/12/20
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
      .Selection.PageSetup.FooterDistance = .CentimetersToPoints(3)
      .Selection.PageSetup.CharsLine = 40
      .Selection.PageSetup.LinesPage = 38
      .Selection.Orientation = wdTextOrientationHorizontal
      
      '信頭尾
      If PUB_ReadDB2File(stFileName, iPicNo) = True Then
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
         Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
         oShape.ZOrder 4
         oShape.LockAnchor = True
         oShape.LockAspectRatio = -1
         oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
         oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
         oShape.Left = 0
         oShape.Top = 0
         oShape.Width = .CentimetersToPoints(21)
         oShape.WrapFormat.Type = wdWrapNone
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         
         If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
            .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
            oShape.ZOrder 4
            oShape.LockAnchor = True
            oShape.LockAspectRatio = -1
            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
            oShape.Left = 0
            oShape.Top = .CentimetersToPoints(27)
            oShape.Width = .CentimetersToPoints(21)
            oShape.WrapFormat.Type = wdWrapNone
            .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         End If
         
         .Selection.HomeKey Unit:=wdStory
      End If
      
      
      .Selection.TypeParagraph
      '行距
      With .Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .DisableLineHeightGrid = True
      End With
      
      '新增表格(1*2)
      Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumRows:=1, NumColumns:=2)
      With oTable
         '無邊框
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
      End With
            
      oTable.Select
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop '靠上對齊
      .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
      .Selection.InsertRows 8

      '代理人名稱,POBox/地址
      oTable.Cell(1, 1).Merge oTable.Cell(4, 1)
      oTable.Cell(1, 1).Select
      .Selection.Text = strExc(1)
      
      '月份
      strExc(0) = Format(ChangeTStringToWDateString(Text3 & "01"), "mmmm, yyyy")
      'Added by Morgan 2021/2/22
      If Text4 <> Text3 Then
         strExc(0) = strExc(0) & " - " & Format(ChangeTStringToWDateString(Text4 & "01"), "mmmm, yyyy")
      End If
      'end 2021/2/22
      
      'Added by Morgan 2022/11/23
      stTitle = "Monthly Invoice"
      If m_strNo = "Y55666000" Then
         'Modified by Morgan 2024/1/24
         'stTitle = stTitle & " (for " & strExc(0) & ")"
         strInvNo = rsReprot("a1k01") & "/" & Text3
         stTitle = stTitle & " No. " & strInvNo & " (for " & strExc(0) & ")"
         'end 2024/1/24
      Else
      'end 2022/11/23
         oTable.Cell(1, 2).Select
         .Selection.Text = "Month: " & strExc(0)
      End If
      
      'Added by Morgan 2017/5/10
      If bolInvDate Then
         strExc(0) = "Date: " & Format(ChangeTStringToWDateString(strSrvDate(2)), "mmmm dd, yyyy")
         oTable.Cell(2, 2).Select
         .Selection.Text = strExc(0)
      
      'Added by Morgan 2024/6/24
      ElseIf m_strNo = "Y55973000" Then
         strExc(0) = "Date of the invoice: " & Format(ChangeTStringToWDateString(strSrvDate(2)), "dd.mm.yyyy")
         oTable.Cell(2, 2).Select
         .Selection.Text = strExc(0)
      End If
      'end 2017/5/10
      
      'Added by Morgan 2018/11/22 --Lina
      If m_strNo = "Y34412010" Then
         strExc(0) = "Contract No.: CM0051A22420350579"
         oTable.Cell(4, 2).Select
         .Selection.Text = strExc(0)
      End If
      'end 2018/11/22
            
      iRow = 5
      
      'Added by Morgan 2024/6/24
      If m_strNo = "Y55973000" Then
         oTable.Cell(iRow, 1).Select
         .Selection.InsertRows 2
         iRow = iRow + 1
         
         strExc(0) = Format(ChangeTStringToWDateString(Text3 & "01"), "mmmm, yyyy")
         oTable.Cell(iRow, 1).Select
         .Selection.Text = "Service Period: " & strExc(0)
         
         iRow = iRow + 1
         stTitle = stTitle & " No. " & rsReprot("a1k01") & "/" & Text3
      End If
      'end 2024/6/24
         
      'Added by Morgan 2022/3/17 Y55666000 若有財務編號也要印 --Ryan
      If m_strNo = "Y55666000" Then
         strExc(0) = PUB_GetACCNO(m_strNo)
         oTable.Cell(iRow, 1).Select
         .Selection.Text = strExc(0)
         iRow = iRow + 1
         oTable.Cell(iRow, 1).Select
         .Selection.InsertRows 1
      End If
      'end 2022/3/17
      
      oTable.Cell(iRow, 1).Merge oTable.Cell(iRow, 2)
      oTable.Cell(iRow, 1).Select
      .Selection.Cells(1).SetHeight RowHeight:=30, HeightRule:=wdRowHeightAtLeast
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      'Modified by Morgan 2022/11/23
      '.Selection.Text = "Monthly Invoice"
      .Selection.Text = stTitle
      'end 2022/11/23
      iRow = iRow + 1
      
      oTable.Cell(iRow, 1).Select
      .Selection.SelectRow

      With .Selection.Cells
        '有邊框
        .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
        .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
      End With
      
      'Added by Morgan 2024/1/24
      'Modified by Morgan 2024/6/24 +Y55973000
      If m_strNo = "Y55666000" Or m_strNo = "Y55973000" Then
         iCols = 6
      'end 2024/1/24
      'Added by Morgan 2017/5/10
      ElseIf bolInvDate Then
         iCols = 7
      Else
         'Modified by Morgan 2019/10/17 +Case No.
         'iCols = 6
         iCols = 7
      End If
      'end 2017/5/10
      
      .Selection.Cells.Split NumRows:=1, NumColumns:=iCols, MergeBeforeSplit:=True
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      .Selection.Font.Size = 10
      
      'Added by Morgan 2024/1/24
      'Modified by Morgan 2024/6/24 +Y55973000
      If m_strNo = "Y55666000" Or m_strNo = "Y55973000" Then
         '設定表格高度欄寬
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.7), RulerStyle:=wdAdjustProportional
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
         .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.7), RulerStyle:=wdAdjustProportional
         .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
         .Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(8), RulerStyle:=wdAdjustProportional
      'end 2024/1/14
      
      'Added by Morgan 2017/5/10
      ElseIf bolInvDate Then
         '設定表格高度欄寬
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.7), RulerStyle:=wdAdjustProportional
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
         .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
         .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.7), RulerStyle:=wdAdjustProportional
         .Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
         'Added by Morgan 2022/7/1
         'Removed by Morgan 2022/8/5 -Y55751 --Franny
         'If m_strNo = "Y55751000" Then
         '   .Selection.Cells(6).SetWidth ColumnWidth:=.CentimetersToPoints(5.3), RulerStyle:=wdAdjustProportional
         'Else
         'end 2022/8/5
         'end 2022/7/1
         
            .Selection.Cells(6).SetWidth ColumnWidth:=.CentimetersToPoints(5.8), RulerStyle:=wdAdjustProportional
         'End If 'Removed by Morgan 2022/8/5
      Else
      'end 2017/5/10
      
         '設定表格高度欄寬
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.7), RulerStyle:=wdAdjustProportional
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
         'Modified by Morgan 2019/10/17 +Case No.
         '.Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.6), RulerStyle:=wdAdjustProportional
         '.Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
         '.Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(8#), RulerStyle:=wdAdjustProportional
         .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
         .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.7), RulerStyle:=wdAdjustProportional
         .Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
         .Selection.Cells(6).SetWidth ColumnWidth:=.CentimetersToPoints(5.8), RulerStyle:=wdAdjustProportional
         'end 2019/10/17
      End If 'Added by Morgan 2017/5/10
      
      .Selection.Cells(1).SetHeight RowHeight:=36, HeightRule:=wdRowHeightAtLeast
      .Selection.InsertRows rsReprot.RecordCount + 1
      
      oTable.Cell(iRow, 1).Select
      .Selection.SelectRow
      .Selection.Font.Bold = True
      oTable.Cell(iRow, 1).Select
      .Selection.Text = "No."
      
      
      'Added by Morgan 2024/1/24
      'Modified by Morgan 2024/6/24 +Y55973000
      If m_strNo = "Y55666000" Or m_strNo = "Y55973000" Then
         oTable.Cell(iRow, 2).Select
         .Selection.Text = "Invoice Date"
         oTable.Cell(iRow, 3).Select
         .Selection.Text = "Your Ref"
         oTable.Cell(iRow, 4).Select
         .Selection.Text = "Our Ref"
         oTable.Cell(iRow, 5).Select
         .Selection.Text = "Description"
         oTable.Cell(iRow, 6).Select
         .Selection.Text = "Total Fee" & vbCrLf & "(USD)"
      'end 2024/1/24
      
      'Added by Morgan 2017/5/10
      ElseIf bolInvDate Then
         oTable.Cell(iRow, 2).Select
         .Selection.Text = "Invoice No."
         oTable.Cell(iRow, 3).Select
         .Selection.Text = "Invoice Date"
         oTable.Cell(iRow, 4).Select
         .Selection.Text = "Your Ref"
         oTable.Cell(iRow, 5).Select
         .Selection.Text = "Our Ref"
         oTable.Cell(iRow, 6).Select
         .Selection.Text = "Description"
         oTable.Cell(iRow, 7).Select
         'Added by Morgan 2022/7/1
         'Removed by Morgan 2022/8/5 -Y55751 --Franny
         'If m_strNo = "Y55751000" Then
         '   stDNCurr = "" & rsReprot("a1k18")
         '   .Selection.Text = "Total Fee" & vbCrLf & "(NTD/" & stDNCurr & ")"
         'Else
         'end 2022/8/5
         'end 2022/7/1
         
            .Selection.Text = "Total Fee" & vbCrLf & "(USD)"
         'End If 'Removed by Morgan 2022/8/5
      Else
      'end 2017/5/10
         oTable.Cell(iRow, 2).Select
         .Selection.Text = "Invoice No."
         oTable.Cell(iRow, 3).Select
         'Added by Morgan 2023/12/20
         If m_strNo = "X82995000" Then
            .Selection.Text = "US firm's " & vbCrLf & "Ref"
         Else
            .Selection.Text = "Your Ref"
         End If
         oTable.Cell(iRow, 4).Select
         'Modified by Morgan 2019/10/17 +Case No.
         .Selection.Text = "Case No."
         oTable.Cell(iRow, 5).Select
         .Selection.Text = "Our Ref"
         oTable.Cell(iRow, 6).Select
         .Selection.Text = "Description"
         oTable.Cell(iRow, 7).Select
         .Selection.Text = "Total Fee" & vbCrLf & "(USD)"
         
      End If 'Added by Morgan 2017/5/10
      
      .Selection.SelectRow
      '.Selection.Cells.Shading.BackgroundPatternColorIndex = wdTurquoise
      .Selection.Cells.Shading.Texture = wdTexture15Percent
      .Selection.Cells(1).SetHeight RowHeight:=36, HeightRule:=wdRowHeightAtLeast
      iSNo = 0
      Do While Not rsReprot.EOF
         iRow = iRow + 1
         iSNo = iSNo + 1
         oTable.Cell(iRow, 1).Select
         .Selection.Text = iSNo
         'Added by Morgan 2024/1/24
         If m_strNo = "Y55666000" Then
            iCol = 2
            oTable.Cell(iRow, iCol).Select
            .Selection.Text = "" & rsReprot("dt")
         
         'Added by Morgan 2024/6/24
         ElseIf m_strNo = "Y55973000" Then
            iCol = 2
            oTable.Cell(iRow, iCol).Select
            .Selection.Text = "" & rsReprot("dt2")
         'end 2024/6/24
         Else
         'end 2024/1/24
            oTable.Cell(iRow, 2).Select
            .Selection.Text = "" & rsReprot("a1k01")
            iCol = 2
            'Added by Morgan 2017/5/10
            If bolInvDate Then
               iCol = iCol + 1
               oTable.Cell(iRow, iCol).Select
               .Selection.Text = "" & rsReprot("dt")
            End If
            'end 2017/5/10
         End If 'Added by Morgan 2024/1/24
      
         '彼所案號
         'Added by Morgan 2025/6/13
         '請款對象Y55666000 NOVOCURE GMBH 彼號欄位 (Your ref:) 優先抓客戶案件案號--Franny
         'Modified by Morgan 2025/6/25 更代後客戶案號會改放到彼號 Ex:X11408163--Franny
         If m_strNo = "Y55666000" Then
            If "" & rsReprot("CuRef") <> "" Then
               strExc(1) = "" & rsReprot("CuRef")
            Else
               strExc(1) = "" & rsReprot("YrRef")
            End If
         Else
         'end 2025/6/13
            strExc(1) = "" & rsReprot("YrRef")
            If GetXYrRef(rsReprot("a1k13"), rsReprot("A1K01"), strExc(2)) = True Then
               strExc(1) = strExc(2)
            End If
         End If
         
         iCol = iCol + 1
         oTable.Cell(iRow, iCol).Select
         .Selection.Text = strExc(1)
         'Added by Morgan 2019/10/17 +Case No.
         'Modified by Morgan 2024/6/24 Y55973000除外
         If bolInvDate = False And m_strNo <> "Y55973000" Then
            iCol = iCol + 1
            oTable.Cell(iRow, iCol).Select
            .Selection.Text = "" & rsReprot("CuRef")
         End If
         'end 2019/10/17
         iCol = iCol + 1
         oTable.Cell(iRow, iCol).Select
         .Selection.Text = "" & rsReprot("OrRef")
         iCol = iCol + 1
         oTable.Cell(iRow, iCol).Select
         'Added by Morgan 2022/5/13
         If m_strNo = "Y55666000" Then
            .Selection.Text = "" & GetNOVODesc("" & rsReprot("a1k01"))
         Else
         'end 2022/5/13
            'Added by Morgan 2024/11/4
            If (rsReprot("a1k13") = "FCP" Or rsReprot("a1k13") = "P" Or rsReprot("a1k13") = "CFP") And (rsReprot("X003") = "601" Or rsReprot("X003") = "605") Then
               .Selection.Text = PUB_GetAnnuityDesc(rsReprot("a1k01"), rsReprot("X003"), "" & rsReprot("IDesc"))
            Else
            'end 2024/11/4
               .Selection.Text = "" & rsReprot("IDesc")
            End If
         End If 'Added by Morgan 2022/5/13
         
         iCol = iCol + 1
         oTable.Cell(iRow, iCol).Select
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         'Added by Morgan 2022/7/1
         'Removed by Morgan 2022/8/5 -Y55751 --Franny
         'If m_strNo = "Y55751000" Then
         '   .Selection.TypeText "NTD" & Format(Val("" & rsReprot("a1k11")), cfmtDollar)
         '   .Selection.TypeParagraph
         '   .Selection.TypeText stDNCurr & Format(Val("" & rsReprot("TFee")), cfmtDollar)
         'Else
         'end 2022/8/5
         'end 2022/7/1
         
            .Selection.Text = Format(Val("" & rsReprot("TFee")), cfmtDollar)
         'End If 'Removed by Morgan 2022/8/5 --Franny
         dblTFee = dblTFee + Val("" & rsReprot("TFee"))
         dblTFeeNT = dblTFeeNT + Val("" & rsReprot("a1k11")) 'Added by Morgan 2022/7/1
         rsReprot.MoveNext
      Loop
      
      iRow = iRow + 1
      oTable.Cell(iRow, 1).Merge oTable.Cell(iRow, iCols - 1)
      oTable.Cell(iRow, 1).Select
      .Selection.SelectRow
      .Selection.Font.Bold = True
      
      oTable.Cell(iRow, 1).Select
      .Selection.Text = "Total"
      oTable.Cell(iRow, 2).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      'Added by Morgan 2022/7/1
      'Removed by Morgan 2022/8/5 -Y55751 --Franny
      'If m_strNo = "Y55751000" Then
      '   .Selection.TypeText "NTD" & Format(dblTFeeNT, cfmtDollar)
      '   .Selection.TypeParagraph
      '   .Selection.TypeText stDNCurr & Format(dblTFee, cfmtDollar)
      'Else
      'end 2022/8/5
      'end 2022/7/1
      
         .Selection.Text = Format(dblTFee, cfmtDollar)
      'End If 'Removed by Morgan 2022/8/5 --Franny
      
      '帳號
      iRow = iRow + 2
      oTable.Cell(iRow, 1).Select
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
      .Selection.Text = vbCrLf & ReportSum(71001) & vbCrLf & ReportSum(72) & vbCrLf & ReportSum(73001) & vbCrLf & ReportSum(85) & vbCrLf & ReportSum(74) & vbCrLf & ReportSum(121) & vbCrLf
      
      '建議電匯提醒
      oTable.Cell(iRow, 2).Select
      .Selection.Cells.Split NumRows:=3, NumColumns:=1, MergeBeforeSplit:=False
      .Selection.Cells(1).SetHeight RowHeight:=28, HeightRule:=wdRowHeightAtLeast
      iRow = iRow + 1
      oTable.Cell(iRow, 2).Select
      .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
      With .Selection.Cells(1)
          With .Borders(wdBorderLeft)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderRight)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderTop)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderBottom)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
      End With
      .Selection.ParagraphFormat.LeftIndent = .CentimetersToPoints(0.2)
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
      .Selection.Cells(1).SetHeight RowHeight:=52, HeightRule:=wdRowHeightAtLeast
      .Selection.Text = "Wire" & vbCrLf & "Transfer" & vbCrLf & "Preferred"
      iRow = iRow + 1
      oTable.Cell(iRow, 2).Select
      .Selection.Cells(1).SetHeight RowHeight:=0, HeightRule:=wdRowHeightAtLeast
      
      '備註
      iRow = iRow + 1
      oTable.Cell(iRow, 1).Select
      .Selection.SelectRow
      .Selection.Font.Bold = True
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
      oTable.Cell(iRow, 1).Select
      .Selection.Text = "PS:"
      oTable.Cell(iRow, 2).Select
      .Selection.Text = "Please return a copy of the invoice(s) or indicate the invoice number(s) paid with remittance"
      .Selection.EndKey
      
      'Added by Morgan 2022/5/18
      'Modified by Morgan 2024/6/24 +Y55973000
      If m_RptNo = "9" Or m_strNo = "Y55973000" Then
         rsReprot.MoveFirst
         Do While Not rsReprot.EOF
            '請款單電子檔
            Load Frmacc2480
            With Frmacc2480
               .Text1.Text = rsReprot("a1k01")
               .Text2.Text = .Text1.Text
               .m_bEditDoc = True
               .m_bBeCalled = True
               .m_CallPrevForm = Me.Name
               .m_bolNoPic = True 'Added by Morgan 2023/12/20
               'Added by Morgan 2024/6/24 明細不印請款單號
               If m_strNo = "Y55973000" Then
                  .m_strInvoiceNo = " "
               Else
               'end 2024/6/24
                  .m_strInvoiceNo = strInvNo 'Added by Morgan 2024/1/24
               End If
               .Command2_Click
            End With
            Unload Frmacc2480
            strFormName = Me.Name
            tool3_enabled
            
            oWordAp.Selection.EndKey Unit:=wdStory
            oWordAp.Selection.TypeText Chr(12)
            g_WordAp.Selection.WholeStory
            g_WordAp.Selection.Copy
            oWordAp.Selection.Paste
            g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
            g_WordAp.Quit wdDoNotSaveChanges
            
            rsReprot.MoveNext
         Loop
      End If
      
      If Text2 = "Y" Then
         .Activate
      Else
         'Added by Morgan 2023/12/20
         If m_strNo = "X82995000" Then
            stPdfName = "Monthly Invoice (" & Format(ChangeTStringToWDateString(Text3 & "01"), "mmmm, yyyy") & ").pdf"
         Else
         'end 2023/12/20
            stPdfName = m_strNo & Text3 & ".pdf"
         End If
         
'Modified by Morgan 2022/6/8
'         If Pub_GetPrinterIndex("PDFCreator") < 0 Then
'            MsgBox "請先安裝 PDFCreator 印表機，pdf 轉檔失敗！", vbExclamation
'            Exit Function
'         Else

'            pub_OsPrinter = PUB_GetOsDefaultPrinter
'            frmPDF.Show
'            frmPDF.StartProcess m_strSavePath, stPdfName
'            PUB_SetOsDefaultPrinter Printer.DeviceName
'            PUB_SetWordActivePrinter
'            .ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
'            .ActiveDocument.Close wdDoNotSaveChanges
'            .Quit wdDoNotSaveChanges
'            frmPDF.EndtProcess
'            Unload frmPDF
'            PUB_SetOsDefaultPrinter pub_OsPrinter
'         End If
         .ActiveDocument.ExportAsFixedFormat OutputFileName:=m_strSavePath & "\" & stPdfName, ExportFormat:=17, OpenAfterExport:=False
         .ActiveDocument.Close wdDoNotSaveChanges
         .Quit wdDoNotSaveChanges
'end 2022/6/8

      End If
   End With
   PdfSave2 = True
   
   'Added by Morgan 2016/9/30
   If pSaveExcel Then
      ExcelSave2 rsReprot
   End If
   'end 2016/9/30
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   Set rsReprot = Nothing
   Set oWordAp = Nothing 'Added by Morgan 2022/5/18
End Function

'Added by Morgan 2015/3/5
'Modified by Morgan 2024/9/3 +pAllDN
Private Function PdfSave3(Optional pAllDN As Boolean = False) As Boolean
   Const cFontSize = 12
   Const cfmtDollar = "#,##0.00"
   Dim oTable As Word.Table
   Dim oShape As Word.Shape
   Dim dblAFee As Double, dblOFee As Double, dblDFee As Double, dblTFee As Double
   Dim iRow As Integer
   Dim stFileName As String
   Dim stPdfName As String, stFullPath As String
   Dim rsReprot As ADODB.Recordset
   
   
On Error GoTo ErrHnd
   
   '表頭
   If Left(m_strNo, 1) = "Y" Then
      strExc(0) = "select fa05,fa63,fa64,fa65,fa32,fa33,fa34,fa35,fa36,fa18,fa19,fa20,fa21,fa22,fa70,fa17,fa23" & _
         " from fagent where fa01='" & Left(m_strNo, 8) & "' and fa02='" & Mid(m_strNo, 9) & "'"
   Else
      strExc(0) = "select cu05 as fa05,cu88 as fa63,cu89 as fa64,cu90 as fa65,cu65 as fa32, cu66 as fa33, cu67 as fa34" & _
         ", cu68 as fa35, cu69 as fa36, cu24 as fa18, cu25 as fa19, cu26 as fa20, cu27 as fa21, cu28 as fa22,cu102 fa70" & _
         ", cu23 as fa17, cu29 as fa23 from customer where cu01='" & Left(m_strNo, 8) & "' and cu02='" & Mid(m_strNo, 9) & "'"
   End If
   intI = 1
   Set rsReprot = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      MsgBox "表頭資料讀取失敗!!"
      Exit Function
   End If
   
   With rsReprot
   '代理人名稱 strexc(1)
   strExc(1) = "" & .Fields("fa05")
   If Not IsNull(.Fields("fa63")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa63")
   End If
   If Not IsNull(.Fields("fa64")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa64")
   End If
   If Not IsNull(.Fields("fa65")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa65")
   End If
   '代理人POBox/地址
   If Not IsNull(.Fields("fa32")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa32")
      If Not IsNull(.Fields("fa33")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa33")
      End If
      If Not IsNull(.Fields("fa34")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa34")
      End If
      If Not IsNull(.Fields("fa35")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa35")
      End If
      If Not IsNull(.Fields("fa36")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa36")
      End If
   ElseIf Not IsNull(.Fields("fa18")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa18")
      If Not IsNull(.Fields("fa19")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa19")
      End If
      If Not IsNull(.Fields("fa20")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa20")
      End If
      If Not IsNull(.Fields("fa21")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa21")
      End If
      If Not IsNull(.Fields("fa22")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa22")
      End If
      If Not IsNull(.Fields("fa70")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa70")
      End If
   End If
   End With
   
   '項目中文雜費者加總
   strExc(0) = "select a1k01,a1k13,to_char(to_date(a1k02+19110000,'yyyymmdd'),'FMMM/DD/yyyy') dt" & _
      ",tm45||pa77||lc23||sp27 YrRef,a1k13||'-'||a1k14||decode(a1k15||a1k16,'000','','-'||a1k16||'-'||a1k17) OrRef,X.*" & _
      ",rtrim(decode(a2607,null,X004,a2607||' '||a2608||' '||a2609)) IDesc" & _
      " from (select a1l01,sum(amt)-sum(decode(substrb(a1l04,-2),'99',amt,0))-sum(decode(a1j03,'雜費',amt,0)) AFee" & _
      ",sum(decode(substrb(a1l04,-2),'99',amt,0)) OFee,sum(decode(a1j03,'雜費',amt,0)) DFee,sum(amt) TFee" & _
      ",min(a1k28) X001,min(a1l03) X002,substr(min(a1l02||a1l04),4) X003,substr(min(a1l02||a1j04),4) X004" & _
      " from (select a.a1l01,a.a1l02,a.a1l03,a.a1l04" & _
      ",decode( nvl(a.a1l17,0), 0, trunc((a.a1l05-nvl(a.a1l07,0)+nvl(b.a1l05,0)-nvl(b.a1l07,0))/a1k10)" & _
      ",trunc((a.a1l17+nvl(b.a1l17,0))* round(1-nvl(a.a1l07,0)/a.a1l05,2))) Amt,a1j03" & _
      ",a1k28,rtrim(a1j04||' '||a1j05||' '||a1j06) a1j04" & _
      " from acc1k0,acc1l0 a,acc1l0 b,acc1j0" & _
      " where nvl(a1k12,0)=0 and a1k25" & IIf(pAllDN, "", "||a1k29") & " is null and a1k02>=" & Text3 & "01 and a1k02<=" & Text3 & "31 and a1k28='" & m_strNo & "'" & _
      " and a.a1l01(+)=a1k01 and substr(a.a1l04(+),-2)<>'98' and b.a1l01(+)=a.a1l01 and b.a1l03(+)=a.a1l03 and b.a1l04(+)=a.a1l04||'98'" & _
      " and a1j01(+)=a.a1l03 and a1j02(+)=a.a1l04" & _
      ") group by a1l01) X,acc1k0,trademark,patent,lawcase,servicepractice,acc260 where a1k01(+)=a1l01" & _
      " and tm01(+)=a1k13 and tm02(+)=a1k14 and tm03(+)=a1k15 and tm04(+)=a1k16" & _
      " and pa01(+)=a1k13 and pa02(+)=a1k14 and pa03(+)=a1k15 and pa04(+)=a1k16" & _
      " and sp01(+)=a1k13 and sp02(+)=a1k14 and sp03(+)=a1k15 and sp04(+)=a1k16" & _
      " and lc01(+)=a1k13 and lc02(+)=a1k14 and lc03(+)=a1k15 and lc04(+)=a1k16" & _
      " and a2601(+)=substr(X001,1,8) and a2602(+)=X002 and a2603(+)=X003" & _
      " order by 1,2"
   intI = 1
   Set rsReprot = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      MsgBox "請款明細資料讀取失敗!!"
      Exit Function
   End If
      
   '請款單號 strexc(2)
   strExc(2) = rsReprot("a1k01") & "/" & Text3
   
   If NewWordDoc = False Then Exit Function
   
   With g_WordAp.Application
      
      .Selection.Font.Name = "Times New Roman"
      .Selection.Font.Size = cFontSize
      
      '版面設定
      .Selection.PageSetup.PaperSize = wdPaperA4
      .Selection.PageSetup.Orientation = wdOrientLandscape
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.5)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.3)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3.2)
      .Selection.PageSetup.FooterDistance = .CentimetersToPoints(3)
      '.Selection.PageSetup.CharsLine = 40
      '.Selection.PageSetup.LinesPage = 38
      '.Selection.Orientation = wdTextOrientationHorizontal
      
      '信頭尾
      If PUB_ReadDB2File(stFileName, iPicNo) = True Then
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
         Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
         oShape.ZOrder 4
         oShape.LockAnchor = True
         oShape.LockAspectRatio = -1
         oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
         oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
         oShape.Left = 0
         oShape.Top = 0
         oShape.Width = .CentimetersToPoints(21)
         oShape.WrapFormat.Type = wdWrapNone
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         
         If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
            .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
            oShape.ZOrder 4
            oShape.LockAnchor = True
            oShape.LockAspectRatio = -1
            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
            oShape.Left = 0
            'oShape.Top = .CentimetersToPoints(27)
            oShape.Top = .CentimetersToPoints(18)
            oShape.Width = .CentimetersToPoints(21)
            oShape.WrapFormat.Type = wdWrapNone
            .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         End If
         
         .Selection.HomeKey Unit:=wdStory
      End If
      
      
      '.Selection.TypeParagraph
      '行距
      With .Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .DisableLineHeightGrid = True
      End With
      
      '新增表格(1*2)
      Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumRows:=1, NumColumns:=2)
      With oTable
         '無邊框
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
      End With
            
      oTable.Select
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop '靠上對齊
      .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
      .Selection.InsertRows 9

      '代理人名稱,POBox/地址
      oTable.Cell(1, 1).Merge oTable.Cell(5, 1)
      oTable.Cell(1, 1).Select
      .Selection.Text = strExc(1)
      
      'Attention
      strExc(0) = "Attention: Patent Invoices"
      oTable.Cell(1, 2).Select
      .Selection.Text = strExc(0)
      
      'Time period covered
      strExc(0) = "Time period covered: " & ChangeTStringToWDateString(Text3 & "01") & "~" & ChangeWStringToWDateString(GetLastDay(Text3 & "01"))
      oTable.Cell(2, 2).Select
      .Selection.Text = strExc(0)
      
      If m_strNo = "Y45493000" Then 'Added by Morgan 2024/9/3 因為要共用,增加判斷Lundbeck才印
         'Lundbeck cost center
         strExc(0) = "Lundbeck cost center: 10008133"
         oTable.Cell(3, 2).Select
         .Selection.Text = strExc(0)
      End If
      
      'Invoice No
      strExc(0) = "Invoice No: " & strExc(2)
      oTable.Cell(4, 2).Select
      .Selection.Text = strExc(0)
      
      oTable.Cell(6, 1).Merge oTable.Cell(6, 2)
      oTable.Cell(6, 1).Select
      .Selection.Cells(1).SetHeight RowHeight:=30, HeightRule:=wdRowHeightAtLeast
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      .Selection.Text = "Monthly Invoice"
      
      oTable.Cell(7, 1).Select
      .Selection.SelectRow

      With .Selection.Cells
        '有邊框
        .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
        .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
      End With
      
      .Selection.Cells.Split NumRows:=1, NumColumns:=9, MergeBeforeSplit:=True
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      .Selection.Font.Size = 10
      '設定表格高度欄寬
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.7), RulerStyle:=wdAdjustProportional
      .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.6), RulerStyle:=wdAdjustProportional
      .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.6), RulerStyle:=wdAdjustProportional
      .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
      .Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(9.6), RulerStyle:=wdAdjustProportional
      .Selection.Cells(6).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
      .Selection.Cells(7).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
      .Selection.Cells(8).SetWidth ColumnWidth:=.CentimetersToPoints(2.4), RulerStyle:=wdAdjustProportional
      
      .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
      .Selection.InsertRows rsReprot.RecordCount + 1
      
      oTable.Cell(7, 1).Select
      .Selection.SelectRow
      .Selection.Font.Bold = True
      
      oTable.Cell(7, 1).Select
      .Selection.Text = "No."
      oTable.Cell(7, 2).Select
      .Selection.Text = "Invoice Date" & vbCrLf & "<mm/dd/yyyy>"
      oTable.Cell(7, 3).Select
      .Selection.Text = "Your Ref"
      oTable.Cell(7, 4).Select
      .Selection.Text = "Our Ref"
      oTable.Cell(7, 5).Select
      .Selection.Text = "Description"
      oTable.Cell(7, 6).Select
      .Selection.Text = "Attorney" & vbCrLf & "Fee" & vbCrLf & "(USD)"
      oTable.Cell(7, 7).Select
      .Selection.Text = "Official Fee" & vbCrLf & "(USD)"
      oTable.Cell(7, 8).Select
      .Selection.Text = "Disbursement" & vbCrLf & "Fee" & vbCrLf & "(USD)"
      oTable.Cell(7, 9).Select
      .Selection.Text = "Total Fee" & vbCrLf & "(USD)"
      
      .Selection.SelectRow
      '.Selection.Cells.Shading.BackgroundPatternColorIndex = wdTurquoise
      .Selection.Cells.Shading.Texture = wdTexture15Percent
      .Selection.Cells(1).SetHeight RowHeight:=36, HeightRule:=wdRowHeightAtLeast
      
      iRow = 7
      Do While Not rsReprot.EOF
         iRow = iRow + 1
         'No.
         oTable.Cell(iRow, 1).Select
         .Selection.Text = iRow - 7
         'Invoice Date
         oTable.Cell(iRow, 2).Select
         .Selection.Text = "" & rsReprot("dt")
         'Your Ref
         strExc(1) = "" & rsReprot("YrRef")
         If GetXYrRef(rsReprot("a1k13"), rsReprot("A1K01"), strExc(2)) = True Then
            strExc(1) = strExc(2)
         End If
         oTable.Cell(iRow, 3).Select
         .Selection.Text = strExc(1)
         'Our Ref
         oTable.Cell(iRow, 4).Select
         .Selection.Text = "" & rsReprot("OrRef")
         'Description
         oTable.Cell(iRow, 5).Select
         .Selection.Text = "" & rsReprot("IDesc")
         'Attorney Fee
         oTable.Cell(iRow, 6).Select
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Text = Format(Val("" & rsReprot("AFee")), cfmtDollar)
         'Official Fee
         oTable.Cell(iRow, 7).Select
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Text = Format(Val("" & rsReprot("OFee")), cfmtDollar)
         'Disbursement
         oTable.Cell(iRow, 8).Select
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Text = Format(Val("" & rsReprot("DFee")), cfmtDollar)
         'Total Fee
         oTable.Cell(iRow, 9).Select
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Text = Format(Val("" & rsReprot("TFee")), cfmtDollar)
         
         
         dblAFee = dblAFee + Val("" & rsReprot("AFee"))
         dblOFee = dblOFee + Val("" & rsReprot("OFee"))
         dblDFee = dblDFee + Val("" & rsReprot("DFee"))
         dblTFee = dblTFee + Val("" & rsReprot("TFee"))
         rsReprot.MoveNext
      Loop
      
      iRow = iRow + 1
      oTable.Cell(iRow, 1).Merge oTable.Cell(iRow, 5)
      oTable.Cell(iRow, 1).Select
      .Selection.SelectRow
      .Selection.Font.Bold = True
      
      oTable.Cell(iRow, 1).Select
      .Selection.Text = "Total"
      
      oTable.Cell(iRow, 2).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.Text = Format(dblAFee, cfmtDollar)
      oTable.Cell(iRow, 3).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.Text = Format(dblOFee, cfmtDollar)
      oTable.Cell(iRow, 4).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.Text = Format(dblDFee, cfmtDollar)
      oTable.Cell(iRow, 5).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.Text = Format(dblTFee, cfmtDollar)
      
      '帳號
      iRow = iRow + 2
      oTable.Cell(iRow, 1).Select
      
'      '設標籤
'      With .ActiveDocument.Bookmarks
'         .Add Range:=g_WordAp.Application.Selection.Range, Name:="BreakPos"
'         .DefaultSorting = wdSortByLocation
'         .ShowHidden = False
'      End With
      
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
      .Selection.Text = vbCrLf & ReportSum(71001) & vbCrLf & ReportSum(72) & vbCrLf & ReportSum(73001) & vbCrLf & ReportSum(85) & vbCrLf & ReportSum(74) & vbCrLf & ReportSum(121) & vbCrLf
      
      '建議電匯提醒
      oTable.Cell(iRow, 2).Select
      .Selection.Cells.Split NumRows:=3, NumColumns:=1, MergeBeforeSplit:=False
      .Selection.Cells(1).SetHeight RowHeight:=28, HeightRule:=wdRowHeightAtLeast
      iRow = iRow + 1
      oTable.Cell(iRow, 2).Select
      .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
      With .Selection.Cells(1)
          With .Borders(wdBorderLeft)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderRight)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderTop)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderBottom)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
      End With
      .Selection.ParagraphFormat.LeftIndent = .CentimetersToPoints(0.2)
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
      .Selection.Cells(1).SetHeight RowHeight:=52, HeightRule:=wdRowHeightAtLeast
      .Selection.Text = "Wire" & vbCrLf & "Transfer" & vbCrLf & "Preferred"
      iRow = iRow + 1
      oTable.Cell(iRow, 2).Select
      .Selection.Cells(1).SetHeight RowHeight:=0, HeightRule:=wdRowHeightAtLeast
      
      '備註
      iRow = iRow + 1
      oTable.Cell(iRow, 1).Select
      .Selection.SelectRow
      .Selection.Font.Bold = True
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
      oTable.Cell(iRow, 1).Select
      .Selection.Text = "PS:"
      oTable.Cell(iRow, 2).Select
      .Selection.Text = "Please return a copy of the invoice(s) or indicate the invoice number(s) paid with remittance"
      .Selection.EndKey
      
      .ActiveDocument.Repaginate
      '超過1頁時插入頁碼
      If .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) > 1 Then
'         .Selection.GoTo what:=wdGoToBookmark, Name:="BreakPos"
'         If .Selection.Information(wdActiveEndPageNumber) = 1 Then
'            .Selection.InsertBreak Type:=wdPageBreak
'            .Selection.TypeParagraph
'            .Selection.TypeParagraph
'            .ActiveDocument.Repaginate
'         End If
         
         If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
            .ActiveWindow.ActivePane.View.Type = wdPageView
         Else
            .ActiveWindow.View.Type = wdPageView
         End If
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
         '.Selection.TypeParagraph
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldPage
         .Selection.TypeText Text:="/"
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldEmpty, Text:="NUMPAGES ", PreserveFormatting:=True
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
      End If
'      .ActiveDocument.Bookmarks("BreakPos").Delete
      
      If Text2 = "Y" Then
         .Activate
      Else
         stPdfName = m_strNo & Text3 & ".pdf"
         
'Modified by Morgan 2022/6/8
'         If Pub_GetPrinterIndex("PDFCreator") < 0 Then
'            MsgBox "請先安裝 PDFCreator 印表機，pdf 轉檔失敗！", vbExclamation
'            Exit Function
'         Else
'            pub_OsPrinter = PUB_GetOsDefaultPrinter
'            frmPDF.Show
'            frmPDF.StartProcess m_strSavePath, stPdfName
'            PUB_SetOsDefaultPrinter Printer.DeviceName
'            PUB_SetWordActivePrinter
'            .ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
'            .ActiveDocument.Close wdDoNotSaveChanges
'            .Quit wdDoNotSaveChanges
'            frmPDF.EndtProcess
'            Unload frmPDF
'            PUB_SetOsDefaultPrinter pub_OsPrinter
'         End If
         .ActiveDocument.ExportAsFixedFormat OutputFileName:=m_strSavePath & "\" & stPdfName, ExportFormat:=17, OpenAfterExport:=False
'end 2022/6/8
         
         'Added by Morgan 2024/9/4
         .ActiveDocument.Close wdDoNotSaveChanges
         .Quit wdDoNotSaveChanges
         'end 2024/9/4
      End If
   End With
   PdfSave3 = True
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   Set rsReprot = Nothing
   
End Function

'Added by Morgan 2015/3/9
'台幣+美金格式
'Modified by Morgan 2024/9/3 +pAllDN
Private Function PdfSave4(Optional pAllDN As Boolean = False) As Boolean
   Const cFontSize = 12
   Const cfmtDollar = "#,##0.00"
   Dim oTable As Word.Table
   Dim oShape As Word.Shape
   Dim dblAFee As Double, dblOFee As Double, dblDFee As Double, dblTFee As Double, dblTot As Double
   Dim iRow As Integer
   Dim stFileName As String
   Dim stPdfName As String, stFullPath As String
   Dim rsReprot As ADODB.Recordset
   
   
On Error GoTo ErrHnd
   
   '表頭
   If Left(m_strNo, 1) = "Y" Then
      strExc(0) = "select fa05,fa63,fa64,fa65,fa32,fa33,fa34,fa35,fa36,fa18,fa19,fa20,fa21,fa22,fa70,fa17,fa23" & _
         " from fagent where fa01='" & Left(m_strNo, 8) & "' and fa02='" & Mid(m_strNo, 9) & "'"
   Else
      strExc(0) = "select cu05 as fa05,cu88 as fa63,cu89 as fa64,cu90 as fa65,cu65 as fa32, cu66 as fa33, cu67 as fa34" & _
         ", cu68 as fa35, cu69 as fa36, cu24 as fa18, cu25 as fa19, cu26 as fa20, cu27 as fa21, cu28 as fa22,cu102 fa70" & _
         ", cu23 as fa17, cu29 as fa23 from customer where cu01='" & Left(m_strNo, 8) & "' and cu02='" & Mid(m_strNo, 9) & "'"
   End If
   intI = 1
   Set rsReprot = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      MsgBox "表頭資料讀取失敗!!"
      Exit Function
   End If
   
   With rsReprot
   '代理人名稱 strexc(1)
   strExc(1) = "" & .Fields("fa05")
   If Not IsNull(.Fields("fa63")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa63")
   End If
   If Not IsNull(.Fields("fa64")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa64")
   End If
   If Not IsNull(.Fields("fa65")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa65")
   End If
   '代理人POBox/地址
   If Not IsNull(.Fields("fa32")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa32")
      If Not IsNull(.Fields("fa33")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa33")
      End If
      If Not IsNull(.Fields("fa34")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa34")
      End If
      If Not IsNull(.Fields("fa35")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa35")
      End If
      If Not IsNull(.Fields("fa36")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa36")
      End If
   ElseIf Not IsNull(.Fields("fa18")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa18")
      If Not IsNull(.Fields("fa19")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa19")
      End If
      If Not IsNull(.Fields("fa20")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa20")
      End If
      If Not IsNull(.Fields("fa21")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa21")
      End If
      If Not IsNull(.Fields("fa22")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa22")
      End If
      If Not IsNull(.Fields("fa70")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa70")
      End If
   End If
   End With
   
   '項目中文雜費者加總
   strExc(0) = "select a1k01,a1k13,to_char(to_date(a1k02+19110000,'yyyymmdd'),'FMMM/DD/yyyy') dt" & _
      ",tm45||pa77||lc23||sp27 YrRef,a1k13||'-'||a1k14||decode(a1k15||a1k16,'000','','-'||a1k16||'-'||a1k17) OrRef,a1k08,X.*" & _
      ",rtrim(decode(a2607,null,X004,a2607||' '||a2608||' '||a2609)) IDesc" & _
      " from (select a1l01,sum(amt)-sum(decode(substrb(a1l04,-2),'99',amt,0))-sum(decode(a1j03,'雜費',amt,0)) AFee" & _
      ",sum(decode(substrb(a1l04,-2),'99',amt,0)) OFee,sum(decode(a1j03,'雜費',amt,0)) DFee,sum(amt) TFee" & _
      ",min(a1k28) X001,min(a1l03) X002,substr(min(a1l02||a1l04),4) X003,substr(min(a1l02||a1j04),4) X004" & _
      " from (select a.a1l01,a.a1l02,a.a1l03,a.a1l04" & _
      ",decode( nvl(a.a1l17,0), 0, trunc((a.a1l05-nvl(a.a1l07,0)+nvl(b.a1l05,0)-nvl(b.a1l07,0)))" & _
      ",trunc((a.a1l17+nvl(b.a1l17,0))* round(1-nvl(a.a1l07,0)/a.a1l05,2))) Amt,a1j03" & _
      ",a1k28,rtrim(a1j04||' '||a1j05||' '||a1j06) a1j04" & _
      " from acc1k0,acc1l0 a,acc1l0 b,acc1j0" & _
      " where nvl(a1k12,0)=0 and a1k25" & IIf(pAllDN, "", "||a1k29") & " is null and a1k02>=" & Text3 & "01 and a1k02<=" & Text3 & "31 and a1k28='" & m_strNo & "'" & _
      " and a.a1l01(+)=a1k01 and substr(a.a1l04(+),-2)<>'98' and b.a1l01(+)=a.a1l01 and b.a1l03(+)=a.a1l03 and b.a1l04(+)=a.a1l04||'98'" & _
      " and a1j01(+)=a.a1l03 and a1j02(+)=a.a1l04" & _
      ") group by a1l01) X,acc1k0,trademark,patent,lawcase,servicepractice,acc260 where a1k01(+)=a1l01" & _
      " and tm01(+)=a1k13 and tm02(+)=a1k14 and tm03(+)=a1k15 and tm04(+)=a1k16" & _
      " and pa01(+)=a1k13 and pa02(+)=a1k14 and pa03(+)=a1k15 and pa04(+)=a1k16" & _
      " and sp01(+)=a1k13 and sp02(+)=a1k14 and sp03(+)=a1k15 and sp04(+)=a1k16" & _
      " and lc01(+)=a1k13 and lc02(+)=a1k14 and lc03(+)=a1k15 and lc04(+)=a1k16" & _
      " and a2601(+)=substr(X001,1,8) and a2602(+)=X002 and a2603(+)=X003" & _
      " order by 1,2"
   intI = 1
   Set rsReprot = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      MsgBox "請款明細資料讀取失敗!!"
      Exit Function
   End If
      
   '請款單號 strexc(2)
   strExc(2) = rsReprot("a1k01") & "/" & Text3
   
   If NewWordDoc = False Then Exit Function
   
   With g_WordAp.Application
      
      .Selection.Font.Name = "Times New Roman"
      .Selection.Font.Size = cFontSize
      
      '版面設定
      .Selection.PageSetup.PaperSize = wdPaperA4
      .Selection.PageSetup.Orientation = wdOrientLandscape
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.5)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.3)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3.2)
      .Selection.PageSetup.FooterDistance = .CentimetersToPoints(3)
      '.Selection.PageSetup.CharsLine = 40
      '.Selection.PageSetup.LinesPage = 38
      '.Selection.Orientation = wdTextOrientationHorizontal
      
      '信頭尾
      If PUB_ReadDB2File(stFileName, iPicNo) = True Then
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
         Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
         oShape.ZOrder 4
         oShape.LockAnchor = True
         oShape.LockAspectRatio = -1
         oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
         oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
         oShape.Left = 0
         oShape.Top = 0
         oShape.Width = .CentimetersToPoints(21)
         oShape.WrapFormat.Type = wdWrapNone
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         
         If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
            .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
            oShape.ZOrder 4
            oShape.LockAnchor = True
            oShape.LockAspectRatio = -1
            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
            oShape.Left = 0
            'oShape.Top = .CentimetersToPoints(27)
            oShape.Top = .CentimetersToPoints(18)
            oShape.Width = .CentimetersToPoints(21)
            oShape.WrapFormat.Type = wdWrapNone
            .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         End If
         
         .Selection.HomeKey Unit:=wdStory
      End If
      
      
      '.Selection.TypeParagraph
      '行距
      With .Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .DisableLineHeightGrid = True
      End With
      
      '新增表格(1*2)
      Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumRows:=1, NumColumns:=2)
      With oTable
         '無邊框
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
      End With
            
      oTable.Select
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop '靠上對齊
      .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
      .Selection.InsertRows 9

      '代理人名稱,POBox/地址
      oTable.Cell(1, 1).Merge oTable.Cell(5, 1)
      oTable.Cell(1, 1).Select
      .Selection.Text = strExc(1)
      
      'Attention
      strExc(0) = "Attention: Patent Invoices"
      oTable.Cell(1, 2).Select
      .Selection.Text = strExc(0)
      
      'Time period covered
      strExc(0) = "Time period covered: " & Format(ChangeTStringToWDateString(Text3 & "01"), "m/d/yyyy") & "~" & Format(ChangeWStringToWDateString(GetLastDay(Text3 & "01")), "m/d/yyyy")
      oTable.Cell(2, 2).Select
      .Selection.Text = strExc(0)
      
      If m_strNo = "Y45493000" Then 'Added by Morgan 2024/9/3 因為要共用,增加判斷Lundbeck才印
         'Lundbeck cost center
         strExc(0) = "Lundbeck cost center: 10008133"
         oTable.Cell(3, 2).Select
         .Selection.Text = strExc(0)
      End If
      
      'Invoice No
      strExc(0) = "Invoice No: " & strExc(2)
      oTable.Cell(4, 2).Select
      .Selection.Text = strExc(0)
      
      oTable.Cell(6, 1).Merge oTable.Cell(6, 2)
      oTable.Cell(6, 1).Select
      .Selection.Cells(1).SetHeight RowHeight:=30, HeightRule:=wdRowHeightAtLeast
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      .Selection.Text = "Monthly Invoice"
      
      oTable.Cell(7, 1).Select
      .Selection.SelectRow

      With .Selection.Cells
        '有邊框
        .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
        .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
      End With
      
      .Selection.Cells.Split NumRows:=1, NumColumns:=10, MergeBeforeSplit:=True
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      .Selection.Font.Size = 10
      '設定表格高度欄寬
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.7), RulerStyle:=wdAdjustProportional
      .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.6), RulerStyle:=wdAdjustProportional
      .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.6), RulerStyle:=wdAdjustProportional
      .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
      .Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(7.4), RulerStyle:=wdAdjustProportional
      .Selection.Cells(6).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
      .Selection.Cells(7).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
      .Selection.Cells(8).SetWidth ColumnWidth:=.CentimetersToPoints(2.4), RulerStyle:=wdAdjustProportional
      .Selection.Cells(9).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
      
      .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
      .Selection.InsertRows rsReprot.RecordCount + 1
      
      oTable.Cell(7, 1).Select
      .Selection.SelectRow
      .Selection.Font.Bold = True
      
      oTable.Cell(7, 1).Select
      .Selection.Text = "No."
      oTable.Cell(7, 2).Select
      .Selection.Text = "Invoice Date" & vbCrLf & "<m/d/yyyy>"
      oTable.Cell(7, 3).Select
      .Selection.Text = "Your Ref"
      oTable.Cell(7, 4).Select
      .Selection.Text = "Our Ref"
      oTable.Cell(7, 5).Select
      .Selection.Text = "Description"
      oTable.Cell(7, 6).Select
      .Selection.Text = "Attorney Fee" & vbCrLf & "(NTD)"
      oTable.Cell(7, 7).Select
      .Selection.Text = "Official Fee" & vbCrLf & "(NTD)"
      oTable.Cell(7, 8).Select
      .Selection.Text = "Disbursement" & vbCrLf & "Fee" & vbCrLf & "(NTD)"
      oTable.Cell(7, 9).Select
      .Selection.Text = "Total Fee" & vbCrLf & "(NTD)"
      oTable.Cell(7, 10).Select
      .Selection.Text = "Total Fee" & vbCrLf & "(USD)"
      
      .Selection.SelectRow
      '.Selection.Cells.Shading.BackgroundPatternColorIndex = wdTurquoise
      .Selection.Cells.Shading.Texture = wdTexture15Percent
      .Selection.Cells(1).SetHeight RowHeight:=36, HeightRule:=wdRowHeightAtLeast
      
      iRow = 7
      Do While Not rsReprot.EOF
         iRow = iRow + 1
         'No.
         oTable.Cell(iRow, 1).Select
         .Selection.Text = iRow - 7
         'Invoice Date
         oTable.Cell(iRow, 2).Select
         .Selection.Text = "" & rsReprot("dt")
         'Your Ref
         strExc(1) = "" & rsReprot("YrRef")
         If GetXYrRef(rsReprot("a1k13"), rsReprot("A1K01"), strExc(2)) = True Then
            strExc(1) = strExc(2)
         End If
         oTable.Cell(iRow, 3).Select
         .Selection.Text = strExc(1)
         'Our Ref
         oTable.Cell(iRow, 4).Select
         .Selection.Text = "" & rsReprot("OrRef")
         'Description
         oTable.Cell(iRow, 5).Select
         .Selection.Text = "" & rsReprot("IDesc")
         'Attorney Fee
         oTable.Cell(iRow, 6).Select
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Text = Format(Val("" & rsReprot("AFee")), cfmtDollar)
         'Official Fee
         oTable.Cell(iRow, 7).Select
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Text = Format(Val("" & rsReprot("OFee")), cfmtDollar)
         'Disbursement
         oTable.Cell(iRow, 8).Select
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Text = Format(Val("" & rsReprot("DFee")), cfmtDollar)
         'Total Fee
         oTable.Cell(iRow, 9).Select
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Text = Format(Val("" & rsReprot("TFee")), cfmtDollar)
         
         oTable.Cell(iRow, 10).Select
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Text = Format(Val("" & rsReprot("a1k08")), cfmtDollar)
         
         
         dblAFee = dblAFee + Val("" & rsReprot("AFee"))
         dblOFee = dblOFee + Val("" & rsReprot("OFee"))
         dblDFee = dblDFee + Val("" & rsReprot("DFee"))
         dblTFee = dblTFee + Val("" & rsReprot("TFee"))
         dblTot = dblTot + Val("" & rsReprot("a1k08"))
         rsReprot.MoveNext
      Loop
      
      iRow = iRow + 1
      oTable.Cell(iRow, 1).Merge oTable.Cell(iRow, 5)
      oTable.Cell(iRow, 1).Select
      .Selection.SelectRow
      .Selection.Font.Bold = True
      
      oTable.Cell(iRow, 1).Select
      .Selection.Text = "Total"
      
      oTable.Cell(iRow, 2).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.Text = Format(dblAFee, cfmtDollar)
      oTable.Cell(iRow, 3).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.Text = Format(dblOFee, cfmtDollar)
      oTable.Cell(iRow, 4).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.Text = Format(dblDFee, cfmtDollar)
      oTable.Cell(iRow, 5).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.Text = Format(dblTFee, cfmtDollar)
      oTable.Cell(iRow, 6).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.Text = Format(dblTot, cfmtDollar)
      
      '帳號
      iRow = iRow + 2
      oTable.Cell(iRow, 1).Select
      
'      '設標籤
'      With .ActiveDocument.Bookmarks
'         .Add Range:=g_WordAp.Application.Selection.Range, Name:="BreakPos"
'         .DefaultSorting = wdSortByLocation
'         .ShowHidden = False
'      End With
      
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
      .Selection.Text = vbCrLf & ReportSum(71001) & vbCrLf & ReportSum(72) & vbCrLf & ReportSum(73001) & vbCrLf & ReportSum(85) & vbCrLf & ReportSum(74) & vbCrLf & ReportSum(121) & vbCrLf
      
      '建議電匯提醒
      oTable.Cell(iRow, 2).Select
      .Selection.Cells.Split NumRows:=3, NumColumns:=1, MergeBeforeSplit:=False
      .Selection.Cells(1).SetHeight RowHeight:=28, HeightRule:=wdRowHeightAtLeast
      iRow = iRow + 1
      oTable.Cell(iRow, 2).Select
      .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
      With .Selection.Cells(1)
          With .Borders(wdBorderLeft)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderRight)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderTop)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderBottom)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
      End With
      .Selection.ParagraphFormat.LeftIndent = .CentimetersToPoints(0.2)
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
      .Selection.Cells(1).SetHeight RowHeight:=52, HeightRule:=wdRowHeightAtLeast
      .Selection.Text = "Wire" & vbCrLf & "Transfer" & vbCrLf & "Preferred"
      iRow = iRow + 1
      oTable.Cell(iRow, 2).Select
      .Selection.Cells(1).SetHeight RowHeight:=0, HeightRule:=wdRowHeightAtLeast
      
      '備註
      iRow = iRow + 1
      oTable.Cell(iRow, 1).Select
      .Selection.SelectRow
      .Selection.Font.Bold = True
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
      oTable.Cell(iRow, 1).Select
      .Selection.Text = "PS:"
      oTable.Cell(iRow, 2).Select
      .Selection.Text = "Please return a copy of the invoice(s) or indicate the invoice number(s) paid with remittance"
      .Selection.EndKey
      
      .ActiveDocument.Repaginate
      '超過1頁時插入頁碼
      If .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) > 1 Then
'         .Selection.GoTo what:=wdGoToBookmark, Name:="BreakPos"
'         If .Selection.Information(wdActiveEndPageNumber) = 1 Then
'            .Selection.InsertBreak Type:=wdPageBreak
'            .Selection.TypeParagraph
'            .Selection.TypeParagraph
'            .ActiveDocument.Repaginate
'         End If
         
         If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
            .ActiveWindow.ActivePane.View.Type = wdPageView
         Else
            .ActiveWindow.View.Type = wdPageView
         End If
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
         '.Selection.TypeParagraph
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldPage
         .Selection.TypeText Text:="/"
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldEmpty, Text:="NUMPAGES ", PreserveFormatting:=True
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
      End If
      
'      .ActiveDocument.Bookmarks("BreakPos").Delete
      If Text2 = "Y" Then
         .Activate
      Else
         stPdfName = m_strNo & Text3 & ".pdf"
         
'Modified by Morgan 2022/6/8
'         If Pub_GetPrinterIndex("PDFCreator") < 0 Then
'            MsgBox "請先安裝 PDFCreator 印表機，pdf 轉檔失敗！", vbExclamation
'            Exit Function
'         Else
'            pub_OsPrinter = PUB_GetOsDefaultPrinter
'            frmPDF.Show
'            frmPDF.StartProcess m_strSavePath, stPdfName
'            PUB_SetOsDefaultPrinter Printer.DeviceName
'            PUB_SetWordActivePrinter
'            .ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
'            .ActiveDocument.Close wdDoNotSaveChanges
'            .Quit wdDoNotSaveChanges
'            frmPDF.EndtProcess
'            Unload frmPDF
'            PUB_SetOsDefaultPrinter pub_OsPrinter
'         End If
         .ActiveDocument.ExportAsFixedFormat OutputFileName:=m_strSavePath & "\" & stPdfName, ExportFormat:=17, OpenAfterExport:=False
'end 2022/6/8

         'Added by Morgan 2024/9/4
         .ActiveDocument.Close wdDoNotSaveChanges
         .Quit wdDoNotSaveChanges
         'end 2024/9/4
      End If
   End With
   PdfSave4 = True
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   Set rsReprot = Nothing
   
End Function

'Added by Morgan 2018/6/26
Private Function SaveLEDES() As Boolean
   Dim strLedes() As String
   Dim strTmp() As String
   Dim iUpper As Integer, ii As Integer, jj As Integer
   Dim rsQuery As ADODB.Recordset
   Dim rsLEDES As ADODB.Recordset
   Dim iRows As Integer, iCols As Integer
   Dim lngTot As Long
   Dim strError As String
   Dim strMaxNo As String 'Added by Morgan 2022/8/18
   Dim strConA1K28 As String 'Added by Morgan 2023/7/12
   
On Error GoTo ErrHnd
   
   iUpper = 0
   
   'Modified by Morgan 2023/7/12 Y48279000要同時抓X48279000 --Kahn
   If m_strNo = "Y48279000" Then
      strConA1K28 = " and (a1k28='" & m_strNo & "' or a1k28='X48279000')"
   
   'Added by Morgan 2025/6/30--Tim
   '新增Y22457選項(將指定月份請款對象為 Y22457000,Y48048000,Y52322000,Y52322B10 的帳單合併為一張ledes帳單)
   ElseIf m_strNo = "Y22457000" Then
      strConA1K28 = " and a1k28 in ('Y22457000','Y48048000','Y52322000','Y52322B10')"
   Else
      strConA1K28 = " and a1k28='" & m_strNo & "'"
   End If
   'end 2023/7/12
   
   strExc(0) = "select a1k01,a1k08 from acc1k0 where a1k02>=" & Text3 & "01 and a1k02<=" & Text3 & "31" & _
      " and nvl(a1k12,0)=0 and a1k25 is null" & strConA1K28
      
   'Modified by Morgan 2020/11/9 Y51971000 不必排除已結清請款單(LEDES檔是報告用，請款是用pdf檔給請款對象)--莊瑄凡
   'Modified by Morgan 2021/3/16 Y20412010 也不必排除已結清 --莊瑄凡
   'Modified by Morgan 2023/8/1 +Y51467020 --Franny
   If m_strNo <> "Y51971000" And m_strNo <> "Y20412010" And m_strNo <> "Y51467020" Then
      strExc(0) = strExc(0) & " and a1k29 is null"
   End If
   'end 2020/11/9
   strExc(0) = strExc(0) & "  order by 1"
      
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      InsertQueryLog rsQuery.RecordCount
      SetOutPutPath
      
      With rsQuery
      .MoveFirst
      Do While Not .EOF
         lngTot = lngTot + .Fields("a1k08")
         strMaxNo = .Fields("a1k01") 'Added by Morgan 2022/8/18
         .MoveNext
      Loop
      
      .MoveFirst
      Do While Not .EOF
         Load Frmacc2480
         With Frmacc2480
            .Visible = False
            .Text1.Text = rsQuery("a1k01")
            .Text2.Text = .Text1.Text
            .Check2.Value = vbChecked
            .m_bLedesOnly = True
            .Command2.Value = True
            strError = .m_sEBillingMsg
            .GetLedes strTmp
         End With
         Unload Frmacc2480
         If strError <> "" Then
            MsgBox strError
            GoTo ErrHnd
         Else
            iRows = UBound(strTmp, 2)
            iCols = UBound(strTmp, 1)
            
            For ii = 1 To iRows
               iUpper = iUpper + 1
               ReDim Preserve strLedes(iCols, iUpper)
               'Added by Morgan 2023/3/14
               'DuPont一案號一請款單(一LEDES檔多請款單)
               'Modified by Morgan 2023/7/10 +Y54570000,Y48279000
               'Modified by Morgan 2023/12/20 +X82995000 --Kahn
               'Modified by Morgan 2025/6/19 +Y22327000 --Lisa
               'Modified by Morgan 2025/6/30 +Y22457000 --Tim
               If m_strNo = "Y55240000" Or m_strNo = "Y54570000" Or m_strNo = "Y48279000" Or m_strNo = "X82995000" Or m_strNo = "Y22327000" Or m_strNo = "Y22457000" Then
                  If iUpper = 1 Then strLedes(1, 0) = strTmp(1, 0) 'LEDES版本
                  For jj = 1 To iCols
                     strLedes(jj, iUpper) = strTmp(jj, ii)
                  Next
                  strLedes(9, iUpper) = iUpper
               'end 2023/3/14
               '只要設定第一筆,其他相同
               ElseIf iUpper = 1 Then
                  strLedes(1, 0) = strTmp(1, 0) 'LEDES版本
                  'Added by Morgan 2022/8/18
                  If m_strNo = "Y48292000" Then
                     '1 INVOICE_DATE
                     strLedes(1, iUpper) = strTmp(1, ii)
                     '2 'INVOICE_NUMBER
                     strLedes(2, iUpper) = strTmp(2, ii)
                     If strLedes(2, iUpper) <> strMaxNo Then
                        strLedes(2, iUpper) = strLedes(2, iUpper) & "TO" & Right(strMaxNo, 3)
                     End If
                  Else
                  'end 2022/8/18
                  
                     '1 INVOICE_DATE
                     'strLedes(1, iUpper) = PUB_GetWorkDay1(CompDate(2, -1, CompDate(1, 1, Text3 & "01")), True) '最後一個工作天
                     strLedes(1, iUpper) = CompDate(2, -1, CompDate(1, 1, Text3 & "01")) '最後一天
                     '2 'INVOICE_NUMBER
                     strLedes(2, iUpper) = .Fields("a1k01") & "/" & Text3 '第１張請款單號/年月
                  End If
                  
                  '3 CLIENT_ID
                  strLedes(3, iUpper) = strTmp(3, ii)
                  '5 INVOICE_TOTAL
                  strLedes(5, iUpper) = lngTot
                  
                  If m_strNo = "Y48292000" Then
                     '6 BILLING_START_DATE
                     strLedes(6, iUpper) = strTmp(6, ii)
                     '7 BILLING_END_DATE
                     strLedes(7, iUpper) = strTmp(7, ii)
                  Else
                     '6 BILLING_START_DATE
                     strLedes(6, iUpper) = DBDATE(Text3 & "01") '請款月1號
                     '7 BILLING_END_DATE
                     strLedes(7, iUpper) = strLedes(1, iUpper) 'INVOICE_DATE
                  End If
                  '8 INVOICE_DESCRIPTION
                  'Modified by Morgan 2020/9/21
                  'strLedes(8, iUpper) = "[" & Text1 & "]"
                  If m_strNo = "Y54225B10" Then
                     strLedes(8, iUpper) = "[" & Text1 & "]"
                     'Added by Morgan 2021/2/4 改用 98BI
                     '4 LAW_FIRM_MATTER_ID
                     strLedes(4, iUpper) = strLedes(2, iUpper)
                     '29 INVOICE_NET_TOTAL
                     strLedes(29, iUpper) = lngTot
                     'end 2021/2/4
                  Else
                     strLedes(8, iUpper) = strTmp(8, ii)
                  End If
                  'end 2020/9/21
                  '20 LAW_FIRM_ID
                  strLedes(20, iUpper) = strTmp(20, ii)
               Else
                  strLedes(1, iUpper) = strLedes(1, 1)
                  strLedes(2, iUpper) = strLedes(2, 1)
                  strLedes(3, iUpper) = strLedes(3, 1)
                  strLedes(5, iUpper) = strLedes(5, 1)
                  strLedes(6, iUpper) = strLedes(6, 1)
                  strLedes(7, iUpper) = strLedes(7, 1)
                  strLedes(8, iUpper) = strLedes(8, 1)
                  strLedes(20, iUpper) = strLedes(20, 1)
                  
                  'Added by Morgan 2021/2/4
                  If m_strNo = "Y54225B10" Then
                     strLedes(4, iUpper) = strLedes(4, 1)
                     strLedes(29, iUpper) = strLedes(29, 1)
                  End If
                  'end 2021/2/4
               End If
               
               '4 LAW_FIRM_MATTER_ID(本所案號)
               'Modified by Morgan 2020/9/21
               'Modified by Morgan 2021/3/16 +Y20412010 --Franny
               'Modified by Morgan 2023/8/1 +Y51467020 --Franny
               If m_strNo = "Y51971000" Or m_strNo = "Y20412010" Or m_strNo = "Y51467020" Then
                  strLedes(4, iUpper) = strTmp(4, ii) & "-TW"
               Else
                  If m_strNo <> "Y54225B10" Then 'Added by Morgan 2021/2/4
                     strLedes(4, iUpper) = strTmp(4, ii)
                  End If
               End If
               '9 LINE_ITEM_NUMBER(項次)
               strLedes(9, iUpper) = iUpper
               '10 EXP/FEE/INV_ADJ_TYPE(項目類別)
               strLedes(10, iUpper) = strTmp(10, ii)
               '11 LINE_ITEM_NUMBER_OF_UNITS
               strLedes(11, iUpper) = strTmp(11, ii)
               '12 LINE_ITEM_ADJUSTMENT_AMOUNT
               strLedes(12, iUpper) = strTmp(12, ii)
               '13 LINE_ITEM_TOTAL
               strLedes(13, iUpper) = strTmp(13, ii)
               '14 LINE_ITEM_DATE
               strLedes(14, iUpper) = strTmp(1, ii) '請款日期
               '15 LINE_ITEM_TASK_CODE
               strLedes(15, iUpper) = strTmp(15, ii)
               '16 LINE_ITEM_EXPENSE_CODE
               strLedes(16, iUpper) = strTmp(16, ii)
               '17 LINE_ITEM_ACTIVITY_CODE
               strLedes(17, iUpper) = strTmp(17, ii)
               '18 TIMEKEEPER_ID
               strLedes(18, iUpper) = strTmp(18, ii)
               '19 LINE_ITEM_DESCRIPTION
               strLedes(19, iUpper) = strTmp(19, ii)
               '21 LINE_ITEM_UNIT_COST(單價)
               strLedes(21, iUpper) = strTmp(21, ii)
               '22 TIMEKEEPER_NAME
               strLedes(22, iUpper) = strTmp(22, ii)
               '23 TIMEKEEPER_CLASSIFICATION
               strLedes(23, iUpper) = strTmp(23, ii)
               '24 CLIENT_MATTER_ID
               strLedes(24, iUpper) = strTmp(24, ii)
               
               'LEDES98BI欄位
               For jj = 25 To iCols
                  If jj <> 29 Then 'Added by Morgan 2021/2/4
                     strLedes(jj, iUpper) = strTmp(jj, ii)
                  End If
               Next
            Next ii
         End If
         .MoveNext
      Loop
      End With
      SaveLEDES = WriteLEDES(strLedes)
   Else
      InsertQueryLog rsQuery.RecordCount
      MsgBox "無符合資料！"
   End If
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   Set rsQuery = Nothing
   Set rsLEDES = Nothing
End Function

'Added by Morgan 2018/11/6
'目前僅適用 Sandoz
Private Function SaveLEDES2() As Boolean
   Dim strLedes() As String
   Dim strTmp() As String
   Dim iUpper As Integer, ii As Integer, jj As Integer, kk As Integer
   Dim rsQuery As ADODB.Recordset
   Dim rsLEDES As ADODB.Recordset
   Dim iRows As Integer, iCols As Integer
   Dim lngTot As Long
   Dim strError As String
   Dim strDate As String, strDNoList As String, iCount As Integer
   Dim arrNo() As String
   
On Error GoTo ErrHnd
   
   '只抓請款項目有查名的請款單,同一請款日合併為一個電子帳單,新請款單號=最小請款單號/請款月日(Ex:X10709773/0704)
   strExc(0) = "select a1k01,a1k02,a1k08 from acc1k0,acc1l0 where a1k02>=" & Text3 & "01 and a1k02<=" & Text3 & "31" & _
      " and a1k28='" & m_strNo & "' and nvl(a1k12,0)=0 and a1k25||a1k29 is null and a1l01(+)=a1k01 and a1l04='001' order by a1k02,a1k01"
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      InsertQueryLog rsQuery.RecordCount
      SetOutPutPath
      
      With rsQuery
      .MoveFirst
      
   Do While Not .EOF
      iUpper = 0
      lngTot = 0
      iCount = 0
      strDNoList = ""
      strDate = .Fields("a1k02")
      Do While Not .EOF
         If strDate = .Fields("a1k02") Then
            iCount = iCount + 1
            lngTot = lngTot + .Fields("a1k08")
            strDNoList = strDNoList & .Fields("a1k01") & ";"
         Else
            Exit Do
         End If
         .MoveNext
      Loop
      
      arrNo = Split(strDNoList, ";")
      For kk = LBound(arrNo) To UBound(arrNo)
         If arrNo(kk) <> "" Then
            Load Frmacc2480
            With Frmacc2480
               .Visible = False
               .Text1.Text = arrNo(kk)
               .Text2.Text = .Text1.Text
               .Check2.Value = vbChecked
               .m_bLedesOnly = True
               .Command2.Value = True
               strError = .m_sEBillingMsg
               .GetLedes strTmp
            End With
            Unload Frmacc2480
            If strError <> "" Then
               MsgBox strError
               GoTo ErrHnd
            Else
               iRows = UBound(strTmp, 2)
               iCols = UBound(strTmp, 1)
               
               For ii = 1 To iRows
                  iUpper = iUpper + 1
                  ReDim Preserve strLedes(iCols, iUpper)
                  '只要設定第一筆,其他相同
                  If iUpper = 1 Then
                     strLedes(1, 0) = strTmp(1, 0) 'LEDES版本
                     '1 INVOICE_DATE
                     strLedes(1, 1) = strTmp(1, 1)
                     '2 'INVOICE_NUMBER
                     If iCount = 1 Then
                        strLedes(2, 1) = strTmp(2, 1)
                     Else
                        strLedes(2, 1) = strTmp(2, 1) & "/" & Right(strTmp(1, 1), 4)
                     End If
                     '3 CLIENT_ID
                     strLedes(3, 1) = strTmp(3, 1)
                     '5 INVOICE_TOTAL
                     strLedes(5, 1) = lngTot
                     '6 BILLING_START_DATE
                     strLedes(6, 1) = strTmp(6, 1)
                     '7 BILLING_END_DATE
                     strLedes(7, 1) = strTmp(7, 1) 'INVOICE_DATE
                     '8 INVOICE_DESCRIPTION
                     strLedes(8, 1) = strTmp(8, 1)
                     '20 LAW_FIRM_ID
                     strLedes(20, 1) = strTmp(20, 1)
                  Else
                     strLedes(1, iUpper) = strLedes(1, 1)
                     strLedes(2, iUpper) = strLedes(2, 1)
                     strLedes(3, iUpper) = strLedes(3, 1)
                     strLedes(5, iUpper) = strLedes(5, 1)
                     strLedes(6, iUpper) = strLedes(6, 1)
                     strLedes(7, iUpper) = strLedes(7, 1)
                     strLedes(8, iUpper) = strLedes(8, 1)
                     strLedes(20, iUpper) = strLedes(20, 1)
                  End If
                  '4 LAW_FIRM_MATTER_ID(本所案號)
                  strLedes(4, iUpper) = strTmp(4, ii)
                  '9 LINE_ITEM_NUMBER(項次)
                  strLedes(9, iUpper) = iUpper
                  '10 EXP/FEE/INV_ADJ_TYPE(項目類別)
                  strLedes(10, iUpper) = strTmp(10, ii)
                  '11 LINE_ITEM_NUMBER_OF_UNITS
                  strLedes(11, iUpper) = strTmp(11, ii)
                  '12 LINE_ITEM_ADJUSTMENT_AMOUNT
                  strLedes(12, iUpper) = strTmp(12, ii)
                  '13 LINE_ITEM_TOTAL
                  strLedes(13, iUpper) = strTmp(13, ii)
                  '14 LINE_ITEM_DATE
                  strLedes(14, iUpper) = strTmp(1, ii) '請款日期
                  '15 LINE_ITEM_TASK_CODE
                  strLedes(15, iUpper) = strTmp(15, ii)
                  '16 LINE_ITEM_EXPENSE_CODE
                  strLedes(16, iUpper) = strTmp(16, ii)
                  '17 LINE_ITEM_ACTIVITY_CODE
                  strLedes(17, iUpper) = strTmp(17, ii)
                  '18 TIMEKEEPER_ID
                  strLedes(18, iUpper) = strTmp(18, ii)
                  '19 LINE_ITEM_DESCRIPTION
                  strLedes(19, iUpper) = strTmp(19, ii)
                  '21 LINE_ITEM_UNIT_COST(單價)
                  strLedes(21, iUpper) = strTmp(21, ii)
                  '22 TIMEKEEPER_NAME
                  strLedes(22, iUpper) = strTmp(22, ii)
                  '23 TIMEKEEPER_CLASSIFICATION
                  strLedes(23, iUpper) = strTmp(23, ii)
                  '24 CLIENT_MATTER_ID
                  strLedes(24, iUpper) = strTmp(24, ii)
                  
                  'LEDES98BI欄位(目前沒用)
                  For jj = 25 To iCols
                     strLedes(jj, iUpper) = strTmp(jj, ii)
                  Next
               Next ii
            End If
         End If
      Next
      SaveLEDES2 = WriteLEDES(strLedes, m_strNo & strDate)
   Loop
   End With
   
   Else
      InsertQueryLog rsQuery.RecordCount
      MsgBox "無符合資料！"
   End If
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   Set rsQuery = Nothing
   Set rsLEDES = Nothing
End Function

Private Function WriteLEDES(pLedes() As String, Optional pFileName As String) As Boolean
   Dim strPath As String, strFile As String
   Dim F1 As Integer, ii As Integer, jj As Integer
   Dim stOut As String
   
   SetOutPutPath
   If pFileName <> "" Then
      strFile = pFileName & ".txt"
   Else
      strFile = m_strNo & Text3 & ".txt"
   End If
   strPath = m_strSavePath & "\" & strFile
   
   F1 = FreeFile
   Open strPath For Output As F1
   
   If pLedes(1, 0) = 2 Then
      Print #F1, "LEDES98BI V2[]"
      Print #F1, "INVOICE_DATE|INVOICE_NUMBER|CLIENT_ID|LAW_FIRM_MATTER_ID|INVOICE_TOTAL|BILLING_START_DATE|BILLING_END_DATE|INVOICE_DESCRIPTION|LINE_ITEM_NUMBER|EXP/FEE/INV_ADJ_TYPE|LINE_ITEM_NUMBER_OF_UNITS|LINE_ITEM_ADJUSTMENT_AMOUNT|LINE_ITEM_TOTAL|LINE_ITEM_DATE|LINE_ITEM_TASK_CODE|LINE_ITEM_EXPENSE_CODE|LINE_ITEM_ACTIVITY_CODE|TIMEKEEPER_ID|LINE_ITEM_DESCRIPTION|LAW_FIRM_ID|LINE_ITEM_UNIT_COST|TIMEKEEPER_NAME|TIMEKEEPER_CLASSIFICATION|CLIENT_MATTER_ID|PO_NUMBER|CLIENT_TAX_ID|MATTER_NAME|INVOICE_TAX_TOTAL|INVOICE_NET_TOTAL|INVOICE_CURRENCY|TIMEKEEPER_LAST_NAME|TIMEKEEPER_FIRST_NAME|ACCOUNT_TYPE|LAW_FIRM_NAME|LAW_FIRM_ADDRESS_1|LAW_FIRM_ADDRESS_2|LAW_FIRM_CITY|LAW_FIRM_STATEorREGION|LAW_FIRM_POSTCODE|LAW_FIRM_COUNTRY|CLIENT_NAME|CLIENT_ADDRESS_1|CLIENT_ADDRESS_2|CLIENT_CITY|CLIENT_STATEorREGION|CLIENT_POSTCODE|CLIENT_COUNTRY|LINE_ITEM_TAX_RATE|LINE_ITEM_TAX_TOTAL|LINE_ITEM_TAX_TYPE|INVOICE_REPORTED_TAX_TOTAL|INVOICE_TAX_CURRENCY[]"
   Else
      Print #F1, "LEDES1998B[]"
      Print #F1, "INVOICE_DATE|INVOICE_NUMBER|CLIENT_ID|LAW_FIRM_MATTER_ID|INVOICE_TOTAL|BILLING_START_DATE|BILLING_END_DATE|INVOICE_DESCRIPTION|LINE_ITEM_NUMBER|EXP/FEE/INV_ADJ_TYPE|LINE_ITEM_NUMBER_OF_UNITS|LINE_ITEM_ADJUSTMENT_AMOUNT|LINE_ITEM_TOTAL|LINE_ITEM_DATE|LINE_ITEM_TASK_CODE|LINE_ITEM_EXPENSE_CODE|LINE_ITEM_ACTIVITY_CODE|TIMEKEEPER_ID|LINE_ITEM_DESCRIPTION|LAW_FIRM_ID|LINE_ITEM_UNIT_COST|TIMEKEEPER_NAME|TIMEKEEPER_CLASSIFICATION|CLIENT_MATTER_ID[]"
   End If
   
   For ii = 1 To UBound(pLedes, 2)
      stOut = pLedes(1, ii)
      For jj = 2 To UBound(pLedes, 1)
         stOut = stOut & "|" & pLedes(jj, ii)
      Next
      stOut = stOut & "[]"
      Print #F1, stOut
   Next
   Close #F1
   
   WriteLEDES = True
   
End Function

Private Sub AddBorder(pWks As Worksheet, pFromCell As String, pToCell As String)
  '加邊框
   With pWks.Range(pFromCell, pToCell)
      .Borders(xlDiagonalDown).LineStyle = xlNone
      .Borders(xlDiagonalUp).LineStyle = xlNone
      
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeLeft).ColorIndex = xlAutomatic
      .Borders(xlEdgeLeft).tintandshade = 0
      .Borders(xlEdgeLeft).Weight = xlThin
      
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeTop).ColorIndex = xlAutomatic
      .Borders(xlEdgeTop).tintandshade = 0
      .Borders(xlEdgeTop).Weight = xlThin
      
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
      .Borders(xlEdgeBottom).tintandshade = 0
      .Borders(xlEdgeBottom).Weight = xlThin
      
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeRight).ColorIndex = xlAutomatic
      .Borders(xlEdgeRight).tintandshade = 0
      .Borders(xlEdgeRight).Weight = xlThin
      
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideVertical).ColorIndex = xlAutomatic
      .Borders(xlInsideVertical).tintandshade = 0
      .Borders(xlInsideVertical).Weight = xlThin
      
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
      .Borders(xlInsideHorizontal).tintandshade = 0
      .Borders(xlInsideHorizontal).Weight = xlThin
   End With
End Sub

'Added by Morgan 2020/10/6
'Y22327 MKS 月帳單首頁+明細
Private Sub runWordCoverPage(ByRef pWordAp As Word.Application, ByRef pRst As ADODB.Recordset)
   Dim iPicNo As Integer, iPicNo2 As Integer
   Dim iRowCount As Integer, ii As Integer
   Dim stFileName As String
   Dim oShape
   Dim dblAFee As Double, dblOFee As Double
   
On Error GoTo ErrHnd
   
   iRowCount = 0
   With pWordAp
      '版面設定
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.5)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
      .Selection.PageSetup.FooterDistance = .CentimetersToPoints(3)
      
      .Selection.PageSetup.CharsLine = 40
      .Selection.PageSetup.LinesPage = 38
      
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = 12
      '保留信頭空間
      .Selection.TypeParagraph
      '行距
      With .Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .DisableLineHeightGrid = True
      End With
      
      '新增表格(1*1)
      .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=1
      
      With .Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
      End With
      
      '設定表格高度欄寬
      .Selection.SelectRow
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      
      .Selection.InsertRows 1
      .Selection.Collapse Direction:=wdCollapseStart
      iRowCount = 0
      '列印對象
      strExc(0) = "select cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu27 as fa21, cu28 as fa22, cu68 as fa35, cu69 as fa36, cu06 as fa06, cu29 as fa23" & _
         " From customer where cu01='X4584600' and cu02='0'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
         '名稱
         If IsNull(RsTemp.Fields("fa05").Value) = False Then
            .Selection.TypeText Text:=RsTemp.Fields("fa05").Value
         End If
         If IsNull(RsTemp.Fields("fa63").Value) = False Then
            .Selection.MoveDown Unit:=wdLine, Count:=1
            .Selection.InsertRows 1
            .Selection.TypeText Text:=RsTemp.Fields("fa63").Value
         End If
         If IsNull(RsTemp.Fields("fa64").Value) = False Then
            .Selection.MoveDown Unit:=wdLine, Count:=1
            .Selection.InsertRows 1
            .Selection.TypeText Text:=RsTemp.Fields("fa64").Value
         End If
         If IsNull(RsTemp.Fields("fa65").Value) = False Then
            .Selection.MoveDown Unit:=wdLine, Count:=1
            .Selection.InsertRows 1
            .Selection.TypeText Text:=RsTemp.Fields("fa65").Value
         End If
         
         'POB
         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.InsertRows 1
         .Selection.Collapse Direction:=wdCollapseStart
         If IsNull(RsTemp.Fields("fa32").Value) = False Then
            .Selection.TypeText Text:=RsTemp.Fields("fa32").Value
            If IsNull(RsTemp.Fields("fa33").Value) = False Then
               .Selection.MoveDown Unit:=wdLine, Count:=1
               .Selection.InsertRows 1
               .Selection.TypeText Text:=RsTemp.Fields("fa33").Value
               iRowCount = iRowCount + 1
            End If
            If IsNull(RsTemp.Fields("fa34").Value) = False Then
               .Selection.MoveDown Unit:=wdLine, Count:=1
               .Selection.InsertRows 1
               .Selection.TypeText Text:=RsTemp.Fields("fa34").Value
               iRowCount = iRowCount + 1
            End If
            If IsNull(RsTemp.Fields("fa35").Value) = False Then
               .Selection.MoveDown Unit:=wdLine, Count:=1
               .Selection.InsertRows 1
               .Selection.TypeText Text:=RsTemp.Fields("fa35").Value
               iRowCount = iRowCount + 1
            End If
            If IsNull(RsTemp.Fields("fa36").Value) = False Then
               .Selection.MoveDown Unit:=wdLine, Count:=1
               .Selection.InsertRows 1
               .Selection.TypeText Text:=RsTemp.Fields("fa36").Value
               iRowCount = iRowCount + 1
            End If
         '地址
         ElseIf IsNull(RsTemp.Fields("fa18").Value) = False Then
            .Selection.TypeText Text:=RsTemp.Fields("fa18").Value
            If IsNull(RsTemp.Fields("fa19").Value) = False Then
               .Selection.MoveDown Unit:=wdLine, Count:=1
               .Selection.InsertRows 1
               .Selection.TypeText Text:=RsTemp.Fields("fa19").Value
               iRowCount = iRowCount + 1
            End If
            If IsNull(RsTemp.Fields("fa20").Value) = False Then
               .Selection.MoveDown Unit:=wdLine, Count:=1
               .Selection.InsertRows 1
               .Selection.TypeText Text:=RsTemp.Fields("fa20").Value
               iRowCount = iRowCount + 1
            End If
            If IsNull(RsTemp.Fields("fa21").Value) = False Then
               .Selection.MoveDown Unit:=wdLine, Count:=1
               .Selection.InsertRows 1
               .Selection.TypeText Text:=RsTemp.Fields("fa21").Value
               iRowCount = iRowCount + 1
            End If
            If IsNull(RsTemp.Fields("fa22").Value) = False Then
               .Selection.MoveDown Unit:=wdLine, Count:=1
               .Selection.InsertRows 1
               .Selection.TypeText Text:=RsTemp.Fields("fa22").Value
               iRowCount = iRowCount + 1
            End If
         End If
      End If
      
      .Selection.MoveDown Unit:=wdLine, Count:=1
      .Selection.InsertRows 3
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.MoveDown Unit:=wdLine, Count:=2
      .Selection.TypeText Text:="Date: "
      .Selection.Font.Bold = True
      .Selection.TypeText Text:=Format(Now, "mmmm dd, yyyy")
      
      .Selection.MoveDown Unit:=wdLine, Count:=1
      .Selection.InsertRows 3
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.MoveDown Unit:=wdLine, Count:=2
      
      .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(12), RulerStyle:=wdAdjustProportional
      .Selection.InsertRows 1
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.TypeText Text:="Re: Monthly Invoice "
      .Selection.Font.Bold = True
      strExc(2) = Format(ChangeTStringToWDateString(Text3 & "01"), "mmmm yyyy")
      .Selection.TypeText Text:="(" & strExc(2) & ")"
      .Selection.MoveDown Unit:=wdLine, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.TypeText Text:="Invoice Number: "
      .Selection.Font.Bold = True
      .Selection.TypeText Text:="Y22327" & (Val(Text3) + 191100)
            
      .Selection.MoveDown Unit:=wdLine, Count:=1
      .Selection.InsertRows 3
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.MoveDown Unit:=wdLine, Count:=2
      .Selection.Collapse Direction:=wdCollapseStart
      strExc(1) = "Please find below the service fees incurred for our professional services rendered in "
      strExc(1) = strExc(1) & "connection with your matters during the month of " & strExc(2) & ". The separate detailed "
      strExc(1) = strExc(1) & "invoices for the listed cases are attached herein."
      .Selection.TypeText Text:=strExc(1)
            
      .Selection.MoveDown Unit:=wdLine, Count:=1
      .Selection.InsertRows 3
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.MoveDown Unit:=wdLine, Count:=1
            
      With .Selection.Cells(1)
         With .Borders(wdBorderLeft)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderRight)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderTop)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderBottom)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
      End With
            
      .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=False
      .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(3.2), RulerStyle:=wdAdjustProportional
      .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(3.2), RulerStyle:=wdAdjustProportional
      .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(4.2), RulerStyle:=wdAdjustProportional

      .Selection.InsertRows pRst.RecordCount + 2
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText "Your Ref."
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText "Our Ref."
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText "Service Fees (USD)"
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText "Disbursements: Official" & vbCrLf & "Fees & Other Charges (USD)"
      
      pRst.MoveFirst
      dblAFee = 0
      dblOFee = 0
      Do While Not pRst.EOF
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.TypeText "" & pRst.Fields("YrRef")
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.TypeText "" & pRst.Fields("OrRef")
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Format(Val("" & pRst.Fields("AFee")), "#,##0")
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Format(Val("" & pRst.Fields("OFee")), "#,##0")
         dblAFee = dblAFee + Val("" & pRst.Fields("AFee"))
         dblOFee = dblOFee + Val("" & pRst.Fields("OFee"))
         pRst.MoveNext
      Loop
      .Selection.MoveRight Unit:=wdCharacter, Count:=2
      .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
      .Selection.Cells.Merge
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.TypeText "Sub-Total"
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Format(dblAFee, "#,##0")
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Format(dblOFee, "#,##0")
      .Selection.MoveRight Unit:=wdCharacter, Count:=2
      .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
      .Selection.Cells.Merge
      .Selection.Font.Bold = True
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.TypeText "Total"
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
      .Selection.Cells.Merge
      .Selection.Font.Bold = True
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Format(dblAFee + dblOFee, "#,##0")
      
      .Selection.MoveRight Unit:=wdCharacter, Count:=2
      .Selection.MoveDown Unit:=wdLine, Count:=1
      .Selection.InsertRows 6
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.TypeText ReportSum(71001)
      .Selection.MoveDown Unit:=wdLine, Count:=1
      .Selection.TypeText ReportSum(72)
      .Selection.MoveDown Unit:=wdLine, Count:=1
      .Selection.TypeText ReportSum(73001)
      .Selection.MoveDown Unit:=wdLine, Count:=1
      .Selection.TypeText ReportSum(85)
      .Selection.MoveDown Unit:=wdLine, Count:=1
      .Selection.TypeText ReportSum(74)
      .Selection.MoveDown Unit:=wdLine, Count:=1
      .Selection.TypeText ReportSum(121)
      
      .Selection.WholeStory
      .Selection.Font.Name = "Times New Roman"
      
      .Selection.HomeKey Unit:=wdStory
      
      PUB_GetLetterPicID "2", "FCP", iPicNo, iPicNo2, 2, True, Pub_StrUserSt03
      If PUB_ReadDB2File(stFileName, iPicNo) Then
         Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
         oShape.ZOrder 4
         oShape.LockAnchor = True
         oShape.LockAspectRatio = -1
         oShape.Width = .CentimetersToPoints(21)
         oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
         oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
         oShape.Left = .CentimetersToPoints(0)
         oShape.Top = .CentimetersToPoints(0)
         oShape.WrapFormat.Type = wdWrapNone
         .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
         .Selection.EndKey Unit:=wdStory
         .Selection.HomeKey Unit:=wdStory
         If iPicNo2 > 0 Then
            If PUB_ReadDB2File(stFileName, iPicNo2) Then
               For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
                  Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                  oShape.ZOrder 4
                  oShape.LockAnchor = True
                  oShape.LockAspectRatio = -1
                  oShape.Width = .CentimetersToPoints(21)
                  oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                  oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                  oShape.Left = .CentimetersToPoints(0)
                  oShape.Top = .CentimetersToPoints(27.6)
                  oShape.WrapFormat.Type = wdWrapNone
                  .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
                  .Selection.EndKey Unit:=wdStory
               Next ii
               .Selection.HomeKey Unit:=wdStory
            End If
         End If
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox "錯誤 : " & Err.Description, vbCritical
   End If
End Sub

'Added by Morgan 2022/5/17
'請款對象Y55666 NOVOCURE GMBH 特殊說明
Private Function GetNOVODesc(pDNo As String) As String
   Dim stSQL As String
   Dim intQ As Integer
   Dim rstQ As ADODB.Recordset
   Dim strA1L04s As String
   Dim strDesc As String
   
   stSQL = "select decode(substr(a1l04,-2),'99',substr(a1l04,1,length(a1l04)-2),a1l04) as Itm from acc1l0 where a1l01='" & pDNo & "'"
   intQ = 1
   Set rstQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      rstQ.MoveFirst
      strA1L04s = "," & rstQ.GetString(, , , ",")
      '中說請款: 101發明申請 or 102新型申請 or 103 設計申請
      If InStr(strA1L04s, ",101,") + InStr(strA1L04s, ",102,") + InStr(strA1L04s, ",103,") > 0 Then
         strDesc = "Filing Stage"
         
      '實審: 416實體審查
      ElseIf InStr(strA1L04s, ",416,") > 0 Then
         strDesc = "Substantive Examination Stage"
         
      '中間程序
      '主動修正: 203主動修正
      ElseIf InStr(strA1L04s, ",203,") > 0 Then
         strDesc = "Voluntary amendments"
      
      '申復  1202審查意見 or 205 申復
      ElseIf InStr(strA1L04s, ",1202,") + InStr(strA1L04s, ",205,") > 0 Then
         strDesc = "Office Action Stage"
      
      '再審  1002核駁通知 or 107 再審
      ElseIf InStr(strA1L04s, ",1002,") + InStr(strA1L04s, ",107,") > 0 Then
         strDesc = "Re-examination Stage"
         
      '分割  307分割
      ElseIf InStr(strA1L04s, ",307,") > 0 Then
         strDesc = "Divisional Application"
         
      '面詢  408面詢
      ElseIf InStr(strA1L04s, ",408,") > 0 Then
         strDesc = "Interview with examiner"
         
      '修正  204修正
      ElseIf InStr(strA1L04s, ",204,") > 0 Then
         strDesc = "Amendments"
      
      '核准
      '領証  601領證+第1年年費
      ElseIf InStr(strA1L04s, ",601,") > 0 Then
         strDesc = "Issue fee and 1st annuity fee"
         
      '二次核對 926核對已准專利
      ElseIf InStr(strA1L04s, ",926,") > 0 Then
         strDesc = "Final Review"
      '其他
      '讓與  701讓與 合併  702合併 繼承  703繼承 變更  401變更
      ElseIf InStr(strA1L04s, ",701,") + InStr(strA1L04s, ",702,") + InStr(strA1L04s, ",703,") + InStr(strA1L04s, ",401,") > 0 Then
         strDesc = "Recordal"
      End If
   End If
   GetNOVODesc = strDesc
   Set rstQ = Nothing
End Function

'Added by Morgan 2024/9/2
'複製/建立請款單pdf
Private Function CopyDN() As Boolean
   Dim rsQuery As ADODB.Recordset
   Dim strFileName As String, strSource As String, strDestination As String, strCaseNo As String
      
On Error GoTo ErrHnd
   'Modified by Morgan 2024/11/1 +a1k32
   strExc(0) = "select a1k01,a1k13,a1k14,a1k15,a1k16,st03,a1k32,a1k33 from acc1k0,staff where nvl(a1k12,0)=0 and a1k25||a1k29 is null and a1k02>=" & Text3 & "01 and a1k02<=" & Text3 & "31 and a1k28='" & m_strNo & "' and st01(+)=a1k21 order by a1k01 asc"
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   InsertQueryLog rsQuery.RecordCount
   If intI = 1 Then
      SetOutPutPath
      With rsQuery
      If .Fields("a1k33") = "3" Then
         If PdfSave3() = False Then GoTo ErrHnd
      ElseIf .Fields("a1k33") = "2" Then
         If PdfSave4() = False Then GoTo ErrHnd
      End If
      Do While Not .EOF
         strCaseNo = .Fields("a1k13") & .Fields("a1k14") & IIf(.Fields("a1k15") & .Fields("a1k16") <> "000", .Fields("a1k15") & .Fields("a1k16"), "")
         strFileName = strCaseNo & "_DN" & .Fields("a1k01") & ".pdf"
         'Memo by Morgan 2024/11/1 若請款單建立/修改人員換部門時會抓不到檔案(路徑不對)而跑列印 Ex:B3033
         strSource = PUB_GetEFilePath(.Fields("a1k13"), "" & .Fields("st03")) & "\" & .Fields("a1k13") & "\" & Left(.Fields("a1k14"), 3) & "\" & strCaseNo & "\" & strFileName
         strDestination = m_strSavePath & "\" & strFileName
         If Dir(strSource) <> "" Then
            FileCopy strSource, strDestination
         'Added by Morgan 2024/11/1
         ElseIf Not IsNull(rsQuery("a1k32")) Then
            MsgBox rsQuery("a1k01") & " 為 " & IIf(rsQuery("a1k32") = "Y", "特殊", "") & IIf(rsQuery("a1k32") = "C", "整批", "") & " 請款單，請人工處理！", vbExclamation
         'end 2024/11/1
         Else
            '請款單電子檔
            Load Frmacc2480
            With Frmacc2480
               .Text1.Text = rsQuery("a1k01")
               .Text2.Text = .Text1.Text
               .txtOutMode = "2"
               .m_bBeCalled = True
               .m_CallPrevForm = Me.Name
               .m_bEMail = True
               .m_SavePath = m_strSavePath
               .Command2_Click
            End With
            Unload Frmacc2480
            strFormName = Me.Name
            tool3_enabled
         End If
         rsQuery.MoveNext
      Loop
      End With
      CopyDN = True
   Else
      MsgBox "無符合資料！"
   End If
   
ErrHnd:
   Set rsQuery = Nothing
End Function

'Added by Morgan 2024/9/3
'複製/建立LEDES帳單電子檔
Private Function CopyLEDES() As Boolean
   Dim rsQuery As ADODB.Recordset
   Dim strFileName As String, strSource As String, strDestination As String, strCaseNo As String
   Dim strError As String
      
On Error GoTo ErrHnd

   strExc(0) = "select a1k01,a1k13,a1k14,a1k15,a1k16,a1k33,st03 from acc1k0,staff where nvl(a1k12,0)=0 and a1k25 is null and a1k02>=" & Text3 & "01 and a1k02<=" & Text3 & "31 and a1k28='" & m_strNo & "' and st01(+)=a1k21 order by a1k01 asc"
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   InsertQueryLog RsTemp.RecordCount
   If intI = 1 Then
      SetOutPutPath
      With rsQuery
      If .Fields("a1k33") = "3" Then
         If PdfSave3(True) = False Then GoTo ErrHnd
      ElseIf .Fields("a1k33") = "2" Then
         If PdfSave4(True) = False Then GoTo ErrHnd
      End If
      Do While Not .EOF
         strCaseNo = .Fields("a1k13") & .Fields("a1k14") & IIf(.Fields("a1k15") & .Fields("a1k16") <> "000", .Fields("a1k15") & .Fields("a1k16"), "")
         
         strFileName = strCaseNo & "_DN" & .Fields("a1k01") & ".txt"
         strSource = PUB_GetEFilePath(.Fields("a1k13"), "" & .Fields("st03")) & "\" & .Fields("a1k13") & "\" & Left(.Fields("a1k14"), 3) & "\" & strCaseNo & "\" & strFileName
         strDestination = m_strSavePath & "\" & strFileName
         If Dir(strSource) <> "" Then
            FileCopy strSource, strDestination
         Else
            'LEDES帳單電子檔
            Load Frmacc2480
            With Frmacc2480
               .Visible = False
               .Text1.Text = rsQuery("a1k01")
               .Text2.Text = .Text1.Text
               .Check2.Value = vbChecked
               .m_bBeCalled = True
               .m_CallPrevForm = Me.Name
               .m_SavePath = m_strSavePath
               .Command2.Value = True
               strError = .m_sEBillingMsg
            End With
            Unload Frmacc2480
            strFormName = Me.Name
            tool3_enabled
            
            If strError <> "" Then
               MsgBox strError
               GoTo ErrHnd
            End If
            
         End If
         rsQuery.MoveNext
      Loop
      End With
      CopyLEDES = True
   Else
      MsgBox "無符合資料！"
   End If
   
ErrHnd:
   Set rsQuery = Nothing
End Function


'Added by Morgan 2025/8/21
'英文請款項目敘述(此為簡化版，完整需參看frmacc2480)
Private Function GetDNItemDesc(pDNo As String, pItemNo As String, Optional pAmount As String) As String
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stDesc As String
   
   stSQL = "select a1k13,decode(a2607,null,rtrim(a1j04||' '||a1j05||' '||a1j06),rtrim(a2607||' '||a2608||' '||a2609)) a1j04" & _
      " from acc1k0,acc1j0,acc260 where a1k01='" & pDNo & "'" & _
      " and a1j01(+)=a1k13 and a1j02(+)='" & pItemNo & "' and a2601(+)=substr(a1k27,1,8) and a2602(+)=a1k13 and a2603(+)=a1j02"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      stDesc = "" & rsQuery("a1j04")
      If (rsQuery("a1k13") = "P" Or rsQuery("a1k13") = "FCP" Or rsQuery("a1k13") = "CFP") And (pItemNo = "601" Or pItemNo = "605") Then
         stDesc = PUB_GetAnnuityDesc(pDNo, pItemNo, stDesc)
      End If
      If pAmount > 0 Then
         stDesc = PUB_ParseItemDesc(pAmount, stDesc)
      End If
   End If
   GetDNItemDesc = stDesc
   Set rsQuery = Nothing
End Function

Private Function PdfSave6(Optional pXlsOnly As Boolean = False) As Boolean
   Const cFontSize = 12
   Dim oTable As Word.Table
   Dim oShape As Word.Shape
   Dim dblAFee As Double, dblOFee As Double, dblDFee As Double, dblTFee As Double, dblTFeeNT As Double
   Dim iRow As Integer, iSNo As Integer
   Dim stFileName As String
   Dim stPdfName As String, stFullPath As String
   Dim rsReprot As ADODB.Recordset
   Dim bolInvDate As Boolean, iCol As Integer, iCols As Integer 'Added by Morgan 2017/5/10
   Dim stAddrNo As String 'Added by Morgan 2021/2/22 列印對象
   Dim oWordAp As Word.Application
   Dim stCon0K0 As String
   Dim stDNCurr As String 'Added by Morgan 2022/7/1
   Dim stTitle As String 'Added by Morgan 2022/11/23
   Dim strInvNo As String 'Added by Morgan 2024/1/24
   
On Error GoTo ErrHnd

   stAddrNo = m_strNo
   
   bolInvDate = True
   
   '表頭
   If Left(stAddrNo, 1) = "Y" Then
      strExc(0) = "select fa05,fa63,fa64,fa65,fa32,fa33,fa34,fa35,fa36,fa18,fa19,fa20,fa21,fa22,fa70,fa17,fa23" & _
         " from fagent where fa01='" & Left(stAddrNo, 8) & "' and fa02='" & Mid(stAddrNo, 9) & "'"
   Else
      strExc(0) = "select cu05 as fa05,cu88 as fa63,cu89 as fa64,cu90 as fa65,cu65 as fa32, cu66 as fa33, cu67 as fa34" & _
         ", cu68 as fa35, cu69 as fa36, cu24 as fa18, cu25 as fa19, cu26 as fa20, cu27 as fa21, cu28 as fa22,cu102 fa70" & _
         ", cu23 as fa17, cu29 as fa23 from customer where cu01='" & Left(stAddrNo, 8) & "' and cu02='" & Mid(stAddrNo, 9) & "'"
   End If
   intI = 1
   Set rsReprot = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      MsgBox "表頭資料讀取失敗!!"
      Exit Function
   End If
   
   With rsReprot
   '代理人名稱 strexc(1)
   strExc(1) = "" & .Fields("fa05")
   If Not IsNull(.Fields("fa63")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa63")
   End If
   If Not IsNull(.Fields("fa64")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa64")
   End If
   If Not IsNull(.Fields("fa65")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa65")
   End If
   '代理人POBox/地址
   If Not IsNull(.Fields("fa32")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa32")
      If Not IsNull(.Fields("fa33")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa33")
      End If
      If Not IsNull(.Fields("fa34")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa34")
      End If
      If Not IsNull(.Fields("fa35")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa35")
      End If
      If Not IsNull(.Fields("fa36")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa36")
      End If
   ElseIf Not IsNull(.Fields("fa18")) Then
      strExc(1) = strExc(1) & vbCrLf & .Fields("fa18")
      If Not IsNull(.Fields("fa19")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa19")
      End If
      If Not IsNull(.Fields("fa20")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa20")
      End If
      If Not IsNull(.Fields("fa21")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa21")
      End If
      If Not IsNull(.Fields("fa22")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa22")
      End If
      If Not IsNull(.Fields("fa70")) Then
         strExc(1) = strExc(1) & vbCrLf & .Fields("fa70")
      End If
   End If
   End With
   
   'Added by Morgan 2022/3/17 Y55666000 若有財務編號也要印 --Ryan
   strExc(2) = PUB_GetACCNO(m_strNo)
   'end 2022/3/17
      
   stCon0K0 = " and a1k02>=" & Text3 & "01 and a1k02<=" & Text4 & "31 and a1k28='" & m_strNo & "'"
   
   '彼所案號
   '請款對象Y55666000 NOVOCURE GMBH 彼號欄位 (Your ref:) 優先抓客戶案件案號--Franny
   '更代後客戶案號會改放到彼號 Ex:X11408163--Franny
   strExc(0) = "select a1k01,a1k13,to_char(to_date(a1k02+19110000,'yyyymmdd'),'FMMM/DD/yyyy') dt" & _
      ",nvl(tm35||pa48||lc17||sp29,tm45||pa77||lc23||sp27) YrRef,tm12||pa11||sp11 AppNo,a1k13||'-'||a1k14||decode(a1k15||a1k16,'000','','-'||a1k16||'-'||a1k17) OrRef,a1k08 TFee" & _
      ",a1k11,a1k18,X.*,rtrim(decode(a2607,null,X004,a2607||' '||a2608||' '||a2609)) IDesc,to_char(to_date(a1k02+19110000,'yyyymmdd'),'YYYY/MM/DD') dt2" & _
      " from (select a1l01,min(a1k28) X001,min(a1l03) X002,substr(min(a1l02||a1l04),4) X003,substr(min(a1l02||a1j04),4) X004" & _
      " from (select a.a1l01,a.a1l02,a.a1l03,a.a1l04,a1j03" & _
      ",a1k28,rtrim(a1j04||' '||a1j05||' '||a1j06) a1j04" & _
      " from acc1k0,acc1l0 a,acc1l0 b,acc1j0" & _
      " where nvl(a1k12,0)=0 and a1k25 is null " & stCon0K0 & _
      " and a.a1l01(+)=a1k01 and substr(a.a1l04(+),-2)<>'98' and b.a1l01(+)=a.a1l01 and b.a1l03(+)=a.a1l03 and b.a1l04(+)=a.a1l04||'98'" & _
      " and a1j01(+)=a.a1l03 and a1j02(+)=a.a1l04" & _
      ") group by a1l01) X,acc1k0,trademark,patent,lawcase,servicepractice,acc260 where a1k01(+)=a1l01" & _
      " and tm01(+)=a1k13 and tm02(+)=a1k14 and tm03(+)=a1k15 and tm04(+)=a1k16" & _
      " and pa01(+)=a1k13 and pa02(+)=a1k14 and pa03(+)=a1k15 and pa04(+)=a1k16" & _
      " and sp01(+)=a1k13 and sp02(+)=a1k14 and sp03(+)=a1k15 and sp04(+)=a1k16" & _
      " and lc01(+)=a1k13 and lc02(+)=a1k14 and lc03(+)=a1k15 and lc04(+)=a1k16" & _
      " and a2601(+)=substr(X001,1,8) and a2602(+)=X002 and a2603(+)=X003" & _
      " order by a1k02,a1k01"
   intI = 1
   Set rsReprot = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      MsgBox "請款明細資料讀取失敗!!"
      Exit Function
   End If
   
   ExcelSave6 rsReprot, strExc(1), strExc(2)
   rsReprot.MoveFirst
   
   Set oWordAp = New Word.Application
   oWordAp.Visible = True
   oWordAp.Documents.add
   With oWordAp
      .Selection.Font.Name = "Times New Roman"
      .Selection.Font.Size = cFontSize
      
      '版面設定
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.5)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
      .Selection.PageSetup.FooterDistance = .CentimetersToPoints(3)
      .Selection.PageSetup.CharsLine = 40
      .Selection.PageSetup.LinesPage = 38
      .Selection.Orientation = wdTextOrientationHorizontal
      
      '信頭尾
      If PUB_ReadDB2File(stFileName, iPicNo) = True Then
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
         Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
         oShape.ZOrder 4
         oShape.LockAnchor = True
         oShape.LockAspectRatio = -1
         oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
         oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
         oShape.Left = 0
         oShape.Top = 0
         oShape.Width = .CentimetersToPoints(21)
         oShape.WrapFormat.Type = wdWrapNone
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         
         If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
            .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
            oShape.ZOrder 4
            oShape.LockAnchor = True
            oShape.LockAspectRatio = -1
            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
            oShape.Left = 0
            oShape.Top = .CentimetersToPoints(27)
            oShape.Width = .CentimetersToPoints(21)
            oShape.WrapFormat.Type = wdWrapNone
            .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         End If
         
         .Selection.HomeKey Unit:=wdStory
      End If
      
      .Selection.TypeParagraph
      '行距
      With .Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .DisableLineHeightGrid = True
      End With
      
      '新增表格(1*2)
      Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumRows:=1, NumColumns:=2)
      With oTable
         '無邊框
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
      End With
            
      oTable.Select
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop '靠上對齊
      .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
      .Selection.InsertRows 8

      '代理人名稱,POBox/地址
      oTable.Cell(1, 1).Merge oTable.Cell(4, 1)
      oTable.Cell(1, 1).Select
      .Selection.Text = strExc(1)
      
      '月份
      strExc(0) = Format(ChangeTStringToWDateString(Text3 & "01"), "mmmm, yyyy")
      If Text4 <> Text3 Then
         strExc(0) = strExc(0) & " - " & Format(ChangeTStringToWDateString(Text4 & "01"), "mmmm, yyyy")
      End If
      
      stTitle = "Monthly Invoice"
      strInvNo = rsReprot("a1k01") & "/" & Text3
      stTitle = stTitle & " No. " & strInvNo & " (for " & strExc(0) & ")"
      
      strExc(0) = "Date: " & Format(ChangeTStringToWDateString(strSrvDate(2)), "mmmm dd, yyyy")
      oTable.Cell(2, 2).Select
      .Selection.Text = strExc(0)
            
      iRow = 5
      
      '財務編號
      oTable.Cell(iRow, 1).Select
      .Selection.Text = strExc(2)
      iRow = iRow + 1
      oTable.Cell(iRow, 1).Select
      .Selection.InsertRows 1
      
      oTable.Cell(iRow, 1).Merge oTable.Cell(iRow, 2)
      oTable.Cell(iRow, 1).Select
      .Selection.Cells(1).SetHeight RowHeight:=30, HeightRule:=wdRowHeightAtLeast
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      .Selection.Text = stTitle
      
      iRow = iRow + 1
      
      oTable.Cell(iRow, 1).Select
      .Selection.SelectRow

      With .Selection.Cells
        '有邊框
        .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
        .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
      End With
      
      
      iCols = 7
      
      .Selection.Cells.Split NumRows:=1, NumColumns:=iCols, MergeBeforeSplit:=True
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      .Selection.Font.Size = 10
      
      '設定表格高度欄寬
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.7), RulerStyle:=wdAdjustProportional
      .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
      .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.7), RulerStyle:=wdAdjustProportional
      .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.7), RulerStyle:=wdAdjustProportional
      .Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
      .Selection.Cells(6).SetWidth ColumnWidth:=.CentimetersToPoints(5), RulerStyle:=wdAdjustProportional
      
      .Selection.Cells(1).SetHeight RowHeight:=36, HeightRule:=wdRowHeightAtLeast
      .Selection.InsertRows rsReprot.RecordCount + 1
      
      oTable.Cell(iRow, 1).Select
      .Selection.SelectRow
      .Selection.Font.Bold = True
      oTable.Cell(iRow, 1).Select
      .Selection.Text = "No."
      
      oTable.Cell(iRow, 2).Select
      .Selection.Text = "Invoice Date"
      oTable.Cell(iRow, 3).Select
      .Selection.Text = "Application" & vbCrLf & "Number"
      oTable.Cell(iRow, 4).Select
      .Selection.Text = "Novocure" & vbCrLf & "Matter"
      oTable.Cell(iRow, 5).Select
      .Selection.Text = "Our Ref"
      oTable.Cell(iRow, 6).Select
      .Selection.Text = "Task"
      oTable.Cell(iRow, 7).Select
      .Selection.Text = "Charge" & vbCrLf & "(USD)"
      
      
      .Selection.SelectRow
      .Selection.Cells.Shading.Texture = wdTexture15Percent
      .Selection.Cells(1).SetHeight RowHeight:=36, HeightRule:=wdRowHeightAtLeast
      iSNo = 0
      Do While Not rsReprot.EOF
         iRow = iRow + 1
         iSNo = iSNo + 1
         oTable.Cell(iRow, 1).Select
         .Selection.Text = iSNo
         
         iCol = 2
         oTable.Cell(iRow, iCol).Select
         .Selection.Text = "" & rsReprot("dt")
         
         iCol = iCol + 1
         oTable.Cell(iRow, iCol).Select
         .Selection.Text = "" & rsReprot("AppNo")
         
         iCol = iCol + 1
         oTable.Cell(iRow, iCol).Select
         .Selection.Text = "" & rsReprot("YrRef")
         
         iCol = iCol + 1
         oTable.Cell(iRow, iCol).Select
         .Selection.Text = "" & rsReprot("OrRef")
         
         iCol = iCol + 1
         oTable.Cell(iRow, iCol).Select
         .Selection.Text = "" & GetNOVODesc("" & rsReprot("a1k01"))
         
         iCol = iCol + 1
         oTable.Cell(iRow, iCol).Select
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Text = Format(Val("" & rsReprot("TFee")), cfmtDollar)
         dblTFee = dblTFee + Val("" & rsReprot("TFee"))
         dblTFeeNT = dblTFeeNT + Val("" & rsReprot("a1k11"))
         rsReprot.MoveNext
      Loop
      
      iRow = iRow + 1
      oTable.Cell(iRow, 1).Merge oTable.Cell(iRow, iCols - 1)
      oTable.Cell(iRow, 1).Select
      .Selection.SelectRow
      .Selection.Font.Bold = True
      
      oTable.Cell(iRow, 1).Select
      .Selection.Text = "Total"
      oTable.Cell(iRow, 2).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.Text = Format(dblTFee, cfmtDollar)
      
      '帳號
      iRow = iRow + 2
      oTable.Cell(iRow, 1).Select
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
      .Selection.Text = vbCrLf & ReportSum(71001) & vbCrLf & ReportSum(72) & vbCrLf & ReportSum(73001) & vbCrLf & ReportSum(85) & vbCrLf & ReportSum(74) & vbCrLf & ReportSum(121) & vbCrLf
      
      '建議電匯提醒
      oTable.Cell(iRow, 2).Select
      .Selection.Cells.Split NumRows:=3, NumColumns:=1, MergeBeforeSplit:=False
      .Selection.Cells(1).SetHeight RowHeight:=28, HeightRule:=wdRowHeightAtLeast
      iRow = iRow + 1
      oTable.Cell(iRow, 2).Select
      .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
      With .Selection.Cells(1)
          With .Borders(wdBorderLeft)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderRight)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderTop)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderBottom)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
      End With
      .Selection.ParagraphFormat.LeftIndent = .CentimetersToPoints(0.2)
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
      .Selection.Cells(1).SetHeight RowHeight:=52, HeightRule:=wdRowHeightAtLeast
      .Selection.Text = "Wire" & vbCrLf & "Transfer" & vbCrLf & "Preferred"
      iRow = iRow + 1
      oTable.Cell(iRow, 2).Select
      .Selection.Cells(1).SetHeight RowHeight:=0, HeightRule:=wdRowHeightAtLeast
      
      '備註
      iRow = iRow + 1
      oTable.Cell(iRow, 1).Select
      .Selection.SelectRow
      .Selection.Font.Bold = True
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
      oTable.Cell(iRow, 1).Select
      .Selection.Text = "PS:"
      oTable.Cell(iRow, 2).Select
      .Selection.Text = "Please return a copy of the invoice(s) or indicate the invoice number(s) paid with remittance"
      .Selection.EndKey
      
      rsReprot.MoveFirst
      Do While Not rsReprot.EOF
         '請款單電子檔
         Load Frmacc2480
         With Frmacc2480
            .Text1.Text = rsReprot("a1k01")
            .Text2.Text = .Text1.Text
            .m_bEditDoc = True
            .m_bBeCalled = True
            .m_CallPrevForm = Me.Name
            .m_bolNoPic = True
            .m_strInvoiceNo = strInvNo
            .Command2_Click
         End With
         Unload Frmacc2480
         strFormName = Me.Name
         tool3_enabled
         
         oWordAp.Selection.EndKey Unit:=wdStory
         oWordAp.Selection.TypeText Chr(12)
         g_WordAp.Selection.WholeStory
         g_WordAp.Selection.Copy
         oWordAp.Selection.Paste
         g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
         g_WordAp.Quit wdDoNotSaveChanges
         
         rsReprot.MoveNext
      Loop
      
      If Text2 = "Y" Then
         .Activate
      Else
         stPdfName = m_strNo & Text3 & ".pdf"
         .ActiveDocument.ExportAsFixedFormat OutputFileName:=m_strSavePath & "\" & stPdfName, ExportFormat:=17, OpenAfterExport:=False
         .ActiveDocument.Close wdDoNotSaveChanges
         .Quit wdDoNotSaveChanges
      End If
   End With
   
XlsOnly:

   PdfSave6 = True
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   Set rsReprot = Nothing
   Set oWordAp = Nothing
End Function

'Added by Morgan 2025/11/4
Private Function ExcelSave6(pRst As ADODB.Recordset, pAddr As String, pAccNo As String) As Boolean
   Dim xlsReport As New Excel.Application
   Dim wksReport As New Worksheet
   Dim stFullPath As String
   Dim ii As Integer, dblTFee As Double
   Dim bolInvDate As Boolean, stCol As String, iCols As Integer, jj As Integer
   
On Error GoTo ErrHnd
   
   xlsReport.Visible = True
   With pRst
   .MoveFirst
   xlsReport.SheetsInNewWorkbook = 1
   xlsReport.Workbooks.add
   Set wksReport = xlsReport.Worksheets(1)
   
   wksReport.Cells.NumberFormatLocal = "@"
   wksReport.Cells.Font.Name = "Times New Roman"
   ii = 1
   wksReport.Range("A" & ii, "F" & ii).Merge
   wksReport.Range("A" & ii) = pAddr
   
   '自動調整高度
   jj = 0
   For intI = 1 To Len(pAddr)
      If Mid(pAddr, intI, 2) = vbCrLf Then
         jj = jj + 1
         intI = intI + 1
      End If
   Next
   wksReport.Rows(ii).EntireRow.AutoFit
   wksReport.Rows(ii).RowHeight = wksReport.Rows(ii).Height * (jj + 1)
   
   ii = ii + 1
   wksReport.Range("A" & ii) = pAccNo
   wksReport.Range("A" & ii, "F" & ii).Merge
   ii = ii + 1
   stCol = "A"
   wksReport.Range(stCol & ii) = "Invoice Date"
   stCol = Chr(Asc(stCol) + 1)
   wksReport.Range(stCol & ii) = "Application Number"
   stCol = Chr(Asc(stCol) + 1)
   wksReport.Range(stCol & ii) = "Novocure Matter"
   stCol = Chr(Asc(stCol) + 1)
   wksReport.Range(stCol & ii) = "Our Ref"
   stCol = Chr(Asc(stCol) + 1)
   wksReport.Range(stCol & ii) = "Task"
   stCol = Chr(Asc(stCol) + 1)
   wksReport.Range(stCol & ii) = "Charge (USD)"
   wksReport.Range("A" & ii, stCol & ii).Font.Bold = True
   
   iCols = Asc(stCol) - Asc("A") + 1
   Do While Not .EOF
      ii = ii + 1
      stCol = "A"
      wksReport.Range(stCol & ii) = "" & .Fields("dt")
      stCol = Chr(Asc(stCol) + 1)
      wksReport.Range(stCol & ii) = "" & .Fields("AppNo")
      stCol = Chr(Asc(stCol) + 1)
      wksReport.Range(stCol & ii) = "" & .Fields("YrRef")
      stCol = Chr(Asc(stCol) + 1)
      wksReport.Range(stCol & ii) = "" & .Fields("OrRef")
      stCol = Chr(Asc(stCol) + 1)
      wksReport.Range(stCol & ii) = GetNOVODesc("" & .Fields("a1k01"))
      stCol = Chr(Asc(stCol) + 1)
      wksReport.Range(stCol & ii) = Val("" & .Fields("TFee"))
      wksReport.Range(stCol & ii).NumberFormatLocal = "#,##0.00_ "
      dblTFee = dblTFee + Val("" & .Fields("TFee"))
      .MoveNext
   Loop
   End With
      
   For jj = 0 To iCols - 1
      stCol = Chr(Asc("A") + jj)
      wksReport.Columns(stCol & ":" & stCol).EntireColumn.AutoFit
      If wksReport.Range(stCol & "1").ColumnWidth > 80 Then
         wksReport.Range(stCol & "1").ColumnWidth = 80
      End If
   Next
    
'   ii = ii + 1
'   stCol = Chr(Asc("A") + iCols - 2)
'   wksReport.Range("A" & ii) = "Total"
'   wksReport.Range("A" & ii, stCol & ii).Merge
'   wksReport.Range("A" & ii, stCol & ii).HorizontalAlignment = xlCenter
'
'   stCol = Chr(Asc(stCol) + 1)
'   wksReport.Range(stCol & ii) = dblTFee
'   wksReport.Range(stCol & ii).NumberFormatLocal = "#,##0.00_ "
'   wksReport.Range("A" & ii, stCol & ii).Font.Bold = True

   xlsReport.Range("A1").Select
      
   stFullPath = m_strSavePath & "\" & m_FileName
   If Dir(stFullPath & ".*") <> "" Then
      Kill stFullPath & ".*"
   End If
   xlsReport.Workbooks(1).SaveAs stFullPath
   xlsReport.Workbooks.Close
   xlsReport.Quit
   
   ExcelSave6 = True
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
ExitPoint:
   Set xlsReport = Nothing
   
End Function
