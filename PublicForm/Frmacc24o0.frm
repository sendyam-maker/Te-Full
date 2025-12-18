VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc24o0 
   AutoRedraw      =   -1  'True
   Caption         =   "請款單折扣案件明細"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   6075
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   630
      TabIndex        =   30
      Top             =   2640
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.CommandButton CmdPrt1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   720
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   3840
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
      _Version        =   393216
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   360
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
      _Version        =   393216
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   16
      Left            =   2760
      TabIndex        =   29
      Top             =   1952
      Width           =   255
      VariousPropertyBits=   8388627
      Caption         =   "~"
      Size            =   "450;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   15
      Left            =   2760
      TabIndex        =   28
      Top             =   1541
      Width           =   255
      VariousPropertyBits=   8388627
      Caption         =   "~"
      Size            =   "450;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   14
      Left            =   2760
      TabIndex        =   27
      Top             =   1130
      Width           =   255
      VariousPropertyBits=   8388627
      Caption         =   "~"
      Size            =   "450;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   13
      Left            =   2760
      TabIndex        =   26
      Top             =   308
      Width           =   255
      VariousPropertyBits=   8388627
      Caption         =   "~"
      Size            =   "450;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   12
      Left            =   4560
      TabIndex        =   25
      Top             =   719
      Width           =   1455
      VariousPropertyBits=   8388627
      Caption         =   "(ALL: 全部)"
      Size            =   "2566;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label lblSName 
      Height          =   225
      Left            =   2520
      TabIndex        =   24
      Top             =   2774
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "lblSName"
      Size            =   "2143;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   585
      Index           =   10
      Left            =   4680
      TabIndex        =   23
      Top             =   1440
      Width           =   1335
      VariousPropertyBits=   8388627
      Caption         =   "外商：F10~F19  外專：F20~F29  外法：F30~F49"
      Size            =   "2355;1032"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   9
      Left            =   4440
      TabIndex        =   22
      Top             =   1125
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "業務區說明："
      Size            =   "2143;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   8
      Left            =   4560
      TabIndex        =   21
      Top             =   2325
      Width           =   1455
      VariousPropertyBits=   8388627
      Caption         =   "(以逗點區隔)"
      Size            =   "2566;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   9
      Left            =   1440
      TabIndex        =   11
      Top             =   3120
      Width           =   495
      VariousPropertyBits=   679495707
      MaxLength       =   1
      Size            =   "873;635"
      Value           =   "2"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   7
      Left            =   120
      TabIndex        =   20
      Top             =   3188
      Width           =   4215
      VariousPropertyBits=   8388627
      Caption         =   "輸出方式：　　　(1.螢幕 2.印表機)"
      Size            =   "7435;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   2789
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "智權人員："
      Size            =   "2143;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   2363
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "案件性質："
      Size            =   "2143;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   1952
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "請款對象："
      Size            =   "2143;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   1541
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "國　　籍："
      Size            =   "2143;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   330
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   1100
      Width           =   1290
      VariousPropertyBits=   8388627
      Caption         =   "業 務 區："
      Size            =   "2275;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   704
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "系統類別："
      Size            =   "2143;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   8
      Left            =   1440
      TabIndex        =   10
      Top             =   2706
      Width           =   1000
      VariousPropertyBits=   679495707
      MaxLength       =   6
      Size            =   "1764;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   7
      Left            =   1440
      TabIndex        =   9
      Top             =   2295
      Width           =   3015
      VariousPropertyBits=   679495707
      Size            =   "5318;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   6
      Left            =   3120
      TabIndex        =   8
      Top             =   1884
      Width           =   1200
      VariousPropertyBits=   679495707
      MaxLength       =   9
      Size            =   "2117;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   5
      Left            =   1440
      TabIndex        =   7
      Top             =   1884
      Width           =   1200
      VariousPropertyBits=   679495707
      MaxLength       =   9
      Size            =   "2117;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   4
      Left            =   3120
      TabIndex        =   6
      Top             =   1473
      Width           =   1200
      VariousPropertyBits=   679495707
      MaxLength       =   3
      Size            =   "2117;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   3
      Left            =   1440
      TabIndex        =   5
      Top             =   1473
      Width           =   1200
      VariousPropertyBits=   679495707
      MaxLength       =   3
      Size            =   "2117;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   2
      Left            =   3120
      TabIndex        =   4
      Top             =   1062
      Width           =   1200
      VariousPropertyBits=   679495707
      MaxLength       =   3
      Size            =   "2117;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   1062
      Width           =   1200
      VariousPropertyBits=   679495707
      MaxLength       =   3
      Size            =   "2117;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   651
      Width           =   3015
      VariousPropertyBits=   679495707
      Size            =   "5318;635"
      Value           =   "ALL"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   308
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "請款日期："
      Size            =   "2143;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
End
Attribute VB_Name = "Frmacc24o0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2018/11/22 請款單折扣案件明細
'Memo by Lydia 2018/11/22 使用Form 2.0 (Label和TextBox)
Option Explicit

Dim RsAcc24o0 As New ADODB.Recordset

Dim oText As MSForms.TextBox
Dim ii As Integer

Dim PLeft() As Integer, PColName() As String, PColWidth() As String, strTemp() As String
Dim iPrint As Integer, iPage As Integer
Dim m_iStartX As Integer, m_iStartY As Integer, m_iGap As Integer
Dim m_iPageHeight As Integer, m_iLineHeight As Integer, m_iMargin As Integer
Dim m_bPrinter As Boolean, m_iPages As Integer, m_Device
Dim mStrGrp(1 To 5) As String '請款對象、小計(折扣金額及未折扣總額)、合計

Private Sub CmdPrt1_Click()
Dim tmpArr1 As Variant
Dim stCon As String
Dim stConCP As String
Dim intP As Integer

   If FormCheck = False Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   
   ClearQueryLog (Me.Name) '清除查詢印表記錄檔欄位
   
   '請款日期
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      stCon = stCon & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      stCon = stCon & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If (MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29)) Or _
      (MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29)) Then
      pub_QL05 = pub_QL05 & ";" & LblFM2(0) & MaskEdBox1 & "-" & MaskEdBox2
   End If
   '系統類別
   tmpArr1 = Empty
   If txtFM2(0) <> "ALL" Then
      tmpArr1 = Split(txtFM2(0).Text, ",")
      strExc(1) = ""
      For ii = LBound(tmpArr1) To UBound(tmpArr1)
          If Trim(tmpArr1(ii)) <> "" Then
               strExc(1) = strExc(1) & "'" & tmpArr1(ii) & "',"
          End If
      Next ii
      strExc(1) = Left(strExc(1), Len(strExc(1)) - 1)
      stCon = stCon & " AND A1K13 IN ( " & GetAddStr(txtFM2(0)) & " ) "
   Else
      stCon = stCon & " AND A1K13 IN ( " & GetAddStr(Systemkind_g) & " ) "
   End If
   If Trim(txtFM2(0)) <> "" Then
      pub_QL05 = pub_QL05 & ";" & LblFM2(1) & txtFM2(0)
   End If
   
   '業務區
   If txtFM2(1) <> "" Then
      stConCP = " and CP12 >= '" & txtFM2(1) & "'"
   End If
   If txtFM2(2) <> "" Then
      stConCP = stConCP & " and CP12 <= '" & txtFM2(2) & "'"
   End If
   If txtFM2(1) <> "" Or txtFM2(2) <> "" Then
      pub_QL05 = pub_QL05 & ";" & LblFM2(2) & txtFM2(1) & "-" & txtFM2(2)
   End If
   
   '智權人員
   If txtFM2(8) <> "" Then
      stConCP = stConCP & " AND CP13= '" & txtFM2(8) & "'"
      pub_QL05 = pub_QL05 & ";" & LblFM2(6) & txtFM2(8)
   End If
   
   '請款對象-國籍
   If txtFM2(3) <> MsgText(601) Then
      stCon = stCon & " AND FA10 >= '" & txtFM2(3) & "'"
   End If
   If txtFM2(4) <> MsgText(601) Then
      stCon = stCon & " AND FA10 <= '" & txtFM2(4) & "z'"
   End If
   If txtFM2(3) <> MsgText(601) Or txtFM2(4) <> MsgText(601) Then
      pub_QL05 = pub_QL05 & ";" & LblFM2(3) & txtFM2(3) & "-" & txtFM2(4)
   End If
   '請款對象
   If txtFM2(5) <> "" Then
      stCon = stCon & " AND A1K28 >= '" & txtFM2(5) & "'"
   End If
   If txtFM2(6) <> "" Then
      stCon = stCon & " AND A1K28 <= '" & txtFM2(6) & "'"
   End If
   If txtFM2(5) <> "" Or txtFM2(6) <> "" Then
      pub_QL05 = pub_QL05 & ";" & LblFM2(4) & txtFM2(5) & "-" & txtFM2(6)
   End If

   '案件性質
   If Trim(txtFM2(7)) <> "" Then
      tmpArr1 = Split(txtFM2(7).Text, ",")
      strExc(1) = ""
      For ii = LBound(tmpArr1) To UBound(tmpArr1)
          If Trim(tmpArr1(ii)) <> "" Then
               strExc(1) = strExc(1) & " OR A1L04 LIKE " & CNULL(tmpArr1(ii) & "%")
          End If
      Next ii
      strExc(1) = Mid(strExc(1), 4)
      stCon = stCon & " AND (" & strExc(1) & ")"
      pub_QL05 = pub_QL05 & ";" & LblFM2(5) & txtFM2(7)
   End If

   If txtFM2(9) = "1" Then
      pub_QL05 = pub_QL05 & ";" & LblFM2(7) & "1.螢幕" '
   Else
      pub_QL05 = pub_QL05 & ";" & LblFM2(7) & "2.印表機"
   End If
   
   strSql = " SELECT A1K02,SUBSTR(NA01,1,3) NA01,NA03,A1K28,DECODE(FA31,'1',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65))," & _
               " '2',DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65),NVL(FA06,NVL(FA04,FA05))) A1K28N," & _
               " A1K13||'-'||A1K14||DECODE(A1K15||A1K16,'000','',A1K15||'-'||A1K16) CASENO,A1J03,A1L01,A1L02,A1L03,A1L04,A1L05,A1L07,A1L19,A1K06" & _
               " FROM ACC1K0 A,ACC1L0 B,ACC1J0,FAGENT,NATION" & _
               " WHERE A1K01=A1L01 AND A1K12 IS NULL AND A1K25 IS NULL AND A1L07>0 AND A1L19> 0 AND A1L03=A1J01(+) AND A1L04=A1J02(+)" & _
               " AND SUBSTR(A1K28,1,8)=FA01(+) AND SUBSTR(A1K28,9,1)=FA02(+) AND FA10=NA01(+)" & stCon & _
               IIf(stConCP <> "", " AND A1L01 IN (SELECT CP60 FROM CASEPROGRESS WHERE CP01=A.A1K13 AND CP02=A.A1K14 AND CP03=A.A1K15 AND CP04=A.A1K16" & stConCP & ")", "")
   '以國籍、請款對象(no+name)、請款日期排序
   strSql = strSql & " ORDER BY 2 ASC,A1K28,A1K02" & IIf(InStr(stCon, "A1L04") > 0, ",A1L04,A1L01,A1L02", ",A1L01,A1L02")
   If RsAcc24o0.State = adStateOpen Then
      RsAcc24o0.Close
   End If
   RsAcc24o0.CursorLocation = adUseClient

   RsAcc24o0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If RsAcc24o0.RecordCount <> 0 Then
      InsertQueryLog (RsAcc24o0.RecordCount)
      
      If txtFM2(9) = "1" Then
         m_bPrinter = False
         Set m_Device = Picture1
         m_Device.AutoRedraw = True
         m_Device.Width = 16838
         m_Device.Height = 11899
         DelPic
      Else
         m_bPrinter = True
         Set m_Device = Printer
      End If
      Call GetPleft
      Erase mStrGrp '請款對象和小計
      
      PrintPageHeader
      With RsAcc24o0
           .MoveFirst
           Do While Not .EOF
                 intP = 0
                '列印小計(2~3)
                If mStrGrp(1) <> "" And mStrGrp(1) <> "" & .Fields("A1K28") Then
                    Call PrintSubTot(mStrGrp(1))
                    mStrGrp(2) = "0"
                    mStrGrp(3) = "0"
                End If
                If mStrGrp(1) <> "" & .Fields("A1K28") Then '只在第一筆顯示
                    '國籍
                    intP = intP + 1
                    strTemp(intP) = convForm("" & .Fields("NA03"), Val(PColWidth(intP)))
                    '請款對象
                    intP = intP + 1
                    strTemp(intP) = "" & .Fields("A1K28")
                    '請款對象名稱
                    intP = intP + 1
                    strTemp(intP) = convForm("" & .Fields("A1K28N"), Val(PColWidth(intP)))
                Else
                    '國籍
                    intP = intP + 1
                    strTemp(intP) = String(Val(PColWidth(intP)), " ")
                    '請款對象
                    intP = intP + 1
                    strTemp(intP) = String(Val(PColWidth(intP)), " ")
                    '請款對象名稱
                    intP = intP + 1
                    strTemp(intP) = String(Val(PColWidth(intP)), " ")
                End If
                
                '請款單號
                intP = intP + 1
                strTemp(intP) = "" & .Fields("A1L01")
                '請款日期
                intP = intP + 1
                strTemp(intP) = ChangeTStringToTDateString("" & .Fields("A1K02"))
                '本所案號
                intP = intP + 1
                strTemp(intP) = "" & .Fields("CASENO")
                '請款項目
                intP = intP + 1
                strTemp(intP) = convForm("" & .Fields("A1J03"), Val(PColWidth(intP)))
                '折扣金額(台幣)
                intP = intP + 1
                strTemp(intP) = Format("" & .Fields("A1L07"), DDollar2)
                '折扣%
                intP = intP + 1
                strTemp(intP) = Format("" & .Fields("A1L19"), "##0%")
                '未折扣總額(台幣)
                intP = intP + 1
                strTemp(intP) = Format("" & .Fields("A1L05"), DDollar2)
                '備註(有折讓)
                intP = intP + 1
                strTemp(intP) = convForm(IIf(Val("" & .Fields("A1K06")) > 0, "有折讓", ""), Val(PColWidth(intP)))
                
                Call PrintDetail
                
                '小計(2~3)和合計(4~5)
                mStrGrp(1) = "" & .Fields("A1K28")
                mStrGrp(2) = Val(mStrGrp(2)) + Val("" & .Fields("A1L07")) '折扣金額
                mStrGrp(3) = Val(mStrGrp(3)) + Val("" & .Fields("A1L05")) '未折扣總額
                mStrGrp(4) = Val(mStrGrp(4)) + Val("" & .Fields("A1L07"))
                mStrGrp(5) = Val(mStrGrp(5)) + Val("" & .Fields("A1L05"))
                .MoveNext
           Loop
           Call PrintSubTot(mStrGrp(1)) '最後一筆小計
           '列印合計(4~5)
           Call PrintSubTot("ALL")
      End With
      
      '列印表尾
      Call PrintReportFooter
      
      If m_bPrinter = True Then
         m_Device.EndDoc
         ShowPrintOk
      ElseIf m_iPages > 0 Then
         SetPic m_iPages
         Frmacc24c0_1.m_ImageW = m_Device.Width
         Frmacc24c0_1.m_ImageH = m_Device.Height
         Frmacc24c0_1.m_iPages = m_iPages
         Frmacc24c0_1.Caption = Me.Caption
         Frmacc24c0_1.Show
      End If

      m_Device.DrawWidth = 1
   Else
      ShowNoData
      InsertQueryLog (0)
   End If
   RsAcc24o0.Close
   '執行完不清除條件

   Screen.MousePointer = vbDefault
   StatusView MsgText(101)
   
End Sub

Private Sub Form_Activate()
   If IsObject(mdiMain) Then
      ToolShow
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      StatusView MsgText(101)
   End If
End Sub

Private Sub Form_Load()
  
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   PUB_InitForm Me, 6200, 4900, strBackPicPath4

   FormClear
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc24o0 = Nothing
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   For Each oText In txtFM2
        oText.Text = ""
   Next
   lblSName.Caption = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   
   txtFM2(0).Text = "ALL"
   txtFM2(9).Text = "2"
  
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Private Function FormCheck() As Boolean
Dim bolTmp As Boolean
   
   FormCheck = False
   For Each oText In txtFM2
       txtFM2_Validate oText.Index, bolTmp
       If bolTmp = True Then
           Exit Function
       End If
   Next
   
   If MaskEdBox1.Text = MsgText(29) And MaskEdBox2.Text = MsgText(29) Then
      FormCheck = False
      MaskEdBox1.SetFocus
      MsgBox "請款日期不可空白！", , MsgText(5)
      Exit Function
   End If
   If MaskEdBox1.Text <> MsgText(29) And MaskEdBox2.Text <> MsgText(29) Then
      If Val(FCDate(MaskEdBox1.Text)) > Val(FCDate(MaskEdBox2.Text)) Then
            FormCheck = False
            MaskEdBox1.SetFocus
            MsgBox "請款日期起值不可大於迄值！", , MsgText(5)
            Exit Function
      End If
   End If
   
   If txtFM2(1) <> "" And txtFM2(2) <> "" And txtFM2(1) > txtFM2(2) Then
       MsgBox "業務區起值不可大於迄值！", vbCritical
       txtFM2(1).SetFocus
       Call txtFM2_GotFocus(1)
       Exit Function
   End If
   
   If txtFM2(3) <> "" And txtFM2(4) <> "" And txtFM2(3) > txtFM2(4) Then
       MsgBox "國籍起值不可大於迄值！", vbCritical
       txtFM2(3).SetFocus
       Call txtFM2_GotFocus(3)
       Exit Function
   End If
   
   If txtFM2(5) <> "" And txtFM2(6) <> "" And txtFM2(5) > txtFM2(6) Then
       MsgBox "請款對象起值不可大於迄值！", vbCritical
       txtFM2(5).SetFocus
       Call txtFM2_GotFocus(5)
       Exit Function
   End If

   FormCheck = True
End Function


Private Sub PrintSubTot(strSalesNo As String)
   
   For ii = LBound(strTemp) To UBound(strTemp)
      strTemp(ii) = ""
   Next ii
   If strSalesNo <> "ALL" Then
      strTemp(4) = "小計："
      strTemp(8) = Format(mStrGrp(2), DDollar2)
      strTemp(10) = Format(mStrGrp(3), DDollar2)
   Else
      strTemp(4) = "合計："
      strTemp(8) = Format(mStrGrp(4), DDollar2)
      strTemp(10) = Format(mStrGrp(5), DDollar2)
   End If
   
    PrintNewLine
    DrawLine 4
    m_Device.FontBold = True
    PrintDetail
    m_Device.FontBold = False
  
End Sub

Private Sub DelPic()
   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_*.tmp"
   If Dir(strPicFileName) <> "" Then
      Kill strPicFileName
   End If
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
End Sub

Private Sub SetPic(idx As Integer)

   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_" & idx & ".tmp"
   SavePicture Picture1.Image, strPicFileName
   
   '要用覆蓋的否則會錯誤--VB Bug
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
   
End Sub

Sub GetPleft()
   Dim iCols As Integer
   Dim tmpArr1 As Variant, tmpArr2 As Variant
   If m_bPrinter = True Then
      m_Device.Orientation = 2
   End If
   m_iStartX = 300
   m_iStartY = 400
   m_iGap = 70
   m_iPageHeight = m_Device.ScaleHeight
   m_iLineHeight = 300
   m_iMargin = 500
   iPage = 0: m_iPages = 0
   
   strExc(1) = "國　籍,請款對象,請款對象名稱,請款單號,請款日期,本 所 案 號,請 款 項 目,折扣金額,折扣%,未折扣總額,備　註"
   strExc(2) = "6,9,28,9,9,15,16,8,6,10,6"
   tmpArr1 = Split(strExc(1), ",")
   tmpArr2 = Split(strExc(2), ",")
   iCols = UBound(tmpArr1) + 1
   
   Erase PLeft
   Erase PColName
   Erase PColWidth
   
   ReDim PLeft(1 To iCols + 1)
   ReDim PColName(1 To iCols)
   ReDim PColWidth(1 To iCols)
   ReDim strTemp(1 To iCols)
   
   ii = 1
   iCols = 125 '字元寬度(大約)
   PLeft(ii) = m_iStartX
   
   For ii = 1 To UBound(PColName)
        PColName(ii) = tmpArr1(ii - 1)
        PColWidth(ii) = tmpArr2(ii - 1)
        PLeft(ii + 1) = PLeft(ii) + iCols * Val(PColWidth(ii)) + m_iGap
   Next ii
 
End Sub

Private Sub PrintPageHeader()
   Dim strTmp As String
   
   iPage = iPage + 1
   m_iPages = m_iPages + 1
   If m_iPages > 1 Then
      If m_bPrinter = False Then
         SetPic m_iPages - 1
      ElseIf iPage > 1 Then
         m_Device.NewPage
      End If
   End If

   m_Device.FontName = "新細明體"
   
   iPrint = m_iStartY
   m_Device.Font.Size = 14
   m_Device.Font.Bold = True
   
   strTmp = Me.Caption
   m_Device.CurrentX = (m_Device.ScaleWidth - m_Device.TextWidth(strTmp)) / 2
   m_Device.CurrentY = iPrint
   m_Device.Print strTmp
   
   PrintNewLine 500
   
   m_Device.Font.Size = 10
   m_Device.Font.Bold = True
   m_Device.Font.Underline = False
   
   strTmp = "列印人員:" & strUserName
   m_Device.CurrentX = m_iStartX
   m_Device.CurrentY = iPrint
   m_Device.Print strTmp
   strTmp = "請款日期:" & MaskEdBox1 & " ~ " & MaskEdBox2
   m_Device.CurrentX = (m_Device.ScaleWidth - m_Device.TextWidth(strTmp)) / 2
   m_Device.CurrentY = iPrint
   m_Device.Print strTmp
   strTmp = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   m_Device.CurrentX = m_Device.ScaleWidth - m_iMargin - 2500
   m_Device.CurrentY = iPrint
   m_Device.Print strTmp
   PrintNewLine
   
   strTmp = "頁    次：" & str(iPage)
   m_Device.CurrentX = m_Device.ScaleWidth - m_iMargin - 2500
   m_Device.CurrentY = iPrint
   m_Device.Print strTmp
   
   If txtFM2(0) <> "ALL" Then
      strTmp = "系統類別:" & txtFM2(0)
      m_Device.CurrentX = (m_Device.ScaleWidth - m_Device.TextWidth(strTmp)) / 2
      m_Device.CurrentY = iPrint
      m_Device.Print strTmp
      PrintNewLine
   End If
   
   If Trim(txtFM2(7)) <> "" Then
      strTmp = "案件性質:" & txtFM2(7)
      m_Device.CurrentX = (m_Device.ScaleWidth - m_Device.TextWidth(strTmp)) / 2
      m_Device.CurrentY = iPrint
      m_Device.Print strTmp
      PrintNewLine
   End If
   
   If Trim(txtFM2(1) & txtFM2(2)) <> "" Then
      strTmp = "業務區:" & txtFM2(1) & " ~ " & txtFM2(2)
      m_Device.CurrentX = (m_Device.ScaleWidth - m_Device.TextWidth(strTmp)) / 2
      m_Device.CurrentY = iPrint
      m_Device.Print strTmp
      PrintNewLine
   End If
   
   If Trim(txtFM2(3) & txtFM2(4)) <> "" Then
      strTmp = "國籍:" & txtFM2(3) & " ~ " & txtFM2(4)
      m_Device.CurrentX = (m_Device.ScaleWidth - m_Device.TextWidth(strTmp)) / 2
      m_Device.CurrentY = iPrint
      m_Device.Print strTmp
      PrintNewLine
   End If
   
   If Trim(txtFM2(5) & txtFM2(6)) <> "" Then
      strTmp = "請款對象:" & txtFM2(5) & " ~ " & txtFM2(6)
      m_Device.CurrentX = (m_Device.ScaleWidth - m_Device.TextWidth(strTmp)) / 2
      m_Device.CurrentY = iPrint
      m_Device.Print strTmp
      PrintNewLine
   End If
   
   If Trim(txtFM2(8)) <> "" Then
      strTmp = "智權人員:" & txtFM2(8) & " " & lblSName.Caption
      m_Device.CurrentX = (m_Device.ScaleWidth - m_Device.TextWidth(strTmp)) / 2
      m_Device.CurrentY = iPrint
      m_Device.Print strTmp
      PrintNewLine
   End If

   '判斷只有請款日期，要多跳一行
   strExc(1) = ""
   For Each oText In txtFM2
        If oText.Index <> 9 Then
            strExc(1) = strExc(1) & Trim(oText.Text)
        End If
   Next
   If strExc(1) = "ALL" Then
       PrintNewLine
   End If
   
   PrintPageHeader1
   m_Device.Font.Bold = False
End Sub

Private Sub PrintPageHeader1()
   For intI = 1 To UBound(PColName)
      If intI > 7 Then
         m_Device.CurrentX = PLeft(intI + 1) - m_Device.TextWidth(PColName(intI)) - m_iGap
      Else
         m_Device.CurrentX = PLeft(intI)
      End If
      m_Device.CurrentY = iPrint
      m_Device.Print PColName(intI)
   Next
   PrintNewLine
   DrawLine
End Sub

Private Sub PrintDetail()
   Dim iCol As Integer
   PrintNewLine
   For iCol = LBound(strTemp) To UBound(strTemp)
      If iCol > 7 Then
         m_Device.CurrentX = PLeft(iCol + 1) - m_Device.TextWidth(strTemp(iCol)) - m_iGap
      Else
         m_Device.CurrentX = PLeft(iCol)
      End If
      m_Device.CurrentY = iPrint
      m_Device.Print strTemp(iCol)
   Next
End Sub

'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)
   Dim iSize As Integer
   
   PrintNewLine , 3
   DrawLine
   PrintNewLine
   
   iSize = m_Device.Font.Size
   m_Device.Font.Size = 12
   m_Device.Font.Bold = True
   strExc(1) = "*** 結束 ***"
   m_Device.CurrentX = (m_Device.ScaleWidth - m_Device.TextWidth(strExc(1))) / 2
   m_Device.CurrentY = iPrint
   m_Device.Print strExc(1)

   m_Device.Font.Size = iSize
   m_Device.Font.Bold = False
End Sub

Private Sub DrawLine(Optional iStartCol As Integer, Optional iEndCol As Integer, Optional lngEndPoint As Long)
   Dim lngFrom As Long, lngTo As Long
   If iStartCol = 0 Then
      lngFrom = PLeft(LBound(PLeft))
   Else
      lngFrom = PLeft(iStartCol)
   End If
   If iEndCol = 0 Then
      If lngEndPoint > 0 Then
         lngTo = lngEndPoint
      Else
         lngTo = PLeft(UBound(PLeft))
      End If
   Else
      lngTo = PLeft(iEndCol)
   End If
   m_Device.DrawWidth = 4
   m_Device.Line (lngFrom, iPrint)-(lngTo, iPrint)
   iPrint = iPrint - m_iLineHeight / 2
End Sub

Private Sub PrintNewLine(Optional ByVal iHeight As Integer = 0, Optional ByVal p_iExtraLines As Integer = 2)
   If iHeight = 0 Then
      iHeight = m_iLineHeight
   End If
   
   iPrint = iPrint + iHeight
   
   If iPrint >= (m_iPageHeight - m_iMargin - p_iExtraLines * m_iLineHeight) Then
      DrawLine
      PrintPageHeader
      iPrint = iPrint + m_iLineHeight
   End If
End Sub

Private Sub txtFM2_GotFocus(Index As Integer)
    TextInverse txtFM2(Index)
End Sub

Private Sub txtFM2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
     Select Case Index
         Case 3, 4, 9
              KeyAscii = Pub_NumAscii(KeyAscii)
         Case Else
              KeyAscii = UpperCase(KeyAscii)
     End Select
End Sub

Private Sub txtFM2_LostFocus(Index As Integer)
    Select Case Index
        Case 3 '國籍
             If txtFM2(3) <> "" And txtFM2(4) = "" Then
                txtFM2(4) = txtFM2(3)
             End If
        Case 5 '請款對象
             If txtFM2(5) <> "" Then
                txtFM2(6) = Left(txtFM2(5), 6) & "ZZZ"
             End If
    End Select
End Sub

Private Sub txtFM2_Validate(Index As Integer, Cancel As Boolean)

    Select Case Index
        Case 0      '系統類別
            '檢查跨部門系統,不可輸入案件性質
            txtFM2(Index) = Replace(txtFM2(Index), " ", "")
            If Trim(txtFM2(Index)) = "" Then
                 MsgBox "系統類別不可空白！", vbCritical
                 GoTo EXITSUB
            ElseIf txtFM2(Index) <> "ALL" Then
                If PUB_CheckSKAddCross(strUserNum, Systemkind_g, True, txtFM2(Index)) = False Then
                       GoTo EXITSUB
                End If
                strExc(0) = txtFM2(Index)
            Else
                strExc(0) = Systemkind_g
            End If
            '檢查跨部門系統,不可輸入案件性質
            strSql = "select count(distinct sk02) cnt from systemkind where sk01 in (" & GetAddStr(strExc(0)) & ") "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
                If Val("" & RsTemp.Fields("cnt")) > 1 Then
                    txtFM2(7).Text = ""
                    txtFM2(7).Locked = True
                Else
                    txtFM2(7).Locked = False
                End If
            End If
        Case 1, 2   '業務區
            If Index = 2 And txtFM2(1) <> "" Then
                If txtFM2(3) > txtFM2(4) Then
                    MsgBox "業務區迄值不可小於起值！", vbCritical
                    GoTo EXITSUB
                End If
            End If

        Case 3, 4 '國籍
            If txtFM2(Index).Text = "" Then Exit Sub
            If Val(txtFM2(Index)) <> txtFM2(Index) Then
                MsgBox "請輸入國籍代號！", vbCritical
                GoTo EXITSUB
            End If
            
        Case 5, 6   '請款對象
            If txtFM2(Index).Text = "" Then Exit Sub
            If Left(txtFM2(Index), 1) <> "Y" Then
                MsgBox "請款對象請輸入代理人編號！", vbCritical
                GoTo EXITSUB
            Else
                If Len(txtFM2(Index)) < 9 Then
                    If Index = 5 Then
                        txtFM2(Index).Text = Left(txtFM2(Index) & String(9, "0"), 9)
                    Else
                        txtFM2(Index).Text = Left(txtFM2(Index) & String(9, "Z"), 9)
                    End If
                End If
            End If

        Case 8   '智權人員
             If txtFM2(Index).Text = "" Then Exit Sub
             lblSName.Caption = GetStaffName(txtFM2(Index), True)
             If lblSName.Caption = "" Then
                MsgBox "智權人員輸入錯誤！", vbCritical
                GoTo EXITSUB
             End If
        Case 9   '輸出方式
             If txtFM2(Index) <> "1" And txtFM2(Index) <> "2" Then
                MsgBox "輸出方式請輸入1或2！", vbCritical
                GoTo EXITSUB
             End If
    End Select
    
    Exit Sub
    
EXITSUB:
    txtFM2(Index).SetFocus
    txtFM2_GotFocus Index
    Cancel = True
End Sub


