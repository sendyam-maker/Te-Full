VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm081031_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "智財顧問專業分配比例"
   ClientHeight    =   4980
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8028
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8028
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   30
      Left            =   6720
      MaxLength       =   5
      TabIndex        =   60
      Text            =   "30"
      Top             =   3180
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   31
      Left            =   6720
      MaxLength       =   5
      TabIndex        =   59
      Text            =   "31"
      Top             =   3516
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   32
      Left            =   6720
      MaxLength       =   5
      TabIndex        =   58
      Text            =   "32"
      Top             =   3864
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   33
      Left            =   6720
      MaxLength       =   5
      TabIndex        =   57
      Text            =   "33"
      Top             =   4212
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   34
      Left            =   6720
      MaxLength       =   5
      TabIndex        =   56
      Text            =   "34"
      Top             =   4560
      Width           =   825
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "存檔及通知相關人"
      Height          =   375
      Left            =   1500
      TabIndex        =   49
      Top             =   60
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton CmdCaculate 
      Caption         =   "重新計算比例"
      Height          =   375
      Left            =   60
      TabIndex        =   48
      Top             =   60
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Index           =   20
      Left            =   5292
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   47
      TabStop         =   0   'False
      Text            =   "20"
      Top             =   3180
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Index           =   21
      Left            =   5292
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   46
      TabStop         =   0   'False
      Text            =   "21"
      Top             =   3516
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Index           =   22
      Left            =   5292
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   45
      TabStop         =   0   'False
      Text            =   "22"
      Top             =   3864
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Index           =   23
      Left            =   5292
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   44
      TabStop         =   0   'False
      Text            =   "23"
      Top             =   4212
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Index           =   24
      Left            =   5292
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   43
      TabStop         =   0   'False
      Text            =   "24"
      Top             =   4560
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtCP15 
      Height          =   300
      Left            =   1008
      MaxLength       =   4
      TabIndex        =   16
      Top             =   2496
      Width           =   855
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "確定(&E)"
      Default         =   -1  'True
      Height          =   375
      Left            =   5580
      TabIndex        =   5
      Top             =   60
      Width           =   885
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Index           =   14
      Left            =   3444
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "14"
      Top             =   4560
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Index           =   13
      Left            =   3444
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "13"
      Top             =   4212
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Index           =   12
      Left            =   3444
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   35
      TabStop         =   0   'False
      Text            =   "12"
      Top             =   3864
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Index           =   11
      Left            =   3444
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   34
      TabStop         =   0   'False
      Text            =   "11"
      Top             =   3516
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Index           =   10
      Left            =   3444
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   33
      TabStop         =   0   'False
      Text            =   "10"
      Top             =   3180
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   4
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "4"
      Top             =   4560
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   3
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "3"
      Top             =   4212
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   2
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "2"
      Top             =   3864
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   1
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "1"
      Top             =   3516
      Width           =   825
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   0
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "0"
      Top             =   3180
      Width           =   825
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "回前畫面(&U)"
      Height          =   375
      Left            =   6540
      TabIndex        =   6
      Top             =   75
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "調整比例％"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   228
      Index           =   15
      Left            =   6720
      TabIndex        =   55
      Top             =   2880
      Width           =   1140
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(9)"
      Height          =   288
      Index           =   9
      Left            =   4704
      TabIndex        =   54
      Top             =   2184
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "（收文服務費—行政流程費用）"
      Height          =   228
      Index           =   14
      Left            =   5424
      TabIndex        =   53
      Top             =   2196
      Width           =   2532
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(7)"
      Height          =   288
      Index           =   7
      Left            =   1752
      TabIndex        =   52
      Top             =   2208
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "專業部分配費用："
      Height          =   228
      Index           =   13
      Left            =   3150
      TabIndex        =   51
      Top             =   2198
      Width           =   1524
   End
   Begin VB.Label Label1 
      Caption         =   "行政流程費用20%："
      Height          =   228
      Index           =   12
      Left            =   90
      TabIndex        =   50
      Top             =   2193
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "未發文次數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   228
      Index           =   11
      Left            =   5292
      TabIndex        =   42
      Top             =   2880
      Visible         =   0   'False
      Width           =   1128
   End
   Begin MSForms.Label lblFM2 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   41
      Top             =   1230
      Width           =   5715
      VariousPropertyBits=   27
      Caption         =   "lblFM2(1)"
      Size            =   "10081;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(8)"
      Height          =   285
      Index           =   8
      Left            =   1050
      TabIndex        =   40
      Top             =   1230
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "當事人："
      Height          =   225
      Index           =   17
      Left            =   90
      TabIndex        =   39
      Top             =   1230
      Width           =   945
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(6)"
      Height          =   288
      Index           =   6
      Left            =   4104
      TabIndex        =   38
      Top             =   2496
      Width           =   2592
   End
   Begin VB.Label Label1 
      Caption         =   "顧問期間："
      Height          =   228
      Index           =   16
      Left            =   3156
      TabIndex        =   37
      Top             =   2520
      Width           =   948
   End
   Begin VB.Label lblSys 
      Caption         =   "CFT"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Index           =   4
      Left            =   120
      TabIndex        =   32
      Top             =   4596
      Width           =   552
   End
   Begin VB.Label lblSys 
      Caption         =   "CFP"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Index           =   3
      Left            =   120
      TabIndex        =   31
      Top             =   4248
      Width           =   552
   End
   Begin VB.Label lblSys 
      Caption         =   "ACS"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Index           =   2
      Left            =   120
      TabIndex        =   30
      Top             =   3900
      Width           =   552
   End
   Begin VB.Label lblSys 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   3564
      Width           =   552
   End
   Begin VB.Label lblSys 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Index           =   0
      Left            =   120
      TabIndex        =   28
      Top             =   3216
      Width           =   552
   End
   Begin MSForms.Label lblFM2 
      Height          =   285
      Index           =   0
      Left            =   4860
      TabIndex        =   27
      Top             =   1560
      Width           =   1845
      VariousPropertyBits=   27
      Caption         =   "lblFM2(0)"
      Size            =   "3254;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(5)"
      Height          =   288
      Index           =   5
      Left            =   4104
      TabIndex        =   26
      Top             =   1872
      Width           =   732
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(4)"
      Height          =   288
      Index           =   4
      Left            =   1248
      TabIndex        =   25
      Top             =   1872
      Width           =   1632
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(3)"
      Height          =   285
      Index           =   3
      Left            =   4104
      TabIndex        =   24
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(2)"
      Height          =   285
      Index           =   2
      Left            =   1020
      TabIndex        =   23
      Top             =   1551
      Width           =   1635
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(1)"
      Height          =   285
      Index           =   1
      Left            =   4110
      TabIndex        =   22
      Top             =   510
      Width           =   1425
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(0)"
      Height          =   285
      Index           =   0
      Left            =   1050
      TabIndex        =   21
      Top             =   510
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "實際分配比例％"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   228
      Index           =   10
      Left            =   3444
      TabIndex        =   20
      Top             =   2880
      Width           =   1548
   End
   Begin VB.Label Label1 
      Caption         =   "收文分配比例％"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   228
      Index           =   9
      Left            =   1596
      TabIndex        =   19
      Top             =   2880
      Width           =   1548
   End
   Begin VB.Label Label1 
      Caption         =   "專業部門"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   228
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   948
   End
   Begin VB.Label Label1 
      Caption         =   "簽約時數："
      Height          =   228
      Index           =   7
      Left            =   90
      TabIndex        =   17
      Top             =   2520
      Width           =   948
   End
   Begin VB.Label Label1 
      Caption         =   "收文點數："
      Height          =   225
      Index           =   6
      Left            =   3150
      TabIndex        =   15
      Top             =   1879
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   225
      Index           =   5
      Left            =   3180
      TabIndex        =   14
      Top             =   510
      Width           =   945
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1050
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   870
      Width           =   6615
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11668;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "收文日期："
      Height          =   225
      Index           =   4
      Left            =   90
      TabIndex        =   11
      Top             =   1551
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   225
      Index           =   3
      Left            =   3150
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   225
      Index           =   2
      Left            =   90
      TabIndex        =   9
      Top             =   870
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "收文服務費："
      Height          =   228
      Index           =   1
      Left            =   96
      TabIndex        =   8
      Top             =   1872
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   510
      Width           =   945
   End
End
Attribute VB_Name = "frm081031_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/04/26 Form2.0已修改 lblFM2、Combo1
'Create by Lydia 2021/04/26 智財顧問專業分配比例
Option Explicit
Dim m_PrevForm As Form  '前一畫面
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String '本所案號
Dim m_CP09 As String  '收文號
Dim m_CP53 As String, m_CP54 As String '顧問期間
Dim m_Status As String '狀態: M-可修改資料, Q-查詢, U-重新計算作業
Dim strTmpQ As String
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset
Dim oObj
'Modified by Lydia 2023/10/06
'Const m_Max As Integer = 6  '系統別的數量
'Const strOrderBy As String = " 'P','0','T','1','ACS','2','CFP','3','CFT','4','L','5','CFL','6','9' " 'SQL系統別的順序
Const m_Max As Integer = 4  '系統別的數量
'Modified by Lydia 2024/05/15 調整系統別，若系統別有變更請一併更新ACS112STATISTICS;ex.ACS-000173的相關案為S案
'Const strOrderBy As String = " 'P','0','T','1','ACS','2','CFP','3','CFT','4','9' " 'SQL系統別的順序
Const strOrderBy As String = " DECODE(SK02||SK03,'10','0','50','0','20','1','60','1','11','3','51','3','21','4','61','4',DECODE(CP01,'ACS','2','9')) " 'SQL系統別的順序
Dim m_Caculate As String  '重新計算作業之狀態：0-顯示未收文次數, 1-重新計算

Public Sub SetParent(ByVal pFrm As Form, ByVal pCP09 As String, ByVal pStatus As String)
    Set m_PrevForm = pFrm
    m_CP09 = pCP09
    m_Status = pStatus
End Sub

Private Sub Cmd1_Click()
 
 If CheckDataValidate = True Then 'Added by Lydia 2023/11/29 原本的檢查改成模組
    cmd1.Enabled = False
    If FormSave = True Then
        cmd1.Enabled = True
        Call cmdExit_Click
        Exit Sub
    End If
    cmd1.Enabled = True
 End If
End Sub

'Added by Lydia 2023/11/29
Private Function CheckDataValidate() As Boolean
Dim tmpBol As Boolean

   For intI = 0 To m_Max
       Call Txtdata_Validate(intI, tmpBol)
       If tmpBol = True Then
         Exit Function
       End If
   Next intI
   
   strExc(1) = "0"
   For intI = 0 To m_Max
       strExc(1) = Val(strExc(1)) + Val(txtData(intI))
   Next intI
   If Val(strExc(1)) < 100 Then
       MsgBox "收文分配比例不可小於100%", vbCritical, "檢核資料"
       Exit Function
   End If
   'Added by Lydia 2023/11/29
    If txtData(30).Visible = True And txtData(30).Locked = False Then
       strExc(1) = "0"
       For intI = 30 To 30 + m_Max
           strExc(1) = Val(strExc(1)) + Val(txtData(intI))
       Next intI
       If Val(strExc(0)) > 0 And Val(strExc(1)) < 100 Then
          MsgBox "調整比例不可小於100%", vbCritical, "檢核資料"
          Exit Function
       End If
    End If
   'end 2023/11/29
   
   'Modified by Lydia 2023/11/22 簽約時數欄位仍保留，但屬非必填欄位(公告1121011-03)
   'If Val(Trim(txtCP15)) <= 0 Then
   '    MsgBox "請輸入簽約時數 ！", vbCritical, "檢核資料"
   If Trim(txtCP15) = "" Then
       MsgBox "簽約時數不可空白！", vbCritical, "檢核資料"
   'end 2023/11/22
       txtCP15.SetFocus
       txtCP15_GotFocus
       Exit Function
   End If
   
   CheckDataValidate = True
End Function

Private Sub CmdCaculate_Click()
   If CheckDataValidate = True Then 'Added by Lydia 2023/11/29
      m_Caculate = "1" '重新計算：並且帶出實際分配比例
      Call ProcCaculate
      cmdSave.Enabled = True
   End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub doQuery()
    
    'Modified by Lydia 2023/10/06 +acc1u0
    'strTmpQ = "select cp01||'-'||cp02||decode(cp03,'0',null,'-'||cp03)||decode(cp04,'00',null,'-'||cp04) caseno,cp09," & _
                     "lc05,lc06,lc07,lc11, nvl(cu04,nvl(cu05,cu06)) lc11n,cp05,cp13,st02 as cp13n," & _
                     "cp16,cp18,cp15,cp01,cp02,cp03,cp04,cp53,cp54 " & _
                     "from caseprogress, lawcase, staff, customer " & _
                     "where cp09='" & m_CP09 & "' and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp13=st01(+) " & _
                     "and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) "
    strTmpQ = "select cp01||'-'||cp02||decode(cp03,'0',null,'-'||cp03)||decode(cp04,'00',null,'-'||cp04) caseno,cp09," & _
                     "lc05,lc06,lc07,lc11, nvl(cu04,nvl(cu05,cu06)) lc11n,cp05,cp13,st02 as cp13n," & _
                     "cp16,cp17,cp18,cp15,cp01,cp02,cp03,cp04,cp53,cp54,sum(nvl(a1u07,0)/1000) a1u07,sum(nvl(a1u09,0)/1000) a1u09 " & _
                     "from caseprogress, lawcase, staff, customer, acc1u0 " & _
                     "where cp09='" & m_CP09 & "' and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp13=st01(+) " & _
                     "and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) and cp09=a1u03(+) and cp60=a1u02(+) " & _
                     "group by cp01||'-'||cp02||decode(cp03,'0',null,'-'||cp03)||decode(cp04,'00',null,'-'||cp04),cp09,lc05,lc06,lc07,lc11, nvl(cu04,nvl(cu05,cu06)),cp05,cp13,st02,cp16,cp17,cp18,cp15,cp01,cp02,cp03,cp04,cp53,cp54 "
   
    intQ = 0
    Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
    If intQ = 0 Then
        Exit Sub
    End If
    
    intQ = 0
    Combo1.AddItem "中：" & rsQuery.Fields("lc05"), 0
    If rsQuery.Fields("lc05") <> "" Then intQ = 1
    Combo1.AddItem "英：" & rsQuery.Fields("lc06"), 1
    If rsQuery.Fields("lc06") <> "" Then intQ = 2
    Combo1.AddItem "日：" & rsQuery.Fields("lc07"), 2
    If rsQuery.Fields("lc07") <> "" Then intQ = 3
    Combo1.ListIndex = intQ - 1
    
    '本所案號
    lblData(0).Caption = "" & rsQuery.Fields("caseno")
    m_CP01 = "" & rsQuery.Fields("cp01")
    m_CP02 = "" & rsQuery.Fields("cp02")
    m_CP03 = "" & rsQuery.Fields("cp03")
    m_CP04 = "" & rsQuery.Fields("cp04")
    '收文號
    lblData(1).Caption = "" & rsQuery.Fields("cp09")
    '當事人
    lblData(8).Caption = "" & rsQuery.Fields("lc11")
    lblFM2(1).Caption = "" & rsQuery.Fields("lc11n")
    '收文日期、費用、點數
    lblData(2).Caption = ChangeWStringToTDateString("" & rsQuery.Fields("cp05"))
    'Modified by Lydia 2023/10/06 費用改成「收文服務費」＝收文費用-（收文規費-銷帳規費）-銷帳服務費
    'lblData(4).Caption = Format("" & rsQuery.Fields("cp16"), DDollar2)
    'lblData(5).Caption = "" & rsQuery.Fields("cp18")
    lblData(4).Caption = Format(Val("" & rsQuery.Fields("cp16")) - (Val("" & rsQuery.Fields("cp17")) - Val("" & rsQuery.Fields("a1u09"))) - Val("" & rsQuery.Fields("a1u07")), DDollar2)
    lblData(5).Caption = Format(Val(Format(lblData(4), "##0.000")) / 1000, "##,##0.000")
    'Added by Lydia 2023/10/06 行政流程費用，設定為原收文費用之20%，直接歸顧服組，另一欄位為分配費用，即收文費用扣除行政流程費用後之費用。
    If Val("" & rsQuery.Fields("cp16")) > 0 Then
      lblData(7).Caption = Format(Val(Format(lblData(4), "##0.000")) * 0.2, DDollar2)
      lblData(9).Caption = Format((Val(Format(lblData(4), "##0.000"))) - Val(Format(lblData(7), "##0")), DDollar2)
    End If
    'end 2023/10/06
    '智權人員
    lblData(3).Caption = "" & rsQuery.Fields("cp13")
    lblFM2(0).Caption = "" & rsQuery.Fields("cp13n")
    '簽約時數
    txtCP15.Text = "" & rsQuery.Fields("cp15")
    txtCP15.Tag = txtCP15.Text
    '顧問期間
    lblData(6).Caption = ChangeWStringToTDateString("" & rsQuery.Fields("cp53")) & " ～ " & ChangeWStringToTDateString("" & rsQuery.Fields("cp54"))
    m_CP53 = "" & rsQuery.Fields("cp53")
    m_CP54 = "" & rsQuery.Fields("cp54")
    
    'Modified by Lydia 2024/05/15
    'strTmpQ = "select decode(ar02," & strOrderBy & " ) ord1, a1.* from ACSPFrate a1 " & _
                     "where ar01='" & m_CP09 & "' order by ord1 "
    strTmpQ = "select " & Replace(UCase(strOrderBy), "CP01", "AR02") & " as ord1, a1.* from ACSPFrate a1, systemkind " & _
                     "where ar01='" & m_CP09 & "' and ar02=sk01(+) order by ord1 "
    intQ = 1
    Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
    If intQ = 0 Then
        '先預設P 40%、T 40%、ACS 20%
        If m_Status = "M" Then  '限維護模式
           txtData(0) = "40"
           txtData(1) = "40"
           txtData(2) = "20"
        End If
    Else
        rsQuery.MoveFirst
        Do While Not rsQuery.EOF
            If "" & rsQuery.Fields("ord1") <> "9" Then
                '收文分配比例
                If Val("" & rsQuery.Fields("ar03")) <> 0 Then
                    txtData(Val("" & rsQuery.Fields("ord1"))).Text = Val("" & rsQuery.Fields("ar03"))
                    txtData(Val("" & rsQuery.Fields("ord1"))).Tag = txtData(Val("" & rsQuery.Fields("ord1"))).Text
                    lblSys(Val("" & rsQuery.Fields("ord1"))).Tag = "Y"
                End If
                '實際分配比例
                If Val("" & rsQuery.Fields("ar04")) <> 0 Then
                    txtData(Val("" & rsQuery.Fields("ord1")) + 10).Text = Val("" & rsQuery.Fields("ar04"))
                    txtData(Val("" & rsQuery.Fields("ord1")) + 10).Tag = txtData(Val("" & rsQuery.Fields("ord1")) + 10).Text
                    lblSys(Val("" & rsQuery.Fields("ord1"))).Tag = "Y"
                End If
                'Added by Lydia 2023/11/29 人工調整分配比例(調整比例)
                If Val("" & rsQuery.Fields("ar10")) > 0 Then
                    txtData(Val("" & rsQuery.Fields("ord1")) + 30).Text = Val("" & rsQuery.Fields("ar10"))
                    txtData(Val("" & rsQuery.Fields("ord1")) + 30).Tag = txtData(Val("" & rsQuery.Fields("ord1")) + 30).Text
                    lblSys(Val("" & rsQuery.Fields("ord1"))).Tag = "Y"
                End If
                'end 2023/11/29
            'Added by Lydia 2024/05/15
            Else
                MsgBox rsQuery.Fields("ar02") & "尚未列入專業部門，請洽電腦中心！"
                CmdCaculate.Enabled = False
                cmdSave.Enabled = False
            'end 2024/05/15
            End If
            rsQuery.MoveNext
        Loop
    End If
    
    If m_Status = "U" Then '預設帶出未發文次數
         m_Caculate = "0"
         Call ProcCaculate
         cmdSave.Enabled = False
    End If
End Sub

Private Sub ProcCaculate()
Dim intR As Integer, strR1 As String
Dim rsRd As New ADODB.Recordset
Dim strTmpA As String, strTmp2 As String
Dim strRate As String, strTRate As String
Dim pWHours As String
    
    '計算方法若有異動,請一併修改ACS112STATISTICS
    Call Proc_R100101_i(m_CP01, m_CP02, m_CP03, m_CP04)
    'ACS本案非112之收文
    strR1 = "select cp01,cp02,cp03,cp04,cp09,cp10,cpm03,nvl(cp113,0) cp113,cp158 from caseprogress,casepropertymap " & _
                "where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "' " & _
                "and cp05>=" & m_CP53 & " and cp05<=" & m_CP54 & " and cp10 <> '112' and cp159=0 " & _
                "and cp09<'C' and cp01=cpm01(+) and cp10=cpm02(+) "
    '相關卷號：排除C,D類收文
    strR1 = strR1 & "Union All select cp01,cp02,cp03,cp04,cp09,cp10,cpm03,nvl(cp113,0) cp113,cp158 from caseprogress,casepropertymap " & _
                 "where (cp01,cp02,cp03,cp04) in (select R001001,R001002,R001003,R001004 from R100101_i where id='" & strUserNum & "' and R001001<>'ACS' ) " & _
                 "and cp05>=" & m_CP53 & " and cp05<=" & m_CP54 & " and cp159=0 and cp09<'C' and cp01=cpm01(+) and cp10=cpm02(+) "
                 
    '先抓已服務時數(工作時數總計)
    strTmpA = "select count(cp09) cnt ,sum(cp113) tot1 from (" & strR1 & " ) "
    intR = 1
    Set rsRd = ClsLawReadRstMsg(intR, strTmpA)
    If intR = 1 Then
        pWHours = Val("" & rsRd.Fields("tot1"))
    End If
    
    '以系統類別統計收文次數及工作時數、尚未發文次數
    'Modified by Lydia 2024/05/15
    'strTmpA = "select decode(cp01, " & strOrderBy & ") as ord1 ,cp01, count(cp09) cnt ,sum(cp113) tot1, sum(decode(cp158,0,1,0)) tot2 " & _
                     "from (" & strR1 & " ) group by decode(cp01, " & strOrderBy & "), cp01 order by ord1 "
    strTmpA = "select " & strOrderBy & " as ord1 ,cp01, count(cp09) cnt ,sum(cp113) tot1, sum(decode(cp158,0,1,0)) tot2 " & _
                     "from (" & strR1 & " ),systemkind where cp01=sk01(+) group by " & strOrderBy & ", cp01 order by ord1 "
    intR = 1
    Set rsRd = ClsLawReadRstMsg(intR, strTmpA)
    If intR = 1 Then
        strTmp2 = "0"
        rsRd.MoveFirst
        Do While Not rsRd.EOF
           'Added by Lydia 2024/05/15
           If "" & rsRd.Fields("ord1") = "9" Then
               MsgBox rsRd.Fields("cp01") & "尚未列入專業部門，請洽電腦中心！"
               CmdCaculate.Enabled = False
               cmdSave.Enabled = False
           Else
           'end 2024/05/15
              If Val("" & rsRd.Fields("cnt")) <> 0 Or Val("" & rsRd.Fields("tot1")) <> 0 Or Val("" & rsRd.Fields("tot2")) <> 0 Then
                 '實際分配比例
                 If Val(pWHours) = 0 Or Val("" & rsRd.Fields("tot1")) = 0 Then
                    strRate = "0"
                 Else
                    strRate = Round((Val("" & rsRd.Fields("tot1")) / Val(pWHours)) * 100, 2)
                    '記錄實際分配比例, 於最後一筆修正全部加總=100%
                    strTRate = Val(strTRate) + Val(strRate)
                    strTmp2 = Val(strTmp2) + Val("" & rsRd.Fields("tot1"))
                    If Val(pWHours) = Val(strTmp2) And Val(strTRate) <> 100 Then
                       strRate = strRate + (100 - strTRate)
                    End If
                 End If
              End If
              '記錄實際分配比例
              If m_Caculate = "1" Then txtData(10 + Val("" & rsRd.Fields("ord1"))).Text = strRate
              '未發文次數
              txtData(20 + Val("" & rsRd.Fields("ord1"))).Text = Val("" & rsRd.Fields("tot2"))
           End If 'Added by Lydia 2024/05/15
           rsRd.MoveNext
        Loop
    End If
    
End Sub

Private Sub CmdSave_Click()
  'Added by Lydia 2023/11/29
  If CheckDataValidate = True Then
     If FormSave = True Then
  'end 2023/11/29
       If PUB_RunStrMenu105(m_CP01, m_CP02, m_CP03, m_CP04, m_CP09) = True Then
          MsgBox "完成存檔及通知相關人作業！", vbInformation
          '重新載入資料
          Call ClearForm
          Call doQuery
       Else
          MsgBox "存檔失敗！", vbCritical
       End If
  'Added by Lydia 2023/11/29
     End If
  End If
  'end 2023/11/29
End Sub

Private Sub Form_Load()

    MoveFormToCenter Me
    
    If m_Status = "M" Then
        Me.Caption = "智財顧問專業分配比例"
    ElseIf m_Status = "U" Then
        Me.Caption = "智財顧問專業分配比例-重新計算"
    Else
        Me.Caption = "智財顧問專業分配比例-查詢"
    End If
    
    Call ClearForm
    Call doQuery
    Call TxtLocked

End Sub
 
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If m_Status = "U" Then
       strExc(1) = "": strExc(2) = ""
       For intI = 0 To m_Max
           strExc(1) = strExc(1) & "," & txtData(10 + intI).Text
           strExc(2) = strExc(2) & "," & txtData(10 + intI).Tag
       Next intI
       
       If strExc(1) <> strExc(2) Then
            If MsgBox("重新計算後尚未存檔，是否要存檔及通知相關人？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
                Call CmdSave_Click
                Cancel = True
            End If
       End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If TypeName(m_PrevForm) <> "Nothing" Then
       m_PrevForm.Show
   End If
   
   Set frm081031_3 = Nothing
End Sub

Private Sub ClearForm()

    For Each oObj In txtData
        oObj.Text = "0"
        oObj.Tag = "0"
    Next
    For Each oObj In lblData
       oObj.Caption = ""
    Next
    For Each oObj In lblFM2
       oObj.Caption = ""
    Next
    For Each oObj In lblSys
       oObj.Tag = ""
    Next
    Combo1.Clear
    txtCP15.Text = ""
    txtCP15.Tag = ""
    
End Sub

Private Function FormSave() As Boolean
Dim intP As Integer
Dim tmpUpd As Variant
Dim bolConn As Boolean
On Error GoTo ErrHandle
      
   strSql = ""
   
   '簽約時數
   If txtCP15.Text <> txtCP15.Tag Then
       strSql = strSql & ";" & "update caseprogress set cp15=" & CNULL(txtCP15, True) & " where cp09='" & m_CP09 & "' "
   End If
   
   '收文分配比例
   For intP = 0 To m_Max
      'Modified by Lydia 2023/11/29 人工調整分配比例(調整比例)
      'If txtData(intP).Text & txtData(intP + 10).Text <> txtData(intP).Tag & txtData(intP + 10).Tag Then
      '   If Val(txtData(intP)) = 0 And Val(txtData(intP + 10)) = 0 Then
      '      strSql = strSql & ";" & "delete from ACSPFrate where ar01='" & m_CP09 & "' and ar02='" & lblSys(intP) & "' "
      '   Else
      '      If lblSys(intP).Tag <> "Y" Then
      '         strSql = strSql & ";" & "insert into ACSPFrate (ar01, ar02, ar03, ar04, ar05, ar06, ar07) values (" & _
      '                    CNULL(m_CP09) & ", '" & lblSys(intP).Caption & "', " & CNULL(txtData(intP), True) & ", 0, '" & strUserNum & "', '" & strSrvDate(1) & "', '" & Left(Format(ServerTime, "000000"), 4) & "' ) "
      '      Else
      '         strSql = strSql & ";" & "update ACSPFrate set ar03=" & CNULL(txtData(intP), True) & ", ar05='" & strUserNum & "', ar06=" & strSrvDate(1) & ", ar07=" & Left(Format(ServerTime, "000000"), 4) & _
      '                                   " where ar01='" & m_CP09 & "' and ar02='" & lblSys(intP) & "' "
      '      End If
      '   End If
      'End If
      If txtData(intP).Text & txtData(intP + 10).Text & txtData(intP + 30).Text <> txtData(intP).Tag & txtData(intP + 10).Tag & txtData(intP + 30).Tag Then
         If Val(txtData(intP)) = 0 And Val(txtData(intP + 10)) = 0 And Val(txtData(intP + 10)) = 0 And Val(txtData(intP + 30)) = 0 Then
            strSql = strSql & ";" & "delete from ACSPFrate where ar01='" & m_CP09 & "' and ar02='" & lblSys(intP) & "' "
         Else
            If lblSys(intP).Tag <> "Y" Then
               strSql = strSql & ";" & "insert into ACSPFrate (ar01, ar02, ar03, ar04, ar05, ar06, ar07, ar10) values (" & _
                          CNULL(m_CP09) & ", '" & lblSys(intP).Caption & "', " & CNULL(txtData(intP), True) & ", 0, '" & strUserNum & "', '" & strSrvDate(1) & "', '" & Left(Format(ServerTime, "000000"), 4) & "', " & CNULL(txtData(intP + 30), True) & " ) "
               lblSys(intP).Tag = "Y" 'Added by Lydia 2024/05/15
            Else
               strSql = strSql & ";" & "update ACSPFrate set ar03=" & CNULL(txtData(intP), True) & ", ar10=" & CNULL(txtData(intP + 30), True) & ", ar05='" & strUserNum & "', ar06=" & strSrvDate(1) & ", ar07=" & Left(Format(ServerTime, "000000"), 4) & _
                                         " where ar01='" & m_CP09 & "' and ar02='" & lblSys(intP) & "' "
            End If
         End If
      End If
      'end 2023/11/29
   Next intP
   If strSql <> "" Then
      strSql = Mid(strSql, 2)
      tmpUpd = Split(strSql, ";")
      cnnConnection.BeginTrans
      For intP = 0 To UBound(tmpUpd)
          If Trim(tmpUpd(intP)) <> "" Then
             If InStr(UCase(tmpUpd(intP)), "CASEPROGRESS") > 0 Then '新增log
                 Pub_SeekTbLog Trim(tmpUpd(intP))
                 cnnConnection.Execute "begin user_data.user_enabled:=1; " & Trim(tmpUpd(intP)) & "; end; "
             Else
                 cnnConnection.Execute Trim(tmpUpd(intP))
             End If
          End If
      Next intP
      cnnConnection.CommitTrans
   End If
   FormSave = True
   Exit Function
   
ErrHandle:
   If Err.Number <> 0 Then
        If strSql <> "" Then cnnConnection.RollbackTrans
        MsgBox "存檔失敗：" & vbCrLf & Err.Description, vbCritical
   End If
   
End Function

Private Sub txtCP15_GotFocus()
   TextInverse txtCP15
End Sub

Private Sub txtCP15_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
    TextInverse txtData(Index)
End Sub

Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
Dim intP As Integer

    strExc(1) = "0"
    For intP = 0 To m_Max
        strExc(1) = Val(strExc(1)) + Val(txtData(intP))
    Next intP
   
    If Val(strExc(1)) > 100 Then
        MsgBox "收文分配比例不可大於100%", vbCritical, "檢核資料"
        GoTo EXITSUB
    End If
    
    If Val(txtData(Index)) < 0 Then
        MsgBox "收文分配比例不可輸入負數", vbCritical, "檢核資料"
        GoTo EXITSUB
    End If
    
    If InStr(txtData(Index), ".") > 0 Then
        MsgBox "收文分配比例不可輸入小數點", vbCritical, "檢核資料"
        GoTo EXITSUB
    End If
    
    'Added by Lydia 2023/11/29
    If txtData(30).Locked = False Then
       strExc(1) = "0"
       For intP = 30 To 30 + m_Max
           strExc(1) = Val(strExc(1)) + Val(txtData(intP))
       Next intP
      
       If Val(strExc(1)) > 100 Then
           MsgBox "調整比例不可大於100%", vbCritical, "檢核資料"
           GoTo EXITSUB
       End If
       
       If Val(txtData(Index)) < 0 Then
           MsgBox "調整比例不可輸入負數", vbCritical, "檢核資料"
           GoTo EXITSUB
       End If
    End If
    'end 2023/11/29
    
    Exit Sub
    
EXITSUB:
    Cancel = True
    txtData(Index).SetFocus
End Sub

Private Sub TxtLocked()
Dim bolLock  As Boolean
Dim intJ As Integer

    If m_Status <> "M" Then  'Q-查詢, U-重新計算作業
         bolLock = True
         cmd1.Visible = False
         If m_Status = "U" Then
            CmdCaculate.Visible = True
            cmdSave.Visible = True
            Label1(11).Visible = True
            For intJ = 20 To 20 + m_Max
                txtData(intJ).Visible = True
            Next intJ
         End If
    Else   'M-可修改資料
         bolLock = False
         cmd1.Visible = True
    End If
    
    For Each oObj In txtData
        oObj.Locked = bolLock
    Next
    
    txtCP15.Locked = bolLock
    
    'Added by Lydia 2023/11/29 人工調整分配比例(調整比例)
    If m_Status = "U" Then
       Label1(15).Visible = True
       For intJ = 30 To 30 + m_Max
          txtData(intJ).Locked = False
          txtData(intJ).TabIndex = intJ - 30
       Next intJ
       For intJ = 0 To 0 + m_Max
          txtData(intJ).BorderStyle = False
          txtData(intJ).BackColor = &H80000004
       Next intJ
    Else
       For intJ = 30 To 30 + m_Max
          txtData(intJ).Locked = True
          If m_Status = "Q" Then
             txtData(intJ).Left = txtData(20).Left
          End If
       Next intJ
       If m_Status = "Q" Then
          Label1(15).Left = Label1(11).Left
       End If
    End If
    'end 2023/11/29
End Sub
