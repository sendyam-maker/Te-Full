VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050705_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "被介紹資料"
   ClientHeight    =   4644
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   8292
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   8292
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7380
      TabIndex        =   9
      Top             =   30
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   3300
      Left            =   120
      TabIndex        =   7
      Top             =   1230
      Width           =   8055
      _ExtentX        =   14203
      _ExtentY        =   5821
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
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
      _Band(0).Cols   =   8
   End
   Begin VB.Label Label1 
      Caption         =   "被介紹者資料："
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   8
      Top             =   1020
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "申請人/代理人編號："
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   6
      Top             =   135
      Width           =   1710
   End
   Begin VB.Label Label1 
      Caption         =   "名稱："
      Height          =   180
      Index           =   21
      Left            =   150
      TabIndex        =   5
      Top             =   450
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "國籍："
      Height          =   180
      Index           =   8
      Left            =   3420
      TabIndex        =   4
      Top             =   135
      Width           =   540
   End
   Begin MSForms.Label Lbl1 
      Height          =   285
      Index           =   0
      Left            =   4020
      TabIndex        =   3
      Top             =   135
      Width           =   495
      VariousPropertyBits=   27
      Caption         =   "Lbl1(0)"
      Size            =   "873;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl1 
      Height          =   285
      Index           =   1
      Left            =   4530
      TabIndex        =   2
      Top             =   135
      Width           =   1065
      VariousPropertyBits=   27
      Caption         =   "Lbl1(1)"
      Size            =   "1870;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl1 
      Height          =   570
      Index           =   3
      Left            =   750
      TabIndex        =   1
      Top             =   450
      Width           =   7305
      VariousPropertyBits=   27
      Caption         =   "Lbl1(3)"
      Size            =   "12876;1005"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtNo 
      Height          =   285
      Left            =   1860
      TabIndex        =   0
      Top             =   105
      Width           =   1200
      VariousPropertyBits=   671105051
      BackColor       =   -2147483633
      MaxLength       =   8
      Size            =   "2117;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm050705_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Amy 2022/11/30
Option Explicit

Dim m_PrevForm As Form '前一畫面
Dim RsQ As New ADODB.Recordset, strQ As String, strAllField As String, intQ As Integer
Dim ii As Integer, jj As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim arrF, arrW

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    strAllField = "編號|名稱|國籍|建立日期"
    GRD1.FormatString = strAllField
    arrF = Split(strAllField, "|")
    arrW = Split("1200;4500;1000;1000", ";")
    
    MoveFormToCenter Me
End Sub

Public Sub SetParent(ByRef fm As Form)
    Set m_PrevForm = fm
End Sub

Public Function QueryData() As Boolean
    strQ = "Select XYS01,Nvl(fa04,Decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) as stName,na03,XYS05 " & _
                "From XYNoSource,Fagent,Nation Where XYS02='" & txtNo & "' And XYS01=FA01(+) And FA01 is not null And FA02='0' And FA10=NA01(+) " & _
    "Union Select XYS01,Nvl(cu04,Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)) as stName,na03,XYS05 " & _
                "From XYNoSource,Customer,Nation Where XYS02='" & txtNo & "' And XYS01=CU01(+) And CU01 is not null And CU02='0' And CU10=NA01(+) "
   'Add by Amy 2023/07/12 +國內/外 潛在客戶
   strQ = strQ & "Union " & _
               "Select XYS01,Nvl(poc03,Decode(poc23,null,poc27,poc23||' '||poc24||' '||poc25||' '||poc26)) as stName,na03,XYS05 " & _
                "From XYNoSource,PotCustomer1,Nation Where XYS02='" & txtNo & "' And XYS01=POC01(+) And POC01 is not null And POC02='0' And POC04=NA01(+) " & _
   "Union Select XYS01,Nvl(pcu08,Decode(pcu03,null,pcu07,pcu03||' '||pcu04||' '||pcu05||' '||pcu06)) as stName,na03,XYS05 " & _
                "From XYNoSource,PotCustomer,Nation Where XYS02='" & txtNo & "' And XYS01=PCU01(+) And PCU01 is not null And PCU02='0' And PCU09=NA01(+) " & _
                "Order by XYS01"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        Set GRD1.Recordset = RsQ
    End If
    SetGrdField
End Function

'初始化 Grid
Private Sub SetGrdField()
    For ii = LBound(arrF) To UBound(arrF)
        GRD1.TextMatrix(0, ii) = arrF(ii)
        GRD1.ColWidth(ii) = arrW(ii)
    Next ii
End Sub

Private Function GetValue(pFieldN As String) As Integer
    For jj = LBound(arrF) To UBound(arrF)
        If UCase(arrF(jj)) = UCase(pFieldN) Then
            GetValue = jj
            Exit For
        End If
    Next jj
End Function

Private Sub Form_Unload(Cancel As Integer)
    If TypeName(m_PrevForm) <> "Nothing" Then
        m_PrevForm.Show
    End If
    Set frm050705_1 = Nothing
End Sub
