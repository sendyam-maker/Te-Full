VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090202_2_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "多國案-選取資料"
   ClientHeight    =   5600
   ClientLeft      =   2790
   ClientTop       =   3720
   ClientWidth     =   8120
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5600
   ScaleWidth      =   8120
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "取消"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6750
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   75
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0FF&
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Left            =   5640
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   75
      Width           =   1000
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4725
      Left            =   60
      TabIndex        =   1
      Top             =   810
      Width           =   7995
      _ExtentX        =   14093
      _ExtentY        =   8326
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
   Begin VB.Label Label2 
      Caption         =   "取消：放棄選取並且回前畫面"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   5610
      TabIndex        =   5
      Top             =   540
      Width           =   2445
   End
   Begin VB.Label Label1 
      Caption         =   "可複選或全選　全選：在左上角 V 方框點選即可"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4395
   End
   Begin VB.Label LblNote 
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   120
      TabIndex        =   3
      Top             =   30
      Width           =   5385
   End
End
Attribute VB_Name = "frm090202_2_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/10/7 Form2.0已修改
Option Explicit

Dim i As Integer
Dim m_strQueryType As String
Public m_EEP01 As String '總收文號
Public m_bolCustSend As Boolean '是否已客戶會稿
Dim intKind As Integer


Private Sub cmdok_Click()
   frm090202_2.m_RetrunRecv = "" '回傳總收文號
   frm090202_2.cmdManyCase.Tag = "確定" 'Add By Sindy 2018/10/24
   For i = 1 To grdDataList.Rows - 1
      If grdDataList.TextMatrix(i, 0) = "V" Then
         If frm090202_2.m_RetrunRecv = "" Then
            frm090202_2.m_RetrunRecv = grdDataList.TextMatrix(i, 10)
         Else
            frm090202_2.m_RetrunRecv = frm090202_2.m_RetrunRecv & "," & grdDataList.TextMatrix(i, 10)
         End If
      End If
   Next i
   Unload Me
End Sub

Private Sub Command1_Click()
   frm090202_2.m_RetrunRecv = "" '取消:清除總收文號
   frm090202_2.cmdManyCase.Tag = "" 'Add By Sindy 2020/9/17
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090202_2_1 = Nothing
End Sub

'Modify By Sindy 2023/9/11 +, Optional strCP01 As String
Public Function QueryData(strQueryType As String, Optional strManyAppl As String = "", _
   Optional strEEP05 As String = "", _
   Optional strSales As String = "", Optional strEmp As String = "", Optional strCP05 As String = "", _
   Optional strCP10 As String = "", Optional strCP44 As String = "", Optional strCurrFlowEEP04 As String = "", _
   Optional strCountry As String, Optional strCP01 As String) As Boolean

Dim RsQ As New ADODB.Recordset
Dim strQ As String
Dim strVal As String
Dim arrID As Variant, strCP09 As String
Dim ii As Integer, jj As Integer, kk As Integer
Dim strConSql As String
Dim strCP163 As String, strTM15 As String
   
On Error GoTo ErrHnd
   
   QueryData = False
   m_strQueryType = strQueryType
   If strCP44 <> "" Then strCP44 = ChangeCustomerL(strCP44) 'Add By Sindy 2020/6/8
   
   'Add By Sindy 2023/9/11 檢查系統別是屬那一類
   If strCP01 <> "" Then
      Call ClsPDGetSystemKind(strCP01, intKind)
   End If
   '2023/9/11 END
   
   SetDataListWidth
   LblNote.Visible = False
   
   '多國案-選取資料
   If m_strQueryType = "0" Then
      Me.Caption = "多國案-選取資料"
      strVal = PUB_GetSameCaseSQL(frm090202_2.lblCP09) '相同案語法(收文號)
      '未發文未取消收文
      strQ = "select '' V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,PA05||PA06||PA07 as 案件名稱,NA03 申請國家,Decode(PA09,'000',PTM03,PTM04) as ""類別/種類"",sqldatet(cp05) 收文日,DECODE(PA09,'000',CPM03,CPM04) 案件性質,'' 客戶會稿日,s1.st02 承辦人,s2.st02 智權人員,cp09 總收文號" & _
             " from caseprogress,patent,nation,casepropertyMAP,staff s1,staff s2,PatentTradeMarkMap," & _
             "(" & strVal & ") V1" & _
             " Where substr(V1.CNo, 1, Length(V1.CNo) - 9) = CP01" & _
             " and substr(V1.cno,-9,6)=cp02" & _
             " and substr(V1.cno,-3,1)=cp03" & _
             " and substr(V1.cno,-2)=cp04" & _
             " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
             " and na01(+)=pa09 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
             " and cp14=s1.st01(+) and cp13=s2.st01(+) AND cp158=0 AND cp159=0" & _
             " and cp01||'-'||cp02||'-'||cp03||'-'||cp04<>'" & frm090202_2.lblCaseNo & "'" & _
             " and '1'=PTM01(+) AND PA08=PTM02(+)"
   
   '請勾選需整批客戶會稿案件(限A類收文)
   ElseIf m_strQueryType = EMP_客戶會稿 Then
      Me.Caption = "請勾選需整批客戶會稿案件(限A類收文)"
      Me.LblNote.Caption = "Email客戶會稿時，一律帶入「送會」附件，" & vbCrLf & _
                           "若有修改需求，請直接於附件區增刪處理"
      'select EEP01,max(EEP02) as EEP02 from(
      'Union select null,null from dual where 1=0
      If m_bolCustSend = True Then '已客戶會稿
         strVal = "select EEP01,max(EEP02) as EEP02 from EmpElectronProcess e3" & _
                  " where EEP09='Y'" & _
                  " And EEP04='" & EMP_送會 & "' and substr(EEP01,1,1)='A'" & _
                  " And exists(select e2.eep01 from EmpElectronProcess e2 where e2.eep01=e3.eep01 and e2.eep06||lpad(e2.eep07,6,'0')>e3.eep06||lpad(e3.eep07,6,'0') and e2.eep04='" & EMP_客戶會稿 & "')" & _
                  " And EEP05='" & strEEP05 & "'" & _
                  " group by EEP01"
      Else '未客戶會稿
         '" And not exists(select e2.eep01 from EmpElectronProcess e2 where e2.eep01=e3.eep01 and e2.eep06<e3.eep06 and e2.eep04='" & EMP_客戶會稿 & "')"
         strVal = "select EEP01,max(EEP02) as EEP02 from EmpElectronProcess e3" & _
                  " where EEP09='Y'" & _
                  " And EEP04='" & EMP_送會 & "' and substr(EEP01,1,1)='A'" & _
                  " And not exists(select e2.eep01 from EmpElectronProcess e2 where e2.eep01=e3.eep01 and e2.eep06||lpad(e2.eep07,6,'0')>e3.eep06||lpad(e3.eep07,6,'0') and e2.eep04='" & EMP_客戶會稿 & "')" & _
                  " And EEP05='" & strEEP05 & "'" & _
                  " group by EEP01"
      End If
      
      If intKind = 商標 Then
         'Modify By Sindy 2024/8/13 cp14 => ep05
         strQ = "select '' V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,TM05||TM06||TM07 as 案件名稱,NA03 申請國家,TM09 as ""類別/種類"",sqldatet(cp05) 收文日,DECODE(TM10,'000',CPM03,CPM04) 案件性質,sqldatet(EP37) 客戶會稿日,s1.st02 承辦人,s2.st02 智權人員,cp09 總收文號" & _
                " from EmpElectronProcess e1,caseprogress,trademark,nation,casepropertyMAP,staff s1,staff s2,engineerprogress" & _
                " Where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
                " AND e1.EEP01=CP09(+) AND e1.EEP01=EP02(+)" & _
                " and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04" & _
                " and na01(+)=tm10 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
                " and cp13='" & strSales & "'" & _
                " and ep05=s1.st01(+) and cp13=s2.st01(+) AND cp158=0 AND cp159=0"
         'Modify By Sindy 2023/9/5 轉讓案應抓讓與申請人 ex:T-245082,T-202355 14筆
         If strCP10 = "501" Then
            strQ = strQ & " and (instr('" & strManyAppl & "',substr(cp56,1,8))>0 or instr('" & strManyAppl & "',substr(cp89,1,8))>0 or instr('" & strManyAppl & "',substr(cp90,1,8))>0 or instr('" & strManyAppl & "',substr(cp91,1,8))>0 or instr('" & strManyAppl & "',substr(cp92,1,8))>0)"
         Else
         '2023/9/5 END
            strQ = strQ & " and (instr('" & strManyAppl & "',substr(tm23,1,8))>0 or instr('" & strManyAppl & "',substr(tm78,1,8))>0 or instr('" & strManyAppl & "',substr(tm79,1,8))>0 or instr('" & strManyAppl & "',substr(tm80,1,8))>0 or instr('" & strManyAppl & "',substr(tm81,1,8))>0)"
         End If
         
      'Modify By Sindy 2023/9/11 +TM案
      Else
      '2023/9/11 END
         'Modify By Sindy 2024/8/13 cp14 => ep05
         strQ = "select '' V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,sp05||sp06||sp07 as 案件名稱,NA03 申請國家,sp73 as ""類別/種類"",sqldatet(cp05) 收文日,DECODE(sp09,'000',CPM03,CPM04) 案件性質,sqldatet(EP37) 客戶會稿日,s1.st02 承辦人,s2.st02 智權人員,cp09 總收文號" & _
                " from EmpElectronProcess e1,caseprogress,servicepractice,nation,casepropertyMAP,staff s1,staff s2,engineerprogress" & _
                " Where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
                " AND e1.EEP01=CP09(+) AND e1.EEP01=EP02(+)" & _
                " and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04" & _
                " and na01(+)=sp09 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
                " and cp13='" & strSales & "'" & _
                " and ep05=s1.st01(+) and cp13=s2.st01(+) AND cp158=0 AND cp159=0"
         If strCP10 = "501" Then
            strQ = strQ & " and (instr('" & strManyAppl & "',substr(cp56,1,8))>0 or instr('" & strManyAppl & "',substr(cp89,1,8))>0 or instr('" & strManyAppl & "',substr(cp90,1,8))>0 or instr('" & strManyAppl & "',substr(cp91,1,8))>0 or instr('" & strManyAppl & "',substr(cp92,1,8))>0)"
         Else
            strQ = strQ & " and (instr('" & strManyAppl & "',substr(sp08,1,8))>0 or instr('" & strManyAppl & "',substr(sp58,1,8))>0 or instr('" & strManyAppl & "',substr(sp59,1,8))>0 or instr('" & strManyAppl & "',substr(sp65,1,8))>0 or instr('" & strManyAppl & "',substr(sp66,1,8))>0)"
         End If
      End If
      
   '同一客戶會稿案件-整批會完(限A類收文)
   ElseIf m_strQueryType = EMP_會完 Then
      Me.Caption = "同一客戶會稿案件-整批會完(A類收文限有客戶會稿)"
      Me.LblNote.Caption = "注意：執行整批會完時，請確認案件均未修改內容" & vbCrLf & _
                           "　　　若有修改，請以單筆會完處理"
      'Add By Sindy 2023/5/10
      strCP163 = ""
      strSql = "select cp09,cp163 from caseprogress where cp09='" & m_EEP01 & "' and cp163 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strCP163 = RsTemp.Fields("cp163")
      End If
      '2023/5/10 END
      
      'Modify By Sindy 2023/5/9 A類收文限有客戶會稿,加C類
      strVal = "select EEP01,max(EEP02) as EEP02 from EmpElectronProcess e3" & _
               " where EEP09='Y'" & _
               " And EEP04='" & EMP_送會 & "'" & _
               " and ((substr(EEP01,1,1)='A' And exists(select e2.eep01 from EmpElectronProcess e2 where e2.eep01=e3.eep01 and e2.eep06>=e3.eep06 and e2.eep04='" & EMP_客戶會稿 & "'))" & _
                    " or substr(EEP01,1,1)='C')" & _
               " And EEP05='" & strEEP05 & "'" & _
               " group by EEP01"
      
      If intKind = 商標 Then
         'Modify By Sindy 2023/5/9 + IIf(strCP163 <> "", " and cp163='" & strCP163 & "' and cp163<>cp09", "")
         'Modify By Sindy 2024/6/17 + and cp14='" & strEmp & "' ex:T-249005
         'Modify By Sindy 2024/8/13 cp14 => ep05
         strQ = "select '' V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,TM05||TM06||TM07 as 案件名稱,NA03 申請國家,TM09 as ""類別/種類"",sqldatet(cp05) 收文日,DECODE(TM10,'000',CPM03,CPM04) 案件性質,sqldatet(EP37) 客戶會稿日,s1.st02 承辦人,s2.st02 智權人員,cp09 總收文號" & _
                " from EmpElectronProcess e1,caseprogress,trademark,nation,casepropertyMAP,staff s1,staff s2,engineerprogress" & _
                " Where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
                " AND e1.EEP01=CP09(+) AND e1.EEP01=EP02(+)" & _
                " and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04" & _
                " and na01(+)=tm10 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
                " and cp13='" & strSales & "' and ep05='" & strEmp & "'" & _
                " and ep05=s1.st01(+) and cp13=s2.st01(+) AND cp158=0 AND cp159=0" & IIf(strCP163 <> "", " and cp163='" & strCP163 & "' and cp163<>cp09", "")
         'Modify By Sindy 2023/9/5 轉讓案應抓讓與申請人 ex:T-245082,T-202355 14筆
         If strCP10 = "501" Then
            strQ = strQ & " and (instr('" & strManyAppl & "',substr(cp56,1,8))>0 or instr('" & strManyAppl & "',substr(cp89,1,8))>0 or instr('" & strManyAppl & "',substr(cp90,1,8))>0 or instr('" & strManyAppl & "',substr(cp91,1,8))>0 or instr('" & strManyAppl & "',substr(cp92,1,8))>0)"
         Else
         '2023/9/5 END
            strQ = strQ & " and (instr('" & strManyAppl & "',substr(tm23,1,8))>0 or instr('" & strManyAppl & "',substr(tm78,1,8))>0 or instr('" & strManyAppl & "',substr(tm79,1,8))>0 or instr('" & strManyAppl & "',substr(tm80,1,8))>0 or instr('" & strManyAppl & "',substr(tm81,1,8))>0)"
         End If
         
      'Modify By Sindy 2023/9/11 +TM案
      Else
      '2023/9/11 END
         'Modify By Sindy 2024/6/17 + and cp14='" & strEmp & "' ex:T-249005
         'Modify By Sindy 2024/8/13 cp14 => ep05
         strQ = "select '' V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,sp05||sp06||sp07 as 案件名稱,NA03 申請國家,sp73 as ""類別/種類"",sqldatet(cp05) 收文日,DECODE(sp09,'000',CPM03,CPM04) 案件性質,sqldatet(EP37) 客戶會稿日,s1.st02 承辦人,s2.st02 智權人員,cp09 總收文號" & _
                " from EmpElectronProcess e1,caseprogress,servicepractice,nation,casepropertyMAP,staff s1,staff s2,engineerprogress" & _
                " Where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
                " AND e1.EEP01=CP09(+) AND e1.EEP01=EP02(+)" & _
                " and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04" & _
                " and na01(+)=sp09 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
                " and cp13='" & strSales & "' and ep05='" & strEmp & "'" & _
                " and ep05=s1.st01(+) and cp13=s2.st01(+) AND cp158=0 AND cp159=0" & IIf(strCP163 <> "", " and cp163='" & strCP163 & "' and cp163<>cp09", "")
         If strCP10 = "501" Then
            strQ = strQ & " and (instr('" & strManyAppl & "',substr(cp56,1,8))>0 or instr('" & strManyAppl & "',substr(cp89,1,8))>0 or instr('" & strManyAppl & "',substr(cp90,1,8))>0 or instr('" & strManyAppl & "',substr(cp91,1,8))>0 or instr('" & strManyAppl & "',substr(cp92,1,8))>0)"
         Else
            strQ = strQ & " and (instr('" & strManyAppl & "',substr(sp08,1,8))>0 or instr('" & strManyAppl & "',substr(sp58,1,8))>0 or instr('" & strManyAppl & "',substr(sp59,1,8))>0 or instr('" & strManyAppl & "',substr(sp65,1,8))>0 or instr('" & strManyAppl & "',substr(sp66,1,8))>0)"
         End If
      End If
      
   'Add By Sindy 2020/9/25 其他歷程的多案
   Else
      Me.LblNote.Caption = "多案操作歷程：歷程只記錄在操作的案號中。" & vbCrLf & _
                           "【若為同一" & IIf(Left(m_EEP01, 1) = "A", "指示信", "定稿") & "內容時，才需勾選合併處理的案件。】" & vbCrLf & vbCrLf
      If strCountry <> "000" And _
         (Left(m_EEP01, 1) = "A" Or Left(m_EEP01, 1) = "B") Then
         Me.Caption = "指示信(限非台灣案AB類收文)"
         If m_strQueryType = EMP_送件 Or m_strQueryType = EMP_退件重送 Then
            Me.LblNote.Caption = Me.LblNote.Caption & _
                                 "發指示信時，一律帶入「送件」PDF附件，" & vbCrLf & _
                                 "該筆文號若有修改電子檔的需求，請直接於附件區增刪處理"
         End If
         strConSql = strConSql & " AND cp44='" & strCP44 & "'"
         If strCP10 = "102" Or strCP10 = "301" Then
            strConSql = strConSql & " AND cp10 in('102','301')"
         Else
            strConSql = strConSql & " AND cp10='" & strCP10 & "'"
         End If
         
      'Add By Sindy 2022/4/26 ex:尚待收款-完稿日
      ElseIf frm090202_2.m_strSpecState = "尚待收款-完稿日" And m_strQueryType = EMP_聯絡 Then
         Me.Caption = "台灣案(尚待收款-完稿日)"
         strConSql = strConSql & " AND cp10='" & strCP10 & "' AND cp141='2' AND cp79>0 AND cp13='" & strSales & "'"
      
      'Modify By Sindy 2024/11/28 本國客戶
      '+ And (frm090202_2.bolTMFlow = True Or frm090202_2.bolCFTFlow = True)
      ElseIf Left(m_EEP01, 1) = "C" _
         And (frm090202_2.bolTMFlow = True Or frm090202_2.bolCFTFlow = True) Then
         Me.Caption = "來函同申請人案件"
         strConSql = strConSql & " AND cp10='" & strCP10 & "'"
      
      'Modify By Sindy 2024/8/23 + frm090202_2.bolTMFlow = True
      ElseIf frm090202_2.bolTMFlow = True And strCountry = "000" And _
         (strCP10 = "308" Or strCP10 = "301" Or strCP10 = "501") Then
         Me.Caption = "台灣案(限分割、移轉、變更AB類收文)"
         strConSql = strConSql & " AND cp10='" & strCP10 & "'"
      
      'Add By Sindy 2024/8/23
      ElseIf frm090202_2.bolFCTFlow = True And strCountry = "000" And _
         (strCP10 = "301" Or strCP10 = "501") Then
         Me.Caption = "台灣案(限變更、移轉收文)"
         strConSql = strConSql & " AND cp10='" & strCP10 & "'"
      Else
         '不可操作多案
         Exit Function
      End If
      
      If m_strQueryType = EMP_送會 Then
         strConSql = strConSql & " AND cp13='" & strSales & "'"
      End If
      
      '承辦進度要一致
      'Add By Sindy 2022/4/26 排除 尚待收款-完稿日
      If frm090202_2.m_strSpecState = "尚待收款-完稿日" And m_strQueryType = EMP_聯絡 Then
         strConSql = strConSql & " AND (ep02='" & m_EEP01 & "' or ep09 is null)"
      Else
      '2022/4/26 END
         If RsQ.State = 1 Then RsQ.Close
         strQ = "select * from engineerprogress where ep02='" & m_EEP01 & "'"
         RsQ.CursorLocation = adUseClient
         RsQ.Open strQ, cnnConnection, adOpenDynamic, adLockBatchOptimistic
         If RsQ.RecordCount = 1 Then
            If Val("" & RsQ.Fields("ep07")) > 0 Then
               strConSql = strConSql & " And ep07>0"
            Else
               strConSql = strConSql & " And nvl(ep07,0)=0"
            End If
            If Val("" & RsQ.Fields("ep08")) > 0 Then
               strConSql = strConSql & " And ep08>0"
            Else
               strConSql = strConSql & " And nvl(ep08,0)=0"
            End If
            If Val("" & RsQ.Fields("ep09")) > 0 Then
               strConSql = strConSql & " And ep09>0"
            Else
               strConSql = strConSql & " And nvl(ep09,0)=0"
            End If
         End If
      End If
      '無進行中的歷程
      'Modify By Sindy 2025/1/22 排除已送件
      strConSql = strConSql & " And not exists(select e2.eep01 from EmpElectronProcess e2 where e2.eep01=cp09 and e2.eep09='Y')" & _
                              " And ((instr(GetEEPCurState(cp09),'送件')=0" & _
                                    " And instr(GetEEPCurState(cp09),'退件重送')=0" & _
                                    " And instr(GetEEPCurState(cp09),'發文歸檔')=0) or GetEEPCurState(cp09) is null)"
      If intKind = 商標 Then
         'Add By Sindy 2025/11/7 台灣案的變更有分註冊前和註冊後,主要為了區分申請書的不同
         strTM15 = ""
         strSql = "select cp09,TM15 from caseprogress,trademark where cp09='" & m_EEP01 & "'" & _
                  " and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strTM15 = "" & RsTemp.Fields("TM15")
         End If
         If strCountry = "000" And strCP10 = "301" Then '台灣變更案
            If Trim(strTM15) = "" Then '註冊號數
               strConSql = strConSql & " AND TM15 is null"
            Else
               strConSql = strConSql & " AND TM15 is not null"
            End If
         End If
         '2025/11/7 END
      
         '文件必須齊備
         'Modify By Sindy 2024/8/13 cp14 => ep05
         strQ = "select '' V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,TM05||TM06||TM07 as 案件名稱,NA03 申請國家,TM09 as ""類別/種類"",sqldatet(cp05) 收文日,DECODE(TM10,'000',CPM03,CPM04) 案件性質,sqldatet(EP37) 客戶會稿日,s1.st02 承辦人,s2.st02 智權人員,cp09 總收文號" & _
                " FROM caseprogress,trademark,nation,casepropertyMAP,staff s1,staff s2,engineerprogress" & _
                " Where cp09=EP02(+) AND cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04" & _
                " AND tm10='" & strCountry & "'" & _
                " AND na01(+)=tm10 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
                " AND ep05='" & strEmp & "' AND ep05=s1.st01(+) AND cp13=s2.st01(+)" & _
                " AND cp158=0 AND cp159=0 AND ep06>0" & _
                strConSql
         'Modify By Sindy 2023/9/5 轉讓案應抓讓與申請人 ex:T-245082,T-202355 14筆
         If strCP10 = "501" Then
            strQ = strQ & " AND (instr('" & Left(strManyAppl, 9) & "',substr(cp56,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(cp89,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(cp90,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(cp91,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(cp92,1,8))>0)"
         Else
         '2023/9/5 END
            strQ = strQ & " AND (instr('" & Left(strManyAppl, 9) & "',substr(tm23,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(tm78,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(tm79,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(tm80,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(tm81,1,8))>0)"
         End If
         
      'Modify By Sindy 2023/9/11 +TM案 (intKind=6)
      Else
      '2023/9/11 END
         'Modify By Sindy 2024/8/13 cp14 => ep05
         strQ = "select '' V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,sp05||sp06||sp07 as 案件名稱,NA03 申請國家,sp73 as ""類別/種類"",sqldatet(cp05) 收文日,DECODE(sp09,'000',CPM03,CPM04) 案件性質,sqldatet(EP37) 客戶會稿日,s1.st02 承辦人,s2.st02 智權人員,cp09 總收文號" & _
                " FROM caseprogress,servicepractice,nation,casepropertyMAP,staff s1,staff s2,engineerprogress" & _
                " Where cp09=EP02(+) AND cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04" & _
                " AND sp09='" & strCountry & "'" & _
                " AND na01(+)=sp09 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
                " AND ep05='" & strEmp & "' AND ep05=s1.st01(+) AND cp13=s2.st01(+)" & _
                " AND cp158=0 AND cp159=0 AND ep06>0" & _
                strConSql
         If strCP10 = "501" Then
            strQ = strQ & " AND (instr('" & Left(strManyAppl, 9) & "',substr(cp56,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(cp89,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(cp90,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(cp91,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(cp92,1,8))>0)"
         Else
            strQ = strQ & " AND (instr('" & Left(strManyAppl, 9) & "',substr(sp08,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(sp58,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(sp59,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(sp65,1,8))>0 or instr('" & Left(strManyAppl, 9) & "',substr(sp66,1,8))>0)"
         End If
      End If
      
   End If
   If Trim(Me.LblNote.Caption) <> "" Then Me.LblNote.Visible = True '*****
   If RsQ.State = 1 Then RsQ.Close
   RsQ.CursorLocation = adUseClient
   RsQ.Open strQ, cnnConnection, adOpenDynamic, adLockBatchOptimistic
   If (m_strQueryType = "0" And RsQ.RecordCount > 0) Or _
      (m_strQueryType <> "0" And RsQ.RecordCount > 1) Then
      QueryData = True
      Set grdDataList.Recordset = RsQ
      grdDataList.Visible = False
      If frm090202_2.m_RetrunRecv <> "" Then
         arrID = Split(frm090202_2.m_RetrunRecv, ",")
         For ii = 0 To UBound(arrID)
            strCP09 = arrID(ii)
            For jj = 1 To grdDataList.Rows - 1
               If grdDataList.TextMatrix(jj, 10) = strCP09 Then
                  grdDataList.TextMatrix(jj, 0) = "V"
                  grdDataList.row = jj
                  For kk = 0 To grdDataList.Cols - 1
                     grdDataList.col = kk
                     grdDataList.CellBackColor = &HFFC0C0
                  Next kk
               End If
            Next jj
         Next ii
      End If
      grdDataList.Visible = True
   End If
   RsQ.Close
   Set RsQ = Nothing
   
   Exit Function
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   Set RsQ = Nothing
End Function

Private Sub GrdDataList_Click()
Dim j As Integer
Dim ii As Integer
   
   grdDataList.col = 0
   grdDataList.row = grdDataList.MouseRow
   If m_EEP01 <> "" Then
      If grdDataList.TextMatrix(grdDataList.row, 10) = m_EEP01 Then Exit Sub '總收文號
   End If
   
   grdDataList.Visible = False
   If grdDataList.row <> 0 Then
      If grdDataList.TextMatrix(grdDataList.row, 0) = "V" Then
         grdDataList.TextMatrix(grdDataList.row, 0) = ""
         grdDataList.row = grdDataList.row
         For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = QBColor(15)
         Next j
      Else
         grdDataList.TextMatrix(grdDataList.row, 0) = "V"
         grdDataList.row = grdDataList.row
         For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = &HFFC0C0
         Next j
      End If
   'Add By Sindy 2018/10/22
   ElseIf grdDataList.MouseCol = 0 Then
      If grdDataList.Tag = "V" Then '全部取消
         For ii = 1 To grdDataList.Rows - 1
            If grdDataList.TextMatrix(ii, 10) <> m_EEP01 Then
               grdDataList.TextMatrix(ii, 0) = ""
               grdDataList.row = ii
               For j = 0 To grdDataList.Cols - 1
                  grdDataList.col = j
                  grdDataList.CellBackColor = QBColor(15)
               Next j
            End If
         Next ii
         grdDataList.Tag = ""
      Else '全選
         For ii = 1 To grdDataList.Rows - 1
            If grdDataList.TextMatrix(ii, 10) <> m_EEP01 Then
               grdDataList.TextMatrix(ii, 0) = "V"
               grdDataList.row = ii
               For j = 0 To grdDataList.Cols - 1
                  grdDataList.col = j
                  grdDataList.CellBackColor = &HFFC0C0
               Next j
            End If
         Next ii
         grdDataList.Tag = "V"
      End If
      '2018/10/22 END
   End If
   grdDataList.Visible = True
End Sub

Private Sub SetDataListWidth()
Dim iCol As Integer
   
   ReDim strColN(12)
   ReDim intWidth(12)
'   If m_strQueryType = EMP_送件 Or m_strQueryType = EMP_退件重送 Then 'TM指示信
'      '                0    1           2           3           4            5         6           7         8         9           10          11
'      strColN = Array("V", "本所案號", "案件名稱", "申請國家", "類別/種類", "收文日", "案件性質", "會稿日", "承辦人", "智權人員", "總收文號", "發文日")
'      intWidth = Array(200, 1200, 700, 700, 500, 800, 700, 800, 700, 700, 800, 800)
'   Else
      '                0    1           2           3           4            5         6           7         8         9           10
      strColN = Array("V", "本所案號", "案件名稱", "申請國家", "類別/種類", "收文日", "案件性質", "會稿日", "承辦人", "智權人員", "總收文號")
      If m_strQueryType <> "0" Then
         intWidth = Array(200, 1200, 700, 700, 500, 800, 700, 800, 700, 700, 800)
      Else
         intWidth = Array(200, 1200, 700, 700, 500, 800, 700, 0, 700, 700, 800)
      End If
'   End If
   With grdDataList
      .Visible = False
      For iCol = 0 To UBound(strColN)
         .ColWidth(iCol) = intWidth(iCol)
         .TextMatrix(0, iCol) = strColN(iCol)
      Next
      .Visible = True
   End With
End Sub
