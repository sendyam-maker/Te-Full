VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm12040149 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "MAPISession1"
   ClientHeight    =   4548
   ClientLeft      =   3048
   ClientTop       =   1512
   ClientWidth     =   7800
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4548
   ScaleWidth      =   7800
   Begin VB.CommandButton Command2 
      Caption         =   "FCP³]­p±M¥Î´Á©µªø³qª¾"
      Height          =   465
      Left            =   3300
      TabIndex        =   19
      Top             =   60
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FCP·s¬KÀu´f³qª¾"
      Height          =   435
      Left            =   4995
      TabIndex        =   18
      Top             =   60
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   588
      Left            =   45
      TabIndex        =   15
      Top             =   2850
      Width           =   7635
   End
   Begin VB.Timer Timer1 
      Left            =   45
      Top             =   30
   End
   Begin VB.Frame Frame1 
      Caption         =   "¥[³t¼f¬d³qª¾¨ç"
      Height          =   1395
      Left            =   90
      TabIndex        =   1
      Top             =   570
      Width           =   7575
      Begin VB.TextBox txtMailTo 
         Height          =   315
         Left            =   3960
         TabIndex        =   13
         Top             =   870
         Width           =   2040
      End
      Begin VB.CommandButton cmdMail 
         Caption         =   "µo°eMail"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6210
         TabIndex        =   11
         Top             =   840
         Width           =   1230
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   3960
         MaxLength       =   9
         TabIndex        =   10
         Top             =   480
         Width           =   2040
      End
      Begin VB.Frame Frame3 
         Caption         =   "®æ¦¡"
         Height          =   975
         Left            =   1530
         TabIndex        =   6
         Top             =   270
         Width           =   1230
         Begin VB.OptionButton Option2 
            Caption         =   "EMail"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   8
            Top             =   300
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton Option2 
            Caption         =   "¶Ç¯u"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   7
            Top             =   630
            Width           =   960
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "»y¤å"
         Height          =   975
         Left            =   180
         TabIndex        =   3
         Top             =   270
         Width           =   1230
         Begin VB.OptionButton Option1 
            Caption         =   "¤é¤å"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   5
            Top             =   630
            Width           =   960
         End
         Begin VB.OptionButton Option1 
            Caption         =   "­^¤å"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   4
            Top             =   300
            Value           =   -1  'True
            Width           =   960
         End
      End
      Begin VB.CommandButton cmdLetter 
         Caption         =   "²£¥ÍWord"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6210
         TabIndex        =   2
         Top             =   420
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "¦¬¥ó¤H"
         Height          =   180
         Index           =   1
         Left            =   2925
         TabIndex        =   12
         Top             =   900
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "¥N²z¤H½s¸¹"
         Height          =   180
         Index           =   0
         Left            =   2925
         TabIndex        =   9
         Top             =   540
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "µ²§ô"
      Height          =   435
      Left            =   6885
      TabIndex        =   0
      Top             =   60
      Width           =   825
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   540
      Top             =   60
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   345
      Left            =   45
      TabIndex        =   14
      Top             =   2160
      Width           =   7665
      _ExtentX        =   13526
      _ExtentY        =   614
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1668
      Top             =   0
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   1032
      Top             =   0
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "¯S®í«H¨ç¹w³]¤£§PÂ_¹q¤l³øÄæ¦ì¡I"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   180
      TabIndex        =   17
      Top             =   3990
      Width           =   7215
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  '¸m¤¤¹ï»ô
      Caption         =   "( 0/0 )"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      TabIndex        =   16
      Top             =   2550
      Width           =   2490
   End
End
Attribute VB_Name = "frm12040149"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/5 §ï¦¨Form2.0 (µL)
'Memo By Sonia 2012/12/6 ´¼Åv¤H­ûÄæ¤w­×§ï
'2010/12/1 memo by sonia ­û¤u½s¸¹Äæ¤w­×§ï
'sonia 2010/8/19 ¤é´ÁÄæ¤w­×§ï
'Create by Morgan 2009/2/26
Option Explicit

Const MailBefore$ = "IMCEAEX-_O=TAIE_OU=DOMAIN_CN=RECIPIENTS_CN="
Const MailAfter$ = "@taie.com.tw"
Const MailName$ = "Tai E International Patent & Law Office"

Dim Result$, Sec%
Dim fso As New FileSystemObject

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdLetter_Click()
   Screen.MousePointer = vbHourglass
   Process
   Screen.MousePointer = vbDefault
End Sub

Private Sub Process()
   Dim ErrFA() As String, iUB As Integer, bolErr As Boolean
   Dim stKey As String, stLstKey As String, iSNo As Integer, stCP10 As String
   Dim stET02 As String, strLstCuNo As String, iLine As Integer
   Dim stAppList As String, stAppCaseList As String, iListRows As Integer
   Dim iExPages As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stCP(1 To 4) As String
   Dim stTag As String
   Dim stVTB1 As String, stVTB2 As String, stVTB3 As String, stVTB4 As String
   Dim bJap As Boolean
   Dim stCon As String, stPA As String
   Dim stFileName As String
   Dim bByEmail As Boolean
   Dim bSkip As Boolean
   
   ReDim ErrFA(0)
   
    stCon = ""
    stPA = ""
    
   '¥N²z¤H
   If Text1 <> "" Then
      stPA = stPA & " and pa75='" & Text1 & "'"
      If Option2(0).Value = True Then
         bByEmail = True
      Else
         bByEmail = False
      End If
   Else
      '­^¤å
      If Option1(0).Value = True Then
         stCon = stCon & " and (fa31 is null or fa31<>'3')"
         'EMail
         If Option2(0).Value = True Then
            stCon = stCon & " and instr(fa16,'@')>0"
            bByEmail = True
         '¶Ç¯u
         ElseIf Option2(1).Value = True Then
            stCon = stCon & " and (fa16 is null or instr(fa16,'@')=0)"
            bByEmail = False
         End If
      '¤é¤å(¦³³y¦r°ÝÃD¬G¤@²v¥Î¶Ç¯u)
      ElseIf Option1(1).Value = True Then
         stCon = stCon & " and fa31='3'"
         bByEmail = False
      End If
   End If

   
   'ÂÂªk
   stVTB1 = "select pa01,pa02,pa03,pa04,pa75,pa26,pa77,pa48,pa11,pa22" & _
      " From patent" & _
      " where pa01='FCP' and pa08='1' and pa57 is null and pa108 is null and pa10<20021026 and (pa16 is null or pa16='2')" & stPA & _
      " and not exists(select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10 in ('422','501','413','929') and cp57 is null)" & _
      " and not exists(select * from nextprogress where np02=pa01 and np03=pa02 and np04=pa03 and np05=pa04 and np07 in ('107','205') and (np06 is null or np06='N'))"
      
   
   stVTB2 = " union select pa01,pa02,pa03,pa04,pa75,pa26,pa77,pa48,pa11,pa22" & _
      " From patent" & _
      " where pa01='FCP' and pa08='1' and pa57 is null and pa108 is null and pa10>=20021026 and (pa16 is null or pa16='2')" & stPA & _
      " and exists(select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='416' and cp27>0)" & _
      " and not exists(select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10 in ('422','501','413','929') and cp57 is null)" & _
      " and not exists(select * from nextprogress where np02=pa01 and np03=pa02 and np04=pa03 and np05=pa04 and np07 in ('107','205') and (np06 is null or np06='N'))"
   
     'Add by Lydia 2014/11/14 FCP©Ó¿ì°Ï°ì¯S®íª¬ªp¤§´¼Åv¤H­û¹º¤À¤è¦¡
    '¥N²z¤HY51333010=Pub_GetSpecMan("¥_¨Ê»ÈÀsFCP®×©Ó¿ì·~°È")
'   strExc(0) = "select *" & _
      " from (" & stVTB1 & stVTB2 & stVTB3 & stVTB4 & ") X,fagent,Nation,CUSTOMER,staff " & _
      " where fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9)" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
      " and na01(+)=fa10 and st01(+)=na51" & stCon & _
      " order by na16,na01,fa05,pa75,cu05,pa26,pa01,pa02,pa03,pa04"
   'Modified by Lydia 2024/05/27 §ï¦¨¥ÎY½s¸¹+¤é¤å©Ó¿ì·~°È
   'strExc(0) = "select *" & _
      " from (" & stVTB1 & stVTB2 & stVTB3 & stVTB4 & ") X,fagent,Nation,CUSTOMER,staff " & _
      " where fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9)" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
      " and na01(+)=fa10 and st01(+)=decode(pa75,'Y51333010','" & Pub_GetSpecMan("¥_¨Ê»ÈÀsFCP®×©Ó¿ì·~°È") & "',na51)" & stCon & _
      " order by na16,na01,fa05,pa75,cu05,pa26,pa01,pa02,pa03,pa04"
   strExc(0) = "select *" & _
      " from (" & stVTB1 & stVTB2 & stVTB3 & stVTB4 & ") X,fagent,Nation,CUSTOMER,staff " & _
      " where fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9)" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
      " and na01(+)=fa10 and st01(+)=decode(pa75," & Pub_GetSpecFCP & ",na51)" & stCon & _
      " order by na16,na01,fa05,pa75,cu05,pa26,pa01,pa02,pa03,pa04"
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With rsQuery
      .MoveFirst
      iSNo = 0
      iLine = 0
      iListRows = 0
      iExPages = 0
      strLstCuNo = ""
      stAppList = ""
      stAppCaseList = ""
      
      stKey = ""
      stLstKey = "" & .Fields("pa75")
      stCP(1) = .Fields("pa01").Value
      stCP(2) = .Fields("pa02").Value
      stCP(3) = .Fields("pa03").Value
      stCP(4) = .Fields("pa04").Value
      '§ï¨Ì´¼Åv¤H­û¥N½X©R¦W
      stFileName = .Fields("st17") & "-" & stLstKey & ".doc"
      stFileName = Replace(stFileName, "/", "-")
      stET02 = stCP(1) & stCP(2) & stCP(3) & stCP(4) & "&000"
      If .Fields("fa31") = "3" Then
         bJap = True
      Else
         bJap = False
      End If
      bSkip = True
      Do While Not .EOF
'         If bSkip = True Then
'            If .Fields("na16") = "89030" And .Fields("pa75") = "Y51553000" Then
'               iSNo = 0
'               iLine = 0
'               iListRows = 0
'               iExPages = 0
'               strLstCuNo = ""
'               stAppList = ""
'               stAppCaseList = ""
'
'               stKEY = ""
'               stLstKey = "" & .Fields("pa75")
'               stCP(1) = .Fields("pa01").Value
'               stCP(2) = .Fields("pa02").Value
'               stCP(3) = .Fields("pa03").Value
'               stCP(4) = .Fields("pa04").Value
'               stFileName = .Fields("na16") & "_" & stLstKey & ".doc"
'               stET02 = stCP(1) & stCP(2) & stCP(3) & stCP(4) & "&000"
'               If .Fields("fa31") = "3" Then
'                  bJap = True
'               Else
'                  bJap = False
'               End If
'               bSkip = False
'            End If
'            If bSkip = True Then GoTo NextStep
'         End If
         
         stKey = "" & .Fields("pa75")
         If stKey <> stLstKey Then
            If bByEmail = True Then
               bolErr = Not PrintLetter(stET02, bJap, stAppList, iLine, stAppCaseList, iExPages, stFileName)
            Else
               pub_AddressListSN = pub_AddressListSN + 1
               PUB_AddNewAddressList strUserNum, stCP(1), stCP(2), stCP(3), stCP(4), "" & pub_AddressListSN, "0"
               If PrintLetter(stET02, bJap, stAppList, iLine, stAppCaseList, iExPages, stFileName, False, True) = False Then
                  bolErr = True
               ElseIf PrintLetter(stET02, bJap, stAppList, iLine, stAppCaseList, iExPages, , False, False) = False Then
                  bolErr = True
               Else
                  bolErr = False
               End If
            End If
            
            If bolErr = True Then
               iUB = iUB + 1
               ReDim Preserve ErrFA(iUB)
               ErrFA(iUB) = stLstKey
               MsgBox "[" & stLstKey & "]²£¥Í«H¨ç¿ù»~¡I"
            End If
            
            iSNo = 0
            iLine = 0
            iListRows = 0
            iExPages = 0
            strLstCuNo = ""
            stAppList = ""
            stAppCaseList = ""
            stLstKey = stKey
            stCP(1) = .Fields("pa01").Value
            stCP(2) = .Fields("pa02").Value
            stCP(3) = .Fields("pa03").Value
            stCP(4) = .Fields("pa04").Value
            '§ï¨Ì´¼Åv¤H­û¥N½X©R¦W
            stFileName = .Fields("st17") & "-" & stLstKey & ".doc"
            stFileName = Replace(stFileName, "/", "-")
            stET02 = stCP(1) & stCP(2) & stCP(3) & stCP(4) & "&000"
            If .Fields("fa31") = "3" Then
               bJap = True
            Else
               bJap = False
            End If
         End If
         
         strExc(1) = .Fields("pa01") & "-" & .Fields("pa02") & IIf(.Fields("pa03") & .Fields("pa04") = "000", "", "-" & .Fields("pa03") & "-" & .Fields("pa04"))
         
         If "" & .Fields("pa26") <> strLstCuNo Then
            iSNo = iSNo + 1

            If iSNo > 1 Then
               stAppList = stAppList & vbCrLf & Space(14)
            End If

            '®×¥ó²M³æ¦æ¼Æ¶W¹L­n³sªíÀY¤@°_¸õ­¶
            If iListRows > 50 Then
               stAppCaseList = stAppCaseList & vbCrLf & Chr(12)
               iListRows = 0
               iExPages = iExPages + 1
            ElseIf iSNo > 1 Then
               stAppCaseList = stAppCaseList & vbCrLf & vbCrLf
               iListRows = iListRows + 1
            End If
            
            If bJap = True Then
               stAppList = stAppList & .Fields("cu06")
               stAppCaseList = stAppCaseList & "¥XÄ@¤H¡G" & .Fields("cu06")
               iLine = iLine + 1
               iListRows = iListRows + 1
            Else
               If Not IsNull(.Fields("cu05")) Then
                  stAppList = stAppList & .Fields("cu05") & " " & .Fields("cu88")
                  stAppCaseList = stAppCaseList & "Applicant: " & .Fields("cu05") & " " & .Fields("cu88")
                  iLine = iLine + 1
                  iListRows = iListRows + 1
                  If Not IsNull(.Fields("cu89")) Then
                     stAppList = stAppList & vbCrLf & Space(14) & .Fields("cu89") & " " & .Fields("cu90")
                     iLine = iLine + 1
                     stAppCaseList = stAppCaseList & vbCrLf & Space(11) & .Fields("cu89") & " " & .Fields("cu90")
                     iListRows = iListRows + 1
                  End If
               Else
                  stAppList = stAppList & .Fields("cu06")
                  stAppCaseList = stAppCaseList & "Applicant: " & .Fields("cu06")
                  iLine = iLine + 1
                  iListRows = iListRows + 1
               End If
            End If
            
            If bJap = True Then
               stAppCaseList = stAppCaseList & vbCrLf & vbCrLf & _
                  "¶Q©Ò¡]ªÀ¡^¾ã²zµf†A             þÝ«È¼Ë¾ã²zµf†A                 ’U©Ò¾ã²zµf†A    ¥XÄ@µf†A"
            Else
               stAppCaseList = stAppCaseList & vbCrLf & vbCrLf & _
                  "Your Ref                       Case No.                       Our Ref         Application No."
            End If
            stAppCaseList = stAppCaseList & vbCrLf & _
               "----------------------------------------------------------------------------------------------" & vbCrLf & _
               Left(.Fields("pa77") & Space(31), 31) & Left(.Fields("pa48") & Space(31), 31) & Left(strExc(1) & Space(16), 16) & Left(.Fields("pa11") & Space(18), 18)
            iListRows = iListRows + 4
            strLstCuNo = "" & .Fields("pa26")
            
         Else
            If iListRows > 50 Then
               stAppCaseList = stAppCaseList & vbCrLf & Chr(12)
               If bJap = True Then
                  stAppCaseList = stAppCaseList & vbCrLf & vbCrLf & _
                     "¶Q©Ò¡]ªÀ¡^¾ã²zµf†A             þÝ«È¼Ë¾ã²zµf†A                 ’U©Ò¾ã²zµf†A    ¥XÄ@µf†A"
               Else
                  stAppCaseList = stAppCaseList & vbCrLf & vbCrLf & _
                     "Your Ref                       Case No.                       Our Ref         Application No."
               End If
               stAppCaseList = stAppCaseList & vbCrLf & _
               "----------------------------------------------------------------------------------------------"

               iListRows = 2
               iExPages = iExPages + 1
            End If
            stAppCaseList = stAppCaseList & vbCrLf & _
               Left(.Fields("pa77") & Space(31), 31) & Left(.Fields("pa48") & Space(31), 31) & Left(strExc(1) & Space(16), 16) & Left(.Fields("pa11") & Space(18), 18)

            iListRows = iListRows + 1
         End If
Nextstep:
         .MoveNext
      Loop
      
      If bByEmail = True Then
         bolErr = Not PrintLetter(stET02, bJap, stAppList, iLine, stAppCaseList, iExPages, stFileName)
      Else
         pub_AddressListSN = pub_AddressListSN + 1
         PUB_AddNewAddressList strUserNum, stCP(1), stCP(2), stCP(3), stCP(4), "" & pub_AddressListSN, "0"
         If PrintLetter(stET02, bJap, stAppList, iLine, stAppCaseList, iExPages, stFileName, False, True) = False Then
            bolErr = True
         ElseIf PrintLetter(stET02, bJap, stAppList, iLine, stAppCaseList, iExPages, , False, False) = False Then
            bolErr = True
         Else
            bolErr = False
         End If
      End If
      
      If bolErr = True Then
         iUB = iUB + 1
         ReDim Preserve ErrFA(iUB)
         ErrFA(iUB) = stLstKey
         MsgBox "[" & stLstKey & "]²£¥Í«H¨ç¿ù»~¡I"
      End If
      
      End With
      
      If UBound(ErrFA) > 0 Then
RePrint:
         Printer.Print "«H¨ç¿ù»~²M³æ"
         Printer.Print Join(ErrFA, vbCrLf)
         Printer.EndDoc
         If MsgBox("¬O§_­«¦L«H¨ç¿ù»~²M³æ¡H", vbYesNo + vbDefaultButton1) = vbYes Then
            GoTo RePrint
         End If
      End If
   Else
      MsgBox "µL¸ê®Æ¡I"
   End If
   Set rsQuery = Nothing
End Sub




Private Function PrintLetter1(ET02 As String, p_CP10 As String, p_stAppList As String, p_iLineCnt As Integer, p_stAppCaseList As String, Optional p_iExtPages As Integer, Optional p_stStatus As String) As Boolean
   Dim ET01 As String, ET03(5) As String, intJ As Integer
   Dim stExpFld(2) As String
   Dim bolRetry As Boolean '¬O§_¤wµo¥Í¿ù»~¥B­«¸Õ
   Dim stLetter(6) As String
   Dim stFileName As String '¼È¦s¹ÏÀÉÀÉ¦W
   Dim iPicNo As Integer '¹ÏÀÉ¥N½X 1:¥~°Ó 2:¥~±M/¥~ªk 3.CFP 4.¨ä¥L
   'Added by Lydia 2016/09/29
   Dim oShape
   Dim oShape2
   
   ET01 = "04"
   Erase ET03
   ET03(1) = "97" '¶Ç¯u«Ê­±
   ET03(3) = "91" 'ªþ¥ó1(©e¥ô®Ñ)
   iPicNo = 2
   
   Select Case p_CP10
      Case "928" '­«·s©e¥ô
         Select Case p_stStatus
            Case "A1" '¥¼­ã
               stExpFld(1) = "Taiwan Patent Application(s)"
               ET03(2) = "88"
            Case "A2" '¤w­ã
               stExpFld(1) = "Taiwan Granted Patent(s)"
               ET03(2) = "87"
            Case "A3" '¦h¥Ó½Ð¤H
               stExpFld(1) = "Taiwan Patent Application(s) and Patent(s)"
               ET03(2) = "86"
         End Select
         stExpFld(2) = 4 + p_iExtPages
         
      Case "701" 'Åý»P
         stExpFld(1) = "Assignment on Taiwan Patent Application(s) and Patent(s)"
         stExpFld(2) = 6 + p_iExtPages
         ET03(2) = "85"
         ET03(4) = "90"
         ET03(5) = "89"
         
      Case "703" 'Ä~©Ó
         stExpFld(1) = "Taiwan Patent Application(s) and Patent(s)"
         stExpFld(2) = 5 + p_iExtPages
         ET03(2) = "84"
         ET03(4) = "90"

      Case "702" '¦X¨Ö
         stExpFld(1) = "Merger on Taiwan Patent Application(s) and Patent(s)"
         stExpFld(2) = 5 + p_iExtPages
         ET03(2) = "83"
         ET03(4) = "90"
         
      Case "401" '§ó¦W
         stExpFld(1) = "Name Change on Taiwan Patent Application(s) and Patent(s)"
         stExpFld(2) = 5 + p_iExtPages
         ET03(2) = "82"
         ET03(4) = "90"
   End Select
   
   EndLetter ET01, ET02, ET03(1), strUserNum
   
   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03(1) & "','" & strUserNum & _
      "','¥DÃD','" & stExpFld(1) & "')"

   cnnConnection.Execute strSql
   
   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03(1) & "','" & strUserNum & _
      "','¶Ç¯u­¶¼Æ','" & stExpFld(2) & "')"
            
   cnnConnection.Execute strSql
   
   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03(1) & "','" & strUserNum & _
      "','¥Ó½Ð¤H²M³æ','" & ChgSQL(p_stAppList) & "')"
            
   cnnConnection.Execute strSql
   
   
   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03(1) & "','" & strUserNum & _
      "','¼u©Ê¸õ¦æ','" & String(15 - p_iLineCnt, vbCrLf) & "')"
            
   cnnConnection.Execute strSql
   
   Erase stLetter
   
   For intJ = 1 To 5
      If ET03(intJ) <> Empty Then
         NowPrint ET02, "04", ET03(intJ), False, strUserNum, , , True, stLetter(intJ), 1
      End If
   Next
   stLetter(6) = p_stAppCaseList
   
   If stLetter(1) = "" Then
      MsgBox "©w½ZÅª¨ú¥¢±Ñ¡I"
      Exit Function
   End If
   
   If ReadDB2File(stFileName, iPicNo) = False Then
      MsgBox "­¶­º¹ÏÀÉÅª¨ú¥¢±Ñ¡I"
      Exit Function
   End If
   bolRetry = False
    
On Error GoTo ERRORSECTION1
   
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   g_WordAp.Documents.add
   With g_WordAp
   
      '.Visible = True
      '.WindowState = wdWindowStateMaximize

      '³]©w¦r«¬ª©­±(°Ñ·Ó©w½Z)
      .Selection.Font.Name = "Times New Roman"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(3.125)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(3.125)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      
      '´¡¤J¹Ï¤ùÀÉ®×(²Ä¤@­¶)
      .Selection.HomeKey Unit:=wdStory
       'Modified by Lydia 2016/09/29 ¥ÎÂÂ¼gªk·|³y¦¨Word2010¥X¿ù
      '.ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True
      '.ActiveDocument.Shapes("Picture " & Trim(.ActiveDocument.Shapes.Count + 1)).Select
      Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
      oShape.Select
      'Modified by Lydia 2016/09/29 .Selection.ShapeRange=> oShape
      oShape.ZOrder 4
      oShape.LockAnchor = True
      oShape.LockAspectRatio = -1
      oShape.Width = 546.5
      oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
      oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
      oShape.Left = .CentimetersToPoints(1)
      oShape.Top = .CentimetersToPoints(1)
      'end 2016/09/29
      
      '²Ä¤@­¶¤º®e
      .Selection.EndKey Unit:=wdStory
      .Selection.Font.Size = 12
      .Selection.TypeText stLetter(1)
      
      '´¡¤J¹Ï¤ùÀÉ®×(²Ä¤G­¶)
      '.Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Count:=1
      .Selection.TypeText vbCrLf & Chr(12)
      'Modified by Lydia 2016/09/29 ¥ÎÂÂ¼gªk·|³y¦¨Word2010¥X¿ù
      '.ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True
      '.ActiveDocument.Shapes("Picture " & Trim(.ActiveDocument.Shapes.Count + 1)).Select
      Set oShape2 = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
      oShape2.Select
      'Modified by Lydia 2016/09/29 .Selection.ShapeRange=> oShape2
      oShape2.ZOrder 4
      oShape2.LockAnchor = True
      oShape2.LockAspectRatio = -1
      oShape2.Width = 546.5
      oShape2.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
      oShape2.RelativeVerticalPosition = wdRelativeVerticalPositionPage
      oShape2.Left = .CentimetersToPoints(1)
      oShape2.Top = .CentimetersToPoints(1)
      'end 2016/09/29
      
      .Selection.EndKey Unit:=wdStory
      
      '²Ä¤G­¶¤º®e
      .Selection.EndKey Unit:=wdStory
      .Selection.TypeText stLetter(2)
      
      '­«³]¤WÃä¬É
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(2.5)
      .Selection.PageSetup.SectionStart = wdSectionNewPage
      
      '²Ä¤T­¶¤º®e
      .Selection.TypeText vbCrLf & Chr(12)
      .Selection.Font.Size = 14
      .Selection.TypeText stLetter(3)
      
      stLetter(0) = stLetter(1) & vbCrLf & Chr(12) & stLetter(2) & vbCrLf & Chr(12) & stLetter(3)
      
      '²Ä¥|­¶¤º®e
      If stLetter(4) <> "" Then
         .Selection.TypeText vbCrLf & Chr(12)
         .Selection.Font.Size = 13
         .Selection.TypeText stLetter(4)
         stLetter(0) = stLetter(0) & vbCrLf & Chr(12) & stLetter(4)
      End If
      
      '²Ä¤­­¶¤º®e
      If stLetter(5) <> "" Then
         .Selection.TypeText vbCrLf & Chr(12)
         .Selection.Font.Size = 12
         .Selection.TypeText stLetter(5)
         stLetter(0) = stLetter(0) & vbCrLf & Chr(12) & stLetter(5)
      End If
      
      ChgWordFormat g_WordAp.Application, stLetter(0)
      .Selection.EndKey Unit:=wdStory
      
      '²Ä¤»­¶¤º®e
      '§ï¾î¦L
      
      .ActiveDocument.Range(Start:=.Selection.Start, End:=.Selection.Start).InsertBreak Type:=wdSectionBreakNextPage
      .Selection.Start = .Selection.Start + 1
      With .ActiveDocument.Range(Start:=.Selection.Start, End:=.ActiveDocument.Content.End).PageSetup
         .Orientation = wdOrientLandscape
      End With
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.Font.Size = 12
      .Selection.Font.Bold = True
      .Selection.Font.Name = "²Ó©úÅé"
      .Selection.TypeText stLetter(6)
            
      '¦C¦L
      'Modify by Morgan 2008/1/23 ¾ã§å¤w§¹¦¨§ïºûÅ@
      .PrintOut Copies:=1, Collate:=True: DoEvents
      '' ²M°£¤å¥ó¤º®e
      '.Selection.WholeStory
      '.Selection.Delete
      '.ActiveDocument.Close savechanges:=wdDoNotSaveChanges
      '.Visible = True
   End With
   
   PrintLetter1 = True
   
ERRORSECTION1:
   If Err.Number <> 0 Then
      Select Case Err.Number
         Case 91, 462:
            If bolRetry = True Then
               MsgBox Err.Description, vbCritical
            Else
               Set g_WordAp = New Word.Application
               g_WordAp.Documents.add
               bolRetry = True
               Resume
            End If
         Case Else:
            MsgBox Err.Description, vbCritical
      End Select
   End If
End Function

Private Function PrintLetter(ET02 As String, p_bolJap As Boolean, p_stAppList As String, p_iLineCnt As Integer, p_stAppCaseList As String, Optional p_iExtPages As Integer, Optional p_FileName As String, Optional p_byEmail As Boolean = True, Optional p_FaxPage As Boolean = False) As Boolean
   Dim ET01 As String, ET03(5) As String, intJ As Integer
   Dim stExpFld(2) As String
   Dim bolRetry As Boolean '¬O§_¤wµo¥Í¿ù»~¥B­«¸Õ
   Dim stLetter(6) As String
   Dim stFileName As String '¼È¦s¹ÏÀÉÀÉ¦W
   Dim iPicNo As Integer '¹ÏÀÉ¥N½X 1:¥~°Ó 2:¥~±M/¥~ªk 3.CFP 4.¨ä¥L
   Dim stLetterPath As String, stCoverPath As String
   Dim stDocPath As String
   Dim stET03 As String
   'Added by Lydia 2016/09/29
   Dim oShape
   Dim oShape2
   
   ET01 = "04"
   Erase ET03
   
   If p_FileName <> "" Then
      stDocPath = PUB_Getdesktop & "\AE"
      If p_bolJap = True Then
         stDocPath = stDocPath & "\JAP"
      Else
         stDocPath = stDocPath & "\ENG"
      End If
      If p_byEmail = True Then
         stDocPath = stDocPath & "\EMail"
      Else
         stDocPath = stDocPath & "\Fax"
      End If
      stDocPath = stDocPath & "\" & p_FileName
   Else
      stDocPath = ""
   End If
   
   '¤é¤å«H¨ç
   If p_bolJap Then
      If p_FaxPage = True Then ET03(1) = "78"
      ET03(2) = "80"
      stLetterPath = PUB_Getdesktop & "\Jap.doc"
      stCoverPath = PUB_Getdesktop & "\JapCover.doc"
   '­^¤å«H¨ç
   Else
      '¶Ç¯u«Ê­±
      If p_FaxPage = True Then ET03(1) = "97"
      If p_byEmail = True Then
         ET03(2) = "79"
      Else
         ET03(2) = "81"
      End If
      stLetterPath = PUB_Getdesktop & "\Eng.doc"
   End If
   
   iPicNo = 2
   
   stExpFld(1) = "Accelerated Examination in Taiwan"
   stExpFld(2) = 4 + p_iExtPages
   If ET03(1) <> "" Then
      stET03 = ET03(1)
   Else
      stET03 = ET03(2)
   End If
   
   EndLetter ET01, ET02, stET03, strUserNum
   
   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & stET03 & "','" & strUserNum & _
      "','¶Ç¯u­¶¼Æ','" & stExpFld(2) & "')"
            
   cnnConnection.Execute strSql
   
   If p_bolJap = False Then
   
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & stET03 & "','" & strUserNum & _
         "','¥DÃD','" & stExpFld(1) & "')"
   
      cnnConnection.Execute strSql
      
      
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & stET03 & "','" & strUserNum & _
         "','¥Ó½Ð¤H²M³æ','" & ChgSQL(p_stAppList) & " ')"
               
      cnnConnection.Execute strSql
      
      
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & stET03 & "','" & strUserNum & _
         "','¼u©Ê¸õ¦æ','" & String(15 - p_iLineCnt, vbCrLf) & "')"
               
      cnnConnection.Execute strSql
      
   End If
   
   Erase stLetter
   
   For intJ = 1 To 5
      If ET03(intJ) <> Empty Then
         NowPrint ET02, "04", ET03(intJ), False, strUserNum, , , True, stLetter(intJ), 1
      End If
   Next
   stLetter(6) = p_stAppCaseList
   
   If stLetter(2) = "" Then
      MsgBox "©w½ZÅª¨ú¥¢±Ñ¡I"
      Exit Function
   End If
   
   If ReadDB2File(stFileName, iPicNo) = False Then
      MsgBox "­¶­º¹ÏÀÉÅª¨ú¥¢±Ñ¡I"
      Exit Function
   End If
   bolRetry = False
   
   
On Error GoTo ERRORSECTION1
   
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   g_WordAp.Documents.add
   With g_WordAp
   
      '.Visible = True
      '.WindowState = wdWindowStateMaximize

      '³]©w¦r«¬ª©­±(°Ñ·Ó©w½Z)
      .Selection.Font.Name = "Times New Roman"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      If p_bolJap Then
         .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
         .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
      Else
         .Selection.PageSetup.LeftMargin = .CentimetersToPoints(3.125)
         .Selection.PageSetup.RightMargin = .CentimetersToPoints(3.125)
      End If
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      
      '´¡¤J¹Ï¤ùÀÉ®×(²Ä¤@­¶)
      .Selection.HomeKey Unit:=wdStory
      'Modified by Lydia 2016/09/29 ¥ÎÂÂ¼gªk·|³y¦¨Word2010¥X¿ù
      '.ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True
      '.ActiveDocument.Shapes("Picture " & Trim(.ActiveDocument.Shapes.Count + 1)).Select
      Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
      oShape.Select
      'Modified by Lydia 2016/09/29 .Selection.ShapeRange =>oShape
      oShape.ZOrder 4
      oShape.LockAnchor = True
      oShape.LockAspectRatio = -1
      oShape.Width = 546.5
      oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
      oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
      oShape.Left = .CentimetersToPoints(1)
      oShape.Top = .CentimetersToPoints(1)
      'end 2016/09/29
      
      If stLetter(1) <> "" Then
         '²Ä¤@­¶¤º®e
         .Selection.EndKey Unit:=wdStory
         .Selection.Font.Size = 12
         .Selection.TypeText stLetter(1)
         
         If stCoverPath <> "" Then
            .Selection.EndKey Unit:=wdStory
            .Selection.InsertFile FileName:=stCoverPath, Range:="", ConfirmConversions:= _
              False, Link:=False, Attachment:=False
            .Selection.EndKey Unit:=wdStory
         End If
         
         
         '´¡¤J¹Ï¤ùÀÉ®×(²Ä¤G­¶)
         '.Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Count:=1
         .Selection.TypeText vbCrLf & Chr(12)
         'Modified by Lydia 2016/09/29 ¥ÎÂÂ¼gªk·|³y¦¨Word2010¥X¿ù
         '.ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True
         '.ActiveDocument.Shapes("Picture " & Trim(.ActiveDocument.Shapes.Count + 1)).Select
         Set oShape2 = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
         oShape2.Select
         'Modified by Lydia 2016/09/29 .Selection.ShapeRange =>oShape2
         oShape2.ZOrder 4
         oShape2.LockAnchor = True
         oShape2.LockAspectRatio = -1
         oShape2.Width = 546.5
         oShape2.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
         oShape2.RelativeVerticalPosition = wdRelativeVerticalPositionPage
         oShape2.Left = .CentimetersToPoints(1)
         oShape2.Top = .CentimetersToPoints(1)
         'end 2016/09/29
         
         stLetter(0) = stLetter(1) & vbCrLf & Chr(12) & stLetter(2)
      Else
         stLetter(0) = stLetter(2)
      End If
      
      '²Ä¤G­¶¤º®e
      .Selection.EndKey Unit:=wdStory
      .Selection.TypeText stLetter(2)
      
      ChgWordFormat g_WordAp.Application, stLetter(0)
      
      .Selection.EndKey Unit:=wdStory
      .Selection.InsertFile FileName:=stLetterPath, Range:="", ConfirmConversions:= _
        False, Link:=False, Attachment:=False
      .Selection.EndKey Unit:=wdStory
      
      'ªþ¥ó¤º®e
      .ActiveDocument.Range(Start:=.Selection.Start, End:=.Selection.Start).InsertBreak Type:=wdSectionBreakNextPage
      .Selection.Start = .Selection.Start + 1
      With .ActiveDocument.Range(Start:=.Selection.Start, End:=.ActiveDocument.Content.End).PageSetup
         '.Orientation = wdOrientLandscape
         .LeftMargin = g_WordAp.CentimetersToPoints(2)
         .RightMargin = g_WordAp.CentimetersToPoints(2)
      End With
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.Font.Size = 10
      .Selection.Font.Bold = True
      .Selection.Font.Name = "²Ó©úÅé"
      .Selection.TypeText stLetter(6)
            
      '¦C¦L
      'Modify by Morgan 2008/1/23 ¾ã§å¤w§¹¦¨§ïºûÅ@
      '.PrintOut Copies:=1, Collate:=True: DoEvents
      ' ²M°£¤å¥ó¤º®e
      '.Selection.WholeStory
      '.Selection.Delete
      If stDocPath <> "" Then
         .ActiveDocument.SaveAs stDocPath
      End If
      'If p_byEmail = False Then
      '   .PrintOut Copies:=1, Collate:=True: DoEvents
      'End If
      
      Sleep 1
      .ActiveDocument.Close savechanges:=wdDoNotSaveChanges
      '.Visible = True
   End With
   
   PrintLetter = True
   
ERRORSECTION1:
   If Err.Number <> 0 Then
      Select Case Err.Number
         Case 91, 462:
            If bolRetry = True Then
               MsgBox Err.Description, vbCritical
            Else
               Set g_WordAp = New Word.Application
               g_WordAp.Documents.add
               bolRetry = True
               Resume
            End If
         Case Else:
            MsgBox Err.Description, vbCritical
      End Select
   End If
End Function
'±q¸ê®Æ®wÅª¥XÀÉ®×
Private Function ReadDB2File(ByRef p_FileName As String, Optional p_iPicNo As Integer = 1) As Boolean

   Dim iFileNo As Integer
   Dim bytes() As Byte
   p_FileName = ""
   
On Error GoTo ErrHnd

   strSql = "select * from ImgByteFile where ibf01='M51' and ibf02='" & Format(p_iPicNo, "00000#") & "'"
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         'Add By Sindy 2017/8/10
'         If "" & .Fields("IBF15") <> "" Then
            ReadDB2File = PUB_GetFtpFile(.Fields("IBF15"), App.path & "\TempFile", UCase("ImgByteFile"))
'         Else
'         '2017/8/10 END
'            ReDim bytes(Val(.Fields("ibf13").Value))
'            bytes() = .Fields("ibf14").GetChunk(Val(.Fields("ibf13").Value))
'            iFileNo = FreeFile
'            Open App.path & "\TempFile" For Binary Access Write As #iFileNo
'            Put #iFileNo, , bytes()
'            Close #iFileNo
'            ReadDB2File = True
'         End If
         p_FileName = App.path & "\TempFile"
      End If
   End With
   Exit Function
   
ErrHnd:
   MsgBox Err.Description
End Function

Private Sub Process1()
   Dim ErrFA() As String, iUB As Integer, bolErr As Boolean
   Dim stKey As String, stLstKey As String, iSNo As Integer, stCP10 As String
   Dim stET02 As String, strLstCuNo As String, iLine As Integer
   Dim stAppList As String, stAppCaseList As String, iListRows As Integer
   Dim iExPages As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stCP(1 To 4) As String
   Dim stTag As String
   Dim stVTB1 As String, stVTB2 As String, stVTB3 As String, stVTB4 As String
   
   
   ReDim ErrFA(0)
   
   'strExc(0) = "select na01,pa75,fa05,pa26,nvl(cu05,nvl(cu06,cu04)) cu05,cu88,cu89,cu90" & _
      ",np02,np03,np04,np05,nvl(b.cp10,a.cp10) cp10,nvl(b.cp09,a.cp09) cp09" & _
      ",pa77 ,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) OurRef,pa11,pa22,pa48" & _
      ",decode(b.cp10,null,decode(pa27,null,decode(nvl(pa16,'2'),'2',0,1),2),3) TAG" & _
      " from nextprogress,caseprogress a,patent,fagent,nation,customer,caseprogress b" & _
      " where NP02='FCP' and np06 is null and np07='202'" & _
      " and a.cp09(+)=np01 and a.cp10='928' and a.cp27>19221111 and a.cp57 is null" & _
      " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04 and pa57 is null and pa108 is null" & _
      " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9,1)" & _
      " and na01(+)=fa10" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
      " and b.cp43(+)=a.cp09 AND (B.CP10 IS NULL OR B.CP10 IN ('401','701','702','703'))" & _
      " and pa27 is not null and B.CP10 is null" & _
      " order by na01,fa05,pa75,decode(b.cp10,null,decode(pa27,null,decode(nvl(pa16,'2'),'2',0,1),2),3),nvl(cu05,nvl(cu06,cu04)),pa26,np03"
   
   stVTB1 = "select na01,na03,pa75,fa05,pa26,nvl(cu05,nvl(cu06,cu04)) cu05" & _
      ",cu88,cu89,cu90,np02,np03,np04,np05,nvl(b.cp10,a.cp10) cp10,nvl(b.cp09,a.cp09) cp09" & _
      ",pa77,pa11,pa22,pa48,na16" & _
      ",decode(b.cp10,null,decode(pa27,null,decode(nvl(pa16,'2'),'2','A1','A2'),'A3'),'B'||decode(B.cp10,'701','Åý','702','¦X','703','Ä~','401','§ó')) TAG" & _
      " from nextprogress,caseprogress a,patent,fagent,nation,customer,caseprogress b" & _
      " where NP02='FCP' and np06 is null and np07='202'" & _
      " and a.cp09(+)=np01 and a.cp10='928' and a.cp27>19221111 and a.cp57 is null" & _
      " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04 and pa57 is null and pa108 is null" & _
      " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9,1)" & _
      " and na01(+)=fa10" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
      " and b.cp43(+)=a.cp09 AND (B.CP10 IS NULL OR B.CP10 IN ('401','701','702','703'))" & _
      " and pa75 is not null and not(b.cp10 is not null and pa27 is not null)"
      
   stVTB2 = " union all select na01,na03,pa75,fa05,pa27,nvl(cu05,nvl(cu06,cu04)) cu05" & _
      ",cu88,cu89,cu90,np02,np03,np04,np05,nvl(b.cp10,a.cp10) cp10,nvl(b.cp09,a.cp09) cp09" & _
      ",pa77,pa11,pa22,pa48,na16" & _
      ",decode(b.cp10,null,decode(pa27,null,decode(nvl(pa16,'2'),'2','A1','A2'),'A3'),'B'||decode(B.cp10,'701','Åý','702','¦X','703','Ä~','401','§ó')) TAG" & _
      " from nextprogress,caseprogress a,patent,fagent,nation,customer,caseprogress b" & _
      " where NP02='FCP' and np06 is null and np07='202'" & _
      " and a.cp09(+)=np01 and a.cp10='928' and a.cp27>19221111 and a.cp57 is null" & _
      " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04 and pa57 is null and pa108 is null" & _
      " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9,1)" & _
      " and na01(+)=fa10" & _
      " and cu01(+)=substr(pa27,1,8) and cu02(+)=substr(pa27,9,1)" & _
      " and b.cp43(+)=a.cp09 AND (B.CP10 IS NULL OR B.CP10 IN ('401','701','702','703'))" & _
      " and pa75 is not null and not(b.cp10 is not null and pa27 is not null)" & _
      " and pa27 is not null"
      
   stVTB3 = " union all select na01,na03,pa75,fa05,pa28,nvl(cu05,nvl(cu06,cu04)) cu05" & _
      ",cu88,cu89,cu90,np02,np03,np04,np05,nvl(b.cp10,a.cp10) cp10,nvl(b.cp09,a.cp09) cp09" & _
      ",pa77,pa11,pa22,pa48,na16" & _
      ",decode(b.cp10,null,decode(pa27,null,decode(nvl(pa16,'2'),'2','A1','A2'),'A3'),'B'||decode(B.cp10,'701','Åý','702','¦X','703','Ä~','401','§ó')) TAG" & _
      " from nextprogress,caseprogress a,patent,fagent,nation,customer,caseprogress b" & _
      " where NP02='FCP' and np06 is null and np07='202'" & _
      " and a.cp09(+)=np01 and a.cp10='928' and a.cp27>19221111 and a.cp57 is null" & _
      " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04 and pa57 is null and pa108 is null" & _
      " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9,1)" & _
      " and na01(+)=fa10" & _
      " and cu01(+)=substr(pa28,1,8) and cu02(+)=substr(pa28,9,1)" & _
      " and b.cp43(+)=a.cp09 AND (B.CP10 IS NULL OR B.CP10 IN ('401','701','702','703'))" & _
      " and pa75 is not null and not(b.cp10 is not null and pa27 is not null)" & _
      " and pa28 is not null"
      
   stVTB4 = " union all select na01,na03,pa75,fa05,pa29,nvl(cu05,nvl(cu06,cu04)) cu05" & _
      ",cu88,cu89,cu90,np02,np03,np04,np05,nvl(b.cp10,a.cp10) cp10,nvl(b.cp09,a.cp09) cp09" & _
      ",pa77,pa11,pa22,pa48,na16" & _
      ",decode(b.cp10,null,decode(pa27,null,decode(nvl(pa16,'2'),'2','A1','A2'),'A3'),'B'||decode(B.cp10,'701','Åý','702','¦X','703','Ä~','401','§ó')) TAG" & _
      " from nextprogress,caseprogress a,patent,fagent,nation,customer,caseprogress b" & _
      " where NP02='FCP' and np06 is null and np07='202'" & _
      " and a.cp09(+)=np01 and a.cp10='928' and a.cp27>19221111 and a.cp57 is null" & _
      " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04 and pa57 is null and pa108 is null" & _
      " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9,1)" & _
      " and na01(+)=fa10" & _
      " and cu01(+)=substr(pa29,1,8) and cu02(+)=substr(pa29,9,1)" & _
      " and b.cp43(+)=a.cp09 AND (B.CP10 IS NULL OR B.CP10 IN ('401','701','702','703'))" & _
      " and pa75 is not null and not(b.cp10 is not null and pa27 is not null)" & _
      " and pa29 is not null"
      
   strExc(0) = "select na01,pa75,fa05,pa26,cu05,cu88,cu89,cu90,np02,np03,np04,np05,cp10,cp09,pa77" & _
      ",np02||'-'||np03||decode(np04||np05,'000','','-'||np04||'-'||np04) OurRef,pa11,pa22,pa48,TAG,na16" & _
      " from (" & stVTB1 & stVTB2 & stVTB3 & stVTB4 & ") where na16='86013' and na01='207' and pa75='Y20267000' and TAG='A2'" & _
      " order by na16,na01,fa05,pa75,TAG,cu05,pa26,np02,np03,np04,np05"
   
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With rsQuery
      .MoveFirst
      iSNo = 0
      iLine = 0
      iListRows = 0
      iExPages = 0
      strLstCuNo = ""
      stAppList = ""
      stAppCaseList = ""
      
      stKey = ""
      stTag = .Fields("TAG")
      stLstKey = "" & .Fields("pa75") & .Fields("cp10") & .Fields("TAG")
      stET02 = "" & .Fields("cp09")
      stCP10 = "" & .Fields("cp10")
      stCP(1) = .Fields("np02").Value
      stCP(2) = .Fields("np03").Value
      stCP(3) = .Fields("np04").Value
      stCP(4) = .Fields("np05").Value
      Do While Not .EOF
         bolErr = False
         stKey = "" & .Fields("pa75") & .Fields("cp10") & .Fields("TAG")
         If stKey <> stLstKey Then
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewAddressList strUserNum, stCP(1), stCP(2), stCP(3), stCP(4), "" & pub_AddressListSN, "0"
            If PrintLetter1(stET02, stCP10, stAppList, iLine, stAppCaseList, iExPages, stTag) = False Then
               bolErr = True
            End If
            If bolErr = True Then
               iUB = iUB + 1
               ReDim Preserve ErrFA(iUB)
               ErrFA(iUB) = stLstKey
               MsgBox "[" & stLstKey & "]²£¥Í«H¨ç¿ù»~¡I"
            End If
            iSNo = 0
            iLine = 0
            iListRows = 0
            iExPages = 0
            strLstCuNo = ""
            stET02 = "" & .Fields("cp09")
            stCP10 = "" & .Fields("cp10")
            stAppList = ""
            stAppCaseList = ""
            stLstKey = stKey
            stTag = .Fields("TAG")
            stCP(1) = .Fields("np02").Value
            stCP(2) = .Fields("np03").Value
            stCP(3) = .Fields("np04").Value
            stCP(4) = .Fields("np05").Value
            
         End If
            
         If "" & .Fields("pa26") <> strLstCuNo Then
            iSNo = iSNo + 1

            If iSNo > 1 Then
               stAppList = stAppList & vbCrLf & Space(14)
            End If

            '®×¥ó²M³æ¦æ¼Æ¶W¹L­n³sªíÀY¤@°_¸õ­¶
            If iListRows > 21 Then
               stAppCaseList = stAppCaseList & vbCrLf & Chr(12)
               iListRows = 0
               iExPages = iExPages + 1
            ElseIf iSNo > 1 Then
               stAppCaseList = stAppCaseList & vbCrLf & vbCrLf
               iListRows = iListRows + 1
            End If
            'Modify by Morgan 2007/11/19 ¤£¦L§Ç¸¹
            'stAppList = stAppList & iSNo & "." & .Fields("cu05") & " " & .Fields("cu88")
            stAppList = stAppList & .Fields("cu05") & " " & .Fields("cu88")
            iLine = iLine + 1
            stAppCaseList = stAppCaseList & "Applicant: " & .Fields("cu05") & " " & .Fields("cu88")
            iListRows = iListRows + 1

            If Not IsNull(.Fields("cu89")) Then
               'Modify by Morgan 2007/11/19 ¤£¦L§Ç¸¹
               'stAppList = stAppList & vbCrLf & Space(16) & .Fields("cu89") & " " & .Fields("cu90")
               stAppList = stAppList & vbCrLf & Space(14) & .Fields("cu89") & " " & .Fields("cu90")
               iLine = iLine + 1
               stAppCaseList = stAppCaseList & vbCrLf & Space(11) & .Fields("cu89") & " " & .Fields("cu90")
               iListRows = iListRows + 1
            End If

            stAppCaseList = stAppCaseList & vbCrLf & vbCrLf & _
               "Your Ref                       Case No.                       Our Ref         Application No.   Patent No.     " & vbCrLf & _
               "---------------------------------------------------------------------------------------------------------------" & vbCrLf & _
               Left(.Fields("pa77") & Space(31), 31) & Left(.Fields("pa48") & Space(31), 31) & Left(.Fields("OurRef") & Space(16), 16) & Left(.Fields("pa11") & Space(18), 18) & Left(.Fields("pa22") & Space(15), 15)

            iListRows = iListRows + 4
            strLstCuNo = "" & .Fields("pa26")
         Else
            If iListRows > 25 Then
               stAppCaseList = stAppCaseList & vbCrLf & Chr(12) & _
               "Your Ref                       Case No.                       Our Ref         Application No.   Patent No.     " & vbCrLf & _
               "---------------------------------------------------------------------------------------------------------------"

               iListRows = 2
               iExPages = iExPages + 1
            End If
            stAppCaseList = stAppCaseList & vbCrLf & _
               Left(.Fields("pa77") & Space(31), 31) & Left(.Fields("pa48") & Space(31), 31) & Left(.Fields("OurRef") & Space(16), 16) & Left(.Fields("pa11") & Space(18), 18) & Left(.Fields("pa22") & Space(15), 15)

            iListRows = iListRows + 1
         End If

         .MoveNext
      Loop
      
      pub_AddressListSN = pub_AddressListSN + 1
      PUB_AddNewAddressList strUserNum, stCP(1), stCP(2), stCP(3), stCP(4), "" & pub_AddressListSN, "0"
      
      If PrintLetter1(stET02, stCP10, stAppList, iLine, stAppCaseList, iExPages, stTag) = False Then
         iUB = iUB + 1
         ReDim Preserve ErrFA(iUB)
         ErrFA(iUB) = stLstKey
         MsgBox "[" & stLstKey & "]²£¥Í«H¨ç¿ù»~¡I"
      End If
      
      End With
      
      If UBound(ErrFA) > 0 Then
RePrint:
         Printer.Print "«H¨ç¿ù»~²M³æ"
         Printer.Print Join(ErrFA, vbCrLf)
         Printer.EndDoc
         If MsgBox("¬O§_­«¦L«H¨ç¿ù»~²M³æ¡H", vbYesNo + vbDefaultButton1) = vbYes Then
            GoTo RePrint
         End If
      End If
   Else
      MsgBox "µL¸ê®Æ¡I"
   End If
   Set rsQuery = Nothing
End Sub

Private Sub cmdMail_Click()
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   BatchMail
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

'FCT±H«H
'Add By Sindy 2009/08/18 FCT°w¹ïªk°ê¤Î¼w°ê¥N²z¤H±H«H(ªñ¤T¦~¤º»P¥»©Ò¦³®×¥ó©¹¨Óªº¨Æ°È©Ò¤Îª½±µ«È¤á)
Private Sub Porcess980818()
Dim strSql As String
Dim adoRst As New ADODB.Recordset
Dim strMailText As String
Dim iErrNo As Integer

On Error GoTo DebugErr
     
   strSql = "select distinct(tm44),fa16,fa08,substr(fa10,1,3) " & _
                  "From caseprogress, trademark, fagent " & _
                  "Where CP01 = tm01 And cp02 = tm02 And cp03 = tm03 And cp04 = tm04 " & _
                  "and cp01='FCT' " & _
                  "and cp05>=20070101 " & _
                  "and cp57 is null " & _
                  "and substr(cp09,1,1)='A' " & _
                  "and not tm44 is null " & _
                  "and tm44=fa01||fa02 " & _
                  "and (substr(fa10,1,3)='203' or substr(fa10,1,3)='231') " & _
                  "Union All " & _
                  "select distinct(tm23),cu20,cu59,substr(cu10,1,3) " & _
                  "From caseprogress, trademark, customer " & _
                  "Where CP01 = tm01 And cp02 = tm02 And cp03 = tm03 And cp04 = tm04 " & _
                  "and cp01='FCT' " & _
                  "and cp05>=20070101 " & _
                  "and cp57 is null " & _
                  "and substr(cp09,1,1)='A' " & _
                  "and tm44 is null " & _
                  "and not tm23 is null " & _
                  "and tm23=cu01||cu02 " & _
                  "and (substr(cu10,1,3)='203' or substr(cu10,1,3)='231') "
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
   With adoRst
         If .RecordCount > 0 And .RecordCount <> 0 Then
         .MoveFirst
         Do While Not .EOF
            If Trim("" & .Fields(1)) <> "" And Not IsNull(.Fields(1)) And _
               UCase(Trim("" & .Fields(1))) <> "NO" Then
               '±H«H¤º®e
               'strMailText = "Attn:" & "" & .Fields(2) & vbCrLf & vbCrLf
               strMailText = "Dear Sirs," & vbCrLf & vbCrLf & _
               "Kindly be advised that the person "
               If .Fields(3) = "203" Then '¹ï¶H¬°ªk°ê¥Î
                  strMailText = strMailText & "Ms. Jerry Lo "
               ElseIf .Fields(3) = "231" Then '¹ï¶H¬°¼w°ê¥Î
                  strMailText = strMailText & "Ms. Anette Kao "
               End If
               strMailText = strMailText & "is no longer working in our firm." & vbCrLf & _
                                                           "Please do not hesitate to contact us via ipdept@taie.com.tw for any orders and inquiries." & vbCrLf & _
                                                           "Thanks." & vbCrLf & vbCrLf & vbCrLf & _
               "Best regards," & vbCrLf & vbCrLf & _
               "Frances Chen¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@Fred C.T. Yen" & vbCrLf & _
               "Head of Trademark Department¡@¡@¡@¡@Patent Attorney" & vbCrLf & _
               "¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@ Managing Partner"
               'µo mail
               'SendMAPIMail "97038@taie.com.tw", "test", strMailText, ""
               'strusernum="ipdept@taie.com.tw"
               'Call PUB_SendMail("", .Fields(1), "", "notice", strMailText, , , True, True)
               'Call PUB_SendMail("", "97038@taie.com.tw", "", "notice", strMailText, , , True, True)
               'Call PUB_SendMail("", "68005@taie.com.tw", "", "notice", strMailText, , , True, True)
            End If
            .MoveNext
         Loop
         End If
   End With
   End If
   
   MsgBox "±H«H§¹²¦¡I"
   Exit Sub
DebugErr:
        MsgBox Err.Description
End Sub
'2009/08/18 End

'Add by Morgan 2005/10/13
'µoMail
Private Function SendMAPIMail(p_RecID$, p_Sub$, p_Text$, p_AttPath$) As Boolean
   Dim tmpnext As String, bolResume As Boolean
   Dim ii As Integer, arrAtt
   Dim strRecID As String
   
   'Add by Morgan 2009/2/23
   'strRecID = ChkMailId(p_RecID)
   
On Error GoTo ErrHandle
   
   'µo mail
   tmpnext = "·Ç³Æµo mail..."
   MAPISession1.LogonUI = False
   MAPISession1.UserName = "ipdept@taie.com.tw" '"administrator"
   tmpnext = "·Ç³Æµn¤J¶l¥ó¦øªA¾¹..."
   MAPISession1.SignOn
   tmpnext = "µn¤J¶l¥ó¦øªA¾¹..."
   MAPIMessages1.SessionID = MAPISession1.SessionID
   
   MAPIMessages1.MsgIndex = -1
   tmpnext = "«Ø¥ß¶l¥ó..."
   MAPIMessages1.Compose
   MAPIMessages1.MsgSubject = p_Sub$
   'Add by Morgan 2008/8/27 §ï¦hªþ¥ó
   If p_AttPath$ <> Empty Then
      arrAtt = Split(p_AttPath$, ";")
      p_Text$ = Space(UBound(arrAtt) + 1) & vbCrLf & p_Text$
   End If
   MAPIMessages1.MsgNoteText = p_Text$
   If p_AttPath$ <> Empty Then
      'Add by Morgan 2008/8/27 §ï¦hªþ¥ó
      For ii = 0 To UBound(arrAtt)
         If arrAtt(ii) <> "" Then
            MAPIMessages1.AttachmentIndex = ii
            MAPIMessages1.AttachmentPosition = ii
            MAPIMessages1.AttachmentPathName = arrAtt(ii)
         End If
      Next
   End If
   MAPIMessages1.RecipIndex = 0
   MAPIMessages1.RecipDisplayName = p_RecID 'strRecID
   MAPIMessages1.ResolveName
   tmpnext = "·Ç³Æ¦s¤J¶l¥ó..."
   MAPIMessages1.Send
   MAPISession1.SignOff
   'Shell "net send /domain:taient4 '¨C¤é¦Û°Ê¶l¥ó¸ê®Æ¤w°e¥X¡A½Ð²M°£¶l¥ó³Æ¥÷' ", vbNormalNoFocus
   SendMAPIMail = True
   Exit Function
ErrHandle:
'Modify by Morgan 2007/9/17 §ï¼gLog,§_«hMail±¾±¼¤£·|°±
'   If bolResume = False And Err.Number = 32050 Then
'      Err.Clear
'      bolResume = True
'      Resume Next
'   End If
   WLog tmpnext & " " & Err.Description
'end 2007/9/17
   'Add by Morgan 2009/4/3 µo«H¥¢±Ñ­nµn¥X
   If MAPISession1.SessionID <> 0 Then MAPISession1.SignOff
End Function

'FCP²×¤î¿ì²z
Private Sub Process980828()
   Dim stSQL As String, stSQL1 As String, stSQL2 As String, ii As Integer
   Dim stDate1 As String, stDate2 As String, stDate3 As String, stDate4 As String
   Dim stCon1 As String, stCon2 As String
   Dim intR As Integer
      
   stSQL = ""
   stCon1 = "": stCon2 = ""
   '²Ä¤@¦¸
   If strSrvDate(1) = "20090901" Then
      stCon1 = " and c1.cp05<" & 20080101 & " and nvl(pa58,0)<" & 20090901
   ElseIf strSrvDate(1) = "20091001" Then
      stCon1 = " and c1.cp05>=20080101 and c1.cp05<20080201" & " and nvl(pa58,0)<20091001"
      stCon2 = " and c1.cp05<20080101 and pa58>=20090901 and pa58<20091001"
   ElseIf InStr("0101,0401,0701,1001", Mid(strSrvDate(1), 5)) > 0 Then
      stDate1 = strSrvDate(1) '³øªí¤é
      stDate2 = CompDate(1, -20, stDate1) '³øªí¤é-20¤ë
      stDate3 = CompDate(1, -23, stDate1) '³øªí¤é-23¤ë
      stDate4 = CompDate(1, -3, stDate1) '³øªí¤é-3¤ë
      stCon1 = " and c1.cp05>=" & stDate3 & " and c1.cp05<" & stDate2 & " and nvl(pa58,0)<" & stDate1
      stCon2 = " and c1.cp05<" & stDate3 & " and pa58>=" & stDate4 & " and pa58<" & stDate1
   End If
   If stCon1 <> "" Then
      'C1:´¼Åv¤H­û,C2:¥N²z¤H½s¸¹¦WºÙ,C3:¥Ó½Ð¤H½s¸¹¦WºÙ,C4:¥»©Ò®×¸¹
      'C5:¥Ó½Ð¤é,C6:¹ê¼fµo¤å¤é,C7:1204¦¬¤å¤é,C8:ºM¦^µo¤å¤é
      'C9:ªL«ß®v®×¥ó,C10:­«·s©e¥ôµo¤å¤é,C11 ­«·s©e¥ô­ã»é
      stSQL1 = "select na51 C0,st02 C1" & _
         ",substrb(rtrim(fa05||' '||fa63||' '||fa64||' '||fa65),1,34) C2" & _
         ",substrb(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),1,40) C3" & _
         ",c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04) C4" & _
         ",pa10 C5,c2.cp27 C6,c1.cp05 C7,c3.cp27 C8,decode(lc01,null,'',' Y ') C9" & _
         ",c4.cp27 C10,c4.cp24 C11,st17,pa75,cu01||cu02" & _
         " from caseprogress c1,patent,fagent,nation,staff,caseprogress c2,caseprogress c3,lincase,caseprogress c4,customer" & _
         " where c1.cp01='FCP' and c1.cp10='1204'" & _
         " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04" & _
         " and pa08='1' and pa57='Y'" & _
         " and not exists(select * from caseprogress x where x.cp01=c1.cp01 and x.cp02=c1.cp02" & _
         " and x.cp03=c1.cp03 and x.cp04=c1.cp04 and x.cp10 in ('1001','1002','1201','1202','1203','1221','1307')" & _
         " and x.cp05>c1.cp05) and fa01(+)=substrb(pa75,1,8) and fa02(+)=substrb(pa75,9)" & _
         " and na01(+)=fa10 and st01(+)=na51 and c2.cp09(+)=c1.cp43" & _
         " and c3.cp01(+)=c1.cp01 and c3.cp02(+)=c1.cp02 and c3.cp03(+)=c1.cp03 and c3.cp04(+)=c1.cp04 and c3.cp10(+)='413'" & _
         " and lc01(+)=c1.cp01 and lc02(+)=c1.cp02 and lc03(+)=c1.cp03 and lc04(+)=c1.cp04" & _
         " and c4.cp01(+)=c1.cp01 and c4.cp02(+)=c1.cp02 and c4.cp03(+)=c1.cp03 and c4.cp04(+)=c1.cp04 and c4.cp10(+)='928'"
      
      stSQL2 = "select na51 C0,st02 C1" & _
         ",substrb(rtrim(fa05||' '||fa63||' '||fa64||' '||fa65),1,34) C2" & _
         ",substrb(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),1,40) C3" & _
         ",c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04) C4" & _
         ",pa10 C5,c2.cp27 C6,c1.cp05 C7,c3.cp27 C8,decode(lc01,null,'',' Y ') C9" & _
         ",c4.cp27 C10,c4.cp24 C11,st17,pa75,cu01||cu02" & _
         " from patent,caseprogress c1,fagent,nation,staff,caseprogress c2,caseprogress c3,lincase,caseprogress c4,customer" & _
         " where pa01='FCP' and pa08='1' and pa57='Y'" & _
         " and c1.cp01(+)=pa01 and c1.cp02(+)=pa02 and c1.cp03(+)=pa03 and c1.cp04(+)=pa04" & _
         " and c1.cp10='1204'" & _
         " and not exists(select * from caseprogress x where x.cp01=c1.cp01 and x.cp02=c1.cp02" & _
         " and x.cp03=c1.cp03 and x.cp04=c1.cp04 and x.cp10 in ('1001','1002','1201','1202','1203','1221','1307')" & _
         " and x.cp05>c1.cp05) and fa01(+)=substrb(pa75,1,8) and fa02(+)=substrb(pa75,9)" & _
         " and na01(+)=fa10 and st01(+)=na51 and c2.cp09(+)=c1.cp43" & _
         " and c3.cp01(+)=c1.cp01 and c3.cp02(+)=c1.cp02 and c3.cp03(+)=c1.cp03 and c3.cp04(+)=c1.cp04 and c3.cp10(+)='413'" & _
         " and lc01(+)=c1.cp01 and lc02(+)=c1.cp02 and lc03(+)=c1.cp03 and lc04(+)=c1.cp04" & _
         " and c4.cp01(+)=c1.cp01 and c4.cp02(+)=c1.cp02 and c4.cp03(+)=c1.cp03 and c4.cp04(+)=c1.cp04 and c4.cp10(+)='928'"
      For ii = 0 To 4
         If ii > 0 Then
            stSQL = stSQL & " union "
         End If
         stSQL = stSQL & stSQL1 & " and pa" & 26 + ii & " is not null and cu01(+)=substr(pa" & 26 + ii & ",1,8) and cu02(+)=substr(pa" & 26 + ii & ",9)" & stCon1
      Next ii
      
      If stCon2 <> "" Then
         For ii = 0 To 4
            stSQL = stSQL & " union " & stSQL2 & " and pa" & 26 + ii & " is not null and cu01(+)=substr(pa" & 26 + ii & ",1,8) and cu02(+)=substr(pa" & 26 + ii & ",9)" & stCon2
         Next ii
      End If
      stSQL = stSQL & " order by 1,2,3,4,5"
      intR = 1
      Set RsTemp = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
      
      End If
   End If
End Sub
'Add by Morgan 2009/12/14
'¥xÆW³W¶O½Õº¦³qª¾
Private Sub Command981214()
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   If Process981214 = True Then
      BatchMail3
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Function Process981214() As Boolean
   Dim stSQL As String, stVTB As String, lngMS01 As Long, lngRec As Long, ii As Integer
   Dim arrToMail
   
   lngMS01 = 18
   
   stVTB = "select fa16 MSD02,FA01||FA02 MSD07 from fagent where fa02='0' and instr(fa16,'@')>0" & _
      " and fa69 is null and fa02='0' and fa76<'C' and instr(upper(fa29),'FAIL')=0" & _
      " Union select fa80 MSD02,FA01||FA02 MSD07 from fagent where fa02='0' and instr(fa80,'@')>0" & _
      " and fa69 is null and fa02='0' and fa76<'C' and instr(upper(fa29),'FAIL')=0" & _
      " Union select fa81 MSD02,FA01||FA02 MSD07 from fagent where fa02='0' and instr(fa81,'@')>0" & _
      " and fa69 is null and fa02='0' and fa76<'C' and instr(upper(fa29),'FAIL')=0" & _
      " Union select fa82 MSD02,FA01||FA02 MSD07 from fagent where fa02='0' and instr(fa82,'@')>0" & _
      " and fa69 is null and fa02='0' and fa76<'C' and instr(upper(fa29),'FAIL')=0" & _
      " Union select pcc08 MSD02,FA01||FA02||'-'||PCC02 MSD07 from fagent,potcustcont where fa02='0'" & _
      " and fa69 is null and fa02='0' and fa76<'C'" & _
      " and pcc01(+)=fa01 and instr(pcc08,'@')>0 and instr(upper(pcc13),'FAIL')=0"
   
   stSQL = "insert into MailScheduleDetail(MSD01,MSD02,MSD06,MSD07)" & _
      " SELECT " & lngMS01 & ",MSD02,'None',min(MSD07) FROM (" & stVTB & ") X group by MSD02"
   
   cnnConnection.BeginTrans

On Error GoTo ErrHnd

   cnnConnection.Execute stSQL, lngRec
   
   '¦h­Ó«H½c©ñ¤@°_ªº¸ê®Æ
   stSQL = "select msd02,msd07 from MailScheduleDetail where msd01=" & lngMS01 & " and instr(msd02,';')>0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         arrToMail = Split("" & .Fields(0), ";") '¥h°£«e«áªºªÅ¥Õ©M¸õ¦æ²Å¸¹
         For ii = 0 To UBound(arrToMail)
            arrToMail(ii) = Trim(Replace(arrToMail(ii), vbCrLf, ""))
            If InStr(arrToMail(ii), "@") > 0 Then
               stSQL = "select msd02 from MailScheduleDetail where msd01=" & lngMS01 & " and msd02='" & arrToMail(ii) & "'"
               intI = 1
               Set AdoRecordSet3 = ClsLawReadRstMsg(intI, stSQL)
               If intI = 0 Then
                  stSQL = " insert into MailScheduleDetail(MSD01,MSD02,MSD06,MSD07) values(" & lngMS01 & ",'" & arrToMail(ii) & "','None','" & RsTemp("msd07") & "')"
                  cnnConnection.Execute stSQL, intI
               End If
            End If
         Next
         .MoveNext
      Loop
      End With
      stSQL = "delete from MailScheduleDetail where msd01=" & lngMS01 & " and instr(msd02,';')>0"
      cnnConnection.Execute stSQL, lngRec
   End If
   
   stSQL = "update MailSchedule set ms10=(select count(*) from MailScheduledetail where msd01=ms01) where ms01=" & lngMS01
   cnnConnection.Execute stSQL, lngRec
   
   cnnConnection.CommitTrans
   Process981214 = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
   
End Function
'¥D¦®¥[Our Ref
Private Sub BatchMail3(Optional pReMail As Boolean)

   Dim stSQL As String, Rs As ADODB.Recordset
   Dim stSubjectLead As String, stSubject As String
   Dim stFromName As String, stFromMail As String, stAttPath As String, stScript As String
   Dim arrToMail, stToMail As String, stToName As String, stMIME As String
   Dim adoRst As New ADODB.Recordset, lngRec As Long, iTurn As Integer
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte
   Dim iErrNo As Integer
   Dim stMS01 As String
   Dim bFail As Boolean
   Dim ii As Integer
   Dim iWait As Integer

  On Error GoTo ErrHnd
  
   AlertMsg "¶}©l±Æµ{!"

   stAttPath = App.path & "\edm.mht"
   
   If pReMail = False Then
      stSQL = "SELECT * FROM MailSchedule,MailScheduleTemplet WHERE MS08*1000000+MS09<=TO_CHAR(SYSDATE,'YYYYMMDDHH24MISS') AND  MS11 IS NULL and mst01(+)=ms01"
   Else
      stSQL = "SELECT * FROM MailSchedule,MailScheduleTemplet WHERE MS08*1000000+MS09<=TO_CHAR(SYSDATE,'YYYYMMDDHH24MISS') and mst01(+)=ms01 and ms01=(select max(msd01) from mailscheduledetail where msd05=1)"
   End If
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      With Rs
      Do While Not .EOF
         stMS01 = "" & .Fields("ms01")
         stSubjectLead = "" & Rs.Fields("ms02") '¥D¦®
         stFromMail = "" & .Fields("ms03") '±H¥ó«H½c
         stFromName = "" & .Fields("ms14") '±H¥ó¦WºÙ
         If stFromName = "" Then
            stFromName = MailName
         End If
         lngSize = Val(.Fields("mst02").Value)
         
         ReDim bytes(lngSize)
         bytes() = .Fields("mst03").GetChunk(lngSize)
         iFileNo = FreeFile
         If fso.FileExists(stAttPath) Then
            Kill stAttPath
         End If
         Open stAttPath For Binary Access Write As #iFileNo
         Put #iFileNo, , bytes()
         Close #iFileNo
         
         stMIME = GetMime(stAttPath, IIf(.Fields("mst04") = "Y", True, False))
         
         stScript = "select msd02,msd07 from MailScheduleDetail where msd01='" & stMS01 & "' and (msd03 is null or msd05=1)"
         If adoRst.State <> adStateClosed Then adoRst.Close
         With adoRst
         .CursorLocation = adUseClient
         .Open stScript, cnnConnection, adOpenForwardOnly, adLockReadOnly
         If .RecordCount > 0 Then
            ProgressBar1.max = .RecordCount
            ProgressBar1.Min = 0
            ProgressBar1.Value = 0
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            lngRec = 0
            If InStr(stFromName, "¥x¤@") > 0 Then
               iWait = 60
            Else
               iWait = 1
            End If
            Do While Not .EOF
               '¥D¦®
               If Not IsNull(.Fields("msd07")) Then
                  stSubject = stSubjectLead & " ( O/Ref: " & .Fields("msd07") & " )"
               Else
                  stSubject = stSubjectLead
               End If
               lngRec = lngRec + 1
               arrToMail = Split(.Fields(0), ";")  '¥h°£«e«áªºªÅ¥Õ©M¸õ¦æ²Å¸¹
               stToMail = Trim(Replace(arrToMail(0), vbCrLf, ""))
               stToName = stToMail
               'Modify by Morgan 2009/5/6 §ï1¤ÀÄÁ
               'Sleep 'µ¥1¬í¦A±H
               'Modify by Morgan 2009/7/13 °ê¥~¹q¤l³ø¼È®É§ï¦^
               'Sleep 60
               Sleep iWait
               If SendXMail(stFromName, stFromMail, stToName, stToMail, stSubject, stMIME, iErrNo) = True Then
                  If UpdateDetail(stMS01, "" & .Fields(0)) = False Then
                     AlertMsg "#" & lngRec & ",email:" & stToName & ",§ó·s¥¢±Ñ !"
                  End If
                  For ii = 1 To UBound(arrToMail)
                     stToMail = Trim(Replace(arrToMail(ii), vbCrLf, ""))
                     stToName = stToMail
                     Sleep 'µ¥1¬í¦A±H
                     SendXMail stFromName, stFromMail, stToName, stToMail, stSubject, stMIME, iErrNo
                  Next
               Else
                  bFail = True
                  AlertMsg "#" & lngRec & ",email:" & stToName & ",±H«H¥¢±Ñ(" & iErrNo & ") !"
                  If UpdateDetail(stMS01, "" & .Fields(0), True) = False Then
                     AlertMsg "#" & lngRec & ",email:" & stToName & ",§ó·s¥¢±Ñ !"
                  End If
               End If
               ProgressBar1.Value = ProgressBar1.Value + 1
               lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
               DoEvents
               .MoveNext
            Loop
            UpdateSchedule stMS01, str(lngRec)
         End If
         End With
         .MoveNext
      Loop
      End With
   End If
   AlertMsg "µ²§ô±Æµ{!"
   
ErrHnd:
   If Err.Number <> 0 Then
      AlertMsg Err.Description
   End If
   
   For iTurn = 0 To List1.ListCount
      WLog List1.List(iTurn), 1
   Next

End Sub
'99°²¤é³qª¾
Private Sub Command981215()
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   Process981215
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub


Private Sub Command1_Click()
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   Process990128 24
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   BatchMail1081016
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

'Added by Morgan 2019/10/16
'FCP³]­p±M¥Î´Á©µªø³qª¾(108±M§Q·sªk)
Private Sub BatchMail1081016()
   If MsgBox("¬O§_«Ø¥ß XLS ªþ¥ó¡H", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
      If Process1081016 = False Then
         Exit Sub
      End If
   End If
   
   If MsgBox("¬O§_¶}©l±HµoEMail¡H", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
      Exit Sub
   End If
   
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   Dim stSubject As String, stSubjectLead As String
   Dim stFromMail As String, stAttPath As String
   Dim stToMail As String, stToName As String
   Dim stBoundryTag As String, stMimeHead As String, stMIME As String
   Dim lngRec As Long
   
   Dim iErrNo As Integer
   Dim stMS01 As String
   Dim bFail As Boolean
   Dim stMontherPath As String
   Dim stSamplePath As String
   Dim stCon As String

   

  On Error GoTo ErrHnd
  
   stMS01 = 594
   stMontherPath = PUB_Getdesktop & "\EMailAtt"
   stSamplePath = stMontherPath & "\SampleMail.eml"
   
   If GetTemplete(stMS01, stSamplePath) = False Then
      MsgBox "¶l¥ó½d¥»ÀÉ¤U¸ü¥¢±Ñ¡I", vbCritical
      GoTo ExitFlag
   End If
   
   stMimeHead = GetMimeHead(stSamplePath, stBoundryTag)
   If stBoundryTag = "" Then
      MsgBox "ªþ¥[ÀÉ®×ªººX¼ÐµLªk¨ú±o!", vbCritical
      GoTo ExitFlag
   End If
  
   '±Æµ{¥DÀÉ
   stSQL = "SELECT * FROM MailSchedule WHERE ms01=" & stMS01
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      'stSubjectLead = "" & rsQuery.Fields("ms02") '¥D¦®
      stSubjectLead = "Your Design Patent Term Extended for 15 Years"
      stFromMail = "" & rsQuery.Fields("ms03") '±H¥ó«H½c
      End With
   Else
      MsgBox "¶l¥ó±Æµ{Åª¨ú¥¢±Ñ!", vbCritical
      GoTo ExitFlag
   End If
      
   If Text1 <> "" Then
      stCon = " and msd06='" & Text1 & "'"
   End If
   
   'stCon = " and msd02='77015@taie.com.tw'"
   
   '±Æµ{©ú²Ó
   stSQL = "SELECT st07,msd06,msd02,na51 FROM MailScheduledetail,fagent,nation,staff" & _
      " WHERE msd01=" & stMS01 & " and msd03 is null" & stCon & _
      " and fa01(+)=substr(msd06,1,8) and fa02(+)=substr(msd06,9) and na01(+)=fa10 and st01(+)=na51" & _
      " order by 1,2"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      .MoveFirst
      ProgressBar1.max = .RecordCount
      ProgressBar1.Min = 0
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      lngRec = 0
      Do While Not .EOF
         lngRec = lngRec + 1
         stAttPath = stMontherPath & "\" & .Fields("msd06") & "\" & "Design patent term extended.xls"
         If Dir(stAttPath) = "" Then
            MsgBox "#" & lngRec & ",path:" & stAttPath & ",ªþ¥óÅª¨ú¥¢±Ñ !", vbCritical
            GoTo ExitFlag
         Else
            stSubject = .Fields("st07") & " " & stSubjectLead & " [Our Ref: " & .Fields("msd06") & ".A11]"
            stToMail = Trim(Replace(Replace("" & .Fields("msd02"), vbCrLf, ""), " ", "")) '¥h°£«e«áªºªÅ¥Õ©M¸õ¦æ²Å¸¹
            stToName = stToMail
            Sleep 2 'µ¥1¬í¦A±H
            stMIME = stMimeHead & GetAttMime(stAttPath, stBoundryTag)
                        
            If SendXMail(MailName, stFromMail, stToName, stToMail, stSubject, stMIME, iErrNo, True) = True Then
               If UpdateDetail(stMS01, "" & .Fields("msd02"), , "" & .Fields("msd06")) = False Then
                  MsgBox "#" & lngRec & ",email:" & stToName & ",§ó·s¥¢±Ñ !", vbCritical
                  GoTo ExitFlag
               End If
            Else
               bFail = True
               'MsgBox "#" & lngRec & ",email:" & stToName & ",±H«H¥¢±Ñ(" & iErrNo & ") !", vbCritical
               If UpdateDetail(stMS01, "" & .Fields("msd02"), True, "" & .Fields("msd06")) = False Then
                  MsgBox "#" & lngRec & ",email:" & stToName & ",§ó·s¥¢±Ñ !", vbCritical
                  GoTo ExitFlag
               End If
            End If
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
         End If
         .MoveNext
      Loop
      End With
   End If
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
ExitFlag:
   Set rsQuery = Nothing
End Sub

Private Function Process1081016() As Boolean
   Dim stSQL As String
   Dim intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stFaNo As String
   Dim strTempFile As String, stXLSFileName As String, stXLSPath As String, stXLSFullPath As String
   Dim xlsReport, wksReport
   Dim iRow As Integer
   Dim lngRec As Integer
   
   stSQL = "select st07 C1,pa75||' '||rtrim(fa05||' '||fa63||' '||fa64||' '||fa65) C2" & _
      ",pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C3" & _
      ",pa77 C4,pa48 C5,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) C6" & _
      ",pa11 C7,pa22 C8,sqldatew(pa25) C9" & _
      ",sqldatew(to_char(add_months(to_date(pa10,'yyyymmdd'),180)-1,'yyyymmdd')) C10" & _
      ",decode(np02,'',to_char(to_date(pa14+LASTYEAR(pa72)*10000,'yyyymmdd')-1,'yyyy/mm/dd')) C11" & _
      ",decode(np02,'','X') C12,decode(np02,'','','X') C13" & _
      ",pa75 From patent, fagent, nation, staff, customer" & _
      ",(select distinct np02,np03,np04,np05" & _
      " from patent,nextprogress a where pa01='FCP' and pa08='3' and pa24<20191101 and pa25>=20191101 and pa57 is null" & _
      " and a.np02(+)=pa01 and a.np03(+)=pa02 and a.np04(+)=pa03 and a.np05(+)=pa04 and a.np06='N' and a.np07='605'" & _
      " and not exists(select * from nextprogress b where b.np02=pa01 and b.np03=pa02 and b.np04=pa03 and b.np05=pa04 and b.np07='605' and b.np09>a.np09)" & _
      ") N where pa01='FCP' and pa08='3' and pa24<20191101 and pa25>=20191101 and pa57 is null" & _
      " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9)" & _
      " and na01(+)=fa10 and st01(+)=na51" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
      " and np02(+)=pa01 and np03(+)=pa02 and np04(+)=pa03 and np05(+)=pa04" & _
      " order by 1,2,3,4,5"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      stXLSFileName = "Design patent term extended.xls"
      stXLSPath = PUB_Getdesktop & "\EMailAtt"
      If Dir(stXLSPath, vbDirectory) = "" Then MkDir stXLSPath
      strTempFile = stXLSPath & "\" & stXLSFileName
      
      Set xlsReport = CreateObject("Excel.Application")
      xlsReport.Visible = True
         
      With rsQuery
      
      ProgressBar1.max = .RecordCount
      ProgressBar1.Min = 0
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      lngRec = 0
      
      Do While Not .EOF
         lngRec = lngRec + 1
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         If stFaNo <> .Fields("pa75") Then
            If stFaNo <> "" Then
               stXLSFullPath = stXLSPath & "\" & stFaNo
               If Dir(stXLSFullPath, vbDirectory) = "" Then MkDir stXLSFullPath
               stXLSFullPath = stXLSFullPath & "\" & stXLSFileName
               xlsReport.Workbooks(1).SaveAs stXLSFullPath
               xlsReport.Workbooks.Close
            End If
            
            xlsReport.Workbooks.Open strTempFile
            Set wksReport = xlsReport.Worksheets(1)
            wksReport.Cells.NumberFormatLocal = "@"
            stFaNo = .Fields("pa75")
            iRow = 1
         End If
         
         iRow = iRow + 1
         wksReport.Range("A" & iRow) = "" & .Fields("C1") 'TAI E ID
         wksReport.Range("B" & iRow) = "" & .Fields("C2") 'ASSOICIATE
         wksReport.Range("C" & iRow) = "" & .Fields("C3") 'OUR REF
         wksReport.Range("D" & iRow) = "" & .Fields("C4") 'YOUR REF
         wksReport.Range("E" & iRow) = "" & .Fields("C5") 'CASE NO.
         wksReport.Range("F" & iRow) = "" & .Fields("C6") 'APPLICANT
         wksReport.Range("G" & iRow) = "" & .Fields("C7") 'APPLN NO.
         wksReport.Range("H" & iRow) = "" & .Fields("C8") 'PATENT NO.
         wksReport.Range("I" & iRow) = "" & .Fields("C9") 'ORIGINAL EXPIRY DATE
         wksReport.Range("J" & iRow) = "" & .Fields("C10") 'NEW EXPIRY DATE
         wksReport.Range("K" & iRow) = "" & .Fields("C11") 'NEXT ANNUITY DUE
         wksReport.Range("L" & iRow) = "" & .Fields("C12") 'ANNUITY HANDLED BY TAI E
         wksReport.Range("M" & iRow) = "" & .Fields("C13") 'ANNUITY PAID BY OTHER CHANNEL
         
         .MoveNext
      Loop
      
      stXLSFullPath = stXLSPath & "\" & stFaNo
      If Dir(stXLSFullPath, vbDirectory) = "" Then MkDir stXLSFullPath
      stXLSFullPath = stXLSFullPath & "\" & stXLSFileName
      xlsReport.Workbooks(1).SaveAs stXLSFullPath
      xlsReport.Workbooks.Close
      xlsReport.Quit
      
      End With
   End If
   
   Process1081016 = True
   
ErrHandle:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
   Set rsQuery = Nothing
   Set wksReport = Nothing
   Set xlsReport = Nothing
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   txtMailTo = strUserNum
End Sub

'¼È°±
Private Sub Sleep(Optional iSec As Integer = 1)
    frmWait.iWaitSec = iSec
    frmWait.Timer1.Interval = 1000
    frmWait.Show vbModal
End Sub


'¦Û°Ê°õ¦æ
Private Sub BatchMail(Optional pReMail As Boolean)

   Dim stSQL As String
   Dim stSubject As String, stSubjectLead As String
   Dim stFromMail As String, stAttPath As String, stScript As String
   Dim arrToMail, stToMail As String, stToName As String, stMIME As String
   Dim adoRst As New ADODB.Recordset, lngRec As Long, iTurn As Integer
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte
   Dim iErrNo As Integer
   Dim stMS01 As String
   Dim bFail As Boolean
   Dim ii As Integer
   Dim stMontherPath As String
   Dim stSamplePath As String, stMimeHead As String
   Dim stCon As String

   Dim stBoundryTag As String

  On Error GoTo ErrHnd
  
   stMS01 = 3
   stSamplePath = "c:\sample.eml"
   stMontherPath = PUB_Getdesktop & "\AE\Eng\EMail\pdf"
   
   AlertMsg "¶}©l±H«H!"
   
   stSQL = "SELECT * FROM MailSchedule,MailScheduleTemplet WHERE ms01=" & stMS01 & " and mst01(+)=ms01"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      With adoRst
      stSubjectLead = "" & .Fields("ms02") '¥D¦®
      stFromMail = "" & .Fields("ms03") '±H¥ó«H½c
      lngSize = Val(.Fields("mst02").Value) '¼Ë¥»¶l¥ó¤j¤p
         
      ReDim bytes(lngSize)
      bytes() = .Fields("mst03").GetChunk(lngSize)
      iFileNo = FreeFile
      If fso.FileExists(stSamplePath) Then
         Kill stSamplePath
      End If
      Open stSamplePath For Binary Access Write As #iFileNo
      Put #iFileNo, , bytes()
      Close #iFileNo
      End With
      
      stMimeHead = GetMimeHead(stSamplePath, stBoundryTag)
      If stBoundryTag = "" Then
         AlertMsg "ªþ¥[ÀÉ®×ªººX¼ÐµLªk¨ú±o!"
         GoTo ExitFlag
      End If
      
      If Text1 <> "" Then
         stCon = " and msd06='" & Text1 & "'"
      End If
      
      stSQL = "SELECT * FROM MailScheduledetail" & _
         " WHERE msd01=" & stMS01 & " and instr(msd02,'@')>0 and msd03 is null" & stCon
      
      intI = 1
      Set adoRst = ClsLawReadRstMsg(intI, stSQL)
      If intI = 1 Then
         With adoRst
         .MoveFirst
         ProgressBar1.max = .RecordCount
         ProgressBar1.Min = 0
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         lngRec = 0
         Do While Not .EOF
            lngRec = lngRec + 1
            stAttPath = stMontherPath & "\" & .Fields("msd07")
            If Dir(stAttPath) = "" Then
               AlertMsg "#" & lngRec & ",email:" & stToName & ",ªþ¥óÅª¨ú¥¢±Ñ !"
            Else
               stSubject = stSubjectLead & " ( O/Ref: " & Mid(.Fields("msd07"), 18, 15) & " )"
               arrToMail = Split(Trim(Replace("" & .Fields("msd02"), vbCrLf, "")), ";") '¥h°£«e«áªºªÅ¥Õ©M¸õ¦æ²Å¸¹
               stToMail = arrToMail(0) '¥u±H°e²Ä¤@­Ó«H½c
               stToName = stToMail
               Sleep 2 'µ¥1¬í¦A±H
               stMIME = stMimeHead & GetAttMime(stAttPath, stBoundryTag)
               
               If SendXMail(MailName, stFromMail, stToName, stToMail, stSubject, stMIME, iErrNo) = True Then
                  If UpdateDetail(stMS01, "" & .Fields("msd02")) = False Then
                     AlertMsg "#" & lngRec & ",email:" & stToName & ",§ó·s¥¢±Ñ !"
                  End If
               Else
                  bFail = True
                  AlertMsg "#" & lngRec & ",email:" & stToName & ",±H«H¥¢±Ñ(" & iErrNo & ") !"
                  If UpdateDetail(stMS01, "" & .Fields("msd02"), True) = False Then
                     AlertMsg "#" & lngRec & ",email:" & stToName & ",§ó·s¥¢±Ñ !"
                  End If
               End If
               ProgressBar1.Value = ProgressBar1.Value + 1
               lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
               DoEvents
            End If
            'If txtMailTo <> "" Then Exit Do '´ú¸Õ
            .MoveNext
         Loop
         End With
      End If
   End If
   AlertMsg "±H¥óµ²§ô!"
   
ErrHnd:
   If Err.Number <> 0 Then
      AlertMsg Err.Description
   End If
   
   For iTurn = 0 To List1.ListCount
      WLog List1.List(iTurn), 1
   Next
ExitFlag:
   Set adoRst = Nothing
End Sub

Private Function UpdateDetail(SMD01 As String, SMD02 As String, Optional bolErr As Boolean, Optional SMD06 As String) As Boolean
   Dim stSQL As String, intR As Integer
On Error GoTo ErrHnd
   stSQL = "update MailScheduleDetail set msd03=to_char(sysdate,'yyyymmdd'),msd04=to_char(sysdate,'hh24mmss')"
   If bolErr Then
      stSQL = stSQL & ",MSD05='1'"
   Else
      stSQL = stSQL & ",MSD05=null" '­«±H¦¨¥\­n²M°£
   End If
   stSQL = stSQL & " where msd01=" & SMD01 & " and msd02='" & ChgSQL(SMD02) & "'" & IIf(SMD06 <> "", " and msd06='" & SMD06 & "'", "")
   
   cnnConnection.Execute stSQL, intR
   If intR = 1 Then
      UpdateDetail = True
   End If
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function


Private Function GetMimeHead(stSamplePath As String, Optional stOutBoundryTag As String) As String
      
   Dim ts As TextStream
   Dim strLine As String
   Dim bStart As Boolean
   Dim stMIME As String
   Dim iPos As Integer
   
   stMIME = ""
   stOutBoundryTag = ""
   If fso.FileExists(stSamplePath) Then
      Set ts = fso.OpenTextFile(stSamplePath)
      bStart = False
      Do While Not ts.AtEndOfStream
         strLine = ts.ReadLine
         If stOutBoundryTag = "" Then
            iPos = InStr(UCase(strLine), UCase("boundary="))
            If iPos > 0 Then
               '¦©°£«e«áªºÂù¤Þ¸¹
               stOutBoundryTag = Mid(strLine, iPos + 10)
               stOutBoundryTag = Left(stOutBoundryTag, Len(stOutBoundryTag) - 1)
            End If
         End If
         If bStart = False Then
            '¶}©l
            If InStr(UCase(strLine), UCase("MIME-Version: 1.0")) > 0 Then
               bStart = True
            'Added by Morgan 2019/10/17
            ElseIf InStr(UCase(strLine), UCase("Content-Type: ")) > 0 Then
               bStart = True
            End If
         End If
         
         If bStart = True Then
            'µ²§ô
            'Modified by Morgan 2019/10/17
            'If InStr(UCase(strLine), UCase("application/octet-stream")) > 0 Then
            If InStr(UCase(strLine), UCase("Content-Type: application/vnd.ms-excel")) > 0 Then
            'end 2019/10/17
               Exit Do
            Else
               stMIME = stMIME & strLine & vbCrLf
            End If
         End If
      Loop
      ts.Close
   End If
   GetMimeHead = FixFirstDot(stMIME)
End Function

Private Function GetAttMime(stAttPath As String, stBoundryTag As String) As String
   Dim stMIME As String, sFil64 As String
   Dim stFileName As String, stExt As String
   
   stFileName = Mid(stAttPath, InStrRev(stAttPath, "\") + 1)
   stExt = Mid(stFileName, InStrRev(stFileName, ".") + 1)
   
   Select Case LCase(stExt)
      Case "xls"
         stMIME = "Content-Type: application/vnd.ms-excel;" & vbCrLf
      Case "pdf"
         stMIME = "Content-Type: application/pdf;" & vbCrLf
      Case Else
         stMIME = "Content-Type: application/octet-stream;" & vbCrLf
   End Select
   
   stMIME = stMIME & "  name=""" & stFileName & """" & vbCrLf
   stMIME = stMIME & "Content-Description: " & stFileName & vbCrLf
   stMIME = stMIME & "Content-Disposition: attachment;" & vbCrLf
   stMIME = stMIME & "  filename=""" & stFileName & """" & vbCrLf
   stMIME = stMIME & "Content-Transfer-Encoding: base64" & vbCrLf & vbCrLf

   sFil64 = ConvertToBase64(stAttPath, True, True)
   stMIME = stMIME & sFil64 & vbCrLf & vbCrLf & cDASH2 & stBoundryTag & cDASH2 & vbCrLf & vbCrLf
   GetAttMime = stMIME
End Function
Private Function SendXMail(FromName$, FromMail$, ToName$, ToMail$, Subj$, strMime$, Optional iErrCode As Integer, Optional pByMailServer As Boolean = False) As Boolean
   
   Dim strData(0 To 9) As String
   Dim DateNow As String
   Dim SMTP As String
   Dim iRetry As Integer
   Dim stBas64 As String
   
   Dim ArrTmpMail, ArrTmpName
   Dim MailCnt As Integer
   Dim tmpMailStr As String
           
On Error GoTo ErrHnd

   iErrCode = 0
   Result = ""
   DoEvents
      
   SMTP = GetSMTP(pByMailServer)
   
   strData(1) = "mail from:" + Chr(32) + FromMail + vbCrLf
   stBas64 = ConvertToBase64(FromName, False, False)
   strData(3) = "From: =?Big5?B?" & stBas64 & "?= <" & FromMail & ">" & vbCrLf
   
   
   'Modified by Morgan 2019/10/18 §ï¥i¦h¦¬¥óªÌ
   'stBas64 = ConvertToBase64(ToName, False, False)
   'strData(4) = "To: =?Big5?B?" & stBas64 & "?= <" & ToMail & ">" & vbCrLf
   ArrTmpMail = Split(ToMail, ";")
   If ToName = "" Then ToName = ";"
   ArrTmpName = Split(ToName, ";")
   strData(4) = ""
   For MailCnt = 0 To UBound(ArrTmpMail)
      If ArrTmpMail(MailCnt) <> "" Then
         tmpMailStr = ConvertToBase64(CStr(ArrTmpName(MailCnt)), False, False)
         If tmpMailStr <> ArrTmpName(MailCnt) And ArrTmpName(MailCnt) <> "" Then
            strData(4) = strData(4) & "To: =?Big5?B?" & tmpMailStr & "?= <" & ArrTmpMail(MailCnt) & ">" & vbCrLf
         Else
            strData(4) = strData(4) & "To: """ & ArrTmpName(MailCnt) & """ <" & ArrTmpMail(MailCnt) & ">" & vbCrLf
         End If
      End If
   Next MailCnt
   'end 2019/10/18
      
   stBas64 = ConvertToBase64(Subj, False, False)
   strData(5) = "Subject: =?Big5?B?" & stBas64 & "?=" & vbCrLf
   
   DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(time, "hh:mm:ss") & "" & " +0800"
   
   stBas64 = ConvertToBase64(FromName, False, False)
   strData(6) = "Date:" + Chr(32) + DateNow + vbCrLf
   
   'strData(6) = strData(6) & _
      "Importance: high" & vbCrLf & _
      "X-Priority: 1" & vbCrLf & _
      "Return-Receipt-To: =?Big5?B?" & stBas64 & "?= <" & FromMail & ">" & vbCrLf
      
   strData(0) = strData(3) + strData(4) + strData(5) + strData(6)
   
   strData(9) = strMime
   If strData(9) = "" Then
      strData(7) = "MIME-Version: 1.0" & vbCrLf & _
                   "Content-Type: text/plain;" + vbCrLf & _
                   "   charset=""big5""" + vbCrLf
      strData(8) = "testing..." + vbCrLf
      strData(9) = strData(7) & strData(8)
   End If
     
   strData(0) = strData(0) + strData(9)
   
RetryPoint:
   
   If Winsock1.State <> sckClosed Then Winsock1.Close
   
   Winsock1.LocalPort = 0
   Winsock1.Protocol = sckTCPProtocol
   Winsock1.RemoteHost = SMTP
   Winsock1.RemotePort = 25
   DoEvents
   
   List1.AddItem Now & " -> SMTP:" & SMTP & "," & strData(4), 0
   
   Winsock1.Connect
   If Not Response("220") Then
      Winsock1.Close
      iErrCode = 1
      GoTo ERRORMail
   End If
   
   DoEvents
   Winsock1.SendData ("HELO " & Winsock1.LocalHostName & ".taie.com.tw" & vbCrLf)
   If Not Response("250") Then
      iErrCode = 2
      GoTo ERRORMail
   End If
     
   DoEvents
   Winsock1.SendData (strData(1))
   If Not Response("250") Then
      iErrCode = 3
      GoTo ERRORMail
   End If
   
   DoEvents
   'Modified by Morgan 2019/10/18 §ï¥i¦h¦¬¥óªÌ
   strData(2) = ""
   For MailCnt = 0 To UBound(ArrTmpMail)
      If ArrTmpMail(MailCnt) <> "" Then
         If InStr(ArrTmpMail(MailCnt), "@") > 0 Then
            strData(2) = "rcpt to: <" & ArrTmpMail(MailCnt) & ">" & vbCrLf
         Else
            strData(2) = "rcpt to: <" & Trim(ArrTmpMail(MailCnt)) & "@taie.com.tw>" & vbCrLf
         End If
         If txtMailTo <> "" Then
            If InStr(txtMailTo, "@") > 0 Then
               strData(2) = "rcpt to: <" & txtMailTo & ">" & vbCrLf
            Else
               strData(2) = "rcpt to: <" & txtMailTo & "@taie.com.tw>" & vbCrLf
            End If
         End If
         Winsock1.SendData (strData(2))
         If Not Response("250") Then
            iErrCode = 4
            GoTo ERRORMail
         End If
         
      End If
   Next
   'end 2019/10/18
   
   DoEvents
   Winsock1.SendData ("data" + vbCrLf)
   If Not Response("354") Then
      iErrCode = 5
      GoTo ERRORMail
   End If
   
   DoEvents
   Winsock1.SendData (strData(0) & vbCrLf & "." & vbCrLf)
   If Not Response("250") Then
      iErrCode = 6
      GoTo ERRORMail
   End If
   
   DoEvents
   Winsock1.SendData ("quit" + vbCrLf)
   If Not Response("221") Then
      iErrCode = 7
      GoTo ERRORMail
   End If
   Winsock1.Close
   SendXMail = True
   Exit Function

ERRORMail:
   iRetry = iRetry + 1
   If iRetry < 3 Then
      GoTo RetryPoint
   End If
   
ErrHnd:
   

End Function

Private Sub AlertMsg(p_Msg As String)
   MsgBox p_Msg
End Sub

'Àx¦s¼Ë¥»ÀÉ
Private Function SaveFile() As Boolean
   Dim lngMS01 As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte
   Dim lngSize As Long 'ÀÉ®×¤j¤p
   Dim stSQL As String
   Dim stFilePath As String
   Dim adoRst As New ADODB.Recordset
   
On Error GoTo ErrHandle
   
   lngMS01 = 3
   stFilePath = PUB_Getdesktop & "\Sample.eml"
   iFileNo = FreeFile
   Open stFilePath For Binary Access Read As #iFileNo
   lngSize = LOF(iFileNo)
   ReDim bytes(lngSize)
   Get #iFileNo, , bytes()

   cnnConnection.BeginTrans
   
On Error GoTo ErrHandle1

   cnnConnection.Execute "delete MailScheduleTemplet where mst01=3", intI
   
   stSQL = "select * from MailScheduleTemplet where rownum<1"
   If adoRst.State <> adStateClosed Then adoRst.Close
   With adoRst
   .CursorLocation = adUseClient
   .Open stSQL, cnnConnection, adOpenStatic, adLockOptimistic
   .AddNew
   .Fields("mst01").Value = lngMS01
   .Fields("mst02").Value = lngSize
   .Fields("mst03").AppendChunk bytes()
   .UPDATE
   End With
   cnnConnection.CommitTrans
   
   SaveFile = True
   Exit Function
              
ErrHandle:
   cnnConnection.RollbackTrans
   
ErrHandle1:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Function

'¼g°O¿ý
Private Function WLog(oStrLog As String, Optional iFile As Integer = 0)
   Dim ffa As Integer
   ffa = FreeFile
   If iFile = 1 Then
      Open App.path & "\" & App.EXEName & "X.log" For Append As ffa
   Else
      Open App.path & "\" & App.EXEName & ".log" For Append As ffa
   End If
   Print #ffa, Trim(Now) & "  ==>  " & oStrLog
   Close ffa
End Function

Private Function Response(RCode$, Optional IsShow As Boolean = True) As Boolean
   
   Const TimeOut% = 20
   Sec = 0
   Timer1.Interval = 500
   Timer1.Enabled = True
   Response = True
  
   Do While Left$(Result, 3) <> RCode
      '¦¬¥óªÌ³Q©Ú504,Unsupport Option 555
      If Left(Result, 3) = "504" Or Left(Result, 3) = "555" Then
         Response = False
         Exit Do
      End If
      DoEvents
      If Sec > TimeOut * 2 Then
         Response = False
         Exit Do
      End If
   Loop
   Result = ""
   Timer1.Enabled = False
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040149 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   If Index = 1 Then
      Option2(0).Enabled = False
      Option2(1).Value = True
   Else
      Option2(0).Enabled = True
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Winsock1.GetData Result, vbString
    List1.AddItem Now & " -> " & Result, 0
End Sub

Private Sub Timer1_Timer()
  Sec = Sec + 1
  DoEvents
End Sub

Private Function UpdateSchedule(SM01 As String, SM10 As String) As Boolean
   Dim stSQL As String, intR As Integer
On Error GoTo ErrHnd

   stSQL = "update MailSchedule set ms11=to_char(sysdate,'yyyymmdd'),ms12=to_char(sysdate,'hh24miss'),ms13=" & Val(SM10) & " where ms01=" & SM01
   cnnConnection.Execute stSQL, intR '
   UpdateSchedule = True
   Exit Function
   
ErrHnd:
   AlertMsg Err.Description
End Function


Private Function GetMime(stAttPath As String, Optional bolAtt As Boolean) As String
   Const cBoundaryA As String = "Boundary_A_3435FE2_6617A_AA"
   Const cDASH2 As String = "--"
   Dim sCharset As String, sCTEnc As String
   Dim ts As TextStream
   Dim strLine As String, strPrefix As String
   Dim bStart As Boolean
   Dim stMIME As String
   Dim iPos As Integer
   
   sCharset = "charset=" & Chr$(34) & "Big5" & Chr$(34) & vbCrLf
   sCTEnc = "Content-Transfer-Encoding: quoted-printable" & vbCrLf
   
   If bolAtt = False Then
      stMIME = "MIME-Version: 1.0" & vbCrLf & _
         "Content-Type: multipart/alternative;" & vbCrLf & _
         vbTab & "boundary=" & Chr$(34) & cBoundaryA & Chr$(34) & vbCrLf & _
         "X-Mailer: Taie" & vbCrLf & vbCrLf & _
         cDASH2 & cBoundaryA & vbCrLf & "Content-Type: text/plain;" & vbCrLf & _
         vbTab & sCharset & sCTEnc & vbCrLf & _
         TextBlurb() & _
         cDASH2 & cBoundaryA & vbCrLf & "Content-Type: text/html;" & vbCrLf & _
         vbTab & sCharset & sCTEnc & vbCrLf
   End If
   
   If fso.FileExists(stAttPath) Then
      Set ts = fso.OpenTextFile(stAttPath)
      bStart = False
      Do While Not ts.AtEndOfStream
         strLine = ts.ReadLine
         If bolAtt = False Then
            '¹Ï¤£¥²±H
            If InStr(UCase(strLine), UCase("</HTML>")) > 0 Then
               stMIME = stMIME & strLine & vbCrLf
               Exit Do
            End If
            
            If bStart = False Then
               'If InStr(UCase(strLine), UCase("MIME-Version:")) > 0 Then
               iPos = InStr(UCase(strLine), UCase("<HTML>"))
               If iPos > 0 Then
                  bStart = True
                  '«e­±¥i¯à·|¦³µù¸Ñ,­n©¿²¤
                  stMIME = stMIME & Mid(strLine, iPos) & vbCrLf
               End If
            Else
               stMIME = stMIME & strLine & vbCrLf
            End If
         Else
            If bStart = False Then
               iPos = InStr(UCase(strLine), UCase("MIME-Version"))
               If iPos > 0 Then
                  bStart = True
                  stMIME = strLine & vbCrLf
               End If
            Else
               stMIME = stMIME & strLine & vbCrLf
            End If
         End If
      Loop
      ts.Close
   End If
   
   stMIME = stMIME & cDASH2 & cBoundaryA & cDASH2 & vbCrLf
   GetMime = stMIME
End Function


Private Function Process981215() As Boolean
   Dim stSQL As String, stVTB As String, lngMS01 As Long, lngRec As Long, ii As Integer
   Dim arrToMail
   
   lngMS01 = 19
   
   stVTB = "select distinct msd02 from (" & _
      " select fa16 MSD02 from fagent where fa02='0' and instr(fa16,'@')>0 and fa69 is null and fa02='0' and fa76<'C'" & _
      " Union select fa80 MSD02 from fagent where fa02='0' and instr(fa80,'@')>0 and fa69 is null and fa02='0' and fa76<'C'" & _
      " Union select fa81 MSD02 from fagent where fa02='0' and instr(fa81,'@')>0 and fa69 is null and fa02='0' and fa76<'C'" & _
      " Union select fa82 MSD02 from fagent where fa02='0' and instr(fa82,'@')>0 and fa69 is null and fa02='0' and fa76<'C'" & _
      " Union select pcc08 MSD02 from fagent,potcustcont where fa02='0' and fa69 is null and fa02='0' and fa76<'C' and pcc01(+)=fa01 and instr(pcc08,'@')>0" & _
      " union select pcu18 from potcustomer where pcu02='0' and pcu39 is null and instr(pcu18,'@')>0" & _
      " Union select pcc08 from potcustomer,potcustcont where pcu02='0' and pcu39 is null and pcc01(+)=pcu01 and instr(pcc08,'@')>0" & _
      " ) X,(select msd02 Y1 from mailscheduledetail where msd01=19) Y" & _
      " where y1(+)=msd02 and y1 is null"
   
   stSQL = "insert into MailScheduleDetail(MSD01,MSD02,MSD06)" & _
      " SELECT " & lngMS01 & ",MSD02,'None' FROM (" & stVTB & ") X"
   
   cnnConnection.BeginTrans

On Error GoTo ErrHnd

   cnnConnection.Execute stSQL, lngRec
   
   '¦h­Ó«H½c©ñ¤@°_ªº¸ê®Æ
   stSQL = "select msd02,msd07 from MailScheduleDetail where msd01=" & lngMS01 & " and instr(msd02,';')>0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         arrToMail = Split("" & .Fields(0), ";") '¥h°£«e«áªºªÅ¥Õ©M¸õ¦æ²Å¸¹
         For ii = 0 To UBound(arrToMail)
            arrToMail(ii) = Trim(Replace(arrToMail(ii), vbCrLf, ""))
            If InStr(arrToMail(ii), "@") > 0 Then
               stSQL = "select msd02 from MailScheduleDetail where msd01=" & lngMS01 & " and msd02='" & arrToMail(ii) & "'"
               intI = 1
               Set AdoRecordSet3 = ClsLawReadRstMsg(intI, stSQL)
               If intI = 0 Then
                  stSQL = " insert into MailScheduleDetail(MSD01,MSD02,MSD06,MSD07) values(" & lngMS01 & ",'" & arrToMail(ii) & "','None','" & RsTemp("msd07") & "')"
                  cnnConnection.Execute stSQL, intI
               End If
            End If
         Next
         .MoveNext
      Loop
      End With
      stSQL = "delete from MailScheduleDetail where msd01=" & lngMS01 & " and instr(msd02,';')>0"
      cnnConnection.Execute stSQL, lngRec
   End If
   
   stSQL = "update MailSchedule set ms10=(select count(*) from MailScheduledetail where msd01=ms01) where ms01=" & lngMS01
   cnnConnection.Execute stSQL, lngRec
   
   cnnConnection.CommitTrans
   Process981215 = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
   
End Function

Private Function Process990120(ByVal lngMS01 As Long) As Boolean
   Dim stSQL As String, stVTB As String, lngRec As Long, ii As Integer
   Dim stCon As String
   Dim arrToMail
   
   '¤é¤å
   If lngMS01 = 23 Then
      stCon = " and FA31='3'"
   '¨ä¥L
   Else
      stCon = " and (FA31 is null or FA31<>'3')"
   End If
   
   stVTB = "select fa16 MSD02,FA01||FA02 MSD07 from fagent where instr(fa16,'@')>0" & _
      " and fa02='0' and fa69 is null and fa76<'C'" & stCon & _
      " Union select fa80 MSD02,FA01||FA02 MSD07 from fagent where instr(fa80,'@')>0" & _
      " and fa02='0' and fa69 is null and fa76<'C'" & stCon & _
      " Union select fa81 MSD02,FA01||FA02 MSD07 from fagent where instr(fa81,'@')>0" & _
      " and fa02='0' and fa69 is null and fa76<'C'" & stCon & _
      " Union select fa82 MSD02,FA01||FA02 MSD07 from fagent where instr(fa82,'@')>0" & _
      " and fa02='0' and fa69 is null and fa76<'C'" & stCon & _
      " Union select pcc08 MSD02,FA01||FA02||'-'||PCC02 MSD07" & _
      " from fagent,potcustcont where  pcc01(+)=fa01 and instr(pcc08,'@')>0" & _
      " and fa02='0' and fa69 is null and fa76<'C'" & stCon
   
   stSQL = "insert into MailScheduleDetail(MSD01,MSD02,MSD06,MSD07)" & _
      " SELECT " & lngMS01 & ",MSD02,'None',min(MSD07) FROM (" & stVTB & ") X group by MSD02"
   
   cnnConnection.BeginTrans

On Error GoTo ErrHnd

   cnnConnection.Execute stSQL, lngRec
   
   '¦h­Ó«H½c©ñ¤@°_ªº¸ê®Æ
   stSQL = "select msd02,msd07 from MailScheduleDetail where msd01=" & lngMS01 & " and instr(msd02,';')>0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         arrToMail = Split("" & .Fields(0), ";") '¥h°£«e«áªºªÅ¥Õ©M¸õ¦æ²Å¸¹
         For ii = 0 To UBound(arrToMail)
            arrToMail(ii) = Trim(Replace(arrToMail(ii), vbCrLf, ""))
            If InStr(arrToMail(ii), "@") > 0 Then
               stSQL = "select msd02 from MailScheduleDetail where msd01=" & lngMS01 & " and msd02='" & arrToMail(ii) & "'"
               intI = 1
               Set AdoRecordSet3 = ClsLawReadRstMsg(intI, stSQL)
               If intI = 0 Then
                  stSQL = " insert into MailScheduleDetail(MSD01,MSD02,MSD06,MSD07) values(" & lngMS01 & ",'" & arrToMail(ii) & "','None','" & RsTemp("msd07") & "')"
                  cnnConnection.Execute stSQL, intI
               End If
            End If
         Next
         .MoveNext
      Loop
      End With
      stSQL = "delete from MailScheduleDetail where msd01=" & lngMS01 & " and instr(msd02,';')>0"
      cnnConnection.Execute stSQL, lngRec
   End If
   
   stSQL = "update MailSchedule set ms10=(select count(*) from MailScheduledetail where msd01=ms01) where ms01=" & lngMS01
   cnnConnection.Execute stSQL, lngRec
   
   cnnConnection.CommitTrans
   Process990120 = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
   
End Function

'Add by Morgan 2010/1/28
Private Function Process990128(ByVal lngMS01 As Long) As Boolean
   Dim stSQL As String, stVTB As String, lngRec As Long, ii As Integer
   Dim stCon As String
   Dim arrToMail
      
   stVTB = "select T03 MSD02,ST07||' '||T01 MSD07 from T20100128,fagent,nation,staff" & _
      " where instr(T03,'@')>0 and fa01(+)=substr(T01,1,8) and fa02(+)=substr(T01,9)" & _
      " and na01(+)=fa10 and st01(+)=na51"
   
   stSQL = "insert into MailScheduleDetail(MSD01,MSD02,MSD06,MSD07)" & _
      " SELECT " & lngMS01 & ",MSD02,'None',min(MSD07) FROM (" & stVTB & ") X group by MSD02"
   
   cnnConnection.BeginTrans

On Error GoTo ErrHnd

   cnnConnection.Execute stSQL, lngRec
   
   '¦h­Ó«H½c©ñ¤@°_ªº¸ê®Æ
   stSQL = "select msd02,msd07 from MailScheduleDetail where msd01=" & lngMS01 & " and instr(msd02,';')>0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         arrToMail = Split("" & .Fields(0), ";") '¥h°£«e«áªºªÅ¥Õ©M¸õ¦æ²Å¸¹
         For ii = 0 To UBound(arrToMail)
            arrToMail(ii) = Trim(Replace(arrToMail(ii), vbCrLf, ""))
            If InStr(arrToMail(ii), "@") > 0 Then
               stSQL = "select msd02 from MailScheduleDetail where msd01=" & lngMS01 & " and msd02='" & arrToMail(ii) & "'"
               intI = 1
               Set AdoRecordSet3 = ClsLawReadRstMsg(intI, stSQL)
               If intI = 0 Then
                  stSQL = " insert into MailScheduleDetail(MSD01,MSD02,MSD06,MSD07) values(" & lngMS01 & ",'" & arrToMail(ii) & "','None','" & RsTemp("msd07") & "')"
                  cnnConnection.Execute stSQL, intI
               End If
            End If
         Next
         .MoveNext
      Loop
      End With
      stSQL = "delete from MailScheduleDetail where msd01=" & lngMS01 & " and instr(msd02,';')>0"
      cnnConnection.Execute stSQL, lngRec
   End If
   
   stSQL = "update MailSchedule set ms10=(select count(*) from MailScheduledetail where msd01=ms01) where ms01=" & lngMS01
   cnnConnection.Execute stSQL, lngRec
   
   cnnConnection.CommitTrans
   Process990128 = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
   
End Function

'Added by Morgan 2019/10/16
'Åª¨ú¶l¥ó½d¥»
Private Function GetTemplete(p_MST01 As String, p_TempletPath As String) As Boolean
   strExc(0) = "select * from mailscheduleTemplet b where mst01=" & p_MST01
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Dir(p_TempletPath) <> "" Then Kill p_TempletPath
      GetTemplete = PUB_GetFtpFile(RsTemp.Fields("mst06"), p_TempletPath, UCase("MAILSCHEDULETEMPLET"))
   End If
   
End Function

Private Function GetSMTP(Optional pByMailServer As Boolean = False) As String
   Dim stSQL As String, intR As Integer
   
   If pByMailServer Then
      stSQL = "select oMan from setSpecMan where ocode='SMTP_IP_MS'"
   Else
      stSQL = "select oMan from setSpecMan where ocode='SMTP_IP_FW'"
   End If
   intR = 1
   Set RsTemp = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      GetSMTP = RsTemp(0)
   End If
End Function

Private Function FixFirstDot(pText As String) As String
   Dim arr() As String, ii As Integer, stNew As String
   arr = Split(pText, vbCrLf)
   stNew = ""
   For ii = LBound(arr) To UBound(arr)
      If Left(arr(ii), 1) = "." Then
         arr(ii) = "." & arr(ii)
      End If
      stNew = stNew & IIf(stNew = "", "", vbCrLf) & arr(ii)
   Next
   FixFirstDot = stNew
End Function
