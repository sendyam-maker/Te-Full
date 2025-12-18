VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHTA850 
   Caption         =   "HTA850VB6DEMO"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   10665
   StartUpPosition =   2  '螢幕中央
   Begin VB.ComboBox Combo2 
      Height          =   300
      Index           =   0
      ItemData        =   "frmHTA850.frx":0000
      Left            =   1800
      List            =   "frmHTA850.frx":0007
      TabIndex        =   24
      Text            =   "4660"
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frmHTA850.frx":0011
      Left            =   120
      List            =   "frmHTA850.frx":0018
      TabIndex        =   23
      Text            =   "192.168.0.1"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtELid 
      Height          =   270
      Left            =   9360
      TabIndex        =   22
      Text            =   "1"
      Top             =   2370
      Visible         =   0   'False
      Width           =   615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Basic"
      TabPicture(0)   =   "frmHTA850.frx":002B
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Line5"
      Tab(0).Control(1)=   "Line7"
      Tab(0).Control(2)=   "btninitialgcu"
      Tab(0).Control(3)=   "btngetgcuinfo"
      Tab(0).Control(4)=   "btnwritegcutime"
      Tab(0).Control(5)=   "btnreadgcutime"
      Tab(0).Control(6)=   "Command4"
      Tab(0).Control(7)=   "Command6"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Table"
      TabPicture(1)   =   "frmHTA850.frx":0047
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command7"
      Tab(1).Control(1)=   "Command8"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Card"
      TabPicture(2)   =   "frmHTA850.frx":0063
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(2)=   "btninsertmusrecord"
      Tab(2).Control(3)=   "btndeluserRecord"
      Tab(2).Control(4)=   "btnqueryuserR"
      Tab(2).Control(5)=   "btndelalluserR"
      Tab(2).Control(6)=   "txtCardNo"
      Tab(2).Control(7)=   "txtDisplayName"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Polling"
      TabPicture(3)   =   "frmHTA850.frx":007F
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "btnclearlist"
      Tab(3).Control(1)=   "List1"
      Tab(3).Control(2)=   "btngetpolldata"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "EEPRom"
      TabPicture(4)   =   "frmHTA850.frx":009B
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command2"
      Tab(4).Control(1)=   "Command1"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "MifareParams"
      TabPicture(5)   =   "frmHTA850.frx":00B7
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Command3"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Fingerprint"
      TabPicture(6)   =   "frmHTA850.frx":00D3
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Label3"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Label8"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Label9"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Label10"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "Text1"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "Command5"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "Command9"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "List2"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "Text2"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).Control(9)=   "Command10"
      Tab(6).Control(9).Enabled=   0   'False
      Tab(6).Control(10)=   "Command11"
      Tab(6).Control(10).Enabled=   0   'False
      Tab(6).Control(11)=   "Text3"
      Tab(6).Control(11).Enabled=   0   'False
      Tab(6).Control(12)=   "Text4"
      Tab(6).Control(12).Enabled=   0   'False
      Tab(6).Control(13)=   "Command12"
      Tab(6).Control(13).Enabled=   0   'False
      Tab(6).Control(14)=   "Command13"
      Tab(6).Control(14).Enabled=   0   'False
      Tab(6).ControlCount=   15
      Begin VB.CommandButton Command13 
         Caption         =   "ClearFG"
         Height          =   345
         Left            =   8640
         TabIndex        =   48
         Top             =   2370
         Width           =   1050
      End
      Begin VB.CommandButton Command12 
         Caption         =   "ClearList"
         Height          =   345
         Left            =   8640
         TabIndex        =   47
         Top             =   1950
         Width           =   1050
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   900
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   44
         Top             =   5310
         Width           =   8835
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   900
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   43
         Top             =   4470
         Width           =   8835
      End
      Begin VB.CommandButton Command11 
         Caption         =   "hsHTA850QueryMasterFP"
         Height          =   400
         Left            =   1305
         TabIndex        =   42
         Top             =   2340
         Width           =   3375
      End
      Begin VB.CommandButton Command10 
         Caption         =   "hsHTA850UpdateMasterFP"
         Height          =   400
         Left            =   4860
         TabIndex        =   41
         Top             =   2340
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   6570
         TabIndex        =   39
         Top             =   1320
         Width           =   2085
      End
      Begin VB.TextBox txtDisplayName 
         Height          =   315
         Left            =   -70005
         TabIndex        =   37
         Top             =   1980
         Width           =   2085
      End
      Begin VB.ListBox List2 
         Height          =   1500
         ItemData        =   "frmHTA850.frx":00EF
         Left            =   855
         List            =   "frmHTA850.frx":00F1
         TabIndex        =   36
         Top             =   2880
         Width           =   8895
      End
      Begin VB.CommandButton Command9 
         Caption         =   "hsHTA850InsertMultiUserFingerPrinter"
         Height          =   400
         Left            =   4860
         TabIndex        =   34
         Top             =   1920
         Width           =   3375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "hsHTA850QueryUserFingerPrinter"
         Height          =   400
         Left            =   1320
         TabIndex        =   33
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   2280
         TabIndex        =   32
         Text            =   "11"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "hsHTA850SetMifareReader"
         Height          =   495
         Left            =   -74880
         TabIndex        =   31
         Top             =   840
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "hsHTA850SetEEPROM"
         Height          =   495
         Left            =   -74760
         TabIndex        =   30
         Top             =   1320
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Caption         =   "hsHTA850ReadEEPROM"
         Height          =   495
         Left            =   -74760
         TabIndex        =   29
         Top             =   2280
         Width           =   3015
      End
      Begin VB.CommandButton Command7 
         Caption         =   "hsHTA850WriteTable"
         Height          =   375
         Left            =   -74520
         TabIndex        =   28
         Top             =   2100
         Width           =   2055
      End
      Begin VB.CommandButton Command8 
         Caption         =   "hsHTA850ReadTable"
         Height          =   375
         Left            =   -74520
         TabIndex        =   27
         Top             =   1500
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
         Caption         =   "hsHTA850WriteParameter"
         Height          =   375
         Left            =   -69720
         TabIndex        =   26
         Top             =   2100
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "hsHTA850ReadParameter"
         Height          =   375
         Left            =   -72120
         TabIndex        =   25
         Top             =   2100
         Width           =   2055
      End
      Begin VB.CommandButton btnclearlist 
         Caption         =   "Clear list"
         Height          =   400
         Left            =   -69360
         TabIndex        =   19
         Top             =   1020
         Width           =   2055
      End
      Begin VB.ListBox List1 
         Height          =   4920
         ItemData        =   "frmHTA850.frx":00F3
         Left            =   -74880
         List            =   "frmHTA850.frx":00F5
         TabIndex        =   18
         Top             =   1860
         Width           =   10215
      End
      Begin VB.CommandButton btngetpolldata 
         Caption         =   "hsHTA850PollingData"
         Height          =   400
         Left            =   -72480
         TabIndex        =   17
         Top             =   1020
         Width           =   2055
      End
      Begin VB.TextBox txtCardNo 
         Height          =   270
         Left            =   -74400
         TabIndex        =   15
         Text            =   "0000135724"
         Top             =   1500
         Width           =   2055
      End
      Begin VB.CommandButton btndelalluserR 
         Caption         =   "hsHTA850DeleteAllUserRecord"
         Height          =   400
         Left            =   -74520
         TabIndex        =   14
         Top             =   4320
         Width           =   2655
      End
      Begin VB.CommandButton btnqueryuserR 
         Caption         =   "hsHTA850QueryUserRecord"
         Height          =   400
         Left            =   -74520
         TabIndex        =   13
         Top             =   2640
         Width           =   2655
      End
      Begin VB.CommandButton btndeluserRecord 
         Caption         =   "hsHTA850DeleteUserRecord"
         Height          =   400
         Left            =   -74520
         TabIndex        =   12
         Top             =   3480
         Width           =   2655
      End
      Begin VB.CommandButton btninsertmusrecord 
         Caption         =   "hsHTA850InsertMultiUserRecord"
         Height          =   400
         Left            =   -74520
         TabIndex        =   11
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CommandButton btnreadgcutime 
         Caption         =   "hsHTA850ReadTime"
         Height          =   400
         Left            =   -72120
         TabIndex        =   10
         Top             =   1140
         Width           =   2055
      End
      Begin VB.CommandButton btnwritegcutime 
         Caption         =   "hsHTA850WriteTime"
         Height          =   400
         Left            =   -74520
         TabIndex        =   9
         Top             =   1140
         Width           =   2055
      End
      Begin VB.CommandButton btngetgcuinfo 
         Caption         =   "hsHTA850GetInfo"
         Height          =   400
         Left            =   -74520
         TabIndex        =   8
         Top             =   2100
         Width           =   2055
      End
      Begin VB.CommandButton btninitialgcu 
         Caption         =   "hsHTA850Initial"
         Height          =   400
         Left            =   -69720
         TabIndex        =   7
         Top             =   1140
         Width           =   2055
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "FG2:"
         Height          =   180
         Left            =   495
         TabIndex        =   46
         Top             =   5370
         Width           =   345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "FG1:"
         Height          =   180
         Left            =   495
         TabIndex        =   45
         Top             =   4530
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Display Message:"
         Height          =   180
         Left            =   5220
         TabIndex        =   40
         Top             =   1380
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Display Message:"
         Height          =   180
         Left            =   -71400
         TabIndex        =   38
         Top             =   2010
         Width           =   1200
      End
      Begin VB.Label Label3 
         Caption         =   "Card NO."
         Height          =   255
         Left            =   840
         TabIndex        =   35
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Line Line7 
         X1              =   -74640
         X2              =   -65160
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Label Label4 
         Caption         =   "Card NO."
         Height          =   255
         Left            =   -74280
         TabIndex        =   16
         Top             =   1020
         Width           =   1695
      End
      Begin VB.Line Line5 
         X1              =   -74640
         X2              =   -65160
         Y1              =   1860
         Y2              =   1860
      End
   End
   Begin VB.TextBox txtreturn 
      Height          =   270
      Left            =   3960
      TabIndex        =   3
      Top             =   465
      Width           =   6495
   End
   Begin VB.CommandButton btnclose 
      Caption         =   "Close"
      Height          =   400
      Left            =   8400
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton btndisconnect 
      Caption         =   "hsCloseChannel"
      Height          =   400
      Left            =   2280
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton btnconnect 
      Caption         =   "hsOpenChannel"
      Height          =   400
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Gcu ID"
      Height          =   255
      Left            =   6600
      TabIndex        =   21
      Top             =   4680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Rate/iPort"
      Height          =   255
      Left            =   1800
      TabIndex        =   20
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Dll Return"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Com NO./Ip"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmHTA850"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/****************************************************************************
'* Name ........... RAC2000EL SDK Sample
'* Parameter.......
'* Author ......... Yi-Pin Wang (Raphael Wang)
'* Date ........... 2007/02/10
'* Company ........ HUNDURE TECHNOLOGY CO.,LTD
'****************************************************************************/
 
 'Dim sFPtData1(122) As Byte
 'Dim sFPtData2(122) As Byte
 Dim sFPtData1(386) As Byte
 Dim sFPtData2(386) As Byte


Private Sub btnclearlist_Click()
   List1.Clear
End Sub

Private Sub btnclose_Click()
    If ghComm > 0 Then
       ireturn = hsCloseChannel(ghComm)
    End If
    End
End Sub

Private Sub btnconnect_Click()
        
  Dim ireturn, iport, iELID As Integer
  
  iport = Int(Trim(Combo2(0).Text))
  txtreturn.Text = ""
  

    ' Open the communication channel
    
    ireturn = hsOpenChannel(ghComm, Trim(Combo1.Text), iport)
    If ireturn = 0 Then
      txtreturn.Text = "ip connect: ok!"
    Else
      txtreturn.Text = "ip connect: error!(" & ireturn & ")"
    End If

    
End Sub

Private Sub btnconnect_Click_Click()

End Sub

Private Sub btndelalluserR_Click()
        Dim ireturn As Integer
        Dim ireturncode As Integer
        ireturn = 0
        ireturncode = 0
        txtreturn.Text = ""
        ireturn = hsHTA850DeleteAllUserRecord(ghComm, ireturncode, 1000)
        If ireturn = 0 Then
            txtreturn.Text = "hsHTA850DeleteAllUserRecord : ok! "
        Else
            txtreturn.Text = "hsHTA850DeleteAllUserRecord : error!(" & ireturn & ") Error Number:" & Str(ireturncode)
        End If
    
End Sub

Private Sub btndeluserRecord_Click()
    
        Dim i As Integer
        Dim ireturn As Integer
        Dim ireturncode As Integer
        Dim sCardFormatData(255) As Byte
        Dim stRecord(255) As Byte
        Dim Card As String

        Card = txtCardNo.Text
        i = 0
        For i = 0 To 255
            stRecord(i) = 0
        Next
                
        For i = 0 To Len(Card) - 1 '.Length - 1
            stRecord(i) = Asc(Mid(Card, i + 1, 1))
        Next

        ireturn = hsHTA850DeleteUserRecord(ghComm, 16, stRecord(0), ireturncode, 1000)
        If ireturn = 0 Then
            txtreturn.Text = "hsHTA850DeleteUserRecord : ok!"
        Else
            txtreturn.Text = "hsHTA850DeleteUserRecord : error!(" & ireturn & ") Error Number:" & Str(ireturncode)
        End If
    
End Sub

Private Sub btndisconnect_Click()
        Dim ireturn As Integer
        ireturn = 0
        txtreturn.Text = ""
        ' Close the communication channel
        ireturn = hsCloseChannel(ghComm)
        If ireturn = 0 Then
            txtreturn.Text = "Close channel: ok!"
        Else
            txtreturn.Text = "Close channel: error!(" & ireturn & ")"
        End If
   
End Sub


Private Sub btngetgcuinfo_Click()
        Dim ireturncode, ireturn As Integer
        Dim iinfolen As Integer
        Dim i As Integer
        Dim j As Integer
        Dim sinfodata As String
        Dim sshow As String
        Dim sstr1 As String
        Dim nInfo(255) As Byte

        ireturn = 0
        ireturncode = 0
        iinfolen = 0
        txtreturn.Text = ""
        sshow = ""
        ireturn = hsHTA850GetInfo(ghComm, nInfo(0), iinfolen, ireturncode, 1000)
        If ireturn = 0 Then

            For i = 1 To iinfolen

                j = nInfo(i - 1)
                sstr1 = Hex(j)
                If Len(sstr1) < 2 Then
                    sstr1 = "0" & sstr1
                ElseIf Len(sstr1) > 2 Then
                    sstr1 = Mid(sstr1, 1, 2) & " " & Mid(sstr1, 3, 2)
                End If
                sshow = sshow & sstr1 & " "
            Next i
            txtreturn.Text = "hsHTA850GetInfo: ok! " & sshow
        Else
            txtreturn.Text = "hsHTA850GetInfo: error!(" & ireturn & ") Error Number:" & Str(ireturncode)
        End If
    
End Sub


'/////////////**************************

    '/////////////**************************
    Function polling(ByVal iELID As Integer, ByVal icardtype As Integer) As Boolean
        Dim ireturn As Integer
        Dim ireturncode As Integer
        Dim iRecord As Integer
        Dim i As Integer
        Dim j As Integer
        Dim flag As Boolean
        Dim rByte(1024) As Byte
        Dim k, iRecCount As Integer
        Dim stRecord As String

        Dim sEventCode As String
        Dim sDateTime As String
        Dim sReaderNo As String
        Dim sInputType As String
        Dim sSection As String
        Dim sClass As String

        Dim sCard As String
        Dim gStr As String
        Dim sshow As String
        Dim AEvent As TEvent

        ireturn = 0
        ireturncode = 0
        iRecord = 0
        txtreturn.Text = ""
        iRecCount = 0
        ireturn = hsHTA850PollingData(ghComm, iRecCount, rByte(0), iRecord, ireturncode, 1000)
        If ireturn = 0 Then

            'iPrevRe(iELID) = iRecord



            '   char cDate[10];
            'char cTime[10];
            ' char Reader;
            'char InputType;
            'char ASection;
            'char AClass;
            'char EventCode;
            'char Card[16];
            '/////////////////////
            k = 0
            For i = 0 To iRecord - 1
                sshow = ""
                sEvnetCode = "Event Date: "
                gStr = ""
                For j = 1 To 10
                    If (rByte(k) <> 0) Then
                        sEvnetCode = sEvnetCode & Chr(rByte(k))
                    End If
                    k = k + 1
                Next j
                AEvent.sDate = sEvnetCode

                sDateTime = "Date Time: "
                For j = 1 To 10
                    If (rByte(k) <> 0) Then
                        sDateTime = sDateTime & Chr(rByte(k))
                    End If

                    k = k + 1
                Next j

                AEvent.sTime = sDateTime

                sReaderNo = "Reader NO.: "
                sReaderNo = sReaderNo & rByte(k)
                AEvent.Reader = rByte(k)
                k = k + 1



                sInputType = "Input Type: "
                sInputType = sInputType & rByte(k)
                AEvent.InputType = rByte(k)
                k = k + 1

                sSection = "Section: "
                sSection = sSection & rByte(k)
                AEvent.ASection = rByte(k)
                k = k + 1

                sClass = "Class: "
                sClass = sClass & rByte(k)
                AEvent.AClass = rByte(k)
                k = k + 1

                sEventCode = "Event Code: "
                sEventCode = sEventCode & rByte(k)
                AEvent.EventCode = rByte(k)
                k = k + 1

                sCard = ""
                For j = 1 To 16
                    If (rByte(k) <> 0) Then
                        sCard = sCard & Chr(rByte(k))
                    End If

                    k = k + 1
                Next j
                AEvent.Card = sCard
                sshow = sEvnetCode & " " & sDateTime & " " & sCard & " " & sReaderNo & " " & sSection & " " & sInputType & " " & sClass & " " & sEventCode
                List1.AddItem sshow

            Next i
            flag = True
        ElseIf ghComm <> 0 Then
            If ireturn = 1010 Then
                txtreturn.Text = "The HTA850" & Trim(Str(iELID)) & " has not data!"
                flag = True
                iRecCount = 0
            Else
                txtreturn.Text = "Polling the HTA850" & Trim(Str(iELID)) & " is failure!"
                iRecCount = 0
            End If
        Else
            txtreturn.Text = "Polling the HTA850" & Trim(Str(iELID)) & " is failure!"
            iRecCount = 0
            flag = False

        End If
        polling = flag
        iRecCount = iRecord
        While iRecCount > 0
            ireturn = hsHTA850PollingData(ghComm, iRecCount, rByte(0), iRecord, ireturncode, 1000)
            If ireturn = 0 Then

                'iPrevRe(iELID) = iRecord
                iRecCount = iRecord


                '   char cDate[10];
                'char cTime[10];
                ' char Reader;
                'char InputType;
                'char ASection;
                'char AClass;
                'char EventCode;
                'char Card[16];
                '/////////////////////
                k = 0
                For i = 0 To iRecord - 1
                    sshow = ""
                    sEvnetCode = "Event Date: "
                    gStr = ""
                    For j = 1 To 10
                        If (rByte(k) <> 0) Then
                            sEvnetCode = sEvnetCode & Chr(rByte(k))
                        End If
                        k = k + 1
                    Next j
                    AEvent.sDate = sEvnetCode

                    sDateTime = "Date Time: "
                    For j = 1 To 10
                        If (rByte(k) <> 0) Then
                            sDateTime = sDateTime & Chr(rByte(k))
                        End If

                        k = k + 1
                    Next j

                    AEvent.sTime = sDateTime

                    sReaderNo = "Reader NO.: "
                    sReaderNo = sReaderNo & rByte(k)
                    AEvent.Reader = rByte(k)
                    k = k + 1



                    sInputType = "Input Type: "
                    sInputType = sInputType & rByte(k)
                    AEvent.InputType = rByte(k)
                    k = k + 1

                    sSection = "Section: "
                    sSection = sSection & rByte(k)
                    AEvent.ASection = rByte(k)
                    k = k + 1

                    sClass = "Class: "
                    sClass = sClass & rByte(k)
                    AEvent.AClass = rByte(k)
                    k = k + 1

                    sEventCode = "Event Code: "
                    sEventCode = sEventCode & rByte(k)
                    AEvent.EventCode = rByte(k)
                    k = k + 1

                    sCard = ""
                    For j = 1 To 16
                        If (rByte(k) <> 0) Then
                            sCard = sCard & Chr(rByte(k))
                        End If

                        k = k + 1
                    Next j
                    AEvent.Card = sCard
                    sshow = sEvnetCode & " " & sDateTime & " " & sCard & " " & sReaderNo & " " & sSection & " " & sInputType & " " & sClass & " " & sEventCode

                    List1.AddItem sshow

                Next i
                flag = True
            ElseIf ghComm <> 0 Then
                If ireturn = 1010 Then
                    txtreturn.Text = "The HTA850" & Trim(Str(iELID)) & " has not data!"
                    flag = True
                    iRecCount = 0
                Else
                    txtreturn.Text = "Polling the HTA850" & Trim(Str(iELID)) & " is failure!"
                    iRecCount = 0
                End If
            Else
                txtreturn.Text = "Polling the HTA850" & Trim(Str(iELID)) & " is failure!"
                iRecCount = 0
                flag = False

            End If
            polling = flag
        Wend
    End Function


Private Sub btngetpolldata_Click()
    flag = polling(1, 0)
End Sub

Private Sub btninitialgcu_Click()
0   Dim ireturn, iELID, ireturncode As Integer
   Dim binitflag As Byte
   
   ireturn = 0
   ireturncode = 0
   txtreturn.Text = ""
   binitflag = 255

   ireturn = hsHTA850Initial(ghComm, binitflag, ireturncode, 5000)
   If ireturn = 0 Then
      txtreturn.Text = "hsHTA850Initialize :  OK!"
   Else
      txtreturn.Text = "hsHTA850Initialize : error!(" & ireturn & ") Error Number:" + Str(ireturncode)
   End If
      
     
End Sub

Private Sub btninsertmusrecord_Click()
        Dim bstCard() As Byte
        ReDim bstCard(5 * 224)

        Dim ireturn As Integer
        Dim ireturncode As Integer
        Dim iRecord As Integer
        Dim i As Integer, k As Integer
        Dim stRecord(255) As Byte
        Dim Card, Msg As String

        Card = txtCardNo.Text
        Msg = txtDisplayName
        i = 0
        For i = 0 To 255
            stRecord(i) = 0
        Next

        For i = 0 To Len(Card) - 1
            stRecord(i) = Asc(Mid(Card, i + 1, 1))
        Next

'        For i = 18 To 18 + Len(Msg) - 1
'            stRecord(i) = Asc(Mid(Msg, (i - 17), 1))
'        Next
        
        k = 0
        For i = 18 To 18 + Len(Msg) - 1
            If Asc(Mid(Msg, (i - 17), 1)) > 0 Then
               stRecord(i + k) = Asc(Mid(Msg, (i - 17), 1))
            Else
               stRecord(i + k) = Val("&H" & Left(Hex(Asc(Mid(Msg, (i - 17), 1))), 2))
               k = k + 1
               stRecord(i + k) = Val("&H" & Right(Hex(Asc(Mid(Msg, (i - 17), 1))), 2))
            End If
        Next

        ireturn = 0
        ireturncode = 0
        iRecord = 1



        ireturn = hsHTA850InsertMultiUserRecord(ghComm, 16, 16, 1, stRecord(0), ireturncode, 1000)
        If ireturn = 0 Then
            txtreturn.Text = "hsHTA850InsertMultiUserRecord : ok!"
        Else
            If ireturncode = 2 Then
                txtreturn.Text = "The length of card is not true!"
            Else
                txtreturn.Text = "hsHTA850InsertMultiUserRecord : error!(" & ireturn & ") Error Number:" & Str(ireturncode)
            End If
        End If
       
 End Sub

Private Sub btnmatrixwritet_Click()
  
End Sub

Private Sub btnqueryuserR_Click()
        Dim i As Integer
        Dim ireturn As Integer
        Dim ireturncode As Integer
        Dim sCardFormatData(255) As Byte
        Dim iCardFormatLen As Integer
        Dim stRecord(255) As Byte
        Dim Card As String
        Dim gStr As String

        Card = txtCardNo.Text
        i = 0
        For i = 0 To 255
            stRecord(i) = 0
        Next

        For i = 0 To Len(Card) - 1
            stRecord(i) = Asc(Mid(Card, (i + 1), 1))
        Next

        iCardFormatLen = 0

        ireturn = 0
        ireturncode = 0

        txtreturn.Text = ""
        ireturn = hsHTA850QueryUserRecord(ghComm, 16, stRecord(0), sCardFormatData(0), iCardFormatLen, ireturncode, 1000)
        If ireturn = 0 Then
            gStr = ""
            For i = 0 To iCardFormatLen - 1
                gStr = gStr & Hex(sCardFormatData(i)) & " "
            Next
            txtreturn.Text = "Query user Record: ok! " & gStr
            
            gStr = ""
            For i = 18 To iCardFormatLen - 1
               If Val(sCardFormatData(i)) < 128 Then
                  gStr = gStr & Chr(sCardFormatData(i))
               Else
                  gStr = gStr & Chr(Val("&H" & Hex(sCardFormatData(i)) & Hex(sCardFormatData(i + 1))))
                  i = i + 1
               End If
            Next
            txtDisplayName = RTrim(gStr)
        Else
            If ireturncode = 6 Then
                txtreturn.Text = "Query Card:Not Exist!"
            Else
                txtreturn.Text = "Query user Record: error!(" & ireturn & ") Error Number:" & Str(ireturncode)
            End If
        End If
      
   
End Sub

Private Sub btnreadgcutime_Click()
        Dim sDate As String
        Dim sTime As String
        Dim ireturn As Integer
        Dim ireturncode As Integer
        sDate = "                                  "
        sTime = "                    "
        
        ireturn = 0
        ireturncode = 0
        txtreturn.Text = ""
        ireturn = hsHTA850ReadTime(ghComm, sDate, sTime, ireturncode, 5000)
        If ireturn = 0 Then
            txtreturn.Text = "hsELGetTime: " & sDate
            txtreturn.Text = txtreturn.Text & " " & sTime
        Else
            txtreturn.Text = "hsELGetTime: error!(" & ireturn & ") Error Number:" & Str(ireturncode)
        End If
End Sub

 
Private Sub btnwritegcutime_Click()
        Dim ek As Integer
        Dim ireturncode, ireturn, iELID As Integer
        Dim iweek As Integer
        Dim sDate As String
        Dim sTime As String
        sDate = "                   "
        sTime = "                      "
        
        iweek = Weekday(Now)
        If iweek = 1 Then
            iweek = 7
        ElseIf ek = 7 Then
            iweek = 6
        Else
            iweek = iweek - 1
        End If
        sDate = Format(Now, "yyyymmddw")
        sDate = sDate & Trim(Str(iweek))
        sTime = Format(Now, "HHmmss")
        ireturn = 0
        ireturncode = 0

        txtreturn.Text = ""
        ireturn = hsHTA850WriteTime(ghComm, sDate, sTime, ireturncode, 500)
        If ireturn = 0 Then
            txtreturn.Text = "hsELSetTime : OK!"
        Else
            txtreturn.Text = "hsELSetTime : error!(" & ireturn & ") Error Number:" & Str(ireturncode)
        End If
       
End Sub

Private Sub btnwriteopara_Click()
    
      
End Sub

Private Sub btnwriteparameter_Click()
  
       
  
End Sub

Private Sub btnwritepollid_Click()
      
   
 
         
End Sub

Private Sub btnwritepollidTimeout_Click()
    
End Sub

Private Sub hsOpenChannel_Click()

End Sub

Private Sub Command1_Click()
        Dim ireturn, ireturncode As Integer
        Dim cdata(255) As Byte
        ireturn = 0
        cdata(0) = &H13
        cdata(1) = 0
        cdata(2) = 1
        cdata(3) = 0
        cdata(4) = 2


        txtreturn.Text = ""

        ireturn = hsHTA850SetEEPROM(ghComm, cdata(0), 5, ireturncode, 5000)
        If ireturn = 0 Then
            txtreturn.Text = "hsHTA850SetEEPROM : ok! "
        Else
            txtreturn.Text = "hsHTA850SetEEPROM : error!(" & ireturn & ") "
        End If


End Sub

Private Sub Command10_Click()
        Dim i As Integer
        Dim ireturn As Integer
        Dim ireturncode As Integer
        Dim gStr As String
        Dim newFP1, newFP2
        
        For i = 0 To 385
            sFPtData1(i) = 0
        Next
        For i = 0 To 385
            sFPtData2(i) = 0
        Next
        
        newFP1 = Split(Text3, " ")
        newFP2 = Split(Text4, " ")
        
        For i = 0 To UBound(newFP1)
            sFPtData1(i) = Val("&H" & newFP1(i))
        Next
        For i = 0 To UBound(newFP2)
            sFPtData2(i) = Val("&H" & newFP2(i))
        Next

        ireturn = 0
        ireturncode = 0

        txtreturn.Text = ""
        ireturn = hsHTA850UpdateMasterFP(ghComm, sFPtData1(0), sFPtData2(0), ireturncode, 10000)
        Text3.Text = "": Text4.Text = ""
        If ireturn = 0 Then
            txtreturn.Text = "Update Master Finger Printer: ok!"
        Else
            txtreturn.Text = "Update Master Finger Printer: error!(" & ireturn & ")" & "  Error Number:" & Str(ireturncode)
        End If

End Sub

Private Sub Command11_Click()

        Dim i As Integer
        Dim ireturn As Integer
        Dim ireturncode As Integer
        Dim gStr As String
        
        For i = 0 To 385
            sFPtData1(i) = 0
        Next
        For i = 0 To 385
            sFPtData2(i) = 0
        Next

        ireturn = 0
        ireturncode = 0

        txtreturn.Text = ""
        ireturn = hsHTA850QueryMasterFP(ghComm, sFPtData1(0), sFPtData2(0), ireturncode, 10000)
        Text3.Text = "": Text4.Text = ""
        If ireturn = 0 Then
            gStr = ""
            For i = 0 To 385
                gStr = gStr & Right("0" & Hex(sFPtData1(i)), 2) & " "
            Next
            Text3 = gStr
            
            gStr = ""
            For i = 0 To 385
                gStr = gStr & Right("0" & Hex(sFPtData2(i)), 2) & " "
            Next
            Text4 = gStr
            
        ElseIf ireturn = 2225 Then
            txtreturn.Text = "Query Master Finger Printer:未建立指紋資料!(" & ireturn & ")  Error Number:" & Str(ireturncode)
        Else
            txtreturn.Text = "Query Master Finger Printer: error!(" & ireturn & ")  Error Number:" & Str(ireturncode)
        End If

End Sub

Private Sub Command12_Click()
List2.Clear
End Sub

Private Sub Command13_Click()
   Text3.Text = ""
   Text4.Text = ""
End Sub

Private Sub Command2_Click()
        Dim ireturn, ireturncode As Integer
        Dim i, RecLen As Integer
        Dim cdata(255) As Byte
        Dim rdata(255) As Byte
        Dim gStr As String
        ireturn = 0
        cdata(0) = &H13
        cdata(1) = 0
        cdata(2) = 1
        cdata(3) = 0



        txtreturn.Text = ""

        ireturn = hsHTA850ReadEEPROM(ghComm, cdata(0), 4, rdata(0), RecLen, ireturncode, 5000)
        If ireturn = 0 Then
            gStr = ""

            For i = 0 To RecLen - 1
                gStr = gStr & Hex(rdata(i))
            Next
            txtreturn.Text = "hsHTA850ReadEEPROM : ok! " & gStr
        Else
            txtreturn.Text = "hsHTA850ReadEEPROM : error!(" & ireturn & ")"
        End If
End Sub

Private Sub Command3_Click()
        Dim iparalen As Integer
        Dim ireturncode, ireturn As Integer
        Dim cParaData(30) As Byte
        

        ireturn = 0
        ireturncode = 0
        txtreturn.Text = ""

        cParaData(0) = &H70
        cParaData(1) = 0
        cParaData(2) = 0
        cParaData(3) = 0
        cParaData(4) = &H80
        cParaData(5) = 255
        cParaData(6) = 255
        cParaData(7) = 255
        cParaData(8) = 255
        cParaData(9) = 255
        cParaData(10) = 255
        iparalen = 11
        ireturn = hsHTA850SetMifareReader(ghComm, cParaData(0), iparalen, ireturncode, 5000)
        If ireturn = 0 Then
            txtreturn.Text = "hsHTA850SetMifareReader:  OK!"
        Else
            txtreturn.Text = "hsHTA850SetMifareReader : error!(" & ireturn & ") Error Number:" & Str(ireturncode)
        End If

End Sub

Private Sub Command4_Click()
        Dim ireturn As Integer
        Dim iELID, ireturncode, itablelen As Integer
        Dim i As Integer
        Dim j As Integer
        Dim sshow As String
        Dim sstr1 As String
        sshow = ""
        sstr1 = ""

        Dim btabledata() As Byte
        ReDim btabledata(280)
        ireturn = 0
        ireturncode = 0
        itablelen = 4

        txtreturn.Text = ""

        btabledata(0) = 0 ' 讀取位址0x00 ---> 位址0
        btabledata(1) = 0
        btabledata(2) = 19 ' 0x13    19Bytes
        btabledata(3) = 0
        ireturn = hsHTA850ReadParameter(ghComm, btabledata(0), itablelen, ireturncode, 1000)
        If ireturn = 0 Then
            For i = 0 To itablelen - 1
                sstr1 = Hex(btabledata(i))
                If Len(sstr1) < 2 Then
                    sstr1 = "0" & sstr1
                End If

                sshow = sshow & sstr1 & " "
            Next i
            txtreturn.Text = "hsELReadParamenter : ok! " & sshow
        Else
            txtreturn.Text = "hsELReadParamenter : error!(" & ireturn & ")"
        End If
            
End Sub

Private Sub Command5_Click()
 Dim i As Integer
        Dim ireturn As Integer
        Dim ireturncode As Integer
        Dim iCardFormatLen As Integer
        Dim stRecord(255) As Byte
        Dim Card As String
        Dim gStr As String
        Card = Text1.Text
        i = 0
        For i = 0 To 255
            stRecord(i) = 0
        Next
        
        'For i = 0 To 122
        For i = 0 To 385
            sFPtData1(i) = 0
        Next
        
        'For i = 0 To 122
        For i = 0 To 385
            sFPtData2(i) = 0
        Next

        For i = 0 To Len(Card) - 1
            stRecord(i) = Asc(Mid(Card, (i + 1), 1))
        Next

        iCardFormatLen = 0

        ireturn = 0
        ireturncode = 0

        txtreturn.Text = ""
        'ireturn = hsHTA850QueryUserRecord(ghComm, 16, stRecord(0), sCardFormatData(0), iCardFormatLen, ireturncode, 1000)
        'ireturn = hsHTA850QueryUserFingerPrinter(ghComm, 16, stRecord(0), sFPtData1(0), sFPtData2(0), iCardFormatLen, ireturncode, 5000)
        ireturn = hsHTA850QueryUserFingerPrinter2(ghComm, 16, stRecord(0), sFPtData1(0), sFPtData2(0), iCardFormatLen, ireturncode, 5000)
        List2.Clear
        If ireturn = 0 Then
            gStr = ""
            'For i = 0 To 122 - 1
            For i = 0 To 385
                gStr = gStr & Right("0" & Hex(sFPtData1(i)), 2) & " "
            Next
            List2.AddItem ("First FP:")
            List2.AddItem (gStr)
            'txtreturn.Text = "Query user Record: ok! " & gStr
            Text3.Text = gStr
            gStr = ""
            'For i = 0 To 122 - 1
            For i = 0 To 385
                gStr = gStr & Right("0" & Hex(sFPtData2(i)), 2) & " "
            Next
           ' txtreturn.Text = "Query user Record: ok! " & gStr
            List2.AddItem ("Second FP:")
            List2.AddItem (gStr)
            Text4.Text = gStr
            
        Else
            If ireturncode = 6 Then
                txtreturn.Text = "Query Card:Not Exist!"
            Else
                txtreturn.Text = "Query user Record: error!(" & ireturn & ")  Error Number:" & Str(ireturncode)
            End If
        End If
      
End Sub

Private Sub Command6_Click()
        Dim iparalen As Integer
        Dim ireturncode, ireturn As Integer
        Dim cParaData(30) As Byte
        

        ireturn = 0
        ireturncode = 0
        txtreturn.Text = ""

        cParaData(0) = 1 '0x00 Legal Card Max seting
        cParaData(1) = 3 '0x03 History Max Seting
        cParaData(2) = 16 '0x00 Card Length
        cParaData(3) = 0 '0x00 reserve
        cParaData(4) = 16 '0x00 Message Length
        cParaData(5) = 0 '0x00  reserve
        cParaData(6) = 0 '0xE8  reserve
        cParaData(7) = 0 '0x03  reserve
        cParaData(8) = 0 '0x00  reserve
        cParaData(9) = 0 '0x00  reserve
        iparalen = 10
        ireturn = hsHTA850WriteParameter(ghComm, cParaData(0), iparalen, ireturncode, 5000)
        If ireturn = 0 Then
            txtreturn.Text = "hsHTA850WriteParamenter :  OK!"
        Else
            txtreturn.Text = "hsHTA850WriteParamenter : error!(" & ireturn & ") Error Number:" & Str(ireturncode)
        End If

 
End Sub

Private Sub Command7_Click()
        Dim ireturn As Integer
        Dim ireturncode As Integer
        Dim iwritelen As Integer
        Dim i As Integer
        Dim itable As Integer
        Dim btabledata() As Byte
        ReDim btabledata(279)

        For i = 0 To 159
            btabledata(i) = 0
        Next i

        ireturn = 0
        ireturncode = 0
        iwritelen = 160
        txtreturn.Text = ""

        ireturn = hsHTA850WriteTable(ghComm, 1, btabledata(0), iwritelen, ireturncode, 1000)
        If ireturn = 0 Then
            txtreturn.Text = "hsHTA850WriteTable: ok!"
        Else
            txtreturn.Text = "hsHTA850WriteTable: error!(" & ireturn & ") Error Number:" & Str(ireturncode)
        End If
End Sub

Private Sub Command8_Click()
   
        Dim ireturn, itablelen, ireturncode As Integer
        Dim i As Integer
        Dim sshow As String
        Dim sstr1 As String
        sshow = ""
        sstr1 = ""

        Dim btabledata() As Byte
        ReDim btabledata(280)
        ireturn = 0
        ireturncode = 0
        itablelen = 10

        txtreturn.Text = ""

        ireturn = hsHTA850ReadTable(ghComm, 1, btabledata(0), itablelen, ireturncode, 1000)
        If ireturn = 0 Then
            For i = 0 To itablelen - 1
                sstr1 = Hex(btabledata(i))
                If Len(sstr1) < 2 Then
                    sstr1 = "0" & sstr1
                End If

                sshow = sshow & sstr1 & " "
            Next i
            txtreturn.Text = "hsHTA850ReadTable : ok! " & sshow
        Else
            txtreturn.Text = "hsHTA850ReadTable : error!(" & ireturn & ")"
        End If
End Sub

Private Sub Command9_Click()
Dim bstCard() As Byte
        ReDim bstCard(5 * 224)

    Dim ireturn As Integer
    Dim ireturncode As Integer
    Dim iRecord As Integer
    Dim i As Integer, k As Integer
    'Dim stRecord(512) As Byte
    Dim stRecord(806) As Byte
    Dim Card, Msg As String
    
    Dim newFP1, newFP2
    
   For i = 0 To 385
      sFPtData1(i) = 0
   Next
   For i = 0 To 385
       sFPtData2(i) = 0
   Next
      
   If Text3 <> "" Then
      newFP1 = Split(Text3, " ")
      newFP2 = Split(Text4, " ")
      For i = 0 To UBound(newFP1)
          sFPtData1(i) = Val("&H" & newFP1(i))
      Next
      For i = 0 To UBound(newFP2)
          sFPtData2(i) = Val("&H" & newFP2(i))
      Next
   End If
    
    
    If sFPtData1(0) <> "0" Then
        Card = Text1.Text
        'Msg = "VBOK"
        Msg = Text2.Text
        i = 0
        'For i = 0 To 511
        For i = 0 To 805
            stRecord(i) = 0
        Next

        For i = 0 To Len(Card) - 1
            stRecord(i) = Asc(Mid(Card, i + 1, 1))
        Next

        stRecord(16) = 0
        stRecord(17) = 1
        
        'For i = 18 To 18 + Len(Msg) - 1
        '    stRecord(i) = Asc(Mid(Msg, (i - 17), 1))
        'Next
         k = 0
        For i = 18 To 18 + Len(Msg) - 1
            If Asc(Mid(Msg, (i - 17), 1)) > 0 Then
               stRecord(i + k) = Asc(Mid(Msg, (i - 17), 1))
            Else
               stRecord(i + k) = Val("&H" & Left(Hex(Asc(Mid(Msg, (i - 17), 1))), 2))
               k = k + 1
               stRecord(i + k) = Val("&H" & Right(Hex(Asc(Mid(Msg, (i - 17), 1))), 2))
            End If
        Next
        
        
        'For i = 34 To 155
        Msg = ""
        For i = 0 To 385
            stRecord(34 + i) = sFPtData1(i)
            Msg = Msg & Right("0" & Hex(sFPtData1(i)), 2) & " "
        Next
        Debug.Print "Fg1:" & Msg
        
        'For i = 156 To 278
        '    stRecord(i) = sFPtData2(i - 156)
        Msg = ""
        For i = 0 To 385
            stRecord(i + 420) = sFPtData2(i)
            Msg = Msg & Right("0" & Hex(sFPtData2(i)), 2) & " "
        Next
        Debug.Print "Fg2:" & Msg
        
        ireturn = 0
        ireturncode = 0
        iRecord = 1



        'ireturn = hsHTA850InsertMultiUserFingerPrinter(ghComm, 16, 16, iRecord, stRecord(0), ireturncode, 5000)
        ireturn = hsHTA850InsertMultiUserFingerPrinter2(ghComm, 16, 16, 1, stRecord(0), ireturncode, 30000)
        List2.Clear
        If ireturn = 0 Then
             List2.AddItem ("hsHTA850InsertMultiUserFingerPrinter : ok!")
        Else
            'If ireturncode = 2 Then
            '    List2.AddItem ("The length of card is not true!")
           ' Else
               List2.AddItem ("hsHTA850InsertMultiUserFingerPrinter : error!(" & ireturn & ") Error Number:" & Str(ireturncode))
           ' End If
        End If
    Else
        List2.AddItem ("Please click hsHTA850QueryUserFingerPrinter button first!!")
    End If




End Sub

Private Sub Form_Unload(Cancel As Integer)

    If ghComm > 0 Then
       ireturn = hsCloseChannel(ghComm)
    End If
End Sub

