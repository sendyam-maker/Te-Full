VERSION 5.00
Begin VB.Form frm140102 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "北所銷卷案號輸入作業"
   ClientHeight    =   5610
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8740
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8740
   Begin VB.PictureBox pic2 
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   4905
      Left            =   60
      ScaleHeight     =   4910
      ScaleWidth      =   8420
      TabIndex        =   15
      Top             =   660
      Width           =   8415
      Begin VB.PictureBox pic1 
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000008&
         Height          =   4905
         Left            =   0
         ScaleHeight     =   4910
         ScaleWidth      =   8420
         TabIndex        =   16
         Top             =   0
         Width           =   8415
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   249
            Left            =   0
            TabIndex        =   266
            Top             =   10125
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   248
            Left            =   6660
            TabIndex        =   265
            Top             =   9900
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   247
            Left            =   4995
            TabIndex        =   264
            Top             =   9900
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   246
            Left            =   3330
            TabIndex        =   263
            Top             =   9900
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   245
            Left            =   1665
            TabIndex        =   262
            Top             =   9900
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   244
            Left            =   0
            TabIndex        =   261
            Top             =   9900
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   243
            Left            =   6660
            TabIndex        =   260
            Top             =   10125
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   242
            Left            =   4995
            TabIndex        =   259
            Top             =   10125
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   241
            Left            =   3330
            TabIndex        =   258
            Top             =   10125
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   240
            Left            =   1665
            TabIndex        =   257
            Top             =   10125
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   239
            Left            =   0
            TabIndex        =   256
            Top             =   10575
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   238
            Left            =   6660
            TabIndex        =   255
            Top             =   10350
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   237
            Left            =   4995
            TabIndex        =   254
            Top             =   10350
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   236
            Left            =   3330
            TabIndex        =   253
            Top             =   10350
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   235
            Left            =   1665
            TabIndex        =   252
            Top             =   10350
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   234
            Left            =   0
            TabIndex        =   251
            Top             =   10350
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   233
            Left            =   6660
            TabIndex        =   250
            Top             =   10575
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   232
            Left            =   4995
            TabIndex        =   249
            Top             =   10575
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   231
            Left            =   3330
            TabIndex        =   248
            Top             =   10575
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   230
            Left            =   1665
            TabIndex        =   247
            Top             =   10575
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   229
            Left            =   0
            TabIndex        =   246
            Top             =   11025
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   228
            Left            =   6660
            TabIndex        =   245
            Top             =   10800
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   227
            Left            =   4995
            TabIndex        =   244
            Top             =   10800
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   226
            Left            =   3330
            TabIndex        =   243
            Top             =   10800
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   225
            Left            =   1665
            TabIndex        =   242
            Top             =   10800
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   224
            Left            =   0
            TabIndex        =   241
            Top             =   10800
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   223
            Left            =   6660
            TabIndex        =   240
            Top             =   11025
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   222
            Left            =   4995
            TabIndex        =   239
            Top             =   11025
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   221
            Left            =   3330
            TabIndex        =   238
            Top             =   11025
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   220
            Left            =   1665
            TabIndex        =   237
            Top             =   11025
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   219
            Left            =   0
            TabIndex        =   236
            Top             =   4950
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   218
            Left            =   1665
            TabIndex        =   235
            Top             =   4950
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   217
            Left            =   3330
            TabIndex        =   234
            Top             =   4950
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   216
            Left            =   4995
            TabIndex        =   233
            Top             =   4950
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   215
            Left            =   6660
            TabIndex        =   232
            Top             =   4950
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   214
            Left            =   0
            TabIndex        =   231
            Top             =   5175
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   213
            Left            =   6660
            TabIndex        =   230
            Top             =   5175
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   212
            Left            =   4995
            TabIndex        =   229
            Top             =   5175
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   211
            Left            =   3330
            TabIndex        =   228
            Top             =   5175
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   210
            Left            =   1665
            TabIndex        =   227
            Top             =   5175
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   209
            Left            =   0
            TabIndex        =   226
            Top             =   5625
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   208
            Left            =   6660
            TabIndex        =   225
            Top             =   5400
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   207
            Left            =   4995
            TabIndex        =   224
            Top             =   5400
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   206
            Left            =   3330
            TabIndex        =   223
            Top             =   5400
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   205
            Left            =   1665
            TabIndex        =   222
            Top             =   5400
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   204
            Left            =   0
            TabIndex        =   221
            Top             =   5400
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   203
            Left            =   6660
            TabIndex        =   220
            Top             =   5625
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   202
            Left            =   4995
            TabIndex        =   219
            Top             =   5625
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   201
            Left            =   3330
            TabIndex        =   218
            Top             =   5625
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   200
            Left            =   1665
            TabIndex        =   217
            Top             =   5625
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   199
            Left            =   0
            TabIndex        =   216
            Top             =   6075
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   198
            Left            =   6660
            TabIndex        =   215
            Top             =   5850
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   197
            Left            =   4995
            TabIndex        =   214
            Top             =   5850
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   196
            Left            =   3330
            TabIndex        =   213
            Top             =   5850
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   195
            Left            =   1665
            TabIndex        =   212
            Top             =   5850
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   194
            Left            =   0
            TabIndex        =   211
            Top             =   5850
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   193
            Left            =   6660
            TabIndex        =   210
            Top             =   6075
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   192
            Left            =   4995
            TabIndex        =   209
            Top             =   6075
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   191
            Left            =   3330
            TabIndex        =   208
            Top             =   6075
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   190
            Left            =   1665
            TabIndex        =   207
            Top             =   6075
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   189
            Left            =   0
            TabIndex        =   206
            Top             =   6525
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   188
            Left            =   6660
            TabIndex        =   205
            Top             =   6300
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   187
            Left            =   4995
            TabIndex        =   204
            Top             =   6300
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   186
            Left            =   3330
            TabIndex        =   203
            Top             =   6300
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   185
            Left            =   1665
            TabIndex        =   202
            Top             =   6300
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   184
            Left            =   0
            TabIndex        =   201
            Top             =   6300
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   183
            Left            =   6660
            TabIndex        =   200
            Top             =   6525
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   182
            Left            =   4995
            TabIndex        =   199
            Top             =   6525
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   181
            Left            =   3330
            TabIndex        =   198
            Top             =   6525
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   180
            Left            =   1665
            TabIndex        =   197
            Top             =   6525
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   179
            Left            =   0
            TabIndex        =   196
            Top             =   6975
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   178
            Left            =   6660
            TabIndex        =   195
            Top             =   6750
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   177
            Left            =   4995
            TabIndex        =   194
            Top             =   6750
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   176
            Left            =   3330
            TabIndex        =   193
            Top             =   6750
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   175
            Left            =   1665
            TabIndex        =   192
            Top             =   6750
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   174
            Left            =   0
            TabIndex        =   191
            Top             =   6750
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   173
            Left            =   6660
            TabIndex        =   190
            Top             =   6975
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   172
            Left            =   4995
            TabIndex        =   189
            Top             =   6975
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   171
            Left            =   3330
            TabIndex        =   188
            Top             =   6975
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   170
            Left            =   1665
            TabIndex        =   187
            Top             =   6975
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   169
            Left            =   0
            TabIndex        =   186
            Top             =   7425
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   168
            Left            =   6660
            TabIndex        =   185
            Top             =   7200
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   167
            Left            =   4995
            TabIndex        =   184
            Top             =   7200
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   166
            Left            =   3330
            TabIndex        =   183
            Top             =   7200
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   165
            Left            =   1665
            TabIndex        =   182
            Top             =   7200
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   164
            Left            =   0
            TabIndex        =   181
            Top             =   7200
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   163
            Left            =   6660
            TabIndex        =   180
            Top             =   7425
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   162
            Left            =   4995
            TabIndex        =   179
            Top             =   7425
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   161
            Left            =   3330
            TabIndex        =   178
            Top             =   7425
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   160
            Left            =   1665
            TabIndex        =   177
            Top             =   7425
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   159
            Left            =   0
            TabIndex        =   176
            Top             =   7875
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   158
            Left            =   6660
            TabIndex        =   175
            Top             =   7650
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   157
            Left            =   4995
            TabIndex        =   174
            Top             =   7650
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   156
            Left            =   3330
            TabIndex        =   173
            Top             =   7650
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   155
            Left            =   1665
            TabIndex        =   172
            Top             =   7650
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   154
            Left            =   0
            TabIndex        =   171
            Top             =   7650
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   153
            Left            =   6660
            TabIndex        =   170
            Top             =   7875
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   152
            Left            =   4995
            TabIndex        =   169
            Top             =   7875
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   151
            Left            =   3330
            TabIndex        =   168
            Top             =   7875
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   150
            Left            =   1665
            TabIndex        =   167
            Top             =   7875
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   149
            Left            =   0
            TabIndex        =   166
            Top             =   8325
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   148
            Left            =   6660
            TabIndex        =   165
            Top             =   8100
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   147
            Left            =   4995
            TabIndex        =   164
            Top             =   8100
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   146
            Left            =   3330
            TabIndex        =   163
            Top             =   8100
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   145
            Left            =   1665
            TabIndex        =   162
            Top             =   8100
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   144
            Left            =   0
            TabIndex        =   161
            Top             =   8100
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   143
            Left            =   6660
            TabIndex        =   160
            Top             =   8325
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   142
            Left            =   4995
            TabIndex        =   159
            Top             =   8325
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   141
            Left            =   3330
            TabIndex        =   158
            Top             =   8325
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   140
            Left            =   1665
            TabIndex        =   157
            Top             =   8325
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   139
            Left            =   0
            TabIndex        =   156
            Top             =   8775
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   138
            Left            =   6660
            TabIndex        =   155
            Top             =   8550
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   137
            Left            =   4995
            TabIndex        =   154
            Top             =   8550
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   136
            Left            =   3330
            TabIndex        =   153
            Top             =   8550
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   135
            Left            =   1665
            TabIndex        =   152
            Top             =   8550
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   134
            Left            =   0
            TabIndex        =   151
            Top             =   8550
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   133
            Left            =   6660
            TabIndex        =   150
            Top             =   8775
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   132
            Left            =   4995
            TabIndex        =   149
            Top             =   8775
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   131
            Left            =   3330
            TabIndex        =   148
            Top             =   8775
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   130
            Left            =   1665
            TabIndex        =   147
            Top             =   8775
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   129
            Left            =   0
            TabIndex        =   146
            Top             =   9225
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   128
            Left            =   6660
            TabIndex        =   145
            Top             =   9000
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   127
            Left            =   4995
            TabIndex        =   144
            Top             =   9000
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   126
            Left            =   3330
            TabIndex        =   143
            Top             =   9000
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   125
            Left            =   1665
            TabIndex        =   142
            Top             =   9000
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   124
            Left            =   0
            TabIndex        =   141
            Top             =   9000
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   123
            Left            =   6660
            TabIndex        =   140
            Top             =   9225
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   122
            Left            =   4995
            TabIndex        =   139
            Top             =   9225
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   121
            Left            =   3330
            TabIndex        =   138
            Top             =   9225
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   120
            Left            =   1665
            TabIndex        =   137
            Top             =   9225
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   119
            Left            =   0
            TabIndex        =   136
            Top             =   9675
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   118
            Left            =   6660
            TabIndex        =   135
            Top             =   9450
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   117
            Left            =   4995
            TabIndex        =   134
            Top             =   9450
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   116
            Left            =   3330
            TabIndex        =   133
            Top             =   9450
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   115
            Left            =   1665
            TabIndex        =   132
            Top             =   9450
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   114
            Left            =   0
            TabIndex        =   131
            Top             =   9450
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   113
            Left            =   6660
            TabIndex        =   130
            Top             =   9675
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   112
            Left            =   4995
            TabIndex        =   129
            Top             =   9675
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   111
            Left            =   3330
            TabIndex        =   128
            Top             =   9675
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   110
            Left            =   1665
            TabIndex        =   127
            Top             =   9675
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   126
            Top             =   0
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   1
            Left            =   1665
            TabIndex        =   125
            Top             =   0
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   2
            Left            =   3330
            TabIndex        =   124
            Top             =   0
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   3
            Left            =   4995
            TabIndex        =   123
            Top             =   0
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   4
            Left            =   6660
            TabIndex        =   122
            Top             =   0
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   5
            Left            =   0
            TabIndex        =   121
            Top             =   225
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   9
            Left            =   6660
            TabIndex        =   120
            Top             =   225
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   8
            Left            =   4995
            TabIndex        =   119
            Top             =   225
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   7
            Left            =   3330
            TabIndex        =   118
            Top             =   225
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   6
            Left            =   1665
            TabIndex        =   117
            Top             =   225
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   15
            Left            =   0
            TabIndex        =   116
            Top             =   675
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   14
            Left            =   6660
            TabIndex        =   115
            Top             =   450
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   13
            Left            =   4995
            TabIndex        =   114
            Top             =   450
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   12
            Left            =   3330
            TabIndex        =   113
            Top             =   450
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   11
            Left            =   1665
            TabIndex        =   112
            Top             =   450
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   10
            Left            =   0
            TabIndex        =   111
            Top             =   450
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   19
            Left            =   6660
            TabIndex        =   110
            Top             =   675
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   18
            Left            =   4995
            TabIndex        =   109
            Top             =   675
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   17
            Left            =   3330
            TabIndex        =   108
            Top             =   675
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   16
            Left            =   1665
            TabIndex        =   107
            Top             =   675
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   25
            Left            =   0
            TabIndex        =   106
            Top             =   1125
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   24
            Left            =   6660
            TabIndex        =   105
            Top             =   900
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   23
            Left            =   4995
            TabIndex        =   104
            Top             =   900
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   22
            Left            =   3330
            TabIndex        =   103
            Top             =   900
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   21
            Left            =   1665
            TabIndex        =   102
            Top             =   900
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   20
            Left            =   0
            TabIndex        =   101
            Top             =   900
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   29
            Left            =   6660
            TabIndex        =   100
            Top             =   1125
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   28
            Left            =   4995
            TabIndex        =   99
            Top             =   1125
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   27
            Left            =   3330
            TabIndex        =   98
            Top             =   1125
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   26
            Left            =   1665
            TabIndex        =   97
            Top             =   1125
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   35
            Left            =   0
            TabIndex        =   96
            Top             =   1575
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   34
            Left            =   6660
            TabIndex        =   95
            Top             =   1350
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   33
            Left            =   4995
            TabIndex        =   94
            Top             =   1350
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   32
            Left            =   3330
            TabIndex        =   93
            Top             =   1350
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   31
            Left            =   1665
            TabIndex        =   92
            Top             =   1350
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   30
            Left            =   0
            TabIndex        =   91
            Top             =   1350
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   39
            Left            =   6660
            TabIndex        =   90
            Top             =   1575
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   38
            Left            =   4995
            TabIndex        =   89
            Top             =   1575
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   37
            Left            =   3330
            TabIndex        =   88
            Top             =   1575
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   36
            Left            =   1665
            TabIndex        =   87
            Top             =   1575
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   45
            Left            =   0
            TabIndex        =   86
            Top             =   2025
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   44
            Left            =   6660
            TabIndex        =   85
            Top             =   1800
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   43
            Left            =   4995
            TabIndex        =   84
            Top             =   1800
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   42
            Left            =   3330
            TabIndex        =   83
            Top             =   1800
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   41
            Left            =   1665
            TabIndex        =   82
            Top             =   1800
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   40
            Left            =   0
            TabIndex        =   81
            Top             =   1800
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   49
            Left            =   6660
            TabIndex        =   80
            Top             =   2025
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   48
            Left            =   4995
            TabIndex        =   79
            Top             =   2025
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   47
            Left            =   3330
            TabIndex        =   78
            Top             =   2025
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   46
            Left            =   1665
            TabIndex        =   77
            Top             =   2025
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   55
            Left            =   0
            TabIndex        =   76
            Top             =   2475
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   54
            Left            =   6660
            TabIndex        =   75
            Top             =   2250
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   53
            Left            =   4995
            TabIndex        =   74
            Top             =   2250
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   52
            Left            =   3330
            TabIndex        =   73
            Top             =   2250
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   51
            Left            =   1665
            TabIndex        =   72
            Top             =   2250
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   50
            Left            =   0
            TabIndex        =   71
            Top             =   2250
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   59
            Left            =   6660
            TabIndex        =   70
            Top             =   2475
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   58
            Left            =   4995
            TabIndex        =   69
            Top             =   2475
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   57
            Left            =   3330
            TabIndex        =   68
            Top             =   2475
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   56
            Left            =   1665
            TabIndex        =   67
            Top             =   2475
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   65
            Left            =   0
            TabIndex        =   66
            Top             =   2925
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   64
            Left            =   6660
            TabIndex        =   65
            Top             =   2700
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   63
            Left            =   4995
            TabIndex        =   64
            Top             =   2700
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   62
            Left            =   3330
            TabIndex        =   63
            Top             =   2700
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   61
            Left            =   1665
            TabIndex        =   62
            Top             =   2700
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   60
            Left            =   0
            TabIndex        =   61
            Top             =   2700
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   69
            Left            =   6660
            TabIndex        =   60
            Top             =   2925
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   68
            Left            =   4995
            TabIndex        =   59
            Top             =   2925
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   67
            Left            =   3330
            TabIndex        =   58
            Top             =   2925
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   66
            Left            =   1665
            TabIndex        =   57
            Top             =   2925
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   75
            Left            =   0
            TabIndex        =   56
            Top             =   3375
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   74
            Left            =   6660
            TabIndex        =   55
            Top             =   3150
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   73
            Left            =   4995
            TabIndex        =   54
            Top             =   3150
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   72
            Left            =   3330
            TabIndex        =   53
            Top             =   3150
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   71
            Left            =   1665
            TabIndex        =   52
            Top             =   3150
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   70
            Left            =   0
            TabIndex        =   51
            Top             =   3150
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   79
            Left            =   6660
            TabIndex        =   50
            Top             =   3375
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   78
            Left            =   4995
            TabIndex        =   49
            Top             =   3375
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   77
            Left            =   3330
            TabIndex        =   48
            Top             =   3375
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   76
            Left            =   1665
            TabIndex        =   47
            Top             =   3375
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   85
            Left            =   0
            TabIndex        =   46
            Top             =   3825
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   84
            Left            =   6660
            TabIndex        =   45
            Top             =   3600
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   83
            Left            =   4995
            TabIndex        =   44
            Top             =   3600
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   82
            Left            =   3330
            TabIndex        =   43
            Top             =   3600
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   81
            Left            =   1665
            TabIndex        =   42
            Top             =   3600
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   80
            Left            =   0
            TabIndex        =   41
            Top             =   3600
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   89
            Left            =   6660
            TabIndex        =   40
            Top             =   3825
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   88
            Left            =   4995
            TabIndex        =   39
            Top             =   3825
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   87
            Left            =   3330
            TabIndex        =   38
            Top             =   3825
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   86
            Left            =   1665
            TabIndex        =   37
            Top             =   3825
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   95
            Left            =   0
            TabIndex        =   36
            Top             =   4275
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   94
            Left            =   6660
            TabIndex        =   35
            Top             =   4050
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   93
            Left            =   4995
            TabIndex        =   34
            Top             =   4050
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   92
            Left            =   3330
            TabIndex        =   33
            Top             =   4050
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   91
            Left            =   1665
            TabIndex        =   32
            Top             =   4050
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   90
            Left            =   0
            TabIndex        =   31
            Top             =   4050
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   99
            Left            =   6660
            TabIndex        =   30
            Top             =   4275
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   98
            Left            =   4995
            TabIndex        =   29
            Top             =   4275
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   97
            Left            =   3330
            TabIndex        =   28
            Top             =   4275
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   96
            Left            =   1665
            TabIndex        =   27
            Top             =   4275
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   105
            Left            =   0
            TabIndex        =   26
            Top             =   4725
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   104
            Left            =   6660
            TabIndex        =   25
            Top             =   4500
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   103
            Left            =   4995
            TabIndex        =   24
            Top             =   4500
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   102
            Left            =   3330
            TabIndex        =   23
            Top             =   4500
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   101
            Left            =   1665
            TabIndex        =   22
            Top             =   4500
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   100
            Left            =   0
            TabIndex        =   21
            Top             =   4500
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   109
            Left            =   6660
            TabIndex        =   20
            Top             =   4725
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   108
            Left            =   4995
            TabIndex        =   19
            Top             =   4725
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   107
            Left            =   3330
            TabIndex        =   18
            Top             =   4725
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Height          =   225
            Index           =   106
            Left            =   1665
            TabIndex        =   17
            Top             =   4725
            Visible         =   0   'False
            Width           =   1665
         End
      End
   End
   Begin VB.VScrollBar V1 
      Height          =   4965
      Left            =   8490
      TabIndex        =   14
      Top             =   630
      Width           =   255
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "全部取消(&U)"
      Height          =   345
      Index           =   2
      Left            =   5370
      TabIndex        =   10
      Top             =   0
      Width           =   1245
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "全部選取(&C)"
      Height          =   345
      Index           =   1
      Left            =   4065
      TabIndex        =   9
      Top             =   0
      Width           =   1275
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   3180
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   7
      Left            =   4590
      MaxLength       =   2
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   6
      Left            =   4350
      MaxLength       =   1
      TabIndex        =   6
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   3510
      MaxLength       =   6
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   3030
      MaxLength       =   3
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   2670
      MaxLength       =   2
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   2430
      MaxLength       =   1
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1590
      MaxLength       =   6
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1110
      MaxLength       =   3
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   4
      Left            =   7530
      TabIndex        =   12
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "存檔(&O)"
      Height          =   345
      Index           =   3
      Left            =   6645
      TabIndex        =   11
      Top             =   0
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   2010
      X2              =   3300
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   420
      Width           =   900
   End
End
Attribute VB_Name = "frm140102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2022/2/9 Form2.0不用改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit

Const MaxChk As Integer = 250
Dim i As Integer
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim IsSave As Boolean


Private Sub cmdOK_Click(Index As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse

Select Case Index
Case 0
     StrMenu
Case 1
     IsSave = False
     For i = 0 To MaxChk - 1
        If Check1(i).Visible = True And Check1(i).Enabled = True Then IsSave = True: Check1(i).Value = vbChecked
     Next i
     If IsSave = False Then
        MsgBox "沒有可以選取的！", vbExclamation
     End If
Case 2
     IsSave = False
     For i = 0 To MaxChk - 1
        If Check1(i).Visible = True And Check1(i).Enabled = True Then IsSave = True: Check1(i).Value = vbUnchecked
     Next i
     If IsSave = False Then
        MsgBox "沒有可以取消的！", vbExclamation
     End If
Case 3
     Screen.MousePointer = vbHourglass
     'Me.Enabled = False
     IsSave = False
     DoEvents
     On Error GoTo oErr
     strDate = DBDATE(DateAdd("yyyy", -1, ChangeWStringToWDateString(strSrvDate(1)))) 'Add By Sindy 2012/1/4
     'cnnConnection.BeginTrans
     For i = 0 To MaxChk - 1
        If Check1(i).Visible = True And Check1(i).Enabled = True And Check1(i).Value = vbChecked Then
            cnnConnection.BeginTrans 'Modify By Sindy 2012/1/4 因要逐筆Commit
            'Add By Sindy 2012/1/4
            strSql = "SELECT count(*) FROM CaseProgress WHERE cp01='" & SystemNumber(Check1(i).Caption, 1) & "' and cp02='" & SystemNumber(Check1(i).Caption, 2) & "' and cp03='" & SystemNumber(Check1(i).Caption, 3) & "' and cp04='" & SystemNumber(Check1(i).Caption, 4) & "' and cp05>" & strDate & " and substr(cp09,1,1)='A' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
                If RsTemp.Fields(0) > 0 Then
                   strTit = "詢問"
                   strMsg = "此案(" & Check1(i).Caption & ")一年內仍有收文，是否確定銷卷？"
                   nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
                   If nResponse = vbNo Then
                     cnnConnection.CommitTrans
                     Screen.MousePointer = vbDefault
                     Exit Sub
                   End If
                End If
            End If
            '2012/1/4 End
            cnnConnection.Execute "update patent set pa108=to_number(to_char(sysdate,'YYYYMMDD')) where pa01='" & SystemNumber(Check1(i).Caption, 1) & "' and pa02='" & SystemNumber(Check1(i).Caption, 2) & "' and pa03='" & SystemNumber(Check1(i).Caption, 3) & "' and pa04='" & SystemNumber(Check1(i).Caption, 4) & "' and pa108 is null  "
            cnnConnection.Execute "update trademark set tm57=to_number(to_char(sysdate,'YYYYMMDD')) where tm01='" & SystemNumber(Check1(i).Caption, 1) & "' and tm02='" & SystemNumber(Check1(i).Caption, 2) & "' and tm03='" & SystemNumber(Check1(i).Caption, 3) & "' and tm04='" & SystemNumber(Check1(i).Caption, 4) & "' and tm57 is null  "
            cnnConnection.Execute "update servicepractice set sp61=to_number(to_char(sysdate,'YYYYMMDD')) where sp01='" & SystemNumber(Check1(i).Caption, 1) & "' and sp02='" & SystemNumber(Check1(i).Caption, 2) & "' and sp03='" & SystemNumber(Check1(i).Caption, 3) & "' and sp04='" & SystemNumber(Check1(i).Caption, 4) & "' and sp61 is null  "
            cnnConnection.Execute "update lawcase set lc34=to_number(to_char(sysdate,'YYYYMMDD')) where lc01='" & SystemNumber(Check1(i).Caption, 1) & "' and lc02='" & SystemNumber(Check1(i).Caption, 2) & "' and lc03='" & SystemNumber(Check1(i).Caption, 3) & "' and lc04='" & SystemNumber(Check1(i).Caption, 4) & "' and lc34 is null  "
            cnnConnection.Execute "update hirecase set hc19=to_number(to_char(sysdate,'YYYYMMDD')) where hc01='" & SystemNumber(Check1(i).Caption, 1) & "' and hc02='" & SystemNumber(Check1(i).Caption, 2) & "' and hc03='" & SystemNumber(Check1(i).Caption, 3) & "' and hc04='" & SystemNumber(Check1(i).Caption, 4) & "' and hc19 is null  "
            IsSave = True
            cnnConnection.CommitTrans 'Modify By Sindy 2012/1/4 逐筆Commit
            Check1(i).Enabled = False
        End If
     Next i
     'cnnConnection.CommitTrans
     StrMenu
     If IsSave = True Then
        MsgBox "存檔成功！", vbInformation
     Else
        MsgBox "沒有可以存檔的！", vbExclamation
     End If
     'Me.Enabled = True
     Screen.MousePointer = vbDefault
Case 4
     Unload Me
Case Else
End Select
Exit Sub
oErr:
    cnnConnection.RollbackTrans
    MsgBox "存檔失敗！", vbExclamation
    Me.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Sub StrMenu()
Dim strSQL1 As String
Dim strSQL2 As String
Dim StrSQL3 As String
Dim StrSQL4 As String
Dim strSQL5 As String
'Add By Sindy 2014/9/4
Dim bolAmt As Boolean
Dim lngAmt As Long
Dim lngNote As Long
Dim lngFee As Long, bolFee As Boolean 'Add By Sindy 2015/3/9
'2014/9/4 END
   
   If Text1(0) <> Text1(4) And Text1(4) <> "" Then
      MsgBox "不可以針對不同系統類別案件作業！", vbExclamation
      Text1(4).SetFocus
      Exit Sub
   End If
   If Trim(Text1(0)) = "" Then
      MsgBox "系統別不可空白！", vbExclamation
      Text1(0).SetFocus
      Exit Sub
   End If
   If Trim(Text1(1)) = "" Then
      MsgBox "流水號起不可空白！", vbExclamation
      Text1(1).SetFocus
      Exit Sub
   End If
   If Trim(Text1(2)) = "" Then Text1(2) = "0"
   If Text1(4) <> "" Then
      If Trim(Text1(5)) = "" Then
         MsgBox "流水號起不可空白！", vbExclamation
         Text1(5).SetFocus
         Exit Sub
      End If
      If Trim(Text1(6)) = "" Then Text1(6) = "9"
   End If
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   DoEvents
   If Text1(0) <> "" And Text1(4) <> "" Then
'edit by nickc 2007/10/04 修正 BUG
'        strSQL1 = "pa01='" & Text1(0) & "' and pa02>='" & Text1(1) & "' and pa04='00' and pa03>='" & Text1(2) & "' and pa02<='" & Text1(5) & "' and pa03<='" & Text1(6) & "' "
'        strSQL2 = "tm01='" & Text1(0) & "' and tm02>='" & Text1(1) & "' and tm02<='" & Text1(5) & "'  and tm04='00' and tm03>='" & Text1(2) & "' and tm03<='" & Text1(6) & "' "
'        StrSQL3 = "sp01='" & Text1(0) & "' and sp02>='" & Text1(1) & "' and sp02<='" & Text1(5) & "'  and sp04='00' and sp03>='" & Text1(2) & "' and sp03<='" & Text1(6) & "' "
'        StrSQL4 = "lc01='" & Text1(0) & "' and lc02>='" & Text1(1) & "' and lc02<='" & Text1(5) & "'  and lc04='00' and lc03>='" & Text1(2) & "' and lc03<='" & Text1(6) & "' "
'        strSQL5 = "hc01='" & Text1(0) & "' and hc02>='" & Text1(1) & "' and hc02<='" & Text1(5) & "'  and hc04='00' and hc03>='" & Text1(2) & "' and hc03<='" & Text1(6) & "' "
      strSQL1 = "pa01='" & Text1(0) & "' and pa02||pa03>='" & Text1(1) & Text1(2) & "' and pa04='00'  and pa02||pa03<='" & Text1(5) & Text1(6) & "' "
      strSQL2 = "tm01='" & Text1(0) & "' and tm02||tm03>='" & Text1(1) & Text1(2) & "'  and tm04='00' and tm02||tm03<='" & Text1(5) & Text1(6) & "' "
      StrSQL3 = "sp01='" & Text1(0) & "' and sp02||sp03>='" & Text1(1) & Text1(2) & "'  and sp04='00' and sp02||sp03<='" & Text1(5) & Text1(6) & "' "
      StrSQL4 = "lc01='" & Text1(0) & "' and lc02||lc03>='" & Text1(1) & Text1(2) & "'  and lc04='00' and lc02||lc03<='" & Text1(5) & Text1(6) & "' "
      strSQL5 = "hc01='" & Text1(0) & "' and hc02||hc03>='" & Text1(1) & Text1(2) & "'  and hc04='00' and hc02||hc03<='" & Text1(5) & Text1(6) & "' "
   ElseIf Text1(0) <> "" And Text1(4) = "" Then
      strSQL1 = "pa01='" & Text1(0) & "' and pa02='" & Text1(1) & "' and pa04='00' and pa03='" & Text1(2) & "' "
      strSQL2 = "tm01='" & Text1(0) & "' and tm02='" & Text1(1) & "' and tm04='00' and tm03='" & Text1(2) & "' "
      StrSQL3 = "sp01='" & Text1(0) & "' and sp02='" & Text1(1) & "' and sp04='00' and sp03='" & Text1(2) & "' "
      StrSQL4 = "lc01='" & Text1(0) & "' and lc02='" & Text1(1) & "' and lc04='00' and lc03='" & Text1(2) & "' "
      strSQL5 = "hc01='" & Text1(0) & "' and hc02='" & Text1(1) & "' and hc04='00' and hc03='" & Text1(2) & "' "
   End If
'Modify By Sindy 2014/9/4 +案件備註是否有不銷卷字樣者
'+,instr(pa91,'不銷卷')
'+,instr(tm58,'不銷卷')
'+,instr(sp18,'不銷卷')
'+,instr(lc27,'不銷卷')
'+,instr(hc12,'不銷卷')
strSql = "select * from (     "
strSql = strSql & "select pa01||'-'||pa02||'-'||pa03||'-'||pa04,pa108,instr(pa91,'不銷卷') from patent where " & strSQL1
strSql = strSql & " union select tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm57,instr(tm58,'不銷卷') from trademark where " & strSQL2
strSql = strSql & " union select sp01||'-'||sp02||'-'||sp03||'-'||sp04,sp61,instr(sp18,'不銷卷') from servicepractice where " & StrSQL3
strSql = strSql & " union select lc01||'-'||lc02||'-'||lc03||'-'||lc04,lc34,instr(lc27,'不銷卷') from lawcase where " & StrSQL4
strSql = strSql & " union select hc01||'-'||hc02||'-'||hc03||'-'||hc04,hc19,instr(hc12,'不銷卷') from hirecase where " & strSQL5
strSql = strSql & ") where rownum <" & MaxChk + 1 & " order by 1 "
Set rsTmp = New ADODB.Recordset
If rsTmp.State = 1 Then rsTmp.Close
rsTmp.CursorLocation = adUseClient
rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If rsTmp.RecordCount <> 0 Then
   ReSizePic rsTmp.RecordCount + 1
   rsTmp.MoveFirst
   lngAmt = 0: lngNote = 0 'Add By Sindy 2014/9/4
   lngFee = 0: 'Add By Sindy 2015/3/9
   Do While Not rsTmp.EOF
      Check1(rsTmp.AbsolutePosition - 1).Visible = True
      Check1(rsTmp.AbsolutePosition - 1).Caption = CheckStr(rsTmp.Fields(0))
      'Check1(rsTmp.AbsolutePosition - 1).Enabled = (CheckStr(rsTmp.Fields(1)) = "")
      'Add By Sindy 2014/9/4 案件備註有'不銷卷'字樣者,或是該案號有應收帳款(ACC1K0及ACC0K0)未收者,在畫面上的案號改為不可選取
      bolAmt = PUB_ChkReceivables(rsTmp.Fields(0))
      If bolAmt = True Then lngAmt = lngAmt + 1
      If Val(CheckStr(rsTmp.Fields(2))) > 0 Then lngNote = lngNote + 1
      'Add By Sindy 2015/3/9 檢查財務的規費資料是否不平
      bolFee = False
      '2015/7/24 MODIFY BY SONIA (CFP-013815)TF母案領土延伸分開算,CFP僅EPC子案與母案一起算,接續案或集體設計個別算
      'strSql = "select sum(ax206),sum(ax207) from acc021 where ax214='" & SystemNumber(rsTmp.Fields(0), 1) & SystemNumber(rsTmp.Fields(0), 2) & SystemNumber(rsTmp.Fields(0), 3) & SystemNumber(rsTmp.Fields(0), 4) & "' and substr(ax205,1,4)='2201'"
      If SystemNumber(rsTmp.Fields(0), 1) = "TF" Then
         strSql = "select sum(ax206),sum(ax207) from acc021 where ax214>='" & SystemNumber(rsTmp.Fields(0), 1) & SystemNumber(rsTmp.Fields(0), 2) & "000' and ax214<='" & SystemNumber(rsTmp.Fields(0), 1) & SystemNumber(rsTmp.Fields(0), 2) & "999' and substr(ax205,1,4)='2201'"
      Else
         strSql = "select sum(ax206),sum(ax207) from acc021 where ax214>='" & SystemNumber(rsTmp.Fields(0), 1) & SystemNumber(rsTmp.Fields(0), 2) & SystemNumber(rsTmp.Fields(0), 3) & "00' and ax214<='" & SystemNumber(rsTmp.Fields(0), 1) & SystemNumber(rsTmp.Fields(0), 2) & SystemNumber(rsTmp.Fields(0), 3) & "99' and substr(ax205,1,4)='2201'"
      End If
      '2015/7/24 end
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Val("" & RsTemp.Fields(0)) <> Val("" & RsTemp.Fields(1)) Then
            bolFee = True
            lngFee = lngFee + 1
         End If
      End If
      '2015/3/9 END
      '無銷卷日、無備註不銷卷並且無應收帳款者
      'Modify By Sindy 2015/3/9 +And bolFee = False
      If CheckStr(rsTmp.Fields(1)) = "" And Val(CheckStr(rsTmp.Fields(2))) = 0 And bolAmt = False And bolFee = False Then
         Check1(rsTmp.AbsolutePosition - 1).Enabled = True
      Else
         Check1(rsTmp.AbsolutePosition - 1).Enabled = False
      End If
      '2014/9/4 End
      Check1(rsTmp.AbsolutePosition - 1).Value = IIf(CheckStr(rsTmp.Fields(1)) = "", vbUnchecked, vbChecked)
      rsTmp.MoveNext
   Loop
   'Add By Sindy 2014/9/4
   'Modify By Sindy 2015/3/9 +Or lngFee > 0
   If lngAmt > 0 Or lngNote > 0 Or lngFee > 0 Then
      MsgBox "該區間有 " & lngNote & " 筆不可銷卷案件, " & lngAmt & " 筆有未收款案件, " & lngFee & " 筆財務的規費資料不平", vbInformation
   End If
   '2015/3/9 END
   '2014/9/4 End
Else
   MsgBox "查無任何資料！", vbExclamation
   '將 pic 物件定義高度
   ReSizePic MaxChk
   '將捲軸定義
   V1.Enabled = False
End If
Me.Enabled = True
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   '將checkbox 物件定位
   For i = 0 To MaxChk - 1
      Check1(i).Top = 0 + ((i \ 5) * 225)
      Check1(i).Left = 0 + ((i Mod 5) * 1665)
      Check1(i).TabIndex = i + 8
      Check1(i).Caption = ""
      'Check1(i).Caption = i + 1
      'Check1(i).Visible = True
   Next i
   '將 pic 物件定義高度
   ReSizePic MaxChk
   '將捲軸定義
   V1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140102 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 4
         If Text1(Index) = "" Then Exit Sub
         strExc(0) = "SELECT SK02 FROM SYSTEMKIND WHERE SK01='" & Text1(Index) & "'"
         intI = 1
         'edit by nickc 2007/02/08 不用 dll 了
         'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI <> 1 Then
            MsgBox "無此系統類別，請重新輸入 !", vbCritical
            Cancel = True
         End If
         If Index = 4 Then
            If Text1(0).Text <> "" And Text1(0).Text <> Text1(4).Text And Text1(4).Text <> "" Then
               MsgBox "不可跨不同系統類別做修改 !", vbCritical
               Cancel = True
            End If
         End If
      Case 2
         If Text1(Index) = "" Then Text1(Index) = "0"
      Case 3
         If Text1(Index) = "" Then Text1(Index) = "00"
      Case 6
         If Text1(Index) = "" Then Text1(Index) = "9"
      Case 7
         If Text1(Index) = "" Then Text1(Index) = "99"
   End Select
   If Cancel = True Then TextInverse Text1(Index)
End Sub

Private Sub V1_Change()
pic1.Move 0, V1.Value * 200, pic1.Width, pic1.Height
End Sub

Private Sub V1_Scroll()
pic1.Move 0, V1.Value * 200, pic1.Width, pic1.Height
End Sub

Sub ReSizePic(oCount As Integer)
     For i = 0 To MaxChk - 1
        Check1(i).Visible = False
     Next i
   '將 pic 物件定義高度
   pic1.Height = ((oCount / 5) * 225) + 100
   '將捲軸定義
   If oCount > 110 Then
      V1.max = (pic2.Height - pic1.Height) / 200
      V1.Min = 0
      V1.Value = 0
      V1.Enabled = True
   Else
      V1.Enabled = False
   End If
End Sub
