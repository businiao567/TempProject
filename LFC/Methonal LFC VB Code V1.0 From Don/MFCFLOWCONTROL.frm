VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form MainFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MTL Liqyuid Mass Flow Control, - HORIBA  LV-F30P 1.100"
   ClientHeight    =   9840
   ClientLeft      =   3465
   ClientTop       =   2595
   ClientWidth     =   5880
   FillColor       =   &H000000FF&
   Icon            =   "MFCFLOWCONTROL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   5880
   Begin VB.Frame Frame1 
      Caption         =   "LFC 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   75
      TabIndex        =   127
      Top             =   45
      Width           =   1845
      Begin VB.TextBox FlSet_1OverrideTxt 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   193
         Text            =   "CLOSED"
         Top             =   480
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton LFC1VlvOverRideCmd 
         Height          =   195
         Left            =   1590
         TabIndex        =   192
         ToolTipText     =   "Change LFC valve status."
         Top             =   1185
         Width           =   195
      End
      Begin VB.TextBox FlCV_1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         TabIndex        =   132
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.OptionButton OptMFC1 
         Caption         =   "Normal"
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   131
         Top             =   1470
         Width           =   1725
      End
      Begin VB.OptionButton OptMFC1 
         Caption         =   "Open"
         Height          =   255
         Index           =   1
         Left            =   75
         TabIndex        =   130
         Top             =   1740
         Width           =   1335
      End
      Begin VB.OptionButton OptMFC1 
         Caption         =   "Close"
         Height          =   255
         Index           =   2
         Left            =   75
         TabIndex        =   129
         Top             =   2025
         Width           =   1335
      End
      Begin VB.CommandButton FlowSet1 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   128
         Top             =   480
         Width           =   660
      End
      Begin VB.TextBox FlSet_1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   133
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   135
         Top             =   870
         Width           =   1845
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   134
         Top             =   210
         Width           =   1845
      End
   End
   Begin VB.CommandButton EmoCmd 
      BackColor       =   &H000000FF&
      Caption         =   "EMERGENCY OFF"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   191
      Top             =   8940
      Width           =   5745
   End
   Begin VB.TextBox LogRateTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   8295
      TabIndex        =   189
      Text            =   "1"
      Top             =   6840
      Width           =   540
   End
   Begin VB.TextBox LogDateTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   6480
      TabIndex        =   188
      Text            =   "Enter Log Date"
      Top             =   6795
      Width           =   1665
   End
   Begin VB.TextBox LogNameTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   6495
      TabIndex        =   187
      Text            =   "Enter Log Name"
      Top             =   6435
      Width           =   3375
   End
   Begin VB.TextBox AlarmValTxt 
      Height          =   375
      Left            =   7080
      TabIndex        =   184
      Text            =   "5"
      Top             =   5970
      Width           =   405
   End
   Begin VB.Timer AlarmTmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6585
      Top             =   5925
   End
   Begin VB.TextBox WarnValTxt 
      Height          =   375
      Left            =   7065
      TabIndex        =   182
      Text            =   "5"
      Top             =   5505
      Width           =   405
   End
   Begin VB.Timer WarningTmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6585
      Top             =   5475
   End
   Begin VB.Timer CommStatusTmr 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6810
      Top             =   9810
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   181
      Top             =   9450
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   688
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "COM: Disconnected "
            TextSave        =   "COM: Disconnected "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   176
            MinWidth        =   176
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Text            =   "Rx: 0"
            TextSave        =   "Rx: 0"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Text            =   "Tx: 0"
            TextSave        =   "Tx: 0"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   2355
            MinWidth        =   1764
            Text            =   "Data Log: Disabled "
            TextSave        =   "Data Log: Disabled "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   1667
            MinWidth        =   1676
            Text            =   "Records: 0 "
            TextSave        =   "Records: 0 "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame6 
      Caption         =   "LFC 6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   60
      TabIndex        =   172
      Top             =   7470
      Width           =   1845
      Begin VB.CommandButton FlowSet6 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   178
         Top             =   480
         Width           =   660
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Close"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   177
         Top             =   2475
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Open"
         Height          =   255
         Index           =   4
         Left            =   135
         TabIndex        =   176
         Top             =   2190
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Normal"
         Height          =   255
         Index           =   5
         Left            =   135
         TabIndex        =   175
         Top             =   1905
         Width           =   1215
      End
      Begin VB.TextBox FlSet_6 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   174
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox FlCV_6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         MaxLength       =   6
         TabIndex        =   173
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   180
         Top             =   870
         Width           =   1845
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   179
         Top             =   210
         Width           =   1845
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "LFC 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   60
      TabIndex        =   163
      Top             =   5985
      Width           =   1845
      Begin VB.TextBox FlCV_5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         MaxLength       =   6
         TabIndex        =   169
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.TextBox FlSet_5 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   168
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Normal"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   167
         Top             =   1905
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Open"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   166
         Top             =   2190
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Close"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   165
         Top             =   2475
         Width           =   1215
      End
      Begin VB.CommandButton FlowSet5 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   164
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   171
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   170
         Top             =   870
         Width           =   1845
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "LFC 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   60
      TabIndex        =   154
      Top             =   4500
      Width           =   1845
      Begin VB.TextBox FlCV_4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         MaxLength       =   6
         TabIndex        =   160
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.TextBox FlSet_4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   159
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.OptionButton OptMFC4 
         Caption         =   "Normal"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   158
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC4 
         Caption         =   "Open"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   157
         Top             =   2205
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC4 
         Caption         =   "Close"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   156
         Top             =   2490
         Width           =   1215
      End
      Begin VB.CommandButton FlowSet4 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   155
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   162
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   161
         Top             =   870
         Width           =   1845
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "LFC 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   60
      TabIndex        =   145
      Top             =   3015
      Width           =   1845
      Begin VB.TextBox FlCV_3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         MaxLength       =   6
         TabIndex        =   151
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.TextBox FlSet_3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   150
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.OptionButton OptMFC3 
         Caption         =   "Normal"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   149
         Top             =   1890
         Width           =   1095
      End
      Begin VB.OptionButton OptMFC3 
         Caption         =   "Open"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   148
         Top             =   2175
         Width           =   1095
      End
      Begin VB.OptionButton OptMFC3 
         Caption         =   "Close"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   147
         Top             =   2460
         Width           =   1095
      End
      Begin VB.CommandButton FlowSet3 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   146
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   153
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   152
         Top             =   870
         Width           =   1845
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "LFC 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   60
      TabIndex        =   136
      Top             =   1530
      Width           =   1845
      Begin VB.TextBox FlCV_2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         MaxLength       =   6
         TabIndex        =   142
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.TextBox FlSet_2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   141
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.OptionButton OptMFC2 
         Caption         =   "Normal"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   140
         Top             =   1890
         Width           =   1095
      End
      Begin VB.OptionButton OptMFC2 
         Caption         =   "Open"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   139
         Top             =   2175
         Width           =   1095
      End
      Begin VB.OptionButton OptMFC2 
         Caption         =   "Close"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   138
         Top             =   2460
         Width           =   1095
      End
      Begin VB.CommandButton FlowSet2 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   137
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   144
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   143
         Top             =   870
         Width           =   1845
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "LFC 7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   2025
      TabIndex        =   118
      Top             =   45
      Width           =   1845
      Begin VB.CommandButton FlowSet7 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   124
         Top             =   480
         Width           =   660
      End
      Begin VB.OptionButton OptMFC1 
         Caption         =   "Close"
         Height          =   255
         Index           =   3
         Left            =   75
         TabIndex        =   123
         Top             =   2955
         Width           =   1335
      End
      Begin VB.OptionButton OptMFC1 
         Caption         =   "Open"
         Height          =   255
         Index           =   4
         Left            =   75
         TabIndex        =   122
         Top             =   2670
         Width           =   1335
      End
      Begin VB.OptionButton OptMFC1 
         Caption         =   "Normal"
         Height          =   255
         Index           =   5
         Left            =   75
         TabIndex        =   121
         Top             =   2385
         Width           =   1335
      End
      Begin VB.TextBox FlCV_7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         TabIndex        =   120
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.TextBox FlSet_7 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   119
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   126
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   125
         Top             =   870
         Width           =   1845
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "LFC 8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   2025
      TabIndex        =   109
      Top             =   1530
      Width           =   1845
      Begin VB.CommandButton FlowSet8 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   115
         Top             =   480
         Width           =   660
      End
      Begin VB.OptionButton OptMFC2 
         Caption         =   "Close"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   114
         Top             =   2460
         Width           =   1095
      End
      Begin VB.OptionButton OptMFC2 
         Caption         =   "Open"
         Height          =   255
         Index           =   4
         Left            =   135
         TabIndex        =   113
         Top             =   2175
         Width           =   1095
      End
      Begin VB.OptionButton OptMFC2 
         Caption         =   "Normal"
         Height          =   255
         Index           =   5
         Left            =   135
         TabIndex        =   112
         Top             =   1890
         Width           =   1095
      End
      Begin VB.TextBox FlSet_8 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   111
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox FlCV_8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         MaxLength       =   6
         TabIndex        =   110
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   117
         Top             =   870
         Width           =   1845
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   116
         Top             =   210
         Width           =   1845
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "LFC 9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   2025
      TabIndex        =   100
      Top             =   3015
      Width           =   1845
      Begin VB.CommandButton FlowSet9 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   106
         Top             =   480
         Width           =   660
      End
      Begin VB.OptionButton OptMFC3 
         Caption         =   "Close"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   105
         Top             =   2460
         Width           =   1095
      End
      Begin VB.OptionButton OptMFC3 
         Caption         =   "Open"
         Height          =   255
         Index           =   4
         Left            =   150
         TabIndex        =   104
         Top             =   2175
         Width           =   1095
      End
      Begin VB.OptionButton OptMFC3 
         Caption         =   "Normal"
         Height          =   255
         Index           =   5
         Left            =   150
         TabIndex        =   103
         Top             =   1890
         Width           =   1095
      End
      Begin VB.TextBox FlSet_9 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   102
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox FlCV_9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         MaxLength       =   6
         TabIndex        =   101
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   108
         Top             =   870
         Width           =   1845
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   107
         Top             =   210
         Width           =   1845
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "LFC 10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   2025
      TabIndex        =   91
      Top             =   4500
      Width           =   1845
      Begin VB.CommandButton FlowSet10 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   97
         Top             =   480
         Width           =   660
      End
      Begin VB.OptionButton OptMFC4 
         Caption         =   "Close"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   96
         Top             =   2490
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC4 
         Caption         =   "Open"
         Height          =   255
         Index           =   4
         Left            =   135
         TabIndex        =   95
         Top             =   2205
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC4 
         Caption         =   "Normal"
         Height          =   255
         Index           =   5
         Left            =   135
         TabIndex        =   94
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox FlSet_10 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   93
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox FlCV_10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         MaxLength       =   6
         TabIndex        =   92
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   99
         Top             =   870
         Width           =   1845
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   98
         Top             =   210
         Width           =   1845
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "LFC 11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   2025
      TabIndex        =   82
      Top             =   5985
      Width           =   1845
      Begin VB.CommandButton FlowSet11 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   88
         Top             =   480
         Width           =   660
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Close"
         Height          =   255
         Index           =   6
         Left            =   135
         TabIndex        =   87
         Top             =   2475
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Open"
         Height          =   255
         Index           =   7
         Left            =   135
         TabIndex        =   86
         Top             =   2190
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Normal"
         Height          =   255
         Index           =   8
         Left            =   135
         TabIndex        =   85
         Top             =   1905
         Width           =   1215
      End
      Begin VB.TextBox FlSet_11 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   84
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox FlCV_11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         MaxLength       =   6
         TabIndex        =   83
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   90
         Top             =   870
         Width           =   1845
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   89
         Top             =   210
         Width           =   1845
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "LFC 12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   2025
      TabIndex        =   73
      Top             =   7470
      Width           =   1845
      Begin VB.TextBox FlCV_12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         MaxLength       =   6
         TabIndex        =   79
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.TextBox FlSet_12 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   78
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Normal"
         Height          =   255
         Index           =   9
         Left            =   135
         TabIndex        =   77
         Top             =   1905
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Open"
         Height          =   255
         Index           =   10
         Left            =   135
         TabIndex        =   76
         Top             =   2190
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Close"
         Height          =   255
         Index           =   11
         Left            =   135
         TabIndex        =   75
         Top             =   2475
         Width           =   1215
      End
      Begin VB.CommandButton FlowSet12 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   74
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   0
         TabIndex        =   81
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   0
         TabIndex        =   80
         Top             =   870
         Width           =   1845
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "LFC 13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   3975
      TabIndex        =   64
      Top             =   45
      Width           =   1845
      Begin VB.TextBox FlSet_13 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   70
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox FlCV_13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         TabIndex        =   69
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.OptionButton OptMFC1 
         Caption         =   "Normal"
         Height          =   255
         Index           =   6
         Left            =   75
         TabIndex        =   68
         Top             =   2385
         Width           =   1335
      End
      Begin VB.OptionButton OptMFC1 
         Caption         =   "Open"
         Height          =   255
         Index           =   7
         Left            =   75
         TabIndex        =   67
         Top             =   2670
         Width           =   1335
      End
      Begin VB.OptionButton OptMFC1 
         Caption         =   "Close"
         Height          =   255
         Index           =   8
         Left            =   75
         TabIndex        =   66
         Top             =   2955
         Width           =   1335
      End
      Begin VB.CommandButton FlowSet13 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   65
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   0
         TabIndex        =   72
         Top             =   870
         Width           =   1845
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   0
         TabIndex        =   71
         Top             =   210
         Width           =   1845
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "LFC 14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   3975
      TabIndex        =   55
      Top             =   1530
      Width           =   1845
      Begin VB.TextBox FlCV_14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         MaxLength       =   6
         TabIndex        =   61
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.TextBox FlSet_14 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   60
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.OptionButton OptMFC2 
         Caption         =   "Normal"
         Height          =   255
         Index           =   6
         Left            =   135
         TabIndex        =   59
         Top             =   1890
         Width           =   1095
      End
      Begin VB.OptionButton OptMFC2 
         Caption         =   "Open"
         Height          =   255
         Index           =   7
         Left            =   135
         TabIndex        =   58
         Top             =   2175
         Width           =   1095
      End
      Begin VB.OptionButton OptMFC2 
         Caption         =   "Close"
         Height          =   255
         Index           =   8
         Left            =   135
         TabIndex        =   57
         Top             =   2460
         Width           =   1095
      End
      Begin VB.CommandButton FlowSet14 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   56
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   0
         TabIndex        =   63
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   13
         Left            =   0
         TabIndex        =   62
         Top             =   870
         Width           =   1845
      End
   End
   Begin VB.Frame Frame15 
      Caption         =   "LFC 15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   3975
      TabIndex        =   46
      Top             =   3015
      Width           =   1845
      Begin VB.TextBox FlCV_15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         MaxLength       =   6
         TabIndex        =   52
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.TextBox FlSet_15 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   51
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.OptionButton OptMFC3 
         Caption         =   "Normal"
         Height          =   255
         Index           =   6
         Left            =   150
         TabIndex        =   50
         Top             =   1890
         Width           =   1095
      End
      Begin VB.OptionButton OptMFC3 
         Caption         =   "Open"
         Height          =   255
         Index           =   7
         Left            =   150
         TabIndex        =   49
         Top             =   2175
         Width           =   1095
      End
      Begin VB.OptionButton OptMFC3 
         Caption         =   "Close"
         Height          =   255
         Index           =   8
         Left            =   150
         TabIndex        =   48
         Top             =   2460
         Width           =   1095
      End
      Begin VB.CommandButton FlowSet15 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   47
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   0
         TabIndex        =   54
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   14
         Left            =   0
         TabIndex        =   53
         Top             =   870
         Width           =   1845
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "LFC 16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   3975
      TabIndex        =   37
      Top             =   4500
      Width           =   1845
      Begin VB.TextBox FlCV_16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         MaxLength       =   6
         TabIndex        =   43
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.TextBox FlSet_16 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   42
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.OptionButton OptMFC4 
         Caption         =   "Normal"
         Height          =   255
         Index           =   6
         Left            =   135
         TabIndex        =   41
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC4 
         Caption         =   "Open"
         Height          =   255
         Index           =   7
         Left            =   135
         TabIndex        =   40
         Top             =   2205
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC4 
         Caption         =   "Close"
         Height          =   255
         Index           =   8
         Left            =   135
         TabIndex        =   39
         Top             =   2490
         Width           =   1215
      End
      Begin VB.CommandButton FlowSet16 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   38
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   0
         TabIndex        =   45
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   0
         TabIndex        =   44
         Top             =   870
         Width           =   1845
      End
   End
   Begin VB.Frame Frame17 
      Caption         =   "LFC 17"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   3975
      TabIndex        =   28
      Top             =   5985
      Width           =   1845
      Begin VB.TextBox FlCV_17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         MaxLength       =   6
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.TextBox FlSet_17 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   33
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Normal"
         Height          =   255
         Index           =   12
         Left            =   135
         TabIndex        =   32
         Top             =   1905
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Open"
         Height          =   255
         Index           =   13
         Left            =   135
         TabIndex        =   31
         Top             =   2190
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Close"
         Height          =   255
         Index           =   14
         Left            =   135
         TabIndex        =   30
         Top             =   2475
         Width           =   1215
      End
      Begin VB.CommandButton FlowSet17 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   29
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   0
         TabIndex        =   36
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   0
         TabIndex        =   35
         Top             =   870
         Width           =   1845
      End
   End
   Begin VB.Frame Frame18 
      Caption         =   "LFC 18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1450
      Left            =   3975
      TabIndex        =   19
      Top             =   7470
      Width           =   1845
      Begin VB.CommandButton FlowSet18 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1120
         TabIndex        =   25
         Top             =   480
         Width           =   660
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Close"
         Height          =   255
         Index           =   15
         Left            =   135
         TabIndex        =   24
         Top             =   2475
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Open"
         Height          =   255
         Index           =   16
         Left            =   135
         TabIndex        =   23
         Top             =   2190
         Width           =   1215
      End
      Begin VB.OptionButton OptMFC5 
         Caption         =   "Normal"
         Height          =   255
         Index           =   17
         Left            =   135
         TabIndex        =   22
         Top             =   1905
         Width           =   1215
      End
      Begin VB.TextBox FlSet_18 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   75
         MaxLength       =   6
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox FlCV_18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   45
         MaxLength       =   6
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   17
         Left            =   0
         TabIndex        =   27
         Top             =   870
         Width           =   1845
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Flow (g/min)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   0
         TabIndex        =   26
         Top             =   210
         Width           =   1845
      End
   End
   Begin VB.TextBox DataLogSeqTxt 
      Height          =   285
      Left            =   6495
      TabIndex        =   18
      Text            =   "0"
      Top             =   7170
      Width           =   360
   End
   Begin VB.Timer DataLogShpTmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9975
      Top             =   6420
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9975
      Top             =   6000
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6495
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   4485
      Width           =   1455
   End
   Begin VB.Timer T_Delay500ms 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   12210
      Top             =   4050
   End
   Begin VB.Timer AFCTimer 
      Interval        =   500
      Left            =   12210
      Top             =   3090
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   6360
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   195
      Width           =   2775
   End
   Begin VB.Frame Frame90 
      Caption         =   "COM Status"
      Height          =   1425
      Left            =   8850
      TabIndex        =   5
      Top             =   2820
      Width           =   2640
      Begin VB.CommandButton Command8 
         Caption         =   "Clear Counter"
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   360
         Shape           =   3  'Circle
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Rx"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Tx"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "0"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Timer RFVTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12210
      Top             =   2130
   End
   Begin VB.Frame Frame91 
      Caption         =   "COM Setting"
      Height          =   1425
      Left            =   6435
      TabIndex        =   0
      Top             =   2820
      Width           =   2145
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "MFCFLOWCONTROL.frx":628A
         Left            =   960
         List            =   "MFCFLOWCONTROL.frx":628C
         TabIndex        =   2
         Text            =   "COM1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Text            =   "38400"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Port"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Baud Rate"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9330
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   38400
      ParitySetting   =   1
      DataBits        =   7
      SThreshold      =   1
   End
   Begin VB.Frame Frame92 
      Height          =   1425
      Left            =   6345
      TabIndex        =   10
      Top             =   1125
      Width           =   5715
      Begin VB.CommandButton Command9 
         Caption         =   "OPEN Comm"
         Height          =   375
         Left            =   75
         TabIndex        =   17
         Top             =   150
         Width           =   1250
      End
      Begin VB.CommandButton Command10 
         Caption         =   "CLOSE Comm"
         Height          =   375
         Left            =   1365
         TabIndex        =   16
         Top             =   150
         Width           =   1250
      End
   End
   Begin VB.Label RecordCountLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8985
      TabIndex        =   190
      Top             =   6870
      Width           =   210
   End
   Begin VB.Label Label29 
      Caption         =   "1"
      Height          =   255
      Index           =   3
      Left            =   7620
      TabIndex        =   186
      Top             =   5565
      Width           =   435
   End
   Begin VB.Label Label29 
      Caption         =   "1"
      Height          =   255
      Index           =   2
      Left            =   7665
      TabIndex        =   185
      Top             =   6045
      Width           =   1680
   End
   Begin VB.Label Label29 
      Caption         =   "Alarms and Warnings"
      Height          =   255
      Index           =   1
      Left            =   6330
      TabIndex        =   183
      Top             =   5190
      Width           =   1605
   End
   Begin VB.Label Label32 
      Caption         =   "AFC Timer"
      Height          =   255
      Left            =   12810
      TabIndex        =   14
      Top             =   3210
      Width           =   975
   End
   Begin VB.Label Label29 
      Caption         =   "RFV Timer"
      Height          =   255
      Index           =   0
      Left            =   12810
      TabIndex        =   13
      Top             =   2250
      Width           =   855
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Setup 
      Caption         =   "Setup"
      Begin VB.Menu Comm 
         Caption         =   "COMM Settings"
      End
      Begin VB.Menu Alarms 
         Caption         =   "Alarms Setup"
      End
      Begin VB.Menu DataLog 
         Caption         =   "Data Log Setup"
      End
   End
   Begin VB.Menu LogData 
      Caption         =   "Data Logging"
      Begin VB.Menu StartLog 
         Caption         =   "Start Data Collecting"
      End
      Begin VB.Menu StopLog 
         Caption         =   "Stop Data Collecting"
      End
      Begin VB.Menu OpenLog 
         Caption         =   "Open Data Log"
      End
   End
   Begin VB.Menu AlarmsStatus 
      Caption         =   "Alarms"
      Begin VB.Menu EnableAlarms 
         Caption         =   "Enable"
      End
      Begin VB.Menu DisableAlarms 
         Caption         =   "Disable"
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SW_SHOW = 5

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public MFC_Counter As Integer
Public FeedbackType As Integer  '1=RFV, 2=Mode 3= Flow Set, 4=FS
Public RFS_Read  As Boolean
Dim RFS_ARY(9) As Byte
Dim RFV_ARY(9) As Byte
Dim FlagFlowSP As Variant
Public FlowSetNo As Integer
Public CurrCMD As String
Public FlSetStatus As Boolean
Public MFC1_NewSP, MFC2_NewSP, MFC3_NewSP, MFC4_NewSP, MFC5_NewSP _
, MFC6_NewSP, MFC7_NewSP, MFC8_NewSP, MFC9_NewSP, MFC10_NewSP, MFC11_NewSP _
, MFC12_NewSP, MFC13_NewSP, MFC14_NewSP, MFC15_NewSP, MFC16_NewSP _
, MFC17_NewSP, MFC18_NewSP As Integer
Public LogMaker As Boolean

Private Sub ASCII_CMD_Click()
''Open port
    MSComm1.CommPort = Combo1.Text 'ListIndex + 1
    MSComm1.Settings = Str(Combo2) + "O" + "7" + "1"
    MSComm1.PortOpen = True
 MSComm1.InBufferCount = 0   '清除输入缓冲区的数据
End Sub

Private Sub AFCTimer_Timer()

    Dim strBuff As String
    Dim CMD_Buf As String
    Dim i As Integer
    Dim j As Integer
    Dim AFC_ARY() As Byte
    Dim StrFlowText As String
    Dim StrFlow() As Byte
    Dim FlagSetFlow(30) As Byte
   
    FlagSetFlow(0) = &H40
    FlagSetFlow(1) = &H30
    FlagSetFlow(2) = &H30       'Address
    FlagSetFlow(3) = &H2
    FlagSetFlow(4) = &H41       'Command
    FlagSetFlow(5) = &H46
    FlagSetFlow(6) = &H43
        'Clear send buffer for flow set command
        
    If MFC_Counter = 18 Then 'changed from 6 to 18 (08/02/14)
        MFC_Counter = 1
        AFCTimer.Enabled = False
        RFVTimer.Enabled = True
    End If
    
    MSComm1.OutBufferCount = 0
    
    'get ASC address
    FlagSetFlow(2) = FlagSetFlow(2) + MFC_Counter
    
    
    'get data from input
    Select Case MFC_Counter
                
        Case 1
            StrFlowText = FlSet_1.Text
          
        Case 2
            StrFlowText = FlSet_2.Text
            
        Case 3
            StrFlowText = FlSet_3.Text
            
        Case 4
            StrFlowText = FlSet_4.Text
            
        Case 5
            StrFlowText = FlSet_5.Text
            
        Case 6
            StrFlowText = FlSet_6.Text
            
        Case 7
            StrFlowText = FlSet_7.Text
            
        Case 8
            StrFlowText = FlSet_8.Text
            
        Case 9
            StrFlowText = FlSet_9.Text
            
        Case 10
            StrFlowText = FlSet_10.Text
            
        Case 11
            StrFlowText = FlSet_11.Text
            
        Case 12
            StrFlowText = FlSet_12.Text
            
        Case 13
            StrFlowText = FlSet_13.Text
            
        Case 14
            StrFlowText = FlSet_14.Text
            
        Case 15
            StrFlowText = FlSet_15.Text
            
        Case 16
            StrFlowText = FlSet_16.Text
            
        Case 17
            StrFlowText = FlSet_17.Text
            
        Case 18
            StrFlowText = FlSet_18.Text

    End Select
    
    'get length of flow setpoint
    IntFlowLen = Len(StrFlowText)
    
    'calc flag length
    LenFlag = 7 + IntFlowLen + 4
    LenFlag = 7 + IntFlowLen + 3
    
    StrFlow() = StrConv(StrFlowText, vbFromUnicode)
    
    For i = 0 To IntFlowLen - 1
        FlagSetFlow(7 + i) = StrFlow(i)
    
    'FlagSetFlow(LenFlag - 4) = &
    'FlagSetFlow(LenFlag - 3) = &H42
    'FlagSetFlow(LenFlag - 2) = &H3
    
    
    'FlagSetFlow(LenFlag - 4) = &H2C
    FlagSetFlow(LenFlag - 3) = &H42  'NEW 08/03 for more than 9 MHCs
    FlagSetFlow(LenFlag - 2) = &H3 'NEW 08/03 for more than 9 MHCs
    
    'get check sum
    ASCChcSum = 0
    
    For X = 4 To LenFlag - 2
    
        VarChcSum = VarChcSum + FlagSetFlow(X)
    Next
    
    ByChcSum = VarChcSum Mod 128
     
    FlagSetFlow(LenFlag - 1) = ByChcSum
    
    'get flag string
    For j = 0 To LenFlag - 1
    
        strBuff1 = strBuff1 & Chr(FlagSetFlow(j))
    Next
    Next  ' won't compile unless I include a second NEXT
    
    
    MSComm1.Output = strBuff1
    
    MFC_Counter = 1 + MFC_Counter
    
    Text1.Text = strBuff1
    
    On Error GoTo uerror
    
    MSComm1.Output = strBuff
    Label11.Caption = Label11.Caption + Len(strBuff) '发送计数
        
    ClearErrCounter
    
uerror: Exit Sub

End Sub

Private Sub Alarms_Click()
    AlarmSetUpFrm.Visible = True
End Sub

Private Sub Combo1_Change()
MSComm1.CommPort = Combo1.Text 'ListIndex + 1 '读取com口号
End Sub

Private Sub Combo1_Click()
If MSComm1.PortOpen = True Then  '如果串口打开先关闭后再进行其他操作
   MSComm1.PortOpen = False
 End If
MSComm1.CommPort = Combo1.ListIndex + 1 '读取com口号
End Sub

Private Sub Comm_Click()
    CommFrm.Visible = True
End Sub

Private Sub Command10_Click() 'old

    FlSet_1.Enabled = False
    FlSet_2.Enabled = False
    FlSet_3.Enabled = False
    FlSet_4.Enabled = False
    FlSet_5.Enabled = False
    Combo1.Enabled = True
    Combo2.Enabled = True

    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False

    End If
    
    
    Shape1.FillColor = &HFFFFC0    '灯颜色改变
    RFVTimer.Enabled = False
    
    Command9.Enabled = True         '使能启动按钮
    Command10.Enabled = False       '关闭停止按钮
    
    PageOperation (False)
        
End Sub
Private Sub Command8_Click() '清零计数器
Label10.Caption = 0
Label11.Caption = 0
End Sub

Private Sub Command9_Click()
     On Error GoTo uerror      '发现错误跳转到错误处理
     
    MSComm1.CommPort = Combo1.Text 'ListIndex + 1
    MSComm1.Settings = Str(Combo2) + "O" + "7" + "1"
    MSComm1.PortOpen = True

    Command9.Enabled = False         '停止启动按钮
    Command10.Enabled = True         '使能停止按钮
    
    Shape1.FillColor = &HFF00&
    RFVTimer.Enabled = True         '开始通讯
    
    RFS_Read = False
    
    PageOperation (True)
    
    OptMFC1(0).Value = True
    OptMFC2(0).Value = True
    OptMFC3(0).Value = True
    OptMFC4(0).Value = True
    OptMFC5(0).Value = True
    
 Exit Sub
 

uerror:
    msg$ = "Check connections or change Port"          '错误显示
    Title$ = "MFC Flow Control"
    X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
End Sub

Private Sub CommStatusTmr_Timer()
If StatusBar1.Panels.Item(2).Bevel = sbrInset Then
    StatusBar1.Panels.Item(2).Bevel = sbrRaised
Else
    StatusBar1.Panels.Item(2).Bevel = sbrInset
End If

End Sub

Private Sub DataLog_Click()
DataLogFrm.Visible = True
End Sub

Private Sub DisableAlarms_Click()
    WarningTmr.Enabled = False
    AlarmTmr.Enabled = False
    'clear any existing alarm conditions
    Frame1.BackColor = &H8000000F
    FlCV_1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    FlCV_2.BackColor = &H8000000F
    Frame3.BackColor = &H8000000F
    FlCV_3.BackColor = &H8000000F
    Frame4.BackColor = &H8000000F
    FlCV_4.BackColor = &H8000000F
    Frame5.BackColor = &H8000000F
    FlCV_5.BackColor = &H8000000F
    Frame6.BackColor = &H8000000F
    FlCV_6.BackColor = &H8000000F
    Frame7.BackColor = &H8000000F
    FlCV_7.BackColor = &H8000000F
    Frame8.BackColor = &H8000000F
    FlCV_8.BackColor = &H8000000F
    Frame9.BackColor = &H8000000F
    FlCV_9.BackColor = &H8000000F
    Frame10.BackColor = &H8000000F
    FlCV_10.BackColor = &H8000000F
    Frame11.BackColor = &H8000000F
    FlCV_11.BackColor = &H8000000F
    Frame12.BackColor = &H8000000F
    FlCV_12.BackColor = &H8000000F
    Frame13.BackColor = &H8000000F
    FlCV_13.BackColor = &H8000000F
    Frame14.BackColor = &H8000000F
    FlCV_14.BackColor = &H8000000F
    Frame15.BackColor = &H8000000F
    FlCV_15.BackColor = &H8000000F
    Frame16.BackColor = &H8000000F
    FlCV_16.BackColor = &H8000000F
    Frame17.BackColor = &H8000000F
    FlCV_17.BackColor = &H8000000F
    Frame18.BackColor = &H8000000F
    FlCV_18.BackColor = &H8000000F
End Sub

Private Sub EmoCmd_Click()
'force all setpoints to "0" and send command
FlSet_1.Text = "0.00"
FlowSet1.Value = True
FlSet_2.Text = "0.00"
FlowSet2.Value = True
FlSet_3.Text = "0.00"
FlowSet3.Value = True
FlSet_4.Text = "0.00"
FlowSet4.Value = True
FlSet_5.Text = "0.00"
FlowSet5.Value = True
FlSet_6.Text = "0.00"
FlowSet6.Value = True
FlSet_7.Text = "0.00"
FlowSet7.Value = True
FlSet_8.Text = "0.00"
FlowSet8.Value = True
FlSet_9.Text = "0.00"
FlowSet9.Value = True
FlSet_10.Text = "0.00"
FlowSet10.Value = True
FlSet_11.Text = "0.00"
FlowSet11.Value = True
FlSet_12.Text = "0.00"
FlowSet12.Value = True
FlSet_13.Text = "0.00"
FlowSet13.Value = True
FlSet_14.Text = "0.00"
FlowSet14.Value = True
FlSet_15.Text = "0.00"
FlowSet15.Value = True
FlSet_16.Text = "0.00"
FlowSet16.Value = True
FlSet_17.Text = "0.00"
FlowSet17.Value = True
FlSet_18.Text = "0.00"
FlowSet18.Value = True
End Sub

Private Sub EnableAlarms_Click()
    WarningTmr.Enabled = True
    AlarmTmr.Enabled = True
End Sub

Private Sub Exit_Click()
UnloadAll
End
End Sub

Private Sub FlowSet1_Click()
    If FlSet_1.Text = "" Then
        FlSet_1.Text = 0
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_1.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
    
    Else
        FlowSetNo = 1
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet2_Click()
        If FlSet_2.Text = "" Then
            msg$ = "Value Cannot be Empty"          '错误显示
            Title$ = "Error info"
            X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
        ElseIf Val(FlSet_2.Text) > 0.2 Then 'full range scale of LFM
    
            msg$ = "Value Cannot be more than Max scale"          '错误显示
            Title$ = "Error info"
            X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
        Else
            FlowSetNo = 2
            FlagFlowSet (FlowSetNo)
        End If
End Sub

Private Sub FlowSet3_Click()
    If FlSet_3.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
    
    ElseIf Val(FlSet_3.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
    
    Else
        FlowSetNo = 3
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet4_Click()
    If FlSet_4.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_4.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
    
    Else
        FlowSetNo = 4
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet5_Click()
    If FlSet_5.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_5.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标

    Else
        FlowSetNo = 5
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet6_Click()
    If FlSet_6.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_6.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标

    Else
        FlowSetNo = 6
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet7_Click()
    If FlSet_7.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_7.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标

    Else
        FlowSetNo = 7
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet8_Click()
    If FlSet_8.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_8.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标

    Else
        FlowSetNo = 8
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet9_Click()
    If FlSet_9.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_9.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标

    Else
        FlowSetNo = 9
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet10_Click()
    If FlSet_10.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_10.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标

    Else
        FlowSetNo = 10
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet11_Click()
    If FlSet_11.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_11.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标

    Else
        FlowSetNo = 11
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet12_Click()
    If FlSet_12.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_12.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标

    Else
        FlowSetNo = 12
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet13_Click()
    If FlSet_13.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_13.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标

    Else
        FlowSetNo = 13
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet14_Click()
    If FlSet_14.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_14.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标

    Else
        FlowSetNo = 14
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet15_Click()
    If FlSet_15.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_15.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标

    Else
        FlowSetNo = 15
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet16_Click()
    If FlSet_16.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_16.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标

    Else
        FlowSetNo = 16
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet17_Click()
    If FlSet_17.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_17.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标

    Else
        FlowSetNo = 17
        FlagFlowSet (FlowSetNo)
    End If
End Sub

Private Sub FlowSet18_Click()
    If FlSet_18.Text = "" Then
        msg$ = "Value Cannot be Empty"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
        
    ElseIf Val(FlSet_18.Text) > 0.2 Then 'full range scale of LFM
    
        msg$ = "Value Cannot be more than Max scale"          '错误显示
        Title$ = "Error info"
        X = MsgBox(msg$, 48, Title$) '48标示显示警告图标

    Else
        FlowSetNo = 18
        FlagFlowSet (FlowSetNo)
    End If
End Sub
Private Sub FlSet_1_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_1, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select

End Sub
Private Sub FlSet_2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_2, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_3_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_3, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_4_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_4, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_5_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_5, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_6_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_6, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_7_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_7, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_8_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_8, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_9_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_9, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_10_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_10, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_11_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_11, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_12_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_12, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_13_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_13, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_14_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_14, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_15_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_15, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_16_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_16, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_17_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_17, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub
Private Sub FlSet_18_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8
           Case 46
                If InStr(FlSet_18, ".") <> 0 Then KeyAscii = 0
           Case 48 To 57
           Case Else
                KeyAscii = 0
    End Select
End Sub

Private Sub Form_Load()
 'On Error GoTo uerror      '发现错误跳转到错误处理
 
MFC_Counter = 0         '初始化MFC_Counter
FlSetStatus = False

If MSComm1.PortOpen = True Then
   MSComm1.PortOpen = False
Else
End If

Combo1.AddItem "COM1"
Combo1.AddItem "COM2"
Combo1.AddItem "COM3"
Combo1.AddItem "COM4"
Combo1.AddItem "COM5"
Combo1.AddItem "COM6"
Combo1.AddItem "COM7"
Combo1.AddItem "COM8"
Combo1.AddItem "COM9"
Combo1.AddItem "COM10"
Combo1.AddItem "COM11"
Combo1.AddItem "COM12"
Combo1.AddItem "COM13"
Combo1.AddItem "COM14"
Combo1.AddItem "COM15"
Combo1.AddItem "COM16"
Combo1.ListIndex = 0

MSComm1.CommPort = Combo1.ListIndex + 1
MSComm1.Settings = "38400,O,7,1"

Shape1.FillColor = &H808080
   
Command10.Enabled = False

AFCTimer.Enabled = False

Combo2.AddItem "38400"
Combo2.AddItem "28800"
Combo2.AddItem "19200"
Combo2.AddItem "9600"


FeedbackType = 1
RFS_Read = False

PageOperation (False)
StartLog.Enabled = False
StopLog.Enabled = False

 Exit Sub
 

uerror:
    msg$ = "无效端口号"          '错误显示
    Title$ = "错误提示"
    X = MsgBox(msg$, 48, Title$) '48标示显示警告图标
End Sub

Private Sub Label11_Change()
Dim NumTx As Single
NumTx = Label11.Caption
If NumTx > 10000 Then
    Command8.Value = True
End If
StatusBar1.Panels(3).Text = "Rx: " & Label10.Caption
StatusBar1.Panels(4).Text = "Tx: " & Label11.Caption
End Sub

Private Sub LFC1VlvOverRideCmd_Click()
If Frame1.Height = 1450 Then
    Frame1.Height = 2355
    Frame1.ZOrder (0)
Else
    Frame1.Height = 1450
End If
End Sub

Private Sub MSComm1_OnComm()
    
    On Error GoTo uerror
    
    DelayTime
    
    Dim BytReceived() As Byte
    Dim strBuff As String
    Dim i As Integer
    Dim FullScale As Integer
    Dim FlCV As Single
    Dim ASC_02, ASC_03 As String
    Dim LenstrBuff, IntFlowLen As Integer
    Dim STXposi, ETXposi As Integer
    Dim RFV_Value As String
    Dim CsngShapHgt As Single
    Dim IntShapHgt As Integer
    Dim FullScle1 As Integer 'FullScale2, FullScale3, FullScale4, FullScale5 As Integer
    Dim Now_Time As String
    Dim ReciValue As String
    Dim FeedbackMode As Integer
    
    FullScle1 = 0.2
    'FullScle2 = 10
    'FullScle3 = 3
    'FullScle4 = 1
    'FullScle5 = 300
    STXposi = 0
    ETXposi = 0
    FlCV = 0
    RFV_Value = ""
    ASC_02 = Chr(2)
    ASC_03 = Chr(3)
    FeedbackMode = 1
    
    
    MSComm1.InputLen = 0     '读入缓冲区全部内容
    strBuff = MSComm1.Input  '读入到缓冲区
    MSComm1.InBufferCount = 0   '清除输入缓存区的数据
    
    '接收数据
'    Text2.Text = strBuff

    '接收的数据长度
    LenstrBuff = Len(strBuff)
    
    Label10.Caption = Label10.Caption + LenstrBuff '接收计数
    
    If LogMaker = True Then
        Open "C:\Logdata.txt" For Append As 7
        Print #7, Now(); vbTab + "MFC" + MFC_Counter; vbTab + strBuff
        Close #7
    End If
    
    '在输入缓冲区中搜索开始和结束的字符
    STXposi = InStr(1, strBuff, ASC_02, vbTextCompare)
    ETXposi = InStr(1, strBuff, ASC_03, vbTextCompare)
    
    If ETXposi - STXposi > 0 And STXposi <> 0 And ETXposi <> 0 Then
    
        STXposi = InStr(1, strBuff, ASC_02, vbTextCompare)
        ETXposi = InStr(1, strBuff, ASC_03, vbTextCompare)
        IntFlowLen = ETXposi - STXposi - 1
        ReciValue = Mid(strBuff, STXposi + 1, IntFlowLen)
    Else
        ERRCnt.Caption = ERRCnt.Caption + 1

    Exit Sub
        
    End If
    
    
    Select Case FeedbackMode
        
        Case 1          'RFV
        
            FlCV = Val(ReciValue)
                       
            Select Case MFC_Counter
                Case 1      'MFC1 RFV
                    FlCV = FlCV
                    FlCV_1.Text = Format(FlCV, "0.000")
                Case 2      'MFC2 RFV
                    FlCV = FlCV
                    FlCV_2.Text = Format(FlCV, "0.000")
                Case 3         'MFC3 RFV
                    FlCV = FlCV
                    FlCV_3.Text = Format(FlCV, "0.000")
                Case 4      'MFC4 RFV
                    FlCV = FlCV
                    FlCV_4.Text = Format(FlCV, "0.000")
                Case 5       'MFC5 RFV
                    FlCV = FlCV
                    FlCV_5.Text = Format(FlCV, "0.000")
                Case 6       'MFC6 RFV  (added here down 08/02/14)
                    FlCV = FlCV
                    FlCV_6.Text = Format(FlCV, "0.000")
                Case 7       'MFC7 RFV
                    FlCV = FlCV
                    FlCV_7.Text = Format(FlCV, "0.000")
                Case 8       'MFC5 RFV
                    FlCV = FlCV
                    FlCV_8.Text = Format(FlCV, "0.000")
                Case 9       'MFC9 RFV
                    FlCV = FlCV
                    FlCV_9.Text = Format(FlCV, "0.000")
                Case 10       'MFC10 RFV
                    FlCV = FlCV
                    FlCV_10.Text = Format(FlCV, "0.000")
                Case 11       'MFC11 RFV
                    FlCV = FlCV
                    FlCV_11.Text = Format(FlCV, "0.000")
                Case 12       'MFC12 RFV
                    FlCV = FlCV
                    FlCV_12.Text = Format(FlCV, "0.000")
                Case 13       'MFC13 RFV
                    FlCV = FlCV
                    FlCV_13.Text = Format(FlCV, "0.000")
                Case 14       'MFC14 RFV
                    FlCV = FlCV
                    FlCV_14.Text = Format(FlCV, "0.000")
                Case 15       'MFC15 RFV
                    FlCV = FlCV
                    FlCV_15.Text = Format(FlCV, "0.000")
                Case 16       'MFC16 RFV
                    FlCV = FlCV
                    FlCV_16.Text = Format(FlCV, "0.000")
                Case 17       'MFC17 RFV
                    FlCV = FlCV
                    FlCV_17.Text = Format(FlCV, "0.000")
                Case 18       'MFC18 RFV
                    FlCV = FlCV
                    FlCV_18.Text = Format(FlCV, "0.000")
            End Select
        Case 19, 20       'MVO, MVC, IAC, AFC 'changed from 2, 3 to 19, 20 (08/02/14)
        Case 21          'FS changed from 4 to 20 (08/02/14)
    End Select
uerror:    Exit Sub
    
End Sub

Private Sub OpenLog_Click()
ShellExecute Me.hwnd, "open", App.Path + "\Data\NewFile.csv", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub OptMFC1_Click(Index As Integer)
Dim tOptMFC As OptionButton

    For Each tOptMFC In OptMFC1
    
        If tOptMFC.Value Then Call MFCModeChange(1, tOptMFC.Index)
        
    Next
    
End Sub

Private Sub OptMFC2_Click(Index As Integer)
Dim tOptMFC As OptionButton

    For Each tOptMFC In OptMFC2
    
        If tOptMFC.Value Then Call MFCModeChange(2, tOptMFC.Index)
        
    Next
End Sub

Private Sub OptMFC3_Click(Index As Integer)

Dim tOptMFC As OptionButton

    For Each tOptMFC In OptMFC3
    
        If tOptMFC.Value Then Call MFCModeChange(3, tOptMFC.Index)
        
    Next
    
End Sub

Private Sub OptMFC4_Click(Index As Integer)

Dim tOptMFC As OptionButton

    For Each tOptMFC In OptMFC4
    
        If tOptMFC.Value Then Call MFCModeChange(4, tOptMFC.Index)
        
    Next
    
End Sub

Private Sub OptMFC5_Click(Index As Integer)

Dim tOptMFC As OptionButton

    For Each tOptMFC In OptMFC5
    
        If tOptMFC.Value Then Call MFCModeChange(5, tOptMFC.Index)
        
    Next
    
End Sub

Private Sub RecordCountLbl_Change()
    StatusBar1.Panels.Item(6).Text = "Records: " & RecordCountLbl.Caption & " "
End Sub

Private Sub RFVTimer_Timer()

    Dim strBuff As String
    Dim CMD_Buf As String
    Dim MFC1Flw1 As Single
    'Dim MFC1Flw2 As Single
    'Dim MFC1Flw3 As Single
    'Dim MFC1Flw4 As Single
    'Dim MFC1Flw5 As Single

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    
    'Read Full Scal
    RFS_ARY(0) = &H40
    RFS_ARY(1) = &H30
    RFS_ARY(2) = &H31
    RFS_ARY(3) = &H2
    RFS_ARY(4) = &H52
    RFS_ARY(5) = &H46
    RFS_ARY(6) = &H53
    RFS_ARY(7) = &H3
    RFS_ARY(8) = &H6E
    
    'Read Current Flow
    RFV_ARY(0) = &H40
    RFV_ARY(1) = &H30
    RFV_ARY(2) = &H31
    RFV_ARY(3) = &H2
    RFV_ARY(4) = &H52
    RFV_ARY(5) = &H46
    RFV_ARY(6) = &H56
    RFV_ARY(7) = &H3
    RFV_ARY(8) = &H71
    
    MFC_Counter = 1 + MFC_Counter
    
    MSComm1.InBufferCount = 0   'Clear Data in input buffer

    
    If RFS_Read = False Then
        
        If MFC_Counter = 18 Then 'changed from 6 to 18 (08/02/14)
            MFC_Counter = 1
            
        End If
        
        If MFC_Counter > 9 Then 'new code for LFC
            k = Int(MFC_Counter / 10)
            RFV_ARY(1) = RFV_ARY(1) + k
            RFV_ARY(2) = RFV_ARY(2) + MFC_Counter - k * 10 - 1
        Else
        
        'Combine RFV
        RFV_ARY(2) = RFV_ARY(2) + MFC_Counter - 1
        End If
        
        For j = 0 To 8
            CMD_Buf = CMD_Buf & Chr(RFV_ARY(j))
        Next
        
    End If
    
    strBuff = CMD_Buf
    Text1.Text = CMD_Buf
    
    On Error GoTo uerror
    
    MSComm1.Output = strBuff 'READ CURRENT FLOW
    Label11.Caption = Label11.Caption + Len(strBuff) '发送计数
        
'If FlCV_1.Text > 0 Then
    'MFC1Flw1 = FlCV_1.Text
'Else
    'MFC1Flw1 = 0
'End If
'If FlCV_2.Text > 0 Then
    'MFC1Flw2 = FlCV_2.Text
'Else
    'MFC1Flw2 = 0
'End If
'If FlCV_3.Text > 0 Then
    'MFC1Flw3 = FlCV_3.Text
'Else
    'MFC1Flw3 = 0
'End If
'If FlCV_4.Text > 0 Then
    'MFC1Flw4 = FlCV_4.Text
'Else
    'MFC1Flw4 = 0
'End If
'If FlCV_5.Text > 0 Then
    'MFC1Flw5 = FlCV_5.Text
'Else
    'MFC1Flw5 = 0
'End If
uerror:
End Sub
Private Sub FlagFlowSet(ByVal FlowSP As Integer)
    Dim FlowSetSng As Single
    Dim FlowSetStr As String
    Dim FlagSetFlow(30) As Byte
    Dim StrFlow() As Byte
    Dim ChkSum As Byte
    Dim StrFlowText As String
    Dim IntFlowLen As Integer
    Dim X, Y, i, j As Integer
    Dim strBuff(20) As Byte
    Dim LenFlag As Integer
    Dim intChcSum As Integer
    Dim VarChcSum As Integer
    Dim ByChcSum As Byte
    Dim strBuff1 As String
    
    FlagSetFlow(0) = &H40
    FlagSetFlow(1) = &H30
    FlagSetFlow(2) = &H30       'Address
    FlagSetFlow(3) = &H2
    FlagSetFlow(4) = &H41       'Command
    FlagSetFlow(5) = &H46
    FlagSetFlow(6) = &H43
    
    RFVTimer = False
    T_Delay500ms = True
    
    'Clear send buffer for flow set command
    MSComm1.OutBufferCount = 0
       
    'get ASC address
    FlagSetFlow(2) = FlagSetFlow(2) + FlowSP
    
    
    'get data from input
    Select Case FlowSP
                
        Case 1
            StrFlowText = Trim(Str(FlSet_1.Text))
        Case 2
            StrFlowText = Trim(Str(FlSet_2.Text))
        Case 3
            StrFlowText = Trim(Str(FlSet_3.Text))
        Case 4
            StrFlowText = Trim(Str(FlSet_4.Text))
        Case 5
            StrFlowText = Trim(Str(FlSet_5.Text))
        Case 6
            StrFlowText = Trim(Str(FlSet_6.Text))
        Case 7
            StrFlowText = Trim(Str(FlSet_7.Text))
        Case 8
            StrFlowText = Trim(Str(FlSet_8.Text))
        Case 9
            StrFlowText = Trim(Str(FlSet_9.Text))
        Case 10
            StrFlowText = Trim(Str(FlSet_10.Text))
        Case 11
            StrFlowText = Trim(Str(FlSet_11.Text))
        Case 12
            StrFlowText = Trim(Str(FlSet_12.Text))
        Case 13
            StrFlowText = Trim(Str(FlSet_13.Text))
        Case 14
            StrFlowText = Trim(Str(FlSet_14.Text))
        Case 15
            StrFlowText = Trim(Str(FlSet_15.Text))
        Case 16
            StrFlowText = Trim(Str(FlSet_16.Text))
        Case 17
            StrFlowText = Trim(Str(FlSet_17.Text))
        Case 18
            StrFlowText = Trim(Str(FlSet_18.Text))
    End Select
    
    'get length of flow setpoint
    IntFlowLen = Len(StrFlowText)
    
    'calc flag length
    LenFlag = 7 + IntFlowLen + 4
    
    StrFlow() = StrConv(StrFlowText, vbFromUnicode)
    
    For i = 0 To IntFlowLen - 1
        FlagSetFlow(7 + i) = StrFlow(i)
    Next
    
    'FlagSetFlow(LenFlag - 4) = &H2C
    FlagSetFlow(LenFlag - 3) = &H42  'NEW 08/03 for more than 9 MHCs
    FlagSetFlow(LenFlag - 2) = &H3 'NEW 08/03 for more than 9 MHCs
    
    'get check sum
    ASCChcSum = 0
    
    For X = 4 To LenFlag - 2
    
        VarChcSum = VarChcSum + FlagSetFlow(X)
    Next
    
    ByChcSum = VarChcSum Mod 128
     
    FlagSetFlow(LenFlag - 1) = ByChcSum
    
    'get flag string
    For j = 0 To LenFlag - 1
    
        strBuff1 = strBuff1 & Chr(FlagSetFlow(j)) 'strBuff1 "@01AFC0.1[B]" '
    Next
    
    
    MSComm1.Output = strBuff1 'send new flow rate
    Text1.Text = strBuff1
    
End Sub
Private Sub DelayTime()
  Dim bDT As Boolean
  Dim sPrevious As Single, sLast As Single
  bDT = True
  sPrevious = Timer '(Timer可以计算从子夜到现在所经过的秒数，在Microsoft Windows中，Timer函数可以返回一秒的小数部分)
  Do While bDT
    If Timer - sPrevious >= 0.05 Then bDT = False
  Loop
  bDT = True
End Sub
Private Sub ClearErrCounter()

    If Label10.Caption > 10000 Or Label11.Caption > 10000 Then
        
        Label10.Caption = 0
        Label11.Caption = 0
        ERRCnt.Caption = 0
    End If
    
End Sub
Private Sub MFCModeChange(sMFCNo As Integer, sNewMode As Integer)
    
    Dim ModeFrame(30) As Byte
    Dim tChkSum As Long
    
    'MsgBox sMFCNo & sNewMode
    
    
    RFVTimer = False
    T_Delay500ms = True
    
    ModeFrame(0) = &H40
    ModeFrame(1) = &H30
    ModeFrame(2) = &H30 + sMFCNo
    ModeFrame(3) = &H2
    
    Select Case sNewMode
        
        Case 0
            ModeFrame(4) = Asc("I")     'H49
            ModeFrame(5) = Asc("A")     'H41
            ModeFrame(6) = Asc("C")     'H43
            
        Case 1
            ModeFrame(4) = Asc("M")     'H4D
            ModeFrame(5) = Asc("V")     'H56
            ModeFrame(6) = Asc("O")     'H4F
            
        Case 2
            ModeFrame(4) = Asc("M")     'H4D
            ModeFrame(5) = Asc("V")     'H56
            ModeFrame(6) = Asc("C")     'H43
            
    End Select
    
    ModeFrame(7) = &H3
    
    tChkSum = ModeFrame(4) + ModeFrame(5) + ModeFrame(6) + ModeFrame(7)
    ModeFrame(8) = tChkSum Mod 128
    
    Text1.Text = Chr(ModeFrame(0)) & Chr(ModeFrame(1)) & Chr(ModeFrame(2)) & Chr(ModeFrame(3)) & Chr(ModeFrame(4)) & Chr(ModeFrame(5)) & Chr(ModeFrame(6)) & Chr(ModeFrame(7)) & Chr(ModeFrame(8))
    
        'Clear send buffer for flow set command
    MSComm1.OutBufferCount = 0
    MSComm1.Output = ModeFrame
End Sub

Private Sub PageOperation(sStatus As Boolean)
Dim OperationStatus As Boolean

OperationStatus = sStatus

FlSet_1.Enabled = sStatus
FlowSet1.Enabled = sStatus

FlSet_2.Enabled = sStatus
FlowSet2.Enabled = sStatus

FlSet_3.Enabled = sStatus
FlowSet3.Enabled = sStatus

FlSet_4.Enabled = sStatus
FlowSet4.Enabled = sStatus

FlSet_5.Enabled = sStatus
FlowSet5.Enabled = sStatus

FlSet_6.Enabled = sStatus
FlowSet6.Enabled = sStatus

FlSet_7.Enabled = sStatus
FlowSet7.Enabled = sStatus

FlSet_8.Enabled = sStatus
FlowSet8.Enabled = sStatus

FlSet_9.Enabled = sStatus
FlowSet9.Enabled = sStatus

FlSet_10.Enabled = sStatus
FlowSet10.Enabled = sStatus

FlSet_11.Enabled = sStatus
FlowSet11.Enabled = sStatus

FlSet_12.Enabled = sStatus
FlowSet12.Enabled = sStatus

FlSet_13.Enabled = sStatus
FlowSet13.Enabled = sStatus

FlSet_14.Enabled = sStatus
FlowSet14.Enabled = sStatus

FlSet_15.Enabled = sStatus
FlowSet15.Enabled = sStatus

FlSet_16.Enabled = sStatus
FlowSet16.Enabled = sStatus

FlSet_17.Enabled = sStatus
FlowSet17.Enabled = sStatus

FlSet_18.Enabled = sStatus
FlowSet18.Enabled = sStatus


OptMFC1(0).Enabled = sStatus
OptMFC1(1).Enabled = sStatus
OptMFC1(2).Enabled = sStatus

OptMFC2(0).Enabled = sStatus
OptMFC2(1).Enabled = sStatus
OptMFC2(2).Enabled = sStatus

OptMFC3(0).Enabled = sStatus
OptMFC3(1).Enabled = sStatus
OptMFC3(2).Enabled = sStatus

OptMFC4(0).Enabled = sStatus
OptMFC4(1).Enabled = sStatus
OptMFC4(2).Enabled = sStatus

OptMFC5(0).Enabled = sStatus
OptMFC5(1).Enabled = sStatus
OptMFC5(2).Enabled = sStatus

'must change name of remining Opt and add status state code

End Sub

Private Sub StartLog_Click()
    answer = MsgBox("This will overwrite data in the csv file.  Do you want to continue?", vbExclamation + vbYesNo, "Confirm")
    If answer = vbYes Then
    RecordCountLbl.Caption = "0"
    Timer1.Enabled = True
    StartLog.Enabled = False
    OpenLog.Enabled = False
    StopLog.Enabled = True
    StatusBar1.Panels.Item(5).Text = "Data Log: Logging "
        
    End If
End Sub

Private Sub StopLog_Click()
DataLogSeqTxt.Text = "4"
StartLog.Enabled = True
StopLog.Enabled = False
OpenLog.Enabled = True
Timer1.Enabled = False
StatusBar1.Panels.Item(5).Text = "Data Log: Disabled "
End Sub

Private Sub T_Delay500ms_Timer()
RFVTimer = True
T_Delay500ms = False
End Sub

Private Sub Timer1_Timer()
Dim RecordCount As Single
On Error Resume Next
Open App.Path + "\Data\" & "NewFile.csv" For Output As 1
If DataLogSeqTxt.Text = "0" Then
    Timer1.Interval = LogRateTxt.Text * 1000
    Print #1, LogDateTxt.Text
    DataLogSeqTxt.Text = "1"
ElseIf DataLogSeqTxt.Text = "1" Then
    Print #1, LogNameTxt.Text
    DataLogSeqTxt.Text = "3"
ElseIf DataLogSeqTxt.Text = "2" Then
    'Print #1, "Time" & "," & "LFC1" & "," & "LFC2" & "," & "LFC3" & "," & "LFC4" & "," _
    & "LFC5" & "," & "LFC6" & "," & "LFC7" & "," & "LFC8" & "," & "LFC9" & "," & "LFC10" & "," _
    & "LFC11" & "," & "LFC12" & "," & "LFC13" & "," & "LFC14" & "," & "LFC15" & "," & "LFC16" & "," _
    & "LFC17" & "," & "LFC18"
    'DataLogSeqTxt.Text = "3"
ElseIf DataLogSeqTxt.Text = "3" Then
    Timer1.Interval = LogRateTxt.Text * 1000
    Print #1, Format(Time, "hh:mm:ss") & "," & FlCV_1.Text & "," & FlCV_2.Text _
    & "," & FlCV_3.Text & "," & FlCV_4.Text & "," & FlCV_5.Text & "," & FlCV_6.Text _
    & "," & FlCV_7.Text & "," & FlCV_8.Text & "," & FlCV_9.Text & "," & FlCV_10.Text _
    & "," & FlCV_11.Text & "," & FlCV_12.Text & "," & FlCV_13.Text & "," & FlCV_14.Text _
    & "," & FlCV_15.Text & "," & FlCV_16.Text & "," & FlCV_17.Text & "," & FlCV_18.Text
    RecordCount = RecordCountLbl.Caption + 1
    RecordCountLbl.Caption = RecordCount
ElseIf DataLogSeqTxt.Text = "4" Then
    StopLog.Enabled = False
    OpenLog.Enabled = True
    Close 1
    Timer1.Enabled = False
    DataLogSeqTxt.Text = "0"
    StatusBar1.Panels.Item(5).Text = "Data Log: Disabled "
    End If
End Sub

Private Sub WarningTmr_Timer()
Dim WarnVal1 As Single
Dim AlarmVal1 As Single
Dim SetVal1 As Single
SetVal1 = FlSet_1.Text
WarnVal1 = WarnValTxt.Text * 0.01
AlarmVal1 = AlarmValTxt.Text * 0.01
'flow rate warnings
If SetVal1 > 0 Then
    If FlCV_1.Text < SetVal1 - (SetVal1 * WarnVal1) Or _
    FlCV_1.Text > SetVal1 + (SetVal1 * WarnVal1) And _
    Not ActFlw1 < SetVal1 - (SetVal1 * AlarmVal1) And _
    Not ActFlw1 > SetVal1 + (SetVal1 * AlarmVal1) Then
        Frame1.BackColor = &HFFFF&
        FlCV_1.BackColor = &HFFFF&
    ElseIf Not FlCV_1.Text < SetVal1 - (SetVal1 * WarnVal1) And _
    Not FlCV_1.Text > SetVal1 + (SetVal1 * AlarmVal1) Then
        Frame1.BackColor = &H8000000F
        FlCV_1.BackColor = &H8000000F
    End If
    Label29(3).Caption = SetVal1 + (SetVal1 * WarnVal1)
Else
    Frame1.BackColor = &H8000000F
    FlCV_1.BackColor = &H8000000F
End If
End Sub

Private Sub AlarmTmr_Timer()
Dim WarnVal1 As Single
Dim AlarmVal1 As Single
Dim ActFlw1 As Single
Dim SetVal1 As Single
WarnVal1 = WarnValTxt.Text * 0.01
AlarmVal1 = AlarmValTxt.Text * 0.01
ActFlw1 = FlCV_1.Text
SetVal1 = FlSet_1.Text
Label29(2).Caption = SetVal1 + (SetVal1 * AlarmVal1)
'flow rate alarms
If SetVal1 > 0 Then
    If ActFlw1 < SetVal1 - (SetVal1 * AlarmVal1) Or _
    ActFlw1 > SetVal1 + (SetVal1 * AlarmVal1) Then
        Frame1.BackColor = &HFF&
        FlCV_1.BackColor = &HFF&
    ElseIf Not ActFlw1 < SetVal1 - (SetVal1 * WarnVal1) Or _
    Not ActFlw1 > SetVal1 + (SetVal1 * AlarmVal1) Then
        Frame1.BackColor = &H8000000F
        FlCV_1.BackColor = &H8000000F
    End If
Else
    Frame1.BackColor = &H8000000F
    FlCV_1.BackColor = &H8000000F
End If
End Sub

