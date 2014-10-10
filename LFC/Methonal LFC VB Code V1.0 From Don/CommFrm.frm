VERSION 5.00
Begin VB.Form CommFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "COMM Settings"
   ClientHeight    =   2160
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   1890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   1890
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE COMM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   6
      Top             =   1755
      Width           =   1845
   End
   Begin VB.Frame Frame3 
      Caption         =   "COM Setting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   1830
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   630
         TabIndex        =   3
         Text            =   "38400"
         Top             =   810
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "CommFrm.frx":0000
         Left            =   630
         List            =   "CommFrm.frx":0002
         TabIndex        =   2
         Text            =   "4"
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Baud Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "SET and CONNECT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   1365
      Width           =   1845
   End
End
Attribute VB_Name = "CommFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MainFrm.Command10.Value = True
    MainFrm.StatusBar1.Panels.Item(1).Text = "COM STATUS: Disconnected "
    CommFrm.Visible = False
    MainFrm.StatusBar1.Panels(3).Text = "Rx: 0"
    MainFrm.StatusBar1.Panels(4).Text = "Tx: 0"
    MainFrm.CommStatusTmr.Enabled = False
    MainFrm.Command8.Value = True
    MainFrm.Timer1.Enabled = False
    MainFrm.DataLogSeqTxt.Text = "0"
    MainFrm.StatusBar1.Panels.Item(5).Text = "Data Log: Disabled "
    MainFrm.StartLog.Enabled = False
    MainFrm.StopLog.Enabled = False
    MainFrm.OpenLog.Enabled = True
    MainFrm.DataLogSeqTxt.Text = "4"
    MainFrm.Timer1.Enabled = True
    MainFrm.EmoCmd.Enabled = False
End Sub

Private Sub Form_Load()
Combo1.AddItem "1"
Combo1.AddItem "2"
Combo1.AddItem "3"
Combo1.AddItem "4"
Combo1.AddItem "5"
Combo1.AddItem "6"
Combo1.AddItem "7"
Combo1.AddItem "8"
Combo1.AddItem "9"
Combo1.AddItem "10"
Combo1.AddItem "11"
Combo1.AddItem "12"
Combo1.AddItem "13"
Combo1.AddItem "14"
Combo1.AddItem "15"
Combo1.AddItem "16"
Combo1.ListIndex = 0
End Sub

Private Sub OKButton_Click()
    MainFrm.Combo1 = Combo1.Text
    MainFrm.Combo2 = Combo2.Text
    MainFrm.Command9.Value = True
    CommFrm.Visible = False
    MainFrm.StatusBar1.Panels.Item(1).Text = "COM STATUS: Connected "
    MainFrm.CommStatusTmr.Enabled = True
    MainFrm.StartLog.Enabled = True
    MainFrm.DataLogSeqTxt.Text = "0"
    MainFrm.EmoCmd.Enabled = True
End Sub

