VERSION 5.00
Begin VB.Form AlarmSetUpFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alarm Setup"
   ClientHeight    =   2490
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Flow Deviation Alarm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   30
      TabIndex        =   2
      Top             =   1035
      Width           =   2700
      Begin VB.TextBox AlarmTimeDelayTxt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2205
         TabIndex        =   10
         Text            =   "2"
         Top             =   585
         Width           =   420
      End
      Begin VB.TextBox AlarmTxt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2205
         TabIndex        =   8
         Text            =   "5"
         Top             =   270
         Width           =   420
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Delay (sec):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   630
         TabIndex        =   9
         Top             =   600
         Width           =   1530
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deviation Setpoint (%):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   285
         Width           =   1980
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Flow Deviation Warning"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   2700
      Begin VB.TextBox WarnTimeDelayTxt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2175
         TabIndex        =   6
         Text            =   "2"
         Top             =   585
         Width           =   420
      End
      Begin VB.TextBox WarnTxt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2175
         TabIndex        =   4
         Text            =   "5"
         Top             =   270
         Width           =   420
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Delay (sec):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   1530
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deviation Setpoint (%):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   285
         Width           =   1980
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Save and CLOSE"
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
      Left            =   45
      TabIndex        =   0
      Top             =   2085
      Width           =   2700
   End
End
Attribute VB_Name = "AlarmSetUpFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()

If WarnTimeDelayTxt.Text = "" Then
    MainFrm.WarningTmr.Interval = 1000
Else
    MainFrm.WarningTmr.Interval = WarnTimeDelayTxt.Text * 1000
End If

If WarnTxt.Text = "" Then
    MainFrm.WarnValTxt.Text = "1"
Else
    MainFrm.WarnValTxt.Text = WarnTxt.Text
End If

If AlarmTimeDelayTxt.Text = "" Then
    MainFrm.AlarmTmr.Interval = 1000
Else
    MainFrm.AlarmTmr.Interval = AlarmTimeDelayTxt.Text * 1000
End If

If AlarmTxt.Text = "" Then
    MainFrm.AlarmValTxt.Text = "1"
Else
    MainFrm.AlarmValTxt.Text = AlarmTxt.Text
End If
AlarmSetUpFrm.Visible = False
End Sub
