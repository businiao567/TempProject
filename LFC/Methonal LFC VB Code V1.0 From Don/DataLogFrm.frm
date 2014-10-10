VERSION 5.00
Begin VB.Form DataLogFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Logging Setup"
   ClientHeight    =   2460
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Log Entry Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   3135
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
         Left            =   1980
         TabIndex        =   6
         Text            =   "1"
         Top             =   1515
         Width           =   405
      End
      Begin VB.TextBox LogDateTxt 
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
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   615
         TabIndex        =   3
         Text            =   "Enter Log Date"
         Top             =   1080
         Width           =   1770
      End
      Begin VB.TextBox LogNameTxt 
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
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   60
         TabIndex        =   2
         Text            =   "Enter Log Name"
         Top             =   480
         Width           =   3000
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Data Log Rate (sec):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   7
         Top             =   1530
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Data Log Test Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   255
         Width           =   3000
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Log Test Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   630
         TabIndex        =   4
         Top             =   855
         Width           =   1695
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Save and CLOSE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   2025
      Width           =   3135
   End
End
Attribute VB_Name = "DataLogFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
MainFrm.LogNameTxt.Text = LogNameTxt.Text
MainFrm.LogDateTxt.Text = LogDateTxt.Text
MainFrm.LogRateTxt.Text = LogRateTxt.Text
DataLogFrm.Visible = False
End Sub
