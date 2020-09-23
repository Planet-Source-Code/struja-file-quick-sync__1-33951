VERSION 5.00
Begin VB.Form frm_question 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sinchronize"
   ClientHeight    =   1290
   ClientLeft      =   4530
   ClientTop       =   5175
   ClientWidth     =   5925
   Icon            =   "frm_question.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk_showallinreport 
      Caption         =   "Include (no criteria matched) files in report"
      Height          =   240
      Left            =   810
      TabIndex        =   0
      Top             =   540
      Width           =   3615
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   1710
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   900
      Width           =   1005
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3285
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   900
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Are you shure you want to synchronize files with defined actions?"
      Height          =   285
      Left            =   810
      TabIndex        =   3
      Top             =   180
      Width           =   4920
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "frm_question.frx":0442
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "frm_question"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_cancel_Click()
    Unload Me
    
End Sub

Private Sub cmd_ok_Click()

If chk_showallinreport = 1 Then
    FlagShowAllInReport = True
Else
    FlagShowAllInReport = False
End If

frm_sync.Show

Unload Me

End Sub

Private Sub Form_Load()

CenterForm_global Me

Beep

chk_showallinreport = 0

End Sub
