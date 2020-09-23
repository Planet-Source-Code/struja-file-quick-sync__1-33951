VERSION 5.00
Begin VB.Form frm_scheduler 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scheduler"
   ClientHeight    =   1770
   ClientLeft      =   5400
   ClientTop       =   4920
   ClientWidth     =   3810
   Icon            =   "frm_scheduler.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1575
      Top             =   225
   End
   Begin VB.CheckBox chk_enabled 
      Caption         =   "Enabled"
      Height          =   195
      Left            =   2745
      TabIndex        =   4
      Top             =   1035
      Width           =   915
   End
   Begin VB.TextBox txt_time 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1665
      TabIndex        =   2
      Top             =   945
      Width           =   960
   End
   Begin VB.TextBox txt_date 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   945
      Width           =   1140
   End
   Begin VB.CheckBox chk_daily 
      Caption         =   "Daily"
      Height          =   195
      Left            =   2745
      TabIndex        =   3
      Top             =   765
      Width           =   825
   End
   Begin VB.ComboBox cmb_template 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   315
      Width           =   3615
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   585
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1350
      Width           =   1005
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2205
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1350
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "at"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1305
      TabIndex        =   10
      Top             =   990
      Width           =   285
   End
   Begin VB.Label Label1 
      Caption         =   "Time:"
      Height          =   195
      Index           =   2
      Left            =   1710
      TabIndex        =   9
      Top             =   720
      Width           =   870
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   8
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "Template:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   7
      Top             =   90
      Width           =   1050
   End
End
Attribute VB_Name = "frm_scheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chk_daily_Click()

If chk_daily = 1 Then
    txt_date.Enabled = False
Else
    txt_date.Enabled = True
End If

End Sub

Private Sub cmd_cancel_Click()
    Unload Me
    
End Sub

Private Sub cmd_ok_Click()

If cmb_template = "- none -" Then
    tmp = MsgBox("Please select template!", vbOKOnly + vbCritical, "Error")
    Exit Sub
End If

txt_date = Trim(txt_date)
txt_time = Trim(txt_time)

If txt_time = "" Then
    tmp = MsgBox("Please enter time for scheduler", vbOKOnly + vbCritical, "Error")
    txt_time.SetFocus
    Exit Sub
End If

If txt_date = "" And chk_daily = 0 Then
    tmp = MsgBox("Please enter date for scheduler", vbOKOnly + vbCritical, "Error")
    txt_date.SetFocus
    Exit Sub
End If


' upisi u ini file
sIniFile = GlobalAppPath & "fqs-timer.tmr"

IniPisi "Timer", "Template", cmb_template

IniPisi "Timer", "Date", txt_date
IniPisi "Timer", "Time", txt_time

If chk_daily = 1 Then
    IniPisi "Timer", "Daily", "YES"
Else
    IniPisi "Timer", "Daily", "NO"
End If

If chk_enabled = 1 Then
    IniPisi "Timer", "Enabled", "YES"
Else
    IniPisi "Timer", "Enabled", "NO"
End If


frm_main.LoadTimerSettings

Unload Me
End Sub

Private Sub Form_Load()
    CenterForm_global Me
    
For qwe = 0 To frm_main.cmb_template.ListCount - 1

    cmb_template.AddItem frm_main.cmb_template.List(qwe)
    
Next qwe

Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

Timer1.Enabled = False

If IsTemplateExists(tmrTemplate) = True Then
    cmb_template = tmrTemplate
Else
    cmb_template = "- none -"
End If

txt_date = tmrDate
txt_time = tmrTime

If tmrDaily = True Then
    chk_daily = 1
Else
    chk_daily = 0
End If

If tmrEnabled = True Then
    chk_enabled = 1
Else
    chk_enabled = 0
End If

End Sub

Private Sub txt_date_DblClick()

    txt_date = Format(Date, "dd.mm.yyyy")
    
End Sub

Private Sub txt_date_LostFocus()

If txt_date <> "" Then
    txt_date = NapraviDatumStruja(txt_date)
End If

End Sub

Private Sub txt_time_LostFocus()

If txt_time <> "" Then
    txt_time = NapraviVrijemeStruja(txt_time)
End If

End Sub
