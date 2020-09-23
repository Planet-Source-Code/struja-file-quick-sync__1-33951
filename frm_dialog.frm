VERSION 5.00
Begin VB.Form frm_dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ..."
   ClientHeight    =   975
   ClientLeft      =   5220
   ClientTop       =   5055
   ClientWidth     =   4320
   ClipControls    =   0   'False
   Icon            =   "frm_dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   855
      Top             =   45
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   2565
      TabIndex        =   1
      Top             =   675
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   1
      Left            =   1890
      TabIndex        =   4
      Top             =   630
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "File Quick Sync"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   510
      Index           =   0
      Left            =   1260
      TabIndex        =   0
      Top             =   135
      Width           =   2985
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "File Quick Sync"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   510
      Index           =   1
      Left            =   1290
      TabIndex        =   3
      Top             =   170
      Width           =   2985
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "struja-online@vip.hr"
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   1905
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frm_dialog.frx":0442
      Top             =   135
      Width           =   480
   End
End
Attribute VB_Name = "frm_dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    CenterForm_global Me
    Timer1.Enabled = True
    
    Label2(0).Caption = "V " & App.Major & "." & App.Minor & "." & App.Revision
    Label2(1).Caption = "V " & App.Major & "." & App.Minor & "." & App.Revision
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    
End Sub

Private Sub Image1_Click()
    Unload Me
    
End Sub

Private Sub Label1_Click(Index As Integer)

Unload Me

End Sub

Private Sub Label2_Click(Index As Integer)

Unload Me

End Sub

Private Sub Label3_Click()
    Unload Me

End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
    Unload Me
    
End Sub
