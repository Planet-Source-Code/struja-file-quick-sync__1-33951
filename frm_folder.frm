VERSION 5.00
Begin VB.Form frm_folder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select folder"
   ClientHeight    =   4170
   ClientLeft      =   4845
   ClientTop       =   2865
   ClientWidth     =   3915
   Icon            =   "frm_folder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_ok 
      Caption         =   "OK"
      Height          =   330
      Left            =   720
      TabIndex        =   3
      Top             =   3780
      Width           =   1005
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2295
      TabIndex        =   2
      Top             =   3780
      Width           =   1005
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   3705
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   3375
      Width           =   3705
   End
End
Attribute VB_Name = "frm_folder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_cancel_Click()
    Unload Me
    
End Sub

Private Sub cmd_ok_Click()
    SelectedFolder = Dir1.path
    
    Unload Me
    
End Sub

Private Sub Drive1_Change()
    Dir1.path = Drive1.Drive
    
End Sub

Private Sub Form_Load()
    CenterForm_global Me
    
End Sub
