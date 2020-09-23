VERSION 5.00
Begin VB.Form frm_templates 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1710
   ClientLeft      =   4230
   ClientTop       =   5160
   ClientWidth     =   5265
   Icon            =   "frm_templates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_name 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Top             =   315
      Width           =   2895
   End
   Begin VB.TextBox txt_description 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Top             =   900
      Width           =   5145
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   1350
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1305
      Width           =   1005
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2880
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1305
      Width           =   1005
   End
   Begin VB.Label Label5 
      Caption         =   "Name:"
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   1140
   End
   Begin VB.Label Label5 
      Caption         =   "Description:"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   675
      Width           =   1050
   End
End
Attribute VB_Name = "frm_templates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_cancel_Click()
    Unload Me
    
End Sub

Private Sub cmd_ok_Click()

txt_name = Trim(txt_name)

If txt_name = "" Then
    tmp = MsgBox("Please input template name!", vbOKOnly + vbCritical, "Error")
    txt_name.SetFocus
    Exit Sub
End If
    
Dim rs As Recordset
    
Select Case Me.Caption
    Case "New template"
        Set rs = DbSys.OpenRecordset("Select Templates.* FROM Templates;")
        
        rs.AddNew
            rs!Name = txt_name
            rs!Description = txt_description
        rs.Update
    
    Case "Edit template"
        Set rs = DbSys.OpenRecordset("Select Templates.* FROM Templates WHERE (((Templates.TemplateNo)=" & TemplateNo & "));")
        
        rs.Edit
            rs!Name = txt_name
            rs!Description = txt_description
        rs.Update
    
End Select
    
frm_main.LoadTemplates
frm_main.cmb_template = txt_name

Unload Me

End Sub

Private Sub Form_Activate()

Select Case Me.Caption
    Case "New template"
        txt_name.SetFocus
    
    Case "Edit template"
        txt_name = frm_main.cmb_template
        txt_description = frm_main.cmb_template.ToolTipText
    
End Select


End Sub

Private Sub Form_Load()
    CenterForm_global Me
    
    
End Sub
