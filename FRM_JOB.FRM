VERSION 5.00
Begin VB.Form frm_action 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2880
   ClientLeft      =   4215
   ClientTop       =   4050
   ClientWidth     =   5235
   Icon            =   "frm_job.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk_includesubfolders 
      Caption         =   "Include subfolders"
      Height          =   195
      Left            =   1755
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   495
      Width           =   1680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1080
      Top             =   765
   End
   Begin VB.TextBox txt_description 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   45
      TabIndex        =   3
      Top             =   2070
      Width           =   5145
   End
   Begin VB.CheckBox chk_or 
      Caption         =   "Or"
      Height          =   195
      Left            =   4455
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2610
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CheckBox chk_and 
      Caption         =   "And"
      Height          =   195
      Left            =   4050
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2430
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton cmd_browseto 
      Height          =   285
      Left            =   4815
      Picture         =   "frm_job.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1485
      Width           =   375
   End
   Begin VB.CommandButton cmd_browsefrom 
      Height          =   285
      Left            =   4815
      Picture         =   "frm_job.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   900
      Width           =   375
   End
   Begin VB.CheckBox chk_sizebig 
      Caption         =   "Size >"
      Height          =   195
      Left            =   4365
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   495
      Width           =   870
   End
   Begin VB.CheckBox chk_sizedif 
      Caption         =   "Size <>"
      Height          =   195
      Left            =   4365
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   270
      Width           =   870
   End
   Begin VB.CheckBox chk_datebig 
      Caption         =   "Date >"
      Height          =   195
      Left            =   3465
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   495
      Width           =   870
   End
   Begin VB.CheckBox chk_datedif 
      Caption         =   "Date <>"
      Height          =   195
      Left            =   3465
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   270
      Width           =   870
   End
   Begin VB.CheckBox chk_enabled 
      Caption         =   "Enabled"
      Height          =   195
      Left            =   1755
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   225
      Width           =   960
   End
   Begin VB.TextBox txt_folderto 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Top             =   1485
      Width           =   4695
   End
   Begin VB.TextBox txt_folderfrom 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Top             =   900
      Width           =   4695
   End
   Begin VB.TextBox txt_filter 
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
      Left            =   45
      TabIndex        =   0
      Top             =   315
      Width           =   1635
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2970
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2475
      Width           =   1005
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   1395
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2475
      Width           =   1005
   End
   Begin VB.Label Label5 
      Caption         =   "Description:"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   18
      Top             =   1845
      Width           =   1050
   End
   Begin VB.Label Label5 
      Caption         =   "File filter:"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "From folder:"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   7
      Top             =   675
      Width           =   1140
   End
   Begin VB.Label Label5 
      Caption         =   "To folder:"
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   6
      Top             =   1260
      Width           =   1140
   End
End
Attribute VB_Name = "frm_action"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chk_and_Click()
    If chk_and = 1 Then chk_or = 0
    
End Sub

Private Sub chk_datebig_Click()
    If chk_datebig = 1 Then chk_datedif = 0
    
If chk_datebig = 0 And chk_datedif = 0 Then chk_datebig = 1

End Sub

Private Sub chk_datedif_Click()
    If chk_datedif = 1 Then chk_datebig = 0
    
If chk_datebig = 0 And chk_datedif = 0 Then chk_datedif = 1
    
End Sub

Private Sub chk_or_Click()
    If chk_or = 1 Then chk_and = 0
    
End Sub

Private Sub chk_sizebig_Click()
    If chk_sizebig = 1 Then chk_sizedif = 0
    
If chk_sizebig = 0 And chk_sizedif = 0 Then chk_sizebig = 1
    
End Sub

Private Sub chk_sizedif_Click()
    If chk_sizedif = 1 Then chk_sizebig = 0
    
If chk_sizebig = 0 And chk_sizedif = 0 Then chk_sizedif = 1
    
End Sub

Private Sub cmd_browsefrom_Click()

On Error GoTo ErrH

Dim sPath As String

            '   Open the Browse Folder dialog box and
            '   return the folder selected
    
    'StartFolder = txt_folderfrom
    'sPath = FolderBrowse(Me.hWnd, "Please select source folder", 64)
    sPath = SelectFolder(Me, "Please select source folder")
    
            '   If no folder was selected, exit here
    If Len(sPath) = 0 Then
        Exit Sub
    Else
        txt_folderfrom = sPath
    End If

Exit Sub
    
ErrH:
    MsgBox "Unhandled error: " & Err.Number & vbCr & Err.Description, vbOKOnly + vbCritical, "Error"

End Sub

Private Sub cmd_browseto_Click()

On Error GoTo ErrH

Dim sPath As String

            '   Open the Browse Folder dialog box and
            '   return the folder selected
    'StartFolder = txt_folderto
    'sPath = FolderBrowse(Me.hWnd, "Please select destination folder", 64)
    sPath = SelectFolder(Me, "Please select destination folder")
  
            '   If no folder was selected, exit here
    If Len(sPath) = 0 Then
        Exit Sub
    Else
        txt_folderto = sPath
    End If

Exit Sub
    
ErrH:
    MsgBox "Unhandled error: " & Err.Number & vbCr & Err.Description, vbOKOnly + vbCritical, "Error"

End Sub

Private Sub cmd_cancel_Click()
    Unload Me
    
End Sub

Private Sub cmd_ok_Click()

txt_filter = Trim(txt_filter)
txt_folderfrom = Trim(txt_folderfrom)
txt_folderto = Trim(txt_folderto)

If txt_filter = "" Then
    tmp = MsgBox("Please input file filter criteria!", vbOKOnly + vbCritical, "Error")
    txt_filter.SetFocus
    Exit Sub
End If

If txt_folderfrom = "" Then
    tmp = MsgBox("Please select folder to copy from!", vbOKOnly + vbCritical, "Error")
    txt_folderfrom.SetFocus
    Exit Sub
End If

If txt_folderto = "" Then
    tmp = MsgBox("Please select folder to copy to!", vbOKOnly + vbCritical, "Error")
    txt_folderto.SetFocus
    Exit Sub
End If

If chk_datedif = 0 And chk_datebig = 0 And chk_sizedif = 0 And chk_sizebig = 0 Then
    chk_and = 0
    chk_or = 0
End If

If chk_datedif = 0 And chk_datebig = 0 And chk_sizedif = 0 And chk_sizebig = 0 Then
    chk_and = 0
    chk_or = 0
Else
    If chk_and = 0 And chk_or = 0 Then
        chk_or = 1
        chk_and = 0
    End If
End If




Select Case Me.Caption
    Case "New action"
        NewFlag = True
        
        With frm_main.Data1.Recordset
        .AddNew
            
            !TemplateNo = TemplateNo
            
            If chk_enabled = 1 Then
                !Enabled = "YES"
            Else
                !Enabled = "NO"
            End If
            
            !FileFilter = txt_filter
            !FromFolder = txt_folderfrom
            !ToFolder = txt_folderto
            !Description = txt_description
            
            If chk_includesubfolders = 1 Then
                !InclSub = "YES"
            Else
                !InclSub = "NO"
            End If
            
            !FileDate = ""
            If chk_datebig = 1 Then !FileDate = ">"
            If chk_datedif = 1 Then !FileDate = "<>"
            
            !FileSize = ""
            If chk_sizebig = 1 Then !FileSize = ">"
            If chk_sizedif = 1 Then !FileSize = "<>"
            
            !Operator = ""
            If chk_and = 1 Then !Operator = "AND"
            If chk_or = 1 Then !Operator = "OR"
            
        .Update
        End With
            
        NewFlag = False
        frm_main.Data1.Refresh
        
    Case "Edit action"
        With frm_main.Data1.Recordset
        .Edit
            
            If chk_enabled = 1 Then
                !Enabled = "YES"
            Else
                !Enabled = "NO"
            End If
            
            !FileFilter = txt_filter
            !FromFolder = txt_folderfrom
            !ToFolder = txt_folderto
            !Description = txt_description
            
            If chk_includesubfolders = 1 Then
                !InclSub = "YES"
            Else
                !InclSub = "NO"
            End If
            
            !FileDate = ""
            If chk_datebig = 1 Then !FileDate = ">"
            If chk_datedif = 1 Then !FileDate = "<>"
            
            !FileSize = ""
            If chk_sizebig = 1 Then !FileSize = ">"
            If chk_sizedif = 1 Then !FileSize = "<>"
            
            !Operator = ""
            If chk_and = 1 Then !Operator = "AND"
            If chk_or = 1 Then !Operator = "OR"
            
        .Update
        End With
    
    
    
    
End Select

Unload Me

End Sub

Private Sub Form_Load()
    CenterForm_global Me
    

    Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False

Select Case Me.Caption
    Case "New action"
        chk_enabled = 1
        chk_datebig = 1
        chk_sizedif = 1
        chk_or = 1
        
        txt_filter = "*.*"
        txt_filter.SetFocus
    
    
    Case "Edit action"
        With frm_main.Data1.Recordset
            
            If UCase("" & !Enabled) = "YES" Then
                chk_enabled = 1
            Else
                chk_enabled = 0
            End If
            
            txt_filter = "" & !FileFilter
            txt_folderfrom = "" & !FromFolder
            txt_folderto = "" & !ToFolder
            txt_description = "" & !Description
            
            If UCase("" & !InclSub) = "YES" Then
                chk_includesubfolders = 1
            Else
                chk_includesubfolders = 0
            End If
            
            If "" & !FileDate = ">" Then chk_datebig = 1
            If "" & !FileDate = "<>" Then chk_datedif = 1
            
            If "" & !FileSize = ">" Then chk_sizebig = 1
            If "" & !FileSize = "<>" Then chk_sizedif = 1
            
            If "" & !Operator = "AND" Then chk_and = 1
            If "" & !Operator = "OR" Then chk_or = 1
    
        End With
    
End Select

End Sub
