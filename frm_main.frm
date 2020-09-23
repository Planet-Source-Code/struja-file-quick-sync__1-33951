VERSION 5.00
Object = "{00028C4A-0000-0000-0000-000000000046}#5.0#0"; "TDBG5.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Quick Sync - Copyright(c) By 2001 by Struja"
   ClientHeight    =   5370
   ClientLeft      =   3840
   ClientTop       =   3480
   ClientWidth     =   6720
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6720
   Begin VB.Frame Frame2 
      Caption         =   "Actions"
      Height          =   4425
      Left            =   45
      TabIndex        =   3
      Top             =   855
      Width           =   6630
      Begin VB.CheckBox chk_includesubfolders 
         Caption         =   "Include subfolders"
         Height          =   420
         Left            =   5445
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1140
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Actions"
         Top             =   990
         Visible         =   0   'False
         Width           =   3480
      End
      Begin TrueDBGrid50.TDBGrid TDBGrid1 
         Bindings        =   "frm_main.frx":08CA
         Height          =   1815
         Left            =   135
         OleObjectBlob   =   "frm_main.frx":08DE
         TabIndex        =   1
         Top             =   270
         Width           =   6360
      End
      Begin VB.Label Label5 
         Caption         =   "File date:"
         Height          =   195
         Index           =   7
         Left            =   4680
         TabIndex        =   19
         Top             =   2160
         Width           =   690
      End
      Begin VB.Label Label5 
         Caption         =   "Operator:"
         Height          =   195
         Index           =   6
         Left            =   3915
         TabIndex        =   18
         Top             =   2160
         Width           =   690
      End
      Begin VB.Label Label5 
         Caption         =   "File size:"
         Height          =   195
         Index           =   5
         Left            =   3150
         TabIndex        =   17
         Top             =   2160
         Width           =   645
      End
      Begin VB.Label lbl_filedate 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4635
         TabIndex        =   16
         Top             =   2385
         Width           =   735
      End
      Begin VB.Label lbl_operator 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3870
         TabIndex        =   15
         Top             =   2385
         Width           =   735
      End
      Begin VB.Label lbl_filesize 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3105
         TabIndex        =   14
         Top             =   2385
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Description:"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   13
         Top             =   3780
         Width           =   1140
      End
      Begin VB.Label lbl_description 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   135
         TabIndex        =   12
         Top             =   4005
         Width           =   6360
      End
      Begin VB.Label Label5 
         Caption         =   "To folder:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   11
         Top             =   3240
         Width           =   1140
      End
      Begin VB.Label Label5 
         Caption         =   "From folder:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   10
         Top             =   2700
         Width           =   1140
      End
      Begin VB.Label Label5 
         Caption         =   "File filter:"
         Height          =   195
         Index           =   1
         Left            =   1170
         TabIndex        =   9
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Enabled:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lbl_tofolder 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   135
         TabIndex        =   7
         Top             =   3465
         Width           =   6360
      End
      Begin VB.Label lbl_fromfolder 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   135
         TabIndex        =   6
         Top             =   2925
         Width           =   6360
      End
      Begin VB.Label lbl_filter 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1125
         TabIndex        =   5
         Top             =   2385
         Width           =   1860
      End
      Begin VB.Label lbl_enabled 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   135
         TabIndex        =   4
         Top             =   2385
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Templates"
      Height          =   735
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   6630
      Begin VB.Timer tmr_sch 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5535
         Top             =   180
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1890
         Top             =   135
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer Timer1 
         Interval        =   250
         Left            =   2790
         Top             =   225
      End
      Begin VB.CommandButton cmd_sync 
         Caption         =   "&Synchronize! >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3915
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   225
         Width           =   2130
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
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   3615
      End
      Begin VB.Image Img_Red 
         Height          =   480
         Left            =   6210
         Picture         =   "frm_main.frx":3ECD
         ToolTipText     =   "Scheduler - disabled"
         Top             =   315
         Width           =   480
      End
      Begin VB.Image Img_Green 
         Height          =   480
         Left            =   6210
         Picture         =   "frm_main.frx":41D7
         ToolTipText     =   "Scheduler - enabled"
         Top             =   315
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Menu cap_file 
      Caption         =   "File"
      Begin VB.Menu mnu_import 
         Caption         =   "Import template"
      End
      Begin VB.Menu mnu_export 
         Caption         =   "Export template"
      End
      Begin VB.Menu x 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_viewreport 
         Caption         =   "View REPORT.LOG"
      End
      Begin VB.Menu x3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu cap_sync 
      Caption         =   "Synchronize"
      Begin VB.Menu mnu_sync 
         Caption         =   "Synchronize! >>"
      End
      Begin VB.Menu x2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_simulate 
         Caption         =   "Simulate sync"
      End
      Begin VB.Menu mnu_scheduler 
         Caption         =   "Scheduler sync"
      End
   End
   Begin VB.Menu cap_template 
      Caption         =   "Templates"
      Begin VB.Menu mnu_newtemplate 
         Caption         =   "New template"
      End
      Begin VB.Menu mnu_edittemplate 
         Caption         =   "Edit template"
      End
      Begin VB.Menu mnu_deletetemplate 
         Caption         =   "Delete template"
      End
      Begin VB.Menu x1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exporttoini 
         Caption         =   "Export to INI"
      End
   End
   Begin VB.Menu cap_action 
      Caption         =   "Actions"
      Begin VB.Menu mnu_newaction 
         Caption         =   "New action"
      End
      Begin VB.Menu mnu_editaction 
         Caption         =   "Edit action"
      End
      Begin VB.Menu mnu_deleteaction 
         Caption         =   "Delete action"
      End
   End
   Begin VB.Menu cap_help 
      Caption         =   "Help"
      Begin VB.Menu mnu_about 
         Caption         =   "About ..."
      End
   End
   Begin VB.Menu cap_sysmenu 
      Caption         =   "SysMenu"
      Visible         =   0   'False
      Begin VB.Menu mnu_restoreFQS 
         Caption         =   "Restore FQS"
      End
      Begin VB.Menu x4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exitFQS 
         Caption         =   "Exit FQS"
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chk_includesubfolders_Click()

    If NewFlag = True Then Exit Sub

Dim xx As String

If Data1.Recordset.RecordCount = 0 Then
    chk_includesubfolders = 0
    Exit Sub
End If

xx = UCase("" & Data1.Recordset!InclSub)

            If xx = "YES" Then
                chk_includesubfolders = 1
            Else
                chk_includesubfolders = 0
            End If

End Sub

Private Sub cmb_template_Click()

PromjeniViewReportMeni

If cmb_template = "- none -" Then
    TemplateNo = 0
    cmb_template.ToolTipText = ""
    
    Data1.RecordSource = "Blank"
    Data1.Refresh
    
    Exit Sub
End If


Dim rs As Recordset

Set rs = DbSys.OpenRecordset("SELECT Templates.* FROM Templates WHERE (((Templates.Name)='" & cmb_template & "'));")

If rs.RecordCount = 0 Then
    cmb_template = "- none -"
    TemplateNo = 0
    cmb_template.ToolTipText = ""
    
    Data1.RecordSource = "Blank"
    Data1.Refresh
    
    Exit Sub
End If

TemplateNo = rs!TemplateNo
cmb_template.ToolTipText = "" & rs!Description

Data1.RecordSource = "SELECT Actions.* FROM Actions WHERE (((Actions.TemplateNo)=" & TemplateNo & "));"
Data1.Refresh

End Sub



Private Sub cmd_sync_Click()

mnu_sync_Click

End Sub

Private Sub Data1_Reposition()
    ClearFields
    
    If Data1.Recordset.RecordCount = 0 Then Exit Sub
    
    If NewFlag = True Then Exit Sub
    
With Data1.Recordset
    
    lbl_description = "" & !Description
    lbl_enabled = "" & !Enabled
    lbl_filedate = "" & !FileDate
    lbl_filesize = "" & !FileSize
    lbl_filter = "" & !FileFilter
    lbl_fromfolder = "" & !FromFolder
    lbl_operator = "" & !Operator
    lbl_tofolder = "" & !ToFolder
    
            If UCase("" & !InclSub) = "YES" Then
                chk_includesubfolders = 1
            Else
                chk_includesubfolders = 0
            End If
    
End With

End Sub

Private Sub Form_Load()
    
On Error GoTo greska

    CenterForm_global Me

If Right$(App.path, 1) = "\" Then
    GlobalAppPath = App.path
Else
    GlobalAppPath = App.path & "\"
End If

Set DbSys = OpenDatabase(GlobalAppPath & "File Quick Sync.mdb")

Data1.DatabaseName = DbSys.Name
Data1.RecordSource = "SELECT Actions.* FROM Actions ORDER BY Actions.FileFilter;"
Data1.Refresh

LoadTemplates

LoadTimerSettings

tmrDailyStarted = ""

Exit Sub

greska:

MsgBox "Error: " & Err.Number & vbCr & "Description: " & Error, vbOKOnly + vbCritical, "Error"

End

End Sub

Sub LoadTemplates()
cmb_template.Clear
cmb_template.AddItem "- none -"
TemplateNo = 0

Dim rs As Recordset

Set rs = DbSys.OpenRecordset("SELECT Templates.* FROM Templates ORDER BY Templates.Name;")

If rs.RecordCount = 0 Then
    cmb_template = "- none -"
    Exit Sub
End If

rs.MoveLast
rs.MoveFirst

For qwe = 1 To rs.RecordCount
    cmb_template.AddItem "" & rs!Name
    rs.MoveNext
Next qwe


    cmb_template = "- none -"

End Sub

Sub ClearFields()

lbl_description = ""
lbl_enabled = ""
lbl_filedate = ""
lbl_filesize = ""
lbl_filter = ""
lbl_fromfolder = ""
lbl_operator = ""
lbl_tofolder = ""
chk_includesubfolders = 0

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Dim Sys As Long
    Sys = x / Screen.TwipsPerPixelX

Debug.Print Sys

Select Case Sys
    Case WM_RBUTTONUP
        'Me.PopupMenu cap_sysmenu
        frm_main.WindowState = vbNormal
        Me.Show

    Case WM_LBUTTONUP
        'Me.PopupMenu cap_sysmenu
        frm_main.WindowState = vbNormal
        Me.Show

End Select

End Sub

Private Sub Form_Resize()

If WindowState = vbMinimized Then
    Me.Hide
    Me.Refresh

    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = Me.Caption & vbNullChar
    End With

    Shell_NotifyIcon NIM_ADD, nid
Else
    Shell_NotifyIcon NIM_DELETE, nid
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

Shell_NotifyIcon NIM_DELETE, nid

DbSys.Close

Unload frm_action
Unload frm_dialog
Unload frm_folder
Unload frm_scheduler
Unload frm_sync
Unload frm_templates
Unload frm_main
Unload frm_question

End

End Sub

Private Sub Img_Green_DblClick()
    mnu_scheduler_Click
    
End Sub

Private Sub Img_Red_DblClick()
    mnu_scheduler_Click
    
End Sub

Private Sub mnu_about_Click()
    frm_dialog.Show vbModal
    
End Sub

Private Sub mnu_deleteaction_Click()

    If Data1.Recordset.RecordCount = 0 Then Exit Sub

tmp = MsgBox("Are you shure you want to delete selected action?", vbYesNo + vbQuestion, "Delete")

If tmp = vbNo Then Exit Sub

Data1.Recordset.Delete
Data1.Refresh

End Sub

Private Sub mnu_deletetemplate_Click()

    If cmb_template = "- none -" Then Exit Sub
    
    tmp = MsgBox("Are you shure you want to delete selected template?", vbYesNo + vbQuestion, "Delete")
    
    If tmp = vbNo Then Exit Sub
    
    Dim rs As Recordset
    
    ' obrisi akcije
    Set rs = DbSys.OpenRecordset("SELECT Actions.* FROM Actions WHERE (((Actions.TemplateNo)=" & TemplateNo & "));")
    If rs.RecordCount > 0 Then
        NewFlag = True
        
        rs.MoveLast
        rs.MoveFirst
        
        For qwe = 1 To rs.RecordCount
            rs.Delete
            rs.MoveNext
        Next qwe
        
        NewFlag = False
    End If
    
    ' obrisi template
    Set rs = DbSys.OpenRecordset("SELECT Templates.* FROM Templates WHERE (((Templates.TemplateNo)=" & TemplateNo & "));")
    rs.Delete
    
LoadTemplates

End Sub

Private Sub mnu_editaction_Click()

    If Data1.Recordset.RecordCount = 0 Then Exit Sub
    
    frm_action.Caption = "Edit action"
    frm_action.Show vbModal

End Sub

Private Sub mnu_edittemplate_Click()

If cmb_template = "- none -" Then Exit Sub

    frm_templates.Caption = "Edit template"
    frm_templates.Show vbModal

End Sub

Private Sub mnu_exit_Click()
    Unload Me
    
End Sub

Private Sub mnu_exitFQS_Click()
    Unload frm_main
    
End Sub

Private Sub mnu_export_Click()
On Error GoTo greska

CommonDialog1.FileName = cmb_template & ".fqs"
CommonDialog1.Filter = "*.fqs (File Quick Sync template)"
CommonDialog1.DefaultExt = "fqs"
CommonDialog1.CancelError = True

CommonDialog1.ShowSave

Debug.Print CommonDialog1.FileTitle

If CommonDialog1.FileTitle = "" Or LCase(Right$(CommonDialog1.FileTitle, 4)) <> ".fqs" Then
    tmp = MsgBox("Error in template filename!", vbOKOnly + vbCritical, "Error")
    Exit Sub
Else
    
    sIniFile = CommonDialog1.FileName
    
    With Data1.Recordset
    .MoveFirst
    NewFlag = True
    
    IniPisi "Main", "TemplateName", cmb_template  ' template name
    IniPisi "Main", "Description", cmb_template.ToolTipText ' template description
    IniPisi "Main", "Actions", .RecordCount ' broj akcija
    
    For qwe = 1 To .RecordCount
        
        IniPisi "Action" & qwe, "Enabled", "" & !Enabled   ' enabled
        IniPisi "Action" & qwe, "FileFilter", "" & !FileFilter    ' filefilter
        IniPisi "Action" & qwe, "FromFolder", "" & !FromFolder ' fromfolder
        IniPisi "Action" & qwe, "ToFolder", "" & !ToFolder   ' tofolder
        IniPisi "Action" & qwe, "Description", "" & !Description 'description
        IniPisi "Action" & qwe, "FileDate", "" & !FileDate      ' filedate
        IniPisi "Action" & qwe, "FileSize", "" & !FileSize      ' filesize
        IniPisi "Action" & qwe, "Operator", "" & !Operator      ' operator
        IniPisi "Action" & qwe, "InclSub", "" & !InclSub       'include subfolders
    
        .MoveNext
    Next qwe
    
    NewFlag = False
    .MoveFirst
    End With
    
    tmp = MsgBox("Template successfull exported as:" & vbCr & CommonDialog1.FileName & vbCr & vbCr & "Name: " & cmb_template & ", actions#: " & Data1.Recordset.RecordCount, vbOKOnly + vbInformation, "Info")
    
End If

Exit Sub
    

greska:

End Sub

Private Sub mnu_exporttoini_Click()

    
    If cmb_template = "- none -" Then Exit Sub

tmp = MsgBox("Are you shure you want to create INI file for current template: " & cmb_template & " ?", vbYesNo + vbQuestion, "Create INI file")

If tmp = vbNo Then Exit Sub

sIniFile = GlobalAppPath & "FQS_RunMe.ini"

tmp = MsgBox("Do you want to enable user to view FQS status report before exit?", vbYesNo + vbQuestion, "Status report")

If tmp = vbYes Then
    IniPisi "AutoRunTemplate", "Name", cmb_template
    IniPisi "AutoRunTemplate", "ShowReport", "Yes"
Else
    IniPisi "AutoRunTemplate", "Name", cmb_template
    IniPisi "AutoRunTemplate", "ShowReport", "No"
End If

tmp = MsgBox("FQS ini file created as: " & GlobalAppPath & "FQS_RunMe.ini" & vbCr & "Template: " & cmb_template & vbCr & vbCr & "Note: Please rename FQS_RunMe.exe and FQS_RunMe.ini to same name", vbOKOnly + vbInformation, "INI file created")

End Sub

Private Sub mnu_import_Click()
On Error GoTo greska

CommonDialog1.FileName = "*.fqs"
CommonDialog1.Filter = "*.fqs (File Quick Sync template)"
CommonDialog1.CancelError = True

CommonDialog1.ShowOpen

Debug.Print CommonDialog1.FileTitle

If CommonDialog1.FileTitle = "" Or LCase(Right$(CommonDialog1.FileTitle, 4)) <> ".fqs" Then
    tmp = MsgBox("Error in template filename!", vbOKOnly + vbCritical, "Error")
    Exit Sub
Else
    
    If FileExist(CommonDialog1.FileName) Then
        
        sIniFile = CommonDialog1.FileName
    
        Dim tmpTemplName As String
        
        tmpTemplName = IniDaj("Main", "TemplateName")

        'check if template exists
        If IsTemplateExists(tmpTemplName) Then
            tmp = MsgBox("Template with name: " & tmpTemplName & " allready exists!" & vbCr & vbCr & "Overwrite existing template and actions?", vbYesNo + vbExclamation, "Overwrite template")
            
            If tmp = vbNo Then
                Exit Sub
            Else
                cmb_template = tmpTemplName
                mnu_deletetemplate_Click
                LoadTemplates
            
                ' if No is clicked - quit
                If IsTemplateExists(tmpTemplName) Then Exit Sub
                
            End If
        End If
        
        
        ' make new template
        Dim trs As Recordset
        Set trs = DbSys.OpenRecordset("Templates")
        
        trs.AddNew
            trs!Name = tmpTemplName
            trs!Description = IniDaj("Main", "Description")
        trs.Update

        ' load template
        LoadTemplates
        cmb_template = tmpTemplName
        
        ' load action count
        Dim tcnt As Long
        tcnt = Val(IniDaj("Main", "Actions", "0"))
        
        With Data1.Recordset
        NewFlag = True
        
        For qwe = 1 To tcnt
            .AddNew
            
            !TemplateNo = TemplateNo
            
            !Enabled = IniDaj("Action" & qwe, "Enabled")
            !FileFilter = IniDaj("Action" & qwe, "FileFilter")
            !FromFolder = IniDaj("Action" & qwe, "FromFolder")
            !ToFolder = IniDaj("Action" & qwe, "ToFolder")
            !Description = IniDaj("Action" & qwe, "Description")
            !FileDate = IniDaj("Action" & qwe, "FileDate")
            !FileSize = IniDaj("Action" & qwe, "FileSize")
            !Operator = IniDaj("Action" & qwe, "Operator")
            !InclSub = IniDaj("Action" & qwe, "InclSub")
    
            .Update
        Next qwe
    
            NewFlag = False
            .MoveFirst
        End With
    
        tmp = MsgBox("Template successfull imported from:" & vbCr & CommonDialog1.FileName & vbCr & vbCr & "Name: " & cmb_template & ", actions#: " & Data1.Recordset.RecordCount, vbOKOnly + vbInformation, "Info")
    
    Else
    
        tmp = MsgBox("Cannot open file: " & CommonDialog1.FileName, vbOKOnly + vbCritical, "Error")
        Exit Sub
        
    End If
    
End If

Exit Sub
    

greska:

End Sub

Private Sub mnu_newaction_Click()

If TemplateNo = 0 Then
    tmp = MsgBox("Please select template!", vbOKOnly + vbCritical, "Error")
    cmb_template.SetFocus
    Exit Sub
End If


    frm_action.Caption = "New action"
    frm_action.Show vbModal

End Sub

Private Sub mnu_newtemplate_Click()

    frm_templates.Caption = "New template"
    frm_templates.Show vbModal

End Sub

Private Sub mnu_restoreFQS_Click()

frm_main.WindowState = vbNormal
Me.Show

End Sub

Private Sub mnu_scheduler_Click()
tmr_sch.Enabled = False

    frm_scheduler.Show vbModal
    
tmr_sch.Enabled = True
End Sub

Private Sub mnu_simulate_Click()

    SyncMode = "Simulate"
    
    If cmb_template = "- none -" Then Exit Sub

If Data1.Recordset.RecordCount = 0 Then
    tmp = MsgBox("No actions defined! Please create some actions and then try again", vbOKOnly + vbCritical, "Error")
    Exit Sub
End If

tmp = MsgBox("Are you shure you want to synchronize files with defined actions?", vbYesNo + vbQuestion, SyncMode)

If tmp = vbNo Then Exit Sub

frm_sync.Show vbModal

End Sub

Private Sub mnu_sync_Click()
    SyncMode = "Synchronize"
    
    If cmb_template = "- none -" Then Exit Sub

If Data1.Recordset.RecordCount = 0 Then
    tmp = MsgBox("No actions defined! Please create some actions and then try again", vbOKOnly + vbCritical, "Error")
    Exit Sub
End If

frm_question.Show vbModal

End Sub

Private Sub mnu_viewreport_Click()
On Error Resume Next
Dim tmpReportFilename As String

tmpReportFilename = cmb_template & ".log"

If FileExist(GlobalAppPath & tmpReportFilename) Then
    Shell "notepad.exe" & " " & GlobalAppPath & tmpReportFilename, vbNormalFocus
Else
    tmp = MsgBox("Cannot find file:" & vbCr & GlobalAppPath & tmpReportFilename, vbOKOnly + vbCritical, "Error")
End If

End Sub

Private Sub TDBGrid1_DblClick()
    mnu_editaction_Click
    
End Sub

Private Sub TDBGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbRightButton Then
    
    Me.PopupMenu Me.cap_action

End If

End Sub

Private Sub Timer1_Timer()

If NewFlag = True Then Exit Sub


If cmb_template = "- none -" Then
    mnu_edittemplate.Enabled = False
    mnu_deletetemplate.Enabled = False
    mnu_newaction.Enabled = False
Else
    mnu_edittemplate.Enabled = True
    mnu_deletetemplate.Enabled = True
    mnu_newaction.Enabled = True
End If


If Data1.Recordset.RecordCount > 0 Then
    mnu_editaction.Enabled = True
    mnu_deleteaction.Enabled = True
    mnu_exporttoini.Enabled = True
    mnu_export.Enabled = True

    'mnu_scheduler.Enabled = True
    mnu_simulate.Enabled = True
    mnu_sync.Enabled = True
    cmd_sync.Enabled = True
Else
    mnu_editaction.Enabled = False
    mnu_deleteaction.Enabled = False
    mnu_export.Enabled = False
    mnu_exporttoini.Enabled = False

    'mnu_scheduler.Enabled = False
    mnu_simulate.Enabled = False
    mnu_sync.Enabled = False
    cmd_sync.Enabled = False
End If


End Sub


Sub SchOn()

Img_Green.Visible = True
Img_Red.Visible = False

End Sub

Sub SchOff()

Img_Green.Visible = False
Img_Red.Visible = True

End Sub

Sub LoadTimerSettings()

sIniFile = GlobalAppPath & "fqs-timer.tmr"

tmrTemplate = IniDaj("Timer", "Template", "- none -")
tmrDate = IniDaj("Timer", "Date", "")
tmrTime = IniDaj("Timer", "Time", "")

If UCase(IniDaj("Timer", "Daily", "NO")) = "YES" Then
    tmrDaily = True
Else
    tmrDaily = False
End If

If UCase(IniDaj("Timer", "Enabled", "NO")) = "YES" Then
    tmrEnabled = True
Else
    tmrEnabled = False
End If

tmr_sch.Enabled = True

End Sub

Private Sub tmr_sch_Timer()

If tmrEnabled = False Then
    Img_Green.Visible = False
    Img_Red.Visible = True
    
    Img_Red.ToolTipText = "Scheduler - disabled"
    Exit Sub
Else
    Img_Green.Visible = True
    Img_Red.Visible = False
    
    If tmrDaily = True Then
        Img_Green.ToolTipText = "Scheduler - enabled <> " & "Template: " & tmrTemplate & ", starting: Daily at " & tmrTime
    Else
        Img_Green.ToolTipText = "Scheduler - enabled <> " & "Template: " & tmrTemplate & ", starting: " & tmrDate & " at " & tmrTime
    End If
    
End If
    
    
Dim tmpDateNow As String
Dim tmpTimeNow As String

tmpDateNow = Format(Date, "dd.mm.yyyy")
tmpTimeNow = Format(Time, "hh:mm")

' check for date and time condition
If tmpDateNow = tmrDate And tmpTimeNow = tmrTime Then
    tmr_sch.Enabled = False
   
    ' if is not daily - write to ini ENABLED=NO
    If tmrDaily = False Then
        sIniFile = GlobalAppPath & "fqs-timer.tmr"
        IniPisi "Timer", "Enabled", "NO"
        
        LoadTimerSettings
    Else
        ' check if is started today
        If tmrDailyStarted = tmpDateNow Then
            Exit Sub
        End If
    End If
    
    ' ++++++++++ start sync +++++++++++++++++
    
    ' set 'today is started=True'
    tmrDailyStarted = tmpDateNow
    
    SyncMode = "Scheduler"
    
    Dim oldTemplate As String

    oldTemplate = cmb_template
    cmb_template = tmrTemplate
    
    If Data1.Recordset.RecordCount = 0 Then
        tmp = MsgBox("Scheduler cannot start becouse no actions are defined! Please create some actions and then try again", vbOKOnly + vbCritical, "Error")
        Exit Sub
    End If

    frm_sync.Show vbModal
    
    cmb_template = oldTemplate
    tmr_sch.Enabled = True
End If

End Sub

Sub PromjeniViewReportMeni()

mnu_viewreport.Caption = "View '" & cmb_template & ".log'"

If cmb_template = "- none -" Then
    mnu_viewreport.Enabled = False
Else
    mnu_viewreport.Enabled = True
End If

End Sub
