VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_sync 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Synchronizing - please wait ..."
   ClientHeight    =   6090
   ClientLeft      =   3735
   ClientTop       =   2805
   ClientWidth     =   6570
   Icon            =   "frm_sync.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "!"
      Height          =   285
      Left            =   90
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   225
      TabIndex        =   7
      Top             =   1575
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.ListBox l3 
      Height          =   1620
      Left            =   225
      TabIndex        =   6
      Top             =   3105
      Visible         =   0   'False
      Width           =   5955
   End
   Begin VB.ListBox l2 
      Height          =   1230
      Left            =   3330
      TabIndex        =   5
      Top             =   225
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.ListBox l1 
      Height          =   1230
      Left            =   225
      TabIndex        =   4
      Top             =   225
      Visible         =   0   'False
      Width           =   2940
   End
   Begin MSComctlLib.ProgressBar Pb1 
      Height          =   285
      Left            =   45
      TabIndex        =   3
      Top             =   5355
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   900
      Top             =   5670
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   3420
      TabIndex        =   1
      Top             =   1530
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5235
      Left            =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   45
      Width           =   6450
   End
   Begin VB.CommandButton cmd_close 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   330
      Left            =   2745
      TabIndex        =   0
      Top             =   5715
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Press ESC to cancel !"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4140
      TabIndex        =   9
      Top             =   5760
      Visible         =   0   'False
      Width           =   2310
   End
End
Attribute VB_Name = "frm_sync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpPathFrom As String   ' source folder
Dim tmpPathTo As String     ' destination folder
Dim tmpFilter As String     ' File1 filter
Dim tmpInclSub As Boolean   ' Include Subfolders

Dim tmpSomeErrorsOcurred As Boolean

Dim ActionEnabled As String
Dim ActionFileDate As String
Dim ActionFileSize As String

Dim tmpFilename As String   ' selected filename from File1

Dim tmpSsize As Long    ' source file size
Dim tmpSdate As Date    ' source file date

Dim tmpDsize As Long    ' destination file size
Dim tmpDdate As Date    ' destination file date

Dim tmpSExists As Boolean   ' source exists?
Dim tmpDExists As Boolean   ' destination exists?

Dim tmpSDexists As Boolean  ' source folder exists?
Dim tmpDDexists As Boolean  ' destination folder exists?

Dim FileCopyed As Boolean

Dim SumFiles As Long        ' number of copyed files
Dim SumBytes As Long        ' sum of bytes copyed

Dim GlobalSumFiles As Long        ' number of copyed files
Dim GlobalSumBytes As Long        ' sum of bytes copyed

Dim SyncCanceled As Boolean

Private Sub cmd_close_Click()
    Unload Me
    
End Sub

Private Sub Command1_Click()

SearchForFiles


End Sub


Private Sub SearchForFiles()

Me.Caption = "Creating file list, please wait ..."

l3.Clear

    Dim A As Long, B As Long
    Dim tmpPathFromLen As Long
    
    l2.Clear
    Dir1.path = tmpPathFrom
    tmpPathFromLen = Len(tmpPathFrom)
    
    l2.AddItem Dir1.path
    
    File1.Pattern = tmpFilter

    DoEvents


        Do
            DoEvents
            If SyncCanceled = True Then Exit Sub
            
                For A = 0 To Dir1.ListCount - 1
                    DoEvents
                    
                    If SyncCanceled = True Then Exit Sub
                    
                    l1.AddItem Dir1.List(A)
                Next A
                    
                    Dir1.path = l1.List(0)
                    
                    l2.AddItem l1.List(0)
                    l1.RemoveItem (0)
                    'Dir1.path = l1.List(0)
         
         Loop Until l1.ListCount = 0

                
                
                For B = 0 To l2.ListCount - 1
                    Dir1.path = l2.List(0)

                    DoEvents
                    If SyncCanceled = True Then Exit Sub
                    
                        For A = 0 To File1.ListCount - 1

                            DoEvents
                            If SyncCanceled = True Then Exit Sub
                            
                                If l3.ListCount > 32000 Then
                                
                                    tmp = MsgBox("Error: File count overflw! (max. 32000 files allowed)", vbOKOnly + vbCritical, "Error")
                                    Exit Sub
                                    
                                Else

                                    If Mid(l2.List(0), Len(l2.List(0)), 1) = "\" Then
                                        l3.AddItem Right$(l2.List(0) & File1.List(A), Len(l2.List(0) & File1.List(A)) - tmpPathFromLen)
                                        Debug.Print Right$(l2.List(0) & File1.List(A), Len(l2.List(0) & File1.List(A)) - tmpPathFromLen)
                                    Else
                                        l3.AddItem Right$(l2.List(0) & "\" & File1.List(A), Len(l2.List(0) & "\" & File1.List(A)) - tmpPathFromLen)
                                        Debug.Print Right$(l2.List(0) & "\" & File1.List(A), Len(l2.List(0) & "\" & File1.List(A)) - tmpPathFromLen)
                                    End If
                                
                                End If
                        
                        Next A
                            
'                            Label1.Caption = Label1.Caption + File1.ListCount


                            DoEvents
                            l2.RemoveItem (0)
                Next B

Me.Caption = "File list created! " & l3.ListCount & " files selected by criteria"

End Sub

Private Sub Dir1_Change()
    File1.path = Dir1.path

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next

If KeyCode = vbKeyEscape Then
    
        Label1.Visible = False
    
        SyncCanceled = True
        DoEvents
        
        List1.AddItem " "
        List1.AddItem "          " & "-----------------------------------"
        List1.AddItem "          Synchronization canceled by user!!"
        List1.AddItem "          " & "-----------------------------------"
        List1.AddItem "          "

        Select Case SyncMode
            Case "Scheduler"
                List1.AddItem "          " & "Finished: " & Format(Date, "dd.mm.yyyy") & " at " & Format(Time, "hh:mm:ss")
    
            Case "Simulate"
                List1.AddItem "          " & "Finished: " & Format(Date, "dd.mm.yyyy") & " at " & Format(Time, "hh:mm:ss")
    
            Case Else
                List1.AddItem "          " & "Finished: " & Format(Date, "dd.mm.yyyy") & " at " & Format(Time, "hh:mm:ss")

        End Select

        List1.AddItem " "

        'List1.AddItem "          " & "Total copyed = " & GlobalSumFiles & " files"
        'List1.AddItem "          " & "Total copyed = " & Format(GlobalSumBytes, "#,##0") & " bytes"
        'List1.AddItem " "

        List1.AddItem "========================================="

        List1.AddItem " "
        List1.AddItem "     report saved in file: " & frm_main.cmb_template & ".log"

        ' save report to LOG file
        Open GlobalAppPath & frm_main.cmb_template & ".log" For Append As #1

        For qwe = 0 To List1.ListCount - 2

            Print #1, List1.List(qwe)

        Next qwe

        Close #1

        NewFlag = False
        frm_main.Data1.Refresh

        Me.Caption = "User canceled!"
        
        MousePointer = vbNormal
        DoEvents

        cmd_close.Enabled = True
        ChDir GlobalAppPath

        ' close window if is started by scheduler
        If SyncMode = "Scheduler" Then cmd_close_Click
        
        Label1.Visible = False

End If

End Sub

Private Sub Form_Load()
    CenterForm_global Me
    Pb1.Min = 0
    Pb1.Value = 0
    
Unload frm_question
frm_main.Hide

SyncCanceled = False

cmd_close.Enabled = False

    Timer1.Enabled = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If cmd_close.Enabled = False Then
    Cancel = True
    Exit Sub
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frm_main.Show
    
End Sub

Private Sub Timer1_Timer()

Timer1.Enabled = False

On Error GoTo greska

MousePointer = vbHourglass
DoEvents

NewFlag = True              ' LOCK reposition of DATA1

List1.Clear

List1.AddItem " "

Select Case SyncMode
    Case "Scheduler"
        List1.AddItem "Started: " & Format(Date, "dd.mm.yyyy") & " at " & Format(Time, "hh:mm:ss") & " (Scheduler for template: " & frm_main.cmb_template & ")"
    
    Case "Simulate"
        List1.AddItem "Started: " & Format(Date, "dd.mm.yyyy") & " at " & Format(Time, "hh:mm:ss") & " (Simulate template: " & frm_main.cmb_template & ")"
    
    Case Else
        List1.AddItem "Started: " & Format(Date, "dd.mm.yyyy") & " at " & Format(Time, "hh:mm:ss") & " (template: " & frm_main.cmb_template & ")"

End Select


Label1.Visible = True
DoEvents


With frm_main.Data1.Recordset

.MoveFirst

GlobalSumBytes = 0
GlobalSumFiles = 0

'=================================================
' -------- begin of ACTIONS loop --------------
For qwe = 1 To .RecordCount

    If SyncCanceled = True Then Exit Sub
    
    SumBytes = 0
    SumFiles = 0
    
    ' load atributes for files
    tmpFilter = "" & !FileFilter
    tmpPathFrom = CheckDirSep("" & !FromFolder)
    tmpPathTo = CheckDirSep("" & !ToFolder)
    
    tmpInclSub = False
    If UCase("" & !InclSub) = "YES" Then tmpInclSub = True
    
    ActionEnabled = UCase("" & !Enabled)
    ActionFileDate = "" & !FileDate
    ActionFileSize = "" & !FileSize
    
    List1.AddItem " "
    List1.AddItem "Action " & qwe & " of " & .RecordCount
    List1.AddItem "    From folder: " & tmpPathFrom
    List1.AddItem "    To folder: " & tmpPathTo
    

    If SyncCanceled = True Then Exit Sub

    ' select way to copy files
    
    If tmpInclSub = True Then
        List1.AddItem "    Include subfolders: Yes"
        CopyWithFolders
    Else
        List1.AddItem "    Include subfolders: No"
        CopyWithoutSubFolders
    End If
    
        List1.Selected(List1.ListCount - 1) = True

    .MoveNext

Next qwe
'===============================================
' end of all actions

.MoveFirst

End With


If SyncCanceled = True Then Exit Sub


List1.AddItem " "
List1.AddItem "          " & "-----------------------------------"

Select Case SyncMode
    Case "Scheduler"
        List1.AddItem "          " & "Finished: " & Format(Date, "dd.mm.yyyy") & " at " & Format(Time, "hh:mm:ss")
    
    Case "Simulate"
        List1.AddItem "          " & "Finished: " & Format(Date, "dd.mm.yyyy") & " at " & Format(Time, "hh:mm:ss")
    
    Case Else
        List1.AddItem "          " & "Finished: " & Format(Date, "dd.mm.yyyy") & " at " & Format(Time, "hh:mm:ss")

End Select

List1.AddItem " "

List1.AddItem "          " & "Total copyed = " & GlobalSumFiles & " files"
List1.AddItem "          " & "Total copyed = " & Format(GlobalSumBytes, "#,##0") & " bytes"
List1.AddItem " "

List1.AddItem "========================================="

List1.AddItem " "
List1.AddItem "     report saved in file: " & frm_main.cmb_template & ".log"

List1.Selected(List1.ListCount - 1) = True

If SyncCanceled = True Then Exit Sub


' save report to LOG file
Open GlobalAppPath & frm_main.cmb_template & ".log" For Append As #1

For qwe = 0 To List1.ListCount - 2

    Print #1, List1.List(qwe)

Next qwe

Close #1

NewFlag = False
frm_main.Data1.Refresh

Me.Caption = "Finished!"
MousePointer = vbNormal
DoEvents

cmd_close.Enabled = True
ChDir GlobalAppPath

' close window if is started by scheduler
If SyncMode = "Scheduler" Then cmd_close_Click

Label1.Visible = False

Exit Sub




' error handler
greska:

Label1.Visible = False

List1.AddItem " "
List1.AddItem ">> Unhandled fatal error: " & Err & " - " & Error & " !!!!!"
List1.AddItem " "
List1.AddItem "========================================="

List1.AddItem " "
List1.AddItem "     report saved in file: " & frm_main.cmb_template & ".log"

List1.Selected(List1.ListCount - 1) = True

' save report to LOG file
Open GlobalAppPath & frm_main.cmb_template & ".log" For Append As #1

For qwe = 0 To List1.ListCount - 2

    Print #1, List1.List(qwe)

Next qwe

Close #1

NewFlag = False
frm_main.Data1.Refresh

Me.Caption = "Finished!"
MousePointer = vbNormal
DoEvents

cmd_close.Enabled = True
ChDir GlobalAppPath

' close window if is started by scheduler
If SyncMode = "Scheduler" Then cmd_close_Click

tmp = MsgBox("Unhandled fatal error occured! Please see report or LOG file", vbOKOnly + vbCritical, "Error")

End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Sub CopyWithoutSubFolders()

l3.Clear
Pb1.Value = 0
tmpSomeErrorsOcurred = False

Me.Caption = "Synchronizing - please wait ..."

Dim xtmp
xtmp = 1

' skip if source folder don't exists
    If DirExists(tmpPathFrom) Then
        File1.path = tmpPathFrom
        File1.Pattern = tmpFilter
        
        ' move File1 to l3
        If File1.ListCount > 0 Then
            For wer = 0 To File1.ListCount - 1
            l3.AddItem File1.List(wer)
            Next wer
        End If
        
    Else
        List1.AddItem " "
        List1.AddItem "    - Error: source folder not exists!"
        
        GoTo disabledskip
    End If
    
' skip action if disabled
    If ActionEnabled <> "YES" Then
        List1.AddItem " "
        List1.AddItem "    File filter: " & tmpFilter & " - ACTION DISABLED!"
        List1.AddItem " "
                
        GoTo disabledskip
    End If
    
' check l3 for files
    If l3.ListCount = 0 Then
        List1.AddItem "    File filter: " & tmpFilter & " (no files found!)"
        List1.AddItem " "
        
    Else
    
        List1.AddItem "    File filter: " & tmpFilter & " (" & l3.ListCount & " files found)"
        List1.AddItem " "
    
    ' copy loop
    Pb1.Max = l3.ListCount - 1
    
    If SyncCanceled = True Then Exit Sub
    
    For wer = 0 To l3.ListCount - 1
        
        If SyncCanceled = True Then Exit Sub
        
        Pb1.Value = wer
        
        tmpFilename = ""
        tmpSsize = Empty
        tmpSdate = Empty
        tmpDsize = Empty
        tmpDdate = Empty
        
        tmpFilename = l3.List(wer)
        
        'check for source folder and file
        tmpSExists = False
        tmpSDexists = False
        
        If DirExists(tmpPathFrom) Then
            tmpSDexists = True
            If FileExist(tmpPathFrom & tmpFilename) Then
                tmpSExists = True
                tmpSsize = FileLen(tmpPathFrom & tmpFilename)
                tmpSdate = FileDateTime(tmpPathFrom & tmpFilename)
            End If
        End If
        
        'check for destination folder and file
        tmpDExists = False
        tmpDDexists = False
        
        If DirExists(tmpPathTo) Then
            tmpDDexists = True
            If FileExist(tmpPathTo & tmpFilename) Then
                tmpDExists = True
                tmpDsize = FileLen(tmpPathTo & tmpFilename)
                tmpDdate = FileDateTime(tmpPathTo & tmpFilename)
            End If
        End If
        
        
        ' if source or destination folder not exists - exit for
        If tmpSDexists = False Or tmpDDexists = False Then
            If tmpSDexists = False Then
                List1.AddItem "    - Error: source folder not exists!"
            End If
            
            If tmpDDexists = False Then
                List1.AddItem "    - Error: destination folder not exists!"
            End If
            
            Exit For
        End If

        FileCopyed = False
        
        If SyncCanceled = True Then Exit Sub
        
        ' +++++++++++ start copy +++++++++++++++
        If tmpDExists = False Then
            ' if destination file not exists, just copy file
            If SyncMode <> "Simulate" Then xtmp = CopyFile(tmpPathFrom & tmpFilename, tmpPathTo & tmpFilename, False)
            FileCopyed = True
            
            SumBytes = SumBytes + tmpSsize
            SumFiles = SumFiles + 1
        Else
            
            ' if destination file exists, check for conditions
        Select Case ActionFileSize
            Case "<>"
                If tmpSsize <> tmpDsize Then
                    If SyncMode <> "Simulate" Then xtmp = CopyFile(tmpPathFrom & tmpFilename, tmpPathTo & tmpFilename, False)
                    FileCopyed = True
                
                    SumBytes = SumBytes + tmpSsize
                    SumFiles = SumFiles + 1
                End If
                
            Case ">"
                If tmpSsize > tmpDsize Then
                    If SyncMode <> "Simulate" Then xtmp = CopyFile(tmpPathFrom & tmpFilename, tmpPathTo & tmpFilename, False)
                    FileCopyed = True
                
                    SumBytes = SumBytes + tmpSsize
                    SumFiles = SumFiles + 1
                End If
            
            Case Else
                    If SyncMode <> "Simulate" Then xtmp = CopyFile(tmpPathFrom & tmpFilename, tmpPathTo & tmpFilename, False)
                    FileCopyed = True
                    
                    SumBytes = SumBytes + tmpSsize
                    SumFiles = SumFiles + 1
                    
        End Select
                
        If FileCopyed = False Then
                
                Select Case ActionFileDate
                    Case "<>"
                        If tmpSdate <> tmpDdate Then
                            If SyncMode <> "Simulate" Then xtmp = CopyFile(tmpPathFrom & tmpFilename, tmpPathTo & tmpFilename, False)
                            FileCopyed = True
                        
                            SumBytes = SumBytes + tmpSsize
                            SumFiles = SumFiles + 1
                        End If
                
                    Case ">"
                        If tmpSdate > tmpDdate Then
                            If SyncMode <> "Simulate" Then xtmp = CopyFile(tmpPathFrom & tmpFilename, tmpPathTo & tmpFilename, False)
                            FileCopyed = True
                        
                            SumBytes = SumBytes + tmpSsize
                            SumFiles = SumFiles + 1
                        End If
                    
                    Case Else
                            If SyncMode <> "Simulate" Then xtmp = CopyFile(tmpPathFrom & tmpFilename, tmpPathTo & tmpFilename, False)
                            FileCopyed = True
                            
                            SumBytes = SumBytes + tmpSsize
                            SumFiles = SumFiles + 1
                End Select
        End If
        
        End If
        
        If FileCopyed = True Then
            If xtmp <> 0 Then
                List1.AddItem "     " & LCase(tmpFilename) & " - (copyed)"
            End If
        Else
            If FlagShowAllInReport = True Then
                List1.AddItem "     " & LCase(tmpFilename) & " - (no criteria matched!)"
            End If
        End If
        
        If xtmp = 0 Then
            List1.AddItem "     >>  " & LCase(tmpFilename) & " - (Error: File is Read-only or locked by the system!)"
            tmpSomeErrorsOcurred = True
            xtmp = 1
        End If
        
        
        List1.Selected(List1.ListCount - 1) = True
        DoEvents
    
    'if LOG list count>30000 dump log list and clear list
    If List1.ListCount > 30000 Then
        DumpLogList
    End If
    
    
    Next wer
    ' end of loop for current action


    End If
    
disabledskip:
    
    DoEvents
    
    If tmpSomeErrorsOcurred = True Then
        List1.AddItem " "
        List1.AddItem "        >> Some errors ocurred, see log file for details!"
    End If
    
    List1.AddItem " "
    List1.AddItem "        " & SumFiles & " of " & l3.ListCount & " file(s) copyed, size of file(s)= " & Format(SumBytes, "#,##0") & " bytes"
    
    GlobalSumBytes = GlobalSumBytes + SumBytes
    GlobalSumFiles = GlobalSumFiles + SumFiles
    
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Sub CopyWithFolders()

tmpSomeErrorsOcurred = False
l3.Clear
Pb1.Value = 0

Dim xtmp
xtmp = 1

' skip action if disabled
    If ActionEnabled <> "YES" Then
        List1.AddItem " "
        List1.AddItem "    File filter: " & tmpFilter & " - ACTION DISABLED!"
        List1.AddItem " "
                
        GoTo disabledskip
    End If

' skip if source folder don't exists
    If DirExists(tmpPathFrom) Then
        
        Dir1.path = tmpPathFrom
        
        ' if subfolders not exists - just copy files
        If Dir1.ListCount = 0 Then
            
            List1.AddItem " "
            List1.AddItem "    (No subfolders detected)"
            
            CopyWithoutSubFolders   ' if subfolders not exists - just copy files
            Exit Sub
        
        Else
            SearchForFiles  ' load l3 with files
                
            If SyncCanceled = True Then Exit Sub
            
            Me.Caption = "Synchronizing - please wait ..."
        End If

    Else
        List1.AddItem " "
        List1.AddItem "    - Error: source folder not exists!"
        
        'Me.Caption = "Synchronizing - please wait ..."
        
        GoTo disabledskip
    End If
    
' skip if destination-root folder don't exists
    If Not DirExists(tmpPathTo) Then
        
        List1.AddItem " "
        List1.AddItem "    - Error: destination-root folder not exists!"
        
        GoTo disabledskip
    End If

' check l3 for files
    If l3.ListCount = 0 Then
        List1.AddItem "    File filter: " & tmpFilter & " (no files found!)"
        List1.AddItem " "
        
    Else
    
        List1.AddItem "    File filter: " & tmpFilter & " (" & l3.ListCount & " files found)"
        List1.AddItem " "
    
    'check and copy
    Pb1.Max = l3.ListCount - 1
    
    If SyncCanceled = True Then Exit Sub
    
    For wer = 0 To l3.ListCount - 1
        
        If SyncCanceled = True Then Exit Sub
        
        Pb1.Value = wer
        
        tmpFilename = ""
        tmpSsize = Empty
        tmpSdate = Empty
        tmpDsize = Empty
        tmpDdate = Empty
        
        tmpFilename = l3.List(wer)
        
        Dim tmpSegmentFolder As String
        Dim cnt
        
        ' select Segment folder and Filename from tmpFileName
        For cnt = Len(tmpFilename) To 1 Step -1
            If Mid$(tmpFilename, cnt, 1) = "\" Then
                tmpSegmentFolder = Left$(tmpFilename, cnt)
                tmpFilename = Right$(tmpFilename, Len(tmpFilename) - cnt)
                Exit For
            End If
        Next cnt
        
        
        If cnt = 0 Then
            tmpSegmentFolder = ""
        End If
      
        'check for source folder and file
        tmpSExists = False
        tmpSDexists = False
        
        If DirExists(tmpPathFrom & tmpSegmentFolder) Then
            tmpSDexists = True
            If FileExist(tmpPathFrom & tmpSegmentFolder & tmpFilename) Then
                tmpSExists = True
                tmpSsize = FileLen(tmpPathFrom & tmpSegmentFolder & tmpFilename)
                tmpSdate = FileDateTime(tmpPathFrom & tmpSegmentFolder & tmpFilename)
            End If
        Else
            List1.AddItem "    - Error: source folder not exists! (" & LCase(tmpPathFrom & tmpSegmentFolder) & ")"
            GoTo SourceFolderNotExist
        End If
        
        
        'check for destination folder and file
        tmpDExists = False
        tmpDDexists = False
        
        If DirExists(tmpPathTo & tmpSegmentFolder) Then
            tmpDDexists = True
            If FileExist(tmpPathTo & tmpSegmentFolder & tmpFilename) Then
                tmpDExists = True
                tmpDsize = FileLen(tmpPathTo & tmpSegmentFolder & tmpFilename)
                tmpDdate = FileDateTime(tmpPathTo & tmpSegmentFolder & tmpFilename)
            End If
        Else
            ' check for destination root folder
            If DirExists(tmpPathTo) Then
                
                ' make destination folders
                KreirajPut (tmpPathTo & tmpSegmentFolder)
                List1.AddItem "    - Note: created destination folder (" & LCase(tmpPathTo & tmpSegmentFolder) & ")"
            
            Else
                List1.AddItem "    - Error: destination folder not exists! (" & LCase(tmpPathTo) & ")"
                GoTo SourceFolderNotExist
            
            End If
        
        End If
  
        If SyncCanceled = True Then Exit Sub
        
        FileCopyed = False
        
        ' +++++++++++ start copy +++++++++++++++
        If tmpDExists = False Then
            ' if destination file not exists - just copy
            If SyncMode <> "Simulate" Then xtmp = CopyFile(tmpPathFrom & tmpSegmentFolder & tmpFilename, tmpPathTo & tmpSegmentFolder & tmpFilename, False)
            FileCopyed = True
            
            SumBytes = SumBytes + tmpSsize
            SumFiles = SumFiles + 1
        Else
            
            ' if exists, check for conditions
        Select Case ActionFileSize
            Case "<>"
                If tmpSsize <> tmpDsize Then
                    If SyncMode <> "Simulate" Then xtmp = CopyFile(tmpPathFrom & tmpSegmentFolder & tmpFilename, tmpPathTo & tmpSegmentFolder & tmpFilename, False)
                    FileCopyed = True
                
                    SumBytes = SumBytes + tmpSsize
                    SumFiles = SumFiles + 1
                End If
                
            Case ">"
                If tmpSsize > tmpDsize Then
                    If SyncMode <> "Simulate" Then xtmp = CopyFile(tmpPathFrom & tmpSegmentFolder & tmpFilename, tmpPathTo & tmpSegmentFolder & tmpFilename, False)
                    FileCopyed = True
                
                    SumBytes = SumBytes + tmpSsize
                    SumFiles = SumFiles + 1
                End If
            
            Case Else
                    If SyncMode <> "Simulate" Then xtmp = CopyFile(tmpPathFrom & tmpSegmentFolder & tmpFilename, tmpPathTo & tmpSegmentFolder & tmpFilename, False)
                    FileCopyed = True
                    
                    SumBytes = SumBytes + tmpSsize
                    SumFiles = SumFiles + 1
                    
        End Select
                
        If FileCopyed = False Then
                
                Select Case ActionFileDate
                    Case "<>"
                        If tmpSdate <> tmpDdate Then
                            If SyncMode <> "Simulate" Then xtmp = CopyFile(tmpPathFrom & tmpSegmentFolder & tmpFilename, tmpPathTo & tmpSegmentFolder & tmpFilename, False)
                            FileCopyed = True
                        
                            SumBytes = SumBytes + tmpSsize
                            SumFiles = SumFiles + 1
                        End If
                
                    Case ">"
                        If tmpSdate > tmpDdate Then
                            If SyncMode <> "Simulate" Then xtmp = CopyFile(tmpPathFrom & tmpSegmentFolder & tmpFilename, tmpPathTo & tmpSegmentFolder & tmpFilename, False)
                            FileCopyed = True
                        
                            SumBytes = SumBytes + tmpSsize
                            SumFiles = SumFiles + 1
                        End If
                    
                    Case Else
                            If SyncMode <> "Simulate" Then xtmp = CopyFile(tmpPathFrom & tmpSegmentFolder & tmpFilename, tmpPathTo & tmpSegmentFolder & tmpFilename, False)
                            FileCopyed = True
                            
                            SumBytes = SumBytes + tmpSsize
                            SumFiles = SumFiles + 1
                End Select
        End If
        
        End If
        
        If FileCopyed = True Then
            If xtmp <> 0 Then
                List1.AddItem "     " & LCase(tmpSegmentFolder & tmpFilename) & " - (copyed)"
            End If
        Else
            If FlagShowAllInReport = True Then
                List1.AddItem "     " & LCase(tmpSegmentFolder & tmpFilename) & " - (no criteria matched!)"
            End If
        End If
        
        If xtmp = 0 Then
            List1.AddItem "     >>  " & LCase(tmpSegmentFolder & tmpFilename) & " - (Error: File is Read-only or locked by the system!)"
            tmpSomeErrorsOcurred = True
            xtmp = 1
        End If
        
        
SourceFolderNotExist:

        List1.Selected(List1.ListCount - 1) = True
        DoEvents
    
    
    'if LOG list count>30000 dump log list and clear list
    If List1.ListCount > 30000 Then
        DumpLogList
    End If
    
    
    Next wer
    ' end of loop fo rcurrent action


    End If
    
disabledskip:
    
    DoEvents
    
    If tmpSomeErrorsOcurred = True Then
        List1.AddItem " "
        List1.AddItem "        >> Some errors ocurred, see log file for details!"
    End If
    
    List1.AddItem " "
    List1.AddItem "        " & SumFiles & " of " & l3.ListCount & " file(s) copyed, size of file(s)= " & Format(SumBytes, "#,##0") & " bytes"
    
    GlobalSumBytes = GlobalSumBytes + SumBytes
    GlobalSumFiles = GlobalSumFiles + SumFiles

End Sub

Sub DumpLogList()

On Error Resume Next

Dim tx

        ' save report to LOG file
        Open GlobalAppPath & frm_main.cmb_template & ".log" For Append As #1

        For tx = 0 To List1.ListCount - 2

            Print #1, List1.List(tx)

        Next tx

        Close #1

List1.Clear

End Sub
