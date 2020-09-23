Attribute VB_Name = "modMakeDir"
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Public Function MakeDirectory(szDirectory As String) As Boolean

Dim strFolder As String
Dim szRslt As String

On Error GoTo IllegalFolderName

If Right(szDirectory, 1) <> "\" Then szDirectory = szDirectory & "\"

strFolder = szDirectory

szRslt = Dir(strFolder, 63)

While szRslt = ""
    DoEvents
    szRslt = Dir(strFolder, 63)
    strFolder = Left(strFolder, Len(strFolder) - 1)
    If strFolder = "" Then GoTo IllegalFolderName
Wend

If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

While strFolder <> szDirectory
    strFolder = Left(szDirectory, Len(strFolder) + 1)
    If Right(strFolder, 1) = "\" Then MkDir strFolder
Wend

MakeDirectory = True

Exit Function

IllegalFolderName:
    Call MsgBox("Could not Create Destination Folder.", vbExclamation)
    
End Function

Public Function GetText(hWnd As Long) As String

Dim hWndChild  As Long, nSize As Long
Dim sBuffer As String * 32
Dim lmsg As String * 260
hWndChild = GetWindow(hWnd, GW_CHILD)
   
Do While hWndChild <> 0

    nSize = GetClassName(hWndChild, sBuffer, 32)
       
    If nSize Then
        If Left$(sBuffer, nSize) = "Edit" Then
            lmsg = Space(64)
            Call GetWindowText(hWndChild, lmsg, 260)
            GetText = Left(lmsg, InStr(lmsg, vbNullChar) - 1)
            Exit Function
        End If
    End If
      
    hWndChild = GetWindow(hWndChild, GW_HWNDNEXT)
       
Loop

End Function

