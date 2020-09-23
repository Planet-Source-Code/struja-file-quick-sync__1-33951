Attribute VB_Name = "modFolderBrowse"
Option Explicit

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

'BrowseInfo ulFlags
Public Const BIF_RETURNONLYFSDIRS = &H1        'Only return file system directories.
Public Const BIF_DONTGOBELOWDOMAIN = &H2       'Do not include network folders below the domain level in the dialog box's tree view control.
Public Const BIF_STATUSTEXT = &H4              'Include a status area in the dialog box.
Public Const BIF_RETURNFSANCESTORS = &H8       'Only return file system ancestors. An ancestor is a subfolder that is beneath the root folder in the namespace hierarchy.
Public Const BIF_EDITBOX = &H10                '(SHELL32.DLL Version 4.71). Include an edit control in the browse dialog box that allows the user to type the name of an item.
Public Const BIF_VALIDATE = &H20               '(SHELL32.DLL Version 4.71). If the user types an invalid name into the edit box, the browse dialog will call the application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message.
Public Const BIF_USENEWUI = &H40               '(SHELL32.DLL Version 5.0). Use the new user interface, including an edit box.
Public Const BIF_NEWDIALOGSTYLE = &H50         '(SHELL32.DLL Version 5.0). Use the new user interface.
Public Const BIF_BROWSEINCLUDEURLS = &H80      '(SHELL32.DLL Version 5.0). The browse dialog box can display URLs. The BIF_USENEWUI and BIF_BROWSEINCLUDEFILES flags must also be set.
Public Const BIF_BROWSEFORCOMPUTER = &H1000    'Only return computers.
Public Const BIF_BROWSEFORPRINTER = &H2000     'Only return network printers.
Public Const BIF_BROWSEINCLUDEFILES = &H4000   '(SHELL32.DLL Version 4.71). The browse dialog will display files as well as folders.
Public Const BIF_SHAREABLE = &H8000            '(SHELL32.DLL Version 5.0). The browse dialog box can display shareable resources on remote systems. The BIF_USENEWUI flag must also be set.

'BrowseInfo pIDLRoot(Do not use these with new style dialog)
Const Default = 0
Const Internet = 1
Const Programs = 2
Const ControlPanel = 3
Const Printers = 4
Const MyDocuments = 5
Const Favorites = 6
Const StartUp = 7
Const Recent = 8
Const SendTo = 9
Const RecycleBin = 10
Const StartMenu = 11
Const Desktop = 16
Const MyComputer = 17
Const Network = 18
Const Nethood = 19
Const Fonts = 20
Const Templates = 21
Const ApplicationData = 26
Const PrintHood = 27
Const TemporaryInternetFiles = 32
Const Cookies = 33
Const History = 34

Const BFFM_ENABLEOK = &H465
Const BFFM_SETSELECTION = &H466
Const BFFM_SETSTATUSTEXT = &H464

Const BFFM_INITIALIZED = 1
Const BFFM_SELCHANGED = 2
Const BFFM_VALIDATEFAILED = 3

Const MAX_PATH = 255

Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallBack As Long
    lParam As Long
    iImage As Integer
End Type

Public StartFolder As String
Public SpecialFolder As Long
Public CurrentSelection As String * MAX_PATH
Public OKEnable As Boolean
Public szDisplay As String
Public hWndText As Long

Public Function FolderBrowse(hwndForm As Long, szInstruction As String, Optional lFlags As Long) As String

    Dim BI As BrowseInfo
    Dim lRslt As Long
    Dim strReturn As String

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Public Const BIF_RETURNONLYFSDIRS = &H1        'Only return file system directories.
'Public Const BIF_DONTGOBELOWDOMAIN = &H2       'Do not include network folders below the domain level in the dialog box's tree view control.
'Public Const BIF_STATUSTEXT = &H4              'Include a status area in the dialog box.
'Public Const BIF_RETURNFSANCESTORS = &H8       'Only return file system ancestors. An ancestor is a subfolder that is beneath the root folder in the namespace hierarchy.
'Public Const BIF_EDITBOX = &H10                '(SHELL32.DLL Version 4.71). Include an edit control in the browse dialog box that allows the user to type the name of an item.
'Public Const BIF_VALIDATE = &H20               '(SHELL32.DLL Version 4.71). If the user types an invalid name into the edit box, the browse dialog will call the application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message.
'Public Const BIF_USENEWUI = &H40               '(SHELL32.DLL Version 5.0). Use the new user interface, including an edit box.
'Public Const BIF_NEWDIALOGSTYLE = &H50         '(SHELL32.DLL Version 5.0). Use the new user interface.
'Public Const BIF_BROWSEINCLUDEURLS = &H80      '(SHELL32.DLL Version 5.0). The browse dialog box can display URLs. The BIF_USENEWUI and BIF_BROWSEINCLUDEFILES flags must also be set.
'Public Const BIF_BROWSEFORCOMPUTER = &H1000    'Only return computers.
'Public Const BIF_BROWSEFORPRINTER = &H2000     'Only return network printers.
'Public Const BIF_BROWSEINCLUDEFILES = &H4000   '(SHELL32.DLL Version 4.71). The browse dialog will display files as well as folders.
'Public Const BIF_SHAREABLE = &H8000            '(SHELL32.DLL Version 5.0). The browse dialog box can display shareable resources on remote systems. The BIF_USENEWUI flag must also be set.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    With BI
        .hWndOwner = hwndForm
        .lpszTitle = szInstruction
        .pIDLRoot = SpecialFolder
        .ulFlags = lFlags + BIF_VALIDATE
        '.ulFlags = BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE  '+ BIF_USENEWUI
        .pszDisplayName = String$(MAX_PATH, 0)
        .lpfnCallBack = DummyFunction(AddressOf BrowseCallbackProc)
    End With

    lRslt = SHBrowseForFolder(BI)

    If lRslt Then
        lRslt = SHGetPathFromIDList(lRslt, CurrentSelection)
        strReturn = Left(CurrentSelection, InStr(CurrentSelection, vbNullChar) - 1)
        szDisplay = Left$(BI.pszDisplayName, InStr(BI.pszDisplayName, vbNullChar) - 1)
    End If

    FolderBrowse = strReturn

    CoTaskMemFree (lRslt)

End Function

Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    
    On Error Resume Next
    
    Dim retVal As Long

    Select Case uMsg
        Case BFFM_INITIALIZED
            If StartFolder > "" Then Call SendMessage(hWnd, BFFM_SETSELECTION, 0, ByVal StartFolder)
        
        Case BFFM_SELCHANGED
            retVal = SHGetPathFromIDList(lParam, CurrentSelection)
            If retVal <> 0 Then
                Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, ByVal CurrentSelection)
            End If

            If SpecialFolder = 4 Then Call SendMessage(hWnd, BFFM_ENABLEOK, 0, ByVal True)
            If Not OKEnable Then Call SendMessage(hWnd, BFFM_ENABLEOK, 0, ByVal OKEnable)

            CoTaskMemFree (retVal)

        Case BFFM_VALIDATEFAILED
            If MsgBox("The Path You Typed Does Not Exist!" & vbCrLf _
                    & "Would you like to create it?", vbYesNo Or vbQuestion) = vbYes Then
                szDisplay = GetText(hWnd)
                If szDisplay > "" Then
                    MakeDirectory (szDisplay)
                    Call SendMessage(hWnd, BFFM_SETSELECTION, 0, ByVal szDisplay)
                    BrowseCallbackProc = 1
                    Exit Function
                End If
            Else
                BrowseCallbackProc = 1
                Exit Function
            End If
    End Select

    BrowseCallbackProc = 0

End Function

Public Function DummyFunction(ByVal lParam As Long) As Long
    DummyFunction = lParam
End Function

