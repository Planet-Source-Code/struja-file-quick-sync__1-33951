Attribute VB_Name = "mod_main"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public TemplateNo As Integer
Public SelectedFolder As String

Public SyncMode As String

Public DbSys As Database

Public GlobalAppPath As String

Public Q As String
Public Qtmp As String

Public tmp
Public qwe
Public wer

Public AllOk As Boolean

Public NewFlag As Boolean
Public NewOk As Boolean

Public tmrTemplate As String
Public tmrDate As String
Public tmrTime As String
Public tmrEnabled As Boolean
Public tmrDaily As Boolean

Public FlagShowAllInReport As Boolean

Public tmrDailyStarted As String


Sub CenterForm(f As Form)
   f.Move Screen.Width / 2 - f.Width / 2, Screen.Height / 2 - (f.Height / 2) - 1400

End Sub

Sub CenterForm_global(f As Form)
   f.Move Screen.Width / 2 - f.Width / 2, Screen.Height / 2 - f.Height / 2

End Sub

Public Function FileExist(sFileName$, Optional sPath$) As Boolean
   FileExist = False
   'If sPath = "" Then sPath = App.Path
   Dim sPutDoSys$
   
   sPutDoSys = Dir(sPath & sFileName, vbNormal)
   If sPutDoSys = "" Then Exit Function
   
   FileExist = True
End Function

Sub FormOnTop(f As Form)
   SetWindowPos f.hwnd, -1, 0, 0, 0, 0, 3
End Sub

Sub UnsetColor(contr As Control)
If ContactMode = "View" Then
   contr.BackColor = vbButtonFace
Else
   contr.BackColor = vbWhite
End If

   contr.ForeColor = vbWindowText
   DoEvents
End Sub

Sub SetColor(contr As Control)
   contr.BackColor = vbCyan
   contr.ForeColor = vbBlack
   If TypeOf contr Is TextBox Then
      contr.SelStart = 0: contr.SelLength = 10000
   End If
   DoEvents
End Sub

Sub UnsetColorGeneral(contr As Control)
   contr.BackColor = vbWhite
   contr.ForeColor = vbWindowText
   DoEvents
End Sub

Sub SetColorGeneral(contr As Control)
   contr.BackColor = vbCyan
   contr.ForeColor = vbBlack
   If TypeOf contr Is TextBox Then
      contr.SelStart = 0: contr.SelLength = 10000
   End If
   DoEvents
End Sub

Public Function NapraviDatumStruja(txt As Control) As String

If txt.Text = "" Then
    NapraviDatumStruja = ""
    Exit Function
End If

Dim tmpDate As String

tmpDate = Format(txt.Text, "dd.mm.yyyy")

If Len(tmpDate) <> 10 Then
    tmp = MsgBox("Incorrect Date format entered!", vbOKOnly + vbCritical, "Error")
    txt.SetFocus
End If

NapraviDatumStruja = tmpDate
End Function

Public Function NapraviDatum(txt As String) As String
On Error GoTo GreskaDatum

If txt = "" Then
    NapraviDatum = ""
    Exit Function
End If

Dim tmpDate As String

tmpDate = Format(txt, "dd.mm.yyyy")

NapraviDatum = tmpDate
Exit Function

GreskaDatum:
    NapraviDatum = ""

End Function

Public Function DirExists(Dpath As String) As Boolean
On Error GoTo DEerror

DirExists = False

If Dpath = "" Then Exit Function

ChDir Dpath
DirExists = True
Exit Function

DEerror:
DirExists = False

End Function

Sub KreirajPut(ByVal pth As String)
'Robert_kreira bilo koji path
   Dim s As String, n As Integer, i As Integer, t As String, p As String
   Dim disk As String
   On Error Resume Next
   disk = "": s = "": p = ""
   
   
   If Mid(pth, 2, 1) = ":" Then
      disk = Left(pth, 2)
      pth = Right(pth, Len(pth) - 2)
   Else
      disk = Left(CurDir, 2)
   End If
   If Left(pth, 1) = "\" Then disk = disk & "\": pth = Right(pth, Len(pth) - 1)
   If Right(pth, 1) <> "\" Then pth = pth & "\"
   If disk <> "" Then ChDir disk
   
   'pth = UCase(pth)
   
   For n = 1 To Len(pth)
      t = Mid(pth, n, 1)
      If t <> "\" Then
         s = s & t
      Else
         p = p & s & "\"
         MkDir Left(disk & p, Len(disk & p) - 1)
         s = ""
      End If
   Next n
End Sub

Function CheckText(c As Control, sTip$) As Boolean
   CheckText = False
   Select Case LCase(sTip)
      Case Is = "numeric"
         If Not IsNumeric(c) Or c = "" Then
            MsgBox "Unesite podatak!", vbExclamation, "Neispravan unos"
            c.SetFocus
            Exit Function
         End If
      Case Is = "text"
         If Len(Trim(c)) = 0 Then
            MsgBox "Unesite podatak!", vbExclamation, "Neispravan unos"
            c.SetFocus
            Exit Function
         End If
      Case Else
         MsgBox "Krivi tip podatka", vbExclamation, "Greška"
   End Select
   CheckText = True
End Function


Function NapraviDatumCRP(dat As String) As String
  Dim LokDat, tmp$
  
  LokDat = DateValue(dat)
  tmp = "(" & Format$(LokDat, "yyyy") & ", "
  tmp = tmp & Format$(LokDat, "mm") & ", "
  tmp = tmp & Format$(LokDat, "dd") & ")"
  NapraviDatumCRP = tmp
End Function

Function NapraviDatum2(dat As String) As String
  Dim LokDat, tmp$
  Dim Kako As String
  If dat = "" Then Exit Function
  Kako = Left$(dat, 1)
  If Kako = "<" Or Kako = "=" Or Kako = ">" Then
    dat = Right$(dat, Len(dat) - 1)
  ElseIf Kako = "#" Then
     
     tmp = Mid(dat, 5, 2) & "."
     tmp = tmp & Mid(dat, 2, 2) & "."
     tmp = tmp & Mid(dat, 8, 4)
     NapraviDatum2 = tmp
     Exit Function
  
  End If
  
  LokDat = DateValue(dat)
  tmp = "#" & Format$(LokDat, "MM") & "/"
  tmp = tmp & Format$(LokDat, "dd") & "/"
  tmp = tmp & Format$(LokDat, "yyyy") & "#"
  NapraviDatum2 = tmp
End Function

Public Function DajPath(sStaza$) As String
   DajPath = ""
   Dim iLastBSlash%, i%
   For i = 1 To Len(sStaza)
      If Mid(sStaza, i, 1) = "\" Then iLastBSlash = i
   Next i
   DajPath = Left(sStaza, iLastBSlash)
End Function

Public Function NapraviVrijemeStruja(txt As Control) As String

If txt.Text = "" Then
    NapraviVrijemeStruja = ""
    Exit Function
End If

Dim tmpTime As String

tmpTime = Format(txt.Text, "hh:mm")

If Len(tmpTime) <> 5 Then
    tmp = MsgBox("Incorrect Time format entered!", vbOKOnly + vbCritical, "Error")
    txt.SetFocus
End If

NapraviVrijemeStruja = tmpTime
End Function

Public Function IsTemplateExists(tmpTname As String) As Boolean

IsTemplateExists = False

Dim trs As Recordset
Dim tcnt

Set trs = DbSys.OpenRecordset("Templates")

If trs.RecordCount = 0 Then
    IsTemplateExists = False
    Exit Function
End If

trs.MoveLast
trs.MoveFirst

For tcnt = 1 To trs.RecordCount
    If LCase("" & trs!Name) = LCase(tmpTname) Then
        IsTemplateExists = True
        Exit Function
    End If
    trs.MoveNext
Next tcnt

End Function

