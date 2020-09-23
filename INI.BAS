Attribute VB_Name = "mod_ini"
Declare Function ReadINI Lib "kernel32" Alias "GetPrivateProfileStringA" _
   (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
   ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
   ByVal lpFileName As String) As Long
Public Declare Function WriteINI Lib "kernel32" Alias "WritePrivateProfileStringA" _
   (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
   ByVal lpFileName As String) As Long

Public sIniFile As String

Public Function IniDaj(Sekcija$, Kljuc$, Optional DefaultVrijednost$ = "")
   Dim x%, temp$
   temp = Space(255)
   x = ReadINI(Sekcija, Kljuc, DefaultVrijednost, temp, Len(temp), sIniFile) 'Text1(n).Text
   temp = Trim(temp)
   If Right(temp, 1) = vbNullChar Then temp = Left(temp, Len(temp) - 1)
   IniPisi Sekcija, Kljuc, Trim(temp)
   IniDaj = Trim(temp)
End Function

Public Sub IniPisi(Sekcija$, Kljuc$, Vrijednost$)
   Dim x%
   x = WriteINI(Sekcija, Kljuc, Vrijednost, sIniFile)
End Sub
Function CheckDirSep(sDir)
   If Len(Trim(sDir)) = 0 Then Exit Function
   If Right(sDir, 1) = "\" Then
      CheckDirSep = sDir
   Else
      CheckDirSep = sDir & "\"
   End If
End Function

