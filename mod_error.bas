Attribute VB_Name = "mod_error"
Option Explicit

Sub ErrorHandler(sProcedure As String, sModule As String)
Dim tmpErrMsg

If sProcedure = "" Then sProcedure = "Unspecified"
If sModule = "" Then sModule = "Unspecified"

tmpErrMsg = MsgBox("Code: " & Err & vbCr & "Description: " & Error & vbCr & vbCr & "Procedure: " & sProcedure & vbCr & "Module: " & sModule & vbCr & vbCr & "Note: error saved in LOG file: " & App.Title & "_err.log", vbOKOnly + vbCritical, "Error")

Open GlobalAppPath & App.Title & "_err.log" For Append As #2

Print #2, ">>> " & Format(Date, "dd.mm.yyyy") & " ** " & Format(Time, "hh:mm") & " ** Code: " & Err & " ** Description: " & Error & " ** Procedure: " & sProcedure & " ** Module: " & sModule

Close #2

End Sub
