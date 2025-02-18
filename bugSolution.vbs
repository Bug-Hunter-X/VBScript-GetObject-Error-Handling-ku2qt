Function GetObjectSafe(path)
  Dim obj, fso, fileExists
  Set fso = CreateObject("Scripting.FileSystemObject")
  fileExists = fso.FileExists(path)
  Set fso = Nothing

  If fileExists Then
    On Error Resume Next
    Set obj = GetObject(path)
    If Err.Number <> 0 Then
      Err.Clear
      MsgBox "Error accessing file: " & Err.Description, vbCritical
      Set obj = Nothing
    End If
    On Error GoTo 0
  Else
    MsgBox "File not found: " & path, vbCritical
  End If

  Set GetObjectSafe = obj
End Function

' Example usage
Set objExcel = GetObjectSafe("C:\\path\\to\\your\\excel\\file.xls")
If objExcel Is Nothing Then
  ' Handle the case where the file could not be opened
Else
  ' Work with the Excel object
  objExcel.Application.Quit
  Set objExcel = Nothing
End If