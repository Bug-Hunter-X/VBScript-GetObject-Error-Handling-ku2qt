Function GetObject(path)
  On Error Resume Next
  Set obj = GetObject(path)
  If Err.Number <> 0 Then
    Err.Clear
    Set obj = Nothing
  End If
  Set GetObject = obj
End Function

' Example usage
Set objExcel = GetObject("C:\\path\\to\\your\\excel\\file.xls")
If objExcel Is Nothing Then
  MsgBox "Could not open Excel file."
Else
  ' Work with the Excel object
  objExcel.Application.Quit
  Set objExcel = Nothing
End If