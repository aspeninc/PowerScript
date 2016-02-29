' ASPEN PowerScript Sample Program
'
' SAVEDATAFILE.BAS
'
' Report total line impedance and length.
' Lines with tap buses are handled correctly
Sub main
  sFile$ = ""
  If 0 = SaveDataFile (sFile, 1+2) Then
    Print ErrorString()
  Else
    Print "File Saved"
  End If
End Sub
