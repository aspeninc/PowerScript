' FILESAVEDLG.BAS
'
' ASPEN sample script
'
' Version 1.0
'
Sub main
 sOlrPathName$ = GetOlrFileName()
 sOlrPath$ = ExtractFilePath(sOlrPathName)
 Print sOlrPath
 sPath = FileSaveDialog(sOlrPath, "Excel CSV (*.csv)|*.csv|", ".csv", 2+16 )
End Sub

Function ExtractFilePath( sFullPathName$ ) As String
 ExtractFilePath = ""
 nLen& = Len(sFullPathName)
 Do
   If nLen <= 1 Then GoTo breakWhile
   nLen = nLen - 1
   aChar$ = Mid(sFullPathName, nLen, 1)
 Loop While aChar <> "\"
 breakWhile:
 ExtractFilePath$ = Left(sFullPathName,nLen)
End Function
