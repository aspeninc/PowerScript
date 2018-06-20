' ASPEN PowerScrip sample program
'
' FileDlg.BAS
'
' Show file dialog boxes
'
' Version: 1.0
'
Sub main
' sPath$ = FileOpenDialog( "", "Change files (*.chf)|*.chf|OneLiner files (*.olr)|*.olr||", 4 )
' sPath$ = FileSaveDialog( "", "", ".txt", 2+16 )
 sPath$ = FolderSelectDialog( "My Dialog", "c:\" )
 Print sPath
End Sub
