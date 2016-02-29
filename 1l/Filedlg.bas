' ASPEN PowerScrip sample program
'
' FileDlg.BAS
'
Sub main
' sPath$ = FileOpenDialog( "", "Change files (*.chf)|*.chf|OneLiner files (*.olr)|*.olr||", 4 )
' sPath$ = FileSaveDialog( "", "", ".txt", 2+16 )
 sPath$ = FolderSelectDialog( "My Dialog", "c:\" )
 Print sPath
End Sub
