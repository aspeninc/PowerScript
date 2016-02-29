' ASPEN PowerScript Sample Program
'
' EXPORT.BAS
'
'
' Demo the ExportNetwork function
'
Sub main
  dim Options(10) As long
  
  ' OLR file
  OLRFile$ = "Sample30.olr"
  ExportFile$ = "export"
  
  If 0 = LoadBinary( OLRFile$ ) Then 
    Print "Error opening OLR file"
    Stop
  End If

  
  'ASPEN Format
  Options(1) = 0 ' 0-Entire network;1-Area;2-Zone
  Options(2) = 1 ' Area/Zone #
  Options(3) = 1 ' Inclue Tie
  
'  If ExportNetwork( ExportFile, Options ) Then
'    Print "ASPEN format export OK"
'  Else
'    Print "ASPEN format export Not OK"
'  End If

  Options(4) = 32		' PTI version 23-32
  Options(5) = 18000	' First fictitious bus number
  Options(6) = 15001    ' First bus number

  If ExportNetworkPSSE( ExportFile, ExportFile, Options ) Then
    Print "PSSE format export OK"
  Else
    Print "PSSE format export Not OK"
  End If
    
End Sub
