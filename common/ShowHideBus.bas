' ASPEN PowerScript Sample Program
'
' ShowHideBus.BAS
'
' Demonstrate steps to Show and Hide bus in PowerScript program
'
Sub main
    
  ' OLR file
  OLRFile$ = "Sample30.olr"
  
  If 0 = LoadDataFile( OLRFile$ ) Then 
    Print "Error opening OLR file"
    Stop
  End If

  nCount = 0
  nBusHnd& = 0
  While NextBusByName(nBusHnd&) > 0
    Call GetData( nBusHnd&, BUS_nArea, nArea& )
    If nArea = 1 Then
      If 0 = SetData( nBusHnd&, BUS_nVisible, 1 ) Then GoTo hasError
      If 0 = PostData(nBusHnd&) Then GoTo hasError
      nCount = nCount + 1
    End If
  Wend
  
Break:
  Print nCount, " buses updated"
  Stop
HasError:
  Print "Error: ", ErrorString( )
End Sub