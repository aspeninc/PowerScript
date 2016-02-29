' ASPEN PowerScript Sample Program
'
' BOUNDARY.BAS
'
'
' Demonstrate the BoundaryEquivalent function
'
Sub main
  dim BusList(40) As long
  dim Options(5) As double
  
  ' OLR file
  OLRFile$ = "Sample30.olr"
  EqOLRFile$ = "Equivalent.olr"
  
  If 0 = LoadBinary( OLRFile$ ) Then 
    Print "Error opening OLR file"
    Stop
  End If
  
  nBusHnd& = 0
  nCount = 0
  While NextBusByName( nBusHnd& ) > 0
    Call GetData( nBusHnd&, BUS_dKVnorminal, dVal1# )
    If dVal1# > 100 Then
      nCount = nCount + 1
      BusList(nCount) = nBusHnd
    End If
  Wend
  BusList(nCount+1) = -1
  Options(1) = 99
  Options(2) = 1
  Options(3) = 0
  If BoundaryEquivalent( EqOLRFile, BusList, Options ) Then
    Print "OK"
  Else
    Print "Not OK"
  End If

  If 0 = LoadBinary( OLRFile$ ) Then 
    Print "Error opening OLR file"
    Stop
  End If
  
  If 0 = LoadBinary( OLRFile$ ) Then 
    Print "Error opening OLR file"
    Stop
  End If
  EqOLRFile$ = "c:\Documents and Settings\thanh\My Documents\Aspen\TestData\PowerScript\Equivalent2.olr"
  nBusHnd& = 0
  nCount = 0
  While NextBusByName( nBusHnd& ) > 0
    Call GetData( nBusHnd&, BUS_dKVnorminal, dVal1# )
    If dVal1# < 100 Then
      nCount = nCount + 1
      BusList(nCount) = nBusHnd
    End If
  Wend
  BusList(nCount+1) = -1
  Options(1) = 99
  Options(2) = 1
  Options(3) = 0
  If BoundaryEquivalent( EqOLRFile, BusList, Options ) Then
    Print "OK"
  Else
    Print "Not OK"
  End If

    
End Sub
