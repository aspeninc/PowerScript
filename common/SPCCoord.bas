' ASPEN PowerScript Sample Program
'
' SPCoord.BAS
'
' Demonstrate the use of State Plane Coordinates
'
' Version 1.0
'
Sub main
    
  ' OLR file
  OLRFile$ = "Sample30.olr"
  SPCFile$ = "SPC.txt"
  
  If 0 = LoadBinary( OLRFile$ ) Then 
    Print "Error opening OLR file"
    Stop
  End If

  Open SPCFile$ For Input As 1
  
  nCount = 0
  Do While Not EOF(1)
    Line Input #1, aLine$
    If 1 = ParseLine(aLine$, nBusHnd&, dX#, dY#) Then
      If 0 = SetData( nBusHnd&, BUS_dSPCx, dX ) Then GoTo hasError
      If 0 = SetData( nBusHnd&, BUS_dSPCy, dY ) Then GoTo hasError
      If 0 = PostData(nBusHnd&) Then GoTo hasError
      nCount = nCount + 1
    End If
  Loop
  
Break:
  Print nCount, " buses updated"
  Stop
HasError:
  Print "Error: ", ErrorString( )
End Sub

Function ParseLine( ByRef aLine$, ByRef nBusHnd&, ByRef dX#, ByRef dY# ) As long
  ParseLine = 0
  ' Bus name
  nStart& = 1
  nEnd& = InStr( nStart&, aLine$, "," )
  If nEnd = 0 Then exit Function
  Token$ = Mid( aLine$, nStart&, nEnd&-nStart& )
  Token$ = Trim(Token$)
  If Token$ = "" Then exit Function
  BName$ = Mid(Token$,2,len(Token$)-2)
  ' kV nominal
  nStart& = nEnd& + 1
  nEnd& = InStr( nStart&, aLine$, "," )
  If nEnd = 0 Then exit Function
  Token$ = Mid( aLine$, nStart&, nEnd&-nStart& )
  Token$ = Trim(Token$)
  If Token$ = "" Then exit Function
  BKV# = Val(Token$)
  If BKV# = 0.0 Then exit Function
  If 1 <> FindBusByName( BName$, BKV#, nBusHnd& ) Then exit Function
  ' X
  nStart& = nEnd& + 1
  nEnd& = InStr( nStart&, aLine$, "," )
  If nEnd = 0 Then exit Function
  Token$ = Mid( aLine$, nStart&, nEnd&-nStart& )
  Token$ = Trim(Token$)
  If Token$ = "" Then exit Function
  dX# = Val(Token$)
  ' Y
  nStart& = nEnd& + 1
  nEnd&  = Len(aLine$) + 1
  Token$ = Mid( aLine$, nStart&, nEnd&-nStart& )
  Token$ = Trim(Token$)
  If Token$ = "" Then exit Function
  dY# = Val(Token$)
  
  ParseLine = 1
End Function
