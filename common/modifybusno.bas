' ASPEN PowerScript sample program
'
' MODIFYBUSNO.BAS
'
' Network database access demo: Reset bus number to zero.
' 
'
Sub main()
  Count   = 0
  BusHnd& = 0	'This will make next cmd to seek the first bus
  While NextBusByName( BusHnd& ) > 0
   Call GetData( BusHnd&, BUS_nNumber, BusNo& )
   If BusNo > 1000000 Then
    BusNo = 0
    If SetData(   BusHnd&, BUS_nNumber, BusNo ) = 0 Then GoTo HasError
    If PostData( BusHnd& ) > 0 Then Count = Count + 1
   End If
  Wend
  Print Count, " buses had been modified"
  Exit Sub
HasError:
  Print "Error: ", ErrorString( )
End Sub