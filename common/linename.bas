' ASPEN PowerScrip sample program
' LINENAME.BAS
'
' Create line name from end-bus IDs
'
' Demonstrate PowerScript can update data in the nework database
'
' PowerScript functions used:
'   GetEquipment()
'   GetData()
'   SetData()
'   PostData()
'
Sub main()
LineHnd = 0
Counts  = 0
While GetEquipment( TC_LINE, LineHnd ) > 0
  If GetData( LineHnd, LN_nBus1Hnd, nBusHnd ) = 0 Then GoTo HasError
  If GetData( nBusHnd, BUS_dKVNorminal, dKV ) = 0 Then GoTo HasError
  If dKV > 100 Then 'Do 100kV and above
    If GetData( nBusHnd, BUS_sName, sBusName1$ ) = 0 Then GoTo HasError
    If GetData( LineHnd, LN_nBus2Hnd, nBusHnd ) = 0 Then GoTo HasError
    If GetData( nBusHnd, BUS_sName, sBusName2$ ) = 0 Then GoTo HasError
    sLineName$ = Left(sBusName1$,3) & " : " & Left(sBusName2$,3)
    If SetData( LineHnd, LN_sName, sLineName$ ) = 0 Then GoTo HasError
    If PostData( LineHnd ) = 0 Then GoTo HasError
    Counts = Counts + 1
  End If
Wend  'Each line
Print Counts; " line names created"
Exit Sub
HasError:
Print "Error: " , ErrorString()
End Sub