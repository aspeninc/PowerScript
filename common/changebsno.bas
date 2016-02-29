' ASPEN PowerScript sample program
'
' CHANGEBSNO.BAS
'
' Change bus data.
' Demonstrate how to modify network database through a change file
' 
' PowerScript functions called:
'   NextBusByName()
'   GetData()
'   SetData()
'   PostData()
'
Sub main()

  sFileName = "c:\number.chf"		'TODO: make sure this path is writeable
  
  Open sFileName For output As 1
  Print #1, "[ONELINER AND POWER FLOW CHANGE FILE]"
  Print #1, "delimiter='"
  Print #1, "ver= 2006 'E'"
  Print #1, "[MODIFY BUS DATA]"
  
  Count   = 0
  BusHnd& = 0	'This will make the cmd on the next line to seek the first bus
  While NextBusByName( BusHnd& ) > 0
    ' Bus name
    If GetData( BusHnd&, BUS_sName, sString$ ) = 0 Then GoTo HasError
    sLine$ = "'" + sString + "' "
    ' Bus kV
    If GetData( BusHnd&, BUS_dkvNominal, dNumber# ) = 0 Then GoTo HasError
    sLine$ = sLine$ + Format(dNumber, "0.00" ) + "= "
    
    ' Bus Number
    If GetData( BusHnd&, BUS_nNumber, nNumber& ) = 0 Then GoTo HasError
    If nNumber > 0 Then nNumber = nNumber + 1000	'TODO: complete this bus number modification logic
    sLine$ = sLine$ + Str(nNumber) + " "
    
    ' Bus Area
    If GetData( BusHnd&, BUS_nArea, nNumber& ) = 0 Then GoTo HasError
    sLine$ = sLine$ + Str(nNumber) + " "
    ' Bus Zone
    If GetData( BusHnd&, BUS_nZone, nNumber& ) = 0 Then GoTo HasError
    sLine$ = sLine$ + Str(nNumber) + " "
    ' Bus location
    If GetData( BusHnd&, BUS_sLocation, sString$ ) = 0 Then GoTo HasError
    sLine$ = sLine$ + "'" + sString$ + ", "
    ' Tap Bus flag
    If GetData( BusHnd&, BUS_nTapBus, nNumber& ) = 0 Then GoTo HasError
    sLine$ = sLine$ + Str(nNumber) + " "
    ' Power flow solution, state plane coords and mid-point flag
    sLine$ = sLine$ + " 1 0 0 0 0"
    ' Substation group
    If GetData( BusHnd&, BUS_nSubGroup, nNumber& ) = 0 Then GoTo HasError
    sLine$ = sLine$ + Str(nNumber) + " "
    ' Bus comments
    If GetData( BusHnd&, BUS_sComment, sString$ ) = 0 Then GoTo HasError
    sLine$ = sLine$ + "'" + sString$ + "'"
    Print #1, sLine
    Count = Count + 1
  Wend
  Close 1
  Print Count, " buses have been modified. The change file is at " + sFileName
  Exit Sub
  
HasError:
  Print "Error: ", ErrorString( )
End Sub