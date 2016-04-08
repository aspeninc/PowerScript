' ASPEN Sample PowerScript
' SEA.BAS
'
' Demo of the DoSteppedEvent() function
'
Sub main
 If GetEquipment( TC_PICKED, nBusHnd ) = 0 Or EquipmentType( nBusHnd ) <> TC_BUS Then
    Print "No bus is selected"
    Stop
  End If
  dim vdFltOpt(10) As double
  dim vnDevOpt(10) As long
  vdFltOpt(1) = 1    '3LG
  vdFltOpt(2) = 0    'Intermediate percent between 0.01-99.99
  vdFltOpt(3) = 0    'Fault resistance
  vdFltOpt(4) = 0    'Fault reactance
  vdFltOpt(5) = 0    'Zero or nFltconn of additional event
  vnDevOpt(1) = 1    'Consider OCGnd
  vnDevOpt(2) = 1    'Consider OCPh
  vnDevOpt(3) = 1    'Consider DSGnd
  vnDevOpt(4) = 1    'Consider DSPh
  nTiers& = 5

  If 0 = DoSteppedEvent( nBusHnd, vdFltOpt, vnDevOpt, nTiers ) Then GoTo HasError
  ' Call GetSteppedEvent with 0 to get total number of events simulated
  nSteps = GetSteppedEvent( 0, dTime#, dCurrent#, nUserEvwent&, sEventDesc$, sFaultDest$ )
  Print "Stepped-event simulation completed successfully with ", nSteps-1, " events"
  For ii = 1 to nSteps
    Call GetSteppedEvent( ii, dTime#, dCurrent#, nUserEvwent&, sEventDesc$, sFaultDest$ )
    Print "Fault: ", sFaultDest$
    Print sEventDesc$
    Print "Time = ", dTime, " Current= ", dCurrent
    ''''''''''''''''''''''''''''''''''''''
    '' ADD LOGIC TO GET LINE CURRENT HERE
    ''''''''''''''''''''''''''''''''''''''
  Next

  Stop
  HasError:
  Print "Error: ", ErrorString( )
  Stop
End Sub