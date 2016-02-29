' ASPEN PowerScript sample program
'
' DOSTEPPEDEVENT.BAS
'
' Demo the DoSpeppedEvent Command
'
Sub main()

  'Variable declarations
  Dim vdFltOpt(15) As Double, vnDevOpt(10) As long
  

  ' Get picked bus handle
  If (GetEquipment( TC_PICKED, nBusHnd& ) = 0) Or _
     (EquipmentType( nBusHnd& ) <> TC_BUS And EquipmentType( nBusHnd& ) <> TC_RLYGROUP ) Then
   Print "Please select a bus or relay group"
   Exit Sub
  End If

  vdFltOpt(1) = 5	'nFltConn
                    '    1=3LG,
                    '    2=2LG_BC,3=2LG_CA,4=2LG_AB
                    '    5=1LG_A,6=1LG_B,7=1LG_C,
                    '    8=LL_BC,9=LL_CA,10=LL_AB
  vdFltOpt(2) = 10	'Intermediate percent between 0.01-99.99
                    '   0 For Close-in
  vdFltOpt(3) = 0	'Fault resistance
  vdFltOpt(4) = 0	'Fault reactance
  vdFltOpt(4+1) = 8	'LL BC
  vdFltOpt(4+2) = 0.5	'time
  vdFltOpt(4+3) = 0.5	'Fault resistance
  vdFltOpt(4+4) = 0	'Fault reactance
  vdFltOpt(4+5) = 0	'Zero additional event
  vnDevOpt(1) = 1	'Consider OCGnd
  vnDevOpt(2) = 1	'Consider OCPh
  vnDevOpt(3) = 1	'Consider DSGnd
  vnDevOpt(4) = 1	'Consider DSPh
  
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
  Next
  
  Stop
  HasError:
  Print "Error: ", ErrorString( )
  Stop
End Sub