' ASPEN PowerScript sample program
' VSag.BAS
'
' Demonstrate the Voltage Save Analysis Command
'
' Version 1.0
' Category: OneLiner
'
' PowerScript functions used
Sub main()

  'Variable declarations
  Dim vdOption(7) As Double, vdMag(4) As Double
  Dim vnFltConn(4) As Long

  ' Output file
  CSVFile$ = "c:\0tmp\vs.csv"

  ' Get picked bus handle
  If GetEquipment( TC_PICKED, nBusHnd& ) = 0 Then
   Print "Must select a bus"
   Exit Sub
  End If

  vnFltConn(1) = 0	'1LG
  vnFltConn(2) = 0	'2LG
  vnFltConn(3) = 1	'3PH
  vnFltConn(4) = 0	'LL
  
  vdOption(1) = 0.5	'Sag threshold
  vdOption(2) = 50.0	'Line percent
  vdOption(3) = 1	'Ouput all
  vdOption(4) = 0.0	'Zground.imag
  vdOption(5) = 0.0	'Zground.real
'  vdOption(6) = 1.0	'Stepped event
'  vdOption(7) = 5.0	'Stepped event extent

  
  If 0 = DoVS( nBusHnd, vnFltConn, vdOption, CSVFile ) Then GoTo HasError
  
  
  Print "Voltage sag simulation complete. Output is in " + CSVFile
  Stop
  HasError:
  Print "Error: ", ErrorString( )
  Stop
End Sub
