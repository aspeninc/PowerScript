' ASPEN PowerScript sample program
'
' ARCFLASH.BAS
'
' Demonstrate the Voltage Sag Analysis Command
'
' PowerScript functions used
Sub main()

  'Variable declarations
  Dim vdOption(10) As Double, vdResult(15) As Double
  
  ' Output file
'  CSVFile$ = "c:\0tmp\vs.csv"

  ' Get picked bus handle
  If GetEquipment( TC_PICKED, nBusHnd& ) = 0 Then
   Print "Must select a bus"
   Exit Sub
  End If

  vdOption(1)  = 0	'0-Switchgear; 1-Cable; 2- open air
  vdOption(2) = 1	'0-Ungrounded;1-Grounded
  vdOption(3) = 0	'0-No enclosure;1-Enclosed
  vdOption(4) = 153	'Conductor gap in mm
  vdOption(5) = 36	'Working distance in inches
  
  vdOption(4) = 153	'Conductor gap in mm
  vdOption(4) = 153	'Conductor gap in mm
  
  vdOption(1) = 0.6	'Sag threshold
  vdOption(2) = 0.0	'Line percent
  vdOption(3) = 1	'Ouput all
  vdOption(4) = 0.0	'Zground.imag
  vdOption(5) = 0.0	'Zground.real
  
  If 0 = DoVS( nBusHnd, vnFltConn, vdOption, CSVFile ) Then GoTo HasError
  
  
  Print "Voltage sag simulation complete. Output is in " + CSVFile
  Stop
  HasError:
  Print "Error: ", ErrorString( )
  Stop
End Sub