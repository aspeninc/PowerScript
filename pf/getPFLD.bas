' ASPEN PowerScript sample program
' GETPFGEN.BAS
'
' Get power and current of a generator
'
' Demonstrate how to access PF result from a PowerScript program
'
' PowerScript functions used
'  GetEquipment()
'  GetData()
'  GetPFVoltage()
'  GetPFCurrent()
'  GetFlow()
Sub main()

  'Variable declarations
  Dim vdVal1(12) As Double, vdVal2(12) As Double
  Dim vnShowRelay(4) As Long

  ' Prepare output file
  Open "\0tmp\1.out" For Output As 1

  ' Get picked load handle number
  If GetEquipment( TC_PICKED, nBusHnd& ) = 0 Then
   Print "Must select a bus with load"
   Exit Sub
  End If
  If EquipmentType( nBusHnd ) <> TC_BUS Then
   Print "Must select a bus with load"
   Exit Sub
  End If

  If GetBusEquipment( nBusHnd, TC_GENUNIT, nDevHnd& ) <= 0 Then
   Print "Must select a bus with generator"
   Exit Sub
  End If


  busName$ = FullBusName( nBusHnd& )
  'Get voltagge at the end bus
  If GetPFVoltage( nBusHnd&, vdVal1, vdVal2, ST_PU ) = 0 Then GoTo HasError
  ' Show it
  Print _ 
        "Voltage at bus: "; busName$; ": "; Chr(13); _
   "V1 = "; Format( vdVal1(1), "#0.00"); "@"; Format( vdVal2(1), "#0.0")
  ' Print it to file
  Print #1, _
        "Voltage at bus: "; busName$; ": "; Chr(13); _
   "V1 = "; Format( vdVal1(1), "#0.00"); "@"; Format( vdVal2(1), "#0.0")

    ' Get current from Gen
  If GetPFCurrent( nDevHnd&, vdVal1, vdVal2, 1 ) = 0 Then GoTo HasError
  ' Print it to file
  Print #1, _
        "Current from Gen: "; Chr(13); _
   "I1 = "; Format( vdVal1(1), "#0.0"); "@"; Format( vdVal2(1), "#0.0")
  ' Show it
  Print _
        "Current from Gen: "; Chr(13); _
   "I1 = "; Format( vdVal1(1), "#0.0"); "@"; Format( vdVal2(1), "#0.0")
  ' Get power from gen
  If GetFlow( nDevHnd&, vdVal1, vdVal2 ) = 0 Then GoTo HasError
  ' Print it to file
  Print #1, _
        "Gen: "; Chr(13); _
   "P = "; Format( vdVal1(1), "#0.0"); " Q= "; Format( vdVal2(1), "#0.0")
  ' Show it
  Print _
        "Gen: "; Chr(13); _
   "P = "; Format( vdVal1(1), "#0.0"); " Q= "; Format( vdVal2(1), "#0.0")
  Close 1
  Print "Output written to \0tmp\1.out"
  Stop
  HasError:
  Print "Error: ", ErrorString( )
  Close 
  Stop
End Sub