' ASPEN PowerScript sample program
' GETPFBUS.BAS
'
' Demonstrate how to access PF result at a bus
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

  ' Prepare output file
  Open "\0tmp\1.out" For Output As 1

  ' Get picked bus handle
  If GetEquipment( TC_PICKED, nBusHnd& ) = 0 Then
   Print "Must select a bus"
   Exit Sub
  End If

  busName$ = FullBusName( nBusHnd& )
  'Get voltagge at the bus
  If GetPFVoltage( nBusHnd&, vdVal1, vdVal2, ST_PU ) = 0 Then GoTo HasError
  ' Show it
  Print _ 
        "Voltage at bus: "; busName$; ": "; Chr(13); _
   "V1 = "; Format( vdVal1(1), "#0.00"); "@"; Format( vdVal2(1), "#0.0")
  ' Print it to file
  Print #1, _
        "Voltage at bus: "; busName$; ": "; Chr(13); _
   "V1 = "; Format( vdVal1(1), "#0.00"); "@"; Format( vdVal2(1), "#0.0")

  ' GenUnits
  nDevHnd = 0
  While GetBusEquipment( nBusHnd, TC_GENUNIT, nDevHnd ) > 0
    If GetData( nDevHnd, GU_sID, DevName$ ) = 0 Then GoTo HasError
    ' Get current from Gen
    If GetPFCurrent( nDevHnd&, vdVal1, vdVal2, 1 ) = 0 Then GoTo HasError
    ' Print it to file
    Print #1, _
          "Current from GenUnit ", DevName$, ":"; Chr(13); _
     "I1 = "; Format( vdVal1(1), "#0.0"); "@"; Format( vdVal2(1), "#0.0")
    ' Show it
    Print _
          "Current from GenUnit ", DevName$, ":"; Chr(13); _
     "I1 = "; Format( vdVal1(1), "#0.0"); "@"; Format( vdVal2(1), "#0.0")
    ' Get power from gen
    If GetFlow( nDevHnd&, vdVal1, vdVal2 ) = 0 Then GoTo HasError
    ' Print it to file
    Print #1, _
          "GenUnit", DevName$, " Power: "; Chr(13); _
     "P = "; Format( vdVal1(1), "#0.0"); " Q= "; Format( vdVal2(1), "#0.0")
    ' Show it
    Print _
          "GenUnit", DevName$, " Power: "; Chr(13); _
     "P = "; Format( vdVal1(1), "#0.0"); " Q= "; Format( vdVal2(1), "#0.0")
  Wend

  ' GenUnits
  nDevHnd = 0
  While GetBusEquipment( nBusHnd, TC_LOADUNIT, nDevHnd ) > 0
    If GetData( nDevHnd, LU_sID, DevName$ ) = 0 Then GoTo HasError
    ' Get current from Gen
    If GetPFCurrent( nDevHnd&, vdVal1, vdVal2, 1 ) = 0 Then GoTo HasError
    ' Print it to file
    Print #1, _
          "Current from LoadUnit ", DevName$, ":"; Chr(13); _
     "I1 = "; Format( vdVal1(1), "#0.0"); "@"; Format( vdVal2(1), "#0.0")
    ' Show it
    Print _
          "Current from LoadUnit ", DevName$, ":"; Chr(13); _
     "I1 = "; Format( vdVal1(1), "#0.0"); "@"; Format( vdVal2(1), "#0.0")
    ' Get power from gen
    If GetFlow( nDevHnd&, vdVal1, vdVal2 ) = 0 Then GoTo HasError
    ' Print it to file
    Print #1, _
          "LoadUnit", DevName$, " Power: "; Chr(13); _
     "P = "; Format( vdVal1(1), "#0.0"); " Q= "; Format( vdVal2(1), "#0.0")
    ' Show it
    Print _
          "LoadUnit", DevName$, " Power: "; Chr(13); _
     "P = "; Format( vdVal1(1), "#0.0"); " Q= "; Format( vdVal2(1), "#0.0")
  Wend

  Close 1
  Print "Output written to \0tmp\1.out"
  Stop
  HasError:
  Print "Error: ", ErrorString( )
  Close 
  Stop
End Sub