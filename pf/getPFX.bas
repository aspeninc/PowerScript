' ASPEN PowerScript sample program
' GETPFX.BAS
'
' Get power and current of a load
' Demonstrate how to access PF result from a PowerScript program
'
' Version 1.0
' Category: Power Flow
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
  If GetEquipment( TC_PICKED, nDevHnd& ) = 0 Then
   Print "Must select a line"
   Exit Sub
  End If
  If EquipmentType( nDevHnd ) <> TC_LOAD Then
   Print "Must select a load"
   Exit Sub
  End If
  If GetData( nDevHnd&, XR_nBus1Hnd, nBusHnd& ) = 0 Then GoTo HasError
  sVal2$ = FullBusname( nBusHnd& )
  If GetData( nDevHnd&, XR_nBus2Hnd, nBusHnd& ) = 0 Then GoTo HasError
  sVal3$ = FullBusName( nBusHnd& )
  'Get voltagge at the end bus
  If GetPFVoltage( nDevHnd&, vdVal1, vdVal2, ST_PU ) = 0 Then GoTo HasError
  ' Show it
  Print "Voltage on xfmr: "; sVal2$ & "-"; sVal3$ & " ID= "; sVal1$; ": "; Chr(13); _
   "V1 = "; Format( vdVal1(1), "#0.00"); "@"; Format( vdVal2(1), "#0.0"); _
        "; V2 = "; Format( vdVal1(2), "#0.00"); "@"; Format( vdVal2(2), "#0.0")
  ' Print it to file
  Print #1, _
        "Voltage on xfmr: "; sVal2$ & "-"; sVal3$ & " ID= "; sVal1$; ": "; Chr(13); _
   "V1 = "; Format( vdVal1(1), "#0.00"); "@"; Format( vdVal2(1), "#0.0"); _
        "; V2 = "; Format( vdVal1(2), "#0.00"); "@"; Format( vdVal2(2), "#0.0")
  ' Get current into end buses
  If GetPFCurrent( nDevHnd&, vdVal1, vdVal2, 1 ) = 0 Then GoTo HasError
  ' Print it to file
  Print #1, _
        "Current on xfmr: "; sVal2$ & "-"; sVal3$ & " ID= "; sVal1$; ": "; Chr(13); _
   "I1 = "; Format( vdVal1(1), "#0.0"); "@"; Format( vdVal2(1), "#0.0"); _
        "; I2 = "; Format( vdVal1(2), "#0.0"); "@"; Format( vdVal2(2), "#0.0")
  ' Show it
  Print "Current on xfmr: "; sVal2$ & "-"; sVal3$ & " ID= "; sVal1$; ": "; Chr(13); _
   "I1 = "; Format( vdVal1(1), "#0.0"); "@"; Format( vdVal2(1), "#0.0"); _
        "; I2 = "; Format( vdVal1(2), "#0.0"); "@"; Format( vdVal2(2), "#0.0")
  ' Get power into end buses
  If GetFlow( nDevHnd&, vdVal1, vdVal2 ) = 0 Then GoTo HasError
  ' Print it to file
  Print #1, _
        "Power on xfmr: "; sVal2$ & "-"; sVal3$ & " ID= "; sVal1$; ": "; Chr(13); _
   "P1 = "; Format( vdVal1(1), "#0.0"); " Q1= "; Format( vdVal2(1), "#0.0"); _
        "; P2 = "; Format( vdVal1(2), "#0.0"); " Q2 = "; Format( vdVal2(2), "#0.0")
  ' Show it
  Print "Power on xfmr: "; sVal2$ & "-"; sVal3$ & " ID= "; sVal1$; ": "; Chr(13); _
   "P1 = "; Format( vdVal1(1), "#0.0"); " Q1= "; Format( vdVal2(1), "#0.0"); _
        "; P2 = "; Format( vdVal1(2), "#0.0"); " Q2 = "; Format( vdVal2(2), "#0.0")
  Close 1
  Print "Output written to \0tmp\1.out"
  Stop
  HasError:
  Print "Error: ", ErrorString( )
  Close 
  Stop
End Sub