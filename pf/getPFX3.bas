' ASPEN PowerScript example
' GETPFX3.BAS
'
' Get voltage and flow on a 3-W transformer
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
  Open "c:\0tmp\1.out" For Output As 1

  ' Get picked line handle number
  If GetEquipment( TC_PICKED, nDevHnd& ) = 0 Then
   Print "Must select a line"
   Exit Sub
  End If
  If EquipmentType( nDevHnd ) <> TC_XFMR3 Then
   Print "Must select a 3-winding transformer"
   Exit Sub
  End If

  If GetData( nDevHnd&, X3_nBus1Hnd, nBusHnd& ) = 0 Then GoTo HasError
  sVal2$ = FullBusname( nBusHnd& )
  If GetData( nDevHnd&, X3_nBus2Hnd, nBusHnd& ) = 0 Then GoTo HasError
  sVal3$ = FullBusName( nBusHnd& )
  If GetData( nDevHnd&, X3_nBus3Hnd, nBusHnd& ) = 0 Then GoTo HasError
  sVal4$ = FullBusName( nBusHnd& )

  'Get voltage at the end bus
  If GetPFVoltage( nDevHnd&, vdVal1, vdVal2, ST_KV ) = 0 Then GoTo HasError
  ' Print it
  Print #1, _
        "Voltage on Xfmr3: "; sVal2$ & "-"; sVal3$ & " - " & sVal4$; ": "; Chr(13); _
   "V1 = "; Format( vdVal1(1), "#0.0"); "@"; Format( vdVal2(1), "#0.0"); _
        "; V2 = "; Format( vdVal1(2), "#0.0"); "@"; Format( vdVal2(2), "#0.0"); _
        "; V3 = "; Format( vdVal1(3), "#0.0"); "@"; Format( vdVal2(3), "#0.0")
  ' Show it
  Print "Voltage on Xfmr3: "; sVal2$ & "-"; sVal3$ & " - " & sVal4$; ": "; Chr(13); _
   "V1 = "; Format( vdVal1(1), "#0.0"); "@"; Format( vdVal2(1), "#0.0"); _
        "; V2 = "; Format( vdVal1(2), "#0.0"); "@"; Format( vdVal2(2), "#0.0"); _
        "; V3 = "; Format( vdVal1(3), "#0.0"); "@"; Format( vdVal2(3), "#0.0")
  ' Get current from end buses
  If GetPFCurrent( nDevHnd&, vdVal1, vdVal2, 1 ) = 0 Then GoTo HasError
  ' Print it
  Print #1, _
        "Current on Xfmr3: "; sVal2$ & "-"; sVal3$ & " - "; sVal4$; ": "; Chr(13); _
   "I1 = "; Format( vdVal1(1), "#0.0"); "@"; Format( vdVal2(1), "#0.0"); _
        "; I2 = "; Format( vdVal1(2), "#0.0"); "@"; Format( vdVal2(2), "#0.0"); _
        "; I3 = "; Format( vdVal1(3), "#0.0"); "@"; Format( vdVal2(3), "#0.0")
  ' Show it
  Print "Current on Xfmr3: "; sVal2$ & "-"; sVal3$ & " - "; sVal4$; ": "; Chr(13); _
   "I1 = "; Format( vdVal1(1), "#0.0"); "@"; Format( vdVal2(1), "#0.0"); _
        "; I2 = "; Format( vdVal1(2), "#0.0"); "@"; Format( vdVal2(2), "#0.0"); _
        "; I3 = "; Format( vdVal1(3), "#0.0"); "@"; Format( vdVal2(3), "#0.0")
  ' Get Power
  If GetFlow( nDevHnd&, vdVal1, vdVal2 ) = 0 Then GoTo HasError
  ' Print it
  Print #1, _
        "Power on Xfmr3: "; sVal2$ & "-"; sVal3$ & " - "; sVal4$; ": "; Chr(13); _
   "P1 = "; Format( vdVal1(1), "#0.0"); "Q1 = "; Format( vdVal2(1), "#0.0"); _
        "; P2 = "; Format( vdVal1(2), "#0.0"); "Q2 = "; Format( vdVal2(2), "#0.0"); _
        "; P3 = "; Format( vdVal1(3), "#0.0"); "Q3 = "; Format( vdVal2(3), "#0.0")
  ' Show it
  Print "Power on Xfmr3: "; sVal2$ & "-"; sVal3$ & " - "; sVal4$; ": "; Chr(13); _
   "P1 = "; Format( vdVal1(1), "#0.0"); "Q1 = "; Format( vdVal2(1), "#0.0"); _
        "; P2 = "; Format( vdVal1(2), "#0.0"); "Q2 = "; Format( vdVal2(2), "#0.0"); _
        "; P3 = "; Format( vdVal1(3), "#0.0"); "Q3 = "; Format( vdVal2(3), "#0.0")
  Close 1
  Print "Output written to \0tmp\1.out"
  Stop
HasError:
  Print "Error: ", ErrorString( )
  Close 
  Stop
End Sub