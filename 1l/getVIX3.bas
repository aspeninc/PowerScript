' ASPEN PowerScript example
' GETVX3.BAS
'
' Get and show post fault voltage and current on every 3-W transformer
'
' Demonstrate how PowerScript program can access Shortcircuit
' simulation result
'
' Version 1.0
' Category: OneLiner
'
' PowerScript functions called:
'   GetEquipment()
'   FullBusName()
'   GetBusEquipment()
'   GetData()
'   ShowFault()
'   GetSCVoltage()
'   GetSCCurrent()
'
Sub main()

  'Variable declarations
  Dim MagArray(12) As Double, AngArray(12) As Double

  ' Prepare file for output
  Open "c:\0tmp\1.out" For Output As 1

   ' Loop through all 3-w transformers
   DeviceHnd = 0
   While GetEquipment( TC_XFMR3, DeviceHnd ) > 0
      ' Print line info
      If GetData( DeviceHnd, X3_sID, IDstring ) = 0 Then GoTo HasError
      If GetData( DeviceHnd, X3_nBus1Hnd, BusHnd ) = 0 Then GoTo HasError
      Bus1ID = FullBusName( BusHnd )
      If GetData( DeviceHnd, X3_nBus2Hnd, BusHnd ) = 0 Then GoTo HasError
      Bus2ID = FullBusName( BusHnd )
      If GetData( DeviceHnd, X3_nBus3Hnd, BusHnd ) = 0 Then GoTo HasError
      Bus3ID = FullBusName( BusHnd )

      ' Show the fault
      If PickFault( 1 ) = 0 Then GoTo HasError
      'Get voltagge at the end bus
      If GetSCVoltage( DeviceHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
      ' Print it
      Print #1, _
            "Voltage on Xfmr3: "; Bus1ID & "-"; Bus2ID & " - " & Bus3ID; ": "; Chr(10); _
	         "V1a = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
            "; V1b = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
            "; V1c = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0"); Chr(10); _
            "V2a = "; Format( MagArray(4), "#0.0"); "@"; Format( AngArray(4), "#0.0"); _
            "; V2b = "; Format( MagArray(5), "#0.0"); "@"; Format( AngArray(5), "#0.0"); _
            "; V2c = "; Format( MagArray(6), "#0.0"); "@"; Format( AngArray(6), "#0.0"); Chr(10); _
            "V3a = "; Format( MagArray(7), "#0.0"); "@"; Format( AngArray(7), "#0.0"); _
            "; V3b = "; Format( MagArray(8), "#0.0"); "@"; Format( AngArray(8), "#0.0"); _
            "; V3c = "; Format( MagArray(9), "#0.0"); "@"; Format( AngArray(9), "#0.0")
      ' Show it
      Print _
            "Voltage on Xfmr3: "; Bus1ID & "-"; Bus2ID & " - " & Bus3ID; ": "; Chr(10); _
	         "V1a = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
            "; V1b = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
            "; V1c = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0"); Chr(10); _
            "V2a = "; Format( MagArray(4), "#0.0"); "@"; Format( AngArray(4), "#0.0"); _
            "; V2b = "; Format( MagArray(5), "#0.0"); "@"; Format( AngArray(5), "#0.0"); _
            "; V2c = "; Format( MagArray(6), "#0.0"); "@"; Format( AngArray(6), "#0.0"); Chr(10); _
            "V3a = "; Format( MagArray(7), "#0.0"); "@"; Format( AngArray(7), "#0.0"); _
            "; V3b = "; Format( MagArray(8), "#0.0"); "@"; Format( AngArray(8), "#0.0"); _
            "; V3c = "; Format( MagArray(9), "#0.0"); "@"; Format( AngArray(9), "#0.0")

      ' Get current from transformer end buses
      If GetSCCurrent( DeviceHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
      ' Print it
      Print #1, _
            "Current on Xfmr3: "; Bus1ID & "-"; Bus2ID & " - "; Bus3ID; ": "; Chr(10); _
   	      "I1a = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
            "; I1b = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
            "; I1c = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0"); _
            "; I1g = "; Format( MagArray(4), "#0.0"); "@"; Format( AngArray(4), "#0.0"); Chr(10); _
            "I2a = "; Format( MagArray(5), "#0.0"); "@"; Format( AngArray(5), "#0.0"); _
            "; I2b = "; Format( MagArray(6), "#0.0"); "@"; Format( AngArray(6), "#0.0"); _
            "; I2c = "; Format( MagArray(7), "#0.0"); "@"; Format( AngArray(7), "#0.0"); _
            "; I2g = "; Format( MagArray(8), "#0.0"); "@"; Format( AngArray(8), "#0.0"); Chr(10);  _
            "I3a = "; Format( MagArray(9), "#0.0"); "@"; Format( AngArray(9), "#0.0"); _
            "; I3b = "; Format( MagArray(10), "#0.0"); "@"; Format( AngArray(10), "#0.0"); _
            "; I3c = "; Format( MagArray(11), "#0.0"); "@"; Format( AngArray(11), "#0.0"); _
            "; I3g = "; Format( MagArray(12), "#0.0"); "@"; Format( AngArray(12), "#0.0")
      ' Show it
      Print _
            "Current on Xfmr3: "; Bus1ID & "-"; Bus2ID & " - "; Bus3ID; ": "; Chr(10); _
	         "I1a = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
            "; I1b = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
            "; I1c = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0"); _
            "; I1g = "; Format( MagArray(4), "#0.0"); "@"; Format( AngArray(4), "#0.0"); Chr(10); _
            "I2a = "; Format( MagArray(5), "#0.0"); "@"; Format( AngArray(5), "#0.0"); _
            "; I2b = "; Format( MagArray(6), "#0.0"); "@"; Format( AngArray(6), "#0.0"); _
            "; I2c = "; Format( MagArray(7), "#0.0"); "@"; Format( AngArray(7), "#0.0"); _
            "; I2g = "; Format( MagArray(8), "#0.0"); "@"; Format( AngArray(8), "#0.0"); Chr(10);  _
            "I3a = "; Format( MagArray(9), "#0.0"); "@"; Format( AngArray(9), "#0.0"); _
            "; I3b = "; Format( MagArray(10), "#0.0"); "@"; Format( AngArray(10), "#0.0"); _
            "; I3c = "; Format( MagArray(11), "#0.0"); "@"; Format( AngArray(11), "#0.0"); _
            "; I3g = "; Format( MagArray(12), "#0.0"); "@"; Format( AngArray(12), "#0.0")
   Wend  'Each 3-w transformer
   Close 1
   Print "Output written to \0tmp\1.out"
   Stop
HasError:
  Print "Error: ", ErrorString( )
  Close 
End Sub
