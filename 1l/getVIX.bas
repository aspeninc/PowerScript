' ASPEN PowerScript example
' GETVIX.BAS
'
' Get and show post fault voltage and current on 
' every transformer
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
  Dim MagArray(12) As Double, AngArray(12) As Double
  Dim DummyArray(4) As Long

  ' Prepare output file
  Open "c:\0tmp\1.out" For Output As 1

   ' Loop through all transformer in the network
   DeviceHnd = 0
   While GetEquipment( TC_XFMR, DeviceHnd ) > 0
      ' Get transformer info
      If GetData( DeviceHnd, XR_sID, IDstring ) = 0 Then GoTo HasError
      If GetData( DeviceHnd, XR_nBus1Hnd, BusHnd ) = 0 Then GoTo HasError
      Bus1ID = FullBusName( BusHnd )
      If GetData( DeviceHnd, XR_nBus2Hnd, BusHnd ) = 0 Then GoTo HasError
      Bus2ID = FullBusName( BusHnd )

      ' Must always show fault before getting V and I
      If ShowFault( 1, 3, 4, 0, DummyArray ) = 0 Then GoTo HasError

      ' Get voltagge at the end bus
      If GetSCVoltage( DeviceHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
      ' Show it
      Print _
            "Voltage on Xfmr: "; Bus1ID & "-"; Bus2ID & " ID= "; IDstring; ": "; Chr(10); _
	         "V1a = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
            "; V1b = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
            "; V1c = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0"); Chr(10); _
            "V2a = "; Format( MagArray(4), "#0.0"); "@"; Format( AngArray(4), "#0.0"); _
            "; V2b = "; Format( MagArray(5), "#0.0"); "@"; Format( AngArray(5), "#0.0"); _
            "; V2c = "; Format( MagArray(6), "#0.0"); "@"; Format( AngArray(6), "#0.0")
      ' Print it
      Print #1, _
            "Voltage on Xfmr: "; Bus1ID & "-"; Bus2ID & " ID= "; IDstring; ": "; Chr(10); _
	         "V1a = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
            "; V1b = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
            "; V1c = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0"); Chr(10); _
            "V2a = "; Format( MagArray(4), "#0.0"); "@"; Format( AngArray(4), "#0.0"); _
            "; V2b = "; Format( MagArray(5), "#0.0"); "@"; Format( AngArray(5), "#0.0"); _
            "; V2c = "; Format( MagArray(6), "#0.0"); "@"; Format( AngArray(6), "#0.0")

      ' Get current from end buses
      If GetSCCurrent( DeviceHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
      ' Show it
      Print _
            "Current on Xfmr: "; Bus1ID & "-"; Bus2ID & " ID= "; IDstring; ": "; Chr(10); _
	         "I1a = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
            "; I1b = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
            "; I1c = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0"); _
            "; I1g = "; Format( MagArray(4), "#0.0"); "@"; Format( AngArray(4), "#0.0"); Chr(10); _
            "I2a = "; Format( MagArray(5), "#0.0"); "@"; Format( AngArray(5), "#0.0"); _
            "; I2b = "; Format( MagArray(6), "#0.0"); "@"; Format( AngArray(6), "#0.0"); _
            "; I2c = "; Format( MagArray(7), "#0.0"); "@"; Format( AngArray(7), "#0.0"); _
            "; I2g = "; Format( MagArray(8), "#0.0"); "@"; Format( AngArray(8), "#0.0")
      ' Print it
      Print #1, _
            "Current on Xfmr: "; Bus1ID & "-"; Bus2ID & " ID= "; IDstring; ": "; Chr(10); _
	         "I1a = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
            "; I1b = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
            "; I1c = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0"); _
            "; I1g = "; Format( MagArray(4), "#0.0"); "@"; Format( AngArray(4), "#0.0"); Chr(10); _
            "I2a = "; Format( MagArray(5), "#0.0"); "@"; Format( AngArray(5), "#0.0"); _
            "; I2b = "; Format( MagArray(6), "#0.0"); "@"; Format( AngArray(6), "#0.0"); _
            "; I2c = "; Format( MagArray(7), "#0.0"); "@"; Format( AngArray(7), "#0.0"); _
            "; I2g = "; Format( MagArray(8), "#0.0"); "@"; Format( AngArray(8), "#0.0")
   Wend  ' Each transformer

   ' Clean up
   Close 1
   Print "Output written to \0tmp\1.out"
   Exit Sub
HasError:
  Print "Error: ", ErrorString( )
  Close 
End Sub
