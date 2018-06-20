' ASPEN PowerScript example
' GETVISW.BAS
'
' Get and show post fault voltage and current of 
' a switch
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
  Dim vnShowRelay(4) As Long

  ' Prepare output file
  If 0 = GetEquipment( TC_PICKED, DeviceHnd ) Then 
    Print "Please select a switch"
    Exit Sub
  End If
  If TC_SWITCH <> EquipmentType( DeviceHnd ) Then
    Print "Please select a switch"
    Exit Sub
  End If

  ' Get Switch info
  If GetData( DeviceHnd, SW_nBus1Hnd, BusHnd ) = 0 Then GoTo HasError
  Bus1ID = FullBusName( BusHnd )
  If GetData( DeviceHnd, SW_nBus2Hnd, BusHnd ) = 0 Then GoTo HasError
  Bus2ID = FullBusName( BusHnd )

  ' Must always pick or show fault before getting short circuit V and I
  If PickFault( 1 ) = 0 Then GoTo HasError

  ' Get voltagge at the end bus
  If GetSCVoltage( DeviceHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
  ' Show it
  Print _
            "Voltage on switch: "; Bus1ID & "-"; Bus2ID ; Chr(10); _
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
            "Current on switch: "; Bus1ID & "-"; Bus2ID & " ID= "; PsID; ": "; Chr(10); _
	      "I1a = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
            "; I1b = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
            "; I1c = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0"); _
            "; I1g = "; Format( MagArray(4), "#0.0"); "@"; Format( AngArray(4), "#0.0"); Chr(10); _
            "I2a = "; Format( MagArray(5), "#0.0"); "@"; Format( AngArray(5), "#0.0"); _
            "; I2b = "; Format( MagArray(6), "#0.0"); "@"; Format( AngArray(6), "#0.0"); _
            "; I2c = "; Format( MagArray(7), "#0.0"); "@"; Format( AngArray(7), "#0.0"); _
            "; I2g = "; Format( MagArray(8), "#0.0"); "@"; Format( AngArray(8), "#0.0")
  Exit Sub
HasError:
  Print "Error: ", ErrorString( )
  Close 
End Sub
