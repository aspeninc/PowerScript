' ASPEN PowerScript sample program
'
' ARCFLASH.BAS
'
' Demonstrate the DoArcFlash Command
'
' PowerScript functions used
Sub main()

  'Variable declarations
  Dim vdOption(10) As Double, vdResult(15) As Double
  

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
  vdOption(6) = -1	'Fault clearing:-1- Auto;-2- manual clearing Time;>0- clearing device handle
  vdOption(7) = 1.5	'Breaker interrupting time in cycles or manual clearing time in seconds
  vdOption(8) = 1	'Ignore 2 second flag: 0-reset; 1-set;
  
  
  If 0 = DoAcrFlash( nBusHnd, vdOption, vdResult ) Then GoTo HasError
  
  
'  Print  FullBusName(nBusHnd) & "," & _
'            "Isc(kA),Iarc(kA)," & _
'            "Device1,T1(sec),E1(cal/cm2)," & _
'            "Device2,T2(sec),E2(cal/cm2)," & _
'            "PPE level,PPECat1,PPECat2,PPECat3,PPECat4"
  Print  vdResult(1), "," , vdResult(2) , _
            "," , FullRelayName(vdResult(3)), ",", vdResult(4), ",", vdResult(5), _
            "," , FullRelayName(vdResult(6)), ",", vdResult(7), ",", vdResult(8), _
            "," , vdResult(9), ",", vdResult(10), ",", vdResult(11), ",", vdResult(12), ",", vdResult(13)
            

  Stop
  HasError:
  Print "Error: ", ErrorString( )
  Stop
End Sub