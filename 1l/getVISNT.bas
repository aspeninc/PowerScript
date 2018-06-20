' ASPEN PowerScript sample program
' GETVISNT.BAS
'
' Get and show post fault voltage and current on every 
' shunt connected to selected bus
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
 Dim MagArray(12) As Double
 Dim AngArray(12) As Double
 Dim DummyArray(6) As Long

 'Prepare output file
 Open "\0tmp\1.out" For Output As 1

 ' Get the picked bus handle number
 If GetEquipment( TC_PICKED, BusHnd ) = 0 Then
   Print "Must select a bus"
   Exit Sub
 End If
 If EquipmentType( BusHnd ) <> TC_BUS Then
   Print "Must select a bus"
   Exit Sub
 End If
 BusID = FullBusName( BusHnd )
 ' Show the fault
 If ShowFault( 1, 1, ST_A, 0, DummyArray ) = 0 Then GoTo HasError
 ' Get bus voltage
 If GetSCVoltage( BusHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
 ' Show it
 Print "Voltage at "; sString; ":"; Chr(10); _
      "Va = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
      "; Vb = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
      "; Vc = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
 ' Print it
 Print #1, _
      "Voltage at "; sString; ":"; Chr(10); _
      "Va = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
      "; Vb = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
      "; Vc = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")

 ' Retrieve load units and get their currents
 ShuntHnd = 0
 While GetBusEquipment( BusHnd, TC_LOADUNIT, ShuntHnd ) > 0
  ' Get load unit ID
  If GetData( ShuntHnd, LU_sID, IDstring ) = 0 Then GoTo HasError
  ' Get load current
  If GetSCCurrent( ShuntHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
  ' Show it
  Print _
      "Current from Load Unit "; IDstring; ":"; Chr(10); _
      "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
      "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
      "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
  ' Print it
  Print #1, _
      "Current from Load Unit "; IDstring; ":"; Chr(10); _
      "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
      "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
      "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
 Wend ' Each unit

 ' Get load handle number
 If GetBusEquipment( BusHnd, TC_LOAD, ShuntHnd ) > 0 Then
    ' Get load current
    If GetSCCurrent( ShuntHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
    ' Show it
    Print _
         "Current from load on this bus:"; Chr(10); _
         "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
         "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
         "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
    ' Print it
    Print #1, _
         "Current from load on this bus:"; Chr(10); _
         "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
         "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
         "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
 End If

 ' Get Switched Shunt handle number
 If GetBusEquipment( BusHnd, TC_SVD, ShuntHnd ) > 0 Then 
    ' Get switched shunt current
    If GetSCCurrent( ShuntHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
    ' Show it
    Print _
         "Current from Switched shunt on this bus:"; Chr(10); _
         "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
         "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
         "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
    ' Print it
    Print #1, _
         "Current from Switched shunt on this bus:"; Chr(10); _
         "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
         "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
         "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
 End If

 ' Retrieve all shunt units and get their currents
 ShuntHnd = 0
 While GetBusEquipment( BusHnd, TC_SHUNTUNIT, ShuntHnd ) > 0
  ' Get Shutn unit ID
  If GetData( ShuntHnd, SU_sID, IDstring ) = 0 Then GoTo HasError
  ' Get Gen current
  If GetSCCurrent( ShuntHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
  ' Show it
  Print _
      "Current from Shunt Unit "; IDstring; ":"; Chr(10); _
      "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
      "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
      "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
  ' Print it
  Print #1, _
      "Current from Shunt Unit "; IDstring; ":"; Chr(10); _
      "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
      "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
      "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
 Wend ' Each unit

 ' Get Shunt handle
 If GetBusEquipment( BusHnd, TC_SHUNT, ShuntHnd ) > 0 Then
    ' Get Shunt current
    If GetSCCurrent( ShuntHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
    ' Show it
    Print _
         "Current from Shunt on this bus:"; Chr(10); _
         "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
         "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
         "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
    ' Print it
    Print #1, _
         "Current from Shunt on this bus:"; Chr(10); _
         "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
         "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
         "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
 End If

 ' Retrieve all generating units and get their currents
 ShuntHnd = 0
 While GetBusEquipment( BusHnd, TC_GENUNIT, ShuntHnd ) > 0
  ' Get Generator ID
  If GetData( ShuntHnd, GU_sID, IDstring ) = 0 Then GoTo HasError
  ' Get Gen current
  If GetSCCurrent( ShuntHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
  ' Show it
  Print "Current from GenUnit "; IDstring; ":"; Chr(10); _
      "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
      "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
      "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
  ' Print it
  Print #1, _
      "Current from GenUnit "; IDstring; ":"; Chr(10); _
      "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
      "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
      "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
 Wend ' Each unit

 ' Get Generator handle
 If GetBusEquipment( BusHnd, TC_GEN, ShuntHnd ) > 0 Then
    ' Get Generator current
    If GetSCCurrent( ShuntHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
    ' Show it
    Print _
         "Current from Generator on this bus:"; Chr(10); _
         "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
         "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
         "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
    ' Print it
    Print #1, _
         "Current from Generator on this bus:"; Chr(10); _
         "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
         "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
         "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
 End If
 Close
 Exit Sub
HasError:
 Print "Error: ", ErrorString( )
End Sub
