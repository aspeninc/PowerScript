' ASPEN PowerScript sample program
' GETVIBr.BAS
'
' Get and show post fault voltage and current on each 
' branch connected to selected bus
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
 Dim DummyArray(6) As Long   '

 Open "\0tmp\1.out" For Output As 1 ' For stroring output

 ' Get picked bus handle number
 If GetEquipment( TC_PICKED, Bus1Hnd ) = 0 Then
   Print "Must select a bus"
   Exit Sub
 End If
 If EquipmentType( Bus1Hnd ) <> TC_BUS Then
   Print "Must select a bus"
   Exit Sub
 End If

 StringVal$ = FullBusName( Bus1Hnd )
 ' Must alway show fault before getting V and I
 If ShowFault( 1, 1, 4, 0, DummyArray ) = 0 Then GoTo HasError

 ' Get bus voltage
 If GetSCVoltage( Bus1Hnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
 ' Show it
 Print "Voltage at "; StringVal; ":"; Chr(13); _
      "Va = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
      "; Vb = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
      "; Vc = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
 ' Print it
 Print #1, _
      "Voltage at "; StringVal; ":"; Chr(13); _
      "Va = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
      "; Vb = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
      "; Vc = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")

 ' Retrieve all branches and get their currents
 BranchHnd = 0
 While GetBusEquipment( Bus1Hnd, TC_BRANCH, BranchHnd ) > 0
  ' Get branch type
  If GetData( BranchHnd, BR_nType, TypeCode ) = 0 Then GoTo HasError
  Select Case TypeCode
   Case TC_LINE
    TypeString = "Line"
   Case TC_XFMR
    TypeString = "Transformer"
   Case TC_PS
    TypeString = "Phase shifter"
   Case TC_XFMR3
    TypeString = "3-w transformer"
   Case Else
    TypeString = "Device"
  End Select

  ' Get far bus handle number
  If GetData( BranchHnd, BR_nBus2Hnd, Bus2Hnd ) = 0 Then GoTo HasError
  ' Find it's  name
  StringVal$ = FullBusName( Bus2Hnd )

  ' Get branch current
  If GetSCCurrent( BranchHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
  ' Show it
  Print TypeString; " current to "; StringVal; ":"; Chr(13); _
      "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
      "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
      "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
  ' Print it
  Print #1, _
      TypeString; " Current to "; StringVal; ":"; Chr(13); _
      "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
      "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
      "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
 Wend    'Each branch
 Close   'Close output file
 Exit Sub
HasError:
 Print "Error: "; ErrorString( ) 
End Sub
