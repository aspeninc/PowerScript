' ASPEN PowerScript sample program
'
' GETVIRelay.BAS
'
' Show post fault voltage and current on a relay group
'
' Version 1.0
' Category: OneLiner
'
Sub main()
 Dim MagArray(12) As Double
 Dim AngArray(12) As Double
 Dim DummyArray(6) As Long   '

 ' Get picked object number
 If GetEquipment( TC_PICKED, ObjHnd ) = 0 Then 
   Print "Please select a relay group"
   Exit Sub
 End If
 
 If EquipmentType( ObjHnd ) <> TC_RLYGROUP Then 
   Print "Please select a relay group"
   Exit Sub
 End If

 ' Must alway show fault before getting V and I
 If ShowFault( 1, 1, 4, 0, DummyArray ) = 0 Then GoTo HasError

 ' Get branch and bus handles
 Call GetData( ObjHnd, RG_nBranchHnd, BranchHnd )
 Call GetData( BranchHnd, BR_nBus1Hnd, Bus1Hnd )
 
 StringVal$ = FullBusName( Bus1Hnd )
 
 ' Get bus voltage
 If GetSCVoltage( Bus1Hnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
 ' Show it
 Print "Voltage at "; StringVal; ":"; Chr(13); _
      "Va = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
      "; Vb = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
      "; Vc = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")

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
      
 Exit Sub
HasError:
 Print "Error: "; ErrorString( ) 
End Sub
