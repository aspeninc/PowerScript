' ASPEN PowerScript sample program
'
' Directional32Q.BAS
'
' Simulated 32Q directional logic
'
' Version 1.0
' Category: OneLiner
'
Sub main()
 Dim MagArray(12) As Double
 Dim AngArray(12) As Double
 Dim DummyArray(6) As Long   '

 
 MTA = 75
 
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
 If GetSCVoltage( Bus1Hnd, MagArray, AngArray, 2 ) = 0 Then GoTo HasError

  OutString$ = "Voltage at " & StringVal & ":" & Chr(13) & _
      "Vo = " & Format( MagArray(1), "#0.0") & "@" & Format( AngArray(1), "#0.0") & _
      "; V+ = " & Format( MagArray(2), "#0.0") & "@" & Format( AngArray(2), "#0.0") & _
      "; V- = " & Format( MagArray(3), "#0.0") & "@" & Format( AngArray(3), "#0.0") & Chr(13)
      
  V2ang = AngArray(3)
  
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
  If GetSCCurrent( BranchHnd, MagArray, AngArray, 2 ) = 0 Then GoTo HasError
  
  OutString$ = OutString$ & _
      TypeString & " current to " & StringVal & ":" & Chr(13) & _
      "Io = " & Format( MagArray(1), "#0.0") & "@" & Format( AngArray(1), "#0.0") & _
      "; I+ = " & Format( MagArray(2), "#0.0") & "@" & Format( AngArray(2), "#0.0") & _
      "; I- = " & Format( MagArray(3), "#0.0") & "@" & Format( AngArray(3), "#0.0")

   I2ang = AngArray(3)
  ' Show it
  
   OutString$ = OutString$ & Chr(13) & "MTA=" & Format(MTA, "#0.0") & chr(13) & _
      "COS(V2Angle-I2Angle-MTA)=" & Format( Cos(V2ang-I2ang-MTA), "#0.00" )
      
   Print OutString$
            
 Exit Sub
HasError:
 Print "Error: "; ErrorString( ) 
End Sub
