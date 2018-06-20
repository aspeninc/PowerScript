' ASPEN PowerScript sample program
'
' Directional32I.BAS
'
' Simulate 32I directional logic
'
' Version 1.0
' Category: OneLiner
'
Sub main()
 Dim MagArray(12) As Double
 Dim AngArray(12) As Double
 Dim DummyArray(6) As Long   '

 
 MTA = 75
 
 UseTertiary = 1  '0 or 1
 
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
 Call GetData( ObjHnd, RG_nBranchHnd, BranchHndRly )
 Call GetData( BranchHndRly, BR_nBus1Hnd, Bus1Hnd )
 
 ' Look for transformer at the bus
 BranchHnd& = 0
 IpolMag   = 0
 IpolAng   = 0
 While GetBusEquipment( Bus1Hnd, TC_BRANCH, BranchHnd& ) > 0
   ' Get branch type
   Call GetData( BranchHnd&, BR_nInService, nInService )
   If nInService <> 1 Then GoTo continueLoop
   Call GetData( BranchHnd&, BR_nType, TypeCode )
   Call GetData( BranchHnd&, BR_nHandle, DeviceHnd )
   Select Case TypeCode
    Case TC_LINE
     TypeString = "Line"
    Case TC_XFMR
     TypeString = "Transformer"
     Call GetData( DeviceHnd, XR_sID, sID$ )
    Case TC_PS
     TypeString = "Phase shifter"
    Case TC_XFMR3
     TypeString = "3-w transformer"
     Call GetData( DeviceHnd, X3_sID, sID$ )
    Case Else
     TypeString = "Device"
   End Select

   If (UseTertiary=1 And TypeCode=TC_XFMR3) Or (UseTertiary<>1 And TypeCode=TC_XFMR) Then
    ' Get far bus handle number
    If GetData( BranchHnd&, BR_nBus2Hnd, Bus2Hnd ) = 0 Then GoTo HasError
    ' Find it's  name
    StringVal$ = FullBusName( Bus2Hnd )

    ' Get branch current
    If GetSCCurrent( BranchHnd&, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
    IpolMag = MagArray(4)
    IpolAng = AngArray(4)
    ' Show it
    OutString$ = TypeString & sID$ & " to " & StringVal$ & " neutral current In=" & Format( IpolMag, "#0.0") & "@" & Format( IpolAng, "#0.0")
    exit Do
   End If
   continueLoop:
  Wend    'Each branch
 
  If IpolMag = 0 And IpolAng = 0 Then
    Print "No active transformer found at this location"
    exit Sub
  End If
 
 
  ' Get branch type
  If GetData( BranchHndRly, BR_nType, TypeCode ) = 0 Then GoTo HasError
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
  If GetData( BranchHndRly, BR_nBus2Hnd, Bus2Hnd ) = 0 Then GoTo HasError
  ' Find it's  name
  StringVal$ = FullBusName( Bus2Hnd )

  ' Get branch current
  If GetSCCurrent( BranchHndRly, MagArray, AngArray, 2 ) = 0 Then GoTo HasError
  
  OutString$ = OutString$ & Chr(13) & _
      TypeString & " current to " & StringVal & ": " &  _
      "Io = " & Format( MagArray(1), "#0.0") & "@" & Format( AngArray(1), "#0.0")

  I0mag = MagArray(1)
  I0ang = AngArray(1)
  ' Show it
  
   OutString$ = OutString$ & Chr(13) & "MTA=" & Format(MTA, "#0.0") & chr(13) & _
      "COS(IpolAngle-IoAngle-MTA)=" & Format( Cos(3.14156/180*IpolAng-I0ang-MTA), "#0.00" )
      
   Print OutString$
            
 Exit Sub
HasError:
 Print "Error: "; ErrorString( ) 
End Sub
