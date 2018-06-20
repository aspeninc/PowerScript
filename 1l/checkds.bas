' ASPEN PowerScrip sample program
'
' DSCHECK.BAS
'
' Check zone reach setting of distance relays
'
' Version 1.0
' Category: OneLiner
'
' PowerScript functions called:
'
Sub main()

 BusHnd& = 0
 While NextBusByName( BusHnd ) > 0
  If 0 = GetData( BusHnd, BUS_nTapBus, nFlag& ) Then GoTo HasError
  If nFlag <> 0 Then GoTo Continue1
  Call DoOneBus(BusHnd)
 Continue1:
 Wend

 exit Sub
HasError:
  Print "Error: ", ErrorString( )
End Sub

Sub DoOneBus( ByVal Bus1Hnd& )
 Bus1Name$ = ""
 ' Retrieve all branches and get their currents
 BranchHnd& = 0
 While GetBusEquipment( Bus1Hnd, TC_BRANCH, BranchHnd ) > 0
  Bus2Name$ = ""
  ' Branch must be a line
  Call GetData( BranchHnd, BR_nType, TypeCode& )
  If TypeCode <> TC_LINE Then GoTo Continue
  ' Line must have a relay group
  If GetData( BranchHnd, BR_nRlyGrp1Hnd, GroupHnd& ) <= 0 Then GoTo continue
  
  Call DoOneGroup( GroupHnd )
  
  Continue:
 Wend
 exit Sub
HasError:
  Print "Error: ", ErrorString( )
End Sub

Function GetZoneReach( ByVal RlyHnd&, ByVal ZoneNo& ) As double
 GetZoneReach = 0.0
End Function

Function GetImpedance( ByVal BranchHnd& ) As double
 GetImpedance = 0.0
End Function

Function DoOneGroup( ByVal GroupHnd ) As long
 Call GetData( GroupHnd, RG_nBranchHnd, BranchHnd& )
 Call GetData( BranchHnd, BR_nBus1Hnd, Bus1Hnd& )
 Call GetData( BranchHnd, BR_nBus2Hnd, Bus2Hnd& )
 Bus1Name$ = FullBusName( Bus1Hnd )
 Bus2Name$ = FullBusName( Bus2Hnd )
 LineName$ = Bus1Name + " - " + Bus2Name
 LineImpedance# = GetImpedance( BranchHnd )

 ' Group must have a phase DS relay
 Count& = 0
 RelayHnd& = 0
 While GetRelay( GroupHnd, RelayHnd ) > 0
  TypeCode& = EquipmentType( RelayHnd )
  If TypeCode = TC_RLYDSP Then 
   Call GetData( RelayHnd, DP_sID, sRelayID$ )
   Zone1Reach# = GetZoneReach( RelayHnd, 1 )
   aString$ = "DS Phase Relay " + sRelayID + Chr(13) + "Zone 1 reach=" + Str(Zone1Reach)   
   If LineImpedance*Threshold < Zone1Reach Then
    aString$ = aString$ + Chr(13) + "NOT OK"
   Else 
    aString$ = aString$ + Chr(13) + "OK"
   End If
   Print LineName + Chr(13) + aString
   Count = Count + 1
  End If
 Wend  'Each relay
 DoOneGroup = Count
End Function
