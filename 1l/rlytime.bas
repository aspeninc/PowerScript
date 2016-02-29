' ASPEN PowerScript Sample Program
'
' RELAY.BAS
'
' Show operating time of relays in the selected group
'
' Demonstrate how to access Shortcircuit simulation result
' from a PowerScript program
'
' PowerScript functions called:
'   GetEquipment()
'   GetRelay()
'   ShowFault()
'   GetRelayTime()
' 
Sub main()
   Dim ShowRelayFlag (4) As Long

   ShowRelayFlag(1) = 1
   ShowRelayFlag(2) = 1
   ShowRelayFlag(3) = 1
   ShowRelayFlag(4) = 1
   ' Get picked device
   If GetEquipment( TC_PICKED, PickedHnd ) = 0 Then GoTo hasError
   ' Probe to see what's being picked
   TypeCode = EquipmentType( PickedHnd )
   If TypeCode <> TC_RLYGROUP Then
     ' Must be a relay group
     Print "Must select a relay group"
     Stop
   End If

   ' Must always show fault first
   If ShowFault( 1, 3, 7, 0, ShowRelayFlag ) = 0 Then GoTo HasError

   ' Loop through all relays and find their operating times
   RelayCount = 0
   RelayHnd   = 0
   While GetRelay( PickedHnd, RelayHnd ) > 0
     RelayCount = RelayCount + 1
     TypeCode = EquipmentType( RelayHnd )
     If TypeCode = TC_RLYOCG Then ParamID = OG_sID
     If TypeCode = TC_RLYOCP Then ParamID = OP_sID
     If TypeCode = TC_RLYDSG Then ParamID = DG_sID
     If TypeCode = TC_RLYDSP Then ParamID = DP_sID
     If TypeCode = TC_FUSE   Then ParamID = FS_sID
     If GetData( RelayHnd, ParamID, sID$ ) = 0 Then GoTo HasError
     If GetRelayTime( RelayHnd, 1.0, OpTime ) = 0 Then GoTo HasError
     Print "Relay " & sID & ": "; Format( OpTime, "#0.#0s" )
   Wend  'Each relay
   Print "Relays in this group = "; RelayCount
   Exit Sub
   ' Error handling
   HasError:
   Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()