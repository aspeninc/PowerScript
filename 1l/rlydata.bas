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

   ' Get picked device
   If GetEquipment( TC_PICKED, PickedHnd ) = 0 Then GoTo hasError
   ' Probe to see what's being picked
   TypeCode = EquipmentType( PickedHnd )
   If TypeCode <> TC_RLYGROUP Then
     ' Must be a relay group
     Print "Must select a relay group"
     Stop
   End If

   ' Loop through all relays and find their operating times
   RelayCount = 0
   RelayHnd   = 0
   While GetRelay( PickedHnd, RelayHnd ) > 0
     RelayCount = RelayCount + 1
     TypeCode = EquipmentType( RelayHnd )
     Select Case TypeCode
       Case TC_RLYOCG
         RlyDevice$ = "OCG: "
         ParamID1& = OG_sID
         ParamID2& = OG_dTap
         ParamID3& = OG_dTDial
         ParamID4& = OG_sComment
       Case TC_RLYOCP
         RlyDevice$ = "OCP: "
         ParamID1& = OP_sID
         ParamID2& = OP_dTap
         ParamID3& = OP_dTDial
         ParamID4& = OP_sComment
       Case TC_FUSE
         RlyDevice$ = "Fuse: "
         ParamID1& = FS_sID
         ParamID4& = FS_sComment
       Case TC_RLYDSP
         RlyDevice$ = "DSP: "
         ParamID1& = DP_sID
         ParamID4& = DP_sComment
       Case TC_RLYDSG
         RlyDevice$ = "DSG: "
         ParamID1& = DG_sID
         ParamID4& = DG_sComment
       Case Else
         GoTo Cont
     End Select
     Call GetData( RelayHnd, ParamID1, sID$ )
     Call GetData( RelayHnd, ParamID4, sInfo$ )
     If (RlyDevice = "OCG: ") Or (RlyDevice = "OCP: ") Then
       Call GetData( RelayHnd, ParamID2, dTap# )
       Call GetData( RelayHnd, ParamID3, dTD# )
       Print RlyDevice & sID & " Tap="; Format( dTap, "#0.#0" ); "; TD="; Format( dTD, "#0,#0" ); _
             "; Comment: " & sInfo
     Else
       Print RlyDevice & sID & "; Comment: " & sInfo
     End If
   Cont:
   Wend  'Each relay
   Print "Relays in this group = "; RelayCount
   Exit Sub
   ' Error handling
   HasError:
   Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()