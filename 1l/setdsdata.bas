' ASPEN PowerScript Sample Program
'
' SETDSDATA.BAS
'
' Demonstrate how to set DS relay data
'
' Version 1.0
' Category: OneLiner
'
' PowerScript functions called:
'   GetEquipment()
'   GetRelay()
'   GetData()
'   SetData()
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
       Case TC_RLYDSP
         RlyDevice$ = "DSP: "
         ParamID1& = DP_sID
         ParamID2& = DP_dCT
         ParamID3& = DP_dVT
       Case TC_RLYDSG
         RlyDevice$ = "DSG: "
         ParamID1& = DG_sID
         ParamID2& = DG_dCT
         ParamID3& = DG_dVT
       Case Else
         GoTo Cont
     End Select
     Call GetData( RelayHnd, ParamID1, sID$ )
     dCT# = 111
     dVT# = 222
     Call SetData( RelayHnd, ParamID2, dCT )
     Call SetData( RelayHnd, ParamID3, dVT )
     Call PostData( RelayHnd )
   Cont:
   Wend  'Each relay
   Print "Relays in this group = "; RelayCount
   Exit Sub
   ' Error handling
   HasError:
   Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()
