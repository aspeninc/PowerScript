' ASPEN PowerScript Sample Program
'
' GETDSDATA.BAS
'
' Demonstrate how to get distance relay data
'
' Version: 1.0
' Category: OneLiner
'
' PowerScript functions called:
'   GetEquipment()
'   GetRelay()
'   GetData()
'   SetData()
' 
Sub main()

	dim arrSetting(256) As variant
	dim arrLabel(256) As variant
	
   ' Get picked device
   If GetEquipment( TC_PICKED, PickedHnd ) = 0 Then GoTo hasError
   ' Probe to see what's being picked
   TypeCode = EquipmentType( PickedHnd )
   If TypeCode <> TC_RLYGROUP Then
     ' Must be a relay group
     Print "Must select a relay group"
     Stop
   End If

   ' Loop through all DS relays to change CT and VT ratios
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
         ParamID4& = DP_vParams
         ParamID5& = DP_vParamLabels
       Case TC_RLYDSG
         RlyDevice$ = "DSG: "
         ParamID1& = DG_sID
         ParamID2& = DG_dCT
         ParamID3& = DG_dVT
         ParamID4& = DG_vParams
         ParamID5& = DG_vParamLabels
       Case Else
         GoTo Cont
     End Select
     Call GetData( RelayHnd, ParamID1, sID$ )
     Print "Relay " + sID$
     Call GetData( RelayHnd, ParamID4, arrSetting )
     Call GetData( RelayHnd, ParamID5, arrLabel )
     Print arrLabel(1) & "=" & arrSetting(1)
   Cont:
   Wend  'Each relay
   Print "Relays in this group = "; RelayCount
   Exit Sub
   ' Error handling
   HasError:
   Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()