' ASPEN PowerScript Sample Program
'
' RLYDATA.BAS
'
' Retrieve data of relays in the selected group
'
' Demonstrate how to access Shortcircuit simulation result
' from a PowerScript program
'
' Version 1.0
' Category: OneLiner
'
' PowerScript functions called:
'   GetEquipment()
'   GetRelay()
' 
Sub main()
   Dim ShowRelayFlag (4) As Long
   Dim vDSParams( 40 ) As Double 
   Dim ii As Long
   Dim nCount As Long 

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

   ' Loop through all relays and find their settings
   RelayCount = 0
   RelayHnd   = 0
   While GetRelay( PickedHnd, RelayHnd ) > 0
     RelayCount = RelayCount + 1
     TypeCode = EquipmentType( RelayHnd )
     Select Case TypeCode
       Case TC_RLYOCG
         ParamID1& = OG_sID
         ParamID2& = OG_dTap
         ParamID3& = OG_dTDial
         ParamID4& = OG_sType
         ParamID5& = OG_dInst
         ParamID6& = OG_dInstDelay
         ParamID7& = OG_dTimeAdd
         ParamID8& = OG_dTimeMult
         ParamID9& = OG_dTimeReset
       Case TC_RLYOCP
         ParamID1& = OP_sID
         ParamID2& = OP_dTap
         ParamID3& = OP_dTDial
         ParamID4& = OP_sType
         ParamID5& = OP_dInst
         ParamID6& = OP_dInstDelay
         ParamID7& = OP_dTimeAdd
         ParamID8& = OP_dTimeMult
         ParamID9& = OP_dTimeReset
       Case TC_RLYDSP
         ParamID1& = DP_sID
         ParamID4& = DP_sType
         ParamID5& = DP_sDSType
         ParamID6& = DP_vdParams
         ParamID7& = DP_nParamCount
       Case TC_RLYDSG
         ParamID1& = DG_sID
         ParamID4& = DG_sType
         ParamID5& = DG_sDSType
         ParamID6& = DG_vdParams
         ParamID7& = DG_nParamCount
       Case Else
         GoTo Cont
     End Select
     If TypeCode=TC_RLYOCG Or TypeCode=TC_RLYOCP Then 
       Call GetData( RelayHnd, ParamID1, sID$ )
       Call GetData( RelayHnd, ParamID2, dTap# )
       Call GetData( RelayHnd, ParamID3, dTD# )
       Call GetData( RelayHnd, ParamID4, sType$ )
       Call GetData( RelayHnd, ParamID5, dInst# )
       Call GetData( RelayHnd, ParamID6, dInstDelay# )
       Call GetData( RelayHnd, ParamID7, dTimeAdd# )
       Call GetData( RelayHnd, ParamID8, dTimeMult# )
       Call GetData( RelayHnd, ParamID9, dTimeReset# )
       Print "OC Relay " & sID & "(" & sType & ")" & ": Tap=" & Format( dTap, "#0.#0" ) & "; TD="; Format( dTD, "#0,#0" ) & _
              "; Inst=" & Format( dInst, "#0" ) & "; InstDelay=" & Format( dInstDelay, "#0.0" ) & _
              "; Tmult=" & Format( dTimeMult, "#0.0" ) & "; Tadd=" & Format( dTimeAdd, "#0.0" ) & "; Treset=" & Format( dTimeReset, "#0.0" )
     Else
       Call GetData( RelayHnd, ParamID1, sID$ )
       Call GetData( RelayHnd, ParamID4, sType$ )
       Call GetData( RelayHnd, ParamID5, sDSType$ )
       Call GetData( RelayHnd, ParamID7, nCount& )
       If 0 = GetData( RelayHnd, ParamID6, vDSParams ) Then GoTo hasError
       sPrint$ = "DS Relay " & sID & "(" & sType & ")" & "(" & sDSType & ")" & Chr(13)
       For ii=1 To nCount
         sPrint$ = sPrint$ + " Param[" + Str(ii) + "]=" + Str(vDSParams(ii)) + Chr(13)
       Next ii
       Print sPrint
     End If
   Cont:
   Wend  'Each relay
   Print "Relays in this group = "; RelayCount
   Exit Sub
   ' Error handling
   HasError:
   Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()
