' ASPEN PowerScript Sample Program
'
' SETFUSE.BAS
'
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

   ' Loop through relays to find fuse
   RelayCount = 0
   RelayHnd   = 0
   While GetRelay( PickedHnd, RelayHnd ) > 0
     RelayCount = RelayCount + 1
     If TC_FUSE <> EquipmentType( RelayHnd ) then GoTo Cont
     Call GetData( RelayHnd, FS_sID, sID$ )
     Call GetData( RelayHnd, FS_nCurve, nCurve& )
     If nCurve > 1 Then 
       Print "Fuse=" + sID$ + " Curve= Total clear"
       nCurve = 1
     Else
       Print "Fuse=" + sID$ + " Curve= Min. melt"
       nCurve = 2
     End If
     If SetData( RelayHnd, FS_nCurve, nCurve& ) = 0 Then GoTo HasError
     If PostData( RelayHnd ) = 0 Then GoTo HasError
     If nCurve > 1 Then 
       Print "Fuse=" + sID$ + " Curve= Total clear"
     Else
       Print "Fuse=" + sID$ + " Curve= Min. melt"
     End If
   Cont:
   Wend  'Each relay
   Print "Fuses in this group = "; RelayCount
   Exit Sub
   ' Error handling
   HasError:
   Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()