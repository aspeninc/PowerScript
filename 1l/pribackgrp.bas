' ASPEN PowerScript sample program
'
' PRIBACK.BAS
'
' Report relay group coordination pairs
' 
' Version 1.0
' Category: OneLiner
'
' ===================== main() ================================================
'
Sub main()

  PickedHnd& = 0
  bPicked = 0
  Call GetEquipment( TC_PICKED, PickedHnd )
  ' Probe to see what's being picked
  If TC_RLYGROUP = EquipmentType( PickedHnd ) Then
   bPicked = 1
  Else
   resp = MsgBox( "Do you want to print all coordination pairs?", 4+32, "PowerScript" ) 
   If 6 <> resp Then Stop
   PickedHnd = 0
  End If

  If PickedHnd = 0 Then Call GetEquipment( TC_RLYGROUP, PickedHnd )

  While PickedHnd <> 0  
   PrintTTY("Relay group: " & FullBranchName( PickedHnd ) )
  
   PrintTTY("  Backups for this group:")
   GrHnd = 0
   While 0 < Getdata( PickedHnd, RG_nBackupHnd, GrHnd )
    PrintTTY( "   " & FullBranchName( GrHnd ) )
   Wend
  
   PrintTTY("  This group backs up:")
   GrHnd = 0
   While 0 < Getdata( PickedHnd, RG_nPrimaryHnd, GrHnd )
    PrintTTY( "   " & FullBranchName( GrHnd ) )
   Wend
   If bPicked = 1 Then exit Do
   ' Get next relay group
   If 0 >= GetEquipment( TC_RLYGROUP, PickedHnd ) Then exit Do
  Wend
  
  Print "See report in TTY window"
  Stop
Exit Sub
  ' Error handling
  HasError:
  Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()

Function FullBranchName( ByVal GrHnd& ) As String
  Call Getdata( GrHnd, RG_nBranchHnd, BrHnd )
  Call Getdata( BrHnd, BR_nBus1Hnd, Bus1Hnd& )
  Call Getdata( BrHnd, BR_nBus2Hnd, Bus2Hnd& )
  Call Getdata( BrHnd, BR_nHandle, EqHnd& )
  cType = ""
  cID   = ""
  select case EquipmentType( EqHnd )
   case TC_LINE
     cType = "L"
     Call Getdata( EqHnd, LN_sID, cID )
   case TC_XFMR
     cType = "T"
     Call Getdata( EqHnd, XR_sID, cID )
   case TC_XFMR3
     cType = "X"
     Call Getdata( EqHnd, X3_sID, cID )
   case TC_PS
     cType = "P"
     Call Getdata( EqHnd, PS_sID, cID )
  End select
  FullBranchName = FullBusName(Bus1Hnd&) & " - " & FullBusName(Bus2Hnd&) & " " & cID & " " & cType
End Function
