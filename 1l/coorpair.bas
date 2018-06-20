' ASPEN PowerScript Sample Program
'
' COORPAIR.BAS
'
' Demonstrate how to get primary and backups of the 
' selected realy group in a PowerScript program
'
' Version 1.0
' Category: OneLiner
'
' PowerScript functions called:
'   GetEquipment()
'   GetData()
'   FullBusName()
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
   strText$ = "Picked relay group: " & RlyBranchID( PickedHnd ) & Chr(10)

   ' List all primaries
   strText$ = StrText$ & "This group is backed up by:" & Chr(10)
   GroupHnd& = 0   ' Force GetData to retrieve the first group
   While GetData( PickedHnd, RG_nPrimaryHnd, GroupHnd ) > 0
     strText$ = StrText$ & "    " & RlyBranchID( GroupHnd ) & Chr(10)
   Wend  'Each group

   ' List all backups
   strText$ = StrText$ & "This group backs up:" & Chr(10)
   GroupHnd& = 0   ' Force GetData to retrieve the first group
   While GetData( PickedHnd, RG_nBackupHnd, GroupHnd ) > 0
     strText$ = StrText$ & "    " & RlyBranchID( GroupHnd ) & Chr(10)
   Wend  'Each group

   ' Output
   Print strText
   Exit Sub

   ' Error handling
   HasError:
   Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()

Function RlyBranchID( ByVal RlyGroupHnd& ) As String
   ' Identify the relay branch
   If GetData( RlyGroupHnd, RG_nBranchHnd, RlyBranchHnd& ) = 0 Then GoTo HasErr1
   If GetData( RlyBranchHnd, BR_nHandle,  DevHnd& )        = 0 Then GoTo HasErr1
   If GetData( RlyBranchHnd, BR_nBus1Hnd, Bus1Hnd& )       = 0 Then GoTo HasErr1
   If GetData( RlyBranchHnd, BR_nBus2Hnd, Bus2Hnd& )       = 0 Then GoTo HasErr1
   If GetData( RlyBranchHnd, BR_nType, RlyBranchType& )    = 0 Then GoTo HasErr1
   Select Case RlyBranchType
     Case TC_LINE
       TypeCode$ = "L"
       If GetData( DevHnd, LN_sID, BrID$ ) = 0 Then GoTo HasErr1
     Case TC_XFMR
       TypeCode$ = "T"
       If GetData( DevHnd, XR_sID, BrID$ ) = 0 Then GoTo HasErr1
     Case TC_XFMR3
       TypeCode$ = "X"
       If GetData( DevHnd, X3_sID, BrID$ ) = 0 Then GoTo HasErr1
     Case TC_PS
       TypeCode$ = "P"
       If GetData( DevHnd, PS_sID, BrID$ ) = 0 Then GoTo HasErr1
   End Select
   strText$ = FullBusName( Bus1Hnd ) & " - " & FullBusName( Bus2Hnd ) & " " _
              & BrID$ & " " & TypeCode$
   RlyBranchID = strText$
   GoTo ExitFunction
   HasErr1:
   RlyBranchID = ""
   ExitFunction:
End Function
