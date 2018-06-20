' ASPEN PowerScript sample program
'
' RlyInTier.bas
'
' Find list of relay group vicinity of
'
' Version 1.0
' Category: OneLiner
'

Sub main
  If 0 = GetEquipment( TC_PICKED, PickedHnd ) Or _
     TC_BUS <> EquipmentType( PickedHnd ) Then
     Print "Please select a bus"
     stop
  End If
  
 
  dim RlyGrpTodo(50) As long
  
  nCount = FindRlyGrpInVicinity( PickedHnd, 2, RlyGrpTodo )
  
  print nCount
  
End Sub

Const MXNEIGHBOR = 50
Function FindRlyGrpInVicinity( ByVal nBusHnd&, ByVal TierWanted&, ByRef RlyGroupTodo() As long ) As long
  dim BusToDo(MXNEIGHBOR) As long
  dim BusTier(MXNEIGHBOR) As long
  dim countRlyGrp As long
  dim countBus As long
  dim nBusTodo As long
  
  BusToDo(1) = nBusHnd
  BusTier(1) = 0
  countBus   = 1
  nBusTodo   = 1
  countRlyGrp& = 0
  While nBusTodo > 0
    nIdx = countBus - nBusTodo + 1
    Bus1Hnd  = BusToDo(nIdx)
    nTier    = BusTier(nIdx)
    nBusTodo = nBusTodo - 1
    BranchHnd   = 0
    countRlyGrp = 0
    While GetBusEquipment( Bus1Hnd, TC_BRANCH, BranchHnd ) > 0
      RlyGrp1Hnd = 0
      Call GetData( BranchHnd, BR_nRlyGrp1Hnd, RlyGrp1Hnd )
      If RlyGrp1Hnd > 0 Then
        nFound = 0
        For ii = 1 to countRlyGrp
          If RlyGroupTodo(ii) = RlyGrp1Hnd Then 
            nFound = 1
            GoTo BreakFor
          End If
        Next
        BreakFor:
        If nFound = 0 Then
          countRlyGrp = countRlyGrp + 1
          RlyGroupTodo(countRlyGrp) = RlyGrp1Hnd
        End If
      End If
      Call GetData( BranchHnd, BR_nBus2Hnd, nBus2Hnd )
      nFound = 0
      For ii = 1 to countBus
        If BusToDo(ii) = nBus2Hnd Then 
          nFound = 1
          GoTo BreakFor2
        End If
      Next
      BreakFor2:
      If nFound = 0 And nTier < TierWanted And countBus < MXNEIGHBOR - 1Then
        countBus = countBus + 1
        BusToDo(countBus) = nBus2Hnd
        BusTier(countbus) = nTier + 1
        nBusTodo = nBusTodo + 1
      End If
    Wend
  Wend
  FindRlyGrpInVicinity = countRlyGrp
End Function 
