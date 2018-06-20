' ASPEN PowerScript Sample Program
'
' RELAYTIMEVICINITY.BAS
'
' Show operating time of relays near selected bus
'
' Demonstrate how to access Shortcircuit simulation result
' from a PowerScript program
'
' Version 1.0
' Category: OneLiner
' 

Const MXNEIGHBOR = 50

Sub main
  If 0 = GetEquipment( TC_PICKED, PickedHnd ) Or _
    TC_BUS <> EquipmentType( PickedHnd ) Then
    Print "Please select a bus"
    Stop
  End If
  
 
  ' Must always show fault first
  Dim ShowRelayFlag (4) As Long
  ShowRelayFlag(1) = 1
  ShowRelayFlag(2) = 1
  ShowRelayFlag(3) = 1
  ShowRelayFlag(4) = 1
   
  If ShowFault( 1, 3, 7, 0, ShowRelayFlag ) = 0 Then GoTo HasError

  dim RlyGrpTodo(50) As long
  
  nCount = FindRlyGrpInVicinity( PickedHnd, 2, RlyGrpTodo )
  
  For ii = 1 to nCount 
   nRlyGroup = RlyGrpTodo(ii)
   dTime = fastestTime( nRlyGroup )
  Next   
  
  Stop
HasError:
  Print ErrorString()
End Sub


Function fastestTime( ByVal nRlyGrHnd& ) As double
   fastestTime = 9999
   ' Loop through all relays and find their operating times
   RelayHnd   = 0
   While GetRelay( nRlyGrHnd, RelayHnd ) > 0
     RelayCount = RelayCount + 1
     TypeCode = EquipmentType( RelayHnd )
     If TypeCode = TC_RLYOCG Then ParamID = OG_sID
     If TypeCode = TC_RLYOCP Then ParamID = OP_sID
     If TypeCode = TC_RLYDSG Then ParamID = DG_sID
     If TypeCode = TC_RLYDSP Then ParamID = DP_sID
     If TypeCode = TC_FUSE   Then ParamID = FS_sID
     Call GetData( RelayHnd, ParamID, sID$ )
     If GetRelayTime( RelayHnd, 1.0, OpTime ) > 0 Then
       If OpTime < fastestTime Then fastestTime = OpTime
     End If
   Wend  'Each relay
   

   'Get the output ready
   Call getdata(nRlyGrHnd, RG_nBranchHnd, BranchHnd)
   Call GetData( BranchHnd, BR_nBus1Hnd, Bus1Hnd )
   Call GetData( BranchHnd, BR_nBus2Hnd, Bus2Hnd )
   Call GetData( BranchHnd, BR_nHandle, DeviceHnd )
   Call GetData( BranchHnd, BR_nType, TypeCode )
   Select Case TypeCode
    Case TC_LINE
     TypeString = "L"
     Call GetData( DeviceHnd, LN_sID, sID$ )
    Case TC_XFMR
     TypeString = "T"
     Call GetData( DeviceHnd, XR_sID, sID$ )
    Case TC_PS
     TypeString = "P"
     Call GetData( DeviceHnd, PS_sID, sID$ )
    Case TC_XFMR3
     TypeString = "X"
     Call GetData( DeviceHnd, X3_sID, sID$ )
    Case Else
     TypeString = ""
     sID$ = ""
   End Select
   StringVal$ = FullBusName( Bus1Hnd ) & "-" & FullBusName( Bus2Hnd ) & " " & sID & TypeString 

  ' Show it
  Print StringVal$ & "=" & Format( fastestTime, "#0.00")
  ' Print it
'  Print #1, StringVal$ & "," & Format( fastestTime, "#0.00")

  exit Function
  
HasError:
  Print ErrorString()

End Function

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
  countRlyGrp = 0
  While nBusTodo > 0
    nIdx = countBus - nBusTodo + 1
    Bus1Hnd  = BusToDo(nIdx)
    nTier    = BusTier(nIdx)
    nBusTodo = nBusTodo - 1
    BranchHnd   = 0
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
          
'          Call GetData( BranchHnd, BR_nBus1Hnd, Bus1Hnd )
'          Call GetData( BranchHnd, BR_nBus2Hnd, Bus2Hnd )
'          Call GetData( BranchHnd, BR_nHandle, DeviceHnd )
'          StringVal$ = FullBusName( Bus1Hnd ) & "-" & FullBusName( Bus2Hnd ) & " " & sID & TypeString
          ' Show it
'          Print countRlyGrp, StringVal$ 
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


