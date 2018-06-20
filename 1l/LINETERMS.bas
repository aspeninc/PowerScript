' ASPEN PowerScript sample program
'
' LINETERMS.BAS
'
' Find substation relay groups on line terminals
'
' Version 1.0
' Category: OneLiner
'
'
Const DEBUGPRINT = 0
Const MXSUBS = 200

Type TLine
 nRlyGrp1Hnd As long   ' nRlyGrp1Hnd < nRlyGrp2Hnd
 nRlyGrp1Hnd as long
End Type

Sub Main()
  dim SubsList(MXSUBS) as TLine
  dim nSubs as long
  nSubs = 0
  
  bDoOneGroup& = False
  PickedHnd& = 0
  If GetEquipment( TC_PICKED, PickedHnd& ) = 0  Or _
     EquipmentType( PickedHnd ) <> TC_RLYGROUP Then
   PickedHnd& = 0
  End If
  If PickedHnd& <> 0 Then
    Call GetData( PickedHnd, RG_nBranchHnd, nBranchHnd& )
    Call GetData( nBranchHnd, BR_nType, nBrType& )
    If nBrType <> TC_LINE Then PickedHnd = 0
  End If
  If PickedHnd > 0 Then 
    bDoOneGroup& = True
  Else 
    InputStr$ = InputBox( "Study kV level (Enter bank to cancel)", "Scope" )
    dKVWanted# = 0
    If IsNumeric(InputStr) Then 
      dKVWanted# = Val(InputStr)
    End If
    If dKVWanted = 0 Then Stop
  End If

  dim nTermHnds(5) As long
  While True
    if not bDoOneGroup then
      ' Look for the next relay group
      bFound = false
      While GetEquipment( TC_RLYGROUP, PickedHnd ) > 0
       Call GetData( PickedHnd, RG_nBranchHnd, nBranchHnd& )
       Call GetData( PickedHnd, RG_nInService, nInService& )
       If nInService = 1 Then
         Call GetData( nBranchHnd, BR_nType, nType& )
         If nType = TC_LINE Then
           Call GetData( nBranchHnd, BR_nBus1Hnd, nBus1Hnd& )
           Call GetData( nBus1Hnd, BUS_dKvNominal, dKV# )
           If Abs(dKV-dKVWanted) < 0.0001 Then
              If DEBUGPRINT = 1 then
                sText$ = "Found " & FullBranchString(PickedHnd) 
                printTTY( sText )
              End If
            bFound& = True
            exit For
           End If
         End If
       End If
      Wend
      If Not bFound Then exit For
    end if
    nTermHnds(1) = PickedHnd
    nNOTerms& = FindLineTerms( nTermHnds )
    sText$ = "Local group: " & FullBranchString( PickedHnd )
    printTTY(sText)
    For ii& = 2 to nNOTerms
     nThisEnd& = nTermHnds(ii)
     sText = "  Remote group:" & FullBranchString( nThisEnd )
     printTTY(sText)
    Next 
    If bDoOneGroup Then exit For
  Wend

  Exit Sub
HasError:
   Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()

Const MXNEIBORS = 100
Const MXTIERS = 2
Function FindLineTerms( ByRef nHndTerms() As long ) As long
  ' Traverse the line to find remote end relay groups
  dim BrHndsTodo(MXNEIBORS/2) As long
  dim BrHndsTodoTier(MXNEIBORS/2) As long
  dim nBrsTodo As long
  dim BusHndsSeen(MXNEIBORS) As long
  dim nBusesSeen As long
  dim BrHnds(2) As long

  nBusesSeen = 0
  nFound& = 0
  nTermHnds& = nHndTerms(1)
  Call GetData( nTermHnds, RG_nBranchHnd, nBranchHnd1& )
  nBrsTodo = 1
  BrHndsTodo(nBrsTodo) = nBranchHnd1
  BrHndsTodoTier(nBrsTodo) = 0
  While nBrsTodo > 0 
    nBranchHnd1& = BrHndsTodo(nBrsTodo)
    nThisTier& = BrHndsTodoTier(nBrsTodo)
    nBrsTodo = nBrsTodo - 1 
    If DEBUGPRINT And 1 Then
      sText$ = "Near end branch (tier = " & Str(nThisTier) & "): " & FullBranchString(nBranchHnd1)
      PrintTTY( sText )
    End If
    Call GetData( nBranchHnd1, BR_nBus1Hnd, nBus1Hnd& )
    idx& = FindInList( nBus1Hnd, BusHndsSeen, nBusesSeen )
    If idx > nBusesSeen Then
      If nBusesSeen = MXNEIBORS Then
        Print "MXNEIBORS is exceeded. Cannot continue"
        Stop
      End If
      nBusesSeen = nBusesSeen + 1
      BusHndsSeen(nBusesSeen) = nBus1Hnd
    End If
    BrHnds(1) = nBranchHnd1
    nRemotes& = OppositeBranchHnds( BrHnds )
    For ii = 1 to nRemotes
      ' One or two remote branches seen from this end
      nRemoteBrHnd = BrHnds(ii)
      If DEBUGPRINT And 1 Then
        sText$ = "  Looking at: " & FullBranchString(nRemoteBrHnd)
        printTTY( sText )
      End If
      nRlyGrp1Hnd& = 0
      Call GetData( nRemoteBrHnd, BR_nRlyGrp1Hnd, nRlyGrp1Hnd& )
      If nRlyGrp1Hnd > 0 Then
        nFound = nFound + 1
        nHndTerms(1+nFound) = nRlyGrp1Hnd
        If DEBUGPRINT And 1 Then
          sText$ = "***Found terminal: " & FullBranchString(nRlyGrp1Hnd)
          printTTY( sText )
        End If
      Else
        ' No relay group here. Must go further if possible
        If nThisTier < MXTIERS And nBrsTodo < MXNEIBORS/2 - 1 Then
          Call GetData( nRemoteBrHnd, BR_nBus1Hnd, nRemoteBus& )
          nABrHnd& = 0
          While GetBusEquipment( nRemoteBus, TC_BRANCH, nABrHnd ) > 0
            If nABrHnd <> nRemoteBrHnd Then
              Call GetData( nABrHnd, BR_nType, nType& )
              If nType = TC_LINE Then
                Call GetData( nABrHnd, BR_nInService, nInService& )
                If nInService = 1 Then
                  Call GetData( nABrHnd, BR_nBus2Hnd, nBus2Hnd& )
                  idx& = FindInList( nBus2Hnd, BusHndsSeen, nBusesSeen )
                  If idx > nBusesSeen Then
                    ' Remote bus was not encountered before
                    nBrsTodo = nBrsTodo + 1
                    BrHndsTodo(nBrsTodo) = nABrHnd
                    Call GetData( nBus2Hnd, BUS_nTapBus, nTapBus& )
                    If nTapBus = 1 Then
                      BrHndsTodoTier(nBrsTodo) = nThisTier + 1
                    Else
                      BrHndsTodoTier(nBrsTodo) = nThisTier
                    End If
                    If DEBUGPRINT And 1 Then
                      sText$ = "  Added to up next: " & FullBranchString(nABrHnd)
                      printTTY( sText )
                    End If
                  End If  'If idx > nBusesSeen
                End If  'If nInService = 1
              End If  'If nType = TC_LINE
            End If  'If nABrHnd <> nRemoteBrHnd
          Wend
        End If
      End If
    Next
  Wend  'While nBrsTodo > 0
  FindLineTerms = 1+nFound
End Function


Function FindInList( ByVal nVal&, ByRef List() As long, ByVal nLen ) As long
  For ii& = 1 to nLen
   If List(ii) = nVal Then 
     FindInList = ii
     exit Function
   End If
  Next
  FindInList = ii
End Function

Function OppositeBranchHnds( ByRef nBrHnds() As long ) As long
  nNearHnd& = nBrHnds(1)
  Call GetData( nNearHnd, BR_nHandle, nHandle& )  ' Handle of branch's equipment for FindBusBranch()
  Call GetData( nNearHnd, BR_nBus2Hnd, nFarBus1Hnd& )
  nBrHnds(1) = FindBusBranch( nFarBus1Hnd, nHandle )
  OppositeBranchHnds = 1
  Call GetData( nNearHnd, BR_nBus3Hnd, nFarBus2Hnd& )
  If nFarBus2Hnd <> 0 Then 
    nBrHnds(1) = FindBusBranch( nFarBus2Hnd, nHandle )
    OppositeBranchHnds = 2
  End If
End Function

Function FindBusBranch( nBusHnd&, nDevHnd& ) As long
 FindBusBranch = 0
 nBrHnd& = 0
 While GetBusEquipment( nBusHnd, TC_BRANCH, nBrHnd ) > 0
  Call GetData( nBrHnd, BR_nHandle, nHandle& )
  If nHandle = nDevHnd Then
    FindBusBranch = nBrHnd
    exit Sub
  End If
 Wend
End Function

Function FullBranchString( ByVal nHnd& ) As String
  nBranchHnd& = 0
  nType = EquipmentType( nHnd )
  If nType = TC_RLYGROUP Then
    Call GetData( nHnd, RG_nBranchHnd, nBranchHnd& )
  Else 
    If nType = TC_BRANCH Then nBranchHnd& = nHnd
  End If
  FullBranchString = ""
  If nBranchHnd = 0 Then exit Sub
  Call GetData( nBranchHnd, BR_nBus1Hnd, nBus1Hnd& )
  Call GetData( nBranchHnd, BR_nBus2Hnd, nBus2Hnd& )
  Call GetData( nBranchHnd, BR_nType, nBrType& )
  Call GetData( nBranchHnd, BR_nHandle, nHandle& )
  sType$ = ""
  select case nBrType
   case TC_LINE
    sType$ = "L"
    Call GetData( nHandle, LN_sID, sID$ )
   case TC_XFMR
    sType$ = "T"
    Call GetData( nHandle, XR_sID, sID$ )
   case TC_XFMR3
    sType$ = "X"
    Call GetData( nHandle, X3_sID, sID$ )
   case TC_PS
    sType$ = "P"
    Call GetData( nHandle, PS_sID, sID$ )
  End select
  If sType = "" Then exit Sub
  FullBranchString = FullBusName(nBus1Hnd) & "-" & FullBusName(nBus2Hnd) & " " & sID & sType
End Function
