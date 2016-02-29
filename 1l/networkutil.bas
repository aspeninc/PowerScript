' ASPEN PowerScript sample program
'
'NETWORKUTIL.BAS
'
Sub main
  
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' Test GetRemoteTerminals()
  If 0 = GetEquipment( TC_PICKED, PickedHnd& ) Or _
     EquipmentType( PickedHnd ) <> TC_RLYGROUP Then
    Print "Please select a relay group on transmission line"
    exit Sub
  End If

  Call GetData( PickedHnd, RG_nBranchHnd, branchHnd& )
  
  dim TerminalList(50) As long
  
  nCount = GetRemoteTerminals( branchHnd, TerminalList )

  Call GetData( branchHnd, BR_nBus1Hnd, nBusHnd& )
  sText = "Local: " & FullBusName(nBusHnd) & "; remote: "
  For ii = 1 to nCount
    branchHnd = TerminalList(ii)
    Call GetData( branchHnd, BR_nBus1Hnd, nBusHnd& )
    sText = sText & " " & FullBusName(nBusHnd)
  Next

  Print sText
  ' End Test GetRemoteTerminals()
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


End Sub


'=============================================================================================================
' GetRemoteTerminals()
'
' Purpose: Find all remote ends of a line. All taps are ignored. Close switches are included
' Usage: 
'   BranchHnd [in] branch handle of the local terminal
'   TermsHnd  [in] array to hold list of branch handle at remote ends
' Return: Number of remote ends  
'
Const MXSIZE = 100
dim TempBrList(MXSIZE) As long, TempListSize As long  ' We use this list to store branches at tap buses
dim TempList2(MXSIZE) As long, TempListSize2 As long  ' We use this list to store tap buses

Function GetRemoteTerminals( ByVal BranchHnd&, ByRef TermsHnd() As long ) As long
  ListSize&     = 0
  TempListSize2 = 0
  TempListSize  = 1
  TempBrList(TempListSize) = BranchHnd  
  While TempListSize  > 0
    NearEndBrHnd& = TempBrList(TempListSize)
    TempListSize = TempListSize  - 1
    Call FindOppositeBranch( NearEndBrHnd, TermsHnd, ListSize )
  Wend
  GetRemoteTerminals = ListSize
End Function

Function FindOppositeBranch( ByVal NearEndBrHnd&, ByRef OppositeBrList() As long, ByRef ListSize& ) As long
  FindOppositeBranch = 0
  Call GetData( NearEndBrHnd, BR_nInService, nStatus& )
  If nStatus <> 1 Then exit Function
  
  Call GetData( NearEndBrHnd, BR_nBus2Hnd, nBus2Hnd& )
  ' Check if we had encountered this bus before
  For ii = 1 to TempListSize2  
    If TempList2(ii) = nBus2Hnd Then exit Function
  Next
  If TempListSize2 = MXSIZE Then 
    Print "Ran out of buffer spase. Edit script code to incread MXSIZE" 
    Stop
  End If
  TempListSize2 = TempListSize2 + 1
  TempList2(TempListSize2) = nBus2Hnd
  
  Call GetData( NearEndBrHnd, BR_nHandle,  nThisLineHnd& )
  Call GetData( nBus2Hnd,     BUS_nTapBus, nTapBus& )
  nBranchHnd& = 0
  If nTapBus <> 1 Then
    While GetBusEquipment( nBus2Hnd, TC_BRANCH, nBranchHnd ) > 0
      Call GetData( nBranchHnd, BR_nHandle, nLineHnd& )
      If nThisLineHnd = nLineHnd Then exit Do
    Wend
    ListSize = ListSize + 1
    OppositeBrList(ListSize) = nBranchHnd
    FindOppositeBranch = 1
  Else
    While GetBusEquipment( nBus2Hnd, TC_BRANCH, nBranchHnd ) > 0
      Call GetData( nBranchHnd, BR_nHandle, nLineHnd& )
      If nThisLineHnd = nLineHnd Then GoTo cont1
      Call GetData( nBranchHnd, BR_nType, nType& )
      If nType <> TC_LINE And nType <> TC_SWITCH Then GoTo cont1
      If nType = TC_SWITCH Then 
        Call GetData( nLineHnd, SW_nInService, nStatus& )
        If nStatus <> 1 Then GoTo cont1
        Call GetData( nLineHnd, SW_nStatus, nStatus& )
        If nStatus <> 1 Then GoTo cont1
      End If
      If nType = TC_LINE Then 
        Call GetData( nLineHnd, LN_nInService, nStatus& )
        If nStatus <> 1 Then GoTo cont1
      End If
      If TempListSize = MXSIZE Then 
        Print "Ran out of buffer spase. Edit script code to incread MXSIZE" 
        Stop
      End If
      TempListSize = TempListSize + 1
      TempBrList(TempListSize) = nBranchHnd  
      FindOppositeBranch = FindOppositeBranch  + 1
    cont1:
    Wend
  End If
End Function
'===============================================================================================================

