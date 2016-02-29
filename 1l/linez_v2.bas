' ASPEN PowerScript Sample Program
'
' LINEZ.BAS
'
' Report total line impedance and length.
' Lines with tap buses are handled correctly
'
' Version 2.0
'
'========================== 
' Change these constants as as needed
Const KVMin = 0        
Const KVMax = 999
Const maxLines  = 10000
'
'==========================
' Do not change this constant
Const hndOffset = 3     
'==========================

' Global var declaration
dim ProcessedHnd(maxLines) As long
dim BusHndList(100) As long, BusListCount As long

Sub main()
   LineCount  = 0
   PickedHnd& = 0
   While GetEquipment( TC_LINE, PickedHnd& ) > 0
     Index = PickedHnd-hndOffset
     If Index >= maxLines Then
       Print "Too many lines in this network. Increase maxLines and try again."
       Stop
     End If
     ProcessedHnd(Index) = 0
   Wend
   
   If GetEquipment( TC_PICKED, PickedHnd& ) > 0 And _
      EquipmentType( PickedHnd ) = TC_LINE Then
      Call compuOneLine( PickedHnd )
      ' Do it for only selected line
      LineCount  = 1
   Else
      ' Do it for all lines withing kV range
      PickedHnd& = 0
      LineCount  = 0
      While GetEquipment( TC_LINE, PickedHnd& ) > 0
        If ProcessedHnd(PickedHnd-hndOffset) = 0 Then
          Call GetData( PickedLineHnd, LN_nBus1Hnd, Bus1Hnd& )
          Call getdata( Bus1Hnd, BUS_dKVNominal,dKV# )
          If dKV >= KVMin And dKV <= KVmax Then
            Call GetData( Bus1Hnd, BUS_nTapBus, TapCode1& )
            Call GetData( PickedLineHnd, LN_nBus2Hnd, Bus2Hnd& )
            Call GetData( Bus2Hnd, BUS_nTapBus, TapCode2& )
            If TapCode1 = 0 Or TapCode2 = 0 Then
              Call compuOneLine( PickedHnd )       ' Want to start from a real bus
              LineCount = LineCount + 1
            End If
          End If
        End If 
      Wend
   End If
 
   Print LineCount, " lines processed. See result in TTY window"
     
End Sub  
   
Function compuOneLine( ByVal PickedLineHnd& )   
   
   ' Get the branch bus handle
   Call GetData( PickedLineHnd, LN_nBus1Hnd, Bus1Hnd& )
   Call GetData( PickedLineHnd, LN_nBus2Hnd, Bus2Hnd& )
   Call GetData( PickedLineHnd, LN_dR, dR# )
   Call GetData( PickedLineHnd, LN_dX, dX# )
   Call GetData( PickedLineHnd, LN_dR0, dR0# )
   Call GetData( PickedLineHnd, LN_dX0, dX0# )
   Call GetData( PickedLineHnd, LN_dLength, dLength# )
   Call GetData( PickedLineHnd, LN_sName, sName$ )
   BusListCount = 2
   If Bus1Hnd > Bus2Hnd Then
     BusHndList(0) = Bus2Hnd
     BusHndList(1) = Bus1Hnd
   Else
     BusHndList(0) = Bus1Hnd
     BusHndList(1) = Bus2Hnd
   End If
   aLine1$ = FullBusName(Bus1Hnd) + " - " + FullBusName(Bus2Hnd) + ": " + _
                      "Z=" + Format(dR#,"0.00000") + "+j" + Format(dX#,"0.00000") + " " + _
                      "Zo=" + Format(dR0#,"0.00000") + "+j" + Format(dX0#,"0.00000") + " " + _
                      "L=" + Format(dLength#,"0.00000")
   PrintTTY(" ")
'   PrintTTY(aLine$)
   ProcessedHnd(PickedLineHnd-hndOffset) = 1
   
   ' find tap segments on Bus1 side
   BusHnd&  = Bus1Hnd
   Do 
     LineHnd = FindTapSegmentAtBus(BusHnd&, sName)
     If LineHnd = 0 Then exit Do
     ProcessedHnd(LineHnd-hndOffset) = 1
     Call GetData( LineHnd, LN_dR, dRn# )
     Call GetData( LineHnd, LN_dX, dXn# )
     Call GetData( LineHnd, LN_dR0, dR0n# )
     Call GetData( LineHnd, LN_dX0, dX0n# )
     Call GetData( LineHnd, LN_dLength, dL# )
     Call GetData( LineHnd, LN_nBus2Hnd, BusFarHnd )  ' Get the far end bus
     If BusFarHnd = BusHnd Then _
       Call GetData( LineHnd, LN_nBus1Hnd, BusFarHnd )  ' Get the far end bus
     dLength = dLength + dL
     dR  = dR  + dRn
     dX  = dX  + dXn
     dR0 = dR0 + dR0n
     dX0 = dX0 + dX0n
     aLine$ = FullBusName(BusHnd) + " - " + FullBusName(BusFarHnd) + ": " + _
                      "Z=" + Format(dRn#,"0.00000") + "+j" + Format(dXn#,"0.00000") + " " + _
                      "Zo=" + Format(dR0n#,"0.00000") + "+j" + Format(dX0n#,"0.00000") + " " + _
                      "L=" + Format(dL#,"0.00000")
     PrintTTY("Segment: " & aLine$)
     ProcessedHnd(LineHnd-hndOffset) = 1
     BusHndList(BusListCount) = BusHnd
     BusListCount = BusListCount+1
     BusHnd  = BusFarHnd
     BusHndList(BusListCount) = BusFarHnd
     BusListCount = BusListCount+1
   Loop

   ' find tap segments on Bus1 side
   BusHnd&  = Bus2Hnd
   Do 
     LineHnd = FindTapSegmentAtBus(BusHnd&, sName)
     If LineHnd = 0 Then exit Do
     ProcessedHnd(LineHnd-hndOffset) = 1
     Call GetData( LineHnd, LN_dR, dRn# )
     Call GetData( LineHnd, LN_dX, dXn# )
     Call GetData( LineHnd, LN_dR0, dR0n# )
     Call GetData( LineHnd, LN_dX0, dX0n# )
     Call GetData( LineHnd, LN_dLength, dL# )
     Call GetData( LineHnd, LN_nBus2Hnd, BusFarHnd )  ' Get the far end bus
     If BusFarHnd = BusHnd Then _
       Call GetData( LineHnd, LN_nBus1Hnd, BusFarHnd )  ' Get the far end bus
     dLength = dLength + dL
     dR  = dR + dRn
     dX  = dX + dXn
     dR0 = dR0 + dR0n
     dX0 = dX0 + dX0n
     aLine$ = FullBusName(BusHnd) + " - " + FullBusName(BusFarHnd) + ": " + _
                      "Z=" + Format(dRn#,"0.00000") + "+j" + Format(dXn#,"0.00000") + " " + _
                      "Zo=" + Format(dR0n#,"0.00000") + "+j" + Format(dX0n#,"0.00000") + " " + _
                      "L=" + Format(dL#,"0.00000")
     PrintTTY("Segment: " & aLine$)
     ProcessedHnd(LineHnd-hndOffset) = 1
     BusHndList(BusListCount) = BusHnd
     BusListCount = BusListCount+1
     BusHnd  = BusFarHnd
     BusHndList(BusListCount) = BusFarHnd
     BusListCount = BusListCount+1
   Loop

   If BusListCount > 2 Then
     PrintTTY("Segment: " & aLine1$)
     ' Find the two real end buses by sorting the bus list and keep
     ' only entries that do not repeat
     Do 
       Changed = 0
       For ii = 0 to BusListCount-2
         If BusHndList(ii) > BusHndList(ii+1) Then
           nTemp& = BusHndList(ii) 
           BusHndList(ii) = BusHndList(ii+1)
           BusHndList(ii+1) = nTemp
           Changed = 1
         End If
       Next
     Loop While(changed > 0)
     For ii = 0 to BusListCount-2  ' Find 
       If BusHndList(ii) = BusHndList(ii+1) Then
         BusHndList(ii)   = 0
         BusHndList(ii+1) = 0
       End If
     Next 
     jj& = 0
     For ii = 0 to BusListCount-1
       If BusHndList(ii) > 0 Then 
         BusHndList(jj) = BusHndList(ii)
         jj = jj + 1
       End If
       If jj = 2 Then GoTo breakFor
     Next 
     breakFor:
     aLine1$ = FullBusName(BusHndList(0)) + " - " + FullBusName(BusHndList(1)) + ": " + _
                      "Z=" + Format(dR,"0.00000") + "+j" + Format(dX,"0.00000") + " " + _
                      "Zo=" + Format(dR0,"0.00000") + "+j" + Format(dX0,"0.00000") + " " + _
                      "L=" + Format(dLength,"0.00000")
   End If
   PrintTTY("Line: " & aLine1$)
End Function

Function FindTapSegmentAtBus( BusHnd&, sName$ ) As long
  FindTapSegmentAtBus = 0
  Call GetData( BusHnd, BUS_nTapBus, TapCode& )
  If TapCode = 0 Then Exit Function
  BranchHnd& = 0
  While GetBusEquipment( BusHnd, TC_BRANCH, BranchHnd& ) > 0
    Call GetData( BranchHnd&, BR_nType, TypeCode )
    If TypeCode <> TC_LINE Then GoTo ContinueWhile
    Call GetData( BranchHnd&, BR_nHandle, LineHnd& )
    If ProcessedHnd(LineHnd-hndOffset) = 1 Then GoTo continueWhile
    Call GetData( LineHnd, LN_sName, sNameThis$ )
    If sNameThis = sName Then 
      FindTapSegmentAtBus = LineHnd
      exit Do
    End If
    If Left(sNameThis,3) = "[T]" Then GoTo ContinueWhile
    Call GetData( LineHnd, LN_sID, sIDThis$ )
    If sIDThis = "T" Then GoTo ContinueWhile
    FindTapSegmentAtBus = LineHnd
    ContinueWhile:
  Wend
End Function
