' ASPEN PowerScrip sample program
'
' LineFaultProfile3.BAS
'
' Perform intermediate faults on all lines
' to find locations (in percent and mile) where fault current 
' exceed given threshold
'
' Usage: Follow screen instruction
'
' Version 1.0
' Category: OneLiner
'
'
'   If dKV1 < 500 Then
'    vdLimits(1) = 42000
'    vdLimits(2) = 64000
'    vdLimits(3) = 79000
'   Else
'    vdLimits(1) = 26000
'    vdLimits(2) = 49000
'    vdLimits(3) = 64000
'   End If
'
'

dim nCountBuses As long
dim nCountTabBuses As long
dim tabBuseHnds() As long
dim tabBuses() As long

Sub main()
   
      
      ' Specify output file
   Print "Please select or create an excel file for report output"
   ExcelFile$ = FileSaveDialog( "", "Excel File (*.csv)||", "csv", 2+16 )
   If Len(ExcelFile) = 0 Then exit Sub
   Open ExcelFile For Output As 1
   
   Call ResetTapBuses()

   Print #1, "Line Fault Profile"
   Print #1, ",>500kV,<=500kV"
'   Print #1, "Level 1,26kA,42kA"
'   Print #1, "Level 2,49kA,64kA"
'   Print #1, "Level 3,64kA,79kA"
   Print #1, "Level 1,26kA,10kA"
   Print #1, "Level 2,49kA,15kA"
   Print #1, "Level 3,64kA,25kA"
   Print #1, ""
   Print #1, "Area, Zone, Bus1, Bno1, Bus2, Bno2, kV, ID, >= Level 3,,< Level 3 & >= Level 2,, < Level 2 & >= Level 1,, < Level 1"
   ' Count total of lines
   LineCount = 0
   DevHandle& = 0
   While GetEquipment( TC_LINE, DevHandle ) > 0
     LineCount = LineCount + 1
   Wend
   Counts&   = 0
   DevHandle& = 0
   jj = 0
   While GetEquipment( TC_BUS, DevHandle ) > 0
     jj = jj + 1
     Button =  ProgressDialog( 1, "Computing line faults", "Bus " + Str(jj) +" of " + Str(nCountBuses), 100 * jj / nCountBuses )
     If Button = 2 Then 
       Print "Cancel button pressed"
       GoTo Done
     End If
     
'     Call GetData( DevHandle, BUS_sName, sBusName$ )
'     Print sBusName
     
     Call GetData( DevHandle, BUS_nTapBus, nTap& )
     If nTap <> 0 Then GoTo next1
     
     nBrHnd& = 0
     While GetBusEquipment( DevHandle, TC_BRANCH, nBrHnd ) > 0
       Call GetData( nBrHnd, BR_nType, nType& )
       If nType = TC_LINE Then
         Call GetData( nBrHnd, BR_nHandle, nLineHnd& )
         Call GetData( nLineHnd, LN_nInService, nInService& )
         If nInService = 1 Then Call faultOneLine( nBrHnd )
       End If
     Wend
'   GoTo done
   next1:
   Wend
   
   Done:
   
   Call RestoreTapBuses()
   Call ProgressDialog( 0, "", "", 0 )   
   Close 1
   StrOut$ = "The report has been saved to " & ExcelFile
   Print StrOut
   Exit Sub
HasError:
   Call RestoreTapBuses()
   Call ProgressDialog( 0, "", "", 0 )   
   Close 1
   Print "Error: ", ErrorString( )
End Sub

Function RestoreTapBuses() As long
   ' Restore tap buses
   For ii = 1 to nCountTabBuses
     nTapBus = tabBuses(ii)
     nBusHnd = tabBuseHnds(ii)
     Call SetData(nBusHnd, BUS_nTapBus, nTapBus)
     Call PostData(nBusHnd)
   Next
End Function

Function ResetTapBuses() As long
   ' Temporary remove all tap bus attributes
   nCountBuses = 0
   nCountTabBuses& = 0
   nBusHnd& = 0
   While GetEquipment( TC_BUS, nBusHnd ) > 0
     Call GetData( nBusHnd, BUS_nTapBus, nTapBus& )
     If nTapBus <> 0 Then nCountTabBuses = nCountTabBuses + 1
   Wend
   maxsizeTabBuses& = nCountTabBuses + 2
   ReDim tabBuseHnds(maxsizeTabBuses)
   ReDim tabBuses(maxsizeTabBuses)
   nCountTabBuses& = 0
   nBusHnd& = 0
   While GetEquipment( TC_BUS, nBusHnd ) > 0
     nCountBuses = nCountBuses + 1
     Call GetData( nBusHnd, BUS_nTapBus, nTapBus& )
     If nTapBus <> 0 Then
       nCountTabBuses = nCountTabBuses + 1
       tabBuseHnds(nCountTabBuses) = nBusHnd
       tabBuses(nCountTabBuses) = nTapBus
       nTapBus = 0
       nResult = SetData(nBusHnd, BUS_nTapBus, nTapBus)
       nResult = PostData(nBusHnd)
     End If
   Wend
End Function


Function faultOneLine( BranchHnd& ) As long
   faultOneLine = 0
   ' Variable declaration
   Dim FltConn(4) As Long
   Dim FltOption(14) As Double
   Dim OutageList(20) As Long
   dim vdVal1(6) As double
   dim vdVal2(6) As double
   dim vdLimits(3) As double
   
   Call getdata(nBrHnd, BR_nBus1Hnd, nBus1Hnd&) 
   Call getdata(nBus1Hnd, BUS_dKVNominal, dKV1#)
   
   If dKV1 < 500 Then
 '   vdLimits(1) = 42000
 '   vdLimits(2) = 64000
 '   vdLimits(3) = 79000
    vdLimits(1) = 10000
    vdLimits(2) = 15000
    vdLimits(3) = 25000
   Else
    vdLimits(1) = 26000
    vdLimits(2) = 49000
    vdLimits(3) = 64000
   End If
   
   sCur$ = printBranchID(BranchHnd)
   Call intermFaultCurr(BranchHnd, 0.1,  dFCurr#) ' Get From End Fault Current
   Call intermFaultCurr(BranchHnd, 99.9, dTCurr#) ' Get To End Fault Current
   
   Call GetData( BranchHnd, BR_nHandle, nLineHnd& )
   Call GetData( nLineHnd, LN_dLength, dLength# )
   Call GetData( nLineHnd, LN_sLengthUnit, sUnit$ )
   
   If dFCurr > vdLimits(3) Then
     If dTCurr > vdLimits(3) Then
     ' No Fault Test (All above hightest limit)
       sLoc$ = getLoc( dLength, 0.0, 100.0, sUnit )
       sCur$ = sCur$ & "," & "0.00% - 100.00%," & sLoc & "," & "N/A" & ",," & "N/A" & ",," & "N/A"
     ElseIf dTCurr > vdLimits(2) Then
     ' Two Sections: No Fault Test + vdLimits(3) 
       Call locateFault(BranchHnd, 0.1, 99.9, vdLimits(3), dPerCent# )
       sLoc$ = getLoc( dLength, 0.0, dPerCent, sUnit )
       sCur$ = sCur$ & "," & "0.00% - " & Format( dPerCent, "#0.#0" ) & "%," & sLoc
       sLoc$ = getLoc( dLength, dPerCent, 100.0, sUnit )
       sCur$ = sCur$ & "," & Format( dPerCent, "#0.#0" ) & "% - 100.00%," & sLoc & "," & "N/A" & ",," & "N/A"
     ElseIf dTCurr > vdLimits(1) Then
     ' Three Secions: No Fault Test + vdLimits(3) + vdLimits(2)
       Call locateFault(BranchHnd, 0.1, 99.9, vdLimits(3), dPerCent# )
       sLoc$ = getLoc( dLength, 0.0, dPerCent, sUnit )
       sCur$ = sCur$ & "," & "0.00% - " & Format( dPerCent, "#0.#0" ) & "%," & sLoc
       dTemp = dPerCent
       Call locateFault(BranchHnd, dTemp, 99.9, vdLimits(2), dPerCent# )
       sLoc$ = getLoc( dLength, dTemp, dPerCent, sUnit )
       sCur$ = sCur$ & "," & Format( dTemp, "#0.#0" ) & "% - " & Format( dPerCent, "#0.#0" ) & "%," & sLoc
       sLoc$ = getLoc( dLength, dPerCent, 100.0, sUnit ) 
       sCur$ = sCur$ & "," & Format( dPerCent, "#0.#0" ) & "% - 100.00%," & sLoc & "," & "N/A"  
     Else
     ' Four Secions: No Fault Test + vdLimits(3) + vdLimits(2) + vdLimits(1)
       Call locateFault(BranchHnd, 0.1, 99.9, vdLimits(3), dPerCent# )
       sLoc$ = getLoc( dLength, 0.0, dPerCent, sUnit )
       sCur$ = sCur$ & "," & "0.00% - " & Format( dPerCent, "#0.#0" ) & "%," & sLoc
       dTemp = dPerCent
       Call locateFault(BranchHnd, dTemp, 99.9, vdLimits(2), dPerCent# )
       sLoc$ = getLoc( dLength, dTemp, dPerCent, sUnit )
       sCur$ = sCur$ & "," & Format( dTemp, "#0.#0" ) & "% - " & Format( dPerCent, "#0.#0" ) & "%," & sLoc
       dTemp = dPerCent
       Call locateFault(BranchHnd, dTemp, 99.9, vdLimits(1), dPerCent# )
       sLoc$ = getLoc( dLength, dTemp, dPerCent, sUnit )
       sCur$ = sCur$ & "," & Format( dTemp, "#0.#0" ) & "% - " & Format( dPerCent, "#0.#0" ) & "%," & sLoc
       sLoc$ = getLoc( dLength, dPerCent, 100.0, sUnit ) 
       sCur$ = sCur$ & "," & Format( dPerCent, "#0.#0" ) & "% - 100.00%," & sLoc
     End If
   ElseIf dFCurr > vdLimits(2) Then
     If dTCurr > vdLimits(2) Then
     ' One Section: vdLimits(3)
       sLoc$ = getLoc( dLength, 0.0, 100.0, sUnit )
       sCur$ = sCur$ & "," & "N/A" & ",," & "0.00% - 100.00%," & sLoc & "," & "N/A" & ",," & "N/A"
     ElseIf dTCurr > vdLimits(1) Then
     ' Two Sections: vdLimits(3) + vdLimits(2)
       Call locateFault(BranchHnd, 0.1, 99.9, vdLimits(2), dPerCent# )
       sLoc$ = getLoc( dLength, 0.0, dPerCent, sUnit )
       sCur$ = sCur$ & "," & "N/A" & ",," & "0.00% - " & Format( dPerCent, "#0.#0" ) & "%," & sLoc
       sLoc$ = getLoc( dLength, dPerCent, 100.0, sUnit ) 
       sCur$ = sCur$ & "," & Format( dPerCent, "#0.#0" ) & "% - 100.00%," & sLoc & "," & "N/A"
     Else
     ' Three Sections: vdLimits(3) + vdLimits(2) + vdLimits(1)
       Call locateFault(BranchHnd, 0.1, 99.9, vdLimits(2), dPerCent# )
       sLoc$ = getLoc( dLength, 0.0, dPerCent, sUnit )
       sCur$ = sCur$ & "," & "N/A" & ",," & "0.00% - " & Format( dPerCent, "#0.#0" ) & "%," & sLoc
       dTemp = dPerCent
       Call locateFault(BranchHnd, dTemp, 99.9, vdLimits(1), dPerCent# )
       sLoc$ = getLoc( dLength, dTemp, dPerCent, sUnit )
       sCur$ = sCur$ & "," & Format( dTemp, "#0.#0" ) & "% - " & Format( dPerCent, "#0.#0" ) & "%," & sLoc
       sLoc$ = getLoc( dLength, dPerCent, 100.0, sUnit ) 
       sCur$ = sCur$ & "," & Format( dPerCent, "#0.#0" ) & "% - 100.00%," & sLoc    
     End If  
   ElseIf dFCurr > vdLimits(1) Then
     If dTCurr > vdLimits(1) Then
     ' One Section: vdLimits(2)
       sLoc$ = getLoc( dLength, 0.0, 100.0, sUnit )
       sCur$ = sCur$  & "," & "N/A" & ",," & "N/A" & ",," & "0.00% - 100.00%," & sLoc & "," & "N/A" 
     Else
     ' Two Sections: vdLimits(2) + vdLimits(1)
       Call locateFault(BranchHnd, 0.1, 99.9, vdLimits(1), dPerCent# )
       sLoc$ = getLoc( dLength, 0.0, dPerCent, sUnit )
       sCur$ = sCur$  & "," & "N/A" & ",," & "N/A" & ",," & "0.00% - " & Format( dPerCent, "#0.#0" ) & "%," & sLoc
       sLoc$ = getLoc( dLength, dPerCent, 100.0, sUnit ) 
       sCur$ = sCur$ & "," & Format( dPerCent, "#0.#0" ) & "% - 100.00%," & sLoc 
     End If 
   Else
     ' One Section: vdLimits(1)
     sLoc$ = getLoc( dLength, 0.0, 100.0, sUnit )
     sCur$ = sCur$ & "," & "N/A" & ",," & "N/A" & ",," & "N/A" & ",," & "0.00% - 100.00%" & sLoc
   End If   
   Print #1, sCur

   faultOneLine = 1
   Exit Function
HasError:
   Print "Error: ", ErrorString( )
    
   
End Function

Function intermFaultCurr( BranchHnd&, dPcnt#, ByRef dCur# ) As long
   intermFaultCurr = 0
   ' Variable declaration
   Dim FltConn(4) As Long
   Dim FltOption(14) As Double
   Dim OutageList(20) As Long
   dim vdVal1(6) As double
   dim vdVal2(6) As double
   
   For ii = 1 To 4 
     FltConn(ii) = 0
   Next 
   For ii = 1 To 12
     FltOption(ii) = 0.0
   Next   

   ' Fault connection
   FltConn(1)    = 1	' 3LG 
   FltConn(3)    = 1	' 1LG 
   
   FltOption(11)  = dPcnt   ' Intermediate with end-open
   ' Simulate fault
   If 0 = DoFault( BranchHnd, FltConn, FltOption, OutageType, OutageList, 0.0, 0.0, 1 ) Then GoTo HasError1
   FaultFlag = 1
   While PickFault( FaultFlag ) <> 0
     If GetSCCurrent( HND_SC, vdVal1, vdVal2, 4 ) = 0 Then GoTo HasError1
     dCur = vdVal1(1) ' Phase A is max
     If pickFault( SF_NEXT ) = 0 Then GoTo HasError1 
     If GetSCCurrent( HND_SC, vdVal1, vdVal2, 4 ) = 0 Then GoTo HasError1
     If dCur < vdVal1(1) Then dCur = vdVal1(1) ' Phase A is max
     FaultFlag = SF_NEXT	' Check next fault   
   Wend
   intermFaultCurr = 1
   Exit Function
HasError1:
   Print "Error: ", ErrorString( )
   Stop
End Function

Function locateFault( BranchHnd&, dPcnt1, dPcnt2, dLimit, ByRef dPCnt# ) As long
   locateFault = 0
   ' Variable declaration
   Dim FltConn(4) As Long
   Dim FltOption(14) As Double
   Dim OutageList(20) As Long
   dim vdVal1(6) As double
   dim vdVal2(6) As double
   
   dFPcnt = dPcnt1
   dTPcnt = dPcnt2
   For ii = 1 To 4 
     FltConn(ii) = 0
   Next 
   For ii = 1 To 12
     FltOption(ii) = 0.0
   Next   

   ' Fault connection
   FltConn(1)    = 1	 ' 3LG 
   FltConn(3)    = 1	 ' 1LG 
   dPCnt = (dFPcnt + dTPcnt) / 2.0
   FltOption(11) = dPCnt ' Intermediate with end-open
   If 0 = DoFault( BranchHnd, FltConn, FltOption, OutageType, OutageList, 0.0, 0.0, 1 ) Then GoTo HasError1
   FaultFlag = 1
   While PickFault( FaultFlag ) <> 0
     If GetSCCurrent( HND_SC, vdVal1, vdVal2, 4 ) = 0 Then GoTo HasError1
     dCur = vdVal1(1) ' Phase A is max
     If pickFault( SF_NEXT ) = 0 Then GoTo HasError1 
     If GetSCCurrent( HND_SC, vdVal1, vdVal2, 4 ) = 0 Then GoTo HasError1
     If dCur < vdVal1(1) Then dCur = vdVal1(1) ' Phase A is max
     FaultFlag = SF_NEXT	' Check next fault   
   Wend
   
   While Abs(dCur - dLimit) > 1
     If dCur > dLimit Then
       dFPcnt = dPCnt
       dPCnt = (dPCnt + dTPcnt) / 2.0
       FltOption(11) = dPCnt ' Intermediate with end-open
     Else 
       dTPcnt = dPcnt
       dPCnt = (dPCnt + dFPcnt) / 2.0
       FltOption(11) = dPCnt ' Intermediate with end-open   
     End If
     If 0 = DoFault( BranchHnd, FltConn, FltOption, OutageType, OutageList, 0.0, 0.0, 1 ) Then GoTo HasError1
     FaultFlag = 1
     While PickFault( FaultFlag ) <> 0
       If GetSCCurrent( HND_SC, vdVal1, vdVal2, 4 ) = 0 Then GoTo HasError1
       dCur = vdVal1(1) ' Phase A is max
       If pickFault( SF_NEXT ) = 0 Then GoTo HasError1 
       If GetSCCurrent( HND_SC, vdVal1, vdVal2, 4 ) = 0 Then GoTo HasError1
       If dCur < vdVal1(1) Then dCur = vdVal1(1) ' Phase A is max
       FaultFlag = SF_NEXT	' Check next fault   
     Wend
   Wend
   locateFault = 1
   Exit Function	
HasError1:
   Print "Error: ", ErrorString( )
   Stop
End Function

Function printBranchID( nBrHnd& ) As String
  Call getdata(nBrHnd, BR_nBus1Hnd, nBus1Hnd&) 
  Call getdata(nBrHnd, BR_nBus2Hnd, nBus2Hnd&)
  Call getdata(nBus1Hnd, Bus_nArea, nArea&)
  Call getdata(nBus1Hnd, Bus_nZone, nZone&) 
  Call getdata(nBus1Hnd, BUS_sName, sName1$)
  Call getdata(nBus1Hnd, BUS_nNumber, nNumber1&)
  Call getdata(nBus1Hnd, BUS_dKVNominal, dKV1#)
  Call getdata(nBus2Hnd, BUS_sName, sName2$)
  Call getdata(nBus2Hnd, BUS_nNumber, nNumber2&)
  Call getdata(nBus2Hnd, BUS_dKVNominal, dKV2#)
  Call getdata(nBrHnd, BR_nHandle, nHandle&)
  Call getdata(nHandle, LN_sID, sID$)
  printBranchID = nArea & "," & nZone & "," & sName1 & "," & nNumber1 & "," & sName2 & "," & nNumber2 & "," & dKV1 & "," & sID
End Function

Function findLineBranch( lineHnd&, side& ) As long
 findLineBranch = 0
 If side = 1 Then nCode& = LN_nBus1Hnd Else nCode& = LN_nBus2Hnd
 Call GetData( lineHnd, nCode, nBusHnd& )
 nBrHnd& = 0
 While GetBusEquipment( nBusHnd, TC_BRANCH, nBrHnd ) > 0
   Call GetData( nBrHnd, BR_nHandle, nThisHnd& )
   If nThisHnd = lineHnd Then
     findLineBranch = nBrHnd
     exit Function
   End If
 Wend 
End Function

Function getLoc( dLineLength, dFromPct, dToPct, sUnit ) As String
  If dLineLength > 0 Then
    dFrom = dFromPct*dLineLength/100
    dTo   = dToPct*dLineLength/100
    getLoc= Format( dFrom, "#0.#0" ) & " - " & Format( dTo, "#0.#0" ) & " " & sUnit
  Else
    getLoc = ""
  End If
End Function

