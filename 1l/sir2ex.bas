' ASPEN PowerScript Sample Program
'
' SIR2EX.BAS
'
' Compute source to line impedance ratio of selected buses and generate report
'
' !!!!! Warning !!!!!!!!!!
'   The source impedance calculation in this script does not always
'   produce accurate results. See SIR.BAS script for correct SIR calculation.
' !!!!!!!!!!!!!!!!!!!!!!!!

' Version: 1.0
' Category: OneLiner
'
'
Sub main()
   ' Variable declaration
   Dim vnBusHnd1(10) As Long
   Dim vnBusHnd2(10000) As Long
   Dim FltConn(4) As Long
   Dim FltOption(14) As Double
   Dim OutageList(20) As Long
   Dim OutageType(3) As Long
   Dim BrHndList(50) As Long
   Dim ToBusHndList(50) As Long
   Dim BranchList(50) As String, BusList(3) As String
   
   
   ' Initialize DoFault options using dialog data
   For ii = 1 To 4 
     FltConn(ii) = 0
   Next 
   For ii = 1 To 14
     FltOption(ii) = 0.0
   Next
   For ii = 1 To 20
     OutageList(ii) = 0
   Next
   For ii = 1 To 3
     OutageType(ii) = 0
   Next
   
   ' Fault bus selection
   sWindowText$ = "Select buses for SIR calculation"
   nPicked& = BusPicker( sWindowText$, vnBusHnd1, vnBusHnd2 )  
   If nPicked = 0 Then exit Sub 
   ' Output file
   Print "Please select or create an excel file for report output"
   ExcelFile$ = FileSaveDialog( "", "Excel File (*.csv)||", 0 )
   If Len(ExcelFile) = 0 Then exit Sub
   Open ExcelFile For Output As 1
   Print #1, "ASPEN OneLiner Bus SIR Report"
   OlrFile$ = GetOLRFileName()
   StrOut$ = "OLR File: " & OlrFile$
   Print #1, StrOut$
   Print #1, "From Bus,To Bus,Line End Branch,Line Impedance(ohm),Source Impedance(ohm),SIR"
   Button = 0
          
   For ii = 1 To nPicked   
     BusHnd = vnBusHnd2(ii)
     BusName = FullBusName(BusHnd)
     BrHnd&   = 0
     BrCounts = 0
     While GetBusEquipment( BusHnd, TC_BRANCH, BrHnd ) > 0
       Call GetData( BrHnd, BR_nType, BrType& )
       Call GetData( BrHnd, BR_nInService, nFlag& )
       If nFlag = 1 And BrType = TC_LINE Then
         Call GetData( BrHnd, BR_nBus2Hnd, BusHnd2& )
         BrHndList(BrCounts) = BrHnd
         BrHndStr$ = BrHnd
         ToBusHnd$ = FindLineRemoteBus(BrHnd)
         If ToBusHnd <> "-999999" Then
           BrList$ = "[" + BrHndStr$ + "] "
           BrList$ = BrList$ + FullBusName(ToBusHnd$)
           ToBusHndList(BrCounts) = ToBusHnd
           BranchList(BrCounts) = BrList$
           BrCounts = BrCounts + 1
         End If
       End If
     Wend
   
     For jj = 0 To Brcounts - 1
       ' Fault connection
       FltConn(1)    = 1	' 3PH 
       ' Fault type
       FltOption(2)  = 1   ' Bus fault with outage
       FltOption(7)  = 0

       ' Extract handle numbers and prepare the outage list
       StrLine$ = BranchList(jj)
       BranchHnd = BrHndList(jj)
       nPos = InStr( 1, StrLine, "]" )
       OutageList(1) = Val( Mid(StrLine, 2, nPos-2) )
       OutageList(2) = 0
       OutageType(1) = 1	' Outage one at a time

       ' Simulate fault
       If 0 = DoFault( BusHnd, FltConn, FltOption, OutageType, OutageList, 0.0, 0.0, 1 ) Then GoTo HasError
                       
       ' Must alway pick a fault before getting V and I
       If PickFault( 1 ) = 0 Then GoTo HasError
       Call GetData( HND_SC, FT_dRPt, R1t# )
       Call GetData( HND_SC, FT_dXPt, X1t# )
   
       dZs# = Sqr(R1t*R1t + X1t*X1t)
 
       ' Fault connection
       FltConn(1)    = 1	' 3PH 
       ' Fault type
       FltOption(2)  = 0
       FltOption(7)  = 1    ' Close-in fault with end opened
     
       ' Simulate fault
       If 0 = DoFault( BranchHnd, FltConn, FltOption, OutageType, OutageList, 0.0, 0.0, 1 ) Then GoTo HasError
                       
       ' Must alway pick a fault before getting V and I
       If PickFault( 1 ) = 0 Then GoTo HasError
       Call GetData( HND_SC, FT_dRPt, R2t# )
       Call GetData( HND_SC, FT_dXPt, X2t# )

       dZl# = Sqr((R2t-R1t)*(R2t-R1t) + (X2t-X1t)*(X2t-X1t))
   
       ' Print output to excel file
       If (dZs > 1000000) And (dZl > 1000000) Then
         sZs = "N/A"
         sZl = "N/A"
         sSir= "N/A"
       Else
         sZs = Format(dZs#, "0.00")
         sZl = Format(dZl#, "0.00")
         sSir= Format(dZs#/dZl#, "0.000")
       End If
       FltStr = FaultDescription()
       nPos1 = InStr(1,FltStr,": ") + 2
       nPos2 = InStr(1,FltStr,"3LG")
       BranchName = Mid(FltStr, nPos1, nPos2-nPos1)
       StrOut$ = FullBusName(BusHnd) & "," & FullBusName(ToBusHndList(jj)) & "," & BranchName & "," & sZl & "," & sZs & "," & sSir
       Print #1, StrOut     
     Next jj
     Button = ProgressDialog( 1, "Process SIR Data", "Record " & Str(ii) & " of " & Str(nPicked), 100*ii/nPicked )
     If Button = 2 Then 
       Print "Cancel button pressed"
       Call ProgressDialog( 0, "", "", 0 )
       Close 1
       Exit Sub
     End If   
   Next ii
   Call ProgressDialog( 0, "", "", 0 )
   Close 1
   StrOut$ = "The report has been saved to " & ExcelFile
   Print StrOut
   Exit Sub
HasError:
   Print "Error: ", ErrorString( )
   Call ProgressDialog( 0, "", "", 0 )
   Close 1
End Sub

Function FindLineRemoteBus( ByVal Branch1Hnd& ) As String
  nLoop1 = 0
  ' Skip all taps on the line
  Do 
    Call GetData( Branch1Hnd, BR_nHandle, LineHnd )
    Call GetData( LineHnd, LN_sName, LineName )
    Call GetData( Branch1Hnd, BR_nBus1Hnd, BusHnd )
    Call GetData( Branch1Hnd, BR_nBus2Hnd, Bus1Hnd )
    Call GetData( Bus1Hnd, BUS_nTapBus, TapCode )
    If TapCode = 0 Then Exit Do			' real bus
    ' Only for tap bus
    Branch1Hnd& = 0
    ttt = GetBusEquipment( Bus1Hnd, TC_BRANCH, Branch1Hnd& )
    nLoop2 = 0
    While ttt <> 0
      Call GetData( Branch1Hnd, BR_nBus2Hnd, Bus2Hnd )	' Get the far end bus
      If Bus2Hnd <> BusHnd Then	' for different branch
        Call GetData( Branch1Hnd, BR_nType, TypeCode )	' Get branch type
        Call GetData( Branch1Hnd, BR_nInService, nFlag& )
        If nFlag = 1 And TypeCode = TC_LINE Then 
          ' Get line name
          Call GetData( Branch1Hnd, BR_nHandle, LineHnd )
          Call GetData( LineHnd, LN_sName, StringVal )
          If StringVal = LineName Then GoTo ExitWhile		' can go further on line with same name
          ttt = GetBusEquipment( Bus1Hnd, TC_BRANCH, Branch1Hnd )
          If ttt = -1 Then GoTo ExitWhile		' It is the last line, no choice but further on line
        End If
      Else		' for same branch
        If ttt = -1 Then GoTo ExitLoop		' If the end bus is tap bus, stop
        ttt = GetBusEquipment( Bus1Hnd, TC_BRANCH, Branch1Hnd )
      End If
      nLoop2 = nLoop2 + 1
      If nLoop2 = 4 Then GoTo ExitWhile
    Wend
    ExitWhile:
    BusHnd  = Bus1Hnd
    Bus1Hnd = Bus2Hnd
    nLoop = nLoop + 1
    If nLoop1 = 4 Or nLoop2 = 4 Then
      Bus1Hnd = "-999999"
      Exit Do
    End If	
  Loop While TapCode = 1
   
  ExitLoop:
  
  FindLineRemoteBus = Bus1Hnd
End Function


