' ASPEN PowerScript Sample Program
'
' SIR1.BAS
'
' Compute source to line impedance ratio with voltage

Sub main()
   ' Variable declaration
   Dim FltConn(4) As Long
   Dim FltOption(14) As Double
   Dim OutageList(20) As Long
   Dim OutageType(3) As Long
   Dim BranchNameList(50) As String
   Dim LineNameList(50) As String
   Dim RmtBusHnd(50) As Long
   Dim MagArray(3) As Double
   Dim AngArray(3) As Double
   Dim vnBusHnd1(2) As long, vnBusHnd2(200) As long
   Dim BranchCheckList(5000) As String

   ' bus selection
   sWindowText$ = "Select bus to check SIR (200 or fewer)"
   vnBusHnd1(1) = 0
   nPicked& = BusPicker( sWindowText$, vnBusHnd1, vnBusHnd2 )
   If nPicked = 0 Then Exit Sub
      
   ' open output *csv file  
   CsvFile$ = InputBox("Enter output excel file full path name")
   Open CsvFile For Output As 1
   Print #1, "Aspen Onliner Program SIR Check Report"
   Print #1, "" & "," & "" & "," & "" & "," & "" & "," & "3ph fault on Bus 2" & "," & "" & "," & "3ph fault on Bus 1"
   Print #1, "Bus 1" & "," & "Bus 2" & "," & "Branch" & "," & "Line Name" & "," & "Bus 1 Vmag (pu)"  & "," & "Bus 1 SIR" & "," & "Bus 2 Vmag (pu)"  & "," & "Bus 2 SIR"  & "," & "Max SIR"
   
   ' Initialize DoFault options using dialog data
   For ii = 1 To 4 
     FltConn(ii) = 0
   Next 
   For ii = 1 To 12
     FltOption(ii) = 0.0
   Next
   For ii = 1 To 20
     OutageList(ii) = 0
   Next
   For ii = 1 To 3
     OutageType(ii) = 0
   Next
   
   ' Fault connection
   FltConn(1)    = 1	' 3PH
   
    ' Fault option
   FltOption(1)  = 1	' Close-in
   
   CheckIndex = 0 
   
   ' Inventory of all lines at the bus 
   For ii& = 1 to nPicked  
     BusHnd = vnBusHnd2(ii)
     Call GetData( BusHnd, BUS_nTapBus, TapCode )
     While TapCode = 0
       BusName = FullBusName(BusHnd)
       BrHnd&   = 0
       BrCounts = 0
       BrCheck  = 0
       nIndex = CheckIndex
       TapCode = 1
       While GetBusEquipment( BusHnd, TC_BRANCH, BrHnd ) > 0
         Call GetData( BrHnd, BR_nType, BrType& )
         Call GetData( BrHnd, BR_nInService, nFlag& )
         If nFlag = 1 And BrType = TC_LINE Then
           BusHnd2 = FindLineRemoteBus( BrHnd, 3 )
           If BusHnd2 <> "-999999" Then
             Call GetData( BrHnd, BR_nHandle, nHandle& )
             Call GetData( nHandle, LN_sID, sID )
             BranchName = FullBusName(BusHnd) & " - " & FullBusName(BusHnd2) & " " & sID & " " & "L"
             If nIndex = 0 Then
               Call GetData( nHandle, LN_sName, sName )       
               BrCounts = BrCounts + 1
               BranchNameList(BrCounts) = BranchName
               LineNameList(BrCounts) = sName
               RmtBusHnd(BrCounts) = BusHnd2
               CheckIndex = CheckIndex + 1
               BranchName = FullBusName(BusHnd2) & " - " & FullBusName(BusHnd) & " " & sID & " " & "L"
               BranchCheckList(CheckIndex) = BranchName   
             Else
               nCheck = 0
               For jj = 1 to nIndex
                 If StrComp( BranchName, BranchCheckList(jj) ) = 0 Then
                   nCheck = 1
                   BrCheck = 1
                   jj = nIndex + 1
                 End If
               Next
               If nCheck = 0 Then
                 Call GetData( nHandle, LN_sName, sName )       
                 BrCounts = BrCounts + 1
                 BranchNameList(BrCounts) = BranchName
                 LineNameList(BrCounts) = sName
                 RmtBusHnd(BrCounts) = BusHnd2
                 CheckIndex = CheckIndex + 1
                 BranchName = FullBusName(BusHnd2) & " - " & FullBusName(BusHnd) & " " & sID & " " & "L"
                 BranchCheckList(CheckIndex) = BranchName
               End If
             End If  
           Else
             Call GetData( BrHnd, BR_nHandle, nHandle )
             Call GetData( nHandle, LN_sName, LineName )
             Call GetData( BrHnd, BR_nBus1Hnd, BusHnd )
             Call GetData( BrHnd, BR_nBus2Hnd, BusHnd2 )
             Call GetData( BrHnd, BR_nHandle, nHandle& )
             Call GetData( nHandle, LN_sID, sID )
             BranchName = FullBusName(BusHnd) & " - " & FullBusName(BusHnd2) & " " & sID & " " & "L (" & LineName & ")"
             aString = "Can't find remote bus of " & BranchName
             Print aString
	         Close 1
             Exit Sub
           End If             
         End If
       Wend
   
       If BrCounts = 0 And BrCheck = 0 Then
         aString = "No active line found at the bus:" & BusName
         Print aString
       End If
   
       Call GetData( BusHnd, Bus_dKVnorminal, dVBase )
       dVBase = dVBase/Sqr(3.0)

       ' Simulate fault
       For jj = 1 To BrCounts
         If 0 = DoFault( RmtBusHnd(jj), FltConn, FltOption, OutageType, OutageList, 0.0, 0.0, 1 ) Then GoTo HasError                     
         ' Must alway pick a fault before getting V and I
         If PickFault( 1 ) = 0 Then GoTo HasError
         If GetSCVoltage( BusHnd, MagArray, AngArray, 2 ) = 0 Then GoTo HasError
         dMag1 = MagArray(2)/dVBase
         dAng1 = AngArray(2)
         If 0 = DoFault( BusHnd, FltConn, FltOption, OutageType, OutageList, 0.0, 0.0, 1 ) Then GoTo HasError                     
         ' Must alway pick a fault before getting V and I
         If PickFault( 1 ) = 0 Then GoTo HasError
         If GetSCVoltage( RmtBusHnd(jj), MagArray, AngArray, 2 ) = 0 Then GoTo HasError
         dMag2 = MagArray(2)/dVBase
         dAng2 = AngArray(2)
         SIR1 = (1.0 - dMag1)/dMag1
         SIR2 = (1.0 - dMag2)/dMag2
         If SIR1 > SIR2 Then
           maxSIR = SIR1
         Else
           maxSIR = SIR2
         End If 
         aString = BusName & "," & FullBusName(RmtBusHnd(jj)) & "," & BranchNameList(jj) & "," & LineNameList(jj) & "," & Format(dMag1, "0.00") & "," & Format(SIR1, "0.0") & "," & Format(dMag2, "0.00") & "," & Format(SIR2, "0.0")  & "," & Format(maxSIR, "0.0")
         Print #1, aString
       Next
     Wend
   Next
   Close 1
   aString = "The output has been saved to file " & CsvFile$
   Print aString
   Exit Sub
HasError:
   Close 1
   Print "Error: ", ErrorString( )
End Sub

Function FindLineRemoteBus( ByVal Branch1Hnd&, ByVal nMethod ) As String 
  ' Skip all taps on the line
  nLoop1 = 0
  Do 
    nLoop1 = nLoop1 + 1
    If nLoop1 > 10 Then GoTo ExitLoop
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
      nLoop2 = nLoop2 + 1
      If nLoop2 > 10 Then GoTo ExitLoop
      Call GetData( Branch1Hnd, BR_nBus2Hnd, Bus2Hnd )	' Get the far end bus
      If Bus2Hnd <> BusHnd Then	' for different branch
        Call GetData( Branch1Hnd, BR_nType, TypeCode )	' Get branch type
        Call GetData( Branch1Hnd, BR_nInService, nFlag& )
        If nFlag = 1 And TypeCode = TC_LINE Then 
          ' Get line name
          Call GetData( Branch1Hnd, BR_nHandle, LineHnd )
          If nMethod = 1 Then
            Call GetData( LineHnd, LN_sName, StringVal )
            If StringVal = LineName Then GoTo ExitWhile
          End If
          If nMethod = 2 Then
            Call GetData( LineHnd, LN_sName, StringVal )
            nPos1 = InStr( 1, StringVal, "[T]" )
            nPos2 = InStr( 1, StringVal, "[t]" )
            If nPos1 = 0 And nPos2 = 0 Then GoTo ExitWhile
          End If
          If nMethod = 3 Then
            Call GetData( LineHnd, LN_sID, StringVal )
            If StringVal <> "T" Then GoTo ExitWhile
          End If
        End If
        ttt = GetBusEquipment( Bus1Hnd, TC_BRANCH, Branch1Hnd )
        If ttt = -1 Then GoTo ExitWhile      
      Else		' for same branch
        If ttt = -1 Then
          Bus1Hnd = "-999999" 
          GoTo ExitLoop
        End If
        ttt = GetBusEquipment( Bus1Hnd, TC_BRANCH, Branch1Hnd )
      End If
    Wend
    ExitWhile:
    BusHnd  = Bus1Hnd
    Bus1Hnd = Bus2Hnd	
  Loop While TapCode = 1
   
  ExitLoop:
  If nLoop1 > 10 Or nLoop2 > 10 Then
    FindLineRemoteBus = -999999
  Else
    FindLineRemoteBus = Bus1Hnd
  End If
End Function
