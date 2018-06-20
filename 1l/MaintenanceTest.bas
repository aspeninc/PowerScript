' ASPEN PowerScrip sample program
'
' MaintenanceTest.BAS
'
' Run fault simulation on specified transmission line and two adjoining shortest lines
' Steps:
' 1) Select bus A, B, C and D
' 2) Check if branch exist between A and B, B and C, C and D
' 3) Run 11 fault simulations
'
' Version 1.0
' Category: OneLiner
'

Sub Main()

    Dim vnBusHnd1(10) As Long
    Dim vnBusHnd2(10) As Long
	Dim vnFltConn(4) As Long			' Fault connection flags
	Dim LinePercent As Double			' Intermediate %
	Dim vdFltOpt(14) As Double			' Fault options flags
	Dim vnOutageOpt(4) As Long			' Branch outage option flags
	Dim vnOutageLst(30) As Long			' List of handles of branches to be outaged
	Dim dRflt As Double					' Fault resistance, in Ohms
	Dim dXflt As Double					' Fault reactance, in Ohms
	Dim	nClearPrev As Long				' Clear previous result flag
	Dim nStyle As Long					' Current result style
	Dim vdOut1(12) As Double			' Current result magnitude, into equipment terminals
	Dim vdOut2(12) As Double			' Current result angle in degree, into equipment terminals
	Dim Imag(2) As Double				' Saved fault current magnitude values for the four fault types
	Dim Iang(2) As Double				' Saved fault current angle values in degrees for the four fault types
	Dim MaxAns(2) As Double				' The maximum fault current magnitudes for the four types of faults
	Dim MaxPercentAns(2) As Double		' The maximum fault current percentage value for the four types of faults
	Dim MinAns(2) As Double				' The minimum fault current magnitudes for the four types of faults
	Dim MinPercentAns(2) As Double		' The minimum fault current percentage value for the four types of faults
	
	' Bus A selection
	sWindowText$ = "Bus A Selection"
	nPicked = 0
	nPicked& = BusPicker( sWindowText$, vnBusHnd1, vnBusHnd2 )
	If nPicked = 0 Then
	   Print "Must specify Bus A"
	   Exit Sub
	ElseIf nPicked > 1 Then
	   Print "You could only pick one bus"
	   Exit Sub
	Else
	   BusAHnd = vnBusHnd2(1)
	   vnBusHnd2(1) = 0
	End If
	If GetData( BusAHnd&, BUS_sName, sVal$ ) = 0 Then GoTo HasError
	strMsg = "Bus A: " + sVal$ + Chr(13)
	
	' Bus B selection
	sWindowText$ = "Bus B Selection"
	nPicked = 0
	nPicked& = BusPicker( sWindowText$, vnBusHnd1, vnBusHnd2 )
	If nPicked = 0 Then
	   Print "Must specify Bus B"
	   Exit Sub
	ElseIf nPicked > 1 Then
	   Print "You could only pick one bus"
	   Exit Sub
	Else
	   BusBHnd     = vnBusHnd2(1)
	   vnBusHnd2(1) = 0  
	End If	
	If GetData( BusBHnd&, BUS_sName, sVal$ ) = 0 Then GoTo HasError
	strMsg = strMsg + "Bus B: " + sVal$ + Chr(13)

	
	' Bus C selection
    sWindowText$ = "Bus C Selection"
	nPicked = 0
	nPicked& = BusPicker( sWindowText$, vnBusHnd1, vnBusHnd2 )	
	If nPicked = 0 Then
	   Print "Must specify Bus C"
	   Exit Sub
	ElseIf nPicked > 1 Then
	   Print "You could only pick one bus"
	   Exit Sub
	Else
	   BusCHnd = vnBusHnd2(1)
	   vnBusHnd2(1) = 0
	End If
	If GetData( BusCHnd&, BUS_sName, sVal$ ) = 0 Then GoTo HasError
	strMsg = strMsg + "Bus C: " + sVal$ + Chr(13) 
	
	' Bus D selection
    sWindowText$ = "Bus D Selection"
	nPicked = 0
	nPicked& = BusPicker( sWindowText$, vnBusHnd1, vnBusHnd2 )
	If nPicked = 0 Then
	   Print "Must specify Bus D"
	   Exit Sub
	ElseIf nPicked > 1 Then
	   Print "You could only pick one bus"
	   Exit Sub
	Else
	   BusDHnd = vnBusHnd2(1)
	   vnBusHnd2(1) = 0
	End If
    If GetData( BusDHnd&, BUS_sName, sVal$ ) = 0 Then GoTo HasError
	strMsg = strMsg + "Bus D: " + sVal$ + chr(13)
	
	' Find Branch A-B
    BranchABHnd = branchSearch(BusAHnd&, BusBHnd&)
    If BranchABHnd = 0 Then
       Print "Error: Cant't find branch between bus A and B"
	   	GoTo ErrorFindBranch
	End If
    
    ' Find Branch B-C
    BranchBCHnd = branchSearch(BusBHnd&, BusCHnd&)
    If BranchBCHnd = 0 Then
       Print "Error: Cant't find branch between bus B and C"
	   GoTo ErrorFindBranch
	End If
    
    ' Find Branch C-D
    BranchCDHnd = branchSearch(BusCHnd&, BusDHnd&) 
	If BranchCDHnd = 0 Then
       Print "Error: Cant't find branch between bus C and D"
	   GoTo ErrorFindBranch
	End If
	
	' Fault connection flags
	vnFltConn(1) = 0	' Do 3PH
	vnFltConn(2) = 0	' Don't do 2LG
	vnFltConn(3) = 0	' Don't do 1LG
	vnFltConn(4) = 0	' Don't do LL

	' User Defined Intermediate Percentage
	LinePercent = 0.0			' Select 10 %
	
	' Fault options flags
	vdFltOpt(1) = 0				' Close-in
	vdFltOpt(2) = 0				' Close-in w/ outage
	vdFltOpt(3) = 0				' Close-in with end opened
	vdFltOpt(4) = 0				' Close-in with end opened w/ outage
	vdFltOpt(5) = 0				' Remote bus
	vdFltOpt(6) = 0				' Remote bus w/ outage
	vdFltOpt(7) = 0				' Line end
	vdFltOpt(8) = 0				' Line end w/ outage
	vdFltOpt(9) = 0         	' Intermediate %
	vdFltOpt(10) = 0			' Intermediate % w/ outage
	vdFltOpt(11) = 0			' Intermediate % with end opened
	vdFltOpt(12) = 0			' Intermediate % with end opened w/ outage
	vdFltOpt(13) = 0			' Auto seq. Intermediate % from
	vdFltOpt(14) = 0			' Auto seq. Intermediate % to
		
	' Branch outage option flags
	vnOutageOpt(1) = 0	' One at a time
	vnOutageOpt(2) = 0	' Two at a time
	vnOutageOpt(3) = 0	' All at once
	
    vnOutageLst(1) = -1 ' Terminate the list for good measure  
	
		
	' Other fault option variables
	dRflt = 0			' Fault resistance, in Ohms
	dXflt = 0			' Fault reactance, in Ohms
	nClearPrev = 1		' Clear previous result flag
	
    'Do fault study
    'Test 1: AG fault at Bus A
    vnFltConn(3) = 1
    vdFltOpt(1)  = 1    
    If DoFault(BusAHnd, vnFltConn, vdFltOpt, vnOutageOpt, vnOutageLst, dRflt, dXflt, nClearPrev) = 0 Then GoTo ErrorApplyFault 
    vnFltConn(3) = 0
    vdFltOpt(1)  = 0 
    
    nClearPrev = 0
    
    'Test 2: BC fault 5% from Bus B - Bus A
    vnFltConn(4) = 1
    vdFltOpt(9)  = 95.0
    If DoFault(BranchABHnd&, vnFltConn, vdFltOpt, vnOutageOpt, vnOutageLst, dRflt, dXflt, nClearPrev) = 0 Then GoTo ErrorApplyFault
    vnFltConn(4) = 0
    vdFltOpt(9)  = 0.0
    
    'Test 3: ABC fault 5% from Bus B - Bus A
    vnFltConn(1) = 1
    vdFltOpt(9)  = 95.0
    If DoFault(BranchABHnd&, vnFltConn, vdFltOpt, vnOutageOpt, vnOutageLst, dRflt, dXflt, nClearPrev) = 0 Then GoTo ErrorApplyFault
    vnFltConn(1) = 0
    vdFltOpt(9)  = 0.0
    
    'Test 4: BG fault 5% from Bus B - Bus C
    vnFltConn(3) = 2
    vdFltOpt(9)  = 5.0
    If DoFault(BranchBCHnd&, vnFltConn, vdFltOpt, vnOutageOpt, vnOutageLst, dRflt, dXflt, nClearPrev) = 0 Then GoTo ErrorApplyFault
    vnFltConn(3) = 0
    vdFltOpt(9)  = 0.0   
    
    'Test 5: ACG fault 5% from Bus B - Bus C
    vnFltConn(2) = 2
    vdFltOpt(9)  = 5.0
    If DoFault(BranchBCHnd&, vnFltConn, vdFltOpt, vnOutageOpt, vnOutageLst, dRflt, dXflt, nClearPrev) = 0 Then GoTo ErrorApplyFault
    vnFltConn(2) = 0
    vdFltOpt(9)  = 0.0
    
    'Test 6: ABC fault 50% from Bus C - Bus B
    vnFltConn(1) = 1
    vdFltOpt(9)  = 50.0
    If DoFault(BranchBCHnd&, vnFltConn, vdFltOpt, vnOutageOpt, vnOutageLst, dRflt, dXflt, nClearPrev) = 0 Then GoTo ErrorApplyFault
    vnFltConn(1) = 0
    vdFltOpt(9)  = 0.0     
    
    'Test 7: AB fault 5% from Bus C - Bus B
    vnFltConn(4) = 3
    vdFltOpt(9)  = 95.0
    If DoFault(BranchBCHnd&, vnFltConn, vdFltOpt, vnOutageOpt, vnOutageLst, dRflt, dXflt, nClearPrev) = 0 Then GoTo ErrorApplyFault
    vnFltConn(4) = 0
    vdFltOpt(9)  = 0.0    
    
    'Test 8: BCG fault 5% from Bus C - Bus B
    vnFltConn(2) = 1
    vdFltOpt(9)  = 95.0
    If DoFault(BranchBCHnd&, vnFltConn, vdFltOpt, vnOutageOpt, vnOutageLst, dRflt, dXflt, nClearPrev) = 0 Then GoTo ErrorApplyFault
    vnFltConn(2) = 0
    vdFltOpt(9)  = 0.0    
    
    'Test 9: CG fault 5% from Bus C - Bus D
    vnFltConn(3) = 3
    vdFltOpt(9)  = 5.0
    If DoFault(BranchCDHnd&, vnFltConn, vdFltOpt, vnOutageOpt, vnOutageLst, dRflt, dXflt, nClearPrev) = 0 Then GoTo ErrorApplyFault
    vnFltConn(3) = 0
    vdFltOpt(9)  = 0.0    
    
    'Test 10: ABC fault 5% from Bus C - Bus D
    vnFltConn(1) = 1
    vdFltOpt(9)  = 5.0
    If DoFault(BranchCDHnd&, vnFltConn, vdFltOpt, vnOutageOpt, vnOutageLst, dRflt, dXflt, nClearPrev) = 0 Then GoTo ErrorApplyFault
    vnFltConn(1) = 0
    vdFltOpt(9)  = 0.0  
    
    'Test 11: AC fault at Bus D
    vnFltConn(4) = 2
    vdFltOpt(1)  = 1
    If DoFault(BusDHnd&, vnFltConn, vdFltOpt, vnOutageOpt, vnOutageLst, dRflt, dXflt, nClearPrev) = 0 Then GoTo ErrorApplyFault
    vnFltConn(4) = 0
    vdFltOpt(1)  = 0           
    strMsg = strMsg + "11 fault simulations are done."
	Print strMsg
	Exit Sub    
	
HasError :
    Stop		
	
ErrorFindBranch :
    Stop
		
ErrorApplyFault :
	Print "Error - Can't Apply Fault"
	Stop

End Sub  ' End of Sub Main()

'************************************************************************************************************************************************
Function  branchSearch( nHndBus1&, nHndBus2& ) As long
  branchSearch = 0
  BranchHnd = 0
  While GetBusEquipment( nHndBus1, TC_BRANCH, BranchHnd ) > 0
      Call GetData( BranchHnd, BR_nBus2Hnd, nItemHnd& )
      If nItemHnd = nHndBus2 Then
        branchSearch = BranchHnd
        exit Function
      End If
  Wend
  branchSearchReturn:
End Function
'************************************************************************************************************************************************
	

