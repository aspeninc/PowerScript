' ASPEN PowerScrip sample program
'
' LineFault.BAS
'
' Run intermediate fault simulation with fixed fault location on all transmission lines
'
' Version 1.0
' Category: OneLiner
'

Sub Main()

	' Output starting text to TTY window 
	PrintTTY("Starting ..." & Chr(10))

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

	' Fault connection flags
	vnFltConn(1) = 1	' Do 3PH
	vnFltConn(2) = 0	' Don't do 2LG
	vnFltConn(3) = 0	' Don't do 1LG
	vnFltConn(4) = 0	' Don't do LL

	' User Defined Intermediate Percentage
	LinePercent = 10.0			' Select 10 %
	
	' Fault options flags
	vdFltOpt(1) = 0				' Close-in
	vdFltOpt(2) = 0				' Close-in w/ outage
	vdFltOpt(3) = 0				' Close-in with end opened
	vdFltOpt(4) = 0				' Close-in with end opened w/ outage
	vdFltOpt(5) = 0				' Remote bus
	vdFltOpt(6) = 0				' Remote bus w/ outage
	vdFltOpt(7) = 0				' Line end
	vdFltOpt(8) = 0				' Line end w/ outage
	vdFltOpt(9) = LinePercent	' Intermediate %
	vdFltOpt(10) = 0			' Intermediate % w/ outage
	vdFltOpt(11) = 0			' Intermediate % with end opened
	vdFltOpt(12) = 0			' Intermediate % with end opened w/ outage
	vdFltOpt(13) = 0			' Auto seq. Intermediate % from
	vdFltOpt(14) = 0			' Auto seq. Intermediate % to
		
	' Branch outage option flags
	vnOutageOpt(1) = 0	' One at a time
	vnOutageOpt(2) = 0	' Two at a time
	vnOutageOpt(3) = 0	' All at once
	
	' List of handles of branches to be outaged
	For q = 1 To 30
		vnOutageLst(q) = 0  
	Next q
		
	' Other fault option variables
	dRflt = 0			' Fault resistance, in Ohms
	dXflt = 0			' Fault reactance, in Ohms
	nClearPrev = 1		' Clear previous result flag
	
	i = 0 	' Temporary variable so the program does not run long 
	' Do fault calculation on every line (just see if it can perform a fault)
	While GetEquipment(TC_LINE, LineHandle&) > 0
		
		' Apply the fault with the above parameters
		'If DoFault(LineHandle&, vnFltConn, vdFltOpt, vnOutageOpt, vnOutageLst, dRflt, dXflt, nClearPrev) = 0 Then GoTo ErrorApplyFault

'The LineHandle needs to be replaced by BranchHnd
'************************************************************************************************************************************************
       'Find from bus handle of the line
       Call GetData(LineHandle, LN_nBus1Hnd, nHndBus1&)
       'Find branch handle, function branchSearch() is defined at the bottom of this file
       BranchHnd = branchSearch(nHndBus1&, LineHandle)
       If BranchHnd = 0 Then GoTo ErrorApplyFault
       'Do fault study
       If DoFault(BranchHnd&, vnFltConn, vdFltOpt, vnOutageOpt, vnOutageLst, dRflt, dXflt, nClearPrev) = 0 Then GoTo ErrorApplyFault 
'************************************************************************************************************************************************

		' Temporary code for running only a few lines
		i = i + 1		
		PrintTTY("i = " & i)
		If i > 11 Then GoTo Done
	EndLoop:
	Wend

	GoTo Done	
	
HasError:
	PrintTTY("ERROR - HasError")
	Stop
		
ErrorApplyFault:
	PrintTTY("ERROR - Apply Fault")
	Stop
	
Done:
	PrintTTY(" " & Chr(10) & "Done!")	
	Exit Sub
End Sub  ' End of Sub Main()

'************************************************************************************************************************************************
Function  branchSearch( nHndBus1&, LineHandle ) As long
  branchSearch = 0
  BranchHnd = 0
  While GetBusEquipment( nHndBus1, TC_BRANCH, BranchHnd ) > 0
      Call GetData( BranchHnd, BR_nHandle, nItemHnd& )
      If nItemHnd = LineHandle Then
        branchSearch = BranchHnd
        exit Function
      End If
  Wend
  branchSearchReturn:
End Function
'************************************************************************************************************************************************
	
