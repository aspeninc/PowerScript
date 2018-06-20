' ASPEN PowerScript Sample Program
'
' FltBusVCheck.BAS
'
' Run 1LG, 3LG, L-L & 2LG at picked bus
' Generate report with a list of buses with post-fault positive-sequence voltage lower
' than the threshold within specified tiers
'
' Version 1.0
' Category: OneLiner
'


Const MXNEIGHBOR = 50

Sub main
  Dim vnFltConn(4) As Long
  Dim vdFltOpt(14) As Double 
  Dim vnOutageOpt(3) As Long
  Dim OutageList(20) As Long
  Dim MagArray(3) As Double
  Dim AngArray(3) As Double
  
  dVthd = 0.75  ' Voltage threshold
  nTier = 1     ' Tier number

  If 0 = GetEquipment( TC_PICKED, PickedHnd ) Or _
    TC_BUS <> EquipmentType( PickedHnd ) Then
    Print "Please select a bus"
    Stop
  End If
  
  ' Fault connection flags
  vnFltConn(1) = 1	' Do 3PH
  vnFltConn(2) = 0	' Don't do 2LG
  vnFltConn(3) = 0	' Don't do 1LG
  vnFltConn(4) = 0	' Don't do LL 
  
  ' Fault options flags
  vdFltOpt(1) = 1				' Close-in
  vdFltOpt(2) = 0				' Close-in w/ outage
  vdFltOpt(3) = 0				' Close-in with end opened
  vdFltOpt(4) = 0				' Close-in with end opened w/ outage
  vdFltOpt(5) = 0				' Remote bus
  vdFltOpt(6) = 0				' Remote bus w/ outage
  vdFltOpt(7) = 0				' Line end
  vdFltOpt(8) = 0				' Line end w/ outage
  vdFltOpt(9) = 0               ' Intermediate %
  vdFltOpt(10) = 0			    ' Intermediate % w/ outage
  vdFltOpt(11) = 0			    ' Intermediate % with end opened
  vdFltOpt(12) = 0			    ' Intermediate % with end opened w/ outage
  vdFltOpt(13) = 0			    ' Auto seq. Intermediate % from
  vdFltOpt(14) = 0			    ' Auto seq. Intermediate % to
		
  ' Branch outage option flags
  vnOutageOpt(1) = 0	' One at a time
  vnOutageOpt(2) = 0	' Two at a time
  vnOutageOpt(3) = 0	' All at once
			
  ' Other fault option variables
  dFltR = 0			' Fault resistance, in Ohms
  dFltX = 0			' Fault reactance, in Ohms
		
  If 0 = DoFault( PickedHnd, vnFltConn, vdFltOpt, vnOutageOpt, OutageList, dFltR, dFltX, 1 ) Then GoTo HasError
  
  
  Print "Please select or create an excel file for output report"
  OutputFile$ = FileSaveDialog( "", "Excel File (*.csv)|*.csv||", ".csv", 2+16 )
  If Len(OutputFile) = 0 Then exit Sub
  Open OutputFile For Output As 1
  If PickFault( 1 ) = 0 Then GoTo HasError

  Print #1, "Bus Post-Fault Positive-Sequence Voltage Check Report"
  Print #1, "Date: ", Date()
  Print #1, "OneLiner file name: ", GetOLRFileName()
  Print #1, "Study date: N/A"
  Print #1, "Fault Description: " & FaultDescription()
  Print #1, ""     
  Print #1, "Bus Name, V1(pu)"  
  
  dim BusTodo(MXNEIGHBOR) As long
  nCount = FindBusInVicinity( PickedHnd, nTier, BusTodo )
  nCount1 = 0
  
  For ii = 2 to nCount 
    nBusHnd = BusTodo(ii)
    Call GetData( nBusHnd, Bus_dKVnorminal, dVBase )
    dVBase = dVBase/Sqr(3.0)
    If GetSCVoltage( nBusHnd, MagArray, AngArray, 2 ) = 0 Then GoTo HasError
    dMag = MagArray(2)/dVBase
    If dMag < dVthd Then
      nCount1 = nCount1 + 1
      sOut$ = FullBusName(nBusHnd) & "," & Format(dMag,"#0.000")
      Print #1, sOut$  
    End If
  Next 
  Close 1
  sMsg$ = "Checked " & Str(nCount-1) & " buses, " & Str(nCount1) & " buses are with post-fault positive-sequence voltage below " + Str(dVthd) & "pu"
  Print sMsg$
  exit Sub
HasError:
  Close
  Print ErrorString()
End Sub


Function FindBusInVicinity( ByVal nBusHnd&, ByVal TierWanted&, ByRef BusTodo() As long ) As long
  dim BusTier(MXNEIGHBOR) As long
  dim countBus As long
  dim nBusTodo As long
  
  BusToDo(1) = nBusHnd
  BusTier(1) = 0
  countBus   = 1
  nBusTodo   = 1
  While nBusTodo > 0
    nIdx = countBus - nBusTodo + 1
    Bus1Hnd  = BusToDo(nIdx)
    nTier    = BusTier(nIdx)
    nBusTodo = nBusTodo - 1
    BranchHnd   = 0
    While GetBusEquipment( Bus1Hnd, TC_BRANCH, BranchHnd ) > 0
      Call GetData( BranchHnd, BR_nBus2Hnd, nBus2Hnd )
      nFound = 0
      For ii = 1 to countBus
        If BusToDo(ii) = nBus2Hnd Then 
          nFound = 1
          GoTo BreakFor
        End If
      Next
      BreakFor:
      If nFound = 0 And nTier < TierWanted And countBus < MXNEIGHBOR - 1Then
        countBus = countBus + 1
        BusToDo(countBus) = nBus2Hnd
        BusTier(countbus) = nTier + 1
        nBusTodo = nBusTodo + 1
      End If
    Wend
  Wend
  FindBusInVicinity = countBus
End Function 