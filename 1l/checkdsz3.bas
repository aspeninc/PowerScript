' ASPEN PowerScrip sample program
'
' CHECKDSZ3.BAS
'
' Find distance relay zone 1, 2 or 3 reach by checking relay operating time
' in intermediate faults on transmission line.
'
'
' Global variables
'
Const DS_DELAY = 400
Const DS_ZNO1 = 10000
Const DS_ZNO2 = 20000
Const DS_ZNO3 = 30000
Const DS_ZNO4 = 40000
Const DS_ZNO5 = 50000
Const DS_ZNO6 = 52000
Const DS_ZNO7 = 54000
Const DS_ZNO8 = 56000

Dim FltConnection(4) As Long
Dim StepSize As Double
Dim Z2ThresholdMax As double
Dim Z2ThresholdMin As double
Dim Z3ThresholdMax As double
Dim Z3ThresholdMin As double
Dim DSType As long
dim sFile As String, sTag As String
dim nChecked As long
dim nConnP As long
dim nConnG As long
dim nZone as long
dim nUseFltType  As long
dim FltRmin As double, FltRmax As double, FltXmin As double, FltXmax As double
dim NOFltZ As long
dim vnBusHnd1(2) As long, vnBusHnd2(200) As long
 
Sub main()

'******************************************************************************************************
'TODO: adjust checking parameters in this section if needed
 sFile$      = "c:\0tmp\zone2check.csv"		' output file name. Must have csv extension to open in Excel
 DSType      = TC_RLYDSP	' Check relay type: TC_RLYDSP=Phase;TC_RLYDSG=Ground
 StepSize    = 1.0			' intermediate fault percent step
 FltRmin     = 0.0          ' Fault resistance min
 FltRmax     = 5.0          ' Fault resistance max
 FltXmin     = 0.0          ' Fault reactance min
 FltXmax     = 0.0          ' Fault reactance max
 NOFltZ      = 0            ' Number of fault Z to consider
 nConnP      = 1			' Fault type to check phase relay (1=3LG;2=2LG;3=LG;4=LL)
 nConnG      = 3			' Fault type to check phase relay (1=3LG;2=2LG;3=LG;4=LL)
 nUseFltType = 11			' Intermediate fault w/ end open (9=intermediate;11=Inter. w/ end open)
 sTag        = ""           ' Tag string to check
 nZone       = 3            ' Zone to check
 nOutage     = 1            ' Check with line outage
 Z2ThresholdMax = 83.0      ' Acceptable zone reach in percent of line impedance
 Z2ThresholdMin = 78.0

'******************************************************************************************************

 nChecked = 0

 If sFile$ <> "" Then 
  Open sFile For output As 1 
 Else 
  Print "Problem creating output file " & sFile$
  exit Sub
 End If
 
 Call printHeader()		' Print report header
 
 nChecked = 0
 
 If GetEquipment( TC_PICKED, PickedHnd ) > 0 Then
  If EquipmentType( PickedHnd ) = TC_RLYGROUP Then
   nChecked = CheckOneGroup(PickedHnd)
  Else
   If EquipmentType( PickedHnd ) = TC_BUS Then nChecked = CheckOneBus(PickedHnd)
  End If
 End If
 If nChecked = 0 Then
  sWindowText$ = "Select bus to check (200 or fewer)"
  vnBusHnd1(1) = 0
  nPicked& = BusPicker( sWindowText$, vnBusHnd1, vnBusHnd2 )
  If nPicked = 0 Then exit Sub
  For ii& = 1 to nPicked
   nChecked = nChecked + CheckOneBus(vnBusHnd2(ii))
  Next
 End If
 
 If nChecked > 0 Then
  sMsg$ = "Checked " + Str(nChecked) + " relays. Report is in " + sFile _
           + Chr(13) + "Do you want to open this file in Excel?"
  If 6 = MsgBox( sMsg, 4, "Check DS Zone" ) Then
   Set xlApp = CreateObject("excel.application")
   xlApp.Workbooks.Open Filename := sFile
   xlApp.Visible = True
  End If
 Else
  Print "Found no relay matching given criteria"
 End If
 exit Sub
HasError:
  Print "Error: ", ErrorString( )
End Sub

Function CheckOneBus( ByVal Bus1Hnd& ) As long
 CheckOneBus = 0
 ' Retrieve all branches
 BranchHnd& = 0
 While GetBusEquipment( Bus1Hnd, TC_BRANCH, BranchHnd ) > 0
  ' Branch must be a line
  Call GetData( BranchHnd, BR_nType, TypeCode& )
  If TypeCode <> TC_LINE Then GoTo Continue
  ' Line must have a relay group
  If GetData( BranchHnd, BR_nRlyGrp1Hnd, GroupHnd& ) <= 0 Then GoTo continue
  CheckOneBus = CheckOneBus + CheckOneGroup( GroupHnd )
  Continue:
 Wend
End Function

Function ReachPercent() As double
  ' Look for reach as intermediate fault percentage
  aString$ = FaultDescription()
  Pos2& = InStr(1,aString,"%)")
  If Pos2 = 0 Then Pos2& = InStr(1,aString,"%")
  Pos1& = Pos2 - 1
  While True
   StrTmp$ = Mid(aString,Pos1,1) 
   If (StrTmp = " ") Or (StrTmp$="(") Then exit Do Else Pos1 = Pos1 - 1
  Wend
  StrTmp = Mid(aString,Pos1+1,Pos2-Pos1-1)
  ReachPercent = Val(StrTmp)
End Function

Function DSZoneReach( ByVal RlyHnd&, ByVal ZoneNo&, ByRef ReachS#, ByRef ReachE# ) As double
 Dim ShowFlagRly(4) As Long

 ' Make sure relay operating time output is enable
 For ii = 1 To 4 
   ShowFlagRly(ii) = 1
 Next 
 
 DSZone2Reach = 1.0
 ReachS = 999
 ReachE = -999

 ZoneDelay = 0 
 If nZone > 1 Then
  ' Retrieve relay zone 2 delay setting
  If nZone = 2 Then 
   sDelay$ = Format(DS_DELAY + DS_ZNO2,"0") ' Zone 2 delay code
  Else
   sDelay$ = Format(DS_DELAY + DS_ZNO3,"0") ' Zone 3 delay code
  End If
  If EquipmentType(RlyHnd) = TC_RLYDSG Then
   Call GetData( RlyHnd, DG_sParam, sDelay$ )
  Else
   Call GetData( RlyHnd, DP_sParam, sDelay$ )
  End If
  ZoneDelay = Val(sDelay)
 End if
  
 ' Go through all faults. Find the range of zone tripping   
 ShowFaultFlag& = 1 ' Start with first fault
 OpTime#  = 0.0
 While PickFault( ShowFaultFlag ) > 0
  ShowFaultFlag = SF_NEXT
  Reach = ReachPercent()
  
  ' Get relay times
  Call GetRelayTime( RlyHnd, 1.0, OpTime )
  If OpTime >= 9998 Then 'NOP
   If Reach > ReachS And Reach < ReachE Then DSZone2Reach = -1.0 ' No NOP inside zone
   GoTo SkipIt
  End If
  
  If Abs(OpTime - ZoneDelay) < 0.0001 Then
   If ReachE < Reach Then ReachE = Reach
   If ReachS > Reach Then ReachS = Reach
  End If
  SkipIt:
 Wend
End Function		'DSZone2Reach

Function CheckOneGroup( ByVal GroupHnd ) As long
 ' Check zone reach of DS relay in the group
 
 dim FltBranchArray(20) As long
 dim LineArray(20) As String
 dim RlyHndArray(20) As long
 
 CheckOneGroup = 0
 
 ' Inventory relays in this group
 nRlyCount = 0
 RelayHnd& = 0
 While GetRelay( GroupHnd, RelayHnd ) > 0
  TypeCode& = EquipmentType( RelayHnd )
  If TypeCode& = DSType Then 
   ' Check active flag 
   If TypeCode = TC_RLYDSP Then
    Call GetData(RelayHnd, DP_nInService, nFlag& )
   Else
    Call GetData(RelayHnd, DG_nInService, nFlag& )
   End If
   If nFlag = 1 Then
    RlyHndArray(nRlyCount) = RelayHnd
    nRlyCount = nRlyCount + 1
   End If
  End If
 Wend
 
 If nRlyCount = 0 Then exit Function
 
 Call GetData( GroupHnd, RG_nBranchHnd, BranchHnd& )
 Call GetData( BranchHnd, BR_nBus1Hnd, RlyBus1Hnd& )
 Call GetData( BranchHnd, BR_nBus2Hnd, RlyBus2Hnd& )
 Call GetData( BranchHnd, BR_nHandle, LineHnd& )
 Call GetData( LineHnd, LN_sID, RlyLineID$ )
 Print #1, "Relay group at:," & FullBusName( RlyBus1Hnd ) & "-" & FullBusName( RlyBus2Hnd ) & " " & RlyLineID$

 FltBranchArray(0) = BranchHnd
 nLineCount = 1

 If nZone > 1 Then 
  ' Inventory lines at the remote bus   
  nRemoteBusHnd& = FindLineRemoteBus( BranchHnd, dLineZ# )
  BranchHnd& = 0
  While GetBusEquipment( nRemoteBusHnd, TC_BRANCH, BranchHnd& ) > 0
   Call GetData(BranchHnd, BR_nType, BrType& )
   If BrType = TC_LINE Then
    ' Must exclude the relay line
    If RlyBus1Hnd <> FindLineRemoteBus( BranchHnd, dLineZ# ) Then
     Call GetData( BranchHnd, BR_nInservice, nFlag& )
     If nFlag = 1 Then
      FltBranchArray(nLineCount) = BranchHnd
      nLineCount = nLineCount + 1
     End If
    End If
   End If
  Wend
 End If
 For OutageNo = 0 to nOutageCount 
  if OutageArray(OutageNo) > 0 then
  For LineNo = 0 to nLineCount -1
   BranchHnd = FltBranchArray(LineNo)
   If 0 = SimulateLineFaults( BranchHnd ) Then exit For
   Call GetData( BranchHnd, BR_nBus1Hnd, Bus1Hnd& )
   Call GetData( BranchHnd, BR_nBus2Hnd, Bus2Hnd& )
   Call GetData( BranchHnd, BR_nHandle, LineHnd& )
   Call GetData( LineHnd, LN_sID, FltLineID$ )
   FltLineName$ = FullBusName( Bus1Hnd ) & "-" & FullBusName( Bus2Hnd ) & " " & FltLineID$
   ' Check each relay
   For RlyNo = 0 to nRlyCount - 1
    RelayHnd = RlyHndArray(RlyNo)
    If LineNo = 0 Then 
     TypeCode& = EquipmentType( RelayHnd )
     If (TypeCode = TC_RLYDSP) Then
      Call GetData( RelayHnd, DP_sID, RelayID$ )
     Else
      Call GetData( RelayHnd, DG_sID, RelayID$ )
     End If
     LineArray(RlyNo) = RelayID$ & " zone" & Str(nZone)
    End If
    Zone1Reach# = DSZoneReach( RelayHnd, nZone, dReachS#, dReachE# )
    If dReachS < 999 Then
     aString$ = Chr(13) & "," & Format(dReachS,"0") & " - " & Format(dReachE,"0") & "%," & FltLineName
    Else
     aString$ = Chr(13) & "," & "NOP" & FltLineName
    End If
    LineArray(RlyNo) = LineArray(RlyNo) & aString
   Next  'Each relay
  Next ' Each line
  For RlyNo = 0 to nRlyCount-1
   Print #1, LineArray(RlyNo)
  Next
 Next'Each outage
 CheckOneGroup = nRlyCount
End Function 	'CheckOneGroup

Function SimulateLineFaults( ByVal BranchHnd& ) As long
 Dim FltOption(14) As Double
 Dim OutageType(3) As Long
 Dim OutageList(15) As Long
 Dim FltConnStr(4) As String
 
 For ii = 1 To 14
  FltOption(ii) = 0.0
 Next
 For ii = 1 To 4
  FltConnection(1) = 0
 Next
 For ii = 1 To 3
   OutageType(ii) = 0
 Next
 
 If DSType = TC_RLYDSP Then  FltConnection(nConnP) = 1 Else FltConnection(nConnG) = 1	
 FltOption(nUseFltType) = StepSize	'Intermediate
 FltOption(13)          = 0			'Intermediate percent from
 FltOption(14)          = 100		'Intermediate percent to

 If NOFltZ > 0 Then
  dXstep = (FltXmax - FltXmin) / NOFltZ
  dRstep = (FltRmax - FltRmin) / NOFltZ
  dFltR     = FltRmin   '
  dFltX     = FltXmin
 Else
  dFltR     = 0
  dFltX     = 0
 End If
 
 'Simulate faults
 For ii = 0 to NOFltZ
  If ii = 0 Then ClearPrev = 1 Else ClearPrev = 0
  SimulateLineFaults = DoFault( BranchHnd, FltConnection, FltOption, OutageType, OutageList, _
                   dFltR, dFltX, ClearPrev )
  dFltR = dFltR + dRstep
  dFltX = dFltX + dXstep
 Next
                   
End Function	'SimulateLineFaults

Sub printHeader()
  Print #1, "Date:,", Date()
  Print #1, "Time:,", Time()
  Print #1, "Name of this file:,",sFile
  Print #1, "OLR file:,", GetOlrFileName()
  If DSType = TC_RLYDSP Then
   Print #1, "DS relay type:,Phase"
  Else
   Print #1, "DS relay type:,Ground"
  End If
  If NOFltZ = 0 Then
   Print #1, "Fault Z:,0"
  Else
   Print #1, "Fault Z  (ohm):,", FltRMin, "+j", FltXmin, " to ", FltRmax, "+j", FltXmax
  End If
'  Print #1, "Reach % Max:,", Z2ThresholdMax
'  Print #1, "Reach % Min:,", Z2ThresholdMin
  Print #1, ""
'  Print #1, "Bus1,Bus2,CktID,RelayID,Zone%,Flag"
End Sub		'printHeader

Function FindLineRemoteBus( ByVal Branch1Hnd&, ByRef dLineZ# ) As long
  dLineZR = 0.0
  dLineZX = 0.0
  
  ' Skip all taps on the line
  Do 
    Call GetData( Branch1Hnd, BR_nHandle, LineHnd )
    Call GetData( LineHnd, LN_sName, LineName )
    Call GetData( LineHnd, LN_dR, LineR )
    Call GetData( LineHnd, LN_dX, LineX )
    dLineZR = dLineZR + LineR
    dLineZX = dLineZX + LineX
    Call GetData( Branch1Hnd, BR_nBus1Hnd, BusHnd )
    Call GetData( Branch1Hnd, BR_nBus2Hnd, Bus1Hnd )
    Call GetData( Bus1Hnd, BUS_nTapBus, TapCode )
    If TapCode = 0 Then Exit Do			' real bus
    ' Only for tap bus
    Branch1Hnd& = 0
    ttt = GetBusEquipment( Bus1Hnd, TC_BRANCH, Branch1Hnd& )
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
    Wend
    ExitWhile:
    BusHnd  = Bus1Hnd
    Bus1Hnd = Bus2Hnd	
  Loop While TapCode = 1
   
  ExitLoop:
  
  dLineZ# = Sqr(dLineZR*dLineZR + dLineZX*dLineZX)
  
  FindLineRemoteBus = Bus1Hnd
End Function

