' ASPEN PowerScrip sample program
'
' CHECKDSZ1.BAS
'
' Find distance relay zone 1 reach by checking relay operating time
' in intermediate faults on transmission line.
'
' Version 1.0
' Category: OneLiner
'
' Global variables
'
Dim FltConnection(4) As Long
Dim StepSize As Double
Dim Z1ThresholdMax As double
Dim Z1ThresholdMin As double
Dim Z2ThresholdMax As double
Dim Z2ThresholdMin As double
Dim DSType As long
dim sFile As String, sTag As String
dim nChecked As long
dim nConnP As long
dim nConnG As long
dim nUseFltType  As long
dim FltRmin As double, FltRmax As double, FltXmin As double, FltXmax As double
dim NOFltZ As long
 
Sub main()

'******************************************************************************************************
'TODO: adjust checking parameters in this section if needed
 sFile$      = "c:\0tmp\zone1check.csv"		' output file name. Must have csv extension to open in Excel
 DSType      = TC_RLYDSP	' Check relay type: TC_RLYDSP=Phase;TC_RLYDSG=Ground
 StepSize    = 1.0			' intermediate fault percent step
 FltRmin     = 0.0          ' Fault resistance min
 FltRmax     = 5.0          ' Fault resistance max
 FltXmin     = 0.0          ' Fault reactance min
 FltXmax     = 0.0          ' Fault reactance max
 NOFltZ      = 4            ' Number of fault Z to consider
 nConnP      = 1			' Fault type to check phase relay (1=3LG;2=2LG;3=LG;4=LL)
 nConnG      = 3			' Fault type to check phase relay (1=3LG;2=2LG;3=LG;4=LL)
 nUseFltType = 11			' Intermediate fault w/ end open (9=intermediate;11=Inter. w/ end open)
 sTag        = ""           ' Tag string to check

 Z1ThresholdMax = 83.0      ' Acceptable zone reach in percent of line impedance
 Z1ThresholdMin = 78.0

'******************************************************************************************************

 nChecked = 0

 If sFile$ <> "" Then 
  Open sFile For output As 1 
 Else 
  Print "Problem creating output file " & sFile$
  exit Sub
 End If
 
 Call printHeader()		' Print report header
 
 If (GetEquipment( TC_PICKED, PickedHnd ) <> 0) And (EquipmentType( PickedHnd ) = TC_RLYGROUP) Then
  ' Check selected relay group only
  nChecked = CheckOneGroup(PickedHnd)
 Else 
  sMsg$ = "No relay group is selected. Do you want to check the entire system?" _
    + Chr(13) + "(NOTE: This operation may take up to several minutes to complete)"
  If 6 <> MsgBox( sMsg, 4, "Check DS Zone Reach" ) Then exit Sub
  ' Check the entire file
  BusHnd& = 0
  While NextBusByName( BusHnd ) > 0
   Call GetData( BusHnd, BUS_nTapBus, nFlag& )
   If nFlag <> 0 Then GoTo Continue1 ' Skip tap bus
   If sTag <> "" Then
    If InStr( 1, GetObjTags(BusHnd), sTag ) < 1 Then GoTo Continue1 
   End If
   Call CheckOneBus(BusHnd)
  Continue1:
  Wend
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

Sub CheckOneBus( ByVal Bus1Hnd& )
 Bus1Name$ = ""
 ' Retrieve all branches
 BranchHnd& = 0
 While GetBusEquipment( Bus1Hnd, TC_BRANCH, BranchHnd ) > 0
  Bus2Name$ = ""
  ' Branch must be a line
  Call GetData( BranchHnd, BR_nType, TypeCode& )
  If TypeCode <> TC_LINE Then GoTo Continue
  ' Line must have a relay group
  If GetData( BranchHnd, BR_nRlyGrp1Hnd, GroupHnd& ) <= 0 Then GoTo continue
  
  nChecked = nChecked + CheckOneGroup( GroupHnd )
  
  Continue:
 Wend
 exit Sub
HasError:
  Print "Error: ", ErrorString( )
End Sub

Function DSZone1Reach( ByVal RlyHnd&, ByRef Reach1S#, ByRef Reach1E#, ByRef Reach2S#, ByRef Reach2E# ) As double
 Dim ShowFlagRly(4) As Long

 ' Initialize 
 For ii = 1 To 4 
   ShowFlagRly(ii) = 1
 Next 
 
 DSZone1Reach = -1.0
 Reach1S = 999
 Reach1E = -999
 Reach2S = 999
 Reach2E = -999
 
 ' Go through all faults. Find the smallest and largest zone 1 reach   
 ShowFaultFlag& = 1
 OpTime#  = 0.0
 While PickFault( ShowFaultFlag ) > 0
  ShowFaultFlag = SF_NEXT
  ' Get relay times
  Call GetRelayTime( RlyHnd, 1.0, OpTime )
  If OpTime = 9999 Then GoTo SkipIt
  
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
  Reach = Val(StrTmp)
  
  If OpTime = 0 Then
   If Reach1E < Reach Then Reach1E = Reach
   If Reach1S > Reach Then Reach1S = Reach
   DSZone1Reach = 100.0		'Zone 1 tripped
  Else
   If Reach2S > Reach Then Reach2S = Reach
   If Reach2E < Reach Then Reach2E = Reach
  End If
  SkipIt:
 Wend
End Function		'DSZone1Reach

Function CheckOneGroup( ByVal GroupHnd ) As long
 CheckOneGroup = 0
 FirstTime& = 1
 
 ' Group must have a phase DS relay
 Count& = 0
 RelayHnd& = 0
 While GetRelay( GroupHnd, RelayHnd ) > 0
 ' Check active flag to be on the safe side
  If (DSType = TC_RLYDSP) Then
   nCode& = GetData(RelayHnd, DP_nInService, nFlag& )
  Else
   Call GetData(RelayHnd, DG_nInService, nFlag& )
  End If
  If nFlag <> 1 Then GoTo NextRelay
 
  TypeCode& = EquipmentType( RelayHnd )
  If TypeCode <> DSType Then GoTo NextRelay
  If FirstTime = 1 Then
   Call GetData( GroupHnd, RG_nBranchHnd, BranchHnd& )
   Call GetData( BranchHnd, BR_nBus1Hnd, Bus1Hnd& )
   Call GetData( BranchHnd, BR_nBus2Hnd, Bus2Hnd& )
   Call GetData( BranchHnd, BR_nHandle, LineHnd& )
   Call GetData( LineHnd, LN_sID, LineID$ )
   Bus1Name$ = FullBusName( Bus1Hnd )
   Bus2Name$ = FullBusName( Bus2Hnd )
   LineName$ = Bus1Name + " - " + Bus2Name
   If 0 = SimulateLineFaults( BranchHnd ) Then exit Function
   FirstTime = 0
  End If
  If DSType = TC_RLYDSP Then
   Call GetData( RelayHnd, DP_sID, RelayID$ )
   aString$ = "DS Phase Relay "
  Else
   Call GetData( RelayHnd, DG_sID, RelayID$ )
   aString$ = "DS Ground Relay "
  End If
  Zone1Reach# = DSZone1Reach( RelayHnd, dReach1S#, dReach1E#, dReach2S#, dReach2E# )
  aString$ = Bus1Name & "," & Bus2Name & "," & LineID & "," & RelayID & "," _
              & Str(dReach1S) & "-" & Str(dReach1E) _
              & "," & Str(dReach2S) & "-" & Str(dReach2E)

  If Zone1Reach < 0 Then
   aString$ = aString$ + ",RESTRAINED,,"
  Else
   If (Z1ThresholdMin > dReach2S) Then _
    aString$ = aString$ + ",UNDER_REACH" _
   Else _
    aString$ = aString$ + ","
   If (Z1ThresholdMax < dReach1E) Then _
    aString$ = aString$ + ",OVER_REACH" _
   Else _
    aString$ = aString$ + ","
  End If
  Print #1, aString
  Count = Count + 1
  
  NextRelay:
 Wend  'Each relay
 CheckOneGroup = Count
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

 dXstep = (FltXmax - FltXmin) / NOFltZ
 dRstep = (FltRmax - FltRmin) / NOFltZ
 
 'Simulate faults
 dFltR     = FltRmin   '
 dFltX     = FltXmin
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
  Print #1, "Reach % Max:,", Z1ThresholdMax
  Print #1, "Reach % Min:,", Z1ThresholdMin
  Print #1, ""
  Print #1, "Bus1,Bus2,CktID,RelayID,Zone1%,Zone2%,Flag"
End Sub		'printHeader
