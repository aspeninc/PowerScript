' ASPEN PowerScrip sample program
'
' CHECKDSZONE.BAS
'
' Simulate series of intermediate faults on line
' to check zone 1 reach of distance relays
'
' PowerScript functions called:
'
' Global variables
Dim FltConnection(4) As Long
Dim StepSize As Double
Dim ThresholdMax As double
Dim ThresholdMin As double
Dim DSType As long
dim sFile As String
dim nChecked As long

Sub main()

 ' Initialize global variables
 StepSize  = 5.0
 FltConnection(1) = 0	'3LG
 FltConnection(2) = 0	'2LG
 FltConnection(3) = 0	'1LG
 FltConnection(4) = 0	'LL
 ThresholdMax = 75.0
 ThresholdMin = 70.0
' DSType = TC_RLYDSP
 DSType = TC_RLYDSG
 sFile$  = "c:\0tmp\dscheck.csv"
 nChecked = 0

 If (GetEquipment( TC_PICKED, PickedHnd ) <> 0) And _
    (EquipmentType( PickedHnd ) = TC_RLYGROUP) Then
  ' Check selected group only
  CheckOneGroup(PickedHnd)
  exit Sub
 End If
 
 If sFile<>"" Then Open sFile For output As 1 Else exit Sub
 
 Call printHeader()
 
 BusHnd& = 0
 While NextBusByName( BusHnd ) > 0
  If 0 = GetData( BusHnd, BUS_nTapBus, nFlag& ) Then GoTo HasError
  If nFlag <> 0 Then GoTo Continue1
  Call CheckOneBus(BusHnd)
 Continue1:
 Wend
 
 If nChecked > 0 Then
  sMsg$ = "Checked " + str(nChecked) + " relays. Report is in " + sFile _
           + chr(13) + "Do you want to open this file in Excel?"
  If 6 = MsgBox( sMsg, 4, "Check DS Zone" ) Then
   Set xlApp = CreateObject("excel.application")
   xlApp.Workbooks.Open Filename:=sFile
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
 ' Retrieve all branches and get their currents
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

Function DSZone1Reach( ByVal RlyHnd&, ByVal ZoneNo& ) As double
 Dim ShowFlagRly(4) As Long

 ' Initialize 
 For ii = 1 To 4 
   ShowFlagRly(ii) = 1
 Next 
   
 ShowFaultFlag& = 1
 OpTime#  = 0.0
'  While ShowFault( ShowFaultFlag, 1, 1, 0, ShowFlagRly ) > 0
 While PickFault( ShowFaultFlag ) > 0
  ShowFaultFlag = SF_NEXT
  ' Get relay times
  Call GetRelayTime( RlyHnd, 1.0, OpTime )
  If OpTime > 0 Then
   aString$ = FaultDescription()
   ' Look for intermediate fault percentage
   Pos2& = InStr(1,aString,"%)")
   If Pos2 = 0 Then Pos2& = InStr(1,aString,"%")
   Pos1& = Pos2 - 1
   While True
    StrTmp$ = Mid(aString,Pos1,1) 
    If (StrTmp = " ") Or (StrTmp$="(") Then exit Do Else Pos1 = Pos1 - 1
   Wend
   StrTmp = Mid(aString,Pos1+1,Pos2-Pos1-1)
   DSZone1Reach = Val(StrTmp)
   exit Function
  End If
 Wend
 DSZone1Reach = 1.0
End Function

Function CheckOneGroup( ByVal GroupHnd ) As long
 CheckOneGroup = 0
 FirstTime& = 1
 
 ' Group must have a phase DS relay
 Count& = 0
 RelayHnd& = 0
 While GetRelay( GroupHnd, RelayHnd ) > 0
 ' Check active flag to be on the safe side
'  If DSType = TC_RLYDSP Then
'   GetData(RelayHnd, DP_nInSevice, nFlag& )
'  Else
'   GetData(RelayHnd, DG_nInSevice, nFlag& )
'  End If
'  If nFlag <> 1 Then exit Function
 
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
  Zone1Reach# = DSZone1Reach( RelayHnd, 1 )
'  aString$ = aString + RelayID + Chr(13) + "Zone 1 reach=" + Str(Zone1Reach) + "%"  
  aString$ = Bus1Name + "," + Bus2Name + "," + LineID + "," + RelayID + "," + Str(Zone1Reach)
  If (ThresholdMin > Zone1Reach) Then
   aString$ = aString$ + ",UNDER_REACH"
  Else 
   If (ThresholdMax < Zone1Reach) Then
    aString$ = aString$ + ",OVER_REACH"
   Else
    aString$ = aString$ + ",OK"
   End If
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
 For ii = 1 To 3
   OutageType(ii) = 0
 Next
 dFltR     = 0.0   '
 dFltX     = 0.0
 ClearPrev = 1
 If DSType = TC_RLYDSP Then  FltConnection(1) = 1 _	
   Else FltConnection(3) = 1						
 'Simulate faults
 FltOption(11) = StepSize	'Intermediate w/ end open
 FltOption(13) = 0			'Intermediate percent from
 FltOption(14) = 100		'Intermediate percent to
     
 SimulateLineFaults = DoFault( BranchHnd, FltConnection, FltOption, OutageType, OutageList, _
                   dFltR, dFltX, ClearPrev )
                   
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
  Print #1, "Reach % Max:,", ThresholdMax
  Print #1, "Reach % Min:,", ThresholdMin
  Print #1, ""
  Print #1, "Bus1,Bus2,CktID,RelayID,Zone1Reach%,Flag"
End Sub		'printHeader
