' ASPEN PowerScrip sample program
'
' BusFltwContingency.BAS
'
' Simulate faults on the selected bus with N-1 contingency analysis (branch outage)
'

Sub main()
  Dim vnBusHnd1(1000) As Long
  Dim vnBusHnd2(1000) As Long     'The maximum number of selected fault buses is set to 1000  
  Dim vnFltConn(4) As Long
  Dim vdFltOpt(14) As Double
  Dim vnOutageOpt(4) As Long 
  Dim vnOutageList(2000) As Long  'The maximum number of line outages is set to 2000
  

  ' Fault bus selection
  sWindowText$ = "My Bus Picker"
  nPicked& = BusPicker( sWindowText$, vnBusHnd1, vnBusHnd2 )  

  ' Set up outage list in fault bus zone
  nBusHnd = vnBusHnd2(1)
  If GetData( nBusHnd, BUS_nZone, nZone& ) = 0 Then GoTo HasError
  nSelectedZone = nZone
  nOutage = 0 
  For ii = 1 To 2000
    vnoutageList(ii) = 0
  Next ii
  

  ' Line
  LineHandle = 0
  While GetEquipment( TC_Line, LineHandle& ) > 0
    If GetData( LineHandle, LN_nBus1Hnd, nBus1Hnd& ) = 0 Then GoTo HasError
    If GetData( LineHandle, LN_nBus2Hnd, nBus2Hnd& ) = 0 Then GoTo HasError
    If GetData( nBus1Hnd, BUS_nZone, nZone1& ) = 0 Then GoTo HasError
    If GetData( nBus2Hnd, BUS_nZone, nZone2& ) = 0 Then GoTo HasError
    If nZone1 = nSelectedZone And nZone2 = nSelectedZone Then
      nOutage = nOutage + 1
      vnOutageList(nOutage) = LineHandle
    End If
  Wend
  
  ' 2-winding transformer  
  XfmrHandle = 0  
  While GetEquipment( TC_XFMR, XfmrHandle& ) > 0
    If GetData( XfmrHandle, XR_nBus1Hnd, nBus1Hnd& ) = 0 Then GoTo HasError
    If GetData( XfmrHandle, XR_nBus2Hnd, nBus2Hnd& ) = 0 Then GoTo HasError
    If GetData( nBus1Hnd, BUS_nZone, nZone1& ) = 0 Then GoTo HasError
    If GetData( nBus2Hnd, BUS_nZone, nZone2& ) = 0 Then GoTo HasError
    If nZone1 = nSelectedZone And nZone2 = nSelectedZone Then
      nOutage = nOutage + 1
      vnOutageList(nOutage) = XfmrHandle
    End If
  Wend   
  
  ' Phase shifter
  PsHandle = 0  
  While GetEquipment( TC_PS, PsHandle& ) > 0
    If GetData( PsHandle, PS_nBus1Hnd, nBus1Hnd& ) = 0 Then GoTo HasError
    If GetData( PsHandle, PS_nBus2Hnd, nBus2Hnd& ) = 0 Then GoTo HasError
    If GetData( nBus1Hnd, BUS_nZone, nZone1& ) = 0 Then GoTo HasError
    If GetData( nBus2Hnd, BUS_nZone, nZone2& ) = 0 Then GoTo HasError
    If nZone1 = nSelectedZone And nZone2 = nSelectedZone Then
      nOutage = nOutage + 1
      vnOutageList(nOutage) = PsHandle
    End If
  Wend  
  
  ' Switch
  SwHandle = 0  
  While GetEquipment( TC_SWITCH, SwHandle& ) > 0
    If GetData( SwHandle, SW_nBus1Hnd, nBus1Hnd& ) = 0 Then GoTo HasError
    If GetData( SwHandle, SW_nBus2Hnd, nBus2Hnd& ) = 0 Then GoTo HasError
    If GetData( nBus1Hnd, BUS_nZone, nZone1& ) = 0 Then GoTo HasError
    If GetData( nBus2Hnd, BUS_nZone, nZone2& ) = 0 Then GoTo HasError
    If nZone1 = nSelectedZone And nZone2 = nSelectedZone Then
      nOutage = nOutage + 1
      vnOutageList(nOutage) = SwHandle
    End If
  Wend  

  ' 3-winding transformer
  Xfmr3Handle = 0  
  While GetEquipment( TC_XFMR3, Xfmr3Handle& ) > 0
    If GetData( Xfmr3Handle, X3_nBus1Hnd, nBus1Hnd& ) = 0 Then GoTo HasError
    If GetData( Xfmr3Handle, X3_nBus2Hnd, nBus2Hnd& ) = 0 Then GoTo HasError
    If GetData( Xfmr3Handle, X3_nBus3Hnd, nBus3Hnd& ) = 0 Then GoTo HasError
    If GetData( nBus1Hnd, BUS_nZone, nZone1& ) = 0 Then GoTo HasError
    If GetData( nBus2Hnd, BUS_nZone, nZone2& ) = 0 Then GoTo HasError
    If GetData( nBus3Hnd, BUS_nZone, nZone3& ) = 0 Then GoTo HasError
    If nZone1 = nSelectedZone And nZone2 = nSelectedZone And nZone3 = nSelectedZone Then
      nOutage = nOutage + 1
      vnOutageList(nOutage) = Xfmr3Handle
    End If
  Wend   

  ' Fault connections
  vnFltConn(1) = 1	' Do 3PH
  vnFltConn(2) = 0	' Don't do 2LG
  vnFltConn(3) = 0	' Don't do 1LG
  vnFltConn(4) = 0	' Don't do LL

  ' Fault options
  vdFltOpt(1) = 0	' Close-in
  vdFltOpt(2) = 1	' Close-in w/ outage
  vdFltOpt(3) = 0	' Close-in with end opened
  vdFltOpt(4) = 0	' Close-in with end opened w/ outage
  vdFltOpt(5) = 0	' Remote bus
  vdFltOpt(6) = 0	' Remote bus w/ outage
  vdFltOpt(7) = 0	' Line end
  vdFltOpt(8) = 0	' Line end w/ outage
  vdFltOpt(9) = 0	' Intermediate %
  vdFltOpt(10) = 0	' Intermediate % w/ outage
  vdFltOpt(11) = 0	' Intermediate % with end opened
  vdFltOpt(12) = 0	' Intermediate % with end opened w/ outage
  vdFltOpt(13) = 0	' Auto seq. Intermediate % from
  vdFltOpt(14) = 0	' Auto seq. Intermediate % to

  ' Branch outage option flags
  vnOutageOpt(1) = 1	' One at a time
  vnOutageOpt(2) = 0	' Two at a time
  vnOutageOpt(3) = 0	' All at once

  ' Other fault option variables
  dRflt = 0		' Fault resistance, in Ohms
  dXflt = 0		' Fault reactance, in Ohms
  nClearPrev = 1	' Clear previous result flag

  For ii = 1 To nPicked
    BusHnd = vnBusHnd2(ii)
    If DoFault( BusHnd, vnFltConn, vdFltOpt, vnOutageOpt, vnOutageList, dRflt, dXflt, nClearPrev ) = 0 Then GoTo HasError
    If ii = 1 Then nClearPrev = 0
  Next ii

  Exit Sub
  ' Error handling
  HasError:
  Print ErrorString()
End Sub