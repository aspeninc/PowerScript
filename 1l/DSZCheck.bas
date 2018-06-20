' ASPEN PowerScrip sample program
'
' DSZCheck.BAS
'
' Find distance relay zone reach by simulating intermediate faults on line
' in select bus 3ph&1ph fault.
'
' Version: 1.0
' Category: OneLiner
'

Sub main()
 Dim MagArray(12) As Double
 Dim AngArray(12) As Double
 Dim DummyArray(6) As Long   '

 ' Get picked object number
 If GetEquipment( TC_PICKED, ObjHnd ) = 0 Then 
   Print "Please select a relay group"
   Exit Sub
 End If
 If EquipmentType( ObjHnd ) <> TC_RLYGROUP Then 
   Print "Please select a relay group"
   Exit Sub
 End If

 ' Loop through all relays in the group and find desired relay
 nDesiredType = TC_RLYDSP
 RelayHnd&   = 0
 While GetRelay( ObjHnd, RelayHnd ) > 0
  RelayCount = RelayCount + 1
  TypeCode = EquipmentType( RelayHnd )
  
  If TypeCode = TC_RLYOCG Then 
   ParamID = OG_sID
   ParamInService = OG_nInService
  End If
  If TypeCode = TC_RLYOCP Then 
   ParamID = OP_sID
   ParamInService = OP_nInService
  End If
  If TypeCode = TC_RLYDSG Then 
   ParamID = DG_sID
   ParamInService = DG_nInService
  End If
  If TypeCode = TC_RLYDSP Then 
   ParamID = DP_sID
   ParamInService = DP_nInService
  End If
  Call GetData( RelayHnd, ParamInService, nInService& )
  If TypeCode = nDesiredType And nInService = 1 Then GoTo FoundRelay
 Wend  'Each relay
 Print "No relay of desired type is found in this group"
 Stop
FoundRelay: 
 Call GetData( RelayHnd, ParamID, sRlyID$ )
 PrintTTY( "Found relay: " & sRlyID )

 timeStart# = Timer

 nProgress = 5
 nCount = 0 
 For dPcnt = 0.1 to 99.9 step 99.8/5
  sLine$ = "Pcnt = " & Format( dPcnt/100, "#0.#0%" ) & "; T = " & Format( OpTime, "#0.#0s" )
  If 2 = ProgressDialog( 1, "Computing zone reach: DS relay " & sRlyID, sLine$, nProgress ) Then GoTo doStop
  nProgress = nProgress + 10
  If nProgress > 100 Then nProgress = 0
  
  If OneIntmFault(ObjHnd, RelayHnd, 1, dPcnt, OpTime#) = 0 Then GoTo HasError
  
  nCount = nCount + 1
  If nCount = 1 Then
    OpTime1 = OpTime
    dPcnt1 = dPcnt
  Else
   If OpTime <> OpTime1 Then
    dPcnt2 = dPcnt
    OpTime2 = OpTime
    GoTo exitFor
   Else
    OpTime1 = OpTime
    dPcnt1 = dPcnt
   End If
  End If
 Next
 exitFor:
 
 dThreshold# = 0.1
 While OpTime1 <> OpTime2
  If dPcnt2 - dPcnt1 <= dThreshold Then GoTo exitWhile
  dPcnt = (dPcnt2 + dPcnt1) / 2
  If OneIntmFault(ObjHnd, RelayHnd, 1, dPcnt, OpTime#) = 0 Then GoTo HasError

  sLine$ = "Pcnt = " & Format( dPcnt/100, "#0.#0%" ) & "; T = " & Format( OpTime, "#0.#0s" )
  If 2 = ProgressDialog( 1, "DS relay " & sRlyID & " reach", sLine$, nProgress ) Then GoTo doStop
  nProgress = nProgress + 10
  If nProgress > 100 Then nProgress = 5

    nCount = nCount + 1
  If OpTime = OpTime1 Then dPcnt1 = dPcnt Else dPcnt2 = dPcnt
 Wend
 exitWhile:
 timeEnd# = Timer
 
 aLine$ = "nCount=" & Str(nCount) & " dTime=" & Format( timeEnd-timeStart, "#0.#0s" ) _
          & " dPcnt1=" & Format( dPcnt1/100, "#0.#0%" )  & " OpTime1=" & Format( OpTime1, "#0.#0s" ) & _
          " dPcnt2=" & Format( dPcnt2/100, "#0.#0%" )  & " OpTime2=" & Format( OpTime2, "#0.#0s" )
 PrintTTY( aLine )
 Print aLine
doStop:
  Call ProgressDialog( 0, "", "", 0 )
Exit Sub
HasError:
  Call ProgressDialog( 0, "", "", 0 )
  Print "Error: ", ErrorString( )
End Sub

Function OneIntmFault( ByVal nGroupHnd, ByVal nRlyHnd&, ByVal nConn&, ByVal dPcnt#, ByRef OpTime# ) As long
 Dim FltConn(4) As Long
 Dim FltOpt(15) As Double
 Dim OutageOpt(4) As Long
 Dim OutageLst(30) As Long
 
 doOneFault = 0
 ' fault connections
 FltConn(nConn) = 1
 Rflt       = 0   ' Fault R
 Xflt       = 0   ' Fault X
 ClearPrev  = 0   ' Not clear previous results will make the script run faster.
 For ii = 1 to 15
  FltOpt(ii)  = 0
 Next 
 FltOpt(9)  = dPcnt  ' Intermediate %
 If DoFault( nGroupHnd, FltConn, FltOpt, OutageOpt, OutageLst, Rflt, Xflt, ClearPrev ) = 0 Then GoTo HasError1
 If PickFault( SF_LAST ) = 0 Then GoTo HasError1
' Dim vnShowRelay(4)

' If ShowFault( 1, 2, 4, 0, vnShowRelay ) = 0 Then GoTo HasError1


 If GetRelayTime( nRlyHnd, 1.0, OpTime# ) = 0 Then GoTo HasError1
 OneIntmFault = 1
 exit Function
 HasError1:
End Function
