' ASPEN PowerScrip sample program
'
' Z1Check.BAS
'
' Find distance relay zone 1 is over reaching by checking relay operating time
' in select bus 3ph&1ph fault.
'
' Version: 1.0
' Category: OneLiner
'
' Global variables
'
Const MaxBus = 500
Const MaxRlyGrp = 500
Const nTier = 3  ' define tier number
Const FilePath = "c:\Users\xu\Desktop\Z1Check.csv"  'output report file path
Const ExcludeBus = 0 ' 0 = Run fault on all selected buses; 1= exclude bus with no local relays
Dim vnBusHnd1(2) As long, vnBusHnd2(200) As long
Dim TierBusHnd(MaxBus) As Long
Dim RlyGrpHnd(MaxRlyGrp) As Long
Dim FltConn(4) As Long
Dim FltOpt(14) As Double
Dim OutageOpt(4) As Long
Dim OutageLst(30) As Long

Sub main()
   ' fault bus selection
   sWindowText$ = "Select bus to check (200 or fewer)"
   vnBusHnd1(1) = 0
   nPicked& = BusPicker( sWindowText$, vnBusHnd1, vnBusHnd2 )
   If nPicked = 0 Then exit Sub
      
   ' open output *csv file  
   CsvFile$ = FilePath
   Open CsvFile For Output As 1
   Print #1, "Aspen Onliner Program Zone 1 Over Reaching Check Report"
   Print #1, "Relay Group" & "," & "Relay ID" & "," & "Relay Type" & "," & "Operating Time (s)" & "," & "Flag" & "," & "Fault Bus" & "," & "Fault type (3PH/1PH)" 
   
   ' fault connections
   FltConn(1) = 1   ' Do 3PH
   FltConn(2) = 0   ' Do 2LG
   FltConn(3) = 1   ' Do 1LG
   FltConn(4) = 0   ' Do LL
   FltOpt(1)  = 1   ' Bus fault no outage
   FltOpt(2)  = 0
   Rflt       = 0   ' Fault R
   Xflt       = 0   ' Fault X
   ClearPrev  = 1   ' Clear previous result initially
      
   ' do bus fault one by one   
   For ii& = 1 to nPicked
      BusHnd = vnBusHnd2(ii)
      
      If ExcludeBus = 1 Then
         BranchHnd = 0
         While GetBusEquipment( BusHnd, TC_BRANCH, BranchHnd& ) > 0
            sTemp = FullBranchString( BranchHnd )
            If GetData( BranchHnd, BR_nInservice, nFlag& ) = 0 Then Exit Function
            If nFlag = 1 Then
               If GetData( BranchHnd, BR_nType, BrType& ) = 0 Then Exit Function
               If BrType = TC_LINE Then  ' branch has to be a line
                  If GetData( BranchHnd, BR_nRlyGrp1Hnd, RlyGrp1Hnd& ) > 0 Then GoTo BusFault
               End If
            End If
         Wend
         GoTo Continue
      End If
      
      BusFault:
      ' find remote relay or bus, goto next tier if bus is on the remote side, 
      nTierIndexStart  = 0
      nTierIndexEnd    = 1
      nTierIndexUpdate = 1
      TierBusHnd(1) = BusHnd
      nRlyGrpIndexUpdate  = 0
      
      For nTierIndex = 1 to nTier
         For nBusIndex = nTierIndexStart+1 to nTierIndexEnd
            BusHnd = TierBusHnd( nBusIndex )    
            nCheck = SearchRlyGroup( BusHnd, nTierIndexUpdate, TierBusHnd, nRlyGrpIndexUpdate, RlyGrpHnd )
            If nCheck = 0 Then GoTo HasError
            If nCheck = 2 Then
               Print "Tier bus number exceeds limit"
               Exit Sub
            End If
            If nCheck = 3 Then
               Print "Relay group number exceeds limit"
               Exit Sub  
            End If
         Next nBusIndex 
         nTierIndexStart = nTierIndexEnd
         nTierIndexEnd   = nTierIndexUpdate
      Next nTierIndex
   
      ' check tier buses in tty window
      printTTY("Tier Buses:")
      For nBusIndex = 1 to nTierIndexUpdate
         BusHnd = TierBusHnd( nBusIndex )
         Call GetData( BusHnd, Bus_sName, sName$ )
         If nBusIndex = 1 Then
            sFltBusName = sName
         End If
         printTTY ( sName )
      Next nBusIndex
      printTTY("")
      ' check relay group in tty window
      printTTY("Relay Group:")
      For nRlyGrpIndex = 1 to nRlyGrpIndexUpdate
         PrintDSRly( RlyGrpHnd( nRlyGrpIndex ) )
      Next nRlyGrpIndex

      If DoFault( TierBusHnd(1), FltConn, FltOpt, OutageOpt, OutageLst, _
                  Rflt, Xflt, ClearPrev ) = 0 _
         Then GoTo HasError

      ' check whether z1 is overreaching by comparing operation time
   	  FaultFlag = 1
      While PickFault( FaultFlag ) <> 0
         FltString$ = FaultDescription()     
         For nRlyGrpIndex = 1 to nRlyGrpIndexUpdate
            Call GetData( RlyGrpHnd( nRlyGrpIndex ), RG_nBranchHnd, BranchHnd& )
            BranchName = FullBranchString( RlyGrpHnd( nRlyGrpIndex ) )
            RelayHnd& = 0
   		    While GetRelay( RlyGrpHnd( nRlyGrpIndex ), RelayHnd ) > 0 		        
               TypeCode& = EquipmentType( RelayHnd )
               If TypeCode = TC_RLYDSP Then
                  If GetData( RelayHnd, DP_nInService, nFlag& ) = 0 Then GoTo HasError
                  If nFlag = 1 Then
                     If GetRelayTime( RelayHnd, 1.0, dTime# ) = 0 Then GoTo HasError	
                     If GetData( RelayHnd, DP_sID, sID$ ) = 0 Then GoTo HasError
                     sType$ = "Z1 Overreach"
                     If dTime = 0 Then 
                        If InStr( 1, FltString$, " 3LG" ) > 0 Then
                           aString = BranchName & "," & sID & "," & "Phase Distance Relay" & "," & Format( dTime, "#0.#0" ) & "," & sType & "," & sFltBusName & "," & "3PH"
                        ElseIf InStr( 1, FltString$, " 1LG" ) > 0 Then
                           aString = BranchName & "," & sID & "," & "Phase Distance Relay" & "," & Format( dTime, "#0.#0" ) & "," & sType & "," & sFltBusName & "," & "1PH"
                        End If
                        Print #1, aString
                     End If
                  End If
	           End If
	           If TypeCode = TC_RLYDSG Then
	              If GetData( RelayHnd, DG_nInService, nFlag& ) = 0 Then GoTo HasError
	              If nFlag = 1 Then
	                 If GetRelayTime( RelayHnd, 1.0, dTime# ) = 0 Then GoTo HasError	
                     If GetData( RelayHnd, DG_sID, sID$ ) = 0 Then GoTo HasError
                     sType$ = "Z1 Overreach"
                     If dTime = 0 Then
                        If InStr( 1, FltString$, " 3LG" ) > 0 Then
                           aString = BranchName & "," & sID & "," & "Ground Distance Relay" & "," & Format( dTime, "#0.#0" ) & "," & sType & "," & sFltBusName & "," & "3PH"
                        ElseIf InStr( 1, FltString$, " 1LG" ) > 0 Then
                           aString = BranchName & "," & sID & "," & "Ground Distance Relay" & "," & Format( dTime, "#0.#0" ) & "," & sType & "," & sFltBusName & "," & "1PH"
                        End If
                        Print #1, aString
                     End If
	              End If    
	           End If    
            Wend 
         Next nRlyGrpIndex
         FaultFlag = SF_NEXT	' Show next fault
      Wend   
   Continue:  
   Next               
   Close 1
   Print "Report is saved as " & FilePath
Exit Sub
HasError:
   Close 1
   Print "Error: ", ErrorString( )   
End Sub

Function CheckOneBus( ByVal BusHnd, ByVal nTier ) As long

End Function

' Search relay group on the remote end, if not, save remote bus or next round search
Function SearchRlyGroup( ByVal BusHnd, ByRef nBusIndex, ByRef TierBusHnd() As long, ByRef nRlyGrpIndex, ByRef RlyGrpHnd() As long ) As long
   SerchRlyGroup = 0
   BranchHnd = 0
   nStartIndex = nBusIndex
   While GetBusEquipment( BusHnd, TC_BRANCH, BranchHnd& ) > 0
   	  If GetData( BranchHnd, BR_nInservice, nFlag& ) = 0 Then Exit Function
      If nFlag = 1 Then
         If GetData( BranchHnd, BR_nType, BrType& ) = 0 Then Exit Function
         If BrType = TC_LINE Then  ' branch has to be a line
            If GetData( BranchHnd, BR_nBus2Hnd, Bus2Hnd& ) = 0 Then Exit Function
            For nIndex = 1 to nStartIndex
               If TierBusHnd(nIndex) = Bus2Hnd Then GoTo Continue
            Next nIndex
            If GetData( BranchHnd, BR_nRlyGrp2Hnd, RlyGrp2Hnd& ) <= 0 Then         
               nBusIndex = nBusIndex + 1
               If nBusIndex > MaxBus Then 
                  SerchRlyGroup = 2
                  Exit Function
               End If
               TierBusHnd(nBusIndex) = Bus2Hnd
            Else
               If CheckRlyGroup( RlyGrp2Hnd ) = 1 Then
                  For nIndex = 1 to nRlyGrpIndex
                     If RlyGrpHnd(nIndex) = RlyGrp2Hnd Then GoTo Continue
                  Next nIndex
                  nRlyGrpIndex = nRlyGrpIndex + 1
                  If nRlyGrpIndex > MaxRlyGrp Then 
                     SerchRlyGroup = 3
                     Exit Function
                  End If
                  RlyGrpHnd(nRlyGrpIndex) = RlyGrp2Hnd
               End If          
            End If
         End If
      End If
   Continue:   
   Wend
   SearchRlyGroup = 1
End Function

' Check relay group include distance relay or not 
Function CheckRlyGroup( ByVal GroupHnd ) As long
   ' Group must have a phase/ground DS relay
   ChekRlyGroup = 0
   RelayHnd& = 0
   While GetRelay( GroupHnd, RelayHnd ) > 0
      TypeCode& = EquipmentType( RelayHnd )
      If TypeCode = TC_RLYDSP Or TypeCode = TC_RLYDSG Then 
         CheckRlyGroup = 1
         Exit Function
      End If
   Wend
End Function

Function PrintDSRly( ByVal GroupHnd ) As long
   PrintDSRly = 0
   RelayHnd& = 0
   While GetRelay( GroupHnd, RelayHnd ) > 0
      TypeCode& = EquipmentType( RelayHnd )
      If TypeCode = TC_RLYDSP Or TypeCode = TC_RLYDSG Then 
         If TypeCode = TC_RLYDSP Then
            If GetData( RelayHnd, DP_sID, sID ) = 0 Then Exit Function
            PrintTTY( sID )
         End If
         If TypeCode = TC_RLYDSG Then
            If GetData( RelayHnd, DG_sID, sID ) = 0 Then Exit Function
            PrintTTY( sID )
         End If
      End If
   Wend  'Each relay
   PrintDSRly = 1
End Function

Function FullBranchString( ByVal nHnd& ) As String
  nBranchHnd& = 0
  nType = EquipmentType( nHnd )
  If nType = TC_RLYGROUP Then
    Call GetData( nHnd, RG_nBranchHnd, nBranchHnd& )
  Else 
    If nType = TC_BRANCH Then nBranchHnd& = nHnd
  End If
  FullBranchString = ""
  If nBranchHnd = 0 Then exit Sub
  Call GetData( nBranchHnd, BR_nBus1Hnd, nBus1Hnd& )
  Call GetData( nBranchHnd, BR_nBus2Hnd, nBus2Hnd& )
  Call GetData( nBranchHnd, BR_nType, nBrType& )
  Call GetData( nBranchHnd, BR_nHandle, nHandle& )
  sType$ = ""
  select case nBrType
   case TC_LINE
    sType$ = "L"
    Call GetData( nHandle, LN_sID, sID$ )
   case TC_XFMR
    sType$ = "T"
    Call GetData( nHandle, XR_sID, sID$ )
   case TC_XFMR3
    sType$ = "X"
    Call GetData( nHandle, X3_sID, sID$ )
   case TC_PS
    sType$ = "P"
    Call GetData( nHandle, PS_sID, sID$ )
  End select
  If sType = "" Then exit Sub
  FullBranchString = FullBusName(nBus1Hnd) & "-" & FullBusName(nBus2Hnd) & " " & sID & sType
End Function


