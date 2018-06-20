' ASPEN PowerScript sample program
'
' BKRFAILURE.BAS
'
' Run DoFault() with breaker failure simulation option
'
' Version 1.0
' Category: OneLiner
'
' PowerScript functions called:
'
'
Sub main()
   ' Variable declaration
   Dim FltConn(4) As Long
   Dim FltOption(14) As Double
   Dim OutageList(20) As Long
   Dim BranchList(50) As String, BusList(3) As String
   dim LineZList(50) As double
   Dim OutageType(4) As Long

  If GetEquipment( TC_PICKED, DevHandle& ) = 0 Then
    Print "Please select a relay group"
    Exit Sub
  End If
  
  If EquipmentType( DevHandle& ) <> TC_RLYGROUP Then
    Print "Please select a relay group"
    Exit Sub
  End If

  Call GetData( DevHandle, RG_nBranchHnd, RlyBranchHnd& )
  Call GetData( RlyBranchHnd, BR_nBus1Hnd, BusHnd& )

  ' Inventory of all lines at the bus
   
  BrHnd&   = 0
  BrCounts = 0
  While GetBusEquipment( BusHnd, TC_BRANCH, BrHnd ) > 0
    If BrHnd = RlyBranchHnd then goto contWhile
    Call GetData( BrHnd, BR_nType, BrType& )
    Call GetData( BrHnd, BR_nInService, nFlag& )
    If nFlag = 1 And BrType = TC_LINE Then
       Call GetData( BrHnd, BR_nBus2Hnd, BusHnd2& )
       BrHndStr$ = BrHnd
       BrList$ = "[" + BrHndStr$ + "] "
       BrList$ = BrList$ + FullBusName(FindLineRemoteBus(BrHnd, dLineZ#))
       BranchList(BrCounts) = BrList$
       BrCounts = BrCounts + 1
    End If
  contWhile:
  Wend
   
  If BrCounts < 2 Then
    Print "There must be 3 or more branches at the bus"
    Stop
  End If
  
'=============Dialog Spec=============
Begin Dialog DIALOG_1 134,63, 203, 154, "Select Breaker's 2nd Line"
  Text 8,8,84,8, "Line to"
  ListBox 8,20,184,108, BranchList(), .ListBox_1
  CancelButton 120,136,40,12
  OKButton 44,136,68,12
End Dialog

'=====================================

   Dim Dlg As Dialog_1

   For ii = 1 To 4 
     FltConn(ii) = 0
   Next 
   For ii = 1 To 12
     FltOption(ii) = 0.0
   Next
   For ii = 1 To 3
     OutageType(ii) = 0
   Next
   
   ' show the dialog
   Button = Dialog( dlg )
   If Button = 0 Then Exit Sub	' Canceled
   
   ' Initialize DoFault options using dialog data
   ' Fault connection
   FltConn(1)    = 1	' 3LG 
   FltConn(3)    = 1	' 1LG 
   
   ' Fault type
   FltOption(8)  = 1   ' Line end fault with outage
   OutageType(4) = 1	' Breaker failure - stub


   ' Extract handle numbers and prepare the outage list
   StrLine$ = BranchList(Dlg.ListBox_1)
   nPos = InStr( 1, StrLine, "]" )
   OutageList(1) = Val( Mid(StrLine, 2, nPos-2) )
   OutageList(2) = Val( RlyBranchHnd )
   OutageList(3) = 0
   

   ' Simulate fault
   If 0 = DoFault( RlyBranchHnd, FltConn, FltOption, OutageType, OutageList, 0.0, 0.0, 1 ) Then GoTo HasError

                          
   Dim vnShowRelay(4)

   If ShowFault( 1, 0, 4, 0, vnShowRelay ) = 0 Then GoTo HasError

   
   Exit Sub
HasError:
   Print "Error: ", ErrorString( )
   Close
End Sub

Function FindLineRemoteBus( ByVal Branch1Hnd&, ByRef dLineZ# ) As String

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

