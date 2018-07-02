' ASPEN PowerScript Sample Program
'
' SIR2.BAS
'
' Compute source to line impedance ratio
'
' !!!!! Warning !!!!!!!!!!
'   The source impedance calculation in this script does not always
'   produce accurate results. See SIR.BAS script for correct SIR calculation.
' !!!!!!!!!!!!!!!!!!!!!!!!
'
' Version: 1.0
' Category: OneLiner
'
'
Sub main()
   ' Variable declaration
   Dim FltConn(4) As Long
   Dim FltOption(14) As Double
   Dim OutageList(20) As Long
   Dim BrHndList(50) As Long
   Dim BranchList(50) As String, BusList(3) As String
   Dim OutageType(3) As Long
   
   ' Make sure a bus is selected
   If GetEquipment( TC_PICKED, BusHnd& ) = 0 Or _
      EquipmentType( BusHnd& ) <> TC_BUS Then
     Print "Please select a bus"
     Exit Sub
   End If
  
   ' Inventory of all lines at the bus
   
   BrHnd&   = 0
   BrCounts = 0
   While GetBusEquipment( BusHnd, TC_BRANCH, BrHnd ) > 0
     Call GetData( BrHnd, BR_nType, BrType& )
     Call GetData( BrHnd, BR_nInService, nFlag& )
     If nFlag = 1 And BrType = TC_LINE Then
       Call GetData( BrHnd, BR_nBus2Hnd, BusHnd2& )
       BrHndList(BrCounts) = BrHnd
       BrHndStr$ = BrHnd
       BrList$ = "[" + BrHndStr$ + "] "
       BrList$ = BrList$ + FullBusName(FindLineRemoteBus(BrHnd))
       BranchList(BrCounts) = BrList$
       BrCounts = BrCounts + 1
     End If
   Wend
   
   If BrCounts = 0 Then
     Print "No active line found at the bus"
     Stop
   End If

'=============Dialog Spec=============
Begin Dialog DIALOG_1 134,63, 203, 154, "Source to Line Impedance Ratio"
  Text 8,8,84,8, "Line to"
  ListBox 8,20,184,108, BranchList(), .ListBox_1
  CancelButton 120,136,40,12
  OKButton 44,136,68,12
End Dialog

'=====================================

   Dim Dlg As Dialog_1

   ' show the dialog
   Button = Dialog( dlg )
   If Button = 0 Then Exit Sub	' Canceled
   
   ' Initialize DoFault options using dialog data
   For ii = 1 To 4 
     FltConn(ii) = 0
   Next 
   For ii = 1 To 12
     FltOption(ii) = 0.0
   Next
   For ii = 1 To 3
     OutageType(ii) = 0
   Next
   
   ' Fault connection
   FltConn(1)    = 1	' 3PH 
   ' Fault type
   FltOption(2)  = 1   ' Bus fault with outage


   ' Extract handle numbers and prepare the outage list
   StrLine$ = BranchList(Dlg.ListBox_1)
   BranchHnd = BrHndList(Dlg.ListBox_1)
   nPos = InStr( 1, StrLine, "]" )
   OutageList(1) = Val( Mid(StrLine, 2, nPos-2) )
   OutageList(2) = 0
   OutageType(1) = 1	' Outage one at a time

   ' Simulate fault
   If 0 = DoFault( BusHnd, FltConn, FltOption, OutageType, OutageList, 0.0, 0.0, 1 ) Then GoTo HasError
                       
   ' Must alway pick a fault before getting V and I
   If PickFault( 1 ) = 0 Then GoTo HasError
   Call GetData( HND_SC, FT_dRPt, R1t# )
   Call GetData( HND_SC, FT_dXPt, X1t# )
   
   dZs# = Sqr(R1t*R1t + X1t*X1t)

   For ii = 1 To 4 
     FltConn(ii) = 0
   Next 
   For ii = 1 To 12
     FltOption(ii) = 0.0
   Next
   For ii = 1 To 20
     OutageList(ii) = 0
   Next
   For ii = 1 To 3
     OutageType(ii) = 0
   Next
   
   ' Fault connection
   FltConn(1)    = 1	' 3PH 
   ' Fault type
   FltOption(7)  = 1    ' Close-in fault with end opened
     
   ' Simulate fault
   If 0 = DoFault( BranchHnd, FltConn, FltOption, OutageType, OutageList, 0.0, 0.0, 1 ) Then GoTo HasError
                       
   ' Must alway pick a fault before getting V and I
   If PickFault( 1 ) = 0 Then GoTo HasError
   Call GetData( HND_SC, FT_dRPt, R2t# )
   Call GetData( HND_SC, FT_dXPt, X2t# )

   dZl# = Sqr((R2t-R1t)*(R2t-R1t) + (X2t-X1t)*(X2t-X1t))
   
   ' Print output to TTY
   PrintTTY( "Line: " & FullBusName( BusHnd ) & " - " & Mid(StrLine$, nPos+1, 99 ) )
   strOut$ = " Line Z = " & Format(dZl#, "0.00 Ohm") & _
              " Source Z = " & Format(dZs#, "0.00 Ohm") & _
              " SIR = " & Format(dZs#/dZl#, "0.000") 
   PrintTTY(strOut$)
   Print strOut$
   
   Exit Sub
HasError:
   Print "Error: ", ErrorString( )
   Close
End Sub

Function FindLineRemoteBus( ByVal Branch1Hnd& ) As String
 
  ' Skip all taps on the line
  Do 
    Call GetData( Branch1Hnd, BR_nHandle, LineHnd )
    Call GetData( LineHnd, LN_sName, LineName )
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
  
  FindLineRemoteBus = Bus1Hnd
End Function
