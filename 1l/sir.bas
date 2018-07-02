' ASPEN PowerScript Sample Program
'
' SIR.BAS
'
' Compute source to line impedance ratio
'
'   SIR = Z_source/Z_line
'
'   Z_source_3PH = (V_relay_prefault - V_relay_faulted)/I_relay
'   Z_source_SLG = (V_relay_prefault - V_relay_faulted)/(I_relay + 3*Io_relay*Ko)
'
' where 
'   Z_line: line positive sequence impedance
'   V_relay_prefault: pre-fault positive sequence relay voltage
'   V_relay_faulted:  positive sequence relay voltage in fault on remote bus
'   I_relay: positive sequence relay current in fault on remote bus
'   Io_relay: zero sequence relay current in SLG fault on remote bus
'
'
' Version 2.0
' Category: OneLiner
'
'
Sub main()
   ' Variable declaration
   Dim FltConn(4) As Long
   Dim FltOption(14) As Double
   Dim OutageType(3) As Long
   Dim OutageList(20) As Long
   Dim BranchList(50) As String
   Dim RmtBusList(50) As Long
   dim LineZ1RList(50) As double
   dim LineZ1XList(50) As double
   dim LineZ0RList(50) As double
   dim LineZ0XList(50) As double
   
   Dim LineList(50) As Long
   Dim MagArray(3) As Double
   Dim AngArray(3) As Double
   Dim VSeqReal(3) As Double
   Dim VSeqImag(3) As Double
   Dim ISeqReal(3) As Double
   Dim ISeqImag(3) As Double    
   
   
   
   ' Make sure a bus is selected
   If GetEquipment( TC_PICKED, BusHnd& ) = 0 Or _
      EquipmentType( BusHnd& ) <> TC_BUS Then
     Print "A bus must be selected."
     Exit Sub
   End If
   
   Call GetData( HND_SYS, SY_dBaseMVA, dMVA# )  
   Call GetData( BusHnd, BUS_dKVnorminal, dKV# )
   dZbase = dKV*dKV/dMVA 
   
   ' Inventory of all lines at the bus
   BrHnd&   = 0
   BrCounts = 0
   While GetBusEquipment( BusHnd, TC_BRANCH, BrHnd ) > 0
     Call GetData( BrHnd, BR_nType, BrType& )
     Call GetData( BrHnd, BR_nInService, nFlag& )
     If nFlag = 1 And BrType = TC_LINE Then
       Call GetData( BrHnd, BR_nBus2Hnd, BusHnd2& )
       BrHndStr$ = BrHnd
       BrList$ = "[" + BrHndStr$ + "] "
       RmtBusHnd = FindLineRemoteBus(BrHnd, dLineZ1R#, dLineZ1X#, dLineZ0R#, dLineZ0X#)
       Call GetData( BrHnd, BR_nHandle, LineHnd )
       BrList$ = BrList$ + FullBusName(RmtBusHnd)
       BranchList(BrCounts) = BrList$
       RmtBusList(BrCounts) = RmtBusHnd
       LineZ1RList(BrCounts) = dLineZ1R
       LineZ1XList(BrCounts) = dLineZ1X
       LineZ0RList(BrCounts) = dLineZ0R
       LineZ0XList(BrCounts) = dLineZ0X
       LineList(BrCounts) = LineHnd
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
   StrLine$ = BranchList(Dlg.ListBox_1)
   
   ' Get remote bus handle
   FltBusHnd = RmtBusList(Dlg.ListBox_1)
   
   ' Calculate line impedance
   dZ1R# = LineZ1RList(Dlg.ListBox_1)
   dZ1X# = LineZ1XList(Dlg.ListBox_1)
   dZl# = Sqr(dZ1R*dZ1R + dZ1X*dZ1X)*dZbase
   
   ' Initialize DoFault options using dialog data
   For ii = 1 To 4 
     FltConn(ii) = 0
   Next 
   For ii = 1 To 14
     FltOption(ii) = 0.0
   Next
   For ii = 1 To 3
     OutageType(ii) = 0
   Next
   
   ' Fault connection
   FltConn(1)    = 1	' 3PH 
   FltConn(3)    = 1    ' 1LG
   ' Fault type
   FltOption(1)  = 1    ' Close-in

   ' Simulate 3LG fault
   If 0 = DoFault( FltBusHnd, FltConn, FltOption, OutageType, OutageList, 0.0, 0.0, 1 ) Then GoTo HasError
                        
   ' Must alway pick a fault before getting V and I
   FaultFlag = 1 
   ' 3LG
   If PickFault( FaultFlag ) = 0 Then GoTo HasError
   
   ' get prefault voltage
   If GetPSCVoltage( BusHnd, MagArray, AngArray, 1 ) = 0 Then	' get prefault voltage on bus
      Print "Get Bus prefault voltage failed."
      Exit Function
   End If    
   Vpre_Real = MagArray(1)*Cos(AngArray(1) *3.14159/180.0)*1000
   Vpre_Imag = MagArray(1)*Sin(AngArray(1) *3.14159/180.0)*1000
    
   ' get postfault voltage and current 
   If GetSCVoltage( BusHnd, VSeqReal, VSeqImag, 1 ) = 0 Then	' get postfault voltage on bus
      Print "Get Bus postfault voltage failed."
      Exit Function
   End If 
   If GetSCCurrent( LineList(Dlg.ListBox_1), ISeqReal, ISeqImag, 1 ) = 0 Then	' get postfault current on line
      Print "Get Bus postfault voltage failed."
      Exit Function
   End If 
   ' 3LG Vpost
   Vpost_Real = VSeqReal(2)*1000
   Vpost_Imag = VSeqImag(2)*1000
   ' 3LG Ipost
   Ipost_Real = ISeqReal(2)
   Ipost_Imag = ISeqImag(2)
   ' 3LG Vdrop
   Call ComplexSub( Vpre_Real, Vpre_Imag, Vpost_Real, Vpost_Imag, Vdrop_Real#, Vdrop_Imag# )
   ' 3LG souce impedance
   Call ComplexDiv( Vdrop_Real, Vdrop_Imag, Ipost_Real, Ipost_Imag, Zs_Real#, Zs_Imag# )
   dZs_3LG# = Sqr(Zs_Real*Zs_Real + Zs_Imag*Zs_Imag)
   ' 3PH SIR 
   SIR_3LG = dZs_3LG/dZl
   
   ' 1LG
   FaultFlag = SF_NEXT
   If PickFault( FaultFlag ) = 0 Then GoTo HasError
   ' get postfault voltage and current 
   If GetSCVoltage( BusHnd, VSeqReal, VSeqImag, 1 ) = 0 Then	' get postfault voltage on bus
      Print "Get Bus postfault voltage failed."
      Exit Function
   End If 
   If GetSCCurrent( LineList(Dlg.ListBox_1), ISeqReal, ISeqImag, 1 ) = 0 Then	' get postfault current on line
      Print "Get Bus postfault voltage failed."
      Exit Function
   End If 
   ' 1LG Vpost
   Vpost_Real = VSeqReal(2)*1000
   Vpost_Imag = VSeqImag(2)*1000
   ' Calculate 3K0
   dZ0R# = LineZ0RList(Dlg.ListBox_1)
   dZ0X# = LineZ0XList(Dlg.ListBox_1)
   Call ComplexSub(dZ0R, dZ0X, dZ1R, dZ1X, dAR#, dAX#)
   Call ComplexDiv(dAR, dAX, dZ1R, dZ1X, d3K0_Real#, d3K0_Imag#)
   ' 1LG Ipost
   Call ComplexMul( ISeqReal(1), ISeqImag(1), d3K0_Real, d3K0_Imag, Itemp_Real#, Itemp_Imag# )
   Call ComplexAdd( ISeqReal(2), ISeqImag(2), Itemp_Real, Itemp_Imag, Ipost_Real#, Ipost_Imag# )
   ' 1LG Vdrop
   Call ComplexSub( Vpre_Real, Vpre_Imag, Vpost_Real, Vpost_Imag, Vdrop_Real#, Vdrop_Imag# )
   ' 1LG source impedance
   Call ComplexDiv( Vdrop_Real, Vdrop_Imag, Ipost_Real, Ipost_Imag, Zs_Real#, Zs_Imag# )
   dZs_1LG# = Sqr(Zs_Real*Zs_Real + Zs_Imag*Zs_Imag)
   ' 1LG SIR 
   SIR_1LG = dZs_1LG/dZl
     
   ' Print output to TTY
   PrintTTY( "Line: " & FullBusName( BusHnd ) & " - " & Mid(StrLine$, nPos+1, 99 ) )
   strOut$ = " Line Z = " & Format(dZl#, "0.00 Ohm") & _
              "    Source Z_3LG = " & Format(dZs_3LG#, "0.00 Ohm") & _
              "    SIR_3LG = " & Format(SIR_3LG#, "0.000") & _   
              "    Source Z_1LG = " & Format(dZs_1LG#, "0.00 Ohm") & _
              "    SIR_1LG = " & Format(SIR_1LG#, "0.000") 
   PrintTTY(strOut$)
   strOut$ =  "Line Z = " & Format(dZl#, "0.00 Ohm") & Chr(13) & _
              "Source Z_3LG = " & Format(dZs_3LG#, "0.00 Ohm") & _
              "   SIR_3LG = " & Format(SIR_3LG#, "0.000") & Chr(13) & _   
              "Source Z_1LG = " & Format(dZs_1LG#, "0.00 Ohm") & _
              "   SIR_1LG = " & Format(SIR_1LG#, "0.000") & Chr(13) & Chr(13) & _
              "Details are in TTY window."
   Print strOut$
   
   Exit Sub
HasError:
   Print "Error: ", ErrorString( )
   Close
End Sub

Function FindLineRemoteBus( ByVal Branch1Hnd&, ByRef dLineZ1R#, ByRef dLineZ1X#, ByRef dLineZ0R#, ByRef dLineZ0X# ) As String

  dLineZ1R = 0.0
  dLineZ1X = 0.0
  dLineZ0R = 0.0
  dLineZ0X = 0.0
  
  ' Skip all taps on the line
  Do 
    Call GetData( Branch1Hnd, BR_nHandle, LineHnd )
    Call GetData( LineHnd, LN_sName, LineName )
    Call GetData( LineHnd, LN_dR, LineR1 )
    Call GetData( LineHnd, LN_dX, LineX1 )
    Call GetData( LineHnd, LN_dR0, LineR0 )
    Call GetData( LineHnd, LN_dX0, LineX0 )
    dLineZ1R = dLineZ1R + LineR1
    dLineZ1X = dLineZ1X + LineX1
    dLineZ0R = dLineZ0R + LineR0
    dLineZ0X = dLineZ0X + LineX0
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

Function ComplexDiv( ByVal X_R#, ByVal X_I#, ByVal Y_R#, ByVal Y_I#, ByRef Z_R#, ByRef Z_I# )
' Z = X/Y
  Z_R = 0.0
  Z_I = 0.0
  Z_R = (X_R*Y_R + X_I*Y_I)/(Y_R*Y_R + Y_I*Y_I)
  Z_I = (X_I*Y_R - X_R*Y_I)/(Y_R*Y_R + Y_I*Y_I)
End Function

Function ComplexMul( ByVal X_R#, ByVal X_I#, ByVal Y_R#, ByVal Y_I#, ByRef Z_R#, ByRef Z_I# )
' Z = X*Y
  Z_R = 0.0
  Z_I = 0.0
  Z_R = X_R*Y_R - X_I*Y_I
  Z_I = X_R*Y_I + X_I*Y_R 
End Function

Function ComplexAdd( ByVal X_R#, ByVal X_I#, ByVal Y_R#, ByVal Y_I#, ByRef Z_R#, ByRef Z_I# )
' Z = X+Y
  Z_R = 0.0
  Z_I = 0.0
  Z_R = X_R + Y_R
  Z_I = X_I + Y_I 
End Function

Function ComplexSub( ByVal X_R#, ByVal X_I#, ByVal Y_R#, ByVal Y_I#, ByRef Z_R#, ByRef Z_I# )
' Z = X-Y
  Z_R = 0.0
  Z_I = 0.0
  Z_R = X_R - Y_R
  Z_I = X_I - Y_I 
End Function
