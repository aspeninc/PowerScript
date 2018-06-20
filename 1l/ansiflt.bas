' ASPEN PowerScript Sample Program
'
' ANSIFLT.BAS
'
' Simulate fault with R only and X only network to compute ANSI X/R ratio
'
' Version: 1.0
' Category: OneLiner
'
' PowerScript functions called:
'   GetData()
'   SetData()
'   PostData()
'   GetBusEquipment()
'   GetEquipment()
'   DoFault()
'   ShowFault()
'   FaultDescription()
'
' Global constants and variables
Const StorageLimit = 5000  ' Size of temp storage for X and R
Const XRfactor = 0.03      ' if R=0, R = XRfactor * X
Const SmallX   = 0.0001    ' if X=0, X = SmallX
'
Dim TempStorage( StorageLimit ) As Double	' Temporary storage for R and X
Dim FlagArray( StorageLimit ) As Integer	' Zero impedance flag
Dim ArrayIndex&					' Index to the above arrays

Sub main()
   Dim ZtArray(500) As Double
   Dim Fltconn(4) As Long
   Dim FltOption(14) As Double
   Dim OutageOpt(3) As Long
   Dim OutageLst(15) As Long
   Dim ShowRelayOpt(4) As Long
   Dim FltR As Double, FltX As Double

   ' Storage size must be enough for all the impedances in network
   SysSize = 0
   NOdev   = 0
   If GetData( HND_SYS, SY_nNOgen,  NOdev ) = 0 Then GoTo HasError
   SysSize = SysSize + NOdev*4
   If GetData( HND_SYS, SY_nNOline, NOdev ) = 0 Then GoTo HasError
   SysSize = SysSize + NOdev*12
   If GetData( HND_SYS, SY_nNOload, NOdev ) = 0 Then GoTo HasError
   SysSize = SysSize + NOdev*4
   If GetData( HND_SYS, SY_nNOps,   NOdev ) = 0 Then GoTo HasError
   SysSize = SysSize + NOdev*12
   If GetData( HND_SYS, SY_nNOseriescap, NOdev ) = 0 Then GoTo HasError
   SysSize = SysSize + NOdev*4
   If GetData( HND_SYS, SY_nNOshunt, NOdev ) = 0 Then GoTo HasError
   SysSize = SysSize + NOdev*4
   If GetData( HND_SYS, SY_nNOxfmr,  NOdev ) = 0 Then GoTo HasError
   SysSize = SysSize + NOdev*12
   If GetData( HND_SYS, SY_nNOxfmr3, NOdev ) = 0 Then GoTo HasError
   SysSize = SysSize + NOdev*18
   If SysSize > StorageLimit Then
     Print "Not enough space allocated for impedances. StorageLimit should be increased to "; SysSize
     Exit Sub
   End If   

   ' Initialize 
   For ii = 1 To 4 
     Fltconn(ii) = 0
   Next 
   For ii = 1 To 12
     FltOption(ii) = 0.0
   Next
   OutageLst(1) = 0  ' Must terminate the list
   For ii = 1 To 3
     OutageOpt(ii) = 0
   Next
   For ii = 1 To 4
     ShowRelayOpt(ii) = 0
   Next
   FltR        = 0.0
   FltX        = 0.0

   If GetEquipment( TC_PICKED, PickedHnd ) = 0 Then GoTo hasError
   ' Probe to see what's being picked
   DeviceType = EquipmentType( PickedHnd )
   If DeviceType = TC_RLYGROUP Then
     ' Must be a relay group
     If FaultDialog( Fltconn, FltOption, 0, FltR, FltX ) = 0 Then Stop
   ElseIf DeviceType = TC_BUS Then
     ' Must be a bus group
     If FaultDialog( Fltconn, FltOption, 1, FltR, FltX ) = 0 Then Stop
   Else
     Print "Must select a relay group or a bus"
     Stop
   End If

  ' Make R only network
  Call ProcessNetwork( 1 )
  ' Simulate the fault
  If 0 = DoFault( PickedHnd, Fltconn, FltOption, OutageOpt, OutageLst, _
            FltR, FltX, 1 ) Then GoTo HasError
  ' show fault and get thevenin result
  ShowFltSelector = 1
  For NumFaults& = 1 To 500 ' Max 500 faults
    If ShowFault( ShowFltSelector, 1, 4, 0, ShowRelayOpt ) = 0 Then Exit For
    If GetData( HND_SC, FT_dRt, ZtArray(NumFaults) ) = 0 Then GoTo HasError
    ShowFltSelector = SF_NEXT   ' Show next fault in line
  Next   ' Each fault
  ' Make X only network
  Call ProcessNetwork( 2 )
  ' Simulate the fault
  If 0 = DoFault( PickedHnd, Fltconn, FltOption, OutageOpt, OutageLst, _
            FltR, FltX, 1 ) Then GoTo HasError
  ' show fault and compute thevenin X/R result
  ShowFltSelector = 1
  For NumFaults& = 1 To 500 ' Max 500 faults
    If ShowFault( ShowFltSelector, 1, 4, 0, ShowRelayOpt ) = 0 Then Exit For
    If GetData( HND_SC, FT_dXt, Zt ) = 0 Then GoTo HasError
    AnsiXR = Zt/ZtArray(NumFaults)
    FltString$ = FaultDescription( 1 )
    TempString$ = FltString$ & Chr(10) & _
      "     ANSI X/R = " & Format( AnsiXR, "#0.####0" ) & Chr(10) & _
      "-----------------------------------------------------------------" & _
      "-----------------------------------------------------------------"
    rc& = PrintTTY( TempString$ )
    ShowFltSelector = SF_NEXT   ' Show next fault in line
  Next   ' Each fault
  NumFaults = NumFaults - 1
  ' Restore full network
  Call ProcessNetwork( 3 )
  If 0 = DoFault( PickedHnd, Fltconn, FltOption, OutageOpt, OutageLst, _
            FltR, FltX, 1 ) Then GoTo HasError
  ' show fault and get out
  If ShowFault( 1, 1, 4, 0, ShowRelayOpt ) = 0 Then GoTo HasError
  Print FltString$ & Chr(10) & " ANSI X/R = " & Format( AnsiXR, "#0.####0" ) _
         & Chr(10) & "Result also printed in the TTY window"
  Exit Sub
  HasError:
  Print ErrorString()
End Sub ' Main sub

Sub ProcessNetwork( ByVal NetworkType )
   Dim ArrayR(16) As Double
   Dim ArrayX(16) As Double

   ArrayIndex = 0
   ' Loop thru all buses
   BusHandle& = 0
   While NextBusByName( BusHandle& ) > 0
     ' Loop thru all genunits
     DevHandle& = 0
     While GetBusEquipment( BusHandle&, TC_GENUNIT, DevHandle& ) > 0
       If GetData( DevHandle&, GU_nOnLine, StatusFlag& ) = 0 Then GoTo HasError
       If StatusFlag = 1 Then
         If GetData( DevHandle&, GU_vdR, ArrayR() ) = 0 Then GoTo HasError
         If GetData( DevHandle&, GU_vdX, ArrayX() ) = 0 Then GoTo HasError
         Call ProcessZ( ArrayX, ArrayR, 5, NetworkType )
         If SetData( DevHandle&, GU_vdR, ArrayR() ) = 0 Then GoTo HasError
         If SetData( DevHandle&, GU_vdX, ArrayX() ) = 0 Then GoTo HasError
         If PostData( DevHandle& ) = 0 Then GoTo HasError
       End If	' Each active device
     Wend	' Each genunit
     ' Loop thru all shunt units
     ' Loop thru all load units
   Wend	' Each bus
   ' Loop thru all lines
   DevHandle& = 0
   While GetEquipment( TC_LINE, DevHandle& ) > 0
     If GetData( DevHandle&, LN_nInService, StatusFlag& ) = 0 Then GoTo HasError
     If StatusFlag = 1 Then
       ' Do impedances
       If GetData( DevHandle&, LN_dR, ArrayR(1) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, LN_dX, ArrayX(1) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, LN_dR0, ArrayR(2) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, LN_dX0, ArrayX(2) ) = 0 Then GoTo HasError
       Call ProcessZ( ArrayX, ArrayR, 2, NetworkType )
       If SetData( DevHandle&, LN_dR, ArrayR(1) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, LN_dX, ArrayX(1) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, LN_dR0, ArrayR(2) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, LN_dX0, ArrayX(2) ) = 0 Then GoTo HasError
       ' Do admitances
       If GetData( DevHandle&, LN_dG1, ArrayR(1) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, LN_dB1, ArrayX(1) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, LN_dG2, ArrayR(2) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, LN_dB2, ArrayX(2) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, LN_dG10, ArrayR(3) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, LN_dB10, ArrayX(3) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, LN_dG20, ArrayR(4) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, LN_dB20, ArrayX(4) ) = 0 Then GoTo HasError
       Call ProcessY( ArrayX, ArrayR, 4, NetworkType )
       If SetData( DevHandle&, LN_dG1, ArrayR(1) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, LN_dB1, ArrayX(1) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, LN_dG2, ArrayR(2) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, LN_dB2, ArrayX(2) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, LN_dG10, ArrayR(3) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, LN_dB10, ArrayX(3) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, LN_dG20, ArrayR(4) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, LN_dB20, ArrayX(4) ) = 0 Then GoTo HasError
       ' Post the data
       If PostData( DevHandle& ) = 0 Then GoTo HasError
     End If	' Each active device
   Wend	' Each line
   ' Loop thru all mutual pair
   DevHandle& = 0
   While GetEquipment( TC_MU, DevHandle& ) > 0
     ' Do impedances
     If GetData( DevHandle&, MU_dR, ArrayR(1) ) = 0 Then GoTo HasError
     If GetData( DevHandle&, MU_dX, ArrayX(1) ) = 0 Then GoTo HasError
     Call ProcessY( ArrayX, ArrayR, 1, NetworkType )
     If SetData( DevHandle&, MU_dR, ArrayR(1) ) = 0 Then GoTo HasError
     If SetData( DevHandle&, MU_dX, ArrayX(1) ) = 0 Then GoTo HasError
     ' Post the data
     If PostData( DevHandle& ) = 0 Then GoTo HasError
   Wend	' Each Mutual pair
   ' Loop thru all Phase shifter
   DevHandle& = 0
   While GetEquipment( TC_PS, DevHandle& ) > 0
     If GetData( DevHandle&, PS_nInService, StatusFlag& ) = 0 Then GoTo HasError
     If StatusFlag = 1 Then
       If GetData( DevHandle&, PS_dR, ArrayR(1) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, PS_dX, ArrayX(1) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, PS_dR0, ArrayR(2) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, PS_dX0, ArrayX(2) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, PS_dR2, ArrayR(3) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, PS_dX2, ArrayX(3) ) = 0 Then GoTo HasError
       Call ProcessZ( ArrayX, ArrayR, 3, NetworkType )
       If SetData( DevHandle&, PS_dR, ArrayR(1) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, PS_dX, ArrayX(1) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, PS_dR0, ArrayR(2) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, PS_dX0, ArrayX(2) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, PS_dR2, ArrayR(3) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, PS_dX2, ArrayX(3) ) = 0 Then GoTo HasError
       ' Do admitances
       If GetData( DevHandle&, PS_dB,  ArrayX(1) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, PS_dB0, ArrayX(2) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, PS_dB2, ArrayX(3) ) = 0 Then GoTo HasError
       Call ProcessY( ArrayX, ArrayR, 3, NetworkType )
       If SetData( DevHandle&, PS_dB,  ArrayX(1) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, PS_dB0, ArrayX(2) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, PS_dB2, ArrayX(3) ) = 0 Then GoTo HasError
       If PostData( DevHandle& ) = 0 Then GoTo HasError
     End If	' Each active device
   Wend	' Each Phase shifter
   ' Loop thru all xfmr
   DevHandle& = 0
   While GetEquipment( TC_XFMR, DevHandle& ) > 0
     If GetData( DevHandle&, XR_nInService, StatusFlag& ) = 0 Then GoTo HasError
     If StatusFlag = 1 Then
       If GetData( DevHandle&, XR_dR, ArrayR(1) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dX, ArrayX(1) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dR0, ArrayR(2) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dX0, ArrayX(2) ) = 0 Then GoTo HasError
       Call ProcessZ( ArrayX, ArrayR, 2, NetworkType )
       If SetData( DevHandle&, XR_dR, ArrayR(1) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dX, ArrayX(1) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dR0, ArrayR(2) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dX0, ArrayX(2) ) = 0 Then GoTo HasError
       ' Do admitances
       If GetData( DevHandle&, XR_dB,  ArrayX(1) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dB0, ArrayX(2) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dG1, ArrayR(3) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dB1, ArrayX(3) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dG2, ArrayR(4) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dB2, ArrayX(4) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dG10, ArrayR(5) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dB10, ArrayX(5) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dG20, ArrayR(6) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dB20, ArrayX(6) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dRG1, ArrayR(7) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dXG1, ArrayX(7) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dRG2, ArrayR(8) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dXG2, ArrayX(8) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dRGN, ArrayR(9) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, XR_dXGN, ArrayX(9) ) = 0 Then GoTo HasError
       Call ProcessY( ArrayX, ArrayR, 9, NetworkType )
       If SetData( DevHandle&, XR_dB,  ArrayX(1) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dB0, ArrayX(2) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dG1, ArrayR(3) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dB1, ArrayX(3) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dG2, ArrayR(4) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dB2, ArrayX(4) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dG10, ArrayR(5) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dB10, ArrayX(5) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dG20, ArrayR(6) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dB20, ArrayX(6) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dRG1, ArrayR(7) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dXG1, ArrayX(7) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dRG2, ArrayR(8) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dXG2, ArrayX(8) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dRGN, ArrayR(9) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, XR_dXGN, ArrayX(9) ) = 0 Then GoTo HasError
       If PostData( DevHandle& ) = 0 Then GoTo HasError
     End If	' Each active device
   Wend	' Each xfmr
   ' Loop thru all xfmr3
   DevHandle& = 0
   While GetEquipment( TC_XFMR3, DevHandle& ) > 0
     If GetData( DevHandle&, X3_nInService, StatusFlag& ) = 0 Then GoTo HasError
     If StatusFlag = 1 Then
       If GetData( DevHandle&, X3_dRps,  ArrayR(1) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dXps,  ArrayX(1) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dR0ps, ArrayR(2) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dX0ps, ArrayX(2) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dRpt,  ArrayR(3) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dXpt,  ArrayX(3) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dR0pt, ArrayR(4) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dX0pt, ArrayX(4) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dRst,  ArrayR(5) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dXst,  ArrayX(5) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dR0st, ArrayR(6) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dX0st, ArrayX(6) ) = 0 Then GoTo HasError
       Call ProcessZ( ArrayX, ArrayR, 6, NetworkType )
       If SetData( DevHandle&, X3_dRps,  ArrayR(1) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dXps,  ArrayX(1) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dR0ps, ArrayR(2) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dX0ps, ArrayX(2) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dRpt,  ArrayR(3) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dXpt,  ArrayX(3) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dR0pt, ArrayR(4) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dX0pt, ArrayX(4) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dRst,  ArrayR(5) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dXst,  ArrayX(5) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dR0st, ArrayR(6) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dX0st, ArrayX(6) ) = 0 Then GoTo HasError
       ' Do admitances
       If GetData( DevHandle&, X3_dB,   ArrayX(1) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dB0,  ArrayX(2) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dRG1, ArrayR(3) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dXG1, ArrayX(3) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dRG2, ArrayR(4) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dXG2, ArrayX(4) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dRG3, ArrayR(5) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dXG3, ArrayX(5) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dRGN, ArrayR(6) ) = 0 Then GoTo HasError
       If GetData( DevHandle&, X3_dXGN, ArrayX(6) ) = 0 Then GoTo HasError
       Call ProcessY( ArrayX, ArrayR, 6, NetworkType )
       If SetData( DevHandle&, X3_dB,   ArrayX(1) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dB0,  ArrayX(2) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dRG1, ArrayR(3) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dXG1, ArrayX(3) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dRG2, ArrayR(4) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dXG2, ArrayX(4) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dRG3, ArrayR(5) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dXG3, ArrayX(5) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dRGN, ArrayR(6) ) = 0 Then GoTo HasError
       If SetData( DevHandle&, X3_dXGN, ArrayX(6) ) = 0 Then GoTo HasError
       If PostData( DevHandle& ) = 0 Then GoTo HasError
     End If	' Each active device
   Wend	' Each xfmr  
  Exit Sub
  HasError:
  Print ErrorString()
End Sub

Sub ProcessZ( ArrayX() As Double, ArrayR() As Double, _
              ByVal ArraySize&, ByVal NetworkType& )
' Process all impedances in the array
  For ii = 1 To ArraySize
    ArrayIndex = ArrayIndex + 1
    If NetworkType = 1 Then	' R only
      ' Store X
      TempStorage(ArrayIndex) = ArrayX(ii)
      ' Process R
      If ArrayR(ii) = 0 Then 
        ArrayR(ii) = XRfactor * ArrayX(ii)
        FlagArray(ArrayIndex) = 1
      Else
        FlagArray(ArrayIndex) = 0
      End If
      ' Process X
      ArrayX(ii) = 0
    ElseIf NetworkType = 2 Then	' X only
      ' Restore X
      ArrayX(ii) = TempStorage(ArrayIndex)
      ' Restore R
      If FlagArray(ArrayIndex) = 1 Then ArrayR(ii) = 0
      ' Store R
      TempStorage(ArrayIndex) = ArrayR(ii)
      ' Process X
      If ArrayX(ii) = 0 Then
        ArrayX(ii) = SmallX
        FlagArray(ArrayIndex) = 1
      Else
        FlagArray(ArrayIndex) = 0
      End If
      ' Process R
      ArrayR(ii) = 0
    Else
      ' Restore R
      ArrayR(ii) = TempStorage(ArrayIndex)
      ' Restore X
      If FlagArray(ArrayIndex) = 1 Then ArrayX(ii) = 0
    End If
  Next
End Sub

Sub ProcessY( vdValB() As Double, vdValG() As Double, _
              ByVal ArraySize&, ByVal NetworkType& )
' Process all admitance in the array
  For ii = 1 To ArraySize
    ArrayIndex = ArrayIndex + 1
    If NetworkType = 1 Then	' G only
      ' Store B
      TempStorage(ArrayIndex) = vdValB(ii)
      ' Process B
      vdValB(ii) = 0
    ElseIf NetworkType = 2 Then	' X only
      ' Restore B
      vdValB(ii) = TempStorage(ArrayIndex)
      ' Store G
      TempStorage(ArrayIndex) = vdValG(ii)
      ' Process G
      vdValG(ii) = 0
    Else
      ' Restore G
      vdValG(ii) = TempStorage(ArrayIndex)
    End If
  Next
End Sub

Begin Dialog FAULTDLG 48,46, 258, 128, "Specify fault"
  OptionGroup .FLTCONN
    OptionButton 12,16,24,12, "3PH"
    OptionButton 40,16,24,12, "2LG"
    OptionButton 68,16,24,12, "1LG"
    OptionButton 96,16,24,12, "L-L"
  OptionGroup .FLTOPT
    OptionButton 12,36,44,8, "Close-in"
    OptionButton 12,44,100,8, "Close-in with end opened"
    OptionButton 12,52,68,8, "Remote bus"
    OptionButton 12,60,88,8, "Line end"
    OptionButton 12,68,60,8, "Intermediate"
    OptionButton 12,76,112,8, "Intermediate with end opened"
  Text 8,8,64,8, "Fault Connection"
  Text 8,28,64,8, "Fault Location"
  Text 128,8,64,8, "Fault impedance"
  Text 136,20,12,8, "Z="
  TextBox 148,16,40,12, .EditBox_1
  Text 192,20,12,8, "+ j"
  TextBox 204,16,40,12, .EditBox_2
  TextBox 144,68,16,12, .EditBox_3
  Text 128,72,16,8, "At %"
  TextBox 96,84,24,12, .EditBox_4
  Text 124,88,12,8, "To"
  TextBox 136,84,24,12, .EditBox_5
  Text 24,88,68,8, "Auto sequence from"
  OKButton 72,104,48,12
  CancelButton 136,104,48,12
End Dialog
Begin Dialog BUSFAULTDLG 48,46, 258, 59, "Specify fault"
  OptionGroup .FLTCONN
    OptionButton 12,16,24,12, "3PH"
    OptionButton 40,16,24,12, "2LG"
    OptionButton 68,16,24,12, "1LG"
    OptionButton 96,16,24,12, "L-L"
  Text 8,8,64,8, "Fault Connection"
  Text 128,8,64,8, "Fault impedance"
  Text 136,20,12,8, "Z="
  TextBox 148,16,40,12, .EditBox_1
  Text 192,20,12,8, "+ j"
  TextBox 204,16,40,12, .EditBox_2
  OKButton 72,40,48,12
  CancelButton 136,40,48,12
End Dialog
'
' Function:
'   Get Fault spec. inputs from user
'
Function FaultDialog( Fltconn() As Long, FltOption() As Double, nStyle As Long, _
      ByRef dR As Double, ByRef dX As Double ) As Long
'
' Dialog specifications (generated by dialog editor)
'
  If nStyle = 1 Then  ' Picked Bus
    Dim dlg As BUSFAULTDLG
    button = Dialog( dlg )
    If button = 0 Then ' Canceled
      FaultDialog = 0
      Exit Function
    End If
    Fltconn( 1 + dlg.FLTCONN ) = 1
    FltOption(1) = 1.0
    dR = Val(dlg.EditBox_1)
    dX = Val(dlg.EditBox_2)
  Else  ' Picked Relay group
    Dim dlg1 As FAULTDLG
    dlg1.EditBox_3 = 10  '10%
    button = Dialog( dlg1 )
    If button = 0 Then ' Canceled
      FaultDialog = 0
      Exit Function
    End If
    Fltconn( 1 + dlg1.FLTCONN ) = 1
    If dlg1.FLTOPT = 4 Or dlg1.FLTOPT = 5  Then 
      FltOption(1 + 2*dlg1.FLTOPT ) = Val(dlg1.EditBox_3)
      FltOption(13) = Val(dlg1.EditBox_4)
      FltOption(14) = Val(dlg1.EditBox_5)
    Else
      FltOption( 1 + 2*dlg1.FLTOPT ) = 1.0
    End If
    dR = Val(dlg1.EditBox_1)
    dX = Val(dlg1.EditBox_2)
  End If
  FaultDialog = 1
End Function
