' ASPEN PowerScript Sample Program
'
' PROCZ.BAS
'
' Make X-only and R-only network
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
Dim XRfactor As Double
Dim SmallX  As Double
'

Begin Dialog Dialog_1 26,37, 171, 89, "Process Network"
  OptionGroup .NNETWORK
    OptionButton 12,8,88,12, "Make R-only network"
    OptionButton 12,36,88,12, "Make X-only network"
  Text 28,24,84,8, "When R is zero, set R = "
  TextBox 112,20,28,12, .EditBox_1
  Text 28,52,84,8, "When X is zero, set X = "
  TextBox 112,48,40,12, .EditBox_2
  Text 144,24,16,8, "* X"
  OKButton 36,68,44,12
  CancelButton 88,68,44,12
End Dialog

Sub main()

  Dim dlg As Dialog_1

  XRfactor = 0.03      ' if R=0, R = XRfactor * X
  SmallX   = 0.0001    ' if X=0, X = SmallX
  dlg.EditBox_1 = XRfactor
  dlg.EditBox_2 = SmallX

  If 0 = Dialog( dlg ) Then 
    Exit Sub
  End If

  XRfactor = Val(dlg.EditBox_1)
  XmallX   = Val(dlg.EditBox_2)

  ' Make R or X only network
  Call ProcessNetwork( dlg.nNetwork + 1 )

  If dlg.nNetwork = 0 Then
    Print "Network is now PURELY RESISTIVE." +Chr(10) + Chr(13) + _
            "Save file under new name to avoid losing original data"
  Else
    Print "Network is now PURELY REACTIVE." +Chr(10) + Chr(13) + _
            "Save file under new name to avoid losing original data"
  End If

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
 R1 = ArrayR(1)
 R2 = ArrayR(2)
 R3 = ArrayR(3)
 R4 = ArrayR(4)
 R5 = ArrayR(5)
 X1 = ArrayX(1)
 X2 = ArrayX(2)
 X3 = ArrayX(3)
 X4 = ArrayX(4)
 X5 = ArrayX(5)
 Open "c:\0tmp\ansiflt.log" For Append As 1
 Print #1, FullBusName( BusHandle ), ":", DevHandle, ",",R1, ",", R2, ",", R3, ",", R4, _
     ",", R5, ",", X1, ",", X2, ",", X3, ",", X4, ",", X5
 Close 1

         Call ProcessZ( ArrayX, ArrayR, 5, NetworkType )

 R1 = ArrayR(1)
 R2 = ArrayR(2)
 R3 = ArrayR(3)
 R4 = ArrayR(4)
 R5 = ArrayR(5)
 X1 = ArrayX(1)
 X2 = ArrayX(2)
 X3 = ArrayX(3)
 X4 = ArrayX(4)
 X5 = ArrayX(5)
' Open "c:\0tmp\ansiflt.log" For Append As 1
' Print #1, FullBusName( BusHandle ), ":", DevHandle, ",",R1, ",", R2, ",", R3, ",", R4, _
'    ",", R5, ",", X1, ",", X2, ",", X3, ",", X4, ",", X5
' Close 1

         If SetData( DevHandle&, GU_vdR, ArrayR() ) = 0 Then GoTo HasError
         If SetData( DevHandle&, GU_vdX, ArrayX() ) = 0 Then GoTo HasError
         If PostData( DevHandle& ) = 0 Then GoTo HasError

 R1 = ArrayR(1)
 R2 = ArrayR(2)
 R3 = ArrayR(3)
 R4 = ArrayR(4)
 R5 = ArrayR(5)
 X1 = ArrayX(1)
 X2 = ArrayX(2)
 X3 = ArrayX(3)
 X4 = ArrayX(4)
 X5 = ArrayX(5)
 'Open "c:\0tmp\ansiflt.log" For Append As 1
 'Print #1, FullBusName( BusHandle ), ":", DevHandle, ",",R1, ",", R2, ",", R3, ",", R4, _
 '   ",", R5, ",", X1, ",", X2, ",", X3, ",", X4, ",", X5
 'Close 1

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
      ' Process R
      If ArrayR(ii) = 0 Then 
        ArrayR(ii) = XRfactor * ArrayX(ii)
      End If
      ' Process X
      ArrayX(ii) = 0
    ElseIf NetworkType = 2 Then	' X only
      ' Process X
      If ArrayX(ii) = 0 Then
        ArrayX(ii) = SmallX
      End If
      ' Process R
      ArrayR(ii) = 0
    End If
  Next
End Sub

Sub ProcessY( vdValB() As Double, vdValG() As Double, _
              ByVal ArraySize&, ByVal NetworkType& )
' Process all admitance in the array
  For ii = 1 To ArraySize
    ArrayIndex = ArrayIndex + 1
    If NetworkType = 1 Then	' G only
      ' Process B
      vdValB(ii) = 0
    ElseIf NetworkType = 2 Then	' X only
      ' Process G
      vdValG(ii) = 0
    End If
  Next
End Sub

