' ASPEN PowerScript sample program
'
' FLTPARAM.BAS
'
' Demonstrate how to get fault simulation input parameters
'
' Version 1.0
' Category: OneLiner

Sub main
  If pickFault(1) = 0 Then GoTo HasError
  
  If GetData(HND_SC,FT_dRPt,dR#) = 0 Then GoTo HasError
  If GetData(HND_SC,FT_dXPt,dX#) = 0 Then GoTo HasError
  Print "ZP =", dR, "+j", dX

  If GetData(HND_SC,FT_dRNt,dR#) = 0 Then GoTo HasError
  If GetData(HND_SC,FT_dXNt,dX#) = 0 Then GoTo HasError
  Print "ZN =", dR, "+j", dX

  If GetData(HND_SC,FT_dRZt,dR#) = 0 Then GoTo HasError
  If GetData(HND_SC,FT_dXZt,dX#) = 0 Then GoTo HasError
  Print "ZZ =", dR, "+j", dX

  If GetData(HND_SC,FT_dXR,dXR#) = 0 Then GoTo HasError
  Print "X/R =", dXR

  If GetData(HND_SC,FT_dMVA,dMVA#) = 0 Then GoTo HasError
  Print "MVA =", dMVA

  exit Sub
  HasError:
  Print "Error: ", ErrorString( )
  Close 
End Sub
