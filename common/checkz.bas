Sub main
smallZ = 9999
countP = 0
countZ = 0
countL = 0
LnHnd& = 0
While GetEquipment( TC_LINE, LnHnd ) > 0
  Call GetData( LnHnd, LN_dR, dR# )
  Call GetData( LnHnd, LN_dX, dX# )
  Call GetData( LnHnd, LN_dR0, dR0# )
  Call GetData( LnHnd, LN_dX0, dX0# )
  If dR=0.0 And dX < smallZ Then countP = countP + 1
  If dR0=0.0 And dX0 < smallZ Then countZ = countZ + 1
  countL = countL + 1
Wend
Print "Lines count = ", countL, "; countP = ", countP, ";  countZ = ", countZ
countP = 0
countZ = 0
countL = 0
LnHnd& = 0
While GetEquipment( TC_XFMR, LnHnd ) > 0
  Call GetData( LnHnd, XR_dR, dR# )
  Call GetData( LnHnd, XR_dX, dX# )
  Call GetData( LnHnd, XR_dR0, dR0# )
  Call GetData( LnHnd, XR_dX0, dX0# )
  If dR=0.0 And dX < smallZ Then countP = countP + 1
  If dR0=0.0 And dX0 < smallZ Then countZ = countZ + 1
  countL = countL + 1
Wend
Print "Xfmr count = ", countL, "; countP = ", countP, ";  countZ = ", countZ
End Sub