'ASPEN Sample Script

Sub main

DevHnd& = 0

While GetEquipment( TC_MU, DevHnd ) > 0
  Call GetData( DevHnd, MU_dX,        dX# )
  Call GetData( DevHnd, MU_dR,        dR# )
  'Consider only pair with none-zero impedance
  If Abs(dX)+Abs(dR) > 0.00001 Then     
    Call GetData( DevHnd, MU_nOrient1,  nOrient1& )
    Call GetData( DevHnd, MU_nOrient2,  nOrient2& )
    Call GetData( DevHnd, MU_dFrom1,    dFrom1# )
    Call GetData( DevHnd, MU_dFrom2,    dFrom2# )
    Call GetData( DevHnd, MU_dTo1,      dTo1# )
    Call GetData( DevHnd, MU_dTo2,      dTo2# )
    Call GetData( DevHnd, MU_nHndLine1, nHndLine1& )
    Call GetData( DevHnd, MU_nHndLine2, nHndLine2& )
    Print printLineID(nHndLine1, nOrient1) & Str(dFrom1) & "-" & Str(dTo1) & "%" & Chr(13) & Chr(10) _
          & printLineID(nHndLine2, nOrient2) & Str(dFrom2) & "-" & Str(dTo2) & "%" & Chr(13) & Chr(10) _
          & Str(dR) & " + j" & Str(dX)
  End If
Wend
End Sub

Function printLineID( nHnd&, nOrient& ) As String
  Call GetData( nHnd, LN_nBus1Hnd, nBus1Hnd& )
  Call GetData( nHnd, LN_nBus2Hnd, nBus2Hnd& )
  Call GetData( nHnd, LN_sID,      sID$ )
  If nOrient = 1 Then
    printLineID = FullBusName(nBus1Hnd) & "-" & FullBusName(nBus2Hnd) & " " & sID
  Else
    printLineID = FullBusName(nBus2Hnd) & "-" & FullBusName(nBus1Hnd) & " " & sID
  End If
End Function
