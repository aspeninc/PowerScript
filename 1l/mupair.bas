' ASPEN PowerScript Sample Program
'
'MUPAIR.BAS
'
' Retrieve mutual pair data
'
'
Sub main

 nPicked& = 0

 Call getequipment( TC_PICKED, nPicked& )

 If nPicked = 0 Or equipmentType( nPicked ) <> TC_LINE Then
  Print "Must select a line"
 End If

 pairHnd& = 0

 While getdata( nPicked, LN_nMuPairHnd, pairHnd ) > 0
  Call getdata( pairHnd, MU_dX, dX# )
  Call getdata( pairHnd, MU_dR, dR# )
  ' Only pair(s) with non-zero X and R are usable.
  If dX <> 0.0 Or dR <> 0.0 Then
   Call getdata( pairHnd, MU_nHndLine2, hLine2& )
   If hLine2 = nPicked Then Call getdata( pairHnd, MU_nHndLine1, hLine2& )
   Print lineStr( hLine2 )
  End If
 Wend

End Sub

Function lineStr( ByVal nHnd& ) As String
 Call getdata( nHnd, LN_nBus1Hnd, nBus1& )
 Call getdata( nHnd, LN_nBus2Hnd, nBus1& )
 Call getdata( nHnd, LN_sID, sID$ )
 sStr$ = fullbusname(nBus1) + "-" + fullbusname(nBus2) + " " + sID
 lineStr = sStr
End Function
