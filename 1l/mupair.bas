' ASPEN PowerScript Sample Program
'
'MUPAIR.BAS
'
' Retrieve mutual pair data
'
' Version 1.0
' Category: OneLiner
'
Sub main

 nPicked& = 0

 Call getequipment( TC_PICKED, nPicked& )

 If nPicked = 0 Or equipmentType( nPicked ) <> TC_LINE Then
  Print "Must select a line"
  Stop
 End If

 dim dX(10) As double
 dim dR(10) As double
 dim dFrom1(10) As double
 dim dFrom2(10) As double
 dim dTo1(10) As double
 dim dTo2(10) As double
 pairHnd& = 0
 While getdata( nPicked, LN_nMuPairHnd, pairHnd ) > 0
  If 0 = getdata( pairHnd, MU_vdX, dX ) Then GoTo hasError
  If 0 = getdata( pairHnd, MU_vdR, dR ) Then GoTo hasError
  ' Only pair(s) with non-zero X and R are usable.
  If dX(1) <> 0.0 Or dR(1) <> 0.0 Then
   If 0 = getdata( pairHnd, MU_vdFrom1, dFrom1 ) Then GoTo hasError
   If 0 = getdata( pairHnd, MU_vdFrom2, dFrom2 ) Then GoTo hasError
   If 0 = getdata( pairHnd, MU_vdTo1, dTo1 ) Then GoTo hasError
   If 0 = getdata( pairHnd, MU_vdTo2, dTo2 ) Then GoTo hasError
   Call getdata( pairHnd, MU_nHndLine2, hLine2& )
   If hLine2 = nPicked Then Call getdata( pairHnd, MU_nHndLine1, hLine2& )
   Print lineStr( hLine2 ) & Chr(13) & Chr(10) & _
         "From1=" & Format(dFrom1(1), "#0.00") & Format(dFrom1(2), " #0.00") & Format(dFrom1(3), " #0.00") & Format(dFrom1(4), " #0.00") & Format(dFrom1(5), " #0.00") & Chr(13) & Chr(10) & _
         "To1=" & Format(dTo1(1), "#0.00") & Format(dTo1(2), " #0.00") & Format(dTo1(3), " #0.00") & Format(dTo1(4), " #0.00") & Format(dTo1(5), " #0.00") & Chr(13) & Chr(10) & _
         "From2=" & Format(dFrom2(1), "#0.00") & Format(dFrom2(2), " #0.00") & Format(dFrom2(3), " #0.00") & Format(dFrom2(4), " #0.00") & Format(dFrom2(5), " #0.00") & Chr(13) & Chr(10) & _
         "To2=" & Format(dTo2(1), "#0.00") & Format(dTo2(2), " #0.00") & Format(dTo2(3), " #0.00") & Format(dTo2(4), " #0.00") & Format(dTo2(5), " #0.00") & Chr(13) & Chr(10) & _
         "R=" & Format(dR(1), "#0.000") & Format(dR(2), " #0.000") & Format(dR(3), " #0.000") & Format(dR(4), " #0.000") & Format(dR(5), " #0.000") & Chr(13) & Chr(10) & _
         "X=" & Format(dX(1), "#0.000") & Format(dX(2), " #0.000") & Format(dX(3), " #0.000") & Format(dX(4), " #0.000") & Format(dX(5), " #0.000") & Chr(13) & Chr(10) 
   For ii = 1 to 5
     dX(ii) = dX(ii)*2
     dR(ii) = dR(ii)*2
     dFrom1(ii) = dFrom1(ii) + 2
     dFrom2(ii) = dFrom2(ii) + 2
     dTo1(ii) = dTo1(ii) + 2
     If dTo1(ii) > 100 Then dTo1(ii) = 100
     dTo2(ii) = dTo2(ii) + 2
     If dTo2(ii) > 100 Then dTo2(ii) = 100
   Next
   If 0 = SetData(pairHnd, MU_vdFrom1, dFrom1 ) Then GoTo hasError
   If 0 = SetData(pairHnd, MU_vdFrom2, dFrom2 ) Then GoTo hasError
   If 0 = setdata( pairHnd, MU_vdTo1, dTo1 ) Then GoTo hasError
   If 0 = setdata( pairHnd, MU_vdTo2, dTo2 ) Then GoTo hasError
   If 0 = setdata( pairHnd, MU_vdX, dX ) Then GoTo hasError
   If 0 = setdata( pairHnd, MU_vdR, dR ) Then GoTo hasError
   If 0 = postdata( pairHnd ) Then GoTo hasError
   If 0 = getdata( pairHnd, MU_vdX, dX ) Then GoTo hasError
   If 0 = getdata( pairHnd, MU_vdR, dR ) Then GoTo hasError
   If 0 = getdata( pairHnd, MU_vdFrom1, dFrom1 ) Then GoTo hasError
   If 0 = getdata( pairHnd, MU_vdFrom2, dFrom2 ) Then GoTo hasError
   If 0 = getdata( pairHnd, MU_vdTo1, dTo1 ) Then GoTo hasError
   If 0 = getdata( pairHnd, MU_vdTo2, dTo2 ) Then GoTo hasError
   Print lineStr( hLine2 ) & Chr(13) & Chr(10) & _
         "From1=" & Format(dFrom1(1), "#0.00") & Format(dFrom1(2), " #0.00") & Format(dFrom1(3), " #0.00") & Format(dFrom1(4), " #0.00") & Format(dFrom1(5), " #0.00") & Chr(13) & Chr(10) & _
         "To1=" & Format(dTo1(1), "#0.00") & Format(dTo1(2), " #0.00") & Format(dTo1(3), " #0.00") & Format(dTo1(4), " #0.00") & Format(dTo1(5), " #0.00") & Chr(13) & Chr(10) & _
         "From2=" & Format(dFrom2(1), "#0.00") & Format(dFrom2(2), " #0.00") & Format(dFrom2(3), " #0.00") & Format(dFrom2(4), " #0.00") & Format(dFrom2(5), " #0.00") & Chr(13) & Chr(10) & _
         "To2=" & Format(dTo2(1), "#0.00") & Format(dTo2(2), " #0.00") & Format(dTo2(3), " #0.00") & Format(dTo2(4), " #0.00") & Format(dTo2(5), " #0.00") & Chr(13) & Chr(10) & _
         "R=" & Format(dR(1), "#0.000") & Format(dR(2), " #0.000") & Format(dR(3), " #0.000") & Format(dR(4), " #0.000") & Format(dR(5), " #0.000") & Chr(13) & Chr(10) & _
         "X=" & Format(dX(1), "#0.000") & Format(dX(2), " #0.000") & Format(dX(3), " #0.000") & Format(dX(4), " #0.000") & Format(dX(5), " #0.000") & Chr(13) & Chr(10) 

  End If
 Wend
Stop
hasError:
 Print ErrorString()
End Sub

Function lineStr( ByVal nHnd& ) As String
 Call getdata( nHnd, LN_nBus1Hnd, nBus1& )
 Call getdata( nHnd, LN_nBus2Hnd, nBus1& )
 Call getdata( nHnd, LN_sID, sID$ )
 sStr$ = fullbusname(nBus1) + "-" + fullbusname(nBus2) + " " + sID
 lineStr = sStr
End Function
