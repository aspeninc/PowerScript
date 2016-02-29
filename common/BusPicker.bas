' ASPEN PowerScript Sample Program
'
' BusPicker.BAS
'
' Demonstrate how to use BusPicker Power Script fuction
'
' PowerScript functions called:
'

Sub main

 dim vnBusHnd1(30) As long  ' Important note: Bus handle list must have adequate dimension
 dim vnBusHnd2(30) As long  ' Important note: Bus handle list must have adequate dimension


 vnBusHnd1(1) = 0           'Important: must always terminate the list with 0
 vnBusHnd2(2) = 0           'Important: must always terminate the list with 0
 
 Do
  sWindowText$ = "My Bus Picker"
  nPicked& = BusPicker( sWindowText$, vnBusHnd1, vnBusHnd2 )
  nTemp&   = vnBusHnd2(1)
  strTemp$ = ""
  For ii& = 1 to nPicked
   nBsHnd& = vnBusHnd2(ii)
   strName$ = FullBusName( nBsHnd& )
   strTemp$ = strTemp$ + strName$ + Chr(10) + Chr(13)
  Next
  If nPicked > 0 then Print "Picked:", Chr(10), Chr(13), strTemp
 Loop While nPicked > 0

End Sub
