Sub main

 dim vnList1(6)
 dim vnList2(6)
 
 nExcludeTap = InputBox( "Exclude tap buses? (1-yes;0-no)", "Tap buses", 0 )
 Call BusPicker( "Text", vnList1, vnList2, nExcludeTap )
End Sub
