' ASPEN PowerScript sample program
'
' SCDATA.BAS
'
' Modify and display series capacitor/reactor data.
'
' Demonstrate how to access network data from a PowerScript program
' 
' PowerScript functions called:
'   FindBusByName()
'   GetData()
'
Sub main()

   If GetEquipment( TC_PICKED, DevHandle& ) = 0 Then
     Print "Please select a series capacitor/reactor"
     Exit Sub
   End If
   If EquipmentType( DevHandle& ) <> TC_SCAP Then
     Print "Please select a series capacitor/reactor"
     Exit Sub
   End If 
  
   ' Modify data
   sFlag$ = InputBox("Enter flag", "Flag" )
   If sFlag$ = "" Then 
     Print "Bye"
     Stop
   End If
   nFlag& = Val(sFlag$)
'   If 0 = SetData( DevHandle&, SC_nSComp, nFlag ) Then GoTo HasError
   If 0 = SetData( DevHandle&, SC_nInService, nFlag ) Then GoTo HasError

   If 0 = PostData( DevHandle& ) Then GoTo HasError

   Print "Done"
   
   Exit Sub
 HasError:
   Print "Error: ", ErrorString( )
End Sub