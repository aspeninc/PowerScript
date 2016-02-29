' ASPEN PowerScript sample program
'
' GETDATA.BAS
'
' Retrieve and display bus data.
'
' Demonstrate how to access network data from a PowerScript program
' 
' PowerScript functions called:
'   FindBusByName()
'   GetData()
'
Sub main()

   If 0 = FindBusByName( "Reusens", 132, DevHandle ) Then GoTo HasError
   Print "Bus handle found: ", DevHandle

GoTo jump

   
   If 0 = GetData( DevHandle&, BUS_dKVnorminal, OutputVal ) Then GoTo HasError
   Print "fKVnorminal = ", OutputVal

   If 0 = GetData( DevHandle, BUS_nNumber, OutputVal ) Then GoTo HasError
   Print "nNumber= ", OutputVal

   If 0 = GetData( DevHandle, BUS_nArea, OutputVal ) Then GoTo HasError
   Print "nArea= ", OutputVal 

   If 0 = GetData( DevHandle, BUS_nTapBus, OutputVal ) Then GoTo HasError
   Print "nTapBus= ", OutputVal 
jump:
   If 0 = GetData( DevHandle, BUS_sLocation, OutputStr) Then GoTo HasError
   Print "sLocation =", OutputStr

   If 0 = GetData( DevHandle, BUS_sName, OutputStr ) Then GoTo HasError
   Print "sName= " & OutputStr

   Exit Sub
 HasError:
   Print "Error: ", ErrorString( )
End Sub