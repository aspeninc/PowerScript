' ASPEN PowerScript sample program
'
' GETSWDATA.BAS
'
' Retrieve and display switch data.
'
' Demonstrate how to access network data from a PowerScript program
' 
' PowerScript functions called:
'   FindBusByName()
'   GetData()
'
Sub main()

   If GetEquipment( TC_PICKED, DevHandle& ) = 0 Then
     Print "Please select a switch"
     Exit Sub
   End If

   If 0 = GetData( DevHandle&, SW_dRating, OutputVal ) Then GoTo HasError
   Print "Rating = ", OutputVal

   If 0 = GetData( DevHandle&, SW_sName, OutputVal ) Then GoTo HasError
   Print "Name = ", OutputVal

   If 0 = GetData( DevHandle&, SW_nInService, nNum& ) Then GoTo HasError
   Print "InService = ", nNum&

   If 0 = GetData( DevHandle&, SW_nStatus, nNum& ) Then GoTo HasError
   Print "Status = ", nNum&

   Exit Sub
 HasError:
   Print "Error: ", ErrorString( )
End Sub