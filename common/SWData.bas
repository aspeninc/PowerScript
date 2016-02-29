' ASPEN PowerScript sample program
'
' SWDATA.BAS
'
' Modify and display switch data.
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

   ' Modify switch data
   
   NewName$ = "NewName"
   If 0 = SetData( DevHandle&, SW_sName, NewName$ ) Then GoTo HasError

   NewRating# = "999.9"
   If 0 = SetData( DevHandle&, SW_dRAting, NewRating# ) Then GoTo HasError
 
   NewFlag& = 0
   If 0 = SetData( DevHandle&, SW_nInService, NewFlag& ) Then GoTo HasError

   NewStatus& = 0
   If 0 = SetData( DevHandle&, SW_nStatus, NewStatus& ) Then GoTo HasError

   If 0 = PostData( DevHandle& ) Then GoTo HasError

   ' Get switch data
   If 0 = GetData( DevHandle&, SW_nBus1Hnd, nBus1& ) Then GoTo HasError
   Print "Bus 1: ", FullBusName( nBus1& )

   If 0 = GetData( DevHandle&, SW_nBus2Hnd, nBus2& ) Then GoTo HasError
   Print "Bus 2: ", FullBusName( nBus2& )

   If 0 = GetData( DevHandle&, SW_sName, NewName$ ) Then GoTo HasError
   Print "Name = ", NewName$

   If 0 = SetData( DevHandle&, SW_dRating, NewRating# ) Then GoTo HasError
   Print "Rating = ", NewRating#

   If 0 = GetData( DevHandle&, SW_nInService, NewFlag& ) Then GoTo HasError
   If NewFlag& = 1 Then
     Print "Flag = In Service"
   Else
     Print "Flag = Out-of-service"
   End If

   If 0 = GetData( DevHandle&, SW_nStatus, NewStatus& ) Then GoTo HasError
   If NewStatus& = 1 Then
     Print "Status = Close"
   Else
     Print "Status = Open"
   End If

   Exit Sub
 HasError:
   Print "Error: ", ErrorString( )
End Sub