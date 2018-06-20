' ASPEN PowerScrip sample program
'
' PREFAULT.BAS
'
' Simulate prefault condition in the network.
'
' Version 1.0
' Category: OneLiner
'
'
Sub main()
   ' Variable declaration
   Dim FltConn(4) As Long
   Dim FltOpt(14) As Double
   Dim ShowRelayFlag(4) As Long

   If GetEquipment( TC_PICKED, BusHnd& ) = 0 Then 
     Print "Please select a bus"
     Exit Sub
   End If
   If EquipmentType( BusHnd ) <> TC_BUS Then
     Print "Please select a bus"
     Exit Sub
   End If

   
   'fault connections
   FltConn(1) = 1	' Do 3PH
   FltConn(2) = 0   ' Do 2LG
   FltConn(3) = 0   ' Do 1LG
   FltConn(4) = 0   ' Do LL
   For ii = 1 to 14
     FltOpt(ii) = 0
   Next  
   FltOpt(1)  = 1   ' Bus fault no outage
   FltOpt(2)  = 0
   Rflt         = 9999 ' Fault R
   Xflt         = 9999 ' Fault X
   ClearPrev    = 0 ' Clear previous result
   
   If DoFault( BusHnd, FltConn, FltOpt, OutageOpt, OutageLst, _
                 Rflt, Xflt, ClearPrev ) = 0 _
           Then GoTo HasError
   For ii = 1 To 4
      ShowRelayFlag(ii) = 0
   Next
   Call ShowFault( SF_LAST, 99, 2, 0, ShowRelayFlag )
   Stop
   HasError:
   Print "Error: ", ErrorString( )
End Sub
