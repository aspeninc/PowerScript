' ASPEN PowerScript Sample Program
'
' LINEOUT.BAS
'
' Simulate fault at selected bus with and without branch outage
'
' Demonstrate how to simulate fault from a PowerScript program
'
' PowerScript functions called:
'   GetEquipment()
'   FullBusName()
'   DoFault
'   ShowFault()
'   GetSCVoltage()
'   GetSCCurrent()
'
Sub main()
   ' Variable declaration
   Dim MagArray(16) As Double
   Dim AngArray(16) As Double
   Dim FltConn(4) As Long
   Dim FltOption(14) As Double
   Dim OutageList(20) As Long
   Dim OutageType(3) As Long
   Dim ShowRelayFlag(4) As Long

   If GetEquipment( TC_PICKED, BusHnd ) = 0 Then 
     Print "Must select a bus"
     Exit Sub
   End If
   If EquipmentType( BusHnd ) <> TC_BUS Then
     Print "Must select a bus"
     Exit Sub
   End If

   ' Prepare file for output
   FileName = "output.rep"
   Open FileName For Output As #1

   ' Initialize the arrays
   For ii = 1 To 4 
     FltConn(ii) = 0
   Next 
   For ii = 1 To 12
     FltOption(ii) = 0.0
   Next
   For ii = 1 To 3
     OutageType(ii) = 0
   Next
   For ii = 1 To 4
     ShowRelayFlag(ii) = 0
   Next
   Rflt       = 0.0   '
   Xflt       = 0.0
   nClearPrev = 1 ' Don't keep previous result

   FltConn(1)    = 1	' 3PH 
   FltConn(3)    = 1	' 1LG
   FltOption(1)  = 1   ' Bus fault
   FltOption(2)  = 1   ' Bus fault with outage
   OutageType(1) = 1	' Outage one at a time

   ' Prepare the outage list
   BranchHnd = 0
   For ii = 1 To 10  ' max 10 outage
     If GetBusEquipment( BusHnd, TC_BRANCH, BranchHnd ) > 0 Then
       OutageList(ii) = BranchHnd
     Else
       OutageList(ii) = 0	' Must always close the list
       Exit For
     End If
   Next

   ' Simulate the faults
   If 0 = DoFault( BusHnd, FltConn, FltOption, OutageType, OutageList, Rflt, _
                   Xflt, nClearPrev ) Then GoTo HasError

   ' Print output
   BusID = FullBusName( BusHnd )
   Print #1, "Fault simulation at Bus: ", BusID
   Print #1, ""
   Print #1, "                                  Phase A      Phase B      Phase C"
   Print #1, ""

   ' Must alway pick a fault before getting V and I
   If PickFault( 1 ) = 0 Then GoTo HasError
   Do
     If GetSCCurrent( HND_SC, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
     Print #1, _
        FaultDescription(); Chr(13); _
        "                                     "; _
        Format( MagArray(1), "####0.0"); "@"; Format( AngArray(1), "#0.0"), Space(5), _
        Format( MagArray(2), "####0.0"); "@"; Format( AngArray(2), "#0.0"), Space(5), _
        Format( MagArray(3), "####0.0"); "@"; Format( AngArray(3), "#0.0")
   Loop While PickFault( SF_NEXT ) > 0
   Print "Simulation complete. Report is in " & FileName
   Close
   Exit Sub
HasError:
   Print "Error: ", ErrorString( )
   Close
End Sub