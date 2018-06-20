' ASPEN PowerScript example
' VOLTAGESAG.BAS
'
' Print post fault voltage at every bus to a csv file
'
' Demonstrate how PowerScript program can access Shortcircuit
' simulation result
'
' Version 1.0
' Category: OneLiner
'
' PowerScript functions called:
'   NextBusByName()
'   FullBusName()
'   PickFault()
'   GetSCVoltage()
'
Sub main()
  Dim MagArray(12) As Double, AngArray(12) As Double
  Dim DummyArray(4) As Long

  ' Must always show fault before getting V and I
  If PickFault( 1 ) = 0 Then 
    Print "No fault result available"
    exit Sub
  End If

  ' Prepare output file
  Open "c:\000tmp\voltagesag.csv" For Output As 1
   Print #1, "Bus,Va,,Vb,,Vc," 

   ' Loop through all buses in the network
   BusHnd& = 0
   While NextBusByName( BusHnd ) > 0
      ' Get Bus info
      Bus1ID = FullBusName( BusHnd )

      ' Get voltagge at the end bus
      If GetSCVoltage( BusHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
      ' Print it
      Print #1, _
            Bus1ID & "," & _
            Format( MagArray(1), "#0.0") & "," & Format( AngArray(1), "#0.0") & "," & _
            Format( MagArray(2), "#0.0") & "," & Format( AngArray(2), "#0.0") & "," & _
            Format( MagArray(3), "#0.0") & "," & Format( AngArray(3), "#0.0")

   Wend  ' Each Bus

   ' Clean up
   Close 1
   Print "Output written to c:\voltagesag.csv"
   Exit Sub
HasError:
  Print "Error: ", ErrorString( )
  Close 
End Sub
