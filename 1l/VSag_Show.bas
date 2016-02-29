' ASPEN PowerScript sample program
' VSag_Show.BAS
'
' Demonstrate the Voltage Save Analysis Command
'
' PowerScript functions used
Sub main()

  Dim vdMag(4) As double
  
  ' Get picked bus handle
  If GetEquipment( TC_PICKED, nBusHnd& ) = 0 Then
   Print "Must select a bus"
   Exit Sub
  End If

  If 0 = GetVSVoltage( nBusHnd,  vdMag ) Then GoTo HasError
  
  dSag# = -1.0
  For ii=1 to 4
    If dSag < vdMag(ii) then dSag = vdmag(ii)
  Next
  
  Print "Voltage sag on fault at: " + FullBusName(nBusHnd) + " = " + Str(dSag)
  
  Stop
  HasError:
  Print "Error: ", ErrorString( )
  Stop
End Sub