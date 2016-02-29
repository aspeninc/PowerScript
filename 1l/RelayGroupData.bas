' ASPEN PowerScript sample program
'
' RelayGroupData.BAS
'
' Demo of relay group data access from a PowerScript program
' 
'
Sub main()

  ' Get picked andle
  If (GetEquipment( TC_PICKED, DevHandle& ) = 0) Or _
     (EquipmentType( DevHandle& ) <> TC_RLYGROUP ) Then
   Print "Please select a relay group"
   Exit Sub
  End If


   If 0 = GetData( DevHandle&, RG_sNote, sMemo$ ) Then GoTo HasError
   Print "Memo = ", sMemo$

   If 0 = GetData( DevHandle, RG_dBreakerTime, dBkrTime# ) Then GoTo HasError
   Print "dBkrTime#= ", dBkrTime#

   sMemo$ = sMemo$ & " + More text"
   If 0 = SetData( DevHandle&, RG_sNote, sMemo$ ) Then GoTo HasError
   dBkrTime# = dBkrTime# * 2
   If 0 = SetData( DevHandle, RG_dBreakerTime, dBkrTime# ) Then GoTo HasError
   If 0 = PostData( DevHandle ) Then GoTo HasError

   If 0 = GetData( DevHandle&, RG_sNote, sMemo$ ) Then GoTo HasError
   Print "Memo = ", sMemo$

   If 0 = GetData( DevHandle, RG_dBreakerTime, dBkrTime# ) Then GoTo HasError
   Print "dBkrTime#= ", dBkrTime#


   Exit Sub
 HasError:
   Print "Error: ", ErrorString( )
End Sub