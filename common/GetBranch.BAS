Sub main
  hnd& = 0
  If 0 = GetEquipment(TC_PICKED,hnd) Or EquipmentType(hnd) <> TC_BUS Then
    Print "No bus is selected"
    Stop
  End If
  hndBr& = 0
  While 1 = GetBusEquipment( hnd, TC_BRANCH, hndBr )
    Print PrintObj1LPF( hndBr )
    Call GetData( hndBr, BR_nHandle, hndLine& )
    'Print PrintObj1LPF( hndLine )
    Call GetData( hndLine, LN_sID, sID$ )
    Print sID
  Wend
End Sub

