' ASPEN PowerScript sample program
'
' LINEZS.BAS
'
' Display line impedance in ohm
' 
'
' ===================== main() ================================================
'
Sub main()
  Dim dlg As PUDLG
  ' Get system MVA
  If GetData( HND_SYS, SY_dBaseMVA, BaseMVA ) = 0 Then BaseMVA = 100
  ' Figure out kV base from picked object
  If 0 <> GetEquipment( TC_PICKED, PickedHnd ) Then
    ' Probe to see what's being picked
    Select Case EquipmentType( PickedHnd )
      Case TC_LINE
        If 0 = GetData( PickedHnd, LN_dX, dX# ) Then GoTo HasError
        If 0 = GetData( PickedHnd, LN_dX0, dX0# ) Then GoTo HasError
        If 0 = GetData( PickedHnd, LN_dR, dR# ) Then GoTo HasError
        If 0 = GetData( PickedHnd, LN_dR0, dR0# ) Then GoTo HasError
        If 0 = GetData( PickedHnd, LN_nBus1Hnd, nBusHnd& ) Then GoTo HasError
        If 0 = GetData( nBusHnd&, BUS_dKVNorminal, BaseKV ) Then GoTo HasError
        Equipment$= "Line impedance: "
      Case Else
        Print "Please select a line"
        exit Sub
    End Select
  Else
    Print "Please select a line"
    exit Sub
  End If
  BaseZ = BaseKV * BaseKV / BaseMVA
  dX  = dX * BaseZ
  dX0 = dX0 * BaseZ
  dR  = dR * BaseZ
  dR0 = dR0* BaseZ
  aLine$ = Equipment$ & "Z1=" & Str(dR) & " +j " & Str(dX) & "  Z0= " & Str(dR0) & " +j " & Str(dX0)
  Print aLine
  PrintTTY( aLine )
Exit Sub
  ' Error handling
  HasError:
  Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()