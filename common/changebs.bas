' ASPEN PowerScript sample program
'
' CHANGEBS.BAS
'
' Change bus data.
' Demonstrate how to access network database.
' 
' PowerScript functions called:
'   NextBusByName()
'   GetData()
'   SetData()
'   PostData()
'
Const IDOK = -1
Sub main()
Begin Dialog DIALOG_1 45,37, 171, 77, "Change Bus"
  OptionGroup .GROUP_1
    OptionButton 40,4,32,12, "Area"
    OptionButton 76,4,32,12, "Zone"
  Text 4,4,33,12, "Change"
  Text 4,20,117,12, "From (leave blank to change all):"
  TextBox 124,20,37,12, .EditBox_2
  Text 4,36,12,12, "to: "
  TextBox 20,36,32,12, .EditBox_1
  OKButton 40,56,60,12
  CancelButton 112,56,28,12
End Dialog

  Dim dlg As Dialog_1
  dlg.EditBox_1 = 0
  If IDOK <> Dialog( dlg ) Then Exit Sub
  ChangeWhat = dlg.Group_1
  OldValue& = Val( dlg.EditBox_2 )
  NewValue& = Val( dlg.EditBox_1 )
  Count   = 0
  BusHnd& = 0	'This will make next cmd to seek the first bus
  If ChangeWhat = 0 Then strTemp$ = "Change Area" Else strTemp$ = "Change Zone"
  If OldValue > 0 Then strTemp$ = strTemp$ & " from " & Str(OldValue)
  strTemp$ = strTemp$ & " to " & Str( NewValue )
  Call PrintTTY( strTemp )
  While NextBusByName( BusHnd& ) > 0
   If ChangeWhat = 0 Then
     If GetData( BusHnd&, BUS_nArea, ExistingValue& ) = 0 Then GoTo HasError
     If OldValue = 0 Or OldValue = ExistingValue Then
       If SetData( BusHnd&, BUS_nArea, NewValue ) = 0 Then GoTo HasError
       If PostData( BusHnd& ) > 0 Then Count = Count + 1
       Call PrintTTY(FullBusName(BusHnd))
     End If
   Else
     If GetData( BusHnd&, BUS_nZone, ExistingValue& ) = 0 Then GoTo HasError
     If OldValue = 0 Or OldValue = ExistingValue Then
       If SetData( BusHnd&, BUS_nZone, NewValue ) = 0 Then GoTo HasError
       If PostData( BusHnd& ) > 0 Then Count = Count + 1
     End If
   End If
  Wend
  If Count > 0 Then
    Print Count, " buses have been modified. Full list is in TTY"
  Else
    Print "No change made"
  End If
  Exit Sub
HasError:
  Print "Error: ", ErrorString( )
End Sub