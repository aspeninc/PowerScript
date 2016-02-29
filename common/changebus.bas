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
Begin Dialog DIALOG_1 18,35, 116, 76, "Change Bus"
  OptionGroup .GROUP_1
    OptionButton 44,4,32,12, "Area"
    OptionButton 80,4,32,12, "Zone"
  Text 4,4,41,12, "Change all"
  Text 20,20,21,12, "from:"
  TextBox 44,20,32,12, .EditBox_2
  Text 28,36,12,12, "to: "
  TextBox 44,36,32,12, .EditBox_1
  OKButton 20,56,40,12
  CancelButton 64,56,28,12
End Dialog


  Dim dlg As Dialog_1
  dlg.EditBox_1 = 0
  If IDOK <> Dialog( dlg ) Then Exit Sub
  ChangeWhat = dlg.Group_1
  NewValue& = Val( dlg.EditBox_1 )
  OldValue& = Val( dlg.EditBox_2 )
  Count   = 0
  BusHnd& = 0	'This will make next cmd to seek the first bus
  While NextBusByName( BusHnd& ) > 0
   If ChangeWhat = 0 Then nParamCode& = BUS_nArea Else nParamCode = BUS_nZone
   Call GetData( BusHnd&, nParamCode, ExistingValue& )
   If ExistingValue = OldValue Then
    If SetData(   BusHnd&, nParamCode, NewValue ) = 0 Then GoTo HasError
    If PostData( BusHnd& ) > 0 Then Count = Count + 1
   End If
  Wend
  Print Count, " buses have been modified"
  Exit Sub
HasError:
  Print "Error: ", ErrorString( )
End Sub