' ASPEN PowerScript sample program
'
' LOSES.BAS
'
' Compute total loses.
' 
'
Const IDOK = -1
Const MxLines = 1000
Begin Dialog Dialog_1 34,52, 148, 58, "Compute Loses in"
  OptionGroup .GROUP_1
    OptionButton 4,4,32,12, "Area"
    OptionButton 36,4,60,12, "Zone number ="
  OKButton 32,40,44,12
  CancelButton 84,40,28,12
  Text 20,20,28,12, "kV from:"
  TextBox 52,20,24,12, .KVFROM
  Text 84,20,12,12, "to:"
  TextBox 96,20,24,12, .KVTO
  TextBox 100,4,36,12, .AreaZone
End Dialog

Sub main()
  Dim dlg As Dialog_1
  Dim Qarray(4) As Double 
  Dim Parray(4) As Double 
  Dim LnMap(MxLines) As Long 

  dlg.AreaZone = 0
  dlg.KVFROM   = 0
  dlg.KVTO     = 9999
  If IDOK <> Dialog( dlg ) Then Exit Sub
  AreaFlag     = dlg.Group_1
  AreaZone&    = Val( dlg.AreaZone )
  kvFrom#      = Val( dlg.KVFROM )
  kvTo#        = Val( dlg.KVTO )
  'Initialization
  Loses#   = 0.0
  CountLn& = 0
  LnMap(1) = -1   'map to avoid double counting

  BusHnd&  = 0	'This will make next cmd to seek the first bus
  While NextBusByName( BusHnd& ) > 0
    'Check if bus is in area/zone
    If AreaFlag = 0 Then
      Call GetData( BusHnd&, BUS_nArea, AreaValue& )
    Else
      Call GetData( BusHnd&, BUS_nZone, AreaValue& )
    End If
    If AreaValue <> AreaZone Then GoTo Cont    ' Wrond area/zone
    'Check kV level
    Call GetData( BusHnd&, BUS_dKVNominal, BusKV# )
    If BusKV < kvFrom Or BusKV > kvTo Then GoTo cont  ' Not in kV range
    'Get branch loses and sum them up
    BrHnd& = 0    'This will make the next cmd to seek the first branch
    While GetBusEquipment( BusHnd&, TC_BRANCH, BrHnd& ) > 0
      Call GetData( BrHnd, BR_nHandle, LnHnd& )
      'Check if this line has been counted
      ii& = 1
      While LnMap(ii) > -1
        If LnMap(ii) = LnHnd Then GoTo ExitWhile
        ii = ii + 1
      Wend
      ExitWhile:
      If LnMap(ii) = -1 Then
        'Line has not been counted before
        LnMap(ii)   = LnHnd    'Update map
        LnMap(ii+1) = -1
        'Get line flow
        If GetFlow( LnHnd, Parray, Qarray ) = 0 Then GoTo HasError
        'Compute loses
        Loses# = Loses# + Abs( Parray(1)+Parray(2) )
        CountLn = CountLn + 1
      End If
    Wend  'Each branch
  Cont:
  Wend
  Print "Found ", CountLn, " branches. Loses = ", Loses#, "MW"
  Exit Sub
HasError:
  Print "Error: ", ErrorString( )
End Sub