' ASPEN PowerScript sample program
'
' WINDFARM.BAS
'
' Simulate fault on a relay group or a bus near large wind farm sources.
'
' Fault simulation is carried out in two iterations. First simulation 
' is done with all wind farm generators taken off-line. 
' Generators at buses with voltage sag not exceeding 
' dThreshold are turned on in the second iteration to model contribution 
' from the wind farm
'
' List of all wind farm generators must be stored in a comma
' delimited text file with path name GenFile$.
'
' Last modified: 6/15/2004
'
' Version 1.0
' Category: OneLiner
'
' Program options:
Const dThreshold = 0.7
Const GenFile$   = "c:\generator.csv" 'Generator list
Const TempFile$  = "c:\wndfarm.res"   'Temporary file
Const MX_GEN     = 50                 'Max number of of wind farm generator
'
' Program constants
Const FlagActive = 1
Const FlagInActive = 2
'
' Main program code
Sub main()
 ' Variable declaration
 Dim FltConnection(4) As Long
 Dim FltOption(14) As Double
 Dim OutageType(3) As Long
 Dim OutageList(15) As Long
 Dim ShowRelayFlag(4) As Long
 Dim Rflt As Double, Xflt As Double
 Dim vdMag(4) As Double, vdAng(4) As Double 
 Dim BsHndList(MX_GEN) As Long

 If GetEquipment( TC_PICKED, PickedHandle ) = 0 Then
   Print "You must select a relay group or a bus before running this program."
   Exit Sub
 End If

 ' Check the TempFile. If one found, restore out-of-service generators listed.
 On Error GoTo Skip1
 Open TempFile For Input As 1
 Do While Not EOF(1)
  Line Input #1, TextLine
  'Extract bus name and kV
  Pos2 = InStr( 2, TextLine, "kV" )
  Pos1 = Pos2 - 1
  Do While " " <> Mid( TextLine, Pos1, 1 )
   Pos1 = Pos1 - 1
  Loop
  sKV$ = Trim(Mid( TextLine, Pos1, Pos2-Pos1 ))
  KV# = Val(sKV)
  BName$ = Trim(Mid( TextLine, 1, Pos1-1 ))
  'Find bus handle
  If 0 <> FindBusByName( bName, KV, nBusHnd& ) Then
   'Restore in-service all generators units at the bus
   nGenUnitHnd& = 0
   Do While 1 = GetBusEquipment( nBusHnd, TC_GENUNIT, nGenUnitHnd )
    If 1 <> SetData( nGenUnitHnd, GU_nOnLine, FlagActive ) Then GoTo haserror
    PostData(nGenUnitHnd)
   Loop
  End If
 Loop
 Close #1
 Kill( TempFile )

 Skip1:
 ' Initialize dofault parameters
 For ii = 1 To 4 
  FltConnection(ii) = 0
 Next 
 For ii = 1 To 12
  FltOption(ii) = 0.0
 Next
 OutageList(1) = -1
 For ii = 1 To 3
  OutageType(ii) = 0
 Next
 For ii = 1 To 4
   ShowRelayFlag(ii) = 0
 Next
 Rflt        = 0.0   ' No fault impedance
 Xflt        = 0.0
 ClearPrev   = 1     ' Don't keep previous result

 ' Determine what's being picked
 TypeCode = EquipmentType( PickedHandle )
 If TypeCode = TC_RLYGROUP Then
   ' Relay group fault dialog
   If FaultDialog( FltConnection, FltOption, 0, Rflt, Xflt ) = 0 Then Stop
 ElseIf TypeCode = TC_BUS Then
   ' Bus fault dialog
   If FaultDialog( FltConnection, FltOption, 1, Rflt, Xflt ) = 0 Then Stop
 Else
   Print "You must select a relay group or a bus before running this program."
   Exit Sub
 End If

 bFirstTime = 1
 'Read file with generator list. Take them off-line
 nGenCount = 0
 Open GenFile For Input As #1
 Do While Not EOF(1)
  Line Input #1, TextLine
  'Extract bus name and kV
  Pos2 = InStr( 2, TextLine, "kV" )
  Pos1 = Pos2 - 1
  Do While " " <> Mid( TextLine, Pos1, 1 )
    Pos1 = Pos1 - 1
  Loop
  sKV$ = Trim(Mid( TextLine, Pos1, Pos2-Pos1 ))
  KV# = Val(sKV)
  BName$ = Trim(Mid( TextLine, 1, Pos1-1 ))
  'Find bus handle
  If 0 <> FindBusByName( bName, KV, nBusHnd& ) Then
   'Find generator handle
   nGenHnd& = 0
   If  1=GetData( nBusHnd, GE_nActive, nFlag& ) Then
    If nFlag = 1 Then
     If bFirstTime=1 Then
      'Open temporary file to save list of outaged generators
      On Error GoTo Skip2
      Open TempFile For Output As #2
      bFirstTime = 0
     End If
     nGenUnitHnd& = 0
     Do While 1 = GetBusEquipment( nBusHnd, TC_GENUNIT, nGenUnitHnd )
      If 1 <> SetData( nGenUnitHnd, GU_nOnLine, FlagInActive ) Then GoTo HasError
      PostData( nGenUnitHnd)
     Loop
     sLine$ = BName+" "+sKV+"kV"
     Print #2, sLine ' Keep list of outaged generators in the TempFile
     nGenCount = nGenCount + 1
     BsHndList(nGenCount) = nBusHnd  'Keep bus handle for voltage sag checking
    End If
   End If
  End If
 Loop
 Close #1
 Close #2

 'Simulate the selected fault
 If 1 <> DoFault( PickedHandle, FltConnection, FltOption, OutageType, OutageList, _
          Rflt, Xflt, ClearPrev ) Then Exit Sub

 Call PickFault( 1 ) ' Must always pick a fault before attempting to read result

 'Check bus voltage sag
 
 If nGenCount > 0 Then
  'Prepare TTY output
  Call PrintTTY( "**********************************************************************************************************************************" )
  Call PrintTTY( "* Wind Farm Generation Cut-off Simulation " )
  Call PrintTTY( "*" )
  sLine$ = "* Voltage cut-off threshold=" + Str(dThreshold) + "p.u"
  Call PrintTTY( sLine )
  Call PrintTTY( "*" )
  Call PrintTTY( "* Wind farm generators:" ) 
 End If
 For ii=1 To nGenCount
  nBusHnd = BsHndList(ii)
  sLine$ = "* " + FullBusName(nBusHnd)
  Call GetSCVoltage( nBusHnd, vdMag, vdAng, 2 ) '012 sequence voltage in polar form
  Call GetData( nBusHnd, BUS_dKVnominal, BsKV# )
  Sag = vdMag(2)/BsKV*Sqr(3)
  If Sag >= dThreshold Then
   'Turn generator on
   nGenUnitHnd& = 0
   Do While 1 = GetBusEquipment( nBusHnd, TC_GENUNIT, nGenUnitHnd )
    If 1 <> SetData( nGenUnitHnd, GU_nOnLine, FlagActive ) Then GoTo HasError
    PostData( nGenUnitHnd)
   Loop
   sLine = sLine + ":" + Format(Sag," 0.00") + "p.u. (on-line)"
  Else
   sLine = sLine + ":" + Format(Sag," 0.00") + "p.u. (off-line)"
  End If
  Call PrintTTY( sLine )
 Next

 If nGenCount > 0 Then
  If 1 <> DoFault( PickedHandle, FltConnection, FltOption, OutageType, OutageList, _
          Rflt, Xflt, ClearPrev ) Then GoTo HasError
  Call ShowFault( 1, 5, 4, 0, ShowRelayFlag )
  sMessage$ = Str(nGenCount) + " wind generators were considered in this simulation. Details are in TTY."
  Print sMessage
 Else
  'Show fault result on the 1-line diagram
  Call ShowFault( 1, 5, 4, 0, ShowRelayFlag )
 End If
 
 Skip2:
 Exit Sub
 ' Error handling
 HasError:
 Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()
' ===================== End of Main() =========================================


' ===================== Dialog box spec (generated by Dialog Editor) ==========
'
Begin Dialog FAULTDLG 48,46, 258, 128, "Specify fault"
  OptionGroup .FLTCONN
    OptionButton 12,16,24,12, "3PH"
    OptionButton 40,16,24,12, "2LG"
    OptionButton 68,16,24,12, "1LG"
    OptionButton 96,16,24,12, "L-L"
  OptionGroup .FLTOPT
    OptionButton 12,36,44,8, "Close-in"
    OptionButton 12,44,100,8, "Close-in with end opened"
    OptionButton 12,52,68,8, "Remote bus"
    OptionButton 12,60,88,8, "Line end"
    OptionButton 12,68,60,8, "Intermediate"
    OptionButton 12,76,112,8, "Intermediate with end opened"
  Text 8,8,64,8, "Fault Connection"
  Text 8,28,64,8, "Fault Location"
  Text 128,8,64,8, "Fault impedance"
  Text 136,20,12,8, "Z="
  TextBox 148,16,40,12, .EditBox_1
  Text 192,20,12,8, "+ j"
  TextBox 204,16,40,12, .EditBox_2
  TextBox 144,68,16,12, .EditBox_3
  Text 128,72,16,8, "At %"
  TextBox 96,84,24,12, .EditBox_4
  Text 124,88,12,8, "To"
  TextBox 136,84,24,12, .EditBox_5
  Text 24,88,68,8, "Auto sequence from"
  OKButton 72,104,48,12
  CancelButton 136,104,48,12
End Dialog
Begin Dialog BUSFAULTDLG 48,46, 258, 59, "Specify fault"
  OptionGroup .FLTCONN
    OptionButton 12,16,24,12, "3PH"
    OptionButton 40,16,24,12, "2LG"
    OptionButton 68,16,24,12, "1LG"
    OptionButton 96,16,24,12, "L-L"
  Text 8,8,64,8, "Fault Connection"
  Text 128,8,64,8, "Fault impedance"
  Text 136,20,12,8, "Z="
  TextBox 148,16,40,12, .EditBox_1
  Text 192,20,12,8, "+ j"
  TextBox 204,16,40,12, .EditBox_2
  OKButton 72,40,48,12
  CancelButton 136,40,48,12
End Dialog
' ===================== End of Dialog box spec ================================

' ===================== FaultDialog() =========================================
' Purpose:
'   Get Fault spec. inputs from user
'
Function FaultDialog( FltConnection() As Long, FltOption() As Double, _ 
             DlgStyle As Long, ByRef dR As Double, ByRef Xcap As Double ) As Long
  If DlgStyle = 1 Then  ' Picked Bus
    Dim dlg As BUSFAULTDLG
    button = Dialog( dlg )
    If button = 0 Then ' Canceled
      FaultDialog = 0
      Exit Function
    End If
    FltConnection( 1 + dlg.FLTCONN ) = 1
    FltOption(1) = 1.0
    dR = Val(dlg.EditBox_1)
    Xcap = Val(dlg.EditBox_2)
  Else  ' Picked Relay group
    Dim dlg1 As FAULTDLG
    dlg1.EditBox_3 = 10  '10%
    button = Dialog( dlg1 )
    If button = 0 Then ' Canceled
      FaultDialog = 0
      Exit Function
    End If
    FltConnection( 1 + dlg1.FLTCONN ) = 1
    If dlg1.FLTOPT = 4 Or dlg1.FLTOPT = 5  Then 
      FltOption(1 + 2*dlg1.FLTOPT ) = Val(dlg1.EditBox_3)
      FltOption(13) = Val(dlg1.EditBox_4)
      FltOption(14) = Val(dlg1.EditBox_5)
    Else
      FltOption( 1 + 2*dlg1.FLTOPT ) = 1.0
    End If
    dR = Val(dlg1.EditBox_1)
    Xcap = Val(dlg1.EditBox_2)
  End If
  FaultDialog = 1
End Function
' ===================== End of FaultDialog() ==================================
