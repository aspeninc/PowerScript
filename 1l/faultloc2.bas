' ASPEN PowerScript sample program
'
' FAULTLOC.BAS
'
' Find location of a fault on a line and its neighbors using trial-and-error.
' method. First the script simulates an intermediate fault on every 1% of 
' the selected line(s). The simulation result is then  compared to recorded 
' quantied values. One or several matches will be reported.
' 
' Version 3.0
' Category: OneLiner
'
' PowerScript functions called:
'   DoFault()
'   PickFault()
'   EquipmentType()
'   GetSCCurrent()
'   FaultDescription()
'
' Main program code

Sub main()
   ' Variable declaration
   Dim ShowFlagRly(5) As Long
   Dim FltConnection(5) As Long
   Dim FltOption(15) As Double
   Dim OutageType(5) As Long
   Dim Number As Long
   Dim OutageList(20) As Long
   Dim Current(10) As Double, StepSize As Double
   Dim Voltage(10) As Double
   Dim FltR(4) As Double
   Dim MinErr(100) As Double, ErrVal As Double
   Dim PickedHnd As Long, BranchHnd As Long, DoneFlag As Long, ShowFaultFlag As Long
   Dim FaultString(100) As String
   Dim Choice(10) As Long, dFltR(100) As String
   Dim dFltX As Double
   Dim temp As Double
   Dim StringVal As String, StringVal1 As String
   Dim LargeNo As Long
   Dim Relay1Hnd As Long, Bus1Hnd As Long, Bus2Hnd As Long
   Dim Branch1Hnd As Long, TapCode As Long
   Dim SmallFloat As Double, ttt As Long, LineHnd As Long, LineName As String
   Dim BusName(50) As String, BusName1(50) As String, Counter As Long
   Dim FaultCount As Long      ' Total fault simulated
   

   ' Initialize 
   For ii = 1 To 4 
      ShowFlagRly(ii) = 1
   Next ii
   For ii = 1 To 4 
      FltConnection(ii) = 0
   Next ii
   For ii = 1 To 14
      FltOption(ii) = 0.0
   Next ii
   For ii = 1 To 3
      OutageType(ii) = 0
   Next ii
   For ii = 1 To 6
      Choice(ii) = 0
   Next ii
   dFltX     = 0.0

   LargeNo = 1  
   Counter = 1	' keep name index
   
   If GetEquipment( TC_PICKED, PickedHnd ) = 0 Then 
      Print "Please select relay group where fault recorder is located"
      Exit Sub
   End If
   ' Must be a relay group
   If EquipmentType( PickedHnd ) <> TC_RLYGROUP Then
      Print "Please select relay group where fault recorder is located"
      Exit Sub
   End If
   
   ' Get the relay group branch handle
   If GetData( PickedHnd, RG_nBranchHnd, RlyBranchHnd& ) = 0 Then GoTo HasError
   ' Get the branch near bus handle
   If GetData( RlyBranchHnd, BR_nBus1Hnd, BusHnd& ) = 0 Then GoTo HasError
   Call GetData( RlyBranchHnd, BR_nBus2Hnd, RlyBus2Hnd& )
   RlyBranchID$ = FullBusName( BusHnd ) + " - " + FullBusName( RlyBus2Hnd )
   
   BranchHnd& = LineDialog( BusHnd )
   
   If BranchHnd = 0 Then
     exit Sub
   End If
   
   ' Get line name
   Call GetData( BranchHnd, BR_nHandle,  LineHnd& )
   Call GetData( LineHnd, LN_nInService, nFlag& )
   If nFlag <> 1 Then
    Print "Line is out of service"
    exit Sub
   End If
   Call GetData( BranchHnd, BR_nBus1Hnd, LineBus1Hnd& )
   Call GetData( BranchHnd, BR_nBus2Hnd, LineBus2Hnd& )
   Call GetData( LineHnd, LN_sID, CktID$ )
   BusName(Counter) = FullBusName( LineBus1Hnd )  	' Get the near bus name
   BusName1(Counter) = FullBusName( LineBus2Hnd )	' Get the far end bus name
   LineID$ = BusName(Counter) + " - " + BusName1(Counter) + " " + CktID
   Counter = Counter + 1

   If InputDialog( FltConnection, Choice, FltR, Current, Voltage, Number, LineID, RlyBranchID ) = 0 Then Stop

   For ii = 1 To Number	' Initialize
      MinErr(ii) = 1e12
   Next ii 

   ' Initialization
   StepSize      = 1  	' Every 1%
   DoneFlag      = 0
   FltOption(1)  = 1	' Close-in
   FltOption(5)  = 1	' Remote bus
   FltOption(9)  = StepSize
   FltOption(13) = 0  	' Start from 0% 
   FltOption(14) = 100  	' To 100% 
   For ii = 1 To 15
      OutageList(ii) = 0
   Next ii
 
   ' Simulate line faults
   FaultCount = SimulateFault( FltR, BranchHnd, FltConnection, FltOption, OutageType, OutageList, dFltX, _
      ShowFlagRly, RlyBranchHnd, BusHnd, Choice, Current, Voltage, Number, MinErr, LargeNo, _
      FaultString, dFltR, 9 )
      
   If Choice(7) = 0 Then GoTo PrintResult	' Not including neighboring lines

   ' Do reverse fault
   Branch1Hnd = 0              ' Must always start from zero
   While GetBusEquipment( LineBus1Hnd, TC_BRANCH, Branch1Hnd ) > 0
      If GetData( Branch1Hnd, BR_nType, TypeCode ) = 0 Then GoTo HasError
      Call GetData( Branch1Hnd, BR_nInService, nFlag )
      If nFlag = 1 And TypeCode = TC_LINE Then
         If GetData( Branch1Hnd, BR_nBus2Hnd, Bus2Hnd ) = 0 Then GoTo HasError	' Get the far end bus
         If Bus2Hnd <> LineBus2Hnd Then ' Omit the original branch
            ' Simulate intermediate faults
            FaultCount1& = SimulateFault( FltR, Branch1Hnd, FltConnection, FltOption, OutageType, _
                                OutageList, dFltX, ShowFlagRly, RlyBranchHnd, BusHnd, Choice, _
                                Current, Voltage, Number, MinErr, LargeNo, FaultString, dFltR, 9 )
            FaultCount = FaultCount + FaultCount1
            BusName(Counter)  = FullBusName( LineBus1Hnd )	' Get the near bus name
            BusName1(Counter) = FullBusName( Bus2Hnd )	' Get the far end bus name
            Counter = Counter + 1
         End If
      End If
   Wend ' Each branch

      
   ' Get branches connected to the far bus
   Branch1Hnd = 0
   While GetBusEquipment( LineBus2Hnd, TC_BRANCH, Branch1Hnd ) > 0
      ' Get branch type
      If GetData( Branch1Hnd, BR_nType, TypeCode ) = 0 Then GoTo HasError
      Call GetData( Branch1Hnd, BR_nInService, nFlag )
      If nFlag = 1 And TypeCode = TC_LINE Then
         If GetData( Branch1Hnd, BR_nBus2Hnd, Bus2Hnd ) = 0 Then GoTo HasError	' Get the far end bus
         If Bus2Hnd <> LineBus1Hnd Then ' Omit the original branch
            ' Simulate intermediate faults on the  line
            FaultCount1 = SimulateFault( FltR, Branch1Hnd, FltConnection, FltOption, _
                                OutageType, OutageList, dFltX, ShowFlagRly, RlyBranchHnd, _
                                BusHnd, Choice, Current, Voltage, Number, MinErr, LargeNo, _
                                FaultString, dFltR, 9 )
            FaultCount = FaultCount + FaultCount1
            BusName(Counter)  = FullBusName( LineBus2Hnd )	' Get the near bus name
            BusName1(Counter) = FullBusName( Bus2Hnd )	' Get the far end bus name
            Counter = Counter + 1
         End If
      End If
   Wend	' Each branch

PrintResult:
   StringVal = FaultString(ii) & Chr(10)
   Print Str( FaultCount ) & " faults have been simulated. " & _
         "The best matched case is:    " & Chr(10) & Chr(10) & _
         FaultString(1) & Chr(10) & Chr(10) & "See TTY window for simulation details."
   
   ' Print to TTY window
   Call PrintTTY( " " )
   Call PrintTTY( " " )
   Call PrintTTY( "==================================================================================================" )
   Call PrintTTY( "Fault location report" )
   Call PrintTTY( " " )
   Call PrintTTY( "The following fault connections have been simulated:" )
   If FltConnection(1) = 1 Then Call PrintTTY( " - Three phase fault" )
   If FltConnection(2) = 1 Then Call PrintTTY( " - Two-phase to ground fault" )
   If FltConnection(3) = 1 Then Call PrintTTY( " - Single phase to ground fault" )
   If FltConnection(4) = 1 Then Call PrintTTY( " - Line to line fault" )
   Call PrintTTY( "On every 1% of following lines:" )
   ' Print lines names
   StringVal1 = ""
   For ii = 1 To Counter-1
      StringVal1 = StringVal1 & Chr(10) & "Line " & Format( ii, "#0:   " ) & "from    " _
         & BusName(ii) & "    to    " & BusName1(ii) & Chr(10)
   Next ii
   Call PrintTTY( StringVal1 )
   Call PrintTTY( " " )
     
   StringVal1 = Chr(10) & "With fault resistance from: " & Format( FltR(1), "#0.00" ) & " ohm " & " to " & _
		Format( FltR(2), "#0.00" ) & " ohm " & " increment step " & Format( FltR(3), "#0.00" ) & " ohm " & Chr(10)
   Call PrintTTY( StringVal1 )
   Call PrintTTY( " " )
   
   Call PrintTTY( "Reference quantities:" )
   If Choice(1) = 1 Then
       StringVal1 = Chr(10) & "Phase current:     " & "Ia = " & Format( Current(1), "#00000.00" ) _
		& " A        " & "Ib = " &  Format( Current(2), "#00000.00" ) & " A        " & _
		"Ic = " &  Format( Current(3), "#00000.00" ) & " A        " & Chr(10)
	Call PrintTTY( StringVal1)
   End If
   If Choice(2) = 1 Then
       StringVal1 = Chr(10) & "Phase voltage:     " & "Va = " & Format( Voltage(1), "#000.00" ) & _
		" kV         " & "Vb = " &  Format( Voltage(2), "#000.00" ) & " kV         " _
		& "Vc = " &  Format( Voltage(3), "#000.00" ) & " kV        " & Chr(10)
	Call PrintTTY( StringVal1 )
   End If
   If Choice(3) = 1 Then
       StringVal1 = Chr(10) & "3I0 = " & Format( Current(4), "#00000.00" ) & " A " & Chr(10)
	Call PrintTTY( StringVal1 )
   End If
   If Choice(4) = 1 Then
       StringVal1 = Chr(10) & "V0  = " & Format( Voltage(4), "#000.00" ) & " kV " & Chr(10)
	Call PrintTTY( StringVal1)
   End If
   If Choice(5) = 1 Then
       StringVal1 = Chr(10) & "I2  = " & Format( Current(5), "#00000.00" ) & " A " & Chr(10)
	Call PrintTTY( StringVal1)
   End If
   If Choice(6) = 1 Then
       StringVal1 = Chr(10) & "V2  = " & Format( Voltage(5), "#000.00" ) & " kV " & Chr(10)
	Call PrintTTY( StringVal1)
   End If
   Call PrintTTY( " " )
   
   StringVal1 = Chr(10) & "Best matched cases are:" & Chr(10)
   Call PrintTTY( StringVal1 )
   Call PrintTTY( " " )
   For ii =1 To Number
      StringVal1 = Chr(10) & "Case " & Format( ii, "#0:  " ) & FaultString(ii) & Chr(10) & _
                   "   Faulted quantities:" & Chr(10) & "      " & dFltR(ii) & Chr(10)
      Call PrintTTY( StringVal1 )
      StringVal1 = "   error = " & Format( MinErr(ii), "#0" ) & Chr(10)
      Call PrintTTY( StringVal1 )
      Call PrintTTY( " " )
   Next ii 
   Exit Sub
HasError:
   Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()
' ===================== End of Main() =========================================

' ===================== Dialog box spec (generated by Dialog Editor) ==========
'
'
'Try these different styles or-ed together as the last parameter of Textbox
' to define the text box style.
Const ES_LEFT             = &h0000&
Const ES_CENTER           = &h0001&
Const ES_RIGHT            = &h0002&
Const ES_MULTILINE        = &h0004&
Const ES_UPPERCASE        = &h0008&
Const ES_LOWERCASE        = &h0010&
Const ES_PASSWORD         = &h0020&
Const ES_AUTOVSCROLL      = &h0040&
Const ES_AUTOHSCROLL      = &h0080&
Const ES_NOHIDESEL        = &h0100&
Const ES_OEMCONVERT       = &h0400&
Const ES_READONLY         = &h0800&
Const ES_WANTRETURN       = &h1000&
Const ES_NUMBER           = &h2000&
Const WS_VSCROLL          = &h00200000&
Const WSTYLE = WS_VSCROLL Or ES_MULTILINE Or ES_AUTOVSCROLL Or ES_READONLY
Begin Dialog DIALOG_LINE 72,34, 238, 138, "Select Transmission Line"
  TextBox 104,120,37,12, .EditBox_2
  Text 4,4,105,12, "Transmission lines in vicinity:"
  TextBox 4,16,224,96, .EditBox_1, WSTYLE
  Text 8,120,88,12, "Enter line index number:"
  OKButton 148,120,37,12
  CancelButton 192,120,37,12
End Dialog
' ===================== LineDialog() =========================================
' Purpose:
'   Get Fault spec. inputs from user
'
Function LineDialog( ByVal BusHnd As long ) As long
 Dim dlg As DIALOG_LINE
 dim map(1000) As long
 dim stack(20) As long
 dim stackTop As long
 dim tiers(20) As long
 
 LineDialog   = 0
 
 For ii& = 1 to 1000
   map(ii) = 0
 Next
 
 stackTop& = 1
 stack(stackTop) = BusHnd
 tiers(stackTop) = 0
 
 dlg.EditBox_1 = ""
 tierLimit = 1
 
 Do 
  Do
   Bus1Hnd&  = stack(stackTop)
   thisTier& = tiers(stackTop)
   stackTop  = stackTop -1
  Loop While thisTier > tierLimit And stackTop > 0
  If thisTier > tierLimit Then exit Do
  
  ' Get all transmission lines near this bus
  Branch1Hnd& = 0              ' Must always start from zero
  While GetBusEquipment( Bus1Hnd, TC_BRANCH, Branch1Hnd ) > 0
   Call GetData( Branch1Hnd, BR_nType, TypeCode )
   If TypeCode = TC_LINE Then
    Call GetData( Branch1Hnd, BR_nHandle, LineHnd& ) 	' Get the equipment handle
    If map(LineHnd) = 0 Then
     Call GetData( Branch1Hnd, BR_nBus2Hnd, Bus2Hnd& )    ' Get the far end bus
     bFound = False
     For jj& = 1 to stackTop
      If Bus2Hnd = stack(jj) Then
       bFound = True
       exit For
      End If
     Next
     If Not bFound Then
      stackTop = stackTop + 1
      stack(stackTop) = Bus2Hnd
      tiers(stackTop) = thisTier + 1
     End If
     Call GetData( LineHnd, LN_sID, CktID$ )
     LineName$ = "# " + Str(Branch1Hnd) + ": " + FullBusName( Bus1Hnd ) + _
                " - " + FullBusName( Bus2Hnd ) + " " + CktID
     dlg.EditBox_1 = dlg.EditBox_1 + LineName + Chr(13) + Chr(10)
     map(LineHnd) = 1
    End If
   End If
  Wend ' Each branch
 Loop While stackTop > 0
 
 button = Dialog( dlg )
 If button = 0 Then ' Canceled
   Exit Function
 End If
 LineDialog = val(dlg.EditBox_2)
 Exit Function
HasErrorL:
 Print "Error"
End Function
'
' ===================== End of Dialog box spec ================================

'TODO: Tidy up the dialog box

' ===================== Dialog box spec (generated by Dialog Editor) ==========
'
'
Begin Dialog FAULTDLG 131,-56, 287, 289, "Locate a fault"
  GroupBox 20,120,244,116, "Reference quantities"
  Text 16,16,120,8, "Simulate Fault Connections:"
  CheckBox 28,28,28,8, "3PH", .CheckBox_1
  CheckBox 100,28,28,8, "2LG", .CheckBox_2
  CheckBox 172,28,28,8, "1LG", .CheckBox_3
  CheckBox 236,28,24,8, "LL", .CheckBox_4
  Text 16,48,120,8, "With Fault Resistance (ohm):"
  CheckBox 16,84,128,8, "Include Neighboring Lines", .CheckBox_11
  CheckBox 24,132,132,8, "Phase Current (A):", .CheckBox_5
  Text 32,148,36,12, "Phase A ="
  TextBox 68,148,32,12, .EditBox_1
  Text 32,164,40,8, "Phase B ="
  TextBox 68,164,32,12, .EditBox_2
  CheckBox 184,132,76,8, "Phase Voltage (kV):", .CheckBox_6
  Text 192,148,40,8, "Phase A ="
  TextBox 228,148,32,12, .EditBox_4
  Text 192,164,36,12, "Phase B ="
  TextBox 228,164,32,12, .EditBox_5
  Text 192,180,44,8, "Phase C ="
  TextBox 228,180,32,12, .EditBox_6
  CheckBox 24,200,36,12, "3I0 (A) =", .CheckBox_7
  TextBox 64,200,32,12, .EditBox_7
  CheckBox 184,200,36,12, "V0 (kV) =", .CheckBox_8
  TextBox 224,200,32,12, .EditBox_8
  CheckBox 24,216,36,12, "I2 (A)   =", .CheckBox_9
  TextBox 64,216,32,12, .EditBox_9
  CheckBox 184,216,36,12, "V2 (kV) =", .CheckBox_10
  TextBox 224,216,32,12, .EditBox_10
  OKButton 76,268,92,12
  CancelButton 172,268,40,12
  Text 24,248,96,12, "No. of Best Cases to Output ="
  TextBox 120,248,20,12, .EditBox_14
  Text 28,60,24,12, "From ="
  TextBox 56,60,20,12, .EditBox_11
  Text 124,60,20,12, "To ="
  TextBox 148,60,20,12, .EditBox_12
  Text 32,180,40,8, "Phase C ="
  TextBox 68,180,32,12, .EditBox_3
  Text 216,60,24,12, "Step ="
  TextBox 240,60,20,12, .EditBox_13
  Text 20,0,29,12, "Line ID:"
  Text 4,104,45,12, "Recorder at:"
  TextBox 56,0,216,12, .EditBox_15, ES_READONLY
  TextBox 56,104,216,12, .EditBox_16, ES_READONLY
End Dialog

'
' ===================== End of Dialog box spec ================================
'
' ===================== InputDialog() =========================================
' Purpose:
'   Get Fault spec. inputs from user
'
Function InputDialog( FltConnection() As Long, Choice() As Long, FltR() As Double, _
	ByRef Current() As Double, ByRef Voltage() As Double, ByRef Number As Long, _
	ByRef LineID As String, ByRef RlyID As String ) As Long
  Dim dlg As FAULTDLG
  ' Check all fault connection
  dlg.CheckBox_1 = 1
  dlg.CheckBox_2 = 1
  dlg.CheckBox_3 = 1
  dlg.CheckBox_4 = 1
  dlg.CheckBox_5 = 1
  dlg.CheckBox_6 = 1
  dlg.CheckBox_7 = 0
  dlg.CheckBox_8 = 0
  dlg.CheckBox_9 = 0
  dlg.CheckBox_10= 0
  dlg.CheckBox_11= 0
  dlg.EditBox_1  = 2967
  dlg.EditBox_2  = 37
  dlg.EditBox_3  = 20
  dlg.EditBox_4  = 41.5
  dlg.EditBox_5  = 77.0
  dlg.EditBox_6  = 76.8
  dlg.EditBox_7  = 2961
  dlg.EditBox_8  = 12.6
  dlg.EditBox_9  = 995
  dlg.EditBox_10 = 11.3
  dlg.EditBox_11 = 0.0
  dlg.EditBox_12 = 0.0
  dlg.EditBox_13 = 0.2
  dlg.EditBox_14 = 5
  dlg.EditBox_15 = LineID
  dlg.EditBox_16 = RlyID

  ' Initialization
  For ii = 1 To 5
     Current(ii) = 0.0
     Voltage(ii) = 0.0
  Next
  For ii = 1 To 3	
     FltR(ii) = 0.0 
  Next  
  
  DoneFlag = 0
  While DoneFlag = 0
    button = Dialog( dlg )
    If button = 0 Then ' Canceled
      InputDialog = 0
      Exit Function
    End If
    FltConnection(1) = dlg.CheckBox_1
    FltConnection(2) = dlg.CheckBox_2
    FltConnection(3) = dlg.CheckBox_3
    FltConnection(4) = dlg.CheckBox_4
    Choice(1) = dlg.CheckBox_5	' phase current
    Choice(2) = dlg.CheckBox_6	' phase voltage
    Choice(3) = dlg.CheckBox_7	' 3I0
    Choice(4) = dlg.CheckBox_8	' V0
    Choice(5) = dlg.CheckBox_9	' I2
    Choice(6) = dlg.CheckBox_10	' V2
    Choice(7) = dlg.CheckBox_11	' Include neighboring lines

    If (Choice(1) = 1) Then	' phase current
	Current(1) = Val(dlg.EditBox_1)
    	Current(2) = Val(dlg.EditBox_2)
    	Current(3) = Val(dlg.EditBox_3)
    End If
    If (Choice(2) = 1) Then	' phase voltage
	Voltage(1) = Val(dlg.EditBox_4)
    	Voltage(2) = Val(dlg.EditBox_5)
    	Voltage(3) = Val(dlg.EditBox_6)
    End If
    If (Choice(3) = 1) Then	' 3I0
	Current(4) = Val(dlg.EditBox_7)
    End If
    If (Choice(4) = 1) Then	' V0
	Voltage(4) = Val(dlg.EditBox_8)
    End If
    If (Choice(5) = 1) Then	' I2
	Current(5) = Val(dlg.EditBox_9)
    End If
    If (Choice(6) = 1) Then	' V2
	Voltage(5) = Val(dlg.EditBox_10)
    End If
    FltR(1) = Val(dlg.EditBox_11)	' From
    FltR(2) = Val(dlg.EditBox_12)	' To
    FltR(3) = Val(dlg.EditBox_13)	' Step
    Number = Val(dlg.EditBox_14)	' No.

    If (FltConnection(1)=1 Or FltConnection(2)=1 Or FltConnection(3)=1 _
           Or FltConnection(4)=1) Then
	If (FltR(1) < 0 Or FltR(2) < 0 Or FltR(3) < 0) Then
	   Print "Fault resistance must be >=0"
	Else
	   If (Current(1) > 0 Or Current(2) > 0 Or Current(3) > 0) Then 
    	      DoneFlag = 1
          ElseIf (Voltage(1) > 0 Or Voltage(2) > 0 Or Voltage(3) > 0) Then
             DoneFlag = 1
	   ElseIf (Current(4) > 0) Then
             DoneFlag = 1
	   ElseIf (Voltage(4) > 0) Then
             DoneFlag = 1
	   ElseIf (Current(5) > 0) Then
             DoneFlag = 1
	   ElseIf (Voltage(5) > 0) Then
             DoneFlag = 1
          Else
             Print "Must input fault current or voltage"
   	   End If   
	End If 
    Else
    	Print "Must select a fault connection"
    End If
  Wend
  InputDialog = 1
End Function
' ===================== End of InputDialog() ==================================

' ===================== SimulateFault() =======================================
' Purpose:
'    Simulate intermediate faults
Function SimulateFault( ByRef FltR() As Double, ByRef PickedHnd As Long, ByRef FltConnection() As Long, _
   ByRef FltOption() As Double, ByRef OutageType() As Long, ByRef OutageList() As Long, ByVal dFltX As Double, _
   ByRef ShowFlagRly() As Long, ByVal BranchHnd As Long, ByVal BusHnd As Long, ByRef Choice() As Long, ByRef Current() As Double, _
   ByRef Voltage() As Double, ByVal Number As Long, ByRef MinErr() As Double, ByRef LargeNo As Long, _
   ByRef FaultString() As String, ByRef dFltR() As String, ByVal Tier As Long) As Long

   Dim MagArray(12) As Double
   Dim AngArray(12) As Double
   Dim MagArray1(12) As Double
   Dim AngArray1(12) As Double
   Dim VmagArray(3) As Double
   Dim VangArray(3) As Double
   Dim VmagArray1(3) As Double
   Dim VangArray1(3) As Double
   Dim TempVal As Long, IntVal As Double, IntString As String

   SimulateFault = 0
   For temp = FltR(1) To FltR(2) Step FltR(3)
      If 0 = DoFault( PickedHnd, FltConnection, FltOption, OutageType, OutageList, _
                   temp, dFltX, 1 ) Then GoTo HasError1
      ' Check fault result and find the best match
      ShowFaultFlag = 1 ' Starting from the first one
      ' Pick fault does not update single line diagram screen like ShowFault
      ' While PickFault( ShowFaultFlag ) > 0
      While ShowFault( ShowFaultFlag, Tier, 1, 0, ShowFlagRly ) > 0
         ' output ABC phase branch current in polar form
    	   If GetSCCurrent( BranchHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError1
    	   ' output 012 sequence branch current in polar form
    	   If GetSCCurrent( BranchHnd, MagArray1, AngArray1, 2 ) = 0 Then GoTo HasError1
    	   ' output ABC phase bus voltage in polar form
    	   If GetSCVoltage( BusHnd, VmagArray, VangArray, 4 ) = 0 Then GoTo HasError1
    	   ' output 012 sequence bus voltage in polar form
    	   If GetSCVoltage( BusHnd, VmagArray1, VangArray1, 2 ) = 0 Then GoTo HasError1

    	   ' Compute squared error
    	   ErrVal = 0.0
    	   If Choice(1) = 1 Then	' phase current
	      ErrVal = ErrVal + (MagArray(1) - Current(1))*(MagArray(1) - Current(1)) + _
             	  (MagArray(2) - Current(2))*(MagArray(2) - Current(2)) + _
             	  (MagArray(3) - Current(3))*(MagArray(3) - Current(3))
    	   End If
    	   If Choice(2) = 1 Then	' phase voltage
	      ErrVal = ErrVal + (VmagArray(1) - Voltage(1))*(VmagArray(1) - Voltage(1)) + _
                (VmagArray(2) - Voltage(2))*(VmagArray(2) - Voltage(2)) + _
                (VmagArray(3) - Voltage(3))*(VmagArray(3) - Voltage(3))
    	   End If
    	   If Choice(3) = 1 Then	' 3I0
	      ErrVal = ErrVal + (MagArray1(1) - Current(4)/3.0)*(MagArray1(1) - Current(4)/3.0)
    	   End If
    	   If Choice(4) = 1 Then	' V0
	      ErrVal = ErrVal + (VmagArray1(1) - Voltage(4))*(VmagArray1(1) - Voltage(4))
    	   End If
    	   If Choice(5) = 1 Then	' I2
	      ErrVal = ErrVal + (MagArray1(3) - Current(5))*(MagArray1(3) - Current(5))
    	   End If
    	   If Choice(6) = 1 Then	' V2
	      ErrVal = ErrVal + (VmagArray1(3) - Voltage(5))*(VmagArray1(3) - Voltage(5))
    	   End If

         If ErrVal < MinErr(LargeNo) Then
            ' Get Fault description string
            FltDesc$ = FaultDescription()
            FltDesc  = Right( FltDesc, Len( FltDesc ) - 4 )
            
            ' record
            MinErr(LargeNo)= ErrVal
            FaultString(LargeNo) = FltDesc
            dFltR(LargeNo) = ""
            If Choice(1) =1 Then 	' phase current
               dFltR(LargeNo) = dFltR(LargeNo) & "Ia = " & Format( MagArray(1), "#00000.00" ) & " A        " & _
                       "Ib = " &  Format( MagArray(2), "#00000.00" ) & " A        " & _
		          "Ic = " &  Format( MagArray(3), "#00000.00" ) & " A        " & Chr(10)
            End If
            If Choice(2) = 1 Then	' phase voltage
               dFltR(LargeNo) = dFltR(LargeNo) & "Va = " & Format( VmagArray(1), "#000.00" ) & " kV         " & _
                       "Vb = " &  Format( VmagArray(2), "#000.00" ) & " kV         " & _
		          "Vc = " &  Format( VmagArray(3), "#000.00" ) & " kV        " & Chr(10)
            End If
            If Choice(3) = 1 Then	' 3I0
               dFltR(LargeNo) = dFltR(LargeNo) & "3I0 = " & Format( MagArray1(1)*3.0, "#00000.00" ) & " A " & Chr(10)
            End If
            If Choice(4) = 1 Then	' V0
               dFltR(LargeNo) = dFltR(LargeNo) & "V0  = " & Format( VmagArray1(1), "#000.00" ) & " kV " & Chr(10)
            End If
            If Choice(5) = 1 Then	' I2
               dFltR(LargeNo) = dFltR(LargeNo) & "I2  = " & Format( MagArray1(3), "#00000.00" ) & " A " & Chr(10)
            End If
            If Choice(6) = 1 Then	' V2
               dFltR(LargeNo) = dFltR(LargeNo) & "V2  = " & Format( VmagArray1(3), "#000.00" ) & " kV " & Chr(10)
            End If

      	      ' Insert one
            If LargeNo < Number Then
               TempVal = LargeNo	' remember the inserted element
               LargeNo = LargeNo + 1
            Else ' replace the largest
               TempVal = LargeNo	' remember the inserted element
            End If
              
            ' Sorting from smallest to largest
            While (TempVal-1) > 0
               If MinErr(TempVal) < MinErr(TempVal-1) Then
                  ' Switch elements
                  IntVal = MinErr(TempVal)
                  MinErr(TempVal) = MinErr(TempVal-1)
                  MinErr(TempVal-1) = IntVal
                  IntString = dFltR(TempVal)
                  dFltR(TempVal) = dFltR(TempVal-1)
                  dFltR(TempVal-1) = IntString
                  IntString = FaultString(TempVal)
                  FaultString(TempVal) = FaultString(TempVal-1)
                  FaultString(TempVal-1) = IntString
                  TempVal = TempVal - 1
		   Else
		      GoTo EndSort
		   End If
		Wend
    	   End If
          
	   EndSort:
          SimulateFault = SimulateFault + 1
     	   ShowFaultFlag = SF_NEXT   ' Show next fault
   	Wend
   Next temp
HasError1:
   
End Function
' ===================== End of SimulateFault() ==================================
