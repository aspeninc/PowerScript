' ASPEN PowerScrip sample program
'
' Submitted by: "Smith, Nolan" <NSmith@belco.bhl.bm>
'
' MyFaults.BAS
'
' Simulate fault at one bus in the network 
' Record fault current, Thevenin impedance and X/R
' in a text or csv file and write out all line flows
'
' Version 1.0
' Category: OneLiner
'
' PowerScript functions called:
'   GetEquipment()
'   NextBusByName()
'   GetBusEquipment()
'   FaultDescription()
'   DoFault()
'   PickFault()
'   GetSCCurrent()
'
Sub main()
   ' Variable declaration
   Dim Imag(16) As Double
   Dim Iang(16) As Double
   Dim FltConn(4) As Long
   Dim FltOpt(14) As Double
   Dim OutageOpt(4) As Long
   Dim OutageLst(30) As Long

Begin Dialog Dialog_1 49,60, 152, 95, "Specify Fault"
  OptionGroup .GROUP_1
    OptionButton 84,56,28,8, "Text"
    OptionButton 116,56,28,8, "CSV"
  Text 12,44,60,12, "Output file name: "
  TextBox 12,56,68,12, .EditBox_1
  OKButton 20,76,52,12
  CancelButton 84,76,48,12
  GroupBox 8,8,132,32, "Fault connection"
  CheckBox 16,20,28,12, "1LG", .CheckBox_1
  CheckBox 48,20,24,12, "2LG", .CheckBox_2
  CheckBox 80,20,28,12, "3PH", .CheckBox_3
  CheckBox 112,20,20,12, "LL", .CheckBox_4
End Dialog
   Dim dlg As Dialog_1

   dlg.CheckBox_1 = 1
   dlg.CheckBox_2 = 0
   dlg.CheckBox_3 = 1
   dlg.CheckBox_4 = 0
   dlg.EditBox_1 = "c:\AspenFault\busflt"
   dlg.group_1   = 0
   
   If 0 = Dialog(dlg) Then Exit Sub
   
   If dlg.CheckBox_1=0 And dlg.CheckBox_2=0 And dlg.CheckBox_3=0 And dlg.CheckBox_4=0 Then
     Print "Must specify fault connection"
     Exit Sub
   End If

   If dlg.group_1 = 0 Then
     OutputText = 1
   Else
     OutputText = 0
   End If

   'fault connections
   FltConn(1) = dlg.CheckBox_1   ' Do 3PH
   FltConn(2) = dlg.CheckBox_2   ' Do 2LG
   FltConn(3) = dlg.CheckBox_3   ' Do 1LG
   FltConn(4) = dlg.CheckBox_4   ' Do LL
   FltOpt(1)  = 1   ' Bus fault no outage
   FltOpt(2)  = 0
   Rflt         = 0 ' Fault R
   Xflt         = 0 ' Fault X
   ClearPrev    = 1 ' Clear previous result initially

   If GetEquipment( TC_PICKED, BusHnd ) = 0 Then 
     Print "Must select a bus"
     Exit Sub
   End If
   If EquipmentType( BusHnd ) <> TC_BUS Then
     Print "Must select a bus"
     Exit Sub
   End If

   
    ' Get picked bus handle number
 If GetEquipment( TC_PICKED, BusHnd& ) = 0 Then 
     BusHnd   = 0
     If 0 = NextBusByName( BusHnd ) Then GoTo HasError
     DoOneBus = 0
   Else
     DoOneBus = 1
   End If
   Delim$ = Chr(34) & "," & Chr(34)

   Open Dlg.EditBox_1 For Output As 1

   ' Print report title
   If OutputText = 1 Then
     Print #1, "                      BUS FAULT REPORT"
     Print #1, ""
     Print #1, "                      Fault current                               Thevenin Impedance"
     Print #1, "______________________Phase_A________Phase B____________Phase C___R0+jX0______R1+jX1_____R2+jX2_____X/R"
     Print #1, ""
   Else
     Print #1, Chr(34) & "Fault" & Delim & "PhaseA" & Delim & "PhaseB" & Delim & "PhaseC" & Delim & _
               "R0+jX0" & Delim & "R1+jX1" & Delim & "R2+jX2" & Delim & "X/R" & Chr(34)
   End If
   NoFaults = 0
   MaxCurrent = 0
   Do
     ' Simulate the NoFaults
     If DoFault( BusHnd, FltConn, FltOpt, OutageOpt, OutageLst, _
                 Rflt, Xflt, ClearPrev ) = 0 _
           Then GoTo HasError
     ClearPrev = 0	' Keep result from now on
   Loop Until (DoOneBus = 1 Or NextBusByName( BusHnd ) <= 0)

   FaultFlag = 1	' Show First fault first
   While PickFault( FaultFlag ) <> 0
     ' Get current
     If GetSCCurrent( HND_SC, Imag, Iang, 4 ) = 0 Then GoTo HasError
     ' Get Thevenin equivalent
     If GetData( HND_SC, FT_dRPt, R1t# ) = 0 Then GoTo HasError
     If GetData( HND_SC, FT_dRNt, R2t# ) = 0 Then GoTo HasError
     If GetData( HND_SC, FT_dRZt, R0t# ) = 0 Then GoTo HasError
     If GetData( HND_SC, FT_dXPt, X1t# ) = 0 Then GoTo HasError
     If GetData( HND_SC, FT_dXNt, X2t# ) = 0 Then GoTo HasError
     If GetData( HND_SC, FT_dXZt, X0t# ) = 0 Then GoTo HasError
     If GetData( HND_SC, FT_dXR,  XRt# ) = 0 Then GoTo HasError
     FltDescription$ = FaultDescription( )
     If OutputText = 1 Then 
       Print #1, FltDescription$
       Print #1,    "                      " ; _
          Format( Imag(1), "####0.0") & "@" & Format( Iang(1), "#0.0"), "   ", _
          Format( Imag(2), "####0.0") & "@" & Format( Iang(2), "#0.0"), "   ", _
          Format( Imag(3), "####0.0") & "@" & Format( Iang(3), "#0.0"), "   ", _
          Format( R0t,    "####0.0") & "+j" & Format( X0t,    "#0.0"), "   ", _
          Format( R1t,    "####0.0") & "+j" & Format( X1t,    "#0.0"), "   ", _
          Format( R2t,    "####0.0") & "+j" & Format( X2t,    "#0.0"), "   ", _
          Format( XRt,    "####0.0") 
'*********************************
'Get Line Flows
   DevHandle = 0
   LineID    = ""
   Bus1ID    = ""
   Bus2ID    = ""
   Counts    = 0
   While GetEquipment( TC_LINE, DevHandle ) > 0
     ' Get Line ID and end bus names
     If GetData( DevHandle , LN_sID,   LineID ) = 0 Then GoTo HasError
     If GetData( DevHandle , LN_nBus1Hnd, Bus1Handle ) = 0 Then GoTo HasError
     If GetData( DevHandle , LN_nBus2Hnd, Bus2Handle ) = 0 Then GoTo HasError
     ' Get current flows
     If GetSCCurrent( DevHandle, Imag, Iang, 4 ) = 0 Then GoTo HasError
	Print #1, _
        FullBusName( Bus1Handle ) & " - "; FullBusName( Bus2Handle) & " " & LineID _
        & "  " & Format( Imag(1),  "    ####.000  " ) & Format( Iang(1),  "    ####.000  " )
     Counts = Counts + 1
   Wend
   Print Counts; " Lines Exported"
'Get Trans Flows
   DevHandle = 0
   LineID    = ""
   Bus1ID    = ""
   Bus2ID    = ""
   Counts    = 0
   While GetEquipment( TC_XFMR, DevHandle ) > 0
     ' Get Line ID and end bus names
     If GetData( DevHandle , XR_sID,   LineID ) = 0 Then GoTo HasError
     If GetData( DevHandle , XR_nBus1Hnd, Bus1Handle ) = 0 Then GoTo HasError
     If GetData( DevHandle , XR_nBus2Hnd, Bus2Handle ) = 0 Then GoTo HasError
     ' Get current flows
     If GetSCCurrent( DevHandle, Imag, Iang, 4 ) = 0 Then GoTo HasError
	Print #1, _
        FullBusName( Bus1Handle ) & " - "; FullBusName( Bus2Handle) & " " & LineID _
        & "  " & Format( Imag(1),  "    ####.000  " ) & Format( Iang(1),  "    ####.000  " )
     Counts = Counts + 1
   Wend
   Print Counts; " Transformers Exported"


'*********************************

     Else
       Print #1, Chr(34) & FltDescription$ & Delim & _
          Format( Imag(1), "####0.0") & "@" & Format( Iang(1), "#0.0") & Delim & _
          Format( Imag(2), "####0.0") & "@" & Format( Iang(2), "#0.0") & Delim & _
          Format( Imag(3), "####0.0") & "@" & Format( Iang(3), "#0.0") & Delim & _
          Format( R0t,    "####0.0") & "+j" & Format( X0t,    "#0.0")  & Delim & _
          Format( R1t,    "####0.0") & "+j" & Format( X1t,    "#0.0")  & Delim & _
          Format( R2t,    "####0.0") & "+j" & Format( X2t,    "#0.0")  & Delim & _
          Format( XRt,    "####0.0") & Chr(34)
'*********************************
'Get all Flows
 DevHandle = 0
   LineID    = ""
   Bus1ID    = ""
   Bus2ID    = ""
   Counts    = 0
   While GetEquipment( TC_LINE, DevHandle ) > 0
     ' Get Line ID and end bus names
     If GetData( DevHandle , LN_sID,   LineID ) = 0 Then GoTo HasError
     If GetData( DevHandle , LN_nBus1Hnd, Bus1Handle ) = 0 Then GoTo HasError
     If GetData( DevHandle , LN_nBus2Hnd, Bus2Handle ) = 0 Then GoTo HasError
     ' Get current flows
     If GetSCCurrent( DevHandle, Imag, Iang, 4 ) = 0 Then GoTo HasError
	Print #1, _
        FullBusName( Bus1Handle ) & Delim ; FullBusName( Bus2Handle) & Delim & LineID _
        & Delim & Format(Imag(1),  "    ####.000  " ) & Delim & Format(Iang(1),  "    ####.000  " )
     Counts = Counts + 1
   Wend
   Print Counts; " Lines Exported"

'Get Trans Flows
   DevHandle = 0
   LineID    = ""
   Bus1ID    = ""
   Bus2ID    = ""
   Counts    = 0
   While GetEquipment( TC_XFMR, DevHandle ) > 0
     ' Get Line ID and end bus names
     If GetData( DevHandle , XR_sID,   LineID ) = 0 Then GoTo HasError
     If GetData( DevHandle , XR_nBus1Hnd, Bus1Handle ) = 0 Then GoTo HasError
     If GetData( DevHandle , XR_nBus2Hnd, Bus2Handle ) = 0 Then GoTo HasError
     ' Get current flows
     If GetSCCurrent( DevHandle, Imag, Iang, 4 ) = 0 Then GoTo HasError
	Print #1, _
        FullBusName( Bus1Handle ) & Delim ; FullBusName( Bus2Handle) & Delim & LineID _
        & Delim & Format(Imag(1),  "    ####.000  " ) & Delim & Format(Iang(1),  "    ####.000  " )
     Counts = Counts + 1
     Counts = Counts + 1
   Wend
   Print Counts; " Transformers Exported"



'*********************************

     End If
     FaultFlag = SF_NEXT	' Show next fault
     NoFaults  = NoFaults + 1 ' Update # of imulated NoFaults
   Wend

   Close ' Close output file
   Print NoFaults & " Faults Simulated"
   Exit Sub
   HasError:
   Print "Error: ", ErrorString( )
   Close
End Sub

Function OpenOutFile() As Long
   ' Open file for output

   ' Dialog data generated by Dialog Edito
Begin Dialog Dialog_1 49,60, 152, 95, "Specify Fault"
  OptionGroup .GROUP_1
    OptionButton 84,56,28,8, "Text"
    OptionButton 116,56,28,8, "CSV"
  Text 12,44,60,12, "Output file name: "
  TextBox 12,56,68,12, .EditBox_1
  OKButton 20,76,52,12
  CancelButton 84,76,48,12
  GroupBox 8,8,132,32, "Fault connection"
  CheckBox 16,20,28,12, "1LG", .CheckBox_1
  CheckBox 48,20,24,12, "2LG", .CheckBox_2
  CheckBox 80,20,28,12, "3PH", .CheckBox_3
  CheckBox 112,20,20,12, "LL", .CheckBox_4
End Dialog
   Dim dlg As Dialog_1
   Dlg.EditBox_1 = "c:\AspenFault\b0.rep"         ' Default name
   ' Dialog returns -1 for OK, 0 for Cancel, button # for PushButtons
   button = Dialog( Dlg )
   If button = 0 Then 
      OpenOutFile = 0
      Exit Function
   End If
   fileName = Dlg.EditBox_1
   Open fileName For Output As #1
   OpenOutFile = 1
End Function
