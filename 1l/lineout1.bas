' ASPEN PowerScript Sample Program
'
' LINEOUT1.BAS
'
' Simulate fault at selected relay group with and without branch outage
' and up to 3 fault impedances
'
' Demonstrate how to simulate fault from a PowerScript program
'
' Version 2.0
' Category: OneLiner
'
' PowerScript functions called:
'   GetEquipment()
'   FullBusName()
'   DoFault
'   ShowFault()
'   GetSCVoltage()
'   GetSCCurrent()
'


Sub main()
   ' Variable declaration
   Dim MagArray(16) As Double
   Dim AngArray(16) As Double
   Dim FltConn(4) As Long
   Dim FltOption(14) As Double
   Dim OutageList(20) As Long
   Dim BranchList(50) As String, BusList(3) As String
   Dim OutageType(3) As Long
   Dim ShowRelayFlag(4) As Long
   Dim FltR(3) As Double, FltX(3) As Double
   Dim RlyHnd(3)

   ' Make sure a line with relay group is being selected
   If GetEquipment( TC_PICKED, DevHnd& ) = 0 Then 
     Print "Must select a line with relay group"
     Exit Sub
   End If
   If EquipmentType( DevHnd ) <> TC_LINE Then
     Print "Must select a line with relay group"
     Exit Sub
   End If

   If GetData( DevHnd, LN_nRlyGr1Hnd, TempHnd& ) > 0 Then
     RlyHnd(1) = TempHnd
   Else
     RlyHnd(1) = 0
   End If 
   If GetData( DevHnd, LN_nRlyGr2Hnd, TempHnd& ) > 0 Then
     RlyHnd(2) = TempHnd
   Else
     RlyHnd(2) = 0
   End If 

   If RlyHnd(1) = 0 And RlyHnd(2) = 0 Then
     Print "Must select a line with relay group"
     Exit Sub
   End If
   
   ' Inventory of all branches at 2 end buses
   If GetData( DevHnd, LN_nBus1Hnd, BusHnd1& ) = 0 Then GoTo HasError
   If GetData( DevHnd, LN_nBus2Hnd, BusHnd2& ) = 0 Then GoTo HasError
   BusList(0) = "Simulate fault on:"
   BusList(1) = FullBusName( BusHnd1 )
   BusList(2) = FullBusName( BusHnd2 )
   BrHnd&   = 0
   BrCounts = 0
   While GetBusEquipment( BusHnd1, TC_BRANCH, BrHnd ) > 0
     If GetData( BrHnd, BR_nBus1Hnd, BusHndA& ) = 0 Then GoTo HasError
     If GetData( BrHnd, BR_nBus2Hnd, BusHndB& ) = 0 Then GoTo HasError
     If (BusHnd1 <> BusHndA Or BusHnd2 <> BusHndB) And _
        (BusHnd1 <> BusHndB Or BusHnd2 <> BusHndA ) Then    ' Must not count the faulted branch
       BrHndStr$ = BrHnd
       BrList$ = "[" + BrHndStr$ + "] "
       BrList$ = BrList$ + FullBusName( BusHndA ) + " - "
       BrList$ = BrList$ + FullBusName( BusHndB )
       If GetData( BrHnd, BR_nType, BrType& ) = 0 Then GoTo HasError
       Select Case BrType
         Case TC_LINE
            BrList$ = BrList$ + " L"
         Case TC_XFMR
            BrList$ = BrList$ + " T"
         Case TC_XFMR3
            BrList$ = BrList$ + " X"
         Case TC_SHIFTER
            BrList$ = BrList$ + " L"
       End Select
       BranchList(BrCounts) = BrList$
       BrCounts = BrCounts + 1
     End If
   Wend
   BrHnd&   = 0
   While GetBusEquipment( BusHnd2, TC_BRANCH, BrHnd ) > 0
     If GetData( BrHnd, BR_nBus1Hnd, BusHndA& ) = 0 Then GoTo HasError
     If GetData( BrHnd, BR_nBus2Hnd, BusHndB& ) = 0 Then GoTo HasError
     If (BusHnd1 <> BusHndA Or BusHnd2 <> BusHndB) And _
        (BusHnd1 <> BusHndB Or BusHnd2 <> BusHndA ) Then    ' Must not count the faulted branch
       BrHndStr$ = BrHnd
       BrList$ = "[" + BrHndStr$ + "] "
       BrList$ = BrList$ + FullBusName( BusHndA ) + " - "
       BrList$ = BrList$ + FullBusName( BusHndB )
       If GetData( BrHnd, BR_nType, BrType& ) = 0 Then GoTo HasError
       Select Case BrType
         Case TC_LINE
            BrList$ = BrList$ + " L"
         Case TC_XFMR
            BrList$ = BrList$ + " T"
         Case TC_XFMR3
            BrList$ = BrList$ + " X"
         Case TC_SHIFTER
            BrList$ = BrList$ + " L"
       End Select
       BranchList(BrCounts) = BrList$
       BrCounts = BrCounts + 1
     End If
   Wend

'=============Dialog Spec=============
Begin Dialog Dialog_1 25,58, 338, 245, "Specify Fault"
  GroupBox 216,52,112,28, "Phase connection"
  GroupBox 216,136,109,64, "Fault impedances"
  Text 8,8,84,8, "Connected branch list"
  Text 8,208,160,12, "List of handle of branches to be outaged:"
  TextBox 8,224,200,12, .EditBox_OList
  ListBox 8,20,200,180, BranchList(), .ListBox_1
  CancelButton 268,224,40,12
  OKButton 220,224,40,12
  CheckBox 220,68,28,8, "3PH", .CheckBox_1
  CheckBox 248,68,28,8, "2LG", .CheckBox_2
  CheckBox 276,68,28,8, "1LG", .CheckBox_3
  CheckBox 304,68,20,8, "LL", .CheckBox_4
  CheckBox 216,84,40,12, "Close-in", .CheckBox_5
  CheckBox 264,84,52,12, "with outage", .CheckBox_6
  Text 256,148,12,12, "+ j"
  TextBox 228,148,24,12, .EditBox_1
  TextBox 268,148,20,12, .EditBox_2
  Text 256,164,12,12, "+ j"
  TextBox 228,164,24,12, .EditBox_3
  TextBox 268,164,20,12, .EditBox_4
  Text 256,180,12,12, "+ j"
  TextBox 228,180,24,12, .EditBox_5
  TextBox 268,180,20,12, .EditBox_6
  Text 292,148,20,12, "Ohm"
  Text 292,164,20,12, "Ohm"
  Text 292,180,20,12, "Ohm"
  ListBox 216,7,96,32, BusList, .ListBox_2
  CheckBox 315,18,8,8, "CheckBox_7 ", .CheckBox_7
  CheckBox 315,26,8,8, "CheckBox_7 ", .CheckBox_8
End Dialog
'=====================================

   Dim Dlg As Dialog_1

   ' Initialize dialog data
   Dlg.CheckBox_1 = 1
   Dlg.CheckBox_3 = 1
   Dlg.CheckBox_5 = 1
   Dlg.CheckBox_6 = 1
   dlg.EditBox_1 = "0.0"
   dlg.EditBox_2 = "0.0"
   dlg.EditBox_3 = "0.0"
   dlg.EditBox_4 = "0.0"
   dlg.EditBox_5 = "0.0"
   dlg.EditBox_6 = "0.0"
   dlg.EditBox_OList = ""

   If ( RlyHnd(1) > 0 ) Then dlg.CheckBox_7 = 1   
   If ( RlyHnd(2) > 0 ) Then dlg.CheckBox_8 = 1   

   ' show the dialog
   Button = Dialog( dlg )
   If Button = 0 Then Exit Sub	' Canceled

   ' Initialize DoFault options using dialog data
   For ii = 1 To 4 
     FltConn(ii) = 0
   Next 
   For ii = 1 To 12
     FltOption(ii) = 0.0
   Next
   For ii = 1 To 3
     OutageType(ii) = 0
   Next
   Rflt       = 0.0   '
   Xflt       = 0.0
   nClearPrev = 0 ' Keep previous result

   ' Fault connection
   FltConn(1)    = dlg.CheckBox_1	' 3PH 
   FltConn(2)    = dlg.CheckBox_2	' 2LG
   FltConn(3)    = dlg.CheckBox_3	' 1LG 
   FltConn(4)    = dlg.CheckBox_4	' 2LL

   ' Fault type
   FltOption(1)  = dlg.CheckBox_5   ' Bus fault
   FltOption(2)  = dlg.CheckBox_6   ' Bus fault with outage

   ' Fault impedance
   FltR(1) = Dlg.EditBox_1
   FltX(1) = Dlg.EditBox_2
   FltR(2) = Dlg.EditBox_3
   FltX(2) = Dlg.EditBox_4
   FltR(3) = Dlg.EditBox_5
   FltX(3) = Dlg.EditBox_6

   OutageType(1) = 1	' Outage one at a time

   If FltOption(2) = 1 Then
     ' Extract handle numbers and prepare the outage list
     SpcPos1   = 1
     DoneFlag  = 0
     HndCounts = 0
     Do While Len( Dlg.EditBox_OList ) > 2 And DoneFlag = 0
       SpcPos2 = InStr( SpcPos1, Dlg.EditBox_OList, " " )
       If SpcPos2 = 0 Then
         SpcPos2  = Len( Dlg.EditBox_OList ) - 1
         DoneFlag = 1
       End If
       WLength = SpcPos2 - SpcPos1
       StrHnd$ = Mid( Dlg.EditBox_OList, SpcPos1, WLength )
       HndCounts = HndCounts + 1
       OutageList(HndCounts) = Val( StrHnd$ )
       SpcPos1 = SpcPos2 + 1
     Loop 

     If HndCounts > 0 Then
        OutageList(HndCounts+1) = 0
     Else
        FltOption(2) = 0
     End If
   End If

   If dlg.CheckBox_7 = 0 Then RlyHnd(1) = 0
   If dlg.CheckBox_8 = 0 Then RlyHnd(2) = 0

   ' Prepare file for output
   FileName = "output.rep"
   Open FileName For Output As #1

   ' Simulate fault on each end
   For jj = 1 To 2
     If RlyHnd(jj) > 0 Then
       ' Simulate the faults three times
       For ii = 1 To 3
         If 0 = DoFault( RlyHnd(jj), FltConn, FltOption, OutageType, OutageList, FltR(ii), _
                       FltX(ii), 0 ) Then GoTo HasError
       Next
     End If
   Next

   ' Print output
   BusID = FullBusName( BusHnd )
   Print #1, "Fault simulation at Bus: ", BusID
   Print #1, ""
   Print #1, "                                  Phase A      Phase B      Phase C"
   Print #1, ""

   ' Must alway pick a fault before getting V and I
   If PickFault( 1 ) = 0 Then GoTo HasError
   Do
     If GetSCCurrent( HND_SC, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
     Print #1, _
        FaultDescription(); Chr(13); _
        "                                     "; _
        Format( MagArray(1), "####0.0"); "@"; Format( AngArray(1), "#0.0"), Space(5), _
        Format( MagArray(2), "####0.0"); "@"; Format( AngArray(2), "#0.0"), Space(5), _
        Format( MagArray(3), "####0.0"); "@"; Format( AngArray(3), "#0.0")
   Loop While PickFault( SF_NEXT ) > 0
   Print "Simulation complete. Report is in " & FileName
   Close
   Exit Sub
HasError:
   Print "Error: ", ErrorString( )
   Close
End Sub
