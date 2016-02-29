' ASPEN PowerScript Sample Program
'
' FLTEXPRT.BAS
'
' This program exports fault simulation results to disk files for relay 
' testing purposes. Four types of files can be exported:
' - Tab-delimited file (.out)
' - Comma-delimited file (.csv)
' - Doble format file (.ss1, .txt)
' - COMTRADE format file (.dat, .cfg, .hdr)
'
' Following simulation results are saved in the files (in secondary quantities):
' - Three-phase faulted voltages and currents
' - Zero and negative-sequence faulted voltages and currents
' - Pre-fault voltages (Doble and COMTRADE only)
'
' A relay group has to be selected before running the script program. The export
' file will contain quantities seen by a relay in the selected group.
' If the selected relay group is on a line that has two relay groups at it two ends,
' the script will also export the result at the remote group to a second set
' of files.
'
' Multiple fault simulations can be saved in a single file except in the COMTRADE 
' format file which can contain only a one fault result.
'
'
Const pi = 3.141592654
Const SQRT2 = 1.41421356
Const MXFLT = 100
' Global variable declaration
Dim FltList(MXFLT) As Long
Dim FltDesc(MXFLT) As String
Dim Time1(MXFLT) As Double, Time2(MXFLT) As Double, Time3(MXFLT) As Double, Time4(MXFLT) As Double
Dim TripOption(MXFLT) As Long, Comments(MXFLT) As String, FName(MXFLT) As String
Dim outVA(MXFLT) As Long, outVB(MXFLT) As Long, outVC(MXFLT) As Long
Dim outIA(MXFLT) As Long, outIB(MXFLT) As Long, outIC(MXFLT) As Long
Dim outV0(MXFLT) As Long, outV2(MXFLT) As Long
Dim outI0(MXFLT) As Long, outI2(MXFLT) As Long
Dim CTratio(3) As Double, PTratio(3) As Double 
Dim Bus1Hnd As Long, Bus2Hnd As Long, Branch1Hnd As Long 
Dim FltCount As Long, SingleFlt As Long
Dim BranchName As String, Bus1Name As String, Bus2Name As String 

Sub main()
   ' Get picked relay group handle
   If GetEquipment( TC_PICKED, PickedHnd& ) = 0 Then 
      Print "Must select a relay group"
      Exit Sub
   End If
   ' Must be a relay group
   If EquipmentType( PickedHnd& ) <> TC_RLYGROUP Then
      Print "Must select a relay group"
      Exit Sub
   End If

   ' There should be some faults simulated
   If PickFault(1) = 0 Then 	
     Print "No fault simulation result available"
     Exit Sub
   End If

   ' Get the relay group branch handle
   If GetData( PickedHnd, RG_nBranchHnd, Branch1Hnd& ) = 0 Then GoTo HasError
   ' Get the branch near bus handle
   If GetData( Branch1Hnd, BR_nBus1Hnd, Bus1Hnd& ) = 0 Then GoTo HasError
   If GetData( Branch1Hnd, BR_nBus2Hnd, Bus2Hnd& ) = 0 Then GoTo HasError	' far bus handle

   If GetData( Branch1Hnd, BR_nType, TypeCode& ) = 0 Then GoTo HasError
   If TypeCode = TC_LINE Then
      ' Get line info
      If GetData( Branch1Hnd, BR_nHandle, LineHnd& ) = 0 Then GoTo HasError
      If GetData( LineHnd, LN_sName, LineName$ ) = 0 Then GoTo HasError
      If GetData( LineHnd, LN_sID, sID$ ) = 0 Then GoTo HasError
      sCode = "L"
      BusHnd& = Bus1Hnd		' temporary near bus handle
      ' Must skip all tap buses on the lines
      Do 
         If GetData( Bus2Hnd, BUS_nTapBus, TapCode& ) = 0 Then GoTo HasError
         If TapCode = 0 Then Exit Do			' real bus
         ' Only for tap bus
         Branch2Hnd& = 0
         ttt = GetBusEquipment( Bus2Hnd, TC_BRANCH, Branch2Hnd& )
         While ttt <> 0
            If GetData( Branch2Hnd, BR_nBus2Hnd, Bus3Hnd& ) = 0 Then GoTo HasError	' Get the far end bus
            If Bus3Hnd <> BusHnd Then	' for different branch
               If GetData( Branch2Hnd, BR_nType, TypeCode& ) = 0 Then GoTo HasError	' Get branch type
               If TypeCode = TC_LINE Then 
                  ' Get line name
                  If GetData( Branch2Hnd, BR_nHandle, LineHnd& ) = 0 Then GoTo HasError
                  If GetData( LineHnd, LN_sName, StringVal$ ) = 0 Then GoTo HasError
                  If StringVal = LineName Then GoTo ExitWhile		' can go further on line with same name
                  ttt = GetBusEquipment( Bus2Hnd, TC_BRANCH, Branch2Hnd )
                  If ttt = -1 Then GoTo ExitWhile		' It is the last line, no choice but further on line
               End If
            Else		' for same branch
               If ttt = -1 Then GoTo ExitLoop		' If the end bus is tap bus, stop
               ttt = GetBusEquipment( Bus2Hnd, TC_BRANCH, Branch2Hnd )
            End If
         Wend
         ExitWhile:
         BusHnd  = Bus2Hnd
         Bus2Hnd = Bus3Hnd	
      Loop While TapCode = 1
   
      ExitLoop:
      ' get the far bus branch handle
      Branch2Hnd = 0
      While GetBusEquipment( Bus2Hnd, TC_BRANCH, Branch2Hnd ) <> 0
         If GetData( Branch2Hnd, BR_nBus2Hnd, Bus3Hnd ) = 0 Then GoTo HasError	' Get the far end bus
         If Bus3Hnd = BusHnd Then Exit Do	' for different branch
      Wend
   elseif TypeCode = TC_XFMR Then
      ' Get circuit ID
      If GetData( Branch1Hnd, BR_nHandle, DevHnd& ) = 0 Then GoTo HasError
      If GetData( DevHnd, XR_sID, sID$ ) = 0 Then GoTo HasError
      sCode = "T"
   elseif TypeCode = TC_XFMR3 Then
      ' Get circuit ID
      If GetData( Branch1Hnd, BR_nHandle, DevHnd& ) = 0 Then GoTo HasError
      If GetData( DevHnd, X3_sID, sID$ ) = 0 Then GoTo HasError
      sCode = "X"
   elseif TypeCode = TC_PS Then
      ' Get circuit ID
      If GetData( Branch1Hnd, BR_nHandle, DevHnd& ) = 0 Then GoTo HasError
      If GetData( DevHnd, PS_sID, sID$ ) = 0 Then GoTo HasError
      sCode = "P"
   End If

   ' Bus information
   Bus1Name$ = FullBusName( Bus1Hnd )	      ' near bus name
   Bus2Name$ = FullBusName( Bus2Hnd )            ' far bus name
   BranchName$ = Bus1Name + "-" + Bus2Name + " " + sID + sCode

   doPage1:
   nCode = PageOne( FltList, FltDesc, FltCount&, SingleFlt& )
   If nCode = 0 Then Exit Sub

   ' Copy location and fault description to use as comment
   For ii& = 1 To Fltcount
     Comments(ii) = "Quantities at " + BranchName + " in:" + Chr(13) + Chr( 10) + FltDesc(ii)
     FName(ii) = CurDir() + "\Fault" + LTrim(Str(ii))
   Next
   Time1(1) = 5
   Time2(1) = 5
   Time3(1) = 0
   Time4(1) = 5
   TripOption(1) = 0
   outVA(1) = 1
   outVB(1) = 1
   outVC(1) = 1
   outV0(1) = 1
   outV2(1) = 1
   outIA(1) = 1
   outIB(1) = 1
   outIC(1) = 1
   outI0(1) = 1
   outI2(1) = 1
   CT# = 100
   PT# = 1000
   If SingleFlt = 0 Then    ' Export selected faults to separate files
     For ii& = 1 To FltCount
       TitleText$ = FltDesc(ii)
       nCode = PageTwo( ii, TitleText$ )
       If nCode = 0 Then Exit Sub
       If nCode = 2 Then 
         If ii = 1 Then GoTo doPage1 Else ii = ii - 2
       Else
         ' Use as default
         Time1(ii+1) = Time1(ii)
         Time2(ii+1) = Time2(ii)
         Time3(ii+1) = Time3(ii)
         Time4(ii+1) = Time4(ii)
         TripOption(ii+1) = TripOption(ii)
         outVA(ii+1) = outVA(ii)
         outVB(ii+1) = outVB(ii)
         outVC(ii+1) = outVC(ii)
         outV0(ii+1) = outV0(ii)
         outV2(ii+1) = outV2(ii)
         outIA(ii+1) = outIA(ii)
         outIB(ii+1) = outIB(ii)
         outIC(ii+1) = outIC(ii)
         outI0(ii+1) = outI0(ii)
         outI2(ii+1) = outI2(ii)
       End If
     Next
     TitleText = "Export Relay Quantities at " + BranchName
     nCode = PageThree( CT#, PT#, Infor$, FileName$, More&, TitleText )
     If nCode = 0 Then Exit Sub
     If nCode = 2 Then GoTo doPage1
     CTratio(1) = CT
     PTratio(1) = PT
     ' Now do the export
     For ii = 1 To Fltcount
       Call exportCOMTRADE(ii) 
     Next
     Print "Export completed successfully. Data files are in: " + CurDir()
   Else
     ' Export all selected faults to a single file
     doPage2_1:
     For ii& = 1 To FltCount
       TitleText$ = FltDesc(ii)
       nCode = PageTwo_1( ii, TitleText )
       If nCode = 0 Then Exit Sub
       If nCode = 2 Then 
         If ii = 1 Then GoTo doPage1 Else ii = ii - 2
       Else
         ' Use as default
         Time2(ii+1) = Time2(ii)
       End If
     Next
     TitleText = "Evolving Fault of" + Str( fltCount ) + " Steps"
     nCode = PageTwo_2( TitleText$ )
     If nCode = 0 Then Exit Sub
     If nCode = 2 Then GoTo doPage2_1
     ' Compose comment string
     Infor$ = "Quantities at " + BranchName + " in evolving fault:" + Chr(13) + Chr(10) + _
                   FltDesc(1) + ", " + Str(Time2(1)) + " cycles"
     For ii = 2 To FltCount
       Infor$ = Infor$ + Chr(13) + Chr(10) + FltDesc(ii) + ", " + Str(Time2(ii)) + " cycles"
     Next
     FileName$ = CurDir() + "\FaultData"
     Infor$ = Comments(1)
     TitleText = "Export Relay Quantities at " + BranchName
     nCode = PageThree( CT#, PT#, Infor$, FileName$, More&, TitleText )
     If nCode = 0 Then Exit Sub
     If nCode = 2 Then GoTo doPage1
     CTratio(1) = CT
     PTratio(1) = PT
     FName(1)   = FileName
     Comments(1) = Infor
     ' Now do the export
     Call exportCOMTRADE(-1)
     Print "Export completed successfully."
   End If

 
   Exit Sub      
HasError:
   Print "Error: ", ErrorString( )
   Close
End Sub
'===================================End of Main()====================================

'====================================Dialog Spec=====================================
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
'====================Dialog box spec (generated by Dialog Editor)===================
Const WSTYLE1 = WS_VSCROLL Or ES_MULTILINE Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL
Begin Dialog PAGE1 15,53, 325, 197, "Select Fault Result to Export"
  Text 4,4,264,12, "Simulation results for following faults are available. Edit the list  to keep only"
  Text 4,12,280,12, "the faults that you want to export. You must not change any fault index number."
  PushButton 176,180,44,12, "Next", .Next
  TextBox 4,28,308,84, .EditBox_1, WSTYLE1
  CheckBox 4,112,196,12, "Export as a single evolving fault", .CheckBox_1
  GroupBox 4,136,308,28, " Export Quantities "
  CheckBox 16,148,25,12, "Va", .CheckBox_2
  CheckBox 44,148,25,12, "Vb", .CheckBox_3
  CheckBox 72,148,25,12, "Vc", .CheckBox_4
  CheckBox 100,148,25,12, "V0", .CheckBox_5
  CheckBox 132,148,25,12, "V2", .CheckBox_6
  CheckBox 160,148,25,12, "Ia", .CheckBox_7
  CheckBox 184,148,25,12, "Ib", .CheckBox_8
  CheckBox 208,148,25,12, "Ic", .CheckBox_9
  CheckBox 236,148,25,12, "I0", .CheckBox_10
  CheckBox 260,148,25,12, "I2", .CheckBox_11
  Text 4,164,312,8, "_____________________________________________________________________________"
  CancelButton 236,180,36,12
End Dialog
'====================================PageOne()=========================================
' Purpose:
'   Solicit user input on the scope of the export
'
'======================================================================================
Function PageOne( ByRef FltList() As Long, ByRef FltDesc() As String, _
                  ByRef FltCount As Long, ByRef SingleFlt ) As Long
  Dim dlg As PAGE1

  ' Prepare fault list 
  AString$ = ""
  If PickFault( 1 ) = 0 Then Exit Sub	' No fault
  Do
    FltString$ = FaultDescription()
    ' Need to insert chr(13) at the end of each line to make it
    ' show up properly in the edit box
    CharPos = InStr( 1, FltString, Chr(10) )
    While CharPos > 0
      TempStr$   = Left$( FltString, CharPos - 1 )
      If Right( TempStr, 3 ) <> "on:" Then TempStr = TempStr + Chr(13) + Chr(10)
      TempStr    = TempStr + " " + LTrim( Mid$(FltString, CharPos+1, 9999 ) )
      FltString$ = TempStr$
      CharPos    = InStr( CharPos+2, FltString, Chr(10) )
    Wend
    AString$ = AString$ + TempStr$ + Chr(13) + Chr(10) 
  Loop While PickFault( SF_NEXT ) > 0

  ' Initialize dialog box
  dlg.EditBox_1 = AString$

  PageOne = Dialog(dlg)	' 2 = Previous; 1 = Next; 0 = Cancel;

  If PageOne = 0 Then Exit Function

  ' Parse FaultString to get the list of fault number to output
  FltCount       = 0
  AString$       = dlg.EditBox_1
  CharPos&       = InStr( 1, AString$, Chr(10) )
  While CharPos > 0
    ALine$    = Left$( AString$, CharPos - 2 )
    CharPos1& = InStr( 1, ALine$, ". " )
    If CharPos1 > 0 And CharPos1 < 10 Then
      TempStr$       = Left$( ALine$, CharPos1 - 1 )
      FltCount          = FltCount + 1
      FltList(FltCount) = Val( TempStr$ )
      FltDesc(FltCount) = ALine
    End If
    AString$ = Mid$(AString$, CharPos+1, 9999 )
    CharPos  = InStr( 1, AString$, Chr(10) )
  Wend
  SingleFlt = dlg.CheckBox_1
  If FltCount = 0 Then 
    PageOne   = 0
    Exit Function   ' Nothing to report. Cancel
  End If
End Function 'Page1

'====================Dialog box spec (generated by Dialog Editor)===================
Const WSTYLE2 = ES_READONLY Or ES_CENTER
Begin Dialog PAGE2 15,53, 325, 197, "Fault Export Options"
  OptionGroup .GROUP_1
    OptionButton 144,32,52,12, "Three-pole"
    OptionButton 208,32,52,12, "Single pole"
  GroupBox 12,20,104,80, " Durations in 60Hz cycles "
  GroupBox 120,20,160,28, " Fault tripping option"
  GroupBox 120,52,160,48, " Export quantities "
  Text 20,36,36,12, "Prefault"
  Text 60,36,9,12, "="
  TextBox 68,36,28,12, .EditBox_2
  Text 20,52,36,12, "Fault"
  Text 60,52,9,12, "="
  TextBox 68,52,28,12, .EditBox_3
  Text 20,68,36,12, "Reclosing"
  Text 60,68,9,12, "="
  TextBox 68,68,28,12, .EditBox_4
  Text 20,84,36,12, "Post-Fault"
  Text 60,84,9,12, "="
  TextBox 68,84,28,12, .EditBox_5
  CheckBox 136,64,21,12, "Va", .CheckBox_1
  CheckBox 160,64,21,12, "Vb", .CheckBox_2
  CheckBox 184,64,21,12, "Vc", .CheckBox_3
  CheckBox 208,64,21,12, "V0", .CheckBox_4
  CheckBox 236,64,21,12, "V2", .CheckBox_5
  CheckBox 136,80,21,12, "Ia", .CheckBox_6
  CheckBox 160,80,21,12, "Ib", .CheckBox_7
  CheckBox 184,80,21,12, "Ic", .CheckBox_8
  CheckBox 208,80,21,12, "I0", .CheckBox_9
  CheckBox 236,80,21,12, "I2", .CheckBox_10
  Text 12,108,264,8, "Comments in COMTRADE header (Use Ctrl-Enter to start a new line):"
  TextBox 12,120,300,24, .EditBox_6, WSTYLE1
  Text 12,152,84,12, "COMTRADE file name ="
  TextBox 96,152,216,12, .EditBox_7
  Text 4,164,312,8, "_____________________________________________________________________________"
  PushButton 176,180,44,12, "Next", .Next
  PushButton 124,180,44,12, "Back", .Back
  CancelButton 236,180,36,12
  TextBox 8,4,308,12, .EditBox_1, WSTYLE2
End Dialog
'====================================PageTwo()=========================================
' Purpose:
'   Solicit user input on the scope of the export
'
'======================================================================================
Function PageTwo( ByVal ii, ByRef TitleText As String ) As Long
  Dim dlg As PAGE2

  dlg.EditBox_1 = TitleText
  dlg.EditBox_6 = Comments(ii)
  dlg.EditBox_7 = FName(ii)
  dlg.EditBox_2 = Time1(ii)
  dlg.EditBox_3 = Time2(ii)
  dlg.EditBox_4 = Time3(ii)
  dlg.EditBox_5 = Time4(ii)
  dlg.Group_1   = TripOption(ii)
  dlg.CheckBox_1  = outVA(ii)
  dlg.CheckBox_2  = outVB(ii)
  dlg.CheckBox_3  = outVC(ii)
  dlg.CheckBox_4  = outV0(ii)
  dlg.CheckBox_5  = outV2(ii)
  dlg.CheckBox_6  = outIA(ii)
  dlg.CheckBox_7  = outIB(ii)
  dlg.CheckBox_8  = outIC(ii)
  dlg.CheckBox_9  = outI0(ii)
  dlg.CheckBox_10 = outI2(ii)
  DoneLoop = 0
  Do
    PageTwo = Dialog( dlg )
    If PageTwo <> 1 Then Exit Function
    If Val(dlg.EditBox_2) < 0 Then
      Print "Negative prefault duration"
    ElseIf Val(dlg.EditBox_3) <= 0 Then
      Print "Fault duration not positive"
    ElseIf Val(dlg.EditBox_4) < 0 Then
      Print "Negative reclosing duration"
    ElseIf Val(dlg.EditBox_5) < 0 Then
      Print "Negative post-fault duration"
    elseif dlg.EditBox_7 = "" Then
      Print "Please enter file name"
    Else
      Time1(ii) = Val(dlg.EditBox_2)
      Time2(ii) = Val(dlg.EditBox_3)
      Time3(ii) = Val(dlg.EditBox_4)
      Time4(ii) = Val(dlg.EditBox_5)
      TripOption(ii) = dlg.Group_1
      outVA(ii) = dlg.CheckBox_1
      outVB(ii) = dlg.CheckBox_2
      outVC(ii) = dlg.CheckBox_3
      outV0(ii) = dlg.CheckBox_4
      outV2(ii) = dlg.CheckBox_5
      outIA(ii) = dlg.CheckBox_6
      outIB(ii) = dlg.CheckBox_7
      outIC(ii) = dlg.CheckBox_8
      outI0(ii) = dlg.CheckBox_9
      outI2(ii) = dlg.CheckBox_10
      Comments(ii)   = dlg.EditBox_6
      FName(ii) = dlg.EditBox_7
      DoneLoop = 1
    End If
  Loop While DoneLoop = 0
End Function

'====================Dialog box spec (generated by Dialog Editor)===================
Begin Dialog PAGE2_1 15,53, 325, 197, "Fault Export Options"
  GroupBox 20,52,104,68, " Durations in 60Hz cycles "
  Text 28,72,36,12, "Fault"
  Text 68,72,9,12, "="
  TextBox 76,72,28,12, .EditBox_3
  Text 28,92,36,12, "Reclosing"
  Text 68,92,9,12, "="
  TextBox 76,92,28,12, .EditBox_4
  GroupBox 128,52,160,68, " Fault tripping option"
  OptionGroup .GROUP_1
    OptionButton 152,80,52,12, "Three-pole"
    OptionButton 216,80,52,12, "Single pole"
  Text 4,164,312,8, "_____________________________________________________________________________"
  PushButton 176,180,44,12, "Next", .Next
  PushButton 124,180,44,12, "Back", .Back
  CancelButton 236,180,36,12
  TextBox 8,4,308,12, .EditBox_1,WSTYLE2
End Dialog
'====================================PageTwo_1()=========================================
' Purpose:
'   Solicit user input on the scope of the export
'
'======================================================================================
Function PageTwo_1( ByVal ii, ByRef TitleText As String ) As Long
  Dim dlg As PAGE2_1

  dlg.EditBox_1 = TitleText
  dlg.EditBox_3 = Time2(ii)
  dlg.EditBox_4 = Time3(ii)
  dlg.Group_1   = TripOption(ii)
  DoneLoop = 0
  Do
    PageTwo_1 = Dialog( dlg )
    If PageTwo_1 <> 1 Then Exit Function
    If Val(dlg.EditBox_3) <= 0 Then
      Print "Fault duration not positive"
    ElseIf Val(dlg.EditBox_4) < 0 Then
      Print "Negative reclosing duration"
    Else
      Time2(ii) = Val(dlg.EditBox_3)
      Time3(ii) = Val(dlg.EditBox_4)
      DoneLoop = 1
    End If
  Loop While DoneLoop = 0
End Function
'====================Dialog box spec (generated by Dialog Editor)===================
Begin Dialog PAGE2_2 15,53, 325, 197, "Fault Export Options"
  GroupBox 32,44,104,68, " Durations in 60Hz cycles "
  GroupBox 140,44,144,68, " Export quantities "
  Text 40,68,36,12, "Prefault"
  Text 80,68,9,12, "="
  TextBox 88,68,28,12, .EditBox_2
  Text 40,88,36,12, "Post-Fault"
  Text 80,88,9,12, "="
  TextBox 88,88,28,12, .EditBox_5
  CheckBox 156,64,21,12, "Va", .CheckBox_1
  CheckBox 180,64,21,12, "Vb", .CheckBox_2
  CheckBox 204,64,21,12, "Vc", .CheckBox_3
  CheckBox 228,64,21,12, "V0", .CheckBox_4
  CheckBox 256,64,21,12, "V2", .CheckBox_5
  CheckBox 156,80,21,12, "Ia", .CheckBox_6
  CheckBox 180,80,21,12, "Ib", .CheckBox_7
  CheckBox 204,80,21,12, "Ic", .CheckBox_8
  CheckBox 228,80,21,12, "I0", .CheckBox_9
  CheckBox 256,80,21,12, "I2", .CheckBox_10
  Text 4,164,308,8, "_____________________________________________________________________________"
  PushButton 176,180,44,12, "Next", .Next
  PushButton 124,180,44,12, "Back", .Back
  CancelButton 236,180,36,12
  TextBox 8,4,304,12, .EditBox_1, WSTYLE2
End Dialog
'====================================PageTwo_2()=========================================
' Purpose:
'   Solicit user input on the scope of the export
'
'======================================================================================
Function PageTwo_2( ByRef TitleText As String ) As Long
  Dim dlg As PAGE2_2

  dlg.EditBox_1 = TitleText
  dlg.EditBox_2 = Time1(1)
  dlg.EditBox_5 = Time4(1)
  dlg.CheckBox_1  = outVA(1)
  dlg.CheckBox_2  = outVB(1)
  dlg.CheckBox_3  = outVC(1)
  dlg.CheckBox_4  = outV0(1)
  dlg.CheckBox_5  = outV2(1)
  dlg.CheckBox_6  = outIA(1)
  dlg.CheckBox_7  = outIB(1)
  dlg.CheckBox_8  = outIC(1)
  dlg.CheckBox_9  = outI0(1)
  dlg.CheckBox_10 = outI2(1)
  DoneLoop = 0
  Do
    PageTwo_2 = Dialog( dlg )
    If PageTwo_2 <> 1 Then Exit Function
    If Val(dlg.EditBox_2) < 0 Then
      Print "Negative prefault duration"
    ElseIf Val(dlg.EditBox_5) < 0 Then
      Print "Negative post-fault duration"
    Else
      Time1(1) = Val(dlg.EditBox_2)
      Time4(1) = Val(dlg.EditBox_5)
      outVA(1) = dlg.CheckBox_1
      outVB(1) = dlg.CheckBox_2
      outVC(1) = dlg.CheckBox_3
      outV0(1) = dlg.CheckBox_4
      outV2(1) = dlg.CheckBox_5
      outIA(1) = dlg.CheckBox_6
      outIB(1) = dlg.CheckBox_7
      outIC(1) = dlg.CheckBox_8
      outI0(1) = dlg.CheckBox_9
      outI2(1) = dlg.CheckBox_10
      DoneLoop = 1
    End If
  Loop While DoneLoop = 0
End Function

'====================Dialog box spec (generated by Dialog Editor)===================
Begin Dialog PAGE3 15,53, 325, 197, "Relay location"
  Text 64,32,36,12, "CT ratio"
  Text 104,32,9,12, "="
  TextBox 112,32,28,12, .EditBox_2
  Text 152,32,36,12, "PT ratio"
  Text 192,32,9,12, "="
  TextBox 200,32,28,12, .EditBox_3
  Text 8,52,117,12, "COMTRADE header comments:"
  TextBox 10,64,300,48, .EditBox_4, WSTYLE1
  Text 8,132,37,12, "File name:"
  TextBox 48,132,264,12, .EditBox_5
  CheckBox 152,152,149,12, "I want to enter more relay location", .CheckBox_1
  Text 4,164,312,8, "_____________________________________________________________________________"
  PushButton 176,180,44,12, "Next", .Next
  PushButton 124,180,44,12, "Back", .Back
  CancelButton 236,180,36,12
  TextBox 8,4,308,12, .EditBox_1, WSTYLE2
  Text 188,112,121,12, "(Use Ctrl-Enter to start a new line)"
End Dialog
'===================================== exportComtrade()================================
' Purpose:
'   Get user input
'
'======================================================================================
Function PageThree( ByRef CT As Double, ByRef PT As Double, ByRef Infor As String, _
                    ByRef FileName As String, ByRef More As Long, ByRef TitleText As String ) As Long
    Dim dlg As PAGE3

  dlg.EditBox_1 = TitleText
  dlg.EditBox_2 = CT
  dlg.EditBox_3 = PT
  dlg.EditBox_4 = Infor
  dlg.EditBox_5 = FileName
  DoneLoop = 0
  Do
    PageThree = Dialog( dlg )
    If PageThree <> 1 Then Exit Function
    If Val(dlg.EditBox_2) <= 0 Then
      Print "CT ratio not positive"
    ElseIf Val(dlg.EditBox_3) <= 0 Then
      Print "PT ratio not positive"
    elseif dlg.EditBox_5 = "" Then
      Print "Please enter file name"
    Else
      CT = dlg.EditBox_2
      PT = dlg.EditBox_3
      Infor = dlg.EditBox_4
      FileNmae = dlg.EditBox_5
      More = dlg.CheckBox_1
      DoneLoop = 1
    End If
  Loop While DoneLoop = 0
End Function
'===================================== exportComtrade()================================
' Purpose:
'   Print out fault result to Comtrade file (*.CFG, *.DAT)
'
'======================================================================================
Function exportCOMTRADE( ByVal nCase As Long ) As Long

   Dim MagArray1(3) As Double, AngArray1(3) As Double
   Dim MagArray(16) As Double, AngArray(16) As Double
   Dim MagArray2(10) As Double, AngArray2(10) As Double
   Dim DataValue(10) As Long
   Dim BusNominal As Double
   Dim aV As Double, bV As Double, aI As Double, bI As Double
   Dim PtCount As Long, Point As Long, Omeg As Double, Sample As Long, PtNo As Long

   exportCOMTRADE = 0
 
   If nCase = -1 Then
     nCase  = 1
     nFault = 1
     bSingleFile& = 1
   Else
     bSingleFile& = 0
   End If   
   If GetData( Bus1Hnd, BUS_nNumber, BusNumber& ) = 0 Then
      Print "Get Bus number failed."
      Exit Function
   End If
   If GetData( Bus1Hnd, BUS_dKVnorminal, BusNominal ) = 0 Then
      Print "Get Bus nominal voltage failed."
      Exit Function
   End If

   ' Select the first fault and export data
   If PickFault( FltList(nCase) ) = 0 Then 
      Print "Problem reading fault simulation result #", FltList(nCase)
      Exit Function
   End If

   Open FName(nCase) + ".CFG" For Output As 1   ' CFG file
   Open FName(nCase) + ".HDR" For Output As 2   ' HDR file
   Open FName(nCase) + ".DAT" For Output As 3   ' DAT file

   ' Calculate conversion factors a & b for V & I
   ' The maximum voltage is 1.3*Vnominal, the maximum current is 25*Irate
   ' Assume 12-bit AD converter (+2047, -2048)
   ' PtNo samples per cycle
   aV   = SQRT2*1.3*( BusNominal*1000/PTratio(1) )/( 2047+2048 )
   bV   = 0.0
   aI   = SQRT2*25*5/( 2047+2048 )
   bI   = 0.0
   PtNo = 100
   Sample = 60.0*PtNo

   ' Count number of out series
   nCountS& = 0
   If outVA(nCase) = 1 Then nCount = nCount + 1
   If outVB(nCase) = 1 Then nCount = nCount + 1
   If outVC(nCase) = 1 Then nCount = nCount + 1
   If outIA(nCase) = 1 Then nCount = nCount + 1
   If outIB(nCase) = 1 Then nCount = nCount + 1
   If outIC(nCase) = 1 Then nCount = nCount + 1
   If outV0(nCase) = 1 Then nCount = nCount + 1
   If outV2(nCase) = 1 Then nCount = nCount + 1
   If outI0(nCase) = 1 Then nCount = nCount + 1
   If outI2(nCase) = 1 Then nCount = nCount + 1

   sCount$ = Trim(Str(nCount)) + "," + Trim(Str(nCount)) + "A,0D"

   ' Print out to *.CFG
   Print #1, Bus1Name & "," & BusNumber & ",1999"
   Print #1, sCount$
   If outVA(nCase) Then _
     Print #1, "1," & Bus1Name & " Va-g,,,V," & Format( aV, "#####0.0000000000" ) & "," & _
      Format(bV, "#####0.0000000000") & ",0,-2048,2047," & Format( PTratio(1), "#####0.000" ) & ",1,S"
   If outVB(nCase) Then _
     Print #(FileNo+1), "2," & Bus1Name & " Vb-g,,,V," & Format( aV, "#####0.0000000000" ) & "," & _
      Format(bV, "#####0.0000000000") & ",0,-2048,2047," & Format( PTratio(1), "#####0.000" ) & ",1,S"
   If outVC(nCase) Then _
     Print #(FileNo+1), "3," & Bus1Name & " Vc-g,,,V," & Format( aV, "#####0.0000000000" ) & "," & _
      Format(bV, "#####0.0000000000") & ",0,-2048,2047," & Format( PTratio(1), "#####0.000" ) & ",1,S"
   If outIA(nCase) Then _
     Print #(FileNo+1), "4," & BranchName & " Ia,,,A," & Format( aI, "#####0.0000000000" ) & "," & _
      Format(bI, "#####0.0000000000") & ",0,-2048,2047," & Format( CTratio(1), "#####0.000" ) & ",1,S"
   If outIB(nCase) Then _
     Print #(FileNo+1), "5," & BranchName & " Ib,,,A," & Format( aI, "#####0.0000000000" ) & "," & _
      Format(bI, "#####0.0000000000") & ",0,-2048,2047," & Format( CTratio(1), "#####0.000" ) & ",1,S"
   If outIC(nCase) Then _
     Print #(FileNo+1), "6," & BranchName & " Ic,,,A," & Format( aI, "#####0.0000000000" ) & "," & _
      Format(bI, "#####0.0000000000") & ",0,-2048,2047," & Format( CTratio(1), "#####0.000" ) & ",1,S"
   If outV0(nCase) Then _
     Print #(FileNo+1), "7," & Bus1Name & " V0,,,V," & Format( aV, "#####0.0000000000" ) & "," & _
      Format(bV, "#####0.0000000000") & ",0,-2048,2047," & Format( PTratio(1), "#####0.000" ) & ",1,S"
   If outI0(nCase) Then _
     Print #(FileNo+1), "8," & BranchName & " I0,,,A," & Format( aI, "#####0.0000000000" ) & "," & _
      Format(bI, "#####0.0000000000") & ",0,-2048,2047," & Format( CTratio(1), "#####0.000" ) & ",1,S"
   If outV2(nCase) Then _
     Print #(FileNo+1), "9," & Bus1Name & " V2,,,V," & Format( aV, "#####0.0000000000" ) & "," & _
      Format(bV, "#####0.0000000000") & ",0,-2048,2047," & Format( PTratio(1), "#####0.000" ) & ",1,S"
   If outI2(nCase) Then _
     Print #(FileNo+1), "10," & BranchName & " I2,,,A," & Format( aI, "#####0.0000000000" ) & "," & _
      Format(bI, "#####0.0000000000") & ",0,-2048,2047," & Format( CTratio(1), "#####0.000" ) & ",1,S"
   Print #1, "60"
   Print #1, "1"
   ' Total cycles
   If bSingleFile Then
     nTotalCycles& = Time1(1) + Time4(1)
     For ii = 1 To FltCount
       nTotalCycles = nTotalCycltes + Time2(ii)+Time3(ii)
     Next
   Else
     nTotalCycles& = Time1(nCase)+Time2(nCase)+Time3(nCase)+Time4(nCase)
   End If
   Print #1, Format( Sample, "##0000.000" ) & "," & _
      Format( PtNo*nTotalCycles, "0" )
   
   CurrentDate$ = Date()
   CurrentTime$ = Time()
   Print #1, Format( Day(CurrentDate), "00" ) & "/" & Format( Month(CurrentDate), "00" ) & _
      "/" & Format( Year(CurrentDate), "0000" ) & "," & Format( Hour(CurrentTime), "00" ) & ":" & _
      Format( Minute(CurrentTime), "00" ) & ":" & Format( Second(CurrentTime), "00" ) & ".000000"

   ' Non-critical, just assign a value
   Print #1, Format( Day(CurrentDate), "00" ) & "/" & Format( Month(CurrentDate), "00" ) & _
      "/" & Format( Year(CurrentDate), "0000" ) & "," & Format( Hour(CurrentTime), "00" ) & ":" & _
      Format( Minute(CurrentTime), "00" ) & ":" & Format( Second(CurrentTime), "00" ) & ".100000"
   Print #1, "ASCII"
   Print #1, "1"	' timemult

   ' Print to *.HDR file
   Print #2, "*** Fault simulation result from ASPEN OneLiner ***"
   Print #2, "* Date/Time: " & CurrentDate & " " & CurrentTime
   Print #2, "* Description: "
   Print #2, Comments(nCase)
   Print #2, "*"
   Print #2, "***************************************************"

   ' Print *.DAT file
   ' Prefault quantities
   If GetPSCVoltage( Bus1Hnd, MagArray1, AngArray1, 1 ) = 0 Then	' get prefault voltage on bus
      Print "Get prefault voltage failed."
      Exit Function
   End If
   MagArray1(1) = SQRT2*MagArray1(1)*1000/PTratio(1)
   AngArray1(1) = AngArray1(1)*pi/180
   AngArray1(2) = AngArray1(1) - 2*pi/3	' phase B
   If AngArray1(2) <=-2*pi Then AngArray1(2) = AngArray1(2) + 2*pi
   AngArray1(3) = AngArray1(1) + 2*pi/3	' phase C
   If AngArray1(3) >=2*pi  Then AngArray1(3) = AngArray1(3) - 2*pi
   
   PtCount = 1
   Omeg    = 120*pi/Sample	' PtNo pts per cycle 

   ' Pre-fault period
   Point = Int( PtNo*Time1(nCase) )	' get integer part of prefault points
   While PtCount <= Point
      ' Va, Vb, Vc
      DataValue(1) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(1) )-bV )/aV )
      DataValue(2) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(2) )-bV )/aV )
      DataValue(3) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(3) )-bV )/aV )
      ' Ia, Ib, Ic, V0, I0, V2, I2
      DataValue(4) = 0
      DataValue(5) = 0
      DataValue(6) = 0
      DataValue(7) = 0
      DataValue(8) = 0
      DataValue(9) = 0
      DataValue(10)= 0

      strText$ = Format( PtCount, "0" ) & "," & Format( 1e6*(PtCount-1)/Sample,"#######0" ) 
      If outVA(nCase) Then strText = strText & "," & Format( DataValue(1),  "####0" )
      If outVB(nCase) Then strText = strText & "," & Format( DataValue(2),  "####0" )
      If outVC(nCase) Then strText = strText & "," & Format( DataValue(3),  "####0" )
      If outIA(nCase) Then strText = strText & "," & Format( DataValue(4),  "####0" )
      If outIB(nCase) Then strText = strText & "," & Format( DataValue(5),  "####0" )
      If outIC(nCase) Then strText = strText & "," & Format( DataValue(6),  "####0" )
      If outV0(nCase) Then strText = strText & "," & Format( DataValue(7),  "####0" )
      If outI0(nCase) Then strText = strText & "," & Format( DataValue(8),  "####0" )
      If outV2(nCase) Then strText = strText & "," & Format( DataValue(9),  "####0" )
      If outI2(nCase) Then strText = strText & "," & Format( DataValue(10), "####0" )
      Print #3, strText$
      PtCount = PtCount+1
   Wend
   
   ' Fault period
   Do
     ' Bus voltage
     If GetSCVoltage( Bus1Hnd, MagArray, AngArray, 4 ) = 0 Then 
       Print "Get Bus voltage failed."
       Exit Function
     End If
     For jj = 1 To 3
       MagArray2(jj) = SQRT2*MagArray(jj)*1000/PTratio(1)
       AngArray2(jj) = AngArray(jj)*pi/180
     Next jj
     ' Bbranch current
     If GetSCCurrent( Branch1Hnd, MagArray, AngArray, 4 ) = 0 Then 
       Print "Get Branch current failed."
       Exit Function
     End If
     For jj = 1 To 3
       MagArray2(jj+3) = SQRT2*MagArray(jj)/CTratio(1)
       AngArray2(jj+3) = AngArray(jj)*pi/180
     Next jj

     ' Bus voltage in sequence
     If GetSCVoltage( Bus1Hnd, MagArray, AngArray, 2 ) = 0 Then 
       Print "Get Bus sequence voltage failed."
       Exit Function
     End If
     MagArray2(7) = SQRT2*MagArray(1)*1000/PTratio(1)	' V0
     AngArray2(7) = AngArray(1)*pi/180			' V0
     MagArray2(9) = SQRT2*MagArray(3)*1000/PTratio(1)	' V2
     AngArray2(9) = AngArray(3)*pi/180			' V2
     ' Branch current in sequence
     If GetSCCurrent( Branch1Hnd, MagArray, AngArray, 2 ) = 0 Then 
       Print "Get Branch sequence current failed."
       Exit Function
     End If
     MagArray2(8)  = SQRT2*MagArray(1)/CTratio(1)		' I0
     AngArray2(8)  = AngArray(1)*pi/180			' I0
     MagArray2(10) = SQRT2*MagArray(3)/CTratio(1)		' I2
     AngArray2(10) = AngArray(3)*pi/180			' I2
  
     If bSingleFile = 1 Then Point = Point + Int(PtNo*Time2(nFault)) Else Point = Point + Int(PtNo*Time2(nCase))
     While PtCount <= Point
       ' Va, Vb, Vc
       DataValue(1) = Int( ( MagArray2(1) * Sin( Omeg*(PtCount-1)+AngArray2(1) )-bV )/aV )
       DataValue(2) = Int( ( MagArray2(2) * Sin( Omeg*(PtCount-1)+AngArray2(2) )-bV )/aV )
       DataValue(3) = Int( ( MagArray2(3) * Sin( Omeg*(PtCount-1)+AngArray2(3) )-bV )/aV )
       ' Ia, Ib, Ic
       DataValue(4) = Int( ( MagArray2(4) * Sin( Omeg*(PtCount-1)+AngArray2(4) )-bI )/aI )
       DataValue(5) = Int( ( MagArray2(5) * Sin( Omeg*(PtCount-1)+AngArray2(5) )-bI )/aI )
       DataValue(6) = Int( ( MagArray2(6) * Sin( Omeg*(PtCount-1)+AngArray2(6) )-bI )/aI )
       ' V0, I0, V2, I2
       DataValue(7) = Int( ( MagArray2(7) * Sin( Omeg*(PtCount-1)+AngArray2(7) )-bV )/aV )
       DataValue(8) = Int( ( MagArray2(8) * Sin( Omeg*(PtCount-1)+AngArray2(8) )-bI )/aI )
       DataValue(9) = Int( ( MagArray2(9) * Sin( Omeg*(PtCount-1)+AngArray2(9) )-bV )/aV )
       DataValue(10)= Int( ( MagArray2(10)* Sin( Omeg*(PtCount-1)+AngArray2(10) )-bI )/aI )

       strText$ = Format( PtCount, "0" ) & "," & Format( 1e6*(PtCount-1)/Sample,"#######0" ) 
       If outVA(nCase) Then strText = strText & "," & Format( DataValue(1),  "####0" )
       If outVB(nCase) Then strText = strText & "," & Format( DataValue(2),  "####0" )
       If outVC(nCase) Then strText = strText & "," & Format( DataValue(3),  "####0" )
       If outIA(nCase) Then strText = strText & "," & Format( DataValue(4),  "####0" )
       If outIB(nCase) Then strText = strText & "," & Format( DataValue(5),  "####0" )
       If outIC(nCase) Then strText = strText & "," & Format( DataValue(6),  "####0" )
       If outV0(nCase) Then strText = strText & "," & Format( DataValue(7),  "####0" )
       If outI0(nCase) Then strText = strText & "," & Format( DataValue(8),  "####0" )
       If outV2(nCase) Then strText = strText & "," & Format( DataValue(9),  "####0" )
       If outI2(nCase) Then strText = strText & "," & Format( DataValue(10), "####0" )
       Print #3, strText$
       PtCount = PtCount+1
     Wend

     ' Reclosing
     If bSingleFile = 1 Then 
       nCycle3 =  Time3(nFault)
       nTrip   = TripOption(nFault)
     Else
       nCycle3 = Time3(nCase)
       nTrip   = TripOption(nCase)
     End If
     If nCycle3 > 0.0 Then
       Point = Point + Int( PtNo*nCycle3 )	' get integer part of reclosing points
      ' Determine if this is a ground fault
       FaultString$ = FaultDescription()
       n1LG& = InStr( 1, FaultString, "1LG" )
       While PtCount <= Point
         ' Va, Vb, Vc
         DataValue(1) = 0.0
         If nTrip > 0 And n1LG > 0 Then
           DataValue(2) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(2) )-bV )/aV )
           DataValue(3) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(3) )-bV )/aV )
         Else
           DataValue(2) = 0.0
           DataValue(3) = 0.0
         End If

         ' Ia, Ib, Ic, V0, I0, V2, I2
         DataValue(4) = 0.0
         DataValue(5) = 0.0
         DataValue(6) = 0.0
         DataValue(7) = 0.0
         DataValue(8) = 0.0
         DataValue(9) = 0.0
         DataValue(10)= 0.0

         strText$ = Format( PtCount, "0" ) & "," & Format( 1e6*(PtCount-1)/Sample,"#######0" ) 
         If outVA(nCase) Then strText = strText & "," & Format( DataValue(1),  "####0" )
         If outVB(nCase) Then strText = strText & "," & Format( DataValue(2),  "####0" )
         If outVC(nCase) Then strText = strText & "," & Format( DataValue(3),  "####0" )
         If outIA(nCase) Then strText = strText & "," & Format( DataValue(4),  "####0" )
         If outIB(nCase) Then strText = strText & "," & Format( DataValue(5),  "####0" )
         If outIC(nCase) Then strText = strText & "," & Format( DataValue(6),  "####0" )
         If outV0(nCase) Then strText = strText & "," & Format( DataValue(7),  "####0" )
         If outI0(nCase) Then strText = strText & "," & Format( DataValue(8),  "####0" )
         If outV2(nCase) Then strText = strText & "," & Format( DataValue(9),  "####0" )
         If outI2(nCase) Then strText = strText & "," & Format( DataValue(10), "####0" )
         Print #3, strText$
         PtCount = PtCount+1
       Wend
     End If
     If bSingleFile = 1 Then 
       ' Retrieve next fault
       If nFault = FltCount Then Exit Do
       nFault = nFault + 1
       If PickFault( FltList(nFault) ) = 0 Then Exit Do
     End If
   Loop While bSingleFile = 1
 
   ' Post-fault
   Point = Point + Int( PtNo*Time4(nCase) )	' get integer part of post-fault points
   While PtCount <= Point
      ' Va, Vb, Vc
      DataValue(1) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(1) )-bV )/aV )
      DataValue(2) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(2) )-bV )/aV )
      DataValue(3) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(3) )-bV )/aV )
      ' Ia, Ib, Ic, V0, I0, V2, I2
      DataValue(4) = 0
      DataValue(5) = 0
      DataValue(6) = 0
      DataValue(7) = 0
      DataValue(8) = 0
      DataValue(9) = 0
      DataValue(10)= 0

      strText$ = Format( PtCount, "0" ) & "," & Format( 1e6*(PtCount-1)/Sample,"#######0" ) 
      If outVA(nCase) Then strText = strText & "," & Format( DataValue(1),  "####0" )
      If outVB(nCase) Then strText = strText & "," & Format( DataValue(2),  "####0" )
      If outVC(nCase) Then strText = strText & "," & Format( DataValue(3),  "####0" )
      If outIA(nCase) Then strText = strText & "," & Format( DataValue(4),  "####0" )
      If outIB(nCase) Then strText = strText & "," & Format( DataValue(5),  "####0" )
      If outIC(nCase) Then strText = strText & "," & Format( DataValue(6),  "####0" )
      If outV0(nCase) Then strText = strText & "," & Format( DataValue(7),  "####0" )
      If outI0(nCase) Then strText = strText & "," & Format( DataValue(8),  "####0" )
      If outV2(nCase) Then strText = strText & "," & Format( DataValue(9),  "####0" )
      If outI2(nCase) Then strText = strText & "," & Format( DataValue(10), "####0" )
      Print #3, strText$
      PtCount = PtCount+1
   Wend

   Print #3, Chr(26)		' End of file marker
   Close ' All files
   Comtrade = 1
End Function
'=================================End of Comtrade()====================================