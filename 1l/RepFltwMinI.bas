' ASPEN PowerScript Sample Program
'
' RepFltwMinI.BAS
'
' This program exports the fault event with minimum fault current at selected relay group
' The fault current is
' - Positive sequence current if the fault type is 3LG
' - Zero sequence current if the fault type is 1LG
' - Negative sequence current or Phase B current if the fault type is LL or LLG
' The output will be displayed in the TTY window
'
'
Const MXFLT = 2000
' Global variable declaration
Dim FltList(MXFLT) As Long
Dim FltTitle(MXFLT) As String
Dim FltDesc(MXFLT) As String
Dim FltIndex(2) As Long
Dim FltMinCur(2) As Double
Dim Bus1Hnd As Long, Bus2Hnd As Long, Branch1Hnd As Long 
Dim FltCount As Long, FltType As Long
Dim BranchName As String, Bus1Name As String, Bus2Name As String
Dim ExportDetail As Long 

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
   nCode = PageOne( FltList, FltTitle, FltDesc, FltCount&, ExportDetail& )
   If nCode = 0 Then Exit Sub

   If ExportDetail = 1 Then 
     If OpenOutFile() = 0 Then Exit Sub
     Print #1, "Fault Current at ", BranchName$
   End If
   
   If PickMinFlt( FltIndex, FltMinCur ) = 0  Then Exit Sub 
   If ExportDetail = 1 Then 
     Close
   End If  
   
   Call PrintTTY( "" )
   Call PrintTTY( "" )
   Call PrintTTY( "-----------------------------------------------------------------------------------------------------------------------------------" )
   Call PrintTTY( "--------------------------------- Fault Event With Minimum Fault Current At Selected Relay Group ----------------------------------" )
   Call PrintTTY( "-----------------------------------------------------------------------------------------------------------------------------------" )
   Call PrintTTY( "Selected relay group: " + BranchName$ ) 
   StrTemp = "Selected relay group: " + Chr(10)
   StrTemp = StrTemp + BranchName$ + Chr(10)
   StrTemp = StrTemp + Chr(10)
   If FltType = 0 Then
     Call PrintTTY( "The minimum value of I1 is " + Format(FltMinCur(1),"####0.00") + "A in fault event:" )
     Call PrintTTY( FltDesc(FltIndex(1)) )
     StrTemp = StrTemp + "The minimum value of I1 is " + Format(FltMinCur(1),"####0.00") + "A in fault event:" + Chr(10)
     FltString$ = FltDesc(FltIndex(1))
     ' Need to insert chr(13) at the end of each line to make it
     ' show up properly in the edit box
     CharPos = InStr( 1, FltString, Chr(13) + Chr(10) )
     TempStr$ = FltString$
     While CharPos > 0
       TempStr$   = Trim(Left$( FltString, CharPos - 1 ))
       StrTemp = StrTemp + TempStr$
       TempStr$ = Trim( Mid$(FltString, CharPos+1, 9999 ) ) 
       FltString$ = TempStr$
       CharPos    = InStr( 1, FltString, Chr(13) + Chr(10) )
     Wend    
     StrTemp = StrTemp + TempStr$
   End If
   

   If FltType = 1 Or FltType = 2 Then
     Call PrintTTY( "The minimum value of 3I0 is " + Format(FltMinCur(1),"####0.00") + "A in fault event:" )
     Call PrintTTY( FltDesc(FltIndex(1)) )
     StrTemp = StrTemp + "The minimum value of 3I0 is " + Format(FltMinCur(1),"####0.00") + "A in fault event:" + Chr(10)
     FltString$ = FltDesc(FltIndex(1))
     ' Need to insert chr(13) at the end of each line to make it
     ' show up properly in the edit box
     CharPos = InStr( 1, FltString, Chr(13) + Chr(10) )
     TempStr$ = FltString$
     While CharPos > 0
       TempStr$   = Trim(Left$( FltString, CharPos - 1 ))
       StrTemp = StrTemp + TempStr$
       TempStr$ = Trim( Mid$(FltString, CharPos+1, 9999 ) ) 
       FltString$ = TempStr$
       CharPos    = InStr( 1, FltString, Chr(13) + Chr(10) )
     Wend    
     StrTemp = StrTemp + TempStr$
   End If 

   If FltType = 3 Then
     Call PrintTTY( "The minimum value of I2 is " + Format(FltMinCur(1),"####0.00") + "A in fault event:" )  
     Call PrintTTY( FltDesc(FltIndex(1)) )
     Call PrintTTY( "" )
     Call PrintTTY( "The minimum value of IB is " + Format(FltMinCur(2),"####0.00") + "A in fault event:" )
     Call PrintTTY( FltDesc(FltIndex(2)) )
     StrTemp = StrTemp + "The minimum value of I2 is " + Format(FltMinCur(1),"####0.00") + "A in fault event:" + Chr(10)
     
     FltString$ = FltDesc(FltIndex(1))
     ' Need to insert chr(13) at the end of each line to make it
     ' show up properly in the edit box
     CharPos = InStr( 1, FltString, Chr(13) + Chr(10) )
     TempStr$ = FltString$
     While CharPos > 0
       TempStr$   = Trim(Left$( FltString, CharPos - 1 ))
       StrTemp = StrTemp + TempStr$
       TempStr$ = Trim( Mid$(FltString, CharPos+1, 9999 ) ) 
       FltString$ = TempStr$
       CharPos    = InStr( 1, FltString, Chr(13) + Chr(10) )
     Wend    
     StrTemp = StrTemp + TempStr$     
     StrTemp = StrTemp + Chr(10) + Chr(10)
     StrTemp = StrTemp + "The minimum value of IB is " + Format(FltMinCur(2),"####0.00") + "A in fault event:" + Chr(10)
     FltString$ = FltDesc(FltIndex(2))
     ' Need to insert chr(13) at the end of each line to make it
     ' show up properly in the edit box
     CharPos = InStr( 1, FltString, Chr(13) + Chr(10) )
     TempStr$ = FltString$
     While CharPos > 0
       TempStr$   = Trim(Left$( FltString, CharPos - 1 ))
       StrTemp = StrTemp + TempStr$
       TempStr$ = Trim( Mid$(FltString, CharPos+1, 9999 ) ) 
       FltString$ = TempStr$
       CharPos    = InStr( 1, FltString, Chr(13) + Chr(10) )
     Wend    
     StrTemp = StrTemp + TempStr$
   End If
   
   StrTemp = StrTemp + Chr(10) 
   StrTemp = StrTemp + Chr(10)
   Print StrTemp + "(The same result is printed in TTY window)"
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
  CheckBox 4,170,196,12, "Export details to *.csv file", .CheckBox_1
  PushButton 176,180,44,12, "Export", .Next
  TextBox 4,28,308,135, .EditBox_1, WSTYLE1
  CancelButton 236,180,36,12
End Dialog

'====================================PageOne()=========================================
' Purpose:
'   Solicit user input on the scope of the export
'
'======================================================================================
Function PageOne( ByRef FltList() As Long, ByRef FltTitle() As String, ByRef FltDesc() As String, _
                  ByRef FltCount As Long, ByRef ExportDetail As Long ) As Long
  Dim dlg As PAGE1

  ' Prepare fault list 
  AString$ = ""
  If PickFault( 1 ) = 0 Then Exit Sub	' No fault
  Do
    FltString$ = FaultDescription()
    ' Need to insert chr(13) at the end of each line to make it
    ' show up properly in the edit box
    CharPos = InStr( 1, FltString, Chr(10) )
    TempStr$ = FltString$
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
      FltTitle(FltCount) = ALine
      FltDesc(FltCount) = ALine 
    Else
      FltDesc(FltCount) = FltDesc(FltCount) + Chr(13) + Chr(10) + ALine$
    End If
    nLength  = Len(AString$)
    If nLength > 65000 Then 
      Print "Too many fault simulation."
      PageOne = 0
      Exit Function
    End If
    AString$ = Mid$(AString$, CharPos+1, nLength )
    CharPos  = InStr( 1, AString$, Chr(10) )
  Wend
  ExportDetail = dlg.CheckBox_1
  If FltCount = 0 Then 
    PageOne   = 0
    Exit Function   ' Nothing to report. Cancel
  End If
End Function 'Page1
'=================================OpenOutFile()========================================
' Purpose:
'   Specify output file name and path
'
'======================================================================================
Function OpenOutFile() As Long
   ' Open file for output

   ' Dialog data generated by Dialog Edito
   Begin Dialog Dialog_1 49,60, 202, 56, "Output File"
      Text 24,12,56,12, "Enter file name: "
      TextBox 84,12,84,12, .EditBox_1
      OKButton 44,36,52,12
      CancelButton 108,36,48,12
   End Dialog
   Dim dlg As Dialog_1
   Dlg.EditBox_1 = ".csv"         ' Default name
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
'==================================PickMinFlt()========================================
' Purpose:
'   Return the fault index with minimum fault value
'
'======================================================================================
Function PickMinFlt( ByRef FltIndex() As Long, ByRef FltMinCur() As Double ) As Long
  Dim MagArray(3) As Double, AngArray(3) As Double
  PickMinFlt = 0
  MinValue   = 0
  MinValue1  = 0
  MinIndex   = 0
  MinIndex1  = 0
  FltType = -1
  If InStr( 1, FltTitle(1), " 3LG" ) > 0 Then FltType = 0
  If InStr( 1, FltTitle(1), " 2LG" ) > 0 Then FltType = 1
  If InStr( 1, FltTitle(1), " 1LG" ) > 0 Then FltType = 2
  If InStr( 1, FltTitle(1), " LL" )  > 0 Then FltType = 3

  For ii& = 1 To Fltcount
    If FltType = 0 And InStr( 1, FltTitle(ii), " 3LG" ) = 0 Then FltType = -1
    If FltType = 1 And InStr( 1, FltTitle(ii), " 2LG" ) = 0 Then FltType = -1
    If FltType = 2 And InStr( 1, FltTitle(ii), " 1LG" ) = 0 Then FltType = -1
    If FltType = 3 And InStr( 1, FltTitle(ii), " LL" )  = 0 Then FltType = -1
    If FltType = -1 Then
      Print "The fault type is not consistent" 
      If ExportDetail = 1 Then 
        Close
      End If  
      Exit Sub
    End If
    If PickFault( FltList(ii) ) = 0 Then 
      Print "Problem reading fault simulation result #", FltList(nCase)
      Exit Function
    End If
    ' Bbranch current
    If GetSCCurrent( Branch1Hnd, MagArray, AngArray, 2 ) = 0 Then 
      Print "Get Branch current failed."
      Exit Function
    End If 
    If FltType = 0 Then 
      If ii = 1 Then 
        If ExportDetail = 1 Then Print #1, "Index" + "," + "Fault Description" + "," + "I1"
        MinValue = MagArray(2)
        MinIndex = 1
      Else
        If MagArray(2) < MinValue Then
          MinValue = MagArray(2)
          MinIndex = ii
        End If  
      End If
      Temp = MagArray(2)
    End If

    If FltType = 1 Or FltType = 2 Then 
      If ii = 1 Then 
        If ExportDetail = 1 Then Print #1, "Index" + "," + "Fault Description" + "," + "I0"
        MinValue = MagArray(1)
        MinIndex = 1
      Else
        If MagArray(1) < MinValue Then
          MinValue = MagArray(1)
          MinIndex = ii
        End If  
      End If
      Temp = 3*MagArray(1)
    End If
    
    If FltType = 3 Then 
      If ii = 1 Then 
        If ExportDetail = 1 Then Print #1, "Index" + "," + "Fault Description" + "," + "I2" + "," + "IB"
        MinValue = MagArray(3)
        MinIndex = 1
      Else
        If MagArray(3) < MinValue Then
          MinValue = MagArray(3)
          MinIndex = ii
        End If  
      End If
      Temp = MagArray(3)
      If GetSCCurrent( Branch1Hnd, MagArray, AngArray, 4 ) = 0 Then 
        Print "Get Branch current failed."
      Exit Function
      End If
      If ii = 1 Then 
        MinValue1 = MagArray(2)
        MinIndex1 = 1
      Else
        If MagArray(2) < MinValue1 Then
          MinValue1 = MagArray(2)
          MinIndex1 = ii
        End If  
      End If
    End If 
    If ExportDetail = 1 Then
      StrTemp = ""
      FltString$ = FaultDescription()
      CharPos = InStr( 1, FltString, Chr(10) )
      TempStr$ = FltString$
      While CharPos > 0
        TempStr$   = Trim(Left$( FltString, CharPos - 1 ))
        StrTemp = StrTemp + " " + TempStr$
        TempStr$ = Trim( Mid$(FltString, CharPos+1, 9999 ) ) 
        FltString$ = TempStr$
        CharPos    = InStr( 1, FltString, Chr(10) )
      Wend    
      StrTemp = StrTemp + " " + TempStr$
      If FltType = 3 Then  
        Print #1, Str(ii) + "," + StrTemp + "," + Format(Temp,"####0.00") + "," + Format(MagArray(2),"####0.00") 
      Else 
        Print #1, Str(ii) + "," + StrTemp + "," + Format(Temp,"####0.00")
      End If                
    End If
  Next
  FltIndex(1) = MinIndex
  FltIndex(2) = MinIndex1
  If FltType = 1 Or FltType = 2 Then
    FltMinCur(1) = MinValue*3
  Else
    FltMinCur(1) = MinValue
  End If
  FltMinCur(2) = MinValue1  
  PickMinFlt = 1             
End Function




