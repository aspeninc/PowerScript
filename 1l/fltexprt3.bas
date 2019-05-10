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
' Version 3.0
' Category: OneLiner
'
Const pi = 3.141592654
Dim OutZ As Long, OutN As Long
Sub main()
   ' Variable declaration
   Dim Choice(4) As Long, FileType As Long
   Dim FltList(200) As Long
   Dim FltDesc(200) As String
   Dim Branch1Hnd As Long, Branch2Hnd As Long, BusHnd As Long, Bus1Hnd As Long, Bus2Hnd As Long, Bus3Hnd As Long
   Dim PickedHnd As Long, LineHnd As Long
   Dim LineName As String
   Dim CT_Ratio1 As Double, CT_Ratio2 As Double, PT_Ratio1 As Double, PT_Ratio2 As Double 
   Dim TapCode As Long, ttt As Long, Bus1kV As Double
   Dim FName1(7) As String, FName2(7) As String, FileName1 As String, FileName2 As String
   Dim StringVal As String
   Dim Bus1Name As String, Bus2Name As String
   Dim Bus1No As Long, Bus2No As Long, LineID As Long
   Dim TimePeriod1(3) As Double	'For doble
   Dim TimePeriod2(6) As Double	'For comptrade
   Dim Flag As Long
   Dim Comments As String

   ' Get picked relay group handle
   If GetEquipment( TC_PICKED, PickedHnd ) = 0 Then 
      Print "Must select a relay group"
      Exit Sub
   End If
   ' Must be a relay group
   If EquipmentType( PickedHnd ) <> TC_RLYGROUP Then
      Print "Must select a relay group"
      Exit Sub
   End If

   ' There should be some faults simulated
   If PickFault(1) = 0 Then 	
     Print "No fault simulation result available"
     Exit Sub
   End If

   Flag = 0	' Flag = 0 for a line; Flag = 1 for other equipments
   ' Get the relay group branch handle
   If GetData( PickedHnd, RG_nBranchHnd, Branch1Hnd ) = 0 Then GoTo HasError
   ' Get the branch near bus handle
   If GetData( Branch1Hnd, BR_nBus1Hnd, Bus1Hnd ) = 0 Then GoTo HasError

   ' Judge if it is a line
   If GetData( Branch1Hnd, BR_nType, TypeCode ) = 0 Then GoTo HasError
   If TypeCode <> TC_LINE Then
      Flag = 1	' Flag = 1 for non-line equipments
      Bus2Hnd    = -1	' non-existent
      Branch2Hnd = -1	' non-existent
   End If
 
   If Flag = 0 Then		' for line equipment
      ' Get line info
      If GetData( Branch1Hnd, BR_nHandle, LineHnd ) = 0 Then GoTo HasError
      If GetData( LineHnd, LN_sName, LineName ) = 0 Then GoTo HasError
      If GetData( Bus1Hnd, BUS_dkVnorminal, Bus1kV ) = 0 Then GoTo HasError	' near bus nominal kV
      BusHnd = Bus1Hnd		' temporary near bus handle

      ' far bus
      If GetData( Branch1Hnd, BR_nBus2Hnd, Bus2Hnd ) = 0 Then GoTo HasError	' far bus handle
   
      ' Must skip all taps on original lines
      Do 
         If GetData( Bus2Hnd, BUS_nTapBus, TapCode ) = 0 Then GoTo HasError
         If TapCode = 0 Then Exit Do			' real bus
         ' Only for tap bus
         Branch2Hnd = 0
         ttt = GetBusEquipment( Bus2Hnd, TC_BRANCH, Branch2Hnd )
         While ttt <> 0
            If GetData( Branch2Hnd, BR_nBus2Hnd, Bus3Hnd ) = 0 Then GoTo HasError	' Get the far end bus
            If Bus3Hnd <> BusHnd Then	' for different branch
               If GetData( Branch2Hnd, BR_nType, TypeCode ) = 0 Then GoTo HasError	' Get branch type
               If TypeCode = TC_LINE Then 
                  ' Get line name
                  If GetData( Branch2Hnd, BR_nHandle, LineHnd ) = 0 Then GoTo HasError
                  If GetData( LineHnd, LN_sName, StringVal ) = 0 Then GoTo HasError
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
   End If

   ' Bus information and line information
   Bus1Name = FullBusName( Bus1Hnd )	' near bus name
   If GetData( Bus1Hnd, BUS_nNumber, Bus1No ) = 0 Then GoTo HasError	' Get bus 1 No.

   If Flag = 0 Then	' for line
      Bus2Name = FullBusName( Bus2Hnd )	' far bus name
      If GetData( Bus2Hnd, BUS_nNumber, Bus2No ) = 0 Then GoTo HasError		' Get bus 2 No.
      If GetData( Branch1Hnd, BR_nHandle, LineHnd ) = 0 Then GoTo HasError	' Get line 1 handle
      If GetData( LineHnd, LN_sID, Line1No ) = 0 Then GoTo HasError		' Get line 1 ID
      If GetData( Branch2Hnd, BR_nHandle, LineHnd ) = 0 Then GoTo HasError	' Get line 2 handle
      If GetData( LineHnd, LN_sID, Line2No ) = 0 Then GoTo HasError		' Get line 2 ID
      If GetData( LineHnd, LN_NRlyGr1Hnd, RlyGrHnd ) = 0 Then 			' Get line 2 relay group handle
         Flag = 1	' Flag = 1 for non-line equipments or no relay group at the far end
         Bus2Hnd    = -1	' non-existent
         Branch2Hnd = -1	' non-existent
      End If
   End If
      
   FileName1 = "c:\Aspen01\script\bus1"
   FileName2 = "c:\Aspen01\script\bus2" 
   If DiaScope( FltList, FltDesc, Choice, FileType, OutZ, OutN, CT_Ratio1, CT_Ratio2, PT_Ratio1, PT_Ratio2 ) = 0 Then Exit Sub
   If DiaFile( FileName1, 1 ) = 0 Then Exit Sub
   If Flag = 0 Then	' line with two relay groups
      If DiaFile( FileName2, 2 ) = 0 Then Exit Sub
   End If
   If Choice(3) = 1 Then	' Doble file
      If DiaTime ( TimePeriod1 ) = 0 Then Exit Sub
   End If
   If Choice(4) = 1 Then	' COMTRADE file
      If DiaTime1 ( TimePeriod2, Comments ) = 0 Then Exit Sub
      'If PickFault(2) <> 0 Then 	
      '   Print "You have simulated more than one fault. Results in the first fault will only be output to the COMTRADE files"
      'End If
   End If
   
   FName1(1) = FileName1 & ".out"
   FName2(1) = FileName2 & ".out"
   FName1(2) = FileName1 & ".csv"
   FName2(2) = FileName2 & ".csv"
   FName1(3) = FileName1 & ".ss1"
   FName2(3) = FileName2 & ".ss1"
   FName1(4) = FileName1 & ".txt"
   FName2(4) = FileName2 & ".txt"
   FName1(5) = FileName1 & ".dat"
   FName2(5) = FileName2 & ".dat"
   FName1(6) = FileName1 & ".cfg"
   FName2(6) = FileName2 & ".cfg"
   FName1(7) = FileName1 & ".hdr"
   FName2(7) = FileName2 & ".hdr"
   
   ' Prepare Output file
   If FileType = 0 Then	' append
      If Choice(1) = 1 Then 	' .out
         If Flag = 0 Then	' line
            Open FName1(1) For Append As 1
            Open FName2(1) For Append As 5
         Else
            Open FName1(1) For Append As 1
         End If
      End If
      If Choice(2) = 1 Then 	' .csv
         If Flag = 0 Then	' line
            Open FName1(2) For Append As 2
            Open FName2(2) For Append As 6
         Else
            Open FName1(2) For Append As 2
         End If
      End If
      If Choice(3) = 1 Then 	' .ss1, .txt
         Print "Doble file can not use append as output"
      End If
      If Choice(4) = 1 Then 	' .dat, .cfg
         Print "COMTRADE file can not use append as output"
      End If
   Else	' overwrite
      If Choice(1) = 1 Then 	' .out
         If Flag = 0 Then	' line
            Open FName1(1) For Output As 1
            Open FName2(1) For Output As 5
         Else
            Open FName1(1) For Output As 1
         End If
      End If
      If Choice(2) = 1 Then 	' .csv
         If Flag = 0 Then	' line
            Open FName1(2) For Output As 2
            Open FName2(2) For Output As 6
         Else
            Open FName1(2) For Output As 2
         End If
      End If
      If Choice(3) = 1 Then 	' .ss1
         If Flag = 0 Then	' line
            Open FName1(3) For Output As 3
            Open FName2(3) For Output As 7
            Open FName1(4) For Output As 4
            Open FName2(4) For Output As 8
         Else
            Open FName1(3) For Output As 3
         End If
      End If
      If Choice(4) = 1 Then 	' .dat, .cfg, .hdr
         If Flag = 0 Then	' line
            Open FName1(5) For Output As 9
            Open FName1(6) For Output As 10
            Open FName1(7) For Output As 11
            Open FName2(5) For Output As 12
            Open FName2(6) For Output As 13
            Open FName2(7) For Output As 14
         Else
            Open FName1(5) For Output As 9
            Open FName1(6) For Output As 10
            Open FName1(7) For Output As 11
         End If
      End If  
   End If

   If Choice(1) = 1 Then	' Tab delimited (.out)
      If TextFile( Bus1Hnd, Bus2Hnd, Branch1Hnd, Branch2Hnd, FileType, Chr(9), 1, FltList, FltDesc, _
                   CT_Ratio1, CT_Ratio2, PT_Ratio1, PT_Ratio2 ) = 0 Then
         Print "Output tab delimited file failed."
         Exit Sub
      End If
      If Flag = 0 Then	'line
         Print "Results have been exported in tab-delimited format to: "; FName1(1); " & "; FName2(1)
      Else
         Print "Results have been exported in tab-delimited format to: "; FName1(1)
      End If
   End If
   If Choice(2) = 1 Then	' Comma delimited (.csv)
      If TextFile( Bus1Hnd, Bus2Hnd, Branch1Hnd, Branch2Hnd, FileType, Chr(44), 2, FltList, FltDesc, _
                   CT_Ratio1, CT_Ratio2, PT_Ratio1, PT_Ratio2 ) = 0 Then
         Print "Output text file failed."
         Exit Sub
      End If
      If Flag = 0 Then 	'line
         Print "Results have been exported in comma-delimited format to: "; FName1(2); " & "; FName2(2)
      Else
         Print "Results have been exported in tab-delimited format to: "; FName1(2)
      End If
   End If
   If Choice(3) = 1 Then	' Doble (.ss1, .txt)
      If Doble( Bus1Hnd, Bus2Hnd, Branch1Hnd, Branch2Hnd, FltList, FltDesc, _
                CT_Ratio1, CT_Ratio2, PT_Ratio1, PT_Ratio2, TimePeriod1 ) = 0 Then
         Print "Output Doble file failed."
         Exit Sub
      End If
      If Flag = 0 Then	' line
         ' Output to key file (.txt)
         If KeyFile( FName1(3), CT_Ratio1, PT_Ratio1,Bus1kV, Bus1No, Bus2No, Bus1Name, Bus2Name, _
                  Line1ID, FltList, FltDesc, 4 ) = 0 Then
            Print "Output Key file for bus 1 failed."
            Exit Sub
         End If
         If KeyFile( FName2(3), CT_Ratio2, PT_Ratio2, Bus1kV, Bus2No, Bus1No, Bus2Name, Bus1Name, _
                  Line2ID, FltList, FltDesc, 8 ) = 0 Then
            Print "Output Key file for bus 2 failed."
            Exit Sub
         End If
      End If
      If Flag = 0 Then	'line
         Print "Results have been exported in Doble format to: "; FName1(3); " & "; FName2(3); " & "; FName1(4); " & "; FName2(4)
      Else
         Print "Results have been exported in Doble format to: "; FName1(3)
      End If
   End If
   If Choice(4) = 1 Then	' COMTRADE (.dat, .cfg)
      If Comtrade( Bus1Hnd, Branch1Hnd, FltList, FltDesc(), CT_Ratio1, PT_Ratio1, 9, TimePeriod2, Comments ) = 0 Then
         Print "Output COMTRADE file 1 failed."
         Exit Sub
      End If
      If Flag = 0 Then	'line
         If Comtrade( Bus2Hnd, Branch2Hnd, FltList, FltDesc(), CT_Ratio2, PT_Ratio2, 12, TimePeriod2, Comments ) = 0 Then
            Print "Output COMTRADE file 2 failed."
            Exit Sub
         End If
      End If
      If Flag = 0 Then	'line
         Print "Results have been exported in COMTRADE format to: "; FName1(5); " & "; FName2(5); " & "; FName1(6); " & "; FName2(6)
      Else
         Print "Results have been exported in COMTRADE format to: "; FName1(5); " & "; FName1(6)
      End If
   End If

   Close 'Close all files

   Exit Sub
HasError:
   Print "Error: ", ErrorString( )
   Close
End Sub
'===================================End of Main()====================================

'===================================dAdd()===========================================
' Calculate prefault branch current
Function dAdd( ByVal Mag1 As Double, ByVal Ang1 As Double, _
               ByVal Mag2 As Double, ByVal Ang2 As Double, _
               ByRef Mag3 As Double, ByRef Ang3 As Double ) As Long
   Ang1   = Ang1 *3.14159/180.0
   Real1# = Mag1*Cos(Ang1)
   Imag1# = Mag1*Sin(Ang1)
   Ang2   = Ang2 *3.14159/180.0
   Real2# = Mag2*Cos(Ang2)
   Imag2# = Mag2*Sin(Ang2)
   Real3# = Real1+Real2
   Imag3# = Imag1+Imag2
   Call dR2P ( Real3, Imag3, Mag3, Ang3 )
   dSub = 1
End Function

'===================================dR2P()===========================================
' Convert complext # from rect to polar
Function dR2P( ByVal Real As Double, ByVal Imag As Double, _
               ByRef Mag As Double, ByRef Ang As Double ) As Long
   Ang  = Atn(Abs(Imag/Real))
   Mag  = Abs(Real / Cos(Ang))
   Ang  = Ang*180.0/3.14159
   If Imag > 0 And Real < 0 Then Ang = Ang + 90
   If Imag < 0 And Real < 0 Then Ang = Ang + 180
   If Imag < 0 And Real > 0 Then Ang = Ang - 90
   dSub = 1
End Function

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
Const WSTYLE = WS_VSCROLL Or ES_MULTILINE Or ES_AUTOVSCROLL
'====================Dialog box spec (generated by Dialog Editor)===================
Begin Dialog OUTPUTDIA 90,55, 345, 260, "Export Fault Result for Relay Testing"
  OptionGroup .GROUP_1
    OptionButton 108,132,68,12, "&Append file"
    OptionButton 232,132,64,12, "&Overwrite file"
  GroupBox 4,120,336,88, "Output Options"
  TextBox 3,16,336,80, .EditBox_1, WSTYLE
  Text 4,4,276,8, "Following fault results are available for export:"
  Text 8,132,60,12, "Output action:"
  Text 4,100,336,8, "Edit the list above to keep only faults that you want to export. You may also edit fault description strings"
  Text 4,108,292,8, "but fault index numbers in front of each fault must be kept unchanged."
  Text 8,150,40,12, "File type:"
  Text 8,190,76,12, "Sequence quantities:"
  CheckBox 108,148,105,16, "Tab delimited (.out)", .CheckBox_1
  CheckBox 232,148,104,16, "Comma delimited (.csv)", .CheckBox_2
  CheckBox 108,164,105,16, "Doble (.ss1, .txt)", .CheckBox_3
  CheckBox 232,164,105,16, "COMTRADE (.dat, .cfg, .hdr)", .CheckBox_4
  CheckBox 108,188,81,12, "Io and Vo", .CheckBox_5
  CheckBox 232,188,81,12, "I2 and V2", .CheckBox_6
  Text 4,216,41,12, "CT ratio 1 = "
  TextBox 44,216,37,12, .EditBox_2
  Text 92,216,41,12, "CT ratio 2 = "
  TextBox 132,216,37,12, .EditBox_3
  Text 172,216,41,12, "PT ratio 1 = "
  TextBox 212,216,37,12, .EditBox_4
  Text 260,216,41,12, "PT ratio 2 = "
  TextBox 300,216,37,12, .EditBox_5
  OKButton 96,242,76,12
  CancelButton 188,242,40,12
  PushButton 236,242,40,12, "Help", .PushButton_1
End Dialog


'=============================End of Dialog box spec===================================


'====================================DiaScope()========================================
' Purpose:
'   Solicit user input on the scope of the export
'
'======================================================================================
Function DiaScope( ByRef FltList() As Long, ByRef FltDesc() As String, ByRef Choice() As Long, _
         ByRef FileType As Long, ByRef OutZ As Long, ByRef OutN As Long, _
         ByRef CT_Ratio1 As Double, ByRef CT_Ratio2 As Double, _
         ByRef PT_Ratio1 As Double, ByRef PT_Ratio2 As Double ) As Long
Dim dlg As OUTPUTDIA

' DiaScope = 0: dialog fail;	1: dialog success
' FileType = 0: Append;		1: overwrite
' Choice(1)= 1: Tab delimited (.out);		Choice(2)= 1: Comma delimited (.csv)
' Choice(3)= 1: Doble (.ss1, .txt);		Choice(4)= 1: COMTRADE (.dat, .cfg)

DiaScope      = 0
' Initial file names
dlg.EditBox_2 = 300
dlg.EditBox_3 = 300
dlg.EditBox_4 = 2000
dlg.EditBox_5 = 2000
dlg.CheckBox_1= 1
dlg.CheckBox_2= 0
dlg.CheckBox_3= 0
dlg.CheckBox_4= 0
dlg.CheckBox_5= 1
dlg.CheckBox_6= 1
dlg.GROUP_1   = 1

' Prepare fault list 
AString$ = ""
If PickFault( 1 ) = 0 Then Exit Sub	' No fault
Do
   FltString$ = FaultDescription()
   ' Need to insert chr(13) at the end of each line to make it
   ' show up properly in the edit box
   CharPos = InStr( 1, FltString, Chr(10) )
   If CharPos = 0 Then 
    TempStr$ = FltString$
   Else
    While CharPos > 0
      TempStr$   = Left$( FltString, CharPos - 1 )
      If Right( TempStr, 3 ) <> "on:" Then TempStr = TempStr + Chr(13) + Chr(10)
      TempStr    = TempStr + " " + LTrim( Mid$(FltString, CharPos+1, 9999 ) )
      FltString$ = TempStr$
      CharPos    = InStr( CharPos+2, FltString, Chr(10) )
    Wend
   End If
   AString$ = AString$ + TempStr$ + Chr(13) + Chr(10) 
Loop While PickFault( SF_NEXT ) > 0

' Initialize dialog box
dlg.EditBox_1 = AString$

DoneFlag = 0
While DoneFlag = 0
   button = Dialog(dlg)	' -1 = OK; 0 = Cancel; 1 = Help
   If button = 0 Then	' Cancel
      DiaScope = 0
      Exit Function
   End If
   If button = 1 Then	' Help
	Print _
        "This program exports fault simulation results to disk files for relay" & Chr(13) & Chr(10) & _
        "testing purposes. Four types of files can be exported:" & Chr(13) & Chr(10) & _
        " - Tab-delimited file (.out)" & Chr(13) & Chr(10) & _
        " - Comma-delimited file (.csv)" & Chr(13) & Chr(10) & _
        " - Doble format file (.ss1, .txt)" & Chr(13) & Chr(10) & _
        " - COMTRADE format file (.dat, .cfg, .hdr)" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
        "Following simulation results are saved in the files (in secondary quantities):" & Chr(13) & Chr(10) & _
        " - Three-phase faulted voltages and currents" & Chr(13) & Chr(10) & _
        " - Zero and negative-sequence faulted voltages and currents" & Chr(13) & Chr(10) & _
        " - Pre-fault voltages (Doble and COMTRADE only)" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
        "A relay group has to be highlighted before running the script program. The export" & Chr(13) & Chr(10) & _
        "file will contain quantities seen by a relay in the selected group." & Chr(13) & Chr(10) & _
        "If the selected relay group is on a line that has two relay groups at it two ends," & Chr(13) & Chr(10) & _
        "the script will also export the result at the remote relay group to second set" & Chr(13) & Chr(10) & _
        "of files." & Chr(13) & Chr(10) &  Chr(13) & Chr(10) & _
        "Multiple fault simulations can be saved in a single file except in the COMTRADE" & Chr(13) & Chr(10) & _
        "format file which can contain only a one fault result."
   End If             

   ' Parse FaultString to get the list of fault number to output
   Count          = 1
   FltList(Count) = -1
   AString$       = dlg.EditBox_1
   CharPos&       = InStr( 1, AString$, Chr(10) )
   While CharPos > 0
      ALine$    = Left$( AString$, CharPos - 2 )
      CharPos1& = InStr( 1, ALine$, ". " )
      If CharPos1 > 0 And CharPos1 < 10 Then
         TempStr$       = Left$( ALine$, CharPos1 - 1 )
         FltList(Count) = Val( TempStr$ )
         FltDesc(Count) = ALine
         Count          = Count + 1
         FltList(Count) = -1 ' Mark the list end
      End If
      AString$ = Mid$(AString$, CharPos+1, 9999 )
      CharPos  = InStr( 1, AString$, Chr(10) )
   Wend
   If Count = 1 Then 
      DiaScope = 0
      Exit Function   ' Nothing to report. Cancel
   End If

   CT_Ratio1 = Val(dlg.EditBox_2)
   CT_Ratio2 = Val(dlg.EditBox_3)
   PT_Ratio1 = Val(dlg.EditBox_4)
   PT_Ratio2 = Val(dlg.EditBox_5)
   Choice(1) = dlg.CheckBox_1	' Tab delimited (.out)
   Choice(2) = dlg.CheckBox_2	' Comma delimited (.csv)
   Choice(3) = dlg.CheckBox_3	' Doble (.ss1, .txt)
   Choice(4) = dlg.CheckBox_4	' COMTRADE (.dat, .cfg)
   OutZ      = dlg.CheckBox_5
   OutN      = dlg.CheckBox_6
   FileType  = dlg.GROUP_1 	' 0 = Append, 1 = overwrite

   If FileType = 0 Then
      If ( Choice(3) = 1 Or Choice(4) = 1 ) Then
         dlg.GROUP_1 = 1
      End If
   End If
   If ( CT_Ratio1 <= 0 Or CT_Ratio2 <= 0 Or PT_Ratio1 <= 0 Or PT_Ratio2 <= 0 ) Then 
      Print "CT ratio and PT ratio must be >0"
   ElseIf ( Choice(1) = 0 And Choice(2) = 0 And Choice(3) = 0 And Choice(4) = 0 ) Then
      Print "At least one output type must be selected"
   Else
      If FileType = dlg.GROUP_1 Then 
         DoneFlag = 1
      Else 
         Print "Doble or COMTRADE file can not use append as output"
      End If
   End If
   If button <> -1 Then DoneFlag = 0	' ok button
Wend
DiaScope = 1
End Function
'===================================End of DiaScope()==================================


'====================Dialog box spec (generated by Dialog Editor)===================
Begin Dialog FILE1 57,49, 171, 98, "Output the First File Name"
  Text 12,12,145,12, "Please enter the first output file name:"
  TextBox 12,36,145,12, .EditBox_1
  OKButton 12,68,37,12
  CancelButton 120,68,37,12
End Dialog
Begin Dialog FILE2 57,49, 182, 98, "Output the Second File Name"
  Text 12,12,153,12, "Please enter the second output file name:"
  TextBox 12,36,153,12, .EditBox_1
  OKButton 12,68,37,12
  CancelButton 128,68,37,12
End Dialog
'=============================End of Dialog box spec===================================


'====================================DiaFile()========================================
' Purpose:
'   Solicit user enter output file name
'
'======================================================================================
Function DiaFile( ByRef FileName As String, ByVal FileNo As Long ) As Long
DiaFile = 0

If FileNo = 1 Then
   Dim dlg1 As FILE1
Else
   Dim dlg2 As FILE2
End If

' Initial values
If FileNo = 1 Then
   dlg1.EditBox_1 = FileName
Else
   dlg2.EditBox_1 = FileName
End If

DoneFlag = 0
While DoneFlag = 0
   If FileNo = 1 Then
      button = Dialog(dlg1)
   Else
      button = Dialog(dlg2)
   End If
   If button = 0 Then	' cancelled
      DiaFile = 0
      Exit Function
   End If
   If FileNo = 1 Then
      FileName = dlg1.EditBox_1
   Else
      FileName = dlg2.EditBox_1
   End If
   If button = -1 Then 
      If FileName = "" Then 
         Print "Must enter a file name"
         DoneFlag = 0
      Else 
         DoneFlag = 1
      End If
   End If
Wend
DiaFile = 1
End Function
'===================================End of DiaTime()==================================


'====================Dialog box spec (generated by Dialog Editor)===================
Begin Dialog DURATION 57,49, 129, 126, "Export Time"
  Text 16,36,37,12, "Pre-fault   = "
  TextBox 56,36,37,12, .EditBox_1
  Text 96,36,29,12, "cycles"
  Text 16,56,41,12, "Fault         = "
  TextBox 56,56,37,12, .EditBox_2
  Text 96,56,29,12, "cycles"
  Text 16,76,37,12, "Post-fault = "
  TextBox 56,76,37,12, .EditBox_3
  Text 96,76,29,12, "cycles"
  Text 16,16,85,12, "   Time Duration for Doble"
  OKButton 16,100,37,12
  CancelButton 72,100,37,12
End Dialog
'=============================End of Dialog box spec===================================


'====================================DiaTime()========================================
' Purpose:
'   Solicit user input time duration for Doble file
'
'======================================================================================
Function DiaTime( ByRef TimePeriod() As Double ) As Long
Dim dlg As DURATION

DiaTime = 0
' Initial values
dlg.EditBox_1 = 5.0
dlg.EditBox_2 = 5.0
dlg.EditBox_3 = 5.0

DoneFlag = 0
While DoneFlag = 0
   button = Dialog(dlg)
   If button = 0 Then	' cancelled
      DiaTime = 0
      Exit Function
   End If
   TimePeriod(1) = dlg.EditBox_1
   TimePeriod(2) = dlg.EditBox_2
   TimePeriod(3) = dlg.EditBox_3
   If ( TimePeriod(1) >= 0 And TimePeriod(2) >= 0 And TimePeriod(1) >= 0 ) Then
      DoneFlag = 1
   Else
      Print "Time duration must be >= 0"
   End If
Wend
DiaTime = 1
End Function
'===================================End of DiaTime()==================================


'====================Dialog box spec (generated by Dialog Editor)===================
Begin Dialog CYCLE 57,49, 185, 189, "Export to COMTRADE"
  OptionGroup .GROUP_1
    OptionButton 26,97,58,8, "Three pole"
    OptionButton 100,95,69,12, "Single pole"
  GroupBox 4,4,172,80, " Time Durations (60Hz cycles) "
  GroupBox 4,88,172,22, " Tripping Option "
  Text 51,18,37,12, "Pre-fault = "
  TextBox 89,16,24,12, .EditBox_1
  Text 60,35,27,12, "Fault  = "
  TextBox 89,33,24,12, .EditBox_2
  Text 44,51,44,8, "Re-closing ="
  TextBox 89,48,24,12, .EditBox_3
  Text 48,64,39,12, "Post-fault = "
  TextBox 89,64,24,12, .EditBox_4
  Text 5,114,134,12, "Comments to save in the header file:"
  TextBox 5,126,171,33, .EditBox_7, WSTYLE
  Text 6,160,129,10, "(Use Ctrl-Enter to start a new line)"
  OKButton 45,173,53,12
  CancelButton 102,173,37,12
End Dialog


'=============================End of Dialog box spec===================================


'====================================DiaTime1()========================================
' Purpose:
'   Solicit user input time cycle for COMTRADE file
'
'======================================================================================
Function DiaTime1( ByRef TimePeriod() As Double, ByRef Comments As String ) As Long
Dim dlg As CYCLE

DiaTime1 = 0
' Initial values
dlg.EditBox_1 = 5.0
dlg.EditBox_2 = 5.0
dlg.EditBox_3 = 0.0
dlg.EditBox_4 = 5.0
dlg.Group_1   = 0

DoneFlag = 0
While DoneFlag = 0
   button = Dialog(dlg)
   If button = 0 Then	' cancelled
      DiaTime1 = 0
      Exit Function
   End If
   If ( dlg.EditBox_1 < 0.0 Or dlg.EditBox_2 < 0 Or dlg.EditBox_3 < 0 Or dlg.EditBox_4 < 0 ) Then
      Print "Time duration cannot be negative"
   Else
      DoneFlag = 1
   End If
Wend
TimePeriod(1) = dlg.EditBox_1
TimePeriod(2) = dlg.EditBox_2
TimePeriod(3) = dlg.EditBox_3
TimePeriod(4) = dlg.EditBox_4
If dlg.Group_1 = 1 Then TimePeriod(5) = 1.0 Else TimePeriod(5) = 0.0
Comments = dlg.EditBox_7
DiaTime1 = 1
End Function
'===================================End of DiaTime1()==================================


'=====================================TextFile()====================================
' Purpose:
'   Print out fault result to tab-delimited and comma-delimited file
'
'======================================================================================
Function TextFile( ByVal Bus1Hnd As Long, ByVal Bus2Hnd As Long, ByVal Branch1Hnd As Long, ByVal Branch2Hnd As Long, _
         ByVal FileType As Long, ByVal Delim As String, ByVal FileNO As Long, ByRef FltList() As Long, _
         ByRef FltDesc() As String, ByRef CT_Ratio1 As Double, ByRef CT_Ratio2 As Double, _
         ByRef PT_Ratio1 As Double, ByRef PT_Ratio2 As Double ) As Long

   Dim MagArray(16) As Double
   Dim AngArray(16) As Double
   Dim TempHnd As Long

   TextFile = 0
   ' near bus name
   If GetData( Bus1Hnd, BUS_sName, Bus1Name$ ) = 0 Then	
      Print "Get Bus 1 name failed."
      Exit Function
   End If
   ' far bus name
   If Bus2Hnd <> -1 Then
      If GetData( Bus2Hnd, BUS_sName, Bus2Name$ ) = 0 Then	
         Print "Get Bus 2 name failed."
         Exit Function
      End If
   Else
      If GetData( Branch1Hnd, BR_nBus2Hnd, TempHnd ) = 0 Then	
         Print "Get Bus 2 handle failed."
         Exit Function
      End If
      If GetData( TempHnd, BUS_sName, Bus2Name$ ) = 0 Then	
         Print "Get Bus 2 name failed."
         Exit Function
      End If
   End If

   ' Print column header if overwrite
   If FileType = 1 Then 
      strTemp$ = _
         Chr(34) & "Comment" & Chr(34) & Delim & _
         Chr(34) & Bus1Name & "  ->  " & Bus2Name & Chr(34) & Delim & _
         Chr(34) & "Va"     & Chr(34)  & Delim & _
         Chr(34) & " "      & Chr(34)  & Delim & _
         Chr(34) & "Ia"     & Chr(34)  & Delim & _
         Chr(34) & " "      & Chr(34)  & Delim & _
         Chr(34) & "Vb"     & Chr(34)  & Delim & _
         Chr(34) & " "      & Chr(34)  & Delim & _
         Chr(34) & "Ib"     & Chr(34)  & Delim & _
         Chr(34) & " "      & Chr(34)  & Delim & _
         Chr(34) & "Vc"     & Chr(34)  & Delim & _
         Chr(34) & " "      & Chr(34)  & Delim & _
         Chr(34) & "Ic"     & Chr(34)  & Delim & _
         Chr(34) & " "      & Chr(34)  & Delim
      If OutZ = 1 Then
         strTemp$ = strTemp$ & _
           Chr(34) & "V0"     & Chr(34)  & Delim & _
           Chr(34) & " "      & Chr(34)  & Delim & _
           Chr(34) & "I0"     & Chr(34)  & Delim & _
           Chr(34) & " "      & Chr(34)  & Delim
      End If
      If OutN = 1 Then
         strTemp$ = strTemp$ & _
           Chr(34) & "V2"     & Chr(34)  & Delim & _
           Chr(34) & " "      & Chr(34)  & Delim & _
           Chr(34) & "I2"     & Chr(34)  & Delim & _
           Chr(34) & " "      & Chr(34)
      End If
      Print #FileNO, strTemp$

      If Bus2Hnd <> -1 Then 
         strTemp$ = _
           Chr(34) & "Comment" & Chr(34) & Delim & _
           Chr(34) & Bus2Name & "  ->  " & Bus1Name & Chr(34) & Delim & _
           Chr(34) & "Va"     & Chr(34)  & Delim & _
           Chr(34) & " "      & Chr(34)  & Delim & _
           Chr(34) & "Ia"     & Chr(34)  & Delim & _
           Chr(34) & " "      & Chr(34)  & Delim & _
           Chr(34) & "Vb"     & Chr(34)  & Delim & _
           Chr(34) & " "      & Chr(34)  & Delim & _
           Chr(34) & "Ib"     & Chr(34)  & Delim & _
           Chr(34) & " "      & Chr(34)  & Delim & _
           Chr(34) & "Vc"     & Chr(34)  & Delim & _
           Chr(34) & " "      & Chr(34)  & Delim & _
           Chr(34) & "Ic"     & Chr(34)  & Delim & _
           Chr(34) & " "      & Chr(34)  & Delim
        If OutZ = 1 Then
           strTemp$ = strTemp$ & _
             Chr(34) & "V0"     & Chr(34)  & Delim & _
             Chr(34) & " "      & Chr(34)  & Delim & _
             Chr(34) & "I0"     & Chr(34)  & Delim & _
             Chr(34) & " "      & Chr(34)  & Delim
        End If
        If OutN = 1 Then
           strTemp$ = strTemp$ & _
             Chr(34) & "V2"     & Chr(34)  & Delim & _
             Chr(34) & " "      & Chr(34)  & Delim & _
             Chr(34) & "I2"     & Chr(34)  & Delim & _
             Chr(34) & " "      & Chr(34)
        End If
        Print #(FileNO+4), strTemp$
      End If
   End If

   ' Prefault quantities
   ' get prefault voltage on bus 1
   If GetPSCVoltage( Bus1Hnd, MagArray, AngArray, 1 ) = 0 Then	
      Print "Get Bus 1 prefault voltage failed."
      Exit Function
   End If
   MagArray(1) = MagArray(1)*1000/PT_Ratio1
   AngArray(2) = AngArray(1) - 120	' phase B
   If AngArray(2) <=-360 Then AngArray(2) = AngArray(2) + 360
   AngArray(3) = AngArray(1) + 120	' phase C
   If AngArray(3) >=360 Then AngArray(3) = AngArray(3) - 360
   VA1$ = Format( MagArray(1), "#####0.00000") & Delim & Format( AngArray(1), "##0.0000")
   VB1$ = Format( MagArray(1), "#####0.00000") & Delim & Format( AngArray(2), "##0.0000")
   VC1$ = Format( MagArray(1), "#####0.00000") & Delim & Format( AngArray(3), "##0.0000")
   IA1$ = "0.00000" & Delim & "0.0000"

   ' get prefault voltage on bus 2
   If Bus2Hnd <> -1 Then
      If GetPSCVoltage( Bus2Hnd, MagArray, AngArray, 1 ) = 0 Then	
         Print "Get Bus 2 prefault voltage failed."
         Exit Function
      End If
      MagArray(1) = MagArray(1)*1000/PT_Ratio2
      AngArray(2) = AngArray(1) - 120	' phase B
      If AngArray(2) <=-360 Then AngArray(2) = AngArray(2) + 360
      AngArray(3) = AngArray(1) + 120	' phase C
      If AngArray(3) >=360 Then AngArray(3) = AngArray(3) - 360
      VA2$ = Format( MagArray(1), "#####0.00000") & Delim & Format( AngArray(1), "##0.0000")
      VB2$ = Format( MagArray(1), "#####0.00000") & Delim & Format( AngArray(2), "##0.0000")
      VC2$ = Format( MagArray(1), "#####0.00000") & Delim & Format( AngArray(3), "##0.0000")
   End If

   ' get prefault current on branches
   If Branch1Hnd <> -1 Then
      If GetPSCCurrent( Branch1Hnd, MagArray, AngArray, 4 ) = 0 Then 
         Print "Get Branch 1 current failed."
         Exit Function
      End If
      MagArray(1) = MagArray(1)/CT_Ratio1
      AngArray(2) = AngArray(1) - 120	' phase B
      If AngArray(2) <=-360 Then AngArray(2) = AngArray(2) + 360
      AngArray(3) = AngArray(1) + 120	' phase C
      IA1$ = Format( MagArray(1), "####0.00000") & Delim & Format( AngArray(1), "##0.0000")
      IB1$ = Format( MagArray(1), "####0.00000") & Delim & Format( AngArray(2), "##0.0000")
      IC1$ = Format( MagArray(1), "####0.00000") & Delim & Format( AngArray(3), "##0.0000")
   End If

   ' get prefault current on branches
   If Branch2Hnd <> -1 Then
      If GetPSCCurrent( Branch2Hnd, MagArray, AngArray, 4 ) = 0 Then 
         Print "Get Branch 1 current failed."
         Exit Function
      End If
      MagArray(1) = MagArray(1)/CT_Ratio2
      AngArray(2) = AngArray(1) - 120	' phase B
      If AngArray(2) <=-360 Then AngArray(2) = AngArray(2) + 360
      AngArray(3) = AngArray(1) + 120	' phase C
      IA1$ = Format( MagArray(1), "####0.00000") & Delim & Format( AngArray(1), "##0.0000")
      IB1$ = Format( MagArray(1), "####0.00000") & Delim & Format( AngArray(2), "##0.0000")
      IC1$ = Format( MagArray(1), "####0.00000") & Delim & Format( AngArray(3), "##0.0000")
   End If

   Print #FileNO, _
      Chr(34) & "  Prefault" & Chr(34) & Delim & _
      Chr(34) & Bus1Name & Chr(34) & Delim & _
      VA1 & Delim & IA1 & Delim & _
      VB1 & Delim & IA1 & Delim & _
      VC1 & Delim & IA1 & Delim & _
      IA1 & Delim & IA1 & Delim & _
      IA1 & Delim & IA1
   If Bus2Hnd <> -1 Then 
      Print #(FileNO+4), _
      Chr(34) & "  Prefault" & Chr(34) & Delim & _
      Chr(34) & Bus2Name & Chr(34) & Delim & _
      VA2 & Delim & IA1 & Delim & _
      VB2 & Delim & IA1 & Delim & _
      VC2 & Delim & IA1 & Delim & _
      IA1 & Delim & IA1 & Delim & _
      IA1 & Delim & IA1
   End If
 
   ' Loop over selected faults and export data
   Index = 1
   While FltList(Index) > -1
      If PickFault( FltList(Index) ) = 0 Then 
         TextFile = 0
         Exit Function
      End If

      ' Print out data
      ' Get bus 1 Voltage
      If GetSCVoltage( Bus1Hnd, MagArray, AngArray, 4 ) = 0 Then 
         Print "Get Bus 1 voltage failed."
         Exit Function
      End If
      For ii = 1 To 3
         MagArray(ii) = MagArray(ii)*1000/PT_Ratio1
      Next ii
      VA1$ = Format( MagArray(1), "#####0.00000") & Delim & Format( AngArray(1), "##0.0000")
      VB1$ = Format( MagArray(2), "#####0.00000") & Delim & Format( AngArray(2), "##0.0000")
      VC1$ = Format( MagArray(3), "#####0.00000") & Delim & Format( AngArray(3), "##0.0000")

      ' Get bus 2 Voltage
      If Bus2Hnd <> -1 Then
         If GetSCVoltage( Bus2Hnd, MagArray, AngArray, 4 ) = 0 Then 
            Print "Get Bus 2 voltage failed."
            Exit Function
         End If
         For ii = 1 To 3
            MagArray(ii) = MagArray(ii)*1000/PT_Ratio2
         Next ii
         VA2$ = Format( MagArray(1), "#####0.00000") & Delim & Format( AngArray(1), "##0.0000")
         VB2$ = Format( MagArray(2), "#####0.00000") & Delim & Format( AngArray(2), "##0.0000")
         VC2$ = Format( MagArray(3), "#####0.00000") & Delim & Format( AngArray(3), "##0.0000")
      End If
 
      ' Get bus 1 Voltage in sequence
      If GetSCVoltage( Bus1Hnd, MagArray, AngArray, 2 ) = 0 Then 
         Print "Get Bus 1 sequence voltage failed."
         Exit Function
      End If
      For ii = 1 To 3
         MagArray(ii) = MagArray(ii)*1000/PT_Ratio1
      Next ii
      V01$ = Format( MagArray(1), "#####0.00000") & Delim & Format( AngArray(1), "##0.0000")
      V21$ = Format( MagArray(3), "#####0.00000") & Delim & Format( AngArray(3), "##0.0000")
   
      ' Get bus 2 Voltage in sequence
      If Bus2Hnd <> -1 Then
         If GetSCVoltage( Bus2Hnd, MagArray, AngArray, 2 ) = 0 Then 
            Print "Get Bus 2 sequence voltage failed."
            Exit Function
         End If
         For ii = 1 To 3
            MagArray(ii) = MagArray(ii)*1000/PT_Ratio2
         Next ii
         V02$ = Format( MagArray(1), "#####0.00000") & Delim & Format( AngArray(1), "##0.0000")
         V22$ = Format( MagArray(3), "#####0.00000") & Delim & Format( AngArray(3), "##0.0000")
      End If
 
      ' Get branch 1 current
      If GetSCCurrent( Branch1Hnd, MagArray, AngArray, 4 ) = 0 Then 
         Print "Get Branch 1 current failed."
         Exit Function
      End If
      For ii = 1 To 3
         MagArray(ii) = MagArray(ii)/CT_Ratio1
      Next ii
      IA1$ = Format( MagArray(1), "####0.00000") & Delim & Format( AngArray(1), "##0.0000")
      IB1$ = Format( MagArray(2), "####0.00000") & Delim & Format( AngArray(2), "##0.0000")
      IC1$ = Format( MagArray(3), "####0.00000") & Delim & Format( AngArray(3), "##0.0000")

      ' Get branch 2 current
      If Branch2Hnd <> -1 Then
         If GetSCCurrent( Branch2Hnd, MagArray, AngArray, 4 ) = 0 Then 
            Print "Get Branch 2 current failed."
            Exit Function
         End If
         For ii = 1 To 3
            MagArray(ii) = MagArray(ii)/CT_Ratio2
         Next ii
         IA2$ = Format( MagArray(1), "####0.00000") & Delim & Format( AngArray(1), "##0.0000")
         IB2$ = Format( MagArray(2), "####0.00000") & Delim & Format( AngArray(2), "##0.0000")
         IC2$ = Format( MagArray(3), "####0.00000") & Delim & Format( AngArray(3), "##0.0000")
      End If

      ' Get branch 1 current in sequence
      If GetSCCurrent( Branch1Hnd, MagArray, AngArray, 2 ) = 0 Then 
         Print "Get Branch 1 sequence current failed."
         Exit Function
      End If
      For ii = 1 To 3
         MagArray(ii) = MagArray(ii)/CT_Ratio1
      Next ii
      I01$ = Format( MagArray(1), "####0.00000") & Delim & Format( AngArray(1), "##0.0000")
      I21$ = Format( MagArray(3), "####0.00000") & Delim & Format( AngArray(3), "##0.0000")
   
      ' Get branch 2 current in sequence
      If Branch2Hnd <> -1 Then
         If GetSCCurrent( Branch2Hnd, MagArray, AngArray, 2 ) = 0 Then 
            Print "Get Branch 2 sequence current failed."
            Exit Function
         End If
         For ii = 1 To 3
            MagArray(ii) = MagArray(ii)/CT_Ratio2
         Next ii
         I02$ = Format( MagArray(1), "####0.00000") & Delim & Format( AngArray(1), "##0.0000")
         I22$ = Format( MagArray(3), "####0.00000") & Delim & Format( AngArray(3), "##0.0000")
      End If
      strText$ = _
         Chr(34) & FltDesc(Index) & Chr(34) & Delim & _
         Chr(34) & Bus1Name & Chr(34) & Delim & _
         VA1 & Delim & IA1 & Delim & _
         VB1 & Delim & IB1 & Delim & _
         VC1 & Delim & IC1 & Delim
      If OutZ Then strText$ = strText$ & V01 & Delim & I01 & Delim
      If OutN Then strText$ = strText$ & V21 & Delim & I21
      Print #FileNO, strText$
      If Bus2Hnd <> -1 Then 
         strText$ = _
         Chr(34) & FltDesc(Index) & Chr(34) & Delim & _
         Chr(34) & Bus2Name & Chr(34) & Delim & _
         VA2 & Delim & IA2 & Delim & _
         VB2 & Delim & IB2 & Delim & _
         VC2 & Delim & IC2 & Delim
         If OutZ Then strText$ = strText$ & V02 & Delim & I02 & Delim
         If OutN Then strText$ = strText$ & V22 & Delim & I22
         Print #(FileNO+4), strText$
      End If
      Index = Index + 1
   Wend
   TextFile = 1
End Function
'=================================End of TextFile()====================================


'=====================================Doble()====================================
' Purpose:
'   Print out fault result to Doble file
'
'======================================================================================
Function Doble( ByVal Bus1Hnd As Long, ByVal Bus2Hnd As Long, ByVal Branch1Hnd As Long, ByVal Branch2Hnd As Long, _
         ByRef FltList() As Long, ByRef FltDesc() As String, ByRef CT_Ratio1 As Double, ByRef CT_Ratio2 As Double, _
         ByRef PT_Ratio1 As Double, ByRef PT_Ratio2 As Double, ByRef TimePeriod() As Double ) As Long

   Dim MagArray1(3) As Double
   Dim AngArray1(3) As Double
   Dim MagArray2(3) As Double
   Dim AngArray2(3) As Double
   Dim MagArray3(3) As Double
   Dim AngArray3(3) As Double
   Dim MagArray4(3) As Double
   Dim AngArray4(3) As Double
   Dim MagArray(16) As Double
   Dim AngArray(16) As Double

   Doble = 0
   Bus1Name$ = FullBusName( Bus1Hnd )	' near bus name
   If Bus2Hnd <> -1 Then
      Bus2Name$ = FullBusName( Bus2Hnd )	' far bus name
   End If
   CurrentDate$ = Date()
   CurrentTime$ = Time()
   ' Get the number of simulations
   Index = 1
   While FltList(Index) > -1
      Index = Index + 1
   Wend

   ' Print column header if overwrite
   Print #3, _
      "VER;" & "0003" & Chr(13) & Chr(10) & _
      Chr(34) & Bus1Name & Chr(34) & " file created by SS1 Version 3 -> " & CurrentDate & " " & CurrentTime & Chr(13) & Chr(10) & _
      Format( Index-1, "##0" )
         
   If Bus2Hnd <> -1 Then	' line 
      Print #7, _		
      "VER;" & "0003" & Chr(13) & Chr(10) & _
      Chr(34) & Bus2Name & Chr(34) & " file created by SS1 Version 3 -> " & CurrentDate & " " & CurrentTime & Chr(13) & Chr(10) & _
      Format( Index-1, "##0" )
   End If

   ' Prefault quantities
   ' get prefault voltage on bus 1
   If GetPSCVoltage( Bus1Hnd, MagArray1, AngArray1, 1 ) = 0 Then	
      Print "Get Bus 1 prefault voltage failed."
      Exit Sub
   End If
   MagArray1(1) = MagArray1(1)*1000/PT_Ratio1	' V
   AngArray1(2) = AngArray1(1) - 120	' phase B
   If AngArray1(2) <=-360 Then AngArray1(2) = AngArray1(2) + 360
   AngArray1(3) = AngArray1(1) + 120	' phase C
   If AngArray1(3) >=360 Then AngArray1(3) = AngArray1(3) - 360
   
   If Bus2Hnd <> -1 Then
      ' get prefault voltage on bus 2
      If GetPSCVoltage( Bus2Hnd, MagArray2, AngArray2, 1 ) = 0 Then	
         Print "Get Bus 2 prefault voltage failed."
         Exit Function
      End If
      MagArray2(1) = MagArray2(1)*1000/PT_Ratio2	' V
      AngArray2(2) = AngArray2(1) - 120	' phase B
      If AngArray2(2) <=-360 Then AngArray2(2) = AngArray2(2) + 360
      AngArray2(3) = AngArray2(1) + 120	' phase C
      If AngArray2(3) >=360 Then AngArray2(3) = AngArray2(3) - 360
   End If

   ' get prefault current on branches
   If Branch1Hnd <> -1 Then
      If GetPSCCurrent( Branch1Hnd, MagArray3, AngArray3, 4 ) = 0 Then 
         Print "Get Branch 1 current failed."
         Exit Function
      End If
      MagArray3(1) = MagArray3(1)/CT_Ratio1
      AngArray3(2) = AngArray3(1) - 120	' phase B
      If AngArray3(2) <=-360 Then AngArray3(2) = AngArray3(2) + 360
      AngArray3(3) = AngArray3(1) + 120	' phase C
   End If

   ' get prefault current on branches
   If Branch2Hnd <> -1 Then
      If GetPSCCurrent( Branch2Hnd, MagArray4, AngArray4, 4 ) = 0 Then 
         Print "Get Branch 2 current failed."
         Exit Function
      End If
      MagArray4(1) = MagArray4(1)/CT_Ratio1
      AngArray4(2) = AngArray4(1) - 120	' phase B
      If AngArray4(2) <=-360 Then AngArray4(2) = AngArray4(2) + 360
      AngArray4(3) = AngArray4(1) + 120	' phase C
   End If

   ' Loop over selected faults and export data
   Index = 1
   While FltList(Index) > -1
      If PickFault( FltList(Index) ) = 0 Then 
         Doble = 0
         Exit Function
      End If
      
      ' Print out data
      ' Get bus 1 Voltage
      If GetSCVoltage( Bus1Hnd, MagArray, AngArray, 4 ) = 0 Then 
         Print "Get Bus 1 voltage failed."
         Exit Function
      End If
      For ii = 1 To 3
         MagArray(ii) = MagArray(ii)*1000/PT_Ratio1	' V
      Next ii
      VA1$ = Format( MagArray(1), "#####0.00000") & Chr(13) & Chr(10) & Format( AngArray(1), "##0.0000")
      VB1$ = Format( MagArray(2), "#####0.00000") & Chr(13) & Chr(10) & Format( AngArray(2), "##0.0000")
      VC1$ = Format( MagArray(3), "#####0.00000") & Chr(13) & Chr(10) & Format( AngArray(3), "##0.0000")

      If Bus2Hnd <> -1 Then	' line
         ' Get bus 2 Voltage
         If GetSCVoltage( Bus2Hnd, MagArray, AngArray, 4 ) = 0 Then 
            Print "Get Bus 2 voltage failed."
            Exit Function
         End If
         For ii = 1 To 3
            MagArray(ii) = MagArray(ii)*1000/PT_Ratio2	' V
         Next ii
         VA2$ = Format( MagArray(1), "#####0.00000") & Chr(13) & Chr(10) & Format( AngArray(1), "##0.0000")
         VB2$ = Format( MagArray(2), "#####0.00000") & Chr(13) & Chr(10) & Format( AngArray(2), "##0.0000")
         VC2$ = Format( MagArray(3), "#####0.00000") & Chr(13) & Chr(10) & Format( AngArray(3), "##0.0000")
      End If
 
      ' Get bus 1 Voltage in sequence
      If GetSCVoltage( Bus1Hnd, MagArray, AngArray, 2 ) = 0 Then 
         Print "Get Bus 1 sequence voltage failed."
         Exit Function
      End If
      For ii = 1 To 3
         MagArray(ii) = MagArray(ii)*1000/PT_Ratio1
      Next ii
      V01$ = Format( MagArray(1), "#####0.00000") & Chr(13) & Chr(10) & Format( AngArray(1), "##0.0000")
      V21$ = Format( MagArray(3), "#####0.00000") & Chr(13) & Chr(10) & Format( AngArray(3), "##0.0000")
   
      If Bus2Hnd <> -1 Then
         ' Get bus 2 Voltage in sequence
         If GetSCVoltage( Bus2Hnd, MagArray, AngArray, 2 ) = 0 Then 
            Print "Get Bus 2 sequence voltage failed."
            Exit Function
         End If
         For ii = 1 To 3
            MagArray(ii) = MagArray(ii)*1000/PT_Ratio2
         Next ii
         V02$ = Format( MagArray(1), "#####0.00000") & Chr(13) & Chr(10) & Format( AngArray(1), "##0.0000")
         V22$ = Format( MagArray(3), "#####0.00000") & Chr(13) & Chr(10) & Format( AngArray(3), "##0.0000")
      End If
 
      ' Get branch 1 current
      If GetSCCurrent( Branch1Hnd, MagArray, AngArray, 4 ) = 0 Then 
         Print "Get Branch 1 current failed."
         Exit Function
      End If
      For ii = 1 To 3
         MagArray(ii) = MagArray(ii)/CT_Ratio1
      Next ii
      IA1$ = Format( MagArray(1), "####0.00000") & Chr(13) & Chr(10) & Format( AngArray(1), "##0.0000")
      IB1$ = Format( MagArray(2), "####0.00000") & Chr(13) & Chr(10) & Format( AngArray(2), "##0.0000")
      IC1$ = Format( MagArray(3), "####0.00000") & Chr(13) & Chr(10) & Format( AngArray(3), "##0.0000")

      If Branch2Hnd <> -1 Then
         ' Get branch 2 current
         If GetSCCurrent( Branch2Hnd, MagArray, AngArray, 4 ) = 0 Then 
            Print "Get Branch 2 current failed."
            Exit Function
         End If
         For ii = 1 To 3
            MagArray(ii) = MagArray(ii)/CT_Ratio2
         Next ii
         IA2$ = Format( MagArray(1), "####0.00000") & Chr(13) & Chr(10) & Format( AngArray(1), "##0.0000")
         IB2$ = Format( MagArray(2), "####0.00000") & Chr(13) & Chr(10) & Format( AngArray(2), "##0.0000")
         IC2$ = Format( MagArray(3), "####0.00000") & Chr(13) & Chr(10) & Format( AngArray(3), "##0.0000")
      End If

      ' Get branch 1 current in sequence
      If GetSCCurrent( Branch1Hnd, MagArray, AngArray, 2 ) = 0 Then 
         Print "Get Branch 1 sequence current failed."
         Exit Function
      End If
      For ii = 1 To 3
         MagArray(ii) = MagArray(ii)/CT_Ratio1
      Next ii
      I01$ = Format( MagArray(1), "####0.00000") & Chr(13) & Chr(10) & Format( AngArray(1), "##0.0000")
      I21$ = Format( MagArray(3), "####0.00000") & Chr(13) & Chr(10) & Format( AngArray(3), "##0.0000")
   
      If Branch2Hnd <> -1 Then
         ' Get branch 2 current in sequence
         If GetSCCurrent( Branch2Hnd, MagArray, AngArray, 2 ) = 0 Then 
            Print "Get Branch 2 sequence current failed."
            Exit Function
         End If
         For ii = 1 To 3
            MagArray(ii) = MagArray(ii)/CT_Ratio2
         Next ii
         I02$ = Format( MagArray(1), "####0.00000") & Chr(13) & Chr(10) & Format( AngArray(1), "##0.0000")
         I22$ = Format( MagArray(3), "####0.00000") & Chr(13) & Chr(10) & Format( AngArray(3), "##0.0000")
      End If

      nCount& = 6
      If OutZ Then nCount& = nCount& + 2
      If OutN Then nCount& = nCount& + 2
      ' Print simulation information
      Print #3, Right$( FltDesc(Index), Len( FltDesc(Index) ) - 2 )	' Fault description
      Print #3, "Y" & Chr(13) & Chr(10) & "0" & Chr(13) & Chr(10) & "0" & Chr(13) & Chr(10) & "0" 
      Print #3, " "
      Print #3, "V" & Chr(13) & Chr(10) & "V" & Chr(13) & Chr(10) & "60.00000" & Chr(13) & Chr(10) & Str( nCount& ) & Chr(13) & Chr(10) & "3"
      Print #3, "VA" & Chr(13) & Chr(10) & "1" & Chr(13) & Chr(10) & "VB" & Chr(13) & Chr(10) & "1" 
      Print #3, "VC" & Chr(13) & Chr(10) & "1"	' source name
      Print #3, "IA" & Chr(13) & Chr(10) & "1" & Chr(13) & Chr(10) & "IB" & Chr(13) & Chr(10) & "1"
      Print #3, "IC" & Chr(13) & Chr(10) & "1"
      If OutZ = 1 Then _
        Print #3, "V0" & Chr(13) & Chr(10) & "1" & Chr(13) & Chr(10) & "I0" & Chr(13) & Chr(10) & "1"
      If OutN = 1 Then _ 
        Print #3, "V2" & Chr(13) & Chr(10) & "1" & Chr(13) & Chr(10) & "I2" & Chr(13) & Chr(10) & "1"
      ' Prefault
      Print #3, "Pre-Fault" & Chr(13) & Chr(10) & "TIME"
      ' Voltage
      Print #3, Format( MagArray1(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray1(1), "##0.0000")
      Print #3, Format( MagArray1(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray1(2), "##0.0000")
      Print #3, Format( MagArray1(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray1(3), "##0.0000")
      ' Current
      Print #3, Format( MagArray3(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray3(1), "##0.0000")
      Print #3, Format( MagArray3(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray3(2), "##0.0000")
      Print #3, Format( MagArray3(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray3(3), "##0.0000")
      ' V0, I0
      If OutZ = 1 Then
        Print #3, "0.00000" & Chr(13) & Chr(10) & "0.0000"
        Print #3, "0.00000" & Chr(13) & Chr(10) & "0.0000"
      End If
      ' V2, I2
      If OutN = 1 Then
        Print #3, "0.00000" & Chr(13) & Chr(10) & "0.0000"
        Print #3, "0.00000" & Chr(13) & Chr(10) & "0.0000"
      End If
      Print #3, Format( TimePeriod(1), "#####0.000" ) & Chr(13) & Chr(10) & "60.00000" & Chr(13) & Chr(10) & "60.00000"
      ' Fault
      Print #3, "Fault" & Chr(13) & Chr(10) & "TIME"
      Print #3, VA1 & Chr(13) & Chr(10) & VB1 & Chr(13) & Chr(10) & VC1
      Print #3, IA1 & Chr(13) & Chr(10) & IB1 & Chr(13) & Chr(10) & IC1
      If OutZ = 1 Then Print #3, V01 & Chr(13) & Chr(10) & I01
      If OutN = 1 Then Print #3, V21 & Chr(13) & Chr(10) & I21
      Print #3, Format( TimePeriod(2), "#####0.000" ) & Chr(13) & Chr(10) & "60.00000" & Chr(13) & Chr(10) & "60.00000"
      ' Postfault
      Print #3, "Post-Fault" & Chr(13) & Chr(10) & "TIME"
      ' Voltage
      Print #3, Format( MagArray1(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray1(1), "##0.0000")
      Print #3, Format( MagArray1(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray1(2), "##0.0000")
      Print #3, Format( MagArray1(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray1(3), "##0.0000")
      ' Current
      Print #3, Format( MagArray3(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray3(1), "##0.0000")
      Print #3, Format( MagArray3(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray3(2), "##0.0000")
      Print #3, Format( MagArray3(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray3(3), "##0.0000")
      ' V0, I0
      If OutZ = 1 Then 
        Print #3, "0.00000" & Chr(13) & Chr(10) & "0.0000"
        Print #3, "0.00000" & Chr(13) & Chr(10) & "0.0000"
      End If
      ' V2, I2
      If OutN = 1 Then
        Print #3, "0.00000" & Chr(13) & Chr(10) & "0.0000"
        Print #3, "0.00000" & Chr(13) & Chr(10) & "0.0000"
      End If
      Print #3, Format( TimePeriod(3), "#####0.000" ) & Chr(13) & Chr(10) & "60.00000" & Chr(13) & Chr(10) & "60.00000"
      ' Timer section
      ' Voltage
      Print #3, "1" & Chr(13) & Chr(10) & "VA" & Chr(13) & Chr(10) & "C" & Chr(13) & Chr(10) & "O" & Chr(13) & Chr(10) & "P"
      Print #3, "0.00000 " & CurrentDate & " " & CurrentTime
      Print #3, "MS"
      ' Current
      Print #3, "1" & Chr(13) & Chr(10) & "IA" & Chr(13) & Chr(10) & "C" & Chr(13) & Chr(10) & "O" & Chr(13) & Chr(10) & "T"
      Print #3, "0.00000 " & CurrentDate & " " & CurrentTime
      Print #3, "MS"
      ' V0, I0
      If OutZ = 1 Then
        Print #3, "1" & Chr(13) & Chr(10) & "I0" & Chr(13) & Chr(10) & "C" & Chr(13) & Chr(10) & "O" & Chr(13) & Chr(10) & "T"
        Print #3, "0.00000 " & CurrentDate & " " & CurrentTime
        Print #3, "MS"
      End If
  
      If Bus2Hnd <> -1 Then  	' line  
         Print #7, Right$( FltDesc(Index), Len( FltDesc(Index) ) - 2 )	' Fault description
         Print #7, "Y" & Chr(13) & Chr(10) & "0" & Chr(13) & Chr(10) & "0" & Chr(13) & Chr(10) & "0" 
         Print #7, " "
         Print #7, "V" & Chr(13) & Chr(10) & "V" & Chr(13) & Chr(10) & "60.00000" & Chr(13) & Chr(10) & Str(nCount&) & Chr(13) & Chr(10) & "3"
         Print #7, "VA" & Chr(13) & Chr(10) & "1" & Chr(13) & Chr(10) & "VB" & Chr(13) & Chr(10) & "1" 
         Print #7, "VC" & Chr(13) & Chr(10) & "1"	' source name
         Print #7, "IA" & Chr(13) & Chr(10) & "1" & Chr(13) & Chr(10) & "IB" & Chr(13) & Chr(10) & "1"
         Print #7, "IC" & Chr(13) & Chr(10) & "1"
         If OutZ = 1 Then _
           Print #7, "V0" & Chr(13) & Chr(10) & "1" & Chr(13) & Chr(10) & "I0" & Chr(13) & Chr(10) & "1"
         If OutN Then _
           Print #7, "V2" & Chr(13) & Chr(10) & "1" & Chr(13) & Chr(10) & "I2" & Chr(13) & Chr(10) & "1"
         ' Prefault
         Print #7, "Pre-Fault" & Chr(13) & Chr(10) & "TIME"
         ' Voltage
         Print #7, Format( MagArray2(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray2(1), "##0.0000")
         Print #7, Format( MagArray2(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray2(2), "##0.0000")
         Print #7, Format( MagArray2(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray2(3), "##0.0000")
         ' Current
         Print #7, Format( MagArray4(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray4(1), "##0.0000")
         Print #7, Format( MagArray4(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray4(2), "##0.0000")
         Print #7, Format( MagArray4(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray4(3), "##0.0000")
         ' V0, I0
         If OutZ = 1 Then
           Print #7, "0.00000" & Chr(13) & Chr(10) & "0.0000"
           Print #7, "0.00000" & Chr(13) & Chr(10) & "0.0000"
         End If
         ' V2, I2
         If OutN = 1 Then
           Print #7, "0.00000" & Chr(13) & Chr(10) & "0.0000"
           Print #7, "0.00000" & Chr(13) & Chr(10) & "0.0000"
         End If
         Print #7, Format( TimePeriod(1), "#####0.000" ) & Chr(13) & Chr(10) & "60.00000" & Chr(13) & Chr(10) & "60.00000"
         ' Fault
         Print #7, "Fault" & Chr(13) & Chr(10) & "TIME"
         Print #7, VA2 & Chr(13) & Chr(10) & VB2 & Chr(13) & Chr(10) & VC2
         Print #7, IA2 & Chr(13) & Chr(10) & IB2 & Chr(13) & Chr(10) & IC2
         If OutZ = 1 Then _
           Print #7, V02 & Chr(13) & Chr(10) & I02
         If OutN = 1 Then _
           Print #7, V22 & Chr(13) & Chr(10) & I22
         Print #7, Format( TimePeriod(2), "#####0.000" ) & Chr(13) & Chr(10) & "60.00000" & Chr(13) & Chr(10) & "60.00000"
         ' Postfault
         Print #7, "Post-Fault" & Chr(13) & Chr(10) & "TIME"
         ' Voltage
         Print #7, Format( MagArray2(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray2(1), "##0.0000")
         Print #7, Format( MagArray2(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray2(2), "##0.0000")
         Print #7, Format( MagArray2(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray2(3), "##0.0000")
         ' Current
         Print #7, Format( MagArray4(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray4(1), "##0.0000")
         Print #7, Format( MagArray4(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray4(2), "##0.0000")
         Print #7, Format( MagArray4(1), "#####0.00000" ) & Chr(13) & Chr(10) & Format( AngArray4(3), "##0.0000")
         ' V0, I0
         If OutZ = 1 Then
           Print #7, "0.00000" & Chr(13) & Chr(10) & "0.0000"
           Print #7, "0.00000" & Chr(13) & Chr(10) & "0.0000"
         End If
         ' V2, I2
         If OutN = 1 Then 
           Print #7, "0.00000" & Chr(13) & Chr(10) & "0.0000"
           Print #7, "0.00000" & Chr(13) & Chr(10) & "0.0000"
         End If
         Print #7, Format( TimePeriod(3), "#####0.000" ) & Chr(13) & Chr(10) & "60.00000" & Chr(13) & Chr(10) & "60.00000"
         ' Timer section
         ' Voltage
         Print #7, "1" & Chr(13) & Chr(10) & "VA" & Chr(13) & Chr(10) & "C" & Chr(13) & Chr(10) & "O" & Chr(13) & Chr(10) & "P"
         Print #7, "0.00000 " & CurrentDate & " " & CurrentTime
         Print #7, "MS"
         ' Current
         Print #7, "1" & Chr(13) & Chr(10) & "IA" & Chr(13) & Chr(10) & "C" & Chr(13) & Chr(10) & "O" & Chr(13) & Chr(10) & "T"
         Print #7, "0.00000 " & CurrentDate & " " & CurrentTime
         Print #7, "MS"
         ' V0, I0
         If OutZ = 1 Then
           Print #7, "1" & Chr(13) & Chr(10) & "I0" & Chr(13) & Chr(10) & "C" & Chr(13) & Chr(10) & "O" & Chr(13) & Chr(10) & "T"
           Print #7, "0.00000 " & CurrentDate & " " & CurrentTime
           Print #7, "MS"
         End If
      End If

      Index = Index + 1
   Wend
   Doble = 1
   Print #3, "**XX**F"	' End of File Marker
   If Bus2Hnd <> -1 Then Print #7, "**XX**F"
End Function
'=================================End of Doble()====================================


'=====================================KeyFile()====================================
' Purpose:
'   Print out the key file associated with ss1 file
'
'======================================================================================
Function KeyFile( ByVal FileName As String, ByVal CT_Ratio As Double, ByVal PT_Ratio As Double, ByVal kVNominal As Double, _
                  ByVal Bus1No As Long, ByVal Bus2No As Long, ByVal Bus1Name As String, ByVal Bus2Name As String, _
                  ByVal LineID As Long, ByRef FltList() As Long, ByRef FltDesc() As String, ByVal FileNo As Long ) As Long

   Dim TempString As String, Pos As Long

   KeyFile = 0
   CurrentDate$ = Date()
   CurrentTime$ = Time()

   Print #FileNo, "This is the key file associated with SS1 file: " & FileName
   Print #FileNo, "Created -> " & CurrentDate & " " & CurrentTime
   Print #FileNo, " "
   Print #FileNo, "SS1 data created for branch " & Format( Bus1No, "0" ) & " " & Format( Bus2No, "0" ) _
         & " " & Format( LineID, "0" )
   Print #FileNo, "CT ratio used for IA, IB, IC, I0, I2:  " & Format( CT_Ratio, "0" )
   Print #FileNo, "VT ratio used for VA, VB, VC, V0, V2:  " & Format( PT_Ratio, "0" )
   Print #FileNo, "Voltage base used:  " & Format( kVNominal, "0" ) & "kV"
   Print #FileNo, " "
   Print #FileNo, " "
   Print #FileNo, "The following information describes the buses in " & FileName
   Print #FileNo, " "
   Print #FileNo, "Bus #     Bus name               Bus kV     Substation name"
   Print #FileNo, Format( Bus1No, "0" ) & "        " & Bus1Name & "       " & Format( kVNominal, "0" )
   Print #FileNo, Format( Bus2No, "0" ) & "        " & Bus2Name & "       " & Format( kVNominal, "0" )
   Print #FileNo, "------------------------------------------------------------------------------------------"
   Print #FileNo, " "
   Print #FileNo, "The following data describes the 1 current(s) summed into the polarizing"
   Print #FileNo, "circuit to produce the current source (H1) used in " & FileName
   Print #FileNo, " "
   Print #FileNo, "From      To        Circuit     CT        Base"
   Print #FileNo, "Bus #     Bus #     #           Ratio     kV"
   Print #FileNo, "------------------------------------------------------------------------------------------"
   Print #FileNo, Format( Bus1No, "0" ) & "         " & Format( Bus2No, "0" ) & "         " & Format( LineID, "0" ) & _
         "           " & Format( CT_Ratio, "0" ) & "         " & Format( kVNominal, "0" )
   Print #FileNo, "------------------------------------------------------------------------------------------"
   Print #FileNo, " "
   Print #FileNo, "The following are the fault locations:"
   Print #FileNo, " "
   Print #FileNo, "<------------------------- Actual fault location ----------------------><---- User input data -------->"
   Print #FileNo, "TYPE        FROM BUS         TO BUS          CIRCUIT        % DIST        A END         B END"     
   Print #FileNo, "-------------------------------------------------------------------------------------------------------"

   ' Loop over selected faults and export data
   Index = 1
   While FltList(Index) > -1
      If PickFault( FltList(Index) ) = 0 Then 
         KeyFile = 0
         Exit Function
      End If
      Pos = InStr( 1, FltDesc(Index), "(" )
      'Print Format(Pos,"0") & FltList(Index)
      TempString = Mid( FltDesc(Index), Pos+1, 6 )
      Print #FileNo, "LINE        " & Format( Bus1No, "0" ) & "                " & Format( Bus2No, "0" ) & _
         "               " & Format( LineID, "0" ) & "              " & TempString & "        " & Format( Bus1No, "0" ) & _
         "             " & Format( Bus2No, "0" )
      Index = Index + 1
   Wend
   Print #FileNo, "-------------------------------------------------------------------------------------------------------"
   KeyFile = 1
End Function
'=================================End of KeyFile()====================================


'=====================================Comtrade()====================================
' Purpose:
'   Print out fault result to Comtrade file (*.CFG, *.DAT)
'
'======================================================================================
Function Comtrade( ByVal BusHnd As Long, ByVal BranchHnd As Long, ByRef FltList() As Long, _
         ByRef FaultDesc() As String, ByRef CT_Ratio As Double, ByRef PT_Ratio As Double, _
         ByVal FileNo As Long, ByRef Cycles() As Double, _
         ByRef Comments As String ) As Long

   Dim MagArray1(3) As Double, AngArray1(3) As Double
   Dim MagArray3(3) As Double, AngArray3(3) As Double
   Dim MagArray(16) As Double, AngArray(16) As Double
   Dim MagArray2(10) As Double, AngArray2(10) As Double
   Dim DataValue(10) As Long
   Dim BusNumber As Long, BusName As String, BusNominal As Double
   Dim aV As Double, bV As Double, aI As Double, bI As Double
   Dim PtCount As Long, Point As Long, Omeg As Double, Sample As Long, PtNo As Long

   Comtrade = 0
    
   If GetData( BusHnd, BUS_sName, BusName ) = 0 Then
      Print "Get Bus name failed."
      Exit Function
   End If
   If GetData( BusHnd, BUS_nNumber, BusNumber ) = 0 Then
      Print "Get Bus number failed."
      Exit Function
   End If
   If GetData( BusHnd, BUS_dKVnorminal, BusNominal ) = 0 Then
      Print "Get Bus nominal voltage failed."
      Exit Function
   End If

   ' Select the first fault and export data
   If PickFault( FltList(1) ) = 0 Then 
      Print "No fault is simulated"
      Exit Function
   End If

   ' Calculate conversion factors a & b for V & I
   ' The maximum voltage is 1.3*Vnominal, the maximum current is 25*Irate
   ' Assume 12-bit AD converter (+2047, -2048)
   ' PtNo samples per cycle
   aV = Sqr(2)*1.3*( BusNominal*1000/PT_Ratio )/( 2047+2048 )
   bV = 0.0
   aI = Sqr(2)*25*5/( 2047+2048 )
   bI = 0.0
   PtNo =100
   Sample = 60*PtNo

   If OutZ = 1 And OutN = 1 Then sCount$ = "10,10A,0D" _
   Else If OutZ <> 1 And OutN <>1 Then sCount$ = "6,6A,0D" _
   Else sCount$ = "8,8A,0D"

   ' Print out results to *.CFG
   Print #(FileNo+1), BusName & "," & BusNumber & ",1999"
   Print #(FileNo+1), sCount$ '"10,10A,0D"
   Print #(FileNo+1), "1," & BusName & " Va-g,,,V," & Format( aV, "#####0.0000000000" ) & "," & _
      Format(bV, "#####0.0000000000") & ",0,-2048,2047," & Format( PT_Ratio, "#####0.000" ) & ",1,S"
   Print #(FileNo+1), "2," & BusName & " Vb-g,,,V," & Format( aV, "#####0.0000000000" ) & "," & _
      Format(bV, "#####0.0000000000") & ",0,-2048,2047," & Format( PT_Ratio, "#####0.000" ) & ",1,S"
   Print #(FileNo+1), "3," & BusName & " Vc-g,,,V," & Format( aV, "#####0.0000000000" ) & "," & _
      Format(bV, "#####0.0000000000") & ",0,-2048,2047," & Format( PT_Ratio, "#####0.000" ) & ",1,S"
   Print #(FileNo+1), "4," & BusName & " Ia,,,A," & Format( aI, "#####0.0000000000" ) & "," & _
      Format(bI, "#####0.0000000000") & ",0,-2048,2047," & Format( CT_Ratio, "#####0.000" ) & ",1,S"
   Print #(FileNo+1), "5," & BusName & " Ib,,,A," & Format( aI, "#####0.0000000000" ) & "," & _
      Format(bI, "#####0.0000000000") & ",0,-2048,2047," & Format( CT_Ratio, "#####0.000" ) & ",1,S"
   Print #(FileNo+1), "6," & BusName & " Ic,,,A," & Format( aI, "#####0.0000000000" ) & "," & _
      Format(bI, "#####0.0000000000") & ",0,-2048,2047," & Format( CT_Ratio, "#####0.000" ) & ",1,S"
   If OutZ = 1 Then 
     Print #(FileNo+1), "7," & BusName & " V0,,,V," & Format( aV, "#####0.0000000000" ) & "," & _
      Format(bV, "#####0.0000000000") & ",0,-2048,2047," & Format( PT_Ratio, "#####0.000" ) & ",1,S"
     Print #(FileNo+1), "8," & BusName & " I0,,,A," & Format( aI, "#####0.0000000000" ) & "," & _
      Format(bI, "#####0.0000000000") & ",0,-2048,2047," & Format( CT_Ratio, "#####0.000" ) & ",1,S"
   End If
   If OutN = 1 Then 
     Print #(FileNo+1), "9," & BusName & " V2,,,V," & Format( aV, "#####0.0000000000" ) & "," & _
      Format(bV, "#####0.0000000000") & ",0,-2048,2047," & Format( PT_Ratio, "#####0.000" ) & ",1,S"
     Print #(FileNo+1), "10," & BusName & " I2,,,A," & Format( aI, "#####0.0000000000" ) & "," & _
      Format(bI, "#####0.0000000000") & ",0,-2048,2047," & Format( CT_Ratio, "#####0.000" ) & ",1,S"
   End If
   Print #(FileNo+1), "60"
   Print #(FileNo+1), "1"
   Print #(FileNo+1), Format( Sample, "##0000.000" ) & "," & _
          Format( PtNo*(Cycles(1)+Cycles(2)+Cycles(3)+Cycles(4)), "0" )
   
   CurrentDate$ = Date()
   CurrentTime$ = Time()
   Print #(FileNo+1), Format( Day(CurrentDate), "00" ) & "/" & Format( Month(CurrentDate), "00" ) & _
      "/" & Format( Year(CurrentDate), "0000" ) & "," & Format( Hour(CurrentTime), "00" ) & ":" & _
      Format( Minute(CurrentTime), "00" ) & ":" & Format( Second(CurrentTime), "00" ) & ".000000"
   ' Non-critical, just assign a value
   Print #(FileNo+1), Format( Day(CurrentDate), "00" ) & "/" & Format( Month(CurrentDate), "00" ) & _
      "/" & Format( Year(CurrentDate), "0000" ) & "," & Format( Hour(CurrentTime), "00" ) & ":" & _
      Format( Minute(CurrentTime), "00" ) & ":" & Format( Second(CurrentTime), "00" ) & ".100000"
   Print #(FileNo+1), "ASCII"
   Print #(FileNo+1), "1"	' timemult

   ' Print our results to *.HDR
   Print #(FileNo+2), "*** Fault simulation result from ASPEN OneLiner ***"
   Print #(FileNo+2), "* Date/Time: " & CurrentDate & " " & CurrentTime
   Print #(FileNo+2), "* Fault description: " & FaultDesc(1)
   Print #(FileNo+2), "* Bus name: " & BusName & " " & Str(BusNominal) & "kV"
   Print #(FileNo+2), "* Comments: " & Comments

   ' Print out results to *.DAT
   ' Prefault voltage quantities
   If GetPSCVoltage( BusHnd, MagArray1, AngArray1, 1 ) = 0 Then	' get prefault voltage on bus
      Print "Get Bus prefault voltage failed."
      Exit Function
   End If
   MagArray1(1) = Sqr(2)*MagArray1(1)*1000/PT_Ratio
   AngArray1(1) = AngArray1(1)*pi/180
   AngArray1(2) = AngArray1(1) - 2*pi/3	' phase B
   If AngArray1(2) <=-2*pi Then AngArray1(2) = AngArray1(2) + 2*pi
   AngArray1(3) = AngArray1(1) + 2*pi/3	' phase C
   If AngArray1(3) >=2*pi  Then AngArray1(3) = AngArray1(3) - 2*pi
   ' Prefault current quantities
   If GetPSCCurrent( BranchHnd, MagArray3, AngArray3, 4 ) = 0 Then 
      Print "Get Branch current failed."
      Exit Function
   End If
   MagArray3(1) = Sqr(2)*MagArray3(1)/CT_Ratio
   AngArray3(1) = AngArray3(1)*pi/180
   AngArray3(2) = AngArray3(1) - 2*pi/3	' phase B
   If AngArray3(2) <=-2*pi Then AngArray3(2) = AngArray3(2) + 2*pi
   AngArray3(3) = AngArray3(1) + 2*pi/3	' phase C
   If AngArray3(3) >=2*pi  Then AngArray3(3) = AngArray3(3) - 2*pi

   PtCount = 1
   Omeg = 120*pi/Sample	' PtNo pts per cycle 

   ' Fault connection
   FaultString$ = FaultDescription()
   n1LG = InStr( 1, FaultString, "1LG" )

   ' Pre-fault
   Point = Int( PtNo*Cycles(1) )	' get integer part of prefault points
   While PtCount <= Point
      ' Va, Vb, Vc
      DataValue(1) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(1) )-bV )/aV )
      DataValue(2) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(2) )-bV )/aV )
      DataValue(3) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(3) )-bV )/aV )
      ' Ia, Ib, Ic, V0, I0, V2, I2
      DataValue(4) = Int( ( MagArray3(1) * Sin( Omeg*(PtCount-1)+AngArray3(1) )-bI )/aI )
      DataValue(5) = Int( ( MagArray3(1) * Sin( Omeg*(PtCount-1)+AngArray3(2) )-bI )/aI )
      DataValue(6) = Int( ( MagArray3(1) * Sin( Omeg*(PtCount-1)+AngArray3(3) )-bI )/aI )
      DataValue(7) = 0
      DataValue(8) = 0
      DataValue(9) = 0
      DataValue(10)= 0

      strText$ = Format( PtCount, "0" ) & "," & Format( 1e6*(PtCount-1)/Sample,"#######0" ) & "," & Format( DataValue(1), "####0" ) _
         & "," & Format( DataValue(2), "####0" ) & "," & Format( DataValue(3), "####0" ) & "," & Format( DataValue(4), "####0" ) _
         & "," & Format( DataValue(5), "####0" ) & "," & Format( DataValue(6), "####0" )
      If OutZ = 1 Then strText$ = strText$ & "," & Format( DataValue(7), "####0" ) & "," & Format( DataValue(8), "####0" )
      If OutN = 1 Then strText$ = strText$ & "," & Format( DataValue(9), "####0" ) & "," & Format( DataValue(10), "####0" )
      Print #FileNo, strText$
      PtCount = PtCount+1
   Wend
   
   ' Fault
   ' Get bus Voltage
   If GetSCVoltage( BusHnd, MagArray, AngArray, 4 ) = 0 Then 
      Print "Get Bus voltage failed."
      Exit Function
   End If
   For ii = 1 To 3
      MagArray2(ii) = Sqr(2)*MagArray(ii)*1000/PT_Ratio
      AngArray2(ii) = AngArray(ii)*pi/180
   Next ii
'   AngArray2(1) = AngArray(1)*pi/180
   ' Get branch current
   If GetSCCurrent( BranchHnd, MagArray, AngArray, 4 ) = 0 Then 
      Print "Get Branch current failed."
      Exit Function
   End If
   For ii = 1 To 3
      MagArray2(ii+3) = Sqr(2)*MagArray(ii)/CT_Ratio
      AngArray2(ii+3) = AngArray(ii)*pi/180
   Next ii
'   AngArray2(1) = AngArray2(1)-AngArray(1)*pi/180		' V-I
'   If AngArray2(1) <=-2*pi Then AngArray2(1) = AngArray2(1) + 2*pi
'   If AngArray2(1) >=2*pi  Then AngArray2(1) = AngArray2(1) - 2*pi

   ' Get bus Voltage in sequence
   If GetSCVoltage( BusHnd, MagArray, AngArray, 2 ) = 0 Then 
      Print "Get Bus sequence voltage failed."
      Exit Function
   End If
   MagArray2(7) = Sqr(2)*MagArray(1)*1000/PT_Ratio	' V0
   AngArray2(7) = AngArray(1)*pi/180			' V0
   MagArray2(9) = Sqr(2)*MagArray(3)*1000/PT_Ratio	' V2
   AngArray2(9) = AngArray(3)*pi/180			' V2
   ' Get branch current in sequence
   If GetSCCurrent( BranchHnd, MagArray, AngArray, 2 ) = 0 Then 
      Print "Get Branch sequence current failed."
      Exit Function
   End If
   MagArray2(8)  = Sqr(2)*MagArray(1)/CT_Ratio		' I0
   AngArray2(8)  = AngArray(1)*pi/180			' I0
   MagArray2(10) = Sqr(2)*MagArray(3)/CT_Ratio		' I2
   AngArray2(10) = AngArray(3)*pi/180			' I2
   
   Point = Int( PtNo*(Cycles(1)+Cycles(2)) )	' get integer part of fault points
   While PtCount <= Point
      ' Va, Vb, Vc
      DataValue(1) = Int( ( MagArray2(1) * Sin( Omeg*(PtCount-1)+AngArray1(1) )-bV )/aV )
      DataValue(2) = Int( ( MagArray2(2) * Sin( Omeg*(PtCount-1)+AngArray1(2) )-bV )/aV )
      DataValue(3) = Int( ( MagArray2(3) * Sin( Omeg*(PtCount-1)+AngArray1(3) )-bV )/aV )
      ' Ia, Ib, Ic
'      DataValue(4) = Int( ( MagArray2(4) * Sin( Omeg*(PtCount-1)+AngArray1(1)-AngArray2(1) )-bI )/aI )
'      DataValue(5) = Int( ( MagArray2(5) * Sin( Omeg*(PtCount-1)+AngArray1(2)-AngArray2(1) )-bI )/aI )
'      DataValue(6) = Int( ( MagArray2(6) * Sin( Omeg*(PtCount-1)+AngArray1(3)-AngArray2(1) )-bI )/aI )
      DataValue(4) = Int( ( MagArray2(4) * Sin( Omeg*(PtCount-1)+AngArray2(4) )-bI )/aI )
      DataValue(5) = Int( ( MagArray2(5) * Sin( Omeg*(PtCount-1)+AngArray2(5) )-bI )/aI )
      DataValue(6) = Int( ( MagArray2(6) * Sin( Omeg*(PtCount-1)+AngArray2(6) )-bI )/aI )


      ' V0, I0, V2, I2
      DataValue(7) = Int( ( MagArray2(7) * Sin( Omeg*(PtCount-1)+AngArray2(7) )-bV )/aV )
      DataValue(8) = Int( ( MagArray2(8) * Sin( Omeg*(PtCount-1)+AngArray2(8) )-bI )/aI )
      DataValue(9) = Int( ( MagArray2(9) * Sin( Omeg*(PtCount-1)+AngArray2(9) )-bV )/aV )
      DataValue(10)= Int( ( MagArray2(10)* Sin( Omeg*(PtCount-1)+AngArray2(10) )-bI )/aI )

      strText$ = Format( PtCount, "0" ) & "," & Format( 1e6*(PtCount-1)/Sample,"#######0" ) & "," & Format( DataValue(1), "####0" ) _
         & "," & Format( DataValue(2), "####0" ) & "," & Format( DataValue(3), "####0" ) & "," & Format( DataValue(4), "####0" ) _
         & "," & Format( DataValue(5), "####0" ) & "," & Format( DataValue(6), "####0" )
      If OutZ = 1 Then strText$ = strText$ & "," & Format( DataValue(7), "####0" ) & "," & Format( DataValue(8), "####0" )
      If OutN = 1 Then strText$ = strText$ & "," & Format( DataValue(9), "####0" ) & "," & Format( DataValue(10), "####0" )
      Print #FileNo, strText$
      PtCount = PtCount+1
   Wend

   ' Reclosing
   If Cycles(3) > 0.0 Then
     Point = Point + Int( PtNo*Cycles(3) )	' get integer part of reclosing points
     While PtCount <= Point
        ' Va, Vb, Vc
        DataValue(1) = 0.0
        If Cycles(5) > 0.0 And n1LG > 0 Then _
          DataValue(2) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(2) )-bV )/aV ) _
        Else DataValue(2) = 0.0
        If Cycles(5) > 0.0 And n1LG Then _
          DataValue(3) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(3) )-bV )/aV ) _
        Else DataValue(3) = 0.0
        ' Ia, Ib, Ic, V0, I0, V2, I2
        DataValue(4) = 0.0
        If Cycles(5) > 0.0 And n1LG > 0 Then _
          DataValue(5) = Int( ( MagArray3(1) * Sin( Omeg*(PtCount-1)+AngArray3(2) )-bI )/aI ) _
        Else DataValue(5) = 0.0
        If Cycles(5) > 0.0 And n1LG > 0 Then _
          DataValue(6) = Int( ( MagArray3(1) * Sin( Omeg*(PtCount-1)+AngArray3(3) )-bI )/aI ) _
        Else DataValue(6) = 0.0
        DataValue(7) = 0.0
        DataValue(8) = 0.0
        DataValue(9) = 0.0
        DataValue(10)= 0.0

        strText$ = Format( PtCount, "0" ) & "," & Format( 1e6*(PtCount-1)/Sample,"#######0" ) & "," & Format( DataValue(1), "####0" ) _
           & "," & Format( DataValue(2), "####0" ) & "," & Format( DataValue(3), "####0" ) & "," & Format( DataValue(4), "####0" ) _
           & "," & Format( DataValue(5), "####0" ) & "," & Format( DataValue(6), "####0" )
        If OutZ = 1 Then strText$ = strText$ & "," & Format( DataValue(7), "####0" ) & "," & Format( DataValue(8), "####0" )
        If OutN = 1 Then strText$ = strText$ & "," & Format( DataValue(9), "####0" ) & "," & Format( DataValue(10), "####0" )
        Print #FileNo, strText$
        PtCount = PtCount+1
     Wend
   End If

   ' Post-fault
   Point = Point + Int( PtNo*Cycles(4) )	' get integer part of post-fault points
   While PtCount <= Point
      ' Va, Vb, Vc
      DataValue(1) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(1) )-bV )/aV )
      DataValue(2) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(2) )-bV )/aV )
      DataValue(3) = Int( ( MagArray1(1) * Sin( Omeg*(PtCount-1)+AngArray1(3) )-bV )/aV )
      ' Ia, Ib, Ic, V0, I0, V2, I2
      DataValue(4) = Int( ( MagArray3(1) * Sin( Omeg*(PtCount-1)+AngArray3(1) )-bI )/aI )
      DataValue(5) = Int( ( MagArray3(1) * Sin( Omeg*(PtCount-1)+AngArray3(2) )-bI )/aI )
      DataValue(6) = Int( ( MagArray3(1) * Sin( Omeg*(PtCount-1)+AngArray3(3) )-bI )/aI )
      DataValue(7) = 0
      DataValue(8) = 0
      DataValue(9) = 0
      DataValue(10)= 0

      strText$ = Format( PtCount, "0" ) & "," & Format( 1e6*(PtCount-1)/Sample,"#######0" ) & "," & Format( DataValue(1), "####0" ) _
         & "," & Format( DataValue(2), "####0" ) & "," & Format( DataValue(3), "####0" ) & "," & Format( DataValue(4), "####0" ) _
         & "," & Format( DataValue(5), "####0" ) & "," & Format( DataValue(6), "####0" )
      If OutZ = 1 Then strText$ = strText$ & "," & Format( DataValue(7), "####0" ) & "," & Format( DataValue(8), "####0" )
      If OutN = 1 Then strText$ = strText$ & "," & Format( DataValue(9), "####0" ) & "," & Format( DataValue(10), "####0" )
      Print #FileNo, strText$
      PtCount = PtCount+1
   Wend

   Print #FileNo, Chr(26)		' End of file marker
   Comtrade = 1
End Function
'=================================End of Comtrade()====================================
