' ASPEN PowerScript Sample Program
'
' TESTXPRT.BAS
'
' Export fault simulation data in a file to use in relay testing
'
' Version 1.0
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
   Dim FltList(200) As Long
   Dim FltDesc(200) As String

   ' Make sure a line with relay group is being selected
   If GetEquipment( TC_PICKED, DevHnd& ) = 0 Then
     Print "Must select a relay group"
     Exit Sub
   End If
   
   If EquipmentType( DevHnd ) <> TC_RLYGROUP Then
     Print "Must select a relay group"
     Exit Sub
   End If

   ' Get Branch and bus handle
   If GetData( DevHnd, RG_nBranchHnd, BranchHnd& ) = 0 Then GoTo hasError
   If GetData( BranchHnd, BR_nBus1Hnd, BusHnd& ) = 0 Then GoTo hasError

   If PickFault( 1 ) = 0 Then 
     Print "No fault simulation result available"
     Exit Sub
   End If


   FName$   = "c:\temp\testxprt.out"
   OutCode& = DiaScope( FltList, FltDesc, FName$ )
   If OutCode = 0  Then Exit Sub
   ' Prepare output file
   If OutCode = 1 Then Open FName$ For Append As 1
   If OutCode = 2 Then Open FName$ For Output As 1
   
   Index = 1
   While FltList(Index) > -1
     If PickFault( FltList(Index) ) = 0 Then GoTo HasError
     Call XportAFault( FltDesc(Index), BusHnd, BranchHnd )
     Index = Index + 1
   Wend
   Close
   Print Index-1; "Faults have been exported succesfully to: "; FName$
   Exit Sub
HasError:
   Print "Error: ", ErrorString( )
   Close
End Sub


'=============Dialog Spec============================================================
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
'=============Dialog Spec============================================================
Begin Dialog OUTPUTDIA 57,49, 300, 188, "Export Scope"
  TextBox 68,156,128,12, .EditBox_2
  TextBox 4,16,292,112, .EditBox_1, WSTYLE
  OKButton 192,172,48,12
  CancelButton 248,172,48,12
  OptionGroup .GROUP_1
    OptionButton 200,156,40,8, "&Append"
    OptionButton 244,156,44,8, "&Overwrite"
  Text 4,132,296,8, "Note: You may edit the list to keep only faults that you want to export and to customize"
  Text 4,4,276,8, "Following fault results are available for export:"
  Text 4,156,60,12, "Output file name:"
  Text 24,140,268,8, "fault description string. The original fault index number must be left unchanged"
End Dialog

'======================================================================================
' DiaScope
'   Solicit user input on the scope of the export
'
'======================================================================================
Function DiaScope( ByRef FltList() As Long, ByRef FltDesc() As String, ByRef FName$ ) As Long
Dim dlg As OUTPUTDIA

DiaScope      = 0
' Prepare fault list 
AString$ = ""
If PickFault( 1 ) = 0 Then Exit Sub
Do
  FltString$ = FaultDescription()
  ' Need to insert chr(13) at the end of each line to make it
  ' show up properly in the edit box
  CharPos = InStr( 1, FltString, Chr(10) )
  While CharPos > 0
    TempStr$   = Left$( FltString, CharPos - 1 )
    If Right( TempStr, 3 ) <> "on:" Then TempStr = TempStr + Chr(13) + Chr(10)
    TempStr    = TempStr + " " + LTrim(Mid$(FltString, CharPos+1, 9999 ))
    FltString$ = TempStr$
    CharPos    = InStr( CharPos+2, FltString, Chr(10) )
  Wend
  AString$ = AString$ + TempStr$ + Chr(13) + Chr(10) 
Loop While PickFault( SF_NEXT ) > 0
' Initialize dialog box
dlg.EditBox_1 = AString$
dlg.EditBox_2 = Fname$
RetCode&      = Dialog( dlg )
If RetCode = 0 Then Exit Function	' Canceled
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
If Count = 1 Then Exit Function   ' Nothing to report. Cancel
DiaScope = dlg.GROUP_1 + 1 ' Append or overwrite
Fname$   = dlg.EditBox_2
End Function

'======================================================================================
' XportAFault
'   Print out fault result to file #1
'
'======================================================================================
Sub XportAFault( ByVal Comment$, ByVal BusHnd&, ByVal BranchHnd& )
  Dim MagArray(16) As Double
  Dim AngArray(16) As Double

  ' Get Voltage
  rCode = GetSCVoltage( BusHnd, MagArray, AngArray, 4 )
  VA1$ = Format( MagArray(1), "####0.0") & Chr(9) & Format( AngArray(1), "#0.0")
  VB1$ = Format( MagArray(2), "####0.0") & Chr(9) & Format( AngArray(2), "#0.0")
  VC1$ = Format( MagArray(3), "####0.0") & Chr(9) & Format( AngArray(3), "#0.0")
  ' Get Voltage in Sequence
  rCode = GetSCVoltage( BusHnd, MagArray, AngArray, 2 )
  V01$ = Format( MagArray(1), "####0.0") & Chr(9) & Format( AngArray(1), "#0.0")
  V21$ = Format( MagArray(3), "####0.0") & Chr(9) & Format( AngArray(3), "#0.0")
  ' Get Current
  rCode = GetSCCurrent( BranchHnd, MagArray, AngArray, 4 )
  IA1$= Format( MagArray(1), "####0.0") & Chr(9) & Format( AngArray(1), "#0.0")
  IB1$= Format( MagArray(2), "####0.0") & Chr(9) & Format( AngArray(2), "#0.0")
  IC1$= Format( MagArray(3), "####0.0") & Chr(9) & Format( AngArray(3), "#0.0")
  ' Get Current in Sequence
  rCode = GetSCCurrent( BranchHnd, MagArray, AngArray, 2 )
  I01$= Format( MagArray(1), "####0.0") & Chr(9) & Format( AngArray(1), "#0.0")
  Print #1, Comment$ & Chr(9) & _
    VA1 & Chr(9) & IA1 & Chr(9) & _
    VB1 & Chr(9) & IB1 & Chr(9) & _
    VC1 & Chr(9) & IC1 & Chr(9) & _
    V01 & Chr(9) & I01 & Chr(9) & V21

End Sub
