' ASPEN PowerScript Sample Program
'
' TESTXPRT1.BAS
'
' Export fault simulation data to a file for use in relay testing
' Following quantities are being reported for both ends of the
' selected line:
' - Phase voltages and currents
' - Zero and Negative sequence voltage and currents
'
' Possible Output file formats:
' - Formated text file (.txt)
' - Comma delimited text file (.cvs)
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
   If GetEquipment( TC_PICKED, BranchHnd& ) = 0 _
      Or EquipmentType( BranchHnd ) <> TC_LINE Then
     Print "Must select a line"
     Exit Sub
   End If

   If PickFault( 1 ) = 0 Then 
     Print "No fault simulation result available"
     Exit Sub
   End If


   FName$   = "c:\temp\testxprt"
   OutCode& = DiaScope( FltList, FltDesc, FName$ )
   If OutCode = 0  Then Exit Sub ' Cancel

   If Right$( FName, 3 ) = "csv" Then
     Delim$ = ","
   Else
     Delim$ = Chr( 9 )
   End If

   ' Prepare output file
   If OutCode = 1 Then Open FName$ For Append As 1
   If OutCode = 2 Then Open FName$ For Output As 1

   ' Print column header
   If OutCode = 2 Then Print #1, _
      Chr(34) & "Comment" & Chr(34) & Delim$ & _
      Chr(34) & "Bus1"    & Chr(34) & Delim$ & _
      Chr(34) & "Bus2"    & Chr(34) & Delim$ & _
      Chr(34) & "Va_1"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Ia_1"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Vb_1"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Ib_1"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Vc_1"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Ic_1"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Vz_1"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Iz_1"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Vn_1"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Va_2"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Ia_2"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Vb_2"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Ib_2"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Vc_2"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Ic_2"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Vz_2"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Iz_2"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Vn_2"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$
   ' Loop over selected faults and export data
   Index = 1
   While FltList(Index) > -1
     If PickFault( FltList(Index) ) = 0 Then GoTo HasError
     Call XportAFault( FltDesc(Index), Delim$, BranchHnd )
     Index = Index + 1
   Wend
   Close ' Save the file

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
Begin Dialog OUTPUTDIA 57,49, 300, 208, "Export Scope"
  OptionGroup .GROUP_1
    OptionButton 200,156,40,8, "&Append"
    OptionButton 240,156,44,8, "&Overwrite"
  OptionGroup .GROUP_2
    OptionButton 68,172,64,8, "Tab delimited"
    OptionButton 132,172,36,8, "CSV"
  TextBox 68,156,128,12, .EditBox_2
  TextBox 3,16,292,112, .EditBox_1, WSTYLE
  OKButton 104,190,48,12
  CancelButton 160,190,48,12
  Text 4,4,276,8, "Following fault results are available for export:"
  Text 4,156,60,12, "Output file name:"
  Text 4,132,296,8, "Edit the list above to keep only faults that you want to export. You may customize"
  Text 4,140,268,8, "fault description strings while keeping the fault index numbers unchanged"
  Text 28,170,36,12, "File type:"
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
If dlg.Group_2 = 0 Then
  Fname$ = Fname$ + ".out"
Else
  Fname$ = Fname$ + ".csv"
End If
End Function

'======================================================================================
' XportAFault
'   Print out fault result to file #1
'
'======================================================================================
Sub XportAFault( ByVal Comment$, ByVal Delim$, ByVal BranchHnd& )
  Dim MagArray(16) As Double
  Dim AngArray(16) As Double

  ' Get end bus name
  rCode   = GetData( BranchHnd, LN_nBus1Hnd, BusHnd& )
  BName1$ = FullBusName( BusHnd )
  rCode   = GetData( BranchHnd, LN_nBus2Hnd, BusHnd& )
  BName2$ = FullBusName( BusHnd )

  ' Get Voltage
  rCode = GetSCVoltage( BranchHnd, MagArray, AngArray, 4 )
  VA1$ = Format( MagArray(1), "####0.0") & Delim$ & Format( AngArray(1), "#0.0")
  VB1$ = Format( MagArray(2), "####0.0") & Delim$ & Format( AngArray(2), "#0.0")
  VC1$ = Format( MagArray(3), "####0.0") & Delim$ & Format( AngArray(3), "#0.0")
  VA2$ = Format( MagArray(4), "####0.0") & Delim$ & Format( AngArray(4), "#0.0")
  VB2$ = Format( MagArray(5), "####0.0") & Delim$ & Format( AngArray(5), "#0.0")
  VC2$ = Format( MagArray(6), "####0.0") & Delim$ & Format( AngArray(6), "#0.0")
  ' Get Voltage in Sequence
  rCode = GetSCVoltage( BranchHnd, MagArray, AngArray, 2 )
  V01$ = Format( MagArray(1), "####0.0") & Delim$ & Format( AngArray(1), "#0.0")
  V21$ = Format( MagArray(3), "####0.0") & Delim$ & Format( AngArray(3), "#0.0")
  V02$ = Format( MagArray(4), "####0.0") & Delim$ & Format( AngArray(4), "#0.0")
  V22$ = Format( MagArray(6), "####0.0") & Delim$ & Format( AngArray(6), "#0.0")
  ' Get Current
  rCode = GetSCCurrent( BranchHnd, MagArray, AngArray, 4 )
  IA1$= Format( MagArray(1), "####0.0") & Delim$ & Format( AngArray(1), "#0.0")
  IB1$= Format( MagArray(2), "####0.0") & Delim$ & Format( AngArray(2), "#0.0")
  IC1$= Format( MagArray(3), "####0.0") & Delim$ & Format( AngArray(3), "#0.0")
  IA2$= Format( MagArray(4), "####0.0") & Delim$ & Format( AngArray(4), "#0.0")
  IB2$= Format( MagArray(5), "####0.0") & Delim$ & Format( AngArray(5), "#0.0")
  IC2$= Format( MagArray(6), "####0.0") & Delim$ & Format( AngArray(6), "#0.0")
  ' Get Current in Sequence
  rCode = GetSCCurrent( BranchHnd, MagArray, AngArray, 2 )
  I01$= Format( MagArray(1), "####0.0") & Delim$ & Format( AngArray(1), "#0.0")
  I02$= Format( MagArray(5), "####0.0") & Delim$ & Format( AngArray(5), "#0.0")
  Print #1, _
    Chr(34) & Comment$ & Chr(34) & Delim$ & _
    Chr(34) & BName1$ & Chr(34) & Delim$ & _
    Chr(34) & BName2$ & Chr(34) & Delim$ & _
    VA1 & Delim$ & IA1 & Delim$ & _
    VB1 & Delim$ & IB1 & Delim$ & _
    VC1 & Delim$ & IC1 & Delim$ & _
    V01 & Delim$ & I01 & Delim$ & V21 & Delim$ & _
    VA2 & Delim$ & IA2 & Delim$ & _
    VB2 & Delim$ & IB2 & Delim$ & _
    VC2 & Delim$ & IC2 & Delim$ & _
    V02 & Delim$ & I02 & Delim$ & V22

End Sub