' ASPEN PowerScript Sample Program
'
' FLTREPORT.BAS
'
' Export fault simulation result to file for use in relay setting.
' Following fault quantities are being reported at relay group location:
' - Phase voltages and currents
' - Zero, negative and positive sequence voltage and currents
' Output file formats:
' - Comma delimited text file (.cvs)
'
' Version 1.0
' Category: OneLiner
'
'
Sub main()

   dim ShowFlagRly(5)
   
   ' Initialize 
   For ii = 1 To 4 
     ShowFlagRly(ii) = 1
   Next 
   
   VersionNum$ = "1.0"

   
   ' Make sure a relay group is being selected
   If GetEquipment( TC_PICKED, GroupHnd& ) = 0 Then
     Print "Please select a relay group must before running this script program"
     Exit Sub
   End If
   
   If EquipmentType( GroupHnd ) <> TC_RLYGROUP Then
     Print "Please select a relay group must before running this script program"
     Exit Sub
   End If
   
   If PickFault( 1 ) = 0 Then 
     Print "No fault simulation result available"
     Exit Sub
   End If

   Call GetData( GroupHnd, RG_nBranchHnd, BrHnd& )
   Call GetData( BrHnd, BR_nBus1Hnd, Bus1Hnd& )
   Call GetData( Bus1Hnd, BUS_sName, BName1$ )
   Call GetData( Bus1Hnd, BUS_dKVNominal, KV1# )
   Call GetData( BrHnd, BR_nBus2Hnd, Bus2Hnd& )
   Call GetData( Bus2Hnd, BUS_sName, BName2$ )
   Call GetData( Bus2Hnd, BUS_dKVNominal, KV2# )


   BrName$   = BName1 + Trim(Str(KV1)) + "-" + BName2 + Trim(Str(KV2))
   
   FName$    = "C:\000tmp\" + BrName

   OutCode& = DiaScope( FName$ )
   If OutCode = 0  Then Exit Sub ' Cancel

   Delim$ = ","
   
   ' Prepare output file
   If OutCode = 2 Then Open FName$ For Append As 1
   If OutCode = 1 Then Open FName$ For Output As 1

   
   ' Print column header
   If OutCode = 1 Then 
   ' Print file header
   
     DateInfo$ = Date()
     TimeInfo$ = time$(Now)
  
     Print #1, "Fault solution report"
     Print #1, "Version:,", VersionNum
     Print #1, "OneLiner Version:,", GetProgramVersion()
     Print #1, "Date/Time:,", DateInfo, " ", TimeInfo
     Print #1, "ORL file name:,", GetOlrFileName()
     Print #1, "Name of this file:,", FName
     Print #1," "
     Print #1, _
      Chr(34) & "Fault Info"    & Chr(34) & Delim$ & _
      Chr(34) & "Fault Conn"    & Chr(34) & Delim$ & _
      Chr(34) & "Va"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Ia"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Vb"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Ib"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Vc"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Ic"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Vo"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "Io"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "V1"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "I1"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "V2"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ & _
      Chr(34) & "I2"    & Chr(34) & Delim$ & _
      Chr(34) & " "       & Chr(34) & Delim$ 
   End If
   Print #1," "
   Print #1, "Relay group:,", BrName
   ' Loop over selected faults and export data
   Index = 1
   While PickFault(Index) > 0 
'   While ShowFault( Index, 9, 1, 0, ShowFlagRly ) > 0
     Call XportAFault( Index, Delim$, BrHnd )
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
Begin Dialog OUTPUTDIA 57,49, 227, 63, "Fault Simulation Report"
  OptionGroup .GROUP_1
    OptionButton 88,28,40,8, "&Overwrite"
    OptionButton 132,28,44,8, "&Append"
  TextBox 68,8,128,12, .EditBox_2
  OKButton 60,46,76,12
  CancelButton 148,46,36,12
  Text 4,8,60,12, "Output file name:"
  Text 200,8,21,12, ".CSV"
End Dialog



'======================================================================================
' DiaScope
'   Solicit user input on the scope of the export
'
'======================================================================================
Function DiaScope( ByRef FName$ ) As Long
Dim dlg As OUTPUTDIA

DiaScope      = 0
dlg.EditBox_2 = Fname$
RetCode&      = Dialog( dlg )
If RetCode = 0 Then Exit Function	' Canceled
DiaScope = dlg.GROUP_1 + 1 ' Append or overwrite
Fname$   = dlg.EditBox_2
Fname$ = Fname$ + ".csv"
End Function

'======================================================================================
' XportAFault
'   Print out fault result to file #1
'
'======================================================================================
Sub XportAFault( ByVal Index&, ByVal Delim$, ByVal BranchHnd& )
  Dim MagArray(16) As Double
  Dim AngArray(16) As Double

  ' Get fault info
  FltInfo$ = FaultDescription()
  If InStr(1, FltInfo, " 3LG ") > 0 Then FltInfo2$ = "3LG"
  If InStr(1, FltInfo, " 2LG ") > 0 Then FltInfo2$ = "2LG"
  If InStr(1, FltInfo, " 1LG ") > 0 Then FltInfo2$ = "1LG"
  If InStr(1, FltInfo, " LL ")  > 0 Then FltInfo2$ = "LL"
  
  ' Get Voltage
  rCode = GetSCVoltage( BranchHnd, MagArray, AngArray, 4 )
  VA1$ = Format( MagArray(1), "####0.0") & Delim$ & Format( AngArray(1), "#0.0")
  VB1$ = Format( MagArray(2), "####0.0") & Delim$ & Format( AngArray(2), "#0.0")
  VC1$ = Format( MagArray(3), "####0.0") & Delim$ & Format( AngArray(3), "#0.0")
  ' Get Voltage in Sequence
  rCode = GetSCVoltage( BranchHnd, MagArray, AngArray, 2 )
  V01$ = Format( MagArray(1), "####0.0") & Delim$ & Format( AngArray(1), "#0.0")
  V11$ = Format( MagArray(2), "####0.0") & Delim$ & Format( AngArray(2), "#0.0")
  V21$ = Format( MagArray(3), "####0.0") & Delim$ & Format( AngArray(3), "#0.0")
  ' Get Current
  rCode = GetSCCurrent( BranchHnd, MagArray, AngArray, 4 )
  IA1$= Format( MagArray(1), "####0.0") & Delim$ & Format( AngArray(1), "#0.0")
  IB1$= Format( MagArray(2), "####0.0") & Delim$ & Format( AngArray(2), "#0.0")
  IC1$= Format( MagArray(3), "####0.0") & Delim$ & Format( AngArray(3), "#0.0")
  ' Get Current in Sequence
  rCode = GetSCCurrent( BranchHnd, MagArray, AngArray, 2 )
  I01$= Format( MagArray(1), "####0.0") & Delim$ & Format( AngArray(1), "#0.0")
  I11$= Format( MagArray(2), "####0.0") & Delim$ & Format( AngArray(2), "#0.0")
  I21$= Format( MagArray(3), "####0.0") & Delim$ & Format( AngArray(3), "#0.0")
  Print #1, _
    Chr(34) & FltInfo$ & Chr(34) & Delim$ & _
    Chr(34) & FltInfo2$ & Chr(34) & Delim$ & _
    VA1 & Delim$ & IA1 & Delim$ & _
    VB1 & Delim$ & IB1 & Delim$ & _
    VC1 & Delim$ & IC1 & Delim$ & _
    V01 & Delim$ & I01 & Delim$ & _
    V11 & Delim$ & I11 & Delim$ & _
    V21 & Delim$ & I21 & Delim$

End Sub