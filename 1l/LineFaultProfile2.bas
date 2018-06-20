' ASPEN PowerScrip sample program
'
' LineFaultProfile2.BAS
'
' Run intemmediate fault simulation on transmission line
' Read list of lines from an input CSV file
' Report fault currents in CSV file.
'
' Version :1.0
' Category: OneLiner
'
'******************************************************************************************************
'TODO: Modify parameters in this section if needed
 Const DataPath$        = "c:\000tmp\"
 Const FileNameIn$      = "linelist.txt"     ' Input file name
 Const PercentStepSize  = 0.1	             ' intermediate fault percent step
'
'Global variables

Dim StepSize As Double

Sub main()
  dim nChecked As long
  dim dlg As LINEFLT
  dim sFileNameOut As String
  
  StepSize          = PercentStepSize
  sList$            = ""
  
  Print "Please select an input file"   
  InputFile$ = FileOpenDialog( "", "Excel File (*.csv)||", 0 )
  If Len(InputFile) = 0 Then Exit Sub
  
  Open InputFile For Input As 1
  sList$ = ""
  nLine = 0
  Line Input #1, aLine$ ' skip the first line
  Do While Not EOF(1)
    Line Input #1, aLine$ ' Read a line of data.
    If Len(sList) > 0 Then sList = sList & Chr(13) & Chr(10)
    sList = sList & aLine
    nLine = nLine + 1
  Loop
  Close #1
  
  Print "Please select or create an excel file for output report"
  OutputFile$ = FileSaveDialog( "", "Excel File (*.csv)|*.csv||", ".csv", 2+16 )
  If Len(OutputFile) = 0 Then exit Sub
  Open OutputFile For Output As 2

  Print #2, "Line Intermediate Fault Calculation Report"
  Print #2, "Date: ", Date()
  Print #2, "OneLiner file name: ", GetOLRFileName()
  Print #2, "Study date: N/A"
  Print #2, ""    

  Print #2, "Bus 1,kV,Bus 2,kV,CktID,Flt Type,Pct,IA_Real(Total),IA_Imag(Total),IB_Real(Total),IB_Imag(Total),IC_Real(Total),IC_Imag(Total)," & _
            "IA_Real(Bus 1),IA_Imag(Bus 1),IB_Real(Bus 1),IB_Imag(Bus 1),IC_Real(Bus 1),IC_Imag(Bus 1)," & _
            "IA_Real(Bus 2),IA_Imag(Bus 2),IB_Real(Bus 2),IB_Imag(Bus 2),IC_Real(Bus 2),IC_Imag(Bus 2)"
  nChecked = 0

  nLineIndex = 0
  sData$ = sList$
  Do While Len(sData$) > 0 
    Call parseALine( sData, Chr(13) & Chr(10), aLine$, sData )
    ' Bus1, kV1, Bus2, kV2, cktID
    If Len(aLine$) > 0 Then Call parseALine( aLine$, ",", sBus1$,  aLine$ )
    If Len(aLine$) > 0 Then Call parseALine( aLine$, ",", sKV1$,   aLine$ )
    If Len(aLine$) > 0 Then Call parseALine( aLine$, ",", sBus2$,  aLine$ )
    If Len(aLine$) > 0 Then Call parseALine( aLine$, ",", sKV2$,   aLine$ )
    If Len(aLine$) > 0 Then Call parseALine( aLine$, ",", sCktID$, aLine$ )
   
    sOut = sBus1 & "," & sKV1 & "," & sBus2 & "," & sKV2 & "," & sCktID
    Print #2, sOut 
        
    BranchHnd = branchSearch( sBus1$, Val(sKV1$), sBus2$, Val(sKV2$), sCktID$ )
    Call FindBus( sBus1$, Val(sKV1$), nBus1Hnd& )
    Call FindBus( sBus2$, Val(sKV2$), nBus2Hnd& )
    If BranchHnd > 0 Then
      ' Get branch type
      Call GetData( BranchHnd, BR_nType, TypeCode )
      If TypeCode = TC_LINE Then
        If SimulateLineFaults( BranchHnd, nBus1Hnd, nBus2Hnd ) > 0 Then nChecked = nChecked + 1 Else GoTo HasError
      End If
    End If
    nLineIndex = nLineIndex + 1
    nDone& = nLineIndex*100/nLine
    sMsg$ = "Record " & Str(nLineIndex) & " of " & Str(nLine)
    If 2 = ProgressDialog( 1, "Reading line data", sMsg$, nDone& ) Then 
      Print "Cancel button pressed"
      Close 2 
      exit Sub 
    End If
  Loop
  Close 2
  Call ProgressDialog( 0, "", "", 0 )
  exit Sub
HasError:
  Close
  Print "Error: ", ErrorString( )
End Sub


Function SimulateLineFaults( ByVal BranchHnd&, nBus1Hnd&, nBus2Hnd& ) As long
  Dim FltConnection(4) As Long
  Dim FltOption(14) As Double
  Dim OutageType(3) As Long
  Dim OutageList(15) As Long
  Dim FltConnStr(4) As String
  dim vdReal(12) As double
  dim vdImag(12) As double
  dim vdRealBranch(12) As double
  dim vdImagBranch(12) As double 
  Dim DummyArray(6) As Long   
  
  For ii = 1 To 14
  FltOption(ii) = 0.0
  Next
  For ii = 1 To 4
  FltConnection(ii) = 0
  Next
  For ii = 1 To 3
   OutageType(ii) = 0
  Next
  
  FltConnection(1) = 1      '3LG enabled
  FltConnection(3) = 1      '1LG enabled
  
  FltOption(13) = 0			'Intermediate percent from
  FltOption(14) = 0   		'Intermediate percent to

  dFltR     = 0
  dFltX     = 0
  
  nSwap = 0
  Call GetData( BranchHnd, BR_nHandle, nItemHnd& )
  Call GetData( nItemHnd, LN_nBus1Hnd, nHndBus1& )
  Call GetData( nItemHnd, LN_nBus2Hnd, nHndBus2& ) 
  If nHndBus1 = nBus2Hnd And nHndBus2 = nBus1Hnd Then
    nSwap = 1
  End If
  
  Percent# = StepSize
  Do While Percent <= 100
    sOut$ = ""
    'Simulate faults
    If Percent = StepSize Then
      FltOption(1)  = 1
      If 0 = DoFault( BranchHnd, FltConnection, FltOption, OutageType, OutageList, dFltR, dFltX, 1 ) Then
        SimulateLineFaults = 0
        exit Function
      End If
      PickFault( SF_FIRST )
      If ShowFault( SF_FIRST, 1, 4, 0, DummyArray ) = 0 Then GoTo HasError
      Do
        ' Total fault current
        If 0 = GetSCCurrent( HND_SC, vdReal, vdImag, 4 ) Then GoTo HasError      
        If 0 = GetSCCurrent( nItemHnd, vdRealBranch, vdImagBranch, 4 ) Then GoTo HasError
        
        sFltDesc$ = FaultDescription()
        If InStr( 1, sFltDesc, " 3LG " ) > 0 Then _
          sFltConn = "3LG" _
        Else If InStr( 1, sFltDesc, " LL " ) > 0 Then _
          sFltConn = "LL" _
        Else If InStr( 1, sFltDesc, " 1LG " ) > 0 Then _
          sFltConn = "1LG" _
        Else If InStr( 1, sFltDesc, " 2LG " ) > 0 Then _
          sFltConn = "2LG"
        If nSwap = 1 Then
          sOut$ = ",,,,," & sFltConn & "," & "Close-in" & "," _
                   & Format(vdReal(1),"#0.000") & "," & Format(vdImag(1),"#0.000") & "," _
                   & Format(vdReal(2),"#0.000") & "," & Format(vdImag(2),"#0.000") & "," _
                   & Format(vdReal(3),"#0.000") & "," & Format(vdImag(3),"#0.000") & "," _
                   & Format(vdRealBranch(5),"#0.000") & "," & Format(vdImagBranch(5),"#0.000") & "," _
                   & Format(vdRealBranch(6),"#0.000") & "," & Format(vdImagBranch(6),"#0.000") & "," _
                   & Format(vdRealBranch(7),"#0.000") & "," & Format(vdImagBranch(7),"#0.000") & "," _
                   & Format(vdRealBranch(1),"#0.000") & "," & Format(vdImagBranch(1),"#0.000") & "," _
                   & Format(vdRealBranch(2),"#0.000") & "," & Format(vdImagBranch(2),"#0.000") & "," _
                   & Format(vdRealBranch(3),"#0.000") & "," & Format(vdImagBranch(3),"#0.000") 
        Else  
          sOut$ = ",,,,," & sFltConn & "," & "Close-in" & "," _
                   & Format(vdReal(1),"#0.000") & "," & Format(vdImag(1),"#0.000") & "," _
                   & Format(vdReal(2),"#0.000") & "," & Format(vdImag(2),"#0.000") & "," _
                   & Format(vdReal(3),"#0.000") & "," & Format(vdImag(3),"#0.000") & "," _
                   & Format(vdRealBranch(1),"#0.000") & "," & Format(vdImagBranch(1),"#0.000") & "," _
                   & Format(vdRealBranch(2),"#0.000") & "," & Format(vdImagBranch(2),"#0.000") & "," _
                   & Format(vdRealBranch(3),"#0.000") & "," & Format(vdImagBranch(3),"#0.000") & "," _
                   & Format(vdRealBranch(5),"#0.000") & "," & Format(vdImagBranch(5),"#0.000") & "," _
                   & Format(vdRealBranch(6),"#0.000") & "," & Format(vdImagBranch(6),"#0.000") & "," _
                   & Format(vdRealBranch(7),"#0.000") & "," & Format(vdImagBranch(7),"#0.000")
        End If
        Print #2, sOut$          
      Loop While PickFault( SF_NEXT ) > 0
      FltOption(1)  = 0
            
      FltOption(7)  = 1
      If 0 = DoFault( BranchHnd, FltConnection, FltOption, OutageType, OutageList, dFltR, dFltX, 1 ) Then
        SimulateLineFaults = 0
        exit Function
      End If
      PickFault( SF_FIRST )
      If ShowFault( SF_FIRST, 1, 4, 0, DummyArray ) = 0 Then GoTo HasError
      Do
        ' Total fault current
        If 0 = GetSCCurrent( HND_SC, vdReal, vdImag, 4 ) Then GoTo HasError
        If 0 = GetSCCurrent( nItemHnd, vdRealBranch, vdImagBranch, 4 ) Then GoTo HasError

        sFltDesc$ = FaultDescription()
        If InStr( 1, sFltDesc, " 3LG " ) > 0 Then _
          sFltConn = "3LG" _
        Else If InStr( 1, sFltDesc, " LL " ) > 0 Then _
          sFltConn = "LL" _
        Else If InStr( 1, sFltDesc, " 1LG " ) > 0 Then _
          sFltConn = "1LG" _
        Else If InStr( 1, sFltDesc, " 2LG " ) > 0 Then _
          sFltConn = "2LG"
        If nSwap = 1 Then
          sOut$ = ",,,,," & sFltConn & "," & "Line end" & "," _
                   & Format(vdReal(1),"#0.000") & "," & Format(vdImag(1),"#0.000") & "," _
                   & Format(vdReal(2),"#0.000") & "," & Format(vdImag(2),"#0.000") & "," _
                   & Format(vdReal(3),"#0.000") & "," & Format(vdImag(3),"#0.000") & "," _
                   & Format(vdRealBranch(5),"#0.000") & "," & Format(vdImagBranch(5),"#0.000") & "," _
                   & Format(vdRealBranch(6),"#0.000") & "," & Format(vdImagBranch(6),"#0.000") & "," _
                   & Format(vdRealBranch(7),"#0.000") & "," & Format(vdImagBranch(7),"#0.000") & "," _
                   & Format(vdRealBranch(1),"#0.000") & "," & Format(vdImagBranch(1),"#0.000") & "," _
                   & Format(vdRealBranch(2),"#0.000") & "," & Format(vdImagBranch(2),"#0.000") & "," _
                   & Format(vdRealBranch(3),"#0.000") & "," & Format(vdImagBranch(3),"#0.000") 
        Else  
          sOut$ = ",,,,," & sFltConn & "," & "Line end" & "," _
                   & Format(vdReal(1),"#0.000") & "," & Format(vdImag(1),"#0.000") & "," _
                   & Format(vdReal(2),"#0.000") & "," & Format(vdImag(2),"#0.000") & "," _
                   & Format(vdReal(3),"#0.000") & "," & Format(vdImag(3),"#0.000") & "," _
                   & Format(vdRealBranch(1),"#0.000") & "," & Format(vdImagBranch(1),"#0.000") & "," _
                   & Format(vdRealBranch(2),"#0.000") & "," & Format(vdImagBranch(2),"#0.000") & "," _
                   & Format(vdRealBranch(3),"#0.000") & "," & Format(vdImagBranch(3),"#0.000") & "," _
                   & Format(vdRealBranch(5),"#0.000") & "," & Format(vdImagBranch(5),"#0.000") & "," _
                   & Format(vdRealBranch(6),"#0.000") & "," & Format(vdImagBranch(6),"#0.000") & "," _
                   & Format(vdRealBranch(7),"#0.000") & "," & Format(vdImagBranch(7),"#0.000")
        End If
        Print #2, sOut$          
      Loop While PickFault( SF_NEXT ) > 0
      FltOption(7)  = 0
    End If
    
    FltOption(9)  = Percent  	'Intermediate percent
    If 0 = DoFault( BranchHnd, FltConnection, FltOption, OutageType, OutageList, dFltR, dFltX, 1 ) Then
      SimulateLineFaults = 0
      exit Function
    End If
    
    PickFault( SF_FIRST )
    If ShowFault( SF_FIRST, 1, 4, 0, DummyArray ) = 0 Then GoTo HasError
    Do
      ' Total fault current
      If 0 = GetSCCurrent( HND_SC, vdReal, vdImag, 4 ) Then GoTo HasError
      If 0 = GetSCCurrent( nItemHnd, vdRealBranch, vdImagBranch, 4 ) Then GoTo HasError
      
      sFltDesc$ = FaultDescription()
      If InStr( 1, sFltDesc, " 3LG " ) > 0 Then _
        sFltConn = "3LG" _
      Else If InStr( 1, sFltDesc, " LL " ) > 0 Then _
        sFltConn = "LL" _
      Else If InStr( 1, sFltDesc, " 1LG " ) > 0 Then _
        sFltConn = "1LG" _
      Else If InStr( 1, sFltDesc, " 2LG " ) > 0 Then _
        sFltConn = "2LG"
      If nSwap = 1 Then
        sOut$ = ",,,,," & sFltConn & "," & Format(Percent, "#0.00") & "," _
                 & Format(vdReal(1),"#0.000") & "," & Format(vdImag(1),"#0.000") & "," _
                 & Format(vdReal(2),"#0.000") & "," & Format(vdImag(2),"#0.000") & "," _
                 & Format(vdReal(3),"#0.000") & "," & Format(vdImag(3),"#0.000") & "," _
                 & Format(vdRealBranch(5),"#0.000") & "," & Format(vdImagBranch(5),"#0.000") & "," _
                 & Format(vdRealBranch(6),"#0.000") & "," & Format(vdImagBranch(6),"#0.000") & "," _
                 & Format(vdRealBranch(7),"#0.000") & "," & Format(vdImagBranch(7),"#0.000") & "," _
                 & Format(vdRealBranch(1),"#0.000") & "," & Format(vdImagBranch(1),"#0.000") & "," _
                 & Format(vdRealBranch(2),"#0.000") & "," & Format(vdImagBranch(2),"#0.000") & "," _
                 & Format(vdRealBranch(3),"#0.000") & "," & Format(vdImagBranch(3),"#0.000") 
      Else  
        sOut$ = ",,,,," & sFltConn & "," & Format(Percent, "#0.00") & "," _
                 & Format(vdReal(1),"#0.000") & "," & Format(vdImag(1),"#0.000") & "," _
                 & Format(vdReal(2),"#0.000") & "," & Format(vdImag(2),"#0.000") & "," _
                 & Format(vdReal(3),"#0.000") & "," & Format(vdImag(3),"#0.000") & "," _
                 & Format(vdRealBranch(1),"#0.000") & "," & Format(vdImagBranch(1),"#0.000") & "," _
                 & Format(vdRealBranch(2),"#0.000") & "," & Format(vdImagBranch(2),"#0.000") & "," _
                 & Format(vdRealBranch(3),"#0.000") & "," & Format(vdImagBranch(3),"#0.000") & "," _
                 & Format(vdRealBranch(5),"#0.000") & "," & Format(vdImagBranch(5),"#0.000") & "," _
                 & Format(vdRealBranch(6),"#0.000") & "," & Format(vdImagBranch(6),"#0.000") & "," _
                 & Format(vdRealBranch(7),"#0.000") & "," & Format(vdImagBranch(7),"#0.000")
      End If
      Print #2, sOut$          
    Loop While PickFault( SF_NEXT ) > 0
    Percent = Percent + StepSize
  Loop 
  SimulateLineFaults = 1
  exit Function
HasError:
  Print "Error: ", ErrorString( )
  stop
End Function	'SimulateLineFaults

Sub parseALine( ByVal aLine$, ByVal Delim$, ByRef sLeft$,  ByRef sRight$ )
  nPos = InStr( 1, aLine$, Delim$ )
  If nPos = 0 Then
    sLeft = aLine$
    sRight = ""
  Else
    sLeft = Left(aLine$, nPos-1)
    sRight = Mid(aLine$, nPos+Len(Delim), 9999 )
  End If
  sLeft  = Trim(sLeft)
  sRight = Trim(sRight)
End Sub

Function  branchSearch( sBus1$, KV1#, sBus2$, KV2#, sCktID$ ) As long
  branchSearch = 0
  If 0 = FindBus( sBus1, KV1, nHndBus1& ) Then exit Function
  If 0 = FindBus( sBus2, KV2, nHndBus2& ) Then exit Function
  BranchHnd = 0
  While GetBusEquipment( nHndBus1, TC_BRANCH, BranchHnd ) > 0
    Call GetData( BranchHnd, BR_nBus2Hnd, nHndFarBus& )
    If nHndFarBus = nHndBus2 Then
      Call GetData( BranchHnd, BR_nType, nBrType& )
      select case nBrType
        case TC_LINE
          nCode = LN_sID
        case TC_XFMR
          nCode = XF_sID
        case TC_XFMR3
          nCode = X3_sID
        case TC_PS
          nCode = PS_sID
      End select
      Call GetData( BranchHnd, BR_nHandle, nItemHnd& )
      Call GetData( nItemHnd, nCode, sID$ )
      If sID = sCktID Then
        branchSearch = BranchHnd
        exit Function
      End If
    End If
  Wend
  branchSearchReturn:
End Function

Function FindBus( sName$, dKV#, ByRef BusHnd& ) As long
  BusHnd& = 0
  Do While GetEquipment( TC_BUS, BusHnd ) = 1
    Call GetData( BusHnd, BUS_sName, thisName$ )
    If thisName = sName Then
      Call GetData( BusHnd, BUS_dKVnominal, thisKV# )
      If Abs(dKV-thisKV) < 0.000001 Then
        FindBus = 1
        exit Function
      End If
    End If
  Loop
  FindBus = 0
End Function
