' ASPEN PowerScrip sample program
'
' FiveFlt.BAS
'
' Simulate 5 faults on 3 transmission lines and save results to a txt file
'
'  Line A (Bus A - Bus B), Line B(Bus B - Bus C), Line C(Bus C - Bus D) 
'  Line A, B and C are stored in rows 2-4 of the input csv file in the format:
'  Bus 1 Name,Bus 1 KV,Bus 2 Name,Bus 2 KV,Ckt
'  Example:
'  GLEN LYN,132,CLAYTOR,132,1
'  CLAYTOR,132,FIELDALE,132,1
'  FIELDALE,132,OHIO,132,1
'
'  Location: 90% Line A, 10% LineC, 10%|50%|90% Line B
'  Fault Type: AG, BC   
'
' Usage instructions:
' 1) Launch OneLiner program and open *.OLR file
' 2) Tools->Script Manager...->Run script on disk->Select FiveFlt.BAS and click open   (run the power script code)
' 3) Select *.csv and click open (Power script code needs to read the input file to locate Line A, B, C)
' 4) Select *.txt and click Save (Power Script code needs an output file to save the result)
' 5) Open *.txt and check the result   
'
' Version 1.0
' Category: OneLiner
'

Dim xlApp As Object     ' Declare variable to hold the reference.
Dim wkbook As Object    ' Declare variable to hold the reference.
Dim dataSheet As Object
Sub main
  ' Variable declaration
  Dim FltConn(4) As Long
  Dim FltOption(14) As Double
  Dim OutageType(3) As Long
  Dim OutageList(20) As Long
  Dim MagArray(3) As Double
  Dim AngArray(3) As Double
  
  ExcelFile$ = FileOpenDialog( "Input File", "Excel File (*.csv)||", 0 )
  
  If Len(ExcelFile) = 0 Then 
    Print "Bye"
    Stop
  End If
  
  OutputFile$ = FileSaveDialog( "Output Report File", "Text files (*.txt)", ".txt", 2+16 )
  If OutputFile = "" Then Exit Sub
  Open OutputFile For Output As 1

  ' Get Pointer to Excel application
  On Error GoTo excelErr  
  Set xlApp = CreateObject("excel.application")
  Set wkbook = xlApp.Workbooks.Open( ExcelFile, True, True) 
  On Error GoTo dataSheetErr
  Set dataSheet = xlApp.Worksheets(1)
  On Error GoTo dataSheetErr
  
  For ii = 2 to 4
    sBus$ = dataSheet.Cells(ii,1).Value
    sKV$  = dataSheet.Cells(ii,2).Value
    Call findBusByName( sBus$, sKV$, nBus1Hnd& )
    sBus$ = dataSheet.Cells(ii,3).Value
    sKV$  = dataSheet.Cells(ii,4).Value
    Call findBusByName( sBus$, sKV$, nBus2Hnd& )
    sID$  = dataSheet.Cells(ii,5).Value
    branchHnd& = branchSearch( TC_LINE, nBus1Hnd&, nBus2Hnd&, 0, sID$ )
    If branchHnd = 0 Then GoTo dataSheetErr
    If ii = 2 Then 
      branchAHnd& = branchHnd
    ElseIf ii = 3 Then 
      branchBHnd& = branchHnd
    Else 
      branchCHnd& = branchHnd
    End If
  Next

  wkbook.Close
  ' Free memory  
  Set dataSheet = Nothing
  Set wkbook    = Nothing
  Set xlApp     = Nothing
  
  
  ' Initialize DoFault options using dialog data
  For ii = 1 To 4 
    FltConn(ii) = 0
  Next 
  For ii = 1 To 14
    FltOption(ii) = 0.0
  Next
  For ii = 1 To 3
    OutageType(ii) = 0
  Next
   
   ' Fault connection
  FltConn(3)    = 1	' 1LG_A 
  FltConn(4)    = 1 ' LL_BC
 
  For ii = 1 To 5 
    If ii = 1 Then 
      FltOption(9) = 90   ' Intermediate fault
      branchHnd = branchAHnd 
      nClear = 1 
    ElseIf ii = 2 Then 
      FltOption(9) = 10   ' Intermediate fault
      branchHnd = branchCHnd
      nClear = 0
    Else
      If ii = 3 Then FltOption(9) = 50
      If ii = 4 Then FltOption(9) = 10
      If ii = 5 Then FltOption(9) = 90
      branchHnd = branchBHnd
      nClear = 0
    End If
    
    ' Simulate fault
    If 0 = DoFault( branchHnd, FltConn, FltOption, OutageType, OutageList, 0.0, 0.0, nClear ) Then GoTo HasError    
  Next
  
  FaultFlag = 1 
  While PickFault( FaultFlag ) <> 0
    Print #1, FaultDescription(0)
    If GetSCCurrent( HND_SC, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
    sOutput = "Ia = " & _
            Format( MagArray(1), "#0.0") & "@" & Format( AngArray(1), "#0.0") & "," & _
            "Ib = " & _
            Format( MagArray(2), "#0.0") & "@" & Format( AngArray(2), "#0.0") & "," & _
            "Ic = " & _
            Format( MagArray(3), "#0.0") & "@" & Format( AngArray(3), "#0.0")
    Print #1, sOutput
    FaultFlag = SF_NEXT
  Wend

  GoTo Done
  
excelErr:	
dataSheetErr:
  Close 1
  ' Free memory  
  Set dataSheet = Nothing
  Set wkbook    = Nothing
  Set xlApp     = Nothing
  If Err.Number > 0 Then Print "Excecution error: " & Err.Description
  Stop 		
	
HasError:
  Close 1
  PrintTTY("ERROR - HasError")
  Stop
	
Done:
  Close 1
  PrintTTY(" " & Chr(10) & "Done!")		
  Exit Sub
End Sub  ' End of Sub Main()

 ' Find branch handle 
Function branchSearch( nType&, bus1Hnd&, bus2Hnd&, bus3Hnd&, CktID$ )
  branchSearch = 0
  branchHnd&   = 0
  select case nType
    case TC_LINE
      thisTypeID = LN_sID
    case TC_XFMR
      thisTypeID = XR_sID
    case TC_XFMR3
      thisTypeID = X3_sID
    case TC_PS
      thisTypeID = PS_sID
    case TC_SWITCH
      thisTypeID = SW_sID  
    case default
      Print "Error in branchSearch()"
      Stop
  End select
  While GetBusEquipment( bus1Hnd, TC_BRANCH, branchHnd ) > 0
    Call GetData(branchHnd, BR_nHandle, thisItemHnd&)
    If EquipmentType(thisItemHnd) = nType Then
      Call GetData(branchHnd, BR_nBus1Hnd, farBusHnd&)
      If farBusHnd = bus1Hnd Then Call GetData(branchHnd, BR_nBus2Hnd, farBusHnd&)
      If farBusHnd = bus2Hnd Then
        Call GetData(thisItemHnd, thisTypeID, myID$)
        myID = Trim(myID)
        CktIDTmp$ = "0" + CktID$
        If myID = CktID Or myID = CktIDTmp Or (Len(myID) = 0 And Len(CktID) = 0) Then
          branchSearch = branchHnd
          exit Do
        End If
      End If
    End If
  Wend
End Function
