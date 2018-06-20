' ASPEN PowerScript sample program
'
' UpdateBreakerName.BAS
'
' Update breaker name in OLR file 
'
' Version: 1.0
' Category: OneLiner
'

Sub main
  
  ExcelFile$ = FileOpenDialog( "", "Excel File (*.csv)||", 0 )
  
  If Len(ExcelFile) = 0 Then 
    Print "Bye"
    Stop
  End If

  printTTY("")
  printTTY("====================================================================================================================================")
  printTTY(ExcelFile)


  ' Get Pointer to Excel application
  On Error GoTo excelErr  
  Set xlApp = CreateObject("excel.application")
  Set wkbook = xlApp.Workbooks.Open( ExcelFile, True, True) 
  On Error GoTo dataSheetErr
  Set dataSheet = xlApp.Worksheets(1)

  On Error GoTo endProgram 

  ' Process the spreadsheet row by row
  rowCount& = 0
  Do
    aVal$ = dataSheet.Cells(2+rowCount,1).Value    
    If "" = aVal$ Then exit Do
    rowCount = rowCount + 1
  Loop While true
  
  If rowCount = 0 Then
    Print "Table has no data row"
    GoTo endProgram
  End If 
 
  For ii& = 1 to rowCount
    sBusName$    = dataSheet.Cells(1+ii,1).Value
    sBrkName$    = dataSheet.Cells(1+ii,2).Value
    sBrkNameNew$ = dataSheet.Cells(1+ii,3).Value
    busHnd& = findBusHnd(sBusName$)
    If busHnd = -1 Then 
      printTTY("  Error: Bus not found")
      GoTo endProgram   
    End If
    BreakerHnd& = breakerSearch( busHnd&, sBrkName$ )
    If BreakerHnd = 0 Then
      printTTY("  Error: Breaker not found")
      GoTo endProgram
    End If
    If SetData(BreakerHnd, BK_sID, sBrkNameNew$ ) = 0 Then
      PrintTTY("  Error: " + ErrorString())
      GoTo endProgram
    End If  
    If PostData(BreakerHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString())
      GoTo endProgram
    End If
    sMsg$ = "Record " & Str(ii) & " of " & Str(rowCount)
    If 2 = ProgressDialog( 1, "Reading XLS data", sMsg$, nDone& ) Then exit For
  Next
  Print rowCount & " breakers were updated successfully."

endProgram:
  Call ProgressDialog( 0, "", "", 0 )
  wkbook.Close
excelErr:
dataSheetErr:
  ' Free memory  
  Set dataSheet = Nothing
  Set wkbook    = Nothing
  Set xlApp     = Nothing
  If Err.Number > 0 Then Print "Excecution error: " & Err.Description
  Stop    
End Sub

 ' Find bus handle with bus name
Function findBusHnd( ByVal busStr$ ) As long
  findBusHnd = -1
  nLen& = Len(busStr)
  If UCase(Right(busStr,2)) = "KV" Then 
    busStr = Trim(Left(busStr,nLen-2))
    nLen   = Len(busStr)
  End If
  For ii = 1 to nLen - 1
    If Mid(busStr, nLen-ii, 1) = " " Then
      nPos = nLen - ii
      sToken$ = Trim(Mid(busStr, nPos+1, ii))
      aKV#    = Val(sToken)
      aName$  = Trim(Left(busStr, nPos-1))
      If findBusByName( aName, aKV, nHnd& ) Then findBusHnd = nHnd
      exit For
    End If
  Next
End Function

 ' Find breaker handle
Function breakerSearch( busHnd&, breakerName$ )
  breakerSearch = 0
  breakerHnd&   = 0
  While GetBusEquipment( busHnd, TC_BREAKER, breakerHnd& ) > 0
    Call GetData(breakerHnd, Bk_sID, myID$)
    myID = Trim(myID)
    If myID = breakerName Then
      breakerSearch = breakerHnd
    End If
  Wend
End Function

