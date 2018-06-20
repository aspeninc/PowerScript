' ASPEN PowerScript sample program
'
' UpdateBusNameNo.BAS
'
' Update bus name and bus number in OLR file 
' The input *.csv file includes 5 columns
' The first 3 columns are the bus data exported from OneLiner: Bus Name (Original), Bus No. (Original), Bus kV
' The last 2 columns are the updated bus data: Bus Name (Updated), Bus No. (Updated)
' The script uses Bus Name (Original) and Bus kV to locate the bus, then apply Bus Name (Updated), Bus No. (Updated)
' to the bus. 
' Bus No. is not unique for each bus, so Bus No. (Original) is for display only.
'
' Version 2.1
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
    sBusNameOrg$ = dataSheet.Cells(1+ii,1).Value
    nBusNoOrg    = Val(dataSheet.Cells(1+ii,2).Value)
    dBusKv       = Val(dataSheet.Cells(1+ii,3).Value)
    sBusNameNew$ = dataSheet.Cells(1+ii,4).Value
    nBusNoNew    = Val(dataSheet.Cells(1+ii,5).Value)
    If findBusByName( sBusNameOrg$, dBusKv, busHnd& ) = 0 Then 
      printTTY("  Error: Object not found")
      GoTo endProgram
    End If
    If SetData( busHnd, BUS_sName, sBusNameNew$ ) = 0 Then
      printTTY("  Error: " + ErrorString())
      GoTo endProgram
    End If
    If SetData( busHnd, BUS_nNumber, nBusNoNew$ ) = 0 Then
      printTTY("  Error: " + ErrorString())
      GoTo endProgram
    End If  
    If PostData(busHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString())
      GoTo endProgram
    End If
    sMsg$ = "Record " & Str(ii) & " of " & Str(rowCount)
    If 2 = ProgressDialog( 1, "Reading XLS data", sMsg$, nDone& ) Then exit For
  Next
  Print rowCount & " buses were updated successfully."

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
