' ASPEN PowerScript sample program
'
' ARCFLASHEX.BAS
'
' Run DoArcFlash with input data from a template file
' in Excel format with columns (fields) in following order:
'  1.  "No."                - Bus number  (optional)
'  2.  "Bus Name"	       - bus name
'  3.  "kV	               - bus kv 
'  4.  "Equipment Category" - 0: switch	1: cable	2: open air
'  5.  "Grounded"           - 1=yes, 2=no
'  6.  "Enclosed"           - 1=yes, 2=no
'  7.  "Conductor Gap (mm)"
'  8.  "Working Distance (inches)"
'  9.  "Breaker inerrupting Time (cycles)"
'  10. "Ignore 2Sec"        - 1=yes, 2=no
'  11. "Fault Clearing"     - 1: Auto	2: Manual Clearing Time	3: Step Event Analysis
'  12. "Fixed Delay"        - For clearing time option 2  (manual)
'  13. "Tier Num"           - For clearing time option 1 (auto)
'
' Version: 1.0
' Category: OneLiner
'
' PowerScript functions called:
' DoArcFlash()
'


Const IDOK = -1
Sub main()
  Begin Dialog Dialog_1 68,78,221,60, "Dialog"
    Text 10,6,201,27,"This script will read an input file with bus list plus calculation settings and generate arc-flash hazard report for all buses in the list."
    OKButton 123,39,40,13
    CancelButton 172,39,40,13
  End Dialog
  
  Begin Dialog Dialog_3 68,78,221,46, "Dialog"
    Text 10,6,201,17,"Select file to save calulcation results."
    OKButton 123,29,40,13
    CancelButton 172,29,40,13
  End Dialog  
  

  ' Variable declarations
  Dim fieldNames(50) As String 
  Dim fieldValue(50) As String
  Dim vdOption(10) As Double, vdResult(15) As Double
  Dim dlg1 As Dialog_1
  Dim dlg2 As Dialog_2
  Dim dlg3 As Dialog_3

  ' Read input file
  If IDOK <> Dialog( dlg1 ) Then Exit Sub
  
  Print "Select input file"
  ExcelFile$ = FileOpenDialog( "", "Excel File (*.csv, *.xls, *.xlsx)|*.csv;*.xls;*.xlsx||", 4 )
  
  If Len(ExcelFile) = 0 Then 
    Print "Bye"
    Stop
  End If

  ' Get Pointer to Excel application
  On Error GoTo excelErr  
  Set xlApp = CreateObject("excel.application")
  Set wkbook = xlApp.Workbooks.Open( ExcelFile, True, True) 
  On Error GoTo dataSheetErr
  Set dataSheet = xlApp.Worksheets(1)

  On Error GoTo endProgram 

  ' Read file header row
  rowCount& = 1
  Do
    aHeader1$ = dataSheet.Cells(rowCount,1).Value
    aHeader2$ = dataSheet.Cells(rowCount,2).Value
    aHeader3$ = dataSheet.Cells(rowCount,3).Value
    If "No." = aHeader1$ And "Bus Name" = aHeader2$ And "kV" = aHeader3$ Then exit Do
    rowCount = rowCount + 1
    If rowCount > 10 Then
       Print "Couldn't find import data."
       GoTo endProgram
    End If    
  Loop While true
  
  HeaderRow = rowCount
  
  ' Count rows in spreadsheet and read field names
  colCount = readXLSRow( dataSheet, HeaderRow, fieldNames, 49 ) 
  For ii = 1 to colCount 
    fieldNames(ii) = Trim(fieldNames(ii))
  Next

  ' Count rows in spreadsheet
  rowCount& = 0
  Do
    aVal$ = dataSheet.Cells(HeaderRow+1+rowCount,1).Value & dataSheet.Cells(HeaderRow+1+rowCount,2).Value & dataSheet.Cells(HeaderRow+1+rowCount,3).Value
    If "" = aVal$ Then exit Do
    rowCount = rowCount + 1
  Loop While true

  If rowCount = 0 Then
    Print "Table has no data row"
    GoTo endProgram
  End If 

  Print "Select file to save calculation results"
  OutputFile$ = FileSaveDialog( "", "Excel File (*.csv)|*.csv||", ".csv", 2+16 )
  If Len(OutputFile) = 0 Then 
    Print "Bye"
    Stop
  End If

  Open OutputFile For Output As 1
  Print #1, "Arc-flash Hazard Calculation Report"
  Print #1, "Date: ", Date()
  Print #1, "OneLiner file name: ", GetOLRFileName()
'  Print #1, "Study date: N/A"
  Print #1, ""
  Print #1, "BUS" & "," & "EQUI.CAT." & "," & "GROUNDED" & "," & "ENCLOSED" & "," & _
            "BKRTIME" & "," & "WORKDIST." & "," & "COND.GAP" & "," & "I3P" & "," & "IARC" & "," & _
            "CLRDEV" & "," & "CLRT" & "," & "IE" & "," & "CLRDEV85%" & "," & "CLRT85%" & "," & _
            "IE85%" & "," & "REQPPE" & "," & "BDRY_PPE1" & "," & "BDRY_PPE2" & "," & "BDRY_PPE3" & "," & _
            "BDRY_PPE4" & "," & "BDRY_PP4EX"
  ' Process the spreadsheet row by row
  nSuccess = 0
  For ii& = 1 to rowCount 
    Call readXLSRow( dataSheet, HeaderRow+ii, fieldValue, colCount )
    sBusName$   = Trim(fieldValue(2))
    dBusKv      = Val(fieldValue(3))
    nBusNumber& = Val(fieldValue(1))
    vdOption(1) = Int(Val(fieldValue(4)))    '0-Switchgear; 1-Cable; 2- open air
    vdOption(2) = Int(Val(fieldValue(5)))    '0-Ungrounded;1-Grounded
    vdOption(3) = Int(Val(fieldValue(6)))    '0-No enclosure;1-Enclosed
    vdOption(4) = Val(fieldValue(7))	     'Conductor gap in mm
    vdOption(5) = Val(fieldValue(8))	     'Working distance in inches
    vdOption(6) = -Int(Val(fieldValue(11)))  'Fault clearing:-1- Auto;-2- manual clearing Time;-3- Step Event Analysis 
    If vdOption(6) = -2 Then 
      vdOption(7) = Val(fieldValue(12))      'manual clearing Time in seconds
    Else
      vdOption(7) = Val(fieldValue(9))	     'Breaker interrupting time in cycles
    End If
    vdOption(8) = Int(Val(fieldValue(10)))   'Ignore 2 second flag: 0-reset; 1-set;
    vdOption(9) = Int(Val(fieldValue(13)))   'Number of tiers to include in protective device list
    
    nBusHnd& = 0
    If sBusName <> "" Then Call findBusByName( sBusName, dBusKv, nBusHnd& )
    If nBusHnd = 0 And nBusNumber > 0 Then Call findBusByNumber( nBusNumber, nBusHnd& )
    If nBusHnd > 0 Then
      bOK = DoAcrFlash( nBusHnd, vdOption, vdResult ) 
      If bOK > 0 Then 
        Select Case vdOption(1)
          Case 0
            sEquipment = "Switchgear"
          Case 1
            sEquipment = "Cable"
          Case Else
            sEquiment = "open air"
        End Select  
        If vdOption(2) = 0 Then
          sGrounded = "No"
        Else
          sGrounded = "Yes"
        End If
        If vdOption(3) = 0 Then
          sEnclosed = "No"
        Else
          sEnclosed = "Yes"
        End If   
        Print #1,  FullBusName(nBusHnd) & "," & sEquipment & "," & sGrounded & "," & sEnclosed & _ 
                   "," & fieldValue(9) & "," & fieldValue(8) & "," & fieldValue(7) & "," & vdResult(1) & "," & vdResult(2) & _
                   "," & FullRelayName(vdResult(3)) & ",", vdResult(4) & "," & vdResult(5) & _
                   "," & FullRelayName(vdResult(6)) & "," & vdResult(7) & "," & vdResult(8) & _
                   "," & vdResult(9) & ",", vdResult(10) & "," & vdResult(11) & "," & vdResult(12) & "," & vdResult(13)& "," & vdResult(14) 
        nSuccess = nSuccess + 1
      End If
      nDone& = ii*100/rowCount
      sMsg$ = "Record " & Str(ii) & " of " & Str(rowCount)
    Else
      PrintTTY( "Bus not found:" & Str(nBusNumber) & " " & sBusName & Str(dBusKv) & " KV" )
    End If
    If 2 = ProgressDialog( 1, "Reading XLS data", sMsg$, nDone& ) Then exit For
  Next  
  Close 1
  Print "Arc-flash calculation ran successfully on " & Str(nSuccess) & " buses" & Chr(13) & Chr(10) & _
        "Results are in: " & OutputFile$
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
HasError:
  ' Free memory  
  Close 1
  Set dataSheet = Nothing
  Set wkbook    = Nothing
  Set xlApp     = Nothing
  If Err.Number > 0 Then Print "Excecution error: " & Err.Description
  Stop
End Sub

 ' Read data of the specified row in the spreadsheet
Function readXLSRow( ByRef aSheet As Object, ByVal rowNo As long, _
          ByRef outArray() As String, ByVal maxSize As long )  As long
  
  readXLSRow = 0        
  For Col = 1 To maxSize
    outArray(Col) = aSheet.Cells(rowNo,Col).Value
    If outArray(Col) <> "" Then readXLSRow  = readXLSRow  + 1
  Next
End Function

Function findBusByNumber( nBusNumber&, ByRef nBusHnd& ) As long
  findBusByNumber = 0
  nBusHnd = 0
  While 1 = GetEquipment( TC_BUS, nBusHnd )
    Call getdata( nBusHnd, BUS_nNumber, nNo& )
    If nNo = nBusNumber Then
      findBusByNumber = 1
      exit Function
    End If
  Wend
End Function
