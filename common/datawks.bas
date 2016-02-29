' ASPEN PowerScript sample program
'
' DATAWKS.BAS
'
' Version 2.1
'
' Import network data from excel spreadsheet 
'
' Spreadsheet must be in the same directory as OLR file. 
'
' Spreadsheet must have:
' 1- FILESIGNATURE in cell A1
' 2- aType_XXX in cell A2
' 3- Header in row 3 as in the corresponding data browser table
' 4- Data rows in same format as in the corresponding browser table
' 5- Key columns must present. All other columns are optional
'
'
'
' Global vars and consts
Const DebugDataPath$ = "c:\Source\bas\dev\datawks\"
Const FILESIGNATURE$ = "ASPEN OneLiner/Power Flow"
Const aType_Line$    = "Browser Report for Lines"
Const aType_Mu$      = "Browser Report for Zero-Sequence Mutuals"
Const aType_Xfmr$    = "Browser Report for Transformers: 2-Winding"
Const aType_Xfmr3$   = "Browser Report for Transformers: 3-Winding"
Const aType_Ps$      = "Browser Report for Phase Shifters"
Const aType_Gen$     = "Browser Report for Generators"
Const aType_Load$    = "Browser Report for Loads"
Const aType_Shc$     = "Browser Report for Shunts"

dim sLabels(50) As String
dim nCodes(50) As long     
dim nCountCodes As long
Dim xlApp As Object     ' Declare variable to hold the reference.
Dim wkbook As Object    ' Declare variable to hold the reference.
Dim dataSheet As Object

Sub main
  Dim outArr(50) As String
  dim fieldNames(50) As String 
  
  DataFile$ = InputBox("Enter excel file name")
  
  If Len(DataFile) = 0 Then 
    Print "Bye"
    Stop
  End If
  
  ExcelFile$ = GetOLRFilePath()  + DataFile
  

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
    
  ' Read file header row
  aHeader$ = dataSheet.Cells(1,1).Value
  aType$   = dataSheet.Cells(4,1).Value
  If aType = "" Or InStr(1, aHeader, FILESIGNATURE) <> 1 Then
    Print "Invalid ASPEN OneLiner Data file"
    GoTo endProgram
  End If
  
  ' Count rows in spreadsheet and read field names
  colCount = readXLSRow( dataSheet, 6, fieldNames, 49 )
  If Not checkHeader( fieldNames, colCount, aType ) Then 
     Print aType + " header row is not recognized"
     GoTo endProgram
  End If
  
  For ii = 1 to colCount 
    fieldNames(ii) = Trim(fieldNames(ii))
  Next

  ' Count rows in spreadsheet
  rowCount& = 0
  Do
    aVal$ = dataSheet.Cells(8+rowCount,1).Value & dataSheet.Cells(8+rowCount,2).Value & dataSheet.Cells(8+rowCount,3).Value
    If "" = aVal$ Then exit Do
    rowCount = rowCount + 1
  Loop While true

  If rowCount = 0 Then
    Print "Table has no data row"
    GoTo endProgram
  End If 
  
  ' Process the spreadsheet row by row
  nUpdated = 0
  nFailed  = 0
  For ii& = 1 to rowCount 
    Call readXLSRow( dataSheet, 7+ii, outArr, 49 )
    If 0 < processRow( fieldNames, outArr, colCount, aType ) Then 
      nUpdated = nUpdated + 1
    Else
      nFailed = nFailed + 1
    End If
    nDone& = ii*100/rowCount
    sMsg$ = "Record " & Str(ii) & " of " & Str(rowCount)
    If 2 = ProgressDialog( 1, "Reading XLS data", sMsg$, nDone& ) Then exit For
  Next
  If nUpdated > 0 Then _
    Print nUpdated & " records were updated successfully. Details are in TTY"
  If nFailed > 0 Then _
    Print nFailed & " records failed to update. See details in TTY"
  
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

 'Find the directory of the spreadsheet
Function GetOLRFilePath() As String
  FilePath$ = GetOLRFileName()
  For ii = Len(FilePath) to 1 step -1
    If Mid(FilePath, ii, 1) = "\" Then
      FilePath = Left(FilePath, ii)
      exit For
    End If
  Next
  GetOLRFilePath = FilePath
End Function
 
 'Identify the data type stored in the spreadsheet
Function checkHeader( ByRef Array() As String, ByVal colCount, ByRef aType As String )
  checkHeader = false
  select case aType
    case aType_Line
      checkHeader = checkHeader_Lines( Array, colcount )
      If checkHeader Then nCountCodes& = InitParamCode_Line(sLabels,nCodes)
    case aType_Mu
      checkHeader = checkHeader_Mu( Array, colcount )
      If checkHeader Then nCountCodes& = InitParamCode_Mu(sLabels,nCodes)
    case aType_Xfmr
      checkHeader = checkHeader_Xfmr( Array, colcount )
      If checkHeader Then nCountCodes& = InitParamCode_Xfmr(sLabels,nCodes)
    case aType_Xfmr3
      checkHeader = checkHeader_Xfmr3( Array, colcount )
      If checkHeader Then nCountCodes& = InitParamCode_Xfmr3(sLabels,nCodes) 
    case aType_Ps
      checkHeader = checkHeader_Ps( Array, colcount )
      If checkHeader Then nCountCodes& = InitParamCode_Ps(sLabels,nCodes)
    case aType_Gen
      checkHeader = checkHeader_Gen( Array, colcount )
      If checkHeader Then nCountCodes& = InitParamCode_Gen(sLabels,nCodes)  
    case aType_Load
      checkHeader = checkHeader_Load( Array, colcount )
      If checkHeader Then nCountCodes& = InitParamCode_Load(sLabels,nCodes) 
    case aType_Shc
      checkHeader = checkHeader_Sh( Array, colcount )
      If checkHeader Then nCountCodes& = InitParamCode_Sh(sLabels,nCodes)   
  End select
End Function

 ' Read data of the specified row in the spreadsheet
Function readXLSRow( ByRef aSheet As Object, ByVal rowNo As long, _
          ByRef outArray() As String, ByVal maxSize As long )  As long
  
  readXLSRow = 0        
  For Col = 1 To maxSize
    outArray(Col) = aSheet.Cells(rowNo,Col).Value
    If outArray(Col) <> "" Then readXLSRow  = readXLSRow  + 1
  Next
End Function

  ' Process the data row by row according to the data type stored in the spreadsheet
Function processRow( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols, ByRef aType$ ) As long
  processRow = 0
  ' Print log
  aText$ = ""
  For ii = 1 to cols
    If ii > 1 Then aText$ = aText$ & ";"
    If FieldName(ii) = "Line 1" Or FieldName(ii) = "Line 2" Then
      nPos1 = InStr(1,FieldVal(ii)," ")
      branch1_bus1Num$  = Trim(Left(FieldVal(ii),nPos1))
      nPos = InStr(1,FieldVal(ii)," - ")    
      branch1_bus1Name$ = Trim(Mid(FieldVal(ii),nPos1,nPos-nPos1))
      nLen = Len(FieldVal(ii))     
      nLen1 = Len(branch1_bus1Name)    
      branch1_bus2Num$  = Trim(Mid(FieldVal(ii),nPos+2,7))
      branch1_bus2Name$ = Trim(Mid(FieldVal(ii),nPos+10,nLen-nPos-10))
      nLen2 = Len(branch1_bus2Name)
      CktID1           = Trim(Right(FieldVal(ii),4))
      aText$ = aText & branch1_bus1Num & " " & Trim(Left(branch1_bus1Name, nLen1-7)) & Right(branch1_bus1Name,7) & " " _
                     & branch1_bus2Num & " " & Trim(Left(branch1_bus2Name, nLen1-7)) & Right(branch1_bus2Name,7) & " " & CktID1  
    ElseIf FieldName(ii) = "Bus 1" Or FieldName(ii) = "Bus 2" Or FieldName(ii) = "Bus 3" Or FieldName(ii) = "Bus Name" Then
      nLen = Len(FieldVal(ii))
      aText$ = aText & Trim(Left(FieldVal(ii), nLen-7)) & " " & Right(FieldVal(ii),7)
    Else
      aText$ = aText & FieldVal(ii)
    End If
  Next
  nLen& = Len(aType) + 4
  If (Len(aText)+nLen-20) <130 Then
    printTTY Mid(aType,20,nLen-20) & " data:" & aText
  Else
    printTTY Mid(aType,20,nLen-20) & " data:" & Left(aText, 126-nLen+20) & "..."
  End If
  select case aType
    case aType_Line
      processRow = processRow_Lines( FieldName, FieldVal, cols )
    case aType_Mu
      processRow  = processRow_Mu( FieldName, FieldVal, cols )
    case aType_Xfmr
      processRow = processRow_Xfmr( FieldName, FieldVal, cols )
    case aType_Xfmr3
      processRow = processRow_Xfmr3( FieldName, FieldVal, cols )
    case aType_Ps
      processRow = processRow_Ps( FieldName, FieldVal, cols )  
    case aType_Gen
      processRow = processRow_Gen( FieldName, FieldVal, cols )
    case aType_Load
      processRow = processRow_Load( FieldName, FieldVal, cols )
    case aType_Shc
      processRow = processRow_Sh( FieldName, FieldVal, cols )  
  End select
End Function

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

 ' Find parameter ID
Function LookupParamCode( l() As String, c() As long, nLen&, sLabel$, ByVal jj ) As long
  LookupParamCode = 0
  nIteration = 0
  For ii = 1 to nLen
    If l(ii) = sLabel Then
      nIteration = nIteration + 1
      If nIteration = jj Then 
        LookupParamCode = c(ii)
        exit Function
      End If      
    End If
  Next
End Function

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
        If myID = CktID Or (Len(myID) = 0 And Len(CktID) = 0) Then
          branchSearch = thisItemHnd
          exit Do
        End If
      End If
    End If
  Wend
End Function

 ' Find mutual line handle
Function muSearch( branch1Hnd&, branch2Hnd& )
  muSearch = 0
  MuHnd&   = 0
  While GetData( branch1Hnd&, LN_nMuPairHnd, MuHnd& ) > 0 
    Call GetData( MuHnd&, MU_nHndLine1,  nHndLine1& )
    Call GetData( MuHnd&, MU_nHndLine2,  nHndLine2& )
    If branch1Hnd = nHndLine1 And branch2Hnd = nHndLine2 Then
      muSearch = MuHnd
      exit Do
    End If
  Wend
End Function

 ' Find generator unit handle
Function genunitSearch( busHnd&, CktID$ )
  genunitSearch = 0
  genunitHnd&   = 0
  While GetBusEquipment( busHnd, TC_GENUNIT, genunitHnd& ) > 0
    Call GetData(genunitHnd, GU_sID, myID$)
    myID = Trim(myID)
    If myID = CktID Or (Len(myID) = 0 And Len(CktID) = 0) Then
      genunitSearch = genunitHnd
      exit Do
    End If
  Wend
End Function

 ' Find generator handle
Function genSearch( busHnd& )
  genSearch = 0
  genHnd&   = 0
  If GetBusEquipment( busHnd, TC_GEN, genHnd& ) > 0  Then
    genSearch = genHnd&
  End If
End Function

 ' Find load unit handle
Function loadunitSearch( busHnd&, CktID$ )
  loadunitSearch = 0
  loadunitHnd&   = 0
  While GetBusEquipment( busHnd, TC_LOADUNIT, loadunitHnd& ) > 0
    Call GetData(loadunitHnd, LU_sID, myID$)
    myID = Trim(myID)
    If myID = CktID Or (Len(myID) = 0 And Len(CktID) = 0) Then
      loadunitSearch = loadunitHnd
    End If
  Wend
End Function

 ' Find shunt unit handle 
Function shunitSearch( busHnd&, CktID$ )
  shunitSearch = 0
  shunitHnd&   = 0
  While GetBusEquipment( busHnd, TC_SHUNTUNIT, shunitHnd& ) > 0
    Call GetData(shunitHnd, SU_sID, myID$)
    myID = Trim(myID)
    If myID = CktID Or (Len(myID) = 0 And Len(CktID) = 0) Then
      shunitSearch = shunitHnd
    End If
  Wend
End Function

 ' Find shunt handle
Function shSearch( busHnd& )
  shSearch = 0
  shHnd&   = 0
  If GetBusEquipment( busHnd, TC_SHUNT, shHnd& ) > 0 Then
    shSearch = shHnd&
  End If
End Function

 ' Set field value 
Function SetFieldValue( sLabels() As String, nCodes() As long, nCountCodes&, sFieldVal$, thisHnd&, paramID&, nIndex)
  SetFieldValue = 0
  Dim vdArray(5) As Double
  paramType& = paramID/1000
  paramType& = paramID - paramType*1000
  paramType& = paramType/100
  select case paramType&
    case 1 ' String
      sVal$ = sFieldVal
      Call GetData( thisHnd, paramID, sValtemp$ )
      If sVal <> sValtemp Then SetFieldValue = SetData( thisHnd, paramID, sVal$ )
    case 2 ' double     
      dVal# = a2d(sFieldVal)
      Call GetData( thisHnd, paramID, dValtemp# )
      If Abs(dVal - dValtemp) > 0.00001 Then SetFieldValue = SetData( thisHnd, paramID, dVal# )     
      Call PostData(thisHnd) 
    case 3 ' Integer
      If UCase(sFieldVal) = "YES" Then 
        nVal& = 1 
      ElseIf UCase(sFieldVal) = "NO" Then
        nVal& = 2 
      ElseIf UCase(sFieldVal) = "PV" Then
        nVal& = 0
      Elseif UCase(sFieldVal) = "PQ" Then
        nVal& = 1    
      Else
        nVal& = Val(sFieldVal)
      End If 
      Call GetData( thisHnd, paramID, nValtemp& )
      If nVal <> nValtemp Then SetFieldValue = SetData( thisHnd, paramID, nVal& )
    case -5 ' Array
      Call GetData( thisHnd, paramID, vdArray() )
      If Abs(vdArray(nIndex) - a2d(sFieldVal)) > 0.000001 Then
        vdArray(nIndex) = a2d(sFieldVal)
        SetFieldValue = SetData( thisHnd, paramID, vdArray() ) 
      End If
      Call PostData(thisHnd)       
  End select
End Function

 ' Convert string to double
Function a2d(v$) As double
  a2d = 0.0
  v   = Trim(v)
  If Len(v) = 0 Then exit Function
  If IsNumeric(v) Then
    a2d = Val(v)
  Else
    v = UCase(v)
    nPos = InStr(1,v,"E")
    If nPos > 0 Then
      aPart1$ = Left(v,nPos-1)
      aPart2$ = Mid(v,nPos+1,99)
      If Len(aPart1) > 0 And Len(aPart2) > 0 And IsNumeric(aPart1) And IsNumeric(aPart2) Then
        a2d = Val(aPart1) * 10^Val(aPart2)
      End If
    End If
  End If
End Function
 
 ' Convert string to complex number [real (Mode = 1) or imaginary (Mode = 2) part]
Function c2ri(v$, ByVal Mode) As String
  c2ri = ""
  v =Trim(v)
  If Len(v) = 0 Then exit Function
  nPos = InStr(1,v,"+")
  Select Case Mode
    case 1 
      c2ri = Left(v,nPos-1)
    case 2
      c2ri = Mid(v,nPos+2,99)
  End Select
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''Line'''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InitParamCode_Line( l() As String, c() As long ) As long
  l(1)  = "In Serv"
  l(2)  = "Name"
  l(3)  = "Length"
  l(4)  = "Type"
  l(5)  = "R"
  l(6)  = "X"
  l(7)  = "R0"
  l(8)  = "X0"
  l(9)  = "G1"
  l(10) = "B1"
  l(11) = "G2"
  l(12) = "B2"
  l(13) = "G10"
  l(14) = "B10"
  l(15) = "G20"
  l(16) = "B20"
  l(17) = "Rating A"
  l(18) = "Rating B"
  l(19) = "Rating C"
  l(20) = "Rating D"
  
  c(1)  = LN_nInService
  c(2)  = LN_sName
  c(3)  = LN_dLength
  c(4)  = 0  'LN_sType
  c(5)  = LN_dR
  c(6)  = LN_dX
  c(7)  = LN_dR0
  c(8)  = LN_dX0
  c(9)  = LN_dG1
  c(10) = LN_dB1
  c(11) = LN_dG2
  c(12) = LN_dB2
  c(13) = LN_dG10
  c(14) = LN_dB10
  c(15) = LN_dG20
  c(16) = LN_dB20
  c(17) = LN_vdRating 
  c(18) = LN_vdRating
  c(19) = LN_vdRating
  c(20) = LN_vdRating
  
  InitParamCode_Line = 20
End Function
Function checkHeader_Lines( ByRef Array() As String, ByVal colCount )
  okBus1  = false
  okBus2  = false
  okCktID = false
  checkHeader_Lines = false
  For ii = 1 to colCount
    If Array(ii) = "Bus 1" Then okBus1  = true
    If Array(ii) = "Bus 2" Then okBus2  = true
    If Array(ii) = "Ckt"   Then okCktID = true
    If okBus1 And okBus2 And okCktID Then
      checkHeader_Lines = true
      exit For
    End If
  Next
End Function
Function processRow_Lines( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols ) As long
  processRow_Lines = 0
  ' Find object handle
  bus1Hnd& = -1
  bus2Hnd& = -1
  CktID$   = "999"
  branchHnd& = 0
  For ii = 1 to cols
    If FieldName(ii) = "Bus 1" Then bus1Hnd = findBusHnd(FieldVal(ii))
    If FieldName(ii) = "Bus 2" Then bus2Hnd = findBusHnd(FieldVal(ii))
    If FieldName(ii) = "Ckt"   Then CktID   = FieldVal(ii)
    If bus1Hnd> -1 And bus2Hnd > -1 And CktID <> "999" Then
      branchHnd& = branchSearch( TC_LINE, bus1Hnd, bus2Hnd, 0, CktID )
      exit For
    End If
  Next
  If branchHnd = 0 Then 
    printTTY("  Error: Object not found")
    exit Function
  End If
  countUpdated = 0
  listUpdated$ = ""
  For ii = 1 to cols
    nIndex = 0
    sFieldVal$ = FieldVal(ii)
    paramID = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),1)
    If paramID = 0 Then 
      GoTo NextIteration
    End If
    select case paramID
      case LN_vdRating
        select case FieldName(ii)
          case "Rating A"
            nIndex = 1
          case "Rating B" 
            nIndex = 2
          case "Rating C"
            nIndex = 3
          case "Rating D"
            nIndex = 4 
        End select
      case LN_dLength 
        nPos   = InStr(sFieldVal," ")
        If nPos > 0 Then
          aPart$ = Mid(sFieldVal, nPos+1,99)
          bPart$ = Left(sFieldVal, nPos-1)
          If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,aPart,branchHnd,LN_sLengthUnit,nIndex) And _
             0 < SetFieldValue(sLabels,nCodes,nCountCodes,bPart,branchHnd,paramID&,nIndex) Then
            countUpdated = countUpdated + 1
            If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
          End If
        End If  
        GoTo NextIteration     
    End select    

    If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,branchHnd,paramID&,nIndex) Then
      countUpdated = countUpdated + 1
      If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
    End If
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(branchHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString())
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Lines = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''Mutual''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InitParamCode_Mu( l() As String, c() As long ) As long
  l(1)  = "R"
  l(2)  = "X"
  l(3)  = "From Ln.1"
  l(4)  = "To Ln.1"
  l(5)  = "From Ln.2"
  l(6)  = "To Ln.2"
  
  c(1)  = MU_dR
  c(2)  = MU_dX
  c(3)  = MU_dFrom1
  c(4)  = MU_dTo1
  c(5)  = MU_dFrom2
  c(6)  = MU_dTo2
  
  InitParamCode_Mu = 6
End Function
Function checkHeader_Mu( ByRef Array() As String, ByVal colCount )
  okLine1  = false
  okLine2  = false
  checkHeader_Mu = false
  For ii = 1 to colCount
    If Array(ii) = "Line 1" Then okLine1  = true
    If Array(ii) = "Line 2" Then okLine2  = true
    If okLine1 And okLine2 Then
      checkHeader_Mu = true
      exit For
    End If
  Next
End Function
Function processRow_Mu( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols ) As long
  processRow_Mu = 0
  ' Find object handle
  muHnd&      = 0
  branch1Hnd& = 0
  branch2Hnd& = 0
  CktID1$     = "999"
  CktID2$     = "999"
  branch1_bus1Hnd& = -1
  branch1_bus2Hnd& = -1
  branch2_bus1Hnd& = -1
  branch2_bus2Hnd& = -1
  
  For ii = 1 to cols 
    If FieldName(ii) = "Line 1" Then 
      nPos1 = InStr(1,FieldVal(ii)," ")
      nPos = InStr(1,FieldVal(ii)," - ")  
      nLen = Len(FieldVal(ii))  
      branch1_bus1Name$ = Trim(Mid(FieldVal(ii),nPos1,nPos-nPos1))                 
      branch1_bus2Name$ = Trim(Mid(FieldVal(ii),nPos+10,nLen-nPos-10))
      branch1_bus1Hnd  = findBusHnd(branch1_bus1Name)
      branch1_bus2Hnd  = findBusHnd(branch1_bus2Name)
      CktID1           = Trim(Right(FieldVal(ii),4))         
    End If
    If FieldName(ii) = "Line 2" Then 
      nPos1 = InStr(1,FieldVal(ii)," ")
      nPos = InStr(1,FieldVal(ii)," - ") 
      nLen = Len(FieldVal(ii))    
      branch2_bus1Name$ = Trim(Mid(FieldVal(ii),nPos1,nPos-nPos1))           
      branch2_bus2Name$ = Trim(Mid(FieldVal(ii),nPos+10,nLen-nPos-10))
      branch2_bus1Hnd  = findBusHnd(branch2_bus1Name)
      branch2_bus2Hnd  = findBusHnd(branch2_bus2Name)
      CktID2           = Trim(Right(FieldVal(ii),4))
    End If
    If branch1_bus1Hnd> -1 And branch1_bus2Hnd > -1 And branch2_bus1Hnd > -1 And branch2_bus2Hnd > -1 Then
      branch1Hnd& = branchSearch( TC_LINE, branch1_bus1Hnd, branch1_bus2Hnd, 0, CktID1 )
      branch2Hnd& = branchSearch( TC_LINE, branch2_bus1Hnd, branch2_bus2Hnd, 0, CktID2 )
      muHnd&      = muSearch( branch1Hnd, branch2Hnd )
      exit For
    End If
  Next
  If muHnd = 0 Then 
    printTTY("  Error: Object not found")
    exit Function
  End If
  countUpdated = 0
  listUpdated$ = ""
  countMuUpdated = 0
  For ii = 1 to cols
    nIndex = 0
    sFieldVal$ = FieldVal(ii)
    paramID = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),1)
    If paramID = 0 Then 
      GoTo NextIteration
    End If
    
    If paramID = MU_dFrom1 Or paramID = MU_dTo1 Or paramID = MU_dFrom2 Or paramID = MU_dTo2 Then
      dVal# = Val(sFieldVal)
      dVal# = dVal#*100
      sFieldVal = Str(dVal#)
    End If

    If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,muHnd,paramID&,nIndex) Then
      countUpdated = countUpdated + 1
      If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
    End If
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(muHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString())
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Mu = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''2-Winding Transformer'''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InitParamCode_Xfmr( l() As String, c() As long ) As long
  l(1)  = "In Serv"
  l(2)  = "Tap1 V(pu)"
  l(3)  = "Tap2 V(pu)"  
  l(4)  = "R"
  l(5)  = "X"
  l(6)  = "B"
  l(7)  = "R0"
  l(8)  = "X0"  
  l(9)  = "B0"
  l(10) = "G1" 
  l(11) = "B1"
  l(12) = "G10"
  l(13) = "B10"
  l(14) = "G2"
  l(15) = "B2"
  l(16) = "G20"
  l(17) = "B20"
  l(18) = "ZG1"
  l(19) = "ZG1"
  l(20) = "ZG2"
  l(21) = "ZG2"
  l(22) = "ZGN"
  l(23) = "ZGN"

  c(1)  = XR_nInService
  c(2)  = XR_dTap1
  c(3)  = XR_dTap2  
  c(4)  = XR_dR
  c(5)  = XR_dX
  c(6)  = XR_dB
  c(7)  = XR_dR0
  c(8)  = XR_dX0  
  c(9)  = XR_dB0
  c(10) = XR_dG1  
  c(11) = XR_dB1
  c(12) = XR_dG10
  c(13) = XR_dB10
  c(14) = XR_dG2
  c(15) = XR_dB2
  c(16) = XR_dG20  
  c(17) = XR_dB20
  c(18) = XR_dRG1
  c(19) = XR_dXG1
  c(20) = XR_dRG2
  c(21) = XR_dXG2
  c(22) = XR_dRGN
  c(23) = XR_dXGN

  InitParamCode_Xfmr = 23
End Function
Function checkHeader_Xfmr( ByRef Array() As String, ByVal colCount )
  okBus1  = false
  okBus2  = false
  okCktID = false
  checkHeader_Xfmr = false
  For ii = 1 to colCount
    If Array(ii) = "Bus 1" Then okBus1  = true
    If Array(ii) = "Bus 2" Then okBus2  = true
    If Array(ii) = "Ckt"   Then okCktID = true
    If okBus1 And okBus2 And okCktID Then
      checkHeader_Xfmr = true
      exit For
    End If
  Next
End Function
Function processRow_Xfmr( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols ) As long
  processRow_Xfmr = 0
  bus1Hnd& = -1
  bus2Hnd& = -1
  CktID$   = "999"
  branchHnd& = 0
  For ii = 1 to cols
    If FieldName(ii) = "Bus 1" Then bus1Hnd = findBusHnd(FieldVal(ii))
    If FieldName(ii) = "Bus 2" Then bus2Hnd = findBusHnd(FieldVal(ii))
    If FieldName(ii) = "Ckt"   Then CktID   = FieldVal(ii)
    If bus1Hnd> -1 And bus2Hnd > -1 And CktID <> "999" Then
      branchHnd& = branchSearch( TC_XFMR, bus1Hnd, bus2Hnd, 0, CktID )
      exit For
    End If
  Next
  If branchHnd = 0 Then 
    printTTY("  Error: Object not found")
    exit Function
  End If
  countUpdated = 0
  listUpdated$ = ""
  For ii = 1 to cols
    nIndex = 0
    sFieldVal$ = FieldVal(ii)
    paramID = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),1)
    If paramID = 0 Or sFieldVal = "NA"Then 
      GoTo NextIteration
    End If
    Select case paramID
      case XR_dTap1, XR_dTap2
        nPos = InStr(sFieldVal, "R")
        If nPos > 0 Then sTemp = Left(sFieldVal,nPos-1)
        If  paramID = XR_dTap1 Then Call GetData(branchHnd, XR_nBus1Hnd, busHnd&) Else Call GetData(branchHnd, XR_nBus2Hnd, busHnd&)
        Call GetData(busHnd, BUS_dKVNominal, kvNom#)
        dVal# = a2d(sFieldVal) * kvNom
        sFieldVal = Str(dval#)         
      case XR_dRG1,XR_dRG2,XR_dRGN     
        paramIDX = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),2)
        Real$ = c2ri(sFieldVal,1)
        Imag$ = c2ri(sFieldVal,2)    
        If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,Real,branchHnd,paramID&,nIndex) Or _
           0 < SetFieldValue(sLabels,nCodes,nCountCodes,Imag,branchHnd,paramIDX&,nIndex) Then
          countUpdated = countUpdated + 1
        If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
        End If
        GoTo NextIteration
    End select

    If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,branchHnd,paramID&,nIndex) Then
      countUpdated = countUpdated + 1
      If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
    End If
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(branchHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString())
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Xfmr = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''3-Winding Transformer'''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InitParamCode_Xfmr3( l() As String, c() As long ) As long
  l(1)  = "In Serv"
  l(2)  = "Base MVA"
  l(3)  = "MVA A"
  l(4)  = "MVA B"
  l(5)  = "MVA C"
  l(6)  = "Tap1 V(pu)"
  l(7)  = "Tap2 V(pu)"
  l(8)  = "Tap3 V(pu)"
  l(9)  = "Zps"
  l(10) = "Zps"
  l(11) = "Zpt"
  l(12) = "Zpt"
  l(13) = "Zst"
  l(14) = "Zst"
  l(15) = "Zps0"
  l(16) = "Zps0"
  l(17) = "Zpt0"
  l(18) = "Zpt0"
  l(19) = "Zst0"  
  l(20) = "Zst0" 
  l(21) = "Zg1"
  l(22) = "Zg1"
  l(23) = "Zg2"
  l(24) = "Zg2"
  l(25) = "Zg3"
  l(26) = "Zg3"
  l(27) = "Zgn"
  l(28) = "Zgn"
  l(29) = "B"
  l(30) = "B0"

  c(1)  = X3_nInService
  c(2)  = X3_dBaseMVA
  c(3)  = X3_dMVA1
  c(4)  = X3_dMVA2
  c(5)  = X3_dMVA3
  c(6)  = X3_dTap1
  c(7)  = X3_dTap2
  c(8)  = X3_dTap3
  c(9)  = X3_dRps
  c(10) = X3_dXps
  c(11) = X3_dRpt
  c(12) = X3_dXpt
  c(13) = X3_dRst
  c(14) = X3_dXst  
  c(15) = X3_dR0ps
  c(16) = X3_dX0ps
  c(17) = X3_dR0pt
  c(18) = X3_dX0pt
  c(19) = X3_dR0st
  c(20) = X3_dX0st
  c(21) = X3_dRG1
  c(22) = X3_dXG1
  c(23) = X3_dRG2
  c(24) = X3_dXG2
  c(25) = X3_dRG3
  c(26) = X3_dXG3
  c(27) = X3_dRGN
  c(28) = X3_dXGN
  c(29) = X3_dB
  c(30) = X3_dB0

'  c(27) = X3_dLTCCenterTap
'  c(28) = X3_dLTCstep
'  c(29) = X3_dMaxTap
'  c(30) = X3_dMaxVW
'  c(31) = X3_nLTCPriority
'  c(32) = X3_nLTCGanged
  

  InitParamCode_Xfmr3 = 30
End Function
Function checkHeader_Xfmr3( ByRef Array() As String, ByVal colCount )
  okBus1  = false
  okBus2  = false
  okBus3  = false
  okCktID = false
  checkHeader_Xfmr3 = false
  For ii = 1 to colCount
    If Array(ii) = "Bus 1" Then okBus1  = true
    If Array(ii) = "Bus 2" Then okBus2  = true
    If Array(ii) = "Tertiary" Then okBus3  = true
    If Array(ii) = "Ckt"   Then okCktID = true
    If okBus1 And okBus2 And okBus3 And okCktID Then
      checkHeader_Xfmr3 = true
      exit For
    End If
  Next
End Function
Function processRow_Xfmr3( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols ) As long
  processRow_Xfmr3 = 0
  ' Find object handle
  bus1Hnd& = -1
  bus2Hnd& = -1
  bus3Hnd& = -1
  CktID$   = "999"
  branchHnd& = 0
  For ii = 1 to cols
'    aField$ = FieldName(ii)
'    If "53-330-T2    400. kV" = Trim(FieldVal(ii)) Then Print "AAA"
    If FieldName(ii) = "Bus 1" Then bus1Hnd = findBusHnd(FieldVal(ii))
    If FieldName(ii) = "Bus 2" Then bus2Hnd = findBusHnd(FieldVal(ii))
    If FieldName(ii) = "Tertiary" Then bus3Hnd = findBusHnd(FieldVal(ii))
    If FieldName(ii) = "Ckt"   Then CktID   = FieldVal(ii)
    If bus1Hnd> -1 And bus2Hnd > -1 And bus3Hnd > -1 And CktID <> "999" Then
      branchHnd& = branchSearch( TC_XFMR3, bus1Hnd, bus2Hnd, bus3Hnd, CktID )
      exit For
    End If
  Next
  If branchHnd = 0 Then 
    printTTY("  Error: Object not found")
    exit Function
  End If
  countUpdated = 0
  listUpdated$ = ""
  For ii = 1 to cols
    nIndex = 0
    sFieldVal$ = FieldVal(ii)
    paramID = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),1)
    If paramID = 0 Or sFieldVal = "N/A"Then 
      GoTo NextIteration
    End If
    Select case paramID    
      case X3_dTap1, X3_dTap2, X3_dTap3
        nPos = InStr(sFieldVal, "R")
        If nPos > 0 Then sTemp = Left(sFieldVal,nPos-1)
        If paramID = X3_dTap1 Then 
          Call GetData(branchHnd, X3_nBus1Hnd, busHnd&) 
        ElseIf paramID = X3_dTap2 Then
          Call GetData(branchHnd, X3_nBus2Hnd, busHnd&)
        Else
          Call GetData(branchHnd, X3_nBus3Hnd, busHnd&)
        End If
        Call GetData(busHnd, BUS_dKVNominal, kvNom#)
        dVal# = a2d(sFieldVal) * kvNom
        sFieldVal = Str(dval#)         
      case X3_dR0ps,X3_dR0pt,X3_dR0st,X3_dRG1,X3_dRG2,X3_dRG3,X3_dRGN,X3_dRps,X3_dRpt,X3_dRst   
        paramIDX = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),2)
        Real$ = c2ri(sFieldVal,1)
        Imag$ = c2ri(sFieldVal,2)    
        If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,Real,branchHnd,paramID&,nIndex) Or _
           0 < SetFieldValue(sLabels,nCodes,nCountCodes,Imag,branchHnd,paramIDX&,nIndex) Then
          countUpdated = countUpdated + 1
        If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)         
        End If
        GoTo NextIteration
    End select
    If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,branchHnd,paramID&,nIndex) Then
      countUpdated = countUpdated + 1
      If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
    End If
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(branchHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString())
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Xfmr3 = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''Phase Shifter'''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InitParamCode_Ps( l() As String, c() As long ) As long
  l(1)  = "In Serv"
  l(2)  = "Name"
  l(3)  = "Angle"
  l(4)  = "R"
  l(5)  = "X"
  l(6)  = "B"
  l(7)  = "R2"
  l(8)  = "X2"
  l(9)  = "B2"
  l(10) = "R0"
  l(11) = "X0"
  l(12) = "B0"
  l(13) = "Angle Min"
  l(14) = "Angle Max."
  l(15) = "Targ.MW Min"
  l(16) = "Targ.MW Max"
  
  c(1)  = PS_nInService
  c(2)  = PS_sName
  c(3)  = PS_dAngle
  c(4)  = PS_dR
  c(5)  = PS_dX
  c(6)  = PS_dB
  c(7)  = PS_dR2
  c(8)  = PS_dX2
  c(9)  = PS_dB2
  c(10) = PS_dR0
  c(11) = PS_dX0
  c(12) = PS_dB0
  c(13) = PS_dAngleMin
  c(14) = PS_dAngleMax
  c(15) = PS_dMWmin
  c(16) = PS_dMWmax
  
  InitParamCode_Ps = 16
End Function
Function checkHeader_Ps( ByRef Array() As String, ByVal colCount )
  okBus1  = false
  okBus2  = false
  okCktID = false
  checkHeader_Ps = false
  For ii = 1 to colCount
    If Array(ii) = "Bus 1" Then okBus1  = true
    If Array(ii) = "Bus 2" Then okBus2  = true
    If Array(ii) = "Ckt"   Then okCktID = true
    If okBus1 And okBus2 And okCktID Then
      checkHeader_Ps = true
      exit For
    End If
  Next
End Function
Function processRow_Ps( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols ) As long
  processRow_Ps = 0
  ' Find object handle
  bus1Hnd& = -1
  bus2Hnd& = -1
  CktID$   = "999"
  branchHnd& = 0
  For ii = 1 to cols
    If FieldName(ii) = "Bus 1" Then bus1Hnd = findBusHnd(FieldVal(ii))
    If FieldName(ii) = "Bus 2" Then bus2Hnd = findBusHnd(FieldVal(ii))
    If FieldName(ii) = "Ckt"   Then CktID   = FieldVal(ii)
    If bus1Hnd> -1 And bus2Hnd > -1 And CktID <> "999" Then
      branchHnd& = branchSearch( TC_PS, bus1Hnd, bus2Hnd, 0, CktID )
      exit For
    End If
  Next
  If branchHnd = 0 Then 
    printTTY("  Error: Object not found")
    exit Function
  End If
  countUpdated = 0
  listUpdated$ = ""
  For ii = 1 to cols
    nIndex = 0
    sFieldVal$ = FieldVal(ii)
    paramID = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),1)
    If paramID = 0 Then 
      GoTo NextIteration
    End If

    If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,branchHnd,paramID&,nIndex) Then
      countUpdated = countUpdated + 1
      If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
    End If
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(branchHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString())
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Ps = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''Generator'''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InitParamCode_Gen( l() As String, c() As long ) As long
  l(1)  = "In Serv."
  l(2)  = "Ref. V"
  l(3)  = "Ref.Ang."
  l(4)  = "Regulation"
  l(5)  = "Limit A"
  l(6)  = "Limit B"
  l(7)  = "Subtransient"
  l(8)  = "Subtransient"
  l(9)  = "Transient"
  l(10) = "Transient"
  l(11) = "Synchronous"
  l(12) = "Synchronous"
  l(13) = "Neg. Seq."
  l(14) = "Neg. Seq."
  l(15) = "Zero Seq."
  l(16) = "Zero Seq."
  l(17) = "Neutral Imp."
  l(18) = "Neutral Imp."
  l(19) = "P Min"
  l(20) = "P Max"
  l(21) = "Q Min"
  l(22) = "Q Max"

  c(1)  = GE_nActive
  c(2)  = GE_dVSourcePU
  c(3)  = GE_dRefAngle
  c(4)  = GE_nFixedPQ
  c(5)  = GE_dCurrLimit1
  c(6)  = GE_dCurrLimit2
  c(7)  = GU_vdR
  c(8)  = GU_vdX
  c(9)  = GU_vdR
  c(10) = GU_vdX
  c(11) = GU_vdR
  c(12) = GU_vdX
  c(13) = GU_vdR
  c(14) = GU_vdX
  c(15) = GU_vdR
  c(16) = GU_vdX
  c(17) = GU_dRz
  c(18) = GU_dXz
  c(19) = GU_dPmin
  c(20) = GU_dPmax
  c(21) = GU_dQmin
  c(22) = GU_dQmax
  
  InitParamCode_Gen = 22
End Function
Function checkHeader_Gen( ByRef Array() As String, ByVal colCount )
  okBusName  = false
  okCktID = false
  checkHeader_Gen = false
  For ii = 1 to colCount
    If Array(ii) = "Bus Name" Then okBusName  = true
    If Array(ii) = "ID"   Then okCktID = true
    If okBusName And okCktID Then
      checkHeader_Gen = true
      exit For
    End If
  Next
End Function
Function processRow_Gen( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols ) As long
  processRow_Gen = 0
  ' Find object handle
  busHnd& = -1
  CktID$   = "999"
  genHnd& = 0
  genunitHnd& = 0
  sFieldVal = ""
  For ii = 1 to cols
    If FieldName(ii) = "Bus Name" Then busHnd = findBusHnd(FieldVal(ii))
    If FieldName(ii) = "ID"   Then CktID   = FieldVal(ii)
    If busHnd> -1 And CktID <> "999" Then
      genHnd& = genSearch( busHnd )
      genunitHnd& = genunitSearch( busHnd, CktID )     
      exit For
    End If
  Next
  If genHnd = 0 Or genunitHnd = 0 Then 
    printTTY("  Error: Object not found")
    exit Function
  End If
  countUpdated = 0
  listUpdated$ = ""
  For ii = 1 to cols
    nIndex = 0
    sFieldVal$ = FieldVal(ii)
    paramID = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),1)
    If paramID = 0 Then 
      GoTo NextIteration
    End If
    If paramID = GE_nActive Or paramID = GE_dVSourcePU Or paramID = GE_dRefAngle Or _
       paramID = GE_nFixedPQ Or paramID = GE_dCurrLimit1 Or paramID = GE_dCurrLimit2 Then
      If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,genHnd,paramID&,nIndex) Then
        countUpdated = countUpdated + 1
        If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
      End If 
    ElseIf paramID = GU_vdR Or paramID = GU_dRz Then
      select Case FieldName(ii)
        case "Subtransient"
          nIndex = 1
        case "Synchronous" 
          nIndex = 2
        case "Transient"
          nIndex = 3
        case "Neg. Seq."
          nIndex = 4
        case "Zero Seq."
          nIndex = 5     
      End select

      paramIDX = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),2)
      Real$ = c2ri(sFieldVal,1)
      Imag$ = c2ri(sFieldVal,2)
      
      If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,Real,genunitHnd,paramID&,nIndex) Or _
         0 < SetFieldValue(sLabels,nCodes,nCountCodes,Imag,genunitHnd,paramIDX&,nIndex) Then
        countUpdated = countUpdated + 1
        If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
      End If

    Else      
      If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,genunitHnd,paramID&,nIndex) Then
        countUpdated = countUpdated + 1
        If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
      End If
    End If 
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(genunitHnd) = 0 Or PostData(genHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString() )
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Gen = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''Load'''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InitParamCode_Load( l() As String, c() As long ) As long
  l(1)  = "In Serv"
  l(2)  = "ID"
  l(3)  = "Const. P" 
  l(4)  = "Const. P"  
  l(5)  = "Const I"
  l(6)  = "Const I"
  l(7)  = "Const. Z"
  l(8)  = "Const. Z"

  c(1)  = LD_nActive
  c(2)  = LU_sID
  c(3)  = LU_vdMW
  c(4)  = LU_vdMVAR
  c(5)  = LU_vdMW
  c(6)  = LU_vdMVAR
  c(7)  = LU_vdMW
  c(8)  = LU_vdMVAR
  
  InitParamCode_Load = 8
End Function
Function checkHeader_Load( ByRef Array() As String, ByVal colCount )
  okBusName  = false
  okCktID = false
  checkHeader_Load = false
  For ii = 1 to colCount
    If Array(ii) = "Bus Name" Then okBusName  = true
    If Array(ii) = "ID"   Then okCktID = true
    If okBusName And okCktID Then
      checkHeader_Load = true
      exit For
    End If
  Next
End Function
Function processRow_Load( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols ) As long
  processRow_Load = 0
  ' Find object handle
  busHnd& = -1
  loadHnd& = -1
  CktID$   = "999"
  sFieldVal = ""
  For ii = 1 to cols
    If FieldName(ii) = "Bus Name" Then busHnd = findBusHnd(FieldVal(ii))
    If FieldName(ii) = "ID"   Then CktID   = FieldVal(ii)
    If busHnd> -1 And CktID <> "999" Then  
      loadHnd& = loadunitSearch( busHnd, CktID )  
      exit For
    End If
  Next
  If loadHnd = 0 Then 
    printTTY("  Error: Object not found")
    exit Function
  End If
  countUpdated = 0
  listUpdated$ = ""
  For ii = 1 to cols
    nIndex = 0
    sFieldVal$ = FieldVal(ii)
    paramID = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),1)
    If paramID = 0 Then 
      GoTo NextIteration
    End If
    If paramID = LU_vdMW Then
      select Case FieldName(ii)
        case "Const. P"
          nIndex = 1
        case "Const I"
          nIndex = 2
        case "Const. Z"
          nIndex = 3    
      End select
      paramIDX = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),2)
      Real$ = c2ri(sFieldVal,1)
      Imag$ = c2ri(sFieldVal,2)
      
      If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,Real,loadHnd,paramID&,nIndex) Or _
         0 < SetFieldValue(sLabels,nCodes,nCountCodes,Imag,loadHnd,paramIDX&,nIndex) Then
        countUpdated = countUpdated + 1
        If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
      End If

    Else      
      If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,loadHnd,paramID&,nIndex) Then
        countUpdated = countUpdated + 1
        If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
      End If
    End If 
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(loadHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString() )
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Load = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''Shunt'''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InitParamCode_Sh( l() As String, c() As long ) As long
  l(1)  = "In Serv"
  l(2)  = "ID"
  l(3)  = "G"
  l(4)  = "B"
  l(5)  = "G0"
  l(6)  = "B0"

  c(1)  = SC_nActive
  c(2)  = SC_sID
  c(3)  = SU_dG
  c(4)  = SU_dB
  c(5)  = SU_dG0
  c(6)  = SU_dB0
  
  InitParamCode_Sh = 6
End Function
Function checkHeader_Sh( ByRef Array() As String, ByVal colCount )
  okBusName  = false
  okCktID = false
  checkHeader_Sh = false
  For ii = 1 to colCount
    If Array(ii) = "Bus Name" Then okBusName  = true
    If Array(ii) = "ID"   Then okCktID = true
    If okBusName And okCktID Then
      checkHeader_Sh = true
      exit For
    End If
  Next
End Function
Function processRow_Sh( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols ) As long
  processRow_Sh = 0
  ' Find object handle
  busHnd& = -1
  shHnd& = -1
  shunitHnd& = -1
  CktID$   = "999"
  sFieldVal = ""
  For ii = 1 to cols
    If FieldName(ii) = "Bus Name" Then busHnd = findBusHnd(FieldVal(ii))
    If FieldName(ii) = "ID"   Then CktID   = FieldVal(ii)
    If busHnd> -1 And CktID <> "999" Then  
      shHnd& = shSearch( busHnd )  
      shunitHnd& = shunitSearch( busHnd, CktID )
      exit For
    End If
  Next
  If shHnd = 0 Or shunitHnd = 0 Then 
    printTTY("  Error: Object not found")
    exit Function
  End If
  countUpdated = 0
  listUpdated$ = ""
  For ii = 1 to cols
    nIndex = 0
    sFieldVal$ = FieldVal(ii)
    paramID = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),1)
    If paramID = 0 Or paramID = SC_nActive Or paramID = SC_sID Then 
      GoTo NextIteration
    End If   
    If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,shunitHnd,paramID&,nIndex) Then
      countUpdated = countUpdated + 1
      If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
    End If
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(shunitHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString() )
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Sh = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function
