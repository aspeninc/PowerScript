' ASPEN PowerScript sample program
'
' DATAWKS.BAS
'
' Version 2.7
'
' Import network data from excel spreadsheets 
'
' Spreadsheet must have data rows in exact same format as in 
' the corresponding OneLiner data browser report table. Specifically
' following rows must be included:
' 1- FILESIGNATURE$
' 2- Table type signature aType_XXX$ 
' 3- Header row that matches the corresponding data browser table
'
'
' Global vars and consts
Const FILESIGNATURE$ = "ASPEN OneLiner/Power Flow"
Const aType_Line$    = "Browser Report for Lines"
Const aType_Mu$      = "Browser Report for Zero-Sequence Mutuals"
Const aType_Xfmr$    = "Browser Report for Transformers: 2-Winding"
Const aType_Xfmr3$   = "Browser Report for Transformers: 3-Winding"
Const aType_Ps$      = "Browser Report for Phase Shifters"
Const aType_Gen$     = "Browser Report for Generators"
Const aType_Load$    = "Browser Report for Loads"
Const aType_Shc$     = "Browser Report for Shunts"
Const aType_Bus$     = "Browser Report for Buses"
Const aType_Breaker$ = "Browser Report for Breakers"
Const aType_Scap$    = "Browser Report for Series Capacitors/Reactors"
Const aType_Ocp$     = "Browser Report for Overcurrent Relay-Phase"
Const aType_Ocg$     = "Browser Report for Overcurrent Relay-Ground"
Const aType_Dsp$     = "Browser Report for Distance Relay-Phase"
Const aType_Dsg$     = "Browser Report for Distance Relay-Ground"

Const code_Memo = -999
Const code_Tags = -998

dim sLabels(50) As String
dim nCodes(50) As long     
dim nCountCodes As long
Dim xlApp As Object     ' Declare variable to hold the reference.
Dim wkbook As Object    ' Declare variable to hold the reference.
Dim dataSheet As Object

Sub main
  Dim outArr(50) As String
  dim fieldNames(50) As String 
  
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
    case aType_Bus
      checkHeader = checkHeader_Bus( Array, colcount )
      If checkHeader Then nCountCodes& = InitParamCode_Bus(sLabels,nCodes)	
    case aType_Breaker
      checkHeader = checkHeader_Breaker( Array, colcount )
      If checkHeader Then nCountCodes& = InitParamCode_Breaker(sLabels,nCodes)	 
    case aType_Scap
      checkHeader = checkHeader_Scap( Array, colcount )
      If checkHeader Then nCountCodes& = InitParamCode_Scap(sLabels,nCodes)	
    case aType_Ocp
      checkHeader = checkHeader_Ocp( Array, colcount )
      If checkHeader Then nCountCodes& = InitParamCode_Ocp(sLabels,nCodes)
    case aType_Ocg
      checkHeader = checkHeader_Ocg( Array, colcount )
      If checkHeader Then nCountCodes& = InitParamCode_Ocg(sLabels,nCodes) 
    case aType_Dsp
      checkHeader = checkHeader_Dsp( Array, colcount )
      If checkHeader Then nCountCodes& = InitParamCode_Dsp(sLabels,nCodes)
    case aType_Dsg
      checkHeader = checkHeader_Dsg( Array, colcount )
      If checkHeader Then nCountCodes& = InitParamCode_Dsg(sLabels,nCodes) 	 
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
      If nLen > 10 Then _
        aText$ = aText & Trim(Left(FieldVal(ii), nLen-7)) & " " & Right(FieldVal(ii),7) _
      Else _
        aText$ = aText & FieldVal(ii)
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
    case aType_Bus
      processRow = processRow_Bus( FieldName, FieldVal, cols )  
    case aType_Breaker
      processRow = processRow_Breaker( FieldName, FieldVal, cols ) 
    case aType_Scap 
      processRow = processRow_Scap( FieldName, FieldVal, cols ) 
    case aType_Ocp 
      processRow = processRow_Ocp( FieldName, FieldVal, cols )
    case aType_Ocg 
      processRow = processRow_Ocg( FieldName, FieldVal, cols )
    case aType_Dsp 
      processRow = processRow_Dsp( FieldName, FieldVal, cols )
    case aType_Dsg 
      processRow = processRow_Dsg( FieldName, FieldVal, cols )    
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

 ' Find branch equipment handle 
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
        CktIDTmp$ = "0" + CktID$
        If myID = CktID Or myID = CktIDTmp Or (Len(myID) = 0 And Len(CktID) = 0) Then
          branchSearch = thisItemHnd
          exit Do
        End If
      End If
    End If
  Wend
End Function

 ' Find branch handle 
Function branchHndSearch( nType&, bus1Hnd&, bus2Hnd&, bus3Hnd&, CktID$ )
  branchHndSearch = 0
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
        CktIDTmp$ = "0" + CktID$
        If myID = CktID Or myID = CktIDTmp Or (Len(myID) = 0 And Len(CktID) = 0) Then
          branchHndSearch = branchHnd
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
      exit Do
    End If
  Wend
End Function

 ' Find load handle
Function loadSearch( busHnd& )
  loadSearch = 0
  loadHnd&   = 0
  If GetBusEquipment( busHnd, TC_LOAD, loadHnd& ) > 0  Then
    loadSearch = loadHnd&
  End If
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
      exit Do
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

 ' Find breaker handle
Function breakerSearch( busHnd&, breakerName$ )
  breakerSearch = 0
  breakerHnd&   = 0
  While GetBusEquipment( busHnd, TC_BREAKER, breakerHnd& ) > 0
    Call GetData(breakerHnd, Bk_sID, myID$)
    myID = Trim(myID)
    If myID = breakerName Then
      breakerSearch = breakerHnd
      exit Do
    End If
  Wend
End Function

' Find series capacitor/reactor handle
Function scapSearch( bus1Hnd&, bus2Hnd&, CktID$ ) 
  scapSearch = 0
  scapHnd&   = 0
  While GetEquipment( TC_SCAP, scapHnd& ) > 0 
    Call GetData(scapHnd, SC_nBus1Hnd, Hnd1&)
    Call GetData(scapHnd, SC_nBus2Hnd, Hnd2&)
    Call GetData(scapHnd, SC_sID, myID$)
    If bus1Hnd = Hnd1 And bus2Hnd = Hnd2 And myID = CktID Then
      scapSearch = scapHnd
      exit Do
    End If 
  Wend
End Function

' Find overcurrent relay-phase handle
Function ocpSearch( brHnd&, ID$ )
  ocpSearch = 0
  ocpHnd&   = 0 
  RlyGrpHnd1& = 0
  Call GetData(brHnd, BR_nRlyGrp1Hnd, RlyGrpHnd1&)
  While GetEquipment( TC_RLYOCP, ocpHnd& ) > 0
    Call GetData(ocpHnd, OP_sID, myID$)
    Call GetData(ocpHnd, OP_nRlyGrHnd, RlyGrpHnd2)
    If myID = ID And RlyGrpHnd1 = RlyGrpHnd2 Then
      ocpSearch = ocpHnd
      exit Do
    End If  
  Wend
End Function

' Find overcurrent relay-ground handle
Function ocgSearch( brHnd&, ID$ )
  ocgSearch = 0
  ocgHnd&   = 0
  RlyGrHnd& = 0
  BranchHnd&= 0
  While GetEquipment( TC_RLYOCG, ocgHnd& ) > 0
    Call GetData(ocgHnd, OG_sID, myID$)
    Call GetData(ocgHnd, OG_nRlyGrHnd, RlyGrHnd)
    Call GetData(RlyGrHnd, RG_nBranchHnd, BranchHnd)
    If myID = ID And brHnd = BranchHnd Then
      ocgSearch = ocgHnd
      exit Do
    End If  
  Wend
End Function

' Find distance relay-phase handle
Function dspSearch( brHnd&, ID$, ID2$ )
  dspSearch = 0
  dspHnd&   = 0
  RlyGrHnd& = 0
  BranchHnd&= 0
  While GetEquipment( TC_RLYDSP, dspHnd& ) > 0
    Call GetData(dspHnd, DP_sID, myID$)
    Call GetData(dspHnd, DP_sType, myID2$)
    Call GetData(dspHnd, DP_nRlyGrHnd, RlyGrHnd)
    Call GetData(RlyGrHnd, RG_nBranchHnd, BranchHnd)
    If myID = ID And myID2 = ID2 And brHnd = BranchHnd Then
      dspSearch = dspHnd
      exit Do
    End If  
  Wend
End Function

' Find distance relay-ground handle
Function dsgSearch( brHnd&, ID$, ID2$ )
  dsgSearch = 0
  dsgHnd&   = 0
  RlyGrHnd& = 0
  BranchHnd&= 0
  While GetEquipment( TC_RLYDSG, dsgHnd& ) > 0
    Call GetData(dsgHnd, DG_sID, myID$)
    Call GetData(dsgHnd, DG_sType, myID2$)
    Call GetData(dsgHnd, DG_nRlyGrHnd, RlyGrHnd)
    Call GetData(RlyGrHnd, RG_nBranchHnd, BranchHnd)
    If myID = ID And myID2 = ID2 And brHnd = BranchHnd Then
      dsgSearch = dsgHnd
      exit Do
    End If  
  Wend
End Function

 ' Set field value 
Function SetFieldValue( sLabels() As String, nCodes() As long, nCountCodes&, sFieldVal$, thisHnd&, paramID&, nIndex) As long
  nChanged$ = 0
  SetFieldValue = 0
  paramType = 0
  nPos = 0
  Dim vdArray(5) As Double
  If ParamID = code_Memo Then
   sVal$ = sFieldVal
   sValtemp = GetObjMemo( thisHnd )
   nPos = InStr( sValtemp$, Chr(13)&Chr(10) )
   While nPos > 0 
     LWord = Left(sValtemp, nPos - 1)  ' Get left word.
     RWord = Right(sValtemp, Len(sValtemp) - nPos - 1)  ' Get right word.
     sValtemp = LWord & "  " & RWord
     nPos = InStr( sValtemp$, Chr(13)&Chr(10) ) 
   Wend 
   If sVal <> sValtemp Then 
    SetFieldValue = SetObjMemo( thisHnd&, sVal$ )
    nChanged = 1
    paramType& = 0
   End If
   exit Function
  End If
  If ParamID = code_Tags Then
   sVal$ = sFieldVal
   sValtemp = GetObjTags( thisHnd )
   If sVal <> sValtemp Then 
    SetFieldValue = SetObjTags( thisHnd&, sVal$ )
    nChanged = 1
    paramType& = 0
   End If
   exit Function
  End If
  If paramID > 1000 Then   ' V12 or earlier
    paramType& = paramID/1000
    paramType& = paramID - paramType*1000
    paramType& = paramType/100
  Else
    paramType& = paramID
    paramType& = paramType/100
  End If
  select case paramType&
    case 1 ' String
      sVal$ = sFieldVal
      Call GetData( thisHnd, paramID, sValtemp$ )
      If sVal <> sValtemp Then 
        SetFieldValue = SetData( thisHnd, paramID, sVal$ )
        nChanged = 1
      End If
    case 2 ' double     
      dVal# = a2d(sFieldVal)
      Call GetData( thisHnd, paramID, dValtemp# )
      If dValtemp <> 0 Then
        If Abs(dVal - dValtemp)/Abs(dValtemp) > 0.00001 Then 
          SetFieldValue = SetData( thisHnd, paramID, dVal# )     
          nChanged = 1
        End If
      Else
        If Abs(dVal - dValtemp) > 0.00001 Then 
          SetFieldValue = SetData( thisHnd, paramID, dVal# )     
          nChanged = 1
        End If 
      End If
    case 3 ' Integer
      If UCase(sFieldVal) = "YES" Then 
        nVal& = 1
      ElseIf UCase(sFieldVal) = "NO" Then
        If paramID = BK_nDontDerate Then
          nVal& = 0
        Else
          nVal& = 2
        End If 
      ElseIf UCase(sFieldVal) = "PV" Then
        nVal& = 0
      Elseif UCase(sFieldVal) = "PQ" Then
        nVal& = 1    
      Elseif UCase(sFieldVal) = "WYE" Then
        nVal& = 0
      Elseif UCase(sFieldVal) = "DELTA" Then
        nVal& = 1
      Elseif UCase(sFieldVal) = "FUSE INSIDE DELTA" Then
        nVal& = 3
      Else
        nVal& = Val(sFieldVal)
      End If 
      Call GetData( thisHnd, paramID, nValtemp& )
      If nVal <> nValtemp Then 
       SetFieldValue = SetData( thisHnd, paramID, nVal& )
       nChanged = 1
      end if
    case 5 ' Array
      Call GetData( thisHnd, paramID, vdArray() )
      If Abs(vdArray(nIndex) - a2d(sFieldVal)) > 0.000001 Then
        vdArray(nIndex) = a2d(sFieldVal)
        SetFieldValue = SetData( thisHnd, paramID, vdArray() ) 
        nChanged = 1
      End If
  End select
    If nChanged And SetFieldValue >= 0 Then SetFieldValue = PostData(thisHnd)       
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
  l(21) = "Memo"
  l(22) = "Tags"
  
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
  c(21) = code_Memo
  c(22) = code_Tags
  
  InitParamCode_Line = 22
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
  l(7)  = "Memo"
  l(8)  = "Tags"
  
  c(1)  = MU_dR
  c(2)  = MU_dX
  c(3)  = MU_dFrom1
  c(4)  = MU_dTo1
  c(5)  = MU_dFrom2
  c(6)  = MU_dTo2
  c(7)  = code_Memo
  c(8)  = code_Tags
  
  InitParamCode_Mu = 8
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
  l(2)  = "Name"
  l(3)  = "Tap1 V(pu)"
  l(4)  = "Tap2 V(pu)" 
  l(5)  = "MVA1"
  l(6)  = "MVA2"
  l(7)  = "MVA3" 
  l(8)  = "R"
  l(9)  = "X"
  l(10) = "B"
  l(11) = "R0"
  l(12) = "X0"  
  l(13) = "B0"
  l(14) = "G1" 
  l(15) = "B1"
  l(16) = "G10"
  l(17) = "B10"
  l(18) = "G2"
  l(19) = "B2"
  l(20) = "G20"
  l(21) = "B20"
  l(22) = "ZG1"
  l(23) = "ZG1"
  l(24) = "ZG2"
  l(25) = "ZG2"
  l(26) = "ZGN"
  l(27) = "ZGN"
  l(28) = "Memo"
  l(29) = "Tags"

  c(1)  = XR_nInService
  c(2)  = XR_sName
  c(3)  = XR_dTap1
  c(4)  = XR_dTap2
  c(5)  = XR_dMVA1
  c(6)  = XR_dMVA2
  c(7)  = XR_dMVA3  
  c(8)  = XR_dR
  c(9)  = XR_dX
  c(10) = XR_dB
  c(11) = XR_dR0
  c(12) = XR_dX0  
  c(13) = XR_dB0
  c(14) = XR_dG1  
  c(15) = XR_dB1
  c(16) = XR_dG10
  c(17) = XR_dB10
  c(18) = XR_dG2
  c(19) = XR_dB2
  c(20) = XR_dG20  
  c(21) = XR_dB20
  c(22) = XR_dRG1
  c(23) = XR_dXG1
  c(24) = XR_dRG2
  c(25) = XR_dXG2
  c(26) = XR_dRGN
  c(27) = XR_dXGN
  c(28) = code_Memo
  c(29) = code_Tags

  InitParamCode_Xfmr = 29
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
  l(2)  = "Name"
  l(3)  = "Base MVA"
  l(4)  = "MVA A"
  l(5)  = "MVA B"
  l(6)  = "MVA C"
  l(7)  = "Tap1 V(pu)"
  l(8)  = "Tap2 V(pu)"
  l(9)  = "Tap3 V(pu)"
  l(10) = "Zps"
  l(11) = "Zps"
  l(12) = "Zpt"
  l(13) = "Zpt"
  l(14) = "Zst"
  l(15) = "Zst"
  l(16) = "Zps0"
  l(17) = "Zps0"
  l(18) = "Zpt0"
  l(19) = "Zpt0"
  l(20) = "Zst0"  
  l(21) = "Zst0" 
  l(22) = "Zg1"
  l(23) = "Zg1"
  l(24) = "Zg2"
  l(25) = "Zg2"
  l(26) = "Zg3"
  l(27) = "Zg3"
  l(28) = "Zgn"
  l(29) = "Zgn"
  l(30) = "B"
  l(31) = "B0"
  l(32) = "Memo"
  l(33) = "Tags"
  
  c(1)  = X3_nInService
  c(2)  = X3_sName
  c(3)  = X3_dBaseMVA
  c(4)  = X3_dMVA1
  c(5)  = X3_dMVA2
  c(6)  = X3_dMVA3
  c(7)  = X3_dTap1
  c(8)  = X3_dTap2
  c(9)  = X3_dTap3
  c(10) = X3_dRps
  c(11) = X3_dXps
  c(12) = X3_dRpt
  c(13) = X3_dXpt
  c(14) = X3_dRst
  c(15) = X3_dXst  
  c(16) = X3_dR0ps
  c(17) = X3_dX0ps
  c(18) = X3_dR0pt
  c(19) = X3_dX0pt
  c(20) = X3_dR0st
  c(21) = X3_dX0st
  c(22) = X3_dRG1
  c(23) = X3_dXG1
  c(24) = X3_dRG2
  c(25) = X3_dXG2
  c(26) = X3_dRG3
  c(27) = X3_dXG3
  c(28) = X3_dRGN
  c(29) = X3_dXGN
  c(30) = X3_dB
  c(31) = X3_dB0
  c(32) = code_Memo
  c(33) = code_Tags

'  c(27) = X3_dLTCCenterTap
'  c(28) = X3_dLTCstep
'  c(29) = X3_dMaxTap
'  c(30) = X3_dMaxVW
'  c(31) = X3_nLTCPriority
'  c(32) = X3_nLTCGanged
  

  InitParamCode_Xfmr3 = 33
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
  l(17) = "Memo"
  l(18) = "Tags"
  
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
  c(17) = code_Memo
  c(28) = code_Tags
  
  InitParamCode_Ps = 18
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
  l(2)  = "MVA" 
  l(3)  = "Ref. V"
  l(4)  = "Ref.Ang."
  l(5)  = "Regulation"
  l(6)  = "Limit A"
  l(7)  = "Limit B"
  l(8)  = "Subtransient"
  l(9)  = "Subtransient"
  l(10)  = "Transient"
  l(11) = "Transient"
  l(12) = "Synchronous"
  l(13) = "Synchronous"
  l(14) = "Neg. Seq."
  l(15) = "Neg. Seq."
  l(16) = "Zero Seq."
  l(17) = "Zero Seq."
  l(18) = "Neutral Imp."
  l(19) = "Neutral Imp."
  l(20) = "P Min"
  l(21) = "P Max"
  l(22) = "Q Min"
  l(23) = "Q Max"
  l(24) = "Memo"
  l(25) = "Tags"

  c(1)  = GU_nOnline
  c(2)  = GU_dMVArating
  c(3)  = GE_dVSourcePU
  c(4)  = GE_dRefAngle
  c(5)  = GE_nFixedPQ
  c(6)  = GE_dCurrLimit1
  c(7)  = GE_dCurrLimit2
  c(8)  = GU_vdR
  c(9)  = GU_vdX
  c(10)  = GU_vdR
  c(11) = GU_vdX
  c(12) = GU_vdR
  c(13) = GU_vdX
  c(14) = GU_vdR
  c(15) = GU_vdX
  c(16) = GU_vdR
  c(17) = GU_vdX
  c(18) = GU_dRz
  c(19) = GU_dXz
  c(20) = GU_dPmin
  c(21) = GU_dPmax
  c(22) = GU_dQmin
  c(23) = GU_dQmax
  c(24) = code_Memo
  c(25) = code_Tags
  
  InitParamCode_Gen = 25
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
    If paramID = GE_dVSourcePU Or paramID = GE_dRefAngle Or _
       paramID = GE_nFixedPQ Or paramID = GE_dCurrLimit1 Or _ 
       paramID = GE_dCurrLimit2 Or paramID = code_Memo Then
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
  l(9)  = "Memo"
  l(10) = "Tags"

  c(1)  = LD_nActive
  c(2)  = LU_sID
  c(3)  = LU_vdMW
  c(4)  = LU_vdMVAR
  c(5)  = LU_vdMW
  c(6)  = LU_vdMVAR
  c(7)  = LU_vdMW
  c(8)  = LU_vdMVAR
  c(9)  = code_Memo
  c(10) = code_Tags
  
  InitParamCode_Load = 10
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
      loadHnd& = loadSearch( busHnd ) 
      loadunitHnd& = loadunitSearch( busHnd, CktID )  
      exit For
    End If
  Next
  If loadHnd = 0 Or loadunitHnd = 0 Then 
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
    If paramID = code_Memo Then
      If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,loadHnd,paramID&,nIndex) Then
        countUpdated = countUpdated + 1
        If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
      End If 
    ElseIf paramID = LU_vdMW Then
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
      
      If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,Real,loadunitHnd,paramID&,nIndex) Or _
         0 < SetFieldValue(sLabels,nCodes,nCountCodes,Imag,loadunitHnd,paramIDX&,nIndex) Then
        countUpdated = countUpdated + 1
        If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
      End If

    Else      
      If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,loadunitHnd,paramID&,nIndex) Then
        countUpdated = countUpdated + 1
        If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
      End If
    End If 
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(loadHnd) = 0 Or PostData(loadunitHnd) = 0 Then
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
  l(7)  = "Memo"
  l(8)  = "Tags"
  
  c(1)  = SC_nActive
  c(2)  = SC_sID
  c(3)  = SU_dG
  c(4)  = SU_dB
  c(5)  = SU_dG0
  c(6)  = SU_dB0
  c(7)  = code_Memo
  c(8)  = code_Tags
  
  InitParamCode_Sh = 8
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
    If paramID = code_Memo Then
      If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,shHnd,paramID&,nIndex) Then
        countUpdated = countUpdated + 1
        If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
      End If 
      GoTo NextIteration
    End If  
    If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,shunitHnd,paramID&,nIndex) Then
      countUpdated = countUpdated + 1
      If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
    End If
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(shHnd) = 0 Or PostData(shunitHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString() )
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Sh = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''Bus'''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InitParamCode_Bus( l() As String, c() As long ) As long
  l(1)  = "Bus Name"
  l(2)  = "Location"
  l(3)  = "No."
  l(4)  = "Area"
  l(5)  = "Zone"
  l(7)  = "Tap Bus"
  l(7)  = "Sub. Gp."
  l(8)  = "X"
  l(9)  = "Y"
  l(10) = "Memo"
  l(11) = "Tags"

  c(1)  = BUS_sName
  c(2)  = BUS_sLocation
  c(3)  = BUS_nNumber
  c(4)  = BUS_nArea
  c(5)  = BUS_nZone
  c(6)  = BUS_nTapBus
  c(7)  = BUS_nSubGroup
  c(8)  = BUS_dSPCx
  c(9)  = BUS_dSPCy
  c(10) = code_Memo
  c(11) = code_Tags

  InitParamCode_Bus = 11
End Function

Function checkHeader_Bus( ByRef Array() As String, ByVal colCount )
  okBusName  = false
  okKV       = false
  checkHeader_Bus = false
  For ii = 1 to colCount
    If Array(ii) = "Bus Name" Then okBusName  = true
    If Array(ii) = "kV"   Then okKV = true
    If okBusName And okKV Then
      checkHeader_Bus = true
      exit For
    End If
  Next
End Function
Function processRow_Bus( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols ) As long
  processRow_Bus = 0
  ' Find object handle
  busHnd& = -1
  sBusName$ = ""
  dBuskV#   = 0.0
  sFieldVal = ""
  For ii = 1 to cols
    If FieldName(ii) = "Bus Name" Then sBusName$ = FieldVal(ii)
    If FieldName(ii) = "kV"   Then dBuskV = Val(FieldVal(ii))
    If sBusName <> "" And dBuskV <> 0.0 Then  
      If findBusByName( sBusName, dBuskV, busHnd& ) = 0 Then 
        printTTY("  Error: Object not found")
        exit Function
      End If
      exit For
    End If
  Next
  countUpdated = 0
  listUpdated$ = ""
  For ii = 1 to cols
    nIndex = 0
    sFieldVal$ = FieldVal(ii)
    paramID = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),1)
    If paramID = 0 Then 
      GoTo NextIteration
    elseIf paramID = BUS_nTapBus Then 
      If UCase(sFieldVal$) = "T" Or UCase(sFieldVal) = "3T" Then sFieldVal$ = "1" Else sFieldVal$ = "0"
    End If   
    If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,busHnd,paramID&,nIndex) Then
      countUpdated = countUpdated + 1
      If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
    End If
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(busHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString() )
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Bus = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''Breaker''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InitParamCode_Breaker( l() As String, c() As long ) As long
  l(1)  = "In Serv"
  l(2)  = "Breaker name"
  l(3)  = "Rating"
  l(4)  = "Rating Type"
  l(5)  = "Int. time"
  l(6)  = "Operating kV"
  l(7)  = "CPT1"
  l(8)  = "CPT2"
  l(9)  = "kV range Factor"
  l(10)  = "Max Design kV"
  l(11) = "Rated Momentary Amps"
  l(12) = "No Derate"
  l(13) = "NACD"
  l(14) = "Memo"
  l(15) = "Tags"

  c(1)  = BK_nInService
  c(2)  = BK_sID
  c(3)  = BK_dRating1
  c(4)  = BK_nRatingType
  c(5)  = BK_dCycles
  c(6)  = BK_dOperatingKV
  c(7)  = BK_dCPT1
  c(8)  = BK_dCPT2
  c(9)  = BK_dK
  c(10) = BK_dRatedKV
  c(11) = BK_dRating2
  c(12) = BK_nDontDerate
  c(13) = BK_dNACD  
  c(14) = code_Memo
  c(15) = code_Tags

  InitParamCode_Breaker = 15
End Function
Function checkHeader_Breaker( ByRef Array() As String, ByVal colCount )
  okBusName     = false
  okBreakerName = false
  checkHeader_Breaker = false
  For ii = 1 to colCount
    If Array(ii) = "Bus name"	Then okBusName  = true
    If Array(ii) = "Breaker name"	Then okBreakerName = true
    If okBusName And okBreakerName Then
      checkHeader_Breaker = true
      exit For
    End If
  Next
End Function
Function processRow_Breaker( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols ) As long
  processRow_Breaker = 0
  ' Find object handle
  busHnd& = -1
  breakerHnd& = -1
  breakerName$ = ""
  sFieldVal = ""
  For ii = 1 to cols
    If FieldName(ii) = "Bus name"	Then busHnd = findBusHnd(FieldVal(ii))
    If FieldName(ii) = "Breaker name"   Then breakerName = FieldVal(ii)
    If busHnd> -1 And breakerName <> "" Then   
      breakerHnd& = breakerSearch( busHnd, breakerName )
      exit For
    End If
  Next
  If breakerHnd = 0 Then 
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
    If paramID = Bk_dRating1 Or ParamID = Bk_dCycles Then
      nPos1 = InStr(1,FieldVal(ii)," ")
      sFieldVal$ = Trim(Left(FieldVal(ii),nPos1))
    End If
    If paramID = BK_nRatingType Then
      If sFieldVal = "IEEE: Symm. current" Then sFieldVal$ = "0"
      If sFieldVal = "IEEE: Total current" Then sFieldVal$ = "1"
      If sFieldVal = "IEC" Then sFieldVal$ = "2"
    End If
    If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,breakerHnd,paramID&,nIndex) Then
      countUpdated = countUpdated + 1
      If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
    End If
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(breakerHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString() )
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Breaker = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''Series Capacitors/Reactors''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InitParamCode_Scap( l() As String, c() As long ) As long
  l(1)  = "In Serv"
  l(2)  = "Name"
  l(3)  = "R"
  l(4)  = "X"
  l(5)  = "Ipr"
  l(6)  = "R0"
  l(7)  = "X0"
  l(8)  = "Memo"
  l(9)  = "Tags"
  
  c(1)  = SC_nInService
  c(2)  = SC_sName
  c(3)  = SC_dR
  c(4)  = SC_dX
  c(5)  = SC_dIpr
  c(6)  = SC_dR0
  c(7)  = SC_dX0
  c(8)  = code_Memo
  c(9)  = code_Tags

  InitParamCode_Scap = 9
End Function
Function checkHeader_Scap( ByRef Array() As String, ByVal colCount )
  okBus1  = false
  okBus2  = false
  okCktID = false
  checkHeader_Scap = false
  For ii = 1 to colCount
    If Array(ii) = "Bus 1" Then okBus1  = true
    If Array(ii) = "Bus 2" Then okBus2  = true
    If Array(ii) = "CktID"   Then okCktID = true
    If okBus1 And okBus2 And okCktID Then
      checkHeader_Scap = true
      exit For
    End If
  Next
End Function
Function processRow_Scap( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols ) As long
  processRow_Scap = 0
  bus1Hnd& = -1
  bus2Hnd& = -1
  CktID$   = "999"
  scapHnd& = 0
  For ii = 1 to cols
    If FieldName(ii) = "Bus 1" Then bus1Hnd = findBusHnd(FieldVal(ii))
    If FieldName(ii) = "Bus 2" Then bus2Hnd = findBusHnd(FieldVal(ii))
    If FieldName(ii) = "CktID"   Then CktID   = FieldVal(ii)
    If bus1Hnd> -1 And bus2Hnd > -1 And CktID <> "999" Then
      scapHnd& = scapSearch( bus1Hnd, bus2Hnd, CktID )
      exit For
    End If
  Next
  If scapHnd = 0 Then 
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
    If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,scapHnd,paramID&,nIndex) Then
      countUpdated = countUpdated + 1
      If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
    End If
    If (FieldName(ii) = "R") Or (FieldName(ii) = "X") Then
	  paramID = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii)&"0",1)
	  If paramID = 0 Or sFieldVal = "NA"Then 
	    GoTo NextIteration
	  End If
	  If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,scapHnd,paramID&,nIndex) Then
	    countUpdated = countUpdated + 1
	    If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii ) & "0" Else listUpdated = FieldName(ii) & "0"
	  End If  
    End If  
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(scapHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString())
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Scap = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''Overcurrent Relay-Phase'''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InitParamCode_Ocp( l() As String, c() As long ) As long
  l(1)  = "In Serv"
  l(2)  = "ID"
  l(3)  = "Asset ID"
  'l(4) = 'Curve'
  'l(5) = 'Library'
  l(4)  = "TD"
  l(5)  = "Pickup (A)"
  l(6)  = "CT Ratio"
  l(7)  = "CT Conn."
  l(8)  = "Dir."
  l(9)  = "Inst. (A)"
  l(10) = "DT Delay"
  l(11) = "Inst.Dir."
  'l(12)= "dc-sensitive"
  'l(12) = "Char Angle"
  'l(13) = "Dir. logic"
  l(12) = "Adder1"
  l(13) = "Mult.1"
  l(14) = "Adder2"
  l(15) = "Mult.2"
  'l(16) = "Voltage Controlled or Restrained"
  l(16) = "Reset"
  l(17) = "Memo"
  l(18) = "Tags"
    
  c(1)  = OP_nInService
  c(2)  = OP_sID
  c(3)  = OP_sAssetID
  c(4)  = OP_dTDial
  c(5)  = OP_dTap
  c(6)  = OP_dCT
  c(7)  = OP_nByCTConnect
  c(8)  = OP_nDirectional
  c(9)  = OP_dInst
  c(10) = OP_dInstDelay
  c(11) = OP_nIDirectional
  c(12) = OP_dTimeAdd
  c(13) = OP_dTimeMult
  c(14) = OP_dTimeAdd2
  c(15) = OP_dTimeMult2
  c(16) = OP_dResetTime
  c(17) = code_Memo
  c(18) = code_Tags
  
  InitParamCode_Ocp = 18
End Function

Function checkHeader_Ocp( ByRef Array() As String, ByVal colCount )
  okBranch  = false
  okID      = false
  okAssetID = false
  checkHeader_Ocp = false
  For ii = 1 to colCount
    If Array(ii) = "Branch" Then okBranch  = true
    If Array(ii) = "ID" Then okID  = true
    If Array(ii) = "Asset ID" Then okAssetID = true
    If okBranch And okID And okAssetID Then
      checkHeader_Ocp = true
      exit For
    End If
  Next
End Function

Function processRow_Ocp( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols ) As long
  processRow_Ocp = 0
  ' Find object handle
  ocpHnd&         = 0
  branchHnd&      = -1
  branch_bus1Hnd& = -1
  branch_bus2Hnd& = -1
  CktID$ = "999"
  ID$    = "999"
  For ii = 1 to cols 
    If FieldName(ii) = "Branch" Then 
      nPos = InStr(1,FieldVal(ii)," - ") 
      nPos1= InStr(nPos,FieldVal(ii),"kV")
      nLen = Len(FieldVal(ii))  
      branch_bus1Name$ = Trim(Mid(FieldVal(ii),1,nPos-1))                 
      branch_bus2Name$ = Trim(Mid(FieldVal(ii),nPos+3,nPos1-nPos-1))
      branch_bus1Hnd   = findBusHnd(branch_bus1Name)
      branch_bus2Hnd   = findBusHnd(branch_bus2Name) 
      CktID$           = Trim(Mid(FieldVal(ii),nPos1+2,nLen-2-nPos1))
      BranchType$      = Right(FieldVal(ii), 1)
      select case BranchType$
        case "L"
          nType& = TC_LINE
        case "T"
          nType& = TC_XFMR
        case "X"
          nType& = TC_XFMR3
        case "W"
          nType& = TC_PS
        case default
          Print "Error in branch equipment type"
          Stop
        End select
      If branch_bus1Hnd > -1 And branch_bus2Hnd > -1 And CktID <> "999" Then         
         branchHnd& = branchHndSearch( nType&, branch_bus1Hnd, branch_bus2Hnd, 0, CktID )
      End If
    End If
    If FieldName(ii) = "ID" Then ID = FieldVal(ii)  
    If branchHnd > -1 And ID <> "999" Then
      ocpHnd&     = ocpSearch( branchHnd, ID )
      exit For
    End If
  Next
  If ocpHnd = 0 Then 
    printTTY("  Error: Object not found")
    exit Function
  End If
  countUpdated = 0
  listUpdated$ = ""
  For ii = 1 to cols
    nIndex = 0
    sFieldVal$ = FieldVal(ii)
    paramID = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),1)
    If paramID = 0 Or sFieldVal = "N/A" Then
      GoTo NextIteration
    End If
    If FieldName(ii) = "Pickup (A)" Then
      nPos = InStr(1,FieldVal(ii)," ")
      sFieldVal$ = Left(FieldVal(ii), nPos - 1) 
    End If
    If FieldName(ii) = "CT Ratio" Then
      nPos = InStr(1,FieldVal(ii),"/")
      dPri = a2d(Left(FieldVal(ii), nPos - 1))
      dSec = a2d(Right(FieldVal(ii),Len(FieldVal(ii))-nPos))      
      sFieldVal$ = Str(dPri/dSec)
    End If
    
    If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,ocpHnd,paramID&,nIndex) Then
      countUpdated = countUpdated + 1
      If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
    End If
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(ocpHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString())
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Ocp = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''Overcurrent Relay-Ground'''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InitParamCode_Ocg( l() As String, c() As long ) As long
  l(1)  = "In Serv"
  l(2)  = "ID"
  l(3)  = "Asset ID"
  'l(4) = 'Curve'
  'l(5) = 'Library'
  l(4)  = "TD"
  l(5)  = "Pickup (A)"
  l(6)  = "CT Ratio"
  l(7)  = "Dir."
  l(8)  = "Inst. (A)"
  l(9)  = "DT Delay"
  l(10) = "Inst.Dir."
  'l(11)= "dc-sensitive"
  'l(11) = "Char Angle"
  l(11) = "Dir. logic"
  'l(13) = "CT Location"
  'l(14) = "Operating Qty"
  l(12) = "Adder1"
  l(13) = "Mult.1"
  l(14) = "Adder2"
  l(15) = "Mult.2"
  l(16) = "Reset"
  l(17) = "Memo"
  l(18) = "Tags"
  
  c(1)  = OG_nInService
  c(2)  = OG_sID
  c(3)  = OG_sAssetID
  c(4)  = OG_dTDial
  c(5)  = OG_dTap
  c(6)  = OG_dCT
  c(7)  = OG_nDirectional
  c(8)  = OG_dInst
  c(9)  = OG_dInstDelay
  c(10) = OG_nIDirectional
  c(11) = OG_nPolar
  c(12) = OG_dTimeAdd
  c(13) = OG_dTimeMult
  c(14) = OG_dTimeAdd2
  c(15) = OG_dTimeMult2
  c(16) = OG_dResetTime
  c(17) = code_Memo
  c(18) = code_Tags
  
  InitParamCode_Ocg = 18
End Function

Function checkHeader_Ocg( ByRef Array() As String, ByVal colCount )
  okBranch  = false
  okID      = false
  okAssetID = false
  checkHeader_Ocg = false
  For ii = 1 to colCount
    If Array(ii) = "Branch" Then okBranch  = true
    If Array(ii) = "ID" Then okID  = true
    If Array(ii) = "Asset ID" Then okAssetID = true
    If okBranch And okID And okAssetID Then
      checkHeader_Ocg = true
      exit For
    End If
  Next
End Function

Function processRow_Ocg( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols ) As long
  processRow_Ocg = 0
  ' Find object handle
  ocgHnd&         = 0
  branchHnd&      = -1
  branch_bus1Hnd& = -1
  branch_bus2Hnd& = -1
  CktID$ = "999"
  ID$    = "999"
  For ii = 1 to cols 
    If FieldName(ii) = "Branch" Then 
      nPos = InStr(1,FieldVal(ii)," - ") 
      nPos1= InStr(nPos,FieldVal(ii),"kV")
      nLen = Len(FieldVal(ii))  
      branch_bus1Name$ = Trim(Mid(FieldVal(ii),1,nPos-1))                 
      branch_bus2Name$ = Trim(Mid(FieldVal(ii),nPos+3,nPos1-nPos-1))
      branch_bus1Hnd   = findBusHnd(branch_bus1Name)
      branch_bus2Hnd   = findBusHnd(branch_bus2Name) 
      CktID$           = Trim(Mid(FieldVal(ii),nPos1+2,nLen-2-nPos1))
      BranchType$      = Right(FieldVal(ii), 1)
      select case BranchType$
        case "L"
          nType& = TC_LINE
        case "T"
          nType& = TC_XFMR
        case "X"
          nType& = TC_XFMR3
        case "W"
          nType& = TC_PS
        case default
          Print "Error in branch equipment type"
          Stop
        End select
      If branch_bus1Hnd > -1 And branch_bus2Hnd > -1 And CktID <> "999" Then         
         branchHnd& = branchHndSearch( nType&, branch_bus1Hnd, branch_bus2Hnd, 0, CktID )
      End If
    End If
    If FieldName(ii) = "ID" Then ID = FieldVal(ii)  
    If branchHnd > -1 And ID <> "999" Then
      ocgHnd&     = ocgSearch( branchHnd, ID )
      exit For
    End If
  Next
  If ocgHnd = 0 Then 
    printTTY("  Error: Object not found")
    exit Function
  End If
  countUpdated = 0
  listUpdated$ = ""
  For ii = 1 to cols
    nIndex = 0
    sFieldVal$ = FieldVal(ii)
    paramID = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),1)
    If paramID = 0 Or sFieldVal = "N/A" Then
      GoTo NextIteration
    End If
    
    If FieldName(ii) = "Pickup (A)" Then
      nPos = InStr(1,FieldVal(ii)," ")
      sFieldVal$ = Left(FieldVal(ii), nPos - 1) 
    End If
    If FieldName(ii) = "CT Ratio" Then
      nPos = InStr(1,FieldVal(ii),"/")
      dPri = a2d(Left(FieldVal(ii), nPos - 1))
      dSec = a2d(Right(FieldVal(ii),Len(FieldVal(ii))-nPos))      
      sFieldVal$ = Str(dPri/dSec)
    End If
    If FieldName(ii) = "Dir. logic" Then
      If (InStr(1,FieldVal(ii),"Vo") > 0) Then 
         sFieldVal$ = "0"
      ElseIf (InStr(1,FieldVal(ii),"V2") > 0) Then 
         sFieldVal$ = "1"
      ElseIf (InStr(1,FieldVal(ii),"32Q") > 0) Then 
         sFieldVal$ = "2"
      ElseIf (InStr(1,FieldVal(ii),"32G") > 0) Then 
         sFieldVal$ = "3"
      End If
    End If
    
    If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,ocgHnd,paramID&,nIndex) Then
      countUpdated = countUpdated + 1
      If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
    End If
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(ocgHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString())
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Ocg = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''Distance Relay-Phase'''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InitParamCode_Dsp( l() As String, c() As long ) As long
  l(1)  = "In Serv"
  l(2)  = "ID"
  l(3)  = "ID2"
  l(4)  = "Asset ID"
  l(5)  = "Type"
  l(6)  = "CT Ratio"
  l(7)  = "PT Ratio"
  'l(8)  = "PT Bus"
  'l(8)  = "Min I"
  'l(9)  = "Zone-2 supervision"
  'l(8)  = "Z1 Delay"
  'l(9)  = "Z1 Reach"
  'l(10) = "Z2 Delay"
  'l(11) = "Z2 Reach"
  'l(12) = "Z3 Delay"
  'l(13) = "Z3 Reach"
  l(8)  = "Memo"
  l(9)  = "Tags"
  
  c(1)  = DP_nInService
  c(2)  = DP_sID
  c(3)  = DP_sType
  c(4)  = DP_sAssetID
  c(5)  = DP_sDSType
  c(6)  = DP_dCT
  c(7)  = DP_dVT
  'c(8)  = DP_vdDelay
  'c(9)  = DP_vdReach
  'c(10) = DP_vdDelay
  'c(11) = DP_vdReach
  'c(12) = DP_vdDelay
  'c(13) = DP_vdReach
  c(8)  = code_Memo
  c(9)  = code_Tags
  
  InitParamCode_Dsp = 9
End Function

Function checkHeader_Dsp( ByRef Array() As String, ByVal colCount )
  okBranch  = false
  okID      = false
  okID2     = false
  okAssetID = false
  checkHeader_Dsp = false
  For ii = 1 to colCount
    If Array(ii) = "Branch" Then okBranch  = true
    If Array(ii) = "ID" Then okID  = true
    If Array(ii) = "ID2" Then okID2  = true
    If Array(ii) = "Asset ID" Then okAssetID = true
    If okBranch And okID And okID2 And okAssetID Then
      checkHeader_Dsp = true
      exit For
    End If
  Next
End Function

Function processRow_Dsp( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols ) As long
  processRow_Dsp = 0
  ' Find object handle
  dspHnd&         = 0
  branchHnd&      = -1
  branch_bus1Hnd& = -1
  branch_bus2Hnd& = -1
  CktID$ = "999"
  ID$    = "999"
  ID2$   = "999"
  For ii = 1 to cols 
    If FieldName(ii) = "Branch" Then 
      nPos = InStr(1,FieldVal(ii)," - ") 
      nPos1= InStr(nPos,FieldVal(ii),"kV")
      nLen = Len(FieldVal(ii))  
      branch_bus1Name$ = Trim(Mid(FieldVal(ii),1,nPos-1))                 
      branch_bus2Name$ = Trim(Mid(FieldVal(ii),nPos+3,nPos1-nPos-1))
      branch_bus1Hnd   = findBusHnd(branch_bus1Name)
      branch_bus2Hnd   = findBusHnd(branch_bus2Name) 
      CktID$           = Trim(Mid(FieldVal(ii),nPos1+2,nLen-2-nPos1))
      BranchType$      = Right(FieldVal(ii), 1)
      select case BranchType$
        case "L"
          nType& = TC_LINE
        case "T"
          nType& = TC_XFMR
        case "X"
          nType& = TC_XFMR3
        case "W"
          nType& = TC_PS
        case default
          Print "Error in branch equipment type"
          Stop
        End select
      If branch_bus1Hnd > -1 And branch_bus2Hnd > -1 And CktID <> "999" Then         
         branchHnd& = branchHndSearch( nType&, branch_bus1Hnd, branch_bus2Hnd, 0, CktID )
      End If
    End If
    If FieldName(ii) = "ID"  Then ID  = FieldVal(ii)  
    If FieldName(ii) = "ID2" Then ID2 = FieldVal(ii) 
    If branchHnd > -1 And ID <> "999" And ID2 <> "999" Then
      dspHnd&     = dspSearch( branchHnd, ID, ID2 )
      exit For
    End If
  Next
  If dspHnd = 0 Then 
    printTTY("  Error: Object not found")
    exit Function
  End If
  countUpdated = 0
  listUpdated$ = ""
  For ii = 1 to cols
    nIndex = 0
    sFieldVal$ = FieldVal(ii)
    paramID = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),1)
    If paramID = 0 Or sFieldVal = "N/A" Then
      GoTo NextIteration
    End If
    
    If (FieldName(ii) = "CT Ratio") Or (FieldName(ii) = "PT Ratio") Then
      nPos = InStr(1,FieldVal(ii),"/")
      dPri = a2d(Left(FieldVal(ii), nPos - 1))
      dSec = a2d(Right(FieldVal(ii),Len(FieldVal(ii))-nPos))      
      sFieldVal$ = Str(dPri/dSec)
    End If
    'If (FieldName(ii) = "Z1 Delay") Or (FieldName(ii) = "Z1 Reach") Then
    '  nIndex = 1
    'ElseIf (FieldName(ii) = "Z2 Delay") Or (FieldName(ii) = "Z2 Reach") Then
    '  nIndex = 2
    'ElseIf (FieldName(ii) = "Z3 Delay") Or (FieldName(ii) = "Z3 Reach") Then
    '  nIndex = 3
    'End If

    If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,dspHnd,paramID&,nIndex) Then
      countUpdated = countUpdated + 1
      If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
    End If
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(dspHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString())
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Dsp = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''Distance Relay-Ground''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InitParamCode_Dsg( l() As String, c() As long ) As long
  l(1)  = "In Serv"
  l(2)  = "ID"
  l(3)  = "ID2"
  l(4)  = "Asset ID"
  l(5)  = "Type"
  l(6)  = "CT Ratio"
  l(7)  = "PT Ratio"
  'l(8)  = "PT Bus"
  'l(8)  = "Min I"
  'l(8)  = "K1"
  'l(9)  = "K2"
  'l(10) = "Zone-2 supervision"
  'l(8)  = "Z1 Delay"
  'l(9) = "Z1 Reach"
  'l(10) = "Z2 Delay"
  'l(11) = "Z2 Reach"
  'l(12) = "Z3 Delay"
  'l(13) = "Z3 Reach"
  l(8) = "Memo"
  l(9) = "Tags"
  
  
  c(1)  = DG_nInService
  c(2)  = DG_sID
  c(3)  = DG_sType
  c(4)  = DG_sAssetID
  c(5)  = DG_sDSType
  c(6)  = DG_dCT
  c(7)  = DG_dVT
  'c(8)  = DG_dKmag 'DG_dKang
  'c(8)  = DG_vdDelay
  'c(9)  = DG_vdReach
  'c(10) = DG_vdDelay
  'c(11) = DG_vdReach
  'c(12) = DG_vdDelay
  'c(13) = DG_vdReach
  c(8)  = code_Memo
  c(9)  = code_Tags
  
  InitParamCode_Dsg = 9
End Function

Function checkHeader_Dsg( ByRef Array() As String, ByVal colCount )
  okBranch  = false
  okID      = false
  okID2     = false
  okAssetID = false
  checkHeader_Dsg = false
  For ii = 1 to colCount
    If Array(ii) = "Branch" Then okBranch  = true
    If Array(ii) = "ID" Then okID  = true
    If Array(ii) = "ID2" Then okID2  = true
    If Array(ii) = "Asset ID" Then okAssetID = true
    If okBranch And okID And okID2 And okAssetID Then
      checkHeader_Dsg = true
      exit For
    End If
  Next
End Function

Function processRow_Dsg( ByRef FieldName() As String, ByRef FieldVal() As String, ByVal cols ) As long
  processRow_Dsg = 0
  ' Find object handle
  dsgHnd&         = 0
  branchHnd&      = -1
  branch_bus1Hnd& = -1
  branch_bus2Hnd& = -1
  CktID$ = "999"
  ID$    = "999"
  ID2$   = "999"
  For ii = 1 to cols 
    If FieldName(ii) = "Branch" Then 
      nPos = InStr(1,FieldVal(ii)," - ") 
      nPos1= InStr(nPos,FieldVal(ii),"kV")
      nLen = Len(FieldVal(ii))  
      branch_bus1Name$ = Trim(Mid(FieldVal(ii),1,nPos-1))                 
      branch_bus2Name$ = Trim(Mid(FieldVal(ii),nPos+3,nPos1-nPos-1))
      branch_bus1Hnd   = findBusHnd(branch_bus1Name)
      branch_bus2Hnd   = findBusHnd(branch_bus2Name) 
      CktID$           = Trim(Mid(FieldVal(ii),nPos1+2,nLen-2-nPos1))
      BranchType$      = Right(FieldVal(ii), 1)
      select case BranchType$
        case "L"
          nType& = TC_LINE
        case "T"
          nType& = TC_XFMR
        case "X"
          nType& = TC_XFMR3
        case "W"
          nType& = TC_PS
        case default
          Print "Error in branch equipment type"
          Stop
        End select
      If branch_bus1Hnd > -1 And branch_bus2Hnd > -1 And CktID <> "999" Then         
         branchHnd& = branchHndSearch( nType&, branch_bus1Hnd, branch_bus2Hnd, 0, CktID )
      End If
    End If
    If FieldName(ii) = "ID"  Then ID  = FieldVal(ii)  
    If FieldName(ii) = "ID2" Then ID2 = FieldVal(ii) 
    If branchHnd > -1 And ID <> "999" And ID2 <> "999" Then
      dsgHnd&     = dsgSearch( branchHnd, ID, ID2 )
      exit For
    End If
  Next
  If dsgHnd = 0 Then 
    printTTY("  Error: Object not found")
    exit Function
  End If
  countUpdated = 0
  listUpdated$ = ""
  For ii = 1 to cols
    nIndex = 0
    sFieldVal$ = FieldVal(ii)
    paramID = LookupParamCode(sLabels,nCodes,nCountCodes,FieldName(ii),1)
    If paramID = 0 Or sFieldVal = "N/A" Then
      GoTo NextIteration
    End If
    
    If (FieldName(ii) = "CT Ratio") Or (FieldName(ii) = "PT Ratio") Then
      nPos = InStr(1,FieldVal(ii),"/")
      dPri = a2d(Left(FieldVal(ii), nPos - 1))
      dSec = a2d(Right(FieldVal(ii),Len(FieldVal(ii))-nPos))      
      sFieldVal$ = Str(dPri/dSec)
    End If
    'If FieldName(ii) = "K1" Then
    '  nPos = InStr(1,FieldVal(ii),"@")
    '  sFieldVal$ = Left(FieldVal(ii), nPos - 1)
    '  sKang$     = Right(FieldVal(ii), Len(FieldVal(ii))-nPos) 
    'End If
    'If (FieldName(ii) = "Z1 Delay") Or (FieldName(ii) = "Z1 Reach") Then
    '  nIndex = 1
    'ElseIf (FieldName(ii) = "Z2 Delay") Or (FieldName(ii) = "Z2 Reach") Then
    '  nIndex = 2
    'ElseIf (FieldName(ii) = "Z3 Delay") Or (FieldName(ii) = "Z3 Reach") Then
    '  nIndex = 3
    'End If

    If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,dsgHnd,paramID&,nIndex) Then
      countUpdated = countUpdated + 1
      If listUpdated <> "" Then listUpdated = listUpdated & "," & FieldName(ii) Else listUpdated = FieldName(ii)
    End If
    'If FieldName(ii) = "K1" Then
    '  sFiledVal$ = sKang
    '  If 0 < SetFieldValue(sLabels,nCodes,nCountCodes,sFieldVal,dsgHnd,DG_dKang,nIndex) Then
	'    countUpdated = countUpdated + 1
	'    If listUpdated <> "" Then listUpdated = listUpdated & ",K1_Ang" Else listUpdated = "K1_Ang"
	'  End If 
    'End If
    NextIteration
  Next
  If countUpdated > 0 Then
    If PostData(dsgHnd) = 0 Then
      PrintTTY("  Error: " + ErrorString())
      exit Function
    End If
    PrintTTY("  Updated: " & listUpdated )
    processRow_Dsg = countUpdated
  Else
    PrintTTY("  Error: Nothing to update" )
  End If
End Function

