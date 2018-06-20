' ASPEN Sample Script 
' UpdateMuData.BAS
'
' Update mutual line data using Change file
'
' Usage:
' 1- Remove existing mutual line data
' 2- Read Excel file to generate change file
' 3- Apply the newly generated change file to the OLR file.
'
' Version: 1.0
'

Dim aBusNumber(5000) As Long
Dim aBusHandle(5000) As Long

Sub main

' Remove existing mutual line data
  dVal# = 0.0
  nBrHnd& = 0
  While GetEquipment( TC_LINE, nBrHnd& ) > 0
   nMuHnd& = 0 
   While GetData( nBrHnd&, LN_nMuPairHnd, nMuHnd& ) > 0
    Call GetData( nMuHnd&, MU_dR, RVal# )
    If Abs(RVal) > 0.00000001 Then
     Call SetData( nMuHnd&, MU_dR, dVal# )
     Call PostData(nMuHnd&) 
    End If 
    Call GetData( nMuHnd&, MU_dX, XVal# )
    If Abs(XVal) > 0.00000001 Then
     Call SetData( nMuHnd&, MU_dX, dVal# )
     Call PostData(nMuHnd&) 
    End If 
   Wend
  Wend

' Read excel file and create change file
  ExcelFile$ = InputBox("Enter excel file name")
  
  If Len(ExcelFile) = 0 Then 
    Print "Bye"
    Stop
  End If
  
  nExtension = InStr( ExcelFile, ".xlsx" )
  If nExtension > 0 Then
   ChangeFile$ = Mid(ExcelFile, 1, nExtension-1)
  Else
   ChangeFile$ = ExcelFile
  End If
  LogFile$ = ChangeFile$
  ChangeFile$ = ChangeFile & "_M.CHF"
  LogFile$    = LogFile & "_Log.txt"

  Open ChangeFile For output As 1
  Open LogFile For output As 2  

' Get Pointer to Excel application
  On Error GoTo excelErr  
  Set xlApp = CreateObject("excel.application")
  Set wkbook = xlApp.Workbooks.Open( ExcelFile, True, True) 
  On Error GoTo dataSheetErr
  Set dataSheet = xlApp.Worksheets(2)
  On Error GoTo endProgram 
  
' Create a list with busnumber and bushandle 
  ii = 0
  While GetEquipment(TC_BUS, BusHandle) > 0
   ii = ii + 1
   Call GetData(BusHandle, BUS_nNumber, BusNumber&)
   aBusNumber(ii) = BusNumber
   aBusHandle(ii) = BusHandle 
  Wend
  nBus = ii
  
' Sort the list with ascending busnumbers  
  For ii = 1 to nBus
   For jj = ii+1 to nBus
     If aBusNumber(ii) > aBusNumber(jj) Then
      Temp1 = aBusNumber(ii)
      Temp2 = aBusHandle(ii)
      aBusNumber(ii) = aBusNumber(jj)
      aBusHandle(ii) = aBusHandle(jj)
      aBusNumber(jj) = Temp1
      aBusHandle(jj) = Temp2
     End If
   Next jj
  Next ii
  
  nRow = 1
  Do While dataSheet.Cells(nRow,2).Value <> "Line / Section" 
   nRow = nRow + 1
  Loop
  nRow = nRow + 1

  Print #1, "[ONELINER AND POWER FLOW CHANGE FILE]"
  Print #1, ""
  Print #1, "[ADD MUTUAL]"
  Do While dataSheet.Cells(nRow,2).Value <> ""
   aLine$ = ""
   BranchBus1Num&  = dataSheet.Cells(nRow,3).Value  
   BranchBus2Num&  = dataSheet.Cells(nRow,4).Value
   CheckID$        = dataSheet.Cells(nRow,5).Value
   If Len(CheckID) = 0 Then
     CheckID$ = " "
   End If
   nHndBus1&       = FindBusHndByNum(aBusNumber(),aBusHandle(),BranchBus1Num&,nBus%)
   nHndBus2&       = FindBusHndByNum(aBusNumber(),aBusHandle(),BranchBus2Num&,nBus%)
   If nHndBus1 = "-9999" Then
     bLine$ = "Row # " & Str(nRow) & " Can Not find Bus Number: " & Str(BranchBus1Num)
     Print #2, bLine$ 
     GoTo ContinueDo
   End If
   If nHndBus2 = "-9999" Then
     bLine$ = "Row # " & Str(nRow) & " Can Not find Bus Number: " & Str(BranchBus2Num)
     Print #2, bLine$    
     GoTo ContinueDo
   End If
   Call GetData(nHndBus1, BUS_sName, Bus1Name$)  
   Call GetData(nHndBus2, BUS_sName, Bus2Name$)    
   BranchBus1Name$ = "'" & Bus1Name$ & "'"
   BranchBus2Name$ = "'" & Bus2Name$ & "'"    
   ID$ = branchIDCheck( nHndBus1&, nHndBus2&, CheckID$ )
   If ID$ = "9999*" Then
    BranchID$       = "'" & CheckID$ & "'" 
   ElseIf ID$ = "-9999*" Then 
    BranchID$       = "'" & CheckID$ & "'"
	bLine$ = "Row # " & Str(nRow) & " Can Not find Line: " & BranchBus1Name$ & "-" & BranchBus2Name & " ChkID: " & CheckID$
    Print #2, bLine$ 
    GoTo ContinueDo 
   End If  
   BranchVolt$     = dataSheet.Cells(nRow,6).Value
 
   aLine$ = aLine$ & BranchBus1Name$ & " " &  BranchVolt$ & " " 
   aLine$ = aLine$ & BranchBus2Name$ & " " &  BranchVolt$ & " " 
   aLine$ = aLine$ & BranchID$ & " "
   
   BranchBus1Num&  = dataSheet.Cells(nRow,9).Value  
   BranchBus2Num&  = dataSheet.Cells(nRow,10).Value
   CheckID$        = dataSheet.Cells(nRow,11).Value
   If Len(CheckID) = 0 Then
     CheckID$ = " "
   End If
   nHndBus1&       = FindBusHndByNum(aBusNumber(),aBusHandle(),BranchBus1Num&,nBus%)
   nHndBus2&       = FindBusHndByNum(aBusNumber(),aBusHandle(),BranchBus2Num&,nBus%)
   
   If nHndBus1 = "-9999" Then
     bLine$ = "Row # " & Str(nRow) & " Can Not find Bus Number: " & Str(BranchBus1Num)
     Print #2, bLine$ 
     GoTo ContinueDo
   End If
   If nHndBus2 = "-9999" Then
     bLine$ = "Row # " & Str(nRow) & " Can Not find Bus Number: " & Str(BranchBus2Num)
     Print #2, bLine$    
     GoTo ContinueDo
   End If   
   Call GetData(nHndBus1, BUS_sName, Bus1Name$)  
   Call GetData(nHndBus2, BUS_sName, Bus2Name$)    
   BranchBus1Name$ = "'" & Bus1Name$ & "'"
   BranchBus2Name$ = "'" & Bus2Name$ & "'"
   ID$ = branchIDCheck( nHndBus1&, nHndBus2&, CheckID$ )
   If ID$ = "9999*" Then
    BranchID$       = "'" & CheckID$ & "'"
   ElseIf ID$ = "-9999*" Then 
    BranchID$       = "'" & CheckID$ & "'"
    bLine$ = "Row # " & Str(nRow) & " Can Not find Line: " & BranchBus1Name$ & "-" & BranchBus2Name & " ChkID: " & CheckID$
    Print #2, bLine$ 
    GoTo ContinueDo    
   End If 
   BranchVolt$     = dataSheet.Cells(nRow,12).Value
   Rpu$            = dataSheet.Cells(nRow,16).Value
   Xpu$            = dataSheet.Cells(nRow,17).Value
   
   aLine$ = aLine$ & BranchBus1Name$ & " " &  BranchVolt$ & " " 
   aLine$ = aLine$ & BranchBus2Name$ & " " &  BranchVolt$ & " " 
   aLine$ = aLine$ & BranchID$ & "=" & " "
   aLine$ = aLine$ & Rpu$ & " " & Xpu & " "
   aLine$ = aLine$ & "0" & " " & "100" & " " & "0" & " " & "100"
   
   Print #1, aLine$
   
   continueDo:
   nRow = nRow + 1
  Loop
  
  Close 1
  Close 2
  
  ReadChangeFile(ChangeFile$)
  

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

' Find bus handle with bus number ( binary search )
Function FindBusHndByNum( ByRef BusNumArray() As Long, ByRef BusHdlArray() As Long, BusNum&, nBus% ) As Long
  FindBusHndByNum = -9999
  first = 1
  last = nBus
  middle = Int((first+last)/2)
  While first <= last
   If BusNumArray(middle) < BusNum Then
    first = middle + 1
   ElseIf BusNumArray(middle) = BusNum Then
   	FindBusHndByNum = BusHdlArray(middle)
    exit Function
   Else
    last = middle - 1
   End If 
   middle = Int((first+last)/2) 
  Wend  
  FindBusHndByNumReturn:
End Function

Function  branchIDCheck( nHndBus1&, nHndBus2&, sCktID$ ) As String
  branchIDCheck = "-9999*"
  BranchHnd = 0
  nLoop = 0
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
        branchIDCheck = "9999*"
        exit Function
      Else
        nLoop = nLoop + 1
      End If
    End If
  Wend
  branchIDCheckReturn:
End Function

