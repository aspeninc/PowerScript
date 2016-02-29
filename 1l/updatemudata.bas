' ASPEN Sample Script 
' UpdateMuData.BAS
'
' Update mutual line data using Change file
'
' Usage:
' 1- Remove existing mutual line data
' 2- Read Excel file to generate change file
' 3- Execute change File to apply the newly generated change file to the OLR file.
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
  ChangeFile$ = ChangeFile & "_M.CHF"

  Open ChangeFile For output As 1

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
   BranchBus1Name$ = "'" & FindBusNameByNum(aBusNumber(),aBusHandle(),BranchBus1Num&,nBus%) & "'"
   BranchBus2Name$ = "'" & FindBusNameByNum(aBusNumber(),aBusHandle(),BranchBus2Num&,nBus%) & "'"
   BranchID$       = "'" & dataSheet.Cells(nRow,5).Value & "'"
   BranchVolt$     = dataSheet.Cells(nRow,6).Value
 
   aLine$ = aLine$ & BranchBus1Name$ & " " &  BranchVolt$ & " " 
   aLine$ = aLine$ & BranchBus2Name$ & " " &  BranchVolt$ & " " 
   aLine$ = aLine$ & BranchID$ & " "
   
   BranchBus1Num&  = dataSheet.Cells(nRow,9).Value  
   BranchBus2Num&  = dataSheet.Cells(nRow,10).Value
   BranchBus1Name$ = "'" & FindBusNameByNum(aBusNumber(),aBusHandle(),BranchBus1Num&,nBus%) & "'"
   BranchBus2Name$ = "'" & FindBusNameByNum(aBusNumber(),aBusHandle(),BranchBus2Num&,nBus%) & "'"
   BranchID$       = "'" & dataSheet.Cells(nRow,11).Value & "'"
   BranchVolt$     = dataSheet.Cells(nRow,12).Value
   Rpu$            = dataSheet.Cells(nRow,16).Value
   Xpu$            = dataSheet.Cells(nRow,17).Value
   
   aLine$ = aLine$ & BranchBus1Name$ & " " &  BranchVolt$ & " " 
   aLine$ = aLine$ & BranchBus2Name$ & " " &  BranchVolt$ & " " 
   aLine$ = aLine$ & BranchID$ & "=" & " "
   aLine$ = aLine$ & Rpu$ & " " & Xpu & " "
   aLine$ = aLine$ & "0" & " " & "100" & " " & "0" & " " & "100"
   
   Print #1, aLine$
   nRow = nRow + 1
  Loop
  
  Close 1
  
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

' Find bus name with bus number ( binary search )
Function FindBusNameByNum( ByRef BusNumArray() As Long, ByRef BusHdlArray() As Long, BusNum&, nBus% ) As String
  FindBusNameByNum = ""
  first = 1
  last = nBus
  middle = Int((first+last)/2)
  While first <= last
   If BusNumArray(middle) < BusNum Then
    first = middle + 1
   ElseIf BusNumArray(middle) = BusNum Then
   	BusHandle = BusHdlArray(middle)
    Call GetData(BusHandle, BUS_sName, BusName$) 
    FindBusNameByNum = BusName
    exit Function
   Else
    last = middle - 1
   End If 
   middle = Int((first+last)/2) 
  Wend  
  FindBusNameByNumReturn:
End Function

