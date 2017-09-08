' ASPEN PowerScript sample program
'
' GETVIRelayEx.BAS
'
' Export branch currents in selected faults to excel file 
'
' Version: 2.0
'
Sub main()
 Dim MagArray(12) As Double
 Dim AngArray(12) As Double
 Dim DummyArray(6) As Long 
 Dim ShowFaultFlag As Long
 Dim BranchHnd As Long
 Dim FaultString As String
 Dim FaultIndex(3000) As Integer
 
 ' Get picked object number
 If GetEquipment( TC_PICKED, ObjHnd ) = 0 Or EquipmentType( ObjHnd ) <> TC_RLYGROUP Then 
   Print "No relay group is selected"
   Exit Sub
 End If
 FtCounts = FaultSelector( FaultIndex, "Fault Selector", "Select faults to report" )
 If FtCounts = 0 Then Stop
 
 On Error GoTo loopBack
 loopBack:
 ExcelFilePath$ = GetOLRFilePath()
 ExcelFile = FileSaveDialog( ExcelFilePath$, "CSV files (*.csv)|*.csv||", ".csv", 2+16 )
 If ExcelFile = "" Then Stop
 Open ExcelFile For Output As 1
 
 
 ' Get branch and bus handles
 If GetData( ObjHnd, RG_nBranchHnd, BranchHnd ) = 0 Then GoTo HasError
 If GetData( BranchHnd, BR_nBus1Hnd, Bus1Hnd )  = 0 Then GoTo HasError
 

 ' Print file header
 StrLine$ = PrintObj1LPF( BranchHnd )
 Print #1, StrLine$
 Print #1, ""
 
 StrLine$ = ",Ia(mag),Ia(ang),Ib(mag),Ib(ang),Ic(mag),Ic(ang)"
 StrLine$ = StrLine$ + ",I+(mag),I+(ang),I-(mag),I-(ang),3I0(mag),3I0(ang)"
 StrLine$ = StrLine$ + ",Va(mag),Va(ang),Vb(mag),Vb(ang),Vc(mag),Vc(ang)"
 StrLine$ = StrLine$ + ",V+(mag),V+(ang),V-(mag),V-(ang),V0(mag),V0(ang)"
 StrLine$ = StrLine$ + ",Fault Description"
 Print #1, StrLine$
	
 Call ProgressDialog(1,"Solution report", "0 of " & FtCounts, 0)
 nStop = 0
 For ii = 1 to FtCounts
  If nStop = 2 Then GoTo doStop
  nFltIdx = FaultIndex(ii-1) + 1
  If PickFault( nFltIdx ) > 0 Then
    nStop = ProgressDialog(1,"Solution report", ii &  " of " & FtCounts, 100*ii/FtCounts)
    If GetSCCurrent( BranchHnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
    StrLine$ = Format( MagArray(1), "#0.0") + "," + Format( AngArray(1), "#0.0") + ","
    StrLine$ = StrLine$ + Format( MagArray(2), "#0.0") + "," + Format( AngArray(2), "#0.0") + ","
    StrLine$ = StrLine$ + Format( MagArray(3), "#0.0") + "," + Format( AngArray(3), "#0.0") + ","      
       
    If ii = 1 Then
          IaMax = MagArray(1)
          Iamin = MagArray(1)
          IbMax = MagArray(2)
          Ibmin = MagArray(2)
          IcMax = MagArray(3)
          Icmin = MagArray(3)
    Else 
          If MagArray(1) > IaMax Then IaMax = MagArray(1)
          If MagArray(1) < IaMin Then IaMin = MagArray(1)
          If MagArray(2) > IbMax Then IbMax = MagArray(2)
          If MagArray(2) < IbMin Then IbMin = MagArray(2)
          If MagArray(3) > IcMax Then IcMax = MagArray(3)
          If MagArray(3) < IcMin Then IcMin = MagArray(3)
    End If 
         
    If GetSCCurrent( BranchHnd, MagArray, AngArray, 2 ) = 0 Then GoTo HasError
    StrLine$ = StrLine$ + Format( MagArray(2), "#0.0") + "," + Format( AngArray(2), "#0.0") + ","
    StrLine$ = StrLine$ + Format( MagArray(3), "#0.0") + "," + Format( AngArray(3), "#0.0") + ","
    MagArray(1) = MagArray(1)*3
    StrLine$ = StrLine$ + Format( MagArray(1), "#0.0") + "," + Format( AngArray(1), "#0.0") + ","  
         
    If ii = 1 Then
          I1Max = MagArray(2)
          I1min = MagArray(2)
          I2Max = MagArray(3)
          I2min = MagArray(3)
          I0Max = MagArray(1)
          I0min = MagArray(1)
    Else 
          If MagArray(2) > I1Max Then I1Max = MagArray(2)
          If MagArray(2) < I1Min Then I1Min = MagArray(2)
          If MagArray(3) > I2Max Then I2Max = MagArray(3)
          If MagArray(3) < I2Min Then I2Min = MagArray(3)
          If MagArray(1) > I0Max Then I0Max = MagArray(1)
          If MagArray(1) < I0Min Then I0Min = MagArray(1)
    End If  
       
    If GetSCVoltage( Bus1Hnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
    StrLine$ = StrLine$ + Format( MagArray(1), "#0.0") + "," + Format( AngArray(1), "#0.0") + ","
    StrLine$ = StrLine$ + Format( MagArray(2), "#0.0") + "," + Format( AngArray(2), "#0.0") + ","
    StrLine$ = StrLine$ + Format( MagArray(3), "#0.0") + "," + Format( AngArray(3), "#0.0") + "," 
       
    If ii = 1 Then
          VaMax = MagArray(1)
          Vamin = MagArray(1)
          VbMax = MagArray(2)
          Vbmin = MagArray(2)
          VcMax = MagArray(3)
          Vcmin = MagArray(3)
    Else 
          If MagArray(1) > VaMax Then VaMax = MagArray(1)
          If MagArray(1) < VaMin Then VaMin = MagArray(1)
          If MagArray(2) > VbMax Then VbMax = MagArray(2)
          If MagArray(2) < VbMin Then VbMin = MagArray(2)
          If MagArray(3) > VcMax Then VcMax = MagArray(3)
          If MagArray(3) < VcMin Then VcMin = MagArray(3)
    End If 
         
    If GetSCVoltage( Bus1Hnd, MagArray, AngArray, 2 ) = 0 Then GoTo HasError
    StrLine$ = StrLine$ + Format( MagArray(2), "#0.0") + "," + Format( AngArray(2), "#0.0") + ","
    StrLine$ = StrLine$ + Format( MagArray(3), "#0.0") + "," + Format( AngArray(3), "#0.0") + ","
    StrLine$ = StrLine$ + Format( MagArray(1), "#0.0") + "," + Format( AngArray(1), "#0.0") + "," 
       
    If ii = 1 Then
          V1Max = MagArray(2)
          V1min = MagArray(2)
          V2Max = MagArray(3)
          V2min = MagArray(3)
          V0Max = MagArray(1)
          V0min = MagArray(1)
    Else 
          If MagArray(2) > V1Max Then V1Max = MagArray(2)
          If MagArray(2) < V1Min Then V1Min = MagArray(2)
          If MagArray(3) > V2Max Then V2Max = MagArray(3)
          If MagArray(3) < V2Min Then V2Min = MagArray(3)
          If MagArray(1) > V0Max Then V0Max = MagArray(1)
          If MagArray(1) < V0Min Then V0Min = MagArray(1)
    End If 
                      
    FaultString$ = FaultDescription()   ' Get fault description  
    nPos = InStr( 1,FaultString,Chr(10) )    
    While nPos > 0
          ALine$ = Left$( FaultString$, nPos - 1 )
          BLine$ = Right$( FaultString$, Len(FaultString$) - nPos )
          FaultString$ = ALine$ + "|" + BLine$
          nPos = InStr( 1,FaultString,Chr(10) )
    Wend
    StrLine$ = StrLine$ + FaultString$       
    StrLine$ = nFltIdx & "," & StrLine$
    Print #1, StrLine$        
  end if
 Next	
 
 ' Print min and max    
 Print #1, ""
 StrLine$ = "Max," + Format( IaMax, "#0.0") + ",," + Format( IbMax, "#0.0") + ",," + Format( IcMax, "#0.0")
 StrLine$ = StrLine$ + ",," + Format( I1Max, "#0.0") + ",," + Format( I2Max, "#0.0") + ",," + Format( I0Max, "#0.0")
 StrLine$ = StrLine$ + ",," + Format( VaMax, "#0.0") + ",," + Format( VbMax, "#0.0") + ",," + Format( VcMax, "#0.0")
 StrLine$ = StrLine$ + ",," + Format( V1Max, "#0.0") + ",," + Format( V2Max, "#0.0") + ",," + Format( V0Max, "#0.0")
 Print #1, StrLine$    
 StrLine$ = "Min," + Format( IaMin, "#0.0") + ",," + Format( IbMin, "#0.0") + ",," + Format( IcMin, "#0.0")
 StrLine$ = StrLine$ + ",," + Format( I1Min, "#0.0") + ",," + Format( I2Min, "#0.0") + ",," + Format( I0Min, "#0.0")
 StrLine$ = StrLine$ + ",," + Format( VaMin, "#0.0") + ",," + Format( VbMin, "#0.0") + ",," + Format( VcMin, "#0.0")
 StrLine$ = StrLine$ + ",," + Format( V1Min, "#0.0") + ",," + Format( V2Min, "#0.0") + ",," + Format( V0Min, "#0.0")
 Print #1, StrLine$
 Close 1
    
 doStop:
 ProgressDialog(0)
 Print "Results have been saved to " + ExcelFile$
 
 Exit Sub
 
HasError:
 Print "Error: "; ErrorString( ) 
End Sub


' ===================== End of Main() =========================================

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
